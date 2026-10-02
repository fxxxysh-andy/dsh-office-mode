/**
 * 脚本运行环境与 SDK。
 *
 * 设计取向（对应「极简」「少调用」「批量」三条要求）：
 *   - 模型只看到少数几个工具（office_help、office_run、office_memory 与检索那几个），
 *     而不是每种格式一堆入口。
 *   - office_run 里给的是「造文档的对象」，不是逐条工具调用：一次脚本可以建几个
 *     文件、改一批 Markdown、算一遍排版风险，全部在一次工具调用里完成。
 *   - SDK 只暴露必要面：word/excel/ppt/pdf/python/files/cache/theme/log。没有
 *     process、没有 require、没有 fs —— 想碰文件只能走 files.*，因此所有产物都可被记账。
 *
 * @module dsh-office-mode/sdk
 */
import { createEnv } from './engine/kit.js';
import { themeCatalog, resolveTheme } from './engine/theme.js';
import { archiveExtract, archiveFind, archiveHintForBytes, archiveInfo, archiveList, archiveText } from './archive.js';
import { avCheck, avExtract, avFrames, avInfo, avTranscribe } from './av.js';
import { imageChannels, imageFetch, imageSearch } from './image-source.js';
import { pdfEngines, pdfInfo, pdfPages, pdfText } from './pdf.js';
import { previewCheck, previewImages, previewPdf } from './preview.js';
import { pythonCheck, pythonFile, pythonRun } from './python.js';
import { loadAllFormats, formatByExtension, formatIds } from './registry.js';

/** 把 `{find, replace}` 的各种写法收敛成数组。 */
function normalizeEdits(edits) {
    if (Array.isArray(edits)) {
        return edits.map((item) => {
            if (Array.isArray(item)) return { find: item[0], replace: item[1], all: true };
            return { find: item?.find, replace: item?.replace, all: item?.all !== false };
        });
    }
    if (edits !== null && typeof edits === 'object') {
        return [{ find: edits.find, replace: edits.replace, all: edits.all !== false }];
    }
    return [];
}

const OFFICE_EXTENSIONS = new Set(['.docx', '.xlsx', '.pptx']);

/**
 * OOXML 文档（也是 ZIP）该用哪个入口读 —— `office.files.read` 的错误里要指对路。
 *
 * 不指路的话，模型读到「这是一份 zip 归档，换 office.archive」会真的去列 XML 部件名，
 * 而它要的其实是 `office.word.read` / `excel.read` / `ppt.read` 的结构报告。
 */
function ooxmlReaderOf(filePath) {
    const at = String(filePath ?? '').lastIndexOf('.');
    if (at === -1) return undefined;
    const ext = String(filePath).slice(at).toLowerCase();
    if (ext === '.docx') return { what: 'Word 文档', how: 'office.word.read(path)' };
    if (ext === '.xlsx') return { what: 'Excel 工作簿', how: 'office.excel.read(path)' };
    if (ext === '.pptx') return { what: 'PPT 演示文稿', how: 'office.ppt.read(path)' };
    if (ext === '.tex') return { what: 'LaTeX 项目', how: 'office.tex.read(path)' };
    return undefined;
}

/**
 * 这个写出的文件要不要按格式复检。
 *
 * 判断以注册表为准（谁注册了扩展名谁负责 read()），而不是写死三件套 ——
 * 否则新加的格式（例如 .tex）写出来也不会被自动复检，反馈里就只剩一个体积。
 * 注册表查不到时退回三件套，保证老行为不变。
 */
export function isOfficeFile(filePath) {
    if (formatByExtension(filePath) !== undefined) return true;
    const at = String(filePath ?? '').lastIndexOf('.');
    return at !== -1 && OFFICE_EXTENSIONS.has(String(filePath).slice(at).toLowerCase());
}

/**
 * 组装脚本可见的 `office` 对象。
 * 格式模块加载失败时给出「会抛清晰错误」的占位入口，而不是静默消失。
 */
export async function buildSdk(options) {
    const { env, cache, config } = options;
    const loaded = await loadAllFormats();
    const modules = new Map();
    const loadIssues = [];
    for (const item of loaded) {
        if (item.module) modules.set(item.entry.id, item.module);
        else loadIssues.push(`${item.entry.id}：${item.error?.message ?? '加载失败'}`);
    }

    const notes = [];
    const logs = [];
    const warnings = [...env.warnings];

    const wrap = (id) => {
        const entry = loaded.find((item) => item.entry.id === id)?.entry;
        const module = modules.get(id);
        if (module === undefined) {
            const reason = loadIssues.find((line) => line.startsWith(`${id}：`)) ?? '模块未加载';
            const broken = () => {
                throw new Error(`office.${id} 不可用（${reason}）`);
            };
            return { create: broken, read: broken, edit: broken };
        }
        const withTheme = (spec) => {
            const input = spec === null || typeof spec !== 'object' ? {} : spec;
            return { ...input, theme: input.theme ?? config.defaultTheme };
        };
        const api = {
            create: (spec) => module.create(withTheme(spec), env),
            read: (filePath) => module.read(filePath, env),
            edit: (filePath, ops, editOptions) => module.edit(filePath, ops, { ...env, ...(editOptions ?? {}) }),
        };
        // PPT 的就地修订走同一个格式入口：`office.ppt.readSlides` / `office.ppt.revise`
        // 与 create/read/edit 并列，调用方仍然只面对「一个格式对象」。
        // 用 getter 惰性挂载：`edit` 在 word/excel 上才是真编辑，
        // ppt 的 readSlides/revise 在别的格式上不存在，不能凭空长出来。
        //
        // 第二个参数是插件上下文 `{cache, config}`：格式模块拿不到插件上下文，
        // 需要落中间产物（例如 office.ppt.images 把图导出到缓存目录）的地方只能从这里取。
        // 已有的工厂（readSlides/revise/tex.compile…）都忽略这个参数，因此是向后兼容的。
        const extra = { cache, config };
        for (const [name, fn] of Object.entries(module.api ?? {})) {
            if (api[name] === undefined) api[name] = fn(env, extra);
        }
        return api;
    };

    const office = {
        word: wrap('word'),
        excel: wrap('excel'),
        ppt: wrap('ppt'),
        // LaTeX：产物是一个「项目」（主文件 + 按章分文件 + 模板文件），
        // create/read/edit 走统一的格式入口，compile 由 module.api 挂上来。
        tex: wrap('tex'),

        /** 文件操作：批量改 Markdown 也只写一次盘。 */
        files: {
            /**
             * 读文本（read / edit / template 共用一个入口）。
             *
             * 读到 ZIP / gzip / zstd 时**不返回乱码**：那种「读出来了但全是问号」
             * 的结果会让模型以为文件坏了或内容就是这样。这里按魔数认出来就抛错，
             * 并把该走的路（office.archive）写在错误里 —— 一次调用就能自我纠正。
             *
             * 三个动作共用它是**必须的**：`edit` / `template` 会拿读出来的文本整份写回，
             * 按乱码读再写回等于把压缩包毁掉（数据丢失，不是「没命中」）。
             */
            read: (filePath) => {
                const data = env.readFile(filePath);
                const hint = archiveHintForBytes(new Uint8Array(data.buffer, data.byteOffset, data.byteLength));
                if (hint !== undefined) {
                    // OOXML 也是 ZIP：对 .docx/.xlsx/.pptx 说「换 office.archive」是指错了路
                    // （那里只能看到 XML 部件名）。按扩展名指到对应的格式读取器。
                    const reader = ooxmlReaderOf(filePath);
                    throw new Error(`office.files.read：${filePath} ${reader === undefined
                        ? hint
                        : `这是 OOXML 文档（${reader.what}），不是文本：用 ${reader.how} 读它。`}`);
                }
                return data.toString('utf8');
            },
            write(filePath, content) {
                const written = env.writeFile(filePath, typeof content === 'string' ? content : JSON.stringify(content, null, 2));
                return { path: written.path, bytes: written.bytes };
            },
            /**
             * 批量字面替换，原子写一次。
             * 没有命中的 find 会进 `missing`，不会静默吃掉。
             */
            edit(filePath, edits) {
                const original = office.files.read(filePath);
                let text = original;
                const applied = [];
                const missing = [];
                for (const edit of normalizeEdits(edits)) {
                    if (typeof edit.find !== 'string' || edit.find === '') {
                        missing.push({ find: String(edit.find ?? ''), reason: 'find 不能为空' });
                        continue;
                    }
                    const target = String(edit.replace ?? '');
                    if (edit.all === false) {
                        if (!text.includes(edit.find)) {
                            missing.push({ find: edit.find, reason: '未命中' });
                            continue;
                        }
                        text = text.replace(edit.find, target);
                        applied.push(edit.find);
                    } else {
                        const count = text.split(edit.find).length - 1;
                        if (count === 0) {
                            missing.push({ find: edit.find, reason: '未命中' });
                            continue;
                        }
                        text = text.split(edit.find).join(target);
                        applied.push(`${edit.find} ×${String(count)}`);
                    }
                }
                if (applied.length === 0) {
                    return { path: filePath, changed: false, applied, missing, bytes: Buffer.byteLength(original, 'utf8') };
                }
                const written = env.writeFile(filePath, text);
                return { path: written.path, changed: true, applied, missing, bytes: written.bytes };
            },
            /** 把 `{{key}}` 占位符一次性替换掉（模板填充）。 */
            template(filePath, variables) {
                const original = office.files.read(filePath);
                let text = original;
                const used = [];
                for (const [key, value] of Object.entries(variables ?? {})) {
                    const token = `{{${key}}}`;
                    if (!text.includes(token)) continue;
                    text = text.split(token).join(String(value));
                    used.push(key);
                }
                if (text === original) return { path: filePath, changed: false, used };
                const written = env.writeFile(filePath, text);
                return { path: written.path, changed: true, used, bytes: written.bytes };
            },
            exists: (filePath) => env.exists(filePath),
            list: (dirPath) => env.list(dirPath),
            remove: (filePath) => env.remove(filePath),
            stat(filePath) {
                const info = env.stat(filePath);
                return info === undefined ? null : { bytes: info.size, modifiedMs: info.mtimeMs };
            },
        },

        /**
         * 归档与会话日志（office.archive）。
         *
         * 会话导出（`session*.zip`，里面是 `session.v4.jsonl` 与 `media/*`）、
         * 单独压缩的 jsonl（`.gz` / `.zst`）以前在办公模式里读不出：按文本读只会
         * 得到乱码。这里按魔数判型（ZIP / gzip / zstd / 纯文本），给**五个**有界的读法。
         */
        archive: {
            info: (filePath, options) => archiveInfo(filePath, env, options ?? {}),
            list: (filePath, options) => archiveList(filePath, env, options ?? {}),
            text: (filePath, options) => archiveText(filePath, env, options ?? {}),
            find: (filePath, options) => archiveFind(filePath, env, options ?? {}),
            extract: (filePath, options) => archiveExtract(filePath, env, options ?? {}),
        },

        /**
         * 渲染预览（office.preview）：把自家写出的文档渲染成页面图 / PDF。
         *
         * 引擎是宿主随部署安装的那份 LibreOffice（`@deepseek-ai/libreoffice-kit`），
         * 不是要另装的插件；解析不到时 check() 会如实说缺什么。
         * 渲染一次几秒，所以按需调用：生成完想看一眼版式再调。
         */
        preview: {
            check: () => previewCheck(),
            render: (filePath, options) => previewImages(filePath, options ?? {}, env, cache),
            pdf: (filePath, options) => previewPdf(filePath, options ?? {}, env, cache),
        },

        /**
         * 配图（office.image）：网络取图与图库检索，并带上署名。
         *
         * 只用免 Key 的公开图库（维基共享资源 / Openverse），HTTP 走与取网页同一套
         * 字节通道（地址校验、限长、超时分类都只有一份实现）。取回来的图落到工作目录，
         * 来源与许可随返回值一起给 —— 嵌进文档时把 `credit` 传给 `source` 选项。
         */
        image: {
            channels: () => imageChannels(),
            search: (query, options) => imageSearch(query, options ?? {}),
            fetch: (ref, options) => imageFetch(ref, options ?? {}, env),
        },

        /**
         * PDF：读进来而不是造出来。三条路对应两类来源 ——
         * 文本型抽文字，图像型（扫描 / 手写）渲染成图片再交给 read_image。
         */
        pdf: {
            info: (filePath, options) => pdfInfo(filePath, env, { ...(options ?? {}), cache }),
            text: (filePath, options) => pdfText(filePath, options ?? {}, env, cache),
            pages: (filePath, options) => pdfPages(filePath, options ?? {}, env, cache, config),
            engines: () => pdfEngines(),
        },

        /**
         * 音频与视频的内容提取（office.av）。
         *
         * 两件事：声音转文字（本机 SenseVoice 离线转写，分块 + 逐句时间戳），
         * 视频按时间点抽帧成图片（交给 read_image 看）。四样能力 —— ffmpeg /
         * ffprobe / 模型 / sherpa-onnx 运行时 —— 缺任何一样都会明确报错，
         * 先问一句用 office.av.check()，不要盲调。
         */
        av: {
            check: (options) => avCheck({ ...(options ?? {}), env, cache, config }),
            info: (filePath, options) => avInfo(filePath, options ?? {}, env, cache, config),
            transcribe: (filePath, options) => avTranscribe(filePath, options ?? {}, env, cache, config),
            frames: (filePath, options) => avFrames(filePath, options ?? {}, env, cache, config),
            extract: (filePath, options) => avExtract(filePath, options ?? {}, env, cache, config),
        },

        /**
         * Python：办公模式里唯一的「编程」出口，只做科学计算与绘图。
         *
         * 命令行仍然是关掉的：这里跑的是插件探测到的解释器，脚本由插件落盘再执行，
         * 产物按修改时间收集回反馈。图直接落缓存目录，可以喂给 word 的
         * builder.image() 或 ppt 的 deck.image()。
         */
        python: {
            run: (code, options) => pythonRun(code, options ?? {}, env, cache, config),
            file: (filePath, options) => pythonFile(filePath, options ?? {}, env, cache, config),
            check: (options) => pythonCheck({ ...(options ?? {}), config }),
        },

        /** 缓存：中间产物只允许放这里。跨调用保留（按 TTL 自动清理），可命中复用。 */
        cache: {
            dir: () => cache.rel,
            path: (name) => cache.path(name),
            write(name, data) {
                const written = cache.write(name, data);
                return { path: written.path, bytes: written.bytes };
            },
            list: () => cache.list(),
            clear: () => cache.clear(),
            /** 当前缓存的体积与命中统计（跨调用复用了多少次、省了多少字节）。 */
            stats: () => ({ ...cache.stats, bytes: cache.bytes(), files: cache.list().filter((entry) => !entry.dir).length }),
        },

        theme: {
            list: () => themeCatalog(),
            resolve: (id) => resolveTheme(id).theme,
            current: () => config.defaultTheme,
        },

        log: (...parts) => {
            const line = parts.map((part) => (typeof part === 'string' ? part : JSON.stringify(part))).join(' ');
            logs.push(line);
            env.log(line);
        },
        warn: (...parts) => {
            const line = parts.map((part) => (typeof part === 'string' ? part : JSON.stringify(part))).join(' ');
            warnings.push(line);
            env.warn(line);
        },
        note: (...parts) => {
            const line = parts.map((part) => (typeof part === 'string' ? part : JSON.stringify(part))).join(' ');
            notes.push(line);
            env.note(line);
        },
        assert(condition, message) {
            if (!condition) throw new Error(`断言失败：${message ?? '条件不成立'}`);
            return true;
        },
        formats: () => [...formatIds(), 'pdf'],
    };

    return { office, loadIssues, notes, logs, warnings };
}

/** 造一个只在测试里用的 env（给格式模块的自测复用）。 */
export function testEnv(root, themeId = 'business') {
    return createEnv({ root, themeResolver: (id) => resolveTheme(id ?? themeId) });
}
