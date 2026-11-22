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
import { createEnv, displayPath } from './engine/kit.js';
import { themeCatalog, resolveTheme } from './engine/theme.js';
import { archiveExtract, archiveFind, archiveHintForBytes, archiveInfo, archiveList, archiveText } from './archive.js';
import { avCheck, avExtract, avFrames, avInfo, avTranscribe } from './av.js';
import { convertCheck, convertFile } from './convert.js';
import { imageChannels, imageExtract, imageFetch, imageFromPage, imageInfo, imageSearch } from './image-source.js';
import { pdfEngines, pdfInfo, pdfPages, pdfText } from './pdf.js';
import { previewCheck, previewImages, previewPdf } from './preview.js';
import { pythonCheck, pythonFile, pythonRun } from './python.js';
import { loadAllFormats, formatByExtension, formatIds } from './registry.js';
import { renderWritingPlan, writingAudit, writingDeai, writingPlan, writingSignals } from './writing.js';

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
 * 二进制入参 → Buffer 视图；文本与其它值返回 `undefined`。
 *
 * 用 `ArrayBuffer.isView` 而不是 `instanceof`：脚本跑在另一个 realm 里，
 * 里面 `new Uint8Array(...)` 造出来的视图在宿主侧 `instanceof Uint8Array` 为 **假**
 * （实测），而 `ArrayBuffer.isView` 是规范级的、跨 realm 也认。
 * 返回的是**视图**（不复制内存），真正的复制发生在 `env.writeFile` 里。
 */
function binaryOf(value) {
    if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
    if (isRawBuffer(value)) return Buffer.from(value);
    return undefined;
}

/** 裸 `ArrayBuffer` / `SharedArrayBuffer`（`ArrayBuffer.isView` 对它们为假）。 */
function isRawBuffer(value) {
    if (value === null || typeof value !== 'object') return false;
    const tag = Object.prototype.toString.call(value);
    return tag === '[object ArrayBuffer]' || tag === '[object SharedArrayBuffer]';
}

/**
 * 按 JSON 写 `content` 时用的 replacer：**嵌套的二进制当场报错**。
 *
 * 依据（第五十一轮独立复核打穿的一条）：`write(path, {img: bytes})` 这种形状下，
 * 顶层不是字节，于是整份走 `JSON.stringify` —— 一个 `Uint8Array([0,255])` 会变成
 * `{"0":0,"1":255}`，落盘「成功」、没有告警，而文档写着「字节按字节写，不会被 JSON 化」。
 * 一份看起来写成功的坏文件正是这一轮要消灭的东西，所以这里选**响亮地失败**：
 * 报错里说清是哪个字段、以及两条出路（字节放顶层，或自己转成文本）。
 *
 * `SharedArrayBuffer` 同类：`JSON.stringify` 给它 `{}`（没有可枚举自有属性），
 * 也是一次静默丢失。
 */
function jsonReplacer(key, value) {
    if (ArrayBuffer.isView(value) || isRawBuffer(value)) {
        throw new Error(`office.files.write：content${key === '' ? ' 本身' : ` 的 ${key} 字段`}是二进制（${Object.prototype.toString.call(value)}）。`
            + 'JSON 化会把它变成一串数字键（静默坏文件），所以这里直接拒绝：'
            + '要写字节就把字节本身当 content 传（office.files.write(path, bytes)），或先自己转成文本 / base64。');
    }
    return value;
}

/**
 * 已知的二进制魔数（不看扩展名）：`read` 的「这不是文本」判据里与空字节并列的一条。
 *
 * 为什么两条都要：只有空字节判据时，一份**纯 ASCII 的合法 PDF**（`%PDF-1.4` 开头、
 * 前 8 KB 里恰好没有 `0x00`）会被当文本返回原文，而只截到魔数的 PNG（8 字节签名）
 * 同样漏过 —— 独立复核用真实夹具量过：工作区 245 个二进制夹具里 128 个靠空字节命中、
 * 117 个走压缩包魔数、2 个假 BMP 漏过；15 份真 PDF 全都有空字节，但**不保证**下一份有。
 *
 * 只收**四字节及以上**的签名（JPEG / ELF 虽然短，但开头是 0xFF / 0x7F 这类
 * 文本里不会出现的字节）。两字节的 `BM` / `MZ` **有意不收**：任何以「BM」「MZ」
 * 开头的文本（例如一份讲 BMP 格式的 Markdown）都会被误判，而真正的 .bmp / .exe
 * 在前 8 KB 里必然带空字节，空字节那条判据接得住。
 */
const BINARY_MAGICS = [
    [0x25, 0x50, 0x44, 0x46], // %PDF
    [0x89, 0x50, 0x4e, 0x47], // ‰PNG
    [0xff, 0xd8, 0xff], // JPEG
    [0x47, 0x49, 0x46, 0x38], // GIF8
    [0x7f, 0x45, 0x4c, 0x46], // \x7fELF
];

/** 开头命中已知二进制魔数吗？ */
function hasBinaryMagic(data) {
    return BINARY_MAGICS.some((magic) => magic.every((byte, index) => data[index] === byte));
}

/**
 * 前 8 KB 里有没有空字节 —— 有就当二进制。文本文件里不会有 0x00，
 * 而「按 utf8 读一张 PNG」得到的是一串替换字符，看着像文件坏了、其实只是走错了路。
 *
 * @returns {number} 第一个空字节的下标，没有则 -1
 */
function firstNulByte(data) {
    const limit = Math.min(data.length, 8_192);
    for (let index = 0; index < limit; index += 1) {
        if (data[index] === 0) return index;
    }
    return -1;
}

/** 二进制文件的「换哪条路」指路。只列真有专门读取器的：指错了比不指更糟。 */
function binaryReaderOf(filePath) {
    const at = String(filePath ?? '').lastIndexOf('.');
    if (at === -1) return undefined;
    const ext = String(filePath).slice(at).toLowerCase();
    if (ext === '.pdf') return { what: 'PDF', how: 'office.pdf.info / text / pages' };
    return undefined;
}

/**
 * 整份复制或移动一个文件（`office.files.copy` / `office.files.move` 的实现）。
 *
 * 四条闸都在**内核动作之前**：
 *   1. 源与目标是同一个文件 —— 移动会先把内容写进目标、再删源，同一个路径上就是自毁
 *      （Windows 的路径大小写不敏感，所以比较时按平台归一）；
 *   2. 目标是已存在的**目录** —— 「已存在」那句报错会把人引向 overwrite:true，
 *      而 overwrite 一个目录只会得到 EISDIR；
 *   3. 目标已存在且没有 `overwrite:true` —— 与 `preview.pdf` / `image.fetch` /
 *      `archive.extract` 同一条口径（安静覆盖一份交付物是数据丢失，不是「方便」）；
 *   4. 源不存在 / 源是目录 —— 报错里给下一步（`list`），不要只说 ENOENT。
 */
function transferFile(env, action, from, to, options) {
    const source = typeof from === 'string' ? from.trim() : '';
    const target = typeof to === 'string' ? to.trim() : '';
    if (source === '' || target === '') {
        throw new Error(`office.files.${action}：要给 from 与 to 两个路径（整份复制 / 移动一个文件）。`);
    }
    // 「同一个文件」不能只比路径字符串：硬链接与符号链接是同一个 inode 的别名，
    // 而 `move` 到自身的实现是「写目标 + 删源」—— 别名下就是自毁。判定交给内核层
    // （`env.sameFile`：realpath + inode，拿不到 inode 时退回路径比较）。
    if (env.sameFile(source, target)) {
        const caseOnly = process.platform === 'win32' && env.resolve(source).toLowerCase() === env.resolve(target).toLowerCase();
        throw new Error(`office.files.${action}：源与目标是同一个文件（${source}）；换一个 to。`
            + (caseOnly && action === 'move' ? '（只改大小写的改名：先 move 到一个中间名，再 move 回目标名。）' : ''));
    }
    const targetInfo = env.stat(target);
    if (targetInfo?.isDirectory?.() === true) {
        throw new Error(`office.files.${action}：${target} 是一个目录，不能当目标文件；换一个 to。`);
    }
    if (options.overwrite !== true && targetInfo !== undefined) {
        throw new Error(`office.files.${action}：${target} 已存在；换一个 to，或显式传 overwrite:true。`);
    }
    const sourceInfo = env.stat(source);
    if (sourceInfo === undefined) {
        throw new Error(`office.files.${action}：源文件不存在（${source}）。用 office.files.list 看目录里有什么。`);
    }
    if (sourceInfo.isDirectory?.() === true) {
        throw new Error(`office.files.${action}：${source} 是一个目录；这个动作只处理单个文件。`);
    }
    try {
        const entry = action === 'copy' ? env.copyFile(source, target) : env.moveFile(source, target);
        return { path: entry.path, from: displayPath(env.root, env.resolve(source)), bytes: entry.bytes };
    } catch (error) {
        // 内核动作的原始 Node 错误（EEXIST / EPERM / ENOSPC / EISDIR…）里是绝对路径与
        // 系统措辞，模型照着做不了。包一层：说清是哪个动作、哪个路径，再给一句下一步；
        // 原始措辞里本机的绝对路径换回调用方给的写法（台账与回执不该被绝对路径灌满）。
        const raw = (error instanceof Error ? error.message : String(error)).split('\n')[0];
        const message = raw
            .split(env.resolve(source)).join(source)
            .split(env.resolve(target)).join(target)
            .replace(/[A-Za-z]:\\[^\s'"]*/g, '<绝对路径>');
        const code = typeof error?.code === 'string' ? error.code : '';
        const hint = code === 'EEXIST' || code === 'ENOTDIR'
            ? '目标路径上有同名目录或中间路径不是目录；换一个 to。'
            : (code === 'EPERM' || code === 'EACCES' || code === 'EBUSY'
                ? '目标被占用或没有写权限（只读文件、被 Office 打开、权限不足）。'
                : '换一个 to 再试；要看目录里有什么用 office.files.list。');
        throw new Error(`office.files.${action}：${action === 'copy' ? '复制' : '移动'} ${source} → ${target} 失败`
            + `${code === '' ? '' : `（${code}）`}：${message}。${hint}`);
    }
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
            // PPT 的 read 打开**内容密度自检**（第六十九轮）：工具面读回来的报告要能
            // 直接回答「这页是不是太空」，而不是等用户来问。word / excel 的 read
            // 忽略第三个参数，保持原样。
            read: (filePath) => module.read(filePath, env, id === 'ppt' ? { density: true } : undefined),
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

        /**
         * 文件操作：批量改 Markdown 也只写一次盘。
         *
         * 文本与二进制分两条读法（`read` / `readBytes`），写入与复制是**二进制安全**的：
         * 一份 .png / .docx 从读进来到写回去不经过 utf8，也不会被 JSON 化。
         */
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
             *
             * 非压缩包的二进制（PNG / PDF / 可执行文件…）同样不返回乱码：前 8 KB 里有
             * 空字节、或开头命中已知二进制魔数，就当二进制报错并指向 `readBytes` /
             * 具体的格式读取器（`新-8`）。两条判据是互补的 —— 独立复核实测：纯 ASCII 的
             * 合法 PDF 与只截到 8 字节签名的 PNG 都**不含**空字节，只靠空字节会漏。
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
                const zero = firstNulByte(data);
                if (zero !== -1 || hasBinaryMagic(data)) {
                    const why = zero === -1 ? '开头就是二进制魔数' : `第 ${zero + 1} 字节就是 0x00`;
                    const reader = binaryReaderOf(filePath);
                    throw new Error(`office.files.read：${filePath} 是二进制文件（${why}），`
                        + `按文本读只会得到乱码。${reader === undefined
                            ? '要字节用 office.files.readBytes，要整份复制用 office.files.copy。'
                            : `这是 ${reader.what}，用 ${reader.how} 读它。`}`);
                }
                return data.toString('utf8');
            },
            /**
             * 读原始字节（二进制安全）。返回一个字节视图，直接交给 `write` / `copy` 用
             * （`office.files.write(path, office.files.readBytes(src))` 就是整份复制，
             * 只是 `copy` 更省一步、且不经过脚本内存）。
             *
             * 文本也能用它读（拿到的是 Buffer，自己 `toString('utf8')`）；反过来不行 ——
             * 二进制不能走 `read`，那条路按 utf8 解，解出来是补不回来的。
             */
            readBytes: (filePath) => env.readFile(filePath),
            /**
             * 整份覆盖写。`content` 是字符串就按 utf8 写，是字节（Buffer / Uint8Array /
             * ArrayBuffer）就按字节写，其它值按 2 空格缩进的 JSON 写。
             *
             * 三类拒绝（都是「静默写出一份坏文件」的入口，见独立复核 2026-10-04）：
             *   · `undefined`（往下是宿主内部的 `Buffer.from` 报错）；
             *   · `JSON.stringify` 结果为 `undefined` 的形状（函数 / `Symbol` / `toJSON` 返回 undefined）；
             *   · 嵌在对象里的二进制（replacer 拒绝）与 `Map` / `Set`（JSON 化只会得到 `{}`）。
             */
            write(filePath, content) {
                if (content === undefined) {
                    throw new Error('office.files.write：content 不能是 undefined；要写空文件传 \'\'，要写对象传对象。');
                }
                const binary = binaryOf(content);
                let payload;
                if (binary !== undefined) {
                    payload = binary;
                } else if (typeof content === 'string') {
                    payload = content;
                } else {
                    // Map / Set 没有可枚举自有属性，replacer 根本不会被调到 —— `JSON.stringify`
                    // 直接给 `{}`，包在里面的字节或数据全丢。这一类只能在进 JSON 之前拦。
                    const tag = Object.prototype.toString.call(content);
                    if (tag === '[object Map]' || tag === '[object Set]') {
                        throw new Error(`office.files.write：content 是 ${tag.slice(8, -1)}，JSON 化只会得到 {}（内容全丢）。`
                            + '先转成数组或普通对象（Array.from / Object.fromEntries）再写。');
                    }
                    try {
                        payload = JSON.stringify(content, jsonReplacer, 2);
                    } catch (error) {
                        // replacer 自己抛的那条已经是可读文案（带 office.files.write 前缀），别套两层。
                        const message = error instanceof Error ? error.message : String(error);
                        if (message.startsWith('office.files.write：')) throw error;
                        throw new Error(`office.files.write：content 无法写成 JSON（${message}）。`);
                    }
                    if (payload === undefined) {
                        throw new Error('office.files.write：content 写不出文本（函数 / Symbol / toJSON 返回了 undefined）；'
                            + '要写空文件传 \'\'，要写对象传对象。');
                    }
                }
                const written = env.writeFile(filePath, payload);
                return { path: written.path, bytes: written.bytes };
            },
            /**
             * 整份复制 / 移动（二进制安全，不经过文本）。`overwrite:true` 才允许覆盖已有的目标。
             * 复制另会登记「读过源文件」；移动不会（移动之后源路径已经不存在）。
             */
            copy: (from, to, options) => transferFile(env, 'copy', from, to, options ?? {}),
            move: (from, to, options) => transferFile(env, 'move', from, to, options ?? {}),
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
         * 文档互转与统一出 PDF（office.convert）。
         *
         * 「手上的这份换个格式」不再只能重新生成一遍：md / docx / tex 之间互转（pandoc），
         * 各种文档都能出 PDF（.tex 走本机 TeX，office 三件套走宿主 LibreOffice 引擎，
         * 纯文本源两条路）。引擎都在运行期探测，`check()` 报实况。
         */
        convert: {
            check: () => convertCheck(),
            file: (filePath, options) => convertFile(filePath, options ?? {}, env, cache),
        },

        /**
         * 配图（office.image）：图片的获取与使用，并带上署名。
         *
         * 四条取图的路各管一段：info 看本地图片的尺寸 / 哈希 / 能不能嵌进 Word、PPT；
         * fetch 把一张图（直链，或 search / fromPage 的候选）取回工作目录；fromPage
         * 列出一个网页里的图片候选（og:image 优先，懒加载与 srcset 也认）；extract
         * 把 docx / pptx / xlsx 里已经嵌进去的媒体抽出来复用。图库只走免 Key 的公开
         * 来源（维基共享资源 / Openverse），HTTP 走与取网页同一套字节通道（地址校验、
         * 限长、超时分类都只有一份实现）。取回来的图落到工作目录，来源与许可随返回值
         * 一起给 —— 嵌进文档时把 `credit` 传给 `source` 选项；网页配图没有许可信息，
         * 正式发布前要自己核对。
         */
        image: {
            channels: () => imageChannels(),
            search: (query, options) => imageSearch(query, options ?? {}),
            fetch: (ref, options) => imageFetch(ref, options ?? {}, env),
            info: (filePath) => imageInfo(filePath, env),
            fromPage: (url, options) => imageFromPage(url, options ?? {}),
            extract: (filePath, options) => imageExtract(filePath, env, options ?? {}),
        },

        /**
         * 写作规划与自检（office.writing）：把「哪一节该多写」变成可计算的表，
         * 写完再按同一张表查。不是提示词，也不是结构模板。
         *
         * - plan(input, options)：按难度 / 读者收益给每节分配目标字数与深度；
         *   返回的 markdown 可直接 office.log，让模型照着写。
         * - audit(text, plan?)：句长、段长、套话、锚点、让步、重点节是否写透；
         *   warnings 每条都能落到具体改稿动作。
         * - deai(text, options)：只报可解释的句式与文体偏离（A1 对举 / A2 提示语冒号 /
         *   M1 元话语），按 register 调口径；safe 只删/改提示语前缀，不给分、不改论证。
         * - signals(text)：只看表面量，不判分；调参或自己做评测时用。
         * - render(plan)：把规划对象重新渲染成 Markdown。
         *
         * 大模型 A/B 评测不在脚本里：用仓库的 scripts/writing-eval.mjs（无 Key 会
         * SKIP），这里只提供确定性的规划、自检与去 AI 味报告，保证评测的两臂差在规划而不是提示词。
         */
        writing: {
            plan: (input, options) => writingPlan(input, options ?? {}),
            audit: (text, planOrOptions) => writingAudit(text, planOrOptions),
            deai: (text, options) => writingDeai(text, options ?? {}),
            signals: (text) => writingSignals(text),
            render: (plan) => renderWritingPlan(plan),
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
