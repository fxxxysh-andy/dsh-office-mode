/**
 * 渲染预览（office.preview）：把自家写出的 docx / xlsx / pptx 渲染成页面图，交给 read_image 看。
 *
 * 为什么需要它（历史遗留 `2-1`）：办公模式能生成三件套、也能**读回结构**
 * （office.word.read / excel.read / ppt.read），但「看起来对不对」一直只能靠
 * 自研解析器报的排版风险推断 —— 模型看不到画面。第三轮的建议是用随部署安装的
 * LibreOffice 补这一环；`2-1` 说的就是这个入口。
 *
 * 三条口径：
 *   1. **引擎是宿主随部署装的那份**（`@deepseek-ai/libreoffice-kit`，原生 LibreOffice
 *      编译出来的引擎），不是让用户另装插件；解析不到时**如实报缺什么**，
 *      不假装能预览。
 *   2. **零依赖**：本模块不 import 那个包，只在运行时从宿主的解析位置取；
 *      取不到就降级成一条可执行的错。
 *   3. **产物只进缓存目录**（`.office/cache/preview/…`），按「文件内容 + dpi + 页选择」
 *      做内容键：同一份文件重复预览命中已有图，不重渲染。
 *
 * 渲染一次 2～5 秒（原生引擎，本机实测），所以预览是**按需**动作：
 * 生成的文档要看一眼就调它，别每写一份都调。
 *
 * @module dsh-office-mode/preview
 */
import { createHash } from 'node:crypto';
import { closeSync, existsSync, mkdirSync, openSync, readFileSync, readSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, isAbsolute, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

import { assertInsideRoot, displayPath } from './engine/kit.js';

/** 引擎包名（宿主随部署安装；插件自己不依赖它）。 */
export const PREVIEW_KIT_SPEC = '@deepseek-ai/libreoffice-kit';

/** 支持的输入扩展名（引擎的 IMAGE_FORMATS 取不到时的兜底清单）。 */
export const PREVIEW_EXTENSIONS = Object.freeze(['.doc', '.docx', '.odt', '.xls', '.xlsx', '.ods', '.ppt', '.pptx', '.odp', '.pdf']);

/** 一次预览最多渲染多少页（引擎自己默认 100；预览是给人看的，不需要那么多）。 */
export const PREVIEW_MAX_PAGES = 30;

/** 预览图默认 dpi：够看清版式与文字块，又不至于一张几 MB。 */
export const PREVIEW_DEFAULT_DPI = 110;

/** 解析缓存：同一进程里只找一次（找不到也记着，别每次调用都翻一遍盘）。 */
let kitResolution;

/**
 * 宿主解析锚点：与 src/index.js 的 resolveHostModule 同一套口径
 * （profile 根 → 宿主安装位置 → 宿主入口目录），另加 cwd 兜底。
 */
function hostAnchors() {
    const anchors = [
        process.env.DSH_PROFILE_DIR,
        process.env.DSH_HOST_ROOT,
        process.env.DSH_CHECKOUT,
    ].filter((value) => typeof value === 'string' && value !== '');
    if (typeof process.argv[1] === 'string' && process.argv[1] !== '') anchors.push(dirname(process.argv[1]));
    anchors.push(process.cwd());
    return [...new Set(anchors)];
}

function requireFrom(anchor) {
    try {
        return createRequire(pathToFileURL(join(anchor, 'package.json')).href);
    } catch {
        return undefined;
    }
}

/** 在锚点链上解析引擎包的入口路径；都找不到返回 undefined。 */
function resolveKitPath() {
    for (const anchor of hostAnchors()) {
        const req = requireFrom(anchor);
        if (req === undefined) continue;
        try {
            return req.resolve(PREVIEW_KIT_SPEC);
        } catch {
            // 换下一个锚点
        }
    }
    // 随部署安装的包在 @deepseek-ai/dsh 自己的 node_modules 里：先解析 dsh 再从它那里解析。
    for (const anchor of hostAnchors()) {
        const req = requireFrom(anchor);
        if (req === undefined) continue;
        let dshDir;
        try {
            dshDir = dirname(req.resolve('@deepseek-ai/dsh/package.json'));
        } catch {
            continue;
        }
        const inner = requireFrom(dshDir);
        if (inner === undefined) continue;
        try {
            return inner.resolve(PREVIEW_KIT_SPEC);
        } catch {
            // 继续找
        }
    }
    return undefined;
}

/**
 * 取渲染引擎。返回 `{ module }` 或 `{ error }`（不抛，调用方决定怎么报）。
 * 用动态 import 而不是 require：那个包是 ESM（`type: module`），
 * 而这里是运行期调用，异步完全可以接受。
 */
export async function loadPreviewKit() {
    if (kitResolution !== undefined) return kitResolution;
    const kitPath = resolveKitPath();
    if (kitPath === undefined) {
        kitResolution = {
            error: `找不到渲染引擎 ${PREVIEW_KIT_SPEC}（它随 DSH 部署安装，不是要另装的插件）。`
                + '预览需要它；找不到时办公模式的其余能力不受影响。',
        };
        return kitResolution;
    }
    try {
        const module = await import(pathToFileURL(kitPath).href);
        if (typeof module.createConverter !== 'function') {
            kitResolution = { error: `渲染引擎 ${PREVIEW_KIT_SPEC} 在 ${kitPath}，但没有 createConverter 导出（版本不匹配）。` };
            return kitResolution;
        }
        kitResolution = { module, path: kitPath };
        return kitResolution;
    } catch (error) {
        kitResolution = { error: `渲染引擎 ${PREVIEW_KIT_SPEC} 加载失败：${error instanceof Error ? error.message : String(error)}` };
        return kitResolution;
    }
}

/** 测试用：清掉解析缓存（改了环境变量或装了引擎之后重新解析）。 */
export function resetPreviewKit() {
    kitResolution = undefined;
}

/**
 * 同步判断渲染引擎包能不能解析到（能力地图用，见 src/capabilities.js）。
 *
 * 不加载引擎、不启动 LibreOffice，也不用 async：解析锚点链本来就是同步的
 * （`resolveKitPath` 只走 createRequire）。已经加载过时直接用那次的结果，
 * 免得「加载失败」与「解析不到」在两处给出不一致的答案。
 */
export function previewKitAvailable() {
    if (kitResolution !== undefined) return kitResolution.error === undefined;
    return resolveKitPath() !== undefined;
}

function supportedExtensions(module) {
    const formats = module?.IMAGE_FORMATS;
    if (Array.isArray(formats) && formats.length > 0) return formats.map((item) => `.${String(item).replace(/^\./, '')}`);
    return [...PREVIEW_EXTENSIONS];
}

/**
 * 能力探针：能不能预览、用什么引擎、支持哪些输入。不启动引擎。
 * @returns {Promise<object>}
 */
export async function previewCheck() {
    const kit = await loadPreviewKit();
    if (kit.error !== undefined) {
        return {
            available: false,
            engine: PREVIEW_KIT_SPEC,
            formats: [...PREVIEW_EXTENSIONS],
            hint: kit.error,
        };
    }
    const value = {
        available: true,
        engine: PREVIEW_KIT_SPEC,
        formats: supportedExtensions(kit.module),
        defaultDpi: PREVIEW_DEFAULT_DPI,
        maxPages: PREVIEW_MAX_PAGES,
        hint: 'office.preview.render(path) 把自家文档渲染成页面图，逐张交给 read_image 看。',
    };
    // discoverRuntime 只读安装元数据，不启动引擎；拿不到也不影响可用性。
    try {
        if (typeof kit.module.discoverRuntime === 'function') {
            const runtime = await kit.module.discoverRuntime();
            value.backend = runtime?.backend;
            value.version = runtime?.version;
        } else if (kit.module.ENGINE_VERSION !== undefined) {
            value.version = String(kit.module.ENGINE_VERSION);
        }
    } catch (error) {
        value.runtimeError = error instanceof Error ? error.message : String(error);
    }
    return value;
}

function previewError(message) {
    const error = new Error(message);
    error.code = 'OFFICE_PREVIEW_UNAVAILABLE';
    return error;
}

/** 校验输入：存在、扩展名在支持清单里。返回绝对路径与扩展名。 */
function resolveInput(filePath, env, module) {
    const absolute = isAbsolute(filePath) ? resolve(filePath) : env.resolve(filePath);
    if (!existsSync(absolute)) throw previewError(`office.preview：找不到文件 ${filePath}`);
    const info = statSync(absolute);
    if (!info.isFile()) throw previewError(`office.preview：${filePath} 不是文件。`);
    const at = absolute.lastIndexOf('.');
    const ext = at === -1 ? '' : absolute.slice(at).toLowerCase();
    const supported = supportedExtensions(module);
    if (!supported.includes(ext)) {
        throw previewError(`office.preview：${ext || '(没有扩展名)'} 不在支持的清单里（${supported.join(' / ')}）。`);
    }
    return { absolute, ext, bytes: info.size, mtimeMs: info.mtimeMs };
}

/**
 * 内容键：文件内容 + dpi + 页选择。
 *
 * 用**内容**而不是 mtime：同一份文件被复制 / 重命名后仍然命中，而「改完再预览」
 * 必然换键（改了内容）。文件读一次是有意的 —— 它同时把这份输入登记进台账的
 * `source`（预览的依据是哪一份文件）。
 */
function contentKey(input, env, extra) {
    const data = env.readFile(input.absolute);
    const hash = createHash('sha1').update(data).digest('hex').slice(0, 12);
    return `${hash}-${extra}`;
}

function stemOf(absolute) {
    const base = absolute.replace(/^.*[\\/]/, '');
    const at = base.lastIndexOf('.');
    return (at === -1 ? base : base.slice(0, at)).replace(/[\\/:*?"<>|\s]+/g, '-').slice(0, 60) || 'preview';
}

/**
 * 从 PNG 的 IHDR 现读像素尺寸（缓存命中时也要给出宽高，否则「命中」与「新渲染」
 * 两条路的返回形状不一样，调用方按 width 排版就会在命中时拿到 undefined）。
 * 读不出来返回 undefined —— 不编一个尺寸出来。
 */
function readPngSize(path) {
    let fd;
    try {
        fd = openSync(path, 'r');
        const head = Buffer.alloc(24);
        const read = readSync(fd, head, 0, 24, 0);
        if (read < 24) return undefined;
        if (head[0] !== 0x89 || head[1] !== 0x50 || head[2] !== 0x4e || head[3] !== 0x47) return undefined;
        return { width: head.readUInt32BE(16), height: head.readUInt32BE(20) };
    } catch {
        return undefined;
    } finally {
        if (fd !== undefined) {
            try {
                closeSync(fd);
            } catch {
                // 关不掉不影响结果
            }
        }
    }
}

/** 页选择 → 引擎认的形状：'all' 或一基页码数组。 */
function normalizePages(raw) {
    if (raw === undefined || raw === null || raw === '' || raw === 'all') return 'all';
    const list = Array.isArray(raw) ? raw : [raw];
    const pages = [];
    for (const item of list) {
        const n = Math.trunc(Number(item));
        if (!Number.isFinite(n) || n < 1) {
            throw previewError(`office.preview：pages 只能给 'all' 或从 1 开始的页码（收到 ${JSON.stringify(item)}）。`);
        }
        pages.push(n);
    }
    if (pages.length === 0) return 'all';
    return [...new Set(pages)].sort((left, right) => left - right);
}

/**
 * 把自家文档渲染成页面图。
 *
 * @param {string} filePath 相对工作目录或绝对路径
 * @param {{pages?: 'all'|number|number[], sheet?: string, range?: string, dpi?: number, maxPages?: number}} [options]
 * @param {object} env 引擎环境
 * @param {object} cache 缓存目录
 * @returns {Promise<object>} images[].path 直接交给 read_image
 */
export async function previewImages(filePath, options = {}, env, cache) {
    const kit = await loadPreviewKit();
    if (kit.error !== undefined) throw previewError(`office.preview：${kit.error}`);
    const input = resolveInput(filePath, env, kit.module);
    const pages = normalizePages(options.pages);
    const sheet = typeof options.sheet === 'string' && options.sheet.trim() !== '' ? options.sheet.trim() : undefined;
    if (sheet !== undefined && pages !== 'all') {
        throw previewError('office.preview：pages 与 sheet 不能同时给（工作表按 sheet / range 定位，不按页码）。');
    }
    // 参数与输入族要对得上：xlsx 按 sheet 定位、docx/pptx 按页码 —— 反着给时引擎会抛一句
    // 与文件无关的 TypeError，这里先拦下来，错误才指得对地方。
    const isWorkbook = ['.xlsx', '.xls', '.ods'].includes(input.ext);
    if (isWorkbook && pages !== 'all') {
        throw previewError(`office.preview：${input.ext} 是工作簿，用 sheet（工作表名）配 range 定位，不用 pages。`);
    }
    if (!isWorkbook && sheet !== undefined) {
        throw previewError(`office.preview：${input.ext} 不是工作簿，用 pages（页码）定位，不用 sheet。`);
    }
    const requestedDpi = Math.trunc(Number(options.dpi) || PREVIEW_DEFAULT_DPI);
    const dpi = Math.min(600, Math.max(24, requestedDpi));
    const maxPages = Math.min(PREVIEW_MAX_PAGES, Math.max(1, Math.trunc(Number(options.maxPages) || 12)));
    const range = typeof options.range === 'string' && options.range.trim() !== '' ? options.range.trim() : undefined;
    const key = contentKey(input, env, `${dpi}-${pages === 'all' ? 'all' : pages.join('_')}-${sheet ?? ''}-${range ?? ''}-${maxPages}`);
    const name = `preview/${stemOf(input.absolute)}-${key}`;
    const dir = cache.locate(name);
    const manifestPath = join(dir, 'manifest.json');

    /** 命中已有渲染：目录里挑出 page-*.png 按序号排好（尺寸从 IHDR 现读）。 */
    const existing = () => {
        if (!existsSync(dir)) return undefined;
        const files = readdirSync(dir).filter((entry) => /^page-\d+\.png$/i.test(entry)).sort();
        if (files.length === 0) return undefined;
        // 有 manifest 才算命中：渲染到一半被中断时目录里会剩几张图，
        // 那时把它当完整结果返回会少报页数（manifest 里的张数是完成时的真值）。
        const manifest = readManifest(manifestPath);
        if (manifest === undefined || manifest.count !== files.length) return undefined;
        return {
            manifest,
            images: files.map((entry, index) => {
                const path = join(dir, entry);
                const info = statSync(path);
                const size = readPngSize(path);
                return {
                    index: index + 1,
                    page: index + 1,
                    path: displayPath(env.root, path),
                    bytes: info.size,
                    width: size?.width,
                    height: size?.height,
                };
            }),
        };
    };

    const hit = existing();
    if (hit !== undefined) {
        for (const image of hit.images) cache.noteHit(image.bytes);
        const manifest = hit.manifest;
        const truncated = manifest.pageCount > hit.images.length;
        return {
            ok: true,
            kind: 'images',
            path: displayPath(env.root, input.absolute),
            dir: displayPath(env.root, dir),
            backend: manifest.backend,
            rasterEngine: manifest.rasterEngine,
            dpi,
            pages,
            pageCount: manifest.pageCount,
            count: hit.images.length,
            images: hit.images,
            missingFonts: manifest.missingFonts ?? [],
            reused: true,
            hint: truncated
                ? `这份文件共 ${manifest.pageCount} 页，上次渲染了 ${hit.images.length} 张`
                    + `（${manifest.truncatedBy === 'pages' ? '按 pages 选的页' : `maxPages=${maxPages}`}）。`
                    + '要看别的页传 pages（如 pages: [3,4]）或放大 maxPages。'
                : '把 images[].path 逐张交给 read_image 看；这是上次渲染的结果（内容未变，命中缓存）。',
        };
    }
    // 目录必须**不存在**才给引擎（它拒绝写已存在的输出目录）：半成品先清掉。
    if (existsSync(dir)) rmSync(dir, { recursive: true, force: true });
    mkdirSync(dirname(dir), { recursive: true });

    const converter = await kit.module.createConverter({ timeoutMs: 120_000 });
    let result;
    let truncatedBy;
    try {
        result = await renderWithPageLimit(converter, {
            inputPath: input.absolute,
            outputDir: dir,
            pages,
            sheet,
            range,
            dpi,
            maxPages,
        }, () => { truncatedBy = 'maxPages'; });
    } catch (error) {
        throw previewError(`office.preview：渲染失败（${input.absolute}）：${error instanceof Error ? error.message : String(error)}`);
    } finally {
        try {
            await converter.dispose();
        } catch {
            // 引擎收尾失败不影响已经渲染出来的图
        }
    }
    const images = (result.images ?? []).map((item) => ({
        index: item.index,
        page: item.page,
        sheet: item.sheet,
        width: item.width,
        height: item.height,
        bytes: item.byteLength,
        path: displayPath(env.root, item.path),
    }));
    for (const image of images) cache.noteArtifact();
    const truncated = result.pageCount > images.length;
    // manifest 是「这次渲染完成了」的标记：下一次命中靠它（见 existing()）。
    writeManifest(manifestPath, {
        count: images.length,
        pageCount: result.pageCount,
        backend: result.backend,
        rasterEngine: result.rasterEngine,
        missingFonts: result.missingFonts ?? [],
        dpi,
        pages,
        ...(truncated ? { truncatedBy: truncatedBy ?? (pages === 'all' ? 'maxPages' : 'pages') } : {}),
    });
    return {
        ok: true,
        kind: 'images',
        path: displayPath(env.root, input.absolute),
        dir: displayPath(env.root, dir),
        backend: result.backend,
        rasterEngine: result.rasterEngine,
        dpi,
        pages,
        pageCount: result.pageCount,
        count: images.length,
        images,
        missingFonts: result.missingFonts ?? [],
        reused: false,
        hint: truncated
            ? `这份文件共 ${result.pageCount} 页，这次只渲染了 ${images.length} 张`
                + `（${truncatedBy === 'maxPages' ? `maxPages=${maxPages}` : '按 pages 选的页'}）。`
                + '要看别的页传 pages（如 pages: [3,4]）或放大 maxPages。'
            : '把 images[].path 逐张交给 read_image 看；渲染结果是中间产物，落在缓存目录。',
    };
}

/**
 * 渲染，并在「页数超过 maxPages」时**降级成只渲前 maxPages 页**。
 *
 * 引擎的行为是抛 `output-too-large`（"Selected page count exceeds maxPages"）而不是截断 ——
 * 于是默认调用（pages:'all' + maxPages:12）在一份 17 页的 pptx 上直接失败，而
 * 「预览前 12 页」正是这个 API 想要的语义。这里接住那一类错误、换成显式页列表重试一次
 * （引擎在抛之前不会写任何 PNG，所以重试是干净的）。
 */
async function renderWithPageLimit(converter, request, onTruncated) {
    const { pages, sheet, range, maxPages, ...rest } = request;
    const base = {
        ...rest,
        ...(sheet === undefined ? {} : { sheet }),
        ...(range === undefined ? {} : { range }),
        maxPages,
    };
    try {
        return await converter.renderImages({ ...base, ...(pages === 'all' ? {} : { pages }) });
    } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        if (pages !== 'all' || !/maxPages|output-too-large|exceeds/i.test(message)) throw error;
        onTruncated();
        const first = Array.from({ length: maxPages }, (_item, index) => index + 1);
        return await converter.renderImages({ ...base, pages: first });
    }
}

/** 写渲染清单（manifest.json）：命中时要靠它还原完整返回形状，也当「这次渲染完成了」的标记。 */
function writeManifest(path, value) {
    try {
        writeFileSync(path, JSON.stringify(value), 'utf8');
        return true;
    } catch {
        return false;
    }
}

function readManifest(path) {
    try {
        if (!existsSync(path)) return undefined;
        const value = JSON.parse(readFileSync(path, 'utf8'));
        return value !== null && typeof value === 'object' && Number.isFinite(Number(value.count)) ? value : undefined;
    } catch {
        return undefined;
    }
}

/**
 * 把自家文档转成 PDF（打印 / 交付用）。
 *
 * 两条落点：缓存目录里那份按内容键存（重复转换命中，不重算），
 * `to` 给出时再把字节经 `env.writeFile` 写到工作目录 —— 走 env 的理由是
 * 「写出的文件要出现在反馈的 otherFiles 里」：预览的 PDF 也是产物，不该只躺在缓存里
 * 而反馈上一个字都没有。
 *
 * @param {string} filePath
 * @param {{to?: string, overwrite?: boolean}} [options] `to` 给出时把 PDF 写到工作目录（交付物）。
 */
export async function previewPdf(filePath, options = {}, env, cache) {
    const kit = await loadPreviewKit();
    if (kit.error !== undefined) throw previewError(`office.preview：${kit.error}`);
    const input = resolveInput(filePath, env, kit.module);
    const wanted = typeof options.to === 'string' && options.to.trim() !== '' ? options.to.trim() : '';
    if (wanted !== '') {
        // 把 PDF 写到源文件自己的路径上 = 把原文档毁掉（引擎自己也会拒绝这种输出，
        // 但这里是我们写盘，所以这道闸必须在写之前）。
        const candidate = isAbsolute(wanted) ? resolve(wanted) : resolve(env.root, wanted);
        if (candidate === input.absolute) {
            throw previewError(`office.preview：to 不能和输入文件是同一个（${wanted}）—— 那会把原文档覆盖成 PDF。`);
        }
    }
    const key = contentKey(input, env, 'pdf');
    const cached = cache.locate(`preview/${stemOf(input.absolute)}-${key}.pdf`);

    let hit = false;
    if (existsSync(cached)) {
        hit = true;
        cache.noteHit(statSync(cached).size);
    } else {
        mkdirSync(dirname(cached), { recursive: true });
        const converter = await kit.module.createConverter({ timeoutMs: 120_000 });
        try {
            await converter.render({ inputPath: input.absolute, outputPath: cached });
        } catch (error) {
            throw previewError(`office.preview：转 PDF 失败（${input.absolute}）：${error instanceof Error ? error.message : String(error)}`);
        } finally {
            try {
                await converter.dispose();
            } catch {
                // 收尾失败不影响已经写出的 PDF
            }
        }
        cache.noteArtifact();
    }
    const bytes = readFileSync(cached);

    if (wanted === '') {
        return {
            ok: true,
            kind: 'pdf',
            path: displayPath(env.root, cached),
            bytes: bytes.length,
            reused: hit,
            hint: '这份 PDF 在缓存目录里（中间产物）。要当交付物就传 to（写到工作目录）。',
        };
    }
    // 目标已经是一模一样的内容就不重写（不改 mtime、不重复登记）。
    const target = assertInsideRoot(wanted, env, 'office.preview.pdf');
    const existedBefore = existsSync(target);
    if (existedBefore) {
        const current = readFileSync(target);
        if (current.length === bytes.length && current.equals(bytes)) {
            return {
                ok: true,
                kind: 'pdf',
                path: displayPath(env.root, target),
                bytes: bytes.length,
                reused: true,
                hint: '目标文件已经是这次转换的结果（内容一致），没有重写。',
            };
        }
        // 内容不同的既有文件默认**不覆盖**：与 office.image.fetch / archive.extract 同一口径
        // （覆盖别人的文件要显式说一声）。
        if (options.overwrite !== true) {
            throw previewError(`office.preview：${wanted} 已存在且内容不同；换一个 to，或显式传 overwrite:true。`);
        }
    }
    const written = env.writeFile(target, bytes);
    return {
        ok: true,
        kind: 'pdf',
        path: written.path,
        bytes: written.bytes,
        reused: false,
        ...(existedBefore ? { overwrote: true } : {}),
        hint: existedBefore
            ? 'PDF 已重写到工作目录（原文件内容不同，按 overwrite:true 覆盖）。'
            : 'PDF 已写到工作目录，可以直接交付。',
    };
}
