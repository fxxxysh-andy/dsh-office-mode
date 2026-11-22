/**
 * 预览面板的四条只读路由：文件清单 / 大纲 / 单页图 / 文本原文。
 *
 * 为什么单独一个模块：记忆面板回答「模型与我看到的是不是同一份记忆」，预览面板回答
 * 「这份产物到底长什么样」—— 数据源、耗时（渲染一次 2～5 秒）与错误口径都不同，
 * 混进 view.js 只会让那边的快照逻辑越来越难读。信任边界（cwd 白名单）与 JSON 写回
 * 直接复用 view.js 导出的那一套，绝不在这里另起一份：预览能列出的文件比记忆更多，
 * 放宽一处就会在另一处静默失效。
 *
 * 三条纪律：
 *   1. 只读：GET/HEAD，渲染结果只进缓存目录，不碰工作目录；文本原文只读前
 *      2 MiB，同样不写任何东西；
 *   2. 任何错误都带「下一步怎么做」，面板照原样展示，不让用户对着一句报错猜；
 *   3. 引擎是可选能力：没装引擎时清单与大纲（纯解析）照常可用，只有取图返回 503。
 *
 * @module dsh-office-mode/preview-view
 */
import { closeSync, existsSync, openSync, readdirSync, readFileSync, readSync, statSync } from 'node:fs';
import { basename, extname, isAbsolute, join, relative as relativePath, resolve } from 'node:path';

import { createCache } from './engine/cache.js';
import { assertInsideRoot, createEnv } from './engine/kit.js';
import { read as readPptx } from './formats/ppt.js';
import {
    PREVIEW_DEFAULT_DPI,
    PREVIEW_EXTENSIONS,
    loadPreviewKit,
    previewImages,
} from './preview.js';
import { publicWorkspaces, sendJson, trustedRootFor } from './view.js';

/** 文本原文路由的路径（单独导出，面板侧拼 URL 用）。 */
export const TEXT_VIEW_PATH = '/office-preview/text';

/** 四条路由的路径（面板侧拼 URL 用，测试也照这里断言）。 */
export const PREVIEW_VIEW_PATHS = Object.freeze({
    files: '/office-preview/files',
    outline: '/office-preview/outline',
    page: '/office-preview/page',
    text: TEXT_VIEW_PATH,
});

/** 文件清单最多给几条：面板左列只有一屏，多了也没人翻。 */
export const PREVIEW_FILE_LIMIT = 20;

/** 扫描深度：根 + 两层子目录（`out/`、`out/sub/` 里的产物都算，再深就不看）。 */
const SCAN_DEPTH = 2;

/** 要列出来的扩展名：能渲染的 + 办公模式会产出的纯文本。 */
const OFFICE_EXTENSIONS = new Set([...PREVIEW_EXTENSIONS, '.md', '.csv', '.tex']);

/**
 * 能按原文预览的文本扩展名。
 *
 * 这里只放纯文本：二进制格式（.pptx/.docx/.xlsx/.pdf/图片/压缩包）走渲染通道，
 * 混进来只会把乱码当正文给用户看。清单口径（OFFICE_EXTENSIONS）不在这里增补 ——
 * 它决定左列列什么，改了就会把真正的交付物挤出 20 条上限。
 */
export const TEXT_EXTENSIONS = Object.freeze([
    '.md', '.markdown', '.mdx', '.csv', '.tsv', '.tex', '.txt', '.log',
    '.json', '.yaml', '.yml', '.ini', '.toml', '.xml', '.html', '.css',
    '.js', '.ts', '.py', '.sh', '.c', '.h', '.java', '.go', '.rs',
]);

/** 文本读取上限：2 MiB。超过只读前 2 MiB，并按 UTF-8 字符边界切。 */
export const TEXT_READ_LIMIT = 2 * 1024 * 1024;

function isTextExt(ext) {
    return TEXT_EXTENSIONS.includes(ext);
}

/** dpi 的可调范围：低于 24 看不清字，高于 300 一张图能到十几 MB。 */
const DPI_MIN = 24;
const DPI_MAX = 300;

/**
 * 一条会被路由改成 HTTP 状态码的失败。
 *
 * 用异常而不是到处 return：嵌套在扫描 / 解析 / 渲染里的校验（越权、扩展名、页码）
 * 都长在同一层，抛出来由外层统一写回，错误文案才不会因为调用点不同而漂移。
 */
class PreviewRouteError extends Error {
    constructor(status, message, hint) {
        super(message);
        this.status = status;
        if (typeof hint === 'string' && hint !== '') this.hint = hint;
    }
}

/** 扩展名的「这一份要不要渲染」口径，与 preview.js 的 PREVIEW_EXTENSIONS 一致。 */
function isRenderable(ext) {
    return PREVIEW_EXTENSIONS.includes(ext);
}

function kindOf(ext) {
    if (ext === '.pptx' || ext === '.ppt') return 'pptx';
    if (isRenderable(ext)) return 'renderable';
    return 'text';
}

/**
 * 把 `?path=` 收敛成 cwd 内的绝对路径。
 *
 * 只收相对路径：清单给的是相对路径，绝对路径既对不上清单，也会让「cwd 内」这条
 * 边界变成摆设。越界统一报 404 —— 与「不在白名单的 cwd」同一个状态码，不额外
 * 告诉调用方「这个路径存在、只是不给你看」。
 */
function readPathFor(root, rawPath) {
    const text = String(rawPath ?? '').trim();
    if (text === '') {
        throw new PreviewRouteError(400, 'path 不能为空：先在左侧文件列表里选一份产物。');
    }
    if (isAbsolute(text)) {
        throw new PreviewRouteError(404, 'path 要写相对 cwd 的路径（例如 out/deck.pptx），不接受绝对路径。');
    }
    try {
        return assertInsideRoot(text, { root }, '预览');
    } catch {
        throw new PreviewRouteError(404, `path 要落在工作目录内：${text.replace(/\\/g, '/')} 会跑到 cwd 之外。`);
    }
}

function requireExisting(absolute, relative) {
    if (!existsSync(absolute)) {
        throw new PreviewRouteError(404, `文件不存在：${relative}。可能刚被删掉或改过名，刷新文件列表再选。`);
    }
}

/**
 * 深度 2 的产物清单。
 *
 * 只按扩展名与 mtime 判断，不读文件内容：打开面板就要立刻列出东西，读内容会把
 * 一次面板挂载变成几十次磁盘 IO。跳过隐藏目录（`.git` / `.office`）与依赖目录，
 * 它们是「实现」不是「产物」，列出来只会把真交付物挤出 20 条上限。
 */
function scanOfficeFiles(root) {
    const found = [];
    const walk = (dir, depth, prefix) => {
        let entries;
        try {
            entries = readdirSync(dir, { withFileTypes: true });
        } catch {
            return;
        }
        for (const entry of entries) {
            if (entry.name.startsWith('.')) continue;
            if (entry.isDirectory() && ['node_modules', 'dist', 'build', 'coverage'].includes(entry.name)) continue;
            const relative = prefix === '' ? entry.name : `${prefix}/${entry.name}`;
            const absolute = join(dir, entry.name);
            if (entry.isDirectory()) {
                if (depth < SCAN_DEPTH) walk(absolute, depth + 1, relative);
                continue;
            }
            if (!entry.isFile()) continue;
            const ext = extname(entry.name).toLowerCase();
            if (!OFFICE_EXTENSIONS.has(ext)) continue;
            let info;
            try {
                info = statSync(absolute);
            } catch {
                continue;
            }
            found.push({
                path: relative,
                name: entry.name,
                ext,
                bytes: info.size,
                mtime: new Date(info.mtimeMs).toISOString(),
                mtimeMs: info.mtimeMs,
                kind: kindOf(ext),
            });
        }
    };
    walk(root, 0, '');
    found.sort((left, right) => right.mtimeMs - left.mtimeMs);
    return found;
}

/**
 * 大纲行 → 面板要的 `slides[{page,layout,title}]`。
 *
 * `read()` 的 outline 是给模型看的一行文字，两种写法并存：
 *   `P3 bullets：本季重点｜4 个文本框`  —— 标题在「｜」前（bullets/table/image/助手页）；
 *   `P7 shape：12 个形状`               —— 自由绘制页没有标题位（页码后面直接是形状计数）。
 * 因此标题只取「｜」前那一段，`shape` 行一律当无标题；认不出格式的行不编页码，
 * 留 page:0 让面板过滤掉，而不是把整行塞进标题里。
 */
function slidesFromOutline(outline, pages) {
    const slides = [];
    for (const raw of Array.isArray(outline) ? outline : []) {
        const matched = /^P(\d+)\s+(\S+?)：([\s\S]*)$/.exec(String(raw).trim());
        if (matched === null) continue;
        const page = Number(matched[1]);
        const layout = matched[2];
        const body = matched[3];
        const bar = body.indexOf('｜');
        const title = layout === 'shape' ? '' : (bar === -1 ? body : body.slice(0, bar)).trim();
        slides.push({ page, layout, title });
    }
    // 页码以 pages 为准时补一段：outline 偶尔漏行（部件缺失会只进 warnings），
    // 面板左列的页码不能因此少一页。
    for (const page of Array.isArray(pages) ? pages : []) {
        if (slides.some((item) => item.page === page.index)) continue;
        slides.push({ page: page.index, layout: page.layout, title: page.title ?? '' });
    }
    slides.sort((left, right) => left.page - right.page);
    return slides;
}

/** 版式分布 → 面板的徽标统计（只留出现过的版式，按多到少）。 */
function layoutStats(layouts) {
    return Object.entries(layouts ?? {})
        .filter(([, count]) => count > 0)
        .sort((left, right) => right[1] - left[1])
        .map(([layout, count]) => ({ layout, count }));
}

/** 大纲响应：pptx 走纯解析（不需要引擎），其余格式如实说「没有页结构」。 */
async function outlinePayload({ root, label, relative, absolute, deps }) {
    const ext = extname(absolute).toLowerCase();
    if (ext !== '.pptx') {
        // 文本类文件（.md/.csv/.tex…）另有一条「正文」通道：面板据此改走 text 路由；
        // renderable 仍为 false —— 它们不是取图格式，取图端点照旧 400。
        const isText = isTextExt(ext);
        return {
            ok: true,
            root,
            label,
            path: relative,
            format: ext.replace(/^\./, ''),
            pages: null,
            slides: [],
            outline: [],
            layouts: [],
            warnings: [],
            renderable: isRenderable(ext),
            ...(isText ? { text: true } : {}),
            hint: isText
                ? '文本文件按原文预览（右列显示正文）'
                : isRenderable(ext)
                    ? '这份格式没有逐页大纲，但可以按页取图看版式。'
                    : '这种格式既没有页结构也不能渲染成图，预览面板里只做文件清单。',
        };
    }
    let report;
    try {
        report = await deps.readPpt(absolute, createEnv({ root }));
    } catch (error) {
        throw new PreviewRouteError(400, `读不出大纲：${error?.message ?? error}。`
            + '文件可能不是有效的 .pptx，或已经损坏；重新生成一份再试。');
    }
    const pages = Array.isArray(report?.pages) ? report.pages : [];
    const outline = Array.isArray(report?.outline) ? report.outline : [];
    return {
        ok: true,
        root,
        label,
        path: relative,
        format: 'pptx',
        pages: pages.length,
        slides: slidesFromOutline(outline, pages),
        outline,
        layouts: layoutStats(report?.stats?.layouts),
        warnings: Array.isArray(report?.warnings) ? report.warnings : [],
        theme: report?.theme,
        size: report?.size,
        stats: {
            slides: pages.length,
            shapes: report?.stats?.shapes ?? 0,
            textBoxes: report?.stats?.textBoxes ?? 0,
            images: report?.stats?.images ?? 0,
            tables: report?.stats?.tables ?? 0,
            charts: report?.stats?.charts ?? 0,
            words: report?.stats?.words ?? 0,
            notes: report?.stats?.notes ?? 0,
        },
        renderable: true,
    };
}

/** 取图响应：渲染单页并把 PNG 读成 data URL（面板直接塞 <img>，不用再发一次请求）。 */
async function pagePayload({ root, label, relative, absolute, pageNo, dpi, deps }) {
    const ext = extname(absolute).toLowerCase();
    if (!isRenderable(ext)) {
        throw new PreviewRouteError(400, `这种格式不渲染成页面图：${ext || '（无扩展名）'}。`
            + `可渲染的是 ${PREVIEW_EXTENSIONS.join(' / ')}；纯文本请直接打开文件看。`);
    }
    const kit = await deps.loadEngine();
    if (kit?.error !== undefined) {
        throw new PreviewRouteError(503, `渲染引擎不可用：${kit.error}`
            + '取图能力依赖宿主随部署安装的引擎，其它格式与清单不受影响；'
            + '装好引擎后刷新这个面板即可（不用重启 DSH）。');
    }

    // pptx 的页数是纯解析出来的，先拿它判越界：越界时报「共几页」比让引擎抛一句
    // 跟文件无关的 TypeError 有用得多。其余格式的页数只有引擎知道，交给它报。
    let pageCount = null;
    if (ext === '.pptx') {
        try {
            const report = await deps.readPpt(absolute, createEnv({ root }));
            pageCount = Array.isArray(report?.pages) ? report.pages.length : null;
        } catch {
            pageCount = null;
        }
    }
    if (pageCount !== null && pageNo > pageCount) {
        throw new PreviewRouteError(400, `第 ${pageNo} 页不存在：这份演示文稿共 ${pageCount} 页。`
            + `下一步：把 page 改成 1–${pageCount} 之间的页码。`);
    }

    const env = createEnv({ root });
    const cache = createCache({ root, dir: deps.cacheDir });
    let rendered;
    try {
        rendered = await deps.render(absolute, { pages: [pageNo], dpi, maxPages: 1 }, env, cache);
    } catch (error) {
        throw new PreviewRouteError(502, `渲染第 ${pageNo} 页失败：${error?.message ?? error}`);
    }
    const image = Array.isArray(rendered?.images) ? rendered.images[0] : undefined;
    if (image === undefined) {
        throw new PreviewRouteError(502, `引擎没有为第 ${pageNo} 页出图（返回的页数是 ${rendered?.pageCount ?? '未知'}）。`
            + '先刷新大纲确认页码，再试一次。');
    }
    const imagePath = env.resolve(image.path);
    let png;
    try {
        png = readFileSync(imagePath);
    } catch (error) {
        throw new PreviewRouteError(500, `渲染图读不回来：${error?.message ?? error}。`
            + '清空缓存目录后重试一次。');
    }
    return {
        ok: true,
        root,
        label,
        path: relative,
        page: pageNo,
        pages: rendered?.pageCount ?? pageCount ?? null,
        dpi: rendered?.dpi ?? dpi,
        width: image.width,
        height: image.height,
        bytes: png.length,
        cached: rendered?.reused === true,
        png: `data:image/png;base64,${png.toString('base64')}`,
    };
}

/**
 * 按 UTF-8 字符边界丢掉被切断的尾巴。
 *
 * 只在前 2 MiB 处截断时用：2 MiB 的切片末尾可能落在多字节字符中间，直接解码会
 * 多出一个替换字符（用户看到正文最后是乱码）。从末尾回退最多 3 个续字节
 * （0b10xxxxxx）找到首字节，算出的字符宽度超出已读范围就把这个残缺字符整段丢弃。
 */
function trimUtf8Tail(buffer) {
    let back = 0;
    while (back < 3 && buffer.length - back > 0 && (buffer[buffer.length - back - 1] & 0xC0) === 0x80) {
        back += 1;
    }
    const leadAt = buffer.length - back - 1;
    if (leadAt < 0) return buffer;
    const lead = buffer[leadAt];
    const width = lead >= 0xF0 ? 4 : lead >= 0xE0 ? 3 : lead >= 0xC0 ? 2 : 1;
    return width > back + 1 ? buffer.subarray(0, leadAt) : buffer;
}

/** 读文件最前面的若干字节（最多 TEXT_READ_LIMIT）：大文件不整个读进内存。 */
function readHeadBytes(absolute, size) {
    const take = Math.min(size, TEXT_READ_LIMIT);
    if (take <= 0) return Buffer.alloc(0);
    const fd = openSync(absolute, 'r');
    try {
        const buffer = Buffer.allocUnsafe(take);
        let offset = 0;
        while (offset < take) {
            const read = readSync(fd, buffer, offset, take - offset, offset);
            if (read <= 0) break;
            offset += read;
        }
        return buffer.subarray(0, offset);
    } finally {
        closeSync(fd);
    }
}

/** 「大量替换字符」的判据：至少 8 个、且占解码文本 1% 以上（一两个则可能是原文本来就有）。 */
function looksBinary(text, replacements) {
    return replacements >= 8 && replacements * 100 > text.length;
}

/** 文本端点的状态码 → 「下一步怎么办」（共用 helper 抛出的 400/404 也带得上）。 */
const TEXT_ERROR_HINTS = Object.freeze({
    400: '下一步：给 path（cwd 下的相对路径，例如 docs/report.md），再请求一次。',
    404: '下一步：刷新一次文件列表再从里面选；cwd 要是会话的工作目录，没登记过的先在那个目录跑一次办公工具。',
    405: '下一步：这条端点只读，改用 GET（或 HEAD）请求。',
    415: '下一步：文本以外的格式用渲染通道看（先取大纲、再按页取图）。',
    500: '下一步：刷新文件列表再选一次；仍失败就确认文件没有被占用或删除。',
});

/**
 * 文本原文响应：扩展名闸 → 读上限 → UTF-8 解码 → text/lines/hint。
 *
 * `bytes` 一律是文件的真实大小（不是这次读到的字节数）：面板要显示「这份文件多大」，
 * 截断时 hint 里的「共 X MiB」也按它算。二进制被误标成文本扩展名时强行解码，
 * 返回 200 + hint，不 500 —— 面板总得有东西可显示，用户才能判断该不该走渲染通道。
 */
function textPayload({ root, label, relative: relativeRaw, absolute }) {
    const ext = extname(absolute).toLowerCase();
    if (!isTextExt(ext)) {
        throw new PreviewRouteError(415, `这种格式不支持原文预览：${ext === '' ? '（无扩展名）' : ext}。`,
            '这种格式用 PPT/文档渲染通道看：先取大纲，再按页取图。');
    }
    let info;
    try {
        info = statSync(absolute);
    } catch (error) {
        if (error?.code === 'ENOENT') {
            throw new PreviewRouteError(404, `文件不存在：${relativeRaw}。可能刚被删掉或改过名。`,
                '刷新文件列表再选一份。');
        }
        throw new PreviewRouteError(500, `读不了文件信息：${error?.message ?? error}。`,
            '刷新文件列表再选一次；仍失败就确认文件没有被别的程序占用。');
    }
    if (!info.isFile()) {
        throw new PreviewRouteError(404, `这不是文件（可能是目录）：${relativeRaw}。`,
            '从左侧文件列表里选一份文件；目录不支持原文预览。');
    }
    const bytes = info.size;
    const truncated = bytes > TEXT_READ_LIMIT;
    let head;
    try {
        head = readHeadBytes(absolute, bytes);
    } catch (error) {
        throw new PreviewRouteError(500, `读文件失败：${error?.message ?? error}。`,
            '刷新文件列表再选一次；仍失败就确认文件没有被删除或占用。');
    }
    const buffer = truncated ? trimUtf8Tail(head) : head;
    const text = new TextDecoder('utf-8', { fatal: false }).decode(buffer);
    const replacements = text.split('\uFFFD').length - 1;
    const hints = [];
    if (truncated) {
        hints.push(`文件较大，只预览前 2 MiB（共 ${(bytes / 1024 / 1024).toFixed(1)} MiB）`);
    }
    if (looksBinary(text, replacements)) {
        hints.push('疑似二进制文件，按 UTF-8 强行解码');
    }
    return {
        ok: true,
        root,
        label,
        path: relativePath(root, absolute).replace(/\\/g, '/'),
        name: basename(absolute),
        ext,
        kind: 'text',
        bytes,
        text,
        truncated,
        lines: text.split('\n').length,
        encoding: 'utf-8',
        ...(hints.length === 0 ? {} : { hint: hints.join('；') }),
    };
}

/** 与 view.js 的 sendJson 同一套响应头；文本端点的 405 还要带 `Allow`（sendJson 收不了额外头）。 */
function sendJsonWithHeaders(res, status, payload, headers) {
    const body = `${JSON.stringify(payload)}\n`;
    res.writeHead(status, {
        'content-type': 'application/json; charset=utf-8',
        'content-length': Buffer.byteLength(body, 'utf8'),
        'cache-control': 'no-store',
        ...headers,
    });
    res.end(body);
}

/**
 * 四个处理函数共用的「方法检查 + 信任根 + 错误写回」外壳。
 *
 * `options.allowHeader`：405 响应带上 Allow（文本端点契约要求，既有三条保持原样）。
 * `options.hints`：按状态码补一句「下一步怎么做」（文本端点用；不给时错误体与之前完全一致）。
 */
function routeHandler(name, body, trustFn, options = {}) {
    const allowHeader = typeof options.allowHeader === 'string' ? options.allowHeader : null;
    const hints = options.hints !== null && typeof options.hints === 'object' ? options.hints : null;
    return async (req, res) => {
        try {
            const method = String(req.method ?? 'GET').toUpperCase();
            if (method !== 'GET' && method !== 'HEAD') {
                throw new PreviewRouteError(405, '预览端点只读：只支持 GET / HEAD。');
            }
            const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
            const root = await trustFn(url);
            if (!root.ok) throw new PreviewRouteError(root.status, root.error);
            const payload = await body(url, root);
            sendJson(res, 200, payload);
        } catch (error) {
            const status = error instanceof PreviewRouteError ? error.status : 500;
            const message = error instanceof PreviewRouteError
                ? error.message
                : `${name} 失败：${error?.message ?? error}`;
            const hint = error instanceof PreviewRouteError && typeof error.hint === 'string'
                ? error.hint
                : (hints === null ? undefined : hints[status]);
            const payload = { ok: false, error: message };
            if (typeof hint === 'string' && hint !== '') payload.hint = hint;
            if (status === 405 && allowHeader !== null) {
                sendJsonWithHeaders(res, status, payload, { allow: allowHeader });
            } else {
                sendJson(res, status, payload);
            }
        }
    };
}

/**
 * 注册预览面板路由。
 *
 * @param {object} ctx 插件上下文（需要 ctx.webServer；信任来源同记忆面板：
 *   `ctx.workspaceRegistry` / `ctx.sessionPersistence`）
 * @param {object} options
 * @param {() => object} options.getMemory 当前已解析的 memory 配置
 * @param {() => object} [options.getConfig] 当前已解析的插件配置（取 cacheDir）
 * @param {Function} [options.render] 换掉 previewImages（测试注入假引擎）
 * @param {Function} [options.readPpt] 换掉 formats/ppt 的 read（测试注入假解析）
 * @param {Function} [options.loadEngine] 换掉 loadPreviewKit（测试模拟「引擎没装」）
 * @returns {(() => void)|undefined} 释放函数；没有 webServer 时返回 undefined
 */
export function registerPreviewView(ctx, options = {}) {
    const webServer = typeof ctx?.get === 'function' ? ctx.get('webServer') : undefined;
    if (webServer === undefined || webServer === null || typeof webServer.register !== 'function') {
        return undefined;
    }
    const getMemory = typeof options.getMemory === 'function' ? options.getMemory : () => ({});
    const getConfig = typeof options.getConfig === 'function' ? options.getConfig : () => ({});
    const deps = {
        readPpt: typeof options.readPpt === 'function' ? options.readPpt : readPptx,
        render: typeof options.render === 'function' ? options.render : previewImages,
        loadEngine: typeof options.loadEngine === 'function' ? options.loadEngine : loadPreviewKit,
        get cacheDir() {
            const config = getConfig() ?? {};
            return typeof config.cacheDir === 'string' && config.cacheDir !== '' ? config.cacheDir : '.office/runtime/cache';
        },
    };
    const trust = (url) => trustedRootFor(ctx, url.searchParams.get('cwd'), getMemory() ?? {});

    const disposers = [
        webServer.register({
            kind: 'exact',
            path: PREVIEW_VIEW_PATHS.files,
            handler: routeHandler('列出产物', async (url, root) => {
                const files = scanOfficeFiles(root.root);
                return {
                    ok: true,
                    root: root.root,
                    label: root.label,
                    files: files.slice(0, PREVIEW_FILE_LIMIT).map(({ mtimeMs, ...item }) => item),
                    total: files.length,
                    truncated: files.length > PREVIEW_FILE_LIMIT,
                    limit: PREVIEW_FILE_LIMIT,
                    depth: SCAN_DEPTH,
                };
            }, trust),
        }),
        webServer.register({
            kind: 'exact',
            path: PREVIEW_VIEW_PATHS.outline,
            handler: routeHandler('读取大纲', async (url, root) => {
                const relative = url.searchParams.get('path');
                const absolute = readPathFor(root.root, relative);
                requireExisting(absolute, String(relative).trim());
                return outlinePayload({
                    root: root.root,
                    label: root.label,
                    relative: String(relative).trim().replace(/\\/g, '/'),
                    absolute,
                    deps,
                });
            }, trust),
        }),
        webServer.register({
            kind: 'exact',
            path: PREVIEW_VIEW_PATHS.page,
            handler: routeHandler('渲染页面', async (url, root) => {
                const relativeRaw = url.searchParams.get('path');
                const absolute = readPathFor(root.root, relativeRaw);
                requireExisting(absolute, String(relativeRaw).trim());
                const rawPage = url.searchParams.get('page');
                const pageNo = Math.trunc(Number(rawPage === null || rawPage === '' ? 1 : rawPage));
                if (!Number.isFinite(pageNo) || pageNo < 1) {
                    throw new PreviewRouteError(400, `page 要从 1 开始的整数（收到 ${JSON.stringify(rawPage)}）。下一步：给 page=1 试第一页。`);
                }
                const rawDpi = url.searchParams.get('dpi');
                const dpi = Math.trunc(Number(rawDpi === null || rawDpi === '' ? PREVIEW_DEFAULT_DPI : rawDpi));
                if (!Number.isFinite(dpi) || dpi < DPI_MIN || dpi > DPI_MAX) {
                    throw new PreviewRouteError(400, `dpi 要在 ${DPI_MIN}–${DPI_MAX} 之间（收到 ${JSON.stringify(rawDpi)}）。`
                        + `下一步：把 dpi 调回 ${DPI_MIN}–${DPI_MAX}，常用 ${PREVIEW_DEFAULT_DPI}。`);
                }
                return pagePayload({
                    root: root.root,
                    label: root.label,
                    relative: String(relativeRaw).trim().replace(/\\/g, '/'),
                    absolute,
                    pageNo,
                    dpi,
                    deps,
                });
            }, trust),
        }),
        webServer.register({
            kind: 'exact',
            path: TEXT_VIEW_PATH,
            handler: routeHandler('预览原文', async (url, root) => {
                const rawPath = url.searchParams.get('path');
                const relativeRaw = String(rawPath ?? '').trim().replace(/\\/g, '/');
                const absolute = readPathFor(root.root, rawPath);
                requireExisting(absolute, relativeRaw);
                return textPayload({
                    root: root.root,
                    label: root.label,
                    relative: relativeRaw,
                    absolute,
                });
            }, trust, { allowHeader: 'GET, HEAD', hints: TEXT_ERROR_HINTS }),
        }),
    ];
    // 一个根共用四条路由：任何一条注册失败（路径冲突）都要把已注册的撤掉，
    // 否则面板会看到「一半端点 200、一半 404」的怪状态。
    return () => {
        for (const dispose of disposers) {
            if (typeof dispose === 'function') dispose();
        }
    };
}

