/**
 * 配图的来源与获取（office.image）：网络取图、图库检索，以及**署名**。
 *
 * 两条历史遗留在这里合流：
 *   - `11-1` 网络取图与图库检索：办公模式以前只能嵌本地已有的图（office.python 画出来的，
 *     或用户自己放进工作目录的），没有任何一条「去哪儿找图」的路。
 *   - `1-2` 配图无来源：就算嵌了图，产物里也不写这张图是哪来的、什么许可 ——
 *     正式文档里的配图必须能追到来源。
 *
 * 三条口径：
 *   1. **只用免 Key 的公开图库**（维基共享资源、Openverse）。要 Key 的图库不进默认
 *      清单：那会让「取一张图」变成先配 Key 才能做的事。任何一张图也可以直接给 URL
 *      走 `fetch`（那时署名信息只能靠调用方自己给）。
 *   2. **所有 HTTP 都走 web.js 的字节通道**（`httpFetchBytes`）：地址校验（SSRF 防线）、
 *      只跟同源跳转、限长、超时分类都是同一份实现，图片不另开一条没护栏的路。
 *   3. **署名随图一起回**：结果里带 `attribution`（完整，含来源页）与 `credit`（一行，
 *      直接写进图注）。嵌进文档时用 word / ppt 的 `source` 选项把它写进产物。
 *
 * 第七十二轮把「获取」补齐三段：本地图片的元信息（info）、网页里的图片候选
 * （fromPage；纯解析在 pageImageCandidates 里，离线可测）、文档里已经嵌进去的
 * 媒体（extract）。三条路复用现有的字节通道、env 落盘与魔数嗅探，不另开出口。
 *
 * @module dsh-office-mode/image-source
 */
import { createHash } from 'node:crypto';

import { assertInsideRoot, displayPath } from './engine/kit.js';
import { unzip } from './engine/zip.js';
import { sniffImageBytes } from './image-bytes.js';
import { charsetOf, httpFetchBytes } from './web.js';
import { WebAccessError } from './web-errors.js';

/** 图库通道（免 Key 的两条；顺序 = auto 的尝试顺序）。 */
export const IMAGE_CHANNELS = Object.freeze([
    {
        id: 'commons',
        label: '维基共享资源（Wikimedia Commons）',
        family: 'commons',
        needsKey: false,
        endpoint: 'https://commons.wikimedia.org/w/api.php',
        note: '免 Key；作者与许可（含许可链接）随结果一起给，署名信息最全',
    },
    {
        id: 'openverse',
        label: 'Openverse（CC 图库聚合）',
        family: 'openverse',
        needsKey: false,
        endpoint: 'https://api.openverse.org/v1/images/',
        note: '免 Key；聚合多个 CC 图库，结果带 creator 与 license（匿名调用有速率限制）',
    },
]);

export const IMAGE_CHANNEL_IDS = Object.freeze(IMAGE_CHANNELS.map((channel) => channel.id));
export const DEFAULT_IMAGE_CHANNEL = 'commons';

/** 一次最多回几条候选；超过就只回这么多（不静默截断，结果里写出来）。 */
export const IMAGE_MAX_RESULTS = 40;
/** 缩略图/正文图的默认宽度（像素）。 */
export const IMAGE_DEFAULT_WIDTH = 1024;
/** 单张图默认体积上限。 */
export const IMAGE_MAX_BYTES = 8 * 1024 * 1024;

/** 能嵌进 **Word** 的图片格式（word.js 只吃 PNG/JPEG；GIF 会被跳过）。 */
const WORD_INSERTABLE_MIME = new Set(['image/png', 'image/jpeg']);
/** 能嵌进 **PPT** 的图片格式（ppt.js 还吃 GIF）。 */
const PPT_INSERTABLE_MIME = new Set(['image/png', 'image/jpeg', 'image/gif']);

/** 取哪一份地址时用的「是不是位图」判断：Word 与 PPT 的交集（PNG/JPEG）优先。 */
const RASTER_MIME = PPT_INSERTABLE_MIME;

function channelOf(id) {
    const wanted = String(id ?? '').trim().toLowerCase();
    return IMAGE_CHANNELS.find((channel) => channel.id === wanted);
}

/** 通道清单（给 office.image.channels() 与文档用）。 */
export function imageChannels() {
    return IMAGE_CHANNELS.map((channel) => ({
        id: channel.id,
        label: channel.label,
        needsKey: channel.needsKey,
        note: channel.note,
    }));
}

function asText(value) {
    return typeof value === 'string' ? value : (value === undefined || value === null ? '' : String(value));
}

/** 去掉 HTML 标签与实体（图库的 Artist 字段常带 `<a>`）。 */
export function stripHtml(value) {
    return asText(value)
        .replace(/<[^>]*>/g, ' ')
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/\s+/g, ' ')
        .trim();
}

function clampInt(value, min, max, fallback) {
    const n = Math.trunc(Number(value));
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

/** 从 URL 的扩展名猜 mime（图库给的 mime 不可靠时兜底）。 */
function mimeFromUrl(url) {
    const path = asText(url).replace(/[?#].*$/, '').toLowerCase();
    if (path.endsWith('.png')) return 'image/png';
    if (path.endsWith('.jpg') || path.endsWith('.jpeg')) return 'image/jpeg';
    if (path.endsWith('.gif')) return 'image/gif';
    if (path.endsWith('.webp')) return 'image/webp';
    if (path.endsWith('.bmp')) return 'image/bmp';
    if (path.endsWith('.tif') || path.endsWith('.tiff')) return 'image/tiff';
    if (path.endsWith('.emf')) return 'image/emf';
    if (path.endsWith('.wmf')) return 'image/wmf';
    if (path.endsWith('.svg')) return 'image/svg+xml';
    return '';
}

/** 一行的署名（写进图注）：`维基共享资源 · 作者 X · CC BY-SA 4.0`。 */
export function creditOf(candidate) {
    const parts = [];
    const provider = channelOf(candidate?.channel)?.label ?? asText(candidate?.channel);
    if (provider !== '') parts.push(provider);
    const author = stripHtml(candidate?.author);
    if (author !== '') parts.push(author);
    const license = asText(candidate?.license);
    if (license !== '') parts.push(license);
    return parts.join(' · ');
}

/** 完整的署名（含来源页与许可链接），给「按来源核对」用。 */
export function attributionOf(candidate) {
    const provider = channelOf(candidate?.channel)?.label ?? asText(candidate?.channel);
    const lines = [`来源：${provider === '' ? '网络' : provider}`];
    const title = asText(candidate?.title);
    if (title !== '') lines.push(`名称：${title}`);
    const author = stripHtml(candidate?.author);
    if (author !== '') lines.push(`作者：${author}`);
    const license = asText(candidate?.license);
    if (license !== '') lines.push(`许可：${license}${asText(candidate?.licenseUrl) === '' ? '' : `（${candidate.licenseUrl}）`}`);
    const pageUrl = asText(candidate?.pageUrl);
    if (pageUrl !== '') lines.push(`来源页：${pageUrl}`);
    lines.push(`图片地址：${asText(candidate?.url)}`);
    return lines.join('\n');
}

/** 把一条候选补全成统一形状。 */
function normalizeCandidate(raw) {
    const url = asText(raw?.url);
    const thumbUrl = asText(raw?.thumbUrl) || url;
    const urlMime = asText(raw?.mime) || mimeFromUrl(url) || mimeFromUrl(thumbUrl);
    const thumbMime = mimeFromUrl(thumbUrl);
    // 取哪一个地址：优先原图；原图不是位图（SVG / TIFF / WebP）而缩略图是 PNG/JPEG/GIF 时，
    // 用缩略图 —— 图库的缩略图服务会把矢量图栅格化成 PNG，这正是能嵌进文档的那一份。
    // 少了这一步，「insertable 说可以嵌」而 fetch 取回一个 SVG，就会在最后一步失败
    // （真联网自测里踩到过：维基共享资源的 SVG 原图）。
    // 原图字节数超过取图上限（默认 8 MB）时，fetchUrl 直接指向缩略图：
    // 图库的一条候选本来就该是「能取回来的那一份」，等到 fetch 撞上限再报错太晚。
    // 原图字段（url）保持原样，调用方仍能自己决定要不要大图。
    const size = Number(raw?.size) || null;
    const tooLarge = size !== null && size > IMAGE_MAX_BYTES;
    const fetchUrl = (tooLarge || !RASTER_MIME.has(urlMime))
        ? (RASTER_MIME.has(thumbMime) ? thumbUrl : url)
        : url;
    const candidate = {
        title: stripHtml(raw?.title) || url.replace(/^.*\//, ''),
        url,
        fetchUrl,
        thumbUrl,
        size,
        pageUrl: asText(raw?.pageUrl),
        width: Number(raw?.width) || null,
        height: Number(raw?.height) || null,
        mime: urlMime,
        license: asText(raw?.license),
        licenseUrl: asText(raw?.licenseUrl),
        author: stripHtml(raw?.author),
        channel: asText(raw?.channel),
        // 分宿主报：GIF 在 PowerPoint 里能用、在 Word 里会被跳过 —— 一个布尔量说不清，
        // 报错了会让「insertable:true」的图在 docx 里静默消失。
        insertable: WORD_INSERTABLE_MIME.has(mimeFromUrl(fetchUrl)),
        insertableInWord: WORD_INSERTABLE_MIME.has(mimeFromUrl(fetchUrl)),
        insertableInPpt: PPT_INSERTABLE_MIME.has(mimeFromUrl(fetchUrl)),
    };
    candidate.credit = creditOf(candidate);
    candidate.attribution = attributionOf(candidate);
    return candidate;
}

/**
 * 用字节通道取一个 JSON（图库的检索接口）。所有图库请求都走这里，
 * 于是「地址校验 / 跳转 / 限长 / 超时」只有一份实现。
 */
async function fetchJson(url, deps, options = {}) {
    const fetchBytes = deps.fetchBytes ?? httpFetchBytes;
    const response = await fetchBytes(url, {
        accept: 'application/json',
        maxBytes: options.maxBytes ?? 4 * 1024 * 1024,
        fetchTimeoutMs: options.fetchTimeoutMs ?? 30_000,
    }, options.signal);
    const text = new TextDecoder('utf-8').decode(response.bytes);
    try {
        return JSON.parse(text);
    } catch (error) {
        throw new WebAccessError(
            `图库接口返回的不是 JSON（多半是拦截页或接口变了）：${url}`,
            'OFFICE_WEB_PROVIDER',
            { cause: error },
        );
    }
}

/** 维基共享资源：action=query + generator=search + imageinfo（含 extmetadata）。 */
function commonsUrl(query, { limit, width }) {
    const params = new URLSearchParams({
        action: 'query',
        format: 'json',
        formatversion: '2',
        generator: 'search',
        gsrsearch: query,
        gsrnamespace: '6',
        gsrlimit: String(limit),
        prop: 'imageinfo',
        iiprop: 'url|size|mime|extmetadata',
        iiurlwidth: String(width),
    });
    return `https://commons.wikimedia.org/w/api.php?${params.toString()}`;
}

/**
 * 图库接口的「错误信封」→ 一句可读的原因；不是错误信封时返回空串。
 *
 * 为什么需要：接口在 HTTP 200 里回 `{"error":{"code":"badvalue",…}}` 或
 * `{"detail":"…"}` 是常态（宽度超过原图、限流、参数写错）。把它们当成「没有结果」，
 * 会让人去改检索词而不是改参数 —— 修法指错地方比不报还糟。
 */
function apiErrorOf(payload) {
    const error = payload?.error;
    if (error !== undefined && error !== null) {
        const code = asText(error.code);
        const info = asText(error.info ?? error.message ?? error.detail);
        return `接口报错${code === '' ? '' : `（${code}）`}${info === '' ? '' : `：${info}`}`;
    }
    const detail = asText(payload?.detail);
    return detail === '' ? '' : `接口报错：${detail}`;
}

function mapCommons(payload, channel) {
    const pages = payload?.query?.pages;
    if (!Array.isArray(pages)) return [];
    return pages.map((page) => {
        const info = Array.isArray(page?.imageinfo) ? page.imageinfo[0] : undefined;
        const meta = info?.extmetadata ?? {};
        return normalizeCandidate({
            title: asText(page?.title).replace(/^File:/i, ''),
            url: asText(info?.url),
            thumbUrl: asText(info?.thumburl),
            pageUrl: asText(info?.descriptionurl),
            width: Number(info?.width) || null,
            height: Number(info?.height) || null,
            mime: asText(info?.mime),
            size: Number(info?.size) || null,
            license: asText(meta?.LicenseShortName?.value),
            licenseUrl: asText(meta?.LicenseUrl?.value),
            author: asText(meta?.Artist?.value) || asText(meta?.Credit?.value),
            channel: channel.id,
        });
    }).filter((candidate) => candidate.url !== '');
}

/** Openverse：/v1/images/?q=…（匿名可用，带速率限制）。 */
function openverseUrl(query, { limit }) {
    const params = new URLSearchParams({ q: query, page_size: String(limit) });
    return `https://api.openverse.org/v1/images/?${params.toString()}`;
}

function mapOpenverse(payload, channel) {
    const results = payload?.results;
    if (!Array.isArray(results)) return [];
    return results.map((item) => normalizeCandidate({
        title: asText(item?.title),
        url: asText(item?.url),
        thumbUrl: asText(item?.thumbnail),
        pageUrl: asText(item?.foreign_landing_url),
        width: Number(item?.width) || null,
        height: Number(item?.height) || null,
        mime: mimeFromUrl(asText(item?.url)),
        license: [asText(item?.license), asText(item?.license_version)].filter((part) => part !== '').join(' ').toUpperCase(),
        licenseUrl: asText(item?.license_url),
        author: asText(item?.creator),
        channel: channel.id,
    })).filter((candidate) => candidate.url !== '');
}

const CHANNEL_MAPPERS = {
    commons: { url: commonsUrl, map: mapCommons },
    openverse: { url: openverseUrl, map: mapOpenverse },
};

/**
 * 检索图库。
 *
 * @param {string} query 检索词（中文也可以：维基共享资源按多语言索引检索）
 * @param {{channel?: string, limit?: number, width?: number, signal?: AbortSignal}} [options]
 * @param {{fetchBytes?: Function}} [deps] 测试注入用
 * @returns {Promise<{query: string, channel: string, count: number, results: Array<object>}>}
 */
export async function imageSearch(query, options = {}, deps = {}) {
    const text = asText(query).trim();
    if (text === '') throw new WebAccessError('office.image.search：query 不能为空。', 'OFFICE_WEB_INVALID_QUERY');
    const limit = clampInt(options.limit, 1, IMAGE_MAX_RESULTS, 8);
    const width = clampInt(options.width, 120, 2400, IMAGE_DEFAULT_WIDTH);
    const wanted = asText(options.channel).trim().toLowerCase();
    if (wanted !== '' && wanted !== 'auto' && channelOf(wanted) === undefined) {
        throw new WebAccessError(
            `office.image.search：没有这条图库通道「${wanted}」。可用：${IMAGE_CHANNEL_IDS.join(' / ')}。`,
            'OFFICE_WEB_PROVIDER',
        );
    }
    const order = wanted === '' || wanted === 'auto' ? IMAGE_CHANNEL_IDS : [wanted];
    const failures = [];
    for (const id of order) {
        const channel = channelOf(id);
        const mapper = CHANNEL_MAPPERS[id];
        if (mapper === undefined) continue;
        try {
            const payload = await fetchJson(mapper.url(text, { limit, width }), deps, options);
            const apiError = apiErrorOf(payload);
            const results = mapper.map(payload, channel).slice(0, limit);
            if (results.length === 0) {
                // 接口自己报了错（参数、限流、接口变了）时照实转述：说成「没有结果」
                // 会让人去改检索词，而真正要改的是参数或通道。
                failures.push(`${id}：${apiError === '' ? '没有结果' : apiError}`);
                continue;
            }
            return {
                query: text,
                channel: id,
                count: results.length,
                truncated: results.length >= limit,
                results: results.map((candidate, index) => ({ index: index + 1, ...candidate })),
                note: '挑一条之后用 office.image.fetch(候选) 取回来（会把来源与许可一起带回）。'
                    + '嵌进文档时用 source 选项把 credit 写进图注。',
            };
        } catch (error) {
            failures.push(`${id}：${error instanceof Error ? error.message : String(error)}`);
        }
    }
    const first = failures[0] ?? '没有可用通道';
    throw new WebAccessError(
        `office.image.search 没有取到图（${failures.join('；') || first}）。`
        + '换一个检索词，或换一条通道（office.image.channels() 看有哪些）；'
        + '报「取资源失败 / fetch failed」通常是本机出不去网（设置页「检索编排 → 出口代理」）。',
        'OFFICE_WEB_NO_RESULTS',
    );
}

// 嗅探实现搬到了 image-bytes.js（切断 gallery→image-source→web→config 的环），
// 这里原样再导出，office.image / excel 等既有调用方不用改。
export { sniffImageBytes };

function slugOf(text, fallback = 'image') {
    const slug = asText(text).replace(/^.*[\\/]/, '').replace(/\.[a-z0-9]+$/i, '')
        .replace(/[\\/:*?"<>|\s]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 50);
    return slug === '' ? fallback : slug;
}

// ── 本地图片的元信息 ───────────────────────────────────────────────────────

/**
 * 读一张**本地**图片的元信息（office.image.info）。
 *
 * 为什么需要：干活时最常问的两句是「这张图多大、能不能直接嵌进 Word / PPT」。
 * 以前只能靠 office.python 或手工读字节；info 一次说清尺寸、体积、内容哈希与
 * 两家宿主的可嵌入性，而且不落盘、不联网。
 *
 * 失败不抛：路径写错、给的是文本文件都是常见调用，返回 {ok:false, reason}
 * 比让脚本在 try/catch 里猜更省事。
 *
 * @param {string} filePath 相对 env.root 或绝对路径
 * @param {object} env 引擎环境
 * @returns {{ok: true, path: string, bytes: number, mime: string, ext: string, width: number|null,
 *   height: number|null, aspectRatio: number|null, sha256: string, insertableInWord: boolean,
 *   insertableInPpt: boolean} | {ok: false, path: string, reason: string}}
 */
export function imageInfo(filePath, env) {
    const shown = asText(filePath).trim();
    if (shown === '') return { ok: false, path: shown, reason: '要给出一个图片路径（相对工作目录或绝对路径）。' };
    const display = displayPath(env.root, env.resolve(shown));
    let bytes;
    try {
        bytes = Buffer.from(env.readFile(shown));
    } catch (error) {
        const reason = error?.code === 'ENOENT'
            ? '文件不存在（相对路径按工作目录解析）。'
            : '读不到这个文件：' + (error?.message ?? String(error));
        return { ok: false, path: display, reason };
    }
    if (bytes.length === 0) return { ok: false, path: display, reason: '文件是空的（0 字节）。' };
    const sniffed = sniffImageBytes(bytes);
    if (sniffed === undefined) {
        return {
            ok: false,
            path: display,
            reason: '不是认得的位图（只按魔数认 PNG / JPEG / GIF / WebP；SVG、TIFF、BMP 这些不做元信息）。',
        };
    }
    const width = Number.isFinite(sniffed.width) ? sniffed.width : null;
    const height = Number.isFinite(sniffed.height) ? sniffed.height : null;
    const inWord = WORD_INSERTABLE_MIME.has(sniffed.mime);
    const inPpt = PPT_INSERTABLE_MIME.has(sniffed.mime);
    return {
        ok: true,
        path: display,
        bytes: bytes.length,
        mime: sniffed.mime,
        ext: sniffed.ext,
        width,
        height,
        // 长宽比保留 4 位：1.3333 比 1.3333333 好读，也比「四舍五入到 1.33」够用。
        aspectRatio: width !== null && height !== null && width > 0 && height > 0
            ? Math.round((width / height) * 10000) / 10000
            : null,
        sha256: createHash('sha256').update(bytes).digest('hex'),
        insertableInWord: inWord,
        insertableInPpt: inPpt,
    };
}

// ── 网页里的图片候选（office.image.fromPage） ──────────────────────────────

/** 网页 HTML 的体积上限：截断的 HTML 解析出来的候选不可信，所以超了直接报错。 */
export const PAGE_MAX_BYTES = 4 * 1024 * 1024;
/** 从一页里最多回几条候选（与图库检索同一个上限）。 */
export const PAGE_MAX_RESULTS = IMAGE_MAX_RESULTS;

/**
 * 抓网页用的浏览器 UA。不带它，不少站点直接回 403 或只给一个「请开 JS」的壳页 ——
 * 那不是「这一页没有图」，是没让我们看见。
 */
const BROWSER_UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

/** HTML 实体 → 字符（属性值、标题与 alt 都要解，否则 &amp; 会原样写进结果）。 */
function decodeHtmlEntities(value) {
    const decodeCode = (whole, raw, radix) => {
        const code = Number.parseInt(raw, radix);
        return Number.isFinite(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    };
    return asText(value)
        .replace(/&#x([0-9a-f]+);/gi, (whole, hex) => decodeCode(whole, hex, 16))
        .replace(/&#(\d+);/g, (whole, dec) => decodeCode(whole, dec, 10))
        .replace(/&nbsp;/gi, ' ')
        .replace(/&quot;/gi, '"')
        .replace(/&apos;|&#39;/gi, "'")
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&amp;/gi, '&');
}

// 标签与属性的正则：引号里的 > 不结束标签（alt="a > b" 这种写法很常见）。
const HTML_TAG_RE = /<([a-zA-Z][a-zA-Z0-9:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*)\/?>/g;
const HTML_ATTR_RE = /([^\s=/>"']+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>]+)))?/g;

/** 一个标签的属性表（小写名、解过实体、同名取第一个）。 */
function htmlAttrs(text) {
    const attrs = {};
    const body = asText(text);
    if (body === '') return attrs;
    HTML_ATTR_RE.lastIndex = 0;
    let match;
    while ((match = HTML_ATTR_RE.exec(body)) !== null) {
        const name = match[1].toLowerCase();
        const value = match[2] ?? match[3] ?? match[4] ?? '';
        if (!Object.prototype.hasOwnProperty.call(attrs, name)) attrs[name] = decodeHtmlEntities(value);
    }
    return attrs;
}

/** 去标签、解实体、压空白（标题与 alt 用）。 */
function pagePlainText(value) {
    return decodeHtmlEntities(asText(value)).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * srcset → 最宽的那一条：优先 w 描述符，其次 x（2x 屏），都没有就取第一条。
 * 返回的 width 只在 w 描述符时可知。
 */
function bestSrcset(value) {
    const text = asText(value).trim();
    if (text === '') return undefined;
    const entries = [];
    for (const raw of text.split(',')) {
        const piece = raw.trim();
        if (piece === '') continue;
        const match = /^(\S+)\s+(\d+(?:\.\d+)?)([wx])$/i.exec(piece);
        if (match !== null) entries.push({ url: match[1], size: Number(match[2]), unit: match[3].toLowerCase() });
        else entries.push({ url: piece.split(/\s+/)[0], size: 0, unit: '' });
    }
    if (entries.length === 0) return undefined;
    const widest = entries.filter((entry) => entry.unit === 'w');
    const densest = entries.filter((entry) => entry.unit === 'x');
    const pool = widest.length > 0 ? widest : densest;
    const chosen = pool.length > 0 ? pool.reduce((best, entry) => (entry.size > best.size ? entry : best)) : entries[0];
    return { url: chosen.url, width: chosen.unit === 'w' ? Math.trunc(chosen.size) : null };
}

/** 把候选地址解析成绝对 http(s) URL；不是能取的那类就返回原因。 */
function resolvePageUrl(raw, baseUrl) {
    const text = asText(raw).trim();
    if (text === '') return { reason: '空地址' };
    const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(text);
    if (scheme !== null) {
        const name = scheme[1].toLowerCase();
        if (name !== 'http' && name !== 'https') return { reason: '不支持的协议（' + name + ':）' };
    } else if (text.startsWith('#')) {
        return { reason: '页内锚点' };
    }
    let parsed;
    try {
        parsed = new URL(text, baseUrl);
    } catch {
        return { reason: '地址解析失败' };
    }
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return { reason: '不支持的协议（' + parsed.protocol + '）' };
    return { url: parsed.toString() };
}

/** 文档头部的 meta charset 声明（只看前 4 KB）；没有就返回空串。 */
function charsetFromHtmlMeta(bytes) {
    const head = Buffer.from(bytes.subarray(0, 4096)).toString('latin1');
    return /<meta[^>]+charset\s*=\s*["']?([a-z0-9._:-]+)/i.exec(head)?.[1]?.toLowerCase() ?? '';
}

/**
 * 按 Content-Type 的 charset 解码 HTML。
 *
 * 注意：web.js 的字节通道会把 Content-Type 的参数（含 charset）去掉，所以生产路径上
 * 通常取不到头里的声明；这时退到文档自己的 <meta charset>（GBK 页面几乎都会写），
 * 都没有就按 utf-8。TextDecoder 认不出的编码名会抛 RangeError —— 换下一条而不是
 * 让整次抓取失败。
 */
function decodeHtmlBytes(bytes, contentType) {
    const declared = charsetOf(contentType) ?? '';
    const declaredInPage = charsetFromHtmlMeta(bytes);
    const labels = [declared, declaredInPage, 'utf-8'].filter((label, index, all) => label !== '' && all.indexOf(label) === index);
    for (const label of labels) {
        try {
            return new TextDecoder(label).decode(bytes);
        } catch {
            // 不认识的编码名：换下一个
        }
    }
    return new TextDecoder('utf-8').decode(bytes);
}

/** 一行署名（写进图注）：网页图片没有作者/许可，只能指到页面。 */
function pageCreditOf(pageUrl) {
    let host = '';
    try {
        host = new URL(pageUrl).host;
    } catch {
        host = '';
    }
    return host === '' ? '网页配图' : '网页配图 · ' + host;
}

/** 完整署名（含来源页与直链）＋一句「许可未知」的提醒。 */
function pageAttributionOf(pageUrl, pageTitle, imageUrl) {
    const lines = ['来源：网页'];
    if (pageTitle !== '') lines.push('页面：' + pageTitle);
    lines.push('来源页：' + pageUrl);
    lines.push('图片地址：' + imageUrl);
    lines.push('许可：未知（网页图片不保证可转载或署名，正式发布前请核对来源与授权）');
    return lines.join('\n');
}

/** 一条网页候选：与 search 同构，但 url 指**所在页面**、fetchUrl 才是图片直链。 */
function pageCandidateOf(fetchUrl, meta, pageUrl, pageTitle) {
    const mime = mimeFromUrl(fetchUrl);
    const inWord = WORD_INSERTABLE_MIME.has(mime);
    const inPpt = PPT_INSERTABLE_MIME.has(mime);
    const alt = asText(meta.alt).trim();
    return {
        title: alt === '' ? slugOf(fetchUrl, 'image') : alt,
        url: pageUrl,
        fetchUrl,
        thumbUrl: fetchUrl,
        pageUrl,
        mime,
        width: Number.isFinite(meta.width) ? meta.width : null,
        height: Number.isFinite(meta.height) ? meta.height : null,
        alt,
        role: meta.role,
        credit: pageCreditOf(pageUrl),
        attribution: pageAttributionOf(pageUrl, pageTitle, fetchUrl),
        insertable: inWord,
        insertableInWord: inWord,
        insertableInPpt: inPpt,
    };
}

/**
 * 从一页 HTML 里解析图片候选（纯函数：不联网、不落盘，fromPage 只负责取字节与解码）。
 *
 * 解析口径：
 *   - <base href> 决定相对地址的基准（浏览器规则：第一个带 href 的 <base>）；
 *   - og:image / twitter:image 排最前（社交卡片选的通常就是正文主图）；
 *   - <img> 依次看 srcset（取最大 w / 2x）、data-src / data-original / data-lazy-src
 *     （懒加载的真实地址）与 src，取第一个能用的；
 *   - data: / blob: / javascript: 与非 http(s) 一律跳过并记进 skipped；
 *   - 按解析后的 URL 去重；已知宽度且小于 minWidth 的也跳过。
 *
 * @param {string} html 页面 HTML
 * @param {string} baseUrl 页面最终地址（重定向之后的那一个）
 * @param {{limit?: number, minWidth?: number}} [options]
 * @returns {{title: string, count: number, results: object[], skipped: Array<{url?: string, reason: string}>}}
 */
export function pageImageCandidates(html, baseUrl, options = {}) {
    const rawBase = asText(baseUrl).trim();
    let page;
    try {
        page = new URL(rawBase);
    } catch {
        throw new Error('office.image.fromPage：' + (rawBase === '' ? '（空地址）' : rawBase) + ' 不是合法的网页地址（要 http / https）。');
    }
    if (page.protocol !== 'http:' && page.protocol !== 'https:') {
        throw new Error('office.image.fromPage：只支持 http / https 页面，收到 ' + page.protocol);
    }
    const pageUrl = page.toString();
    const limit = clampInt(options.limit, 1, PAGE_MAX_RESULTS, 20);
    const minWidth = Math.max(0, Math.trunc(Number(options.minWidth)) || 0);
    // 注释 / <script> / <style> 里也会出现「<img src=…>」这类字符串：先剔掉，
    // 否则 JS 模板与 CSS 里的图片会被当成页面上的真图。
    const cleaned = asText(html)
        .replace(/<!--[\s\S]*?-->/g, ' ')
        .replace(/<script\b[\s\S]*?<\/script>/gi, ' ')
        .replace(/<style\b[\s\S]*?<\/style>/gi, ' ');

    let baseHref = '';
    let ogTitle = '';
    const ogImages = [];
    const imgTags = [];
    HTML_TAG_RE.lastIndex = 0;
    let match;
    while ((match = HTML_TAG_RE.exec(cleaned)) !== null) {
        const name = match[1].toLowerCase();
        if (name === 'base') {
            const attrs = htmlAttrs(match[2]);
            if (baseHref === '' && asText(attrs.href).trim() !== '') baseHref = asText(attrs.href).trim();
            continue;
        }
        if (name === 'meta') {
            const attrs = htmlAttrs(match[2]);
            const key = asText(attrs.property ?? attrs.name ?? attrs.itemprop).trim().toLowerCase();
            if (key === 'og:title') {
                if (ogTitle === '') ogTitle = pagePlainText(attrs.content);
                continue;
            }
            if (key === 'og:image' || key === 'og:image:url' || key === 'og:image:secure_url'
                || key === 'twitter:image' || key === 'twitter:image:src') {
                const content = asText(attrs.content).trim();
                if (content !== '' && !ogImages.some((entry) => entry.url === content)) {
                    ogImages.push({ url: content, width: null, height: null, alt: '' });
                }
                continue;
            }
            // og:image:width / height / alt 按规范只描述**上一条** og:image。
            if (key === 'og:image:width' || key === 'og:image:height' || key === 'og:image:alt') {
                const last = ogImages[ogImages.length - 1];
                if (last === undefined) continue;
                if (key === 'og:image:width') last.width = Math.trunc(Number(attrs.content)) || null;
                else if (key === 'og:image:height') last.height = Math.trunc(Number(attrs.content)) || null;
                else last.alt = pagePlainText(attrs.content);
            }
            continue;
        }
        if (name === 'img') imgTags.push(htmlAttrs(match[2]));
    }

    // 相对地址的基准：<base href> 能解析就用它，写坏了退回页面地址（浏览器也是这么做的）。
    let base = pageUrl;
    if (baseHref !== '') {
        try {
            base = new URL(baseHref, page).toString();
        } catch {
            base = pageUrl;
        }
    }

    const titleTag = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(cleaned);
    const pageTitle = pagePlainText(titleTag === null ? '' : titleTag[1]) || ogTitle;
    const results = [];
    const skipped = [];
    const seen = new Set();
    const accept = (resolved, meta) => {
        if (seen.has(resolved)) {
            skipped.push({ url: resolved, reason: '重复的地址（按 URL 去重）' });
            return;
        }
        if (minWidth > 0 && meta.width !== null && meta.width > 0 && meta.width < minWidth) {
            skipped.push({ url: resolved, reason: '宽度 ' + meta.width + 'px 小于 minWidth ' + minWidth + 'px' });
            return;
        }
        seen.add(resolved);
        results.push(pageCandidateOf(resolved, meta, pageUrl, pageTitle));
    };

    for (const entry of ogImages) {
        const resolved = resolvePageUrl(entry.url, base);
        if (resolved.url === undefined) {
            skipped.push({ url: entry.url, reason: resolved.reason });
            continue;
        }
        accept(resolved.url, { width: entry.width, height: entry.height, alt: entry.alt, role: 'og' });
    }

    for (const attrs of imgTags) {
        // 一张 <img> 的地址来源按可信度排：srcset 最大的一张 → 懒加载属性 → src。
        // 懒加载属性排在 src 前：很多站点把 1px 占位图（甚至是 data:）放在 src，
        // 真实地址只在 data-src / data-original / data-lazy-src 上。
        const sources = [];
        const srcset = bestSrcset(attrs.srcset);
        if (srcset !== undefined) sources.push({ url: srcset.url, width: srcset.width });
        for (const attr of ['data-src', 'data-original', 'data-lazy-src', 'src']) {
            const value = asText(attrs[attr]).trim();
            if (value !== '') sources.push({ url: value, width: null });
        }
        if (sources.length === 0) {
            skipped.push({ url: '', reason: '<img> 没有任何地址属性' });
            continue;
        }
        let picked;
        let failure;
        for (const source of sources) {
            const resolved = resolvePageUrl(source.url, base);
            if (resolved.url !== undefined) {
                picked = { url: resolved.url, width: source.width };
                break;
            }
            if (failure === undefined) failure = resolved.reason;
        }
        if (picked === undefined) {
            skipped.push({ url: sources[0].url, reason: failure ?? '没有能用的图片地址' });
            continue;
        }
        const widthAttr = Math.trunc(Number(attrs.width));
        const heightAttr = Math.trunc(Number(attrs.height));
        accept(picked.url, {
            width: picked.width !== null ? picked.width : (Number.isFinite(widthAttr) && widthAttr > 0 ? widthAttr : null),
            height: Number.isFinite(heightAttr) && heightAttr > 0 ? heightAttr : null,
            alt: attrs.alt,
            role: 'img',
        });
    }

    const capped = results.slice(0, limit);
    return { title: pageTitle, count: capped.length, results: capped, skipped };
}

/**
 * 打开一个网页，把里面的图片候选列出来（office.image.fromPage）。
 *
 * @param {string} url 网页地址（http / https）
 * @param {{limit?: number, minWidth?: number, maxBytes?: number, fetchTimeoutMs?: number, signal?: AbortSignal}} [options]
 * @param {{fetchBytes?: Function}} [deps] 测试注入用
 * @returns {Promise<{ok: true, url: string, finalUrl: string, title: string, count: number,
 *   results: object[], skipped: object[], hint: string}>}
 */
export async function imageFromPage(url, options = {}, deps = {}) {
    const wanted = asText(url).trim();
    if (wanted === '') throw new WebAccessError('office.image.fromPage：要给一个网页地址。', 'OFFICE_WEB_INVALID_URL');
    const fetchBytes = deps.fetchBytes ?? httpFetchBytes;
    const response = await fetchBytes(wanted, {
        accept: 'text/html,application/xhtml+xml;q=0.9,*/*;q=0.8',
        userAgent: BROWSER_UA,
        maxBytes: clampInt(options.maxBytes, 4096, 32 * 1024 * 1024, PAGE_MAX_BYTES),
        fetchTimeoutMs: options.fetchTimeoutMs ?? 30_000,
    }, options.signal);
    const bytes = response.bytes ?? new Uint8Array(0);
    const finalUrl = asText(response.url) || wanted;
    // 把图片直链误当网页传进来是常事：直接说该用哪个入口，别让人对着「0 张候选」猜。
    const sniffedDirect = sniffImageBytes(bytes);
    if (sniffedDirect !== undefined) {
        throw new WebAccessError(
            'office.image.fromPage：' + finalUrl + ' 是图片直链（' + sniffedDirect.mime + '），不是网页；直接用 office.image.fetch(url) 取回。',
            'OFFICE_WEB_UNSUPPORTED_TYPE',
        );
    }
    const html = decodeHtmlBytes(bytes, response.contentType);
    const page = pageImageCandidates(html, finalUrl, { limit: options.limit, minWidth: options.minWidth });
    const hint = page.count === 0
        ? '这一页没解析到图片（可能整页由脚本渲染，或图片只写在 CSS 背景里）。换一个页面，或把图片直链交给 office.image.fetch。'
        : '挑一条后用 office.image.fetch(候选) 取回（魔数会再确认一次真实类型）；网页图片的许可未知，正式发布前请核对来源。';
    return {
        ok: true,
        url: wanted,
        finalUrl,
        title: page.title,
        count: page.count,
        results: page.results,
        skipped: page.skipped,
        hint,
    };
}

// ── 文档里已经嵌进去的媒体（office.image.extract） ────────────────────────

/** 内嵌媒体里认得的图片扩展名；其余（音频、视频、字体、二进制）进 skipped。 */
const MEDIA_IMAGE_EXT = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'emf', 'wmf', 'svg']);
const MEDIA_MIME_BY_EXT = Object.freeze({
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp',
    bmp: 'image/bmp', tiff: 'image/tiff', tif: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf', svg: 'image/svg+xml',
});
/** 三种 OOXML 的内嵌媒体目录。 */
const MEDIA_PREFIX_BY_FORMAT = Object.freeze({ docx: 'word/media/', pptx: 'ppt/media/', xlsx: 'xl/media/' });

/** 文件名安全化：去掉目录与 Windows 非法字符，中文照留。 */
function safeStem(value, fallback = 'document') {
    const text = asText(value).replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '')
        .replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '-').replace(/^-+|-+$/g, '');
    return text === '' ? fallback : text.slice(0, 60);
}

/**
 * 把 OOXML 文档里**已经嵌进去**的图片抽出来（office.image.extract）。
 *
 * 什么时候用：手上这份 docx / pptx / xlsx 里有图，想复用、想看一眼、想重新嵌进
 * 另一份产物。抽取只认包内的 media 目录（word/media、ppt/media、xl/media），
 * 不碰 XML —— 那些是文档结构，不是素材。
 *
 * 每条都经 env.writeFile 落盘，因此会进这次调用的产物清单与台账。文件名带内容哈希
 * 前 8 位：同 basename 的两份稿子（不同目录）即使都叫 image1.png 也不会互相覆盖。
 *
 * @param {string} filePath docx / pptx / xlsx（相对 env.root 或绝对路径）
 * @param {object} env 引擎环境
 * @param {{out?: string, overwrite?: boolean}} [options]
 * @returns {{ok: true, format: string, path: string, dir: string, count: number, files: object[],
 *   skipped: object[], hint: string}}
 */
export function imageExtract(filePath, env, options = {}) {
    const shown = asText(filePath).trim();
    if (shown === '') throw new Error('office.image.extract：要给一个 docx / pptx / xlsx 路径。');
    const sourcePath = displayPath(env.root, env.resolve(shown));
    const format = /\.([^.\\/]+)$/.exec(sourcePath)?.[1]?.toLowerCase() ?? '';
    if (MEDIA_PREFIX_BY_FORMAT[format] === undefined) {
        throw new Error('office.image.extract：' + sourcePath + ' 不是 docx / pptx / xlsx（只抽这三类 OOXML 包的内嵌媒体）。');
    }
    let bytes;
    try {
        bytes = Buffer.from(env.readFile(shown));
    } catch (error) {
        throw new Error('office.image.extract：读不到 ' + sourcePath + '（'
            + (error?.code === 'ENOENT' ? '文件不存在' : (error?.message ?? String(error)))
            + '）—— 相对路径按工作目录解析。');
    }
    let filesMap;
    try {
        filesMap = unzip(bytes);
    } catch (error) {
        throw new Error('office.image.extract：' + sourcePath + ' 不是有效的 OOXML 包（旧版 .doc / .ppt / .xls 或文件损坏）：'
            + (error?.message ?? String(error)));
    }
    const stem = safeStem(sourcePath);
    const wantedOut = asText(options.out).trim();
    // 与 office.image.fetch / archive.extract 同口径：相对路径不得逃出工作目录，
    // 绝对路径视为调用方显式选择。
    const dirAbsolute = wantedOut !== ''
        ? assertInsideRoot(wantedOut, env, 'office.image.extract')
        : assertInsideRoot('assets/' + stem + '-media', env, 'office.image.extract');
    const dir = displayPath(env.root, dirAbsolute);
    const prefix = MEDIA_PREFIX_BY_FORMAT[format];
    const overwrite = options.overwrite === true;
    const files = [];
    const skipped = [];
    for (const [part, data] of filesMap) {
        if (!part.toLowerCase().startsWith(prefix)) continue;
        const base = part.slice(prefix.length);
        const dot = base.lastIndexOf('.');
        const ext = dot === -1 ? '' : base.slice(dot + 1).toLowerCase();
        if (!MEDIA_IMAGE_EXT.has(ext)) {
            skipped.push({
                part,
                bytes: data.length,
                reason: ext === '' ? '没有扩展名（认不出是不是图片）' : '不是图片扩展名（.' + ext + '）',
            });
            continue;
        }
        const partName = (dot === -1 ? base : base.slice(0, dot)).replace(/[\\/]+/g, '-');
        const shortHash = createHash('sha1').update(data).digest('hex').slice(0, 8);
        const fileName = stem + '-' + partName + '-' + shortHash + '.' + ext;
        const target = dir + '/' + fileName;
        if (!overwrite && env.exists(target)) {
            skipped.push({ part, path: target, bytes: data.length, reason: '目标已存在（要覆盖传 overwrite:true）' });
            continue;
        }
        const written = env.writeFile(target, data);
        const sniffed = sniffImageBytes(data);
        files.push({
            path: written.path,
            part,
            bytes: written.bytes,
            mime: sniffed?.mime ?? MEDIA_MIME_BY_EXT[ext] ?? 'application/octet-stream',
            width: sniffed?.width ?? null,
            height: sniffed?.height ?? null,
            sha256: createHash('sha256').update(data).digest('hex'),
        });
    }
    const hint = files.length === 0
        ? (skipped.length === 0
            ? '这份 ' + format + ' 里没有内嵌媒体（' + prefix + ' 下没有条目）。图表、SmartArt 不是图片部件。'
            : '没有抽出图片：' + skipped.length + ' 个条目被跳过（见 skipped；目标已存在时默认不覆盖，要重写传 overwrite:true）。')
        : '共抽出 ' + files.length + ' 张（' + format + ' → ' + dir + '/）。files[].path 直接交给 word 的 builder.image() 或 ppt 的 deck.image()；'
            + 'EMF / WMF / SVG / TIFF 这类 Word / PPT 不认，先用 office.python 转成 PNG。';
    return {
        ok: true,
        format,
        path: sourcePath,
        dir,
        count: files.length,
        files,
        skipped,
        hint,
    };
}

/**
 * 取回一张图：下载 → 校验是不是真图 → 落到工作目录（或指定路径）。
 *
 * @param {string|object} ref 图片 URL，或 office.image.search 里的一条候选
 * @param {{to?: string, maxBytes?: number, overwrite?: boolean, signal?: AbortSignal}} [options]
 * @param {object} env 引擎环境
 * @param {{fetchBytes?: Function}} [deps]
 */
/** 候选对象里「这次该取哪一个地址」：知道字节数时按本次 maxBytes 重新选。 */
function fetchUrlOf(ref, maxBytes) {
    if (ref === null || typeof ref !== 'object') return asText(ref).trim();
    const original = asText(ref.url);
    const thumb = asText(ref.thumbUrl);
    const size = Number(ref.size);
    const originalRaster = RASTER_MIME.has(asText(ref.mime) || mimeFromUrl(original));
    // 知道原图多大、而且这次额度装得下，就取原图（即便候选的默认 fetchUrl 是缩略图）。
    if (originalRaster && Number.isFinite(size) && size > 0 && size <= maxBytes) return original;
    return asText(ref.fetchUrl || ref.url).trim();
}

/** 取原图撞上限时能退到哪：候选自带的位图缩略图。 */
function thumbFallbackOf(ref, triedUrl) {
    if (ref === null || typeof ref !== 'object') return '';
    const thumb = asText(ref.thumbUrl);
    if (thumb === '' || thumb === triedUrl) return '';
    return RASTER_MIME.has(mimeFromUrl(thumb)) ? thumb : '';
}

/** 只认「资源超过 N KB 上限」这一种失败，别把网络错误也重试一遍。 */
function isTooLargeError(error) {
    return /超过\s*[\d.]+\s*(?:KB|MB)\s*上限/i.test(asText(error?.message));
}

export async function imageFetch(ref, options = {}, env, deps = {}) {
    const isObject = ref !== null && typeof ref === 'object';
    const maxBytes = clampInt(options.maxBytes, 1024, 32 * 1024 * 1024, IMAGE_MAX_BYTES);
    // fetchUrl 是候选里已经算好的「能嵌的那一份」（原图不是位图或超过上限时指向缩略图）。
    const url = fetchUrlOf(ref, maxBytes);
    if (url === '') throw new WebAccessError('office.image.fetch：要给图片 URL 或 search 返回的一条候选。', 'OFFICE_WEB_INVALID_URL');
    const fetchBytes = deps.fetchBytes ?? httpFetchBytes;
    const fetchOptions = {
        accept: 'image/*,*/*;q=0.8',
        maxBytes,
        fetchTimeoutMs: options.fetchTimeoutMs ?? 60_000,
    };
    let response;
    try {
        response = await fetchBytes(url, fetchOptions, options.signal);
    } catch (error) {
        // 候选没给字节数（或字节数不准）时，原图可能照样撞上限；退到缩略图再取一次。
        const fallback = thumbFallbackOf(ref, url);
        if (fallback === '' || !isTooLargeError(error)) throw error;
        response = await fetchBytes(fallback, fetchOptions, options.signal);
    }
    const bytes = response.bytes;
    const sniffed = sniffImageBytes(bytes);
    if (sniffed === undefined) {
        // 最常见的一种：图库返回 200 + 一个 HTML 说明页。照实说，别写成「下载失败」。
        const contentType = asText(response.contentType);
        const head = new TextDecoder('utf-8').decode(bytes.subarray(0, 64)).trim().toLowerCase();
        const looksHtml = contentType === 'text/html' || head.startsWith('<!doctype') || head.startsWith('<html');
        const looksVector = /svg|tiff|postscript/.test(contentType) || head.startsWith('<?xml') || head.startsWith('<svg');
        const hint = looksVector
            ? '（这是矢量图，不是位图：用候选里的 fetchUrl / thumbUrl 取 PNG 缩略图，或先用 office.python 转成 PNG）'
            : (looksHtml ? '（像是一个 HTML 页面）' : '');
        throw new WebAccessError(
            `office.image.fetch：这个地址返回的不是图片（${contentType || '未知类型'}${hint}）：${response.url}`,
            'OFFICE_WEB_UNSUPPORTED_TYPE',
        );
    }
    const base = isObject ? ref : {};
    const source = {
        url: response.url,
        pageUrl: asText(base.pageUrl),
        title: asText(base.title) || slugOf(response.url),
        license: asText(base.license),
        licenseUrl: asText(base.licenseUrl),
        author: stripHtml(base.author),
        channel: asText(base.channel),
        width: sniffed.width ?? (Number(base.width) || null),
        height: sniffed.height ?? (Number(base.height) || null),
    };
    const candidate = normalizeCandidate({ ...source, url: response.url, mime: sniffed.mime });
    // 候选自带署名时以它为准：fromPage 的候选没有 license / author，按 channel 重算会把
    // 「网页配图 · 站点」那一行丢掉（网页图片的许可本来就要调用方自己核对）。
    if (asText(base.credit) !== '') candidate.credit = asText(base.credit);
    if (asText(base.attribution) !== '') candidate.attribution = asText(base.attribution);
    const wanted = asText(options.to).trim();
    const target = wanted !== ''
        ? assertInsideRoot(wanted, env, 'office.image.fetch')
        : `assets/${slugOf(source.title)}-${createHash('sha1').update(bytes).digest('hex').slice(0, 6)}.${sniffed.ext}`;
    if (options.overwrite !== true && env.exists(target)) {
        throw new WebAccessError(`office.image.fetch：${target} 已存在；换一个 to，或显式传 overwrite:true。`, 'OFFICE_WEB_ERROR');
    }
    const written = env.writeFile(target, bytes);
    // 依据记进这一次调用的引用列表：office_run 会把它们写进台账的 source。
    if (typeof env.cite === 'function') {
        env.cite({
            url: source.pageUrl || source.url,
            title: source.title,
            license: source.license,
            kind: 'image',
        });
    }
    const inWord = WORD_INSERTABLE_MIME.has(sniffed.mime);
    const inPpt = PPT_INSERTABLE_MIME.has(sniffed.mime);
    return {
        ok: true,
        path: written.path,
        bytes: written.bytes,
        mime: sniffed.mime,
        ext: sniffed.ext,
        width: sniffed.width,
        height: sniffed.height,
        sha256: createHash('sha256').update(bytes).digest('hex'),
        source,
        credit: candidate.credit,
        attribution: candidate.attribution,
        insertable: inWord,
        insertableInWord: inWord,
        insertableInPpt: inPpt,
        hint: inWord
            ? '嵌进文档：office.word 的 builder.image(path, {caption, source}) 或 office.ppt 的 deck.image({path, caption, source}) —— source 传 credit 就会把署名写进产物。'
            : (inPpt
                ? `${sniffed.mime} 能嵌进 PPT（deck.image / deck.images），但 **Word 只吃 PNG/JPEG**（builder.image 会跳过它）：要进 Word 先用 office.python 转成 PNG。`
                : `${sniffed.mime} 不能直接嵌进 Word（Word 只吃 PNG/JPEG）：先用 office.python 转成 PNG 再嵌。`),
    };
}
