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
 * @module dsh-office-mode/image-source
 */
import { createHash } from 'node:crypto';

import { assertInsideRoot } from './engine/kit.js';
import { httpFetchBytes } from './web.js';
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
    const fetchUrl = RASTER_MIME.has(urlMime) ? url : (RASTER_MIME.has(thumbMime) ? thumbUrl : url);
    const candidate = {
        title: stripHtml(raw?.title) || url.replace(/^.*\//, ''),
        url,
        fetchUrl,
        thumbUrl,
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

/** 图片魔数嗅探：认得 PNG / JPEG / GIF / WebP，并读出像素尺寸。 */
export function sniffImageBytes(bytes) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (bytes.length >= 24 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
        return { mime: 'image/png', ext: 'png', width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
        // JPEG：扫段找 SOF0/SOF2。
        let at = 2;
        while (at + 9 < bytes.length) {
            if (bytes[at] !== 0xff) { at += 1; continue; }
            const marker = bytes[at + 1];
            const length = view.getUint16(at + 2);
            if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
                return { mime: 'image/jpeg', ext: 'jpeg', height: view.getUint16(at + 5), width: view.getUint16(at + 7) };
            }
            if (length <= 0) break;
            at += 2 + length;
        }
        return { mime: 'image/jpeg', ext: 'jpeg', width: null, height: null };
    }
    if (bytes.length >= 10 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
        return { mime: 'image/gif', ext: 'gif', width: view.getUint16(6, true), height: view.getUint16(8, true) };
    }
    if (bytes.length >= 30 && bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46
        && bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50) {
        const fourcc = String.fromCharCode(bytes[12], bytes[13], bytes[14], bytes[15]);
        if (fourcc === 'VP8X') {
            const width = 1 + (bytes[24] | (bytes[25] << 8) | (bytes[26] << 16));
            const height = 1 + (bytes[27] | (bytes[28] << 8) | (bytes[29] << 16));
            return { mime: 'image/webp', ext: 'webp', width, height };
        }
        return { mime: 'image/webp', ext: 'webp', width: null, height: null };
    }
    return undefined;
}

function slugOf(text, fallback = 'image') {
    const slug = asText(text).replace(/^.*[\\/]/, '').replace(/\.[a-z0-9]+$/i, '')
        .replace(/[\\/:*?"<>|\s]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 50);
    return slug === '' ? fallback : slug;
}

/**
 * 取回一张图：下载 → 校验是不是真图 → 落到工作目录（或指定路径）。
 *
 * @param {string|object} ref 图片 URL，或 office.image.search 里的一条候选
 * @param {{to?: string, maxBytes?: number, overwrite?: boolean, signal?: AbortSignal}} [options]
 * @param {object} env 引擎环境
 * @param {{fetchBytes?: Function}} [deps]
 */
export async function imageFetch(ref, options = {}, env, deps = {}) {
    const isObject = ref !== null && typeof ref === 'object';
    // fetchUrl 是候选里已经算好的「能嵌的那一份」（原图不是位图时指向 PNG 缩略图）。
    const url = asText(isObject ? (ref.fetchUrl || ref.url) : ref).trim();
    if (url === '') throw new WebAccessError('office.image.fetch：要给图片 URL 或 search 返回的一条候选。', 'OFFICE_WEB_INVALID_URL');
    const maxBytes = clampInt(options.maxBytes, 1024, 32 * 1024 * 1024, IMAGE_MAX_BYTES);
    const fetchBytes = deps.fetchBytes ?? httpFetchBytes;
    const response = await fetchBytes(url, {
        accept: 'image/*,*/*;q=0.8',
        maxBytes,
        fetchTimeoutMs: options.fetchTimeoutMs ?? 60_000,
    }, options.signal);
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
