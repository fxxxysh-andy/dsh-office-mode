/**
 * 知识库图片资产的安全读取（不启动服务器，便于单测）。
 *
 * 之前的 kb-asset 入口按扩展名给 MIME，直接 readFile 字面路径：
 *   - assets/ 里放一个指向根外的符号链接/目录联接，字面路径仍在根内，readFile 会跟出去；
 *   - .svg / .bmp 被当成图片按扩展名回 content-type，SVG 可以带脚本，且响应没有 nosniff；
 *   - 只看后缀不看内容，.png 里放任意字节也会以 image/png 返回。
 *
 * 这里统一收口：先 realpath 再比根、必须是普通文件、限大小、只按魔数认四种位图，
 * 扩展名不参与 MIME 判定。路由层只需加安全响应头。
 *
 * @module dsh-office-mode/image-asset
 */
import { readFile, realpath, stat } from 'node:fs/promises';
import { isAbsolute, relative, resolve } from 'node:path';
import { sniffImageBytes } from './image-bytes.js';

/** 单张资产上限：与图库入库上限一致，避免一个响应把内存打满。 */
export const ASSET_MAX_BYTES = 20 * 1024 * 1024;

/** 允许回给浏览器的图片 MIME（只有可安全当图片渲染的四种位图）。 */
export const ASSET_SAFE_MIME = Object.freeze(['image/png', 'image/jpeg', 'image/gif', 'image/webp']);

const SAFE_MIME = new Set(ASSET_SAFE_MIME);

function asText(value) {
    return typeof value === 'string' ? value : value === null || value === undefined ? '' : String(value);
}

/**
 * 规范化并校验 assets/ 相对路径；不碰文件系统。
 * @returns {{ok:true, relativePath:string}|{ok:false, status:number, code:string, error:string}}
 */
export function normalizeAssetRelative(raw) {
    const shown = asText(raw).trim().replace(/\\/g, '/');
    if (shown === '') {
        return { ok: false, status: 400, code: 'bad-path', error: '要给一个 assets/ 下的图片路径。' };
    }
    if (shown.includes('\0')) {
        return { ok: false, status: 400, code: 'bad-path', error: '路径里不能有 NUL 字符。' };
    }
    if (!shown.startsWith('assets/')) {
        return { ok: false, status: 400, code: 'bad-path', error: '只服务知识库 assets/ 下的图片。' };
    }
    const segments = shown.split('/');
    if (segments.some((segment) => segment === '' || segment === '.' || segment === '..')) {
        return { ok: false, status: 400, code: 'bad-path', error: '路径里不能有空段、. 或 ..。' };
    }
    if (shown.startsWith('/') || isAbsolute(shown)) {
        return { ok: false, status: 400, code: 'bad-path', error: '路径必须是知识库内的相对路径。' };
    }
    return { ok: true, relativePath: shown };
}

function insideOf(root, target) {
    const rel = relative(root, target);
    return rel !== '' && !rel.startsWith('..') && !isAbsolute(rel);
}

/**
 * 读一张知识库资产。
 *
 * @param {{root:string, relativePath:string, maxBytes?:number}} options
 *   root = 知识库根目录（不是 assets 目录）；relativePath 必须形如 assets/x.png。
 * @returns {Promise<{ok:true, bytes:Buffer, mime:string, size:number, relativePath:string}
 *   | {ok:false, status:number, code:string, error:string}>}
 */
export async function readSafeAsset({ root, relativePath, maxBytes = ASSET_MAX_BYTES } = {}) {
    const normalized = normalizeAssetRelative(relativePath);
    if (!normalized.ok) return normalized;

    const rootText = asText(root).trim();
    if (rootText === '') {
        return { ok: false, status: 500, code: 'no-root', error: '知识库目录没有配置。' };
    }

    let realRoot;
    try {
        realRoot = await realpath(resolve(rootText));
    } catch {
        return { ok: false, status: 404, code: 'missing', error: '知识库目录不存在。' };
    }

    const literal = resolve(realRoot, normalized.relativePath);
    let realTarget;
    try {
        realTarget = await realpath(literal);
    } catch {
        return { ok: false, status: 404, code: 'missing', error: '图片不存在。' };
    }

    // 关键：用 realpath 之后再比根。符号链接/目录联接指到根外时这里会拒绝。
    if (!insideOf(realRoot, realTarget)) {
        return { ok: false, status: 403, code: 'outside', error: '图片路径越出了知识库目录。' };
    }

    let info;
    try {
        info = await stat(realTarget);
    } catch {
        return { ok: false, status: 404, code: 'missing', error: '图片不存在。' };
    }
    if (!info.isFile()) {
        return { ok: false, status: 404, code: 'not-file', error: '这个路径不是普通文件。' };
    }
    if (info.size <= 0) {
        return { ok: false, status: 415, code: 'empty', error: '图片是空文件。' };
    }
    if (info.size > maxBytes) {
        return { ok: false, status: 413, code: 'too-large', error: '图片超过 ' + Math.round(maxBytes / (1024 * 1024)) + ' MiB 上限。' };
    }

    let bytes;
    try {
        bytes = await readFile(realTarget);
    } catch {
        return { ok: false, status: 404, code: 'missing', error: '图片不存在或读不到。' };
    }
    // 读到的长度再核一次：stat 与 readFile 之间文件可能被换掉。
    if (bytes.length === 0) {
        return { ok: false, status: 415, code: 'empty', error: '图片是空文件。' };
    }
    if (bytes.length > maxBytes) {
        return { ok: false, status: 413, code: 'too-large', error: '图片超过 ' + Math.round(maxBytes / (1024 * 1024)) + ' MiB 上限。' };
    }

    const sniffed = sniffImageBytes(bytes);
    if (sniffed === undefined || !SAFE_MIME.has(sniffed.mime)) {
        return { ok: false, status: 415, code: 'not-image', error: '只服务 PNG / JPEG / GIF / WebP 位图。' };
    }

    return { ok: true, bytes, mime: sniffed.mime, size: bytes.length, relativePath: normalized.relativePath };
}

/** 安全响应头：不猜类型、禁脚本、禁止被当 HTML 打开。 */
export function assetSecurityHeaders(mime, size) {
    return {
        'content-type': mime,
        'content-length': String(size),
        'x-content-type-options': 'nosniff',
        'content-security-policy': "default-src 'none'; sandbox",
        // 同一路径可能被重新入库；不做长缓存，避免浏览器拿着旧图一小时。
        'cache-control': 'private, no-cache, must-revalidate',
    };
}