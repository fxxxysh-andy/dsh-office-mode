/**
 * 图库的 HTTP 端点（自包含，挂载前可独立测试）。
 *
 * 为什么单独一份而不是写进 view.js：图库与记忆/知识库是两条数据面，
 * 但**信任边界必须共用**（会话工作目录白名单、POST 只收 application/json、
 * body 有上限）。所以这里不自己发明鉴权：调用方把 view.js 的
 * trustedRootFor / sendJson / readJsonBody 注进来即可；测试里用假的。
 *
 * 路由只有一个 exact 路径（DSH 的 webServer.register 目前只支持 exact）：
 *
 *   GET  /office-memory/gallery?view=list|search|item|image|thumb|status
 *   POST /office-memory/gallery  { action: 'add'|'manual'|'remove'|'folder-create'|'folder-rename'|'folder-remove'|'describe'|'probe'|'vectors', ... }
 *
 * 图片只按 manifest 里的 id 取：id 必须是 64 位十六进制，路径永远由服务端拼，
 * 请求里的任何字符串都不会进入 readFile。尺寸/格式在入库时已经按魔数校验过。
 *
 * @module dsh-office-mode/gallery-routes
 */
import { readFile } from 'node:fs/promises';
import { isAbsolute, join, relative, resolve } from 'node:path';

/** 图库唯一端点（面板拼 URL 用）。 */
export const GALLERY_PATH = '/office-memory/gallery';
export const GALLERY_ROUTE_PATHS = Object.freeze([GALLERY_PATH]);

const MEDIA_TYPES = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    jpg: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
};

function asText(value) {
    return typeof value === 'string' ? value : value === null || value === undefined ? '' : String(value);
}

function jsonOf(res, status, payload) {
    const body = JSON.stringify(payload) + '\n';
    res.writeHead(status, {
        'content-type': 'application/json; charset=utf-8',
        'content-length': Buffer.byteLength(body, 'utf8'),
        'cache-control': 'no-store',
    });
    res.end(body);
}

function readBody(req, limit) {
    return new Promise((settle, fail) => {
        let size = 0;
        const chunks = [];
        req.on('data', (chunk) => {
            size += chunk.length;
            if (size > limit) {
                fail(new Error('请求体超过 ' + limit + ' 字节。'));
                try { req.destroy(); } catch { /* 已经断开 */ }
                return;
            }
            chunks.push(chunk);
        });
        req.on('end', () => settle(Buffer.concat(chunks).toString('utf8')));
        req.on('error', (error) => fail(error));
    });
}

function boolOf(value) {
    return value === true || value === '1' || value === 'true';
}

function idOf(raw) {
    const id = asText(raw).trim().toLowerCase();
    return /^[0-9a-f]{64}$/.test(id) ? id : '';
}

/** 锁竞争不是请求格式错误：给 503 让调用方知道可以稍后重试，而不是 400/404。 */
function statusFor(result, fallback) {
    if (result?.code === 'lock-timeout') return 503;
    return result?.ok === true ? 200 : fallback;
}

/**
 * 创建图库 handler。
 *
 * @param {object} options
 * @param {() => object|null} options.galleryFor 返回当前 gallery 实例（未启用/未初始化时 null）
 * @param {object} [options.ctx] 插件上下文（交给 trustedRootFor）
 * @param {() => object} [options.memoryOf] 返回 memory 配置（停用时拒绝服务）
 * @param {Function} [options.trustedRootFor] view.js 的信任边界函数
 * @param {Function} [options.sendJson] view.js 的 JSON 输出（缺省用本地实现）
 * @param {Function} [options.readJsonBody] view.js 的 body 读取（缺省用本地实现）
 * @param {() => object|null} [options.hostOf] 视觉宿主通道（attachments + llm）
 * @returns {(req, res) => Promise<void>}
 */
export function createGalleryHandler(options = {}) {
    const galleryFor = typeof options.galleryFor === 'function' ? options.galleryFor : () => null;
    const memoryOf = typeof options.memoryOf === 'function' ? options.memoryOf : () => ({});
    const trustedRootFor = typeof options.trustedRootFor === 'function' ? options.trustedRootFor : null;
    const send = typeof options.sendJson === 'function' ? options.sendJson : jsonOf;
    const readJson = typeof options.readJsonBody === 'function' ? options.readJsonBody : readBody;
    const hostOf = typeof options.hostOf === 'function' ? options.hostOf : () => null;
    const bodyLimit = Number.isFinite(options.bodyLimit) ? options.bodyLimit : 30 * 1024 * 1024;

    async function trust(url) {
        const memory = memoryOf() ?? {};
        if (memory.enabled === false) {
            return { ok: false, status: 403, error: '记忆已在设置里关掉。' };
        }
        if (trustedRootFor === null) return { ok: true, root: '' };
        const result = await trustedRootFor(options.ctx ?? null, url.searchParams.get('cwd'), memory);
        if (result?.ok === false) {
            return { ok: false, status: result.status ?? 403, error: result.error ?? '工作目录不在信任范围内。' };
        }
        return { ok: true, root: asText(result?.root) };
    }

    async function serveImage(res, method, gallery, id, kind) {
        const item = await gallery.get(id);
        if (item === null) {
            send(res, 404, { ok: false, error: '图库里没有这张图。' });
            return;
        }
        const relativePath = kind === 'thumb' ? item.thumb : item.original;
        if (relativePath === null || typeof relativePath !== 'string') {
            send(res, 404, { ok: false, error: kind === 'thumb' ? '这张图还没有缩略图。' : '原图路径缺失。' });
            return;
        }
        // manifest 是磁盘上的真源，也可能被改；只认这个 id 对应的规范相对路径，
        // 并再确认解析后仍落在图库根内，绝不把 manifest 里的字符串直接当路径用。
        const extension = asText(item.facts?.ext).toLowerCase();
        const expected = kind === 'thumb'
            ? 'thumbs/' + id + '.jpg'
            : /^[a-z0-9]{1,8}$/.test(extension) ? 'originals/' + id + '.' + extension : '';
        if (expected === '' || relativePath !== expected) {
            send(res, 404, { ok: false, error: '这张图的记录路径不合法。' });
            return;
        }
        const target = resolve(gallery.root, relativePath);
        const rel = relative(gallery.root, target);
        if (rel === '' || rel.startsWith('..') || isAbsolute(rel)) {
            send(res, 404, { ok: false, error: '这张图的记录路径不合法。' });
            return;
        }
        let data;
        try {
            data = await readFile(target);
        } catch {
            send(res, 404, { ok: false, error: kind === 'thumb' ? '缩略图文件不存在。' : '原图文件不存在。' });
            return;
        }
        const contentType = kind === 'thumb'
            ? 'image/jpeg'
            : (item.facts?.mediaType || MEDIA_TYPES[item.facts?.ext] || 'application/octet-stream');
        res.writeHead(200, {
            'content-type': contentType,
            'content-length': String(data.length),
            'cache-control': 'private, max-age=31536000, immutable',
            'x-content-type-options': 'nosniff',
            'content-security-policy': "default-src 'none'; sandbox",
        });
        res.end(method === 'HEAD' ? undefined : data);
    }

    async function getRoute(req, res, url, trusted) {
        const gallery = galleryFor(trusted);
        if (gallery === null) {
            send(res, 503, { ok: false, error: '图库还没初始化。' });
            return;
        }
        const view = asText(url.searchParams.get('view')) || 'list';
        if (view === 'status') {
            send(res, 200, await gallery.status());
            return;
        }
        if (view === 'item') {
            const id = idOf(url.searchParams.get('id'));
            if (id === '') {
                send(res, 400, { ok: false, error: '要给一个合法的图片 id（64 位十六进制）。' });
                return;
            }
            const item = await gallery.get(id);
            if (item === null) {
                send(res, 404, { ok: false, error: '图库里没有这张图。' });
                return;
            }
            send(res, 200, { ok: true, item: gallery.summarize(item) });
            return;
        }
        if (view === 'image' || view === 'thumb') {
            const id = idOf(url.searchParams.get('id'));
            if (id === '') {
                send(res, 400, { ok: false, error: '要给一个合法的图片 id（64 位十六进制）。' });
                return;
            }
            await serveImage(res, req.method, gallery, id, view);
            return;
        }
        const query = asText(url.searchParams.get('q')).trim();
        const imageId = idOf(url.searchParams.get('image'));
        if (query !== '' || imageId !== '') {
            const result = await gallery.search({
                query,
                imageId,
                tags: asText(url.searchParams.get('tags')),
                category: asText(url.searchParams.get('category')),
                folder: asText(url.searchParams.get('folder')),
                favorite: boolOf(url.searchParams.get('favorite')),
                allowUpload: boolOf(url.searchParams.get('allowUpload')),
                limit: Number.parseInt(asText(url.searchParams.get('limit')), 10) || 30,
            });
            send(res, result.ok ? 200 : 400, result);
            return;
        }
        const listed = await gallery.list({
            tags: asText(url.searchParams.get('tags')),
            category: asText(url.searchParams.get('category')),
            folder: asText(url.searchParams.get('folder')),
            favorite: boolOf(url.searchParams.get('favorite')),
            limit: Number.parseInt(asText(url.searchParams.get('limit')), 10) || 200,
            offset: Number.parseInt(asText(url.searchParams.get('offset')), 10) || 0,
        });
        send(res, 200, listed);
    }

    async function postRoute(req, res, url, trusted) {
        const gallery = galleryFor(trusted);
        if (gallery === null) {
            send(res, 503, { ok: false, error: '图库还没初始化。' });
            return;
        }
        if (!asText(req.headers?.['content-type']).includes('application/json')) {
            send(res, 415, { ok: false, error: 'POST 要用 application/json。' });
            return;
        }
        let body;
        try {
            body = JSON.parse(await readJson(req, bodyLimit));
        } catch (error) {
            send(res, 400, { ok: false, error: '读不了请求体：' + (error?.message ?? String(error)) });
            return;
        }
        if (body === null || typeof body !== 'object' || Array.isArray(body)) {
            send(res, 400, { ok: false, error: '请求体要是一个 JSON 对象。' });
            return;
        }
        const action = asText(body.action);

        if (action === 'add') {
            const encoded = asText(body.data);
            if (encoded === '') {
                send(res, 400, { ok: false, error: 'add 需要 data（base64 图片）。' });
                return;
            }
            let bytes;
            try {
                bytes = Buffer.from(encoded, 'base64');
            } catch {
                send(res, 400, { ok: false, error: 'data 不是合法的 base64。' });
                return;
            }
            const added = await gallery.addBytes(bytes, {
                name: body.name,
                description: body.description,
                tags: body.tags,
                category: body.category,
                folder: body.folder,
                favorite: boolOf(body.favorite),
                note: body.note,
                origin: { kind: 'upload', label: asText(body.name) },
            });
            send(res, statusFor(added, 400), added);
            return;
        }

        if (action === 'manual') {
            const id = idOf(body.id);
            if (id === '') {
                send(res, 400, { ok: false, error: 'manual 需要合法的图片 id。' });
                return;
            }
            const patch = body.patch !== null && typeof body.patch === 'object' && !Array.isArray(body.patch)
                ? body.patch
                : {
                    name: body.name,
                    description: body.description,
                    tags: body.tags,
                    addTags: body.addTags,
                    removeTags: body.removeTags,
                    category: body.category,
                    folder: body.folder,
                    favorite: body.favorite,
                    note: body.note,
                };
            const updated = await gallery.updateManual(id, patch);
            send(res, statusFor(updated, 404), updated);
            return;
        }

        if (action === 'remove') {
            const id = idOf(body.id);
            if (id === '') {
                send(res, 400, { ok: false, error: 'remove 需要合法的图片 id。' });
                return;
            }
            const removed = await gallery.remove(id);
            send(res, statusFor(removed, 404), removed);
            return;
        }

        if (action === 'folder-create') {
            const created = await gallery.createFolder(body.path ?? body.folder);
            send(res, statusFor(created, 400), created);
            return;
        }

        if (action === 'folder-rename') {
            const renamed = await gallery.renameFolder(body.from, body.to);
            send(res, statusFor(renamed, renamed?.code === 'folder-not-found' ? 404 : 400), renamed);
            return;
        }

        if (action === 'folder-remove') {
            const removed = await gallery.removeFolder(body.path ?? body.folder);
            send(res, statusFor(removed, removed?.code === 'folder-not-found' ? 404 : 400), removed);
            return;
        }

        if (action === 'describe') {
            const id = idOf(body.id);
            if (id === '') {
                send(res, 400, { ok: false, error: 'describe 需要合法的图片 id。' });
                return;
            }
            const described = await gallery.describeImage(id, {
                allowUpload: boolOf(body.allowUpload),
                host: hostOf(),
            });
            send(res, described.ok ? 200 : 400, described);
            return;
        }

        if (action === 'probe') {
            const probed = await gallery.probeImage({ allowUpload: boolOf(body.allowUpload) });
            send(res, probed.ok ? 200 : 400, probed);
            return;
        }

        if (action === 'vectors') {
            const vectors = await gallery.ensureVectors({
                ids: Array.isArray(body.ids) ? body.ids : null,
                text: body.text !== false,
                image: boolOf(body.image),
                force: boolOf(body.force),
                allowUpload: boolOf(body.allowUpload),
            });
            send(res, vectors.ok ? 200 : 400, vectors);
            return;
        }

        send(res, 400, {
            ok: false,
            error: '不认识的 action：' + (action === '' ? '（空）' : action),
            actions: ['add', 'manual', 'remove', 'folder-create', 'folder-rename', 'folder-remove', 'describe', 'probe', 'vectors'],
        });
    }

    return async function handleGalleryRoute(req, res) {
        const method = asText(req.method ?? 'GET').toUpperCase();
        try {
            const url = new URL(asText(req.url ?? '/'), 'http://127.0.0.1');
            const trusted = await trust(url);
            if (!trusted.ok) {
                send(res, trusted.status, { ok: false, error: trusted.error });
                return;
            }
            if (method === 'GET' || method === 'HEAD') {
                await getRoute(req, res, url, trusted);
                return;
            }
            if (method === 'POST') {
                await postRoute(req, res, url, trusted);
                return;
            }
            send(res, 405, { ok: false, error: '这个端点只支持 GET / HEAD / POST。' });
        } catch (error) {
            send(res, 500, { ok: false, error: '图库请求失败：' + (error?.message ?? String(error)) });
        }
    };
}

/** 给 registerMemoryView 用的注册清单（挂载时逐条 webServer.register）。 */
export function galleryRouteRegistrations(handler) {
    return GALLERY_ROUTE_PATHS.map((path) => ({ kind: 'exact', path, handler }));
}
