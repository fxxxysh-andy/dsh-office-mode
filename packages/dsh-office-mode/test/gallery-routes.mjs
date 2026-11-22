/**
 * 图库 HTTP 端点的测试（不启动服务器）。
 *
 * 覆盖：信任边界、只按 id 取图、字节原样返回、HEAD、手工标签写入、
 * 上传 base64、未确认的视觉/探针请求、错误码与错误文案。
 *
 * 跑法：node test/gallery-routes.mjs
 */
import assert from 'node:assert/strict';
import { Readable } from 'node:stream';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { deflateSync } from 'node:zlib';
import { createGallery } from '../src/gallery.js';
import { GALLERY_PATH, createGalleryHandler } from '../src/gallery-routes.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
        let c = n;
        for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

function crc32(buffer) {
    let crc = 0xffffffff;
    for (const byte of buffer) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
    const length = Buffer.alloc(4);
    length.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([length, body, crc]);
}

function pngOf(width, height, rgb = [10, 20, 30]) {
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;
    const stride = width * 3 + 1;
    const raw = Buffer.alloc(stride * height);
    for (let y = 0; y < height; y += 1) {
        raw[y * stride] = 0;
        for (let x = 0; x < width; x += 1) {
            const at = y * stride + 1 + x * 3;
            raw[at] = rgb[0];
            raw[at + 1] = rgb[1];
            raw[at + 2] = rgb[2];
        }
    }
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw)),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

function mockRequest({ method = 'GET', url = '/', headers = {}, body = null } = {}) {
    const stream = Readable.from(body === null ? [] : [Buffer.from(body)]);
    stream.method = method;
    stream.url = url;
    stream.headers = headers;
    return stream;
}

function mockResponse() {
    return {
        statusCode: 0,
        headers: {},
        chunks: [],
        ended: false,
        writeHead(status, headers = {}) {
            this.statusCode = status;
            this.headers = { ...headers };
        },
        end(chunk) {
            if (chunk !== undefined && chunk !== null) this.chunks.push(Buffer.from(chunk));
            this.ended = true;
        },
        json() {
            return JSON.parse(Buffer.concat(this.chunks).toString('utf8') || '{}');
        },
        bytes() {
            return Buffer.concat(this.chunks);
        },
    };
}

const root = await mkdtemp(join(tmpdir(), 'dsh-gallery-routes-'));
const work = join(root, 'work');
await writeFile(join(root, 'placeholder.txt'), 'x');
import { mkdir } from 'node:fs/promises';
await mkdir(work, { recursive: true });

const png = pngOf(4, 3, [90, 120, 150]);

function makeGallery(options = {}) {
    return createGallery({
        root: join(root, 'gallery'),
        cwd: work,
        roots: [work],
        thumbnailer: async ({ targetPath }) => {
            await writeFile(targetPath, Buffer.from([0xff, 0xd8, 0xff, 0xe0]));
            return { ok: true, mediaType: 'image/jpeg' };
        },
        config: { textEmbedding: { mode: 'local' } },
        ...options,
    });
}

let gallery = makeGallery();
let memoryConfig = {};
let trustResult = { ok: true, status: 200 };
const handler = createGalleryHandler({
    galleryFor: () => gallery,
    memoryOf: () => memoryConfig,
    trustedRootFor: async () => trustResult,
    hostOf: () => null,
});
const call = async (options) => {
    const req = mockRequest(options);
    const res = mockResponse();
    await handler(req, res);
    return res;
};
const post = (body, url = GALLERY_PATH) => call({
    method: 'POST',
    url,
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(body),
});

await check('GET status：如实返回未配置能力', async () => {
    const res = await call({ url: GALLERY_PATH + '?view=status' });
    assert.equal(res.statusCode, 200);
    const payload = res.json();
    assert.equal(payload.ok, true);
    assert.equal(payload.imageEmbedding.configured, false);
    assert.equal(payload.counts.items, 0);
});

let addedId = '';
await check('POST add：base64 上传入库，GET image 原字节返回、HEAD 只回头', async () => {
    const res = await post({ action: 'add', data: png.toString('base64'), name: '测试.png', tags: ['参考'], category: '风景' });
    assert.equal(res.statusCode, 200, JSON.stringify(res.json()));
    const payload = res.json();
    assert.equal(payload.added, true);
    assert.equal(payload.item.name, '测试.png', '上传时的文件名要成为图片名称');
    addedId = payload.item.id;

    const image = await call({ url: GALLERY_PATH + '?view=image&id=' + addedId });
    assert.equal(image.statusCode, 200);
    assert.equal(image.headers['content-type'], 'image/png');
    assert.equal(image.headers['x-content-type-options'], 'nosniff');
    assert.deepEqual(image.bytes(), png, '端点必须原样返回原图字节');

    const head = await call({ method: 'HEAD', url: GALLERY_PATH + '?view=image&id=' + addedId });
    assert.equal(head.statusCode, 200);
    assert.equal(head.bytes().length, 0);
    assert.equal(head.headers['content-length'], String(png.length));

    const thumb = await call({ url: GALLERY_PATH + '?view=thumb&id=' + addedId });
    assert.equal(thumb.statusCode, 200);
    assert.equal(thumb.headers['content-type'], 'image/jpeg');

    const item = await call({ url: GALLERY_PATH + '?view=item&id=' + addedId });
    assert.equal(item.statusCode, 200);
    assert.equal(item.json().item.category, '风景');
    assert.deepEqual(item.json().item.manualTags, ['参考']);
});

await check('GET list / search：按标签过滤与查询', async () => {
    const list = await call({ url: GALLERY_PATH + '?view=list&tags=参考' });
    assert.equal(list.statusCode, 200);
    assert.equal(list.json().total, 1);

    const searched = await call({ url: GALLERY_PATH + '?q=' + encodeURIComponent('参考') });
    assert.equal(searched.statusCode, 200);
    assert.equal(searched.json().items[0].id, addedId);
    assert.equal(searched.json().signals.tag.used, true);
});

await check('POST manual：改名称、描述、手工标签、分类、收藏', async () => {
    const res = await post({ action: 'manual', id: addedId, patch: {
        name: '改名后的图', description: '用户改写的描述',
        addTags: ['日落'], category: '情绪', favorite: true,
    } });
    assert.equal(res.statusCode, 200);
    assert.equal(res.json().item.name, '改名后的图');
    assert.equal(res.json().item.description, '用户改写的描述');
    assert.deepEqual(res.json().item.manualTags, ['参考', '日落']);
    assert.equal(res.json().item.favorite, true);
    assert.equal((await post({ action: 'manual', id: 'f'.repeat(64), favorite: true })).statusCode, 404);
});

await check('文件夹：创建、入库归类、过滤、移动、重命名与删除', async () => {
    const created = await post({ action: 'folder-create', path: '参考/风景' });
    assert.equal(created.statusCode, 200, JSON.stringify(created.json()));
    assert.equal(created.json().created, true);

    const added = await post({ action: 'add', data: pngOf(5, 5, [3, 4, 5]).toString('base64'), folder: '参考/风景' });
    assert.equal(added.statusCode, 200, JSON.stringify(added.json()));
    const id = added.json().item.id;
    assert.equal(added.json().item.folder, '参考/风景');

    const filtered = await call({ url: GALLERY_PATH + '?view=list&folder=' + encodeURIComponent('参考/风景') });
    assert.equal(filtered.statusCode, 200);
    assert.equal(filtered.json().total, 1);
    assert.equal(filtered.json().items[0].folder, '参考/风景');

    const moved = await post({ action: 'manual', id, patch: { folder: '参考' } });
    assert.equal(moved.statusCode, 200);
    assert.equal(moved.json().item.folder, '参考');

    const renamed = await post({ action: 'folder-rename', from: '参考', to: '资料' });
    assert.equal(renamed.statusCode, 200, JSON.stringify(renamed.json()));
    assert.equal(renamed.json().moved, 1);
    assert.equal((await call({ url: GALLERY_PATH + '?view=item&id=' + id })).json().item.folder, '资料');

    const blocked = await post({ action: 'folder-remove', path: '资料' });
    assert.equal(blocked.statusCode, 400);
    assert.equal(blocked.json().code, 'folder-not-empty');

    await post({ action: 'manual', id, patch: { folder: '' } });
    const empty = await post({ action: 'folder-remove', path: '资料' });
    assert.equal(empty.statusCode, 200, JSON.stringify(empty.json()));
    assert.equal(empty.json().removedFolders, 2, '父目录与子目录都是显式空目录');

    const bad = await post({ action: 'folder-create', path: '..' });
    assert.equal(bad.statusCode, 400);
    assert.equal(bad.json().code, 'bad-folder');
});

await check('POST add：非位图 base64 被拒绝，不产生条目', async () => {
    const res = await post({ action: 'add', data: Buffer.from('plain text').toString('base64') });
    assert.equal(res.statusCode, 400);
    assert.equal(res.json().code, 'unsupported');
});

await check('POST describe / probe：未确认时明确 needs-confirmation', async () => {
    const describe = await post({ action: 'describe', id: addedId });
    assert.equal(describe.statusCode, 400);
    assert.equal(describe.json().code, 'needs-confirmation');

    const probe = await post({ action: 'probe' });
    assert.equal(probe.statusCode, 400);
    assert.equal(probe.json().code, 'needs-confirmation');
});

await check('POST vectors：图片向量未配置时如实失败', async () => {
    const res = await post({ action: 'vectors', image: true, allowUpload: true });
    assert.equal(res.statusCode, 400);
    assert.equal(res.json().image.code, 'unconfigured');
});

await check('GET search / POST vectors：远端文本向量默认零请求，确认后才发', async () => {
    const localGallery = gallery;
    const calls = [];
    const remote = makeGallery({
        root: join(root, 'gallery-remote-text'),
        fetchImpl: async (url, init) => {
            const body = JSON.parse(init.body);
            calls.push({ url, body });
            return new Response(JSON.stringify({
                model: 'fake-text',
                data: body.input.map(() => ({ embedding: [1, 0, 0, 0] })),
            }), { status: 200, headers: { 'content-type': 'application/json' } });
        },
        config: { textEmbedding: { mode: 'auto', baseURL: 'https://fake.test/v1', model: 'fake-text', dimensions: 4, apiKey: 'k' } },
    });
    const added = await remote.addBytes(png, { name: '远端.png', tags: ['参考'] });
    assert.equal(added.ok, true);
    gallery = remote;
    try {
        const searched = await call({ url: GALLERY_PATH + '?q=' + encodeURIComponent('参考') });
        assert.equal(searched.statusCode, 200);
        assert.equal(calls.length, 0, '未确认时 GET 搜索不能发查询词');
        assert.equal(searched.json().signals.textVector.remoteBlocked, true);
        assert.equal(searched.json().signals.textVector.mode, 'local');

        const indexed = await post({ action: 'vectors', text: true });
        assert.equal(indexed.statusCode, 200);
        assert.equal(calls.length, 0, '未确认时 POST vectors 不能发描述/标签');
        assert.equal(indexed.json().text.mode, 'local');

        const confirmed = await post({ action: 'vectors', text: true, allowUpload: true });
        assert.equal(confirmed.statusCode, 200);
        assert.equal(calls.length, 1, '确认后才发');
        assert.equal(confirmed.json().text.mode, 'remote');
    } finally {
        gallery = localGallery;
    }
});

await check('POST remove：删条目后取图 404', async () => {
    const res = await post({ action: 'remove', id: addedId });
    assert.equal(res.statusCode, 200);
    assert.equal(res.json().ok, true);
    const image = await call({ url: GALLERY_PATH + '?view=image&id=' + addedId });
    assert.equal(image.statusCode, 404);
});

await check('图片 id 只认 64 位十六进制；非法/不存在都不碰文件系统', async () => {
    assert.equal((await call({ url: GALLERY_PATH + '?view=image&id=../../etc/passwd' })).statusCode, 400);
    assert.equal((await call({ url: GALLERY_PATH + '?view=image' })).statusCode, 400);
    assert.equal((await call({ url: GALLERY_PATH + '?view=image&id=' + 'a'.repeat(64) })).statusCode, 404);
    assert.equal((await call({ url: GALLERY_PATH + '?view=thumb&id=' + 'a'.repeat(64) })).statusCode, 404);
});

await check('信任边界：trustedRootFor 拒绝时 403，记忆停用时也 403', async () => {
    trustResult = { ok: false, status: 403, error: '工作目录不在信任范围内。' };
    const rejected = await call({ url: GALLERY_PATH + '?view=status' });
    assert.equal(rejected.statusCode, 403);
    assert.match(rejected.json().error, /信任范围/);

    trustResult = { ok: true, status: 200 };
    memoryConfig = { enabled: false };
    const disabled = await call({ url: GALLERY_PATH + '?view=status' });
    assert.equal(disabled.statusCode, 403);
    assert.match(disabled.json().error, /记忆已在设置里关掉/);
    memoryConfig = {};
});

await check('协议错误：非 JSON content-type 415、未知 action 400、错误方法 405', async () => {
    const badType = await call({
        method: 'POST',
        url: GALLERY_PATH,
        headers: { 'content-type': 'text/plain' },
        body: '{}',
    });
    assert.equal(badType.statusCode, 415);

    const unknown = await post({ action: 'nope' });
    assert.equal(unknown.statusCode, 400);
    assert.ok(unknown.json().actions.includes('add'));

    const wrongMethod = await call({ method: 'PUT', url: GALLERY_PATH });
    assert.equal(wrongMethod.statusCode, 405);
});

await check('请求体超限：明确 400，不把内存当垃圾桶', async () => {
    const small = createGalleryHandler({
        galleryFor: () => gallery,
        memoryOf: () => ({}),
        bodyLimit: 64,
    });
    const req = mockRequest({
        method: 'POST',
        url: GALLERY_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ action: 'add', data: 'x'.repeat(500) }),
    });
    const res = mockResponse();
    await small(req, res);
    assert.equal(res.statusCode, 400);
    assert.match(res.json().error, /读不了请求体/);
});

await check('图库未初始化：503 且文案明确', async () => {
    const none = createGalleryHandler({ galleryFor: () => null });
    const res = mockResponse();
    await none(mockRequest({ url: GALLERY_PATH + '?view=status' }), res);
    assert.equal(res.statusCode, 503);
    assert.match(res.json().error, /还没初始化/);
});

await check('POST probe（已配置 + 假端点）：确认后 200 并回报维度', async () => {
    const fetchImpl = async (url, init) => {
        const body = JSON.parse(init.body);
        const isText = body.inputs[0].content[0].type === 'text';
        const vector = isText ? [1, 0, 0, 0] : [1, 0, 0, 0];
        return {
            ok: true,
            status: 200,
            text: async () => JSON.stringify({ data: [{ embedding: vector, index: 0 }], model: 'voyage-multimodal-3.5' }),
        };
    };
    const mm = makeGallery({
        root: join(root, 'gallery-mm'),
        fetchImpl,
        config: { imageEmbedding: { provider: 'voyage', baseURL: 'https://api.voyageai.com/v1', model: 'voyage-multimodal-3.5' } },
    });
    const mmHandler = createGalleryHandler({ galleryFor: () => mm });
    const res = mockResponse();
    await mmHandler(mockRequest({
        method: 'POST',
        url: GALLERY_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ action: 'probe', allowUpload: true }),
    }), res);
    assert.equal(res.statusCode, 200);
    assert.equal(res.json().ok, true);
    assert.equal(res.json().dim, 4);
});

await check('manifest 被篡改成越界路径：只认规范相对路径，404 且不读根外文件', async () => {
    const tamperRoot = join(root, 'gallery-tamper');
    const first = createGallery({
        root: tamperRoot,
        cwd: work,
        roots: [work],
        thumbnailer: async ({ targetPath }) => {
            await writeFile(targetPath, Buffer.from([0xff, 0xd8, 0xff, 0xe0]));
            return { ok: true, mediaType: 'image/jpeg' };
        },
    });
    const added = await first.addBytes(png);
    assert.equal(added.ok, true);
    const id = added.item.id;
    const manifestPath = join(tamperRoot, 'manifest.json');
    const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
    manifest.items[id].original = 'originals/../../outside.png';
    await writeFile(manifestPath, JSON.stringify(manifest));

    const reopened = createGallery({ root: tamperRoot, cwd: work, roots: [work] });
    const tampered = createGalleryHandler({ galleryFor: () => reopened });
    const res = mockResponse();
    await tampered(mockRequest({ url: GALLERY_PATH + '?view=image&id=' + id }), res);
    assert.equal(res.statusCode, 404);
    assert.match(res.json().error, /不合法/);
});

await check('manifest 锁竞争：add/manual/remove 返回 503 而不是 400/404', async () => {
    const lockRoot = join(root, 'gallery-locked');
    await mkdir(lockRoot, { recursive: true });
    const seeded = createGallery({ root: lockRoot, cwd: work, roots: [work], thumbnailer: null });
    const seededAdd = await seeded.addBytes(png);
    assert.equal(seededAdd.ok, true);
    await writeFile(join(lockRoot, 'manifest.lock'), JSON.stringify({
        token: 'route-live-lock',
        pid: process.pid,
        host: 'test-host',
        at: new Date().toISOString(),
        kind: 'gallery-manifest',
    }));
    const locked = createGallery({
        root: lockRoot,
        cwd: work,
        roots: [work],
        thumbnailer: null,
        lockTimeoutMs: 120,
    });
    const lockedHandler = createGalleryHandler({ galleryFor: () => locked });
    const lockedCall = async (body) => {
        const res = mockResponse();
        await lockedHandler(mockRequest({
            method: 'POST',
            url: GALLERY_PATH,
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify(body),
        }), res);
        return res;
    };

    const addRes = await lockedCall({ action: 'add', data: png.toString('base64') });
    assert.equal(addRes.statusCode, 503);
    assert.equal(addRes.json().code, 'lock-timeout');

    const id = seededAdd.item.id;
    const manualRes = await lockedCall({ action: 'manual', id, note: 'x' });
    assert.equal(manualRes.statusCode, 503);
    assert.equal(manualRes.json().code, 'lock-timeout');

    const removeRes = await lockedCall({ action: 'remove', id });
    assert.equal(removeRes.statusCode, 503);
    assert.equal(removeRes.json().code, 'lock-timeout');
});

await rm(root, { recursive: true, force: true });

const failed = results.filter((result) => !result.ok);
for (const result of results) {
    console.log((result.ok ? 'ok   ' : 'FAIL ') + result.name);
    if (!result.ok) console.log('     ' + (result.error?.stack ?? result.error));
}
console.log('');
console.log(results.length - failed.length + '/' + results.length + ' passed');
if (failed.length > 0) process.exitCode = 1;
