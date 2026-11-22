/**
 * 审美图库（独立模块）的测试。
 *
 * 全部离线：
 *   - 原图用测试里现造的最小 PNG，缩略图用假生成器，不依赖 Pillow；
 *   - 文本向量走 kb-embed 的本地哈希向量（明确标注它不是语义模型）；
 *   - 多模态端点用假 fetch 按 Voyage 真实响应形状返回，验证请求体与维度契约；
 *   - 视觉通道用假宿主 llm / attachments 与假 OpenAI 兼容响应。
 *
 * 跑法：node test/gallery.mjs
 */
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, rm, stat, writeFile, readdir } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { homedir, hostname, tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { deflateSync } from 'node:zlib';
import {
    MAX_TAGS,
    createGallery,
    createPillowThumbnailer,
    documentTextOf,
    galleryRootDir,
    hashBytes,
    imageVectorKey,
    lexicalScores,
    normalizeAesthetic,
    normalizeFolderPath,
    normalizeTags,
    tagScoreOf,
    textVectorKey,
} from '../src/gallery.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

// ── 造一张最小 PNG（RGB，无依赖） ─────────────────────────────────────────
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

function pngOf(width, height, rgb = [200, 120, 80]) {
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

// 只构造合法头部的超大像素 PNG：测试选图逻辑用，图库只按魔数读宽高，不解码。
function headerOnlyPng(width, height) {
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(Buffer.from([0]))),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

function fakeThumbnailer() {
    return async function thumbnail({ targetPath }) {
        await writeFile(targetPath, Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10]));
        return { ok: true, mediaType: 'image/jpeg' };
    };
}

function vec(values) {
    const vector = new Float32Array(8);
    for (let i = 0; i < values.length; i += 1) vector[i] = values[i];
    return vector;
}

const root = await mkdtemp(join(tmpdir(), 'dsh-gallery-test-'));
const cwd = join(root, 'work');
await mkdir(cwd, { recursive: true });
const galleryRoot = join(root, 'gallery');
const outsideDir = await mkdtemp(join(tmpdir(), 'dsh-gallery-outside-'));

function makeGallery(options = {}) {
    return createGallery({
        root: galleryRoot,
        cwd,
        roots: [cwd],
        thumbnailer: fakeThumbnailer(),
        ...options,
    });
}

const gallery = makeGallery({ config: { textEmbedding: { mode: 'local' } } });

await check('addBytes：保存原图、事实字段与缩略图', async () => {
    const bytes = pngOf(4, 3, [10, 20, 30]);
    const added = await gallery.addBytes(bytes, { tags: ['参考', '海边'], category: '风景', favorite: true, note: '手工备注' });
    assert.equal(added.ok, true);
    assert.equal(added.added, true);
    assert.equal(added.duplicate, false);
    assert.equal(added.item.width, 4);
    assert.equal(added.item.height, 3);
    assert.equal(added.item.mediaType, 'image/png');
    assert.equal(added.item.bytes, bytes.length);
    assert.equal(added.item.thumb, 'thumbs/' + hashBytes(bytes) + '.jpg');
    assert.deepEqual(added.item.manualTags, ['参考', '海边']);
    assert.equal(added.item.favorite, true);
    assert.equal(added.item.category, '风景');
    const original = await readFile(join(galleryRoot, 'originals', hashBytes(bytes) + '.png'));
    assert.equal(hashBytes(original), hashBytes(bytes), '原图字节必须原样保存');
    assert.ok((await stat(join(galleryRoot, 'manifest.json'))).size > 0);
});

await check('addBytes：同内容去重，不重复写盘', async () => {
    const bytes = pngOf(4, 3, [10, 20, 30]);
    const again = await gallery.addBytes(bytes, { tags: ['不该被追加'] });
    assert.equal(again.ok, true);
    assert.equal(again.added, false);
    assert.equal(again.duplicate, true);
    assert.deepEqual(again.item.manualTags, ['参考', '海边'], '重复添加不应隐式改手工标签');
});

await check('addBytes：拒绝空文件、非位图、超限与超大像素', async () => {
    assert.equal((await gallery.addBytes(Buffer.alloc(0))).code, 'empty');
    assert.equal((await gallery.addBytes(Buffer.from('plain text'))).code, 'unsupported');
    const small = makeGallery({ maxBytes: 10 });
    assert.equal((await small.addBytes(pngOf(2, 2))).code, 'too-large');
    const huge = Buffer.from([
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x4e, 0x20, 0x00, 0x00, 0x4e, 0x20, 8, 2, 0, 0, 0,
    ]);
    assert.equal((await gallery.addBytes(huge)).code, 'too-many-pixels');
});

await check('addFromPath：只接受信任根内的普通文件', async () => {
    const inside = join(cwd, 'inside.png');
    await writeFile(inside, pngOf(2, 2, [1, 2, 3]));
    const added = await gallery.addFromPath('inside.png');
    assert.equal(added.ok, true);
    assert.equal(added.item.mediaType, 'image/png');
    assert.equal(added.item.name, 'inside.png', '路径导入默认用文件名当名称');

    const outside = join(outsideDir, 'outside.png');
    await writeFile(outside, pngOf(2, 2, [4, 5, 6]));
    const rejected = await gallery.addFromPath(outside);
    assert.equal(rejected.ok, false);
    assert.equal(rejected.code, 'outside-root');

    assert.equal((await gallery.addFromPath(join(cwd, 'missing.png'))).code, 'not-found');
    assert.equal((await gallery.addFromPath(cwd)).code, 'not-file');
    assert.equal((await gallery.addFromPath('')).code, 'bad-path');
});

await check('galleryRootDir：显式目录优先，相对路径相对用户主目录，否则与知识库同级', async () => {
    // 盘符只在 Windows 上是「绝对路径」；POSIX 上它是一段普通相对路径。
    // 这一条要跨平台成立，所以两种语义都按当前平台算期望值，不写死 Windows 的形状。
    const isWindows = process.platform === 'win32';
    const kbRoot = join(tmpdir(), 'dsh-gallery-root-kb');
    assert.equal(
        galleryRootDir({ galleryDir: 'D:/data/gallery' }, kbRoot),
        isWindows ? 'D:\\data\\gallery' : resolve(homedir(), 'D:/data/gallery'),
    );
    assert.equal(galleryRootDir({ galleryDir: 'pics/gallery' }, kbRoot), resolve(homedir(), 'pics/gallery'));
    assert.equal(galleryRootDir({}, kbRoot), join(dirname(resolve(kbRoot)), 'gallery'));
    assert.equal(galleryRootDir({}, ''), '');
});

await check('manifest：重新打开图库能读回同一批条目', async () => {
    const reopened = makeGallery({ config: { textEmbedding: { mode: 'local' } } });
    const listed = await reopened.list({ limit: 10 });
    assert.equal(listed.ok, true);
    assert.ok(listed.total >= 2);
    const detail = await reopened.get(listed.items[0].id);
    assert.ok(detail !== null);
    assert.equal(detail.facts.sha256, detail.id);
});

await check('多实例共享同一物理图库：交错读、改元数据、并发上传都不丢，外部改 manifest 能被读新鲜', async () => {
    const panel = makeGallery({ config: { textEmbedding: { mode: 'local' } } });
    const tool = makeGallery({ config: { textEmbedding: { mode: 'local' } } });
    const before = (await panel.list({ limit: 200 })).total;

    const addedByPanel = await panel.addBytes(pngOf(6, 5, [11, 22, 33]), { tags: ['面板'] });
    assert.equal(addedByPanel.ok, true);
    assert.ok(await tool.get(addedByPanel.item.id), '另一个实例必须立刻读得到新条目');

    const renamed = await tool.updateManual(addedByPanel.item.id, { category: '共享分类', note: '工具实例改的' });
    assert.equal(renamed.ok, true);
    const seenByPanel = await panel.get(addedByPanel.item.id);
    assert.equal(seenByPanel.manual.category, '共享分类');
    assert.equal(seenByPanel.manual.note, '工具实例改的');

    const batch = await Promise.all(Array.from({ length: 6 }, (_, index) => panel.addBytes(
        pngOf(3 + index, 3, [index * 30 + 1, index * 20 + 2, index * 10 + 3]),
        { tags: ['并发' + index] },
    )));
    assert.equal(batch.every((row) => row.ok === true && row.added === true), true);
    const listedByTool = await tool.list({ limit: 200 });
    for (const row of batch) assert.ok(listedByTool.items.some((item) => item.id === row.item.id), '并发写入后另一实例必须看得到：' + row.item.id);
    assert.equal(listedByTool.total, before + 1 + batch.length);

    // 模拟另一个进程直接改 manifest：长活实例下次读取必须按磁盘指纹重读，而不是拿内存旧值。
    const manifestPath = join(galleryRoot, 'manifest.json');
    const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
    manifest.items[addedByPanel.item.id].manual.note = '外部进程改的';
    await writeFile(manifestPath, JSON.stringify(manifest));
    const fresh = await panel.get(addedByPanel.item.id);
    assert.equal(fresh.manual.note, '外部进程改的');
});

await check('跨进程：6 个进程同时上传，成功条目全部可读且 manifest 不丢', async () => {
    const crossRoot = await mkdtemp(join(tmpdir(), 'dsh-gallery-cross-'));
    const moduleUrl = new URL('../src/gallery.js', import.meta.url).href;
    const childCode = [
        'const { createGallery } = await import(process.env.GALLERY_MODULE);',
        'const gallery = createGallery({ root: process.env.GALLERY_ROOT, cwd: process.env.GALLERY_ROOT, roots: [process.env.GALLERY_ROOT], thumbnailer: null, lockTimeoutMs: 60000 });',
        'const result = await gallery.addBytes(Buffer.from(process.env.GALLERY_IMAGE, "base64"), { tags: [process.env.GALLERY_TAG] });',
        'console.log(JSON.stringify({ ok: result.ok, added: result.added, id: result.item && result.item.id, code: result.code, error: result.error }));',
        'process.exitCode = result.ok ? 0 : 1;',
    ].join('\n');
    const children = Array.from({ length: 6 }, (_, index) => new Promise((resolveChild) => {
        const child = spawn(process.execPath, ['--max-old-space-size=96', '--jitless', '--input-type=module', '-e', childCode], {
            env: {
                ...process.env,
                GALLERY_MODULE: moduleUrl,
                GALLERY_ROOT: crossRoot,
                GALLERY_IMAGE: pngOf(4 + index, 4, [10 + index * 20, 20, 30]).toString('base64'),
                GALLERY_TAG: '进程' + index,
            },
            stdio: ['ignore', 'pipe', 'pipe'],
        });
        let out = '';
        let err = '';
        let settled = false;
        const finish = (row) => {
            if (settled) return;
            settled = true;
            resolveChild({ index, out, err, ...row });
        };
        child.stdout.on('data', (chunk) => { out += String(chunk); });
        child.stderr.on('data', (chunk) => { err += String(chunk); });
        child.on('error', (error) => finish({ code: -1, error }));
        child.on('close', (code) => finish({ code }));
    }));
    const rows = await Promise.all(children);
    const ids = new Set();
    for (const row of rows) {
        assert.equal(row.code, 0, '子进程 ' + row.index + ' 失败：' + row.out + row.err);
        const line = row.out.trim().split(/\r?\n/).filter(Boolean).pop();
        const parsed = JSON.parse(line);
        assert.equal(parsed.ok, true, '子进程 ' + row.index + ' 回执：' + line);
        assert.equal(parsed.added, true);
        ids.add(parsed.id);
    }
    assert.equal(ids.size, 6, '6 个进程必须产生 6 个不同条目');

    const reopened = createGallery({ root: crossRoot, cwd: crossRoot, roots: [crossRoot], thumbnailer: null });
    const listed = await reopened.list({ limit: 100 });
    assert.equal(listed.total, 6, '所有成功上传都必须还能读到');
    assert.deepEqual(new Set(listed.items.map((item) => item.id)), ids);
    const manifest = JSON.parse(await readFile(join(crossRoot, 'manifest.json'), 'utf8'));
    assert.equal(Object.keys(manifest.items).length, 6, 'manifest 不能少条');
    await rm(crossRoot, { recursive: true, force: true });
});

await check('缩略图在 manifest 锁外：跨进程慢缩略图不会让另一个上传 lock-timeout', async () => {
    const slowRoot = await mkdtemp(join(tmpdir(), 'dsh-gallery-slow-thumb-'));
    await mkdir(slowRoot, { recursive: true });
    const moduleUrl = new URL('../src/gallery.js', import.meta.url).href;
    const childCode = [
        'const { createGallery } = await import(process.env.GALLERY_MODULE);',
        'const { writeFile } = await import("node:fs/promises");',
        'const slowThumb = async ({ targetPath }) => {',
        '  await new Promise((resolve) => setTimeout(resolve, 1200));',
        '  await writeFile(targetPath, Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10]));',
        '  return { ok: true, mediaType: "image/jpeg" };',
        '};',
        'const gallery = createGallery({ root: process.env.GALLERY_ROOT, cwd: process.env.GALLERY_ROOT, roots: [process.env.GALLERY_ROOT], thumbnailer: slowThumb, lockTimeoutMs: 400 });',
        'const startAt = Number(process.env.START_AT);',
        'while (Date.now() < startAt) await new Promise((resolve) => setTimeout(resolve, 5));',
        'const result = await gallery.addBytes(Buffer.from(process.env.GALLERY_IMAGE, "base64"));',
        'console.log(JSON.stringify({ ok: result.ok, added: result.added, id: result.item && result.item.id, code: result.code, error: result.error }));',
        'process.exitCode = result.ok ? 0 : 1;',
    ].join('\n');
    const startAt = Date.now() + 400;
    const children = [0, 1].map((index) => new Promise((resolveChild) => {
        const child = spawn(process.execPath, ['--max-old-space-size=96', '--jitless', '--input-type=module', '-e', childCode], {
            env: {
                ...process.env,
                GALLERY_MODULE: moduleUrl,
                GALLERY_ROOT: slowRoot,
                GALLERY_IMAGE: pngOf(5 + index, 4, [30 + index * 20, 10, 40]).toString('base64'),
                START_AT: String(startAt),
            },
            stdio: ['ignore', 'pipe', 'pipe'],
        });
        let out = '';
        let err = '';
        let settled = false;
        const finish = (row) => {
            if (settled) return;
            settled = true;
            resolveChild({ index, out, err, ...row });
        };
        child.stdout.on('data', (chunk) => { out += String(chunk); });
        child.stderr.on('data', (chunk) => { err += String(chunk); });
        child.on('error', (error) => finish({ code: -1, error }));
        child.on('close', (code) => finish({ code }));
    }));
    const rows = await Promise.all(children);
    for (const row of rows) {
        assert.equal(row.code, 0, '慢缩略图子进程 ' + row.index + ' 失败：' + row.out + row.err);
        const line = row.out.trim().split(/\r?\n/).filter(Boolean).pop();
        const parsed = JSON.parse(line);
        assert.equal(parsed.ok, true, '慢缩略图上传必须成功：' + line);
        assert.equal(parsed.added, true);
    }
    const reopened = createGallery({ root: slowRoot, cwd: slowRoot, roots: [slowRoot], thumbnailer: null });
    const listed = await reopened.list({ limit: 100 });
    assert.equal(listed.total, 2, '两个慢缩略图上传都必须落进 manifest');
    await rm(slowRoot, { recursive: true, force: true });
});

await check('跨进程锁：活锁/残锁不自动删，超时明确失败并给手动恢复路径', async () => {
    const lockRoot = await mkdtemp(join(tmpdir(), 'dsh-gallery-lock-'));
    await mkdir(lockRoot, { recursive: true });
    const lockPath = join(lockRoot, 'manifest.lock');
    await writeFile(lockPath, JSON.stringify({
        token: 'live-lock',
        pid: process.pid,
        host: hostname(),
        at: new Date().toISOString(),
        kind: 'gallery-manifest',
    }));
    const blocked = createGallery({
        root: lockRoot,
        cwd: lockRoot,
        roots: [lockRoot],
        thumbnailer: null,
        lockTimeoutMs: 150,
    });
    const result = await blocked.addBytes(pngOf(2, 2, [1, 2, 3]));
    assert.equal(result.ok, false);
    assert.equal(result.code, 'lock-timeout');
    assert.ok(result.error.includes(lockPath), '错误里要给出锁文件路径');
    assert.match(result.error, /手动删除/);
    assert.match(result.error, /不会自动删除/);
    assert.equal(JSON.parse(await readFile(lockPath, 'utf8')).token, 'live-lock', '不能误删活锁');

    const staleRoot = join(lockRoot, 'stale');
    await mkdir(staleRoot, { recursive: true });
    const stalePath = join(staleRoot, 'manifest.lock');
    await writeFile(stalePath, JSON.stringify({
        token: 'stale-lock',
        pid: 999999,
        host: hostname(),
        at: '2000-01-01T00:00:00.000Z',
        kind: 'gallery-manifest',
    }));
    const staleGallery = createGallery({
        root: staleRoot,
        cwd: staleRoot,
        roots: [staleRoot],
        thumbnailer: null,
        lockTimeoutMs: 120,
    });
    const staleResult = await staleGallery.addBytes(pngOf(2, 2, [4, 5, 6]));
    assert.equal(staleResult.ok, false);
    assert.equal(staleResult.code, 'lock-timeout');
    assert.match(staleResult.error, /崩溃残留/);
    assert.equal(JSON.parse(await readFile(stalePath, 'utf8')).token, 'stale-lock', '残锁也不自动删，交给人工确认');

    const emptyRoot = join(lockRoot, 'empty');
    await mkdir(emptyRoot, { recursive: true });
    const emptyPath = join(emptyRoot, 'manifest.lock');
    await writeFile(emptyPath, '');
    const emptyGallery = createGallery({
        root: emptyRoot,
        cwd: emptyRoot,
        roots: [emptyRoot],
        thumbnailer: null,
        lockTimeoutMs: 120,
    });
    const emptyResult = await emptyGallery.addBytes(pngOf(2, 2, [7, 8, 9]));
    assert.equal(emptyResult.ok, false);
    assert.equal(emptyResult.code, 'lock-timeout');
    assert.match(emptyResult.error, /锁文件为空/);
    assert.equal((await readFile(emptyPath, 'utf8')), '', '空锁也不自动删');
    await rm(lockRoot, { recursive: true, force: true });
});

await check('updateManual：标签去重、增删、分类、收藏、备注限长', async () => {
    const bytes = pngOf(3, 3, [9, 9, 9]);
    const added = await gallery.addBytes(bytes);
    const id = added.item.id;
    const updated = await gallery.updateManual(id, {
        addTags: ['日落', '日落', ' 海边 '],
        category: ' 情绪 ',
        favorite: true,
        note: 'x'.repeat(5000),
    });
    assert.equal(updated.ok, true);
    assert.deepEqual(updated.item.manualTags, ['日落', '海边']);
    assert.equal(updated.item.category, '情绪');
    assert.equal(updated.item.favorite, true);
    assert.equal(updated.item.note.length, 2000);

    const removed = await gallery.updateManual(id, { removeTags: ['日落'] });
    assert.deepEqual(removed.item.manualTags, ['海边']);

    const replaced = await gallery.updateManual(id, { tags: 'a,b,a' });
    assert.deepEqual(replaced.item.manualTags, ['a', 'b']);
    assert.equal((await gallery.updateManual('f'.repeat(64), { favorite: true })).code, 'not-found');
});

await check('updateManual：名称与可编辑描述、留空回落模型总结', async () => {
    const added = await gallery.addBytes(pngOf(4, 3, [21, 31, 41]), {
        name: '  初始名称  ',
        description: 'y'.repeat(900),
    });
    assert.equal(added.item.name, '初始名称');
    assert.equal(added.item.description.length, 600);
    await gallery.setAesthetic(added.item.id, { subject: '模型主体', summary: '模型总结' }, {
        channel: 'test', provider: 'fake', model: 'fake-vl', promptVersion: 'v1',
    });
    const renamed = await gallery.updateManual(added.item.id, { name: '  新名字  ', description: '用户改写' });
    assert.equal(renamed.item.name, '新名字');
    assert.equal(renamed.item.description, '用户改写');
    const cleared = await gallery.updateManual(added.item.id, { name: '', description: '' });
    assert.equal(cleared.item.name, '');
    assert.equal(cleared.item.description, '模型总结', '描述留空时应回落到模型总结');
});

await check('documentTextOf / 检索：名称与改写描述进入文本，改写后不再重复模型总结', async () => {
    const item = {
        id: 'b'.repeat(64),
        facts: { sha256: 'b'.repeat(64), bytes: 2 },
        aesthetic: { summary: '模型总结', tags: ['模型标签'] },
        manual: { name: '产品首页', description: '用户改写描述', tags: [], category: '', note: '' },
    };
    const text = documentTextOf(item);
    assert.ok(text.includes('产品首页'));
    assert.ok(text.includes('用户改写描述'));
    assert.ok(!text.includes('模型总结'), '有用户改写时不再重复模型总结');
    const added = await gallery.addBytes(pngOf(7, 7, [31, 41, 51]), {
        name: '可检索的名称',
        description: '可检索的改写描述',
    });
    const hit = await gallery.search({ query: '可检索的改写描述' });
    assert.equal(hit.items[0].id, added.item.id, '名称与描述要参与词法检索');
});

await check('normalizeTags / normalizeAesthetic：限长、去重、丢弃作者与版权字段', async () => {
    const tags = normalizeTags(Array.from({ length: 60 }, (_, index) => 'tag' + index));
    assert.equal(tags.length, MAX_TAGS);
    assert.deepEqual(normalizeTags('日落, 海边、日落; 暖色'), ['日落', '海边', '暖色']);

    const aesthetic = normalizeAesthetic({
        subject: '  海边   日落 ',
        composition: '三分构图',
        palette: ['暖橙', '暖橙', ' #ff8800 '],
        light: '逆光',
        material: '海水',
        mood: '安静',
        summary: '黄昏的海面',
        tags: '日落,海边,日落',
        author: '张三',
        copyright: '某公司',
    }, { channel: 'host', provider: 'x', model: 'm', promptVersion: 'v1', generatedAt: '2026-01-01T00:00:00.000Z' });
    assert.equal(aesthetic.subject, '海边 日落');
    assert.deepEqual(aesthetic.palette, ['暖橙', '#ff8800']);
    assert.deepEqual(aesthetic.tags, ['日落', '海边']);
    assert.deepEqual(aesthetic.droppedFields, ['author', 'copyright']);
    assert.equal(aesthetic.author, undefined);
    assert.equal(aesthetic.channel, 'host');
});

await check('documentTextOf：只拼描述与标签，不拼哈希/尺寸/来源等事实', async () => {
    const item = {
        id: 'a'.repeat(64),
        facts: { sha256: 'a'.repeat(64), bytes: 123456, width: 1920, height: 1080, origin: { path: 'C:/secret/photo.png' } },
        aesthetic: { subject: '海边日落', composition: '三分构图', palette: ['暖橙'], light: '逆光', material: '海水', mood: '安静', summary: '黄昏海面', tags: ['日落'] },
        manual: { tags: ['收藏'], category: '风景', note: '喜欢这张' },
    };
    const text = documentTextOf(item);
    assert.ok(text.includes('海边日落'));
    assert.ok(text.includes('收藏'));
    assert.ok(!text.includes('1920'), '尺寸不应进入检索文本');
    assert.ok(!text.includes('C:/secret'), '来源路径不应进入检索文本');
    assert.ok(!text.includes('a'.repeat(64)), '哈希不应进入检索文本');
});

await check('list / search：标签精确命中与词法排序', async () => {
    const first = await gallery.addBytes(pngOf(5, 4, [220, 160, 90]), {
        tags: ['日落', '海边'],
        category: '风景',
        favorite: true,
    });
    const second = await gallery.addBytes(pngOf(6, 4, [20, 20, 40]), { tags: ['夜景', '城市'] });
    await gallery.setAesthetic(first.item.id, {
        subject: '海边日落',
        composition: '三分构图',
        palette: ['暖橙', '淡紫'],
        light: '逆光',
        material: '海水',
        mood: '安静',
        summary: '黄昏的海面',
        tags: ['日落', '暖色'],
    }, { channel: 'test', provider: 'fake', model: 'fake-vision', promptVersion: 'v1' });
    await gallery.setAesthetic(second.item.id, {
        subject: '城市夜景',
        composition: '中心对称',
        palette: ['霓虹蓝'],
        light: '灯光',
        material: '玻璃',
        mood: '繁忙',
        summary: '雨夜街头',
        tags: ['夜景', '霓虹'],
    }, { channel: 'test', provider: 'fake', model: 'fake-vision', promptVersion: 'v1' });

    const byTag = await gallery.search({ query: '日落' });
    assert.equal(byTag.ok, true);
    assert.equal(byTag.items[0].id, first.item.id);
    assert.ok(byTag.items[0].parts.tag >= 0.6, '标签命中要体现在分项分数里');

    const byText = await gallery.search({ query: '海边 黄昏' });
    assert.equal(byText.items[0].id, first.item.id);
    assert.ok(byText.signals.lexical.used);

    const filtered = await gallery.list({ tags: ['夜景'] });
    assert.equal(filtered.total, 1);
    assert.equal(filtered.items[0].id, second.item.id);
    const favorite = await gallery.list({ favorite: true });
    assert.ok(favorite.items.some((item) => item.id === first.item.id));
});

await check('search：有查询词但无命中时返回空列表，不把零分图片当结果', async () => {
    const off = createGallery({
        root: join(root, 'gallery-lexical-only'),
        cwd,
        roots: [cwd],
        thumbnailer: fakeThumbnailer(),
        config: { textEmbedding: { mode: 'off' } },
    });
    await off.addBytes(pngOf(5, 4, [220, 160, 90]), { tags: ['日落'], category: '风景' });
    await off.addBytes(pngOf(6, 4, [20, 20, 40]), { tags: ['夜景'], category: '城市' });

    const miss = await off.search({ query: '量子纠缠' });
    assert.equal(miss.total, 0, '没有命中的查询不能把整库当结果');
    assert.deepEqual(miss.items, []);
    assert.ok(miss.reason, '空结果要说明原因，客户端才能显示「没有匹配」');

    const hit = await off.search({ query: '日落' });
    assert.equal(hit.total, 1);
    assert.ok(hit.items[0].score > 0);
    assert.equal(hit.items[0].parts.lexical > 0 || hit.items[0].parts.tag > 0, true);
    assert.ok(hit.items.every((item) => item.score > 0), '结果里不能混进零分候选');

    const filtered = await off.search({ tags: ['夜景'] });
    assert.equal(filtered.total, 1, '只有硬过滤、没有查询词时仍返回过滤结果');
    assert.equal(filtered.items[0].score, null, '没有查询词时不编造分数');
});

await check('词法打分 / 标签分：纯函数可复现', async () => {
    const items = [
        { id: '1', aesthetic: { subject: '海边日落', tags: ['日落'] }, manual: { tags: [] } },
        { id: '2', aesthetic: { subject: '城市夜景', tags: ['霓虹'] }, manual: { tags: [] } },
    ];
    const scores = lexicalScores(items, '日落');
    assert.ok(scores.get('1') > scores.get('2'));
    assert.equal(tagScoreOf(items[0], '日落'), 1);
    assert.equal(tagScoreOf(items[0], '日'), 0.6);
    assert.equal(tagScoreOf(items[1], '日落'), 0);
});

await check('ensureVectors(text)：写文本向量缓存；未配置时明确失败而不是写占位向量', async () => {
    const ok = await gallery.ensureVectors({ text: true });
    assert.equal(ok.ok, true);
    assert.ok(ok.text.written > 0);
    assert.equal(ok.text.model, 'local-hash-v1');
    assert.equal(ok.text.dim, 256);
    assert.equal(ok.text.degraded, false, '本地模式不应标成降级');

    const textDirs = await readdir(join(galleryRoot, 'vectors', 'text'));
    assert.ok(textDirs.length > 0);
    const key = textVectorKey('local-hash-v1', 256);
    const vectorFiles = await readdir(join(galleryRoot, 'vectors', 'text', key));
    assert.ok(vectorFiles.length > 0);

    const off = makeGallery({ config: { textEmbedding: { mode: 'off' } } });
    const disabled = await off.ensureVectors({ text: true });
    assert.equal(disabled.ok, false);
    assert.equal(disabled.text.code, 'disabled');
    await assert.rejects(readdir(join(galleryRoot, 'vectors', 'text', key, 'no-such-id.json')));
});

await check('search(textVector)：本地向量参与融合，并如实标出不是语义模型', async () => {
    const result = await gallery.search({ query: '黄昏的海面' });
    assert.equal(result.signals.textVector.used, true);
    assert.equal(result.signals.textVector.mode, 'local');
    assert.match(result.signals.textVector.note, /本地哈希向量|语义/);
    assert.ok(result.items[0].parts.textVector > 0);
});

await check('远端文本向量：未确认 allowUpload 时零请求并退回本地，确认后才发', async () => {
    const calls = [];
    const fetchImpl = async (url, init) => {
        const body = JSON.parse(init.body);
        calls.push({ url, body });
        return new Response(JSON.stringify({
            model: 'fake-text',
            data: body.input.map(() => ({ embedding: [1, 0, 0, 0] })),
        }), { status: 200, headers: { 'content-type': 'application/json' } });
    };
    const remote = makeGallery({
        fetchImpl,
        config: { textEmbedding: { mode: 'auto', baseURL: 'https://fake.test/v1', model: 'fake-text', dimensions: 4, apiKey: 'k' } },
    });
    const indexed = await remote.ensureVectors({ text: true, allowUpload: true });
    assert.equal(indexed.ok, true);
    assert.equal(indexed.text.mode, 'remote');
    assert.equal(calls.length, 1, '确认后应发一次文本向量请求');
    calls.length = 0;

    const localSearch = await remote.search({ query: '日落' });
    assert.equal(calls.length, 0, '未确认 allowUpload 时不能把查询词发出去');
    assert.equal(localSearch.signals.textVector.remoteBlocked, true);
    assert.equal(localSearch.signals.textVector.mode, 'local');

    const remoteSearch = await remote.search({ query: '日落', allowUpload: true });
    assert.equal(calls.length, 1, '确认后才发查询词');
    assert.equal(remoteSearch.signals.textVector.mode, 'remote');
    assert.equal(remoteSearch.signals.textVector.remoteBlocked, undefined);

    const st = await remote.status();
    assert.equal(st.textEmbedding.kind, 'text', '文本 embedding 与图文多模态是两类能力');
    assert.equal(st.imageEmbedding.kind, 'multimodal');
    assert.match(st.textEmbedding.note, /可能计费/);
    assert.match(st.textEmbedding.note, /两个能力/);
    assert.match(st.uploadPolicy, /图片，或查询词\/描述\/标签/);
    assert.match(st.uploadPolicy, /两个独立能力/);
});

await check('图片向量未配置：ensureVectors 明确 unconfigured，检索跳过图文信号且不与文本向量混比', async () => {
    const report = await gallery.ensureVectors({ image: true, allowUpload: true });
    assert.equal(report.ok, false);
    assert.equal(report.image.code, 'unconfigured');
    assert.match(report.image.error, /没有配置图文对齐/);
    await assert.rejects(readdir(join(galleryRoot, 'vectors', 'image')).then((list) => {
        if (list.length === 0) throw Object.assign(new Error('空目录'), { code: 'ENOENT' });
        return list;
    }));

    const result = await gallery.search({ query: '日落' });
    assert.equal(result.signals.imageText.used, false);
    assert.match(result.signals.imageText.reason, /没有配置图文对齐/);
    assert.equal(result.items[0].parts.imageText, undefined, '不可用的信号不能出现在分项分数里');
});

// ── 假的多模态端点：按 Voyage 的真实响应形状返回 ─────────────────────────
function voyageFetch() {
    const calls = [];
    const fetchImpl = async (url, init) => {
        const body = JSON.parse(init.body);
        calls.push({ url, body });
        const inputs = body.inputs ?? [];
        const vectors = inputs.map((input, index) => {
            const block = input.content?.[0] ?? {};
            if (block.type === 'text') {
                if (block.text === '审美参考图' || block.text.includes('日落')) return [1, 0, 0, 0, 0, 0, 0, 0];
                return [0, 1, 0, 0, 0, 0, 0, 0];
            }
            if (inputs.length > 1) {
                return index === 0 ? [1, 0, 0, 0, 0, 0, 0, 0] : [0, 1, 0, 0, 0, 0, 0, 0];
            }
            return [1, 0, 0, 0, 0, 0, 0, 0];
        });
        return {
            ok: true,
            status: 200,
            text: async () => JSON.stringify({
                object: 'list',
                data: vectors.map((vector, index) => ({ embedding: vector, index })),
                model: 'voyage-multimodal-3.5',
            }),
        };
    };
    return { fetchImpl, calls };
}

await check('多模态探针：未确认不上传；确认后要求文本与图片同维', async () => {
    const { fetchImpl, calls } = voyageFetch();
    const mm = makeGallery({
        fetchImpl,
        root: join(root, 'gallery-mm-probe'),
        config: { imageEmbedding: { provider: 'voyage', baseURL: 'https://api.voyageai.com/v1', model: 'voyage-multimodal-3.5' } },
    });
    const refused = await mm.probeImage({});
    assert.equal(refused.code, 'needs-confirmation');
    assert.equal(calls.length, 0, '未确认时不能发任何请求');

    const probed = await mm.probeImage({ allowUpload: true });
    assert.equal(probed.ok, true);
    assert.equal(probed.dim, 8);
    assert.equal(calls.length, 2);
    const textCall = calls.find((call) => call.body.inputs[0].content[0].type === 'text');
    const imageCall = calls.find((call) => call.body.inputs[0].content[0].type === 'image_base64');
    assert.ok(textCall && imageCall);
    assert.equal(textCall.body.input_type, 'query');
    assert.equal(textCall.body.model, 'voyage-multimodal-3.5');
    assert.match(textCall.url, /multimodalembeddings$/);
    assert.equal(imageCall.body.input_type, 'document');
    assert.match(imageCall.body.inputs[0].content[0].image_base64, /^data:image\/png;base64,/);
});

await check('多模态探针：一次瞬时失败不覆盖上次成功结果', async () => {
    const { fetchImpl } = voyageFetch();
    let fail = false;
    const flaky = async (url, init) => {
        if (fail) return { ok: false, status: 500, text: async () => 'upstream down' };
        return fetchImpl(url, init);
    };
    const mm = makeGallery({
        fetchImpl: flaky,
        root: join(root, 'gallery-mm-stale'),
        config: { imageEmbedding: { provider: 'voyage', baseURL: 'https://api.voyageai.com/v1', model: 'voyage-multimodal-3.5' } },
    });
    assert.equal((await mm.probeImage({ allowUpload: true })).ok, true);
    fail = true;
    const again = await mm.probeImage({ allowUpload: true });
    assert.equal(again.ok, true, '瞬时失败不能把成功探针抹掉');
    assert.equal(again.stale, true);
    assert.match(again.lastError, /500/);
    const status = await mm.status();
    assert.equal(status.imageEmbedding.probed, true, 'status 仍应报告上次成功探针');
});

await check('多模态探针：文本 8 维 / 图片 4 维时明确 dim-mismatch，不写缓存', async () => {
    const fetchImpl = async (url, init) => {
        const body = JSON.parse(init.body);
        const isText = body.inputs[0].content[0].type === 'text';
        const vector = isText ? [1, 0, 0, 0, 0, 0, 0, 0] : [1, 0, 0, 0];
        return {
            ok: true,
            status: 200,
            text: async () => JSON.stringify({ data: [{ embedding: vector, index: 0 }], model: 'm' }),
        };
    };
    const mm = makeGallery({
        fetchImpl,
        root: join(root, 'gallery-mm-mismatch'),
        config: { imageEmbedding: { provider: 'voyage', baseURL: 'https://api.voyageai.com/v1', model: 'voyage-multimodal-3.5' } },
    });
    const probed = await mm.probeImage({ allowUpload: true });
    assert.equal(probed.ok, false);
    assert.equal(probed.code, 'dim-mismatch');

    await mm.addBytes(pngOf(3, 3, [31, 32, 33]), { tags: ['日落'] });
    const vectors = await mm.ensureVectors({ image: true, allowUpload: true });
    assert.equal(vectors.image.code, 'not-probed');
    const searched = await mm.search({ query: '日落', allowUpload: true });
    assert.equal(searched.signals.imageText.used, false);
    assert.match(searched.signals.imageText.reason, /探针/);
});

await check('图文对齐全链路（假 Voyage）：图片向量化后文字搜图可用，且与文本向量分目录', async () => {
    const { fetchImpl, calls } = voyageFetch();
    const mmRoot = join(root, 'gallery-mm');
    const mm = makeGallery({
        fetchImpl,
        root: mmRoot,
        config: {
            textEmbedding: { mode: 'local' },
            imageEmbedding: { provider: 'voyage', baseURL: 'https://api.voyageai.com/v1', model: 'voyage-multimodal-3.5' },
        },
    });
    const sunny = await mm.addBytes(pngOf(7, 5, [240, 180, 100]), { tags: ['日落'] });
    const night = await mm.addBytes(pngOf(8, 5, [20, 30, 60]), { tags: ['夜景'] });
    await mm.setAesthetic(sunny.item.id, { subject: '海边日落', tags: ['日落'] }, { channel: 'test', provider: 'fake', model: 'v1', promptVersion: 'v1' });
    await mm.setAesthetic(night.item.id, { subject: '城市夜景', tags: ['夜景'] }, { channel: 'test', provider: 'fake', model: 'v1', promptVersion: 'v1' });

    assert.equal((await mm.probeImage({ allowUpload: true })).ok, true);
    const vectors = await mm.ensureVectors({ text: true, image: true, allowUpload: true });
    assert.equal(vectors.ok, true);
    assert.ok(vectors.image.written >= 2);
    assert.equal(vectors.image.dim, 8);

    const imageDirs = await readdir(join(mmRoot, 'vectors', 'image'));
    const textDirs = await readdir(join(mmRoot, 'vectors', 'text'));
    assert.ok(imageDirs.length > 0 && textDirs.length > 0);
    assert.notEqual(imageDirs[0], textDirs[0]);
    const key = imageVectorKey(mm.getConfig().imageEmbedding, 8);
    assert.ok((await readdir(join(mmRoot, 'vectors', 'image', key))).includes(sunny.item.id + '.json'));

    const beforeText = calls.length;
    const unconfirmedText = await mm.search({ query: '日落' });
    assert.equal(unconfirmedText.signals.imageText.used, false);
    assert.match(unconfirmedText.signals.imageText.reason, /allowUpload/);
    assert.equal(calls.length, beforeText, '文字搜索未确认时不能向多模态端点发查询词');

    const result = await mm.search({ query: '日落', allowUpload: true });
    assert.equal(result.signals.imageText.used, true);
    assert.equal(result.signals.imageText.covered, 2);
    assert.equal(result.items[0].id, sunny.item.id, '文字查询应召回图文对齐空间里更近的图');
    assert.ok(result.items[0].parts.imageText > 0);
    assert.equal(result.items.some((item) => item.id === night.item.id), false,
        '图文余弦为 0 是「没有相似信号」，不能再当命中列出来');
    assert.ok(result.signals.imageText.dim === 8);

    const before = calls.length;
    const unconfirmed = await mm.search({ image: sunny.item.id });
    assert.equal(unconfirmed.signals.imageImage.used, false);
    assert.match(unconfirmed.signals.imageImage.reason, /allowUpload/);
    assert.equal(calls.length, before, '以图搜图未确认时不能发查询图');
    const confirmed = await mm.search({ image: sunny.item.id, allowUpload: true });
    assert.equal(confirmed.signals.imageImage.used, true);
    assert.equal(confirmed.items.some((item) => item.id === sunny.item.id), false, '查询图本身不计入结果');
});

await check('图片向量：超过 Voyage 像素上限时用缩略图降采样；没有缩略图则明确跳过', async () => {
    const config = { imageEmbedding: { provider: 'voyage', baseURL: 'https://api.voyageai.com/v1', model: 'voyage-multimodal-3.5' } };
    const { fetchImpl, calls } = voyageFetch();
    const downscaled = makeGallery({ fetchImpl, root: join(root, 'gallery-mm-big'), config });
    const added = await downscaled.addBytes(headerOnlyPng(5000, 4000));
    assert.equal(added.ok, true);
    assert.equal((await downscaled.probeImage({ allowUpload: true })).ok, true);
    const before = calls.length;
    const vectors = await downscaled.ensureVectors({ text: false, image: true, allowUpload: true });
    assert.equal(vectors.image.ok, true);
    assert.equal(vectors.image.written, 1);
    const imageCalls = calls.slice(before).filter((call) => call.body.inputs?.[0]?.content?.[0]?.type === 'image_base64');
    assert.equal(imageCalls.length, 1);
    assert.match(imageCalls[0].body.inputs[0].content[0].image_base64, /^data:image\/jpeg;base64,/, '超像素图应改用 JPEG 缩略图');

    const noThumb = makeGallery({ fetchImpl, root: join(root, 'gallery-mm-big-nothumb'), thumbnailer: null, config });
    assert.equal((await noThumb.addBytes(headerOnlyPng(5000, 4000))).ok, true);
    assert.equal((await noThumb.probeImage({ allowUpload: true })).ok, true);
    const skipped = await noThumb.ensureVectors({ text: false, image: true, allowUpload: true });
    assert.equal(skipped.image.written, 0);
    assert.equal(skipped.image.skipped, 1);
    assert.match(skipped.image.error, /上限/);
});

await check('vision：未配置 / 未确认都不联网；确认后写审美描述并丢弃作者字段', async () => {
    const off = await gallery.describeImage('a'.repeat(64), { allowUpload: true });
    assert.equal(off.code, 'not-found');

    const item = (await gallery.list({ limit: 1 })).items[0];
    const refused = await gallery.describeImage(item.id, {});
    assert.equal(refused.code, 'needs-confirmation');

    const unconfigured = makeGallery({ config: {} });
    const unconfiguredItem = (await unconfigured.list({ limit: 1 })).items[0];
    const noModel = await unconfigured.describeImage(unconfiguredItem.id, { allowUpload: true });
    assert.equal(noModel.code, 'unconfigured');
    assert.match(noModel.error, /没有配置视觉模型/);

    const calls = [];
    const fetchImpl = async (url, init) => {
        calls.push({ url, body: JSON.parse(init.body) });
        return {
            ok: true,
            status: 200,
            text: async () => JSON.stringify({
                choices: [{
                    message: {
                        content: JSON.stringify({
                            subject: '海边日落',
                            composition: '三分构图',
                            palette: ['暖橙', '淡紫'],
                            light: '逆光',
                            material: '海水',
                            mood: '安静',
                            summary: '黄昏的海面',
                            tags: ['日落', '海边'],
                            author: '张三',
                            copyright: '某公司',
                        }),
                    },
                }],
            }),
        };
    };
    const vision = makeGallery({
        fetchImpl,
        root: join(root, 'gallery-vision'),
        config: { vision: { channel: 'openai', baseURL: 'https://vision.example.com/v1', model: 'fake-vl', apiKey: 'k' } },
    });
    const added = await vision.addBytes(pngOf(4, 4, [11, 22, 33]));
    const described = await vision.describeImage(added.item.id, { allowUpload: true });
    assert.equal(described.ok, true);
    assert.equal(described.item.aesthetic.subject, '海边日落');
    assert.deepEqual(described.item.aesthetic.droppedFields, ['author', 'copyright']);
    assert.equal(described.item.aesthetic.author, undefined);
    assert.equal(calls.length, 1);
    assert.equal(calls[0].url, 'https://vision.example.com/v1/chat/completions');
    assert.equal(calls[0].body.model, 'fake-vl');
    assert.match(calls[0].body.messages[0].content[1].image_url.url, /^data:image\/png;base64,/);
    const stored = await vision.get(added.item.id);
    assert.equal(stored.aesthetic.model, 'fake-vl');
});

await check('vision：提示词禁止推断作者/版权；模型输出不是 JSON 时不写描述', async () => {
    const fetchImpl = async () => ({
        ok: true,
        status: 200,
        text: async () => JSON.stringify({ choices: [{ message: { content: '这张图很好看，作者是张三。' } }] }),
    });
    const vision = makeGallery({
        fetchImpl,
        root: join(root, 'gallery-vision-bad'),
        config: { vision: { channel: 'openai', baseURL: 'https://vision.example.com/v1', model: 'fake-vl' } },
    });
    const added = await vision.addBytes(pngOf(3, 3, [5, 6, 7]));
    const described = await vision.describeImage(added.item.id, { allowUpload: true });
    assert.equal(described.ok, false);
    assert.equal(described.code, 'bad-output');
    assert.equal((await vision.get(added.item.id)).aesthetic, null, '不可解析的输出不能写进描述');
});

await check('vision 宿主通道：inputModalities 显式不含 image 时不发请求', async () => {
    let streamed = 0;
    const host = {
        attachments: { saveImage: async () => ({ id: 'ref-1' }) },
        llm: {
            prepareCall: async () => ({
                inputModalities: ['text'],
                config: {},
                stream: async function* stream() {
                    streamed += 1;
                    yield { type: 'text-delta', text: '{"subject":"x"}' };
                },
            }),
        },
    };
    const vision = makeGallery({
        root: join(root, 'gallery-vision-host'),
        config: { vision: { channel: 'host', provider: 'p', model: 'text-only' } },
    });
    const added = await vision.addBytes(pngOf(3, 3, [8, 9, 10]));
    const described = await vision.describeImage(added.item.id, { allowUpload: true, host });
    assert.equal(described.ok, false);
    assert.equal(described.code, 'no-vision');
    assert.equal(streamed, 0, '模型声明不支持图片时绝不能发请求');
});

await check('vision 宿主通道：inputModalities 缺失时照发，但标注能力未经声明确认', async () => {
    const seen = [];
    const host = {
        attachments: { saveImage: async (payload) => { seen.push(payload.mediaType); return { id: 'ref-2' }; } },
        llm: {
            prepareCall: async () => ({
                config: { temperature: 0 },
                stream: async function* stream(options) {
                    seen.push(options.messages[0].content[0].text);
                    seen.push(options.messages[0].content[1].type);
                    yield { type: 'text-delta', text: '{"subject":"海边","tags":["日落"]}' };
                },
            }),
        },
    };
    const vision = makeGallery({
        root: join(root, 'gallery-vision-host2'),
        config: { vision: { channel: 'host', provider: 'p', model: 'unknown-modalities' } },
    });
    const added = await vision.addBytes(pngOf(3, 3, [12, 13, 14]));
    const described = await vision.describeImage(added.item.id, { allowUpload: true, host });
    assert.equal(described.ok, true);
    assert.equal(described.item.aesthetic.subject, '海边');
    assert.match(described.note, /inputModalities/);
    assert.equal(seen[0], 'image/png');
    assert.equal(seen[2], 'image');
});

await check('remove：删条目、原图、缩略图与向量', async () => {
    const added = await gallery.addBytes(pngOf(9, 9, [77, 88, 99]));
    const id = added.item.id;
    await gallery.ensureVectors({ text: true, ids: [id] });
    const removed = await gallery.remove(id);
    assert.equal(removed.ok, true);
    assert.equal((await gallery.get(id)), null);
    await assert.rejects(stat(join(galleryRoot, 'originals', id + '.png')));
    await assert.rejects(stat(join(galleryRoot, 'thumbs', id + '.jpg')));
});

await check('status：如实报告能力，未配置就是未配置', async () => {
    const status = await gallery.status();
    assert.equal(status.ok, true);
    assert.equal(status.imageEmbedding.configured, false);
    assert.match(status.imageEmbedding.note, /没有配置图文对齐/);
    assert.equal(status.vision.configured, false);
    assert.match(status.textEmbedding.note, /本地哈希向量/);
    assert.equal(status.uploadPolicy.includes('allowUpload'), true);
});


await check('文件夹：入库补祖先、按目录过滤、status 计数', async () => {
    const lib = makeGallery({ root: join(root, 'gallery-folders') });
    const a = await lib.addBytes(pngOf(3, 3, [1, 2, 3]), { folder: ' 人物 / 街头 ' });
    assert.equal(a.ok, true, a.error);
    assert.equal(a.item.folder, '人物/街头', '路径要规范化后再存');
    const b = await lib.addBytes(pngOf(4, 4, [4, 5, 6]), { folder: '人物' });
    assert.equal(b.item.folder, '人物');
    const c = await lib.addBytes(pngOf(3, 3, [7, 8, 9]), { folder: '人物/街头' });
    assert.equal(c.ok, true);

    const direct = await lib.list({ folder: '人物/街头' });
    assert.equal(direct.total, 2, '文件夹过滤只算直属图片');
    const parent = await lib.list({ folder: '人物' });
    assert.equal(parent.total, 1);
    assert.equal((await lib.list({ folder: '人物/../街头' })).total, 0, '非法过滤条件不能变成全部');

    const status = await lib.status();
    const byPath = Object.fromEntries(status.folders.map((folder) => [folder.path, folder]));
    assert.equal(byPath['人物'].direct, 1);
    assert.equal(byPath['人物'].total, 3, '父文件夹统计要含后代');
    assert.equal(byPath['人物/街头'].direct, 2);
    assert.equal(status.counts.withFolder, 3);
    assert.equal(status.counts.folders, 2);
});

await check('文件夹：规范化拒绝 . / ..、超长与过深路径', async () => {
    assert.equal(normalizeFolderPath(' A \\ B '), 'A/B');
    assert.equal(normalizeFolderPath('/甲//乙/'), '甲/乙');
    assert.equal(normalizeFolderPath(''), '');
    assert.throws(() => normalizeFolderPath('人物/../街头'), /\.\./);
    assert.throws(() => normalizeFolderPath('x'.repeat(41)), /太长/);
    assert.throws(() => normalizeFolderPath(Array.from({ length: 9 }, (_, index) => 'd' + index).join('/')), /层级/);
    const lib = makeGallery({ root: join(root, 'gallery-folders-bad') });
    const added = await lib.addBytes(pngOf(3, 3, [21, 22, 23]), { folder: '../外面' });
    assert.equal(added.ok, false);
    assert.equal(added.code, 'bad-folder');
});

await check('文件夹：显式空目录、重命名（含后代与图片）、删除只允许空目录', async () => {
    const lib = makeGallery({ root: join(root, 'gallery-folders-manage') });
    const created = await lib.createFolder('风景/夜晚');
    assert.equal(created.ok, true);
    assert.equal(created.created, true);
    const again = await lib.createFolder('风景/夜晚');
    assert.equal(again.ok, true);
    assert.equal(again.created, false, '重复创建是幂等的');

    const added = await lib.addBytes(pngOf(6, 6, [31, 32, 33]), { folder: '风景/夜晚' });
    const moved = await lib.updateManual(added.item.id, { folder: '风景' });
    assert.equal(moved.ok, true);
    assert.equal(moved.item.folder, '风景');

    const renamed = await lib.renameFolder('风景', '风景档案');
    assert.equal(renamed.ok, true, renamed.error);
    assert.equal(renamed.moved, 1);
    const item = await lib.get(added.item.id);
    assert.equal(item.manual.folder, '风景档案');
    assert.equal((await lib.list({ folder: '风景' })).total, 0);
    assert.equal((await lib.list({ folder: '风景档案' })).total, 1);

    const cycle = await lib.renameFolder('风景档案', '风景档案/子目录');
    assert.equal(cycle.ok, false);
    assert.equal(cycle.code, 'bad-folder');

    const blocked = await lib.removeFolder('风景档案');
    assert.equal(blocked.ok, false);
    assert.equal(blocked.code, 'folder-not-empty');
    assert.match(blocked.error, /1 张/);

    await lib.updateManual(added.item.id, { folder: '' });
    await lib.createFolder('空目录/子目录');
    const removed = await lib.removeFolder('空目录');
    assert.equal(removed.ok, true);
    assert.equal(removed.removedFolders, 2, '空父目录连同空子目录一起删');
    const status = await lib.status();
    assert.equal(status.folders.some((folder) => folder.path.startsWith('空目录')), false);

    const bad = await lib.updateManual(added.item.id, { folder: '...' });
    assert.equal(bad.ok, false);
    assert.equal(bad.code, 'bad-folder');
    assert.equal((await lib.get(added.item.id)).manual.folder, '');
});

await check('createPillowThumbnailer：Pillow 不可用时返回 ok:false 与原因', async () => {
    const thumbnail = createPillowThumbnailer({ python: 'definitely-not-a-python-binary' });
    const output = await thumbnail({ sourcePath: 'a.png', targetPath: join(root, 'x.jpg') });
    assert.equal(output.ok, false);
    assert.match(output.reason, /Pillow 缩略图失败/);
});

await rm(root, { recursive: true, force: true });
await rm(outsideDir, { recursive: true, force: true });

const failed = results.filter((result) => !result.ok);
for (const result of results) {
    console.log((result.ok ? 'ok   ' : 'FAIL ') + result.name);
    if (!result.ok) console.log('     ' + (result.error?.stack ?? result.error));
}
console.log('');
console.log(results.length - failed.length + '/' + results.length + ' passed');
if (failed.length > 0) process.exitCode = 1;
