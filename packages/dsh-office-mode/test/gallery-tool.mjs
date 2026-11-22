/**
 * 图库统一入口（office_memory 的 gallery action）的测试，全部离线。
 *
 * 跑法：node test/gallery-tool.mjs
 */
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { deflateSync } from 'node:zlib';
import { createGallery } from '../src/gallery.js';
import { executeGalleryAction, renderGalleryValue } from '../src/gallery-tool.js';

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

function pngOf(width, height, rgb = [80, 130, 160]) {
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

const root = await mkdtemp(join(tmpdir(), 'dsh-gallery-tool-'));
const work = join(root, 'work');
await mkdir(work, { recursive: true });
const imagePath = join(work, '海边.png');
await writeFile(imagePath, pngOf(6, 4));
const gallery = createGallery({
    root: join(root, 'gallery'),
    cwd: work,
    roots: [work],
    thumbnailer: async ({ targetPath }) => {
        await writeFile(targetPath, Buffer.from([0xff, 0xd8, 0xff, 0xe0]));
        return { ok: true, mediaType: 'image/jpeg' };
    },
    config: { textEmbedding: { mode: 'local' } },
    fetchImpl: async (url, init) => {
        const body = JSON.parse(init.body);
        if (String(url).includes('/chat/completions')) {
            assert.equal(body.model, 'fake-vision');
            const content = body.messages[0].content;
            assert.equal(content[0].type, 'text');
            assert.equal(content[1].type, 'image_url');
            assert.match(content[1].image_url.url, /^data:image\/png;base64,/);
            const payload = {
                subject: '海面与礁石',
                composition: '低地平线，三分法',
                palette: ['#4a7fa5', '#d9d2c5'],
                light: '清晨侧逆光',
                material: '水与岩石',
                mood: '安静、开阔',
                summary: '清晨的海边，冷色水面与暖色礁石对照。',
                tags: ['海边', '清晨', '冷色', '风景'],
                author: '某摄影师',
                copyright: '© 2026',
            };
            return {
                ok: true,
                status: 200,
                text: async () => JSON.stringify({ choices: [{ message: { content: JSON.stringify(payload) } }] }),
            };
        }
        throw new Error('不该有别的请求：' + url);
    },
});
gallery.configure({ vision: { channel: 'openai', baseURL: 'https://vision.invalid/v1', model: 'fake-vision', apiKey: 'k' } });
const deps = { galleryFor: () => gallery, hostOf: () => null };

let addedId = '';
await check('status：没有配置时如实报告未配置', async () => {
    const value = await executeGalleryAction({ op: 'status' }, deps);
    assert.equal(value.ok, true);
    const text = renderGalleryValue(value);
    assert.match(text, /共 0 张/);
    assert.match(text, /图文对齐：未配置/);
});

await check('add：路径导入 + 文本里有路径/标签，重复导入不重复保存', async () => {
    const value = await executeGalleryAction({ op: 'add', path: imagePath, name: '参考图', description: '封面参考', tags: ['参考'], category: '风景' }, deps);
    assert.equal(value.ok, true);
    addedId = value.item.id;
    assert.equal(value.item.name, '参考图');
    const text = renderGalleryValue(value);
    assert.match(text, /已入库/);
    assert.match(text, /参考图/, '工具回执优先用名称做标题');
    assert.match(text, /originals/);
    assert.match(text, /参考/);

    const again = await executeGalleryAction({ op: 'add', path: imagePath }, deps);
    assert.equal(again.duplicate, true);
    assert.match(renderGalleryValue(again), /已经在图库里/);
});

await check('list / search：返回图片路径，标签命中并给分项分', async () => {
    const listed = await executeGalleryAction({ op: 'list' }, deps);
    assert.equal(listed.items.length, 1);
    assert.match(renderGalleryValue(listed), /最近 1 张/);

    const found = await executeGalleryAction({ op: 'search', query: '参考' }, deps);
    assert.equal(found.ok, true);
    assert.equal(found.items[0].id, addedId);
    const text = renderGalleryValue(found);
    assert.match(text, /找到 1 张/);
    assert.match(text, /分数：/);
    assert.match(text, /标签/);
});

await check('tag：手工改标签/分类/收藏', async () => {
    const value = await executeGalleryAction({ op: 'tag', id: addedId, name: '改过的名字', description: '工具改写描述', addTags: ['日落'], favorite: true }, deps);
    assert.equal(value.ok, true);
    assert.equal(value.item.name, '改过的名字');
    assert.equal(value.item.description, '工具改写描述');
    assert.deepEqual(value.item.manualTags, ['参考', '日落']);
    assert.match(renderGalleryValue(value), /收藏/);
});

await check('describe：未确认上传时拒绝，并提示 allowUpload', async () => {
    const value = await executeGalleryAction({ op: 'describe', id: addedId }, deps);
    assert.equal(value.ok, false);
    assert.equal(value.code, 'needs-confirmation');
    assert.match(renderGalleryValue(value), /allowUpload: true/);
});

await check('describe：OpenAI 兼容视觉端点写入审美描述，作者/版权字段被丢弃', async () => {
    const value = await executeGalleryAction({ op: 'describe', id: addedId, allowUpload: true }, deps);
    assert.equal(value.ok, true);
    const aesthetic = value.item.aesthetic;
    assert.equal(aesthetic.subject, '海面与礁石');
    assert.equal(aesthetic.mood, '安静、开阔');
    assert.deepEqual(aesthetic.droppedFields.sort(), ['author', 'copyright']);
    const text = renderGalleryValue(value);
    assert.match(text, /主体|subject/);
    assert.match(text, /丢弃/);
});

await check('index：文本向量写入本地缓存，图片向量未确认时如实失败', async () => {
    const value = await executeGalleryAction({ op: 'index', text: true, image: true }, deps);
    assert.equal(value.text.ok, true);
    assert.ok(value.text.written >= 1);
    assert.equal(value.image.ok, false);
    assert.equal(value.image.code, 'unconfigured');
    const text = renderGalleryValue(value);
    assert.match(text, /文本：写入 1 条/);
    assert.match(text, /图片：没有完成/);
});

await check('folder：创建、归类、过滤、重命名与空目录删除', async () => {
    const created = await executeGalleryAction({ op: 'folder-create', path: '参考/风景' }, deps);
    assert.equal(created.ok, true, created.error);
    assert.match(renderGalleryValue(created), /已创建文件夹：参考\/风景/);

    const tagged = await executeGalleryAction({ op: 'tag', id: addedId, folder: '参考/风景' }, deps);
    assert.equal(tagged.ok, true);
    assert.equal(tagged.item.folder, '参考/风景');
    assert.match(renderGalleryValue(tagged), /文件夹：参考\/风景/);

    const listed = await executeGalleryAction({ op: 'folder-list' }, deps);
    assert.equal(listed.ok, true);
    assert.equal(listed.folders.find((folder) => folder.path === '参考/风景').direct, 1);
    assert.match(renderGalleryValue(listed), /参考\/风景/);

    const filtered = await executeGalleryAction({ op: 'list', folder: '参考/风景' }, deps);
    assert.equal(filtered.items.length, 1);

    const renamed = await executeGalleryAction({ op: 'folder-rename', from: '参考/风景', to: '参考/夜景' }, deps);
    assert.equal(renamed.ok, true, renamed.error);
    assert.equal(renamed.moved, 1);
    assert.equal((await executeGalleryAction({ op: 'list', folder: '参考/夜景' }, deps)).items.length, 1);

    const blocked = await executeGalleryAction({ op: 'folder-remove', path: '参考/夜景' }, deps);
    assert.equal(blocked.ok, false);
    assert.equal(blocked.code, 'folder-not-empty');

    await executeGalleryAction({ op: 'tag', id: addedId, folder: '' }, deps);
    const removed = await executeGalleryAction({ op: 'folder-remove', path: '参考/夜景' }, deps);
    assert.equal(removed.ok, true);
    assert.match(renderGalleryValue(removed), /已删除文件夹/);
});

await check('remove：删除后可检索不到，文本明确', async () => {
    const value = await executeGalleryAction({ op: 'remove', id: addedId }, deps);
    assert.equal(value.ok, true);
    assert.match(renderGalleryValue(value), /已删除|移除|删/);
    const listed = await executeGalleryAction({ op: 'list' }, deps);
    assert.equal(listed.items.length, 0);
});

await check('未知 op / 没有图库：都是明确错误而不是抛异常', async () => {
    const bad = await executeGalleryAction({ op: 'nope' }, deps);
    assert.equal(bad.ok, false);
    assert.match(bad.text, /list \/ search \/ add/);

    const missing = await executeGalleryAction({ op: 'list' }, { galleryFor: () => null });
    assert.equal(missing.ok, false);
    assert.match(renderGalleryValue(missing), /没有初始化/);
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
