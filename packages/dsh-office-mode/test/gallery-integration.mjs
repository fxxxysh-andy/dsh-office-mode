/**
 * 图库与记忆/工具/HTTP 的集成测试（不是独立模块单测）。
 *
 * 覆盖三件事：
 *   1. resolveConfig 把 gallery 组收口成运行期配置，memory.galleryDir 生效；
 *   2. office_memory 的 gallery 动作真的接到底层（add / tag / search / status）；
 *   3. view.js 注册的两条真实端点：图库 JSON 与 KB 图片资产（安全头 / 路径边界）。
 *
 * 跑法：node test/gallery-integration.mjs
 */
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { resolveConfig } from '../src/config.js';
import { buildTools } from '../src/tools.js';
import { noteWorkspace, registerMemoryView, KB_ASSET_PATH } from '../src/view.js';
import { GALLERY_PATH } from '../src/gallery-routes.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lzq1WQAAAABJRU5ErkJggg==', 'base64');
const results = [];
async function check(name, fn) {
    try { await fn(); results.push({ name, ok: true }); }
    catch (error) { results.push({ name, ok: false, error }); }
}

function fakeRes() {
    return {
        status: 0,
        headers: {},
        body: '',
        writeHead(status, headers) { this.status = status; this.headers = headers ?? {}; },
        end(body) { if (body !== undefined) this.body += Buffer.isBuffer(body) ? body.toString('utf8') : String(body); },
    };
}

function fakeCtx(webServer) {
    return {
        get(name) { return name === 'webServer' ? webServer : undefined; },
    };
}

await check('resolveConfig：gallery 默认 off/none，显式配置被收口，galleryDir 保留', () => {
    const empty = resolveConfig({});
    assert.equal(empty.gallery.vision.channel, 'off');
    assert.equal(empty.gallery.imageEmbedding.provider, 'none');
    assert.equal(empty.gallery.textEmbedding.model, '', '文本向量缺省沿用知识库 embedding');
    assert.equal(empty.memory.galleryDir, '');

    const cfg = resolveConfig({
        memory: { galleryDir: 'D:/tmp/pics' },
        gallery: {
            vision: { channel: 'openai', baseURL: 'https://vision.example/v1', model: 'v1', apiKey: 'k' },
            imageEmbedding: { provider: 'cohere', model: 'embed-v4.0', dimensions: 1024 },
        },
    });
    assert.equal(cfg.memory.galleryDir, 'D:/tmp/pics');
    assert.equal(cfg.gallery.vision.channel, 'openai');
    assert.equal(cfg.gallery.vision.baseURL, 'https://vision.example/v1');
    assert.equal(cfg.gallery.imageEmbedding.provider, 'cohere');
    assert.equal(cfg.gallery.imageEmbedding.dimensions, 1024);
});

await check('office_memory gallery：add → tag → search → status 全链路，未配置就说未配置', async () => {
    const base = await mkdtemp(join(tmpdir(), 'gi-'));
    try {
        const memory = { dir: join(base, 'memory'), kbDir: join(base, 'kb'), galleryDir: join(base, 'pics') };
        const config = resolveConfig({ memory, gallery: {} });
        const tools = buildTools(config, null);
        const memoryTool = tools.find((tool) => tool.name === 'office_memory');
        assert.ok(memoryTool !== undefined, 'office_memory 必须注册');
        const exec = { agent: { session: { header: { cwd: base } } }, signal: new AbortController().signal };
        const image = join(base, 'shot.png');
        await writeFile(image, PNG);

        const added = await memoryTool.execute({ action: 'gallery', op: 'add', path: image, tags: ['建筑', '暖色'] }, exec);
        assert.equal(added.ok, true, added.text);
        assert.match(added.text, /已入库/);
        const list = await memoryTool.execute({ action: 'gallery', op: 'list' }, exec);
        assert.equal(list.value.items.length, 1);
        const id = list.value.items[0].id;
        assert.match(id, /^[0-9a-f]{64}$/);

        const tagged = await memoryTool.execute({ action: 'gallery', op: 'tag', id, tags: ['建筑', '夜景'], category: '参考', favorite: true }, exec);
        assert.equal(tagged.ok, true, tagged.text);
        assert.equal(tagged.value.item.favorite, true);
        assert.equal(tagged.value.item.category, '参考');
        assert.deepEqual(tagged.value.item.tags, ['建筑', '夜景']);

        const found = await memoryTool.execute({ action: 'gallery', op: 'search', query: '夜景' }, exec);
        assert.equal(found.ok !== false, true);
        assert.equal(found.value.items.length, 1);
        assert.equal(found.value.items[0].id, id);

        const status = await memoryTool.execute({ action: 'gallery', op: 'status' }, exec);
        assert.match(status.text, /图文对齐：未配置/);
        assert.equal(status.value.status.imageEmbedding.configured, false);

        const describe = await memoryTool.execute({ action: 'gallery', op: 'describe', id }, exec);
        assert.equal(describe.ok, false);
        assert.match(describe.text, /allowUpload/);
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('view 路由：图库状态可读，KB 资产按真实类型服务且带安全头', async () => {
    const base = await mkdtemp(join(tmpdir(), 'gi-view-'));
    try {
        const kbDir = join(base, 'kb');
        const galleryDir = join(base, 'pics');
        await mkdir(join(kbDir, 'assets'), { recursive: true });
        await writeFile(join(kbDir, 'assets', 'a.png'), PNG);
        await writeFile(join(kbDir, 'assets', 'evil.svg'), '<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>');
        const memory = { enabled: true, dir: join(base, 'memory'), kbDir, galleryDir };
        const config = resolveConfig({ memory, gallery: {} });
        noteWorkspace(base);

        const routes = new Map();
        const webServer = {
            register({ path, handler }) {
                routes.set(path, handler);
                return () => routes.delete(path);
            },
        };
        const dispose = registerMemoryView(fakeCtx(webServer), { getMemory: () => memory, getConfig: () => config });
        assert.equal(typeof dispose, 'function');
        assert.equal(typeof routes.get(GALLERY_PATH), 'function');
        assert.equal(typeof routes.get(KB_ASSET_PATH), 'function');

        const galleryRes = fakeRes();
        await routes.get(GALLERY_PATH)({ method: 'GET', url: GALLERY_PATH + '?view=status&cwd=' + encodeURIComponent(base) }, galleryRes);
        assert.equal(galleryRes.status, 200, galleryRes.body);
        const status = JSON.parse(galleryRes.body);
        assert.equal(status.imageEmbedding.configured, false);
        assert.equal(status.root, resolve(galleryDir));

        const imageRes = fakeRes();
        await routes.get(KB_ASSET_PATH)({ method: 'GET', url: KB_ASSET_PATH + '?cwd=' + encodeURIComponent(base) + '&path=' + encodeURIComponent('assets/a.png') }, imageRes);
        assert.equal(imageRes.status, 200, imageRes.body);
        assert.equal(imageRes.headers['content-type'], 'image/png');
        assert.equal(imageRes.headers['x-content-type-options'], 'nosniff');
        assert.match(imageRes.headers['content-security-policy'], /default-src 'none'/);
        assert.match(imageRes.headers['cache-control'], /no-cache/);

        const headRes = fakeRes();
        await routes.get(KB_ASSET_PATH)({ method: 'HEAD', url: KB_ASSET_PATH + '?cwd=' + encodeURIComponent(base) + '&path=' + encodeURIComponent('assets/a.png') }, headRes);
        assert.equal(headRes.status, 200);
        assert.equal(headRes.body, '', 'HEAD 不回正文');

        const svgRes = fakeRes();
        await routes.get(KB_ASSET_PATH)({ method: 'GET', url: KB_ASSET_PATH + '?cwd=' + encodeURIComponent(base) + '&path=' + encodeURIComponent('assets/evil.svg') }, svgRes);
        assert.equal(svgRes.status, 415, svgRes.body);

        const escapeRes = fakeRes();
        await routes.get(KB_ASSET_PATH)({ method: 'GET', url: KB_ASSET_PATH + '?cwd=' + encodeURIComponent(base) + '&path=' + encodeURIComponent('assets/../secret.png') }, escapeRes);
        assert.equal(escapeRes.status, 400, escapeRes.body);
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

for (const r of results) {
    if (r.ok) console.log('ok   ' + r.name);
    else { console.log('FAIL ' + r.name); console.log(r.error?.stack ?? r.error); }
}
const failed = results.filter((r) => !r.ok).length;
console.log('');
console.log((results.length - failed) + '/' + results.length + ' passed');
if (failed > 0) process.exitCode = 1;
