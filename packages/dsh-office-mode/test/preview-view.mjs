/**
 * 预览面板四条路由的测试：清单 / 大纲 / 单页图 / 文本原文。
 *
 * 重点四件事（与 test/view.mjs 同一套纪律）：
 *   1. **只给 cwd 里的东西**：未知 cwd、`../` 穿越、绝对 path 一律拒绝；
 *   2. **只读**：非 GET/HEAD 405，且没有任何写工作目录的路径；
 *   3. **错误可执行**：dpi / 页码越界与引擎不可用都带「下一步」；
 *   4. **文本通道有界**：只读前 2 MiB、按 UTF-8 字符边界截断、二进制扩展名 415。
 *
 * 渲染一律用假引擎替身（不启动 LibreOffice）：CI 机器上装没装引擎不该决定
 * 这组测试红不红；真引擎的像素正确性不属于端点契约。
 *
 * 跑法：node test/preview-view.mjs
 */
import { mkdir, rm, writeFile } from 'node:fs/promises';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { displayPath } from '../src/engine/kit.js';
import { PREVIEW_FILE_LIMIT, PREVIEW_VIEW_PATHS, TEXT_READ_LIMIT, TEXT_VIEW_PATH, registerPreviewView } from '../src/preview-view.js';
import { noteWorkspace, resetWorkspaces } from '../src/view.js';

const TMP = fileURLToPath(new URL('../../../.office/runtime/tmp/preview-view/', import.meta.url));

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

let caseId = 0;
async function freshRoot() {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    await rm(root, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    return root;
}

/** 1×1 透明 PNG：端点只要验证「base64 解出来是 PNG」，不需要真画面。 */
const PNG_1PX = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
);

function makeRes() {
    return {
        status: 0,
        headers: null,
        body: '',
        writeHead(status, headers) {
            this.status = status;
            this.headers = headers;
        },
        end(body) {
            this.body = body ?? '';
        },
    };
}

async function callRoute(handler, { method = 'GET', url } = {}) {
    const res = makeRes();
    await handler({ method, url }, res);
    let json = null;
    try {
        json = JSON.parse(res.body);
    } catch {
        json = null;
    }
    return { status: res.status, headers: res.headers, json, raw: res.body };
}

function fakeCtx() {
    const routes = [];
    return {
        routes,
        get(name) {
            if (name === 'webServer') {
                return {
                    register(route) {
                        routes.push(route);
                        return () => {
                            const at = routes.indexOf(route);
                            if (at >= 0) routes.splice(at, 1);
                        };
                    },
                };
            }
            return undefined;
        },
    };
}

/** 假引擎：写一张 1×1 PNG 进缓存目录，按 previewImages 的返回形状回话。 */
function fakeRender({ pageCount = 5, reused = false, fail = null, spy = null } = {}) {
    return async (absolute, options, env, cache) => {
        spy?.calls.push({ absolute, pages: options.pages, dpi: options.dpi, maxPages: options.maxPages });
        if (fail !== null) throw fail;
        const dir = cache.ensureDir('preview');
        const file = join(dir, `page-${options.pages[0]}.png`);
        writeFileSync(file, PNG_1PX);
        return {
            ok: true,
            kind: 'images',
            dpi: options.dpi,
            pages: options.pages,
            pageCount,
            count: 1,
            reused,
            images: [{ index: 1, page: options.pages[0], path: displayPath(env.root, file), width: 1, height: 1, bytes: PNG_1PX.length }],
        };
    };
}

/** 假 pptx 解析：形状与 formats/ppt.js 的 read() 返回对齐（只给端点用到的字段）。 */
function fakeReadPpt({ pages = 5, outline = null, warnings = ['第 3 页没有标题文本'], layouts = null } = {}) {
    return async () => ({
        ok: true,
        format: 'pptx',
        theme: 'business',
        size: '16:9',
        stats: {
            slides: pages,
            layouts: layouts ?? { cover: 1, bullets: 2, shape: 1, table: 1 },
            shapes: 25,
            textBoxes: 11,
            images: 2,
            tables: 1,
            charts: 0,
            words: 180,
            notes: 2,
        },
        pages: Array.from({ length: pages }, (_, index) => ({
            index: index + 1,
            layout: 'bullets',
            title: `第 ${index + 1} 页`,
        })),
        outline: outline ?? [
            'P1 cover：季度汇报',
            'P2 bullets：本季重点｜4 个文本框',
            'P3 shape：12 个形状',
            'P4 table：收入结构｜3 行 × 4 列',
            'P5 closing：谢谢',
        ],
        warnings,
    });
}

function mount({ root, memory = {}, render, readPpt, loadEngine } = {}) {
    resetWorkspaces();
    if (typeof root === 'string') noteWorkspace(root);
    const ctx = fakeCtx();
    registerPreviewView(ctx, {
        getMemory: () => memory,
        render,
        readPpt,
        loadEngine,
    });
    const route = (path) => ctx.routes.find((item) => item.path === path)?.handler;
    const all = {
        files: route(PREVIEW_VIEW_PATHS.files),
        outline: route(PREVIEW_VIEW_PATHS.outline),
        page: route(PREVIEW_VIEW_PATHS.page),
        text: route(TEXT_VIEW_PATH),
    };
    for (const [name, handler] of Object.entries(all)) {
        assert.equal(typeof handler, 'function', `${name} 路由没注册上`);
    }
    return { ctx, ...all };
}

/** 文本端点的 URL：cwd + 相对 path。 */
function textUrl(root, path) {
    return `${TEXT_VIEW_PATH}?cwd=${encodeURIComponent(root)}&path=${encodeURIComponent(path)}`;
}

await rm(TMP, { recursive: true, force: true });

// ── 文件清单 ────────────────────────────────────────────────────────────────

await check('清单：深度 2 内的办公产物按 mtime 倒序列出，隐藏目录与依赖目录不扫', async () => {
    const root = await freshRoot();
    await mkdir(join(root, 'out'), { recursive: true });
    await mkdir(join(root, 'out', 'sub'), { recursive: true });
    await mkdir(join(root, 'out', 'sub', 'deep'), { recursive: true });
    await mkdir(join(root, '.office', 'runtime', 'cache'), { recursive: true });
    await mkdir(join(root, 'node_modules', 'pkg'), { recursive: true });
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    await writeFile(join(root, 'out', 'report.docx'), 'doc');
    await writeFile(join(root, 'out', 'sub', 'notes.md'), '# 提纲');
    await writeFile(join(root, 'out', 'sub', 'deep', 'too-deep.pdf'), 'pdf');
    await writeFile(join(root, '.office', 'runtime', 'cache', 'noise.pptx'), 'x');
    await writeFile(join(root, 'node_modules', 'pkg', 'noise.docx'), 'x');
    await writeFile(join(root, 'cover.png'), 'img');

    const server = mount({ root });
    const res = await callRoute(server.files, { url: `${PREVIEW_VIEW_PATHS.files}?cwd=${encodeURIComponent(root)}` });
    assert.equal(res.status, 200);
    assert.equal(res.json.ok, true);
    assert.equal(res.json.root, root);
    assert.deepEqual(res.json.files.map((item) => item.path).sort(), ['deck.pptx', 'out/report.docx', 'out/sub/notes.md']);
    const ppt = res.json.files.find((item) => item.path === 'deck.pptx');
    assert.equal(ppt.kind, 'pptx');
    assert.equal(ppt.ext, '.pptx');
    assert.equal(typeof ppt.bytes, 'number');
    assert.match(ppt.mtime, /^\d{4}-\d{2}-\d{2}T/);
    assert.equal(res.json.files.find((item) => item.path === 'out/sub/notes.md').kind, 'text');
    assert.equal(res.json.depth, 2);
    assert.equal(res.json.truncated, false);
    assert.doesNotMatch(res.raw, /noise/, '隐藏目录与 node_modules 不该出现在响应里');
    assert.doesNotMatch(res.raw, /cover\.png/, '不支持的扩展名不该出现');
});

await check('清单：最多给 20 条，total/truncated 如实报', async () => {
    const root = await freshRoot();
    for (let index = 0; index < 25; index += 1) {
        await writeFile(join(root, `deck-${String(index).padStart(2, '0')}.pptx`), 'x');
    }
    const server = mount({ root });
    const res = await callRoute(server.files, { url: `${PREVIEW_VIEW_PATHS.files}?cwd=${encodeURIComponent(root)}` });
    assert.equal(res.status, 200);
    assert.equal(res.json.files.length, PREVIEW_FILE_LIMIT);
    assert.equal(res.json.total, 25);
    assert.equal(res.json.truncated, true);
    assert.equal(res.json.limit, PREVIEW_FILE_LIMIT);
});

await check('清单：未知 cwd 404，报错说明下一步', async () => {
    const root = await freshRoot();
    const server = mount({ root });
    const res = await callRoute(server.files, { url: `${PREVIEW_VIEW_PATHS.files}?cwd=${encodeURIComponent(join(root, 'other'))}` });
    assert.equal(res.status, 404);
    assert.equal(res.json.ok, false);
    assert.match(res.json.error, /不在已知列表/);
    assert.match(res.json.error, /刷新/);
});

await check('清单：不给 cwd 时用本进程见过的根；一个根都没有时空态 200', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'x');
    const withRoot = mount({ root });
    const res = await callRoute(withRoot.files, { url: PREVIEW_VIEW_PATHS.files });
    assert.equal(res.status, 200);
    assert.equal(res.json.root, root);

    const empty = mount();
    const blank = await callRoute(empty.files, { url: PREVIEW_VIEW_PATHS.files });
    assert.equal(blank.status, 200, '没有任何根不是错误，是空态');
    assert.equal(blank.json.ok, false);
    assert.match(blank.json.error, /还没有见过任何工作目录/);
});

await check('清单：POST 405，HEAD 200，cache-control: no-store', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'x');
    const server = mount({ root });
    for (const method of ['POST', 'PUT', 'DELETE', 'PATCH']) {
        const res = await callRoute(server.files, { method, url: PREVIEW_VIEW_PATHS.files });
        assert.equal(res.status, 405, `${method} 应被拒绝`);
        assert.match(res.json.error, /只读/);
    }
    const head = await callRoute(server.files, { method: 'HEAD', url: PREVIEW_VIEW_PATHS.files });
    assert.equal(head.status, 200);
    assert.equal(head.headers['cache-control'], 'no-store');
});

// ── 大纲 ────────────────────────────────────────────────────────────────────

await check('大纲：pptx 出 pages/slides/outline/layouts/warnings，outline 行解析成 {page,layout,title}', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    const server = mount({ root, readPpt: fakeReadPpt() });
    const res = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=deck.pptx`,
    });
    assert.equal(res.status, 200);
    assert.equal(res.json.ok, true);
    assert.equal(res.json.format, 'pptx');
    assert.equal(res.json.pages, 5);
    assert.equal(res.json.slides.length, 5);
    assert.deepEqual(res.json.slides[0], { page: 1, layout: 'cover', title: '季度汇报' });
    assert.deepEqual(res.json.slides[1], { page: 2, layout: 'bullets', title: '本季重点' });
    // shape 行的「12 个形状」不是标题：不能把计数塞进 title。
    assert.deepEqual(res.json.slides[2], { page: 3, layout: 'shape', title: '' });
    assert.deepEqual(res.json.slides[3], { page: 4, layout: 'table', title: '收入结构' });
    assert.equal(res.json.outline.length, 5);
    assert.deepEqual(res.json.layouts[0], { layout: 'bullets', count: 2 });
    assert.deepEqual(res.json.warnings, ['第 3 页没有标题文本']);
    assert.equal(res.json.size, '16:9');
    assert.equal(res.json.theme, 'business');
    assert.equal(res.json.stats.words, 180);
    assert.equal('text' in res.json, false, 'pptx 不该标正文通道（它走逐页大纲）');
});

await check('大纲：非 pptx 给 pages:null；文本类标 text:true，渲染类仍说按页取图', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'report.docx'), 'doc');
    await writeFile(join(root, 'notes.md'), '# 提纲');
    await writeFile(join(root, 'score.csv'), 'name,score\n张三,90\n');
    const server = mount({ root });
    const docx = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=report.docx`,
    });
    assert.equal(docx.status, 200);
    assert.equal(docx.json.pages, null);
    assert.deepEqual(docx.json.slides, []);
    assert.equal(docx.json.renderable, true);
    assert.equal('text' in docx.json, false, '能取图的格式不该标正文通道');
    assert.match(docx.json.hint, /按页取图/);

    const md = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=notes.md`,
    });
    assert.equal(md.json.pages, null);
    assert.equal(md.json.renderable, false);
    assert.equal(md.json.text, true, '文本类文件要给面板一个正文通道标记');
    assert.match(md.json.hint, /按原文预览/);
    assert.match(md.json.hint, /正文/);

    const csv = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=score.csv`,
    });
    assert.equal(csv.status, 200);
    assert.equal(csv.json.renderable, false);
    assert.equal(csv.json.text, true, '.csv 也是文本类');
    assert.equal(csv.json.pages, null);
});

await check('大纲：path 穿越与绝对路径都是 404（不告诉调用方文件是否存在）', async () => {
    const root = await freshRoot();
    const outside = join(TMP, `outside-${caseId}.pptx`);
    await writeFile(outside, 'ppt');
    const server = mount({ root, readPpt: fakeReadPpt() });
    const escape = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=${encodeURIComponent('../outside.pptx')}`,
    });
    assert.equal(escape.status, 404);
    assert.match(escape.json.error, /工作目录内/);

    const absolute = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=${encodeURIComponent(outside)}`,
    });
    assert.equal(absolute.status, 404);
    assert.match(absolute.json.error, /相对 cwd/);
});

await check('大纲：缺 path 400、文件不存在 404、坏 pptx 400', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'broken.pptx'), 'not-a-zip');
    const server = mount({
        root,
        readPpt: async () => {
            throw new Error('pptx: 缺少 ppt/presentation.xml（不是有效的演示文稿）');
        },
    });
    const missing = await callRoute(server.outline, { url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}` });
    assert.equal(missing.status, 400);
    assert.match(missing.json.error, /path 不能为空/);

    const ghost = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=ghost.pptx`,
    });
    assert.equal(ghost.status, 404);
    assert.match(ghost.json.error, /刷新文件列表/);

    const broken = await callRoute(server.outline, {
        url: `${PREVIEW_VIEW_PATHS.outline}?cwd=${encodeURIComponent(root)}&path=broken.pptx`,
    });
    assert.equal(broken.status, 400);
    assert.match(broken.json.error, /读不出大纲/);
    assert.match(broken.json.error, /重新生成/);
});

// ── 单页图 ──────────────────────────────────────────────────────────────────

await check('取图：正常路径返回 data:image/png;base64 与宽高/页数/dpi', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    const spy = { calls: [] };
    const server = mount({
        root,
        readPpt: fakeReadPpt(),
        render: fakeRender({ spy }),
        loadEngine: async () => ({ module: {} }),
    });
    const res = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=deck.pptx&page=2&dpi=96`,
    });
    assert.equal(res.status, 200);
    assert.equal(res.json.ok, true);
    assert.equal(res.json.page, 2);
    assert.equal(res.json.pages, 5);
    assert.equal(res.json.dpi, 96);
    assert.equal(res.json.width, 1);
    assert.equal(res.json.height, 1);
    assert.equal(res.json.cached, false);
    assert.ok(res.json.png.startsWith('data:image/png;base64,'));
    const decoded = Buffer.from(res.json.png.slice('data:image/png;base64,'.length), 'base64');
    assert.equal(decoded[0], 0x89, 'PNG 魔数');
    assert.equal(spy.calls.length, 1);
    assert.deepEqual(spy.calls[0].pages, [2]);
    assert.equal(spy.calls[0].dpi, 96);
    assert.equal(spy.calls[0].maxPages, 1);
});

await check('取图：dpi 越界 400，错误写清范围与下一步', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    const spy = { calls: [] };
    const server = mount({ root, readPpt: fakeReadPpt(), render: fakeRender({ spy }), loadEngine: async () => ({ module: {} }) });
    for (const dpi of ['10', '301', 'abc']) {
        const res = await callRoute(server.page, {
            url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=deck.pptx&dpi=${dpi}`,
        });
        assert.equal(res.status, 400, `dpi=${dpi} 应被拒绝`);
        assert.match(res.json.error, /24–300/);
        assert.match(res.json.error, /下一步/);
    }
    assert.equal(spy.calls.length, 0, '参数不合法时不该进渲染');
});

await check('取图：page 越界 400 且写清共几页；不合法时根本不调引擎', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    const spy = { calls: [] };
    const server = mount({
        root,
        readPpt: fakeReadPpt({ pages: 5 }),
        render: fakeRender({ spy }),
        loadEngine: async () => ({ module: {} }),
    });
    const res = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=deck.pptx&page=99`,
    });
    assert.equal(res.status, 400);
    assert.match(res.json.error, /共 5 页/);
    assert.match(res.json.error, /1–5/);
    assert.equal(spy.calls.length, 0, '越界应在渲染之前拦下');

    const zero = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=deck.pptx&page=0`,
    });
    assert.equal(zero.status, 400);
    assert.match(zero.json.error, /从 1 开始的整数/);
});

await check('取图：引擎不可用 503，提示装好后刷新即可，不给出假图', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'report.docx'), 'doc');
    const spy = { calls: [] };
    const server = mount({
        root,
        render: fakeRender({ spy }),
        loadEngine: async () => ({ error: '找不到渲染引擎 @deepseek-ai/libreoffice-kit（它随 DSH 部署安装，不是要另装的插件）。' }),
    });
    const res = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=report.docx`,
    });
    assert.equal(res.status, 503);
    assert.equal(res.json.ok, false);
    assert.match(res.json.error, /渲染引擎不可用/);
    assert.match(res.json.error, /刷新/);
    assert.equal(res.json.png, undefined);
    assert.equal(spy.calls.length, 0);
});

await check('取图：不可渲染的格式 400；非 pptx 不做页数预检（页数只有引擎知道）', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'notes.md'), '# 提纲');
    await writeFile(join(root, 'report.docx'), 'doc');
    const spy = { calls: [] };
    const server = mount({ root, readPpt: fakeReadPpt(), render: fakeRender({ spy, pageCount: 3 }), loadEngine: async () => ({ module: {} }) });
    const md = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=notes.md`,
    });
    assert.equal(md.status, 400);
    assert.match(md.json.error, /不渲染成页面图/);
    assert.equal(spy.calls.length, 0);

    const docx = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=report.docx&page=2`,
    });
    assert.equal(docx.status, 200);
    assert.equal(docx.json.pages, 3, '非 pptx 的页数来自引擎返回');
    assert.equal(spy.calls.length, 1);
});

await check('取图：渲染失败 502、缓存命中回 cached:true、HEAD 200', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    const failing = mount({
        root,
        readPpt: fakeReadPpt(),
        render: fakeRender({ fail: new Error('LibreOffice 进程退出码 1') }),
        loadEngine: async () => ({ module: {} }),
    });
    const bad = await callRoute(failing.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=deck.pptx`,
    });
    assert.equal(bad.status, 502);
    assert.match(bad.json.error, /渲染第 1 页失败/);
    assert.match(bad.json.error, /退出码 1/);

    const cached = mount({
        root,
        readPpt: fakeReadPpt(),
        render: fakeRender({ reused: true }),
        loadEngine: async () => ({ module: {} }),
    });
    const hit = await callRoute(cached.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=deck.pptx`,
    });
    assert.equal(hit.status, 200);
    assert.equal(hit.json.cached, true);
    assert.equal(hit.json.dpi, 110, '默认 dpi 与 preview.js 一致');

    const head = await callRoute(cached.page, { method: 'HEAD', url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=deck.pptx` });
    assert.equal(head.status, 200);
});

await check('取图：cwd 未知 404、path 穿越 404、缺 path 400', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    const server = mount({ root, readPpt: fakeReadPpt(), render: fakeRender(), loadEngine: async () => ({ module: {} }) });
    const untrusted = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(join(root, 'nope'))}&path=deck.pptx`,
    });
    assert.equal(untrusted.status, 404);
    assert.match(untrusted.json.error, /不在已知列表/);

    const escape = await callRoute(server.page, {
        url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}&path=..%2Fdeck.pptx`,
    });
    assert.equal(escape.status, 404);
    assert.match(escape.json.error, /工作目录内/);

    const missing = await callRoute(server.page, { url: `${PREVIEW_VIEW_PATHS.page}?cwd=${encodeURIComponent(root)}` });
    assert.equal(missing.status, 400);
    assert.match(missing.json.error, /path 不能为空/);
});

// ── 文本原文 ────────────────────────────────────────────────────────────────

await check('文本：.md 200 带正文/bytes/lines，path/name/ext/kind 齐全，HEAD 同样 200', async () => {
    const root = await freshRoot();
    const content = '# 标题\n\n第一段\n';
    await writeFile(join(root, 'notes.md'), content);
    const server = mount({ root });
    const res = await callRoute(server.text, { url: textUrl(root, 'notes.md') });
    assert.equal(res.status, 200);
    assert.equal(res.json.ok, true);
    assert.equal(res.json.path, 'notes.md');
    assert.equal(res.json.name, 'notes.md');
    assert.equal(res.json.ext, '.md');
    assert.equal(res.json.kind, 'text');
    assert.equal(res.json.bytes, Buffer.byteLength(content));
    assert.equal(res.json.text, content);
    assert.equal(res.json.truncated, false);
    assert.equal(res.json.lines, 4, '按 \\n 计数：末尾换行也算一段（与 archive/tex 同一口径）');
    assert.equal(res.json.encoding, 'utf-8');
    assert.equal(res.json.hint, undefined);

    const head = await callRoute(server.text, { method: 'HEAD', url: textUrl(root, 'notes.md') });
    assert.equal(head.status, 200);
    assert.equal(head.json.text, content);
});

await check('文本：.csv/.tex/.txt 都能读；空文件、CRLF、BOM 照实返回', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'score.csv'), 'name,score\n张三,90\n');
    await writeFile(join(root, 'paper.tex'), '\\documentclass{article}\n');
    await writeFile(join(root, 'run.txt'), 'plain text');
    await writeFile(join(root, 'empty.md'), '');
    await writeFile(join(root, 'crlf.md'), 'a\r\nb\r\n');
    await writeFile(join(root, 'bom.md'), Buffer.concat([Buffer.from([0xEF, 0xBB, 0xBF]), Buffer.from('# BOM\n')]));
    const server = mount({ root });

    const csv = await callRoute(server.text, { url: textUrl(root, 'score.csv') });
    assert.equal(csv.status, 200);
    assert.equal(csv.json.ext, '.csv');
    assert.equal(csv.json.text, 'name,score\n张三,90\n');
    assert.equal(csv.json.lines, 3);

    const tex = await callRoute(server.text, { url: textUrl(root, 'paper.tex') });
    assert.equal(tex.status, 200);
    assert.equal(tex.json.ext, '.tex');
    assert.equal(tex.json.text, '\\documentclass{article}\n');

    const txt = await callRoute(server.text, { url: textUrl(root, 'run.txt') });
    assert.equal(txt.status, 200, '.txt 不在清单里，但原文端点照样认');
    assert.equal(txt.json.text, 'plain text');
    assert.equal(txt.json.lines, 1);

    const empty = await callRoute(server.text, { url: textUrl(root, 'empty.md') });
    assert.equal(empty.status, 200);
    assert.equal(empty.json.text, '');
    assert.equal(empty.json.bytes, 0);
    assert.equal(empty.json.lines, 1, '空文件的 lines 口径与 archive/tex 一致（split 后 1 段）');
    assert.equal(empty.json.hint, undefined);

    const crlf = await callRoute(server.text, { url: textUrl(root, 'crlf.md') });
    assert.equal(crlf.json.text, 'a\r\nb\r\n', '\\r 原样保留，按 \\n 分行不丢行');
    assert.equal(crlf.json.lines, 3);

    const bom = await callRoute(server.text, { url: textUrl(root, 'bom.md') });
    assert.equal(bom.json.text, '# BOM\n', 'TextDecoder 默认吃掉 UTF-8 BOM');
    assert.equal(bom.json.bytes, 9, 'bytes 是磁盘上的真实大小');
});

await check('文本：超过 2 MiB 只读前 2 MiB，truncated/hint/bytes 对，且不切出半个汉字', async () => {
    const root = await freshRoot();
    // 让截断点正好落在 3 字节汉字的中间：limit-2 个 ASCII 之后放「汉」。
    const big = Buffer.concat([
        Buffer.alloc(TEXT_READ_LIMIT - 2, 0x61),
        Buffer.from('汉', 'utf8'),
        Buffer.alloc(4096, 0x62),
    ]);
    await writeFile(join(root, 'big.md'), big);
    const server = mount({ root });
    const res = await callRoute(server.text, { url: textUrl(root, 'big.md') });
    assert.equal(res.status, 200);
    assert.equal(res.json.bytes, big.length, 'bytes 报真实大小，不是读到的字节数');
    assert.equal(res.json.truncated, true);
    assert.ok(Buffer.byteLength(res.json.text, 'utf8') <= TEXT_READ_LIMIT, '正文不能超过读上限');
    assert.equal(Buffer.byteLength(res.json.text, 'utf8'), TEXT_READ_LIMIT - 2, '残缺汉字整段丢掉，只剩前面的 ASCII');
    assert.ok(res.json.text.endsWith('a'));
    assert.ok(!res.json.text.includes('\uFFFD'), '截断点不能解出一个替换字符');
    assert.ok(!res.json.text.includes('汉'));
    assert.equal(res.json.lines, 1);
    assert.match(res.json.hint, /只预览前 2 MiB/);
    assert.match(res.json.hint, /共 \d+\.\d MiB/);
});

await check('文本：二进制内容被当成 .md 时 200 + 「疑似二进制」hint，不 500', async () => {
    const root = await freshRoot();
    const blob = Buffer.concat([Buffer.from('MZ'), Buffer.alloc(200, 0xFF)]);
    await writeFile(join(root, 'blob.md'), blob);
    const server = mount({ root });
    const res = await callRoute(server.text, { url: textUrl(root, 'blob.md') });
    assert.equal(res.status, 200, '解不出来也要给面板东西看，而不是 500');
    assert.equal(res.json.truncated, false);
    assert.ok(res.json.text.includes('\uFFFD'));
    assert.match(res.json.hint, /疑似二进制/);
    assert.equal(res.json.bytes, blob.length);
});

await check('文本：二进制扩展名与未知扩展名 415，hint 指向渲染通道', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'deck.pptx'), 'ppt');
    await writeFile(join(root, 'report.pdf'), '%PDF-1.4');
    await writeFile(join(root, 'photo.png'), 'img');
    await writeFile(join(root, 'thing.bin'), 'x');
    const server = mount({ root });
    for (const name of ['deck.pptx', 'report.pdf', 'photo.png', 'thing.bin']) {
        const res = await callRoute(server.text, { url: textUrl(root, name) });
        assert.equal(res.status, 415, `${name} 应走渲染通道而不是原文`);
        assert.match(res.json.error, /不支持原文预览/);
        assert.match(res.json.hint, /渲染通道/);
    }
});

await check('文本：缺 path 400、穿越/绝对路径 404、未知 cwd 404、文件不存在 404，都带 hint', async () => {
    const root = await freshRoot();
    const outside = join(TMP, `outside-text-${caseId}.md`);
    await writeFile(outside, '# outside');
    const server = mount({ root });

    const missing = await callRoute(server.text, { url: `${TEXT_VIEW_PATH}?cwd=${encodeURIComponent(root)}` });
    assert.equal(missing.status, 400);
    assert.match(missing.json.error, /path 不能为空/);
    assert.equal(typeof missing.json.hint, 'string');
    assert.match(missing.json.hint, /相对路径/);

    const escape = await callRoute(server.text, {
        url: `${TEXT_VIEW_PATH}?cwd=${encodeURIComponent(root)}&path=${encodeURIComponent('../outside.md')}`,
    });
    assert.equal(escape.status, 404);
    assert.match(escape.json.error, /工作目录内/);
    assert.match(escape.json.hint, /刷新/);

    const absolute = await callRoute(server.text, { url: textUrl(root, outside) });
    assert.equal(absolute.status, 404);
    assert.match(absolute.json.error, /相对 cwd/);

    const untrusted = await callRoute(server.text, {
        url: `${TEXT_VIEW_PATH}?cwd=${encodeURIComponent(join(root, 'nope'))}&path=notes.md`,
    });
    assert.equal(untrusted.status, 404);
    assert.match(untrusted.json.error, /不在已知列表/);
    assert.match(untrusted.json.hint, /刷新|办公工具/);

    const ghost = await callRoute(server.text, { url: textUrl(root, 'ghost.md') });
    assert.equal(ghost.status, 404);
    assert.match(ghost.json.error, /文件不存在/);
    assert.match(ghost.json.hint, /刷新/);
});

await check('文本：POST 等非只读方法 405 + Allow: GET, HEAD', async () => {
    const root = await freshRoot();
    await writeFile(join(root, 'notes.md'), '# x');
    const server = mount({ root });
    for (const method of ['POST', 'PUT', 'DELETE', 'PATCH']) {
        const res = await callRoute(server.text, { method, url: textUrl(root, 'notes.md') });
        assert.equal(res.status, 405, `${method} 应被拒绝`);
        assert.equal(res.headers.allow, 'GET, HEAD');
        assert.match(res.json.error, /只读/);
        assert.match(res.json.hint, /GET/);
    }
});

// ── 注册与释放 ──────────────────────────────────────────────────────────────

await check('注册：四条 exact 路由、没有 webServer 时返回 undefined、dispose 全部撤掉', async () => {
    resetWorkspaces();
    const root = await freshRoot();
    noteWorkspace(root);
    const ctx = fakeCtx();
    const dispose = registerPreviewView(ctx, { getMemory: () => ({}), render: fakeRender(), loadEngine: async () => ({ module: {} }) });
    assert.equal(typeof dispose, 'function');
    assert.deepEqual(ctx.routes.map((item) => item.path).sort(), Object.values(PREVIEW_VIEW_PATHS).sort());
    assert.equal(ctx.routes.length, 4, '清单 / 大纲 / 单页图 / 文本原文');
    assert.ok(ctx.routes.some((item) => item.path === TEXT_VIEW_PATH), '文本端点必须注册上');
    assert.ok(ctx.routes.every((item) => item.kind === 'exact'));
    dispose();
    assert.equal(ctx.routes.length, 0, '释放后不该留下任何路由');

    assert.equal(registerPreviewView({ get: () => undefined }, {}), undefined, '没有 webServer 时静默返回');
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`preview-view: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
