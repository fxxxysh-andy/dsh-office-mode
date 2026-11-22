/**
 * 知识库资源管理器口径：默认只列当前层 + 显式空目录合并 + 新建虚拟文件夹端点。
 *
 * 覆盖：
 *   - buildSnapshot 默认 kbRecursive=false（根/子层都只给直属文件），显式 true 才递归；
 *   - 快照 kbTree 合并 folders.json（空目录 docs=0, virtual=true），与当前页/当前目录无关；
 *   - POST /office-memory/mutate 的 kb-folder-create：新建、幂等、非法路径 400；
 *   - 空目录进入后分页计数为 0（不报错、不落回根目录）。
 *
 * 跑法：node test/kb-folders-view.mjs
 */
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { buildSnapshot, noteWorkspace, registerMemoryView, resetWorkspaces } from '../src/view.js';
import { createMemory } from '../src/memory.js';
import { createKbFolder } from '../src/kb-folders.js';

const TMP = fileURLToPath(new URL('../../../.office/runtime/tmp/kb-folders-view/', import.meta.url));
process.env.DSH_HOME = join(TMP, 'home');

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
    const root = join(TMP, 'case-' + caseId);
    await rm(root, { recursive: true, force: true });
    await mkdir(join(root, 'docs'), { recursive: true });
    const home = join(root, 'home');
    await mkdir(home, { recursive: true });
    const kbDir = join(root, '.office', 'kb');
    const config = { memory: { kbDir } };
    const memory = createMemory({ root, memory: config.memory });
    const seed = async (name, folder, text) => {
        const file = join(root, 'docs', name);
        await writeFile(file, '# ' + name + '\n\n' + (text || '正文。') + '\n', 'utf8');
        await memory.kbIngest({ path: file, folder, displayPath: 'docs/' + name });
    };
    await seed('a.md', '资料');
    await seed('b.md', '资料/子');
    await seed('c.md', '资料');
    return { root, home, kbDir, config, memory, seed };
}

function findNode(node, wanted) {
    if (node === null || typeof node !== 'object') return null;
    if (String(node.path ?? '') === wanted) return node;
    for (const child of Array.isArray(node.children) ? node.children : []) {
        const found = findNode(child, wanted);
        if (found !== null) return found;
    }
    return null;
}

function fakeWebServer() {
    const routes = new Map();
    return {
        routes,
        register(definition) {
            routes.set(definition.path, definition.handler);
            return () => {};
        },
    };
}

function callRoute(handler, { method = 'GET', url = '/', body = '', headers = {} } = {}) {
    return new Promise((resolve, reject) => {
        const response = {
            status: 0,
            headers: null,
            body: '',
            writeHead(status, nextHeaders) { this.status = status; this.headers = nextHeaders; },
            end(chunk) {
                this.body = String(chunk ?? '');
                let json = null;
                try { json = JSON.parse(this.body); } catch { json = null; }
                resolve({ status: this.status, headers: this.headers, body: this.body, json });
            },
        };
        const request = Readable.from([Buffer.from(body)]);
        request.method = method;
        request.url = url;
        request.headers = headers;
        Promise.resolve(handler(request, response)).catch(reject);
    });
}

function snapshotContext(root, memory) {
    const webServer = fakeWebServer();
    resetWorkspaces();
    noteWorkspace(root);
    registerMemoryView({ get: (name) => (name === 'webServer' ? webServer : null) }, {
        getMemory: () => memory,
        getConfig: () => ({}),
    });
    return webServer.routes;
}

await check('快照默认只列当前层；kbTree 始终合并空目录', async () => {
    const { root, kbDir, config } = await freshRoot();
    await createKbFolder(kbDir, '项目/空');

    const snap = await buildSnapshot({ root, memory: config.memory });
    assert.equal(snap.kbRecursive, false, '默认必须是「仅当前层」');
    assert.equal(snap.counts.kb, 0, '根目录直属文档为 0（三篇都在 资料 子树里）');
    assert.equal(snap.kbPage.total, 0);
    assert.equal(snap.kbPage.offset, 0);

    const project = findNode(snap.kbTree, '项目/空');
    assert.ok(project !== null, '空目录要进快照树');
    assert.equal(project.virtual, true);
    assert.equal(project.docs, 0);
    assert.ok(findNode(snap.kbTree, '资料/子') !== null, '派生的子树节点照常保留');
});

await check('进入子层：默认只给直属文件，显式 recursive 才含子文件夹', async () => {
    const { root, config } = await freshRoot();
    await createKbFolder(config.memory.kbDir, '项目/空');

    const direct = await buildSnapshot({ root, memory: config.memory, kbFolder: '资料' });
    assert.equal(direct.kbFolder, '资料');
    assert.equal(direct.counts.kb, 2, '资料 下直属两篇，不含 资料/子 那一篇');

    const recursive = await buildSnapshot({ root, memory: config.memory, kbFolder: '资料', kbRecursive: true });
    assert.equal(recursive.kbRecursive, true);
    assert.equal(recursive.counts.kb, 3, '显式递归时才含子文件夹里的文档');
});

await check('进入空目录：计数 0、仍在树里、面包屑路径保留', async () => {
    const { root, config } = await freshRoot();
    await createKbFolder(config.memory.kbDir, '项目/空');

    const snap = await buildSnapshot({ root, memory: config.memory, kbFolder: '项目/空' });
    assert.equal(snap.kbFolder, '项目/空');
    assert.equal(snap.counts.kb, 0);
    assert.equal(snap.kbPage.total, 0);
    assert.equal(snap.kbPage.page, 1, '空目录也要有一页（0 篇），不能报错');
    const node = findNode(snap.kbTree, '项目/空');
    assert.equal(node.virtual, true);
});

await check('GET 路由：没有 kbrecursive 时按「仅当前层」，显式 1 才递归', async () => {
    const { root, config } = await freshRoot();
    await createKbFolder(config.memory.kbDir, '项目/空');
    const routes = snapshotContext(root, config.memory);
    const handler = routes.get('/office-memory/snapshot');
    assert.ok(handler, '快照路由要注册');

    const plain = await callRoute(handler, { url: '/office-memory/snapshot?cwd=' + encodeURIComponent(root) });
    assert.equal(plain.status, 200);
    assert.equal(plain.json.ok, true);
    assert.equal(plain.json.kbRecursive, false);
    assert.equal(plain.json.counts.kb, 0);
    assert.ok(findNode(plain.json.kbTree, '项目/空') !== null, 'picker 读的同一棵树里必须有空目录');

    const recursive = await callRoute(handler, {
        url: '/office-memory/snapshot?cwd=' + encodeURIComponent(root) + '&kbfolder=' + encodeURIComponent('资料') + '&kbrecursive=1',
    });
    assert.equal(recursive.json.kbRecursive, true);
    assert.equal(recursive.json.counts.kb, 3);

    const empty = await callRoute(handler, {
        url: '/office-memory/snapshot?cwd=' + encodeURIComponent(root) + '&kbfolder=' + encodeURIComponent('项目/空'),
    });
    assert.equal(empty.json.kbFolder, '项目/空');
    assert.equal(empty.json.counts.kb, 0);
});

await check('仅当前层搜索：候选阶段精确过滤，子层不抢走直属命中', async () => {
    const { root, config, seed } = await freshRoot();
    // 直属文档只出现一次关键词，子层每个都堆满关键词；若先取 TopK 再按路径过滤，
    // 直属这篇会被 25 篇子层文档挤出窗口，最后只剩「无命中」。
    await seed('直属.md', '资料', '付款。');
    for (let i = 0; i < 25; i += 1) {
        await seed('子-' + i + '.md', '资料/子', '付款 付款 付款 付款 付款。');
    }

    const snap = await buildSnapshot({
        root, memory: config.memory, kbFolder: '资料', kbQuery: '付款', kbRecursive: false,
    });
    assert.ok(snap.kbSearch !== null, '搜索要有结果对象');
    assert.ok(snap.kbSearch.hits.some((hit) => hit.path === 'docs/直属.md'), '直属命中不能被子层挤掉');
    assert.ok(snap.kbSearch.hits.every((hit) => hit.path === 'docs/直属.md'), '默认不返回子层命中');
});

await check('POST kb-folder-create：新建 / 幂等 / 非法路径 400；只写 folders.json', async () => {
    const { root, kbDir, config } = await freshRoot();
    const routes = snapshotContext(root, config.memory);
    const handler = routes.get('/office-memory/mutate');
    assert.ok(handler, '写入路由要注册');
    const post = (payload) => callRoute(handler, {
        method: 'POST',
        url: '/office-memory/mutate?cwd=' + encodeURIComponent(root),
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(payload),
    });

    const created = await post({ action: 'kb-folder-create', path: '项目/新建' });
    assert.equal(created.status, 200);
    assert.equal(created.json.ok, true);
    assert.equal(created.json.result.created, true);
    assert.equal(created.json.result.path, '项目/新建');
    const folders = await readFile(join(kbDir, 'folders.json'), 'utf8');
    assert.match(folders, /项目\/新建/);

    const again = await post({ action: 'kb-folder-create', path: '项目/新建' });
    assert.equal(again.status, 200);
    assert.equal(again.json.ok, true);
    assert.equal(again.json.result.created, false, '同路径重复创建幂等');

    const invalid = await post({ action: 'kb-folder-create', path: '../逃逸' });
    assert.equal(invalid.status, 400, '非法路径要 400，不能 200 假装成功');
    assert.equal(invalid.json.ok, false);

    // 坏清单：写入必须拒绝，不能把损坏文件当空清单再覆盖掉。
    await writeFile(join(kbDir, 'folders.json'), '这不是 JSON', 'utf8');
    const broken = await post({ action: 'kb-folder-create', path: '项目/另一个' });
    assert.equal(broken.status, 409, '坏 folders.json 要 409，不能静默覆盖');
    assert.equal(broken.json.ok, false);
    assert.equal(broken.json.code, 'KB_FOLDERS_CORRUPT');
    assert.equal(await readFile(join(kbDir, 'folders.json'), 'utf8'), '这不是 JSON', '原文件字节不变');

    // 读取侧也不能装作「没有目录」：快照要带上可显示的原因。
    const brokenSnap = await buildSnapshot({ root, memory: config.memory });
    assert.match(brokenSnap.kbFoldersError, /folders\.json/, '坏清单要在视图里给出简短失败原因');

    await writeFile(join(kbDir, 'folders.json'), JSON.stringify({ version: 1, folders: ['项目/新建'] }), 'utf8');
    const snap = await buildSnapshot({ root, memory: config.memory });
    assert.ok(findNode(snap.kbTree, '项目/新建') !== null);
});

for (const result of results) {
    if (result.ok) console.log('PASS ' + result.name);
    else console.log('FAIL ' + result.name + '\n' + (result.error && result.error.stack ? result.error.stack : String(result.error)));
}
const failed = results.filter((item) => !item.ok).length;
console.log('kb-folders-view: ' + (results.length - failed) + '/' + results.length + ' 通过');
if (failed > 0) process.exitCode = 1;
