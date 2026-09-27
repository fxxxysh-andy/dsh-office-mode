/**
 * 记忆浏览端点的测试：只读路由 + 信任边界。
 *
 * 这一份测的重点不是「面板长什么样」，而是**端点不该给什么**：
 *   - 只接受 GET，其它方法一律 405；
 *   - 只服务「本进程真的跑过办公工具的工作目录」，未知路径 404 而不是照读；
 *   - 记忆关掉时不泄漏内容；
 *   - 快照内容与引擎口径一致（热记忆 / 台账 / 归档 / 关系 / 实体 / 计数）。
 *
 * 跑法：node test/view.mjs
 */
import { mkdir, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import { createMemory } from '../src/memory.js';
import { buildSnapshot, knownWorkspaces, noteWorkspace, registerMemoryView, resetWorkspaces, MEMORY_VIEW_PATH } from '../src/view.js';

// 测试中间产物统一落在仓库根的 .office/tmp/<suite>/ 下（不再用包内的 test/.tmp）。
const TMP = fileURLToPath(new URL('../../../.office/tmp/view/', import.meta.url));

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
async function freshRoot(memoryConfig = {}) {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    await rm(root, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    const memory = { globalDir: join(TMP, `global-${caseId}`), ...memoryConfig };
    return { root, memory };
}

/** 造一个假的 req/res，把响应体收下来。 */
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

async function callRoute(handler, { method = 'GET', url = MEMORY_VIEW_PATH } = {}) {
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

/** 一个只提供 webServer.register 的插件上下文替身。 */
function fakeCtx() {
    const routes = [];
    return {
        routes,
        get(name) {
            if (name !== 'webServer') return undefined;
            return {
                register(route) {
                    routes.push(route);
                    return () => {
                        const at = routes.indexOf(route);
                        if (at >= 0) routes.splice(at, 1);
                    };
                },
            };
        },
    };
}

await rm(TMP, { recursive: true, force: true });

// ── 快照内容 ────────────────────────────────────────────────────────────────

await check('快照：三层内容、关系、实体与计数都按引擎口径给出来', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    const a = await instance.mutate({ action: 'add', target: 'user', content: '用户偏好：汇报用 16:9。', entities: ['汇报模板'] });
    const b = await instance.mutate({ action: 'add', target: 'project', content: '项目约定：主题用 business。' });
    await instance.link({ sourceId: a.entry.id, targetId: b.entry.id, kind: 'related', note: '同一条稿子' });
    await instance.log([{ path: '季度汇报.pptx', format: 'ppt', theme: 'business', purpose: '季度主线' }]);
    await instance.appendArchiveBlock({ reason: 'mnemon-migration', items: [{ id: 'i-1', kind: 'mnemon', content: '迁来的长期记忆', entities: ['实体甲'] }] });

    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.ok, true);
    assert.equal(snap.cwd, root);
    assert.equal(snap.label, root.split(/[\\/]/).filter(Boolean).pop());
    assert.equal(snap.hot.length, 2);
    assert.equal(snap.ledger.length, 1);
    assert.ok(snap.archive.length >= 1);
    assert.equal(snap.links.length, 1);
    assert.equal(snap.counts.hot, 2);
    assert.equal(snap.counts.links, 1);

    const user = snap.hot.find((item) => item.target === 'user');
    assert.deepEqual(user.entities, ['汇报模板'], '实体要随条目出来');
    assert.equal(user.origin, 'workspace');

    assert.equal(snap.links[0].kind, 'related');
    assert.equal(snap.links[0].note, '同一条稿子');
    assert.ok(snap.entities.some((item) => item.name === '实体甲'), '归档里的实体也要进实体视图');

    // 配置回显：面板要按这些值解释自己的显示。
    assert.equal(snap.config.scope, 'workspace');
    assert.equal(snap.config.quality.policy, 'strict-v1');
    assert.equal(snap.config.layers.hot, true);
    assert.deepEqual(snap.stores.map((store) => store.id), ['workspace']);

    // 体积字段（冻结契约）：面板要显示每个根的占用与两层正文的体积。
    assert.equal(snap.stores.length, 1);
    const store = snap.stores[0];
    assert.equal(typeof store.bytes, 'number', 'stores[].bytes 必须是数字');
    assert.equal(typeof store.files, 'number', 'stores[].files 必须是数字');
    assert.ok(store.files >= 3, `至少有 memory.json / USER.md / ledger.jsonl，实际 ${store.files}`);
    assert.ok(store.bytes > 0, `有内容的记忆库体积应大于 0，实际 ${store.bytes}`);

    const userBytes = Buffer.byteLength('用户偏好：汇报用 16:9。', 'utf8') + Buffer.byteLength('\n§\n', 'utf8');
    const projectBytes = Buffer.byteLength('项目约定：主题用 business。', 'utf8') + Buffer.byteLength('\n§\n', 'utf8');
    assert.equal(snap.counts.hotBytes, userBytes, 'counts.hotBytes 只算 target:user 的正文');
    assert.equal(snap.counts.projectBytes, projectBytes, 'counts.projectBytes 只算 target:project 的正文');

    // 容量口径与「列表里有几条」是两个量（面板的容量条用前者）：
    //   ledgerTotal  —— 未过滤、未被浏览上限截断的台账总数；
    //   archiveFiles —— 归档**摘要文件**数（archiveKeep 管的是它，一个月一个）；
    //   archiveItems —— 归档条目数（与 archive.length 同源，但不受过滤影响）。
    assert.equal(snap.counts.ledgerTotal, 1, 'ledgerTotal 是台账总数');
    assert.equal(snap.counts.archiveFiles, 1, '一条 appendArchiveBlock 只产生 1 个摘要文件');
    assert.equal(snap.counts.archiveItems, 1, 'archiveItems 是归档条目数');
    assert.equal(snap.stores[0].archiveFiles, 1, 'stores[] 也要给摘要文件数');
});

await check('快照：容量口径不跟着搜索框走（q 只影响列表条数）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.log([
        { path: '季度汇报.pptx', format: 'ppt' },
        { path: '年会策划.pptx', format: 'ppt' },
        { path: '无关的.docx', format: 'word' },
    ]);
    await instance.appendArchiveBlock({ reason: 'mnemon-migration', items: [
        { id: 'i-1', kind: 'mnemon', content: '归档里的季度结论' },
        { id: 'i-2', kind: 'mnemon', content: '另一条不相干的旧记录' },
    ] });

    const all = await buildSnapshot({ root, memory });
    assert.equal(all.counts.ledger, 3);
    assert.equal(all.counts.ledgerTotal, 3);
    assert.equal(all.counts.archive, 2);
    assert.equal(all.counts.archiveFiles, 1);

    const filtered = await buildSnapshot({ root, memory, query: '季度' });
    assert.equal(filtered.counts.ledger, 1, '列表条数跟着过滤走（这是对的：页签与指标卡要它）');
    assert.equal(filtered.counts.archive, 1);
    assert.equal(filtered.counts.ledgerTotal, 3, '台账容量口径不跟着过滤缩水');
    assert.equal(filtered.counts.archiveItems, 2, '归档条目数不跟着过滤缩水');
    assert.equal(filtered.counts.archiveFiles, 1, '摘要文件数更不该跟着过滤变');
});

await check('快照：空目录的体积字段是 0，不是 undefined', async () => {
    const { root, memory } = await freshRoot();
    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.ok, true);
    assert.equal(snap.hot.length, 0);
    assert.equal(snap.stores.length, 1);
    assert.equal(snap.stores[0].bytes, 0, '目录还没建出来时体积是 0');
    assert.equal(snap.stores[0].files, 0, '目录还没建出来时文件数是 0');
    assert.equal(snap.counts.hotBytes, 0, '没有条目时 hotBytes 是 0 而不是 undefined');
    assert.equal(snap.counts.projectBytes, 0, '没有条目时 projectBytes 是 0 而不是 undefined');
});

await check('快照：q 过滤覆盖四段（热记忆 / 台账 / 归档 / 关系）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    const a = await instance.mutate({ action: 'add', target: 'project', content: '热记忆里的一条：季度汇报用 16:9。' });
    const b = await instance.mutate({ action: 'add', target: 'project', content: '另一条不相干的项目约定。' });
    await instance.link({ sourceId: a.entry.id, targetId: b.entry.id, kind: 'supports', note: '季度汇报的配色依据' });
    await instance.log([
        { path: '季度汇报.pptx', format: 'ppt' },
        { path: '年会策划.pptx', format: 'ppt' },
    ]);
    await instance.appendArchiveBlock({ reason: 'mnemon-migration', items: [{ id: 'i-1', kind: 'mnemon', content: '归档里的季度结论' }] });

    const snap = await buildSnapshot({ root, memory, query: '季度' });
    assert.equal(snap.query, '季度');
    assert.equal(snap.hot.length, 1, '热记忆也要参与过滤');
    assert.match(snap.hot[0].content, /季度/);
    assert.equal(snap.ledger.length, 1);
    assert.equal(snap.ledger[0].path, '季度汇报.pptx');
    assert.equal(snap.archive.length, 1, '归档也要参与过滤');
    assert.equal(snap.links.length, 1, '关系按两端 id / 类型 / 备注参与过滤');
    assert.match(snap.links[0].note, /季度/);

    const all = await buildSnapshot({ root, memory });
    assert.equal(all.hot.length, 2, '不带 q 时不过滤');
    assert.equal(all.ledger.length, 2);
    assert.equal(all.links.length, 1);
});

// ── 路由与信任边界 ──────────────────────────────────────────────────────────

await check('路由：组合里没有 webServer 时安静地不注册', async () => {
    const dispose = registerMemoryView({ get: () => undefined }, { getMemory: () => ({}) });
    assert.equal(dispose, undefined);
});

await check('路由：注册到 MEMORY_VIEW_PATH，且是可释放的', async () => {
    const ctx = fakeCtx();
    const dispose = registerMemoryView(ctx, { getMemory: () => ({}) });
    assert.equal(ctx.routes.length, 1);
    assert.equal(ctx.routes[0].kind, 'exact');
    assert.equal(ctx.routes[0].path, MEMORY_VIEW_PATH);
    assert.equal(typeof dispose, 'function');
    dispose();
    assert.equal(ctx.routes.length, 0, '释放后路由应被摘掉');
});

await check('路由：未知工作目录一律 404，不会照路径去读盘', async () => {
    resetWorkspaces();
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({}) });
    const handler = ctx.routes[0].handler;

    const empty = await callRoute(handler);
    assert.equal(empty.json.ok, false);
    assert.match(empty.json.error, /还没有见过任何工作目录/);

    const unknown = await callRoute(handler, { url: `${MEMORY_VIEW_PATH}?cwd=C:/Windows` });
    assert.equal(unknown.status, 404, '没见过的路径必须是 404');
    assert.match(unknown.json.error, /不在已知列表里/);
    assert.equal(unknown.json.hot, undefined, '拒绝时不能带任何内容');
});

await check('路由：只服务 noteWorkspace 记下的根', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'project', content: '这个目录里的一条约定。' });
    noteWorkspace(root);
    assert.deepEqual(knownWorkspaces().map((item) => item.cwd), [root]);

    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memory });
    const handler = ctx.routes[0].handler;

    const snap = await callRoute(handler);
    assert.equal(snap.status, 200);
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 1);
    assert.equal(snap.headers['cache-control'], 'no-store', '记忆内容不该被缓存');
    assert.equal(snap.json.workspaces.length, 1);

    const byCwd = await callRoute(handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(root)}` });
    assert.equal(byCwd.json.ok, true);
    assert.equal(byCwd.json.cwd, root);
});

await check('路由：非 GET 方法一律 405（这个端点只读）', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    noteWorkspace(root);
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memory });
    const handler = ctx.routes[0].handler;
    for (const method of ['POST', 'PUT', 'DELETE', 'PATCH']) {
        const res = await callRoute(handler, { method });
        assert.equal(res.status, 405, `${method} 应被拒绝`);
        assert.match(res.json.error, /只读/);
    }
    const head = await callRoute(handler, { method: 'HEAD' });
    assert.equal(head.status, 200, 'HEAD 允许');
});

await check('路由：记忆关掉时不返回任何内容', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'project', content: '不该被读出来的内容。' });
    noteWorkspace(root);

    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({ ...memory, enabled: false }) });
    const res = await callRoute(ctx.routes[0].handler);
    assert.equal(res.json.ok, false);
    assert.match(res.json.error, /关掉/);
    assert.equal(res.json.hot, undefined);
    assert.doesNotMatch(res.raw, /不该被读出来的内容/);
});

await check('路由：内部出错时给出可读错误，而不是把异常抛给服务器', async () => {
    resetWorkspaces();
    noteWorkspace(join(TMP, 'case-ghost'));
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({}) });
    // 目录不存在时引擎照常给出空结果；这里要确认的是「不会 500 崩掉」。
    const res = await callRoute(ctx.routes[0].handler);
    assert.ok(res.status === 200 || res.status === 500);
    if (res.status === 500) assert.match(res.json.error, /读取记忆失败/);
    else assert.equal(res.json.ok, true);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`view: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
