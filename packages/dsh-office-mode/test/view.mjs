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
import { mkdir, rm, writeFile } from 'node:fs/promises';
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

/**
 * 一个插件上下文替身。
 *
 * `services` 里的东西按名字返回（webServer 之外还认 workspaceRegistry 与
 * sessionPersistence —— 第三十一轮起端点用它们当工作目录的信任来源）；
 * 没给的按「宿主没有这个服务」处理，返回 undefined。
 */
function fakeCtx(services = {}) {
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
            return services[name];
        },
    };
}

/** 宿主工作区登记表的替身：`list()` 返回 [{path, title}]。 */
function fakeRegistry(items) {
    return { list: () => items };
}

/** 会话存储的替身：`list()` 返回 [{header: {cwd, createdAt}}]。 */
function fakeSessions(items) {
    return { list: async () => items };
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

await check('快照：知识库文档清单（kb 一期）与计数、逐根计数都给出来', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await writeFile(join(root, '来源笔记.md'), '# 方法论\n\n真实 Office 才算证据。\n\n## 第二段\n\n复核要能复跑。\n', 'utf8');
    const ingested = await instance.kbIngest({ path: '来源笔记.md', tier: 'verified' });
    assert.equal(ingested.chunks > 0, true, '夹具要先真的切出块');

    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.counts.kb, 1, '知识库页签的条数 = 清单里有几篇');
    assert.equal(snap.counts.kbTotal, 1, 'kbTotal 是未过滤总数');
    assert.equal(snap.counts.kbChunks, ingested.chunks, 'kbChunks 是块总数');
    assert.equal(snap.counts.kbBytes > 0, true, 'kbBytes 是文档总字节');
    assert.equal(snap.counts.kbTruncated, false, '没到单次上限就不算截断');
    assert.equal(snap.kb.length, 1);

    const doc = snap.kb[0];
    assert.equal(doc.path, '来源笔记.md');
    assert.equal(doc.tier, 'verified');
    assert.equal(doc.chunks, ingested.chunks);
    assert.equal(doc.title.includes('方法论'), true, '标题从正文一级标题取');
    assert.equal(doc.id.startsWith('kb:'), true, '文档 id 是 kb: 形');
    assert.equal(doc.origin, 'workspace');
    assert.equal(typeof doc.at, 'string');

    // 逐根计数：面板的存储卡按它画「知识库 N 篇 / M 块」。
    assert.equal(snap.stores[0].kb.docs, 1, 'stores[].kb.docs 要有');
    assert.equal(snap.stores[0].kb.chunks, ingested.chunks, 'stores[].kb.chunks 要有');
    assert.equal(snap.stores[0].kb.tiers.verified, 1, '档位分布要有');

    // 过滤：与其它层同口径（按路径 / 标题 / 档位匹配）。
    const hit = await buildSnapshot({ root, memory, query: '方法论' });
    assert.equal(hit.counts.kb, 1, 'kb 也参与 q 过滤');
    const miss = await buildSnapshot({ root, memory, query: '不存在的词' });
    assert.equal(miss.counts.kb, 0, '过滤后列表条数跟着变');
    assert.equal(miss.counts.kbTotal, 1, 'kbTotal 不跟着过滤缩水');
});

await check('快照：没有知识库时 kb 是空数组、计数为 0（不是 undefined）', async () => {
    const { root, memory } = await freshRoot();
    const snap = await buildSnapshot({ root, memory });
    assert.deepEqual(snap.kb, [], '.office/kb 不存在时给空数组');
    assert.equal(snap.counts.kb, 0);
    assert.equal(snap.counts.kbTotal, 0);
    assert.equal(snap.counts.kbChunks, 0);
    assert.equal(snap.counts.kbBytes, 0);
    assert.equal(snap.stores[0].kb.docs, 0, 'stores[].kb 也要给 0 而不是 null');
});

await check('快照：知识库文档超过清单上限时如实报「截断 + 总数」', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    // 清单接口的硬上限是 20 篇（KB_LIST_LIMIT）。端点声明的上限必须与它一致，
    // 否则面板会以为「要更多却只回来 20 篇」是丢了东西。
    for (let index = 0; index < 23; index += 1) {
        const rel = `来源/第${String(index).padStart(2, '0')}篇.md`;
        const absolute = join(root, '来源', `第${String(index).padStart(2, '0')}篇.md`);
        await mkdir(join(root, '来源'), { recursive: true });
        await writeFile(absolute, `# 第 ${index} 篇\n\n这是第 ${index} 篇的正文，用来把清单推到上限之外。\n`, 'utf8');
        await instance.kbIngest({ path: rel, tier: 'user' });
    }

    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.counts.kbTotal, 23, 'kbTotal 是真实总数');
    assert.equal(snap.kb.length, 20, '列表按清单上限截断（20 篇）');
    assert.equal(snap.counts.kb, snap.kb.length, '页签条数 = 列表里有几篇');
    assert.equal(snap.counts.kbTruncated, true, '截断了就要说');
    assert.equal(snap.stores[0].kb.docs, 23, '逐根计数仍是真实总数');
    assert.equal(snap.stores[0].kb.chunks, snap.counts.kbChunks, '逐根与合计同源（都从 manifest 现算）');
    assert.equal(snap.stores[0].kb.bytes, snap.counts.kbBytes, '逐根与合计同源');

    // **先过滤、再截断**：搜一篇落在「最新 20 篇」之外的旧文档，必须搜得到。
    // 早期实现是先把 20 篇截出来再过滤，于是这一条会得到 0 条（复核 P2-1）。
    // 注意查询词不能带空格：带空格会被拆成「与」关系的多个词，`第` / `篇` 命中全部 23 篇。
    const oldest = await buildSnapshot({ root, memory, query: '第00篇' });
    assert.equal(oldest.counts.kb, 1, '搜索要覆盖全部文档，不是只覆盖最新 20 篇');
    assert.equal(oldest.kb[0].path, '来源/第00篇.md');
    assert.equal(oldest.counts.kbTotal, 23, 'kbTotal 不跟着过滤缩水');
});

await check('快照：q 过滤覆盖五段（热记忆 / 台账 / 归档 / 知识库 / 关系）', async () => {
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

// ── 第三十一轮：重启之后、以及还没跑过办公工具的工作区 ──────────────────────
//
// 老实现的信任来源只有「本进程真跑过工具的工作目录」（进程内），于是重启 DSH
// 之后面板只能显示「还没有见过任何工作目录」—— 记忆文件明明在盘上。
// 现在补两个来源：宿主工作区登记表（持久、跨重启）与会话史（兜底）。

await check('重启后：工作区登记表里的目录不用先跑工具也能显示记忆', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'project', content: '重启前写下的约定。' });
    await instance.log([{ path: '季度汇报.pptx', format: 'ppt', purpose: '季度主线' }]);

    const ctx = fakeCtx({ workspaceRegistry: fakeRegistry([{ path: root, title: '季度项目' }]) });
    registerMemoryView(ctx, { getMemory: () => memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.status, 200, '登记过的目录不该再 404');
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 1);
    assert.equal(snap.json.ledger.length, 1);
    assert.equal(snap.json.label, '季度项目', '面板用登记表的标题当标签');
    assert.equal(snap.json.workspaces.length, 1);
    assert.equal(snap.json.workspaces[0].source, 'workspace');
    assert.equal(snap.json.workspaces[0].label, '季度项目');
});

await check('未使用的工作区：没有记忆目录时给空态，不是 404 也不是报错', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const ctx = fakeCtx({ workspaceRegistry: fakeRegistry([{ path: root, title: '还没开工的项目' }]) });
    registerMemoryView(ctx, { getMemory: () => memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.status, 200);
    assert.equal(snap.json.ok, true, '工作区里还没有记忆也该给一份空快照');
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 0);
    assert.equal(snap.json.ledger.length, 0);
    assert.equal(snap.json.counts.ledgerTotal, 0);
    assert.equal(snap.json.stores[0].bytes, 0);
});

await check('重启后默认根：先开有记忆、且最近在干活的那个工作区', async () => {
    resetWorkspaces();
    const older = await freshRoot();
    const newer = await freshRoot();
    await createMemory({ root: older.root, memory: older.memory }).mutate({ action: 'add', target: 'project', content: '旧项目的记忆。' });
    // 两个记忆目录的 mtime 必须真的分开：没有工作区时间戳时默认根的判据会退到它，
    // 同毫秒内写完两个目录时这条断言就变成掷骰子。
    await new Promise((resolve) => setTimeout(resolve, 25));
    await createMemory({ root: newer.root, memory: newer.memory }).mutate({ action: 'add', target: 'project', content: '新项目的记忆。' });
    // 登记表顺序是「用户自己排的」：旧项目在前。面板不该照搬它，而该挑有内容的那个。
    const ctx = fakeCtx({
        workspaceRegistry: fakeRegistry([
            { path: older.root, title: '旧项目' },
            { path: newer.root, title: '新项目' },
        ]),
    });
    registerMemoryView(ctx, { getMemory: () => newer.memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.workspaces.length, 2, '两个工作区都要列出来给面板切换');
    // 都没有工作区时间戳时，取记忆库文件最近动过的那个（这里后写的 newer 更新）
    assert.equal(snap.json.cwd, newer.root);
    assert.equal(snap.json.label, '新项目');

    // 显式指定时按指定来
    const byCwd = await callRoute(ctx.routes[0].handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(older.root)}` });
    assert.equal(byCwd.json.ok, true);
    assert.equal(byCwd.json.cwd, older.root);
    assert.match(byCwd.json.hot[0].content, /旧项目/);
});

await check('重启后默认根：工作区记录的最近变动时刻优先于记忆文件的 mtime', async () => {
    resetWorkspaces();
    const busy = await freshRoot();
    const idle = await freshRoot();
    await createMemory({ root: idle.root, memory: idle.memory }).mutate({ action: 'add', target: 'project', content: '刚写过记忆但没人在干活的项目。' });
    await new Promise((resolve) => setTimeout(resolve, 25));
    await createMemory({ root: busy.root, memory: busy.memory }).mutate({ action: 'add', target: 'project', content: '最近在干活的项目。' });
    // 故意把 idle 的记忆改到更晚，再把 busy 的工作区时间戳设成更晚：
    // 「最近在哪儿干活」该赢过「哪个记忆文件最近被碰过」。
    const ctx = fakeCtx({
        workspaceRegistry: fakeRegistry([
            { path: idle.root, title: '闲项目', updatedAt: '2026-01-01T00:00:00.000Z' },
            { path: busy.root, title: '忙项目', updatedAt: '2026-09-28T00:00:00.000Z' },
        ]),
    });
    registerMemoryView(ctx, { getMemory: () => busy.memory });
    const snap = await callRoute(ctx.routes[0].handler);
    assert.equal(snap.json.cwd, busy.root, '工作区时间戳更晚的那个才是「最近在干活」的');
    assert.equal(snap.json.label, '忙项目');
});

await check('工作区登记表为空时：会话史里的工作目录兜底（不登记成项目也能显示）', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    await createMemory({ root, memory }).mutate({ action: 'add', target: 'project', content: '会话史兜底的记忆。' });
    const ctx = fakeCtx({
        sessionPersistence: fakeSessions([
            { header: { cwd: root, createdAt: 1_700_000_000_000 } },
            { header: { cwd: '', createdAt: 1_700_000_000_001 } },
            { header: { createdAt: 1_700_000_000_002 } },
        ]),
    });
    registerMemoryView(ctx, { getMemory: () => memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.status, 200);
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 1);
    assert.equal(snap.json.workspaces[0].source, 'session');

    // 登记表里没有、会话史里有：显式点名时也认
    const byCwd = await callRoute(ctx.routes[0].handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(root)}` });
    assert.equal(byCwd.json.ok, true);
    assert.equal(byCwd.json.cwd, root);
});

await check('工作区登记表里有别的项目时：仍然不认登记表之外的路径', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const ctx = fakeCtx({ workspaceRegistry: fakeRegistry([{ path: root, title: '已知项目' }]) });
    registerMemoryView(ctx, { getMemory: () => memory });
    const res = await callRoute(ctx.routes[0].handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(join(TMP, 'case-unknown'))}` });
    assert.equal(res.status, 404, '登记表之外的路径必须还是 404');
    assert.equal(res.json.hot, undefined);
    assert.equal(res.json.workspaces.length, 1, '拒绝时仍把已知工作区带回给面板');
});

await check('宿主服务坏掉时降级：登记表/会话史抛错不把端点打挂', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const ctx = fakeCtx({
        workspaceRegistry: { list: () => { throw new Error('存储坏了'); } },
        sessionPersistence: { list: async () => { throw new Error('会话存储坏了'); } },
    });
    registerMemoryView(ctx, { getMemory: () => memory });
    const res = await callRoute(ctx.routes[0].handler);
    assert.equal(res.status, 200, '拿不到宿主事实时退回老行为（空态），不该 500');
    assert.equal(res.json.ok, false);
    assert.match(res.json.error, /还没有见过任何工作目录/);

    // 本进程跑过工具的根不受影响
    noteWorkspace(root);
    const snap = await callRoute(ctx.routes[0].handler);
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
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
