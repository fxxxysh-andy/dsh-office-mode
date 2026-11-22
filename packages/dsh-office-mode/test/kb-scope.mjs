/**
 * 会话级知识库范围（kb-session-scope）的测试。
 *
 * 这一份测「服务端强制」那条线：
 *   - 状态库：每会话一个哈希文件、原子写、未设置 / 损坏 / IO 错误三态分得开；
 *   - 范围解析：沿真实 header 链取交集，祖先缺失 / 循环 / 超深都 fail-closed；
 *   - 工具面：清单 / 树 / 检索 / 读取 / 删除 / 改 / 搬 / 入库 / 连边 / 摘要都被范围约束，
 *     参数里的 folder 既不能把自己放出去，也不能把范围外的东西拉进来；
 *   - 无会话（内部 / 测试）路径保持旧契约；真会话缺状态不静默全库；
 *   - /kb-scope 命令：查看 / 设置 / 清除，子代理拒绝写，不存在的文件夹当场说清。
 *
 * 跑法：node test/kb-scope.mjs
 */
import { mkdir, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import { kbDbPath } from '../src/kb-db.js';
import { createMemory } from '../src/memory.js';
import {
    createScopeStore,
    folderWithin,
    intersectFolders,
    kbTreeFolderPresence,
    kbTreeHasFolder,
    MAX_ANCESTRY,
    registerKbScopeCommand,
    resolveCallScope,
    resolveSessionScope,
    SCOPE_DIR,
    SCOPE_VERSION,
    scopeFileOf,
    scopeRefusalText,
} from '../src/kb-session-scope.js';
import { buildTools } from '../src/tools.js';

const TMP = fileURLToPath(new URL('../../../.office/runtime/tmp/kb-scope/', import.meta.url));
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
    await mkdir(root, { recursive: true });
    const home = join(root, 'home');
    await mkdir(home, { recursive: true });
    const memoryRoot = join(root, '.office', 'memory');
    const config = { memory: { kbDir: join(root, '.office', 'kb') } };
    return {
        root,
        home,
        config,
        store: createScopeStore({ root: join(home, SCOPE_DIR) }),
        memory: createMemory({ root, memory: config.memory }),
    };
}

function ctxOf(home) {
    const profileContext = { name: 'test', dir: home, home, cwd: home };
    return { get: (name) => (name === 'profileContext' ? profileContext : null) };
}

function execOf(sessionId, root, header = {}) {
    const id = sessionId;
    return {
        agent: { session: { id, header: { id, cwd: root, ...header } } },
        signal: new AbortController().signal,
    };
}

function headerOfFrom(map) {
    return async (id) => (map[id] === undefined ? undefined : map[id]);
}

/** 构造一个只注册 /kb-scope 的宿主上下文；headers 是活体 agents 表（解析父链用）。 */
function commandHarness({ home, headers, folderExists }) {
    let definition = null;
    const ctx = {
        get: (name) => {
            if (name === 'profileContext') return { name: 'test', home, cwd: home };
            if (name === 'commands') return { register: (def) => { definition = def; return () => {}; } };
            if (name === 'agents') {
                return { get: (id) => (headers[id] === undefined ? null : { session: { header: headers[id] } }) };
            }
            return null;
        },
    };
    registerKbScopeCommand(ctx, { folderExists });
    return {
        call: (rawInput, header) => definition.handler({ rawInput, agent: { session: { id: header.id, header } } }),
    };
}

const A = '资料';
const B = '其他';

await check('状态库：写 / 读 / 换会话；未设置与损坏分开报', async () => {
    const { store } = await freshRoot();
    assert.equal((await store.read('s1')).status, 'unset');
    await store.write('s1', { mode: 'folder', folder: ' 资料 / 合同 ', by: 'command' });
    const read = await store.read('s1');
    assert.equal(read.status, 'ok');
    assert.equal(read.record.mode, 'folder');
    assert.equal(read.record.folder, '资料/合同');
    assert.equal(read.record.v, SCOPE_VERSION);
    // 另一个会话读不到这一条。
    assert.equal((await store.read('s2')).status, 'unset');
    // 同一会话覆盖写：all 之后读回 all。
    await store.write('s1', { mode: 'all' });
    assert.equal((await store.read('s1')).record.mode, 'all');
    // 未设置 / 空串：unset，不是 corrupt。
    assert.equal((await store.read('')).status, 'unset');
});

await check('状态库：坏 JSON / 版本 / session 不符 / mode / 空 folder 都算 corrupt（fail-closed）', async () => {
    const { store } = await freshRoot();
    const id = 'bad-1';
    const file = scopeFileOf(store.root, id);
    await mkdir(store.root, { recursive: true });
    const write = (value) => writeFile(file, JSON.stringify(value) + '\n', 'utf8');
    await writeFile(file, '{ not json', 'utf8');
    assert.equal((await store.read(id)).status, 'corrupt');
    await write({ v: 99, session: id, mode: 'all' });
    assert.equal((await store.read(id)).status, 'corrupt');
    await write({ v: SCOPE_VERSION, session: 'someone-else', mode: 'all' });
    assert.equal((await store.read(id)).status, 'corrupt');
    await write({ v: SCOPE_VERSION, session: id, mode: 'weird' });
    assert.equal((await store.read(id)).status, 'corrupt');
    await write({ v: SCOPE_VERSION, session: id, mode: 'folder', folder: '' });
    assert.equal((await store.read(id)).status, 'corrupt');
    await write({ v: SCOPE_VERSION, session: id, mode: 'folder', folder: '../escape' });
    assert.equal((await store.read(id)).status, 'corrupt');
});

await check('边界：exact 与 descendant 才算范围内，前缀相同的兄弟目录不算', async () => {
    assert.equal(folderWithin('资料', '资料'), true);
    assert.equal(folderWithin('资料/合同', '资料'), true);
    assert.equal(folderWithin('资料/合同/2025', '资料'), true);
    assert.equal(folderWithin('资料夹', '资料'), false);
    assert.equal(folderWithin('资料x/合同', '资料'), false);
    assert.equal(folderWithin('其他', '资料'), false);
    assert.equal(intersectFolders('资料', '资料/合同'), '资料/合同');
    assert.equal(intersectFolders('资料/合同', '资料'), '资料/合同');
    assert.equal(intersectFolders('资料', '其他'), null);
    assert.equal(intersectFolders('', '资料'), '资料');
    assert.equal(intersectFolders('', ''), '');
});

await check('继承：子会话没记录→继承父范围；能往里收；收不进来→空范围；all 不能放大父范围', async () => {
    const { store } = await freshRoot();
    const headers = {
        root: { id: 'root', cwd: '/w' },
        child: { id: 'child', cwd: '/w', parentSession: 'root' },
        grand: { id: 'grand', cwd: '/w', parentSession: 'child' },
    };
    const resolve = (id) => resolveSessionScope({
        sessionId: id,
        header: headers[id],
        store,
        headerOf: headerOfFrom(headers),
    });

    // 链顶没设置：整条链 unset，绝不能当成「全部」。
    assert.equal((await resolve('grand')).kind, 'unset');

    await store.write('root', { mode: 'folder', folder: '资料' });
    assert.deepEqual((await resolve('child')).scope, { mode: 'folder', folder: '资料' });
    assert.deepEqual((await resolve('grand')).scope, { mode: 'folder', folder: '资料' });

    // 子会话收窄到子目录。
    await store.write('child', { mode: 'folder', folder: '资料/合同' });
    assert.deepEqual((await resolve('grand')).scope, { mode: 'folder', folder: '资料/合同' });

    // 子会话「全部」不能放大父范围：仍按父的 资料。
    await store.write('child', { mode: 'all' });
    assert.deepEqual((await resolve('grand')).scope, { mode: 'folder', folder: '资料' });

    // 子会话选到父范围外：交集为空 → 拒绝（empty）。
    await store.write('child', { mode: 'folder', folder: '其他' });
    assert.equal((await resolve('grand')).kind, 'empty');

    // 父放开全部后，子会话可以用自己的 folder 收窄。
    await store.write('root', { mode: 'all' });
    await store.write('child', { mode: 'folder', folder: '其他/子' });
    assert.deepEqual((await resolve('grand')).scope, { mode: 'folder', folder: '其他/子' });
    await store.write('child', { mode: 'all' });
    assert.deepEqual((await resolve('grand')).scope, { mode: 'all', folder: '' });
});

await check('继承：祖先 header 缺失 / 循环 / 超深都 fail-closed（不猜成根）', async () => {
    const { store } = await freshRoot();
    const resolve = (id, header, headerOf) => resolveSessionScope({ sessionId: id, header, store, headerOf });

    // 祖先 header 拿不到。
    const missing = await resolve('child', { id: 'child', parentSession: 'gone' }, async () => undefined);
    assert.equal(missing.kind, 'error');
    assert.equal(missing.code, 'ancestry');

    // 循环。
    const headers = {
        a: { id: 'a', parentSession: 'b' },
        b: { id: 'b', parentSession: 'a' },
    };
    const cycle = await resolve('a', headers.a, headerOfFrom(headers));
    assert.equal(cycle.kind, 'error');
    assert.equal(cycle.code, 'cycle');

    // 超深：链长 > MAX_ANCESTRY。
    const deep = {};
    for (let i = 0; i <= MAX_ANCESTRY + 1; i += 1) {
        deep['d' + i] = { id: 'd' + i, parentSession: i === MAX_ANCESTRY + 1 ? '' : 'd' + (i + 1) };
    }
    const tooDeep = await resolve('d0', deep.d0, headerOfFrom(deep));
    assert.equal(tooDeep.kind, 'error');
    assert.equal(tooDeep.code, 'depth');
});

await check('工具面：真会话缺状态 / 状态损坏 → 拒绝，不静默全库；无会话（内部）保持旧契约', async () => {
    const { root, home, store, config } = await freshRoot();
    const tools = buildTools(resolveConfig(config), ctxOf(home));
    const tool = tools.find((item) => item.name === 'office_memory');

    // 有会话但没设置：全部 kb-* 拒绝，文案指向输入框左侧。
    const unset = await tool.execute({ action: 'kb-list' }, execOf('s-unset', root));
    assert.equal(unset.ok, false);
    assert.match(unset.text, /还没有设置知识库范围/);
    assert.match(unset.text, /输入框左侧/);

    // 损坏：拒绝并说清是状态文件问题。
    const brokenId = 's-broken';
    const file = scopeFileOf(store.root, brokenId);
    await mkdir(store.root, { recursive: true });
    await writeFile(file, '{ broken', 'utf8');
    const corrupt = await tool.execute({ action: 'kb-list' }, execOf(brokenId, root));
    assert.equal(corrupt.ok, false);
    assert.match(corrupt.text, /读不懂/);

    // 有宿主 ctx、exec 却缺会话 id：不能当成内部调用放行（P1 修法）。
    const noId = await tool.execute({ action: 'kb-list' }, { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal });
    assert.equal(noId.ok, false);
    assert.match(noId.text, /知识库范围：无会话。/);

    // exec 给了 id 但没有 header / header id 与 session id 不符：同样无法确认，拒绝。
    const noHeader = await tool.execute({ action: 'kb-list' }, { agent: { session: { id: 's-no-header' } }, signal: new AbortController().signal });
    assert.equal(noHeader.ok, false);
    assert.match(noHeader.text, /知识库范围：无法确认。/);
    const mismatched = await tool.execute({ action: 'kb-list' }, execOf('s-real', root, { id: 's-other' }));
    assert.equal(mismatched.ok, false);
    assert.match(mismatched.text, /知识库范围：无法确认。/);

    // 没有宿主 ctx：内部路径（真会话的工具面一定绑着 ctx），保持旧契约。
    const noCtxTools = buildTools(resolveConfig(config));
    const noCtx = await noCtxTools.find((item) => item.name === 'office_memory')
        .execute({ action: 'kb-list' }, execOf('s-unset', root));
    assert.equal(noCtx.ok, true);

    // resolveCallScope 本身：未设置 → ok:false。
    const denied = await resolveCallScope({ ctx: ctxOf(home), exec: execOf('s-unset', root) });
    assert.equal(denied.ok, false);
});

await check('工具面：清单 / 树 / 检索 / 读取 / 删除 / 改搬 / 入库 / 连边 / 摘要都受范围约束', async () => {
    const { root, home, config, memory, store } = await freshRoot();
    const inDoc = join(root, 'docs', '范围里.md');
    const outDoc = join(root, 'docs', '范围外.md');
    await mkdir(join(root, 'docs'), { recursive: true });
    await writeFile(inDoc, '# 范围里\n\n这份文档讲合同里的付款条款。\n', 'utf8');
    await writeFile(outDoc, '# 范围外\n\n这份文档讲合同里的付款条款，但它在别的文件夹。\n', 'utf8');
    const inRow = await memory.kbIngest({ path: inDoc, folder: A, displayPath: 'docs/范围里.md' });
    const outRow = await memory.kbIngest({ path: outDoc, folder: B, displayPath: 'docs/范围外.md' });
    const inId = inRow.docId;
    const outId = outRow.docId;

    await store.write('s1', { mode: 'folder', folder: A });
    const tools = buildTools(resolveConfig(config), ctxOf(home));
    const tool = tools.find((item) => item.name === 'office_memory');
    const exec = execOf('s1', root);

    // 清单只列范围内的文档；树不牵出范围外的目录名。
    const list = await tool.execute({ action: 'kb-list' }, exec);
    assert.equal(list.ok, true);
    assert.match(list.text, /范围里/);
    assert.doesNotMatch(list.text, /范围外/);
    const tree = await tool.execute({ action: 'kb-tree' }, exec);
    assert.match(tree.text, /资料/);
    assert.doesNotMatch(tree.text, /其他/);

    // 检索：两文档都含同一关键词，只返回范围内的。
    const search = await tool.execute({ action: 'kb-search', query: '合同 付款 条款' }, exec);
    assert.equal(search.ok, true);
    assert.match(search.text, /范围里/);
    assert.doesNotMatch(search.text, /范围外/);

    // 读取：按 id / 按 path 都不能读范围外。
    // 读取按「不存在」处理（不回显范围外那份文档的目录名，避免确认它的存在与位置）。
    await assert.rejects(() => tool.execute({ action: 'kb-read', id: outId }, exec), /没有.*这份文档|不在当前会话/);
    await assert.rejects(() => tool.execute({ action: 'kb-read', path: 'docs/范围外.md' }, exec), /没有.*这份文档|不在当前会话/);

    // 检索时把 folder 写成范围外：直接拒绝，而不是「换个查询」。
    await assert.rejects(() => tool.execute({ action: 'kb-search', query: '合同', folder: B }, exec), /不在当前会话的知识库范围里/);
    // 边界：前缀相同的兄弟目录（资料夹）不算子目录。
    await assert.rejects(() => tool.execute({ action: 'kb-search', query: '合同', folder: '资料夹' }, exec), /不在当前会话的知识库范围里/);

    // 删除 / 改 / 搬 / 改正文：都不许碰范围外。
    await assert.rejects(() => tool.execute({ action: 'kb-drop', id: outId }, exec), /不在当前会话的知识库范围里/);
    await assert.rejects(() => tool.execute({ action: 'kb-update', id: outId, title: '改名' }, exec), /不在当前会话的知识库范围里/);
    await assert.rejects(() => tool.execute({ action: 'kb-update', id: inId, folder: B }, exec), /不在当前会话的知识库范围里/);
    await assert.rejects(() => tool.execute({ action: 'kb-edit', id: outId, text: '# 改' }, exec), /不在当前会话的知识库范围里/);

    // 入库到范围外：拒绝（不能把内容搬进范围里）。
    await assert.rejects(() => tool.execute({ action: 'kb-ingest', path: outDoc, folder: B }, exec), /不在当前会话的知识库范围里/);

    // 连边 / 摘要：两端 / 来源块都要在范围内。
    // 不带 scope 的读（工具之外的内部路径）拿到块 id，再拿它去试受限会话的连边 / 摘要。
    const outReadScoped = await memory.kbRead({ path: 'docs/范围外.md' });
    const outChunk = outReadScoped.chunks?.[0]?.id ?? '';
    const inReadScoped = await memory.kbRead({ path: 'docs/范围里.md' });
    const inChunk = inReadScoped.chunks?.[0]?.id ?? '';
    assert.ok(outChunk !== '' && inChunk !== '', '测试需要拿到两块 id');
    if (outChunk !== '' && inChunk !== '') {
        await assert.rejects(() => tool.execute({ action: 'kb-link', sourceId: inChunk, targetId: outChunk }, exec), /不在当前会话的知识库范围里/);
        await assert.rejects(() => tool.execute({ action: 'kb-summarize', text: '摘要', chunkIds: [outChunk] }, exec), /不在当前会话的知识库范围里/);
    }
});

await check('P1-1：范围外旧文档不能被同内容重入改档 / 搬家 / 回显', async () => {
    const { root, home, config, memory, store } = await freshRoot();
    const doc = join(root, 'same.md');
    await writeFile(doc, '# 同一份内容\n\n同内容重入不该越过会话范围。\n', 'utf8');
    const seeded = await memory.kbIngest({ path: 'same.md', folder: B, displayPath: 'same.md' });
    await store.write('s-p1', { mode: 'folder', folder: A });

    const tools = buildTools(resolveConfig(config), ctxOf(home));
    const tool = tools.find((item) => item.name === 'office_memory');
    const exec = execOf('s-p1', root);

    // folder 省略（null）：docId 按路径命中范围外那份旧文档，必须拒绝，不能 retier。
    await assert.rejects(
        () => tool.execute({ action: 'kb-ingest', path: 'same.md', tier: 'verified' }, exec),
        /不在当前会话的知识库范围里/,
    );
    // 显式指定范围内的 folder 也不行：那等于把范围外文档搬进来。
    await assert.rejects(
        () => tool.execute({ action: 'kb-ingest', path: 'same.md', folder: A }, exec),
        /不在当前会话的知识库范围里/,
    );
    // 直接走 memory（不经过工具）同样拒绝：folder:null 不是绕过闸门的入口。
    await assert.rejects(
        () => memory.kbIngest({ path: 'same.md', folder: null, sessionScope: { mode: 'folder', folder: A } }),
        /不在当前会话的知识库范围里/,
    );

    // 旧文档原样：文件夹仍是 B，档位没有被升级，标题 / 块数没被回显给受限会话。
    const after = (await memory.kbList({})).items.find((row) => row.id === seeded.docId);
    assert.equal(after.folder, B, '范围外旧文档的文件夹不能被改');
    assert.equal(after.tier, 'user', '范围外旧文档的档位不能被改');

    // 范围内文档的 kb-edit 仍然可用（内层重新入库带上 sessionScope 后不能误伤）。
    await writeFile(join(root, 'edit.md'), '# 改前\n\n付款条款。\n', 'utf8');
    const editable = await memory.kbIngest({ path: 'edit.md', folder: A, displayPath: 'edit.md' });
    const edit = await tool.execute({ action: 'kb-edit', id: editable.docId, text: '# 改后\n\n付款条款与验收。\n' }, exec);
    assert.equal(edit.ok, true);
    assert.match(edit.text, /已改并重新入库：改后/);
});

await check('P1-4：受限会话读不到来源块在范围外的摘要（读取 / 检索 / 清单同一条闸门）', async () => {
    const { root, home, config, memory, store } = await freshRoot();
    await writeFile(join(root, 'in.md'), '# 范围内\n\n付款条款与验收。\n', 'utf8');
    await writeFile(join(root, 'out.md'), '# 范围外\n\n海底电缆走向属于别的文件夹。\n', 'utf8');
    const inRow = await memory.kbIngest({ path: 'in.md', folder: A, displayPath: 'in.md' });
    await memory.kbIngest({ path: 'out.md', folder: B, displayPath: 'out.md' });

    // 未受限的旧会话可以生成跨文件夹摘要：这是历史数据，读侧必须自己守住。
    const outRead = await memory.kbRead({ path: 'out.md' });
    const outChunk = outRead.chunks?.[0]?.id ?? '';
    assert.ok(outChunk !== '', '需要范围外文档的块 id');
    const inRead = await memory.kbRead({ path: 'in.md' });
    const inChunk = inRead.chunks?.[0]?.id ?? '';
    assert.ok(inChunk !== '', '需要范围内文档的块 id');
    const secret = '跨文件夹摘要：海底电缆走向。';
    const cross = await memory.kbSummarize({ title: '跨文件夹摘要', text: secret, chunkIds: [outChunk], docId: inRow.docId });
    const own = await memory.kbSummarize({ title: '夹内摘要', text: '夹内摘要：付款条款。', chunkIds: [inChunk] });

    await store.write('s-sum', { mode: 'folder', folder: A });
    const tools = buildTools(resolveConfig(config), ctxOf(home));
    const tool = tools.find((item) => item.name === 'office_memory');
    const exec = execOf('s-sum', root);

    // 读取：摘要所属文档在范围内，但来源块在范围外 → 按「没有这份摘要」处理，不回显正文。
    await assert.rejects(() => tool.execute({ action: 'kb-read', id: 'summary:' + cross.id }, exec), /没有.*这份摘要|不在当前会话/);

    // 检索：摘要不能作为候选混进来（正文 / 预览都不行）。
    const search = await tool.execute({ action: 'kb-search', query: '海底电缆 摘要' }, exec);
    assert.equal(search.ok, true);
    assert.doesNotMatch(search.text, /海底电缆走向/);
    assert.doesNotMatch(search.text, /跨文件夹摘要/);

    // 清单：跨夹摘要的标题也不该列给受限会话。
    const list = await tool.execute({ action: 'kb-summaries' }, exec);
    assert.equal(list.ok, true);
    assert.match(list.text, /夹内摘要/);
    assert.doesNotMatch(list.text, /跨文件夹摘要/);

    // 来源块都在范围内的摘要照常可读（闸门不是「一律隐藏」）。
    const ownRead = await tool.execute({ action: 'kb-read', id: 'summary:' + own.id }, exec);
    assert.equal(ownRead.ok, true);
    assert.match(ownRead.text, /付款条款/);
});

await check('工具面：范围「全部」与无限制老路径行为一致；范围内子目录可见', async () => {
    const { root, home, config, memory, store } = await freshRoot();
    const docA = join(root, 'a.md');
    const docB = join(root, 'b.md');
    const docC = join(root, 'c.md');
    await writeFile(docA, '# A\n\n甲文件夹的文档。\n', 'utf8');
    await writeFile(docB, '# B\n\n甲/子文件夹的文档。\n', 'utf8');
    await writeFile(docC, '# C\n\n乙文件夹的文档。\n', 'utf8');
    await memory.kbIngest({ path: docA, folder: '甲', displayPath: 'a.md' });
    await memory.kbIngest({ path: docB, folder: '甲/子', displayPath: 'b.md' });
    await memory.kbIngest({ path: docC, folder: '乙', displayPath: 'c.md' });

    const tools = buildTools(resolveConfig(config), ctxOf(home));
    const tool = tools.find((item) => item.name === 'office_memory');

    await store.write('all-s', { mode: 'all' });
    const allList = await tool.execute({ action: 'kb-list' }, execOf('all-s', root));
    assert.equal(allList.ok, true);
    assert.match(allList.text, /a\.md|甲/);
    assert.match(allList.text, /c\.md|乙/);

    await store.write('sub-s', { mode: 'folder', folder: '甲' });
    const subList = await tool.execute({ action: 'kb-list' }, execOf('sub-s', root));
    assert.match(subList.text, /a\.md/);
    assert.match(subList.text, /b\.md/, '子目录里的文档要能看见');
    assert.doesNotMatch(subList.text, /c\.md/);
});

await check('/kb-scope 命令：查看 / 设置 / 清除；不存在目录当场拒；子代理拒绝写', async () => {
    const { root, home, store } = await freshRoot();
    let definition = null;
    const ctx = {
        get: (name) => {
            if (name === 'profileContext') return { name: 'test', home, cwd: home };
            if (name === 'commands') return { register: (def) => { definition = def; return () => {}; } };
            return null;
        },
    };
    const dispose = registerKbScopeCommand(ctx, {
        folderExists: async (_session, folder) => folder === '资料',
    });
    assert.equal(typeof dispose, 'function');
    assert.equal(definition.name, 'kb-scope');

    const call = (rawInput, session = { id: 's1', header: { id: 's1', cwd: root } }) => definition.handler({ rawInput, agent: { session } });

    const status = await call('');
    assert.equal(status.kind, 'success');
    assert.match(status.text, /^知识库范围：未设置。/);
    assert.match(status.text, /还没有设置/);

    const set = await call('set 资料');
    assert.equal(set.kind, 'success');
    assert.match(set.text, /^知识库范围：已设为「资料」/);
    assert.deepEqual((await store.read('s1')).record.folder, '资料');

    const status2 = await call('');
    assert.match(status2.text, /^知识库范围：「资料」/);

    const missing = await call('set 不存在');
    assert.equal(missing.kind, 'error');
    assert.match(missing.text, /没有/);

    const bad = await call('set');
    assert.equal(bad.kind, 'error');
    const badRoot = await call('set /');
    assert.equal(badRoot.kind, 'error');

    const clear = await call('clear');
    assert.equal(clear.kind, 'success');
    assert.match(clear.text, /^知识库范围：已清空（全部知识库）/);
    assert.equal((await store.read('s1')).record.mode, 'all');

    // 子代理继承父范围，不能自己写。
    const subagent = await call('set 资料', { id: 'sub', header: { id: 'sub', origin: 'subagent', parentSession: 's1', cwd: root } });
    assert.equal(subagent.kind, 'error');
    assert.match(subagent.text, /子代理/);
    assert.equal((await store.read('sub')).status, 'unset');

    // 没有会话上下文：拒绝。
    const noSession = await call('', null);
    assert.equal(noSession.kind, 'error');
    assert.match(noSession.text, /知识库范围：无会话。/);

    // 有 session id 但没有 header：不能写（P1 修法）。
    const noHeader = await call('set 资料', { id: 's-no-header', header: null });
    assert.equal(noHeader.kind, 'error');
    assert.match(noHeader.text, /知识库范围：无法确认。/);
    assert.equal((await store.read('s-no-header')).status, 'unset');

    // header id 与 session id 不符：同样拒绝，不按一个「看起来像」的 id 写。
    const mismatch = await call('set 资料', { id: 's-mismatch', header: { id: 's-other', cwd: root } });
    assert.equal(mismatch.kind, 'error');
    assert.match(mismatch.text, /知识库范围：无法确认。/);
    assert.equal((await store.read('s-mismatch')).status, 'unset');
});

await check('P1-2：无父记录的 seeded fork 不能当根、不能 set/clear；有父的 fork 正常继承', async () => {
    const { root, home, store } = await freshRoot();
    await store.write('root', { mode: 'folder', folder: A });
    // 真实普通 UI fork：header 带 isSeeded:true，但没有 parentSession。
    await store.write('fork-seeded', { mode: 'all' });

    const orphan = await resolveSessionScope({
        sessionId: 'fork-seeded',
        header: { id: 'fork-seeded', isSeeded: true, cwd: root },
        store,
        headerOf: async () => undefined,
    });
    assert.equal(orphan.kind, 'error');
    assert.equal(orphan.code, 'origin');

    // 工具面：不因为自己的记录写了 all 就拿到全库。
    const viaCall = await resolveCallScope({
        ctx: ctxOf(home),
        exec: execOf('fork-seeded', root, { isSeeded: true }),
    });
    assert.equal(viaCall.ok, false);
    assert.match(viaCall.text, /^知识库范围：来源不可确认。/);

    // 命令面：set / clear 都拒绝，且原记录不被改写。
    let definition = null;
    const ctx = {
        get: (name) => {
            if (name === 'profileContext') return { name: 'test', home, cwd: home };
            if (name === 'commands') return { register: (def) => { definition = def; return () => {}; } };
            return null;
        },
    };
    registerKbScopeCommand(ctx, { folderExists: async () => true });
    const call = (rawInput, header) => definition.handler({ rawInput, agent: { session: { id: 'fork-seeded', header } } });
    const seededHeader = { id: 'fork-seeded', isSeeded: true, cwd: root };
    const setSeeded = await call('set 资料', seededHeader);
    assert.equal(setSeeded.kind, 'error');
    assert.match(setSeeded.text, /^知识库范围：来源不可确认。/);
    const clearSeeded = await call('clear', seededHeader);
    assert.equal(clearSeeded.kind, 'error');
    assert.equal((await store.read('fork-seeded')).record.mode, 'all', '拒绝后原记录不能变');

    // 有 parentSession 的 seeded fork：照常沿父链继承（不因 isSeeded 一概拒绝）。
    const withParent = await resolveSessionScope({
        sessionId: 'fork-child',
        header: { id: 'fork-child', isSeeded: true, parentSession: 'root', cwd: root },
        store,
        headerOf: headerOfFrom({ root: { id: 'root', cwd: root } }),
    });
    assert.deepEqual(withParent.scope, { mode: 'folder', folder: A });
});

await check('生产接缝：真实 kbTree 的 {root} 包装能被 folderExists 的 helper 识别（exact 边界）', async () => {
    const { root, memory } = await freshRoot();
    const doc = join(root, 'tree.md');
    await writeFile(doc, '# 树\n\n真实树的包装形态。\n', 'utf8');
    await memory.kbIngest({ path: doc, folder: '资料/合同', displayPath: 'tree.md' });

    const tree = await memory.kbTree({});
    // 真实返回形态：{ action, root, counts, origin }。不拆 root 就会把包装当节点（恒 false）。
    assert.equal(tree.action, 'kb-tree');
    assert.ok(tree.root !== null && typeof tree.root === 'object');
    assert.equal(kbTreeHasFolder(tree, '资料'), true);
    assert.equal(kbTreeHasFolder(tree, '资料/合同'), true);
    assert.equal(kbTreeHasFolder(tree.root, '资料/合同'), true, '裸节点也要接受');
    assert.equal(kbTreeHasFolder([tree.root], '资料'), true, '数组形态也要接受');
    assert.equal(kbTreeHasFolder(tree, '资料夹'), false, '前缀相同的兄弟目录不算命中');
    assert.equal(kbTreeHasFolder(tree, '不存在'), false);
    assert.equal(kbTreeHasFolder(tree, ''), true, '根 = 全部');
});

await check('生产接缝：folders.json 读坏时 manifest 里确认的目录仍通过，查不到返回 null（重试）', async () => {
    const tree = {
        action: 'kb-tree',
        root: { path: '', children: [{ path: '资料', children: [] }] },
    };
    assert.equal(kbTreeFolderPresence(tree, '资料'), true);
    assert.equal(kbTreeFolderPresence(tree, '资料夹'), false);
    assert.equal(kbTreeFolderPresence(tree.root, '资料'), true);
    const broken = { ...tree, foldersError: 'KB_FOLDERS_CORRUPT: folders.json 读不懂' };
    assert.equal(kbTreeFolderPresence(broken, '资料'), true, 'manifest 能确认的目录不受坏清单影响');
    assert.equal(kbTreeFolderPresence(broken, '项目/空'), null, '查不到时不能断言不存在');
});

await check('/kb-scope：父会话未设置时子会话 set/clear 被拒，不产生假成功；父设置后可收窄', async () => {
    const { root, home, store } = await freshRoot();
    let definition = null;
    const headers = new Map([['root', { id: 'root', cwd: root }]]);
    const ctx = {
        get: (name) => {
            if (name === 'profileContext') return { name: 'test', home, cwd: home };
            if (name === 'commands') return { register: (def) => { definition = def; return () => {}; } };
            if (name === 'agents') {
                return {
                    get: (id) => (headers.has(id) ? { session: { header: headers.get(id) } } : null),
                };
            }
            return null;
        },
    };
    registerKbScopeCommand(ctx, { folderExists: async () => true });
    const call = (rawInput, header) => definition.handler({ rawInput, agent: { session: { id: header.id, header } } });
    const childHeader = { id: 'child', parentSession: 'root', cwd: root };

    // 父会话还没有记录：子会话写自己的记录不会生效（解析以链顶为准），必须拒绝而不是回「已设为」。
    const setBefore = await call('set 资料', childHeader);
    assert.equal(setBefore.kind, 'error');
    assert.match(setBefore.text, /^知识库范围：未设置。/);
    assert.equal((await store.read('child')).status, 'unset', '拒绝后不能落盘');
    const clearBefore = await call('clear', childHeader);
    assert.equal(clearBefore.kind, 'error');
    assert.equal((await store.read('child')).status, 'unset');

    // 父会话明确设置后再来收窄：写入生效，且生效范围是交集。
    const parentSet = await call('set 资料', { id: 'root', cwd: root });
    assert.equal(parentSet.kind, 'success');
    const childSet = await call('set 资料/合同', childHeader);
    assert.equal(childSet.kind, 'success');
    assert.equal((await store.read('child')).record.folder, '资料/合同');
    const effective = await resolveSessionScope({
        sessionId: 'child',
        header: childHeader,
        store,
        headerOf: headerOfFrom({ root: { id: 'root', cwd: root } }),
    });
    assert.deepEqual(effective.scope, { mode: 'folder', folder: '资料/合同' });
});

await check('复核：本地空范围 / 损坏记录可由一次新选择修复（门槛只看父会话）', async () => {
    const { root, home, store } = await freshRoot();
    await store.write('root', { mode: 'folder', folder: A });
    // 本会话记录与父范围不交集：effective 为空，但父会话是好的，应该允许重新选择修复。
    await store.write('child-empty', { mode: 'folder', folder: B });
    // 本会话记录损坏：同样不该被自己的坏文件锁死。
    await mkdir(store.root, { recursive: true });
    await writeFile(scopeFileOf(store.root, 'child-corrupt'), '{ broken json', 'utf8');

    const h = commandHarness({
        home,
        headers: { root: { id: 'root', cwd: root } },
        folderExists: async () => true,
    });
    const emptyHeader = { id: 'child-empty', parentSession: 'root', cwd: root };
    const before = await h.call('', emptyHeader);
    assert.match(before.text, /范围为空/);
    const repaired = await h.call('set 资料/合同', emptyHeader);
    assert.equal(repaired.kind, 'success');
    assert.match(repaired.text, /已设为「资料\/合同」/);
    assert.equal((await store.read('child-empty')).record.folder, '资料/合同');

    const corruptHeader = { id: 'child-corrupt', parentSession: 'root', cwd: root };
    const beforeCorrupt = await h.call('', corruptHeader);
    assert.match(beforeCorrupt.text, /状态损坏/);
    const repairCorrupt = await h.call('set 资料', corruptHeader);
    assert.equal(repairCorrupt.kind, 'success');
    assert.equal((await store.read('child-corrupt')).record.folder, A);
});

await check('复核：越出父范围的 set 写前拒绝，原记录与当前生效范围都不动', async () => {
    const { root, home, store } = await freshRoot();
    await store.write('root', { mode: 'folder', folder: A });
    await store.write('child', { mode: 'folder', folder: '资料/合同' });
    const h = commandHarness({
        home,
        headers: { root: { id: 'root', cwd: root } },
        folderExists: async () => true,
    });
    const header = { id: 'child', parentSession: 'root', cwd: root };

    const disjoint = await h.call('set 其他', header);
    assert.equal(disjoint.kind, 'error');
    assert.match(disjoint.text, /不在父会话的范围「资料」里/);
    assert.match(disjoint.text, /没有保存/);
    assert.equal((await store.read('child')).record.folder, '资料/合同', '拒绝时不能改原记录');

    // 前缀相同的兄弟目录同样是范围外（exact / descendant 边界）。
    const sibling = await h.call('set 资料夹', header);
    assert.equal(sibling.kind, 'error');
    assert.match(sibling.text, /不在父会话的范围「资料」里/);
    assert.equal((await store.read('child')).record.folder, '资料/合同');

    const effective = await resolveSessionScope({
        sessionId: 'child',
        header,
        store,
        headerOf: headerOfFrom({ root: { id: 'root', cwd: root } }),
    });
    assert.deepEqual(effective.scope, { mode: 'folder', folder: '资料/合同' });

    // 父范围里的子目录仍可正常收窄。
    const narrowed = await h.call('set 资料/发票', header);
    assert.equal(narrowed.kind, 'success');
    assert.equal((await store.read('child')).record.folder, '资料/发票');
});

await check('复核：派生会话 clear 的回执按父范围说，不谎报「全部知识库」', async () => {
    const { root, home, store } = await freshRoot();
    await store.write('root', { mode: 'folder', folder: A });
    await store.write('root-all', { mode: 'all' });
    const headers = {
        root: { id: 'root', cwd: root },
        'root-all': { id: 'root-all', cwd: root },
    };
    const h = commandHarness({ home, headers, folderExists: async () => true });
    const child = { id: 'child', parentSession: 'root', cwd: root };

    const cleared = await h.call('clear', child);
    assert.equal(cleared.kind, 'success');
    assert.match(cleared.text, /已清除本会话的附加限制/);
    assert.match(cleared.text, /「资料」/);
    assert.doesNotMatch(cleared.text, /全部知识库/);
    assert.equal((await store.read('child')).record.mode, 'all');
    const effective = await resolveSessionScope({
        sessionId: 'child',
        header: child,
        store,
        headerOf: headerOfFrom({ root: { id: 'root', cwd: root } }),
    });
    assert.deepEqual(effective.scope, { mode: 'folder', folder: A });

    // 父会话本身是全部时，clear 回执才可以说「已清空（全部知识库）」。
    const childAll = { id: 'child-all', parentSession: 'root-all', cwd: root };
    const clearedAll = await h.call('clear', childAll);
    assert.equal(clearedAll.kind, 'success');
    assert.match(clearedAll.text, /已清空（全部知识库）/);
});

await check('复核：folderExists 读失败（null / 抛出）要求重试，不写成「已设为」', async () => {
    const { root, home, store } = await freshRoot();
    const header = { id: 'ui-root', cwd: root };
    const cases = [
        { name: 'null', folderExists: async () => null },
        { name: 'throw', folderExists: async () => { throw new Error('kb tree unavailable'); } },
    ];
    for (const item of cases) {
        const h = commandHarness({ home, folderExists: item.folderExists });
        const result = await h.call('set 资料', header);
        assert.equal(result.kind, 'error', item.name);
        assert.match(result.text, /暂时核对不了「资料」/);
        assert.match(result.text, /重试/);
        assert.equal((await store.read('ui-root')).status, 'unset', item.name + ' 不能落盘');
    }
    // 明确不存在与「读失败」要分开：前者说没有这个目录，后者说稍后重试。
    const missing = commandHarness({ home, folderExists: async () => false });
    const notFound = await missing.call('set 资料', header);
    assert.equal(notFound.kind, 'error');
    assert.match(notFound.text, /没有「资料」/);
    assert.equal((await store.read('ui-root')).status, 'unset');
});

await check('复核：父会话损坏 / 父 header 缺失时，子会话不能写（父门槛 fail-closed）', async () => {
    const { root, home, store } = await freshRoot();
    await mkdir(store.root, { recursive: true });
    await writeFile(scopeFileOf(store.root, 'root-corrupt'), '{ broken json', 'utf8');
    const h = commandHarness({
        home,
        headers: { 'root-corrupt': { id: 'root-corrupt', cwd: root } },
        folderExists: async () => true,
    });

    const corruptChild = { id: 'child-corrupt-parent', parentSession: 'root-corrupt', cwd: root };
    const corrupt = await h.call('set 资料', corruptChild);
    assert.equal(corrupt.kind, 'error');
    assert.match(corrupt.text, /^知识库范围：状态损坏。/);
    assert.equal((await store.read('child-corrupt-parent')).status, 'unset');

    // 父 header 已不可用：不能假定它是根，也不能写。
    const missingChild = { id: 'child-missing-parent', parentSession: 'gone', cwd: root };
    const missing = await h.call('set 资料', missingChild);
    assert.equal(missing.kind, 'error');
    assert.match(missing.text, /^知识库范围：父会话不可用。/);
    assert.equal((await store.read('child-missing-parent')).status, 'unset');
});

await check('P1 复核：受限会话里没有来源块 id 的旧摘要不可见，全库会话照常可读', async () => {
    const { root, home, config, memory, store } = await freshRoot();
    const doc = join(root, 'empty-src.md');
    await writeFile(doc, '# 旧摘要\n\n旧摘要正文：没有来源块。\n', 'utf8');
    await memory.kbIngest({ path: doc, folder: A, displayPath: 'empty-src.md' });
    const read = await memory.kbRead({ path: 'empty-src.md' });
    const chunk = read.chunks?.[0]?.id ?? '';
    assert.ok(chunk !== '', '需要范围内文档的块 id');
    const summary = await memory.kbSummarize({ title: '旧摘要', text: '旧摘要正文：没有来源块。', chunkIds: [chunk] });

    // 模拟旧 / 外部写入的摘要：来源块 id 丢失为空数组。受限会话不能因此 fail-open。
    let sqlite = null;
    try {
        sqlite = await import('node:sqlite');
    } catch {
        sqlite = null;
    }
    if (sqlite === null) {
        console.log('      （本机没有 node:sqlite，跳过空来源摘要这一项）');
        return;
    }
    const db = new sqlite.DatabaseSync(kbDbPath(config.memory.kbDir));
    try {
        db.prepare('UPDATE summaries SET chunk_ids = ? WHERE id = ?').run('[]', summary.id);
    } finally {
        db.close();
    }

    await store.write('s-empty-src', { mode: 'folder', folder: A });
    const tool = buildTools(resolveConfig(config), ctxOf(home)).find((item) => item.name === 'office_memory');
    await assert.rejects(
        () => tool.execute({ action: 'kb-read', id: 'summary:' + summary.id }, execOf('s-empty-src', root)),
        /没有.*这份摘要|不在当前会话/,
    );
    const list = await tool.execute({ action: 'kb-summaries' }, execOf('s-empty-src', root));
    assert.equal(list.ok, true);
    assert.doesNotMatch(list.text, /旧摘要/);

    // 全库会话不受这条收紧影响（不是全局隐藏摘要）。
    await store.write('s-all-src', { mode: 'all' });
    const allRead = await tool.execute({ action: 'kb-read', id: 'summary:' + summary.id }, execOf('s-all-src', root));
    assert.equal(allRead.ok, true);
    assert.match(allRead.text, /旧摘要正文/);
});

await check('状态契约：状态 / 拒绝文案都有「知识库范围：」前缀与稳定 token', async () => {
    assert.match(scopeRefusalText({ kind: 'unset' }), /^知识库范围：未设置。/);
    assert.match(scopeRefusalText({ kind: 'corrupt' }), /^知识库范围：状态损坏。/);
    assert.match(scopeRefusalText({ kind: 'empty' }), /^知识库范围：范围为空。/);
    assert.match(scopeRefusalText({ kind: 'no-session' }), /^知识库范围：无会话。/);
    assert.match(scopeRefusalText({ kind: 'error', code: 'ancestry' }), /^知识库范围：父会话不可用。/);
    assert.match(scopeRefusalText({ kind: 'error', code: 'origin' }), /^知识库范围：来源不可确认。/);
    assert.match(scopeRefusalText({ kind: 'error', code: 'header' }), /^知识库范围：无法确认。/);
    assert.match(scopeRefusalText({ kind: 'error', code: 'cycle' }), /^知识库范围：父链循环。/);
    assert.match(scopeRefusalText({ kind: 'error', code: 'depth' }), /^知识库范围：父链过深。/);
    assert.match(scopeRefusalText({ kind: 'error' }), /^知识库范围：状态读取失败。/);

    // 客户端从回执里认状态的解析与这里的 token 对齐：
    // 只有「全部知识库」或「<folder>」才算读到一个范围；未设置 / 拒绝不能被误读成某个目录。
    const scopeOfText = (text) => {
        const value = typeof text === 'string' ? text : '';
        if (value.indexOf('全部知识库') !== -1) return '全部知识库';
        const quoted = /「([^」]+)」/.exec(value);
        return quoted === null ? '' : quoted[1];
    };
    assert.equal(scopeOfText(scopeRefusalText({ kind: 'unset' })), '', '未设置不能被解析出范围');
    assert.equal(scopeOfText(scopeRefusalText({ kind: 'no-session' })), '');
    assert.equal(scopeOfText(scopeRefusalText({ kind: 'corrupt' })), '');
    assert.equal(scopeOfText(scopeRefusalText({ kind: 'error', code: 'header' })), '');
    assert.equal(scopeOfText(scopeRefusalText({ kind: 'empty' })), '');
    assert.equal(scopeOfText('知识库范围：全部知识库。'), '全部知识库');
    assert.equal(scopeOfText('知识库范围：「资料」（含子目录）。'), '资料');
    // 普通真实会话第一次打开：状态行就是「未设置」，入口保持通用文案而不是显示一个假目录。
    assert.equal(scopeOfText('知识库范围：未设置。这个会话还没有设置知识库范围，所以不能查知识库（不会默认放开全库）。'
        + '在输入框左侧选一个虚拟文件夹，或输入 /kb-scope set <虚拟文件夹>；要查全部用 /kb-scope clear。'), '');
});

await check('真实普通会话：isSeeded:false 无父 header 可查看/设置，且工具面读到同一范围', async () => {
    const { root, home, store } = await freshRoot();
    await store.write('资料', { mode: 'folder', folder: '资料' });
    let definition = null;
    const seenSessions = [];
    const ctx = {
        get: (name) => {
            if (name === 'profileContext') return { name: 'test', home, cwd: home };
            if (name === 'commands') return { register: (def) => { definition = def; return () => {}; } };
            return null;
        },
    };
    registerKbScopeCommand(ctx, {
        folderExists: async (session, folder) => {
            seenSessions.push({ session, folder });
            return folder === '资料';
        },
    });
    // 与持久层真实普通根会话一致：id / cwd / isSeeded:false / delegationDepth:0，没有 parentSession。
    const ordinaryHeader = { id: 'ui-root', cwd: root, isSeeded: false, delegationDepth: 0 };
    const session = { id: 'ui-root', header: ordinaryHeader };
    const call = (rawInput) => definition.handler({
        commandId: 'cmd-fixture-1',
        agent: { session },
        rawInput,
        attachments: [],
        signal: new AbortController().signal,
    });

    const status = await call('');
    assert.equal(status.kind, 'success');
    assert.match(status.text, /^知识库范围：未设置。/);
    assert.equal((await store.read('ui-root')).status, 'unset');

    const set = await call(' set 资料 ');
    assert.equal(set.kind, 'success');
    assert.match(set.text, /^知识库范围：已设为「资料」/);
    // 核对目录时拿到的必须是宿主给的会话身份，而不是命令文本里可伪造的东西。
    assert.equal(seenSessions.length, 1);
    assert.equal(seenSessions[0].session.id, 'ui-root');
    assert.equal(seenSessions[0].session.header, ordinaryHeader);
    assert.equal(seenSessions[0].folder, '资料');

    const readBack = await call('');
    assert.match(readBack.text, /^知识库范围：「资料」（含子目录）。/);

    // 同一条会话的工具面读的是同一份状态（命令写、工具读，不各算各的）。
    const config = resolveConfig({ memory: { kbDir: join(root, '.office', 'kb') } });
    const tools = buildTools(config, ctx);
    const tool = tools.find((item) => item.name === 'office_memory');
    const listed = await tool.execute({ action: 'kb-list' }, execOf('ui-root', root, { isSeeded: false, delegationDepth: 0 }));
    assert.equal(listed.ok, true);
    // 目录是空的：范围本身仍然生效，只是没有可列出的文档。
    assert.match(listed.text, /没有|空|0/);
});

let failed = 0;
for (const result of results) {
    if (result.ok) {
        console.log('PASS  ' + result.name);
    } else {
        failed += 1;
        console.log('FAIL  ' + result.name + '  → ' + String(result.error?.message ?? result.error));
        if (result.error?.stack) console.log(String(result.error.stack).split('\n').slice(1, 4).join('\n'));
    }
}
console.log('\nkb-scope: ' + (results.length - failed) + '/' + results.length + ' 通过');
if (failed > 0) process.exitCode = 1;
