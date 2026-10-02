/**
 * 知识库（kb）的测试 —— 第四棵树：外部原文的结构化切块与可回溯指针（第四十一轮）。
 *
 * 这一份测的是 kb 自己的契约与它和三层记忆的接缝：
 *   - 切块：结构切块（标题路径 + 段落）确定性、span 可回溯（`source.slice(span) === text`）、
 *     块头拼「标题路径 · 来源 · 日期」；
 *   - 存储：manifest 是清单、块文件是正文、index.json 只是计数投影；
 *   - 入库：同路径同内容幂等（一个字节不写）、同路径新内容换代（旧块删、边清、墓碑留）；
 *   - 读取：按文档 id / 块 id / path，有界（4000 字符、12 块），缺块如实报；
 *   - 检索（第二十四轮 24-11）：词法多路（正文 + 块头）BM25，命中给块 id / span / 预览，
 *     词表零重叠时弃答，来源档与路径可筛，次数由每回合配额单独管；
 *   - 删除：正文删掉 + 同事务清边 + 留墓碑，且「悬空引用 == 0」照旧成立；
 *   - 写入期硬规则质量门：unverified（单源）永不进热记忆，引用 unverified 块同样被拒；
 *   - 引用（usedRefs）：`kb:` 形的 source 另存成 kbRefs，台账与记忆都能反着查；
 *   - 冲突三态：contradicts 边带类别与状态，保留双方、不做自动裁决；
 *   - Pack：整包搬走 kb（清单 + 块），导入幂等、被墓碑判死的不复。
 *
 * 跑法：node test/kb.mjs
 */
import { readFileSync } from 'node:fs';
import { mkdir, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import {
    KB_MAX_DOC_BYTES,
    KB_READ_CHARS,
    chunkHeaderOf,
    chunkIdOf,
    docIdOf,
    splitDocument,
} from '../src/kb.js';
import { createMemory, memoryPaths, renderKb, renderMutation, renderRead } from '../src/memory.js';
import { buildTools } from '../src/tools.js';

const TMP = fileURLToPath(new URL('../../../.office/tmp/kb/', import.meta.url));

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 每个用例一个干净的工作目录：kb 是按工作目录存的，串了目录就串了结论。 */
let caseId = 0;
async function freshRoot(memoryConfig = {}) {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    await rm(root, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    return { root, memory: createMemory({ root, memory: memoryConfig }) };
}

/** 写一份夹具文档并返回它的相对路径。 */
async function writeDoc(root, rel, text) {
    const target = join(root, rel);
    await mkdir(join(target, '..'), { recursive: true });
    await writeFile(target, text, 'utf8');
    return rel.split('\\').join('/');
}

const DOC = [
    '# 季度汇报要点',
    '',
    '第一段：营收同比增长 12%，主要来自华东区。',
    '第二段：成本持平，毛利率提高 1.4 个点。',
    '',
    '## 风险',
    '',
    '供应链价格波动是最大不确定项。',
    '',
    '## 下季度动作',
    '',
    '一、把华东区的渠道复盘做一次；二、把毛利口径与财务对齐。',
    '',
    '## 附注',
    '',
    '数据口径与上一版一致，未做重述。',
].join('\n');

// ── 1. 切块：确定性、span 可回溯、块头 ──────────────────────────────────────

await check('切块：结构切块确定性，span 与正文逐字相等（source.slice(span) === text）', () => {
    const source = `${DOC}\n\n## 附录\n\n附注正文。`;
    const first = splitDocument(source, { chunkChars: 400 });
    const again = splitDocument(source, { chunkChars: 400 });
    assert.deepEqual(again, first, '同一份正文两次切块必须完全一样（同 id、同 span、同正文）');
    assert.ok(first.length >= 3, `这份夹具该切出至少三块，实际 ${first.length}`);
    assert.deepEqual(first.map((chunk) => chunk.n), first.map((unused, index) => index + 1), '块号从 1 连续');
    for (const chunk of first) {
        assert.equal(source.slice(chunk.span.start, chunk.span.end), chunk.text,
            `第 ${chunk.n} 块的 span 必须能原样切回正文`);
        assert.ok(chunk.text.trim() !== '', '块正文不能是空白');
    }
    assert.deepEqual(first.map((chunk) => chunk.outline),
        ['季度汇报要点', '季度汇报要点 / 风险', '季度汇报要点 / 下季度动作', '季度汇报要点 / 附注', '季度汇报要点 / 附录'],
        '标题路径要按层级拼出来，并跟着块走');
});

await check('切块：单段超限时按可断点硬切，仍可回溯且不放任一块爆掉', () => {
    const long = `${'甲'.repeat(300)}。\n\n${'乙'.repeat(300)}。`;
    const chunks = splitDocument(long, { chunkChars: 200 });
    assert.ok(chunks.length >= 3, `300 字的段落按 200 字上限该切成多块，实际 ${chunks.length}`);
    for (const chunk of chunks) {
        assert.equal(long.slice(chunk.span.start, chunk.span.end), chunk.text);
        assert.ok(chunk.text.length <= 200 * 2, `硬切的单块不该失控：${chunk.text.length}`);
    }
});

await check('切块：同一标题下正文超限被切开时，每一块都还带着标题路径', () => {
    // 这一条是探针挖出来的：按额度收口之后要**按当前标题路径重新开块**，
    // 否则新块会丢掉标题路径（块头因此少一半，Contextual Retrieval 的收益就没了）。
    const body = Array.from({ length: 5 }, (unused, index) => `第 ${index + 1} 段：${'内容'.repeat(60)}。`).join('\n\n');
    const source = `# 长章节\n\n${body}\n\n## 结尾\n\n末段。`;
    const chunks = splitDocument(source, { chunkChars: 300 });
    assert.ok(chunks.length >= 4, `这份正文该被切成多块，实际 ${chunks.length}`);
    for (const chunk of chunks) {
        assert.equal(source.slice(chunk.span.start, chunk.span.end), chunk.text);
    }
    for (const chunk of chunks.slice(0, -1)) {
        assert.equal(chunk.outline, '长章节', `第 ${chunk.n} 块丢了标题路径：${JSON.stringify(chunk.outline)}`);
    }
    assert.equal(chunks[chunks.length - 1].outline, '长章节 / 结尾');
});

await check('块头：标题路径 + 来源 + 日期，且不混进块正文', () => {
    const header = chunkHeaderOf({ outline: 'A / B', source: '.office/search/x.md', at: '2026-09-30T10:00:00.000Z' });
    assert.equal(header, 'A / B · 来源 .office/search/x.md · 2026-09-30');
    assert.equal(chunkHeaderOf({ outline: '', source: 'a.md', at: '' }), '来源 a.md');
    const chunks = splitDocument(DOC, { chunkChars: 400 });
    assert.ok(!chunks[0].text.includes('来源 '), '块头是渲染时前置的，不该混进 text（否则 span 对不上）');
});

// ── 2. 入库：落盘、幂等、换代 ───────────────────────────────────────────────

await check('入库：manifest / chunks / index 都落盘，读得回来', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, '资料/季度汇报.md', DOC);
    const value = await memory.kbIngest({ path: rel });
    assert.equal(value.unchanged, false);
    assert.equal(value.replaced, false);
    assert.ok(value.chunks >= 3, `该切出多块，实际 ${value.chunks}`);
    assert.match(value.docId, /^kb:[0-9a-f]{12}$/);
    const paths = memoryPaths(root, {}).kb;
    const manifest = JSON.parse((await readFile(paths.manifest, 'utf8')).trim().split('\n')[0]);
    assert.equal(manifest.path, rel);
    assert.equal(manifest.chunks, value.chunks);
    assert.equal(manifest.tier, 'user');
    const index = JSON.parse(await readFile(paths.index, 'utf8'));
    assert.deepEqual(index.counts, { docs: 1, chunks: value.chunks, bytes: manifest.bytes, tiers: { user: 1, verified: 0, unverified: 0 } });
    const chunk = JSON.parse(await readFile(join(paths.chunks, `${manifest.hash}`, '1.json'), 'utf8'));
    assert.equal(chunk.id, chunkIdOf(manifest.hash, 1));
    assert.equal(DOC.slice(chunk.span.start, chunk.span.end), chunk.text);
    const listed = await memory.kbList();
    assert.equal(listed.total, 1);
    assert.equal(listed.items[0].id, docIdOf(rel));
    const read = await memory.kbRead({ id: value.docId });
    assert.ok(read.chunks.length >= 3);
    assert.ok(read.chunks[0].text.includes('第一段'));
    assert.equal(read.truncated, false);
});

await check('入库幂等：同路径同内容再入库不写任何字节（unchanged）', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'a.md', DOC);
    const first = await memory.kbIngest({ path: rel });
    const hash = await contentHashOf(memory, rel);
    const paths = memoryPaths(root, {}).kb;
    const before = {
        manifest: await readFile(paths.manifest, 'utf8'),
        index: await readFile(paths.index, 'utf8'),
        chunk: await readFile(join(paths.chunks, hash, '1.json'), 'utf8'),
    };
    const second = await memory.kbIngest({ path: rel });
    assert.equal(second.unchanged, true);
    assert.equal(second.chunks, first.chunks);
    assert.equal(await readFile(paths.manifest, 'utf8'), before.manifest, 'manifest 不该被重写');
    assert.equal(await readFile(paths.index, 'utf8'), before.index, 'index 不该被重写');
    assert.equal(await readFile(join(paths.chunks, hash, '1.json'), 'utf8'), before.chunk, '块文件不该被重写');
});

await check('入库换代：同路径新内容删旧块、清旧边、留墓碑（悬空引用 == 0）', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'b.md', DOC);
    const first = await memory.kbIngest({ path: rel });
    const hashOld = await contentHashOf(memory, rel);
    const oldChunkIds = (await memory.kbRead({ id: first.docId })).chunks.map((chunk) => chunk.id);
    // 用旧块 id 建一条边：换代时这条边必须被同事务清掉。
    await memory.mutate({ action: 'add', target: 'project', content: '结论：依赖旧版资料。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    await memory.link({ sourceId: entryId, targetId: oldChunkIds[0], kind: 'supports' });
    assert.equal((await memory.danglingRefs()).length, 0);

    await writeFile(join(root, rel), `${DOC}\n\n## 新增\n\n换版之后多了一节。`, 'utf8');
    const second = await memory.kbIngest({ path: rel, tier: 'verified' });
    assert.equal(second.replaced, true);
    assert.equal(second.unchanged, false);
    assert.equal(second.tier, 'verified');
    assert.ok(second.prunedLinks >= 1, '指向旧块的边要被清掉');
    assert.equal((await memory.danglingRefs()).length, 0, '悬空引用必须还是 0');
    const paths = memoryPaths(root, {}).kb;
    await assert.rejects(() => readFile(join(paths.chunks, hashOld, '1.json'), 'utf8'), '旧块文件该被删掉');
    assert.ok((await memory.allTombstones()).some((stone) => stone.id === oldChunkIds[0]), '旧块 id 要留墓碑');
    const after = await memory.kbRead({ id: second.docId });
    assert.equal(after.tier, 'verified');
    assert.ok(after.chunks.some((chunk) => chunk.text.includes('换版之后多了一节')));
});

await check('入库拒绝：越界路径 / 目录 / 二进制 / 超限体积都要明确报错，不静默入库', async () => {
    const { root, memory } = await freshRoot();
    await assert.rejects(() => memory.kbIngest({ path: '../外面.md' }), /工作目录里面/);
    await assert.rejects(() => memory.kbIngest({ path: '没有这份.md' }), /读不了/);
    await mkdir(join(root, 'adir'), { recursive: true });
    await assert.rejects(() => memory.kbIngest({ path: 'adir' }), /不是文件/);
    await writeFile(join(root, 'bin.md'), `前\0后`, 'utf8');
    await assert.rejects(() => memory.kbIngest({ path: 'bin.md' }), /二进制/);
    await writeFile(join(root, 'big.md'), 'x'.repeat(KB_MAX_DOC_BYTES + 1), 'utf8');
    await assert.rejects(() => memory.kbIngest({ path: 'big.md' }), /超过 kb 的单文档上限/);
    assert.equal((await memory.kbList()).total, 0, '被拒的都不该进 kb');
});

// ── 3. 读取：显式挑选 + 有界 ────────────────────────────────────────────────

await check('读取：文档 id 给整份（有界）、块 id 只给一块、path 也能读', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'c.md', DOC);
    const doc = await memory.kbIngest({ path: rel });
    const byDoc = await memory.kbRead({ id: doc.docId });
    assert.equal(byDoc.chunks_total, byDoc.chunks.length, '这份夹具没超额度，该给全');
    const firstId = byDoc.chunks[0].id;
    const byChunk = await memory.kbRead({ id: firstId });
    assert.equal(byChunk.chunks.length, 1);
    assert.equal(byChunk.chunks[0].id, firstId);
    const byPath = await memory.kbRead({ path: rel });
    assert.equal(byPath.docId, doc.docId);
    assert.equal(byPath.chunks.length, byDoc.chunks.length);
    assert.match(byPath.chunks[0].spanText, /^字符 \d+–\d+$/);
    assert.ok(byPath.chunks[0].header.includes(rel), '块头里要有来源路径');
    // 块号越界：要报「这份文档只有几块」，而不是「没有这份文档」。
    const hash = await contentHashOf(memory, rel);
    await assert.rejects(() => memory.kbRead({ id: chunkIdOf(hash, 99) }), /只有 \d+ 块/);
    await assert.rejects(() => memory.kbRead({ id: 'kb:ffffffffffff' }), /kb 里没有/);
});

await check('读取有界：超出 4000 字符额度时明确说「还有块没给」', async () => {
    const { root, memory } = await freshRoot();
    // 每段 900 字、段间空行：按 1200 字一块 → 一段一块，六段必然超过 4000 字符额度。
    const big = Array.from({ length: 6 }, (unused, index) => `## 第 ${index + 1} 节\n\n${'内容'.repeat(450)}`).join('\n\n');
    const rel = await writeDoc(root, 'big-doc.md', big);
    const doc = await memory.kbIngest({ path: rel });
    assert.ok(doc.chunks >= 6, `该切出多块，实际 ${doc.chunks}`);
    const read = await memory.kbRead({ id: doc.docId });
    assert.equal(read.truncated, true, '额度用完就要如实报截断');
    assert.ok(read.chunks.length < doc.chunks);
    assert.ok(read.chunks.reduce((sum, chunk) => sum + chunk.text.length, 0) <= KB_READ_CHARS,
        '给出的块不该超过额度');
    assert.equal(read.budgetChars, KB_READ_CHARS);
});

await check('读取缺块：manifest 说有几块、盘上却缺时如实计数，不当成空块', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'd.md', DOC);
    const doc = await memory.kbIngest({ path: rel });
    const hash = await contentHashOf(memory, rel);
    await rm(join(memoryPaths(root, {}).kb.chunks, hash, '1.json'), { force: true });
    const read = await memory.kbRead({ id: doc.docId });
    assert.equal(read.chunks_missing, 1);
    assert.equal(read.chunks.length, doc.chunks - 1);
});

// ── 4. 删除与引用完整性 ────────────────────────────────────────────────────

await check('删除：正文删掉、边同事务清掉、墓碑写上、悬空 == 0（按 path 与按 id 两条路）', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'e.md', DOC);
    const doc = await memory.kbIngest({ path: rel });
    const oldHash = await contentHashOf(memory, rel);
    const chunks = (await memory.kbRead({ id: doc.docId })).chunks.map((chunk) => chunk.id);
    await memory.mutate({ action: 'add', target: 'project', content: '结论：依据 e.md。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    // 记忆 → 知识块（entryIndex 里有块 id 才建得起来）与 块 → 记忆 两条边。
    await memory.link({ sourceId: entryId, targetId: chunks[0], kind: 'supports' });
    await memory.link({ sourceId: chunks[1], targetId: entryId, kind: 'derives' });
    const dropped = await memory.kbDrop({ path: rel });
    assert.equal(dropped.prunedLinks, 2, '落在被删块上的两条边都要清');
    assert.ok(dropped.tombstones >= doc.chunks + 1, '文档 id 与每个块 id 都要留墓碑');
    assert.equal((await memory.danglingRefs()).length, 0);
    assert.equal((await memory.kbList()).total, 0);
    await assert.rejects(() => memory.kbRead({ id: doc.docId }), /kb 里没有/);
    await assert.rejects(() => memory.kbRead({ id: chunks[0] }), /kb 里没有/);
    await assert.rejects(() => readFile(join(memoryPaths(root, {}).kb.chunks, oldHash, '1.json'), 'utf8'));
    // 墓碑的作用：被删的 id 不能再被 link（alive 判死），但也不该锁死别的 id。
    await assert.rejects(() => memory.link({ sourceId: entryId, targetId: chunks[0], kind: 'supports' }), /已经被删除|找不到/);
    const second = await writeDoc(root, 'e2.md', DOC);
    const again = await memory.kbIngest({ path: second });
    assert.equal((await memory.kbDrop({ id: again.docId })).path, second, '按 id 删也要能删掉');
});

await check('内容相同的两份文档共用块目录：删掉一份，另一份仍读得回来', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 'same-a.md', DOC);
    const relB = await writeDoc(root, 'same-b.md', DOC);
    const a = await memory.kbIngest({ path: relA });
    const b = await memory.kbIngest({ path: relB });
    assert.equal(await contentHashOf(memory, relA), await contentHashOf(memory, relB),
        '内容相同 ⇒ 同一个内容哈希（块目录共用）');
    assert.notEqual(a.docId, b.docId, '文档 id 按路径寻址，两份是不同的文档');
    // 共用块上的边也不能被误清：块 id 还活着（另一份文档还在引用它）。
    await memory.mutate({ action: 'add', target: 'project', content: '结论：依据共用的那一块。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    const shared = (await memory.kbRead({ id: b.docId })).chunks[0].id;
    await memory.link({ sourceId: entryId, targetId: shared, kind: 'supports' });
    await memory.kbDrop({ path: relA });
    const still = await memory.kbRead({ id: b.docId });
    assert.ok(still.chunks.length >= 3, '另一份文档的块不该被一起删掉');
    assert.equal((await memory.related({ id: entryId })).edges.length, 1,
        '这条边指向的块还活着（另一份文档在用），不该被「删一份」顺手清掉');
    assert.equal((await memory.danglingRefs()).length, 0);
    // 反过来：把最后一份也删掉，块 id 才真的死，边这时必须被清。
    await memory.kbDrop({ path: relB });
    assert.equal((await memory.danglingRefs()).length, 0);
    assert.equal((await memory.related({ id: entryId })).danglingCount, 0);
    const status = await memory.status();
    assert.equal(status.stores[0].kb.docs, 0);
    assert.equal(status.stores[0].kb.chunks, 0);
});

await check('跨根共用内容：删工作区那份，不该把全局根正在用的边清掉', async () => {
    // scope=both 时两个根同时活着；同一个内容哈希可能同时落在两个根里（内容寻址 ⇒ 同一批块 id）。
    // 只按被删那个根算「死没死」，就会把另一个根正在用的边静默清掉 —— 与第二十八轮那条
    // 「只锁一侧等于没锁」是同一类错误，所以死活的判据必须跨根算。
    caseId += 1;
    const rootA = join(TMP, `case-${caseId}-a`);
    const rootB = join(TMP, `case-${caseId}-b`);
    const globalDir = join(TMP, `case-${caseId}-global`);
    await rm(rootA, { recursive: true, force: true });
    await rm(rootB, { recursive: true, force: true });
    await rm(globalDir, { recursive: true, force: true });
    await mkdir(rootA, { recursive: true });
    await mkdir(rootB, { recursive: true });
    const shared = { scope: 'both', globalDir };
    const globalSession = createMemory({ root: rootA, memory: { scope: 'global', globalDir } });
    const bothSession = createMemory({ root: rootB, memory: shared });
    const relA = await writeDoc(rootA, 'shared.md', DOC);
    const relB = await writeDoc(rootB, 'shared.md', DOC);
    const a = await globalSession.kbIngest({ path: relA });
    const b = await bothSession.kbIngest({ path: relB });
    assert.equal(await contentHashOf(globalSession, relA), await contentHashOf(bothSession, relB),
        '同一份内容在两个根里算出同一个内容哈希');
    const sharedChunk = (await bothSession.kbRead({ id: b.docId })).chunks[0].id;
    await bothSession.mutate({ action: 'add', target: 'project', content: '结论：依据共用块（跨根）。' });
    const entryId = (await bothSession.snapshot()).entries[0].id;
    await bothSession.link({ sourceId: entryId, targetId: sharedChunk, kind: 'supports' });
    // 删工作区那份：全局根还在用同一批块 id，这条边必须留着。
    await bothSession.kbDrop({ path: relB });
    assert.equal((await bothSession.related({ id: entryId })).edges.length, 1,
        '全局根还在用这块，边不该被清掉');
    assert.equal((await bothSession.danglingRefs()).length, 0);
    // 全局那份由**只开全局根**的会话来删：它看不见工作区根里的边，所以那条边会变成悬空边。
    // 这里要的契约是「如实报出来」，而不是「静默当成还在」——「悬空引用 == 0」这条不变量
    // 的范围本来就是「本次打开的根」（与 remove 的边界一致，见第二十八轮的跨根注释）。
    await globalSession.kbDrop({ path: relA });
    const after = await bothSession.related({ id: entryId });
    assert.equal(after.edges.length, 0, '另一端已经删了，就不该再算一条正常的边');
    assert.equal(after.danglingCount, 1, '跨会话删掉的来源要如实报成悬空边');
    assert.equal((await bothSession.danglingRefs()).length, 1);
});

// ── 5. 写入期硬规则质量门（24-6 的前半） ────────────────────────────────────

await check('质量门：unverified 的单源内容永远进不了热记忆（显式声明这条路）', async () => {
    const { root, memory } = await freshRoot();
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: '单源说法：某地明年放开。', tier: 'unverified' }),
        /不能进热记忆/,
    );
    assert.equal((await memory.snapshot()).entries.length, 0, '被拒的写入不该留下任何条目');
});

await check('质量门：引用 unverified 的块被拒；引用 verified 的块记下 kbRefs', async () => {
    const { root, memory } = await freshRoot();
    const relBad = await writeDoc(root, 'single-source.md', '# 单源\n\n某论坛说，明年政策会放开。');
    const relGood = await writeDoc(root, 'two-sources.md', '# 两源\n\n两家独立媒体确认了这一条。');
    const bad = await memory.kbIngest({ path: relBad, tier: 'unverified' });
    const good = await memory.kbIngest({ path: relGood, tier: 'verified' });
    const badChunk = (await memory.kbRead({ id: bad.docId })).chunks[0].id;
    const goodChunk = (await memory.kbRead({ id: good.docId })).chunks[0].id;
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: '结论：政策会放开。', source: [badChunk] }),
        /单源未核实/,
    );
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: '结论：查无此源。', source: ['kb:deadbeef0000:1'] }),
        /找不到/,
    );
    const written = await memory.mutate({
        action: 'add', target: 'project', content: '结论：政策会放开（两源确认）。', source: [goodChunk],
    });
    assert.deepEqual(written.entry.kbRefs, [goodChunk]);
    // 依据要**回执出来、也要读得回来**：只写进文件而不回执，模型会以为这次引用没生效。
    assert.ok(renderMutation(written).includes(goodChunk), '写入回执要印出依据的块 id');
    const reread = renderRead(await memory.read({ layer: 'hot' }));
    assert.ok(reread.includes('依据') && reread.includes(goodChunk), '读热记忆要能看到这条结论依据哪块来源');
    // 来源档是「声明的」：没有来源档的正常写入照旧允许（用户口述、项目约定都不是单源检索结果）。
    const plain = await memory.mutate({ action: 'add', target: 'project', content: '约定：汇报统一用 business 主题。' });
    assert.deepEqual(plain.entry.kbRefs, []);
});

await check('质量门：来源档可以在「确认之后升级」，内容一字不改', async () => {
    // 硬规则不能是死路：一份单源材料被第二家独立来源确认之后，必须有路把它从
    // 「待核」升成「已验证」—— 否则只能靠改一个字触发换代，那是在教模型造假。
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, '.office/search/单源.md', '# 单源\n\n某论坛说，明年会放开。');
    const first = await memory.kbIngest({ path: rel, tier: 'unverified' });
    const chunk = (await memory.kbRead({ id: first.docId })).chunks[0].id;
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: '说法：明年会放开。', source: [chunk] }),
        /单源未核实/,
    );
    const promoted = await memory.kbIngest({ path: rel, tier: 'verified' });
    assert.equal(promoted.retiered, true);
    assert.equal(promoted.previousTier, 'unverified');
    assert.equal(promoted.tier, 'verified');
    assert.equal(promoted.unchanged, true, '改的是档，不是内容');
    assert.equal((await memory.kbList()).items[0].tier, 'verified');
    const written = await memory.mutate({ action: 'add', target: 'project', content: '说法：明年会放开（两源确认）。', source: [chunk] });
    assert.deepEqual(written.entry.kbRefs, [chunk]);
    // 不带 tier 的重复入库既不改档也不误降级。
    const again = await memory.kbIngest({ path: rel });
    assert.equal(again.unchanged, true);
    assert.equal(again.retiered, undefined);
    assert.equal(again.tier, 'verified');
});

await check('usedRefs：log 里的 kb: 引用另存成 kbRefs，且能反着查（台账 → 知识块）', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, '.office/search/检索结果.md', '# 检索结果\n\n两家来源都写了同一条事实。');
    const doc = await memory.kbIngest({ path: rel, tier: 'verified' });
    const chunkId = (await memory.kbRead({ id: doc.docId })).chunks[0].id;
    await memory.log([{
        path: '汇报.pptx',
        format: 'ppt',
        purpose: '季度汇报',
        source: ['输入稿.md', chunkId],
    }]);
    const ledger = await memory.browse({});
    const record = ledger.ledger[0];
    assert.deepEqual(record.source, ['输入稿.md'], '文件来源与块引用要分开存');
    assert.deepEqual(record.kbRefs, [chunkId]);
    // 反向查：拿块 id 当 query，能命中用过它的那条台账。
    const found = await memory.read({ layer: 'ledger', query: chunkId });
    assert.equal(found.ledger.items.length, 1);
    assert.ok(found.ledger.items[0].path.includes('汇报.pptx'));
});

// ── 6. 冲突三态（24-6 的后半） ────────────────────────────────────────────

await check('冲突边三态：保留双方、显式承认未决，状态可以改而不重复建边', async () => {
    const { root, memory } = await freshRoot();
    const left = await memory.mutate({ action: 'add', target: 'project', content: '约定：发布窗口定在周三。' });
    const right = await memory.mutate({ action: 'add', target: 'project', content: '另一份材料说发布窗口是周五。' });
    const first = await memory.link({
        sourceId: left.entry.id,
        targetId: right.entry.id,
        kind: 'contradicts',
        conflict: 'inter-context',
        state: 'unresolved',
        note: '两份材料口径不同',
    });
    assert.equal(first.added, true);
    assert.equal(first.link.conflictClass, 'inter-context');
    assert.equal(first.link.conflictState, 'unresolved');
    // 认不出来的类别 / 状态要**报错**，不能悄悄退回「未分类 / 未决」——
    // 那会把上一步记下的判断抹掉（静默的信息丢失）。
    await assert.rejects(
        () => memory.link({ sourceId: left.entry.id, targetId: right.entry.id, kind: 'contradicts', conflict: '看不懂' }),
        /conflict 只支持/,
    );
    await assert.rejects(
        () => memory.link({ sourceId: left.entry.id, targetId: right.entry.id, kind: 'contradicts', state: '看不懂' }),
        /state 只支持/,
    );
    // 不带这两个字段地重复建：既不重复写边，也不改动已经记下的判断。
    const again = await memory.link({ sourceId: left.entry.id, targetId: right.entry.id, kind: 'contradicts' });
    assert.equal(again.added, false);
    assert.equal(again.updated, false);
    assert.equal(again.link.conflictClass, 'inter-context');
    assert.equal(again.link.conflictState, 'unresolved');
    const updated = await memory.link({ sourceId: left.entry.id, targetId: right.entry.id, kind: 'contradicts', state: 'prefer-target' });
    assert.equal(updated.added, false, '同一条边不该重复建');
    assert.equal(updated.updated, true, '状态变了就要落盘');
    assert.equal(updated.link.conflictState, 'prefer-target');
    const related = await memory.related({ id: left.entry.id });
    assert.equal(related.edges.length, 1);
    assert.equal(related.edges[0].conflictState, 'prefer-target');
    assert.equal(related.nodes[0].conflictState, 'prefer-target');
    // 非 contradicts 的边不带冲突字段（不给普通关系安一个「未决」）。
    const supports = await memory.link({ sourceId: left.entry.id, targetId: right.entry.id, kind: 'supports', conflict: 'intra-memory' });
    assert.equal(supports.link.conflictClass, '');
    assert.equal(supports.link.conflictState, '');
});

// ── 7. Pack：整包搬走 kb ──────────────────────────────────────────────────

await check('Pack：kb 清单与块正文一起搬走，导入幂等、被墓碑判死的不复', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'packed.md', DOC);
    const doc = await memory.kbIngest({ path: rel, tier: 'verified' });
    const pack = await memory.exportPack();
    assert.equal(pack.stores[0].kb.manifest.length, 1);
    assert.ok(Object.keys(pack.stores[0].kb.files).length >= 3, '块正文要带进包');

    const { root: other, memory: target } = await freshRoot();
    const firstImport = await target.importPack(pack);
    assert.equal(firstImport.kbDocs, 1);
    const read = await target.kbRead({ id: doc.docId });
    assert.ok(read.chunks.length >= 3);
    assert.equal(read.tier, 'verified');
    const secondImport = await target.importPack(pack);
    assert.equal(secondImport.kbDocs, 0);
    assert.equal(secondImport.skippedKbDocs, 1, '同 id 已存在就跳过（只增不改）');
    assert.equal((await target.kbList()).total, 1);

    // 删掉之后再导入这份旧包：墓碑判死，条目不复活。
    await target.kbDrop({ path: rel });
    const thirdImport = await target.importPack(pack);
    assert.equal(thirdImport.kbDocs, 0);
    assert.ok(thirdImport.skippedKbTombstoned >= 1, '被**文档 id** 墓碑挡住的文档要计入 skippedKbTombstoned');
    assert.equal((await target.kbList()).total, 0);
});

// ── 8. status 与工具面端到端 ──────────────────────────────────────────────

await check('status：每个根给出 kb 的文档数 / 块数 / 体积 / 来源三档分布', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 's1.md', DOC);
    const relB = await writeDoc(root, '.office/search/s2.md', '# 检索\n\n单源说法。');
    const a = await memory.kbIngest({ path: relA });
    const b = await memory.kbIngest({ path: relB, tier: 'unverified' });
    const status = await memory.status();
    const kb = status.stores[0].kb;
    assert.equal(kb.docs, 2);
    assert.equal(kb.chunks, a.chunks + b.chunks);
    assert.deepEqual(kb.tiers, { user: 1, verified: 0, unverified: 1 });
    assert.ok(kb.bytes > 0);
    assert.ok(kb.dir.endsWith('.office/kb') || kb.dir.endsWith('.office\\kb'), `kb 目录该在 .office 下：${kb.dir}`);
});

await check('工具面：office_memory 的入库 / 清单 / 读取 / 删除端到端可用，回执带来源档与 span', async () => {
    const { root } = await freshRoot();
    const rel = await writeDoc(root, '.office/search/直查.md', DOC);
    const tools = buildTools(resolveConfig({}));
    const tool = tools.find((item) => item.name === 'office_memory');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    const ingest = await tool.execute({ action: 'kb-ingest', path: rel, tier: 'unverified' }, exec);
    assert.equal(ingest.ok, true);
    assert.ok(ingest.text.includes('待核（单源未核实）'), '入库回执要写清来源档');
    const list = await tool.execute({ action: 'kb-list' }, exec);
    assert.ok(list.text.includes('kb 清单'));
    assert.ok(list.text.includes('待核'));
    const docId = /kb:[0-9a-f]{12}/.exec(list.text)[0];
    const read = await tool.execute({ action: 'kb-read', id: docId }, exec);
    assert.ok(read.text.includes('字符 '), '读取回执要带 span');
    assert.ok(read.text.includes('不要据此 add 热记忆') || read.text.includes('单源未核实'));
    // 工具面这条路也要过质量门：拿这份 unverified 的块去 add，必须被拒。
    const chunkLine = /【第 1 块｜(kb:[0-9a-f]{12}:\d+)｜/.exec(read.text);
    assert.ok(chunkLine !== null, '读取回执里要有块 id');
    await assert.rejects(
        () => tool.execute({ action: 'add', target: 'project', content: '依据这条单源说法。', source: [chunkLine[1]] }, exec),
        /单源未核实/,
    );
    const drop = await tool.execute({ action: 'kb-drop', id: docId }, exec);
    assert.ok(drop.text.includes('已删除'));
    assert.equal((await tool.execute({ action: 'kb-list' }, exec)).text.includes('还没有入库任何来源'), true);
});

// ── 8.5 检索接线（第二十四轮 24-11）：词法多路 + 弃答 + 有界 + 配额 ──────────
//
// 探针（test/memory-probes.mjs 的 C 组）量的是**自测集上的召回率**；这一节量的是
// action 面的契约：命中给什么、弃答怎么表现、筛选与上限、以及配额拒绝给出的是哪条路。

const SEARCH_DOC_A = [
    '# 召回实验记录',
    '',
    '## 查询展开',
    '',
    '中文召回先用 bigram 展开查询，词序颠倒的查询只有 bigram 才召得回。',
    '',
    '## 排序',
    '',
    '排序用 BM25，先按词频饱和，再按块长度归一。',
].join('\n');
const SEARCH_DOC_B = [
    '# 归档分卷说明',
    '',
    '## 分卷',
    '',
    '单月摘要超过 200 条就开下一卷，裁剪按月份加卷号一起算。',
].join('\n');

await check('检索：命中给块 id / 路径 / span / 预览，预览是真正文的前缀（不是摘要）', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    await writeDoc(root, 'docs/归档分卷说明.md', SEARCH_DOC_B);
    await memory.kbIngest({ path: relA });
    await memory.kbIngest({ path: 'docs/归档分卷说明.md' });

    const found = await memory.kbSearch({ query: '词序颠倒 bigram' });
    assert.ok(found.hits.length > 0, '这条查询该命中');
    const top = found.hits[0];
    assert.equal(top.path, relA, `该命中召回实验记录，实际 ${top.path}`);
    assert.ok(/^kb:[0-9a-f]{12}:\d+$/.test(top.chunkId), `块 id 形状不对：${top.chunkId}`);
    assert.ok(top.spanText.startsWith('字符 '), '命中要带 span，否则引用落不回原文');
    assert.ok(top.preview.length > 0 && top.preview.length <= 280, '预览要有界');
    // 预览必须是**真正文的前缀**：它是「让你决定要不要读整块」的样子，不是摘要。
    const read = await memory.kbRead({ id: top.chunkId });
    assert.equal(read.chunks[0].text.startsWith(top.preview.replace(/…$/, '')), true, '预览要是真正文的前缀');
    // 回执必须写明「分数是相对次序」：两路都归一化过，读成置信度就会过度相信排序。
    const receipt = renderKb(found);
    assert.ok(receipt.includes('相对次序'), receipt);
    assert.ok(receipt.includes(top.chunkId), '回执要给出块 id，否则没法 kb-read 逐字引用');
    // 命中的块头是「标题路径 · 来源 · 日期」，不是别处拼来的上下文。
    assert.ok(top.header.includes(relA), `块头要带来源路径：${top.header}`);
});

await check('检索：中文 bigram 召得回词序颠倒；词表零重叠时弃答（不许端出「最像的那块」）', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    await memory.kbIngest({ path: relA });

    // 正文写的是「按块长度归一」：整串匹配会整片落空，bigram 才召得回。
    const reversed = await memory.kbSearch({ query: '归一 长度' });
    assert.ok(reversed.hits.some((hit) => hit.path === relA), `词序颠倒的查询要召得回：${JSON.stringify(reversed.hits)}`);

    const absent = await memory.kbSearch({ query: '量子退相干 咖啡萃取' });
    assert.equal(absent.matched, 0, '词表零重叠必须弃答');
    assert.equal(absent.hits.length, 0);
    assert.ok(renderKb(absent).includes('弃答'), renderKb(absent));
});

await check('检索：来源档与路径可当筛选，拼错的档位当场报错', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 'verified.md', SEARCH_DOC_A);
    const relB = await writeDoc(root, '.office/search/待核.md', SEARCH_DOC_B);
    await memory.kbIngest({ path: relA, tier: 'verified' });
    await memory.kbIngest({ path: relB, tier: 'unverified' });

    const all = await memory.kbSearch({ query: '分卷 卷号' });
    assert.ok(all.hits.length > 0);
    assert.equal(all.hits[0].path, relB);
    const onlyVerified = await memory.kbSearch({ query: '分卷 卷号', tier: 'verified' });
    assert.equal(onlyVerified.hits.length, 0, 'unverified 的那份不该被档位筛出来');
    const onlyPath = await memory.kbSearch({ query: '分卷 卷号', path: '.office/search/待核.md' });
    assert.ok(onlyPath.hits.length > 0);
    assert.equal(onlyPath.hits.every((hit) => hit.path === relB), true);
    // 显式传的档位必须看得懂（与 kb-ingest 同一条纪律：拼错一个字母不能静默降档）。
    await assert.rejects(() => memory.kbSearch({ query: '分卷', tier: 'verifiedd' }), /只支持/);
    await assert.rejects(() => memory.kbSearch({ query: '   ' }), /需要 query/);
});

await check('检索有界：limit 收口到清单硬上限，截断时如实报「还有命中没给」', async () => {
    const { root, memory } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    await writeDoc(root, 'docs/归档分卷说明.md', SEARCH_DOC_B);
    await memory.kbIngest({ path: 'docs/召回实验记录.md' });
    await memory.kbIngest({ path: 'docs/归档分卷说明.md' });

    const found = await memory.kbSearch({ query: 'bigram BM25' });
    assert.ok(found.matched >= 1);
    const one = await memory.kbSearch({ query: 'bigram BM25', limit: 1 });
    assert.equal(one.hits.length, 1);
    assert.equal(one.truncated, found.matched > 1);
    const huge = await memory.kbSearch({ query: 'bigram BM25', limit: 999 });
    assert.ok(huge.hits.length <= 20, 'limit 要比清单硬上限再收一次');
    // 小数 limit 曾经一路传到 `slice(0, 0.5)` → hits 空，而 matched 非 0 —— 回执于是报出一句
    // **假的弃答**（第四十五轮独立复核打穿 P3）。收口口径：小数取整、下限 1，绝不给 0 条。
    const half = await memory.kbSearch({ query: 'bigram BM25', limit: 0.5 });
    assert.equal(half.hits.length, 1, '小数 limit 不许退化成 0 条（那就是假弃答）');
    assert.equal(renderKb(half).includes('弃答'), false);
});

await check('检索：没命中的三种成因分得清；坏块与没扫完都如实说（复核 P2/P3 的回归）', async () => {
    const { root, memory } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    await memory.kbIngest({ path: 'docs/召回实验记录.md' });

    // ① 筛选把文档全筛掉 ≠ 弃答：真实原因是筛选，换词永远查不到。
    const filtered = await memory.kbSearch({ query: 'bigram', path: 'docs/不存在的.md' });
    assert.equal(filtered.reason, 'filter-empty');
    assert.equal(filtered.hits.length, 0);
    const filteredText = renderKb(filtered);
    assert.ok(filteredText.includes('筛选条件下一份文档都没有'), filteredText);
    assert.ok(!filteredText.includes('**弃答**'), '筛掉了就不该写「弃答」把读者引去换词');

    // ② 正文读不到 ≠ 弃答（块文件被外面删掉，与 kb-read 同一条纪律）。
    const listed = await memory.kbList({ limit: 5 });
    const row = listed.items[0];
    const hashDir = join(root, '.office', 'kb', 'chunks', row.hash);
    const chunkFiles = (await readdir(hashDir)).sort();
    await writeFile(join(hashDir, chunkFiles[0]), '{"id":"kb:broken:1","n":1,"text":"被改坏的正文","span":{"start":0,"end":3},"outline":"","hash":"deadbeef"}\n', 'utf8');
    const damaged = await memory.kbSearch({ query: 'bigram 排序' });
    assert.ok(damaged.chunksMissing >= 1, `坏块要如实计数：${JSON.stringify({ chunksMissing: damaged.chunksMissing })}`);
    assert.ok(renderKb(damaged).includes('读不到'), renderKb(damaged));
    // 坏块不许把半截正文端出来（自校验对不上就按缺块处理）——命中的预览里不该有那句被改坏的正文。
    assert.equal(damaged.hits.some((hit) => hit.preview.includes('被改坏的正文')), false);

    // ③ 真·零重叠才是弃答。
    const absent = await memory.kbSearch({ query: '量子退相干 咖啡萃取' });
    assert.equal(absent.reason, 'no-overlap');
    assert.ok(renderKb(absent).includes('**弃答**'));
});

await check('检索：块数撞上单次上限时要报「没扫完」，不许报成弃答（复核 P2 的回归）', async () => {
    const { root, memory } = await freshRoot();
    // 一份「切成 2100 块」的文档：每节很小，所以正文远小于 2 MB 的入库上限，但块数越过了
    // 单次 2000 块的上限。上限**在最后一份文档里**撞到 —— 这正是 bounded 早先算成 false
    // 的形状（`docsRead === docsTotal`），于是回执既不提「没扫完」、还把「没查到」写成弃答。
    const sections = [];
    for (let i = 0; i < 2100; i += 1) sections.push(`# 节${i}\n\nuq${i} 的正文。`);
    await writeDoc(root, 'docs/大文档.md', sections.join('\n\n'));
    await memory.kbIngest({ path: 'docs/大文档.md' });

    // 目标词只在第 2100 块里 —— 它在扫描上限之外。
    const beyond = await memory.kbSearch({ query: 'uq2099' });
    assert.equal(beyond.matched, 0, '第 2100 块超出单次扫描上限，这次确实扫不到');
    assert.equal(beyond.bounded, true, `上限撞到就必须 bounded=true：${JSON.stringify({ docs: beyond.docs, docsRead: beyond.docsRead, chunks: beyond.chunks })}`);
    assert.equal(beyond.chunks, 2000);
    const text = renderKb(beyond);
    assert.ok(text.includes('没扫完'), text);
    assert.ok(text.includes('不是「不存在」'), text);
    // 上限之内的一块仍然查得到，且回执也带「没扫完」的说明（命中不等于扫完）。
    const inside = await memory.kbSearch({ query: 'uq5' });
    assert.equal(inside.matched > 0, true);
    assert.ok(renderKb(inside).includes('没扫完'), renderKb(inside));
});

await check('工具面：kb-search 端到端可用，并单独占「知识块检索」这一档每回合配额', async () => {
    const { root } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    const tool = buildTools(resolveConfig({})).find((item) => item.name === 'office_memory');
    const events = [{ type: 'turn/start', data: { turn: 4 } }];
    const exec = { agent: { session: { id: 'kb-search-quota', header: { cwd: root }, ownEvents: () => events } }, signal: new AbortController().signal };
    assert.equal((await tool.execute({ action: 'kb-ingest', path: 'docs/召回实验记录.md' }, exec)).ok, true);

    const hit = await tool.execute({ action: 'kb-search', query: '排序 BM25' }, exec);
    assert.equal(hit.ok, true);
    assert.ok(hit.text.includes('相对次序'), hit.text);
    // 默认 2 次：第二次放行，第三次拒绝 —— 且拒绝要给出**真的能走**的那条路（kb-read，
    // 不是记忆检索那套「换 layer」，kb 没有 layer 可换）。
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 归一' }, exec)).ok, true);
    const denied = await tool.execute({ action: 'kb-search', query: '排序 词频' }, exec);
    assert.equal(denied.ok, false);
    assert.ok(denied.text.includes('知识库检索配额已用完'), denied.text);
    assert.ok(denied.text.includes('kb-read'), '拒绝要给出真的能走的那条路');
    // 缺 query 时不占配额、直接说清要传什么。
    const empty = await tool.execute({ action: 'kb-search', query: '  ' }, exec);
    assert.equal(empty.ok, false);
    assert.ok(empty.text.includes('需要 query'));
    // 0 = 不限制：换一套配置重建工具面。
    const unlimited = buildTools(resolveConfig({ memory: { quota: { kbSearchPerTurn: 0 } } }))
        .find((item) => item.name === 'office_memory');
    for (let i = 0; i < 4; i += 1) {
        assert.equal((await unlimited.execute({ action: 'kb-search', query: '排序 归一' }, exec)).ok, true, '0 = 不限制');
    }
});

await check('工具面：抛错的检索不占名额（配额要退回来，复核 P3 的回归）', async () => {
    const { root } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    const tool = buildTools(resolveConfig({})).find((item) => item.name === 'office_memory');
    const events = [{ type: 'turn/start', data: { turn: 7 } }];
    const exec = { agent: { session: { id: 'kb-search-refund', header: { cwd: root }, ownEvents: () => events } }, signal: new AbortController().signal };
    assert.equal((await tool.execute({ action: 'kb-ingest', path: 'docs/召回实验记录.md' }, exec)).ok, true);

    // 拼错档位 → 当场抛错。它**不该**吃掉默认 2 次里的一次。
    await assert.rejects(() => tool.execute({ action: 'kb-search', query: '排序', tier: 'verifiedd' }, exec), /只支持/);
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 BM25' }, exec)).ok, true,
        '抛错的那次把名额吃掉了：错拼一个档位就白丢一次机会');
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 归一' }, exec)).ok, true);
    const denied = await tool.execute({ action: 'kb-search', query: '排序 词频' }, exec);
    assert.equal(denied.ok, false, '用完两次之后仍该拒绝');
});

/** 一份文档的**内容哈希**（manifest 行的 `hash`，块目录名就是它；与按路径寻址的文档 id 不同）。 */
async function contentHashOf(memory, rel) {
    const listed = await memory.kbList({ limit: 50 });
    const row = listed.items.find((item) => item.path === rel);
    assert.ok(row !== undefined, `kb 清单里该有 ${rel}`);
    return row.hash;
}

// ── 9. 独立验证挖出来的洞：每条都配一条回归用例 ────────────────────────────

await check('F2 旧 Pack 复活：同内容两份只删一份时，文档 id 的墓碑也要判', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 'twin-a.md', DOC);
    const relB = await writeDoc(root, 'twin-b.md', DOC);
    await memory.kbIngest({ path: relA });
    await memory.kbIngest({ path: relB });
    const pack = await memory.exportPack();
    // 只删 a：块 id 因为 b 还活着而没有墓碑，只有 a 的**文档 id** 留了墓碑。
    await memory.kbDrop({ path: relA });
    assert.equal((await memory.kbList()).total, 1);
    const back = await memory.importPack(pack);
    assert.equal(back.kbDocs, 0, '被文档 id 墓碑判死的那份不该被旧 Pack 搬回来');
    assert.equal(back.skippedKbTombstoned >= 1, true);
    assert.equal((await memory.kbList()).total, 1);
    assert.equal((await memory.kbList()).items[0].path, relB);
});

await check('F1① 洗档通道：带 unverified 引用的热记忆不能靠导入进库', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, '.office/search/单源.md', '# 单源\n\n某论坛说会放开。');
    const doc = await memory.kbIngest({ path: rel, tier: 'unverified' });
    const chunk = (await memory.kbRead({ id: doc.docId })).chunks[0].id;
    // 手工造一份「热记忆条目自带走 unverified 引用」的包：本地 add 会被拒，导入也不许过。
    const poisoned = {
        format: 'dsh-office-memory-pack',
        version: 1,
        exportedAt: new Date().toISOString(),
        workspace: root,
        layers: { hot: true, ledger: true, archive: true },
        stores: [{
            id: 'workspace',
            dir: join(root, '.office/memory'),
            entries: [{ id: 'm-poison', target: 'project', content: '结论：政策会放开。', importance: 'normal', kbRefs: [chunk], createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }],
            ledger: [],
            archiveIndex: { version: 1, digests: [] },
            archiveFiles: {},
            archiveItemFiles: {},
            kb: { manifest: [], files: {} },
            links: [],
            tombstones: [],
        }],
    };
    const { root: other, memory: target } = await freshRoot();
    const report = await target.importPack(poisoned);
    assert.equal(report.entries, 0);
    assert.equal(report.skippedUnverified, 1);
    assert.equal((await target.snapshot()).entries.length, 0, '单源引用的条目不该落进热记忆');
});

await check('F1②/F4 档位取更保守的一侧：同一份内容，本地 unverified 压过包里的 verified', async () => {
    // 边界先说清：Pack 与盘上文件是**同一信任级**（谁能改盘上的文件，就能改包里的字段），
    // 插件不承诺「防伪造」，但它承诺**一致**：同一份内容有两处说法时取更保守的那个。
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'a.md', DOC);
    await memory.kbIngest({ path: rel, tier: 'verified' });
    const pack = await memory.exportPack();
    const tampered = JSON.parse(JSON.stringify(pack));
    tampered.stores[0].kb.manifest[0].tier = 'verified'; // 包里仍是 verified（这份本来就是）
    tampered.stores[0].entries = [];

    // 目标库里先有一份**同样内容**（同哈希、不同路径）的 unverified 来源。
    const { root: other, memory: target } = await freshRoot();
    const relOther = await writeDoc(other, 'b.md', DOC);
    const local = await target.kbIngest({ path: relOther, tier: 'unverified' });
    const chunk = (await target.kbRead({ id: local.docId })).chunks[0].id;
    const report = await target.importPack(tampered);
    assert.equal(report.kbDocs, 1);
    const imported = (await target.kbList()).items.find((item) => item.path === rel);
    assert.equal(imported.tier, 'unverified', '本地同一份内容说了 unverified，包里的说法不能把它抬起来');
    await assert.rejects(
        () => target.mutate({ action: 'add', target: 'project', content: '依据来源。', source: [chunk] }),
        /单源未核实/,
    );
});

await check('F1③ 包里的块正文与 id 不匹配（或正文被改）时整份拒收', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'tamper.md', DOC);
    const doc = await memory.kbIngest({ path: rel });
    const pack = await memory.exportPack();
    const key = Object.keys(pack.stores[0].kb.files)[0];
    const chunk = JSON.parse(pack.stores[0].kb.files[key]);
    chunk.text = '被改过的正文';
    pack.stores[0].kb.files[key] = JSON.stringify(chunk);
    const { root: other, memory: target } = await freshRoot();
    const report = await target.importPack(pack);
    assert.equal(report.kbDocs, 0);
    assert.equal(report.skippedKbDamaged, 1);
    assert.equal((await target.kbList()).total, 0);
    // 自校验也对不上：正文改了但 hash 没改。
    assert.notEqual(doc.docId, '', '夹具本身有效');
});

await check('F1④ replace 要重审条目**已有**的引用；传 source:[] 可以清掉引用', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'good.md', '# 好来源\n\n两家独立媒体确认。');
    const doc = await memory.kbIngest({ path: rel, tier: 'verified' });
    const chunk = (await memory.kbRead({ id: doc.docId })).chunks[0].id;
    const written = await memory.mutate({ action: 'add', target: 'project', content: '结论 A。', source: [chunk] });
    // 把同一块降档成 unverified（模拟「后来发现只有一个来源」）：带这个引用的条目再改就过不了门。
    await memory.kbIngest({ path: rel, tier: 'unverified' });
    await assert.rejects(
        () => memory.mutate({ action: 'replace', target: 'project', id: written.entry.id, content: '结论 A（改一个字）。' }),
        /单源未核实/,
    );
    // 清掉引用（显式传空数组）之后可以改，且条目上不再留引用。
    const cleared = await memory.mutate({ action: 'replace', target: 'project', id: written.entry.id, content: '结论 A（改一个字）。', source: [] });
    assert.deepEqual(cleared.entry.kbRefs, []);
});

await check('F3 同内容不同切块档：块目录与块头都不串（含 .office/search 档）', async () => {
    const { root, memory } = await freshRoot();
    const body = Array.from({ length: 6 }, (unused, index) => `## 第 ${index + 1} 节\n\n${'内容'.repeat(200)}`).join('\n\n');
    const relDoc = await writeDoc(root, 'normal.md', body);
    const relSearch = await writeDoc(root, '.office/search/q.md', body);
    const doc = await memory.kbIngest({ path: relDoc });
    const search = await memory.kbIngest({ path: relSearch });
    // 档不同（1200 / 900）⇒ 哈希不同 ⇒ 块目录不同、块数可以不同。
    assert.notEqual(await contentHashOf(memory, relDoc), await contentHashOf(memory, relSearch),
        '切块档要进内容哈希，否则两份不同边界的块会共用一个 id 空间');
    const readDoc = await memory.kbRead({ id: doc.docId });
    const readSearch = await memory.kbRead({ id: search.docId });
    assert.equal(readDoc.chunks_total, doc.chunks);
    assert.equal(readDoc.chunks.length + readDoc.chunks_missing, doc.chunks);
    assert.equal(readSearch.chunks_total, search.chunks);
    // 块头是**文档级**的：读哪份就写哪份的来源，共用块目录时不会串。
    for (const chunk of readDoc.chunks) assert.ok(chunk.header.includes(relDoc), `块头来源串了：${chunk.header}`);
    for (const chunk of readSearch.chunks) assert.ok(chunk.header.includes(relSearch), `块头来源串了：${chunk.header}`);
});

await check('F4 跨根同一份内容档位撞车：门看的是更保守的那一侧', async () => {
    caseId += 1;
    const rootA = join(TMP, `case-${caseId}-a`);
    const rootB = join(TMP, `case-${caseId}-b`);
    const globalDir = join(TMP, `case-${caseId}-global`);
    await rm(rootA, { recursive: true, force: true });
    await rm(rootB, { recursive: true, force: true });
    await rm(globalDir, { recursive: true, force: true });
    await mkdir(rootA, { recursive: true });
    await mkdir(rootB, { recursive: true });
    const globalSession = createMemory({ root: rootA, memory: { scope: 'global', globalDir } });
    const bothSession = createMemory({ root: rootB, memory: { scope: 'both', globalDir } });
    const relGlobal = await writeDoc(rootA, 'shared.md', DOC);
    const relLocal = await writeDoc(rootB, 'shared.md', DOC);
    await globalSession.kbIngest({ path: relGlobal, tier: 'unverified' });
    const local = await bothSession.kbIngest({ path: relLocal });
    assert.equal(local.tier, 'user');
    const chunk = (await bothSession.kbRead({ id: local.docId })).chunks[0].id;
    // 同一个内容哈希在全局根是 unverified、在工作区根是 user：门必须看更保守的那一侧，
    // 不能因为遍历顺序让后一个把前一个覆盖掉。
    await assert.rejects(
        () => bothSession.mutate({ action: 'add', target: 'project', content: '结论：依据这份内容。', source: [chunk] }),
        /单源未核实/,
    );
});

await check('Pack 导入：块 id 的墓碑挡得住「手改了文档 id」的包', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'orig.md', DOC);
    await memory.kbIngest({ path: rel });
    const pack = await memory.exportPack();
    await memory.kbDrop({ path: rel });
    const forged = JSON.parse(JSON.stringify(pack));
    // 手改文档 id：绕开文档级墓碑，只剩块 id 那一道闸（defense in depth）。
    forged.stores[0].kb.manifest[0].id = 'kb:deadbeef0000';
    const report = await memory.importPack(forged);
    assert.equal(report.kbDocs, 0);
    assert.equal(report.skippedKbDamaged, 1);
    assert.equal((await memory.kbList()).total, 0);
});

await check('F5 换代不误清「指向文档 id」的边（文档 id 换代后还活着）', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'evolve.md', DOC);
    const doc = await memory.kbIngest({ path: rel });
    await memory.mutate({ action: 'add', target: 'project', content: '结论：这份材料是依据。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    await memory.link({ sourceId: entryId, targetId: doc.docId, kind: 'derives' });
    await writeFile(join(root, rel), `${DOC}\n\n## 追加\n\n换版之后多了一节。`, 'utf8');
    const second = await memory.kbIngest({ path: rel });
    assert.equal(second.replaced, true);
    const related = await memory.related({ id: entryId });
    assert.equal(related.edges.length, 1, '文档 id 换代后还活着，这条边不该被清掉');
    assert.equal(related.danglingCount, 0);
    const stones = (await memory.allTombstones()).map((stone) => stone.id);
    assert.ok(!stones.includes(doc.docId), '活着的文档 id 不该有墓碑');
    assert.equal((await memory.danglingRefs()).length, 0);
});

await check('F6 跳级标题（### 出现在 # 之前）不该拼出空层级', () => {
    const chunks = splitDocument('### 深标题\n\n正文一。\n\n#### 更深\n\n正文二。', { chunkChars: 400 });
    assert.equal(chunks[0].outline, '深标题');
    assert.equal(chunks[1].outline, '深标题 / 更深');
    assert.ok(!chunks[0].outline.includes(' /  /'), `空层级没滤掉：${JSON.stringify(chunks[0].outline)}`);
});

await check('F9 add 的材料来源不丢：文件来源落在条目的 sources 上并回执出来', async () => {
    const { root, memory } = await freshRoot();
    const written = await memory.mutate({
        action: 'add', target: 'project', content: '结论：参照输入稿写成。', source: ['输入稿.md', 'https://example.com/a'],
    });
    assert.deepEqual(written.entry.kbRefs, []);
    assert.deepEqual(written.entry.sources, ['输入稿.md', 'https://example.com/a']);
    const reread = renderRead(await memory.read({ layer: 'hot' }));
    assert.ok(reread.includes('来源') && reread.includes('输入稿.md'), '读回执要能看出这条依据哪份材料');
});

await check('F11 log 先分流再封顶：前面堆满文件来源也挤不掉 kb 引用', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, '.office/search/f11.md', '# 来源\n\n两家媒体确认。');
    const doc = await memory.kbIngest({ path: rel, tier: 'verified' });
    const chunk = (await memory.kbRead({ id: doc.docId })).chunks[0].id;
    const files = Array.from({ length: 12 }, (unused, index) => `输入-${index}.md`);
    await memory.log([{ path: '产出.pptx', format: 'ppt', source: [...files, chunk] }]);
    const record = (await memory.browse({})).ledger[0];
    assert.ok(record.kbRefs.includes(chunk), 'kb 引用不该被前面的文件来源挤掉');
    assert.equal(record.source.length, 8);
});

await check('F13c 拼错的来源档要报错，不静默降档', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'tier-typo.md', DOC);
    await assert.rejects(() => memory.kbIngest({ path: rel, tier: 'verifed' }), /tier 只支持/);
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: '结论。', tier: 'unverifed' }),
        /tier 只支持/,
    );
    assert.equal((await memory.kbList()).total, 0, '被拒的档位不该留下任何入库');
});

// ── 收尾 ──────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    if (item.ok) console.log(`PASS  ${item.name}`);
    else console.log(`FAIL  ${item.name}  → ${item.error?.message ?? item.error}`);
}
console.log(`\nkb: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
