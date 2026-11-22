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
import { existsSync, readFileSync } from 'node:fs';
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
    chunkOverlapFor,
    docIdOf,
    overlapChunks,
    splitDocument,
} from '../src/kb.js';
import { createMemory, memoryPaths, renderKb, renderMutation, renderRead } from '../src/memory.js';
import { buildTools } from '../src/tools.js';

// 按进程隔离临时根：caseId 从 0 开始，两个进程同时跑会复用同一批 case 目录；
// 进程内队列管不住跨进程的旧快照回写（实测两路并发 37/64、31/64，出现悬空边 / EBUSY）。
const TMP = fileURLToPath(new URL(`../../../.office/runtime/tmp/kb-${process.pid}/`, import.meta.url));
// 默认知识库落在 $DSH_HOME/.office/kb：测试要把 HOME 指到用例目录，
// 否则 buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } })) 会写进本机真实知识库（第八十八轮单库后尤其明显）。
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

/** 每个用例一个干净的工作目录：kb 是按工作目录存的，串了目录就串了结论。 */
let caseId = 0;

/**
 * 全局根的测试现场。
 *
 * kb 与记忆目录是**兄弟目录**（见 `kbDirOf`：`join(dirname(memoryDir), 'kb')`），
 * 所以 `globalDir` 必须放进一个独立的 `.office/memory`，否则同一 TMP 下的用例会
 * 共用 `<TMP>/kb` —— 一个用例入库的块会被另一个用例看见（第八十轮实测：新用例的
 * `scope.md` 污染了跨根用例）。返回的 `parent` 用于整场清理。
 */
function globalFixture(id, tag) {
    const parent = join(TMP, `case-${id}-${tag}`);
    return { parent, dir: join(parent, '.office', 'memory') };
}

async function freshRoot(memoryConfig = {}) {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    await rm(root, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    // 知识库是单库（第八十八轮）：每个 case 一个库目录，避免用例互相看见文档。
    const config = { kbDir: join(root, '.office', 'kb'), ...memoryConfig };
    return { root, memory: createMemory({ root, memory: config }), memoryConfig: config };
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

/** 入库评估闸（P1-4b）：`tier:'verified'` 的文档必须自带 ≥2 个独立来源，夹具统一补两条。 */
const SOURCES = '\n\n来源：https://news-a.example.com/report、https://news-b.example.org/report\n';

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

await check('切块：相邻块带重叠（起点退到句末边界），仍可回溯且不超过上限', () => {
    const source = DOC;
    const plain = splitDocument(source, { chunkChars: 400, overlap: 0 });
    const chunks = splitDocument(source, { chunkChars: 400 });
    assert.equal(chunks.length, plain.length, '加重叠不该改变块数（它只挪起点）');
    assert.ok(chunks.length >= 3, `这份夹具该切出至少三块，实际 ${chunks.length}`);
    const gap = chunkOverlapFor(400);
    for (const chunk of chunks) {
        assert.equal(source.slice(chunk.span.start, chunk.span.end), chunk.text,
            `第 ${chunk.n} 块带重叠后 span 仍要逐字相等`);
    }
    for (let index = 1; index < chunks.length; index += 1) {
        const prev = chunks[index - 1];
        const cur = chunks[index];
        assert.ok(cur.span.start < prev.span.end, `第 ${index + 1} 块没有与上一块重叠`);
        assert.ok(prev.span.end - cur.span.start <= gap, `重叠量失控：${prev.span.end - cur.span.start} > ${gap}`);
        assert.ok(cur.span.start >= prev.span.start, '重叠不许退到上一块起点之前（否则一块会被另一块包住）');
        // 退让后的起点要落在边界上：要么是上一块正文的第一个字，要么前一个字符是边界。
        const head = cur.text.slice(0, Math.min(12, cur.text.length));
        assert.equal(head, source.slice(cur.span.start, cur.span.start + head.length));
    }
    // 关掉重叠：与旧行为一致，块之间不共享正文。
    for (let index = 1; index < plain.length; index += 1) {
        assert.ok(plain[index].span.start >= plain[index - 1].span.end, 'overlap:0 必须完全关闭重叠');
    }
});

await check('切块：重叠量按块大小派生，夹在 60–300；overlapChunks 对模型切块同样成立', () => {
    assert.equal(chunkOverlapFor(200), 60, '小块至少有 60 字的上下文');
    assert.equal(chunkOverlapFor(1200), 150);
    assert.equal(chunkOverlapFor(8000), 300, '大块最多带 300 字，别让块翻倍');
    // 手工构造两段（模拟模型切片的结果），再由 overlapChunks 补重叠。
    const source = '第一段正文，讲的是营收。第二段正文，讲的是成本与毛利。第三段正文，收尾。';
    const middle = source.indexOf('第二段');
    const last = source.indexOf('第三段');
    const chunks = [
        { n: 1, text: source.slice(0, middle), span: { start: 0, end: middle }, outline: 'A' },
        { n: 2, text: source.slice(middle, last), span: { start: middle, end: last }, outline: 'A' },
        { n: 3, text: source.slice(last), span: { start: last, end: source.length }, outline: 'A' },
    ];
    overlapChunks(source, chunks, 8);
    for (const chunk of chunks.slice(1)) {
        assert.equal(source.slice(chunk.span.start, chunk.span.end), chunk.text, 'span 等式对模型切块也成立');
    }
    assert.ok(chunks[1].span.start < middle, '第二块要带上第一块的句尾');
    assert.ok(chunks[2].span.start < last, '第三块要带上第二块的句尾');
    assert.ok(chunks[1].text.slice(0, 2) === '讲的', '退让应停在句读边界之后（从下一句开头起）：' + chunks[1].text.slice(0, 6));
    assert.ok(/[。！？；，、\n]/.test(source[chunks[1].span.start - 1]), '退让后的起点前一个字符必须是边界');
    // 不改原数组结构：仍然是三块，n 不变。
    assert.deepEqual(chunks.map((chunk) => chunk.n), [1, 2, 3]);
});

await check('切块：同一份正文只在重叠量变化时才换 docHash 档位（升级不会被判成没变）', () => {
    const profileOf = (chunkChars, overlap) => `${chunkChars}:${overlap}`;
    assert.notEqual(profileOf(1200, 0), profileOf(1200, 150), '重叠量必须进档位');
    assert.equal(profileOf(1200, chunkOverlapFor(1200)), profileOf(1200, 150));
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
    const rel = await writeDoc(root, 'b.md', `${DOC}${SOURCES}`);
    const first = await memory.kbIngest({ path: rel });
    const hashOld = await contentHashOf(memory, rel);
    const oldChunkIds = (await memory.kbRead({ id: first.docId })).chunks.map((chunk) => chunk.id);
    // 用旧块 id 建一条边：换代时这条边必须被同事务清掉。
    await memory.mutate({ action: 'add', target: 'project', content: '结论：依赖旧版资料。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    await memory.link({ sourceId: entryId, targetId: oldChunkIds[0], kind: 'supports' });
    assert.equal((await memory.danglingRefs()).length, 0);

    await writeFile(join(root, rel), `${DOC}${SOURCES}\n\n## 新增\n\n换版之后多了一节。`, 'utf8');
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

// ── 2b. OCR 正文入库（第一百零一轮）────────────────────────────────────────

await check('OCR 正文入库：正文进文档、块可回溯、kind 与 manifest 记着「这是识别结果」', async () => {
    const { root, memory } = await freshRoot();
    const text = `# 扫描件${'\n\n'}识别出来的正文，讲的是营收与成本。`.concat('后续细节。'.repeat(300));
    const value = await memory.kbIngest({ path: 'uploads/扫描件.png', bodyText: text, ocr: true });
    assert.equal(value.ocr, true);
    assert.equal(value.kind, 'ocr', 'OCR 正文的 kind 不该退回 text');
    assert.ok(value.chunks >= 2, `长正文该切多块，实际 ${value.chunks}`);
    assert.equal(value.chars, text.length);
    const listed = await memory.kbList();
    assert.equal(listed.total, 1);
    assert.equal(listed.items[0].path, 'uploads/扫描件.png');
    assert.equal(listed.items[0].ocr.chars, text.length);
    assert.match(listed.items[0].ocr.at, /^\d{4}-\d{2}-\d{2}T/, 'manifest 要记住识别时刻');
    const read = await memory.kbRead({ path: 'uploads/扫描件.png' });
    assert.equal(read.chunks_total, value.chunks);
    for (const chunk of read.chunks) {
        assert.ok(text.includes(chunk.text), `第 ${chunk.n} 块的正文必须能在源正文里原样找到`);
    }
});

await check('OCR 正文入库：同内容幂等；换一次识别结果就按同路径换代，不新增文档', async () => {
    const { root, memory } = await freshRoot();
    const first = await memory.kbIngest({ path: 'uploads/扫描件.png', bodyText: '第一次识别：营收增长。', ocr: true });
    const again = await memory.kbIngest({ path: 'uploads/扫描件.png', bodyText: '第一次识别：营收增长。', ocr: true });
    assert.equal(again.unchanged, true, '同一份识别结果再存一次不该重写');
    assert.equal(again.docId, first.docId);
    const second = await memory.kbIngest({ path: 'uploads/扫描件.png', bodyText: '第二次识别：营收与成本都变了。', ocr: true });
    assert.equal(second.replaced, true, '同一路径的新识别结果要换代');
    assert.equal(second.unchanged, false);
    assert.equal(second.docId, first.docId, '文档 id 按路径寻址，换代不换 id');
    const listed = await memory.kbList();
    assert.equal(listed.total, 1, '一份文件在库里只有一份正文');
    const read = await memory.kbRead({ path: 'uploads/扫描件.png' });
    assert.ok(read.chunks.some((chunk) => chunk.text.includes('第二次识别')));
    assert.ok(read.chunks.every((chunk) => !chunk.text.includes('第一次识别')), '旧识别结果不该残留');
});

await check('OCR 正文入库：早先存成 <stem>.ocr.md 的旧笔记一并收掉并留墓碑', async () => {
    const { root, memory } = await freshRoot();
    const legacyRel = await writeDoc(root, 'uploads/扫描件.ocr.md', '# 旧识别\n\n旧的识别正文。\n');
    const legacy = await memory.kbIngest({ path: legacyRel });
    assert.equal((await memory.kbList()).total, 1);
    const saved = await memory.kbUpload({
        name: '扫描件.png', folder: '', text: '新的识别正文。', ocr: true,
    });
    assert.equal(saved.ocr, true);
    assert.equal(saved.removedOcrNote, legacyRel, '回执要点名被收掉的旧笔记');
    assert.equal(saved.replaced, true);
    const listed = await memory.kbList();
    assert.equal(listed.total, 1, '旧笔记不能继续留在库里');
    assert.equal(listed.items[0].path, 'uploads/扫描件.png');
    await assert.rejects(() => memory.kbRead({ path: legacyRel }), /kb 里没有/);
    const stones = await memory.allTombstones();
    assert.ok(stones.some((stone) => stone.id === docIdOf(legacyRel)), '旧笔记要留墓碑，别让旧边指回来');
    assert.equal((await memory.danglingRefs()).length, 0);
});

await check('OCR 正文入库：空正文与超限正文都要明确拒绝，不写 manifest', async () => {
    const { root, memory } = await freshRoot();
    await assert.rejects(() => memory.kbIngest({ path: 'uploads/空.png', bodyText: '   ', ocr: true }), /正文是空的/);
    await assert.rejects(
        () => memory.kbIngest({ path: 'uploads/大.png', bodyText: 'x'.repeat(KB_MAX_DOC_BYTES + 1), ocr: true }),
        /超过 kb 的单文档上限/,
    );
    await assert.rejects(() => memory.kbUpload({ name: '空.png', folder: '', text: '  ', ocr: true }), /正文是空的/);
    assert.equal((await memory.kbList()).total, 0);
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
    assert.equal(status.kb.docs, 0);
    assert.equal(status.kb.chunks, 0);
});

await check('单库共用内容：删一份，另一份正在用的块与边都不该被清掉', async () => {
    // 单库之后「跨根」不再存在，但同一份内容仍可以以两条路径入库（内容寻址 ⇒ 同一批块 id）。
    // 删除一份时，死活判据必须看**整份清单**：另一份文档还在引用这些块，边就不能清。
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 'shared-a.md', DOC);
    const relB = await writeDoc(root, 'shared-b.md', DOC);
    const a = await memory.kbIngest({ path: relA });
    const b = await memory.kbIngest({ path: relB });
    assert.equal(await contentHashOf(memory, relA), await contentHashOf(memory, relB),
        '同一份内容算出同一个内容哈希');
    const sharedChunk = (await memory.kbRead({ id: b.docId })).chunks[0].id;
    await memory.mutate({ action: 'add', target: 'project', content: '结论：依据共用的那一块。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    await memory.link({ sourceId: entryId, targetId: sharedChunk, kind: 'supports' });
    await memory.kbDrop({ path: relA });
    assert.equal((await memory.related({ id: entryId })).edges.length, 1,
        '另一份文档还在用这块，边不该被清掉');
    assert.equal((await memory.danglingRefs()).length, 0);
    // 最后一份也删掉，块 id 才真的死，边这时必须被清。
    await memory.kbDrop({ path: relB });
    assert.equal((await memory.related({ id: entryId })).edges.length, 0);
    assert.equal((await memory.danglingRefs()).length, 0);
});

await check('单库：scope 参数不再选根，入库 / 清单 / 删除始终落在同一个 kbDir', async () => {
    // 第八十八轮：知识库不再分全局库 / 工作区库。旧调用面带 scope 也不该改变存放位置。
    const { root, memory } = await freshRoot({ scope: 'workspace' });
    const rel = await writeDoc(root, 'single.md', DOC);
    const ing = await memory.kbIngest({ path: rel, scope: 'global' });
    assert.equal(ing.origin, 'kb');
    const list = await memory.kbList();
    assert.equal(list.total, 1);
    assert.equal(list.items[0].origin, 'kb');
    assert.doesNotMatch(renderKb(list), /（全局）/);
    assert.doesNotMatch(renderKb(list), /（工作区）/);
    const search = await memory.kbSearch({ query: '营收' });
    assert.ok(search.hits.length > 0);
    assert.equal(search.hits[0].origin, 'kb');
    const dropped = await memory.kbDrop({ path: rel, scope: 'global' });
    assert.equal(dropped.origin, 'kb');
    assert.equal((await memory.kbList()).total, 0);
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
    const relGood = await writeDoc(root, 'two-sources.md', `# 两源\n\n两家独立媒体确认了这一条。${SOURCES}`);
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
    const rel = await writeDoc(root, '.office/search/单源.md', `# 单源\n\n某论坛说，明年会放开。${SOURCES}`);
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
    const rel = await writeDoc(root, '.office/search/检索结果.md', `# 检索结果\n\n两家来源都写了同一条事实。${SOURCES}`);
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
    const rel = await writeDoc(root, 'packed.md', `${DOC}${SOURCES}`);
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

await check('status：单库给出 kb 的文档数 / 块数 / 体积 / 来源三档分布', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 's1.md', DOC);
    const relB = await writeDoc(root, '.office/search/s2.md', '# 检索\n\n单源说法。');
    const a = await memory.kbIngest({ path: relA });
    const b = await memory.kbIngest({ path: relB, tier: 'unverified' });
    const status = await memory.status();
    const kb = status.kb;
    assert.equal(kb.docs, 2);
    assert.equal(kb.chunks, a.chunks + b.chunks);
    assert.deepEqual(kb.tiers, { user: 1, verified: 0, unverified: 1 });
    assert.ok(kb.bytes > 0);
    assert.ok(kb.dir.endsWith('.office/kb') || kb.dir.endsWith('.office\\kb'), `kb 目录该在 .office 下：${kb.dir}`);
});

await check('工具面：office_memory 的入库 / 清单 / 读取 / 删除端到端可用，回执带来源档与 span', async () => {
    const { root } = await freshRoot();
    const rel = await writeDoc(root, '.office/search/直查.md', DOC);
    const tools = buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } }));
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

await check('工具面：kb-link / kb-related / kb-summarize / kb-summaries 端到端可用（第八十五轮）', async () => {
    const { root } = await freshRoot();
    const text = [
        '# 块关系与摘要',
        '',
        '## 第一段',
        '',
        '第一块讲块间联系：一块可以连到另一块，关系双向可走。',
        '',
        '## 第二段',
        '',
        '第二块讲摘要：多块可以合成一个摘要块，摘要也参与检索。',
    ].join('\n');
    const rel = await writeDoc(root, 'docs/块关系.md', text);
    const tool = buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } })).find((item) => item.name === 'office_memory');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    assert.equal((await tool.execute({ action: 'kb-ingest', path: rel }, exec)).ok, true);
    const read = await tool.execute({ action: 'kb-read', path: rel }, exec);
    const ids = [...read.text.matchAll(/【第 \d+ 块｜(kb:[0-9a-f]{12}:\d+)｜/g)].map((m) => m[1]);
    const docId = /文档 id：(kb:[0-9a-f]{12})/.exec(read.text)[1];
    assert.ok(ids.length >= 2, `要至少两块才能测关系，实际 ${ids.length}`);
    const linked = await tool.execute({ action: 'kb-link', sourceId: ids[0], targetId: ids[1], kind: 'supports', note: '同主题' }, exec);
    assert.equal(linked.ok, true, linked.text);
    assert.ok(linked.text.includes('块间关系已建立'), linked.text);
    const related = await tool.execute({ action: 'kb-related', chunkId: ids[0] }, exec);
    assert.equal(related.ok, true, related.text);
    assert.ok(related.text.includes(ids[1]), '邻居回执要列出对面块');
    const summary = await tool.execute({ action: 'kb-summarize', title: '两段摘要', text: '块关系与摘要都能用。', chunkIds: ids }, exec);
    assert.equal(summary.ok, true, summary.text);
    assert.ok(summary.text.includes('已生成摘要块'), summary.text);
    const list = await tool.execute({ action: 'kb-summaries' }, exec);
    assert.equal(list.ok, true, list.text);
    assert.ok(list.text.includes('两段摘要'), list.text);
    const updated = await tool.execute({ action: 'kb-update', id: docId, folder: '参考', title: '改过的标题' }, exec);
    assert.equal(updated.ok, true, updated.text);
    assert.ok(updated.text.includes('已改'), updated.text);
    const edited = await tool.execute({ action: 'kb-edit', id: docId, text: '# 改过的正文\n\n新的正文块。' }, exec);
    assert.equal(edited.ok, true, edited.text);
    assert.ok(edited.text.includes('重新入库'), edited.text);
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
    const relA = await writeDoc(root, 'verified.md', `${SEARCH_DOC_A}${SOURCES}`);
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
    const tool = buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } })).find((item) => item.name === 'office_memory');
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

/*
 * ── 8.5 统一入口路由：`recall`（第五十七轮 19-8 的 P2-8） ────────────────────
 *
 * 门控判据本身在 memory-recall 的 D 组逐档钉住；这里测的是**编排**：
 * 门控说够答就不许再花 kb 配额、说不够就要真的下探、两边配额各守各的界
 * （换一个入口不能换来额外的检索次数），以及下探被挡住时**如实说**。
 */
await check('工具面：recall 记忆层够答 → 不下探 kb，也不吃 kb 配额', async () => {
    const { root, memory } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    const tool = buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } })).find((item) => item.name === 'office_memory');
    const events = [{ type: 'turn/start', data: { turn: 11 } }];
    const exec = { agent: { session: { id: 'recall-enough', header: { cwd: root }, ownEvents: () => events } }, signal: new AbortController().signal };
    assert.equal((await tool.execute({ action: 'kb-ingest', path: 'docs/召回实验记录.md' }, exec)).ok, true);
    // 一条**约束级**且词对得上的条目 —— 门控要判 enough 就得靠它，不是靠「有命中」。
    await memory.mutate({ action: 'add', target: 'project', content: '排序口径：先按词频饱和，再按块长归一，这条是硬约定。', importance: 'critical' });

    const out = await tool.execute({ action: 'recall', query: '排序 归一' }, exec);
    assert.equal(out.ok, true, out.text);
    assert.ok(out.text.includes('记忆层够答'), out.text);
    assert.ok(out.text.includes('未下探知识库'), out.text);
    // 路由行上的**数字**也钉住（只查子串的话，命中数 / 覆盖数写错了也看不出来）。
    assert.ok(out.text.includes('命中 1 条｜查询词覆盖 2/2｜最高重要度 critical'), out.text);
    assert.ok(!out.text.includes('kb 检索「'), '够答时不该把 kb 那段也拉进来');

    // kb 配额**一个都没动**：这一回合仍能查满两次（默认 2）。
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 BM25' }, exec)).ok, true, 'kb 配额不该被 recall 吃掉');
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 归一' }, exec)).ok, true);
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 词频' }, exec)).ok, false);
});

await check('工具面：recall 记忆层不够答 → 真的下探 kb，并吃掉一次知识块检索配额', async () => {
    const { root } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    const tool = buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } })).find((item) => item.name === 'office_memory');
    const events = [{ type: 'turn/start', data: { turn: 12 } }];
    const exec = { agent: { session: { id: 'recall-descent', header: { cwd: root }, ownEvents: () => events } }, signal: new AbortController().signal };
    assert.equal((await tool.execute({ action: 'kb-ingest', path: 'docs/召回实验记录.md' }, exec)).ok, true);

    // 记忆层里一条都没有 → empty → 下探。
    const out = await tool.execute({ action: 'recall', query: '排序 BM25' }, exec);
    assert.equal(out.ok, true, out.text);
    assert.ok(out.text.includes('一条都没命中'), out.text);
    assert.ok(out.text.includes('已下探知识库'), out.text);
    assert.ok(out.text.includes('kb 检索「排序 BM25」'), out.text);
    assert.ok(out.text.includes('相对次序'), '下探拿到的要真的是 kb 那一段回执');

    // 下探吃掉的是 kb 那一档：再查一次放行，第三次拒绝（默认 2）。
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 归一' }, exec)).ok, true);
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 词频' }, exec)).ok, false);
});

await check('工具面：recall 与 read 共用记忆检索配额；kb 被挡时如实说「没下探」而不是装作够答', async () => {
    const { root } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    const tool = buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } })).find((item) => item.name === 'office_memory');
    const sessionOf = (id, turn) => ({
        agent: { session: { id, header: { cwd: root }, ownEvents: () => [{ type: 'turn/start', data: { turn } }] } },
        signal: new AbortController().signal,
    });
    // 两个回合各测一件事：配额是按**回合**记的，混在一个回合里两件事会互相吃掉名额。
    const execA = sessionOf('recall-quota-a', 13);
    const execB = sessionOf('recall-quota-b', 14);
    assert.equal((await tool.execute({ action: 'kb-ingest', path: 'docs/召回实验记录.md' }, execA)).ok, true);

    // ① 参数坏掉的那次不占名额（与 read / kb-search 同一条纪律）。
    const bad = await tool.execute({ action: 'recall', query: '排序', tags: '发布' }, execA);
    assert.equal(bad.ok, false);
    assert.match(bad.text, /需要一个字符串数组/);
    const noQuery = await tool.execute({ action: 'recall' }, execA);
    assert.equal(noQuery.ok, false);
    assert.ok(noQuery.text.includes('需要 query'), noQuery.text);
    // ② 读的时候抛错（时间词解析不了）也要退名额 —— 退回的证据是**之后那次还能过**。
    await assert.rejects(() => tool.execute({ action: 'recall', query: '排序', since: '不是时间' }, execA), /时间/);
    const first = await tool.execute({ action: 'recall', query: '排序 BM25' }, execA);
    assert.equal(first.ok, true, '抛错的那次把记忆侧名额吃掉了');
    assert.ok(first.text.includes('已下探知识库'), first.text);
    // ③ 记忆侧还剩「细化」那一次，read 带 query 照样能过 —— 两个入口确实是同一条界。
    assert.equal((await tool.execute({ action: 'read', query: '排序' }, execA)).ok, true);
    // ④ 两档都用完了：recall 被拒时要说清「知识块检索是**另一档**、很可能还满额」，
    //    否则模型会以为整个检索都用不了、原地转圈。
    const denied = await tool.execute({ action: 'recall', query: '排序 词频' }, execA);
    assert.equal(denied.ok, false);
    assert.ok(denied.text.includes('配额已用完'), denied.text);
    assert.ok(denied.text.includes('知识块检索是另一档'), denied.text);

    // ⑤ 换一个回合：先把 kb 那一档用光，再让 recall 的门控判成 empty —— 它必须说清是
    //    **配额**挡的，并给出真的能走的那条路（kb-read / kb-list / 调设置）。
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 BM25' }, execB)).ok, true);
    assert.equal((await tool.execute({ action: 'kb-search', query: '排序 归一' }, execB)).ok, true);
    const blocked = await tool.execute({ action: 'recall', query: '排序 词频' }, execB);
    assert.equal(blocked.ok, true, blocked.text);
    assert.ok(blocked.text.includes('未下探知识库'), blocked.text);
    assert.ok(blocked.text.includes('知识块检索配额已用完'), blocked.text);
    assert.ok(blocked.text.includes('kb-read'), '被挡住时要给出真的能走的那条路');
    assert.ok(blocked.text.includes('设置页'), '放宽的入口也要给');
    assert.ok(!blocked.text.includes('记忆层够答，不必再查来源原文'), '不许把「被挡住」说成「不必查」');
});

await check('工具面：抛错的检索不占名额（配额要退回来，复核 P3 的回归）', async () => {
    const { root } = await freshRoot();
    await writeDoc(root, 'docs/召回实验记录.md', SEARCH_DOC_A);
    const tool = buildTools(resolveConfig({ memory: { kbDir: join(root, '.office/kb') } })).find((item) => item.name === 'office_memory');
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
    const rel = await writeDoc(root, 'a.md', `${DOC}${SOURCES}`);
    await memory.kbIngest({ path: rel, tier: 'verified' });
    const pack = await memory.exportPack();
    const tampered = JSON.parse(JSON.stringify(pack));
    tampered.stores[0].kb.manifest[0].tier = 'verified'; // 包里仍是 verified（这份本来就是）
    tampered.stores[0].entries = [];

    // 目标库里先有一份**同样内容**（同哈希、不同路径）的 unverified 来源。
    const { root: other, memory: target } = await freshRoot();
    const relOther = await writeDoc(other, 'b.md', `${DOC}${SOURCES}`);
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
    const rel = await writeDoc(root, 'good.md', `# 好来源\n\n两家独立媒体确认。${SOURCES}`);
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

await check('F4 单库同一份内容档位撞车：门看的是更保守的那一侧', async () => {
    // 内容寻址让两份不同路径的文档共用块 id，档位却可以不同。门必须按整份清单取更保守的
    // 那一侧，不能因为遍历顺序让后一条把前一条覆盖掉。
    const { root, memory } = await freshRoot();
    const relBad = await writeDoc(root, 'shared-bad.md', DOC);
    const relGood = await writeDoc(root, 'shared-good.md', DOC);
    await memory.kbIngest({ path: relBad, tier: 'unverified' });
    const local = await memory.kbIngest({ path: relGood });
    assert.equal(local.tier, 'user');
    const chunk = (await memory.kbRead({ id: local.docId })).chunks[0].id;
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: '结论：依据这份内容。', source: [chunk] }),
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
    const rel = await writeDoc(root, '.office/search/f11.md', `# 来源\n\n两家媒体确认。${SOURCES}`);
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

// ── 入库评估闸与台账侧质量门（第五十四轮 19-8 的 P1-4b） ────────────────────
//
// 这一组守的是「tier 是不是**可核对**的」：`verified` 从前是纯声明（拼写对就过），
// 而它是「能进热记忆当依据」的那一档。门必须真的挡得住，且挡下时要说清下一步；
// 台账侧则相反 —— 登记不许拒绝（交付物真实存在），但要把待核引用标出来。

await check('入库评估闸：verified 要数得出 ≥2 个独立来源，数不出就拒绝并给三条去处', async () => {
    const { root, memory } = await freshRoot();
    const none = await writeDoc(root, 'no-url.md', '# 无来源\n\n某论坛说会放开。');
    const message = await memory.kbIngest({ path: none, tier: 'verified' }).then(() => '', (error) => error.message);
    assert.match(message, /入库评估闸拦下了/, message);
    for (const needle of ["tier:'unverified'", "tier:'user'", '第二份独立来源']) {
        assert.ok(message.includes(needle), `拒绝的文案要给出去处「${needle}」：${message}`);
    }
    assert.equal((await memory.kbList()).total, 0, '被闸拦下的文档不许落盘');

    // 同一站点的两个页面算**一个**来源 —— 这是「独立」的定义，不许按 URL 条数放行。
    const one = await writeDoc(root, 'one-host.md', '# 单源\n\n见 https://news-a.example.com/a 与 https://news-a.example.com/b。');
    await assert.rejects(() => memory.kbIngest({ path: one, tier: 'verified' }), /只有 1 个独立来源/);

    // 形状门（复核 P2-1）：单标签的「主机名」与占位词不算来源，否则 `https://a、https://b`
    // 这种写法就能骗到 verified。www 前缀也不算第二份来源。
    const forged = await writeDoc(root, 'forged.md', '# 占位\n\n见 https://a、https://b。');
    await assert.rejects(() => memory.kbIngest({ path: forged, tier: 'verified' }), /≥2 个独立来源/,
        'https://a、https://b 这类单标签主机名不算两份独立来源');
    const forgedCn = await writeDoc(root, 'forged-cn.md', '# 占位\n\n见 https://占位一 与 https://占位二。');
    await assert.rejects(() => memory.kbIngest({ path: forgedCn, tier: 'verified' }), /≥2 个独立来源/);
    const www = await writeDoc(root, 'www.md', '# 同站\n\n见 https://www.a.example.com/a 与 https://a.example.com/b。');
    await assert.rejects(() => memory.kbIngest({ path: www, tier: 'verified' }), /只有 1 个独立来源/,
        'www 前缀与裸域名是同一个站点');
    // 「有地址但一个都解不出主机名」与「一个地址都没有」要分开说（复核 P3-1）。
    const broken = await writeDoc(root, 'broken.md', '# 坏链\n\n见 https://%%/x 与 https:// 。');
    const brokenMessage = await memory.kbIngest({ path: broken, tier: 'verified' }).then(() => '', (error) => error.message);
    assert.match(brokenMessage, /一个都解不出主机名/, brokenMessage);

    // 两个 host 就过，且回执里印出现算出来的证据面。
    const two = await writeDoc(root, 'two-hosts.md', '# 两源\n\n见 https://news-a.example.com/a 与 https://news-b.example.org/b。');
    const ok = await memory.kbIngest({ path: two, tier: 'verified' });
    assert.equal(ok.tier, 'verified');
    assert.deepEqual(ok.evidence.hosts, ['news-a.example.com', 'news-b.example.org']);
    assert.ok(renderKb(ok).includes('2 个独立来源'), renderKb(ok));
});

await check('入库评估闸：只拦「这次要变成 verified」的调用，别的档位与幂等重入放行', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'single.md', '# 单源\n\n某论坛说会放开。');

    const unverified = await memory.kbIngest({ path: rel, tier: 'unverified' });
    assert.equal(unverified.tier, 'unverified', '待核不需要证据（它没声称核实过）');
    assert.equal((await memory.kbIngest({ path: rel, tier: 'unverified' })).unchanged, true, '幂等重入放行');
    await assert.rejects(() => memory.kbIngest({ path: rel, tier: 'verified' }), /入库评估闸拦下了/,
        '内容没变但要把档位升成 verified：没有证据就不许升');

    // 历史库（门之前入库的 verified）不该被一次幂等重入追溯拦下：门是向前的。
    const paths = memoryPaths(root, {}).kb;
    const manifest = (await readFile(paths.manifest, 'utf8')).trim().split('\n')
        .map((line) => JSON.parse(line));
    manifest[0].tier = 'verified';
    await writeFile(paths.manifest, `${manifest.map((row) => JSON.stringify(row)).join('\n')}\n`, 'utf8');
    const again = await memory.kbIngest({ path: rel, tier: 'verified' });
    assert.equal(again.unchanged, true, '同内容同档位 = 幂等，不该被追溯拦下');

    // 证据够了就该放行，并指路升档通道。
    const two = await writeDoc(root, 'two-hosts.md', '# 两源\n\n见 https://news-a.example.com/a 与 https://news-b.example.org/b。');
    const staged = await memory.kbIngest({ path: two, tier: 'unverified' });
    assert.ok(renderKb(staged).includes("可以用 kb-ingest 带 tier:'verified' 升档"), renderKb(staged));
    assert.equal((await memory.kbIngest({ path: two, tier: 'verified' })).tier, 'verified');
});

await check('台账侧质量门：待核 / 找不到的来源**读时现算**并标出来，登记从不拒绝', async () => {
    const { root, memory } = await freshRoot();
    // 两份**都带得动证据**的文档（≥2 主机名），才能在同一份内容上验「升档 / 降档」——
    // 一份没有来源的文档根本升不到 verified（那正是入库闸本身）。
    const weakDoc = await writeDoc(root, 'weak.md', '# 待核\n\n见 https://w1.example.com/a 与 https://w2.example.org/b。');
    const strongDoc = await writeDoc(root, 'strong.md', '# 强\n\n见 https://s1.example.com/a 与 https://s2.example.org/b。');
    const weak = await memory.kbIngest({ path: weakDoc, tier: 'unverified' });
    const strong = await memory.kbIngest({ path: strongDoc, tier: 'verified' });
    const weakChunk = (await memory.kbRead({ id: weak.docId })).chunks[0].id;
    const strongChunk = (await memory.kbRead({ id: strong.docId })).chunks[0].id;

    await memory.log([
        { path: '汇报.pptx', format: 'ppt', purpose: '季度汇报', source: ['输入.md', weakChunk] },
        { path: '综述.pptx', format: 'ppt', purpose: '综述', source: [strongChunk] },
        { path: '草稿.pptx', format: 'ppt', source: ['kb:deadbeef0000:1'] },
    ]);

    const pick = async (name) => (await memory.read({ layer: 'ledger' })).ledger.items.find((row) => String(row.path).includes(name));
    const read = await memory.read({ layer: 'ledger' });
    const rows = read.ledger.items;
    assert.equal(rows.length, 3, '三条都要登记（拒绝登记等于把依据抹成好看的账）');
    assert.deepEqual((await pick('汇报')).untrustedRefs, [weakChunk], '待核引用要标出来');
    assert.deepEqual((await pick('综述')).untrustedRefs, [], '已验证的来源不标');
    assert.deepEqual((await pick('草稿')).unknownRefs, ['kb:deadbeef0000:1'], '找不到的块 id 另标一类');
    const text = renderRead(read);
    assert.ok(text.includes('⚠ 依据里 1 块待核'), text);
    assert.ok(text.includes('1 块在 kb 里找不到'), text);

    // 标记**不是登记时的快照**（复核 P1-2）：档位一变，下一次读就要跟着变。
    await memory.kbIngest({ path: weakDoc, tier: 'verified' });
    assert.deepEqual((await pick('汇报')).untrustedRefs, [],
        '块升成 verified 之后警告要消失（旧实现把它当恒真事实印）');
    await memory.kbIngest({ path: strongDoc, tier: 'unverified' });
    assert.deepEqual((await pick('综述')).untrustedRefs, [strongChunk],
        '块降成 unverified 之后要开始报警（同一块在 add 里过不了硬规则）');
    await memory.kbDrop({ id: weak.docId });
    assert.deepEqual((await pick('汇报')).unknownRefs, [weakChunk], 'kb 里删掉之后要说「找不到」，而不是继续印「待核」');
});

// ── 换代 / 删除的两段式提交协议（第五十六轮 S-03） ──────────────────────────
//
// 这一组守的是「换代的提交点在哪」：`manifest.jsonl`（真源）+ `index.json`（计数投影）
// 落盘才算换代发生，**提交之前一个旧字节都不删**；删除类动作（清旧边 / 墓碑 / 收旧块）
// 全在提交之后，并且先写一份可重放的意图日志（`pending.json`）。
//
// 守它的办法不是读代码，而是**真的把提交后的一步弄坏**（把 `index.json` 换成目录，
// 改名必定失败），再看盘上剩下什么：
//   · 老顺序（旧块先删、manifest 后写）在这里留下「manifest 说有、正文没有」；
//   · 新顺序留下「新版本完整 + 一份能重放的日志」。
// 另外两半：「崩在提交之前」必须只是丢掉日志（旧版本与旧边一根不动），以及「下一次
// kb 写操作会把没做完的清边补完」—— 后者是这条协议敢把清边后置的全部理由。

/** 读当前 manifest 的行（每行一个文档）。 */
async function manifestRows(root) {
    const text = await readFile(memoryPaths(root, {}).kb.manifest, 'utf8');
    return text.trim().split('\n').filter((line) => line !== '').map((line) => JSON.parse(line));
}

await check('提交协议：换代与删除的收尾在同一次调用里做完（不留日志、不留孤儿块）', async () => {
    const { root, memory } = await freshRoot();
    const paths = memoryPaths(root, {}).kb;
    const rel = await writeDoc(root, 'proto.md', DOC);
    const first = await memory.kbIngest({ path: rel });
    const v1 = (await manifestRows(root))[0].hash;
    assert.equal(existsSync(paths.pending), false, '正常入库不该留意图日志');

    await writeDoc(root, 'proto.md', `${DOC}\n\n## 第二节\n\n换代之后的新一段正文。\n`);
    const second = await memory.kbIngest({ path: rel });
    assert.equal(second.replaced, true, '同路径新内容 = 换代');
    const v2 = (await manifestRows(root))[0].hash;
    assert.notEqual(v2, v1, '块 id 按内容寻址，换了一版就是新的哈希');
    assert.equal(second.orphans, 1, '旧版本那一份块目录是被**孤儿清扫**收掉的（删除只有这一条路）');
    assert.deepEqual(await readdir(paths.chunks), [v2], '块目录里只剩这一版（没有孤儿）');
    assert.equal(existsSync(paths.pending), false, '收尾做完之后日志要删掉');
    assert.equal((await memory.danglingRefs()).length, 0);

    const dropped = await memory.kbDrop({ path: rel });
    assert.ok(dropped.tombstones >= 1, '删除要留墓碑');
    assert.deepEqual(await readdir(paths.chunks), [], '删掉之后块目录清空');
    assert.equal(existsSync(paths.pending), false);
    assert.equal((await memory.danglingRefs()).length, 0);
    assert.equal(first.docId, dropped.docId, '文档 id 按路径寻址，换代不换 id');
});

await check('提交协议：崩在提交之后（index 写不进去）→ 新版本完整 + 日志在，下一次写操作补完清边', async () => {
    const { root, memory } = await freshRoot();
    const paths = memoryPaths(root, {}).kb;
    const rel = await writeDoc(root, 'crash-post.md', DOC);
    const first = await memory.kbIngest({ path: rel });
    const v1 = (await manifestRows(root))[0].hash;
    await memory.mutate({ action: 'add', target: 'project', content: '结论：依据换代前的那一块。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    const deadChunk = (await memory.kbRead({ id: first.docId })).chunks[0].id;
    await memory.link({ sourceId: entryId, targetId: deadChunk, kind: 'supports' });
    assert.equal((await memory.danglingRefs()).length, 0);

    // 故障注入：`index.json` 换成目录 → `saveKbIndex` 的改名必定失败。抛错点落在
    // 「manifest 已提交、收尾还没跑」之间，正是这条协议要验的那个窗口。
    await rm(paths.index, { force: true });
    await mkdir(paths.index, { recursive: true });
    await writeDoc(root, 'crash-post.md', `${DOC}\n\n## 换代\n\n新的一段正文。\n`);
    const failure = await memory.kbIngest({ path: rel }).then(() => '', (error) => error?.message ?? String(error));
    assert.notEqual(failure, '', 'index 写不进去时这次调用必须报错（不能默默当成功）');

    const v2 = (await manifestRows(root))[0].hash;
    assert.notEqual(v2, v1, '提交点在 manifest：它已经是新版本');
    assert.equal(existsSync(join(paths.chunks, v2)), true, '新版本的正文完整');
    assert.equal(existsSync(join(paths.chunks, v1)), true, '崩在提交之后，旧块目录还没被收（删除都在提交之后）');
    assert.equal(existsSync(paths.pending), true, '清理意图必须留在盘上，否则这次清边就永远丢了');
    const pending = JSON.parse(await readFile(paths.pending, 'utf8'));
    assert.equal(pending.op, 'kb-ingest');
    assert.equal(pending.expectHash, v1, '判据记的是「提交前活着的那一版」');
    assert.ok(pending.dead.includes(deadChunk), '待清 id 里有指向旧块的那条边');
    // 这一瞬间的代价照实钉住：悬空边**是**存在的，但它可诊断（不是静默少一条）。
    assert.equal((await memory.danglingRefs()).length, 1, '崩在提交后的一瞬间，那条边是悬空边 —— 但它报得出来');

    // 修好 index，再跑一次 kb 写操作 → 日志被重放：清边、墓碑、收旧块一次补完。
    await rm(paths.index, { recursive: true, force: true });
    const other = await writeDoc(root, 'other.md', '# 另一份\n\n与上一份无关的正文。\n');
    const after = await memory.kbIngest({ path: other });
    assert.ok(Array.isArray(after.recoveries) && after.recoveries.length === 1, '回执要把「补做了一次清理」说出来');
    assert.equal(after.recoveries[0].replayed, true);
    assert.equal(after.recoveries[0].stage, 'post-commit');
    assert.equal(existsSync(paths.pending), false, '重放成功后日志要删掉');
    assert.equal(existsSync(join(paths.chunks, v1)), false, '旧块目录被孤儿清扫收掉');
    assert.equal((await memory.related({ id: entryId })).edges.length, 0, '提交后补做的清边必须真的清掉那条边');
    assert.equal((await memory.danglingRefs()).length, 0);
    assert.ok((await memory.status()).tombstones >= 1, '墓碑要补上（旧块 id 不再能被 link）');
});

await check('提交协议：崩在提交之前（日志在、manifest 还是旧的）→ 丢日志，旧版本与旧边都不许动', async () => {
    const { root, memory } = await freshRoot();
    const paths = memoryPaths(root, {}).kb;
    const rel = await writeDoc(root, 'crash-pre.md', DOC);
    const first = await memory.kbIngest({ path: rel });
    const v1 = (await manifestRows(root))[0].hash;
    await memory.mutate({ action: 'add', target: 'project', content: '结论：依据这一块来源。' });
    const entryId = (await memory.snapshot()).entries[0].id;
    const liveChunk = (await memory.kbRead({ id: first.docId })).chunks[0].id;
    await memory.link({ sourceId: entryId, targetId: liveChunk, kind: 'supports' });

    // 手工摆出「提交前崩掉」的现场：日志在，manifest 仍是旧的（expectHash 还活着），
    // 外加一份没人引用的孤儿块目录（提交前写进去的新块就是这个形态）。
    await mkdir(join(paths.chunks, 'deadbeef0001'), { recursive: true });
    await writeFile(join(paths.chunks, 'deadbeef0001', '1.json'), '{"id":"kb:deadbeef0001:1"}\n', 'utf8');
    await writeFile(paths.pending, `${JSON.stringify({
        version: 1,
        op: 'kb-ingest',
        at: '2026-10-06T00:00:00.000Z',
        docId: first.docId,
        expectHash: v1,
        dead: [liveChunk],
        reason: 'kb-reingest',
    })}\n`, 'utf8');

    // 同内容幂等重入 = 最小的一次 kb 写操作：它照样先重放日志。
    const again = await memory.kbIngest({ path: rel });
    assert.equal(again.unchanged, true);
    assert.equal(again.recoveries[0].replayed, false);
    assert.equal(again.recoveries[0].stage, 'pre-commit');
    assert.equal(existsSync(paths.pending), false, '没提交的意图要丢掉 —— 拿它去清边等于清掉活着的边');
    assert.equal((await memory.related({ id: entryId })).edges.length, 1, '旧版本还活着，它的边一根不许动');
    assert.equal((await memory.danglingRefs()).length, 0);
    assert.equal(existsSync(join(paths.chunks, 'deadbeef0001')), false, '孤儿块照扫（它与提交没提交无关）');
    assert.equal((await memory.status()).tombstones, 0, '没提交就不该留墓碑');
});

/** 跨根并发的公共现场：文档在全局根，两个热记忆条目与它们的边在工作区根。 */
async function crossRootScene(round) {
    caseId += 1;
    const rootWork = join(TMP, `case-${caseId}-w-${round}`);
    const global = globalFixture(caseId, `global-${round}`);
    const globalDir = global.dir;
    await rm(rootWork, { recursive: true, force: true });
    await rm(global.parent, { recursive: true, force: true });
    await mkdir(rootWork, { recursive: true });
    // 入库的那个会话只开全局根：文档落进全局根（删除时「被请求的根」是它）。
    const globalSession = createMemory({ root: rootWork, memory: { scope: 'global', globalDir } });
    const both = { scope: 'both', globalDir };
    const writer = createMemory({ root: rootWork, memory: both });
    const deleter = createMemory({ root: rootWork, memory: both });

    const rel = await writeDoc(rootWork, 'shared.md', DOC);
    const doc = await globalSession.kbIngest({ path: rel });
    const chunk = (await writer.kbRead({ id: doc.docId })).chunks[0].id;
    await writer.mutate({ action: 'add', target: 'project', content: `第 ${round} 轮的结论甲。` });
    await writer.mutate({ action: 'add', target: 'project', content: `第 ${round} 轮的结论乙。` });
    const [keeper, linker] = (await writer.snapshot()).entries;
    return { writer, deleter, rel, chunk, keeper, linker };
}

await check('多根队列：跨根清边 × 同一根上的删边并发时，清掉的边不许被旧快照写回来（第五十六轮 S-10）', async () => {
    // 这一条打的是 S-10 的**正身**：`kbDrop` 从前只锁「被请求的那个根」（全局），
    // 却要遍历所有根清边 —— 工作区根上的 `unlink` 用自己的队列，两者谁也挡不住谁。
    // 交错坏掉的样子：清边先把（清掉了 E2 的）结果写回，`unlink` 再拿它**自己更早**
    // 读到的那份（还带着 E2）覆盖一遍 → 已经判死的块 id 又有一条边指着它（悬空边）。
    const rounds = 10;
    for (let round = 0; round < rounds; round += 1) {
        const { writer, deleter, rel, chunk, keeper, linker } = await crossRootScene(round);
        const doomed = await writer.link({ sourceId: keeper.id, targetId: linker.id, kind: 'related' });
        const cross = await writer.link({ sourceId: linker.id, targetId: chunk, kind: 'supports' });
        assert.equal(cross.added, true, '跨根那条边先建好');

        const [dropped, removed] = await Promise.allSettled([
            deleter.kbDrop({ path: rel }),
            writer.unlink({ id: doomed.link.id }),
        ]);
        assert.equal(dropped.status, 'fulfilled', `第 ${round} 轮：删除本身该成功（${dropped.reason?.message ?? ''}）`);
        assert.equal(removed.status, 'fulfilled', `第 ${round} 轮：unlink 该成功（${removed.reason?.message ?? ''}）`);
        assert.equal((await writer.danglingRefs()).length, 0,
            `第 ${round} 轮留下了悬空边（跨根清边写回时被 unlink 的旧快照覆盖）`);
        assert.ok(!(await writer.related({ id: keeper.id })).edges.some((edge) => edge.id === doomed.link.id),
            `第 ${round} 轮被 unlink 的那条边复活了`);
        assert.ok((await writer.status()).tombstones >= 1, `第 ${round} 轮删除要留墓碑`);
    }
});

await check('多根队列：跨根清边 × 另一个根上的建边并发时，不许留下悬空边（第五十六轮 S-10）', async () => {
    // 另一半：`link` 从前是「排队之前先判存活，再排队写边」。判存活与写边之间只要有窗口，
    // 跨根清边（或同根的 `remove`）把这一块删掉，边就写成悬空边。现在校验在锁内，
    // 两种交错都干净：要么建边先成立（随后被一起清掉），要么校验时块已经不在 → 当场拒绝。
    const rounds = 8;
    for (let round = 0; round < rounds; round += 1) {
        const { writer, deleter, rel, chunk, keeper, linker } = await crossRootScene(round);
        await writer.link({ sourceId: keeper.id, targetId: linker.id, kind: 'related' });

        const [dropped, linked] = await Promise.allSettled([
            deleter.kbDrop({ path: rel }),
            writer.link({ sourceId: linker.id, targetId: chunk, kind: 'supports' }),
        ]);
        assert.equal(dropped.status, 'fulfilled', `第 ${round} 轮：删除本身该成功（${dropped.reason?.message ?? ''}）`);
        assert.equal((await writer.danglingRefs()).length, 0,
            `第 ${round} 轮留下了悬空边（校验与写边之间被清边插进来）`);
        assert.equal((await writer.related({ id: keeper.id })).edges.length, 1,
            `第 ${round} 轮丢了无关的那条边`);
        if (linked.status === 'fulfilled') {
            assert.equal(linked.value.added, true);
        } else {
            assert.match(String(linked.reason?.message ?? ''), /找不到|已经被删除/,
                `第 ${round} 轮 link 失败的原因应当是「块已经不在了」，实际：${linked.reason?.message}`);
        }
    }
});

// ── 虚拟文件夹（第八十四轮）──────────────────────────────────────────────

await check('虚拟文件夹：入库带 folder、树按 manifest 现算、移动与子树检索', async () => {
    const { root, memory } = await freshRoot();
    const relA = await writeDoc(root, 'contract-a.md', '# 合同 A\n\n付款周期为 30 天。');
    const relB = await writeDoc(root, 'policy-b.md', '# 制度 B\n\n报销需要两张凭证。');
    const a = await memory.kbIngest({ path: relA, folder: '参考/合同' });
    const b = await memory.kbIngest({ path: relB, folder: '参考/制度' });
    assert.equal(a.folder, '参考/合同');
    assert.equal(b.folder, '参考/制度');
    // 根、中间层、叶子都从 manifest 现算；计数含子树。
    const list = await memory.kbList({ folder: '参考' });
    assert.equal(list.total, 2);
    assert.deepEqual(list.items.map((item) => item.folder).sort(), ['参考/制度', '参考/合同'].sort());
    const tree = list.tree;
    assert.equal(tree.docs, 2);
    assert.equal(tree.children.length, 1);
    assert.equal(tree.children[0].name, '参考');
    assert.equal(tree.children[0].docs, 2);
    assert.equal(tree.children[0].children.length, 2);
    // 只算直接子层时，参考下没有直接文档（两份都在子文件夹里）。
    assert.equal((await memory.kbList({ folder: '参考', recursive: false })).total, 0);
    // 同内容重入但改 folder = 移动，不重写块。
    const moved = await memory.kbIngest({ path: relA, folder: '归档/2026' });
    assert.equal(moved.unchanged, true);
    assert.equal(moved.moved, true);
    assert.equal(moved.folder, '归档/2026');
    assert.equal((await memory.kbList({ folder: '参考' })).total, 1);
    // 检索范围跟着文件夹走：制度只在「参考」子树里，合同已经移到归档。
    const scoped = await memory.kbSearch({ query: '凭证 报销', folder: '参考' });
    assert.equal(scoped.hits.length, 1);
    assert.equal(scoped.hits[0].path, relB);
    const outside = await memory.kbSearch({ query: '付款周期', folder: '参考' });
    assert.equal(outside.hits.length, 0, '子树检索不该捞到子树之外的文档');
    assert.equal((await memory.kbSearch({ query: '付款周期', folder: '归档' })).hits.length, 1);
    // 坏路径当场报错，不静默落到根。
    await assert.rejects(() => memory.kbIngest({ path: relB, folder: '../逃出去' }), /不能有「..」/);
});

await check('单库：记忆有全局 / 工作区两个根，知识库仍只有一个库', async () => {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    const global = globalFixture(caseId, 'global');
    await rm(root, { recursive: true, force: true });
    await rm(global.parent, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    const memory = createMemory({ root, memory: { scope: 'both', globalDir: global.dir } });
    const relG = await writeDoc(root, 'global.md', '# 全局资料\n\n跨项目通用的一份材料。');
    const relW = await writeDoc(root, 'workspace.md', '# 工作区资料\n\n只属于这个项目的一份材料。');
    const g = await memory.kbIngest({ path: relG });
    const w = await memory.kbIngest({ path: relW, scope: 'workspace' });
    assert.equal(g.origin, 'kb');
    assert.equal(w.origin, 'kb');
    const all = await memory.kbList();
    assert.equal(all.total, 2);
    assert.equal(all.tree.docs, 2);
    assert.deepEqual(all.items.map((item) => item.origin), ['kb', 'kb']);
    // origin 过滤只服务热记忆 / 台账；知识库不再按来源切分。
    assert.equal((await memory.kbList({ origin: 'global' })).total, 2);
});

await check('指定全局库目录：kbDir 生效，SQLite 在首次入库后就落盘（第八十五轮）', async () => {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    await rm(root, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    const customKb = join(root, 'custom-kb');
    const memory = createMemory({
        root,
        memory: { scope: 'both', globalDir: join(root, 'global-memory'), kbDir: customKb },
    });
    const rel = await writeDoc(root, '资料.md', '# 指定目录\n\n这份材料应该进指定的全局库。');
    const value = await memory.kbIngest({ path: rel });
    assert.equal(value.origin, 'kb', '单库不再按记忆根分来源');
    assert.ok(existsSync(join(customKb, 'kb.sqlite')), '首次入库后 SQLite 索引要真的存在');
    assert.ok(existsSync(join(customKb, 'manifest.jsonl')), 'manifest 也要落在指定目录');
});

await check('归档并入知识库：下沉条目以只读虚拟文档出现在「归档」文件夹', async () => {
    const { root, memory } = await freshRoot({ projectLimitBytes: 300, archiveKeep: 10 });
    await memory.mutate({ action: 'add', target: 'project', content: `重要约定：${'重'.repeat(40)}`, importance: 'critical' });
    for (let i = 0; i < 6; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `会下沉 ${i}：${'沉'.repeat(30)}`, importance: 'low' });
    }
    const list = await memory.kbList({ folder: '归档' });
    assert.ok(list.total >= 1, '下沉过的条目应当出现在知识库的归档文件夹里');
    const row = list.items.find((item) => item.path.startsWith('归档/'));
    assert.ok(row !== undefined, '归档行要落在 归档/YYYY-MM 下');
    assert.equal(row.synthetic, true, '归档文档是现算的只读投影');
    assert.match(row.folder, /^归档\/\d{4}-\d{2}$/);
    assert.ok(list.tree.children.some((node) => node.name === '归档'), '树里也要有归档这一层');
    // 文档清单与正文都能下钻，正文就是归档条目的原文。
    const doc = await memory.kbDoc({ id: row.id });
    assert.equal(doc.chunks_total, 1);
    assert.equal(doc.chunks.length, 1);
    const read = await memory.kbRead({ id: row.id });
    assert.equal(read.chunks_total, 1);
    assert.ok(read.chunks[0].text.includes('会下沉'), '归档投影要能读到原条目内容');
    const hit = await memory.kbSearch({ query: '会下沉', folder: '归档' });
    assert.ok(hit.hits.some((item) => item.docId === row.id), '文件夹范围检索要能命中归档投影');
    // 关掉归档层时，知识库里也不该再出现归档文件夹与投影文档。
    const off = createMemory({ root, memory: { layers: { archive: false } } });
    const hidden = await off.kbList();
    assert.ok(!hidden.tree.children.some((node) => node.name === '归档'), '关掉归档层后树里不该有归档');
    assert.equal(hidden.items.some((item) => item.synthetic === true), false);
});

// ── 编辑、块关系与摘要（第八十五轮）────────────────────────────────────────

await check('编辑与块关系：kb-update / kbEdit / kbLink / kbRelated / kbSummarize', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'handbook.md', [
        '# 手册', '', '第一段：入库之后还能改。',
        '', '## 第二段', '', '第二段：块之间可以建边。',
        '', '## 第三段', '', '第三段：多块可以合成摘要。',
    ].join('\n'));
    const doc = await memory.kbIngest({ path: rel, tier: 'user', folder: '参考' });
    const updated = await memory.kbUpdate({ id: doc.docId, title: '手册（新）', folder: '参考/规范', tier: 'unverified' });
    assert.equal(updated.title, '手册（新）');
    assert.equal(updated.folder, '参考/规范');
    assert.equal(updated.tier, 'unverified');
    const listed = await memory.kbList({ folder: '参考' });
    assert.equal(listed.items[0].title, '手册（新）', '清单由 manifest 现算，标题要跟着改');

    const edited = await memory.kbEdit({ id: doc.docId, text: [
        '# 手册（改后）', '', '新的第一段：正文可以整体替换。',
        '', '## 新的第二段', '', '新的第二段：块也会跟着重切。',
    ].join('\n') });
    assert.equal(edited.action, 'kb-edit');
    assert.equal(edited.edited, true);
    assert.equal(edited.docId, doc.docId, '文档 id 按路径算，改正文不换 id；换的是内容哈希');
    const read = await memory.kbRead({ id: edited.docId });
    assert.ok(read.chunks.length >= 2, '三节文档应切出至少两块');
    const [first, second] = read.chunks;

    const linked = await memory.kbLink({ sourceId: first.id, targetId: second.id, kind: 'supports', note: '相邻' });
    assert.equal(linked.action, 'kb-link');
    const related = await memory.kbRelated({ chunkId: first.id });
    assert.equal(related.total, 1);
    assert.equal(related.items[0].chunkId, second.id);
    assert.equal(related.items[0].kind, 'supports');
    await memory.kbUnlink({ sourceId: first.id, targetId: second.id, kind: 'supports' });
    assert.equal((await memory.kbRelated({ chunkId: first.id })).total, 0);

    const summary = await memory.kbSummarize({ title: '两份块的意思', text: '编辑与关系都可以用。', chunkIds: [first.id, second.id] });
    assert.equal(summary.action, 'kb-summarize');
    const summaries = await memory.kbSummaries();
    assert.equal(summaries.length, 1);
    assert.equal(summaries[0].id, summary.id);
    assert.deepEqual(summaries[0].chunkIds, [first.id, second.id]);
});

await check('混合检索：SQLite + 本地向量也能给出引擎标记（第八十五轮）', async () => {
    const { root, memory } = await freshRoot({ embedding: { mode: 'local' } });
    const rel = await writeDoc(root, 'vector.md', '# 向量检索\n\n复盘要能复跑，指标要能对齐。');
    const doc = await memory.kbIngest({ path: rel });
    const result = await memory.kbSearch({ query: '复盘 复跑' });
    assert.equal(result.engine, 'sqlite-hybrid');
    assert.ok(result.hits.length >= 1, '本地向量模式仍应返回命中');
    assert.equal(result.embedding.mode, 'local');

    // 摘要也要走混合通道，不能只在纯词法路径里可见。
    const read = await memory.kbRead({ id: doc.docId });
    const summary = await memory.kbSummarize({ title: '摘要命中', text: '语义向量与咖啡萃取。', chunkIds: [read.chunks[0].id] });
    const hybridSummary = await memory.kbSearch({ query: '语义向量' });
    assert.ok(hybridSummary.hits.some((item) => item.docId === 'summary:' + summary.id), '混合检索也要命中摘要块');
});

await check('摘要块参与检索：kb-search 命中摘要，kb-read / kb-doc 能读摘要（第八十五轮）', async () => {
    const { root, memory } = await freshRoot();
    const rel = await writeDoc(root, 'summary-source.md', '# 原文\n\n原文只讲年度计划。');
    const doc = await memory.kbIngest({ path: rel, folder: '参考' });
    const read = await memory.kbRead({ id: doc.docId });
    const summary = await memory.kbSummarize({
        title: '咖啡萃取摘要', text: '量子退相干与咖啡萃取的综合结论。', chunkIds: [read.chunks[0].id],
    });
    const hit = await memory.kbSearch({ query: '量子退相干' });
    const summaryHit = hit.hits.find((item) => item.docId === 'summary:' + summary.id);
    assert.ok(summaryHit, '摘要里的词必须能检索到摘要块');
    assert.equal(summaryHit.chunkId, 'summary:' + summary.id + ':1');
    assert.equal(summaryHit.path, '摘要/' + summary.id + '.md');

    const summaryRead = await memory.kbRead({ id: summaryHit.docId });
    assert.equal(summaryRead.synthetic, true);
    assert.equal(summaryRead.chunks.length, 1);
    assert.equal(summaryRead.chunks[0].text, '量子退相干与咖啡萃取的综合结论。');
    assert.deepEqual(summaryRead.summaryOf, [read.chunks[0].id]);
    // 「摘要盖的是哪几块」必须在摘要本身上读得到：来源块 id / 文档 / 块号 / 字符区间。
    assert.equal(summaryRead.sources.length, 1, '摘要要带来源块清单');
    assert.equal(summaryRead.sources[0].id, read.chunks[0].id);
    assert.equal(summaryRead.sources[0].docId, doc.docId);
    assert.equal(summaryRead.sources[0].docPath, doc.path);
    assert.equal(summaryRead.sources[0].n, 1);
    assert.equal(summaryRead.sourcesMissing, 0);

    const summaryDoc = await memory.kbDoc({ id: summaryHit.docId });
    assert.equal(summaryDoc.synthetic, true);
    assert.equal(summaryDoc.text, '量子退相干与咖啡萃取的综合结论。');
    assert.equal(summaryDoc.chunks_total, 1);
    assert.deepEqual(summaryDoc.summaryOf, [read.chunks[0].id]);
    assert.deepEqual(summaryDoc.sources.map((item) => item.id), [read.chunks[0].id]);
});

await check('多格式入库：图片落资产目录，正文留可检索的图片块（第八十五轮）', async () => {
    const { root, memory } = await freshRoot();
    // 1x1 透明 PNG：只验路径与字节，不验像素。
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');
    await writeFile(join(root, '像素.png'), png);
    const receipt = await memory.kbIngest({ path: '像素.png' });
    assert.equal(receipt.kind, 'image');
    assert.equal(receipt.images.length, 1, '图片本身要作为资产收下');
    const assetRel = receipt.images[0];
    assert.match(assetRel, /^assets\//);
    const kbDir = memoryPaths(root, memory).kb.dir;
    assert.ok(existsSync(join(kbDir, assetRel)), '资产文件要真的落盘');
    const read = await memory.kbRead({ id: receipt.docId });
    assert.match(read.chunks[0].text, /!\[像素\.png\]\(assets\//, '正文里要有指向资产的图片块');
    const hit = await memory.kbSearch({ query: '像素' });
    assert.ok(hit.hits.length >= 1, '图片块也要能被检索到');
});

// ── 收尾 ──────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    if (item.ok) console.log(`PASS  ${item.name}`);
    else console.log(`FAIL  ${item.name}  → ${item.error?.message ?? item.error}`);
}
console.log(`\nkb: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
