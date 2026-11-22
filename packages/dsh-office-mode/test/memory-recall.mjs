/**
 * 召回自测：把「记忆到底能不能被查到」钉成可复跑的用例。
 *
 * 为什么单独一个套件：memory.mjs 测的是记忆的**契约**（存储、写入、生命周期、接线），
 * 而这里测的是**检索效果**——中文查询能不能命中、`layer:'all'` 时三层能不能都被看到。
 * 第二十四轮之前的实测基线是：16 条人工标注中文查询在真实台账上 hit@3 只有 8/16、8 条零结果；
 * `layer:'all'`（不带参数，也就是默认调用）下台账与归档**恒返回 0 条**。这两条现在都修了，
 * 这个套件就是防它们回归的那道闸。
 *
 * 用例分两组：
 *   A 组（预算与渲染）：三层额度、`all` 视图下三层都非空、渲染不许把「没给」说成「空」；
 *   B 组（中文召回）：17 条标注查询 × hit@3，覆盖整串命中 / 词序颠倒 / 合称 /
 *      中英混排 / 黏连 / 单字回落 / 零结果。
 *
 * 跑法：node test/memory-recall.mjs
 */
import { mkdir, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { createMemory, GATE_VERDICT_TEXT, gateVerdictOf, READ_LIMITS, renderRead } from '../src/memory.js';

const TMP = fileURLToPath(new URL('../../../.office/runtime/tmp/memory-recall/', import.meta.url));

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
    return { root, memory: createMemory({ root, memory: memoryConfig }) };
}

await rm(TMP, { recursive: true, force: true });

/**
 * 造一条约 310 字的热记忆条目。
 *
 * 长度是**算出来的**，不是随手写的，这里有两个反直觉的约束：
 *
 * 1. **必须短于 `itemChars`（400）**。理由是容量与读取额度的关系：热记忆两个目标的上限合计
 *    14,336 字节（project 10,240 + user 4,096），中文按 3 字节/字算只能装约 4,700 字。
 *    如果单条 ≥ 400 字，每条会被 `admit()` 截到 400，而容量上限只容得下 9 条 —— 9 × 400 =
 *    3,600 字，**够不到 4,000 字的读取额度**，旧实现里剩余额度还剩 400 字，台账照样拿得到
 *    条目，P0 就复现不出来（我第一版夹具就是这样，回滚探针没红）。
 *    单条压到 310 字，同样字节数能装下 13 条 —— 13 × 310 ≈ 4,000 字，刚好把读取额度吃满。
 * 2. **数量要够**，所以 project 与 user 都要填：只填 project 撑不满（见上）。
 *
 * 这条夹具的判据不是「看起来够长」，而是：**旧实现下热记忆恰好用满 4,000 字、台账与归档
 * 各 0 条**（`.office/runtime/tmp/r24-impl/debug-starve.mjs` 在两版代码上分别量过）。
 */
const longEntry = (n) => {
    let text = `第 ${n} 条项目约定：`;
    const body = '这里是一段刻意写长的正文，用来把热记忆层的字符额度吃满，好验证台账与归档不会被它挤掉。';
    while (text.length < 300) text += body;
    return text;
};

/** 把两个目标一起填到接近上限 —— 这是复现 P0 的前提，只填一个目标复现不出来。 */
async function fillHot(memory) {
    for (let i = 0; i < 20; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: longEntry(i), importance: 'critical' });
    }
    for (let i = 0; i < 10; i += 1) {
        await memory.mutate({ action: 'add', target: 'user', content: `偏好第 ${i} 条：${longEntry(i)}`, importance: 'critical' });
    }
}

/** 一个 20 条的真实感台账：路径与用途都是中文，含词序颠倒与合称的坑位。 */
const LEDGER = [
    ['docs/第十九轮-记忆系统结构优化与知识库结合.md', '记忆系统结构优化与知识库结合的研究报告'],
    ['docs/第二十二轮-音频与视频内容提取.md', '把音频与视频的内容提取做进办公插件'],
    ['docs/第二十三轮-音视频能力缺失的根因与恢复.md', '办公模式无法解析音视频的根因与恢复动作'],
    ['docs/第二十一轮-语音输入SenseVoice能否用于音频处理.md', '语音输入那条链能否当音频转写引擎'],
    ['docs/第十八轮-提示词与工具调用准确性与有效性策略研究.md', '提示词与工具调用的准确性与有效性策略'],
    ['docs/第十二轮-记忆面板视觉重做.md', '记忆面板视觉重做与令牌层'],
    ['docs/第十二轮-记忆面板-宽屏.png', '第十二轮实测截图：面板宽屏与图谱并排'],
    ['packages/dsh-office-mode/lib/client.js', '浏览器半侧单文件 bundle：记忆浏览面板的视觉重做'],
    ['docs/第十一轮-办公插件恢复Python计算与绘图.md', '为什么落点是 office.python、六条实现硬约束'],
    ['packages/dsh-office-mode/src/python.js', 'office.python 的实现：解释器探测 / run / file / check'],
    ['.office/runtime/tmp/python-e2e/Python计算与绘图.pptx', '端到端样例：numpy 计算 + matplotlib 中文图嵌进 PPT'],
    ['办公模式示例/季度回顾.docx', '第十轮用修复后的引擎重新生成的样例'],
    ['办公模式示例/经营回顾.pptx', '第十轮重新生成的样例'],
    ['办公模式示例/分月明细.xlsx', '第十轮重新生成的样例'],
    ['docs/第二十三轮-隐私清理与发布.md', '发布前的隐私清理：宿主锚点、用户名、真实邮箱都要清掉'],
    ['docs/第二十轮-检索通道多路化与设置页滚动修复.md', '检索通道多路化与设置页滚动修复'],
    ['docs/第十七轮-提示词注入瘦身与缓存归因.md', '提示词注入瘦身与缓存归因'],
    ['docs/第十四轮-热记忆投影与容量口径.md', '热记忆投影只在变化时贴全文'],
    ['docs/第六轮-记忆层.md', '三层记忆与容量下沉'],
    ['docs/第五轮-PDF与缓存.md', 'PDF 读取与缓存 TTL'],
];

/**
 * 标注查询集：query → 正确答案的**路径**（用路径当 id，读起来比 L-xxxx 清楚）。
 * 右列写的是这条查询属于哪一类失败，方便回归时看出坏在哪一类。
 */
const LABELED = [
    ['记忆面板', ['docs/第十二轮-记忆面板视觉重做.md', 'docs/第十二轮-记忆面板-宽屏.png', 'packages/dsh-office-mode/lib/client.js'], '整串命中（基线）'],
    ['面板记忆', ['docs/第十二轮-记忆面板视觉重做.md', 'docs/第十二轮-记忆面板-宽屏.png'], '词序颠倒'],
    ['结合知识库', ['docs/第十九轮-记忆系统结构优化与知识库结合.md'], '词序颠倒'],
    ['知识库结合', ['docs/第十九轮-记忆系统结构优化与知识库结合.md'], '整串命中'],
    ['音视频', ['docs/第二十二轮-音频与视频内容提取.md', 'docs/第二十三轮-音视频能力缺失的根因与恢复.md'], '合称 vs 全称（正文写「音频与视频」）'],
    ['音视频提取', ['docs/第二十二轮-音频与视频内容提取.md'], '合称 + 词序'],
    ['音视频根因', ['docs/第二十三轮-音视频能力缺失的根因与恢复.md'], '词序颠倒'],
    ['语音转写', ['docs/第二十一轮-语音输入SenseVoice能否用于音频处理.md'], '改写（正文写「语音输入」「音频转写」）'],
    ['提示词准确性', ['docs/第十八轮-提示词与工具调用准确性与有效性策略研究.md', 'docs/第十七轮-提示词注入瘦身与缓存归因.md'], '词序 + 拆分'],
    ['隐私清理', ['docs/第二十三轮-隐私清理与发布.md'], '整串命中'],
    ['python 绘图', ['docs/第十一轮-办公插件恢复Python计算与绘图.md', '.office/runtime/tmp/python-e2e/Python计算与绘图.pptx'], '中英混排'],
    ['python绘图', ['docs/第十一轮-办公插件恢复Python计算与绘图.md', '.office/runtime/tmp/python-e2e/Python计算与绘图.pptx'], '中英黏连（必须分段处理）'],
    ['计算与绘图', ['docs/第十一轮-办公插件恢复Python计算与绘图.md', '.office/runtime/tmp/python-e2e/Python计算与绘图.pptx'], '整串命中'],
    ['季度回顾', ['办公模式示例/季度回顾.docx'], '整串命中'],
    ['样例 重新生成', ['办公模式示例/季度回顾.docx', '办公模式示例/经营回顾.pptx', '办公模式示例/分月明细.xlsx'], '多词，各自命中'],
    ['结构记忆系统', ['docs/第十九轮-记忆系统结构优化与知识库结合.md'], '词序颠倒'],
    ['面板 宽屏', ['docs/第十二轮-记忆面板-宽屏.png'], '多词，各自命中'],
    ['检索通道', ['docs/第二十轮-检索通道多路化与设置页滚动修复.md'], '整串命中'],
];

async function seedLedger(memory) {
    for (const [path, purpose] of LEDGER) {
        await memory.log([{ path, format: 'md', purpose }]);
    }
}

const TOPK = 3;

// ── A 组：预算与渲染 ────────────────────────────────────────────────────────

await check('A1 默认读取（layer:all 无 query）：三层都要有东西，台账与归档不许是 0 条', async () => {
    const { root, memory } = await freshRoot({ ledgerLimit: 8 });
    await fillHot(memory);
    await seedLedger(memory);
    // 让台账滚出几条进归档：把上限压到 6，20 条会有 14 条下沉。
    const rolled = createMemory({ root, memory: { ledgerLimit: 6 } });
    await rolled.log([{ path: '占位.md', format: 'md', purpose: '触发滚动' }]);

    const r = await memory.read({ layer: 'all' });
    assert.ok(r.hot.items.length > 0, '热记忆要有条目');
    assert.ok(r.ledger.total > 0, '台账有 20 条');
    assert.ok(r.ledger.items.length > 0, '台账不许是 0 条（这就是第二十四轮修的 P0）');
    assert.ok(r.archive.items_total > 0, '归档有下沉条目');
    assert.ok(r.archive.items.length > 0, '归档不许是 0 条（这就是第二十四轮修的 P0）');

    const text = renderRead(r);
    assert.ok(!/【台账】共 \d+ 条\n- （空）/.test(text), '台账有内容时不许渲染成「（空）」');
    // 每层都要报出自己的额度，出问题时才知道是谁的份额。
    assert.equal(r.ledger.budgetChars, READ_LIMITS.allChars.ledger);
    assert.equal(r.archive.budgetChars, READ_LIMITS.allChars.archive);
});

await check('A2 默认读取带 query：命中非 0 时台账与归档仍要给条目', async () => {
    const { root, memory } = await freshRoot({ ledgerLimit: 6 });
    await fillHot(memory);
    await seedLedger(memory);

    const r = await memory.read({ layer: 'all', query: '记忆面板' });
    assert.ok(r.ledger.matched > 0, '「记忆面板」应当命中台账');
    assert.ok(r.ledger.items.length > 0, '命中之后要给条目，不许恒 0 条');
    assert.ok(r.archive.matched > 0, '「记忆面板」应当命中归档（下沉的旧台账）');
    assert.ok(r.archive.items.length > 0, '命中之后要给条目，不许恒 0 条');
});

await check('A3 三层总额度有上界：不会因为「各给各的」变成无界注入', async () => {
    const { root, memory } = await freshRoot({ ledgerLimit: 4 });
    for (let i = 0; i < 14; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: longEntry(i), importance: 'critical' });
    }
    await seedLedger(memory);

    const r = await memory.read({ layer: 'all' });
    const chars = (items) => items.reduce((n, item) => n + String(item.text).length, 0);
    const used = chars(r.hot.items) + chars(r.ledger.items) + chars(r.archive.items);
    const ceiling = READ_LIMITS.allChars.hot + READ_LIMITS.allChars.ledger + READ_LIMITS.allChars.archive;
    assert.ok(used <= ceiling, `三层合计 ${used} 字不该超过上界 ${ceiling}`);
    // 每层各自的上界也要成立 —— 这是「谁也不能吃别人的」的机械保证。
    assert.ok(chars(r.hot.items) <= READ_LIMITS.allChars.hot, '热记忆不超过自己的额度');
    assert.ok(chars(r.ledger.items) <= READ_LIMITS.allChars.ledger, '台账不超过自己的额度');
    assert.ok(chars(r.archive.items) <= READ_LIMITS.allChars.archive, '归档不超过自己的额度');
});

await check('A4 单层读取仍独占 totalChars（没有被 all 的额度改动拖累）', async () => {
    const { root, memory } = await freshRoot();
    for (let i = 0; i < 14; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: longEntry(i), importance: 'critical' });
    }
    const hot = await memory.read({ layer: 'hot' });
    assert.equal(hot.hot.budgetChars, READ_LIMITS.totalChars, '单层读取的额度是 totalChars');
    const all = await memory.read({ layer: 'all' });
    assert.equal(all.hot.budgetChars, READ_LIMITS.allChars.hot, 'all 视图里热记忆用自己的额度');
    assert.ok(hot.hot.items.length >= all.hot.items.length, '单层读取给的不该比 all 视图少');
});

await check('A6 渲染的四种措辞：切片 / 命中却没给 / 没命中 / 真的空', async () => {
    // 这一条直接喂 renderRead 合成结果：它是纯函数，比绕夹具准，也更清楚在测什么。
    const item = { id: 'L-1', at: '2026-09-27T00:00:00.000Z', path: 'a.md', purpose: 'p', format: 'md', theme: '' };
    const base = { layer: 'all', query: '', truncated: false, paths: { memory: '.office/memory' }, qualityDroppedTotal: 0 };

    // (a) 切片：给了 2 条、命中 20 条 → 必须说清给了多少、额度多少、怎么看全部。
    const sliced = renderRead({
        ...base,
        ledger: { items: [item, item], total: 20, matched: 20, budgetChars: 2000 },
    });
    assert.ok(sliced.includes('本层只给了 2/20 条'), '切片要说清给了几条');
    assert.ok(sliced.includes("layer:'ledger'"), '切片要给出看全部的办法');

    // (b) 命中却没给：不许说「空」，必须说清是额度的问题。
    const squeezed = renderRead({
        ...base,
        query: '面板',
        ledger: { items: [], total: 20, matched: 20, budgetChars: 1200 },
    });
    assert.ok(!squeezed.includes('- （空）'), '命中 20 条时不许说「空」');
    assert.ok(squeezed.includes('额度'), '要说清是额度被占满');

    // (c) 没命中：说清换关键词，而不是「空」。
    const missed = renderRead({
        ...base,
        query: '量子纠缠',
        ledger: { items: [], total: 20, matched: 0, budgetChars: 1200 },
    });
    assert.ok(missed.includes('没有命中的条目'), '零命中要给出下一步');
    assert.ok(!missed.includes('- （空）'), '有 20 条在手时不许说「空」');

    // (d) 真的空：这才该说「（空）」。
    const empty = renderRead({
        ...base,
        ledger: { items: [], total: 0, matched: 0, budgetChars: 1200 },
    });
    assert.ok(empty.includes('- （空）'), '确实没有记录时说「空」是对的');

    // (e) 热记忆切片：也要明说给了几条。
    const hotSliced = renderRead({
        ...base,
        hot: {
            items: [{ target: 'user', importance: 'critical', text: '一条偏好', origin: 'workspace' }],
            total: 16,
            budgetChars: 4000,
            targets: { user: { entryCount: 1, used: 12, limit: 4096 }, project: { entryCount: 0, used: 0, limit: 10240 } },
        },
    });
    assert.ok(hotSliced.includes('16 条里给了 1 条'), '热记忆切片要明说');
});

await check('A5 真实读取的零命中措辞：有 20 条在手时不说「空」', async () => {
    const { root, memory } = await freshRoot();
    await seedLedger(memory);

    const miss = await memory.read({ layer: 'ledger', query: '量子纠缠' });
    assert.equal(miss.ledger.matched, 0);
    const missText = renderRead(miss);
    assert.ok(missText.includes('没有命中的条目'), '零命中的措辞');
    assert.ok(!missText.includes('- （空）'), '有 20 条在手时不许说「空」');

    // 把额度压到 0，构造「命中却有给不出」的情形：渲染必须说出真实原因。
    const squeezed = createMemory({ root, memory: { ledgerLimit: 500, recallQuality: { maxMediumResults: 0, maxUnknownResults: 0 } } });
    const hitButEmpty = await squeezed.read({ layer: 'ledger', query: '面板 甲词 乙词 丙词 丁词' });
    if (hitButEmpty.ledger.matched > 0 && hitButEmpty.ledger.items.length === 0) {
        const t = renderRead(hitButEmpty);
        assert.ok(t.includes('额度') || t.includes('没有命中'), '给不出条目要说清原因');
        assert.ok(!/【台账】共 \d+ 条\n- （空）/.test(t), '不许把「没给」说成「空」');
    }
});

// ── B 组：中文召回 ─────────────────────────────────────────────────────────

await check('B1 标注查询集：18 条中文查询的 hit@3 与零结果数', async () => {
    const { root, memory } = await freshRoot();
    await seedLedger(memory);

    const rows = [];
    for (const [query, want, kind] of LABELED) {
        const r = await memory.read({ layer: 'ledger', query });
        const got = r.ledger.items.slice(0, TOPK).map((item) => item.path);
        const hit = got.filter((path) => want.includes(path)).length;
        rows.push({ query, kind, hit: hit > 0 ? 1 : 0, recall: hit / want.length, shown: r.ledger.items.length });
    }
    const hits = rows.reduce((n, r) => n + r.hit, 0);
    const zeros = rows.filter((r) => r.shown === 0).length;
    const recall = rows.reduce((n, r) => n + r.recall, 0) / rows.length;
    const failed = rows.filter((r) => r.hit === 0).map((r) => `${r.query}（${r.kind}）`);

    assert.equal(hits, rows.length, `hit@3 必须全中，未中：${failed.join('、')}`);
    assert.equal(zeros, 0, '不许有零结果查询');
    assert.ok(recall >= 0.9, `recall@3 应当 ≥ 0.9，实际 ${recall.toFixed(3)}`);
    console.log(`      标注查询集：hit@3 ${hits}/${rows.length}，recall@3 ${recall.toFixed(3)}，零结果 ${zeros}`);
});

await check('B2 切词边界：标点分隔、大小写、单字回落、纯英文', async () => {
    const { root, memory } = await freshRoot();
    await memory.log([
        { path: 'a.md', format: 'md', purpose: '记忆面板，宽屏，深色' },
        { path: 'b.md', format: 'md', purpose: 'CaseSensitive Token' },
    ]);

    const punct = await memory.read({ layer: 'ledger', query: '记忆，面板' });
    assert.equal(punct.ledger.matched, 1, '中文逗号也要当分隔符');

    const upper = await memory.read({ layer: 'ledger', query: 'CASESENSITIVE' });
    assert.equal(upper.ledger.matched, 1, '英文大小写不敏感');

    const single = await memory.read({ layer: 'ledger', query: '深' });
    assert.equal(single.ledger.matched, 1, '单字回落成整串匹配');

    const latin = await memory.read({ layer: 'ledger', query: 'token' });
    assert.equal(latin.ledger.matched, 1, '纯英文整词匹配');
});

// ── C 组：读侧主题分区（第五十四轮 新-12） ──────────────────────────────────
//
// 守的是「tags / entities 从**只参与匹配**变成**参与隔离**」这条改动：从前正文里出现
// 关键词就算命中（软），现在没**显式声明**过这些标签 / 实体的条目一律不出现（硬）。
// 这条判据最容易退化的三处：只判热记忆、把合取写成「任一命中」、把分区当排序键。

await check('C1 分区是硬过滤：正文里有那个词但没声明过标签 / 实体的条目不出现', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '发布流程：先过评审再发。', tags: ['发布'], entities: ['发布流程'] });
    await memory.mutate({ action: 'add', target: 'project', content: '这条正文里也写着发布两个字，但没声明标签。', tags: ['财务'] });
    await memory.mutate({ action: 'add', target: 'project', content: '既没有标签也没有实体的条目。' });

    const all = await memory.read({ layer: 'hot' });
    assert.equal(all.hot.items.length, 3, '不给分区就是老行为：全给');
    assert.equal(all.partition, undefined, '没给分区时结果里不该有 partition');

    const published = await memory.read({ layer: 'hot', tags: ['发布'] });
    assert.equal(published.hot.items.length, 1, '只给**声明过**该标签的那一条（正文命中不算）');
    assert.ok(published.hot.items[0].text.includes('发布流程：先过评审'), '命中的是声明过标签的那条');
    assert.equal(published.hot.partitionFiltered, 2, '隔掉的条数要如实报');
    assert.equal(published.hot.total, 3, 'total 是**分区前**的条数（与台账 / 归档同口径）');
    assert.deepEqual(published.partition.tags, ['发布']);

    const byEntity = await memory.read({ layer: 'hot', entities: ['发布流程'] });
    assert.equal(byEntity.hot.items.length, 1, '实体分区同理');

    const miss = await memory.read({ layer: 'hot', tags: ['不存在的分区'] });
    assert.equal(miss.hot.items.length, 0);
    const text = renderRead(miss);
    assert.match(text, /^分区：标签 不存在的分区/m, '回执首部要自报分区口径（谁隔掉了条目）');
    assert.ok(text.includes('没有声明所给的分区标签/实体'), '一条都没给时要说是分区挡的，不是「记忆里没有」');
});

await check('C2 多个词是合取（都要声明过），大小写与首尾空白归一、重复键去重', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '两个标签都有的条目。', tags: ['发布', '模板'] });
    await memory.mutate({ action: 'add', target: 'project', content: '只有发布标签的条目。', tags: ['发布'] });
    await memory.mutate({ action: 'add', target: 'project', content: '大小写不同的标签。', tags: ['Release'] });

    const both = await memory.read({ layer: 'hot', tags: ['发布', '模板'] });
    assert.equal(both.hot.items.length, 1, '两个标签都要声明过（合取，不是「任一命中」）');

    const one = await memory.read({ layer: 'hot', tags: ['发布'] });
    assert.equal(one.hot.items.length, 2, '只给一个标签时给两条');

    const lower = await memory.read({ layer: 'hot', tags: [' release '] });
    assert.equal(lower.hot.items.length, 1, '分区键去首尾空白 + 大小写不敏感');

    const dup = await memory.read({ layer: 'hot', tags: ['发布', '发布', ' 发布 '] });
    assert.equal(dup.hot.items.length, 2, '重复的分区键归一成一条（合取里不影响结果）');

    // 空白的分区键**不是**「没给分区」（复核 P3-2）：静默当没给，回执里连分区行都不印，
    // 读的人会把「没过滤」当成「这一类只有这些」。要「不过滤」就传空数组。
    await assert.rejects(() => memory.read({ layer: 'hot', tags: ['', '  '] }), /分区键不能是空串/,
        '空白的分区键要拒绝，不静默变成「不过滤」');
    assert.equal((await memory.read({ layer: 'hot', tags: [] })).hot.items.length, 3, '空数组 = 不过滤');
});

await check('C3 台账的分区键是 theme；拿实体分区时如实报「分区外的条目」', async () => {
    const { memory } = await freshRoot();
    await memory.log([
        { path: 'a.md', format: 'md', theme: '发布', purpose: '发布流程说明' },
        { path: 'b.md', format: 'md', theme: '财务', purpose: '财务口径' },
        { path: 'c.md', format: 'md', purpose: '没有主题的台账行' },
    ]);

    const themed = await memory.read({ layer: 'ledger', tags: ['发布'] });
    assert.equal(themed.ledger.items.length, 1, '台账按 theme 参与分区');
    assert.equal(themed.ledger.partitionFiltered, 2, 'theme 不是「发布」的两条被隔掉');

    // 同一维度里给两个词是合取 —— 台账一行只有一个 theme，所以必然 0 条。这条钉住
    // 「合取不因为层不同就变成析取」（两层各写一套判据是这类改动最常见的退化）。
    const two = await memory.read({ layer: 'ledger', tags: ['发布', '财务'] });
    assert.equal(two.ledger.items.length, 0, '给两个标签仍然要同时满足（合取）');

    const byEntity = await memory.read({ layer: 'ledger', entities: ['发布流程'] });
    assert.equal(byEntity.ledger.items.length, 0, '台账没有实体，实体分区挡掉全部');
    assert.ok(renderRead(byEntity).includes('分区外的条目'), '要说清是分区挡的，不是「台账里没有」');
});

await check('C4 归档条目带标签下沉后仍按分区可读，隔掉的条数能对账', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 320, archiveKeep: 10 });
    await memory.mutate({ action: 'add', target: 'project', content: `要保住的那条：${'重'.repeat(30)}`, importance: 'critical', tags: ['保留'] });
    for (let i = 0; i < 6; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `会下沉 ${i}：${'旧'.repeat(30)}`, importance: 'low', tags: [i < 3 ? '发布' : '财务'] });
    }

    const all = await memory.read({ layer: 'archive' });
    assert.ok(all.archive.items.length >= 3, `下沉后归档里应有多条（实际 ${all.archive.items.length}）`);

    const published = await memory.read({ layer: 'archive', tags: ['发布'] });
    assert.ok(published.archive.items.length >= 1, '归档里的条目按它自己的 tags 参与分区');
    assert.ok(published.archive.items.every((item) => (item.tags ?? []).includes('发布')), '分区外的归档条目不许出现');
    assert.ok(published.archive.items.every((item) => !String(item.text).includes('财务')), '隔掉的那一半不能漏进来');
    assert.equal(published.archive.partitionFiltered + published.archive.matched, all.archive.matched,
        '「隔掉的 + 命中的」要等于分区前的条数（账要对得上）');
});

await check('C6 归档层的「为什么没给」也要对：分区清空归档时说分区，不说额度', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 320, archiveKeep: 10 });
    for (let i = 0; i < 6; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `会下沉 ${i}：${'旧'.repeat(30)}`, importance: 'low', tags: ['发布'] });
    }
    const all = await memory.read({ layer: 'archive' });
    assert.ok(all.archive.items.length >= 3, `夹具要先真的沉下去（实际 ${all.archive.items.length}）`);

    // 归档那一层的「分区前条数」字段叫 items_total（热记忆与台账叫 total）：只认 total 的
    // 判据会 `>= undefined` 恒为假，于是归档永远掉到最后那句假的「额度被占满」（复核 P1-1）。
    const miss = await memory.read({ layer: 'archive', tags: ['根本没有这个分区'] });
    assert.equal(miss.archive.items.length, 0, '分区外的归档条目不该给出来');
    const text = renderRead(miss);
    assert.ok(text.includes('分区外的条目'), `归档层要说是分区挡的：${text}`);
    assert.ok(!text.includes('被更长的条目占满'), `不许把分区说成额度：${text}`);

    // 带 query 时同理：真因是分区，不是关键词。
    const withQuery = await memory.read({ layer: 'archive', tags: ['根本没有这个分区'], query: '会下沉' });
    assert.ok(renderRead(withQuery).includes('分区外的条目'), '带 query 时也要说是分区挡的');
});

await check('C7 热记忆的排序不随 query 变：low 的笔记命中更多词也压不过 critical', async () => {
    const { memory } = await freshRoot();
    // 关键的一条：critical 是**约束**，low 只是一条笔记。命中数更多也不该把约束挤下去 ——
    // 第五十四轮初版让命中强度当主键，`limit:1` 时这条约束会直接消失（复核 P1-4）。
    await memory.mutate({ action: 'add', target: 'project', content: '永远不要把草稿发给客户。', importance: 'critical' });
    await memory.mutate({ action: 'add', target: 'project', content: '草稿发给客户之前，客户名与日期都要再核一遍，草稿的版本号也要写。', importance: 'low' });

    const ranked = await memory.read({ layer: 'hot', query: '草稿 客户' });
    assert.equal(ranked.hot.items.length, 2, '两条都命中，都要给出来');
    assert.equal(ranked.hot.items[0].importance, 'critical', '重要度仍是主键（命中更多词不改变它）');

    const top = await memory.read({ layer: 'hot', query: '草稿 客户', limit: 1 });
    assert.equal(top.hot.items.length, 1);
    assert.ok(top.hot.items[0].text.includes('永远不要'), 'limit 截断时留下的必须是那条约束');

    // 同族两版：旧版命中词更多也不该冒到当前版前面（第四十二轮量到增益才落地的契约）。
    const older = await memory.mutate({ action: 'add', target: 'project', content: '发布窗口：旧版写着周三，发布窗口的细节在旧版里更全。', importance: 'normal' });
    await new Promise((resolve) => { setTimeout(resolve, 5); });
    const newer = await memory.mutate({ action: 'add', target: 'project', content: '发布窗口：当前版写着周五。', importance: 'normal' });
    const family = await memory.read({ layer: 'hot', query: '发布窗口' });
    const order = family.hot.items.map((item) => item.id);
    assert.ok(order.indexOf(newer.entry.id) < order.indexOf(older.entry.id),
        `同重要度里当前版要在旧版前面（实测次序 ${order.join(' > ')}）`);
});

await check('C5 分区与 query 同时给：先分区后过滤，命中相同时重要度仍是主键', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '甲：季度回顾用模板 A。', tags: ['发布'], importance: 'low' });
    await memory.mutate({ action: 'add', target: 'project', content: '乙：月度经营简报用模板 B。', tags: ['发布'], importance: 'critical' });
    await memory.mutate({ action: 'add', target: 'project', content: '丙：季度回顾用模板 A，但不在这个分区。', tags: ['财务'], importance: 'critical' });

    const hit = await memory.read({ layer: 'hot', tags: ['发布'], query: '季度' });
    assert.equal(hit.hot.items.length, 1, '分区先过滤：财务那条即使命中关键词也不出现');
    assert.ok(hit.hot.items[0].text.includes('甲'), '留下的是分区内的那条');

    const both = await memory.read({ layer: 'hot', tags: ['发布'], query: '模板' });
    assert.equal(both.hot.items.length, 2, '两条都在分区里');
    assert.equal(both.hot.items[0].importance, 'critical', '命中强度相同时重要度仍是主键（不是时间倒序）');

    // 带 query 就是检索：一个词都没命中的条目与台账 / 归档同口径地不出现。
    const none = await memory.read({ layer: 'hot', tags: ['发布'], query: '量子纠缠' });
    assert.equal(none.hot.items.length, 0, '没命中就不给');
    assert.ok(renderRead(none).includes('没有命中关键词'), '要说清是关键词没命中（换词），不是分区挡的');
});

// ── D 组（第五十七轮 P2-8）：统一入口路由的「记忆够答吗」门控 ─────────────────

/*
 * 判据是**纯函数**（`gateVerdictOf`），所以这里可以逐档钉住它；`read` 侧只负责
 * 把三样现算出来的事实（命中条数 / 查询词覆盖 / 最高重要度）交给它。
 * 之所以两层分开：门控一旦混进「读」的实现里，就没有任何东西能挡住下一轮
 * 悄悄把某一档改成别的意思 —— 而这一档决定的是「要不要再花一次 kb 配额」。
 */

await check('D1 门控四档：纯判据逐档可核对（顺序即优先级）', () => {
    assert.equal(gateVerdictOf({ terms: ['甲', '乙'], covered: 0, hits: 0, topImportance: '' }), 'empty');
    // 命中 0 条时连「覆盖」都不看：没有命中就一定是 empty。
    assert.equal(gateVerdictOf({ terms: ['甲', '乙'], covered: 2, hits: 0, topImportance: 'critical' }), 'empty');
    assert.equal(gateVerdictOf({ terms: ['甲', '乙'], covered: 1, hits: 3, topImportance: 'critical' }), 'partial');
    assert.equal(gateVerdictOf({ terms: ['甲', '乙'], covered: 2, hits: 3, topImportance: 'normal' }), 'weak');
    assert.equal(gateVerdictOf({ terms: ['甲', '乙'], covered: 2, hits: 3, topImportance: 'low' }), 'weak');
    assert.equal(gateVerdictOf({ terms: ['甲', '乙'], covered: 2, hits: 3, topImportance: 'critical' }), 'enough');
    // 四档的中文说法与判据同源（回执与测试不各写各的）。
    for (const verdict of ['empty', 'partial', 'weak', 'enough']) {
        assert.equal(typeof GATE_VERDICT_TEXT[verdict], 'string', `缺 ${verdict} 的中文说法`);
    }
    return '4 档 + 1 条边界';
});

await check('D2 read 现算门控：命中数、查询词覆盖与最高重要度都来自真的命中', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '发布流程：打包前必须跑回归。', importance: 'critical' });
    await memory.mutate({ action: 'add', target: 'project', content: '主题色用 plain。', importance: 'normal' });

    // 一条约束级 + 词全覆盖 → 够答。
    const enough = await memory.read({ query: '发布流程' });
    assert.equal(enough.gate.verdict, 'enough', JSON.stringify(enough.gate));
    assert.equal(enough.gate.hits, 1);
    assert.equal(enough.gate.topImportance, 'critical');
    assert.equal(enough.gate.covered.length, enough.gate.terms.length);
    assert.deepEqual(enough.gate.missing, []);

    // 记忆里没有这件事 → empty（不是「额度吃掉」也不是「分区挡的」）。
    const empty = await memory.read({ query: '量子纠缠' });
    assert.equal(empty.gate.verdict, 'empty');
    assert.equal(empty.gate.hits, 0);
    assert.deepEqual(empty.gate.covered, []);
    assert.equal(empty.gate.missing.length, empty.gate.terms.length);

    // 半边对上：命中里有一半词对不上 → partial（并且点出缺的是哪几个词）。
    const partial = await memory.read({ query: '发布流程 量子计算' });
    assert.equal(partial.gate.verdict, 'partial', JSON.stringify(partial.gate));
    assert.ok(partial.gate.missing.includes('量子'), JSON.stringify(partial.gate.missing));
    assert.ok(partial.gate.covered.includes('发布'));

    // 词全覆盖但没有约束级 → weak（「有相关内容」不等于「够答」）。
    const weak = await memory.read({ query: '主题色' });
    assert.equal(weak.gate.verdict, 'weak', JSON.stringify(weak.gate));
    assert.equal(weak.gate.topImportance, 'normal');

    // 不带 query 的浏览没有门控这回事（省掉无意义的字段）。
    const browse = await memory.read({ layer: 'hot' });
    assert.equal(browse.gate, undefined, '不带 query 不该给门控');
});

await check('D3 门控只做过滤不做重排：带 query 的次序 = 不带 query 的次序去掉没命中的那些', async () => {
    const { memory } = await freshRoot();
    // 夹具要能让**两种排法的结果不同**，断言才有鉴别力（独立复核 P1-1 打穿了初版：
    // 拿同一次调用的两份拷贝互比，任何实现下都成立）：甲是 low、命中两个词、且**更旧**；
    // 乙是 critical、只命中一个词、且**更新**。按「命中强度」排会得到 [甲, 乙]，
    // 按「重要度是主键」排必须得到 [乙, 甲]；recency 也只会把甲往前推，
    // 所以这一条真的能红。
    await memory.mutate({ action: 'add', target: 'project', content: '甲：交付档命名口径要带当天日期。', importance: 'low' });
    await new Promise((resolve) => { setTimeout(resolve, 5); });
    await memory.mutate({ action: 'add', target: 'project', content: '乙：交付档不要沿用上一轮。', importance: 'critical' });

    const ranked = await memory.read({ layer: 'hot', query: '交付档 命名' });
    assert.equal(ranked.hot.items.length, 2, '两条都命中');
    assert.equal(ranked.hot.items[0].importance, 'critical', '重要度仍是主键（命中更多词、更旧都不改变它）');

    // 与「不带 query 的浏览」对账：把没命中的那些去掉，次序必须逐字一致 ——
    // 门控的作用是**缩小集合**，不是重排（第五十四轮 P1-4 的契约）。
    const browse = await memory.read({ layer: 'hot' });
    const browsedIds = browse.hot.items.map((item) => item.id);
    const rankedIds = ranked.hot.items.map((item) => item.id);
    assert.deepEqual(rankedIds, browsedIds.filter((id) => rankedIds.includes(id)),
        `带 query 的次序必须是浏览次序的子序列（浏览 ${browsedIds.join(' >')} / 命中 ${rankedIds.join(' >')}）`);
    assert.equal(browse.gate, undefined, '不带 query 的浏览没有门控这回事');

    // 换了 query（门控原料跟着变）也不许改变同一批条目的相对次序。
    const other = await memory.read({ layer: 'hot', query: '上一轮' });
    assert.deepEqual(other.hot.items.map((item) => item.id), [rankedIds[0]], '只有乙命中「上一轮」');
});

await check('D4 归档里的 critical 也算得上门控的「约束级」（否则回执会否认自己正文）', async () => {
    // 独立复核 P2-1：门控最初只从热记忆层取最高重要度，于是 `layer:'archive'` 永远到不了
    // enough，判定行写「命中的没有约束级条目」——而紧跟的那行正文就印着 `热记忆（…/critical）`。
    const { memory } = await freshRoot({ projectLimitBytes: 320, archiveKeep: 10 });
    for (let i = 0; i < 6; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `硬约束 ${i}：永远不要把草稿发给客户，编号 ${i}。`, importance: 'critical' });
    }
    const archived = await memory.read({ layer: 'archive' });
    assert.ok(archived.archive.items.length >= 1, '夹具要先真的沉下去');
    assert.ok(archived.archive.items.some((item) => item.importance === 'critical'), '归档里那条确实是 critical');

    const hit = await memory.read({ layer: 'archive', query: '草稿 客户' });
    assert.equal(hit.archive.items.length >= 1, true, '要能命中');
    assert.equal(hit.gate.topImportance, 'critical', `归档 hot-kind 的重要度要算进门控：${JSON.stringify(hit.gate)}`);
    assert.equal(hit.gate.verdict, 'enough', '词全覆盖 + 归档里那条 critical ⇒ 够答');

    // 台账条目本来没有重要度字段：拿它当唯一命中时，不该凭空判出约束级。
    await memory.log([{ path: 'docs/x.md', format: 'markdown', theme: '甲', purpose: '台账条目：草稿与客户' }]);
    const ledgerOnly = await memory.read({ layer: 'ledger', query: '草稿 客户' });
    assert.equal(ledgerOnly.gate.topImportance, '', '台账没有重要度字段，不该凭空算出一个');
    assert.equal(ledgerOnly.gate.verdict, 'weak', '词对上但没有约束级 ⇒ weak（下探 kb 是合理的）');
});

// ── 汇总 ───────────────────────────────────────────────────────────────────

await rm(TMP, { recursive: true, force: true });

const failed = results.filter((r) => !r.ok);
for (const r of results) {
    if (r.ok) console.log(`PASS  ${r.name}`);
    else console.log(`FAIL  ${r.name}\n      ${r.error?.message ?? r.error}`);
}
console.log(`\nmemory-recall: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
