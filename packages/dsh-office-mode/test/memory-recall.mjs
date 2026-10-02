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

import { createMemory, READ_LIMITS, renderRead } from '../src/memory.js';

const TMP = fileURLToPath(new URL('../../../.office/tmp/memory-recall/', import.meta.url));

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
 * 各 0 条**（`.office/tmp/r24-impl/debug-starve.mjs` 在两版代码上分别量过）。
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
    ['.office/tmp/python-e2e/Python计算与绘图.pptx', '端到端样例：numpy 计算 + matplotlib 中文图嵌进 PPT'],
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
    ['python 绘图', ['docs/第十一轮-办公插件恢复Python计算与绘图.md', '.office/tmp/python-e2e/Python计算与绘图.pptx'], '中英混排'],
    ['python绘图', ['docs/第十一轮-办公插件恢复Python计算与绘图.md', '.office/tmp/python-e2e/Python计算与绘图.pptx'], '中英黏连（必须分段处理）'],
    ['计算与绘图', ['docs/第十一轮-办公插件恢复Python计算与绘图.md', '.office/tmp/python-e2e/Python计算与绘图.pptx'], '整串命中'],
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

// ── 汇总 ───────────────────────────────────────────────────────────────────

await rm(TMP, { recursive: true, force: true });

const failed = results.filter((r) => !r.ok);
for (const r of results) {
    if (r.ok) console.log(`PASS  ${r.name}`);
    else console.log(`FAIL  ${r.name}\n      ${r.error?.message ?? r.error}`);
}
console.log(`\nmemory-recall: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
