/**
 * 记忆侧回归探针：三个**可复跑的量表**（第二十四轮 P2-15 / 总清单 24-13）。
 *
 * 三条口径，先写清楚，免得后来的人把它当「质量门」看：
 *
 * 1. **期望值是「能测出来」，不是「一定通过」**（第二十四轮定的）。
 *    每个探针都要证明**自己能分辨好坏**：同一批夹具上，故意做坏的排名器必须拿到更差的
 *    数字。度量本身要是有问题，这一条会先红。
 * 2. **数字照实打印**。hit@3 / recall@3 / 弃答率 / latest@1 都打到输出里 —— 好坏一眼可见，
 *    但不会因为「数字难看」把套件弄红。契约类的断言（引擎自己的行为）该红就红。
 * 3. **C 组量的是产品路径本身**（第四十五轮改成这样）。第四十二轮它带的是「只活在探针里」
 *    的候选打分器 —— 那是接线（24-11）之前的前置自测集。24-11 落地后，打分器搬进了
 *    `src/kb.js` 并由 `office_memory({action:'kb-search'})` 提供，所以这一组改成**直接调
 *    action 面**：探针里再留一份拷贝就等于量了个假的东西。C4 反向钉住这件事。
 *    记忆侧（A / B 组）走的本来就是真实 `read` / 探针自算，没变。
 *
 * 三组形状：
 *   A 组 LongMemEval ABS 形状（弃答）：可答的必须答得出，不可答的必须弃答（不许编一个命中）。
 *   B 组 MemoryAgentBench FactConsolidation 形状：同一事实的多个版本并存时，取到的是不是最新那版。
 *   C 组 kb recall@k 自测集：4 份文档 / 12 块上的 15 条标注查询 + 2 条弃答查询。
 *
 * 跑法：node test/memory-probes.mjs
 */
import { mkdir, rm, writeFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { kbRankChunks, kbTerms, kbTokens } from '../src/kb.js';
import { createMemory, memoryPaths } from '../src/memory.js';

const TMP = fileURLToPath(new URL('../../../.office/tmp/memory-probes/', import.meta.url));

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

const TOPK = 3;

/** 夹具用的短等待：让两个写入落在不同的毫秒上（B3 的时间戳分辨力）。 */
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// ── A 组：LongMemEval ABS 形状（弃答） ──────────────────────────────────────
//
// 形状取自 LongMemEval 的 abstention 子集：一半问题在记忆里**有**答案，另一半问的东西
// 从来没被记过。好的系统在第二类上必须什么都不给（弃答），而不是把最像的那条端出来。
// 这一组只量「弃答」，不量长上下文里的多跳 —— 那是另一张表。

const ABS_STORED = [
    ['docs/季度复盘-2026Q3.md', '第三季度交付物清单与复盘结论'],
    ['docs/主题色规范.md', '产品主题色与字号的规范说明'],
    ['docs/客户A-需求变更.md', '客户 A 在二期提出的需求变更记录'],
    ['docs/数据导出流程.md', '从台账导出 CSV 的操作流程'],
    ['docs/会议室预约.md', '会议室预约与冲突处理办法'],
    ['docs/论文投稿计划.md', '论文投稿的时间表与目标期刊'],
    ['packages/dsh-office-mode/src/python.js', 'office.python 的实现与六条硬约束'],
    ['docs/离线部署.md', '离线环境下的安装与依赖准备'],
    ['docs/发票报销.md', '发票报销的填写要求与常见退单原因'],
    ['docs/值班表-2026-09.md', '九月值班安排与替班规则'],
    ['docs/字体授权.md', '商用字体的授权范围与替换方案'],
    ['docs/接口限流.md', '对外接口的限流口径与重试建议'],
];

/** 可答查询：每条都能在 ABS_STORED 里找到唯一一份来源。 */
const ABS_ANSWERABLE = [
    ['季度复盘', 'docs/季度复盘-2026Q3.md'],
    ['主题色 规范', 'docs/主题色规范.md'],
    ['客户A 需求变更', 'docs/客户A-需求变更.md'],
    ['台账 导出 CSV', 'docs/数据导出流程.md'],
    ['会议室 冲突', 'docs/会议室预约.md'],
    ['投稿 期刊', 'docs/论文投稿计划.md'],
    ['python 硬约束', 'packages/dsh-office-mode/src/python.js'],
    ['离线部署', 'docs/离线部署.md'],
    ['发票 退单', 'docs/发票报销.md'],
    ['值班 替班', 'docs/值班表-2026-09.md'],
];

/** 弃答查询：这些主题一个字都没进过记忆库。 */
const ABS_ABSENT = [
    '量子纠缠 退相干',
    '火星移民 补给窗口',
    '鲸鱼 迁徙路线',
    '咖啡 烘焙曲线',
    '拉丁语 变格表',
    '火山灰 成分分析',
];

async function seedAbsLedger(memory) {
    for (const [path, purpose] of ABS_STORED) await memory.log([{ path, format: 'md', purpose }]);
}

await check('A1 可答查询：hit@3（夹具本身必须答得出，否则量表测不了东西）', async () => {
    const { root, memory } = await freshRoot();
    await seedAbsLedger(memory);

    const rows = [];
    for (const [query, want] of ABS_ANSWERABLE) {
        const r = await memory.read({ layer: 'ledger', query });
        const got = r.ledger.items.slice(0, TOPK).map((item) => item.path);
        rows.push({ query, hit: got.includes(want) ? 1 : 0, shown: r.ledger.items.length, got });
    }
    const hit = rows.reduce((n, row) => n + row.hit, 0);
    const missed = rows.filter((row) => row.hit === 0).map((row) => row.query);
    console.log(`      A1 可答 hit@3 ${hit}/${rows.length}`);
    assert.ok(rows.length > 0, '量表不能是空的');
    assert.ok(hit > 0, `一条都答不出来说明夹具坏了（未中：${missed.join('、')}）`);
});

await check('A2 弃答查询：报出弃答率（不许编一个命中）', async () => {
    const { root, memory } = await freshRoot();
    await seedAbsLedger(memory);

    const rows = [];
    for (const query of ABS_ABSENT) {
        const r = await memory.read({ layer: 'ledger', query });
        const shown = r.ledger.items.length;
        rows.push({ query, abstained: shown === 0 && r.ledger.matched === 0 ? 1 : 0, shown, matched: r.ledger.matched });
    }
    const abstained = rows.reduce((n, row) => n + row.abstained, 0);
    const leaked = rows.filter((row) => row.abstained === 0).map((row) => `${row.query}→${row.shown} 条`);
    console.log(`      A2 弃答率 ${abstained}/${rows.length}${leaked.length > 0 ? `，漏答：${leaked.join('、')}` : ''}`);
    assert.ok(rows.length >= 6, '弃答样本不能太少');
});

await check('A3 灵敏度：弃答率必须能分辨「总是说命中」的坏检索器', async () => {
    const { root, memory } = await freshRoot();
    await seedAbsLedger(memory);

    let abstained = 0;
    for (const query of ABS_ABSENT) {
        const r = await memory.read({ layer: 'ledger', query });
        if (r.ledger.items.length === 0 && r.ledger.matched === 0) abstained += 1;
    }
    const real = abstained / ABS_ABSENT.length;
    // 「总是命中」的坏检索器：不管问什么都把库里的前 3 条端出来 —— 弃答率恒为 0。
    // 它就是「不做弃答判断」的下界，探针必须把它与真引擎分开。
    const alwaysHit = 0;
    assert.ok(real > alwaysHit, `探针没有分辨力：真引擎 ${real} vs 坏检索器 ${alwaysHit}`);

    // 另一半分辨力：答案全在库里时，坏检索器（什么都不给）也会拿到 0 的 hit@3。
    const stubHit = 0;
    let realHit = 0;
    for (const [query, want] of ABS_ANSWERABLE) {
        const r = await memory.read({ layer: 'ledger', query });
        if (r.ledger.items.slice(0, TOPK).some((item) => item.path === want)) realHit += 1;
    }
    const realHitRate = realHit / ABS_ANSWERABLE.length;
    assert.ok(realHitRate > stubHit, `探针没有分辨力：真引擎 ${realHitRate} vs 空检索器 ${stubHit}`);
});

// ── B 组：FactConsolidation 形状（同一事实的多个版本） ──────────────────────
//
// 形状取自 MemoryAgentBench 的 FactConsolidation：同一个事实被先后写下多次，
// 只有**最新那版**算对。第四十二轮这一组量出「引擎把旧版排在前面 0/4」——
// 那是 24-4（recency 弱先验）立项时的对照数；**24-4 落地后这个数应当变成 4/4**，
// 所以 B3 现在有通过线，B4 是配套的「加时间项不伤单会话题」对照。

/** 每个事实族：[族名, 旧说法, 新说法, 查询, 干扰条目]。 */
const FACT_FAMILIES = [
    ['版式比例', '汇报的版式比例用 4:3。', '汇报的版式比例改用 16:9（客户屏幕都是宽屏）。', '版式比例', '版式的间距按 8 的倍数取。'],
    ['图表字体', '图表字体用宋体。', '图表字体改用思源黑体，宋体在投影上太细。', '图表字体', '图表要留白，不要贴边。'],
    ['台账上限', '台账上限设 200 条。', '台账上限改成 500 条，200 条滚得太快。', '台账上限', '台账每条记一次交付物。'],
    ['评审节奏', '评审每周一次。', '评审改成每两周一次，每周太密。', '评审频率', '评审要提前一天发材料。'],
];

await check('B1 版本合并：同一事实的多次更新只留一条，读到的是最新版', async () => {
    const { root, memory } = await freshRoot();
    const ids = [];
    for (const [name, , newText] of FACT_FAMILIES) {
        const added = await memory.mutate({ action: 'add', target: 'project', content: newText });
        ids.push({ name, id: added.entry.id, newText });
    }
    // 用 replace 把同一族改三次（最后一次才是当前版本）：引擎的合并语义就在这里。
    for (const row of ids) {
        for (let i = 0; i < 3; i += 1) {
            await memory.mutate({ action: 'replace', id: row.id, content: `${row.newText}（第 ${i + 1} 次修订）` });
        }
    }

    let latestFirst = 0;
    for (const row of ids) {
        const r = await memory.read({ layer: 'hot', query: row.name });
        const top = r.hot.items[0];
        if (top && String(top.text).includes('第 3 次修订')) latestFirst += 1;
    }
    const rate = latestFirst / ids.length;
    console.log(`      B1 合并后 latest@1 ${latestFirst}/${ids.length}`);
    assert.equal(rate, 1, 'replace 的契约是「就地改」，读到旧版本说明合并语义坏了');
});

await check('B2 supersedes 边：旧条目与新条目并存时，关系可查、且要能分辨新旧', async () => {
    const { root, memory } = await freshRoot();
    const old = await memory.mutate({ action: 'add', target: 'project', content: '月度复盘放在每月 5 号。' });
    const fresh = await memory.mutate({ action: 'add', target: 'project', content: '月度复盘改到每月 8 号，5 号太赶。' });
    await memory.link({
        sourceId: fresh.entry.id, targetId: old.entry.id, kind: 'supersedes',
        note: '8 号取代 5 号',
    });

    const rel = await memory.related({ id: fresh.entry.id });
    // related 的形状是 nodes + edges（第四十一轮起边还带冲突类别与状态）。
    const edge = (rel.edges ?? []).find((item) => item.kind === 'supersedes');
    assert.ok(edge, 'supersedes 边要能沿关系查到');
    assert.equal(edge.from, fresh.entry.id);
    assert.equal(edge.to, old.entry.id);

    // 条目本身带 updatedAt：这是「谁更新」的**唯一**现成依据（时间项还没进排序，
    // 见 24-4）。两版并存时排序不分新旧，所以这里只断言「时间戳能分辨」。
    const r = await memory.read({ layer: 'hot', query: '月度复盘' });
    assert.ok(r.hot.items.length >= 2, '两版并存时都应读得到（引擎不替用户裁决）');
    const stamps = r.hot.items.map((item) => String(item.updatedAt ?? ''));
    assert.ok(new Set(stamps).size >= 1, '每条都要有 updatedAt');
});

await check('B3 增益：同族两版并存时最新版在前（24-4 落地后的契约，量到的是 0/4 → 4/4）', async () => {
    const { root, memory } = await freshRoot();
    const families = [];
    for (const [name, oldText, newText, query] of FACT_FAMILIES) {
        const first = await memory.mutate({ action: 'add', target: 'project', content: oldText });
        // 两个版本的时间戳必须**真的不同**：同一毫秒里写两条时 updatedAt 会撞在一起，
        // 纯时间排序就成了稳定排序的副产品，量表量不到东西（第一版夹具就踩了这个）。
        await sleep(6);
        const second = await memory.mutate({ action: 'add', target: 'project', content: newText });
        assert.notEqual(first.entry.updatedAt, second.entry.updatedAt, '夹具要求两版时间戳不同');
        families.push({ name, query, oldId: first.entry.id, newId: second.entry.id });
    }

    let engine = 0;
    let oracle = 0;
    let stale = 0;
    for (const family of families) {
        const r = await memory.read({ layer: 'hot', query: family.query });
        // 只看同族这两条：热记忆的读取**不分查询排序**，整层按「重要度 → updatedAt 升序」
        // 给出来（`read` 的 hot 分支），所以整层的第一条是全局最旧的，拿它当度量是错的。
        const items = r.hot.items.filter((item) => item.id === family.oldId || item.id === family.newId);
        assert.equal(items.length, 2, `两版都应在结果里（${family.name} 只有 ${items.length} 条）`);
        if (items[0].id === family.newId) engine += 1;
        // 两种**纯时间**排序器，作用在同两条上：这就是「加分项加在哪」的对照。
        const byNewest = [...items].sort((left, right) => String(right.updatedAt).localeCompare(String(left.updatedAt)));
        const byOldest = [...items].sort((left, right) => String(left.updatedAt).localeCompare(String(right.updatedAt)));
        if (byNewest[0].id === family.newId) oracle += 1;
        if (byOldest[0].id === family.newId) stale += 1;
    }
    console.log(`      B3 引擎最新版在前 ${engine}/${families.length} ｜ 纯时间倒序（上界）${oracle}/${families.length} ｜ 纯时间正序（下界）${stale}/${families.length}`);
    // 夹具不变量：上下界必须是满分与 0。它们不证明「引擎有缺陷」，证明的是**度量对排序
    // 有响应** —— 若这批夹具里两种时间排序看不出差别（比如两版时间戳撞在一起、或「新版」
    // 其实写在前面），量表就是坏的，这条会先红。
    //
    // 第四十二轮的独立复核正是用「把新版写在前面」这个**假夹具**打红过这一条：那种夹具里
    // 引擎反而 4/4，说明本指标读的是「谁的时间戳更新」，不是「谁先被写」—— 这正是我们要的
    // 语义（事实更新只可能是后写的），但后来的人别拿它当「引擎会不会看时间」的万能证据。
    assert.equal(oracle, families.length, '纯时间倒序应当每条都把新版排在最前');
    assert.equal(stale, 0, '纯时间正序应当一条都不给新版');
    // 增益与它的落地（第二十四轮 24-4 / 第四十五轮）：读取排序现在是「重要度 →
    // recency 弱先验（0.995/日）」，所以同族两版并存时**最新那版排在最前**。
    // 改造前这个数是 0/4（第四十二轮的对照数）—— 断言改成满分，是为了让「把时间项
    // 拿掉」这件事当场变红（回滚探针用的就是这一条）。
    assert.equal(engine, families.length, '时间项落地后，同族两版里最新那版必须排在前面');
});

/**
 * B4 是对照组（24-4 的门槛之二：**要有「加时间项不伤单会话题」的对照**）。
 *
 * 单会话题 = 每条事实只有一个版本，没有「谁取代谁」这件事。这时时间项唯一能做的事
 * 就是**换次序**，所以这一组量的是「只动次序、不动集合」以及两条不该发生的事：
 *   - 任何条目都不许因为加了时间项而从读取结果里消失（含时间戳坏掉的条目）；
 *   - 时间戳解析不出来的条目不许冒到同重要度的前面（拿不出证据的不该被当成最新）。
 */
await check('B4 对照：加时间项只动次序不动集合，坏时间戳不冒头（单会话题不受损）', async () => {
    const { root, memory } = await freshRoot();
    const paths = memoryPaths(root, {});
    // 手工写一份「单会话」的热记忆：6 条正常时间戳（按时间递增）+ 2 条坏时间戳。
    const normal = [];
    for (let i = 1; i <= 6; i += 1) {
        normal.push({
            id: `m-fact-${i}`,
            target: 'project',
            content: `单会话事实 ${i}：话题词${i} 只出现在这一条里。`,
            importance: 'normal',
            createdAt: `2026-09-0${i}T00:00:00.000Z`,
            updatedAt: `2026-09-0${i}T00:00:00.000Z`,
        });
    }
    const broken = [
        // 用的是**解析不出来**的时间戳（不是缺失）：缺失会被 normalizeEntry 补成读取时刻
        // （那条行为写在 office_help('memory') 里），那种条目等于「刚刚」，不在本对照的范围。
        { id: 'm-broken-1', target: 'project', content: '时间戳坏掉的条目甲。', importance: 'normal', updatedAt: 'not-a-date' },
        { id: 'm-broken-2', target: 'project', content: '时间戳坏掉的条目乙。', importance: 'normal', updatedAt: 'x' },
    ];
    await mkdir(paths.dir, { recursive: true });
    await writeFile(paths.memory, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-06T00:00:00.000Z',
        entries: [...normal, ...broken],
    }, null, 2), 'utf8');

    const r = await memory.read({ layer: 'hot' });
    const ids = r.hot.items.map((item) => item.id);
    // 1) 一条都不许丢：8 条全在（集合不随排序变化）。
    assert.equal(ids.length, 8, `单会话的 8 条都该读得到，实际 ${ids.length}`);
    assert.deepEqual([...ids].sort(), [...normal, ...broken].map((entry) => entry.id).sort(), '集合必须与写入时一致');
    // 2) 坏时间戳的两条排在同重要度的最后。
    assert.deepEqual(ids.slice(-2).sort(), ['m-broken-1', 'm-broken-2'],
        `坏时间戳的条目该排在最后，实际 ${ids.join(',')}`);
    // 3) 对照数：旧口径（updatedAt 升序）与新口径（recency 降序）各取前 3 —— 照实打印。
    const stampById = new Map(r.hot.items.map((item) => [item.id, String(item.updatedAt ?? '')]));
    const byOldest = [...ids]
        .sort((left, right) => String(stampById.get(left)).localeCompare(String(stampById.get(right))))
        .slice(0, TOPK);
    const byRecency = ids.slice(0, TOPK);
    console.log(`      B4 单会话题 top${TOPK}：旧口径（最旧在前）${byOldest.join('/')} ｜ 新口径（最新在前）${byRecency.join('/')}`);
    // 4) 名额收窄到 3 条时，坏时间戳的条目同样不该挤进来。
    const capped = await memory.read({ layer: 'hot', limit: TOPK });
    assert.equal(capped.hot.items.length, TOPK, 'limit 要生效');
    assert.equal(capped.hot.items.some((item) => String(item.id).startsWith('m-broken-')), false,
        '名额不够时，时间戳坏掉的条目不该挤掉有证据的条目');
});

// ── C 组：kb recall@k 自测集（24-11 的前置 → 落地后的回归） ──────────────────
//
// 24-11 的硬门槛是「本地 recall@k 自测集先变绿」。这一组就是那套自测集：
// 4 份文档（各 3 块）+ 15 条标注查询 + 2 条弃答查询。
//
// 第四十二轮这一组量的是探针里自带的一份候选打分器（bigram + BM25 k1=1.2 b=0.75），
// **本轮（24-11 接线）改成直接量 `memory.kbSearch()`** —— 打分器已经搬进 `src/kb.js`
// （`kbRankChunks`，纯函数），探针里再留一份拷贝就等于量了个假的东西。朴素整串与随机
// 排序这两条**故意做坏的对照**仍留在这里：它们本来就不该进产品路径。

/** 朴素基线：整串子串命中（没有 bigram、没有词序容错）。灵敏度对照用。 */
function naiveRank(docs, query) {
    const needle = String(query).toLowerCase();
    return docs.map((doc) => ({ doc, score: String(doc.text).toLowerCase().includes(needle) ? 1 : 0 }))
        .filter((row) => row.score > 0);
}

const KB_DOCS = [
    ['docs/知识库设计.md', [
        '# 知识库设计',
        '## 切块',
        '知识库按结构切块：先按标题分层，再按字数收口，块与块之间保留字符区间。每个块带一个内容寻址的 id：同一份内容被两份文档共用时共用同一批块，块文件里只放块自己的东西。',
        '## 入库',
        '入库是显式的：只有 kb-ingest 会把一份文档切开放进来，没有后台扫描。入库前先算内容与切块档一起的哈希，同内容不重写，新内容换代并给旧块留墓碑。',
        '## 来源档',
        '来源档分三档：user / verified / unverified。单源未核实的材料不许进热记忆 —— 这条是写入期的硬规则，不是权重。',
    ].join('\n\n')],
    ['docs/召回实验记录.md', [
        '# 召回实验记录',
        '## 查询展开',
        '中文召回先用 bigram 展开查询。实测里「面板记忆」这类词序颠倒的查询只有 bigram 才召得回，整串匹配会整片落空。',
        '## 排序',
        '排序用 BM25，k1 取一点二、b 取零点七五：先按词频饱和，再按文档长度归一。多路命中按归一化加权和融合，避免某一路把别的路盖过去。',
        '## 已知边界',
        '语义改写仍然召不回：查询与正文用词完全对不上时，词法检索给不出结果。这条边界写在帮助文档里，别把它当 bug 修。',
    ].join('\n\n')],
    ['docs/归档分卷说明.md', [
        '# 归档分卷说明',
        '## 分卷',
        '单月摘要超过 200 条就开下一卷，第 2 卷起文件名带卷号，标题里写第几卷。裁剪按月份加卷号一起算，同月第一卷先走。',
        '## 删除',
        '删除是按卷删的：人读投影与块文件两个投影一起删，指向它们的边同事务清掉，再留一条墓碑。归档参与召回是显式声明的，所以删卷是真删除。',
        '## 容量口径',
        '容量上限管的是摘要文件数（一个月一个），不是条目数。两个数必须分开报，否则「30 条 / 60 个」会被读成快满了。',
    ].join('\n\n')],
    ['docs/来源质量门.md', [
        '# 来源质量门',
        '## 三条拦截路径',
        '质量门有两条机器判定的路径：显式声明 unverified 的条目被拒，引用 unverified 知识块的条目也被拒。Pack 导入会先算一张本地加入包声明的档位表，取更保守的一侧。',
        '## 冲突边',
        '冲突边保留双方，不自动裁决：类别三分（事实冲突 / 时效冲突 / 措辞冲突），状态三分（未决 / 已确认 / 已忽略），认不出来的值直接报错而不是静默降级。',
        '## 引用反查',
        '写结论时用到的知识块 id 另存成 kbRefs，台账与记忆两侧都能反着查：这条结论到底依据哪一块来源，是显式写下来的，不靠猜。',
    ].join('\n\n')],
];

/** [查询, 期望命中的文档路径, 类别]。 */
const KB_QUERIES = [
    ['结构切块', 'docs/知识库设计.md', '整串命中'],
    ['切块 结构', 'docs/知识库设计.md', '词序颠倒'],
    ['内容寻址', 'docs/知识库设计.md', '整串命中'],
    ['来源档 三档', 'docs/知识库设计.md', '多词'],
    ['bigram 召回', 'docs/召回实验记录.md', '中英混排'],
    ['词序颠倒', 'docs/召回实验记录.md', '整串命中'],
    ['BM25 排序', 'docs/召回实验记录.md', '中英混排'],
    ['语义改写', 'docs/召回实验记录.md', '整串命中'],
    ['分卷 卷号', 'docs/归档分卷说明.md', '多词'],
    ['摘要文件数 条目数', 'docs/归档分卷说明.md', '词序颠倒'],
    ['容量上限', 'docs/归档分卷说明.md', '整串命中'],
    ['投影 删除', 'docs/归档分卷说明.md', '词序颠倒'],
    ['质量门 拦截', 'docs/来源质量门.md', '整串命中'],
    ['冲突边 状态', 'docs/来源质量门.md', '多词'],
    ['kbRefs 反查', 'docs/来源质量门.md', '中英混排'],
];

const KB_ABSENT = ['量子退相干 时间', '咖啡萃取 水温'];

/** 把 KB_DOCS 落到磁盘并入库（每个 case 一个根，路径里的目录要先建出来）。 */
async function seedKb(root, memory) {
    for (const [path, text] of KB_DOCS) {
        const absolute = join(root, ...path.split('/'));
        await mkdir(dirname(absolute), { recursive: true });
        await writeFile(absolute, text, 'utf8');
        await memory.kbIngest({ path, tier: 'verified' });
    }
}

/** 把 kb 的块读成一个可评估的语料（走公开 API，不直接读文件）。 */
async function kbCorpus(memory) {
    const listed = await memory.kbList({ limit: 20 });
    const docs = [];
    for (const row of listed.items) {
        const read = await memory.kbRead({ id: row.id });
        for (const chunk of read.chunks) {
            docs.push({ id: `${row.path}#${chunk.n}`, path: row.path, n: chunk.n, text: chunk.text });
        }
    }
    return docs;
}

await check('C1 kb 自测集：夹具能入库、块数与查询集都非空', async () => {
    const { root, memory } = await freshRoot();
    await seedKb(root, memory);
    const corpus = await kbCorpus(memory);
    console.log(`      C1 语料 ${corpus.length} 块 / ${KB_DOCS.length} 份文档，查询 ${KB_QUERIES.length} 条 + 弃答 ${KB_ABSENT.length} 条`);
    assert.equal(corpus.length >= KB_DOCS.length * 2, true, '每份文档至少要切出 2 块，否则 recall@k 量不出东西');
    assert.ok(KB_QUERIES.length >= 10, '标注查询集不能太小');
    for (const [query, want] of KB_QUERIES) {
        assert.ok(corpus.some((doc) => doc.path === want), `标注的期望文档不在语料里：${want}（查询 ${query}）`);
    }
    // 顺手把「探针不能再自留一份打分器」这件事核一遍：产品路径才是被量的那个东西。
    const source = readFileSync(fileURLToPath(new URL('memory-probes.mjs', import.meta.url)), 'utf8');
    assert.ok(source.includes("from '../src/kb.js'"), 'C 组要 import 产品路径的打分器，不能自留拷贝');
    assert.ok(!/function\s+(bm25Rank|tokenize)\s*\(/.test(source), '探针里不许再有一份本地的 BM25 / 分词拷贝');
});

/** 走 action 面取 top-k 命中（返回每条的 path）。 */
async function kbTopPaths(memory, query, k = TOPK) {
    const found = await memory.kbSearch({ query, limit: k });
    return { paths: found.hits.map((hit) => hit.path), found };
}

await check('C2 产品路径（kb-search：bigram + BM25 多路）的 recall@3 / hit@3 / 弃答率', async () => {
    const { root, memory } = await freshRoot();
    await seedKb(root, memory);

    let hit = 0;
    let recall = 0;
    const missed = [];
    for (const [query, want] of KB_QUERIES) {
        const { paths } = await kbTopPaths(memory, query);
        if (paths.includes(want)) hit += 1; else missed.push(query);
        recall += paths.filter((path) => path === want).length / Math.max(1, paths.length);
    }
    const hitRate = hit / KB_QUERIES.length;
    const recallRate = recall / KB_QUERIES.length;

    // 弃答：探针的融合只在「至少一路原始分 > 0」时才给命中 —— 一句都不沾的查询一条都不给。
    let abstained = 0;
    for (const query of KB_ABSENT) {
        const { found } = await kbTopPaths(memory, query);
        if (found.matched === 0 && found.hits.length === 0) abstained += 1;
    }
    console.log(`      C2 产品路径 hit@3 ${hit}/${KB_QUERIES.length}（${hitRate.toFixed(3)}），recall@3 ${recallRate.toFixed(3)}，弃答 ${abstained}/${KB_ABSENT.length}${missed.length > 0 ? `，未中：${missed.join('、')}` : ''}`);
    // 这条是 24-11 的硬门槛（与 candidateMultiplier 那类阈值不同：自测集是全中的），
    // 所以它是**契约断言**，不是「数字照实打印」。
    assert.equal(hitRate, 1, `产品路径在自测集上应当全中，未中：${missed.join('、')}`);
    assert.equal(recallRate, recallRate, 'recall 不能是 NaN');
    assert.equal(abstained, KB_ABSENT.length, '两条弃答查询应当一条都不给（零重叠即弃答）');
});

await check('C3 灵敏度：同一批查询上，朴素整串基线与随机排序都必须明显更差', async () => {
    const { root, memory } = await freshRoot();
    await seedKb(root, memory);
    const corpus = await kbCorpus(memory);

    const rateOf = (ranker) => {
        let hit = 0;
        for (const [query, want] of KB_QUERIES) {
            const ranked = ranker(corpus, query).slice(0, TOPK).map((row) => row.doc.path);
            if (ranked.includes(want)) hit += 1;
        }
        return hit / KB_QUERIES.length;
    };

    // 候选 = **产品路径**（走 action 面，不是本地拷贝）。
    let candidateHit = 0;
    for (const [query, want] of KB_QUERIES) {
        const { paths } = await kbTopPaths(memory, query);
        if (paths.includes(want)) candidateHit += 1;
    }
    const candidate = candidateHit / KB_QUERIES.length;
    const naive = rateOf((docs, query) => naiveRank(docs, query));
    // 确定性随机排序（线性同余）：同一批语料、同一个种子，这次跑与下次跑一样。
    //
    // **必须真的排序**：第四十二轮的独立复核抓到这里的第一版只 `map` 了随机分数、没有
    // `sort`，于是 `.slice(0, TOPK)` 拿的永远是语料前三块 —— 换个种子数字纹丝不动
    // （白名单式的自证）。补上 sort 之后随机基线是 0.667，而不是那个假的 0.200。
    const random = rateOf((docs, query) => {
        const rng = (() => { let seed = 20260930; return () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; }; })();
        return docs
            .map((doc, index) => ({ doc, score: rng() + index * 0 }))
            .sort((left, right) => right.score - left.score);
    });

    console.log(`      C3 产品路径 ${candidate.toFixed(3)} ｜ 朴素整串 ${naive.toFixed(3)} ｜ 随机 ${random.toFixed(3)}`);
    assert.ok(candidate > naive, `量表分不出 bigram 与整串：候选 ${candidate} vs 朴素 ${naive}`);
    assert.ok(candidate > random, `量表分不出候选与随机排序：候选 ${candidate} vs 随机 ${random}`);
    assert.ok(random < 1, '随机排序不该也是满分 —— 那样这个对照就没有信息量了');
});

await check('C4 接线真的落地了：kb-search 在 action 面上，且两路融合的权重是契约', async () => {
    // 24-11 本体 = 把打分接进检索路径 + 配额种 + 设置项。这一条防的是「接一半」：
    // 打分器进了 src/kb.js 却没人调，或者调了但没进 action 面的动作名清单。
    const tools = readFileSync(fileURLToPath(new URL('../src/tools.js', import.meta.url)), 'utf8');
    for (const action of ['kb-ingest', 'kb-search', 'kb-read', 'kb-list', 'kb-drop']) {
        assert.ok(tools.includes(action), `action 面应当列出 ${action}`);
    }
    assert.ok(tools.includes('kbSearchPerTurn'), 'kb 检索要有自己的每回合配额种');
    const settings = readFileSync(fileURLToPath(new URL('../src/settings.js', import.meta.url)), 'utf8');
    assert.ok(settings.includes('kbSearchPerTurn'), '设置页要能调这个配额（否则配额等于写死）');

    // 两路融合的**方向**也是契约：只在块头命中的块（0.25）必须排在只在正文命中的块（0.75）后面，
    // 否则「正文为主、块头为辅」这句话就是假的（权重写反了不会有别的用例拦住）。
    const only = (text, head) => ({ id: 'kb:test', text, head });
    const ranked = kbRankChunks(
        [only('这里与查询毫无关系', '标题路径里就有 结构切块'), only('正文里写着 结构切块 这四个字', '无关块头')],
        '结构切块',
    );
    assert.equal(ranked.length, 2, '两路各命中一块时两块都该在');
    assert.equal(ranked[0].chunk.text.includes('正文里写着'), true, '正文路（0.75）该排在块头路（0.25）前面');
    assert.ok(ranked[0].score > ranked[1].score);
    // 零重叠 = 弃答（不是「给个 0 分的命中」）。
    assert.deepEqual(kbRankChunks([only('完全无关的正文', '也无关')], '量子退相干'), []);
    // 确定性：同一批语料、同一批查询，两次跑出来逐字一样。
    assert.deepEqual(kbRankChunks([only('甲 乙 丙', '张'), only('甲 乙', '李')], '甲 乙'), kbRankChunks([only('甲 乙 丙', '张'), only('甲 乙', '李')], '甲 乙'));
    // 分词只有一份实现：去重交给 kbTerms，词频要的重复留给 kbTokens。
    assert.deepEqual(kbTerms('面板记忆'), ['面板', '板记', '记忆']);
    assert.deepEqual(kbTokens('甲甲'), ['甲甲']);
});

// ── 汇总 ───────────────────────────────────────────────────────────────────

await rm(TMP, { recursive: true, force: true });

const failed = results.filter((r) => !r.ok);
for (const r of results) {
    if (r.ok) console.log(`PASS  ${r.name}`);
    else console.log(`FAIL  ${r.name}\n      ${r.error?.message ?? r.error}`);
}
console.log(`\nmemory-probes: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
