/**
 * 提示词注入预算与「固定长句只说一次」的回归测试（第十七轮）。
 *
 * 这一份守的是**上下文成本**，不是功能：办公插件往模型上下文里写的东西会留在
 * 前缀里被后续每一次请求重发，所以「体积」与「重复」都是要被钉住的量。
 * 六节断言口径：
 *
 *   1. **默认层必须比全文层小得多**（每个话题至少 2x）。office_help 默认只给
 *      最小事实集（接口签名 + 必需边界），例子与解释留在 detail:true —— 默认层
 *      一旦被谁悄悄改回全文，这条会当场红。
 *   2. **关键接口不许在分层里丢**：默认层仍要能查出每个 office.* 操作名，
 *      否则「省字节」就变成了「模型写不出对的调用」。
 *   3. **固定长句一个会话只说一次**：派工反馈里的「不要把结果原文搬进对话」、
 *      直查反馈里的两条尾巴，第二次调用不再重复；同时每次都要有的最小提醒
 *      （按 URL 核对 / 外部内容不可信）不能一起被省掉。
 *   4. **工具面总字节**（第二十七轮新增）：7 个工具的 `{name, description, parameters}`
 *      之和不超过实测值钉死的预算，`office_memory` 单项也不许超 —— 它是常驻前缀里
 *      最大的一块，瘦身收益最高，所以断言里直接钉住「它必须是单项最大」。
 *   5. **省字节不能省语义**（第二十七轮新增）：13 个动作名与关键事实、关键语义
 *      （kind 的「只看某一类」、import 的「只增不改」、id 的位置「第二段」）都要留。
 *   6. **按需投影的预算**（第三十六轮新增）：带话题信号的投影要收掉一大半、
 *      「完全未变化」的投影压到一行级、台账登记后要贴回来、心跳重贴不筛。
 *
 * 跑法：node test/injection-budget.mjs
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { buildHelp } from '../src/docs.js';
import { renderDispatch } from '../src/search.js';
import { renderRun } from '../src/tools.js';
import { buildTools } from '../src/tools.js';
import { resolveConfig } from '../src/config.js';
import { hintOnce, hintSeen, resetHints, projectDigest, resetProjection, signalOf, REFRESH_EVERY } from '../src/projection.js';
import { createMemory } from '../src/memory.js';

const results = [];
async function check(name, fn) {
    resetHints();
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const bytes = (text) => Buffer.byteLength(String(text ?? ''), 'utf8');

/** 所有话题：默认层与全文层都要能出，且默认层不许比全文层大。 */
const TOPICS = ['', 'guide', 'run', 'word', 'excel', 'ppt', 'tex', 'pdf', 'python', 'av', 'files', 'cache', 'archive', 'preview', 'image', 'theme', 'search', 'memory', 'settings'];

// ── 1. 分层：每个话题的默认层都比全文层小得多 ──────────────────────────────

await check('office_help：默认层都更小，大话题至少收掉 20% / 30%', async () => {
    const report = [];
    for (const topic of TOPICS) {
        const compact = await buildHelp(topic);
        const full = await buildHelp(topic, { detail: true });
        assert.ok(compact.text.trim() !== '', `${topic || '(index)'} 默认层不能为空`);
        const compactBytes = bytes(compact.text);
        const fullBytes = bytes(full.text);
        // 三条口径：
        //   · 每个话题的默认层都必须更小（挡住「默认层悄悄退回全文」）；
        //   · 全文 ≥ 3 KB 的话题至少收掉 20%；
        //   · 全文 ≥ 8 KB 的话题至少收掉 30%（ppt / memory 这两个才是大头）。
        assert.ok(compactBytes < fullBytes, `${topic || '(index)'} 默认层不该不小于全文层`);
        if (fullBytes >= 3000) {
            assert.ok(compactBytes * 5 <= fullBytes * 4, `${topic || '(index)'} 默认层 ${compactBytes}B 对全文层 ${fullBytes}B 收得太少（< 20%）`);
        }
        if (fullBytes >= 8000) {
            assert.ok(compactBytes * 10 <= fullBytes * 7, `${topic || '(index)'} 默认层 ${compactBytes}B 对全文层 ${fullBytes}B 收得太少（< 30%）`);
        }
        report.push(`${topic || 'index'} ${compactBytes}/${fullBytes}`);
    }
    return report.join(' ');
});

await check('office_help：ppt 与 memory 这两个大话题被压到 1/3 以下', async () => {
    for (const topic of ['ppt', 'memory']) {
        const compact = bytes((await buildHelp(topic)).text);
        const full = bytes((await buildHelp(topic, { detail: true })).text);
        assert.ok(compact * 3 <= full, `${topic} 默认层 ${compact}B 没到全文层 ${full}B 的 1/3`);
    }
});

await check('office_help：索引与 guide 也走默认层，并指路 detail:true', async () => {
    const index = await buildHelp('');
    const guide = await buildHelp('guide');
    assert.ok(index.text.includes('detail:true') || index.text.includes('detail: true'), '索引要告诉模型怎么拿全文');
    assert.ok(bytes(index.text) * 2 <= bytes((await buildHelp('', { detail: true })).text));
    assert.ok(bytes(guide.text) * 2 <= bytes((await buildHelp('guide', { detail: true })).text));
});

// ── 2. 分层不许丢接口 ──────────────────────────────────────────────────────

await check('默认层仍给出每个格式的操作名（省字节不能省接口）', async () => {
    const wanted = {
        word: ['office.word', 'create', 'builder.heading', 'builder.table', 'builder.formula', 'builder.toc', 'edit'],
        excel: ['office.excel', 'sheet.stats', 'sheet.summary', 'sheet.formula', 'edit'],
        ppt: ['office.ppt', 'deck.cards', 'deck.kpi', 'deck.imageText', 'deck.timeline', 'deck.iconGrid', 'readSlides', 'revise', 'insertImages', 'images'],
    };
    for (const [topic, needles] of Object.entries(wanted)) {
        const text = (await buildHelp(topic)).text;
        for (const needle of needles) assert.ok(text.includes(needle), `${topic} 默认层缺 ${needle}`);
    }
});

await check('默认层的 pdf / python / search / settings / memory 仍讲清关键事实', async () => {
    const pdf = (await buildHelp('pdf')).text;
    for (const needle of ['office.pdf.info', 'office.pdf.text', 'office.pdf.pages', 'read_image', 'hasTextLayer', 'reused']) {
        assert.ok(pdf.includes(needle), `pdf 默认层缺 ${needle}`);
    }
    const python = (await buildHelp('python')).text;
    for (const needle of ['office.python.check', 'office.python.run', 'office.python.file', 'OUT_DIR', 'savefig', 'use_cjk_font', 'logFiles.stdout']) {
        assert.ok(python.includes(needle), `python 默认层缺 ${needle}`);
    }
    const av = (await buildHelp('av')).text;
    for (const needle of ['office.av.check', 'office.av.info', 'office.av.transcribe', 'office.av.frames', 'office.av.extract', 'read_image', 'SenseVoice', 'ffmpeg', '时间戳']) {
        assert.ok(av.includes(needle), `av 默认层缺 ${needle}`);
    }
    const search = (await buildHelp('search')).text;
    for (const needle of ['office_search_run', 'office_search_brief', 'office_search_dispatch', 'office_parse_findings', '泛搜', '茧房']) {
        assert.ok(search.includes(needle), `search 默认层缺 ${needle}`);
    }
    const settings = (await buildHelp('settings')).text;
    for (const needle of ['工具开关', 'maxParallel', 'resultLimit', 'subagentTools', 'outputDir', 'defaultTheme', 'userLimitBytes', 'ledgerLimit']) {
        assert.ok(settings.includes(needle), `settings 默认层缺 ${needle}`);
    }
    const memory = (await buildHelp('memory')).text;
    for (const needle of ['热记忆', '台账', '归档', 'replace', 'oldText', '不记', '记忆目录', '按工作目录']) {
        assert.ok(memory.includes(needle), `memory 默认层缺 ${needle}`);
    }
});

// ── 3. 固定长句一个会话只说一次 ────────────────────────────────────────────

// 第十八轮 P1-2 / 18-12：`office_run.script` 的三个真实用例（docx / xlsx / Python 绘图）
// 只放在 **docs.js 的全文层**。理由与 P1-2 的实测一致（工具定义内给例子 72%→90%，
// 但例子进 schema 就每步吃 token）：`script` 是「schema 合法但用法不显然」的重灾区，
// 例子必须有；但它是**查一次就够**的规格，不该跟默认层一起常驻前缀。
await check('office_help(run)：三个真实用例只在全文层，默认层不背例子（P1-2 / 18-12）', async () => {
    const compact = (await buildHelp('run')).text;
    const full = (await buildHelp('run', { detail: true })).text;
    // 三个用例各自的标记与真实调用形状：默认层一个都不许有，全文层一个都不许少。
    const markers = ['【例 1】Word 报告', '【例 2】Excel 表', '【例 3】Python 画图'];
    for (const marker of markers) {
        assert.ok(!compact.includes(marker), `默认层不该带例子：${marker}`);
        assert.ok(full.includes(marker), `全文层缺用例：${marker}`);
    }
    for (const call of ['office.word.create(', 'office.excel.create(', 'office.python.run(']) {
        assert.ok(!compact.includes(call), `默认层不该带例子里的调用形状：${call}`);
        assert.ok(full.includes(call), `全文层缺调用形状：${call}`);
    }
    // 默认层要指路到全文层，否则模型不知道例子在哪。
    assert.ok(compact.includes("detail: true"), '默认层要指路 detail:true');
    // 体积有界：例子是「几百字节」的量级，涨到几千字节就该拆到别的话题去。
    assert.ok(bytes(full) <= 5200, `run 全文层 ${bytes(full)}B 超上限：例子要有界`);
    assert.ok(bytes(full) > bytes(compact) * 2, '全文层仍要明显厚于默认层');
});

await check('派工反馈：同一条失败原因不再逐渠道重复七遍', () => {
    const reason = 'tools.restrict() names unknown global tools "web_search", "advanced_search", "platform_search", "web_fetch"; known global tools: ' + 'x'.repeat(200);
    const jobs = Array.from({ length: 7 }, (unused, index) => ({ channel: `渠道${index + 1}`, path: `c${index + 1}.md`, ok: false, error: reason }));
    const text = renderDispatch({ topic: '某事件', jobs, sessionId: 's1' });
    const occurrences = text.split(reason).length - 1;
    assert.equal(occurrences, 1, `失败原因只该原样出现一次，实际 ${occurrences} 次`);
    assert.ok(text.includes('7 个渠道同因'), '要说清有几个渠道同因');
    for (const job of jobs) assert.ok(text.includes(job.channel), `合并后仍要列出渠道：${job.channel}`);
});

await check('派工反馈：「不要把结果原文搬进对话」一个会话只说一次', () => {
    const value = { topic: '某事件', jobs: [{ channel: '泛搜', path: 'a.md', ok: true }], sessionId: 's-dup' };
    const first = renderDispatch(value);
    const second = renderDispatch(value);
    assert.ok(first.includes('不要把结果原文搬进对话'), '第一次要给这条纪律');
    assert.ok(!second.includes('不要把结果原文搬进对话'), '第二次不再重复');
    assert.ok(second.includes('office_parse_findings'), '第二次仍要指向下一步');
    assert.ok(renderDispatch({ ...value, sessionId: 's-other' }).includes('不要把结果原文搬进对话'), '换一个会话要重新给');
});

await check('office_search_run：同一会话第二次的反馈明显更短，最小提醒仍在', async () => {
    const seam = {
        async search(request) {
            return {
                sources: [
                    { url: 'https://news-a.example.com/' + encodeURIComponent(request.query), title: '来源甲', snippet: '引用片段' },
                    { url: 'https://news-b.example.com/2', title: '来源乙' },
                ],
                truncated: false,
            };
        },
        async fetch(request) {
            return {
                url: request.url,
                statusCode: 200,
                body: { kind: 'html', content: '<p>来源乙的正文段落，长度足够当摘录用的一行文字。</p>' },
                truncated: false,
            };
        },
    };
    const ctx = new Proxy({ inject: ['tools'] }, {
        get(target, prop) {
            if (prop === 'get') return (name) => (name === 'web' ? seam : undefined);
            if (prop in target) return target[prop];
            throw new Error(`cannot get property "${String(prop)}" without inject`);
        },
    });
    const root = mkdtempSync(join(tmpdir(), 'office-budget-'));
    try {
        // 默认顺序里免 Key 抓取在前（会真发 HTTP）；这条用例靠假接缝，钉成 seam-only。
        const tools = buildTools(resolveConfig({ search: { providerOrder: ['seam'] } }), ctx);
        const run = tools.find((item) => item.name === 'office_search_run');
        const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
        const first = await run.execute({ queries: ['某事件 通报'], out: 'a.md' }, exec);
        const second = await run.execute({ queries: ['某事件 通报'], out: 'b.md' }, exec);
        assert.ok(first.text.includes('要按渠道逐项覆盖'), '第一次要给渠道引导');
        assert.ok(!second.text.includes('要按渠道逐项覆盖'), '第二次不再重复渠道引导');
        assert.ok(second.text.includes('不是结论') || second.text.includes('未核实材料'), '最小安全提醒必须留着');
        assert.ok(second.text.includes('https://news-a.example.com/') || second.text.includes('来源甲'), '来源清单照旧要给');
        assert.ok(
            Buffer.byteLength(second.text, 'utf8') < Buffer.byteLength(first.text, 'utf8'),
            '第二次的反馈要更短',
        );
    } finally {
        rmSync(root, { recursive: true, force: true });
    }
});

// ── 4. tool 面字节：常驻前缀的另一半开销 ──────────────────────────────────
//
// 口径（第二十四轮 P2-12 提出「没有断言守工具面」，第二十七轮补上）：
// `JSON.stringify({ name, description, parameters })` 的 UTF-8 字节之和 —— 这正是宿主
// 每次请求都要重发的那一份。第十七轮的断言只守 office_help 的分层，工具面本身没有界。
//
// 数字是**实测值 + 小余量**，不是估算：第二十四轮实测 13,370 B（office_memory 5,592 B），
// 第二十七轮瘦身后实测 12,649 / 4,517 B，阈值定在 12,802 / 4,670（余量 153 B）；
// 第二十八轮显式上调（实测 12,951 / 4,819，阈值 13,100 / 4,960，余量 149 / 141）——
// 加的是 `read` 的 since / until 两个参数（时间维度入口）；第二十九轮再次上调
// （实测 13,110 / 4,819，阈值 13,300 / 4,960），涨的是 `office_search_run` 的描述（+159 B）。
//
// **第三十轮再次显式上调**（工具从七个变九个：新增 office_web_search / office_web_fetch
// 两个抓取工具，办公 preset 同时收掉了宿主的 tool-web 行）。理由：这两个工具是「子代理 +
// 网页抓取」这条最初设想的落点，没有它们的描述，模型在没 tool-web 的办公会话里就不知道
// 「查一个事实点 / 打开一个页面」还能用什么；描述里也写死了「不落盘、不走三方 Key」
// 这两条路由语义。可枚举的例子仍旧只留在 office_help('search')。

// **第三十三轮再次显式上调**（实测 14,980 → 上调后仍只涨两个抓取工具的描述）：两个抓取
// 工具的通道顺序多了「宿主 web 服务兜底」这条路由语义 —— 第三十三轮的真实故障就是
// 它们被钉死在两条不通的抓取通道上、组合里明明有能用的接缝。写清「先抓取、通不了才兜底」
// 才能让模型在通道报错时知道还有一条路，而不是判成「这个环境不能联网检索」。
// 每次上调都要重算余量：这一轮余量 220 B（15,200 − 14,980）。

// **第三十四轮再次显式上调**（实测 15,598 → 阈值 15,800，余量 202 B）：新增「站点清单 +
// 站点优先」，`office_web_search` 与 `office_search_run` 各多一个 `sites` 参数。它的四种
// 形态（不传 / 类型 id / 域名数组 / false）不写清，模型就只能祷告式地猜「优先站点」怎么表达；
// 塞进 office_help 又来不及 —— 这个参数是**每次调用都要写的**，属于工具面该付的钱。
// 两个工具的 `sites` 描述已按字节纪律压到同一条最短口径（不在工具面里举例子、不列站点名）。
// 每次上调都要重算余量：这一轮 15,800 − 15,598 = 202 B。

// **第四十一轮：这次预算是往下调的**（实测 15,548 / 4,736 → 阈值 15,650 / 4,840）。
// 本轮往 `office_memory` 里加了知识库的四个动作（kb-ingest / kb-read / kb-list / kb-drop）、
// 来源档 `tier`、冲突边的 `conflict` / `state` —— 按第二十四轮 §5.1 的规矩，
// **先瘦身再加动作**：把工具描述与参数描述压到「模型写错就调不通」的最小事实集，
// 于是总字节与单项都**低于上一轮**（15,631 → 15,548、4,819 → 4,736）。
// 这条「只减不增」是本轮刻意守住的口径：新增能力不该靠加宽预算来落地。
// 余量口径与历轮一致（每项约 110 B）：总 102 B、单项 104 B。
// **第四十三轮：预算不动，实测值往下走**（实测 15,548 → **15,336** / 4,736）。
// 18-11 把 `office_search_*` 四兄弟的首句改成「我是链路哪一段、输入从哪来」，
// 顺手删掉了首句与尾句重复的路由话术（`office_search_run` 少 228 B、dispatch 少 18 B），
// 所以这轮**不需要动预算**：总阈值仍是 15,650（余量从 102 B 变成 314 B）。
// 口径照旧：新增能力不该靠加宽预算来落地，能减就减。
//
// **第四十五轮：预算仍不动，靠瘦身把新动作塞进去**（实测 15,336 → **15,418** / 4,736 → **4,818**）。
// 24-11 往 `office_memory` 里加了 kb 检索（动作名 `kb-search`、参数面五处提到它：`query` /
// `limit` / `tier` / `path` 与描述那句）。按第四十一轮那条「先瘦身再加动作」的规矩，顺手把
// kb 那句压成一行、并删掉 `query` 描述里与工具描述重复的半句（「read 带 query 走召回分档」
// ——工具描述里已经写了「带 query 的读取走召回分档并占每回合配额」），
// 所以总字节与单项各只涨了 82 B —— 两项都在原预算内（余量 232 B / 22 B）。
//
// **第四十七轮：预算仍不动**（实测 15,418 → **15,451** / 4,818 不变）。
// `office_help` 的 `topic` 描述改成**从真源现取**话题清单（docs.js 的
// `availableHelpTopics`）：原来手写的 14 个话题少了 av / preview / image / archive，
// 而「没有这个话题」的提示里还重复列了一次 `tex`（18-23 的格式指令腐烂）。
// 换来自同步的代价是 +33 B，仍在原预算内（余量 199 B）。
//
// **第四十八轮：预算上调，理由写在下面**（实测 15,451 → **15,624** / 4,886）。
// 这一轮 office_memory 多了两处**语义**（不是文案润色）：① `sunk` 动作
// （新-10：容量维持只报条数、热层忘了哪几条不可见）；② `content` 的单条上限
// （新-11：写入期质量门，阈值从 `DEFAULT_ENTRY_LIMIT_BYTES` 现取，不手写数字）。
// 两处都先按「先瘦身再加动作」压过一轮（content 描述从 205 → 147 B，去掉了
// 「超了怎么办」的重复说明 —— 拒绝时的报错里已经给了三条去处），
// 剩下的 46 B 是净增的语义：总预算 +120 B、单项 +60 B，留 4.8% 余量。
const TOOL_SURFACE_BUDGET_BYTES = 15770;
const OFFICE_MEMORY_BUDGET_BYTES = 4950;

await check('工具面字节：九个工具的常驻字节在预算内，且 office_memory 是单项最大的一块', () => {
    const tools = buildTools(resolveConfig({}));
    assert.ok(tools.length >= 9, `工具面至少要含九个工具，实际 ${tools.length}`);
    const measured = tools.map((tool) => ({
        name: tool.name,
        bytes: bytes(JSON.stringify({ name: tool.name, description: tool.description, parameters: tool.parameters })),
    }));
    const total = measured.reduce((sum, item) => sum + item.bytes, 0);
    const memory = measured.find((item) => item.name === 'office_memory');
    const report = measured.map((item) => `${item.name} ${item.bytes}`).join(' ');
    assert.ok(total <= TOOL_SURFACE_BUDGET_BYTES,
        `工具面合计 ${total}B 超过预算 ${TOOL_SURFACE_BUDGET_BYTES}B：${report}。`
        + '要么瘦身，要么在 test/injection-budget.mjs 里显式调高预算并写清理由。');
    assert.ok(memory !== undefined, 'office_memory 必须在工具面里');
    assert.ok(memory.bytes <= OFFICE_MEMORY_BUDGET_BYTES,
        `office_memory 单项 ${memory.bytes}B 超过预算 ${OFFICE_MEMORY_BUDGET_BYTES}B`);
    // 不要写成「占比 ≤ 37%」那种近乎恒真的断言：绝对上限已经把它钉住了，
    // 这里只钉一个会随现实变化的事实 —— 它是单项最大的一块（改小它收益最高）。
    const biggest = measured.slice().sort((left, right) => right.bytes - left.bytes)[0];
    assert.equal(biggest.name, 'office_memory',
        `单项最大的工具是 ${biggest.name}（${biggest.bytes}B）而不是 office_memory —— 说明瘦身目标该换一个了`);
    return `${report} 合计 ${total}B`;
});

await check('工具面字节：省字节不能省掉 office_memory 的动作名与关键语义', () => {
    const tool = buildTools(resolveConfig({})).find((item) => item.name === 'office_memory');
    const text = `${tool.description}\n${JSON.stringify(tool.parameters)}`;
    for (const action of ['read', 'add', 'replace', 'remove', 'log', 'link', 'unlink', 'related', 'entities', 'status', 'export', 'import', 'migrate']) {
        assert.ok(text.includes(action), `瘦身把动作名「${action}」省掉了：模型会写不出这个调用`);
    }
    for (const needle of ['旧条目', 'layer', 'query', 'oldText', 'id', '配额']) {
        assert.ok(text.includes(needle), `office_memory 的参数面缺关键事实「${needle}」`);
    }
    // 语义也要钉住：光有参数名不够（kind 的「只看某一类」、import 的「只增不改」都是
    // 模型决定怎么调的依据，第二十七轮瘦身一度把它们删掉，是审查发现的）。
    for (const needle of ['只看某一类', '只增不改', '第二段']) {
        assert.ok(text.includes(needle), `瘦身把语义「${needle}」省掉了`);
    }
    // 第二十八轮新增的两个参数也要守住：时间窗是「本月记的」这类查询的唯一入口，
    // 只删描述不删参数的话，模型照样不知道该传什么。
    for (const needle of ['since', 'until', '时间窗']) {
        assert.ok(text.includes(needle), `时间窗参数面缺「${needle}」：模型会不知道怎么按时间过滤`);
    }
    // 第四十一轮：知识库的四个动作与「来源三档」也是**写不出来就调不通**的那类事实。
    // 它们必须落在**参数面**（action 的取值清单 / tier 的取值），因为模型是从参数面里
    // 挑动作名与枚举值的；把描述压到最小事实集之后仍然留下这几项，下一轮瘦身不能顺手删。
    const params = JSON.stringify(tool.parameters);
    for (const needle of ['kb-ingest', 'kb-search', 'kb-read', 'kb-list', 'kb-drop', 'tier', 'unverified']) {
        assert.ok(params.includes(needle), `知识库接口面缺「${needle}」：模型会不知道 kb 怎么用、或不知道来源档怎么填`);
    }
    // 冲突边的三态（保留双方、不自动裁决）同样要留在参数面：模型不知道该传什么时，
    // 会把一条真实的冲突记成普通 related 边。
    for (const needle of ['conflict', 'state', 'contradicts']) {
        assert.ok(params.includes(needle), `冲突边参数面缺「${needle}」`);
    }
});

await check('工具面字节：两条检索路径的分工与 PDF 语义不能被瘦身掉（第三十轮）', () => {
    const tools = buildTools(resolveConfig({}));
    const search = tools.find((item) => item.name === 'office_search_run');
    assert.ok(search !== undefined, 'office_search_run 必须在工具面里');
    const text = `${search.description}\n${JSON.stringify(search.parameters)}`;
    // 没有宿主 tool-web 之后，「快查走哪两个工具」是路由语义：
    // 省掉它，模型会在抓取工具与落盘直查之间反复试探。
    for (const needle of ['office_web_search', 'office_web_fetch', '落盘']) {
        assert.ok(text.includes(needle), `office_search_run 描述缺路由事实「${needle}」`);
    }
    // PDF 来源的正文来自 office.pdf 而不是网页预处理：写错了会让模型按 HTML 去理解摘录。
    assert.ok(text.includes('PDF') && text.includes('office.pdf'), 'PDF 来源的抽取路径要写清');
    for (const needle of ['provider', 'preprocess']) {
        assert.ok(text.includes(needle), `按次点名通道 / 改预处理强度的入口不能省：「${needle}」`);
    }
    // 两个抓取工具自己的路由语义：免 Key 抓取、不落盘、PDF 的去处。
    for (const name of ['office_web_search', 'office_web_fetch']) {
        const tool = tools.find((item) => item.name === name);
        assert.ok(tool !== undefined, `抓取工具 ${name} 必须在工具面里`);
        const body = `${tool.description}\n${JSON.stringify(tool.parameters)}`;
        for (const needle of ['免 Key', 'office_search_run']) {
            assert.ok(body.includes(needle), `${name} 描述缺路由事实「${needle}」`);
        }
    }
});

// ── 5. office_run 反馈不许回灌模型自己刚写的正文 ────────────────────────────

await check('office_run 反馈：脚本返回里的 outline 不再原样回灌', () => {
    const outline = Array.from({ length: 41 }, (unused, index) => `P${index + 1} cover：标题${index + 1}`);
    const returned = { ok: true, path: 'a.pptx', outline, note: '保留我' };
    const text = renderRun({
        ok: true,
        elapsedMs: 12,
        files: [],
        otherFiles: [],
        returned,
        logs: [],
        notes: [],
        warnings: [],
        cache: { dir: '.office/cache', kept: 0, keptBytes: 0, hits: 0, pruned: [], ttlMinutes: 720 },
    });
    for (const line of outline) assert.ok(!text.includes(line), `outline 不该回灌：${line}`);
    assert.ok(text.includes('共 41 项'), '要报出 outline 有条数');
    assert.ok(text.includes('保留我'), 'outline 之外的字段要留下');
    assert.ok(bytes(text) < 400, `这一条反馈要够小，实际 ${bytes(text)}B`);
});

await check('office_run 反馈：脚本返回超长时按小上限截断', () => {
    const text = renderRun({
        ok: true,
        elapsedMs: 12,
        files: [],
        otherFiles: [],
        returned: { blob: 'x'.repeat(5000) },
        logs: [],
        notes: [],
        warnings: [],
        cache: { dir: '.office/cache', kept: 0, keptBytes: 0, hits: 0, pruned: [], ttlMinutes: 720 },
    });
    assert.ok(text.includes('…'), '要能看出是被截断的');
    assert.ok(bytes(text) < 2000, `截断上限要够小，实际 ${bytes(text)}B`);
});

await check('office_run 反馈：缓存块压成一行，TTL 说明不再占三行', () => {
    const text = renderRun({
        ok: true, elapsedMs: 12, files: [], otherFiles: [], returned: null, logs: [], notes: [], warnings: [],
        cache: { dir: '.office/cache', kept: 131, keptBytes: 19890000, hits: 2, hitBytes: 4000, pruned: ['a', 'b'], prunedBytes: 12400, ttlMinutes: 720 },
    });
    const cacheLines = text.split('\n').filter((line) => line.includes('缓存：') || line.trimStart().startsWith('720 分钟'));
    assert.equal(cacheLines.length, 1, `缓存块只该一行，实际 ${cacheLines.length} 行`);
    assert.ok(cacheLines[0].includes('命中复用'), '命中信息要留下');
    assert.ok(cacheLines[0].includes('office.cache.clear'), '清空入口要留下');
});

// ── 6. 按需投影的预算（第三十六轮：需要用什么才给什么）────────────────────

await check('按需投影：真实量级夹具下，信号过滤、台账折叠与心跳重贴都在预算内', async () => {
    resetProjection();
    const root = mkdtempSync(join(tmpdir(), 'office-injection-'));
    try {
        const memory = createMemory({ root });
        // 夹具按本机实测量级造（单条 300–1500 B）：折叠行是固定开销，条目太短时
        // 大小断言会失真 —— 与 test/memory.mjs 的「省略形态」用例同一口径。
        const filler = '这条约定用来把热记忆正文撑到本机实测量级，让大小断言立在真实刻度上而不是感觉上。'.repeat(5);
        await memory.mutate({ action: 'add', target: 'user', content: `偏好：PPT 汇报先给结论。${filler}` });
        for (let index = 0; index < 6; index += 1) {
            await memory.mutate({ action: 'add', target: 'project', content: `约定 ${index}：检索与面板的事项。${filler}` });
        }
        await memory.log([{ path: '汇报.pptx', format: 'ppt', theme: 'business', purpose: '季度汇报' }]);
        const digest = await memory.digest({ recentLedger: 3 });

        // ① 有信号时：贴正文的形态也要收掉一大半（1/7 条命中 + 折叠行）。
        const filtered = projectDigest(digest, { sessionId: 'budget-filtered', context: 'help', signal: signalOf({ topic: 'ppt' }) });
        const full = projectDigest(digest, { sessionId: 'budget-full', context: 'help' });
        assert.ok(bytes(filtered.text) * 5 <= bytes(full.text) * 3,
            `带信号的投影该收掉一大半：${bytes(filtered.text)}B / ${bytes(full.text)}B`);
        assert.ok(filtered.text.includes('PPT 汇报先给结论'), '命中的条目要在');
        assert.ok(filtered.text.includes('与「ppt」相关的条目 0 条，6 条全部折叠'), '无关条目要有折叠行');

        // ② 完全未变化（热记忆与台账指纹都没动）：投影压到一行级。
        const unchanged = projectDigest(digest, { sessionId: 'budget-full', context: 'help' });
        assert.equal(unchanged.mode, 'unchanged');
        assert.ok(bytes(unchanged.text) <= 1400, `「未变化」投影要压小，实际 ${bytes(unchanged.text)}B`);
        assert.ok(!unchanged.text.includes('主动记录'), '两部分都省略时固定尾巴不再重复');
        assert.ok(!unchanged.text.includes('汇报.pptx'), '台账行不重复贴');
        assert.ok(unchanged.text.includes('与上次投影相同'), '折叠要说清并留入口');

        // ③ 台账指纹变了（登记新产物）：台账行要贴回来 —— 省字节不能省掉
        //    「office_run 之后能看到刚登记的那条」这条反馈。
        await memory.log([{ path: '第二份.pptx', format: 'ppt', purpose: '复检' }]);
        const afterLog = projectDigest(await memory.digest({ recentLedger: 3 }), { sessionId: 'budget-full', context: 'help' });
        assert.equal(afterLog.reason, 'ledger');
        assert.ok(afterLog.text.includes('第二份.pptx'), '新登记的台账行要贴出来');

        // ④ 心跳重贴必须不筛：先前折叠的条目要能回到上下文（完整性兜底）。
        //    循环到出现 refresh 为止（起点的 since 计数不一定从 0 开始）。
        let last = afterLog;
        for (let index = 0; index < REFRESH_EVERY + 3 && last.reason !== 'refresh'; index += 1) {
            last = projectDigest(await memory.digest({ recentLedger: 3 }), { sessionId: 'budget-full', context: 'help' });
        }
        assert.equal(last.reason, 'refresh');
        assert.ok(last.text.includes('约定 5：'), '心跳重贴把被折叠的条目带回来');
        return `filtered ${bytes(filtered.text)}/${bytes(full.text)}B，unchanged ${bytes(unchanged.text)}B`;
    } finally {
        rmSync(root, { recursive: true, force: true });
    }
});

// ── 投影预算（第四十八轮 P0-2） ────────────────────────────────────────────

await check('投影预算：超出预算的条目折叠但留入口，0 = 不限，首行报本次字节', async () => {
    resetProjection();
    const root = mkdtempSync(join(tmpdir(), 'office-projection-budget-'));
    try {
        const memory = createMemory({ root });
        const filler = '这一条用来把正文撑到真实量级，让预算断言立在刻度上而不是感觉上。'.repeat(5);
        for (let index = 0; index < 5; index += 1) {
            await memory.mutate({ action: 'add', target: 'project', content: `预算条目 ${index}：${filler}` });
        }
        // 最后写一条 critical：预算装不下全部时，先被贴出来的必须是它（排序不是「谁新谁上」）
        await memory.mutate({ action: 'add', target: 'project', content: `预算条目 关键：${filler}`, importance: 'critical' });
        const digest = await memory.digest();

        const unlimited = projectDigest(digest, { sessionId: 'p-unlimited', context: 'help' });
        const budgeted = projectDigest(digest, { sessionId: 'p-budget', context: 'help', budget: 1400 });
        assert.ok(bytes(budgeted.text) < bytes(unlimited.text),
            `预算要真的省字节：${bytes(budgeted.text)}B vs ${bytes(unlimited.text)}B`);
        assert.ok(budgeted.text.includes('预算条目 关键'), '预算先保证重要度最高的条目（critical 优先于新旧）');
        assert.ok(budgeted.text.includes('超出本次投影预算'), '折叠行要报「超出预算」的条数');
        assert.ok(budgeted.text.includes("office_memory({ action: 'read', layer: 'hot' })"), '折叠行必须留读取入口');
        assert.ok(/本次投影 \d+ B，其中条目 \d+ B（上限 1400 B）/.test(budgeted.text),
            '首行要报这次投影的字节与条目占用（成本可见）');
        assert.ok(budgeted.text.includes('项目与环境 6 条'), '用量行不受预算影响，照旧给全量');

        // 0 = 不限（设置页与 config 的下限就是它）：行为回到老样子
        const zero = projectDigest(digest, { sessionId: 'p-zero', context: 'help', budget: 0 });
        assert.equal(bytes(zero.text), bytes(unlimited.text), '预算 0 = 不限');
        assert.ok(!zero.text.includes('本次投影'), '不限时不印成本行');
        return `不限额 ${bytes(unlimited.text)}B / 预算 1400B 时 ${bytes(budgeted.text)}B`;
    } finally {
        rmSync(root, { recursive: true, force: true });
    }
});

// ── 收尾 ──────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    if (item.ok) console.log(`PASS  ${item.name}${item.note ? `  → ${item.note}` : ''}`);
    else console.log(`FAIL  ${item.name}  → ${item.error?.message ?? item.error}`);
}
console.log(`\ninjection-budget: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
