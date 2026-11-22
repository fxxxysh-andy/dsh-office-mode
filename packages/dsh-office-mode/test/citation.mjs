/**
 * 引用纪律（第五十七轮 19-8 的 P2-9）的测试。
 *
 * 这条账的原文是「落盘断言强制句级来源引用，评测套 ALCE 三指标（正确性 / 召回 /
 * 精确率）」。ALCE 的三项里只有**两项**能在离线、无模型、零依赖的本项目里算出来，
 * 第三项（正确性）要蕴含判断 —— 所以这个套件守的第一件事不是「算得准」，而是
 * **名实一致**：
 *   1. 判据表（事实句词表、句界与「句号后的出处并回上一句」）—— 判据可被逐条核对；
 *   2. 两项能算的指标算得出、算得对（召回 = 事实句里带引用的比例；精确率 = 正文里的
 *      地址对不对得上本次材料）；
 *   3. 第三项**照实标 null**，且回执里写明「不许拿上面两项冒充」—— 这一条是这条
 *      欠账拖到第五十七轮的原因，也是它最容易在下一轮被悄悄改回去的地方；
 *   4. 落盘断言真的**断言**：`office_parse_findings({ draft })` 在没过时 `ok:false`，
 *      并逐句点名要改哪几句、哪条地址。
 *
 * 跑法：node test/citation.mjs
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { auditCitations, hasCitation, isFactSentence, normalizeCitationUrl, renderCitationAudit, splitSentences } from '../src/citation.js';
import { resolveConfig } from '../src/config.js';
import { parseResultFiles } from '../src/search.js';
import { buildTools } from '../src/tools.js';

const root = mkdtempSync(join(tmpdir(), 'office-tmp-citation-'));
const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const URL_A = 'https://www.example.com/a';
const URL_B = 'https://news.example.org/2024/0701.html';

// ── 1. 判据表：事实句词表 ────────────────────────────────────────────────────

await check('事实句判据：年份 / 百分比 / 带外部量词的数字 / 归因词命中，页码与建议句不命中', () => {
    const yes = [
        '该企业在 2024 年的产能翻了一倍。',
        '同类事故同比上升 15%。',
        '事故造成 12 人受伤。',
        '据应急管理部通报，现场明火已扑灭。',
        '通报显示，现场明火已扑灭。',
        '公司注册资本为 3.2 亿元。',
    ];
    for (const sentence of yes) assert.equal(isFactSentence(sentence), true, `该判成事实句：${sentence}`);
    // 反例里前五条是本轮词表剔掉的误判面（独立复核实测都曾判成事实句）：自指、方法、
    // 计划、修辞 —— 它们被判成事实句的后果是 `ok:false` 逼着改一句正常的话。
    const no = [
        '## 事故简报',
        '本报告共 12 页。',
        '建议后续加强巡检。',
        '这一步要注意先备份。',
        '（来源：https://www.example.com/a）',
        '显示器分辨率建议不低于 1920×1080。',
        '这一步会把产物发布到目标目录。',
        '按 Enter 继续。',
        '按照设置页里的通道顺序，逐条试。',
        '根据判据，会漏也会误。',
        '唯一需要注意的是时间安排。',
        '图 3 显示了整体架构。',
        '我们预计下周完成初稿。',
        '这一节超过一半的篇幅在讲风险。',
    ];
    for (const sentence of no) assert.equal(isFactSentence(sentence), false, `不该判成事实句：${sentence}`);
    return `正例 ${yes.length} / 反例 ${no.length}`;
});

await check('句界：分号不算句界，句号后面的「（来源：…）」并回上一句', () => {
    // 分号是**句内**分句界：`一是…；二是…；三是…。（来源：X）` 是最常见的写法，
    // 按分号切会把前两个分句判成「没出处」（独立复核 P1 的假红）。
    const semi = splitSentences('事故造成 12 人受伤；另有 3 人失联。（来源：https://www.example.com/a）');
    assert.equal(semi.length, 1, JSON.stringify(semi));
    assert.ok(hasCitation(semi[0]), '分号连起来的一整句要算作有出处');

    const sentences = splitSentences([
        '## 标题不当事实句',
        '事故造成 12 人受伤。（来源：https://www.example.com/a）',
        '该企业 2024 年营业收入 3.2 亿元。',
        '（来源：https://news.example.org/2024/0701.html）',
    ].join('\n'));
    assert.equal(sentences.length, 2, JSON.stringify(sentences));
    assert.ok(hasCitation(sentences[0]), '同行的出处要并回上一句');
    assert.ok(hasCitation(sentences[1]), '另起一行的出处也要并回上一句（跨行同理）');
    // 纯引用段并进上一句之后，实际句数不该虚高（从前按句号硬切会多出一句「没出处的事实句」）。
    const only = splitSentences('（来源：https://www.example.com/a）');
    assert.deepEqual(only, ['（来源：https://www.example.com/a）'], '开头就是出处时无处可并，保留原样');
    // 分隔线不是句子、也是合并屏障（`---` 之后通常是附录 / 参考文献，不该并回上一节）。
    assert.deepEqual(
        splitSentences('正文一句。\n\n---\n\n（来源：https://www.example.com/a）'),
        ['正文一句。', '（来源：https://www.example.com/a）'],
    );
    return sentences[0].slice(0, 24) + '…';
});

await check('合并屏障：出处**绝不跨标题**并回上一节（否则上一节白捡一个出处）', () => {
    const anchors = splitSentences([
        '事故造成 12 人受伤。',
        '## 来源',
        '（来源：https://www.example.com/a）',
    ].join('\n'));
    assert.equal(anchors.length, 2, JSON.stringify(anchors));
    assert.equal(hasCitation(anchors[0]), false, '跨标题并回会让没出处的那句变成「有出处」（假绿）');
    assert.ok(hasCitation(anchors[1]), '下一节那条出处自成一格，要保留下来');
    // 逐句各自成篇地审：第一句是「有事实句没出处」，第二句只剩出处（没有事实句）。
    const first = auditCitations({ text: anchors[0], materialUrls: [URL_A], label: 'd.md' });
    assert.equal(first.verdict, 'fix', '结论必须是「有事实句没出处」，不是通过');
    assert.equal(first.uncited.length, 1);
    return '跨标题不合并';
});

await check('代码块与行内代码里的地址是示例，不进分母', () => {
    const text = [
        '用法如下：',
        '```bash',
        'curl https://api.example.com/v1',
        '```',
        '行内的 `https://inline.example.com/x` 也只是举例。',
        '事故造成 12 人受伤。（来源：https://www.example.com/a）',
    ].join('\n');
    const audit = auditCitations({ text, materialUrls: [URL_A], label: 'd.md' });
    assert.equal(audit.unresolved.length, 0, JSON.stringify(audit.unresolved));
    assert.deepEqual(audit.precision, { got: 1, total: 1, rate: 1, site: 0 });
    assert.equal(audit.verdict, 'pass', JSON.stringify(audit.uncited));
    return '两种代码形状都剔掉';
});

// ── 2. 两项算得出的指标 + 一项照实标 null ────────────────────────────────────

await check('审计通过：事实句都有出处、地址都能对得上材料，correctness 恒 null', () => {
    const text = [
        '事故造成 12 人受伤。（来源：https://www.example.com/a）',
        `该企业 2024 年营业收入为 3.2 亿元 (${URL_B})。`,
        '建议后续加强巡检。',
    ].join('\n');
    const audit = auditCitations({ text, materialUrls: [URL_A, URL_B], label: 'docs/简报.md' });
    assert.equal(audit.verdict, 'pass', JSON.stringify(audit.uncited));
    assert.equal(audit.facts, 2);
    assert.deepEqual(audit.recall, { got: 2, total: 2, rate: 1 });
    assert.deepEqual(audit.precision, { got: 2, total: 2, rate: 1, site: 0 });
    assert.equal(audit.correctness, null, 'ALCE 第三项离线算不出，必须是 null');
    assert.ok(audit.notes.some((note) => note.includes('correctness') && note.includes('不许')), '口径要写进 notes');
    return `召回 ${audit.recall.got}/${audit.recall.total}｜精确率 ${audit.precision.got}/${audit.precision.total}`;
});

await check('审计未过：缺出处的事实句与对不上材料的地址分别点名，且带句号', () => {
    const text = [
        '事故造成 12 人受伤。',
        '据应急管理部通报，现场明火已扑灭。',
        `该企业 2024 年营业收入为 3.2 亿元 (https://bogus.example.net/x)。`,
    ].join('\n');
    const audit = auditCitations({ text, materialUrls: [URL_A], label: 'docs/简报.md' });
    assert.equal(audit.verdict, 'fix');
    assert.equal(audit.uncited.length, 2, '两句没有出处');
    assert.deepEqual(audit.uncited.map((item) => item.index), [0, 1], '每条要带句号（否则 >8 条时没法定位）');
    assert.equal(audit.unresolved.length, 1, '一条地址材料里没有');
    assert.equal(audit.unresolved[0].url, 'https://bogus.example.net/x');
    assert.deepEqual(audit.recall, { got: 1, total: 3, rate: 1 / 3 });
    assert.deepEqual(audit.precision, { got: 0, total: 1, rate: 0, site: 0 });
    const rendered = renderCitationAudit(audit);
    assert.ok(rendered.includes('未过'), rendered);
    assert.ok(rendered.includes('2 句声称外部事实却没给出处'), rendered);
    assert.ok(rendered.includes('第 1 句'), '要点到具体句号');
    assert.ok(rendered.includes('1 条地址对不上本次材料'), rendered);
    assert.ok(rendered.includes('correctness：算不出'), rendered);
    return `未过：缺出处 ${audit.uncited.length} / 地址对不上 ${audit.unresolved.length}`;
});

await check('地址对账：同站不同页只算「站点级」，同一路径不同参数算「对不上」', () => {
    const text = `事故造成 12 人受伤 (https://www.example.com/other-page)。`;
    const audit = auditCitations({ text, materialUrls: [URL_A], label: 'd.md' });
    assert.equal(audit.verdict, 'pass', '站点级对上不算「对不上」');
    assert.equal(audit.siteOnly.length, 1);
    // 分子**只算逐字对上的那批**：把 site 档算进分子，等于把唯一那条硬判据悄悄放松。
    assert.deepEqual(audit.precision, { got: 0, total: 1, rate: 0, site: 1 });
    assert.ok(audit.notes.some((note) => note.includes('只对到**站点**')), '要写明是同站不同页');
    const rendered = renderCitationAudit(audit);
    assert.ok(rendered.includes('另有 1 条只对到站点'), rendered);

    // `?id=1` 与 `?id=2` 常是两个页面：归一时**保留 query**，这一条要算对不上。
    // 归一的口径本身也钉一下：去协议 / www / fragment / 尾斜杠、小写，**query 留着**。
    assert.equal(normalizeCitationUrl('https://www.Example.com/a/?id=1#frag'), 'example.com/a?id=1');
    const query = auditCitations({
        text: '事故造成 12 人受伤（来源：https://news.example.com/article?id=2）。',
        materialUrls: ['https://news.example.com/article?id=1'],
        label: 'd.md',
    });
    assert.equal(query.verdict, 'fix');
    assert.equal(query.unresolved.length, 1);
    assert.equal(query.unresolved[0].status, 'query');
    return '站点级 1 条 / 参数不同 1 条';
});

await check('不算进分母的东西：kb 块 id 只核形状，出处标记与编号引用只算标了出处', () => {
    const text = [
        '统计显示，该结论来自一块来源 kb:0123456789abcdef:3。',
        '另有编号引用[1][2]、脚注[^3]与锚点式[4](#ref)，以及出处标记（出处：某通报）在此。',
    ].join('\n');
    const audit = auditCitations({ text, materialUrls: [], label: 'd.md' });
    assert.deepEqual(audit.precision, { got: 0, total: 0, rate: null, site: 0 }, '没有可对账的地址时分母为 0、比率标 null');
    assert.ok(audit.notes.some((note) => note.includes('kb 块 id 只核了形状')), audit.notes.join(' | '));
    // 四种编号写法都要认得出来（`[1][2]` 曾只数到 1 处）。
    assert.ok(audit.notes.some((note) => note.includes('4 处编号引用')), audit.notes.join(' | '));
    assert.ok(audit.notes.some((note) => note.includes('出处标记') || note.includes('（来源：…）')), audit.notes.join(' | '));
    // 标了出处**不算失败**（对账不了的东西不该造成假红），但也不是「通过」：
    // 这一次确实什么都没验证到 —— 回执要这么写。
    assert.equal(audit.verdict, 'empty', '没有可对账的地址 = 没验证到东西，但不判失败');
    assert.equal(audit.recall.rate, 1, '句级引用率仍然算得出来（假都不算，就没法说引用齐不齐）');
    // 形状坏的 kb id 认不出来 —— 不该被当成一条合法引用。
    assert.equal(hasCitation('统计显示，来源 kb:xyz:1。'), false);
    return '分母只算地址';
});

await check('空判不当通过：没有事实句或没有地址时，verdict 是 empty 而不是 pass', () => {
    const empty = auditCitations({ text: '', materialUrls: [] });
    assert.equal(empty.verdict, 'empty');
    assert.equal(empty.facts, 0);
    assert.equal(empty.recall.rate, null);
    assert.ok(empty.notes.some((note) => note.includes('没有可判的句子')));
    const prose = auditCitations({ text: '建议后续加强巡检。\n先把备份做完。', materialUrls: [] });
    assert.equal(prose.facts, 0);
    assert.equal(prose.verdict, 'empty', '没判出事实句不是「引用都齐了」');
    assert.ok(prose.notes.some((note) => note.includes('没有判出任何事实句') && note.includes('这不是')), prose.notes.join(' | '));
    assert.ok(renderCitationAudit(prose).includes('没判出可核的东西'), renderCitationAudit(prose));
    // 整份都是引用块 / 表格的正文也一样（那两种行整行不判，是**有意的**边界）。
    const table = auditCitations({ text: '| 伤亡 | 2024 年共 12 人 |\n| --- | --- |', materialUrls: [] });
    assert.equal(table.verdict, 'empty');
    return '空与全散文两种形状都照实说';
});

// ── 3. 落盘断言：office_parse_findings 的 draft 通道 ─────────────────────────

/** 一份材料文件：一条结论挂一个地址（对账用的地址集合就来自这里）。 */
function writeMaterial(name, url) {
    const file = join(root, name);
    writeFileSync(file, ['## 泛搜', `- 事故造成 12 人受伤 (${url})`].join('\n'), 'utf8');
    return file;
}

/** 走一次真的工具面（与模型调用同一条路）。 */
async function parseViaTool(args) {
    const tool = buildTools(resolveConfig({})).find((item) => item.name === 'office_parse_findings');
    return tool.execute(args, exec);
}

await check('工具面：没给 draft 时行为逐字不变（ok 仍为 true、不给引用审计）', async () => {
    const material = writeMaterial('m1.md', URL_A);
    const out = await parseViaTool({ paths: [material], topic: '某事件' });
    assert.equal(out.ok, true);
    assert.ok(!out.text.includes('引用审计'), out.text);
    return '旧行为不变';
});

await check('工具面：给了 draft 且引用没过 → ok:false，并逐句点名', async () => {
    const material = writeMaterial('m2.md', URL_A);
    const draft = join(root, 'draft-bad.md');
    writeFileSync(draft, [
        '# 简报',
        '事故造成 12 人受伤。',
        '据通报，现场明火已扑灭。（来源：https://bogus.example.net/x）',
    ].join('\n'), 'utf8');
    const out = await parseViaTool({ paths: [material], draft, topic: '某事件' });
    assert.equal(out.ok, false, '落盘断言没过就要标错，不能只是一个提示');
    assert.ok(out.text.includes('引用审计'), out.text);
    assert.ok(out.text.includes('未过'), out.text);
    assert.ok(out.text.includes('bogus.example.net'), '对不上的地址要点名');
    assert.ok(out.text.includes('事故造成 12 人受伤。'), '缺出处的事实句要点名');
    assert.ok(out.text.includes('不要当成定稿交出去'), '要说清下一步');
    return 'ok:false + 点名';
});

await check('工具面：补上出处之后同一次调用变绿（同一份材料与同一份正文）', async () => {
    const material = writeMaterial('m3.md', URL_A);
    const draft = join(root, 'draft-good.md');
    writeFileSync(draft, [
        '# 简报',
        `事故造成 12 人受伤。（来源：${URL_A}）`,
        '建议后续加强巡检。',
    ].join('\n'), 'utf8');
    const out = await parseViaTool({ paths: [material], draft, topic: '某事件' });
    assert.equal(out.ok, true, out.text);
    assert.ok(out.text.includes('引用审计'), out.text);
    assert.ok(out.text.includes('通过'), out.text);
    return '改完即通过';
});

await check('工具面：draft 收到 Office 文档时给可执行的错（并指路转换 / 抽文本）', async () => {
    const material = writeMaterial('m4.md', URL_A);
    const docx = join(root, '正文.docx');
    writeFileSync(docx, 'not really a docx', 'utf8');
    const refused = await parseViaTool({ paths: [material], draft: docx, topic: '某事件' }).then(() => null, (error) => error);
    assert.match(refused.message, /Office \/ PDF \/ ODF/);
    assert.match(refused.message, /office\.convert\.file/, '要给出真的能走的那条路');
    await assert.rejects(() => parseViaTool({ paths: [material], draft: join(root, 'nope.md') }), /读不到/);
    // 参数形状坏掉 / 空串也要当场拒绝：静默当「没给」会让调用方以为审过了。
    await assert.rejects(() => parseViaTool({ paths: [material], draft: 123 }), /需要一个字符串/);
    await assert.rejects(() => parseViaTool({ paths: [material], draft: '   ' }), /是空的/);
    return '四种坏入参都当场说清';
});

await check('工具面：材料地址写在**散文行**里也算对得上（只认列表项会造成假红）', async () => {
    const file = join(root, 'm5.md');
    writeFileSync(file, `## 泛搜\n另见 ${URL_A} 的通报，还有一段散文说明。\n`, 'utf8');
    const draft = join(root, 'draft-prose.md');
    writeFileSync(draft, `# 简报\n事故造成 12 人受伤。（来源：${URL_A}）\n`, 'utf8');
    const out = await parseViaTool({ paths: [file], draft, topic: '某事件' });
    assert.equal(out.ok, true, out.text);
    assert.ok(out.text.includes('引用审计通过'), out.text);
    return '散文行的地址进对账集合';
});

await check('工具面：没判出可核的东西时不印「通过」（空判文案要自相不打脸）', async () => {
    const material = writeMaterial('m6.md', URL_A);
    const draft = join(root, 'draft-prose-only.md');
    writeFileSync(draft, '# 简报\n建议后续加强巡检。\n先把备份做完。\n', 'utf8');
    const out = await parseViaTool({ paths: [material], draft, topic: '某事件' });
    assert.equal(out.ok, true, 'empty 不是失败：它照实说「没验证到什么」');
    assert.ok(out.text.includes('没判出可核的东西'), out.text);
    assert.ok(!out.text.includes('引用审计通过'), '不许把空判说成通过');
    return 'empty 有独立文案';
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`citation: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
