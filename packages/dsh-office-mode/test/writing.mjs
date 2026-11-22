#!/usr/bin/env node
/**
 * 写作规划与自检的回归测试（第七十三轮）。
 *
 * 钉住四件事：
 *   1. 规划按难度 / 读者收益分配字数，目标总字数精确加总，单节有上限；
 *   2. 文体、节型别名与逐节字段都能落到 plan，markdown 可读；
 *   3. 自检的确定性信号与规划对表真的会报（套话、句长、重点节、平均用力、最长节给错）；
 *   4. 去 AI 味报告只报 A1/A2/M1 与文体偏离，不给分；safe 只删提示语；
 *   5. SDK 的 office.writing 五个入口真的挂上了。
 *
 * 跑法：node test/writing.mjs（加 --keep 保留 .office/runtime/tmp/writing-test/）。
 * 本套件不联网，进 CI。
 */
import { mkdirSync, rmSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { writingPlan, writingAudit, writingDeai, writingSignals, renderWritingPlan } from '../src/writing.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..', '..', '..');
const TMP = join(ROOT, '.office', 'runtime', 'tmp', 'writing-test');
const KEEP = process.argv.includes('--keep');

const failures = [];
let checks = 0;

function check(condition, label, detail = '') {
    checks += 1;
    if (!condition) failures.push(detail ? label + ' —— ' + detail : label);
}
function checkEqual(actual, expected, label) {
    check(actual === expected, label, '期望 ' + JSON.stringify(expected) + '，实际 ' + JSON.stringify(actual));
}
function checkIncludes(haystack, needle, label) {
    const hit = Array.isArray(haystack) ? haystack.includes(needle) : String(haystack).includes(needle);
    check(hit, label, '没有找到 ' + JSON.stringify(needle));
}

function sumTargets(plan) {
    return plan.sections.reduce((sum, section) => sum + section.targetChars, 0);
}

function repeat(sentence, times) {
    return Array.from({ length: times }, () => sentence).join('');
}

// ── 1. 规划：分配、深度、文体与别名 ─────────────────────────────────────────

{
    const plan = writingPlan({
        title: '检索噪声为什么会拖垮答案',
        genre: 'academic',
        audience: '做过 RAG 的工程师',
        goal: '让读者相信扩库后的正确率下降主要来自噪声',
        targetChars: 1200,
        sections: [
            { heading: '反差的来处', kind: 'context', difficulty: 1, value: 1.5 },
            { heading: '两条机制假设', kind: 'theory', difficulty: 2.8, value: 3 },
            { heading: '可证伪的预测与实验', kind: 'evidence', difficulty: 2.4, value: 3 },
            { heading: '边界', kind: 'limitation', difficulty: 1.8, value: 2 },
        ],
    });
    checkEqual(plan.totalChars, 1200, '目标总字数保留');
    checkEqual(sumTargets(plan), 1200, '各节目标字数精确加总');
    checkEqual(plan.brief.genre, 'academic', '文体识别');
    checkEqual(plan.sections.length, 4, '节数保留');
    checkEqual(plan.sections[1].kind, 'theory', '显式 kind 保留');
    checkIncludes(plan.markdown, '篇幅（按难度与读者收益分，不平均）', 'markdown 有篇幅表');
    checkIncludes(plan.markdown, '三问①', 'markdown 有写前三问');
    checkIncludes(plan.markdown, '每 150 字', 'markdown 有锚点密度要求');
    checkIncludes(plan.markdown, '写透', 'markdown 有深度标签');
    check(plan.sections.every((section) => ['compress', 'touch', 'develop', 'dwell'].includes(section.depth)), 'depth 取值合法');
    const dwell = plan.sections.filter((section) => section.depth === 'dwell');
    check(dwell.length >= 1, '至少一节写透');
    check(plan.sections[1].targetChars > plan.sections[0].targetChars * 2, '难点节比背景节多笔墨');
    checkIncludes(plan.checks.join('\n'), '三问③', 'checks 有最长节检查');
    checkIncludes(plan.checks.join('\n'), '最长的一节', '权重差异会触发最长节提醒');
}

{
    const plan = writingPlan({
        title: '给普通读者解释幻觉',
        genre: '科普',
        targetChars: 900,
        sections: [
            { heading: '一个不存在的生平' },
            { heading: '机制是什么' },
            { heading: '它能被怎么发现' },
        ],
    });
    checkEqual(plan.brief.genre, 'explainer', '中文文体别名');
    checkEqual(plan.sections[0].kind, 'example', '节名猜类型 example');
    checkEqual(plan.sections[1].kind, 'theory', '节名猜类型 theory');
    checkEqual(plan.sections[1].register, 'plain', '说明文的机制节用平实语言');
    checkEqual(plan.sections[1].figurative, 'everyday-metaphor', '说明文允许日常比喻');
    checkEqual(sumTargets(plan), 900, '默认节名也精确加总');
}

{
    const plan = writingPlan({ title: '平均用力', genre: 'general', targetChars: 1000, sections: ['A', 'B', 'C', 'D'] });
    const maxTarget = Math.max(...plan.sections.map((section) => section.targetChars));
    check(maxTarget <= 400, '四节时单节不超过 40%', String(maxTarget));
    checkEqual(sumTargets(plan), 1000, '四节目标字数加总');
    const shares = plan.sections.map((section) => section.share);
    check(Math.max(...shares) / Math.min(...shares) < 1.3, '四节默认权重接近时份额接近');
}

{
    const short = writingPlan({ title: '短', targetChars: 100 });
    checkEqual(short.totalChars, 300, '总字数下限夹到 300');
    const long = writingPlan({ title: '长', targetChars: 999999 });
    checkEqual(long.totalChars, 20000, '总字数上限夹到 20000');
}

// ── 2. 渲染与边界 ───────────────────────────────────────────────────────────

{
    const plan = writingPlan({ title: '渲染', genre: 'report', targetChars: 800 });
    checkEqual(renderWritingPlan(plan), plan.markdown, 'render(plan) 与 markdown 一致');
    checkIncludes(plan.markdown, '## 1.', '逐节编号从 1 开始');
    checkIncludes(plan.markdown, '读者问题', '逐节给读者问题');
    checkIncludes(plan.markdown, '必须写到', '逐节给必须写到');
    checkIncludes(plan.markdown, '不要', '逐节给反模式');
    let threw = false;
    try {
        renderWritingPlan({ nope: true });
    } catch (error) {
        threw = /写作规划|writingPlan/.test(String(error.message));
    }
    check(threw, 'render 对非法对象报错并指路');
}

// ── 3. 自检：套话、句长、规划对表 ───────────────────────────────────────────

{
    const text = repeat('首先，这个方案非常重要。', 8) + '其次，它可能有效。最后，我们要重视。';
    const report = writingAudit(text);
    checkIncludes(report.warnings.map((item) => item.code), 'cliches', '套话会被报出');
    checkIncludes(report.warnings.map((item) => item.code), 'uniform-sentences', '句长过匀会被报出');
    check(report.score < 80, '套话 + 句长过匀会扣分', String(report.score));
    check(report.suggestions.length > 0, '有改稿建议');
    checkIncludes(report.note, '不是「像人程度」的证明', 'note 说清 score 的边界');
}

{
    const plan = writingPlan({
        title: '重点节',
        genre: 'academic',
        targetChars: 1200,
        sections: [
            { heading: '背景', kind: 'context', difficulty: 1, value: 1 },
            { heading: '机制', kind: 'theory', difficulty: 3, value: 3 },
            { heading: '边界', kind: 'limitation', difficulty: 2, value: 2 },
        ],
    });
    const text = [
        '# 背景',
        repeat('这是一段背景，读者已经知道。', 4),
        '',
        '# 机制',
        '这里只有一句机制。',
        '',
        '# 边界',
        repeat('这里有一个边界，条件是明确的。', 3),
    ].join('\n');
    const report = writingAudit(text, plan);
    checkIncludes(report.warnings.map((item) => item.code), 'underdeveloped', '重点节没写透会被报出');
    check(report.sections.some((section) => section.planned !== null), '对表后 sections 带 planned');
    check(report.sections.some((section) => Number.isFinite(section.targetChars)), '对表后带 targetChars');
}

{
    const plan = writingPlan({
        title: '平均分配',
        genre: 'general',
        targetChars: 1600,
        sections: [
            { heading: '背景', kind: 'context', difficulty: 0.5, value: 1 },
            { heading: '机制', kind: 'theory', difficulty: 3, value: 3 },
            { heading: '证据', kind: 'evidence', difficulty: 2.5, value: 3 },
            { heading: '边界', kind: 'limitation', difficulty: 1.5, value: 2 },
        ],
    });
    const body = repeat('这一段写了一些内容，句子长度都差不多。', 5);
    const text = ['# 背景', body, '', '# 机制', body, '', '# 证据', body, '', '# 边界', body].join('\n');
    const report = writingAudit(text, plan);
    const codes = report.warnings.map((item) => item.code);
    checkIncludes(codes, 'uniform-allocation', '各节实际篇幅平均会被报出');
}

{
    const manualPlan = {
        brief: { title: '手写计划', genre: 'academic' },
        sections: [
            { heading: '背景', kind: 'context', targetChars: 320, share: 0.4, depth: 'dwell', depthLabel: '写透', difficulty: 0.5, value: 1 },
            { heading: '机制', kind: 'theory', targetChars: 260, share: 0.325, depth: 'dwell', depthLabel: '写透', difficulty: 3, value: 3 },
            { heading: '边界', kind: 'limitation', targetChars: 220, share: 0.275, depth: 'dwell', depthLabel: '写透', difficulty: 1.8, value: 2 },
        ],
    };
    const text = [
        '# 背景',
        repeat('背景写得非常长，读者已经知道，但这里还是在铺垫。', 12),
        '',
        '# 机制',
        repeat('机制只写了一点点。', 2),
        '',
        '# 边界',
        repeat('边界写了一点点。', 2),
    ].join('\n');
    const report = writingAudit(text, manualPlan);
    checkIncludes(report.warnings.map((item) => item.code), 'longest-not-hardest', '最长节不是最难节会被报出');
}

{
    const text = '# 一段\n' + repeat('这个机制的复杂性体现在系统性的逻辑层面，它是一种范式的变化。', 20);
    const report = writingAudit(text);
    const codes = report.warnings.map((item) => item.code);
    checkIncludes(codes, 'anchor-sparse', '锚点不足会被报出');
    checkIncludes(codes, 'no-turn', '通篇无让步会被报出');
    check(report.signals.abstractHits > 0, 'signals 统计抽象词');
    check(report.signals.chars >= 300, 'signals 统计字数');
}

{
    const signals = writingSignals('第一句。第二句稍微长一点点。第三句。');
    checkEqual(signals.sentences, 3, 'signals 句数');
    checkEqual(signals.shortRatio, 1, 'signals 短句比');
    checkEqual(signals.clicheTotal, 0, 'signals 无套话');
    let threw = false;
    try {
        writingAudit('   ');
    } catch (error) {
        threw = /没有正文|正文可查/.test(String(error.message));
    }
    check(threw, '空正文会抛错');
}

// ── 4. 去 AI 味报告 ─────────────────────────────────────────────────────────

{
    const text = '先说结论：这不是效率问题，而是协作问题。Q3 慢了两周，原因不复杂：财务的分摊口径还没定。';
    const report = writingDeai(text);
    checkEqual(report.mode, 'report', 'deai 默认 report');
    checkEqual(report.score, undefined, 'deai 不返回分数');
    checkEqual(report.register.id, 'general', 'deai 默认 register=general');
    const codes = report.flags.map((flag) => flag.code);
    checkIncludes(codes, 'A1', 'deai 报对举');
    checkIncludes(codes, 'A2', 'deai 报提示语冒号');
    checkIncludes(codes, 'M1', 'deai 报元话语起手式');
    const a1 = report.flags.find((flag) => flag.code === 'A1');
    check(a1.spans.length > 0 && a1.spans[0].text.includes('不是'), 'A1 带 span');
    checkEqual(a1.confidence, 'high', '通用体的 A1 是高置信');
}

{
    const report = writingDeai('先说结论：原因不复杂：预算还没批。', { register: 'official' });
    checkEqual(report.register.id, 'official', 'register 认 official');
    check(!report.flags.some((flag) => flag.code === 'A2'), 'official 把 A2 降为提示');
    checkIncludes(report.notes.join('\n'), '提示语冒号', 'official 的 A2 进 notes');
    checkIncludes(report.flags.map((flag) => flag.code), 'M1', 'official 仍报元话语');
    const chinese = writingDeai('原因不复杂：预算还没批。', { register: '公文' });
    checkEqual(chinese.register.id, 'official', 'register 认中文别名');
}

{
    const safe = writingDeai('先说结论：海拔三千米水开得早。Q3 慢两周，原因不复杂：台账对不上。不是A，是B。', { mode: 'safe' });
    checkEqual(safe.mode, 'safe', 'deai safe 模式');
    checkIncludes(safe.text, '海拔三千米水开得早', 'safe 保留正文');
    check(!safe.text.includes('先说结论：'), 'safe 删元话语前缀');
    check(safe.text.includes('原因是'), 'safe 把原因不复杂改成原因是');
    check(safe.text.includes('不是A，是B'), 'safe 不动对举');
    check(safe.edits.length >= 2, 'safe 返回 edits', String(safe.edits.length));
}

{
    const ref = '上周去了趟成都。早上吃面，中午吃面，晚上还是面。不是没别的，是懒得找。店老板说，这条街的面馆比便利店多。我信了，因为走两步就一家。';
    const report = writingDeai('这是一段很长的正文，句子长度明显超过参考稿的平均句长，而且连续写了很多内容。', { styleRef: ref });
    checkEqual(report.styleRef.provided, true, 'styleRef 会被比较');
    check(report.styleRef.metrics.some((metric) => metric.key === 'meanSentence' && metric.direction === 'higher'), 'styleRef 报句长偏离');
    const weak = writingDeai('短句。', { styleRef: '更短。' });
    checkEqual(weak.styleRef.weak, true, '短参考稿只作方向');
}

{
    let threw = false;
    try {
        writingDeai('   ');
    } catch (error) {
        threw = /没有正文|正文可查/.test(String(error.message));
    }
    check(threw, 'deai 空正文会抛错');
}

// ── 5. SDK 挂载面 ───────────────────────────────────────────────────────────

{
    const { buildSdk } = await import('../src/sdk.js');
    const { testEnv } = await import('../src/sdk.js');
    const { createCache } = await import('../src/engine/cache.js');
    rmSync(TMP, { recursive: true, force: true });
    mkdirSync(TMP, { recursive: true });
    const env = testEnv(TMP);
    const cache = createCache({ root: TMP, dir: 'cache' });
    const sdk = await buildSdk({ env, cache, config: { defaultTheme: 'plain' } });
    checkEqual(typeof sdk.office.writing?.plan, 'function', 'SDK 挂上 office.writing.plan');
    checkEqual(typeof sdk.office.writing?.audit, 'function', 'SDK 挂上 office.writing.audit');
    checkEqual(typeof sdk.office.writing?.deai, 'function', 'SDK 挂上 office.writing.deai');
    checkEqual(typeof sdk.office.writing?.signals, 'function', 'SDK 挂上 office.writing.signals');
    checkEqual(typeof sdk.office.writing?.render, 'function', 'SDK 挂上 office.writing.render');
    const plan = sdk.office.writing.plan({ title: 'SDK', genre: 'explainer', targetChars: 600 });
    checkIncludes(plan.markdown, '写作规划：SDK', 'SDK plan 返回 markdown');
    const report = sdk.office.writing.audit(plan.markdown);
    checkEqual(report.ok, true, 'SDK audit 返回 ok');
    check(Array.isArray(report.warnings), 'SDK audit 返回 warnings');
    const deai = sdk.office.writing.deai('先说结论：测试。');
    checkEqual(deai.mode, 'report', 'SDK deai 返回 report');
    if (!KEEP) rmSync(TMP, { recursive: true, force: true });
}

// ── 汇总 ───────────────────────────────────────────────────────────────────

if (failures.length > 0) {
    console.error('\n失败：');
    for (const failure of failures) console.error('  ✗ ' + failure);
}
console.log('\n' + (checks - failures.length) + ' 通过 / ' + failures.length + ' 失败（共 ' + checks + ' 项）');
process.exit(failures.length === 0 ? 0 : 1);
