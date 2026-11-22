#!/usr/bin/env node
/**
 * 大模型写作 A/B 评测（第七十三轮）。
 *
 * 为什么要有它：writing.js 给的是「怎么分配注意力」的规划，但规划有没有让模型写得更好，
 * 只能拿同一批任务在「普通提示」与「按规划写」两个臂下对照。这个脚本把对照做成可复跑的一条命令：
 * 任务来自 .office/runtime/diag/round73-writing-research.md §4（A/B/C 三题），盲评口径来自 §5
 * （三维 1–5 锚点 + 10 条反模板清单），并且刻意控住两个已知偏差：
 *   · 位置偏差：每任务第 1 轮 baseline→planned，第 2 轮 planned→baseline（换序各测一轮）；
 *   · 长度偏差：目标字数 ±10% 之外标 truncated，报告里单列长度，不拿长短当质量。
 *
 * 三条边界：
 *   1. 无 Key 就 SKIP（exit 0），**不写任何假结果**；--mock 是本地确定性假模型，只验证流程
 *      与产物结构（自测 / CI 用），报告与 key.json 都会标明 mock，不许当成模型表现。
 *   2. samples/*.md 只含编号与正文；编号→任务/臂/轮次/模型在 key.json —— 盲评者拿不到臂名。
 *   3. 零第三方依赖，Node 22+；除 --mock / --self-test 外只访问 --base-url 的
 *      /chat/completions（OpenAI 兼容）。
 *
 * 跑法：
 *   node scripts/writing-eval.mjs --tasks a --mock --out .office/runtime/tmp/writing-eval-demo
 *   node scripts/writing-eval.mjs --self-test
 *   DEEPSEEK_API_KEY=... node scripts/writing-eval.mjs --tasks a,b,c --runs 2 --out .office/runtime/tmp/writing-eval
 *   DEEPSEEK_API_KEY=... node scripts/writing-eval.mjs --judge deepseek-chat --out .office/runtime/tmp/writing-eval-judged
 *
 * 退出码：0 成功（含 SKIP）；1 生成/写盘失败；2 参数错误。
 */
import { spawn } from 'node:child_process';
import {
    existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { writingAudit, writingPlan, writingSignals } from '../src/writing.js';

const SELF = fileURLToPath(import.meta.url);

const DEFAULT_BASE_URL = 'https://api.deepseek.com';
const DEFAULT_MODEL = 'deepseek-chat';
const DEFAULT_OUT_DIR = '.office/runtime/tmp/writing-eval';
const DEFAULT_RUNS = 2;
const DEFAULT_TEMPERATURE = 0.8;
const TOLERANCE = 0.1;
const ARMS = ['baseline', 'planned'];

// ---------------------------------------------------------------- 评测任务（§4）

/**
 * 三个任务直接采用研究材料 §4 的题干与详略要求。sections 只写「篇幅与动作」，
 * 具体必须写到什么由 writing.js 的 kind 档案补齐 —— 评测比的是规划带来的差异，
 * 不是让评测脚本自己另写一套写作规则。
 */
export const WRITING_TASKS = [
    {
        id: 'a',
        title: '检索扩到 1 亿文档后，召回率升了、答案正确率却降了',
        genre: 'academic',
        genreLabel: '学术论文',
        audience: '做过 RAG 但不熟悉检索噪声的工程师',
        goal: '让读者能判断两条机制假设哪条更可能，并知道怎么用小实验区分。',
        targetChars: 550,
        mockBaselineHeadings: ['现象', '解释', '收尾'],
        sections: [
            { heading: '反差与背景', kind: 'context', difficulty: 0.8, value: 1 },
            { heading: '两条机制假设', kind: 'theory', difficulty: 2.6, value: 3 },
            { heading: '哪条更可能：可证伪的预测', kind: 'leap', difficulty: 2.4, value: 2.8 },
            { heading: '区分两条假设的小实验', kind: 'method', difficulty: 2.2, value: 2.6 },
            { heading: '适用边界', kind: 'limitation', difficulty: 1.6, value: 1.8 },
        ],
        brief: '某团队把检索增强问答系统从 10 万文档扩到 1 亿文档：召回率 0.71→0.86，但端到端答案正确率 0.64→0.57。请写一节解释这个反差：给出至少两条机制假设并说明哪条更可能；给一条可证伪的预测；设计一个能区分两条假设的小实验（写清变量与判据）；说明结论的适用边界。读者是做过 RAG 但不熟悉检索噪声的工程师。不要写文献综述，不要罗列所有可能原因。',
    },
    {
        id: 'b',
        title: '为什么聊天机器人会一本正经地编造事实',
        genre: 'explainer',
        genreLabel: '说明文',
        audience: '完全没接触过大模型的普通读者',
        goal: '让读者明白幻觉为什么发生，并拿到一个可用的自查方法。',
        targetChars: 475,
        mockBaselineHeadings: ['开篇', '展开', '收尾'],
        sections: [
            { heading: '一个具体场景', kind: 'example', difficulty: 1, value: 1.6 },
            { heading: '幻觉是怎么发生的', kind: 'theory', difficulty: 2.4, value: 3 },
            { heading: '主比喻与它失效的地方', kind: 'example', difficulty: 1.6, value: 2.2 },
            { heading: '你可以怎么自查', kind: 'howto', difficulty: 1.6, value: 2.2 },
            { heading: '边界与常见误解', kind: 'limitation', difficulty: 1.4, value: 1.6 },
        ],
        brief: '向完全没接触过大模型的读者解释：为什么聊天机器人会一本正经地编造事实（幻觉）。可以用 1–2 个日常比喻，但每个比喻之后必须说明它在哪里不成立；至少给一个具体场景（如询问一个不存在的人物的生平）和一个读者可用的自查方法。不要讲 Transformer 结构，不要堆术语，不要用「简单来说」重复上一句。',
    },
    {
        id: 'c',
        title: '首页改版后新用户次日留存下滑的季度汇报',
        genre: 'report',
        genreLabel: '工作汇报',
        audience: '管理层（没有时间看细节，只做决策）',
        goal: '让管理层在 5 分钟内决定是否回滚首页改版并继续观察。',
        targetChars: 350,
        mockBaselineHeadings: ['情况', '分析', '建议'],
        sections: [
            { heading: '结论', kind: 'summary', difficulty: 1, value: 3 },
            { heading: '依据（按证据强度排序）', kind: 'evidence', difficulty: 1.6, value: 3 },
            { heading: '风险（如果…则…）', kind: 'limitation', difficulty: 1.6, value: 2.4 },
            { heading: '决策请求', kind: 'howto', difficulty: 1.2, value: 2.6 },
        ],
        brief: '给管理层写一页季度汇报。事实：新用户次日留存 42%→38%（同期上线首页改版，改版组 n=12,400、对照组 n=11,900）；老用户留存持平；客服工单 +18%，其中「找不到功能」占 46%。要求：第一句给结论；三条依据按证据强度排序并标口径；两条风险写成「如果…则…」；最后给一个明确的决策请求。不要复述背景，不要写感谢。',
    },
];

// ---------------------------------------------------------------- 基础工具

class EvalParamError extends Error {
    constructor(message) {
        super(message);
        this.name = 'EvalParamError';
        this.code = 'param';
    }
}

function countChars(value) {
    return Array.from(String(value ?? '').replace(/\s+/gu, '')).length;
}

function mean(values) {
    if (values.length === 0) return 0;
    return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function formatNumber(value, digits = 1) {
    return Number.isFinite(value) ? value.toFixed(digits) : '—';
}

function hashString(text) {
    let hash = 2166136261;
    for (const char of String(text)) {
        hash ^= char.codePointAt(0);
        hash = Math.imul(hash, 16777619);
    }
    return hash >>> 0;
}

/** 确定性 PRNG（不碰 Math.random，保证 --mock 与产物可复现）。 */
function mulberry32(seed) {
    let state = seed >>> 0;
    return () => {
        state = (state + 0x6d2b79f5) >>> 0;
        let value = Math.imul(state ^ (state >>> 15), 1 | state);
        value = (value + Math.imul(value ^ (value >>> 7), 61 | value)) ^ value;
        return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
    };
}

function shuffle(list, random) {
    const out = list.slice();
    for (let index = out.length - 1; index > 0; index -= 1) {
        const swapAt = Math.floor(random() * (index + 1));
        [out[index], out[swapAt]] = [out[swapAt], out[index]];
    }
    return out;
}

/** 去掉 Markdown 标题后按空行切段（只用于反模板清单，不替代 writing.js 的自检）。 */
function paragraphsOf(text) {
    return String(text ?? '')
        .replace(/^#{1,6}\s+.*$/gmu, '')
        .replace(/\x60\x60\x60[\s\S]*?\x60\x60\x60/gu, '')
        .split(/\n\s*\n/gu)
        .map((paragraph) => paragraph.replace(/\n/gu, ' ').trim())
        .filter((paragraph) => countChars(paragraph) >= 2);
}

function ngramSet(text, size) {
    const chars = Array.from(String(text).replace(/\s+/gu, ''));
    const grams = new Set();
    for (let index = 0; index + size <= chars.length; index += 1) {
        grams.add(chars.slice(index, index + size).join(''));
    }
    return grams;
}

// ---------------------------------------------------------------- §5 反模板清单（0/1）

const ANTI_TEMPLATE_ITEMS = [
    '以「首先/其次/最后」为唯一结构',
    '每节结尾有「综上所述/总之」',
    '连续 3 句字数差 <5',
    '黑名单套话 ≥2 次/千字',
    '三段并列排比无新增信息',
    '每个论点后必接「这不仅…而且…」',
    '比喻只出现不回收、不设边界',
    '结论段复述前文原句',
    '全文无一处让步或反例',
    '数字无口径/无基线/无来源',
];

const NUMBER_QUALIFIER_RE = /n\s*=|N\s*=|样本|口径|对照|基线|来源|数据|同比|环比|置信|区间|±|标准差|均值/u;
const METAPHOR_BOUNDARY_RE = /不成立|不像|不适用|边界|失效|并非|并不是|只在|例外/u;

function findUniformTriple(lengths) {
    for (let index = 0; index + 2 < lengths.length; index += 1) {
        const triple = lengths.slice(index, index + 3);
        if (Math.min(...triple) < 8) continue;
        if (Math.max(...triple) - Math.min(...triple) < 5) return triple;
    }
    return null;
}

function hasParallelParagraphs(paragraphs) {
    if (paragraphs.length < 3) return false;
    for (let index = 0; index + 2 < paragraphs.length; index += 1) {
        const triple = paragraphs.slice(index, index + 3);
        const lengths = triple.map(countChars);
        if (Math.min(...lengths) < 30) continue;
        if (Math.max(...lengths) / Math.min(...lengths) > 1.15) continue;
        const bare = triple.every((paragraph) => !/\d/u.test(paragraph)
            && !/例如|比如|举个例子|具体来说|有一次|设想|假设你/u.test(paragraph)
            && !/但|不过|然而|反例|问题在于|恰恰|局限|失效|不成立|例外|代价|反对/u.test(paragraph));
        if (bare) return true;
    }
    return false;
}

function repeatsLastParagraph(paragraphs) {
    if (paragraphs.length < 3) return false;
    const last = paragraphs[paragraphs.length - 1];
    if (countChars(last) < 20) return false;
    const earlier = paragraphs.slice(0, -1).join('');
    if (countChars(earlier) < 40) return false;
    const earlierGrams = ngramSet(earlier, 6);
    const lastGrams = [...ngramSet(last, 6)];
    if (lastGrams.length === 0) return false;
    const hit = lastGrams.filter((gram) => earlierGrams.has(gram)).length;
    return hit / lastGrams.length >= 0.5;
}

/**
 * 研究材料 §5 的 10 条反模板清单，逐条 0/1。
 *
 * 这是**可解释的启发式**，不是检测器：每条都写明命中的证据，报告只数命中样本数。
 * 阈值按 §5 原文取（连续三句差 <5、套话 ≥2 次/千字等），改阈值要连同测试一起改。
 */
export function antiTemplateChecklist(text) {
    const body = String(text ?? '');
    const signals = writingSignals(body);
    const paragraphs = paragraphsOf(body);
    const items = [];
    const add = (id, hit, evidence = '') => {
        items.push({ id, label: ANTI_TEMPLATE_ITEMS[id - 1], hit: hit ? 1 : 0, evidence: hit ? evidence : '' });
    };
    const clicheCount = (term) => signals.cliches.find((item) => item.term === term)?.count ?? 0;

    const structure = clicheCount('首先/其次/最后');
    add(1, structure >= 2, '首先/其次/最后 ×' + structure);
    add(2, signals.summaryHits >= 2, '总结性连接词 ×' + signals.summaryHits);
    const triple = findUniformTriple(signals.sentenceLengths);
    add(3, triple !== null, triple === null ? '' : '连续三句 ' + triple.join('/') + ' 字');
    add(4, signals.clicheDensity >= 2, '套话密度 ' + signals.clicheDensity + '/千字');
    add(5, hasParallelParagraphs(paragraphs), '连续三段等长且无数字/例子/让步');
    add(6, clicheCount('不仅……而且') >= 2, '不仅…而且 ×' + clicheCount('不仅……而且'));
    add(7, signals.metaphorHits >= 2 && !METAPHOR_BOUNDARY_RE.test(body), '比喻 ×' + signals.metaphorHits + '，无失效边界');
    add(8, repeatsLastParagraph(paragraphs), '结论段与前文 6-gram 重合 ≥50%');
    add(9, signals.turnHits === 0, '无让步/反例/竞争解释');
    const digits = (body.match(/\d+(?:\.\d+)?/gu) ?? []).length;
    add(10, digits > 0 && !NUMBER_QUALIFIER_RE.test(body), '数字 ×' + digits + '，无口径/基线/来源');

    return { items, hits: items.reduce((sum, item) => sum + item.hit, 0) };
}

// ---------------------------------------------------------------- 限长

/** 目标字数 ±10%：超出标 truncated、不足标 short（两者都进报告，不改写正文）。 */
export function lengthVerdict(chars, targetChars, tolerance = TOLERANCE) {
    const minChars = Math.floor(targetChars * (1 - tolerance));
    const maxChars = Math.ceil(targetChars * (1 + tolerance));
    return {
        chars,
        targetChars,
        minChars,
        maxChars,
        truncated: chars > maxChars,
        short: chars < minChars,
    };
}

// ---------------------------------------------------------------- 提示词

function lengthHint(targetChars) {
    const verdict = lengthVerdict(targetChars, targetChars);
    return '约 ' + targetChars + ' 字（允许 ±10%，即 ' + verdict.minChars + '–' + verdict.maxChars + ' 字）';
}

/**
 * 两个臂的唯一差别在这里：baseline 只给任务与篇幅；planned 额外把 writingPlan 的
 * markdown 原样塞进 user 消息（评测要量的就是这份规划带来的差异）。
 */
export function buildMessages({ task, arm, plan }) {
    const planned = arm === 'planned';
    const system = planned
        ? '你是一位中文写作者。下面有一份写作规划：它规定每一节的篇幅、深度与必须完成的动作。严格按规划写；规划是注意力分配，不是小标题模板。只输出成品正文。'
        : '你是一位中文写作者。按下面的要求写一篇中文文稿。只输出成品正文。';
    const lines = [];
    if (planned) lines.push('## 写作规划', '', plan.markdown, '');
    lines.push('## 写作任务');
    lines.push('标题：' + task.title);
    lines.push('文体：' + task.genreLabel);
    lines.push('读者：' + task.audience);
    lines.push('目标篇幅：' + lengthHint(task.targetChars));
    lines.push('', '## 任务要求', task.brief);
    lines.push('', '只输出正文（可用 Markdown 小标题）；不要输出提纲、写作说明、自评或「以下是正文」。');
    return [
        { role: 'system', content: system },
        { role: 'user', content: lines.join('\n') },
    ];
}

const JUDGE_RUBRIC = [
    '评分锚点（只按这三条，禁止参考任何 AI 检测器分数）：',
    '- 人类自然度：5=像有判断的作者，句长与用词随内容变；3=通顺但匀速，能看出模板；1=套话连篇或机械排比。',
    '- 规划合理性：5=最长段落正是最难处，该略处真的略了；3=结构完整但平均用力；1=每节总结 + 三段式。',
    '- 内容可信度：5=每个数字有口径，论断有边界，无编造；3=大体可信，个别来源不清；1=编造事实或相关当因果。',
].join('\n');

function buildJudgeMessages(task, text) {
    const user = [
        '你是严格的中文写作评审。只按下面三个维度给 1–5 分，禁止参考任何 AI 检测器分数。',
        '',
        JUDGE_RUBRIC,
        '',
        '写作任务：' + task.title + '（' + task.genreLabel + '）',
        '任务要求：' + task.brief,
        '',
        '待评正文：',
        '',
        text,
        '',
        '只输出一个 JSON 对象，不要 Markdown 代码块、不要多余解释：',
        '{"naturalness": 1-5, "planning": 1-5, "credibility": 1-5, "comment": "一句话理由"}',
    ].join('\n');
    return [
        { role: 'system', content: '你是中文写作评审。严格按锚点打分，只输出 JSON。' },
        { role: 'user', content: user },
    ];
}

/** 解析评审 JSON；失败抛错（调用方记 judgeError，不让整轮评测崩掉）。 */
export function parseJudgeOutput(text) {
    const raw = String(text ?? '');
    const start = raw.indexOf('{');
    const end = raw.lastIndexOf('}');
    if (start < 0 || end <= start) throw new Error('评审输出不是 JSON：' + raw.slice(0, 120));
    let data;
    try {
        data = JSON.parse(raw.slice(start, end + 1));
    } catch (error) {
        throw new Error('评审输出不是合法 JSON：' + String(error.message));
    }
    const pick = (key) => {
        const value = Number(data[key]);
        if (!Number.isFinite(value) || value < 1 || value > 5) {
            throw new Error('评审字段 ' + key + ' 不在 1–5：' + JSON.stringify(data[key]));
        }
        return Math.round(value);
    };
    return {
        naturalness: pick('naturalness'),
        planning: pick('planning'),
        credibility: pick('credibility'),
        comment: typeof data.comment === 'string' ? data.comment.slice(0, 200) : '',
    };
}

// ---------------------------------------------------------------- 本地假模型（--mock）

const MOCK_POOLS = {
    a: [
        '先把两个数字摆在一起：召回率 0.71→0.86，端到端正确率 0.64→0.57。',
        '方向相反，值得停下来看。',
        '一种解释是检索噪声：候选变多，正确片段被大量相似片段淹没。',
        '另一种解释是排序漂移：证据还在，只是被排到了模型读不到的位置。',
        '我倾向噪声假设，因为召回率的提升说明片段确实进了候选集。',
        '但这不是定论。',
        '如果是噪声，把候选压到 5 条应当让正确率回升；排序假设则预测它基本不动（同一批 1,200 条问题，固定模型与提示词）。',
        '这个判据依赖一个前提：生成模型对位置不敏感，而长上下文里它未必成立。',
        '所以结论只适用于单轮、同语言、答案可由单一片段支撑的问答。',
        '多跳问题里噪声与排序会纠缠，实验分不开两者。',
        '换句话说，先量候选集的信噪比，再谈模型能力。',
        '这条边界比结论本身更重要。',
    ],
    b: [
        '设想你问聊天机器人：请介绍一位并不存在的地方志作者。',
        '它会给出生卒年、代表作，甚至一段引文。',
        '听起来像真的，其实一个字都查不到。',
        '原因不神秘：它在逐词预测下一段最像样的文字，而不是先去核对事实。',
        '这就像闭着眼睛接话，靠语感把句子接完。',
        '比喻到这里就失效了：接话的人知道自己在猜，模型没有「我不知道」这个开关。',
        '训练语料里「人物介绍」长得太像，模型就把格式套在空气上。',
        '你可以这样自查：把回答里的专名、年份、引文逐条搜一遍。',
        '搜不到，就当它没说过。',
        '更稳的做法是要求它给出可点开的来源，再自己点开。',
        '这不是万能的，公开资料本身也会错。',
        '但至少能把「编造」与「有据」分开。',
    ],
    c: [
        '结论：首页改版与次日留存下滑相关，建议今晚回滚。',
        '依据按强度排序。',
        '第一，改版组次日留存 38%，对照组 42%（n=12,400 与 n=11,900），差异稳定。',
        '第二，客服工单上升 18%，其中「找不到功能」占 46%，方向一致。',
        '第三，老用户留存持平，问题集中在新用户首日路径。',
        '如果回滚后留存回到 42% 附近，则改版是主因。',
        '如果工单结构不变，则入口文案不是唯一问题。',
        '这两条风险都能用七天数据检验。',
        '请批准回滚，并保留对照组继续观察一周。',
        '不要同时改文案，否则分不清是哪一个动作起了作用。',
    ],
};

function fillMockSection(targetChars, pool, cursor) {
    // 共享游标：整篇按同一份打乱后的句子池顺序取句，避免每节重复同一批句子
    // （重复会被人为判成「结论复读」，干扰报告里的反模板统计）。
    const picked = [];
    let chars = 0;
    let guard = 0;
    while (chars < targetChars * 0.9 && guard < pool.length * 6) {
        guard += 1;
        const sentence = pool[cursor.index % pool.length];
        cursor.index += 1;
        if (chars > 0 && chars + countChars(sentence) > targetChars * 1.05) continue;
        picked.push(sentence);
        chars += countChars(sentence);
    }
    if (picked.length === 0) {
        const shortest = [...pool].sort((left, right) => countChars(left) - countChars(right))[0];
        picked.push(shortest);
    }
    return picked.join('');
}

/** 确定性假模型：不联网、不用 Key；产物只用于验证流程（报告会标明 mock）。 */
function mockCompletion({ task, arm, plan, run }) {
    const random = mulberry32(hashString(task.id + ':' + arm + ':' + run));
    const pool = shuffle(MOCK_POOLS[task.id], random);
    const cursor = { index: 0 };
    const sections = arm === 'planned'
        ? plan.sections.map((section) => ({ heading: section.heading, targetChars: section.targetChars }))
        : task.mockBaselineHeadings.map((heading) => ({
            heading,
            targetChars: Math.round(task.targetChars / task.mockBaselineHeadings.length),
        }));
    const parts = [];
    for (const section of sections) {
        parts.push('## ' + section.heading, '', fillMockSection(section.targetChars, pool, cursor), '');
    }
    return parts.join('\n').trim();
}

function mockJudge(sample) {
    const score = Number(sample.score) || 70;
    const base = Math.max(1, Math.min(5, Math.round(score / 22)));
    return {
        naturalness: base,
        planning: base,
        credibility: Math.max(1, base - (sample.warnings.length > 3 ? 1 : 0)),
        comment: 'mock 评审：只验证 JSON 解析与报告结构。',
    };
}

// ---------------------------------------------------------------- OpenAI 兼容调用

async function callChat({
    baseUrl, apiKey, model, messages, maxTokens, temperature, fetchImpl, timeoutMs = 120000,
}) {
    const url = String(baseUrl).replace(/\/+$/u, '') + '/chat/completions';
    const doFetch = typeof fetchImpl === 'function' ? fetchImpl : globalThis.fetch;
    let response;
    try {
        response = await doFetch(url, {
            method: 'POST',
            headers: {
                'content-type': 'application/json',
                authorization: 'Bearer ' + apiKey,
            },
            body: JSON.stringify({
                model,
                messages,
                max_tokens: maxTokens,
                temperature,
                stream: false,
            }),
            signal: AbortSignal.timeout(timeoutMs),
        });
    } catch (error) {
        throw new Error('请求 ' + url + ' 失败：' + String(error?.message ?? error));
    }
    const raw = await response.text();
    if (response.ok !== true) {
        throw new Error('POST ' + url + ' 返回 ' + response.status + '：' + raw.slice(0, 300));
    }
    let payload;
    try {
        payload = JSON.parse(raw);
    } catch {
        throw new Error('响应不是 JSON：' + raw.slice(0, 200));
    }
    const content = payload?.choices?.[0]?.message?.content;
    if (typeof content !== 'string' || content.trim() === '') {
        throw new Error('模型没有返回正文（choices[0].message.content 为空）');
    }
    return { content: content.trim(), usage: payload.usage ?? null };
}

// ---------------------------------------------------------------- 选项

function apiKeyNames(options) {
    const names = [];
    if (typeof options.apiKeyEnv === 'string' && options.apiKeyEnv.trim() !== '') names.push(options.apiKeyEnv.trim());
    names.push('WRITING_EVAL_API_KEY', 'DEEPSEEK_API_KEY');
    return [...new Set(names)];
}

function resolveApiKey(options, env) {
    for (const name of apiKeyNames(options)) {
        const value = env?.[name];
        if (typeof value === 'string' && value.trim() !== '') return { key: value.trim(), envName: name };
    }
    return null;
}

function normalizeTaskIds(input) {
    const list = Array.isArray(input) ? input : String(input ?? '').split(',');
    const ids = [];
    for (const item of list) {
        const id = String(item ?? '').trim().toLowerCase();
        if (id === '') continue;
        if (!WRITING_TASKS.some((task) => task.id === id)) {
            throw new EvalParamError('不认识的任务「' + item + '」；可选 ' + WRITING_TASKS.map((task) => task.id).join('/'));
        }
        if (!ids.includes(id)) ids.push(id);
    }
    if (ids.length === 0) throw new EvalParamError('--tasks 至少给一个任务（a/b/c）');
    return ids;
}

function normalizeOptions(input = {}) {
    const options = {
        tasks: normalizeTaskIds(input.tasks ?? WRITING_TASKS.map((task) => task.id)),
        runs: input.runs ?? DEFAULT_RUNS,
        outDir: input.outDir ?? DEFAULT_OUT_DIR,
        mock: input.mock === true,
        model: input.model ?? DEFAULT_MODEL,
        baseUrl: input.baseUrl ?? DEFAULT_BASE_URL,
        apiKey: typeof input.apiKey === 'string' && input.apiKey.trim() !== '' ? input.apiKey.trim() : null,
        apiKeyEnv: input.apiKeyEnv ?? null,
        maxTokens: input.maxTokens ?? null,
        temperature: input.temperature ?? DEFAULT_TEMPERATURE,
        judgeModel: input.judgeModel ?? null,
        env: input.env ?? process.env,
        fetchImpl: input.fetchImpl ?? null,
        log: typeof input.log === 'function' ? input.log : () => {},
    };
    if (!Number.isInteger(options.runs) || options.runs < 1 || options.runs > 20) {
        throw new EvalParamError('--runs 必须是 1–20 的整数（收到 ' + JSON.stringify(input.runs) + '）');
    }
    if (typeof options.outDir !== 'string' || options.outDir.trim() === '') {
        throw new EvalParamError('--out 需要一个目录');
    }
    if (typeof options.model !== 'string' || options.model.trim() === '') {
        throw new EvalParamError('--model 需要一个模型名');
    }
    if (typeof options.baseUrl !== 'string' || !/^https?:\/\//u.test(options.baseUrl)) {
        throw new EvalParamError('--base-url 必须是 http(s) 地址');
    }
    if (options.maxTokens !== null && (!Number.isInteger(options.maxTokens) || options.maxTokens < 64 || options.maxTokens > 32000)) {
        throw new EvalParamError('--max-tokens 必须是 64–32000 的整数');
    }
    if (typeof options.temperature !== 'number' || !Number.isFinite(options.temperature) || options.temperature < 0 || options.temperature > 2) {
        throw new EvalParamError('--temperature 必须是 0–2 之间的数');
    }
    if (options.judgeModel !== null && (typeof options.judgeModel !== 'string' || options.judgeModel.trim() === '')) {
        throw new EvalParamError('--judge 需要一个模型名');
    }
    if (options.apiKeyEnv !== null && (typeof options.apiKeyEnv !== 'string' || options.apiKeyEnv.trim() === '')) {
        throw new EvalParamError('--api-key-env 需要一个环境变量名');
    }
    options.outDir = resolve(String(options.outDir));
    return options;
}

function defaultMaxTokens(targetChars) {
    return Math.max(600, Math.min(8000, Math.ceil(targetChars * 1.8) + 300));
}

function planFor(task) {
    return writingPlan({
        title: task.title,
        genre: task.genre,
        audience: task.audience,
        goal: task.goal,
        targetChars: task.targetChars,
        sections: task.sections,
    });
}

// ---------------------------------------------------------------- 产物

function renderSample(id, text) {
    return '<!-- sample ' + id + ' -->\n\n' + String(text).trim() + '\n';
}

function buildKey(context) {
    const { options, tasks, samples, failures, keyEnv } = context;
    return {
        version: 1,
        generatedAt: new Date().toISOString(),
        model: options.mock ? 'mock' : options.model,
        baseUrl: options.mock ? null : options.baseUrl,
        mock: options.mock,
        judgeModel: options.judgeModel,
        keyEnv: options.mock ? null : keyEnv,
        tasks: tasks.map((task) => task.id),
        runs: options.runs,
        arms: ARMS.slice(),
        samples: samples.map((sample) => ({
            id: sample.id,
            task: sample.task,
            arm: sample.arm,
            run: sample.run,
            position: sample.position,
            model: sample.model,
            chars: sample.chars,
            targetChars: sample.targetChars,
            minChars: sample.minChars,
            maxChars: sample.maxChars,
            truncated: sample.truncated,
            short: sample.short,
            score: sample.score,
            band: sample.band,
            warnings: sample.warnings,
            antiTemplate: sample.antiTemplate,
            antiTemplateHits: sample.antiTemplateHits,
            judge: sample.judge ?? null,
            judgeError: sample.judgeError ?? null,
        })),
        failures,
    };
}

function armStats(samples, taskId, arm) {
    const list = samples.filter((sample) => sample.task === taskId && sample.arm === arm);
    if (list.length === 0) return null;
    const scores = list.map((sample) => sample.score);
    return {
        count: list.length,
        meanScore: mean(scores),
        minScore: Math.min(...scores),
        maxScore: Math.max(...scores),
        meanChars: mean(list.map((sample) => sample.chars)),
        targetChars: list[0].targetChars,
        truncated: list.filter((sample) => sample.truncated).length,
        short: list.filter((sample) => sample.short).length,
        antiTemplate: mean(list.map((sample) => sample.antiTemplateHits)),
    };
}

function buildReport(context) {
    const { options, tasks, samples, failures } = context;
    const lines = [];
    lines.push('# 写作 A/B 评测报告', '');
    lines.push('- 生成时间：' + new Date().toISOString());
    lines.push('- 生成模型：' + (options.mock
        ? '\`mock\`（本地确定性假模型）—— **只验证流程与产物结构，不代表模型写作表现**'
        : '\`' + options.model + '\`｜baseUrl：' + options.baseUrl));
    if (options.judgeModel !== null) {
        lines.push('- 评审模型：\`' + options.judgeModel + '\`' + (options.mock ? '（mock）' : ''));
    }
    lines.push('- 任务：' + tasks.map((task) => task.id + '（' + task.title + '）').join('；'));
    lines.push('- 每臂轮次：' + options.runs + '｜样本：' + samples.length + (failures.length > 0 ? '｜失败：' + failures.length : ''));
    lines.push('- 臂：baseline（普通写作提示）、planned（先 writingPlan，再按规划写）');
    lines.push('- 位置控制：每任务第 1 轮 baseline→planned，第 2 轮 planned→baseline（交替）；key.json 的 position 字段记录左右位置。');
    lines.push('- 限长：目标字数 ±10%；超出标 truncated、不足标 short，正文按模型原样保留（截断会破坏盲评可读性）。');
    lines.push('');

    lines.push('## 汇总', '');
    lines.push('| 任务 | 臂 | 样本 | 平均分 | 分数区间 | 平均字数 | 目标字数 | 超限 | 偏短 | 反模板均值 |');
    lines.push('| --- | --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: |');
    for (const task of tasks) {
        for (const arm of ARMS) {
            const stats = armStats(samples, task.id, arm);
            if (stats === null) continue;
            lines.push('| ' + task.id + ' | ' + arm + ' | ' + stats.count + ' | ' + formatNumber(stats.meanScore)
                + ' | ' + stats.minScore + '–' + stats.maxScore + ' | ' + formatNumber(stats.meanChars, 0)
                + ' | ' + stats.targetChars + ' | ' + stats.truncated + ' | ' + stats.short + ' | '
                + formatNumber(stats.antiTemplate, 1) + ' |');
        }
    }
    lines.push('');

    lines.push('## 反模板清单（§5，0/1 计数）', '');
    lines.push('| # | 条目 | baseline 命中 | planned 命中 | 合计 |');
    lines.push('| ---: | --- | ---: | ---: | ---: |');
    for (let index = 0; index < ANTI_TEMPLATE_ITEMS.length; index += 1) {
        const id = index + 1;
        const countFor = (arm) => samples.filter((sample) => sample.arm === arm
            && sample.antiTemplate[String(id)] === 1).length;
        const base = countFor('baseline');
        const planned = countFor('planned');
        lines.push('| ' + id + ' | ' + ANTI_TEMPLATE_ITEMS[index] + ' | ' + base + ' | ' + planned + ' | ' + (base + planned) + ' |');
    }
    lines.push('');

    lines.push('## 长度与限长', '');
    lines.push('| 任务 | 目标 | 允许区间 | 平均字数 | 最短 | 最长 | 超限 | 偏短 |');
    lines.push('| --- | ---: | --- | ---: | ---: | ---: | ---: | ---: |');
    for (const task of tasks) {
        const list = samples.filter((sample) => sample.task === task.id);
        if (list.length === 0) continue;
        const chars = list.map((sample) => sample.chars);
        const verdict = lengthVerdict(task.targetChars, task.targetChars);
        lines.push('| ' + task.id + ' | ' + task.targetChars + ' | ' + verdict.minChars + '–' + verdict.maxChars
            + ' | ' + formatNumber(mean(chars), 0) + ' | ' + Math.min(...chars) + ' | ' + Math.max(...chars)
            + ' | ' + list.filter((sample) => sample.truncated).length
            + ' | ' + list.filter((sample) => sample.short).length + ' |');
    }
    lines.push('');

    lines.push('## 自检分数与常见警告', '');
    const bands = new Map();
    for (const sample of samples) bands.set(sample.band, (bands.get(sample.band) ?? 0) + 1);
    lines.push('- band 计数：' + [...bands.entries()].map(([band, count]) => band + ' ' + count).join('；'));
    const warningCounts = new Map();
    for (const sample of samples) {
        for (const code of sample.warnings) warningCounts.set(code, (warningCounts.get(code) ?? 0) + 1);
    }
    const topWarnings = [...warningCounts.entries()].sort((left, right) => right[1] - left[1]).slice(0, 8);
    lines.push('- 最常见 warnings：' + (topWarnings.length === 0
        ? '（无）'
        : topWarnings.map(([code, count]) => code + '×' + count).join('、')));
    lines.push('- 每份样本的 score / band / warnings / 反模板清单见 key.json；正文见 samples/。');
    lines.push('');

    if (options.judgeModel !== null) {
        lines.push('## 评审模型打分（§5 三维 1–5）', '');
        lines.push('| 任务 | 臂 | 样本 | 自然度 | 规划合理性 | 可信度 | 解析失败 |');
        lines.push('| --- | --- | ---: | ---: | ---: | ---: | ---: |');
        for (const task of tasks) {
            for (const arm of ARMS) {
                const list = samples.filter((sample) => sample.task === task.id && sample.arm === arm);
                const judged = list.filter((sample) => sample.judge !== null && sample.judge !== undefined);
                if (list.length === 0) continue;
                const dim = (key) => (judged.length === 0 ? '—' : formatNumber(mean(judged.map((sample) => sample.judge[key])), 1));
                lines.push('| ' + task.id + ' | ' + arm + ' | ' + list.length + ' | ' + dim('naturalness')
                    + ' | ' + dim('planning') + ' | ' + dim('credibility')
                    + ' | ' + list.filter((sample) => sample.judgeError).length + ' |');
            }
        }
        lines.push('');
    }

    if (failures.length > 0) {
        lines.push('## 生成失败', '');
        for (const failure of failures) {
            lines.push('- ' + failure.task + ' / ' + failure.arm + ' / run ' + failure.run + '：' + failure.error);
        }
        lines.push('');
    }

    lines.push('## 盲评指引', '');
    lines.push('- samples/*.md 只含编号与正文；编号→任务/臂/轮次/模型在 key.json，评审时不要打开 key.json。');
    lines.push('- 按 §5 三维 1–5 锚点评分，**禁止参考任何 AI 检测器分数**；≥3 人取中位数，分歧 ≥2 分交第三人仲裁。');
    lines.push('- 位置偏差：同一任务两轮左右各出现一次；单轮差 >1 分记「位置敏感」，看两轮均值。');
    lines.push('- 长度偏差：超出目标 ±10% 的样本已标 truncated；按 §5 另算「分数 vs 字数」相关，>0.5 则该维度作废重测。');
    lines.push('- 反模板清单是启发式（可解释但非检测器），只用来定位模板味，不替代人工评分。');
    return lines.join('\n') + '\n';
}

function writeArtifacts(context) {
    const { options, key, report, samples } = context;
    const outDir = options.outDir;
    const samplesDir = join(outDir, 'samples');
    rmSync(samplesDir, { recursive: true, force: true });
    mkdirSync(samplesDir, { recursive: true });
    for (const sample of samples) {
        writeFileSync(join(samplesDir, sample.id + '.md'), renderSample(sample.id, sample.text), 'utf8');
    }
    const keyPath = join(outDir, 'key.json');
    const reportPath = join(outDir, 'report.md');
    writeFileSync(keyPath, JSON.stringify(key, null, 2) + '\n', 'utf8');
    writeFileSync(reportPath, report, 'utf8');
    return { samplesDir, keyPath, reportPath };
}

// ---------------------------------------------------------------- 主流程

/**
 * 跑一轮评测。
 *
 * 返回值：{ ok, skipped?, outDir, samples, failures, samplesDir, keyPath, reportPath }
 * 没有 Key 且不是 mock 时返回 { ok: true, skipped: true, envNames }，**不写任何文件**。
 */
export async function runWritingEval(input = {}) {
    const options = normalizeOptions(input);
    const log = options.log;
    const tasks = options.tasks.map((id) => WRITING_TASKS.find((task) => task.id === id));
    const plans = new Map(tasks.map((task) => [task.id, planFor(task)]));

    let apiKey = options.apiKey;
    let keyEnv = options.apiKey === null ? null : '(--api-key)';
    if (!options.mock && apiKey === null) {
        const found = resolveApiKey(options, options.env);
        if (found === null) {
            return { ok: true, skipped: true, reason: 'no-api-key', envNames: apiKeyNames(options) };
        }
        apiKey = found.key;
        keyEnv = found.envName;
    }

    const samples = [];
    const failures = [];
    const total = tasks.length * options.runs * ARMS.length;
    let index = 0;
    for (const task of tasks) {
        const plan = plans.get(task.id);
        for (let run = 1; run <= options.runs; run += 1) {
            // 换序：第 1 轮 baseline 在前，第 2 轮 planned 在前（控位置偏差）。
            const order = run % 2 === 1 ? ['baseline', 'planned'] : ['planned', 'baseline'];
            for (let position = 1; position <= order.length; position += 1) {
                const arm = order[position - 1];
                index += 1;
                const label = '[' + index + '/' + total + '] task=' + task.id + ' arm=' + arm + ' run=' + run;
                try {
                    let text;
                    if (options.mock) {
                        text = mockCompletion({ task, arm, plan, run });
                    } else {
                        const result = await callChat({
                            baseUrl: options.baseUrl,
                            apiKey,
                            model: options.model,
                            messages: buildMessages({ task, arm, plan }),
                            maxTokens: options.maxTokens ?? defaultMaxTokens(task.targetChars),
                            temperature: options.temperature,
                            fetchImpl: options.fetchImpl,
                        });
                        text = result.content;
                    }
                    const audit = arm === 'planned' ? writingAudit(text, plan) : writingAudit(text);
                    const length = lengthVerdict(audit.chars, task.targetChars);
                    const checklist = antiTemplateChecklist(text);
                    samples.push({
                        id: null,
                        task: task.id,
                        arm,
                        run,
                        position,
                        model: options.mock ? 'mock' : options.model,
                        text,
                        chars: audit.chars,
                        targetChars: task.targetChars,
                        minChars: length.minChars,
                        maxChars: length.maxChars,
                        truncated: length.truncated,
                        short: length.short,
                        score: audit.score,
                        band: audit.band,
                        warnings: audit.warnings.map((item) => item.code),
                        antiTemplate: Object.fromEntries(checklist.items.map((item) => [String(item.id), item.hit])),
                        antiTemplateHits: checklist.hits,
                        judge: null,
                        judgeError: null,
                    });
                    log(label + ' chars=' + audit.chars + ' score=' + audit.score
                        + (length.truncated ? ' truncated' : '') + (length.short ? ' short' : ''));
                } catch (error) {
                    const message = String(error?.message ?? error);
                    failures.push({ task: task.id, arm, run, position, error: message });
                    log(label + ' ERROR ' + message);
                }
            }
        }
    }

    // 匿名化：先确定性洗牌再编号，编号顺序不泄露臂与轮次。
    const random = mulberry32(hashString('shuffle:' + options.tasks.join(',') + ':' + options.runs + ':' + options.model));
    const ordered = shuffle(samples, random);
    ordered.forEach((sample, position) => {
        sample.id = String(position + 1).padStart(4, '0');
    });

    if (options.judgeModel !== null) {
        for (const sample of ordered) {
            try {
                if (options.mock) {
                    sample.judge = mockJudge(sample);
                } else {
                    const task = tasks.find((item) => item.id === sample.task);
                    const judged = await callChat({
                        baseUrl: options.baseUrl,
                        apiKey,
                        model: options.judgeModel,
                        messages: buildJudgeMessages(task, sample.text),
                        maxTokens: 400,
                        temperature: 0,
                        fetchImpl: options.fetchImpl,
                    });
                    sample.judge = parseJudgeOutput(judged.content);
                }
            } catch (error) {
                sample.judge = null;
                sample.judgeError = String(error?.message ?? error);
                log('judge ERROR sample=' + sample.id + ' ' + sample.judgeError);
            }
        }
    }

    const context = { options, tasks, samples: ordered, failures, keyEnv };
    const key = buildKey(context);
    const report = buildReport(context);
    const paths = writeArtifacts({ options, key, report, samples: ordered });
    return {
        ok: failures.length === 0,
        skipped: false,
        outDir: options.outDir,
        samples: ordered,
        failures,
        ...paths,
    };
}

// ---------------------------------------------------------------- CLI

const HELP = [
    '用法：node scripts/writing-eval.mjs [选项]',
    '',
    '  --tasks a,b,c        任务（默认全部 a,b,c）',
    '  --runs 2             每臂轮次（默认 2；两轮交换臂顺序）',
    '  --out <dir>          输出目录（默认 ' + DEFAULT_OUT_DIR + '）',
    '  --mock               本地确定性假模型，不联网（自测 / CI 用）',
    '  --self-test          用 mock 跑完整流程并检查产物结构（含无 Key 路径 exit 0）',
    '  --model <id>         生成模型（默认 ' + DEFAULT_MODEL + '）',
    '  --base-url <url>     OpenAI 兼容端点（默认 ' + DEFAULT_BASE_URL + '）',
    '  --api-key-env <名>   指定读哪个环境变量取 Key（优先于内置顺序）',
    '  --max-tokens <n>     单次生成上限（默认按任务字数自动算）',
    '  --temperature <n>    采样温度（默认 ' + DEFAULT_TEMPERATURE + '）',
    '  --judge <model>      额外让一个模型按 §5 三维锚点打分（默认关闭）',
    '  -h, --help           显示本说明',
    '',
    'Key 优先级：--api-key-env 指定名 → WRITING_EVAL_API_KEY → DEEPSEEK_API_KEY。',
    '没有 Key 时打印 SKIP 并 exit 0（不写任何假结果）；--mock 不需要 Key。',
    '退出码：0 成功（含 SKIP）；1 生成/写盘失败；2 参数错误。',
].join('\n');

function parseCliArgs(argv) {
    const options = {
        tasks: null,
        runs: DEFAULT_RUNS,
        outDir: DEFAULT_OUT_DIR,
        mock: false,
        selfTest: false,
        model: DEFAULT_MODEL,
        baseUrl: DEFAULT_BASE_URL,
        apiKeyEnv: null,
        maxTokens: null,
        temperature: DEFAULT_TEMPERATURE,
        judgeModel: null,
        help: false,
    };
    for (let index = 0; index < argv.length; index += 1) {
        let arg = argv[index];
        let inline = null;
        const equalsAt = arg.indexOf('=');
        if (arg.startsWith('--') && equalsAt > 2) {
            inline = arg.slice(equalsAt + 1);
            arg = arg.slice(0, equalsAt);
        }
        const value = () => {
            if (inline !== null) return inline;
            const next = argv[index + 1];
            if (next === undefined || (next.startsWith('--') && next !== '--')) {
                throw new EvalParamError(arg + ' 需要一个值');
            }
            index += 1;
            return next;
        };
        switch (arg) {
            case '--tasks': options.tasks = value(); break;
            case '--runs': options.runs = Number(value()); break;
            case '--out': options.outDir = value(); break;
            case '--mock': options.mock = true; break;
            case '--self-test': options.selfTest = true; break;
            case '--model': options.model = value(); break;
            case '--base-url': options.baseUrl = value(); break;
            case '--api-key-env': options.apiKeyEnv = value(); break;
            case '--max-tokens': options.maxTokens = Number(value()); break;
            case '--temperature': options.temperature = Number(value()); break;
            case '--judge': options.judgeModel = value(); break;
            case '--help': case '-h': options.help = true; break;
            default:
                throw new EvalParamError('不认识的参数：' + argv[index]);
        }
    }
    return options;
}

function printSummary(result) {
    const byArm = new Map();
    for (const sample of result.samples) {
        const entry = byArm.get(sample.arm) ?? { count: 0, chars: 0, score: 0 };
        entry.count += 1;
        entry.chars += sample.chars;
        entry.score += sample.score;
        byArm.set(sample.arm, entry);
    }
    console.log('writing-eval：' + result.samples.length + ' 份样本 → ' + result.outDir);
    for (const arm of ARMS) {
        const entry = byArm.get(arm);
        if (entry === undefined) continue;
        console.log('  ' + arm + '：' + entry.count + ' 份，平均分 ' + formatNumber(entry.score / entry.count)
            + '，平均 ' + formatNumber(entry.chars / entry.count, 0) + ' 字');
    }
    console.log('  samples/ → ' + result.samplesDir);
    console.log('  key.json → ' + result.keyPath + '（编号→任务/臂/轮次/模型，盲评时不要打开）');
    console.log('  report.md → ' + result.reportPath);
    if (result.failures.length > 0) {
        console.log('  失败 ' + result.failures.length + ' 份：' + result.failures.map((item) => item.task + '/' + item.arm + '/run' + item.run).join('、'));
    }
}

function samePath(left, right) {
    const a = resolve(left);
    const b = resolve(right);
    return process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b;
}

function stripApiKeys(env) {
    const copy = { ...env };
    delete copy.WRITING_EVAL_API_KEY;
    delete copy.DEEPSEEK_API_KEY;
    delete copy.WRITING_EVAL_TEST_KEY;
    delete copy.WRITING_EVAL_JUDGE_TEST_KEY;
    return copy;
}

function spawnNode(args, env) {
    return new Promise((done) => {
        const child = spawn(process.execPath, [SELF, ...args], {
            cwd: process.cwd(),
            env,
            stdio: ['ignore', 'pipe', 'pipe'],
            windowsHide: true,
        });
        let stdout = '';
        let stderr = '';
        const timer = setTimeout(() => {
            try { child.kill('SIGKILL'); } catch { /* ignore */ }
        }, 180000);
        child.stdout.on('data', (chunk) => { stdout += chunk; });
        child.stderr.on('data', (chunk) => { stderr += chunk; });
        child.on('close', (code) => {
            clearTimeout(timer);
            done({ code, stdout, stderr });
        });
    });
}

/** --self-test：mock 跑完整流程 + 无 Key 子进程 exit 0；全部离线。 */
async function runSelfTest() {
    const failures = [];
    let checks = 0;
    const check = (condition, label, detail = '') => {
        checks += 1;
        if (!condition) failures.push(detail === '' ? label : label + ' —— ' + detail);
    };
    const root = mkdtempSync(join(tmpdir(), 'writing-eval-selftest-'));
    try {
        const result = await runWritingEval({
            tasks: ['a'],
            runs: 1,
            outDir: join(root, 'run'),
            mock: true,
            env: {},
            log: () => {},
        });
        check(result.ok === true && result.skipped === false, 'mock 全流程应成功');
        check(result.samples.length === 2, '1 任务 × 2 臂 × 1 轮 = 2 份样本', '实际 ' + result.samples.length);
        const files = existsSync(result.samplesDir) ? readdirSync(result.samplesDir).filter((name) => name.endsWith('.md')) : [];
        check(files.length === 2, 'samples/ 下有 2 个 .md', '实际 ' + files.length);
        check(existsSync(result.keyPath), 'key.json 已写出');
        check(existsSync(result.reportPath), 'report.md 已写出');
        const key = JSON.parse(readFileSync(result.keyPath, 'utf8'));
        check(key.samples.length === 2, 'key.json 记录 2 份样本');
        check(key.samples.every((sample) => sample.task === 'a' && ARMS.includes(sample.arm) && sample.run === 1), 'key 里的任务/臂/轮次');
        check(key.samples.every((sample) => Number.isInteger(sample.score) && typeof sample.band === 'string'
            && Array.isArray(sample.warnings) && typeof sample.antiTemplate === 'object'), 'key 里的 score/band/warnings/反模板');
        check(new Set(key.samples.map((sample) => sample.id)).size === 2, '样本编号不重复');
        const report = readFileSync(result.reportPath, 'utf8');
        check(report.includes('反模板'), 'report 含反模板清单');
        check(report.includes('长度'), 'report 含长度统计');
        check(report.includes('planned'), 'report 标明两个臂');
        check(report.includes('mock'), 'report 标明 mock 模式');
        const bodies = files.map((name) => readFileSync(join(result.samplesDir, name), 'utf8'));
        check(bodies.every((body) => body.startsWith('<!-- sample ')), 'samples 只以编号开头');
        check(bodies.every((body) => !/baseline|planned|mock/iu.test(body)), 'samples 不泄露臂名/模型名');

        const skipOut = join(root, 'skip');
        const child = await spawnNode(['--tasks', 'a', '--out', skipOut], stripApiKeys(process.env));
        check(child.code === 0, '无 Key 子进程 exit 0', '实际 ' + child.code + ' stderr=' + child.stderr.slice(0, 200));
        check(/SKIP/u.test(child.stdout), '无 Key 子进程打印 SKIP');
        check(/DEEPSEEK_API_KEY/u.test(child.stdout), '无 Key 子进程提示所需环境变量');
        check(!existsSync(join(skipOut, 'samples')), '无 Key 不写假结果');
    } finally {
        rmSync(root, { recursive: true, force: true });
    }
    if (failures.length > 0) {
        console.error('SELF-TEST FAIL：' + failures.length + ' 项不通过（共 ' + checks + ' 项检查）');
        for (const line of failures) console.error('  ✗ ' + line);
        return 1;
    }
    console.log('SELF-TEST PASS checks=' + checks);
    return 0;
}

async function main(argv) {
    let options;
    try {
        options = parseCliArgs(argv);
    } catch (error) {
        console.error('RED  参数错误：' + String(error?.message ?? error));
        return 2;
    }
    if (options.help) {
        console.log(HELP);
        return 0;
    }
    if (options.selfTest) return runSelfTest();
    try {
        const result = await runWritingEval({
            ...options,
            env: process.env,
            log: (line) => console.log(line),
        });
        if (result.skipped) {
            console.log('SKIP writing-eval：没有可用的 API Key。');
            console.log('  需要设置 ' + result.envNames.join(' 或 ') + '；--api-key-env 可指定别的变量名。');
            console.log('  没有写出任何结果（不伪造模型输出）；离线跑流程用 --mock，自测用 --self-test。');
            return 0;
        }
        printSummary(result);
        return result.ok ? 0 : 1;
    } catch (error) {
        if (error instanceof EvalParamError) {
            console.error('RED  参数错误：' + error.message);
            return 2;
        }
        console.error('RED  writing-eval 失败：' + String(error?.message ?? error));
        return 1;
    }
}

const invokedDirectly = process.argv[1] !== undefined && samePath(process.argv[1], SELF);
if (invokedDirectly) {
    const code = await main(process.argv.slice(2));
    process.exit(code);
}
