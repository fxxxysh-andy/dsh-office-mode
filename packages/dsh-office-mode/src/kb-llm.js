/**
 * 知识库的可选大模型增强（第九十二轮）：语义切片 + 片段摘要，默认关闭。
 *
 * 三条硬纪律：
 *   1. **默认不调用**：mode='off' 时这里一个网络请求都不发；调用方仍旧走
 *      kb.js 的确定性 splitDocument。开关是用户的显式选择，不是部署副作用。
 *   2. **原文保真**：模型**不产出正文**。它只拿到「小节编号」，返回
 *      {from,to,title,summary} 这样的分组意图；真正的块正文一律由
 *      source.slice(start,end) 得到。因此模型无法改写、增删、遗漏原文 ——
 *      这是结构上的保证，不是靠提示词祈祷。
 *   3. **有界可续**：千万字不整篇进模型。按小节装箱，每箱 <= maxInputChars；
 *      单个文档的任务可以只跑一段（paused）后恢复；总调用数有硬上限，
 *      超预算明确报 over-budget / budget，绝不把「没跑完」报成成功。
 *
 * 模型路由（不硬编码厂商）：
 *   - mode='host'：走宿主 ctx.llm（dsh-llm 的 LlmRuntime）已注册的 provider/model，
 *     与宿主设置页里配置的模型是同一套路由；插件不 import 裸包名，只用结构化对象。
 *   - mode='openai'：任何 OpenAI 兼容的 POST {baseURL}/chat/completions 端点，
 *     provider/baseURL/model/apiKey 全部可配，不绑定单一厂商。
 *
 * @module dsh-office-mode/kb-llm
 */
import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile, rm } from 'node:fs/promises';
import path from 'node:path';
import { asText } from './web-errors.js';
import { chunkIdOf, chunkOverlapFor } from './kb.js';

/**
 * 提示词/回执协议版本：改了分组语义就要进 docHash（不同切法的块不能共用 id 空间）。
 * 第 2 版（第一百零一轮）：块起点按 `chunkOverlapFor` 向前重叠，块正文与第 1 版不同，
 * 旧进度文件按批数不匹配会被拒续跑 —— 这是有意的，不拿新切法去接旧游标。
 * 第 3 版（第一百一十轮）：单箱分组数受输出 token 预算约束，同一份原文的箱与块边界变了。
 */
export const KB_LLM_PROTOCOL_VERSION = 3;

/**
 * 每个分组在 JSON 回复里大约要占多少 token（含 from/to/title/summary 与标点）。
 * 这是保守上限：中文摘要 120 字约 80~100 token，标题/编号/JSON 约 40~60。
 * 用它把「单箱小节数」压到输出预算以内，避免 60 组 JSON 被 1500 token 截断。
 */
export const KB_LLM_TOKENS_PER_GROUP = 180;

/** 可选默认值；全部可被设置页覆盖。 */
export const DEFAULT_KB_LLM = {
    mode: 'off',
    chunking: false,
    summary: false,
    provider: '',
    model: '',
    baseURL: '',
    apiKey: '',
    apiKeyEnv: '',
    timeoutMs: 30_000,
    maxInputChars: 6_000,
    maxOutputTokens: 1_500,
    maxCallsPerRun: 20,
    maxCalls: 400,
    maxChunkChars: 3_000,
    minChunkChars: 400,
    maxUnitsPerBatch: 60,
};

function numberIn(value, fallback, min, max) {
    const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, Math.trunc(n)));
}

/** 配置项：缺省、无法解析或小于下限（含负数）时落默认；超过上限才夹到上限。 */
function pickIn(value, fallback, min, max) {
    const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
    if (!Number.isFinite(n) || n < min) return fallback;
    return Math.min(max, Math.trunc(n));
}

/**
 * 把配置收口成运行期参数（坏值落默认，绝不抛；与 kb-embed 的 normalize 同一条纪律）。
 * 注意：这里**不**因为 mode='host' 就自动把 chunking/summary 打开 ——
 * 「选了通道」不等于「选了用途」，两个开关必须由用户分别明确。
 */
export function normalizeKbLlmConfig(raw = {}) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const mode = ['off', 'host', 'openai'].includes(source.mode) ? source.mode : 'off';
    const cfg = {
        mode,
        chunking: source.chunking === true,
        summary: source.summary === true,
        provider: asText(source.provider),
        model: asText(source.model),
        baseURL: asText(source.baseURL).replace(/\/+$/, ''),
        apiKey: asText(source.apiKey),
        apiKeyEnv: asText(source.apiKeyEnv),
        timeoutMs: pickIn(source.timeoutMs, DEFAULT_KB_LLM.timeoutMs, 50, 120_000),
        // 区间与设置页 schema 保持一致：设置页能存下的值，运行期不得再悄悄夹一次。
        maxInputChars: pickIn(source.maxInputChars, DEFAULT_KB_LLM.maxInputChars, 1_000, 60_000),
        maxOutputTokens: pickIn(source.maxOutputTokens, DEFAULT_KB_LLM.maxOutputTokens, 200, 8_000),
        maxCallsPerRun: pickIn(source.maxCallsPerRun, DEFAULT_KB_LLM.maxCallsPerRun, 1, 100),
        maxCalls: pickIn(source.maxCalls, DEFAULT_KB_LLM.maxCalls, 1, 5_000),
        maxChunkChars: pickIn(source.maxChunkChars, DEFAULT_KB_LLM.maxChunkChars, 500, 10_000),
        minChunkChars: pickIn(source.minChunkChars, DEFAULT_KB_LLM.minChunkChars, 100, 3_000),
        maxUnitsPerBatch: pickIn(source.maxUnitsPerBatch, DEFAULT_KB_LLM.maxUnitsPerBatch, 10, 200),
    };
    if (cfg.minChunkChars >= cfg.maxChunkChars) cfg.minChunkChars = Math.max(200, Math.floor(cfg.maxChunkChars / 3));
    return cfg;
}

/**
 * 单箱最多允许模型返回几个分组。
 *
 * 只有 maxUnitsPerBatch 是不够的：它表达「输入侧一箱放多少小节」，模型还要
 * 为每一组写出 title + summary。1500 输出 token 装不下 60 组 JSON，于是
 * finish_reason=length、整批 JSON 作废，既浪费 token 又只能降级。
 * 这里按输出预算收紧，让「一箱能问」和「一箱答得完」是同一个数。
 */
export function kbLlmBatchLimit(config = {}) {
    const cfg = normalizeKbLlmConfig(config);
    return Math.max(1, Math.min(cfg.maxUnitsPerBatch, Math.floor(cfg.maxOutputTokens / KB_LLM_TOKENS_PER_GROUP)));
}

/** 取 Key：字面量优先，其次环境变量（与 embedding 同一套口径）。 */
export function kbLlmKeyOf(config, env = process.env) {
    const cfg = normalizeKbLlmConfig(config);
    if (cfg.apiKey !== '') return cfg.apiKey;
    if (cfg.apiKeyEnv !== '') return asText(env?.[cfg.apiKeyEnv]);
    return '';
}

/**
 * 这份配置现在能不能真的调用模型？不能时给出**用户看得懂**的原因。
 * @returns {{ok:boolean, reason:string}}
 */
export function kbLlmReady(config) {
    const cfg = normalizeKbLlmConfig(config);
    if (cfg.mode === 'off') return { ok: false, reason: '模型增强未启用（默认关闭，本地确定性切片照常工作）。' };
    if (cfg.mode === 'openai' && (cfg.baseURL === '' || cfg.model === '')) {
        return { ok: false, reason: 'OpenAI 兼容通道缺少 baseURL 或 model。' };
    }
    if (cfg.mode === 'host' && (cfg.provider === '' || cfg.model === '')) {
        return { ok: false, reason: '宿主通道缺少 provider 或 model；请先选择宿主里已配置的模型。' };
    }
    if (!cfg.chunking && !cfg.summary) {
        return { ok: false, reason: '已选模型通道，但「模型切片」和「模型摘要」都没勾选。' };
    }
    return { ok: true, reason: '' };
}

/** 回执用的短标签（绝不带 Key）。 */
export function describeKbLlm(config) {
    const cfg = normalizeKbLlmConfig(config);
    const route = cfg.mode === 'openai'
        ? (cfg.model === '' ? 'openai-compatible' : 'openai-compatible / ' + cfg.model)
        : cfg.mode === 'host'
            ? [cfg.provider, cfg.model].filter((item) => item !== '').join(' / ')
            : '';
    return {
        mode: cfg.mode,
        model: route,
        chunking: cfg.chunking,
        summary: cfg.summary,
        maxInputChars: cfg.maxInputChars,
        maxCalls: cfg.maxCalls,
        maxCallsPerRun: cfg.maxCallsPerRun,
    };
}

/**
 * 切块档（进 docHash）：模型切片与确定性切片必须落在不同的块 id 空间。
 * 同一模型 + 同一装箱上限 = 同一个档，重入才幂等。
 */
export function kbLlmProfile(config) {
    const cfg = normalizeKbLlmConfig(config);
    const model = cfg.mode === 'host' ? [cfg.provider, cfg.model].join('/') : cfg.model;
    // 这些参数都会改变真实切块（装箱上限、单块上限、回退下限、单箱小节数），
    // 必须进 profile；否则改参数后 docHash 不变，旧块会被误判成同一档。
    // maxOutputTokens 也进来：它决定单箱能答完几组，从而改变实际箱数。
    // 只开摘要时（chunking=false）调用方仍用确定性 docHash，这里变化不会触发重切。
    // 重叠量由 maxChunkChars 决定（chunkOverlapFor），所以它已经随 maxChunkChars 进 profile；
    // 这里显式再写一遍是为了让「改了重叠算法」也换 profile —— 只写一个派生量容易被
    // 后来的重构悄悄改掉而不换档。
    return 'llm:' + KB_LLM_PROTOCOL_VERSION + ':' + model
        + ':' + cfg.maxInputChars
        + ':' + cfg.maxChunkChars
        + ':' + cfg.minChunkChars
        + ':' + cfg.maxUnitsPerBatch
        + ':' + cfg.maxOutputTokens
        + ':' + chunkOverlapFor(cfg.maxChunkChars);
}

// ── 原文装箱（纯函数，确定性） ──────────────────────────────────────────────

/** Markdown 标题行（与 kb.js 的 HEADING 同口径）。 */
const HEADING = /^(#{1,6})[ \t]+(\S.*?)[ \t]*$/;
const BREAK_AFTER = /[\n。！？；：.!?;:、，,）)】」』]$/;

/**
 * 把正文拆成「小节」：标题行各成一个小节，连续非空行合成一个段落小节。
 *
 * 为什么以小节而不是字符为单位：模型只被允许在这些边界上分组，于是切点永远
 * 落在原文的换行处 —— 块正文是原文切片，天然不可能被改写或漏字。
 * 元数据（span）来自这里，**不来自模型**。
 */
export function unitsOf(text, { maxUnitChars = 2_000 } = {}) {
    const source = typeof text === 'string' ? text : '';
    const limit = Math.max(200, numberIn(maxUnitChars, 2_000, 200, 48_000));
    const units = [];
    const lines = [];
    {
        let offset = 0;
        for (const raw of source.split('\n')) {
            const clean = raw.endsWith('\r') ? raw.slice(0, -1) : raw;
            lines.push({ text: clean, start: offset, end: offset + clean.length });
            offset += raw.length + 1;
        }
    }
    let stack = [];
    let paragraph = null;
    const outlineOf = () => stack.filter((item) => typeof item === 'string' && item !== '').join(' / ');
    const flush = () => {
        if (paragraph === null) return;
        const { start, end } = paragraph;
        paragraph = null;
        const outline = outlineOf();
        const chapter = stack.find((item) => typeof item === 'string' && item !== '') ?? '';
        pushUnit(units, source, { start, end, outline, chapter, limit });
    };
    for (const line of lines) {
        const heading = HEADING.exec(line.text);
        if (heading !== null) {
            flush();
            const depth = heading[1].length;
            stack = stack.slice(0, depth - 1);
            stack[depth - 1] = heading[2].trim();
            const outline = outlineOf();
            units.push({
                index: units.length,
                start: line.start,
                end: line.end,
                outline,
                chapter: stack.find((item) => typeof item === 'string' && item !== '') ?? '',
                heading: true,
            });
            continue;
        }
        if (line.text.trim() === '') {
            flush();
            continue;
        }
        if (paragraph === null) paragraph = { start: line.start, end: line.end };
        else paragraph.end = line.end;
    }
    flush();
    // 让 index 与位置一致（pushUnit 可能追加多个硬切小节）
    for (let i = 0; i < units.length; i += 1) units[i].index = i;
    return units;
}

/** 超长段落按标点硬切成若干小节（仍是原文切片，不引入任何新字符）。 */
function pushUnit(units, source, { start, end, outline, chapter, limit }) {
    if (end - start <= limit) {
        units.push({ index: units.length, start, end, outline, chapter, heading: false });
        return;
    }
    let cursor = start;
    while (end - cursor > limit) {
        const hard = cursor + limit;
        const floor = Math.max(cursor + 60, hard - Math.floor(limit / 6));
        let cut = hard;
        for (let index = hard; index > floor; index -= 1) {
            if (BREAK_AFTER.test(source.slice(index - 1, index))) {
                cut = index;
                break;
            }
        }
        units.push({ index: units.length, start: cursor, end: cut, outline, chapter, heading: false });
        cursor = cut;
    }
    if (end > cursor) {
        units.push({ index: units.length, start: cursor, end, outline, chapter, heading: false });
    }
}

/**
 * 把小节装箱：每箱 <= maxInputChars，尽量不跨章，单箱小节数同时受
 * maxUnitsPerBatch 与输出 token 预算（kbLlmBatchLimit）约束。
 * 这个划分是**确定性**的：同一份文本、同一份配置，任何时候都得到同样的箱。
 */
export function planKbBatches(text, config = {}) {
    const cfg = normalizeKbLlmConfig(config);
    const units = unitsOf(text, { maxUnitChars: Math.max(cfg.minChunkChars, Math.floor(cfg.maxInputChars / 2)) });
    const batchLimit = kbLlmBatchLimit(cfg);
    const batches = [];
    let current = null;
    for (const unit of units) {
        const size = (unit.end - unit.start) + 8; // 8 = 编号与换行开销
        const chapterChanged = current !== null && unit.chapter !== current.chapter && current.chapter !== '';
        if (current === null
            || current.chars + size > cfg.maxInputChars
            || current.units.length >= batchLimit
            || chapterChanged) {
            current = { index: batches.length, from: unit.index, to: unit.index, chars: 0, units: [], chapter: unit.chapter };
            batches.push(current);
        }
        current.units.push(unit);
        current.chars += size;
        current.to = unit.index;
    }
    return { units, batches };
}

/** 预估：这份正文要切几箱、几章、大约多少次调用（用户先看数，再决定开不开）。 */
export function estimateKbLlmJob(text, config = {}) {
    const cfg = normalizeKbLlmConfig(config);
    const { units, batches } = planKbBatches(text, cfg);
    const chars = typeof text === 'string' ? text.length : 0;
    const calls = batches.length;
    return {
        chars,
        units: units.length,
        chapters: new Set(units.map((unit) => unit.chapter)).size,
        batches: batches.length,
        calls,
        inputCharsPerCall: cfg.maxInputChars,
        unitsPerBatch: kbLlmBatchLimit(cfg),
        maxOutputTokens: cfg.maxOutputTokens,
        maxCalls: cfg.maxCalls,
        overBudget: calls > cfg.maxCalls,
        // 中文按 ~1.5 字符/token 粗估；只用于让用户有量级概念，不保证精确。
        estimatedInputTokens: Math.ceil(chars / 1.5),
        estimatedOutputTokens: calls * cfg.maxOutputTokens,
    };
}

// ── 模型输出 → 分组（严格校验；坏输出只降级，不猜） ────────────────────────

/** 从模型回复里抠出 JSON（容忍代码围栏与前后废话）。 */
export function parseJsonLoose(raw) {
    const text = String(raw ?? '').trim();
    if (text === '') return null;
    const fenced = /\x60{3}(?:json)?\s*([\s\S]*?)\x60{3}/i.exec(text);
    const candidates = fenced === null ? [text] : [fenced[1], text];
    for (const candidate of candidates) {
        const start = candidate.indexOf('{');
        const end = candidate.lastIndexOf('}');
        if (start < 0 || end <= start) continue;
        try {
            return JSON.parse(candidate.slice(start, end + 1));
        } catch { /* 继续试下一个候选 */ }
    }
    return null;
}

/**
 * 校验分组：必须从 0 开始、连续、不重叠、覆盖全部小节。
 * 任何一条不满足就返回 null —— 调用方按「本箱降级」处理，绝不拿半份分组凑。
 */
export function validateGroups(parsed, unitCount) {
    if (parsed === null || typeof parsed !== 'object') return null;
    const raw = Array.isArray(parsed.groups) ? parsed.groups : null;
    if (raw === null || raw.length === 0) return null;
    const groups = [];
    for (const item of raw) {
        if (item === null || typeof item !== 'object') return null;
        const from = Number.isInteger(item.from) ? item.from : Number.parseInt(String(item.from), 10);
        const to = Number.isInteger(item.to) ? item.to : Number.parseInt(String(item.to), 10);
        if (!Number.isFinite(from) || !Number.isFinite(to) || from > to) return null;
        groups.push({
            from,
            to,
            title: asText(item.title),
            summary: asText(item.summary),
        });
    }
    groups.sort((left, right) => left.from - right.from);
    let expect = 0;
    for (const group of groups) {
        if (group.from !== expect) return null;
        expect = group.to + 1;
    }
    if (expect !== unitCount) return null;
    return groups;
}

/** 本箱降级：按确定性块大小把小节打包（不调模型、不留摘要）。 */
function fallbackGroups(units, { minChunkChars, maxChunkChars }) {
    const groups = [];
    let current = null;
    for (const unit of units) {
        const size = unit.end - unit.start;
        if (current === null
            || (current.chars >= minChunkChars && current.chars + size > maxChunkChars)) {
            current = { from: unit.index, to: unit.index, chars: 0, title: '', summary: '', fallback: true };
            groups.push(current);
        }
        current.chars += size;
        current.to = unit.index;
    }
    return groups.length === 0
        ? [{ from: 0, to: units.length - 1, title: '', summary: '', fallback: true }]
        : groups;
}

/** 把一箱分组映射成块（span 全部来自原文；摘要另存，绝不混进正文）。 */
function chunksFromGroups(source, units, groups, { startN, maxChunkChars }) {
    const chunks = [];
    let n = startN;
    for (const group of groups) {
        const sliceUnits = units.slice(group.from, group.to + 1);
        if (sliceUnits.length === 0) continue;
        // 一箱已经 <= maxInputChars；这里再守一道「单块 <= maxChunkChars」，
        // 超了就按小节继续切开（摘要跟着第一块走，其余块标 fallback）。
        let bucket = [];
        let chars = 0;
        // 一段被 maxChunkChars 再拆开时，摘要只跟第一块走，避免同一段摘要在多块里重复。
        let firstPiece = true;
        const flush = () => {
            if (bucket.length === 0) return;
            const start = bucket[0].start;
            const end = bucket[bucket.length - 1].end;
            const text = source.slice(start, end);
            const outline = bucket[0].outline;
            chunks.push({
                n,
                text,
                span: { start, end },
                outline,
                title: group.title,
                summary: firstPiece ? group.summary : '',
                fallback: group.fallback === true,
            });
            firstPiece = false;
            n += 1;
            bucket = [];
            chars = 0;
        };
        for (const unit of sliceUnits) {
            const size = unit.end - unit.start;
            if (bucket.length > 0 && chars + size > maxChunkChars) flush();
            bucket.push(unit);
            chars += size;
        }
        flush();
    }
    return chunks;
}

// ── 请求构造与执行 ─────────────────────────────────────────────────────────

const SYSTEM_PROMPT = [
    '你是文本结构分析助手。只输出 JSON，不要输出解释、不要复述原文。',
    '你的任务是把用户给出的原文小节分成若干语义连续的片段，并为每段写一个短标题和 1~3 句摘要。',
    '硬规则：',
    '1. 只能用给出的 [编号] 分组；from/to 是编号（含两端），必须从 0 开始、连续、不重叠、覆盖全部编号；',
    '2. 不得改写、补写、翻译或省略原文内容；摘要只写「这一段讲了什么」，不是原文摘抄；',
    '3. 每段目标长度 400~2500 字，太短就合并，太长就拆开；',
    '4. 只输出 {"groups":[{"from":0,"to":1,"title":"...","summary":"..."}]} 这一种形状。',
    '5. 本箱最多 {maxGroups} 个分组；编号更多时把相邻编号合并成更少的组，每组摘要不超过 120 字。',
].join('\n');

/** 把本箱允许的分组数写进系统提示（输出预算与装箱上限是同一个数）。 */
export function buildKbLlmSystemPrompt(maxGroups) {
    const cap = Math.max(1, Math.trunc(Number(maxGroups) || 1));
    return SYSTEM_PROMPT.replace('{maxGroups}', String(cap));
}

export function buildKbLlmPrompt({ units, source }) {
    const lines = ['原文小节：'];
    for (const unit of units) {
        const body = source.slice(unit.start, unit.end);
        const label = unit.heading === true ? '[标题] ' : '';
        lines.push('[' + unit.index + '] ' + label + body);
    }
    return lines.join('\n');
}

/**
 * 跑一次文档级的「切片 + 摘要」任务。可恢复：同一份 state 再调一次会从上次的
 * 箱继续；单次最多跑 config.maxCallsPerRun 箱，总调用数硬上限 config.maxCalls。
 *
 * @returns {Promise<{
 *   status: 'off'|'disabled'|'done'|'paused'|'budget'|'over-budget',
 *   chunks: object[], summaries: object[], segments: object[],
 *   calls: number, state: object, batches: number, doneBatches: number,
 *   degraded: number, errors: string[], reason: string, estimate: object
 * }>}
 */
export async function runKbLlmJob({
    text = '',
    config = {},
    generate = null,
    state = null,
    runLimit = null,
    docHash = '',
    env = process.env,
    onProgress = null,
} = {}) {
    const cfg = normalizeKbLlmConfig(config);
    const source = typeof text === 'string' ? text : '';
    const ready = kbLlmReady(cfg);
    const plan = planKbBatches(source, cfg);
    const estimate = estimateKbLlmJob(source, cfg);
    const st = state !== null && typeof state === 'object'
        ? state
        : { version: KB_LLM_PROTOCOL_VERSION, cursor: 0, chunks: [], summaries: [], segments: [], calls: 0, degraded: 0, truncated: 0, retried: 0, errors: [] };
    const base = {
        chunks: [], summaries: [], segments: [], calls: 0, state: st,
        batches: plan.batches.length, doneBatches: 0, degraded: 0, truncated: 0, retried: 0, errors: [], estimate,
    };
    if (!ready.ok) {
        return { ...base, status: cfg.mode === 'off' ? 'off' : 'disabled', doneBatches: st.cursor ?? 0, reason: ready.reason };
    }
    if (typeof generate !== 'function') {
        return { ...base, status: 'disabled', reason: '没有可用的模型调用器（generate 未注入）。' };
    }
    if (st.batches !== undefined && st.batches !== plan.batches.length) {
        // 正文换了（docHash 不同）却拿旧进度来续：拒绝续跑，避免把两次切块拼在一起。
        return { ...base, status: 'disabled', reason: '进度与当前正文不匹配（正文或配置变了），请重新开始。' };
    }
    const startCursor = Number.isFinite(st.cursor) ? Math.max(0, Math.trunc(st.cursor)) : 0;
    if (startCursor === 0 && estimate.overBudget) {
        return {
            ...base,
            status: 'over-budget',
            state: { ...st, batches: plan.batches.length },
            reason: '预估需要 ' + estimate.calls + ' 次模型调用，超过硬上限 ' + cfg.maxCalls
                + ' 次；没有发起任何调用。请调大上限、按章分次入库，或改回本地确定性切片。',
        };
    }
    const limit = numberIn(runLimit, cfg.maxCallsPerRun, 1, cfg.maxCalls);
    const chunks = Array.isArray(st.chunks) ? st.chunks : [];
    const summaries = Array.isArray(st.summaries) ? st.summaries : [];
    const segments = Array.isArray(st.segments) ? st.segments : [];
    const report = (event) => {
        if (typeof onProgress !== 'function') return;
        try { onProgress(event); } catch { /* 进度回调不能打断入库 */ }
    };
    report({
        phase: 'plan',
        doneBatches: startCursor,
        totalBatches: plan.batches.length,
        calls: Number.isFinite(st.calls) ? st.calls : 0,
        degraded: Number.isFinite(st.degraded) ? st.degraded : 0,
        truncated: Number.isFinite(st.truncated) ? st.truncated : 0,
    });
    let calls = Number.isFinite(st.calls) ? st.calls : 0;
    let degraded = Number.isFinite(st.degraded) ? st.degraded : 0;
    let truncated = Number.isFinite(st.truncated) ? st.truncated : 0;
    let retried = Number.isFinite(st.retried) ? st.retried : 0;
    const errors = Array.isArray(st.errors) ? [...st.errors] : [];
    let cursor = startCursor;
    let ran = 0;
    let stop = '';
    while (cursor < plan.batches.length) {
        if (ran >= limit) { stop = 'paused'; break; }
        if (calls >= cfg.maxCalls) { stop = 'budget'; break; }
        const batch = plan.batches[cursor];
        // 分组编号必须相对本箱：模型的 from/to 是箱内编号，块正文仍按原文 span 切。
        const units = batch.units.map((unit, index) => ({ ...unit, index }));
        const prompt = buildKbLlmPrompt({ units, source });
        let groups = null;
        let note = '';
        let result = await callWithDeadline(generate, {
            system: buildKbLlmSystemPrompt(units.length),
            prompt,
            maxTokens: cfg.maxOutputTokens,
            timeoutMs: cfg.timeoutMs,
        });
        calls += 1;
        ran += 1;
        if (!result.ok) {
            note = result.error || '模型调用失败';
        } else {
            groups = validateGroups(parseJsonLoose(result.text), units.length);
            if (groups === null) {
                note = result.truncated === true
                    ? '模型输出被截断（finish_reason=' + result.finish + '），不是完整分组'
                    : '模型输出不是可用的分组（已按确定性规则降级本批）';
            }
        }
        if (result.ok && result.truncated === true) truncated += 1;
        // 截断才值得重试：输出上限翻倍再问一次同一箱。重试也算一次调用与一轮预算，
        // 预算不够就老实降级，不能为了救一箱把总上限冲掉。
        const retryTokens = Math.min(8_000, Math.max(2_000, cfg.maxOutputTokens * 2));
        if (groups === null && result.ok && result.truncated === true
            && retryTokens > cfg.maxOutputTokens && calls < cfg.maxCalls && ran < limit) {
            retried += 1;
            result = await callWithDeadline(generate, {
                system: buildKbLlmSystemPrompt(units.length),
                prompt,
                maxTokens: retryTokens,
                timeoutMs: cfg.timeoutMs,
            });
            calls += 1;
            ran += 1;
            if (!result.ok) {
                note = '重试失败：' + (result.error || '模型调用失败');
            } else {
                groups = validateGroups(parseJsonLoose(result.text), units.length);
                if (result.truncated === true) truncated += 1;
                if (groups === null) {
                    note = result.truncated === true
                        ? '模型输出在 ' + retryTokens + ' token 下仍被截断'
                        : '模型输出不是可用的分组（重试后仍无效）';
                }
            }
        }
        // 全局块号：跨箱连续（n 在整篇里唯一）。
        const nextN = chunks.length + 1;
        if (groups === null) {
            groups = fallbackGroups(units, cfg);
            degraded += 1;
            if (note !== '' && errors.length < 5) errors.push('批次 ' + cursor + '：' + note);
        }
        if (cfg.chunking) {
            const made = chunksFromGroups(source, units, groups, { startN: nextN, maxChunkChars: cfg.maxChunkChars });
            for (const chunk of made) {
                chunks.push(chunk);
                if (cfg.summary && chunk.summary !== '') {
                    summaries.push({
                        id: 'kbllm:' + (docHash || 'doc') + ':' + chunk.n,
                        title: chunk.title || '片段摘要',
                        text: chunk.summary,
                        n: chunk.n,
                        chunkIds: [chunkIdOf(docHash, chunk.n)],
                    });
                }
            }
        } else if (cfg.summary) {
            // 只开摘要：不重切正文，只把每段摘要记下来，由调用方按 span 映射到已有块。
            for (const group of groups) {
                const first = units[group.from];
                const last = units[group.to];
                segments.push({
                    span: { start: first.start, end: last.end },
                    title: group.title,
                    summary: group.summary,
                    n: group.from + 1,
                });
            }
        }
        cursor += 1;
        report({ phase: 'model', doneBatches: cursor, totalBatches: plan.batches.length, calls, degraded, truncated });
    }
    const status = cursor >= plan.batches.length ? 'done' : (stop === '' ? 'paused' : stop);
    const nextState = {
        version: KB_LLM_PROTOCOL_VERSION,
        batches: plan.batches.length,
        cursor,
        chunks,
        summaries,
        segments,
        calls,
        degraded,
        truncated,
        retried,
        errors,
        done: status === 'done',
    };
    const reason = status === 'done'
        ? ''
        : status === 'budget'
            ? '已达到总调用上限 ' + cfg.maxCalls + ' 次，任务停在 ' + cursor + '/' + plan.batches.length + ' 批；已完成的块可以先用，续跑从下一批开始。'
            : '本轮跑完 ' + cursor + '/' + plan.batches.length + ' 批（单轮上限 ' + limit + ' 次），可续跑。';
    return {
        ...base,
        status,
        chunks,
        summaries,
        segments,
        calls,
        degraded,
        truncated,
        retried,
        errors,
        state: nextState,
        doneBatches: cursor,
        reason,
    };
}

/** 单次调用加壳：生成器自己忽略 signal 时也能被兜住（不让一次调用挂死整轮）。 */
async function callWithDeadline(generate, { system, prompt, maxTokens, timeoutMs }) {
    const controller = new AbortController();
    let settled = false;
    const timer = setTimeout(() => controller.abort(new Error('timeout')), timeoutMs);
    try {
        const raced = await Promise.race([
            Promise.resolve().then(() => generate({ system, prompt, maxTokens, signal: controller.signal })),
            new Promise((resolve) => {
                controller.signal.addEventListener('abort', () => {
                    resolve({ ok: false, text: '', error: '模型调用超时（' + timeoutMs + 'ms）或已取消。' });
                }, { once: true });
            }),
        ]);
        settled = true;
        return raced !== null && typeof raced === 'object' ? raced : { ok: false, text: '', error: '模型调用返回了意外的空结果。' };
    } catch (error) {
        settled = true;
        return { ok: false, text: '', error: '模型调用抛错：' + String(error?.message ?? error) };
    } finally {
        clearTimeout(timer);
        // 超时后主动 abort，让守规矩的适配器尽快收尾（已 settled 的请求不受影响）。
        if (!settled) {
            try { controller.abort(new Error('timeout')); } catch { /* 已 abort */ }
        }
    }
}

// ── 模型适配器（两个真实协议，不做假接口） ─────────────────────────────────

/**
 * OpenAI 兼容适配器：POST {baseURL}/chat/completions。
 * 这是**协议**而不是厂商绑定：任何兼容端点都可用（自建网关 / vLLM / Ollama 网关等）。
 * 不支持 chat/completions 的厂商只能走 mode='host'，回执里如实说明。
 */
export function createOpenAiGenerator({ config = {}, fetchImpl = globalThis.fetch, env = process.env } = {}) {
    const cfg = normalizeKbLlmConfig(config);
    return async function generate({ system, prompt, maxTokens = cfg.maxOutputTokens, signal = null }) {
        if (typeof fetchImpl !== 'function') {
            return { ok: false, text: '', error: '当前运行环境没有 fetch，OpenAI 兼容通道不可用。' };
        }
        if (cfg.baseURL === '' || cfg.model === '') {
            return { ok: false, text: '', error: 'OpenAI 兼容通道缺少 baseURL 或 model。' };
        }
        const key = kbLlmKeyOf(cfg, env);
        const headers = { 'content-type': 'application/json' };
        if (key !== '') headers.authorization = 'Bearer ' + key;
        const timeout = timeoutSignal(cfg.timeoutMs, signal);
        try {
            const response = await fetchImpl(cfg.baseURL + '/chat/completions', {
                method: 'POST',
                headers,
                body: JSON.stringify({
                    model: cfg.model,
                    temperature: 0,
                    max_tokens: maxTokens,
                    messages: [
                        { role: 'system', content: system },
                        { role: 'user', content: prompt },
                    ],
                }),
                signal: timeout.signal,
            });
            const raw = await response.text();
            if (!response.ok) {
                return { ok: false, text: '', error: '模型端点返回 HTTP ' + response.status + '：' + raw.slice(0, 200) };
            }
            let parsed = null;
            try { parsed = JSON.parse(raw); } catch { return { ok: false, text: '', error: '模型端点返回的不是 JSON。' }; }
            const choice = Array.isArray(parsed?.choices) ? parsed.choices[0] : null;
            const text = asText(choice?.message?.content);
            if (text === '') return { ok: false, text: '', error: '模型端点返回了空正文。' };
            const finish = asText(choice?.finish_reason) || 'stop';
            return {
                ok: true,
                text,
                model: asText(parsed?.model) || cfg.model,
                finish,
                truncated: finish === 'length',
                error: '',
            };
        } catch (error) {
            const timedOut = timeout.timedOut === true || error?.name === 'AbortError' || error?.name === 'TimeoutError';
            return { ok: false, text: '', error: timedOut ? '模型调用超时（' + cfg.timeoutMs + 'ms）。' : '模型调用失败：' + String(error?.message ?? error) };
        } finally {
            timeout.dispose();
        }
    };
}

/**
 * 宿主适配器：走 ctx.llm（dsh-llm 的 LlmRuntime）已注册的 provider / model。
 *
 * 这里**不 import 宿主的裸包名**（插件运行时解析不到），只用结构化对象 +
 * 文档化的 prepareCall/stream 契约；宿主不在时明确返回不可用，而不是造一个假的。
 */
export function createHostGenerator({ llm = null, config = {}, purpose = 'office-kb', sessionId = '' } = {}) {
    const cfg = normalizeKbLlmConfig(config);
    return async function generate({ system, prompt, maxTokens = cfg.maxOutputTokens, signal = null }) {
        if (llm === null || typeof llm.prepareCall !== 'function') {
            return { ok: false, text: '', error: '宿主没有可用的 llm 服务（ctx.llm 不可达）。' };
        }
        if (cfg.provider === '' || cfg.model === '') {
            return { ok: false, text: '', error: '宿主通道缺少 provider 或 model。' };
        }
        const timeout = timeoutSignal(cfg.timeoutMs, signal);
        try {
            const call = await llm.prepareCall({ provider: cfg.provider, model: cfg.model, maxTokens }, timeout.signal);
            const options = {
                ...(call?.config ?? {}),
                provider: cfg.provider,
                model: cfg.model,
                maxTokens,
                messages: [{ role: 'user', content: [{ type: 'text', text: prompt }] }],
                system,
                purpose,
                signal: timeout.signal,
            };
            if (sessionId !== '') options.sessionId = sessionId;
            let text = '';
            let finish = 'stop';
            for await (const chunk of call.stream(options)) {
                if (chunk?.type === 'text-delta') text += String(chunk.text ?? '');
                else if (chunk?.type === 'finish') {
                    finish = String(chunk.reason?.kind ?? 'stop');
                    if (finish === 'error' || finish === 'aborted') {
                        const message = chunk.reason?.failure?.message ?? '宿主模型调用被中断。';
                        return { ok: false, text: '', error: message };
                    }
                }
            }
            if (text.trim() === '') return { ok: false, text: '', error: '宿主模型没有返回正文。' };
            return { ok: true, text, model: cfg.model, finish, truncated: finish === 'max-tokens', error: '' };
        } catch (error) {
            const timedOut = timeout.timedOut === true || error?.name === 'AbortError' || error?.name === 'TimeoutError';
            return { ok: false, text: '', error: timedOut ? '宿主模型调用超时（' + cfg.timeoutMs + 'ms）。' : '宿主模型调用失败：' + String(error?.message ?? error) };
        } finally {
            timeout.dispose();
        }
    };
}

/** 合成超时信号：外部取消与超时任一触发即可。 */
function timeoutSignal(ms, external = null) {
    let timer = null;
    let timedOut = false;
    const controller = new AbortController();
    timer = setTimeout(() => { timedOut = true; controller.abort(new Error('timeout')); }, ms);
    if (external !== null && typeof external.addEventListener === 'function') {
        const forward = () => controller.abort(external.reason ?? new Error('aborted'));
        if (external.aborted) forward();
        else external.addEventListener('abort', forward, { once: true });
    }
    return {
        signal: controller.signal,
        get timedOut() { return timedOut; },
        dispose() { if (timer !== null) clearTimeout(timer); },
    };
}

// ── 任务进度落盘（可续跑） ─────────────────────────────────────────────────

/**
 * 一次任务的稳定键：正文 + 配置 profile（换模型或换装箱上限就是另一份任务）。
 * 用它做文件名，避免把两次不同切法的进度拼在一起。
 */
export function kbLlmJobKeyOf(text, config = {}) {
    const cfg = normalizeKbLlmConfig(config);
    const digest = createHash('sha1')
        .update(kbLlmProfile(cfg) + '\0' + (typeof text === 'string' ? text : ''))
        .digest('hex');
    return 'kbllm-' + digest.slice(0, 16);
}

function kbLlmJobFile(dir, key) {
    if (!/^kbllm-[0-9a-f]{8,40}$/.test(String(key))) throw new Error('非法的知识库模型任务键：' + key);
    return path.join(dir, '.llm-jobs', key + '.json');
}

/** 读回上次的进度；没有或坏了都返回 null（坏进度绝不拿来续跑）。 */
export async function readKbLlmJob(dir, key) {
    try {
        const raw = await readFile(kbLlmJobFile(dir, key), 'utf8');
        const parsed = JSON.parse(raw);
        return parsed !== null && typeof parsed === 'object' && parsed.state ? parsed : null;
    } catch {
        return null;
    }
}

/** 落盘进度（只写我们自己的 .llm-jobs 子目录，不碰 manifest/chunks）。 */
export async function writeKbLlmJob(dir, key, payload) {
    const file = kbLlmJobFile(dir, key);
    await mkdir(path.dirname(file), { recursive: true });
    await writeFile(file, JSON.stringify(payload), 'utf8');
    return file;
}

/** 任务完成后清掉进度文件；清理失败不影响已提交的块。 */
export async function dropKbLlmJob(dir, key) {
    try {
        await rm(kbLlmJobFile(dir, key), { force: true });
    } catch { /* 已经没了 */ }
}

/**
 * 只开摘要时，把模型给的 segments（原文区间 + 摘要）映射到**已有的**确定性块上。
 *
 * 为什么需要：只开摘要不改切块，所以不能拿摘要去重切正文；但检索要能顺着摘要找到
 * 原文块。这里按区间重叠给每个块选一段摘要（重叠最大者），一个块只归一段，块不会丢。
 * 返回 [{ index, span, title, text, chunkIndexes }]，chunkIndexes 是传入 chunks 数组的下标，
 * 由调用方换成真实块 id（块 id 的生成属于 kb.js，不在这一层）。
 */
export function mapSummariesToChunks({ segments = [], chunks = [] } = {}) {
    const segs = (Array.isArray(segments) ? segments : [])
        .filter((seg) => seg !== null && typeof seg === 'object' && seg.span !== null && seg.span !== undefined
            && Number.isFinite(seg.span.start) && Number.isFinite(seg.span.end) && seg.span.end > seg.span.start)
        .map((seg) => ({ span: { start: seg.span.start, end: seg.span.end }, title: asText(seg.title), summary: asText(seg.summary) }));
    const list = (Array.isArray(chunks) ? chunks : [])
        .filter((chunk) => chunk !== null && typeof chunk === 'object' && chunk.span !== null && chunk.span !== undefined
            && Number.isFinite(chunk.span.start) && Number.isFinite(chunk.span.end));
    if (segs.length === 0 || list.length === 0) return [];
    const buckets = segs.map((seg) => ({ ...seg, chunkIndexes: [] }));
    for (let at = 0; at < list.length; at += 1) {
        const chunk = list[at];
        let best = -1;
        let bestOverlap = 0;
        for (let si = 0; si < segs.length; si += 1) {
            const overlap = Math.min(chunk.span.end, segs[si].span.end) - Math.max(chunk.span.start, segs[si].span.start);
            if (overlap > bestOverlap) { bestOverlap = overlap; best = si; }
        }
        if (best >= 0) buckets[best].chunkIndexes.push(at);
    }
    return buckets
        .filter((bucket) => bucket.chunkIndexes.length > 0)
        .map((bucket, index) => ({ index, span: bucket.span, title: bucket.title, text: bucket.summary, chunkIndexes: bucket.chunkIndexes }));
}

/**
 * 列出宿主里**已配置可用**的 provider / model，给设置页的模型选择器用。
 * 形状随宿主版本可能有细节差异，这里只做保守投影（id + 显示名），不猜字段。
 */
export async function listHostModels(llm) {
    if (llm === null || typeof llm !== 'object' || typeof llm.listProviders !== 'function') {
        return { ok: false, providers: [], error: '宿主没有可用的 llm 服务。' };
    }
    try {
        const providers = await llm.listProviders();
        const list = Array.isArray(providers) ? providers : [];
        const out = [];
        for (const provider of list) {
            const id = typeof provider === 'string' ? provider : asText(provider?.id ?? provider?.provider ?? provider?.name);
            if (id === '') continue;
            let models = [];
            if (typeof llm.listModels === 'function') {
                try {
                    const raw = await llm.listModels(id);
                    const items = Array.isArray(raw) ? raw : [];
                    models = items.map((item) => {
                        if (typeof item === 'string') return { id: item, label: item };
                        const modelId = asText(item?.id ?? item?.model ?? item?.name);
                        return modelId === '' ? null : { id: modelId, label: asText(item?.name ?? item?.label ?? item?.displayName ?? item?.title) || modelId };
                    }).filter((item) => item !== null);
                } catch { models = []; }
            }
            out.push({ provider: id, models });
        }
        return { ok: true, providers: out, error: '' };
    } catch (error) {
        return { ok: false, providers: [], error: '读取宿主模型清单失败：' + String(error?.message ?? error) };
    }
}
