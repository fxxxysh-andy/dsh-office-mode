/**
 * 知识库「可选大模型增强」的测试（第九十二轮）。
 *
 * 这一份只测 kb-llm.js 自己的契约，全部用假的 generate / fetch 或假宿主 llm：
 *   - 默认不调用：mode='off' 时零请求；选了通道但没选用途也不调用；
 *   - 模型可选：OpenAI 兼容端点把 model 原样发出；宿主通道把 provider/model 原样传给 prepareCall；
 *   - 失败可观察：HTTP 错、非 JSON、超时、宿主不可用，全部降级但把原因写进回执；
 *   - 切片不丢字：块正文 === source.slice(span)，且全部非空白字符恰好被覆盖一次；
 *   - 摘要不替代原文：摘要另存，正文永远是原文切片；
 *   - 长文有界可续：60k 字符分多箱，单次只跑上限箱数，可续跑，超预算明确暂停；
 *   - 千万字量级：预估出上千次调用，默认预算下直接 over-budget（零调用）。
 *
 * 跑法：node test/kb-llm.mjs
 */
import assert from 'node:assert/strict';
import { rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { chunkOverlapFor } from '../src/kb.js';
import {
    DEFAULT_KB_LLM,
    buildKbLlmPrompt,
    buildKbLlmSystemPrompt,
    createHostGenerator,
    createOpenAiGenerator,
    describeKbLlm,
    dropKbLlmJob,
    estimateKbLlmJob,
    kbLlmBatchLimit,
    kbLlmJobKeyOf,
    kbLlmProfile,
    kbLlmReady,
    listHostModels,
    mapSummariesToChunks,
    normalizeKbLlmConfig,
    parseJsonLoose,
    planKbBatches,
    readKbLlmJob,
    runKbLlmJob,
    writeKbLlmJob,
    unitsOf,
    validateGroups,
} from '../src/kb-llm.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 一段带 Markdown 结构的测试正文（3 章、每章 8 段、每段约 90 字）。 */
function sampleDoc(chapters = 3, paragraphs = 8) {
    const parts = [];
    for (let c = 1; c <= chapters; c += 1) {
        parts.push('# 第' + c + '章 标题' + c);
        parts.push('');
        for (let p = 1; p <= paragraphs; p += 1) {
            parts.push('这是第' + c + '章第' + p + '段。' + '剧情推进到这里，人物做了若干选择。'.repeat(3));
            parts.push('');
        }
    }
    return parts.join('\n');
}

/** 断言：每块正文都等于源文切片，且全部非空白字符恰好被覆盖一次。 */
function assertCoverage(source, chunks) {
    assert.ok(chunks.length > 0, '至少要切出一块');
    const covered = new Uint8Array(source.length);
    let lastEnd = -1;
    for (const chunk of chunks) {
        assert.equal(chunk.text, source.slice(chunk.span.start, chunk.span.end), '块正文必须等于源文切片');
        assert.ok(chunk.span.start >= 0 && chunk.span.end <= source.length, 'span 必须在源文范围内');
        assert.ok(chunk.span.start < chunk.span.end, 'span 不能为空');
        assert.ok(chunk.span.start >= lastEnd, '块区间必须单调不重叠');
        lastEnd = chunk.span.end;
        for (let i = chunk.span.start; i < chunk.span.end; i += 1) covered[i] += 1;
    }
    for (let i = 0; i < source.length; i += 1) {
        if (/\s/.test(source[i])) continue;
        assert.equal(covered[i], 1, '字符 ' + i + ' 的覆盖次数应为 1，实际 ' + covered[i]);
    }
}

/** 从提示词里数出这一箱有几个小节。 */
function unitsInPrompt(prompt) {
    const found = [...prompt.matchAll(/^\[(\d+)\] /gm)].map((match) => Number(match[1]));
    return found.length === 0 ? 0 : Math.max(...found) + 1;
}

/**
 * 假模型：把这一箱小节按 2~3 段一组切开，返回合法 JSON（边界永远落在小节上）。
 * 同时记录每次 prompt 长度，供「单箱不超过 maxInputChars」断言。
 */
function makeFakeGenerator(calls = []) {
    return async ({ prompt }) => {
        calls.push({ prompt });
        const count = unitsInPrompt(prompt);
        const size = Math.max(1, Math.ceil(count / 3));
        const groups = [];
        for (let from = 0; from < count; from += size) {
            groups.push({
                from,
                to: Math.min(count - 1, from + size - 1),
                title: '片段' + (groups.length + 1),
                summary: '这是摘要，不是正文。' + (groups.length + 1),
            });
        }
        return { ok: true, text: JSON.stringify({ groups }), model: 'fake', finish: 'stop', error: '' };
    };
}

const BASE_CONFIG = {
    mode: 'openai',
    chunking: true,
    summary: true,
    baseURL: 'http://127.0.0.1:1/v1',
    model: 'fake-model',
};

// ── 默认与开关 ─────────────────────────────────────────────────────────────

await check('默认关闭：mode=off 时零调用、状态 off', async () => {
    let called = 0;
    const result = await runKbLlmJob({
        text: sampleDoc(),
        config: {},
        generate: async () => { called += 1; return { ok: true, text: '{}' }; },
    });
    assert.equal(result.status, 'off');
    assert.equal(result.calls, 0);
    assert.equal(called, 0);
    assert.ok(result.reason.includes('默认关闭'));
});

await check('选了通道但没选用途：disabled，零调用', async () => {
    let called = 0;
    const result = await runKbLlmJob({
        text: sampleDoc(),
        config: { mode: 'openai', baseURL: 'http://x/v1', model: 'm' },
        generate: async () => { called += 1; return { ok: true, text: '{}' }; },
    });
    assert.equal(result.status, 'disabled');
    assert.equal(called, 0);
    assert.ok(result.reason.includes('都没勾选'));
    assert.equal(kbLlmReady({ mode: 'off' }).ok, false);
    assert.equal(kbLlmReady({ mode: 'openai', baseURL: 'http://x/v1', model: 'm', summary: true }).ok, true);
});

// ── 原文保真 ───────────────────────────────────────────────────────────────

await check('切片不丢字：块正文等于源文切片且非空白字符全覆盖', async () => {
    const source = sampleDoc();
    const calls = [];
    const result = await runKbLlmJob({ text: source, config: BASE_CONFIG, generate: makeFakeGenerator(calls) });
    assert.equal(result.status, 'done');
    assertCoverage(source, result.chunks);
    assert.ok(result.chunks.length >= 3, '应切出多块');
    assert.ok(calls.length >= 1);
});

await check('摘要不替代原文：摘要另存、正文仍是原文切片', async () => {
    const source = sampleDoc();
    const result = await runKbLlmJob({ text: source, config: BASE_CONFIG, generate: makeFakeGenerator() });
    assert.equal(result.status, 'done');
    assert.equal(result.summaries.length, result.chunks.length);
    for (const summary of result.summaries) {
        assert.ok(summary.text.includes('这是摘要'), '摘要正文应来自模型');
        const chunk = result.chunks.find((item) => item.n === summary.n);
        assert.ok(chunk !== undefined);
        assert.ok(!chunk.text.includes('这是摘要'), '摘要不得混进原文块');
        assert.equal(chunk.summary, summary.text);
        assert.deepEqual(summary.chunkIds, ['kb::' + summary.n]);
    }
});

await check('只开摘要：不重切正文，只产出按 span 的 segments', async () => {
    const source = sampleDoc();
    const result = await runKbLlmJob({
        text: source,
        config: { ...BASE_CONFIG, chunking: false, summary: true },
        generate: makeFakeGenerator(),
    });
    assert.equal(result.status, 'done');
    assert.equal(result.chunks.length, 0);
    assert.ok(result.segments.length > 0);
    for (const segment of result.segments) {
        assert.ok(segment.span.end > segment.span.start);
        assert.ok(source.slice(segment.span.start, segment.span.end).length > 0);
        assert.ok(segment.summary.includes('这是摘要'));
    }
});

// ── 模型可选与适配器协议 ───────────────────────────────────────────────────

await check('OpenAI 兼容适配器：model/baseURL/Key 都按配置发出', async () => {
    let seen = null;
    const fetchImpl = async (url, init) => {
        seen = { url, body: JSON.parse(init.body), headers: init.headers };
        return new Response(JSON.stringify({
            model: 'echo-model',
            choices: [{ message: { content: '{"groups":[]}' }, finish_reason: 'stop' }],
        }), { status: 200, headers: { 'content-type': 'application/json' } });
    };
    const generate = createOpenAiGenerator({
        config: { mode: 'openai', baseURL: 'http://gw.local/v1', model: 'my-custom-model', apiKey: 'sk-test' },
        fetchImpl,
    });
    const result = await generate({ system: 'sys', prompt: 'hello', maxTokens: 64 });
    assert.equal(result.ok, true);
    assert.equal(seen.url, 'http://gw.local/v1/chat/completions');
    assert.equal(seen.body.model, 'my-custom-model');
    assert.equal(seen.body.max_tokens, 64);
    assert.equal(seen.body.messages[0].role, 'system');
    assert.equal(seen.headers.authorization, 'Bearer sk-test');
});

await check('宿主适配器：provider/model 原样进 prepareCall，文本从流里拼出', async () => {
    const seen = [];
    const llm = {
        async prepareCall(controls) {
            seen.push(controls);
            return {
                config: { provider: controls.provider, model: controls.model, maxTokens: controls.maxTokens },
                stream: async function* () {
                    yield { type: 'text-delta', text: '{"groups":' };
                    yield { type: 'text-delta', text: '[]}' };
                    yield { type: 'finish', reason: { kind: 'stop' } };
                },
            };
        },
    };
    const generate = createHostGenerator({ llm, config: { mode: 'host', provider: 'deepseek', model: 'deepseek-chat', chunking: true } });
    const result = await generate({ system: 'sys', prompt: 'hi', maxTokens: 32, signal: new AbortController().signal });
    assert.equal(result.ok, true);
    assert.equal(result.text, '{"groups":[]}');
    assert.deepEqual(seen[0], { provider: 'deepseek', model: 'deepseek-chat', maxTokens: 32 });
});

await check('宿主不可用：明确不可用，不抛错', async () => {
    const generate = createHostGenerator({ llm: null, config: { mode: 'host', provider: 'p', model: 'm', chunking: true } });
    const result = await generate({ system: 's', prompt: 'p' });
    assert.equal(result.ok, false);
    assert.ok(result.error.includes('ctx.llm'));
    const models = await listHostModels(null);
    assert.equal(models.ok, false);
});

await check('宿主模型清单：投影成 provider/models', async () => {
    const llm = {
        async listProviders() { return ['deepseek', { id: 'openai' }]; },
        async listModels(provider) {
            return provider === 'deepseek'
                ? [{ id: 'deepseek-chat', label: 'DeepSeek Chat' }, 'deepseek-reasoner']
                : [{ id: 'gpt-x' }];
        },
    };
    const listed = await listHostModels(llm);
    assert.equal(listed.ok, true);
    assert.equal(listed.providers.length, 2);
    assert.deepEqual(listed.providers[0].models[0], { id: 'deepseek-chat', label: 'DeepSeek Chat' });
    assert.deepEqual(listed.providers[0].models[1], { id: 'deepseek-reasoner', label: 'deepseek-reasoner' });
});

// ── 失败与降级（都可观察） ─────────────────────────────────────────────────

await check('协议失败：HTTP 500 → 降级本批，正文仍完整', async () => {
    const source = sampleDoc();
    const fetchImpl = async () => new Response('boom', { status: 500 });
    const generate = createOpenAiGenerator({ config: BASE_CONFIG, fetchImpl });
    const result = await runKbLlmJob({ text: source, config: BASE_CONFIG, generate });
    assert.equal(result.status, 'done');
    assert.ok(result.degraded >= 1);
    assert.ok(result.errors[0].includes('HTTP 500'));
    assertCoverage(source, result.chunks);
    assert.equal(result.summaries.length, 0);
});

await check('非 JSON 响应：降级且原因可读', async () => {
    const fetchImpl = async () => new Response('<html>not json</html>', { status: 200 });
    const generate = createOpenAiGenerator({ config: BASE_CONFIG, fetchImpl });
    const result = await runKbLlmJob({ text: sampleDoc(), config: BASE_CONFIG, generate });
    assert.ok(result.degraded >= 1);
    assert.ok(result.errors[0].includes('不是 JSON'));
});

await check('超时：适配器按 timeoutMs 收口，任务降级不挂死', async () => {
    const fetchImpl = async (url, init) => new Promise((resolve, reject) => {
        init.signal.addEventListener('abort', () => {
            const error = new Error('aborted');
            error.name = 'AbortError';
            reject(error);
        }, { once: true });
    });
    const generate = createOpenAiGenerator({ config: { ...BASE_CONFIG, timeoutMs: 60 }, fetchImpl });
    const result = await runKbLlmJob({ text: sampleDoc(), config: { ...BASE_CONFIG, timeoutMs: 60 }, generate });
    assert.ok(result.degraded >= 1);
    assert.ok(result.errors.some((line) => line.includes('超时')), result.errors.join(' | '));
});

await check('生成器忽略 signal 也会被硬超时兜住', async () => {
    const result = await runKbLlmJob({
        text: sampleDoc(),
        config: { ...BASE_CONFIG, timeoutMs: 80, maxCallsPerRun: 2 },
        generate: () => new Promise(() => { /* 永不返回 */ }),
    });
    assert.ok(result.degraded >= 1);
    assert.ok(result.errors.some((line) => line.includes('超时')));
});

await check('非法分组：缺号/乱序按本批降级，不拼半份分组', async () => {
    const generate = async ({ prompt }) => {
        const count = unitsInPrompt(prompt);
        return { ok: true, text: JSON.stringify({ groups: [{ from: 0, to: Math.max(0, count - 2), title: 'x', summary: 'y' }] }) };
    };
    const source = sampleDoc();
    const result = await runKbLlmJob({ text: source, config: BASE_CONFIG, generate });
    assert.ok(result.degraded >= 1);
    assertCoverage(source, result.chunks);
});

// ── 长文：分批、可续、预算 ─────────────────────────────────────────────────

/** 构造约 targetChars 字符的多章正文。 */
function longDoc(targetChars) {
    const parts = [];
    let size = 0;
    let chapter = 0;
    while (size < targetChars) {
        chapter += 1;
        const heading = '# 第' + chapter + '章';
        parts.push(heading);
        size += heading.length;
        for (let p = 0; p < 10 && size < targetChars; p += 1) {
            const line = '第' + chapter + '章第' + (p + 1) + '段：' + '人物在这条线索上继续推进，留下伏笔与证据。'.repeat(4);
            parts.push(line);
            size += line.length;
        }
        parts.push('');
    }
    return parts.join('\n');
}

await check('长文分批有界：单箱不超过 maxInputChars，可暂停可续跑', async () => {
    const source = longDoc(60_000);
    const calls = [];
    const config = { ...BASE_CONFIG, maxInputChars: 4_000, maxCallsPerRun: 4 };
    const estimate = estimateKbLlmJob(source, config);
    assert.ok(estimate.batches >= 10, '60k 字符应分成至少 10 箱，实际 ' + estimate.batches);

    const first = await runKbLlmJob({ text: source, config, generate: makeFakeGenerator(calls) });
    assert.equal(first.status, 'paused');
    assert.equal(first.calls, 4);
    assert.equal(first.doneBatches, 4);
    for (const call of calls) assert.ok(call.prompt.length <= config.maxInputChars + 40, '单个 prompt 不得超过装箱上限：' + call.prompt.length);

    let state = first.state;
    let guard = 0;
    let last = first;
    while (last.status === 'paused' && guard < 100) {
        guard += 1;
        last = await runKbLlmJob({ text: source, config, generate: makeFakeGenerator(calls), state });
        state = last.state;
    }
    assert.equal(last.status, 'done');
    assert.equal(last.calls, estimate.batches);
    assert.equal(calls.length, estimate.batches);
    assertCoverage(source, last.chunks);
});

await check('超预算：默认预算下 60k 字符直接 over-budget 且零调用', async () => {
    const source = longDoc(60_000);
    let called = 0;
    const result = await runKbLlmJob({
        text: source,
        config: { ...BASE_CONFIG, maxInputChars: 4_000, maxCalls: 3 },
        generate: async () => { called += 1; return { ok: true, text: '{}' }; },
    });
    assert.equal(result.status, 'over-budget');
    assert.equal(result.calls, 0);
    assert.equal(called, 0);
    assert.ok(result.reason.includes('超过硬上限'));
});

await check('跑到一半到顶：budget 暂停（可续），不报成功', async () => {
    const source = longDoc(60_000);
    const calls = [];
    const config = { ...BASE_CONFIG, maxInputChars: 4_000, maxCallsPerRun: 3, maxCalls: 100 };
    const estimate = estimateKbLlmJob(source, config);
    const paused = await runKbLlmJob({ text: source, config, generate: makeFakeGenerator(calls) });
    assert.equal(paused.status, 'paused');
    assert.equal(paused.doneBatches, 3);

    // 用户把总上限调低后，续跑会在预算处停住（此时才叫 budget，不是静默成功）。
    const capped = await runKbLlmJob({
        text: source,
        config: { ...config, maxCalls: 5 },
        generate: makeFakeGenerator(calls),
        state: paused.state,
    });
    assert.equal(capped.status, 'budget');
    assert.equal(capped.doneBatches, 5);
    assert.ok(capped.reason.includes('停在 5/'));

    // 再把预算调回来，必须从第 5 批继续，而不是重头再来（单轮仍受 maxCallsPerRun 限制）。
    let resumed = await runKbLlmJob({ text: source, config, generate: makeFakeGenerator(calls), state: capped.state });
    assert.equal(resumed.doneBatches, 8, '续跑应从第 5 批继续，再跑本轮上限 3 批');
    assert.equal(resumed.status, 'paused');
    let guard = 0;
    while (resumed.status === 'paused' && guard < 50) {
        guard += 1;
        resumed = await runKbLlmJob({ text: source, config, generate: makeFakeGenerator(calls), state: resumed.state });
    }
    assert.equal(resumed.status, 'done');
    assert.equal(resumed.doneBatches, estimate.batches);
    assert.equal(calls.length, estimate.batches);
});

await check('千万字量级：预估上千次调用，默认预算直接 over-budget', async () => {
    const source = '字'.repeat(10_000_000);
    const estimate = estimateKbLlmJob(source, DEFAULT_KB_LLM);
    assert.ok(estimate.calls > 1_500, '千万字应预估出上千次调用，实际 ' + estimate.calls);
    assert.equal(estimate.overBudget, true);
    let called = 0;
    const result = await runKbLlmJob({
        text: source,
        config: DEFAULT_KB_LLM,
        generate: async () => { called += 1; return { ok: true, text: '{}' }; },
    });
    assert.equal(result.status, 'off', '默认配置仍是 off');
    assert.equal(called, 0);
});

// ── 纯函数零件 ─────────────────────────────────────────────────────────────

await check('分组校验：缺口 / 重叠 / 越界一律拒绝', async () => {
    assert.equal(validateGroups({ groups: [{ from: 0, to: 1 }, { from: 2, to: 3 }] }, 4) !== null, true);
    assert.equal(validateGroups({ groups: [{ from: 0, to: 1 }, { from: 3, to: 3 }] }, 4), null);
    assert.equal(validateGroups({ groups: [{ from: 0, to: 2 }, { from: 1, to: 3 }] }, 4), null);
    assert.equal(validateGroups({ groups: [{ from: 1, to: 3 }] }, 4), null);
    assert.equal(validateGroups({ groups: [] }, 4), null);
    assert.equal(validateGroups({ groups: [{ from: 0, to: 3 }] }, 4)[0].title, '');
});

await check('JSON 容错：围栏 / 前后废话都能抠出来', async () => {
    assert.deepEqual(parseJsonLoose('好的，结果如下：\n{"groups":[{"from":0,"to":1}]}\n完毕'), { groups: [{ from: 0, to: 1 }] });
    const fenced = '\x60\x60\x60json\n{"groups":[]}\n\x60\x60\x60';
    assert.deepEqual(parseJsonLoose(fenced), { groups: [] });
    assert.equal(parseJsonLoose('完全不是 JSON'), null);
});

await check('装箱与小节：确定性、覆盖全文', async () => {
    const source = sampleDoc();
    const a = planKbBatches(source, { maxInputChars: 2_000 });
    const b = planKbBatches(source, { maxInputChars: 2_000 });
    assert.deepEqual(a.batches.map((batch) => [batch.from, batch.to]), b.batches.map((batch) => [batch.from, batch.to]));
    assert.equal(a.batches[0].from, 0);
    assert.equal(a.batches[a.batches.length - 1].to, a.units.length - 1);
    for (const batch of a.batches) assert.ok(batch.chars <= 2_000 + 40);
    const units = unitsOf(source);
    for (const unit of units) assert.ok(source.slice(unit.start, unit.end).length > 0);
});

await check('配置归一化：坏值落默认，不抛；回执不带 Key', async () => {
    const cfg = normalizeKbLlmConfig({ mode: 'nope', timeoutMs: -5, maxInputChars: 'x', apiKey: 'sk-secret' });
    assert.equal(cfg.mode, 'off');
    assert.equal(cfg.timeoutMs, DEFAULT_KB_LLM.timeoutMs);
    assert.equal(cfg.maxInputChars, DEFAULT_KB_LLM.maxInputChars);
    const receipt = JSON.stringify(describeKbLlm({ mode: 'openai', model: 'm', baseURL: 'http://x', apiKey: 'sk-secret' }));
    assert.ok(!receipt.includes('sk-secret'));
    assert.notEqual(kbLlmProfile({ ...BASE_CONFIG, model: 'a' }), kbLlmProfile({ ...BASE_CONFIG, model: 'b' }));
    // 第 2 版（第一百零一轮）：块起点带重叠，重叠量也进 profile —— 同内容但重叠不同
    // 必须算成不同档，否则「补上重叠」的升级会被判成内容没变、旧的非重叠块永远留着。
    assert.ok(kbLlmProfile(BASE_CONFIG).startsWith('llm:3:'));
    // 重叠量按**归一化后**的 maxChunkChars 派生：BASE_CONFIG 没显式给这一项时，
    // kbLlmProfile 用的是 DEFAULT_KB_LLM 的值，直接读原始对象会断言到错误的档。
    assert.ok(kbLlmProfile(BASE_CONFIG).includes(':' + chunkOverlapFor(normalizeKbLlmConfig(BASE_CONFIG).maxChunkChars)),
        'profile 必须带重叠量');
    // 四个装箱参数都会改变真实切块，必须各自进 profile（否则改参数后 docHash 不变）。
    for (const [key, value] of [['maxInputChars', 2_000], ['maxChunkChars', 800], ['minChunkChars', 150], ['maxUnitsPerBatch', 20], ['maxOutputTokens', 2_000]]) {
        assert.notEqual(
            kbLlmProfile({ ...BASE_CONFIG, [key]: value }),
            kbLlmProfile(BASE_CONFIG),
            key + ' 变化必须改变 llm profile',
        );
    }
});

await check('提示词只带编号与原文，不带块正文元数据', async () => {
    const source = sampleDoc(1, 2);
    const plan = planKbBatches(source, { maxInputChars: 10_000 });
    const prompt = buildKbLlmPrompt({ units: plan.batches[0].units, source });
    assert.ok(prompt.includes('[0] '));
    assert.ok(prompt.includes('第1章'));
    assert.ok(!prompt.includes('span'));
});


await check('输出预算收紧：默认 1500 token 每箱最多 8 组，提示词写明上限', async () => {
    assert.equal(kbLlmBatchLimit(DEFAULT_KB_LLM), 8);
    assert.equal(kbLlmBatchLimit({ ...BASE_CONFIG, maxUnitsPerBatch: 60, maxOutputTokens: 8_000 }), 44);
    const source = longDoc(20_000);
    const plan = planKbBatches(source, { maxInputChars: 60_000, maxUnitsPerBatch: 60, maxOutputTokens: 1_500 });
    assert.ok(plan.batches.length >= 2, '20k 字符在 8 组/箱下应分成至少两箱');
    for (const batch of plan.batches) assert.ok(batch.units.length <= 8, '单箱分组数不得超过输出预算：' + batch.units.length);
    const estimate = estimateKbLlmJob(source, { maxInputChars: 60_000, maxUnitsPerBatch: 60, maxOutputTokens: 1_500 });
    assert.equal(estimate.unitsPerBatch, 8);
    const system = buildKbLlmSystemPrompt(8);
    assert.ok(system.includes('最多 8 个分组'), system);
    assert.ok(system.includes('120 字'), system);
});

await check('截断重试：finish=length 且分组不完整，翻倍输出上限重试一次', async () => {
    const source = sampleDoc(1, 2);
    const seen = [];
    const generate = async ({ prompt, maxTokens }) => {
        seen.push({ maxTokens, units: unitsInPrompt(prompt) });
        if (seen.length === 1) {
            return { ok: true, text: JSON.stringify({ groups: [] }), finish: 'length', truncated: true };
        }
        return {
            ok: true,
            text: JSON.stringify({ groups: [{ from: 0, to: seen[seen.length - 1].units - 1, title: '全段', summary: '总结' }] }),
            finish: 'stop',
            truncated: false,
        };
    };
    const result = await runKbLlmJob({ text: source, config: { ...BASE_CONFIG, maxOutputTokens: 1_500 }, generate });
    assert.equal(result.status, 'done');
    assert.equal(result.calls, 2, '首答加一次重试');
    assert.equal(result.retried, 1);
    assert.equal(result.truncated, 1);
    assert.equal(result.degraded, 0, '重试成功后不算降级');
    assert.equal(seen[0].maxTokens, 1_500);
    assert.equal(seen[1].maxTokens, 3_000);
    assertCoverage(source, result.chunks);
});

await check('截断两次：明确降级并写明截断，不静默拼接半份 JSON', async () => {
    const source = sampleDoc(1, 2);
    let called = 0;
    const generate = async () => {
        called += 1;
        return { ok: true, text: JSON.stringify({ groups: [] }), finish: 'length', truncated: true };
    };
    const result = await runKbLlmJob({ text: source, config: BASE_CONFIG, generate });
    assert.equal(called, 2, '同一箱只重试一次');
    assert.equal(result.calls, 2);
    assert.equal(result.retried, 1);
    assert.equal(result.truncated, 2);
    assert.ok(result.degraded >= 1);
    assert.ok(result.errors.some((line) => line.includes('截断')), result.errors.join(' | '));
    assertCoverage(source, result.chunks);
});

await check('预算不够时不重试：只降级，不越总调用上限', async () => {
    const source = sampleDoc(1, 2);
    let called = 0;
    const generate = async () => {
        called += 1;
        return { ok: true, text: JSON.stringify({ groups: [] }), finish: 'length', truncated: true };
    };
    const result = await runKbLlmJob({ text: source, config: { ...BASE_CONFIG, maxCalls: 1 }, generate });
    assert.equal(called, 1);
    assert.equal(result.retried, 0);
    assert.equal(result.truncated, 1);
    assert.ok(result.degraded >= 1);
});

await check('任务进度落盘：读写往返、换模型换键、完成后可清', async () => {
    const dir = fileURLToPath(new URL('../../../.office/runtime/tmp/kb-llm-jobs', import.meta.url));
    await rm(dir, { recursive: true, force: true });
    const text = sampleDoc(1, 2);
    const config = { ...BASE_CONFIG, maxInputChars: 5_000 };
    const key = kbLlmJobKeyOf(text, config);
    assert.match(key, /^kbllm-[0-9a-f]{16}$/);
    assert.notEqual(key, kbLlmJobKeyOf(text, { ...config, model: 'another-model' }), '换模型必须换任务键');

    const first = await runKbLlmJob({ text, config, generate: makeFakeGenerator() });
    assert.equal(first.status, 'done');
    await writeKbLlmJob(dir, key, { version: 1, key, state: first.state });
    const back = await readKbLlmJob(dir, key);
    assert.equal(back.state.cursor, first.state.cursor);
    const resumed = await runKbLlmJob({ text, config, generate: makeFakeGenerator(), state: back.state });
    assert.equal(resumed.status, 'done');

    await dropKbLlmJob(dir, key);
    assert.equal(await readKbLlmJob(dir, key), null);
    assert.equal(await readKbLlmJob(dir, 'kbllm-0000000000000000'), null);
    // 非法键不抛给上层，也绝不能读到目录外（返回 null）。
    assert.equal(await readKbLlmJob(dir, '../../escape'), null);
});


await check('摘要映射：每个确定性块都归到一段摘要，块不丢', async () => {
    const chunks = [{ span: { start: 0, end: 10 } }, { span: { start: 10, end: 20 } }, { span: { start: 20, end: 30 } }];
    const segments = [
        { span: { start: 0, end: 20 }, title: '第一段', summary: '讲了 A' },
        { span: { start: 20, end: 30 }, title: '第二段', summary: '讲了 B' },
    ];
    const mapped = mapSummariesToChunks({ segments, chunks });
    assert.equal(mapped.length, 2);
    assert.deepEqual(mapped[0].chunkIndexes, [0, 1]);
    assert.deepEqual(mapped[1].chunkIndexes, [2]);
    assert.equal(mapped[0].title, '第一段');
    assert.equal(mapped[0].text, '讲了 A');
});

await check('摘要映射：跨段的块归重叠更大的那一段', async () => {
    const chunks = [{ span: { start: 0, end: 10 } }];
    const segments = [
        { span: { start: 0, end: 7 }, title: '左', summary: 'x' },
        { span: { start: 5, end: 10 }, title: '右', summary: 'y' },
    ];
    const mapped = mapSummariesToChunks({ segments, chunks });
    assert.equal(mapped.length, 1);
    assert.equal(mapped[0].title, '左');
});

await check('摘要映射：空输入与坏区间返回空，不抛', async () => {
    assert.deepEqual(mapSummariesToChunks({ segments: [], chunks: [] }), []);
    assert.deepEqual(mapSummariesToChunks({ segments: [{ span: { start: 5, end: 5 } }], chunks: [{ span: { start: 0, end: 9 } }] }), []);
    assert.deepEqual(mapSummariesToChunks({ segments: [{ span: null }], chunks: [{ span: { start: 0, end: 9 } }] }), []);
});

await check('摘要映射：覆盖全文时块恰好各归一段，没有重复', async () => {
    const chunks = Array.from({ length: 10 }, (_, i) => ({ span: { start: i * 10, end: i * 10 + 10 } }));
    const segments = [
        { span: { start: 0, end: 30 }, title: 'a', summary: 'a' },
        { span: { start: 30, end: 70 }, title: 'b', summary: 'b' },
        { span: { start: 70, end: 100 }, title: 'c', summary: 'c' },
    ];
    const mapped = mapSummariesToChunks({ segments, chunks });
    const all = mapped.flatMap((item) => item.chunkIndexes);
    assert.equal(all.length, 10);
    assert.deepEqual([...new Set(all)].sort((x, y) => x - y), [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
});

// ── 汇总 ───────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log((item.ok ? '  ok  ' : ' FAIL ') + item.name);
    if (!item.ok) console.log('       ' + String(item.error?.message ?? item.error).split('\n').join('\n       '));
}
console.log('\n' + (results.length - failed.length) + '/' + results.length + ' 通过');
process.exitCode = failed.length === 0 ? 0 : 1;
