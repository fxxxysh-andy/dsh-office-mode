/**
 * `src/kb-selector.js` 的验收测试（第一百零三轮）。
 *
 * 筛选步骤的边界全在这里：只发编号 + 截断块头、只解析编号、失败一律原样放行。
 * 不联网：fetch 是注入的假函数。
 *
 * 跑法：`node test/kb-selector.mjs`
 */
import assert from 'node:assert/strict';
import test from 'node:test';

import {
    DEFAULT_SELECTOR,
    canUseSelector,
    normalizeSelectorConfig,
    selectWithLlm,
    selectorKeyOf,
} from '../src/kb-selector.js';

const HITS = [
    { chunkId: 'c1', title: '甲', head: '甲块的开头', text: '甲块的开头'.repeat(50) },
    { chunkId: 'c2', title: '', head: '', text: '乙块正文'.repeat(80) },
    { chunkId: 'c3', title: '丙', head: '丙块', text: '丙块' },
];

function reply(content, status = 200) {
    return async () => ({ ok: status === 200, status, text: async () => JSON.stringify({ choices: [{ message: { content } }] }) });
}

test('配置收口：默认 off、坏值落默认、clamp 生效', () => {
    const base = normalizeSelectorConfig({});
    assert.equal(base.mode, 'off');
    assert.equal(base.maxCandidates, DEFAULT_SELECTOR.maxCandidates);
    assert.equal(base.maxSelected, DEFAULT_SELECTOR.maxSelected);
    assert.equal(base.headChars, DEFAULT_SELECTOR.headChars);
    const odd = normalizeSelectorConfig({
        mode: 'maybe', baseURL: 'http://x/v1///', timeoutMs: 5, maxCandidates: 999, maxSelected: 0, headChars: 10,
    });
    assert.equal(odd.mode, 'off');
    assert.equal(odd.baseURL, 'http://x/v1', '尾部斜杠要去掉，拼端点时不出现双斜杠');
    assert.equal(odd.timeoutMs, 1000);
    assert.equal(odd.maxCandidates, 30);
    assert.equal(odd.maxSelected, 1);
    assert.equal(odd.headChars, 80);
});

test('可调用条件：必须 llm + 端点 + 模型', () => {
    assert.equal(canUseSelector({ mode: 'off', baseURL: 'http://x/v1', model: 'm' }), false);
    assert.equal(canUseSelector({ mode: 'llm', baseURL: '', model: 'm' }), false);
    assert.equal(canUseSelector({ mode: 'llm', baseURL: 'http://x/v1', model: '' }), false);
    assert.equal(canUseSelector({ mode: 'llm', baseURL: 'http://x/v1', model: 'm' }), true);
});

test('Key：字面量优先，其次环境变量', () => {
    assert.equal(selectorKeyOf({ apiKey: 'literal', apiKeyEnv: 'DSH_X' }, { DSH_X: 'env' }), 'literal');
    assert.equal(selectorKeyOf({ apiKey: '', apiKeyEnv: 'DSH_X' }, { DSH_X: 'env' }), 'env');
    assert.equal(selectorKeyOf({ apiKey: '', apiKeyEnv: '' }, {}), '');
});

test('筛选：只发编号 + 截断块头，不把正文送出去', async () => {
    let captured = null;
    const fetchImpl = async (url, options) => {
        captured = { url, options, body: JSON.parse(options.body) };
        return reply('2,1')();
    };
    const result = await selectWithLlm({
        query: '丙块讲什么',
        hits: HITS,
        config: { mode: 'llm', baseURL: 'http://x/v1', model: 'small', maxCandidates: 2, maxSelected: 2, headChars: 80 },
        fetchImpl,
    });
    assert.equal(result.ok, true);
    assert.deepEqual(result.ids, ['c2', 'c1'], '编号 2、1 要映射回块 id，按模型给出的顺序');
    assert.equal(captured.url, 'http://x/v1/chat/completions');
    assert.equal(captured.body.model, 'small');
    assert.equal(captured.body.temperature, 0);
    assert.equal(captured.body.messages.length, 1);
    const prompt = captured.body.messages[0].content;
    assert.ok(prompt.includes('[1] 甲｜'), '每行要有编号与标题');
    assert.ok(prompt.includes('[2] 乙块正文'), '没有块头时退回正文');
    assert.ok(!prompt.includes('[3]'), 'maxCandidates=2 时第三条不该出现在提示里');
    assert.ok(!prompt.includes('乙块正文'.repeat(40)), '正文要被 headChars 截断，不能整块发出去');
    assert.ok(prompt.includes('只输出编号'));
    assert.ok(prompt.includes('最多 2 个'));
});

test('筛选：返回垃圾 / 越界编号 / 零 时不当作成功', async () => {
    const cfg = { mode: 'llm', baseURL: 'http://x/v1', model: 'm', maxCandidates: 3, maxSelected: 2 };
    const junk = await selectWithLlm({ query: 'q', hits: HITS, config: cfg, fetchImpl: reply('我觉得都不错') });
    assert.equal(junk.ok, false);
    assert.match(junk.error, /没有返回可用编号/);

    const outOfRange = await selectWithLlm({ query: 'q', hits: HITS, config: cfg, fetchImpl: reply('8,9') });
    assert.equal(outOfRange.ok, false, '越界编号要丢掉，丢光就算失败');

    const zero = await selectWithLlm({ query: 'q', hits: HITS, config: cfg, fetchImpl: reply('0') });
    assert.equal(zero.ok, false, '模型输出 0 = 没有相关块，调用方应保留融合结果');
});

test('筛选：HTTP 错误 / 超时 / 未配置都只返回 ok:false，不抛', async () => {
    const cfg = { mode: 'llm', baseURL: 'http://x/v1', model: 'm' };
    const http = await selectWithLlm({ query: 'q', hits: HITS, config: cfg, fetchImpl: async () => ({ ok: false, status: 503, text: async () => 'busy' }) });
    assert.equal(http.ok, false);
    assert.match(http.error, /HTTP 503/);

    const timeout = await selectWithLlm({
        query: 'q', hits: HITS, config: cfg,
        fetchImpl: async () => { const error = new Error('The operation was aborted due to timeout'); error.name = 'TimeoutError'; throw error; },
    });
    assert.equal(timeout.ok, false);
    assert.match(timeout.error, /候选筛选失败/);

    const unconfigured = await selectWithLlm({ query: 'q', hits: HITS, config: { mode: 'off' }, fetchImpl: reply('1') });
    assert.equal(unconfigured.ok, false);
    assert.match(unconfigured.error, /没有配置筛选端点或模型/);
});

test('筛选：空候选直接 ok 且不调模型', async () => {
    let called = 0;
    const result = await selectWithLlm({
        query: 'q', hits: [], config: { mode: 'llm', baseURL: 'http://x/v1', model: 'm' },
        fetchImpl: async () => { called += 1; return reply('1')(); },
    });
    assert.deepEqual(result, { ok: true, ids: [], error: '' });
    assert.equal(called, 0);
});
