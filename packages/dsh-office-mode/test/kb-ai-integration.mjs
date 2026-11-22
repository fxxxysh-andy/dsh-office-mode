/**
 * 知识库「可选大模型切片 / 摘要」的离线集成测试（不是 kb-llm 纯函数单测）。
 *
 * 覆盖真实链路，而不是只测参数折叠：
 *   1. 默认离线：不传 ai 时一次模型请求都不发，docHash 仍是确定性档；
 *   2. AI 未启用时 ai:true 明确拒绝、aiEstimate 只预检不落库；
 *   3. 真模型切片 + 摘要：块 id 来自真实 docHash（llm profile），摘要带来源块 id、
 *      能 kb-read、能进 kb-search；
 *   4. 模型失败按箱降级，但正文逐字仍是原文切片（span 覆盖无缺口）；
 *   5. 单轮上限 → paused 不落库，同内容 + aiResume 续跑完成后才落库；
 *   6. 上传路由默认离线、AI/模型/切片/摘要/预估/续跑参数原样透传；
 *   7. 宿主模型清单投影；kb-asset GET/HEAD 安全链；图库路由真实挂载且未配置如实。
 *
 * 模型端点用 127.0.0.1 上的假 OpenAI 兼容服务（node:http），不碰任何真实 Key / 云端。
 *
 * 跑法：node test/kb-ai-integration.mjs
 */
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { mkdir, mkdtemp, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { resolveConfig } from '../src/config.js';
import { zip } from '../src/engine/zip.js';
import { createMemory, renderKbIngest } from '../src/memory.js';
import { chunkCharsFor, chunkIdOf, chunkOverlapFor, docIdOf, kbHash } from '../src/kb.js';
import { kbLlmProfile, normalizeKbLlmConfig } from '../src/kb-llm.js';
import {
    KB_ASSET_PATH,
    KB_UPLOAD_PATH,
    LLM_MODELS_PATH,
    GALLERY_PATH,
    noteWorkspace,
    registerMemoryView,
    resetWorkspaces,
} from '../src/view.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
const results = [];
async function check(name, fn) {
    try { await fn(); results.push({ name, ok: true }); }
    catch (error) { results.push({ name, ok: false, error }); }
}

// ── 假模型端点：按 prompt 里的小节编号回分组，可按模型名注入失败 ──────────────
const hits = [];
const failuresLeft = new Map();
const server = createServer((req, res) => {
    let raw = '';
    req.on('data', (chunk) => { raw += chunk; });
    req.on('end', () => {
        if (!String(req.url ?? '').endsWith('/chat/completions')) {
            res.writeHead(404, { 'content-type': 'text/plain' });
            res.end('not found');
            return;
        }
        let body = null;
        try { body = JSON.parse(raw); } catch { body = null; }
        const model = String(body?.model ?? 'unknown');
        hits.push(model);
        const remaining = failuresLeft.get(model) ?? 0;
        if (remaining > 0) {
            failuresLeft.set(model, remaining - 1);
            res.writeHead(500, { 'content-type': 'text/plain' });
            res.end('fake upstream failure');
            return;
        }
        const prompt = String((Array.isArray(body?.messages) ? body.messages : [])
            .find((item) => item?.role === 'user')?.content ?? '');
        const numbers = [...prompt.matchAll(/^\[(\d+)\] /gm)].map((match) => Number(match[1]));
        const unitCount = numbers.length === 0 ? 0 : Math.max(...numbers) + 1;
        const groups = [];
        for (let from = 0; from < unitCount; from += 2) {
            groups.push({
                from,
                to: Math.min(from + 1, unitCount - 1),
                title: '第 ' + from + ' 段',
                summary: '萤火虫摘要 ' + model + ' 第' + from + '段',
            });
        }
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify({
            model,
            choices: [{ message: { content: JSON.stringify({ groups }) }, finish_reason: 'stop' }],
        }));
    });
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const port = server.address().port;
const BASE_URL = 'http://127.0.0.1:' + port + '/v1';

/** 5 个约 265 字的小节 → maxInputChars=1000 时正好 2 箱（3 小节 + 2 小节）。 */
function longText(tag) {
    return Array.from({ length: 5 }, (_, index) => tag + '第' + index + '段：' + '河'.repeat(260)).join('\n\n');
}

/** 递归列出目录里的相对路径（排序后精确比对，用来断言「预估一个字节都没写」）。 */
async function listTree(dir, prefix = '') {
    let entries = [];
    try { entries = await readdir(dir, { withFileTypes: true }); } catch { return []; }
    const out = [];
    for (const entry of [...entries].sort((left, right) => left.name.localeCompare(right.name))) {
        const rel = prefix === '' ? entry.name : prefix + '/' + entry.name;
        if (entry.isDirectory()) out.push(rel + '/', ...await listTree(join(dir, entry.name), rel));
        else out.push(rel);
    }
    return out;
}

/** 最小 DOCX：一份 document.xml + 一张嵌入图（测试内现造，不依赖仓库内夹具）。 */
function minimalDocx(bodyText) {
    const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        + '<w:p><w:r><w:t>' + bodyText + '</w:t></w:r></w:p>'
        + '</w:body></w:document>';
    return zip([
        { name: '[Content_Types].xml', data: '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>' },
        { name: 'word/document.xml', data: documentXml },
        { name: 'word/media/image1.png', data: PNG },
    ]);
}

function memoryConfig(base, llm = {}) {
    return resolveConfig({
        memory: {
            enabled: true,
            dir: join(base, 'memory'),
            kbDir: join(base, 'kb'),
            llm: {
                mode: 'openai',
                baseURL: BASE_URL,
                model: 'trace-1',
                apiKey: 'local-noop',
                chunking: true,
                summary: true,
                maxInputChars: 1_000,
                maxChunkChars: 700,
                minChunkChars: 200,
                maxCallsPerRun: 20,
                maxCalls: 50,
                timeoutMs: 5_000,
                ...llm,
            },
        },
        gallery: {},
    });
}

/**
 * 原文覆盖判据：按 span 排序检查「块正文 = 原文切片」，块之间的空隙只能是空白，
 * 末尾也不许丢正文。
 *
 * 第一百零一轮起相邻块**带重叠**（overlapChunks），所以这里不再要求 span 严格
 * 不重叠，而是：重叠量不超过 overlap 上限；起点不许倒退（否则一块会被另一块
 * 整个包住）；覆盖只增不减，最终必须到达原文末尾。
 */
function assertVerbatim(source, chunks, label, overlap = 300) {
    const sorted = [...chunks].sort((left, right) => left.span.start - right.span.start);
    let cursor = 0;
    let previousStart = -1;
    for (const chunk of sorted) {
        assert.equal(chunk.text, source.slice(chunk.span.start, chunk.span.end), label + '：块正文必须等于原文切片');
        assert.ok(chunk.span.start > previousStart, label + '：块起点必须严格递增');
        previousStart = chunk.span.start;
        if (chunk.span.start >= cursor) {
            assert.equal(source.slice(cursor, chunk.span.start).trim(), '', label + '：块之间的空隙只能是空白');
        } else {
            assert.ok(cursor - chunk.span.start <= overlap,
                label + `：重叠量超限 ${cursor - chunk.span.start} > ${overlap}`);
        }
        cursor = Math.max(cursor, chunk.span.end);
    }
    assert.equal(source.slice(cursor).trim(), '', label + '：末尾剩下的只能是空白');
}

function mockResponse() {
    return {
        status: 0,
        headers: {},
        chunks: [],
        writeHead(status, headers = {}) { this.status = status; this.headers = { ...headers }; },
        end(chunk) { if (chunk !== undefined && chunk !== null) this.chunks.push(Buffer.from(chunk)); },
        json() { return JSON.parse(Buffer.concat(this.chunks).toString('utf8') || '{}'); },
        bytes() { return Buffer.concat(this.chunks); },
    };
}

function mockRequest({ method = 'GET', url = '/', headers = {}, body = null } = {}) {
    const stream = Readable.from(body === null ? [] : [Buffer.from(body)]);
    stream.method = method;
    stream.url = url;
    stream.headers = headers;
    return stream;
}

function fakeGalleryCtx(webServer, llm = null, workspaceRegistry = null) {
    return {
        get(name) {
            if (name === 'webServer') return webServer;
            if (name === 'llm') return llm;
            if (name === 'workspaceRegistry') return workspaceRegistry;
            return undefined;
        },
    };
}

function routeTable() {
    const routes = new Map();
    return {
        routes,
        webServer: { register({ path, handler }) { routes.set(path, handler); return () => routes.delete(path); } },
    };
}

async function postJson(routes, path, payload) {
    const body = JSON.stringify(payload);
    const handler = routes.get(new URL(path, 'http://127.0.0.1').pathname);
    assert.equal(typeof handler, 'function', '路由未注册：' + path);
    const res = mockResponse();
    await handler(mockRequest({
        method: 'POST',
        url: path,
        headers: { 'content-type': 'application/json', 'content-length': String(Buffer.byteLength(body)) },
        body,
    }), res);
    return res;
}

await check('默认离线：ai 不传时零模型请求，docHash 仍是确定性档，重开实例仍读得到且正文逐字保真', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-offline-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const text = longText('离线');
        await writeFile(join(root, '笔记.md'), text);
        const cfg = memoryConfig(base);
        const memory = createMemory({ root, memory: cfg.memory });
        const before = hits.length;
        const result = await memory.kbIngest({ path: '笔记.md' });
        assert.equal(hits.length, before, '默认离线不得发任何模型请求');
        assert.equal(result.ai, undefined, '默认离线回执不应带 ai 段');
        assert.equal(result.unchanged, false);
        const expectedHash = kbHash(text, chunkCharsFor('笔记.md') + ':' + chunkOverlapFor(chunkCharsFor('笔记.md')));
        assert.equal(result.docId, docIdOf('笔记.md'), '文档 id 按路径寻址');
        assert.equal(result.chunks >= 1, true);

        const reopened = createMemory({ root, memory: cfg.memory });
        const list = await reopened.kbList();
        assert.equal(list.total, 1);
        assert.equal(list.items[0].path, '笔记.md');
        assert.equal(list.items[0].hash, expectedHash);
        const read = await reopened.kbRead({ path: '笔记.md' });
        assert.equal(read.docId, 'kb:' + kbHash('笔记.md'));
        for (const chunk of read.chunks) {
            assert.equal(chunk.id, chunkIdOf(expectedHash, chunk.n), '块 id 必须来自确定性 docHash');
        }
        assertVerbatim(text, read.chunks, '确定性切块');
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('AI 未启用：ai:true 明确拒绝；aiEstimate 只预检、不调模型也不落库', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-off-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const text = longText('未启用');
        await writeFile(join(root, '未启用.md'), text);
        const cfg = memoryConfig(base, { mode: 'off' });
        const memory = createMemory({ root, memory: cfg.memory });
        const before = hits.length;
        await assert.rejects(
            () => memory.kbIngest({ path: '未启用.md', ai: true }),
            /AI 建库未启用/,
        );
        const estimate = await memory.kbIngest({ path: '未启用.md', aiEstimate: true });
        assert.equal(estimate.dryRun, true);
        assert.equal(estimate.estimate.calls, 0);
        assert.equal(estimate.estimate.batches, 0);
        assert.equal(estimate.ai.mode, 'off');
        assert.equal(hits.length, before, '预检与拒绝都不得发模型请求');
        const list = await memory.kbList();
        assert.equal(list.total, 0, '预检不得落库');
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('只读预估不落盘：含嵌入图的 DOCX 预估后知识库目录一个字节都不多，也不发模型请求', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-estimate-nowrite-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        await writeFile(join(root, '嵌入图.docx'), minimalDocx(longText('嵌入')));
        const cfg = memoryConfig(base);
        const memory = createMemory({ root, memory: cfg.memory });
        const beforeTree = await listTree(base);
        const beforeHits = hits.length;
        const result = await memory.kbIngest({
            path: '嵌入图.docx', ai: true, aiEstimate: true, aiChunking: true, aiSummary: true,
        });
        assert.equal(result.dryRun, true);
        assert.ok(result.estimate.calls >= 2, '预估要给出真实调用数，实际 ' + result.estimate.calls);
        assert.equal(result.ai.available, true);
        assert.equal((await memory.kbList()).total, 0, '预估不得落库');
        assert.deepEqual(await listTree(base), beforeTree, '预估不得写任何资产 / 清单 / 进度文件');
        assert.equal(hits.length, beforeHits, '预估不得发模型请求');
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('宿主 llm 不可用：预估如实报不可用、0 次调用，且没有任何出站请求', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-host-unavailable-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const text = longText('宿主');
        await writeFile(join(root, '宿主.md'), text);
        const cfg = memoryConfig(base, { mode: 'host', provider: 'fake-provider', model: 'fake-1' });
        const memory = createMemory({ root, memory: cfg.memory }); // 没有宿主 llm
        const beforeHits = hits.length;
        const realFetch = globalThis.fetch;
        let outbound = 0;
        globalThis.fetch = (...args) => { outbound += 1; return realFetch(...args); };
        let result;
        try {
            result = await memory.kbIngest({ path: '宿主.md', ai: true, aiEstimate: true, aiChunking: true, aiSummary: true });
        } finally {
            globalThis.fetch = realFetch;
        }
        assert.equal(result.dryRun, true);
        assert.equal(result.ai.available, false);
        assert.match(result.ai.reason, /宿主没有可用的 llm/);
        assert.equal(result.estimate.calls, 0, '不可用时不得报可执行的调用数');
        assert.equal(result.estimate.batches, 0, '0 次调用时不应再报模型批数');
        assert.match(result.note, /不会发起任何模型请求/);
        assert.equal(outbound, 0, '预估不得发起任何出站请求');
        assert.equal(hits.length, beforeHits);
        assert.match(renderKbIngest(result), /通道不可用/);
        assert.equal((await memory.kbList()).total, 0);
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('同内容重入预估为 0 次调用：不写盘、不发请求；retier 与低预算/无宿主 llm 下真实重入不跑模型', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-estimate-duplicate-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const text = longText('重入');
        await writeFile(join(root, '重入.md'), text);
        // host 侧的 profile 是 provider/model；openai 侧就是 model 字符串。
        // 让首次入库用 "fake-provider/fake-1" 作 openai model，后面就能用
        // host + 同一对 provider/model 命中同一个 docHash，验证「重入优先于
        // 预算 / 宿主可用性」，而不需要再造一个假宿主流。
        const cfg = memoryConfig(base, { model: 'fake-provider/fake-1' });
        const memory = createMemory({ root, memory: cfg.memory });
        const first = await memory.kbIngest({ path: '重入.md', ai: true, aiChunking: true, aiSummary: true });
        assert.ok(first.ai.calls >= 2, '首次入库应有真实模型调用');
        const hitsAfterFirst = hits.length;
        const beforeTree = await listTree(base);

        const estimate = await memory.kbIngest({ path: '重入.md', ai: true, aiEstimate: true, aiChunking: true, aiSummary: true });
        assert.equal(estimate.dryRun, true);
        assert.equal(estimate.duplicate, true);
        assert.equal(estimate.estimate.calls, 0, '同内容重入实际不会调模型，预估必须给 0');
        assert.equal(estimate.estimate.batches, 0, '0 次调用时不应再报模型批数');
        assert.equal(estimate.estimate.estimatedInputTokens, 0);
        assert.equal(estimate.estimate.estimatedOutputTokens, 0);
        assert.equal(estimate.estimate.overBudget, false);
        assert.equal(estimate.ai.calls, 0);
        assert.equal(hits.length, hitsAfterFirst, '重复预估不得发模型请求');
        assert.deepEqual(await listTree(base), beforeTree, '重复预估不得写盘');

        const retierEstimate = await memory.kbIngest({
            path: '重入.md', tier: 'unverified', ai: true, aiEstimate: true, aiChunking: true, aiSummary: true,
        });
        assert.equal(retierEstimate.duplicate, true);
        assert.equal(retierEstimate.estimate.calls, 0, 'retier 不重跑模型，预估也是 0');

        const retiered = await memory.kbIngest({
            path: '重入.md', tier: 'unverified', ai: true, aiChunking: true, aiSummary: true,
        });
        assert.equal(retiered.retiered, true);
        assert.equal(retiered.unchanged, true);
        assert.equal(hits.length, hitsAfterFirst, 'retier 重入不得调模型');
        assert.equal((await memory.kbList()).items[0].tier, 'unverified', 'retier 行为保持不变');

        // 同一份内容再换一个「预算 1 次 + 没有宿主 llm」的实例：预估和真实入库
        // 都必须先认出「不需要模型」，不能被 overBudget / host 不可用挡住。
        const tightCfg = memoryConfig(base, {
            mode: 'host', provider: 'fake-provider', model: 'fake-1', maxCalls: 1,
        });
        const tight = createMemory({ root, memory: tightCfg.memory }); // 没有宿主 llm
        const hitsBeforeTight = hits.length;
        const tightEstimate = await tight.kbIngest({
            path: '重入.md', ai: true, aiEstimate: true, aiChunking: true, aiSummary: true,
        });
        assert.equal(tightEstimate.duplicate, true, '重入识别必须早于预算 / 宿主可用性');
        assert.equal(tightEstimate.estimate.calls, 0);
        assert.equal(tightEstimate.estimate.batches, 0);
        assert.equal(tightEstimate.estimate.overBudget, false);
        assert.equal(tightEstimate.ai.available, true, '重入不需要模型，不应误报通道不可用');
        const tightIngest = await tight.kbIngest({
            path: '重入.md', ai: true, aiChunking: true, aiSummary: true,
        });
        assert.equal(tightIngest.unchanged, true, '预算 1 次 / 无宿主 llm 下重入仍应直接回执未改动');
        assert.equal(hits.length, hitsBeforeTight, '重入不得改走模型');
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('AI 真切片 + 摘要：块 id 来自 llm profile docHash，摘要可追溯、可 kb-read、可 kb-search', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-trace-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const text = longText('追溯');
        await writeFile(join(root, '追溯.md'), text);
        const cfg = memoryConfig(base);
        const memory = createMemory({ root, memory: cfg.memory });
        const result = await memory.kbIngest({
            path: '追溯.md', ai: true, aiModel: 'trace-1', aiChunking: true, aiSummary: true,
        });
        assert.equal(result.ai.mode, 'openai');
        assert.equal(result.ai.model, 'trace-1');
        assert.equal(result.ai.degraded, 0);
        assert.ok(result.ai.calls >= 2, '两箱正文至少要两次模型调用，实际 ' + result.ai.calls);
        assert.ok(result.ai.summaries >= 1, '摘要必须落库');
        const expectedHash = kbHash(text, kbLlmProfile(normalizeKbLlmConfig(cfg.memory.llm)));
        assert.equal(result.docId, docIdOf('追溯.md'), '文档 id 按路径寻址');
        assert.notEqual(expectedHash, kbHash(text, chunkCharsFor('追溯.md')), 'AI 档与确定性档不能共用 hash');

        const read = await memory.kbRead({ path: '追溯.md' });
        assertVerbatim(text, read.chunks, 'AI 切片');
        for (const chunk of read.chunks) {
            assert.equal(chunk.id, chunkIdOf(expectedHash, chunk.n));
        }
        // 模型切片的块也要带重叠（第一百零一轮）：后一块的起点退进前一块的正文里，
        // 退让量由块大小派生（与确定性切块同一个 chunkOverlapFor）。
        const aiGap = chunkOverlapFor(normalizeKbLlmConfig(cfg.memory.llm).maxChunkChars);
        for (let index = 1; index < read.chunks.length; index += 1) {
            const prev = read.chunks[index - 1];
            const cur = read.chunks[index];
            const shared = prev.span.end - cur.span.start;
            assert.ok(shared > 0, `AI 第 ${index + 1} 块没有与上一块重叠`);
            assert.ok(shared <= aiGap, `AI 切块的重叠量超限：${shared} > ${aiGap}`);
            assert.equal(prev.text.slice(prev.text.length - shared), cur.text.slice(0, shared),
                '重叠窗口里两块必须是同一段原文');
        }

        const summaries = await memory.kbSummaries();
        assert.equal(summaries.length, result.ai.summaries);
        const chunkIds = new Set(read.chunks.map((chunk) => chunk.id));
        for (const summary of summaries) {
            assert.ok(Array.isArray(summary.chunkIds) && summary.chunkIds.length > 0, '摘要必须带来源块 id');
            for (const id of summary.chunkIds) assert.ok(chunkIds.has(id), '摘要来源块必须在正文里：' + id);
            assert.match(summary.text, /trace-1/);
        }
        const summaryRead = await memory.kbRead({ id: 'summary:' + summaries[0].id });
        assert.equal(summaryRead.chunks[0].text, summaries[0].text);
        assert.deepEqual(summaryRead.summaryOf, summaries[0].chunkIds);

        const found = await memory.kbSearch({ query: 'trace-1' });
        const hit = found.hits.find((row) => row.kind === 'summary');
        assert.ok(hit !== undefined, '摘要必须能被 kb-search 命中');
        assert.ok(hit.sourceChunkIds.length > 0);
        assert.match(hit.preview, /萤火虫摘要/);

        const reopened = createMemory({ root, memory: cfg.memory });
        assert.equal((await reopened.kbList()).total, 1);
        assert.equal((await reopened.kbSummaries()).length, summaries.length);
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('切块参数进真实 docHash：改 maxChunkChars 换块 id；只开摘要不重切', async () => {
    const baseA = await mkdtemp(join(tmpdir(), 'kbai-profile-a-'));
    const baseB = await mkdtemp(join(tmpdir(), 'kbai-profile-b-'));
    const baseC = await mkdtemp(join(tmpdir(), 'kbai-profile-c-'));
    try {
        const text = longText('参数');
        async function ingest(base, file, llm, extra) {
            const root = join(base, 'work');
            await mkdir(root, { recursive: true });
            await writeFile(join(root, file), text);
            const cfg = memoryConfig(base, llm);
            const memory = createMemory({ root, memory: cfg.memory });
            const result = await memory.kbIngest({ path: file, ai: true, aiChunking: true, aiSummary: true, ...extra });
            const hash = (await memory.kbList()).items[0].hash;
            return { cfg, result, hash };
        }
        const a = await ingest(baseA, '参数.md', { model: 'profile-1', maxChunkChars: 700 });
        const b = await ingest(baseB, '参数.md', { model: 'profile-1', maxChunkChars: 1_200 });
        assert.equal(a.hash, kbHash(text, kbLlmProfile(normalizeKbLlmConfig(a.cfg.memory.llm))));
        assert.equal(b.hash, kbHash(text, kbLlmProfile(normalizeKbLlmConfig(b.cfg.memory.llm))));
        assert.notEqual(a.hash, b.hash, 'maxChunkChars 变化必须换真实块 id');

        const rootC = join(baseC, 'work');
        await mkdir(rootC, { recursive: true });
        await writeFile(join(rootC, '只摘要.md'), text);
        const cfgC = memoryConfig(baseC, { model: 'profile-1', maxChunkChars: 1_200, minChunkChars: 200 });
        const memoryC = createMemory({ root: rootC, memory: cfgC.memory });
        const summaryOnly = await memoryC.kbIngest({
            path: '只摘要.md', ai: true, aiChunking: false, aiSummary: true,
        });
        assert.equal(summaryOnly.ai.chunking, false);
        assert.ok(summaryOnly.ai.summaries >= 1);
        const hashC = (await memoryC.kbList()).items[0].hash;
        // 确定性档 = 块大小 + 重叠量（两条都改变块边界，见 kbHash 与 chunkOverlapFor）。
        assert.equal(hashC, kbHash(text, chunkCharsFor('只摘要.md') + ':' + chunkOverlapFor(chunkCharsFor('只摘要.md'))),
            '只开摘要必须沿用确定性 docHash，不重切');
        assert.notEqual(hashC, kbHash(text, kbLlmProfile(normalizeKbLlmConfig(cfgC.memory.llm))), '摘要-only 不能落进 AI 切块档');
        const readC = await memoryC.kbRead({ path: '只摘要.md' });
        assertVerbatim(text, readC.chunks, '摘要-only');
    } finally {
        await rm(baseA, { recursive: true, force: true });
        await rm(baseB, { recursive: true, force: true });
        await rm(baseC, { recursive: true, force: true });
    }
});

await check('失败回退：模型 500 时该箱按确定性规则降级，但正文逐字仍是原文', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-degrade-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const text = longText('降级');
        await writeFile(join(root, '降级.md'), text);
        failuresLeft.set('flaky-1', 1);
        const cfg = memoryConfig(base, { model: 'flaky-1' });
        const memory = createMemory({ root, memory: cfg.memory });
        const result = await memory.kbIngest({ path: '降级.md', ai: true, aiChunking: true, aiSummary: true });
        assert.equal(result.ai.degraded, 1, '只应降级失败的那一箱');
        assert.equal(result.ai.errors.length, 1);
        assert.match(result.ai.errors[0], /批次 0/);
        assert.match(result.ai.errors[0], /HTTP 500/);
        const read = await memory.kbRead({ path: '降级.md' });
        assertVerbatim(text, read.chunks, '降级切块');
        assert.equal(read.chunks.length >= 2, true);
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('预估续跑：单轮上限先 paused 不落库，同内容 aiResume 后续跑完成才落库', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-resume-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const text = longText('续跑');
        await writeFile(join(root, '续跑.md'), text);
        const cfg = memoryConfig(base, { model: 'resume-1', maxCallsPerRun: 1 });
        const memory = createMemory({ root, memory: cfg.memory });
        const first = await memory.kbIngest({ path: '续跑.md', ai: true, aiChunking: true, aiSummary: true });
        assert.equal(first.paused, true);
        assert.equal(first.status, 'paused');
        assert.equal(first.calls, 1);
        assert.match(first.stateKey, /^kbllm-[0-9a-f]{16}$/);
        assert.equal((await memory.kbList()).total, 0, 'paused 不得落库');

        const second = await memory.kbIngest({
            path: '续跑.md', ai: true, aiChunking: true, aiSummary: true, aiResume: true,
        });
        assert.equal(second.paused, undefined);
        assert.equal(second.unchanged, false);
        assert.ok(second.ai.calls >= 2, '续跑后总调用数应覆盖两箱，实际 ' + second.ai.calls);
        assert.equal((await memory.kbList()).total, 1);
        const read = await memory.kbRead({ path: '续跑.md' });
        assertVerbatim(text, read.chunks, '续跑切块');
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('上传路由：默认 ai=false 零请求；AI/模型/切片/摘要/预估/续跑参数透传', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-upload-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        noteWorkspace(root);
        const text = longText('上传');
        const data = Buffer.from(text, 'utf8').toString('base64');
        const cfg = memoryConfig(base, { model: 'upload-1', maxCallsPerRun: 20 });
        const { routes, webServer } = routeTable();
        const hostLlm = {
            async listProviders() { return ['fake-provider']; },
            async listModels(provider) { return provider === 'fake-provider' ? [{ id: 'fake-1', name: '假模型一号' }] : []; },
        };
        const dispose = registerMemoryView(fakeGalleryCtx(webServer, hostLlm), {
            getMemory: () => cfg.memory,
            getConfig: () => cfg,
        });
        assert.equal(typeof dispose, 'function');
        assert.ok(routes.has(KB_UPLOAD_PATH));
        const cwdQuery = '?cwd=' + encodeURIComponent(root);

        const before = hits.length;
        const plain = await postJson(routes, KB_UPLOAD_PATH + '?cwd=' + encodeURIComponent(root), {
            name: '上传笔记.md', folder: '项目', data,
        });
        const plainBody = plain.json();
        assert.equal(plainBody.ok, true, JSON.stringify(plainBody));
        assert.equal(plainBody.result.path, 'uploads/项目/上传笔记.md');
        assert.equal(plainBody.result.ai, undefined, '默认不得开 AI');
        assert.equal(hits.length, before, '默认上传不得发模型请求');

        const estimated = await postJson(routes, KB_UPLOAD_PATH + cwdQuery, {
            name: '预检.md', data, ai: true, aiModel: 'upload-1', aiChunking: true, aiSummary: true, aiEstimate: true,
        });
        const estimatedBody = estimated.json();
        assert.equal(estimatedBody.ok, true, JSON.stringify(estimatedBody));
        assert.equal(estimatedBody.result.dryRun, true);
        assert.ok(estimatedBody.result.estimate.calls >= 2, '预检要给出真实调用数');
        assert.ok(estimatedBody.result.estimate.batches >= 2);
        assert.equal(estimatedBody.result.ai.model, 'upload-1');
        assert.equal(hits.length, before, '预检不得发模型请求');

        const first = await postJson(routes, KB_UPLOAD_PATH + cwdQuery, {
            name: '续传.md', data, ai: true, aiModel: 'upload-1', aiChunking: true, aiSummary: true, aiRunLimit: 1,
        });
        const firstBody = first.json();
        assert.equal(firstBody.ok, true, JSON.stringify(firstBody));
        assert.equal(firstBody.result.paused, true, 'aiRunLimit=1 应让上传暂停');
        assert.equal(firstBody.result.calls, 1, '上传路由必须透传 aiRunLimit');
        assert.equal(firstBody.result.ai.model, 'upload-1');

        const resumed = await postJson(routes, KB_UPLOAD_PATH + cwdQuery, {
            name: '续传.md', data, ai: true, aiModel: 'upload-1', aiChunking: true, aiSummary: true, aiResume: true,
        });
        const resumedBody = resumed.json();
        assert.equal(resumedBody.ok, true, JSON.stringify(resumedBody));
        assert.equal(resumedBody.result.paused, undefined);
        assert.equal(resumedBody.result.ai.summary, true);
        assert.ok(resumedBody.result.ai.summaries >= 1, '续跑完成必须写摘要');
        assert.equal(resumedBody.result.path, 'uploads/续传.md');

        const llmRes = mockResponse();
        await routes.get(LLM_MODELS_PATH)(mockRequest({
            url: LLM_MODELS_PATH + cwdQuery,
        }), llmRes);
        const listedBody = llmRes.json();
        assert.equal(listedBody.ok, true);
        assert.deepEqual(listedBody.providers, [{ provider: 'fake-provider', models: [{ id: 'fake-1', label: '假模型一号' }] }]);
        dispose();
        assert.equal(routes.size, 0);
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

await check('kb-asset 安全链：GET 带安全头返回真实字节，HEAD 无正文；图库路由真实挂载且未配置如实', async () => {
    const base = await mkdtemp(join(tmpdir(), 'kbai-asset-'));
    try {
        const root = join(base, 'work');
        await mkdir(root, { recursive: true });
        const kbDir = join(base, 'kb');
        await mkdir(join(kbDir, 'assets'), { recursive: true });
        await writeFile(join(kbDir, 'assets', '图.png'), PNG);
        noteWorkspace(root);
        const cfg = memoryConfig(base);
        const { routes, webServer } = routeTable();
        const dispose = registerMemoryView(fakeGalleryCtx(webServer), {
            getMemory: () => cfg.memory,
            getConfig: () => cfg,
        });
        const cwdQuery = '?cwd=' + encodeURIComponent(root);
        const assetUrl = KB_ASSET_PATH + cwdQuery + '&path=' + encodeURIComponent('assets/图.png');
        const got = mockResponse();
        await routes.get(KB_ASSET_PATH)(mockRequest({ url: assetUrl }), got);
        assert.equal(got.status, 200);
        assert.equal(got.headers['content-type'], 'image/png');
        assert.equal(got.headers['x-content-type-options'], 'nosniff');
        assert.match(String(got.headers['cache-control'] ?? ''), /private/);
        assert.deepEqual(got.bytes(), PNG);

        const head = mockResponse();
        await routes.get(KB_ASSET_PATH)(mockRequest({ method: 'HEAD', url: assetUrl }), head);
        assert.equal(head.status, 200);
        assert.equal(head.bytes().length, 0, 'HEAD 不得有正文');

        const status = mockResponse();
        await routes.get(GALLERY_PATH)(mockRequest({ url: GALLERY_PATH + cwdQuery + '&view=status' }), status);
        const statusBody = status.json();
        assert.equal(status.status, 200);
        assert.equal(statusBody.ok, true);
        assert.equal(statusBody.vision.configured, false);
        assert.equal(statusBody.imageEmbedding.configured, false);
        dispose();
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

server.close();
for (const item of results) {
    if (item.ok) console.log('PASS  ' + item.name);
    else console.log('FAIL  ' + item.name + '\n      ' + (item.error?.stack ?? item.error));
}
const failed = results.filter((item) => !item.ok);
console.log('\n' + (results.length - failed.length) + '/' + results.length + ' 通过');
process.exitCode = failed.length === 0 ? 0 : 1;
