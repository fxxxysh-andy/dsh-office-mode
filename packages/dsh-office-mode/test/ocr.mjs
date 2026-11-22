/**
 * 文件识别（OCR）通道的测试（第九十一轮）。
 *
 * 钉住五件事：
 *   1. **契约形状**：端点 https://api.deepseek.com/v1/tools/ocr，multipart 只有 file + purpose，
 *      带 Bearer Key —— 与官方文档 api-docs.deepseek.com/zh-cn/api/ocr 一致（多传字段官方 400）。
 *   2. **200 也可能失败**：识别开始后状态码已是 200，失败信息在响应体 { error } 里。
 *   3. **错误可执行**：缺 Key / 文件缺失 / 超限 / 413 / 429 都给出 code、retryable 与下一步。
 *   4. **不吞错**：超时、非法 JSON、没有 pages 都要抛 OcrError，而不是返回空结果。
 *   5. **纯函数**：ref/box 标记清理、按页拼接。
 *   6. **缺 Key 说清是哪一份环境变量**（第九十五轮）：报设置里配的名字，说明它取的是
 *      启动 DSH 的那个进程继承的环境变量（同一个终端、改完要重启），而名字不合法时
 *      一个字都不回显 —— 那种值往往是误粘进去的 Key。
 *
 * 不联网：全部用注入的 fetchImpl。跑法：node test/ocr.mjs
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { resolveConfig } from '../src/config.js';
import {
    DEFAULT_OCR,
    OCR_MAX_BYTES,
    OcrError,
    cleanOcrText,
    ocrDocument,
    ocrEndpoint,
    ocrResultText,
    officialOcrEndpointOf,
    resolveOcrConfig,
    safeOcrKeyEnvLabel,
} from '../src/ocr.js';
import {
    HOST_OCR_OWNER,
    HOST_OCR_PROVIDER,
    identifyHostOcrProvider,
    resolveHostOcrCredential,
} from '../src/ocr-host-credential.js';
import { isLocalBrowserRequest, officialReuseGate } from '../src/ocr-local-request.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const root = mkdtempSync(join(tmpdir(), 'office-ocr-'));
const sample = join(root, 'sample.pdf');
writeFileSync(sample, Buffer.from('%PDF-1.4 test bytes'));

function jsonResponse(status, body, contentType = 'application/json') {
    return {
        ok: status >= 200 && status < 300,
        status,
        headers: { get: (name) => (String(name).toLowerCase() === 'content-type' ? contentType : null) },
        text: async () => (typeof body === 'string' ? body : JSON.stringify(body)),
    };
}

await check('配置：默认值、purpose 回落、Key 从环境变量补上', () => {
    const resolved = resolveConfig({});
    assert.equal(resolved.ocr.baseURL, DEFAULT_OCR.baseURL);
    assert.equal(resolved.ocr.purpose, 'text_recognition');
    assert.equal(resolved.ocr.enabled, true);
    assert.equal(resolved.ocr.apiKeyEnv, 'DEEPSEEK_API_KEY');
    assert.equal(resolveOcrConfig({ purpose: '不认识的模式' }).purpose, 'text_recognition');
    assert.equal(resolveOcrConfig({ purpose: 'rich_layout_recognition' }).purpose, 'rich_layout_recognition');
});

await check('端点：根地址与已带 /v1 的地址都拼成 /v1/tools/ocr', () => {
    assert.equal(ocrEndpoint(''), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(ocrEndpoint('https://api.deepseek.com'), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(ocrEndpoint('https://api.deepseek.com/'), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(ocrEndpoint('https://gateway.example.com/v1'), 'https://gateway.example.com/v1/tools/ocr');
});

/** 抓一次「缺 Key」的错误（断言写在调用点，失败消息才不会只有一句 rejects 超时）。 */
async function missingKeyError(options = {}) {
    let called = 0;
    try {
        await ocrDocument({
            filePath: sample, apiKey: '', apiKeyEnv: undefined, ...options,
            fetchImpl: async () => { called += 1; return jsonResponse(200, {}); },
        });
    } catch (error) {
        return { error, called };
    }
    throw new Error('缺 Key 应当抛 missing_api_key，但没有抛');
}

await check('缺 Key：明确报 missing_api_key 并指向设置页', async () => {
    const { error } = await missingKeyError();
    assert.ok(error instanceof OcrError);
    assert.equal(error.code, 'missing_api_key');
    assert.match(error.hint, /设置/);
});

await check('缺 Key：环境变量名用设置里的那个，并说清它来自启动 DSH 的终端（第九十五轮）', async () => {
    const { error, called } = await missingKeyError({ apiKeyEnv: 'OFFICE_OCR_TEST_KEY' });
    assert.equal(error.code, 'missing_api_key');
    assert.match(error.message, /OFFICE_OCR_TEST_KEY/, '要报设置里配的那个名字，而不是永远报默认名');
    assert.equal(/DEEPSEEK_API_KEY/.test(error.message), false, '名字不是用户配的就不该出现在消息里');
    // 可执行的三件事：UI 路径、取的是宿主继承的环境变量、要重启。
    assert.match(error.hint, /设置 → 办公模式 → 文件识别（OCR）/);
    assert.match(error.hint, /OFFICE_OCR_TEST_KEY/);
    assert.match(error.hint, /那个进程继承的环境变量/);
    assert.match(error.hint, /启动 DSH 的那个终端/);
    assert.match(error.hint, /重启 DSH/);
    assert.equal(called, 0, '缺 Key 是本地判断，不该发请求');
    // 不填名字时保持旧文案：仍是默认的 DEEPSEEK_API_KEY。
    const fallback = await missingKeyError();
    assert.match(fallback.error.message, /DEEPSEEK_API_KEY/);
    assert.equal(fallback.called, 0);
});

await check('缺 Key：变量名不像变量名（可能是误粘的 Key）时原样不回显（第九十五轮）', async () => {
    const pasted = 'sk-live-0123456789abcdef';
    const { error, called } = await missingKeyError({ apiKeyEnv: pasted });
    assert.equal(error.code, 'missing_api_key');
    // 名字里带 key 的值一个字都不能回来：那正是「把 Key 粘到变量名框」的样子。
    assert.equal(error.message.includes(pasted), false, '不合法的变量名不能出现在消息里');
    assert.equal(error.hint.includes(pasted), false, '不合法的变量名不能出现在提示里');
    assert.match(error.message, /不是合法变量名/);
    assert.match(error.hint, /文件识别（OCR）/);
    assert.match(error.hint, /DEEPSEEK_API_KEY/, '提示里给一个合法变量名当例子');
    assert.equal(called, 0, '缺 Key 是本地判断，不该发请求');
    // 纯函数：空 = 默认名；合法名原样；带连字符 / 空白 / 超长的一律不给。
    assert.equal(safeOcrKeyEnvLabel(''), 'DEEPSEEK_API_KEY');
    assert.equal(safeOcrKeyEnvLabel('  OFFICE_KEY_2  '), 'OFFICE_KEY_2');
    assert.equal(safeOcrKeyEnvLabel('sk-live-x'), '');
    assert.equal(safeOcrKeyEnvLabel('OFFICE KEY'), '');
    assert.equal(safeOcrKeyEnvLabel('1LEADING_DIGIT'), '');
    assert.equal(safeOcrKeyEnvLabel('A'.repeat(129)), '');
});

await check('文件检查：缺失 / 空文件 / 超过 100 MiB 都先被本地拦下', async () => {
    await assert.rejects(
        () => ocrDocument({ filePath: join(root, 'nope.pdf'), apiKey: 'k' }),
        (error) => error.code === 'file_missing',
    );
    const empty = join(root, 'empty.png');
    writeFileSync(empty, Buffer.alloc(0));
    await assert.rejects(
        () => ocrDocument({ filePath: empty, apiKey: 'k' }),
        (error) => error.code === 'file_empty',
    );
    const big = join(root, 'big.pdf');
    writeFileSync(big, Buffer.alloc(OCR_MAX_BYTES + 1));
    await assert.rejects(
        () => ocrDocument({ filePath: big, apiKey: 'k' }),
        (error) => error.code === 'file_too_large' && /100 MiB/.test(error.message),
    );
});

await check('成功路径：POST multipart 只有 file + purpose，解析保活换行后的 JSON', async () => {
    let seen = null;
    const fetchImpl = async (url, init) => {
        seen = { url, init };
        const body = JSON.stringify({
            id: 'ocr-1', object: 'ocr.result', created: 1, purpose: 'text_recognition', page_count: 2,
            pages: [
                { index: 0, text: '第一页 <｜ref｜>标题<｜/ref｜><｜box｜>[[10, 20, 300, 40]]<｜/box｜>', finish_reason: 'stop' },
                { index: 1, text: '第二页', finish_reason: 'length' },
            ],
            usage: { input_tokens: 100, output_tokens: 20, total_tokens: 120 },
        });
        return jsonResponse(200, '\n\n' + body + '\n');
    };
    const result = await ocrDocument({ filePath: sample, apiKey: 'sk-test', fetchImpl, purpose: 'rich_layout_recognition' });
    assert.equal(seen.url, 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(seen.init.method, 'POST');
    const headerOf = (name) => {
        const headers = seen.init.headers || {};
        for (const key of Object.keys(headers)) { if (key.toLowerCase() === name) return headers[key]; }
        return undefined;
    };
    assert.equal(headerOf('authorization'), 'Bearer sk-test');
    assert.equal(headerOf('content-type'), undefined, 'multipart 的 Content-Type 必须由 FormData 自己带 boundary');
    const form = seen.init.body;
    assert.equal(typeof form.get, 'function', '请求体要是 FormData');
    assert.equal(form.get('file').name, 'sample.pdf');
    assert.equal(form.get('purpose'), 'rich_layout_recognition');
    // 多传字段官方直接 400：这里逐项核对字段集合，回归时能立刻发现。
    const keys = [];
    for (const key of form.keys()) keys.push(key);
    assert.deepEqual(keys.sort(), ['file', 'purpose']);
    assert.equal(result.pageCount, 2);
    assert.equal(result.pages[0].index, 0);
    assert.equal(result.pages[1].finishReason, 'length');
    assert.equal(result.usage.total_tokens, 120);
    assert.equal(result.requestURL, 'https://api.deepseek.com/v1/tools/ocr');
});

await check('200 但响应体是 { error }：必须当失败抛出（官方语义）', async () => {
    const fetchImpl = async () => jsonResponse(200, { error: { message: 'OCR did not finish in time', type: 'gateway_timeout_error', code: 'gateway_timeout_error' } });
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl }),
        (error) => error instanceof OcrError && error.code === 'gateway_timeout_error' && error.retryable === true,
    );
});

await check('HTTP 413 / 429：状态码翻成中文，可重试标记跟着官方语义', async () => {
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl: async () => jsonResponse(413, 'file is too large: max 100 MiB') }),
        (error) => error.status === 413 && error.retryable === false && /100 MiB/.test(error.message),
    );
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl: async () => jsonResponse(429, { error: { message: 'too many concurrent jobs' } }) }),
        (error) => error.status === 429 && error.retryable === true,
    );
});

await check('超时：按 timeoutMs 中止并给出 timeout，而不是空结果', async () => {
    const fetchImpl = (url, init) => new Promise((resolve, reject) => {
        init.signal.addEventListener('abort', () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' })));
    });
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl, timeoutMs: 30 }),
        (error) => error instanceof OcrError && error.code === 'timeout',
    );
});

await check('响应体卡住：超时继续管到 body 读完，报 timeout（不是空响应）', async () => {
    const fetchImpl = async (url, init) => ({
        ok: true, status: 200,
        headers: { get: () => 'application/json' },
        text: () => new Promise((resolve, reject) => {
            const onAbort = () => reject(Object.assign(new Error('The operation was aborted'), { name: 'AbortError' }));
            if (init.signal.aborted) onAbort();
            else init.signal.addEventListener('abort', onAbort);
        }),
    });
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl, timeoutMs: 40 }),
        (error) => {
            assert.equal(error.code, 'timeout');
            assert.equal(error.status, 200, '上游 HTTP 状态保留给诊断（端点再映射 504）');
            assert.match(error.message, /响应体/);
            assert.equal(/空响应/.test(error.message), false, 'body 卡住不能被当成空响应');
            return true;
        },
    );
});

await check('响应体读到一半断开：body_interrupted + 上游状态，只带安全错误码，不回显异常文字', async () => {
    const interrupted = Object.assign(new Error('read ECONNRESET at /home/alice/secret/gateway'), { code: 'ECONNRESET' });
    await assert.rejects(
        () => ocrDocument({
            filePath: sample, apiKey: 'k',
            fetchImpl: async () => ({
                ok: true, status: 200,
                headers: { get: () => 'application/json' },
                text: () => Promise.reject(interrupted),
            }),
        }),
        (error) => {
            assert.equal(error.code, 'body_interrupted');
            assert.equal(error.status, 200);
            assert.equal(error.retryable, true);
            assert.match(error.message, /ECONNRESET/);
            assert.equal(error.message.includes('/home/alice'), false, '异常文字里的本机路径不能外发');
            assert.equal(error.message.includes('read ECONNRESET'), false, '异常 message 原样不引用');
            return true;
        },
    );
    // 没有「像错误码」的 code 的异常：连错误码都不带，更不带 message。
    const bare = Object.assign(new Error('socket hang up at /home/bob'), { code: 'not a safe code!' });
    await assert.rejects(
        () => ocrDocument({
            filePath: sample, apiKey: 'k',
            fetchImpl: async () => ({
                ok: true, status: 200,
                headers: { get: () => 'application/json' },
                text: () => Promise.reject(bare),
            }),
        }),
        (error) => {
            assert.equal(error.code, 'body_interrupted');
            assert.equal(/socket hang up|\/home\/bob|not a safe code/.test(error.message), false);
            return true;
        },
    );
});

await check('耗时按响应收完计：body 晚到时 ms 覆盖读取时间', async () => {
    const fetchImpl = async () => ({
        ok: true, status: 200,
        headers: { get: () => 'application/json' },
        text: async () => {
            await new Promise((resolve) => setTimeout(resolve, 40));
            return JSON.stringify({ page_count: 1, pages: [{ index: 0, text: 'x', finish_reason: 'stop' }] });
        },
    });
    const result = await ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl });
    assert.ok(result.ms >= 35, 'ms 要包含 body 读取（当前 ' + result.ms + ' ms）');
});

await check('坏响应：不是 JSON / 没有 pages 都不能当成功', async () => {
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl: async () => jsonResponse(200, 'not json at all', 'text/plain') }),
        (error) => error.code === 'bad_response',
    );
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl: async () => jsonResponse(200, { object: 'ocr.result' }) }),
        (error) => error.code === 'bad_response',
    );
});

// 上游原文的回显会把网关页面、内网报文、本机路径一起带回浏览器：下面几条钉住
// 「只回报状态 + content-type + 分类」，且官方结构化 body.error 的 message 仍然保留。
const RAW_MARKER = 'C:\\Users\\alice\\.secret\\gateway-internal.html';

await check('200 + HTML 网关页：错误里只有状态/类型/分类，不回显页面原文', async () => {
    const page = '<html><body><h1>502 Bad Gateway</h1><p>' + RAW_MARKER + '</p></body></html>';
    await assert.rejects(
        () => ocrDocument({
            filePath: sample, apiKey: 'k',
            fetchImpl: async () => jsonResponse(200, page, 'text/html; charset=utf-8'),
        }),
        (error) => {
            assert.equal(error.code, 'bad_response');
            assert.equal(error.status, 200);
            assert.equal(error.contentType, 'text/html');
            assert.equal(error.bodyKind, 'html');
            assert.match(error.message, /HTTP 200/);
            assert.match(error.message, /text\/html/);
            assert.match(error.message, /HTML\/XML/);
            assert.equal(error.message.includes(RAW_MARKER), false, '上游原文（可能是本机路径）不能出现在错误消息里');
            assert.equal(error.message.includes('<html'), false, '页面标记不能出现在错误消息里');
            return true;
        },
    );
});

await check('200 + 空响应体：分类为「空响应体」，不编造原文', async () => {
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: 'k', fetchImpl: async () => jsonResponse(200, '   ', 'application/json') }),
        (error) => {
            assert.equal(error.code, 'bad_response');
            assert.equal(error.bodyKind, 'empty');
            assert.match(error.message, /空响应体/);
            assert.match(error.message, /HTTP 200/);
            return true;
        },
    );
});

await check('HTTP 502 + HTML 网关页：保留 502 与分类，不回显页面原文', async () => {
    const page = '<!doctype html><title>Bad Gateway</title>' + RAW_MARKER;
    await assert.rejects(
        () => ocrDocument({
            filePath: sample, apiKey: 'k',
            fetchImpl: async () => jsonResponse(502, page, 'text/html'),
        }),
        (error) => {
            assert.equal(error.status, 502);
            assert.equal(error.code, 'http_502');
            assert.equal(error.retryable, true);
            assert.equal(error.bodyKind, 'html');
            assert.match(error.message, /HTTP 502/);
            assert.match(error.message, /HTML\/XML/);
            assert.equal(error.message.includes(RAW_MARKER), false, '网关页面原文不能进错误消息');
            return true;
        },
    );
});

await check('HTTP 502 + 结构化 body.error：官方 message 与 code 照旧保留', async () => {
    await assert.rejects(
        () => ocrDocument({
            filePath: sample, apiKey: 'k',
            fetchImpl: async () => jsonResponse(502, {
                error: { message: 'OCR did not finish in time', type: 'gateway_timeout_error', code: 'gateway_timeout_error' },
            }),
        }),
        (error) => {
            assert.equal(error.code, 'gateway_timeout_error');
            assert.match(error.message, /OCR did not finish in time/);
            // 结构化 body.error 不是原文回显，contentType/bodyKind 也就不必标。
            assert.equal(error.bodyKind, '');
            return true;
        },
    );
});

await check('纯函数：清掉 ref/box 版面标记，按页拼接', () => {
    const raw = '标题<｜ref｜>第一章<｜/ref｜><｜box｜>[[1, 2, 3, 4]]<｜/box｜>\n正文';
    assert.equal(cleanOcrText(raw), '标题第一章\n正文');
    const merged = ocrResultText({ pages: [{ text: 'a', finishReason: 'stop' }, { text: 'b', finishReason: 'stop' }] });
    assert.equal(merged, 'a\n\nb');
});

// ── 宿主官方供应商 Key 复用（第九十七轮） ──────────────────────────────────
// 只在内存里拼一个假的宿主 ctx：不读真实 config / credentials / profile，也绝不
// 对真实凭据库调用 resolve —— 下面所有 Key 都是固定的假串。

function fakeHostCtx({
    ref = 'DEEPSEEK_API_KEY', owner = HOST_OCR_OWNER, state = 2, resolve, environment,
    baseURL = 'https://api.deepseek.com/anthropic', baseURLGetter,
} = {}) {
    const calls = { resolve: 0 };
    const ctx = {
        get(name) {
            if (name === 'llm') {
                return {
                    listProviders: () => [{ id: HOST_OCR_PROVIDER }],
                    listConfigurableProviders: () => [{
                        provider: HOST_OCR_PROVIDER,
                        settingsNs: 'llm',
                        settingsPath: [],
                    }],
                };
            }
            if (name === 'loader') {
                return {
                    entries: () => [{
                        options: { id: 'llm', name: owner },
                        fiber: {
                            state,
                            effectiveConfig: {
                                apiKeyEnv: { get: () => ref },
                                baseURL: { get: baseURLGetter ?? (() => baseURL) },
                            },
                        },
                    }],
                };
            }
            if (name === 'credentials') {
                return resolve === undefined ? undefined : {
                    resolve: async (credentialRef) => {
                        calls.resolve += 1;
                        return resolve(credentialRef);
                    },
                };
            }
            if (name === 'launchEnvironment') return environment;
            return undefined;
        },
    };
    return { ctx, calls };
}

await check('官方地址矩阵：只认 api.deepseek.com 的 HTTPS 根 / /v1，其它一律不借 Key', () => {
    assert.equal(officialOcrEndpointOf(''), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(officialOcrEndpointOf('  https://api.deepseek.com  '), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(officialOcrEndpointOf('https://api.deepseek.com:443'), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(officialOcrEndpointOf('https://api.deepseek.com/'), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(officialOcrEndpointOf('https://api.deepseek.com/v1'), 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(officialOcrEndpointOf('https://api.deepseek.com/v1/'), 'https://api.deepseek.com/v1/tools/ocr');
    for (const rejected of [
        'http://api.deepseek.com',
        'https://user:pass@api.deepseek.com',
        'https://api.deepseek.com:8443',
        'https://api.deepseek.com?x=1',
        'https://api.deepseek.com#frag',
        'https://api.deepseek.com/v2',
        'https://api.deepseek.com/v1/tools/ocr',
        'https://api.deepseek.com.evil.example',
        'https://api.deepseek.com./v1',
    ]) {
        assert.equal(officialOcrEndpointOf(rejected), undefined, '不该把 Key 发到 ' + rejected);
    }
});

await check('供应商 Key 模式：端点钉死官方、拒绝重定向（redirect: error）', async () => {
    let seen = null;
    const fetchImpl = async (url, init) => {
        seen = { url, init };
        return jsonResponse(200, {
            id: 'ocr-host', object: 'ocr.result', purpose: 'text_recognition', page_count: 1,
            pages: [{ index: 0, text: 'ok', finish_reason: 'stop' }],
        });
    };
    const result = await ocrDocument({
        filePath: sample,
        apiKey: 'sk-fake-host-key',
        apiKeySource: 'official-provider',
        baseURL: 'https://api.deepseek.com/v1',
        fetchImpl,
    });
    assert.equal(seen.url, 'https://api.deepseek.com/v1/tools/ocr');
    assert.equal(seen.init.redirect, 'error', '3xx 必须直接失败，不能把 Key 带去别处');
    assert.equal(seen.init.headers.authorization, 'Bearer sk-fake-host-key');
    assert.equal(result.requestURL, 'https://api.deepseek.com/v1/tools/ocr');
});

await check('供应商 Key 模式：自定义 baseURL 在发请求前就拒绝（不换地址、不发出）', async () => {
    let called = 0;
    await assert.rejects(
        () => ocrDocument({
            filePath: sample,
            apiKey: 'sk-fake-host-key',
            apiKeySource: 'official-provider',
            baseURL: 'https://ocr-proxy.example.com/v1',
            fetchImpl: async () => { called += 1; throw new Error('不该请求'); },
        }),
        (error) => error instanceof OcrError && error.code === 'endpoint_not_official'
            && error.message.includes('ocr-proxy.example.com') === false,
    );
    assert.equal(called, 0, '地址不匹配时连请求都不能发');
});

await check('供应商 Key 模式：错误投影剥掉借来的 Key，保留安全传输码', async () => {
    const secret = 'sk-fake-host-key-sentinel';
    const base = {
        filePath: sample,
        apiKey: secret,
        apiKeySource: 'official-provider',
        baseURL: 'https://api.deepseek.com',
    };
    const assertNoSecret = (error) => {
        const projected = JSON.stringify({
            message: error.message ?? '',
            code: error.code ?? '',
            detail: error.detail ?? '',
            hint: error.hint ?? '',
        });
        assert.equal(projected.includes(secret), false, projected);
    };

    await assert.rejects(
        () => ocrDocument({
            ...base,
            fetchImpl: async () => jsonResponse(401, {
                error: { code: secret, message: 'bad key ' + secret },
            }),
        }),
        (error) => {
            assertNoSecret(error);
            assert.equal(error.status, 401);
            assert.equal(error.code, 'http_401', 'code 整个是 Key 时要退回可执行的状态码');
            assert.equal(error.message.includes('bad key'), true);
            assert.equal(error.detail.includes(secret), false);
            return true;
        },
    );

    await assert.rejects(
        () => ocrDocument({
            ...base,
            fetchImpl: async () => jsonResponse(200, {
                error: { code: 'ocr_' + secret, message: 'failed ' + secret },
            }),
        }),
        (error) => {
            assertNoSecret(error);
            assert.match(error.message, /OCR 识别失败/);
            assert.equal(error.detail.includes(secret), false);
            return true;
        },
    );

    const refusal = Object.assign(new Error('connect failed with ' + secret), { code: 'ECONNREFUSED' });
    await assert.rejects(
        () => ocrDocument({ ...base, fetchImpl: async () => { throw refusal; } }),
        (error) => {
            assertNoSecret(error);
            assert.equal(error.code, 'network');
            assert.match(error.message, /ECONNREFUSED/, '安全传输码要保留，便于排查');
            assert.equal(error.message.includes('connect failed'), false, '借来的 Key 下不回显任意异常原文');
            return true;
        },
    );
});

await check('缺 Key 文案：只有官方供应商行存在时才提 Models，默认不改变旧提示', async () => {
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: '', hostProviderHint: true }),
        (error) => {
            assert.equal(error.code, 'missing_api_key');
            assert.match(error.hint, /设置 → 模型（Models）/);
            assert.match(error.hint, /文件识别（OCR）/, '旧的环境变量指引仍在');
            return true;
        },
    );
    await assert.rejects(
        () => ocrDocument({ filePath: sample, apiKey: '' }),
        (error) => {
            assert.equal(error.code, 'missing_api_key');
            assert.equal(/Models/.test(error.hint), false, '没确认官方供应商前不能引导用户去 Models');
            return true;
        },
    );
});

await check('凭据桥：自定义 baseURL 连供应商配置都不读（resolver 零调用）', async () => {
    const { ctx, calls } = fakeHostCtx({ resolve: async () => ({ value: 'sk-fake' }) });
    const result = await resolveHostOcrCredential(ctx, { baseURL: 'https://ocr-proxy.example.com' });
    assert.deepEqual(result, { ok: false, eligible: false, reason: 'endpoint_not_official' });
    assert.equal(calls.resolve, 0);
});

await check('凭据桥：供应商实际 baseURL 不是官方 origin 时 fail closed，resolver 零调用', async () => {
    for (const baseURL of [
        'https://gateway.example/anthropic',
        'http://api.deepseek.com/anthropic',
        'https://api.deepseek.com:8443/anthropic',
    ]) {
        const { ctx, calls } = fakeHostCtx({ baseURL, resolve: async () => ({ value: 'sk-fake-gateway' }) });
        assert.deepEqual(
            await resolveHostOcrCredential(ctx, { baseURL: 'https://api.deepseek.com' }),
            { ok: false, eligible: false, reason: 'provider_endpoint_not_official' },
            baseURL,
        );
        assert.equal(calls.resolve, 0, '自定义供应商网关的 token 不能走到 resolver');
        assert.equal(identifyHostOcrProvider(ctx), undefined);
    }

    const officialPath = fakeHostCtx({
        baseURL: 'https://api.deepseek.com/anthropic/v1',
        resolve: async () => ({ value: 'sk-fake-official' }),
    });
    assert.deepEqual(
        await resolveHostOcrCredential(officialPath.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: true, apiKey: 'sk-fake-official', ref: 'DEEPSEEK_API_KEY' },
    );
});

await check('凭据桥：supplier baseURL 读不出来 / 缺 launchEnvironment 时 fail closed', async () => {
    const thrown = fakeHostCtx({
        baseURLGetter: () => { throw new Error('read /home/alice/secret-config failed'); },
        resolve: async () => ({ value: 'sk-fake' }),
    });
    const thrownResult = await resolveHostOcrCredential(thrown.ctx, { baseURL: 'https://api.deepseek.com' });
    assert.deepEqual(thrownResult, { ok: false, eligible: false, reason: 'provider_endpoint_unreadable' });
    assert.equal(thrown.calls.resolve, 0);
    assert.equal(JSON.stringify(thrownResult).includes('/home/alice'), false);

    const noService = fakeHostCtx({ baseURL: null, resolve: async () => ({ value: 'sk-fake' }) });
    assert.deepEqual(
        await resolveHostOcrCredential(noService.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: false, eligible: false, reason: 'provider_endpoint_unreadable' },
    );
    assert.equal(noService.calls.resolve, 0, '不能把缺失的启动环境当成 process.env 或官方默认');

    const broken = fakeHostCtx({
        baseURL: null,
        environment: { get: () => { throw new Error('env unreadable'); } },
        resolve: async () => ({ value: 'sk-fake' }),
    });
    assert.deepEqual(
        await resolveHostOcrCredential(broken.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: false, eligible: false, reason: 'provider_endpoint_unreadable' },
    );
    assert.equal(broken.calls.resolve, 0);
});

await check('凭据桥：supplier baseURL 的后备只走 launchEnvironment，显式值不二次解析', async () => {
    const envCustom = fakeHostCtx({
        baseURL: null,
        environment: { get: (name) => (name === 'DEEPSEEK_BASE_URL' ? { value: 'https://gateway.example' } : undefined) },
        resolve: async () => ({ value: 'sk-fake' }),
    });
    assert.deepEqual(
        await resolveHostOcrCredential(envCustom.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: false, eligible: false, reason: 'provider_endpoint_not_official' },
    );
    assert.equal(envCustom.calls.resolve, 0);

    let environmentReads = 0;
    const explicit = fakeHostCtx({
        baseURL: 'https://api.deepseek.com/anthropic',
        environment: { get: () => { environmentReads += 1; return { value: 'https://gateway.example' }; } },
        resolve: async () => ({ value: 'sk-fake-official' }),
    });
    assert.deepEqual(
        await resolveHostOcrCredential(explicit.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: true, apiKey: 'sk-fake-official', ref: 'DEEPSEEK_API_KEY' },
    );
    assert.equal(environmentReads, 0, 'baseURL getter 已有显式值时不能再猜环境');
    assert.equal(explicit.calls.resolve, 1);

    const defaulted = fakeHostCtx({
        baseURL: null,
        environment: { get: () => undefined },
        resolve: async () => ({ value: 'sk-fake-default' }),
    });
    assert.deepEqual(
        await resolveHostOcrCredential(defaulted.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: true, apiKey: 'sk-fake-default', ref: 'DEEPSEEK_API_KEY' },
    );
    assert.equal(defaulted.calls.resolve, 1);
});

await check('凭据桥：只解析官方供应商行里实际配置的那个引用', async () => {
    const { ctx, calls } = fakeHostCtx({
        ref: 'OFFICE_PERSONAL_DEEPSEEK_KEY',
        resolve: async (ref) => {
            assert.equal(ref, 'OFFICE_PERSONAL_DEEPSEEK_KEY', '必须按实时配置的引用解析，不能猜默认名');
            return { value: 'sk-fake-resolved', source: 'file' };
        },
    });
    assert.deepEqual(identifyHostOcrProvider(ctx), { ref: 'OFFICE_PERSONAL_DEEPSEEK_KEY' });
    const result = await resolveHostOcrCredential(ctx, { baseURL: 'https://api.deepseek.com' });
    assert.deepEqual(result, { ok: true, apiKey: 'sk-fake-resolved', ref: 'OFFICE_PERSONAL_DEEPSEEK_KEY' });
    assert.equal(calls.resolve, 1);
});

await check('凭据桥：owner 不对 / fiber 未激活 / 引用名非法都不借', async () => {
    for (const options of [
        { owner: '@deepseek-ai/dsh-llm-deepseek' },
        { state: 1 },
        { ref: 'bad-ref' },
    ]) {
        const { ctx, calls } = fakeHostCtx(options);
        const result = await resolveHostOcrCredential(ctx, { baseURL: 'https://api.deepseek.com' });
        assert.deepEqual(result, { ok: false, eligible: false, reason: 'provider_absent' });
        assert.equal(calls.resolve, 0);
    }
});

await check('凭据桥：没 Key 时 eligible=true；解析异常只回结构化原因，不回显异常文字', async () => {
    const missing = fakeHostCtx({ resolve: async () => undefined });
    assert.deepEqual(
        await resolveHostOcrCredential(missing.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: false, eligible: true, reason: 'credential_missing' },
    );

    const broken = fakeHostCtx({
        resolve: async () => { throw new Error('read /home/alice/.credentials.yaml failed for sk-live-secret'); },
    });
    const result = await resolveHostOcrCredential(broken.ctx, { baseURL: 'https://api.deepseek.com' });
    assert.deepEqual(result, { ok: false, eligible: true, reason: 'credential_unavailable' });
    assert.equal(JSON.stringify(result).includes('/home/alice'), false);
    assert.equal(JSON.stringify(result).includes('sk-live-secret'), false);
});

await check('凭据桥：只在没有 credentials 服务时读 launchEnvironment，绝不直接翻 process.env', async () => {
    const environment = { get: (ref) => (ref === 'DEEPSEEK_API_KEY' ? { value: 'sk-fake-launch-env' } : undefined) };
    const fromEnv = fakeHostCtx({ environment });
    assert.deepEqual(
        await resolveHostOcrCredential(fromEnv.ctx, { baseURL: 'https://api.deepseek.com' }),
        { ok: true, apiKey: 'sk-fake-launch-env', ref: 'DEEPSEEK_API_KEY' },
    );

    const ambientName = 'OFFICE_TEST_HOST_OCR_AMBIENT';
    process.env[ambientName] = 'sk-fake-ambient-must-not-be-read';
    try {
        const noServices = fakeHostCtx({ ref: ambientName });
        const result = await resolveHostOcrCredential(noServices.ctx, { baseURL: 'https://api.deepseek.com' });
        assert.deepEqual(result, { ok: false, eligible: true, reason: 'credential_missing' });
        assert.equal(JSON.stringify(result).includes('sk-fake-ambient'), false);
    } finally {
        delete process.env[ambientName];
    }
});

await check('本机守卫：只认精确回环地址 + 同源 Host/Origin（IPv4 / IPv6 / localhost）', () => {
    const request = (remoteAddress, host, origin, extra = {}) => ({
        socket: { remoteAddress, localPort: 3080, encrypted: false },
        headers: { host, origin, ...extra },
    });
    for (const [remote, host, origin] of [
        ['127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3080'],
        ['::1', '[::1]:3080', 'http://[::1]:3080'],
        ['::ffff:127.0.0.1', 'localhost:3080', 'http://localhost:3080'],
    ]) {
        assert.equal(isLocalBrowserRequest(request(remote, host, origin)).ok, true, remote + ' ' + host);
    }
    const rejected = [
        request('127.0.0.2', '127.0.0.1:3080', 'http://127.0.0.1:3080'),
        request('127.0.0.1', '127.0.0.1:9999', 'http://127.0.0.1:9999'),
        request('127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3080', { 'x-forwarded-for': '203.0.113.5' }),
        request('127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3080', { 'X-Forwarded-Custom': '1' }),
        request('127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3080', { 'sec-fetch-site': 'cross-site' }),
        request('127.0.0.1', '127.0.0.1:3080', undefined),
        request('127.0.0.1', '127.0.0.1:3080', 'https://127.0.0.1:3080'),
        request('127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3081'),
        request('127.0.0.1', '127.0.0.1:3080', 'http://user@127.0.0.1:3080'),
        request('127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3080/path'),
        request('127.0.0.1', '127.0.0.1:3080/evil', 'http://127.0.0.1:3080'),
        request('127.0.0.1', '127.0.0.1:3080?x', 'http://127.0.0.1:3080'),
        request('127.0.0.1', '127.0.0.1:3080 ', 'http://127.0.0.1:3080'),
        request('127.0.0.1', '127.0.0.1:3080,evil', 'http://127.0.0.1:3080'),
        request('127.0.0.1', '127.1:3080', 'http://127.0.0.1:3080'),
        request('127.0.0.1', 'localhost.:3080', 'http://127.0.0.1:3080'),
    ];
    for (const item of rejected) assert.equal(isLocalBrowserRequest(item).ok, false, JSON.stringify(item));
    // HTTPS 连接上的 Host :80 不能被 http 归一化后当成 :443 放过。
    assert.equal(isLocalBrowserRequest({
        socket: { remoteAddress: '127.0.0.1', localPort: 443, encrypted: true },
        headers: { host: '127.0.0.1:80', origin: 'https://127.0.0.1' },
    }).ok, false);
    assert.equal(isLocalBrowserRequest({
        socket: { remoteAddress: '127.0.0.1', localPort: 443, encrypted: true },
        headers: { host: '127.0.0.1', origin: 'https://127.0.0.1' },
    }).ok, true);
    // socket.localPort 无效时不能退化成「任意本机端口都行」。
    assert.equal(isLocalBrowserRequest({
        socket: { remoteAddress: '127.0.0.1', encrypted: false },
        headers: { host: '127.0.0.1:3080', origin: 'http://127.0.0.1:3080' },
    }).ok, false);
    assert.deepEqual(officialReuseGate({ enabled: false, officialEndpoint: true, req: request('127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3080') }), { ok: false, reason: 'disabled' });
    assert.deepEqual(officialReuseGate({ enabled: true, officialEndpoint: false, req: request('127.0.0.1', '127.0.0.1:3080', 'http://127.0.0.1:3080') }), { ok: false, reason: 'custom_endpoint' });
    assert.deepEqual(officialReuseGate({ enabled: true, officialEndpoint: true, req: request('192.168.1.20', '127.0.0.1:3080', 'http://127.0.0.1:3080') }), { ok: false, reason: 'nonlocal' });
});

// ── 汇总 ────────────────────────────────────────────────────────────────────
let failed = 0;
for (const result of results) {
    if (result.ok) console.log('ok   ' + result.name + (result.note ? '  (' + result.note + ')' : ''));
    else {
        failed += 1;
        console.log('FAIL ' + result.name);
        console.log(result.error && result.error.stack ? result.error.stack : String(result.error));
    }
}
try {
    rmSync(root, { recursive: true, force: true });
} catch { /* Windows 占用交给临时目录清理 */ }
console.log('');
console.log(String(results.length - failed) + '/' + String(results.length) + ' passed');
if (failed > 0) process.exitCode = 1;
