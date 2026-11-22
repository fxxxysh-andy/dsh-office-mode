/**
 * OCR 端点（POST /office-memory/ocr）的测试：只测路由这一层。
 *
 * ocr.js 的契约形状另有 test/ocr.mjs；这里守的是「谁能调用、能读哪些文件、
 * 失败怎么如实报、成功怎么把逐页文本拼回去」。上游用假 fetch，不联网。
 *
 * 跑法：node test/view-ocr.mjs
 */
import { existsSync, readdirSync } from 'node:fs';
import { mkdir, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { registerMemoryView, OCR_PATH, noteWorkspace, readJsonBody, resetWorkspaces } from '../src/view.js';

const TMP = fileURLToPath(new URL('../../../.office/runtime/tmp/view-ocr/', import.meta.url));

const results = [];
async function check(name, fn) {
    try { await fn(); results.push({ name, ok: true }); }
    catch (error) { results.push({ name, ok: false, error }); }
}

function makeRes() {
    return {
        status: 0, headers: null, body: "",
        writeHead(status, headers) { this.status = status; this.headers = headers; },
        end(body) { this.body = body ?? ""; },
    };
}

function makeReq({
    method = "POST", url = OCR_PATH, headers = {}, body = "",
    socket = { remoteAddress: "127.0.0.1", localPort: 3080, encrypted: false },
    host = "127.0.0.1:3080",
    origin = "http://127.0.0.1:3080",
} = {}) {
    const handlers = new Map();
    const allHeaders = { host, ...headers };
    if (origin === null) delete allHeaders.origin;
    else if (allHeaders.origin === undefined) allHeaders.origin = origin;
    const socketObj = {
        ...socket,
        destroyed: false,
        once() { return socketObj; },
        off() { return socketObj; },
        end() { socketObj.destroyed = true; },
        destroy() { socketObj.destroyed = true; },
    };
    const req = {
        method, url, headers: allHeaders, socket: socketObj,
        on(event, fn) {
            const list = handlers.get(event) ?? [];
            list.push(fn);
            handlers.set(event, list);
            setTimeout(() => {
                const live = handlers.get(event) ?? [];
                for (const handler of [...live]) {
                    if (event === "data") handler(Buffer.from(body));
                    else if (event === "end") handler();
                }
            }, 0);
            return req;
        },
        once(event, fn) {
            const wrapper = (arg) => { req.off(event, wrapper); fn(arg); };
            const list = handlers.get(event) ?? [];
            list.push(wrapper);
            handlers.set(event, list);
            return req;
        },
        off(event, fn) {
            const list = handlers.get(event);
            if (list) {
                const index = list.indexOf(fn);
                if (index >= 0) list.splice(index, 1);
            }
            return req;
        },
        destroy() { req.destroyed += 1; },
        resume() { req.resumed += 1; },
        destroyed: 0,
        resumed: 0,
    };
    return req;
}

function fakeCtx() {
    const routes = [];
    return {
        routes,
        get(name) {
            if (name === "webServer") {
                return { register(route) { routes.push(route); return () => {}; } };
            }
            return undefined;
        },
    };
}

function ocrFetch(body, { status = 200, text = null, contentType = null } = {}) {
    const calls = [];
    const impl = async (url, init) => {
        calls.push({ url: String(url), init });
        const payload = text === null ? JSON.stringify(body) : text;
        return {
            ok: status >= 200 && status < 300,
            status,
            headers: { get: (name) => (String(name).toLowerCase() === "content-type" ? contentType : null) },
            text: async () => payload,
        };
    };
    return { impl, calls };
}

/** 手动控制事件的请求流：用来单独钉住 readJsonBody 的连接行为（不经过路由）。 */
function streamReq() {
    const handlers = new Map();
    const req = {
        paused: 0, destroyed: 0, resumed: 0,
        on(event, fn) { handlers.set(event, fn); return req; },
        emit(event, arg) { const fn = handlers.get(event); if (fn) fn(arg); },
        pause() { req.paused += 1; },
        destroy() { req.destroyed += 1; },
        resume() { req.resumed += 1; },
    };
    return req;
}

const SUCCESS = {
    id: "ocr-1", object: "ocr.result", created: 1, purpose: "text_recognition", page_count: 2,
    pages: [
        { index: 0, text: "第一页<｜ref｜>标题<｜/ref｜><｜box｜>[[10, 20, 30, 40]]<｜/box｜>", finish_reason: "stop" },
        { index: 1, text: "第二页", finish_reason: "length" },
    ],
    usage: { input_tokens: 10, output_tokens: 20, total_tokens: 30 },
};

async function setup({ memory = { enabled: true, kbDir: KB }, ocr = {}, hostResolver } = {}) {
    const ctx = fakeCtx();
    const dispose = registerMemoryView(ctx, {
        getMemory: () => memory,
        getConfig: () => ({ search: { sites: {} }, ocr: Object.assign({ enabled: true, baseURL: "https://api.deepseek.com", apiKey: "sk-test", purpose: "text_recognition", timeoutMs: 180000 }, ocr) }),
        ...(hostResolver === undefined ? {} : { resolveHostOcrCredential: hostResolver }),
    });
    const route = ctx.routes.find((item) => item.path === OCR_PATH);
    return { route, dispose };
}

async function call(route, options) {
    const res = makeRes();
    const req = makeReq(options);
    await route.handler(req, res);
    let json = null;
    try { json = JSON.parse(res.body); } catch { json = null; }
    return { status: res.status, json, raw: res.body, req, headers: res.headers || {} };
}

await rm(TMP, { recursive: true, force: true });
await mkdir(TMP, { recursive: true });
resetWorkspaces();
const root = join(TMP, "work");
await mkdir(root, { recursive: true });
const file = join(root, "sample.png");
await writeFile(file, Buffer.from("89504e470d0a1a0a", "hex"));
noteWorkspace(root);
const KB = join(TMP, "kb");
await mkdir(KB, { recursive: true });
const kbFile = join(KB, "scan.png");
await writeFile(kbFile, Buffer.from("89504e470d0a1a0a", "hex"));
const outside = join(TMP, "outside.png");
await writeFile(outside, Buffer.from("89504e470d0a1a0a", "hex"));

await check('OCR 端点：只认 POST + JSON', async () => {
    const { route, dispose } = await setup();
    try {
        const get = await call(route, { method: "GET", url: OCR_PATH + "?cwd=" + encodeURIComponent(root) });
        assert.equal(get.status, 405);
        assert.match(get.json.error, /只支持 POST/);
        const form = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "text/plain" }, body: "{}" });
        assert.equal(form.status, 415);
    } finally { dispose(); }
});

await check('OCR 端点：记忆或文件识别关掉时拒绝，并说清是哪一档', async () => {
    const off = await setup({ memory: { enabled: false } });
    try {
        const res = await call(off.route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 403);
        assert.match(res.json.error, /记忆已在设置里关掉/);
    } finally { off.dispose(); }
    const ocrOff = await setup({ ocr: { enabled: false } });
    try {
        const res = await call(ocrOff.route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 403);
        assert.match(res.json.error, /文件识别已在设置里关掉/);
    } finally { ocrOff.dispose(); }
});

await check('OCR 端点：只读工作目录与知识库目录里的文件', async () => {
    const { route, dispose } = await setup();
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: join(TMP, "outside.png") }) });
        assert.equal(res.status, 403);
        assert.match(res.json.error, /只识别工作目录或知识库目录里的文件/);
        const empty = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: "{}" });
        assert.equal(empty.status, 400);
        assert.match(empty.json.error, /本地文件路径/);
    } finally { dispose(); }
});

await check('OCR 端点：知识库目录里的文件可以识别（正例）', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: kbFile }) });
        assert.equal(res.status, 200, res.raw);
        assert.equal(res.json.ok, true);
        assert.equal(fake.calls.length, 1, "知识库目录在允许根内，应当真的请求上游");
        assert.match(res.json.path, /scan\.png$/);
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：../ 跳出工作目录会被 403 拦住', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: join(root, "..", "outside.png") }) });
        assert.equal(res.status, 403, res.raw);
        assert.equal(fake.calls.length, 0, "越界文件不该被送到上游");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：符号链接指向边界外会被 403 拦住', async () => {
    const link = join(root, "escape.png");
    try { await symlink(outside, link, "file"); }
    catch (error) { console.log("SKIP  符号链接（当前账户无权创建）：" + error.message); return; }
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: link }) });
        assert.equal(res.status, 403, res.raw);
        assert.equal(fake.calls.length, 0, "realpath 之后越界的符号链接不该被送到上游");
    } finally { globalThis.fetch = realFetch; dispose(); await rm(link, { force: true }); }
});

await check('OCR 端点：成功时把逐页文本拼回去，并回传契约字段', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 200, res.raw);
        assert.equal(res.json.ok, true);
        assert.equal(res.json.pageCount, 2);
        assert.equal(res.json.id, "ocr-1");
        assert.equal(res.json.text, "第一页标题\n\n第二页", "路由返回清洗过的 text：版面标记与坐标框不能进知识库正文");
        assert.match(res.json.pages[0].text, /<｜ref｜>/, "pages 原样保留标记，供需要还原版面的人");
        assert.equal(res.json.pages[1].finishReason, "length", "finish_reason 要带出来");
        assert.equal(res.json.usage.total_tokens, 30);
        assert.equal(fake.calls.length, 1);
        assert.match(fake.calls[0].url, /\/v1\/tools\/ocr$/);
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：上游 200 但 body.error 时报 502（不是 400），错误与 code 原样带出', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch({ error: { message: "The OCR service is temporarily unavailable. Please try again later.", type: "server_error", code: "gateway_timeout_error" } });
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        // 上游状态是 200、失败装在 body.error 里：那是上游/网关的问题，不是用户请求写错了，
        // 本机端点必须报 502，让页面分清「我的请求」与「上游」。
        assert.equal(res.status, 502, res.raw);
        assert.equal(res.json.ok, false);
        assert.match(res.json.error, /temporarily unavailable/);
        assert.equal(res.json.code, "gateway_timeout_error");
        assert.equal(res.json.retryable, true);
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：直接带文件内容（data）时落临时文件、不回临时路径（第九十四轮）', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    const tmpBefore = readdirSync(tmpdir()).filter((name) => name.startsWith("dsh-ocr-"));
    try {
        const bytes = Buffer.from("89504e470d0a1a0a", "hex");
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ name: "面板扫描.png", data: bytes.toString("base64") }),
        });
        assert.equal(res.status, 200, res.raw);
        assert.equal(res.json.ok, true);
        assert.equal(fake.calls.length, 1, "内容入口也要真的请求上游");
        assert.equal(res.json.fileName, "面板扫描.png", "文件名用用户给的那个");
        assert.equal("path" in res.json, false, "内容入口不回临时路径（文件随后就删了）");
        assert.equal(res.json.text, "第一页标题\n\n第二页");
        assert.equal(existsSync(join(root, ".office", "runtime", "ocr")), false,
            "内容入口不往工作区写临时文件（私有 OS 临时目录）");
        const tmpAfter = readdirSync(tmpdir()).filter((name) => name.startsWith("dsh-ocr-"));
        assert.deepEqual(tmpAfter, tmpBefore, "识别结束后私有临时目录要删掉");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：内容入口的扩展名与空内容各有明确拒绝（第九十四轮）', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const bad = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ name: "脚本.exe", data: Buffer.from("MZ").toString("base64") }),
        });
        assert.equal(bad.status, 415);
        assert.equal(bad.json.code, "unsupported_format");
        assert.equal(fake.calls.length, 0, "不支持的格式不该被送到上游");

        const empty = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ name: "空.png", data: "" }),
        });
        assert.equal(empty.status, 400);
        assert.match(empty.json.error, /路径|内容/);
        assert.equal(fake.calls.length, 0);
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：文件不存在时报本地原因，不打上游', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: join(root, "nope.png") }) });
        assert.equal(res.status, 400);
        assert.equal(res.json.code, "file_missing");
        assert.equal(fake.calls.length, 0, "本地文件都没读到，不该请求上游");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：请求体超限时报 413 + body_too_large，且不掐断连接（否则页面只能看到连接重置）', async () => {
    const { route, dispose } = await setup();
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        // 比 OCR_BODY_LIMIT（34,603,008 B）大一截：走的是「超限」而不是「JSON 语法错」。
        const huge = JSON.stringify({ name: "大.pdf", data: "x".repeat(35 * 1024 * 1024) });
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: huge,
        });
        assert.equal(res.status, 413, String(res.status) + " / " + String(res.raw).slice(0, 200));
        assert.equal(res.json.ok, false);
        assert.equal(res.json.code, "body_too_large");
        assert.match(res.json.error, /字节/);
        assert.equal(res.json.retryable, false);
        // 不在响应头里立刻关连接：那会让还在写 body 的客户端撞上 ECONNRESET，
        // 已经刷出的 413 反而读不到；真实 socket 上的收尾由 view-ocr-socket.mjs 钉。
        assert.notEqual(String(res.headers.connection || "").toLowerCase(), "close", "413 不能立刻声明 Connection: close");
        // 文案不能自相矛盾：这里确实会读空（有上界）再关，不是「没有截断读取」。
        assert.equal(/没有截断读取/.test(res.json.error), false, "超限文案不能声称没有截断读取");
        assert.match(res.json.error, /这次请求没有被解析/);
        assert.equal(res.req.destroyed, 0, "超限不能 destroy 请求流：浏览器会只拿到 ECONNRESET，状态码与 JSON 都到不了页面");
        assert.ok(res.req.resumed >= 1, "超限后要把剩余字节读空丢掉（resume），而不是把连接掐断");
        assert.equal(fake.calls.length, 0, "请求体都没收完，不该请求上游");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：上游 200 但返回 HTML 时报 502 bad_response，只带状态/类型/分类（不回显页面原文）', async () => {
    const { route, dispose } = await setup();
    const page = "<html><body><h1>502 Bad Gateway</h1><p>C:\\Users\\alice\\.secret\\gw.html</p></body></html>";
    const fake = ocrFetch(null, { status: 200, text: page, contentType: "text/html; charset=utf-8" });
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 502, res.raw);
        assert.equal(res.json.ok, false);
        assert.equal(res.json.code, "bad_response");
        // 端点把上游 body 的分类原样带出：状态码在 JSON 里已经由 HTTP 502 表达，
        // content-type 与 bodyKind 让页面能说清「上游回的是 HTML 而不是识别结果」。
        assert.equal(res.json.contentType, "text/html");
        assert.equal(res.json.bodyKind, "html");
        assert.match(res.json.error, /HTML\/XML/);
        assert.equal(res.raw.includes("Bad Gateway"), false, "响应体原文（可能是网关内部页）不能出现在端点 JSON 里");
        assert.equal(res.raw.includes(".secret"), false, "上游页面里的本机路径不能出现在端点 JSON 里");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('readJsonBody：非 destroy 行为只对显式 drainOverflow 的调用方生效，且有读空上界', async () => {
    // 默认（其它路由沿用的旧语义）：超限立刻 destroy，不读空。
    const legacy = streamReq();
    const legacyFailed = readJsonBody(legacy, 8).then(() => null, (error) => error);
    legacy.emit("data", Buffer.alloc(16));
    const legacyError = await legacyFailed;
    assert.equal(legacyError.status, 413);
    assert.equal(legacyError.code, "body_too_large");
    assert.equal(legacy.destroyed, 1, "未声明 drainOverflow 的调用方保持旧语义：超限即 destroy");
    assert.equal(legacy.resumed, 0, "旧语义不读空");

    // OCR 式调用：超限 still 回 413，但连接不掐断；到 drainCap 后暂停读取（背压）。
    const ocr = streamReq();
    const ocrFailed = readJsonBody(ocr, 8, { drainOverflow: true, drainCap: 32 }).then(() => null, (error) => error);
    ocr.emit("data", Buffer.alloc(16));
    const ocrError = await ocrFailed;
    assert.equal(ocrError.status, 413);
    assert.equal(ocr.destroyed, 0, "drainOverflow：先回 413，不把浏览器打成 connection reset");
    assert.equal(ocr.resumed, 1, "drainOverflow：把剩余字节读空丢掉");
    ocr.emit("data", Buffer.alloc(64));
    assert.equal(ocr.paused, 1, "读空也有上界：超过 drainCap 就暂停读取，不无限灌数据");
    assert.equal(ocr.destroyed, 0, "暂停而不是掐断：413 已经发出，响应结束后由 Node 关闭连接");
});

await check('OCR 端点：上游网络不通报 502 / 超时报 504（不再一律压成 400）', async () => {
    const realFetch = globalThis.fetch;
    const net = await setup();
    globalThis.fetch = async () => { throw new Error("connect ECONNREFUSED 127.0.0.1:9"); };
    try {
        const res = await call(net.route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 502, res.raw);
        assert.equal(res.json.code, "network");
        assert.equal(res.json.retryable, true);
    } finally { net.dispose(); }
    // 超时：让假 fetch 在 AbortController 触发时才 reject，走的才是 ocr.js 的 timeout 分支。
    const slow = await setup({ ocr: { timeoutMs: 25 } });
    globalThis.fetch = (url, init) => new Promise((settle, fail) => {
        const abort = () => fail(new Error("This operation was aborted"));
        if (init && init.signal) {
            if (init.signal.aborted) abort();
            else init.signal.addEventListener("abort", abort);
        }
    });
    try {
        const res = await call(slow.route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 504, res.raw);
        assert.equal(res.json.code, "timeout");
        assert.equal(res.json.retryable, true);
    } finally { globalThis.fetch = realFetch; slow.dispose(); }
});

await check('OCR 端点：上游 headers 到了但 body 卡住时报 504 timeout（超时管到响应读完）', async () => {
    const { route, dispose } = await setup({ ocr: { timeoutMs: 40 } });
    const realFetch = globalThis.fetch;
    globalThis.fetch = async (url, init) => ({
        ok: true, status: 200,
        headers: { get: () => "application/json" },
        text: () => new Promise((resolve, reject) => {
            const abort = () => reject(Object.assign(new Error("The operation was aborted"), { name: "AbortError" }));
            if (init && init.signal) {
                if (init.signal.aborted) abort();
                else init.signal.addEventListener("abort", abort);
            }
        }),
    });
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 504, res.raw);
        assert.equal(res.json.code, "timeout");
        assert.equal(res.json.retryable, true);
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：上游 body 中途断开设 502 body_interrupted，保留状态且不回显异常文字', async () => {
    const { route, dispose } = await setup();
    const realFetch = globalThis.fetch;
    globalThis.fetch = async () => ({
        ok: true, status: 200,
        headers: { get: () => "application/json" },
        text: () => Promise.reject(Object.assign(new Error("read ECONNRESET at /home/alice/secret"), { code: "ECONNRESET" })),
    });
    try {
        const res = await call(route, { url: OCR_PATH + "?cwd=" + encodeURIComponent(root), headers: { "content-type": "application/json" }, body: JSON.stringify({ path: file }) });
        assert.equal(res.status, 502, res.raw);
        assert.equal(res.json.code, "body_interrupted");
        assert.equal(res.json.retryable, true);
        assert.match(res.json.error, /ECONNRESET/);
        assert.equal(res.raw.includes("/home/alice"), false, "异常里的本机路径不能出现在端点 JSON 里");
        assert.equal(res.raw.includes("read ECONNRESET"), false, "异常 message 原样不引用");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：data 入口在上游失败时也必须回结构化 JSON（脱敏引用不能越界）', async () => {
    // 事故：fileName 原本声明在下面的 try 里，catch 的 scrub 在临时文件已建立时
    // 引用它 → ReferenceError 逃出 handler → 宿主兜底裸 400、空 body。这里用空
    // Key 让 ocrDocument 在建立临时文件之后失败，钉住「错误仍以 JSON 回话」。
    const { route, dispose } = await setup({ ocr: { apiKey: "" } });
    try {
        const png = Buffer.from("89504e470d0a1a0a", "hex").toString("base64");
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ name: "sample.png", data: png }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.match(res.json.error, /API Key/);
        assert.match(res.json.hint, /文件识别/);
    } finally { dispose(); }
});

await check('OCR 端点：缺 Key 时把设置里的环境变量名带进提示，请求体不参与（第九十五轮）', async () => {
    const { route, dispose } = await setup({ ocr: { apiKey: "", apiKeyEnv: "OFFICE_ROUTE_OCR_KEY" } });
    const realFetch = globalThis.fetch;
    let called = 0;
    globalThis.fetch = async () => { called += 1; throw new Error("不该请求上游"); };
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            // 请求体里故意塞一个变量名：它绝不能出现在诊断文案里（只信设置）。
            body: JSON.stringify({ path: file, apiKeyEnv: "FROM_REQUEST_MUST_BE_IGNORED" }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.match(res.json.error, /OFFICE_ROUTE_OCR_KEY/, "要点出设置里配的那个名字");
        assert.equal(res.json.error.includes("FROM_REQUEST_MUST_BE_IGNORED"), false, "环境变量名只信设置，不信请求体");
        assert.match(res.json.hint, /设置 → 办公模式 → 文件识别（OCR）/);
        assert.match(res.json.hint, /那个进程继承的环境变量/);
        assert.match(res.json.hint, /启动 DSH 的那个终端/);
        assert.match(res.json.hint, /重启 DSH/);
        assert.equal(called, 0, "缺 Key 是本地判断，不该请求上游");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：设置里的变量名不合法时不回显（误粘的 Key 不能被打出来）（第九十五轮）', async () => {
    const pasted = "sk-live-0123456789abcdef";
    const { route, dispose } = await setup({ ocr: { apiKey: "", apiKeyEnv: pasted } });
    const realFetch = globalThis.fetch;
    let called = 0;
    globalThis.fetch = async () => { called += 1; throw new Error("不该请求上游"); };
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.equal(res.raw.includes(pasted), false, "整个响应体里都不能出现那段疑似 Key 的字符串");
        assert.match(res.json.error, /不是合法变量名/);
        assert.match(res.json.hint, /文件识别（OCR）/);
        assert.equal(called, 0, "缺 Key 是本地判断，不该请求上游");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

// ── 宿主官方供应商 Key（第九十七轮） ─────────────────────────────────────────
// 只注入假 resolver，不读真实 config / credentials。

await check('OCR 端点：自定义 baseURL 不调用宿主 Key，显式 Key 优先且不带 redirect', async () => {
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    let resolverCalls = 0;
    globalThis.fetch = fake.impl;
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://ocr-proxy.example.com", apiKey: "sk-explicit" },
        hostResolver: async () => { resolverCalls += 1; return { ok: true, apiKey: "sk-host" }; },
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 200, res.raw);
        assert.equal(resolverCalls, 0, "自定义端点连宿主 Key 都不该看");
        assert.equal(fake.calls[0].url, "https://ocr-proxy.example.com/v1/tools/ocr");
        assert.equal(fake.calls[0].init.headers.authorization, "Bearer sk-explicit");
        assert.equal(fake.calls[0].init.redirect, undefined, "显式 Key 保持旧行为");
        assert.equal(res.raw.includes("sk-explicit"), false, "Key 不能出现在响应里");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：官方地址 + 空 Key 时复用宿主 Key，端点钉死并拒绝重定向', async () => {
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    const seenOptions = [];
    globalThis.fetch = fake.impl;
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://api.deepseek.com/v1", apiKey: "", reuseOfficialProvider: true },
        hostResolver: async (options) => {
            seenOptions.push(options);
            return { ok: true, apiKey: "sk-host-secret", ref: "DEEPSEEK_API_KEY" };
        },
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 200, res.raw);
        assert.equal(seenOptions.length, 1);
        assert.equal(seenOptions[0].baseURL, "https://api.deepseek.com/v1");
        assert.equal(fake.calls[0].url, "https://api.deepseek.com/v1/tools/ocr");
        assert.equal(fake.calls[0].init.headers.authorization, "Bearer sk-host-secret");
        assert.equal(fake.calls[0].init.redirect, "error");
        assert.equal(res.raw.includes("sk-host-secret"), false, "解析出的 Key 不能回给浏览器");
        assert.equal(res.raw.includes("DEEPSEEK_API_KEY"), false, "凭据引用也不必回给浏览器");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：官方地址 + 空 Key + 自定义端点配置时 resolver 不调用，缺 Key 不提 Models', async () => {
    const realFetch = globalThis.fetch;
    let resolverCalls = 0;
    globalThis.fetch = async () => { throw new Error("不该请求"); };
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://ocr-proxy.example.com", apiKey: "", reuseOfficialProvider: true },
        hostResolver: async () => { resolverCalls += 1; return { ok: true, apiKey: "sk-host-secret" }; },
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.equal(resolverCalls, 0, "自定义端点必须先被 officialOcrEndpointOf 拦住");
        assert.equal(/Models/.test(res.json.hint), false);
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：官方地址 + 空 Key + 供应商已配置但没取到 Key 时，提示 Models', async () => {
    const realFetch = globalThis.fetch;
    globalThis.fetch = async () => { throw new Error("不该请求"); };
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://api.deepseek.com", apiKey: "", reuseOfficialProvider: true },
        hostResolver: async () => ({ ok: false, eligible: true, reason: "credential_missing" }),
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.match(res.json.hint, /设置 → 模型（Models）/);
        assert.match(res.json.hint, /文件识别（OCR）/, "旧的环境变量提示还在");
        assert.equal(res.raw.includes("credential_missing"), false, "内部 reason 不能回给浏览器");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：resolver 抛异常只按缺 Key 处理，异常文字不外泄', async () => {
    const realFetch = globalThis.fetch;
    globalThis.fetch = async () => { throw new Error("不该请求"); };
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://api.deepseek.com", apiKey: "", reuseOfficialProvider: true },
        hostResolver: async () => { throw new Error("read /home/alice/secret failed"); },
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.equal(res.raw.includes("/home/alice"), false);
        assert.equal(res.raw.includes("secret failed"), false);
        assert.equal(/Models/.test(res.json.hint), false, "借 Key 失败不等于确认供应商可复用");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

// ── 复用开关 + 本机直连守卫（第九十八轮） ───────────────────────────────────

await check('OCR 端点：复用开关默认关闭时，官方地址也不调用 resolver，并提示怎么打开', async () => {
    const realFetch = globalThis.fetch;
    let resolverCalls = 0;
    globalThis.fetch = async () => { throw new Error("不该请求"); };
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://api.deepseek.com", apiKey: "" },
        hostResolver: async () => { resolverCalls += 1; return { ok: true, apiKey: "sk-host" }; },
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.equal(resolverCalls, 0, "设置没同意就不能碰宿主 Key");
        assert.match(res.json.hint, /复用官方供应商 Key/);
        assert.match(res.json.hint, /打开/);
        assert.equal(/已配置|已经保存/.test(res.json.hint), false, "没同意时不要暗示供应商已配置");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：非本机 remoteAddress 不调用 resolver，提示未通过本机来源检查', async () => {
    const realFetch = globalThis.fetch;
    let resolverCalls = 0;
    globalThis.fetch = async () => { throw new Error("不该请求"); };
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://api.deepseek.com", apiKey: "", reuseOfficialProvider: true },
        hostResolver: async () => { resolverCalls += 1; return { ok: true, apiKey: "sk-host" }; },
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            socket: { remoteAddress: "192.168.1.20", localPort: 3080, encrypted: false },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 400, res.raw);
        assert.equal(res.json.code, "missing_api_key");
        assert.equal(resolverCalls, 0, "LAN 客户端不能借用本机 Models 的 Key");
        assert.match(res.json.hint, /这次请求未通过本机来源检查/);
        assert.match(res.json.hint, /不是身份认证/);
        assert.equal(/Models/.test(res.json.hint), false, "非本机请求不能看到供应商配置状态");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：代理转发标记 / 跨站 Sec-Fetch-Site / 缺 Origin / 非本机 Host 都不借', async () => {
    const realFetch = globalThis.fetch;
    let resolverCalls = 0;
    globalThis.fetch = async () => { throw new Error("不该请求"); };
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://api.deepseek.com", apiKey: "", reuseOfficialProvider: true },
        hostResolver: async () => { resolverCalls += 1; return { ok: true, apiKey: "sk-host" }; },
    });
    const cases = [
        { name: "x-forwarded-for", options: { headers: { "content-type": "application/json", "x-forwarded-for": "203.0.113.5" } } },
        { name: "forwarded", options: { headers: { "content-type": "application/json", forwarded: "for=203.0.113.5" } } },
        { name: "cross-site", options: { headers: { "content-type": "application/json", "sec-fetch-site": "cross-site" } } },
        { name: "missing origin", options: { headers: { "content-type": "application/json" }, origin: null } },
        { name: "foreign host", options: { host: "evil.example:3080", origin: "http://evil.example:3080", headers: { "content-type": "application/json" } } },
    ];
    try {
        for (const item of cases) {
            const res = await call(route, {
                url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
                body: JSON.stringify({ path: file }),
                ...item.options,
            });
            assert.equal(res.status, 400, item.name + ": " + res.raw);
            assert.equal(res.json.code, "missing_api_key", item.name);
        }
        assert.equal(resolverCalls, 0, "任何一条元数据不合格都不能读宿主 Key");
    } finally { globalThis.fetch = realFetch; dispose(); }
});

await check('OCR 端点：显式 Key 的行为不受本机守卫影响（非本机也能照旧识别）', async () => {
    const fake = ocrFetch(SUCCESS);
    const realFetch = globalThis.fetch;
    globalThis.fetch = fake.impl;
    const { route, dispose } = await setup({
        ocr: { baseURL: "https://api.deepseek.com", apiKey: "sk-explicit", reuseOfficialProvider: true },
        hostResolver: async () => { throw new Error("显式 Key 下不该调用"); },
    });
    try {
        const res = await call(route, {
            url: OCR_PATH + "?cwd=" + encodeURIComponent(root),
            headers: { "content-type": "application/json" },
            socket: { remoteAddress: "192.168.1.20", localPort: 3080, encrypted: false },
            body: JSON.stringify({ path: file }),
        });
        assert.equal(res.status, 200, res.raw);
        assert.equal(fake.calls[0].init.headers.authorization, "Bearer sk-explicit");
        assert.equal(fake.calls[0].init.redirect, undefined);
    } finally { globalThis.fetch = realFetch; dispose(); }
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log((item.ok ? "PASS  " : "FAIL  ") + item.name);
    if (!item.ok) console.log("      " + (item.error && item.error.message ? item.error.message : item.error));
}
console.log("view-ocr: " + (results.length - failed.length) + "/" + results.length + " 通过");
process.exit(failed.length === 0 ? 0 : 1);