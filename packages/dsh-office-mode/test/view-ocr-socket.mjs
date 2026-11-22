/**
 * OCR 413 的真实 loopback socket 测试（**不进 CI**：CI 套件禁真实网络/服务器）。
 *
 * 背景：reviewer 用 40/60 MB 复现过——413 响应带立即的 Connection: close 时，客户端
 * 还在写 body 就会撞上 ECONNRESET，连已经发出的 413 都读不到。去掉立即 close 后，
 * 36/60/67/75 MB 都能收到结构化 413，并在请求体结束或安全截止后关闭。
 *
 * 这里用 127.0.0.1:0 的临时 http server + 真实 TCP 连接（不是宿主的 3080），钉住：
 *   1. 60 MiB（> 50 MiB）分块流式上传：413 在 body 收完之前就刷出、JSON 可解析；
 *      客户端把 body 写完后，服务端在 req end 之后优雅关闭，客户端没有 ECONNRESET。
 *   2. 客户端收到 413 后停住不写完：服务端在 ~1.5s 安全截止后关闭，413 仍可读。
 *   3. 无整包保留：两个用例都只持有 64 KiB 块；响应到达时客户端才写了 ~limit 量级
 *      的字节（若服务端等整包收完才回，sentAtResponse 会接近 60 MiB）。
 *
 * 客户端故意用 node:net 手写 HTTP 请求，而不是 http.request：后者在响应先于请求体
 * 到达时，最后一次背压的 'drain' 会丢（writableLength=0 却不发 drain），把测试卡死；
 * 手写 socket 也更贴近「浏览器/代理往一个裸连接里灌字节」的真实形态。
 *
 * 跑法：node test/view-ocr-socket.mjs
 */
import assert from 'node:assert/strict';
import { mkdir, rm } from 'node:fs/promises';
import { createServer as createHttpServer } from 'node:http';
import { connect } from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { registerMemoryView, OCR_PATH, noteWorkspace, resetWorkspaces } from '../src/view.js';

const MiB = 1024 * 1024;
const results = [];
async function check(name, fn) {
    try { await fn(); results.push({ name, ok: true }); }
    catch (error) { results.push({ name, ok: false, error }); }
}

const root = join(tmpdir(), 'dsh-ocr-socket-' + process.pid);
await mkdir(root, { recursive: true });
resetWorkspaces();
noteWorkspace(root);

const routes = [];
const ctx = {
    get(name) {
        if (name === 'webServer') return { register(route) { routes.push(route); return () => {}; } };
        return undefined;
    },
};
const dispose = registerMemoryView(ctx, {
    getMemory: () => ({ enabled: true }),
    getConfig: () => ({
        search: { sites: {} },
        ocr: { enabled: true, baseURL: 'https://api.deepseek.com', apiKey: 'sk-test', purpose: 'text_recognition', timeoutMs: 5000 },
    }),
});
const route = routes.find((item) => item.path === OCR_PATH);
assert.ok(route, 'OCR 路由要注册出来');

// 上游不该被调用：超限请求在本地就被 413 拦下。
const realFetch = globalThis.fetch;
const upstreamCalls = [];
globalThis.fetch = async (url) => { upstreamCalls.push(String(url)); throw new Error('should not reach upstream'); };

const sockets = new Set();
const server = createHttpServer((req, res) => { route.handler(req, res).catch(() => {}); });
server.on('connection', (socket) => { sockets.add(socket); socket.on('close', () => sockets.delete(socket)); });
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const port = server.address().port;

/**
 * 手写 HTTP/1.1 POST：分块流式发 content-length 声明为 totalBytes 的 body，只持有
 * 64 KiB 块。stopAfterResponse：收到响应头后立刻停写（模拟「客户端还卡着」）。
 */
function sendStreaming({ totalBytes, stopAfterResponse = false, chunkBytes = 64 * 1024 }) {
    return new Promise((resolve, reject) => {
        const chunk = Buffer.alloc(chunkBytes, 0x78);
        const state = {
            sent: 0, sentAtResponse: -1, status: 0, body: '',
            responseEnded: false, closed: false, errorCode: '',
            responseAt: 0, closeAfterResponseMs: -1, stoppedWriting: false,
            connection: '',
        };
        let pendingHeaders = '';
        let headerEnd = -1;
        let remainingBody = -1;
        const bodyParts = [];
        const socket = connect(port, '127.0.0.1');
        const timer = setTimeout(() => {
            state.errorCode = state.errorCode || 'test_timeout';
            socket.destroy();
        }, 10000);
        timer.unref?.();
        socket.setNoDelay?.(true);
        const head = 'POST ' + OCR_PATH + ' HTTP/1.1\r\n'
            + 'Host: 127.0.0.1:' + port + '\r\n'
            + 'Content-Type: application/json\r\n'
            + 'Content-Length: ' + totalBytes + '\r\n'
            + '\r\n';
        socket.on('connect', () => { socket.write(head); void pump(); });
        socket.on('data', (part) => {
            if (state.status === 0) {
                pendingHeaders += part.toString('latin1');
                headerEnd = pendingHeaders.indexOf('\r\n\r\n');
                if (headerEnd === -1) return;
                const headerText = pendingHeaders.slice(0, headerEnd);
                state.status = Number((headerText.match(/^HTTP\/1\.[01] (\d{3})/) || [])[1] || 0);
                state.sentAtResponse = state.sent;
                state.responseAt = Date.now();
                state.connection = String((headerText.match(/^connection:\s*(.+)$/im) || [])[1] || '').trim().toLowerCase();
                const lengthMatch = headerText.match(/^content-length:\s*(\d+)/im);
                remainingBody = lengthMatch ? Number(lengthMatch[1]) : -1;
                const rest = Buffer.from(pendingHeaders.slice(headerEnd + 4), 'latin1');
                pendingHeaders = '';
                appendBody(rest);
                return;
            }
            appendBody(part);
        });
        function appendBody(part) {
            if (remainingBody < 0) { bodyParts.push(part); return; }
            const take = part.subarray(0, Math.max(0, remainingBody));
            bodyParts.push(take);
            remainingBody -= take.length;
            if (remainingBody <= 0) state.responseEnded = true;
        }
        socket.on('error', (error) => { state.errorCode = String(error && (error.code || error.message) || ''); });
        socket.on('close', () => {
            clearTimeout(timer);
            state.closed = true;
            state.body = Buffer.concat(bodyParts).toString('utf8');
            state.closeAfterResponseMs = state.responseAt === 0 ? -1 : Date.now() - state.responseAt;
            resolve(state);
        });
        const pump = async () => {
            while (state.sent < totalBytes) {
                if (stopAfterResponse && state.status !== 0) { state.stoppedWriting = true; return; }
                const remaining = totalBytes - state.sent;
                const buf = remaining >= chunkBytes ? chunk : chunk.subarray(0, remaining);
                state.sent += buf.length;
                if (!socket.write(buf)) await new Promise((r) => socket.once('drain', r));
                else await new Promise((r) => setImmediate(r));
            }
            socket.end();
        };
    });
}

async function waitClosed(timeoutMs = 3000) {
    const deadline = Date.now() + timeoutMs;
    while (sockets.size > 0 && Date.now() < deadline) {
        await new Promise((resolve) => setTimeout(resolve, 20));
    }
    return sockets.size === 0;
}

await check('60 MiB 流式上传：413 先刷出（不等整包），请求体结束后优雅关闭、无 ECONNRESET', async () => {
    const total = 60 * MiB;
    const result = await sendStreaming({ totalBytes: total });
    assert.equal(result.status, 413, result.errorCode || '没有拿到 413');
    const payload = JSON.parse(result.body);
    assert.equal(payload.code, 'body_too_large');
    assert.match(payload.error, /这次请求没有被解析/);
    assert.equal(/没有截断读取/.test(payload.error), false, '文案不能自相矛盾');
    assert.ok(result.body.length < 1024, '413 响应体要小（不回显请求）');
    // 响应先于整包：服务端在 limit 附近就判定，不需要把 60 MiB 收完/留在内存里。
    assert.ok(result.sentAtResponse < total - 8 * MiB,
        '响应到达时客户端应远未写完（sentAtResponse=' + result.sentAtResponse + '）');
    assert.notEqual(result.connection, 'close', '413 不能带 Connection: close（还在写 body 的客户端会 ECONNRESET）');
    assert.equal(result.responseEnded, true, '413 body 要完整读得到');
    assert.equal(result.errorCode, '', '客户端不能收到 ECONNRESET 之类的传输错误：' + result.errorCode);
    assert.equal(result.closed, true);
    assert.ok(result.closeAfterResponseMs >= 0 && result.closeAfterResponseMs < 3000,
        '请求体结束后应在 3s 内关闭（实际 ' + result.closeAfterResponseMs + 'ms）');
    assert.equal(await waitClosed(), true, '服务端 socket 要真的关掉');
    // 「不整包保留」由 sentAtResponse 钉住：服务端在 limit 附近就做了决定并回了响应，
    // 不需要、也没有等 60 MiB 收完。客户端也只持有 64 KiB 的块。
    assert.equal(upstreamCalls.length, 0, '超限请求不该打到上游');
});

await check('客户端收到 413 后停住：~1.5s 安全截止后关闭，413 仍可读', async () => {
    const total = 60 * MiB;
    const result = await sendStreaming({ totalBytes: total, stopAfterResponse: true });
    assert.equal(result.status, 413, result.errorCode || '没有拿到 413');
    assert.equal(JSON.parse(result.body).code, 'body_too_large', '截止关闭前 413 必须已经可读');
    assert.equal(result.stoppedWriting, true);
    assert.equal(result.closed, true, '截止后连接要关闭');
    assert.ok(result.closeAfterResponseMs >= 1200 && result.closeAfterResponseMs < 4000,
        '应在 ~1.5s 截止后关闭（实际 ' + result.closeAfterResponseMs + 'ms）');
    assert.equal(await waitClosed(), true, '服务端 socket 要真的关掉');
    // 强制销毁后客户端可能看到 ECONNRESET —— 这时 413 已经读到，允许。
});

try {
    for (const socket of sockets) { try { socket.destroy(); } catch { /* 已经断开 */ } }
    await new Promise((resolve) => server.close(resolve));
} finally {
    globalThis.fetch = realFetch;
    dispose();
    await rm(root, { recursive: true, force: true });
}

let failed = 0;
for (const item of results) {
    console.log((item.ok ? 'PASS  ' : 'FAIL  ') + item.name);
    if (!item.ok) {
        failed += 1;
        console.log('      ' + (item.error && item.error.stack ? item.error.stack : String(item.error)));
    }
}
console.log('');
console.log('view-ocr-socket: ' + (results.length - failed) + '/' + results.length + ' 通过');
process.exitCode = failed === 0 ? 0 : 1;
