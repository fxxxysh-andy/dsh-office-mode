/**
 * 知识库上传 413 的真实 loopback socket 测试（**不进 CI**：CI 套件禁真实网络/服务器）。
 *
 * 背景（第九十八轮）：只含 pad 字段、略超 12 MiB 的 JSON POST 到 /office-memory/kb-upload
 * 以前会让服务端在读到超限时直接 destroy，浏览器只拿到 UND_ERR_SOCKET / fetch failed。
 * 现在 KB 复用 OCR 已经验证过的「结构化 413 + 有界关闭」：
 *   1. 略超 12 MiB（12 MiB + 64 KiB）分块上传：413 在请求体结束前就刷出、JSON 可解析，
 *      不带 Connection: close；请求体写完后服务端优雅关闭，客户端没有 ECONNRESET。
 *   2. 客户端收到 413 后停住不写完：服务端在 ~1.5s 安全截止后关闭，413 仍可读。
 *
 * 客户端故意用 node:net 手写 HTTP 请求（理由同 view-ocr-socket.mjs）：更贴近浏览器/代理
 * 往裸连接里灌字节的真实形态，也能精确记录「响应到达时已经写了多少」。
 *
 * 跑法：node test/kb-upload-socket.mjs
 */
import assert from 'node:assert/strict';
import { mkdir, rm } from 'node:fs/promises';
import { createServer as createHttpServer } from 'node:http';
import { connect } from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { registerMemoryView, KB_UPLOAD_PATH, noteWorkspace, resetWorkspaces } from '../src/view.js';

const MiB = 1024 * 1024;
const KB_LIMIT = 12 * MiB;
const results = [];
async function check(name, fn) {
    try { await fn(); results.push({ name, ok: true }); }
    catch (error) { results.push({ name, ok: false, error }); }
}

const root = join(tmpdir(), 'dsh-kb-upload-socket-' + process.pid);
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
    getMemory: () => ({ enabled: true, kbDir: join(root, 'kb') }),
});
const route = routes.find((item) => item.path === KB_UPLOAD_PATH);
assert.ok(route, '知识库上传路由要注册出来');

// 超限请求不该碰上游（模型/OCR/文件系统都不该被叫醒）。
const realFetch = globalThis.fetch;
const upstreamCalls = [];
globalThis.fetch = async (url) => { upstreamCalls.push(String(url)); throw new Error('should not reach upstream'); };

const sockets = new Set();
const server = createHttpServer((req, res) => { route.handler(req, res).catch(() => {}); });
server.on('connection', (socket) => { sockets.add(socket); socket.on('close', () => sockets.delete(socket)); });
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const port = server.address().port;

/**
 * 手写 HTTP/1.1 POST，body 是「只含 pad 字段」的 JSON：{ "pad": "xxx…" }。
 * 分块流式发送，只持有 64 KiB 块；content-length 声明真实总长。
 * stopAfterResponse：收到响应头后立刻停写（模拟「浏览器还卡着」）。
 */
function sendStreaming({ padBytes, stopAfterResponse = false, chunkBytes = 64 * 1024 }) {
    const prefix = Buffer.from('{"pad":"', 'utf8');
    const suffix = Buffer.from('"}', 'utf8');
    const totalBytes = prefix.length + padBytes + suffix.length;
    return new Promise((resolve, reject) => {
        const chunk = Buffer.alloc(chunkBytes, 0x78);
        const state = {
            sent: 0, sentAtResponse: -1, status: 0, body: '',
            responseEnded: false, closed: false, errorCode: '',
            responseAt: 0, closeAfterResponseMs: -1, stoppedWriting: false,
            connection: '', totalBytes,
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
        const head = 'POST ' + KB_UPLOAD_PATH + ' HTTP/1.1\r\n'
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
            let padSent = 0;
            const screen = (buf) => {
                if (stopAfterResponse && state.status !== 0) { state.stoppedWriting = true; return false; }
                state.sent += buf.length;
                return true;
            };
            if (!screen(prefix)) return;
            if (!socket.write(prefix)) await new Promise((r) => socket.once('drain', r));
            while (padSent < padBytes) {
                if (stopAfterResponse && state.status !== 0) { state.stoppedWriting = true; return; }
                const remaining = padBytes - padSent;
                const buf = remaining >= chunkBytes ? chunk : chunk.subarray(0, remaining);
                padSent += buf.length;
                if (!screen(buf)) return;
                if (!socket.write(buf)) await new Promise((r) => socket.once('drain', r));
                else await new Promise((r) => setImmediate(r));
            }
            if (!screen(suffix)) return;
            socket.write(suffix);
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

await check('略超 12 MiB 的 KB JSON：413 先刷出（在请求体结束前），写完请求体后优雅关闭、无 ECONNRESET', async () => {
    const padBytes = KB_LIMIT - 16 + 64 * 1024; // 总长约 12 MiB + 64 KiB
    const result = await sendStreaming({ padBytes });
    assert.equal(result.status, 413, result.errorCode || '没有拿到 413');
    const payload = JSON.parse(result.body);
    assert.equal(payload.ok, false);
    assert.equal(payload.code, 'body_too_large');
    assert.match(payload.error, /12582912/);
    assert.match(payload.error, /这次请求没有被解析/);
    assert.equal(payload.retryable, false);
    assert.ok(result.body.length < 1024, '413 响应体要小（不回显请求）');
    assert.ok(result.sentAtResponse <= KB_LIMIT + 256 * 1024,
        '响应应在上限附近就到达（sentAtResponse=' + result.sentAtResponse + '）');
    assert.notEqual(result.connection, 'close', '413 不能带 Connection: close（还在写 body 的客户端会 ECONNRESET）');
    assert.equal(result.responseEnded, true, '413 body 要完整读得到');
    assert.equal(result.errorCode, '', '客户端不能收到 ECONNRESET 之类的传输错误：' + result.errorCode);
    assert.equal(result.closed, true);
    assert.ok(result.closeAfterResponseMs >= 0 && result.closeAfterResponseMs < 3000,
        '请求体结束后应在 3s 内关闭（实际 ' + result.closeAfterResponseMs + 'ms）');
    assert.equal(await waitClosed(), true, '服务端 socket 要真的关掉');
    assert.equal(upstreamCalls.length, 0, '超限请求不该打到上游');
});

await check('客户端收到 413 后停住：~1.5s 安全截止后关闭，413 仍可读', async () => {
    const padBytes = KB_LIMIT + MiB; // 总长约 13 MiB
    const result = await sendStreaming({ padBytes, stopAfterResponse: true });
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
console.log('kb-upload-socket: ' + (results.length - failed) + '/' + results.length + ' 通过');
process.exitCode = failed === 0 ? 0 : 1;
