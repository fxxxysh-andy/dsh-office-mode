/**
 * 本地 embedding 进程（第一百零三轮）。
 *
 * 目标：不配任何远端端点也能有真正的语义向量，代价可控 —— **按需启动**、空闲自动
 * 退出、峰值内存有硬上限。插件本身不实现推理，只负责起一个本机服务，然后把它当
 * 「OpenAI 兼容的远端 embedding 通道」用（POST /v1/embeddings），这样向量化、
 * 降级、去重这些逻辑都只有一份（见 kb-embed.js）。
 *
 * 为什么是 Python + transformers 而不是 llama.cpp：本机（Windows）没有
 * llama-server 二进制，GitHub 不可达；Python 3.13 + torch + transformers 已就绪。
 * 实测 bge-small-zh-v1.5（512 维）在 batch8 / maxlen256 / 2 线程下峰值 RSS 约
 * 480–515MB，低于 600MB 预算；Qwen3-Embedding-0.6B 即使 q4 也要 350MB 权重
 * 再加 KV cache，必然超预算，所以默认模型换成小模型。
 *
 * 失败路径：缺解释器 / 缺模型 / 启动超时 / 健康检查失败 / 内存超限，都返回
 * { ok:false, error }（写清缺什么），由调用方退回词法检索并标 degraded。
 *
 * @module dsh-office-mode/kb-local-embed
 */
import { spawn as nodeSpawn, execFile as nodeExecFile } from 'node:child_process';
import { existsSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { asText } from './web-errors.js';

export const LOCAL_EMBEDDING_MODEL = 'bge-small-zh-v1.5';
export const LOCAL_EMBEDDING_DIMENSIONS = 512;
export const LOCAL_EMBED_READY = 'DSH_LOCAL_EMBED_READY ';
export const SERVER_SCRIPT = fileURLToPath(new URL('./kb-local-embed-server.py', import.meta.url));

export const DEFAULT_LOCAL_EMBEDDING = {
    enabled: false,
    command: '',
    modelDir: '',
    idleTimeoutMs: 120_000,
    maxRssMb: 600,
    startupTimeoutMs: 60_000,
    maxLength: 256,
    threads: 2,
};

/** 收口配置（坏值落默认，绝不抛）。 */
export function normalizeLocalEmbeddingConfig(raw = {}) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const clamp = (value, fallback, min, max) => (Number.isFinite(value) ? Math.min(max, Math.max(min, Math.trunc(value))) : fallback);
    return {
        enabled: source.enabled === true,
        command: asText(source.command),
        modelDir: asText(source.modelDir),
        idleTimeoutMs: clamp(source.idleTimeoutMs, DEFAULT_LOCAL_EMBEDDING.idleTimeoutMs, 5_000, 3_600_000),
        maxRssMb: clamp(source.maxRssMb, DEFAULT_LOCAL_EMBEDDING.maxRssMb, 200, 4_096),
        startupTimeoutMs: clamp(source.startupTimeoutMs, DEFAULT_LOCAL_EMBEDDING.startupTimeoutMs, 5_000, 300_000),
        maxLength: clamp(source.maxLength, DEFAULT_LOCAL_EMBEDDING.maxLength, 64, 512),
        threads: clamp(source.threads, DEFAULT_LOCAL_EMBEDDING.threads, 1, 16),
    };
}

/** 默认模型目录：<知识库目录>/models/bge-small-zh-v1.5（与 kb-model 的下载约定一致）。 */
export function defaultLocalModelDir(kbDir) {
    const root = asText(kbDir);
    return root === '' ? '' : `${root.replace(/[\\/]+$/, '')}/models/${LOCAL_EMBEDDING_MODEL}`;
}

/** 模型目录是否像一份能加载的权重：config.json + 权重 + 分词器，缺什么报什么。 */
export function checkModelDir(dir, fsx = { existsSync, statSync }) {
    const root = asText(dir);
    if (root === '') return { ok: false, error: '没有配置本地模型目录，也没有默认知识库目录可用。' };
    if (!fsx.existsSync(root)) return { ok: false, error: `模型目录不存在：${root}。在设置页「知识库模型」里点下载，或把 ${LOCAL_EMBEDDING_MODEL} 放进去。` };
    try {
        if (!fsx.statSync(root).isDirectory()) return { ok: false, error: `模型路径不是目录：${root}。` };
    } catch {
        return { ok: false, error: `模型目录不可读：${root}。` };
    }
    const has = (name) => fsx.existsSync(`${root}/${name}`);
    if (!has('config.json')) return { ok: false, error: `模型目录缺 config.json：${root}。` };
    if (!has('model.safetensors') && !has('pytorch_model.bin')) {
        return { ok: false, error: `模型目录缺权重（model.safetensors / pytorch_model.bin）：${root}。` };
    }
    if (!has('tokenizer.json') && !has('vocab.txt')) return { ok: false, error: `模型目录缺分词器（tokenizer.json / vocab.txt）：${root}。` };
    return { ok: true, error: '' };
}

function sleep(ms) {
    return new Promise((resolve) => { setTimeout(resolve, ms); });
}

function execFileP(file, args) {
    return new Promise((resolve, reject) => {
        nodeExecFile(file, args, { windowsHide: true, timeout: 15_000 }, (error, stdout) => {
            if (error) reject(error);
            else resolve(String(stdout ?? ''));
        });
    });
}

/**
 * 读子进程峰值内存（MB）。Windows 走 tasklist，POSIX 走 ps —— 不用
 * process.memoryUsage()，那是本进程的堆，不是子进程占用。
 */
export async function readProcessRssMb(pid) {
    if (!Number.isFinite(pid) || pid <= 0) return 0;
    if (process.platform === 'win32') {
        const out = await execFileP('tasklist', ['/FI', `PID eq ${pid}`, '/FO', 'CSV', '/NH']);
        // "python.exe","1234","Console","1","123,456 K"
        const match = /,"([\d., ]*)\s*K"/.exec(out);
        if (match === null) return 0;
        return Math.round(Number.parseFloat(match[1].replace(/[.,]/g, '')) / 1024);
    }
    const out = await execFileP('ps', ['-o', 'rss=', '-p', String(pid)]);
    const kb = Number.parseFloat(out.trim());
    return Number.isFinite(kb) ? Math.round(kb / 1024) : 0;
}

/** 杀掉整棵进程树：Windows taskkill /T /F；POSIX 先 SIGTERM 再 SIGKILL（进程组）。 */
export async function killProcessTree(pid) {
    if (!Number.isFinite(pid) || pid <= 0) return;
    if (process.platform === 'win32') {
        try {
            await execFileP('taskkill', ['/PID', String(pid), '/T', '/F']);
        } catch { /* 进程已经没了就算了 */ }
        return;
    }
    try { process.kill(-pid, 'SIGTERM'); } catch { /* 同上 */ }
    await sleep(300);
    try { process.kill(-pid, 'SIGKILL'); } catch { /* 同上 */ }
}

/**
 * 找解释器：显式 command 优先；否则按平台依次试 py / python / python3。
 * 每个候选都真起一次 `-c "import sys"`，只认能跑通的。
 */
export async function locateInterpreter({ command = '', spawnImpl = nodeSpawn, env = process.env } = {}) {
    const wanted = asText(command);
    const candidates = wanted !== '' ? [wanted] : (process.platform === 'win32' ? ['python', 'py', 'python3'] : ['python3', 'python']);
    const failures = [];
    for (const candidate of candidates) {
        const ok = await new Promise((resolve) => {
            let settled = false;
            let child;
            try {
                child = spawnImpl(candidate, ['-c', 'import sys; print(sys.version_info[0])'], {
                    stdio: ['ignore', 'ignore', 'pipe'], windowsHide: true, env,
                });
            } catch (error) {
                failures.push(`${candidate}: ${error?.message ?? error}`);
                resolve(false);
                return;
            }
            const finish = (value) => { if (!settled) { settled = true; resolve(value); } };
            const timer = setTimeout(() => { try { child.kill(); } catch { /**/ } finish(false); }, 10_000);
            child.on('error', (error) => { clearTimeout(timer); failures.push(`${candidate}: ${error?.message ?? error}`); finish(false); });
            child.on('exit', (code) => { clearTimeout(timer); if (code === 0) finish(true); else { failures.push(`${candidate}: 退出码 ${code}`); finish(false); } });
        });
        if (ok) return { ok: true, command: candidate, error: '' };
    }
    return {
        ok: false,
        command: '',
        error: `找不到可用的 Python 解释器（试过 ${candidates.join(' / ')}）。在设置里填「解释器路径」，或把 Python 加进 PATH。${failures.length === 0 ? '' : ` 明细：${failures.join('；')}`}`,
    };
}

/**
 * 造一个管理器：单例状态 + 单飞启动 + 空闲监督 + RSS 看门狗。
 *
 * 所有外部动作（spawn / fetch / 读内存 / 杀进程 / 计时器）都可注入，测试不必
 * 真的起 Python —— 覆盖「只启动一次、并发复用、空闲被杀、启动失败、超预算」
 * 靠一个假子进程就够了。
 */
export function createLocalEmbeddingManager(options = {}) {
    const spawnImpl = options.spawnImpl ?? nodeSpawn;
    const fetchImpl = options.fetchImpl ?? ((...args) => globalThis.fetch(...args));
    const now = options.now ?? Date.now;
    const timers = options.timers ?? { setTimeout, clearTimeout, setInterval, clearInterval };
    const readRss = options.readRss ?? readProcessRssMb;
    const killTree = options.killTree ?? killProcessTree;
    const locate = options.locate ?? ((params) => locateInterpreter(params));
    const fsx = options.fs ?? { existsSync, statSync };
    const env = options.env ?? process.env;
    const commandOverride = options.command ?? '';
    const rssIntervalMs = options.rssIntervalMs ?? 1_000;
    const retryCooldownMs = options.retryCooldownMs ?? 5_000;

    let state = 'idle'; // idle | starting | ready | failed
    let child = null;
    // 服务端自己的 pid。Windows 上命令是 py.exe 时，child.pid 是启动器而不是真正吃内存的
    // python.exe：RSS 必须读就绪行里报的 pid，否则看门狗看到的是个 8MB 的空壳。
    let processPid = 0;
    let port = 0;
    let dim = 0;
    let model = LOCAL_EMBEDDING_MODEL;
    let lastError = '';
    let startedAt = 0;
    let lastUsedAt = 0;
    let rssMb = 0;
    let rssPeakMb = 0;
    let restarts = 0;
    let idleTimer = null;
    let rssTimer = null;
    let failedAt = 0;
    let starting = null;

    function baseURLOf() {
        return port > 0 ? `http://127.0.0.1:${port}/v1` : '';
    }

    function status() {
        return {
            state,
            pid: processPid > 0 ? processPid : (child !== null && Number.isFinite(child.pid) ? child.pid : 0),
            port,
            dim,
            model,
            rssMb,
            rssPeakMb,
            restarts,
            lastError,
            startedAt,
            lastUsedAt,
        };
    }

    function clearTimers() {
        if (idleTimer !== null) { timers.clearTimeout(idleTimer); idleTimer = null; }
        if (rssTimer !== null) { timers.clearInterval(rssTimer); rssTimer = null; }
    }

    async function shutdown(reason) {
        clearTimers();
        const pid = processPid > 0 ? processPid : (child !== null && Number.isFinite(child.pid) ? child.pid : 0);
        const dying = child;
        child = null;
        processPid = 0;
        port = 0;
        if (dying !== null) {
            try { dying.kill(); } catch { /* 已退出 */ }
            if (pid > 0) await killTree(pid);
        }
        if (reason === 'failed') state = 'failed';
        else if (reason === 'idle') {
            state = 'idle';
            lastError = '';
            // 进程没了就把「当前占用」清零；峰值留着，设置页还要显示它。
            rssMb = 0;
        }
        else state = 'idle';
    }

    function armIdle(cfg) {
        if (idleTimer !== null) timers.clearTimeout(idleTimer);
        idleTimer = timers.setTimeout(() => {
            idleTimer = null;
            const idleFor = now() - lastUsedAt;
            if (state === 'ready' && idleFor >= cfg.idleTimeoutMs) {
                void shutdown('idle');
            } else if (state === 'ready') {
                armIdle(cfg);
            }
        }, Math.min(cfg.idleTimeoutMs, 30_000));
    }

    function armWatchdog(cfg) {
        if (rssTimer !== null) timers.clearInterval(rssTimer);
        rssTimer = timers.setInterval(async () => {
            if (state !== 'ready') return;
            const pid = processPid > 0 ? processPid : (child !== null && Number.isFinite(child.pid) ? child.pid : 0);
            if (pid <= 0) return;
            let value = 0;
            try { value = await readRss(pid); } catch { return; }
            if (!Number.isFinite(value) || value <= 0) return;
            rssMb = Math.round(value);
            rssPeakMb = Math.max(rssPeakMb, rssMb);
            if (rssMb > cfg.maxRssMb) {
                lastError = `本地推理进程内存超过上限：实测 ${rssMb}MB > 预算 ${cfg.maxRssMb}MB，已杀进程并退回词法检索。`;
                state = 'failed';
                failedAt = now();
                await shutdown('failed');
                lastError = `本地推理进程内存超过上限：实测 ${rssMb}MB > 预算 ${cfg.maxRssMb}MB，已杀进程并退回词法检索。`;
            }
        }, rssIntervalMs);
    }

    async function health() {
        const url = baseURLOf();
        if (url === '') return false;
        try {
            const response = await fetchImpl(`${url}/health`, { signal: AbortSignal.timeout(1_500) });
            return response.ok === true;
        } catch {
            return false;
        }
    }

    function startProcess(cfg, params) {
        const modelDir = cfg.modelDir !== '' ? cfg.modelDir : defaultLocalModelDir(params.kbDir);
        const modelCheck = checkModelDir(modelDir, fsx);
        if (!modelCheck.ok) {
            state = 'failed';
            failedAt = now();
            lastError = modelCheck.error;
            return Promise.resolve({ ok: false, error: modelCheck.error, degraded: true });
        }
        return locate({ command: cfg.command !== '' ? cfg.command : commandOverride, spawnImpl, env }).then((interpreter) => {
            if (!interpreter.ok) {
                state = 'failed';
                failedAt = now();
                lastError = interpreter.error;
                return { ok: false, error: interpreter.error, degraded: true };
            }
            return new Promise((resolve) => {
                let settled = false;
                let stdout = '';
                let stderr = '';
                let readyTimer = null;
                const args = [
                    SERVER_SCRIPT,
                    '--model', modelDir,
                    '--host', '127.0.0.1',
                    '--port', '0',
                    '--idle', String(Math.ceil(cfg.idleTimeoutMs / 1000)),
                    '--threads', String(cfg.threads),
                    '--max-length', String(cfg.maxLength),
                ];
                const fail = (message) => {
                    if (settled) return;
                    settled = true;
                    if (readyTimer !== null) timers.clearTimeout(readyTimer);
                    lastError = message;
                    state = 'failed';
                    failedAt = now();
                    const pid = processPid > 0 ? processPid : (child !== null && Number.isFinite(child.pid) ? child.pid : 0);
                    if (pid > 0) void killTree(pid);
                    child = null;
                    processPid = 0;
                    port = 0;
                    resolve({ ok: false, error: message, degraded: true });
                };
                const succeed = (info) => {
                    if (settled) return;
                    settled = true;
                    if (readyTimer !== null) timers.clearTimeout(readyTimer);
                    port = Number(info.port) || 0;
                    processPid = Number.isInteger(Number(info.pid)) && Number(info.pid) > 0 ? Number(info.pid) : (child !== null && Number.isFinite(child.pid) ? child.pid : 0);
                    dim = Number(info.dim) || 0;
                    model = asText(info.model) || LOCAL_EMBEDDING_MODEL;
                    state = 'ready';
                    startedAt = now();
                    lastUsedAt = now();
                    restarts += 1;
                    lastError = '';
                    armIdle(cfg);
                    armWatchdog(cfg);
                    resolve({ ok: true, baseURL: baseURLOf(), model, dim, pid: processPid, degraded: false });
                };
                try {
                    child = spawnImpl(interpreter.command, args, {
                        stdio: ['ignore', 'pipe', 'pipe'],
                        windowsHide: true,
                        detached: process.platform !== 'win32',
                        env,
                    });
                } catch (error) {
                    fail(`启动本地推理进程失败：${error?.message ?? error}`);
                    return;
                }
                child.on('error', (error) => fail(`启动本地推理进程失败：${error?.message ?? error}`));
                child.stdout?.on('data', (chunk) => {
                    stdout += String(chunk);
                    let index = stdout.indexOf('\n');
                    while (index >= 0) {
                        const line = stdout.slice(0, index);
                        stdout = stdout.slice(index + 1);
                        if (line.startsWith(LOCAL_EMBED_READY)) {
                            try { succeed(JSON.parse(line.slice(LOCAL_EMBED_READY.length))); } catch { fail('本地推理进程的就绪行不是合法 JSON。'); }
                        }
                        index = stdout.indexOf('\n');
                    }
                });
                child.stderr?.on('data', (chunk) => { stderr += String(chunk); if (stderr.length > 8_000) stderr = stderr.slice(-8_000); });
                child.on('exit', (code) => {
                    if (settled) {
                        if (state === 'ready' || state === 'starting') {
                            // 服务自己退了（崩溃 / 空闲自杀）：下次检索重新拉起
                            state = 'failed';
                            failedAt = now();
                            lastError = lastError !== '' ? lastError : `本地推理进程意外退出（退出码 ${code}）。`;
                            clearTimers();
                            child = null;
                            processPid = 0;
                            port = 0;
                        }
                        return;
                    }
                    const detail = stderr.trim().split(/\r?\n/).slice(-3).join(' / ');
                    fail(`本地推理进程启动失败（退出码 ${code}）${detail === '' ? '' : `：${detail}`}`);
                });
                readyTimer = timers.setTimeout(() => fail(`本地推理进程启动超时（${cfg.startupTimeoutMs}ms 内没有输出就绪行）。`), cfg.startupTimeoutMs);
            });
        });
    }

    /**
     * 只等一小会儿：冷启动要 import torch + transformers 再加载权重，本机实测 45–120 秒（机器空闲时约 50 秒，和别的重活抢 CPU 时到过 118 秒），
     * 检索不能跟着卡这么久。等不到就返回 { starting:true }，进程继续在后台起，
     * 下一次检索多半就能用；调用方按「这次没有向量」退回词法，并把原因写进回执。
     */
    function bounded(promise, waitMs) {
        const budget = Number(waitMs);
        if (!Number.isFinite(budget) || budget <= 0) return promise;
        return Promise.race([
            promise,
            new Promise((resolve) => {
                const timer = timers.setTimeout(() => {
                    resolve({
                        ok: false,
                        starting: true,
                        degraded: false,
                        model: LOCAL_EMBEDDING_MODEL,
                        dim: LOCAL_EMBEDDING_DIMENSIONS,
                        error: `本地推理进程正在启动（首次要加载 torch/transformers 与权重，实测 45–120 秒）。这次先用词法，稍后重试即可。`,
                    });
                }, budget);
                promise.then(
                    () => timers.clearTimeout(timer),
                    () => timers.clearTimeout(timer),
                );
            }),
        ]);
    }

    async function ensure({ config, kbDir = '', waitMs } = {}) {
        const cfg = normalizeLocalEmbeddingConfig(config);
        if (!cfg.enabled) return { ok: false, disabled: true, error: '本地推理没有开启。', degraded: false };
        if (state === 'ready') {
            if (await health()) {
                lastUsedAt = now();
                armIdle(cfg);
                return { ok: true, baseURL: baseURLOf(), model, dim, pid: child !== null ? child.pid : 0, degraded: false, reused: true };
            }
            lastError = '本地推理服务的健康检查失败，重启进程。';
            await shutdown('failed');
        }
        if (starting !== null) return bounded(starting, waitMs);
        if (state === 'failed' && failedAt > 0 && now() - failedAt < retryCooldownMs) {
            return { ok: false, error: lastError, degraded: true };
        }
        state = 'starting';
        starting = startProcess(cfg, { kbDir }).finally(() => { starting = null; });
        return bounded(starting, waitMs);
    }

    async function stop() {
        state = 'idle';
        lastError = '';
        await shutdown('stopped');
    }

    return { ensure, stop, status, refresh: () => { lastUsedAt = now(); } };
}

/** 进程内单例：插件里所有检索共用一次启动。 */
export const localEmbeddingManager = createLocalEmbeddingManager();

/** 便捷入口：把本地服务当成远端 embedding 通道拿到 baseURL/model。 */
export async function ensureLocalEmbedding({ config, kbDir = '', waitMs } = {}) {
    return localEmbeddingManager.ensure({ config, kbDir, waitMs });
}

export function localEmbeddingStatus() {
    return localEmbeddingManager.status();
}
