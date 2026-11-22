/**
 * `src/kb-local-embed.js` 的验收测试（第一百零三轮）。
 *
 * 不真的起 Python：spawn / fetch / 计时器 / 读内存 / 杀进程全部注入。这里验的是
 * 「什么时候启动、只启动一次、空闲怎么被杀、超预算怎么退、失败怎么报」。
 *
 * 跑法：`node test/kb-local-embed.mjs`
 */
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import test from 'node:test';

import {
    LOCAL_EMBED_READY,
    LOCAL_EMBEDDING_DIMENSIONS,
    checkModelDir,
    createLocalEmbeddingManager,
    defaultLocalModelDir,
    normalizeLocalEmbeddingConfig,
} from '../src/kb-local-embed.js';

const tick = async () => { for (let i = 0; i < 5; i += 1) await new Promise((resolve) => setImmediate(resolve)); };

/** 假时钟：setTimeout / setInterval 全进一张表，advance 按到期顺序跑。 */
function makeClock() {
    let time = 10_000;
    let seq = 0;
    const tasks = new Map();
    return {
        now: () => time,
        setTimeout(fn, ms = 0) { const id = ++seq; tasks.set(id, { fn, at: time + Math.max(0, Number(ms)), every: 0 }); return id; },
        clearTimeout(id) { tasks.delete(id); },
        setInterval(fn, ms = 0) { const id = ++seq; tasks.set(id, { fn, at: time + Math.max(1, Number(ms)), every: Math.max(1, Number(ms)) }); return id; },
        clearInterval(id) { tasks.delete(id); },
        pending: () => tasks.size,
        async advance(ms) {
            const target = time + ms;
            for (;;) {
                let pick = null;
                for (const [id, task] of tasks) {
                    if (task.at <= target && (pick === null || task.at < pick[1].at)) pick = [id, task];
                }
                if (pick === null) break;
                const [id, task] = pick;
                time = task.at;
                if (task.every > 0) task.at += task.every;
                else tasks.delete(id);
                await task.fn();
                await tick();
            }
            time = target;
            await tick();
        },
    };
}

/** 假子进程：stdout/stderr 是 EventEmitter，ready/fail 由用例手动触发。 */
function makeSpawner() {
    const children = [];
    const calls = [];
    const spawnImpl = (command, args) => {
        const child = new EventEmitter();
        child.pid = 7000 + children.length;
        child.stdout = new EventEmitter();
        child.stderr = new EventEmitter();
        child.killed = false;
        child.kill = () => { child.killed = true; };
        child.ready = (info = {}) => {
            child.stdout.emit('data', Buffer.from(`${LOCAL_EMBED_READY}${JSON.stringify({ port: 51234, dim: LOCAL_EMBEDDING_DIMENSIONS, model: 'bge-small-zh-v1.5', maxLength: 256, ...info })}\n`));
        };
        child.fail = (code, text = '') => {
            if (text !== '') child.stderr.emit('data', Buffer.from(text));
            child.emit('exit', code);
        };
        children.push(child);
        calls.push({ command, args });
        return child;
    };
    return { spawnImpl, children, calls };
}

function makeManager(overrides = {}) {
    const clock = makeClock();
    const spawner = makeSpawner();
    const killed = [];
    let healthOk = true;
    const manager = createLocalEmbeddingManager({
        spawnImpl: spawner.spawnImpl,
        fetchImpl: async () => ({ ok: healthOk, status: healthOk ? 200 : 503, text: async () => '{}' }),
        now: clock.now,
        timers: clock,
        readRss: async () => 300,
        killTree: async (pid) => { killed.push(pid); },
        locate: async () => ({ ok: true, command: 'python', error: '' }),
        fs: {
            existsSync: () => true,
            statSync: () => ({ isDirectory: () => true }),
        },
        env: {},
        command: '',
        rssIntervalMs: 1_000,
        retryCooldownMs: 5_000,
        ...overrides,
    });
    return {
        manager, clock, spawner, killed,
        setHealth: (value) => { healthOk = value; },
    };
}

const CONFIG = { enabled: true, modelDir: '/models/bge-small-zh-v1.5', idleTimeoutMs: 5_000, maxRssMb: 600, startupTimeoutMs: 60_000, maxLength: 256, threads: 2 };

test('配置与模型目录：默认关闭、clamp、缺文件报清楚', () => {
    assert.equal(normalizeLocalEmbeddingConfig({}).enabled, false);
    const cfg = normalizeLocalEmbeddingConfig({ enabled: 'yes', idleTimeoutMs: 10, maxRssMb: 99, startupTimeoutMs: 1, maxLength: 9_999, threads: 99 });
    assert.equal(cfg.enabled, false, '只有显式 true 才算开启');
    assert.equal(cfg.idleTimeoutMs, 5_000);
    assert.equal(cfg.maxRssMb, 200);
    assert.equal(cfg.startupTimeoutMs, 5_000);
    assert.equal(cfg.maxLength, 512);
    assert.equal(cfg.threads, 16);
    assert.equal(defaultLocalModelDir('C:/kb/'), 'C:/kb/models/bge-small-zh-v1.5');
    assert.equal(checkModelDir('').ok, false);
    const missing = checkModelDir('/x', { existsSync: (p) => p !== '/x/config.json', statSync: () => ({ isDirectory: () => true }) });
    assert.match(missing.error, /config\.json/);
    const ok = checkModelDir('/x', { existsSync: () => true, statSync: () => ({ isDirectory: () => true }) });
    assert.equal(ok.ok, true);
});

test('未开启：直接说没开，不 spawn', async () => {
    const { manager, spawner } = makeManager();
    const result = await manager.ensure({ config: { enabled: false }, kbDir: '/kb' });
    assert.equal(result.ok, false);
    assert.equal(result.disabled, true);
    assert.equal(spawner.calls.length, 0);
});

test('并发只启动一次，成功后复用；参数含模型、空闲与线程', async () => {
    const { manager, spawner } = makeManager();
    const p1 = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    const p2 = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    await tick();
    assert.equal(spawner.calls.length, 1, '并发两次也只能起一个进程');
    spawner.children[0].ready();
    const [a, b] = await Promise.all([p1, p2]);
    assert.equal(a.ok, true);
    assert.equal(a.baseURL, 'http://127.0.0.1:51234/v1');
    assert.equal(a.dim, 512);
    assert.equal(b.ok, true);
    const args = spawner.calls[0].args;
    assert.ok(args.includes('--model') && args.includes('/models/bge-small-zh-v1.5'));
    assert.ok(args.includes('--port') && args.includes('0'), '端口要交给系统分配，避免撞车');
    assert.ok(args.includes('--idle') && args.includes('5'));
    assert.ok(args.includes('--threads') && args.includes('2'));
    assert.ok(args.includes('--max-length') && args.includes('256'));
    assert.equal(manager.status().state, 'ready');

    const again = await manager.ensure({ config: CONFIG, kbDir: '/kb' });
    assert.equal(again.ok, true);
    assert.equal(again.reused, true);
    assert.equal(spawner.calls.length, 1, '健康检查通过就不该再起进程');
});

test('启动很慢时只等 waitMs，返回 starting 而不是阻塞检索', async () => {
    const { manager, clock, spawner } = makeManager();
    const p = manager.ensure({ config: CONFIG, kbDir: '/kb', waitMs: 100 });
    await tick();
    await clock.advance(101);
    const result = await p;
    assert.equal(result.ok, false);
    assert.equal(result.starting, true, '进程还在后台起，下一次检索就能用');
    assert.match(result.error, /正在启动/);
    spawner.children[0].ready();
    await tick();
    assert.equal(manager.status().state, 'ready');
    const next = await manager.ensure({ config: CONFIG, kbDir: '/kb', waitMs: 100 });
    assert.equal(next.ok, true, '后台起好后下一次要直接用');
});

test('空闲到期杀进程树，状态回到 idle 且不留计时器', async () => {
    const { manager, clock, spawner, killed } = makeManager();
    const p = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    await tick();
    spawner.children[0].ready();
    await p;
    assert.equal(killed.length, 0);
    await clock.advance(5_000);
    assert.deepEqual(killed, [7_000], '空闲到期要杀整棵进程树');
    assert.equal(spawner.children[0].killed, true);
    assert.equal(manager.status().state, 'idle');
    assert.equal(manager.status().port, 0);
    assert.equal(clock.pending(), 0, '空闲计时器与看门狗都要清掉');
});

test('RSS 超预算：杀进程、标 failed、错误里带实测数字', async () => {
    const { manager, clock, spawner, killed } = makeManager({ readRss: async () => 712 });
    const p = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    await tick();
    spawner.children[0].ready();
    await p;
    await clock.advance(1_000);
    const status = manager.status();
    assert.equal(status.state, 'failed');
    assert.equal(status.rssMb, 712);
    assert.equal(status.rssPeakMb, 712);
    assert.match(status.lastError, /712MB/);
    assert.match(status.lastError, /600MB/);
    assert.deepEqual(killed, [7_000]);
    assert.equal(spawner.children[0].killed, true);
});

test('启动失败：带退出码与 stderr，冷却期内不反复拉起', async () => {
    const { manager, clock, spawner } = makeManager();
    const p = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    await tick();
    spawner.children[0].fail(2, 'ImportError: No module named torch\n');
    const result = await p;
    assert.equal(result.ok, false);
    assert.match(result.error, /退出码 2/);
    assert.match(result.error, /ImportError/, '失败原因要原样带出来，用户才知道缺什么');

    const again = await manager.ensure({ config: CONFIG, kbDir: '/kb' });
    assert.equal(again.ok, false);
    assert.equal(spawner.calls.length, 1, '冷却期内第二次检索不再起进程');

    await clock.advance(5_001);
    const retry = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    await tick();
    assert.equal(spawner.calls.length, 2, '冷却期过了可以重试');
    spawner.children[1].ready();
    assert.equal((await retry).ok, true);
});

test('模型目录缺失：不起进程，错误指向下载或目录', async () => {
    const { manager, spawner } = makeManager({
        fs: {
            existsSync: (p) => !String(p).includes('config.json'),
            statSync: () => ({ isDirectory: () => true }),
        },
    });
    const result = await manager.ensure({ config: CONFIG, kbDir: '/kb' });
    assert.equal(result.ok, false);
    assert.match(result.error, /缺 config\.json/);
    assert.equal(spawner.calls.length, 0);
});

test('健康检查失败：重启一个新进程，不继续用死端口', async () => {
    const ctx = makeManager();
    const { manager, spawner } = ctx;
    const p1 = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    await tick();
    spawner.children[0].ready();
    await p1;
    ctx.setHealth(false);
    const p2 = manager.ensure({ config: CONFIG, kbDir: '/kb' });
    await tick();
    assert.equal(spawner.calls.length, 2);
    assert.equal(spawner.children[0].killed, true);
    spawner.children[1].ready({ port: 51235 });
    const result = await p2;
    assert.equal(result.ok, true);
    assert.equal(result.baseURL, 'http://127.0.0.1:51235/v1');
});

test('启动超时：按 startupTimeoutMs 报超时并杀进程', async () => {
    const { manager, clock, spawner, killed } = makeManager();
    const config = { ...CONFIG, startupTimeoutMs: 5_000 };
    const p = manager.ensure({ config, kbDir: '/kb' });
    await tick();
    await clock.advance(5_001);
    const result = await p;
    assert.equal(result.ok, false);
    assert.match(result.error, /启动超时（5000ms/);
    assert.deepEqual(killed, [7_000]);
});
