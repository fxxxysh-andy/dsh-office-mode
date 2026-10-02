/**
 * 出口代理（`search.proxy`，第二十九轮）的单元测试。
 *
 * 这里测的是**决定**与**接线**，不测真实网络：
 *   - 地址收敛与打码（留空 / http / https / SOCKS / 垃圾串 / 超长）；
 *   - 装了哪条路：拿得到 undici 就建插件私有的分派器，拿不到才退回进程级
 *     `http.setGlobalProxyFromEnv()`（两条都用替身，不碰真实 process.env）；
 *   - 换值与清空：私有分派器会被关掉，进程级那条清不掉要照实说；
 *   - 接线：装了私有分派器之后，插件自己发请求真的带上了 `dispatcher`
 *     （用假的 globalThis.fetch 抓 init，不真发请求）。
 */
import assert from 'node:assert/strict';

import { createWebAccess } from '../src/web.js';
import {
    activeDispatcher,
    disposeProxy,
    installProxy,
    maskProxyUrl,
    normalizeProxyUrl,
    proxyStatus,
    proxySupport,
    resetProxyStateForTest,
    shouldBypassProxy,
} from '../src/web-proxy.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    } finally {
        resetProxyStateForTest();
    }
}

/** 一个假的 undici：ProxyAgent 记下构造参数，close() 留下痕迹。 */
function fakeUndici() {
    const built = [];
    class ProxyAgent {
        constructor(options) {
            this.options = options;
            this.closed = false;
            built.push(this);
        }

        close() {
            this.closed = true;
        }
    }
    return { module: { ProxyAgent }, built };
}

/** 一个假的运行时：只有「支持 / 不支持 setGlobalProxyFromEnv」两种。 */
function fakeRuntime(supported = true) {
    const calls = { install: 0 };
    return {
        calls,
        runtime: supported
            ? { setGlobalProxyFromEnv() { calls.install += 1; } }
            : {},
    };
}

await check('地址收敛：留空=直连；只认 http/https，SOCKS 与垃圾串明确拒掉', () => {
    assert.deepEqual(normalizeProxyUrl(''), { ok: true, url: '' });
    assert.deepEqual(normalizeProxyUrl('   '), { ok: true, url: '' });
    assert.equal(normalizeProxyUrl(' http://127.0.0.1:7897 ').url, 'http://127.0.0.1:7897');
    assert.equal(normalizeProxyUrl('https://proxy.example.com:8080/').url, 'https://proxy.example.com:8080');
    // 带鉴权的代理必须原样保留 user:pass@ —— URL.origin 会把它丢掉，那是静默失联。
    assert.equal(normalizeProxyUrl('http://user:pw@proxy.example.com:8080').url, 'http://user:pw@proxy.example.com:8080');
    for (const bad of ['socks5://127.0.0.1:1080', 'ftp://x/y', '127.0.0.1:7897', '不是地址']) {
        const result = normalizeProxyUrl(bad);
        assert.equal(result.ok, false, `${bad} 该被拒`);
        assert.ok(result.reason.length > 0);
    }
    assert.equal(normalizeProxyUrl('http://127.0.0.1:7897/' + 'a'.repeat(2100)).ok, false, '超长地址该被拒');
});

await check('打码：带账号密码的代理地址不把凭据原样回给对话与日志', () => {
    assert.equal(maskProxyUrl('http://127.0.0.1:7897'), 'http://127.0.0.1:7897');
    assert.equal(maskProxyUrl('http://user:secret@proxy.example.com:8080'), 'http://***:***@proxy.example.com:8080');
    assert.equal(maskProxyUrl(''), '');
});

await check('首选私有分派器：只影响本插件，改值当场重装并关掉旧的', () => {
    const undici = fakeUndici();
    const { runtime } = fakeRuntime();
    const env = {};
    const first = installProxy('http://127.0.0.1:7897', { env, runtime, resolveModule: () => undici.module });
    assert.equal(first.ok, true);
    assert.equal(first.mode, 'dispatcher');
    assert.equal(activeDispatcher(), undici.built[0], '响应里要用同一个分派器');
    assert.deepEqual(undici.built[0].options, { uri: 'http://127.0.0.1:7897' });
    assert.deepEqual(env, {}, '走私有分派器时不该动进程环境变量');

    // 同一个值重复装：不重建（设置页每改一个字段都会走到这里）。
    const again = installProxy('http://127.0.0.1:7897', { env, runtime, resolveModule: () => undici.module });
    assert.equal(again.applied, false);
    assert.equal(undici.built.length, 1);

    // 换值：旧的关掉、新的上岗。
    const second = installProxy('http://192.0.2.10:3128', { env, runtime, resolveModule: () => undici.module });
    assert.equal(second.applied, true);
    assert.equal(undici.built[0].closed, true, '旧分派器要关掉，否则连接池留着不走');
    assert.equal(activeDispatcher(), undici.built[1]);
});

await check('清空设置：私有分派器当场下线，回到直连', () => {
    const undici = fakeUndici();
    const { runtime } = fakeRuntime();
    installProxy('http://127.0.0.1:7897', { env: {}, runtime, resolveModule: () => undici.module });
    const cleared = installProxy('', { env: {}, runtime, resolveModule: () => undici.module });
    assert.equal(cleared.ok, true);
    assert.equal(cleared.mode, 'none');
    assert.equal(activeDispatcher(), undefined);
    assert.equal(undici.built[0].closed, true);
    assert.match(cleared.reason, /直连/);
});

await check('兜底进程级：没有 undici 时写环境变量并调 setGlobalProxyFromEnv，清楚交代只能装不能清', () => {
    const { runtime, calls } = fakeRuntime(true);
    const env = {};
    const done = installProxy('http://127.0.0.1:7897', {
        env,
        runtime,
        resolveModule: () => { throw new Error('没有 undici'); },
    });
    assert.equal(done.mode, 'global');
    assert.equal(calls.install, 1, '要真的调那个 API');
    assert.equal(env.HTTPS_PROXY, 'http://127.0.0.1:7897');
    assert.equal(env.HTTP_PROXY, 'http://127.0.0.1:7897');
    assert.match(env.NO_PROXY, /127\.0\.0\.1/, '回环地址要进 NO_PROXY（自建 SearXNG 走 127.0.0.1）');
    assert.match(done.reason, /只能装不能清/);

    // 进程级那条清不掉：清空时要说清「本进程仍走代理，要重启」。
    const cleared = installProxy('', { env, runtime, resolveModule: () => undefined });
    assert.equal(cleared.mode, 'none');
    assert.match(cleared.reason, /重启/);
});

await check('两条路都不通时返回失败与原因，不假装生效', () => {
    const { runtime } = fakeRuntime(false);
    const result = installProxy('http://127.0.0.1:7897', { env: {}, runtime, resolveModule: () => undefined });
    assert.equal(result.ok, false);
    assert.equal(result.mode, 'none');
    assert.match(result.reason, /setGlobalProxyFromEnv|undici/);
    assert.equal(activeDispatcher(), undefined);
    // 非法地址也走失败分支（且不碰环境变量）。
    const env = {};
    const bad = installProxy('socks5://127.0.0.1:1080', { env, runtime, resolveModule: () => undefined });
    assert.equal(bad.ok, false);
    assert.deepEqual(env, {});
});

await check('proxySupport 按运行时如实回答（Node 24+ 才有那个 API）', () => {
    assert.equal(proxySupport({ setGlobalProxyFromEnv() {} }).supported, true);
    assert.equal(proxySupport({}).supported, false);
    assert.match(proxySupport({}).reason, /setGlobalProxyFromEnv/);
});

await check('本地地址不过代理：私有分派器不吃 NO_PROXY，这一步由请求侧判', () => {
    for (const local of ['http://127.0.0.1:8080/search', 'http://localhost:8080/', 'http://[::1]:8080/', 'http://[::ffff:127.0.0.1]/', 'http://0.0.0.0:1/']) {
        assert.equal(shouldBypassProxy(local), true, `${local} 该绕开代理`);
    }
    for (const remote of ['https://example.com/x', 'http://93.184.216.34/x', 'http://::ffff:127.0.0.1/', '不是地址']) {
        assert.equal(shouldBypassProxy(remote), false, `${remote} 不该绕开代理`);
    }
});

await check('接线：装了私有分派器后，插件自己发请求带上 dispatcher（本地地址不带）', async () => {
    const undici = fakeUndici();
    installProxy('http://127.0.0.1:7897', { env: {}, runtime: fakeRuntime().runtime, resolveModule: () => undici.module });
    const real = globalThis.fetch;
    const captured = [];
    globalThis.fetch = async (url, init) => {
        captured.push({ url, init });
        const bytes = new TextEncoder().encode('代理接线用例：这是一段够长的正文，用来过内容级成败哨兵。'.repeat(3));
        return {
            status: 200,
            headers: { get: (name) => (name === 'content-type' ? 'text/plain; charset=utf-8' : null) },
            body: {
                getReader() {
                    let sent = false;
                    return {
                        async read() {
                            if (sent) return { done: true, value: undefined };
                            sent = true;
                            return { done: false, value: bytes };
                        },
                        async cancel() {},
                    };
                },
            },
        };
    };
    try {
        const access = createWebAccess({ get: () => undefined }, { apiKey: 'k' }, {
            // 故意不给 hooks.fetch：要走的正是插件自己的 defaultNetwork.fetch。
            // lookup 对回环主机名报一个公网地址，好让 SSRF 防线放行、走到 fetch 这一步。
            lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        });
        const page = await access.fetch('https://example.com/x');
        assert.equal(page.kind, 'text');
        assert.equal(captured.length, 1);
        assert.equal(captured[0].init.dispatcher, undici.built[0], '公网请求要带上装好的那个分派器');

        // 本地服务（自建 SearXNG 默认端点在 127.0.0.1:8080）走的是通道适配器那条路，
        // 不经取正文的 SSRF 防线 —— 这里断言它**不带** dispatcher，也就是不被转发给代理。
        // 适配器会因为响应形状不对而失败，那与本次断言无关。
        await access.search('x', { provider: 'searxng' }).catch(() => undefined);
        const local = captured.filter((item) => String(item.url).startsWith('http://127.0.0.1:8080'));
        assert.equal(local.length, 1, `本地通道要真的发出去一次，实际 ${captured.map((i) => i.url).join(' , ')}`);
        assert.equal(local[0].init.dispatcher, undefined, '回环地址不带 dispatcher');
    } finally {
        globalThis.fetch = real;
    }
    // 清空之后不该再带 dispatcher。
    installProxy('', { env: {}, runtime: fakeRuntime().runtime, resolveModule: () => undici.module });
    const after = proxyStatus();
    assert.equal(after.mode, 'none');
    assert.equal(activeDispatcher(), undefined);
});

await check('office_help({topic:"search"}) 末尾给出运行期实况（装成没装说得清）', async () => {
    const { buildTools } = await import('../src/tools.js');
    const { resolveConfig } = await import('../src/config.js');
    // memory.enabled:false 让 office_help 跳过记忆投影（这里只关心代理那一行）。
    const tools = buildTools(resolveConfig({ memory: { enabled: false } }));
    const help = tools.find((item) => item.name === 'office_help');

    const before = await help.execute({ topic: 'search' }, { agent: { session: { header: { cwd: process.cwd() } } } });
    assert.ok(before.text.includes('【运行期实况】'), '检索这一页要带运行期状态');
    assert.match(before.text, /出口代理：未配置|出口代理：http/, '没配代理时要照实说未配置');

    const undici = fakeUndici();
    installProxy('http://user:pw@proxy.example.com:8080', { env: {}, runtime: fakeRuntime().runtime, resolveModule: () => undici.module });
    const after = await help.execute({ topic: 'search' }, { agent: { session: { header: { cwd: process.cwd() } } } });
    assert.match(after.text, /出口代理：http:\/\/\*\*\*:\*\*\*@proxy\.example\.com:8080/, '要报出代理（凭据打码）');
    assert.match(after.text, /插件私有分派器/, '要说清走的是哪条路');

    // 别的話題不贴这一行（省字节，且与代理无关）。
    const word = await help.execute({ topic: 'word' }, { agent: { session: { header: { cwd: process.cwd() } } } });
    assert.ok(!word.text.includes('【运行期实况】'), '只有检索这一页带运行期状态');
});

await check('非法地址不回显凭据，且不动已经生效的代理', async () => {
    // 这些 reason 会进 ctx.logger.warn 与 office_help 的运行期实况 —— 口令不能出现在里面（审查 P1-1）。
    const bad = installProxy('http://user:sup3rsecret@127.0.0.1:99999', { env: {}, runtime: fakeRuntime().runtime, resolveModule: () => undefined });
    assert.equal(bad.ok, false);
    assert.ok(!bad.reason.includes('sup3rsecret'), '口令不能出现在 reason 里');
    assert.match(bad.reason, /不是合法的代理地址：http:\/\/\*\*\*:\*\*\*@127\.0\.0\.1:99999/);
    assert.match(bad.reason, /当前没有生效的代理/);

    // 生效中的代理遇到一次非法输入：状态不变，reason 要说清「仍走原来那条」。
    const undici = fakeUndici();
    installProxy('http://127.0.0.1:7897', { env: {}, runtime: fakeRuntime().runtime, resolveModule: () => undici.module });
    const stillBad = installProxy('socks5://user:pw@127.0.0.1:1080', { env: {}, runtime: fakeRuntime().runtime, resolveModule: () => undici.module });
    assert.equal(stillBad.mode, 'dispatcher');
    assert.equal(activeDispatcher(), undici.built[0], '非法输入不该把生效中的代理拆掉');
    assert.match(stillBad.reason, /当前仍走 http:\/\/127\.0\.0\.1:7897/);

    // 解析不了的串，打码函数自己也要兜住。
    assert.equal(maskProxyUrl('http://user:pw@host:99999'), 'http://***:***@host:99999');

    // 端到端：经 office_help 也看不到口令。
    const { buildTools } = await import('../src/tools.js');
    const { resolveConfig } = await import('../src/config.js');
    const help = buildTools(resolveConfig({ memory: { enabled: false } })).find((item) => item.name === 'office_help');
    const out = await help.execute({ topic: 'search' }, { agent: { session: { header: { cwd: process.cwd() } } } });
    assert.ok(!out.text.includes('sup3rsecret') && !out.text.includes('socks5://user:pw'), 'office_help 的运行期实况不能漏凭据');
});

await check('兜底失败时把环境变量恢复原样，不留半装状态', () => {
    const env = { NO_PROXY: 'example.com' };
    const runtime = { setGlobalProxyFromEnv() { throw new Error('boom'); } };
    const result = installProxy('http://127.0.0.1:7897', { env, runtime, resolveModule: () => undefined });
    assert.equal(result.ok, false);
    assert.match(result.reason, /恢复原样/);
    assert.deepEqual(env, { NO_PROXY: 'example.com' }, '写坏的 6 个键都要回到快照');
});

await check('从进程级换成私有分派器时，说清旧代理清不掉、仍在生效', () => {
    const env = {};
    const { runtime } = fakeRuntime(true);
    installProxy('http://127.0.0.1:7897', { env, runtime, resolveModule: () => undefined });
    const undici = fakeUndici();
    const second = installProxy('http://192.0.2.10:3128', { env, runtime, resolveModule: () => undici.module });
    assert.equal(second.mode, 'dispatcher');
    assert.match(second.reason, /清不掉|重启 DSH/);
    assert.match(second.reason, /127\.0\.0\.1:7897/, '要点名旧代理是哪一个');
    // 这类说明也要覆盖「清空设置」那条分支。
    const cleared = installProxy('', { env: {}, runtime, resolveModule: () => undefined });
    assert.match(cleared.reason, /进程级那条仍在生效|直连/);
});

await check('dispose 关掉私有分派器（进程级那条不是本插件的资源，不碰）', () => {
    const undici = fakeUndici();
    installProxy('http://127.0.0.1:7897', { env: {}, runtime: fakeRuntime().runtime, resolveModule: () => undici.module });
    disposeProxy();
    assert.equal(undici.built[0].closed, true);
    assert.equal(activeDispatcher(), undefined);
    assert.match(proxyStatus().note, /卸载/);
    assert.equal(proxyStatus().mode, 'none');
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`web-proxy: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
