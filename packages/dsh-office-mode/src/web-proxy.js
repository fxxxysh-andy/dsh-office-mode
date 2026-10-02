/**
 * 内置检索的出口代理：把设置页里填的代理接到插件的 HTTP 上。
 *
 * ## 两条路，优先走私有的那条
 *
 * **1) 私有分派器（首选）** —— 宿主进程里装着 `undici`（DSH 自己的依赖）。
 * 借宿主解析位置把它取出来（与 index.js 取 schemastery / cosmokit 同一条锚点链），
 * 建一个 `ProxyAgent`，只在本插件发请求时通过 `dispatcher` 选项传进去。
 * 好处是**只有本插件的联网走代理**：宿主的 web 服务、别的插件都不受影响；
 * 改设置或清空设置当场生效（旧 agent 关掉就行）。
 *
 * **2) 进程级环境变量（兜底）** —— 拿不到 undici 时改走 Node 24+ 的
 * `http.setGlobalProxyFromEnv()`：把代理写进 HTTP_PROXY / HTTPS_PROXY，再调它装成
 * 全局 dispatcher。这条路有两个代价，必须照实说：
 *   - 它是**进程级**的：宿主自己的 web 服务、别的插件的 fetch 也一起走代理；
 *   - 它是**只能装、不能清**的（本机 Node 26.7.0 实测：删掉环境变量再调一次，
 *     连接仍旧走原来的代理）。所以清空设置要重启 DSH 才回直连。
 *
 * 两条都不可用（没有 undici，运行时又没有那个 API）时返回 ok:false 与原因，
 * 调用方照实报给用户 —— 不假装设置生效了。
 *
 * ## 本地地址不进代理
 *
 * ## 本地地址不经代理（两条路都成立，但做的地方不同）
 *
 * searxng 的默认端点就是 http://127.0.0.1:8080，把它转发给代理必然失败。兜底那条靠
 * `NO_PROXY`（只增不改）；私有分派器不吃 `NO_PROXY`（undici 的 proxy-agent 没有这个
 * 处理），所以由**请求侧**判 —— `web.js` 的 `defaultNetwork.fetch` 对回环地址不挂
 * `dispatcher`（见 `shouldBypassProxy`）。两条路的承诺因此一致（审查 P2-12b）。
 *
 * ## 兜底那条的两个副作用（必须让用户知道）
 *
 *   1. 它会**顶掉**宿主启动时按环境变量装好的那份全局 dispatcher（宿主
 *      `@deepseek-ai/dsh-http-proxy` 用 npm undici 的 `setGlobalDispatcher`，而
 *      `http.setGlobalProxyFromEnv()` 会把同一个全局槽换成它自己的 agent）；
 *   2. 代理地址（含 `user:pass@`）会写进 `process.env`，宿主拉起的子进程
 *      （bash / pwsh / python）会继承它。
 *
 * 这两条只在「拿不到 undici」的部署里才会发生；正常 DSH 走的是私有分派器。
 *
 * @module dsh-office-mode/web-proxy
 */

import http from 'node:http';
import { isIP } from 'node:net';

/** 兜底那条会写的环境变量名（大小写各一套：Node 两套都读）。 */
export const PROXY_ENV_KEYS = Object.freeze(['HTTP_PROXY', 'HTTPS_PROXY', 'NO_PROXY', 'http_proxy', 'https_proxy', 'no_proxy']);

/** 回环地址永远不经代理（本地服务、自建 SearXNG 都要能直连）。 */
const LOOPBACK_NO_PROXY = ['localhost', '127.0.0.1', '::1'];

/** 当前生效的代理：值、方式（dispatcher / global / none）、私有分派器、说明。 */
let active = { proxy: '', mode: 'none', dispatcher: undefined, note: '' };

/**
 * 把设置里填的代理串收敛成一个 URL；留空表示「不用代理」。
 *
 * 只接受 http / https：Node 在 `NODE_USE_ENV_PROXY` 那条路上遇到别的 scheme 会直接
 * 让进程起不来，宿主包因此也把非 http(s) 的值挡在子进程之外。SOCKS 这类值在这里
 * 明确拒掉，而不是装一个用不了的策略。
 *
 * **失败信息里不回显原文**：代理地址可能带 `user:password@`，而这些 reason 会进
 * `ctx.logger.warn` 与 `office_help` 的运行期实况（审查 P1-1：原先非法地址会把明文
 * 口令回显出去）。所以错误信息只给「形状」与打码后的样子。
 */
export function normalizeProxyUrl(input) {
    const text = typeof input === 'string' ? input.trim() : '';
    if (text === '') return { ok: true, url: '' };
    if (text.length > 2048) return { ok: false, reason: '代理地址超过 2048 字符上限。' };
    const shape = maskProxyUrl(text);
    let url;
    try {
        url = new URL(text);
    } catch {
        return { ok: false, reason: `不是合法的代理地址：${shape}（形如 http://127.0.0.1:7897）` };
    }
    if (url.protocol !== 'http:' && url.protocol !== 'https:') {
        return { ok: false, reason: `只支持 http / https 代理，收到 ${url.protocol}（SOCKS 请交给系统代理，别填这里）` };
    }
    if (url.hostname === '') return { ok: false, reason: '代理地址里没有主机名。' };
    // 用 href 而不是 origin：**origin 会把 user:pass@ 丢掉**，带鉴权的代理会因此静默失联。
    let href = url.href;
    if (url.pathname === '/' && url.search === '' && url.hash === '') href = href.replace(/\/$/, '');
    return { ok: true, url: href };
}

/**
 * 反馈里显示的代理：账号密码打码（它可能带凭据，不该原样进对话与日志）。
 *
 * 解析不了的字符串也走一次正则兜底：`http://user:pw@host:99999` 这种 `new URL()`
 * 会抛的输入，恰恰是最可能带凭据的手误形态。
 */
export function maskProxyUrl(input) {
    const text = typeof input === 'string' ? input.trim() : '';
    if (text === '') return '';
    try {
        const url = new URL(text);
        if (url.username === '' && url.password === '') return text;
        return `${url.protocol}//***:***@${url.host}`;
    } catch {
        return maskInlineCredentials(text);
    }
}

/** 兜底打码：把 `//user:pass@` 整段换成 `//***:***@`（只认 URL 里那一段）。 */
function maskInlineCredentials(text) {
    return String(text).replace(/\/\/[^/@\s]*@/g, '//***:***@');
}

/** 兜底那条路要靠的官方口子在不在这台运行时里。 */
export function proxySupport(runtime = http) {
    if (typeof runtime?.setGlobalProxyFromEnv === 'function') {
        return { supported: true, reason: `Node ${process.version} 支持 http.setGlobalProxyFromEnv()` };
    }
    return {
        supported: false,
        reason: `本运行时（Node ${process.version}）没有 http.setGlobalProxyFromEnv()`,
    };
}

/** 把回环地址并进 NO_PROXY（只增不改）。 */
function mergeNoProxy(existing) {
    const parts = String(existing ?? '').split(',').map((part) => part.trim()).filter((part) => part !== '');
    for (const entry of LOOPBACK_NO_PROXY) {
        if (!parts.includes(entry)) parts.push(entry);
    }
    return parts.join(',');
}

/** 借宿主解析位置取 undici（取不到返回 undefined，不抛）。 */
function loadUndici(resolveModule) {
    if (typeof resolveModule !== 'function') return undefined;
    try {
        const loaded = resolveModule('undici');
        return loaded !== null && typeof loaded === 'object' && typeof loaded.ProxyAgent === 'function' ? loaded : undefined;
    } catch {
        return undefined;
    }
}

/** 关掉我们自己建的私有分派器（只关这一种，别人建的不管）。 */
function closeDispatcher(dispatcher) {
    if (dispatcher === undefined) return;
    for (const method of ['close', 'destroy']) {
        if (typeof dispatcher[method] === 'function') {
            try {
                // undici 的 close() 返回 Promise：不接住的话，拒绝会变成
                // unhandledRejection（审查 P2-9）。
                Promise.resolve(dispatcher[method]()).catch(() => {});
            } catch {
                // 关不掉不影响换代理：新的分派器照样建得起来。
            }
            return;
        }
    }
}

/** 兜底那条会写的 6 个键：装失败时要能原样rollback。 */
function snapshotProxyEnv(env) {
    const snapshot = {};
    for (const key of PROXY_ENV_KEYS) {
        snapshot[key] = Object.prototype.hasOwnProperty.call(env, key) ? env[key] : undefined;
    }
    return snapshot;
}

/** 把 6 个键恢复成快照（undefined = 当时不存在，删掉）。 */
function restoreProxyEnv(env, snapshot) {
    for (const key of PROXY_ENV_KEYS) {
        if (snapshot[key] === undefined) delete env[key];
        else env[key] = snapshot[key];
    }
}

/**
 * 按设置装一次代理。值没变就原地不动（设置页每改一个字段都会走到这里）。
 *
 * @param {string} input 设置里的 `search.proxy`
 * @param {{env?: object, runtime?: object, resolveModule?: (specifier: string) => unknown}} [deps] 测试注入点
 * @returns {{ok: boolean, applied: boolean, mode: 'dispatcher'|'global'|'none', proxy: string, reason: string}}
 */
export function installProxy(input, deps = {}) {
    const env = deps.env ?? process.env;
    const runtime = deps.runtime ?? http;
    const previous = active;
    const normalized = normalizeProxyUrl(input);
    if (normalized.ok !== true) {
        // 地址不合法**不动**已经生效的代理，但要把两件事都说清：这次为什么没装、现在实际走什么。
        const still = previous.proxy === ''
            ? '当前没有生效的代理（联网按本进程原有出口直连）。'
            : `当前仍走 ${maskProxyUrl(previous.proxy)}（这次没有改动它）。`;
        const reason = normalized.reason + ' ' + still;
        active = { ...previous, note: reason };
        return { ok: false, applied: false, mode: previous.mode, proxy: previous.proxy, reason };
    }
    const proxy = normalized.url;

    if (proxy === active.proxy) {
        return { ok: true, applied: false, mode: active.mode, proxy, reason: active.note };
    }

    // 换值/清空：先把自己建的分派器关掉，再从头决定。
    if (active.mode === 'dispatcher') closeDispatcher(active.dispatcher);
    // 进程级那条装上去就**清不掉**（Node 只提供装），所以换值时要把这件事带进说明里。
    const pinnedGlobal = previous.mode === 'global' && previous.proxy !== '' ? previous.proxy : '';
    active = { proxy: '', mode: 'none', dispatcher: undefined, note: '' };

    if (proxy === '') {
        const reason = pinnedGlobal !== ''
            ? `设置已留空：本插件自己发起的请求回到直连，但**进程级那条仍在生效**`
                + `（宿主自己的联网与它拉起的子进程仍走 ${maskProxyUrl(pinnedGlobal)}）：`
                + '它是只能装不能清的，要彻底回直连得重启 DSH。'
            : '未配置代理：联网按本进程原有出口直连。';
        active.note = reason;
        return { ok: true, applied: false, mode: 'none', proxy: '', reason };
    }

    const undici = loadUndici(deps.resolveModule);
    if (undici !== undefined) {
        try {
            const dispatcher = new undici.ProxyAgent({ uri: proxy });
            const reason = `已让本插件自己发起的联网走 ${maskProxyUrl(proxy)}（私有分派器，只影响本插件；`
                + '宿主 web 服务与别的插件不受影响，改设置当场生效）。'
                + (pinnedGlobal === ''
                    ? ''
                    : `注意：之前装过进程级代理 ${maskProxyUrl(pinnedGlobal)}，它清不掉，`
                        + '宿主自己的联网与子进程仍走那一条，要回直连得重启 DSH。');
            active = { proxy, mode: 'dispatcher', dispatcher, note: reason };
            return { ok: true, applied: true, mode: 'dispatcher', proxy, reason };
        } catch (error) {
            // 建不出来就继续往兜底那条路走，把原因带上。
            active.note = `建立代理分派器失败：${error?.message ?? error}`;
        }
    }

    const support = proxySupport(runtime);
    if (!support.supported) {
        const reason = `${active.note === '' ? '' : active.note + '；'}${support.reason}，也没有可用的 undici 分派器`
            + '：代理设置不生效。请用环境变量（HTTPS_PROXY）启动 DSH，或换 Node 24+ 运行。';
        active = { proxy: '', mode: 'none', dispatcher: undefined, note: reason };
        return { ok: false, applied: false, mode: 'none', proxy, reason };
    }
    // 兜底：先快照再写，装不上就把 6 个键原样恢复（否则「返回 ok:false」与
    // 「环境变量已经改了」自相矛盾，子进程还会继承一个父进程没用上的代理 —— 审查 P1-2）。
    const snapshot = snapshotProxyEnv(env);
    try {
        env.HTTP_PROXY = proxy;
        env.HTTPS_PROXY = proxy;
        env.http_proxy = proxy;
        env.https_proxy = proxy;
        env.NO_PROXY = mergeNoProxy(env.NO_PROXY ?? env.no_proxy);
        env.no_proxy = env.NO_PROXY;
        runtime.setGlobalProxyFromEnv();
        const reason = `已把**整个进程**的 fetch 出口切到 ${maskProxyUrl(proxy)}（拿不到 undici，只能走进程级这条路；`
            + '本地地址在插件这一侧按请求绕过）。它是只能装不能清的：清空设置后要重启 DSH 才回直连。'
            + '两点副作用要记住：① 它会**顶掉**宿主启动时按环境变量装好的那份全局 dispatcher；'
            + '② 代理地址（含凭据）会写进 process.env，宿主拉起的子进程（bash / pwsh / python）会继承它。';
        active = { proxy, mode: 'global', dispatcher: undefined, note: reason };
        return { ok: true, applied: true, mode: 'global', proxy, reason };
    } catch (error) {
        restoreProxyEnv(env, snapshot);
        const reason = `装代理失败：${error?.message ?? error}（已把 HTTP_PROXY / HTTPS_PROXY / NO_PROXY 恢复原样）`;
        active = { proxy: '', mode: 'none', dispatcher: undefined, note: reason };
        return { ok: false, applied: false, mode: 'none', proxy, reason };
    }
}

/** 本插件发请求时要用的私有分派器；没有（未配置或走了进程级那条）时 undefined。 */
export function activeDispatcher() {
    return active.mode === 'dispatcher' ? active.dispatcher : undefined;
}

/**
 * 这个地址要不要绕过代理。
 *
 * 兜底那条路靠 `NO_PROXY` 让回环地址直连；私有分派器不吃 `NO_PROXY`，所以这一步
 * 必须由调用方按请求判 —— 两条路对「本地服务（自建 SearXNG 默认就在
 * http://127.0.0.1:8080）不经代理」的承诺才一致。地址解析不出来时**不绕**（宁可
 * 让代理去失败，也不静默把请求发到一个本该被代理的地址）。
 */
export function shouldBypassProxy(input) {
    let raw = '';
    try {
        raw = new URL(String(input)).hostname.toLowerCase();
    } catch {
        return false;
    }
    const host = raw.replace(/^\[|\]$/g, '');
    if (host === 'localhost') return true;
    const family = isIP(host);
    if (family === 4) return host.startsWith('127.') || host === '0.0.0.0';
    if (family === 6) {
        // ::1 与 :: 是环回/未指定；::ffff:127.x 与规范化后的 ::ffff:7f00:x 都是
        // IPv4-mapped 环回（`new URL()` 会把后者写成十六进制形式）。
        return host === '::1' || host === '::' || host.startsWith('::ffff:7f') || host.startsWith('::ffff:127.');
    }
    return false;
}

/** 当前代理状态（office_help 与设置页照实说用）。 */
export function proxyStatus() {
    return {
        proxy: active.proxy,
        proxyMasked: maskProxyUrl(active.proxy),
        mode: active.mode,
        note: active.note,
        globalSupported: proxySupport().supported,
    };
}

/** 测试用：把模块级状态清回初始值（顺带关掉测试里建的分派器）。 */
export function resetProxyStateForTest() {
    if (active.mode === 'dispatcher') closeDispatcher(active.dispatcher);
    active = { proxy: '', mode: 'none', dispatcher: undefined, note: '' };
}

/**
 * 插件卸载时放掉私有的连接池（审查 P2-9）。
 *
 * 只关我们自己建的那个 `ProxyAgent`；进程级那条（全局 dispatcher）**故意不碰** ——
 * 它不是本插件的资源，而且宿主可能正靠它工作（它也清不掉）。
 */
export function disposeProxy() {
    if (active.mode === 'dispatcher') closeDispatcher(active.dispatcher);
    active = { proxy: '', mode: 'none', dispatcher: undefined, note: '插件已卸载：私有代理分派器已关闭。' };
}
