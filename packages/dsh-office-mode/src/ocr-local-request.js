/**
 * OCR 官方供应商 Key 复用的「本机浏览器直连」守卫（第九十八轮）。
 *
 * 背景：POST /office-memory/ocr 是 dsh-host-webserver 的 raw handler，不经过
 * Gateway 的 Peer 认证；trusted cwd 只限制读哪个文件，不是用户认证。因此复用
 * 宿主 Models 里的官方 Key 前，必须 fail-closed 地确认这次请求来自本机浏览器的
 * 直连，而不是 LAN 上的另一个客户端、反向代理或端口转发。
 *
 * 注意：Host / Origin / Sec-Fetch-Site 都是**浏览器会带的元数据**，不是认证。
 * 本模块只能阻止「普通跨站 / 局域网 / 代理」误用，挡不住本机上故意伪造请求头的
 * 程序；也不承诺识别「先剥掉转发头、再从 127.0.0.1 转进来」的本机代理 —— 那种
 * 请求与真实本机浏览器看起来没有区别。这一点必须如实写进 UI / 文档，不能宣称
 * 它认证了本机进程。
 *
 * @module dsh-office-mode/ocr-local-request
 */

/** Node 在 Windows / Linux 上可能给出的本机回环地址（精确匹配，不做 DNS）。 */
const LOOPBACK_ADDRESSES = new Set([
    '127.0.0.1',
    '::1',
    '::ffff:127.0.0.1',
]);

/** 出现任意一个就说明请求经过（或声称经过）代理 / 转发，一律不借 Key。 */
const FORWARDING_HEADERS = new Set([
    'forwarded',
    'via',
    'x-real-ip',
    'x-client-ip',
    'cf-connecting-ip',
    'true-client-ip',
]);

/** 所有 x-forwarded-* 都按转发标记处理，不只列出的几个名字。 */
const FORWARDING_PREFIXES = ['x-forwarded-'];

/** Host / Origin 只接受这三种字面量；不接受 127.1、整数、DNS 名或带点的变体。 */
const AUTHORITY_PATTERN = /^(127\.0\.0\.1|localhost|\[::1\])(?::([0-9]{1,5}))?$/i;

/** Host / Origin 允许的本机主机名。`localhost` 仅作为字面量接受，不做解析。 */
const LOCAL_HOSTNAMES = new Set(['127.0.0.1', 'localhost', '[::1]']);

/**
 * 只接受回环 socket 地址。
 * @param {unknown} value `req.socket.remoteAddress`
 * @returns {boolean}
 */
export function isLoopbackRemoteAddress(value) {
    return typeof value === 'string' && LOOPBACK_ADDRESSES.has(value.trim().toLowerCase());
}

/**
 * @param {Record<string, unknown>} headers Node 已经小写化的请求头
 * @returns {boolean} 是否带任何代理 / 转发标记
 */
export function hasForwardingHeaders(headers) {
    if (headers === null || typeof headers !== 'object') return false;
    for (const name of Object.keys(headers)) {
        const key = String(name).toLowerCase();
        if (FORWARDING_HEADERS.has(key)) return true;
        if (FORWARDING_PREFIXES.some((prefix) => key.startsWith(prefix))) return true;
    }
    return false;
}

/**
 * 解析 Host / Origin 里的 authority，只接受本机字面量与当前监听端口。
 *
 * 不能用 `new URL('http://' + text)` 偷懒：在 HTTPS 连接上，Host `:80` 会被
 * 当成 http 的默认端口而归一化成空，随后又被误读成 :443，等于放过了一个
 * 指向本机 80 端口的 authority。这里按实际 scheme 直接校验字面量语法。
 *
 * @param {unknown} raw 形如 `127.0.0.1:3080`、`localhost:3080`、`[::1]:3080`
 * @param {number|string|undefined} localPort `req.socket.localPort`（必须是有效端口）
 * @param {'http'|'https'} scheme 当前连接的实际 scheme（TLS 终止后为 https）
 * @returns {{hostname: string, port: number} | undefined}
 */
export function parseLocalAuthority(raw, localPort, scheme) {
    if (scheme !== 'http' && scheme !== 'https') return undefined;
    // 实际监听端口必须真实存在且有效；缺了它就不能拿 Host 去和「默认端口」比。
    const actualPort = Number(localPort);
    if (!Number.isInteger(actualPort) || actualPort < 1 || actualPort > 65535) return undefined;

    const text = typeof raw === 'string' ? raw : String(raw ?? '');
    // 不 trim：前后空白、路径、查询、逗号、userinfo 等任何多余字符都直接拒绝。
    if (text === '' || text.trim() !== text) return undefined;
    const match = AUTHORITY_PATTERN.exec(text);
    if (match === null) return undefined;

    const hostname = match[1].toLowerCase();
    if (!LOCAL_HOSTNAMES.has(hostname)) return undefined;
    const defaultPort = scheme === 'https' ? 443 : 80;
    const port = match[2] === undefined ? defaultPort : Number(match[2]);
    if (!Number.isInteger(port) || port < 1 || port > 65535) return undefined;
    // Host 必须指回这个连接真正监听的端口；这也挡住「Host 伪造到别的本机服务」。
    if (port !== actualPort) return undefined;
    return { hostname, port };
}

/**
 * 判断这次请求是否是可复用官方供应商 Key 的本机浏览器直连。
 *
 * 全部条件（fail-closed）：
 *   1. socket.remoteAddress 精确是 127.0.0.1 / ::1 / ::ffff:127.0.0.1；
 *   2. 没有任何 Forwarded / X-Forwarded-* / X-Real-IP / Via 等转发标记；
 *   3. Host 必须是 127.0.0.1 / localhost / [::1] 字面量（可带端口），端口必须
 *      等于真实存在的 socket.localPort，且按实际 http/https scheme 解析默认端口；
 *   4. POST 必须带 Origin，且 Origin 的 scheme 与实际连接一致（TLS 判断
 *      socket.encrypted），authority 与 Host 完全一致；
 *   5. 若带 Sec-Fetch-Site，必须恰好是 same-origin。
 *
 * 这不是认证：本机程序仍可伪造这些头。也不承诺识别「先剥掉转发头、再从
 * 127.0.0.1 转进来」的本机代理 —— 那种请求看起来与真实本机浏览器没有区别。
 *
 * @param {object} req node:http IncomingMessage
 * @returns {{ok: true, authority: string} | {ok: false, reason: string}}
 *   reason 只用于服务端内部分支，不回给浏览器。
 */
export function isLocalBrowserRequest(req) {
    const socket = req?.socket;
    if (socket === null || socket === undefined || !isLoopbackRemoteAddress(socket.remoteAddress)) {
        return { ok: false, reason: 'remote_not_local' };
    }
    const headers = req.headers;
    if (hasForwardingHeaders(headers)) return { ok: false, reason: 'proxy_headers' };

    const scheme = socket.encrypted === true ? 'https' : 'http';
    const host = parseLocalAuthority(headers?.host, socket.localPort, scheme);
    if (host === undefined) return { ok: false, reason: 'host_invalid' };

    const site = headers?.['sec-fetch-site'];
    if (site !== undefined && String(site).toLowerCase() !== 'same-origin') {
        return { ok: false, reason: 'fetch_site' };
    }

    const rawOrigin = headers?.origin;
    if (rawOrigin === undefined || rawOrigin === '') return { ok: false, reason: 'origin_missing' };
    let origin;
    try {
        origin = new URL(String(rawOrigin));
    } catch {
        return { ok: false, reason: 'origin_invalid' };
    }
    if (origin.protocol !== scheme + ':') return { ok: false, reason: 'origin_scheme' };
    // Origin 必须是纯 origin：不允许 userinfo / path / query / fragment。
    if (origin.username !== '' || origin.password !== ''
        || origin.pathname !== '/' || origin.search !== '' || origin.hash !== '') {
        return { ok: false, reason: 'origin_invalid' };
    }
    const originHost = parseLocalAuthority(origin.host, socket.localPort, scheme);
    if (originHost === undefined) return { ok: false, reason: 'origin_host_invalid' };
    if (originHost.hostname !== host.hostname || originHost.port !== host.port) {
        return { ok: false, reason: 'origin_mismatch' };
    }
    return { ok: true, authority: host.hostname + ':' + host.port };
}

/**
 * 统一的门：设置显式打开 + 官方端点 + 本机浏览器直连。
 * 只返回结论，不读配置 / 凭据，方便单测。
 *
 * @param {{enabled: boolean, officialEndpoint: boolean, req: object}} input
 * @returns {{ok: true, authority: string} | {ok: false, reason: 'disabled'|'custom_endpoint'|'nonlocal'}}
 */
export function officialReuseGate({ enabled, officialEndpoint, req }) {
    if (enabled !== true) return { ok: false, reason: 'disabled' };
    if (officialEndpoint !== true) return { ok: false, reason: 'custom_endpoint' };
    const local = isLocalBrowserRequest(req);
    return local.ok ? { ok: true, authority: local.authority } : { ok: false, reason: 'nonlocal' };
}
