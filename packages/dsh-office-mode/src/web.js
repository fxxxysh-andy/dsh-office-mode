/**
 * 联网检索与取正文 —— office 自己的那条通道，第二十轮起是**多条**。
 *
 * ## 为什么要有这个模块
 *
 * 「办公模式」是一个完整的组合（preset 的 plugins 列表就是那个会话的全部插件行），
 * 里面**没有** @deepseek-ai/dsh-tool-web。于是办公会话的工具面里根本没有
 * web_search / advanced_search / platform_search / web_fetch：
 * office_search_dispatch 把渠道派给子代理时会用工具白名单圈那几个名字，
 * 子代理一启动就报
 *   tools.restrict() names unknown global tools "web_search", …
 * 整条检索链 0/N 全部失败（2026-09-25 的真实会话 session6 就是这样）。
 *
 * ## 三层结构（第二十轮）
 *
 *   1. **接缝（seam）**：ctx.get('web') 在就用它。它带着宿主的取正文策略
 *      （公网地址校验与地址钉死、同源跳转、体积与超时上限、代理路由），首选。
 *   2. **多路通道（src/web-providers.js）**：Anthropic 兼容 + 原生 web_search、
 *      OpenAI 兼容、Tavily、Brave、博查、Exa、Serper、自建 SearXNG、
 *      DuckDuckGo（免 Key）。每条通道自己发 HTTP、自己解析，返回统一形状。
 *      单条通道是**可选的**：`search.provider` 指定一条就只用那一条；
 *      `auto`（默认）按 `search.providerOrder` 依次尝试，第一条成功的就用它。
 *      这样「只有 DeepSeek 一条路」不再是硬约束：换通道不用改代码，
 *      零配置的通路（DuckDuckGo）保证「刚装好就能查」。
 *   3. **取正文**：接缝优先，其次自带的安全抓取；两者拿到的页面都要过
 *      **成败哨兵**与**脚本式预处理**（见下）。
 *
 * 三层都不可用时给出可执行的说明（缺哪个 Key、怎么配），而不是沉默失败。
 *
 * ## 成败怎么判（第十八轮补：P0-8 / P0-9）
 *
 * 状态码**不足以**判断一次取正文成没成：区域封锁页、Cloudflare 拦截页与登录墙
 * 常常连状态码都是 200。本机实测：直连一个被墙站点的文档页拿到 **HTTP 200**，
 * 内容却是 447 KB 的「本区域不可用」—— 只看 `response.ok` 就会把封锁页当正文喂给模型，
 * 那比直接报错更糟，因为它看起来是成功的。所以每次取正文都过一遍
 * `inspectFetchedPage`（状态码 → 最终地址 → 正文特征 → 长度）。
 *
 * 抛出的 `WebAccessError.code` 经 `webFailureKind` 归到四类之一，反馈里按分类计数：
 * **配置缺失 / 网络出口不可达 / 目标站拒绝 / 没拿到结果**。四类的修法完全不同，
 * 压成一句「查不到」会让模型把「被墙」读成「环境不提供这项能力」（session6 就这么
 * 写进过长期记忆）。`test/web.mjs` 有一条漂移守卫：新加错误码必须登记分类。
 *
 * ## 取回来的页面先清洗（第二十轮补）
 *
 * 取正文的出口从「`htmlToText` 一发了之」换成「哨兵判成败 → `preprocessPage`
 * 清洗 → 再进上下文」：
 *
 *   - 哨兵读**清洗前**的文本（清洗会把拦截页的样板一起删掉，用清洗后的文本判
 *     会把「目标站拒绝」误判成「正文几乎为空」）；
 *   - 返回给调用方的是**清洗后**的正文，附一份 `preprocess.stats`
 *     （去掉多少行、去重几行、留下多少字符），判错时看得见、也能退回 `off`。
 *
 * @module dsh-office-mode/web
 */
import { lookup as dnsLookup } from 'node:dns/promises';
import { isIP } from 'node:net';

import { DEFAULT_BUILTIN_WEB } from './config.js';
import {
    asText,
    describeWebFailure,
    messageOf,
    serviceOf,
    summarizeWebFailures,
    WebAccessError,
    webFailureKind,
    webFailureLabel,
    WEB_FAILURE_CODES,
    WEB_FAILURE_KINDS,
    withTimeout,
} from './web-errors.js';
import {
    mapSearchResponse,
    PROVIDER_IDS,
    providerOf,
    probeProvider,
    resolveApiKey,
    resolveProviderOptions,
    resolveProviderOrder,
    searchProvider,
} from './web-providers.js';
import { preprocessPage, resolvePreprocessOptions } from './web-preprocess.js';

/** 接缝引擎：走宿主的 ctx.web。 */
export const WEB_ENGINE_SEAM = 'seam';
/** 自带引擎：本模块（或某条通道）自己发 HTTP。 */
export const WEB_ENGINE_BUILTIN = 'builtin';

// 错误层与通道层的公开面在这里**原样再导出**：外部（tools.js / search.js / 测试）
// 从第十六轮起就只认 `./web.js` 这一个入口，第二十轮拆模块不改变这个契约。
export {
    describeWebFailure,
    mapSearchResponse,
    summarizeWebFailures,
    WebAccessError,
    webFailureKind,
    webFailureLabel,
    WEB_FAILURE_CODES,
    WEB_FAILURE_KINDS,
    PROVIDER_IDS,
    providerOf,
    resolveApiKey,
    resolveProviderOptions,
    resolveProviderOrder,
    searchProvider,
    resolvePreprocessOptions,
};
/**
 * 多条查询的结果按 URL 去重合并，整体截到 maxResults 条。
 *
 * 第十六轮就在这个位置：直查（office_search_run）一条查询一次，多条查询的结果
 * 要合并成一份清单。与通道无关，所以留在编排层。
 */
export function mergeSources(results, maxResults) {
    const merged = [];
    const seen = new Set();
    for (const result of results ?? []) {
        for (const source of result?.sources ?? []) {
            const url = asText(source?.url);
            if (url === '' || seen.has(url)) continue;
            seen.add(url);
            merged.push({ ...source, url });
        }
    }
    if (Number.isFinite(maxResults) && merged.length > maxResults) {
        return { sources: merged.slice(0, maxResults), truncated: true };
    }
    return { sources: merged, truncated: false };
}

/**
 * 取宿主的 web 接缝（ctx.web）。
 *
 * 与 search.js 的 subagentsOf 同一条理由：cordis 的 ctx 代理读一个没写进
 * inject 的服务属性会**当场抛错**，所以一律经 ctx.get。接缝只是本插件的一项
 * 能力，缺了它其余能力照常，因此**不**写进模块级 inject。
 *
 * @param {object} [ctx] 插件上下文（cordis 的 ctx 或测试替身）
 * @returns {{search: Function, fetch: Function}|undefined}
 */
export function webSeamOf(ctx) {
    const service = serviceOf(ctx, 'web');
    if (service === undefined || service === null) return undefined;
    if (typeof service.search !== 'function' || typeof service.fetch !== 'function') return undefined;
    return service;
}

/**
 * 合并调用方给的选项与默认值。
 *
 * 接受的形状有**两种**，这是刻意的：
 *   - 第十六～十九轮的形状：直接传 `search.builtin` 那一份参数（老配置与老测试）；
 *   - 第二十轮的形状：传整个 `search` 段（含 provider / providerOrder /
 *     providers / preprocess）。
 * 判据是「有没有那些新键」。返回对象仍然是**扁平的既有一组字段**（apiKey、
 * baseURL、model、maxResults…），因为外部一直按扁平字段读它；同时多出
 * `provider` / `providerOrder` / `providers` / `preprocess` 四个新字段。
 */
export function resolveBuiltinOptions(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const isSearchConfig = ['builtin', 'provider', 'providerOrder', 'providers', 'preprocess']
        .some((key) => Object.prototype.hasOwnProperty.call(source, key));
    const builtinRaw = isSearchConfig ? (source.builtin ?? {}) : source;
    const pick = (key, fallback) => (asText(builtinRaw[key]) !== '' ? asText(builtinRaw[key]) : fallback);
    const num = (key, fallback, min, max) => {
        const value = typeof builtinRaw[key] === 'number' ? builtinRaw[key] : Number.parseFloat(String(builtinRaw[key]));
        if (!Number.isFinite(value)) return fallback;
        return Math.min(max, Math.max(min, value));
    };
    const maxResultsDefault = DEFAULT_BUILTIN_WEB.maxResults;
    const providers = isSearchConfig && source.providers !== null && typeof source.providers === 'object' ? source.providers : {};
    return {
        apiKey: asText(builtinRaw.apiKey),
        apiKeyEnv: pick('apiKeyEnv', DEFAULT_BUILTIN_WEB.apiKeyEnv),
        baseURL: pick('baseURL', DEFAULT_BUILTIN_WEB.baseURL).replace(/\/+$/, ''),
        model: pick('model', DEFAULT_BUILTIN_WEB.model),
        apiVersion: pick('apiVersion', DEFAULT_BUILTIN_WEB.apiVersion),
        maxUses: num('maxUses', DEFAULT_BUILTIN_WEB.maxUses, 1, 20),
        maxTokens: num('maxTokens', DEFAULT_BUILTIN_WEB.maxTokens, 256, 16_384),
        searchTimeoutMs: num('searchTimeoutMs', DEFAULT_BUILTIN_WEB.searchTimeoutMs, 5_000, 300_000),
        fetchTimeoutMs: num('fetchTimeoutMs', DEFAULT_BUILTIN_WEB.fetchTimeoutMs, 5_000, 300_000),
        maxResults: num('maxResults', maxResultsDefault, 1, 20),
        fetchPages: num('fetchPages', DEFAULT_BUILTIN_WEB.fetchPages, 0, 5),
        maxBytes: num('maxBytes', DEFAULT_BUILTIN_WEB.maxBytes, 32_768, 32 * 1024 * 1024),
        maxChars: num('maxChars', DEFAULT_BUILTIN_WEB.maxChars, 1_000, 400_000),
        maxRedirects: num('maxRedirects', DEFAULT_BUILTIN_WEB.maxRedirects, 0, 10),
        userAgent: pick('userAgent', DEFAULT_BUILTIN_WEB.userAgent),
        provider: isSearchConfig && asText(source.provider) !== '' ? asText(source.provider) : 'auto',
        providerOrder: isSearchConfig
            ? resolveProviderOrder('auto', source.providerOrder)
            : [...PROVIDER_IDS],
        providers,
        preprocess: resolvePreprocessOptions(isSearchConfig ? source.preprocess : undefined),
    };
}

/**
 * 某条通道这一次要用的参数：`search.builtin` 是 Anthropic 通道的老家，
 * `search.providers.<id>` 是每条通道各自的家（后者覆盖前者）。
 */
export function providerSourceOf(options, id) {
    const overrides = options?.providers?.[id];
    if (id === 'anthropic') return { ...options, ...(overrides ?? {}) };
    return overrides ?? {};
}

/** 一条通道在反馈里的中文名。 */
export function providerLabelOf(id) {
    return providerOf(id)?.label ?? id;
}

/** 引擎 + 通道 → 一句话（反馈与结果文件用）。 */
export function engineDescription(engine, provider) {
    if (engine === WEB_ENGINE_SEAM) return '宿主 web 服务（seam）';
    return providerLabelOf(provider ?? 'anthropic');
}

// ── 取正文：自带的 HTTP 实现 ────────────────────────────────────────────────

/** 允许抓取的协议与 URL 上限（与宿主 web-fetch-http 的策略同一口径）。 */
const MAX_URL_LENGTH = 2048;

/** 只在测试里替换：DNS 解析与 fetch。 */
const defaultNetwork = { lookup: dnsLookup, fetch: (...args) => globalThis.fetch(...args) };

/**
 * 判断一个 IP 字面量是不是公网可路由地址。
 *
 * 这一段是自带的取正文唯一的 SSRF 防线：环回、私网、链路本地、CGNAT、
 * 组播、保留段与 IPv6 的 ULA/链路本地/NAT64 前缀都不许连。
 * 宿主接缝在的时候用不到它（那边的策略更严：解析后把地址钉死在连接上）。
 */
export function isPublicAddress(input) {
    const text = asText(input).replace(/^\[|\]$/g, '');
    const family = isIP(text);
    if (family === 4) {
        const parts = text.split('.').map((part) => Number.parseInt(part, 10));
        if (parts.length !== 4 || parts.some((part) => !Number.isInteger(part) || part < 0 || part > 255)) return false;
        const [a, b, c] = parts;
        if (a === 0 || a === 10 || a === 127) return false;
        if (a === 100 && b >= 64 && b <= 127) return false;
        if (a === 169 && b === 254) return false;
        if (a === 172 && b >= 16 && b <= 31) return false;
        if (a === 192 && b === 168) return false;
        if (a === 192 && b === 0 && c === 0) return false;
        if (a === 192 && b === 0 && c === 2) return false;
        if (a === 198 && (b === 18 || b === 19)) return false;
        if (a === 198 && b === 51 && c === 100) return false;
        if (a === 203 && b === 0 && c === 113) return false;
        if (a >= 224) return false;
        return true;
    }
    if (family === 6) {
        const groups = expandIpv6(text);
        if (groups === undefined) return false;
        const [g0, g1] = groups;
        const allZero = groups.every((group) => group === 0);
        if (allZero) return false;                                   // ::
        if (groups.slice(0, 7).every((group) => group === 0) && groups[7] === 1) return false; // ::1
        if ((g0 & 0xfe00) === 0xfc00) return false;                  // fc00::/7 ULA
        if ((g0 & 0xffc0) === 0xfe80) return false;                  // fe80::/10 链路本地
        if ((g0 & 0xff00) === 0xff00) return false;                  // ff00::/8 组播
        if (g0 === 0x2001 && g1 === 0x0db8) return false;            // 2001:db8::/32 文档用
        if (g0 === 0x0064 && g1 === 0xff9b) return false;            // 64:ff9b::/96 NAT64
        if (g0 === 0x2002) return false;                             // 2002::/16 6to4
        if (groups.slice(0, 5).every((group) => group === 0) && groups[5] === 0xffff) {
            // ::ffff:a.b.c.d —— 按内嵌的 IPv4 判。
            const embedded = [groups[6] >> 8, groups[6] & 0xff, groups[7] >> 8, groups[7] & 0xff].join('.');
            return isPublicAddress(embedded);
        }
        return true;
    }
    return false;
}

/** 把 IPv6 文本展开成 8 组 16 位整数；形状不认识时返回 undefined。 */
function expandIpv6(text) {
    const zoneFree = text.split('%')[0];
    const halves = zoneFree.split('::');
    if (halves.length > 2) return undefined;
    const parse = (part) => (part === '' ? [] : part.split(':').map((piece) => {
        if (piece.includes('.')) {
            const octets = piece.split('.').map((octet) => Number.parseInt(octet, 10));
            if (octets.length !== 4 || octets.some((octet) => !Number.isInteger(octet) || octet < 0 || octet > 255)) return undefined;
            return [(octets[0] << 8) | octets[1], (octets[2] << 8) | octets[3]];
        }
        if (!/^[0-9a-fA-F]{1,4}$/.test(piece)) return undefined;
        return [Number.parseInt(piece, 16)];
    }));
    const left = parse(halves[0]);
    const right = halves.length === 2 ? parse(halves[1]) : [];
    if (left === undefined || right === undefined) return undefined;
    const flatLeft = left.flat();
    const flatRight = right.flat();
    if (flatLeft.includes(undefined) || flatRight.includes(undefined)) return undefined;
    if (halves.length === 2) {
        const fill = 8 - flatLeft.length - flatRight.length;
        if (fill < 1) return undefined;
        return [...flatLeft, ...new Array(fill).fill(0), ...flatRight];
    }
    return flatLeft.length === 8 ? flatLeft : undefined;
}

/** 校验 URL：只许 http(s)、不许内嵌账号密码、长度受限。 */
export function validateUrl(input) {
    const text = asText(input);
    if (text.length > MAX_URL_LENGTH) throw new WebAccessError(`URL 超过 ${MAX_URL_LENGTH} 字符上限。`, 'OFFICE_WEB_INVALID_URL');
    let url;
    try {
        url = new URL(text);
    } catch (error) {
        throw new WebAccessError(`不是合法的 URL：${text}`, 'OFFICE_WEB_INVALID_URL', { cause: error });
    }
    if (url.protocol !== 'http:' && url.protocol !== 'https:') {
        throw new WebAccessError(`只支持 http / https，收到 ${url.protocol}`, 'OFFICE_WEB_INVALID_URL');
    }
    if (url.username !== '' || url.password !== '') {
        throw new WebAccessError('URL 里不允许带账号密码。', 'OFFICE_WEB_BLOCKED_URL');
    }
    return url;
}

/** 主机名必须解析到公网地址（IP 字面量直接判）。 */
async function assertPublicHost(hostname, network) {
    const bare = hostname.replace(/^\[|\]$/g, '');
    if (isIP(bare) !== 0) {
        if (!isPublicAddress(bare)) throw new WebAccessError(`主机 ${hostname} 不是公网地址，拒绝连接。`, 'OFFICE_WEB_BLOCKED_URL');
        return;
    }
    let addresses;
    try {
        addresses = await network.lookup(bare, { all: true, verbatim: true });
    } catch (error) {
        throw new WebAccessError(`解析不了主机 ${hostname}：${messageOf(error)}`, 'OFFICE_WEB_NETWORK');
    }
    if (!Array.isArray(addresses) || addresses.length === 0) {
        throw new WebAccessError(`主机 ${hostname} 没有解析到任何地址。`, 'OFFICE_WEB_NETWORK');
    }
    for (const entry of addresses) {
        if (!isPublicAddress(entry?.address)) {
            throw new WebAccessError(`主机 ${hostname} 解析到非公网地址（${entry?.address}），拒绝连接。`, 'OFFICE_WEB_BLOCKED_URL');
        }
    }
}

/** 把 Content-Type 归类成可读的正文类型；不支持的返回 undefined。 */
export function classifyContentType(contentType) {
    const mime = asText(contentType).replace(/;.*$/s, '').trim().toLowerCase();
    if (mime === 'text/html' || mime === 'application/xhtml+xml') return 'html';
    if (mime.startsWith('text/')) return 'text';
    if (mime === 'application/json' || mime === 'application/xml' || mime.endsWith('+json') || mime.endsWith('+xml')) return 'text';
    return undefined;
}

/** 从 Content-Type 里取 charset。 */
export function charsetOf(contentType) {
    return /;\s*charset\s*=\s*"?([^";]+)"?/i.exec(asText(contentType))?.[1]?.trim().toLowerCase();
}

/** 按声明的编码解码；编码不认识时报错而不是给出乱码。 */
function decodeBytes(bytes, contentType) {
    const charset = charsetOf(contentType);
    try {
        return new TextDecoder(charset ?? 'utf-8').decode(bytes);
    } catch (error) {
        throw new WebAccessError(`响应声明的编码 ${charset} 不认识。`, 'OFFICE_WEB_UNSUPPORTED_TYPE', { cause: error });
    }
}

/** 读取响应体，最多 maxBytes 字节（超出即截断，不整包吞下）。 */
async function readCapped(response, maxBytes, signal) {
    if (response.body === null || response.body === undefined) return { bytes: new Uint8Array(0), truncated: false };
    const reader = response.body.getReader();
    const chunks = [];
    let total = 0;
    let truncated = false;
    for (;;) {
        if (signal?.aborted === true) {
            await reader.cancel().catch(() => {});
            throw new WebAccessError('取正文被取消。', 'OFFICE_WEB_ABORTED');
        }
        const { done, value } = await reader.read();
        if (done) break;
        const remaining = maxBytes - total;
        if (value.byteLength > remaining) {
            chunks.push(value.subarray(0, remaining));
            total += remaining;
            truncated = true;
            break;
        }
        chunks.push(value);
        total += value.byteLength;
    }
    await reader.cancel().catch(() => {});
    const bytes = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
        bytes.set(chunk, offset);
        offset += chunk.byteLength;
    }
    return { bytes, truncated };
}

// ── 内容级成败哨兵（P0-8）──────────────────────────────────────────────────

/**
 * 正文短到什么程度算「几乎为空」。
 *
 * 定得很低（20 字符）是**故意保守**：真正的拦截页、区域封锁页与 JS 空壳几乎都会
 * 落在这条线以下，而合法但简短的页面不该被误判 —— 误判的代价是丢掉一个真实来源，
 * 比漏判更贵。20–200 字符之间只给 `notice`，不判失败。
 */
const MIN_USEFUL_CHARS = 20;
/** 短于这个长度只提醒（不当失败）。 */
const SHORT_NOTICE_CHARS = 200;
/** 超过这个长度的正文不再做关键词扫描：避免长正文里一句巧合短语把真实来源判掉。 */
const SIGNATURE_SCAN_MAX = 4_000;

/** 命中即判「目标站拒绝」的正文特征。 */
const BLOCK_SIGNATURES = [
    { re: /attention required/i, what: 'Cloudflare 拦截页' },
    { re: /you have been blocked/i, what: 'Cloudflare 拦截页' },
    { re: /cf-error-details|cf_chl_/i, what: 'Cloudflare 拦截页' },
    { re: /just a moment\.\.\./i, what: 'Cloudflare 人机校验页' },
    { re: /enable javascript and cookies to continue/i, what: 'Cloudflare 人机校验页' },
    { re: /checking your browser before accessing/i, what: 'Cloudflare 人机校验页' },
    { re: /verify you are human/i, what: '人机校验页' },
    { re: /g-recaptcha|hcaptcha/i, what: '验证码页' },
    { re: /unavailable in your (region|country)/i, what: '区域封锁页' },
    { re: /not available in your (region|country)/i, what: '区域封锁页' },
    { re: /sign in to continue|please (sign|log) in to/i, what: '登录墙' },
    { re: /accounts\.google\.com\/o\/oauth2|consent\.google\.com/i, what: '登录/同意墙' },
];

/** 最终地址本身的特征 —— 比正文更可靠，区域封锁页常常连状态码都是 200。 */
const BLOCK_URL_SIGNATURES = [
    { re: /app-unavailable-in-region|unavailable-in-region/i, what: '区域封锁页' },
    { re: /challenge-platform|\/cdn-cgi\/challenge/i, what: '人机校验页' },
];

/**
 * 判定一次取正文的结果到底算不算成功 —— **不能只看状态码**。
 *
 * 2026-09-26 实测（这就是本函数存在的理由）：同一个 URL 在没走代理时返回
 * **HTTP 200**，内容却是 `claude.com/app-unavailable-in-region`（447 KB 的区域封锁页）。
 * 只看 `response.ok` 会把这页当成正文喂给模型 —— 比直接报错更糟，因为它看起来是成功的。
 *
 * 判定顺序按可靠性排：状态码 → 最终地址 → 正文特征 → 长度。
 * 传入的必须是**清洗前**的正文文本（见模块头「取回来的页面先清洗」）。
 *
 * @param {{statusCode?: number, url?: string, content?: string}} input
 * @returns {{ok: true, notice?: string}|{ok: false, code: string, reason: string}}
 */
export function inspectFetchedPage(input) {
    const statusCode = Number.isFinite(input?.statusCode) ? Number(input.statusCode) : 0;
    const url = asText(input?.url);
    const content = asText(input?.content);
    const length = content.length;

    // 1) 状态码最可靠。接缝可能回 0（未知），这时不判，交给后面几条。
    if (statusCode >= 400) {
        if (statusCode >= 500) {
            return { ok: false, code: 'OFFICE_WEB_NETWORK', reason: `目标站返回 HTTP ${statusCode}（服务端错误）。` };
        }
        const why = statusCode === 404 || statusCode === 410 ? '页面不存在' : '被挡或需要授权';
        return { ok: false, code: 'OFFICE_WEB_BLOCKED', reason: `目标站拒绝：HTTP ${statusCode}（${why}）。` };
    }

    // 2) 最终地址的特征。
    for (const { re, what } of BLOCK_URL_SIGNATURES) {
        if (re.test(url)) {
            return { ok: false, code: 'OFFICE_WEB_BLOCKED', reason: `抓到的不是原文，是${what}：${url}` };
        }
    }

    // 3) 正文特征：只在短文上扫。
    if (length <= SIGNATURE_SCAN_MAX) {
        for (const { re, what } of BLOCK_SIGNATURES) {
            if (re.test(content)) {
                return { ok: false, code: 'OFFICE_WEB_BLOCKED', reason: `抓到的不是原文，是${what}。` };
            }
        }
    }

    // 4) 几乎为空。
    if (length < MIN_USEFUL_CHARS) {
        return { ok: false, code: 'OFFICE_WEB_EMPTY', reason: `正文几乎为空（清洗后 ${length} 字符）。` };
    }

    if (length < SHORT_NOTICE_CHARS) {
        return { ok: true, notice: `正文偏短（${length} 字符），当来源用之前先打开核对。` };
    }
    return { ok: true };
}

const REDIRECT_STATUS = new Set([301, 302, 303, 307, 308]);

/** 预处理的结果接到取正文的返回值上：正文换成清洗后的那一份，报告一起带走。 */
function withPreprocess(input) {
    const result = preprocessPage(
        { html: input.kind === 'html' ? input.raw : '', text: input.content, url: input.url },
        input.options,
    );
    // 清洗后几乎为空时**退回清洗前**：哨兵已经判过成败，能到这里说明原文是正文；
    // 此时丢了内容比留着噪声更糟（例如页面本身就是一张清单）。
    const text = asText(result.text) === '' ? input.content : result.text;
    return {
        content: text,
        preprocess: {
            mode: result.mode,
            stats: { ...result.stats, outputChars: text.length },
            ...(result.meta.title === '' && result.meta.site === '' ? {} : { meta: result.meta }),
        },
    };
}

/**
 * 自带的取正文：校验地址 → GET → 只跟随同源跳转 → 限长解码 → 哨兵 → 清洗。
 *
 * 与宿主 web-fetch-http 的差别要照实说：那边解析一次 DNS 后把地址**钉**在连接上，
 * 本实现只做「连接前校验解析结果」，理论上挡不住 DNS 重绑定。所以接缝在的时候
 * 优先用接缝；这里是没有接缝时才走的兜底。
 */
export async function httpFetch(input, options, signal, network = defaultNetwork) {
    const url = validateUrl(input);
    const requestSignal = withTimeout(signal, options.fetchTimeoutMs);
    let current = url;
    let hop = 0;
    for (;;) {
        await assertPublicHost(current.hostname, network);
        let response;
        try {
            response = await network.fetch(current.toString(), {
                method: 'GET',
                redirect: 'manual',
                headers: {
                    'user-agent': options.userAgent,
                    accept: 'text/html,application/xhtml+xml,text/*;q=0.9,application/json;q=0.8',
                },
                signal: requestSignal,
            });
        } catch (error) {
            if (signal?.aborted === true) throw new WebAccessError('取正文被取消。', 'OFFICE_WEB_ABORTED', { cause: error });
            throw new WebAccessError(`取正文失败：${messageOf(error)}`, 'OFFICE_WEB_NETWORK', { cause: error });
        }
        if (REDIRECT_STATUS.has(response.status)) {
            const location = response.headers.get('location');
            await response.body?.cancel?.().catch(() => {});
            if (location === null || location === undefined) {
                throw new WebAccessError(`HTTP ${response.status} 跳转但没有 Location 头。`, 'OFFICE_WEB_NETWORK');
            }
            if (hop >= options.maxRedirects) {
                throw new WebAccessError(`跳转超过 ${options.maxRedirects} 次上限。`, 'OFFICE_WEB_REDIRECT');
            }
            const next = validateUrl(new URL(location, current).toString());
            if (next.origin !== current.origin) {
                throw new WebAccessError(`跨站跳转到 ${next.origin} 不自动跟随，请直接用那个地址。`, 'OFFICE_WEB_REDIRECT');
            }
            current = next;
            hop += 1;
            continue;
        }
        // 4xx/5xx 先判、先抛：不必把一个拦截页的正文读完（上限之内也是浪费），
        // 更不能把它当正文返回。200 的那些拦截页由后面的完整哨兵兜。
        if (response.status >= 400) {
            await response.body?.cancel?.().catch(() => {});
            const statusVerdict = inspectFetchedPage({ statusCode: response.status, url: current.toString(), content: '' });
            throw new WebAccessError(statusVerdict.reason, statusVerdict.code);
        }
        const contentType = response.headers.get('content-type');
        const kind = classifyContentType(contentType);
        if (kind === undefined) {
            await response.body?.cancel?.().catch(() => {});
            throw new WebAccessError(`不支持的内容类型：${asText(contentType) || '未知'}`, 'OFFICE_WEB_UNSUPPORTED_TYPE');
        }
        const { bytes, truncated } = await readCapped(response, options.maxBytes, requestSignal).catch((error) => {
            if (error instanceof WebAccessError) throw error;
            // 读流中断：超时、调用方取消与网络断开在这里都会冒出来，按信号归类。
            const timedOut = signal?.aborted !== true && requestSignal?.aborted === true;
            throw new WebAccessError(
                timedOut ? `取正文超过 ${options.fetchTimeoutMs} 毫秒。` : `读取响应体失败：${messageOf(error)}`,
                timedOut ? 'OFFICE_WEB_TIMEOUT' : 'OFFICE_WEB_NETWORK',
                { cause: error },
            );
        });
        const decoded = decodeBytes(bytes, contentType);
        const clipped = decoded.length > options.maxChars;
        const raw = clipped ? decoded.slice(0, options.maxChars) : decoded;
        const content = kind === 'html' ? htmlToText(raw, options.maxChars) : raw;
        // 过哨兵：状态码、最终地址、正文特征、长度，任一命中就按失败抛出去，
        // 而不是把一个拦截页当作正文返回（P0-8）。判的是**清洗前**的文本。
        const verdict = inspectFetchedPage({ statusCode: response.status, url: current.toString(), content });
        if (verdict.ok !== true) throw new WebAccessError(verdict.reason, verdict.code);
        const processed = withPreprocess({ raw, content, kind, url: current.toString(), options: options.preprocess });
        return {
            engine: WEB_ENGINE_BUILTIN,
            url: current.toString(),
            statusCode: response.status,
            kind,
            content: processed.content,
            preprocess: processed.preprocess,
            truncated: truncated || clipped,
            ...(verdict.notice === undefined ? {} : { notice: verdict.notice }),
        };
    }
}

// ── HTML → 文本 ─────────────────────────────────────────────────────────────

const ENTITIES = {
    nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", ldquo: '“', rdquo: '”',
    mdash: '—', ndash: '–', hellip: '…', middot: '·', times: '×', copy: '©', reg: '®',
};

/** 解开常见实体（命名 + 十进制 + 十六进制）。 */
export function decodeEntities(text) {
    return String(text).replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body) => {
        if (body.startsWith('#x') || body.startsWith('#X')) {
            const code = Number.parseInt(body.slice(2), 16);
            return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        if (body.startsWith('#')) {
            const code = Number.parseInt(body.slice(1), 10);
            return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        const known = ENTITIES[body.toLowerCase()];
        return known ?? whole;
    });
}

/**
 * 把 HTML 收成可读文本。
 *
 * 不追求还原排版，只要求「读得出内容、链接留得下」：脚本/样式整段丢掉，
 * 块级标签换成换行，列表项加 `- `，标题加 `#`，链接写成 `[文字](地址)`，
 * 其余标签剥掉，实体解开，空行压缩。
 *
 * 第二十轮起它只用在**哨兵**与 **plain/off 模式**上：要进上下文的正文走
 * `web-preprocess.js` 那条管线（去样板 + 挑主容器 + 行级清洗）。
 *
 * @param {string} html 原始 HTML
 * @param {number} maxChars 结果上限
 */
export function htmlToText(html, maxChars = 20_000) {
    let text = String(html ?? '');
    text = text.replace(/<!--[\s\S]*?-->/g, '');
    text = text.replace(/<(script|style|noscript|template|svg|head)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, ' ');
    text = text.replace(/<a\b[^>]*href\s*=\s*("([^"]*)"|'([^']*)'|([^\s">]+))[^>]*>([\s\S]*?)<\/a\s*>/gi,
        (_whole, _attr, quoted, single, bare, inner) => {
            const label = String(inner).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
            const href = quoted ?? single ?? bare ?? '';
            if (href === '' || href.startsWith('#') || href.startsWith('javascript:')) return label === '' ? ' ' : ' ' + label + ' ';
            return ` [${label === '' ? href : label}](${href}) `;
        });
    text = text.replace(/<h([1-6])\b[^>]*>/gi, (_whole, level) => '\n\n' + '#'.repeat(Number(level)) + ' ');
    text = text.replace(/<\/(h[1-6]|p|div|section|article|header|footer|main|blockquote|table|tr|ul|ol|dl|pre|figure)\s*>/gi, '\n');
    text = text.replace(/<br\s*\/?>/gi, '\n');
    text = text.replace(/<li\b[^>]*>/gi, '\n- ');
    text = text.replace(/<t[dh]\b[^>]*>/gi, ' | ');
    text = text.replace(/<[^>]*>/g, ' ');
    text = decodeEntities(text);
    text = text.replace(/\r\n?/g, '\n');
    text = text.split('\n').map((line) => line.replace(/[ \t\u00a0]+/g, ' ').trim()).join('\n');
    text = text.replace(/\n{3,}/g, '\n\n').trim();
    if (text.length > maxChars) text = text.slice(0, maxChars);
    return text;
}

/** 取网页标题。 */
export function htmlTitle(html) {
    const matched = /<title[^>]*>([\s\S]*?)<\/title\s*>/i.exec(String(html ?? ''));
    return matched === null ? '' : decodeEntities(matched[1]).replace(/\s+/g, ' ').trim();
}

// ── 统一入口 ────────────────────────────────────────────────────────────────

/**
 * 造一个联网访问器：按通道顺序检索，取正文接缝优先，失败原因按通道归并。
 *
 * 返回对象上的每个方法都把「用了哪条通道」带在结果里（`engine` + `provider`），
 * 反馈里就能照实写；所有通道都失败时抛 WebAccessError，消息里逐条给出原因
 * ——「没配 Key」「连不上」「被挡」三类修法不同，压成一句会让人配错东西。
 *
 * @param {object} [ctx] 插件上下文
 * @param {object} [rawOptions] `search` 段或 `search.builtin`（两种形状都认）
 * @param {object} [hooks] 测试注入点：{ web, fetch, lookup, home }
 */
export function createWebAccess(ctx, rawOptions, hooks = {}) {
    const options = resolveBuiltinOptions(rawOptions);
    const network = {
        lookup: hooks.lookup ?? defaultNetwork.lookup,
        fetch: hooks.fetch ?? defaultNetwork.fetch,
    };
    // hooks.web 给单测直接塞一个假接缝；没给才去 ctx 里取（undefined 表示「没有」）。
    const seam = 'web' in hooks ? hooks.web : webSeamOf(ctx);

    /** 这一次调用里已经明确失败的通道：记住原因，不再重复试（`auto` 下很常见）。 */
    const failed = new Map();
    /** 每条通道解析后的参数（同一通道重复检索不必重算）。 */
    const optionCache = new Map();

    function optionsOf(id) {
        if (!optionCache.has(id)) optionCache.set(id, resolveProviderOptions(id, providerSourceOf(options, id)));
        return optionCache.get(id);
    }

    /** 一次接缝检索。 */
    async function seamSearch(query, request, maxResults) {
        const result = await seam.search({ query, ...(maxResults === undefined ? {} : { maxResults }) }, request.signal);
        const sources = Array.isArray(result?.sources)
            ? result.sources.map((source) => ({ ...source, url: asText(source?.url) })).filter((source) => source.url !== '')
            : [];
        return {
            engine: WEB_ENGINE_SEAM,
            provider: 'seam',
            providerLabel: providerLabelOf('seam'),
            content: asText(result?.content),
            sources,
            truncated: result?.truncated === true,
        };
    }

    /** 一次「自带通道」检索：先解析 Key，再交给对应适配器。 */
    async function channelSearch(id, query, request, maxResults) {
        const channelOptions = optionsOf(id);
        const apiKey = await resolveApiKey(ctx, channelOptions);
        const mapped = await searchProvider(id, {
            query,
            maxResults,
            options: channelOptions,
            fetchImpl: network.fetch,
            signal: request.signal,
            apiKey,
        });
        const sources = mapped.sources ?? [];
        const capped = Number.isFinite(maxResults) && sources.length > maxResults ? sources.slice(0, maxResults) : sources;
        return {
            engine: WEB_ENGINE_BUILTIN,
            provider: id,
            providerLabel: providerLabelOf(id),
            content: mapped.content ?? '',
            sources: capped,
            truncated: capped.length !== sources.length,
        };
    }

    /**
     * 每条通道现在能不能用（不联网，除了接缝那一条只看服务在不在）。
     * 设置页与失败提示都靠它说清「还差哪一步」。
     */
    const probeChannels = async () => {
        const list = [];
        for (const id of PROVIDER_IDS) {
            if (failed.has(id)) {
                list.push({ id, label: providerLabelOf(id), ok: false, reason: failed.get(id) });
                continue;
            }
            const verdict = await probeProvider(id, optionsOf(id), ctx, seam !== undefined);
            list.push({ id, label: providerLabelOf(id), ok: verdict.ok, reason: verdict.reason });
        }
        return list;
    };

    /** 当前可用的引擎与理由（不联网，设置页与反馈都用它）。 */
    const probe = async () => {
        const channels = await probeChannels();
        // 候选集合就是**这次真的会按顺序试的那些通道**：点名一条时只看那一条 ——
        // 否则「指名 anthropic 但没配 Key」会被别的免 Key 通道兜成「可用」，
        // 调用方以为配好了，实际每次检索都在报错。
        const order = options.provider === 'auto' ? options.providerOrder : [options.provider];
        const candidates = order.map((id) => channels.find((item) => item.id === id)).filter((item) => item !== undefined);
        const usable = candidates.filter((item) => item.ok);
        if (usable.length === 0) {
            return {
                ok: false,
                engine: null,
                provider: null,
                channels,
                reason: candidates.map((item) => item.label + '（' + item.reason + '）').join('；'),
            };
        }
        const first = usable[0];
        return {
            ok: true,
            engine: first.id === 'seam' ? WEB_ENGINE_SEAM : WEB_ENGINE_BUILTIN,
            provider: first.id,
            channels,
            reason: first.reason + (usable.length > 1 ? `；另有 ${usable.length - 1} 条可用（可在设置页换）` : ''),
        };
    };

    /** 把所有通道的失败原因压成一句可执行的话。 */
    function failureSummary(errors) {
        const parts = errors.map((item) => `${item.label}：${describeWebFailure(item.error)}`);
        const kinds = [...new Set(errors.map((item) => webFailureKind(item.error)))];
        const advice = kinds.includes('config')
            ? '缺 Key 或端点的通道去设置页「办公模式 → 检索编排」配；'
            : '';
        return `检索不到结果（${summarizeWebFailures(errors.map((item) => item.error))}）—— ${parts.join('；')}。${advice}`
            + '也可以直接用一条免 Key 的通道（duckduckgo / searxng），或在设置页把 provider 固定成某一条。';
    }

    return {
        probe,
        /** 只看某条通道的可用性（设置页与单测用）。 */
        async probeChannel(id) {
            return probeProvider(id, optionsOf(id), ctx, seam !== undefined);
        },

        /**
         * 一次检索：按通道顺序试，第一条成功的就是结果。
         *
         * `request.provider` 给了就只走那一条（用户/模型明确点名的通道不该被
         * 悄悄换掉）；没给就按 `search.provider` / `providerOrder` 走。
         */
        async search(query, request = {}) {
            const text = asText(query);
            if (text === '') throw new WebAccessError('检索关键词不能为空。', 'OFFICE_WEB_INVALID_QUERY');
            const maxResults = Number.isFinite(request.maxResults) ? request.maxResults : options.maxResults;
            const explicit = asText(request.provider);
            let order;
            try {
                order = resolveProviderOrder(explicit !== '' ? explicit : options.provider, options.providerOrder);
            } catch (error) {
                throw error;
            }
            const errors = [];
            for (const id of order) {
                if (failed.has(id)) {
                    errors.push({ id, label: providerLabelOf(id), error: new WebAccessError(failed.get(id), 'OFFICE_WEB_ERROR') });
                    continue;
                }
                try {
                    const result = id === 'seam'
                        ? (seam === undefined
                            ? (() => { throw new WebAccessError('组合里没有 web 服务（ctx.web）。', 'OFFICE_WEB_PROVIDER'); })()
                            : await seamSearch(text, request, maxResults))
                        : await channelSearch(id, text, request, maxResults);
                    return result;
                } catch (error) {
                    const message = messageOf(error);
                    failed.set(id, message);
                    errors.push({ id, label: providerLabelOf(id), error });
                }
            }
            throw new WebAccessError(failureSummary(errors), errors[0]?.error?.code ?? 'OFFICE_WEB_ERROR', { cause: errors[0]?.error });
        },

        /**
         * 取一个 URL 的正文：接缝优先，失败就退到自带实现。
         *
         * `request.preprocess` 可以按次覆盖预处理模式（'article' / 'plain' / 'off'），
         * 用来给「这次要原始一点」的调用留出口。
         */
        async fetch(url, request = {}) {
            const preprocessOptions = request.preprocess === undefined
                ? options.preprocess
                : resolvePreprocessOptions({ ...options.preprocess, mode: request.preprocess });
            const errors = [];
            if (seam !== undefined && !failed.has('seam')) {
                try {
                    const result = await seam.fetch({ url }, request.signal);
                    const body = result?.body ?? {};
                    const content = asText(body.content);
                    const finalUrl = asText(result?.url) || url;
                    const statusCode = Number.isFinite(result?.statusCode) ? result.statusCode : 0;
                    const kind = body.kind === 'html' ? 'html' : 'text';
                    const text = kind === 'html' ? htmlToText(content, options.maxChars) : content.slice(0, options.maxChars);
                    // 接缝也可能把拦截页当正文回来（它直接回响应体），所以这里同样过哨兵；
                    // 判定失败就落进 catch → 换自带实现再试一次（两条路的出口本来就不同）。
                    const verdict = inspectFetchedPage({ statusCode, url: finalUrl, content: text });
                    if (verdict.ok !== true) throw new WebAccessError(verdict.reason, verdict.code);
                    const processed = withPreprocess({ raw: content, content: text, kind, url: finalUrl, options: preprocessOptions });
                    return {
                        engine: WEB_ENGINE_SEAM,
                        provider: 'seam',
                        url: finalUrl,
                        statusCode,
                        kind,
                        content: processed.content,
                        preprocess: processed.preprocess,
                        truncated: result?.truncated === true,
                        ...(verdict.notice === undefined ? {} : { notice: verdict.notice }),
                    };
                } catch (error) {
                    failed.set('seam', messageOf(error));
                    errors.push('宿主 web 服务：' + messageOf(error));
                }
            } else if (failed.has('seam')) {
                errors.push('宿主 web 服务：' + failed.get('seam'));
            }
            try {
                return await httpFetch(url, { ...options, preprocess: preprocessOptions }, request.signal, network);
            } catch (error) {
                errors.push('自带取正文：' + messageOf(error));
                throw new WebAccessError(
                    `取不到正文（${webFailureLabel(error)}）—— ${errors.join('；')}`,
                    error?.code ?? 'OFFICE_WEB_ERROR',
                    { cause: error },
                );
            }
        },

        /** 供反馈与 office_help 用的一句话状态。 */
        async describe() {
            const status = await probe();
            if (status.ok) return `内置检索：可用（${status.reason}）`;
            return `内置检索：不可用（${status.reason}）`;
        },

        /** 引擎、通道与解析后的参数，测试和调试用。 */
        options,
    };
}

export { messageOf as webMessageOf };
