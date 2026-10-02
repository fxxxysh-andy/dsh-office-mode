/**
 * 联网通道的错误层：错误类型 + 失败分类 + 几个共用小工具。
 *
 * ## 为什么单独一个模块
 *
 * 第二十轮起联网通道不再只有一条：宿主 web 接缝（seam）、Anthropic 兼容 +
 * 原生 web_search、OpenAI 兼容、Tavily、Brave、博查、Exa、Serper、SearXNG、
 * DuckDuckGo…… 通道实现各自 throw，但「失败该怎么说」必须是**通道无关**的
 * —— 反馈层（office_search_run / 派工结果文件）不该为每个通道写一套措辞，
 * 也不该因为换了通道就把「配置缺失」说成「没搜到」。
 *
 * 所以错误类型与分类下沉到这一层：`web-errors.js` 不依赖任何通道实现，
 * `web.js`（通道编排 + 取正文）与 `web-providers.js`（各通道适配器）都从这里取。
 * `web.js` 仍然把这些名字**原样再导出**，外部（测试与 tools.js）的导入面不变。
 *
 * ## 失败分类（P0-9 的口径，通道多了以后更重要）
 *
 * 四类的**修法完全不同**，压成一句「查不到」会让模型把「被墙」读成
 * 「环境不提供这项能力」：
 *
 *   config  配置缺失      缺 Key、端点或凭据没配        → 去设置页配
 *   network 网络出口不可达  连不上、超时、5xx、跳转       → 换时间或换出口再试
 *   blocked 目标站拒绝     401/403/429/451、区域封锁页、
 *                          拦截页、验证码、页面不存在    → **换一个来源**，不要重试同一个
 *   empty   没拿到结果     接口没触发检索、正文几乎为空   → 换关键词或换角度
 *
 * @module dsh-office-mode/web-errors
 */

/** 联网访问的错误类型：`code` 用来把「没 Key」与「网络不通」分开。 */
export class WebAccessError extends Error {
    constructor(message, code, options) {
        super(message, options);
        this.name = 'WebAccessError';
        this.code = code ?? 'OFFICE_WEB_ERROR';
    }
}

export const WEB_FAILURE_KINDS = Object.freeze({
    config: '配置缺失',
    network: '网络出口不可达',
    blocked: '目标站拒绝',
    empty: '没拿到结果',
    other: '其它',
});

/** 错误码 → 失败分类。没有登记的一律落 `other`（宁可说「其它」也不要猜）。 */
const KIND_BY_CODE = Object.freeze({
    OFFICE_WEB_NO_KEY: 'config',
    OFFICE_WEB_PROVIDER: 'config',
    // 自建 SearXNG 这类通道必须给端点：没配是配置问题，不是网络问题。
    OFFICE_WEB_NO_BASE: 'config',
    OFFICE_WEB_NETWORK: 'network',
    OFFICE_WEB_TIMEOUT: 'network',
    OFFICE_WEB_REDIRECT: 'network',
    OFFICE_WEB_BLOCKED: 'blocked',
    OFFICE_WEB_HTTP_ERROR: 'blocked',
    // 地址被 SSRF 防线挡下：对调用方来说动作与「目标站拒绝」一样 —— 换个来源。
    // 具体原因仍在 message 里（「不是公网地址，拒绝连接」），没有被这句话吃掉。
    OFFICE_WEB_BLOCKED_URL: 'blocked',
    OFFICE_WEB_EMPTY: 'empty',
    OFFICE_WEB_NO_RESULTS: 'empty',
    OFFICE_WEB_UNSUPPORTED_TYPE: 'other',
    // 第二十九轮：来源是 PDF。没引擎 / 抽不出文本属于**两类修法**，所以分成两个码：
    // 「装 poppler 或 PyMuPDF」是配置问题，「扫描件没有文本层」是换来源。
    OFFICE_WEB_PDF_ENGINE: 'config',
    OFFICE_WEB_PDF_TOO_BIG: 'other',
    OFFICE_WEB_INVALID_QUERY: 'other',
    OFFICE_WEB_INVALID_URL: 'other',
    OFFICE_WEB_ABORTED: 'other',
    OFFICE_WEB_ERROR: 'other',
});

/**
 * 已经登记过分类的全部错误码。
 *
 * `test/web.mjs` 拿它当**漂移守卫**：从 `src/web*.js` 里扫出全部抛出的错误码，
 * 逐个断言在这个集合里。新加一个码却忘了给分类，会当场红 —— 否则新码会静默
 * 落进 `other`，「四类分开报」就在没人察觉的情况下退化成「一律其它」。
 */
export const WEB_FAILURE_CODES = Object.freeze(Object.keys(KIND_BY_CODE));

/** 一个异常属于哪一类失败。 */
export function webFailureKind(error) {
    const code = typeof error?.code === 'string' ? error.code.trim() : '';
    return KIND_BY_CODE[code] ?? 'other';
}

/** 那一类失败的中文名。 */
export function webFailureLabel(error) {
    return WEB_FAILURE_KINDS[webFailureKind(error)];
}

/** 「目标站拒绝：HTTP 403（被挡或需要授权）。」这种一行话，给反馈与结果文件用。 */
export function describeWebFailure(error) {
    return webFailureLabel(error) + '：' + messageOf(error);
}

/**
 * 把一批失败归并成「目标站拒绝 2、网络出口不可达 1」这样一行。
 *
 * 与第十七轮 dispatch 的「同因合并」同一条取向：**信息不丢，重复的字省掉**。
 * 元素可以是异常对象，也可以是已经算好的分类字符串。
 */
export function summarizeWebFailures(items) {
    const counts = new Map();
    for (const item of items ?? []) {
        const kind = typeof item === 'string' ? item : webFailureKind(item);
        counts.set(kind, (counts.get(kind) ?? 0) + 1);
    }
    if (counts.size === 0) return '';
    return [...counts.entries()].map(([kind, n]) => `${WEB_FAILURE_KINDS[kind] ?? kind} ${n}`).join('、');
}

/** 收敛成非空字符串。 */
export function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

/** 把任意异常收敛成一句可读的话。 */
export function messageOf(error) {
    const text = String((error && error.message) || error);
    return text.length > 300 ? text.slice(0, 300) + '…' : text;
}

/**
 * 取一个服务而不触发 cordis 的「未 inject 就读属性」异常。
 *
 * 宿主侧与客户端半侧同一条语义：上下文代理读一个没写进 `inject` 的服务属性会
 * **当场抛错**，不是返回 undefined。所以一律经 `ctx.get`，拿不到就当没有。
 */
export function serviceOf(ctx, name) {
    if (ctx === undefined || ctx === null || typeof ctx.get !== 'function') return undefined;
    try {
        return ctx.get(name);
    } catch {
        return undefined;
    }
}

/**
 * 把调用方的取消信号与一个超时信号合成一个。
 *
 * AbortSignal.any 在 Node 20.3 之前没有；没有它就退回「只有超时」——
 * 超时仍然在，只是少一层调用方取消。
 */
export function withTimeout(signal, timeoutMs) {
    if (typeof AbortSignal.timeout !== 'function') return signal;
    const timer = AbortSignal.timeout(timeoutMs);
    if (signal === undefined || signal === null) return timer;
    if (typeof AbortSignal.any === 'function') return AbortSignal.any([signal, timer]);
    return timer;
}
