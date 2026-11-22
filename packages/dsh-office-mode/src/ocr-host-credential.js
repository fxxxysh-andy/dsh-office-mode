/**
 * 官方 DeepSeek 供应商凭据的只读桥（第九十七轮）。
 *
 * 目的：用户在 设置 → 模型（Models）里给宿主官方供应商 `deepseek-official` 保存过
 * Key 之后，OCR 不必再让用户把同一份 Key 抄一遍。这里只做两件事：
 *
 *   1. 找到**确实注册着**的官方供应商那一行，并确认它当前实际生效的 baseURL
 *      origin 就是官方 `https://api.deepseek.com`；
 *   2. 从同一份实时配置里读出真正的 `apiKeyEnv`（凭据引用，默认 DEEPSEEK_API_KEY），
 *      再用宿主自己的 `ctx.credentials` 解析这个引用，把值交给一次 OCR 请求。
 *
 * 为什么必须查 baseURL：供应商行是可配置的。用户可以把它指向中转网关，这时
 * 保存的 Key 是网关的 token，不是 DeepSeek 官方 Key；把它发给官方 OCR 既会认证
 * 失败，也等于把网关凭据送到了另一个目的地。因此 origin 不精确等于
 * `https://api.deepseek.com` 时一律 fail closed，连凭据解析都不调用。
 *
 * baseURL 的解析与宿主适配器保持一致（dsh-llm-deepseek/src/config.ts）：
 *   - `effectiveConfig.baseURL.get()` 有值（含已由宿主归一化/注入的显式值）就是它；
 *   - 只有 getter 返回 undefined/null 才走宿主的后备顺序：
 *     `launchEnvironment.get('DEEPSEEK_BASE_URL')?.value`，再退到 PUBLIC_BASE_URL；
 *   - 后备阶段需要 `launchEnvironment` 服务。服务缺失时宿主会改读 process.env，
 *     而我们不替用户猜测未知环境，直接 fail closed；
 *   - PUBLIC_BASE_URL 取自 dsh-llm-deepseek/src/config.ts:116，值为
 *     `https://api.deepseek.com/anthropic`，origin 正是官方地址。
 *
 * 安全约束（按优先级）：
 *   - 只在 OCR 目标端点就是官方 `https://api.deepseek.com/v1/tools/ocr`、且供应商
 *     实际 baseURL 的 origin 也是官方时才继续；自定义 baseURL 一律不碰供应商 Key。
 *   - 不写 Anywhere：不调 `credentials.set`，不把 Key 放进配置 / 快照 / 日志。
 *   - 解析失败、服务缺失、配置不匹配都只返回结构化的「不可用」，异常原文不外发。
 *   - 没有 `credentials` 服务时只读宿主注入的 launchEnvironment（与
 *     dsh-llm-deepseek/src/credential.ts 的兜底同一份契约），不直接翻 process.env。
 *
 * @module dsh-office-mode/ocr-host-credential
 */

import { officialOcrEndpointOf } from './ocr.js';

/** 官方供应商的注册 route id。 */
export const HOST_OCR_PROVIDER = 'deepseek-official';
/** 唯一拥有该 route 的宿主插件名；名字不对就不读它的配置。 */
export const HOST_OCR_OWNER = '@deepseek-ai/dsh-llm-deepseek-api-key';
/** 合法凭据引用（与 dsh-credentials 的 REF_PATTERN 同形）。 */
const CREDENTIAL_REF = /^[A-Za-z_][A-Za-z0-9_]*$/;
/** cordis FiberState.ACTIVE（const enum 在纯 JS 里没有运行时导出）。 */
const FIBER_ACTIVE = 2;

/** 宿主缺失显式 baseURL 且没有环境覆盖时使用的公开默认值（源：config.ts:116）。 */
const HOST_PUBLIC_BASE_URL = 'https://api.deepseek.com/anthropic';
/** 宿主适配器读取的 baseURL 环境名（源：config.ts:119）。 */
const HOST_BASE_URL_ENV = 'DEEPSEEK_BASE_URL';
/** 唯一允许复用供应商 Key 的 supplier origin。 */
const HOST_OFFICIAL_ORIGIN = 'https://api.deepseek.com';

/**
 * 按宿主适配器的顺序解析供应商实际生效的 baseURL origin。
 *
 * 只用实时 `effectiveConfig` 与宿主注入的 launchEnvironment；不读进程环境，
 * 不读用户配置文件，也不解析 Key。失败一律返回不可用原因。
 *
 * @param {object} ctx 插件上下文
 * @param {object} config loader 条目的 fiber.effectiveConfig
 * @returns {{ok: true, origin: string} | {ok: false, reason: string}}
 */
function supplierEndpointOf(ctx, config) {
    let raw;
    try {
        const getter = config?.baseURL;
        if (getter === null || getter === undefined || typeof getter.get !== 'function') {
            return { ok: false, reason: 'provider_endpoint_unreadable' };
        }
        raw = getter.get();
    } catch {
        return { ok: false, reason: 'provider_endpoint_unreadable' };
    }

    // 宿主 resolveAdapterOptions 的 `??`：显式值（包括空串）优先；null/undefined 才后备。
    if (raw !== undefined && raw !== null) return originOfSupplierBaseURL(raw);

    let environment;
    try {
        environment = ctx?.get?.('launchEnvironment');
    } catch {
        return { ok: false, reason: 'provider_endpoint_unreadable' };
    }
    // 服务缺失时宿主会退到 process.env；我们不猜未知环境，直接 fail closed。
    if (environment === null || environment === undefined || typeof environment.get !== 'function') {
        return { ok: false, reason: 'provider_endpoint_unreadable' };
    }

    let hit;
    try {
        hit = environment.get(HOST_BASE_URL_ENV);
    } catch {
        return { ok: false, reason: 'provider_endpoint_unreadable' };
    }
    const environmentValue = hit?.value;
    // 与宿主的 `environment?.get(...)?.value ?? PUBLIC_BASE_URL` 完全一致。
    return originOfSupplierBaseURL(
        environmentValue === undefined || environmentValue === null ? HOST_PUBLIC_BASE_URL : environmentValue,
    );
}

/**
 * 把 supplier baseURL 收敛成 origin；只接受宿主也会接受的 HTTP(S) 根地址。
 *
 * @param {unknown} raw
 * @returns {{ok: true, origin: string} | {ok: false, reason: string}}
 */
function originOfSupplierBaseURL(raw) {
    if (typeof raw !== 'string' || raw.trim() === '') {
        return { ok: false, reason: 'provider_endpoint_unreadable' };
    }
    let parsed;
    try {
        parsed = new URL(raw);
    } catch {
        return { ok: false, reason: 'provider_endpoint_unreadable' };
    }
    // 与 dsh-llm-deepseek 的 baseURL 校验同形：HTTP(S)、无 userinfo/query/fragment。
    if ((parsed.protocol !== 'http:' && parsed.protocol !== 'https:')
        || parsed.username !== '' || parsed.password !== ''
        || parsed.search !== '' || parsed.hash !== '') {
        return { ok: false, reason: 'provider_endpoint_unreadable' };
    }
    if (parsed.origin !== HOST_OFFICIAL_ORIGIN) {
        return { ok: false, reason: 'provider_endpoint_not_official' };
    }
    return { ok: true, origin: parsed.origin };
}

/**
 * 找到官方供应商那一行（不读 Key、不读 baseURL）。
 *
 * @param {object} ctx
 * @returns {object | undefined}
 */
function hostOcrProviderEntry(ctx) {
    try {
        if (ctx === null || typeof ctx !== 'object' || typeof ctx.get !== 'function') return undefined;
        const llm = ctx.get('llm');
        if (llm === undefined || typeof llm.listProviders !== 'function'
            || typeof llm.listConfigurableProviders !== 'function') return undefined;
        if (!llm.listProviders().some((row) => row?.id === HOST_OCR_PROVIDER)) return undefined;

        const directory = llm.listConfigurableProviders()
            .find((row) => row?.provider === HOST_OCR_PROVIDER);
        // settingsPath 非空 = 这个 profile 嵌在别人的配置节里，不是官方那一行。
        if (directory === undefined || !Array.isArray(directory.settingsPath)
            || directory.settingsPath.length !== 0) return undefined;

        const loader = ctx.get('loader');
        if (loader === undefined || typeof loader.entries !== 'function') return undefined;
        const entry = [...loader.entries()].find((row) =>
            row?.options?.id === directory.settingsNs
            && row?.options?.name === HOST_OCR_OWNER
            && row?.fiber?.state === FIBER_ACTIVE);
        return entry ?? undefined;
    } catch {
        // 宿主服务、loader 行或配置形状异常都当作「没有可复用的 Key」；
        // 异常文字可能带配置细节，不能往上传。
        return undefined;
    }
}

/**
 * 从已确认的供应商条目里读实时的凭据引用。
 *
 * @param {object} entry
 * @returns {string | undefined}
 */
function credentialRefOf(entry) {
    try {
        const ref = entry?.fiber?.effectiveConfig?.apiKeyEnv?.get?.();
        if (typeof ref !== 'string' || !CREDENTIAL_REF.test(ref)) return undefined;
        return ref;
    } catch {
        return undefined;
    }
}

/**
 * 找出官方供应商当前实际配置的凭据引用；不解析值、不碰网络。
 *
 * 只有供应商实际生效的 baseURL origin 就是官方地址时才返回；自定义网关、
 * baseURL 读不出来或后备环境服务缺失都返回 undefined（fail closed）。
 *
 * @param {object} ctx 插件上下文
 * @returns {{ref: string} | undefined}
 */
export function identifyHostOcrProvider(ctx) {
    const entry = hostOcrProviderEntry(ctx);
    if (entry === undefined) return undefined;
    const endpoint = supplierEndpointOf(ctx, entry.fiber.effectiveConfig);
    if (!endpoint.ok) return undefined;
    const ref = credentialRefOf(entry);
    return ref === undefined ? undefined : { ref };
}

/**
 * 解析官方供应商保存的 Key，供一次 OCR 请求使用。
 *
 * 先校验 OCR 目标端点与供应商实际 baseURL origin，再解析 Key：
 * 自定义端点、自定义供应商网关、配置读不出来时都不调用 credentials.resolve。
 *
 * @param {object} ctx 插件上下文
 * @param {{baseURL?: string}} [options] 这次 OCR 将使用的地址（来自设置）
 * @returns {Promise<
 *   {ok: true, apiKey: string, ref: string} |
 *   {ok: false, eligible: boolean, reason: string}
 * >}
 *   eligible 只表示「OCR 端点与官方供应商行都匹配、只差 Key」；调用方可用它决定
 *   缺 Key 文案里要不要提 Models 页，但绝不能把 reason 原文回给浏览器。
 */
export async function resolveHostOcrCredential(ctx, options = {}) {
    if (officialOcrEndpointOf(options?.baseURL) === undefined) {
        return { ok: false, eligible: false, reason: 'endpoint_not_official' };
    }

    const entry = hostOcrProviderEntry(ctx);
    if (entry === undefined) return { ok: false, eligible: false, reason: 'provider_absent' };

    const endpoint = supplierEndpointOf(ctx, entry.fiber.effectiveConfig);
    if (!endpoint.ok) return { ok: false, eligible: false, reason: endpoint.reason };

    const ref = credentialRefOf(entry);
    if (ref === undefined) return { ok: false, eligible: false, reason: 'provider_absent' };

    let credentials;
    try {
        credentials = ctx.get('credentials');
    } catch {
        // 连服务都取不到只按「取不到凭据」处理，不把异常文字带出去。
        return { ok: false, eligible: true, reason: 'credential_unavailable' };
    }
    if (credentials === undefined || typeof credentials.resolve !== 'function') {
        let environment;
        try {
            environment = ctx.get('launchEnvironment');
        } catch {
            return { ok: false, eligible: true, reason: 'credential_unavailable' };
        }
        if (environment === undefined || typeof environment.get !== 'function') {
            return { ok: false, eligible: true, reason: 'credential_missing' };
        }
        let hit;
        try {
            hit = environment.get(ref);
        } catch {
            return { ok: false, eligible: true, reason: 'credential_unavailable' };
        }
        if (hit === undefined || typeof hit.value !== 'string' || hit.value.trim() === '') {
            return { ok: false, eligible: true, reason: 'credential_missing' };
        }
        return { ok: true, apiKey: hit.value, ref };
    }

    try {
        const hit = await credentials.resolve(ref);
        if (hit === undefined || typeof hit.value !== 'string' || hit.value.trim() === '') {
            return { ok: false, eligible: true, reason: 'credential_missing' };
        }
        return { ok: true, apiKey: hit.value, ref };
    } catch {
        // 凭据库读失败不能把异常原文（可能含路径 / 引用）带回 HTTP 层。
        return { ok: false, eligible: true, reason: 'credential_unavailable' };
    }
}
