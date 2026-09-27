/**
 * 多路联网检索通道 —— 每个通道一个适配器，一个统一的返回形状。
 *
 * ## 为什么要有这一层
 *
 * 第十六轮的内置通道只有一条路：Anthropic 兼容 Messages + 原生 `web_search`
 * （默认指 DeepSeek 官方端点）。那条路能用，但把它当成唯一选择有三个代价：
 *
 *   1. **单点**：端点、Key、模型名任一没配好，办公会话就查不了资料 ——
 *      而「查资料」是写文档的前置步骤，不是可选项。
 *   2. **覆盖偏差**：原生 web_search 的中文长尾覆盖一般；只做中文的场景
 *      （博查）、要语义召回的（Exa）、只要 Google 结果的（Serper / Brave）
 *      各有各的强项，锁死一条路等于把选择权拿掉。
 *   3. **零配置不可用**：一条需要 Key 的路，在没配 Key 的机器上就是死的。
 *      免 Key 的通路（DuckDuckGo、自建 SearXNG）能让「刚装好就能查」成立。
 *
 * 所以第二十轮把检索拆成「通道」：同一份查询可以交给任意一条通路，返回统一的
 * `{sources, content}`；`web.js` 负责编排（用哪条、按什么顺序、失败怎么退），
 * 本模块只管「怎么问、怎么解析」。设置页里能选，`office_search_run` 里也能按次选。
 *
 * ## 通道清单（截至第二十轮）
 *
 * | id | 通道 | 要不要 Key | 备注 |
 * | --- | --- | --- | --- |
 * | seam | 宿主 web 服务（ctx.web） | 不要 | 首选：宿主自带地址钉死、体积与超时上限 |
 * | anthropic | Anthropic 兼容 + 原生 web_search | 要 | 默认指 DeepSeek 官方；换 baseURL 即可指别的兼容端点 |
 * | openai | OpenAI 兼容 chat/completions | 要 | 用 `web_search_options`；不支持的端点会明确报「没触发检索」 |
 * | tavily | Tavily Search API | 要 | 专为 Agent 设计的检索 API，返回摘要字段 |
 * | brave | Brave Search API | 要 | 独立索引，隐私口径好 |
 * | bocha | 博查 BochaAI | 要 | 中文覆盖好，国内网络友好 |
 * | exa | Exa | 要 | 语义检索，返回正文片段 |
 * | serper | Serper（Google 结果代理） | 要 | 要「就是 Google 那一页」时用 |
 * | searxng | 自建 SearXNG | 不要（要端点） | 自己的实例，无外键依赖 |
 * | duckduckgo | DuckDuckGo HTML | 不要 | 兜底：无需任何配置，抓 HTML 解析 |
 *
 * ## 解析失败也算失败
 *
 * 每个适配器都区分三种失败：**没配好**（缺 Key / 缺端点 → config）、
 * **连不上或状态码不对**（network / blocked）、**连上了但没解析出来源**
 * （empty）。第三种最常见也最容易被静默吞掉 —— 一个返回 200 但字段名变了
 * 的接口，如果不报 `empty`，模型会以为「这个词确实没有结果」。
 *
 * @module dsh-office-mode/web-providers
 */

import { readFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';

import { asText, messageOf, serviceOf, WebAccessError, withTimeout } from './web-errors.js';

// ── 通道元数据 ──────────────────────────────────────────────────────────────

/** 通道 id（顺序即设置页里的展示顺序）。 */
export const PROVIDER_IDS = Object.freeze([
    'seam',
    'anthropic',
    'openai',
    'tavily',
    'brave',
    'bocha',
    'exa',
    'serper',
    'searxng',
    'duckduckgo',
]);

/**
 * 通道注册表。每个条目是「这条通道怎么用」的全部可读信息：
 * 展示名、要不要 Key、Key 的环境变量名、默认端点、走哪个适配器家族。
 *
 * `note` 会进设置页与 office_help —— 用户要能一眼看出「这条路要准备什么」。
 */
export const PROVIDERS = Object.freeze({
    seam: {
        id: 'seam', label: '宿主 web 服务', family: 'seam', needsKey: false,
        note: '优先用宿主的 ctx.web：它自带公网地址校验、地址钉死、体积与超时上限。',
    },
    anthropic: {
        id: 'anthropic', label: 'Anthropic 兼容 + 原生 web_search', family: 'anthropic',
        needsKey: true, keyEnv: 'DEEPSEEK_API_KEY', endpoint: 'https://api.deepseek.com/anthropic/v1',
        note: '默认指 DeepSeek 官方端点（要支持原生 web_search 的模型）；填别的兼容端点也能用。',
    },
    openai: {
        id: 'openai', label: 'OpenAI 兼容（chat/completions）', family: 'openai',
        needsKey: true, keyEnv: 'OPENAI_API_KEY', endpoint: 'https://api.openai.com/v1',
        note: '用 web_search_options 触发联网；端点不支持时会明确报「没触发检索」，不会静默返回空。',
    },
    tavily: {
        id: 'tavily', label: 'Tavily Search API', family: 'tavily',
        needsKey: true, keyEnv: 'TAVILY_API_KEY', endpoint: 'https://api.tavily.com',
        note: '专给 Agent 用的检索 API，返回结果自带摘要片段，不需要再取正文。',
    },
    brave: {
        id: 'brave', label: 'Brave Search API', family: 'brave',
        needsKey: true, keyEnv: 'BRAVE_API_KEY', endpoint: 'https://api.search.brave.com/res/v1',
        note: '独立索引；免费档有频率限制，被限流会报「目标站拒绝」。',
    },
    bocha: {
        id: 'bocha', label: '博查 BochaAI（中文友好）', family: 'bocha',
        needsKey: true, keyEnv: 'BOCHA_API_KEY', endpoint: 'https://api.bochaai.com/v1',
        note: '中文长尾覆盖好、国内直连，中文主题的检索建议优先选它。',
    },
    exa: {
        id: 'exa', label: 'Exa（语义检索）', family: 'exa',
        needsKey: true, keyEnv: 'EXA_API_KEY', endpoint: 'https://api.exa.ai',
        note: '语义检索，返回正文片段；适合「找观点/找论文」而不是「找官网」。',
    },
    serper: {
        id: 'serper', label: 'Serper（Google 结果代理）', family: 'serper',
        needsKey: true, keyEnv: 'SERPER_API_KEY', endpoint: 'https://google.serper.dev',
        note: '拿到的就是 Google 那一页的有机结果；要「按 Google 的口径」时选它。',
    },
    searxng: {
        id: 'searxng', label: '自建 SearXNG', family: 'searxng',
        needsKey: false, needsBase: true, endpoint: 'http://127.0.0.1:8080',
        note: '自己的实例，不用外键；要填实例地址（含协议与端口），实例必须允许 JSON 输出。',
    },
    duckduckgo: {
        id: 'duckduckgo', label: 'DuckDuckGo（免 Key）', family: 'duckduckgo',
        needsKey: false, endpoint: 'https://html.duckduckgo.com/html/',
        note: '不需要任何配置的兜底通道：抓 HTML 结果页再解析。可能被人机校验拦住，那时换别的通道。',
    },
});

/** 一条通道的元数据；未知 id 返回 undefined。 */
export function providerOf(id) {
    return PROVIDERS[id];
}

/**
 * 解析「一次检索该用哪些通道」。
 *
 *   provider 给具体 id（含 'seam'）→ 只用那一条；
 *   provider = 'auto'（默认）→ 按 providerOrder 依次尝试，第一条成功的就用它。
 *
 * 顺序里认不出来的 id 直接丢掉：配置里写错一个通道名，不该让整条检索失败。
 */
export function resolveProviderOrder(provider, order, fallbackOrder = PROVIDER_IDS) {
    const single = asText(provider);
    if (single !== '' && single !== 'auto') {
        if (!PROVIDER_IDS.includes(single)) {
            throw new WebAccessError(
                `不认识的检索通道「${single}」。可用的是：${PROVIDER_IDS.join(' / ')}。`,
                'OFFICE_WEB_PROVIDER',
            );
        }
        return [single];
    }
    const list = Array.isArray(order) ? order.map((item) => asText(item)).filter((item) => PROVIDER_IDS.includes(item)) : [];
    return list.length > 0 ? list : [...fallbackOrder];
}

/** 每个通道的参数默认值（数值项在这里夹取，坏值退回默认）。 */
export function providerDefaults(id) {
    const meta = PROVIDERS[id];
    if (meta === undefined) return {};
    return {
        apiKey: '',
        apiKeyEnv: meta.keyEnv ?? '',
        baseURL: meta.endpoint ?? '',
        model: id === 'anthropic' ? 'deepseek-v4-flash' : (id === 'openai' ? 'gpt-4o-mini' : ''),
        apiVersion: '2023-06-01',
        maxUses: 5,
        maxTokens: 2048,
        maxResults: id === 'seam' ? 8 : 8,
        // 与 config.js 的 DEFAULT_SEARCH.providers 对齐：兼容端点（要模型跑一轮检索）
        // 给 60 秒，纯检索 API 给 30 秒。
        timeoutMs: id === 'anthropic' || id === 'openai' ? 60_000 : 30_000,
        userAgent: 'dsh-office-mode/0.1 (+built-in web access)',
    };
}

/**
 * 把一份通道配置收敛成可用的参数。
 *
 * `raw` 可以是整个 `search` 段（新形状），也可以直接是某条通道的参数（老形状，
 * 例如原来的 `search.builtin`）—— 第十六～十九轮的配置与测试都按后者写，
 * 这条兼容不能断。
 */
export function resolveProviderOptions(id, raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const defaults = providerDefaults(id);
    const pick = (key, fallback) => (asText(source[key]) !== '' ? asText(source[key]) : fallback);
    const num = (key, fallback, min, max) => {
        const value = typeof source[key] === 'number' ? source[key] : Number.parseFloat(String(source[key]));
        if (!Number.isFinite(value)) return fallback;
        return Math.min(max, Math.max(min, value));
    };
    return {
        id,
        label: PROVIDERS[id]?.label ?? id,
        apiKey: asText(source.apiKey),
        apiKeyEnv: pick('apiKeyEnv', defaults.apiKeyEnv ?? ''),
        baseURL: pick('baseURL', defaults.baseURL ?? '').replace(/\/+$/, ''),
        model: pick('model', defaults.model ?? ''),
        apiVersion: pick('apiVersion', defaults.apiVersion ?? '2023-06-01'),
        maxUses: num('maxUses', defaults.maxUses ?? 5, 1, 20),
        maxTokens: num('maxTokens', defaults.maxTokens ?? 2048, 256, 16_384),
        maxResults: num('maxResults', defaults.maxResults ?? 8, 1, 20),
        timeoutMs: num('timeoutMs', defaults.timeoutMs ?? 60_000, 5_000, 300_000),
        userAgent: pick('userAgent', defaults.userAgent ?? 'dsh-office-mode/0.1'),
    };
}

// ── 凭据 ────────────────────────────────────────────────────────────────────

/**
 * 从 $DSH_HOME/.credentials.yaml 里按名字取一个 Key（凭据服务不在时的兜底）。
 *
 * 只做「NAME: value」这一种形状的行解析，不引 YAML 依赖：解析不到就返回空，
 * 宁可不工作也不猜。文件不存在是正常情况（有人只用环境变量）。
 */
async function readCredentialFile(name, home) {
    if (name === '') return '';
    const candidates = [];
    const dshHome = asText(process.env.DSH_HOME);
    if (dshHome !== '') candidates.push(join(dshHome, '.credentials.yaml'));
    const userHome = asText(home) || homedir();
    if (userHome !== '') candidates.push(join(userHome, '.dsh', '.credentials.yaml'));
    const pattern = new RegExp('^[ \\t]*' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '[ \\t]*:[ \\t]*(.+?)[ \\t]*$');
    for (const candidate of candidates) {
        let raw;
        try {
            raw = await readFile(candidate, 'utf8');
        } catch {
            continue;
        }
        for (const line of raw.split(/\r?\n/)) {
            const matched = pattern.exec(line);
            if (matched === null) continue;
            const value = matched[1].replace(/^["']|["']$/g, '').trim();
            if (value !== '') return value;
        }
    }
    return '';
}

/**
 * 解析一次检索要用的 API Key。
 *
 * 顺序：调用参数里的字面量 → 宿主的凭据服务 → 进程环境 → $DSH_HOME/.credentials.yaml。
 * 「字面量优先」是给测试与显式配置留的后门，正常部署走中间两条。
 */
export async function resolveApiKey(ctx, options) {
    if (asText(options?.apiKey) !== '') return asText(options.apiKey);
    const env = asText(options?.apiKeyEnv);
    if (env === '') return '';
    const credentials = serviceOf(ctx, 'credentials');
    if (credentials !== undefined && typeof credentials.resolve === 'function') {
        try {
            const resolved = await credentials.resolve(env);
            const value = asText(resolved?.value);
            if (value !== '') return value;
        } catch {
            // 凭据服务在但取不到：继续往下试环境与凭据文件。
        }
    }
    const ambient = asText(process.env[env]);
    if (ambient !== '') return ambient;
    return readCredentialFile(env);
}

// ── HTTP 小工具 ─────────────────────────────────────────────────────────────

/** 状态码 → 错误码。401/403 是「Key 不对或没权限」，429/451 是被挡。 */
function codeForStatus(status) {
    if (status === 401 || status === 403) return 'OFFICE_WEB_NO_KEY';
    if (status === 429 || status === 451) return 'OFFICE_WEB_BLOCKED';
    if (status >= 500) return 'OFFICE_WEB_NETWORK';
    return 'OFFICE_WEB_PROVIDER';
}

/**
 * 发一次请求并解析 JSON。
 *
 * 每个通道都要走这一段：连不上、超时、状态码不对、返回的不是 JSON —— 四类
 * 失败各有各的错误码，反馈层才分得清「该去配 Key」还是「该换个通道重试」。
 */
async function requestJson(input) {
    const { url, method = 'GET', headers = {}, body, fetchImpl, signal, timeoutMs, what } = input;
    const requestSignal = withTimeout(signal, timeoutMs);
    let response;
    try {
        response = await fetchImpl(url, {
            method,
            redirect: 'error',
            headers: { accept: 'application/json', 'user-agent': input.userAgent ?? 'dsh-office-mode/0.1', ...headers },
            ...(body === undefined ? {} : { body: typeof body === 'string' ? body : JSON.stringify(body) }),
            ...(requestSignal === undefined ? {} : { signal: requestSignal }),
        });
    } catch (error) {
        const timedOut = signal?.aborted !== true && requestSignal?.aborted === true;
        throw new WebAccessError(
            timedOut ? `${what}超过 ${timeoutMs} 毫秒。` : `${what}连不上：${messageOf(error)}`,
            timedOut ? 'OFFICE_WEB_TIMEOUT' : 'OFFICE_WEB_NETWORK',
            { cause: error },
        );
    }
    const status = Number.isFinite(response?.status) ? response.status : 0;
    if (response?.ok !== true) {
        let detail = '';
        try {
            detail = messageOf(await response.text());
        } catch {
            detail = '';
        }
        throw new WebAccessError(
            `${what}返回 HTTP ${status}${detail === '' ? '' : '：' + detail}`,
            codeForStatus(status),
        );
    }
    try {
        return await response.json();
    } catch (error) {
        throw new WebAccessError(`${what}返回的不是 JSON：${messageOf(error)}`, 'OFFICE_WEB_PROVIDER', { cause: error });
    }
}

/** 从 HTML 里抓文本用的小工具：去标签 + 解实体 + 压空白。 */
function stripTags(text) {
    return String(text ?? '')
        .replace(/<[^>]*>/g, ' ')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&quot;/gi, '"')
        .replace(/&#(\d+);/g, (whole, code) => {
            const value = Number.parseInt(code, 10);
            return Number.isFinite(value) ? String.fromCodePoint(value) : whole;
        })
        .replace(/\s+/g, ' ')
        .trim();
}

/** 统一收敛来源数组：只要 http(s) 的绝对地址，按 URL 去重。 */
export function normalizeSources(items, maxResults) {
    const sources = [];
    const seen = new Set();
    for (const item of items ?? []) {
        const url = asText(item?.url);
        if (url === '' || seen.has(url)) continue;
        if (!/^https?:\/\//i.test(url)) continue;
        seen.add(url);
        const title = asText(item?.title);
        const snippet = asText(item?.snippet).replace(/\s+/g, ' ');
        sources.push({
            url,
            ...(title === '' ? {} : { title }),
            ...(snippet === '' ? {} : { snippet: snippet.length > 400 ? snippet.slice(0, 400) + '…' : snippet }),
            ...(asText(item?.publishedAt) === '' ? {} : { publishedAt: asText(item.publishedAt) }),
        });
        if (Number.isFinite(maxResults) && sources.length >= maxResults) break;
    }
    return sources;
}

// ── 各家族的响应映射 ────────────────────────────────────────────────────────

/**
 * Anthropic 兼容响应 → 统一检索结果（第十六轮那段映射，原样搬过来）。
 *
 * 引用片段来自 text 块里的 citations（DeepSeek 的 web_search_result 条目本身
 * 通常不带 snippet），按 URL 取第一次出现的那条。没有任何 web_search_tool_result
 * 块时报错而不是去猜正文 —— 那说明这次请求根本没触发原生检索。
 */
export function mapSearchResponse(body) {
    const blocks = Array.isArray(body?.content) ? body.content : [];
    const resultBlocks = blocks.filter((block) => block?.type === 'web_search_tool_result');
    if (resultBlocks.length === 0) {
        throw new WebAccessError('检索接口没有返回 web_search_tool_result 块，这次请求没有触发原生检索。', 'OFFICE_WEB_NO_RESULTS');
    }
    const snippets = new Map();
    for (const block of blocks) {
        if (block?.type !== 'text') continue;
        for (const cite of block.citations ?? []) {
            const url = asText(cite?.url);
            const cited = asText(cite?.cited_text);
            if (url === '' || cited === '' || snippets.has(url)) continue;
            snippets.set(url, cited);
        }
    }
    const raw = [];
    for (const block of resultBlocks) {
        for (const item of block?.content ?? []) {
            if (item?.type !== 'web_search_result') continue;
            const url = asText(item.url);
            if (url === '') continue;
            raw.push({ url, title: asText(item.title), snippet: asText(snippets.get(url)), publishedAt: asText(item.page_age) });
        }
    }
    const summary = blocks.filter((block) => block?.type === 'text').map((block) => asText(block.text)).filter((line) => line !== '').join('\n');
    return { content: summary, sources: normalizeSources(raw, undefined) };
}

/**
 * OpenAI 兼容响应 → 统一检索结果。
 *
 * 引用有两种形状：`annotations` 里的 `url_citation`（新版），或正文里的
 * Markdown 链接（老版 / 兼容端点）。两种都收，一个都没有就报「没触发检索」——
 * 沉默地返回 0 条会让模型以为「这个词没有结果」。
 */
export function mapOpenAiResponse(body) {
    const message = body?.choices?.[0]?.message ?? {};
    const content = asText(message.content);
    const raw = [];
    for (const annotation of Array.isArray(message.annotations) ? message.annotations : []) {
        const cite = annotation?.url_citation ?? annotation;
        const url = asText(cite?.url);
        if (url === '') continue;
        raw.push({ url, title: asText(cite?.title), snippet: asText(cite?.snippet ?? cite?.text) });
    }
    if (raw.length === 0) {
        for (const matched of content.matchAll(/\[([^\]]{0,200})\]\((https?:\/\/[^\s)]{1,600})\)/g)) {
            raw.push({ url: matched[2], title: matched[1] });
        }
    }
    if (raw.length === 0 && content === '') {
        throw new WebAccessError('OpenAI 兼容端点没有返回内容，这次请求没有触发检索。', 'OFFICE_WEB_NO_RESULTS');
    }
    return { content, sources: normalizeSources(raw, undefined) };
}

/** Tavily `/search` → 统一检索结果。 */
export function mapTavilyResponse(body) {
    const results = Array.isArray(body?.results) ? body.results : [];
    return {
        content: asText(body?.answer),
        sources: normalizeSources(results.map((item) => ({ url: item?.url, title: item?.title, snippet: item?.content })), undefined),
    };
}

/** Brave `/web/search` → 统一检索结果。 */
export function mapBraveResponse(body) {
    const results = Array.isArray(body?.web?.results) ? body.web.results : [];
    return {
        content: '',
        sources: normalizeSources(results.map((item) => ({
            url: item?.url,
            title: item?.title,
            snippet: item?.description,
            publishedAt: item?.age,
        })), undefined),
    };
}

/** 博查 `/web-search` → 统一检索结果（两种外层形状都认）。 */
export function mapBochaResponse(body) {
    const pages = body?.data?.webPages?.value ?? body?.webPages?.value ?? body?.data?.value ?? [];
    const results = Array.isArray(pages) ? pages : [];
    return {
        content: asText(body?.data?.summary ?? body?.summary),
        sources: normalizeSources(results.map((item) => ({
            url: item?.url ?? item?.siteUrl,
            title: item?.name ?? item?.title,
            snippet: item?.snippet ?? item?.summary ?? item?.description,
            publishedAt: item?.datePublished ?? item?.dateLastCrawled,
        })), undefined),
    };
}

/** Exa `/search` → 统一检索结果。 */
export function mapExaResponse(body) {
    const results = Array.isArray(body?.results) ? body.results : [];
    return {
        content: '',
        sources: normalizeSources(results.map((item) => ({
            url: item?.url,
            title: item?.title,
            snippet: item?.text ?? item?.summary ?? item?.highlights?.[0],
            publishedAt: item?.publishedDate,
        })), undefined),
    };
}

/** Serper `/search` → 统一检索结果。 */
export function mapSerperResponse(body) {
    const organic = Array.isArray(body?.organic) ? body.organic : [];
    return {
        content: asText(body?.answerBox?.answer ?? body?.answerBox?.snippet),
        sources: normalizeSources(organic.map((item) => ({
            url: item?.link ?? item?.url,
            title: item?.title,
            snippet: item?.snippet,
            publishedAt: item?.date,
        })), undefined),
    };
}

/** SearXNG `/search?format=json` → 统一检索结果。 */
export function mapSearxngResponse(body) {
    const results = Array.isArray(body?.results) ? body.results : [];
    return {
        content: '',
        sources: normalizeSources(results.map((item) => ({
            url: item?.url,
            title: item?.title,
            snippet: item?.content,
            publishedAt: item?.publishedDate,
        })), undefined),
    };
}

/**
 * DuckDuckGo HTML 结果页 → 统一检索结果。
 *
 * 只需要两类节点：结果链接（`class="result__a"`）与摘要（`class="result__snippet"`）。
 * DuckDuckGo 的链接是 `//duckduckgo.com/l/?uddg=<encoded>` 这种跳转形式，
 * 这里把 `uddg` 参数解回来；解不出就原样留着（比丢掉好，调用方取正文时
 * 还会再跟一次同源跳转）。
 */
export function parseDuckDuckGo(html) {
    const text = String(html ?? '');
    const items = [];
    const linkRe = /<a\b[^>]*class="[^"]*result__a[^"]*"[^>]*href="([^"]{1,1000})"[^>]*>([\s\S]{0,600}?)<\/a>/gi;
    for (const matched of text.matchAll(linkRe)) {
        items.push({ url: unwrapDuckUrl(matched[1]), title: stripTags(matched[2]) });
    }
    if (items.length === 0) {
        // 结果页改版时的退化读法：所有 result__url / 普通外链。
        for (const matched of text.matchAll(/<a\b[^>]*class="[^"]*result__url[^"]*"[^>]*href="([^"]{1,1000})"[^>]*>([\s\S]{0,400}?)<\/a>/gi)) {
            items.push({ url: unwrapDuckUrl(matched[1]), title: stripTags(matched[2]) });
        }
    }
    const snippets = [...text.matchAll(/<a\b[^>]*class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]{0,1200}?)<\/a>/gi)]
        .map((matched) => stripTags(matched[1]));
    items.forEach((item, index) => {
        if (snippets[index] !== undefined && snippets[index] !== '') item.snippet = snippets[index];
    });
    return normalizeSources(items, undefined);
}

/** 解开 `//duckduckgo.com/l/?uddg=…` 这类跳转；解不出就补上协议返回原串。 */
function unwrapDuckUrl(href) {
    const raw = asText(href);
    if (raw === '') return '';
    try {
        const url = new URL(raw, 'https://duckduckgo.com');
        const target = url.searchParams.get('uddg');
        if (target !== null && target !== '') return decodeURIComponent(target);
        return url.toString();
    } catch {
        return raw.startsWith('//') ? 'https:' + raw : raw;
    }
}

// ── 通道适配器 ──────────────────────────────────────────────────────────────

/**
 * 用一条通道检索一次。
 *
 * @param {string} id 通道 id
 * @param {object} request
 * @param {string} request.query 查询
 * @param {number} [request.maxResults] 最多几条来源
 * @param {object} request.options `resolveProviderOptions(id, …)` 的结果
 * @param {Function} request.fetchImpl 注入的 fetch（便于单测）
 * @param {AbortSignal} [request.signal]
 * @param {string} [request.apiKey] 已解析的 Key（seam 用不到）
 * @returns {Promise<{sources: Array<object>, content: string}>}
 */
export async function searchProvider(id, request) {
    const meta = PROVIDERS[id];
    if (meta === undefined) throw new WebAccessError(`不认识的检索通道「${id}」。`, 'OFFICE_WEB_PROVIDER');
    const options = request.options ?? providerDefaults(id);
    const maxResults = Number.isFinite(request.maxResults) ? request.maxResults : options.maxResults;
    const base = asText(options.baseURL).replace(/\/+$/, '');
    const common = {
        fetchImpl: request.fetchImpl,
        signal: request.signal,
        timeoutMs: options.timeoutMs,
        userAgent: options.userAgent,
    };

    if (meta.needsBase === true && base === '') {
        throw new WebAccessError(`「${meta.label}」需要实例地址（baseURL），现在没配。`, 'OFFICE_WEB_NO_BASE');
    }
    if (meta.needsKey === true && asText(request.apiKey) === '') {
        throw new WebAccessError(
            `「${meta.label}」需要一个 API Key：没有 ${options.apiKeyEnv || '（未指定环境变量名）'}`
            + '（凭据服务、环境变量、$DSH_HOME/.credentials.yaml 都没有）。',
            'OFFICE_WEB_NO_KEY',
        );
    }

    switch (meta.family) {
        case 'anthropic': {
            const body = {
                model: options.model,
                max_tokens: options.maxTokens,
                messages: [{ role: 'user', content: [{ type: 'text', text: `Perform a web search for the query: ${request.query}` }] }],
                tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: options.maxUses }],
            };
            const parsed = await requestJson({
                ...common,
                url: base + '/messages',
                method: 'POST',
                what: '检索接口',
                headers: {
                    'x-api-key': request.apiKey,
                    authorization: `Bearer ${request.apiKey}`,
                    'anthropic-version': options.apiVersion,
                    'content-type': 'application/json',
                },
                body,
            });
            const mapped = mapSearchResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'openai': {
            const body = {
                model: options.model,
                max_tokens: options.maxTokens,
                messages: [{ role: 'user', content: `Perform a web search for the query: ${request.query}` }],
                web_search_options: {},
            };
            const parsed = await requestJson({
                ...common,
                url: base + '/chat/completions',
                method: 'POST',
                what: '检索接口',
                headers: { authorization: `Bearer ${request.apiKey}`, 'content-type': 'application/json' },
                body,
            });
            const mapped = mapOpenAiResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'tavily': {
            const parsed = await requestJson({
                ...common,
                url: base + '/search',
                method: 'POST',
                what: 'Tavily 检索接口',
                headers: { 'content-type': 'application/json' },
                body: {
                    api_key: request.apiKey,
                    query: request.query,
                    max_results: maxResults,
                    search_depth: 'basic',
                    include_answer: false,
                },
            });
            const mapped = mapTavilyResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'brave': {
            const url = base + '/web/search?q=' + encodeURIComponent(request.query) + '&count=' + String(Math.min(20, Math.max(1, maxResults ?? 8)));
            const parsed = await requestJson({
                ...common,
                url,
                what: 'Brave 检索接口',
                headers: { 'x-subscription-token': request.apiKey, accept: 'application/json' },
            });
            const mapped = mapBraveResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'bocha': {
            const parsed = await requestJson({
                ...common,
                url: base + '/web-search',
                method: 'POST',
                what: '博查检索接口',
                headers: { authorization: `Bearer ${request.apiKey}`, 'content-type': 'application/json' },
                body: { query: request.query, count: Math.min(20, Math.max(1, maxResults ?? 8)), summary: false },
            });
            const mapped = mapBochaResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'exa': {
            const parsed = await requestJson({
                ...common,
                url: base + '/search',
                method: 'POST',
                what: 'Exa 检索接口',
                headers: { 'x-api-key': request.apiKey, 'content-type': 'application/json' },
                body: {
                    query: request.query,
                    numResults: Math.min(20, Math.max(1, maxResults ?? 8)),
                    contents: { text: { maxCharacters: 1_200 } },
                },
            });
            const mapped = mapExaResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'serper': {
            const parsed = await requestJson({
                ...common,
                url: base + '/search',
                method: 'POST',
                what: 'Serper 检索接口',
                headers: { 'x-api-key': request.apiKey, 'content-type': 'application/json' },
                body: { q: request.query, num: Math.min(20, Math.max(1, maxResults ?? 8)) },
            });
            const mapped = mapSerperResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'searxng': {
            const url = base + '/search?q=' + encodeURIComponent(request.query) + '&format=json';
            const parsed = await requestJson({ ...common, url, what: 'SearXNG 检索接口' });
            const mapped = mapSearxngResponse(parsed);
            return { ...mapped, sources: normalizeSources(mapped.sources, maxResults) };
        }
        case 'duckduckgo': {
            const url = (base === '' ? 'https://html.duckduckgo.com/html/' : base + '/') + '?q=' + encodeURIComponent(request.query);
            const requestSignal = withTimeout(request.signal, options.timeoutMs);
            let response;
            try {
                response = await request.fetchImpl(url, {
                    method: 'GET',
                    redirect: 'error',
                    headers: {
                        accept: 'text/html',
                        'user-agent': options.userAgent,
                    },
                    ...(requestSignal === undefined ? {} : { signal: requestSignal }),
                });
            } catch (error) {
                throw new WebAccessError(`DuckDuckGo 连不上：${messageOf(error)}`, 'OFFICE_WEB_NETWORK', { cause: error });
            }
            if (response?.ok !== true) {
                throw new WebAccessError(`DuckDuckGo 返回 HTTP ${response?.status ?? 0}`, codeForStatus(Number(response?.status) || 0));
            }
            const html = await response.text().catch(() => '');
            const sources = parseDuckDuckGo(html);
            if (sources.length === 0) {
                throw new WebAccessError(
                    'DuckDuckGo 没解析出任何来源（多半是人机校验页）。换一条通道，例如配一个 Key 用 Tavily / 博查。',
                    'OFFICE_WEB_NO_RESULTS',
                );
            }
            return { content: '', sources: normalizeSources(sources, maxResults) };
        }
        default:
            throw new WebAccessError(`通道「${id}」还没有适配器。`, 'OFFICE_WEB_PROVIDER');
    }
}

/**
 * 一条通道现在能不能用（不联网）。
 *
 * `seam` 自己由 web.js 判（它要先问 ctx.web 在不在），这里只处理需要
 * Key / 端点的通路：缺什么就照实说缺什么 —— 设置页要靠这句话告诉用户
 * 「还差哪一步」。
 */
export async function probeProvider(id, options, ctx, seamAvailable = false) {
    const meta = PROVIDERS[id];
    if (meta === undefined) return { id, ok: false, reason: `不认识的检索通道「${id}」` };
    if (meta.family === 'seam') {
        return seamAvailable
            ? { id, ok: true, reason: '宿主 web 服务可用' }
            : { id, ok: false, reason: '组合里没有 web 服务（ctx.web）' };
    }
    if (meta.needsBase === true && asText(options?.baseURL) === '') {
        return { id, ok: false, reason: `「${meta.label}」还没填实例地址（baseURL）` };
    }
    if (meta.needsKey === true) {
        const apiKey = await resolveApiKey(ctx, options);
        if (apiKey === '') {
            return { id, ok: false, reason: `「${meta.label}」没配 ${options?.apiKeyEnv || 'API Key'}` };
        }
    }
    const where = meta.family === 'seam' ? '' : (asText(options?.baseURL) === '' ? '' : ` @ ${options.baseURL}`);
    return { id, ok: true, reason: `「${meta.label}」可用${where}` };
}
