/**
 * 知识库候选筛选（第一百零三轮，替代原来的「重排」）。
 *
 * 旧的重排把候选正文塞给对话模型，再把结果塞回主会话：候选全文进过一次模型，
 * 融合结果又整批回到主上下文，两头都在烧 token。这里改成一道**独立的筛选步骤**：
 * 只把「编号 + 每块前 headChars 字」发给侧路模型，要求它**只回编号**；主上下文
 * 只拿被选中的那几块。副作用是筛选模型看不到块的后半段——所以它只做「先粗筛、
 * 保序截断」，不做排序输出。
 *
 * 失败路径一律不阻断检索：未配置 / 超时 / HTTP 错误 / 回垃圾编号，都原样返回
 * 融合结果，只在返回的 selector 字段里标 applied:false。
 *
 * @module dsh-office-mode/kb-selector
 */
import { asText } from './web-errors.js';

export const DEFAULT_SELECTOR = {
    mode: 'off',
    baseURL: '',
    model: '',
    apiKey: '',
    apiKeyEnv: '',
    timeoutMs: 12_000,
    maxCandidates: 12,
    maxSelected: 5,
    headChars: 300,
};

/** 把配置收口成一份运行期参数（坏值落默认，绝不抛）。 */
export function normalizeSelectorConfig(raw = {}) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const clamp = (value, fallback, min, max) => (Number.isFinite(value) ? Math.min(max, Math.max(min, Math.trunc(value))) : fallback);
    return {
        mode: source.mode === 'llm' ? 'llm' : 'off',
        baseURL: asText(source.baseURL).replace(/\/+$/, ''),
        model: asText(source.model),
        apiKey: asText(source.apiKey),
        apiKeyEnv: asText(source.apiKeyEnv),
        timeoutMs: clamp(source.timeoutMs, DEFAULT_SELECTOR.timeoutMs, 1_000, 120_000),
        maxCandidates: clamp(source.maxCandidates, DEFAULT_SELECTOR.maxCandidates, 2, 30),
        maxSelected: clamp(source.maxSelected, DEFAULT_SELECTOR.maxSelected, 1, 20),
        headChars: clamp(source.headChars, DEFAULT_SELECTOR.headChars, 80, 2_000),
    };
}

/** 筛选是否具备调用条件：要显式 llm、要有端点与模型。 */
export function canUseSelector(config) {
    const cfg = normalizeSelectorConfig(config);
    return cfg.mode === 'llm' && cfg.baseURL !== '' && cfg.model !== '';
}

/** 取 Key：字面量优先，其次环境变量。 */
export function selectorKeyOf(config, env = process.env) {
    const cfg = normalizeSelectorConfig(config);
    if (cfg.apiKey !== '') return cfg.apiKey;
    if (cfg.apiKeyEnv !== '') return asText(env?.[cfg.apiKeyEnv]);
    return '';
}

/** 每个候选给模型看的头：优先块头，没有就截正文。 */
function headOf(hit, chars) {
    const head = asText(hit?.head);
    if (head !== '') return head.slice(0, chars);
    return asText(hit?.text).slice(0, chars);
}

/**
 * 调一次侧路模型，返回它挑中的块 id（按模型给出的顺序）。
 *
 * 结果**不含正文**，调用方拿 id 回到融合列表里过滤——省的是主上下文，
 * 而不是筛选模型自己的输入（那部分本来就被 headChars 卡住了）。
 */
export async function selectWithLlm({ query, hits = [], config = {}, fetchImpl = globalThis.fetch, env = process.env } = {}) {
    const cfg = normalizeSelectorConfig(config);
    const list = (Array.isArray(hits) ? hits : []).slice(0, cfg.maxCandidates);
    if (!canUseSelector(cfg)) {
        return { ok: false, ids: [], error: '没有配置筛选端点或模型。' };
    }
    if (list.length === 0) return { ok: true, ids: [], error: '' };
    const lines = list.map((hit, index) => {
        const title = asText(hit?.title);
        const head = headOf(hit, cfg.headChars);
        return `[${index + 1}] ${title === '' ? '' : `${title}｜`}${head}`;
    }).join('\n');
    const prompt = `下面是从知识库检索出的资料块，请挑出能回答问题的那些。只输出编号，英文逗号分隔，从相关到不相关，最多 ${cfg.maxSelected} 个；一个都不相关就输出 0。

问题：${asText(query)}

${lines}`;
    const headers = { 'content-type': 'application/json' };
    const key = selectorKeyOf(cfg, env);
    if (key !== '') headers.authorization = `Bearer ${key}`;
    try {
        const response = await fetchImpl(`${cfg.baseURL}/chat/completions`, {
            method: 'POST',
            headers,
            body: JSON.stringify({
                model: cfg.model,
                temperature: 0,
                messages: [{ role: 'user', content: prompt }],
            }),
            signal: AbortSignal.timeout(cfg.timeoutMs),
        });
        const raw = await response.text();
        if (!response.ok) throw new Error(`筛选端点返回 HTTP ${response.status}：${raw.slice(0, 160)}`);
        const parsed = JSON.parse(raw);
        const text = asText(parsed?.choices?.[0]?.message?.content);
        const ids = [];
        for (const match of text.matchAll(/\d+/g)) {
            const index = Number.parseInt(match[0], 10) - 1;
            if (index < 0 || index >= list.length) continue;
            const id = asText(list[index].chunkId);
            if (id !== '' && !ids.includes(id)) ids.push(id);
        }
        if (ids.length === 0) throw new Error('模型没有返回可用编号。');
        return { ok: true, ids, error: '' };
    } catch (error) {
        return { ok: false, ids: [], error: `候选筛选失败：${error?.message ?? error}` };
    }
}
