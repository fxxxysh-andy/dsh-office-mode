/**
 * 知识库向量化（第八十五轮；第一百零三轮移除重排，候选筛选独立到 kb-selector.js）。
 *
 * 设计目标是**永远有结果**：远端 embedding 超时或报错时退回词法并标 degraded；
 * 只有显式 mode:'local' 且本地推理没开时才用确定性的哈希向量。auto 不会在
 * 没有端点时偷偷换成本地哈希（它与词法高度重合，会把「弃答」变松）。
 *
 * 本地哈希不是语义模型：它只把中文 bigram / 英文词哈希到固定维度，能补一点
 * 「词面不同但字面相近」的召回，不能替代真正的 embedding。它的价值在于零依赖、
 * 可复现、离线可用 —— 用于自测与显式 local 的兜底。
 *
 * @module dsh-office-mode/kb-embed
 */
import { asText } from './web-errors.js';
import { kbTokens } from './kb.js';

export const LOCAL_EMBEDDING_MODEL = 'local-hash-v1';
export const DEFAULT_LOCAL_DIMENSIONS = 256;

/** 把配置收口成一份运行期参数（坏值落默认，绝不抛）。 */
export function normalizeEmbeddingConfig(raw = {}) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const mode = ['auto', 'remote', 'local', 'off'].includes(source.mode) ? source.mode : 'auto';
    return {
        mode,
        baseURL: asText(source.baseURL).replace(/\/+$/, ''),
        model: asText(source.model),
        apiKey: asText(source.apiKey),
        apiKeyEnv: asText(source.apiKeyEnv),
        dimensions: Number.isFinite(source.dimensions) ? Math.min(4096, Math.max(0, Math.trunc(source.dimensions))) : 0,
        batchSize: Number.isFinite(source.batchSize) ? Math.min(64, Math.max(1, Math.trunc(source.batchSize))) : 16,
        timeoutMs: Number.isFinite(source.timeoutMs) ? Math.min(120_000, Math.max(1_000, Math.trunc(source.timeoutMs))) : 15_000,
    };
}

/** 远端配置是否齐（不要求 Key：本地/自建端点常常不要 Key）。 */
export function canUseRemoteEmbedding(config) {
    const cfg = normalizeEmbeddingConfig(config);
    if (cfg.mode === 'local' || cfg.mode === 'off') return false;
    return cfg.baseURL !== '' && cfg.model !== '';
}

/**
 * 设置页的「通道」与实际检索路径的对照（第一百零四轮）：
 *
 * - off    → off        只用词法；
 * - local  → local      本地推理（起不来由调用方退回词法，**不会**悄悄换成哈希）；
 * - local  → local-hash 本地推理没开时，显式 local 才允许零依赖哈希向量；
 * - remote → remote     端点 + 模型齐了才走远端，否则 none（退回词法）；
 * - auto   → local（开了本地推理）/ remote（配了远端）/ none（都没有，纯词法）。
 *
 * 以前这段优先级写在 hybridKbSearch 里，且**先看 localEnabled、再看 mode**，
 * 导致 mode:'off' 也会起本地进程、mode:'remote' 会被本地抢占。抽成纯函数，
 * 设置页与服务端读同一份口径。
 */
export function embeddingRouteOf(raw = {}) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const { mode = 'auto', localEnabled = false, remoteReady = false } = source;
    const wanted = normalizeEmbeddingConfig({ mode }).mode;
    if (wanted === 'off') return 'off';
    if (wanted === 'local') return localEnabled ? 'local' : 'local-hash';
    if (wanted === 'remote') return remoteReady ? 'remote' : 'none';
    if (localEnabled) return 'local';
    return remoteReady ? 'remote' : 'none';
}

/** 取 Key：字面量优先，其次环境变量。 */
export function embeddingKeyOf(config, env = process.env) {
    const cfg = normalizeEmbeddingConfig(config);
    if (cfg.apiKey !== '') return cfg.apiKey;
    if (cfg.apiKeyEnv !== '') return asText(env?.[cfg.apiKeyEnv]);
    return '';
}

/** FNV-1a 32 位：本地向量的哈希槽，跨进程稳定。 */
function fnv1a(text) {
    let hash = 0x811c9dc5;
    for (let i = 0; i < text.length; i += 1) {
        hash ^= text.charCodeAt(i);
        hash = Math.imul(hash, 0x01000193) >>> 0;
    }
    return hash >>> 0;
}

/**
 * 本地确定性向量：每个 token 按哈希落到两个槽（带符号），最后 L2 归一化。
 * 同一段文本在任何机器、任何进程都得到同一个向量。
 */
export function localEmbedding(text, dimensions = DEFAULT_LOCAL_DIMENSIONS) {
    const dim = Math.min(4096, Math.max(8, Math.trunc(Number(dimensions) || DEFAULT_LOCAL_DIMENSIONS)));
    const vector = new Float32Array(dim);
    const tokens = kbTokens(asText(text));
    for (const token of tokens) {
        const hash = fnv1a(String(token));
        const slot = hash % dim;
        const sign = (hash >>> 8) & 1 ? -1 : 1;
        vector[slot] += sign * Math.min(4, 1 + Math.log10(1 + String(token).length));
    }
    let norm = 0;
    for (let i = 0; i < dim; i += 1) norm += vector[i] * vector[i];
    if (norm > 0) {
        const scale = 1 / Math.sqrt(norm);
        for (let i = 0; i < dim; i += 1) vector[i] *= scale;
    }
    return vector;
}

/** 余弦相似度；维度不同返回 0（不同模型/不同配置的向量不混用）。 */
export function cosineSimilarity(left, right) {
    if (!left || !right || left.length !== right.length || left.length === 0) return 0;
    let dot = 0;
    let leftNorm = 0;
    let rightNorm = 0;
    for (let i = 0; i < left.length; i += 1) {
        dot += left[i] * right[i];
        leftNorm += left[i] * left[i];
        rightNorm += right[i] * right[i];
    }
    if (leftNorm === 0 || rightNorm === 0) return 0;
    return dot / Math.sqrt(leftNorm * rightNorm);
}

/** 把 number[] / Buffer / Float32Array 统一成 Float32Array（读 SQLite BLOB 时用）。 */
export function asVector(value) {
    if (value instanceof Float32Array) return value;
    if (value instanceof Uint8Array) {
        if (value.byteLength % 4 !== 0) return new Float32Array(0);
        return new Float32Array(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength));
    }
    if (Array.isArray(value)) return new Float32Array(value.map(Number));
    return new Float32Array(0);
}

async function remoteEmbeddings({ texts, config, fetchImpl }) {
    const cfg = normalizeEmbeddingConfig(config);
    const body = { model: cfg.model, input: texts };
    if (cfg.dimensions > 0) body.dimensions = cfg.dimensions;
    const headers = { 'content-type': 'application/json' };
    const key = embeddingKeyOf(cfg);
    if (key !== '') headers.authorization = `Bearer ${key}`;
    const response = await fetchImpl(`${cfg.baseURL}/embeddings`, {
        method: 'POST',
        headers,
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(cfg.timeoutMs),
    });
    const raw = await response.text();
    if (!response.ok) throw new Error(`embedding 端点返回 HTTP ${response.status}：${raw.slice(0, 200)}`);
    let parsed = null;
    try { parsed = JSON.parse(raw); } catch { throw new Error('embedding 端点返回的不是 JSON。'); }
    const data = Array.isArray(parsed?.data) ? parsed.data : [];
    if (data.length !== texts.length) throw new Error(`embedding 端点返回 ${data.length} 条，期望 ${texts.length} 条。`);
    const vectors = data.map((item) => {
        const vector = Array.isArray(item?.embedding) ? item.embedding : null;
        if (vector === null || vector.length === 0) throw new Error('embedding 端点返回了空向量。');
        return new Float32Array(vector.map(Number));
    });
    return { vectors, model: asText(parsed?.model) || cfg.model };
}

/**
 * 批量向量化。远端失败时退回本地向量（除非 mode='remote' 且调用方要求严格失败）。
 * @returns {Promise<{ok:boolean, vectors:Float32Array[], model:string, mode:'remote'|'local'|'off', degraded:boolean, error:string}>}
 */
export async function embedTexts({ texts = [], config = {}, fetchImpl = globalThis.fetch, strict = false } = {}) {
    const list = (Array.isArray(texts) ? texts : []).map((text) => String(text ?? ''));
    const cfg = normalizeEmbeddingConfig(config);
    if (cfg.mode === 'off') {
        return { ok: false, vectors: [], model: '', mode: 'off', degraded: false, error: '向量检索已在设置里关掉。' };
    }
    if (list.length === 0) {
        return { ok: true, vectors: [], model: cfg.model, mode: 'local', degraded: false, error: '' };
    }
    const local = () => list.map((text) => localEmbedding(text, cfg.dimensions || DEFAULT_LOCAL_DIMENSIONS));
    if (!canUseRemoteEmbedding(cfg)) {
        return { ok: true, vectors: local(), model: LOCAL_EMBEDDING_MODEL, mode: 'local', degraded: cfg.mode === 'remote', error: cfg.mode === 'remote' ? '没有配置远端 embedding（baseURL / model），已退回本地向量。' : '' };
    }
    try {
        const remote = await remoteEmbeddings({ texts: list, config: cfg, fetchImpl });
        return { ok: true, vectors: remote.vectors, model: remote.model, mode: 'remote', degraded: false, error: '' };
    } catch (error) {
        if (strict) {
            return { ok: false, vectors: [], model: cfg.model, mode: 'remote', degraded: true, error: `embedding 失败：${error?.message ?? error}` };
        }
        return { ok: true, vectors: local(), model: LOCAL_EMBEDDING_MODEL, mode: 'local', degraded: true, error: `embedding 远端不可用，已退回本地向量：${error?.message ?? error}` };
    }
}


