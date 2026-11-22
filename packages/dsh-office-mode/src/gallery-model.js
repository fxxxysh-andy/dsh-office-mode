/**
 * 审美参考图库的「模型层」：视觉描述与多模态（图文对齐）embedding 适配器。
 *
 * 为什么单独一层：图库本体只负责存储、标签与检索融合，**不直接联网**。
 * 所有会花钱 / 会把图片发出去的动作都收在这里，并且满足三条硬约束：
 *
 *   1. **默认不上传**：通道缺省是 off；调用方必须显式传 allowUpload 才会发请求。
 *   2. **不假成功**：没有配置端点、没有 Key、探针没过、维度对不上，都返回
 *      ok:false + code + 原因；绝不写占位向量，也不拿文本向量冒充图片向量。
 *   3. **契约按官方文档**：Voyage 用 /v1/multimodalembeddings，Cohere 用 /v2/embed，
 *      vLLM / OpenAI 兼容用 /embeddings + messages 形态。普通 OpenAI /embeddings
 *      只吃字符串，不能当多模态用 —— 所以它不被当作图文对齐通道。
 *
 * 视觉描述：宿主通道（ctx.attachments + ctx.llm）与 OpenAI 兼容 chat 通道。
 * 宿主通道先看 prepared.inputModalities 是否声明支持 image：显式不含 image 时
 * **不发请求**，直接说这个模型不接受图片；字段缺失（未知）时照发，成功后回执
 * 标 modalitiesKnown:false，失败也照实回原因。
 *
 * @module dsh-office-mode/gallery-model
 */
import { asText } from './web-errors.js';
import { normalizeEmbeddingConfig } from './kb-embed.js';

/** 视觉描述提示词的版本号：模型或提示词一改，缓存就能区分（写进 aesthetic）。 */
export const GALLERY_PROMPT_VERSION = 'gallery-vision-v1';
/** 多模态 embedding 适配器版本号：契约变了就升它，向量缓存随之隔离。 */
export const IMAGE_EMBED_VERSION = 'mm-v1';

/** 1×1 PNG：只用于探针，证明端点接受图片输入并给出与文本同维的向量。 */
export const ONE_PIXEL_PNG = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
);

/** 多模态 embedding 的已知通道。默认端点与模型都来自官方文档。 */
export const IMAGE_PROVIDERS = Object.freeze({
    voyage: {
        label: 'Voyage',
        baseURL: 'https://api.voyageai.com/v1',
        model: 'voyage-multimodal-3.5',
        apiKeyEnv: 'VOYAGE_API_KEY',
        endpoint: '/multimodalembeddings',
    },
    cohere: {
        label: 'Cohere',
        baseURL: 'https://api.cohere.com',
        model: 'embed-v4.0',
        apiKeyEnv: 'COHERE_API_KEY',
        endpoint: '/v2/embed',
    },
    vllm: {
        label: 'vLLM / OpenAI 兼容',
        baseURL: '',
        model: '',
        apiKeyEnv: '',
        endpoint: '/embeddings',
    },
});

/** 视觉描述通道。off 表示一个都不配，describe 只会得到「未配置」。 */
export const VISION_CHANNELS = Object.freeze(['off', 'host', 'openai']);

function clampInt(value, min, max, fallback) {
    if (!Number.isFinite(value)) return fallback;
    return Math.min(max, Math.max(min, Math.trunc(value)));
}

function asObject(value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value) ? value : {};
}

/** 通道别名收口：off/none/disabled → none；openai 兼容 → vllm。 */
function providerOf(raw) {
    const name = asText(raw).trim().toLowerCase();
    if (name === '' || name === 'off' || name === 'none' || name === 'disabled') return 'none';
    if (name === 'voyage' || name === 'voyageai' || name === 'voyage-ai') return 'voyage';
    if (name === 'cohere') return 'cohere';
    if (name === 'vllm' || name === 'openai' || name === 'openai-compatible' || name === 'local') return 'vllm';
    return 'none';
}

/** 把图库的 imageEmbedding 配置收口成运行期参数（坏值落默认，绝不抛）。 */
export function normalizeImageEmbedding(raw = {}) {
    const source = asObject(raw);
    const provider = providerOf(source.provider ?? source.channel);
    const spec = IMAGE_PROVIDERS[provider] ?? null;
    const baseURL = (asText(source.baseURL) || spec?.baseURL || '').replace(/\/+$/, '');
    const model = asText(source.model) || spec?.model || '';
    const apiKey = asText(source.apiKey);
    const apiKeyEnv = asText(source.apiKeyEnv) || spec?.apiKeyEnv || '';
    return {
        provider,
        label: spec?.label ?? '未配置',
        baseURL,
        model,
        apiKey,
        apiKeyEnv,
        dimensions: clampInt(source.dimensions, 0, 4096, 0),
        batchSize: clampInt(source.batchSize, 1, 16, 4),
        timeoutMs: clampInt(source.timeoutMs, 5000, 180000, 30000),
    };
}

/** 图文对齐通道是否配齐（不要求 Key：自建 vLLM 常常不要 Key）。 */
export function imageEmbeddingReady(raw = {}) {
    const cfg = normalizeImageEmbedding(raw);
    if (cfg.provider === 'none') {
        return { ok: false, code: 'unconfigured', error: '没有配置图文对齐的多模态 embedding（Voyage / Cohere / vLLM 兼容）。' };
    }
    if (cfg.model === '') return { ok: false, code: 'unconfigured', error: '多模态 embedding 缺少模型名（model）。' };
    if (cfg.provider === 'vllm' && cfg.baseURL === '') {
        return { ok: false, code: 'unconfigured', error: 'vLLM / OpenAI 兼容通道需要填端点（baseURL）。' };
    }
    return { ok: true, code: '', error: '', config: cfg };
}

/** 视觉描述配置收口。channel 缺省 off —— 默认不会把任何图片发出去。 */
export function normalizeVisionConfig(raw = {}) {
    const source = asObject(raw);
    const channel = VISION_CHANNELS.includes(asText(source.channel)) ? asText(source.channel) : 'off';
    return {
        channel,
        provider: asText(source.provider),
        model: asText(source.model),
        baseURL: asText(source.baseURL).replace(/\/+$/, ''),
        apiKey: asText(source.apiKey),
        apiKeyEnv: asText(source.apiKeyEnv),
        maxTokens: clampInt(source.maxTokens, 200, 2000, 900),
        timeoutMs: clampInt(source.timeoutMs, 5000, 180000, 60000),
        language: source.language === 'en' ? 'en' : 'zh',
    };
}

/** 视觉通道是否可用（host 通道还要运行期再检查宿主服务）。 */
export function visionReady(raw = {}) {
    const cfg = normalizeVisionConfig(raw);
    if (cfg.channel === 'off') {
        return { ok: false, code: 'unconfigured', error: '没有配置视觉模型：图库只会保存图片与手工标签，不生成描述。' };
    }
    if (cfg.channel === 'host') {
        if (cfg.provider === '' || cfg.model === '') {
            return { ok: false, code: 'unconfigured', error: '宿主视觉通道要先选 provider 与 model。' };
        }
        return { ok: true, code: '', error: '', config: cfg };
    }
    if (cfg.baseURL === '' || cfg.model === '') {
        return { ok: false, code: 'unconfigured', error: 'OpenAI 兼容视觉通道需要端点与模型名。' };
    }
    return { ok: true, code: '', error: '', config: cfg };
}

/** 整份图库配置（文本向量 + 图片向量 + 视觉通道）。 */
export function normalizeGalleryConfig(raw = {}) {
    const source = asObject(raw);
    return {
        textEmbedding: normalizeEmbeddingConfig(source.textEmbedding ?? {}),
        imageEmbedding: normalizeImageEmbedding(source.imageEmbedding ?? {}),
        vision: normalizeVisionConfig(source.vision ?? {}),
    };
}

function keyOf(cfg, env) {
    const literal = asText(cfg.apiKey);
    if (literal !== '') return literal;
    const name = asText(cfg.apiKeyEnv);
    if (name === '') return '';
    return asText(env?.[name]);
}

/** 图片 → data URL（Voyage / Cohere / OpenAI 兼容都接受 base64 data URL）。 */
export function dataUrlOf(bytes, mediaType = 'image/png') {
    const buffer = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes ?? []);
    return 'data:' + mediaType + ';base64,' + buffer.toString('base64');
}

async function postJson({ fetchImpl, url, headers, body, timeoutMs, signal }) {
    const controller = new AbortController();
    let timedOut = false;
    const timer = setTimeout(() => { timedOut = true; controller.abort(); }, timeoutMs);
    const onAbort = () => controller.abort();
    if (signal !== undefined) {
        if (signal.aborted) controller.abort();
        else signal.addEventListener('abort', onAbort, { once: true });
    }
    try {
        const response = await fetchImpl(url, {
            method: 'POST',
            headers: { 'content-type': 'application/json', ...headers },
            body: JSON.stringify(body),
            signal: controller.signal,
        });
        const raw = await response.text();
        if (!response.ok) {
            return { ok: false, status: response.status, error: '端点返回 HTTP ' + response.status + '：' + raw.slice(0, 240) };
        }
        try {
            return { ok: true, status: response.status, json: JSON.parse(raw) };
        } catch {
            return { ok: false, status: response.status, error: '端点返回的不是 JSON。' };
        }
    } catch (error) {
        const reason = timedOut ? '请求超时（' + timeoutMs + 'ms）' : (error?.message ?? String(error));
        return { ok: false, error: '请求失败：' + reason };
    } finally {
        clearTimeout(timer);
        if (signal !== undefined) signal.removeEventListener('abort', onAbort);
    }
}

function vectorsOf(list, expected) {
    const vectors = [];
    for (const item of list) {
        const vector = Array.isArray(item) ? item : (Array.isArray(item?.embedding) ? item.embedding : null);
        if (vector === null || vector.length === 0) {
            return { ok: false, error: '端点返回了空向量。' };
        }
        if (Number.isFinite(item?.index) && item.index >= 0 && item.index < expected) {
            vectors[item.index] = vector;
            continue;
        }
        vectors.push(vector);
    }
    if (vectors.length !== expected || vectors.some((vector) => !Array.isArray(vector))) {
        return { ok: false, error: '端点返回 ' + vectors.filter(Boolean).length + ' 条向量，期望 ' + expected + ' 条。' };
    }
    const dims = new Set(vectors.map((vector) => vector.length));
    if (dims.size !== 1) {
        return { ok: false, error: '端点返回的向量维度不一致：' + [...dims].join(' / ') };
    }
    return {
        ok: true,
        vectors: vectors.map((vector) => new Float32Array(vector.map(Number))),
        dim: vectors[0].length,
    };
}

/**
 * 创建多模态 embedding 适配器。只做请求与响应解析，不写任何缓存。
 *
 * @param {object} rawConfig normalizeImageEmbedding 的输入
 * @param {{fetchImpl?:Function, env?:object}} [deps]
 * @returns {{provider:string, model:string, baseURL:string, label:string,
 *   embedTexts:Function, embedImages:Function, embed:Function}}
 */
export function createImageEmbedder(rawConfig = {}, deps = {}) {
    const cfg = normalizeImageEmbedding(rawConfig);
    const fetchImpl = deps.fetchImpl ?? globalThis.fetch;
    const env = deps.env ?? process.env;
    const headers = {};
    const apiKey = keyOf(cfg, env);
    if (apiKey !== '') headers.authorization = 'Bearer ' + apiKey;

    async function call(inputs, { inputType = 'document', signal } = {}) {
        const ready = imageEmbeddingReady(cfg);
        if (!ready.ok) return { ok: false, code: ready.code, error: ready.error };

        if (cfg.provider === 'voyage') {
            const body = {
                inputs: inputs.map((input) => ({ content: input.content })),
                model: cfg.model,
                input_type: inputType === 'query' ? 'query' : 'document',
            };
            // Voyage 的 multimodalembeddings 参考页没有 output_dimension 参数：
            // 模型维度固定，实际维度以响应/探针为准，不把用户填的 dimensions 发出去。
            const result = await postJson({
                fetchImpl,
                url: cfg.baseURL + '/multimodalembeddings',
                headers,
                body,
                timeoutMs: cfg.timeoutMs,
                signal,
            });
            if (!result.ok) return { ok: false, code: 'http', error: result.error };
            const data = Array.isArray(result.json?.data) ? result.json.data : null;
            if (data === null) return { ok: false, code: 'bad-response', error: 'Voyage 响应里没有 data 数组。' };
            const parsed = vectorsOf(data, inputs.length);
            if (!parsed.ok) return { ok: false, code: 'bad-response', error: 'Voyage：' + parsed.error };
            return { ok: true, vectors: parsed.vectors, dim: parsed.dim, model: asText(result.json?.model) || cfg.model };
        }

        if (cfg.provider === 'cohere') {
            const body = {
                model: cfg.model,
                inputs: inputs.map((input) => ({ content: input.content })),
                input_type: inputType === 'query' ? 'search_query' : 'search_document',
                embedding_types: ['float'],
            };
            if (cfg.dimensions > 0) body.output_dimension = cfg.dimensions;
            const result = await postJson({
                fetchImpl,
                url: cfg.baseURL + '/v2/embed',
                headers,
                body,
                timeoutMs: cfg.timeoutMs,
                signal,
            });
            if (!result.ok) return { ok: false, code: 'http', error: result.error };
            const list = result.json?.embeddings?.float;
            if (!Array.isArray(list)) {
                return { ok: false, code: 'bad-response', error: 'Cohere 响应里没有 embeddings.float（检查 embedding_types）。' };
            }
            const parsed = vectorsOf(list, inputs.length);
            if (!parsed.ok) return { ok: false, code: 'bad-response', error: 'Cohere：' + parsed.error };
            return { ok: true, vectors: parsed.vectors, dim: parsed.dim, model: cfg.model };
        }

        // vLLM / OpenAI 兼容：文本走 input 数组，图片走 messages 里的 image_url。
        const body = { model: cfg.model, encoding_format: 'float' };
        if (inputs.length === 1 && inputs[0].kind === 'image') {
            body.messages = [{ role: 'user', content: [{ type: 'image_url', image_url: { url: inputs[0].url } }] }];
        } else if (inputs.every((input) => input.kind === 'text')) {
            body.input = inputs.map((input) => input.text);
        } else {
            return {
                ok: false,
                code: 'unsupported',
                error: 'vLLM / OpenAI 兼容通道不支持同一请求里混发图片与文本；请分开调用。',
            };
        }
        const result = await postJson({
            fetchImpl,
            url: cfg.baseURL + '/embeddings',
            headers,
            body,
            timeoutMs: cfg.timeoutMs,
            signal,
        });
        if (!result.ok) return { ok: false, code: 'http', error: result.error };
        const data = Array.isArray(result.json?.data) ? result.json.data : null;
        if (data === null) return { ok: false, code: 'bad-response', error: 'embedding 响应里没有 data 数组。' };
        const parsed = vectorsOf(data, inputs.length);
        if (!parsed.ok) return { ok: false, code: 'bad-response', error: 'embedding：' + parsed.error };
        return { ok: true, vectors: parsed.vectors, dim: parsed.dim, model: asText(result.json?.model) || cfg.model };
    }

    return {
        provider: cfg.provider,
        label: cfg.label,
        model: cfg.model,
        baseURL: cfg.baseURL,
        config: cfg,
        embedTexts(texts, options = {}) {
            const list = (Array.isArray(texts) ? texts : []).map((text) => String(text ?? ''));
            if (list.length === 0) return Promise.resolve({ ok: true, vectors: [], dim: 0, model: cfg.model });
            return call(list.map((text) => ({ kind: 'text', text, content: [{ type: 'text', text }] })), options);
        },
        embedImages(images, options = {}) {
            const list = Array.isArray(images) ? images : [];
            if (list.length === 0) return Promise.resolve({ ok: true, vectors: [], dim: 0, model: cfg.model });
            return call(list.map((image) => {
                const url = dataUrlOf(image.bytes, image.mediaType ?? 'image/png');
                return {
                    kind: 'image',
                    url,
                    content: cfg.provider === 'cohere'
                        ? [{ type: 'image_url', image_url: { url } }]
                        : [{ type: 'image_base64', image_base64: url }],
                };
            }), options);
        },
        embed(inputs, options = {}) {
            const list = Array.isArray(inputs) ? inputs : [];
            return call(list.map((input) => {
                if (input.kind === 'image') {
                    const url = dataUrlOf(input.bytes, input.mediaType ?? 'image/png');
                    return {
                        kind: 'image',
                        url,
                        content: cfg.provider === 'cohere'
                            ? [{ type: 'image_url', image_url: { url } }]
                            : [{ type: 'image_base64', image_base64: url }],
                    };
                }
                const text = String(input.text ?? '');
                return { kind: 'text', text, content: [{ type: 'text', text }] };
            }), options);
        },
    };
}

/**
 * 探针：真的发一次文本 + 一次 1×1 PNG，要求两者维度相同。
 * 只有探针通过，图库才会把「图文对齐」当作可用信号；否则检索时明确跳过。
 */
export async function probeImageEmbedding(embedder, { signal, text = '审美参考图' } = {}) {
    if (embedder === null || typeof embedder !== 'object') {
        return { ok: false, code: 'unconfigured', error: '没有可用的多模态 embedding 适配器。' };
    }
    const textResult = await embedder.embedTexts([text], { inputType: 'query', signal });
    if (!textResult.ok) {
        return { ok: false, code: 'text-failed', error: '探针文本向量失败：' + textResult.error };
    }
    const imageResult = await embedder.embedImages([{ bytes: ONE_PIXEL_PNG, mediaType: 'image/png' }], { inputType: 'document', signal });
    if (!imageResult.ok) {
        return { ok: false, code: 'image-failed', error: '探针图片向量失败：' + imageResult.error };
    }
    const textDim = textResult.vectors[0]?.length ?? 0;
    const imageDim = imageResult.vectors[0]?.length ?? 0;
    if (textDim === 0 || imageDim === 0) {
        return { ok: false, code: 'empty-vector', error: '探针返回了空向量。' };
    }
    if (textDim !== imageDim) {
        return {
            ok: false,
            code: 'dim-mismatch',
            error: '文本向量 ' + textDim + ' 维、图片向量 ' + imageDim + ' 维，不是同一个空间，不能做图文对齐检索。',
            textDim,
            imageDim,
        };
    }
    return {
        ok: true,
        dim: textDim,
        provider: embedder.provider,
        model: embedder.model,
        at: new Date().toISOString(),
    };
}

/** 视觉描述提示词：只要画面可见的审美特征，禁止推断作者/版权/出处。 */
export function buildVisionPrompt({ language = 'zh' } = {}) {
    if (language === 'en') {
        return [
            'You are describing an image for an aesthetic reference library.',
            'Describe only what is visible. Do NOT infer the author, copyright, source, price, or identity of anyone.',
            'Return ONE JSON object, no markdown, with exactly these keys:',
            '{"subject":"...","composition":"...","palette":["color1","color2"],"light":"...","material":"...","mood":"...","summary":"...","tags":["tag1","tag2"]}',
            'Keep each text field under 30 words; 5-10 tags, each 1-3 words, in English.',
        ].join('\n');
    }
    return [
        '你在为「审美参考图库」描述一张图片。',
        '只描述画面里看得见的审美特征；不要推断作者、版权、出处、价格或任何人的身份。',
        '只返回一个 JSON 对象，不要 Markdown、不要解释，键名固定为：',
        '{"subject":"主体","composition":"构图","palette":["主色1","主色2"],"light":"光线","material":"材质与质感","mood":"氛围","summary":"一句话概括","tags":["标签1","标签2"]}',
        '每个文字字段不超过 60 字；tags 给 5-10 个，每个 2-8 字，用中文。',
    ].join('\n');
}

/** 把模型输出解析成对象：容忍代码围栏、前后解释，失败时给可读错误与预览。 */
export function parseAestheticJson(text) {
    const raw = String(text ?? '').trim();
    if (raw === '') return { ok: false, error: '模型没有返回任何文本。', preview: '' };
    const unfenced = raw.replace(/^\u0060{3,}[a-zA-Z]*\s*/u, '').replace(/\u0060{3,}\s*$/u, '').trim();
    const candidates = [unfenced, raw];
    for (const candidate of candidates) {
        const start = candidate.indexOf('{');
        const end = candidate.lastIndexOf('}');
        if (start < 0 || end <= start) continue;
        try {
            const value = JSON.parse(candidate.slice(start, end + 1));
            if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
                return { ok: true, value };
            }
        } catch { /* 继续试下一个候选 */ }
    }
    return {
        ok: false,
        error: '模型没有返回可解析的 JSON 对象。',
        preview: raw.length > 200 ? raw.slice(0, 200) + '…' : raw,
    };
}

function contentTextOf(content) {
    if (typeof content === 'string') return content;
    if (!Array.isArray(content)) return '';
    return content
        .filter((block) => block?.type === 'text' || block?.type === 'output_text')
        .map((block) => String(block.text ?? ''))
        .join('');
}

/**
 * 宿主视觉通道：attachments.saveImage → llm.prepareCall → stream。
 * 只读既有服务，不缓存、不落盘。
 */
export async function describeWithHost({
    host = null,
    bytes,
    mediaType = 'image/png',
    prompt,
    provider = '',
    model = '',
    maxTokens = 900,
    signal,
} = {}) {
    const attachments = host?.attachments;
    const llm = host?.llm;
    if (typeof attachments?.saveImage !== 'function') {
        return { ok: false, code: 'host-unavailable', error: '宿主没有附件服务（attachments.saveImage），宿主视觉通道不可用。' };
    }
    if (typeof llm?.prepareCall !== 'function') {
        return { ok: false, code: 'host-unavailable', error: '宿主没有 llm.prepareCall，宿主视觉通道不可用。' };
    }
    if (provider === '' || model === '') {
        return { ok: false, code: 'no-model', error: '没有选择视觉模型（provider / model）。' };
    }
    let ref;
    try {
        ref = await attachments.saveImage({
            data: new Uint8Array(bytes),
            mediaType,
            name: 'gallery-' + Date.now() + (mediaType === 'image/jpeg' ? '.jpg' : '.png'),
        });
    } catch (error) {
        return { ok: false, code: 'attach-failed', error: '图片交给宿主附件服务失败：' + (error?.message ?? String(error)) };
    }
    let prepared;
    try {
        prepared = await llm.prepareCall({ provider, model, maxTokens }, signal);
    } catch (error) {
        return { ok: false, code: 'prepare-failed', error: '模型 ' + provider + '/' + model + ' 不可用：' + (error?.message ?? String(error)) };
    }
    const modalities = prepared?.inputModalities;
    if (Array.isArray(modalities) && !modalities.includes('image')) {
        return {
            ok: false,
            code: 'no-vision',
            error: '模型 ' + provider + '/' + model + ' 不接受图片输入（inputModalities 不含 image）。',
            modalities: [...modalities],
        };
    }
    const config = prepared?.config !== null && typeof prepared?.config === 'object' ? prepared.config : {};
    const options = {
        ...config,
        provider,
        model,
        maxTokens,
        purpose: host?.purpose ?? 'office-gallery',
        messages: [{
            role: 'user',
            content: [
                { type: 'text', text: prompt },
                { type: 'image', attachment: ref },
            ],
        }],
        ...(signal === undefined ? {} : { signal }),
    };
    let text = '';
    try {
        for await (const chunk of prepared.stream(options)) {
            if (chunk?.type === 'text-delta') text += String(chunk.text ?? '');
        }
    } catch (error) {
        return { ok: false, code: 'stream-failed', error: '视觉模型调用失败：' + (error?.message ?? String(error)) };
    }
    if (text.trim() === '') {
        return { ok: false, code: 'empty', error: '视觉模型没有返回文本（可能这个模型不支持图片，或输出被截断）。' };
    }
    return {
        ok: true,
        text,
        channel: 'host',
        provider,
        model,
        modalitiesKnown: Array.isArray(modalities),
    };
}

/** OpenAI 兼容视觉通道：POST {baseURL}/chat/completions，图片用 data URL。 */
export async function describeWithOpenAi({
    bytes,
    mediaType = 'image/png',
    prompt,
    baseURL = '',
    apiKey = '',
    model = '',
    maxTokens = 900,
    timeoutMs = 60000,
    fetchImpl = globalThis.fetch,
    signal,
} = {}) {
    if (baseURL === '' || model === '') {
        return { ok: false, code: 'unconfigured', error: 'OpenAI 兼容视觉通道需要端点与模型名。' };
    }
    const headers = {};
    if (apiKey !== '') headers.authorization = 'Bearer ' + apiKey;
    const result = await postJson({
        fetchImpl,
        url: baseURL.replace(/\/+$/, '') + '/chat/completions',
        headers,
        body: {
            model,
            temperature: 0.2,
            max_tokens: maxTokens,
            messages: [{
                role: 'user',
                content: [
                    { type: 'text', text: prompt },
                    { type: 'image_url', image_url: { url: dataUrlOf(bytes, mediaType) } },
                ],
            }],
        },
        timeoutMs,
        signal,
    });
    if (!result.ok) return { ok: false, code: 'http', error: result.error };
    const text = contentTextOf(result.json?.choices?.[0]?.message?.content);
    if (text.trim() === '') {
        return { ok: false, code: 'empty', error: '视觉模型没有返回文本。' };
    }
    return { ok: true, text, channel: 'openai', provider: 'openai', model, modalitiesKnown: true };
}

/**
 * 按配置选通道描述一张图。没有配置 / 调用失败都返回 ok:false，绝不编造描述。
 * 是否会联网由调用方决定：图库只在显式 allowUpload 时才调到这里。
 */
export async function describeImage({
    config = {},
    bytes,
    mediaType = 'image/png',
    host = null,
    fetchImpl = globalThis.fetch,
    env = process.env,
    signal,
} = {}) {
    const cfg = normalizeVisionConfig(config);
    const ready = visionReady(cfg);
    if (!ready.ok) return { ok: false, code: ready.code, error: ready.error };
    const prompt = buildVisionPrompt({ language: cfg.language });
    if (cfg.channel === 'host') {
        return describeWithHost({
            host,
            bytes,
            mediaType,
            prompt,
            provider: cfg.provider,
            model: cfg.model,
            maxTokens: cfg.maxTokens,
            signal,
        });
    }
    return describeWithOpenAi({
        bytes,
        mediaType,
        prompt,
        baseURL: cfg.baseURL,
        apiKey: keyOf(cfg, env),
        model: cfg.model,
        maxTokens: cfg.maxTokens,
        timeoutMs: cfg.timeoutMs,
        fetchImpl,
        signal,
    });
}
