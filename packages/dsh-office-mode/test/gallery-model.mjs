/**
 * 图库模型层（gallery-model.js）的契约测试。
 *
 * 不联网、不需要 Key：多模态 adapter 用假 fetch 按官方响应形状返回，
 * 重点验证「请求体是否符合已核实的契约」与「失败是否被如实报告」。
 *
 * 跑法：node test/gallery-model.mjs
 */
import assert from 'node:assert/strict';
import { sniffImageBytes } from '../src/image-source.js';
import {
    GALLERY_PROMPT_VERSION,
    ONE_PIXEL_PNG,
    buildVisionPrompt,
    createImageEmbedder,
    dataUrlOf,
    describeImage,
    describeWithHost,
    describeWithOpenAi,
    imageEmbeddingReady,
    normalizeGalleryConfig,
    normalizeImageEmbedding,
    normalizeVisionConfig,
    parseAestheticJson,
    probeImageEmbedding,
    visionReady,
} from '../src/gallery-model.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

function jsonResponse(payload, status = 200) {
    return {
        ok: status >= 200 && status < 300,
        status,
        text: async () => (typeof payload === 'string' ? payload : JSON.stringify(payload)),
    };
}

function vec(values) {
    return values;
}

await check('normalizeImageEmbedding：通道别名与官方默认值', async () => {
    const voyage = normalizeImageEmbedding({ provider: 'voyage' });
    assert.equal(voyage.provider, 'voyage');
    assert.equal(voyage.baseURL, 'https://api.voyageai.com/v1');
    assert.equal(voyage.model, 'voyage-multimodal-3.5');
    assert.equal(voyage.apiKeyEnv, 'VOYAGE_API_KEY');

    const cohere = normalizeImageEmbedding({ provider: 'cohere' });
    assert.equal(cohere.baseURL, 'https://api.cohere.com');
    assert.equal(cohere.model, 'embed-v4.0');

    assert.equal(normalizeImageEmbedding({ provider: 'off' }).provider, 'none');
    assert.equal(normalizeImageEmbedding({ provider: 'openai' }).provider, 'vllm');
    assert.equal(normalizeImageEmbedding({ provider: 'wat' }).provider, 'none');
    assert.equal(normalizeImageEmbedding({ provider: 'vllm', dimensions: 9999 }).dimensions, 4096);
});

await check('imageEmbeddingReady / visionReady：缺配置就明确说缺什么', async () => {
    assert.equal(imageEmbeddingReady({}).code, 'unconfigured');
    assert.match(imageEmbeddingReady({}).error, /没有配置图文对齐/);
    assert.equal(imageEmbeddingReady({ provider: 'vllm', model: '', baseURL: 'http://127.0.0.1:8000/v1' }).code, 'unconfigured');
    assert.equal(imageEmbeddingReady({ provider: 'vllm', model: 'm', baseURL: '' }).code, 'unconfigured');
    assert.equal(imageEmbeddingReady({ provider: 'voyage' }).ok, true);

    assert.equal(normalizeVisionConfig({}).channel, 'off');
    assert.equal(visionReady({}).code, 'unconfigured');
    assert.equal(visionReady({ channel: 'host', provider: 'p' }).ok, false);
    assert.equal(visionReady({ channel: 'openai', baseURL: 'https://x', model: 'm' }).ok, true);
});

await check('normalizeGalleryConfig：文本 / 图片 / 视觉三段各自收口', async () => {
    const cfg = normalizeGalleryConfig({
        textEmbedding: { mode: 'local', dimensions: 128 },
        imageEmbedding: { provider: 'voyage' },
        vision: { channel: 'host', provider: 'p', model: 'm' },
    });
    assert.equal(cfg.textEmbedding.mode, 'local');
    assert.equal(cfg.textEmbedding.dimensions, 128);
    assert.equal(cfg.imageEmbedding.model, 'voyage-multimodal-3.5');
    assert.equal(cfg.vision.channel, 'host');
    assert.equal(cfg.vision.maxTokens, 900);
});

await check('ONE_PIXEL_PNG：探针用的图确实是可识别的 1×1 PNG', async () => {
    const sniffed = sniffImageBytes(ONE_PIXEL_PNG);
    assert.ok(sniffed);
    assert.equal(sniffed.mime, 'image/png');
    assert.equal(sniffed.width, 1);
    assert.equal(sniffed.height, 1);
    assert.equal(dataUrlOf(ONE_PIXEL_PNG, 'image/png').startsWith('data:image/png;base64,'), true);
});

await check('Voyage：请求体按 multimodalembeddings 契约，响应按 index 归位', async () => {
    const calls = [];
    const fetchImpl = async (url, init) => {
        calls.push({ url, body: JSON.parse(init.body), headers: init.headers });
        return jsonResponse({
            object: 'list',
            model: 'voyage-multimodal-3.5',
            data: [
                { embedding: vec([0, 1, 0]), index: 1 },
                { embedding: vec([1, 0, 0]), index: 0 },
            ],
        });
    };
    const embedder = createImageEmbedder({ provider: 'voyage', apiKey: 'k', dimensions: 3 }, { fetchImpl, env: {} });
    const result = await embedder.embed([
        { kind: 'text', text: '海边' },
        { kind: 'image', bytes: ONE_PIXEL_PNG, mediaType: 'image/png' },
    ], { inputType: 'query' });
    assert.equal(result.ok, true);
    assert.equal(result.dim, 3);
    assert.deepEqual([...result.vectors[0]], [1, 0, 0], 'index 1 的向量必须归位到第 2 条');
    assert.deepEqual([...result.vectors[1]], [0, 1, 0]);
    assert.equal(calls[0].url, 'https://api.voyageai.com/v1/multimodalembeddings');
    assert.equal(calls[0].headers.authorization, 'Bearer k');
    assert.equal(calls[0].body.model, 'voyage-multimodal-3.5');
    assert.equal(calls[0].body.input_type, 'query');
    assert.equal('output_dimension' in calls[0].body, false, 'Voyage multimodal 没有 output_dimension 参数，不能发');
    assert.deepEqual(calls[0].body.inputs[0].content, [{ type: 'text', text: '海边' }]);
    assert.equal(calls[0].body.inputs[1].content[0].type, 'image_base64');
    assert.match(calls[0].body.inputs[1].content[0].image_base64, /^data:image\/png;base64,/);
});

await check('Cohere：/v2/embed + embeddings.float，input_type 映射', async () => {
    const calls = [];
    const fetchImpl = async (url, init) => {
        calls.push({ url, body: JSON.parse(init.body) });
        return jsonResponse({ embeddings: { float: [vec([1, 2, 3])] }, model: 'embed-v4.0' });
    };
    const embedder = createImageEmbedder({ provider: 'cohere', apiKeyEnv: 'COHERE_KEY', dimensions: 3 }, { fetchImpl, env: { COHERE_KEY: 'ck' } });
    const result = await embedder.embedImages([{ bytes: ONE_PIXEL_PNG, mediaType: 'image/png' }], { inputType: 'query' });
    assert.equal(result.ok, true);
    assert.deepEqual([...result.vectors[0]], [1, 2, 3]);
    assert.equal(calls[0].url, 'https://api.cohere.com/v2/embed');
    assert.equal(calls[0].body.input_type, 'search_query');
    assert.deepEqual(calls[0].body.embedding_types, ['float']);
    assert.equal(calls[0].body.output_dimension, 3);
    assert.equal(calls[0].body.inputs[0].content[0].type, 'image_url');
    assert.match(calls[0].body.inputs[0].content[0].image_url.url, /^data:image\/png;base64,/);
});

await check('vLLM / OpenAI 兼容：文本走 input 数组，图片走 messages.image_url', async () => {
    const calls = [];
    const fetchImpl = async (url, init) => {
        calls.push({ url, body: JSON.parse(init.body) });
        return jsonResponse({ data: [{ embedding: vec([1, 2]), index: 0 }], model: 'vlm2vec' });
    };
    const embedder = createImageEmbedder({ provider: 'vllm', baseURL: 'http://127.0.0.1:8000/v1', model: 'vlm2vec' }, { fetchImpl, env: {} });
    assert.equal((await embedder.embedTexts(['日落'], { inputType: 'query' })).ok, true);
    assert.deepEqual(calls[0].body.input, ['日落']);
    assert.equal(calls[0].body.encoding_format, 'float');

    assert.equal((await embedder.embedImages([{ bytes: ONE_PIXEL_PNG, mediaType: 'image/png' }])).ok, true);
    assert.equal(calls[1].body.messages[0].content[0].type, 'image_url');
    assert.match(calls[1].body.messages[0].content[0].image_url.url, /^data:image\/png;base64,/);

    const mixed = await embedder.embed([{ kind: 'text', text: 'x' }, { kind: 'image', bytes: ONE_PIXEL_PNG }]);
    assert.equal(mixed.ok, false);
    assert.equal(mixed.code, 'unsupported', '混发会被端点误解，必须明确拒绝');
});

await check('响应维度不一致 / HTTP 错误 / 非 JSON：都返回可读失败，不算成功', async () => {
    const badDims = createImageEmbedder({ provider: 'voyage' }, {
        env: {},
        fetchImpl: async () => jsonResponse({ data: [{ embedding: [1, 2], index: 0 }, { embedding: [1, 2, 3], index: 1 }] }),
    });
    const dims = await badDims.embed([{ kind: 'text', text: 'a' }, { kind: 'text', text: 'b' }]);
    assert.equal(dims.ok, false);
    assert.match(dims.error, /维度不一致/);

    const http = createImageEmbedder({ provider: 'voyage' }, {
        env: {},
        fetchImpl: async () => jsonResponse('rate limited', 429),
    });
    const failed = await http.embedTexts(['x']);
    assert.equal(failed.ok, false);
    assert.equal(failed.code, 'http');
    assert.match(failed.error, /HTTP 429/);

    const notJson = createImageEmbedder({ provider: 'voyage' }, {
        env: {},
        fetchImpl: async () => jsonResponse('<html>not json</html>'),
    });
    const broken = await notJson.embedTexts(['x']);
    assert.equal(broken.ok, false);
    assert.match(broken.error, /不是 JSON/);
});

await check('探针：要求文本与图片同维；不同维返回 dim-mismatch', async () => {
    const okEmbedder = createImageEmbedder({ provider: 'voyage' }, {
        env: {},
        fetchImpl: async (url, init) => {
            const isText = JSON.parse(init.body).inputs[0].content[0].type === 'text';
            return jsonResponse({ data: [{ embedding: isText ? vec([1, 0]) : vec([0, 1]), index: 0 }] });
        },
    });
    const probed = await probeImageEmbedding(okEmbedder, {});
    assert.equal(probed.ok, true);
    assert.equal(probed.dim, 2);

    const mismatch = createImageEmbedder({ provider: 'voyage' }, {
        env: {},
        fetchImpl: async (url, init) => {
            const isText = JSON.parse(init.body).inputs[0].content[0].type === 'text';
            return jsonResponse({ data: [{ embedding: isText ? vec([1, 0, 0, 0]) : vec([1, 0]), index: 0 }] });
        },
    });
    const rejected = await probeImageEmbedding(mismatch, {});
    assert.equal(rejected.ok, false);
    assert.equal(rejected.code, 'dim-mismatch');
    assert.equal(rejected.textDim, 4);
    assert.equal(rejected.imageDim, 2);
});

await check('parseAestheticJson：容忍围栏与前后解释，失败给预览', async () => {
    const fenced = parseAestheticJson('\u0060\u0060\u0060json\n{"subject":"海边"}\n\u0060\u0060\u0060');
    assert.equal(fenced.ok, true);
    assert.equal(fenced.value.subject, '海边');
    const prose = parseAestheticJson('这是描述：{"subject":"夜景","tags":["霓虹"]} 希望有帮助');
    assert.equal(prose.ok, true);
    assert.deepEqual(prose.value.tags, ['霓虹']);
    const bad = parseAestheticJson('完全不是 JSON 的一段话');
    assert.equal(bad.ok, false);
    assert.ok(bad.preview.includes('完全不是'));
});

await check('buildVisionPrompt：要求只描述可见审美、禁止推断作者与版权、给出固定 JSON 键', async () => {
    const prompt = buildVisionPrompt({ language: 'zh' });
    assert.match(prompt, /不要推断作者、版权/);
    assert.match(prompt, /subject/);
    assert.match(prompt, /palette/);
    assert.match(prompt, /tags/);
    assert.equal(GALLERY_PROMPT_VERSION, 'gallery-vision-v1');
});

await check('describeWithOpenAi：请求形状、Key、失败传播与空输出', async () => {
    const calls = [];
    const fetchImpl = async (url, init) => {
        calls.push({ url, body: JSON.parse(init.body), headers: init.headers });
        return jsonResponse({ choices: [{ message: { content: '{"subject":"海边"}' } }] });
    };
    const described = await describeWithOpenAi({
        bytes: ONE_PIXEL_PNG,
        mediaType: 'image/png',
        prompt: 'p',
        baseURL: 'https://vision.example.com/v1/',
        apiKey: 'sk',
        model: 'vl',
        maxTokens: 300,
        fetchImpl,
    });
    assert.equal(described.ok, true);
    assert.equal(described.text, '{"subject":"海边"}');
    assert.equal(calls[0].url, 'https://vision.example.com/v1/chat/completions');
    assert.equal(calls[0].headers.authorization, 'Bearer sk');
    assert.equal(calls[0].body.max_tokens, 300);
    assert.equal(calls[0].body.messages[0].content[0].type, 'text');
    assert.equal(calls[0].body.messages[0].content[1].type, 'image_url');

    const noModel = await describeWithOpenAi({ bytes: ONE_PIXEL_PNG, prompt: 'p', baseURL: 'https://x', model: '' });
    assert.equal(noModel.code, 'unconfigured');

    const http = await describeWithOpenAi({
        bytes: ONE_PIXEL_PNG,
        prompt: 'p',
        baseURL: 'https://x',
        model: 'm',
        fetchImpl: async () => jsonResponse('boom', 500),
    });
    assert.equal(http.ok, false);
    assert.match(http.error, /HTTP 500/);

    const empty = await describeWithOpenAi({
        bytes: ONE_PIXEL_PNG,
        prompt: 'p',
        baseURL: 'https://x',
        model: 'm',
        fetchImpl: async () => jsonResponse({ choices: [{ message: { content: '  ' } }] }),
    });
    assert.equal(empty.ok, false);
    assert.equal(empty.code, 'empty');
});

await check('describeWithHost：缺服务 / 缺模型都明确失败', async () => {
    assert.equal((await describeWithHost({ bytes: ONE_PIXEL_PNG })).code, 'host-unavailable');
    assert.equal((await describeWithHost({ bytes: ONE_PIXEL_PNG, host: { attachments: { saveImage: async () => ({}) } } })).code, 'host-unavailable');
    assert.equal((await describeWithHost({
        bytes: ONE_PIXEL_PNG,
        host: { attachments: { saveImage: async () => ({}) }, llm: { prepareCall: async () => ({}) } },
        provider: '',
        model: '',
    })).code, 'no-model');
});

await check('describeImage：通道 off 时直接说未配置，不构造请求', async () => {
    let fetched = 0;
    const result = await describeImage({
        config: {},
        bytes: ONE_PIXEL_PNG,
        fetchImpl: async () => { fetched += 1; return jsonResponse({}); },
    });
    assert.equal(result.ok, false);
    assert.equal(result.code, 'unconfigured');
    assert.equal(fetched, 0);
});

const failed = results.filter((result) => !result.ok);
for (const result of results) {
    console.log((result.ok ? 'ok   ' : 'FAIL ') + result.name);
    if (!result.ok) console.log('     ' + (result.error?.stack ?? result.error));
}
console.log('');
console.log(results.length - failed.length + '/' + results.length + ' passed');
if (failed.length > 0) process.exitCode = 1;
