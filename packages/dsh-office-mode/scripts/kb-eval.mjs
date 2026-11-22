#!/usr/bin/env node
/**
 * 知识库检索离线评测（第一百零三轮）。
 *
 * 为什么单独一套而不是复用 `test/kb.mjs`：那一套测契约（入库、切块、范围、回执），
 * 这里量**检索质量**。评测要在同一份语料上反复改参数，所以不落盘、不进 sqlite：
 * 直接 `splitDocument` 切块 + `kbRankChunks` 打分 + 余弦向量，融合公式与
 * `memory.js` 的 `fuseKbHits` 完全一致（见 `fuse` 的注释）。
 *
 * 指标按**文档级**算：一个查询的相关文档是 jsonl 里的 `relevant` 路径，命中任一
 * 块就算这篇文档被召回。块级指标会被块大小直接影响，改 `chunkChars` 时没有可比性；
 * 文档级才是用户能感知的「找没找到」。
 *
 * 用法：
 *   node scripts/kb-eval.mjs                          # 默认 hybrid（本地哈希向量）
 *   node scripts/kb-eval.mjs --embedding none         # 纯词法基线
 *   node scripts/kb-eval.mjs --embedding local --local-model <模型目录>
 *   node scripts/kb-eval.mjs --sweep chunkChars=600,900,1200,1600
 *   node scripts/kb-eval.mjs --sweep efSearch=32,64,128
 *   node scripts/kb-eval.mjs --config lexicalWeight=0.4,candidateLimit=20
 *   node scripts/kb-eval.mjs --save-baseline default-hybrid
 *   node scripts/kb-eval.mjs --compare default-hybrid
 *
 * 语料是仓库根 `docs/*.md`（全量，不挑）；查询与相关文档在
 * `test/fixtures/kb-eval.jsonl`。全程离线：本地哈希不需要网络，本地小模型要
 * 本机 Python + 已下载的权重（`--local-model` 指过去）。
 */
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { chunkHeaderOf, kbRankChunks, splitDocument } from '../src/kb.js';
import { cosineSimilarity, embedTexts, localEmbedding } from '../src/kb-embed.js';
import { createAnnIndex } from '../src/kb-ann.js';
import { createLocalEmbeddingManager } from '../src/kb-local-embed.js';

const ROOT = fileURLToPath(new URL('../../../', import.meta.url));
const FIXTURE = join(ROOT, 'packages/dsh-office-mode/test/fixtures/kb-eval.jsonl');
const BASELINES = join(ROOT, 'packages/dsh-office-mode/test/fixtures/kb-eval-baselines.json');
const DOCS_DIR = join(ROOT, 'docs');

const DEFAULTS = {
    chunkChars: 1200,
    overlap: 0.125, // 与 chunkOverlapFor(1200)=150 对齐
    lexicalWeight: 0.55,
    vectorWeight: 0.45,
    candidateLimit: 30,
    semanticGate: 0.25,
    efSearch: 64,
};

function parseConfig(text) {
    const out = {};
    for (const piece of String(text ?? '').split(',')) {
        const [key, raw] = piece.split('=');
        if (key === undefined || raw === undefined) continue;
        const value = Number(raw);
        if (Number.isFinite(value)) out[key.trim()] = value;
    }
    return out;
}

function parseArgs(argv) {
    const out = {
        config: {}, sweep: null, embedding: 'hash', limit: 10, json: false,
        save: '', compare: '', localModel: '', python: '',
    };
    for (let i = 0; i < argv.length; i += 1) {
        const arg = argv[i];
        const next = () => argv[i + 1];
        if (arg === '--config') { out.config = parseConfig(next()); i += 1; }
        else if (arg === '--sweep') {
            const [key, values] = String(next() ?? '').split('=');
            if (key && values) { out.sweep = { key: key.trim(), values: values.split(',').map(Number).filter(Number.isFinite) }; }
            i += 1;
        } else if (arg === '--embedding') { out.embedding = next(); i += 1; }
        else if (arg === '--local-model') { out.localModel = next(); i += 1; }
        else if (arg === '--python') { out.python = next(); i += 1; }
        else if (arg === '--limit') { out.limit = Number(next()) || 10; i += 1; }
        else if (arg === '--save-baseline') { out.save = next(); i += 1; }
        else if (arg === '--compare') { out.compare = next(); i += 1; }
        else if (arg === '--json') out.json = true;
    }
    return out;
}

function loadFixture() {
    return readFileSync(FIXTURE, 'utf8').split(/\r?\n/).filter((line) => line.trim() !== '').map((line) => JSON.parse(line));
}

function loadDocs() {
    return readdirSync(DOCS_DIR)
        .filter((name) => name.endsWith('.md'))
        .sort()
        .map((name) => {
            const rel = `docs/${name}`;
            return { path: rel, text: readFileSync(join(DOCS_DIR, name), 'utf8') };
        });
}

function buildChunks(docs, { chunkChars = 1200, overlap = 0.125 } = {}) {
    const chunks = [];
    for (const doc of docs) {
        const pieces = splitDocument(doc.text, { chunkChars, overlap: Math.round(chunkChars * overlap) });
        for (const piece of pieces) {
            chunks.push({
                id: `${doc.path}#${piece.n}`,
                doc: doc.path,
                text: piece.text,
                head: chunkHeaderOf({ outline: piece.outline, source: doc.path }),
            });
        }
    }
    return chunks;
}

/** 与 `memory.js` 的 fuseKbHits 逐行一致：各自按最大值归一化后加权相加。 */
function fuse(lexical, vector, weights) {
    const maxLex = lexical.reduce((max, hit) => Math.max(max, Math.abs(Number(hit.score) || 0)), 0) || 1;
    const maxVec = vector.reduce((max, hit) => Math.max(max, Number(hit.score) || 0), 0) || 1;
    const fused = new Map();
    const put = (hit, score) => {
        const id = String(hit.chunkId ?? '');
        if (id === '') return;
        const before = fused.get(id);
        if (before === undefined) fused.set(id, { ...hit, score, lexical: 0, vector: 0 });
        else before.score += score;
    };
    for (const hit of lexical) {
        const score = weights.lexicalWeight * (Math.abs(Number(hit.score) || 0) / maxLex);
        put(hit, score);
        fused.get(String(hit.chunkId)).lexical = Math.abs(Number(hit.score) || 0);
    }
    for (const hit of vector) {
        const score = weights.vectorWeight * (Number(hit.score) || 0) / maxVec;
        put(hit, score);
        fused.get(String(hit.chunkId)).vector = Number(hit.score) || 0;
    }
    return [...fused.values()].sort((left, right) => (right.score - left.score) || String(left.chunkId).localeCompare(String(right.chunkId)));
}

function rankDocs(fused, limit) {
    const seen = [];
    for (const hit of fused) {
        const doc = String(hit.doc ?? '');
        if (doc !== '' && !seen.includes(doc)) seen.push(doc);
        if (seen.length >= limit) break;
    }
    return seen;
}

function score(rankings, fixture) {
    const k = 10;
    let recall5 = 0;
    let recall10 = 0;
    let mrr = 0;
    let ndcg = 0;
    const misses = [];
    fixture.forEach((row, index) => {
        const relevant = new Set(row.relevant ?? []);
        const ranked = rankings[index] ?? [];
        const top5 = ranked.slice(0, 5);
        const top10 = ranked.slice(0, k);
        recall5 += relevant.size === 0 ? 0 : top5.filter((doc) => relevant.has(doc)).length / relevant.size;
        recall10 += relevant.size === 0 ? 0 : top10.filter((doc) => relevant.has(doc)).length / relevant.size;
        const at = ranked.findIndex((doc) => relevant.has(doc));
        if (at >= 0) mrr += 1 / (at + 1);
        else misses.push(row.query);
        let dcg = 0;
        top10.forEach((doc, i) => { if (relevant.has(doc)) dcg += 1 / Math.log2(i + 2); });
        let idcg = 0;
        for (let i = 0; i < Math.min(relevant.size, k); i += 1) idcg += 1 / Math.log2(i + 2);
        ndcg += idcg === 0 ? 0 : dcg / idcg;
    });
    const total = fixture.length || 1;
    return {
        recall5: recall5 / total,
        recall10: recall10 / total,
        mrr: mrr / total,
        ndcg10: ndcg / total,
        misses,
    };
}

function pct(value) { return `${(value * 100).toFixed(1)}%`; }

async function makeEmbedder(kind, options) {
    if (kind === 'none') return null;
    if (kind === 'hash') {
        const dim = 256;
        return {
            name: 'local-hash-v1',
            dim,
            embed: async (texts) => texts.map((text) => localEmbedding(text, dim)),
        };
    }
    if (kind !== 'local') throw new Error(`不认识的 --embedding：${kind}（none / hash / local）`);
    const manager = createLocalEmbeddingManager();
    const config = {
        enabled: true,
        command: options.python,
        modelDir: options.localModel,
        idleTimeoutMs: 900_000,
        maxRssMb: 600,
        startupTimeoutMs: 120_000,
        maxLength: 256,
        threads: 2,
    };
    const info = await manager.ensure({ config, kbDir: join(ROOT, 'packages/dsh-office-mode/.office'), waitMs: 600_000 });
    if (!info.ok) throw new Error(`本地推理没起来：${info.error}`);
    console.log(`本地推理：${info.baseURL} • ${info.model} • ${info.dim} 维（PID ${info.pid}）`);
    return {
        name: info.model,
        dim: info.dim,
        manager,
        embed: async (texts) => {
            // embedTexts 不按 batchSize 切批（一个请求原样发出去），这里必须自己切：
            // 2067 块一起发会让服务端单个请求编码几十万 token，把推理锁占住好几分钟。
            const vectors = [];
            for (let start = 0; start < texts.length; start += 8) {
                const batch = texts.slice(start, start + 8);
                const result = await embedTexts({ texts: batch, config: { mode: 'remote', baseURL: info.baseURL, model: info.model, apiKey: '', apiKeyEnv: '', dimensions: info.dim } });
                if (!result.ok) throw new Error(`向量化失败：${result.error}`);
                vectors.push(...result.vectors);
            }
            return vectors;
        },
    };
}

/** 每段文本的向量只算一次：扫参数时同一块反复用。 */
function vectorCacheOf(embedder) {
    const cache = new Map();
    return async (texts) => {
        const missing = [...new Set(texts.filter((text) => !cache.has(text)))];
        if (missing.length > 0 && embedder !== null) {
            const vectors = await embedder.embed(missing);
            missing.forEach((text, i) => cache.set(text, vectors[i]));
        }
        return texts.map((text) => cache.get(text) ?? []);
    };
}

// 扫参数时语料与切块反复用：按「块大小 + 重叠」缓存块、向量与每问的完整排序。
// 没有这层缓存，扫一轮 chunkChars 要把 103 篇文档重新切七遍、BM25 重算七遍。
const chunkCache = new Map();
const lexicalCache = new Map();
const vectorRankCache = new Map();
const annCache = new Map();
const annIndexCache = new Map();

function signatureOf(config) {
    return `${config.chunkChars}:${config.overlap}`;
}

async function prepare(config, docs, embedder, embedText) {
    const sig = signatureOf(config);
    if (!chunkCache.has(sig)) {
        const chunks = buildChunks(docs, config);
        if (embedder !== null) {
            const vectors = await embedText(chunks.map((chunk) => chunk.text));
            chunks.forEach((chunk, i) => { chunk.vector = vectors[i]; });
        }
        chunkCache.set(sig, chunks);
    }
    return { sig, chunks: chunkCache.get(sig) };
}

async function evaluate({ fixture, docs, config, embedder, embedText, limit = 10 }) {
    const started = Date.now();
    const { sig, chunks } = await prepare(config, docs, embedder, embedText);
    const rankings = [];
    const annRankings = [];
    let index = null;
    if (embedder !== null && config.efSearch !== undefined) {
        if (!annIndexCache.has(sig)) {
            const built = createAnnIndex({ m: 16, efConstruction: 200, seed: 0x5eed1e55 });
            for (const chunk of chunks) built.add(chunk.id, chunk.vector);
            annIndexCache.set(sig, built);
        }
        index = annIndexCache.get(sig);
    }
    for (const row of fixture) {
        const lexKey = `${sig}|${row.query}`;
        if (!lexicalCache.has(lexKey)) {
            lexicalCache.set(lexKey, kbRankChunks(chunks, row.query).map((entry) => ({
                chunkId: entry.chunk.id, doc: entry.chunk.doc, score: entry.score,
            })));
        }
        const lexical = lexicalCache.get(lexKey).slice(0, config.candidateLimit);
        let vector = [];
        if (embedder !== null) {
            const vecKey = `${sig}|${row.query}`;
            if (!vectorRankCache.has(vecKey)) {
                const [queryVector] = await embedText([row.query]);
                vectorRankCache.set(vecKey, chunks
                    .map((chunk) => ({ chunkId: chunk.id, doc: chunk.doc, score: cosineSimilarity(queryVector, chunk.vector) }))
                    .filter((hit) => hit.score > 0)
                    .sort((left, right) => right.score - left.score));
            }
            vector = vectorRankCache.get(vecKey).slice(0, config.candidateLimit);
        }
        let fused = fuse(lexical, vector, config);
        if (lexical.length === 0 && fused.length > 0 && (fused[0].vector || 0) < config.semanticGate) fused = [];
        // 筛选截断上限：假设筛选模型把候选全留下，只量「最多保留 N 块」对召回的影响。
        if (Number.isFinite(config.maxSelected) && config.maxSelected > 0) fused = fused.slice(0, config.maxSelected);
        rankings.push(rankDocs(fused, limit));

        if (index !== null) {
            const annKey = `${sig}|${config.efSearch}|${row.query}`;
            if (annCache.has(annKey)) {
                annRankings.push(annCache.get(annKey));
            } else {
                const [queryVector] = await embedText([row.query]);
                const exact = chunks
                    .map((chunk) => ({ id: chunk.id, score: cosineSimilarity(queryVector, chunk.vector) }))
                    .sort((left, right) => right.score - left.score).slice(0, 10).map((hit) => hit.id);
                const approx = index.search(queryVector, 10, { efSearch: config.efSearch }).map((hit) => hit.id);
                const recall = approx.filter((id) => exact.includes(id)).length / 10;
                annCache.set(annKey, recall);
                annRankings.push(recall);
            }
        }
    }
    const metrics = score(rankings, fixture);
    metrics.annRecall10 = annRankings.length === 0 ? null : annRankings.reduce((sum, value) => sum + value, 0) / annRankings.length;
    metrics.chunks = chunks.length;
    metrics.ms = Date.now() - started;
    return metrics;
}

function printRow(label, metrics) {
    const ann = metrics.annRecall10 === null ? '' : `  ANN R@10 ${pct(metrics.annRecall10)}`;
    console.log(`${label.padEnd(38)} R@5 ${pct(metrics.recall5)}  R@10 ${pct(metrics.recall10)}  MRR ${metrics.mrr.toFixed(3)}  nDCG@10 ${metrics.ndcg10.toFixed(3)}  块 ${metrics.chunks}  ${metrics.ms}ms${ann}`);
}

async function main() {
    const args = parseArgs(process.argv.slice(2));
    const fixture = loadFixture();
    const docs = loadDocs();
    const embedder = await makeEmbedder(args.embedding, args);
    const embedText = vectorCacheOf(embedder);
    const config = { ...DEFAULTS, ...args.config };
    console.log(`语料 ${docs.length} 篇 / 查询 ${fixture.length} 条 / 向量通道 ${embedder === null ? '无（纯词法）' : embedder.name}`);
    try {
        const base = await evaluate({ fixture, docs, config, embedder, embedText, limit: args.limit });
        printRow('基线', base);
        if (base.misses.length > 0) console.log(`  未召回：${base.misses.slice(0, 6).join(' / ')}${base.misses.length > 6 ? ' …' : ''}`);

        const results = { base, sweep: [] };
        if (args.sweep !== null && args.sweep.values.length > 0) {
            for (const value of args.sweep.values) {
                const changed = { ...config, [args.sweep.key]: value };
                const metrics = await evaluate({ fixture, docs, config: changed, embedder, embedText, limit: args.limit });
                results.sweep.push({ value, metrics });
                printRow(`${args.sweep.key}=${value}`, metrics);
            }
        }
        if (args.save !== '') {
            let store = {};
            try { store = JSON.parse(readFileSync(BASELINES, 'utf8')); } catch { store = {}; }
            store[args.save] = { at: new Date().toISOString(), config, embedding: args.embedding, metrics: base };
            writeFileSync(BASELINES, JSON.stringify(store, null, 2) + '\n');
            console.log(`已记基线 ${args.save} → ${BASELINES}`);
        }
        if (args.compare !== '') {
            const store = JSON.parse(readFileSync(BASELINES, 'utf8'));
            const saved = store[args.compare];
            if (!saved) throw new Error(`没有这条基线：${args.compare}`);
            console.log(`对比 ${args.compare}（${saved.metrics.chunks} 块）：R@5 ${pct(saved.metrics.recall5)} → ${pct(base.recall5)}  R@10 ${pct(saved.metrics.recall10)} → ${pct(base.recall10)}  MRR ${saved.metrics.mrr.toFixed(3)} → ${base.mrr.toFixed(3)}`);
        }
        if (args.json) console.log(JSON.stringify(results, null, 2));
    } finally {
        if (embedder?.manager) await embedder.manager.stop();
    }
}

main().catch((error) => {
    console.error(`评测失败：${error?.message ?? error}`);
    process.exitCode = 1;
});
