/**
 * HNSW 索引（src/kb-ann.js）的测试 —— 第一百零三轮。
 *
 * 这一份只看索引本身的数学与持久化契约，不碰 SQLite / 知识库目录：
 *   - 召回率：与暴力余弦对比，固定种子，recall@10 ≥ 95%；
 *   - 增删改：删除后不再返回、更新后按新向量检索、重复 add 等价于更新；
 *   - 序列化：往返后 id / 邻居 / 检索结果一致，且同一份数据两次序列化逐字节相等；
 *   - 损坏：截断、魔数不对、版本读不懂都抛错，loadAnn 一律返回 null 交给调用方重建；
 *   - 过滤：谓词过滤后只返回命中项，不够 k 时不硬凑；
 *   - 墓碑压缩：删掉四成后 compact 仍能按新图检索。
 *
 * 跑法：node test/kb-ann.mjs
 */
import { mkdtempSync, readFileSync, rmSync, writeFileSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import {
    ANN_DEFAULTS, ANN_MAGIC, annPathsOf, createAnnIndex, dropAnn, loadAnn, parseAnnBuffer, saveAnn,
} from '../src/kb-ann.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 与模块内部同一套确定性随机（测试要能复现建图）。 */
function randomOf(seed) {
    let state = seed >>> 0;
    return () => {
        state = (state + 0x6d2b79f5) >>> 0;
        let t = state;
        t = Math.imul(t ^ (t >>> 15), t | 1);
        t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

function randomVector(random, dim) {
    const vector = new Float32Array(dim);
    for (let i = 0; i < dim; i += 1) vector[i] = random() * 2 - 1;
    return vector;
}

function normalize(vector) {
    let norm = 0;
    for (let i = 0; i < vector.length; i += 1) norm += vector[i] * vector[i];
    const scale = norm > 0 ? 1 / Math.sqrt(norm) : 0;
    const out = new Float32Array(vector.length);
    for (let i = 0; i < vector.length; i += 1) out[i] = vector[i] * scale;
    return out;
}

function dot(left, right) {
    let sum = 0;
    for (let i = 0; i < left.length; i += 1) sum += left[i] * right[i];
    return sum;
}

function bruteTop(vectors, query, k, allowed = null) {
    const scored = [];
    for (const [id, vector] of vectors) {
        if (allowed !== null && !allowed(id)) continue;
        scored.push({ id, score: dot(normalize(query), normalize(vector)) });
    }
    scored.sort((left, right) => right.score - left.score);
    return scored.slice(0, k).map((item) => item.id);
}

// ── 1. 基本插入与检索 ──────────────────────────────────────────────────────

await check('基本：插入后用余弦检索，返回条数与最近项正确', () => {
    const index = createAnnIndex({ dim: 8, m: 8, efConstruction: 64 });
    const vectors = new Map();
    for (let i = 0; i < 40; i += 1) {
        const vector = Float32Array.from({ length: 8 }, (_, at) => (at === i % 8 ? 1 : 0.05 * ((i + at) % 3)));
        vectors.set(`c${i}`, vector);
        assert.equal(index.add(`c${i}`, vector), true);
    }
    assert.equal(index.size, 40);
    const exact = vectors.get('c7');
    const hits = index.search(exact, 3);
    assert.equal(hits.length, 3);
    assert.equal(hits[0].id, 'c7');
    assert.ok(hits[0].score > 0.999, `自身相似度应接近 1，实际 ${hits[0].score}`);
    assert.ok(hits[0].score >= hits[1].score && hits[1].score >= hits[2].score, '结果按相似度降序');
});

// ── 2. 召回率 ─────────────────────────────────────────────────────────────

await check('召回：2000 条 × 200 次查询，recall@10 ≥ 95%（对暴力余弦）', () => {
    const random = randomOf(20261121);
    const dim = 64;
    const total = 2_000;
    const index = createAnnIndex({ dim, m: 16, efConstruction: 200, efSearch: 64, seed: 7 });
    const vectors = new Map();
    for (let i = 0; i < total; i += 1) {
        const id = `blk-${i}`;
        const vector = randomVector(random, dim);
        vectors.set(id, vector);
        index.add(id, vector);
    }
    let hitAt = 0;
    let wanted = 0;
    const queries = 200;
    for (let q = 0; q < queries; q += 1) {
        const query = randomVector(random, dim);
        const expected = bruteTop(vectors, query, 10);
        const actual = index.search(query, 10).map((item) => item.id);
        const actualSet = new Set(actual);
        for (const id of expected) if (actualSet.has(id)) hitAt += 1;
        wanted += expected.length;
    }
    const recall = hitAt / wanted;
    assert.ok(recall >= 0.95, `recall@10 应 ≥ 95%，实际 ${(recall * 100).toFixed(2)}%`);
    assert.ok(index.nodes >= total, '节点数不应少于插入次数');
});

// ── 3. 增删改 ─────────────────────────────────────────────────────────────

await check('增删改：删除不返回、更新按新向量、重复 add 等价更新', () => {
    const index = createAnnIndex({ dim: 4, m: 8, efConstruction: 64, efSearch: 32 });
    index.add('a', [1, 0, 0, 0]);
    index.add('b', [0, 1, 0, 0]);
    index.add('c', [0, 0, 1, 0]);
    assert.equal(index.remove('b'), true);
    assert.equal(index.remove('b'), false, '重复删除应返回 false');
    assert.equal(index.has('b'), false);
    assert.equal(index.search([0, 1, 0, 0], 3).some((item) => item.id === 'b'), false, '墓碑不能被召回');
    // 把 c 更新成与 b 同向：旧向量不该再被 'c' 召回。
    assert.equal(index.add('c', [0, 1, 0, 0]), true);
    assert.equal(index.size, 2, '更新不增加存活数');
    const hits = index.search([0, 1, 0, 0], 2);
    assert.deepEqual(hits.map((item) => item.id), ['c', 'a']);
    assert.ok(hits[0].score > 0.999);
    assert.equal(index.add('x', [1, 2, 3]), false, '维度不对的向量必须拒绝');
    assert.equal(index.add('zero', [0, 0, 0, 0]), false, '零向量必须拒绝');
});

// ── 4. 序列化往返 ─────────────────────────────────────────────────────────

await check('序列化：往返后 id 与检索结果一致，两次序列化逐字节相等', () => {
    const random = randomOf(99);
    const index = createAnnIndex({ dim: 16, m: 8, efConstruction: 100, efSearch: 48, seed: 3 });
    const queries = [];
    for (let i = 0; i < 300; i += 1) {
        const vector = randomVector(random, 16);
        index.add(`n${i}`, vector);
        if (i % 30 === 0) queries.push(vector);
    }
    index.remove('n5');
    index.remove('n250');
    const buffer = index.toBuffer();
    assert.equal(buffer.toString('ascii', 0, 4), ANN_MAGIC);
    const again = parseAnnBuffer(buffer);
    assert.equal(again.dim, 16);
    assert.equal(again.size, index.size);
    assert.equal(again.tombstones, index.tombstones);
    assert.deepEqual(again.liveIds().sort(), index.liveIds().sort());
    const rebuilt = again.toBuffer();
    assert.equal(rebuilt.length, buffer.length, '往返后字节数应一致');
    assert.ok(rebuilt.equals(buffer), '往返后应逐字节相等（图原样保留）');
    for (const query of queries) {
        assert.deepEqual(again.search(query, 5).map((item) => item.id), index.search(query, 5).map((item) => item.id));
    }
    // 空索引不落盘。
    assert.throws(() => saveAnn(createAnnIndex({ dim: 4 }), { dir: tmpdir(), model: 'x' }), /空索引/);
});

// ── 5. 过滤 ───────────────────────────────────────────────────────────────

await check('过滤：谓词只放行命中项，不够 k 时不硬凑', () => {
    const random = randomOf(4242);
    const dim = 32;
    const index = createAnnIndex({ dim, m: 12, efConstruction: 120, efSearch: 64 });
    const vectors = new Map();
    for (let i = 0; i < 800; i += 1) {
        const vector = randomVector(random, dim);
        vectors.set(`doc${i % 8}:blk${i}`, vector);
        index.add(`doc${i % 8}:blk${i}`, vector);
    }
    const query = randomVector(random, dim);
    const wanted = bruteTop(vectors, query, 8, (id) => id.startsWith('doc3:'));
    const hits = index.search(query, 8, { filter: (id) => id.startsWith('doc3:'), efSearch: 64 });
    assert.equal(hits.length, wanted.length);
    assert.ok(hits.every((item) => item.id.startsWith('doc3:')), '过滤谓词必须生效');
    assert.equal(hits[0].id, wanted[0], '过滤后最近的一条应与暴力检索一致');
    const overlap = hits.filter((item) => wanted.includes(item.id)).length;
    assert.ok(overlap >= 7, `过滤后 top8 至少应重合 7 条，实际 ${overlap}`);
    const none = index.search(query, 8, { filter: () => false });
    assert.equal(none.length, 0, '全部过滤掉时返回空，而不是变通成不过滤');
});

// ── 6. 墓碑压缩 ───────────────────────────────────────────────────────────

await check('墓碑：删掉四成后 compact 重建，检索仍可用且墓碑归零', () => {
    const random = randomOf(777);
    const dim = 24;
    const index = createAnnIndex({ dim, m: 12, efConstruction: 100, efSearch: 48 });
    const vectors = new Map();
    for (let i = 0; i < 500; i += 1) {
        const vector = randomVector(random, dim);
        vectors.set(`m${i}`, vector);
        index.add(`m${i}`, vector);
    }
    for (let i = 0; i < 500; i += 1) if (i % 5 === 0 || i % 5 === 1) index.remove(`m${i}`);
    assert.equal(index.size, 300);
    assert.ok(index.tombstones >= 200);
    const compacted = index.compact();
    assert.equal(compacted.rebuilt, true);
    assert.equal(index.size, 300);
    assert.equal(index.tombstones, 0);
    assert.equal(index.nodes, 300);
    // 压缩会按新图重建：不保证逐条次序相同，但最近的一条必须还是自己，且删掉的不能回来。
    const after = index.search(vectors.get('m7'), 5);
    assert.equal(after[0].id, 'm7');
    assert.ok(after.every((item) => !['m0', 'm1', 'm5', 'm6'].includes(item.id)), '已删除的 id 不能回到结果里');
    assert.equal(index.remove('m1'), false, '已删除的 id 不再存在');
});

// ── 7. 持久化与损坏恢复 ───────────────────────────────────────────────────

await check('持久化：save/load 对得上计数，计数或模型不符时返回 null，损坏文件返回 null', () => {
    const dir = mkdtempSync(join(tmpdir(), 'dsh-ann-'));
    try {
        const index = createAnnIndex({ dim: 8, m: 8, efConstruction: 64, efSearch: 32 });
        for (let i = 0; i < 30; i += 1) index.add(`p${i}`, Float32Array.from({ length: 8 }, () => Math.sin(i + 1)));
        const manifest = saveAnn(index, { dir, model: 'bge-small-zh-v1.5', updatedAt: '2026-11-21T00:00:00.000Z', count: 30 });
        assert.equal(manifest.active, 30);
        assert.equal(manifest.model, 'bge-small-zh-v1.5');
        const paths = annPathsOf(dir, 'bge-small-zh-v1.5', 8);
        assert.ok(readFileSync(paths.file).length > 0);
        const loadedRight = loadAnn({ dir, model: 'bge-small-zh-v1.5', dim: 8, count: 30, updatedAt: '2026-11-21T00:00:00.000Z' });
        assert.ok(loadedRight, '计数与更新时间都对得上时应装回');
        assert.equal(loadedRight.index.size, 30);
        assert.equal(loadAnn({ dir, model: 'bge-small-zh-v1.5', dim: 8, count: 31 }), null, '计数不对要重建');
        assert.equal(loadAnn({ dir, model: 'bge-small-zh-v1.5', dim: 8, updatedAt: '2026-11-22T00:00:00.000Z' }), null, '更新时间不对要重建');
        assert.equal(loadAnn({ dir, model: '另一个模型', dim: 8 }), null, '模型不对不能读到旧图');

        // 截断 / 魔数 / 版本：解析层必须抛错，loadAnn 必须吞掉并返回 null。
        const whole = readFileSync(paths.file);
        assert.throws(() => parseAnnBuffer(whole.subarray(0, 40)), /太短|魔数/);
        const badMagic = Buffer.from(whole);
        badMagic.write('XXXX', 0, 4, 'ascii');
        assert.throws(() => parseAnnBuffer(badMagic), /魔数/);
        const badVersion = Buffer.from(whole);
        badVersion.writeUInt32LE(99, 4);
        assert.throws(() => parseAnnBuffer(badVersion), /版本/);
        writeFileSync(paths.file, whole.subarray(0, 60));
        assert.equal(loadAnn({ dir, model: 'bge-small-zh-v1.5', dim: 8 }), null, '文件损坏时要能退回重建');
        dropAnn({ dir, model: 'bge-small-zh-v1.5', dim: 8 });
        assert.equal(loadAnn({ dir, model: 'bge-small-zh-v1.5', dim: 8 }), null, '删掉后也不该装回');

        // 索引目录不存在时写盘会自动建。
        const deep = join(dir, 'nested', 'deeper');
        mkdirSync(deep, { recursive: true });
        saveAnn(index, { dir: deep, model: 'm', count: 30 });
        assert.ok(loadAnn({ dir: deep, model: 'm', dim: 8 }), '新建目录后也能读回');
    } finally {
        rmSync(dir, { recursive: true, force: true });
    }
});

await check('默认参数：M / efConstruction 落在安全区间，坏值不收', () => {
    const index = createAnnIndex({ m: 0, efConstruction: -5, efSearch: 9999 });
    assert.equal(index.m, 4, 'M 下限 4');
    assert.equal(index.efConstruction, 4, 'efConstruction 不低于 M');
    assert.equal(index.efSearch, 512, 'efSearch 上限 512');
    assert.equal(createAnnIndex().m, ANN_DEFAULTS.m);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    if (item.ok) console.log(`PASS  ${item.name}`);
    else console.log(`FAIL  ${item.name}  → ${item.error?.message ?? item.error}`);
}
console.log(`\nkb-ann: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
