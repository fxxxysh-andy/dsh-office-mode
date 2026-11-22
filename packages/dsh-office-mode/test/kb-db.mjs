/**
 * `src/kb-db.js` 的验收测试。
 *
 * 只用 `node:test` + 系统临时目录：索引库是可重建的缓存，测试不该碰真实的 `.office/kb`。
 *
 * 本机没有 `node:sqlite`（Node 22 早期或没带 `--experimental-sqlite`）时**整组跳过**并打印
 * 原因 —— 那正是模块「动态 import + 优雅降级」要求存在的那条路，跳过而不是失败。
 *
 * 跑法：`node test/kb-db.mjs`
 */
import assert from 'node:assert/strict';
import { readdirSync } from 'node:fs';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test, { after } from 'node:test';
import { KB_DB_FILE, KB_DB_VERSION, kbDbPath, openKbDb } from '../src/kb-db.js';

const ROOT = await mkdtemp(join(tmpdir(), 'dsh-kb-db-'));
after(() => rm(ROOT, { recursive: true, force: true }));

let seq = 0;

/** 每个用例一个独立目录：索引库按目录分文件，互不影响。 */
async function fresh(t) {
    seq += 1;
    const dir = join(ROOT, `case-${seq}`, 'kb');
    const handle = await openKbDb({ dir });
    assert.equal(handle.ok, true, `openKbDb 失败：${handle.code} ${handle.error?.message ?? handle.error}`);
    assert.equal(handle.fts === true || handle.fts === false, true, 'handle.fts 必须是布尔');
    t.after(() => handle.close());
    return { handle, dir };
}

/** manifest 一行的最小形状（行里额外的字段要原样传给 loadChunks）。 */
function manifestRow(id, over = {}) {
    return {
        id,
        path: `${id}.md`,
        folder: '',
        title: id,
        hash: `hash-${id}`,
        bytes: 10,
        chars: 5,
        chunks: 1,
        tier: 'user',
        kind: 'doc',
        at: '2026-11-18T00:00:00.000Z',
        ...over,
    };
}

function chunk(id, n, text, over = {}) {
    return {
        id,
        n,
        text,
        head: '',
        outline: '',
        start: n * 10,
        end: n * 10 + 9,
        kind: 'text',
        image_path: '',
        summary_of: '',
        ...over,
    };
}

const probe = await openKbDb({ dir: join(ROOT, 'probe') });

if (!probe.ok && probe.code !== 'no-sqlite') {
    test('打开索引库（open-failed 不该被当成 skip）', () => {
        assert.fail(`openKbDb 在临时目录上失败：${probe.code} ${probe.error?.message ?? probe.error}`);
    });
} else if (!probe.ok) {
    test('kb 索引（跳过：本机没有 node:sqlite）', (t) => {
        console.log(`      本机 Node ${process.version} 没有可用的 node:sqlite：${probe.error?.message ?? probe.error}`);
        t.skip('node:sqlite 不可用，整个索引测试组跳过');
    });
} else {
    probe.close();

    test('openKbDb：建表、meta、能力位与路径', async (t) => {
        const { handle, dir } = await fresh(t);
        assert.equal(handle.path, kbDbPath(dir));
        assert.ok(handle.path.endsWith(KB_DB_FILE));

        const tables = new Set(handle.db.prepare("SELECT name FROM sqlite_master WHERE type = 'table'").all().map((row) => String(row.name)));
        for (const name of ['meta', 'docs', 'chunks', 'chunks_lex', 'embeddings', 'chunk_links', 'summaries']) {
            assert.ok(tables.has(name), `应该建出 ${name} 表`);
        }
        assert.equal(tables.has('chunks_fts'), handle.fts, 'chunks_fts 的存在要与 handle.fts 一致');

        const meta = new Map(handle.db.prepare('SELECT key, value FROM meta').all().map((row) => [String(row.key), String(row.value)]));
        assert.equal(meta.get('schema_version'), String(KB_DB_VERSION));
        assert.ok(meta.get('created_at'), 'meta 里要有 created_at');

        assert.deepEqual(handle.stats(), { docs: 0, chunks: 0, embeddings: 0, links: 0, summaries: 0, fts: handle.fts, bytes: 0 });
        assert.deepEqual(handle.searchLexical({ query: '记忆' }), []);
        assert.deepEqual(handle.searchVector({ vector: [1, 0] }), []);
    });

    test('syncDocuments：新增 / 跳过 / 改哈希 / 删文档 / origin 隔离', async (t) => {
        const { handle } = await fresh(t);
        const rows = [manifestRow('kb:a', { folder: '参考', chunks: 2 }), manifestRow('kb:b', { folder: '参考/合同' })];
        const blocks = {
            'kb:a': [chunk('kb:a:1', 1, '记忆面板可以按文件夹查看知识库'), chunk('kb:a:2', 2, '第二块正文')],
            'kb:b': [chunk('kb:b:1', 1, '付款周期为三十天')],
        };
        const load = (row) => blocks[row.id] ?? [];

        const first = await handle.syncDocuments(rows, { origin: 'workspace', loadChunks: async (row) => load(row) });
        assert.deepEqual(first, { added: 2, updated: 0, removed: 0, skipped: 0, chunks: 3 });
        assert.deepEqual(
            { docs: handle.stats().docs, chunks: handle.stats().chunks },
            { docs: 2, chunks: 3 },
        );

        // hash 没变：一个块文件都不该再读（loadChunks 直接抛也能过）。
        const second = await handle.syncDocuments(rows, {
            origin: 'workspace',
            loadChunks: async () => { throw new Error('hash 未变时不该再读块'); },
        });
        assert.deepEqual(second, { added: 0, updated: 0, removed: 0, skipped: 2, chunks: 0 });

        // 只改 folder（知识库里的「移动」不改内容哈希）：docs 行要跟着走，块不动。
        const moved = rows.map((row) => (row.id === 'kb:a' ? { ...row, folder: '归档' } : row));
        const third = await handle.syncDocuments(moved, { origin: 'workspace', loadChunks: async () => [] });
        assert.deepEqual(third, { added: 0, updated: 0, removed: 0, skipped: 2, chunks: 0 });
        assert.equal(handle.db.prepare('SELECT folder FROM docs WHERE id = ?').get('kb:a').folder, '归档');

        // 改哈希：整份重写，旧块与它的倒排都要清掉。
        const changed = [{ ...moved[0], hash: 'hash-kb:a-v2', chunks: 1 }, moved[1]];
        const nextBlocks = { 'kb:a': [chunk('kb:a-v2:1', 1, '换了一份正文')], 'kb:b': blocks['kb:b'] };
        const fourth = await handle.syncDocuments(changed, { origin: 'workspace', loadChunks: (row) => nextBlocks[row.id] ?? [] });
        assert.deepEqual(fourth, { added: 0, updated: 1, removed: 0, skipped: 1, chunks: 1 });
        assert.equal(handle.stats().chunks, 2, '旧的两块要删掉，只剩新的 a 一块 + b 一块');
        assert.deepEqual(handle.searchLexical({ query: '记忆 面板' }), [], '旧正文不该还能被搜到');
        assert.equal(handle.searchLexical({ query: '换了' }).length, 1);

        // 删文档。
        const fifth = await handle.syncDocuments([changed[0]], { origin: 'workspace', loadChunks: () => [] });
        assert.deepEqual(fifth, { added: 0, updated: 0, removed: 1, skipped: 1, chunks: 0 });
        assert.equal(handle.stats().docs, 1);
        assert.deepEqual(handle.searchLexical({ query: '付款 周期' }), []);

        // origin 隔离：同步 workspace 不该动到 global 的文档。
        await handle.syncDocuments([manifestRow('kb:g', { hash: 'hash-kb:g' })], {
            origin: 'global',
            loadChunks: () => [chunk('kb:g:1', 1, '全局层的文档')],
        });
        assert.equal(handle.stats().docs, 2);
        await handle.syncDocuments([], { origin: 'workspace' });
        assert.equal(handle.stats().docs, 1, '只该删掉 workspace 的那份');
        assert.equal(handle.db.prepare("SELECT origin FROM docs WHERE id = 'kb:g'").get().origin, 'global');
    });

    test('词法检索：中文 bigram、块头、filters、特殊字符', async (t) => {
        const { handle } = await fresh(t);
        const rows = [
            manifestRow('kb:mem', { folder: '', tier: 'user' }),
            manifestRow('kb:contract', { folder: '参考/合同', tier: 'verified' }),
            manifestRow('kb:other', { folder: '参考', tier: 'unverified' }),
            manifestRow('kb:sub', { folder: '参考/合同', tier: 'user' }),
        ];
        const blocks = {
            'kb:mem': [
                chunk('kb:mem:1', 1, '记忆面板可以按文件夹查看知识库', { head: '说明' }),
                chunk('kb:mem:2', 2, '完全无关的段落', { head: '记忆面板 · 合同' }),
            ],
            'kb:contract': [chunk('kb:contract:1', 1, '付款周期为三十天', { head: '合同' })],
            'kb:other': [chunk('kb:other:1', 1, '记忆面板的另一种说法')],
            'kb:sub': [chunk('kb:sub:1', 1, '记忆面板里的合同条款')],
        };
        await handle.syncDocuments(rows, { origin: 'workspace', loadChunks: (row) => blocks[row.id] ?? [] });

        // 「记忆 面板」两个词分开写，也要命中正文里连着的「记忆面板」。
        const ids = (hits) => hits.map((hit) => hit.chunkId).sort();
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆 面板' })), ['kb:mem:1', 'kb:mem:2', 'kb:other:1', 'kb:sub:1']);
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆面板' })), ['kb:mem:1', 'kb:mem:2', 'kb:other:1', 'kb:sub:1']);

        // 块头也要进检索：kb:mem:2 的正文里没有「合同」，只有块头里有。
        assert.ok(ids(handle.searchLexical({ query: '合同' })).includes('kb:mem:2'));

        // filters：folder 递归 / 只算直接子层 / tier / docId / origin。
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆', filters: { folder: '参考' } })), ['kb:other:1', 'kb:sub:1']);
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆', filters: { folder: '参考', recursive: false } })), ['kb:other:1']);
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆', filters: { tier: 'user' } })), ['kb:mem:1', 'kb:mem:2', 'kb:sub:1']);
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆', filters: { docId: 'kb:mem' } })), ['kb:mem:1', 'kb:mem:2']);
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆', filters: { origin: 'global' } })), []);

        // 特殊字符不能把 FTS 查询语法弄崩（引号 / 冒号 / 括号）。
        const special = handle.searchLexical({ query: '记忆"面板 (测试' });
        assert.ok(Array.isArray(special));
        assert.deepEqual(handle.searchLexical({ query: 'node:sqlite' }), []);

        assert.equal(handle.searchLexical({ query: '记忆', limit: 1 }).length, 1);
        assert.deepEqual(handle.searchLexical({ query: '   ' }), []);

        const hit = handle.searchLexical({ query: '记忆' })[0];
        assert.equal(typeof hit.score, 'number');
        assert.ok(Number.isFinite(hit.score));
        assert.ok(hit.docId.startsWith('kb:'));
        assert.equal(typeof hit.title, 'string');
        assert.equal(typeof hit.path, 'string');
    });

    test('词法检索：FTS 表坏掉时退回 LIKE，结果集不退化', async (t) => {
        const { handle } = await fresh(t);
        await handle.syncDocuments([manifestRow('kb:a'), manifestRow('kb:b')], {
            origin: 'workspace',
            loadChunks: (row) => (row.id === 'kb:a'
                ? [chunk('kb:a:1', 1, '记忆面板可以按文件夹查看', { head: '说明' })]
                : [chunk('kb:b:1', 1, '完全无关的段落', { head: '记忆面板' })]),
        });
        const ids = (hits) => hits.map((hit) => hit.chunkId).sort();
        const withFts = ids(handle.searchLexical({ query: '记忆 面板' }));
        assert.deepEqual(withFts, ['kb:a:1', 'kb:b:1']);

        // 真出现「FTS 表在、但这一轮查不动」时（换运行时 / 库损坏），查询要退到 LIKE
        // 而不是把异常抛给调用方 —— 索引是可重建的，检索不该因此整段失效。
        const hasFts = handle.db.prepare("SELECT name FROM sqlite_master WHERE name = 'chunks_fts'").get() !== undefined;
        if (hasFts) handle.db.exec('DROP TABLE chunks_fts'); // 本来是 LIKE 模式的机器上没有这张表
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆 面板' })), withFts);
        assert.deepEqual(ids(handle.searchLexical({ query: '记忆', filters: { docId: 'kb:a' } })), ['kb:a:1']);
    });

    test('向量：写入 / 读回 / 余弦排序 / filters', async (t) => {
        const { handle } = await fresh(t);
        await handle.syncDocuments([manifestRow('kb:a', { chunks: 3 }), manifestRow('kb:b')], {
            origin: 'workspace',
            loadChunks: (row) => (row.id === 'kb:a'
                ? [chunk('kb:a:1', 1, '甲'), chunk('kb:a:2', 2, '乙'), chunk('kb:a:3', 3, '丙')]
                : [chunk('kb:b:1', 1, '丁')]),
        });

        const stored = handle.upsertEmbeddings([
            { chunkId: 'kb:a:1', model: 'm1', vector: [1, 0, 0] },
            { chunkId: 'kb:a:2', model: 'm1', vector: [0.8, 0.6, 0] },
            { chunkId: 'kb:a:3', model: 'm1', vector: [0, 1, 0] },
            { chunkId: 'kb:b:1', model: 'm1', vector: new Float32Array([0.6, 0.8, 0]) },
            { chunkId: '', model: 'm1', vector: [1, 1, 1] },
        ]);
        assert.deepEqual(stored, { stored: 4 }, '空 chunkId 要跳过');
        assert.equal(handle.stats().embeddings, 4);

        const got = handle.embeddingFor('kb:a:2');
        assert.equal(got.model, 'm1');
        assert.equal(got.dim, 3);
        assert.ok(got.vector instanceof Float32Array);
        assert.deepEqual(Array.from(got.vector).map((n) => Math.round(n * 100) / 100), [0.8, 0.6, 0]);
        assert.equal(handle.embeddingFor('kb:missing'), null);

        const hits = handle.searchVector({ vector: [1, 0, 0], limit: 10 });
        assert.deepEqual(hits.map((hit) => hit.chunkId), ['kb:a:1', 'kb:a:2', 'kb:b:1', 'kb:a:3']);
        assert.ok(Math.abs(hits[0].score - 1) < 1e-6, '同向向量余弦应为 1');
        assert.ok(hits[0].score > hits[1].score && hits[1].score > hits[2].score, '要按余弦降序');
        assert.equal(hits[0].docId, 'kb:a');
        assert.ok(Math.abs(hits[2].score - 0.6) < 1e-6);

        const scoped = handle.searchVector({ vector: [1, 0, 0], filters: { docId: 'kb:b' } });
        assert.deepEqual(scoped.map((hit) => hit.chunkId), ['kb:b:1']);
        assert.deepEqual(handle.searchVector({ vector: [1, 0, 0], filters: { origin: 'global' } }), []);
        assert.deepEqual(handle.searchVector({ vector: [] }), []);

        // 同一个块重写：按 chunk_id 覆盖，不新增行。
        handle.upsertEmbeddings([{ chunkId: 'kb:a:2', model: 'm2', vector: [0, 0, 1] }]);
        assert.equal(handle.embeddingFor('kb:a:2').model, 'm2');
        assert.equal(handle.stats().embeddings, 4);
        assert.deepEqual(Array.from(handle.embeddingFor('kb:a:2').vector), [0, 0, 1]);

        assert.deepEqual(handle.allChunkIds({ filters: { docId: 'kb:a' } }), ['kb:a:1', 'kb:a:2', 'kb:a:3']);
        assert.equal(handle.allChunkIds({}).length, 4);
    });

    test('关系：link / linksFor / related 两跳 / unlink', async (t) => {
        const { handle } = await fresh(t);
        await handle.syncDocuments([manifestRow('kb:x', { chunks: 3, folder: '参考' })], {
            origin: 'workspace',
            loadChunks: () => [chunk('kb:x:1', 1, '第一块'), chunk('kb:x:2', 2, '第二块'), chunk('kb:x:3', 3, '第三块')],
        });

        const a = handle.link({ sourceId: 'kb:x:1', targetId: 'kb:x:2', note: '同源' });
        assert.ok(a.id.startsWith('kb-link:'));
        assert.equal(a.kind, 'related');
        assert.equal(a.note, '同源');
        const b = handle.link({ sourceId: 'kb:x:2', targetId: 'kb:x:3', kind: 'supports' });
        assert.equal(b.kind, 'supports');
        assert.equal(handle.stats().links, 2);

        // 同一三元组再 link：同一条边，笔记更新，不新增行。
        const again = handle.link({ sourceId: 'kb:x:1', targetId: 'kb:x:2', note: '改过的备注' });
        assert.equal(again.id, a.id);
        assert.equal(again.note, '改过的备注');
        assert.equal(handle.stats().links, 2);

        const out = handle.linksFor({ chunkId: 'kb:x:1', direction: 'out' });
        assert.equal(out.length, 1);
        assert.equal(out[0].neighborId, 'kb:x:2');
        assert.equal(out[0].note, '改过的备注');
        assert.deepEqual(handle.linksFor({ chunkId: 'kb:x:1', direction: 'in' }), []);
        assert.equal(handle.linksFor({ chunkId: 'kb:x:2' }).length, 2, '中间节点出入两条都要给');

        const two = handle.related({ chunkId: 'kb:x:1', depth: 2, limit: 10 });
        const byId = new Map(two.map((item) => [item.chunkId, item]));
        assert.equal(two.length, 2);
        assert.equal(byId.get('kb:x:2').hop, 1);
        assert.equal(byId.get('kb:x:3').hop, 2);
        assert.equal(byId.get('kb:x:2').title, 'kb:x');
        assert.equal(byId.get('kb:x:2').folder, '参考');
        assert.deepEqual(handle.related({ chunkId: 'kb:x:1', depth: 1 }).map((item) => item.chunkId), ['kb:x:2']);

        assert.deepEqual(handle.unlink({ sourceId: 'kb:x:2', targetId: 'kb:x:3' }), { removed: 1 });
        assert.deepEqual(handle.related({ chunkId: 'kb:x:1', depth: 2 }).map((item) => item.chunkId), ['kb:x:2']);
        assert.deepEqual(handle.unlink({ sourceId: 'kb:none', targetId: 'kb:none' }), { removed: 0 });
    });

    test('摘要：保存 / 列出 / 单条 / 文档不存在要抛', async (t) => {
        const { handle } = await fresh(t);
        await handle.syncDocuments([manifestRow('kb:a', { chunks: 2 })], {
            origin: 'workspace',
            loadChunks: () => [chunk('kb:a:1', 1, '第一块'), chunk('kb:a:2', 2, '第二块')],
        });

        const saved = handle.saveSummary({ title: '一段摘要', text: '摘要正文', docId: 'kb:a', chunkIds: ['kb:a:1', 'kb:a:2'] });
        assert.equal(saved.docId, 'kb:a');
        assert.ok(saved.id.startsWith('kb-sum:'));

        const list = handle.listSummaries({});
        assert.equal(list.length, 1);
        assert.equal(list[0].title, '一段摘要');
        assert.deepEqual(list[0].chunkIds, ['kb:a:1', 'kb:a:2']);
        assert.equal(list[0].docId, 'kb:a');

        const one = handle.summaryFor(saved.id);
        assert.equal(one.text, '摘要正文');
        assert.deepEqual(one.chunkIds, ['kb:a:1', 'kb:a:2']);
        assert.equal(handle.summaryFor('kb-sum:missing'), null);

        assert.throws(
            () => handle.saveSummary({ title: 'x', text: 'x', docId: 'kb:missing', chunkIds: [] }),
            /不在索引里/,
            '文档不存在时必须抛',
        );

        // 指定 id 再存：覆盖同一行，不新增。
        handle.saveSummary({ id: saved.id, title: '改过的摘要', text: '新正文', docId: 'kb:a', chunkIds: ['kb:a:1'] });
        assert.equal(handle.listSummaries({}).length, 1);
        assert.equal(handle.summaryFor(saved.id).title, '改过的摘要');
        assert.deepEqual(handle.summaryFor(saved.id).chunkIds, ['kb:a:1']);
    });

    test('close 后重开：数据还在，旧 handle 不可再用', async () => {
        const dir = join(ROOT, 'reopen', 'kb');
        const first = await openKbDb({ dir });
        assert.equal(first.ok, true);
        await first.syncDocuments([manifestRow('kb:a')], {
            origin: 'workspace',
            loadChunks: () => [chunk('kb:a:1', 1, '记忆面板在重开之后还在')],
        });
        first.upsertEmbeddings([{ chunkId: 'kb:a:1', model: 'm', vector: [0.1, 0.2, 0.3] }]);
        first.link({ sourceId: 'kb:a:1', targetId: 'kb:elsewhere' });
        first.saveSummary({ title: '摘要', text: '正文', docId: 'kb:a', chunkIds: ['kb:a:1'] });
        const before = first.stats();
        first.close();
        assert.throws(() => first.stats(), /已关闭/, 'close 之后调用方法要明确报错');
        first.close(); // 幂等

        const second = await openKbDb({ dir });
        assert.equal(second.ok, true);
        assert.deepEqual(second.stats(), before);
        assert.equal(second.searchLexical({ query: '记忆 面板' }).length, 1);
        assert.deepEqual(Array.from(second.embeddingFor('kb:a:1').vector).map((n) => Math.round(n * 10) / 10), [0.1, 0.2, 0.3]);
        assert.equal(second.linksFor({ chunkId: 'kb:a:1' }).length, 1);
        assert.equal(second.listSummaries({}).length, 1);

        // 只读 handle：能查、不能写（写会抛 SQLITE_READONLY）。
        const reader = await openKbDb({ dir, readonly: true });
        assert.equal(reader.ok, true);
        assert.equal(reader.stats().docs, 1);
        assert.equal(reader.searchLexical({ query: '重开' }).length, 1);
        assert.throws(() => reader.upsertEmbeddings([{ chunkId: 'kb:a:1', model: 'm', vector: [1] }]));
        reader.close();
        second.close();
    });

    test('ANN：图候选与精确扫描同序、落盘复用、删除与过滤补足', async (t) => {
        const { handle, dir } = await fresh(t);
        const total = 120;
        const rows = [manifestRow('kb:n', { chunks: total })];
        const blocks = Array.from({ length: total }, (_, i) => chunk(`kb:n:${i + 1}`, i + 1, `第 ${i + 1} 块`));
        await handle.syncDocuments(rows, { origin: 'workspace', loadChunks: () => blocks });

        // 固定伪随机向量：同一个种子每次跑都一样，不依赖 Math.random。
        let seed = 0x5eed1e55;
        const random = () => {
            seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
            return seed / 4294967296;
        };
        const items = [];
        for (let i = 0; i < total; i += 1) {
            const raw = Array.from({ length: 16 }, () => random() * 2 - 1);
            const norm = Math.hypot(...raw);
            items.push({ chunkId: `kb:n:${i + 1}`, model: 'm1', vector: raw.map((x) => x / norm) });
        }
        handle.upsertEmbeddings(items);

        const query = items[7].vector;
        assert.equal(
            handle.searchVector({ vector: query, limit: 5, ann: { enabled: false } })[0].chunkId,
            'kb:n:8',
            '查询向量就是第 8 块自己的向量，精确扫描第一名必须是它',
        );

        const approx = handle.searchVector({ vector: query, limit: 10, ann: { enabled: true, efSearch: 64 } });
        const exact = handle.searchVector({ vector: query, limit: 10, ann: { enabled: false } });
        assert.equal(approx[0].chunkId, 'kb:n:8');
        const overlap = approx.filter((hit) => exact.some((other) => other.chunkId === hit.chunkId)).length;
        assert.ok(overlap >= 9, `ANN 前 10 与精确扫描应基本重合，实际 ${overlap}/10`);

        const status = handle.annStatus();
        assert.equal(status.enabled, true);
        assert.equal(status.active, total);
        assert.equal(status.model, 'm1');
        assert.equal(status.dim, 16);
        assert.equal(status.tombstones, 0);
        const files = readdirSync(join(dir, 'index')).filter((name) => name.endsWith('.bin'));
        assert.equal(files.length, 1, '索引目录里应有一份 ann-<model>-<dim>.bin');

        // 过滤：图只给候选，范围条件在精确重算时生效；前 5 全是同文档，直接命中。
        const filtered = handle.searchVector({ vector: query, limit: 5, filters: { docId: 'kb:n' }, ann: { enabled: true, efSearch: 32 } });
        assert.equal(filtered.length, 5);
        assert.ok(filtered.every((hit) => hit.docId === 'kb:n'));

        // 关库 → 重开：manifest 计数一致时应复用文件，而不是当场重建。
        handle.close();
        const reopened = await openKbDb({ dir });
        assert.equal(reopened.ok, true);
        const again = reopened.searchVector({ vector: query, limit: 10, ann: { enabled: true, efSearch: 64 } });
        assert.equal(again[0].chunkId, 'kb:n:8');
        assert.equal(reopened.annStatus().active, total);
        assert.equal(reopened.annStatus().tombstones, 0);

        // 删除整篇文档：块行、向量、图里的节点一起消失，检索不再返回旧块。
        await reopened.syncDocuments([], { origin: 'workspace', loadChunks: () => [] });
        assert.deepEqual(reopened.searchVector({ vector: query, limit: 5, ann: { enabled: true } }), []);
        reopened.close();
    });
}
