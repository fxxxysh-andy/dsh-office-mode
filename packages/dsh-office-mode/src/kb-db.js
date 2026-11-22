/**
 * 知识库的 SQLite 全局索引位（`kb.sqlite`）—— **可重建的索引**，不是真源。
 *
 * 真源仍然是 `.office/kb/manifest.jsonl` + `chunks/<docHash>/<n>.json`：这一层只负责
 * 「跨文档、跨来源」的检索（词法 + 向量 + 关系 + 摘要），任何时候都能从真源整份重建。
 * 所以这里没有迁移脚本，只有 `meta.schema_version`；版本对不上就重建，不需要把索引
 * 当成要小心伺候的数据。
 *
 * 三条实现纪律（都踩过，别再改回去）：
 *
 *   1. **顶层不许 import `node:sqlite`**。它到 Node 22.5 才有，而且早期版本还要
 *      `--experimental-sqlite`；本插件 CI 跑 Node 22 —— 顶层 import 会让整个插件在
 *      import 阶段就崩，连「降级」都来不及做。所以只在 `openKbDb()` 里
 *      `await import()`，拿不到就返回 `{ ok:false, code:'no-sqlite', error }`。
 *
 *   2. **FTS5 也当成可选能力**。Node 自带的那份 sqlite 有没有编进 FTS5，取决于版本与
 *      发行版（本机 Node 26 有、Node 22 的某些构建没有）。拿不到就退普通表
 *      `chunks_lex` + LIKE；两张表**一直同时维护**（FTS 表可用时也写一份词袋），
 *      这样降级与升级两个方向都不用重建数据。降级期间写过的库会被 `meta.fts_dirty`
 *      标脏，下次拿到 FTS5 时按它整表重建。
 *
 *   3. **中文口径由本插件掌握**。FTS5 的 unicode61 对中文基本失效（一整段算一个词），
 *      所以入库的是 `kbTokens()` 切好的 bigram 空格串，查询侧也把每个 token 用双引号
 *      包起来再 OR —— FTS5 只做它擅长的倒排与 bm25，切词与权重始终是同一份实现。
 *
 * 调用侧统一 `await`（方法本身同步或 async 都可以，但不要用回调）。
 *
 * @module dsh-office-mode/kb-db
 */
import { createHash } from 'node:crypto';
import { mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { ANN_DEFAULTS, createAnnIndex, loadAnn, saveAnn } from './kb-ann.js';
import { kbTokens, kbTierOf, normalizeKbFolder } from './kb.js';

/** 索引库文件名（与 manifest.jsonl、chunks/ 同一个目录）。 */
export const KB_DB_FILE = 'kb.sqlite';
/** 表结构版本；对不上就整份重建（索引本来就可重建，不需要逐版迁移）。 */
export const KB_DB_VERSION = 1;

/** 一次词法检索最多用多少个查询词：超长查询不该把 FTS 查询串撑爆。 */
const MAX_TERMS = 64;
/** 一次向量检索最多扫多少块（精确余弦的语料上界，超出的部分不参与打分）。 */
const MAX_VECTOR_SCAN = 50_000;
/**
 * ANN 索引最多覆盖多少个块：超过就退回精确余弦。
 *
 * 首次检索时索引是**同步**建的，建图成本 O(n log n)；两万块在纯 JS 里已经是
 * 几十秒级，再大就会让一次检索看起来像卡死。这个数是实测后定的（见文档评测表），
 * 不是拍脑袋：超过就不建，检索仍走精确扫描，只是慢一点、并不出错。
 */
export const ANN_MAX_BUILD = 20_000;
/** ANN 候选池相对返回条数的倍数：宽一点，过滤与精确重算之后才还有货。 */
const ANN_CANDIDATE_FACTOR = 6;
/** 一条 `IN (...)` 最多带多少个绑定参数（老版本 SQLite 的变量上限是 999）。 */
const IN_BATCH = 400;
/** 一次检索最多返回多少条。 */
const MAX_LIMIT = 200;
/** 关系图最多走几跳：再深就不是「相关」而是「全图」了。 */
const MAX_DEPTH = 5;

/** 基础表。列名与 `SPEC` 一致；`start` / `end` 双引号包起来（END 是 SQL 关键字）。 */
const SCHEMA_DDL = `
CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS docs (
    id TEXT PRIMARY KEY,
    path TEXT NOT NULL,
    folder TEXT NOT NULL DEFAULT '',
    title TEXT NOT NULL,
    hash TEXT NOT NULL,
    origin TEXT NOT NULL,
    tier TEXT NOT NULL DEFAULT 'user',
    kind TEXT NOT NULL DEFAULT 'text',
    bytes INTEGER NOT NULL DEFAULT 0,
    chars INTEGER NOT NULL DEFAULT 0,
    chunks INTEGER NOT NULL DEFAULT 0,
    at TEXT NOT NULL DEFAULT '',
    text TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS docs_origin ON docs(origin);
CREATE INDEX IF NOT EXISTS docs_hash ON docs(hash);
CREATE TABLE IF NOT EXISTS chunks (
    id TEXT PRIMARY KEY,
    doc_id TEXT NOT NULL,
    n INTEGER NOT NULL,
    text TEXT NOT NULL,
    head TEXT NOT NULL DEFAULT '',
    outline TEXT NOT NULL DEFAULT '',
    "start" INTEGER,
    "end" INTEGER,
    kind TEXT NOT NULL DEFAULT 'text',
    image_path TEXT NOT NULL DEFAULT '',
    summary_of TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS chunks_doc_id ON chunks(doc_id);
CREATE TABLE IF NOT EXISTS chunks_lex (
    chunk_id TEXT PRIMARY KEY,
    lex TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS embeddings (
    chunk_id TEXT PRIMARY KEY,
    model TEXT NOT NULL,
    dim INTEGER NOT NULL,
    vector BLOB NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS chunk_links (
    id TEXT PRIMARY KEY,
    source_id TEXT NOT NULL,
    target_id TEXT NOT NULL,
    kind TEXT NOT NULL DEFAULT 'related',
    note TEXT NOT NULL DEFAULT '',
    at TEXT NOT NULL,
    UNIQUE(source_id, target_id, kind)
);
CREATE INDEX IF NOT EXISTS chunk_links_source ON chunk_links(source_id);
CREATE INDEX IF NOT EXISTS chunk_links_target ON chunk_links(target_id);
CREATE TABLE IF NOT EXISTS summaries (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    text TEXT NOT NULL,
    doc_id TEXT NOT NULL,
    chunk_ids TEXT NOT NULL,
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS summaries_doc_id ON summaries(doc_id);
`;

/** FTS5 表：`chunk_id` 只存不索引（它由 chunks 表带出来），两个词袋列都要能打分。 */
const FTS_DDL = "CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(chunk_id UNINDEXED, lex, head_lex, tokenize='unicode61')";

// ── 小工具 ──────────────────────────────────────────────────────────────────

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

function nowIso() {
    return new Date().toISOString();
}

function toInt(value, fallback) {
    const n = typeof value === 'number' ? value : Number.parseInt(String(value ?? ''), 10);
    return Number.isFinite(n) ? n : fallback;
}

function clampInt(value, fallback, min, max) {
    const n = toInt(value, fallback);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(Math.max(n, min), max);
}

/** LIKE 里的元字符转义（配 `ESCAPE '\'` 用）。 */
function escapeLike(value) {
    return String(value).replace(/[\\%_]/g, (char) => `\\${char}`);
}

/** 按固定大小切批：一条 SQL 的绑定参数不该无限长。 */
function batches(list, size) {
    const out = [];
    for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
    return out;
}

function placeholders(count) {
    return new Array(count).fill('?').join(', ');
}

/** 库文件路径（`dir` 是 kb 目录，不是记忆目录）。 */
export function kbDbPath(dir) {
    return join(String(dir ?? ''), KB_DB_FILE);
}

// ── node:sqlite 的按需加载 ──────────────────────────────────────────────────
//
// 模块顶层绝不 import：Node 22.5 之前根本没有这个内置模块，顶层 import 会让插件
// 在加载阶段直接抛 ERR_UNKNOWN_BUILTIN_MODULE。这里缓存「加载结果」而不是 Promise，
// 失败也只失败一次，后面的调用直接拿同一个错误。

let sqliteState = null;

async function loadSqlite() {
    if (sqliteState !== null) return sqliteState;
    try {
        sqliteState = { ok: true, sqlite: await import('node:sqlite') };
    } catch (error) {
        sqliteState = { ok: false, error };
    }
    return sqliteState;
}

// ── 建表 / 能力探测 ─────────────────────────────────────────────────────────

function tableExists(db, name) {
    const row = db.prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name = ?").get(name);
    return row !== undefined && row !== null;
}

function metaGet(db, key) {
    const row = db.prepare('SELECT value FROM meta WHERE key = ?').get(key);
    return row === undefined || row === null ? null : String(row.value);
}

function metaSet(db, key, value) {
    db.prepare('INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value').run(key, String(value));
}

function countOf(db, sql) {
    const row = db.prepare(sql).get();
    return row === undefined || row === null ? 0 : Number(row.n ?? 0);
}

/** 用 chunks_lex 的词袋整表重建 FTS（降级期写过的库升级回来时走这里）。 */
function rebuildFts(db) {
    db.exec('BEGIN IMMEDIATE');
    try {
        db.exec('DELETE FROM chunks_fts');
        const insert = db.prepare('INSERT INTO chunks_fts (chunk_id, lex, head_lex) VALUES (?, ?, ?)');
        for (const row of db.prepare('SELECT chunk_id, lex FROM chunks_lex ORDER BY chunk_id').all()) {
            insert.run(String(row.chunk_id), String(row.lex ?? ''), '');
        }
        db.exec('COMMIT');
    } catch (error) {
        try { db.exec('ROLLBACK'); } catch { /* 回滚失败只能把原始错误抛出去 */ }
        throw error;
    }
}

/**
 * 建表并探测能力，返回 `fts`。
 *
 * 判断「要不要重建 FTS」看三件事：上次是不是降级写的（`fts` = '0'）、降级期有没有
 * 继续写过（`fts_dirty` = '1'）、以及行数与 chunks 对不对得上（崩在中间的情况）。
 */
function initSchema(db) {
    db.exec(SCHEMA_DDL);
    let fts = true;
    try {
        db.exec(FTS_DDL);
    } catch {
        fts = false;
    }
    const previous = metaGet(db, 'fts');
    const dirty = metaGet(db, 'fts_dirty') === '1';
    if (fts) {
        const sameSize = countOf(db, 'SELECT count(*) AS n FROM chunks_fts') === countOf(db, 'SELECT count(*) AS n FROM chunks');
        if (previous === '0' || dirty || !sameSize) rebuildFts(db);
        metaSet(db, 'fts', '1');
        metaSet(db, 'fts_dirty', '0');
    } else if (previous === '1' || tableExists(db, 'chunks_fts')) {
        // 降级运行时 FTS 表还留着，但这一轮写不进去：标脏，等能写的那天整表重建。
        metaSet(db, 'fts_dirty', '1');
    }
    metaSet(db, 'fts', fts ? '1' : '0');
    if (metaGet(db, 'created_at') === null) metaSet(db, 'created_at', nowIso());
    metaSet(db, 'schema_version', String(KB_DB_VERSION));
    return fts;
}

// ── 打开 ────────────────────────────────────────────────────────────────────

/**
 * 打开（必要时新建）知识库索引。
 *
 * **不抛错**：sqlite 不可用返回 `{ ok:false, code:'no-sqlite', error }`，
 * 路径 / 权限 / 库文件损坏返回 `{ ok:false, code:'open-failed', error }`。
 * 调用方据此决定「这一轮没有全局索引」还是真出事了 —— 两种情况都不该让插件崩。
 *
 * @param {{dir?: string, readonly?: boolean}} [options]
 * @returns {Promise<object>} handle 或降级回执
 */
export async function openKbDb({ dir, readonly = false } = {}) {
    const target = asText(dir);
    if (target === '') {
        return { ok: false, code: 'open-failed', error: new Error('openKbDb 需要 dir（知识库目录）') };
    }
    const loaded = await loadSqlite();
    if (!loaded.ok) return { ok: false, code: 'no-sqlite', error: loaded.error };

    const file = join(target, KB_DB_FILE);
    const readOnly = readonly === true;
    let db = null;
    try {
        if (!readOnly) mkdirSync(target, { recursive: true });
        db = new loaded.sqlite.DatabaseSync(file, readOnly ? { readOnly: true } : {});
        db.exec('PRAGMA busy_timeout = 5000');
        let fts;
        if (readOnly) {
            fts = tableExists(db, 'chunks_fts');
        } else {
            // WAL 让「读写同时进行」不至于互相等；某些文件系统不支持，退回默认模式即可。
            try { db.exec('PRAGMA journal_mode = WAL'); } catch { /* 不支持 WAL 不是错误 */ }
            fts = initSchema(db);
        }
        return makeHandle({
            db,
            path: file,
            fts,
            readonly: readOnly,
            // ANN 是**按需懒加载**的缓存：这里只占位，第一次向量检索才会读/建图。
            ann: null,
            annKey: '',
            annMeta: null,
            annDirty: false,
            annSkipped: '',
        });
    } catch (error) {
        try { db?.close(); } catch { /* 已经关了 / 打不开：没有更好的处理 */ }
        return { ok: false, code: 'open-failed', error };
    }
}

// ── handle ──────────────────────────────────────────────────────────────────

function assertOpen(ctx) {
    if (ctx.closed) throw new Error('知识库索引已关闭（close() 之后不能再调用）');
}

function makeHandle(ctx) {
    return {
        ok: true,
        path: ctx.path,
        db: ctx.db,
        fts: ctx.fts,
        readonly: ctx.readonly,
        close: () => closeHandle(ctx),
        syncDocuments: (rows, options) => syncDocuments(ctx, rows, options),
        searchLexical: (params) => searchLexical(ctx, params),
        upsertEmbeddings: (items) => upsertEmbeddings(ctx, items),
        embeddingFor: (chunkId) => embeddingFor(ctx, chunkId),
        searchVector: (params) => searchVector(ctx, params),
        annStatus: () => annStatusOf(ctx),
        link: (params) => link(ctx, params),
        unlink: (params) => unlink(ctx, params),
        linksFor: (params) => linksFor(ctx, params),
        related: (params) => related(ctx, params),
        saveSummary: (params) => saveSummary(ctx, params),
        listSummaries: (params) => listSummaries(ctx, params),
        countSummaries: () => countSummaries(ctx),
        searchSummaries: (params) => searchSummaries(ctx, params),
        summaryFor: (id) => summaryFor(ctx, id),
        chunksByIds: (params) => chunksByIds(ctx, params),
        stats: () => stats(ctx),
        allChunkIds: (params) => allChunkIds(ctx, params),
    };
}

function closeHandle(ctx) {
    if (ctx.closed) return;
    ctx.closed = true;
    flushAnn(ctx);
    try { ctx.db.close(); } catch { /* close 失败也不该把调用方带崩 */ }
}

// ── 文档同步（manifest → 索引） ─────────────────────────────────────────────

/** 规范化 manifest 行：坏行（没 id / 没 hash）返回 null，由调用方跳过。 */
function normalizeDocRow(raw) {
    if (raw === null || typeof raw !== 'object') return null;
    const id = asText(raw.id);
    const hash = asText(raw.hash);
    if (id === '' || hash === '') return null;
    let folder = '';
    try { folder = normalizeKbFolder(raw.folder); } catch { folder = ''; }
    const path = asText(raw.path);
    return {
        id,
        path,
        folder,
        title: asText(raw.title) || path || id,
        hash,
        tier: kbTierOf(raw.tier),
        kind: asText(raw.kind) || 'text',
        bytes: Math.max(0, toInt(raw.bytes, 0)),
        chars: Math.max(0, toInt(raw.chars, 0)),
        at: asText(raw.at),
        text: typeof raw.text === 'string' ? raw.text : '',
    };
}

/** 块数组规范化：块 id 缺了就用 `<docId>:<n>` 兜底（只为了有个稳定主键，不猜内容）。 */
function normalizeChunks(docId, raw) {
    const list = Array.isArray(raw) ? raw : [];
    const out = [];
    const seen = new Set();
    list.forEach((item, index) => {
        if (item === null || typeof item !== 'object') return;
        const n = toInt(item.n, index + 1);
        const id = asText(item.id) || `${docId}:${n}`;
        if (seen.has(id)) return;
        seen.add(id);
        const start = toInt(item.start, Number.NaN);
        const end = toInt(item.end, Number.NaN);
        out.push({
            id,
            n,
            text: typeof item.text === 'string' ? item.text : '',
            head: typeof item.head === 'string' ? item.head : '',
            outline: typeof item.outline === 'string' ? item.outline : '',
            start: Number.isFinite(start) ? start : null,
            end: Number.isFinite(end) ? end : null,
            kind: asText(item.kind) || 'text',
            image_path: asText(item.image_path),
            summary_of: asText(item.summary_of),
        });
    });
    return out;
}

/** 一块的两种词袋：FTS 要词频（保留重复），回退表只做 LIKE（去重足够）。 */
function lexicalOf(chunk) {
    const textLex = kbTokens(chunk.text).join(' ');
    const headLex = kbTokens(chunk.head).join(' ');
    const bag = [...new Set([...kbTokens(chunk.text), ...kbTokens(chunk.head)])].join(' ');
    return { textLex, headLex, bag };
}

function chunkIdsOfDoc(ctx, docId) {
    return ctx.db.prepare('SELECT id FROM chunks WHERE doc_id = ? ORDER BY n ASC').all(docId).map((row) => String(row.id));
}

function deleteChunkRows(ctx, ids) {
    if (ids.length === 0) return;
    annRemoveIds(ctx, ids);
    for (const batch of batches(ids, IN_BATCH)) {
        const marks = placeholders(batch.length);
        // FTS 表只在「这一轮能用 FTS5」时才存在；降级运行时它可能根本不在库里。
        if (ctx.fts) ctx.db.prepare(`DELETE FROM chunks_fts WHERE chunk_id IN (${marks})`).run(...batch);
        ctx.db.prepare(`DELETE FROM chunks_lex WHERE chunk_id IN (${marks})`).run(...batch);
        ctx.db.prepare(`DELETE FROM embeddings WHERE chunk_id IN (${marks})`).run(...batch);
        ctx.db.prepare(`DELETE FROM chunk_links WHERE source_id IN (${marks})`).run(...batch);
        ctx.db.prepare(`DELETE FROM chunk_links WHERE target_id IN (${marks})`).run(...batch);
    }
}

/** 同一份内容的兄弟文档（块 id 是内容寻址的，两份同内容文档共用一个 id 空间）。 */
function siblingOfHash(ctx, hash, exceptId) {
    if (hash === '') return null;
    const row = ctx.db.prepare('SELECT id FROM docs WHERE hash = ? AND id <> ? LIMIT 1').get(hash, exceptId);
    return row === undefined || row === null ? null : String(row.id);
}

/**
 * 一份文档「退出索引」时释放它的块。
 *
 * 这里有个只有内容寻址才会遇到的边角：两份**内容完全相同**的文档共用同一批块 id。
 * 所以不能无脑删 —— 有兄弟文档还拿着这个 hash 时，把块的归属转给它（正文没变），
 * 只有最后一个兄弟走了才真删。索引可重建，但「删 A 把 B 的块也删了」不该发生。
 */
function releaseDocChunks(ctx, docId, hash) {
    const sibling = siblingOfHash(ctx, hash, docId);
    if (sibling !== null) {
        ctx.db.prepare('UPDATE chunks SET doc_id = ? WHERE doc_id = ?').run(sibling, docId);
        return;
    }
    const ids = chunkIdsOfDoc(ctx, docId);
    deleteChunkRows(ctx, ids);
    ctx.db.prepare('DELETE FROM chunks WHERE doc_id = ?').run(docId);
}

/**
 * 按 manifest 行增量同步索引。
 *
 * `hash` 未变就跳过（不重写块）；变了的整份重写；`rows` 里没有的同一 `origin` 文档删掉。
 * `loadChunks` 在事务**之前**读完 —— 读写块文件是 async 的，而事务里只留同步写，
 * 免得把别处的 await 卷进一个开着的事务里。
 *
 * @param {object[]} rows manifest 行
 * @param {{origin?: string, loadChunks?: (row: object) => object[]|Promise<object[]>}} [options]
 * @returns {Promise<{added:number, updated:number, removed:number, skipped:number, chunks:number}>}
 */
async function syncDocuments(ctx, rows, options = {}) {
    assertOpen(ctx);
    const origin = asText(options?.origin);
    if (origin === '') throw new Error('syncDocuments 需要 origin（索引按来源库分区，空 origin 会把别人的文档删掉）');
    const loadChunks = typeof options?.loadChunks === 'function' ? options.loadChunks : () => [];

    // 原始行要留着：loadChunks 收到的是 manifest 那一行（它可能带调用方自己的字段）。
    const list = [];
    const seen = new Set();
    for (const raw of Array.isArray(rows) ? rows : []) {
        const row = normalizeDocRow(raw);
        if (row === null || seen.has(row.id)) continue;
        seen.add(row.id);
        list.push({ row, source: raw });
    }

    const existing = new Map();
    for (const row of ctx.db.prepare('SELECT id, hash, path, folder, title, tier, kind, bytes, chars, at FROM docs WHERE origin = ?').all(origin)) {
        existing.set(String(row.id), row);
    }

    const touched = [];
    const stale = [];
    let skipped = 0;
    for (const entry of list) {
        const { row } = entry;
        const before = existing.get(row.id);
        if (before !== undefined && String(before.hash) === row.hash) {
            skipped += 1;
            // 内容没变但元数据可能变了（知识库里的「移动文件夹」不改变内容哈希）：
            // 只更新 docs 行，不碰块，也不改 updated_at 之外的东西。
            const changed = String(before.path) !== row.path
                || String(before.folder) !== row.folder
                || String(before.title) !== row.title
                || String(before.tier) !== row.tier
                || String(before.kind) !== row.kind
                || Number(before.bytes) !== row.bytes
                || Number(before.chars) !== row.chars
                || String(before.at) !== row.at;
            if (changed) stale.push({ row, before });
            continue;
        }
        const chunks = normalizeChunks(row.id, await loadChunks(entry.source));
        touched.push({ row, chunks, isUpdate: before !== undefined, oldHash: before === undefined ? '' : String(before.hash) });
    }

    const removed = [...existing.keys()].filter((id) => !seen.has(id));
    const stamp = nowIso();
    const insertDoc = ctx.db.prepare(`
        INSERT INTO docs (id, path, folder, title, hash, origin, tier, kind, bytes, chars, chunks, at, text, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            path = excluded.path, folder = excluded.folder, title = excluded.title, hash = excluded.hash,
            origin = excluded.origin, tier = excluded.tier, kind = excluded.kind, bytes = excluded.bytes,
            chars = excluded.chars, chunks = excluded.chunks, at = excluded.at, text = excluded.text,
            updated_at = excluded.updated_at
    `);
    const updateDocMeta = ctx.db.prepare(`
        UPDATE docs SET path = ?, folder = ?, title = ?, tier = ?, kind = ?, bytes = ?, chars = ?, at = ?, updated_at = ?
        WHERE id = ?
    `);
    const insertChunk = ctx.db.prepare(`
        INSERT OR IGNORE INTO chunks (id, doc_id, n, text, head, outline, "start", "end", kind, image_path, summary_of, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const insertLex = ctx.db.prepare('INSERT OR REPLACE INTO chunks_lex (chunk_id, lex) VALUES (?, ?)');
    // 降级运行时库里没有 chunks_fts：连语句都不能 prepare（prepare 就会报 no such table）。
    const insertFts = ctx.fts ? ctx.db.prepare('INSERT INTO chunks_fts (chunk_id, lex, head_lex) VALUES (?, ?, ?)') : null;

    let chunksWritten = 0;
    let added = 0;
    let updated = 0;
    ctx.db.exec('BEGIN IMMEDIATE');
    try {
        for (const item of touched) {
            const { row } = item;
            if (item.isUpdate) {
                releaseDocChunks(ctx, row.id, item.oldHash);
                updated += 1;
            } else {
                added += 1;
            }
            insertDoc.run(
                row.id, row.path, row.folder, row.title, row.hash, origin, row.tier, row.kind,
                row.bytes, row.chars, item.chunks.length, row.at, row.text, stamp,
            );
            for (const chunk of item.chunks) {
                const info = insertChunk.run(
                    chunk.id, row.id, chunk.n, chunk.text, chunk.head, chunk.outline,
                    chunk.start, chunk.end, chunk.kind, chunk.image_path, chunk.summary_of, stamp,
                );
                if (Number(info.changes) > 0) chunksWritten += 1;
                const lex = lexicalOf(chunk);
                insertLex.run(chunk.id, lex.bag);
                if (insertFts !== null) insertFts.run(chunk.id, lex.textLex, lex.headLex);
            }
        }
        for (const item of stale) {
            updateDocMeta.run(
                item.row.path, item.row.folder, item.row.title, item.row.tier, item.row.kind,
                item.row.bytes, item.row.chars, item.row.at, stamp, item.row.id,
            );
        }
        for (const id of removed) {
            const before = existing.get(id);
            releaseDocChunks(ctx, id, before === undefined ? '' : String(before.hash));
            ctx.db.prepare('DELETE FROM docs WHERE id = ?').run(id);
        }
        ctx.db.exec('COMMIT');
    } catch (error) {
        try { ctx.db.exec('ROLLBACK'); } catch { /* 回滚失败只能把原始错误抛出去 */ }
        throw error;
    }
    return { added, updated, removed: removed.length, skipped, chunks: chunksWritten };
}

// ── 词法检索 ────────────────────────────────────────────────────────────────

/** filters → 绑在 docs 上的 WHERE 片段（`d.` 前缀由调用侧保证）。 */
function docFilters(filters) {
    const where = [];
    const params = [];
    const source = filters !== null && typeof filters === 'object' ? filters : {};
    const origin = asText(source.origin);
    if (origin !== '') {
        where.push('d.origin = ?');
        params.push(origin);
    }
    const tier = asText(source.tier).toLowerCase();
    if (tier !== '') {
        where.push('d.tier = ?');
        params.push(tier);
    }
    const docId = asText(source.docId);
    if (docId !== '') {
        where.push('d.id = ?');
        params.push(docId);
    }
    // 空串 = 根 = 不筛；非法路径让 normalizeKbFolder 自己抛（静默忽略等于换了个查询）。
    const folder = normalizeKbFolder(source.folder);
    if (source.exact === true) {
        // 资源管理器「仅当前层」：含根目录在内都按精确匹配，不能让空 folder 退化成全库。
        where.push('d.folder = ?');
        params.push(folder);
    } else if (folder !== '') {
        if (source.recursive === false) {
            where.push('d.folder = ?');
            params.push(folder);
        } else {
            where.push("(d.folder = ? OR d.folder LIKE ? ESCAPE '\\')");
            params.push(folder, `${escapeLike(folder)}/%`);
        }
    }
    return { where, params };
}

function hitOf(row) {
    return {
        chunkId: String(row.chunkId),
        docId: String(row.docId),
        n: Number(row.n ?? 0),
        text: String(row.text ?? ''),
        head: String(row.head ?? ''),
        folder: String(row.folder ?? ''),
        path: String(row.path ?? ''),
        title: String(row.title ?? ''),
        origin: String(row.origin ?? ''),
        tier: String(row.tier ?? ''),
        start: row.start === null || row.start === undefined ? null : Number(row.start),
        end: row.end === null || row.end === undefined ? null : Number(row.end),
        score: Number(row.score ?? 0),
    };
}

/** FTS5 查询串：每个 token 加双引号再 OR（引号内的特殊字符原样，不受查询语法影响）。 */
function ftsQueryOf(terms) {
    return terms.map((term) => `"${term.replace(/"/g, '""')}"`).join(' OR ');
}

function searchFts(ctx, terms, limit, filters) {
    const { where, params } = docFilters(filters);
    const sql = `
        SELECT c.id AS chunkId, c.doc_id AS docId, c.n AS n, c.text AS text, c.head AS head,
               d.folder AS folder, d.path AS path, d.title AS title, d.origin AS origin, d.tier AS tier,
               c."start" AS start, c."end" AS end,
               bm25(chunks_fts) AS rank
          FROM chunks_fts
          JOIN chunks c ON c.id = chunks_fts.chunk_id
          JOIN docs d ON d.id = c.doc_id
         WHERE chunks_fts MATCH ?${where.map((item) => ` AND ${item}`).join('')}
         ORDER BY rank ASC, c.id ASC
         LIMIT ?
    `;
    return ctx.db.prepare(sql).all(ftsQueryOf(terms), ...params, limit).map((row) => hitOf({ ...row, score: -Number(row.rank ?? 0) }));
}

function searchLike(ctx, terms, limit, filters) {
    const patterns = terms.map((term) => `%${escapeLike(term)}%`);
    const score = terms.map(() => "CASE WHEN l.lex LIKE ? ESCAPE '\\' THEN 1 ELSE 0 END").join(' + ');
    const { where, params } = docFilters(filters);
    const sql = `
        SELECT c.id AS chunkId, c.doc_id AS docId, c.n AS n, c.text AS text, c.head AS head,
               d.folder AS folder, d.path AS path, d.title AS title, d.origin AS origin, d.tier AS tier,
               c."start" AS start, c."end" AS end,
               (${score}) AS score
          FROM chunks_lex l
          JOIN chunks c ON c.id = l.chunk_id
          JOIN docs d ON d.id = c.doc_id
         WHERE (${score}) > 0${where.map((item) => ` AND ${item}`).join('')}
         ORDER BY score DESC, c.id ASC
         LIMIT ?
    `;
    return ctx.db.prepare(sql).all(...patterns, ...patterns, ...params, limit).map(hitOf);
}

/**
 * 词法检索：FTS5 优先，不可用或查询语法出错时退 LIKE 命中计数。
 *
 * 打分口径两种模式不同（bm25 是负数、取反后越大越好；LIKE 是命中词数），但返回
 * 形状与「score 越大越相关」的约定一致 —— 调用侧不必知道底下是哪条通路。
 */
function searchLexical(ctx, { query, limit = 20, filters = {} } = {}) {
    assertOpen(ctx);
    const terms = [...new Set(kbTokens(query))].slice(0, MAX_TERMS);
    if (terms.length === 0) return [];
    const cap = clampInt(limit, 20, 1, MAX_LIMIT);
    if (ctx.fts) {
        try {
            return searchFts(ctx, terms, cap, filters);
        } catch {
            // FTS5 运行时不可用（比如库是在有 FTS5 的 Node 上建的）：这一次退 LIKE，
            // 不改 handle.fts —— 调用方看到的能力位不该被一次查询失败改写。
        }
    }
    return searchLike(ctx, terms, cap, filters);
}

// ── 向量 ────────────────────────────────────────────────────────────────────

/** 向量入参 → Float32Array；认不出来返回 null（不猜维度）。 */
function toVector(value) {
    if (value instanceof Float32Array) return value.length > 0 ? value : null;
    if (value instanceof ArrayBuffer) {
        const view = new Float32Array(value);
        return view.length > 0 ? view : null;
    }
    if (Array.isArray(value) || ArrayBuffer.isView(value)) {
        const out = new Float32Array(value.length);
        for (let i = 0; i < value.length; i += 1) {
            const n = Number(value[i]);
            out[i] = Number.isFinite(n) ? n : 0;
        }
        return out.length > 0 ? out : null;
    }
    return null;
}

/** BLOB 读回来还是字节：对齐时直接开视图，不对齐就复制一份（Uint8Array 可能带偏移）。 */
function fromBlob(bytes, dim) {
    if (!(bytes instanceof Uint8Array)) return null;
    if (bytes.byteOffset % 4 === 0 && bytes.byteLength % 4 === 0) {
        const view = new Float32Array(bytes.buffer, bytes.byteOffset, Math.floor(bytes.byteLength / 4));
        return dim > 0 && dim < view.length ? view.slice(0, dim) : view;
    }
    const copy = new Uint8Array(bytes.byteLength);
    copy.set(bytes);
    return new Float32Array(copy.buffer);
}

function normOf(vector) {
    let sum = 0;
    for (let i = 0; i < vector.length; i += 1) sum += vector[i] * vector[i];
    return Math.sqrt(sum);
}

function cosineOf(left, right, leftNorm, rightNorm) {
    if (!(leftNorm > 0) || !(rightNorm > 0)) return 0;
    const len = Math.min(left.length, right.length);
    let dot = 0;
    for (let i = 0; i < len; i += 1) dot += left[i] * right[i];
    return dot / (leftNorm * rightNorm);
}

function upsertEmbeddings(ctx, items) {
    assertOpen(ctx);
    const list = Array.isArray(items) ? items : [];
    if (list.length === 0) return { stored: 0 };
    const stamp = nowIso();
    const stmt = ctx.db.prepare(`
        INSERT INTO embeddings (chunk_id, model, dim, vector, updated_at) VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(chunk_id) DO UPDATE SET
            model = excluded.model, dim = excluded.dim, vector = excluded.vector, updated_at = excluded.updated_at
    `);
    let stored = 0;
    ctx.db.exec('BEGIN IMMEDIATE');
    try {
        for (const item of list) {
            const chunkId = asText(item?.chunkId);
            const vector = toVector(item?.vector);
            if (chunkId === '' || vector === null) continue;
            const bytes = new Uint8Array(vector.buffer, vector.byteOffset, vector.byteLength);
            stmt.run(chunkId, asText(item?.model), vector.length, bytes, stamp);
            stored += 1;
        }
        ctx.db.exec('COMMIT');
    } catch (error) {
        try { ctx.db.exec('ROLLBACK'); } catch { /* 见上 */ }
        throw error;
    }
    // 图索引跟着库走：事务提交后再改图，回滚时图不会先跑一步（图坏了也只是缓存）。
    annUpsertRows(ctx, list);
    return { stored };
}

function embeddingFor(ctx, chunkId) {
    assertOpen(ctx);
    const id = asText(chunkId);
    if (id === '') return null;
    const row = ctx.db.prepare('SELECT model, dim, vector FROM embeddings WHERE chunk_id = ?').get(id);
    if (row === undefined || row === null) return null;
    const vector = fromBlob(row.vector, 0);
    if (vector === null || vector.length === 0) return null;
    return { model: String(row.model ?? ''), dim: vector.length, vector };
}

// ── ANN 索引（第一百零三轮） ─────────────────────────────────────────────────
//
// 索引是 embeddings 表的**派生缓存**：模型/维度/条数/更新时间任何一项对不上就重建，
// 文件损坏也只是重建一次，检索不会因此整块失败（最坏退回精确余弦）。
// 落盘走 kb-ann.saveAnn 的「临时文件 + rename」，多进程同时写只会后写者覆盖，
// 读方要么看到旧图要么看到新图，不会读到半截文件。

/** 收口 ANN 参数：坏值落默认，enabled 只认显式 false。 */
function annOptions(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const m = clampInt(source.m, ANN_DEFAULTS.m, 4, 64);
    return {
        enabled: source.enabled !== false,
        m,
        efConstruction: clampInt(source.efConstruction, ANN_DEFAULTS.efConstruction, m, 1_000),
        efSearch: clampInt(source.efSearch, ANN_DEFAULTS.efSearch, 1, 512),
        maxBuild: clampInt(source.maxBuild, ANN_MAX_BUILD, 100, ANN_MAX_BUILD),
    };
}

/**
 * 当前 embeddings 表能建图的「唯一模型/维度」。混用两套向量时**不建图**：
 * 不同模型/维度的余弦不可比，硬塞进一张图只会给出错的结果，不如整表精确扫描。
 */
function annSourceOf(ctx) {
    const rows = ctx.db.prepare(
        'SELECT model, dim, count(*) AS n, max(updated_at) AS updated FROM embeddings GROUP BY model, dim ORDER BY n DESC',
    ).all();
    if (rows.length === 0) return null;
    if (rows.length > 1) {
        return { mixed: true, models: rows.map((row) => `${String(row.model)}:${Number(row.dim)}(${Number(row.n)})`) };
    }
    const row = rows[0];
    return {
        model: String(row.model ?? ''),
        dim: Number(row.dim) || 0,
        count: Number(row.n) || 0,
        updatedAt: String(row.updated ?? ''),
    };
}

function annKeyOf(source) {
    return `${source.model}\u0000${source.dim}`;
}

/** 落盘（只写不抛）：索引写不进去不影响检索，下次重建就是。 */
function persistAnn(ctx, source) {
    if (ctx.ann === null || ctx.ann.nodes === 0 || ctx.readonly) return false;
    try {
        saveAnn(ctx.ann, {
            dir: dirname(ctx.path),
            model: source.model,
            updatedAt: source.updatedAt,
            count: source.count,
        });
        return true;
    } catch {
        return false;
    }
}

/**
 * 懒加载 / 懒构建：第一次向量检索时按 embeddings 表现状装回或重建。
 *
 * 返回的图只覆盖「计数最多的那个模型/维度」；调用方据此重算，过滤后不够再走
 * 精确扫描补足。`maxBuild` 之上不建图（同步建图会把一次检索拖成几十秒）。
 */
function ensureAnn(ctx, options = {}) {
    const opts = annOptions(options);
    if (!opts.enabled) return null;
    const source = annSourceOf(ctx);
    if (source === null || source.count === 0) return null;
    if (source.mixed === true) {
        ctx.ann = null;
        ctx.annKey = '';
        ctx.annMeta = null;
        ctx.annSkipped = 'mixed-models';
        return null;
    }
    const key = annKeyOf(source);
    if (ctx.ann !== null && ctx.annKey === key) return ctx.ann;
    if (source.count > opts.maxBuild) {
        ctx.ann = null;
        ctx.annKey = '';
        ctx.annMeta = null;
        ctx.annSkipped = `too-large:${source.count}`;
        return null;
    }
    const dir = dirname(ctx.path);
    const loaded = loadAnn({ dir, model: source.model, dim: source.dim, count: source.count, updatedAt: source.updatedAt });
    if (loaded !== null) {
        ctx.ann = loaded.index;
        ctx.annKey = key;
        ctx.annMeta = { model: source.model, dim: source.dim };
        ctx.annDirty = false;
        ctx.annSkipped = '';
        // 墓碑是删除留下的：攒多了就整图重建，检索不会被越来越多的死节点拖慢。
        if (ctx.ann.tombstones > Math.max(64, ctx.ann.size * 0.3)) {
            ctx.ann.compact();
            ctx.annDirty = true;
        }
        return ctx.ann;
    }
    const index = createAnnIndex({ dim: source.dim, m: opts.m, efConstruction: opts.efConstruction, efSearch: opts.efSearch });
    const rows = ctx.db.prepare(
        'SELECT chunk_id, vector FROM embeddings WHERE model = ? AND dim = ? ORDER BY chunk_id ASC',
    ).all(source.model, source.dim);
    for (const row of rows) {
        const vector = fromBlob(row.vector, source.dim);
        if (vector !== null) index.add(String(row.chunk_id), vector);
    }
    if (index.nodes === 0) return null;
    ctx.ann = index;
    ctx.annKey = key;
    ctx.annMeta = { model: source.model, dim: source.dim };
    ctx.annDirty = false;
    ctx.annSkipped = '';
    persistAnn(ctx, source);
    return ctx.ann;
}

/** 新增/更新的向量增量进图；模型或维度一变就丢掉整张图（下一次检索重建）。 */
function annUpsertRows(ctx, items) {
    if (ctx.ann === null || ctx.annMeta === null) return;
    let changed = false;
    for (const item of items) {
        const vector = toVector(item?.vector);
        if (vector === null) continue;
        if (asText(item?.model) !== ctx.annMeta.model || vector.length !== ctx.annMeta.dim) {
            ctx.ann = null;
            ctx.annKey = '';
            ctx.annMeta = null;
            ctx.annDirty = false;
            return;
        }
        if (ctx.ann.add(asText(item?.chunkId), vector)) changed = true;
    }
    if (changed) ctx.annDirty = true;
}

/** 删除只打墓碑（保住图上的连通性），攒够了在 ensureAnn 里 compact。 */
function annRemoveIds(ctx, ids) {
    if (ctx.ann === null) return;
    let changed = false;
    for (const id of ids) {
        if (ctx.ann.remove(asText(id))) changed = true;
    }
    if (changed) ctx.annDirty = true;
}

/**
 * 关库前落盘。行数与存活数对不上说明还有别处的写入没进图：丢掉缓存、
 * 下一次检索重建 —— 存一张对不上账的图比不存更危险。
 */
function flushAnn(ctx) {
    if (ctx.ann === null || !ctx.annDirty) return false;
    ctx.annDirty = false;
    if (ctx.readonly) return false;
    const source = annSourceOf(ctx);
    if (source === null || source.mixed === true || source.count !== ctx.ann.size) {
        ctx.ann = null;
        ctx.annKey = '';
        ctx.annMeta = null;
        return false;
    }
    const saved = persistAnn(ctx, source);
    ctx.annDirty = !saved;
    return saved;
}

function annStatusOf(ctx) {
    const meta = ctx.annMeta;
    return {
        enabled: ctx.ann !== null,
        model: meta === null ? '' : meta.model,
        dim: meta === null ? 0 : meta.dim,
        nodes: ctx.ann === null ? 0 : ctx.ann.nodes,
        active: ctx.ann === null ? 0 : ctx.ann.size,
        tombstones: ctx.ann === null ? 0 : ctx.ann.tombstones,
        skipped: ctx.annSkipped,
    };
}

/** ANN 候选 → 按库里的原始向量精确重算 + 范围过滤（ANN 本身不认识 filters）。 */
function rescoreByIds(ctx, ids, query, queryNorm, filters) {
    if (ids.length === 0) return [];
    const { where, params } = docFilters(filters);
    const hits = [];
    for (const batch of batches(ids, IN_BATCH)) {
        const marks = placeholders(batch.length);
        const sql = `
            SELECT e.chunk_id AS chunkId, e.vector AS vector, c.doc_id AS docId, c.n AS n,
                   c.text AS text, c.head AS head, d.folder AS folder, d.path AS path,
                   d.title AS title, d.origin AS origin, d.tier AS tier,
                   c."start" AS start, c."end" AS end
              FROM embeddings e
              JOIN chunks c ON c.id = e.chunk_id
              JOIN docs d ON d.id = c.doc_id
             WHERE e.chunk_id IN (${marks})${where.length === 0 ? '' : ` AND ${where.join(' AND ')}`}
        `;
        const rows = ctx.db.prepare(sql).all(...batch, ...params);
        for (const row of rows) {
            const stored = fromBlob(row.vector, 0);
            if (stored === null) continue;
            const score = cosineOf(query, stored, queryNorm, normOf(stored));
            hits.push(hitOf({ ...row, score }));
        }
    }
    hits.sort((left, right) => (right.score - left.score) || left.chunkId.localeCompare(right.chunkId));
    return hits;
}

/**
 * ANN 路径：先取 `候选池 = max(efSearch, 返回条数 × 6)`，精确重算并按 filters 收窄；
 * 不够就翻倍再来（最多三次），仍不够再用精确扫描补足。
 *
 * 「过滤后召回不足」是图索引的固有短板：索引不知道 scope，只能在候选里筛。补足
 * 的做法是把精确扫描当兜底，代价是这一次退化成原来的全表余弦 —— 能给出完整结果，
 * 不会静默少给。
 */
function searchVectorAnn(ctx, ann, query, queryNorm, cap, filters, opts) {
    const ceiling = Math.min(MAX_VECTOR_SCAN, ann.size);
    let pool = Math.min(ceiling, Math.max(opts.efSearch, cap * ANN_CANDIDATE_FACTOR));
    const best = new Map();
    for (let attempt = 0; attempt < 3; attempt += 1) {
        let candidates;
        try {
            candidates = ann.search(query, pool, { efSearch: Math.max(opts.efSearch, pool) });
        } catch {
            return null;
        }
        const hits = rescoreByIds(ctx, candidates.map((item) => item.id), query, queryNorm, filters);
        for (const hit of hits) if (!best.has(hit.chunkId)) best.set(hit.chunkId, hit);
        if (best.size >= cap) break;
        if (pool >= ceiling) break;
        pool = Math.min(ceiling, pool * 2);
    }
    if (best.size >= cap) {
        return [...best.values()]
            .sort((left, right) => (right.score - left.score) || left.chunkId.localeCompare(right.chunkId))
            .slice(0, cap);
    }
    const exact = searchVectorExact(ctx, { vector: query, limit: cap, filters });
    for (const hit of exact) if (!best.has(hit.chunkId)) best.set(hit.chunkId, hit);
    return [...best.values()]
        .sort((left, right) => (right.score - left.score) || left.chunkId.localeCompare(right.chunkId))
        .slice(0, cap);
}

/**
 * 向量检索：有 ANN 就走「图候选 + 精确重算」，没有（关掉 / 规模超限 / 多模型
 * 混用 / 构建失败）就走精确余弦。两条路返回同一种命中形状，调用方不用关心。
 */
function searchVector(ctx, { vector, limit = 20, filters = {}, ann = {} } = {}) {
    assertOpen(ctx);
    const query = toVector(vector);
    if (query === null) return [];
    const queryNorm = normOf(query);
    if (!(queryNorm > 0)) return [];
    const cap = clampInt(limit, 20, 1, MAX_LIMIT);
    const opts = annOptions(ann);
    let index = null;
    try {
        index = ensureAnn(ctx, opts);
    } catch {
        index = null;
    }
    if (index !== null && index.dim === query.length) {
        const hits = searchVectorAnn(ctx, index, query, queryNorm, cap, filters, opts);
        if (hits !== null) return hits;
    }
    return searchVectorExact(ctx, { vector: query, limit: cap, filters });
}

/**
 * 精确余弦：只扫命中的块，最多 `MAX_VECTOR_SCAN` 个。
 *
 * 它现在是 ANN 的兜底通道（关闭索引、规模超限、模型混用、过滤后补足）。
 * 几万块以内仍是毫秒到百毫秒级；真到扫不完的规模，说明该换 ANN 或分库了。
 */
function searchVectorExact(ctx, { vector, limit = 20, filters = {} } = {}) {
    assertOpen(ctx);
    const query = toVector(vector);
    if (query === null) return [];
    const queryNorm = normOf(query);
    if (!(queryNorm > 0)) return [];
    const cap = clampInt(limit, 20, 1, MAX_LIMIT);
    const { where, params } = docFilters(filters);
    const sql = `
        SELECT e.chunk_id AS chunkId, e.vector AS vector, c.doc_id AS docId, c.n AS n,
               c.text AS text, c.head AS head, d.folder AS folder, d.path AS path,
               d.title AS title, d.origin AS origin, d.tier AS tier,
               c."start" AS start, c."end" AS end
          FROM embeddings e
          JOIN chunks c ON c.id = e.chunk_id
          JOIN docs d ON d.id = c.doc_id
         ${where.length === 0 ? '' : `WHERE ${where.join(' AND ')}`}
         ORDER BY e.chunk_id ASC
         LIMIT ?
    `;
    const rows = ctx.db.prepare(sql).all(...params, MAX_VECTOR_SCAN);
    const hits = [];
    for (const row of rows) {
        const stored = fromBlob(row.vector, 0);
        if (stored === null) continue;
        const score = cosineOf(query, stored, queryNorm, normOf(stored));
        hits.push(hitOf({ ...row, score }));
    }
    hits.sort((left, right) => (right.score - left.score) || left.chunkId.localeCompare(right.chunkId));
    return hits.slice(0, cap);
}

// ── 关系（块 ↔ 块） ─────────────────────────────────────────────────────────

function linkIdOf(sourceId, targetId, kind) {
    const hash = createHash('sha1').update(`${sourceId}\u0000${targetId}\u0000${kind}`, 'utf8').digest('hex').slice(0, 12);
    return `kb-link:${hash}`;
}

function linkRowOf(row) {
    return {
        id: String(row.id),
        sourceId: String(row.source_id),
        targetId: String(row.target_id),
        kind: String(row.kind),
        note: String(row.note ?? ''),
        at: String(row.at ?? ''),
    };
}

function link(ctx, { sourceId, targetId, kind = 'related', note = '' } = {}) {
    assertOpen(ctx);
    const source = asText(sourceId);
    const target = asText(targetId);
    if (source === '' || target === '') throw new Error('link 需要 sourceId 与 targetId');
    const type = asText(kind) || 'related';
    ctx.db.prepare(`
        INSERT INTO chunk_links (id, source_id, target_id, kind, note, at) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(source_id, target_id, kind) DO UPDATE SET note = excluded.note
    `).run(linkIdOf(source, target, type), source, target, type, asText(note), nowIso());
    const row = ctx.db.prepare('SELECT * FROM chunk_links WHERE source_id = ? AND target_id = ? AND kind = ?').get(source, target, type);
    return linkRowOf(row);
}

function unlink(ctx, { sourceId, targetId, kind } = {}) {
    assertOpen(ctx);
    const source = asText(sourceId);
    const target = asText(targetId);
    if (source === '' || target === '') return { removed: 0 };
    const type = asText(kind);
    const where = ['source_id = ?', 'target_id = ?'];
    const params = [source, target];
    if (type !== '') {
        where.push('kind = ?');
        params.push(type);
    }
    const info = ctx.db.prepare(`DELETE FROM chunk_links WHERE ${where.join(' AND ')}`).run(...params);
    return { removed: Number(info.changes ?? 0) };
}

/** 一条块的出入边（默认两边都要：关系是双向可走的）。 */
function linksFor(ctx, { chunkId, direction = 'both' } = {}) {
    assertOpen(ctx);
    const id = asText(chunkId);
    if (id === '') return [];
    const dir = direction === 'out' || direction === 'in' ? direction : 'both';
    const out = [];
    if (dir !== 'in') {
        for (const row of ctx.db.prepare('SELECT * FROM chunk_links WHERE source_id = ? ORDER BY kind ASC, target_id ASC').all(id)) {
            out.push({ ...linkRowOf(row), direction: 'out', neighborId: String(row.target_id), chunkId: String(row.target_id) });
        }
    }
    if (dir !== 'out') {
        for (const row of ctx.db.prepare('SELECT * FROM chunk_links WHERE target_id = ? ORDER BY kind ASC, source_id ASC').all(id)) {
            out.push({ ...linkRowOf(row), direction: 'in', neighborId: String(row.source_id), chunkId: String(row.source_id) });
        }
    }
    return out;
}

function chunkMetaByIds(ctx, ids) {
    const meta = new Map();
    if (ids.length === 0) return meta;
    for (const batch of batches(ids, IN_BATCH)) {
        const sql = `
            SELECT c.id AS chunkId, c.doc_id AS docId, c.n AS n, c.text AS text, c.head AS head,
                   d.folder AS folder, d.path AS path, d.title AS title, d.origin AS origin, d.tier AS tier
              FROM chunks c JOIN docs d ON d.id = c.doc_id
             WHERE c.id IN (${placeholders(batch.length)})
        `;
        for (const row of ctx.db.prepare(sql).all(...batch)) {
            meta.set(String(row.chunkId), {
                docId: String(row.docId),
                n: Number(row.n ?? 0),
                text: String(row.text ?? ''),
                head: String(row.head ?? ''),
                folder: String(row.folder ?? ''),
                path: String(row.path ?? ''),
                title: String(row.title ?? ''),
                origin: String(row.origin ?? ''),
                tier: String(row.tier ?? ''),
            });
        }
    }
    return meta;
}

/** 广度优先走关系图；每条结果带 `hop`（第几跳）。 */
function related(ctx, { chunkId, depth = 1, limit = 20 } = {}) {
    assertOpen(ctx);
    const start = asText(chunkId);
    if (start === '') return [];
    const cap = clampInt(limit, 20, 1, MAX_LIMIT);
    const hops = clampInt(depth, 1, 1, MAX_DEPTH);
    const seen = new Set([start]);
    const out = [];
    let frontier = [start];
    const neighbors = ctx.db.prepare('SELECT * FROM chunk_links WHERE source_id = ? OR target_id = ? ORDER BY kind ASC, source_id ASC, target_id ASC');
    for (let hop = 1; hop <= hops && out.length < cap && frontier.length > 0; hop += 1) {
        const next = [];
        for (const id of frontier) {
            for (const row of neighbors.all(id, id)) {
                const neighbor = String(row.source_id) === id ? String(row.target_id) : String(row.source_id);
                if (neighbor === '' || seen.has(neighbor)) continue;
                seen.add(neighbor);
                out.push({ chunkId: neighbor, hop, kind: String(row.kind), via: id, linkId: String(row.id) });
                next.push(neighbor);
                if (out.length >= cap) break;
            }
            if (out.length >= cap) break;
        }
        frontier = next;
    }
    const meta = chunkMetaByIds(ctx, out.map((item) => item.chunkId));
    return out.map((item) => ({
        ...item,
        ...(meta.get(item.chunkId) ?? {
            docId: '', n: 0, text: '', head: '', folder: '', path: '', title: '', origin: '', tier: '',
        }),
    }));
}

// ── 摘要 ────────────────────────────────────────────────────────────────────

function summaryRowOf(row) {
    let chunkIds = [];
    try {
        const parsed = JSON.parse(String(row.chunk_ids ?? '[]'));
        if (Array.isArray(parsed)) chunkIds = parsed.map((item) => String(item));
    } catch {
        chunkIds = [];
    }
    return {
        id: String(row.id),
        title: String(row.title ?? ''),
        text: String(row.text ?? ''),
        docId: String(row.doc_id ?? ''),
        chunkIds,
        createdAt: String(row.created_at ?? ''),
    };
}

function saveSummary(ctx, { title, text, docId, chunkIds, id } = {}) {
    assertOpen(ctx);
    const doc = asText(docId);
    if (doc === '') throw new Error('saveSummary 需要 docId');
    const found = ctx.db.prepare('SELECT id FROM docs WHERE id = ?').get(doc);
    if (found === undefined || found === null) throw new Error(`saveSummary：文档不在索引里（${doc}）`);
    const ids = (Array.isArray(chunkIds) ? chunkIds : []).map((item) => asText(item)).filter((item) => item !== '');
    const summaryId = asText(id) || `kb-sum:${createHash('sha1').update(`${doc}\u0000${ids.join(',')}\u0000${String(text ?? '')}`, 'utf8').digest('hex').slice(0, 12)}`;
    ctx.db.prepare(`
        INSERT INTO summaries (id, title, text, doc_id, chunk_ids, created_at) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            title = excluded.title, text = excluded.text, doc_id = excluded.doc_id, chunk_ids = excluded.chunk_ids
    `).run(summaryId, asText(title), String(text ?? ''), doc, JSON.stringify(ids), nowIso());
    return { id: summaryId, docId: doc };
}

function listSummaries(ctx, { limit = 50, offset = 0 } = {}) {
    assertOpen(ctx);
    const skip = clampInt(offset, 0, 0, Number.MAX_SAFE_INTEGER);
    const cap = clampInt(limit, 50, 1, MAX_LIMIT);
    // 带出来源文档的 path / folder：面板要能一眼看出「这条摘要属于哪份文档、盖了几块」，
    // 只有 docId（内容寻址的哈希）对人没有意义。
    return ctx.db.prepare(`
        SELECT s.*, d.path AS doc_path, d.folder AS folder, d.title AS doc_title
          FROM summaries s JOIN docs d ON d.id = s.doc_id
         ORDER BY s.created_at DESC, s.id ASC
         LIMIT ? OFFSET ?
    `).all(cap, skip).map((row) => ({
        ...summaryRowOf(row),
        docPath: String(row.doc_path ?? ''),
        folder: String(row.folder ?? ''),
        docTitle: String(row.doc_title ?? ''),
    }));
}

/** 摘要块总数（分页用）：与 listSummaries 同一来源，不受当页 LIMIT 影响。 */
function countSummaries(ctx) {
    assertOpen(ctx);
    return Number(ctx.db.prepare('SELECT COUNT(*) AS n FROM summaries').get()?.n ?? 0);
}

/**
 * 按 id 批量取块元数据（不取正文）：摘要块的「来源块」清单靠它拼。
 *
 * 块 id 是内容寻址的，来源块可能已经不在了（文档被删、内容变了），所以返回**按入参
 * 顺序**排列、缺的跳过，调用方自己按 id 对齐 —— 静默补位会让「哪一块没了」看不出来。
 */
function chunksByIds(ctx, { ids = [] } = {}) {
    assertOpen(ctx);
    const list = (Array.isArray(ids) ? ids : []).map((item) => asText(item)).filter((item) => item !== '');
    if (list.length === 0) return [];
    const marks = list.map(() => '?').join(', ');
    const rows = ctx.db.prepare(`
        SELECT c.id AS id, c.doc_id AS doc_id, c.n AS n, c.head AS head, c.kind AS kind,
               c."start" AS "start", c."end" AS "end", length(c.text) AS chars,
               d.path AS doc_path, d.title AS doc_title, d.folder AS folder
          FROM chunks c JOIN docs d ON d.id = c.doc_id
         WHERE c.id IN (${marks})
    `).all(...list);
    const byId = new Map(rows.map((row) => [String(row.id), row]));
    return list.filter((id) => byId.has(id)).map((id) => {
        const row = byId.get(id);
        return {
            id: String(row.id),
            docId: String(row.doc_id ?? ''),
            docPath: String(row.doc_path ?? ''),
            docTitle: String(row.doc_title ?? ''),
            folder: String(row.folder ?? ''),
            n: Number(row.n) || 0,
            kind: String(row.kind ?? 'text'),
            head: String(row.head ?? ''),
            start: row.start === null || row.start === undefined ? null : Number(row.start),
            end: row.end === null || row.end === undefined ? null : Number(row.end),
            chars: Number(row.chars) || 0,
        };
    });
}

/**
 * 摘要块的词法检索（第八十五轮补齐）。
 *
 * 摘要不是文档，不写 chunks / FTS / embeddings 表：它由若干来源块生成，来源文档
 * 变化后旧摘要也不该继续冒充原文块。所以这里单独 JOIN docs 取来源筛选字段，
 * 对 title / text 做与 chunks_lex 同口径的 token LIKE，命中词数就是分数。
 */
function searchSummaries(ctx, { query, limit = 20, filters = {} } = {}) {
    assertOpen(ctx);
    const terms = [...new Set(kbTokens(query))].slice(0, MAX_TERMS);
    if (terms.length === 0) return [];
    const cap = clampInt(limit, 20, 1, MAX_LIMIT);
    const patterns = terms.map((term) => `%${escapeLike(term)}%`);
    const score = terms.map(() => "CASE WHEN (s.title || ' ' || s.text) LIKE ? ESCAPE '\\' THEN 1 ELSE 0 END").join(' + ');
    const { where, params } = docFilters(filters);
    const sql = `
        SELECT s.*, d.origin AS origin, d.path AS path, d.folder AS folder, d.tier AS tier,
               (${score}) AS score
          FROM summaries s
          JOIN docs d ON d.id = s.doc_id
         WHERE (${score}) > 0${where.map((item) => ` AND ${item}`).join('')}
         ORDER BY score DESC, s.created_at DESC, s.id ASC
         LIMIT ?
    `;
    return ctx.db.prepare(sql).all(...patterns, ...patterns, ...params, cap).map((row) => ({
        ...summaryRowOf(row),
        origin: String(row.origin ?? ''),
        path: String(row.path ?? ''),
        folder: String(row.folder ?? ''),
        tier: String(row.tier ?? ''),
        score: Number(row.score ?? 0),
    }));
}

function summaryFor(ctx, id) {
    assertOpen(ctx);
    const value = asText(id);
    if (value === '') return null;
    const row = ctx.db.prepare(`
        SELECT s.*, d.origin AS origin, d.folder AS folder, d.tier AS tier
          FROM summaries s JOIN docs d ON d.id = s.doc_id
         WHERE s.id = ?
    `).get(value);
    if (row === undefined || row === null) return null;
    return {
        ...summaryRowOf(row),
        origin: String(row.origin ?? ''),
        folder: String(row.folder ?? ''),
        tier: String(row.tier ?? ''),
    };
}

// ── 统计 / 补嵌入 ───────────────────────────────────────────────────────────

function stats(ctx) {
    assertOpen(ctx);
    return {
        docs: countOf(ctx.db, 'SELECT count(*) AS n FROM docs'),
        chunks: countOf(ctx.db, 'SELECT count(*) AS n FROM chunks'),
        embeddings: countOf(ctx.db, 'SELECT count(*) AS n FROM embeddings'),
        links: countOf(ctx.db, 'SELECT count(*) AS n FROM chunk_links'),
        summaries: countOf(ctx.db, 'SELECT count(*) AS n FROM summaries'),
        fts: ctx.fts === true,
        bytes: countOf(ctx.db, 'SELECT coalesce(sum(bytes), 0) AS n FROM docs'),
    };
}

/** 命中的全部块 id（按 filters 收窄），给补嵌入用。 */
function allChunkIds(ctx, { filters = {} } = {}) {
    assertOpen(ctx);
    const { where, params } = docFilters(filters);
    const sql = `
        SELECT c.id AS id FROM chunks c JOIN docs d ON d.id = c.doc_id
        ${where.length === 0 ? '' : `WHERE ${where.join(' AND ')}`}
        ORDER BY c.id ASC
    `;
    return ctx.db.prepare(sql).all(...params).map((row) => String(row.id));
}
