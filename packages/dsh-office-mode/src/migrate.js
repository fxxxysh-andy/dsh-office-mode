/**
 * 把 mnemon 的记忆迁进办公记忆的三层。
 *
 * 迁移是**降级映射**，不是等量搬运 —— 两个系统的层不一样：
 *
 *   mnemon                          →  办公记忆
 *   runtime USER.md（全局）          →  热记忆 target:'user'
 *   runtime MEMORY.md（按工作目录）   →  热记忆 target:'project'
 *   Documents（项目档案）            →  没有对应层（本迁移只报数量，不搬）
 *   Memory Spaces（长期记忆 + 图）    →  归档（只读长期层），类别/重要度/标签/实体/
 *                                         关系边数都随条目保留；图跳转降级为计数
 *
 * 三条实现上的取舍：
 *   1. **幂等**。热记忆按内容去重；归档条目按 mnemon 的原始 id 去重。重复跑迁移
 *      不会产生第二份。
 *   2. **不覆盖**。只做 add，绝不 replace/remove：迁移不该动用户后来改过的条目。
 *   3. **可降级**。读 sqlite 需要 Node 自带的 `node:sqlite`（Node 22.5+）。拿不到时
 *      只报警告并继续迁 runtime，不让整次迁移失败。
 *
 * @module dsh-office-mode/migrate
 */
import { existsSync } from 'node:fs';
import { readFile, readdir } from 'node:fs/promises';
import { homedir } from 'node:os';
import { isAbsolute, join } from 'node:path';
import { createMemory } from './memory.js';

/** mnemon 的数据目录（相对会话工作目录），可用 source 参数覆盖。 */
export const DEFAULT_MNEMON_DIR = '.mnemon';
/** mnemon 的全局存储根；与 dsh-mnemon 的约定一致（见其 createStorageRoot）。 */
export const MNEMON_DATA_DIR_ENV = 'MNEMON_DATA_DIR';

/** 解析 mnemon 的工作目录数据根。 */
export function mnemonWorkspaceRoot(root, source = DEFAULT_MNEMON_DIR) {
    const value = typeof source === 'string' && source.trim() !== '' ? source.trim() : DEFAULT_MNEMON_DIR;
    return isAbsolute(value) ? value : join(root, value);
}

/** 解析 mnemon 的全局数据根（runtimeUserScope = "global" 时的 USER.md 就落在这里）。 */
export function mnemonGlobalRoot(explicit) {
    const value = typeof explicit === 'string' && explicit.trim() !== '' ? explicit.trim() : '';
    if (value !== '') return isAbsolute(value) ? value : join(process.cwd(), value);
    const fromEnv = typeof process.env[MNEMON_DATA_DIR_ENV] === 'string' ? process.env[MNEMON_DATA_DIR_ENV].trim() : '';
    if (fromEnv !== '') return fromEnv;
    return join(homedir(), '.mnemon');
}

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

function asArray(value) {
    return Array.isArray(value) ? value : [];
}

/** 读一份 mnemon runtime 存储（memories.json）。读不到就是空，不报错。 */
export async function readRuntimeEntries(dir) {
    const file = join(dir, 'runtime', 'memories.json');
    if (!existsSync(file)) return { file, exists: false, entries: [] };
    try {
        const parsed = JSON.parse(await readFile(file, 'utf8'));
        const entries = asArray(parsed?.entries)
            .filter((entry) => entry !== null && typeof entry === 'object' && asText(entry.content) !== '')
            .map((entry) => ({
                content: asText(entry.content),
                target: entry.target === 'user' ? 'user' : 'project',
                importance: ['critical', 'normal', 'low'].includes(entry.importance) ? entry.importance : 'normal',
                createdAt: asText(entry.created_at),
                updatedAt: asText(entry.updated_at),
            }));
        return { file, exists: true, entries };
    } catch (error) {
        return { file, exists: true, entries: [], error: error instanceof Error ? error.message : String(error) };
    }
}

/**
 * mnemon 的长期记忆库：`<data>/<memoryBodyId>/mnemon.db`。
 * 一个工作目录可能注册了多个记忆体，所以这里是「找出全部」而不是写死一个路径。
 */
export async function findInsightDatabases(dataRoot) {
    const dataDir = join(dataRoot, 'data');
    if (!existsSync(dataDir)) return [];
    let bodies = [];
    try {
        const registry = JSON.parse(await readFile(join(dataDir, '.dsh-memory-bodies.json'), 'utf8'));
        bodies = asArray(registry?.bodies);
    } catch {
        // 没有注册表也能靠目录名找到库
    }
    const found = [];
    for (const entry of await readdir(dataDir, { withFileTypes: true })) {
        if (!entry.isDirectory()) continue;
        const dbPath = join(dataDir, entry.name, 'mnemon.db');
        if (!existsSync(dbPath)) continue;
        const body = bodies.find((item) => item?.id === entry.name);
        found.push({ dbPath, bodyId: entry.name, bodyName: asText(body?.name) });
    }
    return found;
}

/**
 * 默认的长期记忆读取器：直接用 Node 自带的 `node:sqlite` 读 mnemon-native 的库。
 *
 * 为什么不用 mnemon 的 API：迁移的落点正是「关掉 mnemon」，工具可能已经不在；
 * 数据格式是插件的存储契约（insights / edges 两张表），直接读更稳。
 * 返回 `{ insights, edges }`，形参与测试注入的替身一致。
 */
export async function readMnemonInsights(dbPath) {
    let sqlite;
    try {
        sqlite = await import('node:sqlite');
    } catch (error) {
        throw new Error(`本机 Node 没有 node:sqlite（需要 Node 22.5+），无法读长期记忆：${error instanceof Error ? error.message : String(error)}`);
    }
    const db = new sqlite.DatabaseSync(dbPath, { readOnly: true });
    try {
        const insights = db.prepare('SELECT id, content, category, importance, tags, entities, source, stored_at, deleted_at FROM insights').all();
        const edges = db.prepare('SELECT source_id, target_id, edge_type FROM edges').all();
        return { insights, edges };
    } finally {
        db.close();
    }
}

function parseList(value) {
    if (Array.isArray(value)) return value.map((item) => asText(String(item))).filter((item) => item !== '');
    if (typeof value === 'string' && value.trim() !== '') {
        try {
            const parsed = JSON.parse(value);
            if (Array.isArray(parsed)) return parsed.map((item) => asText(String(item))).filter((item) => item !== '');
        } catch {
            return value.split(/[,，、]/).map((item) => item.trim()).filter((item) => item !== '');
        }
    }
    return [];
}

/** 把 451 条关系边折成每条记忆的邻居计数：办公记忆没有图，保留「有多少条关系」这个事实。 */
function linkCounts(edges, id) {
    const counts = { temporal: 0, entity: 0, causal: 0, semantic: 0, total: 0 };
    for (const edge of edges) {
        if (edge?.source_id !== id && edge?.target_id !== id) continue;
        const type = asText(edge?.edge_type);
        if (Object.prototype.hasOwnProperty.call(counts, type)) counts[type] += 1;
        counts.total += 1;
    }
    return counts;
}

/**
 * 执行迁移。
 *
 * @param {object} options
 *   root          会话工作目录（办公记忆的落点）
 *   memory        已解析的记忆配置（config.memory）
 *   source        mnemon 的工作目录数据根，默认 .mnemon
 *   globalSource  mnemon 的全局数据根，默认取 MNEMON_DATA_DIR 或 ~/.mnemon
 *   readInsights  注入的长期记忆读取器（测试用；默认读 sqlite）
 *   dryRun        true 时只统计不落盘
 */
export async function migrateMnemon({
    root,
    memory = {},
    source = DEFAULT_MNEMON_DIR,
    globalSource,
    readInsights = readMnemonInsights,
    dryRun = false,
} = {}) {
    if (asText(root) === '') throw new Error('migrateMnemon 需要会话工作目录。');
    const workspaceRoot = mnemonWorkspaceRoot(root, source);
    const globalRoot = mnemonGlobalRoot(globalSource);
    const store = createMemory({ root, memory });

    const warnings = [];
    const workspace = await readRuntimeEntries(workspaceRoot);
    const global = await readRuntimeEntries(globalRoot);
    if (workspace.error !== undefined) warnings.push(`${workspace.file} 读不动：${workspace.error}`);
    if (global.error !== undefined) warnings.push(`${global.file} 读不动：${global.error}`);

    // ── 热记忆：runtime → 热记忆 ────────────────────────────────────────────
    const snapshot = await store.snapshot();
    const existing = new Set(snapshot.entries.map((entry) => entry.content.trim()));
    const hot = {
        user: { added: 0, skipped: 0 },
        project: { added: 0, skipped: 0 },
        globalCopied: 0,
    };
    for (const [origin, runtime] of [['workspace', workspace], ['global', global]]) {
        for (const entry of runtime.entries) {
            const content = entry.content.trim();
            if (existing.has(content)) {
                hot[entry.target].skipped += 1;
                continue;
            }
            existing.add(content);
            hot[entry.target].added += 1;
            if (origin === 'global') hot.globalCopied += 1;
            if (!dryRun) {
                await store.mutate({ action: 'add', target: entry.target, content, importance: entry.importance });
            }
        }
    }

    // ── 长期记忆：Memory Spaces → 归档 ──────────────────────────────────────
    const databases = await findInsightDatabases(workspaceRoot);
    if (!dryRun && databases.length === 0) warnings.push(`${join(workspaceRoot, 'data')} 下没有找到 mnemon.db，长期记忆未迁移。`);
    const already = new Set(asArray(await store.archiveItems()).map((item) => asText(item.mnemonId)).filter((id) => id !== ''));
    const archive = { added: 0, skipped: 0, deleted: 0, failed: [], items: [], month: null, file: null };
    for (const database of databases) {
        let payload;
        try {
            payload = await readInsights(database.dbPath, database);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            warnings.push(`读不了 ${database.dbPath}：${message}`);
            archive.failed.push({ file: database.dbPath, error: message });
            continue;
        }
        const edges = asArray(payload?.edges);
        for (const row of asArray(payload?.insights)) {
            const id = asText(row?.id);
            const content = asText(row?.content);
            if (id === '' || content === '') continue;
            if (row?.deleted_at !== null && row?.deleted_at !== undefined) {
                archive.deleted += 1;
                continue;
            }
            if (already.has(id)) {
                archive.skipped += 1;
                continue;
            }
            already.add(id);
            archive.added += 1;
            archive.items.push({
                kind: 'mnemon',
                mnemonId: id,
                content,
                category: asText(row?.category),
                importance: Number.isFinite(Number(row?.importance)) ? Number(row.importance) : undefined,
                tags: parseList(row?.tags).slice(0, 20),
                entities: parseList(row?.entities).slice(0, 20),
                storedAt: asText(row?.stored_at),
                source: asText(row?.source),
                memoryBodyId: database.bodyId,
                memoryBodyName: database.bodyName,
                links: linkCounts(edges, id),
            });
        }
    }
    if (archive.items.length > 0 && !dryRun) {
        const written = await store.appendArchiveBlock({ reason: 'mnemon-migration', items: archive.items });
        archive.month = written.month;
        archive.file = written.file;
    } else if (archive.items.length > 0 && dryRun) {
        archive.month = new Date().toISOString().slice(0, 7);
    }

    // ── Documents：办公记忆没有对应层，只报数量 ─────────────────────────────
    const documents = await countDocuments(workspaceRoot, globalRoot);
    if (documents.total > 0) {
        warnings.push(`mnemon Documents 里有 ${documents.total} 份项目档案，办公记忆没有对应层，本次未迁移（原件仍在 ${documents.files.join('、')}）。`);
    }

    return {
        ok: true,
        dryRun,
        sources: {
            workspace: { dir: workspaceRoot, file: workspace.file, exists: workspace.exists, entries: workspace.entries.length },
            global: { dir: globalRoot, file: global.file, exists: global.exists, entries: global.entries.length },
            databases: databases.map((database) => ({ file: database.dbPath, bodyId: database.bodyId, bodyName: database.bodyName })),
            documents,
        },
        hot,
        archive,
        warnings,
    };
}

/** mnemon Documents 的索引（只统计，不迁移）。 */
async function countDocuments(...roots) {
    const files = [];
    let total = 0;
    for (const root of roots) {
        const file = join(root, 'documents', 'index.json');
        if (!existsSync(file)) continue;
        files.push(file);
        try {
            const parsed = JSON.parse(await readFile(file, 'utf8'));
            total += asArray(parsed?.documents).length;
        } catch {
            // 读不动就不计数
        }
    }
    return { total, files };
}

/** 渲染迁移回执（给模型看）。 */
export function renderMigration(receipt) {
    const lines = [`📒 mnemon → 办公记忆${receipt.dryRun ? '（试运行，未落盘）' : ''}`];
    lines.push(`- 热记忆：新增 ${receipt.hot.user.added + receipt.hot.project.added} 条（用户偏好 ${receipt.hot.user.added} / 项目与环境 ${receipt.hot.project.added}），已存在跳过 ${receipt.hot.user.skipped + receipt.hot.project.skipped} 条。`);
    if (receipt.hot.globalCopied > 0) {
        lines.push(`  其中 ${receipt.hot.globalCopied} 条来自全局 USER.md（~/.mnemon）：已复制进本工作目录；换项目要各迁一次。`);
    }
    lines.push(`- 归档：新增 ${receipt.archive.added} 条长期记忆，跳过 ${receipt.archive.skipped} 条（已迁过），忽略已删除 ${receipt.archive.deleted} 条。`);
    if (receipt.archive.file !== null) lines.push(`  写入 ${receipt.archive.file}`);
    if (receipt.archive.failed.length > 0) {
        for (const item of receipt.archive.failed) lines.push(`  ⚠ ${item.file} 读取失败：${item.error}`);
    }
    if (receipt.sources.documents.total > 0) {
        lines.push(`- Documents：${receipt.sources.documents.total} 份项目档案没有对应层，未迁移（原件仍在原地）。`);
    }
    for (const item of receipt.warnings) lines.push(`- 注意：${item}`);
    lines.push('- 检索迁进来的长期记忆：office_memory({ action: \'read\', layer: \'archive\', query: \'...\' })');
    return lines.join('\n');
}
