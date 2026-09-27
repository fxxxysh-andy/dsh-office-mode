/**
 * 三层记忆：热记忆（偏好与约定）/ 台账（交付物登记）/ 归档（下沉的旧条目）。
 *
 * 这是参考 mnemon 记忆插件的分层做法为本插件重做的一套：层与层的区别不在
 * 于「存了什么」，而在于**什么时候被读到、什么时候被写、装满了往哪儿去**。
 *
 *   层        对应 mnemon        装什么                              怎么读
 *   热记忆    working-context   用户偏好、项目约定（小、常驻）      每次 office_help /
 *                                                                  office_run 反馈顺带投影
 *   台账      narrative         每份交付物一条：路径、格式、主题、    按关键词检索（有界）
 *                               来源、复检统计
 *   归档      durable-evidence  下沉的热记忆与滚动的旧台账（只读）   按关键词检索（有界）
 *
 * 三条与 mnemon 一致的原则：
 *   1. **Markdown 是投影，不是存储**。USER.md / MEMORY.md 由 memory.json 生成，
 *      改记忆只能走 office_memory 工具，直接改 Markdown 会在下次写入时被覆盖。
 *   2. **容量满了向下沉，不静默丢**。热记忆超出上限时，按「重要度低、更旧」的顺序
 *      把条目移进归档，并在 MEMORY.md 里留一条指路条目；台账超出条数上限时，
 *      最旧的一批滚成月度归档摘要。
 *   3. **读是有界的**。一次 read 给出条数上限、单条字符上限与总量上限，
 *      被截断时明确标 truncated —— 不要用「把整本台账灌进上下文」换方便。
 *
 * 第二轮扩展（把用户选定的 mnemon 能力并进来，全部落在工作目录的普通文件里）：
 *   - **多根**：`scope` 决定记忆落在工作区、全局还是两者（跨项目层）；`userScope`
 *     让 USER 侧偏好单独常驻全局，换工作目录不必重记。
 *   - **层拓扑开关**：热记忆 / 台账 / 归档每层可单独关掉，关掉不删数据。
 *   - **图关系**：`links.jsonl` 存双向类型化关系（link / related / 实体视图）。
 *     归档时**保留条目 id**，否则下沉之后关系就断了 —— 这是本轮修掉的一个隐患。
 *   - **召回质量**：strict-v1 按「命中词占比」把结果分成高 / 中 / 未知三档，
 *     低于下限的丢掉，中与未知各给一个名额上限。
 *   - **备份与迁移**：整包导出 / 导入（export / import），幂等合并。
 *
 * 与 mnemon 有意不同的两点（办公插件没有那些设施，硬做只会是假的）：
 *   - **没有远端 provider，也没有授权层**。全部落在工作目录的 `.office/memory/`
 *     里，写入即落盘；不存在「需要 Host 授权才生效」的一步。
 *   - **实体是声明的，不是猜出来的**。mnemon 有 LLM 抽实体那一层，本插件只在
 *     条目显式带 `entities` 时才建索引 —— 猜出来的实体比没有实体更误导。
 *     回合级配额由 `quota.js` 从会话日志里读真实 turn 号来管（见那个文件）。
 *
 * @module dsh-office-mode/memory
 */
import { createHash, randomUUID } from 'node:crypto';
import { existsSync } from 'node:fs';
import { mkdir, readFile, rename, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, isAbsolute, join, posix, relative, sep } from 'node:path';

/** 存储格式版本；将来改结构时用它做迁移判断。 */
export const MEMORY_VERSION = 1;
/** 记忆目录默认位置（相对会话工作目录）。 */
export const DEFAULT_MEMORY_DIR = '.office/memory';
/** 热记忆的两个上限（字节，与 mnemon 的 USER.md / MEMORY.md 一致）。 */
export const DEFAULT_USER_LIMIT_BYTES = 4096;
export const DEFAULT_PROJECT_LIMIT_BYTES = 10240;
/** 台账最多保留多少条，超出的最旧条目滚进归档。 */
export const DEFAULT_LEDGER_LIMIT = 500;
/** 归档最多保留多少个摘要文件，超出的最旧文件被删（归档本身也是有界的）。 */
export const DEFAULT_ARCHIVE_KEEP = 60;
/** 条目分隔符：单独一行的 §，与 mnemon 的条目格式一致，便于人读也便于切分。 */
export const ENTRY_DELIMITER = '\n§\n';
/** 归档指路条目用固定 id，保证「只留一条」，不会每次下沉都堆一条。 */
export const ARCHIVE_POINTER_ID = 'pointer:archive';

/** 记忆目标：用户是谁 / 项目与环境是什么。 */
export const MEMORY_TARGETS = ['user', 'project'];
/**
 * office_memory 认识的全部动作（供文档与设置页列清单；实际校验在各自的分支里）。
 * read 只读；add/replace/remove 改热记忆；log 登记台账；link/unlink/related/entities
 * 管图关系与实体；status 给汇总；export/import 是备份与迁移；migrate 从 mnemon 搬。
 */
export const MEMORY_ACTIONS = ['read', 'add', 'replace', 'remove', 'log', 'status', 'link', 'unlink', 'related', 'entities', 'export', 'import', 'migrate'];
/** 重要度：critical 优先保留，low 最先被下沉。 */
export const MEMORY_IMPORTANCE = ['critical', 'normal', 'low'];
/** 重要度越高越靠前（下沉时从后往前挑）。 */
const IMPORTANCE_RANK = { critical: 0, normal: 1, low: 2 };

/** 记忆范围：只工作区 / 只全局 / 两者（跨项目层）。 */
export const MEMORY_SCOPES = ['workspace', 'global', 'both'];
/** 用户档案范围：跟着当前范围，还是单独常驻全局。 */
export const USER_SCOPES = ['memory', 'global'];

/** 记忆层拓扑的默认值：三层全开。 */
export const DEFAULT_LAYERS = { hot: true, ledger: true, archive: true };

/**
 * 召回质量策略默认值，逐项对齐 mnemon 的 `recallQuality`：
 *   policy             strict-v1 = 分档过滤；off = 只按关键词排序、不丢结果
 *   lowScoreThreshold  命中词占比低于它算「未知」档
 *   highScoreThreshold 达到它算「高分」档，直接采纳
 *   candidateMultiplier 先取 limit×倍数 个候选再分档（防止低分结果挤掉高分结果）
 *   maxMediumResults   中档最多采纳几条
 *   maxUnknownResults  未知档最多采纳几条（0 = 一条都不要）
 */
export const DEFAULT_RECALL_QUALITY = {
    policy: 'strict-v1',
    lowScoreThreshold: 0.25,
    highScoreThreshold: 0.6,
    candidateMultiplier: 3,
    maxMediumResults: 4,
    maxUnknownResults: 2,
};

/**
 * 每回合配额默认值，对齐 mnemon 的回合预算：
 *   recallPerTurn       一个回合里第一次带 query 的记忆检索
 *   recallRefinePerTurn 同一回合里后续的换词细化（mnemon 是「一次初查 + 一次细化」）
 *   relatedPerTurn      一个回合里的图关系遍历
 * 0 = 该类不限制。
 */
export const DEFAULT_QUOTA = {
    recallPerTurn: 1,
    recallRefinePerTurn: 1,
    relatedPerTurn: 1,
};

/** 关系的类型。双向存储，所以不必为反向再写一条。 */
export const LINK_KINDS = ['related', 'refines', 'supersedes', 'contradicts', 'supports', 'derives'];
/** 关系的默认类型。 */
export const DEFAULT_LINK_KIND = 'related';

/** Pack 的文件格式标识与版本（导入时用它判断这份包能不能读）。 */
export const PACK_FORMAT = 'dsh-office-memory-pack';
export const PACK_VERSION = 1;

/** 读取的界：条数与字符。数不宜大 —— 这份上下文是给排版与措辞留余地的。 */
export const READ_LIMITS = {
    ledgerResults: 8,
    archiveResults: 6,
    hotResults: 40,
    entityResults: 12,
    relatedResults: 12,
    itemChars: 400,
    totalChars: 4000,
};

/**
 * 给系统提示用的一段**静态**说明（不含任何记忆内容）。
 *
 * 为什么是静态的：记忆每次写入都会变，而 `systemPrompt.section` 的 text 只接受
 * 同步函数（读盘是异步的）。所以这段只说「记忆在哪、怎么用」，真正的内容由
 * office_help / office_run 的返回值带出来。办公模式的 persona 是 complete 的，
 * 这段会被组装丢掉 —— 它服务于其它模式（原来那层「每轮热记忆投影」由 mnemon 提供，
 * 关掉 mnemon 之后需要有东西接住这件事）。
 */
export const MEMORY_PROMPT_HINT = `工作目录里有一份三层办公记忆（默认在 .office/memory/）：
热记忆记用户偏好与项目约定，台账记这个目录里产出过的交付物，归档是下沉的旧条目。
动手前用 office_memory({ action: 'read' }) 看一眼；用户纠正你、说明偏好或要求
「记住」时，用 office_memory({ action: 'add', target: 'user' | 'project', content }) 记一条。
条目之间可以建关系：office_memory({ action: 'link', sourceId, targetId, kind })，
沿关系找相邻条目用 action:'related'，看实体用 action:'entities'。
office_run 写出的文档会自动登记进台账，不用手工补。`;

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

/** 上限按**字节**算：中文一个字三字节，按字符算会让上限失去意义。 */
export function byteLength(text) {
    return Buffer.byteLength(String(text), 'utf8');
}

function nowIso() {
    return new Date().toISOString();
}

function sha1(text) {
    return createHash('sha1').update(String(text), 'utf8').digest('hex');
}

function asArray(value) {
    return Array.isArray(value) ? value : [];
}

function monthOf(iso) {
    const text = asText(iso);
    return /^\d{4}-\d{2}/.test(text) ? text.slice(0, 7) : nowIso().slice(0, 7);
}

/** 枚举收敛：不认识的值退回默认（配置写坏了不该弄崩插件）。 */
function oneOf(value, allowed, fallback) {
    return allowed.includes(value) ? value : fallback;
}

function boundedInt(value, fallback, min, max) {
    const n = typeof value === 'number' ? Math.trunc(value) : Number.parseInt(String(value), 10);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

/** 把一串字符串收敛成去重、去空白、有长度上限的列表（标签与实体共用）。 */
function stringList(value, maximum = 12) {
    const seen = new Set();
    const out = [];
    for (const raw of asArray(value)) {
        const item = asText(raw);
        if (item === '' || seen.has(item)) continue;
        seen.add(item);
        out.push(item);
        if (out.length >= maximum) break;
    }
    return out;
}

// ── 路径 ────────────────────────────────────────────────────────────────────

/** 一个记忆目录里的全部文件位置。 */
function pathsForDir(dir) {
    return {
        dir,
        memory: join(dir, 'memory.json'),
        ledger: join(dir, 'ledger.jsonl'),
        archive: join(dir, 'archive'),
        user: join(dir, 'USER.md'),
        project: join(dir, 'MEMORY.md'),
        index: join(dir, 'archive', 'index.json'),
        links: join(dir, 'links.jsonl'),
    };
}

/**
 * 解析记忆目录。目录可以给相对路径（相对会话工作目录）或绝对路径。
 * @returns {{dir: string, memory: string, ledger: string, archive: string, user: string, project: string, index: string, links: string}}
 */
export function memoryPaths(root, memoryConfig = {}) {
    const configured = asText(memoryConfig.dir) || DEFAULT_MEMORY_DIR;
    const dir = isAbsolute(configured) ? configured : join(root, configured);
    return pathsForDir(dir);
}

/**
 * 全局层目录（跨项目层）。
 *
 * 默认落在 `$DSH_HOME/.office/memory`，取不到 DSH_HOME 就退到 `~/.dsh/.office/memory`：
 * 这是「用户是谁」这类事实该待的地方 —— 它不该随工作目录搬家。
 *
 * **旧路径兜底**：`.office/` 之前，全局层默认在 `$DSH_HOME/.office-memory`。真实用户
 * 的全局记忆可能还在那里，所以当新目录**不存在**、而旧的 `.office-memory` 目录**存在**
 * 时，返回旧路径 —— 否则一次升级就把用户已经攒下的全局偏好悄悄换成了空目录
 * （表现为「记忆丢了」，而文件其实还在盘上）。新目录一旦被建出来，就只认新路径；
 * 用户想搬，把旧目录整体挪到 `.office/memory` 即可（或者显式配 globalDir）。
 *
 * 配了 `globalDir` 就按配置走（相对路径相对用户主目录，不相对工作目录：
 * 全局层跟着工作目录走就失去意义了）—— 显式配置永远优先，不做任何兜底。
 */
export function globalMemoryDir(memoryConfig = {}) {
    const configured = asText(memoryConfig.globalDir);
    if (configured !== '') {
        return isAbsolute(configured) ? configured : join(homedir(), configured);
    }
    const home = asText(process.env.DSH_HOME) || join(homedir(), '.dsh');
    const current = join(home, '.office', 'memory');
    if (!existsSync(current)) {
        const legacy = join(home, '.office-memory');
        if (existsSync(legacy)) return legacy;
    }
    return current;
}

/** 全局层的文件位置。 */
export function globalMemoryPaths(memoryConfig = {}) {
    return pathsForDir(globalMemoryDir(memoryConfig));
}

/** 把绝对路径渲染成相对会话工作目录的写法（给出目录外路径时原样返回）。 */
function relPath(root, target) {
    const value = relative(root, target);
    if (value === '' || value.startsWith('..')) return target;
    return value.split(sep).join(posix.sep);
}

async function readTextOrNull(path) {
    try {
        return await readFile(path, 'utf8');
    } catch {
        return null;
    }
}

/** 原子写：先写同目录临时文件再改名，避免中途失败留下半截文件。 */
async function writeAtomic(path, text) {
    await mkdir(dirname(path), { recursive: true });
    const tmp = `${path}.${process.pid}.${randomUUID().slice(0, 8)}.tmp`;
    await writeFile(tmp, text, 'utf8');
    try {
        await rename(tmp, path);
    } catch (error) {
        await rm(tmp, { force: true });
        throw error;
    }
}

async function readJsonOrNull(path) {
    const text = await readTextOrNull(path);
    if (text === null || text.trim() === '') return null;
    try {
        return JSON.parse(text);
    } catch {
        return null;
    }
}

// ── 单进程内串行化 ──────────────────────────────────────────────────────────
//
// 读-改-写三段之间不能有第二个写入者插进来，否则后写的那次会吃掉先写的那次
// （mnemon 用 revision 栅栏做这件事）。同一个记忆目录用一条 promise 链排队：
// 进程内够用，也不需要跨进程锁 —— 工作目录里的记忆本来就不该被两个会话同时写。

const queues = new Map();

function withQueue(dir, task) {
    const key = String(dir);
    const previous = queues.get(key) ?? Promise.resolve();
    const next = previous.then(task, task);
    // 队列只用于排队，不用于传播失败：catch 掉，下一次调用照样能进。
    queues.set(key, next.then(() => undefined, () => undefined));
    return next;
}

// ── 存储读写 ────────────────────────────────────────────────────────────────

function normalizeEntry(raw, fallbackTarget = 'project') {
    if (raw === null || typeof raw !== 'object') return null;
    const content = asText(raw.content);
    if (content === '') return null;
    const target = MEMORY_TARGETS.includes(raw.target) ? raw.target : fallbackTarget;
    const importance = MEMORY_IMPORTANCE.includes(raw.importance) ? raw.importance : 'normal';
    return {
        id: asText(raw.id) || `m-${randomUUID().slice(0, 8)}`,
        target,
        content,
        importance,
        entities: stringList(raw.entities),
        tags: stringList(raw.tags),
        createdAt: asText(raw.createdAt) || nowIso(),
        updatedAt: asText(raw.updatedAt) || asText(raw.createdAt) || nowIso(),
    };
}

async function loadStore(paths) {
    const raw = await readJsonOrNull(paths.memory);
    const entries = asArray(raw?.entries).map((entry) => normalizeEntry(entry)).filter(Boolean);
    return {
        version: MEMORY_VERSION,
        revision: asText(raw?.revision) || storeRevision(entries),
        updatedAt: asText(raw?.updatedAt) || nowIso(),
        entries,
    };
}

/** 版本号只由内容决定：同一份条目任何时候算出来都一样，便于比对「有没有变」。 */
function storeRevision(entries) {
    const canonical = entries
        .map((entry) => `${entry.id}\u0000${entry.target}\u0000${entry.importance}\u0000${entry.content}`)
        .sort()
        .join('\n');
    return sha1(canonical);
}

async function saveStore(paths, entries) {
    const store = {
        version: MEMORY_VERSION,
        revision: storeRevision(entries),
        updatedAt: nowIso(),
        entries,
    };
    await writeAtomic(paths.memory, `${JSON.stringify(store, null, 2)}\n`);
    return store;
}

async function loadLedger(paths) {
    const text = await readTextOrNull(paths.ledger);
    if (text === null) return [];
    const records = [];
    for (const line of text.split('\n')) {
        const trimmed = line.trim();
        if (trimmed === '') continue;
        try {
            const parsed = JSON.parse(trimmed);
            if (parsed !== null && typeof parsed === 'object') records.push(parsed);
        } catch {
            // 单行坏掉不该让整本台账读不出来：跳过，保留其余记录。
        }
    }
    return records;
}

function ledgerLine(record) {
    return `${JSON.stringify(record)}\n`;
}

async function appendLedger(paths, records) {
    if (records.length === 0) return;
    await mkdir(paths.dir, { recursive: true });
    // 语义上是追加，实现上是「读回 + 拼接 + 原子改名」：直接 appendFile 在写到一半
    // 失败时会留下半行坏 JSON，而台账是给下一次检索用的，宁可整份重写。
    const existing = await readTextOrNull(paths.ledger);
    const prefix = existing === null || existing === '' ? '' : existing.endsWith('\n') ? existing : `${existing}\n`;
    await writeAtomic(paths.ledger, `${prefix}${records.map(ledgerLine).join('')}`);
}

async function loadArchiveIndex(paths) {
    const raw = await readJsonOrNull(paths.index);
    const digests = asArray(raw?.digests).filter((digest) => digest !== null && typeof digest === 'object');
    return { version: MEMORY_VERSION, updatedAt: asText(raw?.updatedAt) || nowIso(), digests };
}

async function saveArchiveIndex(paths, index) {
    await writeAtomic(paths.index, `${JSON.stringify(index, null, 2)}\n`);
}

// ── 图关系（links.jsonl） ───────────────────────────────────────────────────

function normalizeLink(raw) {
    if (raw === null || typeof raw !== 'object') return null;
    const sourceId = asText(raw.sourceId);
    const targetId = asText(raw.targetId);
    if (sourceId === '' || targetId === '' || sourceId === targetId) return null;
    return {
        id: asText(raw.id) || `K-${randomUUID().slice(0, 8)}`,
        sourceId,
        targetId,
        kind: LINK_KINDS.includes(raw.kind) ? raw.kind : DEFAULT_LINK_KIND,
        note: asText(raw.note),
        at: asText(raw.at) || nowIso(),
    };
}

async function loadLinks(paths) {
    const text = await readTextOrNull(paths.links);
    if (text === null) return [];
    const out = [];
    for (const line of text.split('\n')) {
        const trimmed = line.trim();
        if (trimmed === '') continue;
        try {
            const link = normalizeLink(JSON.parse(trimmed));
            if (link !== null) out.push(link);
        } catch {
            // 同台账：单行坏掉跳过，其余照读。
        }
    }
    return out;
}

async function appendLinks(paths, links) {
    if (links.length === 0) return;
    await mkdir(paths.dir, { recursive: true });
    const existing = await readTextOrNull(paths.links);
    const prefix = existing === null || existing === '' ? '' : existing.endsWith('\n') ? existing : `${existing}\n`;
    await writeAtomic(paths.links, `${prefix}${links.map((link) => `${JSON.stringify(link)}\n`).join('')}`);
}

// ── 投影（USER.md / MEMORY.md） ─────────────────────────────────────────────

function targetUsage(entries, target) {
    const items = entries.filter((entry) => entry.target === target);
    return {
        target,
        entryCount: items.length,
        used: items.reduce((sum, entry) => sum + byteLength(entry.content) + byteLength(ENTRY_DELIMITER), 0),
    };
}

/**
 * 把一类条目渲染成 Markdown 投影。
 *
 * 头部写明「本文件是投影不是存储」：用户或模型直接编辑这里是最容易发生的
 * 误用（写完下次 mutation 就被覆盖），所以在文件里就写清楚。
 */
export function renderProjection(target, entries, { limit, used }) {
    const title = target === 'user' ? '办公记忆 · 用户偏好' : '办公记忆 · 项目与环境';
    const hint = target === 'user'
        ? "改这条用 office_memory({ action: 'add' | 'replace' | 'remove', target: 'user', ... })"
        : "改这条用 office_memory({ action: 'add' | 'replace' | 'remove', target: 'project', ... })";
    const lines = [
        `# ${title}`,
        '',
        '> 本文件由 dsh-office-mode 生成，是**投影**不是存储。',
        `> ${hint}；直接编辑这里会在下次写入时被覆盖。`,
        '',
    ];
    if (entries.length === 0) {
        lines.push('（还没有记下任何条目。）', '');
    } else {
        for (const entry of entries) {
            lines.push(`- [${entry.importance}] ${entry.content.replace(/\s*\n\s*/g, ' ')}`);
        }
        lines.push('');
    }
    lines.push(`用量 ${used} / ${limit} 字节，共 ${entries.length} 条。`);
    return `${lines.join('\n')}\n`;
}

async function writeProjections(paths, entries, memoryConfig) {
    const userLimit = memoryConfig.userLimitBytes;
    const projectLimit = memoryConfig.projectLimitBytes;
    const userUsage = targetUsage(entries, 'user');
    const projectUsage = targetUsage(entries, 'project');
    await writeAtomic(paths.user, renderProjection('user', entries.filter((entry) => entry.target === 'user'), { limit: userLimit, used: userUsage.used }));
    await writeAtomic(paths.project, renderProjection('project', entries.filter((entry) => entry.target === 'project'), { limit: projectLimit, used: projectUsage.used }));
    return {
        user: { ...userUsage, limit: userLimit, markdownPath: paths.user },
        project: { ...projectUsage, limit: projectLimit, markdownPath: paths.project },
    };
}

// ── 归档 ────────────────────────────────────────────────────────────────────

function archiveSection(reason) {
    return reason === 'hot-overflow' ? '热记忆下沉'
        : reason === 'ledger-rollover' ? '台账滚动'
            : reason === 'mnemon-migration' ? 'mnemon 迁移'
                : reason === 'pack-import' ? 'Pack 导入'
                    : reason;
}

function renderArchiveDigest(digest) {
    const lines = [`# 归档 ${digest.month}`, ''];
    for (const block of digest.blocks) {
        lines.push(`## ${block.at} · ${archiveSection(block.reason)}（${block.items.length} 条）`, '');
        for (const item of block.items) {
            if (item.kind === 'hot') {
                lines.push(`- [${item.origin}/${item.importance}] ${String(item.content).replace(/\s*\n\s*/g, ' ')}`);
            } else if (item.kind === 'mnemon') {
                const badge = [item.category, item.importance === undefined ? '' : `重要度 ${item.importance}`].filter(Boolean).join(' · ');
                const tags = asArray(item.tags).slice(0, 6);
                const entities = asArray(item.entities).slice(0, 6);
                const linked = item.linkedCount === undefined ? '' : `（关系边 ${item.linkedCount} 条）`;
                lines.push(`- [${badge}]${linked} ${String(item.content).replace(/\s*\n\s*/g, ' ')}`);
                if (tags.length > 0) lines.push(`  - 标签：${tags.join('、')}`);
                if (entities.length > 0) lines.push(`  - 实体：${entities.join('、')}`);
            } else {
                const badge = [item.format, item.theme].filter(Boolean).join(' · ');
                lines.push(`- ${item.at.slice(0, 10)} ${item.path}${badge ? ` [${badge}]` : ''}${item.purpose ? ` —— ${item.purpose}` : ''}`);
            }
        }
        lines.push('');
    }
    lines.push('> 归档是只读的长期层：热记忆装不下时向下沉，台账超出条数上限时最旧的滚进来，');
    lines.push('> 从 mnemon 迁来的长期记忆也在这里（带类别、重要度、标签与实体）。');
    lines.push('> 检索：office_memory({ action: \'read\', layer: \'archive\', query: \'...\' })');
    return `${lines.join('\n')}\n`;
}

/** 归档块的稳定标识：老数据没有 id，就用「时刻+原因+条数」当键（导入去重要用）。 */
function blockKey(block) {
    return asText(block.id) || `${asText(block.at)}|${asText(block.reason)}|${asArray(block.items).length}`;
}

/**
 * 追加一批归档：同一个月份合并进一个摘要文件。
 *
 * 归档文件是给人看的 Markdown（可以 grep、可以打开），索引是给检索用的 JSON，
 * 两者同一次写入更新 —— 只更新索引会出现「检索得到、文件里没有」的假象。
 */
async function appendArchive(paths, blocks, memoryConfig) {
    if (blocks.length === 0) return [];
    await mkdir(paths.archive, { recursive: true });
    const index = await loadArchiveIndex(paths);
    const created = [];
    for (const block of blocks) {
        const month = block.month ?? monthOf(block.at);
        let digest = index.digests.find((item) => item.month === month);
        if (digest === undefined) {
            digest = { id: `a-${month}-${randomUUID().slice(0, 6)}`, month, at: block.at, file: `archive/${month}.md`, blocks: [] };
            index.digests.push(digest);
        }
        digest.blocks.push({ id: asText(block.id) || `b-${randomUUID().slice(0, 8)}`, at: block.at, reason: block.reason, items: block.items });
        digest.at = block.at;
        created.push(digest);
    }
    // 归档本身也封顶：最旧的摘要文件被删掉，索引里同步移除，不留「指不到的条目」。
    // 排序键必须是 month 而不是 at：一个月只产生一个摘要，而 at 是「这次下沉发生的
    // 时刻」——同一秒内的多个摘要 at 完全相同，按它排序等于按写入顺序裁剪，剪掉哪
    // 一个是不确定的。
    const keep = memoryConfig.archiveKeep;
    index.digests.sort((left, right) => String(left.month).localeCompare(String(right.month)));
    const dropped = index.digests.length > keep ? index.digests.splice(0, index.digests.length - keep) : [];
    for (const digest of dropped) {
        await rm(join(paths.archive, `${digest.month}.md`), { force: true });
    }
    for (const digest of new Set(created.map((item) => item.month))) {
        const target = index.digests.find((item) => item.month === digest);
        if (target !== undefined) await writeAtomic(join(paths.archive, `${digest}.md`), renderArchiveDigest(target));
    }
    index.updatedAt = nowIso();
    await saveArchiveIndex(paths, index);
    return created;
}

/**
 * 把一条热记忆切成检索项。
 *
 * **id 必须带过去**：热记忆下沉进归档之后，图关系要还能按 id 找到它。
 * 第一版把 id 丢了，于是「下沉」等于「关系断掉」——本轮修掉。
 */
function hotArchiveItem(entry) {
    return {
        id: entry.id,
        kind: 'hot',
        origin: entry.target,
        importance: entry.importance,
        content: entry.content,
        entities: stringList(entry.entities),
        tags: stringList(entry.tags),
        at: entry.updatedAt,
    };
}

function ledgerArchiveItem(record) {
    const outline = asArray(record.outline).slice(0, 6).join(' / ');
    return {
        id: asText(record.id),
        kind: 'ledger',
        at: asText(record.at),
        path: asText(record.path),
        format: asText(record.format),
        theme: asText(record.theme),
        purpose: asText(record.purpose),
        outline,
    };
}

function ledgerSearchText(record) {
    const parts = [
        record.path,
        record.purpose,
        record.format,
        record.theme,
        asArray(record.outline).join(' '),
        asArray(record.source).join(' '),
        record.note,
    ];
    return parts.filter((part) => typeof part === 'string' && part !== '').join(' ');
}

function archiveSearchText(item) {
    if (item.kind === 'hot') return `${item.origin} ${item.importance} ${item.content} ${asArray(item.entities).join(' ')} ${asArray(item.tags).join(' ')}`;
    if (item.kind === 'mnemon') {
        return [item.category, asArray(item.tags).join(' '), asArray(item.entities).join(' '), item.content]
            .filter((part) => typeof part === 'string' && part !== '')
            .join(' ');
    }
    return `${item.path} ${item.purpose} ${item.format} ${item.theme} ${item.outline}`;
}

/** 归档条目在模型侧的一行文本（不同来源的条目各有各的字段）。 */
function archiveItemText(item) {
    if (item.kind === 'hot') return String(item.content ?? '');
    if (item.kind === 'mnemon') return String(item.content ?? '');
    return [item.path, item.purpose].filter(Boolean).join(' —— ');
}

/** 关键词命中数（大小写不敏感的整串包含），用来排序与过滤。 */
function matchScore(text, terms) {
    const haystack = String(text).toLowerCase();
    let score = 0;
    for (const term of terms) {
        if (term !== '' && haystack.includes(term)) score += 1;
    }
    return score;
}

function queryTerms(query) {
    return String(query ?? '')
        .toLowerCase()
        .split(/[\s,，、;；|/]+/)
        .map((term) => term.trim())
        .filter((term) => term !== '');
}

function truncate(text, maximum) {
    const value = String(text);
    if (value.length <= maximum) return value;
    return `${value.slice(0, Math.max(1, maximum - 1))}…`;
}

// ── 召回质量（strict-v1） ───────────────────────────────────────────────────

/**
 * 把「命中词数」换算成 0..1 的相关度：命中了查询里几成的词。
 *
 * mnemon 的阈值是给向量相似度用的；这里没有向量，就用「命中占比」当同构的量 ——
 * 阈值语义（低分丢掉、中档限名额、高分直采）完全一致，只是打分函数不同。
 * 这一点必须写在文档里，不能让人以为 0.6 是余弦相似度。
 */
export function relevanceOf(score, termCount) {
    if (termCount <= 0) return 1;
    return score / termCount;
}

/**
 * 按 strict-v1 分档过滤。输入是已按 score 降序排好的候选。
 *
 * 顺序上有个关键点：**先取候选池再分档**。如果直接按 limit 截断再分档，
 * 低分结果会把高分结果挤出候选池（candidateMultiplier 存在的理由就是这个）。
 */
export function applyRecallQuality(candidates, { quality, limit, termCount }) {
    if (quality === undefined || quality.policy === 'off' || termCount <= 0) {
        return { items: candidates.slice(0, limit), dropped: 0, medium: 0, unknown: 0 };
    }
    const pool = candidates.slice(0, Math.max(limit, limit * quality.candidateMultiplier));
    const items = [];
    let medium = 0;
    let unknown = 0;
    let dropped = 0;
    for (const candidate of pool) {
        if (items.length >= limit) break;
        const relevance = relevanceOf(candidate.score, termCount);
        if (relevance >= quality.highScoreThreshold) {
            items.push(candidate);
            continue;
        }
        if (relevance >= quality.lowScoreThreshold) {
            if (medium < quality.maxMediumResults) {
                items.push(candidate);
                medium += 1;
            } else {
                dropped += 1;
            }
            continue;
        }
        if (unknown < quality.maxUnknownResults) {
            items.push(candidate);
            unknown += 1;
        } else {
            dropped += 1;
        }
    }
    return { items, dropped, medium, unknown };
}

// ── 有界读取 ────────────────────────────────────────────────────────────────

/**
 * 有界采纳：条数、单条字符、总字符三重上限，外加去重。
 *
 * 与 mnemon 的 boundedModelInsights 同一个目的：宁可给少而准的几条，
 * 也不要把一整本台账倒进上下文把排版与措辞的余地从上下文里挤走。
 */
function admit(items, { resultLimit, itemChars, totalChars, limits }) {
    const seen = new Set();
    const admitted = [];
    let used = 0;
    for (const item of items) {
        if (admitted.length >= resultLimit) break;
        const digest = sha1(String(item.text).trim().replace(/\s+/g, ' '));
        if (seen.has(digest)) continue;
        const remaining = totalChars - used;
        if (remaining <= 0) break;
        const text = truncate(item.text, Math.min(itemChars, remaining));
        if (text === '') continue;
        seen.add(digest);
        used += text.length;
        admitted.push({ ...item, text, truncated: text !== item.text });
    }
    return { items: admitted, used, truncated: admitted.length < items.length || admitted.some((item) => item.truncated) };
}

// ── 对外：一个记忆实例 ──────────────────────────────────────────────────────

/**
 * 把 memory 配置补全成运行期用的那一份（缺省永远可用）。
 * @returns {{dir, userLimitBytes, projectLimitBytes, ledgerLimit, archiveKeep, autoLedger, layers, scope, globalDir, userScope, links, autoCapture, recallQuality, quota}}
 */
export function resolveMemoryConfig(memory = {}) {
    const raw = memory !== null && typeof memory === 'object' ? memory : {};
    const rawLayers = raw.layers !== null && typeof raw.layers === 'object' ? raw.layers : {};
    const layers = {};
    for (const [name, fallback] of Object.entries(DEFAULT_LAYERS)) {
        layers[name] = typeof rawLayers[name] === 'boolean' ? rawLayers[name] : fallback;
    }
    const rawQuality = raw.recallQuality !== null && typeof raw.recallQuality === 'object' ? raw.recallQuality : {};
    const quality = {
        policy: oneOf(rawQuality.policy, ['strict-v1', 'off'], DEFAULT_RECALL_QUALITY.policy),
        lowScoreThreshold: Number.isFinite(rawQuality.lowScoreThreshold) ? rawQuality.lowScoreThreshold : DEFAULT_RECALL_QUALITY.lowScoreThreshold,
        highScoreThreshold: Number.isFinite(rawQuality.highScoreThreshold) ? rawQuality.highScoreThreshold : DEFAULT_RECALL_QUALITY.highScoreThreshold,
        candidateMultiplier: boundedInt(rawQuality.candidateMultiplier, DEFAULT_RECALL_QUALITY.candidateMultiplier, 1, 10),
        maxMediumResults: boundedInt(rawQuality.maxMediumResults, DEFAULT_RECALL_QUALITY.maxMediumResults, 0, 40),
        maxUnknownResults: boundedInt(rawQuality.maxUnknownResults, DEFAULT_RECALL_QUALITY.maxUnknownResults, 0, 40),
    };
    const rawQuota = raw.quota !== null && typeof raw.quota === 'object' ? raw.quota : {};
    const quota = {};
    for (const [name, fallback] of Object.entries(DEFAULT_QUOTA)) {
        quota[name] = boundedInt(rawQuota[name], fallback, 0, 50);
    }
    return {
        dir: asText(raw.dir) || DEFAULT_MEMORY_DIR,
        userLimitBytes: Number.isFinite(raw.userLimitBytes) ? raw.userLimitBytes : DEFAULT_USER_LIMIT_BYTES,
        projectLimitBytes: Number.isFinite(raw.projectLimitBytes) ? raw.projectLimitBytes : DEFAULT_PROJECT_LIMIT_BYTES,
        ledgerLimit: Number.isFinite(raw.ledgerLimit) ? raw.ledgerLimit : DEFAULT_LEDGER_LIMIT,
        archiveKeep: Number.isFinite(raw.archiveKeep) ? raw.archiveKeep : DEFAULT_ARCHIVE_KEEP,
        autoLedger: raw.autoLedger !== false,
        layers,
        scope: oneOf(raw.scope, MEMORY_SCOPES, 'workspace'),
        globalDir: asText(raw.globalDir),
        userScope: oneOf(raw.userScope, USER_SCOPES, 'memory'),
        links: raw.links !== false,
        autoCapture: raw.autoCapture !== false,
        recallQuality: quality,
        quota,
    };
}

/**
 * 算出这次要开哪几个记忆根。
 *
 *   scope      workspace → 只工作区；global → 只全局；both → 两个都开
 *   userScope  global 时强制把全局根加进来（USER 侧偏好要常驻全局）
 *
 * 顺序固定「全局在前、工作区在后」：读的时候全局偏好先出现，工作区约定覆盖它 ——
 * 具体项目的约定比通用偏好更该被最后看到。
 */
function buildStores(root, config) {
    const stores = [];
    const wantGlobal = config.scope === 'global' || config.scope === 'both' || config.userScope === 'global';
    const wantWorkspace = config.scope !== 'global';
    if (wantGlobal) stores.push({ id: 'global', paths: globalMemoryPaths(config) });
    if (wantWorkspace) stores.push({ id: 'workspace', paths: memoryPaths(root, config) });
    if (stores.length === 0) stores.push({ id: 'workspace', paths: memoryPaths(root, config) });
    return stores;
}

/**
 * 合并多个根的 revision。
 *
 * 单根时**必须原样返回那一份 revision**：它是「内容算出来的稳定指纹」这条契约
 * 的一部分（同一份条目任何时候算出来都一样），外面拿它比对「有没有变」。
 * 多根时才退化成一个组合指纹 —— 这时它只用来判断「整体有没有变」。
 */
function combineRevisions(items) {
    if (items.length === 0) return sha1('');
    if (items.length === 1) return items[0].revision;
    return sha1(items.map((item) => `${item.origin}:${item.revision}`).join('|'));
}

/**
 * 建一个记忆实例。所有读写都通过它，路径与上限在构造时定死。
 *
 * @param {{root: string, memory?: object}} options
 *   root   会话工作目录（路径解析的基准）
 *   memory 已解析的记忆配置（见 config.js 的 memory 段）
 */
export function createMemory({ root, memory = {} }) {
    const base = asText(root);
    if (base === '') throw new Error('createMemory 需要会话工作目录。');
    const config = resolveMemoryConfig(memory);
    const stores = buildStores(base, config);
    /** 工作区根（默认的那一个）。`paths` 保持指向它，兼容既有调用方。 */
    const paths = (stores.find((store) => store.id === 'workspace') ?? stores[0]).paths;
    const layers = config.layers;

    /** 某个目标该写进哪个根：USER 侧可以单独常驻全局，其余看 scope。 */
    function storeFor(target) {
        if (target === 'user' && config.userScope === 'global') {
            const global = stores.find((store) => store.id === 'global');
            if (global !== undefined) return global;
        }
        if (config.scope === 'global') {
            return stores.find((store) => store.id === 'global') ?? stores[0];
        }
        return stores.find((store) => store.id === 'workspace') ?? stores[0];
    }

    /** 台账这类「项目产物」该落在工作区；纯全局范围才落全局。 */
    function defaultStore() {
        if (config.scope === 'global') return stores.find((store) => store.id === 'global') ?? stores[0];
        return stores.find((store) => store.id === 'workspace') ?? stores[0];
    }

    /** 读一份快照：热记忆全量 + 用量。热记忆本来就小，全给。 */
    async function snapshot() {
        const entries = [];
        const perStore = [];
        for (const store of stores) {
            if (!layers.hot) break;
            const loaded = await loadStore(store.paths);
            entries.push(...loaded.entries.map((entry) => ({ ...entry, origin: store.id })));
            perStore.push({ origin: store.id, revision: loaded.revision, updatedAt: loaded.updatedAt });
        }
        return {
            revision: combineRevisions(perStore),
            updatedAt: nowIso(),
            entries,
            targets: {
                user: { ...targetUsage(entries, 'user'), limit: config.userLimitBytes },
                project: { ...targetUsage(entries, 'project'), limit: config.projectLimitBytes },
            },
            stores: perStore,
            paths,
        };
    }

    /**
     * 三层读取。layer 决定读哪几层，query 为空时按时间倒序取最近的。
     *
     * 层拓扑关掉的层直接跳过（关掉不删数据，打开就回来）。多根时结果合并，
     * 每条带 `origin` 标出它来自全局还是工作区。
     *
     * @returns {Promise<object>} 紧凑结果（渲染由 renderRead 负责）
     */
    async function read({ layer = 'all', query = '', limit = 0 } = {}) {
        const wanted = layer === 'hot' || layer === 'ledger' || layer === 'archive' ? layer : 'all';
        const terms = queryTerms(query);
        const quality = config.recallQuality;
        const result = { layer: wanted, query: asText(query), truncated: false, origins: stores.map((store) => store.id), quality: quality.policy };
        // layer:'all' 时三层共用一份字符预算（mnemon 也是「一个信封、全层共用」）；
        // 只读单层时这一层独占预算，好让一次精确查询能拿满。
        const shared = wanted === 'all';
        let remaining = READ_LIMITS.totalChars;
        const budget = () => (shared ? remaining : READ_LIMITS.totalChars);
        const cap = (want, hard) => (limit > 0 ? Math.min(limit, hard) : hard);

        if (layers.hot && (wanted === 'hot' || wanted === 'all')) {
            const all = [];
            const storeUsage = [];
            for (const store of stores) {
                const loaded = await loadStore(store.paths);
                for (const entry of loaded.entries) all.push({ ...entry, origin: store.id });
                storeUsage.push({ origin: store.id, revision: loaded.revision });
            }
            const admitted = admit(
                all
                    .slice()
                    .sort((left, right) => (IMPORTANCE_RANK[left.importance] - IMPORTANCE_RANK[right.importance]) || left.updatedAt.localeCompare(right.updatedAt))
                    .map((entry) => ({ id: entry.id, target: entry.target, importance: entry.importance, updatedAt: entry.updatedAt, origin: entry.origin, entities: entry.entities, text: entry.content, digest: sha1(entry.content) })),
                { resultLimit: cap('hot', READ_LIMITS.hotResults), itemChars: READ_LIMITS.itemChars, totalChars: budget() },
            );
            remaining = Math.max(0, remaining - admitted.used);
            result.hot = {
                items: admitted.items,
                revision: combineRevisions(storeUsage),
                targets: {
                    user: { ...targetUsage(all, 'user'), limit: config.userLimitBytes },
                    project: { ...targetUsage(all, 'project'), limit: config.projectLimitBytes },
                },
            };
            result.truncated = result.truncated || admitted.truncated;
        }

        if (layers.ledger && (wanted === 'ledger' || wanted === 'all')) {
            const records = [];
            for (const store of stores) {
                for (const record of await loadLedger(store.paths)) records.push({ ...record, origin: store.id });
            }
            const scored = records
                .map((record) => ({ record, score: terms.length === 0 ? 0 : matchScore(ledgerSearchText(record), terms) }))
                .filter((item) => terms.length === 0 || item.score > 0)
                .sort((left, right) => (right.score - left.score) || String(right.record.at).localeCompare(String(left.record.at)));
            const picked = applyRecallQuality(scored, { quality, limit: cap('ledger', READ_LIMITS.ledgerResults), termCount: terms.length });
            const admitted = admit(
                picked.items.map((item) => ({
                    id: asText(item.record.id),
                    at: asText(item.record.at),
                    path: asText(item.record.path),
                    format: asText(item.record.format),
                    theme: asText(item.record.theme),
                    purpose: asText(item.record.purpose),
                    origin: item.record.origin,
                    score: item.score,
                    text: [asText(item.record.path), asText(item.record.purpose), asArray(item.record.outline).slice(0, 4).join(' / ')].filter(Boolean).join(' —— '),
                })),
                { resultLimit: cap('ledger', READ_LIMITS.ledgerResults), itemChars: READ_LIMITS.itemChars, totalChars: budget() },
            );
            remaining = Math.max(0, remaining - admitted.used);
            result.ledger = {
                items: admitted.items,
                total: records.length,
                matched: scored.length,
                qualityDropped: picked.dropped,
                relativeDir: relPath(base, paths.dir),
            };
            result.truncated = result.truncated || admitted.truncated || scored.length > admitted.items.length;
        }

        if (layers.archive && (wanted === 'archive' || wanted === 'all')) {
            const items = [];
            let digestCount = 0;
            for (const store of stores) {
                const index = await loadArchiveIndex(store.paths);
                digestCount += index.digests.length;
                for (const digest of index.digests) {
                    for (const block of asArray(digest.blocks)) {
                        for (const item of asArray(block.items)) items.push({ ...item, origin: store.id, month: digest.month, file: digest.file });
                    }
                }
            }
            const scored = items
                .map((item) => ({ item, score: terms.length === 0 ? 0 : matchScore(archiveSearchText(item), terms) }))
                .filter((entry) => terms.length === 0 || entry.score > 0)
                .sort((left, right) => (right.score - left.score) || String(right.item.at ?? '').localeCompare(String(left.item.at ?? '')));
            const picked = applyRecallQuality(scored, { quality, limit: cap('archive', READ_LIMITS.archiveResults), termCount: terms.length });
            const admitted = admit(
                picked.items.map((entry) => ({
                    id: entry.item.id,
                    month: entry.item.month,
                    file: entry.item.file,
                    kind: entry.item.kind,
                    at: entry.item.at,
                    origin: entry.item.origin,
                    importance: entry.item.importance,
                    category: entry.item.category,
                    entities: entry.item.entities,
                    tags: entry.item.tags,
                    path: entry.item.path,
                    score: entry.score,
                    text: archiveItemText(entry.item),
                })),
                { resultLimit: cap('archive', READ_LIMITS.archiveResults), itemChars: READ_LIMITS.itemChars, totalChars: budget() },
            );
            remaining = Math.max(0, remaining - admitted.used);
            result.archive = {
                items: admitted.items,
                digests: digestCount,
                items_total: items.length,
                matched: scored.length,
                qualityDropped: picked.dropped,
                relativeDir: relPath(base, paths.dir),
            };
            result.truncated = result.truncated || admitted.truncated || scored.length > admitted.items.length;
        }

        result.qualityDroppedTotal = (result.ledger?.qualityDropped ?? 0) + (result.archive?.qualityDropped ?? 0);
        result.paths = {
            memory: relPath(base, paths.memory),
            user: relPath(base, paths.user),
            project: relPath(base, paths.project),
            ledger: relPath(base, paths.ledger),
            archive: relPath(base, paths.archive),
            links: relPath(base, paths.links),
        };
        result.globalDir = stores.some((store) => store.id === 'global') ? relPath(base, globalMemoryPaths(config).dir) : '';
        return result;
    }

    /**
     * 改热记忆：add / replace / remove。
     *
     * remove 只在用户明确要求或确有证据时用 —— 这条写在工具描述里，不在这里拦
     * （拦不住「模型认为没用了」，但可以要求它给出 oldText 而不是下标）。
     */
    async function mutate({ action, target = 'project', content = '', oldText = '', importance = 'normal', entities = [], tags = [] } = {}) {
        const act = asText(action);
        if (!['add', 'replace', 'remove'].includes(act)) throw new Error(`记忆动作只支持 add / replace / remove，收到「${act}」。`);
        // 层开关是「不再读 / 不再写」，不是「只读」：关掉热记忆之后还能 add，
        // 就会让人以为记忆坏了（写了读不到）。已有条目不受影响，打开就回来。
        if (!layers.hot) throw new Error('热记忆层已在设置里关掉（memory.layers.hot），本次没有写入。要写入请先把它打开；已有条目不会被删除。');
        const realTarget = asText(target);
        if (!MEMORY_TARGETS.includes(realTarget)) throw new Error(`记忆目标只支持 user / project，收到「${realTarget}」。`);
        const wanted = asText(content);
        const needle = asText(oldText);
        if (act !== 'remove' && wanted === '') throw new Error('记忆内容不能为空。');
        if (act !== 'add' && needle === '') throw new Error(`${act} 需要 oldText 指出改哪一条（给一段能唯一命中的原文）。`);
        if (act !== 'remove' && byteLength(wanted) > (realTarget === 'user' ? config.userLimitBytes : config.projectLimitBytes)) {
            throw new Error(`这条记忆 ${byteLength(wanted)} 字节，比 ${realTarget} 侧的容量上限还大，存不下。请压缩到一两句。`);
        }
        const realImportance = MEMORY_IMPORTANCE.includes(importance) ? importance : 'normal';
        const store = storeFor(realTarget);

        return withQueue(store.paths.dir, async () => {
            const loaded = await loadStore(store.paths);
            const entries = loaded.entries.slice();
            // replace / remove 也要能在**另一个根**里命中：多根时用户记不清
            // 那条偏好当初落在全局还是工作区，按当前根找不到就跨根找一遍。
            let owner = store;
            let matched = needle === '' ? [] : entries.filter((entry) => entry.content.includes(needle));
            if (act !== 'add' && matched.length === 0) {
                for (const other of stores) {
                    if (other.id === store.id) continue;
                    const candidate = await loadStore(other.paths);
                    const hits = candidate.entries.filter((entry) => entry.content.includes(needle));
                    if (hits.length > 0) {
                        owner = other;
                        matched = hits;
                        break;
                    }
                }
            }
            const ownerEntries = owner.id === store.id ? entries : (await loadStore(owner.paths)).entries.slice();
            let touched = null;
            let previous = null;

            if (act === 'add') {
                touched = normalizeEntry({ target: realTarget, content: wanted, importance: realImportance, entities, tags, createdAt: nowIso(), updatedAt: nowIso() });
                entries.push(touched);
            } else {
                if (matched.length === 0) throw new Error(`没有哪条记忆包含「${truncate(needle, 60)}」。先用 action:'read' 看现在的条目。`);
                if (matched.length > 1) {
                    // 改和删都要求唯一命中：命中多条时按顺序取第一条，可能改错或删错的是
                    // 用户真正在意的那条，而记忆里的删除是不可见的破坏。
                    throw new Error(`「${truncate(needle, 60)}」命中 ${matched.length} 条，不唯一。oldText 给长一点再试。`);
                }
                const hit = matched[0];
                previous = { ...hit };
                if (act === 'remove') {
                    ownerEntries.splice(ownerEntries.indexOf(hit), 1);
                    touched = hit;
                } else {
                    hit.content = wanted;
                    hit.target = realTarget;
                    hit.importance = realImportance;
                    if (stringList(entities).length > 0) hit.entities = stringList(entities);
                    if (stringList(tags).length > 0) hit.tags = stringList(tags);
                    hit.updatedAt = nowIso();
                    touched = hit;
                }
            }

            const maintenance = await enforceCapacity(owner.paths, ownerEntries, config);
            const saved = await saveStore(owner.paths, ownerEntries);
            const targets = await writeProjections(owner.paths, ownerEntries, config);
            // 刚写入的那条自己也可能是被下沉的那个（它比现有条目更次要）。
            // 这件事必须如实回执，不能让模型以为「记下了」。remove 不算「被下沉」：
            // 删掉本来就是这次调用的目的。
            const evicted = act !== 'remove' && !ownerEntries.some((entry) => entry.id === touched.id);
            return {
                action: act,
                target: realTarget,
                entry: touched,
                previous,
                evicted,
                revision: saved.revision,
                entryCount: ownerEntries.length,
                origin: owner.id,
                usage: { user: targets.user, project: targets.project },
                maintenance,
            };
        });
    }

    /**
     * 登记台账：一次 office_run 写出的每个 Office 文件一条。
     * 检索键用「相对工作目录的路径」，所以同一份文件反复生成会留下多版记录，
     * 但不会混进别的项目。
     */
    async function log(records) {
        const clean = asArray(records)
            .map((record) => ({
                id: `L-${randomUUID().slice(0, 8)}`,
                at: asText(record.at) || nowIso(),
                path: asText(record.path),
                format: asText(record.format),
                theme: asText(record.theme),
                bytes: Number.isFinite(record.bytes) ? record.bytes : 0,
                purpose: asText(record.purpose),
                outline: asArray(record.outline).slice(0, 12).map((item) => String(item)),
                stats: record.stats !== null && typeof record.stats === 'object' ? record.stats : null,
                source: asArray(record.source).slice(0, 8).map((item) => String(item)),
                note: asText(record.note),
            }))
            .filter((record) => record.path !== '');
        if (clean.length === 0) return { added: 0, kept: 0, rolled: 0 };
        if (!layers.ledger) return { added: 0, kept: 0, rolled: 0, disabled: true };
        const store = defaultStore();

        return withQueue(store.paths.dir, async () => {
            await appendLedger(store.paths, clean);
            const all = await loadLedger(store.paths);
            const limit = Math.max(1, config.ledgerLimit);
            if (all.length <= limit) return { added: clean.length, kept: all.length, rolled: 0, origin: store.id };

            const overflow = all.slice(0, all.length - limit);
            const kept = all.slice(all.length - limit);
            const month = monthOf(overflow[0]?.at);
            await appendArchive(store.paths, [{
                at: nowIso(),
                month,
                reason: 'ledger-rollover',
                items: overflow.map(ledgerArchiveItem),
            }], config);
            await writeAtomic(store.paths.ledger, kept.map(ledgerLine).join(''));
            return { added: clean.length, kept: kept.length, rolled: overflow.length, origin: store.id };
        });
    }

    /**
     * 给提示词/工具反馈用的紧凑投影：热记忆全文 + 台账近况。
     *
     * 办公模式的 persona 是 complete 的（唯一的系统提示），插件注册的提示段落
     * 会被丢掉 —— 所以「让模型在动笔前看到记忆」只能挂在它会调的调用上：
     * office_help 与 office_run 的反馈。
     */
    async function digest({ recentLedger = 3 } = {}) {
        const entries = [];
        const records = [];
        if (layers.hot) {
            for (const store of stores) {
                for (const entry of (await loadStore(store.paths)).entries) entries.push({ ...entry, origin: store.id });
            }
        }
        if (layers.ledger) {
            for (const store of stores) {
                for (const record of await loadLedger(store.paths)) records.push({ ...record, origin: store.id });
            }
        }
        const userEntries = entries.filter((entry) => entry.target === 'user');
        const projectEntries = entries.filter((entry) => entry.target === 'project');
        return {
            empty: entries.length === 0 && records.length === 0,
            user: userEntries,
            project: projectEntries,
            usage: {
                user: { ...targetUsage(entries, 'user'), limit: config.userLimitBytes },
                project: { ...targetUsage(entries, 'project'), limit: config.projectLimitBytes },
            },
            ledgerTotal: records.length,
            ledger: records.slice(-Math.max(0, recentLedger)).reverse(),
            revision: sha1(entries.map((entry) => `${entry.id}:${entry.updatedAt}`).sort().join('|')),
            origins: stores.map((store) => store.id),
            autoCapture: config.autoCapture,
            paths: {
                dir: relPath(base, paths.dir),
                user: relPath(base, paths.user),
                project: relPath(base, paths.project),
                ledger: relPath(base, paths.ledger),
                archive: relPath(base, paths.archive),
                links: relPath(base, paths.links),
            },
            globalDir: stores.some((store) => store.id === 'global') ? relPath(base, globalMemoryPaths(config).dir) : '',
        };
    }

    /**
     * 往归档里写一块外部来源的条目（mnemon 迁移与 Pack 导入用它）。
     *
     * 与容量下沉、台账滚动共用同一条写入路径：id 由调用方给（迁移用 mnemon 的
     * 原始 id），所以「同一个来源重复迁移」可以按 id 去重，不必靠内容比对。
     */
    async function appendArchiveBlock({ reason, items, month, store: wantedStore } = {}) {
        const clean = asArray(items).filter((item) => item !== null && typeof item === 'object');
        if (clean.length === 0) return { added: 0, month: null, file: null };
        const store = wantedStore === 'global' ? (stores.find((item) => item.id === 'global') ?? defaultStore()) : defaultStore();
        return withQueue(store.paths.dir, async () => {
            const created = await appendArchive(store.paths, [{
                at: nowIso(),
                month: asText(month) || monthOf(nowIso()),
                reason: asText(reason) || 'external',
                items: clean,
            }], config);
            const digest = created[0];
            return {
                added: clean.length,
                month: digest?.month ?? null,
                origin: store.id,
                file: digest === undefined ? null : relPath(base, join(store.paths.archive, `${digest.month}.md`)),
            };
        });
    }

    /** 归档里已有的全部条目（迁移用它做幂等判断，也用于状态面板与实体索引）。 */
    async function archiveItems() {
        const out = [];
        for (const store of stores) {
            const index = await loadArchiveIndex(store.paths);
            for (const digest of index.digests) {
                for (const block of asArray(digest.blocks)) {
                    for (const item of asArray(block.items)) out.push({ ...item, origin: store.id, month: digest.month, file: digest.file, blockId: block.id });
                }
            }
        }
        return out;
    }

    // ── 图关系 ──────────────────────────────────────────────────────────────

    /**
     * 建一条双向类型化关系。
     *
     * 两个 id 都必须**真的存在**（热记忆、台账或归档里能找到）：允许指向不存在的
     * id 会让关系图慢慢烂成一堆悬空边，而悬空边在「沿关系找相邻条目」时表现为
     * 静默少一条 —— 那比直接报错难查得多。
     */
    async function link({ sourceId = '', targetId = '', kind = '', note = '' } = {}) {
        if (!config.links) throw new Error('图关系已在设置里关掉（memory.links）。');
        const from = asText(sourceId);
        const to = asText(targetId);
        if (from === '' || to === '') throw new Error('link 需要 sourceId 与 targetId。');
        if (from === to) throw new Error('不能把一条记忆连到它自己。');
        const realKind = LINK_KINDS.includes(kind) ? kind : DEFAULT_LINK_KIND;

        const index = await entryIndex();
        for (const id of [from, to]) {
            if (!index.has(id)) throw new Error(`找不到 id 为「${id}」的记忆条目。先用 action:'read' 或 action:'related' 拿到确切 id。`);
        }
        const fromStore = stores.find((store) => store.id === index.get(from).origin) ?? defaultStore();
        return withQueue(fromStore.paths.dir, async () => {
            const existing = await loadLinks(fromStore.paths);
            const duplicate = existing.find((item) => item.sourceId === from && item.targetId === to && item.kind === realKind);
            if (duplicate !== undefined) return { added: false, link: duplicate, origin: fromStore.id };
            const record = normalizeLink({ sourceId: from, targetId: to, kind: realKind, note, at: nowIso() });
            await appendLinks(fromStore.paths, [record]);
            return { added: true, link: record, origin: fromStore.id };
        });
    }

    /** 删掉一条关系（按 id）。 */
    async function unlink({ id = '' } = {}) {
        const wanted = asText(id);
        if (wanted === '') throw new Error('unlink 需要关系 id。');
        for (const store of stores) {
            const links = await loadLinks(store.paths);
            const hit = links.find((item) => item.id === wanted);
            if (hit === undefined) continue;
            return withQueue(store.paths.dir, async () => {
                const kept = (await loadLinks(store.paths)).filter((item) => item.id !== wanted);
                await writeAtomic(store.paths.links, kept.map((item) => `${JSON.stringify(item)}\n`).join(''));
                return { removed: true, link: hit, origin: store.id };
            });
        }
        throw new Error(`找不到 id 为「${wanted}」的关系。`);
    }

    /**
     * 全量条目索引：id → { id, layer, origin, text, item }。
     * 热记忆与台账优先于归档（同 id 只可能出现在热记忆下沉之后的归档里）。
     */
    async function entryIndex() {
        const index = new Map();
        if (layers.archive) {
            for (const item of await archiveItems()) {
                const id = asText(item.id);
                if (id === '') continue;
                index.set(id, { id, layer: 'archive', origin: item.origin, text: archiveItemText(item), item });
            }
        }
        if (layers.ledger) {
            for (const store of stores) {
                for (const record of await loadLedger(store.paths)) {
                    const id = asText(record.id);
                    if (id === '') continue;
                    index.set(id, {
                        id,
                        layer: 'ledger',
                        origin: store.id,
                        text: [asText(record.path), asText(record.purpose)].filter(Boolean).join(' —— '),
                        item: record,
                    });
                }
            }
        }
        if (layers.hot) {
            for (const store of stores) {
                for (const entry of (await loadStore(store.paths)).entries) {
                    index.set(entry.id, { id: entry.id, layer: 'hot', origin: store.id, text: entry.content, item: entry });
                }
            }
        }
        return index;
    }

    /** 全部关系（跨根合并）。 */
    async function allLinks() {
        const out = [];
        for (const store of stores) {
            for (const link of await loadLinks(store.paths)) out.push({ ...link, origin: store.id });
        }
        return out;
    }

    /**
     * 沿关系找相邻条目。
     *
     * 关系是双向的，所以从哪一端查都能找到对面；`depth` 默认 1 跳 ——
     * mnemon 也把遍历限在一跳，图一旦放开很容易变成「把整个记忆库拉进上下文」。
     */
    async function related({ id = '', depth = 1, kind = '', limit = 0 } = {}) {
        if (!config.links) throw new Error('图关系已在设置里关掉（memory.links）。');
        const start = asText(id);
        if (start === '') throw new Error('related 需要 id（先用 action:\'read\' 拿到条目 id）。');
        const index = await entryIndex();
        if (!index.has(start)) throw new Error(`找不到 id 为「${start}」的记忆条目。`);
        const wantedKind = LINK_KINDS.includes(kind) ? kind : '';
        const links = (await allLinks()).filter((link) => wantedKind === '' || link.kind === wantedKind);
        const hops = Math.min(3, Math.max(1, boundedInt(depth, 1, 1, 3)));
        const cap = limit > 0 ? Math.min(limit, READ_LIMITS.relatedResults) : READ_LIMITS.relatedResults;

        const seen = new Set([start]);
        const nodes = [];
        const edges = [];
        let frontier = [start];
        for (let step = 0; step < hops; step += 1) {
            const next = [];
            for (const current of frontier) {
                for (const link of links) {
                    const other = link.sourceId === current ? link.targetId : link.targetId === current ? link.sourceId : '';
                    if (other === '' || seen.has(other)) continue;
                    seen.add(other);
                    const entry = index.get(other);
                    edges.push({ id: link.id, kind: link.kind, from: current, to: other, note: link.note, origin: link.origin });
                    if (entry !== undefined && nodes.length < cap) {
                        nodes.push({ ...entry, via: link.kind, hop: step + 1 });
                    }
                    next.push(other);
                }
            }
            frontier = next;
            if (frontier.length === 0) break;
        }
        const startEntry = index.get(start);
        return {
            id: start,
            text: startEntry.text,
            layer: startEntry.layer,
            origin: startEntry.origin,
            nodes,
            edges,
            truncated: nodes.length >= cap,
        };
    }

    /**
     * 实体视图：条目**显式声明**的 entities（外加 mnemon 迁来的那份）。
     *
     * 不做抽取：猜出来的实体比没有实体更误导。热记忆可以在 add/replace 时带
     * entities 数组，归档条目里的 entities 来自迁移数据。
     */
    async function entities({ query = '', limit = 0 } = {}) {
        const terms = queryTerms(query);
        const table = new Map();
        const add = (name, ref) => {
            const key = asText(name);
            if (key === '') return;
            if (!table.has(key)) table.set(key, { name: key, count: 0, refs: [] });
            const entry = table.get(key);
            entry.count += 1;
            if (entry.refs.length < 6) entry.refs.push(ref);
        };
        if (layers.hot) {
            for (const store of stores) {
                for (const item of (await loadStore(store.paths)).entries) {
                    for (const name of stringList(item.entities)) add(name, { id: item.id, layer: 'hot', origin: store.id, text: truncate(item.content, 80) });
                }
            }
        }
        if (layers.archive) {
            for (const item of await archiveItems()) {
                for (const name of stringList(item.entities)) {
                    add(name, { id: asText(item.id), layer: 'archive', origin: item.origin, text: truncate(archiveItemText(item), 80) });
                }
            }
        }
        let items = [...table.values()];
        if (terms.length > 0) {
            items = items.filter((entry) => terms.some((term) => entry.name.toLowerCase().includes(term)));
        }
        items.sort((left, right) => (right.count - left.count) || left.name.localeCompare(right.name));
        const cap = limit > 0 ? Math.min(limit, READ_LIMITS.entityResults) : READ_LIMITS.entityResults;
        return { items: items.slice(0, cap), total: items.length, truncated: items.length > cap };
    }

    // ── 备份与迁移（Pack） ──────────────────────────────────────────────────

    /**
     * 整包导出：热记忆 + 台账 + 归档（索引与 Markdown 全文）+ 关系，一个 JSON。
     *
     * 为什么不打包成 zip：本插件零依赖，而「读回来」这条路上任何一个压缩库都会
     * 变成新的依赖；JSON 单文件同样能整包搬走，还能直接看。
     */
    async function exportPack({ includeArchiveFiles = true } = {}) {
        const storesOut = [];
        for (const store of stores) {
            const loaded = await loadStore(store.paths);
            const index = await loadArchiveIndex(store.paths);
            const archiveFiles = {};
            if (includeArchiveFiles) {
                for (const digest of index.digests) {
                    const text = await readTextOrNull(join(store.paths.archive, `${digest.month}.md`));
                    if (text !== null) archiveFiles[digest.month] = text;
                }
            }
            storesOut.push({
                id: store.id,
                dir: store.paths.dir,
                entries: loaded.entries,
                ledger: await loadLedger(store.paths),
                archiveIndex: index,
                archiveFiles,
                links: await loadLinks(store.paths),
            });
        }
        return {
            format: PACK_FORMAT,
            version: PACK_VERSION,
            exportedAt: nowIso(),
            workspace: base,
            layers: { ...layers },
            stores: storesOut,
        };
    }

    /**
     * 整包导入：**只增不改**，按 id 幂等。
     *
     * 冲突口径（与 mnemon 的「导入不擦除」一致）：
     *   热记忆  同 id 已存在就跳过，不覆盖现内容
     *   台账    同 id 已存在就跳过
     *   关系    同 (source,target,kind) 已存在就跳过
     *   归档    同月份合并，块按 id（老数据按「时刻+原因+条数」）去重
     */
    async function importPack(pack, { target: wantedTarget = '' } = {}) {
        if (pack === null || typeof pack !== 'object') throw new Error('导入需要一份 Pack 对象（先 export 或读 pack.json）。');
        if (asText(pack.format) !== PACK_FORMAT) throw new Error(`这不是办公记忆 Pack（format=${asText(pack.format) || '空'}）。`);
        const incoming = asArray(pack.stores);
        if (incoming.length === 0) throw new Error('这份 Pack 里没有任何记忆根。');
        const report = { entries: 0, skippedEntries: 0, ledger: 0, skippedLedger: 0, links: 0, skippedLinks: 0, archiveBlocks: 0, skippedArchiveBlocks: 0, stores: [] };

        for (const incomingStore of incoming) {
            // 导入目标：默认把 Pack 里的根**还原到同名根**（global → 全局，其余 → 工作区）。
            const id = asText(incomingStore.id) === 'global' ? 'global' : 'workspace';
            if (wantedTarget !== '' && wantedTarget !== id) continue;
            const store = stores.find((item) => item.id === id);
            if (store === undefined) {
                report.stores.push({ id, skipped: true, reason: '当前记忆范围里没有这个根' });
                continue;
            }
            const storeReport = { id, entries: 0, ledger: 0, links: 0, archiveBlocks: 0 };

            await withQueue(store.paths.dir, async () => {
                // 热记忆
                const loaded = await loadStore(store.paths);
                const entries = loaded.entries.slice();
                const known = new Set(entries.map((entry) => entry.id));
                for (const raw of asArray(incomingStore.entries)) {
                    const entry = normalizeEntry(raw);
                    if (entry === null) continue;
                    if (known.has(entry.id)) {
                        report.skippedEntries += 1;
                        continue;
                    }
                    known.add(entry.id);
                    entries.push(entry);
                    report.entries += 1;
                    storeReport.entries += 1;
                }
                if (storeReport.entries > 0) {
                    await enforceCapacity(store.paths, entries, config);
                    await saveStore(store.paths, entries);
                    await writeProjections(store.paths, entries, config);
                }

                // 台账
                const ledger = await loadLedger(store.paths);
                const knownLedger = new Set(ledger.map((record) => asText(record.id)).filter((item) => item !== ''));
                const fresh = [];
                for (const record of asArray(incomingStore.ledger)) {
                    const rid = asText(record?.id);
                    if (rid === '' || knownLedger.has(rid)) {
                        report.skippedLedger += 1;
                        continue;
                    }
                    knownLedger.add(rid);
                    fresh.push(record);
                    report.ledger += 1;
                    storeReport.ledger += 1;
                }
                if (fresh.length > 0) await appendLedger(store.paths, fresh);

                // 关系
                const links = await loadLinks(store.paths);
                const knownLinks = new Set(links.map((link) => `${link.sourceId}|${link.targetId}|${link.kind}`));
                const freshLinks = [];
                for (const raw of asArray(incomingStore.links)) {
                    const link = normalizeLink(raw);
                    if (link === null) continue;
                    const key = `${link.sourceId}|${link.targetId}|${link.kind}`;
                    if (knownLinks.has(key)) {
                        report.skippedLinks += 1;
                        continue;
                    }
                    knownLinks.add(key);
                    freshLinks.push(link);
                    report.links += 1;
                    storeReport.links += 1;
                }
                if (freshLinks.length > 0) await appendLinks(store.paths, freshLinks);

                // 归档：按月合并，块去重
                const index = await loadArchiveIndex(store.paths);
                const blocks = [];
                for (const digest of asArray(incomingStore.archiveIndex?.digests)) {
                    const month = asText(digest?.month) || monthOf(nowIso());
                    for (const block of asArray(digest?.blocks)) {
                        const targetDigest = index.digests.find((item) => item.month === month);
                        const known = new Set((targetDigest?.blocks ?? []).map(blockKey).concat(blocks.filter((item) => item.month === month).map((item) => blockKey(item))));
                        if (known.has(blockKey(block))) {
                            report.skippedArchiveBlocks += 1;
                            continue;
                        }
                        blocks.push({ month, at: asText(block.at) || nowIso(), reason: asText(block.reason) || 'pack-import', items: asArray(block.items) });
                        report.archiveBlocks += 1;
                        storeReport.archiveBlocks += 1;
                    }
                }
                if (blocks.length > 0) await appendArchive(store.paths, blocks, config);
            });
            report.stores.push(storeReport);
        }
        return report;
    }

    /**
     * 浏览用的一次性读取：**原始记录**，不过召回质量、不按命中排序。
     *
     * 与 read() 的区别是用途不同：read() 服务「模型带着问题查资料」，所以要有
     * 相关性排序、质量分档与字符预算；browse() 服务「人在面板里看现状」，
     * 要的是最近的全量（有界），过滤交给调用方按自己的字段做。
     */
    async function browse({ ledgerLimit = 200, archiveLimit = 200, linkLimit = 500 } = {}) {
        const out = { hot: [], ledger: [], archive: [], links: [] };
        /** 归档**摘要文件**数（一个月一个）：与条目数是两个量，容量上限管的是这个。 */
        let digestCount = 0;
        for (const store of stores) {
            if (layers.hot) {
                for (const entry of (await loadStore(store.paths)).entries) out.hot.push({ ...entry, origin: store.id });
            }
            if (layers.ledger) {
                for (const record of await loadLedger(store.paths)) out.ledger.push({ ...record, origin: store.id });
            }
            if (layers.archive) {
                const index = await loadArchiveIndex(store.paths);
                digestCount += index.digests.length;
                for (const digest of index.digests) {
                    for (const block of asArray(digest.blocks)) {
                        for (const item of asArray(block.items)) {
                            out.archive.push({ ...item, origin: store.id, month: digest.month, file: digest.file, blockId: block.id });
                        }
                    }
                }
            }
            if (config.links) {
                for (const link of await loadLinks(store.paths)) out.links.push({ ...link, origin: store.id });
            }
        }
        const byAt = (left, right) => String(right.at ?? right.updatedAt ?? '').localeCompare(String(left.at ?? left.updatedAt ?? ''));
        out.ledger.sort(byAt);
        out.archive.sort(byAt);
        out.links.sort(byAt);
        return {
            hot: out.hot,
            ledger: out.ledger.slice(0, Math.max(0, ledgerLimit)),
            archive: out.archive.slice(0, Math.max(0, archiveLimit)),
            links: out.links.slice(0, Math.max(0, linkLimit)),
            // 容量口径要用**未截断**的总数：面板的容量条讲的是「离上限还有多远」，
            // 跟着浏览用的截断走会显示成「才 200 条就满了」。
            totals: { hot: out.hot.length, ledger: out.ledger.length, archive: out.archive.length, links: out.links.length },
            // 归档上限 archiveKeep 管的是**摘要文件数**（一个月一个），不是条目数 ——
            // 两者必须分开报，否则「30 条 / 60 个」这种配对会被读成「快满了」。
            digests: digestCount,
        };
    }

    /** 设置页 / 状态面板用的汇总（不读全文，只给计数与体积）。 */
    async function status() {
        const perStore = [];
        for (const store of stores) {
            const loaded = await loadStore(store.paths);
            const index = await loadArchiveIndex(store.paths);
            const archived = index.digests.reduce((sum, digest) => sum + asArray(digest.blocks).reduce((inner, block) => inner + asArray(block.items).length, 0), 0);
            const size = await dirStats(store.paths.dir);
            perStore.push({
                id: store.id,
                dir: relPath(base, store.paths.dir),
                hot: loaded.entries.length,
                ledger: (await loadLedger(store.paths)).length,
                archive: archived,
                // 归档的**摘要文件**数：面板的容量条要的是它（上限 archiveKeep 管的就是
                // 文件数），而 archive 是条目数 —— 两个数都给，别让界面自己猜。
                archiveFiles: index.digests.length,
                links: (await loadLinks(store.paths)).length,
                // 浏览面板要显示「这个根占多少地方」：字节与文件数按目录实算。
                // 目录不存在时是 0 而不是 undefined —— 契约是数字，面板直接渲染它。
                bytes: size.bytes,
                files: size.files,
            });
        }
        return { scope: config.scope, userScope: config.userScope, layers: { ...layers }, stores: perStore };
    }

    return {
        root: base,
        paths,
        config,
        stores: stores.map((store) => ({ id: store.id, dir: store.paths.dir })),
        snapshot,
        read,
        browse,
        mutate,
        log,
        digest,
        appendArchiveBlock,
        archiveItems,
        link,
        unlink,
        related,
        entities,
        entryIndex,
        exportPack,
        importPack,
        status,
    };
}

/**
 * 容量维持：谁超出上限，就把「重要度低、更旧」的条目移进归档。
 *
 * 四个细节：
 *   - 候选包含刚写入的那条（不豁免）。否则「刚写了一条 low、而旧条目都是
 *     critical」时，唯一能下沉的就只剩 critical —— 保护新条目反而赶走了更值钱
 *     的记忆。重要度是唯一的排序主键：同重要度里最新的排最后。
 *   - 下沉后补一条指路条目（固定 id，只留一条），并且**先为它预留空间**，
 *     否则贴着上限的小容量会「下沉完没地方放指路」，投影里内容凭空少了；
 *   - 全都下沉还是装不下（单条就超限）就抛错，而不是静默截断内容。
 */
async function enforceCapacity(paths, entries, config) {
    const maintenance = [];
    for (const target of MEMORY_TARGETS) {
        const limit = target === 'user' ? config.userLimitBytes : config.projectLimitBytes;
        let usage = targetUsage(entries, target);
        if (usage.used <= limit) continue;

        const before = usage.used;
        const moved = [];
        // 指路条目自己也要占空间，所以先把它的位置预留出来再下沉。
        const pointerText = `较旧的${target === 'user' ? '用户偏好' : '项目记忆'}已下沉到归档，检索：office_memory({ action: 'read', layer: 'archive', query: '...' })`;
        const pointerBytes = byteLength(pointerText) + byteLength(ENTRY_DELIMITER);
        const pointerFits = pointerBytes > 0 && pointerBytes < limit;
        // 已经有指路条目时不必再为它预留空间（它已经在用量里了）。按
        // limit - pointerBytes 一路下沉会把「预留」变成「多留一份」，最后连
        // critical 也被吃掉 —— 实测过。
        const hasPointer = entries.some((entry) => entry.id === ARCHIVE_POINTER_ID && entry.target === target);
        const demote = () => {
            // 「更旧」用**插入次序**判断，不用时间戳：同一毫秒内连着写多条时
            // updatedAt 完全相同，按它会退化成按 id 随机挑一个下沉。
            const candidates = entries
                .map((entry, index) => ({ entry, index }))
                .filter(({ entry }) => entry.target === target && entry.id !== ARCHIVE_POINTER_ID)
                .sort((left, right) => (IMPORTANCE_RANK[right.entry.importance] - IMPORTANCE_RANK[left.entry.importance])
                    || (left.index - right.index));
            const victim = candidates[0]?.entry;
            if (victim === undefined) return false;
            entries.splice(entries.indexOf(victim), 1);
            moved.push(victim);
            usage = targetUsage(entries, target);
            return true;
        };
        const reservation = pointerFits && !hasPointer ? limit - pointerBytes : limit;
        while (usage.used > reservation && demote()) {
            // 继续下沉，直到留出指路条目的位置
        }
        if (moved.length === 0) {
            throw new Error(`${target} 侧的热记忆 ${before} 字节超出上限 ${limit}，但没有可下沉的条目（单条就超限）。请压缩内容。`);
        }
        await appendArchive(paths, [{
            at: nowIso(),
            month: monthOf(new Date().toISOString()),
            reason: 'hot-overflow',
            items: moved.map(hotArchiveItem),
        }], config);

        if (pointerFits) {
            const pointer = normalizeEntry({
                id: ARCHIVE_POINTER_ID,
                target,
                importance: 'low',
                content: pointerText,
                createdAt: nowIso(),
                updatedAt: nowIso(),
            });
            const existing = entries.findIndex((entry) => entry.id === ARCHIVE_POINTER_ID && entry.target === target);
            if (existing >= 0) entries.splice(existing, 1);
            entries.push(pointer);
            usage = targetUsage(entries, target);
            if (usage.used > limit) {
                // 兜底：真到了「连一条指路都放不下」的容量，就只保内容不留指路。
                entries.splice(entries.indexOf(pointer), 1);
                usage = targetUsage(entries, target);
            }
        }
        maintenance.push({
            kind: 'mnemon-archive',
            target,
            moved: moved.length,
            movedIds: moved.map((entry) => entry.id),
            usageBefore: before,
            usageAfter: usage.used,
            limit,
        });
    }
    return maintenance;
}

// ── 渲染（给模型看的文本） ──────────────────────────────────────────────────

function usageLine(target, usage) {
    return `${target === 'user' ? '用户偏好' : '项目与环境'} ${usage.entryCount} 条 / ${usage.used} 字节（上限 ${usage.limit}）`;
}

/** 条目来自哪个根：多根时这一小段是唯一能分辨「全局偏好」与「本项目约定」的地方。 */
function originTag(origin) {
    return origin === 'global' ? '（全局）' : '';
}

function hotLines(hot) {
    const lines = [];
    for (const target of MEMORY_TARGETS) {
        const items = hot.items.filter((item) => item.target === target);
        lines.push(`【${target === 'user' ? '用户偏好' : '项目与环境'}】${items.length === 0 ? '（空）' : ''}`);
        for (const item of items) lines.push(`- [${item.importance}]${originTag(item.origin)} ${item.text.replace(/\s*\n\s*/g, ' ')}`);
        lines.push(`  ${usageLine(target, hot.targets[target])}`);
    }
    return lines;
}

/** 渲染 action:'read' 的结果。 */
export function renderRead(result) {
    const lines = [];
    const title = result.layer === 'all' ? '三层' : result.layer === 'hot' ? '热记忆' : result.layer === 'ledger' ? '台账' : '归档';
    lines.push(`📒 记忆读取（${title}）${result.query ? `：${result.query}` : ''}`);
    if (result.origins !== undefined && result.origins.length > 1) {
        lines.push(`范围：${result.origins.map((id) => (id === 'global' ? '全局' : '工作区')).join(' + ')}${result.globalDir ? `（全局层 ${result.globalDir}）` : ''}`);
    }
    if (result.hot !== undefined) {
        lines.push('', ...hotLines(result.hot));
    }
    if (result.ledger !== undefined) {
        lines.push('', `【台账】共 ${result.ledger.total} 条${result.query ? `，命中 ${result.ledger.matched} 条` : ''}`);
        if (result.ledger.items.length === 0) lines.push('- （空）');
        for (const item of result.ledger.items) {
            const badge = [item.format, item.theme].filter(Boolean).join(' · ');
            lines.push(`- ${item.at.slice(0, 10)}${originTag(item.origin)} ${item.text}${badge ? `  [${badge}]` : ''}`);
        }
    }
    if (result.archive !== undefined) {
        lines.push('', `【归档】${result.archive.digests} 个摘要 / ${result.archive.items_total} 条${result.query ? `，命中 ${result.archive.matched} 条` : ''}`);
        if (result.archive.items.length === 0) lines.push('- （空）');
        for (const item of result.archive.items) {
            const label = item.kind === 'hot' ? `热记忆（${item.origin}/${item.importance}）`
                : item.kind === 'mnemon' ? `mnemon 长期记忆${item.category ? `（${item.category}）` : ''}`
                    : '台账';
            const id = asText(item.id) === '' ? '' : ` ${item.id}`;
            lines.push(`- [${item.month}]${id} ${label} ${item.text}`);
        }
    }
    if (result.qualityDroppedTotal !== undefined && result.qualityDroppedTotal > 0) {
        lines.push('', `（召回质量 strict-v1 丢掉了 ${result.qualityDroppedTotal} 条低相关结果：换更准的关键词再试。）`);
    }
    if (result.truncated === true) lines.push('', '（结果已按条数与字符上限截断：要更精确就带 query，或缩小 layer。）');
    lines.push('', `存哪儿：${result.paths.memory}（USER.md / MEMORY.md / ledger.jsonl / links.jsonl / archive 都是它的投影与分片）`);
    return lines.join('\n');
}

/** 渲染一次写入的回执。mnemon 的做法：给回执，不回显整份文件。 */
export function renderMutation(result) {
    const lines = [`📒 记忆已更新（${result.action} / ${result.target}）${originTag(result.origin)}`];
    if (result.action === 'remove') {
        lines.push(`- 移除：[${result.previous?.importance ?? 'normal'}] ${truncate(result.previous?.content ?? '', 120)}`);
    } else if (result.action === 'replace') {
        lines.push(`- 改前：[${result.previous?.importance ?? 'normal'}] ${truncate(result.previous?.content ?? '', 120)}`);
        lines.push(`- 改后：[${result.entry?.importance ?? 'normal'}] ${result.entry?.content ?? ''}`);
    } else {
        lines.push(`- 新增：[${result.entry?.importance ?? 'normal'}] ${result.entry?.content ?? ''}`);
    }
    lines.push(`- 用量：${usageLine('user', result.usage.user)}；${usageLine('project', result.usage.project)}`);
    if (result.evicted === true) {
        lines.push('- 注意：容量维持把刚写的这条也下沉了（它比现有条目更次要）。要留住它就先精简现有条目，或调大上限。');
    }
    for (const item of result.maintenance ?? []) {
        lines.push(`- 容量维持：${item.target} 侧下沉了 ${item.moved} 条较旧条目（${item.usageBefore} → ${item.usageAfter} 字节），已进归档。`);
    }
    lines.push('- 投影已重写：USER.md / MEMORY.md（都是生成物，不要直接编辑）。');
    return lines.join('\n');
}

/** 渲染台账登记的回执。 */
export function renderLog(result) {
    if (result.disabled === true) return '📒 台账层已在设置里关掉（memory.layers.ledger），本次没有登记。';
    const lines = [`📒 台账已登记 ${result.added} 条（现共 ${result.kept} 条）`];
    if (result.rolled > 0) lines.push(`- ${result.rolled} 条最旧记录已滚进归档（台账条数有上限）。`);
    return lines.join('\n');
}

/** 渲染建关系的回执。 */
export function renderLink(result) {
    const link = result.link ?? {};
    const head = result.added === true ? '📒 已建立关系' : '📒 这条关系已经存在（未重复写入）';
    return [
        head,
        `- ${link.id ?? ''} ${link.sourceId ?? ''} —[${link.kind ?? ''}]→ ${link.targetId ?? ''}${link.note ? `（${link.note}）` : ''}`,
        '- 关系是双向的：从任一端用 office_memory({ action: \'related\', id }) 都能走到对面。',
    ].join('\n');
}

/** 渲染删关系的回执。 */
export function renderUnlink(result) {
    return `📒 已删除关系 ${result.link?.id ?? ''}（${result.link?.sourceId ?? ''} — ${result.link?.targetId ?? ''}）。`;
}

/** 渲染沿关系找相邻条目的结果。 */
export function renderRelated(result) {
    const lines = [`📒 关系遍历：${result.id}（${result.layer}）${result.text ? ` —— ${truncate(result.text, 80)}` : ''}`];
    if (result.nodes.length === 0) {
        lines.push('- 这条记忆还没有相邻条目。可以用 action:\'link\' 建一条关系。');
        return lines.join('\n');
    }
    lines.push(`- 相邻 ${result.nodes.length} 条（共 ${result.edges.length} 条边）：`);
    for (const node of result.nodes) {
        lines.push(`  - [${node.layer}/${node.via}${node.hop > 1 ? ` ${node.hop}跳` : ''}] ${node.id}${originTag(node.origin)} ${truncate(node.text, 120)}`);
    }
    if (result.truncated === true) lines.push('- （已达条数上限：要看得更窄就带 kind。）');
    return lines.join('\n');
}

/** 渲染实体视图。 */
export function renderEntities(result) {
    const lines = [`📒 实体视图（${result.total} 个）`];
    if (result.items.length === 0) {
        lines.push('- （还没有声明过任何实体。）');
        lines.push('- 实体来自条目显式带的 entities 数组：add/replace 时传 entities: ["某某系统", "某某人"] 才会出现在这里。');
        return lines.join('\n');
    }
    for (const item of result.items) {
        lines.push(`- ${item.name}（${item.count} 条）`);
        for (const ref of item.refs.slice(0, 3)) {
            lines.push(`  - ${ref.id} [${ref.layer}] ${truncate(ref.text, 100)}`);
        }
    }
    if (result.truncated === true) lines.push('- （已达条数上限。）');
    return lines.join('\n');
}

/** 渲染 Pack 导出回执。 */
export function renderExport(result) {
    const stores = (result.pack?.stores ?? []).map((store) => `${store.id}（热记忆 ${store.entries.length} / 台账 ${store.ledger.length} / 关系 ${store.links.length} / 归档摘要 ${store.archiveIndex?.digests?.length ?? 0}）`);
    return [
        `📒 记忆已导出：${result.path}`,
        `- ${stores.join('；')}`,
        `- 导出时刻：${result.pack?.exportedAt ?? ''}；体积 ${result.bytes} 字节。`,
        '- 这份 Pack 含私有记忆，不要当公开文件传。',
    ].join('\n');
}

/** 渲染 Pack 导入回执。 */
export function renderImport(result) {
    const lines = [`📒 记忆导入完成（只增不改，按 id 去重）`];
    lines.push(`- 热记忆 新增 ${result.entries} 条 / 跳过 ${result.skippedEntries} 条`);
    lines.push(`- 台账 新增 ${result.ledger} 条 / 跳过 ${result.skippedLedger} 条`);
    lines.push(`- 关系 新增 ${result.links} 条 / 跳过 ${result.skippedLinks} 条`);
    lines.push(`- 归档 新增 ${result.archiveBlocks} 块 / 跳过 ${result.skippedArchiveBlocks} 块`);
    for (const store of result.stores ?? []) {
        if (store.skipped === true) lines.push(`- ${store.id}：跳过（${store.reason}）`);
    }
    return lines.join('\n');
}

/** 渲染状态汇总（设置页与 office_memory 的 status 用）。 */
export function renderStatus(result) {
    const lines = [`📒 记忆状态（范围 ${result.scope}，USER 范围 ${result.userScope}）`];
    lines.push(`- 层：${Object.entries(result.layers).map(([name, on]) => `${name}${on ? '开' : '关'}`).join(' / ')}`);
    for (const store of result.stores) {
        lines.push(`- ${store.id === 'global' ? '全局' : '工作区'} ${store.dir}：热记忆 ${store.hot} / 台账 ${store.ledger} / 归档 ${store.archive} / 关系 ${store.links}`);
    }
    return lines.join('\n');
}

/**
 * 渲染紧凑摘要：挂在 office_help 与 office_run 的反馈里，等于「动笔前的记忆投影」。
 *
 * `hot: 'unchanged'` 是「只在变化时贴」用的形态：热记忆的**正文**换成一到两行
 * 说明（条数 + revision + 怎么读全文），台账部分照旧 —— 台账不体现在 revision
 * 里，省掉它会让 office_run 之后看不到刚登记的那条。决定贴哪种形态的是
 * projection.js，这里只负责渲染。
 */
export function renderDigest(digest, { context = 'help', hot = 'full' } = {}) {
    if (digest.empty === true) {
        const lines = [
            `📒 记忆（还没有内容；目录 ${digest.paths.dir}${digest.globalDir ? `，全局层 ${digest.globalDir}` : ''}）`,
            "- 记用户对文档的偏好与要求（target:'user'）、项目约定与踩过的坑（target:'project'）：office_memory({ action: 'add', target, content, importance? })。",
            '- 不记：检索结果原文、一次性进度、猜测，以及助手自己说过的话。',
        ];
        // 空态也要带主动记录指引：第一次用到记忆时正是最该说这句话的时候。
        if (digest.autoCapture === true) {
            lines.push('- 主动记录：用户在这一轮里说出的偏好、纠正、稳定事实，当场用 action:\'add\' 记一条；不必等他说「记住」。');
        }
        return lines.join('\n');
    }
    const omitted = hot === 'unchanged';
    const revision = String(digest.revision ?? '').slice(0, 8);
    const lines = [
        omitted
            ? '📒 记忆（热记忆与上次投影相同，省略正文；用量与最近台账照旧）'
            : `📒 记忆（${context === 'run' ? '本次调用后' : '动笔前'}先看这几条）`,
    ];
    for (const target of MEMORY_TARGETS) {
        const items = target === 'user' ? digest.user : digest.project;
        if (items.length === 0) continue;
        lines.push(`【${target === 'user' ? '用户偏好' : '项目与环境'}】`);
        if (omitted) {
            lines.push(`- 与上一次投影相同（${items.length} 条，revision ${revision}），正文不再重复；要看：office_memory({ action: 'read', layer: 'hot' })`);
        } else {
            for (const entry of items) lines.push(`- [${entry.importance}]${originTag(entry.origin)} ${entry.content.replace(/\s*\n\s*/g, ' ')}`);
        }
        lines.push(`  ${usageLine(target, digest.usage[target])}`);
    }
    if (digest.ledger.length > 0) {
        lines.push(`【台账】共 ${digest.ledgerTotal} 条，最近 ${digest.ledger.length} 条：`);
        for (const record of digest.ledger) {
            const badge = [record.format, record.theme].filter(Boolean).join(' · ');
            lines.push(`- ${String(record.at).slice(0, 10)}${originTag(record.origin)} ${record.path}${badge ? `  [${badge}]` : ''}${record.purpose ? ` —— ${record.purpose}` : ''}`);
        }
    }
    lines.push(`- 完整读取与检索：office_memory({ action: 'read', layer: 'all'|'hot'|'ledger'|'archive', query? })`);
    if (digest.autoCapture === true) {
        lines.push('- 主动记录：用户在这一轮里说出的偏好、纠正、稳定事实，当场用 action:\'add\' 记一条；不必等他说「记住」。');
    }
    return lines.join('\n');
}

/** 目录是否已经存在（用于「第一次使用」的判断，避免为了看一眼又建一次目录）。 */
export function memoryExists(root, memoryConfig = {}) {
    return existsSync(memoryPaths(root, memoryConfig).dir);
}

/**
 * 一个记忆目录的体积与文件数（递归，含 archive/ 里的摘要文件）。
 *
 * 目录不存在时给 `{bytes: 0, files: 0}`，**不是 undefined**：这份数字会直接进
 * 浏览面板与 status 回执，契约是数字 —— undefined 在界面上会渲染成 NaN 或空白。
 */
async function dirStats(dir) {
    if (!existsSync(dir)) return { bytes: 0, files: 0 };
    const files = [];
    async function walk(current) {
        for (const entry of await readdir(current, { withFileTypes: true })) {
            const target = join(current, entry.name);
            if (entry.isDirectory()) await walk(target);
            else files.push(target);
        }
    }
    await walk(dir);
    let bytes = 0;
    for (const file of files) {
        try {
            bytes += (await stat(file)).size;
        } catch {
            // 文件在读的中间被删掉：忽略，统计本来就是个约数。
        }
    }
    return { bytes, files: files.length };
}

/** 记忆目录的体积与条数（设置页与 office.memory.stats 用）。 */
export async function memoryStats(root, memoryConfig = {}) {
    const paths = memoryPaths(root, memoryConfig);
    const exists = existsSync(paths.dir);
    const stats = await dirStats(paths.dir);
    return { dir: relPath(root, paths.dir), exists, bytes: stats.bytes, files: stats.files };
}
