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
import { existsSync, statSync } from 'node:fs';
import { mkdir, readFile, rename, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { basename, dirname, isAbsolute, join, posix, relative, sep } from 'node:path';

import {
    KB_HIT_PREVIEW,
    KB_LIST_LIMIT,
    KB_MAX_DOC_BYTES,
    KB_READ_CHARS,
    KB_READ_CHUNKS,
    KB_SEARCH_CHUNKS,
    KB_SEARCH_LIMIT,
    KB_TIERS,
    KB_VERSION,
    chunkCharsFor,
    chunkHeaderOf,
    chunkIdOf,
    chunkRelOf,
    chunkSelfConsistent,
    docIdOf,
    kbCountsOf,
    kbHash,
    kbLiveIdsOf,
    kbPathsFor,
    kbRankChunks,
    kbTierOf,
    kbTerms,
    normalizeKbChunk,
    normalizeManifestRow,
    scoreText,
    splitDocument,
    spanText,
    strictTierOf,
    strictestTier,
} from './kb.js';

/** 存储格式版本；将来改结构时用它做迁移判断。 */
export const MEMORY_VERSION = 1;
/** 记忆目录默认位置（相对会话工作目录）。 */
export const DEFAULT_MEMORY_DIR = '.office/memory';
/** 热记忆的两个上限（字节，与 mnemon 的 USER.md / MEMORY.md 一致）。 */
export const DEFAULT_USER_LIMIT_BYTES = 4096;
export const DEFAULT_PROJECT_LIMIT_BYTES = 10240;
/**
 * 一次投影里条目正文的字节上限（第四十八轮 P0-2）。
 *
 * 取 4096：与用户层容量上限同量级，比本机实测的「全文 15,444 B」小一个数量级，
 * 又足够装下三五条 300–800 B 的常用约定。超出的条目不丢，折叠行报条数并给读取入口。
 * 这是**投影**的额度，不是记忆库的容量 —— 后者看 userLimitBytes / projectLimitBytes。
 */
export const DEFAULT_PROJECTION_BUDGET_BYTES = 4096;

/**
 * 单条记忆的**硬上限**（第四十八轮 新-11）：超过就拒绝写入，并给压缩模板。
 *
 * 取 2800：本机真实库单条正文中位 599 B、≥1 KB 的 5 条、最长 2,374 B —— 这个默认
 * 不误伤存量，同时拦住「一条 10 KB 的段落塞进项目层」这种把热记忆吃光、又没人读的写法。
 * 热记忆是**给模型每次动笔前照办**的清单，不是文档仓库；长内容的去处是 source / kb / 交付物。
 */
export const DEFAULT_ENTRY_LIMIT_BYTES = 2800;

/**
 * 单条的**建议线**（第四十八轮 新-11）：超过只在回执里提醒，不拒绝。
 *
 * 取 1200：介于中位（599 B）与硬上限（2800 B）之间，用来给「写成长段落」这个习惯
 * 一个可见的反馈，而又不把确实需要一段话的条目卡死。0 = 关掉提醒。
 */
export const DEFAULT_ENTRY_HINT_BYTES = 1200;
/** 台账最多保留多少条，超出的最旧条目滚进归档。 */
export const DEFAULT_LEDGER_LIMIT = 500;
/** 归档最多保留多少**卷**摘要文件，超出的最旧卷被删（归档本身也是有界的）。 */
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
 * 管图关系与实体；status 给汇总；export/import 是备份与迁移；migrate 从 mnemon 搬；
 * kb-ingest / kb-search / kb-list / kb-read / kb-drop 是知识库（第四棵树，检索是词法的）。
 */
export const MEMORY_ACTIONS = [
    'read', 'add', 'replace', 'remove', 'log', 'status', 'sunk', 'link', 'unlink', 'related', 'entities',
    'kb-ingest', 'kb-search', 'kb-list', 'kb-read', 'kb-drop',
    'export', 'import', 'migrate',
];
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
    // 下面三个名额上限在第二十四轮放宽（3/4/2 → 5/6/4）。理由是实证的：**上限由召回决定、
    // 不由排序决定** —— RMM 的 Oracle 检索上限 Acc 90.2 而实际 70.4，加 reranker 只给
    // +1.4 Recall@5，把 Top-M 从 5 放到 10 却给 +4.6；RAGChecker 里换检索器只值 +2.4 F1，
    // 而把名额 k 从 5 放宽到 20 值 +1.7 F1。所以「低分档直接丢弃」比「排序不精」更伤。
    // 放宽只影响候选与采纳的条数，真正的字符开销仍由 READ_LIMITS 的额度封住。
    candidateMultiplier: 5,
    maxMediumResults: 6,
    maxUnknownResults: 4,
};

/**
 * 每回合配额默认值，对齐 mnemon 的回合预算：
 *   recallPerTurn       一个回合里第一次带 query 的记忆检索
 *   recallRefinePerTurn 同一回合里后续的换词细化（mnemon 是「一次初查 + 一次细化」）
 *   kbSearchPerTurn     一个回合里的 kb（来源原文）检索次数
 *   relatedPerTurn      一个回合里的图关系遍历
 * 0 = 该类不限制。
 *
 * kb 检索单独给一类（第二十四轮 24-11）：它与记忆检索的语料与代价都不同 ——
 * 记忆层是几百条短条目，kb 是几百块长原文。给同一类配额会让「查一次记忆」把
 * 「查一次来源」的名额一起吃掉，于是模型只能二选一。
 */
export const DEFAULT_QUOTA = {
    recallPerTurn: 1,
    recallRefinePerTurn: 1,
    kbSearchPerTurn: 2,
    relatedPerTurn: 1,
};

/** 关系的类型。双向存储，所以不必为反向再写一条。 */
export const LINK_KINDS = ['related', 'refines', 'supersedes', 'contradicts', 'supports', 'derives'];
/** 关系的默认类型。 */
export const DEFAULT_LINK_KIND = 'related';

/**
 * 冲突边（`kind:'contradicts'`）的两个附加字段（第二十四轮 §5.3 / P1-8）。
 *
 * 类别取 2403.08319 的三分：上下文内 / 上下文之间 / 记忆内部；状态取「未决 / 偏向
 * 甲 / 偏向乙」——**保留双方、显式承认未决**，不做自动裁决（自动裁决各有反面实证）。
 */
export const CONFLICT_CLASSES = ['context-memory', 'inter-context', 'intra-memory'];
export const CONFLICT_STATES = ['unresolved', 'prefer-source', 'prefer-target'];

/** 类别收敛：不认识的值算「没分类」（它只是给人看的标签，不该拒绝一条真实的冲突）。 */
function linkConflictClass(value) {
    const wanted = asText(value);
    return CONFLICT_CLASSES.includes(wanted) ? wanted : '';
}

/** 状态收敛：不认识的值算 `unresolved` —— 冲突默认是**未决**，这是安全的那一侧。 */
function linkConflictState(value) {
    const wanted = asText(value);
    return CONFLICT_STATES.includes(wanted) ? wanted : 'unresolved';
}

/** 冲突状态在回执里的中文写法。 */
export function conflictStateText(state) {
    if (state === 'prefer-source') return '偏向起点';
    if (state === 'prefer-target') return '偏向终点';
    return '未决';
}

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
    /**
     * `layer:'all'` 时每一层**各自的**字符额度。
     *
     * 为什么不再是三层共用一份 `totalChars`：热记忆的正文有 11.8 KB，而它是被**先**采纳的
     * 一层，一层就能把 4,000 字吃干净 —— 台账与归档随后拿到的剩余额度是 0，`admit()` 直接
     * `break`，于是**恒定返回 0 条**（哪怕命中 20/54 条），而渲染还把它印成「（空）」。
     * 台账与归档正是带 query 的检索层，「找上次那份东西」因此在默认调用上完全失效。
     * （第二十四轮实测：`.office/tmp/memory-r24/probe-budget.mjs`。）
     *
     * 改成每层一份有界额度后：① 任何一层都不会被别人饿死；② 总量仍有上界（三层额度之和），
     * 不会变成无界注入；③ 单层读取（`layer:'hot'` 等）仍独占 `totalChars`。
     * 每层额度可以取满自己的 `totalChars`（热记忆与单层读取一致），其余两层按实际条目数花。
     */
    allChars: { hot: 4000, ledger: 2000, archive: 1600 },
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
        tombstones: join(dir, 'tombstones.jsonl'),
        // 知识库是**兄弟目录**（`.office/kb`）：它装的是外部原文，不随三层生命周期走。
        kb: kbPathsFor(dir),
    };
}

/**
 * 解析记忆目录。目录可以给相对路径（相对会话工作目录）或绝对路径。
 * @returns {{dir: string, memory: string, ledger: string, archive: string, user: string, project: string, index: string, links: string, tombstones: string}}
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

/**
 * 严格的读：只有「文件不存在」算空，其它错误照抛。
 *
 * 用在**读-改-写**的那两份 append-only 文件上（`links.jsonl` / `tombstones.jsonl`）：
 * `readTextOrNull` 把 EACCES / EISDIR / EBUSY 统统变成 null，于是 `pruneLinks` 会算出
 * 「一条边都没删」、`appendLinks` 会拿空前缀把整份文件重写成只有新边 —— 一次读失败就等于
 * 整个关系图静默丢失（第二十八轮独立审查实测：把 `links.jsonl` 换成目录，删除流程照样
 * 「成功」返回）。宁可这次调用失败，也不要假装成功。
 */
async function readTextStrict(path) {
    try {
        return await readFile(path, 'utf8');
    } catch (error) {
        if (error?.code === 'ENOENT') return null;
        throw new Error(`读不出 ${path}（${error?.code ?? error?.message ?? '未知错误'}）：为避免把它当成空文件、进而丢掉整份内容，这次调用中止。`);
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

/**
 * 同时持有**多个目录**的队列：按目录名排序后逐个进队。
 *
 * 为什么必须有它（第二十八轮独立审查挖出的最重一处）：跨根命中时，真正被写的是**另一个根**
 * 的条目与边，而只锁「被请求的那个根」等于没锁 —— 并发的 `link` / `importPack` 会用同一份
 * `links.jsonl` 的旧快照回写，把刚清掉的边复活（悬空引用）或把新边抹掉（静默丢写）；Windows 上
 * 两个并发 rename 打同一个目标还会直接抛 `EPERM`。实测（审查脚本）：跨根并发 60 轮有 59 轮
 * 出现悬空引用、33/40 轮 `EPERM`、7/40 轮两边都报成功却丢写；同根对照全部干净。
 *
 * **按字符串排序**是关键：所有多目录获取都走同一个全局顺序，才不会两个调用各持有对方要的
 * 那把锁形成环。已经持有其中某个目录的调用方**不要**再进这里（同目录嵌套会自锁）。
 */
function withQueues(dirs, task) {
    const keys = [...new Set(dirs.map((dir) => String(dir)))].sort();
    const acquire = (index) => (index >= keys.length ? task() : withQueue(keys[index], () => acquire(index + 1)));
    return acquire(0);
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
        // 这条结论依据哪些知识块（`kb:<docHash>:<n>`）：写入期从 source 里认出来。
        // 引用是**声明**的，不做推断 —— 与 entities 同一条纪律。
        kbRefs: stringList(raw.kbRefs),
        // 非知识块的材料来源（文件路径 / URL）：与 kbRefs 分开存，免得「依据」被 id 混住，
        // 也免得参数面承诺了 `source` 却在条目上什么都没留下（独立验证 F9）。
        sources: stringList(raw.sources),
        createdAt: asText(raw.createdAt) || nowIso(),
        updatedAt: asText(raw.updatedAt) || asText(raw.createdAt) || nowIso(),
    };
}

async function loadStore(paths) {
    const raw = await readJsonOrNull(paths.memory);
    const entries = asArray(raw?.entries).map((entry) => normalizeEntry(entry)).filter(Boolean);
    return {
        version: MEMORY_VERSION,
        // 老文件里这个键叫 `revision`；改名后仍要读得回来（改名不改算法，见 storeRevision）。
        contentRevision: asText(raw?.contentRevision) || asText(raw?.revision) || storeRevision(entries),
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
        contentRevision: storeRevision(entries),
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

/** 台账 `note` 的字符上限：它参与检索匹配，必须与条数一样有界。 */
const LEDGER_NOTE_CHARS = 400;

/**
 * 归档条目的规范化：补稳定 `id` 与关系边数。
 *
 * 为什么在读入与写入两处都过一遍：mnemon 迁来的条目当初只写了 `mnemonId`，没有 `id`
 * —— 于是 `entryIndex`（`if (id === '') continue`，`link` / `related` 的唯一入口）
 * 把它们全跳过，18 条历史长期记忆的图关系**整条断掉**；`renderArchiveDigest` 读的
 * 是 `linkedCount`，而迁移写的是 `links` 对象，于是人可读的摘要里关系边数**全空**
 * （第十九轮 P0-1 / 第二十四轮 P2-10，2026-09-27 复核仍成立）。
 *
 * 在读入时补一次，老库不必重迁：下一次写入会把补好的值固化进 `index.json`。
 */
/** 只接受真正的数字或能当数字用的非空字符串；`null` / `''` / 布尔一律算「没有」。 */
function finiteNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : undefined;
    if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value))) return Number(value);
    return undefined;
}

function normalizeArchiveItem(item) {
    const raw = item !== null && typeof item === 'object' ? item : {};
    const mnemonId = asText(raw.mnemonId);
    const fromLinks = raw.links !== null && typeof raw.links === 'object' ? finiteNumber(raw.links.total) : undefined;
    const linkedCount = finiteNumber(raw.linkedCount) ?? fromLinks;
    return {
        ...raw,
        id: asText(raw.id) || (mnemonId === '' ? '' : `mnemon:${mnemonId}`),
        ...(linkedCount === undefined ? {} : { linkedCount }),
    };
}

/**
 * 单卷归档的条目上限：超了就开下一卷（`2026-09.md` → `2026-09-2.md`）。
 *
 * 第二十四轮 P2-11 / 第一梯队 `14-1`：此前一个月只产生一个摘要文件，条数与字节
 * **月内无界** —— 一个高频月份能把单份 Markdown（以及那次写入要整份重写的索引）
 * 顶到几十万字节。分卷让「不丢」与「有界」两全：月份仍是检索与展示的口径，
 * 容量拐点落在卷上。
 *
 * 为什么是常量而不是设置项：它是**存储形状**的一部分，不是用户偏好；做成配置要同时
 * 动 config / 设置页 / 两侧测试，而调大调小都只在「单卷多大」上有差别，不影响语义。
 */
export const ARCHIVE_VOLUME_MAX_ITEMS = 200;

/** 卷名（也是两个投影文件的主干）：第 1 卷就是月份，第 2 卷起带序号。 */
export function archiveVolumeStem(month, volume) {
    const value = finiteNumber(volume);
    const index = value === undefined || value < 2 ? 1 : Math.trunc(value);
    return index > 1 ? `${month}-${index}` : month;
}

/** 从 `archive/<stem>.md` 取回主干（老索引只有 `file`，没有 `volume` / `itemFile`）。 */
function archiveStem(file) {
    const value = asText(file);
    if (value === '') return '';
    // 老索引里的 `file` 可能是 Windows 写法（`archive\2026-09.md`）：先统一分隔符，
    // 否则主干会带上整段反斜杠路径，进而写出 `archive/archive\2026-09.jsonl` 这种名字。
    const base = value.replace(/\\/g, '/').split('/').pop();
    return base.endsWith('.md') ? base.slice(0, -3) : base.endsWith('.jsonl') ? base.slice(0, -6) : base;
}

/** 从主干反推卷号：`2026-09` → 1、`2026-09-2` → 2（认不出就当第 1 卷）。 */
function archiveVolumeOf(stem, month) {
    if (stem === month) return 1;
    const tail = stem.startsWith(`${month}-`) ? stem.slice(month.length + 1) : '';
    const value = finiteNumber(tail);
    return value !== undefined && value >= 2 ? Math.trunc(value) : 1;
}

/**
 * 归档卷的规范化：补 `volume` / `itemFile` 两个新字段（第二十四轮 P2-11 / `24-9`）。
 *
 * 老库（本轮之前）的卷只有 `month` + `file: archive/2026-09.md`，正文内联在
 * `index.json` 里 —— 读入时照旧可用，**不必重迁**：下一次写入会把它固化成
 * `.jsonl` + 指针索引。第 1 卷的主干就是月份，所以老卷反推出来仍然是第 1 卷。
 */
function normalizeDigest(raw) {
    const digest = raw !== null && typeof raw === 'object' ? raw : {};
    const month = asText(digest.month) || monthOf(asText(digest.at) || nowIso());
    const stem = archiveStem(digest.file) || archiveVolumeStem(month, finiteNumber(digest.volume) ?? 1);
    const volume = finiteNumber(digest.volume) ?? archiveVolumeOf(stem, month);
    return {
        ...digest,
        month,
        volume: volume < 1 ? 1 : Math.trunc(volume),
        // 老索引里的 `file` 可能是 Windows 写法（`archive\2026-09.md`）：统一成 POSIX
        // 分隔符再落盘，否则 `itemFile` 会派生出 `archive/archive\2026-09.jsonl` 这种名字。
        file: (asText(digest.file) || `archive/${stem}.md`).replace(/\\/g, '/'),
        itemFile: (asText(digest.itemFile) || `archive/${stem}.jsonl`).replace(/\\/g, '/'),
        blocks: asArray(digest.blocks)
            .filter((block) => block !== null && typeof block === 'object')
            .map((block) => ({ ...block, items: asArray(block.items).map(normalizeArchiveItem) })),
    };
}

async function loadArchiveIndex(paths) {
    const raw = await readJsonOrNull(paths.index);
    const digests = asArray(raw?.digests)
        .filter((digest) => digest !== null && typeof digest === 'object')
        .map(normalizeDigest);
    return { version: MEMORY_VERSION, updatedAt: asText(raw?.updatedAt) || nowIso(), digests };
}

async function saveArchiveIndex(paths, index) {
    await writeAtomic(paths.index, `${JSON.stringify(index, null, 2)}\n`);
}

// ── 知识库（kb）的存储读写（第二十四轮 P1-7 一期） ──────────────────────────
//
// 与归档同一条纪律：**真源先落、索引后落**。manifest.jsonl 是文档层的真源，
// chunks/<docHash>/*.json 是正文层的真源，index.json 只是「计数与体积」的投影
// （它坏了可以从 manifest 重算，所以 `status` 读不到它时按 manifest 现算）。

/** manifest 是追加写的行文件；坏行跳过，不让半截行毁掉整份清单。 */
async function loadKbManifest(kbPaths) {
    const text = await readTextOrNull(kbPaths.manifest);
    if (text === null) return [];
    const rows = [];
    for (const line of text.split('\n')) {
        const trimmed = line.trim();
        if (trimmed === '') continue;
        try {
            const row = normalizeManifestRow(JSON.parse(trimmed));
            if (row !== null) rows.push(row);
        } catch {
            // 单行坏掉跳过。manifest 只是索引，正文在块文件里，丢一行不会丢正文。
        }
    }
    return rows;
}

async function saveKbManifest(kbPaths, rows) {
    await writeAtomic(kbPaths.manifest, rows.map((row) => `${JSON.stringify(row)}\n`).join(''));
}

async function loadKbIndex(kbPaths) {
    const raw = await readJsonOrNull(kbPaths.index);
    if (raw === null || typeof raw !== 'object') return null;
    const counts = raw.counts !== null && typeof raw.counts === 'object' ? raw.counts : null;
    if (counts === null) return null;
    return { version: KB_VERSION, updatedAt: asText(raw.updatedAt) || nowIso(), counts };
}

/**
 * 索引写回：计数由 manifest 现算（**不读正文**），所以任何时刻都能从真源重建。
 * 调用方拿到返回值就是这次落盘的计数。
 */
async function saveKbIndex(kbPaths, rows) {
    const counts = kbCountsOf(rows);
    await writeAtomic(kbPaths.index, `${JSON.stringify({ version: KB_VERSION, updatedAt: nowIso(), counts }, null, 2)}\n`);
    return counts;
}

/**
 * 一份文档的全部块，按块号升序。
 *
 * 两处「如实」：
 *   - 块文件缺失的不补空块（调用方按 `chunks_missing` 报出来）；
 *   - **自校验对不上的块当缺块处理**：正文被外面动过 / 截断时，宁可说「这一块读不到」，
 *     也不要拿半截内容当原文给模型（独立验证 F1③ 是同一条纪律的导入侧版本）。
 */
async function loadKbChunks(kbPaths, hash, chunks) {
    const dir = join(kbPaths.chunks, hash);
    const out = [];
    for (let n = 1; n <= chunks; n += 1) {
        const text = await readTextOrNull(join(dir, `${n}.json`));
        if (text === null) continue;
        try {
            const chunk = normalizeKbChunk(JSON.parse(text));
            if (chunk === null || chunk.n !== n || !chunkSelfConsistent(chunk)) continue;
            out.push(chunk);
        } catch {
            // 坏块跳过：缺的那块会在回执里体现为「块数对不上」。
        }
    }
    return out;
}

/** Pack 里的块文件键 → 内容（`<docHash>/<n>.json`）。 */
async function readKbChunkFiles(kbPaths, rows) {
    const files = {};
    for (const row of rows) {
        for (let n = 1; n <= row.chunks; n += 1) {
            const text = await readTextOrNull(join(kbPaths.chunks, chunkRelOf(row.hash, n)));
            if (text !== null) files[chunkRelOf(row.hash, n)] = text;
        }
    }
    return files;
}

/** 标题：Markdown 的一级标题优先，否则用文件名（manifest 里给人看的一行）。 */
function kbTitleOf(text, relPath) {
    const match = /^[ \t]{0,3}#[ \t]+(\S.*?)[ \t]*$/m.exec(String(text ?? ''));
    if (match !== null) return match[1].trim();
    return basename(relPath) || relPath;
}

/** 索引里的**指针块**：条目 id 留在索引里，`liveIds` 这类只要 id 的调用就不必读正文。 */
function pointerBlock(block) {
    const items = asArray(block.items);
    return {
        id: asText(block.id),
        at: asText(block.at),
        reason: asText(block.reason),
        count: items.length,
        ids: items.map((item) => asText(item.id)).filter((id) => id !== ''),
    };
}

/** 一块归档里的条目 id：新格式只有指针 `ids`，老格式内联在 `items` 里。 */
function blockIds(block) {
    const ids = asArray(block.ids).map((id) => asText(id)).filter((id) => id !== '');
    if (ids.length > 0) return ids;
    return asArray(block.items).map((item) => asText(item?.id)).filter((id) => id !== '');
}

/**
 * 归档块的**内容寻址 id**：`b-` + 「时刻 + 原因 + 条目」的 sha1 前 12 位。
 *
 * 为什么用内容而不是「时刻+原因+条数」那类形状键（第四十轮独立验证挖出的三条洞）：
 *
 * 1. **墓碑过滤会改条数**：Pack 里的块声明 2 条、其中 1 条已被墓碑挡掉，按**落盘后**
 *    的条数当 id，第二次导入算出来的键就与库里那条不同 → 同一条内容被搬两份。
 *    内容 id 由**入包声明的内容**算（调用方显式传进来），过滤前后都是同一个 id。
 * 2. **形状键会跨月撞**：`at|reason|count` 在两个月里可能一样，靠「全库唯一」加后缀
 *    的结果是同一份 Pack 换个导入顺序就重复。
 * 3. **形状键撞车时是静默丢内容**：同月两块真的同键，块 id 相同 → 第二块被去重逻辑
 *    判成「已存在」直接跳过（旧行为是可见的重复）。内容不同则 id 必不同，两条都在。
 *
 * id 只要求**同一内容稳定**、不同内容不撞：读侧从不解析它，图关系挂的是条目 id。
 */
function archiveBlockId(block) {
    return `b-${sha1(JSON.stringify({
        at: asText(block.at),
        reason: asText(block.reason),
        items: asArray(block.items),
    })).slice(0, 12)}`;
}

/**
 * 只看**条目内容**的签名：合并 `.jsonl` 与老索引内联正文时用来认「同一块」。
 *
 * 为什么不能只看 `blockKey`：老索引里的块常常没有 id（键退化成内容+时刻+原因），
 * 而 `.jsonl` 里同一块带着 id —— 两边算出的身份不同，于是同一条内容会留两份。
 * 内容签名把这种「同一块、两种形态」认出来（第四十轮验证 probe-c C2）。
 */
function blockContentKey(block) {
    return archiveBlockId({ at: '', reason: '', items: asArray(block.items) });
}

/** 一条指针块宣告的条目数（老格式按内联条目算）。 */
function blockCount(block) {
    const value = finiteNumber(block.count);
    if (value !== undefined) return value;
    return asArray(block.items).length;
}

/** 把 `.jsonl` 正文解析成块；单行坏掉跳过，其余照读（与台账同一口径）。 */
function parseDigestItemText(text) {
    const blocks = [];
    for (const line of String(text ?? '').split('\n')) {
        const trimmed = line.trim();
        if (trimmed === '') continue;
        try {
            const raw = JSON.parse(trimmed);
            if (raw === null || typeof raw !== 'object') continue;
            blocks.push({
                id: asText(raw.id),
                at: asText(raw.at),
                reason: asText(raw.reason),
                items: asArray(raw.items).map(normalizeArchiveItem),
            });
        } catch {
            // 与台账一样：一行坏了不该让整卷作废。
        }
    }
    return blocks;
}

/** 一卷归档的 `.jsonl` 文本：一块一行（与索引里的指针块一一对应）。 */
function renderDigestItems(blocks) {
    return asArray(blocks).map((block) => `${JSON.stringify({
        id: asText(block.id),
        at: asText(block.at),
        reason: asText(block.reason),
        items: asArray(block.items),
    })}\n`).join('');
}

/**
 * 一卷归档的**全部块（含正文）**—— 读归档正文的唯一入口。
 *
 * 正文只有两个落点：人读的 `.md` 与结构化的 `.jsonl`；`index.json` 只存指针。
 * 这里优先读 `.jsonl`（**真源**），读不到才退回索引里内联的老格式正文。
 *
 * 两种形态同时存在（`.jsonl` 已落、索引还没落就崩了；或从备份恢复了老索引）时，
 * **以 `.jsonl` 为准，再把索引里多出来的块并回来** —— 反过来「有内联正文就不看
 * `.jsonl`」会把崩掉那一批静默丢掉，而之后每一次写入都会按索引重写 `.jsonl`，
 * 丢掉的正文再也回不来（第四十轮独立验证 probe-c C2 实测）。
 *
 * `damaged` 的判据是「**索引声明了几条、正文实际给了几条**」，不是「文件在不在」：
 *   - 正文比索引**多**（索引落后）不算受损 —— 那是可以自愈的状态，调用方会把多出来的
 *     块并入索引；
 *   - 正文比索引**少**（文件被删、被截断、有坏行）才算受损。写入路径对这种卷**只隔离、
 *     不重写**：拿指针块去当正文渲染，等于把「读不到」变成「真的没了」，
 *     连唯一还留着正文的那份 `.md` 也会被空摘要覆盖（probe-e E2 实测）。
 */
async function loadDigestBlocks(paths, digest) {
    const inline = asArray(digest.blocks);
    const declared = inline.reduce((sum, block) => sum + blockCount(block), 0);
    const text = await readTextOrNull(join(paths.archive, basename(asText(digest.itemFile))));
    const stored = text === null ? [] : parseDigestItemText(text);
    const inlineHasItems = inline.some((block) => asArray(block.items).length > 0);
    // 返回值一律是**副本**：调用方（appendArchive）会往这份「正文」里追加块，如果它就是
    // `digest.blocks` 本身，索引指针会被一起追加进正文里 —— 正文里混进没有 items 的
    // 指针块，渲染与检索都会崩（本轮实测踩到过）。
    if (inlineHasItems) {
        const keys = new Set();
        for (const block of stored) {
            keys.add(blockKey(block));
            keys.add(blockContentKey(block));
        }
        const merged = stored.map((block) => ({ ...block }));
        for (const block of inline) {
            if (asArray(block.items).length === 0) continue;
            const key = blockKey(block);
            const content = blockContentKey(block);
            // 身份或内容任一相同就算同一块：老索引里那块没有 id，`.jsonl` 里那块有，
            // 只按身份去重会留下两份同内容块。
            if (keys.has(key) || keys.has(content)) continue;
            keys.add(key);
            keys.add(content);
            merged.push({ ...block });
        }
        // `.jsonl` 还不存在（老格式）：merged 就是内联的那份，交给下一次写入固化。
        return { blocks: merged.length > 0 ? merged : inline.map((block) => ({ ...block })), legacy: true, damaged: false };
    }
    const available = stored.reduce((sum, block) => sum + asArray(block.items).length, 0);
    const damaged = declared > available;
    return { blocks: stored.length > 0 || text !== null ? stored : inline.map((block) => ({ ...block })), legacy: false, damaged };
}

/**
 * 一个记忆根里的全部归档条目（含正文）与卷数。
 *
 * 索引只给指针（`loadArchiveIndex` 便宜），正文按卷从 `.jsonl` 读 —— 需要正文的
 * 调用点（检索 / 实体 / 关系索引 / 导出）都走这里，不要各自去翻 `digest.blocks`。
 */
async function harvestArchive(paths, origin) {
    const index = await loadArchiveIndex(paths);
    const items = [];
    let damaged = 0;
    for (const digest of index.digests) {
        const loaded = await loadDigestBlocks(paths, digest);
        if (loaded.damaged) damaged += 1;
        for (const block of loaded.blocks) {
            for (const item of asArray(block.items)) {
                // reason 带下去（第四十八轮 新-10）：归档读侧与 `sunk` 都要按它区分
                // 「热记忆下沉」与「台账滚动 / mnemon 迁移 / Pack 导入」。
                items.push({
                    ...item,
                    origin,
                    reason: asText(block.reason) || 'external',
                    blockAt: asText(block.at),
                    month: digest.month,
                    volume: digest.volume,
                    file: digest.file,
                    blockId: block.id,
                });
            }
        }
    }
    return { items, volumes: index.digests.length, damaged };
}

/**
 * 一卷的正文文件是不是**给不出内容**（给 `status` 用的廉价近似）。
 *
 * 只 stat 体积，**不读正文** —— `status` 的定位就是「不读全文」。所以：
 *   - 老格式（正文还内联在索引里）**永远不算受损**：它读得出来，只是还没固化。
 *     不排除这一条，每个还没迁移过的库都会被报成「有一卷受损」（第四十轮第二轮
 *     独立验证 P3 实测：真实库 read 说 88/88 可读，status 却说受损 1 卷）；
 *   - 文件不在、或 0 字节（被截断）都算受损；
 *   - **坏行不算**：那要读正文才看得出，`read` 会报 `damagedVolumes`，
 *     `status` 这个近似值不会。两个口径的差别写在这里，别当成 bug。
 */
function digestBodyMissing(paths, digest) {
    const blocks = asArray(digest.blocks);
    const declared = blocks.reduce((sum, block) => sum + blockCount(block), 0);
    if (declared === 0) return false;
    if (blocks.some((block) => asArray(block.items).length > 0)) return false;
    try {
        return statSync(join(paths.archive, basename(asText(digest.itemFile)))).size === 0;
    } catch {
        return true;
    }
}

// ── 图关系（links.jsonl） ───────────────────────────────────────────────────

function normalizeLink(raw) {
    if (raw === null || typeof raw !== 'object') return null;
    const sourceId = asText(raw.sourceId);
    const targetId = asText(raw.targetId);
    if (sourceId === '' || targetId === '' || sourceId === targetId) return null;
    const kind = LINK_KINDS.includes(raw.kind) ? raw.kind : DEFAULT_LINK_KIND;
    return {
        id: asText(raw.id) || `K-${randomUUID().slice(0, 8)}`,
        sourceId,
        targetId,
        kind,
        // 冲突边的三态（第二十四轮 §5.3 / P1-8）：只有 `contradicts` 才带这两个字段。
        // 不做自动裁决 —— ContextConflict 显示模型系统性偏好更早的证据，ARR 证明静态
        // 对比解码必有一侧错；这里只把「冲突属于哪一类、现在偏向谁」**如实记下来**，
        // 让下一次读到它的人（模型或用户）看到「未决」而不是一个被偷偷选好的结论。
        conflictClass: kind === 'contradicts' ? linkConflictClass(raw.conflictClass) : '',
        conflictState: kind === 'contradicts' ? linkConflictState(raw.conflictState) : '',
        note: asText(raw.note),
        at: asText(raw.at) || nowIso(),
    };
}

async function loadLinks(paths) {
    const text = await readTextStrict(paths.links);
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
    const existing = await readTextStrict(paths.links);
    const prefix = existing === null || existing === '' ? '' : existing.endsWith('\n') ? existing : `${existing}\n`;
    await writeAtomic(paths.links, `${prefix}${links.map((link) => `${JSON.stringify(link)}\n`).join('')}`);
}

/**
 * 删边：把两端落在 `deadIds` 里的关系从 `links.jsonl` 里移掉。
 *
 * **悬空边比悬空条目更糟**（第二十四轮 §5.6 / P1-9）：悬空条目只是查不到，
 * 而悬空边会让 `related` 静默少一条，旧实现还会**穿过**已经不存在的节点继续往外走
 * —— 于是「删掉一条记忆」反而成了两个本来无关的条目之间的桥。
 * 删除与清边要在**同一次调用**里做完（中间不让别的写入插进来）；但**不要**为它再排一次
 * `withQueue`：同一个目录嵌套排队会自锁（`previous.then(task, task)`），跨根时更是如此
 * —— 详见 `mutate` 里 remove 分支的注释。
 */
async function pruneLinks(paths, deadIds) {
    if (deadIds.size === 0) return 0;
    const links = await loadLinks(paths);
    const kept = links.filter((link) => !deadIds.has(link.sourceId) && !deadIds.has(link.targetId));
    const removed = links.length - kept.length;
    if (removed > 0) await writeAtomic(paths.links, kept.map((link) => `${JSON.stringify(link)}\n`).join(''));
    return removed;
}

/** 数一批边里有多少悬空引用（不变量：**必须 == 0**）。 */
function danglingOf(links, liveIds) {
    return links.filter((link) => !liveIds.has(link.sourceId) || !liveIds.has(link.targetId));
}

// ── 墓碑（tombstones.jsonl） ─────────────────────────────────────────────────

/**
 * 墓碑：删掉的条目只留 `id + 时刻 + 原因`，正文不进。
 *
 * 为什么删干净了还要留一条：没有墓碑时，「被删掉」与「从来不存在」在下游是同一种形状
 * —— 一次 Pack 导入就能把删掉的条目原样搬回来（GateMem 的结论是检索式记忆**仍会泄漏
 * 已删除的信息**），而 `link` 也只能说「找不到这个 id」。墓碑让这两件事都能被如实说清。
 */
function normalizeTombstone(raw) {
    if (raw === null || typeof raw !== 'object') return null;
    const id = asText(raw.id);
    if (id === '') return null;
    return { id, at: asText(raw.at) || nowIso(), reason: asText(raw.reason) };
}

async function loadTombstones(paths) {
    const text = await readTextStrict(paths.tombstones);
    if (text === null) return [];
    const out = [];
    for (const line of text.split('\n')) {
        const trimmed = line.trim();
        if (trimmed === '') continue;
        try {
            const stone = normalizeTombstone(JSON.parse(trimmed));
            if (stone !== null) out.push(stone);
        } catch {
            // 与台账、边同一口径：单行坏掉跳过，其余照读。
        }
    }
    return out;
}

async function appendTombstones(paths, records) {
    const clean = asArray(records).map(normalizeTombstone).filter((item) => item !== null);
    if (clean.length === 0) return 0;
    await mkdir(paths.dir, { recursive: true });
    const existing = await readTextStrict(paths.tombstones);
    // 按 id 去重：同一个 id 可能先被 remove、后被归档封顶再判一次死；墓碑只增不删，
    // 但**不该长两行**（否则 status.tombstones 会把同一个 id 数成两条）。
    const known = new Set((await loadTombstones(paths)).map((stone) => stone.id));
    const fresh = clean.filter((stone) => !known.has(stone.id));
    if (fresh.length === 0) return 0;
    const prefix = existing === null || existing === '' ? '' : existing.endsWith('\n') ? existing : `${existing}\n`;
    await writeAtomic(paths.tombstones, `${prefix}${fresh.map((stone) => `${JSON.stringify(stone)}\n`).join('')}`);
    return fresh.length;
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
    // 分卷之后同一个月会有多份摘要：标题里写出卷号，否则两个文件看起来一模一样
    // （第 1 卷不加后缀，老库的标题保持原样，历史文档里的引用不会失效）。
    const title = (finiteNumber(digest.volume) ?? 1) > 1 ? `${digest.month}（第 ${digest.volume} 卷）` : digest.month;
    const lines = [`# 归档 ${title}`, ''];
    for (const block of digest.blocks) {
        lines.push(`## ${block.at} · ${archiveSection(block.reason)}（${asArray(block.items).length} 条）`, '');
        for (const item of asArray(block.items)) {
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

/**
 * 归档块的稳定标识：有 id 用 id，没有就按内容算（与写入时的生成口径同一处）。
 *
 * 与写入共用同一个函数是有意的：导入去重是「入包块的键」对「库里块的键」，
 * 两边算法一旦不同就再也对不上（第四十轮独立验证实测过这条）。
 */
function blockKey(block) {
    return asText(block.id) || archiveBlockId(block);
}

/**
 * 追加一批归档：同一个月的**当前卷**里追加，块数超上限就开下一卷。
 *
 * 归档的正文有两个投影，写入顺序与它们的关系是这一段的全部要点：
 *
 *   `archive/<stem>.jsonl`  结构化的真源（块 + 条目正文，一块一行）
 *   `archive/<stem>.md`     人可读的渲染（grep 得到、打不开 JSON 也能读）
 *   `archive/index.json`    只存指针：卷、块、条目数、条目 id 列表
 *
 * **索引不再内联正文**（第二十四轮 P2-11 / 第一梯队 `24-9`）：此前它整份内联了
 * 全部条目正文，于是「索引比正文文件还大」（实测 157 181 B / 97 502 B），而且
 * `liveIds` / `entryIndex` 这类**只要 id** 的调用每次都要解析那几十万字节。现在
 * 索引只随「卷数 + 块数 + 条目数」增长，不随正文长度增长。
 *
 * 顺序仍是「真源先改、投影后删」：先落 `.jsonl`，再落索引，最后才删被挤掉的卷 ——
 * 反过来的话，一步失败就会留下「索引还列着、正文已经没了」的半状态。
 */
async function appendArchive(paths, blocks, memoryConfig) {
    if (blocks.length === 0) return [];
    await mkdir(paths.archive, { recursive: true });
    const index = await loadArchiveIndex(paths);
    // 先把每一卷的正文读回来（新格式在 .jsonl，老格式内联在索引里）：写回时索引里的
    // 指针块与这份正文必须一一对应，所以读写都在这一份上做。
    // `damaged` 的卷这一轮**不碰正文**（见 loadDigestBlocks 的注释）。
    const loaded = new Map();
    for (const digest of index.digests) loaded.set(digest, await loadDigestBlocks(paths, digest));

    const created = [];
    for (const block of blocks) {
        const month = block.month ?? monthOf(block.at);
        const items = asArray(block.items).map(normalizeArchiveItem);
        const volumes = index.digests.filter((item) => item.month === month).sort((left, right) => left.volume - right.volume);
        let digest = volumes[volumes.length - 1];
        const latest = digest === undefined ? undefined : loaded.get(digest);
        const currentCount = digest === undefined ? 0 : asArray(latest?.blocks).reduce((sum, row) => sum + asArray(row.items).length, 0);
        // 分卷在两处发生：当前卷**满了**，或当前卷**受损**（隔离它：往一个正文读不全的
        // 卷里追加，等于用「指针 + 新块」重写它，把读不到的那些条真正抹掉）。
        // 注意**一块本身超上限也不切开** —— 块是去重与导入的单位，切开就再也对不上了
        // （一次台账滚动可能带几百条）。
        if (digest !== undefined && (latest?.damaged === true || (currentCount > 0 && currentCount + items.length > ARCHIVE_VOLUME_MAX_ITEMS))) {
            digest = undefined;
        }
        if (digest === undefined) {
            const volume = (volumes[volumes.length - 1]?.volume ?? 0) + 1;
            const stem = archiveVolumeStem(month, volume);
            digest = {
                id: `a-${stem}-${randomUUID().slice(0, 6)}`,
                month,
                volume,
                at: block.at,
                file: `archive/${stem}.md`,
                itemFile: `archive/${stem}.jsonl`,
                blocks: [],
            };
            index.digests.push(digest);
            loaded.set(digest, { blocks: [], legacy: false, damaged: false });
        }
        const written = {
            // 块 id：调用方给了就用它（Pack 导入会把**入包声明的内容**算出的 id 传进来，
            // 这样墓碑过滤掉几条也不会改变它的身份），没给就按内容算（见 archiveBlockId）。
            id: asText(block.id) || archiveBlockId({ at: block.at, reason: block.reason, items }),
            at: block.at,
            reason: block.reason,
            // 写进去的条目也过一遍规范化：id 与边数在同一处补齐，读侧就不必各写一份兜底。
            items,
        };
        loaded.get(digest).blocks.push(written);
        // 指针同步进索引：`deadAmong` 要在**这次调用之后**的索引上判活，新块必须
        // 已经在里面 —— 否则「固定 id 的条目被重建」会被判成死的、边被误删。
        digest.blocks.push(pointerBlock(written));
        digest.at = block.at;
        created.push(digest);
    }
    // 归档本身也封顶：最旧的**卷**被删掉，索引里同步移除，不留「指不到的条目」。
    // 排序键是 month + volume 而不是 at：at 是「这次下沉发生的时刻」，同一秒里的多个卷
    // at 完全相同，按它排序等于按写入顺序裁剪，剪掉哪一个是不确定的。
    const keep = memoryConfig.archiveKeep;
    index.digests.sort((left, right) => String(left.month).localeCompare(String(right.month)) || (left.volume - right.volume));
    const dropped = index.digests.length > keep ? index.digests.splice(0, index.digests.length - keep) : [];
    // 删之前先把正文摘出来：`loaded` 是这次写入的真源，接下来两步（判死、落盘）都要用它。
    const droppedRows = new Map(dropped.map((digest) => [digest, loaded.get(digest)?.blocks ?? []]));
    for (const digest of dropped) loaded.delete(digest);
    // 「封顶删最旧卷」是**真的删除**（不是下沉）：正文文件删掉、边同事务清掉、留墓碑。
    // 三件事必须一起做 —— 只删文件会留下指不到任何东西的边，只删边则删除变成假的
    // （正文还在盘上，而归档仍可被召回）。
    //
    // 顺序也是有意写的（第二十七轮审查挖出的第 4 条缺陷）：**先墓碑 + 清边，再落正文
    // 与索引，最后删文件**。反过来的话，一步失败就会留下「正文已删、索引仍列着它」的
    // 半状态 —— 而 read(archive) / entryIndex / liveIds 的唯一真源就是索引，于是条目
    // 照样被召回、悬空引用照样是 0，谁也看不出正文已经没了。现在失败时最坏是
    // 「墓碑写了、卷没剪成」，下一次再剪即可（墓碑只增不删，重复剪不会长第二条）。
    //
    // **只判真正死掉的 id**：同一个 id 可能仍活在热记忆或台账里（例如固定 id 的
    // `pointer:archive` 会在下次溢出时被重建）。那种情况下删边是**删错了**，
    // 于是「悬空引用 == 0」是拿删掉合法边换来的。
    const droppedIds = new Set();
    for (const digest of dropped) {
        for (const block of droppedRows.get(digest) ?? []) {
            for (const id of blockIds(block)) droppedIds.add(id);
        }
    }
    const dead = await deadAmong(paths, droppedIds, index);
    if (dead.size > 0) {
        await appendTombstones(paths, [...dead].map((id) => ({ id, reason: 'archive-keep' })));
        await pruneLinks(paths, dead);
    }
    // 摘要 Markdown 与索引同一次写入更新，但**不能只写「这次被触碰到的卷」**：
    // 规范化（补 id / 补边数）可能改的是几个月前那份摘要，而那份 .md 只有被重写才会
    // 带上新字段 —— 只写 created 的话，老月份的「关系边 N 条」永远是空的（第十九轮
    // P0-1 报的就是这个症状，第二十七轮审查发现只修 index 不够）。
    // 做法是「渲染一遍、与盘上比对、不同才写」：幂等、自纠正，正常情况下一个字节都不写。
    for (const digest of index.digests) {
        const state = loaded.get(digest) ?? { blocks: [], damaged: false };
        // **受损的卷只隔离、不重写**：它的正文（`.jsonl` / `.md`）可能是唯一还留着内容的
        // 地方，拿指针块去渲染会把它覆盖成空摘要，同时把 `damagedVolumes` 洗成 0
        // ——「读不到」就这么变成了「真的没了」，还顺手把引用它的边变成悬空边
        // （第四十轮独立验证 probe-e E2 / probe-g 实测）。索引里的指针原样保留，
        // 让 damage 一直报得出来，等人工处置。
        if (state.damaged) continue;
        const rows = state.blocks;
        // 老库的块可能没有 id：落盘时补一个内容寻址的 id（与导入去重同一口径）。
        for (const block of rows) {
            if (asText(block.id) === '') block.id = archiveBlockId(block);
        }
        const itemsText = renderDigestItems(rows);
        const itemsPath = join(paths.archive, basename(asText(digest.itemFile)));
        if ((await readTextOrNull(itemsPath)) !== itemsText) await writeAtomic(itemsPath, itemsText);
        const target = join(paths.archive, basename(asText(digest.file)));
        const next = renderArchiveDigest({ ...digest, blocks: rows });
        if ((await readTextOrNull(target)) !== next) await writeAtomic(target, next);
        // 索引只留指针：正文到这里为止都不再进 index.json。
        digest.blocks = rows.map(pointerBlock);
    }
    index.updatedAt = nowIso();
    await saveArchiveIndex(paths, index);
    // 索引落定之后才删正文：真源先改，投影后删。**两个投影一起删** ——
    // 只删 .md 而把 .jsonl 留在盘上，正文就还读得到，「删最旧卷」就是假删除
    // （第二十四轮 §5.6：归档仍参与召回，所以删除必须落到召回面上）。
    for (const digest of dropped) {
        await rm(join(paths.archive, basename(asText(digest.file))), { force: true });
        await rm(join(paths.archive, basename(asText(digest.itemFile))), { force: true });
    }
    return created;
}

/**
 * 一批候选 id 里**真正死掉**的那些：不再出现在热记忆、台账或（更新后的）归档索引里。
 * 封顶剪枝与墓碑都要用它 —— 只按「这个月被剪了」判死，会把还活在别处的 id 误杀。
 */
async function deadAmong(paths, candidates, index) {
    const alive = new Set();
    for (const entry of (await loadStore(paths)).entries) alive.add(entry.id);
    for (const record of await loadLedger(paths)) {
        const id = asText(record.id);
        if (id !== '') alive.add(id);
    }
    for (const digest of asArray(index?.digests)) {
        for (const block of asArray(digest.blocks)) {
            for (const id of blockIds(block)) alive.add(id);
        }
    }
    return new Set([...candidates].filter((id) => !alive.has(id)));
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
        // 用过的知识块 id 也进检索串：台账 → 知识块这根线要能反着查
        // （「哪份交付物用过这块来源」）。它们不是文件来源，所以另存一个字段。
        asArray(record.kbRefs).join(' '),
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

/**
 * 一条热记忆参与话题信号匹配的文本：正文 + 实体 + 标签。
 *
 * 实体与标签是用户写条目时**显式声明**的归类（「这条在讲什么」），比正文更接近
 * 话题；正文兜底，覆盖没填实体标签的条目。
 */
function entrySignalText(entry) {
    return `${entry.content} ${asArray(entry.entities).join(' ')} ${asArray(entry.tags).join(' ')}`;
}

// ── recency 弱先验（第二十四轮 24-4） ────────────────────────────────────────

/** 每日衰减系数：0.995（第二十四轮点名的那个数）。 */
export const RECENCY_DECAY_PER_DAY = 0.995;

/**
 * 一条记忆的「时间新鲜度」先验，落在 (0, 1]。
 *
 * 三条口径，写清楚免得被当成第二个相关性分：
 *   1. **它只用来定序，不参与取舍**。热记忆的读取按「重要度 → 这个先验」排序 ——
 *      重要度永远是主键，先验只在**同重要度**里决定谁在前。所以它不会让 critical
 *      被 normal 的新条目挤下去。
 *   2. **量的是写入时刻（`updatedAt`），不是取用时刻**。同一事实的两版并存时，
 *      「后写的那版是当前事实」（`supersedes` 的语义），这是可以量到增益的那件事
 *      （第四十二轮量表：同族两版并存时最新版在前 0/4，纯时间倒序上界 4/4）。
 *      取用时刻（`lastAccessAt`）至今没有用处 —— 记录它会让**读变成写**，而
 *      「读不写盘」是面板、只读探针与真库对照共同依赖的性质（24-3r 的判据）。
 *   3. **单调 ⇒ 同重要度内等价于「新的在前」**。用指数衰减而不是线性，是为了让这个
 *      先验的量级**弱**：将来若与相关性分相加，它不会把相关性差异盖过去
 *      （一年前的条目仍有 0.995^365 ≈ 0.16，而不是归零）。
 *
 * 时间戳解析不出来时返回 0（排在同重要度的最后），而不是「当作最新」——
 * 拿不出证据的条目不该冒到前面。
 */
export function recencyPriorOf(updatedAt, now = Date.now()) {
    const value = typeof updatedAt === 'number' ? updatedAt : Date.parse(asText(updatedAt));
    if (!Number.isFinite(value)) return 0;
    const days = Math.max(0, (now - value) / 86_400_000);
    return RECENCY_DECAY_PER_DAY ** days;
}

/**
 * 切词：非 CJK 片段保持整词，CJK 片段展开成字符二元组（bigram），最后去重。
 *
 * 为什么必须这样切：中文查询通常不带空格，原实现按空白/标点切完就是**一个整串**，而
 * `matchScore` 要求整串出现才计分 —— 于是 `relevanceOf = score / termCount` 只能是
 * 0 或 1，低分丢掉、中档与未知档的名额上限、候选池倍数这一整套机械**一次都不触发**。
 * 实测（第二十四轮，16 条人工标注中文查询，台账层）：原实现 hit@3 只有 **8/16**、
 * 8 条查询零结果；只把切词换成二元组、阈值与其余机械一个字不动 → hit@3 **16/16**、
 * 零结果 **0**。典型救回来的查询：`面板记忆`（词序颠倒）、`结合知识库`、
 * `音视频`（正文写的是「音频与视频」）、`提示词准确性`。
 *
 * 粒度依据：LeCaRDv2（3,795 文档 / 159 查询）nDCG@10 —— 完全不切 0.359 ＜ 字符级 0.567
 * ＜ 子词 0.631 ＜ Jieba 词级 0.641；手写字符级相对词级只差约 11.5%，而相对不切高 0.208。
 * SQLite FTS5 官方文档也说明「所有非 ASCII 字符（码点 >127）始终被视为 token 字符」，
 * 默认 unicode61 分析器对中文基本失效 —— **中文必须自己切**，而切得粗一点没关系。
 *
 * 实现搬到了 `kb.js` 的 `kbTerms`（第二十四轮 24-11）：kb 的 BM25 排序要的是**同一套**
 * 词表。两处各留一份的代价是「记忆侧能召回、kb 侧召不回」这类只有真实语料才暴露的漂移，
 * 所以这里只做转发 —— 记忆侧的检索行为逐字未变（`memory-recall` / `memory` 套件钉住）。
 */
function queryTerms(query) {
    return kbTerms(query);
}

function truncate(text, maximum) {
    const value = String(text);
    if (value.length <= maximum) return value;
    return `${value.slice(0, Math.max(1, maximum - 1))}…`;
}

/** 本地时刻的紧凑写法（时间窗回执用）。 */
function localStamp(ms) {
    const value = new Date(ms);
    const pad = (number) => String(number).padStart(2, '0');
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ${pad(value.getHours())}:${pad(value.getMinutes())}`;
}

// ── 时间窗（第二十四轮 P1-5 / 第二十六轮 24-3） ──────────────────────────────

/**
 * 时间维度的**过滤入口**：`read` 的 `since` / `until`。
 *
 * 为什么要有它：在它之前，整个记忆系统没有任何按时间缩小搜索空间的入口 ——
 * 用户问「本月记的偏好」「上周那份交付物」时，召回机械只能把整库拿来做词法匹配，
 * 时间只是排序并列时的兜底键（第二十四轮 §2.3 实测）。
 *
 * 三条实现口径，都是为了避免**假的「没有」**：
 *   1. 解析不出来就**抛错**，不静默忽略 —— 静默忽略会把「没命中」说成「那段时间没有」；
 *   2. 按**本地时区**解析（`2026-09-27` 指本地那一天），带 `Z` / 偏移量的 ISO 串按绝对时刻；
 *   3. 只给日期时，`since` 取当天 00:00:00、`until` 取当天 23:59:59.999（闭区间）；
 *      给了时分秒就按那个时刻算。
 *
 * `lastAccessAt`（「上次取用」）**故意不在这一轮做**：它只有在 recency 作为排序项时才有
 * 消费者，而第二十四轮 P1-6 明确要求「权重未量到增益就不落」（Zep 的反例是时间重排让
 * single-session-assistant 掉 17.7%）。更硬的一条理由是：记录 lastAccessAt 会让 `read`
 * 变成写操作，而「读不写盘」是面板、只读探针与真库对照表共同依赖的性质。
 */
function startOfDay(date) {
    const value = new Date(date);
    value.setHours(0, 0, 0, 0);
    return value;
}

function endOfDay(date) {
    const value = new Date(date);
    value.setHours(23, 59, 59, 999);
    return value;
}

/** 周一为一周的第一天（与中文「本周」的日常读法一致）。 */
function startOfWeek(date) {
    const value = startOfDay(date);
    value.setDate(value.getDate() - ((value.getDay() + 6) % 7));
    return value;
}

function startOfMonth(date) {
    const value = startOfDay(date);
    value.setDate(1);
    return value;
}

function addDays(date, days) {
    const value = new Date(date);
    value.setDate(value.getDate() + days);
    return value;
}

function addMonths(date, months) {
    const value = new Date(date);
    value.setMonth(value.getMonth() + months, 1);
    return startOfDay(value);
}

const ABSOLUTE_ISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:\d{2})$/;
const LOCAL_DATE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
// 前缀与数字之间允许空格：`最近 3 个月` 与 `最近3个月` 是同一件事（审查提的疑点）。
const RELATIVE_DAYS = /^(?:最近|近|过去|last)?\s*(\d+)\s*(天|日|days?)(?:内|以内|之内)?$/;
const RELATIVE_MONTHS = /^(?:最近|近|过去|last)?\s*(\d+)\s*(个?月|months?)(?:内|以内|之内)?$/;
const RELATIVE_HOURS = /^(?:最近|近|过去|last)?\s*(\d+)\s*(个?小时|hours?|h)(?:内|以内|之内)?$/;

/**
 * 时间词表：键 = 能认的字，值 = 相对 `now` 的区间。
 *
 * 解析与报错提示**共用这一份**（第四十七轮 18-23 的整改）：提示里列的字必须真的能解析，
 * 能解析的也要列出来 —— 提示原来手写了 9 个，而这张表里有 今天 / today / 昨天 / yesterday /
 * 前天 / 明天 / 本周 / 这周 / this week / 上周 / last week / 本月 / 这个月 / this month /
 * 上月 / 上个月 / last month / 今年 / 去年，实测 `until:'明天'` 能解析、提示里却没有它。
 * （提示词只列中文常用那几个，见 {@link MEMORY_TIME_HINTS}；同义写法在这里都认。）
 */
const TIME_WORD_TABLE = {
    今天: (now) => [startOfDay(now), endOfDay(now)],
    today: (now) => [startOfDay(now), endOfDay(now)],
    昨天: (now) => [startOfDay(addDays(now, -1)), endOfDay(addDays(now, -1))],
    yesterday: (now) => [startOfDay(addDays(now, -1)), endOfDay(addDays(now, -1))],
    前天: (now) => [startOfDay(addDays(now, -2)), endOfDay(addDays(now, -2))],
    明天: (now) => [startOfDay(addDays(now, 1)), endOfDay(addDays(now, 1))],
    本周: (now) => [startOfWeek(now), endOfDay(addDays(startOfWeek(now), 6))],
    这周: (now) => [startOfWeek(now), endOfDay(addDays(startOfWeek(now), 6))],
    'this week': (now) => [startOfWeek(now), endOfDay(addDays(startOfWeek(now), 6))],
    上周: (now) => [startOfWeek(addDays(now, -7)), endOfDay(addDays(startOfWeek(addDays(now, -7)), 6))],
    'last week': (now) => [startOfWeek(addDays(now, -7)), endOfDay(addDays(startOfWeek(addDays(now, -7)), 6))],
    本月: (now) => [startOfMonth(now), endOfDay(addDays(addMonths(now, 1), -1))],
    这个月: (now) => [startOfMonth(now), endOfDay(addDays(addMonths(now, 1), -1))],
    'this month': (now) => [startOfMonth(now), endOfDay(addDays(addMonths(now, 1), -1))],
    上月: (now) => [addMonths(now, -1), endOfDay(addDays(startOfMonth(now), -1))],
    上个月: (now) => [addMonths(now, -1), endOfDay(addDays(startOfMonth(now), -1))],
    'last month': (now) => [addMonths(now, -1), endOfDay(addDays(startOfMonth(now), -1))],
    今年: (now) => [startOfDay(new Date(now.getFullYear(), 0, 1)), endOfDay(new Date(now.getFullYear(), 11, 31))],
    去年: (now) => [startOfDay(new Date(now.getFullYear() - 1, 0, 1)), endOfDay(new Date(now.getFullYear() - 1, 11, 31))],
};

/** 时间词表里全部能认的字（测试与提示共用；大小写不敏感，解析前会 lower）。 */
export const MEMORY_TIME_WORDS = Object.freeze(Object.keys(TIME_WORD_TABLE));

/** 相对时间的三种写法（解析用的正则源；测试拿它做纯判定）。 */
export const MEMORY_TIME_PATTERNS = Object.freeze([
    RELATIVE_DAYS.source,
    RELATIVE_MONTHS.source,
    RELATIVE_HOURS.source,
]);

/**
 * 报错提示里列的那几个（给人看的**中文**短清单）。
 * **每一个都必须在词表或相对写法里能解析** —— 测试会拿真解析器逐个验一遍。
 * 同义写法（`today` / `这周` / `上个月` / `this week`）不重复列：词表里都认，提示只给中文常用那几个。
 */
export const MEMORY_TIME_HINTS = Object.freeze([
    '今天', '昨天', '前天', '明天', '本周', '上周', '本月', '上月', '今年', '去年',
    '最近7天', '最近3个月',
]);

/** 找不到解析办法时的那句话（提示里的字与能解析的字同一份真源）。 */
function unknownTimeMessage(text) {
    return `看不懂的时间「${text}」。可用：2026-09-27、2026-09-27T10:30、`
        + `${MEMORY_TIME_HINTS.join(' / ')}（timezone：本地时区）。`;
}

/** 条目/记录身上可取的时间戳：归档里 mnemon 迁来的条目用的是 `storedAt`（审查挖出的第 4 条缺陷）。 */
function stampOf(item) {
    return asText(item?.at) || asText(item?.storedAt) || asText(item?.updatedAt) || '';
}

/**
 * 解析时间窗的一端。
 * @param {string} raw 用户给的原话（ISO 日期、带时区的 ISO 时刻，或时间词）
 * @param {{now?: Date, edge?: 'since'|'until'}} options edge='since' 取区间起点，'until' 取终点
 * @returns {number|null} 毫秒时间戳；空串返回 null
 */
function parseTimeBound(raw, { now = new Date(), edge = 'since' } = {}) {
    const text = asText(raw).trim();
    if (text === '') return null;
    const plain = text.toLowerCase().replace(/\s+/g, ' ');
    const pick = (start, end) => (edge === 'since' ? start.getTime() : end.getTime());

    if (ABSOLUTE_ISO.test(text)) {
        const at = Date.parse(text);
        if (Number.isNaN(at)) throw new Error(`看不懂的时间「${text}」。`);
        return at;
    }
    const local = LOCAL_DATE.exec(text);
    if (local !== null) {
        const [, year, month, day, hour, minute, second] = local;
        const hourValue = Number(hour ?? 0);
        const minuteValue = Number(minute ?? 0);
        const secondValue = Number(second ?? 0);
        // 时分秒越界要报错，不能靠 Date 静默进位（`T25:00` 会被算成次日 01:00，
        // 那是「读成另一个时刻」而不是「看不懂」）。
        if (hourValue > 23 || minuteValue > 59 || secondValue > 59) throw new Error(`看不懂的时间「${text}」（时分秒越界）。`);
        const value = new Date(Number(year), Number(month) - 1, Number(day), hourValue, minuteValue, secondValue, 0);
        if (Number.isNaN(value.getTime()) || value.getMonth() !== Number(month) - 1) throw new Error(`看不懂的时间「${text}」（日期不存在）。`);
        if (hour === undefined) return pick(startOfDay(value), endOfDay(value));
        return value.getTime();
    }
    const words = TIME_WORD_TABLE;
    // `Object.hasOwn` 而不是 `words[plain] !== undefined`：对象字面量继承 `Object.prototype`，
    // 于是 `since:'constructor'` / `'__proto__'` 会取到原型上的东西，报出来的是
    // 「Spread syntax requires ...iterable」这种看不懂的 TypeError（复核 F15，老毛病）。
    if (Object.hasOwn(words, plain)) return pick(...words[plain](now));
    const days = RELATIVE_DAYS.exec(plain);
    if (days !== null) {
        const count = Math.max(1, Number(days[1]));
        return pick(startOfDay(addDays(now, -(count - 1))), endOfDay(now));
    }
    const months = RELATIVE_MONTHS.exec(plain);
    if (months !== null) {
        const count = Math.max(1, Number(months[1]));
        return pick(addMonths(now, -(count - 1)), endOfDay(now));
    }
    const hours = RELATIVE_HOURS.exec(plain);
    if (hours !== null) {
        const count = Math.max(1, Number(hours[1]));
        const at = now.getTime() - count * 60 * 60 * 1000;
        return edge === 'since' ? at : now.getTime();
    }
    throw new Error(unknownTimeMessage(text));
}

/**
 * 把 `since` / `until` 解析成一个时间窗。
 * @returns {{since: number|null, until: number|null, sinceInput: string, untilInput: string, sinceText: string, untilText: string}|null}
 */
export function resolveTimeWindow({ since = '', until = '', now = new Date() } = {}) {
    const sinceInput = asText(since);
    const untilInput = asText(until);
    if (sinceInput === '' && untilInput === '') return null;
    const from = parseTimeBound(sinceInput, { now, edge: 'since' });
    const to = parseTimeBound(untilInput, { now, edge: 'until' });
    if (from !== null && to !== null && from > to) {
        throw new Error(`时间窗是空的：since（${sinceInput}）晚于 until（${untilInput}）。`);
    }
    return {
        since: from,
        until: to,
        sinceInput,
        untilInput,
        sinceText: from === null ? '' : new Date(from).toISOString(),
        untilText: to === null ? '' : new Date(to).toISOString(),
    };
}

/** 时间窗过滤：`at` 落在闭区间内。没有时间戳的条目在带时间窗时**一律排除**（拿不出证据就别给结论）。 */
function withinWindow(window, at) {
    if (window === null) return true;
    const value = Date.parse(asText(at));
    if (Number.isNaN(value)) return false;
    if (window.since !== null && value < window.since) return false;
    if (window.until !== null && value > window.until) return false;
    return true;
}

/**
 * 把一层拆成「窗内 / 窗外 / 没有时间戳」三份。
 *
 * 必须分开报（审查挖出的第 4 条缺陷）：把「没有时间戳」混进「在窗外」，回执会印
 * 「N 条都在时间窗之外：把 since/until 放宽」—— 而放宽**永远救不回**没有时间戳的条目
 * （mnemon 迁来的历史长期记忆就是这种形状），这正是本轮立项要消灭的假「没有」。
 */
function partitionByWindow(window, items, atOf) {
    if (window === null) return { inWindow: items, outside: 0, unstamped: 0 };
    const inWindow = [];
    let outside = 0;
    let unstamped = 0;
    for (const item of items) {
        const at = atOf(item);
        if (at === '') {
            unstamped += 1;
            continue;
        }
        if (withinWindow(window, at)) inWindow.push(item);
        else outside += 1;
    }
    return { inWindow, outside, unstamped };
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
 * `limit` 的收口：**先取整、再夹到 [1, hard]**。
 *
 * 为什么不能只写 `Math.min(limit, hard)`：`0.5` 这种小数会一路传到 `slice(0, 0.5)`
 * 或 `admitted.length >= resultLimit`，结果是**0 条** —— 而调用方以为自己要了一条。
 * 在检索面上这就表现为**假的弃答**（第四十五轮独立复核打穿 P3：`limit:0.5` 时
 * `matched=1`、`hits=0`，回执却写「没有一块的词表与查询重叠」）。
 * 非数、0、负数仍按「没传」处理（用 `fallback`），与既有语义一致。
 */
function limitOf(value, hard, fallback) {
    const n = typeof value === 'number' ? value : Number.parseInt(String(value), 10);
    if (!Number.isFinite(n) || n <= 0) return fallback;
    return Math.min(Math.max(1, Math.trunc(n)), hard);
}

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
        // 单条上限与建议线（新-11）、投影预算（P0-2）：memory.js 自己也有一份收敛 ——
        // createMemory({ memory }) 是测试与探针的直接入口，只加在 config.js 上会漏掉它们。
        entryLimitBytes: Number.isFinite(raw.entryLimitBytes) ? raw.entryLimitBytes : DEFAULT_ENTRY_LIMIT_BYTES,
        entryHintBytes: Number.isFinite(raw.entryHintBytes) ? raw.entryHintBytes : DEFAULT_ENTRY_HINT_BYTES,
        projectionBudgetBytes: Number.isFinite(raw.projectionBudgetBytes) ? raw.projectionBudgetBytes : DEFAULT_PROJECTION_BUDGET_BYTES,
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
 * 合并多个根的内容指纹（`contentRevision`）。
 *
 * 单根时**必须原样返回那一份**：它是「内容算出来的稳定指纹」这条契约的一部分
 * （同一份条目任何时候算出来都一样），外面拿它比对「有没有变」。多根时才退化成
 * 一个组合指纹 —— 这时它只用来判断「整体有没有变」。
 *
 * 名字里带 `Content` 是刻意的（第二十四轮 P2-5 / 第二十六轮 24-17）：记忆里有两个
 * 都叫过 `revision` 的指纹，算法与用途完全不同 ——
 *   contentRevision  由**条目内容**算（热记忆本体变没变）
 *   writeRevision    由 `id:updatedAt` 算（**有没有写操作**，台账也吃它）
 * 同名时早晚有人把「写了一次」当成「内容变了」用。
 */
function combineRevisions(items) {
    if (items.length === 0) return sha1('');
    if (items.length === 1) return items[0].contentRevision;
    return sha1(items.map((item) => `${item.origin}:${item.contentRevision}`).join('|'));
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
            perStore.push({ origin: store.id, contentRevision: loaded.contentRevision, updatedAt: loaded.updatedAt });
        }
        return {
            contentRevision: combineRevisions(perStore),
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
    async function read({ layer = 'all', query = '', limit = 0, since = '', until = '' } = {}) {
        const wanted = layer === 'hot' || layer === 'ledger' || layer === 'archive' ? layer : 'all';
        const terms = queryTerms(query);
        const quality = config.recallQuality;
        // 时间窗在**打分之前**生效：先缩小搜索空间，再排名。解析失败会在这里抛错
        // （静默忽略过滤器会把「没命中」说成「那段时间没有」）。
        const window = resolveTimeWindow({ since, until });
        const result = { layer: wanted, query: asText(query), truncated: false, origins: stores.map((store) => store.id), quality: quality.policy };
        if (window !== null) {
            result.window = {
                sinceInput: window.sinceInput,
                untilInput: window.untilInput,
                since: window.sinceText,
                until: window.untilText,
            };
        }
        // layer:'all' 时三层**各有各的字符额度**，不再抢同一份 totalChars。
        // 旧写法（热记忆先采纳、后两层拿剩余）会让热记忆的 11.8 KB 正文把 4,000 字吃干，
        // 台账与归档恒定为 0 条；详见 READ_LIMITS.allChars 的注释与第二十四轮实测。
        // 只读单层时这一层独占 totalChars，好让一次精确查询能拿满。
        const shared = wanted === 'all';
        const budgetFor = (layer) => (shared ? READ_LIMITS.allChars[layer] : READ_LIMITS.totalChars);
        const cap = (want, hard) => limitOf(limit, hard, hard);

        if (layers.hot && (wanted === 'hot' || wanted === 'all')) {
            const everything = [];
            const storeUsage = [];
            for (const store of stores) {
                const loaded = await loadStore(store.paths);
                for (const entry of loaded.entries) everything.push({ ...entry, origin: store.id });
                storeUsage.push({ origin: store.id, contentRevision: loaded.contentRevision });
            }
            // 时间维度过滤热记忆用 `updatedAt`（改动时刻），台账与归档用 `at`
            // （归档里 mnemon 迁来的条目只有 `storedAt`，走 stampOf 的回退）。
            const part = partitionByWindow(window, everything, (entry) => stampOf(entry));
            const all = part.inWindow;
            // 排序：**重要度永远是主键**，同重要度内按 recency 弱先验（新的在前，见
            // `recencyPriorOf` 的三条口径）。第三档是时刻字符串、第四档是 id —— 后两档
            // 只为「同一毫秒写多条」时给一个确定的次序（稳定排序的副产品不该是契约）。
            const nowMs = Date.now();
            const admitted = admit(
                all
                    .slice()
                    .sort((left, right) => (IMPORTANCE_RANK[left.importance] - IMPORTANCE_RANK[right.importance])
                        || (recencyPriorOf(right.updatedAt, nowMs) - recencyPriorOf(left.updatedAt, nowMs))
                        || String(right.updatedAt).localeCompare(String(left.updatedAt))
                        || String(left.id).localeCompare(String(right.id)))
                    .map((entry) => ({ id: entry.id, target: entry.target, importance: entry.importance, updatedAt: entry.updatedAt, origin: entry.origin, entities: entry.entities, kbRefs: entry.kbRefs, sources: entry.sources, text: entry.content, digest: sha1(entry.content) })),
                { resultLimit: cap('hot', READ_LIMITS.hotResults), itemChars: READ_LIMITS.itemChars, totalChars: budgetFor('hot') },
            );
            result.hot = {
                items: admitted.items,
                total: all.length,
                windowFiltered: part.outside,
                windowUnstamped: part.unstamped,
                budgetChars: budgetFor('hot'),
                contentRevision: combineRevisions(storeUsage),
                targets: {
                    user: { ...targetUsage(everything, 'user'), limit: config.userLimitBytes },
                    project: { ...targetUsage(everything, 'project'), limit: config.projectLimitBytes },
                },
            };
            result.truncated = result.truncated || admitted.truncated;
        }

        if (layers.ledger && (wanted === 'ledger' || wanted === 'all')) {
            const records = [];
            for (const store of stores) {
                for (const record of await loadLedger(store.paths)) records.push({ ...record, origin: store.id });
            }
            const part = partitionByWindow(window, records, (record) => stampOf(record));
            const inWindow = part.inWindow;
            const scored = inWindow
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
                    // 用过哪块知识来源（usedRefs）：渲染台账行时要印出来。
                    kbRefs: asArray(item.record.kbRefs),
                    text: [asText(item.record.path), asText(item.record.purpose), asArray(item.record.outline).slice(0, 4).join(' / ')].filter(Boolean).join(' —— '),
                })),
                { resultLimit: cap('ledger', READ_LIMITS.ledgerResults), itemChars: READ_LIMITS.itemChars, totalChars: budgetFor('ledger') },
            );
            result.ledger = {
                items: admitted.items,
                total: inWindow.length,
                windowFiltered: part.outside,
                windowUnstamped: part.unstamped,
                matched: scored.length,
                budgetChars: budgetFor('ledger'),
                qualityDropped: picked.dropped,
                relativeDir: relPath(base, paths.dir),
            };
            result.truncated = result.truncated || admitted.truncated || scored.length > admitted.items.length;
        }

        if (layers.archive && (wanted === 'archive' || wanted === 'all')) {
            const items = [];
            let digestCount = 0;
            let allItems = 0;
            let damagedVolumes = 0;
            for (const store of stores) {
                // 正文按卷从 `.jsonl` 读（索引只有指针）；`damaged` 是「索引说这卷有条目、
                // 正文文件却不在」—— 如实计数，不把它当成空卷。
                const harvested = await harvestArchive(store.paths, store.id);
                digestCount += harvested.volumes;
                damagedVolumes += harvested.damaged;
                for (const item of harvested.items) {
                    allItems += 1;
                    items.push(item);
                }
            }
            // 归档条目的时间戳要回退到 `storedAt`：mnemon 迁来的历史长期记忆只有它。
            const part = partitionByWindow(window, items, (item) => stampOf(item));
            const inWindow = part.inWindow;
            const scored = inWindow
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
                    // 为什么它在归档里（新-10）：热记忆下沉 / 台账滚动 / mnemon 迁移 / Pack 导入 /
                    // 外部写入。读侧与 sunk 用同一个字段，不各判各的。
                    reason: entry.item.reason,
                    importance: entry.item.importance,
                    category: entry.item.category,
                    entities: entry.item.entities,
                    tags: entry.item.tags,
                    path: entry.item.path,
                    score: entry.score,
                    text: archiveItemText(entry.item),
                })),
                { resultLimit: cap('archive', READ_LIMITS.archiveResults), itemChars: READ_LIMITS.itemChars, totalChars: budgetFor('archive') },
            );
            result.archive = {
                items: admitted.items,
                digests: digestCount,
                items_total: inWindow.length,
                items_all: allItems,
                // 「索引有、正文没有」的卷数：正常恒为 0（写入顺序是 jsonl 先落、索引后落），
                // 非 0 只可能是有人在外面删了 store 里的文件 —— 报出来，别让 read 静默给空。
                damagedVolumes,
                windowFiltered: part.outside,
                windowUnstamped: part.unstamped,
                matched: scored.length,
                budgetChars: budgetFor('archive'),
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
            tombstones: relPath(base, paths.tombstones),
        };
        result.globalDir = stores.some((store) => store.id === 'global') ? relPath(base, globalMemoryPaths(config).dir) : '';
        return result;
    }

    /**
     * 改热记忆：add / replace / remove。
     *
     * remove 只在用户明确要求或确有证据时用 —— 这条写在工具描述里，不在这里拦
     * （拦不住「模型认为没用了」，但可以要求它给出目标而不是下标）。
     *
     * 定位支持两种方式（第十九轮 P1-3 / 第二十四轮 P1-4）：`id` 精确寻址，或
     * `oldText` 子串唯一命中。id 是 read 的行首那个（`m-xxxxxxxx`），有它就
     * 不必再猜原文；oldText 那条老路一字未改，向后兼容。
     */
    async function mutate({ action, target = 'project', content = '', oldText = '', id = '', importance = 'normal', entities = [], tags = [], tier = '', source = null } = {}) {
        const act = asText(action);
        if (!['add', 'replace', 'remove'].includes(act)) throw new Error(`记忆动作只支持 add / replace / remove，收到「${act}」。`);
        // 层开关是「不再读 / 不再写」，不是「只读」：关掉热记忆之后还能 add，
        // 就会让人以为记忆坏了（写了读不到）。已有条目不受影响，打开就回来。
        if (!layers.hot) throw new Error('热记忆层已在设置里关掉（memory.layers.hot），本次没有写入。要写入请先把它打开；已有条目不会被删除。');
        const realTarget = asText(target);
        if (!MEMORY_TARGETS.includes(realTarget)) throw new Error(`记忆目标只支持 user / project，收到「${realTarget}」。`);
        const wanted = asText(content);
        const needle = asText(oldText);
        const wantedId = asText(id);
        if (act !== 'remove' && wanted === '') throw new Error('记忆内容不能为空。');
        if (act !== 'add' && needle === '' && wantedId === '') {
            throw new Error(`${act} 需要 oldText 或 id 指出改哪一条（oldText 给一段能唯一命中的原文；id 从 read 的行首拿）。`);
        }
        if (act !== 'remove' && byteLength(wanted) > (realTarget === 'user' ? config.userLimitBytes : config.projectLimitBytes)) {
            throw new Error(`这条记忆 ${byteLength(wanted)} 字节，比 ${realTarget} 侧的容量上限还大，存不下。请压缩到一两句。`);
        }
        // 单条硬上限（第四十八轮 新-11）：比「整层上限」严一档的写入期质量门。
        // 热记忆是每次动笔前照办的清单，不是文档仓库 —— 一条 10 KB 的段落会把整层吃光，
        // 又没人会读完它。长内容的去处是 source 指的输入文件 / kb / 交付物，不是这里。
        const entryLimit = finiteNumber(config.entryLimitBytes) ?? DEFAULT_ENTRY_LIMIT_BYTES;
        if (act !== 'remove' && entryLimit > 0 && byteLength(wanted) > entryLimit) {
            throw new Error(`这条记忆 ${byteLength(wanted)} 字节，超过单条上限 ${entryLimit} 字节，没有写入。`
                + '把它压成一两句（谁、什么、怎么办）；细节放进 source 指的输入文件，或先 kb-ingest 入库再引用块 id。'
                + '确实需要长条目就调大设置页「记忆 → 单条上限」（memory.entryLimitBytes，0 = 不限）。');
        }
        const realImportance = MEMORY_IMPORTANCE.includes(importance) ? importance : 'normal';
        // tier 是**显式输入**：拼错一个字母就报错，不静默降档（strictTierOf）。
        const declaredTier = strictTierOf(tier);
        // `source` 的三态语义：不传（null）= 不动已有引用；传数组（含空数组）= 按这次给的
        // 重建引用。区分「没传」与「传了空」是必要的 —— 否则一条引用了单源块的旧条目
        // 永远清不掉引用（改一个字仍被门拦住，独立验证 F1④ 的修法）。
        const sourceGiven = Array.isArray(source);
        const givenSource = sourceGiven ? stringList(source) : [];
        const givenRefs = givenSource.filter((item) => item.startsWith('kb:'));
        const givenFiles = givenSource.filter((item) => !item.startsWith('kb:'));
        // 写入期硬规则质量门（第二十四轮 §5.3 / P1-8）：**单源未核实的内容永不进热记忆**。
        // 三条路一起拦：显式声明 `tier:'unverified'`、引用了一块来源档为 unverified 的知识块、
        // 或（replace 时）条目**已经带着**这样的引用（独立验证 F1④：旧写法只看这次传了什么，
        // 带 unverified 引用的历史条目改一个字照样留在热记忆里）。
        // 为什么必须是硬规则而不是权重：加性 provenance 权重与完全无防御**不可区分**（p = 0.80），
        // 而按来源类别排除不可信内容把准确率从 0.3167 拉到 0.7000（Utility Under Attack）。
        // 读时过滤/降权被刻意排除：它要付 4.4 个准确率点、并误隔离 33.6% 的合法记忆。
        if (act !== 'remove' && (declaredTier !== '' || givenRefs.length > 0)) {
            if (declaredTier === 'unverified') {
                throw new Error('这条内容的来源档是 unverified（单源未核实），按硬规则**不能进热记忆**。'
                    + '先让它留在 kb 的待核区（kb-list / kb-read 能看到它），等有两个以上独立来源确认之后再记。');
            }
            await assertRefsTrusted(givenRefs);
        }
        const store = storeFor(realTarget);

        // **锁全部根**，不只是被请求的那个：跨根命中时真正被写的是另一个根的条目与边，
        // 只锁一侧等于没锁（审查实测：跨根并发会造出悬空引用、静默丢写与 rename EPERM）。
        // 多目录获取必须排序，见 withQueues 的注释。
        return withQueues(stores.map((item) => item.paths.dir), async () => {
            const loaded = await loadStore(store.paths);
            const entries = loaded.entries.slice();
            // id 与 oldText 两条路都要求唯一命中，也都支持**跨根**搜：多根时用户记不清
            // 那条偏好当初落在全局还是工作区，按当前根找不到就跨根找一遍。
            const hitsOf = (list) => (wantedId !== ''
                ? list.filter((entry) => entry.id === wantedId)
                : list.filter((entry) => entry.content.includes(needle)));
            let owner = store;
            // ownerEntries 必须与 matched **是同一批对象**：跨根命中时若回头再 loadStore 一次，
            // 拿到的是重新 JSON.parse 出来的新对象 —— remove 的 `indexOf(hit)` 会得 -1
            // （splice(-1,1) 删掉最后一条！），replace 只会改到临时对象上、saveStore 写回时
            // 内容没变（静默假回执）。这是本轮之前就存在的缺陷，随「按 id 跨根改删」一起修掉。
            let ownerEntries = entries;
            let matched = act === 'add' ? [] : hitsOf(entries);
            if (act !== 'add' && matched.length === 0) {
                for (const other of stores) {
                    if (other.id === store.id) continue;
                    const candidate = await loadStore(other.paths);
                    const hits = hitsOf(candidate.entries);
                    if (hits.length > 0) {
                        owner = other;
                        ownerEntries = candidate.entries.slice();
                        matched = hits;
                        break;
                    }
                }
            }
            let touched = null;
            let previous = null;

            if (act === 'add') {
                touched = normalizeEntry({ target: realTarget, content: wanted, importance: realImportance, entities, tags, kbRefs: givenRefs, sources: givenFiles, createdAt: nowIso(), updatedAt: nowIso() });
                entries.push(touched);
            } else {
                if (matched.length === 0) {
                    throw new Error(wantedId !== ''
                        ? `热记忆里没有 id 为「${wantedId}」的条目。id 是 read 行首那个（m-…）；台账与归档条目不能改（归档只读）。`
                        : `没有哪条记忆包含「${truncate(needle, 60)}」。先用 action:'read' 看现在的条目。`);
                }
                if (matched.length > 1) {
                    // 改和删都要求唯一命中：命中多条时按顺序取第一条，可能改错或删错的是
                    // 用户真正在意的那条，而记忆里的删除是不可见的破坏。
                    throw new Error(`「${truncate(needle, 60)}」命中 ${matched.length} 条，不唯一。oldText 给长一点，或直接用 read 行首的 id。`);
                }
                const hit = matched[0];
                previous = { ...hit };
                if (act === 'remove') {
                    ownerEntries.splice(ownerEntries.indexOf(hit), 1);
                    touched = hit;
                } else {
                    // replace 也要过门：这次没给 source 时，条目**已有的**引用照样要审一遍
                    // （带着 unverified 引用的旧条目改一个字不能就这么留在热记忆里）。
                    if (!sourceGiven) await assertRefsTrusted(hit.kbRefs);
                    hit.content = wanted;
                    hit.target = realTarget;
                    hit.importance = realImportance;
                    if (stringList(entities).length > 0) hit.entities = stringList(entities);
                    if (stringList(tags).length > 0) hit.tags = stringList(tags);
                    if (sourceGiven) {
                        hit.kbRefs = givenRefs;
                        hit.sources = givenFiles;
                    }
                    hit.updatedAt = nowIso();
                    touched = hit;
                }
            }

            // 删除必须**同事务**清边并留墓碑（第二十四轮 P1-9 / 第二十六轮 24-7）。
            // 悬空边比悬空条目更糟：条目查不到只是没有，边还在则 `related` 会静默少一条、
            // 甚至穿过已经删掉的节点继续往外走。跨根也要清 —— 边记在起点那一侧的根里，
            // 而**所有根此刻都被 withQueues 持有**，所以这里的读-改-写是排他的。
            //
            // 顺序有意写成「先清边 + 墓碑，再落 store / 索引」：这两步失败会抛出去，
            // 此时条目**还没被删**，现场是一致的（可以重试）；反过来先删条目再清边，
            // 中途失败就会留下「条目没了、边还在」的悬空状态。
            let prunedLinks = 0;
            let tombStoned = false;
            if (act === 'remove') {
                const dead = new Set([touched.id]);
                for (const item of stores) prunedLinks += await pruneLinks(item.paths, dead);
                tombStoned = (await appendTombstones(owner.paths, [{ id: touched.id, reason: 'remove' }])) > 0;
            }
            const maintenance = await enforceCapacity(owner.paths, ownerEntries, config);
            const saved = await saveStore(owner.paths, ownerEntries);
            const targets = await writeProjections(owner.paths, ownerEntries, config);
            // 刚写入的那条自己也可能是被下沉的那个（它比现有条目更次要）。
            // 这件事必须如实回执，不能让模型以为「记下了」。remove 不算「被下沉」：
            // 删掉本来就是这次调用的目的。
            const evicted = act !== 'remove' && !ownerEntries.some((entry) => entry.id === touched.id);
            // 建议线（新-11）：超过只在回执里提醒，不拒绝 —— 拒绝一条真实记忆的代价比
            // 多几行字大。replace 把条目改**短**了就不提醒（否则改短时还要被念一遍）。
            const hint = finiteNumber(config.entryHintBytes) ?? DEFAULT_ENTRY_HINT_BYTES;
            const advisoryBytes = act === 'remove' ? 0 : byteLength(wanted);
            const notShorter = previous === null || advisoryBytes >= byteLength(previous.content ?? '');
            const advisory = hint > 0 && advisoryBytes > hint && notShorter
                ? { bytes: advisoryBytes, hint }
                : null;
            return {
                action: act,
                target: realTarget,
                entry: touched,
                previous,
                evicted,
                advisory,
                prunedLinks,
                tombStoned,
                contentRevision: saved.contentRevision,
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
            .map((record) => {
                // 来源里混着两类东西：文件来源（放进 `source`）与**知识块 id**
                // （`kb:<docHash>:<n>`，放进 `kbRefs`）。**先分流再各自封顶** ——
                // 先 slice 再分流会让前面 12 条文件来源把后面真正的 kb 引用挤掉
                // （独立验证 F11）。
                const given = asArray(record.source).map((item) => String(item));
                const refs = given.filter((item) => item.startsWith('kb:')).slice(0, 12);
                const files = given.filter((item) => !item.startsWith('kb:')).slice(0, 8);
                return {
                    id: `L-${randomUUID().slice(0, 8)}`,
                    at: asText(record.at) || nowIso(),
                    path: asText(record.path),
                    format: asText(record.format),
                    theme: asText(record.theme),
                    bytes: Number.isFinite(record.bytes) ? record.bytes : 0,
                    purpose: asText(record.purpose),
                    outline: asArray(record.outline).slice(0, 12).map((item) => String(item)),
                    stats: record.stats !== null && typeof record.stats === 'object' ? record.stats : null,
                    source: files,
                    kbRefs: refs,
                    // note 也要封顶：它是**检索键的一部分**（ledgerSearchText 每次查询都把它拼进
                    // 匹配串），而告警是逐项 push 的（一个 500 处公式错误的 xlsx 能带出几十 KB）。
                    // 上限与单条展示口径（READ_LIMITS.itemChars）同量级，够留下前几条告警。
                    note: truncate(asText(record.note), LEDGER_NOTE_CHARS),
                };
            })
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
        // 台账近况的指纹：最近几条（取各调用方渲染上限里最大的 3 条）的 id/时刻/路径。
        // 它与热记忆的 writeRevision 互不相干 —— 台账只在登记新产物时变，热记忆只在
        // 增删改条目时变。投影（projection.js）用它决定「台账部分要不要重贴」。
        const ledgerWindow = records.slice(-3);
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
            ledgerRevision: sha1(ledgerWindow.map((record) => `${asText(record.id) || asText(record.path)}:${asText(record.at)}`).join('|')),
            // 两个指纹**故意不同名**（第二十四轮 P2-5 / 第二十六轮 24-17）：
            //   writeRevision   由 `id:updatedAt` 算 —— 「有没有写操作」（改一个字也算写）
            //   contentRevision 由**条目内容**算（memory.json 里那个）—— 「内容变没变」
            // 同名过一次，投影因此差点把「写了一次」当成「内容变了」重复贴全文。
            writeRevision: sha1(entries.map((entry) => `${entry.id}:${entry.updatedAt}`).sort().join('|')),
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
                file: digest === undefined ? null : relPath(base, join(store.paths.archive, basename(asText(digest.file)))),
            };
        });
    }

    /** 归档里已有的全部条目（迁移用它做幂等判断，也用于状态面板与实体索引）。 */
    async function archiveItems() {
        const out = [];
        for (const store of stores) {
            const harvested = await harvestArchive(store.paths, store.id);
            for (const item of harvested.items) out.push(item);
        }
        return out;
    }

    /**
     * 「最近下沉了什么」（第四十八轮 新-10）。
     *
     * 为什么要有它：容量维持的回执只报条数（「下沉了 2 条」）—— 热层忘了**哪几条**
     * 不可见，而这几条往往是 critical 的旧结论。查证只能靠人拿关键词去 archive 里撞，
     * 撞不到就以为内容丢了（本机真实库实测：一次写入下沉掉两条第四十四 / 四十五轮的
     * 落地记录，回执里连 id 都没有）。
     *
     * 只读、有界：按 `reason === 'hot-overflow'` 过滤，按时间倒序给最近 `limit` 块。
     */
    async function sunk({ limit = 5 } = {}) {
        const wanted = Math.max(1, Math.min(20, Math.round(finiteNumber(limit) ?? 5)));
        const all = (await archiveItems()).filter((item) => item.reason === 'hot-overflow');
        const blocks = new Map();
        for (const item of all) {
            const key = `${item.origin}|${item.blockId}`;
            const found = blocks.get(key) ?? { at: asText(item.blockAt) || asText(item.at), origin: item.origin, items: [] };
            found.items.push(item);
            blocks.set(key, found);
        }
        const ordered = [...blocks.values()]
            .sort((left, right) => String(right.at).localeCompare(String(left.at)))
            .slice(0, wanted);
        return {
            blocks: ordered.map((block) => ({
                at: block.at,
                origin: block.origin,
                items: block.items.map((item) => ({
                    id: item.id,
                    importance: item.importance,
                    target: item.origin,
                    preview: truncate(String(item.content ?? '').replace(/\s*\n\s*/g, ' '), 60),
                })),
            })),
            totalSinks: blocks.size,
            totalItems: all.length,
            limit: wanted,
        };
    }

    // ── 知识库（kb） ────────────────────────────────────────────────────────
    //
    // 五件事：显式入库、词法检索（24-11 接线）、清单、按需有界整份读、显式删除。
    // 检索是**词法**的（bigram + BM25，正文 + 块头两路融合），打分器是 `kb.js` 的纯函数
    // `kbRankChunks` —— 探针与 action 面量的是同一份实现。
    // kb 不进 `read` 的三层额度：那是「带问题查资料」的机械，读整块有自己的额度
    // （KB_READ_CHARS），检索命中只给预览（KB_HIT_PREVIEW）与次数配额（kbSearchPerTurn）。

    /** 工作目录内的相对路径（绝对路径与越界都拒绝）。 */
    function relativeInside(rawPath) {
        const given = asText(rawPath);
        if (given === '') throw new Error('需要 path（相对会话工作目录的文档路径）。');
        const absolute = isAbsolute(given) ? given : join(base, given);
        const rel = relative(base, absolute);
        if (rel === '' || rel.startsWith('..') || isAbsolute(rel)) {
            throw new Error(`path 要在会话工作目录里面：「${given}」指向目录之外，没有入库。`);
        }
        return rel.split(sep).join(posix.sep);
    }

    /** 跨根的全部 manifest 行（带 origin）。 */
    async function kbRows() {
        const rows = [];
        for (const store of stores) {
            for (const row of await loadKbManifest(store.paths.kb)) rows.push({ ...row, origin: store.id });
        }
        return rows;
    }

    /**
     * 按 id / path 找一份文档时的**查找顺序**：本会话写东西的那个根排第一。
     *
     * 为什么需要它：文档 id 是**按相对路径**寻址的，而相对路径在多个根里可能一样
     * （全局根与工作区根都有 `notes.md`）—— 两个根于是算出同一个文档 id。这时候
     * 按哪个根解释这个 id 就成了必须显式决定的事：以「这个会话往哪儿写」为准
     * （scope=global 写全局、其余写工作区），找不到再去别的根兜底。块 id 不受影响
     * （它按内容寻址，同一份内容本来就是同一批块）。
     */
    function kbLookupStores() {
        const primary = defaultStore();
        return [primary, ...stores.filter((store) => store.id !== primary.id)];
    }

    /**
     * 全部 kb id → 来源档。给写入期质量门用（`add` 引用到的块是单源未核实的就要拦）。
     * 只读 manifest：块 id 是内容寻址的，不必读正文。
     *
     * **同一份内容在多个根里档位不同时取更保守的那个**（独立验证 F4）：全局根标
     * `unverified`、工作区根标 `user` 时，旧写法按遍历顺序被后一个覆盖，门就看不见单源
     * 那一侧了。安全规则的方向永远是「更保守」，所以这里合并而不是覆盖。
     */
    async function kbTierIndex() {
        const map = new Map();
        const put = (id, tier) => {
            if (asText(id) === '') return;
            map.set(id, map.has(id) ? strictestTier(map.get(id), tier) : kbTierOf(tier));
        };
        for (const store of stores) {
            for (const row of await loadKbManifest(store.paths.kb)) {
                put(row.id, row.tier);
                for (let n = 1; n <= row.chunks; n += 1) put(chunkIdOf(row.hash, n), row.tier);
            }
        }
        return map;
    }

    /**
     * 一批知识块引用必须**找得到、且来源档不是 unverified**，否则抛错。
     *
     * 写入期质量门的三条路共用它（`add` 这次给的、`replace` 这次给的、`replace` 条目
     * 已有的）—— 门只在「这次传了什么」上判，就等于允许一条带着单源引用的旧条目
     * 一直留在热记忆里，改一个字也不重审（独立验证 F1④ 的修法）。
     */
    async function assertRefsTrusted(refs) {
        const wanted = stringList(refs).filter((item) => item.startsWith('kb:'));
        if (wanted.length === 0) return;
        const tiers = await kbTierIndex();
        const unknown = wanted.filter((ref) => !tiers.has(ref));
        if (unknown.length > 0) {
            throw new Error(`引用的知识块 id 在 kb 里找不到：${unknown.join('、')}。`
                + 'id 用 action:\'kb-list\' 或 kb-read 的回执拿；找不到的引用会让「依据」变成一句无法核对的话。');
        }
        const untrusted = wanted.filter((ref) => tiers.get(ref) === 'unverified');
        if (untrusted.length > 0) {
            throw new Error(`引用的知识块是 unverified（单源未核实）：${untrusted.join('、')}。`
                + '单源内容按硬规则不能进热记忆（投毒率 < 0.1% 就能把攻击成功率推到 80% 以上）。'
                + '要么换一个 ≥2 独立来源确认过的块，要么改这一条时传一组可信的 source（传空数组 = 清掉引用）。');
        }
    }

    /** 一个根的 kb 计数：索引里的话直接读（`status` 的定位是「不读全文」），否则按 manifest 现算。 */    async function kbStatusOf(store) {
        const index = await loadKbIndex(store.paths.kb);
        if (index !== null) return { ...index.counts, dir: relPath(base, store.paths.kb.dir) };
        const counts = kbCountsOf(await loadKbManifest(store.paths.kb));
        return { ...counts, dir: relPath(base, store.paths.kb.dir) };
    }

    /**
     * 换掉 / 删掉一份文档时，算出**真的死掉**的 id。
     *
     * 块 id 是内容寻址的，所以两份路径不同、内容相同的文档会**共用同一批块 id 与同一个
     * 块目录**。这时候删掉其中一份，另一份的块还活着 —— 按「这份文档的块全死」去清边、
     * 留墓碑、删目录，会把另一份文档的块一起删掉（manifest 说有几块、正文却没了），
     * 或者把另一份文档正在用的边静默清掉。所以死活的判据是「**还活着的 id 集合**里
     * 有没有它」，而这个集合必须**跨根**算：scope 为 both 时同一个内容哈希可能同时
     * 落在全局根与工作区根，只按被删那个根算，仍会把另一个根的块判死（第二十八轮那条
     * 「只锁一侧等于没锁」是同一类错误）。
     */
    function deadKbIds(removed, stillLive) {
        const dead = new Set([removed.id]);
        for (let n = 1; n <= removed.chunks; n += 1) dead.add(chunkIdOf(removed.hash, n));
        for (const id of [...dead]) {
            if (stillLive.has(id)) dead.delete(id);
        }
        return dead;
    }

    /**
     * 跨根的全部活着的 kb id。
     *
     * `rows` 是**被操作那个根**在换掉/删掉这份文档之后的清单（已经不含它）；别的根按
     * 盘上现状读。这里**不能**再按文档 id 过滤一遍：文档 id 按路径寻址，另一个根里
     * 同名的 `notes.md` 恰好就是同一个 id ——「同 id 就跳过」会把那个根正在用的块
     * 一起判死，于是边被清掉（跨根共用内容这条用例就是这么挖出来的）。
     */
    async function kbLiveAcrossStores({ store: excludeStore, rows: replacedRows }) {
        const rows = [];
        for (const store of stores) {
            const source = store.id === excludeStore.id ? replacedRows : await loadKbManifest(store.paths.kb);
            for (const row of source) rows.push(row);
        }
        return kbLiveIdsOf(rows);
    }

    /**
     * 显式入库一份工作目录里的文档。
     *
     * 三条口径：
     *   - **原文不复制**：kb 只存块正文与 span，源文件仍在自己的位置上（manifest
     *     记的是相对路径）。所以源文件被删之后，kb 里剩下的仍是当时切出来的正文，
     *     但它不再是「那一份文件的当前版本」——`kb-read` 会如实带上入库时刻。
     *   - **同路径同内容 = 幂等**：内容哈希没变就直接回`unchanged`，一个字节不重写。
     *   - **同路径新内容 = 换代**：旧块文件删掉、指向旧块的边同事务清掉、旧块 id 留墓碑；
     *     文档 id（按路径寻址）不变，块 id（按内容寻址）换成新的。
     */
    async function kbIngest({ path = '', tier = '' } = {}) {
        const relPosix = relativeInside(path);
        const absolute = join(base, relPosix);
        let info = null;
        try {
            info = await stat(absolute);
        } catch (error) {
            throw new Error(`读不了「${relPosix}」（${error?.code ?? '未知错误'}）。路径要相对会话工作目录。`);
        }
        if (!info.isFile()) throw new Error(`「${relPosix}」不是文件，kb 一期只认工作目录里的文本文件。`);
        if (info.size > KB_MAX_DOC_BYTES) {
            throw new Error(`「${relPosix}」有 ${info.size} 字节，超过 kb 的单文档上限 ${KB_MAX_DOC_BYTES} 字节。`
                + '先把它切小或只放需要的那一节 —— kb 有意不复制大文件进工作目录。');
        }
        let text = '';
        try {
            text = await readFile(absolute, 'utf8');
        } catch (error) {
            throw new Error(`读不了「${relPosix}」的正文（${error?.code ?? error?.message ?? '未知错误'}）。`);
        }
        if (text.includes('\u0000')) {
            throw new Error(`「${relPosix}」看起来是二进制（正文里有 NUL）。kb 一期只收文本：'
                + 'PDF 先用 office.pdf.text 抽成 .md，表格先导出成 .md，再入库。`);
        }
        // 来源档：显式传了就要看得懂（拼错一个字母不该静默降档，见 strictTierOf）。
        const realTier = strictTierOf(tier) || kbTierOf('');
        // 文档哈希 = 内容 + **切块档**（kbHash 的注释说明了为什么档位必须进哈希）。
        const docHash = kbHash(text, chunkCharsFor(relPosix));
        const docId = docIdOf(relPosix);
        const store = defaultStore();

        return withQueue(store.paths.dir, async () => {
            const rows = await loadKbManifest(store.paths.kb);
            const existing = rows.find((row) => row.id === docId);
            if (existing !== undefined && existing.hash === docHash) {
                // **同内容但显式改了来源档** = 改档，不是重写：两份独立来源确认之后，一份
                // 单源材料该能从「待核」升成「已验证」—— 否则硬规则就成了死路（内容一个字
                // 没变，却因为没有升级通道只能一直关在待核区，或者被迫改一个字来触发换代）。
                // 改的是 manifest 行的 `tier`：**块文件里没有档位**（档是文档级的），所以
                // 共用同一批块的另一份文档不会被连带改档（独立验证 F13 的 (a)）。
                const wantedTier = strictTierOf(tier);
                if (wantedTier !== '' && wantedTier !== existing.tier) {
                    const next = rows.map((row) => (row.id === docId ? { ...row, tier: wantedTier } : row));
                    await saveKbManifest(store.paths.kb, next);
                    await saveKbIndex(store.paths.kb, next);
                    return {
                        action: 'kb-ingest', docId, path: relPosix, title: existing.title,
                        chunks: existing.chunks, tier: wantedTier, previousTier: existing.tier,
                        unchanged: true, retiered: true, replaced: false, origin: store.id,
                    };
                }
                return {
                    action: 'kb-ingest', docId, path: relPosix, title: existing.title,
                    chunks: existing.chunks, tier: existing.tier, unchanged: true, replaced: false,
                    origin: store.id,
                };
            }
            // 先切块（纯函数，确定性），切不出正文就明确报错，不写一份空文档进 kb。
            const at = nowIso();
            const pieces = splitDocument(text, { chunkChars: chunkCharsFor(relPosix) });
            if (pieces.length === 0) {
                throw new Error(`「${relPosix}」里没有可入库的正文（全是空白行或标题）。`);
            }
            const title = kbTitleOf(text, relPosix);
            // 块文件只放与块本身有关的东西（正文 / span / 标题路径 / 自校验哈希）：来源路径、
            // 入库时刻、来源档都是文档级的，由 manifest 行现取 —— 两份内容相同的文档共用同一
            // 批块文件，把文档级元数据写进去就会互相覆盖（独立验证 F3）。
            const chunks = pieces.map((piece) => normalizeKbChunk({
                id: chunkIdOf(docHash, piece.n),
                n: piece.n,
                text: piece.text,
                span: piece.span,
                outline: piece.outline,
            }));
            // 顺序有意写成「正文先落 → 清旧件与边 → manifest → index」：与归档同一条纪律
            // （真源先落、索引后落），崩在任何一步都不会留下「索引说有、正文没有」的卷。
            for (const chunk of chunks) {
                await writeAtomic(join(store.paths.kb.chunks, chunkRelOf(docHash, chunk.n)), `${JSON.stringify(chunk)}\n`);
            }
            let prunedLinks = 0;
            let replaced = false;
            const next = rows.filter((item) => item.id !== docId);
            if (existing !== undefined) {
                replaced = true;
                // 墓碑只留**真的死掉**的**块** id：文档 id 按路径寻址，换代之后同一个 id 还活着
                // （新内容），给它留墓碑既与「alive wins」的判死口径多绕一圈，又会让
                // `pruneLinks` 把指向这份文档的边**误清**（独立验证 F5 实测：注释写着「不留」，
                // 代码却在留）。所以这里把文档 id 从死名单里摘掉，再按跨根的活 id 集合过滤块 id。
                const dead = deadKbIds(existing, await kbLiveAcrossStores({ store, rows: next }));
                dead.delete(existing.id);
                for (const item of stores) prunedLinks += await pruneLinks(item.paths, dead);
                if (dead.size > 0) {
                    await appendTombstones(store.paths, [...dead].map((deadId) => ({ id: deadId, reason: 'kb-reingest' })));
                }
                if (!next.some((item) => item.hash === existing.hash)) {
                    await rm(join(store.paths.kb.chunks, existing.hash), { recursive: true, force: true });
                }
            }
            const row = normalizeManifestRow({
                id: docId,
                path: relPosix,
                hash: docHash,
                title,
                bytes: byteLength(text),
                chars: text.length,
                chunks: chunks.length,
                tier: realTier,
                at,
            });
            next.push(row);
            await saveKbManifest(store.paths.kb, next);
            const counts = await saveKbIndex(store.paths.kb, next);
            return {
                action: 'kb-ingest', docId, path: relPosix, title, chunks: chunks.length,
                chars: text.length, bytes: byteLength(text), tier: realTier,
                unchanged: false, replaced, prunedLinks, counts, origin: store.id,
            };
        });
    }

    /** kb 清单：按入库时刻倒序，有界。 */
    async function kbList({ limit = 0 } = {}) {
        const rows = await kbRows();
        rows.sort((left, right) => String(right.at).localeCompare(String(left.at)));
        const cap = limitOf(limit, KB_LIST_LIMIT, KB_LIST_LIMIT);
        return {
            action: 'kb-list',
            items: rows.slice(0, cap),
            total: rows.length,
            truncated: rows.length > cap,
            counts: kbCountsOf(rows),
        };
    }

    /**
     * 按 id 或 path 读块（有界整份读）。
     *
     * `id` 两种：块 id（`kb:<docHash>:<n>`）只读那一块；文档 id（`kb:<docHash>`）按块号
     * 顺序读到额度用完为止。额度是 KB_READ_CHARS，被截断时明确报「还有 N 块没给」。
     */
    async function kbRead({ id = '', path = '' } = {}) {
        const wantedId = asText(id);
        let docId = '';
        let wantedHash = '';
        let wantedChunk = 0;
        if (wantedId !== '') {
            const parts = wantedId.split(':');
            // 两种 id 的**寻址方式不同**，别混：文档 id 按路径寻址（`kb:<路径哈希>`），
            // 块 id 按内容寻址（`kb:<内容哈希>:<块号>`）。所以块 id 不能用文档 id 那把
            // 尺子去找 —— 要看 manifest 行的 `hash` 字段，而不是它的 `id`。
            if (parts[0] === 'kb' && parts.length >= 3) {
                wantedHash = parts[1];
                wantedChunk = Number.parseInt(parts[2], 10);
            } else {
                docId = wantedId;
            }
        } else {
            docId = docIdOf(relativeInside(path));
        }
        for (const store of kbLookupStores()) {
            const rows = await loadKbManifest(store.paths.kb);
            // 块 id 按**内容**寻址，同一批块可能被两个根里的两份文档共用；请求的块号在
            // 这一份里越界时**继续往后找**（多个候选里挑第一个真的有这一块的），
            // 而不是当场报「只有 N 块」——那不是「没有这一块」，只是这一份没有
            // （独立验证 4h）。文档 id 则按路径寻址，直接命中或继续找。
            const candidates = wantedHash !== ''
                ? rows.filter((item) => item.hash === wantedHash)
                : rows.filter((item) => item.id === docId);
            if (candidates.length === 0) continue;
            const row = wantedChunk > 0
                ? candidates.find((item) => wantedChunk <= item.chunks)
                : candidates[0];
            if (row === undefined) {
                throw new Error(`「${candidates[0].path}」只有 ${candidates[0].chunks} 块，没有第 ${wantedChunk} 块。先用 action:'kb-list' 看清单。`);
            }
            const all = await loadKbChunks(store.paths.kb, row.hash, row.chunks);
            const wanted = wantedChunk > 0 ? all.filter((chunk) => chunk.n === wantedChunk) : all;
            const cap = wantedChunk > 0 ? 1 : KB_READ_CHUNKS;
            const picked = [];
            let used = 0;
            for (const chunk of wanted) {
                if (picked.length >= cap) break;
                if (picked.length > 0 && used + chunk.text.length > KB_READ_CHARS) break;
                picked.push(chunk);
                used += chunk.text.length;
            }
            if (wantedChunk > 0 && picked.length === 0) {
                throw new Error(`「${row.path}」第 ${wantedChunk} 块的正文读不到（块文件缺失或坏了）。`);
            }
            return {
                action: 'kb-read',
                docId: row.id,
                path: row.path,
                title: row.title,
                tier: row.tier,
                at: row.at,
                // 块头 / 来源 / 档位都是**文档级**的，在这里现取现拼（块文件里没有这些字段：
                // 两份内容相同的文档共用同一批块，写进块文件就会互相覆盖）。
                chunks: picked.map((chunk) => ({
                    ...chunk,
                    path: row.path,
                    tier: row.tier,
                    at: row.at,
                    header: chunkHeaderOf({ outline: chunk.outline, source: row.path, at: row.at }),
                    spanText: spanText(chunk.span),
                })),
                chunks_total: row.chunks,
                chunks_missing: row.chunks - all.length,
                total: wanted.length,
                truncated: picked.length < wanted.length,
                budgetChars: KB_READ_CHARS,
                origin: store.id,
            };
        }
        throw new Error(`kb 里没有「${wantedId || docId}」这份文档。`
            + (wantedHash === '' ? '' : '（块 id 里的那一段是**内容**哈希，不是文档 id；先 kb-list 拿文档 id 或块 id。）')
            + '先用 action:\'kb-list\' 看有什么；id 从 kb-list 的每行拿，或直接给 path。');
    }

    /**
     * kb 检索（第二十四轮 24-11）：词法多路（正文 + 块头）BM25 融合，**弃答**而不是硬凑。
     *
     * 三条口径，都有出处：
     *   - **词法检索，不做嵌入**。候选打分器在第四十二轮的自测集上量过（15 条标注查询
     *     hit@3 15/15，朴素整串只有 5/15），本轮把它从探针文件搬进 `src/kb.js` 并接在这里。
     *     边界照实说：**语义改写召不回**（查询与正文用词完全对不上时词法检索给不出结果）。
     *   - **弃答**。词表零重叠的块一律不给（候选块的原始分必须有一路大于 0）——
     *     把最像的那块端出来当「命中」是 LongMemEval 的 ABS 口径里最典型的失败。
     *   - **有界**。一次最多读 `KB_SEARCH_CHUNKS` 块正文、最多给 `KB_SEARCH_LIMIT` 条命中；
     *     命中里给预览而不是整块（逐字引用要走 `kb-read`，那块有自己的额度）。
     *
     * 检索**不占读取额度**但占每回合配额（与记忆侧的 `read` 同一条纪律）：词法检索最容易被
     * 「没查到就换个词再查」拖着走，所以次数由 `quota.kbSearchPerTurn` 管。
     */
    async function kbSearch({ query = '', limit = 0, tier = '', path = '' } = {}) {
        const wanted = asText(query);
        if (wanted === '') {
            throw new Error('kb-search 需要 query（要查什么）。它是**词法**检索：中文按 bigram 展开、'
                + '拉丁词整词保留；查询与正文用词完全对不上（语义改写）时召不回 —— 这是这条边界的原文意思，不是故障。');
        }
        const wantedTier = strictTierOf(tier, 'tier');
        const wantedPath = asText(path) === '' ? '' : relativeInside(path);
        const candidates = [];
        const seen = new Set();
        /** 同一个内容哈希可能被多份文档共用（块按内容寻址），块文件只读一次。 */
        const loadedCache = new Map();
        let docsTotal = 0;
        let docsRead = 0;
        // 两个如实计数的量，第四十五轮独立复核挖出来的两处（P2/P3）都靠它们说清楚：
        //   capped        块数撞上单次上限而**中途停下**（可能与「文档读完」无关）
        //   missingChunks manifest 说有几块、盘上却读不到（或自校验对不上）的块数
        let capped = false;
        let missingChunks = 0;
        for (const store of kbLookupStores()) {
            for (const row of await loadKbManifest(store.paths.kb)) {
                if (wantedTier !== '' && row.tier !== wantedTier) continue;
                if (wantedPath !== '' && row.path !== wantedPath) continue;
                docsTotal += 1;
                if (candidates.length >= KB_SEARCH_CHUNKS) {
                    capped = true;
                    continue;
                }
                docsRead += 1;
                let loaded = loadedCache.get(row.hash);
                if (loaded === undefined) {
                    loaded = await loadKbChunks(store.paths.kb, row.hash, row.chunks);
                    loadedCache.set(row.hash, loaded);
                }
                missingChunks += Math.max(0, row.chunks - loaded.length);
                for (const chunk of loaded) {
                    if (candidates.length >= KB_SEARCH_CHUNKS) {
                        capped = true;
                        break;
                    }
                    // 同一批块被两个根里的两份文档共用：内容一样，只算一次（先命中的那个根赢）。
                    if (seen.has(chunk.id)) continue;
                    seen.add(chunk.id);
                    candidates.push({
                        id: chunk.id,
                        docId: row.id,
                        path: row.path,
                        title: row.title,
                        tier: row.tier,
                        at: row.at,
                        origin: store.id,
                        n: chunk.n,
                        span: chunk.span,
                        text: chunk.text,
                        head: chunkHeaderOf({ outline: chunk.outline, source: row.path, at: row.at }),
                    });
                }
            }
        }
        const cap = limitOf(limit, KB_LIST_LIMIT, KB_SEARCH_LIMIT);
        const ranked = kbRankChunks(candidates, wanted);
        const hits = ranked.slice(0, cap).map((row) => ({
            chunkId: row.chunk.id,
            docId: row.chunk.docId,
            path: row.chunk.path,
            title: row.chunk.title,
            tier: row.chunk.tier,
            origin: row.chunk.origin,
            n: row.chunk.n,
            score: row.score,
            scoreText: row.scoreText,
            scoreHead: row.scoreHead,
            header: row.chunk.head,
            spanText: spanText(row.chunk.span),
            chars: row.chunk.text.length,
            preview: truncate(row.chunk.text, KB_HIT_PREVIEW),
        }));
        // 没能给出命中时，**原因必须分得清**（复核 P3：path 筛掉全部文档时也写「弃答 /
        // 没有一块的词表与查询重叠」，把读者推向「换个更贴近原文的词」——而真实原因是筛选）。
        const reason = hits.length > 0
            ? 'hits'
            : docsTotal === 0
                ? 'filter-empty'
                : candidates.length === 0
                    ? 'no-chunks'
                    : 'no-overlap';
        // `bounded` = **这次没把符合筛选的块扫完**。两种成因都要算进去（复核 P2）：
        // ① 上限撞到而中途停下（`capped`）—— 早先只写 `docsRead < docsTotal`，于是
        //    上限在**最后一份**文档里撞到时它算成 false，回执不说「没扫完」，
        //    最坏还报出一句**假的弃答**；
        // ② 有文档被整份跳过。
        const bounded = capped || docsRead < docsTotal;
        return {
            action: 'kb-search',
            query: wanted,
            hits,
            matched: ranked.length,
            truncated: ranked.length > hits.length,
            reason,
            docs: docsTotal,
            docsRead,
            bounded,
            chunks: candidates.length,
            chunksMissing: missingChunks,
            budgetChars: KB_HIT_PREVIEW,
            maxChunks: KB_SEARCH_CHUNKS,
        };
    }

    /**
     * 显式删除一份入库文档。
     *
     * 与 `remove` 同一条纪律：**先清边 + 留墓碑，再动正文与索引**。悬空边比悬空条目更糟，
     * 而 kb 的块是 link 的目标（「这条结论依据哪块来源」），所以删文档时必须把落在它的
     * 文档 id 与全部块 id 上的边一并清掉。墓碑留着：一份旧 Pack 不该把删掉的来源搬回来。
     */
    async function kbDrop({ id = '', path = '' } = {}) {
        const wantedId = asText(id);
        const relPosix = asText(path) === '' ? '' : relativeInside(path);
        // 给的可以是文档 id、块 id（`kb:<内容哈希>:<块号>`，删掉它所属的那份文档）
        // 或 path。三种都指到“一份文档”这个单位上 —— kb 的删除粒度就是文档。
        const parts = wantedId.split(':');
        const byHash = parts[0] === 'kb' && parts.length >= 3 ? parts[1] : '';
        const targetId = wantedId !== '' ? wantedId : docIdOf(relPosix);
        for (const store of kbLookupStores()) {
            const rows = await loadKbManifest(store.paths.kb);
            const hit = byHash !== '' ? rows.find((row) => row.hash === byHash) : rows.find((row) => row.id === targetId);
            if (hit === undefined) continue;
            return withQueue(store.paths.dir, async () => {
                const current = await loadKbManifest(store.paths.kb);
                const again = current.find((row) => row.id === hit.id);
                if (again === undefined) throw new Error(`「${hit.id}」在这一刻已经不在 kb 里了。`);
                const next = current.filter((row) => row.id !== again.id);
                const dead = deadKbIds(again, await kbLiveAcrossStores({ store, rows: next }));
                let prunedLinks = 0;
                for (const item of stores) prunedLinks += await pruneLinks(item.paths, dead);
                if (dead.size > 0) {
                    await appendTombstones(store.paths, [...dead].map((deadId) => ({ id: deadId, reason: 'kb-drop' })));
                }
                if (!next.some((row) => row.hash === again.hash)) {
                    await rm(join(store.paths.kb.chunks, again.hash), { recursive: true, force: true });
                }
                await saveKbManifest(store.paths.kb, next);
                await saveKbIndex(store.paths.kb, next);
                return {
                    action: 'kb-drop', docId: again.id, path: again.path, title: again.title,
                    chunks: again.chunks, prunedLinks, tombstones: dead.size, origin: store.id,
                };
            });
        }
        throw new Error(`kb 里没有「${wantedId || relPosix}」这份文档。先用 action:'kb-list' 看清单。`);
    }

    // ── 图关系 ──────────────────────────────────────────────────────────────

    /**
     * 建一条双向类型化关系。
     *
     * 两个 id 都必须**真的存在**（热记忆、台账或归档里能找到）：允许指向不存在的
     * id 会让关系图慢慢烂成一堆悬空边，而悬空边在「沿关系找相邻条目」时表现为
     * 静默少一条 —— 那比直接报错难查得多。
     */
    async function link({ sourceId = '', targetId = '', kind = '', note = '', conflict = '', state = '' } = {}) {
        if (!config.links) throw new Error('图关系已在设置里关掉（memory.links）。');
        const from = asText(sourceId);
        const to = asText(targetId);
        if (from === '' || to === '') throw new Error('link 需要 sourceId 与 targetId。');
        if (from === to) throw new Error('不能把一条记忆连到它自己。');
        const realKind = LINK_KINDS.includes(kind) ? kind : DEFAULT_LINK_KIND;
        // 冲突的两个字段**显式传了就要能看懂**：认不出来的值一律报错，而不是悄悄退回
        // 「未分类 / 未决」—— 那会把上一步记下的判断抹成「没分类」，是静默的信息丢失。
        // （读入落盘数据时仍是宽松的：老文件、外部改过的 links.jsonl 不该读不动。）
        const wantedConflict = asText(conflict);
        const wantedState = asText(state);
        if (wantedConflict !== '' && !CONFLICT_CLASSES.includes(wantedConflict)) {
            throw new Error(`conflict 只支持 ${CONFLICT_CLASSES.join(' / ')}，收到「${wantedConflict}」。`);
        }
        if (wantedState !== '' && !CONFLICT_STATES.includes(wantedState)) {
            throw new Error(`state 只支持 ${CONFLICT_STATES.join(' / ')}，收到「${wantedState}」。`);
        }

        const index = await entryIndex();
        // 墓碑只在「这个 id 现在确实不活着」时才判死：同一个 id 可能被重建
        // （`pointer:archive` 就是固定 id —— 删掉之后下次容量溢出会重新建它），
        // 陈旧墓碑不该把一条活着的条目永久锁死、更不能让它的合法边被删掉。
        const stones = new Set((await allTombstones()).map((stone) => stone.id));
        const alive = await liveIds();
        for (const id of [from, to]) {
            if (stones.has(id) && !alive.has(id)) throw new Error(`id 为「${id}」的条目已经被删除了（墓碑还在），不能给它建关系。`);
        }
        for (const id of [from, to]) {
            if (!index.has(id)) throw new Error(`找不到 id 为「${id}」的记忆条目。先用 action:'read' 或 action:'related' 拿到确切 id。`);
        }
        const fromStore = stores.find((store) => store.id === index.get(from).origin) ?? defaultStore();
        return withQueue(fromStore.paths.dir, async () => {
            const existing = await loadLinks(fromStore.paths);
            const record = normalizeLink({ sourceId: from, targetId: to, kind: realKind, conflictClass: conflict, conflictState: state, note, at: nowIso() });
            const duplicate = existing.find((item) => item.sourceId === from && item.targetId === to && item.kind === realKind);
            if (duplicate !== undefined) {
                // 冲突的三态是**可以转的**（未决 → 偏向某一侧）：同一条边再调一次带新的
                // 状态时按「改状态」处理，而不是被幂等判成「已存在」悄悄丢掉这次判断。
                const changed = realKind === 'contradicts'
                    && (asText(conflict) !== '' || asText(state) !== '')
                    && (duplicate.conflictClass !== record.conflictClass || duplicate.conflictState !== record.conflictState);
                if (!changed) return { added: false, updated: false, link: duplicate, origin: fromStore.id };
                duplicate.conflictClass = asText(conflict) === '' ? duplicate.conflictClass : record.conflictClass;
                duplicate.conflictState = asText(state) === '' ? duplicate.conflictState : record.conflictState;
                await writeAtomic(fromStore.paths.links, existing.map((item) => `${JSON.stringify(item)}\n`).join(''));
                return { added: false, updated: true, link: duplicate, origin: fromStore.id };
            }
            await appendLinks(fromStore.paths, [record]);
            return { added: true, updated: false, link: record, origin: fromStore.id };
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
        // kb 先放（第四层）：条目 id 是内容寻址的，只读 manifest 就能把文档与块全部摊出来，
        // 所以 `link` 可以指向「哪一块来源」而不必把正文读进内存。同 id 不可能与三层撞
        // （kb 前缀是 `kb:`，热记忆是 `m-`、台账是 `L-`、归档是 `mnemon:` / `b-`）。
        for (const store of stores) {
            for (const row of await loadKbManifest(store.paths.kb)) {
                index.set(row.id, {
                    id: row.id,
                    layer: 'kb',
                    origin: store.id,
                    text: `[kb 文档] ${row.title}（${row.path}，${row.chunks} 块）`,
                    item: row,
                });
                for (let n = 1; n <= row.chunks; n += 1) {
                    const id = chunkIdOf(row.hash, n);
                    index.set(id, {
                        id,
                        layer: 'kb',
                        origin: store.id,
                        text: `[kb 第 ${n} 块] ${row.path}`,
                        item: { ...row, n },
                    });
                }
            }
        }
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

    /** 全部墓碑（跨根合并）。 */
    async function allTombstones() {
        const out = [];
        for (const store of stores) {
            for (const stone of await loadTombstones(store.paths)) out.push({ ...stone, origin: store.id });
        }
        return out;
    }

    /**
     * 全部**活着**的条目 id（热记忆 + 台账 + 归档）。
     *
     * 与 `entryIndex()` 的差别是**不看层开关**：层关掉只是「这次不读」，不是「条目不存在」，
     * 而悬空引用这条不变量要的是事实 —— 拿可见范围去算，会把关掉一层的库报成满是悬空边。
     */
    async function liveIds() {
        const ids = new Set();
        for (const store of stores) {
            for (const entry of (await loadStore(store.paths)).entries) {
                const id = asText(entry.id);
                if (id !== '') ids.add(id);
            }
            for (const record of await loadLedger(store.paths)) {
                const id = asText(record.id);
                if (id !== '') ids.add(id);
            }
            const index = await loadArchiveIndex(store.paths);
            for (const digest of index.digests) {
                for (const block of asArray(digest.blocks)) {
                    for (const id of blockIds(block)) ids.add(id);
                }
            }
            // kb 的文档 id 与块 id 也是关系边可能指向的目标（「这条结论依据哪块来源」）。
            // 只读 manifest：块 id 由「内容哈希 + 块号」算出来，不为这条不变量读正文。
            for (const id of kbLiveIdsOf(await loadKbManifest(store.paths.kb))) ids.add(id);
        }
        return ids;
    }

    /** 悬空引用（不变量：必须为空）。数得出来才守得住。 */
    async function danglingRefs() {
        const live = await liveIds();
        const out = [];
        for (const store of stores) {
            for (const link of await loadLinks(store.paths)) {
                if (!live.has(link.sourceId) || !live.has(link.targetId)) out.push({ ...link, origin: store.id });
            }
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
        if (!index.has(start)) {
            // 只有「真的不活着」才说是删除：层被关掉时 index 里也没有它，但那不是删除。
            const gone = (await liveIds()).has(start) ? undefined : (await allTombstones()).find((stone) => stone.id === start);
            if (gone !== undefined) throw new Error(`id 为「${start}」的条目已经被删除（${gone.reason || 'remove'} / ${String(gone.at).slice(0, 10)}），没有可遍历的节点了。`);
            throw new Error(`找不到 id 为「${start}」的记忆条目。`);
        }
        const wantedKind = LINK_KINDS.includes(kind) ? kind : '';
        const links = (await allLinks()).filter((link) => wantedKind === '' || link.kind === wantedKind);
        const hops = Math.min(3, Math.max(1, boundedInt(depth, 1, 1, 3)));
        const cap = limit > 0 ? Math.min(limit, READ_LIMITS.relatedResults) : READ_LIMITS.relatedResults;

        const seen = new Set([start]);
        const nodes = [];
        const edges = [];
        const dangling = [];
        let frontier = [start];
        for (let step = 0; step < hops; step += 1) {
            const next = [];
            for (const current of frontier) {
                for (const link of links) {
                    const other = link.sourceId === current ? link.targetId : link.targetId === current ? link.sourceId : '';
                    if (other === '' || seen.has(other)) continue;
                    seen.add(other);
                    const entry = index.get(other);
                    if (entry === undefined) {
                        // 悬空边（另一端已经不在了）。旧实现把它当普通邻居 push 进 next，
                        // 于是**删掉的节点变成一座桥**：`related` 会报出隔着一层删除内容
                        // 才够得到的条目。现在只如实记下来，不穿过它继续走。
                        // 同一个缺失对端只报一次（seen 已记），所以这里的数按「指不到的对端」算。
                        dangling.push({ id: link.id, kind: link.kind, from: current, to: other, note: link.note, origin: link.origin });
                        continue;
                    }
                    edges.push({
                        id: link.id,
                        kind: link.kind,
                        from: current,
                        to: other,
                        note: link.note,
                        origin: link.origin,
                        // 冲突边的三态要一起带出去：读到「未决」才知道两边都还没有定论。
                        conflictClass: link.conflictClass ?? '',
                        conflictState: link.conflictState ?? '',
                    });
                    if (nodes.length < cap) {
                        nodes.push({ ...entry, via: link.kind, hop: step + 1, conflictState: link.conflictState ?? '' });
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
            dangling,
            danglingCount: dangling.length,
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
            const archiveItemFiles = {};
            if (includeArchiveFiles) {
                // 归档正文的两个投影都要带走：`.md` 是给人读的，`.jsonl` 是结构化真源 ——
                // 索引现在只有指针，少带 `.jsonl` 的话导入方拿回的只有「有几卷几块」。
                for (const digest of index.digests) {
                    const stem = archiveStem(asText(digest.file));
                    const text = await readTextOrNull(join(store.paths.archive, basename(asText(digest.file))));
                    if (text !== null) archiveFiles[stem] = text;
                    const items = await readTextOrNull(join(store.paths.archive, basename(asText(digest.itemFile))));
                    if (items !== null) archiveItemFiles[stem] = items;
                }
            }
            // 知识库（第四棵树）：manifest 是清单，块正文按 `<docHash>/<n>.json` 逐块带走。
            // 与归档同一个理由 —— 少了正文，导入方拿回的只是「有几份文档、切了几块」。
            const kbManifest = await loadKbManifest(store.paths.kb);
            storesOut.push({
                id: store.id,
                dir: store.paths.dir,
                entries: loaded.entries,
                ledger: await loadLedger(store.paths),
                archiveIndex: index,
                archiveFiles,
                archiveItemFiles,
                kb: {
                    manifest: includeArchiveFiles ? kbManifest : [],
                    files: includeArchiveFiles ? await readKbChunkFiles(store.paths.kb, kbManifest) : {},
                },
                links: await loadLinks(store.paths),
                tombstones: await loadTombstones(store.paths),
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
        const report = { entries: 0, skippedEntries: 0, skippedTombstoned: 0, skippedDangling: 0, skippedUnverified: 0, ledger: 0, skippedLedger: 0, links: 0, skippedLinks: 0, tombstones: 0, skippedTombstones: 0, archiveBlocks: 0, skippedArchiveBlocks: 0, kbDocs: 0, skippedKbDocs: 0, skippedKbDamaged: 0, skippedKbTombstoned: 0, stores: [] };

        for (const incomingStore of incoming) {
            // 导入目标：默认把 Pack 里的根**还原到同名根**（global → 全局，其余 → 工作区）。
            const id = asText(incomingStore.id) === 'global' ? 'global' : 'workspace';
            if (wantedTarget !== '' && wantedTarget !== id) continue;
            const store = stores.find((item) => item.id === id);
            if (store === undefined) {
                report.stores.push({ id, skipped: true, reason: '当前记忆范围里没有这个根' });
                continue;
            }
            const storeReport = { id, entries: 0, ledger: 0, links: 0, archiveBlocks: 0, kbDocs: 0 };

            await withQueue(store.paths.dir, async () => {
                // 本地墓碑先读一次：删掉的条目**不能**靠导入复活（GateMem：检索式记忆会
                // 泄漏已删除的信息）。判死用的是 `allTombstones()`（**全部根**）而不是
                // 只有目标根 —— 墓碑是跨根生效的（`link` 也这么判），只看一个根会出现
                // 「全局根判死的 id 被 import 搬回来」（审查挖出的第 3 条缺陷的一半）。
                const dead = new Set((await allTombstones()).map((stone) => stone.id));
                // 导入热记忆之前先算一张「来源档」表（本地 + 入包声明，取更保守的一侧）：
                // 单源未核实的内容不许借着「恢复一份旧包」重新进热记忆 —— 本地 `add` 被拒的
                // 那条，打包再导入照样得被拒（第四十一轮独立验证 F1①）。
                const kbTiers = await kbTierIndex();
                for (const raw of asArray(incomingStore.kb?.manifest)) {
                    const row = normalizeManifestRow(raw);
                    if (row === null) continue;
                    const put = (id) => kbTiers.set(id, kbTiers.has(id) ? strictestTier(kbTiers.get(id), row.tier) : kbTierOf(row.tier));
                    put(row.id);
                    for (let n = 1; n <= row.chunks; n += 1) put(chunkIdOf(row.hash, n));
                }
                // 热记忆
                const loaded = await loadStore(store.paths);
                const entries = loaded.entries.slice();
                const known = new Set(entries.map((entry) => entry.id));
                for (const raw of asArray(incomingStore.entries)) {
                    const entry = normalizeEntry(raw);
                    if (entry === null) continue;
                    if (dead.has(entry.id)) {
                        report.skippedTombstoned += 1;
                        continue;
                    }
                    if (known.has(entry.id)) {
                        report.skippedEntries += 1;
                        continue;
                    }
                    // 引用单源块（或引用了表里查不到的块）的条目不入库：与 `add` 同一个门。
                    const refs = stringList(entry.kbRefs);
                    const untrusted = refs.filter((ref) => !['verified', 'user'].includes(kbTiers.get(ref)));
                    if (untrusted.length > 0) {
                        report.skippedUnverified += 1;
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

                // 墓碑：本地没有的才补，按 id 幂等（墓碑只增不删 —— 删除是事实，不该被覆盖）。
                // 要在**搬边之前**做：一份自带墓碑的 Pack 若先搬边，会把指向「被它自己
                // 判死的 id」的边装进库（审查挖出的第 3 条缺陷）。
                const freshStones = [];
                for (const raw of asArray(incomingStore.tombstones)) {
                    const stone = normalizeTombstone(raw);
                    if (stone === null) continue;
                    if (dead.has(stone.id)) {
                        report.skippedTombstones += 1;
                        continue;
                    }
                    dead.add(stone.id);
                    freshStones.push(stone);
                    report.tombstones += 1;
                }
                if (freshStones.length > 0) await appendTombstones(store.paths, freshStones);

                // 归档：按月合并，块去重
                const index = await loadArchiveIndex(store.paths);
                const blocks = [];
                // 新格式的 Pack：归档正文在 `archiveItemFiles` 里（`archiveIndex` 只有指针）；
                // 老格式的 Pack：正文内联在 `archiveIndex.digests[].blocks[].items` 里。
                // 两种都要能读回来，否则「升级一次就把老包读成空归档」。
                const incomingItemFiles = incomingStore.archiveItemFiles !== null && typeof incomingStore.archiveItemFiles === 'object'
                    ? incomingStore.archiveItemFiles
                    : {};
                for (const digest of asArray(incomingStore.archiveIndex?.digests)) {
                    const month = asText(digest?.month) || monthOf(nowIso());
                    let incomingBlocks = asArray(digest?.blocks);
                    if (!incomingBlocks.some((block) => asArray(block?.items).length > 0)) {
                        const stem = archiveStem(digest?.file) || archiveVolumeStem(month, finiteNumber(digest?.volume) ?? 1);
                        incomingBlocks = parseDigestItemText(incomingItemFiles[stem] ?? incomingItemFiles[month] ?? '');
                    }
                    // 去重键跨**同月的每一卷**收集：分卷之后一个月可能有好几卷，
                    // 只看一卷会让「同一块被另一卷带进来」重复落库。
                    const known = new Set();
                    for (const row of index.digests) {
                        if (row.month !== month) continue;
                        for (const block of asArray(row.blocks)) known.add(blockKey(block));
                    }
                    for (const item of blocks) {
                        if (item.month === month) known.add(blockKey(item));
                    }
                    for (const block of incomingBlocks) {
                        // 去重键与写入时**同一个算法**：有 id 用 id，没有就按**入包声明的
                        // 内容**算（不能等 appendArchive 按落盘后的内容算 —— 墓碑过滤会
                        // 改变结果，第二次导入就算不出同一个 id、同一条内容被搬两份；
                        // 第四十轮独立验证 probe-a A4 实测）。
                        const key = blockKey(block);
                        if (known.has(key)) {
                            report.skippedArchiveBlocks += 1;
                            continue;
                        }
                        // 归档块里的条目也要过墓碑：封顶删掉的最旧卷不该被一份旧 Pack 搬回来。
                        const items = asArray(block.items).filter((item) => {
                            const id = asText(item?.id);
                            return id === '' || !dead.has(id);
                        });
                        if (items.length === 0) {
                            report.skippedTombstoned += 1;
                            continue;
                        }
                        // **块的标识要原样带过去**：去重键就是它。
                        const fresh = {
                            id: key,
                            month,
                            at: asText(block.at) || nowIso(),
                            reason: asText(block.reason) || 'pack-import',
                            items,
                        };
                        blocks.push(fresh);
                        known.add(key);
                        report.archiveBlocks += 1;
                        storeReport.archiveBlocks += 1;
                    }
                }
                if (blocks.length > 0) await appendArchive(store.paths, blocks, config);

                // 知识库（第四棵树）：manifest 行按**文档 id** 幂等（只增不改），块正文按
                // `<docHash>/<n>.json` 逐个落盘。一条文档只有在「它的每一块都在包里、
                // 自校验对得上、且都没被墓碑判死」时才入库 —— 否则会出现 manifest 说有几块、
                // 正文却缺块（或被改过）的半份文档，那正是 kb 的 `damaged` 口径要报的东西，
                // 不该由导入制造出来。
                //
                // 两道**闸**（第四十一轮独立验证 F1/F2 挖出来的）：
                //   1. **文档 id 的墓碑也要判**：同内容两份只删掉一份时，块 id 因为另一份还活
                //      而没有墓碑，只有文档 id 留了 —— 旧写法只看块 id，于是一份旧 Pack 能把
                //      删掉的那份原样搬回来；
                //   2. **每块都要自校验**（`chunk.hash === kbHash(chunk.text)`）与块号/块 id 一致：
                //      包里的正文与 id 不匹配时照收，等于让「Pack 可以改内容而保留 id」。
                const incomingKb = incomingStore.kb !== null && typeof incomingStore.kb === 'object' ? incomingStore.kb : {};
                const incomingKbFiles = incomingKb.files !== null && typeof incomingKb.files === 'object' ? incomingKb.files : {};
                const localKb = await loadKbManifest(store.paths.kb);
                const localKbIds = new Set(localKb.map((row) => row.id));
                const aliveNow = kbLiveIdsOf(localKb);
                const freshKb = [];
                for (const raw of asArray(incomingKb.manifest)) {
                    const row = normalizeManifestRow(raw);
                    if (row === null) continue;
                    if (localKbIds.has(row.id)) {
                        report.skippedKbDocs += 1;
                        continue;
                    }
                    // 文档 id 的墓碑：删掉的那份不该被旧 Pack 复活（alive wins：本地已经又
                    // 有同 id 的活文档时上面那条 localKbIds 就拦下了）。
                    if (dead.has(row.id) && !aliveNow.has(row.id)) {
                        report.skippedKbTombstoned += 1;
                        continue;
                    }
                    if (row.chunks < 1) {
                        report.skippedKbDamaged += 1;
                        continue;
                    }
                    let blocked = false;
                    const files = {};
                    for (let n = 1; n <= row.chunks; n += 1) {
                        const id = chunkIdOf(row.hash, n);
                        if (dead.has(id) && !aliveNow.has(id)) {
                            blocked = true;
                            break;
                        }
                        const text = incomingKbFiles[chunkRelOf(row.hash, n)];
                        let parsed = null;
                        if (typeof text === 'string') {
                            try {
                                parsed = JSON.parse(text);
                            } catch {
                                parsed = null;
                            }
                        }
                        const chunk = normalizeKbChunk(parsed);
                        if (chunk === null || chunk.id !== id || chunk.n !== n || !chunkSelfConsistent(chunk)) {
                            blocked = true;
                            break;
                        }
                        files[chunkRelOf(row.hash, n)] = `${JSON.stringify(chunk)}\n`;
                    }
                    if (blocked) {
                        report.skippedKbDamaged += 1;
                        continue;
                    }
                    for (const [rel, text] of Object.entries(files)) {
                        await writeAtomic(join(store.paths.kb.chunks, rel), text);
                    }
                    // 档位取**更保守**的一方：包里说 verified、本地同一份内容说 unverified 时，
                    // 相信 unverified（档位可以被人手改包里的字段，本地记录不会）。
                    const localSameHash = localKb.find((item) => item.hash === row.hash);
                    freshKb.push(localSameHash === undefined ? row : { ...row, tier: strictestTier(row.tier, localSameHash.tier) });
                    localKbIds.add(row.id);
                    report.kbDocs += 1;
                    storeReport.kbDocs += 1;
                }
                if (freshKb.length > 0) {
                    await saveKbManifest(store.paths.kb, [...localKb, ...freshKb]);
                    await saveKbIndex(store.paths.kb, [...localKb, ...freshKb]);
                }

                // 关系**最后搬**（归档合并之后）：这样「活着的 id」集合已经包含本次
                // 导入进来的热记忆、台账与归档条目。逐条校验两端真的活着 —— 老 Pack
                // 里带悬空边是常态（本轮修的就是「以前 remove 不清边」），照搬会把脏边
                // 搬进一个干净的库，让「悬空引用 == 0」从此恒不成立，而且不报、不计数。
                const live = await liveIds();
                const links = await loadLinks(store.paths);
                const knownLinks = new Set(links.map((link) => `${link.sourceId}|${link.targetId}|${link.kind}`));
                const freshLinks = [];
                for (const raw of asArray(incomingStore.links)) {
                    const link = normalizeLink(raw);
                    if (link === null) continue;
                    if (dead.has(link.sourceId) || dead.has(link.targetId)) {
                        report.skippedTombstoned += 1;
                        continue;
                    }
                    if (!live.has(link.sourceId) || !live.has(link.targetId)) {
                        report.skippedDangling += 1;
                        continue;
                    }
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
                // 浏览面板要的是条目正文：与 read 同一条路（索引只有指针）。
                const harvested = await harvestArchive(store.paths, store.id);
                digestCount += harvested.volumes;
                for (const item of harvested.items) out.archive.push(item);
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
        // 悬空引用是跨根算的：边记在起点那一侧的根里，而目标可能在另一个根。
        const live = await liveIds();
        let dangling = 0;
        let tombstoneCount = 0;
        for (const store of stores) {
            const loaded = await loadStore(store.paths);
            const index = await loadArchiveIndex(store.paths);
            // 条目数从**指针块**算：索引里就写着每块的条数，不必读正文（`status` 的
            // 定位是「不读全文，只给计数与体积」）。
            const archived = index.digests.reduce((sum, digest) => sum + asArray(digest.blocks).reduce((inner, block) => inner + blockCount(block), 0), 0);
            const size = await dirStats(store.paths.dir);
            const storeLinks = await loadLinks(store.paths);
            const storeDangling = danglingOf(storeLinks, live).length;
            const stones = await loadTombstones(store.paths);
            dangling += storeDangling;
            tombstoneCount += stones.length;
            perStore.push({
                id: store.id,
                dir: relPath(base, store.paths.dir),
                hot: loaded.entries.length,
                ledger: (await loadLedger(store.paths)).length,
                archive: archived,
                // 归档的**摘要文件**数：面板的容量条要的是它（上限 archiveKeep 管的就是
                // 文件数，分卷之后一个月可能有几卷），而 archive 是条目数 —— 两个数都给，
                // 别让界面自己猜。
                archiveFiles: index.digests.length,
                // 受损的卷数（索引声明有条目、正文却给不出来，见 digestBodyMissing）：
                // 面板要能看出「这个根有读不出来的东西」，而不是按索引计数显示成「都读得到」。
                // 注意这是**廉价近似**：老格式不算受损、坏行不算受损（那要读正文才看得出）。
                archiveDamaged: index.digests.filter((digest) => digestBodyMissing(store.paths, digest)).length,
                links: storeLinks.length,
                // 不变量：dangling 必须恒为 0（第二十四轮 P1-9）。它一旦非 0，就是
                // 「删了条目忘了清边」——正是这条统计存在的理由。
                dangling: storeDangling,
                tombstones: stones.length,
                // 浏览面板要显示「这个根占多少地方」：字节与文件数按目录实算。
                // 目录不存在时是 0 而不是 undefined —— 契约是数字，面板直接渲染它。
                bytes: size.bytes,
                files: size.files,
                // 知识库（第四棵树）：文档数 / 块数 / 来源三档的分布。读索引就够了；
                // 索引不在时按 manifest 现算（真源是 manifest 与块文件，索引只是投影）。
                kb: await kbStatusOf(store),
            });
        }
        return { scope: config.scope, userScope: config.userScope, layers: { ...layers }, dangling, tombstones: tombstoneCount, stores: perStore };
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
        sunk,
        kbIngest,
        kbRead,
        kbList,
        kbSearch,
        kbDrop,
        kbRows,
        kbTierIndex,
        link,
        unlink,
        related,
        entities,
        entryIndex,
        liveIds,
        danglingRefs,
        allTombstones,
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
        // 指路条目自己也要占空间，所以先把它的位置预留出来再下沉。它同时是**投影里**
        // 唯一看得见「热层少了东西」的一行，所以带上最近一次的条数与清单入口（新-10）。
        // 但**必须短**：它是热记忆里的一条，多一个字节就少一个字节给真实条目 ——
        // 而且它占用的是「下沉时预留」的额度，写长了会把 critical 也挤下去（实测过）。
        const pointerTextOf = (count) => `较旧的${target === 'user' ? '用户偏好' : '项目记忆'}已下沉到归档`
            + `（${count} 条，清单见 action:'sunk'）；`
            + "检索：office_memory({ action: 'read', layer: 'archive', query: '...' })";
        const pointerText = pointerTextOf(99);
        const pointerBytes = byteLength(pointerText) + byteLength(ENTRY_DELIMITER);
        const pointerFits = pointerBytes > 0 && pointerBytes < limit;
        // 已经有指路条目时不必再为它预留空间（它已经在用量里了）。按
        // limit - pointerBytes 一路下沉会把「预留」变成「多留一份」，最后连
        // critical 也被吃掉 —— 实测过。
        const hasPointer = entries.some((entry) => entry.id === ARCHIVE_POINTER_ID && entry.target === target);
        const demote = (allowCritical) => {
            // 「更旧」用**插入次序**判断，不用时间戳：同一毫秒内连着写多条时
            // updatedAt 完全相同，按它会退化成按 id 随机挑一个下沉。
            const candidates = entries
                .map((entry, index) => ({ entry, index }))
                .filter(({ entry }) => entry.target === target && entry.id !== ARCHIVE_POINTER_ID
                    && (allowCritical || entry.importance !== 'critical'))
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
        while (usage.used > reservation) {
            // 只为了给**指路条目**腾位置时不动 critical（新-10）：指路条目是 low，
            // 职责是服务条目，不该反过来把最该留下的东西挤走。真的超出整层上限
            // （usage > limit）时才允许动 critical —— 那时它是唯一能腾出空间的东西。
            const allowCritical = usage.used > limit;
            if (!demote(allowCritical)) break;
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
                content: pointerTextOf(moved.length),
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
            // 回执要能点名（新-10）：id + 重要度 + 60 字预览，够判断「丢的是不是要紧的那条」
            movedPreview: moved.map((entry) => ({
                id: entry.id,
                importance: entry.importance,
                preview: truncate(String(entry.content ?? '').replace(/\s*\n\s*/g, ' '), 60),
            })),
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
        // id 就印在重要度后面（第十九轮 P1-3）：`link` / `unlink` / `related` / 按 id 改
        // 都要它，而旧渲染只印重要度与正文 —— 热记忆与台账的 id 没有任何合法途径拿到，
        // 图关系有一半入口是暗的（归档与实体视图是唯一入口）。
        for (const item of items) {
            const id = asText(item.id) === '' ? '' : ` ${item.id}`;
            // 依据（kbRefs）与非知识块来源（sources）也要印：这条结论是哪块来源、哪份材料
            // 支撑的，是「引用可回溯」在模型侧唯一的可见入口（第四十一轮）。各印三条，
            // 其余折叠成计数，免得把额度吃光。
            const refs = stringList(item.kbRefs);
            const files = stringList(item.sources);
            const basis = refs.length === 0
                ? ''
                : `｜依据 ${refs.slice(0, 3).join(' ')}${refs.length > 3 ? ` 等 ${refs.length} 块` : ''}`;
            const from = files.length === 0
                ? ''
                : `｜来源 ${files.slice(0, 3).join(' ')}${files.length > 3 ? ' …' : ''}`;
            lines.push(`- [${item.importance}]${id}${originTag(item.origin)} ${item.text.replace(/\s*\n\s*/g, ' ')}${basis}${from}`);
        }
        lines.push(`  ${usageLine(target, hot.targets[target])}`);
    }
    return lines;
}

/**
 * 一层的「为什么这一行没有条目」。
 *
 * 必须把三种情况分开，因为旧实现只有一句「（空）」：台账明明有 20 条、命中 20 条，却因为
 * 预算被上层吃干而一条没给，渲染出来是「【台账】共 20 条 / - （空）」—— 同一段话里既说
 * 20 条又说空，模型会据此判定「台账里没有东西」，然后重做一份已有的交付物。
 * 现在：真的没有（total 0）说「空」；有但没命中说「没命中」；命中却被额度卡住说「额度被占满」。
 */
function emptyLayerNote(layer, { query = '', windowed = false } = {}) {
    if (layer.total === 0 || layer.items_total === 0) {
        // 带时间窗时「一条都没有」有三种：本来就没有、有但都在窗外、有但没有时间戳。
        // 混着说会让模型把「窗选窄了」或「这条根本没有时间」读成「那段时间没记过」；
        // 更糟的是对「没有时间戳」建议「放宽窗口」—— 放宽永远无效。
        if (windowed) {
            const notes = [];
            if ((layer.windowUnstamped ?? 0) > 0) notes.push(`${layer.windowUnstamped} 条没有时间戳（带时间窗时排除，放宽窗口也读不到）`);
            if ((layer.windowFiltered ?? 0) > 0) notes.push(`${layer.windowFiltered} 条在时间窗之外（可以放宽 since / until）`);
            if (notes.length > 0) return `（${notes.join('；')}）`;
        }
        return '（空）';
    }
    if (query !== '' && layer.matched === 0) return '（没有命中的条目：换更准的关键词，或把召回质量 policy 改成 off）';
    const budget = layer.budgetChars === undefined ? '' : `本层 ${layer.budgetChars} 字额度`;
    return `（命中 ${layer.matched ?? layer.total} 条，但${budget}被更长的条目占满，一条都没给）`;
}

/** 一层给了几条、被截掉几条 —— 有截断就明说给的是切片，而不是装作给全了。 */
function sliceNote(layer, label) {
    if (layer.items.length === 0 || layer.matched === undefined) return '';
    if (layer.matched <= layer.items.length) return '';
    const budget = layer.budgetChars === undefined ? '' : `本层额度 ${layer.budgetChars} 字，`;
    return `（本层只给了 ${layer.items.length}/${layer.matched} 条：${budget}看全部用 layer:'${label}'）`;
}

/** 渲染 action:'read' 的结果。 */
export function renderRead(result) {
    const title = result.layer === 'all' ? '三层' : result.layer === 'hot' ? '热记忆' : result.layer === 'ledger' ? '台账' : '归档';
    const lines = [`📒 记忆读取（${title}）${result.query ? `：${result.query}` : ''}`];
    const windowed = result.window !== undefined;
    if (result.window !== undefined) {
        const label = `${result.window.sinceInput || '不限'} ~ ${result.window.untilInput || '不限'}`;
        // 印出来的时刻必须是**本地**时刻：窗口是按本地时区解析的，印 UTC 会让人对不上账。
        const resolved = [result.window.since, result.window.until].filter(Boolean).map(localStamp).join(' ~ ');
        lines.push(`时间窗：${label}${resolved ? `（本地时刻 ${resolved}）` : ''}`);
    }
    if (result.origins !== undefined && result.origins.length > 1) {
        lines.push(`范围：${result.origins.map((id) => (id === 'global' ? '全局' : '工作区')).join(' + ')}${result.globalDir ? `（全局层 ${result.globalDir}）` : ''}`);
    }
    if (result.hot !== undefined) {
        lines.push('', ...hotLines(result.hot));
        const hotSlice = result.hot.total > result.hot.items.length
            ? `（热记忆 ${result.hot.total} 条里给了 ${result.hot.items.length} 条：本层额度 ${result.hot.budgetChars} 字，看全部用 layer:'hot'）`
            : '';
        if (hotSlice !== '') lines.push(hotSlice);
        if (windowed && result.hot.total === 0) {
            const notes = [];
            if ((result.hot.windowUnstamped ?? 0) > 0) notes.push(`${result.hot.windowUnstamped} 条没有时间戳（排除，放宽窗口也读不到）`);
            if ((result.hot.windowFiltered ?? 0) > 0) notes.push(`${result.hot.windowFiltered} 条在时间窗之外（可放宽 since / until）`);
            if (notes.length > 0) lines.push(`（热记忆${notes.join('；')}）`);
        } else if (windowed && (result.hot.windowFiltered ?? 0) + (result.hot.windowUnstamped ?? 0) > 0) {
            // 层内部分过滤也要说：上面那行用量（targets / usageLine）是**全量**口径，
            // 与这里列出的条数不是一个数，不说清会被读成「记忆丢了几条」。
            lines.push(`（时间窗：另有 ${result.hot.windowFiltered ?? 0} 条在窗外、${result.hot.windowUnstamped ?? 0} 条没有时间戳未列出；上面的用量是全量口径）`);
        }
    }
    if (result.ledger !== undefined) {
        lines.push('', `【台账】共 ${result.ledger.total} 条${result.query ? `，命中 ${result.ledger.matched} 条` : ''}`);
        if (result.ledger.items.length === 0) lines.push(`- ${emptyLayerNote(result.ledger, { query: result.query, windowed })}`);
        for (const item of result.ledger.items) {
            const badge = [item.format, item.theme].filter(Boolean).join(' · ');
            const id = asText(item.id) === '' ? '' : ` ${item.id}`;
            // 台账行也带「这份交付物用过哪块来源」（第四十一轮 usedRefs）：反查的入口要在
            // 结果里看得见，只存在字段里没人知道。
            const refs = stringList(item.kbRefs);
            const basis = refs.length === 0 ? '' : `｜用过 ${refs.slice(0, 3).join(' ')}${refs.length > 3 ? ' …' : ''}`;
            lines.push(`- ${item.at.slice(0, 10)}${id}${originTag(item.origin)} ${item.text}${badge ? `  [${badge}]` : ''}${basis}`);
        }
        const note = sliceNote(result.ledger, 'ledger');
        if (note !== '') lines.push(note);
        if (windowed && (result.ledger.windowFiltered ?? 0) + (result.ledger.windowUnstamped ?? 0) > 0) {
            lines.push(`（时间窗：另有 ${result.ledger.windowFiltered ?? 0} 条在窗外、${result.ledger.windowUnstamped ?? 0} 条没有时间戳，未列出）`);
        }
    }
    if (result.archive !== undefined) {
        lines.push('', `【归档】${result.archive.digests} 个摘要 / ${result.archive.items_total} 条${result.query ? `，命中 ${result.archive.matched} 条` : ''}`);
        if (result.archive.items.length === 0) lines.push(`- ${emptyLayerNote(result.archive, { query: result.query, windowed })}`);
        for (const item of result.archive.items) {
            const label = item.kind === 'hot' ? `热记忆（${item.origin}/${item.importance}）`
                : item.kind === 'mnemon' ? `mnemon 长期记忆${item.category ? `（${item.category}）` : ''}`
                    : '台账';
            const id = asText(item.id) === '' ? '' : ` ${item.id}`;
            // 归档的原因（新-10）：只说「在归档里」，读的人不知道它是被容量挤下来的
            // 还是台账滚进来的 —— 前者可能要救，后者只是旧。
            const why = asText(item.reason) === '' || item.reason === 'external' ? '' : ` · ${archiveSection(item.reason)}`;
            lines.push(`- [${item.month}]${id} ${label}${why} ${item.text}`);
        }
        const note = sliceNote(result.archive, 'archive');
        if (note !== '') lines.push(note);
        if (windowed && (result.archive.windowFiltered ?? 0) + (result.archive.windowUnstamped ?? 0) > 0) {
            lines.push(`（时间窗：另有 ${result.archive.windowFiltered ?? 0} 条在窗外、${result.archive.windowUnstamped ?? 0} 条没有时间戳，未列出）`);
        }
        // 「索引列着这卷、正文文件却不在」要当场说清：正常恒为 0（写入顺序是 .jsonl
        // 先落、索引后落），非 0 只可能是有人在外面删了 store 里的文件 —— 那时候
        // 「归档里没有这条」和「归档读不出来」是两件事，不能都显示成空。
        if ((result.archive.damagedVolumes ?? 0) > 0) {
            lines.push(`（警告：有 ${result.archive.damagedVolumes} 卷归档的正文文件读不到，索引里还列着它 —— 这些条目这次给不出来，不是「归档里没有」。）`);
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
    // 回执也印 id：刚写完这一条就能直接拿去 link / 按 id 再改，不必再 read 一次拿键。
    const badge = (entry) => `[${entry?.importance ?? 'normal'}]${asText(entry?.id) === '' ? '' : ` ${entry.id}`}`;
    if (result.action === 'remove') {
        lines.push(`- 移除：${badge(result.previous)} ${truncate(result.previous?.content ?? '', 120)}`);
        // 删了就必须说清边和墓碑怎么了：只报「移除」会让人以为图里还留着这条路。
        lines.push(`- 同事务清掉关系边 ${result.prunedLinks ?? 0} 条；墓碑${result.tombStoned === true ? '已留' : '未写'}（删掉的 id 不会再被 link，也不会靠 Pack 导入复活）。`);
    } else if (result.action === 'replace') {
        lines.push(`- 改前：${badge(result.previous)} ${truncate(result.previous?.content ?? '', 120)}`);
        lines.push(`- 改后：${badge(result.entry)} ${result.entry?.content ?? ''}`);
    } else {
        lines.push(`- 新增：${badge(result.entry)} ${result.entry?.content ?? ''}`);
    }
    // 写了就把它引用的知识块印出来（第四十一轮）：引用是给下一次「这条依据什么」看的，
    // 只写进文件而不回执，模型会以为这次引用没生效。
    const refs = stringList(result.entry?.kbRefs);
    if (result.action !== 'remove' && refs.length > 0) {
        lines.push(`- 依据 ${refs.length} 块来源：${refs.slice(0, 4).join(' ')}${refs.length > 4 ? ' …' : ''}（related 可以顺着这些 id 走）`);
    }
    lines.push(`- 用量：${usageLine('user', result.usage.user)}；${usageLine('project', result.usage.project)}`);
    if (result.evicted === true) {
        lines.push('- 注意：容量维持把刚写的这条也下沉了（它比现有条目更次要）。要留住它就先精简现有条目，或调大上限。');
    }
    for (const item of result.maintenance ?? []) {
        lines.push(`- 容量维持：${item.target} 侧下沉了 ${item.moved} 条较旧条目（${item.usageBefore} → ${item.usageAfter} 字节），已进归档。`);
        // 「哪几条」必须点名（新-10）：只报条数时，模型与用户都不知道热层忘了什么，
        // 而这几条常常是 critical 的旧结论 —— 本机真实库实测过一次下沉掉两条落地记录。
        const moved = Array.isArray(item.movedPreview) ? item.movedPreview : [];
        for (const entry of moved) {
            const id = asText(entry.id) === '' ? '（无 id）' : entry.id;
            lines.push(`  · [${entry.importance ?? 'normal'}] ${id} ${entry.preview}`);
        }
        if (moved.length > 0) {
            lines.push('  要看全文：office_memory({ action: \'sunk\' }) 按时间列最近几次下沉，或 read 的 layer:\'archive\'。');
        }
    }
    // 单条建议线（新-11）：超过建议线但不超硬上限时只提醒，不拒绝 —— 拒绝一条真实记忆
    // 的代价比多几行字大。回执里说清「为什么提醒」与「怎么办」。
    if (result.advisory !== undefined && result.advisory !== null) {
        lines.push(`- 篇幅提醒：这条 ${result.advisory.bytes} 字节，超过建议线 ${result.advisory.hint} 字节。`
            + '热记忆是「动笔前照办」的清单，建议压成一两句（谁、什么、怎么办），细节放 source 或入库到 kb。');
    }
    lines.push('- 投影已重写：USER.md / MEMORY.md（都是生成物，不要直接编辑）。');
    return lines.join('\n');
}

/**
 * 渲染「最近下沉了什么」（新-10）。
 *
 * 只报 id 与 60 字预览，正文用 archive 读 —— 这个动作回答的是「热层忘了哪几条」，
 * 不是「把内容搬回来」，所以它自己是小回执。
 */
export function renderSunk(result) {
    const lines = [`📒 最近下沉（热记忆容量维持）：共 ${result.totalSinks} 次 / ${result.totalItems} 条`];
    if (result.blocks.length === 0) {
        lines.push('- 还没有发生过下沉（热记忆始终在容量以内）。');
        return lines.join('\n');
    }
    for (const block of result.blocks) {
        const when = String(block.at ?? '').slice(0, 16).replace('T', ' ');
        lines.push(`- ${when}${originTag(block.origin)} 下沉 ${block.items.length} 条：`);
        for (const item of block.items) {
            lines.push(`  · [${item.importance ?? 'normal'}] ${asText(item.id) === '' ? '（无 id）' : item.id} ${item.preview}`);
        }
    }
    if (result.totalSinks > result.blocks.length) {
        lines.push(`- （只给最近 ${result.limit} 次；全部 ${result.totalSinks} 次可用 read 的 layer:'archive' 按关键词找。）`);
    }
    lines.push("- 取回正文：office_memory({ action: 'read', layer: 'archive', query: '关键词' })；归档是只读的，改不了也删不了。");
    return lines.join('\n');
}

/** 渲染台账登记的回执。 */
export function renderLog(result) {    if (result.disabled === true) return '📒 台账层已在设置里关掉（memory.layers.ledger），本次没有登记。';
    const lines = [`📒 台账已登记 ${result.added} 条（现共 ${result.kept} 条）`];
    if (result.rolled > 0) lines.push(`- ${result.rolled} 条最旧记录已滚进归档（台账条数有上限）。`);
    return lines.join('\n');
}

/**
 * 来源档在回执里的写法。
 *
 * `unverified` 一定要点名「待核」：它是**写入期硬规则**的依据（单源内容不能进热记忆），
 * 只在界面里写个 `unverified` 等于让模型自己去猜这个词的分量。
 */
export function tierTag(tier) {
    if (tier === 'verified') return '已验证（≥2 独立来源）';
    if (tier === 'unverified') return '待核（单源未核实）';
    return '用户提供';
}

/** 渲染 kb 入库的回执。 */
export function renderKbIngest(result) {
    const head = result.retiered === true
        ? `📚 kb 里那份内容一个字节没改，来源档已改：${result.path}（${tierTag(result.previousTier)} → ${tierTag(result.tier)}）`
        : result.unchanged === true
            ? `📚 kb 里已经有这份文档的同一版内容（未写入任何字节）：${result.path}`
            : `📚 kb 已入库 ${result.chunks} 块：${result.title}（${result.path}）`;
    const lines = [head];
    lines.push(`- 文档 id：${result.docId}（块 id 是 kb:<内容哈希>:<块号>，从 kb-read 的回执里抄。）`);
    lines.push(`- 来源档：${tierTag(result.tier)}${result.tier === 'unverified' ? ' —— 按硬规则不能进热记忆，只能留在 kb 里待核。' : ''}`);
    if (result.unchanged !== true) {
        lines.push(`- 体积：${result.bytes} 字节 / ${result.chars} 字符；切块按结构（Markdown 标题路径 + 段落），块大小按来源分档。`);
    }
    if (result.replaced === true) {
        lines.push(`- 这是**换代**入库（同一路径、内容变了）：旧块文件已删、指向旧块的边同事务清掉 ${result.prunedLinks ?? 0} 条、旧块 id 留了墓碑。`);
    }
    lines.push('- 读法：office_memory({ action: \'kb-read\', id: \'<文档 id 或块 id>\' })；清单：action:\'kb-list\'。');
    lines.push('- 依据引用：把块 id 写进 add 的 source —— 引用 unverified 的块会被当场拒绝（写入期硬规则）。');
    return lines.join('\n');
}

/** 渲染 kb 清单。 */
export function renderKbList(result) {
    const counts = result.counts ?? { docs: 0, chunks: 0, bytes: 0, tiers: {} };
    const lines = [`📚 kb 清单：${counts.docs} 份文档 / ${counts.chunks} 块 / ${counts.bytes} 字节`
        + `（用户提供 ${counts.tiers?.user ?? 0}，已验证 ${counts.tiers?.verified ?? 0}，待核 ${counts.tiers?.unverified ?? 0}）`];
    if (result.items.length === 0) {
        lines.push('- 还没有入库任何来源。入库：office_memory({ action: \'kb-ingest\', path: \'相对工作目录的文档.md\' })。');
        return lines.join('\n');
    }
    for (const item of result.items) {
        lines.push(`- [${tierTag(item.tier)}] ${item.id} ${item.path} —— ${item.title}（${item.chunks} 块，${item.bytes} 字节，${item.at}）`);
    }
    if (result.truncated === true) lines.push(`- （只列了最近 ${result.items.length} 份，共 ${result.total} 份：要全看就按月份或路径自己挑。）`);
    lines.push('- 读一块：office_memory({ action: \'kb-read\', id: \'<块 id>\' })；读整份：id 用文档 id。');
    return lines.join('\n');
}

/** 渲染 kb 读取的回执（块头前置在正文之前，正文与 span 逐字一致）。 */
export function renderKbRead(result) {
    const lines = [`📚 kb 读取：${result.title}（${result.path}）${originTag(result.origin)}`];
    lines.push(`- 文档 id：${result.docId}；来源档：${tierTag(result.tier)}；入库：${result.at}`);
    lines.push(`- 本次给 ${result.chunks.length} 块 / 共 ${result.chunks_total} 块（额度 ${result.budgetChars} 字符）`
        + `${result.truncated ? '，**还有块没给**：接着用块 id 一块块读，或把来源切小再入库。' : '。'}`);
    if (result.tier === 'unverified') {
        lines.push('- 注意：这份来源是**单源未核实**的。可以照它写摘要或待核清单，但不要把它当结论，也不要据此 add 热记忆（会被硬规则拒绝）。');
    }
    if ((result.chunks_missing ?? 0) > 0) {
        lines.push(`- 警告：有 ${result.chunks_missing} 块的正文文件读不到（manifest 说有几块、盘上却缺）——这些块这次没给出来，不是「文档里没有」。`);
    }
    for (const chunk of result.chunks) {
        lines.push('');
        lines.push(`【第 ${chunk.n} 块｜${chunk.id}｜${chunk.spanText}】`);
        if (chunk.header !== '') lines.push(chunk.header);
        lines.push(chunk.text);
    }
    return lines.join('\n');
}

/**
 * 渲染 kb 检索的回执。
 *
 * 「没命中」的三种成因**必须分开说**（第四十五轮独立复核 P3）：筛选把文档全筛掉了、
 * 文档在但块正文读不到、以及真正的词表零重叠。三者混成一句「弃答」会把读者推向
 * 「换个更贴近原文的词」，而前两种换多少词都没用。
 */
export function renderKbSearch(result) {
    const hits = asArray(result.hits);
    const bounded = result.bounded === true;
    const scanned = `打开 ${result.docsRead ?? 0} 份文档 / 打分 ${result.chunks} 块`;
    const head = result.reason === 'filter-empty'
        ? `📚 kb 检索「${result.query}」：**筛选条件下一份文档都没有**（tier / path 把 ${result.docs ?? 0} 份文档全筛掉了）`
        : result.reason === 'no-chunks'
            ? `📚 kb 检索「${result.query}」：筛出的 ${result.docs} 份文档里**一块正文都读不到**（块文件缺失或坏了）`
            : hits.length === 0
                ? `📚 kb 检索「${result.query}」：**弃答**（${scanned}，没有一块的词表与查询重叠${bounded ? '；但这次**没扫完**，见下' : ''}）`
                : `📚 kb 检索「${result.query}」：命中 ${result.matched} 块，给前 ${hits.length} 块（${scanned}）`;
    const lines = [head];
    if (result.reason === 'filter-empty') {
        lines.push('- 这不是「库里没有这份资料」，而是**筛选**把它排除了：去掉 tier / path 再查一次，'
            + '或先用 action:\'kb-list\' 看清单与每份文档的来源档。');
    } else if (result.reason === 'no-chunks') {
        lines.push('- 这不是「没有这块来源」，而是**正文读不到**：块文件被外面删掉或改坏了（kb-read 在同样的情况下会报缺块）。');
    } else if (hits.length === 0) {
        lines.push('- 弃答是刻意的：kb 的检索是**词法**的（中文 bigram + BM25），查询与正文用词完全对不上时它给不出结果，');
        lines.push('  与其端出「最像的那块」当命中，不如说没有。换更贴近原文的词再试一次，或用 kb-list / kb-read 显式挑读。');
    } else {
        lines.push('- 排序：正文与块头两路 BM25 各自归一化后加权（正文 0.75 / 块头 0.25）。分数是**同一批命中里的相对次序**，不是置信度。');
        lines.push('- 这里是预览（每块开头若干字符）。要逐字引用、要看 span 对应的原文，就用块 id 走 action:\'kb-read\' 读整块 —— 它有独立的读取额度。');
    }
    // 「没扫完」这一条**命中与没命中都要说**：没扫完时的「弃答」只说明「已扫过的那些里没有」。
    if (bounded) {
        lines.push(`- **这次没扫完**：单次上限 ${result.maxChunks} 块，符合筛选的共 ${result.docs} 份文档`
            + `（打开了 ${result.docsRead ?? 0} 份）、已打分 ${result.chunks} 块 —— 更深的命中没被扫到，`
            + '不是「不存在」。缩小 path / tier，或换更贴近原文的词。');
    }
    if ((result.chunksMissing ?? 0) > 0) {
        lines.push(`- 警告：有 ${result.chunksMissing} 块的正文读不到（manifest 说有几块、盘上却缺或自校验对不上）`
            + '——这些块这次没进打分器，不是「文档里没有」。');
    }
    if (hits.length > 0) {
        for (const [index, hit] of hits.entries()) {
            lines.push('');
            lines.push(`【${index + 1}｜分数 ${scoreText(hit.score)}（正文 ${scoreText(hit.scoreText)} / 块头 ${scoreText(hit.scoreHead)}）｜${hit.chunkId}｜${hit.path}｜第 ${hit.n} 块｜${hit.spanText}】`);
            lines.push(`来源档：${tierTag(hit.tier)}${hit.header === '' ? '' : `　${hit.header}`}`);
            lines.push(hit.preview);
        }
    }
    return lines.join('\n');
}

/** 渲染 kb 删除的回执。 */
export function renderKbDrop(result) {
    return [
        `📚 kb 已删除：${result.title}（${result.path}）`,
        `- 文档 id：${result.docId}；删掉 ${result.chunks} 块正文。`,
        `- 同事务清掉关系边 ${result.prunedLinks ?? 0} 条；留墓碑 ${result.tombstones ?? 0} 条（删掉的来源不会被旧 Pack 搬回来）。`,
    ].join('\n');
}

/** kb 五个动作的统一渲染入口（按结果里的 action 分派）。 */
export function renderKb(result) {
    if (result.action === 'kb-ingest') return renderKbIngest(result);
    if (result.action === 'kb-list') return renderKbList(result);
    if (result.action === 'kb-read') return renderKbRead(result);
    if (result.action === 'kb-search') return renderKbSearch(result);
    if (result.action === 'kb-drop') return renderKbDrop(result);
    return `📚 kb：${JSON.stringify(result)}`;
}

/** 渲染建关系的回执。 */
export function renderLink(result) {
    const link = result.link ?? {};
    const head = result.added === true
        ? '📒 已建立关系'
        : result.updated === true
            ? '📒 这条关系已经存在，冲突状态已更新'
            : '📒 这条关系已经存在（未重复写入）';
    const lines = [
        head,
        `- ${link.id ?? ''} ${link.sourceId ?? ''} —[${link.kind ?? ''}]→ ${link.targetId ?? ''}${link.note ? `（${link.note}）` : ''}`,
    ];
    if (link.kind === 'contradicts') {
        lines.push(`- 冲突：类别 ${link.conflictClass || '未分类'}，当前 ${conflictStateText(link.conflictState)}`
            + '（保留双方、不做自动裁决；要改判断就再调一次 link，state 取 unresolved / prefer-source / prefer-target）。');
    }
    lines.push('- 关系是双向的：从任一端用 office_memory({ action: \'related\', id }) 都能走到对面。');
    return lines.join('\n');
}

/** 渲染删关系的回执。 */
export function renderUnlink(result) {
    return `📒 已删除关系 ${result.link?.id ?? ''}（${result.link?.sourceId ?? ''} — ${result.link?.targetId ?? ''}）。`;
}

/** 渲染沿关系找相邻条目的结果。 */
export function renderRelated(result) {
    const lines = [`📒 关系遍历：${result.id}（${result.layer}）${result.text ? ` —— ${truncate(result.text, 80)}` : ''}`];
    const dangling = result.dangling ?? [];
    if (result.nodes.length === 0 && dangling.length === 0) {
        lines.push('- 这条记忆还没有相邻条目。可以用 action:\'link\' 建一条关系。');
        return lines.join('\n');
    }
    if (result.nodes.length > 0) {
        lines.push(`- 相邻 ${result.nodes.length} 条（共 ${result.edges.length} 条边）：`);
        for (const node of result.nodes) {
            const conflict = node.via === 'contradicts' && node.conflictState ? `（冲突：${conflictStateText(node.conflictState)}）` : '';
            lines.push(`  - [${node.layer}/${node.via}${node.hop > 1 ? ` ${node.hop}跳` : ''}] ${node.id}${originTag(node.origin)} ${truncate(node.text, 120)}${conflict}`);
        }
    }
    if (dangling.length > 0) {
        // 悬空边必须**报出来**，不能静默当没有邻居：它是「删了条目忘了清边」的证据。
        lines.push(`- 另有 ${dangling.length} 处悬空（对端已经不在了，不会穿过它继续遍历）：`);
        for (const edge of dangling.slice(0, 5)) {
            lines.push(`  - ${edge.from} —[${edge.kind}]→ ${edge.to}（边 ${edge.id}）`);
        }
        lines.push(`  - 处理：把边删掉用 action:'unlink', id:'<边 id>'；或核对是不是某条记忆被删了没清边。`
            + `库里的悬空边总数看 action:'status' 的 dangling（不变量：必须为 0）。`);
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
    const stores = (result.pack?.stores ?? []).map((store) => `${store.id}（热记忆 ${store.entries.length} / 台账 ${store.ledger.length} / 关系 ${store.links.length} / 墓碑 ${store.tombstones?.length ?? 0} / 归档摘要 ${store.archiveIndex?.digests?.length ?? 0}）`);
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
    lines.push(`- 墓碑 新增 ${result.tombstones ?? 0} 条 / 已存在 ${result.skippedTombstones ?? 0} 条`);
    if ((result.skippedTombstoned ?? 0) > 0) {
        lines.push(`- 有 ${result.skippedTombstoned} 条落在本地墓碑上（已被删除），没有搬回来。`);
    }
    if ((result.skippedDangling ?? 0) > 0) {
        // 老 Pack 里的边可能指着已经不存在的条目（本轮之前 remove 不清边）。
        // 这种边**不能**搬进来，否则一个干净库会继承旧库的悬空引用。
        lines.push(`- 有 ${result.skippedDangling} 条边的一端在库里不存在（老 Pack 的悬空边），没有搬进来。`);
    }
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
    // 悬空引用是「有没有删干净」的唯一现成证据，恒定该是 0。
    lines.push(`- 引用完整性：悬空引用 ${result.dangling ?? 0} 条（不变量：必须为 0）/ 墓碑 ${result.tombstones ?? 0} 条`);
    return lines.join('\n');
}

/**
 * 渲染紧凑摘要：挂在 office_help 与 office_run 的反馈里，等于「动笔前的记忆投影」。
 *
 * `hot: 'unchanged'` 是「只在变化时贴」用的形态：热记忆的**正文**换成一到两行
 * 说明（条数 + revision + 怎么读全文）。`ledger: 'unchanged'` 同理管台账部分
 * （第三十六轮：台账近况同样只在指纹变化时贴，两次交付之间不重复付这三行）。
 * 决定贴哪种形态的是 projection.js，这里只负责渲染。
 *
 * `signal: { label, terms }`（第三十六轮：按需投影）= 「本次在做什么」的话题信号。
 * 给了信号且贴正文时，**只贴命中的条目**，其余折叠成一行（带条数与读取入口）——
 * 「需要用什么才给什么」。四条纪律：
 *   1. 折叠行必须带读取入口：折叠是把选择权交回给模型，不是替它决定不需要；
 *   2. 用量行照旧给全量（容量是「记忆库多大」，不是「这次给了多少」）；
 *   3. 信号匹配允许偏宽（一条记忆可能同时属于两个话题），误命中的代价只是多几行，
 *      漏命中的代价是模型看不到该看的约束。
 *   4. （第四十八轮 P0-2）`budget` 给定时条目正文有字节上限，`skip` 给定时不重复贴本会话
 *      已经贴过的条目 —— 两条都不是丢弃：折叠行照旧报条数与入口。排序按
 *      「命中强度 → 重要度 → 写入时刻」，所以预算先吃掉的是最不相关、最次要、最旧的。
 */
export function renderDigest(digest, {
    context = 'help', hot = 'full', ledger = 'full', signal = null,
    budget = null, skip = null, report = null,
} = {}) {
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
    const ledgerOmitted = ledger === 'unchanged';
    const terms = signal !== null && signal !== undefined && Array.isArray(signal.terms) && signal.terms.length > 0
        ? signal.terms
        : null;
    const label = signal !== null && signal !== undefined && typeof signal.label === 'string' && signal.label !== ''
        ? signal.label
        : '本次';
    const writeRevision = String(digest.writeRevision ?? '').slice(0, 8);
    const lines = [
        omitted && ledgerOmitted
            ? '📒 记忆（热记忆与最近台账都与上次投影相同，正文省略；要看：office_memory({ action: \'read\' })）'
            : omitted
                ? '📒 记忆（热记忆与上次投影相同，省略正文；用量与最近台账照旧）'
                : `📒 记忆（${context === 'run' ? '本次调用后' : '动笔前'}先看这几条${terms === null ? '' : `；按「${label}」筛过，无关条目已折叠`}）`,
    ];
    // 预算（第四十八轮 P0-2）：整段投影共享一个字节额度，只作用于「条目正文」那一块。
    // null / 0 / 非有限值 = 不限（老行为）；显式给数才节流。
    let budgetLeft = Number.isFinite(budget) && budget > 0 ? budget : null;
    for (const target of MEMORY_TARGETS) {
        const items = target === 'user' ? digest.user : digest.project;
        if (items.length === 0) continue;
        lines.push(`【${target === 'user' ? '用户偏好' : '项目与环境'}】`);
        if (omitted) {
            lines.push(`- 与上一次投影相同（${items.length} 条，revision ${writeRevision}），正文不再重复；要看：office_memory({ action: 'read', layer: 'hot' })`);
        } else {
            const matched = terms === null
                ? items
                : items.filter((entry) => matchScore(entrySignalText(entry), terms) > 0);
            const seenBefore = skip instanceof Set
                ? matched.filter((entry) => skip.has(entry.id))
                : [];
            const fresh = seenBefore.length === 0
                ? matched
                : matched.filter((entry) => !skip.has(entry.id));
            // 排序：命中强度 → 重要度 → 写入时刻（新在前）。预算与折叠都按这个次序吃，
            // 于是先被折叠的是「最不相关、最次要、最旧」的那些。没有信号时命中强度一律 0，
            // 排序退化成「重要度 → 新在前」—— 预算有限时先给 critical 的那几条。
            const scoreOf = (entry) => (terms === null ? 0 : matchScore(entrySignalText(entry), terms));
            const ranked = [...fresh].sort((left, right) => {
                const score = scoreOf(right) - scoreOf(left);
                if (score !== 0) return score;
                const rank = (IMPORTANCE_RANK[left.importance] ?? 1) - (IMPORTANCE_RANK[right.importance] ?? 1);
                if (rank !== 0) return rank;
                return String(right.updatedAt ?? '').localeCompare(String(left.updatedAt ?? ''));
            });
            let overBudget = 0;
            for (const entry of ranked) {
                const line = `- [${entry.importance}]${originTag(entry.origin)} ${entry.content.replace(/\s*\n\s*/g, ' ')}`;
                if (budgetLeft !== null && byteLength(line) > budgetLeft) {
                    // 放不下的不丢：记进折叠行的计数，入口照给
                    overBudget += 1;
                    continue;
                }
                if (budgetLeft !== null) budgetLeft -= byteLength(line) + 1;
                lines.push(line);
                if (Array.isArray(report?.printedIds)) report.printedIds.push(entry.id);
            }
            const unrelated = items.length - matched.length;
            const parts = [];
            if (unrelated > 0) parts.push(`${unrelated} 条与「${label}」无关`);
            if (seenBefore.length > 0) parts.push(`${seenBefore.length} 条本会话已贴过`);
            if (overBudget > 0) parts.push(`${overBudget} 条超出本次投影预算`);
            if (parts.length > 0) {
                lines.push(matched.length === 0
                    ? `  与「${label}」相关的条目 0 条，${items.length} 条全部折叠；要看：office_memory({ action: 'read', layer: 'hot' })`
                    : `  另有 ${parts.join('、')}，正文折叠；要看：office_memory({ action: 'read', layer: 'hot' })`);
            }
        }
        lines.push(`  ${usageLine(target, digest.usage[target])}`);
    }
    if (digest.ledger.length > 0) {
        if (ledgerOmitted) {
            lines.push(`【台账】共 ${digest.ledgerTotal} 条，最近 ${digest.ledger.length} 条与上次投影相同，正文省略；要看：office_memory({ action: 'read', layer: 'ledger' })`);
        } else {
            lines.push(`【台账】共 ${digest.ledgerTotal} 条，最近 ${digest.ledger.length} 条：`);
            for (const record of digest.ledger) {
                const badge = [record.format, record.theme].filter(Boolean).join(' · ');
                lines.push(`- ${String(record.at).slice(0, 10)}${originTag(record.origin)} ${record.path}${badge ? `  [${badge}]` : ''}${record.purpose ? ` —— ${record.purpose}` : ''}`);
            }
        }
    }
    // 固定的两条尾巴只在「这次贴了正文」时说：热记忆与台账都省略时，投影只剩几行
    // 状态，把同样的指引句重复贴进前缀是纯开销（它们在任何一次贴正文的形态里都有）。
    if (!(omitted && ledgerOmitted)) {
        lines.push(`- 完整读取与检索：office_memory({ action: 'read', layer: 'all'|'hot'|'ledger'|'archive', query? })`);
        if (digest.autoCapture === true) {
            lines.push('- 主动记录：用户在这一轮里说出的偏好、纠正、稳定事实，当场用 action:\'add\' 记一条；不必等他说「记住」。');
        }
    }
    // 成本可见（第四十八轮 P0-2）：给了预算时把这次投影的实际字节印在首行 ——
    // 看不到成本就谈不上节流，排查时也不必再写一支探针。口径要写清：预算管的是
    // **条目正文**那一块，整段投影还包括用量行、折叠行与固定尾巴。
    if (budgetLeft !== null) {
        const entryUsed = budget - budgetLeft;
        lines[0] = `${lines[0].replace(/）$/, '')}；本次投影 ${byteLength(lines.join('\n'))} B，`
            + `其中条目 ${entryUsed} B（上限 ${budget} B））`;
    }
    const text = lines.join('\n');
    if (report !== null && typeof report === 'object') {
        report.bytes = byteLength(text);
        if (!Array.isArray(report.printedIds)) report.printedIds = [];
        report.folded = Math.max(0, (digest.user?.length ?? 0) + (digest.project?.length ?? 0) - report.printedIds.length);
    }
    return text;
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
