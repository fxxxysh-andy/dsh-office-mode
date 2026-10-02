/**
 * 记忆浏览面板的数据通道：宿主侧一个**只读** HTTP 路由。
 *
 * 为什么需要一条路由：浏览器半侧只有 settingsScope（配置读写），拿不到工作目录
 * 里的文件；而「记忆浏览面板」要显示的是热记忆、台账、归档、关系与实体的真实
 * 内容。mnemon 为这件事搭了一整套 typert RPC 通道；本插件零依赖，就借宿主自带
 * 的 webServer 注册一个只读端点，比自建协议栈小得多。
 *
 * 安全边界（这是本文件最该被 review 的部分）：
 *   1. **只读**。只实现 GET / HEAD，其余方法一律 405；没有任何写入路径。
 *   2. **只服务宿主认得的工作目录**。`?cwd=` 必须命中下面三个来源之一，否则 404：
 *        a. 本进程里真跑过工具的工作目录（noteWorkspace() 记的，来源是工具执行
 *           上下文里的 cwd）；
 *        b. 宿主工作区登记表里的项目目录（`ctx.workspaceRegistry`，用户自己建过、
 *           跨重启仍在）；
 *        c. 会话日志里出现过的工作目录（`ctx.sessionPersistence` 的 header.cwd，
 *           只在 a、b 都没命中时查一次，兜住「没登记成项目但在那儿跑过会话」）。
 *      三个来源都是**宿主自己的事实**，不是请求方给的路径 —— 没有这一条，一个本地
 *      HTTP 端点就成了「传任意路径读任意目录」的数据外泄面。
 *
 *      第三十一轮为什么要加 b、c：原来只有 a，而 a 是**进程内**的 —— 重启 DSH
 *      之后、或者一个还没跑过办公工具的工作区里，已知根是空的，面板只能显示
 *      「还没有见过任何工作目录」。记忆文件明明就在盘上，用户却看不到。
 *      b 是持久的（工作区登记表落在宿主存储里），c 是会话史里的，两者都不需要
 *      先跑一次办公工具。
 *   3. **不返回记忆目录之外的任何路径**。响应里只出现工作目录的绝对路径与相对
 *      展示路径（工作目录本身在宿主的侧栏里也是可见的）。
 *   4. **有界**。台账 / 归档 / 关系 / 实体都截断到固定条数，工作目录列表也有上限，
 *      端点不会被一份巨大的记忆库拖垮。
 *
 * @module dsh-office-mode/view
 */

import { statSync } from 'node:fs';
import { resolve } from 'node:path';
import { createMemory, byteLength, ENTRY_DELIMITER, memoryExists, memoryPaths } from './memory.js';
import { KB_TIERS } from './kb.js';

/** 端点的绝对路径（无尾斜杠）。 */
export const MEMORY_VIEW_PATH = '/office-memory/snapshot';

/** 一次响应里各层的条数上限（浏览用，不是全量导出）。 */
export const VIEW_LIMITS = {
    ledger: 200,
    archive: 200,
    // 知识库的**文档**条数上限。这里取 `KB_LIST_LIMIT`（20）而不是随手写一个大的：
    // `kbList()` 内部会把 limit 收口到那个硬上限，所以写大了**也不会多拿到**，只会让读
    // 代码的人以为端点能翻 200 篇；写小了反而少拿。文档超过 20 篇时 `counts.kbTruncated`
    // 为真、`counts.kbTotal` 给真实总数，面板照实说明（`test/view.mjs` 用 23 篇钉住）。
    // 块数不进这个上限：面板只列文档，块正文由会话侧的 kb-read 按需读（总数走 counts.kbChunks）。
    kb: 20,
    links: 500,
    entities: 60,
    /** 面板能切换的工作目录个数上限（工作区登记表本身没有条数上限）。 */
    workspaces: 50,
};

/** 见过的记忆根：cwd（规范化绝对路径）→ { cwd, seenAt }。 */
const knownRoots = new Map();
/** 最多记多少个根，避免长时间运行后无限增长。 */
const MAX_ROOTS = 50;

/**
 * 会话史那条来源的缓存：`{service, at, roots}`。
 *
 * 按**服务实例**判有效性（不是按 ctx）：同一个进程里服务被替换掉会立刻重查，
 * 同一实例的连续请求不重复扫会话存储。
 */
let sessionRootsCache = { service: undefined, at: 0, roots: [] };

/** Windows / macOS 的路径比较要忽略大小写（同一目录可能以不同大小写出现）。 */
const CASE_INSENSITIVE_PATHS = process.platform === 'win32' || process.platform === 'darwin';

/** 路径的比较键：能 resolve 就 resolve，Windows / macOS 再折成小写。 */
function pathKey(value) {
    let absolute;
    try {
        absolute = resolve(value);
    } catch {
        absolute = String(value);
    }
    return CASE_INSENSITIVE_PATHS ? absolute.toLowerCase() : absolute;
}

/** 两个路径是不是同一个目录（大小写按平台判）。 */
function samePath(left, right) {
    return pathKey(left) === pathKey(right);
}

/**
 * 免 inject 取一个宿主服务；拿不到就返回 undefined。
 *
 * 与 search.js 的 subagentsOf 同一套判据：cordis 的 ctx 代理读未 inject 的服务
 * 属性会**抛错**，`ctx.get(name)` 才是「服务没挂载时返回 undefined」的那个入口。
 * 记忆面板只是本插件的一项能力，不能因为宿主没装工作区登记表就整个插件起不来。
 */
function serviceOf(ctx, name) {
    if (ctx === undefined || ctx === null || typeof ctx.get !== 'function') return undefined;
    try {
        const service = ctx.get(name);
        return service === undefined || service === null ? undefined : service;
    } catch {
        return undefined;
    }
}

/**
 * 记下一个「本进程真的在这里跑过工具」的工作目录。
 *
 * 只由工具执行路径调用 —— 这是路由信任边界里最紧的那一条，别的入口都不该调它
 * （工作区登记表与会话史那两条来源由 registryWorkspacesOf / sessionWorkspacesOf
 * 在读的时候现取，不走这里）。
 *
 * @param {string} cwd 工具执行上下文里的会话工作目录
 */
export function noteWorkspace(cwd) {
    const value = typeof cwd === 'string' ? cwd.trim() : '';
    if (value === '') return;
    let absolute;
    try {
        absolute = resolve(value);
    } catch {
        return;
    }
    knownRoots.delete(absolute);
    knownRoots.set(absolute, { cwd: absolute, seenAt: new Date().toISOString() });
    if (knownRoots.size > MAX_ROOTS) {
        const oldest = knownRoots.keys().next().value;
        if (oldest !== undefined && oldest !== absolute) knownRoots.delete(oldest);
    }
}

/** 当前见过的根列表（最近见到的排前面）。 */
export function knownWorkspaces() {
    return [...knownRoots.values()].reverse().map((item) => ({ ...item }));
}

/** 测试用：清空已知根。 */
export function resetWorkspaces() {
    knownRoots.clear();
    sessionRootsCache = { service: undefined, at: 0, roots: [] };
}

/** 取一个展示用的短名字（路径最后一段）。 */
function labelOf(cwd) {
    const parts = String(cwd).split(/[\\/]/).filter((part) => part !== '');
    return parts.length === 0 ? String(cwd) : parts[parts.length - 1];
}

/**
 * 宿主工作区登记表里的项目目录（`ctx.workspaceRegistry.list()`）。
 *
 * 这是「重启之后面板还能显示记忆」的主来源：登记表落在宿主的存储里，进程重启后
 * 仍在，而且**不需要先跑过一次办公工具**。同步读取、不写任何东西；服务不在
 * （精简组合 / headless）或读取失败时返回空数组 —— 面板退回老行为。
 */
export function registryWorkspacesOf(ctx) {
    const service = serviceOf(ctx, 'workspaceRegistry');
    if (service === undefined || typeof service.list !== 'function') return [];
    let list;
    try {
        list = service.list();
    } catch {
        return [];
    }
    if (!Array.isArray(list)) return [];
    const out = [];
    for (const item of list) {
        const path = typeof item?.path === 'string' ? item.path.trim() : '';
        if (path === '') continue;
        const title = typeof item?.title === 'string' ? item.title.trim() : '';
        out.push({
            cwd: path,
            label: title === '' ? labelOf(path) : title,
            source: 'workspace',
            // 工作区记录的最近变动时刻（会话挂上来算一次变动）——默认根用它挑
            // 「最近在哪儿干活」，比目录 mtime 靠谱（目录 mtime 只在增删文件时变）。
            atMs: Date.parse(String(item?.updatedAt ?? '')) || 0,
        });
    }
    return out;
}

/**
 * 会话日志里出现过的工作目录（`ctx.sessionPersistence.list()` 的 header.cwd）。
 *
 * 只当最后一道兜底：请求的目录既不在本进程见过的根里、也不在工作区登记表里时
 * 才查一次。理由两条 —— 它要扫一遍会话存储（可能慢），而且它的覆盖面最宽
 * （连「没登记成项目、但在那儿跑过会话」的目录也算），不该当成主来源。
 *
 * 结果按**服务实例**缓存 30 秒：同一个进程里换掉服务实例会立刻重查，
 * 同一实例的连续请求不重复扫盘。
 */
export async function sessionWorkspacesOf(ctx, { ttlMs = 30_000 } = {}) {
    const service = serviceOf(ctx, 'sessionPersistence');
    if (service === undefined || typeof service.list !== 'function') return [];
    const now = Date.now();
    if (sessionRootsCache.service === service && now - sessionRootsCache.at < ttlMs) {
        return sessionRootsCache.roots;
    }
    let snapshots;
    try {
        snapshots = await service.list();
    } catch {
        // 会话存储读不动时退回上一次的结果（可能是空的），不影响其余来源。
        return sessionRootsCache.service === service ? sessionRootsCache.roots : [];
    }
    const newest = new Map();
    for (const snapshot of Array.isArray(snapshots) ? snapshots : []) {
        const cwd = typeof snapshot?.header?.cwd === 'string' ? snapshot.header.cwd.trim() : '';
        if (cwd === '') continue;
        const at = Number(snapshot?.header?.createdAt);
        const key = pathKey(cwd);
        const previous = newest.get(key);
        if (previous === undefined || (Number.isFinite(at) && at > previous.at)) {
            newest.set(key, { cwd, at: Number.isFinite(at) ? at : 0 });
        }
    }
    const roots = [...newest.values()]
        .sort((left, right) => right.at - left.at)
        .map((item) => ({ cwd: item.cwd, label: labelOf(item.cwd), source: 'session', atMs: item.at }));
    sessionRootsCache = { service, at: now, roots };
    return roots;
}

/**
 * 把三个来源合成「面板现在能服务的根」。
 *
 * 顺序即优先级：本进程刚跑过工具的根（最贴近「我正在这儿干活」）→ 工作区登记表
 * （持久、跨重启）→ 会话史（兜底）。去重按平台口径比路径。
 *
 * @returns {Promise<{candidates: object[], listing: object[], requested: object|null}>}
 *   candidates 全部候选；listing 是响应里给面板的那一份（有上限）；
 *   requested 是 `?cwd=` 命中的那一条，没命中是 null
 */
async function resolveRootCandidates(ctx, requested, memory) {
    const candidates = [];
    const push = (item) => {
        if (item === undefined || typeof item.cwd !== 'string' || item.cwd.trim() === '') return;
        const cwd = item.cwd.trim();
        if (candidates.some((existing) => samePath(existing.cwd, cwd))) return;
        candidates.push({ ...item, cwd });
    };
    for (const item of knownWorkspaces()) {
        push({ cwd: item.cwd, label: labelOf(item.cwd), source: 'session', seenAt: item.seenAt, atMs: Date.parse(item.seenAt) || 0 });
    }
    for (const item of registryWorkspacesOf(ctx)) push(item);

    let requestedItem = null;
    if (requested !== '') {
        requestedItem = candidates.find((item) => samePath(item.cwd, requested)) ?? null;
        if (requestedItem === null) {
            const fallback = await sessionWorkspacesOf(ctx);
            requestedItem = fallback.find((item) => samePath(item.cwd, requested)) ?? null;
        }
    }

    let listing = candidates.length > 0 ? candidates : await sessionWorkspacesOf(ctx);
    if (requestedItem !== null && !listing.some((item) => samePath(item.cwd, requestedItem.cwd))) {
        listing = [...listing, requestedItem];
    }
    listing = listing.slice(0, VIEW_LIMITS.workspaces);
    return { candidates, listing, requested: requestedItem };
}

/**
 * 这个根的记忆库「最后动过」的时间（毫秒）；没有记忆时返回 null。
 *
 * 看的是记忆库里的文件而不是目录：目录的 mtime 只在增删条目时变，改一条热记忆
 * 不会动它 —— 拿目录 mtime 当「最近写过」会把顺序判反。
 */
function memoryMtime(root, memory) {
    const paths = memoryPaths(root, memory);
    let newest = null;
    for (const file of [paths.memory, paths.ledger, paths.user, paths.project, paths.dir]) {
        try {
            const at = statSync(file).mtimeMs;
            if (newest === null || at > newest) newest = at;
        } catch {
            // 这个文件还不存在：跳过
        }
    }
    return newest;
}

/**
 * 不给 `?cwd=` 时开哪一个根。
 *
 *   1. 本进程刚跑过办公工具的根最优先 —— 那才是「用户现在在哪儿干活」；
 *   2. 否则在候选里挑一个：先要有记忆（面板一开就有东西看，而不是空态），
 *      再看工作区记录的最近变动时刻（会话挂上来算一次，等于「最近在哪儿干活」），
 *      最后看记忆库文件的改动时间；都相同就按列表顺序。
 *   3. 一条都没有记忆时退到第 2 步的时间比较，仍没有就取列表第一个。
 */
function pickDefaultRoot(listing, memory) {
    if (listing.length === 0) return '';
    const recent = knownWorkspaces()[0];
    if (recent !== undefined && recent.cwd !== '') return recent.cwd;
    const scored = listing.map((item, order) => ({
        item,
        order,
        hasMemory: memoryExists(item.cwd, memory),
        atMs: Number.isFinite(item.atMs) ? item.atMs : 0,
        mtime: memoryMtime(item.cwd, memory) ?? 0,
    }));
    scored.sort((left, right) => {
        if (left.hasMemory !== right.hasMemory) return left.hasMemory ? -1 : 1;
        if (left.atMs !== right.atMs) return right.atMs - left.atMs;
        if (left.mtime !== right.mtime) return right.mtime - left.mtime;
        return left.order - right.order;
    });
    return scored[0].item.cwd;
}

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

/** 大小写不敏感的子串命中（浏览面板的过滤用）。 */
function matches(text, terms) {
    if (terms.length === 0) return true;
    const haystack = String(text).toLowerCase();
    return terms.some((term) => haystack.includes(term));
}

function termsOf(query) {
    return String(query ?? '')
        .toLowerCase()
        .split(/[\s,，、;；|/]+/)
        .map((term) => term.trim())
        .filter((term) => term !== '');
}

/**
 * 归档条目在面板里显示、也用于过滤的一行文本。
 *
 * 与 memory.js 的 archiveItemText 同口径：热记忆与 mnemon 长期记忆显示内容本身，
 * 台账条目显示「路径 —— 用途」。两处必须一致，否则面板里搜得到的条目在模型侧
 * 检索不到（或反过来），那种不一致最难查。
 */
function archiveTextOf(item) {
    if (item.kind === 'hot' || item.kind === 'mnemon') return String(item.content ?? '');
    return [item.path, item.purpose].filter(Boolean).join(' —— ');
}

/**
 * 组装一份给浏览面板的完整快照。
 *
 * 契约是**冻结**的：浏览器半侧按这里的字段名渲染，改名要两边一起改。
 *
 * @param {object} options
 *   root     会话工作目录（必须已经在 knownRoots 里）
 *   memory   已解析的记忆配置
 *   query    选填：台账 / 归档的过滤词
 *   label    选填：面板标题用的短名字（工作区登记表里有标题时用它，比目录名清楚）
 */
export async function buildSnapshot({ root, memory = {}, query = '', label = '' } = {}) {
    const instance = createMemory({ root, memory });
    const terms = termsOf(query);
    const status = await instance.status();
    const raw = await instance.browse({
        ledgerLimit: VIEW_LIMITS.ledger,
        archiveLimit: VIEW_LIMITS.archive,
        linkLimit: VIEW_LIMITS.links,
    });

    const hot = raw.hot
        .filter((entry) => matches(`${entry.content} ${(entry.entities ?? []).join(' ')} ${(entry.tags ?? []).join(' ')}`, terms))
        .map((entry) => ({
            id: entry.id,
            target: entry.target,
            importance: entry.importance,
            origin: entry.origin,
            updatedAt: entry.updatedAt,
            entities: entry.entities ?? [],
            tags: entry.tags ?? [],
            content: entry.content,
        }));

    const ledger = raw.ledger
        .filter((record) => matches(`${record.path} ${record.purpose} ${record.format} ${record.theme} ${(record.outline ?? []).join(' ')}`, terms))
        .map((record) => ({
            id: record.id ?? '',
            at: record.at ?? '',
            path: record.path ?? '',
            format: record.format ?? '',
            theme: record.theme ?? '',
            purpose: record.purpose ?? '',
            outline: (record.outline ?? []).slice(0, 8),
            origin: record.origin,
        }));

    const archive = raw.archive
        .filter((item) => matches(archiveTextOf(item), terms))
        .map((item) => ({
            id: item.id ?? '',
            month: item.month ?? '',
            kind: item.kind ?? '',
            at: item.at ?? '',
            origin: item.origin,
            importance: item.importance ?? '',
            category: item.category ?? '',
            entities: item.entities ?? [],
            tags: item.tags ?? [],
            text: archiveTextOf(item),
        }));

    // 关系也参与过滤：面板的搜索框是「在看到的这一堆里找」，只筛一半会让人以为
    // 搜出来的关系不存在。关系没有正文，就按两端 id、类型与备注匹配。
    const links = raw.links
        .filter((link) => matches(`${link.sourceId} ${link.targetId} ${link.kind} ${link.note}`, terms))
        .map((link) => ({
            id: link.id,
            sourceId: link.sourceId,
            targetId: link.targetId,
            kind: link.kind,
            note: link.note ?? '',
            at: link.at,
            origin: link.origin,
        }));

    const entities = await instance.entities({ query, limit: VIEW_LIMITS.entities });

    // 知识库（第四棵树）：清单行的字段就是页面要显示的字段（路径 / 标题 / 块数 /
    // 来源档 / 入库时刻），不再加工。过滤与其它层同口径（按「能搜到的文本」匹配）。
    //
    // **先过滤、再截断**（第四十二轮独立复核挖出来的顺序错）：`kbList()` 的语义是
    // 「按入库时刻倒序取前 N」，直接拿它当数据源会让搜索只覆盖最新的 N 篇 —— 搜一篇
    // 更旧的文档会得到「0 条」，而它明明在库里。所以这里取全量清单行（manifest 是真源，
    // 不读块正文），过滤之后再截到 `VIEW_LIMITS.kb`。顺带一个好处：逐根计数也从
    // manifest 现算，与 `counts.kb*` 同源 —— 不会出现「存储卡说 99 篇、指标卡说 3 篇」
    // 那种陈旧的 `index.json` 打架（复核 P2-2）。
    //
    // 面板**不给检索器**：一期 kb 还没有检索器（第二十四轮定的顺序：自测集变绿才启用），
    // 所以这里只做子串过滤，命中范围与左上角那个搜索框的其它页签一致。
    const kbAll = (await instance.kbRows())
        .sort((left, right) => String(right.at).localeCompare(String(left.at)));
    const kbHits = kbAll.filter((row) => matches(`${row.path} ${row.title} ${row.tier}`, terms));
    const kb = kbHits.slice(0, VIEW_LIMITS.kb).map((row) => ({
        id: row.id ?? '',
        path: row.path ?? '',
        title: row.title ?? '',
        hash: row.hash ?? '',
        bytes: row.bytes ?? 0,
        chars: row.chars ?? 0,
        chunks: row.chunks ?? 0,
        tier: row.tier ?? '',
        at: row.at ?? '',
        origin: row.origin,
    }));
    /** 全量清单行的合计（未过滤）：面板上的「库里共 N 篇 / M 块 / 体积」用它。 */
    const kbTotals = kbAll.reduce((sum, row) => ({
        docs: sum.docs + 1,
        chunks: sum.chunks + (Number.isFinite(row.chunks) ? row.chunks : 0),
        bytes: sum.bytes + (Number.isFinite(row.bytes) ? row.bytes : 0),
    }), { docs: 0, chunks: 0, bytes: 0 });
    /** 逐根的 kb 计数（同样从 manifest 现算）：docs / chunks / bytes / tiers。 */
    const kbPerStore = new Map();
    for (const row of kbAll) {
        const key = row.origin ?? '';
        if (!kbPerStore.has(key)) {
            kbPerStore.set(key, { docs: 0, chunks: 0, bytes: 0, tiers: Object.fromEntries(KB_TIERS.map((tier) => [tier, 0])) });
        }
        const bucket = kbPerStore.get(key);
        bucket.docs += 1;
        bucket.chunks += Number.isFinite(row.chunks) ? row.chunks : 0;
        bucket.bytes += Number.isFinite(row.bytes) ? row.bytes : 0;
        const tier = KB_TIERS.includes(row.tier) ? row.tier : KB_TIERS[0];
        bucket.tiers[tier] += 1;
    }

    // 两层热记忆的「正文体积」，口径与 memory.js 的 targetUsage 一致
    // （每条 content 的字节数 + 分隔符）—— 面板上的占用与下沉阈值必须是同一把尺子，
    // 否则「显示还有空间」和「实际已经开始下沉」会同时成立。
    // 按**未过滤**的 raw.hot 算：q 只决定列出哪些条目，不该让体积跟着搜索词变。
    const layerBytes = (target) => raw.hot
        .filter((entry) => entry.target === target)
        .reduce((sum, entry) => sum + byteLength(entry.content) + byteLength(ENTRY_DELIMITER), 0);

    return {
        ok: true,
        generatedAt: new Date().toISOString(),
        cwd: root,
        label: asText(label) || labelOf(root),
        query: asText(query),
        config: {
            scope: instance.config.scope,
            userScope: instance.config.userScope,
            layers: { ...instance.config.layers },
            links: instance.config.links,
            autoCapture: instance.config.autoCapture,
            quality: { ...instance.config.recallQuality },
            quota: { ...instance.config.quota },
            limits: {
                userLimitBytes: instance.config.userLimitBytes,
                projectLimitBytes: instance.config.projectLimitBytes,
                ledgerLimit: instance.config.ledgerLimit,
                archiveKeep: instance.config.archiveKeep,
            },
        },
        // stores / counts 的字段名是**冻结契约**：浏览器半侧按它们渲染，
        // 加字段安全、改名或删字段要两边一起改。
        stores: status.stores.map((store) => ({
            id: store.id,
            dir: store.dir,
            hot: store.hot,
            ledger: store.ledger,
            archive: store.archive,
            // 归档摘要**文件**数（上限 archiveKeep 管的是它，不是条目数）。
            archiveFiles: store.archiveFiles ?? 0,
            links: store.links,
            // 知识库计数：与 `counts.kb*` **同源**（都从 manifest 现算），所以存储卡与
            // 指标卡不会各说一套（复核 P2-2：`status()` 优先读 `index.json`，那份是投影，
            // 万一陈旧就会与真源打架）。`dir` 仍取 `status()` 的展示路径。
            kb: Object.assign(
                kbPerStore.get(store.id) ?? { docs: 0, chunks: 0, bytes: 0, tiers: Object.fromEntries(KB_TIERS.map((tier) => [tier, 0])) },
                { dir: store.kb?.dir ?? '' },
            ),
            // 这个根的总体积与文件数。目录不存在时是 0 而不是 undefined：
            // 面板直接渲染这两个数字，「空记忆库」也该显示 0 B / 0 个文件。
            bytes: store.bytes ?? 0,
            files: store.files ?? 0,
        })),
        counts: {
            hot: hot.length,
            ledger: ledger.length,
            archive: archive.length,
            links: links.length,
            entities: entities.total,
            // 知识库的**文档**数（过滤并截断之后，与其它页签的计数同口径：
            // 页签与指标卡要的就是「这一页列了几条」）。
            kb: kb.length,
            kbTotal: kbTotals.docs,
            kbChunks: kbTotals.chunks,
            kbBytes: kbTotals.bytes,
            // 过滤之后的命中集有没有被单次上限截断：面板要如实说「这一页给了多少、
            // 库里一共多少」，而不是让 20 看起来像全部（与 archiveItems 同一条纪律）。
            kbTruncated: kbHits.length > VIEW_LIMITS.kb,
            hotBytes: layerBytes('user'),
            projectBytes: layerBytes('project'),
            // ── 容量口径（与上面几个「列表里现在有几条」分开）──
            //
            // 上面那几个是**过滤并截断之后**的条数（页签、指标卡、时间线用它们，
            // 搜完就跟着变，这是对的）。容量条问的是另一个问题：「离上限还有多远」，
            // 它不该跟着搜索框走，也不该被浏览用的 200/500 截断。
            //
            // 两个口径必须分开报，否则会出现两个具体的错：
            //   - 搜一下，「台账 25 / 500」就变成「台账 3 / 500」；
            //   - 归档的条目数（30）配上摘要文件的上限（60），读起来像快满了，
            //     而实际只用了 1 个文件。
            ledgerTotal: status.stores.reduce((sum, store) => sum + (store.ledger ?? 0), 0),
            archiveFiles: raw.digests ?? 0,
            archiveItems: raw.totals?.archive ?? archive.length,
        },
        hot,
        ledger,
        archive,
        kb,
        links,
        entities: entities.items,
    };
}

function sendJson(res, status, payload) {
    const body = `${JSON.stringify(payload)}\n`;
    res.writeHead(status, {
        'content-type': 'application/json; charset=utf-8',
        'content-length': Buffer.byteLength(body, 'utf8'),
        'cache-control': 'no-store',
    });
    res.end(body);
}

/**
 * 注册记忆浏览路由。
 *
 * @param {object} ctx 插件上下文（需要 ctx.webServer；另外按需读
 *   `ctx.workspaceRegistry` 与 `ctx.sessionPersistence` 作为工作目录的信任来源）
 * @param {{getMemory: () => object}} options getMemory 返回当前已解析的 memory 配置
 * @returns {() => void|undefined} 释放函数；没有 webServer 时返回 undefined
 */
export function registerMemoryView(ctx, { getMemory }) {
    const webServer = typeof ctx?.get === 'function' ? ctx.get('webServer') : undefined;
    if (webServer === undefined || webServer === null || typeof webServer.register !== 'function') {
        return undefined;
    }
    return webServer.register({
        kind: 'exact',
        path: MEMORY_VIEW_PATH,
        handler: async (req, res) => {
            try {
                const method = String(req.method ?? 'GET').toUpperCase();
                if (method !== 'GET' && method !== 'HEAD') {
                    sendJson(res, 405, { ok: false, error: '这个端点只读：只支持 GET。' });
                    return;
                }
                const memory = getMemory() ?? {};
                const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
                const requested = asText(url.searchParams.get('cwd'));
                const query = asText(url.searchParams.get('q'));

                // 信任边界：只服务宿主认得的工作目录（本进程见过的 / 工作区登记表里的 /
                // 会话史里的）。三个来源都在 resolveRootCandidates 里。
                const { listing, requested: requestedItem } = await resolveRootCandidates(ctx, requested, memory);
                const workspaces = listing.map((item) => ({
                    cwd: item.cwd,
                    label: item.label,
                    source: item.source,
                    ...(item.seenAt === undefined ? {} : { seenAt: item.seenAt }),
                }));

                if (memory.enabled === false) {
                    sendJson(res, 200, { ok: false, error: '记忆已在设置里关掉。', workspaces });
                    return;
                }
                if (requested !== '' && requestedItem === null) {
                    sendJson(res, 404, {
                        ok: false,
                        error: '这个工作目录不在已知列表里。它要是宿主里登记过的工作区、或者跑过会话的目录，'
                            + '刷新一次就能认出来；否则先在那个目录的会话里用一次办公工具（例如 office_help）。',
                        workspaces,
                    });
                    return;
                }
                const root = requestedItem !== null ? requestedItem.cwd : pickDefaultRoot(listing, memory);
                if (root === '') {
                    sendJson(res, 200, {
                        ok: false,
                        error: '还没有见过任何工作目录：本进程没跑过办公工具、工作区登记表是空的、会话史里也没有目录。'
                            + '先在某个会话里用一次办公工具（例如 office_help），再回来刷新。',
                        workspaces: [],
                    });
                    return;
                }
                const selected = listing.find((item) => samePath(item.cwd, root));
                const snapshot = await buildSnapshot({ root, memory, query, label: selected?.label ?? '' });
                snapshot.workspaces = workspaces;
                sendJson(res, 200, snapshot);
            } catch (error) {
                sendJson(res, 500, { ok: false, error: `读取记忆失败：${error?.message ?? error}` });
            }
        },
    });
}
