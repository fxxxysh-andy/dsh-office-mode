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

import { randomUUID } from 'node:crypto';
import { readdirSync, statSync } from 'node:fs';
import { mkdir, mkdtemp, readFile, realpath, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, isAbsolute, join, relative, resolve } from 'node:path';
import { createMemory, byteLength, ENTRY_DELIMITER, kbRootDir, memoryExists, memoryPaths } from './memory.js';
import { KB_SEARCH_LIMIT, KB_TIERS, kbFolderInScope, kbFolderTree, normalizeKbFolder } from './kb.js';
import { mergeKbFolders, readKbFolders } from './kb-folders.js';
import { modelStatus, startModelDownload } from './kb-model.js';
import { kbLlmReady, listHostModels } from './kb-llm.js';
import { createGallery, createPillowThumbnailer, galleryRootDir } from './gallery.js';
import { createGalleryHandler, GALLERY_PATH } from './gallery-routes.js';
import { assetSecurityHeaders, readSafeAsset } from './image-asset.js';
import { ocrDocument, ocrResultText, officialOcrEndpointOf, OCR_PURPOSES, OCR_SUPPORTED_EXTENSIONS } from './ocr.js';
import { officialReuseGate } from './ocr-local-request.js';
import { kbPageWindow, normalizeKbSort, sortKbRows } from './kb-page.js';

/** 端点的绝对路径（无尾斜杠）。 */
export const MEMORY_VIEW_PATH = '/office-memory/snapshot';
/** 面板写入端点（第八十五轮）：JSON POST，动作见 handleMemoryMutateRoute。 */
export const MEMORY_MUTATE_PATH = '/office-memory/mutate';
/** 本地模型状态 / 下载端点（第八十八轮）。 */
export const MODEL_STATUS_PATH = '/office-memory/model-status';
export const MODEL_DOWNLOAD_PATH = '/office-memory/model-download';
/** 目录浏览 / 知识库上传（2026-11-19）：不依赖宿主桌面桥，Web 也能选目录、拖文件。 */
export const FS_LIST_PATH = '/office-memory/fs-list';
export const KB_UPLOAD_PATH = '/office-memory/kb-upload';
/** 宿主已配置模型清单（只读）：给上传工具栏的模型选择器用，不返回任何 Key。 */
export const LLM_MODELS_PATH = '/office-memory/llm-models';
export const OCR_PATH = '/office-memory/ocr';
/**
 * 直接送文件内容时的体积上限（展开后的字节）：24 MiB。
 *
 * 与「给本地路径」那条入口的区别是**没有磁盘原文件可依**，内容全在 JSON 里；上限
 * 取 OCR_MAX_BYTES（100 MiB）显然太大 —— 一份 100 MiB 的扫描件先传 base64 再落盘，
 * 只会让面板像卡死。超限明确 413，并让人先压小。
 */
const OCR_DATA_MAX_BYTES = 24 * 1024 * 1024;
/** JSON 体上限：base64 膨胀 4/3，再留 1 MiB 给字段名与其它元数据。 */
const OCR_BODY_LIMIT = Math.ceil(OCR_DATA_MAX_BYTES / 3) * 4 + 1024 * 1024;
/**
 * 知识库上传的 JSON 体上限：与 OCR 内容通道一致（24 MB 原文 → base64 约 32 MB）。
 *
 * 第一百一十轮起，8–24 MB 的 PDF 会先在服务端按页拆成 ≤4 MB 的分片再预览；
 * 请求体上限若还停在 12 MB，这类文件根本到不了拆分逻辑，只会先吃一个 413。
 * 真正的 8 MB 入库上限仍在 kbUpload / kbPreview 里逐类型校验，这里只放宽传输层。
 */
const KB_UPLOAD_BODY_LIMIT = OCR_BODY_LIMIT;
/** 知识库图片资产：只服务 kbDir/assets/ 下的图片，供块正文里的 Markdown 图片引用。 */
export const KB_ASSET_PATH = '/office-memory/kb-asset';
/** 图库端点前缀：列表 / 搜索 / 上传 / 标签 / 描述 / 索引共用一条前缀（gallery-routes.js）。 */
export { GALLERY_PATH };

/**
 * 记忆导出端点的绝对路径（无尾斜杠）。
 *
 * 与快照端点是**两条路由**：快照给面板画界面（有截断、有搜索），导出给「搬走整份记忆」
 * 用（无损、不截断）。合成一条的话，面板浏览的 200 条上限会静默变成导出的上限。
 */
export const EXPORT_VIEW_PATH = '/office-memory/export';

/**
 * 导出格式 → 响应形态（冻结契约：`format=json|md`，缺省 json）。
 *
 * 为什么只有这两种：agent 记忆没有公认的公有协议（各家有自己的 pack / JSONL），
 * 所以采用**单文件**；JSON 走本插件既有的版本化 Pack（导入侧已存在，保证无损），
 * Markdown 走 CommonMark（人读、别的工具也能读）。两条互斥，都从同一个 GET 下载。
 */
export const EXPORT_FORMATS = Object.freeze({
    json: { extension: 'json', contentType: 'application/json; charset=utf-8' },
    md: { extension: 'md', contentType: 'text/markdown; charset=utf-8' },
});

/**
 * 可读版（Markdown）的体积上限：4 MiB。
 *
 * JSON 包不受这个上限约束（它本来就是无损迁移用的，截断等于损坏）；限的是 md：
 * 归档与热记忆都是纯文本，体积可以很大，而可读版的价值在「能从头读到尾」，不是「全」。
 * 超过就截断并在文末写明，完整内容仍在同端点的 JSON 里。
 */
export const EXPORT_MARKDOWN_MAX_BYTES = 4 * 1024 * 1024;

/** 一次响应里各层的条数上限（浏览用，不是全量导出）。 */
export const VIEW_LIMITS = {
    ledger: 200,
    archive: 200,
    // 知识库的**文档**条数上限。这里取 `KB_LIST_LIMIT`（20）而不是随手写一个大的：
    // `kbList()` 内部会把 limit 收口到那个硬上限，所以写大了**也不会多拿到**，只会让读
    // 代码的人以为端点能翻 200 篇；写小了反而少拿。文档超过 20 篇时 `counts.kbTruncated`
    // 为真、`counts.kbTotal` 给真实总数，面板照实说明（`test/view.mjs` 用 23 篇钉住）。
    // 块数不进这个上限：面板按文档列清单（块总数走 counts.kbChunks），块清单与
    // 单块正文由 kbDoc / kbChunk 两个下钻参数按需取，不把整份长文档塞进首屏。
    kb: 20,
    /** 知识块检索一次最多给几条（与 memory.js 的硬上限同源，不在界面上另造一个数）。 */
    kbHits: KB_SEARCH_LIMIT,
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

/**
 * 信任根解析：把 `?cwd=` 解析成一个「宿主认得、可以服务」的根。
 *
 * 记忆浏览与预览两条路由共用它 —— 同一套信任来源、同一套错误文案。不能各写一份：
 * 预览能列出 cwd 里的文件，边界只要有一处放宽，另一处就形同虚设。
 *
 * @param {object} ctx 插件上下文（读 workspaceRegistry / sessionPersistence）
 * @param {string|null|undefined} requestedRaw 请求里的 cwd，空串表示「给个默认根」
 * @param {object} memory 已解析的 memory 配置
 * @returns {Promise<
 *   | { ok: true, root: string, label: string, listing: object[], requested: object|null }
 *   | { ok: false, status: number, error: string, listing: object[] }>}
 *   status 200 表示「还没有任何根」（面板显示空态而不是报错），404 才是明确越权
 */
export async function trustedRootFor(ctx, requestedRaw, memory) {
    const requested = asText(requestedRaw);
    const { listing, requested: requestedItem } = await resolveRootCandidates(ctx, requested, memory);
    if (requested !== '' && requestedItem === null) {
        return {
            ok: false,
            status: 404,
            listing,
            error: '这个工作目录不在已知列表里。它要是宿主里登记过的工作区、或者跑过会话的目录，'
                + '刷新一次就能认出来；否则先在那个目录的会话里用一次办公工具（例如 office_help）。',
        };
    }
    const root = requestedItem !== null ? requestedItem.cwd : pickDefaultRoot(listing, memory ?? {});
    if (root === '') {
        return {
            ok: false,
            status: 200,
            listing: [],
            error: '还没有见过任何工作目录：本进程没跑过办公工具、工作区登记表是空的、会话史里也没有目录。'
                + '先在某个会话里用一次办公工具（例如 office_help），再回来刷新。',
        };
    }
    const selected = listing.find((item) => samePath(item.cwd, root));
    return { ok: true, root, label: selected?.label ?? labelOf(root), listing, requested: requestedItem };
}

/** 响应里给面板的工作目录清单（listing 的公开口径：只留面板用得上的字段）。 */
export function publicWorkspaces(listing) {
    return (Array.isArray(listing) ? listing : []).map((item) => ({
        cwd: item.cwd,
        label: item.label,
        source: item.source,
        ...(item.seenAt === undefined ? {} : { seenAt: item.seenAt }),
    }));
}

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

/** 查询参数里的数字：空串 / 乱写 / 负数一律回落到缺省（出界由业务层再收口）。 */
function parseNumber(value, fallback) {
    const text = asText(value);
    if (text === '') return fallback;
    const n = Number(text);
    return Number.isFinite(n) && n >= 0 ? Math.trunc(n) : fallback;
}

function asArray(value) {
    return Array.isArray(value) ? value : [];
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
 * 一份条目里三个来源口径的条数：全部 / 全局 / 工作区。
 *
 * 只在面板里用来给「全部 / 全局 / 工作区」那颗分段控件标数，不参与容量判断：
 * 容量问的是「离上限还有多远」，那个数必须与过滤无关（见 counts 里的容量段）。
 * 认不出 origin 的行只计入「全部」—— 老数据里可能有这一种，不该把它算进任一根。
 */
function countsByOrigin(rows) {
    const out = { all: 0, global: 0, workspace: 0 };
    for (const row of Array.isArray(rows) ? rows : []) {
        out.all += 1;
        if (row?.origin === 'global') out.global += 1;
        else if (row?.origin === 'workspace') out.workspace += 1;
    }
    return out;
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
 *   origin   选填：global / workspace —— 只列这一个根；空串表示两层都列
 *   kbQuery  选填：知识块检索词（词法检索，命中来自 kbSearch）
 *   kbFolder 选填：知识库虚拟文件夹（空 = 根）—— 文档清单与检索都只看这个子树
 *   kbRecursive 选填：文件夹范围是否含子文件夹（默认 true）
 *   kbOffset 选填：知识库文档清单的起始行（按 kbSort 排序后的全量序，越界回拉末页）
 *   kbLimit  选填：每页条数（1–KB_PAGE_MAX，缺省 VIEW_LIMITS.kb）
 *   kbSort   选填：name / size / time / chunks（拼错回落到 name）
 *   kbSortDir 选填：asc / desc
 *   kbDoc    选填：文档 id —— 返回这份文档的块清单（只有元数据，正文按块取）
 *   kbChunk  选填：块 id —— 返回这一块的完整正文
 */
/**
 * 归档隔离警告只以**字符串**过桥。
 *
 * 后端的 archiveScope 是判定对象，里面有物理路径、模式与旧档归属；浏览器只需要
 * 「为什么这块归档被隔离了」这句人话。这里只收字符串、去空白、去重，非字符串与
 * 空串一律丢弃 —— 快照是公开契约，不能把路径或未核实内容顺手带出去。
 */
function archiveWarningsOf(store) {
    const scope = store !== null && typeof store === 'object' ? store.archiveScope : null;
    const rows = scope !== null && typeof scope === 'object' && Array.isArray(scope.warnings)
        ? scope.warnings
        : [];
    const seen = new Set();
    const out = [];
    for (const row of rows) {
        if (typeof row !== 'string') continue;
        const text = row.trim();
        if (text === '' || seen.has(text)) continue;
        seen.add(text);
        out.push(text);
    }
    return out;
}

export async function buildSnapshot({ root, memory = {}, query = '', label = '', origin = '', kbQuery = '', kbFolder = '', kbRecursive = false, kbOffset = 0, kbLimit = VIEW_LIMITS.kb, kbSort = 'name', kbSortDir = 'asc', kbDoc = '', kbChunk = '', kbRelated = '', kbSummaries = false, kbSummariesOffset = 0, kbSummariesLimit = 10 } = {}) {
    const instance = createMemory({ root, memory });
    const terms = termsOf(query);
    // 只认两个值：其余（拼错、旧客户端、空串）一律当「全部」，不悄悄变成空列表。
    const originKey = origin === 'global' || origin === 'workspace' ? origin : '';
    const inOrigin = (row) => originKey === '' || (row !== null && typeof row === 'object' && row.origin === originKey);
    const status = await instance.status();
    const raw = await instance.browse({
        ledgerLimit: VIEW_LIMITS.ledger,
        archiveLimit: VIEW_LIMITS.archive,
    });

    const hotAll = raw.hot
        .filter((entry) => matches(`${entry.content} ${(entry.entities ?? []).join(' ')} ${(entry.tags ?? []).join(' ')}`, terms));
    const hot = hotAll
        .filter(inOrigin)
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

    const ledgerAll = raw.ledger
        .filter((record) => matches(`${record.path} ${record.purpose} ${record.format} ${record.theme} ${(record.outline ?? []).join(' ')}`, terms));
    const ledger = ledgerAll
        .filter(inOrigin)
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

    const archiveAll = raw.archive
        .filter((item) => matches(archiveTextOf(item), terms));
    const archive = archiveAll
        .filter(inOrigin)
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

    // 知识库（第四棵树）：清单行的字段就是页面要显示的字段（路径 / 标题 / 块数 /
    // 来源档 / 入库时刻），不再加工。过滤与其它层同口径（按「能搜到的文本」匹配）。
    //
    // **先过滤、再排序、最后按页切片**（第四十二轮修过「先截断后过滤」，第九十四轮
    // 把截断换成真分页）：`kbList()` 的语义是「按入库时刻倒序取前 N」，直接拿它当数据
    // 源会让搜索只覆盖最新的 N 篇 —— 搜一篇更旧的文档会得到「0 条」，而它明明在库里。
    // 所以这里取全量清单行（manifest 是真源，不读块正文），在**全量**上过滤与排序，
    // 再切出请求的那一页；只在页内排序会让翻页出现自相矛盾的顺序（看似 a < b 又 b < a）。
    // 顺带一个好处：逐根计数也从 manifest 现算，与 `counts.kb*` 同源 —— 不会出现
    // 「存储卡说 99 篇、指标卡说 3 篇」那种陈旧的 `index.json` 打架（复核 P2-2）。
    //
    // 这里仍只做**文档清单**的子串过滤（左上角搜索框覆盖的也是文档路径 / 标题 / 档位）；
    // 真正的内容检索是 kbQuery 下钻调用的 kbSearch（BM25，打在块正文与块头上），
    // 两条入口各管一件事，别把文档过滤的命中数当成块检索的结果。
    const kbRowsAll = (await instance.kbRows())
        .sort((left, right) => String(right.at).localeCompare(String(left.at)));
    // 文件夹范围：面板的左树决定「现在看哪一层」，文档清单与检索都跟着它收窄。
    // 非法路径已在 memory 侧抛错，这里沿用同一套纯函数，避免视图与动作面各有一套前缀判断。
    let kbFolderKey = '';
    try { kbFolderKey = normalizeKbFolder(kbFolder); } catch { kbFolderKey = ''; }
    // 资源管理器口径：默认只列当前目录的直属文件；要递归必须显式传 kbRecursive=true。
    const wantRecursive = kbRecursive === true;
    // 浏览口径与检索口径分开：检索里空 scope 表示「整个库」，但资源管理器里
    // 根目录就是根目录 —— 默认只列 folder 恰好等于当前路径的直属文件。
    const inFolder = (row) => wantRecursive
        ? kbFolderInScope(row.folder, kbFolderKey, { recursive: true })
        : String(row.folder ?? "") === kbFolderKey;
    const kbVisibleRows = kbRowsAll.filter(inFolder);
    const kbMatches = kbVisibleRows.filter((row) => matches(`${row.path} ${row.title} ${row.tier}`, terms));
    // 知识库是单库：来源筛选（全部 / 全局 / 工作区）只作用于热记忆、台账、归档与关系，
    // 不再把 kb 分成两半（第八十八轮）。
    const kbAll = kbMatches;
    // 排序与分页交给纯函数（src/kb-page.js）：排序键白名单、方向兜底、平局按 id 稳定，
    // 越界 offset 回拉到末页行首 —— 用户在第 2 页删掉唯一一条时不会看到「空库」。
    const kbOrder = normalizeKbSort(kbSort, kbSortDir);
    const kbSorted = sortKbRows(kbAll, kbOrder.key, kbOrder.dir);
    const kbWindow = kbPageWindow(kbSorted.length, kbOffset, kbLimit);
    const kb = kbSorted.slice(kbWindow.offset, kbWindow.offset + kbWindow.limit).map((row) => ({
        id: row.id ?? '',
        path: row.path ?? '',
        folder: row.folder ?? '',
        title: row.title ?? '',
        hash: row.hash ?? '',
        bytes: row.bytes ?? 0,
        chars: row.chars ?? 0,
        chunks: row.chunks ?? 0,
        tier: row.tier ?? '',
        at: row.at ?? '',
        origin: row.origin,
        ...(Array.isArray(row.images) && row.images.length > 0 ? { images: row.images } : {}),
    }));
    /** 全量清单行的合计（未过滤）：面板上的「库里共 N 篇 / M 块 / 体积」用它。 */
    const kbTotals = kbVisibleRows.reduce((sum, row) => ({
        docs: sum.docs + 1,
        chunks: sum.chunks + (Number.isFinite(row.chunks) ? row.chunks : 0),
        bytes: sum.bytes + (Number.isFinite(row.bytes) ? row.bytes : 0),
    }), { docs: 0, chunks: 0, bytes: 0 });
    // 存储卡问的是「这个库总共占了多少」：不能跟着当前文件夹的浏览口径收窄，
    // 否则停在根目录（只列直属文件）时会显示「文档 0 篇」。
    const kbLibraryTotals = kbRowsAll.reduce((sum, row) => ({
        docs: sum.docs + 1,
        chunks: sum.chunks + (Number.isFinite(row.chunks) ? row.chunks : 0),
        bytes: sum.bytes + (Number.isFinite(row.bytes) ? row.bytes : 0),
    }), { docs: 0, chunks: 0, bytes: 0 });
    // 「全部 / 全局 / 工作区」那颗分段控件上的数：按**当前过滤词**算（与页签计数同口径），
    // 但不受来源筛选本身影响 —— 否则切到「全局」之后「工作区 3」就消失了，没法切回去。
    // 显式创建的空目录没有 manifest 行，单靠 kbFolderTree 永远看不到；这里并进来，
    // 左树与输入框左侧的「范围 picker」读的是同一棵完整树（与当前目录 / 当前页无关）。
    // 只加空节点，不改派生节点的 docs/chunks/bytes。
    let kbTree = kbFolderTree(kbRowsAll);
    let kbFoldersError = '';
    try {
        // 严格读：坏掉的 folders.json 不能当成「一个空目录都没有」，否则界面看起来
        // 正常、用户以为目录丢了；这里记下原因，快照把错误带给面板显示。
        kbTree = mergeKbFolders(kbTree, await readKbFolders(kbRootDir(memory), { strict: true }));
    } catch (error) {
        kbFoldersError = String(error?.message ?? error);
    }

    const originCounts = {
        hot: countsByOrigin(hotAll),
        ledger: countsByOrigin(ledgerAll),
        archive: countsByOrigin(archiveAll),
    };

    // 知识库的两个下钻与检索：只在请求里点名时才算，普通浏览不多付一次扫描。
    // 失败不把整份快照打红：面板要显示「这一块读不到」，其余五层照常可看。
    let kbSearchResult = null;
    let kbSearchError = '';
    if (asText(kbQuery) !== '') {
        try {
            kbSearchResult = await instance.kbSearch({
                query: asText(kbQuery),
                limit: VIEW_LIMITS.kbHits,
                folder: kbFolderKey,
                recursive: wantRecursive,
                // 浏览口径与检索口径分开：默认「仅当前层」时在候选阶段就按精确目录过滤
                // （含根目录）；否则子层先占满 TopK，后过滤会把真正的直属命中全漏掉。
                // 工具面 kb-search 的默认语义不受影响（exactFolder 只有面板传）。
                exactFolder: !wantRecursive,
            });
        } catch (error) {
            kbSearchError = `检索知识块失败：${error?.message ?? error}`;
        }

    }
    let kbDocInfo = null;
    let kbDocError = '';
    if (asText(kbDoc) !== '') {
        try {
            kbDocInfo = await instance.kbDoc({ id: asText(kbDoc) });
        } catch (error) {
            kbDocError = `读取知识块清单失败：${error?.message ?? error}`;
        }
    }
    let kbChunkInfo = null;
    let kbChunkError = '';
    if (asText(kbChunk) !== '') {
        try {
            const read = await instance.kbRead({ id: asText(kbChunk) });
            const first = Array.isArray(read.chunks) ? read.chunks[0] : null;
            if (first === null || first === undefined) {
                kbChunkError = '这一块的正文读不到（块文件缺失或坏了）。';
            } else {
                kbChunkInfo = {
                    id: first.id,
                    n: first.n,
                    text: first.text,
                    spanText: first.spanText,
                    header: first.header,
                    chars: String(first.text ?? '').length,
                    docId: read.docId,
                    path: read.path,
                    title: read.title,
                    tier: read.tier,
                    origin: read.origin,
                };
            }
        } catch (error) {
            kbChunkError = `读取知识块失败：${error?.message ?? error}`;
        }
    }

    // 块间关系与摘要块（第八十五轮）：面板在选中一块时问邻居，知识库页签顶部列摘要。
    // 两者都只在请求里点名时才算 —— 普通浏览不打开 SQLite。
    let kbRelatedInfo = null;
    let kbRelatedError = '';
    if (asText(kbRelated) !== '') {
        try {
            kbRelatedInfo = await instance.kbRelated({ chunkId: asText(kbRelated), depth: 1, limit: 20 });
        } catch (error) {
            kbRelatedError = `读取块间关系失败：${error?.message ?? error}`;
        }
    }
    let kbSummariesInfo = null;
    let kbSummariesPage = null;
    let kbSummariesError = '';
    if (kbSummaries === true) {
        try {
            // 先取总数再回拉 offset，越界时落到末页 —— 与文档清单同一套窗口语义。
            const total = await instance.kbSummariesCount();
            const window = kbPageWindow(total, kbSummariesOffset, kbSummariesLimit);
            kbSummariesInfo = await instance.kbSummaries({ limit: window.limit, offset: window.offset });
            kbSummariesPage = window;
        } catch (error) {
            kbSummariesError = `读取摘要块失败：${error?.message ?? error}`;
        }
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
        origin: originKey,
        kbFolder: kbFolderKey,
        kbRecursive: wantRecursive,
        // 生效的排序与分页窗口（offset 是**回拉之后**的值，客户端要写回状态）。
        kbSort: kbOrder,
        kbPage: kbWindow,
        kbTree,
        kbFoldersError,
        config: {
            scope: instance.config.scope,
            userScope: instance.config.userScope,
            layers: { ...instance.config.layers },
            autoCapture: instance.config.autoCapture,
            quality: { ...instance.config.recallQuality },
            quota: { ...instance.config.quota },
            limits: {
                userLimitBytes: instance.config.userLimitBytes,
                projectLimitBytes: instance.config.projectLimitBytes,
                ledgerLimit: instance.config.ledgerLimit,
                archiveKeep: instance.config.archiveKeep,
            },
            // 知识库大模型是「可选增强」：面板只拿得到是否就绪与短标签，绝不回传 Key。
            // 上传工具栏据此决定「用模型切片/摘要」是否可勾选，以及未就绪时照实说明原因。
            kbLlm: (() => {
                const cfg = memory !== null && typeof memory === 'object' && memory.llm !== null && typeof memory.llm === 'object' ? memory.llm : {};
                const ready = kbLlmReady(cfg);
                return {
                    mode: asText(cfg.mode) || 'off',
                    provider: asText(cfg.provider),
                    model: asText(cfg.model),
                    chunking: cfg.chunking === true,
                    summary: cfg.summary === true,
                    ready: ready.ok === true,
                    reason: ready.reason,
                };
            })(),
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
            // 这个根的总体积与文件数。目录不存在时是 0 而不是 undefined：
            // 面板直接渲染这两个数字，「空记忆库」也该显示 0 B / 0 个文件。
            bytes: store.bytes ?? 0,
            files: store.files ?? 0,
            // 归档隔离提示（字符串数组）：不传 archiveScope 对象，避免路径/旧档
            // 归属等内部判定进入浏览器；没有警告时就是空数组。
            warnings: archiveWarningsOf(store),
        })),
        // 知识库单库（第八十八轮）：与记忆根分开报。docs/chunks/bytes 与 `counts.kb*`
        // 同源（都从 manifest 现算），存储卡不会再出现「99 篇 / 3 篇」两套数字。
        kbStore: {
            id: 'kb',
            dir: status.kb?.dir ?? '',
            docs: kbLibraryTotals.docs,
            chunks: kbLibraryTotals.chunks,
            bytes: kbLibraryTotals.bytes,
            tiers: Object.fromEntries(KB_TIERS.map((tier) => [tier, kbRowsAll.filter((row) => row.tier === tier).length])),
        },
        counts: {
            hot: hot.length,
            // 个人偏好 / 项目记忆的真实条数（未过滤、未截断）：面板的第二级按
            // target 分，而不是按来源根分。其它 target 归到项目一侧。
            hotUser: raw.hot.filter((entry) => entry.target === 'user').length,
            hotProject: raw.hot.filter((entry) => entry.target !== 'user').length,
            ledger: ledger.length,
            archive: archive.length,
            // 知识库的**文档**数：过滤之后的**总数**（不是本页条数）。页签与指标卡问的
            // 是「这一层有多少篇」，分页之后若报页长，第 2 页会显示「3 篇」、第 1 页
            // 显示「20 篇」，看起来像删了 17 篇（第九十四轮改口径）。
            kb: kbAll.length,
            // 整库总数（不受当前文件夹 / 搜索收窄）：kbNote 的「库里共 N 篇」与存储卡
            // 用它；分页状态栏用 kbPage.total，那一份才是当前层的数。
            kbTotal: kbLibraryTotals.docs,
            kbChunks: kbLibraryTotals.chunks,
            kbBytes: kbLibraryTotals.bytes,
            // 还有没有下一页。旧的 `kbTruncated` = 「被 20 条上限截断」，面板据此打
            // 「仅显示这一页 · 要全部用 kb-list」；现在它只表示页码末尾，面板用 kbPage
            // 渲染分页控件，不再教用户换工具。
            kbTruncated: kbWindow.hasNext,
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
        originCounts,
        kbSearch: kbSearchResult,
        kbSearchError,
        kbDoc: kbDocInfo,
        kbDocError,
        kbChunk: kbChunkInfo,
        kbChunkError,
        kbRelated: kbRelatedInfo,
        kbRelatedError,
        kbSummaries: kbSummariesInfo,
        kbSummariesPage,
        kbSummariesError,
        hot,
        ledger,
        archive,
        kb,
    };
}

// ── 记忆导出（单文件 · 只读） ───────────────────────────────────────────────
//
// 契约（冻结）：GET/HEAD `/office-memory/export?cwd=&format=json|md`。
//   json = 记忆层既有的版本化 Pack（`exportPack()` 的原样输出，导入侧 importPack 能吃）；
//   md   = 同一份内容的 CommonMark 可读版（不贴知识库块正文，体积不可控）。
// 导出**不落盘**：Pack 只进这一次响应，用户自己决定存哪儿 —— 端点仍是只读的。

/** 导出文件名里的时间戳：`YYYYMMDD-HHmmss`（本地时间，用户看得懂是哪一次导出的）。 */
function exportStamp(date) {
    const pad = (value) => String(value).padStart(2, '0');
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}`
        + `-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

/**
 * Content-Disposition 里的 slug：取工作区目录名，非 `[A-Za-z0-9._-]` 的字符换成 `-`。
 *
 * 为什么必须折：HTTP 头是 latin-1 传输的，中文目录名直接放进去在部分客户端会乱码；
 * 折成 ASCII 后下载下来的文件名至少是可预测的（目录名本身仍完整写在 md/json 里）。
 */
function exportSlugOf(root) {
    const slug = labelOf(root)
        .replace(/[^A-Za-z0-9._-]+/g, '-')
        .replace(/^[-.]+/g, '')
        .replace(/[.-]+$/g, '');
    return slug === '' ? 'workspace' : slug;
}

/** 导出响应的统一写回：GET 带正文，HEAD 只有同样的头（content-length 仍是正文大小）。 */
function sendDownload(res, { method, contentType, body, disposition }) {
    res.writeHead(200, {
        'content-type': contentType,
        'content-length': Buffer.byteLength(body, 'utf8'),
        'cache-control': 'no-store',
        'content-disposition': disposition,
    });
    if (method === 'HEAD') res.end();
    else res.end(body);
}

/**
 * 导出端点的错误回执。
 *
 * 不复用 `sendJson`：HEAD 的错误响应也不该带正文（sendJson 总是写 body），
 * 而 405 需要额外的 `Allow` 头。两者都不该反过来改快照端点的既有行为。
 */
function sendExportError(res, status, payload, { method, headers = {} } = {}) {
    const body = `${JSON.stringify(payload)}\n`;
    res.writeHead(status, {
        'content-type': 'application/json; charset=utf-8',
        'content-length': Buffer.byteLength(body, 'utf8'),
        'cache-control': 'no-store',
        ...headers,
    });
    if (method === 'HEAD') res.end();
    else res.end(body);
}

/** 一行文本（表格 / bullet 里不能有换行）。 */
function oneLine(value) {
    return String(value ?? '').replace(/\s+/g, ' ').trim();
}

/** Markdown 表格单元格：竖线要转义，否则一条记忆里的 `|` 会把列切断。 */
function cell(value) {
    return oneLine(value).replace(/\|/g, '\\|');
}

/** 人类可读的体积（摘要用；精确字节数在 JSON 里）。 */
function formatBytes(value) {
    const bytes = Number.isFinite(value) ? value : 0;
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MiB`;
}

/** 内容的短摘要（关系表里把 id 翻成一句人话）。 */
function shorten(value, max = 32) {
    const text = oneLine(value);
    return text.length <= max ? text : `${text.slice(0, max - 1)}…`;
}

function rememberId(table, id, text) {
    const key = asText(id);
    const value = oneLine(text);
    if (key === '' || value === '' || table.has(key)) return;
    table.set(key, value);
}

/** 关系表里的 id 提示：`m-abc（用户偏好：汇报用 16:9）`。 */
function idHint(table, id) {
    const text = table.get(asText(id));
    return text === undefined ? '' : `（${shorten(text)}）`;
}

/** 记忆根的展示名（与面板的口径一致）。 */
function storeLabel(id) {
    const value = asText(id);
    if (value === 'global') return '全局';
    if (value === 'workspace') return '工作区';
    return value === '' ? '未知' : value;
}

/** 摘要文件路径 → 主干（`archive/2026-09.md` → `2026-09`），与 memory.js 的 archiveStem 同口径。 */
function archiveStemOf(file) {
    const value = asText(file).replace(/\\/g, '/').split('/').pop();
    if (value === '' || value === undefined) return '';
    if (value.endsWith('.md')) return value.slice(0, -3);
    if (value.endsWith('.jsonl')) return value.slice(0, -6);
    return value;
}

/** `.jsonl` 归档正文 → 块；坏行跳过（与记忆层 parseDigestItemText 同一口径）。 */
function parseArchiveItemBlocks(text) {
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
                items: asArray(raw.items),
            });
        } catch {
            // 一行坏了不影响其余块：归档正文是追加写的行文件。
        }
    }
    return blocks;
}

/** 指针块宣告的条目数（老格式按内联条目算）。 */
function countOfBlock(block) {
    return Number.isFinite(block?.count) ? block.count : asArray(block?.ids).length || asArray(block?.items).length;
}

/** 归档块原因的展示名（与 memory.js 的 archiveSection 同文案）。 */
function archiveSectionOf(reason) {
    switch (asText(reason)) {
        case 'hot-overflow': return '热记忆下沉';
        case 'ledger-rollover': return '台账滚动';
        case 'mnemon-migration': return 'mnemon 迁移';
        case 'pack-import': return 'Pack 导入';
        default: return asText(reason) || '归档';
    }
}

/** 热记忆的一条：重要度保留（critical 加粗），标签 / 实体 / 更新时间收进括注。 */
function hotEntryLine(entry) {
    const importance = asText(entry?.importance) || 'normal';
    const badge = importance === 'critical' ? '**[critical]**' : `[${importance}]`;
    const extras = [];
    const tags = asArray(entry?.tags).slice(0, 8);
    const entities = asArray(entry?.entities).slice(0, 8);
    if (tags.length > 0) extras.push(`标签：${tags.join('、')}`);
    if (entities.length > 0) extras.push(`实体：${entities.join('、')}`);
    const updatedAt = asText(entry?.updatedAt);
    if (updatedAt !== '') extras.push(`更新：${updatedAt.slice(0, 10)}`);
    return `- ${badge} ${oneLine(entry?.content)}${extras.length === 0 ? '' : `（${extras.join(' · ')}）`}`;
}

/** 台账表格的一行：日期 | 标题 | 格式 | 路径/来源。 */
function ledgerRowLine(record) {
    const date = asText(record?.at).slice(0, 10) || '—';
    const title = asText(record?.purpose) || labelOf(asText(record?.path)) || '—';
    const format = [asText(record?.format), asText(record?.theme)].filter((part) => part !== '').join(' · ') || '—';
    const where = asText(record?.path) || asArray(record?.source)[0] || '—';
    return `| ${cell(date)} | ${cell(title)} | ${cell(format)} | ${cell(where)} |`;
}

/** 归档条目的一条：三类来源（热记忆下沉 / mnemon 迁移 / 台账滚动）各有各的字段。 */
function archiveItemLine(item) {
    const kind = asText(item?.kind);
    if (kind === 'hot' || kind === 'mnemon') {
        const badge = kind === 'hot'
            ? [asText(item?.origin) || 'hot', asText(item?.importance) || 'normal'].join('/')
            : [asText(item?.category), asText(item?.importance) === '' ? '' : `重要度 ${item?.importance}`].filter((part) => part !== '').join(' · ');
        const extras = [];
        const tags = asArray(item?.tags).slice(0, 8);
        const entities = asArray(item?.entities).slice(0, 8);
        if (tags.length > 0) extras.push(`标签：${tags.join('、')}`);
        if (entities.length > 0) extras.push(`实体：${entities.join('、')}`);
        if (Number.isFinite(item?.linkedCount)) extras.push(`关系边 ${item.linkedCount} 条`);
        return `- [${badge || kind}] ${oneLine(item?.content)}${extras.length === 0 ? '' : `（${extras.join('；')}）`}`;
    }
    const badge = [asText(item?.format), asText(item?.theme)].filter((part) => part !== '').join(' · ');
    const name = archiveTextOf(item);
    const date = asText(item?.at).slice(0, 10);
    return `- ${date === '' ? '' : `${date} `}${oneLine(name)}${badge === '' ? '' : ` [${badge}]`}`;
}

/**
 * 渲染「可读版」导出（CommonMark，无 YAML front matter）。
 *
 * 数据源就是同一份 `exportPack()` 的结果 —— 不在渲染时另外读盘：JSON 与 Markdown
 * 必须是同一时刻的同一份内容，否则「一边有、一边没有」这种差异最难解释。
 *
 * @param {object} pack exportPack() 的返回值
 * @param {object} [options] root / label / generatedAt 只用于标题与元信息
 * @returns {string} Markdown 文本
 */
export function renderMemoryExportMarkdown(pack, { root = '', label = '', generatedAt = '' } = {}) {
    const stores = asArray(pack?.stores);
    const workspace = asText(root) || asText(pack?.workspace);
    const title = asText(label) || labelOf(workspace) || '工作区';

    const hot = [];
    const ledger = [];
    const links = [];
    const tombstones = [];
    const kbRows = [];
    const archiveVolumes = [];
    /** id → 一行文本：关系表把裸 id 翻成「id（内容摘要）」，读的人不必来回翻。 */
    const idText = new Map();

    for (const store of stores) {
        for (const entry of asArray(store?.entries)) {
            hot.push(entry);
            rememberId(idText, entry?.id, entry?.content);
        }
        for (const record of asArray(store?.ledger)) {
            ledger.push(record);
            rememberId(idText, record?.id, asText(record?.purpose) || asText(record?.path));
        }
        for (const link of asArray(store?.links)) links.push(link);
        for (const stone of asArray(store?.tombstones)) tombstones.push(stone);
        for (const row of asArray(store?.kb?.manifest)) kbRows.push(row);

        // 归档正文有两处来源：新包在 `archiveItemFiles`（`.jsonl` 原文），老包内联在
        // `archiveIndex.digests[].blocks[].items`。两种都要读，否则一份老包导出的 md
        // 会把整段归档显示成「只有指针」。
        const itemFiles = store?.archiveItemFiles !== null && typeof store?.archiveItemFiles === 'object'
            ? store.archiveItemFiles
            : {};
        for (const digest of asArray(store?.archiveIndex?.digests)) {
            const stem = archiveStemOf(digest?.file) || archiveStemOf(digest?.itemFile) || asText(digest?.month);
            const text = itemFiles[stem];
            let blocks = typeof text === 'string' && text.trim() !== '' ? parseArchiveItemBlocks(text) : [];
            if (blocks.length === 0) blocks = asArray(digest?.blocks);
            archiveVolumes.push({ digest, blocks });
            for (const block of blocks) {
                for (const item of asArray(block?.items)) rememberId(idText, item?.id, archiveTextOf(item));
            }
        }
    }

    const userEntries = hot.filter((entry) => asText(entry?.target) === 'user');
    const projectEntries = hot.filter((entry) => asText(entry?.target) !== 'user');
    const hotBytesOf = (entries) => entries
        .reduce((sum, entry) => sum + byteLength(String(entry?.content ?? '')) + byteLength(ENTRY_DELIMITER), 0);
    const ledgerBytes = ledger.reduce((sum, record) => sum + byteLength(JSON.stringify(record)), 0);
    const linkBytes = links.reduce((sum, link) => sum + byteLength(JSON.stringify(link)), 0);
    const archiveSummary = archiveVolumes.reduce((sum, volume) => {
        for (const block of asArray(volume.blocks)) {
            sum.blocks += 1;
            sum.items += asArray(block?.items).length || countOfBlock(block);
            sum.bytes += byteLength(JSON.stringify(block));
        }
        return sum;
    }, { blocks: 0, items: 0, bytes: 0 });
    const kbChunks = kbRows.reduce((sum, row) => sum + (Number.isFinite(row?.chunks) ? row.chunks : 0), 0);
    const kbBytes = kbRows.reduce((sum, row) => sum + (Number.isFinite(row?.bytes) ? row.bytes : 0), 0);

    const lines = [];
    lines.push(`# 办公记忆导出 · ${oneLine(title)}`, '');
    lines.push(`- 生成时间：${asText(generatedAt) || new Date().toISOString()}`);
    if (workspace !== '') lines.push(`- 工作区：\`${workspace.replace(/`/g, "'")}\``);
    lines.push(`- 记忆根：${stores.length === 0 ? '无' : stores.map((store) => storeLabel(store?.id)).join(' / ')}`);
    lines.push(`- 包格式：${asText(pack?.format) || '未知'} v${Number.isFinite(pack?.version) ? pack.version : '?'}`);
    lines.push('');
    lines.push('> 本文件由 dsh-office-mode 记忆导出生成，单文件、可直接阅读；无损迁移请用 JSON 格式。');
    lines.push('');

    lines.push('## 摘要', '');
    lines.push(`- 用户偏好（user）：${userEntries.length} 条 / ${formatBytes(hotBytesOf(userEntries))}`);
    lines.push(`- 项目与环境（project）：${projectEntries.length} 条 / ${formatBytes(hotBytesOf(projectEntries))}`);
    lines.push(`- 台账：${ledger.length} 条 / ${formatBytes(ledgerBytes)}`);
    lines.push(`- 归档：${archiveVolumes.length} 卷 / ${archiveSummary.blocks} 块 / ${archiveSummary.items} 条 / ${formatBytes(archiveSummary.bytes)}`);
    lines.push(`- 关系：${links.length} 条 / ${formatBytes(linkBytes)}`);
    lines.push(`- 知识库：${kbRows.length} 篇 / ${kbChunks} 块 / ${formatBytes(kbBytes)}`);
    if (tombstones.length > 0) {
        lines.push(`- 墓碑：${tombstones.length} 条（已删除条目的标记：导入时不会被旧 Pack 复活）`);
    }
    lines.push('');

    lines.push('## 用户偏好', '');
    if (userEntries.length === 0) lines.push('（还没有记下任何用户偏好。）');
    else for (const entry of userEntries) lines.push(hotEntryLine(entry));
    lines.push('');

    lines.push('## 项目与环境', '');
    if (projectEntries.length === 0) lines.push('（还没有记下任何项目约定。）');
    else for (const entry of projectEntries) lines.push(hotEntryLine(entry));
    lines.push('');

    lines.push('## 台账', '');
    if (ledger.length === 0) {
        lines.push('（还没有台账记录。）');
    } else {
        lines.push('| 日期 | 标题 | 格式 | 路径/来源 |', '| --- | --- | --- | --- |');
        for (const record of ledger) lines.push(ledgerRowLine(record));
    }
    lines.push('');

    lines.push('## 归档', '');
    if (archiveVolumes.length === 0) {
        lines.push('（归档层是空的。）');
    } else {
        // 按月份分组：月份是检索与展示的既有口径（分卷只是同月太大时的容器）。
        const months = new Map();
        for (const volume of archiveVolumes) {
            const month = asText(volume?.digest?.month) || '未标注月份';
            if (!months.has(month)) months.set(month, []);
            months.get(month).push(volume);
        }
        for (const [month, volumes] of months) {
            const itemCount = volumes.reduce((sum, volume) => sum + asArray(volume.blocks)
                .reduce((count, block) => count + (asArray(block?.items).length || countOfBlock(block)), 0), 0);
            lines.push(`### ${month}（${volumes.length} 卷 / ${itemCount} 条）`, '');
            for (const volume of volumes) {
                const volumeNo = Number.isFinite(volume?.digest?.volume) ? volume.digest.volume : 1;
                if (volumes.length > 1) lines.push(`#### 第 ${volumeNo} 卷`, '');
                for (const block of asArray(volume.blocks)) {
                    const items = asArray(block?.items);
                    lines.push(`**${asText(block?.at) || '未标注时刻'} · ${archiveSectionOf(block?.reason)}（${items.length || countOfBlock(block)} 条）**`, '');
                    if (items.length === 0) lines.push('- （这一块只留下了指针：正文不在包里。）');
                    else for (const item of items) lines.push(archiveItemLine(item));
                    lines.push('');
                }
            }
        }
    }

    lines.push('## 关系', '');
    if (links.length === 0) {
        lines.push('（还没有记下任何关系。）');
    } else {
        lines.push('| source → target | kind |', '| --- | --- |');
        for (const link of links) {
            const source = `${asText(link?.sourceId)}${idHint(idText, link?.sourceId)}`;
            const target = `${asText(link?.targetId)}${idHint(idText, link?.targetId)}`;
            const kind = [asText(link?.kind) || 'related', asText(link?.note) === '' ? '' : `备注：${link.note}`]
                .filter((part) => part !== '')
                .join(' · ');
            lines.push(`| ${cell(source)} → ${cell(target)} | ${cell(kind)} |`);
        }
    }
    lines.push('');

    lines.push('## 知识库', '');
    if (kbRows.length === 0) {
        lines.push('（知识库还是空的。）');
    } else {
        for (const row of kbRows) {
            const name = asText(row?.title) || asText(row?.path) || asText(row?.id) || '未命名文档';
            const path = asText(row?.path);
            const trailer = [
                `${asText(row?.tier) || 'unknown'} 档`,
                `${Number.isFinite(row?.chunks) ? row.chunks : 0} 块`,
                formatBytes(Number.isFinite(row?.bytes) ? row.bytes : 0),
            ];
            if (asText(row?.at) !== '') trailer.push(`入库 ${row.at}`);
            lines.push(`- **${oneLine(name)}**${path === '' ? '' : `（\`${path.replace(/`/g, "'")}\`）`} · ${trailer.join(' · ')}`);
        }
        lines.push('');
        lines.push('> 块正文在 JSON 包里（`stores[].kb.files`），本文不贴：块正文体积不可控，贴进来会让这份可读版失去可读性。');
    }
    lines.push('');

    lines.push('---', '');
    lines.push('> 本文件是**可读投影**，不是记忆存储：请勿直接编辑后当存储用；完整无损数据见 JSON 格式导出。');
    return `${lines.join('\n')}\n`;
}

/**
 * Markdown 的体积上限：超了按字节截断，并在文末说明「已截断」。
 *
 * 只截可读版，JSON 包不受影响（这是两条互斥格式的前提：一份用来搬，一份用来看）。
 *
 * @param {string} text 已渲染的 Markdown
 * @param {number} [limit] 字节上限（测试可注入小值，正式路径用 EXPORT_MARKDOWN_MAX_BYTES）
 * @returns {{ text: string, truncated: boolean, bytes: number }}
 */
export function clampMemoryMarkdown(text, limit = EXPORT_MARKDOWN_MAX_BYTES) {
    const source = String(text ?? '');
    const buffer = Buffer.from(source, 'utf8');
    if (buffer.length <= limit) return { text: source, truncated: false, bytes: buffer.length };
    const note = '\n\n---\n\n> **已截断**：这份可读版超过体积上限，后面的内容没有写进本文。'
        + '完整内容（无损迁移包）请用 JSON 格式（同一端点 `format=json`）。\n';
    // 预留 8 字节：按字节截断后解码可能在边界产生一个替换字符，别让结果反超上限。
    const budget = Math.max(0, limit - Buffer.byteLength(note, 'utf8') - 8);
    const head = buffer.subarray(0, budget).toString('utf8');
    const output = `${head}${note}`;
    return { text: output, truncated: true, bytes: Buffer.byteLength(output, 'utf8') };
}

/**
 * 导出路由的处理函数。
 *
 * 顺序：方法 → format → 信任根 → 记忆开关 → 生成 Pack → 写回。
 * 信任根与快照共用 `trustedRootFor()`（同一套来源与错误文案）—— 导出能拿到的东西
 * 严格多于快照（没有条数上限），所以边界更不能有第二份实现。
 */
async function handleExportRoute(req, res, ctx, memoryOf) {
    const method = String(req.method ?? 'GET').toUpperCase();
    if (method !== 'GET' && method !== 'HEAD') {
        sendExportError(res, 405, {
            ok: false,
            error: '这个端点只读：只支持 GET / HEAD。',
            hint: '导出不会写任何文件；换成 GET 再试。',
        }, { method, headers: { allow: 'GET, HEAD' } });
        return;
    }
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const rawFormat = asText(url.searchParams.get('format'));
        const format = rawFormat.toLowerCase() || 'json';
        if (!Object.prototype.hasOwnProperty.call(EXPORT_FORMATS, format)) {
            sendExportError(res, 400, {
                ok: false,
                error: `format 只支持 json 或 md（收到 ${JSON.stringify(rawFormat)}）。`,
                hint: '不传 format 就是 json（无损迁移包）；给人读的 Markdown 用 format=md。',
            }, { method });
            return;
        }
        const memory = memoryOf() ?? {};
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            // 空态（trustedRootFor 的 status 200）在导出这里也必须明确报错：
            // 给一份空文件等于告诉调用方「迁移完了」，而实际什么都没搬走。
            sendExportError(res, 404, {
                ok: false,
                error: trust.error,
                hint: '下一步：先在目标目录的会话里用一次办公工具（例如 office_help），或把该目录登记成工作区，再回来导出。',
            }, { method });
            return;
        }
        if (memory.enabled === false) {
            sendExportError(res, 403, {
                ok: false,
                error: '记忆已在设置里关掉，没有可导出的内容。',
                hint: '在设置里重新打开记忆后再导出。',
            }, { method });
            return;
        }
        const instance = createMemory({ root: trust.root, memory });
        const pack = await instance.exportPack();
        const now = new Date();
        const disposition = `attachment; filename="dsh-office-memory-${exportSlugOf(trust.root)}-${exportStamp(now)}.${EXPORT_FORMATS[format].extension}"`;
        if (format === 'json') {
            // 2 空格：人能看、Git diff 也读得懂；形状与 office_memory 的 export 完全一致，
            // 所以任何一份旧包/新包都能被同一个 importPack 读回来。
            sendDownload(res, {
                method,
                contentType: EXPORT_FORMATS[format].contentType,
                body: JSON.stringify(pack, null, 2),
                disposition,
            });
            return;
        }
        const markdown = renderMemoryExportMarkdown(pack, {
            root: trust.root,
            label: trust.label,
            generatedAt: now.toISOString(),
        });
        sendDownload(res, {
            method,
            contentType: EXPORT_FORMATS[format].contentType,
            body: clampMemoryMarkdown(markdown).text,
            disposition,
        });
    } catch (error) {
        sendExportError(res, 500, {
            ok: false,
            error: `导出记忆失败：${error?.message ?? error}`,
        }, { method });
    }
}

/**
 * 读 JSON 请求体；有上限，超了失败，错误上带 413 语义（不把内存当请求体的垃圾桶）。
 *
 * 默认保持旧语义：超限后立即 destroy，连接被掐断。只有显式传
 * `{ drainOverflow: true }` 的调用方（OCR 与知识库上传）才把剩余字节读空后丢掉
 * （不再往 chunks 里攒）—— 掐断连接的结果是浏览器只看到连接重置，状态码与 JSON
 * 谁都到不了页面，面板上就退化成浏览器的 `Failed to execute 'json' ...` 或
 * `fetch failed`，真实原因（请求体太大）反而被藏起来。
 *
 * 读空本身也有显式上界 drainCap：超过就暂停读取（TCP 背压），不把「等我读完」
 * 变成可以无限灌数据的入口。413 在超限那一刻就已交给上层，响应结束后由 Node
 * 关闭这条连接。
 *
 * @param {import('node:http').IncomingMessage} req
 * @param {number} [limit] 允许的请求体字节数
 * @param {{drainOverflow?: boolean, drainCap?: number}} [options]
 */
export function readJsonBody(req, limit = 512 * 1024, options = {}) {
    const drainOverflow = options.drainOverflow === true;
    const drainCap = Number.isFinite(Number(options.drainCap)) && Number(options.drainCap) > limit
        ? Number(options.drainCap)
        : limit * 2;
    return new Promise((settle, fail) => {
        let size = 0;
        let over = false;
        const chunks = [];
        req.on('data', (chunk) => {
            if (over) {
                if (!drainOverflow) return;
                size += chunk.length;
                if (size > drainCap) { try { req.pause(); } catch { /* 已经断开 */ } }
                return;
            }
            size += chunk.length;
            if (size > limit) {
                over = true;
                chunks.length = 0;
                const error = new Error(`请求体超过 ${limit} 字节。`);
                error.status = 413;
                error.code = 'body_too_large';
                error.limit = limit;
                fail(error);
                if (!drainOverflow) { try { req.destroy(); } catch { /* 已经断开 */ } return; }
                try { req.resume(); } catch { /* 已经断开 */ }
                return;
            }
            chunks.push(chunk);
        });
        req.on('end', () => { if (!over) settle(Buffer.concat(chunks).toString('utf8')); });
        req.on('error', (error) => { if (!over) fail(error); });
    });
}

export function sendJson(res, status, payload, headers = {}) {
    const body = `${JSON.stringify(payload)}\n`;
    res.writeHead(status, {
        'content-type': 'application/json; charset=utf-8',
        'content-length': Buffer.byteLength(body, 'utf8'),
        'cache-control': 'no-store',
        ...headers,
    });
    res.end(body);
}

/**
 * 面板写入端点的动作分发（第八十五轮）。
 *
 * 边界与快照 / 导出共用 trustedRootFor：只服务宿主认得的工作目录；POST 必须是
 * application/json（跨站表单发不了这个 content-type，CSRF 那条路被挡住）；body 有上限。
 * 写入动作本身仍走 memory.js 的同一套方法 —— 面板不是第二套实现。
 */
async function handleMemoryMutateRoute(req, res, ctx, memoryOf) {
    const method = String(req.method ?? 'GET').toUpperCase();
    if (method !== 'POST') {
        sendJson(res, 405, { ok: false, error: '这个端点只支持 POST。' });
        return;
    }
    const contentType = String(req.headers?.['content-type'] ?? '');
    if (!contentType.includes('application/json')) {
        sendJson(res, 415, { ok: false, error: 'POST 要用 application/json（跨站表单发不了这个类型）。' });
        return;
    }
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const memory = memoryOf() ?? {};
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            sendJson(res, trust.status, { ok: false, error: trust.error });
            return;
        }
        if (memory.enabled === false) {
            sendJson(res, 403, { ok: false, error: '记忆已在设置里关掉。' });
            return;
        }
        let payload = null;
        try {
            payload = JSON.parse(await readJsonBody(req));
        } catch (error) {
            sendJson(res, 400, { ok: false, error: '请求体不是 JSON：' + (error?.message ?? error) });
            return;
        }
        const action = asText(payload?.action);
        if (action === '') {
            sendJson(res, 400, { ok: false, error: '请求体要带 action。' });
            return;
        }
        const instance = createMemory({ root: trust.root, memory });
        let result = null;
        if (action === 'memory-add') {
            result = await instance.mutate({
                action: 'add',
                target: payload.target,
                content: payload.content,
                importance: payload.importance,
                entities: payload.entities,
                tags: payload.tags,
                tier: payload.tier,
                source: Array.isArray(payload.source) ? payload.source : null,
                scope: payload.scope,
            });
        } else if (action === 'memory-replace') {
            result = await instance.mutate({
                action: 'replace',
                target: payload.target,
                id: payload.id,
                oldText: payload.oldText,
                content: payload.content,
                importance: payload.importance,
                entities: payload.entities,
                tags: payload.tags,
                tier: payload.tier,
                scope: payload.scope,
            });
        } else if (action === 'memory-remove') {
            result = await instance.mutate({
                action: 'remove',
                target: payload.target,
                id: payload.id,
                oldText: payload.oldText,
                scope: payload.scope,
            });
        } else if (action === 'kb-update') {
            result = await instance.kbUpdate({
                id: payload.id,
                title: payload.title,
                folder: payload.folder,
                tier: payload.tier,
            });
        } else if (action === 'kb-edit') {
            result = await instance.kbEdit({ id: payload.id, text: payload.text });
        } else if (action === 'kb-drop') {
            result = await instance.kbDrop({ id: payload.id });
        } else if (action === 'kb-link') {
            result = await instance.kbLink({ sourceId: payload.sourceId, targetId: payload.targetId, kind: payload.kind, note: payload.note });
        } else if (action === 'kb-unlink') {
            result = await instance.kbUnlink({ sourceId: payload.sourceId, targetId: payload.targetId, kind: payload.kind });
        } else if (action === 'kb-folder-create') {
            try {
                result = await instance.kbFolderCreate({ path: payload.path });
            } catch (error) {
                if (['KB_FOLDER_INVALID', 'KB_FOLDER_FORBIDDEN', 'KB_FOLDERS_CORRUPT', 'KB_FOLDERS_BUSY'].includes(error?.code)) {
                    // 400 参数/权限；409 清单坏了（可识别、不会被当成空清单）；503 有别的进程在写。
                    const status = error.code === 'KB_FOLDERS_BUSY' ? 503 : (error.code === 'KB_FOLDERS_CORRUPT' ? 409 : 400);
                    sendJson(res, status, { ok: false, code: String(error.code), error: String(error?.message ?? error) });
                    return;
                }
                throw error;
            }
        } else if (action === 'kb-summarize') {
            result = await instance.kbSummarize({ title: payload.title, text: payload.text, chunkIds: payload.chunkIds, docId: payload.docId });
        } else {
            sendJson(res, 400, { ok: false, error: '不认识的面板动作：' + action });
            return;
        }
        sendJson(res, 200, { ok: true, action, result });
    } catch (error) {
        sendJson(res, 200, { ok: false, error: '写入失败：' + (error?.message ?? error) });
    }
}

/** 本地模型状态：给设置页轮询用；目录来自 memory 配置，不接受请求方传路径。 */
async function handleModelStatusRoute(req, res, ctx, memoryOf) {
    const method = String(req.method ?? 'GET').toUpperCase();
    if (method !== 'GET' && method !== 'HEAD') {
        sendJson(res, 405, { ok: false, error: '这个端点只读：只支持 GET。' });
        return;
    }
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const memory = memoryOf() ?? {};
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            sendJson(res, trust.status, { ok: false, error: trust.error });
            return;
        }
        if (memory.enabled === false) {
            sendJson(res, 403, { ok: false, error: '记忆已在设置里关掉。' });
            return;
        }
        sendJson(res, 200, await modelStatus({ kbDir: kbRootDir(memory) }));
    } catch (error) {
        sendJson(res, 500, { ok: false, error: `读取模型状态失败：${error?.message ?? error}` });
    }
}

/** 启动一次模型下载（后台任务）：POST {kind:'embedding'}。 */
async function handleModelDownloadRoute(req, res, ctx, memoryOf) {
    if (String(req.method ?? '').toUpperCase() !== 'POST') {
        sendJson(res, 405, { ok: false, error: '这个端点只支持 POST。' });
        return;
    }
    if (!String(req.headers?.['content-type'] ?? '').includes('application/json')) {
        sendJson(res, 415, { ok: false, error: 'POST 要用 application/json。' });
        return;
    }
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const memory = memoryOf() ?? {};
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            sendJson(res, trust.status, { ok: false, error: trust.error });
            return;
        }
        if (memory.enabled === false) {
            sendJson(res, 403, { ok: false, error: '记忆已在设置里关掉。' });
            return;
        }
        let payload = null;
        try {
            payload = JSON.parse(await readJsonBody(req, 16 * 1024));
        } catch (error) {
            sendJson(res, 400, { ok: false, error: '请求体不是 JSON：' + (error?.message ?? error) });
            return;
        }
        const kind = asText(payload?.kind);
        if (kind === '') {
            sendJson(res, 400, { ok: false, error: '请求体要带 kind（embedding）。' });
            return;
        }
        const info = await startModelDownload({ kbDir: kbRootDir(memory), kind });
        sendJson(res, 202, { ok: true, model: info });
    } catch (error) {
        sendJson(res, 400, { ok: false, error: `启动下载失败：${error?.message ?? error}` });
    }
}

/** 目录浏览：设置页「选择目录」用，不依赖宿主桌面桥。只列目录，不读文件内容。 */
async function handleFsListRoute(req, res, ctx, memoryOf) {
    if (String(req.method ?? 'GET').toUpperCase() !== 'GET') {
        sendJson(res, 405, { ok: false, error: '这个端点只读：只支持 GET。' });
        return;
    }
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const memory = memoryOf() ?? {};
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            sendJson(res, trust.status, { ok: false, error: trust.error });
            return;
        }
        const wanted = asText(url.searchParams.get('path'));
        const roots = listRoots();
        let current = wanted;
        if (current === '') current = roots[0] ?? '';
        if (current === '' || !isAbsolute(current)) {
            sendJson(res, 400, { ok: false, error: '要一个绝对路径。' });
            return;
        }
        let entries;
        try {
            entries = readdirSync(current, { withFileTypes: true })
                .filter((entry) => entry.isDirectory() || entry.isFile())
                .map((entry) => {
                    const full = join(current, entry.name);
                    const kind = entry.isDirectory() ? 'dir' : 'file';
                    if (kind === 'dir') return { name: entry.name, path: full, kind: 'dir' };
                    let size = 0;
                    let mtime = 0;
                    try {
                        const info = statSync(full);
                        size = info.size;
                        mtime = info.mtimeMs;
                    } catch { /* 读不到大小就留 0：列目录不该因为一个坏文件整页失败 */ }
                    return { name: entry.name, path: full, kind: 'file', size, mtime, ext: extname(entry.name).slice(1).toLowerCase() };
                })
                // 资源管理器口径：文件夹在前（按名），文件在后（按名）。
                .sort((left, right) => {
                    if (left.kind !== right.kind) return left.kind === 'dir' ? -1 : 1;
                    return left.name.localeCompare(right.name, 'zh-Hans-CN');
                })
                .slice(0, 500);
        } catch (error) {
            sendJson(res, 400, { ok: false, error: `读不了「${current}」：${error?.message ?? error}` });
            return;
        }
        const parent = dirname(current);
        sendJson(res, 200, {
            ok: true,
            path: current,
            parent: parent === current ? '' : parent,
            roots,
            entries,
        });
    } catch (error) {
        sendJson(res, 500, { ok: false, error: `列目录失败：${error?.message ?? error}` });
    }
}

/** 413 之后最多再等多久：请求体一直不结束（客户端还在写或卡住）就强制关连接。 */
const OVERFLOW_CLOSE_MS = 1500;

/**
 * 有界拒绝的收尾：结构化 413 已经刷出，请求体还有没读完的字节。OCR 与知识库上传共用。
 *
 * 立刻 close 会让**还在写大 body** 的客户端撞上 ECONNRESET，连已经发出的 413 都读
 * 不到（reviewer 在 40/60 MB 复现过，KB 上传在略超 12 MiB 时也复现过）；一直等又等于
 * 把超限请求变成可以无限灌数据的入口。所以这里只做两件事：请求体读完了（自然结束）
 * 就优雅关闭；超过 OVERFLOW_CLOSE_MS 还没结束就强制销毁。读空本身的字节上界在
 * readJsonBody 的 drainCap（调用方传 2× 上限），到界后暂停读取、靠这里的截止时间兜底。
 *
 * 计时器 unref 并在连接关闭时清掉：不拖住进程退出，也不留下悬空定时器。
 *
 * @param {import('node:http').IncomingMessage} req
 */
function closeOverflowBody(req) {
    const socket = req.socket;
    if (!socket || socket.destroyed) return;
    let settled = false;
    const timer = setTimeout(() => {
        settled = true;
        try { socket.destroy(); } catch { /* 已经断开 */ }
    }, OVERFLOW_CLOSE_MS);
    timer.unref?.();
    const cleanup = () => {
        clearTimeout(timer);
        socket.off?.('close', cleanup);
        req.off?.('end', onEnd);
    };
    const onEnd = () => {
        if (settled) return;
        settled = true;
        cleanup();
        try { socket.end(); } catch { /* 已经断开 */ }
    };
    req.once('end', onEnd);
    socket.once('close', cleanup);
    // 极端情况：413 送出之前请求体就已经读完了（正好越过上限的请求）。
    if (req.readableEnded === true) onEnd();
}

/**
 * 把 OCR 的失败翻成这个端点的 HTTP 状态。
 *
 * 原先的写法是「带 >=400 的状态就照抄，否则一律 400」，结果把三类完全不同的事都
 * 压成 400：上游超时（应 504）、上游网络不通（应 502）、上游 200 里装着
 * { error }（应 502）。页面看到 400 只能以为是自己的请求写错了，状态码这一层诊断
 * 信息就丢了。上游自己的 4xx/5xx 仍然照抄 —— 那是真实状态，不做二次包装。
 */
function ocrFailureStatus(error) {
    const status = Number(error?.status);
    const code = String(error?.code ?? '');
    if (Number.isInteger(status) && status >= 400) return status;
    if (code === 'timeout') return 504;
    if (code === 'network' || code === 'bad_response' || code === 'ocr_error' || code === 'body_interrupted') return 502;
    // 识别开始后上游状态就是 200，失败装在 body.error 里（见 ocr.js 契约第 5 条）。
    if (Number.isInteger(status) && status >= 200 && status < 300) return 502;
    if (code === 'missing_api_key' || code === 'file_missing' || code === 'file_too_large'
        || code === 'file_empty' || code === 'not_a_file') return 400;
    return 500;
}

/**
 * 文件识别（OCR）：POST JSON { path, purpose? }。
 *
 * 只认两种路径：信任根（宿主认得的工作目录）之下，或知识库目录之下。
 * 官方契约（api-docs.deepseek.com/zh-cn/api/ocr）：POST /v1/tools/ocr，multipart 只有
 * file + purpose 两个字段。端点、校验与错误翻译都在 ocr.js，这里只做 HTTP 边界。
 */
async function handleOcrRoute(req, res, ctx, memoryOf, configOf, resolveHostOcrCredential) {
    if (String(req.method ?? '').toUpperCase() !== 'POST') {
        sendJson(res, 405, { ok: false, error: '这个端点只支持 POST。' });
        return;
    }
    if (!String(req.headers?.['content-type'] ?? '').includes('application/json')) {
        sendJson(res, 415, { ok: false, error: 'POST 要用 application/json。' });
        return;
    }
    const memory = memoryOf() ?? {};
    if (memory.enabled === false) {
        sendJson(res, 403, { ok: false, error: '记忆已在设置里关掉。' });
        return;
    }
    const config = configOf() ?? {};
    const ocr = config.ocr ?? {};
    if (ocr.enabled === false) {
        sendJson(res, 403, { ok: false, error: '文件识别已在设置里关掉（内存 > 文件识别）。' });
        return;
    }
    let payload = null;
    try {
        // OCR 的面板请求最大约 34 MB（24 MiB 文件 base64 后），超限也必须回一个
        // 正式的 413：读空到 drainCap（2× 上限）为止，再大就靠背压停住。
        payload = JSON.parse(await readJsonBody(req, OCR_BODY_LIMIT, {
            drainOverflow: true,
            drainCap: OCR_BODY_LIMIT * 2,
        }));
    } catch (error) {
        // 超限与「JSON 语法错」不是一回事：前者是 413（换个更小的文件/先拆页），
        // 后者是 400（把请求体改成合法 JSON）。压成同一个 400 会让人照着提示改错方向。
        if (error?.status === 413) {
            // 先回结构化 413（body 很短，立即刷出）。**不能**在这里声明
            // Connection: close：那会让 Node 在响应写完时立刻关连接，还在写 body 的
            // 客户端会撞上 ECONNRESET，连 413 都读不到（40/60 MB 复现）。
            // 收尾交给 closeOverflowBody：读完或 1.5s 截止后再关。
            sendJson(res, 413, {
                ok: false,
                code: 'body_too_large',
                error: '请求体超过这个端点允许的 ' + OCR_BODY_LIMIT + ' 字节（面板识别的单文件上限是 '
                    + OCR_DATA_MAX_BYTES + ' 字节），这次请求没有被解析。',
                retryable: false,
            });
            closeOverflowBody(req);
            return;
        }
        sendJson(res, 400, { ok: false, error: '请求体不是 JSON：' + (error?.message ?? error) });
        return;
    }
    // 两条入口一条路：面板里已经是文件内容（data）就落临时文件再识别，不再要求客户端
    // 编一个「服务端路径」出来 —— 浏览器的 File 没有路径，硬编一个只会是假的，
    // 而且会把「浏览器所在机器的路径」当成服务端路径（隐私与正确性都不对）。
    const wanted = asText(payload?.path);
    const encoded = typeof payload?.data === 'string' ? payload.data : '';
    if (wanted === '' && encoded === '') {
        sendJson(res, 400, { ok: false, error: '要给一个本地文件路径（path），或直接带文件内容（data 与 name）。' });
        return;
    }
    const purpose = OCR_PURPOSES.includes(asText(payload?.purpose)) ? asText(payload.purpose) : '';
    let temporary = '';
    let temporaryDir = '';
    // fileName 必须和 temporary 同层：catch 里的 scrub 用它替换临时文件路径。
    // 声明在下面的 try 里时，data 入口一旦在建好临时文件之后失败，catch 会抛
    // ReferenceError，宿主只能回一个裸 400 —— 真实错误反而看不见。
    let fileName = '';
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            sendJson(res, trust.status, { ok: false, error: trust.error });
            return;
        }
        // 两条入口：path 走原有的信任边界（只读工作目录 / 知识库子树的 realpath），
        // data 走服务端自选的临时文件 —— 后者完全没有「客户端路径」参与，所以既不需要、
        // 也不该拿客户端给的字符串去做任何路径判断。
        let ocrFile = '';
        if (encoded !== '') {
            const data = Buffer.from(encoded, 'base64');
            if (data.length === 0) {
                sendJson(res, 400, { ok: false, error: '文件内容是空的（data 解不出字节）。' });
                return;
            }
            if (data.length > OCR_DATA_MAX_BYTES) {
                sendJson(res, 413, {
                    ok: false,
                    code: 'too_large',
                    error: '文件有 ' + data.length + ' 字节，超过面板识别的单文件上限 '
                        + OCR_DATA_MAX_BYTES + ' 字节；先压小或拆页再试。',
                });
                return;
            }
            const rawName = basename(String(payload?.name ?? '').replace(/\\/g, '/'));
            const safeName = rawName.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/^\.+/, '').trim() || 'scan';
            const ext = extname(safeName).toLowerCase();
            if (!OCR_SUPPORTED_EXTENSIONS.includes(ext)) {
                sendJson(res, 415, {
                    ok: false,
                    code: 'unsupported_format',
                    error: '这个格式不能识别文字：' + OCR_SUPPORTED_EXTENSIONS.join(' / '),
                });
                return;
            }
            // 临时文件放**私有 OS 临时目录**（mkdtemp 独占、0700 目录 + 0600 文件），
            // 不放工作区的 .office/runtime/：工作区子目录可能被配置或符号链接改到别处，
            // 而且会与用户/同步盘共享同一棵树。识别结束（无论成败）整个目录删掉。
            temporaryDir = await mkdtemp(join(tmpdir(), 'dsh-ocr-'));
            temporary = join(temporaryDir, 'content' + ext);
            await writeFile(temporary, data, { mode: 0o600 });
            ocrFile = temporary;
            fileName = safeName;
        } else {
            // 信任边界：只读两个地方 —— 工作目录子树与知识库目录子树。
            // 先 realpath 再比：符号链接可以指到边界外，字面路径包含判断会被它绕过。
            const target = resolve(wanted);
            let realTarget;
            try {
                realTarget = await realpath(target);
            } catch {
                sendJson(res, 400, { ok: false, code: 'file_missing', error: '要识别的文件不存在或读不到：' + target });
                return;
            }
            const allowed = [];
            for (const dir of [trust.root, kbRootDir(memory)]) {
                if (typeof dir !== 'string' || dir === '') continue;
                try {
                    allowed.push(await realpath(dir));
                } catch {
                    allowed.push(resolve(dir));
                }
            }
            const inside = allowed.some((dir) => {
                const rel = relative(dir, realTarget);
                return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel));
            });
            if (!inside) {
                sendJson(res, 403, { ok: false, error: '只识别工作目录或知识库目录里的文件：' + realTarget });
                return;
            }
            ocrFile = realTarget;
        }
        // Key 的优先级：设置里的 ocr.apiKey（含它折叠进来的环境变量值）永远第一；
        // 只有它为空、设置显式打开复用、且这次请求是本机浏览器直连时，才向宿主
        // 官方供应商借一次。自定义 baseURL / 非本机请求连 resolver 都不调用。
        let apiKey = String(ocr.apiKey ?? '').trim();
        let apiKeySource = '';
        let hostProviderHint = false;
        let reuseHint = '';
        if (apiKey === '') {
            const gate = officialReuseGate({
                enabled: ocr.reuseOfficialProvider === true,
                officialEndpoint: officialOcrEndpointOf(ocr.baseURL) !== undefined,
                req,
            });
            if (!gate.ok) {
                // 不暴露供应商是否存在：只说明这次为什么没有借。
                reuseHint = gate.reason;
            } else if (typeof resolveHostOcrCredential === 'function') {
                let host;
                try {
                    host = await resolveHostOcrCredential({ baseURL: ocr.baseURL });
                } catch {
                    // 借 Key 失败只是一个本地缺 Key 的原因，不能把异常文字带回 HTTP 层。
                    host = undefined;
                }
                if (host?.ok === true && typeof host.apiKey === 'string' && host.apiKey.trim() !== '') {
                    apiKey = host.apiKey.trim();
                    apiKeySource = 'official-provider';
                } else if (host?.eligible === true) {
                    // 官方供应商行在、只是 Key 没取到：缺 Key 文案才提 Models。
                    hostProviderHint = true;
                }
            }
        }
        const result = await ocrDocument({
            filePath: ocrFile,
            purpose: purpose === '' ? ocr.purpose : purpose,
            baseURL: ocr.baseURL,
            apiKey,
            apiKeySource,
            hostProviderHint,
            reuseHint,
            // 环境变量名只从设置里取：请求体里的任何字段都不参与「缺 Key」的诊断文案，
            // 客户端字符串一旦被拼进错误消息，就既是注入面也是外泄面。
            apiKeyEnv: ocr.apiKeyEnv,
            timeoutMs: ocr.timeoutMs,
        });
        // pages 原样保留（带 ref/box 标记，给要还原版面的人）；text 走清洗版，
        // 因为它是直接进知识库正文/输入框的那一份，标记留在里面会污染检索块。
        const text = ocrResultText(result);
        sendJson(res, 200, {
            ok: true,
            // 内容入口不回路径：临时文件在 finally 里就删了，回一个已不存在的路径只会
            // 误导客户端；客户端给的 name 只用来定扩展名，绝不当服务端路径用。
            ...(temporary === '' ? { path: ocrFile } : {}),
            fileName: fileName || result.fileName,
            id: result.id,
            purpose: result.purpose,
            pageCount: result.pageCount,
            pages: result.pages,
            // 上游按页返回 finishReason：不是 stop 的页可能被截断。原样报出来，
            // 客户端据此提示「正文可能不完整」，而不是把半页正文静默拼进去。
            truncatedPages: (Array.isArray(result.pages) ? result.pages : [])
                .filter((page) => page && typeof page === 'object'
                    && typeof page.finishReason === 'string' && page.finishReason !== ''
                    && page.finishReason !== 'stop')
                .map((page) => ({ index: page.index, finishReason: page.finishReason })),
            usage: result.usage,
            bytes: result.bytes,
            ms: result.ms,
            text,
        });
    } catch (error) {
        const status = ocrFailureStatus(error);
        // 错误里可能带上临时目录的绝对路径（上游/本地读文件失败时）。对客户端只
        // 暴露用户自己给的文件名：临时路径既是实现细节，也是这台机器上不该外泄的路径。
        const scrub = (text) => {
            let out = String(text ?? '');
            if (temporary === '') return out;
            out = out.split(temporary).join(fileName || '文件');
            if (temporaryDir !== '') out = out.split(temporaryDir).join('');
            return out;
        };
        sendJson(res, status, {
            ok: false,
            error: scrub(error?.message ?? '文件识别失败。'),
            code: error?.code ?? '',
            retryable: error?.retryable === true,
            hint: scrub(error?.hint ?? ''),
            // 上游响应体的公开分类：只到「空 / HTML-XML / 非 JSON 文本」这一层，
            // 原文（可能是网关页面或带本机路径的报文）不出服务端。
            ...(error?.contentType ? { contentType: String(error.contentType) } : {}),
            ...(error?.bodyKind ? { bodyKind: String(error.bodyKind) } : {}),
        });
    } finally {
        // 无论成功失败都删临时目录：识别内容可能含敏感信息，本机不该留副本。
        if (temporaryDir !== '') {
            await rm(temporaryDir, { recursive: true, force: true }).catch(() => {});
        } else if (temporary !== '') {
            await rm(temporary, { force: true }).catch(() => {});
        }
    }
}

/** 知识库图片资产：GET/HEAD，只认 assets/ 下的四种位图，安全头由 image-asset 统一给。 */
async function handleKbAssetRoute(req, res, ctx, memoryOf) {
    const method = String(req.method ?? 'GET').toUpperCase();
    if (method !== 'GET' && method !== 'HEAD') {
        sendJson(res, 405, { ok: false, error: '这个端点只读：只支持 GET / HEAD。' });
        return;
    }
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const memory = memoryOf() ?? {};
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            sendJson(res, trust.status, { ok: false, error: trust.error });
            return;
        }
        if (memory.enabled === false) {
            sendJson(res, 403, { ok: false, error: '记忆已在设置里关掉。' });
            return;
        }
        // 路径规范化、realpath 根内校验、大小上限、魔数判定都在 image-asset.js：
        // 扩展名不参与 MIME，SVG/BMP 与伪后缀一律 415，符号链接出根一律 403。
        const asset = await readSafeAsset({
            root: kbRootDir(memory),
            relativePath: asText(url.searchParams.get('path')),
        });
        if (!asset.ok) {
            sendJson(res, asset.status, { ok: false, code: asset.code, error: asset.error });
            return;
        }
        res.writeHead(200, assetSecurityHeaders(asset.mime, asset.size));
        res.end(method === 'HEAD' ? undefined : asset.bytes);
    } catch (error) {
        sendJson(res, 500, { ok: false, error: `读取图片失败：${error?.message ?? error}` });
    }
}

/** Windows 返回存在的盘符，其它平台返回根目录。 */
function listRoots() {
    if (process.platform !== 'win32') return ['/'];
    const found = [];
    for (const letter of 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') {
        const root = letter + ':' + '\\';
        try {
            if (statSync(root).isDirectory()) found.push(root);
        } catch { /* 没有这个盘 */ }
    }
    return found;
}

/** 上传入库：POST JSON { name, folder, data(base64), tier? }，服务端只落切片与图片。 */
async function handleKbUploadRoute(req, res, ctx, memoryOf) {
    let streamMode = false;
    let streamStarted = false;
    const writeStreamEvent = (event) => {
        if (!streamStarted) {
            streamStarted = true;
            res.writeHead(200, {
                'content-type': 'application/x-ndjson; charset=utf-8',
                'cache-control': 'no-store',
                'x-content-type-options': 'nosniff',
            });
            res.flushHeaders?.();
        }
        res.write(JSON.stringify(event) + '\n');
    };
    if (String(req.method ?? '').toUpperCase() !== 'POST') {
        sendJson(res, 405, { ok: false, error: '这个端点只支持 POST。' });
        return;
    }
    if (!String(req.headers?.['content-type'] ?? '').includes('application/json')) {
        sendJson(res, 415, { ok: false, error: 'POST 要用 application/json。' });
        return;
    }
    try {
        const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
        const memory = memoryOf() ?? {};
        const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
        if (!trust.ok) {
            sendJson(res, trust.status, { ok: false, error: trust.error });
            return;
        }
        if (memory.enabled === false) {
            sendJson(res, 403, { ok: false, error: '记忆已在设置里关掉。' });
            return;
        }
        let payload = null;
        try {
            // 与 OCR 端点同一条纪律：超限也必须有正式的 413。默认 readJsonBody 会立刻
            // destroy 连接，浏览器只看到 UND_ERR_SOCKET（略超 12 MiB 的 KB 上传已复现）；
            // 这里改成读空到 drainCap 为止，响应发出后交给 closeOverflowBody 有界关闭。
            payload = JSON.parse(await readJsonBody(req, KB_UPLOAD_BODY_LIMIT, {
                drainOverflow: true,
                drainCap: KB_UPLOAD_BODY_LIMIT * 2,
            }));
        } catch (error) {
            // 超限（413）与「JSON 语法错」（400）不是一回事：压成同一个 400 会让人照着
            // 提示去改 JSON，而真正的问题是文件太大。
            if (error?.status === 413) {
                // 结构化 413 先刷出；不能声明 Connection: close（还在写 body 的客户端会
                // ECONNRESET，连 413 都读不到）。收尾交给 closeOverflowBody：请求体结束
                // 或 1.5s 截止后再关。
                sendJson(res, 413, {
                    ok: false,
                    code: 'body_too_large',
                    error: '请求体超过这个端点允许的 ' + KB_UPLOAD_BODY_LIMIT + ' 字节，这次请求没有被解析。',
                    retryable: false,
                });
                closeOverflowBody(req);
                return;
            }
            sendJson(res, 400, { ok: false, error: '请求体不是 JSON：' + (error?.message ?? error) });
            return;
        }
        const name = asText(payload?.name);
        if (name === '') {
            sendJson(res, 400, { ok: false, error: '请求体要带文件名 name。' });
            return;
        }
        // 两条通道（第一百零一轮）：`data` 是二进制（base64），`text` 是**现成正文**。
        // OCR 结果走 `text`：二进制已经被识别过一遍，再传一次慢且可能超限；正文写入
        // 「源文件自己的那一份文档」，并带上 ocr 标记与旧笔记替换。
        const encoded = typeof payload?.data === 'string' ? payload.data : '';
        const encodedText = typeof payload?.text === 'string' ? payload.text : '';
        if (encoded === '' && encodedText === '') {
            sendJson(res, 400, { ok: false, error: '请求体要带文件内容 data（base64）或正文 text（base64）。' });
            return;
        }
        const textBody = encodedText === '' ? null : Buffer.from(encodedText, 'base64').toString('utf8');
        const data = encoded === '' ? null : Buffer.from(encoded, 'base64');
        if (textBody === null && (data === null || data.length === 0)) {
            sendJson(res, 400, { ok: false, error: '文件内容是空的。' });
            return;
        }
        if (textBody !== null && textBody.trim() === '') {
            sendJson(res, 400, { ok: false, error: '正文是空的。' });
            return;
        }
        // 上传是「面板里真实可点的入库」：把 AI 开关与模型选择一起透传到 kbIngest，
        // 并把宿主 llm 传进来（host 通道必须有它，否则只能报不可用）。默认 ai=false。
        const instance = createMemory({ root: trust.root, memory, llm: ctx?.get?.('llm') ?? null });
        // 统一入口的「先看一眼」：只抽取、不落库。客户端点下文件就调这里，用户还没
        // 决定要不要入库；因此这里绝不能写 manifest、建块或调模型（kbPreview 的注释
        // 里写了同一条契约）。响应里不带路径：预览卡显示的是用户自己给的文件名。
        if (payload?.preview === true) {
            if (data === null) {
                sendJson(res, 400, { ok: false, error: '预览需要文件内容 data；正文 text 是识别之后的结果，不能再预览。' });
                return;
            }
            const preview = await instance.kbPreview({ name, data });
            sendJson(res, 200, { ok: true, action: 'kb-preview', result: preview });
            return;
        }
        streamMode = payload?.stream === true;
        const result = await instance.kbUpload({
            name,
            folder: asText(payload?.folder),
            tier: asText(payload?.tier),
            data,
            text: textBody,
            ocr: payload?.ocr === true,
            ai: payload?.ai === true,
            aiModel: asText(payload?.aiModel),
            aiChunking: typeof payload?.aiChunking === 'boolean' ? payload.aiChunking : null,
            aiSummary: typeof payload?.aiSummary === 'boolean' ? payload.aiSummary : null,
            aiEstimate: payload?.aiEstimate === true,
            aiResume: payload?.aiResume === true,
            aiRunLimit: Number.isFinite(payload?.aiRunLimit) ? payload.aiRunLimit : 0,
            onProgress: streamMode ? (event) => writeStreamEvent({ type: 'progress', ...event }) : null,
        });
        if (streamMode) {
            writeStreamEvent({ type: 'result', ok: true, action: 'kb-upload', result });
            res.end();
        } else {
            sendJson(res, 200, { ok: true, action: 'kb-upload', result });
        }
    } catch (error) {
        if (streamMode) {
            writeStreamEvent({ type: 'error', ok: false, error: '上传失败：' + (error?.message ?? error) });
            res.end();
            return;
        }
        sendJson(res, 200, { ok: false, error: '上传失败：' + (error?.message ?? error) });
    }
}

/**
 * 注册记忆面板的路由：浏览快照、整包导出与写入。
 *
 * @param {object} ctx 插件上下文（需要 ctx.webServer；另外按需读
 *   `ctx.workspaceRegistry` 与 `ctx.sessionPersistence` 作为工作目录的信任来源）
 * @param {{getMemory: () => object}} options getMemory 返回当前已解析的 memory 配置
 * @returns {() => void|undefined} 释放函数；没有 webServer 时返回 undefined
 */
/** 图库实例按「根 + 信任目录 + 配置」缓存；设置一改就建新实例，旧实例自然淘汰。 */
const galleryInstances = new Map();
const MAX_GALLERY_INSTANCES = 8;

function galleryInstanceFor(trustRoot, memoryConfig, config) {
    const kbDir = kbRootDir(memoryConfig);
    const root = galleryRootDir(memoryConfig, kbDir) || join(dirname(kbDir), 'gallery');
    const galleryConfig = config?.gallery ?? { textEmbedding: memoryConfig?.embedding };
    const key = root + '\u0000' + trustRoot + '\u0000' + JSON.stringify(galleryConfig);
    const cached = galleryInstances.get(key);
    if (cached !== undefined) return cached;
    const instance = createGallery({
        root,
        cwd: trustRoot === '' ? dirname(root) : trustRoot,
        roots: [trustRoot, kbDir, root].filter((value) => typeof value === 'string' && value !== ''),
        config: galleryConfig,
        thumbnailer: createPillowThumbnailer(),
    });
    if (galleryInstances.size >= MAX_GALLERY_INSTANCES) galleryInstances.clear();
    galleryInstances.set(key, instance);
    return instance;
}

export function registerMemoryView(ctx, { getMemory, getConfig, resolveHostOcrCredential } = {}) {
    const webServer = typeof ctx?.get === 'function' ? ctx.get('webServer') : undefined;
    if (webServer === undefined || webServer === null || typeof webServer.register !== 'function') {
        return undefined;
    }
    const disposers = [
        webServer.register({
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
                    const query = asText(url.searchParams.get('q'));
                    // 三个下钻参数都只是字符串；buildSnapshot 自己收口（origin 只认两个值，
                    // 文档 / 块 id 找不到时把原因写回对应字段，不抛 500 把整块面板打红）。
                    const origin = asText(url.searchParams.get('origin'));
                    const kbQuery = asText(url.searchParams.get('kbq'));
                    const kbFolder = asText(url.searchParams.get('kbfolder'));
                    // 显式递归：只有 kbrecursive=1 才是子树；缺省 / 0 都只给当前层。
                    const kbRecursive = url.searchParams.get('kbrecursive') === '1';
                    // 分页三个参数原样转发；parseNumber 收口（空串 / 乱写 → 缺省），
                    // 出界由 kbPageWindow 回拉 —— 路由层不做业务判断，避免第二套边界。
                    const kbOffset = parseNumber(url.searchParams.get('kboffset'), 0);
                    const kbLimit = parseNumber(url.searchParams.get('kblimit'), VIEW_LIMITS.kb);
                    const kbSort = asText(url.searchParams.get('kbsort'));
                    const kbSortDir = asText(url.searchParams.get('kbsortdir'));
                    const kbDoc = asText(url.searchParams.get('kbdoc'));
                    const kbChunk = asText(url.searchParams.get('kbchunk'));
                    const kbRelated = asText(url.searchParams.get('kbrelated'));
                    const kbSummaries = url.searchParams.get('kbsummaries') === '1';
                    const kbSummariesOffset = parseNumber(url.searchParams.get('kbsummariesoffset'), 0);
                    const kbSummariesLimit = parseNumber(url.searchParams.get('kbsummarieslimit'), 10);

                    // 信任边界：只服务宿主认得的工作目录。解析与错误文案在 trustedRootFor 里，
                    // 预览面板的三条路由与导出路由共用同一份（见 preview-view.js）。
                    const trust = await trustedRootFor(ctx, url.searchParams.get('cwd'), memory);
                    const workspaces = publicWorkspaces(trust.listing);

                    if (memory.enabled === false) {
                        sendJson(res, 200, { ok: false, error: '记忆已在设置里关掉。', workspaces });
                        return;
                    }
                    if (!trust.ok) {
                        sendJson(res, trust.status, { ok: false, error: trust.error, workspaces });
                        return;
                    }
                    const snapshot = await buildSnapshot({ root: trust.root, memory, query, label: trust.label, origin, kbQuery, kbFolder, kbRecursive, kbOffset, kbLimit, kbSort, kbSortDir, kbDoc, kbChunk, kbRelated, kbSummaries, kbSummariesOffset, kbSummariesLimit });
                    snapshot.workspaces = workspaces;
                    sendJson(res, 200, snapshot);
                } catch (error) {
                    sendJson(res, 500, { ok: false, error: `读取记忆失败：${error?.message ?? error}` });
                }
            },
        }),
        webServer.register({
            kind: 'exact',
            path: EXPORT_VIEW_PATH,
            handler: async (req, res) => {
                await handleExportRoute(req, res, ctx, () => getMemory() ?? {});
            },
        }),
        webServer.register({
            kind: 'exact',
            path: MEMORY_MUTATE_PATH,
            handler: async (req, res) => {
                await handleMemoryMutateRoute(req, res, ctx, () => getMemory() ?? {});
            },
        }),
        webServer.register({
            kind: 'exact',
            path: MODEL_STATUS_PATH,
            handler: async (req, res) => {
                await handleModelStatusRoute(req, res, ctx, () => getMemory() ?? {});
            },
        }),
        webServer.register({
            kind: 'exact',
            path: MODEL_DOWNLOAD_PATH,
            handler: async (req, res) => {
                await handleModelDownloadRoute(req, res, ctx, () => getMemory() ?? {});
            },
        }),
        webServer.register({
            kind: 'exact',
            path: FS_LIST_PATH,
            handler: async (req, res) => {
                await handleFsListRoute(req, res, ctx, () => getMemory() ?? {});
            },
        }),
        webServer.register({
            kind: 'exact',
            path: KB_UPLOAD_PATH,
            handler: async (req, res) => {
                await handleKbUploadRoute(req, res, ctx, () => getMemory() ?? {});
            },
        }),
        webServer.register({
            kind: 'exact',
            path: LLM_MODELS_PATH,
            handler: async (req, res) => {
                try {
                    const method = String(req.method ?? 'GET').toUpperCase();
                    if (method !== 'GET' && method !== 'HEAD') {
                        sendJson(res, 405, { ok: false, error: '这个端点只读：只支持 GET。' });
                        return;
                    }
                    const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
                    const memory = getMemory() ?? {};
                    const trust = await trustedRootFor(ctx, asText(url.searchParams.get('cwd')), memory);
                    if (!trust.ok) {
                        sendJson(res, trust.status, { ok: false, error: trust.error });
                        return;
                    }
                    const listed = await listHostModels(ctx?.get?.('llm') ?? null);
                    sendJson(res, 200, { ok: listed.ok === true, providers: listed.providers, error: listed.error });
                } catch (error) {
                    sendJson(res, 200, { ok: false, providers: [], error: '读取宿主模型清单失败：' + (error?.message ?? error) });
                }
            },
        }),
        webServer.register({
            kind: 'exact',
            path: KB_ASSET_PATH,
            handler: async (req, res) => {
                await handleKbAssetRoute(req, res, ctx, () => getMemory() ?? {});
            },
        }),
        webServer.register({
            kind: 'exact',
            path: OCR_PATH,
            handler: async (req, res) => {
                await handleOcrRoute(
                    req, res, ctx,
                    () => getMemory() ?? {},
                    () => (typeof getConfig === 'function' ? getConfig() : {}) ?? {},
                    resolveHostOcrCredential,
                );
            },
        }),
        webServer.register({
            kind: 'exact',
            path: GALLERY_PATH,
            handler: createGalleryHandler({
                ctx,
                galleryFor: (trusted) => galleryInstanceFor(
                    asText(trusted?.root),
                    getMemory() ?? {},
                    (typeof getConfig === 'function' ? getConfig() : {}) ?? {},
                ),
                memoryOf: () => getMemory() ?? {},
                trustedRootFor,
                sendJson,
                readJsonBody,
                hostOf: () => ({
                    attachments: typeof ctx?.get === 'function' ? ctx.get('attachments') : null,
                    llm: typeof ctx?.get === 'function' ? ctx.get('llm') : null,
                }),
            }),
        }),
    ];
    // 这些路由是同一块面板的能力：任何一条注册失败都要把已注册的撤掉，
    // 否则会出现「快照 200、写入 404」的半状态。
    return () => {
        for (const dispose of disposers) {
            if (typeof dispose === 'function') dispose();
        }
    };
}
