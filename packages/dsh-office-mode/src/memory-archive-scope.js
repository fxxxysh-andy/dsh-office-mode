/**
 * 多工作区归档的物理路径 / 知识库虚拟路径隔离。
 *
 * 背景：记忆目录默认是「每个工作区各一份」（<root>/.office/memory），但配置允许
 * 把 memory.dir 指到工作区外，甚至让多个工作区、全局层指向同一个绝对目录。那时
 * 它们的 archive/ 是同一份物理目录：新归档会互相覆盖卷，知识库里的归档投影也会
 * 用同一个 docHash 互相遮蔽。
 *
 * 这个模块只做**判定**，不创建、不移动、不删除任何文件（唯一的文件系统读取是
 * 读 archive/workspace-owner.json 这个显式归属标记）。调用方按返回的 writeDir
 * 决定新写落点；legacyDir 只用于只读兼容，绝不能把归属不明的旧档并进某个工作区。
 *
 * 三条硬规则：
 *   1. 工作区身份用真实物理路径（realpath）算，不用 basename —— 两个都叫 proj
 *      的工作区必须是两个身份。
 *   2. 只有「memoryDir 在真实工作区内」或「archive/workspace-owner.json 显式声明
 *      归属本工作区」时，旧 archive 才算本工作区的。其它情况一律按未归属处理。
 *   3. 共用目录的新写落到 archive/workspaces/<workspaceKey>/，各写各的索引与卷；
 *      旧 archive 原地保留，不复制、不移动、不覆盖。
 */
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync, realpathSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, dirname, isAbsolute, join, relative, resolve, sep } from 'node:path';

/** 共用 archive 下放各工作区命名空间的目录名。 */
export const ARCHIVE_WORKSPACES_DIR = 'workspaces';
/** 全局层与工作区共用 archive 时，全局新写落在这里。 */
export const ARCHIVE_GLOBAL_DIR = 'global';
/** 旧 archive 的显式归属标记（只读；由需要认领旧档的调用方手工放置）。 */
export const ARCHIVE_OWNER_FILE = 'workspace-owner.json';
/** 全局命名空间里的稳定标记：一旦建立，后续调用都沿用 archive/global。 */
export const ARCHIVE_GLOBAL_MARKER = 'scope.json';

const DEFAULT_MEMORY_DIR = '.office/memory';
const SHA1_CHARS = 12;
const WORKSPACE_HASH_CHARS = 16;
const SHORT_KEY_CHARS = 8;
const LABEL_CHARS = 40;

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

function sha1(text) {
    return createHash('sha1').update(String(text)).digest('hex');
}

function sha256(text) {
    return createHash('sha256').update(String(text)).digest('hex');
}

/** Windows 的 \\?\ 长路径前缀对身份比较没有意义，先去掉。 */
function stripLongPrefix(value) {
    return String(value).replace(/^\\\\\?\\/, '');
}

function normalizeSeparators(value) {
    return String(value).replace(/\\/g, '/');
}

function normalizeDrive(value) {
    const text = String(value);
    // Windows 盘符大小写不改变所指：C:\ws 与 c:\ws 是同一处。
    return process.platform === 'win32' ? text.replace(/^([a-zA-Z]):/, (_, letter) => `${letter.toUpperCase()}:`) : text;
}

/** 参与身份哈希的规范串：Windows 大小写不敏感，其它平台保持原样。 */
function identityOf(canonicalPath) {
    const slashed = normalizeSeparators(canonicalPath);
    return process.platform === 'win32' ? slashed.toLowerCase() : slashed;
}

function sanitizeLabel(value) {
    const cleaned = String(value ?? '')
        .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .replace(/^\.+|\.+$/g, '');
    if (cleaned === '') return 'workspace';
    return Array.from(cleaned).slice(0, LABEL_CHARS).join('').trim() || 'workspace';
}

/**
 * 把路径规范成真实物理路径。
 *
 * 路径可能还不存在（memory.dir 是待建目录）：向上找到第一个存在的祖先做 realpath，
 * 再把剩余段接回去 —— 这样父目录是 junction / symlink 时，两个写法仍归一到同一处。
 */
function canonicalizePath(value) {
    const resolved = resolve(String(value));
    let current = resolved;
    const tail = [];
    while (!existsSync(current)) {
        const parent = dirname(current);
        if (parent === current) break;
        tail.unshift(basename(current));
        current = parent;
    }
    let base = current;
    let reliability = 'resolved';
    try {
        base = stripLongPrefix(realpathSync.native(current));
        reliability = tail.length === 0 ? 'realpath' : 'ancestor-realpath';
    } catch {
        try {
            base = stripLongPrefix(realpathSync(current));
            reliability = tail.length === 0 ? 'realpath' : 'ancestor-realpath';
        } catch {
            base = current;
        }
    }
    const restored = tail.length === 0 ? base : join(base, ...tail);
    return { path: normalizeDrive(restored), reliability };
}

/** 两个路径是不是同一处物理位置（大小写、分隔符、symlink 都已归一）。 */
export function samePhysicalPath(left, right) {
    if (asText(left) === '' || asText(right) === '') return false;
    return identityOf(canonicalizePath(left).path) === identityOf(canonicalizePath(right).path);
}

/** child 在 parent 之内（含相等）；按真实路径判断，不按字符串前缀。 */
export function isPathInside(parent, child) {
    if (asText(parent) === '' || asText(child) === '') return false;
    const from = canonicalizePath(parent).path;
    const to = canonicalizePath(child).path;
    const rel = relative(from, to);
    if (rel === '') return true;
    if (isAbsolute(rel)) return false;
    return rel !== '..' && !rel.startsWith(`..${sep}`) && !rel.startsWith('../');
}

/**
 * 一个工作区的稳定身份。
 *
 * key 取真实路径的 sha256 前 16 位；同一处物理目录的别名（.. / junction / symlink）
 * 归一到同一个 key，同名但不同路径的工作区不会碰撞。
 */
export function canonicalWorkspacePath(root) {
    const raw = asText(root);
    if (raw === '') {
        return {
            ok: false,
            path: '',
            identity: '',
            key: 'w-unknown',
            shortKey: 'unknown',
            label: '未知工作区',
            display: '未知工作区',
            reliability: 'unknown',
            error: '工作目录为空',
        };
    }
    const canonical = canonicalizePath(raw);
    const identity = identityOf(canonical.path);
    const key = `w-${sha256(identity).slice(0, WORKSPACE_HASH_CHARS)}`;
    const label = sanitizeLabel(basename(canonical.path) || canonical.path);
    return {
        ok: true,
        path: canonical.path,
        identity,
        key,
        shortKey: key.slice(2, 2 + SHORT_KEY_CHARS),
        label,
        display: `${label} [${key.slice(2, 2 + SHORT_KEY_CHARS)}]`,
        reliability: canonical.reliability,
        error: '',
    };
}

/** 记忆目录解析（与 memory.js 的 memoryPaths 同一口径，不依赖它以免循环引用）。 */
export function resolveMemoryDir(root, memory = {}) {
    const configured = asText(memory?.dir) || DEFAULT_MEMORY_DIR;
    return isAbsolute(configured) ? configured : join(String(root ?? ''), configured);
}

/** 全局记忆目录解析（与 globalMemoryDir 同一口径，含旧 .office-memory 兜底）。 */
export function resolveGlobalMemoryDir(memory = {}) {
    const configured = asText(memory?.globalDir);
    if (configured !== '') return isAbsolute(configured) ? configured : join(homedir(), configured);
    const home = asText(process.env.DSH_HOME) || join(homedir(), '.dsh');
    const current = join(home, '.office', 'memory');
    if (!existsSync(current)) {
        const legacy = join(home, '.office-memory');
        if (existsSync(legacy)) return legacy;
    }
    return current;
}

function archiveDirOf(memoryDir) {
    return join(memoryDir, 'archive');
}

function hasAnyEntry(dir) {
    try {
        return existsSync(dir) && readdirSync(dir).length > 0;
    } catch {
        return false;
    }
}

/** archive 根下有没有真正的旧档内容；workspaces/、global/ 与归属标记不算旧档。 */
function hasLegacyArchiveEntries(archiveDir) {
    try {
        if (!existsSync(archiveDir)) return false;
        return readdirSync(archiveDir).some((name) => name !== ARCHIVE_WORKSPACES_DIR && name !== ARCHIVE_GLOBAL_DIR && name !== ARCHIVE_OWNER_FILE);
    } catch {
        return false;
    }
}

/** 全局命名空间标记的位置（archive/global/scope.json）。 */
export function globalNamespaceMarkerPath(archiveDir) {
    return join(String(archiveDir ?? ''), ARCHIVE_GLOBAL_DIR, ARCHIVE_GLOBAL_MARKER);
}

/**
 * 读全局命名空间标记。
 *
 * 标记只记录「这个 archive 是否已经按 global 命名空间写」以及当时的旧档归属，
 * 不记录任何路径、正文或文件名。标记存在以后，即使某个工作区的配置从「与全局
 * 共用」改成默认目录，全局层也不会突然切回 archive 根目录。
 */
export function readGlobalNamespaceMarker(archiveDir) {
    const file = globalNamespaceMarkerPath(archiveDir);
    if (!existsSync(file)) return null;
    let raw;
    try {
        raw = JSON.parse(readFileSync(file, 'utf8'));
    } catch {
        return { present: true, valid: false, legacy: 'unattributed', shared: true, reason: '', error: '全局命名空间标记不是合法 JSON' };
    }
    const shared = raw?.shared === true;
    const declared = asText(raw?.legacy);
    const legacy = shared ? 'unattributed' : (declared === 'global' ? 'global' : 'unattributed');
    return {
        present: true,
        valid: declared === 'global' || declared === 'unattributed',
        legacy,
        shared,
        reason: asText(raw?.reason),
        error: '',
    };
}

/**
 * 读旧 archive 的显式归属标记。
 *
 * 形状（手工放置，模块不会写它）：
 *   { "version": 1, "workspace": { "root": "<workspace-root>" } }
 * 也接受 workspace.identity / 顶层 identity（规范化后的身份串）。只有能归一成
 * 本工作区身份时才认领；标记坏了按「没有标记」处理，绝不据此猜归属。
 */
export function legacyOwnerOf(archiveDir) {
    const file = join(String(archiveDir ?? ''), ARCHIVE_OWNER_FILE);
    if (!existsSync(file)) return null;
    let raw;
    try {
        raw = JSON.parse(readFileSync(file, 'utf8'));
    } catch {
        return { present: true, valid: false, identity: '', root: '', error: '归属标记不是合法 JSON' };
    }
    const declaredRoot = asText(raw?.workspace?.root) || asText(raw?.root);
    const declaredIdentity = asText(raw?.workspace?.identity) || asText(raw?.identity);
    let identity = '';
    if (declaredRoot !== '') identity = identityOf(canonicalizePath(declaredRoot).path);
    else if (declaredIdentity !== '') identity = declaredIdentity;
    if (identity === '') return { present: true, valid: false, identity: '', root: declaredRoot, error: '归属标记里没有可用的工作区路径' };
    return { present: true, valid: true, identity, root: declaredRoot, error: '' };
}

/** 配置里的 memory.dir 是绝对路径时，它也可能与全局 archive 重合。 */
function configuredWorkspaceArchiveDirs(memory) {
    const configured = asText(memory?.dir);
    if (configured === '' || !isAbsolute(configured)) return [];
    return [archiveDirOf(configured)];
}

function classifyGlobal({ memory, memoryDir, otherArchiveDirs }) {
    const archiveDir = archiveDirOf(memoryDir);
    const marker = readGlobalNamespaceMarker(archiveDir);
    const owner = legacyOwnerOf(archiveDir);
    const ownerMarkerPresent = owner?.present === true;
    const namespacePresent = marker !== null || hasAnyEntry(join(archiveDir, ARCHIVE_GLOBAL_DIR));
    const workspaceNamespacesPresent = hasAnyEntry(join(archiveDir, ARCHIVE_WORKSPACES_DIR));
    const dirs = new Set([...(otherArchiveDirs ?? []), ...configuredWorkspaceArchiveDirs(memory)]);
    // 归属标记是「这个 archive 其实是工作区归档」的可见证据：即使另一个工作区把自己的
    // globalDir 指到它，也不能把旧档当成真正的全局旧档读取或覆盖。
    const collides = [...dirs].some((dir) => samePhysicalPath(archiveDir, dir))
        || workspaceNamespacesPresent
        || ownerMarkerPresent;
    // 标记一旦存在就固定走 archive/global：即便某个工作区后来把 memory.dir 改回默认，
    // 全局层也不会突然切回 archive 根目录（那会让 A 和 B 看到不同的全局归档）。
    const namespaced = collides || namespacePresent;
    const hasLegacy = hasLegacyArchiveEntries(archiveDir);
    const sharedDir = collides || workspaceNamespacesPresent || marker?.shared === true;
    const markerCanClaimLegacy = marker?.valid === true
        && marker.legacy === 'global'
        && marker.shared !== true
        && !sharedDir;
    let legacyAttribution = 'none';
    if (namespaced) legacyAttribution = markerCanClaimLegacy ? 'global' : (hasLegacy ? 'unattributed' : 'none');
    else legacyAttribution = hasLegacy ? 'global' : 'none';
    return {
        mode: namespaced ? (sharedDir ? 'global-shared' : 'global-namespaced') : 'global',
        writeDir: namespaced ? join(archiveDir, ARCHIVE_GLOBAL_DIR) : archiveDir,
        legacyAttribution,
        sharedDir,
        retentionSafe: !sharedDir,
        reason: namespaced
            ? (ownerMarkerPresent
                ? '全局目录里有工作区归属标记：它实际是工作区 archive，旧档保持只读、不归全局；新全局归档写 archive/global。'
                : (sharedDir
                    ? '全局层与工作区共用同一个物理 archive：旧档归属不明，保持只读；新全局归档写 archive/global。'
                    : '全局归档已按命名空间隔离：旧档归属由显式标记决定，新归档沿用 archive/global。'))
            : '全局层按设计跨工作区共享：旧 archive 属于全局层，新归档沿用原路径。',
    };
}

function classifyWorkspace({ workspace, memoryDir, globalArchiveDir }) {
    const archiveDir = archiveDirOf(memoryDir);
    const owner = legacyOwnerOf(archiveDir);
    const hasLegacy = hasLegacyArchiveEntries(archiveDir);
    const globalCollision = samePhysicalPath(archiveDir, globalArchiveDir);

    if (!workspace.ok) {
        return {
            mode: 'unknown',
            writeDir: join(archiveDir, ARCHIVE_WORKSPACES_DIR, 'unknown'),
            legacyAttribution: hasLegacy ? 'unattributed' : 'none',
            sharedDir: true,
            retentionSafe: false,
            reason: '工作目录无法定位：旧档归属不明；拒绝把新归档写到一个没有身份的命名空间，也不读取未归属旧档。',
            owner,
        };
    }
    const inside = isPathInside(workspace.path, memoryDir);
    const sharedDir = !inside || globalCollision;
    if (owner?.present === true && owner.valid === true && owner.identity === workspace.identity) {
        return {
            mode: 'claimed',
            writeDir: archiveDir,
            legacyAttribution: hasLegacy ? 'own' : 'none',
            sharedDir,
            retentionSafe: !sharedDir,
            reason: `archive/${ARCHIVE_OWNER_FILE} 明确声明这个旧 archive 属于本工作区，沿用原路径。`,
            owner,
        };
    }
    if (globalCollision) {
        return {
            mode: 'collision',
            writeDir: join(archiveDir, ARCHIVE_WORKSPACES_DIR, workspace.key),
            legacyAttribution: hasLegacy ? 'unattributed' : 'none',
            sharedDir: true,
            retentionSafe: false,
            reason: '全局层与工作区指向同一物理 archive：旧档归属不明，保持只读；新归档按 global / workspaces 命名空间分隔。',
            owner,
        };
    }
    if (inside) {
        return {
            mode: 'workspace-owned',
            writeDir: archiveDir,
            legacyAttribution: hasLegacy ? 'own' : 'none',
            sharedDir: false,
            retentionSafe: true,
            reason: '记忆目录在真实工作区内：旧 archive 可明确归本工作区，沿用原路径。',
            owner,
        };
    }
    const otherOwner = owner?.present === true && owner.valid === true && owner.identity !== workspace.identity;
    return {
        mode: 'shared',
        writeDir: join(archiveDir, ARCHIVE_WORKSPACES_DIR, workspace.key),
        legacyAttribution: otherOwner ? 'other' : (hasLegacy ? 'unattributed' : 'none'),
        sharedDir: true,
        retentionSafe: false,
        reason: otherOwner
            ? '记忆目录在工作区外，且旧 archive 的归属标记写的是另一个工作区：旧档保持只读，新归档写本工作区命名空间。'
            : '记忆目录在工作区外：可能与其他工作区共用；旧 archive 没有可核实的归属，保持只读，不并入本工作区；新归档写 archive/workspaces/<key>。',
        owner,
    };
}

/**
 * 判定一个记忆根（工作区或全局）的归档落点。
 *
 * @param {object} options
 * @param {string} [options.root] 真实工作区根（store=workspace 时必需）
 * @param {object} [options.memory] 记忆配置（dir / globalDir）
 * @param {'workspace'|'global'} [options.store]
 * @param {string} [options.memoryDir] 直接给物理记忆目录（调用方已经解析过时用）
 * @param {string} [options.globalDir] 直接给全局物理目录
 * @param {string[]} [options.otherArchiveDirs] 已知会与它碰撞的 archive 物理目录
 * @returns {object} 只读判定；writeDir 是新写落点，legacyDir 是旧档原位
 */
export function inspectArchiveScope({ root = '', memory = {}, store = 'workspace', memoryDir = '', globalDir = '', otherArchiveDirs = [] } = {}) {
    const kind = store === 'global' ? 'global' : 'workspace';
    const workspace = kind === 'workspace' ? canonicalWorkspacePath(root) : null;
    const dir = asText(memoryDir) || (kind === 'global' ? resolveGlobalMemoryDir(memory) : resolveMemoryDir(root, memory));
    const archiveDir = archiveDirOf(dir);
    const globalArchiveDir = kind === 'workspace'
        ? archiveDirOf(asText(globalDir) || resolveGlobalMemoryDir(memory))
        : '';
    const dirs = new Set((otherArchiveDirs ?? []).map((item) => asText(item)).filter(Boolean));
    if (kind === 'workspace') dirs.add(globalArchiveDir);
    const classified = kind === 'global'
        ? classifyGlobal({ memory, memoryDir: dir, otherArchiveDirs: [...dirs] })
        : classifyWorkspace({ workspace, memoryDir: dir, globalArchiveDir });
    const writeNamespaced = classified.writeDir !== archiveDir;
    const canWrite = kind === 'global' || workspace.ok;
    // 显式 globalDir 指向一个没有归属证据的 raw archive：保留旧路径与旧档读取，
    // 但不在这里做封顶删除——它可能与某个工作区目录物理重合，只是调用方没把
    // 那个工作区传进来。默认全局路径不在此列，仍按原 archiveKeep 清理。
    const explicitGlobalDir = kind === 'global' && asText(memory?.globalDir) !== '';
    const ambiguousExplicitGlobal = explicitGlobalDir && classified.mode === 'global';
    const retentionSafe = classified.retentionSafe !== false && !ambiguousExplicitGlobal;
    return {
        ok: true,
        store: kind,
        workspace,
        root: workspace === null ? '' : workspace.path,
        memoryDir: dir,
        archiveDir,
        legacyDir: archiveDir,
        writeDir: classified.writeDir,
        writeNamespaced,
        canWrite,
        mode: classified.mode,
        legacyAttribution: classified.legacyAttribution,
        legacyReadable: classified.legacyAttribution === 'own' || classified.legacyAttribution === 'global',
        sharedDir: classified.sharedDir === true,
        explicitGlobalDir,
        retentionSafe,
        retentionReason: ambiguousExplicitGlobal ? 'explicit-global-dir-unverified' : '',
        reliability: workspace === null ? 'global' : workspace.reliability,
        reason: classified.reason,
        owner: classified.owner ?? null,
        scopeId: kind === 'global'
            ? 'global'
            : (workspace.ok ? `workspace:${workspace.identity}` : 'workspace:unknown'),
        workspaceKey: workspace?.key ?? '',
        shortKey: workspace?.shortKey ?? '',
        label: workspace?.label ?? '全局',
        display: workspace?.display ?? '全局',
    };
}

/** 把 scope 应用到 pathsForDir 形状的对象（纯函数，不碰磁盘）。 */
export function withArchiveScope(paths, scope) {
    return {
        ...paths,
        archive: scope.writeDir,
        index: join(scope.writeDir, 'index.json'),
        legacyArchive: scope.legacyDir,
        archiveScope: scope,
    };
}

/**
 * 知识库虚拟目录。
 *
 * owned / claimed / 独立全局层沿用旧路径 归档/YYYY-MM：这些是既有用户已经
 * 引用过的归档页，不能因为本轮隔离把所有旧 ID 都换掉。只有真正需要命名空间的
 * 共用目录才加 工作区/<label> [key] 或 全局 这一层。
 */
export function kbArchiveFolder(scope, month) {
    const value = /^\d{4}-\d{2}$/.test(String(month ?? '')) ? String(month) : 'unknown-month';
    if (scope.writeNamespaced !== true) return `归档/${value}`;
    if (scope.store === 'global') return `归档/全局/${value}`;
    if (scope.mode === 'unknown') return `归档/未知工作区/${value}`;
    return `归档/工作区/${sanitizeLabel(scope.label)} [${scope.shortKey}]/${value}`;
}

/** 知识库 docHash 的 scope 前缀：只有命名空间化时才必须含真实工作区。 */
export function archiveScopeKey(scope) {
    if (scope.store === 'global') {
        if (scope.writeNamespaced !== true) return 'global';
        return `global:${identityOf(canonicalizePath(scope.archiveDir).path)}`;
    }
    if (scope.mode === 'unknown') return 'workspace:unknown';
    return `workspace:${scope.workspace?.identity ?? scope.scopeId}`;
}

/**
 * 归档条目在知识库里的稳定 docHash。
 *
 * 兼容优先：没有命名空间化的根（默认工作区、显式 owner 认领、独立全局层）继续用
 * 旧的 store.id + itemKey 算法，已有 kb:<hash> 文档 id 与引用都不变；共用/碰撞
 * 场景才改用带真实工作区身份的 scope。全局命名空间也刻意与旧的 global:key 不同，
 * 避免旧的未归属 global 页被悄悄解释成新命名空间里的另一条。
 */
export function archiveDocHash(scope, key) {
    const text = asText(key);
    if (scope.writeNamespaced !== true) return `a${sha1(`${scope.store}:${text}`).slice(0, SHA1_CHARS)}`;
    return `a${sha1(`${archiveScopeKey(scope)}:${text}`).slice(0, SHA1_CHARS)}`;
}

/**
 * 一次检查多个工作区，找出真正的物理碰撞。
 *
 * 返回的 collisions 只描述**可核实**的碰撞：
 *   shared-archive      两个不同身份的工作区解析到同一 archive 物理目录；
 *   duplicate-workspace 同一处物理路径的别名（应视为同一个工作区，不是碰撞）；
 *   same-label          显示名相同但身份不同（提示用，不是物理碰撞）；
 *   global-workspace    全局 archive 与某个工作区 archive 重合。
 */
export function detectArchiveCollisions({ roots = [], memory = {}, globalDir = '' } = {}) {
    const workspaces = [];
    for (const root of roots) workspaces.push(inspectArchiveScope({ root, memory, store: 'workspace', globalDir }));
    const groups = new Map();
    for (const scope of workspaces) {
        const key = identityOf(canonicalizePath(scope.archiveDir).path);
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(scope);
    }
    const collisions = [];
    for (const [key, group] of groups) {
        const identities = [...new Set(group.map((scope) => scope.workspace?.identity ?? ''))];
        if (identities.length > 1) {
            collisions.push({
                kind: 'shared-archive',
                archiveDir: group[0].archiveDir,
                archiveKey: key,
                workspaceKeys: group.map((scope) => scope.workspaceKey),
                roots: group.map((scope) => scope.root),
                legacyAttribution: 'unattributed',
                hint: '这些工作区共用同一个物理 archive；旧 archive 归属不明，各工作区新写落到自己的 namespaces/<key>。',
            });
        }
    }
    const byIdentity = new Map();
    for (const scope of workspaces) {
        if (!scope.workspace?.identity) continue;
        if (!byIdentity.has(scope.workspace.identity)) byIdentity.set(scope.workspace.identity, []);
        byIdentity.get(scope.workspace.identity).push(scope);
    }
    for (const [, group] of byIdentity) {
        if (group.length > 1) {
            collisions.push({
                kind: 'duplicate-workspace',
                workspaceKey: group[0].workspaceKey,
                roots: group.map((scope) => scope.root),
                hint: '这些写法指向同一处物理工作区，应按同一个工作区处理。',
            });
        }
    }
    const byLabel = new Map();
    for (const scope of workspaces) {
        const label = scope.workspace?.label ?? '';
        if (!byLabel.has(label)) byLabel.set(label, []);
        byLabel.get(label).push(scope);
    }
    for (const [label, group] of byLabel) {
        if (group.length > 1 && new Set(group.map((scope) => scope.workspaceKey)).size > 1) {
            collisions.push({
                kind: 'same-label',
                label,
                workspaceKeys: group.map((scope) => scope.workspaceKey),
                roots: group.map((scope) => scope.root),
                hint: '显示名相同但物理路径不同；知识库目录用 [shortKey] 区分，不会合并。',
            });
        }
    }
    const globalScope = inspectArchiveScope({ store: 'global', memory, globalDir, otherArchiveDirs: workspaces.map((scope) => scope.archiveDir) });
    for (const scope of workspaces) {
        if (samePhysicalPath(scope.archiveDir, globalScope.archiveDir)) {
            collisions.push({
                kind: 'global-workspace',
                archiveDir: scope.archiveDir,
                workspaceKey: scope.workspaceKey,
                hint: '全局层与这个工作区共用 archive；旧档归属不明，新写按 global / workspaces 分开。',
            });
        }
    }
    return { ok: true, workspaces, global: globalScope, collisions };
}
