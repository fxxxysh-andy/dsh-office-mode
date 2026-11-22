/**
 * 会话级知识库范围（本轮新增）。
 *
 * 用户在聊天框左侧选一个知识库虚拟文件夹后，**这个会话**的所有 kb 读/写都只能落在
 * 该文件夹及其子目录里。这里存的是「范围的权威状态」，不是提示词：工具面（tools.js）
 * 每次调用都从宿主上下文现读，调用者传什么 folder 参数都只能往里收，不能往外放。
 *
 * 为什么状态放到宿主目录而不是项目 / 知识库目录：
 *   · 项目目录是会话工作区，模型能写；把范围放进去等于让被约束者保管自己的锁。
 *   · 知识库 manifest 也在这棵会被工具写的树里，同理不能当权威。
 *   · 状态根按 resolveScopeRoot 的优先级取（profileContext.home → DSH_HOME → ~/.dsh），
 *     每会话一个文件、按 session id 哈希命名、原子写 —— 并发与半截文件都不会让另一个
 *     会话读到别人的范围，也不会把范围读成「全库」。
 *
 * 继承：会话 header 里有 parentSession / isSeeded（fork）。派生会话（子代理 / fork）
 * 不物化范围快照，而是每次沿**真实 header 链**现读祖先记录再取交集 —— 父会话后来收紧，
 * 已经在跑的子代理跟着收紧；父文件损坏则整条链 fail-closed，绝不退回全库。
 *
 * @module dsh-office-mode/kb-session-scope
 */
import { createHash, randomBytes } from 'node:crypto';
import { chmod, mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { normalizeKbFolder } from './kb.js';

export const SCOPE_VERSION = 1;
/** 状态目录（相对状态根）。 */
export const SCOPE_DIR = '.office/kb-session-scopes';
/** 祖先链最多走这么多层；超了就拒绝，不猜。 */
export const MAX_ANCESTRY = 16;
export const SCOPE_MODES = ['folder', 'all'];

/** header 上的父会话 id；没有就是空串。 */
function parentIdOf(header) {
    return header !== null && header !== undefined && typeof header === 'object'
        && typeof header.parentSession === 'string'
        ? header.parentSession.trim()
        : '';
}

/**
 * 没有 parentSession 的派生会话不能当成链顶的根。
 *
 * 真实普通 UI fork 的 header 是 `isSeeded:true` 却没有 parentSession（来源记录不在 header
 * 上）—— 把它当根来读自己的记录，等于让一个来源不可核对的会话自己决定范围；子代理同理，
 * 它应当继承父范围，而不是从自己的记录起算。两种情况统一 fail-closed。
 */
function derivedWithoutParent(header) {
    if (header === null || header === undefined || typeof header !== 'object') return false;
    if (parentIdOf(header) !== '') return false;
    if (header.isSeeded === true) return true;
    return asText(header.origin) === 'subagent';
}

/** 只接受非空字符串；其它一律当空。 */
function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

function sha256(text) {
    return createHash('sha256').update(text, 'utf8').digest('hex');
}

/**
 * 范围状态的根目录。
 *
 * 优先级与宿主自己的 DSH_HOME 解析一致：dsh 启动的 profile 用 profileContext.home；
 * 否则看 DSH_HOME（空白视为未设）；再退到 ~/.dsh。不读 cwd —— 工作目录是会话自己的，
 * 不能拿它当「可信状态根」。
 */
export function resolveScopeRoot(ctx, env = process.env) {
    let home = '';
    try {
        const profile = ctx !== null && ctx !== undefined && typeof ctx.get === 'function'
            ? ctx.get('profileContext')
            : null;
        if (profile !== null && profile !== undefined && typeof profile.home === 'string') home = profile.home.trim();
    } catch {
        home = '';
    }
    if (home === '') home = typeof env?.DSH_HOME === 'string' ? env.DSH_HOME.trim() : '';
    if (home === '') home = join(homedir(), '.dsh');
    return join(home, SCOPE_DIR);
}

/** 每个会话一个文件；文件名取 session id 的哈希，不在文件名里泄露会话 id 原文。 */
export function scopeFileOf(root, sessionId) {
    return join(root, sha256(asText(sessionId)).slice(0, 32) + '.json');
}

function badRecord(detail) {
    return { status: 'corrupt', reason: 'record', detail };
}

/**
 * 校验并规范化一条记录。任何字段对不上都算损坏 —— 损坏要能被上层看见（fail-closed），
 * 而不是悄悄当成「没设置」或「全部」。
 */
function normalizeRecord(raw, sessionId) {
    if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return badRecord('shape');
    if (raw.v !== SCOPE_VERSION) return badRecord('version');
    if (asText(raw.session) !== asText(sessionId)) return badRecord('session');
    if (!SCOPE_MODES.includes(raw.mode)) return badRecord('mode');
    if (raw.mode === 'all') {
        return {
            status: 'ok',
            record: {
                v: SCOPE_VERSION,
                session: asText(sessionId),
                mode: 'all',
                folder: '',
                updatedAt: typeof raw.updatedAt === 'string' ? raw.updatedAt : '',
                by: typeof raw.by === 'string' ? raw.by : '',
            },
        };
    }
    let folder;
    try {
        folder = normalizeKbFolder(raw.folder);
    } catch {
        return badRecord('folder');
    }
    // 「folder 范围」不允许是根：根 = 全部文档，那应该用 mode:'all' 表达。
    if (folder === '') return badRecord('folder-empty');
    return {
        status: 'ok',
        record: {
            v: SCOPE_VERSION,
            session: asText(sessionId),
            mode: 'folder',
            folder,
            updatedAt: typeof raw.updatedAt === 'string' ? raw.updatedAt : '',
            by: typeof raw.by === 'string' ? raw.by : '',
        },
    };
}

/**
 * 创建一个范围状态库（每次调用现读盘；写用临时文件 + rename 原子提交）。
 *
 * 同一进程里重复创建同一个根没有意义，tools / 命令各自调用 scopeStoreFor 复用即可。
 */
export function createScopeStore({ root }) {
    const dir = root;
    return {
        root: dir,
        fileOf: (sessionId) => scopeFileOf(dir, sessionId),
        /**
         * 读一个会话的记录。
         * @returns {Promise<{status:'ok'|'unset'|'corrupt'|'error', record?:object, reason?:string, detail?:string}>}
         */
        async read(sessionId) {
            const id = asText(sessionId);
            if (id === '') return { status: 'unset', reason: 'no-session' };
            let text;
            try {
                text = await readFile(scopeFileOf(dir, id), 'utf8');
            } catch (error) {
                if (error?.code === 'ENOENT') return { status: 'unset' };
                // 权限 / IO 错误与「没设置」必须分开：前者 fail-closed，后者交给上层决定
                // 是拒绝还是按旧契约（内部调用）放行。
                return { status: 'error', reason: 'read', detail: String(error?.code ?? error) };
            }
            let parsed;
            try {
                parsed = JSON.parse(text);
            } catch {
                return { status: 'corrupt', reason: 'json' };
            }
            return normalizeRecord(parsed, id);
        },
        /** 写一个会话的记录（已存在的会被替换）。 */
        async write(sessionId, { mode, folder = '', by = 'command' } = {}) {
            const id = asText(sessionId);
            if (id === '') throw new Error('没有会话 id，不能设置知识库范围。');
            if (!SCOPE_MODES.includes(mode)) throw new Error('知识库范围的 mode 只能是 folder 或 all。');
            const wantedFolder = mode === 'folder' ? normalizeKbFolder(folder) : '';
            if (mode === 'folder' && wantedFolder === '') {
                throw new Error('「folder」范围需要一个具体的虚拟文件夹；要放开全部知识库请用 clear（mode:all）。');
            }
            const record = {
                v: SCOPE_VERSION,
                session: id,
                mode,
                folder: wantedFolder,
                updatedAt: new Date().toISOString(),
                by: asText(by) || 'command',
            };
            await mkdir(dir, { recursive: true, mode: 0o700 });
            const target = scopeFileOf(dir, id);
            const tmp = target + '.' + randomBytes(6).toString('hex') + '.tmp';
            try {
                await writeFile(tmp, JSON.stringify(record) + '\n', { mode: 0o600 });
                await rename(tmp, target);
            } catch (error) {
                await rm(tmp, { force: true }).catch(() => {});
                throw error;
            }
            // Windows 上 rename 继承临时文件权限的语义与 POSIX 不同，补一次 chmod（失败不致命）。
            try { await chmod(target, 0o600); } catch { /* 平台不支持就算了 */ }
            return record;
        },
    };
}

const storeCache = new Map();
/** 按状态根复用 scope store（根不变就不必每次重建）。 */
export function scopeStoreFor(ctx) {
    const root = resolveScopeRoot(ctx);
    if (!storeCache.has(root)) storeCache.set(root, createScopeStore({ root }));
    return storeCache.get(root);
}

/**
 * 两个文件夹范围取交集；无交集返回 null。
 * 空串表示根（全部），与任何范围的交集都是另一个。
 */
export function intersectFolders(left, right) {
    const a = normalizeKbFolder(left);
    const b = normalizeKbFolder(right);
    if (a === '' || b === '') return a === '' ? b : a;
    if (a === b) return a;
    if (a.startsWith(b + '/')) return a;
    if (b.startsWith(a + '/')) return b;
    return null;
}

/**
 * kb-tree 的真实返回是 { action, root, counts, origin }，树在 `root` 下；这里同时容忍
 * 裸节点 / 节点数组，供命令的 folderExists 与测试复用，避免再把包装当成树（恒 false）。
 * 只做 **exact** 命中：调用者已经把 folder 规范化，父子边界不在这里判断。
 */
export function kbTreeHasFolder(tree, folder) {
    const wanted = normalizeKbFolder(folder);
    if (wanted === '') return true;
    const value = tree !== null && tree !== undefined && typeof tree === 'object' && !Array.isArray(tree)
        && tree.root !== undefined
        ? tree.root
        : tree;
    const roots = Array.isArray(value) ? value : [value];
    const seen = (node) => {
        if (node === null || node === undefined || typeof node !== 'object') return false;
        let path;
        try {
            path = normalizeKbFolder(node.path);
        } catch {
            return false;
        }
        if (path === wanted) return true;
        const children = Array.isArray(node.children) ? node.children : [];
        return children.some(seen);
    };
    return roots.some(seen);
}

/**
 * 目录存在性三态：true（确认存在）/ false（确认不在树里）/ null（树读不全，不能下结论）。
 *
 * folders.json 损坏时 kbTree 只保留 manifest 部分并带 foldersError；此时查不到可能是
 * 因为那份清单没读出来，不能回「不存在」。调用方拿到 null 应让用户重试。
 */
export function kbTreeFolderPresence(tree, folder) {
    if (kbTreeHasFolder(tree, folder)) return true;
    if (tree !== null && typeof tree === 'object' && typeof tree.foldersError === 'string' && tree.foldersError !== '') {
        return null;
    }
    return false;
}

/** 单个文件夹是否落在范围里（exact 或 descendant；空范围 = 根 = 全部）。 */
export function folderWithin(folder, scopeFolder) {
    const value = normalizeKbFolder(folder);
    const scope = normalizeKbFolder(scopeFolder);
    if (scope === '') return true;
    return value === scope || value.startsWith(scope + '/');
}

/** 沿真实 header 链解析一个会话的**生效范围**。 */
export async function resolveSessionScope({ sessionId, header, store, headerOf, maxDepth = MAX_ANCESTRY }) {
    const id = asText(sessionId);
    if (id === '') return { kind: 'no-session' };
    const chain = [];
    const seen = new Set();
    let current = { id, header: header ?? null };
    while (current !== null) {
        if (seen.has(current.id)) return { kind: 'error', code: 'cycle', detail: current.id };
        if (chain.length >= maxDepth) return { kind: 'error', code: 'depth', detail: id };
        seen.add(current.id);
        chain.push(current);
        const parent = parentIdOf(current.header);
        if (parent === '') {
            // 链顶必须是「没有父会话的普通根会话」；seeded fork / 子代理缺父记录时不能当根。
            if (derivedWithoutParent(current.header)) return { kind: 'error', code: 'origin', detail: current.id };
            break;
        }
        let parentHeader;
        try {
            parentHeader = await headerOf(parent);
        } catch {
            parentHeader = undefined;
        }
        // 祖先 header 拿不到（被删 / 存储不可用）时不能假定它是根：拒绝，而不是放大范围。
        if (parentHeader === null || parentHeader === undefined) {
            return { kind: 'error', code: 'ancestry', detail: parent };
        }
        current = { id: parent, header: parentHeader };
    }
    const records = [];
    for (const item of chain) {
        const found = await store.read(item.id);
        if (found.status === 'error') return { kind: 'error', code: 'state', detail: item.id };
        if (found.status === 'corrupt') return { kind: 'corrupt', detail: item.id };
        records.push({ id: item.id, record: found.status === 'ok' ? found.record : null });
    }
    const root = records[records.length - 1];
    // 链顶没设置：整条链都是「未设置」。绝不能在这里假定根会话是全库 —— 一个没设置过
    // 范围的会话只是「还没选」，不是「选了全部」。
    if (root.record === null) return { kind: 'unset', detail: root.id };
    let mode = root.record.mode;
    let folder = mode === 'folder' ? root.record.folder : '';
    // 后代记录只能往里收：folder 取交集，all 不放大（零记录 = 继承父范围）。
    for (const entry of records.slice(0, -1)) {
        const record = entry.record;
        if (record === null || record.mode === 'all') continue;
        if (mode === 'all') {
            mode = 'folder';
            folder = record.folder;
            continue;
        }
        const narrowed = intersectFolders(folder, record.folder);
        if (narrowed === null) return { kind: 'empty', detail: entry.id };
        folder = narrowed;
    }
    return mode === 'all'
        ? { kind: 'ok', scope: { mode: 'all', folder: '' } }
        : { kind: 'ok', scope: { mode: 'folder', folder } };
}

/** 从一次工具执行的上下文里取会话 id / header（与 projection.js 的 sessionIdOf 同一来源）。 */
export function sessionHeaderOf(exec) {
    const session = exec?.agent?.session ?? null;
    const id = typeof session?.id === 'string' ? session.id : (typeof session?.header?.id === 'string' ? session.header.id : '');
    const header = session?.header ?? null;
    return { sessionId: id.trim(), header };
}

/**
 * 为一次宿主调用解析生效范围。
 *
 * 返回值三种：
 *   · {ok:true, scope:null}   —— 没有会话（内部 / 纯测试路径），沿用旧契约（不受限）；
 *   · {ok:true, scope:{...}}  —— 会话有范围（可能 mode:'all'）；
 *   · {ok:false, text}        —— 未设置 / 状态损坏 / 祖先无法解析：**拒绝**，不退回全库。
 */
export async function resolveCallScope({ ctx, exec }) {
    // 没有宿主上下文（内部脚本 / 直接 buildTools(config) 的纯测试）读不到可信状态根：
    // 这条路径本来就不经过会话，按旧契约不受限。真会话的工具面一定绑了 ctx，
    // 调用者也拿不到参数来把它去掉 —— 所以这不是一条绕过通道。
    if (ctx === null || ctx === undefined || typeof ctx.get !== 'function') return { ok: true, scope: null };
    const { sessionId, header } = sessionHeaderOf(exec);
    // 有真实宿主、却拿不到可信会话标识 / header：不能当作内部调用放行 —— 否则
    // 「exec 上少填一个字段」就成了绕过会话范围的通道。拒绝，且不猜成「未设置」。
    if (sessionId === '') return { ok: false, text: scopeRefusalText({ kind: 'no-session' }) };
    if (header === null || header === undefined || typeof header !== 'object' || Array.isArray(header)) {
        return { ok: false, text: scopeRefusalText({ kind: 'error', code: 'header', detail: sessionId }) };
    }
    const headerId = asText(header.id);
    if (headerId !== '' && headerId !== sessionId) {
        return { ok: false, text: scopeRefusalText({ kind: 'error', code: 'header', detail: sessionId }) };
    }
    const store = scopeStoreFor(ctx);
    const result = await resolveSessionScope({
        sessionId,
        header,
        store,
        headerOf: (id) => headerOfSession(ctx, id),
    });
    if (result.kind === 'ok') return { ok: true, scope: result.scope };
    return { ok: false, text: scopeRefusalText(result) };
}

/** 取某个会话的 header：先看活着的 agent，再问持久层（离线会话 / fork 祖先）。 */
export async function headerOfSession(ctx, sessionId) {
    try {
        const agents = typeof ctx?.get === 'function' ? ctx.get('agents') : null;
        const agent = agents !== null && agents !== undefined && typeof agents.get === 'function'
            ? agents.get(sessionId)
            : null;
        if (agent?.session?.header !== undefined && agent.session.header !== null) return agent.session.header;
    } catch {
        // 活体注册表不可用时退到持久层，不在这里下结论。
    }
    try {
        const persistence = typeof ctx?.get === 'function' ? ctx.get('sessionPersistence') : null;
        if (persistence !== null && persistence !== undefined && typeof persistence.stat === 'function') {
            const snapshot = await persistence.stat(sessionId);
            return snapshot?.header;
        }
    } catch {
        return undefined;
    }
    return undefined;
}

/**
 * 把解析失败翻译成一句能照做的拒绝（不泄露范围外的文档是否存在）。
 *
 * **UI 契约**：状态 / 拒绝文案统一以「知识库范围：」开头，紧跟一个固定 token
 * （全部知识库 / 「<folder>」/ 未设置 / 状态损坏 / 范围为空 / 父会话不可用 /
 * 来源不可确认 / 父链循环 / 父链过深 / 状态读取失败 / 无会话 / 无法确认），
 * 之后才是给人看的解释句。宿主的 CommandResult 只保留 kind/text，UI 只能从 text 认状态，
 * 所以 token 必须稳定；加新状态时同步 flash-scope-review 的解析清单。
 */
export function scopeRefusalText(result) {
    const at = asText(result?.detail);
    switch (result?.kind) {
        case 'no-session':
            return '知识库范围：无会话。当前调用没有会话上下文，读不到知识库范围。请在聊天框里操作。';
        case 'unset':
            return '知识库范围：未设置。这个会话还没有设置知识库范围，所以不能查知识库（不会默认放开全库）。'
                + '在输入框左侧选一个虚拟文件夹，或输入 /kb-scope set <虚拟文件夹>；要查全部用 /kb-scope clear。';
        case 'corrupt':
            return '知识库范围：状态损坏。这个会话的知识库范围状态文件读不懂（可能被外部改过）。为安全起见这次拒绝查询，不会退回全库。'
                + '请重新选择目录，或让管理员检查 ' + SCOPE_DIR + ' 下对应的文件。'
                + (at === '' ? '' : '（会话：' + at + '）');
        case 'empty':
            return '知识库范围：范围为空。这个会话继承到的范围与它自己的记录没有交集，所以能查的目录为空。'
                + '请在父会话里放宽范围，或在当前会话重新选择目录。'
                + (at === '' ? '' : '（会话：' + at + '）');
        case 'error':
            if (result.code === 'ancestry') {
                return '知识库范围：父会话不可用。这个会话的父会话记录已不可用，无法确认知识库范围，这次拒绝查询（fail-closed）。'
                    + (at === '' ? '' : '父会话：' + at);
            }
            if (result.code === 'origin') {
                return '知识库范围：来源不可确认。这个会话没有可核对的父会话记录（seeded fork / 没有父链的子代理），'
                    + '无法确认它继承的范围，这次拒绝访问（fail-closed）。'
                    + (at === '' ? '' : '（会话：' + at + '）');
            }
            if (result.code === 'header') {
                return '知识库范围：无法确认。这个调用拿不出可信的会话信息，拒绝访问知识库（fail-closed），不会退回全库。';
            }
            if (result.code === 'cycle') {
                return '知识库范围：父链循环。会话的父链出现循环，无法确认知识库范围，这次拒绝查询（fail-closed）。';
            }
            if (result.code === 'depth') {
                return '知识库范围：父链过深。会话的父链太深（超过 ' + MAX_ANCESTRY + ' 层），无法确认知识库范围，这次拒绝查询（fail-closed）。';
            }
            return '知识库范围：状态读取失败。读知识库范围状态失败，这次拒绝查询（fail-closed），不会退回全库。';
        default:
            return '知识库范围：无法确认。无法确认知识库范围，这次拒绝查询。';
    }
}

/**
 * 注册 /kb-scope 命令：查看 / 设置 / 清除当前会话的知识库范围。
 *
 * 只用 authenticated 的命令通道改状态：命令由宿主按「发命令的那个 agent」执行，
 * 模型改不了自己的会话 id，也拿不到别的会话的范围。子代理拒绝写 —— 它们继承父范围，
 * 允许它们自写就等于给了一条绕开父约束的路。
 */
export function registerKbScopeCommand(ctx, { logger = null, folderExists = null } = {}) {
    const commands = ctx?.commands ?? (typeof ctx?.get === 'function' ? ctx.get('commands') : null);
    if (commands === null || commands === undefined || typeof commands.register !== 'function') return null;
    return commands.register({
        name: 'kb-scope',
        description: '查看或设置当前会话的知识库范围：只检索所选虚拟文件夹及其子目录（clear = 全部知识库）。',
        input: { hint: 'set <虚拟文件夹> | clear' },
        handler: async (invocation) => {
            try {
                const agent = invocation?.agent ?? null;
                const session = agent?.session ?? null;
                const sessionId = typeof session?.id === 'string' ? session.id
                    : (typeof session?.header?.id === 'string' ? session.header.id : '');
                if (sessionId === '') {
                    return { kind: 'error', text: '知识库范围：无会话。没有会话上下文，不能设置知识库范围。' };
                }
                const header = session?.header ?? null;
                // 写状态必须建立在可信的会话身份上：缺 header / header 的 id 与 session 不符时
                // 拒绝，而不是按一个「看起来像」的 id 写文件。status 查询也走同一条门 ——
                // 读不出可信身份时给「无法确认」，不给任何可被误读成范围的状态。
                if (header === null || header === undefined || typeof header !== 'object' || Array.isArray(header)) {
                    return { kind: 'error', text: scopeRefusalText({ kind: 'error', code: 'header', detail: sessionId }) };
                }
                const headerId = asText(header.id);
                if (headerId !== '' && headerId !== sessionId) {
                    return { kind: 'error', text: scopeRefusalText({ kind: 'error', code: 'header', detail: sessionId }) };
                }
                if (header.origin === 'subagent') {
                    return { kind: 'error', text: '知识库范围：子代理只读。子代理继承父会话的知识库范围，不能单独设置。请在父会话里调整。' };
                }
                if (derivedWithoutParent(header)) {
                    return {
                        kind: 'error',
                        text: '知识库范围：来源不可确认。这个会话是 fork，但没有可核对的父会话记录，不能设置或清除知识库范围。请回到原会话调整。',
                    };
                }
                const store = scopeStoreFor(ctx);
                const raw = typeof invocation?.rawInput === 'string' ? invocation.rawInput.trim() : '';
                if (raw === '') return { kind: 'success', text: await statusLineFor(ctx, sessionId, header, store) };
                const space = raw.indexOf(' ');
                const verb = (space === -1 ? raw : raw.slice(0, space)).toLowerCase();
                const rest = space === -1 ? '' : raw.slice(space + 1).trim();
                if (verb !== 'clear' && verb !== 'all' && verb !== 'set') {
                    return { kind: 'error', text: '用法：/kb-scope（看当前范围）、/kb-scope set <虚拟文件夹>、/kb-scope clear。' };
                }
                if ((verb === 'clear' || verb === 'all') && rest !== '') {
                    return { kind: 'error', text: 'clear 不带参数：/kb-scope clear' };
                }
                // 写门槛只看**父会话**的生效范围，不看本会话自己的记录：本会话记录损坏 / 与父范围
                // 没有交集时，用户仍然可以重新选择来修复它。父会话未设置 / 损坏 / 父链不可用 /
                // 父范围为空时，本会话写什么都不会生效 —— 必须拒绝，不能回「已设为」的假成功。
                const parent = await parentScopeFor(ctx, header, store);
                if (parent.kind === 'unset') {
                    return {
                        kind: 'error',
                        text: '知识库范围：未设置。父会话还没有设置知识库范围，当前会话的记录不会生效。'
                            + '请先在父会话选择目录；父会话设好后再回到这里选择要收窄到的子目录。',
                    };
                }
                if (parent.kind !== 'ok' && parent.kind !== 'root') {
                    return { kind: 'error', text: scopeRefusalText(parent) };
                }
                const readback = async () => resolveSessionScope({
                    sessionId,
                    header,
                    store,
                    headerOf: (id) => headerOfSession(ctx, id),
                });
                if (verb === 'clear' || verb === 'all') {
                    await store.write(sessionId, { mode: 'all', folder: '', by: 'command' });
                    const after = await readback();
                    return { kind: 'success', text: clearScopeReceipt(after, parent) };
                }
                if (rest === '') return { kind: 'error', text: 'set 后面要给一个虚拟文件夹，例如 /kb-scope set 参考资料/合同。' };
                let folder;
                try {
                    folder = normalizeKbFolder(rest);
                } catch (error) {
                    return { kind: 'error', text: String(error?.message ?? error) };
                }
                if (folder === '') return { kind: 'error', text: '根目录不是文件夹：要放开全部知识库用 /kb-scope clear。' };
                // 子会话的选择必须落在父范围里。交集为空时先写再回「已设为」会让当前可用的
                // 范围立刻变成空，等于一次误点把正在用的选择弄坏；写前拒绝，原记录不动。
                if (parent.kind === 'ok' && parent.scope.mode === 'folder' && !folderWithin(folder, parent.scope.folder)) {
                    return {
                        kind: 'error',
                        text: '知识库范围：「' + folder + '」不在父会话的范围「' + parent.scope.folder + '」里，这次没有保存。'
                            + '请选择「' + parent.scope.folder + '」或其子目录。',
                    };
                }
                // 选一个库里不存在的文件夹只会得到一个永远空的检索范围（仍然 fail-closed），
                // 但用户会以为是「查不到」。核对失败（抛出 / 返回 null）不能当成「存在」：
                // 这次不保存，让用户重试；未注入核对函数（headless / 纯测试）时没有这条信息，照旧写入。
                if (typeof folderExists === 'function') {
                    let exists = null;
                    try {
                        exists = await folderExists({ id: sessionId, header }, folder);
                    } catch {
                        exists = null;
                    }
                    if (exists === false) {
                        return {
                            kind: 'error',
                            text: '知识库里没有「' + folder + '」这个虚拟文件夹。'
                                + '请从输入框左侧选择已有目录，或先把文档 kb-ingest 到该文件夹。',
                        };
                    }
                    if (exists !== true) {
                        return {
                            kind: 'error',
                            text: '知识库范围：暂时核对不了「' + folder + '」是否存在（读取知识库树失败）。'
                                + '这次没有保存，请稍后重试。',
                        };
                    }
                }
                await store.write(sessionId, { mode: 'folder', folder, by: 'command' });
                const after = await readback();
                return { kind: 'success', text: setScopeReceipt(folder, after) };
            } catch (error) {
                logger?.warn?.('[dsh-office-mode] /kb-scope 失败：' + String(error?.message ?? error));
                return { kind: 'error', text: '设置知识库范围失败：' + String(error?.message ?? error) };
            }
        },
    });
}

/**
 * 解析**父会话**的生效范围（不读本会话记录）。
 *
 * 写门槛必须只卡父链：本会话自己的记录损坏 / 为空 / 与父范围不交集，都可能由一次新的
 * 选择修好；如果拿「本会话的 effective」当门槛，这些情况会被自己的坏状态锁死。
 * 根会话返回 { kind:'root' }（没有父门槛）；父 header 拿不到返回 ancestry，父记录损坏 /
 * 父链未设置等按 resolveSessionScope 的 kind 原样返回。
 */
async function parentScopeFor(ctx, header, store) {
    const parentId = parentIdOf(header);
    if (parentId === '') return { kind: 'root' };
    let parentHeader;
    try {
        parentHeader = await headerOfSession(ctx, parentId);
    } catch {
        parentHeader = undefined;
    }
    if (parentHeader === null || parentHeader === undefined || typeof parentHeader !== 'object' || Array.isArray(parentHeader)) {
        return { kind: 'error', code: 'ancestry', detail: parentId };
    }
    return resolveSessionScope({
        sessionId: parentId,
        header: parentHeader,
        store,
        headerOf: (id) => headerOfSession(ctx, id),
    });
}

/** set 的回执：按写后回读的**生效**范围说话；读不到就明确说无法确认，不假定成功。 */
function setScopeReceipt(folder, after) {
    if (after.kind === 'ok' && after.scope.mode === 'folder') {
        if (after.scope.folder !== folder) {
            return '知识库范围：已设为「' + folder + '」（含子目录）；当前生效范围是「'
                + after.scope.folder + '」（受父会话范围影响）。';
        }
        return '知识库范围：已设为「' + folder + '」（含子目录）。';
    }
    if (after.kind !== 'ok') {
        return '知识库范围：已写入本会话记录，但当前生效范围无法确认，请用 /kb-scope 查看。';
    }
    return '知识库范围：已设为「' + folder + '」，但当前生效范围是全部知识库。';
}

/**
 * clear 的回执：派生会话清掉的只是**自己**的附加限制，父范围仍然生效。
 * 只有回读确实是 all 才说「已清空（全部知识库）」，否则报当前继承到的文件夹。
 */
function clearScopeReceipt(after, parent) {
    if (after.kind === 'ok' && after.scope.mode === 'folder') {
        return '知识库范围：已清除本会话的附加限制；当前生效范围是「'
            + after.scope.folder + '」（含子目录），继承自父会话。';
    }
    if (after.kind === 'ok') {
        return '知识库范围：已清空（全部知识库）。本会话不再限制文件夹。';
    }
    if (parent.kind === 'ok' && parent.scope.mode === 'folder') {
        return '知识库范围：已清除本会话的附加限制；当前生效范围是「'
            + parent.scope.folder + '」（含子目录），继承自父会话。';
    }
    if (parent.kind === 'root' || (parent.kind === 'ok' && parent.scope.mode === 'all')) {
        return '知识库范围：已清空（全部知识库）。本会话不再限制文件夹。';
    }
    return '知识库范围：已清除本会话的附加限制，但当前生效范围无法确认，请用 /kb-scope 查看。';
}

/** 命令不带参数时的状态行（只给一个文件夹名与状态，不列整棵树）。 */
async function statusLineFor(ctx, sessionId, header, store) {
    const result = await resolveSessionScope({
        sessionId,
        header,
        store,
        headerOf: (id) => headerOfSession(ctx, id),
    });
    if (result.kind === 'ok') {
        return result.scope.mode === 'all'
            ? '知识库范围：全部知识库。'
            : '知识库范围：「' + result.scope.folder + '」（含子目录）。';
    }
    return scopeRefusalText(result);
}
