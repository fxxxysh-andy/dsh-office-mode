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
 *   2. **只服务见过的根**。`?cwd=` 必须命中「本进程里真跑过工具的那个工作目录」，
 *      否则 404。没有这一条，一个本地 HTTP 端点就成了「传任意路径读任意目录」的
 *      数据外泄面。已知根由 noteWorkspace() 记录，来源是工具执行上下文里的 cwd。
 *   3. **不返回记忆目录之外的任何路径**。响应里只出现相对工作目录的展示路径。
 *   4. **有界**。台账 / 归档 / 关系 / 实体都截断到固定条数，端点不会被一份巨大的
 *      记忆库拖垮。
 *
 * @module dsh-office-mode/view
 */

import { resolve } from 'node:path';
import { createMemory, byteLength, ENTRY_DELIMITER } from './memory.js';

/** 端点的绝对路径（无尾斜杠）。 */
export const MEMORY_VIEW_PATH = '/office-memory/snapshot';

/** 一次响应里各层的条数上限（浏览用，不是全量导出）。 */
export const VIEW_LIMITS = {
    ledger: 200,
    archive: 200,
    links: 500,
    entities: 60,
};

/** 见过的记忆根：cwd（规范化绝对路径）→ 最后见到的时间。 */
const knownRoots = new Map();
/** 最多记多少个根，避免长时间运行后无限增长。 */
const MAX_ROOTS = 50;

/**
 * 记下一个「本进程真的在这里跑过工具」的工作目录。
 *
 * 只由工具执行路径调用 —— 这是路由信任边界的唯一来源，别的入口都不该调它。
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
    knownRoots.set(absolute, new Date().toISOString());
    if (knownRoots.size > MAX_ROOTS) {
        const oldest = knownRoots.keys().next().value;
        if (oldest !== undefined && oldest !== absolute) knownRoots.delete(oldest);
    }
}

/** 当前见过的根列表（最近见到的排前面）。 */
export function knownWorkspaces() {
    return [...knownRoots.entries()]
        .map(([cwd, seenAt]) => ({ cwd, seenAt }))
        .reverse();
}

/** 测试用：清空已知根。 */
export function resetWorkspaces() {
    knownRoots.clear();
}

/** 取一个展示用的短名字（路径最后一段）。 */
function labelOf(cwd) {
    const parts = String(cwd).split(/[\\/]/).filter((part) => part !== '');
    return parts.length === 0 ? String(cwd) : parts[parts.length - 1];
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
 */
export async function buildSnapshot({ root, memory = {}, query = '' } = {}) {
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
        label: labelOf(root),
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
 * @param {object} ctx 插件上下文（需要 ctx.webServer）
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
                if (memory.enabled === false) {
                    sendJson(res, 200, { ok: false, error: '记忆已在设置里关掉。', workspaces: knownWorkspaces() });
                    return;
                }
                const url = new URL(String(req.url ?? '/'), 'http://127.0.0.1');
                const requested = asText(url.searchParams.get('cwd'));
                const query = asText(url.searchParams.get('q'));

                // 信任边界：只服务本进程真的跑过工具的工作目录。
                let root = '';
                if (requested !== '') {
                    let absolute = '';
                    try {
                        absolute = resolve(requested);
                    } catch {
                        absolute = '';
                    }
                    if (absolute === '' || !knownRoots.has(absolute)) {
                        sendJson(res, 404, {
                            ok: false,
                            error: '这个工作目录不在已知列表里。请先在该工作目录的会话里用一次办公工具（例如 office_help），再回来刷新。',
                            workspaces: knownWorkspaces(),
                        });
                        return;
                    }
                    root = absolute;
                } else {
                    const recent = knownWorkspaces()[0];
                    root = recent?.cwd ?? '';
                }
                if (root === '') {
                    sendJson(res, 200, {
                        ok: false,
                        error: '还没有见过任何工作目录。先在某个会话里用一次办公工具（例如 office_help），再回来刷新。',
                        workspaces: [],
                    });
                    return;
                }
                const snapshot = await buildSnapshot({ root, memory, query });
                snapshot.workspaces = knownWorkspaces().map((item) => ({
                    cwd: item.cwd,
                    label: labelOf(item.cwd),
                    seenAt: item.seenAt,
                }));
                sendJson(res, 200, snapshot);
            } catch (error) {
                sendJson(res, 500, { ok: false, error: `读取记忆失败：${error?.message ?? error}` });
            }
        },
    });
}
