/**
 * 知识库的**显式空文件夹**（2026-11-20）。
 *
 * 为什么单独一个文件：知识库的目录树一直是从 manifest **投影**出来的 ——
 * 有文档的目录才存在。用户要的是普通资源管理器的手感：能先建一个空文件夹，
 * 再把文件放进去；建完刷新 / 重启它还在，进得去、面包屑回得来。
 *
 * 于是引入一份独立的目录清单 `<kbDir>/folders.json`：
 *   - 它只记**用户显式创建**的文件夹路径，不碰 manifest、块文件、SQLite 索引，
 *     也不在磁盘上真的 mkdir —— 知识库的逻辑目录不是文件系统目录；
 *   - 已有的文档推导树保持原样，两边按路径合并；显式目录只补出缺失的节点，
 *     不会改写任何由文档算出来的 docs / chunks / bytes 计数；
 *   - `office_memory` 的工具面没有「建文件夹」这个动作。这条写入只存在于
 *     面板的 POST 通道里（人点的），模型既看不到也调不到。
 *
 * 文件格式（版本化，但读取对旧格式宽容）：
 *   { "version": 1, "folders": ["参考资料", "项目/甲"] }
 * 数组形式也接受，方便以后手工修。
 *
 * 写入纪律（与 gallery 的目录索引同一套思路：临时文件 + 原子替换）：
 *   1. 读 → 改 → 写整段串行化：进程内按目录排队，跨进程用临时目录里的锁文件；
 *      两个窗口同时点「新建」不会互相覆盖（最后写入者不会拿旧清单盖掉新条目）。
 *   2. 临时文件用 `wx` 独占创建、0600 权限；rename 目标被占用时**有界重试**，
 *      绝不先删掉 live 的 folders.json —— 第二次 rename 失败时旧清单还在。
 *   3. 读到坏 JSON / 结构不对时，展示路径退回空数组（页面不红），但**写入路径
 *      拒绝继续**（KB_FOLDERS_CORRUPT）：绝不把损坏文件当成「没有目录」再覆盖掉。
 *
 * @module dsh-office-mode/kb-folders
 */

import { createHash, randomUUID } from 'node:crypto';
import { open, readFile, rename, stat, unlink } from 'node:fs/promises';
import { mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

/** 显式目录清单的文件名（落在知识库根目录里）。 */
export const KB_FOLDERS_FILE = 'folders.json';

/** 单个目录名的字符上限：够长，又不至于把面包屑撑爆。 */
export const KB_FOLDER_SEGMENT_MAX = 64;
/** 最多几层：知识库是给文档分类的，不是网盘。 */
export const KB_FOLDER_MAX_DEPTH = 8;
/** 显式目录条数上限：防止一份坏记录把面板拖垮。 */
export const KB_FOLDER_MAX_COUNT = 500;

/** rename 遇到占用时的重试次数 / 间隔（总计约 1.8s，够杀毒软件放开一次句柄）。 */
const RENAME_RETRY = 12;
const RENAME_RETRY_MS = 15;
/** 跨进程锁：拿不到时的等待上限 / 间隔。 */
const LOCK_RETRY = 100;
const LOCK_RETRY_MS = 20;
/** 锁文件超过这个年龄就当作进程崩溃留下的，允许接管。 */
const LOCK_STALE_MS = 30_000;

/** Windows 文件名里不允许出现的字符，外加控制字符（`/` 与 `\\` 单独处理）。 */
const FORBIDDEN_SEGMENT_CHARS = /[<>:"|?*\u0000-\u001f\u007f]/;

/** 路径校验失败：带一个稳定的 code，路由据此回 400 而不是笼统的「写入失败」。 */
function invalid(message) {
    const error = new Error(message);
    error.code = 'KB_FOLDER_INVALID';
    return error;
}

/** 清单文件本身坏了：读取展示可以退回空树，写入必须停下，不能覆盖旧内容。 */
function corrupt(message, cause) {
    const error = new Error(message);
    error.code = 'KB_FOLDERS_CORRUPT';
    if (cause !== undefined) error.cause = cause;
    return error;
}

function delay(ms) {
    return new Promise((resolveDelay) => {
        setTimeout(resolveDelay, ms);
    });
}

/**
 * 把用户/接口给的一条路径校验收成规范写法。
 *
 * 这里的「严格」是有意的：虚拟目录不进文件系统，但它是**用户看得见的地址**，
 * 也是面包屑与检索范围的键。放行 `a//b`、`a/../b`、`C:/x` 只会让同一层目录
 * 出现两种写法，之后按路径合并 / 收窄就都对不上。
 *
 * @param {unknown} value
 * @returns {string} 规范化后的相对路径（段之间用 `/`，没有首尾斜杠）
 * @throws {Error} 路径不合法时抛出可直接展示给用户的中文原因
 */
export function validateKbFolderPath(value) {
    const raw = typeof value === 'string' ? value.trim() : '';
    if (raw === '') throw invalid('文件夹名不能为空。');
    const unified = raw.replace(/\\/g, '/');
    if (unified.startsWith('/')) throw invalid('只能用知识库里的相对路径，不要以斜杠开头。');
    if (/^[A-Za-z]:/.test(unified)) throw invalid('不要用盘符路径：知识库文件夹是逻辑目录。');
    const parts = unified.split('/');
    if (parts.length > KB_FOLDER_MAX_DEPTH) {
        throw invalid('文件夹最多 ' + KB_FOLDER_MAX_DEPTH + ' 层。');
    }
    const out = [];
    for (const part of parts) {
        if (part === '') {
            throw invalid('路径里不能有空目录名（不要连续斜杠，也不要以斜杠结尾）。');
        }
        if (part === '.' || part === '..') {
            throw invalid('路径里不能有 . 或 ..。');
        }
        if (part !== part.trim()) {
            throw invalid('目录名的首尾不能有空格。');
        }
        if (part.length > KB_FOLDER_SEGMENT_MAX) {
            throw invalid('单个目录名不能超过 ' + KB_FOLDER_SEGMENT_MAX + ' 个字符。');
        }
        if (FORBIDDEN_SEGMENT_CHARS.test(part)) {
            throw invalid('目录名不能包含 < > : " | ? * 或控制字符。');
        }
        if (part.endsWith('.')) {
            throw invalid('目录名不要以点结尾。');
        }
        out.push(part);
    }
    return out.join('/');
}

/** 目录名的稳定排序：与 kbFolderTree 的子节点排序同一口径。 */
function compareFolderPath(left, right) {
    return String(left).localeCompare(String(right), 'zh-Hans-CN', { numeric: true });
}

/**
 * 读显式目录清单。
 *
 * 默认（展示路径）读失败退回空数组：这份文件只是界面元数据，坏掉不该让整页
 * 知识库变红。`strict: true` 供写入路径使用：坏 JSON / 坏结构 / 坏记录都抛
 * `KB_FOLDERS_CORRUPT`，避免把损坏文件当成空清单覆盖掉。
 *
 * 单条不合法的记录在宽容模式下会被跳过，而不是让整个文件作废 —— 手工改坏
 * 一行时其余目录仍然可用。
 *
 * @param {string} dir 知识库根目录
 * @param {{strict?: boolean}} [options]
 * @returns {Promise<string[]>} 规范化、去重、排好序的路径
 */
export async function readKbFolders(dir, options = {}) {
    const strict = options !== null && typeof options === 'object' && options.strict === true;
    const file = join(dir, KB_FOLDERS_FILE);
    let raw = '';
    try {
        raw = await readFile(file, 'utf8');
    } catch (error) {
        if (error && error.code === 'ENOENT') return [];
        if (strict) throw corrupt('读不到 ' + KB_FOLDERS_FILE + '，为免覆盖已有目录，先检查这个文件。', error);
        return [];
    }
    let parsed = null;
    try {
        parsed = JSON.parse(raw);
    } catch (error) {
        if (strict) throw corrupt(KB_FOLDERS_FILE + ' 不是合法 JSON；不会把它当空清单覆盖，请先修好。', error);
        return [];
    }
    const list = Array.isArray(parsed)
        ? parsed
        : (parsed && typeof parsed === 'object' && Array.isArray(parsed.folders) ? parsed.folders : null);
    if (list === null) {
        if (strict) throw corrupt(KB_FOLDERS_FILE + ' 结构不对：应为数组，或带 folders 数组的对象。');
        return [];
    }
    const seen = new Set();
    for (const item of list) {
        if (typeof item !== 'string') {
            if (strict) throw corrupt(KB_FOLDERS_FILE + ' 里有不是字符串的目录记录；不会覆盖，请先修好。');
            continue;
        }
        let path = '';
        try {
            path = validateKbFolderPath(item);
        } catch (error) {
            if (strict) throw corrupt(KB_FOLDERS_FILE + ' 里有不合法的目录记录：' + error.message, error);
            continue;
        }
        seen.add(path);
    }
    return [...seen].sort(compareFolderPath);
}

/** 进程内按知识库目录排队：同一时刻只有一段「读 → 改 → 写」在跑。 */
const FOLDER_QUEUES = new Map();

function serializedByDir(dir, task) {
    const key = resolve(dir);
    const previous = FOLDER_QUEUES.get(key) || Promise.resolve();
    const run = previous.catch(() => {}).then(task);
    const tail = run.catch(() => {});
    FOLDER_QUEUES.set(key, tail);
    tail.finally(() => {
        if (FOLDER_QUEUES.get(key) === tail) FOLDER_QUEUES.delete(key);
    });
    return run;
}

/** 跨进程锁放在系统临时目录里：不往知识库目录里丢锁文件。 */
function lockFileFor(dir) {
    const key = createHash('sha256').update(resolve(dir)).digest('hex').slice(0, 24);
    return join(tmpdir(), 'dsh-office-kb-folders-' + key + '.lock');
}

async function withFolderLock(dir, task) {
    const lockFile = lockFileFor(dir);
    let handle = null;
    for (let attempt = 0; attempt < LOCK_RETRY; attempt += 1) {
        try {
            handle = await open(lockFile, 'wx', 0o600);
            break;
        } catch (error) {
            if (!error || error.code !== 'EEXIST') throw error;
            // 上一个进程崩了会留下锁文件；超过 LOCK_STALE_MS 就当它失效，接管。
            try {
                const info = await stat(lockFile);
                if (Date.now() - info.mtimeMs > LOCK_STALE_MS) {
                    await unlink(lockFile).catch(() => {});
                    continue;
                }
            } catch {
                // 锁刚好被释放：下一轮直接重试。
            }
            await delay(LOCK_RETRY_MS);
        }
    }
    if (handle === null) {
        const error = new Error('知识库文件夹正在被另一个进程修改，稍后再试。');
        error.code = 'KB_FOLDERS_BUSY';
        throw error;
    }
    try {
        return await task();
    } finally {
        await handle.close().catch(() => {});
        await unlink(lockFile).catch(() => {});
    }
}

/**
 * 原子地写整份清单：先写同目录临时文件再 rename。
 *
 * 临时文件用 `wx` 独占创建（0600），避免撞上别的进程的临时文件；rename 目标
 * 被占用（Windows 上常见 EPERM / EACCES / EBUSY）时**有界重试**，绝不先删掉
 * live 的 `folders.json` —— 删了再 rename 失败就真把用户目录弄丢了。重试仍
 * 失败时抛错并清掉临时文件，旧清单原样保留。
 */
async function writeKbFolders(dir, folders) {
    await mkdir(dir, { recursive: true });
    const file = join(dir, KB_FOLDERS_FILE);
    const temp = join(dir, '.folders.json.' + process.pid + '.' + randomUUID() + '.tmp');
    const body = JSON.stringify({ version: 1, folders }, null, 2) + '\n';
    const handle = await open(temp, 'wx', 0o600);
    try {
        await handle.writeFile(body, 'utf8');
        await handle.sync().catch(() => {});
    } finally {
        await handle.close().catch(() => {});
    }
    try {
        for (let attempt = 0; ; attempt += 1) {
            try {
                await rename(temp, file);
                return;
            } catch (error) {
                const retriable = error && (error.code === 'EPERM' || error.code === 'EACCES'
                    || error.code === 'EBUSY' || error.code === 'EEXIST');
                if (!retriable || attempt >= RENAME_RETRY) throw error;
                await delay(RENAME_RETRY_MS);
            }
        }
    } finally {
        await unlink(temp).catch(() => {});
    }
}

/**
 * 建一个显式空文件夹（人从面板触发）。
 *
 * 幂等：同一条规范化路径已经在清单里就返回 `created:false`，不重复写文件、
 * 也不报错 —— 用户连点两次「创建」不应该看到红字。
 *
 * 并发：整段「严格读 → 判断 → 写」在进程内队列 + 跨进程锁里执行；两个请求
 * 同时建不同的目录时都会保留，不会后写者拿旧清单盖掉先写者。
 *
 * @param {string} dir 知识库根目录
 * @param {unknown} value 用户输入的路径（可带 `/` 建多层）
 * @returns {Promise<{action:string,path:string,created:boolean,total:number,folders:string[]}>}
 */
export async function createKbFolder(dir, value) {
    const path = validateKbFolderPath(value);
    return await serializedByDir(dir, async () => withFolderLock(dir, async () => {
        const folders = await readKbFolders(dir, { strict: true });
        if (folders.includes(path)) {
            return { action: 'kb-folder-create', path, created: false, total: folders.length, folders };
        }
        if (folders.length >= KB_FOLDER_MAX_COUNT) {
            throw invalid('显式文件夹太多（上限 ' + KB_FOLDER_MAX_COUNT + ' 个），先清理再建。');
        }
        const next = [...folders, path].sort(compareFolderPath);
        await writeKbFolders(dir, next);
        return { action: 'kb-folder-create', path, created: true, total: next.length, folders: next };
    }));
}

/**
 * 把显式文件夹合并进文档推导树。
 *
 * 规则只有一条：**只补缺，不改已有节点**。文档树里已有同名节点就沿用它
 * （它的 docs/chunks/bytes 是 manifest 算出来的事实），显式目录里缺的祖先
 * 节点补成空目录（docs/chunks/bytes 都是 0）。这样「用户新建的空目录」与
 * 「文档推导出来的目录」在界面上是同一种节点，面包屑、下钻、计数都不必特判。
 *
 * @param {object} root kbFolderTree() 的根节点
 * @param {string[]} folders 已规范化的显式路径
 * @returns {object} 同一个根（就地补节点并排序）
 */
export function mergeKbFolders(root, folders) {
    if (root === null || typeof root !== 'object') return root;
    const list = Array.isArray(folders) ? folders : [];
    const ensureChildren = (node) => {
        if (!Array.isArray(node.children)) node.children = [];
        return node.children;
    };
    for (const raw of list) {
        let path = '';
        try {
            path = validateKbFolderPath(raw);
        } catch {
            continue;
        }
        const parts = path.split('/');
        let node = root;
        let walked = '';
        for (const part of parts) {
            walked = walked === '' ? part : walked + '/' + part;
            const children = ensureChildren(node);
            let next = children.find((child) => child && child.path === walked);
            if (!next) {
                next = { path: walked, name: part, docs: 0, chunks: 0, bytes: 0, children: [], virtual: true };
                children.push(next);
            } else {
                ensureChildren(next);
            }
            node = next;
        }
    }
    const sortDeep = (node) => {
        if (!Array.isArray(node.children) || node.children.length === 0) return;
        node.children.sort((left, right) => compareFolderPath(left && left.name, right && right.name));
        for (const child of node.children) sortDeep(child);
    };
    sortDeep(root);
    return root;
}
