/**
 * 统一缓存目录：所有中间产物（预览、临时导出、渲染出来的页面图）都只写到这里。
 *
 * 办公模式的一条硬规则：工作目录里不出现任何临时文件。需要落地的东西进
 * `.office/cache/`，交付物才写工作目录。
 *
 * 2026-09-23 改的是**生命周期**：以前每次调用开始清空、结束再清空，缓存等于
 * 一次性的 —— 于是任何跨调用复用的东西（最典型的是把 PDF 页面渲染成图片，
 * 下一步用 read_image 读）在第一次调用结束时就没了，第二次调用开始又被清一遍，
 * 缓存永远不可能命中。现在改成：
 *   - 调用开始只按 TTL 清理**过期**条目，不动新文件；
 *   - 调用结束默认保留，`keepCache:false` 才整目录清空；
 *   - 同一份输入重复渲染时按内容键命中已有文件（cache hit），不再重算。
 *
 * @module dsh-office-mode/engine/cache
 */
import { existsSync, mkdirSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { dirname, isAbsolute, join, resolve } from 'node:path';
import { displayPath } from './kit.js';

export const DEFAULT_CACHE_DIR = '.office/cache';

/** 把一段名字收敛成单层文件名，禁止子目录与向上跳。 */
function safeName(name) {
    const text = String(name ?? 'tmp').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-').replace(/^\.+/, '');
    return text === '' ? 'tmp' : text.slice(0, 120);
}

/**
 * 缓存内的相对路径 → 一段段收敛过的安全相对路径。
 *
 * 允许 `pdf/abc/page-1.png` 这样的子目录（渲染结果要按内容键分目录放），
 * 但不允许绝对路径与 `..`：缓存目录之外的任何东西都不该被缓存 API 碰到。
 */
function safeRelative(name) {
    const raw = String(name ?? 'tmp').split(/[\\/]+/);
    const parts = [];
    for (const part of raw) {
        if (part === '' || part === '.') continue;
        if (part === '..') throw new Error(`缓存路径不能跳出缓存目录：${String(name)}`);
        parts.push(safeName(part));
    }
    return parts.length === 0 ? 'tmp' : parts.join('/');
}

export function createCache(options = {}) {
    const root = resolve(options.root ?? process.cwd());
    const requested = options.dir ?? DEFAULT_CACHE_DIR;
    const absolute = isAbsolute(requested) ? resolve(requested) : resolve(root, requested);
    const rel = displayPath(root, absolute);

    /** 命中统计：跨调用复用的次数与省下的字节数（渲染这类贵操作会记账）。 */
    const stats = { hits: 0, hitBytes: 0, artifacts: 0 };

    /** 递归列出缓存内容（含子目录里的文件），相对缓存目录的路径当名字。 */
    function walk(dir, prefix = '', out = []) {
        let entries;
        try {
            entries = readdirSync(dir, { withFileTypes: true });
        } catch {
            return out;
        }
        for (const entry of entries) {
            const full = join(dir, entry.name);
            const name = prefix === '' ? entry.name : `${prefix}/${entry.name}`;
            if (entry.isDirectory()) {
                out.push({ name, dir: true, bytes: 0, modifiedMs: 0 });
                walk(full, name, out);
                continue;
            }
            let bytes = 0;
            let modifiedMs = 0;
            try {
                const info = statSync(full);
                bytes = info.size;
                modifiedMs = info.mtimeMs;
            } catch {
                bytes = 0;
            }
            out.push({ name, dir: false, bytes, modifiedMs });
        }
        return out;
    }

    const cache = {
        dir: absolute,
        rel,
        stats,
        ensure() {
            mkdirSync(absolute, { recursive: true });
            return absolute;
        },
        /** 缓存内文件名（可带子目录）→ 绝对路径。 */
        path(name) {
            const target = join(absolute, ...safeRelative(name).split('/'));
            mkdirSync(dirname(target), { recursive: true });
            return target;
        },
        /** 建一个缓存内的子目录，返回绝对路径。 */
        ensureDir(name) {
            const target = join(absolute, ...safeRelative(name).split('/'));
            mkdirSync(target, { recursive: true });
            return target;
        },
        /** 缓存内文件的绝对路径（不建目录，用于「有没有」这类判断）。 */
        locate(name) {
            return join(absolute, ...safeRelative(name).split('/'));
        },
        write(name, data) {
            const target = cache.path(name);
            // 非字符串内容按 JSON 落盘：脚本里写对象/数组是很自然的写法，
            // 让它直接可用比要求模型先 JSON.stringify 更省一次往返。
            const bytes = typeof data === 'string'
                ? Buffer.from(data, 'utf8')
                : (data instanceof ArrayBuffer || ArrayBuffer.isView(data))
                    ? Buffer.from(data)
                    : Buffer.from(JSON.stringify(data, null, 2), 'utf8');
            writeFileSync(target, bytes);
            stats.artifacts += 1;
            return { path: displayPath(root, target), absolute: target, bytes: bytes.length };
        },
        /** 记一次跨调用命中（复用已有产物，没重算）。 */
        noteHit(bytes = 0) {
            stats.hits += 1;
            stats.hitBytes += Number(bytes) || 0;
        },
        /** 记一次写进缓存的产物（脚本自己写的不算，渲染这类由模块调用）。 */
        noteArtifact() {
            stats.artifacts += 1;
        },
        /** 递归列出全部内容（含子目录）。 */
        list() {
            if (!existsSync(absolute)) return [];
            return walk(absolute);
        },
        /**
         * 按时间与容量清理缓存。
         *
         * @param {{maxAgeMs?: number, maxBytes?: number}} [limits]
         * @returns {{removed: string[], bytes: number, kept: number, keptBytes: number}}
         */
        prune(limits = {}) {
            const maxAgeMs = Number(limits.maxAgeMs) > 0 ? Number(limits.maxAgeMs) : 0;
            const maxBytes = Number(limits.maxBytes) > 0 ? Number(limits.maxBytes) : 0;
            if (maxAgeMs === 0 && maxBytes === 0) {
                const files = cache.list().filter((entry) => !entry.dir);
                return { removed: [], bytes: 0, kept: files.length, keptBytes: files.reduce((sum, f) => sum + f.bytes, 0) };
            }
            const now = Date.now();
            const files = cache.list().filter((entry) => !entry.dir);
            const removed = [];
            let removedBytes = 0;
            let alive = [];
            for (const file of files) {
                if (maxAgeMs > 0 && now - file.modifiedMs > maxAgeMs) {
                    if (cache.removeEntry(file.name)) {
                        removed.push(file.name);
                        removedBytes += file.bytes;
                    }
                    continue;
                }
                alive.push(file);
            }
            if (maxBytes > 0) {
                let total = alive.reduce((sum, file) => sum + file.bytes, 0);
                if (total > maxBytes) {
                    // 从最旧的开始删，直到回到容量以内。
                    for (const file of [...alive].sort((a, b) => a.modifiedMs - b.modifiedMs)) {
                        if (total <= maxBytes) break;
                        if (cache.removeEntry(file.name)) {
                            removed.push(file.name);
                            removedBytes += file.bytes;
                            total -= file.bytes;
                        }
                    }
                    alive = alive.filter((file) => !removed.includes(file.name));
                }
            }
            cache.pruneEmptyDirs();
            return {
                removed,
                bytes: removedBytes,
                kept: alive.length,
                keptBytes: alive.reduce((sum, file) => sum + file.bytes, 0),
            };
        },
        /** 删一个缓存内的文件（相对路径），删掉返回 true。 */
        removeEntry(name) {
            const target = cache.locate(name);
            if (!existsSync(target)) return false;
            try {
                rmSync(target, { recursive: true, force: true });
                return true;
            } catch {
                return false;
            }
        },
        /** 清掉已经空掉的子目录（顶层缓存目录本身保留）。 */
        pruneEmptyDirs() {
            if (!existsSync(absolute)) return;
            for (const entry of readdirSync(absolute, { withFileTypes: true })) {
                if (!entry.isDirectory()) continue;
                const full = join(absolute, entry.name);
                if (walk(full).length === 0) rmSync(full, { recursive: true, force: true });
            }
        },
        /** 缓存总体积（字节）。 */
        bytes() {
            return cache.list().filter((entry) => !entry.dir).reduce((sum, file) => sum + file.bytes, 0);
        },
        /** 清掉整个缓存目录。`keepCache:false` 与 office.cache.clear() 都会调用。 */
        clear() {
            if (!existsSync(absolute)) return false;
            rmSync(absolute, { recursive: true, force: true });
            return true;
        },
        /** 缓存目录相对工作目录的路径（供反馈展示）。 */
        display(name) {
            return displayPath(root, cache.locate(name));
        },
    };
    return cache;
}

export { safeRelative as safeCacheRelative };
