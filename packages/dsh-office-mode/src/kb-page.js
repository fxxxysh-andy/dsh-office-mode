/**
 * 知识库清单的服务端排序与分页（第九十四轮）。
 *
 * 面板要「真分页」：每一页都是服务端在**当前文件夹 + 过滤词**的全量行上排序后切出来
 * 的一小段，而不是把前 20 篇拿来在客户端里翻。两条契约必须一起守：
 *
 *   1. **排序在全量上做**。只在当前页里排序，翻到第 2 页会出现「第 2 页的 a 排在第 1 页
 *      的 b 前面」这种自相矛盾的顺序 —— 那是假分页。
 *   2. **平局要有稳定次序**。同一毫秒入库、同名、同大小在真实库里很常见；不 tie-break
 *      时两页之间可能出现重复或漏行（Array#sort 在旧引擎里不稳定，即便稳定也依赖
 *      输入顺序）。这里用 `id` 升序兜底，`id` 在 manifest 里唯一。
 *
 * 纯函数、不读盘：便于单测，也让 view.js 只负责「取全量行 → 调这里 → 回响应」。
 */

/** 面板默认每页条数（与历史清单上限一致，避免升级后首屏从 20 变多）。 */
export const KB_PAGE_SIZE = 20;

/** 面板可请求的每页上限：再大就不是「翻页」而是「把库灌进快照」。 */
export const KB_PAGE_MAX = 100;

/** 认得的排序键；拼错 / 旧客户端一律回落到 name（不是空排序）。 */
export const KB_SORT_KEYS = ['name', 'size', 'time', 'chunks'];

function textOf(value) {
    return typeof value === 'string' ? value : (value === null || value === undefined ? '' : String(value));
}

function finiteNumber(value) {
    const n = typeof value === 'number' ? value : Number(value);
    return Number.isFinite(n) ? n : 0;
}

/** 名称排序口径与面板一致：中文按拼音、数字按数值（2 在 10 前面）。 */
function compareName(left, right) {
    return textOf(left).localeCompare(textOf(right), 'zh-Hans-CN', { numeric: true, sensitivity: 'base' });
}

/**
 * 把排序键与方向归一：只认白名单里的键，方向只认 asc / desc。
 *
 * @param {string} key
 * @param {string} dir
 * @returns {{key: string, dir: 'asc'|'desc'}}
 */
export function normalizeKbSort(key, dir) {
    const wanted = textOf(key).toLowerCase();
    const safeKey = KB_SORT_KEYS.includes(wanted) ? wanted : 'name';
    const wantedDir = textOf(dir).toLowerCase();
    const safeDir = wantedDir === 'desc' ? 'desc' : 'asc';
    return { key: safeKey, dir: safeDir };
}

/**
 * 在一份行数组上做稳定排序（不改入参）。
 *
 * 比较顺序：主键 → `id` 升序（tie-break）。`id` 不跟着 dir 翻转：翻页时行在页间的
 * 归属必须只由主键决定，否则同一条会在两页里各出现一次。
 *
 * @param {Array<object>} rows manifest 行
 * @param {string} key name / size / time / chunks
 * @param {string} dir asc / desc
 * @returns {Array<object>} 新数组
 */
export function sortKbRows(rows, key, dir) {
    const list = Array.isArray(rows) ? rows.slice() : [];
    const sort = normalizeKbSort(key, dir);
    const sign = sort.dir === 'desc' ? -1 : 1;
    const primary = (row) => {
        const item = row !== null && typeof row === 'object' ? row : {};
        if (sort.key === 'size') return finiteNumber(item.bytes);
        if (sort.key === 'time') return textOf(item.at);
        if (sort.key === 'chunks') return finiteNumber(item.chunks);
        return textOf(item.title || item.path);
    };
    return list.sort((left, right) => {
        const a = primary(left);
        const b = primary(right);
        let delta = 0;
        if (typeof a === 'number' && typeof b === 'number') delta = a - b;
        else if (sort.key === 'name') delta = compareName(a, b);
        else delta = textOf(a).localeCompare(textOf(b));
        if (delta !== 0) return delta * sign;
        return textOf(left && left.id).localeCompare(textOf(right && right.id));
    });
}

/**
 * 把总量与请求的 offset / limit 归一成一页合法的切片窗口。
 *
 * 越界的 offset **回拉到最后一页的行首**而不是返回空页：用户在第 2 页删掉唯一一条时，
 * 客户端会带着旧 offset 再来一次；回空页会显示「没有文档」并让人以为库空了。
 * 返回的 `offset` 是**生效值**，客户端应把它写回状态（否则会一直带着越界值请求）。
 *
 * @param {number} total 过滤后的总行数
 * @param {number|string} offset 请求的起始行
 * @param {number|string} limit 请求的每页条数（缺省 KB_PAGE_SIZE，上限 KB_PAGE_MAX）
 * @returns {{offset: number, limit: number, total: number, pageCount: number, page: number, hasPrev: boolean, hasNext: boolean}}
 */
export function kbPageWindow(total, offset, limit) {
    const count = Math.max(0, Math.trunc(finiteNumber(total)));
    const wantedLimit = Math.trunc(finiteNumber(limit));
    const size = Math.min(KB_PAGE_MAX, Math.max(1, wantedLimit > 0 ? wantedLimit : KB_PAGE_SIZE));
    const pageCount = count === 0 ? 0 : Math.ceil(count / size);
    const lastStart = pageCount === 0 ? 0 : (pageCount - 1) * size;
    const wantedOffset = Math.max(0, Math.trunc(finiteNumber(offset)));
    const start = Math.min(wantedOffset, lastStart);
    return {
        offset: start,
        limit: size,
        total: count,
        pageCount,
        page: pageCount === 0 ? 1 : Math.floor(start / size) + 1,
        hasPrev: start > 0,
        hasNext: start + size < count,
    };
}
