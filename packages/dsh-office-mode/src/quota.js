/**
 * 每回合配额：一个回合里每类检索最多几次。
 *
 * 为什么需要它：召回质量管的是「一次查询返回什么」，配额管的是「一个回合里
 * 能查几次」。没有配额时，模型遇到「检索没查到想要的」时的自然反应是换个词
 * 再查一遍，一直换到上下文被检索结果填满为止 —— mnemon 用回合预算堵的正是
 * 这件事。
 *
 * 回合身份从哪儿来：工具拿到的 `exec.agent.session` 是活的会话对象，它的
 * `ownEvents()` 是本次会话的事件日志，`turn/start` 事件带 `data.turn`。
 * 所以「这是第几回合」是可以**读出来**的，不需要宿主额外授权 ——
 * 向后扫日志取最后一个 turn/start 即可。
 *
 * 两条实现纪律：
 *   1. **扫日志是有界的**。正常情况下一个回合内的事件不会太多，但真遇到超长
 *      回合也不能每次都从头扫：按会话缓存「上次事件数 → 当时的 turn 号」，
 *      只在事件变多时从尾部往前扫，且最多扫 `MAX_SCAN` 条。
 *   2. **拿不到回合身份时不假装有配额**。没有会话对象时直接放行，并把
 *      `enforced: false` 报给调用方 —— 静默「当作已用完」会让检索莫名其妙失效。
 *
 * @module dsh-office-mode/quota
 */

/** 一次回溯最多扫多少条事件：超长回合也不会因此变慢。 */
const MAX_SCAN = 600;

/** 会话缓存：sessionId → { count, turn }。只在事件数变化时重扫。 */
const turnCache = new Map();
/** 缓存条数上限：只保留最近若干个会话，避免长时间运行后无限增长。 */
const MAX_SESSIONS = 32;

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

/**
 * 取这次调用的回合身份。
 *
 * @param {object} exec 工具执行上下文（宿主给的 ToolRunContext）
 * @returns {{sessionId: string, turn: number, key: string}|null} 拿不到就返回 null
 */
export function turnKeyOf(exec) {
    const session = exec?.agent?.session;
    if (session === undefined || session === null) return null;
    let sessionId = '';
    try {
        sessionId = asText(typeof session.id === 'string' ? session.id : session.header?.id);
    } catch {
        sessionId = '';
    }
    if (sessionId === '') return null;

    let events = null;
    try {
        events = typeof session.ownEvents === 'function' ? session.ownEvents() : null;
    } catch {
        events = null;
    }
    if (!Array.isArray(events)) return null;

    const cached = turnCache.get(sessionId);
    let turn = cached?.turn ?? 0;
    if (cached === undefined || cached.count !== events.length) {
        const floor = cached === undefined ? Math.max(0, events.length - MAX_SCAN) : Math.max(0, Math.min(cached.count, events.length - MAX_SCAN));
        for (let i = events.length - 1; i >= floor; i -= 1) {
            const event = events[i];
            if (event?.type !== 'turn/start') continue;
            const value = event.data?.turn;
            if (typeof value === 'number' && Number.isFinite(value)) {
                turn = value;
                break;
            }
        }
        turnCache.set(sessionId, { count: events.length, turn });
        if (turnCache.size > MAX_SESSIONS) {
            // 简单的 LRU：Map 保持插入序，删掉最早那一条。
            const oldest = turnCache.keys().next().value;
            if (oldest !== undefined && oldest !== sessionId) turnCache.delete(oldest);
        }
    }
    return { sessionId, turn, key: `${sessionId}#${turn}` };
}

/**
 * 建一个回合配额计数器。
 *
 * @param {{limits?: object}} options 每类动作的每回合上限；0 或缺失 = 不限制
 */
export function createTurnQuota({ limits = {} } = {}) {
    /** key → { [kind]: 已用次数 }。 */
    const used = new Map();

    /**
     * 申请一次配额。
     *
     * @param {string} kind 动作类别（recall / recallRefine / related / search）
     * @param {object} exec 工具执行上下文
     * @returns {{allowed: boolean, enforced: boolean, used: number, limit: number, turn: number|null, reason: string}}
     */
    function take(kind, exec) {
        const limit = Number.isFinite(limits[kind]) ? Math.trunc(limits[kind]) : 0;
        const identity = turnKeyOf(exec);
        if (identity === null) {
            // 没有回合身份：放行，但如实说「没在管」。
            return { allowed: true, enforced: false, used: 0, limit, turn: null, reason: '拿不到回合身份（会话对象不可用），本次未计入配额' };
        }
        if (limit <= 0) {
            return { allowed: true, enforced: false, used: 0, limit, turn: identity.turn, reason: '该类配额已关闭（0 = 不限制）' };
        }
        const counter = used.get(identity.key) ?? {};
        const current = Number.isFinite(counter[kind]) ? counter[kind] : 0;
        if (current >= limit) {
            return {
                allowed: false,
                enforced: true,
                used: current,
                limit,
                turn: identity.turn,
                reason: `本回合（第 ${identity.turn} 回合）的「${kind}」已用完 ${current}/${limit} 次`,
            };
        }
        counter[kind] = current + 1;
        used.set(identity.key, counter);
        // 只保留最近若干回合的计数，避免长会话里无限增长。
        if (used.size > MAX_SESSIONS * 4) {
            const oldest = used.keys().next().value;
            if (oldest !== undefined && oldest !== identity.key) used.delete(oldest);
        }
        return { allowed: true, enforced: true, used: counter[kind], limit, turn: identity.turn, reason: '' };
    }

    /**
     * 退回一次名额（只给「这次调用抛错了」那条路用）。
     *
     * 为什么需要它：名额是在**参数校验之前**申请的 —— 不先占名额就去做重活，两件事都不对；
     * 但这样「模型把 `tier` / 时间词拼错一个字母」这种当场抛错的调用也会吃掉一次机会。
     * 第四十五轮独立复核实测（P3）：默认 2 次下，一次错拼 tier + 一次正常检索之后，
     * 第 3 次就被拒 —— 模型会以为是检索坏了。
     *
     * 退回的语义是「**没有产生结果，所以不算查过一次**」，不是「事后无限重试」：
     * 只有抛错这条路上调用方才会调它，且名额不会被退成负数。
     */
    function release(kind, exec) {
        const identity = turnKeyOf(exec);
        if (identity === null) return false;
        const counter = used.get(identity.key);
        if (counter === undefined) return false;
        const current = Number.isFinite(counter[kind]) ? counter[kind] : 0;
        if (current <= 0) return false;
        counter[kind] = current - 1;
        used.set(identity.key, counter);
        return true;
    }

    /** 本回合已经用掉多少（给「回合记忆条」与回执用）。 */
    function usedIn(exec) {
        const identity = turnKeyOf(exec);
        if (identity === null) return null;
        const counter = used.get(identity.key) ?? {};
        return { turn: identity.turn, counts: { ...counter } };
    }

    /** 清空计数（测试与设置页「重置」用）。 */
    function reset() {
        used.clear();
        turnCache.clear();
    }

    return { take, release, usedIn, reset };
}
