/**
 * 热记忆投影的「只在变化时贴」。
 *
 * 为什么需要它：热记忆的投影是挂在 office_help / office_run **每一次**调用上的
 * （办公模式的 persona 是 complete 的，插件注册的提示段会被组装丢掉，所以只能
 * 挂在模型必然会调的调用上）。一份 ~10 KB 的全文在同一轮里被贴 10 次，就是 10 份
 * 都留在上下文里 —— 而热记忆绝大多数时候是不变的。所以「没变就只报一行」既省
 * 上下文，也不丢信息：那一行说清了没变、revision 是多少、要看内容怎么读。
 *
 * 三条实现纪律：
 *
 *   1. **不是「永远不重复贴」**。宿主的工具结果压缩
 *      （dsh-compaction-tool-result-pruner）会把很久以前的调用结果裁掉，那时
 *      上下文里就没有热记忆了，而模型不会知道自己缺了什么。所以每
 *      `REFRESH_EVERY` 次投影强制重贴一次全文 —— 这一条是**防丢失**，
 *      不是防「没看到」。拿不到会话身份时同理：老老实实每次都贴全文，
 *      不假装「刚贴过」（静默省略才是真的丢信息）。
 *
 *      **残余风险要如实说**：这个计数是「按次数」的近似，不是「投影还在不在」
 *      的精确判断。如果在第 3 次调用之前恰好发生了压缩、把第一次那份全文裁掉了，
 *      那么第 4 到第 12 次之间模型手里没有热记忆（只看到「与上一次投影相同」，
 *      以及一行「要看：office_memory({ action: 'read', layer: 'hot' })」）。
 *      精确做法是去会话日志里找上一次投影还在不在（压缩会把旧结果换成占位符），
 *      但那要依赖 `ownEvents()` 里工具结果事件的具体形状 —— 猜错的后果是这条守卫
 *      **静默失效**（要么永不触发、要么次次触发），比现在这个有界的近似更糟。
 *      所以这里选了「简单且一定生效」的那个，并把边界写在这里。
 *
 *   2. **台账部分照旧每次都贴**。台账最近几条是「这次交付之后新增了什么」的
 *      直接反馈，而它**不**体现在热记忆的 revision 里（revision 只由热记忆条目
 *      算出来）—— 把它一起省掉会让 office_run 之后看不到刚登记的那条。
 *      省的是热记忆的正文，不是整个投影。
 *
 *   3. **状态按会话记，且是有界的**。会话之间互不影响（一个会话贴过不代表另一个
 *      看过）；Map 只保留最近 `MAX_SESSIONS` 个会话，长跑进程不会因此涨内存。
 *
 * 同一份「省前缀」的取向还用在**固定的长指引**上：`hintOnce(sessionId, key)`
 * 管那些每次调用都一模一样的说明句（怎么读检索结果、不要把原文搬进对话…），
 * 一个会话里只说一次。它与热记忆投影的关键差别是**拿不到会话身份时也去重**，
 * 且不参与 `resetProjection()` —— 理由写在 `hintOnce` 的注释里。
 *
 * 真实量级（本机工作区实测，热记忆 user 302 + project 10155 字节）：
 * 全文 11714 字节 / 省略形态 1442 字节 —— 8.1x；一轮里调 10 次办公工具，
 * 从 117 KB 降到 24 KB。
 *
 * @module dsh-office-mode/projection
 */

import { renderDigest } from './memory.js';

/** 连续多少次「未变化」之后强制重贴一次全文（防宿主压缩把早先那份裁掉）。 */
export const REFRESH_EVERY = 12;
/** 最多记多少个会话的投影状态（超出按插入序淘汰最旧的）。 */
const MAX_SESSIONS = 32;
/** 最多记多少条「这条长指引已经说过」（超出按插入序淘汰最旧的）。 */
const MAX_HINTS = 64;

/** sessionId → { revision, since }。since = 自上次贴全文以来省略了几次。 */
const state = new Map();
/** 「会话|指引标识」→ true。与 state 分开存，见 hintOnce 的注释。 */
const hints = new Map();

/**
 * 取这次调用的会话 id（拿不到返回空串）。
 *
 * 与 quota.js 的 `turnKeyOf` 同一来源：`exec.agent.session`。两处都只读这一个
 * 字段，所以任何一处改了会话形状，另一处也要一起改。
 *
 * @param {object} exec 工具执行上下文（宿主给的 ToolRunContext）
 * @returns {string} 会话 id，拿不到就是空串
 */
export function sessionIdOf(exec) {
    const session = exec?.agent?.session;
    if (session === undefined || session === null) return '';
    try {
        const id = typeof session.id === 'string' ? session.id : session.header?.id;
        return typeof id === 'string' ? id.trim() : '';
    } catch {
        return '';
    }
}

/** 只保留最近若干个会话的投影状态。 */
function evict() {
    while (state.size > MAX_SESSIONS) {
        const oldest = state.keys().next().value;
        if (oldest === undefined) return;
        state.delete(oldest);
    }
}

/**
 * 决定这一次投影贴全文还是只报「未变化」。
 *
 * @param {object} digest `memory.digest()` 的结果（含 revision）
 * @param {{sessionId?: string, context?: 'help'|'run'}} options
 * @returns {{text: string, mode: 'full'|'unchanged', reason: string}}
 *   text 就是要追加到工具反馈末尾的那段；mode / reason 给测试与排查用
 */
export function projectDigest(digest, { sessionId = '', context = 'help' } = {}) {
    // 空记忆：整段本来就小（讲的是「怎么记」，不是内容），而且那几句指引正是
    // 该反复出现的时候 —— 不做省略。
    if (digest?.empty === true) {
        return { text: renderDigest(digest, { context }), mode: 'full', reason: 'empty' };
    }
    const id = typeof sessionId === 'string' ? sessionId.trim() : '';
    if (id === '') {
        // 拿不到会话身份：不假装「刚贴过」。
        return { text: renderDigest(digest, { context }), mode: 'full', reason: 'no-session' };
    }
    const revision = String(digest?.revision ?? '');
    const previous = state.get(id);
    const changed = previous === undefined || previous.revision !== revision;
    const since = changed ? 0 : previous.since + 1;
    const refresh = !changed && since >= REFRESH_EVERY;
    state.set(id, { revision, since: refresh ? 0 : since });
    evict();
    if (changed) return { text: renderDigest(digest, { context }), mode: 'full', reason: 'changed' };
    if (refresh) return { text: renderDigest(digest, { context }), mode: 'full', reason: 'refresh' };
    return { text: renderDigest(digest, { context, hot: 'unchanged' }), mode: 'unchanged', reason: 'unchanged' };
}

/** 清空投影状态（测试用；进程内状态不该被测试互相污染）。 */
export function resetProjection() {
    state.clear();
}

/**
 * 「这条长指引在本会话里已经说过没有」。
 *
 * 用途与热记忆投影同源：工具反馈里那些**每次调用都一模一样**的说明句
 * （怎么读结果、不要把原文搬进对话、改了渠道要重新出提纲…）一份就够。
 * 同一轮里 office_search_run 会被调 5-8 次，每次都带同一段话，等于把它
 * 复制 8 份进前缀 —— 前缀是只增不减的，省下来的是每一次后续请求都在付的钱。
 *
 * 与 `projectDigest` 的三点区别，都是刻意的：
 *   1. **不参与 resetProjection()**：它是「说过了」的账，测试清投影状态不该
 *      把它一起清掉（否则第二个用例又会拿到长指引，断言看着像没过）。
 *   2. **拿不到会话身份时也去重**：监听里的 exec 未必带 session（单测的替身
 *      往往没有），那种情况下按「进程内说过一次」算 —— 重复长句的代价比
 *      「第二次少一句说明」大。
 *   3. **有界**：最多记 `MAX_HINTS` 条，超出按插入序淘汰，长跑进程不会涨内存。
 *
 * @param {string} sessionId 会话 id（可为空串）
 * @param {string} key 指引的标识（如 'search-run' / 'dispatch'）
 * @returns {boolean} true = 这条指引这次该说（之前没说过）
 */
export function hintOnce(sessionId, key) {
    const id = `${typeof sessionId === 'string' ? sessionId.trim() : ''}|${String(key ?? '')}`;
    if (hints.has(id)) return false;
    hints.set(id, true);
    while (hints.size > MAX_HINTS) {
        const oldest = hints.keys().next().value;
        if (oldest === undefined) break;
        hints.delete(oldest);
    }
    return true;
}

/** 看一眼某条指引说过没有（测试与排查用）。 */
export function hintSeen(sessionId, key) {
    return hints.has(`${typeof sessionId === 'string' ? sessionId.trim() : ''}|${String(key ?? '')}`);
}

/**
 * 清空「长指引说过没有」的账（**只给测试用**）。
 *
 * 与 `resetProjection()` 分开：投影状态是每个用例自己造的记忆决定的，必须清；
 * 而指引状态是「说过一次」的账，正常运行时不该被任何人清掉。测试套件在开始
 * 与每个用例之间调用它，才能既验「第一次说」又验「第二次不说」。
 */
export function resetHints() {
    hints.clear();
}

/** 看一眼某个会话的投影状态（测试与排查用）。 */
export function projectionStateOf(sessionId) {
    const found = state.get(sessionId);
    return found === undefined ? null : { revision: found.revision, since: found.since };
}
