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
/** 一个会话最多记多少条「已经贴过」的条目 id（超出按插入序淘汰最旧的）。 */
const MAX_NOTED = 64;
/** 最多记多少个会话的投影状态（超出按插入序淘汰最旧的）。 */
const MAX_SESSIONS = 32;
/** 最多记多少条「这条长指引已经说过」（超出按插入序淘汰最旧的）。 */
const MAX_HINTS = 64;

/**
 * 「本次在做什么」的话题词表（第三十六轮：按需投影）。
 *
 * 口径三条：
 *   1. **词是可以出现在记忆条目正文里的词**，不是 API 名的同义词表 —— 匹配方向是
 *      「条目正文包含词」，所以中文词取两字以上、英文取小写整词，单字（如「域」）
 *      与过泛的词（如「文件」「页面」单独使用时）会大面积误命中，宁可不放。
 *   2. **多给比少给安全**：词表偏宽只损失一点字节，偏窄会把该看到的条目折叠掉。
 *      折叠不是丢信息（折叠行带条数与读取入口），但「看不到」的代价仍大于「多看到」。
 *   3. 话题之间允许重叠（「表格」同时在 word 与 excel、「主题」同时在 ppt 与 theme）：
 *      一条记忆本来就可能同时属于两个话题。
 */
export const SIGNAL_TERMS = {
    word: ['word', 'docx', 'winword', '文档', '排版', '目录', '页码', '字体', '段落', '表格', '弹窗', '标题', '题注'],
    excel: ['excel', 'xlsx', 'sheet', '工作表', '单元格', '公式', '透视', '图表'],
    ppt: ['ppt', 'pptx', '幻灯', '版面', '审阅', '微调', '插图', '母版', '演示', '参考稿', '页面填充'],
    tex: ['tex', 'latex', 'thuthesis', '论文', '学位论文', 'latexmk', 'xelatex', 'texlive'],
    pdf: ['pdf', 'fitz', 'poppler', 'pdftoppm', '页面图', '文本型', '扫描件'],
    python: ['python', 'numpy', 'scipy', 'pandas', 'matplotlib', 'sympy', '绘图', '科学计算'],
    av: ['音频', '视频', '语音', '转写', '抽帧', '分块', 'asr', 'sensevoice', 'sherpa', 'ffmpeg'],
    files: ['批量', '改文件', '回写', '编码', '乱码', 'powershell', '占位符', '模板填充', '换行符'],
    cache: ['缓存', 'ttl', '命中复用', 'keepcache', '中间产物'],
    theme: ['主题', '令牌', '配色', '浅色', '深色', '亮色'],
    search: ['检索', '搜索', '通道', '来源', '代理', 'proxy', 'duckduckgo', 'searxng', '联网', '站点', '网页', '抓取', '预处理', 'url', 'web', 'provider'],
    memory: ['记忆', '召回', '归档', '台账', '投影', 'mnemon', '配额', '墓碑', '迁移', 'pack', '热记忆'],
    settings: ['设置', '面板', '分组', '折叠', '定位', '滚动', '字号', 'localStorage'],
};

/** 脚本里能认出来的文件扩展名 → 话题。office.* 的 API 名与话题同名，走词表的键。 */
const SCRIPT_EXTENSIONS = [
    ['.docx', 'word'], ['.doc', 'word'], ['.xlsx', 'excel'], ['.xls', 'excel'],
    ['.pptx', 'ppt'], ['.pdf', 'pdf'], ['.py', 'python'], ['.tex', 'tex'],
    ['.mp3', 'av'], ['.wav', 'av'], ['.m4a', 'av'], ['.mp4', 'av'], ['.mkv', 'av'],
];

/**
 * 扩展名匹配要带边界：`.pdf` 不能命中 `office.pdf.text(` 里的 `.tex`（那是
 * `.text(` 的一部分），`.py` 也不能命中 `.python`。要求扩展名后面跟的不是
 * 字母数字（引号、括号、空白、行尾都行）。
 */
const extensionRegex = () => new RegExp(
    `\\.(${SCRIPT_EXTENSIONS.map(([extension]) => extension.slice(1)).join('|')})(?![a-z0-9])`,
    'gi',
);

/**
 * 从「这次调用在做什么」里取话题信号。
 *
 * 两个来源：office_help 的 `topic`（模型自己点名的话题，最干净）；office_run 的
 * `script`（里面写的 `office.ppt.…` 与文件扩展名就是它要动的东西）。两处都给时
 * 取并集。认不出来（空话题、guide/run 这类总览话题、没有 office.* 的脚本）返回
 * null —— **没有信号就不筛**，宁可用旧行为也不猜。
 *
 * @param {{topic?: string, script?: string}} source
 * @returns {{key: string, terms: string[]}|null} key 用于「换了话题要重贴」，terms 给匹配用
 */
export function signalOf({ topic = '', script = '' } = {}) {
    const ids = new Set();
    if (typeof topic === 'string' && Object.prototype.hasOwnProperty.call(SIGNAL_TERMS, topic)) ids.add(topic);
    if (typeof script === 'string' && script !== '') {
        for (const match of script.matchAll(/office\.([a-zA-Z]+)/g)) {
            const name = match[1].toLowerCase();
            if (Object.prototype.hasOwnProperty.call(SIGNAL_TERMS, name)) ids.add(name);
        }
        const lowered = script.toLowerCase();
        for (const match of lowered.matchAll(extensionRegex())) {
            const extension = `.${match[1].toLowerCase()}`;
            const id = SCRIPT_EXTENSIONS.find(([wanted]) => wanted === extension)?.[1];
            if (id !== undefined) ids.add(id);
        }
    }
    if (ids.size === 0) return null;
    const terms = [...new Set([...ids].sort().flatMap((id) => SIGNAL_TERMS[id]))];
    return { key: [...ids].sort().join('+'), terms };
}

/**
 * sessionId → { writeRevision, ledgerRevision, since, signalKey, noted, refreshes }。
 * `since` = 自上次贴正文以来省略了几次；`noted` = 本会话**已经贴过正文**的条目 id
 * （第四十八轮 P0-2：节流的依据不再是「上一次是什么话题」，而是「哪些条目真的贴过」——
 * 话题 A→B→A 时第二次 A 不再重付全文）；`refreshes` = 本会话做过几次心跳重贴。
 */
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
 * 决定这一次投影贴什么。
 *
 * 三个维度各有各的时钟（第三十六轮起）：
 *   · **热记忆正文**：`writeRevision`（「有没有写操作」那个指纹）变了、第一次见这个
 *     会话、或**话题信号换了**（同一份热记忆，模型这次问的是另一件事）→ 贴；贴的时候
 *     **按信号筛**，无关条目折叠成一行（带条数与读取入口）。连续 `REFRESH_EVERY` 次
 *     没变 → 强制重贴**不筛的**全文 —— 心跳的职责是完整性（防宿主压缩把先前那份裁掉），
 *     所以这一份故意不筛。
 *   · **台账近况**：用自己的 `ledgerRevision`（最近几条的指纹）判断，变了就贴，
 *     没变也折叠成一行。台账不体现在热记忆的 `writeRevision` 里，但它同样只在登记新产物
 *     时才变 —— 两次交付之间反复贴同样三行，付的是纯重复的钱。
 *   · **读全文的入口**：任何折叠行都必须带 `office_memory` 的读取调用 —— 折叠是
 *     「把选择权交回给模型」，不是「替模型决定它不需要」。
 *
 * 指纹用哪个**故意写清楚**（第二十四轮 P2-5 / 第二十六轮 24-17）：这里要的是
 * 「有没有写操作」`writeRevision`，不是 memory.json 里那个由内容算的 `contentRevision`。
 *
 * @param {object} digest `memory.digest()` 的结果（含 writeRevision 与 ledgerRevision）
 * @param {{sessionId?: string, context?: 'help'|'run', signal?: {key: string, terms: string[]}|null, budget?: number|null}} options
 *   `budget`（第四十八轮 P0-2）= 条目正文的字节上限，透传给 renderDigest；不传 = 不限。
 * @returns {{text: string, mode: 'full'|'unchanged', reason: string, bytes?: number, printed?: number}}
 *   text 就是要追加到工具反馈末尾的那段；mode / reason 给测试与排查用。
 *   mode 为 'full' 表示「这次贴了新内容」（热记忆或台账任一）。
 */
export function projectDigest(digest, { sessionId = '', context = 'help', signal = null, budget = null } = {}) {
    // 空记忆：整段本来就小（讲的是「怎么记」，不是内容），而且那几句指引正是
    // 该反复出现的时候 —— 不做省略。
    if (digest?.empty === true) {
        return { text: renderDigest(digest, { context }), mode: 'full', reason: 'empty' };
    }
    const id = typeof sessionId === 'string' ? sessionId.trim() : '';
    const terms = signal !== null && signal !== undefined && Array.isArray(signal.terms) && signal.terms.length > 0
        ? signal.terms
        : null;
    const signalKey = terms === null ? '' : String(signal?.key ?? terms.join(','));
    if (id === '') {
        // 拿不到会话身份：不假装「刚贴过」。信号照样筛（它与谁是会话无关）。
        return {
            text: renderDigest(digest, { context, signal: terms === null ? null : { label: signalKey, terms } }),
            mode: 'full',
            reason: 'no-session',
        };
    }
    const revision = String(digest?.writeRevision ?? '');
    const ledgerRevision = String(digest?.ledgerRevision ?? '');
    const previous = state.get(id);
    const hotChanged = previous === undefined || previous.writeRevision !== revision;
    const ledgerChanged = previous === undefined || previous.ledgerRevision !== ledgerRevision;
    // 记下「这次真正贴出去的条目」（第四十八轮 P0-2）：下次遇到同一条就直接折叠，
    // 话题来回切（A→B→A）不再重付一份全文。
    const remember = (base, printed) => {
        const noted = new Set(base.noted ?? []);
        for (const entryId of printed) {
            noted.delete(entryId);
            noted.add(entryId);
        }
        while (noted.size > MAX_NOTED) {
            const oldest = noted.keys().next().value;
            if (oldest === undefined) break;
            noted.delete(oldest);
        }
        return noted;
    };
    const emit = (activeSignal, hot, ledger, skip) => {
        const report = { printedIds: [] };
        const text = renderDigest(digest, {
            context,
            hot,
            ledger,
            signal: activeSignal,
            budget,
            skip,
            report,
        });
        return { text, report };
    };
    const signalArg = terms === null ? null : { label: signalKey, terms };
    if (hotChanged) {
        const { text, report } = emit(signalArg, 'full', ledgerChanged ? 'full' : 'unchanged', null);
        state.set(id, {
            writeRevision: revision, ledgerRevision, since: 0, signalKey,
            noted: remember({ noted: new Set() }, report.printedIds), refreshes: 0,
        });
        evict();
        return { text, mode: 'full', reason: 'changed', bytes: report.bytes, printed: report.printedIds.length };
    }
    let since = previous.since + 1;
    const refresh = since >= REFRESH_EVERY;
    // 心跳照旧**不筛**（完整性兜底：宿主压缩把先前那份裁掉时，靠它把条目带回来），
    // 但同样受预算约束 —— 于是「每 12 次一次全文」从 15 KB 级降到预算级。
    if (refresh) {
        const { text, report } = emit(null, 'full', ledgerChanged ? 'full' : 'unchanged', null);
        state.set(id, {
            writeRevision: revision, ledgerRevision, since: 0, signalKey,
            noted: remember({ noted: new Set() }, report.printedIds),
        });
        evict();
        return {
            text,
            mode: 'full',
            reason: 'refresh',
            bytes: report.bytes,
            printed: report.printedIds.length,
        };
    }
    if (ledgerChanged) {
        const { text, report } = emit(null, 'unchanged', 'full', null);
        state.set(id, { ...previous, ledgerRevision, since, signalKey });
        return { text, mode: 'full', reason: 'ledger', bytes: report.bytes, printed: 0 };
    }
    // 没写、没登记、也不是心跳：只有在这次的话题下**还有没贴过的条目**时才贴正文。
    const { text, report } = emit(signalArg, 'full', 'unchanged', previous.noted);
    state.set(id, { ...previous, ledgerRevision, since, signalKey });
    if (report.printedIds.length > 0) {
        state.set(id, {
            ...previous, ledgerRevision, since: 0, signalKey,
            noted: remember(previous, report.printedIds),
        });
        return { text, mode: 'full', reason: 'changed', bytes: report.bytes, printed: report.printedIds.length };
    }
    return {
        text: renderDigest(digest, { context, hot: 'unchanged', ledger: 'unchanged' }),
        mode: 'unchanged',
        reason: 'unchanged',
        bytes: undefined,
        printed: 0,
    };
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
    return found === undefined
        ? null
        : { writeRevision: found.writeRevision, ledgerRevision: found.ledgerRevision, since: found.since, signalKey: found.signalKey };
}

