/**
 * 引用纪律：**句级**引用审计（第五十七轮 19-8 的 P2-9）。
 *
 * 第十九轮附录把这条记成「落盘断言强制句级来源引用，评测套 ALCE 三指标
 * （正确性 / 召回 / 精确率）」。要落地就必须先回答一个绕不开的问题：
 * **ALCE 的三项里，哪几项在离线、无模型、无第三方依赖的本项目里算得出来？**
 *
 *   · **召回**（citation recall）：ALCE 的原定义是「回答里的每个陈述都被它引用的
 *     段落支撑」——「支撑」是蕴含判断，算不出来。**能算的是它的词法代理**：
 *     该给出处的句子有没有给出处。本模块把这一项叫 `recall` 并写成**句级引用率**，
 *     名实一致：它数的是「事实句里有引用的比例」，不是「引用真的支撑了这句话」。
 *   · **精确率**（citation precision）：原定义是「每条引用都支撑对应的陈述」——
 *     同样是蕴含判断。**能算的是**：「正文里写下的地址，是不是本次材料里真的取到的
 *     那批地址」。这一项是硬的：材料是刚落盘的那几份结果文件，地址对不上就是编的。
 *   · **正确性**（correctness）：要有事实真值与蕴含判断，**离线算不出**。
 *     这一项**照实标 `null`**，并在这里写死一句：**不许**拿上面两项的代理量冒充它。
 *     （第十九轮那条欠账拖到现在，卡的就是这句话先得定下来。）
 *
 * 判据全是**显式词表与显式正则**（年份 / 百分比 / 带外部量词的数字 / 归因主语），
 * 不是语义模型。这是刻意的：判据能被人逐条核对，测试能钉住它；代价是会漏判
 * （不在这几类里的外部断言不会被抓），也可能误判。所以审计的结论只有
 * 「按这份词表看来」这一层意思，回执里也照实这么写。
 *
 * 两个方向**不等价**，取舍一律偏向「宁可漏、不可误」：误判的后果是模型被逼着改写
 * 一句本来正常的话（而且是硬失败 `ok:false`），漏判只是少抓一条。所以词表里剔掉了
 * 那些有大量反例的词（见 `ATTRIBUTION_INLINE` 的注释）。
 *
 * @module dsh-office-mode/citation
 */

/**
 * 「这句话在声称一个外部事实」的判据之一：年份。
 *
 * **必须带「年」字**（`2024 年` / `2026 年度`）：光看四位数字会把分辨率、端口号、
 * 型号这类东西一起抓进来 —— 独立复核实测 `1920×1080` 就被判成了年份，而误判的
 * 代价是硬失败。中文稿子里年份几乎都写「年」，漏掉的是 `截至 2024，…` 这种写法
 * （属「宁可漏」那一侧）。
 */
const FACT_YEAR = /(?:19|20)\d{2}\s*年/;

/** 百分比：`42%` / `42.5 ％`。 */
const FACT_PERCENT = /\d+(?:\.\d+)?\s*[%％]/;

/**
 * 带**外部世界量词**的数字。
 *
 * 量词表刻意不含「个 / 项 / 条 / 页 / 份 / 台 / 章 / 节」：那些多半是在说这份产物
 * 自己的形状（「共 12 页」「3 条建议」），要求它们给出处只会制造假红。留下的是
 * 只能来自外部世界的量：钱、人、倍数、重量、距离、面积、温度。
 */
const FACT_NUMBER_UNIT = /\d[\d,.]*\s*(?:元|万元|亿元|亿美元|亿|倍|人|人次|家|所|吨|公里|千米|平方米|℃|摄氏度)/;

/**
 * 句首归因词：`据…` / `根据…` / `依据…`，且后面要出现**来源性名词**。
 *
 * 为什么不是「只要句首有『据 / 根据 / 依据 / 按 / 按照』就算」：`按 Enter 继续`、
 * `按照设置页里的通道顺序逐条试`、`根据判据,会漏也会误` 都是**讲方法**的正常句，
 * 强制它们给出处就是把好好的话逼着改（见文件头的取舍）。要求后面 16 字内出现
 * 数据 / 报告 / 通报 / 用户 / 客户 这类来源性名词，才判成「在引述外部信息」。
 */
const ATTRIBUTION_START = /^(?:据|根据|依据)[^，。；：]{0,16}(?:通报|报道|报告|公告|通知|统计|调查|研究|数据|消息|官方|文件|规定|标准|条例|用户|客户|访谈|问卷|反馈|来源|说明)/;

/**
 * **带归因主语的**引述：`数据显示` / `通报显示` / `报告称` …
 *
 * 为什么不直接收「显示」两个字：它太容易落在自指的句子里（「图 3 显示了整体架构」
 * 「显示器分辨率…」）。要求前面有归因主语，漏掉的是没主语的写法 —— 这正是上面
 * 那条「宁可漏、不可误」。
 */
const ATTRIBUTION_HEADED = /(?:数据|报告|通报|公告|统计|调查|研究|官方|媒体|新闻|问卷|审计|白皮书)(?:显示|表明|指出|称)/;

/**
 * 句中的归因 / 统计 / 断言词。
 *
 * 只收**在中文书面语里几乎只用来引述外部信息**的词。实测被剔掉的那几个都有反例：
 * 「显示」→ 显示器 / 图 3 显示了；「发布」→ 发布产物 / 发布流程；「预计 / 达到 / 超过」
 * → 「预计下周完成」「成本达到预期」「超过一半的篇幅」都是计划或自指；「唯一」→
 * 「唯一需要注意的是」。它们声称外部事实时几乎都带数字，由数字类判据接住。
 */
const ATTRIBUTION_INLINE = /(?:表明|报道|统计|宣布|公布|同比|环比|创下|占比|位居|首次)/;

/** 一层列表标记 / 标题标记：判据只看正文，不把标记算进句子。 */
const LEADING_MARKER = /^\s*(?:[-*+]|\d+[.)])\s+/;

/** Markdown 标题行不是事实句，而且是**合并屏障**（见 splitSentences）。 */
const HEADING_LINE = /^\s*#{1,6}\s+/;

/** 分隔线（`---` / `***` / `___`）既不是句子，也不是合并屏障之外的普通内容。 */
const DIVIDER_LINE = /^\s*(?:-{3,}|\*{3,}|_{3,})\s*$/;

/** 引用块与表格行不作为事实句判据的输入（它们是材料的摘录或结构化行）。 */
const QUOTE_OR_TABLE = /^\s*[|>]/;

/** 围栏代码块与行内代码：里面的地址是**示例**，不是引用。 */
const FENCED_CODE = /^[ \t]*(?:```|~~~)[\s\S]*?^[ \t]*(?:```|~~~)[ \t]*$/gm;
const INLINE_CODE = /`[^`\n]*`/g;

/**
 * 句界：中英文句末标点与换行。
 *
 * **分号不算句界**（独立复核 P1）：中文分号是**句内**的分句界（「一是…；二是…；三是…。」
 * 是极常见的写法），而引用通常只挂在最后一个分号之后 —— 按分号切会把前半句判成
 * 「没出处」，那是判据自己制造的假红。
 */
const SENTENCE_BREAK = /(?<=[。！？!?])/;

/** URL 的尾巴上不该算进地址的字符（含标记污染：加粗的 `**`、行内代码的反引号）。 */
const URL_TAIL = /[.,;:!?、，。；：！？*`_…）】\]]+$/;

/** 一行之内的 URL（与 findings.js 同一口径：通吃再统一剥尾）。 */
const URL_RE = /https?:\/\/[^\s)\]}>，。；、"'）】]+/g;

/** kb 块 id：`kb:<内容哈希>:<块号>`（哈希长度按 sha1 起、允许更长）。 */
const KB_ID_RE = /kb:[0-9a-f]{6,}:\d+/gi;

/**
 * 编号引用：`[3]` / 脚注 `[^3]` / 锚点式 `[3](#ref)`。
 *
 * 不排除后面紧跟 `(` 的写法（那正是锚点式的形状）；Markdown 链接的**文字**不是数字，
 * 天然不会命中。前置用 lookbehind 而不是消费一个字符：`[1][2]` 里第二处的前一个字符
 * 是 `]`，消费式的写法会把它吃掉（独立复核实测只数到 1 处）。
 */
const NUMBERED_RE = /(?<![\[\w])\[\^?(\d{1,3})\]/g;

/** 显式出处标记：`（来源：…）` / `(出处：…)` / `（依据：…）`。 */
const SOURCE_MARK_RE = /[（(]\s*(?:来源|出处|依据)\s*[：:]/g;

/** 把地址剥到可对账的形状：去尾部标点与标记污染。 */
function trimUrlTail(url) {
    return String(url ?? '').trim().replace(URL_TAIL, '');
}

/** 拆地址：主机 / 路径 / 查询串（query 保留 —— 见 normalizeCitationUrl 的注释）。 */
function splitUrl(url) {
    const raw = trimUrlTail(url);
    try {
        const parsed = new URL(/^https?:\/\//i.test(raw) ? raw : `https://${raw}`);
        return {
            host: parsed.host.replace(/^www\./i, '').toLowerCase(),
            path: parsed.pathname.replace(/\/+$/, '').toLowerCase(),
            query: parsed.search.toLowerCase(),
        };
    } catch {
        return null;
    }
}

/**
 * 把 URL 归一成可对账的形状：去协议、去 www、去 fragment、去尾斜杠、小写。
 *
 * **query 保留**（独立复核 P2）：很多站点用 `?id=` / `?p=` 标识页面，去掉 query 会把
 * 「两个不同页面」算成同一个地址 —— 这是本审计唯一那条硬判据被悄悄放松。代价是带
 * `?utm_*` 的同一页面会被判成不同地址；同一路径不同参数单列一档（见 `citeStatusOf`）。
 */
export function normalizeCitationUrl(url) {
    const parts = splitUrl(url);
    if (parts === null) return trimUrlTail(url).toLowerCase();
    return `${parts.host}${parts.path}${parts.query}`;
}

/**
 * 一条引用相对材料的状态。
 *
 *   `exact`      归一化之后完全相同 —— 材料里那一条就是这个地址；
 *   `query`      同主机同路径、**参数不同** —— 可能是另一个页面（`?id=1` / `?id=2`），
 *                算「对不上」，但要与「完全找不到」分开说；
 *   `site`       同主机、不同路径 —— 同站不同页（ALCE 里算**另一条**引用）；
 *   `unresolved` 材料里没有这个站点。
 */
function citeStatusOf(url, material) {
    const parts = splitUrl(url);
    if (parts === null) return 'unresolved';
    if (material.norms.includes(normalizeCitationUrl(url))) return 'exact';
    if (material.parts.some((item) => item.host === parts.host && item.path === parts.path)) return 'query';
    if (material.hosts.has(parts.host)) return 'site';
    return 'unresolved';
}

/**
 * 这一段是不是**只剩引用标记**（没有自己的正文）。
 *
 * 为什么要判它：中文稿子的出处最常见的写法是「……事实。（来源：https://…）」——
 * 引文在句号**后面**。若按句号硬切，出处会变成独立的一段，事实句就成了「没出处」，
 * 句级引用率直接腰斩。所以切完之后把纯引用段并回上一句。
 */
function isCitationOnly(piece) {
    const stripped = String(piece)
        .replace(URL_RE, '')
        .replace(KB_ID_RE, '')
        .replace(/[（(]\s*(?:来源|出处|依据)\s*[：:][^）)]*[）)]/g, '')
        .replace(/\[\^?\d{1,3}\]/g, '')
        .replace(/[\s，。；、,;:：()（）\[\]【】]/g, '');
    return stripped === '';
}

/** 抽出一段文本里的全部地址（材料侧对账用：材料是可信输入，取超集只会少报假红）。 */
export function urlsInText(text) {
    const found = [];
    URL_RE.lastIndex = 0;
    let match;
    while ((match = URL_RE.exec(String(text ?? ''))) !== null) {
        const url = trimUrlTail(match[0]);
        if (url !== '' && !found.includes(url)) found.push(url);
    }
    return found;
}

/**
 * 把一份正文切成句子（保留原文，不去标点）。
 *
 * 三层剔除：围栏 / 行内代码（里面的地址是示例）、标题行、引用块与表格行 ——
 * 它们不是「正文里的陈述」，拿它们要求出处只会制造噪声。
 *
 * **合并屏障**（独立复核 P1）：纯引用段只在**同一节内**并回上一句。跨标题绝不并 ——
 * 否则上一节那句本来没出处的事实句会因为下一节的 `## 来源` 后面那条地址而变成
 * 「有出处」，那是模型看不出来的假绿。
 */
export function splitSentences(text) {
    const out = [];
    let anchor = true;
    const body = String(text ?? '').replace(FENCED_CODE, '\n').replace(INLINE_CODE, ' ');
    for (const rawLine of body.split(/\r?\n/)) {
        const line = rawLine.trim();
        if (line === '') continue;
        if (HEADING_LINE.test(line) || DIVIDER_LINE.test(line)) {
            anchor = false;
            continue;
        }
        if (QUOTE_OR_TABLE.test(line)) continue;
        for (const piece of line.split(SENTENCE_BREAK)) {
            const value = piece.trim();
            if (value === '') continue;
            // 纯引用段并回上一句（见 isCitationOnly 与上面的「合并屏障」）。
            if (anchor && isCitationOnly(value) && out.length > 0) {
                out[out.length - 1] = `${out[out.length - 1]} ${value}`;
                continue;
            }
            out.push(value);
            // 并进上一句之后，锚点仍在（出处常写成多行）。
            anchor = true;
        }
    }
    return out;
}

/**
 * 这句话是不是「在声称一个外部事实」（判据见文件头）。
 *
 * 开销一句话说清：判据是词表，所以它**会漏、也会误**。审计的结论只是
 * 「按这份词表看来」，不是「这句话真的需要出处」。
 */
export function isFactSentence(sentence) {
    const body = String(sentence ?? '').replace(LEADING_MARKER, '').trim();
    if (body.length < 6) return false;
    if (FACT_YEAR.test(body) || FACT_PERCENT.test(body) || FACT_NUMBER_UNIT.test(body)) return true;
    if (ATTRIBUTION_START.test(body) || ATTRIBUTION_HEADED.test(body)) return true;
    return ATTRIBUTION_INLINE.test(body);
}

/**
 * 抽出一句话里的引用标记。
 *
 * 四种都认得出来，但**只有地址能对账**：kb 块 id 只能对形状（本模块没有库的入口）、
 * 出处标记与编号引用只算「标了出处」。后三种不进精确率的分母 —— 不让它们把精确率
 * 算成一个看起来很美的数字。
 */
export function citationsOf(sentence) {
    const body = String(sentence ?? '');
    const urls = [];
    const kbIds = [];
    let marker = 0;
    let numbered = 0;
    let match;
    URL_RE.lastIndex = 0;
    while ((match = URL_RE.exec(body)) !== null) {
        const url = trimUrlTail(match[0]);
        if (url !== '' && !urls.includes(url)) urls.push(url);
    }
    KB_ID_RE.lastIndex = 0;
    while ((match = KB_ID_RE.exec(body)) !== null) {
        const id = match[0].toLowerCase();
        if (!kbIds.includes(id)) kbIds.push(id);
    }
    SOURCE_MARK_RE.lastIndex = 0;
    while (SOURCE_MARK_RE.exec(body) !== null) marker += 1;
    NUMBERED_RE.lastIndex = 0;
    while (NUMBERED_RE.exec(body) !== null) numbered += 1;
    return { urls, kbIds, marker, numbered };
}

/** 一句话里有没有引用标记（四种都算）。 */
export function hasCitation(sentence) {
    const found = citationsOf(sentence);
    return found.urls.length > 0 || found.kbIds.length > 0 || found.marker > 0 || found.numbered > 0;
}

/**
 * 句级引用审计。
 *
 * @param {object} input
 * @param {string} input.text            产物正文（Markdown / 纯文本）
 * @param {string[]} [input.materialUrls] 本次材料里真的取到的地址
 * @param {string} [input.label]         正文的显示名（回执里点名用）
 * @returns {object} 审计结果；`correctness` 恒为 `null`（离线算不出，见文件头）
 */
export function auditCitations({ text = '', materialUrls = [], label = '' } = {}) {
    const raw = (Array.isArray(materialUrls) ? materialUrls : []).map(trimUrlTail).filter((url) => url !== '');
    const material = {
        norms: [...new Set(raw.map(normalizeCitationUrl))],
        parts: raw.map(splitUrl).filter((item) => item !== null),
        hosts: new Set(raw.map((url) => splitUrl(url)?.host).filter(Boolean)),
    };

    const sentences = splitSentences(text);
    const facts = [];
    const uncited = [];
    const citations = [];
    let numbered = 0;
    let markers = 0;
    let kbIds = 0;

    for (const [index, sentence] of sentences.entries()) {
        const found = citationsOf(sentence);
        const fact = isFactSentence(sentence);
        if (fact) facts.push(index);
        if (fact && !hasCitation(sentence)) uncited.push({ index, sentence });
        numbered += found.numbered;
        markers += found.marker;
        kbIds += found.kbIds.length;
        for (const url of found.urls) {
            citations.push({ url, sentence: index, status: citeStatusOf(url, material) });
        }
    }

    // 「对不上」的两档分得清：完全找不到的站点 vs 同一路径不同参数（可能是另一个页面）。
    const unresolved = citations.filter((item) => item.status === 'unresolved' || item.status === 'query');
    const siteOnly = citations.filter((item) => item.status === 'site');
    const exact = citations.filter((item) => item.status === 'exact');
    // 精确率的分母**只有地址**，分子**只算逐字对上的那批**（独立复核 P2：`site` 档
    // 计进分子会把唯一那条硬判据悄悄放松）。同站不同页单独报数，不夸大也不隐藏。
    const precision = {
        got: exact.length,
        total: citations.length,
        rate: citations.length === 0 ? null : exact.length / citations.length,
        site: siteOnly.length,
    };
    const recall = { got: facts.length - uncited.length, total: facts.length, rate: facts.length === 0 ? null : (facts.length - uncited.length) / facts.length };

    const notes = [];
    if (sentences.length === 0) notes.push('正文里没有可判的句子（空文件，或整行都是标题 / 分隔线 / 引用块 / 表格）。');
    if (facts.length === 0) notes.push('按事实句词表（年份 / 百分比 / 带量词的数字 / 归因主语）没有判出任何事实句 —— 这不是「引用都齐了」。');
    if (citations.length === 0) notes.push('正文里没有可对账的地址（URL）—— 出处标记与编号引用无法对账，不能代替地址。');
    if (siteOnly.length > 0) notes.push(`${siteOnly.length} 条地址只对到**站点**、不是材料里那一个页面（同站不同页在 ALCE 里算另一条引用），不计入上面的精确率。`);
    if (kbIds > 0) notes.push(`${kbIds} 条 kb 块 id 只核了形状（本审计没有库的入口，是否真的在库里要用 kb-read 看）。`);
    if (markers > 0) notes.push(`${markers} 处「（来源：…）」标记无法对账，只算标了出处。`);
    if (numbered > 0) notes.push(`${numbered} 处编号引用（[1] 形）无法对账，只算标了出处。`);
    notes.push('correctness（ALCE 第三项）离线算不出：它要蕴含判断与事实真值，这里恒为 null —— 不许用上面的召回 / 精确率冒充。');

    // 三态：`fix` 有硬问题；`empty` 这次什么也没验证（不许渲染成「通过」）；`pass` 真的过了。
    const verdict = uncited.length > 0 || unresolved.length > 0
        ? 'fix'
        : (facts.length === 0 || citations.length === 0 ? 'empty' : 'pass');

    return {
        label: String(label ?? ''),
        sentences: sentences.length,
        facts: facts.length,
        recall,
        precision,
        // 两项之外**照实标出**的那一项：算不出就是 null，不放代理量。
        correctness: null,
        uncited,
        unresolved,
        siteOnly,
        citations,
        notes,
        verdict,
    };
}

/**
 * 把审计结果渲染成给主会话看的几行。
 *
 * 与 `renderFindings` 同一条纪律：进上下文的是结论与要改的那几处，不是整份正文。
 * 每一条都带**句号**（`第 N 句`）：只给原文片段的话，超过 8 条之后要重读整份草稿
 * 才知道其余在哪。
 */
export function renderCitationAudit(audit) {
    const name = audit.label === '' ? '产物正文' : audit.label;
    const percent = (rate) => (rate === null ? '—' : `${Math.round(rate * 1000) / 10}%`);
    const title = audit.verdict === 'pass' ? '通过' : audit.verdict === 'empty' ? '**没判出可核的东西**' : '**未过**';
    const lines = [`引用审计（${name}）：${title}`
        + `｜句子 ${audit.sentences} 句、其中事实句 ${audit.facts} 句`];
    lines.push(`- 句级引用率 ${audit.recall.got}/${audit.recall.total}（${percent(audit.recall.rate)}）`
        + `｜地址逐字对账 ${audit.precision.got}/${audit.precision.total}（${percent(audit.precision.rate)}）`
        + (audit.precision.site > 0 ? `，另有 ${audit.precision.site} 条只对到站点` : '')
        + '｜correctness：算不出（null）');
    if (audit.uncited.length > 0) {
        lines.push(`- **${audit.uncited.length} 句声称外部事实却没给出处**（要么补上材料里那条地址，要么把这句改成不声称事实的说法）：`);
        for (const item of audit.uncited.slice(0, 8)) lines.push(`  · 第 ${item.index + 1} 句：${item.sentence.slice(0, 120)}`);
        if (audit.uncited.length > 8) lines.push(`  · …其余 ${audit.uncited.length - 8} 句按上面的句号在同一份正文里定位`);
    }
    if (audit.unresolved.length > 0) {
        lines.push(`- **${audit.unresolved.length} 条地址对不上本次材料**（材料里没有的地址不要写进正文）：`);
        for (const item of audit.unresolved.slice(0, 8)) {
            const why = item.status === 'query' ? '（同一路径、不同参数：可能是另一个页面）' : '';
            lines.push(`  · 第 ${item.sentence + 1} 句：${item.url}${why}`);
        }
        if (audit.unresolved.length > 8) lines.push(`  · …其余 ${audit.unresolved.length - 8} 条按上面的句号在同一份正文里定位`);
    }
    if (audit.notes.length > 0) {
        lines.push('- 口径：');
        for (const note of audit.notes) lines.push(`  · ${note}`);
    }
    return lines.join('\n');
}
