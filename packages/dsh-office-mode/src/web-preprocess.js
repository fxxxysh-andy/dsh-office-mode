/**
 * 网页正文的脚本式预处理管线 —— 取回来的页面**先清洗再进上下文**。
 *
 * ## 为什么要有这一步
 *
 * 第十八轮之前，取正文的出口是 `htmlToText`：把 HTML 收成文本，链接写成
 * `[文字](地址)`。它对「读得出内容」够用，但喂给模型的文本里混着导航条、
 * 菜单、cookie 条、页脚、推荐位与分享按钮 —— 一个正文 800 字的页面能洗出
 * 6000 字，其中八成是栏目名。两个后果：
 *
 *   1. **上下文白花钱**：检索结果文件与摘要按字节付费，噪声直接换成 token；
 *   2. **摘录变差**：`excerptOf` 只好靠「够长、不带链接标记、不以导航词开头」
 *      这几条启发式去猜正文行（第十七轮的补丁），猜错的概率随页面复杂度上升。
 *
 * 所以这一步是**确定性脚本**（无模型、无网络、无依赖）：去脚本样式 → 去语义
 * 骨架与样板区 → 挑主容器 → 转文本 → 行级清洗（丢样板行、去重、压空行）→
 * 出报告。留下来的趋势是「像正文的行」，摘录也就跟着准了。
 *
 * ## 三种模式
 *
 *   article （默认）去样板 + 挑主容器 + 行级清洗。给网页正文用。
 *   plain           只在文本上做行级清洗（去重、压空行、丢样板行）。
 *                   给「来源不是 HTML」或结构判断不了时的退路。
 *   off             原样返回（等于回到第十八轮的行为）。给「宁可多也不能少」
 *                   的场景，例如页面本身就是一份清单。
 *
 * ## 为什么不是 Readability
 *
 * 零依赖是本插件的一条硬约束（第六轮起），而真正的正文抽取要一棵 DOM 树与
 * 一套评分模型。这套启发式做不到「任何页面都准」，但**每一条判据都可解释、
 * 可回滚**：报告里给出「丢了几行、去重几行、剩下多少字符」，判错时看得见，
 * 也能一键退回 `off`。
 *
 * ## 与哨兵的分工
 *
 * 成败判定（拦截页 / 区域封锁页 / 空壳）仍然用**清洗前**的文本，见 `web.js`
 * 的 `inspectFetchedPage`：清洗会把拦截页的样板也删掉，拿清洗后的文本去判
 * 会把「目标站拒绝」误判成「正文几乎为空」，而这两类失败的修法完全不同。
 *
 * @module dsh-office-mode/web-preprocess
 */

import { asText } from './web-errors.js';

/** 预处理模式。 */
export const PREPROCESS_MODES = Object.freeze(['article', 'plain', 'off']);

/** 预处理默认参数（与 config.js 的 DEFAULT_PREPROCESS 同源同义）。 */
export const DEFAULT_PREPROCESS = Object.freeze({
    mode: 'article',
    /** 行级清洗：丢掉样板行（cookie / 登录 / 订阅 / 版权 …）。 */
    dropBoilerplate: true,
    /** 行级清洗：同一条重复出现的行只留第一次。 */
    dedupeLines: true,
    /** 短于这个长度的行，若也不像句子（不以句末标点收尾、不是标题/列表），就当导航条目丢掉。 */
    minLineChars: 12,
    /** 保留标题与出处元信息。 */
    keepTitle: true,
});

/**
 * 字符串/枚举/数值一律收敛，坏值退回默认（配置写坏不该弄崩插件）。
 * 与 `web.js` 的 `resolveBuiltinOptions` 同一取向。
 */
export function resolvePreprocessOptions(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const min = Number.isFinite(source.minLineChars) ? Math.trunc(source.minLineChars) : DEFAULT_PREPROCESS.minLineChars;
    return {
        mode: PREPROCESS_MODES.includes(source.mode) ? source.mode : DEFAULT_PREPROCESS.mode,
        dropBoilerplate: typeof source.dropBoilerplate === 'boolean' ? source.dropBoilerplate : DEFAULT_PREPROCESS.dropBoilerplate,
        dedupeLines: typeof source.dedupeLines === 'boolean' ? source.dedupeLines : DEFAULT_PREPROCESS.dedupeLines,
        minLineChars: Math.min(80, Math.max(0, min)),
        keepTitle: typeof source.keepTitle === 'boolean' ? source.keepTitle : DEFAULT_PREPROCESS.keepTitle,
    };
}

// ── 结构层 ──────────────────────────────────────────────────────────────────

/**
 * 整段丢掉（连同内容）的标签：脚本、样式、模板、图元、表单控件。
 *
 * 这些「一定不是正文」的标签在浏览器里也要求成对出现，所以可以先用配对正则
 * 快清一遍；清理时一律同时兜住「没闭合」的情形（只剥标签、留内容），
 * 免得一口吃掉半篇文档。
 */
const NON_CONTENT_TAGS = ['script', 'style', 'noscript', 'template', 'svg', 'canvas', 'iframe', 'form', 'select', 'textarea', 'button', 'dialog'];

/** 语义骨架：整段丢掉（导航 / 页头 / 页脚 / 侧栏）。 */
const SKELETON_TAGS = ['nav', 'header', 'footer', 'aside'];

/** 可能承载样板区的容器标签（按属性判）。 */
const CONTAINER_TAGS = 'div|section|ul|ol|span|p|table|li|a';

/**
 * 样板区的 class / id 特征。
 *
 * 判据取「栏目名」而不是「排版名」：sidebar / comment / cookie / consent /
 * advert / banner / promo / share / related / recommend / subscribe /
 * breadcrumb / pagination / toolbar / menu / footer / header / nav。
 * 这些词出现在类名或 id 里时，八成不是正文。
 */
const BOILERPLATE_ATTR = /(?:^|[\s"'/])(?:id|class)\s*=\s*("[^"]{0,240}"|'[^']{0,240}'|[^\s>]{0,240})/i;
const BOILERPLATE_WORDS = /(?:^|[-_. ])(?:nav|navbar|menu|sidebar|side-bar|footer|header|comment|comments|reply|cookie|consent|gdpr|advert|ad|ads|banner|promo|share|sharing|social|related|recommend|subscribe|newsletter|breadcrumb|pagination|pager|toolbar|masthead|skip-link|back-to-top|sitemap|site-map|language|lang-switch)(?:[-_. ]|$)/i;

/** 一个开标签的属性里带没带样板词。单测直接钉这一条。 */
export function looksLikeBoilerplateTag(tag) {
    const attrs = BOILERPLATE_ATTR.exec(String(tag ?? ''));
    if (attrs === null) return false;
    return BOILERPLATE_WORDS.test(attrs[1].replace(/["']/g, ' '));
}

/**
 * 找出一段 HTML 里某个标签的配对区间。
 *
 * 不用配对正则（`<div>…</div>` 遇到嵌套 div 会提前收口），而是从头扫标签、
 * 对同名标签做深度计数。**找不到闭合标签时返回 null**，调用方据此退化成
 * 「只剥标签、保留内容」—— 宁可留下噪声，也不吃掉半篇正文。
 *
 * @param {string} html
 * @param {number} start 开标签的起始下标
 * @returns {{openEnd: number, closeStart: number, closeEnd: number, name: string}|null}
 */
function matchRegion(html, start) {
    const open = /^<([a-zA-Z][a-zA-Z0-9:-]*)\b[^>]{0,4000}>/.exec(html.slice(start, start + 4_100));
    if (open === null) return null;
    const name = open[1].toLowerCase();
    const openEnd = start + open[0].length;
    if (open[0].endsWith('/>')) return { openEnd, closeStart: openEnd, closeEnd: openEnd, name };
    const tagRe = new RegExp('<' + name + '\\b|</' + name + '\\s*>', 'gi');
    tagRe.lastIndex = openEnd;
    let depth = 1;
    for (;;) {
        const matched = tagRe.exec(html);
        if (matched === null) return null;
        if (matched[0].charAt(1) === '/') {
            depth -= 1;
            if (depth === 0) return { openEnd, closeStart: matched.index, closeEnd: matched.index + matched[0].length, name };
        } else {
            depth += 1;
        }
    }
}

/**
 * 按开标签特征丢掉整段区域（含内容）。
 *
 * `accept` 给定时用它筛开标签（样板区那一条）；不给就全收（语义骨架那一条）。
 * 每段处理完把扫描位置挪到替换处，不从头重扫 —— 2 MB 的页面也扛得住。
 *
 * @returns {{text: string, dropped: number}}
 */
function dropRegions(html, opener, accept = null, budget = 600) {
    let text = html;
    let dropped = 0;
    let scanned = 0;
    opener.lastIndex = 0;
    let matched = opener.exec(text);
    while (matched !== null && dropped < budget && scanned < budget * 4) {
        scanned += 1;
        if (accept !== null && !accept(matched[0])) {
            opener.lastIndex = matched.index + matched[0].length;
            matched = opener.exec(text);
            continue;
        }
        const region = matchRegion(text, matched.index);
        if (region === null) {
            // 没闭合：只剥开标签，内容留下。
            text = text.slice(0, matched.index) + ' ' + text.slice(matched.index + matched[0].length);
        } else {
            text = text.slice(0, matched.index) + ' ' + text.slice(region.closeEnd);
        }
        dropped += 1;
        opener.lastIndex = matched.index;
        matched = opener.exec(text);
    }
    return { text, dropped };
}

/** 去注释、去非内容标签、去语义骨架、去样板区。返回清洗后的 HTML 与统计。 */
function stripNoise(html) {
    let text = String(html ?? '');
    text = text.replace(/<!--[\s\S]*?-->/g, ' ');
    text = text.replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, ' ');
    for (const name of NON_CONTENT_TAGS) {
        text = text.replace(new RegExp('<' + name + '\\b[^>]*>[\\s\\S]*?<\\/' + name + '\\s*>', 'gi'), ' ');
        text = text.replace(new RegExp('<' + name + '\\b[^>]*>', 'gi'), ' ');
        text = text.replace(new RegExp('<\\/' + name + '\\s*>', 'gi'), ' ');
    }
    let regions = 0;
    for (const name of SKELETON_TAGS) {
        const result = dropRegions(text, new RegExp('<' + name + '\\b[^>]*>', 'gi'));
        text = result.text;
        regions += result.dropped;
    }
    const boilerplate = dropRegions(
        text,
        new RegExp('<(?:' + CONTAINER_TAGS + ')\\b[^>]*>', 'gi'),
        looksLikeBoilerplateTag,
    );
    return { html: boilerplate.text, regions: regions + boilerplate.dropped };
}

// ── 主容器 ──────────────────────────────────────────────────────────────────

/**
 * 挑主容器：`<article>` / `<main>` / `role="main"` / `id|class` 带
 * article|post|content|entry|main 的容器里**文本最长**的那一块；一个都没有
 * 就返回整篇。
 *
 * 判据只用「剥完标签后的可见字符数」—— 不做链接密度加权：只有正则的情况下
 * 那个权重很容易把短正文页面判给侧栏。挑错时正文仍在（只是噪声多），
 * 也就是退回第十八轮的水平，不会丢内容。
 */
function pickMainRegion(html) {
    const candidates = [];
    const openers = [
        /<article\b[^>]*>/gi,
        /<main\b[^>]*>/gi,
        /<(?:div|section)\b[^>]*\brole\s*=\s*"?main"?[^>]*>/gi,
        /<(?:div|section)\b[^>]*\b(?:id|class)\s*=\s*"[^"]*(?:article|post|content|entry|main)[^"]*"[^>]*>/gi,
    ];
    for (const opener of openers) {
        opener.lastIndex = 0;
        let matched = opener.exec(html);
        let count = 0;
        while (matched !== null && count < 40) {
            const region = matchRegion(html, matched.index);
            if (region !== null && region.closeEnd > region.openEnd) {
                candidates.push(html.slice(region.openEnd, region.closeStart));
                count += 1;
                opener.lastIndex = region.closeEnd;
            } else {
                opener.lastIndex = matched.index + matched[0].length;
            }
            matched = opener.exec(html);
        }
    }
    if (candidates.length === 0) return { html, picked: false };
    let best = '';
    let bestLength = 0;
    for (const candidate of candidates) {
        const length = candidate.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().length;
        if (length > bestLength) {
            bestLength = length;
            best = candidate;
        }
    }
    return { html: best, picked: true };
}

// ── 元信息 ──────────────────────────────────────────────────────────────────

/** 只解最常见的几个实体：预处理是「降噪」不是「还原」，多解没意义。 */
function decodeLite(text) {
    return String(text)
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&quot;/gi, '"')
        .replace(/&#39;|&apos;/gi, "'")
        .replace(/&#(\d+);/g, (whole, code) => {
            const value = Number.parseInt(code, 10);
            return Number.isFinite(value) ? String.fromCodePoint(value) : whole;
        });
}

/** 从 HTML 里取标题、站点名、作者、发布时间、语言。取不到就给空串，不猜。 */
export function extractMeta(html) {
    const text = String(html ?? '');
    const pick = (pattern) => {
        const matched = pattern.exec(text);
        return matched === null ? '' : decodeLite(matched[1]).replace(/\s+/g, ' ').trim();
    };
    return {
        title: pick(/<title[^>]*>([\s\S]{0,400}?)<\/title\s*>/i)
            || pick(/<meta[^>]+property\s*=\s*"og:title"[^>]+content\s*=\s*"([^"]{0,400})"/i),
        site: pick(/<meta[^>]+property\s*=\s*"og:site_name"[^>]+content\s*=\s*"([^"]{0,200})"/i),
        byline: pick(/<meta[^>]+name\s*=\s*"author"[^>]+content\s*=\s*"([^"]{0,200})"/i),
        publishedAt: pick(/<meta[^>]+property\s*=\s*"article:published_time"[^>]+content\s*=\s*"([^"]{0,80})"/i)
            || pick(/<time[^>]+datetime\s*=\s*"([^"]{0,80})"/i),
        lang: pick(/<html[^>]+lang\s*=\s*"([^"]{0,20})"/i),
    };
}

// ── HTML → 纯文本 ───────────────────────────────────────────────────────────

/**
 * HTML → 纯文本（**不带链接标记**）。
 *
 * 与 `web.js` 的 `htmlToText` 的差别只有一处，但很关键：链接不再写成
 * `[文字](地址)`。预处理的产物是「人读得懂、模型不必看到 URL」的正文，
 * 而几百个导航链接收成文本后主要贡献噪声；URL 该出现在来源清单里，
 * 不该混进正文。跳转锚与 `javascript:` 链接的内容照留、地址丢掉。
 */
export function htmlToPlainText(html) {
    let text = String(html ?? '');
    text = text.replace(/<a\b[^>]*>([\s\S]{0,2000}?)<\/a\s*>/gi, ' $1 ');
    text = text.replace(/<h([1-6])\b[^>]*>/gi, (_whole, level) => '\n\n' + '#'.repeat(Math.min(3, Number(level))) + ' ');
    text = text.replace(/<\/(h[1-6]|p|div|section|article|main|blockquote|table|tr|ul|ol|dl|pre|figure)\s*>/gi, '\n');
    text = text.replace(/<br\s*\/?>/gi, '\n');
    // 列表项只在开标签处断行：`</li>` 再断一次会让每条之间多一个空行（`- 一\n\n- 二`）。
    text = text.replace(/<li\b[^>]*>/gi, '\n- ');
    text = text.replace(/<t[dh]\b[^>]*>/gi, ' | ');
    text = text.replace(/<[^>]*>/g, ' ');
    return decodeLite(text);
}

// ── 行级清洗 ────────────────────────────────────────────────────────────────

/**
 * 样板行：cookie 与同意条、登录注册、订阅关注、分享、版权与免责、相关推荐、
 * 客户端下载、隐私政策 …… 中英各一套。这是**行级**判据，比结构判据更保守：
 * 只影响以这些词打头的行。
 */
const BOILERPLATE_LINE = /^(?:cookie|we use cookies|this (?:site|website) uses|accept (?:all )?cookies|by (?:continuing|using)|manage (?:cookies|preferences)|privacy policy|terms of (?:service|use)|all rights reserved|sign ?(?:in|up)|log ?in|register|subscribe|newsletter|follow us|share (?:this|on)|related (?:posts|articles)|recommended|read more|advertisement|sponsored|skip to (?:content|main)|back to top|copyright|©|版权所有|免责声明|隐私政策|用户协议|服务条款|登录|注册|订阅|关注我们|分享到|扫码|下载客户端|相关内容|相关阅读|推荐阅读|热门推荐|上一篇|下一篇|返回顶部|京ICP|沪ICP|粤ICP)/i;

/** 一句话里出现句末标点，说明它更像正文而不是栏目名。 */
const SENTENCE_END = /[.!?。！？…；;：:）)】」』"'”’]$/;

/**
 * 行级清洗：压空白 → 丢样板行 → 丢「短且不像句子」的导航行 → 去重 → 压空行。
 *
 * 为什么不一律丢掉短行：合法的正文里也有短行（人名、数字、引语）。所以短行
 * 只在**同时**不像句子时才丢；`minLineChars` 调到 0 就关掉这条。
 */
export function cleanLines(text, options) {
    const settings = resolvePreprocessOptions(options);
    const inputLines = String(text ?? '').replace(/\r\n?/g, '\n').split('\n');
    const kept = [];
    const seen = new Set();
    let dropped = 0;
    let deduped = 0;
    for (const raw of inputLines) {
        const line = raw.replace(/[ \t\u00a0]+/g, ' ').trim();
        if (line === '') {
            kept.push('');
            continue;
        }
        if (settings.dropBoilerplate && BOILERPLATE_LINE.test(line)) {
            dropped += 1;
            continue;
        }
        if (line.length < settings.minLineChars && !SENTENCE_END.test(line) && !/^[#\-|]/.test(line)) {
            dropped += 1;
            continue;
        }
        if (settings.dedupeLines) {
            const key = line.replace(/^#+\s*/, '').toLowerCase();
            if (key.length >= 6 && seen.has(key)) {
                deduped += 1;
                continue;
            }
            if (key.length >= 6) seen.add(key);
        }
        kept.push(line);
    }
    const joined = kept.join('\n').replace(/\n{3,}/g, '\n\n').trim();
    return { text: joined, dropped, deduped };
}

// ── 统一入口 ────────────────────────────────────────────────────────────────

/** 段落数：非空行的粗略口径，只用于报告。 */
function countParagraphs(text) {
    return String(text ?? '').split('\n').filter((line) => line.trim() !== '').length;
}

/** 去掉的比例（0–1，四位小数够读）。 */
function ratio(input, output) {
    if (input <= 0) return 0;
    return Math.round(Math.max(0, Math.min(1, 1 - output / input)) * 10_000) / 10_000;
}

/**
 * 预处理一个页面。
 *
 * @param {{html?: string, text?: string, url?: string, kind?: 'html'|'text'}} input
 *        `html` 与 `text` 都给时以 `html` 为准（结构判据只有 HTML 才有）。
 * @param {object} [rawOptions] `search.preprocess` 配置
 * @returns {{mode: string, text: string, meta: object, stats: object}}
 */
export function preprocessPage(input, rawOptions) {
    const settings = resolvePreprocessOptions(rawOptions);
    const html = asText(input?.html);
    // 只给了 HTML 没给 text 时（调用方漏传，或 plain/off 模式下），仍然从 HTML 折出文本 ——
    // 否则「有内容却返回空」这种错误很难从报告里看出来。
    const text = String(input?.text ?? '') || (html === '' ? '' : htmlToPlainText(html));
    const source = html !== '' ? html : text;
    const sourceChars = source.length;
    const meta = settings.keepTitle ? extractMeta(html) : extractMeta('');
    if (settings.mode === 'off' || sourceChars === 0) {
        return {
            mode: settings.mode,
            text,
            meta,
            stats: { inputChars: sourceChars, outputChars: text.length, removedRatio: 0, droppedLines: 0, dedupedLines: 0, regions: 0, mainRegion: false, paragraphs: countParagraphs(text) },
        };
    }
    if (settings.mode === 'plain' || html === '') {
        const cleaned = cleanLines(text, settings);
        return {
            mode: settings.mode,
            text: cleaned.text,
            meta,
            stats: {
                inputChars: sourceChars,
                outputChars: cleaned.text.length,
                removedRatio: ratio(sourceChars, cleaned.text.length),
                droppedLines: cleaned.dropped,
                dedupedLines: cleaned.deduped,
                regions: 0,
                mainRegion: false,
                paragraphs: countParagraphs(cleaned.text),
            },
        };
    }
    const stripped = stripNoise(html);
    const main = pickMainRegion(stripped.html);
    const cleaned = cleanLines(htmlToPlainText(main.html), settings);
    return {
        mode: settings.mode,
        text: cleaned.text,
        meta,
        stats: {
            inputChars: sourceChars,
            outputChars: cleaned.text.length,
            removedRatio: ratio(sourceChars, cleaned.text.length),
            droppedLines: cleaned.dropped,
            dedupedLines: cleaned.deduped,
            regions: stripped.regions,
            mainRegion: main.picked,
            paragraphs: countParagraphs(cleaned.text),
        },
    };
}
