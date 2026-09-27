
/**
 * 结果解析：把子代理写下的检索结果文件读成结构化发现。
 *
 * 为什么要有这一步：子代理检索回来的原始材料（几十条结果、整页正文）
 * 如果直接回给主会话，主上下文会被一次检索灌满，之后真正要做的文档
 * 排版与措辞就没有余地了。所以约定是——子代理把结果写成文件，主会话
 * 只交文件路径，在这里读进来、按渠道归类、算出覆盖度与风险，
 * 主会话只拿到几十行的摘要。
 *
 * 解析刻意做得宽容：子代理写的是 Markdown，不是严格 schema。抽不出的
 * 东西进入 notes 而不是报错——宁可让模型看到「这条没带上 URL」，
 * 也不要因为格式不标准就把整份结果判失败。
 *
 * @module dsh-office-mode/findings
 */

import { contentType, guessContentType } from './search-routes.js';

/** 常见小标题别名 → 渠道通称。子代理写法各异，这里做一层归一。 */
const HEADING_ALIASES = [
    { match: /(泛搜|概览|overview|初步)/, channel: '泛搜' },
    { match: /(权威|官方媒体|人民日报|新华社|通稿|通报|媒体)/, channel: '权威媒体' },
    { match: /(百科|wiki|维基)/i, channel: '百科' },
    { match: /(文献|论文|学术|研究|paper)/i, channel: '文献' },
    { match: /(官方|文档|reference|手册|manual)/i, channel: '官方文档' },
    { match: /(视频|bilibili|哔哩)/i, channel: '社交平台：视频' },
    { match: /(社区|论坛|v2ex|知乎|讨论)/, channel: '社交平台：社区' },
    { match: /(国际|reddit|外网)/i, channel: '社交平台：国际' },
    { match: /(问答|stackoverflow|stack)/i, channel: '知识问答' },
    { match: /(仓库|github|release)/i, channel: '官方仓库' },
];

function channelOf(heading) {
    const text = String(heading ?? '');
    for (const item of HEADING_ALIASES) {
        if (item.match.test(text)) return item.channel;
    }
    return text.trim() === '' ? '未分组' : text.trim();
}

/** 从 Markdown 里抽 (标题, 正文块) 序列。只认 # 到 #### 四级标题。 */
function splitSections(markdown) {
    const lines = String(markdown ?? '').split(/\r?\n/);
    const sections = [];
    let current = { heading: '', lines: [] };
    for (const line of lines) {
        const heading = /^#{1,4}\s+(.*)$/.exec(line);
        if (heading !== null) {
            sections.push(current);
            current = { heading: heading[1].trim(), lines: [] };
            continue;
        }
        current.lines.push(line);
    }
    sections.push(current);
    return sections.filter((section) => section.heading !== '' || section.lines.some((line) => line.trim() !== ''));
}

/**
 * 抽一行里的 URL（Markdown 链接或裸链）。
 *
 * 用一条通吃的正则再统一剥尾，而不是加 lookbehind 排除「前面是括号」——
 * 中文行里 URL 常常正好写在 （） 里，lookbehind 会把最常见的那种写法全部漏掉。
 */
function urlsIn(line) {
    const found = [];
    const any = /https?:\/\/[^\s)\]}>，。；、"'）】]+/g;
    let m;
    while ((m = any.exec(line)) !== null) {
        // 剥掉链接末尾常见的标点，保留路径本身。
        const url = m[0].replace(/[.,;:!?、，。；：！？]+$/, '');
        if (url !== '' && !found.includes(url)) found.push(url);
    }
    return found;
}

/** 从域取一个用于「来源是否同源」判断的粗标识。 */
function hostOf(url) {
    try {
        return new URL(url).host.replace(/^www\./, '').toLowerCase();
    } catch {
        return '';
    }
}

/**
 * 提纲里的一个渠道是否被结果里的某个分组覆盖。
 *
 * 判定按「关键区分词」做：手册类的官方文档只看「官方文档 / 手册 / reference」，
 * 不能因为分组名叫「官方仓库」就算覆盖；热点类的权威媒体与它的「回看更长
 * 时间窗」变体要能分别认领。
 */
function channelMatches(channel, sectionName, rawHeading) {
    const name = String(sectionName ?? '');
    const raw = String(rawHeading ?? '');
    if (name === '' || name === '未分组') return false;
    if (channel.platform) {
        // 平台渠道：归一后的分组名或原始小标题里出现平台名即可——
        // 子代理常把标题写成「社交平台：中文社区讨论（platform: v2ex）」，
        // 归一化会把平台名抹掉，所以原始标题必须一起看。
        const needle = channel.platform.toLowerCase();
        if (name.toLowerCase().includes(needle) || raw.toLowerCase().includes(needle)) return true;
    }
    const key = channel.kind.replace(/（[^）]*）/g, '').trim();
    if (key.includes('泛搜')) return name.includes('泛搜') || name.includes('概览') || name.includes('初步');
    if (key.includes('权威媒体')) {
        const broad = channel.timeRange === 'month';
        const base = name.includes('权威媒体') || name.includes('媒体') || name.includes('通报');
        return broad ? base : base && !name.includes('回看');
    }
    if (key.includes('官方参考文档') || key === '官方文档') {
        return name.includes('官方文档') || name.includes('官方参考') || name.includes('手册')
            || name.toLowerCase().includes('reference');
    }
    if (key.includes('官方仓库')) return name.includes('仓库') || name.toLowerCase().includes('github');
    if (key.includes('文献')) return name.includes('文献');
    if (key.includes('百科')) return name.includes('百科') || name.toLowerCase().includes('wiki');
    if (key.includes('知识问答')) return name.includes('问答') || name.toLowerCase().includes('stack');
    if (key.includes('正文核实')) return name.includes('核实') || name.includes('正文');
    return name.includes(key);
}

/** 取中文 2-gram + 英文词，用于判断两条结论是否在讲同一件事。 */
function tokensOf(text) {
    const clean = String(text ?? '').toLowerCase().replace(/[\s，。、；：！？,.;:!?"'（）()\[\]【】]+/g, ' ');
    const tokens = new Set();
    for (const word of clean.split(' ')) {
        if (word === '') continue;
        if (/^[a-z0-9]+$/.test(word)) {
            if (word.length >= 3) tokens.add(word);
            continue;
        }
        for (let i = 0; i + 2 <= word.length; i += 1) tokens.add(word.slice(i, i + 2));
    }
    return tokens;
}

/** 两条结论的字面重合度（Jaccard）。 */
function similarity(a, b) {
    const sa = tokensOf(a);
    const sb = tokensOf(b);
    if (sa.size === 0 || sb.size === 0) return 0;
    let shared = 0;
    for (const token of sa) if (sb.has(token)) shared += 1;
    return shared / (sa.size + sb.size - shared);
}

/**
 * 跨源核对：把讲同一件事的结论归成一簇，看这一簇背后有几个互不相关的站点。
 * 只有一簇、却挂着多个站点的说法才算「已核对」；一簇里只有一个站点的，
 * 就是「单一来源」——这正是茧房效应最容易骗过人的地方：搜出来一堆结果，
 * 其实全是同一篇稿子的转载。
 */
function corroboration(findings) {
    // 用并查集把「讲同一件事」的结论连成簇。
    // 逐个比对 reference 的写法有个坏处：谁先出现谁当代表，A≈B、B≈C 但 A≉C
    // 时 C 会被拆成单独一簇，于是一件被三个来源确认的事被算成两组。
    const parent = findings.map((_, index) => index);
    const find = (index) => {
        let at = index;
        while (parent[at] !== at) {
            parent[at] = parent[parent[at]];
            at = parent[at];
        }
        return at;
    };
    const union = (a, b) => {
        const ra = find(a);
        const rb = find(b);
        if (ra !== rb) parent[rb] = ra;
    };
    for (let i = 0; i < findings.length; i += 1) {
        for (let j = i + 1; j < findings.length; j += 1) {
            if (similarity(findings[i].claim, findings[j].claim) >= 0.5) union(i, j);
        }
    }
    const grouped = new Map();
    for (const [index, item] of findings.entries()) {
        const root = find(index);
        if (!grouped.has(root)) grouped.set(root, { reference: item.claim, items: [], hosts: new Set() });
        const cluster = grouped.get(root);
        cluster.items.push(item);
        for (const host of item.hosts) cluster.hosts.add(host);
    }
    const clusters = [...grouped.values()];
    const corroborated = clusters.filter((cluster) => cluster.hosts.size >= 2).length;
    const unverified = clusters.filter((cluster) => cluster.hosts.size < 2).length;
    return { clusters, corroborated, unverified };
}

/**
 * 把一份检索结果 Markdown 解析成 { findings, coverage, risks, notes }。
 *
 * @param {string} markdown 子代理写下的结果文件内容
 * @param {string} [typeId] 内容类型；省略时按主题文字猜
 * @param {string} [topic]  主题，用来猜类型与写摘要
 */
export function parseFindings(markdown, typeId, topic) {
    const text = String(markdown ?? '');
    const seen = new Map();
    const findings = [];
    const notes = [];

    for (const section of splitSections(text)) {
        const channel = channelOf(section.heading);
        const rawHeading = String(section.heading ?? '');
        for (const rawLine of section.lines) {
            const line = rawLine.trim();
            if (line === '') continue;
            if (/^[|>]/.test(line)) continue;
            const bullet = /^(?:[-*+]|\d+[.)])\s+(.*)$/.exec(line);
            if (bullet === null) continue;
            const claim = bullet[1].trim();
            if (claim === '') continue;
            const urls = urlsIn(claim);
            // 把 URL 从正文里摘掉：结论本身与出处要分开，引用时不该把长链
            // 一起抄进文档。末尾的括号链接、行尾裸链都算出处。
            const cleanClaim = claim
                .replace(/\[[^\]]*\]\(https?:\/\/[^)\s]+\)/g, ' ')
                .replace(/\(\s*https?:\/\/[^)\s]+\s*\)/g, ' ')
                .replace(/(?<![(\w])https?:\/\/[^\s)>\]]+/g, ' ')
                .replace(/\s*[（(]\s*[)）]\s*$/g, '')
                .replace(/[\s,，;；:：、-]+$/g, '')
                .replace(/\s+/g, ' ')
                .trim();
            if (cleanClaim === '') continue;
            const key = cleanClaim.slice(0, 160);
            if (seen.has(key)) {
                // 同一说法在多个渠道被重复搜到：合并出处，并**重算 hosts**。
                // 只 push URL 不重算 hosts，会让「多站点背书」永远算不出来 ——
                // 明明两个来源都提到的事，仍被判成单一来源。
                const existing = seen.get(key);
                for (const url of urls) if (!existing.urls.includes(url)) existing.urls.push(url);
                existing.hosts = existing.urls.map(hostOf).filter((host) => host !== '');
                continue;
            }
            const item = {
                channel,
                heading: rawHeading,
                claim: cleanClaim,
                urls,
                hosts: urls.map(hostOf).filter((host) => host !== ''),
                singleSource: urls.length <= 1,
            };
            seen.set(key, item);
            findings.push(item);
        }
    }

    // 覆盖度：提纲里标了 ★ 的渠道是否都出现过。
    const type = contentType(typeId) ?? guessContentType(topic ?? text);
    const present = new Set(findings.map((item) => item.channel));
    const covered = [];
    const missing = [];
    for (const channel of type.channels) {
        const label = channel.platform ? channel.kind + '（' + channel.platform + '）' : channel.kind;
        // 命中判定要窄：提纲里的一个渠道只能被一个分组认领，否则
        // 「权威媒体」与「权威媒体（回看更长时间窗）」会互相顶替，
        // 覆盖度就永远是满的，漏掉的渠道看不出来。
        const hit = findings.some((item) => channelMatches(channel, item.channel, item.heading));
        if (hit) covered.push(label);
        else if (channel.required) missing.push(label);
    }

    // 风险：只有单一来源的结论、以及完全没带 URL 的行。
    const noUrl = findings.filter((item) => item.urls.length === 0).length;
    const hosts = new Set(findings.flatMap((item) => item.hosts));
    const crossSource = corroboration(findings);
    // 单一来源要按「簇」判：同一件事的多条转载合并成簇后，看这一簇有几个站点。
    for (const cluster of crossSource.clusters) {
        const verified = cluster.hosts.size >= 2;
        for (const item of cluster.items) {
            item.verified = verified;
            item.singleSource = !verified;
        }
    }
    const singleSource = findings.filter((item) => item.singleSource).length;
    if (noUrl > 0) {
        notes.push(noUrl + ' 条结论没有带来源 URL，引用前必须补齐出处。');
    }
    if (singleSource > 0) {
        notes.push(singleSource + ' 条结论只有单一来源，写进正文时要标注「单一来源」。');
    }
    if (findings.length > 0 && crossSource.unverified > 0) {
        notes.push(crossSource.unverified + ' 组结论只有单一来源站点背书，正文里必须写「单一来源」或再补一轮检索。');
    }
    if (findings.length > 0 && hosts.size <= 2) {
        notes.push('全部结论只来自 ' + hosts.size + ' 个站点，存在茧房风险——换渠道再补一轮。');
    }
    if (missing.length > 0) {
        notes.push('提纲里必须覆盖的渠道还缺：' + missing.join('、') + '。');
    }
    if (findings.length === 0) {
        notes.push('没有解析出任何结论条目：确认文件是 Markdown，且每条结论是 - 或 1. 开头的列表项。');
    }

    return {
        topic: String(topic ?? '').trim(),
        type,
        findings,
        coverage: {
            channels: [...present].filter((name) => name !== '未分组'),
            covered,
            missing,
            hosts: [...hosts],
        },
        risks: {
            total: findings.length,
            singleSource,
            noUrl,
            distinctHosts: hosts.size,
            corroborated: crossSource.corroborated,
            unverified: crossSource.unverified,
        },
        notes,
    };
}

/** 把解析结果渲染成给主会话看的紧凑摘要（解析的产物只以这个形式进上下文）。 */
export function renderFindings(result) {
    const lines = [];
    const name = result.type?.name ?? '未指明';
    lines.push('检索结果解析（' + name + (result.topic ? '：' + result.topic : '') + '）');
    lines.push('到手结论 ' + result.risks.total + ' 条｜来源站点 ' + result.risks.distinctHosts
        + ' 个｜已跨源核对 ' + result.risks.corroborated + ' 组｜待核 ' + result.risks.unverified
        + ' 组｜缺 URL ' + result.risks.noUrl + ' 条');
    const covered = result.coverage.covered;
    if (covered.length > 0) lines.push('已覆盖渠道：' + covered.join('、'));
    if (result.coverage.missing.length > 0) lines.push('★ 仍缺渠道：' + result.coverage.missing.join('、'));
    lines.push('');

    let index = 0;
    for (const item of result.findings) {
        index += 1;
        const mark = item.singleSource ? ' [单一来源]' : '';
        lines.push(index + ') [' + item.channel + ']' + mark + ' ' + item.claim.slice(0, 220));
        for (const url of item.urls.slice(0, 3)) lines.push('   ' + url);
        const cap = Number.isFinite(result.limit) ? Math.max(1, Math.trunc(result.limit)) : 40;
        if (index >= cap) {
            lines.push('…其余 ' + (result.findings.length - index) + ' 条见结果文件');
            break;
        }
    }

    if (result.notes.length > 0) {
        lines.push('', '需要注意：');
        for (const note of result.notes) lines.push('- ' + note);
    }
    return lines.join('\n');
}