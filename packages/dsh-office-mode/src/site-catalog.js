/**
 * 站点清单：内置目录 + 选择 + 域名匹配。
 *
 * ## 要解决的事
 *
 * 用户的原话是「搜索时优先选择检索这些网站的内容…如果被墙了就不纠结于这些信息来源」。
 * 落到插件里就是两件事：
 *
 *   1. **目录**：一份可按类型分组的站点清单（学术 / 图书 / 代码 / 自定义），
 *      由设置页的可视化列表编辑（服务端 schema 的默认值就是这份内置目录）。
 *   2. **优先**：检索时先把查询**限定到清单里的站点**去问；命中的来源排在前面并按
 *      类型汇总反馈。限定轮一无所获时（被墙、没收录、站点改版）**退回不限定来源**
 *      的泛搜，并把「哪些站点空手」照实写出来 —— 不因为清单查不到就当成「没有资料」。
 *
 * ## 为什么单独一个模块
 *
 * 目录是**数据**，选择与匹配是**纯函数**，两者都不该长在 tools.js 里：
 * 设置页的浏览器半侧要镜像同一份目录（`lib/client.js` 的 `SITE_CATALOG` 与
 * `SITE_TYPES`），漂移由 `test/client.mjs` 的逐条比对断言守住；检索侧的限定与匹配
 * 由 `test/site-catalog.mjs` 钉住；tools.js 只负责「怎么用这些结果写反馈」。
 *
 * ## 默认开关的口径（第三十四轮，用户点明）
 *
 * 影子图书馆类（sci-hub 镜像、z-lib、Anna's Archive）**只进目录、默认关闭** ——
 * 需要时在设置页打开或按次点名（`sites: ['sci-hub.se']`）。其余为公开的预印本库、
 * 引用索引、开放书库、代码仓库与问答站，默认开启。
 *
 * @module dsh-office-mode/site-catalog
 */

/** 类型表：id 进配置，name/note 给人看（设置页与反馈共用一份）。 */
export const SITE_TYPES = Object.freeze([
    { id: 'academic', name: '学术', note: '预印本、引用索引、文献库与文献检索' },
    { id: 'book', name: '图书', note: '开放书库、书目与电子书' },
    { id: 'code', name: '代码', note: '代码仓库、问答与包索引' },
    { id: 'custom', name: '自定义', note: '自己加的站点（只按域名限定）' },
]);

/** 合法类型 id（schema 与选择逻辑共用）。 */
export const SITE_TYPE_IDS = Object.freeze(SITE_TYPES.map((type) => type.id));

/**
 * 内置目录。字段：
 *   type     类型 id（见 SITE_TYPES）
 *   domain   域名，**不带协议与路径**（限定查询拼成 `site:<domain>`）
 *   label    显示名（设置页、反馈与结果文件里都用它）
 *   note     一句话说明（为什么收它 / 要注意什么）
 *   enabled  默认是否参与优先检索
 *
 * 顺序就是默认优先级顺序：设置页与选择逻辑都按数组顺序取。
 */
export const BUILTIN_SITES = Object.freeze([
    // ── 学术 ──────────────────────────────────────────────────────────────
    { type: 'academic', domain: 'arxiv.org', label: 'arXiv', note: '预印本，理工科一手材料', enabled: true },
    { type: 'academic', domain: 'scholar.google.com', label: 'Google 学术', note: '覆盖广，通常要代理才通', enabled: true },
    { type: 'academic', domain: 'cnki.net', label: '中国知网', note: '中文期刊与学位论文（多为摘要页）', enabled: true },
    { type: 'academic', domain: 'crossref.org', label: 'CrossRef', note: 'DOI 与引用元数据，核实出处最省事', enabled: true },
    { type: 'academic', domain: 'semanticscholar.org', label: 'Semantic Scholar', note: '语义检索 + 引用关系', enabled: true },
    { type: 'academic', domain: 'pubmed.ncbi.nlm.nih.gov', label: 'PubMed', note: '生物医学文献库', enabled: true },
    { type: 'academic', domain: 'researchgate.net', label: 'ResearchGate', note: '作者自存的全文，质量参差', enabled: true },
    { type: 'academic', domain: 'sciencedirect.com', label: 'ScienceDirect', note: '付费墙，通常只有摘要', enabled: false },
    { type: 'academic', domain: 'springer.com', label: 'Springer', note: '付费墙，通常只有摘要', enabled: false },
    { type: 'academic', domain: 'ieee.org', label: 'IEEE Xplore', note: '付费墙，通常只有摘要', enabled: false },
    { type: 'academic', domain: 'sci-hub.se', label: 'Sci-Hub', note: '第三方镜像（影子图书馆），默认关闭；可用性与版权状态自行判断', enabled: false },

    // ── 图书 ──────────────────────────────────────────────────────────────
    { type: 'book', domain: 'openlibrary.org', label: 'Open Library', note: '书目与借阅入口', enabled: true },
    { type: 'book', domain: 'gutenberg.org', label: 'Project Gutenberg', note: '公版电子书全文', enabled: true },
    { type: 'book', domain: 'book.douban.com', label: '豆瓣读书', note: '中文书目、目录与书评', enabled: true },
    { type: 'book', domain: 'standardebooks.org', label: 'Standard Ebooks', note: '校订过的公版电子书', enabled: true },
    { type: 'book', domain: 'libgen.is', label: 'Library Genesis', note: '影子图书馆，默认关闭；镜像域名常变', enabled: false },
    { type: 'book', domain: 'z-lib.io', label: 'Z-Library', note: '影子图书馆，默认关闭；镜像域名常变', enabled: false },
    { type: 'book', domain: 'annas-archive.org', label: "Anna's Archive", note: '影子图书馆聚合，默认关闭', enabled: false },

    // ── 代码 ──────────────────────────────────────────────────────────────
    { type: 'code', domain: 'github.com', label: 'GitHub', note: '仓库、Issue 与讨论', enabled: true },
    { type: 'code', domain: 'stackoverflow.com', label: 'Stack Overflow', note: '问答，常有可复现的答案', enabled: true },
    { type: 'code', domain: 'gitee.com', label: 'Gitee', note: '国内仓库，直连更稳', enabled: true },
    { type: 'code', domain: 'developer.mozilla.org', label: 'MDN', note: 'Web 平台参考文档', enabled: true },
    { type: 'code', domain: 'pypi.org', label: 'PyPI', note: 'Python 包与文档入口', enabled: true },
    { type: 'code', domain: 'npmjs.com', label: 'npm', note: 'Node 包与文档入口', enabled: true },
]);

/** 单次调用最多限定几个站点（设置页与 schema 的区间共用）。 */
export const SITE_PRIORITY_LIMITS = Object.freeze({ maxPerCall: [1, 8], fallbackSites: 2 });

/** 一条站点条目的域名是否合法：只认 `a.b` 形状，不带协议、路径、端口与空格。 */
const DOMAIN_PATTERN = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/i;

/**
 * 把任意写法的域名收敛成比较用的形状：小写、去协议、去路径、去端口、去 `www.`。
 *
 * 收敛只此一处：`normalizeSiteEntry`（存进清单）与 `selectSiteEntries` 的点名
 * （按域名指定）必须走同一条规则 —— 否则 `https://www.arxiv.org/x` 这种写法在
 * 一边认得、另一边认不得（第三十四轮子代理审查抓到的真实不一致）。
 */
export function normalizeDomain(value) {
    const text = String(value ?? '').trim().toLowerCase();
    const noScheme = text.replace(/^[a-z][a-z0-9+.-]*:\/\//, '');
    const noPath = noScheme.split(/[/?#]/)[0];
    const noPort = noPath.replace(/:\d+$/, '');
    return noPort.replace(/^www\./, '');
}

/** 把任意输入收敛成一条合法条目；不合法返回 null（宁可少一条，也不写坏数据）。 */
export function normalizeSiteEntry(raw) {
    if (raw === null || typeof raw !== 'object') return null;
    const domain = normalizeDomain(raw.domain);
    if (!DOMAIN_PATTERN.test(domain)) return null;
    const type = SITE_TYPE_IDS.includes(String(raw.type ?? '').trim()) ? String(raw.type).trim() : 'custom';
    const label = String(raw.label ?? '').trim() || domain;
    return {
        type,
        domain,
        label: label.slice(0, 60),
        note: String(raw.note ?? '').trim().slice(0, 200),
        enabled: raw.enabled !== false,
    };
}

/** 内置目录的副本（深拷贝：调用方改了不会污染常量）。 */
export function siteCatalog() {
    return BUILTIN_SITES.map((item) => ({ ...item }));
}

/** 按类型分组（设置页渲染与反馈汇总都用它；顺序跟着 SITE_TYPES）。 */
export function groupSitesByType(entries) {
    const groups = [];
    for (const type of SITE_TYPES) {
        const items = (entries ?? []).filter((item) => item.type === type.id);
        if (items.length > 0) groups.push({ type, items });
    }
    return groups;
}

/**
 * 这次真正生效的清单。
 *
 * `entries === undefined`（没配过）→ 内置目录；`entries === []`（用户把行删光了）
 * → 真的没有站点，不再兜回内置目录，否则「删空」这个动作在界面上等于没生效。
 * 两条语义分开，是这一层唯一的坑。
 */
export function effectiveSiteEntries(settings) {
    // 畸形输入（裸字符串 / 数字）不当成「没配过」：那会把它悄悄变成一整份内置目录。
    // `undefined` / `null` 才是「没配过」。
    if (settings !== undefined && settings !== null && typeof settings !== 'object') return [];
    const raw = settings?.entries;
    if (raw === undefined || raw === null) return siteCatalog();
    if (!Array.isArray(raw)) return [];
    const cleaned = [];
    for (const item of raw) {
        const entry = normalizeSiteEntry(item);
        if (entry === null) continue;
        if (cleaned.some((existing) => existing.domain === entry.domain)) continue;
        cleaned.push(entry);
    }
    return cleaned;
}

/**
 * 挑出这次要限定的站点。
 *
 * `request.sites` 的四种形态：
 *   undefined / true  按清单里的启用项（清单顺序）
 *   'academic'        只要该类型（id 见 SITE_TYPES）
 *   ['arxiv.org', …]  按域名点名（顺序跟着传进来的顺序；认不出的域名进 skipped）
 *   false             本次不启用站点优先
 *
 * `max` 由调用方给（设置页的 maxPerCall，schema 已夹在 1..8）。
 */
export function selectSiteEntries(entries, request = {}) {
    const all = (entries ?? []).filter((item) => item.enabled !== false);
    const max = Number.isFinite(request.max) ? Math.max(1, Math.min(8, Math.trunc(request.max))) : all.length;
    const asked = request.sites;

    if (asked === false) return { picked: [], skipped: [], reason: '本次关掉了站点优先' };

    if (Array.isArray(asked)) {
        const picked = [];
        const skipped = [];
        for (const wanted of asked) {
            const domain = normalizeDomain(wanted);
            if (domain === '') continue;
            // 点名时不看 enabled：用户明确说了用哪几个站点，比清单上的开关更权威。
            // 收敛走 normalizeDomain（与写进清单时同一条规则），
            // 所以 `https://www.arxiv.org/x` 这种写法也认得。
            const hit = (entries ?? []).find((item) => item.domain === domain);
            if (hit === undefined) skipped.push(String(wanted ?? '').trim() || domain);
            else if (!picked.some((item) => item.domain === hit.domain)) picked.push(hit);
        }
        return { picked: picked.slice(0, max), skipped, reason: '按域名点名' };
    }

    if (typeof asked === 'string' && asked.trim() !== '') {
        const type = asked.trim();
        if (!SITE_TYPE_IDS.includes(type)) {
            return { picked: [], skipped: [], reason: `不认识的站点类型「${type}」` };
        }
        const picked = all.filter((item) => item.type === type);
        return { picked: picked.slice(0, max), skipped: [], reason: `按类型「${type}」` };
    }

    return { picked: all.slice(0, max), skipped: [], reason: '按清单里的启用项' };
}

/** 一条查询的站点限定形态：`site:arxiv.org <查询>`。 */
export function siteQueryFor(query, entry) {
    const text = String(query ?? '').trim();
    const domain = String(entry?.domain ?? '').trim();
    if (domain === '') return text;
    return text === '' ? `site:${domain}` : `site:${domain} ${text}`;
}

/**
 * 一个 URL 是否落在清单站点上（含子域）。
 *
 * `www.` 前缀在两边都无视：清单里写 `developer.mozilla.org`，
 * 结果里回来 `www.developer.mozilla.org` 也算命中。
 */
export function matchSiteEntry(url, entries) {
    let host;
    try {
        host = new URL(String(url ?? '')).hostname.toLowerCase().replace(/^www\./, '');
    } catch {
        return undefined;
    }
    if (host === '') return undefined;
    for (const entry of entries ?? []) {
        const domain = String(entry?.domain ?? '').toLowerCase().replace(/^www\./, '');
        if (domain === '') continue;
        if (host === domain || host.endsWith('.' + domain)) return entry;
    }
    return undefined;
}

/**
 * 按类型汇总命中数，给反馈用：`[{ type, name, total, labels: ['arXiv', 'CrossRef'] }]`。
 *
 * 只统计**命中过的**类型（没命中的类型在反馈里另有「空手」那句话说），
 * labels 按命中次数从多到少，同数保持清单顺序。
 */
export function summarizeSiteHits(hits) {
    const byType = new Map();
    for (const hit of hits ?? []) {
        const type = hit?.entry?.type ?? 'custom';
        if (!byType.has(type)) byType.set(type, new Map());
        const labels = byType.get(type);
        const label = String(hit?.entry?.label ?? '').trim() || hit?.entry?.domain || '未知站点';
        labels.set(label, (labels.get(label) ?? 0) + 1);
    }
    const groups = [];
    for (const type of SITE_TYPES) {
        const labels = byType.get(type.id);
        if (labels === undefined) continue;
        const ordered = [...labels.entries()].sort((a, b) => b[1] - a[1]);
        groups.push({
            type: type.id,
            name: type.name,
            total: ordered.reduce((sum, item) => sum + item[1], 0),
            labels: ordered.map((item) => `${item[0]} ${item[1]}`),
        });
    }
    return groups;
}
