/**
 * 站点清单的单元测试：目录数据质量、条目收敛、清单生效、选择、域名匹配与命中汇总。
 *
 * 这一套钉的都是**纯逻辑**，不联网、不起进程、不落盘：目录是数据，选择与匹配是
 * 纯函数。真正联网的那部分（限定查询打不打得到东西）不在这里测。
 *
 * 几条最容易被后人「顺手改坏」的语义，各占一条独立用例，不许合并：
 *   - `entries === []`（用户删空）是真的空，**不兜回内置目录**；
 *   - 点名（`sites: ['sci-hub.se']`）**不看条目自身的 enabled**，
 *     默认选择与按类型选择才只看 `enabled !== false`；
 *   - 影子图书馆（sci-hub.se / z-lib.io / annas-archive.org）默认关闭 ——
 *     这是用户明确点名的口径（「只放进目录、默认不开」），所以这里逐条点名断言，
 *     以后谁把 enabled 改成 true，先红的是这一条；
 *   - `max` 夹在 1..8（设置页 schema 的区间）；
 *   - `siteCatalog()` 给的是副本，改它不许污染 `BUILTIN_SITES`。
 *
 * 跑法：node test/site-catalog.mjs
 */
import assert from 'node:assert/strict';

import {
    BUILTIN_SITES,
    SITE_PRIORITY_LIMITS,
    SITE_TYPES,
    SITE_TYPE_IDS,
    effectiveSiteEntries,
    groupSitesByType,
    matchSiteEntry,
    normalizeSiteEntry,
    selectSiteEntries,
    siteCatalog,
    siteQueryFor,
    summarizeSiteHits,
} from '../src/site-catalog.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 内置目录里默认开启的域名，顺序与清单一致（默认选择的期望值）。 */
const ENABLED_DOMAINS = BUILTIN_SITES.filter((item) => item.enabled !== false).map((item) => item.domain);

/** 内置目录里默认关闭的域名（影子图书馆 + 付费墙）。 */
const DISABLED_DOMAINS = BUILTIN_SITES.filter((item) => item.enabled === false).map((item) => item.domain);

// ── 目录数据质量 ──────────────────────────────────────────────────────────

await check('内置目录：域名都合法（normalizeSiteEntry 认它）且全局唯一', () => {
    assert.ok(BUILTIN_SITES.length >= 8, `内置目录太短，只有 ${BUILTIN_SITES.length} 条`);
    const seen = new Set();
    for (const item of BUILTIN_SITES) {
        const normalized = normalizeSiteEntry(item);
        assert.ok(normalized !== null, `内置条目的域名不合法：${item.domain}`);
        assert.equal(normalized.domain, item.domain,
            `内置条目的域名被 normalizeSiteEntry 改动过（说明写法不规范）：${item.domain} → ${normalized.domain}`);
        assert.ok(!seen.has(item.domain), `域名重复：${item.domain}`);
        seen.add(item.domain);
    }
    assert.equal(seen.size, BUILTIN_SITES.length);
});

await check('内置目录：类型都是已知 id、label/note 非空，有数据的类型每类至少 1 条', () => {
    for (const item of BUILTIN_SITES) {
        assert.ok(SITE_TYPE_IDS.includes(item.type), `未知类型 id：${item.type}（${item.domain}）`);
        assert.equal(typeof item.label, 'string');
        assert.ok(item.label.trim() !== '', `label 不能为空：${item.domain}`);
        assert.equal(typeof item.note, 'string');
        assert.ok(item.note.trim() !== '', `note 不能为空：${item.domain}`);
        assert.equal(typeof item.enabled, 'boolean', `enabled 必须是布尔：${item.domain}`);
    }
    // 注意：内置目录里**没有** custom 条目 —— custom 是「自己加的站点」那个桶
    //（见 SITE_TYPES 的 note），空着才是对的。所以这里钉的是「除 custom 外每个
    // 类型都有内置条目」，另加一条「custom 仍是合法 id」。
    for (const type of SITE_TYPES) {
        if (type.id === 'custom') continue;
        const count = BUILTIN_SITES.filter((item) => item.type === type.id).length;
        assert.ok(count >= 1, `类型 ${type.id}（${type.name}）一条站点都没有`);
    }
    assert.ok(SITE_TYPE_IDS.includes('custom'), 'custom 必须是合法类型 id（用户自建条目的落点）');
});

await check('内置目录：影子图书馆四条默认关闭（含 libgen.is）', () => {
    // 需求描述里把 libgen.is 列进了「默认开启」那组，但模块自身的口径是
    // 「影子图书馆类只进目录、默认关闭」（见模块头注释与条目 note），
    // libgen.is 的 note 也写着「影子图书馆，默认关闭」。这里按模块口径把
    // **四条**都钉成 false；若产品口径确实是「libgen.is 默认开」，那要改的是 src，
    // 改完这条会先红。
    for (const domain of ['sci-hub.se', 'z-lib.io', 'annas-archive.org', 'libgen.is']) {
        const item = BUILTIN_SITES.find((entry) => entry.domain === domain);
        assert.ok(item !== undefined, `内置目录少了 ${domain}`);
        assert.equal(item.enabled, false, `${domain} 必须默认关闭（只进目录、默认不开）`);
    }
});

await check('内置目录：常用公开站点默认开启', () => {
    // libgen.is 不在这份名单里：它是影子图书馆，默认关闭（见上一条）。
    for (const domain of ['arxiv.org', 'scholar.google.com', 'cnki.net', 'crossref.org', 'github.com', 'stackoverflow.com']) {
        const item = BUILTIN_SITES.find((entry) => entry.domain === domain);
        assert.ok(item !== undefined, `内置目录少了 ${domain}`);
        assert.equal(item.enabled, true, `${domain} 必须默认开启`);
    }
    // 默认关闭的那几条不能被顺手当成开启项：默认选择的期望值就是这两份名单之差。
    assert.equal(ENABLED_DOMAINS.length, BUILTIN_SITES.length - DISABLED_DOMAINS.length);
    for (const domain of DISABLED_DOMAINS) {
        assert.ok(!ENABLED_DOMAINS.includes(domain), `${domain} 是默认关闭项，不该出现在默认选择里`);
    }
});

// ── normalizeSiteEntry ───────────────────────────────────────────────────

await check('normalizeSiteEntry：协议、路径与 www 前缀被收敛成裸域名', () => {
    assert.equal(normalizeSiteEntry({ domain: 'https://www.ArXiv.org/abs/123' }).domain, 'arxiv.org');
    assert.equal(normalizeSiteEntry({ domain: 'http://a.com/x/y' }).domain, 'a.com');
    assert.equal(normalizeSiteEntry({ domain: '  arxiv.org  ' }).domain, 'arxiv.org', '两头空白要 trim');
    assert.equal(normalizeSiteEntry({ domain: 'www.a.com' }).domain, 'a.com');
});

await check('normalizeSiteEntry：非法输入一律返回 null', () => {
    // 端口被剥掉（`a.com:8080` → `a.com`）而不是判非法：从浏览器地址栏复制过来的
    // 写法常常带端口，剥掉比让用户自己删更省事。这条口径在第三十四轮由主代理拍板
    // （原先被当成非法输入，子代理的初版断言按旧口径写），不是断言写错。
    assert.equal(normalizeSiteEntry({ domain: 'a.com:8080' }).domain, 'a.com');
    for (const bad of ['', 'a b.com', 'notadomain', 'http://', 'user@a.com', '.com', 'a..com']) {
        assert.equal(normalizeSiteEntry({ domain: bad }), null, `「${bad}」不该被当成合法域名`);
    }
    for (const bad of [123, 0, true, null, undefined, 'arxiv.org', ['arxiv.org'], () => {}]) {
        assert.equal(normalizeSiteEntry(bad), null, `非对象输入 ${JSON.stringify(bad) ?? String(bad)} 必须返回 null`);
    }
    assert.equal(normalizeSiteEntry({}), null, '没有 domain 字段 → null');
    assert.equal(normalizeSiteEntry({ domain: {} }), null);
});

await check('点名（sites: [域名]）：URL 写法一律收敛，认不出的进 skipped', () => {
    // 这一条钉的是子代理审查抓到的真实不一致：点名侧原先只剥协议与路径、
    // 不剥 `www.`，于是 `https://www.arxiv.org/x` 认不出来而 `www.arxiv.org` 认得出。
    // 现在两侧共用 normalizeDomain 的同一套收敛。
    const entries = siteCatalog();
    for (const written of ['arxiv.org', 'https://arxiv.org/x', 'https://www.arxiv.org', 'http://www.arXiv.org/abs/1']) {
        const selected = selectSiteEntries(entries, { sites: [written] });
        assert.deepEqual(selected.picked.map((item) => item.domain), ['arxiv.org'], `「${written}」该被认出来`);
        assert.deepEqual(selected.skipped, [], `「${written}」不该进 skipped`);
    }
    const mixed = selectSiteEntries(entries, { sites: ['cnki.net', '不存在.com'] });
    assert.deepEqual(mixed.picked.map((item) => item.domain), ['cnki.net']);
    assert.deepEqual(mixed.skipped, ['不存在.com']);
});

await check('normalizeSiteEntry：未知 type 落 custom，已知 type 保留', () => {
    assert.equal(normalizeSiteEntry({ domain: 'a.com', type: 'academic' }).type, 'academic');
    assert.equal(normalizeSiteEntry({ domain: 'a.com', type: '不存在的类型' }).type, 'custom');
    assert.equal(normalizeSiteEntry({ domain: 'a.com' }).type, 'custom', '缺 type → custom');
    assert.ok(SITE_TYPE_IDS.includes(normalizeSiteEntry({ domain: 'a.com' }).type));
});

await check('normalizeSiteEntry：label 缺省（或只有空白）时用域名', () => {
    assert.equal(normalizeSiteEntry({ domain: 'www.Example.com' }).label, 'example.com');
    assert.equal(normalizeSiteEntry({ domain: 'a.com', label: '   ' }).label, 'a.com');
    assert.equal(normalizeSiteEntry({ domain: 'a.com', label: ' 我的站点 ' }).label, '我的站点', 'label 要 trim');
});

await check('normalizeSiteEntry：label 与 note 超长被截断（label 60 / note 200）', () => {
    const entry = normalizeSiteEntry({ domain: 'a.com', label: 'x'.repeat(100), note: 'y'.repeat(300) });
    assert.equal(entry.label.length, 60, 'label 要截到 60');
    assert.equal(entry.note.length, 200, 'note 要截到 200');
    // 没超长时不许顺手截短。
    const short = normalizeSiteEntry({ domain: 'a.com', label: '短', note: '一句话' });
    assert.equal(short.label, '短');
    assert.equal(short.note, '一句话');
});

await check('normalizeSiteEntry：enabled 只有显式 false 才算关', () => {
    assert.equal(normalizeSiteEntry({ domain: 'a.com' }).enabled, true, '缺省 → 开');
    assert.equal(normalizeSiteEntry({ domain: 'a.com', enabled: true }).enabled, true);
    assert.equal(normalizeSiteEntry({ domain: 'a.com', enabled: false }).enabled, false);
    assert.equal(normalizeSiteEntry({ domain: 'a.com', enabled: 0 }).enabled, true, '只有 false 字面量才算关');
});

// ── siteCatalog ──────────────────────────────────────────────────────────

await check('siteCatalog 返回副本：改它不污染 BUILTIN_SITES', () => {
    const copy = siteCatalog();
    assert.ok(Object.isFrozen(BUILTIN_SITES), 'BUILTIN_SITES 必须是冻结的常量');
    assert.equal(copy.length, BUILTIN_SITES.length);
    assert.notEqual(copy, BUILTIN_SITES, '不能把常量本身交出去');
    assert.notEqual(copy[0], BUILTIN_SITES[0], '条目也要是副本（深拷贝一层）');

    copy[0].domain = 'mutated.example.com';
    copy[0].enabled = false;
    copy[0].label = '改过的';
    copy.push({ type: 'custom', domain: 'injected.example.com', label: '塞进来的', note: '', enabled: true });

    assert.equal(BUILTIN_SITES[0].domain, 'arxiv.org', '常量被调用方改坏了');
    assert.equal(BUILTIN_SITES[0].enabled, true, '常量被调用方改坏了');
    assert.equal(BUILTIN_SITES[0].label, 'arXiv', '常量被调用方改坏了');
    assert.equal(BUILTIN_SITES.length, copy.length - 1, '往副本里 push 不该影响常量');
    assert.equal(siteCatalog()[0].domain, 'arxiv.org', '再取一份仍要是原样');
});

// ── effectiveSiteEntries ─────────────────────────────────────────────────

await check('effectiveSiteEntries：undefined / null 兜回内置目录', () => {
    assert.equal(effectiveSiteEntries().length, BUILTIN_SITES.length, '没有 settings 参数 = 没配过');
    assert.equal(effectiveSiteEntries({}).length, BUILTIN_SITES.length, 'settings 里没有 entries = 没配过');
    assert.equal(effectiveSiteEntries({ entries: undefined }).length, BUILTIN_SITES.length);
    assert.equal(effectiveSiteEntries({ entries: null }).length, BUILTIN_SITES.length);
    // 是内置目录的**副本**，不是常量本身。
    const entries = effectiveSiteEntries({});
    entries[0].domain = 'mutated.example.com';
    assert.equal(BUILTIN_SITES[0].domain, 'arxiv.org');
});

await check('effectiveSiteEntries：[] 是真的空，不兜回内置目录（关键语义）', () => {
    const entries = effectiveSiteEntries({ entries: [] });
    assert.equal(entries.length, 0, '用户把行删光就是没有站点，删空这个动作必须生效');
    assert.notEqual(entries.length, BUILTIN_SITES.length, '不许兜回内置目录');
    assert.deepEqual(entries, []);
});

await check('effectiveSiteEntries：重复域名去重（先来的赢）', () => {
    const entries = effectiveSiteEntries({
        entries: [
            { domain: 'arxiv.org', label: '第一个' },
            { domain: 'https://www.arXiv.org/abs/1' },
            { domain: 'arxiv.org/' },
            { domain: 'cnki.net' },
        ],
    });
    assert.equal(entries.length, 2, `去重后应只剩 2 条，实得 ${entries.map((e) => e.domain).join(',')}`);
    assert.equal(entries[0].label, '第一个', '同名时保留先来的那条');
    assert.deepEqual(entries.map((entry) => entry.domain), ['arxiv.org', 'cnki.net']);
});

await check('effectiveSiteEntries：非法条目被丢弃，合法的照留', () => {
    const entries = effectiveSiteEntries({
        entries: [
            { domain: 'arxiv.org' },
            { domain: 'notadomain' },
            'cnki.net',
            null,
            42,
            { domain: '' },
            { domain: 'github.com', type: 'code' },
        ],
    });
    assert.deepEqual(entries.map((entry) => entry.domain), ['arxiv.org', 'github.com'], '只有合法的两条该留下');
    for (const entry of entries) {
        assert.ok(SITE_TYPE_IDS.includes(entry.type), '留下的条目要已经被收敛过');
    }
});

await check('effectiveSiteEntries：非数组（字符串 / 对象）→ []', () => {
    assert.deepEqual(effectiveSiteEntries({ entries: 'arxiv.org' }), []);
    assert.deepEqual(effectiveSiteEntries({ entries: { 0: { domain: 'arxiv.org' } } }), []);
    assert.deepEqual(effectiveSiteEntries({ entries: 42 }), []);
    assert.deepEqual(effectiveSiteEntries({ entries: true }), []);
});

// ── selectSiteEntries ────────────────────────────────────────────────────

await check('selectSiteEntries：默认（undefined / true）只取启用项，顺序跟清单一致', () => {
    for (const sites of [undefined, true]) {
        const out = selectSiteEntries(siteCatalog(), { sites });
        assert.deepEqual(out.picked.map((entry) => entry.domain), ENABLED_DOMAINS,
            `sites=${String(sites)} 时该按清单顺序取启用项`);
        assert.deepEqual(out.skipped, []);
        assert.ok(out.picked.every((entry) => entry.enabled !== false), '默认选择不许带出默认关闭的条目');
        for (const domain of DISABLED_DOMAINS) {
            assert.ok(!out.picked.some((entry) => entry.domain === domain), `默认选择不该含 ${domain}`);
        }
    }
    assert.equal(selectSiteEntries(siteCatalog(), {}).picked.length, ENABLED_DOMAINS.length, '不带 sites 等同默认');
    assert.deepEqual(selectSiteEntries(undefined, {}).picked, [], '条目为空时取不出东西，不该抛错');
});

await check("selectSiteEntries：'academic' 只取该类型的启用项，不含默认关闭的", () => {
    const out = selectSiteEntries(siteCatalog(), { sites: 'academic' });
    assert.ok(out.picked.length >= 1);
    assert.ok(out.picked.every((entry) => entry.type === 'academic'), '混进了别的类型');
    assert.ok(out.picked.every((entry) => entry.enabled !== false), '按类型选也要守 enabled 开关');
    for (const domain of ['sci-hub.se', 'sciencedirect.com', 'springer.com', 'ieee.org']) {
        assert.ok(!out.picked.some((entry) => entry.domain === domain), `按类型选不该带出默认关闭的 ${domain}`);
    }
    const book = selectSiteEntries(siteCatalog(), { sites: 'book' });
    assert.ok(book.picked.every((entry) => entry.type === 'book'));
    const custom = selectSiteEntries(siteCatalog(), { sites: 'custom' });
    assert.deepEqual(custom.picked, [], '内置目录里没有自定义条目');
});

await check('selectSiteEntries：不认识的类型 → picked 空，reason 里出现那个字符串', () => {
    const out = selectSiteEntries(siteCatalog(), { sites: '不存在的类型' });
    assert.deepEqual(out.picked, []);
    assert.deepEqual(out.skipped, []);
    assert.ok(out.reason.includes('不存在的类型'), `reason 要点名那个类型，实得：${out.reason}`);
});

await check('selectSiteEntries：按域名点名，认不出的进 skipped', () => {
    const out = selectSiteEntries(siteCatalog(), { sites: ['cnki.net', '不存在.com'] });
    assert.deepEqual(out.picked.map((entry) => entry.domain), ['cnki.net']);
    assert.deepEqual(out.skipped, ['不存在.com'], '认不出的域名要如实报出来，不能静默丢掉');
    assert.ok(out.reason.includes('点名'));
});

await check('selectSiteEntries：点名不看条目自身的 enabled（明确指定比清单开关更权威）', () => {
    const out = selectSiteEntries(siteCatalog(), { sites: ['sci-hub.se'] });
    assert.equal(out.picked.length, 1, '默认关闭的站点被点名时必须能选中');
    assert.equal(out.picked[0].domain, 'sci-hub.se');
    assert.equal(out.picked[0].enabled, false, '选中的这条本来就是 enabled=false —— 这一条钉的就是「点名不查开关」');
    assert.deepEqual(out.skipped, []);

    const multi = selectSiteEntries(siteCatalog(), { sites: ['arxiv.org', 'z-lib.io', 'annas-archive.org'] });
    assert.deepEqual(multi.picked.map((entry) => entry.domain), ['arxiv.org', 'z-lib.io', 'annas-archive.org'],
        '点名顺序跟着传进来的顺序，且三个默认关闭的都能选中');
    // 点同一个域名两次只算一条。
    const twice = selectSiteEntries(siteCatalog(), { sites: ['cnki.net', 'cnki.net'] });
    assert.equal(twice.picked.length, 1);
});

await check('selectSiteEntries：false → 本次不启用站点优先', () => {
    const out = selectSiteEntries(siteCatalog(), { sites: false });
    assert.deepEqual(out.picked, []);
    assert.deepEqual(out.skipped, []);
    assert.ok(out.reason.includes('关掉'), `reason 要说明是本次关掉了：${out.reason}`);
});

await check('selectSiteEntries：max 生效（传 2 只取 2 条）', () => {
    const out = selectSiteEntries(siteCatalog(), { max: 2 });
    assert.equal(out.picked.length, 2);
    assert.deepEqual(out.picked.map((entry) => entry.domain), ENABLED_DOMAINS.slice(0, 2), '取的是清单最前面两条');
    const academic = selectSiteEntries(siteCatalog(), { sites: 'academic', max: 3 });
    assert.equal(academic.picked.length, 3, 'max 对按类型选同样生效');
    const named = selectSiteEntries(siteCatalog(), { sites: ['arxiv.org', 'cnki.net', 'github.com'], max: 2 });
    assert.equal(named.picked.length, 2, 'max 对点名同样生效');
});

await check('selectSiteEntries：max 夹在 1..8（999 → ≤8，0 → ≥1）', () => {
    const big = selectSiteEntries(siteCatalog(), { max: 999 });
    assert.ok(big.picked.length <= SITE_PRIORITY_LIMITS.maxPerCall[1], `max=999 时取了 ${big.picked.length} 条`);
    assert.equal(big.picked.length, 8, '夹到上限 8');
    const zero = selectSiteEntries(siteCatalog(), { max: 0 });
    assert.ok(zero.picked.length >= SITE_PRIORITY_LIMITS.maxPerCall[0], 'max=0 也要至少取一条');
    assert.equal(zero.picked.length, 1, '夹到下限 1');
    const negative = selectSiteEntries(siteCatalog(), { max: -5 });
    assert.equal(negative.picked.length, 1);
    // 没给 max（或给了非数字）时不截断：有多少启用项取多少。
    assert.equal(selectSiteEntries(siteCatalog(), { max: Number.NaN }).picked.length, ENABLED_DOMAINS.length);
});

// ── siteQueryFor ─────────────────────────────────────────────────────────

await check('siteQueryFor：拼成 site:<域名> <查询>，空查询只剩限定', () => {
    assert.equal(siteQueryFor('某主题', { domain: 'arxiv.org' }), 'site:arxiv.org 某主题');
    assert.equal(siteQueryFor('', { domain: 'arxiv.org' }), 'site:arxiv.org');
    assert.equal(siteQueryFor('   ', { domain: 'arxiv.org' }), 'site:arxiv.org', '只有空白的查询等同空查询');
    assert.equal(siteQueryFor('  某主题  ', { domain: 'arxiv.org' }), 'site:arxiv.org 某主题', '查询两头空白要 trim');
});

await check('siteQueryFor：没有站点（entry 为 undefined / domain 为空）时原样返回查询', () => {
    assert.equal(siteQueryFor('某主题', undefined), '某主题');
    assert.equal(siteQueryFor('某主题', null), '某主题');
    assert.equal(siteQueryFor('某主题', {}), '某主题');
    assert.equal(siteQueryFor('某主题', { domain: '' }), '某主题');
    assert.equal(siteQueryFor('某主题', { domain: '   ' }), '某主题');
    assert.equal(siteQueryFor(undefined, undefined), '', '两个都空时返回空串，不抛错');
});

// ── matchSiteEntry ───────────────────────────────────────────────────────

await check('matchSiteEntry：命中（含 www 前缀、子域与大小写混合）', () => {
    const entries = siteCatalog();
    assert.equal(matchSiteEntry('https://www.arxiv.org/abs/1', entries).domain, 'arxiv.org', 'www 前缀两边都无视');
    assert.equal(matchSiteEntry('https://scholar.google.com/citations', entries).domain, 'scholar.google.com');
    assert.equal(matchSiteEntry('https://pubmed.ncbi.nlm.nih.gov/12345/', entries).domain, 'pubmed.ncbi.nlm.nih.gov');
    assert.equal(matchSiteEntry('HTTPS://WWW.ArXiv.ORG/abs/1', entries).domain, 'arxiv.org', '大小写混合也要命中');
    // 命中的是清单里的原条目（同一次调用里可直接拿 label 用）。
    const hit = matchSiteEntry('https://github.com/x/y', entries);
    assert.equal(hit.label, 'GitHub');
    // 匹配不看 enabled：默认关闭的站点，其结果照样认得出来（要不要用是选择那一层的事）。
    assert.equal(matchSiteEntry('https://sci-hub.se/paper/1', entries).domain, 'sci-hub.se');
});

await check('matchSiteEntry：不命中（后缀伪装与无关站点）', () => {
    const entries = siteCatalog();
    assert.equal(matchSiteEntry('https://arxiv.org.evil.com/x', entries), undefined, '后缀伪装不许命中');
    assert.equal(matchSiteEntry('https://example.com', entries), undefined, '清单外的站点不命中');
    assert.equal(matchSiteEntry('https://notarxiv.org/x', entries), undefined, '前缀不算命中');
    assert.equal(matchSiteEntry('https://arxiv.org/x', []), undefined, '空清单不命中');
    assert.equal(matchSiteEntry('https://arxiv.org/x', undefined), undefined);
});

await check('matchSiteEntry：非法 URL 返回 undefined（不抛错）', () => {
    const entries = siteCatalog();
    assert.equal(matchSiteEntry('not a url', entries), undefined);
    assert.equal(matchSiteEntry('', entries), undefined);
    assert.equal(matchSiteEntry(undefined, entries), undefined);
    assert.equal(matchSiteEntry(null, entries), undefined);
    assert.equal(matchSiteEntry('arxiv.org', entries), undefined, '裸域名不是合法 URL');
});

// ── summarizeSiteHits ────────────────────────────────────────────────────

await check('summarizeSiteHits：按类型聚合，labels 按命中次数降序，没命中的类型不出现', () => {
    const hits = [
        { entry: { type: 'academic', domain: 'arxiv.org', label: 'arXiv' } },
        { entry: { type: 'academic', domain: 'arxiv.org', label: 'arXiv' } },
        { entry: { type: 'academic', domain: 'crossref.org', label: 'CrossRef' } },
        { entry: { type: 'code', domain: 'github.com', label: 'GitHub' } },
    ];
    const groups = summarizeSiteHits(hits);
    assert.equal(groups.length, 2, '只有命中的两个类型出现（图书、自定义一条都没命中）');
    assert.deepEqual(groups.map((group) => group.type), ['academic', 'code'], '组的顺序跟 SITE_TYPES');
    assert.equal(groups[0].type, 'academic');
    assert.equal(groups[0].name, '学术', 'name 用的是中文类型名，不是 id');
    assert.equal(groups[0].total, 3, '总数是命中条数，不是站点数');
    assert.deepEqual(groups[0].labels, ['arXiv 2', 'CrossRef 1'], 'labels 按次数降序，并带上次数');
    assert.equal(groups[1].name, '代码');
    assert.equal(groups[1].total, 1);
    assert.deepEqual(groups[1].labels, ['GitHub 1']);
    assert.ok(!groups.some((group) => group.type === 'book' || group.type === 'custom'), '没命中的类型不许出现');
});

await check('summarizeSiteHits：空输入 / 缺字段时不抛错', () => {
    assert.deepEqual(summarizeSiteHits([]), []);
    assert.deepEqual(summarizeSiteHits(undefined), []);
    // 没有 entry 的命中落到「自定义 / 未知站点」，而不是把整条汇总搞崩。
    const groups = summarizeSiteHits([{}, { entry: { domain: 'x.com' } }]);
    assert.equal(groups.length, 1);
    assert.equal(groups[0].type, 'custom');
    assert.equal(groups[0].name, '自定义');
    assert.equal(groups[0].total, 2);
});

// ── groupSitesByType ─────────────────────────────────────────────────────

await check('groupSitesByType：分组顺序跟 SITE_TYPES、空组不出现、组内保持传入顺序', () => {
    const entries = [
        { type: 'code', domain: 'b.com' },
        { type: 'academic', domain: 'a.com' },
        { type: 'code', domain: 'c.com' },
    ];
    const groups = groupSitesByType(entries);
    assert.deepEqual(groups.map((group) => group.type.id), ['academic', 'code'], '组的顺序跟 SITE_TYPES，不是传入顺序');
    assert.equal(groups.length, 2, '图书与自定义是空组，不许出现');
    assert.deepEqual(groups[0].items.map((item) => item.domain), ['a.com']);
    assert.deepEqual(groups[1].items.map((item) => item.domain), ['b.com', 'c.com'], '组内保持传入顺序');
    assert.equal(groups[1].type.name, '代码');

    assert.deepEqual(groupSitesByType([]), []);
    assert.deepEqual(groupSitesByType(undefined), []);
    assert.deepEqual(groupSitesByType(null), []);
    // 内置目录按类型分组：组的顺序仍是 SITE_TYPES 的顺序，且一条都不许漏
    //（内置目录里没有 custom 条目，所以空组不出现这条同时被验到）。
    const full = groupSitesByType(siteCatalog());
    assert.deepEqual(full.map((group) => group.type.id), ['academic', 'book', 'code']);
    assert.equal(full.reduce((sum, group) => sum + group.items.length, 0), BUILTIN_SITES.length, '一条都不许漏');
});

// ── LIMITS ───────────────────────────────────────────────────────────────

await check('SITE_PRIORITY_LIMITS：maxPerCall 是 [1, 8]，fallbackSites 是正整数', () => {
    assert.deepEqual([...SITE_PRIORITY_LIMITS.maxPerCall], [1, 8]);
    assert.ok(Number.isInteger(SITE_PRIORITY_LIMITS.fallbackSites) && SITE_PRIORITY_LIMITS.fallbackSites > 0,
        `fallbackSites 要是正整数，实得 ${SITE_PRIORITY_LIMITS.fallbackSites}`);
    assert.ok(SITE_PRIORITY_LIMITS.fallbackSites <= SITE_PRIORITY_LIMITS.maxPerCall[1], '兜底站点数不该超过单次上限');
    assert.ok(Object.isFrozen(SITE_PRIORITY_LIMITS), 'LIMITS 是常量');
    // 上限真的和选择逻辑对得上：内置目录的启用项远多于 8，max=8 仍只取 8 条。
    assert.ok(ENABLED_DOMAINS.length > SITE_PRIORITY_LIMITS.maxPerCall[1], '内置启用项要足够多，否则上面那条夹取断言会失去意义');
    assert.equal(selectSiteEntries(siteCatalog(), { max: SITE_PRIORITY_LIMITS.maxPerCall[1] }).picked.length, 8);
});

// ── 收尾 ──────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`site-catalog: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
