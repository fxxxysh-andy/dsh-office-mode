/**
 * 内置联网通道的单元测试（全部离线）。
 *
 * 这些用例钉的是「工具面里缺这几个联网工具的精简部署（第二十九轮之前的办公 preset 就是
 * 这样），检索还能不能用」这条需求的可判定部分：
 * 通道选择（接缝优先、失败回落）、HTML 收文本、地址分类（SSRF 防线）、
 * 响应映射与摘录挑选。真正的联网检索单独用 `node test/web-live.mjs` 跑，
 * 不进默认测试套件（默认套件不许依赖网络）。
 *
 * 跑法：node test/web.mjs
 */
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

import {
    charsetOf,
    classifyContentType,
    createWebAccess,
    decodeEntities,
    describeWebFailure,
    htmlTitle,
    htmlToText,
    inspectFetchedPage,
    isPublicAddress,
    looksLikePdf,
    mapSearchResponse,
    mergeSources,
    resolveBuiltinOptions,
    searchProvider,
    summarizeWebFailures,
    validateUrl,
    WebAccessError,
    WEB_ENGINE_BUILTIN,
    WEB_ENGINE_SEAM,
    WEB_FAILURE_CODES,
    WEB_FAILURE_KINDS,
    webFailureKind,
    webFailureLabel,
    webSeamOf,
} from '../src/web.js';
import { excerptOf } from '../src/search.js';
import { PROVIDER_IDS, PROVIDERS, providerDefaults } from '../src/web-providers.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 与真实 cordis 一致的最小替身：服务只能经 ctx.get 取。 */
function cordisLikeCtx(services = {}) {
    return new Proxy({ inject: ['tools'] }, {
        get(target, prop) {
            if (prop === 'get') return (name) => services[name];
            if (prop in target) return target[prop];
            throw new Error(`cannot get property "${String(prop)}" without inject`);
        },
    });
}

/** 一个可用的假接缝：检索返回两条来源，取正文返回一段 HTML。 */
function fakeSeam() {
    const calls = { search: [], fetch: [] };
    return {
        calls,
        search: async (request) => {
            calls.search.push(request);
            return {
                content: '模型写的概览',
                sources: [
                    { url: 'https://news-a.example.com/1', title: '来源甲', snippet: '引用片段甲' },
                    { url: 'https://news-b.example.com/2', title: '来源乙' },
                ],
                truncated: false,
            };
        },
        fetch: async (request) => {
            calls.fetch.push(request);
            return {
                url: request.url,
                statusCode: 200,
                body: { kind: 'html', content: '<html><head><title>页面乙</title></head><body><p>这是页面乙的正文段落，长度足够当摘录用的一行。</p></body></html>' },
                truncated: false,
            };
        },
    };
}

// ── 接缝识别 ──────────────────────────────────────────────────────────────

await check('webSeamOf 只认经 ctx.get 拿到、且 search/fetch 齐备的服务', () => {
    const seam = fakeSeam();
    assert.equal(webSeamOf(cordisLikeCtx({ web: seam })), seam);
    assert.equal(webSeamOf(cordisLikeCtx({})), undefined);
    assert.equal(webSeamOf(undefined), undefined);
    assert.equal(webSeamOf({ web: seam }), undefined, '没有 ctx.get 的替身不算有接缝');
    assert.equal(webSeamOf(cordisLikeCtx({ web: { search: () => {} } })), undefined, '缺 fetch 不算接缝');
    assert.throws(() => cordisLikeCtx({ web: seam }).web, /without inject/, '替身要如实模拟宿主：未 inject 的属性读取必须抛错');
});

// ── 通道选择 ──────────────────────────────────────────────────────────────

await check('有接缝时检索走接缝，并把 maxResults 传下去', async () => {
    const seam = fakeSeam();
    const access = createWebAccess(cordisLikeCtx({ web: seam }), {});
    const result = await access.search('某事件', { maxResults: 3 });
    assert.equal(result.engine, WEB_ENGINE_SEAM);
    assert.equal(result.sources.length, 2);
    assert.deepEqual(seam.calls.search, [{ query: '某事件', maxResults: 3 }]);
});

await check('接缝报错时自动回落到自带实现（并把接缝的失败原因记下来）', async () => {
    let fetched = 0;
    const seam = {
        search: async () => { throw new Error('WEB_PROVIDER_UNAVAILABLE: 没有可用的 provider'); },
        fetch: async () => { throw new Error('不可用'); },
    };
    const access = createWebAccess(cordisLikeCtx({ web: seam }), { apiKey: 'test-key' }, {
        fetch: async (url, init) => {
            fetched += 1;
            assert.ok(String(url).endsWith('/messages'), '自带检索要打 /messages 端点');
            assert.equal(init.headers['x-api-key'], 'test-key');
            return {
                ok: true,
                status: 200,
                async json() {
                    return {
                        content: [
                            { type: 'web_search_tool_result', content: [{ type: 'web_search_result', url: 'https://c.example.com/3', title: '来源丙' }] },
                            { type: 'text', text: '概览', citations: [{ url: 'https://c.example.com/3', cited_text: '被引用的那句话' }] },
                        ],
                    };
                },
            };
        },
    });
    const result = await access.search('某事件', { maxResults: 5 });
    assert.equal(result.engine, WEB_ENGINE_BUILTIN);
    assert.equal(fetched, 1);
    assert.equal(result.sources[0].snippet, '被引用的那句话');

    // 接缝失败过一次就不再重试：第二次检索直接走自带实现。
    await access.search('再查一次', {});
    assert.equal(fetched, 2);
});

await check('两边都不可用时，报错里同时点明接缝与内置的原因', async () => {
    const seam = { search: async () => { throw new Error('没有可用 provider'); }, fetch: async () => { throw new Error('x'); } };
    const access = createWebAccess(cordisLikeCtx({ web: seam }), { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY' }, { fetch: async () => { throw new Error('不该走到这里'); } });
    await assert.rejects(() => access.search('某事件'), (error) => {
        assert.ok(error instanceof WebAccessError);
        assert.match(error.message, /宿主 web 服务/);
        assert.match(error.message, /OFFICE_TEST_MISSING_KEY/);
        return true;
    });
    const status = await access.probe();
    assert.equal(status.ok, false);
    assert.match(status.reason, /宿主 web 服务/);
});

await check('没有接缝、但有 Key 时，probe 报自带检索可用', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' });
    const status = await access.probe();
    assert.equal(status.ok, true);
    assert.equal(status.engine, WEB_ENGINE_BUILTIN);
});

// ── 响应映射 ──────────────────────────────────────────────────────────────

await check('响应映射：按 URL 去重、引用片段贴在对应来源上、概览单独给', () => {
    const mapped = mapSearchResponse({
        content: [
            { type: 'web_search_tool_result', content: [
                { type: 'web_search_result', url: 'https://a.example.com/1', title: '甲', page_age: '2026-01-01' },
                { type: 'web_search_result', url: 'https://a.example.com/1', title: '甲的重复' },
                { type: 'web_search_result', url: 'https://b.example.com/2', title: '乙' },
            ] },
            { type: 'text', text: '一句话概览', citations: [{ url: 'https://b.example.com/2', cited_text: '乙的引用' }] },
        ],
    });
    assert.equal(mapped.sources.length, 2, '重复 URL 只留一条');
    assert.equal(mapped.sources[0].publishedAt, '2026-01-01');
    assert.equal(mapped.sources[1].snippet, '乙的引用');
    assert.equal(mapped.content, '一句话概览');
});

await check('没有 web_search_tool_result 块时明确报错，而不是拿正文凑数', () => {
    assert.throws(() => mapSearchResponse({ content: [{ type: 'text', text: '我没搜' }] }), /没有触发原生检索/);
});

await check('多查询结果合并去重后按上限截断', () => {
    const merged = mergeSources([
        { sources: [{ url: 'https://a/1' }, { url: 'https://a/2' }] },
        { sources: [{ url: 'https://a/2' }, { url: 'https://a/3' }] },
    ], 2);
    assert.equal(merged.sources.length, 2);
    assert.equal(merged.truncated, true);
});

// ── 取正文的策略 ──────────────────────────────────────────────────────────

await check('URL 校验：只许 http(s)、不许内嵌账号、长度受限', () => {
    assert.equal(validateUrl(' https://example.com/x ').hostname, 'example.com');
    assert.throws(() => validateUrl('file:///etc/passwd'), /只支持 http/);
    assert.throws(() => validateUrl('https://user:pw@example.com/'), /不允许带账号密码/);
    assert.throws(() => validateUrl('https://example.com/' + 'a'.repeat(2100)), /超过/);
});

await check('地址分类：环回、私网、链路本地、CGNAT、组播与 IPv6 内网都判为非公网', () => {
    for (const ip of ['127.0.0.1', '10.1.2.3', '172.16.0.1', '192.168.1.1', '169.254.1.1', '100.64.0.1', '0.0.0.0', '224.0.0.1', '255.255.255.255', '::1', '::', 'fc00::1', 'fe80::1', 'ff02::1', '64:ff9b::7f00:1', '::ffff:127.0.0.1']) {
        assert.equal(isPublicAddress(ip), false, `${ip} 不该被判成公网`);
    }
    for (const ip of ['8.8.8.8', '1.1.1.1', '2606:4700:4700::1111', '::ffff:8.8.8.8']) {
        assert.equal(isPublicAddress(ip), true, `${ip} 应该判成公网`);
    }
});

await check('取正文只跟随同源跳转，跨站跳转直接拒绝', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k', maxRedirects: 3 }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async (url) => ({
            status: 302,
            headers: { get: (name) => (name === 'location' ? 'https://other.example.com/x' : null) },
            body: { cancel: async () => {} },
        }),
    });
    await assert.rejects(() => access.fetch('https://example.com/a'), /跨站跳转|宿主 web 服务/);
});

await check('非公网主机拒绝连接（SSRF 防线）', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '10.0.0.5', family: 4 }],
        fetch: async () => { throw new Error('不该发起请求'); },
    });
    await assert.rejects(() => access.fetch('https://internal.example.com/'), /非公网地址/);
});

await check('取正文：HTML 收成文本、按 charset 解码、超长截断', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k', maxChars: 1_000 }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => ({
            status: 200,
            headers: { get: (name) => (name === 'content-type' ? 'text/html; charset=utf-8' : null) },
            body: {
                getReader() {
                    const chunks = [new TextEncoder().encode('<h1>标题</h1><p>' + '正'.repeat(3_000) + '</p>')];
                    let index = 0;
                    return {
                        async read() {
                            if (index >= chunks.length) return { done: true, value: undefined };
                            return { done: false, value: chunks[index++] };
                        },
                        async cancel() {},
                    };
                },
            },
        }),
    });
    const page = await access.fetch('https://example.com/');
    assert.equal(page.kind, 'html');
    assert.ok(page.content.startsWith('# 标题'), `HTML 标题要收成 Markdown 标题：${page.content.slice(0, 20)}`);
    assert.ok(page.content.length <= 1_000, `要按 maxChars 截断，实际 ${page.content.length}`);
});

await check('内容类型分类与 charset 解析', () => {
    assert.equal(classifyContentType('text/html; charset=GBK'), 'html');
    assert.equal(classifyContentType('application/json'), 'text');
    assert.equal(classifyContentType('image/png'), undefined);
    assert.equal(charsetOf('text/html; charset="GBK"'), 'gbk');
    assert.equal(charsetOf('text/html'), undefined);
});

// ── HTML 收文本 ───────────────────────────────────────────────────────────

await check('htmlToText 丢脚本样式、留标题列表与链接、解开实体', () => {
    const html = '<html><head><title>页面题</title><style>p{color:red}</style></head>'
        + '<body><script>alert(1)</script><h2>小标题</h2><p>正文&nbsp;一行 &amp; 两个</p>'
        + '<ul><li>条目一</li><li>条目二</li></ul><a href="https://x.example.com/y">链接文字</a></body></html>';
    const text = htmlToText(html);
    assert.ok(!text.includes('alert'), '脚本要整段丢掉');
    assert.ok(!text.includes('color:red'), '样式要整段丢掉');
    assert.ok(text.includes('## 小标题'), '标题要有层级标记');
    assert.ok(text.includes('- 条目一'), '列表项要保留');
    assert.ok(text.includes('[链接文字](https://x.example.com/y)'), '链接要留下地址');
    assert.ok(text.includes('正文 一行 & 两个'), '实体要解开');
    assert.equal(htmlTitle(html), '页面题');
});

await check('decodeEntities 处理命名、十进制与十六进制实体', () => {
    assert.equal(decodeEntities('&lt;a&gt; &#65; &#x42; &nbsp;'), '<a> A B  ');
});

// ── 摘录挑选 ──────────────────────────────────────────────────────────────

await check('摘录跳过导航式的链接行，优先取正文句', () => {
    const text = [
        '[首页](https://a/) [新闻](https://a/news) [体育](https://a/sports)',
        '注册 登录 客户端下载',
        '国务院常务会议部署地方政府债务风险化解工作，要求稳妥有序推进',
        '会议还研究了其他事项。',
    ].join('\n');
    const excerpt = excerptOf(text, 40);
    assert.ok(excerpt.startsWith('国务院常务会议'), `不该拿导航当摘录：${excerpt}`);
    assert.ok(excerpt.length <= 41);
});

// ── 内容级成败哨兵（P0-8）与企业级失败分类（P0-9）────────────────────────

await check('哨兵：状态码一类（403/404/429/451 判“目标站拒绝”，5xx 判“网络”）', () => {
    for (const statusCode of [401, 403, 404, 410, 429, 451]) {
        const verdict = inspectFetchedPage({ statusCode, url: 'https://a.example.com/x', content: '正常长度的正文'.repeat(20) });
        assert.equal(verdict.ok, false, `HTTP ${statusCode} 必须判失败`);
        assert.equal(verdict.code, 'OFFICE_WEB_BLOCKED', `HTTP ${statusCode} 该归到目标站拒绝`);
        assert.equal(webFailureKind(verdict), 'blocked');
    }
    const server = inspectFetchedPage({ statusCode: 503, url: 'https://a.example.com/x', content: '正文'.repeat(50) });
    assert.equal(server.code, 'OFFICE_WEB_NETWORK', '5xx 是服务端错误，归到网络出口不可达');
    assert.equal(webFailureKind(server), 'network');
});

await check('哨兵：HTTP 200 也可能是失败（区域封锁页 / 拦截页 / 登录墙 / 验证码）', () => {
    const long = (text) => text + '正文'.repeat(60);
    const cases = [
        [{ statusCode: 200, url: 'https://claude.com/app-unavailable-in-region', content: long('App unavailable in region') }, '区域封锁页（看 URL）'],
        [{ statusCode: 200, url: 'https://a.example.com/x', content: long('Sorry, you have been blocked') }, 'Cloudflare 拦截页'],
        [{ statusCode: 200, url: 'https://a.example.com/x', content: long('Just a moment...') }, 'Cloudflare 人机校验页'],
        [{ statusCode: 200, url: 'https://a.example.com/x', content: long('Enable JavaScript and cookies to continue') }, 'Cloudflare 人机校验页'],
        [{ statusCode: 200, url: 'https://a.example.com/x', content: long('This content is not available in your country') }, '区域封锁页'],
        [{ statusCode: 200, url: 'https://a.example.com/x', content: long('Sign in to continue') }, '登录墙'],
        [{ statusCode: 200, url: 'https://a.example.com/x', content: long('请完成验证 g-recaptcha') }, '验证码页'],
    ];
    for (const [input, what] of cases) {
        const verdict = inspectFetchedPage(input);
        assert.equal(verdict.ok, false, `${what} 必须判失败（这正是只看状态码会漏掉的那一类）`);
        assert.equal(verdict.code, 'OFFICE_WEB_BLOCKED', `${what} 该归到目标站拒绝`);
    }
});

await check('哨兵：正文几乎为空判失败，偏短只提醒，不误伤长正文', () => {
    const empty = inspectFetchedPage({ statusCode: 200, url: 'https://a.example.com/x', content: '  短  ' });
    assert.equal(empty.ok, false, '清洗后没几个字必须判失败');
    assert.equal(empty.code, 'OFFICE_WEB_EMPTY');
    assert.equal(webFailureKind(empty), 'empty');

    const short = inspectFetchedPage({ statusCode: 200, url: 'https://a.example.com/x', content: '这是一段只有几十个字的简短正文，够用但偏短。' });
    assert.equal(short.ok, true, '偏短不该判失败（误判的代价是丢掉真实来源）');
    assert.ok(typeof short.notice === 'string' && short.notice.includes('偏短'), '要给一句提醒');

    // 长正文里出现拦截页的关键词是巧合，不是拦截页 —— 超过扫描上限就不看关键词。
    const big = inspectFetchedPage({ statusCode: 200, url: 'https://a.example.com/x', content: '这是一篇讲 CDN 的长文。'.repeat(400) + 'Sorry, you have been blocked' });
    assert.equal(big.ok, true, '长正文里的巧合短语不该把真实来源判掉');
});

await check('哨兵：未知状态码（接缝回 0）不误判，交给地址与正文判定', () => {
    const unknown = inspectFetchedPage({ statusCode: 0, url: 'https://a.example.com/x', content: '一段足够长的正常正文'.repeat(10) });
    assert.equal(unknown.ok, true, '状态码未知时不能凭空判失败');
});

await check('自带取正文：403 与区域封锁页都当失败抛出，且带分类', async () => {
    const stub = (status, contentType) => ({
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => ({
            status,
            headers: { get: (name) => (name === 'content-type' ? contentType : null) },
            body: { cancel: async () => {}, getReader() { return { async read() { return { done: true, value: undefined }; }, async cancel() {} }; } },
        }),
    });
    const forbidden = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, stub(403, 'text/html'));
    await assert.rejects(() => forbidden.fetch('https://a.example.com/x'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_BLOCKED');
        assert.match(error.message, /目标站拒绝|HTTP 403/);
        return true;
    });

    const regionBlocked = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => ({
            status: 200,
            headers: { get: (name) => (name === 'content-type' ? 'text/html' : null) },
            body: {
                getReader() {
                    const chunks = [new TextEncoder().encode('<html><body><h1>App unavailable in region</h1></body></html>')];
                    let index = 0;
                    return { async read() { return index >= chunks.length ? { done: true } : { done: false, value: chunks[index++] }; }, async cancel() {} };
                },
            },
        }),
    });
    await assert.rejects(() => regionBlocked.fetch('https://claude.com/app-unavailable-in-region'), /区域封锁页|不是原文/);
});

// ── 多路通道（第二十轮）───────────────────────────────────────────────────
//
// 全部离线：每个适配器都拿注入的 fetch 跑，只断言「打到哪个端点、带什么鉴权、
// 怎么解析回来」。真正的联网检索由 test/web-live.mjs 单独跑。

/** 造一个记录调用的假 fetch：返回给定的 JSON 或 HTML。 */
function fakeHttp(reply) {
    const calls = [];
    return {
        calls,
        fetch: async (url, init) => {
            calls.push({ url: String(url), init: init ?? {} });
            return typeof reply === 'function' ? reply(url, init) : reply;
        },
    };
}

/** 造一个 JSON 响应。 */
function jsonReply(body, status = 200) {
    return {
        ok: status < 400,
        status,
        async json() { return body; },
        async text() { return JSON.stringify(body); },
    };
}

/** 造一个 HTML 响应。 */
function htmlReply(html, status = 200) {
    return {
        ok: status < 400,
        status,
        async json() { throw new Error('不是 JSON'); },
        async text() { return html; },
    };
}

/** 一条通道跑一次检索：返回 {result, calls}。 */
async function runChannel(id, options, reply, request = {}) {
    const http = fakeHttp(reply);
    const access = createWebAccess(cordisLikeCtx({}), {
        provider: id,
        providers: { [id]: { ...options, baseURL: options.baseURL ?? PROVIDERS[id].endpoint } },
    }, { fetch: http.fetch });
    const result = await access.search('某事件', request);
    return { result, calls: http.calls };
}

await check('通道清单：每条通道都有展示名、要不要 Key、默认端点', () => {
    assert.ok(PROVIDER_IDS.length >= 8, `通道太少（${PROVIDER_IDS.length} 条）`);
    for (const id of PROVIDER_IDS) {
        const meta = PROVIDERS[id];
        assert.ok(meta, `${id} 没有注册`);
        assert.equal(meta.id, id);
        assert.ok(typeof meta.label === 'string' && meta.label !== '', `${id} 缺展示名`);
        assert.equal(typeof meta.needsKey, 'boolean', `${id} 要说清要不要 Key`);
        if (meta.needsKey) assert.ok(typeof meta.keyEnv === 'string' && meta.keyEnv !== '', `${id} 缺 Key 环境变量名`);
    }
    // 零配置的兜底通道必须存在：没有它，一台没配 Key 的机器就查不了资料。
    const zero = PROVIDER_IDS.filter((id) => PROVIDERS[id].needsKey === false);
    assert.ok(zero.includes('duckduckgo'), '免 Key 兜底通道不能丢');
});

await check('点名通道时不会先去试接缝（用户选了哪条就走哪条）', async () => {
    const seam = fakeSeam();
    const http = fakeHttp(jsonReply({ results: [{ title: '甲', url: 'https://a.example.com/1', content: '摘要甲' }] }));
    const access = createWebAccess(cordisLikeCtx({ web: seam }), {
        provider: 'tavily',
        providers: { tavily: { apiKey: 'tv-key' } },
    }, { fetch: http.fetch });
    const result = await access.search('某事件', {});
    assert.equal(result.provider, 'tavily');
    assert.equal(result.engine, WEB_ENGINE_BUILTIN);
    assert.deepEqual(seam.calls.search, [], '点名 tavily 就不该再问接缝');
    assert.equal(result.sources[0].url, 'https://a.example.com/1');
});

await check('Tavily：POST /search、api_key 在体里、解析 results[].content 当摘录', async () => {
    const { result, calls } = await runChannel('tavily', { apiKey: 'tv-key' }, jsonReply({
        answer: '',
        results: [{ title: '甲', url: 'https://a.example.com/1', content: '摘要甲' }],
    }));
    assert.equal(calls[0].url, 'https://api.tavily.com/search');
    assert.equal(calls[0].init.method, 'POST');
    assert.equal(JSON.parse(calls[0].init.body).api_key, 'tv-key');
    assert.equal(result.sources.length, 1);
    assert.equal(result.sources[0].snippet, '摘要甲');
});

await check('Brave：GET /web/search、令牌走 X-Subscription-Token、解析 web.results', async () => {
    const { result, calls } = await runChannel('brave', { apiKey: 'br-key' }, jsonReply({
        web: { results: [{ title: '乙', url: 'https://b.example.com/2', description: '摘要乙', age: '2 天前' }] },
    }));
    assert.match(calls[0].url, /^https:\/\/api\.search\.brave\.com\/res\/v1\/web\/search\?q=/);
    assert.equal(calls[0].init.headers['x-subscription-token'], 'br-key');
    assert.equal(result.sources[0].snippet, '摘要乙');
    assert.equal(result.sources[0].publishedAt, '2 天前');
});

await check('博查：POST /web-search、Bearer 鉴权、解析 data.webPages.value', async () => {
    const { result, calls } = await runChannel('bocha', { apiKey: 'bc-key' }, jsonReply({
        data: { webPages: { value: [{ name: '丙', url: 'https://c.example.com/3', snippet: '摘要丙' }] } },
    }));
    assert.equal(calls[0].url, 'https://api.bochaai.com/v1/web-search');
    assert.equal(calls[0].init.headers.authorization, 'Bearer bc-key');
    assert.equal(result.sources[0].title, '丙');
    assert.equal(result.sources[0].url, 'https://c.example.com/3');
});

await check('Serper：POST /search、X-API-KEY 鉴权、解析 organic[].link', async () => {
    const { result, calls } = await runChannel('serper', { apiKey: 'sp-key' }, jsonReply({
        organic: [{ title: '丁', link: 'https://d.example.com/4', snippet: '摘要丁' }],
    }));
    assert.equal(calls[0].url, 'https://google.serper.dev/search');
    assert.equal(calls[0].init.headers['x-api-key'], 'sp-key');
    assert.equal(result.sources[0].url, 'https://d.example.com/4');
});

await check('Exa：POST /search、x-api-key 鉴权、解析 results[].text 当摘录', async () => {
    const { result, calls } = await runChannel('exa', { apiKey: 'ex-key' }, jsonReply({
        results: [{ title: '戊', url: 'https://e.example.com/5', text: '正文片段戊' }],
    }));
    assert.equal(calls[0].url, 'https://api.exa.ai/search');
    assert.equal(calls[0].init.headers['x-api-key'], 'ex-key');
    assert.equal(result.sources[0].snippet, '正文片段戊');
});

await check('OpenAI 兼容：POST /chat/completions、带 web_search_options、解析 url_citation', async () => {
    const { result, calls } = await runChannel('openai', { apiKey: 'oa-key', baseURL: 'https://compat.example.com/v1' }, jsonReply({
        choices: [{
            message: {
                content: '这是模型的概览。',
                annotations: [{ type: 'url_citation', url_citation: { url: 'https://f.example.com/6', title: '己' } }],
            },
        }],
    }));
    assert.equal(calls[0].url, 'https://compat.example.com/v1/chat/completions');
    const body = JSON.parse(calls[0].init.body);
    assert.deepEqual(body.web_search_options, {}, '不带 web_search_options 就不会触发检索');
    assert.equal(result.sources[0].url, 'https://f.example.com/6');
    assert.equal(result.content, '这是模型的概览。');
});

await check('SearXNG：GET /search?format=json；缺实例地址时报「配置缺失」而不是网络错', async () => {
    const { result, calls } = await runChannel('searxng', { baseURL: 'https://sx.example.com' }, jsonReply({
        results: [{ title: '庚', url: 'https://g.example.com/7', content: '摘要庚' }],
    }));
    assert.equal(calls[0].url, 'https://sx.example.com/search?q=%E6%9F%90%E4%BA%8B%E4%BB%B6&format=json');
    assert.equal(result.sources[0].snippet, '摘要庚');

    // 缺实例地址由适配器自己守住（配置层会把空端点填回默认值，所以这条防线在适配器上）。
    await assert.rejects(
        () => searchProvider('searxng', { query: '某事件', options: { ...providerDefaults('searxng'), baseURL: '' }, fetchImpl: fakeHttp(jsonReply({})).fetch }),
        (error) => {
            assert.equal(error.code, 'OFFICE_WEB_NO_BASE');
            assert.equal(webFailureKind(error), 'config', '缺实例地址是配置问题，不是网络问题');
            return true;
        },
    );
});

await check('DuckDuckGo：抓 HTML 结果页、解开 uddg 跳转、解析不出时明确报错', async () => {
    const html = '<html><body><div class="result"><a class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fh.example.com%2F8">辛</a>'
        + '<a class="result__snippet">摘要辛</a></div></body></html>';
    const { result, calls } = await runChannel('duckduckgo', {}, htmlReply(html));
    assert.match(calls[0].url, /^https:\/\/html\.duckduckgo\.com\/html\/\?q=/);
    assert.equal(result.sources[0].url, 'https://h.example.com/8');
    assert.equal(result.sources[0].snippet, '摘要辛');

    const captcha = createWebAccess(cordisLikeCtx({}), { provider: 'duckduckgo' }, { fetch: fakeHttp(htmlReply('<html><body>verify you are human</body></html>')).fetch });
    await assert.rejects(() => captcha.search('某事件'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_NO_RESULTS');
        assert.match(error.message, /人机校验|换一条通道/);
        return true;
    });
});

await check('auto：按顺序试通道，第一条成功的就用它；全失败时逐条给出原因', async () => {
    // 只有 tavily 配了 Key：anthropic 会当场失败（不联网），tavily 命中。
    const http = fakeHttp(jsonReply({ results: [{ title: '甲', url: 'https://a.example.com/1', content: '摘要' }] }));
    const access = createWebAccess(cordisLikeCtx({}), {
        provider: 'auto',
        providerOrder: ['anthropic', 'tavily'],
        providers: { anthropic: { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY' }, tavily: { apiKey: 'tv-key' } },
    }, { fetch: http.fetch });
    const result = await access.search('某事件');
    assert.equal(result.provider, 'tavily');
    assert.equal(http.calls.length, 1, 'anthropic 没配 Key 时不该发请求');

    const dead = createWebAccess(cordisLikeCtx({}), {
        provider: 'auto',
        providerOrder: ['anthropic', 'tavily'],
        providers: { anthropic: { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY' }, tavily: { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY_2' } },
    }, { fetch: async () => { throw new Error('不该走到这里'); } });
    await assert.rejects(() => dead.search('某事件'), (error) => {
        assert.match(error.message, /Anthropic 兼容/);
        assert.match(error.message, /Tavily/);
        assert.match(error.message, /配置缺失/);
        return true;
    });
    // probe 也只看这次真的会试的那几条：点名一条没配好的通道，必须报不可用。
    const pinned = createWebAccess(cordisLikeCtx({}), { provider: 'anthropic', providers: { anthropic: { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY' } } }, { fetch: async () => { throw new Error('x'); } });
    const status = await pinned.probe();
    assert.equal(status.ok, false, '点名 anthropic 但没 Key 时不能说「可用」');
    assert.match(status.reason, /Anthropic 兼容/);
});

// ── 失败分类与自检的如实性（第三十三轮）──────────────────────────────────

await check('同一次调用里，后续查询的失败分类不退化成「其它」', async () => {
    // 第三十三轮实测：一条通道第一次失败记的是「网络出口不可达」，缓存下来的却只有
    // message，第二条查询再命中缓存时被重包成通用错误 → 分类变成「其它」。而分类正是
    // 告诉调用方「该换出口」还是「该换来源」的那句话（web-errors.js 的 P0-9 口径）。
    const access = createWebAccess(cordisLikeCtx({}), {
        provider: 'auto',
        providerOrder: ['duckduckgo', 'searxng'],
        providers: {
            // 回环地址在 SSRF 防线处当场被挡：离线、快，且带确定的错误码。
            duckduckgo: { baseURL: 'http://127.0.0.1:1' },
            searxng: { baseURL: 'http://127.0.0.1:1' },
        },
    });
    const first = await access.search('第一次').then(() => null, (error) => error);
    const second = await access.search('第二次').then(() => null, (error) => error);
    assert.ok(first instanceof WebAccessError, '两条通道都不通时必须抛错');
    assert.ok(second instanceof WebAccessError);
    assert.equal(webFailureKind(second), webFailureKind(first), '两次查询的失败分类必须一致');
    assert.notEqual(webFailureKind(second), 'other', `分类不该退化成「其它」：${second.message.slice(0, 120)}`);
});

await check('probe 不把「只判过配置」的通道说成「可用」', async () => {
    // 免 Key 抓取通道 probe 时不联网，所以「配置齐了」不等于「出口通」。
    // 旧措辞一律说「可用」，于是通道全灭时看起来像「搜了但没资料」。
    const access = createWebAccess(cordisLikeCtx({}), {
        provider: 'auto',
        providerOrder: ['duckduckgo', 'searxng'],
        providers: { duckduckgo: { baseURL: 'https://html.duckduckgo.com/html' } },
    });
    const status = await access.probe();
    assert.equal(status.ok, true, '配置齐了就该算「能试」');
    assert.equal(status.verified, false, '没联网确认过就不能说「已验证」');
    assert.ok(!/「DuckDuckGo（免 Key）」可用/.test(status.reason), `不该说成「可用」：${status.reason}`);
    assert.match(status.reason, /未联网确认/, `要说清只是配置齐了：${status.reason}`);
    // 接缝那条真的问过运行时（服务在不在），才配叫「可用」。
    const seam = createWebAccess(cordisLikeCtx({ web: { search: async () => ({ sources: [] }), fetch: async () => ({}) } }), { provider: 'seam' });
    const seamStatus = await seam.probe();
    assert.equal(seamStatus.verified, true, '接缝是问过运行时的');
});

// ── 网页预处理（第二十轮）────────────────────────────────────────────────

/** 一个「正文 + 导航 + 页脚 + cookie 条」的网页。 */
const NOISY_PAGE = '<html lang="zh-CN"><head><title>噪声页</title></head><body>'
    + '<nav><a href="/">首页</a><a href="/a">栏目甲</a></nav>'
    + '<div class="cookie-consent">我们使用 Cookie 以改善体验，点此接受。</div>'
    + '<main><article><h1>正文标题</h1>'
    + '<p>这是正文的第一段，长度足够长，会被完整保留下来。它讲的是第二十轮检索通道与预处理。</p>'
    + '<p>这是正文的第一段，长度足够长，会被完整保留下来。它讲的是第二十轮检索通道与预处理。</p>'
    + '</article></main>'
    + '<footer>版权所有 京ICP备12345号</footer></body></html>';

await check('预处理：article 模式 —— 取正文去掉导航 / cookie 条 / 页脚，重复段落只留一次', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k', maxRedirects: 0 }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => ({
            status: 200,
            headers: { get: (name) => (name === 'content-type' ? 'text/html; charset=utf-8' : null) },
            body: {
                getReader() {
                    const chunks = [new TextEncoder().encode(NOISY_PAGE)];
                    let index = 0;
                    return { async read() { return index >= chunks.length ? { done: true } : { done: false, value: chunks[index++] }; }, async cancel() {} };
                },
            },
        }),
    });
    const page = await access.fetch('https://j.example.com/9');
    assert.ok(page.content.includes('这是正文的第一段'), '正文要留下');
    assert.ok(!page.content.includes('栏目甲'), '导航条目要丢掉');
    assert.ok(!page.content.includes('Cookie'), 'cookie 条要丢掉');
    assert.ok(!page.content.includes('京ICP'), '页脚要丢掉');
    assert.equal(page.preprocess.mode, 'article');
    assert.ok(page.preprocess.stats.removedRatio > 0.5, `压缩比要报出来，实际 ${page.preprocess.stats.removedRatio}`);
    assert.equal(page.preprocess.stats.dedupedLines, 1, '重复的那一段要去重');
    assert.equal(page.preprocess.meta.title, '噪声页', '标题进 meta');
    assert.equal(page.preprocess.meta.lang, 'zh-CN');
});

await check('预处理：preprocess:off 时原样返回（可一键退回老行为）', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k', maxRedirects: 0, preprocess: { mode: 'article' } }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => ({
            status: 200,
            headers: { get: (name) => (name === 'content-type' ? 'text/html; charset=utf-8' : null) },
            body: {
                getReader() {
                    const chunks = [new TextEncoder().encode(NOISY_PAGE)];
                    let index = 0;
                    return { async read() { return index >= chunks.length ? { done: true } : { done: false, value: chunks[index++] }; }, async cancel() {} };
                },
            },
        }),
    });
    const page = await access.fetch('https://j.example.com/9', { preprocess: 'off' });
    assert.ok(page.content.includes('栏目甲'), 'off 模式不去样板');
    assert.equal(page.preprocess.mode, 'off');
    assert.equal(page.preprocess.stats.removedRatio, 0);
});

await check('预处理不接管成败判定：拦截页即使会被洗干净，也仍判「目标站拒绝」', async () => {
    // 拦截页的正文特征藏在会被清洗掉的区域里：哨兵必须读**清洗前**的文本，
    // 否则这里会被误判成「正文几乎为空」，而两类失败的修法完全不同。
    const page = '<html><body><div class="cookie-consent">Just a moment...</div><main><p>短</p></main></body></html>';
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k', maxRedirects: 0, preprocess: { mode: 'article' } }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => ({
            status: 200,
            headers: { get: (name) => (name === 'content-type' ? 'text/html; charset=utf-8' : null) },
            body: {
                getReader() {
                    const chunks = [new TextEncoder().encode(page)];
                    let index = 0;
                    return { async read() { return index >= chunks.length ? { done: true } : { done: false, value: chunks[index++] }; }, async cancel() {} };
                },
            },
        }),
    });
    await assert.rejects(() => access.fetch('https://k.example.com/10'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_BLOCKED');
        assert.equal(webFailureKind(error), 'blocked');
        return true;
    });
});

await check('失败分类：错误码映射到四类，一行话里带分类名', () => {
    const of = (code) => new WebAccessError('随便一句原因', code);
    assert.equal(webFailureKind(of('OFFICE_WEB_NO_KEY')), 'config');
    assert.equal(webFailureKind(of('OFFICE_WEB_NETWORK')), 'network');
    assert.equal(webFailureKind(of('OFFICE_WEB_TIMEOUT')), 'network');
    assert.equal(webFailureKind(of('OFFICE_WEB_REDIRECT')), 'network');
    assert.equal(webFailureKind(of('OFFICE_WEB_BLOCKED')), 'blocked');
    assert.equal(webFailureKind(of('OFFICE_WEB_EMPTY')), 'empty');
    assert.equal(webFailureKind(of('OFFICE_WEB_NO_RESULTS')), 'empty');
    assert.equal(webFailureKind(of('OFFICE_WEB_我没见过的码')), 'other', '没登记的一律落 other，不猜');
    assert.equal(webFailureKind(new Error('没有 code 的普通异常')), 'other');
    assert.equal(webFailureLabel(of('OFFICE_WEB_BLOCKED')), '目标站拒绝');
    assert.match(describeWebFailure(of('OFFICE_WEB_NETWORK')), /^网络出口不可达：/);
    assert.equal(
        summarizeWebFailures([of('OFFICE_WEB_BLOCKED'), of('OFFICE_WEB_BLOCKED'), of('OFFICE_WEB_NETWORK')]),
        '目标站拒绝 2、网络出口不可达 1',
    );
    assert.equal(summarizeWebFailures([]), '');
});

await check('漂移守卫：src/web*.js 里抛出的每个错误码都已登记分类', async () => {
    // 新加错误码却忘了给分类，会静默落进 other，「四类分开报」就退化了 —— 这条守住它。
    // 第二十轮起错误码分布在编排层（web.js）、通道层（web-providers.js）与错误层
    // （web-errors.js），所以扫全部 web*.js 文件，不是一个文件。
    const files = ['../src/web.js', '../src/web-providers.js', '../src/web-errors.js', '../src/web-preprocess.js'];
    const thrown = new Set();
    for (const file of files) {
        const source = await readFile(new URL(file, import.meta.url), 'utf8');
        for (const matched of source.matchAll(/'(OFFICE_WEB_[A-Z_]+)'/g)) thrown.add(matched[1]);
    }
    assert.ok(thrown.size >= 10, `没扫到错误码，正则大概不对了（扫到 ${thrown.size} 个）`);
    const missing = [...thrown].filter((code) => !WEB_FAILURE_CODES.includes(code));
    assert.deepEqual(missing, [], `这些错误码没有登记分类：${missing.join('、')}`);
});

// ── 取正文撞上 PDF（第二十九轮）────────────────────────────────────────────

await check('PDF 识别：MIME 说了算；含糊时才看 URL 后缀', () => {
    assert.equal(looksLikePdf('application/pdf', 'https://a.example.com/x'), true);
    assert.equal(looksLikePdf('application/pdf; charset=binary', 'https://a.example.com/x'), true);
    assert.equal(looksLikePdf('', 'https://a.example.com/x.pdf'), true);
    assert.equal(looksLikePdf('application/octet-stream', 'https://a.example.com/x.pdf?dl=1'), true);
    assert.equal(looksLikePdf('', 'https://a.example.com/x'), false);
    assert.equal(looksLikePdf('text/html', 'https://a.example.com/x.pdf'), false, '站点说是 HTML 就按 HTML 处理，不按后缀猜');
});

/** 一个 PDF 响应的替身。 */
function pdfResponse(bytes) {
    return {
        status: 200,
        headers: { get: (name) => (name === 'content-type' ? 'application/pdf' : null) },
        body: {
            getReader() {
                let sent = false;
                return {
                    async read() {
                        if (sent) return { done: true, value: undefined };
                        sent = true;
                        return { done: false, value: bytes };
                    },
                    async cancel() {},
                };
            },
        },
    };
}

await check('取正文：application/pdf 交给 office.pdf 抽文本，正文进 content、页数进 pdf', async () => {
    const seen = {};
    const extracted = '第一章 不定积分\n\n1. 求 ∫x dx';
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k', pdfMaxBytes: 4 * 1024 * 1024 }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => pdfResponse(new TextEncoder().encode('%PDF-1.4 fake')),
        pdf: {
            async text(path, options) {
                seen.path = path;
                seen.maxPages = options.maxPages;
                seen.maxBytes = options.maxBytes;
                return { text: extracted, pages: 3, extractor: 'stub' };
            },
        },
    });
    const page = await access.fetch('https://a.example.com/paper.pdf');
    assert.equal(page.kind, 'pdf');
    assert.equal(page.content, extracted);
    assert.deepEqual(page.pdf, { pages: 3, chars: extracted.length, bytes: 13, extractor: 'stub' });
    assert.match(page.notice, /PDF（3 页，抽取引擎 stub）/);
    assert.equal(page.preprocess, undefined, 'PDF 不做网页预处理');
    assert.ok(seen.path.endsWith('.pdf'), '抽文本前要落成一个 .pdf 临时文件');
    assert.equal(seen.maxPages, 30, '页数上限来自配置');
});

await check('取正文：PDF 超体积上限直接放弃（截断的 PDF 抽不出文本）', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k', pdfMaxBytes: 1024 * 1024 }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        // 声明 2 MB，实际只给 1 MB 上限的量：读满上限即判截断。
        fetch: async () => pdfResponse(new Uint8Array(1024 * 1024 + 10)),
        pdf: { async text() { throw new Error('超上限时不该走到抽取'); } },
    });
    await assert.rejects(() => access.fetch('https://a.example.com/big.pdf'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_PDF_TOO_BIG');
        assert.match(error.message, /超过 1 MB 上限/);
        return true;
    });
});

await check('取正文：PDF 抽不出文本按「没拿到结果」报，并带出扫描件提示', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => pdfResponse(new TextEncoder().encode('%PDF-1.4')),
        pdf: { async text() { return { text: '   ', pages: 8, extractor: 'fitz', hint: '这份 PDF 没有文本层（扫描/手写）：改用 office.pdf.pages() 渲染成图片。' }; } },
    });
    await assert.rejects(() => access.fetch('https://a.example.com/scan.pdf'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_EMPTY');
        assert.equal(webFailureKind(error), 'empty');
        assert.match(error.message, /没有文本层/);
        return true;
    });
});

await check('取正文：没有 PDF 抽取通道时按「配置缺失」报，不假装读到了页', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => pdfResponse(new TextEncoder().encode('%PDF-1.4')),
        // 注意用 `{}` 而不是 `undefined`：`hooks.pdf ?? defaultNetwork.pdf` 会把 undefined
        // 换成真的读取器，用例就测不到这个守卫了（审查 P1-4）。
        pdf: {},
    });
    await assert.rejects(() => access.fetch('https://a.example.com/x.pdf'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_PDF_ENGINE');
        assert.equal(webFailureKind(error), 'config');
        assert.match(error.message, /没有可用的抽取通道/, '要断言守卫自己的话，别靠 pdftotext|PyMuPDF 碰巧命中');
        return true;
    });
});

await check('取正文：标着 PDF 但内容不是 PDF（HTML 错误页）按「不支持的类型」报', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => ({
            status: 200,
            headers: { get: (name) => (name === 'content-type' ? 'application/octet-stream' : null) },
            body: {
                getReader() {
                    let sent = false;
                    const bytes = new TextEncoder().encode('<!doctype html><title>404</title>');
                    return {
                        async read() {
                            if (sent) return { done: true, value: undefined };
                            sent = true;
                            return { done: false, value: bytes };
                        },
                        async cancel() {},
                    };
                },
            },
        }),
        pdf: { async text() { throw new Error('不是 PDF，不该走到抽取'); } },
    });
    await assert.rejects(() => access.fetch('https://a.example.com/error.pdf'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_UNSUPPORTED_TYPE');
        assert.match(error.message, /内容却不是 PDF/);
        return true;
    });
});

await check('取正文：抽取器抛错时归类为「配置缺失」并带上原因', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => pdfResponse(new TextEncoder().encode('%PDF-1.4')),
        pdf: { async text() { throw new Error('抽文本失败（可用引擎：无）：没有引擎'); } },
    });
    await assert.rejects(() => access.fetch('https://a.example.com/x.pdf'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_PDF_ENGINE');
        assert.match(error.message, /可用引擎：无/);
        return true;
    });
});

await check('取正文：PDF 文本很短时给提醒（不判失败），响应体为空时明确报「没拿到结果」', async () => {
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => pdfResponse(new TextEncoder().encode('%PDF-1.4')),
        pdf: { async text() { return { text: '封面', pages: 1, extractor: 'stub' }; } },
    });
    const page = await access.fetch('https://a.example.com/cover.pdf');
    assert.match(page.notice, /文本很短（2 字符）/, '短文本只提醒，不当失败');
    assert.equal(page.content, '封面');

    const empty = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async () => pdfResponse(new Uint8Array(0)),
        pdf: { async text() { throw new Error('空响应体不该走到抽取'); } },
    });
    await assert.rejects(() => empty.fetch('https://a.example.com/empty.pdf'), (error) => {
        assert.equal(error.code, 'OFFICE_WEB_EMPTY');
        assert.match(error.message, /响应体是空的/);
        return true;
    });
});

// ── 配置收敛 ──────────────────────────────────────────────────────────────

await check('内置检索配置在安全区间内收敛，坏值退回默认', () => {
    const options = resolveBuiltinOptions({ maxResults: 999, fetchTimeoutMs: 1, maxRedirects: -3, baseURL: 'https://x.example.com/v1/', apiKeyEnv: '  K  ', searchTimeoutMs: '不是数字' });
    assert.equal(options.maxResults, 20, '超上限的被夹到 20');
    assert.equal(options.fetchTimeoutMs, 5_000, '低于下限的被抬到 5 秒');
    assert.equal(options.maxRedirects, 0, '负数夹到 0（等于不跟随跳转）');
    assert.equal(options.searchTimeoutMs, 60_000, '解析不出来的退回默认');
    assert.equal(options.baseURL, 'https://x.example.com/v1', '末尾斜杠要去掉，避免拼出 //messages');
    assert.equal(options.apiKeyEnv, 'K');
    // 第二十九轮：PDF 两项也要收敛（1 MB 下限 / 128 MB 上限 / 500 页上限）。
    const pdf = resolveBuiltinOptions({ pdfMaxBytes: 1, pdfMaxPages: 9_999 });
    assert.equal(pdf.pdfMaxBytes, 1024 * 1024, '低于下限抬到 1 MB');
    assert.equal(pdf.pdfMaxPages, 500, '超上限夹到 500 页');
    assert.equal(resolveBuiltinOptions({}).pdfMaxBytes, 24 * 1024 * 1024, '默认 24 MB');
    assert.equal(resolveBuiltinOptions({}).pdfMaxPages, 30, '默认 30 页');
});

// ── 收尾 ──────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`web: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
