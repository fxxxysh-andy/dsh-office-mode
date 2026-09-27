/**
 * 检索能力的单元测试：渠道分流、提纲生成、解析与跨源核对、派工守卫。
 *
 * 这些是不依赖网络的纯逻辑测试 —— 检索本身要联网，但「该走哪些渠道」
 * 「结果有没有出处」「子代理有没有偷懒」这些判断都是本地的，可以钉死。
 *
 * 跑法：node test/search.mjs
 */
import { mkdtempSync, rmSync, mkdirSync, writeFileSync, readFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { buildBrief, CONTENT_TYPES, contentType, contentTypeIds, guessContentType, materializeQueries } from '../src/search-routes.js';
import { parseFindings, renderFindings } from '../src/findings.js';
import { buildChannelPrompt, CHANNEL_TOOLS, defaultOutputPaths, describeSubagentRefusal, dispatchSearch, missingChannelTools, missingToolsFromRefusal, NO_SUBAGENTS_HINT, parseResultFiles, renderDispatch, resetToolGapMemo, runChannelBuiltin, subagentsOf, writeBrief } from '../src/search.js';
import { createWebAccess } from '../src/web.js';
import { buildHelp } from '../src/docs.js';
import { resetHints } from '../src/projection.js';
import { buildTools } from '../src/tools.js';
import { resolveConfig } from '../src/config.js';

const results = [];
async function check(name, fn) {
    // 「这条长指引说过没有」是进程内状态：每个用例都从「没说」开始，
    // 才能既验「第一次会带指引」又验「第二次不再重复」。
    resetHints();
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/**
 * 造一个「行为与真实 cordis 一致」的假插件上下文。
 *
 * 真实 cordis 的 ctx 是个代理：读一个**没写进 inject** 的服务属性会当场抛错
 * （`cannot get property "subagents" without inject`），而不是返回 undefined；
 * 服务只能经 `ctx.get(name)` 取（cordis 文档原话：Read a service from the store
 * without the inject requirement）。
 *
 * 这个替身就是为这条行为存在的：以前测试里用普通对象 `{ subagents: … }` 当 ctx，
 * 恰好把 `ctx.subagents` 这条**在真实会话里必然抛错**的路径测成了「能读到」，
 * 于是 `office_search_dispatch` 上线起就没跑通过、单测却一直是绿的
 * （2026-09-25 复现与修复）。真实 cordis 的对照实验见 test/subagent-seam.mjs。
 */
function cordisLikeCtx(services = {}) {
    return new Proxy({ inject: ['tools'] }, {
        get(target, prop) {
            if (prop === 'get') return (name) => services[name];
            if (prop in target) return target[prop];
            throw new Error(`cannot get property "${String(prop)}" without inject`);
        },
    });
}

/**
 * 假的内置检索接缝（离线）：检索返回两条带 URL 的来源，取正文返回一段 HTML。
 *
 * 内置通道在宿主里优先用 ctx.web，拿不到才自己发 HTTP。测试钉的是「通道选择与
 * 落盘格式」，所以在这里把接缝喂成假的就够，不必真的联网。
 */
function fakeSeam() {
    return {
        async search(request) {
            return {
                sources: [
                    { url: 'https://news-a.example.com/' + encodeURIComponent(request.query), title: '来源甲 ' + request.query, snippet: '引用片段' },
                    { url: 'https://news-b.example.com/2', title: '来源乙' },
                ],
                truncated: false,
            };
        },
        async fetch(request) {
            return {
                url: request.url,
                statusCode: 200,
                body: { kind: 'html', content: '<p>来源乙的正文段落，长度足够当摘录用的一行文字。</p>' },
                truncated: false,
            };
        },
    };
}

const root = mkdtempSync(join(tmpdir(), 'office-tmp-search-'));

// ── 渠道分流 ──────────────────────────────────────────────────────────────

await check('三类内容类型齐备，且每类都有必须覆盖的渠道', () => {
    const ids = contentTypeIds();
    for (const wanted of ['hotspot', 'knowledge', 'manual', 'mixed']) {
        assert.ok(ids.includes(wanted), `缺少内容类型 ${wanted}`);
    }
    for (const type of CONTENT_TYPES) {
        if (type.id === 'mixed') continue;
        const required = type.channels.filter((channel) => channel.required);
        assert.ok(required.length >= 2, `${type.id} 的必须渠道少于 2 个，等于没有分流`);
    }
});

await check('每个类型的第一条渠道都是不限定来源的泛搜（防茧房）', () => {
    for (const type of CONTENT_TYPES) {
        const first = type.channels[0];
        assert.equal(first.engine, 'web', `${type.id} 的第一条不是泛搜`);
        assert.ok(first.kind.includes('泛搜'), `${type.id} 的第一条不叫泛搜：${first.kind}`);
    }
});

await check('热点走权威媒体 + 社交平台两头', () => {
    const type = contentType('hotspot');
    const kinds = type.channels.map((c) => c.kind).join('|');
    assert.ok(type.channels.some((c) => c.kind.includes('权威媒体') && c.required), '热点缺权威媒体必查渠道');
    assert.ok(type.channels.some((c) => c.kind.includes('社交平台') && c.required), '热点缺社交平台必查渠道');
    assert.ok(type.channels.filter((c) => c.kind.includes('社交平台')).length >= 2, '热点至少要覆盖两个社交平台');
});

await check('知识类走百科，并能下沉到文献', () => {
    const type = contentType('knowledge');
    assert.ok(type.channels.some((c) => c.platform === 'wikipedia' && c.required), '知识类缺百科必查渠道');
    assert.ok(type.channels.some((c) => c.kind.includes('文献')), '知识类缺文献下沉渠道');
    assert.ok(Array.isArray(type.depthRule) && type.depthRule.length >= 3, '知识类应有浅/中/深三档深度规则');
});

await check('手册类只认官方平台参考文档', () => {
    const type = contentType('manual');
    assert.ok(type.channels.some((c) => c.engine === 'fetch' && c.required), '手册类必须要求取回官方文档原文');
    assert.ok(type.channels.some((c) => c.platform === 'github' && c.required), '手册类要覆盖官方仓库');
    const community = type.channels.filter((c) => c.platform === 'stackoverflow');
    assert.ok(community.length > 0 && community.every((c) => !c.required), '社区渠道只能作参考，不能是必须项');
});

await check('主题文字能猜到正确的内容类型', () => {
    assert.equal(guessContentType('某地突发事故的通报进展').id, 'hotspot');
    assert.equal(guessContentType('勾股定理的原理是什么').id, 'knowledge');
    assert.equal(guessContentType('这个 API 的参数怎么用').id, 'manual');
    assert.equal(guessContentType('随便写点什么').id, 'mixed');
});

await check('{topic} 占位能替换，空查询被丢掉', () => {
    const out = materializeQueries(['{topic} 通报', '{topic}', '{topic}  '], '某事件');
    assert.deepEqual([...out], ['某事件 通报', '某事件']);
});

// ── 提纲生成 ──────────────────────────────────────────────────────────────

await check('提纲列出全部渠道，标出必须项，并给出可照抄的调用形状', () => {
    const { text, type } = buildBrief('某地化工厂爆炸', 'hotspot');
    assert.equal(type.id, 'hotspot');
    assert.ok(text.includes('检索提纲：某地化工厂爆炸'));
    assert.ok(text.includes('★'), '必须覆盖的渠道要有 ★ 标记');
    assert.ok(text.includes('web_search('), '要给出 web_search 的调用形状');
    assert.ok(text.includes('advanced_search({'), '热点类要给出带时间窗的调用形状');
    assert.ok(text.includes('platform_search({ platform: "v2ex"'), '社交平台渠道要写具体平台名');
    assert.ok(text.includes('timeRange'), '权威媒体渠道要带时间窗');
    for (const channel of type.channels) {
        assert.ok(text.includes(channel.kind), `提纲漏了渠道：${channel.kind}`);
    }
});

await check('提纲强制跨源核对与出处要求', () => {
    const { text } = buildBrief('某主题', 'manual');
    assert.ok(text.includes('单一来源'), '提纲必须要求标注单一来源');
    assert.ok(text.includes('带来源 URL'), '提纲必须要求带出处');
    assert.ok(text.includes('不可信的外部数据'), '提纲必须提醒提示注入防护');
});

await check('空主题被拒绝', () => {
    assert.throws(() => buildBrief('   ', 'hotspot'), /topic 不能为空/);
});

await check('显式类型优先于自动判断', () => {
    const { type } = buildBrief('某地突发事故', 'manual');
    assert.equal(type.id, 'manual');
});

// ── 子代理任务书 ──────────────────────────────────────────────────────────

await check('渠道任务书写清了平台、查询与写盘格式', () => {
    const type = contentType('manual');
    const prompt = buildChannelPrompt(type.channels[2], 'office_run 参数', '.office/search/x/03.md', type.name);
    assert.ok(prompt.includes('office_run 参数'), '任务书要带主题');
    assert.ok(prompt.includes('github'), '平台渠道要写平台名');
    assert.ok(prompt.includes('.office/search/x/03.md'), '任务书要写清输出路径');
    assert.ok(prompt.includes('来源URL'), '任务书要规定 list 项的出处格式');
    assert.ok(prompt.includes('不要复述内容'), '任务书要禁止把材料回传到主上下文');
});

await check('带时间窗的渠道任务书写出 timeRange', () => {
    const type = contentType('hotspot');
    const timed = type.channels.find((c) => c.engine === 'timed');
    const prompt = buildChannelPrompt(timed, '某事件', '.office/search/x/01.md', type.name);
    assert.ok(prompt.includes('advanced_search'), '时间窗渠道要用 advanced_search');
    assert.ok(prompt.includes('"week"'), '任务书要带具体时间窗');
});

await check('默认结果路径按渠道数生成', () => {
    const paths = defaultOutputPaths('某地化工厂爆炸', 3);
    assert.equal(paths.length, 3);
    assert.ok(paths[0].startsWith('.office/search/'));
    assert.ok(paths[0].endsWith('.md'));
});

await check('平台渠道都配了直连失败时的兜底查询', () => {
    for (const type of CONTENT_TYPES) {
        for (const channel of type.channels) {
            if (channel.engine !== 'platform') continue;
            assert.ok(Array.isArray(channel.fallbackQueries) && channel.fallbackQueries.length > 0,
                `${type.id} 的 ${channel.kind} 没有兜底查询；平台直连一旦被网络拦住就整条渠道失效`);
        }
    }
});

await check('提纲写出平台渠道的兜底路径与降级标注要求', () => {
    const { text } = buildBrief('某地化工厂爆炸', 'hotspot');
    assert.ok(text.includes('经搜索引擎间接取得'), '要要求标注降级来源');
    assert.ok(text.includes('直连失败时改用 web_search'), '要写出兜底路径');
});

await check('平台渠道的任务书写明先直连、失败再兜底', () => {
    const type = contentType('knowledge');
    const wiki = type.channels.find((c) => c.platform === 'wikipedia');
    const prompt = buildChannelPrompt(wiki, '勾股定理', '.office/search/x/02.md', type.name);
    assert.ok(prompt.includes('首选直连该平台'), '要先试直连');
    assert.ok(prompt.includes('如果直连报错'), '要有失败分支');
    assert.ok(prompt.includes('经搜索引擎间接取得'), '要要求标注降级');
    assert.ok(prompt.includes('勾股定理 维基百科'), '要给出具体兜底查询');
});

// ── 落盘与三步走的接缝 ────────────────────────────────────────────────────
//
// 这一节是补的：以前只测了 buildBrief（纯函数），没人真的执行过 writeBrief
// 与 office_search_brief 工具，所以「参数 text 遮蔽同名辅助函数、调用它时报
// text is not a function」这种一到手就整条失效的 bug 一路漏到了用户面前。

await check('writeBrief 能落盘，并按 outputDir 与主题收敛路径', async () => {
    const briefPath = await writeBrief(root, '某地化工厂爆炸', 'hotspot', '写一份汇报', '正文一行', '.office/search');
    assert.equal(typeof briefPath, 'string');
    assert.ok(briefPath.endsWith(join('某地化工厂爆炸', 'brief.md')), `落盘路径不对：${briefPath}`);
    assert.ok(existsSync(briefPath), '提纲文件必须真的写出来');
    const onDisk = readFileSync(briefPath, 'utf8');
    assert.ok(onDisk.includes('正文一行'), '提纲正文要原样写入');
    assert.ok(onDisk.includes('# 检索提纲'), '要有标题头');

    // outputDir 优先于默认目录；主题里的路径非法字符要被收敛掉。
    const custom = await writeBrief(root, 'a/b:c*d', 'hotspot', undefined, 'x', '我的提纲');
    assert.ok(custom.includes(join('我的提纲', 'a-b-c-d', 'brief.md')), `自定义目录与 slug 不对：${custom}`);
});

await check('office_search_brief 端到端：出提纲、落盘、给出下一步', async () => {
    const tools = buildTools(resolveConfig({}));
    const tool = tools.find((t) => t.name === 'office_search_brief');
    const exec = { agent: { session: { header: { cwd: root } } } };
    const value = await tool.execute({ topic: '某地化工厂爆炸', type: 'hotspot', audience: '写一份 12 页汇报' }, exec);

    assert.equal(value.ok, true);
    assert.equal(value.type, 'hotspot');
    assert.equal(value.typeName, '热点事件');
    assert.ok(value.briefPath.endsWith('brief.md'), `briefPath 不对：${value.briefPath}`);
    const onDisk = readFileSync(join(root, value.briefPath), 'utf8');
    assert.ok(onDisk.includes('检索提纲：某地化工厂爆炸'), '落盘内容要含主题');
    assert.ok(onDisk.includes('★'), '落盘内容要含必须覆盖标记');
    assert.ok(value.text.includes(value.briefPath), '反馈里要写清提纲落在哪');
    assert.ok(value.text.includes('office_search_dispatch'), '反馈要指向派工这一步');
});

await check('brief 落盘的提纲能被派工读回（第一步与第二步的接缝）', async () => {
    // 派工不单独收主题与类型，是回头从提纲文件里正则反推的。这一条钉的就是
    // 「写出来的提纲，读得回来」——两边的格式约定一旦漂移，这里先红。
    const tools = buildTools(resolveConfig({}));
    const tool = tools.find((t) => t.name === 'office_search_brief');
    const exec = { agent: { session: { header: { cwd: root } } } };
    const brief = await tool.execute({ topic: '某地化工厂爆炸', type: 'hotspot' }, exec);

    const type = contentType('hotspot');
    const outputs = type.channels.map((_, i) => `seam-${i}.md`);
    const ctx = cordisLikeCtx({
        subagents: {
            async start() {
                return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} };
            },
        },
    });
    const out = await dispatchSearch(
        { briefPath: brief.briefPath, outputPaths: outputs },
        { agent: exec.agent, signal: new AbortController().signal },
        {},
        ctx,
    );
    assert.equal(out.topic, '某地化工厂爆炸', '派工要从提纲文件里反推出主题');
    assert.equal(out.jobs.length, type.channels.length, '渠道数要与提纲一致');
});

// ── 结果解析 ──────────────────────────────────────────────────────────────

const SAMPLE = [
    '# 检索结果：某地化工厂爆炸',
    '',
    '## 泛搜',
    '- 化工厂爆炸造成 3 人受伤，官方已介入 (https://news-a.example.com/1)',
    '',
    '## 权威媒体',
    '- 化工厂爆炸造成 3 人受伤，官方已介入 (https://news-b.example.com/2)',
    '- 涉事企业三年前曾被处罚 (https://news-c.example.com/3)',
    '- 另一种说法：伤亡数字为 0',
    '',
    '## 社交平台：中文社区讨论（platform: v2ex）',
    '- 涉事企业三年前曾被处罚，处罚文书可查 (https://v2ex.example.com/t/9)',
    '',
].join('\n');

await check('解析抽出结论、出处与来源站点', () => {
    const res = parseFindings(SAMPLE, 'hotspot', '某地化工厂爆炸');
    assert.equal(res.findings.length, 4);
    assert.equal(res.risks.distinctHosts, 4);
    assert.equal(res.risks.noUrl, 1, '有一条没带 URL');
});

await check('URL 从结论正文里摘掉，不会跟着抄进文档', () => {
    const res = parseFindings(SAMPLE, 'hotspot', '某地化工厂爆炸');
    const withUrl = res.findings.find((f) => f.claim.includes('3 人受伤'));
    assert.ok(!withUrl.claim.includes('https://'), '结论正文里不应残留 URL');
    assert.equal(withUrl.urls.length, 2, '两个来源都要收进来');
});

await check('同一说法多站点支撑时不再判为单一来源', () => {
    const res = parseFindings(SAMPLE, 'hotspot', '某地化工厂爆炸');
    // 同一个事实在两个互不相关的站点出现 → 不标单一来源。
    const corroborated = res.findings.find((f) => f.claim.includes('3 人受伤'));
    assert.equal(corroborated.singleSource, false, '两个站点背书就不该是单一来源');
    assert.equal(corroborated.hosts.length, 2, '两个来源都要记进来');
    // 措辞不同但说的是同一件事时也要合并：否则同一事实会被算成两组、
    // 每一组都只有单一来源，跨源核对就形同虚设。
    const reworded = res.findings.find((f) => f.claim.includes('处罚文书可查'));
    assert.equal(reworded.singleSource, false, '换句话说的同一事实应与原说法合簇');
    assert.ok(res.risks.corroborated >= 2, '两件事都该算跨源核对成功');
    assert.ok(res.risks.unverified >= 1, '没带出处的条目仍要计入待核');
});

await check('覆盖度认出类别，但不会把权威媒体算成回看时间窗那条', () => {
    const res = parseFindings(SAMPLE, 'hotspot', '某地化工厂爆炸');
    assert.ok(res.coverage.covered.some((c) => c.includes('泛搜')), '泛搜应算已覆盖');
    assert.ok(res.coverage.covered.some((c) => c.includes('v2ex')), '平台渠道应算已覆盖');
    const plainMedia = res.coverage.covered.filter((c) => c.includes('权威媒体'));
    assert.ok(plainMedia.some((c) => !c.includes('回看')), '权威媒体主渠道应算已覆盖');
});

await check('缺渠道会被列成必须补的项', () => {
    const onlyOverview = ['## 泛搜', '- 只有概览 (https://a.example.com/1)'].join('\n');
    const res = parseFindings(onlyOverview, 'hotspot', '某事件');
    assert.ok(res.coverage.missing.length >= 1, '只搜了泛搜，必须报出仍缺的渠道');
    assert.ok(res.notes.some((n) => n.includes('仍缺') || n.includes('还缺')), '缺渠道要写进提示');
});

await check('全部结论来自一两个站点时提示茧房风险', () => {
    const narrow = ['## 泛搜', '- 说法一 (https://same.example.com/1)', '- 说法二 (https://same.example.com/2)'].join('\n');
    const res = parseFindings(narrow, 'hotspot', '某事件');
    assert.ok(res.notes.some((n) => n.includes('茧房')), '单一站点要提示茧房风险');
});

await check('空文件不抛错，而是给出可纠正的提示', () => {
    const res = parseFindings('# 标题\n\n没有列表项', 'hotspot', '某事件');
    assert.equal(res.findings.length, 0);
    assert.ok(res.notes.some((n) => n.includes('没有解析出任何结论')));
});

await check('渲染摘要是紧凑的，不是原文搬运', () => {
    const res = parseFindings(SAMPLE, 'hotspot', '某地化工厂爆炸');
    const text = renderFindings(res);
    assert.ok(text.includes('到手结论'), '摘要要有总数');
    assert.ok(text.includes('已跨源核对'), '摘要要报告跨源核对结果');
    assert.ok(text.length < SAMPLE.length * 2, '摘要不该比原文还长');
});

// ── 派工与多文件解析 ──────────────────────────────────────────────────────

await check('派工拒绝空 outputPaths（先出提纲是硬约束）', async () => {
    const briefPath = join(root, 'brief.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    await assert.rejects(
        () => dispatchSearch({ briefPath, outputPaths: [] }, { agent: {}, signal: new AbortController().signal }, {}, undefined),
        /outputPaths 不能为空/,
    );
});

await check('outputPaths 与渠道数对不上时明确报错', async () => {
    const briefPath = join(root, 'brief2.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    await assert.rejects(
        () => dispatchSearch({ briefPath, outputPaths: ['a.md'] }, { agent: {}, signal: new AbortController().signal }, {}, {}),
        /一一对上/,
    );
});

await check('没有 subagents 服务时自动改用内置检索，不再直接失败', async () => {
    // 办公 preset 就是这样：工具面里没有 web_search，也没有子代理。
    // 内置通道必须自己把渠道跑完、把结果文件写出来。
    const type = contentType('hotspot');
    const briefPath = join(root, 'brief3.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `fallback-${i}.md`);
    const parent = { session: { header: { cwd: root } } };
    const out = await dispatchSearch(
        { briefPath, outputPaths: outputs },
        { agent: parent, signal: new AbortController().signal },
        {},
        cordisLikeCtx({ web: fakeSeam() }),
    );
    assert.equal(out.ok, true, '有内置检索时不该整批失败');
    assert.equal(out.engine, 'builtin');
    for (const [index, job] of out.jobs.entries()) {
        assert.equal(job.via, 'builtin', '每个渠道都该走内置通道');
        const text = readFileSync(join(root, outputs[index]), 'utf8');
        assert.ok(text.includes('## ' + type.channels[index].kind), `结果文件要有渠道分组：${type.channels[index].kind}`);
        assert.ok(text.includes('https://'), '结果文件里要有来源 URL');
        assert.ok(!text.startsWith('## '), '文件头要有说明，不是直接甩结论');
    }
});

await check('子代理因工具白名单对不上而失败时，自动回退到内置检索', async () => {
    // 这一条钉的就是 2026-09-25 真实会话里的报错：
    // tools.restrict() names unknown global tools "web_search", …
    // 组合里没有 web 工具时，子代理一启动就抛这个错。派工必须自己兜住。
    const type = contentType('hotspot');
    const briefPath = join(root, 'brief-fallback.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `restrict-${i}.md`);
    const failing = {
        async start() {
            throw new Error('tools.restrict() names unknown global tools "web_search", "advanced_search", "platform_search", "web_fetch"');
        },
    };
    const out = await dispatchSearch(
        { briefPath, outputPaths: outputs },
        { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal },
        {},
        cordisLikeCtx({ subagents: failing, web: fakeSeam() }),
    );
    assert.equal(out.ok, true);
    assert.ok(out.jobs.every((job) => job.via === 'builtin'), '回退后每个渠道都该走内置通道');
    assert.ok(out.jobs[0].notes.some((note) => note.includes('组合里没有联网检索工具')), '反馈里要留下回退原因');
    assert.ok(out.jobs[0].notes.some((note) => note.includes('web_search')), '回退原因要说清缺哪几个工具，不能收短成一句空话');
    assert.ok(!renderDispatch(out).includes('known global tools'), '「已知全局工具」那份清单对模型没有信息量，不该贴出去');
    const rendered = renderDispatch(out);
    assert.ok(rendered.includes('内置检索'), '反馈要说明这条结果是内置检索拿的');
});

await check('engine=builtin 时不起子代理，engine=subagent 时只派子代理', async () => {
    const type = contentType('mixed');
    const briefPath = join(root, 'brief-engine.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：混合或未指明（mixed）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `engine-${i}.md`);
    let started = 0;
    const counting = {
        async start() {
            started += 1;
            return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} };
        },
    };
    const builtinOnly = await dispatchSearch(
        { briefPath, outputPaths: outputs },
        { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal },
        { search: { engine: 'builtin' } },
        cordisLikeCtx({ subagents: counting, web: fakeSeam() }),
    );
    assert.equal(builtinOnly.ok, true);
    assert.equal(started, 0, 'engine=builtin 时一个子代理都不该起');

    const subagentOnly = await dispatchSearch(
        { briefPath, outputPaths: type.channels.map((_, i) => `engine-s-${i}.md`) },
        { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal },
        { search: { engine: 'subagent' } },
        cordisLikeCtx({ subagents: counting, web: fakeSeam() }),
    );
    assert.equal(subagentOnly.ok, true);
    assert.equal(started, type.channels.length, 'engine=subagent 时每渠道一个子代理');
    assert.ok(subagentOnly.jobs.every((job) => job.via === 'subagent'));
});

// ── 装配期探工具面（第十八轮 P1-11）────────────────────────────────────────

await check('装配期探工具面：缺哪些检索工具问得出来，问不出来就不下结论', () => {
    const withSchemas = (names) => cordisLikeCtx({ tools: { schemas: () => names.map((name) => ({ name })) } });
    const officeLike = ['ask_user_question', 'edit', 'glob', 'grep', 'office_help', 'office_memory', 'office_run', 'present', 'read', 'read_image', 'write'];
    assert.deepEqual(
        missingChannelTools(withSchemas(officeLike)),
        ['web_search', 'advanced_search', 'platform_search', 'web_fetch'],
        '办公 preset 缺的正是那四个联网工具（read / read_image / write 是齐的）',
    );
    assert.deepEqual(missingChannelTools(withSchemas(CHANNEL_TOOLS)), [], '齐备时返回空数组，不是 undefined');
    assert.equal(missingChannelTools(cordisLikeCtx({})), undefined, '拿不到 tools 服务 → 不知道，照旧试一次');
    assert.equal(missingChannelTools(cordisLikeCtx({ tools: {} })), undefined, '没有 schemas → 不知道');
    assert.equal(missingChannelTools(cordisLikeCtx({ tools: { schemas: () => { throw new Error('boom'); } } })), undefined, 'schemas 抛错 → 不知道');
    assert.equal(missingChannelTools(cordisLikeCtx({ tools: { schemas: () => [] } })), undefined, '一个名字都读不到 → 不知道，别误判成「组合里什么都没有」');
    assert.equal(missingChannelTools(undefined), undefined);
});

await check('tools.restrict() 的报错收成一句事实，缺失的工具名不丢', () => {
    const raw = new Error('tools.restrict() names unknown global tools "web_search", "advanced_search", "platform_search", "web_fetch"; known global tools: ask_user_question, edit, glob, grow');
    assert.deepEqual(missingToolsFromRefusal(raw), ['web_search', 'advanced_search', 'platform_search', 'web_fetch']);
    const short = describeSubagentRefusal(raw);
    assert.ok(short.includes('web_search') && short.includes('web_fetch'), '收短后仍要看得出缺哪几个');
    assert.ok(!short.includes('known global tools'), '那份「已知全局工具」清单对模型没有信息量');
    assert.ok(short.length < raw.message.length / 2, `要真的短下来：${short.length} vs ${raw.message.length}`);
    // 不是这一类失败就原样返回，别把别的错也说成「组合缺工具」。
    assert.equal(describeSubagentRefusal(new Error('子代理超时')), '子代理超时');
    assert.equal(missingToolsFromRefusal(new Error('子代理超时')), undefined);
});

await check('组合缺检索工具：不再逐渠道去撞失败，反馈里只说一次', async () => {
    resetToolGapMemo();
    const type = contentType('hotspot');
    const briefPath = join(root, 'brief-gap.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `gap-${i}.md`);
    let started = 0;
    const counting = { async start() { started += 1; return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} }; } };
    // 办公 preset 的形态：subagents 服务在，但工具面里没有那四个联网工具。
    const officeTools = {
        schemas: () => ['ask_user_question', 'edit', 'glob', 'grep', 'office_help', 'office_run', 'present', 'read', 'read_image', 'write']
            .map((name) => ({ name })),
    };
    const out = await dispatchSearch(
        { briefPath, outputPaths: outputs },
        { agent: { session: { header: { cwd: root }, id: 'session-gap' } }, signal: new AbortController().signal },
        {},
        cordisLikeCtx({ subagents: counting, tools: officeTools, web: fakeSeam() }),
    );

    assert.equal(out.ok, true, '内置通道照样该把渠道跑完');
    assert.equal(started, 0, '装配期就知道缺工具，不该再起子代理去撞一次（撞了必然被拒）');
    assert.deepEqual(out.toolGap, ['web_search', 'advanced_search', 'platform_search', 'web_fetch']);
    assert.ok(out.jobs.every((job) => job.via === 'builtin'));

    const text = renderDispatch(out);
    assert.ok(text.includes('这个组合里没有联网检索工具'), `反馈要说明白为什么走内置通道：${text.slice(0, 240)}`);
    assert.ok(!text.includes('known global tools'), '不该再把 tools.restrict() 的原始清单贴给模型');
    assert.equal(text.split('这个组合里没有联网检索工具').length - 1, 1, '同一件事只说一次，不是每个渠道行各一遍');
});

await check('工具面齐备时照旧走子代理（探针不能把能用的路掐掉）', async () => {
    resetToolGapMemo();
    const type = contentType('hotspot');
    const briefPath = join(root, 'brief-ok.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    let started = 0;
    const counting = { async start() { started += 1; return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} }; } };
    const full = { schemas: () => [...CHANNEL_TOOLS, 'office_run'].map((name) => ({ name })) };
    const out = await dispatchSearch(
        { briefPath, outputPaths: type.channels.map((_, i) => `full-${i}.md`) },
        { agent: { session: { header: { cwd: root }, id: 'session-full' } }, signal: new AbortController().signal },
        {},
        cordisLikeCtx({ subagents: counting, tools: full, web: fakeSeam() }),
    );
    assert.equal(out.ok, true);
    assert.equal(started, type.channels.length, '工具齐备时每渠道仍该起一个子代理');
    assert.equal(out.toolGap, undefined, '没有缺口就不该报缺口');
    assert.ok(out.jobs.every((job) => job.via === 'subagent'));
    assert.ok(!renderDispatch(out).includes('这个组合里没有联网检索工具'));
});

await check('子代理与内置检索都没有时，仍给出可执行的替代方案', async () => {
    const type = contentType('hotspot');
    const briefPath = join(root, 'brief-none.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `none-${i}.md`);
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    // 没有 subagents、没有 web 接缝，Key 名指向一个不存在的变量 → 两边都不行。
    // 第二十轮起要**显式点名通道**：auto 顺序里还有免 Key 的 duckduckgo / searxng，
    // 「一条都不行」只可能发生在指名选了一条没配好的通道时。
    const barren = { search: { provider: 'anthropic', builtin: { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY' } } };
    await assert.rejects(() => dispatchSearch({ briefPath, outputPaths: outputs }, exec, barren, cordisLikeCtx({})), /没有 subagents 服务/);
    await assert.rejects(() => dispatchSearch({ briefPath, outputPaths: outputs }, exec, barren, undefined), /没有 subagents 服务/);
    // 失败说明必须把降级路径写清楚：办公模式里 Agent Teams 是默认开的，
    // 「换个说法再试一次」不如「改用 spawn_teammate 逐渠道派工」有用。
    assert.ok(NO_SUBAGENTS_HINT.includes('spawn_teammate'), '降级提示要点名 Agent Teams 的入口');
    assert.ok(NO_SUBAGENTS_HINT.includes('office_parse_findings'), '降级后仍要回到解析这一步');
});

await check('engine=subagent 但组合里没有 subagents 服务时，明确报出替代方案', async () => {
    const type = contentType('hotspot');
    const briefPath = join(root, 'brief-subonly.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `subonly-${i}.md`);
    await assert.rejects(
        () => dispatchSearch(
            { briefPath, outputPaths: outputs },
            { agent: {}, signal: new AbortController().signal },
            { search: { engine: 'subagent' } },
            cordisLikeCtx({ web: fakeSeam() }),
        ),
        /没有 subagents 服务/,
    );
});

await check('cordis 形态的 ctx：服务只能经 ctx.get 取，直接读属性会当场抛错', async () => {
    // 这一条钉的就是 2026-09-25 真实会话里的失败：
    //   Error: cannot get property "subagents" without inject
    // 旧代码直接读 ctx.subagents，于是第二步从上线起就没跑通过。
    const service = { async start() { return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} }; } };
    const ctx = cordisLikeCtx({ subagents: service });
    assert.throws(() => ctx.subagents, /without inject/, '替身要如实模拟宿主：未 inject 的属性读取必须抛错');
    assert.equal(subagentsOf(ctx), service, 'subagentsOf 必须经 ctx.get 拿到服务');
    assert.equal(subagentsOf(undefined), undefined);
    assert.equal(subagentsOf({ subagents: service }), undefined, '没有 ctx.get 的替身不再被当成有服务（那条路在宿主上必然抛错）');
    assert.equal(subagentsOf(cordisLikeCtx({ subagents: { notStart: true } })), undefined, '拿到的服务不能 start 时按「没有服务」处理');

    // 端到端：cordis 形态的 ctx 必须能把渠道派出去。
    const type = contentType('hotspot');
    const briefPath = join(root, 'brief-inject.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：热点事件（hotspot）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `inject-out-${i}.md`);
    let started = 0;
    const counting = { async start() { started += 1; return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} }; } };
    const out = await dispatchSearch(
        { briefPath, outputPaths: outputs },
        { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal },
        {},
        cordisLikeCtx({ subagents: counting }),
    );
    assert.equal(out.ok, true);
    assert.equal(started, type.channels.length, '每个渠道都要真的起一个子代理');
});

await check('派工能把渠道派给注入的假 subagents，并收集结果', async () => {
    const type = contentType('mixed');
    const briefPath = join(root, 'brief4.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：混合或未指明（mixed）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `mixed-out-${i}.md`);
    const started = [];
    const fakeCtx = cordisLikeCtx({
        subagents: {
            async start(provider, request) {
                started.push({ provider, label: request.label, text: request.prompt[0].text });
                return {
                    result: Promise.resolve({ stopReason: 'completed' }),
                    dispose: async () => {},
                };
            },
        },
    });
    const parent = { session: { header: { cwd: root } } };
    const out = await dispatchSearch(
        { briefPath, outputPaths: outputs },
        { agent: parent, signal: new AbortController().signal },
        {},
        fakeCtx,
    );
    assert.equal(out.ok, true);
    assert.equal(out.jobs.length, type.channels.length, '每个渠道一个子代理');
    assert.equal(started.length, type.channels.length);
    assert.ok(started.every((job) => job.provider === 'spawn'), '用 spawn 提供方');
    assert.ok(started.every((job) => job.text.includes('写成文件')), '每份任务书都要求写文件');
});

await check('派工把子代理的工具裁成只留检索与写盘', async () => {
    // 子代理默认继承父方组装（办公模式含 office_run 等），必须用白名单裁掉。
    const type = contentType('mixed');
    const briefPath = join(root, 'brief5.md');
    writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某事件\n内容类型：混合或未指明（mixed）\n', 'utf8');
    const outputs = type.channels.map((_, i) => `filter-out-${i}.md`);
    let seenFilter = null;
    const fakeCtx = cordisLikeCtx({
        subagents: {
            async start(provider, request) {
                seenFilter = request.toolFilter;
                return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} };
            },
        },
    });
    const parent = { session: { header: { cwd: root } } };
    await dispatchSearch(
        { briefPath, outputPaths: outputs },
        { agent: parent, signal: new AbortController().signal },
        {},
        fakeCtx,
    );
    assert.ok(seenFilter !== null && Array.isArray(seenFilter.allow), '必须传 toolFilter.allow 白名单');
    for (const wanted of ['web_search', 'advanced_search', 'platform_search', 'web_fetch', 'read', 'read_image', 'write']) {
        assert.ok(seenFilter.allow.includes(wanted), `白名单缺 ${wanted}`);
    }
    for (const unwanted of ['office_run', 'office_help', 'bash', 'pwsh', 'spawn_teammate', 'team_task_create', 'present', 'todo_write', 'skill']) {
        assert.ok(!seenFilter.allow.includes(unwanted), `白名单不该放行 ${unwanted}`);
    }
    assert.ok(!('deny' in seenFilter), '用白名单而不是黑名单：新工具不会悄悄漏给子代理');
});

await check('子代理工具面是极简的：只留检索、取正文、读、写', () => {
    // 这是一条硬约束：子代理挂的工具一多就会跑偏。数量与内容都钉住，
    // 以后谁往里加工具都会先看到这条测试失败。
    assert.deepEqual([...CHANNEL_TOOLS], [
        'web_search',
        'advanced_search',
        'platform_search',
        'web_fetch',
        'read',
        'read_image',
        'write',
    ]);
    for (const unwanted of ['edit', 'glob', 'grep', 'bash', 'pwsh', 'office_run', 'office_help', 'spawn_teammate', 'team_task_create', 'team_task_list', 'wait_agent', 'present', 'todo_write', 'goal', 'skill', 'subagent', 'run_code']) {
        assert.ok(!CHANNEL_TOOLS.includes(unwanted), `极简面不该有 ${unwanted}`);
    }
});

await check('任务书告知子代理只有哪些工具，并给了图片读法', () => {
    const type = contentType('hotspot');
    const prompt = buildChannelPrompt(type.channels[0], '某事件', '.office/search/x/01.md', type.name);
    assert.ok(prompt.includes('read_image'), '要告诉子代理图片用 read_image 看');
    assert.ok(prompt.includes('唯一的交付方式'), '要说明 write 是唯一交付方式');
    assert.ok(prompt.includes('没有命令执行'), '要说明没有命令执行');
    assert.ok(prompt.includes('据图片'), '图片里读出的事实要标来源');
});

await check('派工结果反馈里明确提示下一步解析', () => {
    const text = renderDispatch({ topic: '某事件', jobs: [{ channel: '泛搜', path: 'a.md', ok: true }] });
    assert.ok(text.includes('office_parse_findings'), '要指向下一步');
    assert.ok(text.includes('不要把结果原文搬进对话'), '要重申不搬运原文');
});

await check('多文件解析把各渠道结论合并，并整体重算覆盖度', async () => {
    const a = join(root, 'chan-a.md');
    const b = join(root, 'chan-b.md');
    writeFileSync(a, ['## 泛搜', '- 说法一 (https://one.example.com/1)'].join('\n'), 'utf8');
    writeFileSync(b, ['## 权威媒体', '- 说法一 (https://two.example.com/2)'].join('\n'), 'utf8');
    const out = await parseResultFiles({ paths: [a, b], type: 'hotspot', topic: '某事件' }, { agent: { session: { header: { cwd: root } } } });
    assert.equal(out.ok, true);
    assert.equal(out.files.length, 2);
    assert.equal(out.result.findings.length, 1, '同一说法合并为一条');
    assert.equal(out.result.findings[0].urls.length, 2, '两个渠道的出处都保留');
    assert.equal(out.result.findings[0].singleSource, false, '跨文件也算跨源核对');
});

await check('多文件解析对读不到的文件列入 missing 而不是整单失败', async () => {
    const good = join(root, 'good.md');
    writeFileSync(good, '## 泛搜\n- 有内容 (https://ok.example.com/1)\n', 'utf8');
    const out = await parseResultFiles({ paths: [good, join(root, 'nope.md')] }, { agent: { session: { header: { cwd: root } } } });
    assert.equal(out.files.length, 1);
    assert.equal(out.missing.length, 1);
});

await check('全部文件都读不到时报错，并提示确认子代理已写盘', async () => {
    await assert.rejects(
        () => parseResultFiles({ paths: [join(root, 'absent.md')] }, { agent: { session: { header: { cwd: root } } } }),
        /都读不到/,
    );
});

await check('工具面里有检索三件套，且派工是排他的', () => {
    const tools = buildTools(resolveConfig({}));
    const names = tools.map((t) => t.name);
    for (const wanted of ['office_search_run', 'office_search_brief', 'office_search_dispatch', 'office_parse_findings']) {
        assert.ok(names.includes(wanted), `工具面缺 ${wanted}`);
    }
    const dispatch = tools.find((t) => t.name === 'office_search_dispatch');
    assert.equal(dispatch.isConcurrencySafe(), false, '派工要写文件，不能并发跑');
});

// ── 内置检索直查（office_search_run） ──────────────────────────────────────

await check('office_search_run 直查：走内置通道、落盘、只回紧凑清单', async () => {
    const tools = buildTools(resolveConfig({}), cordisLikeCtx({ web: fakeSeam() }));
    const tool = tools.find((t) => t.name === 'office_search_run');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    const value = await tool.execute({ queries: ['某事件 通报', '某事件 争议'] }, exec);

    assert.equal(value.ok, true);
    assert.equal(value.sources, 4, '两条查询各两条来源');
    assert.ok(existsSync(join(root, value.file)), '结果文件要落盘');
    const text = readFileSync(join(root, value.file), 'utf8');
    assert.ok(text.includes('## 某事件 通报'), '按查询分组');
    assert.ok(text.includes('https://news-a.example.com/'), '来源 URL 要写进文件');
    assert.ok(text.includes('不可信数据'), '要写明外部内容是未核实材料');
    assert.ok(value.text.includes('内置检索'), '反馈要说清楚是内置检索拿的');
    assert.ok(value.text.includes('不是结论'), '清单不是结论，要提醒核对');
    // 聊天里回的是清单：每条来源一行标题 + URL，不搬运摘录原文。
    assert.ok(value.text.includes('https://news-a.example.com/'));

    // out 参数：调用方指定路径时按它落盘（相对会话工作目录解析）。
    const custom = await tool.execute({ queries: ['某事件 通报'], out: '我的检索/结果.md' }, exec);
    assert.equal(custom.file, '我的检索/结果.md');
    assert.ok(existsSync(join(root, '我的检索', '结果.md')), 'out 指定的路径要真的写出来');
});

await check('office_search_run 的落盘文件能被 office_parse_findings 读（格式接得上）', async () => {
    const tools = buildTools(resolveConfig({}), cordisLikeCtx({ web: fakeSeam() }));
    const run = tools.find((t) => t.name === 'office_search_run');
    const parse = tools.find((t) => t.name === 'office_parse_findings');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    const value = await run.execute({ queries: ['某事件 通报'] }, exec);
    const digest = await parse.execute({ paths: [value.file], topic: '某事件' }, exec);
    assert.equal(digest.ok, true);
    assert.ok(digest.text.includes('到手结论'), '解析要真的抽出结论');
});

await check('office_search_run 在通道不可用时说清原因，而不是静默返回空', async () => {
    // 指名一条没配好的通道：这时才真的「一条都不行」（auto 下还有免 Key 的通道可试）。
    const config = resolveConfig({ search: { provider: 'anthropic', builtin: { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY' } } });
    const tools = buildTools(config, cordisLikeCtx({}));
    const tool = tools.find((t) => t.name === 'office_search_run');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    await assert.rejects(() => tool.execute({ queries: ['某事件'] }, exec), /内置检索用不了/);
    await assert.rejects(() => tool.execute({ queries: [] }, exec), /queries 不能为空/);
});

await check('office_search_run 能按次点名通道，并在反馈里写出用的是哪条', async () => {
    // 免 Key 的兜底通道：auto 顺序里最后一条，但**明文点名**时只走它。
    const seam = fakeSeam();
    const config = resolveConfig({});
    const tools = buildTools(config, cordisLikeCtx({ web: seam }));
    const tool = tools.find((t) => t.name === 'office_search_run');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    const value = await tool.execute({ queries: ['某事件'], provider: 'seam', fetchPages: 0 }, exec);
    assert.equal(value.ok, true, value.text);
    assert.ok(value.text.includes('宿主 web 服务'), `反馈要写清走的哪条通道：${value.text}`);
    const text = readFileSync(join(root, value.file), 'utf8');
    assert.ok(text.includes('通道：宿主 web 服务'), '结果文件里也要留一笔');
    // 不认识的通道名要当场拒绝，并列出可选值（不能让模型以为「查不到」）。
    await assert.rejects(() => tool.execute({ queries: ['某事件'], provider: 'gugou' }, exec), /不认识的检索通道/);
});

// ── 内容级哨兵与失败分类接到反馈上（第十八轮 P0-8 / P0-9）──────────────────

await check('内置通道把取正文失败按分类写进渠道文件（原先被静默吞掉）', async () => {
    const channel = contentType('hotspot').channels[0];
    // 检索成功、取正文被目标站拒绝：403 这一类必须能被分开数出来。
    const access = createWebAccess(cordisLikeCtx({}), { apiKey: 'k' }, {
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async (url, init) => {
            if (init?.method === 'POST') {
                return {
                    ok: true,
                    status: 200,
                    headers: { get: (name) => (name === 'content-type' ? 'application/json' : null) },
                    async json() {
                        return {
                            content: [{
                                type: 'web_search_tool_result',
                                content: [{ type: 'web_search_result', url: 'https://blocked.example.com/a', title: '被挡的来源' }],
                            }],
                        };
                    },
                };
            }
            return { status: 403, headers: { get: (name) => (name === 'content-type' ? 'text/html' : null) }, body: { cancel: async () => {} } };
        },
    });
    const outputPath = join(root, 'fetch-fault.md');
    const out = await runChannelBuiltin(channel, '某事件', { outputPath }, access, {}, new AbortController().signal);
    const text = readFileSync(outputPath, 'utf8');
    assert.equal(out.sources, 1, '来源本身还是拿到了，不该整条渠道失败');
    assert.ok(text.includes('这些来源没取到正文'), `取正文失败必须写进渠道文件：${text.slice(-300)}`);
    assert.ok(text.includes('目标站拒绝 1'), `要按分类计数，而不是一句「打不开」：${text.slice(-300)}`);
    assert.ok(text.includes('别重试同一个地址'), '要给可执行的下一步');
});

await check('office_search_run 把取正文失败摆到反馈里，不再静默跳过', async () => {
    // 接缝能搜到来源、但取正文失败；来源用 IP 字面量，好让自带兜底在
    // SSRF 防线处就停住 —— 默认测试套件不许真的联网（连 DNS 都不该发）。
    const seam = {
        async search() {
            return { sources: [{ url: 'http://127.0.0.1/a', title: '取不到的来源' }], truncated: false };
        },
        async fetch() {
            throw new Error('接缝拒绝取这个地址');
        },
    };
    const tools = buildTools(resolveConfig({}), cordisLikeCtx({ web: seam }));
    const tool = tools.find((t) => t.name === 'office_search_run');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    const value = await tool.execute({ queries: ['某事件'] }, exec);

    assert.equal(value.ok, true, '一条来源取不到正文不该让整轮直查失败');
    assert.ok(value.text.includes('取正文失败 1 条'), `反馈要说清取正文失败了几条：${value.text}`);
    const text = readFileSync(join(root, value.file), 'utf8');
    assert.ok(text.includes('这些来源没取到正文'), '结果文件里也要留一笔，供第三步解析时看到');
});

await check('office_search_run 能被设置页关掉', () => {
    const tools = buildTools(resolveConfig({ tools: { office_search_run: false } }));
    assert.ok(!tools.map((t) => t.name).includes('office_search_run'));
    assert.ok(tools.map((t) => t.name).includes('office_search_brief'), '关一个不该连带关别的');
});

await check('office_help 的 search 话题讲清了三步与分流规则', async () => {
    const help = await buildHelp('search');
    assert.equal(help.topic, 'search');
    for (const wanted of ['office_search_brief', 'office_search_dispatch', 'office_parse_findings', '茧房', '泛搜']) {
        assert.ok(help.text.includes(wanted), `search 话题缺 ${wanted}`);
    }
    const index = await buildHelp('');
    assert.ok(index.text.includes('search：'), '索引里要能查到 search');
});

// ── 收尾 ──────────────────────────────────────────────────────────────────

rmSync(root, { recursive: true, force: true });

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`search: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);