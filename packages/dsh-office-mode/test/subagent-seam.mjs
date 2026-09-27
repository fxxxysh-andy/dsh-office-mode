/**
 * 宿主接缝测试：用**真实的 cordis** 复现并钉住「派工取不到 subagents 服务」这个 bug。
 *
 * 为什么单独一个文件：test/search.mjs 是纯逻辑测试（不依赖本机装了什么），而这一条
 * 要真起一个 cordis 组合，验证的是「插件上下文怎么解析服务」这件宿主行为。分开之后，
 * 换台机器跑不到宿主的 cordis 时这里打印「跳过」而不是把纯逻辑测试一起拖红。
 *
 * 被测事实（2026-09-25 的真实 cordis 4.0.4 实测）：
 *
 *   服务由**兄弟节点**提供（dsh-base 的 subagent 行与 dsh-office-mode 行同为根的两
 *   个子节点），而插件 ctx 的 inject 只有 ['tools']。此时：
 *     ctx.subagents      → 抛 Error: cannot get property "subagents" without inject
 *     ctx.get('subagents') → 拿到服务（cordis 文档：Read a service from the store
 *                            without the inject requirement）
 *   没提供的服务经 ctx.get 取到的是 undefined，不抛错。
 *
 * 这正是 session5 里 office_search_dispatch 两次报「cannot get property "subagents"
 * without inject」、检索三步走第二步整条失效的根因；也说明为什么修法只能是
 * ctx.get，而不是把 'subagents' 加进模块级 inject（那会让缺这个服务的组合整个
 * 插件起不来）。
 *
 * 跑法：node test/subagent-seam.mjs（宿主 cordis 找不到时会打印「跳过」）
 */
import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';

import { contentType } from '../src/search-routes.js';
import { dispatchSearch, NO_SUBAGENTS_HINT, subagentsOf } from '../src/search.js';

/** 解析宿主的 cordis：锚点与 src/index.js 的 resolveHostModule 同一套。 */
function resolveCordis() {
    const anchors = [
        process.env.DSH_PROFILE_DIR,
        process.env.DSH_HOST_ROOT,
        process.env.DSH_CHECKOUT,
        typeof process.argv[1] === 'string' && process.argv[1] !== '' ? dirname(process.argv[1]) : undefined,
    ].filter((value) => typeof value === 'string' && value !== '');
    for (const anchor of anchors) {
        try {
            const req = createRequire(pathToFileURL(join(anchor, 'package.json')).href);
            return req.resolve('@deepseek-ai/cordis');
        } catch {
            // 换下一个锚点
        }
    }
    // 兜底：本机 DSH 的安装位置（node 全局 node_modules）。
    const fallback = join(dirname(process.execPath), 'node_modules', '@deepseek-ai', 'dsh', 'node_modules', '@deepseek-ai', 'cordis', 'lib', 'index.js');
    return existsSync(fallback) ? fallback : undefined;
}

const cordisPath = resolveCordis();
if (cordisPath === undefined) {
    console.log('subagent-seam: 跳过（本机找不到宿主的 @deepseek-ai/cordis，这条只验证宿主行为，不影响纯逻辑测试）');
    process.exit(0);
}

const { Context } = await import(pathToFileURL(cordisPath).href);

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const root = mkdtempSync(join(tmpdir(), 'office-seam-'));
const HOTSPOT = contentType('hotspot');
const briefPath = join(root, 'brief.md');
writeFileSync(briefPath, '# 检索提纲\n\n检索提纲：某地化工厂爆炸\n内容类型：热点事件（hotspot）\n', 'utf8');
const outputs = HOTSPOT.channels.map((_, index) => `seam-${index}.md`);
const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };

const starts = [];
const fakeSubagents = {
    async start(provider, request) {
        starts.push({ provider, request });
        return { result: Promise.resolve({ stopReason: 'completed' }), dispose: async () => {} };
    },
};

// 宿主组合：服务由兄弟节点提供（与 dsh-base 的 subagent 行同形），办公插件那一行
// 的 inject 只有 tools。
const ctx = new Context();
const providerFiber = ctx.plugin({
    name: 'subagent-provider',
    apply(fiberCtx) {
        fiberCtx.provide('subagents', fakeSubagents);
        fiberCtx.provide('tools', { register: () => () => {} });
    },
});
await providerFiber;

let officeCtx;
const officeFiber = ctx.plugin({
    name: 'office-mode-like',
    inject: ['tools'],
    apply(fiberCtx) {
        officeCtx = fiberCtx;
    },
});
await officeFiber;

await check('兄弟节点提供的服务：直接读属性抛错，ctx.get 拿得到', () => {
    assert.throws(() => officeCtx.subagents, /cannot get property "subagents" without inject/,
        '未写进 inject 的服务属性必须如实抛错——这就是 session5 里那条报错');
    assert.equal(officeCtx.get('subagents'), fakeSubagents, 'ctx.get 是免 inject 的取服务入口');
    assert.equal(officeCtx.get('这个服务不存在'), undefined, '没提供的服务经 ctx.get 取到 undefined，不抛错');
});

await check('subagentsOf 在真实 ctx 上取到服务，在缺服务时安静返回 undefined', async () => {
    assert.equal(subagentsOf(officeCtx), fakeSubagents);
    const bare = new Context();
    const bareFiber = bare.plugin({
        name: 'bare',
        inject: ['tools'],
        apply(fiberCtx) {
            fiberCtx.provide('tools', { register: () => () => {} });
        },
    });
    await bareFiber;
    const consumerFiber = bare.plugin({ name: 'consumer', inject: ['tools'], apply(fiberCtx) { assert.equal(subagentsOf(fiberCtx), undefined); } });
    await consumerFiber;
});

await check('派工在真实 cordis 上下文里真的能把渠道铺出去', async () => {
    starts.length = 0;
    const out = await dispatchSearch({ briefPath, outputPaths: outputs }, exec, {}, officeCtx);
    assert.equal(out.ok, true, '每个渠道都要成功');
    assert.equal(out.jobs.length, HOTSPOT.channels.length);
    assert.equal(starts.length, HOTSPOT.channels.length, '每个渠道一个子代理');
    for (const item of starts) {
        assert.equal(item.provider, 'spawn', '用 spawn 提供方');
        assert.ok(Array.isArray(item.request.toolFilter?.allow), '检索子代理要带工具白名单');
        assert.ok(!('deny' in (item.request.toolFilter ?? {})), '用白名单而不是黑名单');
        assert.ok(item.request.prompt[0].text.includes('写成文件'), '任务书要求把材料写进文件');
        assert.equal(item.request.parent, exec.agent, '子代理要挂在父 agent 上');
    }
});

await check('缺服务的组合给出的是可执行的降级说明，而不是宿主的代理报错', async () => {
    const bare = new Context();
    const fiber = bare.plugin({
        name: 'no-subagents',
        inject: ['tools'],
        apply(fiberCtx) {
            fiberCtx.provide('tools', { register: () => () => {} });
            assert.equal(subagentsOf(fiberCtx), undefined);
        },
    });
    await fiber;
    // 没有 subagents、又没有联网通道时，才该整条失败：第二十轮起要**显式点名**一条
    // 没配好的通道（auto 顺序里还有免 Key 的 duckduckgo / searxng，它们会让派工跑起来）。
    const barren = { search: { provider: 'anthropic', builtin: { apiKeyEnv: 'OFFICE_TEST_MISSING_KEY' } } };
    await assert.rejects(
        () => dispatchSearch({ briefPath, outputPaths: outputs }, exec, barren, undefined),
        (error) => {
            assert.ok(error.message.startsWith(NO_SUBAGENTS_HINT), '先给降级说明');
            assert.ok(error.message.includes('没有 subagents 服务'));
            assert.ok(error.message.includes('spawn_teammate'), '要指向 Agent Teams 这条降级路径');
            assert.ok(error.message.includes('内置检索也用不了'), '要说明内置检索为什么也不行');
            assert.ok(!/without inject/.test(error.message), '不该把宿主的代理报错原样丢给模型');
            return true;
        },
    );
});

for (const item of results) {
    if (item.ok) console.log('PASS ', item.name);
    else console.log('FAIL ', item.name, '\n      ', item.error?.message ?? item.error);
}
const failed = results.filter((item) => !item.ok);
console.log(`subagent-seam: ${results.length - failed.length}/${results.length} 通过`);
if (failed.length > 0) process.exitCode = 1;
