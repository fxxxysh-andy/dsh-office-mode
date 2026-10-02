/**
 * 能力地图的回归测试（第十八轮 P2-4 / 第二十三轮 23-2）。
 *
 * 这一份守的是「persona 不再写死能力清单、实际能力由插件注入」这条结构：
 *
 *   1. **注入通道真的是变量**。办公 preset 的 persona 是 `complete: true` 的唯一提示段，
 *      插件注册的 section 在办公会话里进不了提示词 —— 能到模型的只有 `{{变量}}` 的插值。
 *      所以断言插件确实注册了 `office_capabilities`，且求值出来是一段可用的文本。
 *   2. **地图只讲工具面讲不了的事**（引擎入口、派工路径），不复述工具清单：
 *      工具面本身就是实际清单（关掉的工具不会出现），复述一遍是常驻成本。
 *   3. **探测是同步、有界、可复算的**：渲染是纯函数（同样快照同样输出），
 *      引擎清单每类封顶，整体不超 `CAPABILITY_MAP_BUDGET_BYTES`。
 *   4. **配置改了就重探**：python.bin / 模型目录这类设置能在地图上立刻反映出来。
 *
 * 跑法：node test/capabilities.mjs
 */
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import {
    CAPABILITY_MAP_BUDGET_BYTES,
    CAPABILITY_VARIABLE,
    renderCapabilityMap,
    resetCapabilities,
    syncCapabilities,
} from '../src/capabilities.js';
import { buildTools } from '../src/tools.js';
import { resolveConfig } from '../src/config.js';
import { findPdfExecutable } from '../src/pdf.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const bytes = (text) => Buffer.byteLength(String(text ?? ''), 'utf8');

/** 与 test/memory.mjs 同一套假 ctx：只够 apply() 走完注册那几步。 */
function fakeCtx({ subagents = undefined, withPrompt = true } = {}) {
    const sections = [];
    const variables = [];
    const tools = [];
    const ctx = {
        tools: { register: (definition) => { tools.push(definition.name); return () => {}; } },
        get: (key) => {
            if (key === 'systemPrompt' && withPrompt) {
                return {
                    section: (value) => { sections.push(value); return () => {}; },
                    variable: (name, provider) => { variables.push({ name, provider }); return () => {}; },
                };
            }
            if (key === 'subagents') return subagents;
            return undefined;
        },
        inject: () => {},
        on: () => {},
        logger: {},
    };
    return { ctx, sections, variables, tools };
}

// ── 1. 注入通道：变量确实注册了，且求值出来是可用的文本 ────────────────────

/** 设置变更事件名（宿主 dsh-settings 的广播口径）。 */
const OFFICE_SETTINGS_NS_EVENT = 'settings/document-updated';

await check('插件注册了 {{office_capabilities}} 变量，求值出来的地图可用（23-2）', async () => {
    const { apply } = await import('../src/index.js');
    // 假 ctx 给的是**能用的**子代理服务（有 start）：判据与 search.js 的 subagentsOf 同一口径，
    // 只有 `ctx.get('subagents')` 返回个对象是不够的（复核 F1）。
    const { ctx, variables, tools } = fakeCtx({ subagents: { start: () => {} } });
    apply(ctx, {});
    assert.ok(tools.length >= 9, `工具面至少要注册九个工具，实际 ${tools.length}`);
    const registered = variables.find((item) => item.name === CAPABILITY_VARIABLE);
    assert.ok(registered, `插件必须注册变量 ${CAPABILITY_VARIABLE}（persona 是 complete 段，section 进不来）`);
    assert.equal(variables.length, 1, '本插件只注册这一个变量（多了就是又一处常驻成本）');
    const text = registered.provider();
    assert.ok(typeof text === 'string' && text.trim() !== '', '变量求值不能是空串');
    assert.ok(text.includes('本会话能力'), '地图要有标题，模型才知道这一段是什么');
    assert.ok(text.includes('office_search_dispatch'), '派工那条路要写清（23-2 的错配就发生在这里）');
    assert.ok(text.includes('有 subagents 服务'), '服务真的能用时要报「走子代理」');
    return `${bytes(text)} B`;
});

await check('没有 subagents 服务时，地图照实说派工走内置检索通道', async () => {
    const { apply } = await import('../src/index.js');
    const { ctx, variables } = fakeCtx({ subagents: undefined });
    apply(ctx, {});
    const text = variables[0].provider();
    assert.ok(text.includes('没有 subagents 服务'), '缺服务时要说清走的是哪条路');
    assert.ok(text.includes('内置检索通道'), '要写出替代路径的名字，否则模型会以为派工不可用');
});

await check('没有 systemPrompt 服务时不报错，只是不注册变量', async () => {
    const { apply } = await import('../src/index.js');
    const { ctx, variables, tools } = fakeCtx({ withPrompt: false });
    apply(ctx, {});
    assert.equal(variables.length, 0);
    assert.ok(tools.length >= 9, '其余功能照常');
});

await check('设置改动重建后变量不重复注册（disposer 真的把旧的拆掉了）', async () => {
    const { apply } = await import('../src/index.js');
    // 这一份假 ctx 会**照着 cordis 的口径**查重：同名变量再注册一次就抛
    // （NamedEntries 的重复检测）。旧的没被 disposer 拆掉时，这条会红。
    const live = new Map();
    const handlers = {};
    const ctx = {
        tools: { register: () => () => {} },
        get: (key) => {
            if (key === 'systemPrompt') {
                return {
                    section: () => () => {},
                    variable: (name, provider) => {
                        if (live.has(name)) throw new Error(`prompt variable "${name}" is already registered`);
                        live.set(name, provider);
                        return () => { live.delete(name); };
                    },
                };
            }
            if (key === 'subagents') return undefined;
            return undefined;
        },
        inject: (deps, callback) => {
            if (deps.includes('settings')) callback({ on: (event, handler) => { handlers[event] = handler; return () => {}; } });
        },
        on: () => () => {},
        logger: {},
    };
    apply(ctx, {});
    assert.equal(live.size, 1, '激活时注册一次');
    const first = live.get(CAPABILITY_VARIABLE);
    handlers[OFFICE_SETTINGS_NS_EVENT]('dsh-office-mode');
    assert.equal(live.size, 1, 'rebuild 之后仍然只有一个（旧的必须被拆掉）');
    assert.notEqual(live.get(CAPABILITY_VARIABLE), first, '重建要换上新的一份（设置改了要重探）');
    handlers[OFFICE_SETTINGS_NS_EVENT]('dsh-office-mode');
    assert.equal(live.size, 1);
});

// ── 2. 地图只讲工具面讲不了的事 ────────────────────────────────────────────

await check('地图不复述工具清单（工具面本身就是实际清单）', async () => {
    const snapshot = syncCapabilities({});
    const text = renderCapabilityMap(snapshot, { subagents: true });
    // 唯一的例外是 `office_search_dispatch`：地图要指名「走子代理还是内置通道」，
    // 那不是复述清单，是路由事实（哪条路通）。
    const allowed = new Set(['office_search_dispatch']);
    for (const name of buildTools(resolveConfig({})).map((tool) => tool.name)) {
        if (allowed.has(name)) continue;
        assert.ok(!text.includes(name),
            `地图里出现了工具名 ${name}：清单由工具面承担（关掉的工具不会出现），复述一遍只是常驻成本`);
    }
    assert.ok(text.includes('工具以本会话工具面为准'), '但要点明「以工具面为准」，否则模型不知道清单去哪看');
});

await check('地图不泄漏本机绝对路径（解释器只报文件名）', async () => {
    const text = renderCapabilityMap({
        pdf: { render: [], text: [] },
        tex: [],
        preview: false,
        python: 'C:\\Users\\someone\\Python313\\python.EXE',
        av: { ffmpeg: false, model: false },
    }, { subagents: false });
    assert.ok(!text.includes('C:\\') && !text.includes('Users'), '路径不该进提示词');
    assert.ok(text.includes('python.EXE'), '但要报出解释器的文件名');
});

// ── 3. 纯函数、有界、可复算 ────────────────────────────────────────────────

await check('渲染是纯函数：同样快照两次渲染逐字相同，省略 options 也不炸', async () => {
    const snapshot = { pdf: { render: ['pdftoppm'], text: ['pdftotext'] }, tex: ['xelatex'], preview: true, python: 'C:\\x\\python.exe', av: { enabled: true, available: true } };
    const a = renderCapabilityMap(snapshot, { subagents: true });
    const b = renderCapabilityMap(snapshot, { subagents: true });
    assert.equal(a, b);
    // 省略 options / 缺字段 / 类型不对，都不该抛（这一份文本进提示词管线，抛出去就是整轮失败）。
    const noOptions = renderCapabilityMap(snapshot);
    assert.ok(noOptions.includes('检索派工'), '省略 options 时按「没有 subagents」处理');
    const oddShapes = [
        { pdf: { render: 'not-an-array', text: 42 }, tex: null, preview: 'yes', python: 7, av: null },
        { pdf: {}, av: {} },
        { av: { enabled: true } },
    ];
    for (const odd of oddShapes) {
        const text = renderCapabilityMap(odd, { subagents: false });
        assert.ok(typeof text === 'string' && text !== '', `形状不对也要出文本：${JSON.stringify(odd)}`);
    }
});

await check('引擎清单每类封顶三个，多的收成「等」', async () => {
    const text = renderCapabilityMap({
        pdf: { render: ['pdftoppm', 'pdftocairo', 'mutool', 'gs'], text: [] },
        tex: [],
        preview: false,
        python: null,
        av: { ffmpeg: false, model: false },
    }, { subagents: false });
    assert.ok(text.includes('pdftoppm / pdftocairo / mutool 等'), '超过三个要收成「等」');
    assert.ok(!text.includes('gs'), '被收掉的那个不该还留着');
});

await check('最坏情况的地图不超预算（地图跟着 persona 进每一次请求）', async () => {
    // 与 preset-check 的 WORST_CAPABILITY_MAP 同一份形状：引擎满 + 音视频「部分缺件」
    // （那句最长）+ 解释器文件名顶到 16 字符上限 + 派工走「没有 subagents」那一支。
    const worst = renderCapabilityMap({
        pdf: { render: ['pdftoppm', 'pdftocairo', 'mutool', 'gs'], text: ['pdftotext'] },
        tex: ['latexmk', 'xelatex', 'lualatex'],
        preview: true,
        python: 'C:/somewhere/aaaaaaaaaaaaaaaaaaaaaaaaaaaa.exe',
        av: { enabled: true, available: false, ffmpeg: true, ffprobe: true, model: true, vad: false, runtime: true },
    }, { subagents: false });
    assert.ok(bytes(worst) <= CAPABILITY_MAP_BUDGET_BYTES,
        `最坏地图 ${bytes(worst)}B 超过预算 ${CAPABILITY_MAP_BUDGET_BYTES}B：要加内容先减内容`);
    return `${bytes(worst)} / ${CAPABILITY_MAP_BUDGET_BYTES} B`;
});

await check('探测失败或没探过时不装作有能力：照实说没探到', async () => {
    for (const empty of [undefined, null, {}, []]) {
        const text = renderCapabilityMap(empty, { subagents: true });
        assert.ok(text.includes('没探到'), `${JSON.stringify(empty)} 要照实说没探到（「没探过」与「探到没有」是两件事）`);
        assert.ok(!text.includes('引擎：'), '不能列出任何一条引擎');
    }
    // 真的探过（有 pdf 字段）才逐项报缺什么。
    const probed = renderCapabilityMap({ pdf: { render: [], text: [] }, tex: [], preview: false, python: null, av: {} }, { subagents: false });
    assert.ok(probed.includes('PDF 渲染 无'), '探过之后才能报「无」');
});

await check('没有 CLI 引擎时不假装知道 PyMuPDF 在不在', async () => {
    const text = renderCapabilityMap({
        pdf: { render: [], text: [] },
        tex: [],
        preview: false,
        python: 'python.exe',
        av: {},
    }, { subagents: false });
    assert.ok(text.includes('未探到 CLI 引擎'), '同步探测看不到 PyMuPDF，照实说没探到 CLI 引擎');
    assert.ok(text.includes('office.pdf.engines'), '并指到真能回答它的地方');
    assert.ok(!text.includes('需 PyMuPDF'), '不能断言「需要装 PyMuPDF」——本机可能已经装了');
});

await check('音视频与 av 通道同口径：五样齐才算可用，关掉了照实说', async () => {
    const base = { pdf: { render: [], text: [] }, tex: [], preview: false, python: null };
    const all = renderCapabilityMap({ ...base, av: { enabled: true, available: true, ffmpeg: true, ffprobe: true, model: true, vad: true, runtime: true } }, { subagents: false });
    assert.ok(all.includes('音视频 ffmpeg + SenseVoice 模型'), '五样齐才报可用');
    // ffprobe 缺：probeAv 的 missing 里有它，所以地图不能报「可用」（复核 D2-a 的方向）。
    const noProbe = renderCapabilityMap({ ...base, av: { enabled: true, available: false, ffmpeg: true, ffprobe: false, model: true, vad: true, runtime: true } }, { subagents: false });
    assert.ok(noProbe.includes('音视频 部分缺件（用 office.av.check 看缺什么）'), 'ffprobe 缺也是缺件');
    const disabled = renderCapabilityMap({ ...base, av: { enabled: false, available: false } }, { subagents: false });
    assert.ok(disabled.includes('音视频 已关闭'), '关掉了就说关掉了，不假装有能力');
    const none = renderCapabilityMap({ ...base, av: { enabled: true, available: false } }, { subagents: false });
    assert.ok(!none.includes('音视频'), '一样都没有时不必占字节');
});

// ── 4. 探测随配置走 ────────────────────────────────────────────────────────

await check('python.bin 指向不存在的路径时，地图不报 Python 可用；存在时才报', async () => {
    const root = mkdtempSync(join(tmpdir(), 'office-capabilities-'));
    try {
        resetCapabilities();
        const missing = syncCapabilities({ config: resolveConfig({ python: { bin: join(root, 'nope.exe') } }), refresh: true });
        assert.equal(missing.python, null, '填错路径不该退回自动探测（与 python.js 的「填了就以它为准」同一口径）');
        assert.ok(!renderCapabilityMap(missing, { subagents: false }).includes('Python '), '地图不该报出不存在的解释器');

        const fake = join(root, 'python.exe');
        writeFileSync(fake, '');
        resetCapabilities();
        const found = syncCapabilities({ config: resolveConfig({ python: { bin: fake } }), refresh: true });
        assert.equal(found.python, fake);
        assert.ok(renderCapabilityMap(found, { subagents: false }).includes('Python python.exe'), '存在时要报出来');
    } finally {
        resetCapabilities();
        rmSync(root, { recursive: true, force: true });
    }
});

await check('python.bin 填命令名时按 PATH 找（填错路径才判「没有」）', async () => {
    // 设置里 `python.bin` 允许填命令名（`python`）：只按 existsSync 判会把它误报成「没有」。
    // 本机没有 python 时这条没有可判的事实，跳过并说明（不把它算成通过）。
    const found = findPdfExecutable(['python']);
    if (found === undefined) {
        resetCapabilities();
        syncCapabilities({ config: resolveConfig({ python: { bin: 'python' } }), refresh: true });
        return '本机 PATH 上没有 python，跳过（不是通过）';
    }
    resetCapabilities();
    const snapshot = syncCapabilities({ config: resolveConfig({ python: { bin: 'python' } }), refresh: true });
    assert.equal(snapshot.python, found, '命令名要按 PATH 解析出真实路径');
    assert.ok(renderCapabilityMap(snapshot, { subagents: false }).includes('Python '), '解析到了就要报出来');
    resetCapabilities();
});

await check('SenseVoice 模型按「文件真的在」判：目录在而权重缺失不算有', async () => {
    const root = mkdtempSync(join(tmpdir(), 'office-capabilities-'));
    try {
        const modelDir = join(root, 'sensevoice-onnx');
        mkdirSync(modelDir, { recursive: true });
        resetCapabilities();
        const empty = syncCapabilities({ config: resolveConfig({ av: { modelDir } }), refresh: true });
        assert.equal(empty.av.model, false, '只有目录、没有权重时不该报「模型在」');

        writeFileSync(join(modelDir, 'tokens.txt'), '');
        writeFileSync(join(modelDir, 'model.int8.onnx'), '');
        resetCapabilities();
        const full = syncCapabilities({ config: resolveConfig({ av: { modelDir } }), refresh: true });
        assert.equal(full.av.model, true, '权重与词表齐了才算有');
    } finally {
        resetCapabilities();
        rmSync(root, { recursive: true, force: true });
    }
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok && item.note !== undefined ? `  → ${item.note}` : ''}`);
    if (!item.ok) console.log(`       | ${item.error?.message ?? item.error}`);
}
console.log(`capabilities: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);