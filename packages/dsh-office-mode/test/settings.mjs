/**
 * 设置界面的单元测试：schema 形状、默认值、越界拒绝、工具开关、参数生效。
 *
 * schema 需要一个 schemastery 实例，而本插件不依赖它（由宿主提供）。
 * 这里从宿主解析位置取一份；取不到就跳过这一组测试并明确说明。
 *
 * 跑法：node test/settings.mjs
 */
import { existsSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';

import { resolveHostPath } from './host-modules.mjs';

import {
    DEFAULT_MEMORY,
    DEFAULT_SEARCH,
    DEFAULT_SUBAGENT_MODEL,
    DEFAULT_TOOLS,
    DEFAULT_CACHE_TTL_MINUTES,
    DEFAULT_PDF_DPI,
    DEFAULT_PDF_MAX_PAGES,
    resolveConfig,
} from '../src/config.js';
import { DEFAULT_QUOTA } from '../src/memory.js';
import { buildTools } from '../src/tools.js';
import { createOfficeSettings, OFFICE_SETTINGS_NS, SUBAGENT_TOOL_CATALOG, LIMITS } from '../src/settings.js';
import { CHANNEL_TOOLS } from '../src/search.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 从宿主的安装位置解析 schemastery（见 host-modules.mjs）。 */
async function loadZ() {
    const resolved = resolveHostPath('@deepseek-ai/schemastery');
    if (resolved === undefined) return undefined;
    try {
        return (await import(pathToFileURL(resolved).href)).default;
    } catch {
        return undefined;
    }
}

/** 从宿主的安装位置解析 cosmokit：拆 volatile 引用要靠它的 isVolatile。 */
async function loadCosmokit() {
    const resolved = resolveHostPath('@deepseek-ai/cosmokit');
    if (resolved === undefined) return undefined;
    try {
        return await import(pathToFileURL(resolved).href);
    } catch {
        return undefined;
    }
}

const cosmokit = await loadCosmokit();

/**
 * 把解析结果里的 volatile 引用拆成纯值。
 *
 * 0.1.7 起可编辑字段都是 volatile（设置页只接受 volatile 路径的写入），
 * 所以 schema 的解析结果是 Volatile 引用而不是裸值。宿主往页面发值时也是先
 * 拆开的（dsh-settings 的 plainConfig 就是这么做的），这里照同一口径还原，
 * 否则下面的类型与默认值断言会全部看到 [object Object]。
 */
function plain(value) {
    if (cosmokit !== undefined && cosmokit.isVolatile(value)) return plain(value.get());
    if (Array.isArray(value)) return value.map(plain);
    if (value !== null && typeof value === 'object') {
        const out = {};
        for (const [key, child] of Object.entries(value)) out[key] = plain(child);
        return out;
    }
    return value;
}

const z = await loadZ();
if (z === undefined) {
    console.log('跳过：本机解析不到 @deepseek-ai/schemastery，无法测 schema 形状。');
    console.log('（工具开关与参数解析的测试仍然执行，它们不依赖 schemastery。）');
}

// ── 工具开关（不依赖 schemastery） ────────────────────────────────────────

await check('默认全开：空配置下七个工具都在', () => {
    const tools = buildTools(resolveConfig({}));
    assert.deepEqual(tools.map((t) => t.name), [
        'office_help', 'office_run', 'office_memory', 'office_search_run', 'office_search_brief', 'office_search_dispatch', 'office_parse_findings',
    ]);
});

await check('关掉某个工具后它不再注册', () => {
    const tools = buildTools(resolveConfig({ tools: { office_run: false } }));
    assert.ok(!tools.some((t) => t.name === 'office_run'), 'office_run 应被关掉');
    assert.ok(tools.some((t) => t.name === 'office_help'), '其它工具不受影响');
});

await check('全部关掉时工具面为空，但不报错', () => {
    const off = {};
    for (const name of Object.keys(DEFAULT_TOOLS)) off[name] = false;
    const tools = buildTools(resolveConfig({ tools: off }));
    assert.equal(tools.length, 0);
});

await check('未知的工具键被忽略，不会让配置失效', () => {
    const tools = buildTools(resolveConfig({ tools: { 未来才有的工具: false, office_run: false } }));
    assert.ok(!tools.some((t) => t.name === 'office_run'));
    assert.equal(tools.length, 6);
});

await check('非布尔值不会把开关误判成关闭', () => {
    // 设置页写回的是布尔；组合层若写了字符串，按「缺省」处理而不是当成 false。
    const tools = buildTools(resolveConfig({ tools: { office_run: 'yes' } }));
    assert.ok(tools.some((t) => t.name === 'office_run'), '非法值应退回默认（开）');
});

// ── 参数解析与区间收敛 ────────────────────────────────────────────────────

await check('检索参数可调，且在安全区间内收敛', () => {
    const cfg = resolveConfig({ search: { maxParallel: 6, resultLimit: 12, outputDir: '.my-search' } });
    assert.equal(cfg.search.maxParallel, 6);
    assert.equal(cfg.search.resultLimit, 12);
    assert.equal(cfg.search.outputDir, '.my-search');
    assert.equal(resolveConfig({ search: { maxParallel: 999 } }).search.maxParallel, 8, '上界收敛');
    assert.equal(resolveConfig({ search: { maxParallel: 0 } }).search.maxParallel, 1, '下界收敛');
});

await check('未设的检索参数退回默认', () => {
    const cfg = resolveConfig({});
    for (const [key, value] of Object.entries(DEFAULT_SEARCH)) {
        assert.deepEqual(cfg.search[key], value, `${key} 应等于默认值`);
    }
});

await check('空 outputDir 退回默认目录', () => {
    assert.equal(resolveConfig({ search: { outputDir: '   ' } }).search.outputDir, DEFAULT_SEARCH.outputDir);
});

await check('keepCache 可调且默认保留（跨调用复用）', () => {
    assert.equal(resolveConfig({}).keepCache, true);
    assert.equal(resolveConfig({ keepCache: false }).keepCache, false);
    assert.equal(resolveConfig({ keepCache: true }).keepCache, true);
});

await check('PDF 与缓存的新参数有默认值且能收敛', () => {
    const cfg = resolveConfig({});
    assert.equal(cfg.cacheTtlMinutes, DEFAULT_CACHE_TTL_MINUTES);
    assert.equal(cfg.pdfDpi, DEFAULT_PDF_DPI);
    assert.equal(cfg.pdfMaxPages, DEFAULT_PDF_MAX_PAGES);
    assert.equal(cfg.pdfEngine, 'auto');
    assert.equal(resolveConfig({ pdfDpi: 9999 }).pdfDpi, 400, '越界要收敛到上界');
    assert.equal(resolveConfig({ pdfMaxPages: 0 }).pdfMaxPages, 1);
    assert.equal(resolveConfig({ cacheTtlMinutes: 0 }).cacheTtlMinutes, 0, '0 = 不按时间清理');
    assert.equal(resolveConfig({ pdfEngine: '不存在' }).pdfEngine, 'auto', '不认识的引擎回落 auto');
    assert.equal(resolveConfig({ pdfEngine: 'pdftoppm' }).pdfEngine, 'pdftoppm');
});

await check('记忆参数可调，且在安全区间内收敛', () => {
    const cfg = resolveConfig({});
    for (const [key, value] of Object.entries(DEFAULT_MEMORY)) {
        assert.deepEqual(cfg.memory[key], value, `memory.${key} 应等于默认值`);
    }
    const tuned = resolveConfig({ memory: { dir: '.my-memory', userLimitBytes: 8192, projectLimitBytes: 32768, ledgerLimit: 120, archiveKeep: 12, autoLedger: false } });
    assert.equal(tuned.memory.dir, '.my-memory');
    assert.equal(tuned.memory.userLimitBytes, 8192);
    assert.equal(tuned.memory.projectLimitBytes, 32768);
    assert.equal(tuned.memory.ledgerLimit, 120);
    assert.equal(tuned.memory.archiveKeep, 12);
    assert.equal(tuned.memory.autoLedger, false);
    assert.equal(resolveConfig({ memory: { userLimitBytes: 1 } }).memory.userLimitBytes, 512, '下界收敛');
    assert.equal(resolveConfig({ memory: { projectLimitBytes: 9_999_999 } }).memory.projectLimitBytes, 262_144, '上界收敛');
    assert.equal(resolveConfig({ memory: { archiveKeep: 0 } }).memory.archiveKeep, 1);
    assert.throws(() => resolveConfig({ memory: { dir: 42 } }), /memory\.dir/);
    assert.equal(resolveConfig({ memory: { enabled: 'yes' } }).memory.enabled, true, '非布尔值按缺省处理');
    assert.equal(resolveConfig({ memory: { enabled: false } }).memory.enabled, false);
});

await check('关掉记忆后 office_memory 工具不再注册', () => {
    const tools = buildTools(resolveConfig({ tools: { office_memory: false } }));
    assert.ok(!tools.some((t) => t.name === 'office_memory'));
    assert.equal(tools.length, 6);
});

await check('文档参数沿用既有默认，不被设置层改坏', () => {
    const cfg = resolveConfig({});
    assert.equal(cfg.defaultTheme, 'plain');
    assert.equal(cfg.scriptTimeoutMs, 60000);
    assert.equal(cfg.cacheDir, '.office/cache');
});

// ── schema 形状（需要 schemastery） ───────────────────────────────────────

await check('schema 覆盖七组：工具 / 记忆 / 检索 / 后台模型 / 文档 / Python / 音频与视频', async () => {
    if (z === undefined) return;
    const S = createOfficeSettings(z);
    const value = plain(S({}));
    assert.deepEqual(Object.keys(value).sort(), ['av', 'documents', 'memory', 'python', 'search', 'subagentModel', 'tools']);
    for (const name of Object.keys(DEFAULT_TOOLS)) {
        assert.equal(typeof value.tools[name], 'boolean', `tools.${name} 应是布尔`);
    }
    for (const name of Object.keys(DEFAULT_MEMORY)) {
        assert.equal(typeof value.memory[name], typeof DEFAULT_MEMORY[name], `memory.${name} 类型应与默认一致`);
    }
    // 层拓扑的三个开关必须是嵌套对象，而不是被压成一个布尔。
    assert.deepEqual(Object.keys(value.memory.layers).sort(), ['archive', 'hot', 'ledger']);
    for (const name of Object.keys(DEFAULT_QUOTA)) {
        assert.equal(typeof value.memory.quota[name], 'number', `memory.quota.${name} 应是数字`);
    }
    for (const name of Object.keys(DEFAULT_SUBAGENT_MODEL)) {
        assert.equal(typeof value.subagentModel[name], typeof DEFAULT_SUBAGENT_MODEL[name], `subagentModel.${name} 类型应与默认一致`);
    }
});

await check('schema 默认值与 config.js 的默认一致', async () => {
    if (z === undefined) return;
    const value = plain(createOfficeSettings(z)({}));
    assert.deepEqual(value.tools, DEFAULT_TOOLS);
    assert.deepEqual(value.memory, resolveConfig({}).memory);
    assert.equal(value.search.maxParallel, DEFAULT_SEARCH.maxParallel);
    assert.equal(value.search.resultLimit, DEFAULT_SEARCH.resultLimit);
    assert.equal(value.search.outputDir, DEFAULT_SEARCH.outputDir);
    assert.equal(value.documents.defaultTheme, resolveConfig({}).defaultTheme);
    assert.equal(value.documents.scriptTimeoutMs, resolveConfig({}).scriptTimeoutMs);
    assert.deepEqual(value.python, resolveConfig({}).python);
    // av 组只把常用项放进设置页（VAD 细项与缓存子目录只在配置文件里），
    // 所以这里逐键比对界面暴露的那些，不用 deepEqual 比整组。
    const avDefaults = resolveConfig({}).av;
    for (const key of Object.keys(value.av)) {
        assert.deepEqual(value.av[key], avDefaults[key], `av.${key} 的默认值要与 config.js 一致`);
    }
});

await check('音频与视频组：默认值与 config 一致，语言是五种加 auto，坏值被拒', async () => {
    if (z === undefined) return;
    const S = createOfficeSettings(z);
    const value = plain(S({}));
    assert.equal(value.av.enabled, true, '默认开启（设备上本来就有 ffmpeg 与 SenseVoice）');
    assert.equal(value.av.language, 'auto');
    assert.equal(value.av.chunkSeconds, 120);
    assert.equal(value.av.maxSeconds, 3600);
    assert.equal(value.av.precision, 'int8');
    assert.equal(value.av.frameFormat, 'jpg');
    const picked = plain(S({ av: { language: 'yue', precision: 'fp32', frameFormat: 'png' } }));
    assert.equal(picked.av.language, 'yue');
    assert.equal(picked.av.precision, 'fp32');
    assert.equal(picked.av.frameFormat, 'png');
    assert.throws(() => S({ av: { language: 'fr' } }), /language/);
    assert.throws(() => S({ av: { precision: 'q4' } }), /precision/);
    assert.throws(() => S({ av: { chunkSeconds: 1 } }), /chunkSeconds/);
    assert.throws(() => S({ av: { maxFrames: 999 } }), /maxFrames/);
    return '七组之一的 av 组已接入设置页 schema';
});

await check('设置页能选检索通道与预处理强度（第二十轮的新字段）', async () => {
    if (z === undefined) return;
    const value = plain(createOfficeSettings(z)({}));
    assert.equal(value.search.provider, DEFAULT_SEARCH.provider, '默认 auto');
    assert.deepEqual(value.search.providerOrder, DEFAULT_SEARCH.providerOrder, '默认顺序与 config 一致');
    assert.deepEqual(value.search.providers, DEFAULT_SEARCH.providers, '每条通道的参数都有默认值');
    assert.deepEqual(value.search.preprocess, DEFAULT_SEARCH.preprocess, '预处理默认与 config 一致');
    // 可选值就是注册表里的那几条（写错的通道名必须被 schema 拒掉，而不是静默退回 auto）。
    const S = createOfficeSettings(z);
    const picked = plain(S({ search: { provider: 'bocha' } }));
    assert.equal(picked.search.provider, 'bocha');
    assert.throws(() => S({ search: { provider: 'gugou' } }), /provider/);
    assert.throws(() => S({ search: { preprocess: { mode: 'reader' } } }), /mode/);
    assert.throws(() => S({ search: { preprocess: { minLineChars: 200 } } }), /minLineChars/);
    const tuned = plain(S({ search: { providerOrder: ['duckduckgo', 'seam'], preprocess: { mode: 'plain', minLineChars: 0 } } }));
    assert.deepEqual(tuned.search.providerOrder, ['duckduckgo', 'seam']);
    assert.equal(tuned.search.preprocess.mode, 'plain');
    assert.equal(tuned.search.preprocess.minLineChars, 0);
});

await check('越界数值被 schema 拒绝（设置页调不坏插件）', async () => {
    if (z === undefined) return;
    const S = createOfficeSettings(z);
    assert.throws(() => S({ search: { maxParallel: 99 } }), /maxParallel/);
    assert.throws(() => S({ documents: { scriptTimeoutMs: 10 } }), /scriptTimeoutMs/);
    assert.throws(() => S({ documents: { defaultTheme: '不存在的主题' } }), /defaultTheme/);
    assert.throws(() => S({ memory: { userLimitBytes: 1 } }), /userLimitBytes/);
    assert.throws(() => S({ memory: { archiveKeep: 0 } }), /archiveKeep/);
    assert.throws(() => S({ python: { timeoutMs: 10 } }), /timeoutMs/);
});

await check('子代理工具白名单默认就是极简七项', async () => {
    if (z === undefined) return;
    const value = plain(createOfficeSettings(z)({}));
    assert.deepEqual(value.search.subagentTools, [...CHANNEL_TOOLS]);
    assert.equal(value.search.subagentTools.length, 7);
});

await check('设置页目录里的工具与实际白名单一一对应', () => {
    // 设置页按 SUBAGENT_TOOL_CATALOG 渲染多选；它若与 CHANNEL_TOOLS 漂移，
    // 用户会看到勾选项和实际生效的工具对不上。
    assert.deepEqual(SUBAGENT_TOOL_CATALOG.map((item) => item.id), [...CHANNEL_TOOLS]);
    for (const item of SUBAGENT_TOOL_CATALOG) {
        assert.equal(typeof item.label, 'string');
        assert.ok(item.label.length > 0, `${item.id} 缺少显示名`);
        assert.equal(typeof item.note, 'string');
    }
});

await check('命名空间合法（小写字母开头，只含小写/数字/连字符）', () => {
    assert.match(OFFICE_SETTINGS_NS, /^[a-z][a-z0-9-]*$/);
});

await check('LIMITS 与 schema 的区间一致', async () => {
    if (z === undefined) return;
    const S = createOfficeSettings(z);
    // 越界两侧都该被拒：证明 schema 用的就是 LIMITS 里的边界。
    assert.throws(() => S({ search: { maxParallel: LIMITS.searchMaxParallel.max + 1 } }));
    assert.throws(() => S({ search: { maxParallel: LIMITS.searchMaxParallel.min - 1 } }));
    const ok = plain(S({ search: { maxParallel: LIMITS.searchMaxParallel.max } }));
    assert.equal(ok.search.maxParallel, LIMITS.searchMaxParallel.max);
});

await check('office_help 的 settings 话题讲清了开关与参数', async () => {
    const { buildHelp } = await import('../src/docs.js');
    const help = await buildHelp('settings');
    assert.equal(help.topic, 'settings');
    for (const wanted of ['工具开关', 'maxParallel', 'resultLimit', 'subagentTools', 'outputDir', 'defaultTheme', '记忆', 'userLimitBytes', 'ledgerLimit']) {
        assert.ok(help.text.includes(wanted), `settings 话题缺 ${wanted}`);
    }
    const index = await buildHelp('');
    assert.ok(index.text.includes('settings：'), '索引里要能查到 settings');
});

await check('office_help 的 memory 话题讲清了三层与写入协议', async () => {
    const { buildHelp } = await import('../src/docs.js');
    const help = await buildHelp('memory');
    assert.equal(help.topic, 'memory');
    for (const wanted of ['热记忆', '台账', '归档', 'replace', 'oldText', '不记', '记忆目录', '按工作目录']) {
        assert.ok(help.text.includes(wanted), `memory 话题缺 ${wanted}`);
    }
    const index = await buildHelp('');
    assert.ok(index.text.includes('memory：'), '索引里要能查到 memory');
});

await check('服务端 schema 能被浏览器侧还原（设置页拿得到值）', async () => {
    // 这条对着浏览器真正做的事：宿主把 schema.toJSON() 和解析后的值发给页面，
    // 页面用 new Schema(serialized) 还原再校验。二者对不上，页面就会一直
    // 停在「正在读取设置…」——所以这个往返必须测。
    if (z === undefined) return;
    const SchemaCtor = z.Schema ?? z;
    if (typeof SchemaCtor !== 'function') return;
    const schema = createOfficeSettings(z);
    const served = { ns: OFFICE_SETTINGS_NS, value: plain(schema({})), schema: schema.toJSON(), revision: 1 };
    const rehydrated = new SchemaCtor(served.schema);
    let failure;
    try { rehydrated(served.value); failure = undefined; } catch (error) { failure = error.message; }
    assert.equal(failure, undefined, '宿主发出去的值必须能通过还原后的 schema 校验');
    const decoded = plain(rehydrated(served.value));
    assert.equal(decoded.tools.office_run, true);
    assert.equal(decoded.search.maxParallel, DEFAULT_SEARCH.maxParallel);
    assert.equal(decoded.documents.defaultTheme, resolveConfig({}).defaultTheme);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`settings: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);