/**
 * 设置保存的宿主侧回归：真实 SettingsForms + 真实 revision/CAS + 本插件真实 schema。
 *
 * 为什么单开一份而不是并进 test/settings.mjs：settings.mjs 测的是 schema 形状与
 * 取值收敛（纯函数）；这一份测的是「把设置写进宿主」这条链路本身 —— 用户现场
 * 保存失败报的是 configForms 那句固定文案，要把「宿主到底为什么拒」和
 * 「客户端拿到的 revision 协议对不对」分开证，必须有真正的 SettingsForms。
 *
 * 现场（2026-11-20）：设置页保存 memory.llm.provider=dsh-opencode-go /
 * memory.llm.model=deepseek-v4.1-flash，页面显示
 * 「宿主拒绝了这次写入（configForms 这条通道只回 true/false，服务端原因被丢掉了；
 *   remote.settings 也不可用，所以这里拿不到细节）」。
 *
 * 这一份要证三件事：
 *   1. 这两个值本身是合法写入 —— schema 的 volatile 覆盖完整、值不越界，
 *      带当前 revision 的一笔 mutate 能成功落进 entry 的用户层。
 *   2. 拒绝只会来自「revision 过期」（settings/conflict：宿主不落盘、可重读重试）
 *      或 schema 校验（settings/rejected：带宿主原话）。前者说明「重试就能成功」，
 *      后者说明「不是这两个字段的问题」。
 *   3. 客户端协议（describe 取 revision → mutate）在真实 CAS 上成立：冲突后重读
 *      到的 revision 能让同一笔写入成功；不传 revision 则会绕过 CAS（所以客户端
 *      必须显式传，否则并发写会互相覆盖）。
 *
 * 边界：真实 configEditor 的 profile patch 落盘 / mergeLayers 是宿主代码，且现场
 * 不许写真实 profile，所以这里的 configEditor 只把「用户层」写回 entry —— 与真
 * ConfigEditor.edit 的语义一致（先 change(current, inherited)，再
 * fiber.resolveConfig(next) 做 schemastery 校验，校验不过直接抛）。
 *
 * 跑法：node test/settings-save.mjs
 */
import assert from 'node:assert/strict';
import { pathToFileURL } from 'node:url';

import { resolveShippedBundle } from './host-modules.mjs';
import { createOfficeSettings, OFFICE_SETTINGS_NS } from '../src/settings.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 从随部署安装的宿主包里解析一个模块（见 host-modules.mjs）。 */
async function loadShipped(spec) {
    const resolved = resolveShippedBundle(spec);
    if (resolved === undefined) return undefined;
    try {
        return await import(pathToFileURL(resolved).href);
    } catch {
        return undefined;
    }
}

const zModule = await loadShipped('@deepseek-ai/schemastery');
const settingsModule = await loadShipped('@deepseek-ai/dsh-settings');
if (zModule === undefined || settingsModule === undefined) {
    console.log('跳过：本机解析不到宿主 @deepseek-ai/schemastery 或 @deepseek-ai/dsh-settings，无法验证宿主写入链路。');
    process.exit(0);
}
const z = zModule.default;
const { SettingsForms, SettingsConflictError } = settingsModule;

// 宿主编译产物里 FiberState.ACTIVE === 2（dsh-settings/lib/index.js:417）；dsh-settings
// 的入口只导出具名类，不导出枚举，所以这里照常量写。
const ACTIVE = 2;
const NS = OFFICE_SETTINGS_NS;
const USER_CONFIG = {
    memory: { llm: { mode: 'host', chunking: true, summary: true } },
};

/**
 * 造一个只承载「一个插件行」的 SettingsForms。
 *
 * 假 ctx 的字段是按 dsh-settings/src/index.ts 真实读到的位置逐个给的：
 * configEditor.entries() / configuration() / edit()、profileContext、fiber.state、
 * logger、emit。少一个都能在真机上炸，所以宁可显式给全。
 */
function makeForms(userConfig = structuredClone(USER_CONFIG)) {
    const schema = createOfficeSettings(z);
    const resolved = (cfg) => schema(cfg);
    const emitted = [];
    const entry = {
        id: NS,
        options: { id: NS, name: NS, config: structuredClone(userConfig) },
        fiber: {
            uid: 7,
            state: ACTIVE,
            runtime: { Config: schema },
            effectiveConfig: resolved(structuredClone(userConfig)),
            ctx: { fiber: { state: ACTIVE } },
            resolveConfig(next) { return resolved(next); },
        },
    };
    const ctx = {
        reflect: { provide() {} },
        effect(fn) { fn(); return () => {}; },
        on() { return () => {}; },
        emit(name, ns, revision) { emitted.push([name, ns, revision]); },
        logger: { error() {}, info() {}, warn() {} },
        // dsh-settings 构造时会 root.loader.await 一个 fiber 就绪信号；这里让它悬着，
        // 等价于「组合已经激活」，不触发 ready 回调。
        root: { loader: { await: () => new Promise(() => {}) } },
        profileContext: {
            home: 'C:/tmp/office-settings-save-home',
            name: 'web',
            dir: 'C:/tmp/office-settings-save-home/profiles/web',
            patchPath: 'C:/tmp/office-settings-save-home/profiles/web/cordis.patch.yml',
            installAnchor: 'C:/tmp',
        },
        configEditor: {
            documentPath: 'C:/tmp/office-settings-save-home/profiles/web/cordis.patch.yml',
            entries() { return [entry]; },
            configuration() { return [{ entry, inherited: {}, override: structuredClone(entry.options.config) }]; },
            async edit(target, change) {
                // 与真 ConfigEditor.edit（dsh-config-editor/src/index.ts:100-110）同序：
                // 先跑 change 得到候选用户层，再做 schemastery 校验 —— 校验不过就不落盘。
                const inherited = {};
                const next = change(structuredClone(target.options.config ?? {}), inherited);
                target.fiber.resolveConfig(next);
                target.options.config = next;
                target.fiber.effectiveConfig = resolved(next);
            },
        },
        fiber: { state: ACTIVE },
    };
    const forms = new SettingsForms(ctx);
    const describe = () => forms.describe().find((row) => row.ns === NS);
    return { forms, entry, describe, emitted, schema };
}

// ── 现场：这两个值能写进去 ────────────────────────────────────────────────

await check('describe 给出 dsh-office-mode：整数 revision、memory.llm 默认空 provider/model', () => {
    const { describe } = makeForms();
    const row = describe();
    assert.ok(row, 'describe 里必须有 dsh-office-mode（插件行 id 与客户端 NAMESPACE 必须一致）');
    assert.equal(row.ns, NS);
    assert.ok(Number.isInteger(row.revision), 'revision 必须是整数，客户端要靠它做 CAS：' + String(row.revision));
    assert.equal(row.value.memory.llm.provider, '', '默认 provider 为空');
    assert.equal(row.value.memory.llm.model, '', '默认 model 为空');
});

await check('现场写入：provider + model 同一笔 mutate，带当前 revision → 成功且 revision 进位', async () => {
    const { forms, entry, describe } = makeForms();
    const before = describe();
    await forms.mutate(NS, [
        { op: 'set', path: ['memory', 'llm', 'provider'], value: 'dsh-opencode-go' },
        { op: 'set', path: ['memory', 'llm', 'model'], value: 'deepseek-v4.1-flash' },
    ], before.revision);
    const after = describe();
    assert.equal(after.value.memory.llm.provider, 'dsh-opencode-go', 'provider 应落进用户层');
    assert.equal(after.value.memory.llm.model, 'deepseek-v4.1-flash', 'model 应落进用户层');
    assert.equal(after.revision, before.revision + 1, '一次成功写入应恰好推进一次 revision');
    assert.equal(entry.options.config.memory.llm.provider, 'dsh-opencode-go', '用户层（要落盘的）也必须是新值');
    assert.equal(entry.options.config.memory.llm.model, 'deepseek-v4.1-flash');
    // 其余用户层字段不许被这一笔抹掉：客户端只 patch 被改的字段。
    assert.equal(entry.options.config.memory.llm.mode, 'host', '不相关的字段应保留');
});

await check('写入后广播 settings/document-updated（插件的即时生效链路靠它）', async () => {
    const { forms, describe, emitted } = makeForms();
    const revision = describe().revision;
    await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm', 'provider'], value: 'dsh-opencode-go' }], revision);
    assert.ok(emitted.some(([name, ns]) => name === 'settings/document-updated' && ns === NS),
        '成功后必须广播 settings/document-updated，否则设置改了工具面不重建：' + JSON.stringify(emitted));
});

// ── 拒绝的两种真实形态 ────────────────────────────────────────────────────

await check('revision 过期 → SettingsConflictError（settings/conflict），且一个字段都不落盘', async () => {
    const { forms, entry, describe } = makeForms();
    const stale = describe().revision;
    await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm', 'provider'], value: 'dsh-opencode-go' }], stale);
    const userBefore = structuredClone(entry.options.config);
    let caught = null;
    try {
        await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm', 'mode'], value: 'off' }], stale);
    } catch (error) {
        caught = error;
    }
    assert.ok(caught instanceof SettingsConflictError, '过期 revision 必须抛 SettingsConflictError，实际：' + String(caught));
    assert.equal(caught.code, 'SETTINGS_CONFLICT', '控制器按这个 code 映射成 settings/conflict（dsh-api-settings-controller/src/index.ts:257-266）');
    assert.equal(caught.expected, stale);
    assert.equal(caught.actual, stale + 1);
    assert.match(caught.message, /settings namespace "dsh-office-mode" changed since it was read/,
        '宿主原话要能透到页面（settings/conflict 的 message 就是它；ns 只在 message 里）');
    assert.deepEqual(entry.options.config, userBefore, '冲突必须不落盘：用户层保持原样');
});

await check('客户端恢复语义：冲突后重读 revision，同一笔写入成功（有限重试的依据）', async () => {
    const { forms, describe } = makeForms();
    const stale = describe().revision;
    await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm', 'provider'], value: 'dsh-opencode-go' }], stale);
    // 客户端第 2 次尝试：重新 describe 拿 revision，再写同一笔。
    await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm', 'model'], value: 'deepseek-v4.1-flash' }], describe().revision);
    assert.equal(describe().value.memory.llm.model, 'deepseek-v4.1-flash', '重读 revision 后必须能写成');
});

await check('不传 revision 会绕过 CAS（所以客户端必须先 describe 再写）', async () => {
    const { forms, describe } = makeForms();
    const stale = describe().revision;
    await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm', 'provider'], value: 'a' }], stale);
    // expectedRevision === undefined：宿主不校验，直接覆盖 —— 这就是客户端不能省 describe 的原因。
    await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm', 'provider'], value: 'b' }], undefined);
    assert.equal(describe().value.memory.llm.provider, 'b');
});

await check('类型不合法 → 宿主 schemastery 原话（settings/rejected 的 message 来源）', async () => {
    const { forms, entry, describe } = makeForms();
    const userBefore = structuredClone(entry.options.config);
    let caught = null;
    try {
        await forms.mutate(NS, [{ op: 'set', path: ['search', 'maxParallel'], value: 'lots' }], describe().revision);
    } catch (error) {
        caught = error;
    }
    assert.ok(caught, '非法类型必须被拒');
    assert.match(String(caught.message), /\$\.search\.maxParallel expected number but got lots/,
        '宿主校验原话要能透到页面：' + String(caught.message));
    assert.deepEqual(entry.options.config, userBefore, '校验失败必须不落盘');
});

await check('非 volatile 路径 → Config field "…" is not volatile（写入面契约）', async () => {
    const { forms, describe } = makeForms();
    let caught = null;
    try {
        await forms.mutate(NS, [{ op: 'set', path: ['memory', 'llm'], value: {} }], describe().revision);
    } catch (error) {
        caught = error;
    }
    assert.ok(caught, '对象容器不是可编辑叶子，必须被拒');
    assert.match(String(caught.message), /Config field "memory\.llm" is not volatile/,
        '消息要带路径：' + String(caught.message));
});

await check('未知命名空间（插件行不在）→ No configurable plugin entry（客户端文案应能区分）', async () => {
    const { forms } = makeForms();
    let caught = null;
    try {
        await forms.mutate('dsh-office-mode-typo', [{ op: 'set', path: ['tools', 'office_help'], value: false }], 0);
    } catch (error) {
        caught = error;
    }
    assert.match(String(caught && caught.message), /No configurable plugin entry "dsh-office-mode-typo"/);
});

// ── 覆盖面：页面能点到的每一条叶子都必须可写 ──────────────────────────────

await check('volatile 全覆盖：describe value 的每一条叶子都能单独写回（不出现 not volatile）', async () => {
    const { forms, describe } = makeForms();
    const leaves = [];
    const walk = (node, path) => {
        if (node === null || typeof node !== 'object') {
            leaves.push([path, node]);
            return;
        }
        for (const [key, child] of Object.entries(node)) walk(child, [...path, key]);
    };
    walk(describe().value, []);
    assert.ok(leaves.length > 20, '叶子数量太少，describe 可能没给全：' + leaves.length);
    const refused = [];
    for (const [path, value] of leaves) {
        try {
            await forms.mutate(NS, [{ op: 'set', path, value }], describe().revision);
        } catch (error) {
            // 只记「not volatile」：其它拒绝说明值本身不合法，不属于标记漏项。
            if (/is not volatile/.test(String(error && error.message))) refused.push(path.join('.') + '：' + String(error.message));
        }
    }
    assert.deepEqual(refused, [], '这些字段没标 volatile，设置页永远写不进去：' + refused.join(' | '));
});

// ── 另一条选择器：subagentModel ───────────────────────────────────────────

await check('subagentModel 的 provider/model 同样可写（页面第二个模型选择器）', async () => {
    const { forms, describe } = makeForms();
    await forms.mutate(NS, [
        { op: 'set', path: ['subagentModel', 'provider'], value: 'dsh-opencode-go' },
        { op: 'set', path: ['subagentModel', 'model'], value: 'deepseek-v4.1-flash' },
    ], describe().revision);
    const row = describe();
    assert.equal(row.value.subagentModel.provider, 'dsh-opencode-go');
    assert.equal(row.value.subagentModel.model, 'deepseek-v4.1-flash');
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const detail = item.ok ? '' : '\n      ' + String(item.error && item.error.message ? item.error.message : item.error);
    console.log((item.ok ? 'PASS  ' : 'FAIL  ') + item.name + detail);
}
console.log('settings-save: ' + (results.length - failed.length) + '/' + results.length + ' 通过');
process.exitCode = failed.length > 0 ? 1 : 0;
