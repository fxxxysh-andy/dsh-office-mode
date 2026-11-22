/**
 * 「极简模式里不注入 office_* 工具」的回归测试（第五十轮）。
 *
 * 这一份守的是三件事，都是这一轮实测出来的结构性事实：
 *
 *   1. **机制是 restrict，不是 preset 的禁用行**。本插件的工具注册在全局层，
 *      preset 里写 `disabled: true` 关不掉它（实测：ptc preset 的
 *      tool-workflow / tool-ralph / tool-plugin-manager 三行都是 disabled，
 *      而 profile 层把这三行打开后，ptc 会话的工具面里三者都在）。
 *      能按 agent 隐藏继承工具的只有 `ctx.tools.restrict({ deny })`。
 *   2. **判据只看 preset id，读不到就不隐藏**。没有 `agentPresets` 服务的部署
 *      （单会话进程级组合）不该被误伤。
 *   3. **掩码跟着工具面走**。`restrict()` 会校验名字是不是已知的全局工具，
 *      所以掩码里的名字必须按**当前注册**的工具现算：设置页关掉一个工具后
 *      掩码不能还带着它（会抛），重新打开后要补进已在跑的极简会话。
 *
 * 跑法：node test/preset-guard.mjs
 */
import assert from 'node:assert/strict';

import {
    GUARDED_TOOL_NAMES,
    OFFICE_DEV_TOOLS,
    TOOL_HIDDEN_PRESETS,
    enabledOfficeToolNames,
    installPresetGuard,
    shouldHideOfficeTools,
} from '../src/preset-guard.js';
import { resolveConfig } from '../src/config.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const OFFICE_TOOLS = [
    'office_help', 'office_run', 'office_memory', 'office_web_search', 'office_web_fetch',
    'office_search_run', 'office_search_dispatch', 'office_parse_findings',
];

/**
 * 与 test/capabilities.mjs 同一套假 ctx（只够 apply() 走完注册那几步），
 * 另加两个服务：`agentPresets`（读 preset）与 `agents`（找活着的 agent）。
 */
function fakeCtx({ presetOf = () => undefined, withPresets = true, withAgents = true, subagents = undefined } = {}) {
    const tools = [];
    const variables = [];
    const handlers = new Map();
    const live = new Map();
    /** 宿主设置文档的广播通道（index.js 通过 ctx.inject(['settings']) 注册）。 */
    const settingsHandlers = new Map();
    const ctx = {
        tools: { register: (definition) => { tools.push(definition.name); return () => {}; } },
        get: (key) => {
            if (key === 'systemPrompt') {
                return {
                    section: () => () => {},
                    variable: (name, provider) => { variables.push({ name, provider }); return () => {}; },
                };
            }
            if (key === 'agentPresets') {
                return withPresets ? { composedPreset: (agentCtx) => presetOf(agentCtx) } : undefined;
            }
            if (key === 'agents') {
                return withAgents ? { list: () => [...live.values()], get: (id) => live.get(id) } : undefined;
            }
            if (key === 'subagents') return subagents;
            return undefined;
        },
        // **立即回调**（与 capabilities.mjs 同一写法）：以前这里是空函数，于是
        // 「设置改动 → 重建 → 给活着的会话重算掩码」那条链路在测试里从来没跑过 ——
        // 复核正是由此发现「对新会话生效」这句文案与实现相反（第五十四轮 P2-1）。
        inject: (deps, callback) => {
            if (typeof callback === 'function' && Array.isArray(deps) && deps.includes('settings')) {
                callback({
                    on: (event, handler) => {
                        const list = settingsHandlers.get(event) ?? [];
                        list.push(handler);
                        settingsHandlers.set(event, list);
                        return () => {
                            const current = settingsHandlers.get(event) ?? [];
                            settingsHandlers.set(event, current.filter((item) => item !== handler));
                        };
                    },
                });
            }
        },
        on: (event, listener) => {
            const list = handlers.get(event) ?? [];
            list.push(listener);
            handlers.set(event, list);
            return () => {
                const current = handlers.get(event) ?? [];
                handlers.set(event, current.filter((item) => item !== listener));
            };
        },
        logger: { warn: () => {} },
    };
    const emit = (event, ...args) => {
        for (const listener of [...(handlers.get(event) ?? [])]) listener(...args);
    };
    /** 广播设置文档变化（宿主 → 插件 → rebuild）。 */
    const broadcast = (ns, ...args) => {
        for (const listener of [...(settingsHandlers.get('settings/document-updated') ?? [])]) listener(ns, ...args);
    };
    return { ctx, tools, variables, handlers, live, emit, broadcast };
}

/** 一个「agent」：只需要 ctx（restrict 挂在它作用域的 tools 代理上）。 */
function fakeAgent(id, { restrict, present } = {}) {
    const calls = [];
    const agent = {
        id,
        ctx: {
            tools: {
                restrict: (filter) => {
                    calls.push(filter);
                    if (typeof restrict === 'function') return restrict(filter, calls.length);
                    return () => {};
                },
                // 公开 API 的全局视图：`get(name)` 返回定义或 undefined。
                // 默认「都在」，用例要模拟「本机没有某个工具」时传 present。
                get: (name) => (typeof present === 'function'
                    ? (present(name) ? { name } : undefined)
                    : { name }),
            },
        },
    };
    return { agent, calls };
}

// ── 1. 纯判据 ─────────────────────────────────────────────────────────────

await check('隐藏清单只含内置的极简模式（minimal），办公与其它模式照旧', () => {
    assert.deepEqual(TOOL_HIDDEN_PRESETS, ['minimal']);
    assert.equal(shouldHideOfficeTools('minimal'), true);
    for (const other of ['office', 'ptc', 'cordis', 'standard', '']) {
        assert.equal(shouldHideOfficeTools(other), false, `${other} 不该被隐藏`);
    }
    assert.equal(shouldHideOfficeTools(undefined), false, '读不到 preset 时不能隐藏（preset-free 部署照旧可用）');
    assert.equal(shouldHideOfficeTools(null), false);
    assert.equal(shouldHideOfficeTools('minimal', []), false, '清单为空时整层守卫是空操作');
    return `清单 ${TOOL_HIDDEN_PRESETS.join(' / ')}`;
});

await check('掩码里的名字按当前工具面现算（关掉的工具不进去，避免 restrict 抛未知工具）', () => {
    assert.deepEqual(GUARDED_TOOL_NAMES, OFFICE_TOOLS, '键集就是八个工具名，顺序与工具面一致');
    assert.deepEqual(enabledOfficeToolNames(resolveConfig({})), OFFICE_TOOLS, '空配置 = 八个全开');
    const off = enabledOfficeToolNames(resolveConfig({ tools: { office_run: false, office_memory: false } }));
    assert.equal(off.includes('office_run'), false);
    assert.equal(off.includes('office_memory'), false);
    assert.equal(off.length, OFFICE_TOOLS.length - 2);
    assert.deepEqual(enabledOfficeToolNames(undefined), OFFICE_TOOLS, '拿不到 resolved 时按全量（守卫自己会再判一次）');
    return `${off.length} / ${OFFICE_TOOLS.length}`;
});

// ── 2. 守卫的接线 ─────────────────────────────────────────────────────────

await check('极简模式的 agent 建起来时被装上拒绝掩码，名字就是八个工具', () => {
    const { ctx, emit } = fakeCtx({ presetOf: () => 'minimal' });
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS });
    const { agent, calls } = fakeAgent('a1');
    emit('agent/created', { agent });
    assert.equal(calls.length, 1, '建 agent 时要判一次');
    assert.deepEqual(calls[0], { deny: OFFICE_TOOLS });
    guard.dispose();
});

await check('别的 preset 的 agent 一动不动（办公 / ptc / 创造模式都要照旧能调 office_run）', () => {
    const presets = new Map();
    const { ctx, emit } = fakeCtx({ presetOf: (agentCtx) => presets.get(agentCtx) });
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS });
    const { agent, calls } = fakeAgent('a2');
    presets.set(agent.ctx, 'ptc');
    emit('agent/created', { agent });
    assert.equal(calls.length, 0);
    guard.dispose();
});

await check('读不到 preset 服务的部署不隐藏（单会话进程级组合不该被误伤）', () => {
    const { ctx, emit } = fakeCtx({ withPresets: false });
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS });
    const { agent, calls } = fakeAgent('a3');
    emit('agent/created', { agent });
    assert.equal(calls.length, 0);
    guard.dispose();
});

await check('composedPreset 抛错时只当「读不到」，不把插件打挂', () => {
    const { ctx, emit } = fakeCtx({ presetOf: () => { throw new Error('mount 还没结算'); } });
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS });
    const { agent, calls } = fakeAgent('a4');
    emit('agent/created', { agent });
    assert.equal(calls.length, 0);
    guard.dispose();
});

await check('空白会话从极简模式换到办公模式：掩码要撤掉，否则办公工具被永久挡住', () => {
    let preset = 'minimal';
    const revoked = [];
    const { ctx, emit, live } = fakeCtx({ presetOf: () => preset });
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS });
    const { agent, calls } = fakeAgent('a5', { restrict: () => () => revoked.push('a5') });
    live.set('a5', agent);
    emit('agent/created', { agent });
    assert.equal(calls.length, 1, '建起来时是极简模式 → 装掩码');
    preset = 'office';
    emit('agent-preset/selected', 'a5', 'office');
    assert.deepEqual(revoked, ['a5'], '换 preset 后要撤销掩码');
    emit('agent/created', { agent });
    assert.equal(calls.length, 1, '撤掉之后不该再装一次（同一次创建只装一次）');
    guard.dispose();
});

await check('设置页关掉工具后重建：掩码按新工具面重算，重新打开又补回去', () => {
    const { ctx, emit, live } = fakeCtx({ presetOf: () => 'minimal' });
    let names = OFFICE_TOOLS.filter((name) => name !== 'office_run');
    const revoked = [];
    const guard = installPresetGuard(ctx, { toolNames: () => names });
    const { agent, calls } = fakeAgent('a6', { restrict: () => () => revoked.push(calls.length) });
    live.set('a6', agent);
    emit('agent/created', { agent });
    assert.equal(calls.length, 1);
    assert.equal(calls[0].deny.includes('office_run'), false, '关掉的工具不能出现在掩码里（restrict 会判未知工具）');
    // 重建：旧的掩码撤掉，新的按「九个全开」重算。
    names = [...OFFICE_TOOLS];
    guard.resync();
    assert.equal(revoked.length, 1, 'resync 要先撤掉旧掩码');
    assert.equal(calls.length, 2);
    assert.deepEqual(calls[1].deny, OFFICE_TOOLS);
    return `掩码 ${calls[0].deny.length} → ${calls[1].deny.length}`;
});

await check('restrict 抛错（宿主判未知工具）只记警告，不往外冒', () => {
    const warnings = [];
    const { ctx, emit } = fakeCtx({ presetOf: () => 'minimal' });
    ctx.logger = { warn: (message) => warnings.push(message) };
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS });
    const { agent } = fakeAgent('a7', { restrict: () => { throw new Error('names unknown global tool "office_help"'); } });
    emit('agent/created', { agent });
    assert.equal(warnings.length, 1);
    assert.match(warnings[0], /隐藏 office 工具失败/);
    // 失败之后可以重试：下次 resync 不该被「已经装过」挡住。
    guard.resync();
    assert.equal(warnings.length, 1, '没有 agents 服务时不重复尝试');
    guard.dispose();
});

await check('dispose 之后不再对新建 agent 装掩码，并把已装的撤掉', () => {
    const { ctx, emit } = fakeCtx({ presetOf: () => 'minimal' });
    const revoked = [];
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS });
    const { agent, calls } = fakeAgent('a8', { restrict: () => () => revoked.push('a8') });
    emit('agent/created', { agent });
    assert.equal(calls.length, 1);
    guard.dispose();
    assert.deepEqual(revoked, ['a8'], 'dispose 要把已装的掩码撤掉');
    emit('agent/created', { agent });
    assert.equal(calls.length, 1, 'dispose 之后不再装');
    guard.dispose();
    assert.deepEqual(revoked, ['a8'], 'dispose 幂等');
});

await check('没有 ctx.on / 清单为空时是安全的空操作', () => {
    const empty = installPresetGuard({}, { toolNames: () => OFFICE_TOOLS });
    empty.dispose();
    empty.resync();
    const { ctx, emit } = fakeCtx({ presetOf: () => 'minimal' });
    const guard = installPresetGuard(ctx, { hidden: [], toolNames: () => OFFICE_TOOLS });
    const { agent, calls } = fakeAgent('a9');
    emit('agent/created', { agent });
    assert.equal(calls.length, 0);
    guard.dispose();
});

// ── 2b. 宿主工具那一半（第五十四轮 新-15①） ────────────────────────────────
//
// 这一路默认关着（配置 `guard.hideDevToolsInOffice`）。打开后办公会话里要把
// preset 声明要关的那批**宿主**工具真正掩掉；与 office 工具那一路的两点不同是
// 判据的核心：**逐个名字装**（宿主工具是别的插件注册的，不同机器装的东西不同），
// 且**只跳过「不认识」这一种失败**（别的失败照旧上报，不许变成静默的假成功）。

await check('宿主工具那一半：清单为空时一个都不掩（默认就是这一档）', () => {
    const { ctx, emit } = fakeCtx({ presetOf: () => 'office' });
    const guard = installPresetGuard(ctx, { toolNames: () => OFFICE_TOOLS, hostHidden: [] });
    const { agent, calls } = fakeAgent('h1');
    emit('agent/created', { agent });
    assert.equal(calls.length, 0, '默认不隐藏：办公会话的工具面一个字不动');
    guard.dispose();
});

await check('宿主工具那一半：打开后在办公 preset 里逐个装掩码，dispose 要把每个都撤掉', () => {
    const revoked = [];
    const { ctx, emit } = fakeCtx({ presetOf: () => 'office' });
    const guard = installPresetGuard(ctx, {
        toolNames: () => OFFICE_TOOLS,
        hostHidden: ['office'],
        hostToolNames: () => ['bash', 'todo_write', 'task_board_list'],
    });
    const { agent, calls } = fakeAgent('h2', { restrict: (filter) => () => revoked.push(filter.deny[0]) });
    emit('agent/created', { agent });
    assert.deepEqual(calls, [{ deny: ['bash'] }, { deny: ['todo_write'] }, { deny: ['task_board_list'] }],
        '宿主工具逐个名字装（合起来装会被一个不认识的名字整批打回）');
    guard.dispose();
    assert.deepEqual(revoked, ['bash', 'todo_write', 'task_board_list'], '每个掩码的撤销函数都要留着并调用');
});

await check('宿主工具那半：本机没有的名字跳过并记警告，别的失败照样上报', () => {
    const warnings = [];
    const { ctx, emit } = fakeCtx({ presetOf: () => 'office' });
    ctx.logger = { warn: (message) => warnings.push(message) };
    const guard = installPresetGuard(ctx, {
        hidden: [],
        hostHidden: ['office'],
        hostToolNames: () => ['bash', 'workflow', 'todo_write'],
    });
    const { agent, calls } = fakeAgent('h3', {
        restrict: (filter) => {
            if (filter.deny[0] === 'workflow') {
                throw new Error('tools.restrict() names unknown global tool "workflow"; known global tools: bash, todo_write');
            }
            if (filter.deny[0] === 'todo_write') throw new Error('tools.restrict() requires a scoped context');
            return () => {};
        },
    });
    emit('agent/created', { agent });
    assert.equal(calls.length, 3, '三个名字都要试过 —— 一个失败不该把整批丢掉');
    assert.ok(warnings.some((line) => /本机不存在，已跳过（同一清单只报一次）：workflow/.test(line)),
        `跳过要留痕（否则「以为藏了、其实没藏」没人发现）：${warnings.join('\n')}`);
    assert.ok(warnings.some((line) => /部分工具隐藏失败.*todo_write/.test(line)),
        `不是「不认识」的失败要照旧上报：${warnings.join('\n')}`);
    guard.dispose();
});

await check('宿主工具那半：办公 → 极简换 preset 时两路掩码按新清单重算', () => {
    let preset = 'office';
    const { ctx, emit, live } = fakeCtx({ presetOf: () => preset });
    const guard = installPresetGuard(ctx, {
        toolNames: () => OFFICE_TOOLS,
        hostHidden: ['office'],
        hostToolNames: () => ['bash'],
    });
    const { agent, calls } = fakeAgent('h4');
    live.set('h4', agent);
    emit('agent/created', { agent });
    assert.deepEqual(calls.map((call) => call.deny[0]), ['bash'], '办公：只掩宿主工具');
    preset = 'minimal';
    emit('agent-preset/selected', 'h4', 'minimal');
    // 极简模式不在 hostHidden 里 → 宿主那一条撤掉；改掩 office 九个。
    assert.deepEqual(calls[1].deny, OFFICE_TOOLS, '极简：掩 office 工具，不再掩宿主工具');
    guard.dispose();
});

await check('宿主工具那半：先用 tools.get 的正向判据筛名字，本机没有的连 restrict 都不试', () => {
    const warnings = [];
    const { ctx, emit } = fakeCtx({ presetOf: () => 'office' });
    ctx.logger = { warn: (message) => warnings.push(message) };
    const guard = installPresetGuard(ctx, {
        hidden: [],
        hostHidden: ['office'],
        hostToolNames: () => ['bash', 'workflow', 'todo_write'],
    });
    // 正向判据：workflow 不在全局视图里（本机没装）→ 直接跳过；
    // 另外两个在 → 才去 restrict。这样 restrict 只可能因为真错误抛。
    const { agent, calls } = fakeAgent('h3b', { present: (name) => name !== 'workflow' });
    emit('agent/created', { agent });
    assert.deepEqual(calls.map((call) => call.deny[0]), ['bash', 'todo_write'], '不存在的名字不该进 restrict');
    assert.ok(warnings.some((line) => /本机不存在，已跳过/.test(line)), warnings.join('\n'));
    guard.dispose();
});

await check('宿主工具那半：上次有跳过的名单会在下一次 sync 重试（本机后来装上了要补掩）', () => {
    const { ctx, emit, live } = fakeCtx({ presetOf: () => 'office' });
    let installed = false;
    const guard = installPresetGuard(ctx, {
        hidden: [],
        hostHidden: ['office'],
        hostToolNames: () => ['bash', 'workflow'],
    });
    const { agent, calls } = fakeAgent('h3c', { present: (name) => name !== 'workflow' || installed });
    live.set('h3c', agent);
    emit('agent/created', { agent });
    assert.deepEqual(calls.map((call) => call.deny[0]), ['bash'], '第一次：workflow 本机没有');
    // 本机把 workflow 装上了 → 下一次 resync 必须补掩（只记「装过的清单」就会漏掉它）。
    installed = true;
    guard.resync();
    assert.deepEqual(calls.map((call) => call.deny[0]), ['bash', 'bash', 'workflow'], '清单没变但有跳过时要重试');
    // 第二次之后没有跳过了 → 再 resync 不该重复装。
    guard.resync();
    assert.equal(calls.length, 3, '装齐之后不再重装');
    guard.dispose();
});

// ── 3. 端到端接线：apply() 真的装上了这层守卫 ──────────────────────────────

await check('apply() 之后 agent/created 就能触发掩码（防「模块写好了没接进 index.js」）', async () => {
    const { apply } = await import('../src/index.js');
    const { ctx, emit, tools } = fakeCtx({ presetOf: () => 'minimal' });
    apply(ctx, {});
    assert.ok(tools.length >= 8, `工具面至少要注册八个工具，实际 ${tools.length}`);
    const { agent, calls } = fakeAgent('a10');
    emit('agent/created', { agent });
    assert.equal(calls.length, 1, 'apply() 必须装上 agent/created 守卫');
    assert.deepEqual(calls[0].deny, OFFICE_TOOLS);
    // 设置改动重建（宿主广播 settings/document-updated）之后守卫仍在。
    emit('settings/document-updated', 'dsh-office-mode');
    const { agent: second, calls: secondCalls } = fakeAgent('a11');
    emit('agent/created', { agent: second });
    assert.equal(secondCalls.length, 1, 'rebuild 之后守卫要重新装上');
    assert.deepEqual(secondCalls[0].deny, OFFICE_TOOLS);
});

await check('apply() 里工具全关时守卫不下发空掩码（restrict({}) 会抛）', async () => {
    const { apply } = await import('../src/index.js');
    const allOff = Object.fromEntries(OFFICE_TOOLS.map((name) => [name, false]));
    const { ctx, emit } = fakeCtx({ presetOf: () => 'minimal' });
    apply(ctx, { tools: allOff });
    const { agent, calls } = fakeAgent('a12');
    emit('agent/created', { agent });
    assert.equal(calls.length, 0);
});

await check('apply()：配置打开守卫时办公会话被掩掉宿主工具，默认（不传 guard）一个字不动', async () => {
    const { apply } = await import('../src/index.js');
    const on = fakeCtx({ presetOf: () => 'office' });
    apply(on.ctx, { guard: { hideDevToolsInOffice: true } });
    const { agent, calls } = fakeAgent('a13');
    on.emit('agent/created', { agent });
    assert.deepEqual(calls.map((call) => call.deny[0]).sort(), [...OFFICE_DEV_TOOLS].sort(),
        '打开后办公会话要逐个掩掉清单里的宿主工具');

    const off = fakeCtx({ presetOf: () => 'office' });
    apply(off.ctx, {});
    const { agent: idle, calls: idleCalls } = fakeAgent('a14');
    off.emit('agent/created', { agent: idle });
    assert.equal(idleCalls.length, 0, '默认不隐藏是**用户定的口径**：这一项不许被谁改成默认开');
});

await check('设置改动**即时生效**：广播 settings/document-updated 后活着的办公会话被掩上、再关掉被撤下', async () => {
    // 这一条钉的是「即时生效」这个口径（复核 P2-1：四处文案曾写成「对新会话生效」，
    // 而实现是 rebuild 末尾 resync() 遍历活着的 agent 重算）。以前这条链路在测试里
    // 跑不到 —— fakeCtx 的 inject 是空函数，于是那条文案错了很久也没人发现。
    const { apply } = await import('../src/index.js');
    const { ctx, emit, live, broadcast } = fakeCtx({ presetOf: () => 'office' });
    // 宿主设置是**就地更新**的引用：插件 refresh() 重新解析的正是这个对象。
    const config = {};
    apply(ctx, config);
    const revoked = [];
    const { agent, calls } = fakeAgent('a15', { restrict: () => () => revoked.push('a15') });
    live.set('a15', agent);
    emit('agent/created', { agent });
    assert.equal(calls.length, 0, '默认关着：建起来时一个掩码都没有');
    // 用户在设置页打开开关 → 宿主广播 → 插件 rebuild → resync 到**同一个**活着的会话。
    config.guard = { hideDevToolsInOffice: true };
    broadcast('dsh-office-mode');
    assert.ok(calls.length >= OFFICE_DEV_TOOLS.length,
        `广播之后活着的会话要被掩上，实际 ${calls.length} 个掩码`);
    // 再关掉：掩码要撤掉（不然那些入口被永久挡住），且不再装新的。
    config.guard = { hideDevToolsInOffice: false };
    const before = calls.length;
    broadcast('dsh-office-mode');
    assert.equal(calls.length, before, '关掉之后不再装新掩码');
    assert.ok(revoked.length >= OFFICE_DEV_TOOLS.length, `关掉要把已装的撤掉，实际撤了 ${revoked.length} 个`);
    // 别的命名空间的事件不该触发重建。
    config.guard = { hideDevToolsInOffice: true };
    broadcast('dsh-other-plugin');
    assert.equal(calls.length, before, '别的命名空间的广播不该重建本插件');
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok && item.note !== undefined ? `  → ${item.note}` : ''}`);
    if (!item.ok) console.log(`       | ${item.error?.message ?? item.error}`);
}
console.log(`preset-guard: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
