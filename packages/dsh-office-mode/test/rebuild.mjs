/**
 * 设置重建事务的回归测试（第五十九轮，档案 §7 第 0 位的 S-04）。
 *
 * 这一份守的是「设置改动 → 重建」那条链路的**失败形态**。第五十五轮结构审查 S-04 说的是：
 * 重建先 `disposeAll()` 再逐项注册，中途任何一步失败都被事件处理器的空 `catch` 吞掉，
 * 而旧资源已经销毁、配置也已经更新 —— 于是运行期既不是旧的稳定版本，也不是完整的新版本。
 * 那时的测试只覆盖了成功路径（`capabilities.mjs` 的「重建后变量不重复」）。
 *
 * 现在的协议（`src/index.js` 的 `rebuildWith`）是：
 *   拆旧的 → 按候选装 → 失败就把候选那半份收干净、把**改动前那份配置整份装回** →
 *   连回滚也失败才进**显式降级态**（不抛错、但一定留日志与可诊断状态）。
 * 而 `installAll()` 把两样「缺了会更糟」的资源排在最前（能力变量、诊断工具 `office_help`），
 * 所以降级态里能剩下什么 = 「回滚失败点之前装到的那几步」。
 *
 * ## 每一条的鉴别力（第五十九轮复核 P3 要求写清）
 *
 * 有鉴别力（换回 HEAD 的 `src/index.js` 会红）：2 / 3 / 4 / 5 / 7 / 8 / 9（实测 2/9 通过）。
 * 只是回归护栏：1（成功路径）与 6（合法设置的边界：工具全关）——
 * 留着它们是防「以后的改动把成功路径或边界弄坏」，不代表它们能发现 S-04 那一类缺陷
 * （第 6 条就是被一次真实的回归换来的：本轮初版让「工具全关」抛错，`preset-guard` 当场红）。
 *
 * 跑法：node test/rebuild.mjs
 */
import assert from 'node:assert/strict';

import { CAPABILITY_VARIABLE } from '../src/capabilities.js';
import { resolveConfig } from '../src/config.js';
import { buildTools } from '../src/tools.js';
import { rebuildState, resetRebuildStateForTest, describeRebuildState } from '../src/rebuild-state.js';
import { proxyStatus, resetProxyStateForTest } from '../src/web-proxy.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 默认配置下的工具面（9 个）：用来数「整份装回来了没有」。 */
const TOOL_COUNT = buildTools(resolveConfig({})).length;

const HEALTHY = Object.freeze({
    ok: true,
    restored: true,
    degraded: false,
    phase: '',
    error: '',
    rollbackPhase: '',
    rollbackError: '',
    droppedProxy: '',
});

/** 查一次用法（不落盘、不碰记忆投影）：读「运行期实况」那一行用。 */
async function helpText(topic) {
    const help = buildTools(resolveConfig({ memory: { enabled: false } })).find((item) => item.name === 'office_help');
    const out = await help.execute({ topic }, { agent: { session: { header: { cwd: process.cwd() } } } });
    return out.text;
}

/**
 * 一套够 `apply()` 走完注册、并且**能按需让某一步注册失败**的假 ctx。
 *
 * 三个关键点，缺一个这套测试就抓不到 S-04 的形态：
 *   1. 工具、提示变量、提示段落都**按名字查重**（照 cordis 的口径 —— 宿主注册表就是这样）：
 *      新旧两份不能并存，「重建后旧资源没被拆掉」与「回滚没把旧资源装回来」都会当场红；
 *   2. `inject(['settings'])` **立即回调**（与 preset-guard.mjs 同一写法）：不然设置事件
 *      那条链路在测试里根本跑不到；
 *   3. 能**在注册过程中插一次重入广播**（`onceReenter`）：宿主读取设置投影时也会广播这个
 *      事件，而 cordis 的 emit 是同步且可重入的。
 *
 * `agents` 要给一个空注册表：`presetGuard.resync()` 读不到 `agents` 服务时会记一条警告，
 * 那是另一条线的事实，不该混进这一套的日志断言里。
 */
function fakeCtx() {
    const tools = new Map();
    const variables = new Map();
    const sections = new Map();
    const warned = [];
    const handlers = new Map();
    /** 注册失败的注入点：`tools` / `variables` 各一个判据函数。 */
    const fail = { tools: null, variables: null };
    /** 下一次工具注册成功后要插的动作（重入广播用）。 */
    let reenter = null;

    const guard = (kind, name) => {
        const matcher = fail[kind];
        return matcher !== null && matcher(name) === true;
    };

    const ctx = {
        tools: {
            get: (name) => tools.get(name),
            register: (definition) => {
                if (guard('tools', definition.name)) throw new Error(`注入的失败：工具注册被打断（${definition.name}）`);
                if (tools.has(definition.name)) throw new Error(`tool "${definition.name}" is already registered`);
                tools.set(definition.name, definition);
                const hook = reenter;
                reenter = null;
                if (typeof hook === 'function') hook();
                return () => { tools.delete(definition.name); };
            },
        },
        get: (key) => {
            if (key === 'systemPrompt') {
                return {
                    section: (value) => {
                        if (sections.has(value.name)) throw new Error(`prompt section "${value.name}" is already registered`);
                        sections.set(value.name, value);
                        return () => { sections.delete(value.name); };
                    },
                    variable: (name, provider) => {
                        if (guard('variables', name)) throw new Error(`注入的失败：变量注册被打断（${name}）`);
                        if (variables.has(name)) throw new Error(`prompt variable "${name}" is already registered`);
                        variables.set(name, provider);
                        return () => { variables.delete(name); };
                    },
                };
            }
            // 空的 agents 注册表：守卫重算时读得到服务，不记那条「读不到 agents」的警告。
            if (key === 'agents') return { list: () => [], get: () => undefined };
            return undefined;
        },
        inject: (deps, callback) => {
            if (typeof callback === 'function' && Array.isArray(deps) && deps.includes('settings')) {
                callback({
                    on: (event, handler) => {
                        handlers.set(event, handler);
                        return () => { handlers.delete(event); };
                    },
                });
            }
        },
        on: (event, listener) => {
            handlers.set(event, listener);
            return () => { handlers.delete(event); };
        },
        logger: { warn: (message) => { warned.push(String(message)); } },
    };

    return {
        ctx,
        tools,
        variables,
        sections,
        warned,
        /** 真的注册在册的那件工具（走注册面，不是另造一份）。 */
        getTool: (name) => tools.get(name),
        /** 宿主广播设置文档更新（只认本插件的命名空间由 index.js 判）。 */
        broadcast: (ns) => {
            const handler = handlers.get('settings/document-updated');
            assert.equal(typeof handler, 'function', '插件必须订阅 settings/document-updated（否则设置改动不会即时生效）');
            handler(ns);
        },
        /** 让下一次（且只有下一次）匹配的注册抛错。 */
        armOnce: (kind, match = () => true) => {
            let armed = true;
            fail[kind] = (name) => {
                if (!armed || match(name) !== true) return false;
                armed = false;
                return true;
            };
        },
        /** 让所有匹配的注册都抛错（连回滚也一起打断）。 */
        armAlways: (kind, match = () => true) => { fail[kind] = (name) => match(name) === true; },
        disarm: (kind) => { fail[kind] = null; },
        /** 下一次工具注册成功后插一次动作（重入广播用）。 */
        onceReenter: (hook) => { reenter = hook; },
    };
}

// ── 1. 成功路径（回归护栏） ────────────────────────────────────────────────

await check('设置改动重建一次：工具面、段落与变量都只留一份，不记日志、状态仍是正常态', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    apply(app.ctx, {});
    assert.equal(app.tools.size, TOOL_COUNT, `激活时工具面是 ${TOOL_COUNT} 个，实际 ${app.tools.size}`);
    assert.equal(app.variables.size, 2, '激活时只注册一个变量');
    assert.equal(app.sections.size, 1, '默认只注册记忆说明段（injectGuide 默认关）');

    const first = app.variables.get(CAPABILITY_VARIABLE);
    app.broadcast('dsh-office-mode');

    assert.equal(app.tools.size, TOOL_COUNT, '重建不该多留或少留工具');
    assert.equal(app.variables.size, 2, '重建之后变量仍然只有一份（旧的必须被拆掉）');
    assert.equal(app.sections.size, 1, '提示段落也必须只有一份（假 ctx 会查重）');
    assert.notEqual(app.variables.get(CAPABILITY_VARIABLE), first, '重建要换上新的一份（设置改了要重探）');
    const text = app.variables.get(CAPABILITY_VARIABLE)();
    assert.ok(typeof text === 'string' && text.includes('本会话能力'), '重装出来的变量要能求值');
    assert.deepEqual(app.warned, [], `成功路径不该记任何日志，实际：${app.warned.join(' | ')}`);
    assert.deepEqual(rebuildState(), HEALTHY, '成功路径要留在正常态');

    // 正常态的**字节纪律**：这一行只有代理那一句，没有换行、没有重建那一段。
    const search = await helpText('search');
    const tail = search.split('【运行期实况】')[1] ?? '';
    assert.ok(tail.includes('出口代理'), '检索这一页要带代理实况');
    assert.equal(tail.split('\n').length, 1, `正常态不许出现第二行，实际：${JSON.stringify(tail)}`);
    assert.ok(!tail.includes('设置重建'), '正常态不该有重建那一段');
});

// ── 2. 候选配置装到一半失败 → 整份装回旧配置 ────────────────────────────────

await check('候选配置装到一半失败：整份装回改动前那份（用工具开关当「旧配置」的指纹）', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    // 宿主设置是**就地更新**的引用：插件每次重建重新解析的正是这个对象（与 preset-guard.mjs 同法）。
    const config = {};
    apply(app.ctx, config);
    const before = app.variables.get(CAPABILITY_VARIABLE);

    // 设置页关掉 office_memory（候选工具面少一个）+ 注入「能力变量注册失败」。
    config.tools = { office_memory: false };
    app.armOnce('variables', (name) => name === CAPABILITY_VARIABLE);
    app.broadcast('dsh-office-mode');

    assert.equal(app.tools.size, TOOL_COUNT, `要整份装回旧配置，工具数应是 ${TOOL_COUNT}，实际 ${app.tools.size}`);
    assert.ok(app.tools.has('office_memory'), '装回的必须是**旧**配置：被关掉的那个工具要回来');
    assert.equal(app.variables.size, 2, '能力变量是 preset 装配的硬依赖：回滚必须把它装回来');
    assert.notEqual(app.variables.get(CAPABILITY_VARIABLE), before, '回滚是重新装一份，不是留一个拆过的空位');
    assert.ok(app.variables.get(CAPABILITY_VARIABLE)().includes('本会话能力'), '装回来的变量要能求值');

    const state = rebuildState();
    assert.equal(state.ok, false, '失败要记进状态');
    assert.equal(state.restored, true, '这一次是「装回旧的」而不是降级态');
    assert.equal(state.degraded, false);
    assert.equal(state.phase, '能力变量', `失败阶段要报得准，实际 ${JSON.stringify(state.phase)}`);
    assert.match(state.error, /注入的失败/, '原因照实记（不吞）');
    assert.equal(state.droppedProxy, '', '这一次没动代理，不该报代理被丢');

    const failures = app.warned.filter((line) => line.includes('设置重建失败'));
    assert.equal(failures.length, 1, `失败必须说出来，实际记了 ${failures.length} 条`);
    assert.match(failures[0], /设置重建失败（能力变量）/);
    assert.match(failures[0], /已把改动前的配置整份装回/);
    assert.ok(describeRebuildState().includes('已装回改动前的配置'), 'office_help 那一段要能说清退路');

    // 故障消失后再广播一次：这次要真的按新配置装（说明回滚没有把状态弄坏）。
    app.disarm('variables');
    app.broadcast('dsh-office-mode');
    assert.equal(app.tools.size, TOOL_COUNT - 1, '候选配置这次装成了：office_memory 应当不在工具面里');
    assert.ok(!app.tools.has('office_memory'));
    assert.equal(app.variables.size, 2);
    assert.deepEqual(rebuildState(), HEALTHY, '恢复成功后状态要回到正常态');
    assert.equal(describeRebuildState(), '', '正常态要回到零字节');
});

// ── 3. 回滚也失败 → 降级态，但**诊断通道真的活着** ──────────────────────────

await check('回滚也失败（工具面）：降级态里诊断工具与能力变量都还在，用户看得见、装配也过得去', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    apply(app.ctx, {});

    // 让「其余工具面」里的 office_run **每次都**炸：候选那一份装不成，装回旧的也装不成。
    // 按 installAll 的顺序，两样硬依赖（能力变量、诊断工具 office_help）排在它前面。
    app.armAlways('tools', (name) => name === 'office_run');
    app.broadcast('dsh-office-mode');

    const state = rebuildState();
    assert.equal(state.ok, false);
    assert.equal(state.restored, false, '回滚也失败时要照实标出来');
    assert.equal(state.degraded, true, '这才是降级态：只剩部分能力');
    assert.equal(state.phase, '工具面', '先报候选那一次失败在哪一步');
    assert.equal(state.rollbackPhase, '工具面', '再报回滚那一次失败在哪一步');
    assert.match(state.rollbackError, /注入的失败/);

    assert.equal(app.variables.size, 2, '能力变量排在工具面之前：它必须还在（缺了每次装配都会抛）');
    assert.ok(app.variables.get(CAPABILITY_VARIABLE)().includes('本会话能力'), '而且还能求值');
    assert.equal(app.tools.size, 1, `只剩诊断工具，实际 ${app.tools.size} 个`);
    assert.ok(app.tools.has('office_help'), 'office_help 要在册 —— 它是「只剩部分能力」唯一能被看见的通道');

    const failures = app.warned.filter((line) => line.includes('装回旧配置也失败'));
    assert.equal(failures.length, 1, '降级态必须说出来，实际记了 ' + failures.length + ' 条');
    assert.match(failures[0], /当前只剩部分能力/);

    // **走注册面**读诊断（不是另造一份 office_help）：这条断言正是复核发现 1 的形态。
    const registered = app.getTool('office_help');
    assert.ok(registered, '注册表里要有 office_help');
    const out = await registered.execute({ topic: 'word' }, { agent: { session: { header: { cwd: process.cwd() } } } });
    assert.ok(out.text.includes('设置重建'), 'office_help 要带这一行，否则用户手上没有任何一处能看出降级');
    assert.ok(out.text.includes('装回旧配置也失败'), '降级态的退路（重挂插件 / 重启）要写出来');
    assert.ok(out.text.includes('只剩部分能力'), '要说清当前是什么状态');
    assert.ok(!describeRebuildState().includes('失败点落在诊断工具或能力变量自己身上'),
        '失败点不在那两样身上时，不该多那句条件句');

    // 故障消失 → 再广播一次要能自己爬出来（降级态不是终点）。
    app.disarm('tools');
    app.broadcast('dsh-office-mode');
    assert.deepEqual(rebuildState(), HEALTHY, '恢复成功要回到正常态');
    assert.equal(app.variables.size, 2);
    assert.equal(app.tools.size, TOOL_COUNT);
    assert.equal(describeRebuildState(), '');
    assert.ok(!(await helpText('word')).includes('【运行期实况】'), '正常态时这一行必须完全不出现（零字节成本）');
});

// ── 4. 硬依赖自己失败 → 照实说「可能什么都看不见」 ───────────────────────────

await check('回滚也失败（能力变量自己）：残留为空，状态照实写「诊断通道可能也不在」', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    apply(app.ctx, {});

    app.armAlways('variables', (name) => name === CAPABILITY_VARIABLE);
    app.broadcast('dsh-office-mode');

    const state = rebuildState();
    assert.equal(state.degraded, true);
    assert.equal(state.phase, '能力变量');
    assert.equal(state.rollbackPhase, '能力变量');
    assert.equal(app.variables.size, 0, '两次都没装成，变量确实缺着（状态没有粉饰）');
    assert.equal(app.tools.size, 0, '失败点在诊断工具之前，所以连 office_help 都没装上 —— 这是最坏形态');
    assert.match(describeRebuildState(), /失败点落在诊断工具或能力变量自己身上/,
        '这一支要照实说「诊断通道可能也不在」，不能宣称一定有');
    assert.equal(app.warned.filter((line) => line.includes('装回旧配置也失败')).length, 1);

    app.disarm('variables');
    app.broadcast('dsh-office-mode');
    assert.deepEqual(rebuildState(), HEALTHY, '恢复成功要回到正常态');
    assert.equal(app.tools.size, TOOL_COUNT);
});

// ── 5. 工具面（第一步之后）失败也要回滚 ─────────────────────────────────────

await check('工具面重新注册失败（一次性的）：整份装回旧配置，工具名与顺序都要一致', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    apply(app.ctx, {});
    const names = [...app.tools.keys()];

    // 其余工具面的第一件（office_run）注册失败：候选那一份只装上了 office_help。
    app.armOnce('tools', (name) => name === 'office_run');
    app.broadcast('dsh-office-mode');

    assert.equal(app.tools.size, TOOL_COUNT, '整份工具面要回来');
    assert.deepEqual([...app.tools.keys()], names, '工具名与顺序都要与激活时一致');
    const state = rebuildState();
    assert.equal(state.phase, '工具面');
    assert.equal(state.restored, true);
    assert.match(app.warned[0], /设置重建失败（工具面）/);
    assert.equal(app.variables.size, 2, '能力变量也要跟着回来');
});

// ── 6. 工具被设置页全部关掉：照常能装、能重建 ——「没有诊断工具」是设置的选择，不是降级 ──

await check('工具全关（含 office_help）时照常激活与重建（这里不该把合法的设置当成插件 bug）', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    const off = Object.fromEntries(buildTools(resolveConfig({})).map((item) => [item.name, false]));
    const config = { tools: off };
    apply(app.ctx, config);
    assert.equal(app.tools.size, 0, '工具全关时工具面就是空的');
    assert.equal(app.variables.size, 2, '能力变量照旧要注册（它跟工具开关无关）');

    app.broadcast('dsh-office-mode');
    assert.deepEqual(app.warned, [], `重建不该报错，实际：${app.warned.join(' | ')}`);
    assert.deepEqual(rebuildState(), HEALTHY, '关掉工具不是失败，状态要留在正常态');

    // 只关 office_help（其余照开）：同样不该被当成 bug。
    const only = Object.fromEntries(buildTools(resolveConfig({})).map((item) => [item.name, item.name !== 'office_help']));
    const second = fakeCtx();
    apply(second.ctx, { tools: only });
    assert.equal(second.tools.size, TOOL_COUNT - 1);
    assert.ok(!second.tools.has('office_help'), '用户关掉的就是不注册');
    assert.deepEqual(second.warned, []);
});

// ── 7. 重入广播不许把事务打断 ──────────────────────────────────────────────
await check('事务进行中又收到一次广播（宿主读设置投影也会广播）：合并成最后一次，不留混合状态', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    const config = {};
    apply(app.ctx, config);

    // 第一次广播带 office_web_search:false；处理它的过程中（工具注册里）重入第二次广播，
    // 而第二次的 config 又多关了 office_web_fetch —— 最终必须按**最后**那一份装。
    // （这两个开关都不影响提示段落：`office_memory` 关掉会连带撤掉记忆说明段，那是另一条判据。）
    config.tools = { office_web_search: false };
    app.onceReenter(() => {
        config.tools = { office_web_search: false, office_web_fetch: false };
        app.broadcast('dsh-office-mode');
    });
    app.broadcast('dsh-office-mode');

    assert.deepEqual(app.warned, [], `重入不该被当成失败，实际：${app.warned.join(' | ')}`);
    assert.deepEqual(rebuildState(), HEALTHY, '重入之后状态必须仍是正常态');
    assert.equal(app.variables.size, 2, '重入不许把变量注册搞成两份或零份');
    assert.equal(app.sections.size, 1, '重入不许把提示段落搞成两份或零份');
    assert.equal(app.tools.size, TOOL_COUNT - 2, `最终要按最后一次配置装（少两个工具），实际 ${app.tools.size}`);
    assert.ok(!app.tools.has('office_web_search') && !app.tools.has('office_web_fetch'), '两次广播的改动都要落地');
});

// ── 8. 代理：不进事务，但「被丢掉」必须说出来 ───────────────────────────────

await check('代理不进重建事务：失败时按旧配置走，且「候选里的代理改动没生效」要报出来（含打码）', async () => {
    const { apply } = await import('../src/index.js');
    resetProxyStateForTest();
    resetRebuildStateForTest();
    const app = fakeCtx();
    const config = {};
    apply(app.ctx, config);

    // 用户填了个带凭据的代理，同时注入工具面失败。
    config.search = { proxy: 'http://user:sup3rsecret@proxy.example.com:8080' };
    app.armOnce('tools', (name) => name === 'office_run');
    app.broadcast('dsh-office-mode');

    const state = rebuildState();
    assert.equal(state.restored, true, '这一轮是「装回旧的」');
    assert.equal(state.droppedProxy, 'http://***:***@proxy.example.com:8080', '被丢掉的代理改动要记进状态并打码');
    assert.ok(!state.droppedProxy.includes('sup3rsecret'), '状态里不许出现口令');
    assert.ok(!proxyStatus().note.includes('proxy.example.com'),
        '事务失败时不该按候选配置去切出口 —— 否则会留下「工具面是旧的、出口是新的」这种组合');
    assert.ok(app.warned.some((line) => line.includes('代理改动没有生效')), '日志要点名这件容易漏掉的事');

    const search = await helpText('search');
    assert.ok(search.includes('本次设置里的代理改动'), 'office_help 也要能看见它（用户不一定会翻日志）');
    assert.ok(search.includes('http://***:***@proxy.example.com:8080'), '报出来的是打码后的地址');
    assert.ok(!search.includes('sup3rsecret'), '口令不许进 office_help');

    // 故障消失：事务提交，代理才装（装得成装不成是另一件事，这一轮只要它被算作生效）。
    app.disarm('tools');
    app.broadcast('dsh-office-mode');
    assert.deepEqual(rebuildState(), HEALTHY, '代理装不成不该让重建事务失败');
    assert.ok(app.warned.some((line) => line.includes('代理设置未生效')) || proxyStatus().mode !== 'none',
        '提交之后才轮到代理：要么装上，要么照实记一条「未生效」');
    assert.ok(!proxyStatus().note.includes('sup3rsecret'), '代理实况里不许回显口令');
    resetProxyStateForTest();
});

// ── 9. 激活期失败：响亮 + 不留半份注册（回归护栏） ──────────────────────────

await check('激活期注册失败要抛出去（宿主记一条加载失败），但不留半份注册在注册表里', async () => {
    const { apply } = await import('../src/index.js');
    resetRebuildStateForTest();
    const app = fakeCtx();
    app.armAlways('tools', (name) => name === 'office_run');
    assert.throws(() => apply(app.ctx, {}), /注入的失败/,
        '激活失败要抛给宿主（那时没有「旧的那份」可回滚，静默半份能力更糟）');
    assert.equal(app.tools.size, 0, '已经装上的 office_help 也要收干净：宿主重试加载时同名注册会撞车');
    assert.equal(app.variables.size, 0, '变量那一步装上了，但它也在 disposers 里，要一起收掉');
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok && item.note !== undefined ? `  → ${item.note}` : ''}`);
    if (!item.ok) console.log(`       | ${item.error?.message ?? item.error}`);
}
console.log(`rebuild: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
