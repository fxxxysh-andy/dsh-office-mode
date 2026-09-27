/**
 * 设置页（浏览器半侧）的单元测试。
 *
 * 客户端 bundle 是给浏览器用的 CJS 产物（window.__ModuleLoader__ + require），
 * 没法在 Node 里直接 import。这里用一套最小的 React 替身把它跑起来，
 * 验证：bundle 形状、slot 注册、设置读写接线、以及交互控件的行为。
 *
 * 跑法：node test/client.mjs
 */
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';

import { resolveHostPath } from './host-modules.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const BUNDLE = join(here, '..', 'lib', 'client.js');

/** 从宿主的安装位置解析 schemastery（与 settings.mjs 同一套锚点，见 host-modules.mjs）。 */
async function loadSchemastery() {
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

/**
 * 解析 schema 并把 volatile 引用拆成纯值。
 *
 * 0.1.7 起可编辑字段都是 volatile，解析结果是 Volatile 引用（一个带 get/write
 * 的对象）。直接拿解析值枚举叶子会得到 "documents.defaultTheme.get" 这种假路径，
 * 下面「界面覆盖服务端每个参数」的核对就会整片错位。
 */
async function resolvePlain(z, schema) {
    const cosmokit = await loadCosmokit();
    const unwrap = (value) => {
        if (cosmokit !== undefined && cosmokit.isVolatile(value)) return unwrap(value.get());
        if (Array.isArray(value)) return value.map(unwrap);
        if (value !== null && typeof value === 'object') {
            const out = {};
            for (const [key, child] of Object.entries(value)) out[key] = unwrap(child);
            return out;
        }
        return value;
    };
    return unwrap(schema({}));
}

const results = [];
/** 最近一次 loadBundle 造出的 React 替身，供 render() 重置 hook 游标。 */
let currentFake = null;
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/**
 * 最小的 React 替身。
 *
 * 只实现 bundle 真正用到的部分：createElement 产出普通对象树，
 * hooks 用数组按调用顺序存值（和 React 的规则一致）。
 * 这样不必引入 react 依赖，也能断言「树里有什么控件」。
 */
function makeFakeReact() {
    // hooks 按「组件实例 + 调用序号」存：同一个组件重渲染时复用上次的值，
    // 不同组件各有一份 —— 和 React 的规则一致。
    //
    // 「实例」是按**同一轮渲染里的出现次序**算的（NumberField#1、NumberField#2…）：
    // 只按组件名做键时，同名的第二个实例会复用第一个的状态，页面上多几个数字
    // 输入框就会互相串值（2026-09-23 加入记忆组的多行数字控件时踩到）。
    const store = new Map();
    let currentComponent = null;
    let cursor = 0;
    /** 本轮渲染里每个组件名已经出现过几次。 */
    let occurrences = new Map();
    /** 有 setState 发生就置脏，render() 据此再渲染一轮。 */
    let dirty = false;

    const resetCursor = () => { cursor = 0; };
    const resetInstances = () => { occurrences = new Map(); };

    const React = {
        createElement: (type, props, ...children) => ({
            $typeof: true,
            type,
            props: Object.assign({}, props || {}, children.length > 0 ? { children } : {}),
        }),
        useState: (initial) => {
            const key = currentComponent + ':' + cursor;
            cursor += 1;
            if (!store.has(key)) store.set(key, { value: initial, rerender: null });
            const slot = store.get(key);
            const setter = (next) => {
                const value = typeof next === "function" ? next(slot.value) : next;
                if (Object.is(value, slot.value)) return;
                slot.value = value;
                // 置脏：render() 见到脏标记就再渲染一轮，直到收敛。
                dirty = true;
            };
            return [slot.value, setter];
        },
        useEffect: (fn) => {
            // 提交后执行一次（同一组件同一序号只跑一次，和 React 的依赖数组行为接近）。
            const key = currentComponent + ':effect:' + cursor;
            cursor += 1;
            if (store.has(key)) return;
            store.set(key, { value: true });
            try {
                const cleanup = fn();
                if (typeof cleanup === 'function') store.set(key + ':cleanup', { value: cleanup });
            } catch (error) {
                store.set(key + ':error', { value: error });
            }
        },
        useRef: (initial) => {
            const key = currentComponent + ':ref:' + cursor;
            cursor += 1;
            if (!store.has(key)) store.set(key, { value: { current: initial } });
            return store.get(key).value;
        },
    };

    return {
        React,
        resetCursor,
        resetInstances,
        store,
        isDirty: () => dirty,
        clearDirty: () => { dirty = false; },
        /** 标记接下来要渲染哪个组件实例（决定 hooks 的归属）。 */
        setComponent(name) {
            const seen = (occurrences.get(name) ?? 0) + 1;
            occurrences.set(name, seen);
            currentComponent = `${name}#${seen}`;
        },
    };
}

/**
 * 面板取数用的假端点响应：一份最小但字段齐全的快照。
 *
 * 字段名照抄 src/view.js 的冻结契约 —— 这里如果比服务端少一个字段，
 * 面板里对应的渲染分支就永远测不到。
 */
function makeSnapshot(overrides = {}) {
    const base = {
        ok: true,
        generatedAt: '2026-09-24T00:00:00.000Z',
        cwd: 'D:/proj/temp1',
        label: 'temp1',
        query: '',
        workspaces: [{ cwd: 'D:/proj/temp1', label: 'temp1', seenAt: '2026-09-24T00:00:00.000Z' }],
        config: {
            scope: 'workspace',
            userScope: 'memory',
            layers: { hot: true, ledger: true, archive: true },
            links: true,
            autoCapture: true,
            quality: {
                policy: 'strict-v1',
                lowScoreThreshold: 0.25,
                highScoreThreshold: 0.6,
                candidateMultiplier: 3,
                maxMediumResults: 4,
                maxUnknownResults: 2,
            },
            quota: { recallPerTurn: 1, recallRefinePerTurn: 1, relatedPerTurn: 1 },
            limits: { userLimitBytes: 4096, projectLimitBytes: 10240, ledgerLimit: 500, archiveKeep: 60 },
        },
        stores: [{ id: 'workspace', dir: '.office/memory', hot: 1, ledger: 1, archive: 1, archiveFiles: 1, links: 1 }],
        // counts 有两个口径（见 src/view.js 的注释）：hot/ledger/archive/links/entities 是
        // 「列表里现在有几条」（过滤并截断之后），ledgerTotal/archiveFiles/archiveItems 是
        // 容量条用的未过滤总数。默认夹具里两者相等，专门的口径用例会造出差异。
        counts: {
            hot: 1, ledger: 1, archive: 1, links: 1, entities: 1,
            ledgerTotal: 1, archiveFiles: 1, archiveItems: 1,
        },
        hot: [{
            id: 'm-1', target: 'user', importance: 'critical', origin: 'global',
            updatedAt: '2026-09-24T00:00:00.000Z', entities: [], tags: [], content: '偏好：书面化表达',
        }],
        ledger: [{
            id: 'L-1', at: '2026-09-24T00:00:00.000Z', path: '季度汇报.pptx', format: 'ppt',
            theme: 'plain', purpose: '季度汇报', outline: ['封面'], origin: 'workspace',
        }],
        archive: [{
            id: '', month: '2026-09', kind: 'mnemon', at: '2026-09-20T00:00:00.000Z', origin: 'workspace',
            importance: 'normal', category: 'fact', entities: ['系统A'], tags: ['迁移'], text: '长期记忆一条',
        }],
        links: [{
            id: 'K-1', sourceId: 'm-1', targetId: 'L-1', kind: 'related', note: '',
            at: '2026-09-24T00:00:00.000Z', origin: 'workspace',
        }],
        entities: [{
            name: '系统A', count: 2,
            refs: [{ id: 'm-1', layer: 'hot', origin: 'workspace', text: '偏好：书面化表达' }],
        }],
    };
    return Object.assign(base, overrides);
}

/** 把 bundle 在一个受控环境里物化，返回它的导出与注册记录。 */
function loadBundle(options = {}) {
    const source = readFileSync(BUNDLE, 'utf8');
    let entry = null;
    const win = { __ModuleLoader__: { load: (value) => { entry = value; } } };
    // 面板里的两处浏览器能力：剪贴板（写入指令的「复制」）与 matchMedia
    // （reduced-motion 下的平滑滚动）。两者都从 window 上读，所以按需注入。
    if (options.navigator !== undefined) win.navigator = options.navigator;
    if (options.matchMedia !== undefined) win.matchMedia = options.matchMedia;
    const fake = makeFakeReact();
    const require = (name) => {
        if (name === 'react') return fake.React;
        throw new Error(`bundle 请求了未声明的模块：${name}`);
    };
    // 面板用的是裸 fetch。把它作为包装函数的形参注入，测试才能换成替身：
    // 不注入的话会落到 Node 的全局 fetch，去打一个相对 URL（必然失败），
    // 测试就变成「跟网络环境比」而不是「跟用例造的数据比」。
    // 注意用 'fetch' in options 判断，而不是 ??：显式传 undefined 表示
    // 「这个环境没有 fetch」，那是一条要测的降级路径。
    const defaultFetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    const fetchImpl = 'fetch' in options ? options.fetch : defaultFetch;
    new Function('window', 'require', 'console', 'fetch', source)(win, require, console, fetchImpl);
    assert.ok(entry !== null, 'bundle 没有调用 window.__ModuleLoader__.load');
    const mod = entry.factory(require);
    const registrations = [];
    const injected = [];
    const target = {
        effect: (fn) => { return fn(); },
        // 宿主客户端上下文的 ctx.inject：只在服务就绪时才回调（缺服务就永不回调，
        // 而不是抛错）。回调拿到的是带该服务的子上下文。
        // configForms 与 layout 都是**可选**服务：没传就当作「这个部署没有」，
        // 于是「拿不到服务时的降级路径」也能被测到。
        inject: (services, callback) => {
            const list = Array.isArray(services) ? services : [services];
            const sub = {};
            for (const service of list) {
                if (service === 'configForms') {
                    if (options.forms === undefined) return;
                    sub.configForms = options.forms;
                    continue;
                }
                if (service === 'layout') {
                    if (options.layout === undefined) return;
                    sub.layout = options.layout;
                    continue;
                }
                return;
            }
            callback(sub);
        },
        slots: {
            // 不再钉死 slot 名：插件现在往五种 slot 里注册（两个设置页 + 侧栏入口
            // + 主面板 + 会话内两处）。这里只记名字，逐项断言交给各条用例。
            inject: (slotName, fn) => {
                assert.equal(typeof slotName, 'string');
                assert.ok(slotName.length > 0, 'slot 名不能为空');
                injected.push(slotName);
                fn();
            },
            register: (opts, render) => { registrations.push({ opts, render }); return () => {}; },
        },
    };
    // DSH 客户端上下文对**没写进 inject 的服务属性**是抛错而不是返回 undefined
    // （0.1.7 实测：直接读 ctx.configForms 会让整个 entry 变 failed，页面显示
    // 「Failed to load plugins / dsh-office-mode: failed」，而宿主与终端都不报错）。
    // 这个 Proxy 把那套严格语义搬进测试，让「偷偷直接读服务」这类回归当场暴露。
    const TOLERATED = new Set(['then', 'toJSON', 'constructor', 'toString', 'valueOf', 'inspect']);
    const ctx = new Proxy(target, {
        get(object, property) {
            if (typeof property === 'symbol' || property in object || TOLERATED.has(property)) {
                return Reflect.get(object, property);
            }
            throw new Error(`Service "${String(property)}" is not available (not declared in inject)`);
        },
        has: (object, property) => property in object,
    });
    currentFake = fake;
    return { entry, mod, ctx, registrations, injected, fake };
}

/**
 * 把 slot 注册产出的元素树渲染成「已展开的」普通对象树。
 *
 * slot 的 render 返回的是 <Component {...props} />（一个元素），真正渲染它的是
 * 外壳。测试里没有外壳，所以这里自己做一次最小渲染：遇到函数组件就调用它，
 * 直到整棵树只剩宿主元素（div / input / select …）。
 *
 * 每次渲染前重置 hook 游标，模拟 React「按调用顺序匹配 hooks」的规则；
 * effect 同步执行一次，这样绑定 settingsScope 的接线才会真的跑起来。
 */
function render(registration, props) {
    // 用一个可重入的渲染循环处理 setState：组件第一次渲染 → effect 里 setState
    // → 再渲染一次，直到不再变化。真实 React 就是这么收敛的。
    let tree = null;
    const expand = (node) => {
        if (node === null || node === undefined || typeof node === 'boolean') return node;
        if (Array.isArray(node)) return node.map(expand);
        if (typeof node !== 'object') return node;
        const type = node.type;
        if (typeof type === 'function') {
            const name = type.name || 'anon';
            currentFake.setComponent(name);
            currentFake.resetCursor();
            const rendered = expand(type(Object.assign({}, node.props)));
            // 记下来源组件的名字与它收到的 props：展开后这些信息就丢了，
            // 而测试需要按 props.label 找到某一行、再取它渲染出的控件。
            if (rendered !== null && typeof rendered === 'object' && !Array.isArray(rendered)) {
                rendered.source = { name, props: node.props };
            }
            return rendered;
        }
        const props = Object.assign({}, node.props);
        if (props.children !== undefined) props.children = expand(props.children);
        return { type, props, source: null };
    };
    // 渲染 → effect 里可能 setState → 再渲染，直到没有新的状态变化。
    for (let pass = 0; pass < 10; pass += 1) {
        currentFake.clearDirty();
        currentFake.resetInstances();
        // 传 props 进去：会话内的两个 slot（回合条 / 存入记忆按钮）靠 owner props
        // 决定画不画，不给 props 就只能测到「读不到时返回 null」那一半。
        tree = expand(registration.render(props));
        if (!currentFake.isDirty()) break;
    }
    return tree;
}

/** 递归收集渲染树里的所有元素。 */
function walk(node, out = []) {
    if (node === null || node === undefined || node === false) return out;
    if (Array.isArray(node)) { for (const item of node) walk(item, out); return out; }
    if (typeof node !== 'object') return out;
    out.push(node);
    const kids = node.props ? node.props.children : undefined;
    if (kids !== undefined) walk(kids, out);
    return out;
}

// ── bundle 形状 ───────────────────────────────────────────────────────────

await check('bundle 注册成模块 id dsh-office-mode，并导出 apply/inject', () => {
    const { entry, mod } = loadBundle();
    assert.equal(entry.id, 'dsh-office-mode');
    assert.equal(typeof mod.apply, 'function');
    assert.deepEqual([...mod.inject], ['slots']);
});

await check('bundle 只请求声明过的模块（react）', () => {
    // require 遇到未声明的模块会抛错；能跑完就说明没有越界请求。
    assert.doesNotThrow(() => loadBundle());
});

await check('注册成两个 settings.section 一级导航项：办公模式 + 记忆系统', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const sections = registrations.filter((item) => item.opts.name === 'settings.section');
    assert.equal(sections.length, 2, '两页共用同一个命名空间，但不是同一格导航');

    const office = officeReg(registrations);
    assert.ok(office, '应有「办公模式」这一格');
    assert.equal(office.opts.name, 'settings.section');
    assert.equal(office.opts.id, 'dsh-office-mode');
    assert.equal(typeof office.opts.order, 'number');
    assert.equal(typeof office.opts.label, 'function', '导航项标题要用 thunk（支持本地化重取）');
    assert.equal(typeof office.render, 'function');

    const memory = memoryReg(registrations);
    assert.ok(memory, '应有「记忆系统」这一格（mnemon 关掉后由本插件接住）');
    assert.equal(memory.opts.name, 'settings.section');
    assert.equal(memory.opts.id, 'dsh-office-memory');
    assert.notEqual(memory.opts.id, office.opts.id, '两格不能用同一个 id：会被当成同一格');
    assert.equal(typeof memory.opts.order, 'number');
    assert.equal(typeof memory.opts.label, 'function');
    assert.ok(memory.opts.order < office.opts.order, '记忆系统排在办公模式之前');
});

await check('没有 slots 服务时静默不注册，不抛错', () => {
    const { mod } = loadBundle();
    assert.doesNotThrow(() => mod.apply({}));
    assert.doesNotThrow(() => mod.apply(null));
});

await check('inject 只声明 slots —— 可选服务不进注入列表，否则缺服务时整页不注册', () => {
    // 这条是踩过坑的：0.1.7 把 settingsScope 整个删了，而它还在 inject 里，
    // 于是 apply 永远不跑 —— 记忆面板与两个设置页一起消失。
    const { mod } = loadBundle();
    assert.ok(mod.inject.includes('slots'), 'inject 必须含 slots（注册导航项与面板）');
    assert.ok(!mod.inject.includes('settingsScope'), 'settingsScope 在 0.1.7 已不存在，不能再声明');
    assert.ok(!mod.inject.includes('remote'), 'remote 不是本插件必需的，不进注入列表');
});

await check('没有 configForms 时不注册出空白面板，而是提示缺少服务', () => {
    const { mod, ctx, registrations } = loadBundle({ forms: undefined });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const texts = walk(tree).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((x) => x.includes('configForms')), '应提示缺少 configForms 服务');
});

await check('不直接触碰没注入的服务：configForms 缺失时 apply 也不抛错，两页仍注册', () => {
    // 这条是踩过坑的，也是本轮真正的启动杀手：DSH 的客户端上下文里，读一个没写进
    // inject 的服务属性会**当场抛错**，整个 entry 变成 failed —— 页面显示
    // 「Failed to load plugins / dsh-office-mode: failed」，而宿主侧一切正常、
    // 连终端都不报这个错。loadBundle 的 ctx 是个对未声明属性抛错的 Proxy，
    // 所以一旦代码里再出现裸读 ctx.<service>，这条会立刻红。
    const { mod, ctx, registrations } = loadBundle({ forms: undefined });
    assert.doesNotThrow(() => mod.apply(ctx), 'apply 不该读未注入的服务属性');
    const sections = registrations.filter((item) => item.opts.name === 'settings.section');
    assert.equal(sections.length, 2, '设置页拿不到服务也要照常注册（面板自己显示降级提示）');
    const panel = registrations.filter((item) => item.opts.name === 'sidebar.panellist');
    assert.equal(panel.length, 1, '记忆面板与设置服务无关，必须照常注册');
});

// ── 设置读写接线 ──────────────────────────────────────────────────────────

/**
 * 造一个假的 configForms 服务，记录取件与写入。
 *
 * 0.1.7 起设置传输是 ctx.configForms（dsh-client-ui-settings 提供的基础服务）：
 * 按「宿主插件行 id」取一份 ConfigForm，写入走
 * mutate([{ op: 'set', path: [...], value }])。这个替身把每次写入压成
 * [点号路径, 值]，好让下面几十条断言继续按原样读。
 */
function makeScope(value, options = {}) {
    const writes = [];
    const snapshot = {
        status: options.status || 'ready',
        value,
        base: undefined,
        user: undefined,
        revision: 1,
        writable: options.writable !== false,
        mode: 'host',
    };
    const form = {
        getSnapshot: () => snapshot,
        subscribe: () => () => {},
        set: (field, next) => { writes.push([field, next]); return Promise.resolve(true); },
        unset: () => Promise.resolve(true),
        mutate: (ops) => {
            for (const op of ops) writes.push([op.path.join('.'), op.value]);
            return Promise.resolve(options.refuse !== true);
        },
    };
    const forms = {
        get: (entryId) => { forms.bound = { entryId }; return form; },
        describe: () => ({
            getSnapshot: () => ({ status: 'ready', view: { namespaces: [], writable: true, hasDocument: false }, error: null }),
            subscribe: () => () => {},
            ensure: () => Promise.resolve(),
        }),
        whileServed: (namespaces, register) => {
            const off = register(new Set(namespaces));
            return typeof off === 'function' ? off : () => {};
        },
    };
    return { forms, writes, snapshot };
}

const SAMPLE_VALUE = {
    tools: {
        office_help: true,
        office_run: false,
        office_memory: true,
        office_search_brief: true,
        office_search_dispatch: true,
        office_parse_findings: true,
    },
    memory: {
        enabled: true,
        dir: '.office/memory',
        autoLedger: true,
        promptHint: true,
        userLimitBytes: 4096,
        projectLimitBytes: 10240,
        ledgerLimit: 500,
        archiveKeep: 60,
        // 三个嵌套组照服务端 schema 的形状给：界面写回用的是点号路径，
        // 扁平化的假数据会让「嵌套有没有写对」这件事测不出来。
        // archive 故意给 false，用来断言「关掉的层显示为未勾选」。
        layers: { hot: true, ledger: true, archive: false },
        scope: 'workspace',
        globalDir: '',
        userScope: 'memory',
        links: true,
        autoCapture: true,
        recallQuality: {
            policy: 'strict-v1',
            lowScoreThreshold: 0.25,
            highScoreThreshold: 0.6,
            candidateMultiplier: 3,
            maxMediumResults: 4,
            maxUnknownResults: 2,
        },
        quota: { recallPerTurn: 1, recallRefinePerTurn: 1, relatedPerTurn: 1 },
    },
    subagentModel: { mode: 'inherit', provider: '', model: '' },
    search: {
        subagentTools: ["web_search", "advanced_search", "platform_search", "web_fetch", "read", "read_image", "write"],
        maxParallel: 4,
        resultLimit: 40,
        maxChannels: 12,
        outputDir: '.office/search',
        requireCrossSource: true,
        fallbackOnPlatformError: true,
    },
    documents: {
        defaultTheme: 'plain',
        scriptTimeoutMs: 60000,
        maxScriptChars: 400000,
        cacheDir: '.office/cache',
        keepCache: false,
    },
    av: {
        enabled: true,
        ffmpegPath: '',
        ffprobePath: '',
        modelDir: '',
        language: 'zh',
        chunkSeconds: 120,
        maxSeconds: 3600,
        timeoutMs: 300000,
        precision: 'int8',
        threads: 2,
        frames: 6,
        maxFrames: 12,
        frameFormat: 'jpg',
        frameWidth: 0,
    },
};

await check('绑定到正确的命名空间', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    render(officeReg(registrations));
    assert.equal(scope.forms.bound.entryId, 'dsh-office-mode');
});

await check('没有 configForms 时给出可读提示，不白屏', () => {
    const { mod, ctx, registrations } = loadBundle({ forms: undefined });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const texts = walk(tree).map((node) => node.props && node.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('configForms')), '应提示缺少 configForms 服务');
});

await check('命名空间不可用时给出可读提示', () => {
    const scope = makeScope(undefined, { status: 'unavailable' });
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const texts = walk(tree).map((node) => node.props && node.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('命名空间')), '应提示命名空间不可用');
});

// ── 交互控件 ──────────────────────────────────────────────────────────────

/**
 * 取某一格导航项的注册记录。
 *
 * 两个 section（办公模式 / 记忆系统）共用同一个设置命名空间，所以不能再按
 * 「第 0 个注册」拿面板 —— 顺序变了就会渲染错的那一页。
 */
function officeReg(registrations) {
    return registrations.find((item) => item.opts.id === 'dsh-office-mode');
}

function memoryReg(registrations) {
    return registrations.find((item) => item.opts.id === 'dsh-office-memory');
}

/**
 * 按 slot 名 + 标识取一条注册记录。
 *
 * list slot 的标识是 `id`，keyed slot（main）的是 `key` —— 两种都认，
 * 免得每处都写一遍。
 */
function regOf(registrations, name, key) {
    return registrations.find((item) => item.opts.name === name && (item.opts.id === key || item.opts.key === key));
}

/** 只渲染记忆系统那一页。 */
function renderMemory(registrations) {
    return render(memoryReg(registrations));
}

/**
 * 在渲染树里按标题文本找到对应的一行。
 *
 * 找的是 Row 组件「渲染出来的那个 div」：展开时在它身上留了 source 标记，
 * 里面存着 Row 收到的 label。直接按展开后的 props.label 找是找不到的 ——
 * 展开后 props 已经换成 div 自己的了。
 */
function findByLabel(tree, labelText) {    return walk(tree).find((node) => node.source && node.source.name === 'Row' && node.source.props.label === labelText);
}

/**
 * 造一个最小的 DOM 事件对象。
 *
 * 控件的事件处理器读的是 event.target.value / event.target.checked（浏览器
 * 真实传进来的形状），所以测试必须给同样的东西，不能直接传裸值。
 */
function eventOf(props) {
    return { target: Object.assign({}, props), currentTarget: Object.assign({}, props) };
}

/**
 * 从一行里取出主控件。
 *
 * Row 右侧可能包着组件（Toggle 渲染出 input、NumberField 渲染出 input），
 * 所以优先找宿主表单元素；找不到再退回「第一个带事件回调的元素」。
 */
function controlOf(row) {
    if (!row) return undefined;
    // Row 渲染成 [左侧文字列, 右侧控件列]；只从右侧那一列里找，
    // 否则会先撞上左侧说明文字里的元素。
    const columns = (row.props && row.props.children) || [];
    const list = Array.isArray(columns) ? columns : [columns];
    const controlColumn = list.length > 1 ? list[list.length - 1] : list[0];
    const nodes = walk(controlColumn);
    const form = nodes.find((node) => (node.type === 'input' || node.type === 'select' || node.type === 'textarea') && node.props);
    if (form !== undefined) return form;
    return nodes.find((node) => node.props && (typeof node.props.onChange === 'function' || typeof node.props.onClick === 'function'));
}

await check('工具开关渲染出勾选态，并能写回 tools.<name>', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const row = findByLabel(tree, 'office_run');
    assert.ok(row, '应渲染 office_run 这一行');
    const toggle = controlOf(row);
    assert.equal(toggle.type, 'input');
    assert.equal(toggle.props.type, 'checkbox');
    assert.equal(toggle.props.checked, false, '配置里 office_run=false，应显示未勾选');
    toggle.props.onChange(eventOf({ checked: true }));
    assert.deepEqual(scope.writes[0], ['tools.office_run', true]);
});

await check('默认开启的工具显示为勾选', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const toggle = controlOf(findByLabel(tree, 'office_help'));
    assert.equal(toggle.props.checked, true);
});

await check('并发数用数字输入，并在提交时写回 search.maxParallel', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const row = findByLabel(tree, '并发子代理数');
    assert.ok(row, '应渲染并发数这一行');
    const input = controlOf(row);
    assert.equal(input.type, 'input');
    assert.equal(input.props.type, 'number');
    assert.equal(String(input.props.value), '4', '数字输入的值（草稿是字符串）');
    assert.equal(input.props.min, 1);
    assert.equal(input.props.max, 8);
});

await check('主题用下拉，选项覆盖服务端支持的主题', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const select = controlOf(findByLabel(tree, '默认主题'));
    assert.equal(select.type, 'select');
    assert.equal(select.props.value, 'plain');
    // React 里 option 是 select 的子元素（不是 options 属性）。
    const options = walk(select.props.children).filter((node) => node.type === 'option');
    assert.equal(options.length, 7, '应渲染全部七个主题');
    const ids = options.map((node) => node.props.value);
    assert.ok(ids.includes('plain') && ids.includes('business') && ids.includes('forest'), '应包含服务端主题');
});

await check('子代理工具多选：点击会增删并写回数组', () => {
    const value = JSON.parse(JSON.stringify(SAMPLE_VALUE));
    value.search.subagentTools = ['web_search', 'write'];
    const scope = makeScope(value);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const row = findByLabel(tree, '子代理可用工具');
    assert.ok(row, '应渲染子代理工具这一行');
    const chips = walk(row).filter((node) => node.props && typeof node.props.onClick === 'function' && node.props.role === 'checkbox');
    assert.equal(chips.length, 7, '应渲染七个可选工具');
    // 勾上一个未选的（platform_search）
    const target = chips.find((c) => walk(c).some((n) => {
        const kids = n.props && n.props.children;
        return Array.isArray(kids) && kids.includes('平台检索');
    }));
    assert.ok(target, '应找到「平台检索」这个标签');
    target.props.onClick(true);
    assert.deepEqual(scope.writes[0], ['search.subagentTools', ['web_search', 'write', 'platform_search']]);
    // 再取消一个已选的（write）
    const writeChip = chips.find((c) => walk(c).some((n) => {
        const kids = n.props && n.props.children;
        return Array.isArray(kids) && kids.includes('写文件');
    }));
    writeChip.props.onClick(false);
    assert.deepEqual(scope.writes[1], ['search.subagentTools', ['web_search']]);
});

await check('取消 write 时给出落盘警告', () => {
    const value = JSON.parse(JSON.stringify(SAMPLE_VALUE));
    value.search.subagentTools = ['web_search'];
    const scope = makeScope(value);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const texts = walk(tree).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('无法落盘')), '去掉 write 应提示结果无法落盘');
});

await check('只读连接下控件被禁用', () => {
    const scope = makeScope(SAMPLE_VALUE, { writable: false });
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const toggle = controlOf(findByLabel(tree, 'office_help'));
    assert.equal(toggle.props.disabled, true, '只读时开关应禁用');
    const texts = walk(tree).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('只读')), '应提示只读');
});

await check('写入失败时把错误显示出来，而不是静默吞掉', async () => {
    const scope = makeScope(SAMPLE_VALUE);
    scope.forms.get = () => ({
        getSnapshot: () => scope.snapshot,
        subscribe: () => () => {},
        set: () => Promise.reject(new Error('revision 冲突')),
        unset: () => Promise.resolve(true),
        mutate: () => Promise.reject(new Error('revision 冲突')),
    });
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const toggle = controlOf(findByLabel(tree, 'office_help'));
    toggle.props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 10));
    const after = render(officeReg(registrations));
    const texts = walk(after).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('写入失败')), '应显示写入失败');
});

// ── 排版结构 ──────────────────────────────────────────────────────────────
// 这里防的是一类「功能对、但看起来坏掉」的回归：Row 用 flex 左右排，
// 左侧标题 flex 可被压到接近 0，而右侧控件 flex: 0 0 auto 不可压。
// 一旦某一行的控件很宽（七个可勾选标签），标题就会被挤成一字一行。
// 所以：宽控件的那一行必须走 stacked（上下排），且标题列必须有宽度下限。

await check('宽控件的那一行改成上下排（标题不再被挤成竖排）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    // 第二十轮起「联网通道」那一行也是 stacked（下拉本身很宽），所以按标签找，
    // 不按「第一个 stacked 行」找 —— 否则以后每加一个宽控件行都会误伤这条断言。
    const chipRow = walk(tree).find((n) => n.source && n.source.name === 'Row'
        && n.source.props.stacked === true && n.source.props.label === '子代理可用工具');
    assert.ok(chipRow, '子代理工具那一行应标记 stacked');
    assert.equal(chipRow.props.style.flexDirection, 'column', 'stacked 行要竖排');
    assert.equal(chipRow.props.style.alignItems, 'stretch');
});

await check('普通行仍是左右排（没有把所有行都改成上下排）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const row = findByLabel(tree, '并发子代理数');
    assert.ok(row, '应找到并发数这一行');
    assert.notEqual(row.source.props.stacked, true, '普通行不该走 stacked');
    assert.equal(row.props.style.display, 'flex');
    assert.notEqual(row.props.style.flexDirection, 'column');
});

await check('标题列有宽度下限（窄屏下也不会被压成竖排）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const rows = walk(tree).filter((n) => n.source && n.source.name === 'Row' && n.source.props.stacked !== true);
    assert.ok(rows.length >= 10, '应渲染出多行');
    for (const row of rows) {
        const label = walk(row.props.children).find((n) => n.props && n.props.style && n.props.style.minWidth === '180px');
        assert.ok(label, `行「${row.source.props.label}」的标题列缺少 minWidth 下限`);
    }
});

// ── 与服务端一致性 ────────────────────────────────────────────────────────

await check('界面上的工具键与服务端 DEFAULT_TOOLS 完全一致', async () => {
    const { DEFAULT_TOOLS } = await import('../src/config.js');
    const source = readFileSync(BUNDLE, 'utf8');
    for (const key of Object.keys(DEFAULT_TOOLS)) {
        assert.ok(source.includes(`key: "${key}"`) || source.includes(`"${key}"`), `界面缺少工具 ${key}`);
    }
});

await check('界面上的子代理工具与服务端白名单一致', async () => {
    const { CHANNEL_TOOLS } = await import('../src/search.js');
    const source = readFileSync(BUNDLE, 'utf8');
    for (const id of CHANNEL_TOOLS) {
        assert.ok(source.includes(`id: "${id}"`), `界面缺少子代理工具 ${id}`);
    }
});

await check('界面上的主题与服务端主题表一致', async () => {
    const { THEMES } = await import('../src/engine/theme.js');
    const source = readFileSync(BUNDLE, 'utf8');
    for (const theme of THEMES) {
        assert.ok(source.includes(`id: "${theme.id}"`), `界面缺少主题 ${theme.id}`);
    }
});

/**
 * 把 schema 的一组字段摊成「点号路径 → 叶子」的清单。
 *
 * 必须递归：layers / recallQuality / quota 是嵌套对象，只查顶层键的话
 * 「memory.layers」这种前缀会被源码里的 memory.layers.hot 顺带命中，断言就
 * 变成永远通过的空转 —— 这正是加了嵌套组之后最容易漏掉的那类回归。
 */
function leafPaths(node, prefix, out = []) {
    for (const key of Object.keys(node ?? {})) {
        const value = node[key];
        const path = prefix === '' ? key : `${prefix}.${key}`;
        if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
            leafPaths(value, path, out);
        } else {
            out.push(path);
        }
    }
    return out;
}

/** 逐项核对界面源码里有没有对应的写回路径；返回摊平后的叶子清单。 */
function assertCovers(source, schema, group) {
    const paths = leafPaths(schema, group);
    assert.ok(paths.length > 0, `${group} 组没有解析出任何参数`);
    for (const path of paths) {
        assert.ok(source.includes(path), `界面缺少参数 ${path}`);
    }
    return paths;
}

await check('界面覆盖服务端 search 组里的每个参数（含嵌套的 builtin）', async () => {
    // 「办公 preset 里没有 web_search」这件事的修法就落在这一组参数上：
    // schema 加了 search.builtin.* 而界面忘了加行 = 用户永远调不到它。
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.search, 'search');
    for (const nested of ['search.engine', 'search.builtin.apiKeyEnv', 'search.builtin.fetchPages']) {
        assert.ok(paths.includes(nested), `没有解析出嵌套参数 ${nested}`);
    }
    assert.ok(paths.length >= 20, `search 组参数过少：${paths.length}`);
});

await check('界面覆盖服务端 documents 组里的每个参数', async () => {
    // 服务端 schema 加了参数、界面忘了加行 = 用户永远调不到它，而两边都不会报错。
    // 这里用源码里的字段名（"documents.xxx"）逐项核对，把「静默少一个开关」钉住。
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.documents, 'documents');
    assert.ok(paths.length >= 10, `documents 组参数过少：${paths.length}`);
});

await check('界面覆盖服务端 memory 组里的每个参数（含嵌套的 layers / recallQuality / quota）', async () => {
    // 与 documents 那条同一个目的：schema 加了记忆参数、界面忘了加行 =
    // 用户永远调不到它，而两边都不会报错。
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.memory, 'memory');
    assert.ok(paths.length >= 20, `memory 组参数过少：${paths.length}`);
    // 嵌套组的叶子必须真的被摊出来，否则上面那条递归等于没跑。
    for (const nested of ['memory.layers.hot', 'memory.recallQuality.policy', 'memory.quota.recallPerTurn']) {
        assert.ok(paths.includes(nested), `没有解析出嵌套参数 ${nested}`);
    }
});

await check('界面覆盖服务端 subagentModel 组里的每个参数', async () => {
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.subagentModel, 'subagentModel');
    assert.deepEqual(paths.sort(), ['subagentModel.mode', 'subagentModel.model', 'subagentModel.provider']);
});

await check('界面覆盖服务端 av 组里的每个参数（音频与视频）', async () => {
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.av, 'av');
    assert.ok(paths.length >= 14, `av 组该有十四个以上参数，实际 ${paths.length}`);
    return `${paths.length} 个参数都有对应控件`;
});

await check('音频与视频：语言下拉覆盖服务端的六种取值', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const select = controlOf(findByLabel(tree, '转写语言'));
    assert.equal(select.type, 'select');
    assert.equal(select.props.value, 'zh', '样例值给的是 zh');
    const options = walk(select.props.children).filter((node) => node.type === 'option');
    const ids = options.map((node) => node.props.value);
    assert.deepEqual(ids, ['auto', 'zh', 'yue', 'en', 'ja', 'ko'], '语言下拉要覆盖服务端 AV_LANGUAGES');
    return `六个选项：${ids.join(' / ')}`;
});

await check('记忆系统面板：开关、目录、容量控件写回 memory.<name>', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    const toggleRow = findByLabel(tree, '开启记忆');
    assert.ok(toggleRow, '应渲染「开启记忆」这一行');
    const toggle = controlOf(toggleRow);
    assert.equal(toggle.props.type, 'checkbox');
    assert.equal(toggle.props.checked, true);
    toggle.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[0], ['memory.enabled', false]);

    const dirRow = findByLabel(tree, '记忆目录');
    assert.ok(dirRow, '应渲染「记忆目录」这一行');
    const dirInput = controlOf(dirRow);
    assert.equal(dirInput.props.defaultValue, '.office/memory');
    dirInput.props.onBlur(eventOf({ value: '.my-office-memory' }));
    assert.deepEqual(scope.writes[1], ['memory.dir', '.my-office-memory']);

    const limitRow = findByLabel(tree, '用户偏好上限');
    assert.ok(limitRow, '应渲染「用户偏好上限」这一行');
    const input = controlOf(limitRow);
    assert.equal(input.props.type, 'number');
    assert.equal(input.props.value, '4096');
    input.props.onChange(eventOf({ value: '8192' }));
    // 数字控件是「本地草稿 + 失焦提交」：草案进 state 之后要重新渲染一次，
    // 新的 onBlur 才拿得到新草稿（真实 React 也是这个次序）。
    const afterTyping = render(memoryReg(registrations));
    const input2 = controlOf(findByLabel(afterTyping, '用户偏好上限'));
    assert.equal(input2.props.value, '8192', '草稿应保留在控件里');
    input2.props.onBlur();
    assert.deepEqual(scope.writes[2], ['memory.userLimitBytes', 8192]);
});

await check('记忆系统面板：写系统提示的开关与迁移说明都在', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    const hintRow = findByLabel(tree, '写入系统提示');
    assert.ok(hintRow, '应渲染「写入系统提示」这一行');
    const hint = controlOf(hintRow);
    assert.equal(hint.props.checked, true);
    hint.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[0], ['memory.promptHint', false]);

    // 迁移是这一页的重点：没有执行通道，所以必须给确切命令。
    const texts = walk(tree).map((node) => node.props && node.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('node scripts/migrate-mnemon.mjs')), '应给出迁移命令');
    assert.ok(texts.some((t) => t.includes("action: 'migrate'")), '应给出工具调用方式');
    assert.ok(texts.some((t) => t.includes("layer: 'archive'")), '应说明迁进去的长期记忆怎么检索');
});

await check('两个面板共用同一个设置命名空间，且各自只渲染自己那些控件', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    // bind 只发生一次：两页读同一份设置（分开 bind 会各自订阅一份镜像）。
    assert.equal(scope.forms.bound.entryId, 'dsh-office-mode');

    const office = render(officeReg(registrations));
    const memory = renderMemory(registrations);
    const officeTexts = walk(office).map((node) => node.props && node.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(findByLabel(office, 'office_run'), '办公模式页有工具开关');
    assert.ok(!findByLabel(office, '开启记忆'), '记忆控件不该在办公模式页重复出现');
    assert.ok(officeTexts.some((t) => t.includes('记忆系统')), '办公模式页要留一句指路');
    assert.ok(findByLabel(memory, '开启记忆'), '记忆系统页有记忆开关');
    assert.ok(!findByLabel(memory, 'office_run'), '记忆系统页不该有工具开关');
});

// ── 新增：记忆参数控件（嵌套组） ──────────────────────────────────────────

await check('记忆层拓扑：三个开关按服务端值显示，写回 memory.layers.<层>', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    const hot = controlOf(findByLabel(tree, '热记忆层'));
    assert.ok(hot, '应渲染「热记忆层」这一行');
    assert.equal(hot.props.type, 'checkbox');
    assert.equal(hot.props.checked, true);

    // SAMPLE_VALUE 里 archive 是关的：关掉的层要显示为未勾选。
    const archive = controlOf(findByLabel(tree, '归档层'));
    assert.equal(archive.props.checked, false, 'archive=false 应显示未勾选');

    archive.props.onChange(eventOf({ checked: true }));
    assert.deepEqual(scope.writes[0], ['memory.layers.archive', true]);
});

await check('记忆范围：下拉与文本框写回 memory.scope / globalDir / userScope', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    const scopeSelect = controlOf(findByLabel(tree, '记忆放哪儿'));
    assert.equal(scopeSelect.type, 'select');
    assert.equal(scopeSelect.props.value, 'workspace');
    const scopeIds = walk(scopeSelect.props.children).filter((n) => n.type === 'option').map((n) => n.props.value);
    assert.deepEqual(scopeIds, ['workspace', 'global', 'both'], '下拉要覆盖服务端 MEMORY_SCOPES');
    scopeSelect.props.onChange(eventOf({ value: 'both' }));
    assert.deepEqual(scope.writes[0], ['memory.scope', 'both']);

    const dirInput = controlOf(findByLabel(tree, '全局层目录'));
    assert.equal(dirInput.props.placeholder, '留空 = $DSH_HOME/.office/memory');
    dirInput.props.onBlur(eventOf({ value: '/data/office-memory' }));
    assert.deepEqual(scope.writes[1], ['memory.globalDir', '/data/office-memory']);

    const userSelect = controlOf(findByLabel(tree, '用户偏好范围'));
    assert.equal(userSelect.props.value, 'memory');
    const userIds = walk(userSelect.props.children).filter((n) => n.type === 'option').map((n) => n.props.value);
    assert.deepEqual(userIds, ['memory', 'global'], '下拉要覆盖服务端 USER_SCOPES');
    userSelect.props.onChange(eventOf({ value: 'global' }));
    assert.deepEqual(scope.writes[2], ['memory.userScope', 'global']);
});

await check('图关系与主动记录两个开关写回 memory.links / memory.autoCapture', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    const links = controlOf(findByLabel(tree, '图关系'));
    assert.ok(links, '应渲染「图关系」这一行');
    assert.equal(links.props.checked, true);
    links.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[0], ['memory.links', false]);

    const capture = controlOf(findByLabel(tree, '主动记录'));
    assert.ok(capture, '应渲染「主动记录」这一行');
    assert.equal(capture.props.checked, true);
    capture.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[1], ['memory.autoCapture', false]);
});

await check('召回质量：策略下拉 + 五个数字，阈值按 0.05 吸附而不是取整', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    const policy = controlOf(findByLabel(tree, '分档策略'));
    assert.equal(policy.type, 'select');
    assert.equal(policy.props.value, 'strict-v1');
    const policyIds = walk(policy.props.children).filter((n) => n.type === 'option').map((n) => n.props.value);
    assert.deepEqual(policyIds, ['strict-v1', 'off']);
    policy.props.onChange(eventOf({ value: 'off' }));
    assert.deepEqual(scope.writes[0], ['memory.recallQuality.policy', 'off']);

    for (const label of ['高分阈值', '候选倍数', '中档最多几条', '未知档最多几条']) {
        assert.ok(findByLabel(tree, label), `应渲染「${label}」这一行`);
    }

    const low = controlOf(findByLabel(tree, '低分阈值'));
    assert.equal(low.props.step, 0.05, '阈值步长必须是 0.05');
    assert.equal(String(low.props.value), '0.25');
    // 关键回归：这一格曾经会被 Math.trunc 把 0.35 变成 0 —— 阈值调到 0 就等于
    // 把整档过滤关掉，而用户看不出来。0.05 步长必须走「吸附」那条路。
    low.props.onChange(eventOf({ value: '0.35' }));
    const after = renderMemory(registrations);
    controlOf(findByLabel(after, '低分阈值')).props.onBlur();
    assert.deepEqual(scope.writes[1], ['memory.recallQuality.lowScoreThreshold', 0.35]);
});

await check('每回合配额：三个数字写回 memory.quota.<项>', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);

    const rows = [
        ['首次检索', 'memory.quota.recallPerTurn'],
        ['换词细化', 'memory.quota.recallRefinePerTurn'],
        ['图关系遍历', 'memory.quota.relatedPerTurn'],
    ];
    for (const [label, path] of rows) {
        const tree = renderMemory(registrations);
        const row = findByLabel(tree, label);
        assert.ok(row, `应渲染「${label}」这一行`);
        const input = controlOf(row);
        assert.equal(input.props.type, 'number');
        assert.equal(input.props.min, 0);
        assert.equal(input.props.max, 50);
        input.props.onChange(eventOf({ value: '3' }));
        const after = renderMemory(registrations);
        controlOf(findByLabel(after, label)).props.onBlur();
        assert.deepEqual(scope.writes[scope.writes.length - 1], [path, 3]);
    }
});

await check('子代理模型：模式下拉与两个文本框写回 subagentModel.<项>', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));

    const mode = controlOf(findByLabel(tree, '模型路由'));
    assert.ok(mode, '应渲染「模型路由」这一行');
    assert.equal(mode.type, 'select');
    assert.equal(mode.props.value, 'inherit');
    const modeIds = walk(mode.props.children).filter((n) => n.type === 'option').map((n) => n.props.value);
    assert.deepEqual(modeIds, ['inherit', 'fixed']);
    mode.props.onChange(eventOf({ value: 'fixed' }));
    assert.deepEqual(scope.writes[0], ['subagentModel.mode', 'fixed']);

    const provider = controlOf(findByLabel(tree, 'Provider 路由名'));
    provider.props.onBlur(eventOf({ value: 'deepseek' }));
    assert.deepEqual(scope.writes[1], ['subagentModel.provider', 'deepseek']);

    const model = controlOf(findByLabel(tree, '模型 id'));
    assert.ok(model, '应渲染「模型 id」这一行');
    model.props.onBlur(eventOf({ value: 'deepseek-chat' }));
    assert.deepEqual(scope.writes[2], ['subagentModel.model', 'deepseek-chat']);
});

// ── 新增：slot 注册面 ─────────────────────────────────────────────────────

await check('注册了记忆浏览面板：侧栏入口 id 与主面板 key 完全一致', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const icon = regOf(registrations, 'sidebar.panellist', 'office-memory');
    const panel = regOf(registrations, 'main', 'office-memory');
    assert.ok(icon, '侧栏应有「记忆」入口');
    assert.ok(panel, '主面板应有 office-memory 这一格');
    // 这一条是整件事的接缝：侧栏按钮靠 id === key 找到要选中的那一格。
    assert.equal(icon.opts.id, panel.opts.key, '侧栏 id 必须等于主面板 key，否则按钮点不亮面板');
    assert.equal(typeof icon.opts.order, 'number');
    assert.equal(typeof icon.opts.label, 'function', 'label 要用 thunk（支持本地化重取）');
    assert.equal(icon.opts.label(), '记忆');
    assert.equal(typeof icon.render, 'function');
    assert.equal(typeof panel.render, 'function');
    // main 是 keyed slot：注册参数里只有 key，多写 id 会被当成别的格子。
    assert.deepEqual(Object.keys(panel.opts).sort(), ['key', 'name']);
});

await check('注册了会话内两处：回合记忆条与存入记忆按钮', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const bar = regOf(registrations, 'conversation.chat.turnTail', 'office-memory-turnbar');
    const save = regOf(registrations, 'conversation.chat.assistant-actions', 'office-memory-save');
    assert.ok(bar, '应有回合记忆条');
    assert.ok(save, '应有存入记忆按钮');
    assert.equal(typeof bar.opts.order, 'number');
    assert.equal(typeof save.opts.order, 'number');
});

await check('五种 slot 都注册了，且 inject 与 register 一一对应', () => {
    const { mod, ctx, registrations, injected } = loadBundle();
    mod.apply(ctx);
    const slots = [...new Set(registrations.map((item) => item.opts.name))].sort();
    assert.deepEqual(slots, [
        'conversation.chat.assistant-actions',
        'conversation.chat.turnTail',
        'main',
        'settings.section',
        'sidebar.panellist',
    ]);
    assert.equal(injected.length, registrations.length, '每个 inject 应对应一次 register');
});

await check('侧栏图标用上 owner 给的 size，并以 currentColor 描边', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const icon = regOf(registrations, 'sidebar.panellist', 'office-memory');
    const tree = render(icon, { size: 20, active: true });
    assert.equal(tree.type, 'svg');
    assert.equal(tree.props.width, 20, 'size 要用上（折叠态与展开态给的尺寸不同）');
    assert.equal(tree.props.height, 20);
    assert.equal(tree.props.stroke, 'currentColor');
    assert.equal(tree.props['aria-hidden'], 'true');
    // 宿主在异常路径下可能不传 props：不能因此抛错。
    assert.doesNotThrow(() => render(icon));
});

// ── 新增：记忆浏览面板 ────────────────────────────────────────────────────

/** 取渲染树里的全部文本节点。 */
function textsOf(tree) {
    return walk(tree).map((node) => node.props && node.props.children).flat().filter((c) => typeof c === 'string');
}

/** 渲染面板并等一轮微任务，让假 fetch 的 promise 落地。 */
async function renderPanelLoaded(panel) {
    render(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    return render(panel);
}

await check('记忆面板：挂载后取数，五个 Tab 显示 counts 里的条数', async () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const panel = regOf(registrations, 'main', 'office-memory');
    const tree = await renderPanelLoaded(panel);
    const texts = textsOf(tree);

    assert.ok(texts.some((t) => t.includes('记忆 · temp1')), '应显示工作目录短名');
    for (const label of ['热记忆（1）', '台账（1）', '归档（1）', '关系（1）', '实体（1）']) {
        assert.ok(texts.some((t) => t === label), `Tab 应显示条数：${label}`);
    }
    // 默认 Tab 是热记忆：显示内容 + 来源徽标 + 重要度
    assert.ok(texts.some((t) => t.includes('偏好：书面化表达')), '应显示热记忆内容');
    assert.ok(texts.some((t) => t === '全局'), '应显示来源徽标（全局 / 工作区）');
    assert.ok(texts.some((t) => t === 'critical'), '应显示重要度');
});

await check('记忆面板：切换 Tab 显示对应层的条目', async () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const panel = regOf(registrations, 'main', 'office-memory');
    let tree = await renderPanelLoaded(panel);

    // 台账 Tab
    const ledgerTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && Array.isArray(n.props.children) && n.props.children.includes('台账（1）'));
    assert.ok(ledgerTab, '应找到台账 Tab');
    ledgerTab.props.onClick();
    tree = render(panel);
    let texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('季度汇报.pptx')), '台账 Tab 应显示产物路径');

    // 归档 Tab
    const archiveTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && Array.isArray(n.props.children) && n.props.children.includes('归档（1）'));
    archiveTab.props.onClick();
    tree = render(panel);
    texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('长期记忆一条')), '归档 Tab 应显示条目文本');
    assert.ok(texts.some((t) => t === 'mnemon'), '归档应显示 kind 徽标');
    assert.ok(texts.some((t) => t === '2026-09'), '归档应显示月份');

    // 关系 Tab
    const linksTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && Array.isArray(n.props.children) && n.props.children.includes('关系（1）'));
    linksTab.props.onClick();
    tree = render(panel);
    texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('m-1 —[related]→ L-1')), '关系 Tab 应显示 sourceId —[kind]→ targetId');

    // 实体 Tab
    const entitiesTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && Array.isArray(n.props.children) && n.props.children.includes('实体（1）'));
    entitiesTab.props.onClick();
    tree = render(panel);
    texts = textsOf(tree);
    assert.ok(texts.some((t) => t === '系统A'), '实体 Tab 应显示实体名');
    assert.ok(texts.some((t) => t.includes('出现 2 次')), '实体 Tab 应显示出现次数');
    assert.ok(texts.some((t) => t.includes('偏好：书面化表达')), '实体 Tab 应显示指向的条目');
});

await check('记忆面板：搜索回车带 q=、换目录带 cwd= 且不丢过滤词、刷新重取', async () => {
    const urls = [];
    const fetchImpl = (url) => {
        urls.push(url);
        return Promise.resolve({
            ok: true,
            status: 200,
            json: () => Promise.resolve(makeSnapshot({
                workspaces: [
                    { cwd: 'D:/proj/temp1', label: 'temp1', seenAt: 'x' },
                    { cwd: 'D:/proj/other', label: 'other', seenAt: 'y' },
                ],
            })),
        });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'main', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    assert.equal(urls.length, 1, '挂载时应取一次数');
    assert.equal(urls[0], '/office-memory/snapshot', '同源相对路径，且首次不带参数');

    const search = walk(tree).find((n) => n.type === 'input' && n.props && typeof n.props.placeholder === 'string' && n.props.placeholder.includes('过滤'));
    assert.ok(search, '应有搜索框');
    search.props.onKeyDown({ key: 'Enter', target: { value: '季度' } });
    tree = await renderPanelLoaded(panel);
    assert.ok(urls[1].includes('q=' + encodeURIComponent('季度')), '搜索应带 q=');

    const select = walk(tree).find((n) => n.type === 'select');
    assert.ok(select, '多个工作目录时应给下拉');
    const options = walk(select.props.children).filter((n) => n.type === 'option').map((n) => n.props.value);
    assert.deepEqual(options, ['D:/proj/temp1', 'D:/proj/other']);
    select.props.onChange(eventOf({ value: 'D:/proj/other' }));
    tree = await renderPanelLoaded(panel);
    assert.ok(urls[2].includes('cwd=' + encodeURIComponent('D:/proj/other')), '换目录应带 cwd=');
    assert.ok(urls[2].includes('q=' + encodeURIComponent('季度')), '换目录不该丢掉已应用的过滤词');

    const before = urls.length;
    const refresh = walk(tree).find((n) => n.type === 'button' && Array.isArray(n.props.children) && n.props.children.includes('刷新'));
    assert.ok(refresh, '应有刷新按钮');
    refresh.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 10));
    assert.equal(urls.length, before + 1, '刷新应重新取数');
});

await check('记忆面板：端点报错时显示服务端给的 error，并保留 Tab 骨架', async () => {
    const fetchImpl = () => Promise.resolve({
        ok: false,
        status: 404,
        json: () => Promise.resolve({ ok: false, error: '这个工作目录不在已知列表里。', workspaces: [] }),
    });
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'main', 'office-memory');
    const tree = await renderPanelLoaded(panel);
    const texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('不在已知列表里')), '应显示服务端给的错误');
    // 出错不等于白屏：Tab 仍在（条数退化成 0）。
    assert.ok(texts.some((t) => t === '热记忆（0）'), '出错时也应保留 Tab 骨架');
});

await check('记忆面板：环境没有 fetch 时给出可读提示，不抛错', () => {
    const { mod, ctx, registrations } = loadBundle({ fetch: undefined });
    mod.apply(ctx);
    const panel = regOf(registrations, 'main', 'office-memory');
    const tree = render(panel);
    const texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('没有 fetch')), '应提示当前环境没有 fetch');
});

// ── 新增：记忆面板的可视化 ────────────────────────────────────────────────
//
// 这一组钉的是「面板不只是两行文字」：指标块、容量条、存储域卡片、时间线、
// 实体频次条、力导向关系图，各自的数值都必须能对回快照里的数。

/** 取渲染树里带某个属性的元素；value 省略 = 只要带这个属性。 */
function byProp(tree, name, value) {
    return walk(tree).filter((node) => node.props
        && (value === undefined ? name in node.props : node.props[name] === value));
}

/** 渲染面板（假 fetch 返回给定快照）并等一次取数落地。 */
async function panelFrom(snapshot, options = {}) {
    const fetchImpl = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(snapshot) });
    const { mod, ctx, registrations } = loadBundle(Object.assign({ fetch: fetchImpl }, options));
    mod.apply(ctx);
    const panel = regOf(registrations, 'main', 'office-memory');
    const tree = await renderPanelLoaded(panel);
    return { tree, panel, registrations };
}

/** 造 n 条热记忆（分页用例用）：user / project 交替，便于核对分组总数。 */
function manyHot(count) {
    const out = [];
    for (let index = 0; index < count; index += 1) {
        out.push({
            id: 'm-' + index,
            target: index % 2 === 0 ? 'user' : 'project',
            importance: 'normal',
            origin: 'workspace',
            updatedAt: 'x',
            entities: [],
            tags: [],
            content: '条目内容 ' + index,
        });
    }
    return out;
}

/** 从当前树里点开某个 Tab，返回重渲染后的树。 */
function openTab(tree, panel, label) {
    const tab = walk(tree).find((n) => n.props && n.props.role === 'tab'
        && Array.isArray(n.props.children) && n.props.children.includes(label));
    assert.ok(tab, `应找到 Tab ${label}`);
    tab.props.onClick();
    return render(panel);
}

/** 树里的全部字符串（含 style 的取值）：用来钉住「不许出现 NaN / undefined」。 */
function stringsOf(tree) {
    const out = [];
    walk(tree).forEach((node) => {
        for (const value of Object.values(node.props || {})) {
            if (typeof value === 'string') out.push(value);
            else if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
                for (const inner of Object.values(value)) if (typeof inner === 'string') out.push(inner);
            }
        }
    });
    return out;
}

/** 读 translate(x y) 里的两个坐标。 */
function parseTranslate(value) {
    const match = /translate\(([-\d.]+)\s+([-\d.]+)\)/.exec(String(value));
    return match === null ? null : { x: Number(match[1]), y: Number(match[2]) };
}

await check('可视化：标题下五枚指标块显示 counts 里的条数', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const tiles = byProp(tree, 'data-metric');
    assert.deepEqual(tiles.map((n) => n.props['data-metric']), ['hot', 'ledger', 'archive', 'links', 'entities'],
        '五枚指标块：热记忆 / 台账 / 归档 / 关系 / 实体');
    const values = byProp(tree, 'data-metric-value');
    assert.equal(values.length, 5);
    for (const tile of values) {
        assert.deepEqual(tile.props.children, ['1'], `${tile.props['data-metric-value']} 应显示 counts 里的 1`);
    }
    const texts = textsOf(tree);
    for (const label of ['热记忆', '台账', '归档', '关系', '实体']) {
        assert.ok(texts.some((t) => t === label), `指标块应有名称 ${label}`);
    }
});

await check('可视化：容量条按 count / limit 算宽度，缺字节字段时整条不画', async () => {
    const { tree } = await panelFrom(makeSnapshot());

    const ledger = byProp(tree, 'data-gauge', 'ledger')[0];
    assert.ok(ledger, '应有台账容量条');
    assert.equal(ledger.props['data-percent'], 0.2, '1 / 500 = 0.2%（取整会变成 0，看不见）');
    assert.equal(byProp(ledger, 'data-gauge-fill', 'ledger')[0].props.style.width, '0.2%');
    assert.ok(textsOf(ledger).some((t) => t.includes('1 / 500')), '容量条要同时给出 n / limit 文本');

    const archive = byProp(tree, 'data-gauge', 'archive')[0];
    assert.ok(archive, '应有归档容量条');
    assert.equal(archive.props['data-percent'], 1.7, '1 / 60 = 1.7%');

    // 快照里还没有 counts.hotBytes / projectBytes（并发在加）：这两条不能画成 NaN。
    assert.equal(byProp(tree, 'data-gauge', 'hotBytes').length, 0, '缺字节数时不画热记忆占用条');
    assert.equal(byProp(tree, 'data-gauge', 'projectBytes').length, 0, '缺字节数时不画项目记忆占用条');
});

await check('可视化：字节数存在时补上两条容量条，并按各自上限算比例', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        counts: {
            hot: 1, ledger: 1, archive: 1, links: 1, entities: 1,
            ledgerTotal: 1, archiveFiles: 1, archiveItems: 1,
            hotBytes: 1024, projectBytes: 5120,
        },
    }));
    const hotBytes = byProp(tree, 'data-gauge', 'hotBytes')[0];
    assert.ok(hotBytes, '有 hotBytes 时应画热记忆占用条');
    assert.equal(hotBytes.props['data-percent'], 25, '1024 / 4096 = 25%');
    assert.ok(textsOf(hotBytes).some((t) => t.includes('1024 / 4096')));

    const project = byProp(tree, 'data-gauge', 'projectBytes')[0];
    assert.ok(project, '有 projectBytes 时应画项目记忆占用条');
    assert.equal(project.props['data-percent'], 50, '5120 / 10240 = 50%');
    assert.equal(byProp(project, 'data-gauge-fill', 'projectBytes')[0].props.style.width, '50%');
});

await check('可视化：存储域卡片显示目录、四层条数与可读的字节数', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        stores: [{ id: 'workspace', dir: '.office/memory', hot: 1, ledger: 1, archive: 1, links: 1, bytes: 2048, files: 7 }],
    }));
    const card = byProp(tree, 'data-store', 'workspace')[0];
    assert.ok(card, '应有 workspace 存储卡');
    const texts = textsOf(card);
    assert.ok(texts.includes('.office/memory'), '卡片要写目录');
    for (const chip of ['热 1', '台账 1', '归档 1', '关系 1']) {
        assert.ok(texts.includes(chip), `卡片缺少层的条数：${chip}`);
    }
    assert.ok(texts.includes('2 KB'), '字节数要人读化（2048 → 2 KB）');
    assert.ok(texts.includes('7 个文件'), '有 files 时要显示文件数');
});

await check('可视化：台账按天、归档按月画时间线，柱高按峰值折算', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 1, ledger: 3, archive: 1, links: 1, entities: 1 },
        ledger: [
            { id: 'L-1', at: '2026-09-24T00:00:00.000Z', path: '季度汇报.pptx', format: 'ppt', theme: 'plain', purpose: '季度汇报', outline: ['封面'], origin: 'workspace' },
            { id: 'L-2', at: '2026-09-24T03:00:00.000Z', path: 'a.docx', format: 'word', origin: 'workspace' },
            { id: 'L-3', at: '2026-09-25T03:00:00.000Z', path: 'b.xlsx', format: 'excel', origin: 'workspace' },
        ],
    }));

    const ledgerTree = openTab(tree, panel, '台账（3）');
    const timeline = byProp(ledgerTree, 'data-panel', 'ledger-timeline')[0];
    assert.ok(timeline, '台账 Tab 应画按天分布');
    const buckets = byProp(timeline, 'data-bucket');
    assert.deepEqual(buckets.map((n) => n.props['data-bucket']), ['2026-09-24', '2026-09-25'], '按天分桶且升序');
    assert.deepEqual(buckets.map((n) => n.props['data-count']), [2, 1]);
    const bars = buckets.map((n) => walk(n).find((x) => x.props && x.props.style && x.props.style.height));
    assert.equal(bars[0].props.style.height, '44px', '峰值时段满高');
    assert.equal(bars[1].props.style.height, '22px', '一半的时段半高');

    const archiveTree = openTab(ledgerTree, panel, '归档（1）');
    const months = byProp(archiveTree, 'data-panel', 'archive-months')[0];
    assert.ok(months, '归档 Tab 应画按月分布');
    assert.deepEqual(byProp(months, 'data-bucket').map((n) => n.props['data-bucket']), ['2026-09']);
});

await check('可视化：实体频次条按 count / max 算宽度，并给出来源层 chip', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 2, ledger: 1, archive: 1, links: 1, entities: 2 },
        entities: [
            {
                name: '系统A', count: 4,
                refs: [
                    { id: 'm-1', layer: 'hot', origin: 'workspace', text: '偏好：书面化表达' },
                    { id: 'm-2', layer: 'hot', origin: 'workspace', text: '约定：只用 edit 工具' },
                    { id: 'L-1', layer: 'ledger', origin: 'workspace', text: '季度汇报.pptx' },
                ],
            },
            { name: '系统B', count: 2, refs: [{ id: 'L-1', layer: 'ledger', origin: 'workspace', text: '季度汇报.pptx' }] },
        ],
    }));
    const entityTree = openTab(tree, panel, '实体（2）');
    const bars = byProp(entityTree, 'data-entity');
    assert.deepEqual(bars.map((n) => n.props['data-entity']), ['系统A', '系统B']);
    assert.deepEqual(bars.map((n) => n.props['data-percent']), [100, 50], '条长 = count / max');
    assert.equal(byProp(bars[1], 'data-entity-fill', '系统B')[0].props.style.width, '50%');
    assert.ok(textsOf(bars[0]).some((t) => t === '出现 4 次'), '条上要写出现次数');

    // 层 chip：按 refs[].layer 去重，同一层多次出现带上次次。
    const chipsA = byProp(bars[0], 'data-entity-layer');
    assert.deepEqual(chipsA.map((n) => n.props['data-entity-layer']), ['hot', 'ledger']);
    assert.deepEqual(chipsA.map((n) => n.props['data-entity-refs']), [2, 1]);
    assert.deepEqual(textsOf(bars[0]).filter((t) => t === '热记忆 ×2'), ['热记忆 ×2'], '多层引用要带次次');
    assert.deepEqual(byProp(bars[1], 'data-entity-layer').map((n) => n.props['data-entity-layer']), ['ledger']);
    assert.deepEqual(textsOf(bars[1]).filter((t) => t === '台账'), ['台账'], '只引用一次就不带次次');
    // 「实体」这一格也能拿到关系图（两个 Tab 是同一张图的两个切面）。
    assert.equal(byProp(entityTree, 'data-graph', 'memory').length, 1, '实体 Tab 也应能看到关系图');
    // 3 条 refs 只画 2 枚 chip；图谱上实体与条目相连的边则按**能解析到的**条目算：
    // m-2 不在快照里，那条边直接丢掉，不画悬空的线（3 条而不是 4 条）。
    assert.equal(byProp(entityTree, 'data-edge').filter((n) => n.props['data-edge'] === 'entity').length, 3);
});

await check('可视化：关系图画 SVG 力导向图，节点按层换形状、边按 kind 上色', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graph = byProp(linksTree, 'data-graph', 'memory')[0];
    assert.ok(graph, '关系 Tab 应画出关系图');
    assert.equal(graph.type, 'svg');
    assert.equal(graph.props.role, 'img', 'SVG 要声明成图片');
    assert.ok(String(graph.props['aria-label']).includes('节点'), 'SVG 要有可读的 aria-label');
    assert.equal(graph.props['data-layout'], 'force', '自写力导向布局');
    assert.equal(graph.props['data-nodes'], 4, '1 热记忆 + 1 台账 + 1 归档 + 1 实体');
    assert.equal(graph.props['data-edges'], 2, 'links 一条 + 实体引用一条');

    const nodes = walk(graph).filter((n) => n.type === 'g' && n.props && n.props['data-node']);
    assert.equal(nodes.length, 4);
    assert.deepEqual(nodes.map((n) => n.props['data-layer']).sort(), ['archive', 'entity', 'hot', 'ledger']);
    for (const node of nodes) {
        assert.equal(node.props.tabIndex, 0, '节点要能聚焦');
        assert.equal(node.props.role, 'button');
    }

    // 形状按层分：圆 = 热记忆、圆角方 = 台账、方 = 归档、菱形 = 实体。
    const core = walk(graph).filter((n) => n.props && n.props['data-shape'] === 'core');
    assert.equal(core.length, 4, '每个节点都有实心核');
    assert.equal(core.filter((n) => n.type === 'circle').length, 1, '热记忆节点是圆');
    assert.equal(core.filter((n) => n.type === 'rect').length, 2, '台账与归档是方');
    assert.equal(core.filter((n) => n.type === 'polygon').length, 1, '实体是菱形');
    for (const shape of core) assert.ok(shape.props.fill, '节点要有填充色');

    const edges = byProp(graph, 'data-edge');
    assert.deepEqual(edges.map((n) => n.props['data-edge']).sort(), ['entity', 'related']);
    for (const edge of edges) {
        assert.equal(edge.type, 'path', '关系边画成二次贝塞尔路径');
        assert.ok(String(edge.props.d).includes('Q'), '路径要用 Q 曲线，不是直线');
        assert.ok(edge.props.stroke, '边要有颜色');
    }

    // 布局必须真的把节点摊开：不能全叠在一个点，也不能跑出画布（GRID 是 720×380，
    // 边距 44）。
    const points = nodes.map((n) => parseTranslate(n.props.transform));
    for (const point of points) {
        assert.ok(point !== null, '节点要有 translate 定位');
        assert.ok(point.x >= 44 && point.x <= 676, `节点 x 越界：${JSON.stringify(point)}`);
        assert.ok(point.y >= 44 && point.y <= 336, `节点 y 越界：${JSON.stringify(point)}`);
    }
    assert.equal(new Set(points.map((p) => `${p.x},${p.y}`)).size, 4, '四个节点不能重叠成一堆');

    // 种子化布局：同一份快照重渲染必须落回同一张图，否则每次悬停整张图都会跳。
    const again = walk(byProp(render(panel), 'data-graph', 'memory')[0])
        .filter((n) => n.type === 'g' && n.props && n.props['data-node']);
    assert.deepEqual(again.map((n) => n.props.transform), nodes.map((n) => n.props.transform),
        '布局对同一份快照必须是确定性的');
});

await check('可视化：关系图的节点数封顶，记忆堆多了也不会把面板拖死', async () => {
    const hot = [];
    for (let index = 0; index < 60; index += 1) {
        hot.push({
            id: 'm-' + index, target: 'project', importance: 'normal', origin: 'workspace',
            updatedAt: 'x', entities: [], tags: [], content: '条目 ' + index,
        });
    }
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 60, ledger: 1, archive: 1, links: 1, entities: 1 },
        hot: hot,
    }));
    const linksTree = openTab(tree, panel, '关系（1）');
    const graph = byProp(linksTree, 'data-graph', 'memory')[0];
    assert.equal(graph.props['data-nodes'], 48, '节点数封顶 48（否则 O(n²) 的迭代会把面板拖死）');
    assert.equal(walk(graph).filter((n) => n.type === 'g' && n.props && n.props['data-node']).length, 48);
    // 有关系的节点优先保留：m-1 连到台账，实体连到 m-1，都不该被裁掉。
    const keys = walk(graph).filter((n) => n.type === 'g' && n.props && n.props['data-node']).map((n) => n.props['data-node']);
    assert.ok(keys.includes('hot:m-1') && keys.includes('ledger:L-1'), '有关系的节点要优先保留');
    assert.equal(graph.props['data-edges'], 2);
});

await check('可视化：关系图支持悬停高亮邻居、点击选中、方向键微调 12px', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graphOf = (rendered) => byProp(rendered, 'data-graph', 'memory')[0];
    const nodeOf = (rendered, key) => walk(graphOf(rendered)).find((n) => n.props && n.props['data-node'] === key);

    assert.equal(byProp(linksTree, 'data-graph-detail', 'none').length, 1, '默认没有聚焦节点');

    nodeOf(linksTree, 'hot:m-1').props.onMouseEnter();
    const hovered = render(panel);
    const detail = byProp(hovered, 'data-graph-detail', 'hot:m-1')[0];
    assert.ok(detail, '悬停后详情行要指向该节点');
    assert.ok(textsOf(detail).some((t) => t.includes('邻居 2 个')), 'm-1 连到台账与实体，共 2 个邻居');
    assert.equal(nodeOf(hovered, 'hot:m-1').props['data-focused'], 'true');
    assert.equal(byProp(hovered, 'data-edge').find((n) => n.props['data-edge'] === 'related').props.strokeWidth, 2.2,
        '与聚焦节点相连的边要加粗');
    nodeOf(hovered, 'hot:m-1').props.onMouseLeave();
    assert.equal(byProp(render(panel), 'data-graph-detail', 'none').length, 1, '移开就恢复');

    // 点击选中
    nodeOf(render(panel), 'ledger:L-1').props.onClick();
    const selected = render(panel);
    assert.ok(textsOf(byProp(selected, 'data-graph-detail', 'ledger:L-1')[0]).some((t) => t.includes('已选中')));
    assert.equal(nodeOf(selected, 'ledger:L-1').props['data-focused'], 'true');

    // 方向键微调：只挪 12px，基础布局不重算。
    const before = parseTranslate(nodeOf(selected, 'ledger:L-1').props.transform);
    assert.ok(before, '节点要有 translate 定位');
    nodeOf(selected, 'ledger:L-1').props.onKeyDown({ key: 'ArrowRight', preventDefault: () => {} });
    const nudged = render(panel);
    const after = parseTranslate(nodeOf(nudged, 'ledger:L-1').props.transform);
    assert.ok(Math.abs((after.x - before.x) - 12) <= 0.2, `右移应约 12px：${before.x} → ${after.x}`);
    assert.equal(after.y, before.y, '左右移动不该改 y');

    // Enter 再按一次取消选中
    nodeOf(nudged, 'ledger:L-1').props.onKeyDown({ key: 'Enter', preventDefault: () => {} });
    assert.equal(byProp(render(panel), 'data-graph-detail', 'none').length, 1, 'Enter 是切换选中');
});

await check('可视化：热记忆按 target 分组，重要度落在左边框上', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        counts: { hot: 3, ledger: 1, archive: 1, links: 1, entities: 1 },
        hot: [
            { id: 'm-1', target: 'user', importance: 'critical', origin: 'global', updatedAt: 'x', entities: [], tags: [], content: '偏好：书面化表达' },
            { id: 'm-2', target: 'project', importance: 'low', origin: 'workspace', updatedAt: 'x', entities: [], tags: [], content: '约定：只用 edit 工具' },
            { id: 'm-3', importance: 'normal', origin: 'workspace', updatedAt: 'x', entities: [], tags: [], content: '没有 target 的条目' },
        ],
    }));
    const groups = byProp(tree, 'data-hot-group');
    assert.deepEqual(groups.map((n) => n.props['data-hot-group']), ['user', 'project', 'other'], '认不出的 target 进「其它」');
    assert.ok(textsOf(groups[0]).some((t) => t === '用户偏好（1）'), '分组标题带条数');
    assert.ok(textsOf(groups[1]).some((t) => t === '项目与环境（1）'));

    const critical = byProp(tree, 'data-importance', 'critical')[0];
    assert.ok(critical, 'critical 的条目要标出来');
    // 侧色带改用 inset 阴影而不是 border-left：这是 mnemon 全站的手法（不改盒模型、
    // 也不会被圆角裁掉）。断言仍然钉住两件事：有 3px 色带，且颜色跟着重要度走。
    assert.ok(String(critical.props.style.boxShadow).includes('inset 3px 0 0'), '左侧画 3px 侧色带');
    assert.ok(String(critical.props.style.boxShadow).includes('state-error-primary'), 'critical 用危险色');
    assert.ok(String(byProp(tree, 'data-importance', 'normal')[0].props.style.boxShadow).includes('state-business-primary'));
    assert.ok(String(byProp(tree, 'data-importance', 'low')[0].props.style.boxShadow).includes('border-l3'));
});

await check('可视化：空快照照样渲染（指标 0、图谱空态），不出现 NaN / undefined', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 0, ledger: 0, archive: 0, links: 0, entities: 0, ledgerTotal: 0, archiveFiles: 0, archiveItems: 0 },
        stores: [],
        hot: [], ledger: [], archive: [], links: [], entities: [],
    }));

    assert.equal(byProp(tree, 'data-metric').length, 5, '空快照也保留五枚指标块');
    assert.deepEqual(byProp(tree, 'data-metric-value').map((n) => n.props.children[0]), ['0', '0', '0', '0', '0']);
    assert.equal(byProp(tree, 'data-panel', 'stores').length, 0, '没有 stores 就不画存储域');
    assert.equal(byProp(tree, 'data-gauge', 'ledger')[0].props['data-percent'], 0, '0 / 500 = 0%');
    assert.equal(byProp(tree, 'data-panel', 'ledger-timeline').length, 0, '没有台账就不画时间线');
    assert.equal(byProp(tree, 'data-entity').length, 0);

    const linksTree = openTab(tree, panel, '关系（0）');
    const graphPanel = byProp(linksTree, 'data-panel', 'graph')[0];
    assert.ok(graphPanel, '关系 Tab 仍要有图谱容器');
    assert.equal(byProp(graphPanel, 'data-graph', 'memory').length, 0, '没有节点就不画 SVG');
    assert.ok(textsOf(graphPanel).some((t) => t.includes('还没有节点或关系')));

    for (const rendered of [tree, linksTree]) {
        for (const text of stringsOf(rendered)) {
            assert.ok(!text.includes('NaN'), `渲染里出现了 NaN：${text}`);
            assert.ok(!text.includes('undefined'), `渲染里出现了 undefined：${text}`);
        }
    }
});

// ── 第十一轮：视觉与交互契约 ──────────────────────────────────────────────
//
// 这一轮把面板从「能用」推到「像 mnemon 一样」：设计系统令牌、分段控件、
// 卡片化指标、4px 容量条、右侧详情栏、以及一张渲染进树里的样式表。
// 下面钉的是**会被无声改坏的**几条：样式表节点还在不在、页签还认不认得出、
// 图谱标签会不会被画布裁掉、选中还看不看得出来。

await check('样式表：面板自带 <style> 节点，含关键帧与减少动效的兜底', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const styleNodes = walk(tree).filter((n) => n.type === 'style');
    assert.equal(styleNodes.length, 1, '面板要渲染恰好一个样式表节点');
    const css = String(styleNodes[0].props.dangerouslySetInnerHTML.__html || '');
    assert.ok(css.includes('@keyframes om-spin'), '加载转圈要靠关键帧（内联样式写不了）');
    assert.ok(css.includes('prefers-reduced-motion'), '动效要能被系统设置关掉');
    assert.ok(css.includes('@container'), '两栏塌成一列用容器查询（面板宽度 ≠ 窗口宽度）');
    assert.ok(css.includes('.om-card:hover'), '悬停态只改描边与阴影，不改几何');
    assert.ok(!css.includes('--dsh-'), '样式表里同样不许出现宿主没有的那套变量');
});

await check('版式：分段控件的页签带层色点，且页签文字仍是整串（含条数）', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const tabs = byProp(tree, 'data-tab');
    assert.deepEqual(tabs.map((n) => n.props['data-tab']), ['hot', 'ledger', 'archive', 'links', 'entities']);
    assert.ok(tabs.every((n) => n.props.role === 'tab'), '页签要有 tab 语义');
    assert.ok(tabs.every((n) => Array.isArray(n.props.children) && n.props.children.length === 2),
        '每个页签 = 层色点 + 整串文字（拆开就丢了条数）');
    const active = tabs.filter((n) => n.props['aria-selected'] === true);
    assert.equal(active.length, 1, '同时只有一个选中页签');
    assert.equal(active[0].props.style.fontWeight, 600, '选中的页签要加粗');
    // 面板根声明 container-type：否则容器查询不生效，窄面板仍会挤成两栏。
    const root = walk(tree).find((n) => n.props && typeof n.props.className === 'string'
        && n.props.className.split(' ').indexOf('om-mem') !== -1);
    assert.ok(root, '面板根要有 om-mem 类名');
    assert.equal(root.props.style.containerType, 'inline-size');
});

await check('面板本身是唯一的滚动容器（宿主那一格不滚，也不给滚动）', async () => {
    // 主面板槽位的宿主容器是 display:flex + overflow:hidden，槽位的 div 是
    // display:contents（不产生盒子）—— 所以本面板根元素就是那个 flex 项，
    // 它必须自己撑满高度并自己滚。第十三轮记的「ITEM DETAIL 落在首屏之外」
    // 就是这一条缺了造成的。
    const { tree } = await panelFrom(makeSnapshot());
    const root = walk(tree).find((n) => n.props && typeof n.props.className === 'string'
        && n.props.className.split(' ').indexOf('om-scroll') !== -1 && n.props['data-memory-panel'] === 'scroll');
    assert.ok(root, '面板根要是滚动容器（om-scroll）');
    assert.equal(root.props.style.height, '100%');
    assert.equal(root.props.style.overflowY, 'auto');
    assert.equal(root.props.style.minHeight, 0);
    assert.equal(root.props.style.overscrollBehavior, 'contain');
    // 头部吸顶：面板自己滚之后，搜索框与「刷新 / 记一条」不能跟着滑出视野。
    const head = walk(tree).find((n) => n.props && n.props.style && n.props.style.position === 'sticky'
        && n.props.style.top === 0);
    assert.ok(head, '面板头部要吸顶（position:sticky; top:0）');
    assert.ok(String(head.props.style.background).includes('bg-base'), '吸顶块要有底色，否则内容会透过去');
    // 列表**不再**自己滚：套两层滚动条会让「滚到一半卡住、要换个滚动条接着滚」。
    const lists = walk(tree).filter((n) => n.props && n.props['data-memory-list'] !== undefined);
    assert.ok(lists.length > 0, '列表还在');
    for (const list of lists) {
        assert.equal(list.props.style.maxHeight, undefined, '列表不该再有内层封顶高度');
        assert.equal(list.props.style.overflow, undefined, '列表不该再自己滚动');
    }
});

await check('图谱：滚轮只在按住修饰键时缩放，普通滚轮留给面板滚动', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graphOf = (rendered) => byProp(rendered, 'data-graph', 'memory')[0];
    const graph = graphOf(linksTree);
    const plain = { deltaY: -100, clientX: 0, clientY: 0, currentTarget: null };
    graph.props.onWheel(plain);
    assert.equal(graphOf(render(panel)).props['data-zoom'], '100', '不按修饰键时不该缩放（页面要能滚）');
    graph.props.onWheel(Object.assign({ ctrlKey: true }, plain));
    assert.equal(graphOf(render(panel)).props['data-zoom'], '115', 'Ctrl + 滚轮上滚放大一格（1.15）');
});

await check('指标卡：数字用等宽字、卡片走抬升表面与悬停态', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const values = byProp(tree, 'data-metric-value');
    assert.equal(values.length, 5);
    for (const value of values) {
        assert.match(String(value.props.style.fontFamily), /mono/, '指标数字要用等宽字（五枚并排才对得齐）');
        assert.deepEqual(value.props.children, ['1']);
    }
    for (const tile of byProp(tree, 'data-metric')) {
        assert.ok(String(tile.props.style.background).includes('bg-layer-1'), '指标卡要落在抬升表面上');
        assert.ok(String(tile.props.className).includes('om-card'), '指标卡要吃到悬停态');
    }
});

await check('容量条：4px 轨道 + 继承圆角 + 百分比，且每条一个语义色', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        counts: {
            hot: 1, ledger: 1, archive: 1, links: 1, entities: 1,
            ledgerTotal: 1, archiveFiles: 1, archiveItems: 1,
            hotBytes: 1024, projectBytes: 5120,
        },
    }));
    const fills = byProp(tree, 'data-gauge-fill');
    assert.equal(fills.length, 4);
    for (const fill of fills) {
        assert.equal(fill.props.style.height, '100%');
        assert.equal(fill.props.style.borderRadius, 'inherit', '填充要继承轨道圆角，否则右端露直角');
        assert.ok(String(fill.props.style.background).length > 0, '每条一个语义色');
    }
    assert.deepEqual(byProp(tree, 'data-gauge-percent').map((n) => n.props.children[0]),
        ['0.2%', '1.7%', '25%', '50%'], '百分比与 data-percent 同源');
    const gauge = byProp(tree, 'data-gauge', 'ledger')[0];
    const track = walk(gauge).find((n) => n.props && n.props.style && n.props.style.height === '4px');
    assert.ok(track, '轨道按 mnemon 的口径做细（4px）');
});

await check('容量口径：归档摘要用**摘要文件数**配 archiveKeep，不是条目数', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        // 一个月一个摘要文件：30 条落在 1 个文件里。上限 archiveKeep 管的是文件数。
        counts: { hot: 1, ledger: 1, archive: 30, links: 1, entities: 1, ledgerTotal: 1, archiveFiles: 1, archiveItems: 30 },
        archive: Array.from({ length: 30 }, (item, index) => ({
            id: 'A-' + index, month: '2026-09', kind: 'hot', at: 'x', origin: 'workspace',
            importance: 'normal', category: '', entities: [], tags: [], text: '归档条目 ' + index,
        })),
    }));
    const gauge = byProp(tree, 'data-gauge', 'archive')[0];
    assert.ok(gauge, '应有归档容量条');
    assert.equal(gauge.props['data-percent'], 1.7, '1 / 60 = 1.7% —— 不是 30 / 60 = 50%');
    assert.ok(textsOf(gauge).some((t) => t.includes('1 / 60')), '要写成「1 / 60 个」');
    assert.ok(!textsOf(gauge).some((t) => t.includes('30 / 60')), '不能拿条目数配文件上限');
    assert.equal(byProp(gauge, 'data-gauge-fill', 'archive')[0].props.style.width, '1.7%');
    // 条目数并没有消失：页签与指标卡照旧按条目数报。
    assert.ok(textsOf(tree).some((t) => t === '归档（30）'), '页签仍报条目数');
    assert.equal(byProp(tree, 'data-metric-value').map((n) => n.props.children[0])[2], '30', '指标卡仍报条目数');
    // 归档页签的按月时间线也不受影响。
    const archiveTree = openTab(tree, panel, '归档（30）');
    assert.deepEqual(byProp(archiveTree, 'data-bucket').map((n) => n.props['data-count']), [30]);
});

await check('容量口径：台账条数用未过滤总数，不跟着搜索框走', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        // 搜过之后列表里只剩 3 条，但库里其实有 25 条 —— 容量条讲的是后者。
        counts: { hot: 1, ledger: 3, archive: 1, links: 1, entities: 1, ledgerTotal: 25, archiveFiles: 1, archiveItems: 1 },
    }));
    const gauge = byProp(tree, 'data-gauge', 'ledger')[0];
    assert.ok(gauge, '应有台账容量条');
    assert.equal(gauge.props['data-percent'], 5, '25 / 500 = 5%');
    assert.ok(textsOf(gauge).some((t) => t.includes('25 / 500')), '容量条要报未过滤的总数');
    assert.ok(!textsOf(gauge).some((t) => t.includes('3 / 500')), '不能跟着过滤结果缩水');
});

await check('存储域卡片：归档条目数与摘要文件数分开显示', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        stores: [{ id: 'workspace', dir: '.office/memory', hot: 1, ledger: 1, archive: 30, archiveFiles: 1, links: 1, bytes: 2048, files: 7 }],
    }));
    const card = byProp(tree, 'data-store', 'workspace')[0];
    const texts = textsOf(card);
    assert.ok(texts.includes('归档 30'), '归档 chip 报条目数');
    assert.ok(texts.includes('摘要 1 个'), '另给一枚摘要文件数的 chip');
});

await check('关系图：贴边节点的标签往回收，且标签描一圈表面色', async () => {
    const hot = [];
    for (let index = 0; index < 14; index += 1) {
        hot.push({ id: 'm-' + index, target: 'project', importance: 'normal', origin: 'workspace', updatedAt: 'x', entities: [], tags: [], content: '很长的条目名字用于把标签撑开 ' + index });
    }
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 14, ledger: 1, archive: 1, links: 1, entities: 1 },
        hot: hot,
    }));
    const graph = byProp(openTab(tree, panel, '关系（1）'), 'data-graph', 'memory')[0];
    const nodes = walk(graph).filter((n) => n.type === 'g' && n.props && n.props['data-node']);
    const labels = byProp(graph, 'data-label');
    assert.ok(labels.length > 0, '要有节点标签');
    const anchors = new Set(labels.map((n) => n.props.textAnchor));
    assert.ok([...anchors].every((a) => ['start', 'middle', 'end'].includes(a)), `textAnchor 三选一：${[...anchors]}`);
    for (const node of nodes) {
        const point = parseTranslate(node.props.transform);
        assert.ok(point.x >= 78 && point.x <= 642, `节点要离画布边更远一点（标签才不会被裁）：${JSON.stringify(point)}`);
    }
    for (const label of labels) {
        assert.ok(String(label.props.paintOrder).includes('stroke'), '标签要描一圈表面色，压在网格上也读得清');
    }
});

await check('关系图：选中画虚线环、详情栏给出层 / 邻居 / 状态', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graph = byProp(linksTree, 'data-graph', 'memory')[0];
    assert.equal(byProp(graph, 'data-shape', 'ring').length, 0, '没选中时不画环');

    const node = walk(graph).find((n) => n.type === 'g' && n.props && n.props['data-node'] === 'hot:m-1');
    node.props.onClick();
    const selectedTree = render(panel);
    const selected = byProp(selectedTree, 'data-graph', 'memory')[0];
    assert.equal(byProp(selected, 'data-shape', 'ring').length, 1, '选中的节点要有虚线环');
    // 详情栏是画布的**兄弟**节点（两栏布局），要在整棵树里找，不能只在 SVG 里找。
    const detail = byProp(selectedTree, 'data-graph-detail', 'hot:m-1')[0];
    assert.ok(detail, '详情栏应指向 hot:m-1');
    const texts = textsOf(detail);
    assert.ok(texts.some((t) => t.includes('已选中')), `详情栏要标出选中态：${JSON.stringify(texts)}`);
    for (const key of ['层', '邻居', '状态']) assert.ok(texts.includes(key), `详情栏要有「${key}」这一行`);
    assert.ok(texts.includes('热记忆'), '详情栏要写层名');
});

await check('主题令牌：不再用宿主里不存在的 --dsh-*，改用 --dsw-alias-* 且保留字面兜底', () => {
    const source = readFileSync(BUNDLE, 'utf8');
    // 宿主主题只定义 --dsw-alias-*；--dsh-border / --dsh-accent / --dsh-input-bg /
    // --dsh-danger 在 52 个 dsh-client-ui-* 包里一个都没有 —— 用它们等于永远走兜底。
    assert.ok(!source.includes('--dsh-'), '不能用宿主里不存在的 --dsh-* 变量');
    for (const token of [
        '--dsw-alias-border-l1', '--dsw-alias-border-l2', '--dsw-alias-state-business-primary',
        '--dsw-alias-state-error-primary', '--dsw-alias-interactive-bg-hover',
    ]) {
        assert.ok(source.includes(token), `应使用宿主真实令牌 ${token}`);
    }
    // 兜底必须还在：令牌缺失时不能变成透明边框 / 无背景。
    assert.ok(source.includes('var(--dsw-alias-state-error-primary, #e5484d)'));
    assert.ok(source.includes('var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))'));
});

await check('回合记忆条：保留边框与背景（S.turnBar 曾被同名键覆盖成无框）', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const bar = regOf(registrations, 'conversation.chat.turnTail', 'office-memory-turnbar');
    const tree = render(bar, {
        turn: { turn: 3 },
        seq: 12,
        openFile: () => {},
        useTrajectory: (selector) => selector(trajectorySnapshot()),
    });
    assert.ok(tree, '应渲染出记忆条');
    assert.ok(String(tree.props.style.border).includes('1px solid'), '回合条要有边框');
    assert.ok(String(tree.props.style.background).length > 0, '回合条要有背景');
    assert.equal(tree.props.style.borderRadius, '8px');
    assert.equal(tree.props.style.marginBottom, '8px');
});

// ── 新增：会话内两处 ──────────────────────────────────────────────────────

await check('回合记忆条：读不到轨迹时返回 null，不崩', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const bar = regOf(registrations, 'conversation.chat.turnTail', 'office-memory-turnbar');
    // 真实 owner props 的形状（TurnLocation）。没有 useTrajectory（测试替身或部署里
    // 没挂 Trajectory）时返回 null —— 画一条假数据或崩掉整条回合尾都更糟。
    const tree = render(bar, { turn: { turn: 3, steps: [] }, seq: 12, openFile: () => {} });
    assert.equal(tree, null, '读不到轨迹就返回 null');
});

/** 造一份轨迹快照：eventNodes 是已完成的调用，runningCalls 是进行中的。 */
function trajectorySnapshot() {
    return {
        eventNodes: [
            { kind: 'tool-result', seq: 5, call: { name: 'office_memory', argsRaw: JSON.stringify({ action: 'read', query: '季度' }) } },
            { kind: 'tool-result', seq: 6, call: { name: 'office_memory', argsRaw: JSON.stringify({ action: 'add', content: 'x' }) } },
            { kind: 'tool-result', seq: 7, call: { name: 'office_search_brief', argsRaw: '{}' } },
            // 下面这些都不该被算进第 3 回合
            { kind: 'tool-result', seq: 8, call: { name: 'office_memory', argsRaw: JSON.stringify({ action: 'read' }) } },
            { kind: 'tool-result', seq: 9, call: { name: 'read', argsRaw: '{}' } },
            { kind: 'tool-result', seq: 10, call: null },
            { kind: 'assistant-step', seq: 11 },
        ],
        eventLocations: new Map([
            [5, { kind: 'step', turn: { turn: 3 }, step: { turn: 3, step: 1 } }],
            [6, { kind: 'turn', turn: { turn: 3 } }],
            [7, { kind: 'step', turn: { turn: 3 } }],
            [8, { kind: 'step', turn: { turn: 2 } }],
            [9, { kind: 'step', turn: { turn: 3 } }],
            [10, { kind: 'step', turn: { turn: 3 } }],
            [11, { kind: 'step', turn: { turn: 3 } }],
        ]),
        runningCalls: [
            { name: 'office_memory', argsRaw: JSON.stringify({ action: 'add' }), turn: 3 },
            { name: 'office_memory', argsRaw: '{}', turn: 2 },
        ],
    };
}

await check('回合记忆条：按本回合统计召回 / 沉淀 / 检索，别的回合与非办公工具都不算', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const bar = regOf(registrations, 'conversation.chat.turnTail', 'office-memory-turnbar');
    const snapshot = trajectorySnapshot();
    const tree = render(bar, {
        turn: { turn: 3 },
        seq: 12,
        openFile: () => {},
        useTrajectory: (selector) => selector(snapshot),
    });
    assert.ok(tree, '本回合有活动时应渲染出记忆条');
    const text = textsOf(tree).join(' ');
    assert.ok(text.includes('召回 1'), `已完成的 read 算召回；实际：${text}`);
    assert.ok(text.includes('沉淀 2'), `已完成 + 进行中的写入算沉淀；实际：${text}`);
    assert.ok(text.includes('检索 1'), `office_search_* 算检索；实际：${text}`);
});

await check('回合记忆条：本回合没有记忆活动时返回 null（不在每个回合尾加噪声）', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const bar = regOf(registrations, 'conversation.chat.turnTail', 'office-memory-turnbar');
    const snapshot = trajectorySnapshot();
    const tree = render(bar, {
        turn: { turn: 99 },
        seq: 12,
        openFile: () => {},
        useTrajectory: (selector) => selector(snapshot),
    });
    assert.equal(tree, null, '没有活动就不画');
});

await check('回合记忆条：选择器返回的是快照本身（SnapshotSelectorHook 才能收敛）', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const bar = regOf(registrations, 'conversation.chat.turnTail', 'office-memory-turnbar');
    const snapshot = trajectorySnapshot();
    let seen = null;
    render(bar, {
        turn: { turn: 3 },
        seq: 12,
        openFile: () => {},
        useTrajectory: (selector) => {
            seen = selector;
            return selector(snapshot);
        },
    });
    assert.equal(typeof seen, 'function', '组件应通过 useTrajectory 取数据');
    // 关键：选择器不能每次返回新建对象（那样宿主的重渲染永远收敛不了）。
    assert.equal(seen(snapshot), snapshot, '选择器必须返回快照里的既有引用');
});

await check('存入记忆按钮：只设草稿、不提交（受监督入口）', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const save = regOf(registrations, 'conversation.chat.assistant-actions', 'office-memory-save');
    const calls = [];
    const inputActions = {
        setDraft: (text) => calls.push(['setDraft', text]),
        submit: () => calls.push(['submit']),
    };
    const tree = render(save, { messageId: 'msg-1', inputActions: inputActions });
    assert.ok(tree, '应渲染出按钮');
    assert.equal(tree.type, 'button');
    assert.ok(textsOf(tree).some((t) => t === '存入记忆'));

    tree.props.onClick();
    assert.equal(calls.length, 1, '只该有 setDraft 一次调用');
    assert.equal(calls[0][0], 'setDraft');
    assert.ok(calls[0][1].includes('office_memory'), '草稿应点名用 office_memory 记');
    assert.ok(!calls.some((call) => call[0] === 'submit'), '绝不能自动提交：要用户看过再发');
});

await check('存入记忆按钮：拿不到 inputActions 时返回 null', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const save = regOf(registrations, 'conversation.chat.assistant-actions', 'office-memory-save');
    assert.equal(render(save, { messageId: 'msg-1' }), null);
    assert.equal(render(save, { messageId: 'msg-1', inputActions: {} }), null, '有 inputActions 但没有 setDraft 也要返回 null');
});

// ── 第十三轮：写入指令 / 分页 / 完整内容 / 视图变换 ──────────────────────
//
// 这一轮补的是第十二轮文档 §5 里列出的五条「已知边界」。下面钉的都是**会被无声
// 改坏**的东西：面板还能不能生成精确的 office_memory 指令、列表还能不能翻页、
// 完整内容还靠不靠原生 title、图谱的视图变换是不是与布局解耦、
// reduced-motion 管不管得住滚动行为。

await check('写入指令：条目动作把一条精确的 office_memory 调用放进待办区', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    assert.deepEqual(byProp(tree, 'data-memory-action').map((n) => n.props['data-memory-action']),
        ['replace', 'remove'], '热记忆条目上有「改写 / 忘记」两个动作');

    byProp(tree, 'data-memory-action', 'remove')[0].props.onClick();
    const withPending = render(panel);
    const card = byProp(withPending, 'data-memory-pending', 'remove')[0];
    assert.ok(card, '点动作后应出现待办指令卡');
    const instruction = byProp(card, 'data-memory-instruction')[0];
    assert.ok(instruction, '待办卡要给出可复制的指令正文');
    const text = String(instruction.props.children);
    assert.ok(text.includes('office_memory('), '指令必须点名 office_memory');
    assert.ok(text.includes('"remove"'), '忘记 → action:"remove"');
    assert.ok(text.includes('"user"'), 'target 跟着条目走（这条是 user）');
    assert.ok(text.includes('oldText'), '删除靠 oldText 唯一命中');
    assert.ok(text.includes('偏好：书面化表达'), 'oldText 要带上这条的原文片段');
    assert.ok(!text.includes('/office-memory/'), '不许把写入做成 HTTP 端点（信任边界不动）');
    assert.ok(textsOf(withPending).some((t) => t.includes('面板不写记忆')), '待办卡要写清「面板不落盘」');

    byProp(withPending, 'data-memory-dismiss')[0].props.onClick();
    assert.equal(byProp(render(panel), 'data-memory-pending').length, 0, '关闭后待办区消失');
});

await check('写入指令：复制进剪贴板、切到会话走 layout 服务', async () => {
    const copied = [];
    const jumped = [];
    const { tree, panel } = await panelFrom(makeSnapshot(), {
        navigator: { clipboard: { writeText: (value) => { copied.push(value); return Promise.resolve(); } } },
        layout: { selectPanel: (id) => { jumped.push(id); } },
    });

    byProp(tree, 'data-memory-action', 'replace')[0].props.onClick();
    const withPending = render(panel);
    assert.equal(byProp(withPending, 'data-memory-jump').length, 1, '拿得到 layout 时给「切到会话」');
    byProp(withPending, 'data-memory-jump')[0].props.onClick();
    assert.deepEqual(jumped, ['conversation'], '切到会话 = layout.selectPanel("conversation")');

    byProp(withPending, 'data-memory-copy')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 5));
    assert.equal(copied.length, 1, '复制应把指令正文交给剪贴板');
    assert.ok(copied[0].includes('office_memory('), '复制的是指令正文');
    assert.ok(copied[0].includes('"replace"'));
    const afterCopy = render(panel);
    assert.ok(textsOf(afterCopy).some((t) => t.includes('已复制')), '复制成功后按钮改成「已复制」');
});

await check('写入指令：拿不到剪贴板或 layout 时如实降级，不画死按钮', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    assert.equal(byProp(tree, 'data-memory-jump').length, 0, '没有 layout 服务时不画「切到会话」');

    byProp(tree, 'data-memory-compose')[0].props.onClick();
    const withPending = render(panel);
    assert.equal(byProp(withPending, 'data-memory-pending', 'add').length, 1, '「记一条」给出 add 骨架');
    const text = String(byProp(withPending, 'data-memory-instruction')[0].props.children);
    assert.ok(text.includes('"add"'));
    assert.ok(text.includes('"project"'), '默认记到项目与环境');

    byProp(withPending, 'data-memory-copy')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 5));
    assert.ok(textsOf(render(panel)).some((t) => t.includes('剪贴板权限')), '没有剪贴板时提示手动复制');
});

await check('写入指令：归档是只读层，给说明而不是给动作', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 1, ledger: 1, archive: 2, links: 1, entities: 1 },
        archive: [
            { id: 'A-1', month: '2026-09', kind: 'mnemon', at: 'x', origin: 'workspace', importance: 'normal', category: 'fact', entities: [], tags: [], text: '长期记忆一条' },
            { id: 'A-2', month: '2026-08', kind: 'mnemon', at: 'x', origin: 'workspace', importance: 'low', category: 'fact', entities: [], tags: [], text: '更旧的一条' },
        ],
    }));
    const archiveTree = openTab(tree, panel, '归档（2）');
    assert.equal(byProp(archiveTree, 'data-memory-action').length, 0, '归档条目没有写入动作');
    assert.equal(byProp(archiveTree, 'data-archive-note').length, 1, '归档要有一句只读说明');
    assert.ok(textsOf(archiveTree).some((t) => t.includes('归档是只读层')));

    // 台账 / 关系 / 实体各自的动作也要在（这三层是可写的）
    const ledgerTree = openTab(archiveTree, panel, '台账（1）');
    assert.deepEqual(byProp(ledgerTree, 'data-memory-action').map((n) => n.props['data-memory-action']), ['log']);
    const linksTree = openTab(ledgerTree, panel, '关系（1）');
    assert.deepEqual(byProp(linksTree, 'data-memory-action').map((n) => n.props['data-memory-action']), ['unlink']);
    const entityTree = openTab(linksTree, panel, '实体（1）');
    assert.deepEqual(byProp(entityTree, 'data-memory-action').map((n) => n.props['data-memory-action']), ['related']);
});

await check('分页：默认只铺一页，「再显示 N 条」才继续，全显示后给「收起」', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 25, ledger: 1, archive: 1, links: 1, entities: 1 },
        hot: manyHot(25),
    }));
    assert.equal(byProp(tree, 'data-item-key').length, 20, '默认只铺一页（20 条）');
    const more = byProp(tree, 'data-show-more')[0];
    assert.ok(more, '还有 5 条没显示时应给「再显示 N 条」');
    assert.deepEqual(more.props.children, ['再显示 5 条'], '按钮写清这一次会多显示几条');
    assert.ok(textsOf(tree).some((t) => t.includes('已显示 20 / 25')), '要写清已显示 / 总数');
    // 分组标题用**整组总数**，不随分页变小：否则越点数字越小，读起来像条目被删了。
    assert.ok(textsOf(tree).some((t) => t === '用户偏好（13）'), '分组标题是整组总数（user 13 条）');
    assert.ok(textsOf(tree).some((t) => t === '项目与环境（12）'), '分组标题是整组总数（project 12 条）');

    more.props.onClick();
    const expanded = render(panel);
    assert.equal(byProp(expanded, 'data-item-key').length, 25, '点一次就把剩下的都铺出来');
    assert.equal(byProp(expanded, 'data-show-more').length, 0, '全显示后不再给「再显示」');
    assert.equal(byProp(expanded, 'data-collapse-list').length, 1, '全显示后给「收起」');
    assert.ok(textsOf(expanded).some((t) => t.includes('已显示 25 / 25')));

    byProp(expanded, 'data-collapse-list')[0].props.onClick();
    assert.equal(byProp(render(panel), 'data-item-key').length, 20, '收起回到第一页');
});

await check('分页：条目数不超过一页时不出现分页条', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    assert.equal(byProp(tree, 'data-list-more').length, 0, '短列表不该多一条分页栏');
});

await check('完整内容：列表项不再用原生 title，改为「展开」与底部详情条', async () => {
    const long = '偏好：' + '书面化表达'.repeat(30);
    const { tree, panel } = await panelFrom(makeSnapshot({
        hot: [{
            id: 'm-1', target: 'user', importance: 'critical', origin: 'global',
            updatedAt: 'x', entities: [], tags: [], content: long,
        }],
    }));
    assert.equal(walk(tree).filter((n) => n.props && typeof n.props.title === 'string' && n.props.title.includes('书面化表达')).length, 0,
        '列表项不许再用原生 title 承载完整内容');

    const expand = byProp(tree, 'data-item-expand', 'm-1')[0];
    assert.ok(expand, '长条目要有「展开」');
    assert.equal(expand.props['aria-expanded'], false);
    assert.ok(String(byProp(tree, 'data-item-body', 'm-1')[0].props.style.maxHeight).includes('em'), '默认限高折叠');

    expand.props.onClick();
    const opened = render(panel);
    assert.equal(byProp(opened, 'data-item-expand', 'm-1')[0].props['aria-expanded'], true, '展开后 aria-expanded 变 true');
    assert.deepEqual(byProp(opened, 'data-item-expand', 'm-1')[0].props.children, ['收起']);
    const body = byProp(opened, 'data-item-body', 'm-1')[0];
    assert.equal(body.props.style.maxHeight, undefined, '展开后去掉限高');
    assert.equal(body.props['data-item-open'], 'true');

    // 详情条：悬停与键盘聚焦走同一套状态，所以两条路都读得到全文
    byProp(opened, 'data-item-key', 'm-1')[0].props.onMouseEnter();
    const hovered = render(panel);
    const detail = byProp(hovered, 'data-item-detail', 'm-1')[0];
    assert.ok(detail, '悬停后底部详情条指向该条目');
    assert.ok(textsOf(detail).some((t) => t === long), '详情条给的是完整内容');
    assert.equal(detail.props.role, 'status', '详情条用 status 语义');
    byProp(hovered, 'data-item-key', 'm-1')[0].props.onBlur();
    assert.equal(byProp(render(panel), 'data-item-detail').length, 0, '移开 / 失焦后详情条收起');
});

await check('完整内容：短条目不给「展开」按钮（折叠与展开没差别）', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    assert.equal(byProp(tree, 'data-item-expand').length, 0, '短条目不该多一个按钮');
});

await check('图谱：滚轮以光标为锚点缩放，视图变换不落在节点上', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graphOf = (rendered) => byProp(rendered, 'data-graph', 'memory')[0];
    const nodeTransforms = (rendered) => walk(graphOf(rendered))
        .filter((n) => n.type === 'g' && n.props && n.props['data-node'])
        .map((n) => n.props.transform);

    assert.equal(graphOf(linksTree).props['data-zoom'], '100', '默认 100%');
    assert.equal(graphOf(linksTree).props['data-pan'], '0,0', '默认没有平移');
    const nodesBefore = nodeTransforms(linksTree);

    graphOf(linksTree).props.onWheel({
        deltaY: -100, clientX: 360, clientY: 190, ctrlKey: true,
        currentTarget: { getBoundingClientRect: () => ({ left: 0, top: 0, width: 720, height: 380 }) },
    });
    const zoomed = graphOf(render(panel));
    assert.equal(zoomed.props['data-zoom'], '115', 'Ctrl + 滚轮上滚放大一格（1.15）');
    const view = walk(zoomed).find((n) => n.props && n.props['data-graph-view']);
    assert.ok(view, '视图变换要落在一个专门的组上');
    assert.ok(String(view.props.transform).includes('scale(1.15)'), '缩放写在视图组的 transform 上');
    assert.ok(String(view.props.transform).includes('translate(-54 -28.5)'),
        '以光标为锚点：光标底下那个点在缩放前后不动');
    assert.deepEqual(nodeTransforms(render(panel)), nodesBefore,
        '节点的 translate 仍是布局坐标，不随视图变');

    // 滚轮缩放的上下限也要钉住：一路放大不该超过 300%
    for (let index = 0; index < 30; index += 1) {
        graphOf(render(panel)).props.onWheel({ deltaY: -100, clientX: 0, clientY: 0, ctrlKey: true, currentTarget: null });
    }
    assert.equal(graphOf(render(panel)).props['data-zoom'], '300', '放大封顶 300%');
});

await check('图谱：拖背景平移、拖节点挪位置，两者互不串台', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graphOf = (rendered) => byProp(rendered, 'data-graph', 'memory')[0];
    const nodeOf = (rendered) => walk(graphOf(rendered)).find((n) => n.type === 'g' && n.props && n.props['data-node'] === 'hot:m-1');

    const pan = byProp(linksTree, 'data-graph-pan')[0];
    assert.ok(pan, '要有整块画布的平移面');
    assert.equal(pan.props.style.cursor, 'grab', '背景是 grab 光标（mnemon 的抓取手势）');

    pan.props.onPointerDown({ button: 0, clientX: 100, clientY: 100, pointerId: 1 });
    const pressed = graphOf(render(panel));
    assert.equal(pressed.props['data-dragging'], 'pan');
    assert.equal(byProp(pressed, 'data-graph-pan')[0].props.style.cursor, 'grabbing');
    // move / up 挂在画布上（指针捕获之后事件仍冒泡到画布，挂一处就够）
    pressed.props.onPointerMove({ clientX: 130, clientY: 88 });
    const moved = graphOf(render(panel));
    assert.equal(moved.props['data-pan'], '30,-12', '拖 30/-12 → 视图平移同样的量');
    moved.props.onPointerUp({ pointerId: 1 });
    assert.equal(graphOf(render(panel)).props['data-dragging'], undefined, '松手后不再标拖拽');

    // 拖节点：只挪那一个节点，视图不动
    const before = parseTranslate(nodeOf(render(panel)).props.transform);
    const panBefore = graphOf(render(panel)).props['data-pan'];
    // 往画布中心方向拖，避免撞上 clampGraphPoint 的边界而测不出位移。
    const stepX = before.x > 360 ? -14 : 14;
    const stepY = before.y > 190 ? -6 : 6;
    nodeOf(render(panel)).props.onPointerDown({ button: 0, clientX: 200, clientY: 200, pointerId: 2 });
    assert.equal(nodeOf(render(panel)).props['data-node-dragging'], 'true', '被拖的节点要标出来');
    graphOf(render(panel)).props.onPointerMove({ clientX: 200 + stepX, clientY: 200 + stepY });
    const after = parseTranslate(nodeOf(render(panel)).props.transform);
    assert.ok(Math.abs((after.x - before.x) - stepX) <= 0.2, `节点应横向移动 ${stepX}：${before.x} → ${after.x}`);
    assert.ok(Math.abs((after.y - before.y) - stepY) <= 0.2, `节点应纵向移动 ${stepY}：${before.y} → ${after.y}`);
    assert.equal(graphOf(render(panel)).props['data-pan'], panBefore, '拖节点不该带动整张画布');

    // 拖过之后那一次 click 不算选中（否则拖完会顺手改选中态）
    nodeOf(render(panel)).props.onClick();
    assert.equal(byProp(render(panel), 'data-graph-detail', 'none').length, 1, '拖过之后不选中');
});

await check('图谱：布局切换换坐标表并清掉手动位移，重置视图只复位视图', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graphOf = (rendered) => byProp(rendered, 'data-graph', 'memory')[0];
    const positions = (rendered) => walk(graphOf(rendered))
        .filter((n) => n.type === 'g' && n.props && n.props['data-node'])
        .map((n) => parseTranslate(n.props.transform));

    assert.equal(graphOf(linksTree).props['data-layout'], 'force', '默认是自然铺开');
    const force = positions(linksTree);
    assert.equal(byProp(linksTree, 'data-graph-layout', 'force')[0].props['aria-pressed'], true);

    byProp(linksTree, 'data-graph-layout', 'uniform')[0].props.onClick();
    const uniformTree = render(panel);
    assert.equal(graphOf(uniformTree).props['data-layout'], 'uniform');
    assert.equal(byProp(uniformTree, 'data-graph-layout', 'uniform')[0].props['aria-pressed'], true);
    const uniform = positions(uniformTree);
    assert.notDeepEqual(uniform, force, '均匀重置要真的换一张坐标表');
    assert.equal(new Set(uniform.map((p) => `${p.x},${p.y}`)).size, uniform.length, '均匀布局不能把节点叠在一起');
    for (const point of uniform) {
        assert.ok(point.x >= 78 && point.x <= 642 && point.y >= 44 && point.y <= 322,
            `均匀布局也不能越界：${JSON.stringify(point)}`);
    }

    // 缩放 + 平移之后「重置视图」回到 100% / 0,0，但布局与节点坐标都不动
    graphOf(uniformTree).props.onWheel({ deltaY: -100, clientX: 100, clientY: 100, ctrlKey: true, currentTarget: null });
    byProp(render(panel), 'data-graph-pan')[0].props.onPointerDown({ button: 0, clientX: 10, clientY: 10, pointerId: 3 });
    graphOf(render(panel)).props.onPointerMove({ clientX: 40, clientY: 25 });
    graphOf(render(panel)).props.onPointerUp({ pointerId: 3 });
    const movedTree = render(panel);
    assert.notEqual(graphOf(movedTree).props['data-pan'], '0,0', '拖过之后有平移量');
    assert.notEqual(graphOf(movedTree).props['data-zoom'], '100', '缩放后不是 100%');

    byProp(movedTree, 'data-graph-reset-view')[0].props.onClick();
    const reset = render(panel);
    assert.equal(graphOf(reset).props['data-zoom'], '100', '重置视图回到 100%');
    assert.equal(graphOf(reset).props['data-pan'], '0,0', '重置视图回到原点');
    assert.equal(graphOf(reset).props['data-layout'], 'uniform', '重置视图不该顺手改布局');
    assert.deepEqual(positions(reset), uniform, '重置视图不该动节点坐标');
});

await check('图谱：视图条上的缩放按钮与百分比读数跟着状态走', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const linksTree = openTab(tree, panel, '关系（1）');
    const graphOf = (rendered) => byProp(rendered, 'data-graph', 'memory')[0];
    assert.equal(byProp(linksTree, 'data-graph-zoom')[0].props.children[0], '100%');
    assert.equal(byProp(linksTree, 'data-graph-controls').length, 1, '画布上方要有视图条');
    assert.equal(byProp(linksTree, 'data-graph-controls')[0].props['data-graph-controls'], 'true');

    byProp(linksTree, 'data-graph-zoom-in')[0].props.onClick();
    assert.equal(byProp(render(panel), 'data-graph-zoom')[0].props.children[0], '115%');
    byProp(render(panel), 'data-graph-zoom-out')[0].props.onClick();
    assert.equal(byProp(render(panel), 'data-graph-zoom')[0].props.children[0], '100%');
    // 缩放按钮以中心为锚点，所以复位之后偏移也回到 0
    assert.equal(graphOf(render(panel)).props['data-pan'], '0,0', '中心锚点的放大缩小是自洽的');
});

await check('减少动效：reduced-motion 也压住 scroll-behavior，JS 平滑滚动同样读一次', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const styleNodes = walk(tree).filter((n) => n.type === 'style');
    assert.equal(styleNodes.length, 1);
    const css = String(styleNodes[0].props.dangerouslySetInnerHTML.__html || '');
    const at = css.indexOf('@media (prefers-reduced-motion:reduce)');
    assert.ok(at >= 0, '要有 reduced-motion 兜底块');
    const block = css.slice(at);
    assert.ok(block.includes('scroll-behavior:auto !important'),
        '减少动效必须覆盖 scroll-behavior —— 只压 transition / animation 时，滚动仍然平滑');
    assert.ok(block.includes('transition-duration:.01ms !important'), 'transition 的兜底不能丢');
    assert.ok(block.includes('animation-duration:.01ms !important'), 'animation 的兜底不能丢');
    assert.ok(block.includes('.om-mem,.om-mem *'), '覆盖整棵子树（宿主也可能在上层设过平滑滚动）');
    assert.ok(css.includes('.om-scroll{scroll-behavior:smooth}'), '列表本身设了平滑滚动，才有一条能被覆盖的声明');

    // JS 那一侧：scrollTo({behavior:'smooth'}) 不读 CSS 的 scroll-behavior，
    // 必须在代码里再读一次 matchMedia，否则「再显示 N 条」照样会平滑滚动。
    const source = readFileSync(BUNDLE, 'utf8');
    assert.ok(source.includes('prefersReducedMotion()'), '平滑滚动前要读一次系统设置');
    assert.ok(source.includes('(prefers-reduced-motion: reduce)'), '读的是 matchMedia 的减少动效查询');
    assert.ok(source.includes('"auto"') && source.includes('"smooth"'), '两条分支都在');
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`client: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);