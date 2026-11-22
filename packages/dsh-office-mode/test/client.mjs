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

import { resolveHostPath } from '../test/host-modules.mjs';
// 站点目录是**服务端**的模块：测试跑在 node 侧，可以直接 import 真源，
// 与 bundle 里镜像的那一份逐条比对（浏览器产物不能 import 它，只能镜像）。
import { BUILTIN_SITES, SITE_PRIORITY_LIMITS, SITE_TYPES } from '../src/site-catalog.js';

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
            // 真实 React 把函数初值当惰性初始化调用；替身也要一致，
            // 否则 useState(loadCustomPresets) 会把函数本身当状态。
            if (!store.has(key)) store.set(key, { value: typeof initial === 'function' ? initial() : initial, rerender: null });
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
        useEffect: (fn, deps) => {
            // 依赖数组按浅比较处理：变了就先跑上一次的 cleanup 再重跑（和 React 一致）。
            // 不传 deps 的调用仍然只跑挂载那一次（本替身不模拟「每次渲染都跑」）。
            const key = currentComponent + ':effect:' + cursor;
            cursor += 1;
            const nextDeps = Array.isArray(deps) ? deps : null;
            const prior = store.get(key);
            if (prior) {
                const priorDeps = Array.isArray(prior.deps) ? prior.deps : null;
                const unchanged = nextDeps !== null && priorDeps !== null
                    && nextDeps.length === priorDeps.length
                    && nextDeps.every((value, index) => Object.is(value, priorDeps[index]));
                if (nextDeps === null || unchanged) return;
                const cleanup = store.get(key + ':cleanup');
                if (cleanup && typeof cleanup.value === 'function') cleanup.value();
                store.delete(key + ':cleanup');
            }
            store.set(key, { value: true, deps: nextDeps });
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
        origin: '',
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
        stores: [{
            id: 'workspace', dir: '.office/memory', hot: 1, ledger: 1, archive: 1, archiveFiles: 1, links: 1,
            // 归档隔离提示只传字符串数组；默认夹具没有冲突时是空数组（冻结契约的一部分）。
            warnings: [],
            // 知识库计数（第四十二轮）：面板的存储卡按它画「知识库 N 篇 / M 块」。
            kb: { docs: 1, chunks: 3, bytes: 4096, tiers: { verified: 1, user: 0, unverified: 0 } },
        }],
        // counts 有两个口径（见 src/view.js 的注释）：hot/ledger/archive/links/entities/kb 是
        // 「列表里现在有几条」（过滤并截断之后），ledgerTotal/archiveFiles/archiveItems/kbTotal
        // 是容量与「库里一共多少」用的未过滤总数。默认夹具里两者相等，专门的口径用例会造出差异。
        counts: {
            hot: 1, ledger: 1, archive: 1, links: 1, entities: 1, kb: 1,
            ledgerTotal: 1, archiveFiles: 1, archiveItems: 1,
            kbTotal: 1, kbChunks: 3, kbBytes: 4096, kbTruncated: false,
        },
        // 来源筛选的三档计数（第八十二轮）：按当前过滤词算、不受来源筛选本身影响。
        originCounts: {
            hot: { all: 1, global: 1, workspace: 0 },
            ledger: { all: 1, global: 0, workspace: 1 },
            archive: { all: 1, global: 0, workspace: 1 },
            links: { all: 1, global: 0, workspace: 1 },
            kb: { all: 1, global: 0, workspace: 1 },
        },
        // 知识库下钻（第八十二轮）：只有请求里点名时才非 null。
        // 知识库文件夹树（第八十四轮）：从 manifest 现算，计数含子树。
        kbFolder: '', kbRecursive: true,
        kbTree: {
            path: '', name: '全部文档', docs: 1, chunks: 3, bytes: 4096,
            children: [{
                path: '来源', name: '来源', docs: 1, chunks: 3, bytes: 4096,
                children: [{ path: '来源/方法论', name: '方法论', docs: 1, chunks: 3, bytes: 4096, children: [] }],
            }],
        },
        kbSearch: null,
        kbSearchError: '',
        kbDoc: null,
        kbDocError: '',
        kbChunk: null,
        kbChunkError: '',
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
        // 知识库文档（kb 一期）：字段名照抄 src/view.js 里 kb[] 的契约。
        kb: [{
            id: 'kb:8f14e45f', path: '来源/方法论.md', title: '方法论笔记', hash: '8f14e45f',
            folder: '来源',
            bytes: 4096, chars: 2048, chunks: 3, tier: 'verified', at: '2026-09-30T10:00:00.000Z',
            origin: 'workspace',
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
    // 第三十五轮：设置页分组的开合状态记在 localStorage 里（拿不到就退回内存）。
    // 注入一个假的才能验「跨刷新恢复」与「坏记录不抛错」这两条。
    if (options.localStorage !== undefined) win.localStorage = options.localStorage;
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
    // ctx.inject 的每次调用都记下来：迟绑定用例要在挂载后模拟「remote 服务
    // 这时才就绪」，再触发真实的注入回调，而不是直接改模块内部变量。
    const injects = [];
    // 宿主把每个远程命名空间挂成**独立的扁平服务** remote.<ns>（settings /
    // session / commands），cordis 只允许在声明过该服务名的注入 fiber 里访问
    // `remote.<ns>`；从只 inject ['remote'] 的回调里顺手读 `remote.session`
    // 会抛「cannot get property "remote.session" without inject」（现场
    // .office/runtime/diag/web-verify.json:1138）。这里把同一套严格语义搬进
    // 测试：remote 替身只在对应的服务名被声明时才放行该命名空间，读别的当场
    // 抛 cordis 原话 —— 否则「在测试里能摸到、到宿主就炸」的回归还会漏过去。
    const REMOTE_NAMESPACES = ['settings', 'session', 'commands'];
    const remoteRuntime = {};
    if (options.remote !== undefined && options.remote !== null) {
        for (const name of REMOTE_NAMESPACES) {
            if (options.remote[name] !== undefined) remoteRuntime[name] = options.remote[name];
        }
    }
    const remoteAvailable = options.remote !== undefined && options.remote !== null;
    const remoteProxyFor = (declared) => new Proxy(remoteRuntime, {
        get(target, prop) {
            if (typeof prop === 'string' && REMOTE_NAMESPACES.includes(prop)) {
                if (!declared.has(prop)) {
                    throw new Error(`cannot get property "remote.${prop}" without inject`);
                }
                return target[prop];
            }
            return Reflect.get(target, prop);
        },
        has(target, prop) {
            if (typeof prop === 'string' && REMOTE_NAMESPACES.includes(prop)) {
                return declared.has(prop) && target[prop] !== undefined;
            }
            return prop in target;
        },
    });
    /**
     * 按严格 Cordis 语义跑一段 inject 回调：先卸载上一轮（如有），再核对该回调
     * 声明的服务；缺任何一个就当作「还没挂上」，不回调。回调返回的函数记作
     * disposer，供卸载/重挂用例触发。
     */
    const runInject = (entry) => {
        if (typeof entry.disposer === 'function') {
            try { entry.disposer(); } catch (error) { /* 卸载失败不影响后续断言 */ }
        }
        entry.disposer = undefined;
        const declared = new Set();
        let remoteProxy = null;
        const sub = {};
        for (const service of entry.services) {
            if (service === 'configForms') {
                if (options.forms === undefined) return false;
                sub.configForms = options.forms;
                continue;
            }
            if (service === 'layout') {
                if (options.layout === undefined) return false;
                sub.layout = options.layout;
                continue;
            }
            if (service === 'remote') {
                if (!remoteAvailable) return false;
                remoteProxy ??= remoteProxyFor(declared);
                sub.remote = remoteProxy;
                continue;
            }
            if (service.startsWith('remote.')) {
                if (!remoteAvailable) return false;
                const name = service.slice('remote.'.length);
                if (!REMOTE_NAMESPACES.includes(name)) return false;
                if (remoteRuntime[name] === undefined) return false;
                declared.add(name);
                sub[service] = remoteRuntime[name];
                remoteProxy ??= remoteProxyFor(declared);
                sub.remote = remoteProxy;
                continue;
            }
            return false;
        }
        const disposer = entry.callback(sub);
        entry.disposer = typeof disposer === 'function' ? disposer : undefined;
        entry.active = true;
        return true;
    };
    /** 模拟宿主异步挂上/重挂一个命名空间服务，并按 cordis 的语义重跑依赖它的回调。 */
    const mountRemoteNamespace = (name, value) => {
        remoteRuntime[name] = value;
        for (const entry of injects.slice()) {
            if (!entry.services.includes(`remote.${name}`)) continue;
            runInject(entry);
        }
    };
    const unmountRemoteNamespace = (name) => {
        remoteRuntime[name] = undefined;
        for (const entry of injects.slice().reverse()) {
            if (!entry.services.includes(`remote.${name}`)) continue;
            if (typeof entry.disposer === 'function') entry.disposer();
            entry.disposer = undefined;
            entry.active = false;
        }
    };
    const target = {
        effect: (fn) => { return fn(); },
        // 宿主客户端上下文的 ctx.inject：只在服务就绪时才回调（缺服务就永不回调，
        // 而不是抛错）。回调拿到的是带该服务的子上下文。
        // configForms 与 layout 都是**可选**服务：没传就当作「这个部署没有」，
        // 于是「拿不到服务时的降级路径」也能被测到。
        inject: (services, callback) => {
            const entry = {
                services: Array.isArray(services) ? services : [services],
                callback: callback,
                disposer: undefined,
                active: false,
            };
            injects.push(entry);
            runInject(entry);
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
    return {
        entry, mod, ctx, registrations, injected, fake, injects,
        remoteRuntime, mountRemoteNamespace, unmountRemoteNamespace,
    };
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

await check('注册成单页设置项：办公与记忆合并，内部八个页签', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const sections = registrations.filter((item) => item.opts.name === 'settings.section');
    assert.equal(sections.length, 1, '合并后只占一格导航；再注册第二格就白合并了');

    const section = settingsReg(registrations);
    assert.ok(section, '应有「办公与记忆」这一格');
    assert.equal(section.opts.name, 'settings.section');
    assert.equal(section.opts.id, 'dsh-office-mode');
    assert.equal(typeof section.opts.order, 'number');
    assert.equal(typeof section.opts.label, 'function', '导航项标题要用 thunk（支持本地化重取）');
    assert.equal(typeof section.render, 'function');

    const tree = render(withTab(section, 'overview'));
    const tabs = walk(tree).filter((node) => node.props && node.props['data-settings-tab']);
    assert.deepEqual(tabs.map((node) => node.props['data-settings-tab']),
        ['overview', 'office', 'search', 'docs', 'gallery', 'memory', 'kb', 'recall'], '八个页签的顺序即用户动线');
    assert.equal(tabs[0].props['aria-selected'], true, '默认停在总览');
    assert.deepEqual(tabs.map((node) => [].concat(node.props.children).join('')),
        ['总览', '办公', '检索', '文档', '图库', '记忆', '知识库', '召回']);
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
    assert.equal(sections.length, 1, '设置页拿不到服务也要照常注册（面板自己显示降级提示）');
    const panel = registrations.filter((item) => item.opts.name === 'conversation.view');
    // 记忆（40）与预览（50）两格都只依赖 slots：设置服务缺失时它们照样在。
    assert.equal(panel.length, 2, '记忆与预览两格与设置服务无关，必须照常注册');
    assert.ok(panel.some((item) => item.opts.id === 'office-memory'), '记忆面板必须照常注册');
    assert.ok(panel.some((item) => item.opts.id === 'office-preview'), '预览面板必须照常注册');
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
        kbDir: '',
        embedding: {
            mode: 'auto', baseURL: '', model: '', apiKeyEnv: '', dimensions: 0, batchSize: 16, timeoutMs: 15000,
            local: {
                enabled: false, command: '', modelDir: '', idleTimeoutMs: 120000,
                maxRssMb: 600, startupTimeoutMs: 60000, maxLength: 256, threads: 2,
            },
        },
        selector: {
            mode: 'off', baseURL: '', model: '', apiKeyEnv: '', timeoutMs: 12000,
            maxCandidates: 12, maxSelected: 5, headChars: 300,
        },
        ann: { enabled: true, m: 16, efConstruction: 200, efSearch: 64, maxBuild: 20000 },
        hybrid: { lexicalWeight: 0.55, vectorWeight: 0.45, candidateLimit: 30 },
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
        maxParallel: 4,
        resultLimit: 40,
        maxChannels: 12,
        outputDir: '.office/search',
        // 站点清单：schema 的默认值就是内置目录（见 src/site-catalog.js），这里给它的
        // 一个子集（四条、跨三组），行数才数得清；「恢复内置默认」另有断言对着
        // BUILTIN_SITES.length，两条互不替代。
        sites: {
            enabled: true,
            fallback: true,
            maxPerCall: 3,
            entries: [
                { type: 'academic', domain: 'arxiv.org', label: 'arXiv', note: '预印本，理工科一手材料', enabled: true },
                { type: 'academic', domain: 'sci-hub.se', label: 'Sci-Hub', note: '影子图书馆，默认关闭', enabled: false },
                { type: 'code', domain: 'github.com', label: 'GitHub', note: '仓库、Issue 与讨论', enabled: true },
                { type: 'custom', domain: 'example.com', label: '示例站', note: '自己加的站点', enabled: true },
            ],
        },
    },
    documents: {
        defaultTheme: 'plain',
        scriptTimeoutMs: 60000,
        maxScriptChars: 400000,
        cacheDir: '.office/runtime/cache',
        keepCache: false,
    },
    // 文件识别（OCR）组：样例刻意给**非默认**取值（环境变量名、精度、超时都改过），
    // 「界面读到的是设置里的值，不是兜底常量」才测得出。apiKey 故意**不给**：
    // 宿主每次远程读都会把 role('secret') 的字段删掉（redactSecrets），所以真实
    // 部署里这里永远是缺席 —— 界面不许依赖它，更不许回填它。
    ocr: {
        enabled: true,
        baseURL: 'https://api.deepseek.com',
        apiKeyEnv: 'OCR_TEST_KEY',
        // 复用 deepseek-official 凭据默认关，样例也保持关：界面必须如实显示「关」。
        reuseOfficialProvider: false,
        purpose: 'rich_layout_recognition',
        timeoutMs: 240000,
    },
    // Python 组：第三十八轮补进界面。样例给一份**非默认**的值（解释器写绝对路径、
    // 超时与产物目录都改过），这样「界面读到的是设置里的值，而不是兜底常量」才测得出。
    python: {
        enabled: true,
        bin: 'C:\\Python313\\python.exe',
        timeoutMs: 60000,
        outDir: 'python/out',
    },
    av: {
        enabled: true,
        ffmpegPath: '',
        ffprobePath: '',
        modelDir: '',
        language: 'zh',
        chunkSeconds: 120,
        overlapSeconds: 0.5,
        maxSeconds: 3600,
        timeoutMs: 300000,
        precision: 'int8',
        threads: 2,
        frames: 6,
        maxFrames: 12,
        frameFormat: 'jpg',
        frameWidth: 0,
    },
    // 图库组：视觉描述与图文对齐默认都关。apiKey / imageEmbedding.apiKey
    // 不给 —— 宿主读设置时会删掉 secret 字段，界面不能依赖它。
    gallery: {
        vision: {
            channel: 'off', provider: '', model: '', baseURL: '', apiKeyEnv: '',
            maxTokens: 900, timeoutMs: 60000, language: 'zh',
        },
        imageEmbedding: {
            provider: 'none', baseURL: '', model: '', apiKeyEnv: '',
            dimensions: 0, batchSize: 4, timeoutMs: 30000,
        },
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

/** 取合并后的设置页注册记录。 */
function settingsReg(registrations) {
    return registrations.find((item) => item.opts.id === 'dsh-office-mode');
}

/**
 * 给注册记录钉一个初始页签。
 *
 * 合并页默认停在总览；测试要针对某一组时，先进入对应页签，等价于用户点了一下。
 * 返回的是浅拷贝，不动注册记录本身（同一份 registrations 会被反复渲染）。
 */
function withTab(registration, tab) {
    if (!registration) return registration;
    return Object.assign({}, registration, {
        render: (props) => registration.render(Object.assign({}, props, { initialTab: tab })),
    });
}

/** 办公页签的注册记录：工具、插件、守卫三组。 */
function officeReg(registrations) {
    return withTab(settingsReg(registrations), 'office');
}

/** 记忆页签的注册记录：记忆、层拓扑、范围、主动记录四组。 */
function memoryReg(registrations) {
    return withTab(settingsReg(registrations), 'memory');
}

/** 渲染任意页签：renderOfficeTab(regs, 'docs') / 'search' / 'kb' / 'recall'。 */
function renderOfficeTab(registrations, tab) {
    return render(withTab(settingsReg(registrations), tab));
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
    const tree = renderOfficeTab(registrations, 'search');
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
    const tree = renderOfficeTab(registrations, 'docs');
    const select = controlOf(findByLabel(tree, '默认主题'));
    assert.equal(select.type, 'select');
    assert.equal(select.props.value, 'plain');
    // React 里 option 是 select 的子元素（不是 options 属性）。
    const options = walk(select.props.children).filter((node) => node.type === 'option');
    assert.equal(options.length, 7, '应渲染全部七个主题');
    const ids = options.map((node) => node.props.value);
    assert.ok(ids.includes('plain') && ids.includes('business') && ids.includes('forest'), '应包含服务端主题');
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

// ── 设置写入：revision 协议与宿主结构化拒绝 ────────────────────────────────
//
// 第九十一轮返工：写入前先 settings.describe() 取 revision，再带 revision 调
// mutate；只有 settings/conflict（宿主按契约拒写、没落盘）才重读 revision 重试，
// 其余拒绝一次都不重试，并把宿主 error.code / message 原样显示出来。

/** 假 remote.settings：describe 依次给 revisions，mutate 的剧本由 handler 决定。 */
function makeRemote(revisions, handler) {
    const calls = [];
    let describeCount = 0;
    const remote = {
        settings: {
            describe: async () => {
                const revision = revisions[Math.min(describeCount, revisions.length - 1)];
                describeCount += 1;
                return { ok: true, value: { namespaces: [{ ns: 'dsh-office-mode', revision }] } };
            },
            mutate: async (ns, ops, revision) => {
                calls.push({ ns, ops, revision });
                return handler(calls.length, revision);
            },
        },
    };
    return { remote, calls, describeCount: () => describeCount };
}

await check('写入：先读 revision，settings/conflict 后重读 revision 有限重试', async () => {
    const fake = makeRemote([3, 4], (n) => (n === 1
        ? { ok: false, error: { code: 'settings/conflict', message: 'revision 3 does not match 4' } }
        : { ok: true }));
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, remote: fake.remote });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(fake.calls.length, 2, '冲突后应恰好重试一次');
    assert.deepEqual(fake.calls[0].ops, [{ op: 'set', path: ['tools', 'office_help'], value: false }]);
    assert.equal(fake.calls[0].ops.length, 1, '只 patch 被改的字段，不整包覆盖');
    assert.ok(!JSON.stringify(fake.calls[0].ops).includes('apiKey'), '写别的字段不得带上已存的 apiKey');
    assert.equal(fake.calls[0].revision, 3, '第一次带读到的 revision');
    assert.equal(fake.calls[1].revision, 4, '冲突后必须重读 revision，而不是沿用旧的');
    const texts = walk(render(officeReg(registrations))).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(!texts.some((t) => t.includes('写入失败')), '写入成功后不应留下错误');
});

await check('写入：settings/rejected 把宿主原话显示出来，且不重试', async () => {
    const fake = makeRemote([7], () => ({ ok: false, error: { code: 'settings/rejected', message: 'Config field "tools.office_help" is not volatile' } }));
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, remote: fake.remote });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(fake.calls.length, 1, '非冲突拒绝不得重试');
    assert.equal(fake.describeCount(), 1, '拒绝后不应再读 revision');
    const texts = walk(render(officeReg(registrations))).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('settings/rejected')), '应带上宿主的 code：' + texts.join(' | '));
    assert.ok(texts.some((t) => t.includes('is not volatile')), '应带上宿主的原始 message');
});

await check('写入：读不到 revision 时退回 configForms（不裸写）', async () => {
    const remote = { settings: {
        describe: async () => ({ ok: false, error: { code: 'gateway/internal', message: 'no settings provider' } }),
        mutate: async () => { throw new Error('不应该走到 mutate'); },
    } };
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, remote });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.ok(scope.writes.some(([path]) => path === 'tools.office_help'), '应回退到 configForms.mutate');
});

// ── 设置写入：直连通道的就绪时序与错误保真（第九十五轮）─────────────────
//
// 现场：页面显示「宿主拒绝了这次写入（configForms 这条通道只回 true/false，
// 服务端原因被丢掉了；remote.settings 也不可用，所以这里拿不到细节）」。
// 根因有两层：
//   1) 宿主把远程命名空间挂成独立服务 remote.<ns>，异步挂载；cordis 的
//      ctx.inject(['remote']) 回调只在 remote 就绪那一刻跑一次，不会再因为
//      remote.settings 事后挂上而重跑。本插件原来只声明 ['remote']，于是
//      注入回调里读到 undefined，被存成「永久 null」—— 直连通道从未生效。
//   2) configForms 的 mutate 只回 true/false，服务端的 error.code / message
//      被它丢掉（dsh-client-ui-settings/src/client/config-form.ts:129-148），
//      所以一旦退回那条通道，页面只能看见一句固定文案。
// 下面三条按这两层分别回归。

await check('写入：remote.settings 晚于 remote 挂上时，服务就绪后下一笔必须走直连', async () => {
    // 现场时序：apply 时 remote 已在、settingsController 还没挂上（它排在
    // dsh-api-remotes $mount 列表第 5 位，见该包 src/client/index.ts:199-209）。
    const remote = {};
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations, injects, mountRemoteNamespace } = loadBundle({ forms: scope.forms, remote });
    mod.apply(ctx);
    assert.ok(injects.some((item) => item.services.includes('remote.settings')),
        '必须把 remote.settings 声明进注入列表：cordis 不会因为事后挂上嵌套服务而重跑旧回调');
    // 挂上之前：只能退回 configForms。
    let tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.ok(scope.writes.some(([path]) => path === 'tools.office_help'), '服务未就绪时应退回 configForms');
    // 命名空间服务这时才挂上（宿主异步 $mount），remote 对象的引用不变。
    const fake = makeRemote([11], () => ({ ok: true }));
    // 不能直接写 remote.settings：严格 Cordis 下服务只能由声明了
    // remote.settings 的注入回调拿到，直接改对象属性会绕过注入、漏掉真实回归。
    mountRemoteNamespace('settings', fake.remote.settings);
    tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: true }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(fake.calls.length, 1, '服务挂上后下一笔必须走直连（旧代码这里恒为 0）');
    assert.deepEqual(fake.calls[0].ops, [{ op: 'set', path: ['tools', 'office_help'], value: true }]);
    assert.equal(fake.calls[0].revision, 11, '直连通道必须带 describe 读到的 revision');
    assert.equal(scope.writes.filter(([path]) => path === 'tools.office_help').length, 1,
        '旧通道只该被用一次（服务挂上之前那笔）');
});

await check('写入：remote.settings 卸载后清掉旧引用，下一笔退回 configForms（不拿死服务）', async () => {
    const fake = makeRemote([2], () => ({ ok: true }));
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations, unmountRemoteNamespace } = loadBundle({ forms: scope.forms, remote: fake.remote });
    mod.apply(ctx);
    let tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(fake.calls.length, 1, '先走直连');
    const before = scope.writes.length;
    unmountRemoteNamespace('settings');
    tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: true }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(fake.calls.length, 1, '卸载后不能再拿旧 namespace 去 describe/mutate');
    assert.equal(scope.writes.length, before + 1, '应退回 configForms');
});

await check('写入：describe 失败时如实说明原因，不谎称「remote.settings 也不可用」', async () => {
    const remote = { settings: {
        describe: async () => ({ ok: false, error: { code: 'gateway/offline', message: 'lost connection to host' } }),
        mutate: async () => { throw new Error('不应该走到 mutate'); },
    } };
    // configForms 也必须拒：只有它也返回 false，页面才会显示那句错误（这条测的是文案内容）。
    const scope = makeScope(SAMPLE_VALUE, { refuse: true });
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, remote });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.ok(scope.writes.some(([path]) => path === 'tools.office_help'), '读不到 revision 时仍应退回 configForms');
    const texts = walk(render(officeReg(registrations))).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('gateway/offline')), '应带出 describe 的失败原因：' + texts.join(' | '));
    assert.ok(!texts.some((t) => t.includes('也不可用')), '服务明明挂着，不得再说「remote.settings 也不可用」：' + texts.join(' | '));
});

await check('写入：没有 remote 服务时，文案只如实说「直连通道还没就绪」', async () => {
    const scope = makeScope(SAMPLE_VALUE, { refuse: true });
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    controlOf(findByLabel(tree, 'office_help')).props.onChange(eventOf({ checked: false }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.ok(scope.writes.some(([path]) => path === 'tools.office_help'), '没有 remote 时必须退回 configForms');
    const texts = walk(render(officeReg(registrations))).map((n) => n.props && n.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(texts.some((t) => t.includes('还没就绪')), '应说明直连通道未就绪：' + texts.join(' | '));
});
// ── 知识库大模型：provider / 模型选择（第九十四轮）─────────────────────────
//
// 问题现场：设置页的「知识库大模型」组只有两个手填输入框，而且「通道」下拉传的
// 是数组形式的 options（Select 只认 {value,label}），三项渲染出来是空白 —— 用户
// 既切不到 host，也不知道宿主里有哪些 provider / 模型，等于整组配不了。
// 这一组用例钉住：通道三选一真的可读；host 通道即使 provider/model 是空配置也会
// 拉 /office-memory/llm-models；模型只列当前 provider 的；切 provider 时
// provider + 清空 model 走**同一笔** mutate；已保存的值始终保留为选项 / 手填兜底，
// 只有成功拿到非空清单时才标注「不在清单里」；清单读取失败与保存失败都原样显示
// 宿主 error（不叠第二层前缀）；乱序 / 切通道的旧响应不落地；openai 仍手填。

/** 带 memory.llm 的设置值（其余组照 SAMPLE_VALUE）。 */
function llmSettings(llm) {
    return Object.assign({}, SAMPLE_VALUE, {
        memory: Object.assign({}, SAMPLE_VALUE.memory, { llm: llm }),
    });
}

/** 设置页取数：llm-models 回给定清单，其它端点（本地模型状态）回空状态。 */
function llmFetch(providers, calls) {
    return (url) => {
        const text = String(url);
        if (calls !== undefined) calls.push(text);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: providers, error: '' });
        }
        return jsonResponse({ ok: true, models: [] });
    };
}

/**
 * 渲染记忆系统设置页（memory.llm 用给定值），等清单请求落地后再渲染一次。
 * fetch 不传时走 loadBundle 的默认替身（返回快照），不会打真实网络。
 */
async function renderLlmSettings(llm, options = {}) {
    const scope = makeScope(llmSettings(llm), options.scope || {});
    const env = { forms: scope.forms };
    if (options.fetch !== undefined) env.fetch = options.fetch;
    // 显式「没有 fetch」：RPC 路径必须能在这种环境里工作（不依赖 HTTP / cwd）。
    if (options.noFetch === true) env.fetch = undefined;
    if (options.remote !== undefined) env.remote = options.remote;
    const { mod, ctx, registrations } = loadBundle(env);
    mod.apply(ctx);
    // 知识库页签才是「知识库大模型」这一组所在的那一页。
    const panel = withTab(settingsReg(registrations), "kb");
    render(panel);
    await new Promise((resolve) => setTimeout(resolve, 20));
    return { tree: render(panel), panel, scope };
}

/** 某个下拉/控件里的 option 取值（按渲染顺序）。 */
function optionValuesOf(control) {
    return walk(control.props.children)
        .filter((node) => node.type === 'option')
        .map((node) => node.props.value);
}

/** 某个下拉/控件里的 option 显示文字。 */
function optionLabelsOf(control) {
    return walk(control.props.children)
        .filter((node) => node.type === 'option')
        // h() 把 children 统一包成数组（h(option, null, "文字") -> children: ["文字"]）。
        .map((node) => [].concat(node.props.children).join(''));
}

const LLM_TWO_PROVIDERS = [
    { provider: 'deepseek-official', models: [{ id: 'deepseek-flash', label: 'DeepSeek-V4.1-Flash' }, { id: 'deepseek-pro', label: 'DeepSeek-Pro' }] },
    { provider: 'gemini', models: [{ id: 'gemini-3-pro', label: 'Gemini 3 Pro' }] },
];

await check('知识库大模型：通道下拉三项都有真实文字，选 host 写回 memory.llm.mode', async () => {
    const { tree, scope } = await renderLlmSettings({});
    const select = controlOf(findByLabel(tree, '通道'));
    assert.ok(select, '应有「通道」这一行');
    assert.equal(select.type, 'select');
    assert.deepEqual(optionValuesOf(select), ['off', 'host', 'openai'], '三个通道的取值');
    assert.deepEqual(optionLabelsOf(select), ['关（默认）', '宿主已配置模型', 'OpenAI 兼容端点'],
        '选项文字不能是 undefined —— 数组 options 曾让三项渲染成空白');
    select.props.onChange(eventOf({ value: 'host' }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.ok(scope.writes.some(([path, value]) => path === 'memory.llm.mode' && value === 'host'),
        '选 host 应写回 memory.llm.mode');
});

await check('知识库大模型：空 provider/model 也拉宿主清单，provider 下拉是真实 id', async () => {
    const calls = [];
    const { tree } = await renderLlmSettings({ mode: 'host', provider: '', model: '' }, {
        fetch: llmFetch(LLM_TWO_PROVIDERS, calls),
    });
    assert.ok(calls.some((url) => url.includes('/office-memory/llm-models')),
        'provider/model 为空时也要拉清单（这正是原来配不了的起点）');
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.ok(providerSelect, '应有 provider 行');
    assert.equal(providerSelect.type, 'select');
    assert.deepEqual(optionValuesOf(providerSelect), ['', 'deepseek-official', 'gemini'],
        'provider 选项来自宿主清单，不是手填');
    assert.deepEqual(optionLabelsOf(providerSelect), ['请选择 provider', 'deepseek-official', 'gemini'],
        'provider 标签只写 id；模型数量不再堆在选项文字里');
    const modelSelect = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelSelect.type, 'select');
    assert.equal(modelSelect.props.disabled, true, '还没选 provider 时模型不可选');
});

await check('知识库大模型：模型只列当前 provider 的，切 provider 用同一笔 mutate 清空 model', async () => {
    const remote = makeRemote([5], () => ({ ok: true }));
    const { tree } = await renderLlmSettings({ mode: 'host', provider: 'deepseek-official', model: 'deepseek-flash' }, {
        fetch: llmFetch(LLM_TWO_PROVIDERS),
        remote: remote.remote,
    });
    const modelSelect = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelSelect.type, 'select');
    assert.equal(modelSelect.props.value, 'deepseek-flash', '已保存的模型是选中的');
    assert.deepEqual(optionValuesOf(modelSelect), ['', 'deepseek-flash', 'deepseek-pro'],
        '模型下拉只含当前 provider（gemini 的模型不该出现）');
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    providerSelect.props.onChange(eventOf({ value: 'gemini' }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(remote.calls.length, 1, 'provider 与 model 必须在同一笔 mutate');
    assert.deepEqual(remote.calls[0].ops, [
        { op: 'set', path: ['memory', 'llm', 'provider'], value: 'gemini' },
        { op: 'set', path: ['memory', 'llm', 'model'], value: '' },
    ], '切 provider 同时把旧模型清空，避免错配');
    assert.equal(remote.calls[0].revision, 5, '仍然带读到的 revision');
});

await check('知识库大模型：切完 provider 还没选模型时，能接着选新 provider 的模型，旧模型不会回填', async () => {
    // 切 provider 后的中间态：provider=gemini、model 已被同一笔 mutate 清空。
    // 这个状态必须「可继续配置」：模型下拉可用、值为空、只列新 provider 的模型；
    // 选中之后一次写入只落 memory.llm.model。
    const remote = makeRemote([7], () => ({ ok: true }));
    const { tree } = await renderLlmSettings({ mode: 'host', provider: 'gemini', model: '' }, {
        fetch: llmFetch(LLM_TWO_PROVIDERS),
        remote: remote.remote,
    });
    const modelSelect = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelSelect.type, 'select', 'provider 已选，模型必须是可选的下拉而不是手填死路');
    assert.equal(modelSelect.props.disabled, false, 'provider 已选，模型下拉不能还是灰的');
    assert.equal(modelSelect.props.value, '', '刚清空的模型不能显示成旧值');
    assert.deepEqual(optionValuesOf(modelSelect), ['', 'gemini-3-pro'], '只列新 provider 的模型');
    assert.ok(!optionValuesOf(modelSelect).includes('deepseek-flash'), '旧 provider 的模型不能还留在下拉里');
    modelSelect.props.onChange(eventOf({ value: 'gemini-3-pro' }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(remote.calls.length, 1);
    assert.deepEqual(remote.calls[0].ops, [
        { op: 'set', path: ['memory', 'llm', 'model'], value: 'gemini-3-pro' },
    ], '选模型只写 model 一个字段');
    assert.equal(remote.calls[0].revision, 7);
});

await check('知识库大模型：provider 不在清单里时保留原值；缺 provider 不误报模型缺失', async () => {
    const { tree } = await renderLlmSettings({ mode: 'host', provider: 'my-proxy', model: 'my-model' }, {
        fetch: llmFetch([LLM_TWO_PROVIDERS[0]]),
    });
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.equal(providerSelect.props.value, 'my-proxy', '不静默改掉用户已保存的 provider');
    assert.ok(optionValuesOf(providerSelect).includes('my-proxy'), '未知 provider 作为额外选项保留');
    assert.deepEqual(optionLabelsOf(providerSelect).filter((text) => text === 'my-proxy'), ['my-proxy'],
        '保留的选项只写原值，不在标签里再断言「不在清单里」');
    const modelControl = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelControl.type, 'input', '列不出模型的 provider 退回手填');
    assert.equal(String(modelControl.props.defaultValue), 'my-model');
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('不在宿主当前清单里'), 'provider 缺失要标注：' + texts);
    assert.ok(!texts.includes('不在当前 provider 的清单里'),
        'provider 未知时没有可核对的模型清单，不能报模型缺失：' + texts);
});

await check('知识库大模型：已知 provider 但已保存模型不在清单里时，保留选项并标注', async () => {
    const { tree } = await renderLlmSettings({ mode: 'host', provider: 'deepseek-official', model: 'deepseek-old' }, {
        fetch: llmFetch(LLM_TWO_PROVIDERS),
    });
    const modelSelect = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelSelect.type, 'select');
    assert.equal(modelSelect.props.value, 'deepseek-old');
    assert.ok(optionValuesOf(modelSelect).includes('deepseek-old'), '未知模型作为额外选项保留');
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('不在当前 provider 的清单里'), '模型缺失要标注：' + texts);
});

await check('知识库大模型：读取失败显示宿主原文一次，不叠前缀、不误报缺失', async () => {
    const calls = [];
    const failing = (url) => {
        const text = String(url);
        calls.push(text);
        if (text.includes('/office-memory/llm-models')) {
            // 真实 src/kb-llm.js 的 error 已经带这层前缀（见 src/kb-llm.js:857）。
            return Promise.resolve({
                ok: false, status: 502,
                json: () => Promise.resolve({
                    ok: false,
                    error: '读取宿主模型清单失败：宿主模型注册表暂时不可用：provider registry offline',
                }),
            });
        }
        return jsonResponse({ ok: true, models: [] });
    };
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'my-proxy', model: 'my-model' },
        { fetch: failing },
    );
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('provider registry offline'), '要显示宿主原始 error：' + texts);
    assert.equal(texts.split('读取宿主模型清单失败').length - 1, 1,
        '宿主原文里的前缀只能出现一次，界面不能再叠一层：' + texts);
    assert.ok(!texts.includes('不在宿主当前清单里') && !texts.includes('不在当前 provider 的清单里'),
        '读取失败时没有可核对清单，不能误报缺失：' + texts);
    const retry = byProp(tree, 'data-llm-catalog-retry', 'true')[0];
    assert.ok(retry, '失败时应给「重试」');
    const before = calls.filter((url) => url.includes('/office-memory/llm-models')).length;
    retry.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(calls.filter((url) => url.includes('/office-memory/llm-models')).length, before + 1,
        '重试只重发模型清单这一条请求');
});

await check('知识库大模型：清单为空给可读提示，且不误报已保存值缺失', async () => {
    const { tree } = await renderLlmSettings({ mode: 'host', provider: 'my-proxy', model: 'my-model' }, {
        fetch: llmFetch([]),
    });
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('宿主清单里没有可用的 provider'), '空清单要有下一步：' + texts);
    assert.ok(!texts.includes('不在宿主当前清单里') && !texts.includes('不在当前 provider 的清单里'),
        '空表没有可核对项，不能报缺失：' + texts);
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.deepEqual(optionValuesOf(providerSelect), ['', 'my-proxy'],
        '已保存的 provider 仍保留为选项');
    assert.equal(providerSelect.props.value, 'my-proxy');
});

await check('知识库大模型：openai 通道仍手填模型，且不打宿主清单请求', async () => {
    const calls = [];
    const { tree } = await renderLlmSettings({ mode: 'openai', provider: '', model: 'gpt-4o-mini' }, {
        fetch: llmFetch(LLM_TWO_PROVIDERS, calls),
    });
    assert.equal(calls.some((url) => url.includes('/office-memory/llm-models')), false,
        'openai 通道不枚举远端模型，不该请求宿主清单');
    const providerControl = controlOf(findByLabel(tree, 'provider'));
    assert.equal(providerControl.type, 'input', 'openai 的 provider 仍手填（一般留空）');
    const modelControl = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelControl.type, 'input');
    assert.equal(String(modelControl.props.defaultValue), 'gpt-4o-mini');
});

await check('知识库大模型：保存失败把宿主 code / message 原样显示，且非冲突不重试', async () => {
    const remote = makeRemote([9], () => ({
        ok: false,
        error: { code: 'settings/rejected', message: 'memory.llm.model is not volatile' },
    }));
    const { tree, panel } = await renderLlmSettings({ mode: 'host', provider: 'deepseek-official', model: '' }, {
        fetch: llmFetch(LLM_TWO_PROVIDERS),
        remote: remote.remote,
    });
    controlOf(findByLabel(tree, '模型')).props.onChange(eventOf({ value: 'deepseek-pro' }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(remote.calls.length, 1, '非冲突拒绝不重试');
    const texts = textsOf(render(panel)).join(' | ');
    assert.ok(texts.includes('settings/rejected'), '应显示宿主 code：' + texts);
    assert.ok(texts.includes('memory.llm.model is not volatile'), '应显示宿主原始 message：' + texts);
});

await check('知识库大模型：清单请求挂起时不算「不在清单里」，已保存的值仍可见', async () => {
    const pending = new Promise(() => {});
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'my-proxy', model: 'my-model' },
        { fetch: (url) => (String(url).includes('/office-memory/llm-models') ? pending : jsonResponse({ ok: true, models: [] })) },
    );
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('正在读取宿主模型清单'), '首帧/挂起应显示加载中：' + texts);
    assert.ok(!texts.includes('不在清单里'), '还没拿到清单，不能报缺失：' + texts);
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.equal(providerSelect.type, 'select');
    assert.equal(providerSelect.props.value, 'my-proxy', '已保存的 provider 仍是选中值');
    assert.ok(optionValuesOf(providerSelect).includes('my-proxy'), '且它必须在 options 里');
    const modelControl = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelControl.type, 'input', '清单没回来时模型走手填兜底');
    assert.equal(String(modelControl.props.defaultValue), 'my-model');
});

await check('知识库大模型：客户端自己生成的读取错误清楚，且不借用宿主前缀', async () => {
    const fetchImpl = (url) => {
        if (String(url).includes('/office-memory/llm-models')) {
            // 200/502 但正文不是 JSON：这是客户端自己识别的错误，不是宿主 error。
            return Promise.resolve({ ok: false, status: 502, json: () => Promise.reject(new Error('not json')) });
        }
        return jsonResponse({ ok: true, models: [] });
    };
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'my-proxy', model: 'my-model' },
        { fetch: fetchImpl },
    );
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('宿主模型清单响应格式不对（HTTP 502）'), '客户端错误要自解释：' + texts);
    assert.ok(!texts.includes('读取宿主模型清单失败'), '客户端错误不能借用宿主的前缀：' + texts);
    assert.ok(!texts.includes('不在清单里'), '失败时不能误报缺失：' + texts);
});

await check('知识库大模型：已知 provider 一个模型都没列时手填，不叠「缺失」提示', async () => {
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'deepseek-official', model: 'deepseek-x' },
        { fetch: llmFetch([{ provider: 'deepseek-official', models: [] }]) },
    );
    const modelControl = controlOf(findByLabel(tree, '模型'));
    assert.equal(modelControl.type, 'input');
    assert.equal(String(modelControl.props.defaultValue), 'deepseek-x');
    const texts = textsOf(tree).join(' | ');
    assert.ok(!texts.includes('不在宿主当前清单里'), 'provider 在清单里，不能报 provider 缺失：' + texts);
    assert.ok(!texts.includes('不在当前 provider 的清单里'), 'models=[] 没有可核对项，不能报模型缺失：' + texts);
});

await check('知识库大模型：连点重试造成乱序响应时，只有最新一次结算', async () => {
    const deferred = [];
    let catalogCalls = 0;
    const fetchImpl = (url) => {
        if (!String(url).includes('/office-memory/llm-models')) return jsonResponse({ ok: true, models: [] });
        catalogCalls += 1;
        if (catalogCalls === 1) {
            return Promise.resolve({
                ok: false, status: 502,
                json: () => Promise.resolve({ ok: false, error: '读取宿主模型清单失败：registry offline' }),
            });
        }
        return new Promise((resolve) => { deferred.push(resolve); });
    };
    const { tree, panel } = await renderLlmSettings({ mode: 'host', provider: '', model: '' }, { fetch: fetchImpl });
    const retry = byProp(tree, 'data-llm-catalog-retry', 'true')[0];
    assert.ok(retry, '首次失败后应有重试按钮');
    retry.props.onClick();
    retry.props.onClick();
    assert.equal(deferred.length, 2, '两次重试各发一条清单请求');
    const okBody = (providers) => ({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, providers, error: '' }) });
    // 后发的先回，先发的后回：旧响应必须被请求号挡掉。
    deferred[1](okBody([LLM_TWO_PROVIDERS[1]]));
    await new Promise((resolve) => setTimeout(resolve, 20));
    deferred[0]({
        ok: false, status: 502,
        json: () => Promise.resolve({ ok: false, error: '读取宿主模型清单失败：旧的失败' }),
    });
    await new Promise((resolve) => setTimeout(resolve, 20));
    const after = render(panel);
    const texts = textsOf(after).join(' | ');
    assert.ok(!texts.includes('旧的失败'), '旧响应不能盖掉新清单或贴回旧错误：' + texts);
    assert.ok(texts.includes('gemini'), '最新一次的清单要生效：' + texts);
    assert.deepEqual(optionValuesOf(controlOf(findByLabel(after, 'provider'))), ['', 'gemini'],
        '下拉只反映最新一次响应');
    assert.equal(byProp(after, 'data-llm-catalog-error', 'true').length, 0, '最新一次成功就不该留错误');
});

await check('知识库大模型：host→openai→host 后，旧通道的清单响应不落地', async () => {
    const scope = makeScope(llmSettings({ mode: 'host', provider: '', model: '' }));
    const deferred = [];
    const fetchImpl = (url) => {
        if (!String(url).includes('/office-memory/llm-models')) return jsonResponse({ ok: true, models: [] });
        return new Promise((resolve) => { deferred.push(resolve); });
    };
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, fetch: fetchImpl });
    mod.apply(ctx);
    const panel = withTab(settingsReg(registrations), "kb");
    render(panel);
    assert.equal(deferred.length, 1, 'host 首帧拉一次清单');
    scope.snapshot.value.memory.llm.mode = 'openai';
    render(panel);
    assert.equal(deferred.length, 1, 'openai 通道不拉清单');
    scope.snapshot.value.memory.llm.mode = 'host';
    render(panel);
    assert.equal(deferred.length, 2, '切回 host 重新拉清单');
    const okBody = (providers) => ({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, providers, error: '' }) });
    // 后一次先回，前一次（旧通道）后回。
    deferred[1](okBody([LLM_TWO_PROVIDERS[1]]));
    await new Promise((resolve) => setTimeout(resolve, 20));
    deferred[0](okBody([LLM_TWO_PROVIDERS[0]]));
    await new Promise((resolve) => setTimeout(resolve, 20));
    const after = render(panel);
    assert.deepEqual(optionValuesOf(controlOf(findByLabel(after, 'provider'))), ['', 'gemini'],
        '只认最后一次 host 请求的清单');
});

await check('知识库大模型：切回 host 后新请求未决时，旧清单不能报「不在清单里」', async () => {
    const scope = makeScope(llmSettings({ mode: 'host', provider: 'my-proxy', model: 'my-model' }));
    const deferred = [];
    let calls = 0;
    const fetchImpl = (url) => {
        if (!String(url).includes('/office-memory/llm-models')) return jsonResponse({ ok: true, models: [] });
        calls += 1;
        if (calls === 1) {
            // 第一次成功，但旧表里没有已保存的 provider。
            return Promise.resolve({
                ok: true, status: 200,
                json: () => Promise.resolve({ ok: true, providers: [LLM_TWO_PROVIDERS[1]], error: '' }),
            });
        }
        return new Promise((resolve) => { deferred.push(resolve); });
    };
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, fetch: fetchImpl });
    mod.apply(ctx);
    const panel = withTab(settingsReg(registrations), "kb");
    render(panel);
    await new Promise((resolve) => setTimeout(resolve, 20));
    let tree = render(panel);
    assert.ok(textsOf(tree).join(' | ').includes('不在宿主当前清单里'),
        '旧表成功但不含已存 provider，这一刻的标注是对的');
    scope.snapshot.value.memory.llm.mode = 'openai';
    render(panel);
    scope.snapshot.value.memory.llm.mode = 'host';
    tree = render(panel);
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('正在读取宿主模型清单'), '切回 host 应显示新请求的加载态：' + texts);
    assert.ok(!texts.includes('不在宿主当前清单里') && !texts.includes('不在当前 provider 的清单里'),
        '新请求未决时不能拿旧清单报缺失：' + texts);
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.equal(providerSelect.props.value, 'my-proxy', '已存 provider 仍保留为选中值');
    assert.ok(optionValuesOf(providerSelect).includes('my-proxy'), '且仍在 options 里');
    assert.equal(deferred.length, 1, '切回 host 确实发了新请求');
    deferred[0]({
        ok: true, status: 200,
        json: () => Promise.resolve({ ok: true, providers: [LLM_TWO_PROVIDERS[0]], error: '' }),
    });
    await new Promise((resolve) => setTimeout(resolve, 20));
});

// ── 知识库大模型：宿主 session.modelCatalog RPC（登录后首选）──────────────
//
// 旧路径是插件自己的 HTTP 只读接口，需要 cwd；新路径走宿主 remote RPC：
// { ok:true, value: ModelCatalog } 或 { ok:false, error }。ModelCatalog 的真实
// 形状见 dsh-api-session-controller/src/types.ts:110-141。这里钉住：RPC 优先且
// 不碰 fetch；权限拒绝不偷偷改走 HTTP；partial failures 照实显示；remote
// 迟绑定时自动刷新；groups/models 的 name 用作显示名。

const LLM_RPC_GROUPS = [
    {
        id: 'deepseek-official', name: 'DeepSeek 官方',
        models: [
            { id: 'deepseek-flash', name: 'DeepSeek-V4.1-Flash' },
            { id: 'deepseek-pro', name: 'DeepSeek-Pro' },
        ],
    },
    { id: 'gemini', name: 'Google Gemini', models: [{ id: 'gemini-3-pro', name: 'Gemini 3 Pro' }] },
];

await check('知识库大模型：登录后优先走 session.modelCatalog，无 fetch / 无 cwd 也能列出真实模型', async () => {
    let rpcCalls = 0;
    const remote = { session: {
        modelCatalog: async () => {
            rpcCalls += 1;
            return { ok: true, value: { groups: LLM_RPC_GROUPS, failures: [] } };
        },
    } };
    // noFetch：环境里根本没有 fetch。RPC 路径不该依赖它，也不该退化成报错。
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'deepseek-official', model: 'deepseek-flash' },
        { remote, noFetch: true },
    );
    assert.equal(rpcCalls, 1, '应只调用一次 modelCatalog');
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.deepEqual(optionValuesOf(providerSelect), ['', 'deepseek-official', 'gemini'],
        'provider 用 groups[].id');
    const modelSelect = controlOf(findByLabel(tree, '模型'));
    assert.deepEqual(optionValuesOf(modelSelect), ['', 'deepseek-flash', 'deepseek-pro'],
        '模型用 groups[].models[].id，且只列当前 provider');
    assert.deepEqual(optionLabelsOf(modelSelect), ['请选择模型', 'DeepSeek-V4.1-Flash', 'DeepSeek-Pro'],
        '模型显示名用 models[].name，不是 id');
    assert.equal(modelSelect.props.value, 'deepseek-flash', '已保存模型仍是选中值');
    const texts = textsOf(tree).join(' | ');
    assert.ok(!texts.includes('当前环境没有 fetch'), 'RPC 成功时不该走 HTTP 兜底：' + texts);
    assert.ok(!texts.includes('读取宿主模型清单失败'), '成功时不应留错误：' + texts);
});

await check('知识库大模型：RPC 权限拒绝原样显示，且绝不自动改走 HTTP', async () => {
    let rpcCalls = 0;
    const calls = [];
    const remote = { session: { modelCatalog: async () => {
        rpcCalls += 1;
        return { ok: false, error: { code: 'permission/denied', message: '需要登录后才能读取模型清单' } };
    } } };
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'my-proxy', model: 'my-model' },
        { remote, fetch: llmFetch(LLM_TWO_PROVIDERS, calls) },
    );
    const catalogCalls = () => calls.filter((url) => url.includes('/office-memory/llm-models')).length;
    assert.equal(catalogCalls(), 0, '权限拒绝不得退回模型清单 HTTP 接口（那会绕过授权）');
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('permission/denied'), '要显示宿主 error.code：' + texts);
    assert.ok(texts.includes('需要登录后才能读取模型清单'), '要显示宿主原始 message：' + texts);
    assert.equal(texts.split('读取宿主模型清单失败').length - 1, 1, '同一个失败只显示一次：' + texts);
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.equal(providerSelect.props.value, 'my-proxy', '读取失败也要保留已保存的 provider');
    assert.ok(optionValuesOf(providerSelect).includes('my-proxy'), '且它仍在 options 里');
    const retry = byProp(tree, 'data-llm-catalog-retry', 'true')[0];
    assert.ok(retry, 'RPC 拒绝后要给重试');
    retry.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(rpcCalls, 2, '重试重发 RPC');
    assert.equal(catalogCalls(), 0, '重试也不得改走模型清单 HTTP 接口');
});

await check('知识库大模型：旧宿主返回 gateway/invocation-unavailable 时才退回只读 HTTP 清单', async () => {
    const calls = [];
    let rpcCalls = 0;
    const remote = { session: { modelCatalog: async () => {
        rpcCalls += 1;
        return {
            ok: false,
            error: {
                code: 'gateway/invocation-unavailable',
                message: 'no active Remote method exports this endpoint',
            },
        };
    } } };
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'deepseek-official', model: 'deepseek-flash' },
        { remote, fetch: llmFetch(LLM_TWO_PROVIDERS, calls) },
    );
    assert.equal(rpcCalls, 1, '先尝试 RPC');
    assert.equal(calls.filter((url) => url.includes('/office-memory/llm-models')).length, 1,
        '只有明确的「宿主没导出这个方法」才退回 HTTP');
    assert.deepEqual(optionValuesOf(controlOf(findByLabel(tree, 'provider'))),
        ['', 'deepseek-official', 'gemini'], '界面用 HTTP 清单');
    const texts = textsOf(tree).join(' | ');
    assert.ok(!texts.includes('读取宿主模型清单失败'), '兼容回退成功时不该显示错误：' + texts);
});

await check('知识库大模型：RPC 抛错不吞、不退回 HTTP，错误原文可见', async () => {
    const calls = [];
    const remote = { session: { modelCatalog: async () => { throw new Error('gateway unavailable'); } } };
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: '', model: '' },
        { remote, fetch: llmFetch(LLM_TWO_PROVIDERS, calls) },
    );
    assert.equal(calls.filter((url) => url.includes('/office-memory/llm-models')).length, 0,
        'RPC 抛错时不得偷偷退回模型清单 HTTP 接口');
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('gateway unavailable'), '要保留真实错误：' + texts);
    assert.ok(texts.includes('读取宿主模型清单失败'), '要说明这是读模型清单失败：' + texts);
});

await check('知识库大模型：RPC 的 partial failures 照实列出，不因部分成功就吞掉', async () => {
    const remote = { session: { modelCatalog: async () => ({
        ok: true,
        value: {
            groups: [LLM_RPC_GROUPS[0]],
            failures: [{ id: 'router', name: 'Router', message: 'provider registry offline' }],
        },
    }) } };
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'deepseek-official', model: 'deepseek-flash' },
        { remote, noFetch: true },
    );
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('Router'), '要点名失败的 provider：' + texts);
    assert.ok(texts.includes('provider registry offline'), '要显示失败原因：' + texts);
    assert.ok(texts.includes('部分 provider 的模型清单没读到'), '要明确这是部分失败：' + texts);
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.deepEqual(optionValuesOf(providerSelect), ['', 'deepseek-official'],
        '成功的 groups 仍然可用');
});

await check('知识库大模型：partial failures 里已保存的 provider 不误报「不在清单里」', async () => {
    const remote = { session: { modelCatalog: async () => ({
        ok: true,
        value: {
            groups: [LLM_RPC_GROUPS[0]],
            failures: [{ id: 'router', name: 'Router', message: 'provider registry offline' }],
        },
    }) } };
    const { tree } = await renderLlmSettings(
        { mode: 'host', provider: 'router', model: 'router-model' },
        { remote, noFetch: true },
    );
    const texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('Router'), '失败原因仍要列出：' + texts);
    assert.ok(!texts.includes('不在宿主当前清单里'),
        '读取失败的 provider 只是没读到，不能说它不在清单里：' + texts);
    assert.ok(!texts.includes('不在当前 provider 的清单里'),
        'provider 清单没读到时也不能判模型缺失：' + texts);
    const providerSelect = controlOf(findByLabel(tree, 'provider'));
    assert.equal(providerSelect.props.value, 'router', '已保存的 provider 保留为选中值');
    assert.ok(optionValuesOf(providerSelect).includes('router'), '且仍在 options 里');
});

await check('知识库大模型：remote 迟绑定后自动改走 RPC，不再重复打旧 HTTP 接口', async () => {
    const calls = [];
    const scope = makeScope(llmSettings({ mode: 'host', provider: '', model: '' }));
    // 第一次注入只有空 remote（session 方法还没有）——选择器按旧宿主走 HTTP。
    const remote = {};
    const { mod, ctx, registrations, injects, mountRemoteNamespace } = loadBundle({
        forms: scope.forms,
        remote,
        fetch: llmFetch([LLM_TWO_PROVIDERS[1]], calls),
    });
    mod.apply(ctx);
    const panel = withTab(settingsReg(registrations), "kb");
    render(panel);
    await new Promise((resolve) => setTimeout(resolve, 20));
    let tree = render(panel);
    assert.equal(calls.filter((url) => url.includes('/office-memory/llm-models')).length, 1,
        '方法还没出现时走旧 HTTP 路径');
    assert.deepEqual(optionValuesOf(controlOf(findByLabel(tree, 'provider'))), ['', 'gemini']);

    // 模拟 remote 服务这时才就绪：触发真实的 inject 回调（换一个对象，表明确实
    // 发生了重新注入），选择器应自己重新取数并改走 RPC。
    let rpcCalls = 0;
    const nextRemote = {
        session: { modelCatalog: async () => {
            rpcCalls += 1;
            return { ok: true, value: { groups: [LLM_RPC_GROUPS[0]], failures: [] } };
        } },
    };
    // 不能直接触发 ['remote'] 的那段回调：session 是独立服务，按 cordis 语义要由
    // `remote.session` 的注入回调重跑；这里用替身复刻那一步。
    assert.ok(injects.some((item) => item.services.includes('remote.session')),
        '必须把 remote.session 声明进注入列表');
    mountRemoteNamespace('session', nextRemote.session);
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(rpcCalls, 1, '迟绑定后应自动走 RPC');
    assert.equal(calls.filter((url) => url.includes('/office-memory/llm-models')).length, 1,
        '改走 RPC 后不应再打旧 HTTP 接口');
    assert.deepEqual(optionValuesOf(controlOf(findByLabel(tree, 'provider'))), ['', 'deepseek-official'],
        '界面换成 RPC 返回的清单');
});

await check('知识库大模型：remote.session 卸载后清掉旧 RPC，下一笔改回只读 HTTP', async () => {
    const calls = [];
    let rpcCalls = 0;
    const remote = { session: {
        modelCatalog: async () => {
            rpcCalls += 1;
            return { ok: true, value: { groups: [LLM_RPC_GROUPS[0]], failures: [] } };
        },
    } };
    const scope = makeScope(llmSettings({ mode: 'host', provider: '', model: '' }));
    const { mod, ctx, registrations, unmountRemoteNamespace } = loadBundle({
        forms: scope.forms,
        remote,
        fetch: llmFetch([LLM_TWO_PROVIDERS[1]], calls),
    });
    mod.apply(ctx);
    const panel = withTab(settingsReg(registrations), "kb");
    render(panel);
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(rpcCalls, 1, '先走 RPC');
    assert.equal(calls.filter((url) => url.includes('/office-memory/llm-models')).length, 0, 'RPC 成功不碰 HTTP');
    unmountRemoteNamespace('session');
    render(panel);
    await new Promise((resolve) => setTimeout(resolve, 20));
    assert.equal(rpcCalls, 1, '卸载后不能再调旧 session');
    assert.equal(calls.filter((url) => url.includes('/office-memory/llm-models')).length, 1, '应退回只读 HTTP');
    const texts = textsOf(render(panel)).join(' | ');
    assert.ok(!texts.includes('without inject'), '不得把 cordis 注入错误当清单错误显示：' + texts);
});

// ── 站点清单（检索编排组里的可视化编辑器）─────────────────────────────────
//
// 数据契约（src/settings.js 的 search.sites 段，见 src/site-catalog.js）：
//   enabled / fallback 两个布尔 + maxPerCall（1..8）+ entries 一条数组。
// 界面把 entries 当**整条数组**写回（没有「只改第 n 条」的协议），所以下面每一条
// 交互断言都要同时核对「路径是 search.sites.entries」与「写回的那条数组长什么样」。

/**
 * 从 bundle 源码里抠出一个顶层常量数组字面量并求值。
 *
 * 浏览器产物是 CJS 单文件、这些常量不导出，而「逐条比对」不能靠
 * `source.includes(...)`（那是子串核对，改了字段名照样过）。数组里只有字符串与
 * 布尔，跳掉字符串后按方括号配平，再用 new Function 求值即可（行注释不影响）。
 */
function bundleArrayLiteral(source, name) {
    const marker = 'const ' + name + ' = [';
    const at = source.indexOf(marker);
    assert.ok(at >= 0, `bundle 里应有 ${name} 常量`);
    const start = at + marker.length - 1;
    let depth = 0;
    let quote = null;
    for (let index = start; index < source.length; index += 1) {
        const ch = source[index];
        if (quote !== null) {
            if (ch === '\\') { index += 1; continue; }
            if (ch === quote) quote = null;
            continue;
        }
        if (ch === '"' || ch === "'") { quote = ch; continue; }
        if (ch === '[') depth += 1;
        else if (ch === ']') {
            depth -= 1;
            if (depth === 0) return new Function('return ' + source.slice(start, index + 1))();
        }
    }
    throw new Error(`${name} 的数组字面量没有闭合`);
}

/** 用带站点清单的夹具渲染办公模式那一页（patch 为 null = 整个 search.sites 缺失）。 */
function renderSites(patch, options = {}) {
    const value = JSON.parse(JSON.stringify(SAMPLE_VALUE));
    value.search.sites = patch === null
        ? {}
        : Object.assign({}, value.search.sites, patch);
    const scope = makeScope(value, options);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const panel = withTab(settingsReg(registrations), 'search');
    return { tree: render(panel), panel, scope };
}

/** 某一行的启用开关（数据属性挂在 Toggle 外面那层，开关本身是里面的 input）。 */
function siteToggleOf(tree, domain) {
    const cell = byProp(tree, 'data-site-toggle', domain)[0];
    assert.ok(cell, `应有 ${domain} 这一行的启用开关`);
    const box = walk(cell).find((node) => node.type === 'input' && node.props.type === 'checkbox');
    assert.ok(box, `${domain} 那一行应渲染出复选框`);
    return box;
}

await check('站点清单：镜像目录与服务端 BUILTIN_SITES 逐条相等（含字段顺序）', () => {
    const mirror = bundleArrayLiteral(readFileSync(BUNDLE, 'utf8'), 'SITE_CATALOG');
    assert.equal(mirror.length, BUILTIN_SITES.length, '条数要与服务端一致');
    assert.deepEqual(mirror, BUILTIN_SITES.map((item) => Object.assign({}, item)),
        '字段名与取值都要与服务端一致（改了一边就必须改另一边）');
    assert.deepEqual(mirror.map((item) => Object.keys(item)), BUILTIN_SITES.map((item) => Object.keys(item)),
        '字段顺序也要一致：界面按这个顺序摆开关 / 域名 / 显示名 / 说明');
    assert.deepEqual(mirror.map((item) => item.domain), BUILTIN_SITES.map((item) => item.domain),
        '条目顺序就是默认优先顺序');
});

await check('站点清单：镜像类型表与服务端 SITE_TYPES 逐条相等', () => {
    const mirror = bundleArrayLiteral(readFileSync(BUNDLE, 'utf8'), 'SITE_TYPES');
    assert.deepEqual(mirror, SITE_TYPES.map((item) => Object.assign({}, item)));
    assert.deepEqual(mirror.map((item) => Object.keys(item)), SITE_TYPES.map((item) => Object.keys(item)));
});

await check('站点清单：单次最多的区间与服务端 SITE_PRIORITY_LIMITS 一致', () => {
    const mirror = bundleArrayLiteral(readFileSync(BUNDLE, 'utf8'), 'SITE_MAX_PER_CALL');
    assert.deepEqual(mirror, [...SITE_PRIORITY_LIMITS.maxPerCall]);
});

await check('站点清单：两个开关分别写回 search.sites.enabled / fallback', () => {
    const { tree, scope } = renderSites({});
    const on = controlOf(findByLabel(tree, '站点优先检索'));
    assert.ok(on, '应渲染「站点优先检索」这一行');
    assert.equal(on.props.type, 'checkbox');
    assert.equal(on.props.checked, true, '夹具里是开着的');
    const back = controlOf(findByLabel(tree, '退回泛搜'));
    assert.equal(back.props.type, 'checkbox');
    assert.equal(back.props.checked, true);
    on.props.onChange(eventOf({ checked: false }));
    back.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[0], ['search.sites.enabled', false]);
    assert.deepEqual(scope.writes[1], ['search.sites.fallback', false]);
    // 值缺失（schema 还没接线 / 从没配过）时按「开」显示，与服务端默认一致。
    const bare = renderSites({ enabled: undefined, fallback: undefined });
    assert.equal(controlOf(findByLabel(bare.tree, '站点优先检索')).props.checked, true);
    assert.equal(controlOf(findByLabel(bare.tree, '退回泛搜')).props.checked, true);
});

await check('站点清单：单次最多限定是 1..8 的数字输入，写回 search.sites.maxPerCall', () => {
    const { tree, panel, scope } = renderSites({});
    const input = controlOf(findByLabel(tree, '单次最多限定'));
    assert.ok(input);
    assert.equal(input.props.type, 'number');
    assert.equal(String(input.props.value), '3', '夹具里是 3');
    assert.equal(input.props.min, 1);
    assert.equal(input.props.max, 8);
    input.props.onChange(eventOf({ value: '6' }));
    // 数字控件是「本地草稿 + 失焦提交」：草稿进 state 之后要重渲染一次，
    // 新的 onBlur 才拿得到新草稿（真实 React 也是这个次序）。
    const after = render(panel);
    controlOf(findByLabel(after, '单次最多限定')).props.onBlur();
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['search.sites.maxPerCall', 6]);
});

await check('站点清单：按类型分组渲染，行数等于条目数', () => {
    const { tree } = renderSites({});
    const rows = byProp(tree, 'data-site-row');
    assert.equal(rows.length, 4, '夹具四条就该有四行');
    assert.deepEqual(rows.map((node) => node.props['data-site-row']),
        ['arxiv.org', 'sci-hub.se', 'github.com', 'example.com']);
    assert.deepEqual(byProp(tree, 'data-site-group').map((node) => node.props['data-site-group']),
        ['academic', 'code', 'custom'], '分组顺序跟着 SITE_TYPES（空组不画）');
    const texts = textsOf(tree);
    for (const name of ['学术', '代码', '自定义']) {
        assert.ok(texts.includes(name), `分组标题要用 SITE_TYPES 的 name：${name}`);
    }
    // 编辑器里不许漏出 undefined / NaN（属性、样式与文本一起查）。
    for (const text of stringsOf(byProp(tree, 'data-site-editor')[0])) {
        assert.ok(!text.includes('undefined') && !text.includes('NaN'), `界面上出现了 ${text}`);
    }
    // 勾选态的条数 = 清单里 enabled 的条数（夹具 4 条里 3 条是开的）。
    const boxes = walk(byProp(tree, 'data-site-editor')[0]).filter((node) => node.type === 'input' && node.props.type === 'checkbox');
    assert.equal(boxes.length, 4, '每条一行，行行一个开关');
    assert.equal(boxes.filter((node) => node.props.checked === true).length, 3);
});

await check('站点清单：行内的启用开关写回整条数组（长度不变、只翻这一条）', () => {
    const { tree, scope } = renderSites({});
    const box = siteToggleOf(tree, 'sci-hub.se');
    assert.equal(box.props.checked, false, '默认关闭的影子图书馆应显示未勾选');
    box.props.onChange(eventOf({ checked: true }));
    assert.equal(scope.writes[0][0], 'search.sites.entries');
    const written = scope.writes[0][1];
    assert.equal(written.length, 4);
    assert.equal(written[1].enabled, true);
    assert.deepEqual(written.map((item) => item.enabled), [true, true, true, true]);
    assert.equal(written[0].enabled, true, '别的条目不受影响');
});

await check('站点清单：上移 / 下移交换数组里相邻的两项', () => {
    const { tree, scope } = renderSites({});
    byProp(tree, 'data-site-up', 'sci-hub.se')[0].props.onClick();
    assert.equal(scope.writes[0][0], 'search.sites.entries');
    assert.deepEqual(scope.writes[0][1].map((item) => item.domain),
        ['sci-hub.se', 'arxiv.org', 'github.com', 'example.com']);
    assert.equal(scope.writes[0][1][0].type, 'academic', '只是换位置，字段不带丢');
    assert.equal(scope.writes[0][1][0].enabled, false, '关闭状态跟着走');
    byProp(tree, 'data-site-down', 'sci-hub.se')[0].props.onClick();
    assert.deepEqual(scope.writes[1][1].map((item) => item.domain),
        ['arxiv.org', 'github.com', 'sci-hub.se', 'example.com']);
    // 首行的「上移」与末行的「下移」没有去处，应当禁用。
    assert.equal(byProp(tree, 'data-site-up', 'arxiv.org')[0].props.disabled, true);
    assert.equal(byProp(tree, 'data-site-down', 'example.com')[0].props.disabled, true);
});

await check('站点清单：删除一行后写回 search.sites.entries，长度减一', () => {
    const { tree, scope } = renderSites({});
    byProp(tree, 'data-site-remove', 'github.com')[0].props.onClick();
    assert.equal(scope.writes[0][0], 'search.sites.entries');
    assert.equal(scope.writes[0][1].length, 3);
    assert.deepEqual(scope.writes[0][1].map((item) => item.domain),
        ['arxiv.org', 'sci-hub.se', 'example.com']);
});

await check('站点清单：恢复内置默认写回内置目录（长度等于 BUILTIN_SITES）', () => {
    const { tree, scope } = renderSites({});
    byProp(tree, 'data-site-reset')[0].props.onClick();
    assert.equal(scope.writes[0][0], 'search.sites.entries');
    const written = scope.writes[0][1];
    assert.equal(written.length, BUILTIN_SITES.length);
    assert.deepEqual(written, BUILTIN_SITES.map((item) => Object.assign({}, item)));
});

await check('站点清单：改域名先规范化（去协议 / 路径 / www. / 大写）再写回', () => {
    const { tree, scope } = renderSites({});
    const input = byProp(tree, 'data-site-domain', 'example.com')[0];
    input.props.onBlur(eventOf({ value: '  HTTPS://WWW.Docs.Example.COM/guide?x=1  ' }));
    assert.equal(scope.writes[0][0], 'search.sites.entries');
    assert.equal(scope.writes[0][1][3].domain, 'docs.example.com');
    assert.equal(scope.writes[0][1].length, 4, '长度不变');
    assert.equal(scope.writes[0][1][3].label, '示例站', '只动域名，别的字段原样带回去');
    // 规范化之后等于原值（只是大小写 / 前缀的差别）就不写设置。
    byProp(tree, 'data-site-domain', 'example.com')[0].props.onBlur(eventOf({ value: 'www.EXAMPLE.com/' }));
    assert.equal(scope.writes.length, 1, '没变就不该写');
});

await check('站点清单：非法域名不写入，并给一行错误提示', () => {
    const { tree, panel, scope } = renderSites({});
    byProp(tree, 'data-site-domain', 'github.com')[0].props.onBlur(eventOf({ value: '不是域名' }));
    assert.equal(scope.writes.length, 0, '不合法的域名不该写进设置');
    const after = render(panel);
    assert.equal(byProp(after, 'data-site-error').length, 1, '应给一行错误提示');
    const texts = textsOf(after);
    assert.ok(texts.some((t) => t.includes('不合法')), '提示要说明哪一条不合法');
    assert.ok(texts.some((t) => t.includes('不是域名')), '提示要带上写进来的那个值');
});

await check('站点清单：域名重复不写入，并点名重复的那个域名', () => {
    const { tree, panel, scope } = renderSites({});
    byProp(tree, 'data-site-domain', 'github.com')[0].props.onBlur(eventOf({ value: 'https://www.arXiv.org/list' }));
    assert.equal(scope.writes.length, 0);
    const texts = textsOf(render(panel));
    assert.ok(texts.some((t) => t.includes('已经在清单里')), '要说明是重复');
    assert.ok(texts.some((t) => t.includes('arxiv.org')), '要点名重复的域名（规范化之后的那一个）');
});

await check('站点清单：新增按表单追加到数组末尾（显示名留空回落域名）', () => {
    const { tree, panel, scope } = renderSites({});
    const typeSelect = walk(byProp(tree, 'data-site-add-type')[0]).find((node) => node.type === 'select');
    assert.ok(typeSelect, '新增表单应有类型下拉');
    assert.equal(typeSelect.props.value, 'custom', '默认类型是自定义');
    typeSelect.props.onChange(eventOf({ value: 'book' }));
    let now = render(panel);
    byProp(now, 'data-site-add-domain')[0].props.onChange(eventOf({ value: 'https://www.Gutenberg.org/' }));
    now = render(panel);
    byProp(now, 'data-site-add-label')[0].props.onChange(eventOf({ value: '' }));
    now = render(panel);
    byProp(now, 'data-site-add-note')[0].props.onChange(eventOf({ value: '公版电子书' }));
    now = render(panel);
    byProp(now, 'data-site-add')[0].props.onClick();
    assert.equal(scope.writes[0][0], 'search.sites.entries');
    assert.equal(scope.writes[0][1].length, 5, '追加到末尾');
    assert.deepEqual(scope.writes[0][1][4],
        { type: 'book', domain: 'gutenberg.org', label: 'gutenberg.org', note: '公版电子书', enabled: true });

    // 表单里的非法域名同样不写入，并且不会顺手把已有条目删掉。
    byProp(render(panel), 'data-site-add-domain')[0].props.onChange(eventOf({ value: 'localhost' }));
    const again = render(panel);
    byProp(again, 'data-site-add')[0].props.onClick();
    assert.equal(scope.writes.length, 1, '缺点的域名不该进清单');
    assert.ok(textsOf(render(panel)).some((t) => t.includes('不合法')));
});

await check('站点清单：空清单给空态说明，按钮不重复（工具条已有「恢复内置默认」）', () => {
    const { tree, scope } = renderSites({ entries: [] });
    assert.equal(byProp(tree, 'data-site-row').length, 0, 'entries=[] 就是真的没有站点');
    assert.ok(textsOf(tree).some((t) => t.includes('清单为空')), '要给一句空态说明');
    assert.equal(byProp(tree, 'data-site-load').length, 0, '空态不再重复一个「载入内置目录」');
    byProp(tree, 'data-site-reset')[0].props.onClick();
    assert.equal(scope.writes[0][0], 'search.sites.entries');
    assert.equal(scope.writes[0][1].length, BUILTIN_SITES.length);
});

await check('站点清单：拿不到 entries 时按内置目录显示（不画空清单）', () => {
    const { tree } = renderSites({ entries: undefined });
    assert.equal(byProp(tree, 'data-site-row').length, BUILTIN_SITES.length);
    assert.equal(byProp(tree, 'data-site-empty').length, 0);
    const missing = renderSites(null);
    assert.equal(byProp(missing.tree, 'data-site-row').length, BUILTIN_SITES.length,
        '整个 search.sites 缺失时也按内置目录显示');
});

await check('站点清单：只读连接下所有控件禁用', () => {
    const { tree } = renderSites({}, { writable: false });
    const editor = byProp(tree, 'data-site-editor')[0];
    assert.ok(editor, '应渲染出清单编辑器');
    const controls = walk(editor).filter((node) => node.type === 'input' || node.type === 'select' || node.type === 'button');
    assert.ok(controls.length >= 16, `应渲染出开关 / 输入框 / 按钮，实际 ${controls.length} 个`);
    for (const node of controls) assert.equal(node.props.disabled, true, '只读时控件必须禁用');
    assert.equal(controlOf(findByLabel(tree, '站点优先检索')).props.disabled, true);
    assert.equal(controlOf(findByLabel(tree, '单次最多限定')).props.disabled, true);
});

await check('站点清单：改显示名与说明沿用同一个整条数组写回', () => {
    const { tree, scope } = renderSites({});
    byProp(tree, 'data-site-label', 'example.com')[0].props.onBlur(eventOf({ value: '示例站点' }));
    byProp(tree, 'data-site-note', 'example.com')[0].props.onBlur(eventOf({ value: '  自己加的站点，随便写  ' }));
    assert.deepEqual(scope.writes.map((item) => item[0]), ['search.sites.entries', 'search.sites.entries']);
    assert.equal(scope.writes[0][1][3].label, '示例站点');
    assert.equal(scope.writes[1][1][3].note, '自己加的站点，随便写', '首尾空白要去掉');
    assert.equal(scope.writes[1][1].length, 4);
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
    const tree = renderOfficeTab(registrations, 'search');
    // 按标签找宽控件行，不按「第一个 stacked 行」找 —— 否则以后每加一个
    // 宽控件行都会误伤这条断言。（第七十轮删掉了「子代理可用工具」行，
    // 改用同样很宽的「清单内容」站点编辑器行当代表。）
    const chipRow = walk(tree).find((n) => n.source && n.source.name === 'Row'
        && n.source.props.stacked === true && n.source.props.label === '清单内容');
    assert.ok(chipRow, '清单内容那一行应标记 stacked');
    assert.equal(chipRow.props.style.flexDirection, 'column', 'stacked 行要竖排');
    assert.equal(chipRow.props.style.alignItems, 'stretch');
});

await check('普通行仍是左右排（没有把所有行都改成上下排）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'search');
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
    const tree = renderOfficeTab(registrations, 'search');
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
    // 「工具面里没有 web_search 时，内置检索的每个参数都要能在界面上调到」这件事就落在这组参数上：
    // schema 加了 search.fetch.* / search.proxy 而界面忘了加行 = 用户永远调不到它。
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.search, 'search');
    for (const nested of ['search.engine', 'search.fetch.maxResults', 'search.fetch.fetchPages']) {
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
    for (const nested of ['memory.layers.hot', 'memory.recallQuality.policy', 'memory.quota.recallPerTurn', 'memory.quota.kbSearchPerTurn']) {
        assert.ok(paths.includes(nested), `没有解析出嵌套参数 ${nested}`);
    }
});

await check('界面覆盖服务端 gallery 组里的每个参数（图库视觉描述与图文对齐）', async () => {
    // 第 106 轮之前 schema 里已有 gallery.vision / gallery.imageEmbedding，
    // 但合并设置页没有入口，用户只能看到「视觉描述未配置」的报错。
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.gallery, 'gallery');
    assert.deepEqual(paths.slice().sort(), [
        'gallery.imageEmbedding.apiKey', 'gallery.imageEmbedding.apiKeyEnv',
        'gallery.imageEmbedding.baseURL', 'gallery.imageEmbedding.batchSize',
        'gallery.imageEmbedding.dimensions', 'gallery.imageEmbedding.model',
        'gallery.imageEmbedding.provider', 'gallery.imageEmbedding.timeoutMs',
        'gallery.vision.apiKey', 'gallery.vision.apiKeyEnv', 'gallery.vision.baseURL',
        'gallery.vision.channel', 'gallery.vision.language', 'gallery.vision.maxTokens',
        'gallery.vision.model', 'gallery.vision.provider', 'gallery.vision.timeoutMs',
    ], 'gallery 组就是这十七个参数，多一个少一个都要在这里显形');
    return `${paths.length} 个参数都有对应控件`;
});

await check('图库设置：默认 off 只留一句说明，不铺十行输入框（第一百零六轮）', () => {
    const scope = makeScope(gallerySample(() => {}));
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'gallery');
    assert.equal(byProp(tree, 'data-gallery-vision-off').length, 1, '视觉描述默认关，只给一句未启用');
    assert.equal(byProp(tree, 'data-gallery-vision-on').length, 0);
    assert.equal(byProp(tree, 'data-gallery-image-off').length, 1, '图文对齐默认 none');
    assert.equal(byProp(tree, 'data-gallery-image-on').length, 0);
    assert.equal(findByLabel(tree, '输出上限'), undefined, 'off 时不显示视觉参数');
    assert.ok(textsOf(tree).some((t) => t.includes('生成视觉描述会提示「未配置」')),
        '说明要写出关掉后的直接后果');
});

await check('图库设置：openai 与 vllm 的值回填、失焦写回对应路径（第一百零六轮）', () => {
    const value = gallerySample((gallery) => {
        gallery.vision.channel = 'openai';
        gallery.vision.baseURL = 'https://vision.example.com/v1';
        gallery.vision.model = 'gpt-4o-mini';
        gallery.vision.apiKeyEnv = 'VISION_TEST_KEY';
        gallery.vision.maxTokens = 1200;
        gallery.vision.timeoutMs = 45000;
        gallery.vision.language = 'en';
        gallery.imageEmbedding.provider = 'vllm';
        gallery.imageEmbedding.baseURL = 'http://127.0.0.1:8000/v1';
        gallery.imageEmbedding.model = 'jina-clip-v2';
        gallery.imageEmbedding.apiKeyEnv = 'IMAGE_TEST_KEY';
        gallery.imageEmbedding.dimensions = 1024;
        gallery.imageEmbedding.batchSize = 8;
        gallery.imageEmbedding.timeoutMs = 20000;
    });
    const scope = makeScope(value);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'gallery');

    assert.equal(byProp(tree, 'data-gallery-vision-on', 'openai').length, 1);
    assert.equal(byProp(tree, 'data-gallery-vision-channel')[0].props.value, 'openai');
    assert.equal(byProp(tree, 'data-gallery-vision-base-url')[0].props.defaultValue, 'https://vision.example.com/v1');
    assert.equal(byProp(tree, 'data-gallery-vision-model')[0].props.defaultValue, 'gpt-4o-mini');
    assert.equal(byProp(tree, 'data-gallery-vision-key-env')[0].props.defaultValue, 'VISION_TEST_KEY');

    byProp(tree, 'data-gallery-vision-base-url')[0].props.onBlur(eventOf({ value: 'https://vision2.example.com/v1' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['gallery.vision.baseURL', 'https://vision2.example.com/v1']);
    byProp(tree, 'data-gallery-vision-channel')[0].props.onChange(eventOf({ value: 'host' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['gallery.vision.channel', 'host']);

    assert.equal(byProp(tree, 'data-gallery-image-on', 'vllm').length, 1);
    assert.equal(byProp(tree, 'data-gallery-image-provider')[0].props.value, 'vllm');
    assert.equal(byProp(tree, 'data-gallery-image-base-url')[0].props.defaultValue, 'http://127.0.0.1:8000/v1');
    assert.equal(byProp(tree, 'data-gallery-image-model')[0].props.defaultValue, 'jina-clip-v2');
    byProp(tree, 'data-gallery-image-model')[0].props.onBlur(eventOf({ value: 'jina-clip-v3' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['gallery.imageEmbedding.model', 'jina-clip-v3']);
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

await check('界面覆盖服务端 python 组里的每个参数（计算与绘图）', async () => {
    // 第二十二轮只给 av 组补了这条核对，python 组一直是「服务端有、界面没有」——
    // 四个参数只能改配置文件，而两边（schema 与界面）都不会报错。
    // 第三十八轮把这一组补进界面，同时把它钉在这条断言上。
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.python, 'python');
    assert.deepEqual(paths.slice().sort(), ['python.bin', 'python.enabled', 'python.outDir', 'python.timeoutMs'],
        'python 组就是这四个参数，多一个少一个都要在这里显形');
    return `${paths.length} 个参数都有对应控件`;
});

await check('界面覆盖服务端 ocr 组里的每个参数（文件识别）', async () => {
    // 服务端 ocr 组从加进来那天起界面就没有对应分组：缺 Key 的报错写着
    // 「设置 → 办公模式 → 文件识别（OCR）」，而那一页根本没有能填 Key 的地方。
    // 这条把七个参数逐项钉在界面上（含 role('secret') 的 apiKey 与默认关的复用开关）。
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.ocr, 'ocr');
    assert.deepEqual(paths.slice().sort(),
        ['ocr.apiKey', 'ocr.apiKeyEnv', 'ocr.baseURL', 'ocr.enabled', 'ocr.purpose', 'ocr.reuseOfficialProvider', 'ocr.timeoutMs'],
        'ocr 组就是这七个参数，多一个少一个都要在这里显形');
    return `${paths.length} 个参数都有对应控件`;
});

await check('界面覆盖服务端 guard 组里的每个参数（新-15① 的工具面守卫）', async () => {
    const { createOfficeSettings } = await import('../src/settings.js');
    const z = await loadSchemastery();
    if (z === undefined) {
        console.log('      （本机解析不到 @deepseek-ai/schemastery，跳过这一项）');
        return;
    }
    const schema = await resolvePlain(z, createOfficeSettings(z));
    const source = readFileSync(BUNDLE, 'utf8');
    const paths = assertCovers(source, schema.guard, 'guard');
    assert.deepEqual(paths, ['guard.hideDevToolsInOffice'], 'guard 组就这一个开关');
    return `${paths.length} 个参数都有对应控件`;
});

await check('工具面守卫：开关默认关（读到的值决定勾选状态），说明里点明代价', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = render(officeReg(registrations));
    const row = rowInGroup(tree, 'guard', '藏掉编程类工具');
    assert.ok(row, 'guard 组缺少「藏掉编程类工具」这一行');
    const toggle = controlOf(row);
    // 样例配置没有 guard 段 → 按默认 false 渲染（默认不隐藏是用户定的口径）。
    assert.equal(toggle.props.checked, false, '默认必须是未勾选');
    // 代价要写在说明里，且**不能**写错的版本：曾经写成「检索派工退回内置通道」，
    // 实测是错的（派工用的是宿主 subagents 服务，不是被掩掉的工具）。
    const text = JSON.stringify(row);
    assert.ok(/命令行/.test(text) && /后台任务/.test(text), '说明里要写清代价（命令行、后台任务这些入口没了）');
    assert.ok(/派工不受影响/.test(text), '要写明检索派工不受影响 —— 这是最容易写错的一条');
    assert.ok(!/退回内置通道/.test(text), '不许写「派工退回内置通道」：与实况不符');
    return '默认未勾选，代价写在说明里且与实况一致';
});

await check('Python 计算与绘图：四个控件都在，读到的是设置里的值而不是兜底常量', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    for (const label of ['启用', '解释器路径', '运行超时', '产物目录']) {
        assert.ok(rowInGroup(tree, 'python', label), `python 组缺少「${label}」这一行`);
    }
    // 解释器是文本输入：样例给的绝对路径要原样显示（不是 placeholder 的「留空自动探测」）。
    const bin = controlOf(rowInGroup(tree, 'python', '解释器路径'));
    assert.equal(bin.type, 'input');
    assert.equal(bin.props.defaultValue, 'C:\\Python313\\python.exe');
    assert.equal(bin.props.placeholder, '留空 = 自动探测');
    const outDir = controlOf(rowInGroup(tree, 'python', '产物目录'));
    assert.equal(outDir.props.defaultValue, 'python/out', '产物目录要显示设置里的值');
    const timeout = controlOf(rowInGroup(tree, 'python', '运行超时'));
    // NumberField 的受控值是字符串（`useState(String(props.value))`），所以比的是 '60000'。
    assert.equal(timeout.props.value, '60000', '超时要显示设置里的值（60 秒），不是默认 120 秒');
    assert.equal(timeout.props.max, 900000, '上界要与服务端 schema 一致');
    return '四个控件都读到设置里的值';
});

await check('Python 计算与绘图：解释器路径失焦写回 python.bin（留空也写，自动探测要能改回来）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    const bin = controlOf(rowInGroup(tree, 'python', '解释器路径'));

    // 写回值故意用正斜杠：Windows 路径里的 `\P` / `\p` 不是合法转义序列，
    // 写在源码里容易被解析器吞掉（本轮踩过：`'  D:\\Python\\python.exe  '` 变成了空串，
    // 断言报的是「写回空值」，看着像功能坏了）。插件的写回是原样字符串，与分隔符无关。
    bin.props.onBlur(eventOf({ value: '  D:/Python/python.exe  ' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['python.bin', 'D:/Python/python.exe'],
        '两端空格要收掉');

    bin.props.onBlur(eventOf({ value: '' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['python.bin', ''],
        '清空要真的写回空串（= 回到自动探测），不能因为「空 = 没改」就吞掉');
    return '写回 python.bin，两端截断、空值照写';
});

await check('Python 计算与绘图：关掉启用开关写回 python.enabled', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    const owner = rowInGroup(tree, 'python', '启用');
    assert.ok(owner, 'python 组应有「启用」这一行');
    const toggle = walk(owner).find((node) => node.type === 'input' && node.props && node.props.type === 'checkbox');
    assert.ok(toggle, '「启用」行里应有勾选框');
    assert.equal(toggle.props.checked, true, '样例里是开着的');
    // Toggle 的 DOM 处理器读 event.target.checked，再把**布尔值**交给面板的 write ——
    // 所以这里给的是事件对象（与既有开关用例同一写法）。
    toggle.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['python.enabled', false]);
    return '写回 python.enabled';
});

// ── 文件识别（OCR，第一百轮）：只写密码框 + 设置值驱动的控件 ──────────────
//
// 这一组的两个硬约束都来自宿主的脱敏契约（dsh-api-settings-controller
// lib/index.js:88-89、:292-297；dsh-settings lib/index.js:73-89）：
//   - 每次远程读都把 role('secret') 字段物理删掉，configForms 快照也没有
//     secrets 旁路，所以界面拿不到已存的 Key，**绝不回填**（回填 = 明文进 DOM）；
//   - 写入是 path op 的 set（覆盖语义），空输入不能写 —— 否则一失焦就把已存的
//     Key 抹掉。下面四条分别钉住「控件值来自设置」「不回填/留空不写/不泄漏」
//     「其余字段写回 ocr.*」「Key 写入仍走 revision 栅栏」。

await check('文件识别（OCR）：七个控件都在，读到的是设置里的值而不是兜底常量', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    for (const label of ['启用', '服务端点', 'DeepSeek API Key', 'Key 的环境变量名', '复用 deepseek-official', '识别精度', '识别超时']) {
        assert.ok(rowInGroup(tree, 'ocr', label), `OCR 组缺少「${label}」这一行`);
    }
    const base = controlOf(rowInGroup(tree, 'ocr', '服务端点'));
    assert.equal(base.type, 'input');
    assert.equal(base.props.type, 'text', '端点用普通文本框（不是密码框）');
    assert.equal(base.props.defaultValue, 'https://api.deepseek.com', '端点要显示设置里的值');
    const env = controlOf(rowInGroup(tree, 'ocr', 'Key 的环境变量名'));
    assert.equal(env.props.defaultValue, 'OCR_TEST_KEY', '环境变量名要显示设置里的值（不是默认的 DEEPSEEK_API_KEY）');
    const purpose = controlOf(rowInGroup(tree, 'ocr', '识别精度'));
    assert.equal(purpose.type, 'select');
    assert.equal(purpose.props.value, 'rich_layout_recognition', '精度要显示设置里的值');
    assert.deepEqual(walk(purpose.props.children).filter((node) => node.type === 'option').map((node) => node.props.value),
        ['text_recognition', 'rich_layout_recognition'], '两个取值要与服务端 OCR_PURPOSES 一致');
    const timeout = controlOf(rowInGroup(tree, 'ocr', '识别超时'));
    assert.equal(timeout.props.value, '240000', '超时要显示设置里的值（240 秒），不是默认 180 秒');
    assert.equal(timeout.props.min, 10000, '下界要与服务端 schema 一致');
    assert.equal(timeout.props.max, 300000);
    assert.equal(timeout.props.step, 5000);
    // 本地 schema 上限 300 秒，官方端点却是 180 秒封顶（src/settings.js:622 的 description、
    // src/ocr.js 的超时文案都是这个口径）。说明里两个数都要出现，否则用户会以为设 300 秒
    // 官方就等 300 秒，或者反过来以为控件写不到 300 秒。
    const timeoutNote = rowInGroup(tree, 'ocr', '识别超时').source.props.note;
    assert.ok(timeoutNote.includes('180') && timeoutNote.includes('300'),
        '超时说明要同时讲清本地范围（300 秒）与官方封顶（180 秒）：' + timeoutNote);
    const toggle = walk(rowInGroup(tree, 'ocr', '启用')).find((node) => node.type === 'input' && node.props && node.props.type === 'checkbox');
    assert.ok(toggle, '「启用」行里应有勾选框');
    assert.equal(toggle.props.checked, true, '样例里是开着的');

    // 复用 deepseek-official 是安全开关：默认关，并且要在原生 input 上留 data 标记，
    // 只读检查据此确认「开关存在」而不必点击、也不去读宿主里存的配置。
    const reuse = walk(rowInGroup(tree, 'ocr', '复用 deepseek-official'))
        .find((node) => node.type === 'input' && node.props && node.props.type === 'checkbox');
    assert.ok(reuse, '「复用 deepseek-official」行应有原生勾选框');
    assert.equal(reuse.props['data-ocr-reuse-official'], 'true', '只读检查靠这个 data 标记确认开关存在');
    assert.equal(reuse.props.checked, false, '样例里是关的（默认关：不能悄悄激活宿主计费）');
    assert.ok(byProp(tree, 'data-ocr-reuse-official', 'true').length >= 1, '整页至少有一个带标记的复用开关');
    const reuseNote = rowInGroup(tree, 'ocr', '复用 deepseek-official').source.props.note;
    assert.ok(reuseNote.includes('默认关') && reuseNote.includes('可信的本机环境'),
        '说明要讲清默认关、只给可信本机环境用：' + reuseNote);
    assert.ok(reuseNote.includes('官方端点') && reuseNote.includes('专用 Key'),
        '仍要讲清官方端点与「专用 Key / 环境变量」优先：' + reuseNote);
    // 不能承诺「网关一定进不来」：来源检查不是身份认证，本机程序或剥掉转发头的代理能绕过。
    assert.ok(reuseNote.includes('来源') && reuseNote.includes('身份认证'),
        '要讲清它只校验请求来源、不是身份认证：' + reuseNote);
    assert.ok(reuseNote.includes('绕过'), '要如实讲清本机程序 / 代理仍可绕过这项检查：' + reuseNote);
    assert.ok(!/已开启|已复用|已配置/.test(reuseNote), '不得凭客户端状态宣称复用已开启');

    // 说明必须讲清「官方端点可复用已存的 deepseek-official 凭据」与「自定义端点不借用」
    // 这两件事 —— 这是用户问过的原话（不必重新输入），也是后端只在官方端点回退的边界。
    // 同时不许出现「已配置」这种没有服务端来源的状态：界面读不到 Key 是否已存
    // （secret 字段在远程读里被删掉，configForms 也没有 secrets 旁路）。
    const ocrStrings = stringsIn(groupBody(tree, 'ocr')).join('\n');
    assert.ok(ocrStrings.includes('deepseek-official'), '要写明可复用已保存的 deepseek-official 凭据');
    assert.ok(ocrStrings.includes('不必重新输入'), '要写明官方端点不必重新输入 Key');
    assert.ok(ocrStrings.includes('两份'), '要写明这份 Key 与对话模型提供方凭据是两份');
    assert.ok(ocrStrings.includes('自定义端点') && ocrStrings.includes('永远不借用'),
        '要写明自定义端点不借用 deepseek-official（否则用户会以为换端点也能白嫖）');
    assert.ok(!/已配置|已保存 Key|Key 已/.test(ocrStrings),
        '不得凭客户端状态宣称 Key 已配置：宿主读回来的 secret 是脱敏的，没有可信来源');
    return '六个控件都读到设置里的值，复用规则与「不显示状态」都写在说明里';
});

await check('文件识别（OCR）：Key 是只写密码框 —— 不回填、留空不写、失焦即清、明文不进 DOM', () => {
    // 样例里故意塞一个「已存的 Key」：真实宿主读回来时这个字段会被删掉，
    // 界面不管拿不拿得到都不许把它渲染出来。
    const value = Object.assign({}, SAMPLE_VALUE, {
        ocr: Object.assign({}, SAMPLE_VALUE.ocr, { apiKey: 'sk-ocr-should-never-render' }),
    });
    const scope = makeScope(value);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    const keyInput = byProp(tree, 'data-ocr-key', 'true').pop();
    assert.ok(keyInput, 'OCR 组应有 Key 输入框');
    assert.equal(keyInput.props.type, 'password', 'Key 必须是密码框，不能是明文输入框');
    assert.equal(keyInput.props.defaultValue, '', '绝不回填已存的 Key（宿主读回来也是脱敏的）');
    assert.equal(keyInput.props.autoComplete, 'new-password',
        '要用 new-password 挡自动填充：off 只是提示，浏览器常忽略，会把已存的密码填进来');

    const strings = textsOf(tree).join('\n');
    assert.ok(!strings.includes('sk-ocr-should-never-render'), '明文不得出现在任何文案里');
    for (const prop of ['value', 'defaultValue', 'placeholder']) {
        assert.equal(byProp(tree, prop, 'sk-ocr-should-never-render').length, 0, `明文不得出现在 ${prop} 属性里`);
    }

    const before = scope.writes.length;
    const blank = eventOf({ value: '   ' });
    keyInput.props.onBlur(blank);
    assert.equal(scope.writes.length, before, '留空不得写回：set 是覆盖语义，写空会把已存的 Key 抹掉');
    assert.equal(blank.target.value, '', '留空失焦也要把输入框清掉，别留下空白');

    const typed = eventOf({ value: '  sk-ocr-new-key  ' });
    keyInput.props.onBlur(typed);
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['ocr.apiKey', 'sk-ocr-new-key'],
        '两端空格要收掉，且只写 ocr.apiKey 这一个字段');
    assert.equal(typed.target.value, '',
        '提交后必须立刻清空输入框：留着值等于把 Key 留在 DOM 里，下次失焦还会再提交一遍');

    // 不受控输入框的经典坑：不清空的话，用户在别处点一下（失焦）就会把同一份 Key
    // 再写一遍；而写入会重读 revision、再走一次鉴权，等于把明文多送一次。
    const afterFirst = scope.writes.length;
    keyInput.props.onBlur(typed);
    assert.equal(scope.writes.length, afterFirst, '同一份值再失焦一次不得重复提交 ocr.apiKey');
    return '密码框不回填，留空不写，提交即清、不重复提交，明文只经过本地输入框';
});

await check('文件识别（OCR）：精度下拉 / 开关 / 端点 / 环境变量名 / 超时都写回 ocr.*', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');

    controlOf(rowInGroup(tree, 'ocr', '识别精度')).props.onChange(eventOf({ value: 'text_recognition' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['ocr.purpose', 'text_recognition']);

    walk(rowInGroup(tree, 'ocr', '启用')).find((node) => node.type === 'input' && node.props && node.props.type === 'checkbox')
        .props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['ocr.enabled', false]);

    controlOf(rowInGroup(tree, 'ocr', '服务端点')).props.onBlur(eventOf({ value: '  https://proxy.example.com/v1  ' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['ocr.baseURL', 'https://proxy.example.com/v1']);

    controlOf(rowInGroup(tree, 'ocr', 'Key 的环境变量名')).props.onBlur(eventOf({ value: '  MY_OCR_KEY  ' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['ocr.apiKeyEnv', 'MY_OCR_KEY'],
        '只写变量名，值不进设置');

    walk(rowInGroup(tree, 'ocr', '复用 deepseek-official'))
        .find((node) => node.type === 'input' && node.props && node.props.type === 'checkbox')
        .props.onChange(eventOf({ checked: true }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['ocr.reuseOfficialProvider', true],
        '复用开关写严格布尔，是否真的去借凭据由后端再判一次');

    // 数字控件是「本地草稿 + 失焦提交」：草稿进 state 后要重渲染，新的 onBlur 才拿得到它。
    controlOf(rowInGroup(tree, 'ocr', '识别超时')).props.onChange(eventOf({ value: '120000' }));
    const afterTyping = renderOfficeTab(registrations, 'docs');
    controlOf(rowInGroup(afterTyping, 'ocr', '识别超时')).props.onBlur();
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['ocr.timeoutMs', 120000]);
    return '六个非敏感字段都写回 ocr.*（含默认关的复用开关）';
});

await check('文件识别（OCR）：Key 写入仍走 revision 栅栏（冲突重读重试），不经过任何网络请求', async () => {
    const fake = makeRemote([5, 6], (n) => (n === 1
        ? { ok: false, error: { code: 'settings/conflict', message: 'revision 5 does not match 6' } }
        : { ok: true }));
    const fetches = [];
    const fakeFetch = (url, init) => {
        fetches.push({ url: String(url), body: init && init.body ? String(init.body) : '' });
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, rows: [] }) });
    };
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, remote: fake.remote, fetch: fakeFetch });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    byProp(tree, 'data-ocr-key', 'true').pop().props.onBlur(eventOf({ value: 'sk-conflict-key' }));
    await new Promise((resolve) => setTimeout(resolve, 20));

    assert.equal(fake.calls.length, 2, '冲突后应恰好重试一次');
    assert.deepEqual(fake.calls[0].ops, [{ op: 'set', path: ['ocr', 'apiKey'], value: 'sk-conflict-key' }],
        '只 patch ocr.apiKey，不整包覆盖别的字段（那会把已存的其它配置一起写回去）');
    assert.equal(fake.calls[0].revision, 5, '第一次带读到的 revision');
    assert.equal(fake.calls[1].revision, 6, '冲突后必须重读 revision，而不是沿用旧的');
    assert.equal(scope.writes.length, 0, 'remote.settings 可用时不得退回 configForms');
    const texts = textsOf(renderOfficeTab(registrations, 'docs'));
    assert.ok(!texts.some((t) => t.includes('写入失败')), '写成功后不应留下错误：' + texts.join(' | '));
    const leaked = fetches.filter((call) => (call.url + call.body).includes('sk-conflict-key'));
    assert.equal(leaked.length, 0, 'Key 不得进入任何网络请求（诊断 / 公开端点）');
    return '冲突重试且只写 ocr.apiKey，没有网络出口';
});

await check('音频与视频：语言下拉覆盖服务端的六种取值', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    const select = controlOf(findByLabel(tree, '转写语言'));
    assert.equal(select.type, 'select');
    assert.equal(select.props.value, 'zh', '样例值给的是 zh');
    const options = walk(select.props.children).filter((node) => node.type === 'option');
    const ids = options.map((node) => node.props.value);
    assert.deepEqual(ids, ['auto', 'zh', 'yue', 'en', 'ja', 'ko'], '语言下拉要覆盖服务端 AV_LANGUAGES');
    return `六个选项：${ids.join(' / ')}`;
});

await check('音频与视频：块间重叠能调到（小数步长，写回 av.overlapSeconds）', () => {
    // 第三十一轮新增：块之间的小重叠是「一个音不被切点切开」的第一道措施，
    // 界面上必须能调；步长是 0.5，不能被取整吃掉（NumberField 有小数步长分支）。
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    const row = findByLabel(tree, '块间重叠');
    assert.ok(row, '应渲染「块间重叠」这一行');
    const input = controlOf(row);
    assert.equal(input.props.type, 'number');
    assert.equal(input.props.min, 0);
    assert.equal(input.props.max, 30);
    assert.equal(input.props.step, 0.5);
    input.props.onChange(eventOf({ value: '1.2' }));
    const after = renderOfficeTab(registrations, 'docs');
    controlOf(findByLabel(after, '块间重叠')).props.onBlur();
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['av.overlapSeconds', 1]);
    return '写回 av.overlapSeconds，0.5 步长吸附';
});

await check('记忆页签：开关与目录写回 memory.<name>；容量控件在召回页签', () => {
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

    const recallTree = renderOfficeTab(registrations, 'recall');
    const limitRow = findByLabel(recallTree, '用户偏好上限');
    assert.ok(limitRow, '容量控件应在召回页签');
    const input = controlOf(limitRow);
    assert.equal(input.props.type, 'number');
    assert.equal(input.props.value, '4096');
    input.props.onChange(eventOf({ value: '8192' }));
    // 数字控件是「本地草稿 + 失焦提交」：草案进 state 之后要重新渲染一次，
    // 新的 onBlur 才拿得到新草稿（真实 React 也是这个次序）。
    const afterTyping = renderOfficeTab(registrations, 'recall');
    const input2 = controlOf(findByLabel(afterTyping, '用户偏好上限'));
    assert.equal(input2.props.value, '8192', '草稿应保留在控件里');
    input2.props.onBlur();
    assert.deepEqual(scope.writes[2], ['memory.userLimitBytes', 8192]);
});

await check('知识库页签：目录与向量/候选筛选/ANN/融合设置写回 memory.*（第八十五轮）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'kb');

    const kbDir = walk(tree).find((n) => n.type === 'input' && n.props
        && String(n.props.placeholder || '').includes('$DSH_HOME/.office/kb'));
    assert.ok(kbDir, '应渲染知识库目录输入框');
    assert.equal(kbDir.props.type, 'text');
    assert.equal(kbDir.props.value, '', '留空 = 默认 $DSH_HOME/.office/kb');
    kbDir.props.onBlur(eventOf({ value: '/data/kb' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['memory.kbDir', '/data/kb']);

    const mode = controlOf(findByLabel(tree, '向量化通道'));
    assert.equal(mode.type, 'select');
    assert.equal(mode.props.value, 'auto');
    assert.deepEqual(walk(mode.props.children).filter((n) => n.type === 'option').map((n) => n.props.value), ['auto', 'remote', 'local', 'off']);
    mode.props.onChange(eventOf({ value: 'remote' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['memory.embedding.mode', 'remote']);

    // 第一百零三轮：重排换成「候选筛选 + 本地推理 + ANN 索引」三组。
    const selectorMode = controlOf(findByLabel(tree, '候选筛选'));
    assert.equal(selectorMode.props.value, 'off');
    selectorMode.props.onChange(eventOf({ value: 'llm' }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['memory.selector.mode', 'llm']);

    const localToggle = controlOf(findByLabel(tree, '本地推理服务'));
    assert.equal(localToggle.props.checked, false, '本地推理默认关');
    localToggle.props.onChange(eventOf({ checked: true }));
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['memory.embedding.local.enabled', true]);

    const annToggle = controlOf(findByLabel(tree, '向量索引'));
    assert.equal(annToggle.props.checked, true, 'ANN 默认开');
    const efSearch = controlOf(findByLabel(tree, '索引 efSearch'));
    assert.equal(Number(efSearch.props.value), 64);
    // NumberField 是「本地草稿 + 失焦提交」：先改草稿，重渲染后 blur 才写回。
    efSearch.props.onChange(eventOf({ value: '128' }));
    const afterAnn = renderOfficeTab(registrations, 'kb');
    controlOf(findByLabel(afterAnn, '索引 efSearch')).props.onBlur();
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['memory.ann.efSearch', 128]);

    const weight = controlOf(findByLabel(tree, '词法权重'));
    assert.equal(weight.props.type, 'number');
    weight.props.onChange(eventOf({ value: '0.7' }));
    const after = renderOfficeTab(registrations, 'kb');
    controlOf(findByLabel(after, '词法权重')).props.onBlur();
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['memory.hybrid.lexicalWeight', 0.7]);
});

/** 第一百零四轮：按 name 取知识库里的折叠块体（收起时只是 display:none）。 */
function foldBodyOf(tree, name) {
    return walk(tree).find((node) => node.props && node.props['data-fold-body'] === name);
}

/** 深拷贝样例值再改 memory，避免一条用例污染下一条。 */
function memorySample(mutate) {
    const value = JSON.parse(JSON.stringify(SAMPLE_VALUE));
    mutate(value.memory);
    return value;
}

/** 深拷贝样例值再改 gallery，避免一条用例污染下一条。 */
function gallerySample(mutate) {
    const value = JSON.parse(JSON.stringify(SAMPLE_VALUE));
    mutate(value.gallery);
    return value;
}

/** 知识库页签里「当前实际通道」那一行：值、路由标识与说明。 */
function retrievalOf(tree) {
    const row = walk(tree).find((node) => node.props && node.props['data-kb-route'] !== undefined);
    const columns = [].concat(row.props.children || []);
    return {
        value: stringsIn(columns[0]).join('').replace('当前实际通道：', ''),
        route: row.props['data-kb-route'],
        note: stringsIn(columns[1]).join(' '),
    };
}

await check('知识库设置：高级参数默认折叠、按通道禁用，常用项不必翻三十行（第一百零四轮）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'kb');

    // 常用项在折叠块外：通道、总开关、筛选、索引。
    for (const label of ['向量化通道', '本地推理服务', '候选筛选', '向量索引']) {
        assert.ok(findByLabel(tree, label), '常用项应直接可见（不在折叠块里）：' + label);
    }
    // 高级参数默认收起，但内容照渲染（收起只是 display:none，定位与单测仍找得到）。
    for (const name of ['kb-remote-params', 'kb-models', 'kb-local-params', 'kb-selector-params', 'kb-ann-params', 'kb-hybrid']) {
        const body = foldBodyOf(tree, name);
        assert.ok(body, '应有折叠块 ' + name);
        assert.deepEqual(body.props.style, { display: 'none' }, name + ' 默认应收起');
    }
    assert.equal(scope.writes.length, 0, '点开折叠块本身不该写设置');

    // 当前是 auto + 本地未开 + 远端未配 + 筛选未开 + ANN 开。
    assert.equal(controlOf(findByLabel(tree, 'embedding 端点')).props.disabled, false, 'auto 仍可能用远端，参数应可编辑');
    assert.equal(controlOf(findByLabel(tree, '本地模型目录')).props.disabled, true, '本地推理未开时本地参数禁用');
    assert.equal(controlOf(findByLabel(tree, '筛选端点')).props.disabled, true, '筛选未开时筛选参数禁用');
    assert.equal(controlOf(findByLabel(tree, '索引 efSearch')).props.disabled, false, 'ANN 默认开，参数可用');
    assert.equal(controlOf(findByLabel(tree, '词法权重')).props.disabled, false, '融合参数始终可用');

    // 点开后真的展开（自管 state，不依赖外层）。先做这一步：下面会再 loadBundle，
    // 测试替身的 currentFake 是模块级变量，跨 bundle 渲染会换掉 hook 存储。
    const toggle = walk(tree).find((node) => node.props && node.props['data-fold-toggle'] === 'kb-selector-params');
    toggle.props.onClick();
    const opened = renderOfficeTab(registrations, 'kb');
    assert.equal(foldBodyOf(opened, 'kb-selector-params').props.style, undefined, '点开折叠块后应展开');
    assert.equal(foldBodyOf(opened, 'kb-local-params').props.style.display, 'none', '只开被点的那一块');
    assert.equal(scope.writes.length, 0, '开合折叠块不写设置');

    // 切到 local：远端整组禁用，本地参数随开关放开；再切 off：本地开关也禁用。
    const localScope = makeScope(memorySample((memory) => {
        memory.embedding.mode = 'local';
        memory.embedding.local.enabled = true;
    }));
    const localBundle = loadBundle({ forms: localScope.forms });
    localBundle.mod.apply(localBundle.ctx);
    const localTree = renderOfficeTab(localBundle.registrations, 'kb');
    assert.equal(controlOf(findByLabel(localTree, 'embedding 端点')).props.disabled, true, 'local 通道下远端参数不生效');
    assert.equal(controlOf(findByLabel(localTree, '本地模型目录')).props.disabled, false, 'local 且已开本地推理，本地参数可编辑');

    const offScope = makeScope(memorySample((memory) => { memory.embedding.mode = 'off'; }));
    const offBundle = loadBundle({ forms: offScope.forms });
    offBundle.mod.apply(offBundle.ctx);
    const offTree = renderOfficeTab(offBundle.registrations, 'kb');
    assert.equal(controlOf(findByLabel(offTree, '本地推理服务')).props.disabled, true, 'off 通道下不该还能开本地推理');
    assert.equal(controlOf(findByLabel(offTree, '候选筛选')).props.disabled, true, 'off 通道下没有候选可筛');

});

await check('知识库实际通道：按配置显示，不再一律「自动混合」（第一百零四轮）', () => {
    const overview = (mutate) => {
        const scope = makeScope(memorySample(mutate));
        const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
        mod.apply(ctx);
        return retrievalOf(renderOfficeTab(registrations, 'kb'));
    };

    const local = overview((memory) => { memory.embedding.local.enabled = true; });
    assert.equal(local.value, '本地推理');
    assert.equal(local.route, 'local');
    assert.match(local.note, /本地推理已开启/);

    const off = overview((memory) => { memory.embedding.mode = 'off'; memory.embedding.local.enabled = true; });
    assert.equal(off.value, '纯词法', 'off 时本地推理开关不该改变实际通道');
    assert.equal(off.route, 'off');

    const remote = overview((memory) => {
        memory.embedding.mode = 'remote';
        memory.embedding.baseURL = 'https://api.example.com/v1';
        memory.embedding.model = 'bge-m3';
    });
    assert.equal(remote.value, '远端向量');
    assert.equal(remote.route, 'remote');

    const remoteMissing = overview((memory) => { memory.embedding.mode = 'remote'; });
    assert.equal(remoteMissing.value, '纯词法', 'remote 端点/模型没配齐只能退回词法');
    assert.equal(remoteMissing.route, 'none');

    const hash = overview((memory) => { memory.embedding.mode = 'local'; });
    assert.equal(hash.value, '本地哈希');
    assert.equal(hash.route, 'local-hash');

    const disabled = overview((memory) => { memory.enabled = false; });
    assert.equal(disabled.value, '已关闭');
    assert.equal(disabled.route, 'disabled');

    const selected = overview((memory) => {
        memory.embedding.local.enabled = true;
        memory.selector.mode = 'llm';
        memory.selector.baseURL = 'https://api.example.com/v1';
        memory.selector.model = 'selector-small';
    });
    assert.match(selected.note, /筛选模型保留前几条/, '筛选真的可用时应写进说明');
});

await check('知识库大模型：off 只留一句说明，切到 openai 才出现端点与 Key（第一百零四轮）', () => {
    const value = memorySample((memory) => { memory.llm = { mode: 'off' }; });
    const scope = makeScope(value);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const off = renderOfficeTab(registrations, 'kb');
    assert.equal(byProp(off, 'data-kb-llm-off').length, 1, 'off 时给一句「未启用」而不是十行输入框');
    assert.equal(byProp(off, 'data-kb-llm-on').length, 0);
    for (const label of ['单次超时', '单次原文上限', '总预算', '大模型 Key']) {
        assert.equal(findByLabel(off, label), undefined, 'off 时不该渲染 ' + label);
    }

    const hostValue = memorySample((memory) => { memory.llm = { mode: 'host' }; });
    const hostScope = makeScope(hostValue);
    const hostBundle = loadBundle({ forms: hostScope.forms });
    hostBundle.mod.apply(hostBundle.ctx);
    const host = renderOfficeTab(hostBundle.registrations, 'kb');
    assert.equal(byProp(host, 'data-kb-llm-on').length, 1);
    assert.ok(findByLabel(host, '单次超时'), 'host 通道仍可调超时与预算');
    assert.equal(findByLabel(host, '大模型 Key'), undefined, 'host 用宿主凭据，不摆 Key 输入框');

    const openaiValue = memorySample((memory) => { memory.llm = { mode: 'openai' }; });
    const openaiScope = makeScope(openaiValue);
    const openaiBundle = loadBundle({ forms: openaiScope.forms });
    openaiBundle.mod.apply(openaiBundle.ctx);
    const openai = renderOfficeTab(openaiBundle.registrations, 'kb');
    assert.ok(findByLabel(openai, '端点'), 'openai 通道要能填端点');
    assert.ok(findByLabel(openai, '大模型 Key'), 'openai 通道要能填 Key');
});

await check('三处 Key 环境变量不再同名：向量 / 筛选 / 大模型各有前缀（第一百零四轮）', () => {
    const value = memorySample((memory) => {
        memory.embedding.local.enabled = true;
        memory.embedding.baseURL = 'https://api.example.com/v1';
        memory.embedding.model = 'bge-m3';
        memory.selector.mode = 'llm';
        memory.selector.baseURL = 'https://api.example.com/v1';
        memory.selector.model = 'selector-small';
        memory.llm = { mode: 'openai', baseURL: 'https://api.example.com/v1' };
    });
    const scope = makeScope(value);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const labels = rowLabels(renderOfficeTab(registrations, 'kb'));
    assert.equal(labels.filter((label) => label === 'Key 环境变量').length, 0, '不该再有裸的「Key 环境变量」');
    for (const label of ['向量 Key 环境变量', '筛选 Key 环境变量', '大模型 Key 环境变量']) {
        assert.equal(labels.filter((item) => item === label).length, 1, label + ' 应唯一');
    }
});

await check('记忆系统面板：一次性迁移入口已移除，系统提示开关仍在', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    const hintRow = findByLabel(tree, '写入系统提示');
    assert.ok(hintRow, '其它模式仍需要这段静态说明，开关保留');
    const hint = controlOf(hintRow);
    assert.equal(hint.props.checked, true);
    hint.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[0], ['memory.promptHint', false]);

    const texts = walk(tree).map((node) => node.props && node.props.children).flat().filter((c) => typeof c === 'string');
    assert.ok(!texts.some((t) => t.includes('node scripts/migrate-mnemon.mjs')),
        '一次性迁移不再占记忆设置页；命令仍留在 office_help 的 memory 文档里');
    assert.ok(!texts.some((t) => t.includes("action: 'migrate'")), '迁移说明不应再出现在设置页');
});

await check('合并页：两个页签共用同一个设置命名空间，控件不重复', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    // bind 只发生一次：两页读同一份设置（分开 bind 会各自订阅一份镜像）。
    assert.equal(scope.forms.bound.entryId, 'dsh-office-mode');

    const office = render(officeReg(registrations));
    const memory = renderMemory(registrations);
    assert.ok(findByLabel(office, 'office_run'), '办公模式页有工具开关');
    assert.ok(!findByLabel(office, '开启记忆'), '记忆控件不该在办公模式页重复出现');
    assert.equal(walk(office).filter((node) => node.props && node.props['data-locator-note'] !== undefined).length, 0,
        '办公模式页不再画一张没有分组头的裸指路卡');
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

await check('主动记录开关写回 memory.autoCapture（图关系开关已随图谱下线）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderMemory(registrations);

    assert.equal(findByLabel(tree, '图关系'), undefined, '图关系开关应已删除');
    const capture = controlOf(findByLabel(tree, '主动记录'));
    assert.ok(capture, '应渲染「主动记录」这一行');
    assert.equal(capture.props.checked, true);
    capture.props.onChange(eventOf({ checked: false }));
    assert.deepEqual(scope.writes[0], ['memory.autoCapture', false]);
});

await check('召回页签：策略下拉 + 五个数字，阈值按 0.05 吸附而不是取整', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'recall');

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
    const after = renderOfficeTab(registrations, 'recall');
    controlOf(findByLabel(after, '低分阈值')).props.onBlur();
    assert.deepEqual(scope.writes[1], ['memory.recallQuality.lowScoreThreshold', 0.35]);
});

await check('容量上限（召回页签）：整数步长吸附到网格，min=512 / step=512 输 1000 得到 1024', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'recall');

    const row = findByLabel(tree, '用户偏好上限');
    assert.ok(row, '应渲染「用户偏好上限」这一行');
    const input = controlOf(row);
    assert.equal(input.props.min, 512);
    assert.equal(input.props.step, 512);
    // 关键回归：整数步长原来只做 Math.trunc，1000 会原样写回 —— 而 512 + k×512 的网格上
    // 没有 1000，schema 会拒掉它，界面看着像「改了但没生效」。吸附之后最近的格点是 1024。
    input.props.onChange(eventOf({ value: '1000' }));
    const after = renderOfficeTab(registrations, 'recall');
    controlOf(findByLabel(after, '用户偏好上限')).props.onBlur();
    assert.deepEqual(scope.writes[scope.writes.length - 1], ['memory.userLimitBytes', 1024]);
    // 网格上已有的值不该被挪动。
    const again = renderOfficeTab(registrations, 'recall');
    controlOf(findByLabel(again, '用户偏好上限')).props.onChange(eventOf({ value: '4096' }));
    const latest = renderOfficeTab(registrations, 'recall');
    const writes = scope.writes.length;
    controlOf(findByLabel(latest, '用户偏好上限')).props.onBlur();
    assert.equal(scope.writes.length, writes, '值本来就在网格上时不该多写一次');
});

await check('每回合配额（召回页签）：四个数字写回 memory.quota.<项>', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);

    const rows = [
        ['首次检索', 'memory.quota.recallPerTurn'],
        ['换词细化', 'memory.quota.recallRefinePerTurn'],
        // 第二十四轮 24-11：kb 检索单独一档（语料与代价都与记忆检索不同，不能共用名额）。
        ['知识块检索', 'memory.quota.kbSearchPerTurn'],
        ['知识库块关系遍历', 'memory.quota.relatedPerTurn'],
    ];
    for (const [label, path] of rows) {
        const tree = renderOfficeTab(registrations, 'recall');
        const row = findByLabel(tree, label);
        assert.ok(row, `应渲染「${label}」这一行`);
        const input = controlOf(row);
        assert.equal(input.props.type, 'number');
        assert.equal(input.props.min, 0);
        assert.equal(input.props.max, 50);
        input.props.onChange(eventOf({ value: '3' }));
        const after = renderOfficeTab(registrations, 'recall');
        controlOf(findByLabel(after, label)).props.onBlur();
        assert.deepEqual(scope.writes[scope.writes.length - 1], [path, 3]);
    }
});

await check('子代理模型：模式下拉与两个文本框写回 subagentModel.<项>', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'search');

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

// ── 新-21：办公模式插件（第六十三轮） ─────────────────────────────────────
//
// 这一格的接缝是**两条 HTTP 端点**（不是 configForms）：列表从磁盘读、开关写回
// preset 文件。所以这里用 `loadBundle({ fetch })` 注入替身，按「GET / POST」分流回执，
// 验四件事：打开就 GET 一次、列出的行 = 端点给的行、开关 POST 的 body 与回执提示、
// 以及「profile 侧没有覆盖行」这种边界照实写在行说明里。
//
// 注意 fake React 的 effect 只在组件实例第一次渲染时跑，而 fetch 的回执是异步落地的：
// 每条用例都要「渲染 → 等一拍 → 再渲染」，第二次才看得到那一格的内容。

const PRESET_PRESET_PATH = 'D:/ws/presets/office/cordis.patch.yml';
const PRESET_PROFILE_PATH = 'C:/Users/someone/.dsh/profiles/web/cordis.patch.yml';

/** 造的插件行（含「profile 侧有 / 没有覆盖行」两种）。 */
function presetRows() {
    return [
        { id: 'persona', name: '@deepseek-ai/dsh-persona', enabled: true, declared: 'default', profileDisabled: false, profileOverridden: true },
        { id: 'tool-bash', name: '@deepseek-ai/dsh-tool-bash', enabled: false, declared: 'disabled', profileDisabled: false, profileOverridden: true },
        { id: 'ui-task-board', name: '@linxin666/dsh-client-ui-task-board', enabled: false, declared: 'disabled', profileDisabled: null, profileOverridden: false },
    ];
}

/** 一个按方法分流的假端点；回执字段照 src/preset-plugins.js 的冻结契约。 */
function presetEndpoint(rows = presetRows()) {
    const calls = [];
    const listBody = (extra = {}) => Object.assign({
        ok: true,
        rows: rows,
        presetPath: PRESET_PRESET_PATH,
        profilePath: PRESET_PROFILE_PATH,
        profileName: 'web',
        backup: { path: `${PRESET_PRESET_PATH}.bak`, name: 'preset-20261013-021010-000.json', at: '2026-10-13T02:10:10.000Z', size: 1024 },
        backups: 1,
        effect: '新建的办公会话',
    }, extra);
    const impl = (url, init) => {
        calls.push({ url: String(url), init: init });
        if (init && init.method === 'POST' && String(url).endsWith('/office-preset/restore')) {
            return Promise.resolve({ status: 200, json: () => Promise.resolve(listBody({
                restored: [PRESET_PRESET_PATH, PRESET_PROFILE_PATH],
                restoredAt: '2026-10-13T02:10:10.000Z',
            })) });
        }
        if (init && init.method === 'POST') {
            const payload = JSON.parse(init.body);
            const next = rows.map((row) => (row.id === payload.id ? Object.assign({}, row, { enabled: payload.enabled === true }) : row));
            rows = next;
            return Promise.resolve({ status: 200, json: () => Promise.resolve(listBody({
                applied: {
                    id: payload.id,
                    enabled: payload.enabled === true,
                    changed: { preset: true, profile: true },
                    profileOverridden: true,
                    diff: '只改了目标行的 disabled。',
                    snapshot: 'C:/Users/someone/.dsh/.office/preset-backups/preset-20261013-021010-000.json',
                },
            })) });
        }
        return Promise.resolve({ status: 200, json: () => Promise.resolve(listBody()) });
    };
    return { calls, impl, empty: () => () => Promise.resolve({ status: 200, json: () => Promise.resolve({ ok: false, error: '找不到办公 preset 的 bundle 目录。' }) }) };
}

/** 渲染办公模式那一页并把异步取数等落地。 */
async function renderOfficeWith(preset) {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, fetch: preset.impl });
    mod.apply(ctx);
    render(officeReg(registrations));
    await new Promise((resolve) => setTimeout(resolve, 0));
    const tree = render(officeReg(registrations));
    return { tree, scope, registrations, mod, ctx };
}

await check('办公模式插件：打开这一格就 GET 一次，行数 = 端点给的行数，说明里带两处路径', async () => {
    const preset = presetEndpoint();
    const { tree } = await renderOfficeWith(preset);
    const got = preset.calls.filter((call) => call.init === undefined || call.init.method === undefined);
    assert.ok(got.length >= 1, '打开这一页应当 GET 一次');
    assert.match(got[0].url, /\/office-preset\/plugins$/);
    const toggles = byProp(tree, 'data-preset-plugin');
    assert.deepEqual(toggles.map((node) => node.props['data-preset-plugin']), ['persona', 'tool-bash', 'ui-task-board']);
    assert.deepEqual(toggles.map((node) => node.props['data-preset-enabled']), ['true', 'false', 'false']);
    const paths = byProp(tree, 'data-preset-paths')[0];
    assert.ok(paths, '应写出「写到哪两处」');
    const text = stringsIn(paths).join(' ');
    assert.ok(text.includes(PRESET_PRESET_PATH) && text.includes(PRESET_PROFILE_PATH), text);
    const backup = byProp(tree, 'data-preset-backup')[0];
    assert.ok(backup, '有备份时应显示备份名与还原说明');
    assert.match(stringsIn(backup).join(' '), /preset-20261013-021010-000\.json/);
    // 只有 profile 侧真有覆盖行时才写；没有就不补一句「没有」
    const noOverride = findByLabel(tree, 'ui-task-board');
    assert.doesNotMatch(stringsIn(noOverride).join(' '), /profile 侧/);
    const withOverride = findByLabel(tree, 'tool-bash');
    assert.match(stringsIn(withOverride).join(' '), /profile 侧有同名覆盖行，开关两处同写/);
    return `${toggles.length} 行`;
});

await check('办公模式插件：开关 POST { id, enabled }，带 JSON content-type', async () => {
    const preset = presetEndpoint();
    const { tree } = await renderOfficeWith(preset);
    const row = findByLabel(tree, 'tool-bash');
    const toggle = controlOf(row);
    assert.equal(toggle.props.checked, false, 'tool-bash 当前是停用');
    toggle.props.onChange(eventOf({ checked: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));
    const post = preset.calls.find((call) => call.init && call.init.method === 'POST');
    assert.ok(post, '开关应当 POST 出去');
    assert.match(post.url, /\/office-preset\/plugins$/);
    assert.deepEqual(JSON.parse(post.init.body), { id: 'tool-bash', enabled: true });
    assert.equal(post.init.headers['content-type'], 'application/json');
    return 'POST 到位';
});

await check('办公模式插件：开关后界面状态即时变化（回执里的新列表直接上屏）', async () => {
    const preset = presetEndpoint();
    const { mod, ctx, registrations } = loadBundle({ forms: makeScope(SAMPLE_VALUE).forms, fetch: preset.impl });
    mod.apply(ctx);
    const reg = officeReg(registrations);
    render(reg);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const before = render(reg);
    controlOf(findByLabel(before, 'tool-bash')).props.onChange(eventOf({ checked: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));
    const after = render(reg);
    const flipped = byProp(after, 'data-preset-plugin', 'tool-bash')[0];
    assert.equal(flipped.props['data-preset-enabled'], 'true', '回执里的新状态要立刻反映在开关上');
    const hint = byProp(after, 'data-preset-hint')[0];
    assert.ok(hint, '写入成功后要有一句回执');
    const text = stringsIn(hint).join(' ');
    assert.match(text, /已启用/);
    assert.match(text, /preset 真源 \+ profile 覆盖行/);
    assert.match(text, /新建的办公会话/);
    assert.match(text, /已经在跑的会话不受影响/);
    return text.slice(0, 40);
});

await check('办公模式插件：端点报错时把服务端给的原因显示出来，不画死开关', async () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, fetch: presetEndpoint().empty() });
    mod.apply(ctx);
    const reg = officeReg(registrations);
    render(reg);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const tree = render(reg);
    const error = byProp(tree, 'data-preset-error')[0];
    assert.ok(error, '失败时必须有一行错误说明');
    assert.match(stringsIn(error).join(' '), /找不到办公 preset 的 bundle 目录/);
    return '照实报错';
});

await check('办公模式插件：「还原上一次备份」POST 到还原端点，没有备份时按钮是禁用的', async () => {
    const preset = presetEndpoint();
    const { tree } = await renderOfficeWith(preset);
    const buttons = byProp(tree, 'data-preset-restore');
    assert.equal(buttons.length, 1);
    assert.equal(buttons[0].props.disabled, false, '有备份时可点');
    buttons[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 0));
    const call = preset.calls.find((item) => item.init && item.init.method === 'POST' && item.url.endsWith('/office-preset/restore'));
    assert.ok(call, '应当 POST 到还原端点');
    // 没有备份时禁用（同一个组件实例换一份回执）
    const empty = presetEndpoint();
    empty.impl = (url, init) => {
        empty.calls.push({ url: String(url), init: init });
        return Promise.resolve({ status: 200, json: () => Promise.resolve({ ok: true, rows: presetRows(), presetPath: PRESET_PRESET_PATH, profilePath: null, backup: null, backups: 0, effect: '新建的办公会话' }) });
    };
    const second = await renderOfficeWith(empty);
    assert.equal(byProp(second.tree, 'data-preset-restore')[0].props.disabled, true, '没有备份时不可点');
    return '还原端点 + 禁用态';
});

// ── 新增：分组折叠与「定位」（第三十五轮） ────────────────────────────────
//
// 这一层做在**元素树**上（不是 DOM 查询）：分组的开合由 data-group-open 表达，
// 组体收起时只是 display:none（控件仍在树里，所以「界面覆盖服务端每个参数」
// 那几条断言不用先展开每一组）。定位过滤则**真的把不命中的行摘掉**，
// 所以下面用「剩下的行标题」来钉住过滤口径。

/** 造一个假的 localStorage（Node 里没有 window.localStorage）。 */
function makeLocalStorage() {
    const map = new Map();
    return {
        map,
        getItem: (key) => (map.has(key) ? map.get(key) : null),
        setItem: (key, value) => { map.set(key, String(value)); },
        removeItem: (key) => { map.delete(key); },
    };
}

/** 某个分组的折叠头按钮。 */
function groupToggle(tree, id) {
    return walk(tree).find((node) => node.props && node.props['data-group-toggle'] === id);
}

/** 某个分组的组体（收起时 style 为 display:none）。 */
function groupBody(tree, id) {
    return walk(tree).find((node) => node.props && node.props['data-group-body'] === id);
}

/** 分组的开合态：读的是渲染出来的 data-group-open，不是内部 state。 */
function groupOpen(tree, id) {
    const box = walk(tree).find((node) => node.props && node.props['data-group'] === id);
    return box === undefined ? undefined : box.props['data-group-open'];
}

/** 把一棵渲染树里的字符串全铺出来（子元素是数组，直接取 props.children 会漏）。 */
function stringsIn(node, out = []) {
    if (typeof node === 'string') { out.push(node); return out; }
    if (Array.isArray(node)) { for (const item of node) stringsIn(item, out); return out; }
    if (node !== null && typeof node === 'object' && node.props) stringsIn(node.props.children, out);
    return out;
}

/** 折叠头上那一串「N 项」/「a / b 项」。 */
function groupCount(tree, id) {
    const toggle = groupToggle(tree, id);
    return stringsIn(toggle).find((text) => /\d+ 项$/.test(text));
}

/** 当前树里还有哪些设置行（按 Row 收到的 label）。 */
function rowLabels(tree) {
    return walk(tree).filter((node) => node.source && node.source.name === 'Row')
        .map((node) => node.source.props.label);
}

/**
 * 「某个分组里的某一行」，取**最后一次渲染**的那一个。
 *
 * 两个坑都在这一行里：
 * ① 不能用 `findByLabel` 全局找：「启用」这种短标签在办公模式页出现多次
 *    （Python 与音频与视频各一行），全局找只会打到最先渲染的那一行；
 * ② 要用**最后一个**匹配节点而不是第一个：render() 为了收敛 setState 会重复展开
 *    （NumberField 的 useState+useEffect 就会触发一次重渲染），walk 会把每一轮的节点
 *    都收进来。第一轮那份的闭包里读到的还是空值 —— 拿它去触发 onBlur，
 *    写回的是空串，而断言看起来像「功能没生效」（本轮踩过）。
 * 组体是收起状态也照样找得到（收起只用 display:none，内容照渲染）。
 */
function rowInGroup(tree, group, label) {
    const body = groupBody(tree, group);
    if (body === undefined) return undefined;
    const found = walk(body).filter((node) => node.source && node.source.name === 'Row'
        && node.source.props.label === label);
    return found[found.length - 1];
}

/** 本页所有分组折叠头的 id（顺序即页面顺序）。 */
function groupIds(tree) {
    return walk(tree).filter((node) => node.props && typeof node.props['data-group-toggle'] === 'string')
        .map((node) => node.props['data-group-toggle']);
}

/** 本页定位框/状态行的元素。 */
function locatorInput(tree, page) {
    return walk(tree).find((node) => node.props && node.props['data-locator-input'] === page);
}

function locatorText(tree, page) {
    const line = walk(tree).find((node) => node.props && node.props['data-locator-stats'] === page);
    assert.ok(line, '应有定位状态行');
    return stringsIn(line).join('');
}

/** 把关键词打进定位框，返回重渲染后的树。 */
function locate(registration, tree, page, query) {
    locatorInput(tree, page).props.onChange(eventOf({ value: query }));
    return render(registration);
}

// 页签 -> 分组。与 lib/client.js 的 SETTINGS_TABS 保持一致；测试不读源码常量，
// 这样页签分错了会当场红，而不是「按同一份错误映射自证正确」。
const TAB_GROUPS = {
    office: ['plugins', 'tools', 'guard'],
    search: ['search', 'subagents'],
    docs: ['documents', 'ocr', 'python', 'av'],
    gallery: ['gallery', 'gallery-image'],
    memory: ['memory', 'layers', 'scope', 'capture'],
    kb: ['kb', 'kb-llm'],
    recall: ['capacity', 'recall', 'quota'],
};

/** 页签对应的定位框 page 标识：办公类四页共用 office，记忆类三页共用 memory。 */
function tabPage(tab) {
    if (tab === 'gallery') return 'gallery';
    return tab === 'memory' || tab === 'kb' || tab === 'recall' ? 'memory' : 'office';
}

await check('页签：每页只渲染自己那几组，且组默认展开', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);

    for (const [tab, ids] of Object.entries(TAB_GROUPS)) {
        const tree = renderOfficeTab(registrations, tab);
        assert.deepEqual(groupIds(tree), ids, tab + ' 页只该有这几组');
        for (const id of ids) {
            assert.equal(groupOpen(tree, id), 'true', tab + ' / ' + id + ' 默认展开');
            assert.equal(groupBody(tree, id).props.style, undefined, tab + ' / ' + id + ' 展开后不设 display:none');
        }
        assert.ok(locatorInput(tree, tabPage(tab)), tab + ' 页应有搜索框');
        assert.equal(locatorInput(tree, tabPage(tab)).props.value, '', tab + ' 页搜索框初始为空');
        assert.match(locatorText(tree, tabPage(tab)), /共 \d+ 项设置/, tab + ' 页给出本页总数');
    }
});

await check('页签：点击切换只画对应分组，并记住页签', () => {
    const store = makeLocalStorage();
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, localStorage: store });
    mod.apply(ctx);
    const reg = settingsReg(registrations);

    const first = render(reg);
    assert.equal(byProp(first, 'data-settings-tab', 'overview')[0].props['aria-selected'], true, '默认停在总览');
    assert.ok(byProp(first, 'data-preset-home')[0], '总览画模式卡');

    byProp(first, 'data-settings-tab', 'search')[0].props.onClick();
    const second = render(reg);
    assert.deepEqual(groupIds(second), ['search', 'subagents'], '切到检索页签只画检索两组');
    assert.equal(byProp(second, 'data-settings-tab', 'search')[0].props['aria-selected'], true);
    assert.equal(byProp(second, 'data-settings-tab', 'overview')[0].props['aria-selected'], false);
    assert.equal(store.getItem('dsh-office-mode.settings-tab.v1'), 'search', '页签写进 localStorage');
});

await check('页签：方向键与 Home/End 走标准 tablist 键盘（第一百一十一轮补）', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const reg = settingsReg(registrations);
    const first = render(reg);
    const tabOf = (tree, id) => byProp(tree, 'data-settings-tab', id)[0];
    assert.equal(tabOf(first, 'overview').props.tabIndex, 0, '当前页签在 Tab 顺序里');
    assert.equal(tabOf(first, 'office').props.tabIndex, -1, '其余页签靠方向键到达');

    tabOf(first, 'overview').props.onKeyDown({ key: 'ArrowRight', preventDefault() {} });
    const second = render(reg);
    assert.equal(tabOf(second, 'office').props['aria-selected'], true, '右方向键切到下一个页签');
    assert.equal(tabOf(second, 'office').props.tabIndex, 0, '焦点页签可再次 Tab 进入');

    tabOf(second, 'office').props.onKeyDown({ key: 'End', preventDefault() {} });
    const third = render(reg);
    assert.equal(tabOf(third, 'recall').props['aria-selected'], true, 'End 到最后一个页签');

    tabOf(third, 'recall').props.onKeyDown({ key: 'Home', preventDefault() {} });
    const fourth = render(reg);
    assert.equal(tabOf(fourth, 'overview').props['aria-selected'], true, 'Home 回第一个页签');

    tabOf(fourth, 'overview').props.onKeyDown({ key: 'ArrowLeft', preventDefault() {} });
    const fifth = render(reg);
    assert.equal(tabOf(fifth, 'recall').props['aria-selected'], true, '左方向键从第一个绕到最后一个');
});

await check('模式：点「轻量」一次写 16 项超参', async () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const reg = withTab(settingsReg(registrations), 'overview');
    const tree = render(reg);

    const button = byProp(tree, 'data-preset-apply', 'lite')[0];
    assert.ok(button, '应有「轻量」模式卡');
    button.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.equal(scope.writes.length, 16, '一套模式正好 16 项，不多写凭据与路径');
    const got = new Map(scope.writes);
    const expected = {
        'memory.embedding.mode': 'off',
        'memory.embedding.local.enabled': false,
        'memory.selector.mode': 'off',
        'memory.ann.enabled': false,
        'memory.ann.efSearch': 64,
        'memory.hybrid.lexicalWeight': 1,
        'memory.hybrid.vectorWeight': 0,
        'memory.hybrid.candidateLimit': 20,
        'memory.recallQuality.policy': 'strict-v1',
        'memory.recallQuality.candidateMultiplier': 3,
        'memory.recallQuality.maxMediumResults': 3,
        'memory.recallQuality.maxUnknownResults': 1,
        'memory.quota.recallPerTurn': 1,
        'memory.quota.recallRefinePerTurn': 1,
        'memory.quota.kbSearchPerTurn': 1,
        'memory.quota.relatedPerTurn': 1,
    };
    for (const [path, value] of Object.entries(expected)) {
        assert.equal(got.get(path), value, path + ' 应写成轻量档的值');
    }

    const after = render(reg);
    const hint = byProp(after, 'data-preset-hint')[0];
    assert.ok(hint, '应给出切换回执');
    assert.match(stringsIn(hint).join(''), /已切换到「轻量」/);
});

await check('模式：保存当前 16 项为自定义模式，跨刷新可再用、可删除', async () => {
    const store = makeLocalStorage();
    const scope = makeScope(SAMPLE_VALUE);
    const first = loadBundle({ forms: scope.forms, localStorage: store });
    first.mod.apply(first.ctx);
    const reg = settingsReg(first.registrations);

    let tree = render(reg);
    byProp(tree, 'data-preset-name')[0].props.onChange(eventOf({ value: '我的模式' }));
    tree = render(reg);
    byProp(tree, 'data-preset-save')[0].props.onClick();
    tree = render(reg);

    const saved = JSON.parse(store.getItem('dsh-office-mode.hyper-presets.v1'));
    assert.equal(saved.length, 1, '只存一个自定义模式');
    assert.equal(saved[0].name, '我的模式');
    assert.equal(saved[0].values['memory.embedding.mode'], 'auto');
    assert.equal(saved[0].values['memory.quota.kbSearchPerTurn'], 2, '样例缺的字段用均衡档补齐');
    assert.ok(byProp(tree, 'data-preset-card', 'custom:我的模式')[0], '保存后立即多一张卡');

    // 新开一个 bundle 读同一份 localStorage：自定义模式还在，并且能再次应用。
    const secondScope = makeScope(SAMPLE_VALUE);
    const second = loadBundle({ forms: secondScope.forms, localStorage: store });
    second.mod.apply(second.ctx);
    let reloaded = render(settingsReg(second.registrations));
    assert.ok(byProp(reloaded, 'data-preset-card', 'custom:我的模式')[0], '跨刷新恢复自定义模式');
    byProp(reloaded, 'data-preset-apply', 'custom:我的模式')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(secondScope.writes.length, 16, '自定义模式同样写 16 项');

    byProp(reloaded, 'data-preset-remove', 'custom:我的模式')[0].props.onClick();
    reloaded = render(settingsReg(second.registrations));
    assert.equal(JSON.parse(store.getItem('dsh-office-mode.hyper-presets.v1')).length, 0, '删除后 localStorage 同步清空');
    assert.equal(byProp(reloaded, 'data-preset-card', 'custom:我的模式').length, 0);
});

await check('模式：内置模式名不能覆盖，空名字不保存', () => {
    const store = makeLocalStorage();
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, localStorage: store });
    mod.apply(ctx);
    const reg = settingsReg(registrations);

    let tree = render(reg);
    byProp(tree, 'data-preset-name')[0].props.onChange(eventOf({ value: '均衡' }));
    tree = render(reg);
    byProp(tree, 'data-preset-save')[0].props.onClick();
    tree = render(reg);
    assert.match(stringsIn(byProp(tree, 'data-preset-hint')[0]).join(''), /内置模式名不能覆盖/);
    assert.equal(store.getItem('dsh-office-mode.hyper-presets.v1'), null, '拒绝时不该落盘');
});

await check('模式：自定义模式最多 12 个，第 13 个明确拒绝且不覆盖已有', () => {
    const store = makeLocalStorage();
    const filled = Array.from({ length: 12 }, (_, index) => ({
        name: '模式' + String(index + 1).padStart(2, '0'),
        values: {},
    }));
    store.setItem('dsh-office-mode.hyper-presets.v1', JSON.stringify(filled));
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, localStorage: store });
    mod.apply(ctx);
    const reg = settingsReg(registrations);

    let tree = render(reg);
    assert.equal(byProp(tree, 'data-preset-card').length, 16, '四张内置 + 十二张自定义');
    byProp(tree, 'data-preset-name')[0].props.onChange(eventOf({ value: '第十三个' }));
    tree = render(reg);
    byProp(tree, 'data-preset-save')[0].props.onClick();
    tree = render(reg);

    const hint = byProp(tree, 'data-preset-hint')[0];
    assert.ok(hint, '应给出拒绝原因');
    assert.match(stringsIn(hint).join(''), /最多保存 12 个模式/);
    assert.equal(JSON.parse(store.getItem('dsh-office-mode.hyper-presets.v1')).length, 12, '存储仍是 12 个');
    assert.equal(byProp(tree, 'data-preset-card', 'custom:第十三个').length, 0, '第 13 个不出现');
});

await check('文案：关闭开关只讲直接后果，不写「其余功能不受影响」', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    for (const [group, tool] of [['ocr', '识别文件'], ['python', 'office.python'], ['av', 'office.av']]) {
        const row = rowInGroup(tree, group, '启用');
        assert.ok(row, group + ' 应有「启用」行');
        const note = String(row.source.props.note || '');
        assert.ok(note.includes('报错'), group + ' 要说明关闭后的直接后果：' + note);
        assert.ok(!note.includes('其余功能不受影响'), group + ' 不该写「其余功能不受影响」：' + note);
        assert.ok(!note.includes(tool + ' 不受影响'), group + ' 不该写「' + tool + ' 不受影响」');
    }
});

await check('总览：只有模式卡，不再画四枚指标', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);

    const overview = render(withTab(settingsReg(registrations), 'overview'));
    assert.equal(byProp(overview, 'data-settings-overview').length, 0, '旧的总览指标块已删除');
    assert.equal(byProp(overview, 'data-settings-fact').length, 0);
    const cards = byProp(overview, 'data-preset-card');
    assert.deepEqual(cards.map((card) => card.props['data-preset-card']),
        ['balanced', 'lite', 'local', 'deep'], '四张内置模式卡');
    const text = stringsIn(cards[0]).join(' ');
    assert.match(text, /均衡/, '第一张卡应显示模式名');

    const office = render(officeReg(registrations));
    assert.equal(byProp(office, 'data-settings-overview').length, 0, '页签内不再重复画总览');
});

await check('模式卡：状态钩子挂在 PresetHomeBody 组件上，不放进 SettingsPanel 的回调', () => {
    // 回归：第 105 轮把 useState 直接写在 SettingsPanel 的 render 回调里。
    // 第一次渲染（还在读设置）回调不执行，第二次才执行，钩子数量从 3 变 6，
    // 真实 React 抛「Rendered more hooks than during the previous render」，
    // 设置页整块空白 —— 假 React 不检查钩子数量，所以这里守住代码形状。
    const src = readFileSync(BUNDLE, 'utf8');
    const wrapperStart = src.indexOf('function PresetHome(props)');
    const bodyStart = src.indexOf('function PresetHomeBody(props)');
    assert.ok(wrapperStart > 0, '应有 PresetHome 包装组件');
    assert.ok(bodyStart > wrapperStart, '应有 PresetHomeBody 内容组件');
    const wrapper = src.slice(wrapperStart, bodyStart);
    assert.doesNotMatch(wrapper, /React\.use(State|Effect|Memo|Ref|Callback)/,
        'PresetHome 的 render 回调里不能再调用钩子，否则读设置完成后会整页空白');
    assert.match(wrapper, /h\(PresetHomeBody, state\)/, '包装组件要把取数结果交给 PresetHomeBody');
    const body = src.slice(bodyStart, src.indexOf('function MergedSettingsSection', bodyStart));
    assert.match(body, /React\.useState\(loadCustomPresets\)/, '自定义模式状态在 PresetHomeBody 里');
});

await check('分组折叠：默认展开；收起只是 display:none，控件仍在树里', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    const labels = rowLabels(tree);
    for (const id of TAB_GROUPS.docs) assert.ok(groupBody(tree, id), id + ' 应有组体');
    assert.ok(labels.length > 20, '文档页不该只有几行，实际 ' + labels.length + ' 行');
    assert.ok(labels.includes('默认主题'), '文档与缓存仍在');
    assert.ok(labels.includes('识别超时'), 'OCR 仍在');

    groupToggle(tree, 'ocr').props.onClick();
    const collapsed = renderOfficeTab(registrations, 'docs');
    assert.equal(groupOpen(collapsed, 'ocr'), 'false');
    assert.deepEqual(groupBody(collapsed, 'ocr').props.style, { display: 'none' });
    assert.ok(rowLabels(collapsed).includes('识别超时'), '收起后控件仍在树里，既有覆盖断言不必先展开');
});

await check('分组折叠：折叠头上的「N 项」等于该组体内的行数', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const tree = renderOfficeTab(registrations, 'docs');
    for (const id of TAB_GROUPS.docs) {
        const rows = walk(groupBody(tree, id)).filter((node) => node.source && node.source.name === 'Row').length;
        assert.equal(groupCount(tree, id), rows + ' 项', id + ' 的项数应等于组内行数');
    }
});

await check('分组折叠：默认展开，点标题收起/展开并写进 localStorage', () => {
    const store = makeLocalStorage();
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, localStorage: store });
    mod.apply(ctx);
    const docs = withTab(settingsReg(registrations), 'docs');

    const start = render(docs);
    assert.equal(groupOpen(start, 'ocr'), 'true', '默认展开');
    assert.equal(groupOpen(start, 'documents'), 'true');

    groupToggle(start, 'ocr').props.onClick();
    const closed = render(docs);
    assert.equal(groupOpen(closed, 'ocr'), 'false', '点标题收起');
    assert.deepEqual(groupBody(closed, 'ocr').props.style, { display: 'none' });
    assert.equal(groupOpen(closed, 'documents'), 'true', '只收被点的那一组');
    assert.equal(JSON.parse(store.getItem('dsh-office-mode.groups.office')).ocr, false, '开合状态写进 localStorage');

    groupToggle(closed, 'ocr').props.onClick();
    const opened = render(docs);
    assert.equal(groupOpen(opened, 'ocr'), 'true', '再点一次展开');
    assert.equal(JSON.parse(store.getItem('dsh-office-mode.groups.office')).ocr, true);
});

await check('分组折叠：localStorage 里有记录时按记录恢复；坏记录退回默认展开', () => {
    const store = makeLocalStorage();
    store.setItem('dsh-office-mode.groups.office', JSON.stringify({ ocr: false }));
    const scope = makeScope(SAMPLE_VALUE);
    const first = loadBundle({ forms: scope.forms, localStorage: store });
    first.mod.apply(first.ctx);
    const tree = renderOfficeTab(first.registrations, 'docs');
    assert.equal(groupOpen(tree, 'ocr'), 'false', '上次收起的那一组应恢复收起');
    assert.equal(groupOpen(tree, 'documents'), 'true', '没记录的组保持默认展开');

    // 坏记录（半截 JSON）：解析失败按默认展开处理，不抛错、不打不开页面。
    const broken = makeLocalStorage();
    broken.setItem('dsh-office-mode.groups.memory', '{oops');
    const second = loadBundle({ forms: makeScope(SAMPLE_VALUE).forms, localStorage: broken });
    second.mod.apply(second.ctx);
    const memoryTree = renderMemory(second.registrations);
    assert.equal(groupOpen(memoryTree, 'memory'), 'true', '坏记录退回默认展开');
});

await check('分组折叠：全部收起 / 展开作用于当前页，另一个页签不受影响', () => {
    const store = makeLocalStorage();
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms, localStorage: store });
    mod.apply(ctx);

    const docs = withTab(settingsReg(registrations), 'docs');
    const start = render(docs);
    walk(start).find((node) => node.props && node.props['data-groups-collapse'] === 'office').props.onClick();
    const closed = render(docs);
    for (const id of TAB_GROUPS.docs) assert.equal(groupOpen(closed, id), 'false', id + ' 应被全部收起');

    const search = renderOfficeTab(registrations, 'search');
    for (const id of TAB_GROUPS.search) assert.equal(groupOpen(search, id), 'true', '另一页签仍默认展开');

    walk(closed).find((node) => node.props && node.props['data-groups-expand'] === 'office').props.onClick();
    const opened = render(docs);
    for (const id of TAB_GROUPS.docs) assert.equal(groupOpen(opened, id), 'true', id + ' 应被全部展开');
});

await check('定位：只留下命中的行、命中组自动展开、报出命中数', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const search = withTab(settingsReg(registrations), 'search');

    const tree = locate(search, render(search), 'office', '出口代理');
    assert.deepEqual(groupIds(tree), ['search'], '只有命中的那一组留下');
    assert.equal(groupOpen(tree, 'search'), 'true', '命中组自动展开');
    const labels = rowLabels(tree);
    assert.deepEqual(labels, ['出口代理'], '只留下命中那一行');
    assert.ok(/匹配 1 项/.test(locatorText(tree, 'office')), '状态行应报匹配 1 项：' + locatorText(tree, 'office'));

    // 过滤后剩下的仍是**活的控件**：改一改照样写回。
    const row = findByLabel(tree, '出口代理');
    const proxyInput = controlOf(row);
    assert.equal(proxyInput.type, 'input');
    proxyInput.props.onBlur(eventOf({ value: 'http://127.0.0.1:7897' }));
    assert.deepEqual(scope.writes[0], ['search.proxy', 'http://127.0.0.1:7897']);
});

await check('定位：命中分组标题时整组保留（搜「音频与视频」）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const docs = withTab(settingsReg(registrations), 'docs');

    const tree = locate(docs, render(docs), 'office', '音频与视频');
    assert.deepEqual(groupIds(tree), ['av'], '只有音频与视频那一组留下');
    const labels = rowLabels(tree);
    assert.ok(labels.includes('ffmpeg 路径') && labels.includes('抽帧宽度'), '整组保留：首尾两行都在');
    const count = groupCount(tree, 'av');
    const parts = count.replace(' 项', '').split(' / ');
    assert.equal(parts[0], parts[1], '整组保留时命中数等于组内行数：' + count);
});

await check('定位：折在 Fold 里的高级参数命中时，折叠块自动展开', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const search = withTab(settingsReg(registrations), 'search');

    const tree = locate(search, render(search), 'office', 'search.providers.deepseek-api.timeoutMs');
    assert.deepEqual(rowLabels(tree), ['超时（毫秒）'], '按配置路径命中了那一行');
    const fold = walk(tree).find((node) => node.props && node.props['data-fold'] === 'search-advanced');
    assert.ok(fold, '高级参数折叠块仍在');
    assert.equal(walk(fold).find((node) => node.props && node.props['data-fold-toggle'] === 'search-advanced').props['aria-expanded'],
        true, '定位期间折叠块要展开，否则命中了也看不见');
    assert.equal(walk(fold).find((node) => node.props && node.props['data-fold-body'] === 'search-advanced').props.style,
        undefined, '折叠块体不再 display:none');
});

await check('定位：多词按「与」匹配、大小写不敏感（搜「LATEX thuthesis」）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const docs = withTab(settingsReg(registrations), 'docs');

    const tree = locate(docs, render(docs), 'office', 'LATEX thuthesis');
    assert.deepEqual(groupIds(tree), ['documents'], '只留文档与缓存那一组');
    assert.deepEqual(rowLabels(tree), ['LaTeX 编译入口'], '两个词都要出现：只有它同时含 latexmk 与 thuthesis');
});

await check('定位：没有命中时给可读提示，而不是一片空白', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const search = withTab(settingsReg(registrations), 'search');

    const tree = locate(search, render(search), 'office', 'zzz-没有这一项');
    assert.deepEqual(groupIds(tree), [], '没有命中的分组');
    assert.deepEqual(rowLabels(tree), [], '没有命中的行');
    assert.ok(/没有匹配项/.test(locatorText(tree, 'office')), '状态行应给出可读提示');
});

await check('定位：清空关键词后回到全量（分组回到各自的收起状态）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const docs = withTab(settingsReg(registrations), 'docs');

    const located = locate(docs, render(docs), 'office', 'latex');
    assert.ok(groupIds(located).length < TAB_GROUPS.docs.length, '定位期间只留部分分组');

    const cleared = locate(docs, located, 'office', '');
    assert.deepEqual(groupIds(cleared), TAB_GROUPS.docs, '清空后分组全回来');
    for (const id of TAB_GROUPS.docs) assert.equal(groupOpen(cleared, id), 'true', id + ' 回到默认展开');
    assert.match(locatorText(cleared, 'office'), /共 \d+ 项设置/, '清空后状态行回到总数（收起状态已逐个断言）');
});

await check('定位：记忆系统页同样过滤（搜「召回」只留召回质量一组）', () => {
    const scope = makeScope(SAMPLE_VALUE);
    const { mod, ctx, registrations } = loadBundle({ forms: scope.forms });
    mod.apply(ctx);
    const recall = withTab(settingsReg(registrations), 'recall');

    const tree = locate(recall, render(recall), 'memory', '召回');
    assert.deepEqual(groupIds(tree), ['recall'], '只留召回质量那一组');
    assert.ok(rowLabels(tree).includes('低分阈值'), '整组保留（标题命中）');
    assert.ok(/匹配 \d+ 项/.test(locatorText(tree, 'memory')));
});

// ── 新增：slot 注册面 ─────────────────────────────────────────────────────

// 第六十二轮（新-22）：上下文与编队在 live slot 树里的 order 是 20 / 30
// （dsh-context 注册，第六十一 / 六十二轮两次 Inspect 实读一致）。这里把它写成
// 一条**对外的锚**：只要它变了，本插件「在编队右侧」这个位置就不再成立，
// 这条断言应该红，而不是悄悄漂走。
const FRAMEWORK_TAB_ORDER = { context: 20, fleet: 30 };

await check('注册了记忆浏览面板：会话页签行里的一格，排在上下文与编队的右侧', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    assert.ok(panel, 'conversation.view 应有 office-memory 这一格');
    assert.equal(panel.opts.name, 'conversation.view');
    assert.equal(typeof panel.opts.order, 'number');
    assert.ok(panel.opts.order > FRAMEWORK_TAB_ORDER.context, '要排在「上下文」右侧');
    assert.ok(panel.opts.order > FRAMEWORK_TAB_ORDER.fleet, '要排在「编队」右侧（这是用户点名的那一句）');
    assert.equal(typeof panel.opts.label, 'function', 'label 要用 thunk（支持本地化重取）');
    assert.equal(panel.opts.label(), '记忆');
    assert.equal(typeof panel.render, 'function');
    // list slot：注册参数就是名字 / id / order / label 这一组，没有第二个名字要跟上。
    assert.deepEqual(Object.keys(panel.opts).sort(), ['id', 'label', 'name', 'order']);
    // 迁移干净：老的两处（侧栏一格 + 主面板一格）都不该再注册 —— 留着就是死代码，
    // 而且会让人以为还有第二条入口。
    assert.equal(regOf(registrations, 'sidebar.panellist', 'office-memory'), undefined, '不该再占侧栏一格');
    assert.equal(regOf(registrations, 'main', 'office-memory'), undefined, '不该再占主面板一格');
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

// ── 输入框左侧：知识库范围选择器 ─────────────────────────────────────────
//
// 这一格只负责「显示当前范围 + 把选择写回」，范围本身在服务端强制（见
// test/kb-scope.mjs）。这里要钉住的是三件事：宿主严格上下文下只注入
// remote.commands 才拿得到命令代理、会话值只经命令通道进出、命令不可用时
// 只降级不抛错。

/** 等一轮微任务，让假 fetch / 假命令的 promise 落地。 */
const tick = () => new Promise((resolve) => setTimeout(resolve, 10));

/** 造一个命令替身：记下每次 (sessionId, line)，按用例给的规则回执。 */
function fakeCommands(reply) {
    const calls = [];
    return {
        calls,
        execute: async (sessionId, line) => { calls.push({ sessionId, line }); return reply(line); },
    };
}

/** 点某个带 data-* 标记的按钮（选择器里的按钮都靠这个找）。 */
function clickMarked(tree, key, value) {
    const node = walk(tree).find((item) => item.props && item.props[key] !== undefined
        && (value === undefined || item.props[key] === value));
    assert.ok(node, `应找到按钮 ${key}${value === undefined ? '' : '=' + value}`);
    node.props.onClick();
}

await check('输入框左侧：注册一格，没有会话时不渲染；没有命令通道时只提示不抛错', async () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    assert.ok(picker, '应有知识库范围选择器');
    assert.equal(picker.opts.name, 'conversation.input.left');
    assert.equal(typeof picker.opts.order, 'number');
    assert.equal(render(picker, {}), null, '没有会话时不渲染（这一格是 session 作用域）');
    const tree = render(picker, { sessionId: 's-1' });
    assert.ok(walk(tree).some((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined), '有会话时应画出入口');
    // 部署里没有命令通道：点开面板后只能给出提示，不能把整页弄挂。
    assert.doesNotThrow(() => clickMarked(tree, 'data-kb-scope-toggle'));
    await tick();
    const texts = textsOf(render(picker, { sessionId: 's-1' }));
    assert.ok(texts.some((text) => text.includes('命令通道还没就绪')), `应提示命令通道不可用：${texts.join('|')}`);
});

await check('输入框左侧：打开时读状态与目录，选目录 / 全部各写一条命令', async () => {
    const urls = [];
    let current = '当前会话的知识库范围：全部知识库（不限目录）。';
    const commands = fakeCommands((line) => {
        if (line === '/kb-scope') {
            return { ok: true, value: { result: { kind: 'success', text: current } } };
        }
        if (line === '/kb-scope clear') {
            current = '当前会话的知识库范围已清空：可以检索全部知识库。';
            return { ok: true, value: { result: { kind: 'success', text: current } } };
        }
        current = '当前会话的知识库范围已设为「' + line.slice('/kb-scope set '.length) + '」及其子目录。';
        return { ok: true, value: { result: { kind: 'success', text: current } } };
    });
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ remote: { commands }, fetch: fetchImpl });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = {
        sessionId: 's-1',
        useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/work/math' } }),
    };
    render(picker, props);
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    let tree = render(picker, props);
    // 状态走命令通道，选项走已有的目录读接口 —— 两边都不带会话身份。
    assert.deepEqual(commands.calls.map((call) => call.line), ['/kb-scope']);
    assert.equal(urls.length, 1, '打开一次只取一次目录');
    assert.ok(urls[0].includes('cwd=' + encodeURIComponent('D:/work/math')), `目录请求应带当前工作区：${urls[0]}`);
    assert.equal(urls[0].includes('s-1'), false, '目录请求不该暴露会话 ID');
    const folders = walk(tree).filter((node) => node.props && node.props['data-kb-scope-folder'] !== undefined)
        .map((node) => node.props['data-kb-scope-folder']);
    assert.deepEqual(folders, ['来源', '来源/方法论'], '应列出知识库里的虚拟文件夹');
    assert.ok(textsOf(tree).some((text) => text.includes('全部知识库')), '状态行应显示当前范围');

    clickMarked(tree, 'data-kb-scope-folder', '来源/方法论');
    await tick();
    tree = render(picker, props);
    assert.deepEqual(commands.calls.map((call) => call.line), ['/kb-scope', '/kb-scope set 来源/方法论', '/kb-scope']);
    assert.ok(commands.calls.every((call) => call.sessionId === 's-1'), '命令必须带当前会话 ID（范围由服务端按会话判定）');
    assert.ok(textsOf(tree).some((text) => text.includes('来源/方法论')), '选完后入口显示该文件夹');
    assert.equal(walk(tree).some((node) => node.props && node.props['data-kb-scope-panel'] !== undefined), false, '选完应收起面板');

    clickMarked(tree, 'data-kb-scope-toggle');
    await tick();
    tree = render(picker, props);
    assert.equal(commands.calls.length, 3, '已就绪时再打开不重复读状态');
    assert.equal(urls.length, 3, '每次打开都要重新取目录（不能只在列表为空时取）；选中目录前也会再核一次');
    clickMarked(tree, 'data-kb-scope-all');
    await tick();
    tree = render(picker, props);
    assert.deepEqual(commands.calls.map((call) => call.line), ['/kb-scope', '/kb-scope set 来源/方法论', '/kb-scope', '/kb-scope clear', '/kb-scope']);
    assert.ok(textsOf(tree).some((text) => text.includes('全部知识库')), '恢复不限制');
});

await check('输入框左侧：每次打开刷新目录，关闭期间新建的目录再次打开即可选', async () => {
    let kbTree = JSON.parse(JSON.stringify(makeSnapshot().kbTree));
    // 只读的「归档」也必须出现在目录树里：它同样是可选的虚拟文件夹。
    kbTree.children.push({ path: '归档', name: '归档', docs: 0, chunks: 0, bytes: 0, children: [] });
    let fetchCount = 0;
    const commands = fakeCommands(() => ({ ok: true, value: { result: { kind: 'success', text: '知识库范围：全部知识库。' } } }));
    const fetchImpl = () => {
        fetchCount += 1;
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot({ kbTree })) });
    };
    const { mod, ctx, registrations } = loadBundle({ remote: { commands }, fetch: fetchImpl });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/work/math' } }) };
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    let tree = render(picker, props);
    assert.equal(walk(tree).filter((node) => node.props && node.props['data-kb-scope-folder'] !== undefined).length, 3, '首次打开应列出三个目录（含只读归档）');
    assert.ok(walk(tree).some((node) => node.props && node.props['data-kb-scope-folder'] === '归档'), '只读归档也要列出来');

    // 关闭选择器；期间知识库新增一个目录（模拟另一个 session / panel 的写入）。
    clickMarked(tree, 'data-kb-scope-toggle');
    await tick();
    const nextTree = JSON.parse(JSON.stringify(kbTree));
    nextTree.children.push({ path: '新建目录', name: '新建目录', docs: 0, chunks: 0, bytes: 0, children: [] });
    kbTree = nextTree;
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    tree = render(picker, props);
    const paths = walk(tree).filter((node) => node.props && node.props['data-kb-scope-folder'] !== undefined)
        .map((node) => node.props['data-kb-scope-folder']);
    assert.ok(paths.includes('新建目录'), '再次打开必须刷新目录，新建目录要可选：' + paths.join(','));
    assert.equal(fetchCount, 2, '每次打开都重新取一次目录');
});

await check('输入框左侧：目录读失败显示原因并可重试，恢复后可选', async () => {
    let kbTree = makeSnapshot().kbTree;
    let failing = true;
    const commands = fakeCommands(() => ({ ok: true, value: { result: { kind: 'success', text: '知识库范围：全部知识库。' } } }));
    const fetchImpl = () => {
        if (failing) {
            return Promise.resolve({ ok: false, status: 403, json: () => Promise.resolve({ ok: false, error: '这个目录不在受信任列表里' }) });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot({ kbTree })) });
    };
    const { mod, ctx, registrations } = loadBundle({ remote: { commands }, fetch: fetchImpl });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/work/math' } }) };
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    let tree = render(picker, props);
    const texts = textsOf(tree);
    assert.ok(texts.some((text) => text.includes('读取目录失败') && text.includes('403') && text.includes('受信任')), '403 必须显示原因：' + texts.join('|'));
    assert.ok(walk(tree).some((node) => node.props && node.props['data-kb-scope-folders-retry'] !== undefined), '应提供重试按钮');
    assert.equal(walk(tree).some((node) => node.props && node.props['data-kb-scope-folders-empty'] !== undefined), false, '读失败不能被当成空目录');
    for (const node of walk(tree).filter((item) => item.props && item.props['data-kb-scope-folder'] !== undefined)) {
        assert.equal(node.props.disabled, true, '读失败时旧目录行必须禁用，不能拿旧列表提交');
    }

    // 恢复后重试：新目录出现、错误消失。
    failing = false;
    const recovered = JSON.parse(JSON.stringify(kbTree));
    recovered.children.push({ path: '恢复目录', name: '恢复目录', docs: 0, chunks: 0, bytes: 0, children: [] });
    kbTree = recovered;
    clickMarked(tree, 'data-kb-scope-folders-retry');
    await tick();
    tree = render(picker, props);
    const paths = walk(tree).filter((node) => node.props && node.props['data-kb-scope-folder'] !== undefined)
        .map((node) => node.props['data-kb-scope-folder']);
    assert.ok(paths.includes('恢复目录'), '重试后应列出新目录：' + paths.join(','));
    assert.equal(walk(tree).some((node) => node.props && node.props['data-kb-scope-folders-error'] !== undefined), false, '重试成功后错误应消失');
});

await check('输入框左侧：列表读取中显示加载标记；刷新后目录已失效的行不提交旧路径', async () => {
    let kbTree = makeSnapshot().kbTree;
    let release = null;
    let defer = true;
    const commands = fakeCommands((line) => ({ ok: true, value: { result: { kind: 'success', text: line === '/kb-scope' ? '知识库范围：全部知识库。' : '知识库范围：已设为「来源」（含子目录）。' } } }));
    const fetchImpl = () => {
        if (!defer) return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot({ kbTree })) });
        return new Promise((resolve) => { release = resolve; });
    };
    const { mod, ctx, registrations } = loadBundle({ remote: { commands }, fetch: fetchImpl });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/work/math' } }) };
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    let tree = render(picker, props);
    assert.ok(walk(tree).some((node) => node.props && node.props['data-kb-scope-folders-loading'] !== undefined), '读目录期间应有加载标记');

    release({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot({ kbTree })) });
    await tick();
    tree = render(picker, props);
    const staleTree = tree;
    // 目录在刷新之间被删掉：点击旧行时必须先核一遍，不能直接 set 旧路径。
    kbTree = JSON.parse(JSON.stringify(kbTree));
    kbTree.children = kbTree.children.filter((child) => child.path !== '来源');
    defer = false;
    clickMarked(staleTree, 'data-kb-scope-folder', '来源');
    await tick();
    tree = render(picker, props);
    assert.equal(commands.calls.some((call) => call.line === '/kb-scope set 来源'), false, '失效目录不能被提交');
    assert.ok(textsOf(tree).some((text) => text.includes('刚刚不可用')), '应提示目录已失效并让用户重选');
});

await check('输入框左侧：命令回执是错误时留在面板里给出原因', async () => {
    const commands = fakeCommands((line) => {
        if (line === '/kb-scope') return { ok: true, value: { result: { kind: 'success', text: '当前会话的知识库范围：全部知识库。' } } };
        return { ok: true, value: { result: { kind: 'error', text: '知识库里没有「不存在」这个虚拟文件夹。' } } };
    });
    const { mod, ctx, registrations } = loadBundle({
        remote: { commands },
        fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) }),
    });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/work/math' } }) };
    render(picker, props);
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    const tree = render(picker, props);
    // 命令失败要能在面板里看到原因，且面板不能收起（收起会看不出失败）。
    const folder = walk(tree).find((node) => node.props && node.props['data-kb-scope-folder'] === '来源');
    assert.ok(folder, '应列出目录按钮');
    // 这一条用回执错误模拟「目录刚被删/名字被改」：服务端拒绝时 UI 只显示。
    clickMarked(tree, 'data-kb-scope-folder', '来源');
    await tick();
    const after = render(picker, props);
    assert.ok(textsOf(after).some((text) => text.includes('知识库里没有')), '应显示命令返回的错误原因');
    assert.ok(walk(after).some((node) => node.props && node.props['data-kb-scope-panel'] !== undefined), '失败时面板保持打开');
});

await check('输入框左侧：注入了 remote 但缺 commands 时按「没就绪」降级', async () => {
    const { mod, ctx, registrations } = loadBundle({ remote: {} });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = { sessionId: 's-1' };
    render(picker, props);
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    const texts = textsOf(render(picker, props));
    assert.ok(texts.some((text) => text.includes('命令通道还没就绪')), '缺少命令服务时只提示');
});

/** 造一个只回 /kb-scope 的假命令通道，并按给定状态文本打开选择器。 */
async function openScopePicker(statusText, snapshotOverrides = {}) {
    // 服务端是有状态的：set/clear 之后 status 必须回读到新值，客户端才允许标 ready。
    let current = statusText;
    const commands = fakeCommands((line) => {
        if (line === '/kb-scope') {
            return { ok: true, value: { result: { kind: 'success', text: current } } };
        }
        if (line === '/kb-scope clear') {
            current = '知识库范围：已清空（全部知识库）。本会话不再限制文件夹。';
            return { ok: true, value: { result: { kind: 'success', text: current } } };
        }
        current = '知识库范围：已设为「' + line.slice('/kb-scope set '.length) + '」（含子目录）。';
        return { ok: true, value: { result: { kind: 'success', text: current } } };
    });
    const { mod, ctx, registrations } = loadBundle({
        remote: { commands },
        fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot(snapshotOverrides)) }),
    });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/work/math' } }) };
    render(picker, props);
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    return { picker, props, commands };
}

await check('输入框左侧：未设置 / 损坏 / 空交集 / 父链异常都明确显示禁止检索，并给出显式选择全部的说明', async () => {
    const cases = [
        ['unset', '📚 这个会话还没有设置知识库范围，所以不能查知识库（不会默认放开全库）。在输入框左侧选一个虚拟文件夹，或输入 /kb-scope set <虚拟文件夹>；要查全部用 /kb-scope clear。', '未选择 · 禁止检索'],
        ['corrupt', '知识库范围：状态损坏。为安全起见这次拒绝查询，不会退回全库。', '范围损坏 · 禁止检索'],
        ['empty', '知识库范围：范围为空。继承到的范围与它自己的记录没有交集。', '范围为空 · 禁止检索'],
        ['ancestry', '知识库范围：父会话不可用。无法确认继承范围，拒绝查询。', '父会话不可用 · 禁止检索'],
        ['cycle', '知识库范围：父链循环。拒绝查询。', '父链循环 · 禁止检索'],
        ['depth', '知识库范围：父链过深。拒绝查询。', '父链过深 · 禁止检索'],
    ];
    for (const [state, text, label] of cases) {
        const { picker, props } = await openScopePicker(text);
        const tree = render(picker, props);
        const toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
        assert.ok(toggle, state + '：应有入口');
        assert.equal(toggle.props['data-kb-scope-blocked'], 'true', state + '：入口要标成禁止检索');
        assert.equal(toggle.props['data-kb-scope-state'], state, state + '：状态 token 要稳定');
        const allText = textsOf(tree).join(' ');
        assert.ok(allText.includes(label), state + '：入口应显示「' + label + '」，实际：' + allText);
        assert.equal(allText.includes('📚 全部知识库（不限制）'), false, state + '：不能显示成已放开全库');
        const notice = walk(tree).find((node) => node.props && node.props['data-kb-scope-notice'] !== undefined);
        assert.ok(notice, state + '：面板应有禁止检索说明');
        const noticeText = textsOf(notice).join(' ');
        assert.ok(noticeText.includes('禁止检索') || noticeText.includes('拒绝检索'), state + '：说明要讲清当前禁止检索');
        if (state === 'unset') {
            assert.ok(noticeText.includes('显式选择'), '未设置：必须说明要显式选择全部才能放开全库');
            assert.ok(noticeText.includes('全部知识库（不限制）'), '未设置：要给出显式的全部选项文案');
        }
        if (state === 'corrupt') assert.ok(noticeText.includes('不会退回全库'), '损坏：要明确不 fail-open');
        if (state === 'ancestry') assert.ok(noticeText.includes('fail-closed'), '父会话不可用：要明确 fail-closed');
    }
});

await check('输入框左侧：恰好名为「全部知识库」的文件夹按文件夹显示，不当成放开全库', async () => {
    const treeBody = {
        path: '', name: '全部文档', docs: 1, chunks: 3, bytes: 4096,
        children: [{ path: '全部知识库', name: '全部知识库', docs: 1, chunks: 3, bytes: 4096, children: [] }],
    };
    const { picker, props, commands } = await openScopePicker('知识库范围：「全部知识库」（含子目录）。', { kbTree: treeBody });
    let tree = render(picker, props);
    let toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
    assert.equal(toggle.props['data-kb-scope-state'], 'folder', '同名文件夹必须解析成 folder');
    assert.ok(textsOf(tree).join(' ').includes('文件夹：全部知识库'), '入口应显示为文件夹');
    const allRow = walk(tree).find((node) => node.props && node.props['data-kb-scope-all'] !== undefined);
    const folderRow = walk(tree).find((node) => node.props && node.props['data-kb-scope-folder'] === '全部知识库');
    assert.ok(allRow && folderRow);
    assert.equal(textsOf(allRow).join(''), '全部知识库（不限制）', '全部行必须带（不限制）');
    assert.equal(textsOf(folderRow).join(''), '全部知识库（文件夹）', '同名文件夹行必须带（文件夹）');

    clickMarked(tree, 'data-kb-scope-folder', '全部知识库');
    await tick();
    tree = render(picker, props);
    assert.ok(commands.calls.some((call) => call.line === '/kb-scope set 全部知识库'), '选择同名文件夹要写文件夹范围');
    toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
    assert.equal(toggle.props['data-kb-scope-state'], 'folder');
    assert.ok(textsOf(tree).join(' ').includes('文件夹：全部知识库'));

    clickMarked(tree, 'data-kb-scope-toggle');
    await tick();
    tree = render(picker, props);
    clickMarked(tree, 'data-kb-scope-all');
    await tick();
    tree = render(picker, props);
    assert.ok(commands.calls.some((call) => call.line === '/kb-scope clear'), '选全部仍要清空范围');
    toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
    assert.equal(toggle.props['data-kb-scope-state'], 'all');
    assert.ok(textsOf(tree).join(' ').includes('全部知识库（不限制）'));
});

await check('输入框左侧：新状态前缀与旧文案都能识别；未知回执不冒充全部', async () => {
    const cases = [
        ['知识库范围：全部知识库。', 'all'],
        ['知识库范围：「资料」（含子目录）。', 'folder'],
        ['当前会话的知识库范围：全部知识库（不限目录）。', 'all'],
        ['当前会话的知识库范围已清空：可以检索全部知识库。', 'all'],
        ['当前会话范围信息格式变了。', 'unknown'],
    ];
    for (const [text, expected] of cases) {
        const { picker, props } = await openScopePicker(text);
        const tree = render(picker, props);
        const toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
        assert.equal(toggle.props['data-kb-scope-state'], expected, text + ' → ' + expected);
        if (expected === 'unknown') {
            assert.equal(toggle.props['data-kb-scope-blocked'], 'true', '未知回执必须按禁止检索处理');
            assert.ok(textsOf(tree).join(' ').includes('范围不可确认'), '未知回执要显示不可确认');
        }
    }

    const { picker, props, commands } = await openScopePicker('知识库范围：全部知识库。');
    let tree = render(picker, props);
    const folderRow = walk(tree).find((node) => node.props && node.props['data-kb-scope-folder'] === '来源');
    assert.ok(folderRow, '目录列表仍应可用');
    clickMarked(tree, 'data-kb-scope-folder', '来源');
    await tick();
    tree = render(picker, props);
    assert.ok(commands.calls.some((call) => call.line === '/kb-scope set 来源'), '选择应写回服务端');
    const toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
    assert.equal(toggle.props['data-kb-scope-state'], 'folder', '回读 status 后要解析成 folder');
    assert.ok(textsOf(tree).join(' ').includes('文件夹：来源'));
});

await check('输入框左侧：写入回执是拒绝态时不用本地选择覆盖，且面板保持打开', async () => {
    const commands = fakeCommands((line) => {
        if (line === '/kb-scope') {
            return { ok: true, value: { result: { kind: 'success', text: '知识库范围：父会话不可用。无法确认继承范围，拒绝查询。' } } };
        }
        return { ok: true, value: { result: { kind: 'error', text: '知识库范围：父会话不可用。这次拒绝查询（fail-closed）。' } } };
    });
    const { mod, ctx, registrations } = loadBundle({
        remote: { commands },
        fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) }),
    });
    mod.apply(ctx);
    const picker = regOf(registrations, 'conversation.input.left', 'office-kb-scope');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/work/math' } }) };
    render(picker, props);
    clickMarked(render(picker, props), 'data-kb-scope-toggle');
    await tick();
    let tree = render(picker, props);
    clickMarked(tree, 'data-kb-scope-folder', '来源');
    await tick();
    tree = render(picker, props);
    const toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
    assert.equal(toggle.props['data-kb-scope-state'], 'ancestry', '拒绝态不能被本地选择覆盖成 folder');
    assert.equal(toggle.props['data-kb-scope-blocked'], 'true');
    assert.ok(textsOf(tree).join(' ').includes('父会话不可用'), '入口显示拒绝原因');
    assert.ok(walk(tree).some((node) => node.props && node.props['data-kb-scope-panel'] !== undefined), '拒绝态下不收起面板');
    const notice = walk(tree).find((node) => node.props && node.props['data-kb-scope-notice'] !== undefined);
    assert.ok(notice && textsOf(notice).join(' ').includes('fail-closed'), '面板里说明 fail-closed');
});

await check('输入框左侧：只认整行正式回执；拒绝/未知里的“全部知识库（不限制）”不冒充放开，同名与嵌套括号目录保真', async () => {
    const cases = [
        ['知识库范围：未设置。要查全部请显式选择 全部知识库（不限制）。', 'unset'],
        ['知识库范围：来源不可确认。这个会话没有可核对的父会话记录。', 'origin'],
        ['知识库范围：无会话。读不到范围。', 'no-session'],
        ['知识库范围：状态读取失败。拒绝查询。', 'state'],
        ['设置失败：已清空。', 'unknown'],
        ['说明：全部知识库（不限制）', 'unknown'],
    ];
    for (const [text, expected] of cases) {
        const { picker, props } = await openScopePicker(text);
        const tree = render(picker, props);
        const toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
        assert.equal(toggle.props['data-kb-scope-state'], expected, text + ' → ' + expected);
        assert.equal(toggle.props['data-kb-scope-blocked'], 'true', text + '：必须按禁止检索处理');
        assert.equal(textsOf(tree).join(' ').includes('📚 全部知识库（不限制）'), false, text + '：不能显示成已放开全库');
        if (expected === 'origin') {
            assert.ok(textsOf(tree).join(' ').includes('来源不可确认'), '来源不可确认要有明确 label');
            const notice = walk(tree).find((node) => node.props && node.props['data-kb-scope-notice'] !== undefined);
            assert.ok(notice && textsOf(notice).join(' ').includes('不可用'), '来源不可确认要说明入口不可用');
        }
    }

    // 目录恰好叫「全部知识库（不限制）」：仍是 folder，不是放开全库。
    const allLikeTree = {
        path: '', name: '全部文档', docs: 1, chunks: 1, bytes: 1,
        children: [{ path: '全部知识库（不限制）', name: '全部知识库（不限制）', docs: 1, chunks: 1, bytes: 1, children: [] }],
    };
    const like = await openScopePicker('知识库范围：「全部知识库（不限制）」（含子目录）。', { kbTree: allLikeTree });
    let tree = render(like.picker, like.props);
    let toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
    assert.equal(toggle.props['data-kb-scope-state'], 'folder', '同名（不限制）目录必须解析成 folder');
    assert.ok(textsOf(tree).join(' ').includes('文件夹：全部知识库（不限制）'));

    // 目录名里带「」：贪婪捕获到最后一个「」（含子目录）为止。
    const nested = await openScopePicker('知识库范围：「A「B」C」（含子目录）。');
    tree = render(nested.picker, nested.props);
    toggle = walk(tree).find((node) => node.props && node.props['data-kb-scope-toggle'] !== undefined);
    assert.equal(toggle.props['data-kb-scope-state'], 'folder', '嵌套括号目录必须解析成 folder');
    assert.ok(textsOf(tree).join(' ').includes('文件夹：A「B」C'), '嵌套括号要保真');
});

await check('五种 slot 都注册了，且 inject 与 register 一一对应', () => {
    const { mod, ctx, registrations, injected } = loadBundle();
    mod.apply(ctx);
    const slots = [...new Set(registrations.map((item) => item.opts.name))].sort();
    assert.deepEqual(slots, [
        'conversation.chat.assistant-actions',
        'conversation.chat.turnTail',
        'conversation.input.left',
        'conversation.view',
        'settings.section',
    ]);
    assert.equal(injected.length, registrations.length, '每个 inject 应对应一次 register');
});

await check('记忆图标仍画在面板头部（size 用上、currentColor 描边）', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    const tree = render(panel);
    const head = walk(tree).find((node) => node.props && node.props['data-memory-head'] === 'top');
    assert.ok(head, '面板头部应存在');
    const icon = walk(head).find((node) => node.type === 'svg');
    assert.ok(icon, '头部应有记忆图标');
    assert.equal(icon.props.width, 16, '面板头部给的尺寸要用上');
    assert.equal(icon.props.height, 16);
    assert.equal(icon.props.stroke, 'currentColor');
    assert.equal(icon.props['aria-hidden'], 'true');
    // 宿主在异常路径下可能不传 owner props：不能因此抛错（旧侧栏入口那条要求，
    // 换成页签之后同样成立 —— 列表 / 面板都可能被不带 props 地渲染一次）。
    assert.doesNotThrow(() => render(panel));
    assert.doesNotThrow(() => render(panel, {}));
});

// ── 新增：记忆浏览面板 ────────────────────────────────────────────────────

/** 取渲染树里的全部文本节点。 */
function textsOf(tree) {
    return walk(tree).map((node) => node.props && node.props.children).flat().filter((c) => typeof c === 'string');
}

/** 渲染面板并等一轮微任务，让假 fetch 的 promise 落地。 */
async function renderPanelLoaded(panel, props) {
    render(panel, props);
    await new Promise((resolve) => setTimeout(resolve, 10));
    return render(panel, props);
}

await check('记忆面板：挂载后取数，五个 Tab 显示 counts 里的条数', async () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    const tree = await renderPanelLoaded(panel);
    const texts = textsOf(tree);

    assert.ok(texts.some((t) => t.includes('记忆 · temp1')), '应显示工作目录短名');
    for (const label of ['记忆', '知识库']) {
        assert.ok(texts.some((t) => t === label), `Tab 只写去处，不再混排三种计数单位：${label}`);
    }
    // 记忆页的二级分档：个人偏好 / 项目记忆 / 台账（不再有来源筛选）。
    for (const label of ['个人偏好（1）', '项目记忆（0）', '台账（1）']) {
        assert.ok(texts.some((t) => t === label), `二级层级应显示条数：${label}`);
    }
    // 默认 Tab 是热记忆：显示内容 + 来源徽标 + 重要度
    assert.ok(texts.some((t) => t.includes('偏好：书面化表达')), '应显示热记忆内容');
    assert.ok(texts.some((t) => t === '全局'), '应显示来源徽标（全局 / 工作区）');
    assert.ok(texts.some((t) => t === 'critical'), '应显示重要度');
});

/**
 * 造一个假的 `useSessions`（宿主给会话作用域视图的标准 hook）。
 *
 * 只支持面板用到的那一件事：按 selector 读 `snapshot.byId[...]`。真实 hook 是
 * 订阅式的（值变了会重渲染），这里不需要 —— 面板只在挂载时读一次当前工作目录。
 */
function fakeUseSessions(byId) {
    return (selector) => selector({ byId });
}

await check('默认根：拿得到会话上下文时，第一次取数就带当前会话的工作目录', async () => {
    const urls = [];
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    const sessionCwd = 'D:/work/math-course';
    await renderPanelLoaded(panel, {
        sessionId: 's-1',
        useSessions: fakeUseSessions({ 's-1': { cwd: sessionCwd } }),
    });
    assert.ok(urls.length >= 1, '应当发过一次取数');
    // **第一次**就得带上：否则会先按宿主挑的默认根取一份（用户看到的是别的工作区），
    // 再被换成当前工作区 —— 那正是用户报的那个现象。
    assert.ok(
        urls[0].includes('cwd=' + encodeURIComponent(sessionCwd)),
        `第一次请求应带当前会话的工作目录：${urls[0]}`,
    );
});

await check('默认根：拿不到会话上下文（老宿主 / 这一格被别的槽位复用）时不带 cwd', async () => {
    const urls = [];
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    await renderPanelLoaded(panel);
    assert.ok(urls.length >= 1, '应当发过一次取数');
    assert.equal(urls[0].includes('cwd='), false, `没有会话上下文时不该带 cwd：${urls[0]}`);
});

await check('默认根：会话快照里没有这条 cwd（空会话 / 读不到）时同样退回老行为', async () => {
    const urls = [];
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    await renderPanelLoaded(panel, {
        sessionId: 's-missing',
        useSessions: fakeUseSessions({ 's-other': { cwd: 'D:/work/other' } }),
    });
    assert.ok(urls.length >= 1, '应当发过一次取数');
    assert.equal(urls[0].includes('cwd='), false, `取不到当前工作区时不该瞎猜：${urls[0]}`);
});

await check('记忆面板：切换 Tab 显示对应层的条目', async () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);

    // 台账 Tab
    const ledgerTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && Array.isArray(n.props.children) && n.props.children.includes('台账（1）'));
    assert.ok(ledgerTab, '应找到台账 Tab');
    ledgerTab.props.onClick();
    tree = render(panel);
    let texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('季度汇报.pptx')), '台账 Tab 应显示产物路径');

    // 归档不再单列（第八十四轮）：它作为知识库的「归档」文件夹出现，面板里不该有第二个入口。
    assert.equal(
        walk(tree).some((n) => n.props && n.props.role === 'tab' && Array.isArray(n.props.children) && n.props.children.includes('归档（1）')),
        false,
        '归档 Tab 应已取消',
    );

    // 知识库 Tab（第四十二轮）
    const kbTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb');
    assert.ok(kbTab, '应找到知识库 Tab');
    kbTab.props.onClick();
    tree = render(panel);
    texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('来源/方法论.md')), '知识库 Tab 应显示入库路径');
    assert.ok(texts.some((t) => t.includes('方法论笔记')), '知识库 Tab 应显示标题');
    assert.ok(texts.some((t) => t === 'verified'), '知识库 Tab 应显示来源档');
    assert.ok(texts.some((t) => t === '3 块'), '知识库 Tab 应显示块数');
    assert.ok(texts.some((t) => t === '4 KB'), '知识库 Tab 应显示人读化的体积');

    // 第九十一轮：旧「图谱」Tab（关系 + 实体）已按用户要求下线，凭证只留知识库块关系。
    assert.equal(
        walk(tree).some((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'links'),
        false,
        '图谱 Tab 应已取消',
    );
});

await check('记忆面板：个人偏好 / 项目记忆 / 台账三档按 target 分流，不再有来源筛选（第八十九轮）', async () => {
    const urls = [];
    const plain = makeSnapshot();
    const mixed = makeSnapshot({
        counts: Object.assign({}, plain.counts, { hot: 2, hotUser: 1, hotProject: 1 }),
        hot: [
            plain.hot[0],
            {
                id: 'm-2', target: 'project', importance: 'normal', origin: 'workspace',
                updatedAt: '2026-09-24T00:00:00.000Z', entities: [], tags: [], content: '约定：只提交一版公开历史',
            },
        ],
    });
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(mixed) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));

    assert.equal(byProp(tree, 'data-memory-origin', 'true').length, 0, '来源筛选应已删除');
    const layers = byProp(tree, 'data-memory-layers', 'true')[0];
    assert.ok(layers, '记忆页要有二级分档');
    assert.deepEqual(
        walk(layers).filter((n) => n.props && n.props['data-layer']).map((n) => n.props['data-layer']),
        ['preference', 'project', 'ledger'],
    );

    byProp(tree, 'data-layer', 'preference')[0].props.onClick();
    tree = render(panel);
    assert.ok(textsOf(tree).some((t) => t.includes('偏好：书面化表达')), '个人偏好档显示 user 条目');
    assert.ok(!textsOf(tree).some((t) => t.includes('只提交一版')), '个人偏好档不显示项目条目');

    byProp(tree, 'data-layer', 'project')[0].props.onClick();
    tree = render(panel);
    assert.ok(textsOf(tree).some((t) => t.includes('只提交一版')), '项目记忆档显示 project 条目');
    assert.ok(!textsOf(tree).some((t) => t.includes('偏好：书面化表达')), '项目记忆档不显示个人偏好');

    byProp(tree, 'data-layer', 'ledger')[0].props.onClick();
    tree = render(panel);
    assert.ok(textsOf(tree).some((t) => t.includes('季度汇报')), '台账档显示交付物');
    assert.ok(urls.every((u) => !u.includes('origin=')), '三档切换不再请求来源筛选');
});

await check('记忆面板：知识库块检索命中后可读块正文（第八十二轮）', async () => {
    const urls = [];
    const hit = {
        chunkId: 'kb:abc:1', docId: 'kb:abc', path: '来源/方法论.md', title: '方法论笔记',
        tier: 'verified', origin: 'workspace', n: 1, score: 1.25, chars: 18,
        spanText: '字符 1–18', header: '# 方法论', preview: '复核要能复跑。',
    };
    const doc = {
        docId: 'kb:abc', path: '来源/方法论.md', title: '方法论笔记', tier: 'verified',
        origin: 'workspace', chunks_total: 2, chunks_missing: 0, selectedChunk: 0,
        chunks: [
            { id: 'kb:abc:1', n: 1, chars: 18, spanText: '字符 1–18', header: '# 方法论', outline: ['方法论'] },
            { id: 'kb:abc:2', n: 2, chars: 20, spanText: '字符 19–38', header: '## 第二段', outline: ['方法论', '第二段'] },
        ],
    };
    const chunk = {
        id: 'kb:abc:1', n: 1, text: '复核要能复跑。', spanText: '字符 1–18',
        header: '# 方法论', chars: 7, docId: 'kb:abc', path: '来源/方法论.md',
        title: '方法论笔记', tier: 'verified', origin: 'workspace',
    };
    const base = makeSnapshot();
    const searched = makeSnapshot({ kbSearch: { hits: [hit], matched: 1, truncated: false, reason: '' } });
    const drilled = makeSnapshot({ kbSearch: { hits: [hit], matched: 1, truncated: false, reason: '' }, kbDoc: doc, kbChunk: chunk });
    const fetchImpl = (url) => {
        const text = String(url);
        urls.push(text);
        const body = text.includes('kbchunk=') ? drilled
            : text.includes('kbdoc=') ? drilled
                : text.includes('kbq=') ? searched
                    : base;
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));

    const kbTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb');
    kbTab.props.onClick();
    tree = render(panel);
    assert.ok(byProp(tree, 'data-kb-bar', 'true').length === 1, '知识库页签要有文档 / 检索工具条');

    byProp(tree, 'data-kb-mode', 'search')[0].props.onClick();
    tree = await renderPanelLoaded(panel);
    let input = byProp(tree, 'data-kb-search-input', 'true')[0];
    assert.ok(input, '块检索模式要有检索框');
    // 受控输入：打字后 value 跟着 state 走，点「清空检索」才能真的把框清掉。
    input.props.onChange(eventOf({ value: '复核' }));
    tree = render(panel);
    input = byProp(tree, 'data-kb-search-input', 'true')[0];
    assert.equal(input.props.value, '复核', '输入框要受控，值来自 state.kbQuery');
    input.props.onKeyDown({ key: 'Enter', target: { value: '复核' } });
    tree = await renderPanelLoaded(panel);
    assert.ok(urls[urls.length - 1].includes('kbq=' + encodeURIComponent('复核')), '检索应带 kbq=');
    assert.ok(urls[urls.length - 1].includes('kbrecursive=1'), '块检索默认递归子文件夹，不跟浏览层的「仅当前层」走');
    assert.ok(textsOf(tree).some((t) => t.includes('复核要能复跑')), '命中卡片要显示正文预览');
    assert.ok(textsOf(tree).some((t) => t.includes('分 1.25')), '命中卡片要显示分数');
    const hitNode = byProp(tree, 'data-kb-hit', 'kb:abc:1')[0];
    assert.equal(hitNode.props.role, 'button', '命中卡片要能被读屏当按钮');
    assert.equal(hitNode.props.tabIndex, 0, '命中卡片要能 Tab 到');

    byProp(tree, 'data-kb-hit', 'kb:abc:1')[0].props.onClick();
    tree = await renderPanelLoaded(panel);
    const last = urls[urls.length - 1];
    assert.ok(last.includes('kbdoc=' + encodeURIComponent('kb:abc')), '点命中应带文档 id');
    assert.ok(last.includes('kbchunk=' + encodeURIComponent('kb:abc:1')), '点命中应带块 id');
    assert.ok(byProp(tree, 'data-kb-chunk-text', 'kb:abc:1').length === 1, '应显示块正文');
    assert.ok(textsOf(tree).some((t) => t.includes('复核要能复跑。')), '块正文要真的上屏');
    // 正文必须跟在选中的那张块卡片后面。挂在清单上方时，用户点第 1 块，
    // 正文跑到清单外、选中的卡片还在下面，看起来像多出一条。
    const nodes = walk(tree);
    const cardAt = nodes.findIndex((n) => n.props && n.props['data-kb-chunk'] === 'kb:abc:1');
    const bodyAt = nodes.findIndex((n) => n.props && n.props['data-kb-chunk-text'] === 'kb:abc:1');
    assert.ok(cardAt >= 0 && bodyAt > cardAt, '块正文要排在选中块卡片之后');

    // 清空检索：输入框与 URL 都要回到「没有 kbq」的状态，不能只在界面上换个空态。
    const clear = byProp(tree, 'data-kb-search-clear', 'true')[0];
    assert.ok(clear, '有过滤词时要给「清空检索」');
    clear.props.onClick();
    tree = await renderPanelLoaded(panel);
    assert.equal(byProp(tree, 'data-kb-search-input', 'true')[0].props.value, '', '清空后输入框不能留着旧词');
    assert.equal(urls[urls.length - 1].includes('kbq='), false, '清空后请求里不该再带 kbq');
});

await check('知识库：块检索无命中时把 reason 译成人话，并写出检索词（第一百一十一轮）', async () => {
    const urls = [];
    const empty = makeSnapshot({ kbSearch: { hits: [], matched: 0, truncated: false, reason: 'no-overlap' } });
    const fetchImpl = (url) => {
        const text = String(url);
        urls.push(text);
        const body = text.includes('kbq=') ? empty : makeSnapshot();
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
    tree = render(panel);
    byProp(tree, 'data-kb-mode', 'search')[0].props.onClick();
    tree = await renderPanelLoaded(panel);
    byProp(tree, 'data-kb-search-input', 'true')[0].props.onChange(eventOf({ value: '量子纠缠' }));
    tree = render(panel);
    byProp(tree, 'data-kb-search-input', 'true')[0].props.onKeyDown({ key: 'Enter', target: { value: '量子纠缠' } });
    tree = await renderPanelLoaded(panel);
    const text = textsOf(tree).join('\n');
    assert.ok(text.includes('搜索「量子纠缠」'), '结果头要写出这次查的词');
    assert.ok(text.includes('查询词与块内容没有重叠'), 'reason 要译成人话');
    assert.ok(!text.includes('no-overlap'), '不要把服务端 reason 原样丢给用户');
});

await check('知识库：顶部搜索按当前看法路由到文档 q 或块 kbq（第一百一十一轮）', async () => {
    const urls = [];
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
    tree = render(panel);

    let top = walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找文档'));
    assert.ok(top, '文档看法下，顶部搜索要如实说只找文档');
    top.props.onChange({ target: { value: '复核' } });
    tree = render(panel);
    top = walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找文档'));
    top.props.onKeyDown({ key: 'Enter', target: { value: '复核' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    let last = urls[urls.length - 1];
    assert.ok(last.includes('q=' + encodeURIComponent('复核')), '文档看法下顶部搜索走文档过滤');
    assert.equal(last.includes('kbq='), false);

    byProp(tree, 'data-kb-mode', 'search')[0].props.onClick();
    tree = render(panel);
    top = walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找块'));
    assert.ok(top, '块检索看法下，顶部搜索要改成找块');
    top.props.onChange({ target: { value: '复核' } });
    tree = render(panel);
    top = walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找块'));
    top.props.onKeyDown({ key: 'Enter', target: { value: '复核' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    last = urls[urls.length - 1];
    assert.ok(last.includes('kbq=' + encodeURIComponent('复核')), '块检索看法下顶部搜索要走 kbq');
    assert.equal(new URL(last, 'http://localhost').searchParams.has('q'), false,
        '块检索不应再带上一轮的文档过滤词');

    const clear = byProp(tree, 'data-memory-search-clear', 'true')[0];
    assert.ok(clear, '块检索后要有清空按钮');
    clear.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    last = urls[urls.length - 1];
    assert.equal(last.includes('kbq='), false, '清空后不该再带 kbq');
    assert.equal(byProp(tree, 'data-kb-search-input', 'true')[0].props.value, '', '工具栏输入框也要清空');
});

await check('记忆面板：摘要卡写清盖了几块、属于哪份文档，摘要页列出来源块并能跳回原文', async () => {
    const urls = [];
    const summaryItem = {
        id: 'kbllm:abc:1', title: '方法论摘要', text: '复核与复跑是一体的。',
        docId: 'kb:abc', chunkIds: ['kb:abc:1', 'kb:abc:2'],
        createdAt: '2026-11-21T02:00:00.000Z', docPath: '来源/方法论.md',
        folder: '来源', docTitle: '方法论笔记',
    };
    const sourceDoc = {
        docId: 'kb:abc', path: '来源/方法论.md', title: '方法论笔记', tier: 'verified',
        origin: 'workspace', chunks_total: 2, chunks_missing: 0, selectedChunk: 0,
        chunks: [
            { id: 'kb:abc:1', n: 1, chars: 7, spanText: '字符 1–7', header: '# 方法论', outline: ['方法论'] },
            { id: 'kb:abc:2', n: 2, chars: 8, spanText: '字符 8–15', header: '## 第二段', outline: ['方法论', '第二段'] },
        ],
    };
    const summaryDoc = {
        docId: 'summary:kbllm:abc:1', path: '摘要/kbllm:abc:1.md', title: '方法论摘要',
        tier: 'verified', origin: 'workspace', synthetic: true,
        text: '复核与复跑是一体的。', chunks_total: 1, chunks_missing: 0,
        summaryOf: ['kb:abc:1', 'kb:abc:2'],
        sources: [
            { id: 'kb:abc:1', docId: 'kb:abc', docPath: '来源/方法论.md', n: 1, head: '# 方法论', chars: 7, spanText: '字符 1–7' },
            { id: 'kb:abc:2', docId: 'kb:abc', docPath: '来源/方法论.md', n: 2, head: '## 第二段', chars: 8, spanText: '字符 8–15' },
        ],
        sourcesMissing: 0,
        chunks: [{ id: 'summary:kbllm:abc:1:1', n: 1, chars: 10, spanText: '字符 1–10', header: '# 方法论摘要', outline: ['方法论摘要'] }],
    };
    const summaryChunk = {
        id: 'summary:kbllm:abc:1:1', n: 1, chars: 10, spanText: '字符 1–10',
        header: '# 方法论摘要', text: '复核与复跑是一体的。', docId: 'summary:kbllm:abc:1',
        path: '摘要/kbllm:abc:1.md', title: '方法论摘要', tier: 'verified', origin: 'workspace',
    };
    const base = makeSnapshot({ kbSummaries: [summaryItem] });
    const drilled = makeSnapshot({
        kbSummaries: [summaryItem], kbDoc: summaryDoc, kbChunk: summaryChunk,
    });
    const jumped = makeSnapshot({
        kbSummaries: [summaryItem], kbDoc: sourceDoc, kbChunk: null,
    });
    const fetchImpl = (url) => {
        const text = String(url);
        urls.push(text);
        const body = text.includes('kbdoc=' + encodeURIComponent('kb:abc')) ? jumped
            : text.includes('kbdoc=' + encodeURIComponent('summary:')) ? drilled
                : base;
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
    tree = render(panel);
    // 第一百零七轮：文档清单在前，摘要块在后；摘要块默认收起，鼠标滚轮不必穿过整段正文。
    const nodes = walk(tree);
    const docsAt = nodes.findIndex((n) => n.props && n.props['data-memory-section'] === 'kb');
    const summariesAt = nodes.findIndex((n) => n.props && n.props['data-kb-summaries'] === 'true');
    assert.ok(docsAt >= 0 && summariesAt > docsAt, '摘要块要排在文档清单之后');
    const summaryToggle = byProp(tree, 'data-fold-toggle', 'kb-summary-kbllm:abc:1')[0];
    assert.ok(summaryToggle, '摘要块要能折叠');
    assert.equal(summaryToggle.props['aria-expanded'], false, '摘要块默认收起');
    const summaryBody = byProp(tree, 'data-fold-body', 'kb-summary-kbllm:abc:1')[0];
    assert.equal(summaryBody.props.style.display, 'none', '收起时摘要正文不占滚动长度');
    assert.ok(textsOf(summaryBody).some((t) => t.includes('2 块来源')), '摘要卡要写清盖了几块');
    assert.ok(textsOf(summaryBody).some((t) => t.includes('来源/方法论.md')), '摘要卡要写清属于哪份文档');

    summaryToggle.props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-fold-toggle', 'kb-summary-kbllm:abc:1')[0].props['aria-expanded'], true, '点标题应展开摘要');
    byProp(tree, 'data-kb-summary', 'kbllm:abc:1')[0].props.onClick();
    tree = await renderPanelLoaded(panel);
    assert.ok(byProp(tree, 'data-kb-summary-sources', 'true').length === 1, '摘要页要有一节「来源块」');
    assert.ok(textsOf(tree).some((t) => t.includes('第 1 块')), '来源块要写清块号');
    assert.ok(textsOf(tree).some((t) => t.includes('字符 8–15')), '来源块要写清字符区间');
    byProp(tree, 'data-kb-summary-source', 'kb:abc:2')[0].props.onClick();
    tree = await renderPanelLoaded(panel);
    assert.ok(urls[urls.length - 1].includes('kbdoc=' + encodeURIComponent('kb:abc')), '点来源块要回到原文文档');
    assert.equal(byProp(tree, 'data-kb-summary-sources', 'true').length, 0, '原文页不该再显示摘要的来源块');
});

await check('记忆面板：摘要块分页走服务端窗口，翻页带 offset 与每页条数（第一百零八轮）', async () => {
    const urls = [];
    const summaryOf = (id, title) => ({
        id, title, text: title + '的正文。', docId: 'kb:abc', chunkIds: ['kb:abc:1'],
        createdAt: '2026-11-21T02:00:00.000Z', docPath: '来源/方法论.md', folder: '来源', docTitle: '方法论笔记',
    });
    const page1 = makeSnapshot({
        kbSummaries: [summaryOf('kbllm:abc:1', '摘要一'), summaryOf('kbllm:abc:2', '摘要二')],
        kbSummariesPage: { offset: 0, limit: 2, total: 5, pageCount: 3, page: 1, hasPrev: false, hasNext: true },
    });
    const page2 = makeSnapshot({
        kbSummaries: [summaryOf('kbllm:abc:3', '摘要三'), summaryOf('kbllm:abc:4', '摘要四')],
        kbSummariesPage: { offset: 2, limit: 2, total: 5, pageCount: 3, page: 2, hasPrev: true, hasNext: true },
    });
    const fetchImpl = (url) => {
        const text = String(url);
        urls.push(text);
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(text.includes('kbsummariesoffset=2') ? page2 : page1) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-kb-summaries-pager', 'true').length, 1, '摘要块要出分页条');
    assert.equal(textsOf(byProp(tree, 'data-kb-summary-page-info', 'true')[0])[0], '第 1–2 条 / 共 5 条');
    assert.equal(byProp(tree, 'data-kb-summary-page-prev', 'true')[0].props.disabled, true, '第一页不能再往前');
    const next = byProp(tree, 'data-kb-summary-page-next', 'true')[0];
    assert.equal(next.props.disabled, false, '还有下一页');
    next.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const last = urls[urls.length - 1];
    assert.ok(last.includes('kbsummariesoffset=2'), '下一页要带新的 offset');
    assert.ok(last.includes('kbsummarieslimit=2'), '每页条数要一起带上');
    assert.equal(textsOf(byProp(tree, 'data-kb-summary-page-info', 'true')[0])[0], '第 3–4 条 / 共 5 条');
});

await check('记忆面板：知识库用文件夹树浏览，当前文件夹决定检索范围（第八十四轮）', async () => {
    const urls = [];
    const base = makeSnapshot();
    const scoped = makeSnapshot({
        kbFolder: '来源',
        kbRecursive: true,
        counts: Object.assign({}, base.counts, { kb: 1, kbTotal: 1, kbChunks: 3, kbBytes: 4096 }),
    });
    const fetchImpl = (url) => {
        const text = String(url);
        urls.push(text);
        return Promise.resolve({
            ok: true,
            status: 200,
            json: () => Promise.resolve(text.includes('kbfolder=') ? scoped : base),
        });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));

    const kbTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb');
    kbTab.props.onClick();
    tree = render(panel);

    assert.equal(byProp(tree, 'data-kb-tree', 'true').length, 1, '知识库要有文件夹树');
    const folderRow = byProp(tree, 'data-kb-folder', '来源')[0];
    assert.ok(folderRow, '树里要有「来源」这一层');
    assert.ok(textsOf(folderRow).includes('1'), '树节点要显示文档数');
    assert.ok(byProp(tree, 'data-kb-crumb', '').length >= 1, '根层要有面包屑');
    assert.ok(textsOf(tree).some((t) => t.includes('全部文档')), '根面包屑要写「全部文档」');

    // 文件夹树要能收起：收起根之后子层不再进 DOM，而不是把箭头画成「收起」却照样铺开。
    const rootRow = byProp(tree, 'data-kb-folder', '')[0];
    const rootToggle = walk(rootRow).find((n) => n.props && n.props['data-kb-toggle'] === 'open');
    assert.ok(rootToggle, '有子层的文件夹要给出展开/收起的小三角');
    rootToggle.props.onClick({ stopPropagation() {} });
    tree = render(panel);
    assert.equal(byProp(tree, 'data-kb-folder', '来源').length, 0, '收起根之后子层不能留在 DOM 里');
    const closedToggle = byProp(tree, 'data-kb-toggle', 'closed')[0];
    assert.ok(closedToggle, '收起后小三角要变成可再展开的状态');
    closedToggle.props.onClick({ stopPropagation() {} });
    tree = render(panel);
    assert.ok(byProp(tree, 'data-kb-folder', '来源')[0], '再点开要恢复子层');

    folderRow.props.onClick();
    tree = await renderPanelLoaded(panel);
    assert.ok(urls[urls.length - 1].includes('kbfolder=' + encodeURIComponent('来源')), '点文件夹要带 kbfolder=');
    assert.ok(textsOf(tree).some((t) => t.includes('文档 · 来源')), '文档分区标题要跟着当前文件夹走');
    assert.ok(textsOf(tree).some((t) => t.includes('仅当前层')), '默认只列当前层');

    const recursive = byProp(tree, 'data-kb-recursive', 'current')[0];
    assert.ok(recursive, '要有「含子文件夹 / 仅当前层」切换');
    recursive.props.onClick();
    tree = await renderPanelLoaded(panel);
    assert.ok(urls[urls.length - 1].includes('kbrecursive=1'), '显式切到子树要带 kbrecursive=1');
    assert.ok(textsOf(tree).some((t) => t.includes('含子文件夹')), '切换后标成含子文件夹');

    // 统一入口（第九十五轮）：上传与识别合并成一颗「＋ 添加文件」，
    // 选完先预览、再显式入库；不再有第二个「识别文件」按钮。
    const upload = byProp(tree, 'data-kb-add-files', 'true')[0];
    assert.ok(upload, '当前文件夹要有「＋ 添加文件」入口');
    assert.ok(textsOf(upload).some((t) => t.includes('添加文件')), '入口文案是添加文件（预览后再入库）');
    assert.equal(byProp(tree, 'data-kb-ocr', 'true').length, 0, '不再有独立的「识别文件」入口');
    assert.equal(byProp(tree, 'data-kb-upload', 'true').length, 0, '不再有独立的「上传文件」入口');
    assert.equal(byProp(tree, 'data-memory-pending', 'kb-ingest').length, 0, '添加文件不该再退回待发送指令');
});

await check('记忆面板：新建文件夹 CTA→输入→POST，成功后进入空子目录、树仍在（第一〇六轮）', async () => {
    const calls = [];
    const base = makeSnapshot();
    const rootTree = JSON.parse(JSON.stringify(base.kbTree));
    rootTree.children.push({ path: '新目录', name: '新目录', docs: 0, chunks: 0, bytes: 0, children: [] });
    const afterCreate = makeSnapshot({
        kbFolder: '新目录',
        kbRecursive: false,
        kb: [],
        counts: Object.assign({}, base.counts, { kb: 0, kbTotal: 1 }),
        kbTree: rootTree,
    });
    const fetchImpl = (url, init) => {
        calls.push({ url: String(url), init });
        if (init && init.method === 'POST') {
            return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({
                ok: true, action: 'kb-folder-create', result: { path: '新目录', created: true },
            }) });
        }
        const text = String(url);
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(
            text.includes('kbfolder=' + encodeURIComponent('新目录')) ? afterCreate : base) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/proj/temp1' } }) };
    let tree = await renderPanelLoaded(panel, props);
    await new Promise((resolve) => setTimeout(resolve, 10));
    const kbTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb');
    kbTab.props.onClick();
    tree = render(panel, props);

    const cta = byProp(tree, 'data-kb-nav', 'new-folder')[0];
    assert.ok(cta, '命令行要有「新建文件夹」入口');
    assert.ok(textsOf(cta).some((t) => t.includes('新建文件夹')));
    cta.props.onClick();
    tree = render(panel, props);
    assert.equal(byProp(tree, 'data-kb-new-folder', 'true').length, 1, '点入口应展开行内新建条');
    assert.ok(textsOf(tree).some((t) => t.includes('在「全部文档」下新建文件夹')), '新建条要写明落在哪一层');
    const input = byProp(tree, 'data-kb-new-folder-input', 'true')[0];
    assert.ok(input, '新建条要有名字输入框');
    input.props.onChange(eventOf({ value: '新目录' }));
    tree = render(panel, props);
    assert.equal(byProp(tree, 'data-kb-new-folder-input', 'true')[0].props.value, '新目录', '输入框受控');
    byProp(tree, 'data-kb-new-folder-ok', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 30));
    tree = render(panel, props);

    const posts = calls.filter((call) => call.url === '/office-memory/mutate');
    assert.equal(posts.length, 1, '只发一次新建请求');
    assert.equal(posts[0].init.method, 'POST', '新建走 POST');
    assert.deepEqual(JSON.parse(posts[0].init.body), { action: 'kb-folder-create', path: '新目录' });
    const gets = calls.filter((call) => call.url.startsWith('/office-memory/snapshot'));
    assert.ok(gets.length >= 2, '成功后要重新取快照');
    const last = gets[gets.length - 1].url;
    assert.ok(last.includes('kbfolder=' + encodeURIComponent('新目录')), '重新取数要进入新目录');
    assert.ok(last.includes('kbrecursive=0'), '新目录默认只看当前层');
    assert.ok(last.includes('cwd=' + encodeURIComponent('D:/proj/temp1')), '重新取数要带当前工作目录');

    assert.equal(byProp(tree, 'data-kb-new-folder', 'true').length, 0, '成功后新建条收起');
    assert.ok(byProp(tree, 'data-kb-crumb', '新目录').length >= 1, '面包屑进入新目录');
    assert.ok(textsOf(tree).some((t) => t.includes('这个文件夹是空的。')), '空子目录要有空态');
    assert.equal(byProp(tree, 'data-kb-empty', 'true').length, 1, '空目录只画空态');
    assert.equal(byProp(tree, 'data-kb-subfolder').length, 0, '空目录没有子文件夹行');
    assert.ok(byProp(tree, 'data-kb-folder', '新目录').length >= 1, '树里仍能看到新目录');
    assert.ok(byProp(tree, 'data-kb-folder', '来源').length >= 1, '原有目录没被挤掉');
    assert.equal(byProp(tree, 'data-memory-action-error', 'true').length, 0);
});

await check('记忆面板：新建文件夹失败时显示原因、不跳转（第一〇六轮）', async () => {
    const calls = [];
    const base = makeSnapshot();
    const fetchImpl = (url, init) => {
        calls.push({ url: String(url), init });
        if (init && init.method === 'POST') {
            return Promise.resolve({ ok: false, status: 409, json: () => Promise.resolve({
                ok: false, code: 'KB_FOLDERS_CORRUPT',
                error: '知识库文件夹清单不是合法 JSON；不会把它当空清单覆盖，请先修好。',
            }) });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(base) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    const props = { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: 'D:/proj/temp1' } }) };
    let tree = await renderPanelLoaded(panel, props);
    await new Promise((resolve) => setTimeout(resolve, 10));
    const kbTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb');
    kbTab.props.onClick();
    tree = render(panel, props);
    byProp(tree, 'data-kb-nav', 'new-folder')[0].props.onClick();
    tree = render(panel, props);
    byProp(tree, 'data-kb-new-folder-input', 'true')[0].props.onChange(eventOf({ value: '坏目录' }));
    tree = render(panel, props);
    byProp(tree, 'data-kb-new-folder-ok', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 30));
    tree = render(panel, props);

    const posts = calls.filter((call) => call.url === '/office-memory/mutate');
    assert.equal(posts.length, 1);
    assert.equal(posts[0].init.method, 'POST');
    assert.deepEqual(JSON.parse(posts[0].init.body), { action: 'kb-folder-create', path: '坏目录' });
    const gets = calls.filter((call) => call.url.startsWith('/office-memory/snapshot'));
    assert.equal(gets.length, 1, '失败后不重新取数，也不假装已经进去');
    assert.equal(gets[0].url.includes('kbfolder='), false, '失败后仍停在根目录');
    assert.ok(byProp(tree, 'data-kb-crumb', '').length >= 1, '面包屑仍在根');
    assert.equal(byProp(tree, 'data-kb-crumb', '坏目录').length, 0, '不能跳到失败的新目录');
    const err = byProp(tree, 'data-memory-action-error', 'true')[0];
    assert.ok(err, '失败原因要显示出来');
    const errText = textsOf(err).join('');
    assert.ok(errText.includes('写入失败'), '要说明是写入失败：' + errText);
    assert.ok(errText.includes('不是合法 JSON'), '要保留服务端给的具体原因：' + errText);
    assert.equal(byProp(tree, 'data-kb-new-folder', 'true').length, 1, '失败后名字还留着，方便改名或重试');
    assert.ok(byProp(tree, 'data-kb-folder', '来源').length >= 1, '失败后树仍完整');
});

await check('记忆面板：编辑器开着时新建文件夹失败仍显示原因（第一〇七轮）', async () => {
    const calls = [];
    const base = makeSnapshot();
    const fetchImpl = (url, init) => {
        calls.push({ url: String(url), init });
        if (init && init.method === 'POST') {
            return Promise.resolve({ ok: false, status: 409, json: () => Promise.resolve({
                ok: false, code: 'KB_FOLDERS_CORRUPT',
                error: '知识库文件夹清单不是合法 JSON；不会把它当空清单覆盖，请先修好。',
            }) });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(base) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    tree = openTab(tree, panel, '知识库（1）');

    byProp(tree, 'data-memory-edit', 'kb:8f14e45f')[0].props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-memory-editor', 'kb-doc').length, 1, '先打开文档编辑器');
    byProp(tree, 'data-kb-nav', 'new-folder')[0].props.onClick();
    tree = render(panel);
    byProp(tree, 'data-kb-new-folder-input', 'true')[0].props.onChange(eventOf({ value: '坏目录' }));
    tree = render(panel);
    byProp(tree, 'data-kb-new-folder-ok', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 30));
    tree = render(panel);

    const err = byProp(tree, 'data-kb-folder-error', 'true')[0];
    assert.ok(err, '新建失败原因要紧贴新建条显示，不能藏进编辑器');
    assert.ok(textsOf(err).join('').includes('不是合法 JSON'), '保留服务端具体原因');
    assert.equal(byProp(tree, 'data-memory-editor', 'kb-doc').length, 1, '错误不能把已打开的编辑器收起');
    assert.equal(byProp(tree, 'data-kb-crumb', '坏目录').length, 0, '失败不能跳转');
    assert.equal(byProp(tree, 'data-kb-new-folder', 'true').length, 1, '草稿保留，方便改名重试');
    assert.equal(calls.filter((call) => call.url.startsWith('/office-memory/snapshot')).length, 1, '失败不再取快照');
});

await check('记忆面板：新建成功后快照失败保留父目录并给出确认/重试（第一〇七轮）', async () => {
    const calls = [];
    const base = makeSnapshot();
    const rootTree = JSON.parse(JSON.stringify(base.kbTree));
    rootTree.children.push({ path: '新目录', name: '新目录', docs: 0, chunks: 0, bytes: 0, children: [] });
    const child = makeSnapshot({
        kbFolder: '新目录', kbRecursive: false,
        kb: [], counts: Object.assign({}, base.counts, { kb: 0, kbTotal: 1 }),
        kbTree: rootTree,
    });
    let childGets = 0;
    const fetchImpl = (url, init) => {
        calls.push({ url: String(url), init });
        if (init && init.method === 'POST') {
            return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({
                ok: true, action: 'kb-folder-create', result: { path: '新目录', created: true },
            }) });
        }
        const text = String(url);
        if (text.includes('kbfolder=' + encodeURIComponent('新目录'))) {
            childGets += 1;
            if (childGets === 1) {
                return Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({
                    ok: false, error: '快照暂时不可用',
                }) });
            }
            return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(child) });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(base) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    tree = openTab(tree, panel, '知识库（1）');
    byProp(tree, 'data-kb-nav', 'new-folder')[0].props.onClick();
    tree = render(panel);
    byProp(tree, 'data-kb-new-folder-input', 'true')[0].props.onChange(eventOf({ value: '新目录' }));
    tree = render(panel);
    byProp(tree, 'data-kb-new-folder-ok', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 40));
    tree = render(panel);

    assert.equal(calls.filter((call) => call.url === '/office-memory/mutate').length, 1);
    assert.equal(childGets, 1, 'POST 成功后只取一次新目录快照，失败先停住');
    assert.ok(byProp(tree, 'data-kb-crumb', '').length >= 1, '读取失败时仍挂在父目录');
    assert.equal(byProp(tree, 'data-kb-crumb', '新目录').length, 0, '不能显示「新目录 + 父目录的行」');
    assert.ok(byProp(tree, 'data-kb-file', 'kb:8f14e45f').length >= 1, '父目录的文档行保持原样');
    const notice = byProp(tree, 'data-kb-folder-notice', 'true')[0];
    assert.ok(notice, '已创建但读取失败要单独说明');
    const noticeText = textsOf(notice).join('');
    assert.ok(noticeText.includes('已创建'), '明确目录已经创建：' + noticeText);
    assert.ok(noticeText.includes('新目录'), '说明是哪一个目录：' + noticeText);
    assert.ok(noticeText.includes('快照暂时不可用'), '保留读取失败的原因：' + noticeText);
    const retry = byProp(tree, 'data-kb-folder-retry', 'true')[0];
    assert.ok(retry, '要给出显式重试，不必先退回再进');

    retry.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 30));
    tree = render(panel);
    assert.equal(childGets, 2, '重试应重新取新目录快照');
    assert.ok(byProp(tree, 'data-kb-crumb', '新目录').length >= 1, '重试成功后再进入新目录');
    assert.equal(byProp(tree, 'data-kb-empty', 'true').length, 1, '空目录给空态');
    assert.equal(byProp(tree, 'data-kb-folder-notice', 'true').length, 0, '成功后清掉失败提示');
});

await check('记忆面板：目录切换快照失败时停在旧目录并给重试（第一〇七轮）', async () => {
    const calls = [];
    const base = makeSnapshot();
    const scoped = makeSnapshot({ kbFolder: '来源', kbRecursive: false });
    let scopedGets = 0;
    const fetchImpl = (url, init) => {
        calls.push({ url: String(url), init });
        const text = String(url);
        if (text.includes('kbfolder=' + encodeURIComponent('来源'))) {
            scopedGets += 1;
            if (scopedGets === 1) {
                return Promise.resolve({ ok: true, status: 500, json: () => Promise.resolve({
                    ok: false, error: '目录清单暂时不可用',
                }) });
            }
            return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(scoped) });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(base) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    tree = openTab(tree, panel, '知识库（1）');
    byProp(tree, 'data-kb-folder', '来源')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 40));
    tree = render(panel);

    assert.equal(scopedGets, 1);
    assert.ok(byProp(tree, 'data-kb-crumb', '').length >= 1, '读取失败时面包屑不能先跳到新目录');
    assert.equal(byProp(tree, 'data-kb-crumb', '来源').length, 0);
    assert.ok(byProp(tree, 'data-kb-file', 'kb:8f14e45f').length >= 1, '旧目录的文档行不能被当成新目录的行');
    const notice = byProp(tree, 'data-kb-folder-notice', 'true')[0];
    assert.ok(notice, '切换失败要说明仍停留在哪一层');
    const noticeText = textsOf(notice).join('');
    assert.ok(noticeText.includes('没能打开'), noticeText);
    assert.ok(noticeText.includes('来源'), noticeText);
    assert.ok(noticeText.includes('仍停留在'), noticeText);
    const retry = byProp(tree, 'data-kb-folder-retry', 'true')[0];
    assert.ok(retry, '切换失败也要给显式重试');

    retry.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 30));
    tree = render(panel);
    assert.equal(scopedGets, 2);
    assert.ok(byProp(tree, 'data-kb-crumb', '来源').length >= 1, '重试成功后再进入目标目录');
    assert.equal(byProp(tree, 'data-kb-folder-notice', 'true').length, 0, '成功后清掉失败提示');
    const back = byProp(tree, 'data-kb-nav', 'back')[0];
    assert.ok(back && back.props.disabled !== true, '重试成功要恢复后退栈，不能丢掉失败前的目录');
});

await check('记忆面板：热记忆就地编辑走 POST，保存后重新取数（第八十五轮）', async () => {
    const calls = [];
    const fetchImpl = (url, init) => {
        calls.push({ url, init });
        const body = init && init.method === 'POST'
            ? { ok: true, action: 'memory-replace', result: { id: 'm-1' } }
            : makeSnapshot();
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    const edit = byProp(tree, 'data-memory-edit', 'm-1')[0];
    assert.ok(edit, '热记忆条目要有「编辑」按钮');
    edit.props.onClick();
    tree = render(panel);
    const textarea = byProp(tree, 'data-memory-editor-text', 'true')[0];
    assert.ok(textarea, '点编辑后应打开正文编辑框');
    textarea.props.onChange(eventOf({ value: '偏好：更短更直' }));
    tree = render(panel);
    const target = byProp(tree, 'data-memory-editor-target', 'true')[0];
    target.props.onChange(eventOf({ value: 'user' }));
    tree = render(panel);
    byProp(tree, 'data-memory-save', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));

    const post = calls.find((call) => call.url === '/office-memory/mutate');
    assert.ok(post, '保存应 POST 到同源的写入端点');
    assert.equal(post.init.method, 'POST');
    const payload = JSON.parse(post.init.body);
    assert.equal(payload.action, 'memory-replace');
    assert.equal(payload.id, 'm-1');
    assert.equal(payload.content, '偏好：更短更直');
    assert.equal(payload.target, 'user');
    assert.ok(calls.filter((call) => String(call.url).startsWith('/office-memory/snapshot')).length >= 2,
        '保存成功后应重新取一次快照，而不是乐观改本地');
    tree = render(panel);
    assert.equal(byProp(tree, 'data-memory-editor-text', 'true').length, 0, '保存成功后编辑器收起');
});

await check('记忆面板：新增记忆直接写 POST，不生成待发送指令（第八十七轮）', async () => {
    const calls = [];
    const fetchImpl = (url, init) => {
        calls.push({ url, init });
        const body = init && init.method === 'POST'
            ? { ok: true, action: 'memory-add', result: { id: 'm-new' } }
            : makeSnapshot();
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);

    const compose = byProp(tree, 'data-memory-compose', 'true')[0];
    assert.ok(compose, '标题栏要有「+ 记一条」');
    compose.props.onClick();
    tree = render(panel);

    assert.equal(byProp(tree, 'data-memory-pending').length, 0, '点新增不应再生成待发送指令卡');
    assert.equal(byProp(tree, 'data-memory-editor', 'new').length, 0, '标题栏按钮不再另开一张草稿卡，避免和页内新增框重复');
    const textarea = byProp(tree, 'data-memory-quick-text', 'true')[0];
    assert.ok(textarea, '页内「直接记一条」的正文框应存在');
    textarea.props.onChange(eventOf({ value: '约定：交付前先跑全量套件。' }));
    tree = render(panel);
    byProp(tree, 'data-memory-quick-target-user', 'true')[0].props.onClick();
    tree = render(panel);
    byProp(tree, 'data-memory-quick-add', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));

    const post = calls.find((call) => call.url === '/office-memory/mutate');
    assert.ok(post, '保存应直接 POST 到同源写入端点');
    const payload = JSON.parse(post.init.body);
    assert.equal(payload.action, 'memory-add');
    assert.equal(payload.content, '约定：交付前先跑全量套件。');
    assert.equal(payload.target, 'user');
    assert.equal(payload.importance, 'normal');
    assert.ok(calls.filter((call) => String(call.url).startsWith('/office-memory/snapshot')).length >= 2,
        '写入成功后重新取快照');
    tree = render(panel);
    assert.equal(byProp(tree, 'data-memory-editor', 'new').length, 0, '保存后草稿收起');
});

await check('记忆面板：知识库文档改标题/文件夹/来源档走 kb-update（第八十五轮）', async () => {
    const calls = [];
    const fetchImpl = (url, init) => {
        calls.push({ url, init });
        const body = init && init.method === 'POST' ? { ok: true, action: 'kb-update' } : makeSnapshot();
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    tree = openTab(tree, panel, '知识库（1）');
    const edit = byProp(tree, 'data-memory-edit', 'kb:8f14e45f')[0];
    assert.ok(edit, '知识库文档也要有「编辑」按钮');
    edit.props.onClick();
    tree = render(panel);
    byProp(tree, 'data-memory-editor-title', 'true')[0].props.onChange(eventOf({ value: '方法论（修订）' }));
    tree = render(panel);
    byProp(tree, 'data-memory-editor-folder', 'true')[0].props.onChange(eventOf({ value: '来源/修订' }));
    tree = render(panel);
    byProp(tree, 'data-memory-editor-tier', 'true')[0].props.onChange(eventOf({ value: 'unverified' }));
    tree = render(panel);
    byProp(tree, 'data-memory-save', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const payload = JSON.parse(calls.find((call) => call.url === '/office-memory/mutate').init.body);
    assert.equal(payload.action, 'kb-update');
    assert.equal(payload.id, 'kb:8f14e45f');
    assert.equal(payload.title, '方法论（修订）');
    assert.equal(payload.folder, '来源/修订');
    assert.equal(payload.tier, 'unverified');
});

await check('记忆面板：知识库改正文与多块合成摘要都走写入端点（第八十五轮）', async () => {
    const calls = [];
    const kbDoc = {
        action: 'kb-doc', docId: 'kb:8f14e45f', path: '来源/方法论.md', title: '方法论笔记',
        tier: 'verified', at: '2026-09-30T10:00:00.000Z', origin: 'workspace',
        text: '# 方法论\n\n第一段。\n\n## 第二段\n\n第二段。',
        chunks: [
            { id: 'kb:8f14e45f:1', n: 1, chars: 8, span: { start: 1, end: 8 }, spanText: '字符 1–8', outline: ['方法论'], header: '# 方法论' },
            { id: 'kb:8f14e45f:2', n: 2, chars: 9, span: { start: 9, end: 18 }, spanText: '字符 9–18', outline: ['第二段'], header: '## 第二段' },
        ],
        chunks_total: 2, chunks_missing: 0, total: 2,
    };
    const fetchImpl = (url, init) => {
        calls.push({ url, init });
        const body = init && init.method === 'POST'
            ? { ok: true, action: 'kb-edit' }
            : (String(url).includes('kbdoc=') ? makeSnapshot({ kbDoc }) : makeSnapshot());
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    tree = openTab(tree, panel, '知识库（1）');
    byProp(tree, 'data-kb-open', 'kb:8f14e45f')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);

    const editText = byProp(tree, 'data-kb-edit-text', 'true')[0];
    assert.ok(editText, '文档抬头要有「改正文」');
    editText.props.onClick();
    tree = render(panel);
    const textarea = byProp(tree, 'data-kb-editor-text', 'true')[0];
    assert.ok(textarea, '改正文应打开整篇编辑框');
    assert.ok(String(textarea.props.value).includes('第一段'), '编辑框要回填整篇正文');
    textarea.props.onChange(eventOf({ value: '# 方法论\n\n改后的正文。' }));
    tree = render(panel);
    const saveText = byProp(tree, 'data-memory-save', 'true')[0];
    assert.ok(saveText, '改正文草稿要有保存按钮');
    saveText.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const editPayload = JSON.parse(calls.find((call) => call.url === '/office-memory/mutate').init.body);
    assert.equal(editPayload.action, 'kb-edit');
    assert.equal(editPayload.id, 'kb:8f14e45f');
    assert.equal(editPayload.text, '# 方法论\n\n改后的正文。');

    // 摘要：选两块 → 打开摘要草稿 → 写正文 → 保存。
    tree = render(panel);
    const boxes = byProp(tree, 'data-kb-select');
    assert.equal(boxes.length, 2, '文档下钻后每块都要有选择框');
    boxes[0].props.onChange();
    tree = render(panel);
    const second = byProp(tree, 'data-kb-select', 'kb:8f14e45f:2')[0];
    assert.ok(second, '第二块的选择框也要在');
    second.props.onChange();
    tree = render(panel);
    const openSummary = byProp(tree, 'data-kb-summarize-open', 'true')[0];
    assert.ok(openSummary, '选满两块后应能打开摘要草稿');
    openSummary.props.onClick();
    tree = render(panel);
    const summaryText = byProp(tree, 'data-kb-summary-text', 'true')[0];
    assert.ok(summaryText, '摘要草稿要有正文框');
    summaryText.props.onChange(eventOf({ value: '两块合起来是方法论。' }));
    tree = render(panel);
    const saveSummary = byProp(tree, 'data-kb-summary-save', 'true')[0];
    assert.ok(saveSummary, '摘要草稿要有保存按钮');
    saveSummary.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const summaryPayload = JSON.parse(calls.filter((call) => call.url === '/office-memory/mutate').pop().init.body);
    assert.equal(summaryPayload.action, 'kb-summarize');
    assert.deepEqual(summaryPayload.chunkIds, ['kb:8f14e45f:1', 'kb:8f14e45f:2']);
    assert.equal(summaryPayload.text, '两块合起来是方法论。');
});

await check('记忆面板：知识库文档可删、块关系可加可解（第八十九轮）', async () => {
    const calls = [];
    const kbDoc = {
        action: 'kb-doc', docId: 'kb:8f14e45f', path: '来源/方法论.md', title: '方法论笔记',
        tier: 'verified', at: '2026-09-30T10:00:00.000Z', origin: 'workspace',
        text: '# 方法论\n\n第一段。',
        chunks: [{ id: 'kb:8f14e45f:1', n: 1, chars: 8, header: '# 方法论', spanText: '字符 1–8', text: '第一段。' }],
        chunks_total: 2, chunks_missing: 0, total: 2,
    };
    const kbChunk = {
        id: 'kb:8f14e45f:1', docId: 'kb:8f14e45f', n: 1, chars: 8,
        header: '# 方法论', spanText: '字符 1–8', text: '第一段。',
    };
    const kbRelated = { items: [{ chunkId: 'kb:8f14e45f:2', docId: 'kb:8f14e45f', kind: 'related', n: 2, title: '第二段', text: '第二段。' }] };
    const fetchImpl = (url, init) => {
        calls.push({ url, init });
        if (init && init.method === 'POST') {
            const body = JSON.parse(String(init.body));
            return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, action: body.action }) });
        }
        const urlText = String(url);
        const body = urlText.includes('kbdoc=')
            ? makeSnapshot({
                kbDoc,
                kbChunk: urlText.includes('kbchunk=') ? kbChunk : null,
                kbRelated: urlText.includes('kbchunk=') ? kbRelated : null,
            })
            : makeSnapshot();
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    tree = openTab(tree, panel, '知识库（1）');
    byProp(tree, 'data-kb-open', 'kb:8f14e45f')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.ok(byProp(tree, 'data-kb-doc-delete', 'kb:8f14e45f').length === 1, '文档抬头要有删除入口');

    byProp(tree, 'data-kb-chunk', 'kb:8f14e45f:1')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.ok(byProp(tree, 'data-kb-link-open', 'kb:8f14e45f:1').length === 1, '块正文要有「加关系」');
    assert.ok(byProp(tree, 'data-kb-unlink', 'kb:8f14e45f:2').length === 1, '相邻块每条要有「解除」');

    byProp(tree, 'data-kb-link-open', 'kb:8f14e45f:1')[0].props.onClick();
    tree = render(panel);
    byProp(tree, 'data-kb-link-target', 'true')[0].props.onChange(eventOf({ value: 'kb:8f14e45f:2' }));
    tree = render(panel);
    byProp(tree, 'data-kb-link-kind', 'true')[0].props.onChange(eventOf({ value: 'supports' }));
    tree = render(panel);
    byProp(tree, 'data-kb-link-note', 'true')[0].props.onChange(eventOf({ value: '同属方法论' }));
    tree = render(panel);
    byProp(tree, 'data-kb-link-save', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const link = JSON.parse(calls.filter((call) => call.url === '/office-memory/mutate').pop().init.body);
    assert.equal(link.action, 'kb-link');
    assert.equal(link.sourceId, 'kb:8f14e45f:1');
    assert.equal(link.targetId, 'kb:8f14e45f:2');
    assert.equal(link.kind, 'supports');
    assert.equal(link.note, '同属方法论');

    tree = render(panel);
    byProp(tree, 'data-kb-unlink', 'kb:8f14e45f:2')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const unlink = JSON.parse(calls.filter((call) => call.url === '/office-memory/mutate').pop().init.body);
    assert.equal(unlink.action, 'kb-unlink');
    assert.equal(unlink.sourceId, 'kb:8f14e45f:1');
    assert.equal(unlink.targetId, 'kb:8f14e45f:2');

    tree = render(panel);
    byProp(tree, 'data-kb-doc-delete', 'kb:8f14e45f')[0].props.onClick();
    tree = render(panel);
    assert.ok(byProp(tree, 'data-memory-delete-card', 'kb:8f14e45f').length === 1, '删除要先出确认卡');
    byProp(tree, 'data-memory-delete-confirm', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const drop = JSON.parse(calls.filter((call) => call.url === '/office-memory/mutate').pop().init.body);
    assert.equal(drop.action, 'kb-drop');
    assert.equal(drop.id, 'kb:8f14e45f', '删的是整篇文档');
});

await check('记忆面板：知识库为空或来源筛选后为空时给可读空态（第八十二轮）', async () => {
    const base = makeSnapshot();
    const emptyKb = makeSnapshot({
        kb: [],
        // 库是空的，文件夹树也必须是空的：夹具默认那棵树里有「来源/方法论」，
        // 留着它就和 counts 自相矛盾，资源管理器会照树画出一堆空文件夹。
        kbTree: { path: '', name: '全部文档', docs: 0, chunks: 0, bytes: 0, children: [] },
        counts: Object.assign({}, base.counts, { kb: 0, kbTotal: 0, kbChunks: 0, kbBytes: 0 }),
        originCounts: Object.assign({}, base.originCounts, { kb: { all: 0, global: 0, workspace: 0 } }),
    });
    const { mod, ctx, registrations } = loadBundle({
        fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(emptyKb) }),
    });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    const kbTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb');
    kbTab.props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-kb-empty', 'true').length, 1, '空库要画空态，不是一块空白');
    assert.ok(textsOf(tree).some((t) => t.includes('添加文件') || t.includes('拖进这一页')), '空态要给添加文件的下一步');
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
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    assert.equal(urls.length, 1, '挂载时应取一次数');
    assert.equal(urls[0], '/office-memory/snapshot', '同源相对路径，且首次不带参数');

    const search = walk(tree).find((n) => n.type === 'input' && n.props && typeof n.props.placeholder === 'string' && n.props.placeholder.includes('找'));
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
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    const tree = await renderPanelLoaded(panel);
    const texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('不在已知列表里')), '应显示服务端给的错误');
    // 出错不等于白屏：Tab 仍在。这一份**根本没有 counts**（失败响应里没有 data），
    // 所以条数按「这个端点没告诉我」显示破折号，而不是谎报 0（与指标块同一口径）。
    assert.ok(texts.some((t) => t === '记忆'), '出错时也应保留 Tab 骨架');
});

await check('记忆面板：环境没有 fetch 时给出可读提示，不抛错', () => {
    const { mod, ctx, registrations } = loadBundle({ fetch: undefined });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
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
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    // options.ownerProps：第六十二轮起宿主会往这一格传 owner props
    // （openView / viewRequest / …），要测「拿得到 openView 时走哪条路」就得传进来。
    const tree = await renderPanelLoaded(panel, options.ownerProps);
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
    // 第八十八轮把「关系」和「实体」合并成「图谱」：旧测试点的是概念，
    // 这里按概念映射到现在的页签，不必每处都改调用。
    const base = label.replace(/（\d+）$/, '');
    const wanted = (base === '实体' || base === '关系') ? '图谱' : base;
    const findTab = (root) => walk(root).find((n) => n.props && n.props.role === 'tab'
        && Array.isArray(n.props.children)
        && n.props.children.some((child) => typeof child === 'string'
            && (child === wanted || child.startsWith(wanted + '（'))));
    let tab = findTab(tree);
    // 一级「记忆」下的层级分段只在记忆视图里渲染；从知识库 / 图谱点「热记忆」「台账」
    // 要先把一级切回「记忆」。
    if (!tab && (wanted === '热记忆' || wanted === '台账')) {
        const memoryTab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'memory');
        assert.ok(memoryTab, '应找到一级「记忆」页签');
        memoryTab.props.onClick();
        tree = render(panel);
        tab = findTab(tree);
    }
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

await check('可视化：标题下四枚指标块显示 counts 里的条数', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const tiles = byProp(tree, 'data-metric');
    assert.deepEqual(tiles.map((n) => n.props['data-metric']), ['hot', 'ledger', 'archive', 'kb'],
        '四枚指标块：热记忆 / 台账 / 归档 / 知识库');
    const values = byProp(tree, 'data-metric-value');
    assert.equal(values.length, 4);
    for (const tile of values) {
        assert.deepEqual(tile.props.children, ['1'], `${tile.props['data-metric-value']} 应显示 counts 里的 1`);
    }
    const texts = textsOf(tree);
    for (const label of ['热记忆', '台账', '归档', '知识库']) {
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
        stores: [{ id: 'workspace', dir: '.office/memory', hot: 1, ledger: 1, archive: 1, bytes: 2048, files: 7 }],
    }));
    const card = byProp(tree, 'data-store', 'workspace')[0];
    assert.ok(card, '应有 workspace 存储卡');
    const texts = textsOf(card);
    assert.ok(texts.includes('.office/memory'), '卡片要写目录');
    for (const chip of ['热 1', '台账 1', '归档 1']) {
        assert.ok(texts.includes(chip), `卡片缺少层的条数：${chip}`);
    }
    assert.ok(texts.includes('2 KB'), '字节数要人读化（2048 → 2 KB）');
    assert.ok(texts.includes('7 个文件'), '有 files 时要显示文件数');
});

await check('可视化：存储卡给出知识库的文档数与块数（第四十二轮）', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        stores: [{
            id: 'workspace', dir: '.office/memory', hot: 1, ledger: 1, archive: 1, links: 1,
            kb: { docs: 2, chunks: 7, bytes: 8192, tiers: { verified: 1, user: 1, unverified: 0 } },
        }],
    }));
    const card = byProp(tree, 'data-store', 'workspace')[0];
    const texts = textsOf(card);
    assert.ok(texts.includes('知识库 2 篇 / 7 块'), '存储卡要同时给文档数与块数');
});

await check('知识库页签：详情条给档位 / 块数 / 文档 id，并如实报「库里一共多少」', async () => {
    // 列表里只给了 1 条，但库里其实有 5 篇：靠服务端分页信息说清「这是第 1 篇、共 5 篇」，
    // 不再用「要全部用 kb-list」把用户赶去命令行。
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: {
            hot: 1, ledger: 1, archive: 1, links: 1, entities: 1, kb: 5,
            ledgerTotal: 1, archiveFiles: 1, archiveItems: 1,
            kbTotal: 5, kbChunks: 3, kbBytes: 4096, kbTruncated: false,
        },
        kbPage: { offset: 0, limit: 20, total: 5, pageCount: 1, page: 1, hasPrev: false, hasNext: false },
    }));
    const kbTree = openTab(tree, panel, '知识库（5）');
    const kbNote = byProp(kbTree, 'data-kb-note')[0];
    assert.ok(kbNote, '知识库页签要有那条只读说明');
    const noteText = textsOf(kbNote).join(' ');
    assert.ok(noteText.includes('库里共 5 篇'), '要报真实总数（不是列表里那 1 篇）');
    assert.ok(noteText.includes('3 块'), '要报块总数');
    // 真分页挪到了底部状态栏：页码 + 总数 + 翻页，旧文案必须消失。
    const statusBar = byProp(kbTree, 'data-kb-status')[0];
    assert.ok(statusBar, '知识库页签要有底部状态栏');
    const statusText = textsOf(statusBar).join(' ');
    assert.ok(!statusText.includes('清单已截断'), '不再有「仅显示这一页」的旧说法');
    assert.ok(!statusText.includes('要全部用 kb-list'), '不再把看全部推给命令行');
    const pageInfo = byProp(kbTree, 'data-kb-page-info', 'true')[0];
    assert.ok(pageInfo, '状态栏要有页码');
    assert.ok(textsOf(pageInfo).join(' ').includes('第 1–1 篇 / 共 5 篇'), '页码要写清当前窗口与总数');
    assert.equal(byProp(kbTree, 'data-kb-page-next', 'true')[0].props.disabled, true, '只有一页时下一页要禁用');
    assert.equal(byProp(kbTree, 'data-kb-page-prev', 'true')[0].props.disabled, true, '首页的上一页要禁用');
    const card = byProp(kbTree, 'data-item-key', 'kb:8f14e45f')[0];
    assert.ok(card, '知识库条目要按文档 id 建卡片');
    const cardTexts = textsOf(card);
    assert.ok(cardTexts.some((t) => t.includes('来源/方法论.md')), '卡片要显示入库路径');
    assert.ok(cardTexts.some((t) => t === 'verified'), '卡片要显示来源档');

    // 悬停 → 底部详情条（键盘聚焦走同一条路）
    card.props.onMouseEnter();
    const detail = byProp(render(panel), 'data-item-detail', 'kb:8f14e45f')[0];
    assert.ok(detail, '悬停知识库条目时要出详情条');
    const detailText = textsOf(detail).join('\n');
    assert.ok(detailText.includes('来源档：verified'), '详情要写档位');
    assert.ok(detailText.includes('知识块：3 块'), '详情要写块数');
    assert.ok(detailText.includes('kb:8f14e45f'), '详情要给会话侧 kb-read 用的文档 id');
});

await check('可视化：台账按天画时间线，柱高按峰值折算', async () => {
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
});

await check('可视化：个人偏好 / 项目记忆分档，重要度落在左边框上', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 3, hotUser: 1, hotProject: 2, ledger: 1, archive: 1, links: 1, entities: 1 },
        hot: [
            { id: 'm-1', target: 'user', importance: 'critical', origin: 'global', updatedAt: 'x', entities: [], tags: [], content: '偏好：书面化表达' },
            { id: 'm-2', target: 'project', importance: 'low', origin: 'workspace', updatedAt: 'x', entities: [], tags: [], content: '约定：只用 edit 工具' },
            { id: 'm-3', importance: 'normal', origin: 'workspace', updatedAt: 'x', entities: [], tags: [], content: '没有 target 的条目' },
        ],
    }));
    const layers = byProp(tree, 'data-layer');
    assert.deepEqual(layers.map((n) => n.props['data-layer']), ['preference', 'project', 'ledger'], '热记忆拆成个人偏好 / 项目记忆两档，并补回只读归档');
    assert.ok(textsOf(tree).some((t) => t === '个人偏好（1）'), '个人偏好档带条数');
    assert.ok(textsOf(tree).some((t) => t === '项目记忆（2）'), '没写 target 的条目归项目记忆');

    byProp(tree, 'data-layer', 'project')[0].props.onClick();
    const project = render(panel);
    assert.ok(textsOf(project).some((t) => t.includes('只用 edit 工具')), '项目记忆档显示 project 条目');
    assert.ok(!textsOf(project).some((t) => t.includes('偏好：书面化表达')), '项目记忆档不显示个人偏好');

    // 侧色带用 inset 阴影而不是 border-left：不改盒模型，也不会被圆角裁掉。
    const normal = byProp(project, 'data-importance', 'normal')[0];
    assert.ok(String(normal.props.style.boxShadow).includes('inset 3px 0 0'), '左侧画 3px 侧色带');
    assert.ok(String(normal.props.style.boxShadow).includes('state-business-primary'));
    assert.ok(String(byProp(project, 'data-importance', 'low')[0].props.style.boxShadow).includes('border-l3'));
});

await check('可视化：空快照照样渲染（指标 0），不出现 NaN / undefined', async () => {
    const { tree } = await panelFrom(makeSnapshot({
        counts: { hot: 0, ledger: 0, archive: 0, ledgerTotal: 0, archiveFiles: 0, archiveItems: 0 },
        stores: [],
        hot: [], ledger: [], archive: [],
    }));

    assert.equal(byProp(tree, 'data-metric').length, 4, '空快照也保留四枚指标块');
    // 老版本端点没有 counts.kb（第四十二轮才加）：那一枚显示破折号，**不谎报 0**
    // ——「知识库 0 篇」与「这个端点不知道知识库」是两件事。
    assert.deepEqual(byProp(tree, 'data-metric-value').map((n) => n.props.children[0]), ['0', '0', '0', '—']);
    assert.equal(byProp(tree, 'data-panel', 'stores').length, 0, '没有 stores 就不画存储域');
    assert.equal(byProp(tree, 'data-gauge', 'ledger')[0].props['data-percent'], 0, '0 / 500 = 0%');
    assert.equal(byProp(tree, 'data-panel', 'ledger-timeline').length, 0, '没有台账就不画时间线');
    assert.equal(byProp(tree, 'data-graph').length, 0, '图谱已下线，任何位置都不该再有图');

    for (const text of stringsOf(tree)) {
        assert.ok(!text.includes('NaN'), `渲染里出现了 NaN：${text}`);
        assert.ok(!text.includes('undefined'), `渲染里出现了 undefined：${text}`);
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

await check('版式：一级页签只写去处、二级分段承载条数', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const tabs = byProp(tree, 'data-tab');
    assert.deepEqual(tabs.map((n) => n.props['data-tab']), ['memory', 'kb', 'gallery'],
        '一级视图：记忆 / 知识库 / 图库');
    assert.deepEqual(byProp(tree, 'data-layer').map((n) => n.props['data-layer']), ['preference', 'project', 'ledger'],
        '记忆视图内部把个人偏好 / 项目记忆 / 台账降为二级层级');
    assert.ok(tabs.every((n) => n.props.role === 'tab'), '页签要有 tab 语义');
    assert.ok(tabs.every((n) => Array.isArray(n.props.children) && n.props.children.length === 1 && typeof n.props.children[0] === 'string'),
        '一级页签只写去处（记忆 / 知识库 / 图库），不再把计数单位混在一条轴上');
    const active = tabs.filter((n) => n.props['aria-selected'] === true);
    assert.equal(active.length, 1, '同时只有一个选中页签');
    assert.equal(active[0].props.style.fontWeight, 600, '选中的页签要加粗');
    // 面板根声明 container-type：否则容器查询不生效，窄面板仍会挤成两栏。
    const root = walk(tree).find((n) => n.props && typeof n.props.className === 'string'
        && n.props.className.split(' ').indexOf('om-mem') !== -1);
    assert.ok(root, '面板根要有 om-mem 类名');
    assert.equal(root.props.style.containerType, 'inline-size');
    // 标题不吸顶（用户要求滚动时让出标题），但页签 + 筛选仍绑在吸顶条里：
    // 否则滚到长列表中间时页签已经滑出视野，想换层只能先把整页滚回顶部。
    const stickyNav = byProp(tree, 'data-memory-nav', 'sticky')[0];
    assert.ok(stickyNav, '页签吸顶条应存在');
    assert.ok(walk(stickyNav).some((n) => n.props && n.props['data-tab']),
        '页签要在吸顶条里');
    const topBar = byProp(tree, 'data-memory-head', 'top')[0];
    assert.ok(topBar && topBar.props.style.position === 'relative', '标题区不吸顶');
});

await check('面板**不是**自己的滚动容器（滚动交给宿主的会话滚动容器）', async () => {
    // 第六十二轮修正。老写法（面板挂在 main 那一格时是对的）是「撑满 + 自己滚」：
    // height:100% + minHeight:0 + flex:'1 1 auto' + overflowY:auto + overscrollBehavior:contain。
    // 迁到 conversation.view 之后窗口变了：宿主只在 `.…_scrollBody:has([data-conversation-composer-overlay])`
    // 命中时才给 viewArea 定高（实测只有「轨迹」页签命中），其余页签（对话 / 上下文 /
    // 编队 / 记忆）的 viewArea 都撑到内容高，滚动归宿主的 `.…_scrollBody`。
    // 于是旧写法在真实运行时变成：面板高＝内容高（滚不动）＋ `overscroll-behavior:contain`
    // 把滚轮链断在这一层（整页也滚不动）。活体实测见 test/web-verify.mjs 的滚动契约。
    const { tree } = await panelFrom(makeSnapshot());
    const root = walk(tree).find((n) => n.props && typeof n.props.className === 'string'
        && n.props.className.split(' ').indexOf('om-scroll') !== -1 && n.props['data-memory-panel'] === 'scroll');
    assert.ok(root, '面板根仍在（om-scroll 管 reduced-motion 下的 scroll-behavior）');
    assert.equal(root.props.style.height, undefined, '不要再撑满高度：宿主那层才是滚动的定高盒子');
    assert.equal(root.props.style.maxHeight, undefined);
    assert.equal(root.props.style.minHeight, undefined);
    assert.equal(root.props.style.flex, undefined, '面板不该再跟 viewArea 抢主轴高度');
    assert.equal(root.props.style.overflowY, undefined, '面板不是滚动容器');
    assert.equal(root.props.style.overscrollBehavior, undefined,
        'overscroll-behavior:contain 会在「有 overflow 但没溢出」的盒子上把滚轮链断掉（实测整页滚不动）');
    assert.equal(root.props.style.overflowX, 'clip',
        '横向只裁切；用 hidden 会把本层变回滚动容器（overflow-y 被算成 auto）');
    assert.equal(root.props.style.containerType, 'inline-size', '容器查询锚点不能丢');
    // 头部吸顶：面板跟宿主一起滚之后，搜索框与「刷新 / 记一条」不能跟着滑出视野。
    // 吸的是宿主滚动容器的顶边（本层不再是滚动容器）。
    const head = walk(tree).find((n) => n.props && n.props['data-memory-nav'] === 'sticky'
        && n.props.style && n.props.style.position === 'sticky' && n.props.style.top === 0);
    assert.ok(head, '页签吸顶条要吸顶（position:sticky; top:0）');
    assert.ok(String(head.props.style.background).includes('bg-base'), '吸顶块要有底色，否则内容会透过去');
    // 列表**也不**自己滚：套两层滚动条会让「滚到一半卡住、要换个滚动条接着滚」。
    const lists = walk(tree).filter((n) => n.props && n.props['data-memory-list'] !== undefined);
    assert.ok(lists.length > 0, '列表还在');
    for (const list of lists) {
        assert.equal(list.props.style.maxHeight, undefined, '列表不该再有内层封顶高度');
        assert.equal(list.props.style.overflow, undefined, '列表不该再自己滚动');
    }
});

await check('指标卡：数字用等宽字、卡片走抬升表面与悬停态', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const values = byProp(tree, 'data-metric-value');
    assert.equal(values.length, 4);
    for (const value of values) {
        assert.match(String(value.props.style.fontFamily), /mono/, '指标数字要用等宽字（并排时才对得齐）');
        assert.deepEqual(value.props.children, ['1']);
    }
    for (const tile of byProp(tree, 'data-metric')) {
        assert.ok(String(tile.props.style.background).includes('bg-layer-1'), '指标卡要落在抬升表面上');
        assert.ok(String(tile.props.className).includes('om-card'), '指标卡要吃到悬停态');
    }
});

await check('仪表盘：统计区有总览标识与三列指标网格，指标卡仍是六枚', async () => {
    const { tree } = await panelFrom(makeSnapshot());
    const overview = byProp(tree, 'data-memory-overview')[0];
    assert.equal(overview.props['data-memory-dashboard'], 'true', '统计区要标成仪表盘');
    assert.ok(textsOf(overview).some((t) => t === 'MEMORY DASHBOARD'), '仪表盘有英文 kicker');
    assert.ok(textsOf(overview).some((t) => t === '内容优先 · 指标在后'), '仪表盘说清内容与指标的顺序');
    const metrics = byProp(tree, 'data-metrics', 'memory')[0];
    assert.match(String(metrics.props.style.gridTemplateColumns), /190px/, '指标网格至少 190px 起跳（标准宽度三列）');
    assert.equal(byProp(tree, 'data-metric').length, 4, '仪表盘保留四张指标卡（图谱下线后同步减少）');
});

await check('版式：内容优先 —— 列表在统计之前', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const nodes = walk(tree);
    const hotList = byProp(tree, 'data-memory-list', 'hot')[0];
    const overview = byProp(tree, 'data-memory-overview')[0];
    assert.ok(hotList, '热记忆列表应存在');
    assert.ok(overview, '统计区应存在');
    assert.ok(nodes.indexOf(hotList) < nodes.indexOf(overview),
        '热记忆列表必须排在指标卡 / 容量 / 存储域之前：打开面板先看到内容');
    assert.equal(byProp(overview, 'data-metric-value').length, 4, '统计区保留四张指标卡');
    assert.ok(textsOf(overview).some((t) => t === '统计与容量'), '统计区要有自己的标题');
});

await check('版式：稿头有主动作与分区标题，深色主题下主动作文字跟着反色', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const sticky = byProp(tree, 'data-memory-head', 'top')[0];
    assert.ok(sticky, '标题区应存在');
    const compose = byProp(sticky, 'data-memory-compose', 'true')[0];
    assert.ok(compose, '稿头要有「记一条」主动作');
    assert.equal(compose.props.style.color, 'var(--dsw-alias-bg-base, #fff)',
        '主动作文字要用底色反色：深色主题的强调色是浅色，写死白字会变成白底白字');
    assert.ok(textsOf(sticky).some((text) => text === 'OFFICE MEMORY'), '稿头要有品牌 kicker');

    const hotSection = byProp(tree, 'data-memory-section', 'preference')[0];
    const hotList = byProp(tree, 'data-memory-list', 'hot')[0];
    assert.ok(hotSection, '内容页签要有分区标题');
    assert.ok(walk(tree).indexOf(hotSection) < walk(tree).indexOf(hotList),
        '分区标题要排在列表之前');

    const kbTree = openTab(tree, panel, '知识库（1）');
    const kbSection = byProp(kbTree, 'data-memory-section', 'kb')[0];
    assert.ok(kbSection, '知识库文档清单也要有分区标题');
    assert.ok(textsOf(kbSection).some((text) => text.includes('篇')), '知识库分区标题要报文档篇数');
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
    // 条目数并没有消失：指标卡照旧按条目数报（归档页签已并入知识库，不再单独报数）。
    assert.equal(byProp(tree, 'data-metric-value').map((n) => n.props.children[0])[2], '30', '指标卡仍报条目数');
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

await check('主题令牌：用到的 --dsw-* 与宿主主题逐个核对（名单从源码现取，不再手写）', () => {
    const source = readFileSync(BUNDLE, 'utf8');
    // 宿主的主题令牌只定义 --dsw-alias-*；--dsh-border / --dsh-accent / --dsh-input-bg /
    // --dsh-danger 在 52 个 dsh-client-ui-* 包里一个都没有 —— 用它们等于永远走兜底。
    //
    // 但**会话版式**是另一套：dsh-client-ui-conversation 在 `.…_body` 上定义了
    // `--dsh-chat-content-width` / `--dsh-chat-user-width` / `--dsh-conversation-column-width`
    // （第七十一轮用它把面板宽度对齐「对话」正文列）。这三个是宿主里真实存在的，
    // 所以守卫从「一个 --dsh- 都不许有」精确成「除了宿主这三个布局变量，别的 --dsh- 都不许有」。
    const HOST_LAYOUT_VARS = new Set([
        '--dsh-chat-content-width',
        '--dsh-chat-user-width',
        '--dsh-conversation-column-width',
    ]);
    const hostLayoutVars = new Set([...source.matchAll(/(--dsh-[a-z0-9-]+)/g)].map((match) => match[1]));
    const unknownDsh = [...hostLayoutVars].filter((token) => !HOST_LAYOUT_VARS.has(token));
    assert.deepEqual(unknownDsh, [], `不能用宿主里不存在的 --dsh-* 变量：${unknownDsh.join('、')}`);
    assert.ok(hostLayoutVars.has('--dsh-chat-content-width'),
        '面板宽度要对齐宿主会话正文列（--dsh-chat-content-width）');
    // 兜底必须还在：令牌缺失时不能变成透明边框 / 无背景。
    assert.ok(source.includes('var(--dsw-alias-state-error-primary, #e5484d)'));
    assert.ok(source.includes('var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))'));

    // 12-3：名单从**源码里现取**，不再手写那五个。手写名单的代价是「加了一个新令牌，
    // 测试仍然全绿」——而界面用了宿主里不存在的令牌时，表现是永远走兜底（不报错、只是难看），
    // 恰恰是测试该拦的那类问题。这里反向做：**用到的每一个**都要在宿主主题的定义里找得到。
    //
    // 两种取法都要有：静态的 `var(--dsw-x, 兜底)` 直接取；**拼出来的**类型刻度
    // （`TYPE("s-14", …)` → `var(--dsw-font-s-14-font-size, …)`）要把后缀补全，
    // 只取静态那一半会漏掉整个字号刻度（那正是第二十轮刚换上的那一套）。
    const used = new Set();
    for (const match of source.matchAll(/var\(\s*(--dsw-[a-z0-9-]+)\s*[,)]/g)) used.add(match[1]);
    const typeScale = [...source.matchAll(/TYPE\(\s*"([a-z0-9-]+)"/g)].map((match) => match[1]);
    assert.ok(typeScale.length >= 5, `宿主类型刻度该被用到，实际只认出 ${typeScale.length} 个`);
    for (const token of typeScale) {
        for (const part of ['font-size', 'line-height', 'font-weight']) used.add(`--dsw-font-${token}-${part}`);
    }
    assert.ok(used.size >= 20, `至少该用到一批宿主令牌，实际 ${used.size}`);

    const themeManifest = resolveHostPath('@deepseek-ai/dsh-client-ui-theme/package.json');
    if (themeManifest === undefined) {
        console.log('      （本机解析不到宿主的主题包 dsh-client-ui-theme，跳过令牌存在性核对）');
        return;
    }
    const themeSource = readFileSync(join(dirname(themeManifest), 'lib', 'client.js'), 'utf8');
    const defined = new Set([...themeSource.matchAll(/--dsw-[a-z0-9-]+(?=\s*:)/g)].map((match) => match[0]));
    assert.ok(defined.size > 20, `宿主主题里该定义了一大批令牌，实际只有 ${defined.size} 个`);
    const missing = [...used].filter((token) => !defined.has(token));
    assert.deepEqual(missing, [], `这些令牌宿主的主题里没有定义（用了等于永远走兜底）：${missing.join('、')}`);
    return `核过 ${used.size} 个令牌（宿主定义 ${defined.size} 个）`;
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
            { kind: 'tool-result', seq: 7, call: { name: 'office_search_dispatch', argsRaw: '{}' } },
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

await check('写入指令：复制进剪贴板、「切到会话」优先走宿主给的 openView', async () => {
    const copied = [];
    const jumped = [];
    // 这一格现在是 conversation.view 的一格，所以宿主给的是 openView；本用例
    // 故意**不给** layout 服务，证明「切到会话」不再依赖它。
    const ownerProps = { openView: (view, focus) => { jumped.push([view, focus]); } };
    const { tree, panel } = await panelFrom(makeSnapshot(), {
        navigator: { clipboard: { writeText: (value) => { copied.push(value); return Promise.resolve(); } } },
        ownerProps,
    });

    byProp(tree, 'data-memory-action', 'replace')[0].props.onClick();
    const withPending = render(panel, ownerProps);
    assert.equal(byProp(withPending, 'data-memory-jump').length, 1, '拿得到 openView 时给「切到会话」');
    byProp(withPending, 'data-memory-jump')[0].props.onClick();
    assert.deepEqual(jumped, [['chat', '']], '切到会话 = openView("chat", "")（选回「对话」那一格）');

    byProp(withPending, 'data-memory-copy')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 5));
    assert.equal(copied.length, 1, '复制应把指令正文交给剪贴板');
    assert.ok(copied[0].includes('office_memory('), '复制的是指令正文');
    assert.ok(copied[0].includes('"replace"'));
    const afterCopy = render(panel, ownerProps);
    assert.ok(textsOf(afterCopy).some((t) => t.includes('已复制')), '复制成功后按钮改成「已复制」');
});

await check('写入指令：老宿主没有 openView 时，退回 layout.selectPanel("conversation")', async () => {
    const jumped = [];
    const { tree, panel } = await panelFrom(makeSnapshot(), {
        layout: { selectPanel: (id) => { jumped.push(id); } },
    });

    byProp(tree, 'data-memory-action', 'replace')[0].props.onClick();
    const withPending = render(panel);
    assert.equal(byProp(withPending, 'data-memory-jump').length, 1, '只有 layout 服务时也给「切到会话」');
    byProp(withPending, 'data-memory-jump')[0].props.onClick();
    assert.deepEqual(jumped, ['conversation'], '切到会话 = layout.selectPanel("conversation")');
});

await check('写入指令：拿不到剪贴板、openView 与 layout 时如实降级，不画死按钮', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    assert.equal(byProp(tree, 'data-memory-jump').length, 0, '两个入口都没有时不画「切到会话」');

    // 新增记忆不再走待发送指令：点「记一条」回到页内唯一的直接写入框。
    byProp(tree, 'data-memory-compose')[0].props.onClick();
    const withEditor = render(panel);
    assert.equal(byProp(withEditor, 'data-memory-editor', 'new').length, 0, '不再另开草稿卡');
    assert.ok(byProp(withEditor, 'data-memory-quick-text', 'true').length === 1, '页内直接写入框是唯一的新增入口');
    assert.equal(byProp(withEditor, 'data-memory-pending').length, 0, '不再生成 add 指令卡');

    // 降级路径仍由条目动作覆盖：拿不到剪贴板时提示手动复制，不画死按钮。
    byProp(tree, 'data-memory-action', 'replace')[0].props.onClick();
    const withPending = render(panel);
    assert.equal(byProp(withPending, 'data-memory-jump').length, 0, '没有跳转入口时不画按钮');
    byProp(withPending, 'data-memory-copy')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 5));
    assert.ok(textsOf(render(panel)).some((t) => t.includes('剪贴板权限')), '没有剪贴板时提示手动复制');
});

await check('直接记一条：填写文字即可写入，目标决定全局还是工作区（第八十九轮）', async () => {
    const posts = [];
    const snapshot = makeSnapshot();
    const fetchImpl = (url, init) => {
        if (init && init.method === 'POST') {
            posts.push({ url: String(url), body: JSON.parse(String(init.body)) });
            return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, action: 'memory-add', result: { id: 'm-9' } }) });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(snapshot) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);

    byProp(tree, 'data-memory-quick-open')[0].props.onClick();
    tree = render(panel);
    const text = byProp(tree, 'data-memory-quick-text')[0];
    assert.ok(text, '记忆页要有直接写入的输入框');
    text.props.onChange(eventOf({ value: '我不喜欢「不是 A 而是 B」这种句式。' }));
    tree = render(panel);
    byProp(tree, 'data-memory-quick-target-project')[0].props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-memory-quick-add')[0].props.disabled, false, '有内容时可点添加');
    byProp(tree, 'data-memory-quick-add')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 10));

    assert.equal(posts.length, 1, '应直接 POST 一次');
    assert.ok(posts[0].url.includes('mutate'), '走记忆写入端点');
    assert.equal(posts[0].body.action, 'memory-add', '动作是 memory-add');
    assert.equal(posts[0].body.content, '我不喜欢「不是 A 而是 B」这种句式。');
    assert.equal(posts[0].body.target, 'project', '切到项目记忆后 target=project');
    assert.equal(byProp(render(panel), 'data-memory-quick-text')[0].props.value, '', '写入成功后清空输入框');
});

await check('写入指令：台账有生成动作，知识库是只读浏览（图谱动作已随图谱下线）', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot());
    const ledgerTree = openTab(tree, panel, '台账（1）');
    assert.deepEqual(byProp(ledgerTree, 'data-memory-action').map((n) => n.props['data-memory-action']), ['log']);
    // 关系 / 实体两类的写入动作已经删除：render 出的条目上不该再出现 unlink / related 按钮。
    const allActions = byProp(tree, 'data-memory-action').map((n) => n.props['data-memory-action']);
    assert.ok(!allActions.includes('unlink') && !allActions.includes('related'), '旧图谱动作应已下线');
});

await check('分页：默认只铺一页，「再显示 N 条」才继续，全显示后给「收起」', async () => {
    const { tree, panel } = await panelFrom(makeSnapshot({
        counts: { hot: 25, hotUser: 25, hotProject: 0, ledger: 1, archive: 1, links: 1, entities: 1 },
        hot: manyHot(25).map((item) => Object.assign({}, item, { target: 'user' })),
    }));
    assert.equal(byProp(tree, 'data-item-key').length, 20, '默认只铺一页（20 条）');
    const more = byProp(tree, 'data-show-more')[0];
    assert.ok(more, '还有 5 条没显示时应给「再显示 N 条」');
    assert.deepEqual(more.props.children, ['再显示 5 条'], '按钮写清这一次会多显示几条');
    assert.ok(textsOf(tree).some((t) => t.includes('已显示 20 / 25')), '要写清已显示 / 总数');
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

// ── 新增：预览面板（conversation.view 第二格，order 50）────────────────────

const PREVIEW_CWD = 'D:/work/demo';

/** 清单响应的默认夹具：两份产物，与 src/preview-view.js 的字段一一对应。 */
function makePreviewFilesBody(overrides = {}) {
    return Object.assign({
        ok: true,
        root: PREVIEW_CWD,
        label: 'demo',
        total: 2,
        truncated: false,
        limit: 20,
        depth: 2,
        files: [
            { name: '季度汇报.pptx', path: '季度汇报.pptx', ext: '.pptx', kind: 'pptx', bytes: 24680, mtime: '2026-10-01T00:00:00.000Z' },
            { name: '备忘.docx', path: 'docs/备忘.docx', ext: '.docx', kind: 'renderable', bytes: 8192, mtime: '2026-09-30T00:00:00.000Z' },
        ],
    }, overrides);
}

const PREVIEW_SLIDES = [
    { page: 1, layout: 'title', title: '开篇' },
    { page: 2, layout: 'content', title: '数据' },
    { page: 3, layout: 'table', title: '预算' },
];

/** 大纲响应的默认夹具：按扩展名分三种口径（与宿主 preview-view.js 的返回形状一致）。 */
function makePreviewOutlineBody(filePath) {
    const value = String(filePath);
    if (value.endsWith('.pptx')) {
        return {
            ok: true, path: value, format: 'pptx', pages: 3, renderable: true,
            slides: PREVIEW_SLIDES,
            outline: ['P1 title：开篇', 'P2 bullets：数据｜4 个文本框', 'P3 table：预算'],
            layouts: [{ layout: 'title', count: 1 }, { layout: 'content', count: 1 }],
            warnings: ['第 2 页 标题过长可能折行', '第 3 页 表格列宽偏窄'],
            hint: '',
        };
    }
    if (value.endsWith('.docx')) {
        return { ok: true, path: value, format: 'docx', pages: null, renderable: true, slides: [], outline: [], layouts: [], warnings: [], hint: '这份格式没有逐页大纲，但可以按页取图看版式。' };
    }
    if (value.endsWith('.md') || value.endsWith('.csv') || value.endsWith('.tex')) {
        // 第七十一轮：文本类文件走 text:true 的「正文通道」，不再只做清单。
        return {
            ok: true, path: value, format: value.slice(value.lastIndexOf('.') + 1), pages: null, renderable: false,
            slides: [], outline: [], layouts: [], warnings: [],
            text: true,
            hint: '文本文件按原文预览（右列显示正文）',
        };
    }
    return {
        ok: true, path: value, format: 'md', pages: null, renderable: false, slides: [], outline: [], layouts: [], warnings: [],
        hint: '这种格式既没有页结构也不能渲染成图，预览面板里只做文件清单。',
    };
}

/** 文本正文响应的夹具：正文里故意混进表格、代码块、危险链接与裸 HTML。 */
function makePreviewTextBody(filePath, overrides = {}) {
    const text = [
        '# 记忆说明',
        '',
        '热记忆是 **常驻** 的，带 [链接](https://example.com/doc) 与 `code`。',
        '',
        '- 第一条',
        '- 第二条',
        '',
        '| 层 | 条数 |',
        '| --- | ---: |',
        '| 热 | 3 |',
        '',
        '> 引用一行',
        '',
        '```js',
        'const a = 1;',
        '```',
        '',
        '[点我](javascript:alert(1)) 与 <script>alert(1)</script> 与 ![图](x.png)',
    ].join('\n');
    return Object.assign({
        ok: true,
        root: PREVIEW_CWD,
        label: 'demo',
        path: String(filePath),
        name: String(filePath).replace(/^.*\//, ''),
        ext: '.md',
        kind: 'text',
        bytes: text.length,
        text,
        truncated: false,
        lines: text.split('\n').length,
        encoding: 'utf-8',
    }, overrides);
}

function previewResponse(status, body) {
    return Promise.resolve({ ok: status >= 200 && status < 300, status, json: () => Promise.resolve(body) });
}

/**
 * 三条预览端点的假实现：按 URL 路径分流，记录每次请求的完整 URL。
 *
 * 与宿主一样把「数据」和「HTTP 状态」分开：`__status` 只在测错误分支时出现，
 * 其余情况一律 200 + ok:true —— 面板照 body.ok 判断，不靠状态码猜。
 */
function makePreviewFetch(options = {}) {
    const calls = [];
    const filesBody = options.files === undefined ? makePreviewFilesBody() : options.files;
    const fetchImpl = (url) => {
        const href = String(url);
        calls.push(href);
        const cut = href.indexOf('?');
        const pathname = cut === -1 ? href : href.slice(0, cut);
        const query = new URLSearchParams(cut === -1 ? '' : href.slice(cut + 1));
        if (pathname === '/office-preview/files') {
            if (filesBody && filesBody.__status !== undefined) return previewResponse(filesBody.__status, filesBody.body);
            return previewResponse(200, filesBody);
        }
        if (pathname === '/office-preview/outline') {
            const body = options.outline === undefined
                ? makePreviewOutlineBody(query.get('path'))
                : options.outline(query.get('path'), query);
            if (body && body.__status !== undefined) return previewResponse(body.__status, body.body);
            return previewResponse(200, body);
        }
        if (pathname === '/office-preview/page') {
            if (options.pageError !== undefined) {
                return previewResponse(options.pageError.status, { ok: false, error: options.pageError.error });
            }
            const page = Number(query.get('page'));
            return previewResponse(200, {
                ok: true, path: query.get('path'), page, pages: 3, dpi: Number(query.get('dpi')),
                width: 960, height: 540, bytes: 2048, cached: page === 1,
                png: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUg==',
            });
        }
        if (pathname === '/office-preview/text') {
            if (options.textError !== undefined) {
                return previewResponse(options.textError.status, { ok: false, error: options.textError.error });
            }
            const body = options.text === undefined
                ? makePreviewTextBody(query.get('path'))
                : options.text(query.get('path'), query);
            if (body && body.__status !== undefined) return previewResponse(body.__status, body.body);
            return previewResponse(200, body);
        }
        return previewResponse(404, { ok: false, error: '没有这条预览路由：' + pathname });
    };
    fetchImpl.calls = calls;
    return fetchImpl;
}

const previewTick = () => new Promise((resolve) => setTimeout(resolve, 10));

/** 面板标准 owner props：会话作用域给的那两个（sessionId / useSessions）。 */
function previewOwnerProps() {
    return { sessionId: 's-1', useSessions: fakeUseSessions({ 's-1': { cwd: PREVIEW_CWD } }) };
}

async function previewPanelFrom(fetchImpl, ownerProps) {
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-preview');
    assert.ok(panel, 'conversation.view 应有 office-preview 这一格');
    const tree = await renderPanelLoaded(panel, ownerProps === undefined ? previewOwnerProps() : ownerProps);
    return { tree, panel };
}

await check('注册了预览面板：会话页签行里 order 50，紧挨记忆右侧', () => {
    const { mod, ctx, registrations } = loadBundle();
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-preview');
    assert.ok(panel, 'conversation.view 应有 office-preview 这一格');
    assert.equal(panel.opts.name, 'conversation.view');
    // list slot 只认这四个字段：多一个少一个都会被宿主忽略 / 报错。
    assert.deepEqual(Object.keys(panel.opts).sort(), ['id', 'label', 'name', 'order']);
    assert.equal(panel.opts.order, 50, '预览排在记忆（40）右侧、settings 之前，order 固定 50');
    const memory = regOf(registrations, 'conversation.view', 'office-memory');
    assert.ok(panel.opts.order > memory.opts.order, '必须排在「记忆」右侧');
    assert.equal(typeof panel.opts.label, 'function', 'label 要用 thunk（支持本地化重取）');
    assert.equal(panel.opts.label(), '预览');
    assert.equal(typeof panel.render, 'function');
    // 没有 props 时也不能抛错（宿主异常路径会这么渲染一次）。
    assert.doesNotThrow(() => render(panel));
    assert.doesNotThrow(() => render(panel, {}));
});

await check('预览面板：没有会话工作目录时显示空态，且一个请求都不发', async () => {
    const fetchImpl = makePreviewFetch();
    const { tree } = await previewPanelFrom(fetchImpl, {});
    const texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('这个会话还没有工作目录')), '应显示「没有工作目录」空态');
    assert.equal(byProp(tree, 'data-preview-empty', 'no-cwd').length, 1, '空态要带可断言的标记');
    // 关键：不能退回到「宿主随便挑一个根」——那会列出别人的文件。
    assert.equal(fetchImpl.calls.length, 0, '没有 cwd 时不该发任何预览请求');
});

await check('预览面板：挂载后按会话 cwd 列产物，自动选中最新一份并取大纲与第 1 页', async () => {
    const fetchImpl = makePreviewFetch();
    const { tree } = await previewPanelFrom(fetchImpl);

    assert.equal(fetchImpl.calls.length, 3, '刚好三次：清单 → 大纲 → 第 1 页');
    assert.ok(fetchImpl.calls[0].startsWith('/office-preview/files?cwd=' + encodeURIComponent(PREVIEW_CWD)),
        '第一次就带当前会话的 cwd：' + fetchImpl.calls[0]);
    assert.ok(fetchImpl.calls[1].startsWith('/office-preview/outline?'));
    assert.ok(fetchImpl.calls[1].includes('path=' + encodeURIComponent('季度汇报.pptx')), '默认选中清单里的第一份（mtime 最新）');
    assert.ok(fetchImpl.calls[2].startsWith('/office-preview/page?'));
    assert.ok(fetchImpl.calls[2].includes('page=1'), '自动渲染第 1 页');
    assert.ok(fetchImpl.calls[2].includes('dpi=110'), '默认 dpi 110');

    const texts = textsOf(tree);
    assert.ok(texts.some((t) => t.includes('预览 · demo')), '标题带工作目录短名');
    // 文件下拉：文件名 + 人读体积 + 选中最新那份。
    const fileSelect = byProp(tree, 'data-preview-file')[0];
    assert.ok(fileSelect, '应有文件下拉');
    assert.equal(fileSelect.props.value, '季度汇报.pptx');
    // 深色主题下列表弹层跟着 select 自己的 background-color 走；早先用的半透明
    // hover 底会把弹层变成灰玻璃（用户截图里的预览文件下拉）。这里守住不透明底
    // 与显式前景色，设置页共用的 select 也一并受益。
    assert.match(String(fileSelect.props.style.background), /--dsw-alias-bg-base/,
        '下拉要铺不透明的表面色');
    assert.match(String(fileSelect.props.style.color), /--dsw-alias-label-primary/,
        '下拉要显式给前景色，别靠 inherit 在弹层里失真');
    const options = walk(fileSelect).filter((node) => node.type === 'option').map((node) => node.props.value);
    assert.deepEqual(options, ['季度汇报.pptx', 'docs/备忘.docx'], '两份产物都在下拉里');
    assert.ok(texts.some((t) => t.includes('季度汇报.pptx · 24 KB')), '下拉项要带体积');
});

await check('预览面板：左列逐页大纲（页码 / 版式 / 标题 / 风险角标），右列大图', async () => {
    const { tree } = await previewPanelFrom(makePreviewFetch());
    const rows = byProp(tree, 'data-preview-page');
    assert.deepEqual(rows.map((node) => node.props['data-preview-page']), ['1', '2', '3'], '三页都在左列');
    const texts = textsOf(tree);
    for (const layout of ['title', 'content', 'table']) assert.ok(texts.includes(layout), '版式徽标：' + layout);
    for (const title of ['开篇', '数据', '预算']) assert.ok(texts.includes(title), '标题：' + title);
    // 风险角标只标被 warnings 点名的页（第 2、3 页），第 1 页没有。
    const riskyRows = rows.filter((node) => textsOf(node).includes('⚠')).map((node) => node.props['data-preview-page']);
    assert.deepEqual(riskyRows, ['2', '3'], '风险角标要按「第 N 页」落到对应行');

    const image = byProp(tree, 'data-preview-image')[0];
    assert.ok(image, '右列应渲染大图');
    assert.ok(String(image.props.src).startsWith('data:image/png;base64,'), '图直接是 data URL（不再发第二次请求）');
    assert.equal(image.props.alt, '第 1 页预览');
    assert.ok(texts.some((t) => t.includes('第 1 页 / 3')), '图上方的页码读数');
    assert.ok(texts.some((t) => t.includes('960×540')), '尺寸读数');
    assert.ok(texts.includes('缓存'), '首屏命中缓存要给标记');
});

await check('预览面板：点左侧页码取那一页（page 参数跟着走，dpi 不变）', async () => {
    const fetchImpl = makePreviewFetch();
    const { tree, panel } = await previewPanelFrom(fetchImpl);
    byProp(tree, 'data-preview-page', '3')[0].props.onClick();
    await previewTick();
    const next = render(panel);
    const last = fetchImpl.calls[fetchImpl.calls.length - 1];
    assert.ok(last.includes('page=3'), '请求最后一页：' + last);
    assert.ok(last.includes('dpi=110'), '只换页，不动 dpi');
    assert.equal(byProp(next, 'data-preview-image')[0].props.alt, '第 3 页预览');
});

await check('预览面板：换 dpi 重新取当前页（page 保留，dpi 带上新值）', async () => {
    const fetchImpl = makePreviewFetch();
    const { tree, panel } = await previewPanelFrom(fetchImpl);
    // 先翻到第 2 页，再换 200 dpi：换 dpi 不能把页码重置回 1。
    byProp(tree, 'data-preview-page', '2')[0].props.onClick();
    await previewTick();
    const afterPage = render(panel);
    const dpiSelect = byProp(afterPage, 'data-preview-dpi')[0];
    assert.ok(dpiSelect, '应有 dpi 下拉');
    dpiSelect.props.onChange(eventOf({ value: '200' }));
    await previewTick();
    const last = fetchImpl.calls[fetchImpl.calls.length - 1];
    assert.ok(last.includes('dpi=200'), 'dpi 要换成新值：' + last);
    assert.ok(last.includes('page=2'), '换 dpi 要留在当前页：' + last);
    assert.equal(byProp(render(panel), 'data-preview-dpi')[0].props.value, '200', '下拉回显新 dpi');
});

await check('预览面板：底部风险清单可点，跳到对应页', async () => {
    const fetchImpl = makePreviewFetch();
    const { tree, panel } = await previewPanelFrom(fetchImpl);
    const warnings = byProp(tree, 'data-preview-warning');
    assert.equal(warnings.length, 2, '两条风险各一行');
    assert.ok(textsOf(tree).some((t) => t.includes('排版风险 2 条')), '清单标题带条数');
    warnings[1].props.onClick(); // 第二条是「第 3 页 表格列宽偏窄」
    await previewTick();
    const last = fetchImpl.calls[fetchImpl.calls.length - 1];
    assert.ok(last.includes('page=3'), '点风险条目跳到它点名的页码：' + last);
    assert.equal(byProp(render(panel), 'data-preview-page', '3')[0].props['aria-pressed'], true, '左列选中态跟着走');
});

await check('预览面板：引擎不可用时右侧给服务端原文，大纲与风险清单照旧可用', async () => {
    const fetchImpl = makePreviewFetch({
        pageError: { status: 503, error: '渲染引擎不可用：这台机器上没有 LibreOffice。装好引擎后刷新这个面板即可（不用重启 DSH）。' },
    });
    const { tree } = await previewPanelFrom(fetchImpl);
    const notice = byProp(tree, 'data-preview-notice');
    assert.equal(notice.length, 1, '取图失败走右侧提示，不是整页错误');
    assert.ok(textsOf(notice[0]).some((t) => t.includes('渲染引擎不可用')), '原样展示服务端文案');
    // 清单 / 大纲 / 风险清单都不受影响：换页或装引擎后继续用，不该整页重来。
    assert.equal(byProp(tree, 'data-preview-pages').length, 1, '左列大纲仍在');
    assert.equal(byProp(tree, 'data-preview-warning').length, 2, '风险清单仍在');
    assert.equal(byProp(tree, 'data-preview-image').length, 0, '没有图就不画 img（不留空框）');
});

await check('预览面板：工作目录里没有产物时给空态，不画空的左右两列', async () => {
    const fetchImpl = makePreviewFetch({ files: { ok: true, root: PREVIEW_CWD, label: 'demo', total: 0, truncated: false, files: [] } });
    const { tree } = await previewPanelFrom(fetchImpl);
    assert.equal(byProp(tree, 'data-preview-empty', 'no-files').length, 1, '空态标记');
    assert.ok(textsOf(tree).some((t) => t.includes('这个工作目录里还没有办公产物')), '空态文案要给出下一步');
    assert.equal(byProp(tree, 'data-preview-viewer').length, 0, '空态时不该画空的浏览区');
    assert.equal(fetchImpl.calls.length, 1, '只请求清单，不再请求大纲 / 取图');
});

await check('预览面板：清单失败（不在白名单 / 404）时把服务端原文顶在顶部', async () => {
    const fetchImpl = makePreviewFetch({
        files: { __status: 404, body: { ok: false, error: '这个工作目录不在已知列表里：加载页面时还没见过它。先在对应目录里跑一次办公工具，或刷新后再试。' } },
    });
    const { tree } = await previewPanelFrom(fetchImpl);
    assert.equal(byProp(tree, 'data-preview-empty', 'error').length, 1, '整页错误标记');
    assert.ok(textsOf(tree).some((t) => t.includes('不在已知列表里')), '服务端已经写好下一步，面板不该换成「加载失败」');
});

await check('预览面板：非 pptx 的 docx 只有「第 1 页」一行，照样能取图', async () => {
    const docxFiles = makePreviewFilesBody({ total: 1, files: [
        { name: '备忘.docx', path: 'docs/备忘.docx', ext: '.docx', kind: 'renderable', bytes: 8192, mtime: '2026-09-30T00:00:00.000Z' },
    ] });
    const fetchImpl = makePreviewFetch({ files: docxFiles });
    const { tree } = await previewPanelFrom(fetchImpl);
    const rows = byProp(tree, 'data-preview-page');
    assert.deepEqual(rows.map((node) => node.props['data-preview-page']), ['1'], '没有章节也要有第 1 页这一行');
    const texts = textsOf(tree);
    assert.ok(texts.includes('页'), 'docx 没有版式，徽标退成「页」');
    assert.ok(texts.includes('（无标题）'));
    assert.ok(fetchImpl.calls.some((call) => call.startsWith('/office-preview/page?') && call.includes('page=1')), 'docx 也能取第 1 页');
});

// ── 第七十一轮：Markdown 在预览面板里能读 ──────────────────────────────────
//
// 第六十八轮只做了清单 + 逐页大纲，于是用户在预览里选中 .md 时右列只有一句
// 「只做文件清单」（正是用户报的「预览无法看 markdown 文件」）。现在服务端在
// outline 里给 text:true，面板改走 /office-preview/text 把正文画出来。

const MD_FILES = makePreviewFilesBody({ total: 1, files: [
    { name: '记忆.md', path: 'docs/记忆.md', ext: '.md', kind: 'text', bytes: 2048, mtime: '2026-09-30T00:00:00.000Z' },
] });

/**
 * 递归收集一棵假元素树里的全部字符串。
 *
 * 不能用 `textsOf`：标记渲染器把行内令牌作为**数组**塞进 children（一个段落的
 * children = [mdInline(...) 的结果]），`textsOf` 只 flat 一层，段落文字会漏掉。
 */
function deepTexts(node) {
    const acc = [];
    const visit = (value) => {
        if (typeof value === 'string') acc.push(value);
        else if (Array.isArray(value)) value.forEach(visit);
        else if (value && value.props) visit(value.props.children);
    };
    visit(node);
    return acc;
}

await check('预览面板：选中 .md 走正文通道，按 Markdown 画出来（不再只有一句 hint）', async () => {
    const fetchImpl = makePreviewFetch({ files: MD_FILES });
    const { tree } = await previewPanelFrom(fetchImpl);

    const outlineCall = fetchImpl.calls.find((call) => call.startsWith('/office-preview/outline'));
    assert.ok(outlineCall.includes(encodeURIComponent('docs/记忆.md')), '先取大纲');
    const textCall = fetchImpl.calls.find((call) => call.startsWith('/office-preview/text'));
    assert.ok(textCall, '文本类文件必须请求正文：' + JSON.stringify(fetchImpl.calls));
    assert.ok(textCall.includes('cwd=' + encodeURIComponent(PREVIEW_CWD)), '正文请求带 cwd');
    assert.ok(textCall.includes('path=' + encodeURIComponent('docs/记忆.md')), '正文请求带路径');
    assert.equal(fetchImpl.calls.filter((call) => call.startsWith('/office-preview/page')).length, 0, '文本不取页面图');

    const md = byProp(tree, 'data-preview-text', 'markdown');
    assert.equal(md.length, 1, '右列要画 Markdown 正文');
    assert.equal(byProp(tree, 'data-preview-viewer', 'true').length, 1, '正文仍在右列 viewer 里');
    const scope = walk(md[0]);
    const texts = deepTexts(md[0]);
    assert.ok(texts.some((t) => t === '记忆说明'), '一级标题要渲染');
    assert.ok(texts.some((t) => t.includes('热记忆是')), '段落要渲染');
    assert.equal(scope.filter((node) => node.type === 'li').length, 2, '两个列表项');
    assert.equal(scope.filter((node) => node.type === 'table').length, 1, '表格要真的画成 table');
    assert.equal(scope.filter((node) => node.type === 'pre').length, 1, '围栏代码块');
    assert.equal(scope.filter((node) => node.type === 'a' && node.props.href === 'https://example.com/doc').length, 1, '安全链接画成 <a>');
    assert.equal(scope.filter((node) => node.type === 'a' && String(node.props.href).indexOf('javascript:') === 0).length, 0,
        'javascript: 链接必须降级成纯文本（面板不注入 HTML，也不给危险链接可点）');
    assert.ok(texts.some((t) => t.includes('<script>')), '裸 HTML 按字面文本显示，不注入');
    assert.equal(byProp(tree, 'data-preview-text-meta').length, 1, '正文头要有文件名/体积/行数');
});

await check('预览面板：非 .md 的文本文件按等宽原文显示', async () => {
    const csvFiles = makePreviewFilesBody({ total: 1, files: [
        { name: '台账.csv', path: '台账.csv', ext: '.csv', kind: 'text', bytes: 64, mtime: '2026-09-30T00:00:00.000Z' },
    ] });
    const fetchImpl = makePreviewFetch({
        files: csvFiles,
        text: (filePath) => makePreviewTextBody(filePath, { name: '台账.csv', ext: '.csv', text: 'a,b\n1,2\n' }),
    });
    const { tree } = await previewPanelFrom(fetchImpl);
    const plain = byProp(tree, 'data-preview-text', 'plain');
    assert.equal(plain.length, 1, '.csv 不按 Markdown 解析');
    assert.equal(byProp(tree, 'data-preview-text', 'markdown').length, 0);
    assert.equal(plain[0].type, 'pre');
    assert.equal(deepTexts(plain[0]).join(''), 'a,b\n1,2\n', '原文一个字符都不改');});

await check('预览面板：正文取不到时给可读错误与下一步，清单照旧可用', async () => {
    const fetchImpl = makePreviewFetch({
        files: MD_FILES,
        textError: { status: 500, error: '读文件失败：EBUSY。' },
    });
    const { tree } = await previewPanelFrom(fetchImpl);
    const err = byProp(tree, 'data-preview-text-error');
    assert.equal(err.length, 1, '正文失败要落在右列，而不是整页白屏');
    const texts = textsOf(err[0]);
    assert.ok(texts.some((t) => t.includes('EBUSY')), '错误原文照显示');
    assert.ok(texts.some((t) => t.includes('下一步')), '附上下一步');
    assert.equal(byProp(tree, 'data-preview-file').length, 1, '文件下拉还在，用户能换一份');
});

await check('预览面板：正文被截断时如实提示（大文件不只读一半还不说）', async () => {
    const fetchImpl = makePreviewFetch({
        files: MD_FILES,
        text: (filePath) => makePreviewTextBody(filePath, {
            truncated: true,
            hint: '文件较大，只预览前 2 MiB（共 12.0 MiB）',
        }),
    });
    const { tree } = await previewPanelFrom(fetchImpl);
    const notice = byProp(tree, 'data-preview-text-truncated');
    assert.equal(notice.length, 1);
    assert.ok(textsOf(notice[0]).some((t) => t.includes('只预览前 2 MiB')), '提示必须是服务端那句');
    assert.equal(byProp(tree, 'data-preview-text-meta').length, 1);
});

await check('预览面板：刷新按钮重新拉清单，并保持当前选择', async () => {
    const fetchImpl = makePreviewFetch();
    const { tree, panel } = await previewPanelFrom(fetchImpl);
    assert.equal(fetchImpl.calls.filter((call) => call.startsWith('/office-preview/files')).length, 1);
    byProp(tree, 'data-preview-refresh')[0].props.onClick();
    await previewTick();
    const after = render(panel);
    assert.equal(fetchImpl.calls.filter((call) => call.startsWith('/office-preview/files')).length, 2, '刷新要重新列清单');
    assert.equal(byProp(after, 'data-preview-file')[0].props.value, '季度汇报.pptx', '刷新后选择不丢');
    assert.ok(fetchImpl.calls.some((call) => call.includes('/office-preview/outline?') && call.includes(encodeURIComponent('季度汇报.pptx'))), '刷新后重新取大纲');
});

await check('预览面板：产物超过 20 份时提示被截断，并给出总数', async () => {
    const many = [];
    for (let index = 0; index < 20; index += 1) {
        many.push({ name: 'deck-' + index + '.pptx', path: 'deck-' + index + '.pptx', ext: '.pptx', kind: 'pptx', bytes: 1024, mtime: 'x' });
    }
    const fetchImpl = makePreviewFetch({ files: makePreviewFilesBody({ files: many, total: 25, truncated: true }) });
    const { tree } = await previewPanelFrom(fetchImpl);
    const hint = byProp(tree, 'data-preview-truncated');
    assert.equal(hint.length, 1, '截断要有提示条');
    assert.ok(textsOf(hint[0]).some((t) => t.includes('共 25 份')), '提示里要有总数：' + textsOf(hint[0]).join('|'));
    assert.equal(walk(byProp(tree, 'data-preview-file')[0]).filter((node) => node.type === 'option').length, 20, '下拉只放 20 份');
});

await check('预览面板：换文件下拉会切到那份文件（大纲 + 第 1 页一起重取）', async () => {
    const fetchImpl = makePreviewFetch();
    const { tree, panel } = await previewPanelFrom(fetchImpl);
    byProp(tree, 'data-preview-file')[0].props.onChange(eventOf({ value: 'docs/备忘.docx' }));
    await previewTick();
    const next = render(panel);
    assert.equal(byProp(next, 'data-preview-file')[0].props.value, 'docs/备忘.docx', '下拉选中跟着变');
    const outlineCalls = fetchImpl.calls.filter((call) => call.startsWith('/office-preview/outline'));
    assert.ok(outlineCalls[outlineCalls.length - 1].includes(encodeURIComponent('docs/备忘.docx')), '重新取这份的大纲');
    assert.ok(textsOf(next).some((t) => t.includes('这份格式没有逐页大纲')), 'docx 的 hint 照原样显示');
});

// ── 第七十一轮：记忆导出入口 ──────────────────────────────────────────────
//
// 用户要「方便迁移」：agent 记忆没有公认的公有协议，所以导出成**单文件**。
// 面板侧只需要两个原生下载锚：JSON（无损，能 importPack 导回）与 MD（人读）。

await check('记忆面板：稿头有「导出」与「MD」两个下载入口，链接带当前工作区与格式', async () => {
    const snapshot = makeSnapshot();
    const { tree } = await panelFrom(snapshot);
    const json = byProp(tree, 'data-memory-export', 'json');
    const md = byProp(tree, 'data-memory-export', 'md');
    assert.equal(json.length, 1, '要有 JSON 导出入口');
    assert.equal(md.length, 1, '要有 Markdown 导出入口');
    assert.equal(json[0].type, 'a', '用原生 <a download>，不在面板里 fetch + Blob');
    assert.ok('download' in json[0].props, '必须带 download 属性');
    const href = String(json[0].props.href);
    assert.ok(href.startsWith('/office-memory/export?'), '直连导出端点：' + href);
    assert.ok(href.includes('format=json'));
    assert.ok(href.includes('cwd=' + encodeURIComponent(snapshot.cwd)), '导出的就是面板正在看的这个工作区');
    assert.ok(String(md[0].props.href).includes('format=md'), 'MD 入口换成 format=md');
});

// ── 第九十一轮：知识库资源管理器 / 文件识别（OCR） ──────────────────────────

await check('知识库：清单是资源管理器的「文件夹 + 文件」两段，不再只是卡片墙', async () => {
    const urls = [];
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    tree = openTab(tree, panel, '知识库（1）');
    assert.equal(byProp(tree, 'data-kb-file-head', 'true').length, 1, '要有表头（名称 / 档位 / 块 / 大小 / 修改 / 操作）');
    const sub = byProp(tree, 'data-kb-subfolder', '来源')[0];
    assert.ok(sub, '当前层的子文件夹要单独成行');
    assert.ok(textsOf(sub).some((t) => t.includes('来源')), '子文件夹行要写文件夹名');
    assert.ok(textsOf(sub).some((t) => t === '1 篇'), '子文件夹行要写文档数');
    const fileRow = byProp(tree, 'data-kb-file', 'kb:8f14e45f')[0];
    assert.ok(fileRow, '文件行要按文档 id 落 data-kb-file');
    assert.ok(textsOf(fileRow).some((t) => t.includes('来源/方法论.md')), '文件行要显示入库路径');
    assert.ok(textsOf(fileRow).some((t) => t === 'verified'), '文件行要显示来源档');
    assert.ok(byProp(fileRow, 'data-kb-open', 'kb:8f14e45f')[0], '文件行保留「打开文档」入口');
    assert.ok(byProp(fileRow, 'data-memory-edit', 'kb:8f14e45f')[0], '文件行保留编辑入口');
    assert.ok(byProp(fileRow, 'data-kb-delete', 'kb:8f14e45f')[0], '文件行保留删除入口');
    // 点文件夹行进入下一层：请求里要带 kbfolder=来源
    sub.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 10));
    assert.ok(urls[urls.length - 1].includes('kbfolder=' + encodeURIComponent('来源')), '点文件夹行要带 kbfolder=');
});
await check('知识库：大图标视图可按 Enter 打开文档（第一百一十一轮补）', async () => {
    const urls = [];
    const fetchImpl = (url) => {
        urls.push(String(url));
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    tree = openTab(tree, panel, '知识库（1）');
    byProp(tree, 'data-kb-view', 'tiles')[0].props.onClick();
    tree = render(panel);
    const tile = byProp(tree, 'data-kb-file', 'kb:8f14e45f')[0];
    assert.ok(tile, '大图标视图要有文档磁贴');
    assert.equal(typeof tile.props.onKeyDown, 'function', '磁贴要能用键盘打开');
    tile.props.onKeyDown({ key: 'Enter', preventDefault() {}, target: {} });
    await new Promise((resolve) => setTimeout(resolve, 20));
    const last = urls[urls.length - 1];
    assert.ok(last.includes('kbdoc=' + encodeURIComponent('kb:8f14e45f')), 'Enter 要打开这份文档');
});

await check('知识库：文档搜索为空时给「清空搜索」，不让人只靠顶部小叉（第一百一十一轮补）', async () => {
    const urls = [];
    const empty = makeSnapshot({
        kb: [],
        kbTree: { path: '', name: '全部文档', docs: 0, chunks: 0, bytes: 0, children: [] },
        counts: Object.assign({}, makeSnapshot().counts, { kb: 0, kbTotal: 0, kbChunks: 0, kbBytes: 0 }),
    });
    const fetchImpl = (url) => {
        const text = String(url);
        urls.push(text);
        const q = new URL(text, 'http://localhost').searchParams.get('q');
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(q ? empty : makeSnapshot()) });
    };
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    await new Promise((resolve) => setTimeout(resolve, 10));
    tree = openTab(tree, panel, '知识库（1）');
    const top = walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找文档'));
    top.props.onChange({ target: { value: '量子纠缠' } });
    tree = render(panel);
    walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找文档'))
        .props.onKeyDown({ key: 'Enter', target: { value: '量子纠缠' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const emptyState = byProp(tree, 'data-kb-empty', 'true')[0];
    assert.ok(emptyState, '没有匹配文档时要显示空态');
    assert.ok(textsOf(emptyState).some((t) => t.includes('没有匹配的知识库文档')), '空态要说清是搜索无结果');
    const action = byProp(tree, 'data-kb-empty-action', 'clear')[0];
    assert.ok(action, '空态要给「清空搜索」');
    action.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(new URL(urls[urls.length - 1], 'http://localhost').searchParams.has('q'), false, '清空后不再带 q');
    assert.ok(byProp(tree, 'data-kb-file', 'kb:8f14e45f').length >= 1, '清空后要回到完整文档清单');
});

// 统一入口（第九十五轮）：选完文件只预览；扫描件必须**显式**点「识别文字」才发给 OCR。
function kbFileReaderFake() {
    const previous = globalThis.FileReader;
    class FakeReader {
        readAsDataURL(file) {
            const encoded = Buffer.from(String(file && file.__body || '')).toString('base64');
            this.result = 'data:application/octet-stream;base64,' + encoded;
            Promise.resolve().then(() => { if (typeof this.onload === 'function') this.onload(); });
        }
    }
    globalThis.FileReader = FakeReader;
    return () => { globalThis.FileReader = previous; };
}

await check('知识库：添加文件先预览、不自动识别；扫描件点「识别文字」才发 OCR（发内容不发客户端路径）', async () => {
    const calls = [];
    const fetchImpl = (url, init) => {
        const text = String(url);
        calls.push({ url: text, init: init || {} });
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            assert.equal(body.preview, true, '选完文件第一步只能是预览');
            return jsonResponse({
                ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'scan', name: '扫描件.png', ext: '.png', bytes: 2048,
                    chars: 0, chunks: 1, head: '', truncated: false, images: 1,
                    note: '这份文件没有可抽取的文本层（是扫描件）：要读文字请点「识别文字」，内容会发给 DeepSeek OCR。',
                },
            });
        }
        if (text.includes('/office-memory/ocr')) {
            const payload = {
                ok: true, fileName: '扫描件.png', pageCount: 1, pages: [{ index: 0, text: '识别出来的正文', finishReason: 'stop' }],
                usage: { total_tokens: 12 }, bytes: 2048, ms: 880, text: '识别出来的正文',
            };
            // 客户端先读文本、再自己解析（空/HTML/非 JSON 都要分类），假响应也照真实响应给出 text()。
            return Promise.resolve({
                ok: true, status: 200, headers: { get: () => 'application/json; charset=utf-8' },
                text: () => Promise.resolve(JSON.stringify(payload)),
            });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const restore = kbFileReaderFake();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '扫描件.png', __body: 'png-bytes' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);

        assert.equal(calls.filter((call) => call.url.includes('/office-memory/ocr')).length, 0, '预览阶段绝不能自动发 OCR');
        const card = byProp(tree, 'data-kb-preview-item', '扫描件.png')[0];
        assert.ok(card, '预览卡要出现');
        assert.equal(card.props['data-kb-preview-kind'], 'scan', '扫描件要标成 scan');
        assert.ok(textsOf(tree).some((t) => t.includes('没有可抽取的文本层')), '预览要说明为什么需要识别');
        assert.equal(byProp(tree, 'data-kb-preview-ingest', '扫描件.png').length, 0, '扫描件没有文本，不给「入库」误导');
        const ocrButton = byProp(tree, 'data-kb-preview-ocr', '扫描件.png')[0];
        assert.ok(ocrButton, '扫描件要给显式的「识别文字」按钮');

        ocrButton.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        const posted = calls.filter((call) => call.url.includes('/office-memory/ocr'));
        assert.equal(posted.length, 1, '点了「识别文字」才发一次 OCR 请求');
        assert.equal(posted[0].init.method, 'POST');
        const body = JSON.parse(posted[0].init.body);
        assert.deepEqual(body, { name: '扫描件.png', data: Buffer.from('png-bytes').toString('base64') },
            'OCR 要发内容与文件名，不发客户端路径');
        assert.equal('path' in body, false, '客户端路径不能当服务端路径发出去');
        assert.equal(byProp(tree, 'data-kb-ocr-panel', 'true').length, 1, '结果要就地展开一个面板');
        assert.equal(byProp(tree, 'data-kb-ocr-text', 'true')[0].props.value, '识别出来的正文');
        assert.equal(byProp(tree, 'data-kb-ocr-save', 'true').length, 1, '结果面板要能「存为知识库笔记」');
    } finally {
        restore();
    }
});

await check('知识库：预览不自动入库，文本给出字数与块数；损坏 PDF 与扫描件分开提示', async () => {
    const uploads = [];
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            uploads.push(body);
            if (body.preview !== true) {
                return jsonResponse({
                    ok: true, action: 'kb-upload', result: { action: 'kb-ingest', path: 'uploads/笔记.txt', chars: 1200, chunks: 3 },
                });
            }
            if (body.name === '坏.pdf') {
                return jsonResponse({
                    ok: true, action: 'kb-preview', result: {
                        action: 'kb-preview', kind: 'broken', name: '坏.pdf', ext: '.pdf', bytes: 12,
                        chars: 0, chunks: 0, head: '', truncated: false, images: 0,
                        note: '这个文件读不出来（可能已损坏或加了密码）：换一份副本再试。',
                    },
                });
            }
            return jsonResponse({
                ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 3600,
                    chars: 1200, chunks: 3, head: '第一段正文……', truncated: true, images: 0, note: '',
                },
            });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const restore = kbFileReaderFake();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [
            { name: '笔记.txt', __body: '第一段正文' },
            { name: '坏.pdf', __body: 'not-a-pdf' },
        ] } });
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);

        assert.equal(uploads.length, 2, '两份文件各预览一次');
        assert.ok(uploads.every((one) => one.preview === true), '预览请求不能夹带入库');
        assert.equal(uploads.filter((one) => one.ai === true).length, 0, '预览绝不调用模型');
        const textCard = byProp(tree, 'data-kb-preview-item', '笔记.txt')[0];
        assert.ok(textCard, '文本预览卡要在');
        assert.ok(textsOf(textCard).some((t) => t.includes('1200 字') && t.includes('3 块')), '预览要报字数与块数');
        assert.ok(textsOf(textCard).some((t) => t.includes('第一段正文')), '预览要给开头一段正文');
        assert.equal(byProp(textCard, 'data-kb-preview-ocr', '笔记.txt').length, 0, '文本不需要 OCR 按钮');
        const brokenCard = byProp(tree, 'data-kb-preview-item', '坏.pdf')[0];
        assert.ok(brokenCard, '损坏文件的预览卡要在');
        assert.equal(brokenCard.props['data-kb-preview-kind'], 'broken', '损坏 PDF 要与扫描件分开');
        assert.ok(textsOf(brokenCard).some((t) => /损坏|密码/.test(t)), '损坏文件要直说原因');
        assert.equal(byProp(brokenCard, 'data-kb-preview-ocr', '坏.pdf').length, 0, '损坏文件不给 OCR 按钮');
        assert.equal(byProp(brokenCard, 'data-kb-preview-ingest', '坏.pdf').length, 0, '损坏文件不给入库按钮');

        const ingestButton = byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0];
        assert.ok(ingestButton, '文本预览卡要有显式「入库」');
        ingestButton.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);
        const ingests = uploads.filter((one) => one.preview !== true);
        assert.equal(ingests.length, 1, '只有点了「入库」才真的写库');
        assert.equal(ingests[0].name, '笔记.txt');
        assert.equal(ingests[0].ai, false, '未勾 AI 时入库走确定性切块');
        assert.equal(byProp(tree, 'data-kb-preview-item', '笔记.txt').length, 0, '入库后这份预览卡要收掉');
        assert.ok(textsOf(tree).some((t) => t.includes('已入库：笔记.txt')), '入库要留一行回执');
    } finally {
        restore();
    }
});

// ── 统一入口的本地预检（第九十六轮）：8 MB 是抽取上限，不是 OCR 上限 ──────
//
// 浏览器端选完文件先看类型/体积：8–24 MB 且 OCR 可识别的文件**不读、不发预览**，
// 只在卡上如实说「没有判定文本层」；用户显式点「识别文字」时才读取并发送。

function countingFileReader() {
    const reads = [];
    const previous = globalThis.FileReader;
    class CountingReader {
        readAsDataURL(file) {
            reads.push(String(file && file.name || ""));
            const encoded = Buffer.from(String(file && file.__body || '')).toString('base64');
            this.result = 'data:application/octet-stream;base64,' + encoded;
            Promise.resolve().then(() => { if (typeof this.onload === 'function') this.onload(); });
        }
    }
    globalThis.FileReader = CountingReader;
    return { reads, restore: () => { globalThis.FileReader = previous; } };
}

await check('知识库：8–24 MB 的扫描件先本地预检——不读、不发；点「识别文字」才读并发 OCR', async () => {
    const calls = [];
    const fetchImpl = (url, init) => {
        const text = String(url);
        calls.push({ url: text, init: init || {} });
        if (text.includes('/office-memory/ocr')) {
            const payload = { ok: true, fileName: '大扫描.png', pageCount: 1, text: '识别出来的正文', usage: null, bytes: 12, ms: 9 };
            return Promise.resolve({
                ok: true, status: 200, headers: { get: () => 'application/json; charset=utf-8' },
                text: () => Promise.resolve(JSON.stringify(payload)),
            });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const reader = countingFileReader();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [
            { name: '大扫描.png', size: 12 * 1024 * 1024, __body: 'png-bytes' },
        ] } });
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);

        assert.equal(calls.filter((call) => call.url.includes('/office-memory/kb-upload')).length, 0,
            '8 MB 以上不发送预览请求（服务端也抽不了）');
        assert.equal(reader.reads.length, 0, '预检阶段不读文件内容');
        const card = byProp(tree, 'data-kb-preview-item', '大扫描.png')[0];
        assert.ok(card, '本地预检也要给出预览卡');
        assert.equal(card.props['data-kb-preview-kind'], 'unknown', '没抽取就不能断言是扫描件');
        assert.ok(textsOf(card).some((t) => t.includes('没有判定文本层') && t.includes('8 MB')), '要说清为什么没判定');
        assert.equal(byProp(card, 'data-kb-preview-ingest', '大扫描.png').length, 0, '没有正文就不给「入库」');
        const ocrButton = byProp(card, 'data-kb-preview-ocr', '大扫描.png')[0];
        assert.ok(ocrButton, '仍要保留显式的「识别文字」通道');

        ocrButton.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);
        assert.equal(reader.reads.length, 1, '点「识别文字」才读文件');
        const ocr = calls.filter((call) => call.url.includes('/office-memory/ocr'));
        assert.equal(ocr.length, 1, '只发一次 OCR');
        const body = JSON.parse(ocr[0].init.body);
        assert.equal(body.name, '大扫描.png');
        assert.equal(body.data, Buffer.from('png-bytes').toString('base64'));
        assert.equal('path' in body, false, '仍不发客户端路径');
    } finally {
        reader.restore();
    }
});


await check('知识库：8–24 MB 的 PDF 预览自动拆分，每份一张待入库卡（第一百一十轮）', async () => {
    const uploads = [];
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            uploads.push(body);
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'split', name: '大文档.pdf', ext: 'pdf', bytes: 12 * 1024 * 1024,
                    ocrEligible: true, note: '文件有 12582912 字节，已按页拆成 2 份（每份不超过 4MB 字节）；逐份确认入库。',
                    parts: [
                        {
                            name: '大文档（第1页）.pdf', kind: 'text', pageFrom: 1, pageTo: 1, pages: 1,
                            bytes: 1024, chars: 120, chunks: 1, head: '第一页正文', truncated: false, images: 0,
                            data: Buffer.from('pdf-one').toString('base64'), ocrEligible: false, note: '',
                        },
                        {
                            name: '大文档（第2页）.pdf', kind: 'text', pageFrom: 2, pageTo: 2, pages: 1,
                            bytes: 2048, chars: 240, chunks: 2, head: '第二页正文', truncated: false, images: 0,
                            data: Buffer.from('pdf-two').toString('base64'), ocrEligible: false, note: '',
                        },
                    ],
                } });
            }
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, chars: 120, chunks: 1, ai: { mode: 'off' },
            } });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const reader = countingFileReader();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [
            { name: '大文档.pdf', size: 12 * 1024 * 1024, __body: 'pdf-bytes' },
        ] } });
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);

        assert.equal(uploads.length, 1, '大 PDF 要像普通文件一样发一次预览请求');
        assert.equal(uploads[0].preview, true);
        assert.equal(reader.reads.length, 1, '大 PDF 预览必须读一次内容');
        assert.equal(byProp(tree, 'data-kb-preview-item', '大文档（第1页）.pdf').length, 1, '第一份拆分卡要在');
        assert.equal(byProp(tree, 'data-kb-preview-item', '大文档（第2页）.pdf').length, 1, '第二份拆分卡要在');
        const first = byProp(tree, 'data-kb-preview-item', '大文档（第1页）.pdf')[0];
        assert.equal(first.props['data-kb-preview-kind'], 'text', '拆分卡按实际抽取类型显示');
        assert.ok(textsOf(first).some((t) => t.includes('120 字') && t.includes('1 块')), '每份要独立报字数与块数');
        assert.ok(byProp(first, 'data-kb-preview-ingest', '大文档（第1页）.pdf')[0], '每份都要有独立的入库按钮');

        byProp(tree, 'data-kb-preview-ingest', '大文档（第1页）.pdf')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);
        assert.equal(uploads.length, 2, '点哪一份才入库哪一份');
        assert.equal(uploads[1].name, '大文档（第1页）.pdf');
        assert.equal(uploads[1].preview, undefined, '第二次不是预览');
        assert.equal(uploads[1].data, Buffer.from('pdf-one').toString('base64'), '发拆分后的那一份内容');
        assert.equal(byProp(tree, 'data-kb-preview-item', '大文档（第1页）.pdf').length, 0, '入库后的那份卡要收掉');
        assert.equal(byProp(tree, 'data-kb-preview-item', '大文档（第2页）.pdf').length, 1, '没点的那份要留下');
    } finally {
        reader.restore();
    }
});

await check('知识库：超过 24 MB 的文件在预检阶段直接拒绝，不读不发', async () => {
    const calls = [];
    const reader = countingFileReader();
    try {
        const { mod, ctx, registrations } = loadBundle({
            fetch: (url) => {
                calls.push(String(url));
                return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
            },
        });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [
            { name: '超大.png', size: 25 * 1024 * 1024, __body: 'png-bytes' },
        ] } });
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);
        assert.equal(reader.reads.length, 0, '拒绝也要发生在读取之前');
        assert.equal(calls.filter((url) => url.includes('kb-upload') || url.includes('/office-memory/ocr')).length, 0,
            '超限文件不能被送出去');
        assert.ok(textsOf(tree).some((t) => t.includes('24 MB')), '要报清上限');
        assert.equal(byProp(tree, 'data-kb-preview-item', '超大.png').length, 0, '超限文件不生成预览卡');
    } finally {
        reader.restore();
    }
});
await check('知识库：2–8 MB 的文本预览就拒绝，不给「入库」（服务端单文档上限 2 MB）', async () => {
    const calls = [];
    const reader = countingFileReader();
    try {
        const { mod, ctx, registrations } = loadBundle({
            fetch: (url) => {
                calls.push(String(url));
                return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
            },
        });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [
            { name: '中等.txt', size: 3 * 1024 * 1024, __body: 'xxxx' },
        ] } });
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);
        assert.equal(reader.reads.length, 0, '2 MB 以上文本不读');
        assert.equal(calls.filter((url) => url.includes('kb-upload')).length, 0, '也不发预览请求');
        const card = byProp(tree, 'data-kb-preview-item', '中等.txt')[0];
        assert.ok(card, '要给本地拒绝卡');
        assert.equal(card.props['data-kb-preview-kind'], 'too_large');
        assert.ok(textsOf(card).some((t) => t.includes('2 MB')), '要报清 2 MB 单文档上限');
        assert.equal(byProp(card, 'data-kb-preview-ingest', '中等.txt').length, 0, '不能给点了才失败的入库按钮');
        assert.equal(byProp(card, 'data-kb-preview-ocr', '中等.txt').length, 0, '文本没有识别通道');
    } finally {
        reader.restore();
    }
});

await check('知识库：图片预览后给显式「入库」，只发内容不发路径', async () => {
    const uploads = [];
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            uploads.push(body);
            if (body.preview === true) {
                return jsonResponse({
                    ok: true, action: 'kb-preview', result: {
                        action: 'kb-preview', kind: 'image', name: '图片.png', ext: 'png', bytes: 2048,
                        chars: 0, chunks: 0, head: '', truncated: false, images: 1,
                        note: '图片可以直接入库（存缩略图与图片块）；要读出里面的文字，请点「识别文字」。',
                    },
                });
            }
            return jsonResponse({
                ok: true, action: 'kb-ingest', result: { action: 'kb-ingest', docId: 'kb:img', path: 'uploads/图片.png', chars: 0, chunks: 1 },
            });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const reader = countingFileReader();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [
            { name: '图片.png', size: 2048, __body: 'png' },
        ] } });
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);
        const card = byProp(tree, 'data-kb-preview-item', '图片.png')[0];
        assert.ok(card, '图片预览卡要在');
        assert.equal(card.props['data-kb-preview-kind'], 'image');
        const ingestButton = byProp(card, 'data-kb-preview-ingest', '图片.png')[0];
        assert.ok(ingestButton, '图片要给显式「入库」（存缩略图与图片块）');
        assert.ok(byProp(card, 'data-kb-preview-ocr', '图片.png')[0], '识别文字入口同时保留');

        ingestButton.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 40));
        tree = render(panel);
        const ingests = uploads.filter((one) => one.preview !== true);
        assert.equal(ingests.length, 1, '点了才入库');
        assert.equal(ingests[0].name, '图片.png');
        assert.equal(ingests[0].data, Buffer.from('png').toString('base64'), '发内容');
        assert.equal('path' in ingests[0], false, '不发客户端路径');
        assert.equal(byProp(tree, 'data-kb-preview-item', '图片.png').length, 0, '入库后预览卡收掉');
        assert.ok(textsOf(tree).some((t) => t.includes('已入库：图片.png')), '要留回执');
    } finally {
        reader.restore();
    }
});

// ── 图库（第九十三轮）：真实端点 + 真实交互 ───────────────────────────────
//
// 这一组不测「树里有个按钮」：它按服务端契约造响应，然后点页签、点卡片、
// 改输入、按保存 / 描述 / 搜索，断言发出去的 URL 与 body 真的是那条请求。
// 云端能力默认必须只写原因、不发请求 —— 这条用「调用记录里没有那个 action」钉死。

const GALLERY_ID_A = 'a'.repeat(64);
const GALLERY_ID_B = 'b'.repeat(64);
const GALLERY_FOLDERS_FIXTURE = [
    { path: '界面', name: '界面', parent: '', depth: 0, direct: 1, total: 1 },
    { path: '界面/深色', name: '深色', parent: '界面', depth: 1, direct: 0, total: 0 },
];
const galleryStatusFixture = {
    ok: true, root: 'D:/proj/gallery', version: 2,
    counts: { items: 2, withThumb: 1, withAesthetic: 0, withFolder: 1, folders: 2 },
    folders: GALLERY_FOLDERS_FIXTURE,
    thumbnail: { configured: true, lastError: '' },
    lexical: { available: true, note: '按标签与描述做词法匹配。' },
    textEmbedding: { mode: 'local', remote: false, model: 'local-hash', dimensions: 256, note: '本地向量，不联网。' },
    imageEmbedding: { provider: 'local', model: 'clip', configured: true, probed: true, dim: 512, note: '端点已通过探针。' },
    vision: { channel: 'openai', provider: 'openai', model: 'gpt-4o-mini', configured: true, note: '描述仍要显式 allowUpload 才会把图片发出去。' },
    uploadPolicy: '需要显式 allowUpload',
};
function galleryItemFixture(id, overrides = {}) {
    return Object.assign({
        id, bytes: 2048, mediaType: 'image/png', width: 640, height: 400,
        addedAt: '2026-11-18T10:00:00.000Z', thumb: null, category: '', folder: '', favorite: false,
        note: '', tags: [], manualTags: [], aesthetic: null,
    }, overrides);
}
const galleryA = galleryItemFixture(GALLERY_ID_A, {
    addedAt: '2026-11-18T10:00:00.000Z', category: '界面参考', folder: '界面', favorite: true, tags: ['深色', '截图'],
});
const galleryB = galleryItemFixture(GALLERY_ID_B, { addedAt: '2026-11-19T10:00:00.000Z', bytes: 4096, width: 800, height: 600 });
function jsonResponse(body) {
    // text() 与 json() 给同一份 body：客户端对 kb-upload 走 text-first 解析（413 空 body
    // 时 response.json() 会把 HTTP 状态吞掉），假响应也要像真实 Response 一样两种都可用。
    const text = JSON.stringify(body);
    return Promise.resolve({
        ok: true, status: 200,
        headers: { get: () => 'application/json; charset=utf-8' },
        json: () => Promise.resolve(body),
        text: () => Promise.resolve(text),
    });
}
/** 造一个可流式读取的 NDJSON 响应：每个事件前隔 gapMs，便于在中间态渲染进度条。 */
function ndjsonResponse(events, gapMs = 0) {
    const encoder = new TextEncoder();
    const body = new ReadableStream({
        async start(controller) {
            for (const event of events) {
                if (gapMs > 0) await new Promise((resolve) => setTimeout(resolve, gapMs));
                controller.enqueue(encoder.encode(JSON.stringify(event) + '\n'));
            }
            controller.close();
        },
    });
    return Promise.resolve({
        ok: true, status: 200,
        headers: { get: () => 'application/x-ndjson; charset=utf-8' },
        body,
    });
}
function galleryListBody() {
    return { ok: true, total: 2, offset: 0, limit: 200, items: [galleryA, galleryB] };
}
function gallerySearchBody(allowUpload = false) {
    // 照服务端真实口径：未带 allowUpload 时 textVector 退回本地哈希、imageText 不发。
    return {
        ok: true, query: '深色', imageId: '', total: 1, limit: 60,
        signals: {
            lexical: { used: true, note: '词法命中 1 张。' },
            tag: { used: true, note: '标签命中 1 张。' },
            textVector: allowUpload
                ? { used: true, mode: 'remote', note: '文本向量相似度已融合。' }
                : { used: true, mode: 'local', note: '本地哈希向量：只补字面召回。' },
            imageText: allowUpload
                ? { used: true, note: '图文对齐已融合。' }
                : { used: false, reason: '文字查询的多模态向量会把查询词发给图文端点（可能计费）。确认后传 allowUpload:true；这次没有发请求。' },
            imageImage: { used: false, reason: '以图搜图会把查询图发给多模态端点（可能计费）。确认后传 allowUpload:true。' },
        },
        activeSignals: allowUpload ? ['lexical', 'tag', 'textVector', 'imageText'] : ['lexical', 'tag', 'textVector'],
        items: [Object.assign({}, galleryA, { score: 0.87, parts: { lexical: 0.5 } })],
        notes: [],
    };
}
/** 图库端点路由：GET 列表 / 搜索 / 状态，POST 写入；别的路径回记忆快照。 */
function galleryFetch(calls, extra = {}) {
    return (url, init) => {
        const text = String(url);
        const request = init || {};
        calls.push({ url: text, method: String(request.method || 'GET'), body: request.body ? JSON.parse(request.body) : null });
        if (text.includes('/office-memory/snapshot')) return jsonResponse(makeSnapshot());
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [{ provider: 'deepseek', models: [{ id: 'deepseek-chat', label: 'DeepSeek Chat' }] }], error: '' });
        }
        if (text.includes('/office-memory/gallery')) {
            if (text.includes('view=status')) {
                return jsonResponse(typeof extra.status === 'function' ? extra.status() : (extra.status || galleryStatusFixture));
            }
            if (request.method === 'POST') {
                const body = JSON.parse(String(request.body || '{}'));
                const answer = extra.post ? extra.post(body) : { ok: true, item: galleryA };
                return jsonResponse(answer);
            }
            if (text.includes('view=search') || text.includes('&q=')) {
                return jsonResponse(extra.search ? extra.search(text) : gallerySearchBody(text.includes('allowUpload=true')));
            }
            return jsonResponse(extra.list ? extra.list() : galleryListBody());
        }
        return jsonResponse(makeSnapshot());
    };
}
/** 打开「图库」页签后的一轮渲染（点页签本身会取状态 + 列表）。 */
async function openGallery(fetchImpl) {
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    const tab = walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'gallery');
    assert.ok(tab, '一级页签应有「图库」');
    tab.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    return { panel, tree };
}

/** 从图库请求 URL 里取一个查询参数；取不到（没带）返回 null。 */
function galleryParamOf(call, key) {
    try {
        return new URL(String(call.url), 'http://localhost').searchParams.get(key);
    } catch {
        return null;
    }
}

await check('图库：点页签按真实端点取状态与列表，左树右网格、能力如实', async () => {
    const calls = [];
    const { panel, tree } = await openGallery(galleryFetch(calls));
    assert.equal(byProp(tree, 'data-gallery-panel', 'true').length, 1, '图库是独立面板');
    assert.ok(calls.some((call) => call.url.includes('view=status')), '进页签要取能力状态');
    assert.ok(calls.some((call) => call.url.includes('view=list')), '进页签要取真实列表');
    assert.ok(byProp(tree, 'data-gallery-sidebar', 'true').length === 1, '左栏树（全部 / 收藏 / 分类）');
    assert.deepEqual(byProp(tree, 'data-gallery-item').map((n) => n.props['data-gallery-item']), [GALLERY_ID_B, GALLERY_ID_A],
        '默认按加入时间倒序：新的在前');
    assert.ok(textsOf(tree).some((t) => t === '全部图片'), '左栏有「全部图片」');
    assert.ok(textsOf(tree).some((t) => t === '界面参考'), '分类来自列表里的真实条目');
    assert.deepEqual(byProp(tree, 'data-gallery-capability').map((n) => n.props['data-gallery-capability']),
        ['thumbnail', 'lexical', 'text-vector', 'image-embedding', 'vision'],
        '五项能力的可用性逐项写出来');
    assert.ok(textsOf(tree).some((t) => t.includes('2 张')), '计数用服务端 total');
    // 回车 = 打开原图预览；再点关闭要能收掉（不留一个盖住面板的遮罩）。
    byProp(tree, 'data-gallery-item', GALLERY_ID_A)[0].props.onKeyDown({ key: 'Enter', preventDefault() {} });
    const opened = render(panel);
    assert.equal(byProp(opened, 'data-gallery-overlay', 'true').length, 1, '回车要打开原图预览');
    byProp(opened, 'data-gallery-preview-close', 'true')[0].props.onClick();
    const closed = render(panel);
    assert.equal(byProp(closed, 'data-gallery-overlay', 'true').length, 0, '关闭预览要收掉遮罩');
});

await check('图库：分类筛选与面包屑都回到真实 list 请求，根可点回全部', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    const category = byProp(first, 'data-gallery-category', '界面参考')[0];
    assert.ok(category, '左栏分类可点');
    category.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    let tree = render(panel);
    const filtered = calls.filter((call) => call.method === 'GET' && call.url.includes('view=search') && call.url.includes('category='));
    assert.equal(filtered.length, 1, '切分类要走 search 的筛选分支');
    const lists = () => calls.filter((call) => call.method === 'GET' && call.url.includes('view=list'));
    assert.ok(byProp(tree, 'data-gallery-crumb-category', '界面参考').length === 1, '面包屑要写出当前分类');
    byProp(tree, 'data-gallery-crumb-root', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const rootList = lists()[lists().length - 1];
    assert.ok(!rootList.url.includes('category='), '点面包屑根要回到全部（请求里不再带分类）');
    assert.equal(byProp(tree, 'data-gallery-crumb-category').length, 0, '清掉分类后面包屑只剩根');
});

await check('图库：选中后改标签 / 收藏 / 备注并保存，发 manual patch', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls, {
        post: (body) => {
            if (body.action === 'manual') {
                return { ok: true, item: Object.assign({}, galleryA, { note: '来自测试的备注', tags: ['深色', '界面'], favorite: true }) };
            }
            return { ok: true, item: galleryA };
        },
    }));
    const card = byProp(first, 'data-gallery-item', GALLERY_ID_A)[0];
    assert.ok(card, '应有图库卡片');
    card.props.onClick();
    let tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-detail-shown', GALLERY_ID_A).length, 1, '选中后出详情');
    const note = byProp(tree, 'data-gallery-note', 'true')[0];
    note.props.onChange({ target: { value: '来自测试的备注' } });
    tree = render(panel);
    const save = byProp(tree, 'data-gallery-save', 'true')[0];
    save.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const manual = calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'manual');
    assert.equal(manual.length, 1, '保存只发一次 manual');
    assert.equal(manual[0].body.id, GALLERY_ID_A);
    assert.equal(manual[0].body.patch.note, '来自测试的备注');
    assert.deepEqual(manual[0].body.patch.tags, ['深色', '截图'], '标签原样切分后写回');
    assert.equal(manual[0].body.patch.category, '界面参考');
    assert.equal(manual[0].body.patch.favorite, true);
});

await check('图库：名称与描述可编辑，描述默认回落模型总结（第一百零八轮）', async () => {
    const calls = [];
    const described = Object.assign({}, galleryA, {
        name: '旧名称',
        aesthetic: { subject: '深色界面', composition: '居中', palette: ['#111'], light: '低', material: '', mood: '冷静', summary: '一张深色界面截图', tags: ['界面'], model: 'gpt-4o-mini', provider: 'openai', channel: 'openai', generatedAt: '2026-11-20T00:00:00.000Z' },
    });
    const { panel, tree: first } = await openGallery(galleryFetch(calls, {
        list: () => Object.assign(galleryListBody(), { items: [described, galleryB] }),
        post: (body) => body.action === 'manual'
            ? { ok: true, item: Object.assign({}, described, { name: body.patch.name, description: body.patch.description }) }
            : { ok: true, item: described },
    }));
    byProp(first, 'data-gallery-item', GALLERY_ID_A)[0].props.onClick();
    let tree = render(panel);
    const name = byProp(tree, 'data-gallery-name', 'true')[0];
    const description = byProp(tree, 'data-gallery-description', 'true')[0];
    assert.equal(name.props.value, '旧名称', '名称要回填');
    assert.equal(description.props.value, '一张深色界面截图', '没改写时描述显示模型总结');
    name.props.onChange({ target: { value: '新名称' } });
    tree = render(panel);
    byProp(tree, 'data-gallery-description', 'true')[0].props.onChange({ target: { value: '用户改写的描述' } });
    tree = render(panel);
    const save = byProp(tree, 'data-gallery-save', 'true')[0];
    assert.equal(textsOf(save)[0], '保存细节', '保存按钮文案要覆盖名称与描述');
    save.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const manual = calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'manual');
    assert.equal(manual.length, 1);
    assert.equal(manual[0].body.patch.name, '新名称');
    assert.equal(manual[0].body.patch.description, '用户改写的描述');
});

await check('图库：没勾「允许上传」时描述 / 以图搜图只写原因、不发请求', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    byProp(first, 'data-gallery-item', GALLERY_ID_A)[0].props.onClick();
    let tree = render(panel);
    byProp(tree, 'data-gallery-describe', 'true')[0].props.onClick();
    tree = render(panel);
    assert.equal(calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'describe').length, 0,
        '未允许上传绝不能把原图发出去');
    assert.ok(byProp(tree, 'data-gallery-error', 'true').length === 1, '要如实写出为什么没发');
    byProp(tree, 'data-gallery-search-image', 'true')[0].props.onClick();
    tree = render(panel);
    assert.equal(calls.filter((call) => call.url.includes('view=search')).length, 0, '未允许时以图搜图也不该发查询图');
});

await check('图库：勾选允许 → 描述要两步确认，确认后 body 带 allowUpload', async () => {
    const calls = [];
    const described = Object.assign({}, galleryA, {
        aesthetic: { subject: '深色界面', composition: '居中', palette: ['#111'], light: '低', material: '', mood: '冷静', summary: '一张深色界面截图', tags: ['界面'], model: 'gpt-4o-mini', provider: 'openai', channel: 'openai', generatedAt: '2026-11-20T00:00:00.000Z' },
    });
    const { panel, tree: first } = await openGallery(galleryFetch(calls, {
        list: () => Object.assign(galleryListBody(), { items: [described, galleryB] }),
        post: (body) => body.action === 'describe' ? { ok: true, item: described, note: '描述已生成。' } : { ok: true, item: galleryA },
    }));
    byProp(first, 'data-gallery-item', GALLERY_ID_A)[0].props.onClick();
    let tree = render(panel);
    byProp(tree, 'data-gallery-allow', 'true')[0].props.onChange({ target: { checked: true } });
    tree = render(panel);
    byProp(tree, 'data-gallery-describe', 'true')[0].props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-describe-confirm', 'true').length, 1, '第一次点只出确认，不发送');
    assert.equal(calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'describe').length, 0);
    // 两步确认之间撤回授权：确认按钮还在，但不能照旧发 allowUpload。
    byProp(tree, 'data-gallery-allow', 'true')[0].props.onChange({ target: { checked: false } });
    tree = render(panel);
    byProp(tree, 'data-gallery-describe-confirm', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 10));
    tree = render(panel);
    assert.equal(calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'describe').length, 0, '撤权后不能发送');
    assert.ok(textsOf(tree).some((t) => t.includes('授权已撤回')), '要说明为什么没发');
    // 重新授权后走正常两步。
    byProp(tree, 'data-gallery-allow', 'true')[0].props.onChange({ target: { checked: true } });
    tree = render(panel);
    byProp(tree, 'data-gallery-describe', 'true')[0].props.onClick();
    tree = render(panel);
    byProp(tree, 'data-gallery-describe-confirm', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const describe = calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'describe');
    assert.equal(describe.length, 1);
    assert.equal(describe[0].body.id, GALLERY_ID_A);
    assert.equal(describe[0].body.allowUpload, true);
    assert.ok(textsOf(tree).some((t) => t.includes('一张深色界面截图')), '描述要显示在详情里');
});

await check('图库：上传真实走 add（base64），重复内容按 duplicate 如实提示', async () => {
    const calls = [];
    const previousReader = globalThis.FileReader;
    class FakeReader {
        readAsDataURL(file) {
            this.result = 'data:image/png;base64,' + Buffer.from(String(file && file.__body || '')).toString('base64');
            Promise.resolve().then(() => { if (typeof this.onload === 'function') this.onload(); });
        }
    }
    globalThis.FileReader = FakeReader;
    try {
        const { panel, tree: first } = await openGallery(galleryFetch(calls, {
            post: (body) => body.action === 'add'
                ? { ok: true, added: true, duplicate: true, item: galleryA, note: '内容相同，复用已有条目。' }
                : { ok: true, item: galleryA },
        }));
        const input = byProp(first, 'data-gallery-upload-input', 'true')[0];
        assert.ok(input, '图库要有真实文件输入');
        input.props.onChange({ target: { files: [Object.assign({ name: '截图.png', __body: 'png-bytes' })], value: '' } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        const tree = render(panel);
        const added = calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'add');
        assert.equal(added.length, 1, '选文件后要发一次 add');
        assert.equal(added[0].body.name, '截图.png');
        assert.equal(added[0].body.data, Buffer.from('png-bytes').toString('base64'), '必须发原始字节的 base64');
        assert.ok(textsOf(tree).some((t) => t.includes('已经在图库里')), '重复内容要如实说，不是假装新增');
        assert.equal(byProp(tree, 'data-gallery-detail-shown', GALLERY_ID_A).length, 1, '重复时选中原来那张');
    } finally {
        globalThis.FileReader = previousReader;
    }
});

await check('图库：详情 / 大图标切换、名称 / 大小排序，header 可点', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    let tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-list', 'true').length, 1, '默认就是详情视图');
    assert.equal(byProp(tree, 'data-gallery-view', 'list')[0].props['aria-pressed'], true);
    assert.deepEqual(byProp(tree, 'data-gallery-subfolder').map((n) => n.props['data-gallery-subfolder']),
        ['界面'], '网格与详情共用同一份当前层文件夹');
    byProp(tree, 'data-gallery-view', 'grid')[0].props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-grid', 'true').length, 1, '大图标视图是网格');
    assert.equal(byProp(tree, 'data-gallery-subfolder', '界面').length, 1, '大图标里文件夹用磁贴呈现');
    byProp(tree, 'data-gallery-view', 'list')[0].props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-list', 'true').length, 1, '能切回详情视图');
    const sort = byProp(tree, 'data-gallery-sort', 'true')[0];
    sort.props.onChange({ target: { value: 'size' } });
    tree = render(panel);
    assert.deepEqual(byProp(tree, 'data-gallery-item').map((n) => n.props['data-gallery-item']), [GALLERY_ID_B, GALLERY_ID_A],
        '按大小也是 B 在前；sort 值要真的进状态');
    byProp(tree, 'data-gallery-sort-dir')[0].props.onClick();
    tree = render(panel);
    assert.deepEqual(byProp(tree, 'data-gallery-item').map((n) => n.props['data-gallery-item']), [GALLERY_ID_A, GALLERY_ID_B],
        '反向排序要真的翻转');
});

await check('图库：当前目录搜索走 view=search，未允许上传时不带 allowUpload', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    const search = walk(first).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找图库'));
    assert.ok(search, '搜索框占位要说明当前在图库');
    search.props.onChange({ target: { value: '深色' } });
    let tree = render(panel);
    const again = walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找图库'));
    again.props.onKeyDown({ key: 'Enter', target: { value: '深色' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const searched = calls.filter((call) => call.url.includes('view=search'));
    assert.equal(searched.length, 1, '回车发一次搜索');
    assert.ok(searched[0].url.includes('q=' + encodeURIComponent('深色')));
    assert.ok(!searched[0].url.includes('allowUpload'), '没勾允许上传就不带这个参数');
    assert.ok(textsOf(tree).some((t) => t.includes('跳过') && (t.includes('图文对齐') || t.includes('以图搜图'))), '未使用通道的原因要写出来');
});

await check('图库：搜索后出现清空按钮，清空与换文件夹都复位输入框（第一百一十一轮）', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    const findSearch = (tree) => walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找图库'));
    findSearch(first).props.onChange({ target: { value: '深色' } });
    let tree = render(panel);
    findSearch(tree).props.onKeyDown({ key: 'Enter', target: { value: '深色' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(findSearch(tree).props.value, '深色', '检索词应留在输入框');
    assert.equal(byProp(tree, 'data-gallery-sort', 'true')[0].props.value, 'score', '检索结果默认按相似度排');
    const clear = byProp(tree, 'data-memory-search-clear', 'true')[0];
    assert.ok(clear, '图库检索后要有清空按钮，否则只能手动删字再回车');

    clear.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(findSearch(tree).props.value, '', '清空后输入框不能留着旧词');
    const last = calls[calls.length - 1];
    assert.equal(galleryParamOf(last, 'q'), null, '清空后请求里不该再带 q');
    assert.equal(galleryParamOf(last, 'view'), 'list', '清空检索回到列表');
    assert.equal(byProp(tree, 'data-gallery-sort', 'true')[0].props.value, 'added', '清空检索回到加入时间');

    findSearch(tree).props.onChange({ target: { value: '深色' } });
    tree = render(panel);
    findSearch(tree).props.onKeyDown({ key: 'Enter', target: { value: '深色' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    byProp(tree, 'data-gallery-folder', '界面')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(findSearch(tree).props.value, '', '换文件夹会清检索，输入框也要跟着复位');
    const afterFolder = calls[calls.length - 1];
    assert.equal(galleryParamOf(afterFolder, 'folder'), '界面');
    assert.equal(galleryParamOf(afterFolder, 'q'), null);
});

await check('图库：搜索无命中显示空态与「清空搜索」，不把整库画成结果（第一百一十一轮）', async () => {
    const calls = [];
    const emptyBody = {
        ok: true, query: '量子纠缠', imageId: '', total: 0, limit: 60,
        signals: { lexical: { used: true, note: '词法命中 0 张。' } },
        activeSignals: ['lexical'], items: [], notes: ['没有匹配的图片：换一个词。'], reason: 'no-match',
    };
    const { panel, tree: first } = await openGallery(galleryFetch(calls, { search: () => emptyBody }));
    const findSearch = (tree) => walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找图库'));
    findSearch(first).props.onChange({ target: { value: '量子纠缠' } });
    let tree = render(panel);
    findSearch(tree).props.onKeyDown({ key: 'Enter', target: { value: '量子纠缠' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.ok(textsOf(tree).some((t) => t.includes('没有匹配的图片')), '无命中要显示空态');
    assert.equal(byProp(tree, 'data-gallery-item').length, 0, '不能把整库当结果画出来');
    const clear = byProp(tree, 'data-gallery-empty-action', 'clear')[0];
    assert.ok(clear, '空态要给出「清空搜索」，不要只让用户自己删字');
    clear.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(findSearch(tree).props.value, '');
    assert.equal(galleryParamOf(calls[calls.length - 1], 'view'), 'list');
    assert.equal(galleryParamOf(calls[calls.length - 1], 'q'), null);
});

await check('图库：勾选「文字搜索走云端（含向量）」后才带 allowUpload，取消又回本地', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    const findSearch = (tree) => walk(tree).find((n) => n.props && typeof n.props.placeholder === 'string'
        && n.props.placeholder.startsWith('找图库'));
    findSearch(first).props.onChange({ target: { value: '深色' } });
    let tree = render(panel);
    findSearch(tree).props.onKeyDown({ key: 'Enter', target: { value: '深色' } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    let searched = calls.filter((call) => call.url.includes('view=search') && call.url.includes('q='));
    assert.equal(searched.length, 1);
    assert.ok(!searched[0].url.includes('allowUpload'), '默认文字搜索是本地通道，不带上传授权');
    assert.ok(textsOf(tree).some((t) => t.includes('本地检索')), '状态栏要写明当前是本地检索');
    // 勾选：正在进行的文字检索立刻带 allowUpload 重查，不能只改状态不发请求。
    byProp(tree, 'data-gallery-cloud-search', 'true')[0].props.onChange({ target: { checked: true } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    searched = calls.filter((call) => call.url.includes('view=search') && call.url.includes('q='));
    assert.equal(searched.length, 2, '勾选后要重发一次');
    assert.ok(searched[1].url.includes('allowUpload=true'), '用户明确勾选后请求才带 allowUpload');
    assert.ok(textsOf(tree).some((t) => t.includes('云端检索')), '状态按服务端真实 textVector.mode / imageText 说话');
    // 取消勾选：回到本地，请求里不再有这个参数。
    byProp(tree, 'data-gallery-cloud-search', 'true')[0].props.onChange({ target: { checked: false } });
    await new Promise((resolve) => setTimeout(resolve, 20));
    searched = calls.filter((call) => call.url.includes('view=search') && call.url.includes('q='));
    assert.equal(searched.length, 3);
    assert.ok(!searched[2].url.includes('allowUpload'), '取消后要撤回授权');
});

await check('图库：文件夹树按状态渲染，点文件夹只改浏览范围、面包屑可回根（第一百零六轮）', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    assert.deepEqual(byProp(first, 'data-gallery-folder-row').map((n) => n.props['data-gallery-folder-row']),
        ['界面', '界面/深色'], '显式空文件夹也要出现，路径按层级排');
    assert.equal(byProp(first, 'data-gallery-folder-depth', '1').length, 1, '子文件夹要有缩进层级');
    const lists = () => calls.filter((call) => call.method === 'GET' && call.url.includes('view=list'));
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), null, '初始列表不带文件夹');

    byProp(first, 'data-gallery-folder', '界面')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    let tree = render(panel);
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), '界面', '点文件夹走同一条 list，只加 folder');
    assert.equal(byProp(tree, 'data-gallery-crumb-folder', '界面').length, 1, '面包屑写出当前文件夹');
    assert.equal(byProp(tree, 'data-gallery-folder', '界面')[0].props['aria-pressed'], true, '当前文件夹高亮');
    assert.ok(byProp(tree, 'data-gallery-folder-rename', 'true').length === 1, '当前文件夹才给改名 / 删除');

    byProp(tree, 'data-gallery-crumb-root', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), null, '面包屑根回到全部');
    assert.equal(byProp(tree, 'data-gallery-crumb-folder').length, 0);
});

await check('图库：默认详细信息、文件夹行与树折叠都按知识库资源管理器（第一百零九轮）', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls));
    assert.equal(byProp(first, 'data-gallery-list', 'true').length, 1, '默认是详细信息列表');
    assert.equal(byProp(first, 'data-gallery-grid', 'true').length, 0);
    assert.equal(byProp(first, 'data-gallery-file-head', 'true').length, 1, '详细信息有列头');
    assert.deepEqual(byProp(first, 'data-gallery-sort-head').map((n) => n.props['data-gallery-sort-head']),
        ['name', 'size', 'added'], '名称 / 大小 / 加入时间可点排序');
    assert.deepEqual(byProp(first, 'data-gallery-subfolder').map((n) => n.props['data-gallery-subfolder']),
        ['界面'], '当前层的文件夹排在图片前面');
    byProp(first, 'data-gallery-toggle-folder', '界面')[0].props.onClick({ stopPropagation() {} });
    let tree = render(panel);
    assert.deepEqual(byProp(tree, 'data-gallery-folder-row').map((n) => n.props['data-gallery-folder-row']),
        ['界面'], '折叠后子文件夹不进 DOM');
    assert.equal(byProp(tree, 'data-gallery-toggle-folder', '界面')[0].props['data-gallery-toggle'], 'closed');
    byProp(tree, 'data-gallery-toggle-folder', '界面')[0].props.onClick({ stopPropagation() {} });
    tree = render(panel);
    assert.deepEqual(byProp(tree, 'data-gallery-folder-row').map((n) => n.props['data-gallery-folder-row']),
        ['界面', '界面/深色'], '再点展开，子文件夹回来');

    byProp(tree, 'data-gallery-subfolder', '界面')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const lists = () => calls.filter((call) => call.method === 'GET' && call.url.includes('view=list'));
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), '界面', '主清单点文件夹只改浏览层');
    assert.equal(byProp(tree, 'data-gallery-subfolder', '界面/深色').length, 1, '子文件夹进入当前层清单');
    assert.equal(byProp(tree, 'data-gallery-nav', 'back')[0].props.disabled, false, '换层后可后退');
    // 面包屑回根会清掉后退栈：根不是「上一页」，再点后退不该弹回旧文件夹。
    byProp(tree, 'data-gallery-crumb-root', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), null, '面包屑回根');
    assert.equal(byProp(tree, 'data-gallery-nav', 'back')[0].props.disabled, true, '面包屑回根同时清空后退栈');
    // 重新进一层，再验证后退确实回根。
    byProp(tree, 'data-gallery-subfolder', '界面')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    byProp(tree, 'data-gallery-nav', 'back')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), null, '后退回根');
    assert.equal(byProp(tree, 'data-gallery-subfolder', '界面').length, 1, '根层文件夹仍在清单里');
});

await check('图库：新建 / 改名 / 删除空文件夹都走真实 POST 并切换浏览位置（第一百零六轮）', async () => {
    const calls = [];
    const folders = GALLERY_FOLDERS_FIXTURE.map((row) => Object.assign({}, row));
    const status = () => Object.assign({}, galleryStatusFixture, {
        folders: folders.map((row) => Object.assign({}, row)),
        counts: Object.assign({}, galleryStatusFixture.counts, { folders: folders.length }),
    });
    const post = (body) => {
        if (body.action === 'folder-create') {
            folders.push({
                path: body.path, name: body.path.split('/').pop(), parent: '',
                depth: body.path.split('/').length - 1, direct: 0, total: 0,
            });
            return { ok: true, created: true, path: body.path };
        }
        if (body.action === 'folder-rename') {
            const from = String(body.from);
            const to = String(body.to);
            folders.splice(0, folders.length, ...folders.map((row) => {
                if (row.path !== from && !row.path.startsWith(from + '/')) return row;
                const path = to + row.path.slice(from.length);
                return Object.assign({}, row, {
                    path, name: path.split('/').pop(),
                    parent: path.includes('/') ? path.slice(0, path.lastIndexOf('/')) : '',
                    depth: path.split('/').length - 1,
                });
            }));
            return { ok: true, from, to, moved: 0, renamedFolders: 1 };
        }
        if (body.action === 'folder-remove') {
            const path = String(body.path);
            folders.splice(0, folders.length,
                ...folders.filter((row) => row.path !== path && !row.path.startsWith(path + '/')));
            return { ok: true, path, removedFolders: 1 };
        }
        return { ok: true, item: galleryA };
    };
    const { panel, tree: first } = await openGallery(galleryFetch(calls, { status, post }));
    const posts = (action) => calls.filter((call) => call.method === 'POST' && call.body && call.body.action === action);
    const lists = () => calls.filter((call) => call.method === 'GET' && call.url.includes('view=list'));

    byProp(first, 'data-gallery-new-folder', 'true')[0].props.onClick();
    let tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-new-folder-bar', 'true').length, 1, '新建文件夹在命令行下方出条，不弹窗');
    assert.ok(textsOf(byProp(tree, 'data-gallery-new-folder-bar', 'true')[0]).join(' ').includes('全部图片'), '条上写清建在哪一层');
    byProp(tree, 'data-gallery-new-folder-input')[0].props.onChange(eventOf({ value: '素材' }));
    tree = render(panel);
    byProp(tree, 'data-gallery-new-folder-submit', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 40));
    tree = render(panel);
    assert.equal(posts('folder-create').length, 1);
    assert.equal(posts('folder-create')[0].body.path, '素材');
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), '素材', '创建后直接进入新文件夹');
    assert.ok(byProp(tree, 'data-gallery-folder', '素材').length === 1, '状态刷新后新文件夹出现在树里');

    byProp(tree, 'data-gallery-folder-rename', 'true')[0].props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-folder-rename-input')[0].props.value, '素材');
    byProp(tree, 'data-gallery-folder-rename-input')[0].props.onChange(eventOf({ value: '参考' }));
    tree = render(panel);
    byProp(tree, 'data-gallery-folder-rename-submit', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 40));
    tree = render(panel);
    assert.equal(posts('folder-rename').length, 1);
    assert.equal(posts('folder-rename')[0].body.from, '素材');
    assert.equal(posts('folder-rename')[0].body.to, '参考');
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), '参考', '改名后停在新路径');
    assert.equal(byProp(tree, 'data-gallery-crumb-folder', '参考').length, 1);

    byProp(tree, 'data-gallery-folder-remove', 'true')[0].props.onClick();
    tree = render(panel);
    assert.ok(textsOf(byProp(tree, 'data-gallery-folder-remove', 'true')[0]).includes('确认删'), '删除要两步确认');
    assert.equal(posts('folder-remove').length, 0, '第一下只进确认态');
    byProp(tree, 'data-gallery-folder-remove', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 40));
    tree = render(panel);
    assert.equal(posts('folder-remove').length, 1);
    assert.equal(posts('folder-remove')[0].body.path, '参考');
    assert.equal(galleryParamOf(lists()[lists().length - 1], 'folder'), null, '删除后回到父目录');
    assert.equal(byProp(tree, 'data-gallery-crumb-folder').length, 0);
});

await check('图库：详情里改文件夹并保存，manual patch 带 folder（第一百零六轮）', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls, {
        post: (body) => body.action === 'manual'
            ? { ok: true, item: Object.assign({}, galleryA, { folder: '成品' }) }
            : { ok: true, item: galleryA },
    }));
    byProp(first, 'data-gallery-item', GALLERY_ID_A)[0].props.onClick();
    let tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-folder-input')[0].props.value, '界面', '详情回填当前文件夹');
    byProp(tree, 'data-gallery-folder-input')[0].props.onChange(eventOf({ value: '成品' }));
    tree = render(panel);
    byProp(tree, 'data-gallery-save', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const manual = calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'manual');
    assert.equal(manual.length, 1);
    assert.equal(manual[0].body.patch.folder, '成品', '移动图片就是改 manual.folder');
    assert.equal(manual[0].body.patch.category, '界面参考');
});

await check('图库：视觉未配置时报错与能力面板都指向设置里的图库页（第一百零六轮）', async () => {
    const calls = [];
    const status = Object.assign({}, galleryStatusFixture, {
        vision: {
            channel: 'off', provider: '', model: '', configured: false,
            note: '没有配置视觉模型：图库只会保存图片与手工标签，不生成描述。',
        },
    });
    const { panel, tree: first } = await openGallery(galleryFetch(calls, { status }));
    byProp(first, 'data-gallery-item', GALLERY_ID_A)[0].props.onClick();
    let tree = render(panel);
    byProp(tree, 'data-gallery-allow', 'true')[0].props.onChange(eventOf({ checked: true }));
    tree = render(panel);
    byProp(tree, 'data-gallery-describe', 'true')[0].props.onClick();
    // 未配置时点描述会先重取一次能力状态（防止旧缓存挡住已配置的模型），等它回来再判。
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const errors = textsOf(byProp(tree, 'data-gallery-error', 'true')[0]).join(' ');
    assert.ok(errors.includes('视觉描述未配置'), '保留服务端的直接原因');
    assert.ok(errors.includes('设置 → 办公与记忆 → 图库 → 视觉描述'), '要告诉用户去哪配');
    assert.equal(calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'describe').length, 0,
        '未配置时一个请求都不发');
    const capText = textsOf(byProp(tree, 'data-gallery-capability', 'vision')[0]).join(' ');
    assert.ok(capText.includes('设置 → 办公与记忆 → 图库 → 视觉描述'), '能力面板也给出同一处设置路径');
});

await check('图库：点描述时重取能力状态，旧缓存不会挡住已配置的视觉模型（第一百零七轮）', async () => {
    const calls = [];
    const stale = Object.assign({}, galleryStatusFixture, {
        vision: {
            channel: 'off', provider: '', model: '', configured: false,
            note: '没有配置视觉模型：图库只会保存图片与手工标签，不生成描述。',
        },
    });
    let statusReads = 0;
    const { panel, tree: first } = await openGallery(galleryFetch(calls, {
        // 第一次读到的还是设置保存前的旧状态；点「生成视觉描述」时重取，服务端已配置。
        status: () => (++statusReads === 1 ? stale : galleryStatusFixture),
    }));
    byProp(first, 'data-gallery-item', GALLERY_ID_A)[0].props.onClick();
    let tree = render(panel);
    byProp(tree, 'data-gallery-allow', 'true')[0].props.onChange(eventOf({ checked: true }));
    tree = render(panel);
    byProp(tree, 'data-gallery-describe', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    assert.equal(byProp(tree, 'data-gallery-error', 'true').length, 0, '重取到已配置后不该再报未配置');
    const confirm = byProp(tree, 'data-gallery-describe-confirm', 'true')[0];
    assert.ok(confirm, '重取到已配置后应进入两步确认');
    assert.ok(statusReads >= 2, '旧缓存不可信时要重新向服务端取能力状态');
    confirm.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const describes = calls.filter((call) => call.method === 'POST' && call.body && call.body.action === 'describe');
    assert.equal(describes.length, 1, '确认后只发一次 describe');
    assert.equal(describes[0].body.allowUpload, true, '确认后才带 allowUpload');
});

await check('图库：服务端在确认后回未配置时仍给出设置路径（第一百零七轮）', async () => {
    const calls = [];
    const { panel, tree: first } = await openGallery(galleryFetch(calls, {
        post: (body) => body.action === 'describe'
            ? { ok: false, code: 'unconfigured', error: '没有配置视觉模型：图库只会保存图片与手工标签，不生成描述。' }
            : { ok: true, item: galleryA },
    }));
    byProp(first, 'data-gallery-item', GALLERY_ID_A)[0].props.onClick();
    let tree = render(panel);
    byProp(tree, 'data-gallery-allow', 'true')[0].props.onChange(eventOf({ checked: true }));
    tree = render(panel);
    byProp(tree, 'data-gallery-describe', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    byProp(tree, 'data-gallery-describe-confirm', 'true')[0].props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    tree = render(panel);
    const error = textsOf(byProp(tree, 'data-gallery-error', 'true')[0]).join(' ');
    assert.ok(error.includes('没有配置视觉模型'), '保留服务端的直接原因');
    assert.ok(error.includes('设置 → 办公与记忆 → 图库 → 视觉描述'), '服务端回执也要指向设置');
});

// ── KB：AI 入库的预估 → 暂停 → 续跑 ───────────────────────────────────────
//
// 这里用假的 FileReader 把文件内容变成 base64；断言三次 kb-upload 的 body：
// 1) aiEstimate:true（不落库）2) 确认入库（非续跑）3) aiResume:true。
// 三次的 data 必须逐字相同 —— 服务端就是靠正文内容 + 模型 profile 认任务的。

await check('知识库：AI 入库先预估、暂停后续跑，三次请求同一份内容', async () => {
    const calls = [];
    const bodies = [];
    const base = makeSnapshot();
    const snapshot = Object.assign({}, base, {
        config: Object.assign({}, base.config, {
            kbLlm: { mode: 'host', provider: 'deepseek', model: 'deepseek-chat', chunking: true, summary: true, ready: true, reason: '' },
        }),
    });
    let round = 0;
    const fetchImpl = (url, init) => {
        const text = String(url);
        calls.push({ url: text, body: init && init.body ? JSON.parse(init.body) : null });
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [{ provider: 'deepseek', models: [{ id: 'deepseek-chat', label: 'DeepSeek Chat' }] }], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String(init.body || '{}'));
            if (body.preview === true) {
                // 统一入口第一步：只抽取、不落库、不调模型。
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文内容', truncated: false, images: 0, note: '',
                } });
            }
            bodies.push(body);
            round += 1;
            if (round === 1) {
                assert.equal(body.aiEstimate, true, '第一步只预估');
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, path: 'uploads/笔记.txt', chars: 1200,
                    estimate: { chars: 1200, units: 3, chapters: 1, batches: 2, calls: 2, inputCharsPerCall: 800, maxCalls: 8, overBudget: false, estimatedInputTokens: 1800, estimatedOutputTokens: 240 },
                    ai: { mode: 'host', calls: 0 }, note: 'AI 未启用：将走确定性切块，不调用模型。',
                } });
            }
            if (round === 2) {
                assert.equal(body.aiResume, false, '确认入库不是续跑');
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', path: 'uploads/笔记.txt', paused: true, status: 'paused',
                    calls: 2, doneBatches: 1, batches: 2, stateKey: 'job-1',
                    ai: { mode: 'host', calls: 2, degraded: 0 }, note: '到本轮预算上限，任务已暂停。',
                } });
            }
            assert.equal(body.aiResume, true, '第三次要续跑');
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', path: 'uploads/笔记.txt', chars: 1200, chunks: 3, ai: { mode: 'host', calls: 4, degraded: 0 },
            } });
        }
        return jsonResponse(snapshot);
    };
    const previousReader = globalThis.FileReader;
    class FakeReader {
        readAsDataURL(file) {
            const encoded = Buffer.from(String(file && file.__body || '')).toString('base64');
            this.result = 'data:text/plain;base64,' + encoded;
            Promise.resolve().then(() => { if (typeof this.onload === 'function') this.onload(); });
        }
    }
    globalThis.FileReader = FakeReader;
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        walk(tree).find((n) => n.props && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        const toggle = byProp(tree, 'data-kb-ai-toggle', 'true')[0];
        assert.ok(toggle, 'AI 建库勾选框要在');
        assert.equal(toggle.props.disabled, false, '模型 ready 时可用');
        toggle.props.onChange({ target: { checked: true } });
        tree = render(panel);
        const input = byProp(tree, 'data-kb-upload-input', 'true')[0];
        assert.ok(input, '文件输入要在');
        input.props.onChange({ target: { files: [Object.assign({ name: '笔记.txt', __body: '一段正文内容' })] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        // 选完文件先给预览卡；勾了 AI 也只是等一句「入库」的明确指令。
        assert.equal(byProp(tree, 'data-kb-preview-item', '笔记.txt').length, 1, '先出预览卡');
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 0, '预览阶段不发预估');
        assert.equal(bodies.length, 0, '预览阶段不写库、不调模型');
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 1, '点「入库」才发预估');
        const texts = textsOf(tree);
        assert.ok(texts.some((t) => t.includes('调用 2')), '预估要写清调用数');
        assert.ok(texts.some((t) => t.includes('未写入任何内容')), '预估不落库要明说');
        const start = byProp(tree, 'data-kb-estimate-start', 'true')[0];
        assert.ok(start, '预估后要有「开始入库」');
        start.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-paused', 'true').length, 1, '暂停要留在界面上');
        assert.ok(textsOf(tree).some((t) => t.includes('批次 1 / 2')));
        assert.equal(byProp(tree, 'data-kb-resume', 'true')[0].props.disabled, false, '内容还在内存里，可续跑');
        byProp(tree, 'data-kb-resume', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-paused', 'true').length, 0, '续跑完成后暂停卡要收掉');
        // 库里已经有这份文档了，界面上不能再挂「待入库」：再点一次「入库」就是重复写。
        assert.equal(byProp(tree, 'data-kb-preview-item', '笔记.txt').length, 0, 'AI 入库完成后预览卡要收掉');
        assert.ok(byProp(tree, 'data-kb-upload-note', 'true').length === 1, '完成要有一行回执');
        assert.ok(textsOf(tree).some((t) => t.includes('已入库：笔记.txt')));
        assert.equal(bodies.length, 3);
        assert.equal(bodies[0].data, bodies[1].data, '确认入库要重发同一份内容');
        assert.equal(bodies[1].data, bodies[2].data, '续跑要重发同一份内容');
        assert.equal(bodies[0].ai, true);
        assert.equal(bodies[0].name, '笔记.txt');
    } finally {
        globalThis.FileReader = previousReader;
    }
});

await check('知识库：kb-upload 返回 HTTP 200 但 ok:false 时按失败处理，不画预估卡', async () => {
    const posts = [];
    const base = makeSnapshot();
    const snapshot = Object.assign({}, base, {
        config: Object.assign({}, base.config, {
            kbLlm: { mode: 'host', provider: 'deepseek', model: 'deepseek-chat', chunking: true, summary: true, ready: true, reason: '' },
        }),
    });
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [{ provider: 'deepseek', models: [{ id: 'deepseek-chat', label: 'DeepSeek Chat' }] }], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 6,
                    chars: 2, chunks: 1, head: '正文', truncated: false, images: 0, note: '',
                } });
            }
            posts.push(body);
            // 视图层把业务失败也放在 200 里：客户端必须看 body.ok，而不是只看 response.ok。
            return jsonResponse({ ok: false, error: 'AI 模型不可用（这是服务端业务失败）' });
        }
        return jsonResponse(snapshot);
    };
    const previousReader = globalThis.FileReader;
    class FakeReader {
        readAsDataURL() {
            this.result = 'data:text/plain;base64,' + Buffer.from('正文').toString('base64');
            Promise.resolve().then(() => { if (typeof this.onload === 'function') this.onload(); });
        }
    }
    globalThis.FileReader = FakeReader;
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        walk(tree).find((n) => n.props && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-preview-item', '笔记.txt').length, 1, '先出预览卡');
        assert.equal(posts.length, 0, '预览不是入库请求');
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(posts.length, 1, '预估要真的发一次请求');
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 0, '业务失败不能画出一张「可开始入库」的预估卡');
        assert.ok(textsOf(tree).some((t) => t.includes('预估失败') && t.includes('AI 模型不可用')), '失败原因要原样显示');
    } finally {
        globalThis.FileReader = previousReader;
    }
});

// ── 知识库上传工具栏：宿主模型清单也走 session.modelCatalog ──────────────
//
// 设置页早已改成 RPC 优先；上传工具栏原来还 fetch /office-memory/llm-models?cwd=…，
// 冷启动（无 cwd / 未登录 HTTP 不可用）时选不了模型。这里钉住同一条来源：
// RPC 优先且不碰 HTTP；RPC 失败只显示错误、不偷偷回退；只有旧宿主缺方法才走 HTTP；
// remote 迟绑定后自动刷新；请求里不带 Key / baseURL。

function kbAiSnapshot() {
    const base = makeSnapshot();
    return Object.assign({}, base, {
        config: Object.assign({}, base.config, {
            kbLlm: {
                mode: 'host', provider: 'p', model: 'm1', chunking: true, summary: true,
                ready: true, reason: '',
            },
        }),
    });
}

/** 上传工具栏取数：快照走假响应，llm-models 单独记一笔（用于断言「没走 HTTP」）。 */
function kbAiFetch(snapshot, calls, providers) {
    return (url, init) => {
        const text = String(url);
        calls.push({ url: text, init: init || null });
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: providers || [], error: '' });
        }
        return jsonResponse(snapshot);
    };
}

/** 打开知识库页签，等清单请求落地后再渲染一次。 */
async function openKbAiPanel(options) {
    const env = { fetch: options.fetch };
    if (options.remote !== undefined) env.remote = options.remote;
    const { mod, ctx, registrations, injects, mountRemoteNamespace } = loadBundle(env);
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);
    const tab = walk(tree).find((node) => node.props && node.props['data-tab'] === 'kb');
    assert.ok(tab, '应有知识库页签');
    tab.props.onClick();
    // 先渲染一次让切页签的 effect 真正跑起来（替身不会自动重渲染），再等 RPC 落地。
    tree = render(panel);
    await new Promise((resolve) => setTimeout(resolve, 30));
    tree = render(panel);
    return { tree, panel, injects, mountRemoteNamespace };
}

await check('上传工具栏：宿主 session.modelCatalog 填真实 provider/model，且不请求 llm-models', async () => {
    const calls = [];
    const rpcArgs = [];
    const remote = { session: {
        modelCatalog: async (...args) => {
            rpcArgs.push(args);
            return {
                ok: true,
                value: { groups: [{ id: 'p', name: 'P', models: [{ id: 'm1', name: 'Model One' }] }], failures: [] },
            };
        },
    } };
    const snapshot = kbAiSnapshot();
    const { tree } = await openKbAiPanel({
        remote,
        fetch: kbAiFetch(snapshot, calls, [{ provider: 'legacy', models: [{ id: 'old', label: 'Old' }] }]),
    });
    const select = byProp(tree, 'data-kb-ai-model', 'true')[0];
    assert.ok(select, 'host 通道要有模型下拉');
    assert.equal(select.type, 'select');
    assert.deepEqual(optionValuesOf(select), ['', 'p/m1'], '选项用 groups[].id / models[].id');
    assert.deepEqual(optionLabelsOf(select), ['宿主默认模型', 'p / Model One'], '显示名用 models[].name');
    assert.equal(calls.filter((call) => call.url.includes('/office-memory/llm-models')).length, 0,
        'RPC 成功时不得再打旧 HTTP 接口');
    assert.deepEqual(rpcArgs, [[]], 'RPC 调用不带参数（也就没有 cwd）');
    const wire = JSON.stringify({ rpcArgs: rpcArgs, calls: calls });
    assert.ok(!wire.includes('sk-'), '请求里不得出现 Key：' + wire);
    assert.ok(!wire.includes('baseURL'), '请求里不得出现 baseURL：' + wire);
});

await check('上传工具栏：RPC 失败只显示宿主错误，不 HTTP 回退，重试仍只走 RPC', async () => {
    const calls = [];
    let rpcCalls = 0;
    const remote = { session: { modelCatalog: async () => {
        rpcCalls += 1;
        return { ok: false, error: { code: 'gateway/forbidden', message: 'no' } };
    } } };
    const snapshot = kbAiSnapshot();
    const first = await openKbAiPanel({
        remote,
        fetch: kbAiFetch(snapshot, calls, [{ provider: 'legacy', models: [{ id: 'old', label: 'Old' }] }]),
    });
    const llmCalls = () => calls.filter((call) => call.url.includes('/office-memory/llm-models')).length;
    assert.equal(llmCalls(), 0, 'RPC 拒绝时不得退回 HTTP（权限拒绝尤其不能绕过）');
    let texts = textsOf(first.tree).join(' | ');
    assert.ok(texts.includes('gateway/forbidden'), '要显示宿主 error.code：' + texts);
    assert.ok(texts.includes('no'), '要显示宿主原始 message：' + texts);
    const retry = byProp(first.tree, 'data-kb-ai-catalog-retry', 'true')[0];
    assert.ok(retry, '失败后要给重试');
    retry.props.onClick();
    await new Promise((resolve) => setTimeout(resolve, 30));
    const tree = render(first.panel);
    assert.equal(rpcCalls, 2, '重试重发 RPC');
    assert.equal(llmCalls(), 0, '重试也不得改走 HTTP');
    texts = textsOf(tree).join(' | ');
    assert.ok(texts.includes('gateway/forbidden'), '重试失败后的错误仍可见：' + texts);
});

await check('上传工具栏：旧宿主返回 gateway/invocation-unavailable 时退回 HTTP 清单', async () => {
    const calls = [];
    const remote = { session: { modelCatalog: async () => ({
        ok: false,
        error: {
            code: 'gateway/invocation-unavailable',
            message: 'no active Remote method exports this endpoint',
        },
    }) } };
    const snapshot = kbAiSnapshot();
    const { tree } = await openKbAiPanel({
        remote,
        fetch: kbAiFetch(snapshot, calls, [{ provider: 'legacy', models: [{ id: 'old', label: 'Old' }] }]),
    });
    const llm = calls.filter((call) => call.url.includes('/office-memory/llm-models'));
    assert.equal(llm.length, 1, '明确的宿主缺方法才走 HTTP');
    const select = byProp(tree, 'data-kb-ai-model', 'true')[0];
    assert.deepEqual(optionValuesOf(select), ['', 'legacy/old'], 'HTTP 清单也投影成同一形状');
});

await check('上传工具栏：旧宿主没有 session.modelCatalog 才走 HTTP，请求不含 Key / baseURL', async () => {
    const calls = [];
    const remote = {};
    const snapshot = kbAiSnapshot();
    const { tree } = await openKbAiPanel({
        remote,
        fetch: kbAiFetch(snapshot, calls, [{ provider: 'legacy', models: [{ id: 'old', label: 'Old' }] }]),
    });
    const llm = calls.filter((call) => call.url.includes('/office-memory/llm-models'));
    assert.equal(llm.length, 1, '只有缺 session 方法时才走 HTTP');
    const select = byProp(tree, 'data-kb-ai-model', 'true')[0];
    assert.deepEqual(optionValuesOf(select), ['', 'legacy/old'], 'HTTP 清单也投影成同一形状');
    const wire = JSON.stringify(llm[0]);
    assert.ok(!wire.includes('sk-'), 'HTTP 请求不得带 Key：' + wire);
    assert.ok(!wire.includes('baseURL'), 'HTTP 请求不得带 baseURL：' + wire);
});

await check('上传工具栏：remote 迟绑定后自动刷新并改走 RPC，不再重复打旧 HTTP 接口', async () => {
    const calls = [];
    const snapshot = kbAiSnapshot();
    const remote = {};
    const { panel, injects, mountRemoteNamespace } = await openKbAiPanel({
        remote,
        fetch: kbAiFetch(snapshot, calls, [{ provider: 'legacy', models: [{ id: 'old', label: 'Old' }] }]),
    });
    const llmCalls = () => calls.filter((call) => call.url.includes('/office-memory/llm-models')).length;
    assert.equal(llmCalls(), 1, '方法还没出现时走旧 HTTP 路径');
    let rpcCalls = 0;
    const nextRemote = { session: { modelCatalog: async () => {
        rpcCalls += 1;
        return { ok: true, value: { groups: [{ id: 'late', name: 'Late', models: [{ id: 'late-model', name: 'Late Model' }] }], failures: [] } };
    } } };
    assert.ok(injects.some((item) => item.services.includes('remote.session')),
        '必须把 remote.session 声明进注入列表');
    mountRemoteNamespace('session', nextRemote.session);
    await new Promise((resolve) => setTimeout(resolve, 30));
    const tree = render(panel);
    assert.equal(rpcCalls, 1, '迟绑定后应自动走 RPC');
    assert.equal(llmCalls(), 1, '改走 RPC 后不应再打旧 HTTP 接口');
    assert.deepEqual(optionValuesOf(byProp(tree, 'data-kb-ai-model', 'true')[0]), ['', 'late/late-model'],
        '界面换成 RPC 返回的清单');
});


await check('归档隔离提示：只在归档视图（含空状态）显示，字符串去重且不透出路径', async () => {
    const warning = '归档目录由多个工作区共用，旧档归属无法核实；新归档已按工作区分开保存。';
    const extra = '共享归档的保留上限已跳过，旧档不会被自动删除。';
    const base = makeSnapshot();
    const store = Object.assign({}, base.stores[0], {
        // 混入重复、非字符串与空白：面板只应显示两条去重后的字符串。
        warnings: [warning, warning, 42, '  ', extra],
    });
    const kbTree = {
        path: '', name: '全部文档', docs: 0, chunks: 0, bytes: 0,
        children: [{ path: '归档', name: '归档', docs: 0, chunks: 0, bytes: 0, children: [] }],
    };
    const counts = Object.assign({}, base.counts, { kb: 0, kbTotal: 0, kbChunks: 0, kbBytes: 0 });
    const rootSnapshot = makeSnapshot({ stores: [store], kb: [], counts: counts, kbTree: kbTree });
    const archiveSnapshot = makeSnapshot({ stores: [store], kb: [], counts: counts, kbTree: kbTree, kbFolder: '归档' });
    const fetchImpl = (url) => Promise.resolve({
        ok: true, status: 200,
        json: () => Promise.resolve(String(url).includes('kbfolder=') ? archiveSnapshot : rootSnapshot),
    });
    const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
    mod.apply(ctx);
    const panel = regOf(registrations, 'conversation.view', 'office-memory');
    let tree = await renderPanelLoaded(panel);

    assert.equal(byProp(tree, 'data-archive-scope-warning', 'true').length, 0,
        '记忆页（非归档视图）不显示归档隔离提示');

    const kbTab = walk(tree).find((node) => node.props && node.props.role === 'tab' && node.props['data-tab'] === 'kb');
    kbTab.props.onClick();
    tree = render(panel);
    assert.equal(byProp(tree, 'data-archive-scope-warning', 'true').length, 0,
        '知识库根目录不显示归档隔离提示');

    const archiveRow = byProp(tree, 'data-kb-folder', '归档')[0];
    assert.ok(archiveRow, '知识库树里要有「归档」虚拟文件夹');
    archiveRow.props.onClick();
    tree = await renderPanelLoaded(panel);

    const boxes = byProp(tree, 'data-archive-scope-warning', 'true');
    assert.equal(boxes.length, 1, '归档文件夹只渲染一个提示框');
    assert.equal(textsOf(tree).filter((text) => text === warning).length, 1, '同一条提示只出现一次');
    assert.ok(textsOf(tree).includes(extra), '其余字符串提示保留');
    assert.ok(byProp(tree, 'data-kb-empty', 'true').length === 1, '空归档文件夹也要显示提示');
    const wire = JSON.stringify(boxes[0]);
    assert.ok(!/(?:^|[\s"'])[A-Za-z]:[\\/]/.test(wire), '提示框不得带本机绝对路径：' + wire);
});

// ── OCR 空响应（第九十八轮）：端点没按契约回话时，页面要说清是「空 / HTML / 状态码」 ──
//
// 事故现场：浏览器抛 `Failed to execute 'json' on 'Response': Unexpected end of JSON
// input`。那是 response.json() 撞上空 body 的原生错误：HTTP 状态、端点是谁、到底
// 是空响应还是 HTML 页面，全被这一句吃掉。这里用假端点把四类回话各放一遍，钉住
// 「分类 + 状态可见 + 不回响应体原文 + 失败不报页数」。

async function ocrFailureScenario({ name = '扫描件.png', response }) {
    const calls = [];
    const fetchImpl = (url, init) => {
        const target = String(url);
        calls.push({ url: target, init: init || {} });
        if (target.includes('/office-memory/kb-upload')) {
            return jsonResponse({
                ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'scan', name, ext: '.png', bytes: 2048,
                    chars: 0, chunks: 1, head: '', truncated: false, images: 1,
                    note: '这份文件没有可抽取的文本层（是扫描件）：要读文字请点「识别文字」。',
                },
            });
        }
        if (target.includes('/office-memory/ocr')) {
            if (response.throw !== undefined) return Promise.reject(new TypeError(response.throw));
            return Promise.resolve({
                ok: response.status >= 200 && response.status < 300,
                status: response.status,
                headers: { get: () => response.contentType || 'application/json; charset=utf-8' },
                text: () => Promise.resolve(response.bodyText),
            });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const restore = kbFileReaderFake();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name, __body: 'png-bytes' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        const button = byProp(tree, 'data-kb-preview-ocr', name)[0];
        assert.ok(button, '要有「识别文字」按钮');
        button.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        return { tree: render(panel), calls };
    } finally {
        restore();
    }
}

function ocrErrorLine(tree) {
    const box = byProp(tree, 'data-kb-ocr-error', 'true')[0];
    assert.ok(box, '失败要有错误行（data-kb-ocr-error）');
    return textsOf(box).join('');
}

await check('知识库：OCR 端点回空 body 时报「空响应 + HTTP 状态」，不把浏览器 JSON 原生错误丢给用户', async () => {
    const { tree } = await ocrFailureScenario({ response: { status: 200, bodyText: '' } });
    const line = ocrErrorLine(tree);
    assert.match(line, /空响应/);
    assert.match(line, /HTTP 200/);
    // 状态只报一次（诊断句不重复同一状态码），也不写「这不是文件本身的问题」这类
    // 没有证据的归因：这里能确认的只有「端点回了个空 body」。
    assert.equal((line.match(/HTTP 200/g) || []).length, 1, 'HTTP 状态只报一次');
    assert.equal(line.includes('这不是文件本身的问题'), false, '不替空响应猜原因');
    const all = textsOf(tree).join('\n');
    assert.equal(/Unexpected end of JSON input/.test(all), false, '浏览器 JSON 解析原生错误不能原样上屏');
    assert.equal(/0 页|0 ms/.test(all), false, '失败时不能显示「0 页 / 0 ms」');
    assert.match(all, /识别失败/, '标题要说失败，而不是报一个假的页数');
    assert.equal(byProp(tree, 'data-kb-ocr-text', 'true').length, 0, '失败不给正文编辑框');
    assert.equal(byProp(tree, 'data-kb-ocr-save', 'true').length, 0, '失败不给保存按钮');
});

await check('知识库：OCR 路由 404 空响应时点名「插件里没有这个路由」，仍带 HTTP 404', async () => {
    const { tree } = await ocrFailureScenario({ response: { status: 404, bodyText: '' } });
    const line = ocrErrorLine(tree);
    assert.match(line, /HTTP 404/);
    assert.match(line, /没有这个路由/);
    assert.equal((line.match(/HTTP 404/g) || []).length, 1, '404 只报一次，不重复同一句');
});

await check('知识库：OCR 端点回 HTML 时判成「HTML/XML 页面」，不把页面原文复制到面板', async () => {
    const { tree } = await ocrFailureScenario({
        response: { status: 502, bodyText: '<html><body><h1>502 Bad Gateway</h1></body></html>', contentType: 'text/html' },
    });
    const line = ocrErrorLine(tree);
    assert.match(line, /HTML\/XML/);
    assert.match(line, /HTTP 502/);
    assert.match(line, /text\/html/);
    assert.equal(line.includes('Bad Gateway') || line.includes('<html'), false, '响应体原文（可能是网关内部页）不能进面板');
    // 归因只到「上游/代理/本地路由没按契约回 JSON」为止：具体是谁没有证据。
    assert.equal(/多半是代理|检查网络与 baseURL|契约不一致/.test(line), false, '不写没有证据的推测归因');
});

await check('知识库：OCR 端点回 502 + body.error 时保留状态与机器可读 code（不再被吞成 400）', async () => {
    const { tree } = await ocrFailureScenario({
        response: {
            status: 502,
            bodyText: JSON.stringify({ ok: false, error: 'OCR 识别失败：识别服务暂时不可用', code: 'gateway_timeout_error', retryable: true }),
        },
    });
    const line = ocrErrorLine(tree);
    assert.match(line, /识别服务暂时不可用/);
    assert.match(line, /HTTP 502/);
    assert.match(line, /gateway_timeout_error/);
});

await check('知识库：OCR 请求本身失败（fetch reject）时只报识别失败，不编页数', async () => {
    const { tree } = await ocrFailureScenario({ response: { throw: 'Failed to fetch' } });
    const line = ocrErrorLine(tree);
    assert.match(line, /识别失败/);
    assert.match(line, /Failed to fetch/);
    assert.equal(/0 页|0 ms/.test(textsOf(tree).join('\n')), false, '失败时不能报「0 页 / 0 ms」');
});

await check('知识库：OCR 页没有正常结束时报出页码与 finishReason（第一百一十轮）', async () => {
    const { tree } = await ocrFailureScenario({
        response: {
            status: 200,
            bodyText: JSON.stringify({
                ok: true, fileName: '扫描件.png', pageCount: 3, text: '识别正文', usage: null, ms: 5,
                truncatedPages: [{ index: 2, finishReason: 'length' }, { index: 3, finishReason: 'content_filter' }],
            }),
        },
    });
    const box = byProp(tree, 'data-kb-ocr-truncated', '2')[0];
    assert.ok(box, '有未正常结束的页时要显示告警');
    const line = textsOf(box).join(' ');
    assert.match(line, /#2 length/);
    assert.match(line, /#3 content_filter/);
    assert.match(line, /可能不完整/);
    assert.ok(byProp(tree, 'data-kb-ocr-text', 'true').length === 1, '正文仍要能编辑');
});

// ── KB 上传的超限 / 非 JSON 回执（第九十八轮）：端点叫「知识库上传」，不是 OCR ──
//
// 事故现场：略超 12 MiB 的 KB JSON 请求被服务端立刻 destroy，浏览器只剩 fetch failed /
// UND_ERR_SOCKET；服务端改成结构化 413 + 有界关闭后，客户端也必须先读文本：413 的空
// body 若直接 response.json()，HTTP 状态又会丢一次。这里钉住四类回话，并防止把
// kb-upload 的失败叫成「OCR 端点」（saveOcrAsNote 也走这条通道）。

async function kbUploadFailureScenario(response) {
    const calls = [];
    const fetchImpl = (url, init) => {
        const target = String(url);
        calls.push({ url: target, init: init || {} });
        if (target.includes('/office-memory/kb-upload')) {
            return Promise.resolve({
                ok: response.status >= 200 && response.status < 300,
                status: response.status,
                headers: { get: () => response.contentType || 'application/json; charset=utf-8' },
                text: () => Promise.resolve(response.bodyText),
            });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const restore = kbFileReaderFake();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [
            { name: '笔记.txt', size: 1024, __body: '正文' },
        ] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        return { tree: render(panel), calls };
    } finally {
        restore();
    }
}

function kbUploadErrorLine(tree) {
    const line = textsOf(tree).find((text) => text.includes('预览失败'));
    assert.ok(line, 'KB 上传失败要有一条可见回执（预览失败：…）');
    return line;
}

await check('知识库上传：413 空 body 时报「空响应 + HTTP 413」，不叫 OCR、不丢状态', async () => {
    const { tree } = await kbUploadFailureScenario({ status: 413, bodyText: '' });
    const line = kbUploadErrorLine(tree);
    assert.match(line, /空响应/);
    assert.match(line, /HTTP 413/);
    assert.match(line, /知识库上传端点/, '失败文案要指向真正出事的端点');
    assert.equal(line.includes('OCR 端点'), false, 'kb-upload 的失败不能叫成 OCR 端点');
    assert.equal(/Unexpected end of JSON input/.test(textsOf(tree).join('\n')), false,
        '浏览器 JSON 解析原生错误不能原样上屏');
});

await check('知识库上传：后端结构化 413（body_too_large）保留服务端文案、HTTP 状态与 code', async () => {
    const { tree } = await kbUploadFailureScenario({
        status: 413,
        bodyText: JSON.stringify({
            ok: false, code: 'body_too_large',
            error: '请求体超过这个端点允许的 12582912 字节，这次请求没有被解析。',
            retryable: false,
        }),
    });
    const line = kbUploadErrorLine(tree);
    assert.match(line, /12582912 字节/);
    assert.match(line, /HTTP 413/);
    assert.match(line, /body_too_large/);
    assert.equal(line.includes('OCR 端点'), false);
});

await check('知识库上传：网关 HTML 判成「HTML/XML 页面」，不回显页面原文', async () => {
    const { tree } = await kbUploadFailureScenario({
        status: 502,
        bodyText: '<html><body><h1>502 Bad Gateway</h1></body></html>',
        contentType: 'text/html',
    });
    const line = kbUploadErrorLine(tree);
    assert.match(line, /HTML\/XML/);
    assert.match(line, /HTTP 502/);
    assert.match(line, /text\/html/);
    assert.equal(line.includes('Bad Gateway') || line.includes('<html'), false,
        '响应体原文（可能是网关内部页）不能进面板');
});

await check('知识库：OCR 结果存为笔记时命中 kb-upload 失败，文案指向知识库上传端点而不是 OCR', async () => {
    let uploads = 0;
    const fetchImpl = (url, init) => {
        const target = String(url);
        if (target.includes('/office-memory/kb-upload')) {
            uploads += 1;
            if (JSON.parse(String((init && init.body) || '{}')).preview === true) {
                return jsonResponse({
                    ok: true, action: 'kb-preview', result: {
                        action: 'kb-preview', kind: 'scan', name: '扫描件.png', ext: '.png', bytes: 2048,
                        chars: 0, chunks: 1, head: '', truncated: false, images: 1,
                        note: '这份文件没有可抽取的文本层（是扫描件）：要读文字请点「识别文字」。',
                    },
                });
            }
            // 第二次 kb-upload 是「存为知识库笔记」：模拟有界关闭前的空 413。
            return Promise.resolve({
                ok: false, status: 413,
                headers: { get: () => 'application/json; charset=utf-8' },
                text: () => Promise.resolve(''),
            });
        }
        if (target.includes('/office-memory/ocr')) {
            return Promise.resolve({
                ok: true, status: 200,
                headers: { get: () => 'application/json; charset=utf-8' },
                text: () => Promise.resolve(JSON.stringify({ ok: true, fileName: '扫描件.png', pageCount: 1, text: '识别正文', usage: null, bytes: 8, ms: 5 })),
            });
        }
        return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) });
    };
    const restore = kbFileReaderFake();
    try {
        const { mod, ctx, registrations } = loadBundle({ fetch: fetchImpl });
        mod.apply(ctx);
        const panel = regOf(registrations, 'conversation.view', 'office-memory');
        let tree = await renderPanelLoaded(panel);
        await new Promise((resolve) => setTimeout(resolve, 10));
        walk(tree).find((n) => n.props && n.props.role === 'tab' && n.props['data-tab'] === 'kb').props.onClick();
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '扫描件.png', __body: 'png-bytes' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ocr', '扫描件.png')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-ocr-save', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(uploads, 2, '预览一次 + 存笔记一次');
        const line = textsOf(tree).find((text) => text.includes('存进知识库失败'));
        assert.ok(line, '存笔记失败要有回执');
        assert.match(line, /知识库上传端点/);
        assert.match(line, /HTTP 413/);
        assert.equal(line.includes('OCR 端点'), false, '这条通道是 kb-upload，不能叫 OCR 端点');
    } finally {
        restore();
    }
});


// ── 上传 UI：目标文件夹锁定 / 回执如实 / 预估通道诊断 ─────────────────────
//
// 这一组只钉 UI 契约：预览时捕获的 folder 一路用到预估与真实入库；回执只说
// 服务端结构化字段（错误只报条数）；available:false 才禁用开始入库，duplicate
// 仍允许 0 次模型调用的清单更新。

/** 宿主模型清单：让 host 通道的模型下拉有真实选项（RPC，不走 HTTP）。 */
function kbCatalogRemote() {
    return { session: {
        modelCatalog: async () => ({
            ok: true,
            value: {
                groups: [{ id: 'p', name: 'P', models: [
                    { id: 'm1', name: 'Model One' },
                    { id: 'm2', name: 'Model Two' },
                ] }],
                failures: [],
            },
        }),
    } };
}

await check('知识库：预览 / 预估 / 入库锁定同一目标文件夹；AI 回执报调用、降级、摘要、错误条数而不回显原文', async () => {
    const previews = [];
    const ingests = [];
    const snapshot = kbAiSnapshot();
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                previews.push(body);
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            ingests.push(body);
            if (body.aiEstimate === true) {
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 1200,
                    estimate: { units: 3, batches: 2, calls: 2, maxCalls: 8, estimatedInputTokens: 1800, estimatedOutputTokens: 240, overBudget: false },
                    ai: { mode: 'host', available: true, calls: 0 },
                    note: '确认后才会写入。',
                } });
            }
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: '来源/笔记.txt', chars: 1200, chunks: 3,
                ai: {
                    mode: 'host', provider: 'p', model: 'm1', chunking: true, summary: true,
                    available: true, calls: 3, degraded: 1, summaries: 2,
                    errors: ['模型端点返回 HTTP 500：RAW_SECRET_BODY'],
                    note: '部分批次已降级',
                },
            } });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        let tree = render(panel);
        byProp(tree, 'data-kb-folder', '来源')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(previews.length, 1);
        assert.equal(previews[0].folder, '来源', '预览请求要带捕获的文件夹');
        assert.equal(byProp(tree, 'data-kb-preview-folder', '来源').length, 1, '预览卡要写入库目标');
        // 回到根目录再勾 AI：预估与真实入库仍必须写「来源」。
        byProp(tree, 'data-kb-folder', '')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 1, '点入库只发一次预估');
        assert.equal(ingests[0].aiEstimate, true);
        assert.equal(ingests[0].folder, '来源', '预估不能读当前目录');
        assert.equal(byProp(tree, 'data-kb-estimate-folder', '来源').length, 1, '预估卡要写目标文件夹');
        // 再切到子层：开始入库仍发 pending 里的「来源」。
        byProp(tree, 'data-kb-folder', '来源/方法论')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-estimate-start', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 2);
        assert.equal(ingests[1].folder, '来源', '真实入库沿用捕获的文件夹');
        assert.equal(ingests[1].ai, true);
        assert.equal(ingests[1].aiEstimate, false);
        const texts = textsOf(tree).join('\n');
        assert.ok(texts.includes('进入「来源」'), '回执写实际进入的文件夹');
        assert.ok(texts.includes('AI 建库') && texts.includes('模型调用 3 次'), 'AI 回执有调用次数');
        assert.ok(texts.includes('摘要 2 条'), 'AI 回执有摘要条数');
        assert.ok(texts.includes('1 批降级为确定性切块'), '降级如实报');
        assert.ok(texts.includes('错误 1 条'), '错误只报条数');
        assert.ok(texts.includes('说明：部分批次已降级'), '服务端说明要附上');
        assert.equal(texts.includes('RAW_SECRET_BODY'), false, '模型端点原文不能上屏');
    } finally {
        restore();
    }
});

await check('知识库：模型入库按 NDJSON 流画进度条，完成后清掉（第一百一十轮）', async () => {
    const snapshot = kbAiSnapshot();
    let finalRequest = null;
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) return jsonResponse({ ok: true, providers: [], error: '' });
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            if (body.aiEstimate === true) {
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 1200,
                    estimate: { units: 3, batches: 2, calls: 2, maxCalls: 8, estimatedInputTokens: 1800, estimatedOutputTokens: 240, overBudget: false },
                    ai: { mode: 'host', available: true, calls: 0 }, note: '确认后才会写入。',
                } });
            }
            finalRequest = body;
            return ndjsonResponse([
                { type: 'progress', phase: 'extract' },
                { type: 'progress', phase: 'chunks', chars: 1200 },
                { type: 'progress', phase: 'plan', doneBatches: 0, totalBatches: 2, calls: 0, degraded: 0, truncated: 0 },
                { type: 'progress', phase: 'model', doneBatches: 1, totalBatches: 2, calls: 1, degraded: 0, truncated: 0 },
                { type: 'progress', phase: 'model', doneBatches: 2, totalBatches: 2, calls: 2, degraded: 1, truncated: 1 },
                { type: 'result', ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', folder: body.folder, path: '笔记.txt', chars: 1200, chunks: 3,
                    ai: { mode: 'host', available: true, calls: 2, degraded: 1, summaries: 1, errors: [] },
                } },
            ], 40);
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        let tree = render(panel);
        byProp(tree, 'data-kb-folder', '来源')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-estimate-start', 'true')[0].props.onClick();
        // 每个事件前隔 40ms：轮询到 1/2 批的中间态再断言，不赌单次渲染时机。
        let bar = byProp(tree, 'data-kb-progress', 'model')[0];
        for (let waited = 0; bar === undefined && waited < 12; waited += 1) {
            await new Promise((resolve) => setTimeout(resolve, 20));
            tree = render(panel);
            bar = byProp(tree, 'data-kb-progress', 'model')[0];
        }
        assert.ok(bar, '模型入库时应画出进度条');
        assert.ok(finalRequest, '应已发出真实入库请求');
        assert.equal(bar.props.role, 'progressbar');
        assert.equal(bar.props['aria-valuenow'], 65, '1/2 批大约是 30% 起的 65%');
        const barText = textsOf(bar).join(' ');
        assert.match(barText, /模型 1\/2 批/);
        assert.match(barText, /已调用 1 次/);
        assert.equal(finalRequest.stream, true, '真实入库要请求流式响应');
        await new Promise((resolve) => setTimeout(resolve, 200));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-progress', 'model').length, 0, '完成后进度条要清掉');
        assert.ok(textsOf(tree).join('\n').includes('AI 建库'), '完成后仍要有入库回执');
    } finally {
        restore();
    }
});

await check('知识库：服务端忽略 stream 时退回整段 JSON（第一百一十轮）', async () => {
    const snapshot = kbAiSnapshot();
    let finalRequest = null;
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) return jsonResponse({ ok: true, providers: [], error: '' });
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            if (body.aiEstimate === true) {
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 1200,
                    estimate: { units: 3, batches: 2, calls: 2, maxCalls: 8, estimatedInputTokens: 1800, estimatedOutputTokens: 240, overBudget: false },
                    ai: { mode: 'host', available: true, calls: 0 }, note: '确认后才会写入。',
                } });
            }
            finalRequest = body;
            // 旧服务端不认 stream：仍回 application/json 的整段结果。真实 fetch 的响应
            // 一定有 body 流，所以这里也带 body；客户端必须看 content-type 才决定按行读。
            const payload = { ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: '笔记.txt', chars: 1200, chunks: 3,
                ai: { mode: 'host', available: true, calls: 2, degraded: 0, summaries: 1, errors: [] },
            } };
            const text = JSON.stringify(payload);
            return Promise.resolve({
                ok: true, status: 200,
                headers: { get: () => 'application/json; charset=utf-8' },
                body: new ReadableStream({
                    start(controller) {
                        controller.enqueue(new TextEncoder().encode(text));
                        controller.close();
                    },
                }),
                text: () => Promise.resolve(text),
                json: () => Promise.resolve(payload),
            });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        let tree = render(panel);
        byProp(tree, 'data-kb-folder', '来源')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-estimate-start', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 120));
        tree = render(panel);
        assert.ok(finalRequest, '应已发出真实入库请求');
        assert.equal(finalRequest.stream, true, '客户端仍请求流式响应');
        assert.equal(byProp(tree, 'data-kb-progress', 'model').length, 0, '旧服务端不应留下进度条');
        assert.ok(textsOf(tree).join('\n').includes('AI 建库'), '旧服务端整段 JSON 也要能收尾');
    } finally {
        restore();
    }
});

await check('知识库：选模型不会自动启用 AI；普通入库回执明说未调用模型', async () => {
    const ingests = [];
    const snapshot = kbAiSnapshot();
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            ingests.push(body);
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: '笔记.txt', chars: 1200, chunks: 3,
            } });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel, tree: first } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        const select = byProp(first, 'data-kb-ai-model', 'true')[0];
        assert.ok(select, 'host 通道要有模型下拉');
        assert.ok(textsOf(first).some((t) => t.includes('选模型不会自动开启')), '要说明选模型不等于启用');
        select.props.onChange({ target: { value: 'p/m1' } });
        let tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.checked, false, '选模型后 AI 开关仍关闭');
        assert.ok(textsOf(tree).some((t) => t.includes('选模型不会自动开启')), '未勾选时提示仍在');
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 1);
        assert.equal(ingests[0].ai, false, '选过模型但没勾开关，入库仍是普通');
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 0, '不勾 AI 就不画预估');
        const texts = textsOf(tree).join('\n');
        assert.ok(texts.includes('已入库：笔记.txt'));
        assert.ok(texts.includes('进入「全部文档」'));
        assert.ok(texts.includes('确定性切块：未调用模型、没有模型摘要'), '普通回执要区分 AI');
    } finally {
        restore();
    }
});

await check('知识库：OCR 正文写回源文件自己的文档，替换旧结果并说明旧笔记会收掉', async () => {
    const saves = [];
    const snapshot = makeSnapshot();
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'scan', name: '扫描件.png', ext: '.png', bytes: 2048,
                    chars: 0, chunks: 1, head: '', truncated: false, images: 1,
                    note: '这份文件没有可抽取的文本层（是扫描件）：要读文字请点「识别文字」。',
                } });
            }
            saves.push(body);
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: 'uploads/来源/扫描件.png', chars: 8, chunks: 1,
                ocr: true, removedOcrNote: 'uploads/来源/扫描件.ocr.md',
            } });
        }
        if (text.includes('/office-memory/ocr')) {
            return Promise.resolve({
                ok: true, status: 200,
                headers: { get: () => 'application/json; charset=utf-8' },
                text: () => Promise.resolve(JSON.stringify({ ok: true, fileName: '扫描件.png', pageCount: 1, text: '识别正文', usage: null, bytes: 8, ms: 5 })),
            });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl });
        let tree = render(panel);
        byProp(tree, 'data-kb-folder', '来源')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '扫描件.png', __body: 'png-bytes' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ocr', '扫描件.png')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-ocr-target', '来源').length, 1, 'OCR 面板要写清存到哪');
        const ocrTexts = textsOf(tree).join('\n');
        assert.ok(ocrTexts.includes('替换它之前的识别结果'), '面板要说明会替换旧识别结果');
        assert.ok(ocrTexts.includes('旧笔记会一并收掉'), '面板要说明旧 .ocr.md 笔记会被收掉');
        assert.equal(byProp(tree, 'data-kb-ocr-ai', 'true').length, 1, '面板要能选择用模型切片 / 摘要');
        // 换回根目录再保存：仍用预览时捕获的「来源」。
        byProp(tree, 'data-kb-folder', '')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-ocr-save', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(saves.length, 1);
        assert.equal(saves[0].folder, '来源', '存正文沿用 OCR 预览文件夹');
        assert.equal(saves[0].name, '扫描件.png', '落点是源文件自己的文档，不改名成 .ocr.md');
        assert.equal(saves[0].ocr, true, '正文要带 ocr 标记');
        assert.equal('data' in saves[0], false, '正文通道不再重传二进制');
        assert.equal(typeof saves[0].text, 'string', '正文用 base64 的 text 字段');
        assert.equal(Buffer.from(saves[0].text, 'base64').toString('utf8'), '识别正文');
        const texts = textsOf(tree).join('\n');
        assert.ok(texts.includes('已入库：扫描件.png'));
        assert.ok(texts.includes('进入「来源」'));
        assert.ok(texts.includes('正文来自 OCR 识别结果'), '回执要写明正文来自 OCR');
        assert.ok(texts.includes('uploads/来源/扫描件.ocr.md 已一并收掉'), '回执要说清旧笔记的去向');
        assert.ok(texts.includes('确定性切块：未调用模型、没有模型摘要'), '没勾模型时回执要明说没有模型切片 / 摘要');
    } finally {
        restore();
    }
});

await check('知识库：识别完成后待入库卡换成 OCR 正文，点「入库」写的是识别结果', async () => {
    const saves = [];
    const snapshot = makeSnapshot();
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'scan', name: '扫描件.pdf', ext: '.pdf', bytes: 4096,
                    chars: 0, chunks: 1, head: '（脏文本层）', truncated: false, images: 0,
                    note: '这份文件没有可抽取的文本层（是扫描件）：要读文字请点「识别文字」。',
                } });
            }
            saves.push(body);
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: 'uploads/来源/扫描件.pdf', chars: 8, chunks: 1, ocr: true,
            } });
        }
        if (text.includes('/office-memory/ocr')) {
            return Promise.resolve({
                ok: true, status: 200,
                headers: { get: () => 'application/json; charset=utf-8' },
                text: () => Promise.resolve(JSON.stringify({ ok: true, fileName: '扫描件.pdf', pageCount: 1, text: '准确的识别正文', usage: null, bytes: 8, ms: 5 })),
            });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl });
        let tree = render(panel);
        byProp(tree, 'data-kb-folder', '来源')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '扫描件.pdf', __body: 'pdf-bytes' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-preview-kind', 'scan').length, 1, '识别前是扫描件卡');
        byProp(tree, 'data-kb-preview-ocr', '扫描件.pdf')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-preview-kind', 'ocr').length, 1, '识别后卡要换成「识别文字」');
        assert.equal(byProp(tree, 'data-kb-preview-head', 'true').length, 1, '识别后卡上要显示识别正文');
        const head = byProp(tree, 'data-kb-preview-head', 'true')[0];
        assert.ok(textsOf(head).join('').includes('准确的识别正文'), '卡上的正文必须是识别结果，不是脏文本层');
        assert.ok(textsOf(tree).join('\n').includes('上面已换成识别结果'));
        byProp(tree, 'data-kb-preview-ingest', '扫描件.pdf')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        assert.equal(saves.length, 1, '点「入库」只发一次');
        assert.equal(saves[0].name, '扫描件.pdf', '仍写源文件名');
        assert.equal(saves[0].folder, '来源', '目标文件夹沿用预览时捕获的');
        assert.equal(saves[0].ocr, true);
        assert.equal('data' in saves[0], false, '入库写的是识别正文，不重传二进制');
        assert.equal(Buffer.from(saves[0].text, 'base64').toString('utf8'), '准确的识别正文');
    } finally {
        restore();
    }
});

await check('知识库：勾了模型时 OCR 正文先预估，确认后原样带 text/ocr 入库', async () => {
    const calls = [];
    const snapshot = kbAiSnapshot();
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) return jsonResponse({ ok: true, providers: [], error: '' });
        if (text.includes('/office-memory/ocr')) {
            return Promise.resolve({
                ok: true, status: 200,
                headers: { get: () => 'application/json; charset=utf-8' },
                text: () => Promise.resolve(JSON.stringify({ ok: true, fileName: '扫描件.png', pageCount: 1, text: '识别正文', usage: null, bytes: 8, ms: 5 })),
            });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'scan', name: '扫描件.png', ext: '.png', bytes: 2048,
                    chars: 0, chunks: 1, head: '', truncated: false, images: 1,
                    note: '这份文件没有可抽取的文本层（是扫描件）：要读文字请点「识别文字」。',
                } });
            }
            calls.push(body);
            if (body.aiEstimate === true) {
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 4,
                    estimate: { units: 1, batches: 1, calls: 1, maxCalls: 8, estimatedInputTokens: 10, estimatedOutputTokens: 5, overBudget: false },
                    ai: { mode: 'host', available: true, chunking: true, summary: true, calls: 0, summaries: 0 },
                    duplicate: false,
                } });
            }
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: 'uploads/扫描件.png', chars: 4, chunks: 1,
                ocr: true,
                ai: { mode: 'host', available: true, chunking: true, summary: true, calls: 1, summaries: 1, degraded: 0, errors: [] },
            } });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl });
        let tree = render(panel);
        // 先勾 AI：OCR 面板上的勾选与顶部开关是同一个状态（面板里也放了一份）。
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '扫描件.png', __body: 'png-bytes' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ocr', '扫描件.png')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-ocr-ai', 'true')[0].props.children[0].props.checked, true,
            'OCR 面板的勾选要跟着全局 AI 开关');
        byProp(tree, 'data-kb-ocr-save', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(calls.length, 1, '勾了模型先只发一次预估');
        assert.equal(calls[0].aiEstimate, true);
        assert.equal(calls[0].ocr, true, '预估也要带 ocr 标记');
        assert.equal(typeof calls[0].text, 'string', '预估用正文而不是二进制');
        assert.equal('data' in calls[0], false, 'OCR 正文不重传二进制');
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 1, '要画预估卡');
        const start = byProp(tree, 'data-kb-estimate-start', 'true')[0];
        start.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(calls.length, 2, '确认后真的入库');
        assert.equal(calls[1].aiEstimate, false);
        assert.equal(calls[1].ocr, true, '确认请求保留 ocr 标记');
        assert.equal(Buffer.from(calls[1].text, 'base64').toString('utf8'), '识别正文');
        const texts = textsOf(tree).join('\n');
        assert.ok(texts.includes('正文来自 OCR 识别结果'), '回执要标明正文来自 OCR');
        assert.ok(texts.includes('模型调用 1 次'), '回执要写清模型调用次数');
    } finally {
        restore();
    }
});

await check('知识库：预估通道不可用时显示诊断并禁用开始入库，程序化触发也不发真实请求', async () => {
    const ingests = [];
    const snapshot = kbAiSnapshot();
    const reason = '宿主没有可用的 llm 服务（host 通道的 ctx.llm 不可达）。';
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            ingests.push(body);
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 1200,
                estimate: { units: 3, batches: 0, calls: 0, maxCalls: 8, estimatedInputTokens: 0, estimatedOutputTokens: 0, overBudget: false },
                ai: { mode: 'host', available: false, calls: 0, reason: reason },
                note: reason + '这次不会发起任何模型请求；改选 openai 兼容通道，或先在宿主里配置模型后再预估。',
            } });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        let tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 1, '预估卡要出来');
        assert.equal(byProp(tree, 'data-kb-estimate-unavailable', 'true').length, 1, '不可用要有诊断');
        const texts = textsOf(tree).join('\n');
        assert.ok(texts.includes(reason), '诊断要带服务端 reason');
        assert.ok(texts.includes('不会发起任何模型请求'), '要说明不会发请求');
        const start = byProp(tree, 'data-kb-estimate-start', 'true')[0];
        assert.equal(start.props.disabled, true, 'available:false 才禁用开始入库');
        assert.ok(textsOf(start).some((t) => t.includes('通道不可用')), '禁用时按钮文案要说明原因');
        start.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        assert.equal(ingests.length, 1, '禁用后不得再发真实入库请求');
    } finally {
        restore();
    }
});

await check('知识库：预估 duplicate 时允许 0 次模型调用的清单更新，不被当成通道不可用', async () => {
    const ingests = [];
    const snapshot = kbAiSnapshot();
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            ingests.push(body);
            if (body.aiEstimate === true) {
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 1200,
                    estimate: { units: 3, batches: 0, calls: 0, maxCalls: 8, estimatedInputTokens: 0, estimatedOutputTokens: 0, overBudget: false },
                    ai: { mode: 'host', available: true, calls: 0, summaries: 0 },
                    duplicate: true,
                    note: '同一版内容与切块档已经在库里：实际入库会直接回执「未改动」，0 次模型调用；来源档 / 文件夹的重入只改清单，不重跑模型。',
                } });
            }
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: '笔记.txt', unchanged: true, moved: true,
                ai: { mode: 'host', available: true, chunking: false, summary: false, calls: 0, summaries: 0 },
            } });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        let tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-estimate-duplicate', 'true').length, 1, 'duplicate 要有诊断');
        assert.equal(byProp(tree, 'data-kb-estimate-unavailable', 'true').length, 0, 'duplicate 不是通道不可用');
        const start = byProp(tree, 'data-kb-estimate-start', 'true')[0];
        assert.equal(start.props.disabled, false, 'duplicate 仍允许入库（0 次模型调用）');
        assert.ok(String(start.props.children).includes('0 次模型调用'), '按钮要说明 0 次模型调用');
        start.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 2, 'duplicate 也要真的发一次入库');
        assert.equal(ingests[1].ai, true);
        assert.equal(ingests[1].aiEstimate, false);
        const texts = textsOf(tree).join('\n');
        assert.ok(texts.includes('内容没变，只把文件夹改到「全部文档」'), '回执如实说未重写');
        assert.ok(texts.includes('模型调用 0 次'), '回执如实说 0 次调用');
    } finally {
        restore();
    }
});

await check('知识库：预估捕获模型，之后换模型不影响入库与续跑', async () => {
    const ingests = [];
    const snapshot = kbAiSnapshot();
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: '笔记.txt', ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            ingests.push(body);
            if (body.aiEstimate === true) {
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 1200,
                    estimate: { units: 3, batches: 0, calls: 0, maxCalls: 8, estimatedInputTokens: 0, estimatedOutputTokens: 0, overBudget: false },
                    ai: { mode: 'host', available: true, calls: 0, summaries: 0 },
                    duplicate: true,
                    note: '同一版内容与切块档已经在库里：0 次模型调用。',
                } });
            }
            if (body.aiResume === true) {
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', folder: body.folder, path: '笔记.txt', chars: 1200, chunks: 3,
                    ai: { mode: 'host', available: true, chunking: false, summary: false, calls: 0, summaries: 0, degraded: 0, errors: [] },
                } });
            }
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: '笔记.txt', paused: true, status: 'paused',
                calls: 2, doneBatches: 1, batches: 3, note: '本轮到预算上限',
            } });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel, tree: first } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        byProp(first, 'data-kb-ai-model', 'true')[0].props.onChange({ target: { value: 'p/m1' } });
        let tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests[0].aiModel, 'p/m1', '预估用当时选中的模型');
        assert.equal(byProp(tree, 'data-kb-estimate-model', 'p/m1').length, 1, '预估卡要标出模型');
        assert.ok(textsOf(tree).some((text) => text.includes('模型：p/m1')), '模型标签可读');
        // 看过预估后改选另一个模型：真实入库和续跑都必须仍用 p/m1。
        byProp(tree, 'data-kb-ai-model', 'true')[0].props.onChange({ target: { value: 'p/m2' } });
        tree = render(panel);
        const start = byProp(tree, 'data-kb-estimate-start', 'true')[0];
        // 连点两次：第二次必须被 kbUploading 守卫挡住，只发一次真实入库。
        start.props.onClick();
        start.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 2, '连点开始入库也只发一次真实请求');
        assert.equal(ingests[1].aiModel, 'p/m1', '换模型后真实入库仍用预估时捕获的模型');
        assert.equal(ingests[1].aiResume, false);
        assert.equal(byProp(tree, 'data-kb-paused', 'true').length, 1, '预算暂停要出现续跑卡');
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 0, '暂停后要撤掉预估卡，避免旧按钮再跑一遍');
        // 暂停前的旧 Start 回调也不能再发请求（aiResume:false 会从游标 0 重跑）。
        start.props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        assert.equal(ingests.length, 2, '暂停后旧 Start 回调不能再发请求');
        // 同一张预览卡再点「入库」也不能绕开暂停重新预估。
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 2, '暂停期间预览卡不能重新发起预估');
        assert.ok(textsOf(tree).some((text) => text.includes('已有暂停的入库任务')), '要提示先续跑或取消');
        byProp(tree, 'data-kb-resume', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 3);
        assert.equal(ingests[2].aiResume, true);
        assert.equal(ingests[2].aiModel, 'p/m1', '续跑不能改用新选中的模型');
    } finally {
        restore();
    }
});

await check('知识库：暂停期间预览卡不能重启、显式取消后才允许重新预估', async () => {
    const ingests = [];
    const snapshot = kbAiSnapshot();
    let estimates = 0;
    const fetchImpl = (url, init) => {
        const text = String(url);
        if (text.includes('/office-memory/snapshot')) return jsonResponse(snapshot);
        if (text.includes('/office-memory/llm-models')) {
            return jsonResponse({ ok: true, providers: [], error: '' });
        }
        if (text.includes('/office-memory/kb-upload')) {
            const body = JSON.parse(String((init && init.body) || '{}'));
            if (body.preview === true) {
                const name = String(body.name || '笔记.txt');
                return jsonResponse({ ok: true, action: 'kb-preview', result: {
                    action: 'kb-preview', kind: 'text', name: name, ext: '.txt', bytes: 18,
                    chars: 1200, chunks: 3, head: '一段正文', truncated: false, images: 0, note: '',
                } });
            }
            ingests.push(body);
            if (body.aiEstimate === true) {
                estimates += 1;
                return jsonResponse({ ok: true, action: 'kb-upload', result: {
                    action: 'kb-ingest', dryRun: true, folder: body.folder, chars: 1200,
                    estimate: { units: 3, batches: estimates === 1 ? 0 : 2, calls: estimates === 1 ? 0 : 2, maxCalls: 8, estimatedInputTokens: 0, estimatedOutputTokens: 0, overBudget: false },
                    ai: { mode: 'host', available: true, calls: 0, summaries: 0 },
                    duplicate: estimates === 1,
                    note: '同一版内容与切块档已经在库里：0 次模型调用。',
                } });
            }
            return jsonResponse({ ok: true, action: 'kb-upload', result: {
                action: 'kb-ingest', folder: body.folder, path: '笔记.txt', paused: true, status: 'paused',
                calls: 2, doneBatches: 1, batches: 3, note: '本轮到预算上限',
            } });
        }
        return jsonResponse(snapshot);
    };
    const restore = kbFileReaderFake();
    try {
        const { panel } = await openKbAiPanel({ fetch: fetchImpl, remote: kbCatalogRemote() });
        let tree = render(panel);
        byProp(tree, 'data-kb-ai-toggle', 'true')[0].props.onChange({ target: { checked: true } });
        tree = render(panel);
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '笔记.txt', __body: '一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-preview-ingest', '笔记.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        byProp(tree, 'data-kb-estimate-start', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-paused', 'true').length, 1);
        assert.equal(byProp(tree, 'data-kb-estimate-start', 'true').length, 0, '暂停后不能再留开始按钮');
        // 暂停期间重新选文件：暂停任务的内存正文要保留，续跑按钮不能失效。
        byProp(tree, 'data-kb-upload-input', 'true')[0].props.onChange({ target: { files: [{ name: '另一.txt', __body: '另一段正文' }] } });
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-paused', 'true').length, 1, '新预览不能顶掉暂停卡');
        assert.equal(byProp(tree, 'data-kb-resume', 'true')[0].props.disabled, false, '暂停任务续跑内容仍在');
        byProp(tree, 'data-kb-preview-ingest', '另一.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 2, '暂停期间新预览卡也不能发预估');
        assert.ok(textsOf(tree).some((text) => text.includes('已有暂停的入库任务')), '要引导用户续跑或取消');
        // 显式取消后才把暂停任务丢掉，允许从预览卡重新开始。
        byProp(tree, 'data-kb-paused-cancel', 'true')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 10));
        tree = render(panel);
        assert.equal(byProp(tree, 'data-kb-paused', 'true').length, 0, '取消后暂停卡要收掉');
        assert.ok(textsOf(tree).some((text) => text.includes('已取消暂停的入库任务')), '取消要有回执');
        byProp(tree, 'data-kb-preview-ingest', '另一.txt')[0].props.onClick();
        await new Promise((resolve) => setTimeout(resolve, 30));
        tree = render(panel);
        assert.equal(ingests.length, 3, '取消后允许重新预估');
        assert.equal(ingests[2].aiEstimate, true);
        assert.equal(byProp(tree, 'data-kb-estimate', 'true').length, 1, '重新预估要出卡');
    } finally {
        restore();
    }
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`client: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);