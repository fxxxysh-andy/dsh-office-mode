/**
 * 知识库范围选择器的配色契约：宿主 --dsw-alias-* token + 浅/深两套的实测对比度。
 *
 * 修的是什么（用户报回）：重启 3080 后，输入框左侧的「知识库范围」选择器
 * 在深色下**整块看不见** —— 面板底色写的是 var(--surface, #fff)，而宿主
 * （@deepseek-ai/dsh-client-ui-theme 的 bundle）从来没有定义过 --surface，
 * 于是 fallback 的 #fff 永远生效；文字用 color:inherit 继承到
 * --dsw-alias-label-primary（深色 #f9fafb）。白底浅字 = 1.04:1。
 * 同一处还硬编码了 rgba(0,0,0,0.6)（说明行）与 #b3261e（错误行），
 * 深色下同样读不出来；选中 / 未选中 / 悬停 / 禁用四态此前没有任何视觉差别。
 *
 * 这个套件钉住四件事：
 *   1. 入口按钮、面板、行、提示条的颜色一律走宿主 token（不再有 --surface、
 *      不再有裸 hex 主色、不用 backdrop-filter 糊一层）；
 *   2. 面板自己显式给 color（不靠从输入行继承）且底色不透明；
 *   3. 选中 / 悬停 / 禁用三态在渲染出的 style 上真的不同，且 aria-pressed
 *      与「看起来被选中」一致；
 *   4. 浅色、深色两套的真实色值算 WCAG 对比度：正文/说明 ≥ 4.5:1，
 *      禁用与左侧语义竖条 ≥ 3:1（禁用态 WCAG 豁免，但仍要能辨认）。
 *
 * 色值来源：宿主安装包里的
 *   node_modules/@deepseek-ai/dsh-client-ui-theme/lib/client.js
 * 中 body{...} 与 body[data-ds-dark-theme]{...} 两段。下面的 FIXTURE 是当时的
 * 解析结果；设了 DSH_THEME_BUNDLE=<该文件路径> 时，本套件会现场重解析并核对
 * fixture 没漂（CI 里没这个文件就跳过核对，只跑 fixture 上的对比度）。
 *
 * 跑法：node test/kb-scope-contrast.mjs
 */
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

// 默认测包里的这份 bundle；`--bundle=<path>` / DSH_SCOPE_BUNDLE 指到别处时改测那份 ——
// 用来做「撤销修复后本套件应当转红」的反证，免得为了验证测试有效去改真正的源文件。
const ARGV = process.argv.slice(2);
const argOf = (name) => {
    const hit = ARGV.find((item) => item.startsWith(name + '='));
    return hit === undefined ? '' : hit.slice(name.length + 1);
};
const BUNDLE = process.env.DSH_SCOPE_BUNDLE || argOf('--bundle') || new URL('../lib/client.js', import.meta.url);

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

// ── 宿主 token 的实际值（浅色 / 深色） ───────────────────────────────────
//
// 只列本选择器用到的 11 个别名。深色块里没被覆盖的（state-warn）仍是浅色那份值；
// 语义色只配给左侧 3px 竖条、不当小字正文，也是因为其中几个自身对比度不够。
const FIXTURE = {
    light: {
        '--dsw-alias-label-primary': '#0f1115',
        '--dsw-alias-label-secondary': '#61666b',
        '--dsw-alias-label-tertiary': '#81858c',
        '--dsw-alias-bg-layer-1': '#ffffff',
        '--dsw-alias-border-l2': '#0000001a',
        '--dsw-alias-interactive-bg-hover': '#2631480f',
        '--dsw-alias-interactive-bg-active': '#2631481a',
        '--dsw-alias-state-business-primary': '#4176e6',
        '--dsw-alias-state-error-primary': '#ec1313',
        '--dsw-alias-state-warn-primary': '#f59e0b',
        '--dsw-alias-bg-mask-2': '#0000001f',
    },
    dark: {
        '--dsw-alias-label-primary': '#f9fafb',
        '--dsw-alias-label-secondary': '#cfd3d6',
        '--dsw-alias-label-tertiary': '#adb2b8',
        '--dsw-alias-bg-layer-1': '#232324',
        '--dsw-alias-border-l2': '#ffffff1f',
        '--dsw-alias-interactive-bg-hover': '#ffffff14',
        '--dsw-alias-interactive-bg-active': '#ffffff24',
        '--dsw-alias-state-business-primary': '#7aaaff',
        '--dsw-alias-state-error-primary': '#f25a5a',
        '--dsw-alias-state-warn-primary': '#f59e0b',
        '--dsw-alias-bg-mask-2': '#0003',
    },
};

// ── 颜色小工具 ───────────────────────────────────────────────────────────

/** '#rgb' / '#rrggbb' / '#rrggbbaa' → [r, g, b, a]。 */
function parseHex(value) {
    const text = String(value).trim().replace(/^#/, '');
    const full = text.length === 3 || text.length === 4 ? text.split('').map((c) => c + c).join('') : text;
    const at = (i) => parseInt(full.slice(i, i + 2), 16);
    return [at(0), at(2), at(4), full.length === 8 ? at(6) / 255 : 1];
}

/** 把带透明度的前景合成到不透明底色上（返回 [r, g, b]）。 */
function compositeOver(foreground, background) {
    const a = foreground[3];
    return [0, 1, 2].map((i) => foreground[i] * a + background[i] * (1 - a));
}

function relativeLuminance(rgb) {
    const linear = rgb.slice(0, 3).map((channel) => {
        const c = channel / 255;
        return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
    });
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
}

/** WCAG 2.x 对比度。 */
function contrast(foreground, background) {
    const a = relativeLuminance(foreground);
    const b = relativeLuminance(background);
    return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

/** 两个不透明色的最大通道差（用来判「两种状态分不分得出来」，对比度对同亮度色不敏感）。 */
function channelDelta(one, other) {
    return Math.max(...[0, 1, 2].map((i) => Math.abs(one[i] - other[i])));
}

/** 把 #fff / #0003 这类简写归一成 #rrggbb[aa]，只用于对账比较。 */
function normalizeHex(value) {
    const [r, g, b, a] = parseHex(value);
    const pair = (n) => n.toString(16).padStart(2, '0');
    return '#' + pair(r) + pair(g) + pair(b) + (a === 1 ? '' : pair(Math.round(a * 255)));
}

/** 解析 style 值里的第一个 --dsw-alias-* token。 */
function tokenOf(value) {
    const matched = /var\((--dsw-alias-[a-z0-9-]+)/.exec(String(value));
    return matched === null ? null : matched[1];
}

/** 断言某个样式值走的是宿主 token（并且这个 token 在宿主里真的有定义）。 */
function assertTokenStyle(value, theme, label) {
    const token = tokenOf(value);
    assert.ok(token !== null, label + ' 必须用宿主 --dsw-alias token，实际：' + String(value));
    assert.ok(Object.prototype.hasOwnProperty.call(FIXTURE[theme], token), label + ' 用了宿主没定义过的 token：' + token);
    return token;
}

/** token 在某个主题下的实际颜色（[r,g,b,a]）。 */
function themeColor(theme, value) {
    const token = tokenOf(value);
    assert.ok(token !== null, '不是 token：' + String(value));
    return parseHex(FIXTURE[theme][token]);
}

// ── 假 React（只到能渲染 KbScopePicker 的程度） ───────────────────────────

function makeFakeReact() {
    const store = new Map();
    let component = null;
    let cursor = 0;
    let occurrences = new Map();
    let dirty = false;
    const React = {
        createElement: (type, props, ...children) => ({
            $typeof: true,
            type,
            props: Object.assign({}, props || {}, children.length > 0 ? { children } : {}),
        }),
        useState: (initial) => {
            const key = component + ':' + cursor;
            cursor += 1;
            if (!store.has(key)) store.set(key, { value: initial });
            const slot = store.get(key);
            return [slot.value, (next) => {
                const value = typeof next === 'function' ? next(slot.value) : next;
                if (Object.is(value, slot.value)) return;
                slot.value = value;
                dirty = true;
            }];
        },
        useRef: (initial) => {
            const key = component + ':ref:' + cursor;
            cursor += 1;
            if (!store.has(key)) store.set(key, { value: { current: initial } });
            return store.get(key).value;
        },
        useEffect: (fn) => { fn(); return undefined; },
    };
    return {
        React,
        resetCursor: () => { cursor = 0; },
        resetInstances: () => { occurrences = new Map(); },
        isDirty: () => dirty,
        clearDirty: () => { dirty = false; },
        setComponent(name) {
            const seen = (occurrences.get(name) ?? 0) + 1;
            occurrences.set(name, seen);
            component = name + '#' + seen;
        },
    };
}

function makeSnapshot() {
    return {
        ok: true,
        cwd: 'D:/proj/temp1',
        kbTree: {
            path: '', name: '全部文档', docs: 0, chunks: 0, bytes: 0,
            children: [{
                path: '来源', name: '来源', docs: 0, chunks: 0, bytes: 0,
                children: [{ path: '来源/方法论', name: '方法论', docs: 0, chunks: 0, bytes: 0, children: [] }],
            }],
        },
    };
}

/** 物化 bundle：只需要 slots 注册表 + remote.commands 注入。 */
function loadBundle(options = {}) {
    const source = readFileSync(BUNDLE, 'utf8');
    const fake = makeFakeReact();
    let entry = null;
    const win = { __ModuleLoader__: { load: (value) => { entry = value; } } };
    const require = (name) => {
        if (name === 'react') return fake.React;
        throw new Error('bundle 请求了未声明的模块：' + name);
    };
    const fetchImpl = options.fetch ?? (() => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(makeSnapshot()) }));
    new Function('window', 'require', 'console', 'fetch', source)(win, require, console, fetchImpl);
    assert.ok(entry !== null, 'bundle 没有调用 window.__ModuleLoader__.load');
    const mod = entry.factory(require);
    const registrations = [];
    const ctx = {
        effect: (fn) => fn(),
        inject: (services, callback) => {
            const wanted = Array.isArray(services) ? services : [services];
            if (wanted.includes('remote.commands') && options.commands) {
                callback({ remote: { commands: options.commands } });
            }
            // configForms / layout / remote.session：本套件用不到，按“这个部署没有”处理。
        },
        slots: {
            inject: (name, fn) => fn(),
            register: (opts, render) => { registrations.push({ opts, render }); return () => {}; },
        },
    };
    mod.apply(ctx);
    return { mod, registrations, fake };
}

/** 展开函数组件，直到只剩宿主元素（同 test/client.mjs 的做法）。 */
function render(registration, props, fake) {
    const expand = (node) => {
        if (node === null || node === undefined || typeof node === 'boolean') return node;
        if (Array.isArray(node)) return node.map(expand);
        if (typeof node !== 'object') return node;
        if (typeof node.type === 'function') {
            const name = node.type.name || 'anon';
            fake.setComponent(name);
            fake.resetCursor();
            return expand(node.type(Object.assign({}, node.props)));
        }
        const next = Object.assign({}, node.props);
        if (next.children !== undefined) next.children = expand(next.children);
        return { type: node.type, props: next };
    };
    let tree = null;
    for (let pass = 0; pass < 10; pass += 1) {
        fake.clearDirty();
        fake.resetInstances();
        tree = expand(registration.render(props));
        if (!fake.isDirty()) break;
    }
    return tree;
}

function walk(node, out = []) {
    if (node === null || node === undefined || node === false) return out;
    if (Array.isArray(node)) { for (const item of node) walk(item, out); return out; }
    if (typeof node !== 'object') return out;
    out.push(node);
    walk(node.props ? node.props.children : undefined, out);
    return out;
}

const tick = () => new Promise((resolve) => setTimeout(resolve, 10));

const useSessionsProp = (cwd) => (selector) => selector({ byId: { 's-1': { cwd } } });

/** 开面板：点一次入口，等状态与目录都回来。 */
function mount(options = {}) {
    const loaded = loadBundle(options);
    const registration = loaded.registrations.find((item) => item.opts.id === 'office-kb-scope');
    assert.ok(registration, '应注册 conversation.input.left / office-kb-scope');
    const props = { sessionId: 's-1', useSessions: useSessionsProp('D:/proj/temp1') };
    return { props, registration, fake: loaded.fake, tree: () => render(registration, props, loaded.fake) };
}

/** 开面板：点一次入口，等状态与目录都回来。 */
async function openPanel(options = {}) {
    const panel = mount(options);
    const toggle = findMarked(panel.tree(), 'data-kb-scope-toggle');
    assert.ok(toggle, '应画出入口按钮');
    toggle.props.onClick();
    await tick();
    return panel;
}

/** 某个样式值的实际底色：transparent 就是面板自己的主题底色。 */
function groundOf(value, theme) {
    const text = String(value ?? '').trim();
    if (text === 'transparent' || text === '') return parseHex(FIXTURE[theme]['--dsw-alias-bg-layer-1']);
    if (text.startsWith('#')) return parseHex(text);
    return themeColor(theme, text);
}

const commandsReturning = (text) => ({ execute: async () => ({ ok: true, value: { result: { kind: 'success', text } } }) });

function findMarked(tree, key, value) {
    return walk(tree).find((node) => node.props && node.props[key] !== undefined
        && (value === undefined || node.props[key] === value));
}

// ── 1. 复现：修之前的写法在深色下是什么样 ────────────────────────────────

await check('复现：--surface 不是宿主 token，深色下白底浅字 = 1.04:1', () => {
    // 宿主 token 表里没有任何 --surface（theme bundle 里 0 处定义），
    // 所以 var(--surface, #fff) 永远取 fallback。
    for (const theme of ['light', 'dark']) {
        assert.equal(Object.prototype.hasOwnProperty.call(FIXTURE[theme], '--surface'), false);
    }
    const inheritedText = parseHex(FIXTURE.dark['--dsw-alias-label-primary']); // #f9fafb
    const buggyPanel = parseHex('#ffffff');                                  // 旧写法的 fallback
    const ratio = contrast(inheritedText, buggyPanel);
    assert.ok(ratio < 1.1, '旧写法的深色对比度应低到看不见，实际 ' + ratio.toFixed(2) + ':1');
    // 也顺手确认：旧说明行 rgba(0,0,0,0.6) 在白底上倒是能读 —— 这正是“只有行文字不见”的现场。
    const oldNote = compositeOver([0, 0, 0, 0.6], buggyPanel);
    assert.ok(contrast(oldNote, buggyPanel) > 4.5);
});

// ── 2. 入口按钮：token、悬停、展开 ───────────────────────────────────────

await check('入口按钮：文字/边框走 token，悬停与展开态各自不同，没有模糊层', async () => {
    const panel = await openPanel({ commands: commandsReturning('知识库范围：已设为「来源/方法论」（含子目录）。') });
    const tree = panel.tree();
    const toggle = findMarked(tree, 'data-kb-scope-toggle');
    const style = toggle.props.style;
    for (const theme of ['light', 'dark']) {
        assertTokenStyle(style.color, theme, '入口文字色');
        assertTokenStyle(style.border, theme, '入口边框');
    }
    assert.equal(style.backdropFilter, undefined, '不要用 backdrop-filter 糊一层');
    assert.equal(style.filter, undefined, '不要用 filter: blur');

    const closed = mount({ commands: commandsReturning('知识库范围：全部知识库。') });
    const closedStyle = findMarked(closed.tree(), 'data-kb-scope-toggle').props.style;
    assert.equal(closedStyle.background, 'transparent', '常态入口不该有底色');
    assert.notEqual(style.background, closedStyle.background, '展开态应与常态不同');

    findMarked(closed.tree(), 'data-kb-scope-toggle').props.onMouseEnter();
    const hoveredStyle = findMarked(closed.tree(), 'data-kb-scope-toggle').props.style;
    assert.notEqual(hoveredStyle.background, closedStyle.background, '悬停态应与常态不同');
    assertTokenStyle(hoveredStyle.background, 'dark', '悬停底');
    findMarked(closed.tree(), 'data-kb-scope-toggle').props.onMouseLeave();
    assert.equal(findMarked(closed.tree(), 'data-kb-scope-toggle').props.style.background, 'transparent', '移开应还原');
});

// ── 3. 面板：显式文字色 + 不透明主题底色 ─────────────────────────────────

await check('面板：底色/文字/边框/阴影都是 token，底色不透明、文字色不靠继承', async () => {
    const panel = await openPanel({ commands: commandsReturning('知识库范围：全部知识库。') });
    const style = findMarked(panel.tree(), 'data-kb-scope-panel').props.style;
    for (const theme of ['light', 'dark']) {
        assertTokenStyle(style.background, theme, '面板底色');
        assertTokenStyle(style.color, theme, '面板文字色');
        assertTokenStyle(style.border, theme, '面板边框');
        assertTokenStyle(style.boxShadow, theme, '面板阴影');
        assert.equal(themeColor(theme, style.background)[3], 1, '面板底色必须不透明（' + theme + '）');
    }
    assert.equal(style.backdropFilter, undefined, '面板不要模糊层');
});

// ── 4. 行：选中 / 未选中 / 悬停 / 禁用 ──────────────────────────────────

await check('行：选中与未选中一眼分得开（底色 + 加粗 + 左侧竖条），aria-pressed 与视觉一致', async () => {
    const panel = await openPanel({ commands: commandsReturning('知识库范围：已设为「来源/方法论」（含子目录）。') });
    const tree = panel.tree();
    const selected = findMarked(tree, 'data-kb-scope-folder', '来源/方法论');
    const plain = findMarked(tree, 'data-kb-scope-folder', '来源');
    assert.equal(selected.props['aria-pressed'], 'true');
    assert.equal(plain.props['aria-pressed'], 'false');
    assert.notEqual(selected.props.style.background, plain.props.style.background, '选中行底色必须不同');
    assert.equal(selected.props.style.fontWeight, '600', '选中行要加粗');
    assertTokenStyle(selected.props.style.boxShadow, 'dark', '选中行左侧竖条');
    assert.equal(selected.props.style.color, plain.props.style.color, '选中不改文字色（对比度已按主色算过）');
    for (const theme of ['light', 'dark']) {
        assertTokenStyle(plain.props.style.color, theme, '行文字色');
        const delta = channelDelta(
            groundOf(selected.props.style.background, theme),
            groundOf(plain.props.style.background, theme),
        );
        assert.ok(delta >= 8, theme + ' 下选中底与未选中底的通道差应可辨认，实际 ' + delta);
    }
});

await check('行：悬停有底色、移开还原；生效期间禁用行用三级文字色且不响应悬停', async () => {
    const commands = commandsReturning('知识库范围：已设为「来源/方法论」（含子目录）。');
    const panel = await openPanel({ commands });
    const before = findMarked(panel.tree(), 'data-kb-scope-folder', '来源');
    assert.equal(before.props.style.background, 'transparent');
    assert.equal(typeof before.props.onMouseEnter, 'function', '可用行应有悬停处理');
    before.props.onMouseEnter();
    const after = findMarked(panel.tree(), 'data-kb-scope-folder', '来源');
    assertTokenStyle(after.props.style.background, 'dark', '悬停底');
    after.props.onMouseLeave();
    assert.equal(findMarked(panel.tree(), 'data-kb-scope-folder', '来源').props.style.background, 'transparent', '移开应还原');

    // 让 pick 悬着：busy 期间行必须禁用（不能拿旧行提交），禁用态要看得出来。
    commands.execute = () => new Promise(() => {});
    findMarked(panel.tree(), 'data-kb-scope-folder', '来源').props.onClick();
    await tick();
    const tree = panel.tree();
    const disabledRow = findMarked(tree, 'data-kb-scope-folder', '来源');
    assert.equal(disabledRow.props.disabled, true, '生效期间行必须禁用');
    assert.equal(tokenOf(disabledRow.props.style.color), '--dsw-alias-label-tertiary', '禁用行用三级文字色');
    assert.equal(disabledRow.props.style.cursor, 'not-allowed');
    assert.equal(disabledRow.props.onMouseEnter, undefined, '禁用行不该响应悬停');
    assert.equal(disabledRow.props.style.background, 'transparent', '禁用行不该有悬停底');

    // 目录读取失败：错误条文字仍是主色（语义色只在竖条上），并给重试入口。
    const broken = await openPanel({
        fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'boom' }) }),
        commands: commandsReturning('知识库范围：全部知识库。'),
    });
    const errorBar = findMarked(broken.tree(), 'data-kb-scope-folders-error');
    assert.ok(errorBar, '应有错误条');
    assert.equal(tokenOf(errorBar.props.style.color), '--dsw-alias-label-primary', '错误条文字用主文字色');
    assertTokenStyle(errorBar.props.style.borderLeft, 'dark', '错误竖条');
    assert.ok(findMarked(broken.tree(), 'data-kb-scope-folders-retry'), '错误条应带重试按钮');
});

await check('提示条：禁止检索的说明用主文字色 + 语义竖条，深色下不再靠 #8a4b00', async () => {
    // 拒绝态：/kb-scope 回「未设置」→ blocked/unset → 面板不收起、显示提示条。
    const panel = await openPanel({ commands: commandsReturning('知识库范围：未设置。') });
    const notice = findMarked(panel.tree(), 'data-kb-scope-notice');
    assert.ok(notice, '应显示禁止检索提示');
    for (const theme of ['light', 'dark']) {
        assertTokenStyle(notice.props.style.color, theme, '提示条文字色');
        assertTokenStyle(notice.props.style.borderLeft, theme, '提示条竖条');
        const text = themeColor(theme, notice.props.style.color);
        const ground = compositeOver(themeColor(theme, notice.props.style.background), parseHex(FIXTURE[theme]['--dsw-alias-bg-layer-1']));
        assert.ok(contrast(text, ground) >= 4.5, theme + ' 提示条文字对比度应 ≥ 4.5，实际 ' + contrast(text, ground).toFixed(2));
    }
});

// ── 5. 两套主题的实际对比度 ─────────────────────────────────────────────

await check('对比度：浅色 / 深色两套都达标（正文说明 ≥4.5，禁用与竖条 ≥3）', () => {
    const lines = [];
    for (const theme of ['light', 'dark']) {
        const panel = parseHex(FIXTURE[theme]['--dsw-alias-bg-layer-1']);
        const hover = compositeOver(parseHex(FIXTURE[theme]['--dsw-alias-interactive-bg-hover']), panel);
        const selected = compositeOver(parseHex(FIXTURE[theme]['--dsw-alias-interactive-bg-active']), panel);
        const text = parseHex(FIXTURE[theme]['--dsw-alias-label-primary']);
        const note = parseHex(FIXTURE[theme]['--dsw-alias-label-secondary']);
        const muted = parseHex(FIXTURE[theme]['--dsw-alias-label-tertiary']);
        const accent = parseHex(FIXTURE[theme]['--dsw-alias-state-business-primary']);
        const pairs = [
            ['正文 on 面板', text, panel, 4.5],
            ['说明 on 面板', note, panel, 4.5],
            ['正文 on 悬停', text, hover, 4.5],
            ['正文 on 选中', text, selected, 4.5],
            ['说明 on 选中', note, selected, 4.5],
            ['禁用 on 面板', muted, panel, 3],
            ['禁用 on 选中', muted, selected, 3],
            ['选中竖条 on 面板', accent, panel, 3],
        ];
        for (const [label, foreground, background, floor] of pairs) {
            const ratio = contrast(foreground, background);
            lines.push(theme + ' ' + label + ' = ' + ratio.toFixed(2) + ':1');
            assert.ok(ratio >= floor, theme + ' ' + label + ' 应 ≥ ' + floor + ':1，实际 ' + ratio.toFixed(2));
        }
        // 三态底色彼此可辨（对比度对“同亮不同色”不敏感，所以另算通道差）。
        assert.ok(channelDelta(selected, panel) >= 8, theme + ' 选中底与面板底应可辨');
        assert.ok(channelDelta(hover, panel) >= 6, theme + ' 悬停底与面板底应可辨');
    }
    console.log('      ' + lines.join('\n      '));
});

// ── 6. 可选：跟宿主安装包对账，确认 fixture 没漂 ─────────────────────────

await check('fixture 与宿主 token 对账（设 DSH_THEME_BUNDLE 时生效）', () => {
    const themePath = process.env.DSH_THEME_BUNDLE || argOf('--theme');
    if (!themePath) {
        console.log('      SKIP 未设 DSH_THEME_BUNDLE，跳过与宿主 bundle 的对账（只跑 fixture 上的对比度）');
        return;
    }
    const source = readFileSync(themePath, 'utf8');
    const statics = new Map();
    for (const matched of source.matchAll(/--dsw-static-([a-z0-9-]+):(#[0-9a-fA-F]{3,8})/g)) {
        statics.set('--dsw-static-' + matched[1], matched[2]);
    }
    /** 从 openIndex（指向 '{'）按花括号配对取出整块。 */
    const blockFrom = (openIndex) => {
        assert.equal(source[openIndex], '{');
        let depth = 0;
        for (let i = openIndex; i < source.length; i += 1) {
            if (source[i] === '{') depth += 1;
            else if (source[i] === '}') {
                depth -= 1;
                if (depth === 0) return source.slice(openIndex + 1, i);
            }
        }
        throw new Error('宿主 bundle 的主题块没有闭合');
    };
    // bundle 里 body{...} / body[data-ds-dark-theme]{...} 会出现多次（分组、局部覆盖），
    // 真正定义整套别名的是**含 label-primary** 的那两块，按选择器取第一块命中。
    const themeBlock = (wantDark) => {
        for (const matched of source.matchAll(/(?:^|[};])body(\[data-ds-dark-theme\])?\{/g)) {
            const open = matched.index + matched[0].length - 1;
            const isDark = matched[1] !== undefined;
            if (isDark !== wantDark) continue;
            const block = blockFrom(open);
            if (block.includes('--dsw-alias-label-primary:')) return block;
        }
        throw new Error('宿主 bundle 里找不到 ' + (wantDark ? '深色' : '浅色') + '主题块');
    };
    const lightBlock = themeBlock(false);
    const darkBlock = themeBlock(true);

    const readToken = (block, name) => {
        const at = block.indexOf(name + ':');
        if (at < 0) return null;
        return block.slice(at + name.length + 1, block.indexOf(';', at));
    };
    const resolve = (value, depth = 0) => {
        assert.ok(depth < 8, 'token 引用层数过深：' + value);
        const matched = /^var\((--dsw-[a-z0-9-]+)\)$/.exec(String(value).trim());
        if (matched === null) return String(value).trim();
        if (statics.has(matched[1])) return statics.get(matched[1]);
        const next = readToken(lightBlock, matched[1]) ?? readToken(darkBlock, matched[1]);
        assert.ok(next !== null, '宿主里找不到 token ' + matched[1]);
        return resolve(next, depth + 1);
    };

    let compared = 0;
    for (const theme of ['light', 'dark']) {
        for (const [name, expected] of Object.entries(FIXTURE[theme])) {
            const raw = theme === 'dark'
                ? (readToken(darkBlock, name) ?? readToken(lightBlock, name))
                : readToken(lightBlock, name);
            assert.ok(raw !== null, '宿主里找不到 ' + name);
            const actual = normalizeHex(resolve(raw));
            assert.equal(actual, normalizeHex(expected), name + '（' + theme + '）的宿主值与 fixture 不一致：' + actual + ' vs ' + expected);
            compared += 1;
        }
    }
    console.log('      宿主 token 对账：' + compared + ' 个值与 fixture 一致');
});

// ── 汇总 ────────────────────────────────────────────────────────────────

let failed = 0;
for (const item of results) {
    if (item.ok) {
        console.log('PASS  ' + item.name);
    } else {
        failed += 1;
        console.log('FAIL  ' + item.name);
        console.log('      ' + (item.error && item.error.message ? item.error.message : String(item.error)));
    }
}
console.log('\nkb-scope-contrast: ' + (results.length - failed) + '/' + results.length + ' 通过');
process.exit(failed === 0 ? 0 : 1);
