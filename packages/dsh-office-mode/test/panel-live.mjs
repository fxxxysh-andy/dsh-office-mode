/**
 * 活体量测：会话页签「记忆」（office-memory, order 40）与「预览」（office-preview, order 50）
 * 的宽度与「拉到底闪烁」，并把宿主自己的四个页签（对话/轨迹/上下文/编队）当参照组。
 *
 * 为什么要活体：这一格在真实运行时由宿主的 `.wSkVaW_*` 布局决定 —— 面板根只是
 * `.wSkVaW_viewArea` 的一个 flex item，宽高、滚动、sticky 全是宿主的 CSS 说了算。
 * 单测里量不到「viewArea 有多宽 / 谁是滚动容器 / 吸底输入框盖住多少」。
 *
 * 手法照抄 test/web-verify.mjs（那边踩过的坑不再重踩）：
 *   1. `dsh web --port 0 --no-open` 起隔离实例，不碰用户的 3080；
 *   2. 必须 `Target.createTarget` 开新标签页（复用会把上一轮导航的缓冲事件重放进来）；
 *   3. 截图用 `Page.captureScreenshot`，不用 `--screenshot` 一次性开关；
 *   4. 每次先删临时 Chrome profile（不清会积累 cookie → DSH 回 431，页面全白）。
 *
 * 三节量测：
 *   A 参照组：对话 / 轨迹 / 上下文 / 编队 —— viewArea 宽度、内容容器 rect 与 max-width、
 *     左右内边距、是否居中、真正的滚动容器、sticky 元素、composer overlay、
 *     `--dsh-chat-content-width` 系列变量的运行时值。
 *   B 目标组：记忆 / 预览 —— 同上，另加「面板根 rect 与 viewArea rect 的左右差」、
 *     首屏文字左边界、1280/1440/1920 三种宽度。
 *   C 拉到底：rAF/定时器把滚动容器从 0 匀速滚到 max（约 1.8s），全程
 *     `PerformanceObserver({type:'layout-shift', buffered:true})` 收 CLS 条目，
 *     每 ~150ms 采一次 `{scrollTop, scrollHeight, clientHeight, 面板头 top, composerSeat rect}`，
 *     在 max-8 / max-2 / max 三个位置各截一张图。同一套对「对话」「上下文」做对照。
 *
 * 判定（任一成立即「复现到问题」）：
 *   · 出现 layout-shift 且 value > 0.001（排除 hadRecentInput）；
 *   · 滚动过程中 scrollHeight（或 clientHeight / viewArea 高度）变化；
 *   · 滚到底时 composerSeat / `[data-conversation-composer-overlay]` 出现或位移；
 *   · 滚到底时内容尾部落到 sticky composerSeat 的上沿之下（=被吸底块盖住）。
 *
 * 用法：node test/panel-live.mjs
 * 产物：.office/runtime/diag/live/report-<ts>.json + .office/runtime/diag/live/*.png
 * 退出码：0 = 复现到问题（打印结论）；3 = 量测成功但未复现；2 = 量测本身失败（打印原因）。
 */
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { mkdirSync, writeFileSync, readFileSync, existsSync, openSync, closeSync, rmSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { dshBinPath } from './host-modules.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '../../..');
const OUT = path.resolve(ROOT, '.office', 'runtime', 'diag', 'live');
const PLUGIN = path.resolve(HERE, '..');
const DSH_BIN = dshBinPath();
const PROFILE = path.join(process.env.USERPROFILE ?? '', '.dsh', 'profiles', 'web');
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
// 与 web-verify.mjs 的 9333 错开：那边可能正在被人跑，撞端口会连到别人的标签页。
const PORT = 9334;
/**
 * 可选：复用已在跑的本机实例（`DSH_OFFICE_PANEL_URL=http://127.0.0.1:3080`）。
 *
 * 为什么要有：本机的 3080 会话占着共享代理，第二个 `dsh web` 经常起不来
 * （实测走不到打印地址那一步）。量面板几何与滚动时，用哪个实例读数一样。
 */
const REUSE_URL = String(process.env.DSH_OFFICE_PANEL_URL ?? '').trim();
const TABS = [
    { text: '对话', slug: 'chat' },
    { text: '轨迹', slug: 'trail' },
    { text: '上下文', slug: 'ctx' },
    { text: '编队', slug: 'team' },
    { text: '记忆', slug: 'mem', root: '[data-memory-panel]' },
    { text: '预览', slug: 'prev', root: '[data-preview-panel]' },
];
const WIDTHS = [1280, 1440, 1920];
const HEIGHT = 900;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const log = (...a) => console.log(...a);
const killTree = (child) => { try { spawn('taskkill', ['/pid', String(child.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* ignore */ } };

function sha256(file) {
    try { return createHash('sha256').update(readFileSync(file)).digest('hex').slice(0, 16); } catch { return null; }
}

if (!existsSync(CHROME)) {
    console.log(`找不到 Chrome：${CHROME}`);
    process.exit(2);
}
if (DSH_BIN === undefined) {
    console.log('找不到宿主入口 lib/bin.js，无法起隔离实例（如实报告，不当成通过）。');
    process.exit(2);
}
const require = createRequire(path.join(PROFILE, 'node_modules', 'x.js'));
let WebSocket;
try {
    WebSocket = require('ws');
} catch (error) {
    console.log(`解析 profile 里的 ws 失败（${PROFILE}）：${error.message}`);
    process.exit(2);
}

// ── 起隔离实例 / CDP ────────────────────────────────────────────────

async function startDsh() {
    // 复用已在跑的本机实例：本机的 3080 会话占着共享代理/锁，第二实例经常起不来
    // （实测 stdout 只到 model-list 就停，120 秒也等不到带 token 的地址）。
    // 量的是「面板在宿主布局里的几何与滚动」，用哪个实例拿到的读数一样。
    if (REUSE_URL !== '') return { child: null, url: REUSE_URL, reused: true };
    mkdirSync(OUT, { recursive: true });
    const out = path.join(OUT, 'web-out.log');
    const err = path.join(OUT, 'web-err.log');
    const fdOut = openSync(out, 'w');
    const fdErr = openSync(err, 'w');
    const child = spawn(process.execPath, [DSH_BIN, 'web', '--port', '0', '--no-open'], {
        cwd: ROOT, stdio: ['ignore', fdOut, fdErr], env: { ...process.env, DSH_NO_OPEN: '1' },
    });
    closeSync(fdOut);
    closeSync(fdErr);
    // 300 秒：本机冷启动实测会停在 opencode-go 模型目录拉取上（网关超时 20–30 秒），
    // 之后才继续加载插件；历史 web-verify 的 100 秒预算不够。
    for (let i = 0; i < 600; i += 1) {
        await sleep(500);
        if (!existsSync(out)) continue;
        const m = readFileSync(out, 'utf8').match(/https?:\/\/127\.0\.0\.1:\d+\/[^\s]*token=[A-Za-z0-9._-]+/);
        if (m) return { child, url: m[0] };
    }
    throw new Error(`没能解析出带 token 的地址：\n${readFileSync(out, 'utf8').slice(-2000)}`);
}

async function cdpConnect() {
    for (let i = 0; i < 60; i += 1) {
        try {
            const json = await (await fetch(`http://127.0.0.1:${PORT}/json/version`)).json();
            if (json.webSocketDebuggerUrl) return json.webSocketDebuggerUrl;
        } catch { /* retry */ }
        await sleep(500);
    }
    throw new Error('Chrome 的 CDP 端点没起来');
}

class Cdp {
    constructor(ws) {
        this.ws = ws; this.id = 0; this.pending = new Map(); this.listeners = [];
        ws.on('message', (raw) => {
            const msg = JSON.parse(raw.toString());
            if (msg.id && this.pending.has(msg.id)) {
                const { resolve, reject } = this.pending.get(msg.id);
                this.pending.delete(msg.id);
                if (msg.error) reject(new Error(`${msg.error.message}`)); else resolve(msg.result);
                return;
            }
            for (const fn of this.listeners) fn(msg);
        });
    }
    send(method, params = {}, sessionId) {
        this.id += 1;
        const id = this.id;
        const payload = { id, method, params };
        if (sessionId) payload.sessionId = sessionId;
        this.ws.send(JSON.stringify(payload));
        return new Promise((resolve, reject) => {
            this.pending.set(id, { resolve, reject });
            setTimeout(() => { if (this.pending.has(id)) { this.pending.delete(id); reject(new Error(`CDP 超时：${method}`)); } }, 30000);
        });
    }
    on(fn) { this.listeners.push(fn); }
}

// ── 页面内探针（String.raw：不写反引号、不写 ${}）──────────────────

const PAGE_HELPERS = String.raw`
window.__plq = function (sel) { return document.querySelector(sel); };
window.__plR = function (el) {
    if (!el) return null;
    var r = el.getBoundingClientRect();
    return { x: Math.round(r.x), y: Math.round(r.y), left: Math.round(r.left), right: Math.round(r.right),
             top: Math.round(r.top), bottom: Math.round(r.bottom), width: Math.round(r.width), height: Math.round(r.height) };
};
window.__plC = function (el) { return el ? String(el.className || '').slice(0, 90) : null; };
window.__plFindScroller = function (el) {
    for (var n = el; n && n !== document.documentElement; n = n.parentElement) {
        var cs = getComputedStyle(n);
        var oy = cs.overflowY;
        if ((oy === 'auto' || oy === 'scroll' || oy === 'overlay') && n.scrollHeight > n.clientHeight + 2) return n;
    }
    return null;
};
window.__plSticky = function (root) {
    var out = [];
    if (!root) return out;
    var all = root.querySelectorAll('*');
    for (var i = 0; i < all.length; i++) {
        var el = all[i];
        var cs = getComputedStyle(el);
        if (cs.position !== 'sticky' && cs.position !== 'fixed') continue;
        var r = el.getBoundingClientRect();
        if (r.width === 0 && r.height === 0) continue;
        out.push({ cls: window.__plC(el), position: cs.position, top: cs.top, bottom: cs.bottom,
                   height: Math.round(r.height), rectTop: Math.round(r.top), rectBottom: Math.round(r.bottom),
                   zIndex: cs.zIndex, opacity: cs.opacity,
                   backgroundImage: cs.backgroundImage === 'none' ? null : String(cs.backgroundImage).slice(0, 120),
                   backgroundColor: cs.backgroundColor,
                   opaque: cs.backgroundColor !== 'rgba(0, 0, 0, 0)' || cs.backgroundImage !== 'none' });
        if (out.length >= 10) break;
    }
    return out;
};
window.__plVars = function (el) {
    var out = {};
    for (var n = el; n && n !== document.documentElement; n = n.parentElement) {
        var cs = getComputedStyle(n);
        var names = ['--dsh-chat-content-width', '--dsh-conversation-column-width', '--dsh-chat-user-width', '--dsh-composer-card-max-width', '--dsh-composer-side-clearance', '--dsh-scrollbar-width'];
        var hit = {};
        var any = false;
        for (var i = 0; i < names.length; i++) {
            var v = (cs.getPropertyValue(names[i]) || '').trim();
            if (v !== '') { hit[names[i]] = v; any = true; }
        }
        if (any) { hit.cls = window.__plC(n); return hit; }
    }
    return null;
};
window.__plFirstText = function (root) {
    if (!root) return null;
    var walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    var node = walker.nextNode();
    while (node) {
        var t = (node.nodeValue || '').trim();
        if (t.length >= 2) {
            var el = node.parentElement;
            var r = el.getBoundingClientRect();
            if (r.width > 0 && r.height > 0) return { text: t.slice(0, 24), x: Math.round(r.left), y: Math.round(r.top), cls: window.__plC(el) };
        }
        node = walker.nextNode();
    }
    return null;
};
window.__plProbe = function (label) {
    var out = { label: label, at: new Date().toISOString() };
    var area = document.querySelector('[class*="viewArea"]');
    var scrollBody = document.querySelector('[class*="scrollBody"]');
    var seat = document.querySelector('[class*="composerSeat"]');
    var panel = document.querySelector('[data-memory-panel]') || document.querySelector('[data-preview-panel]');
    var content = panel || (area ? area.firstElementChild : null);
    out.hasPanel = !!panel;
    out.panelRoot = panel ? window.__plC(panel) : null;
    out.viewArea = window.__plR(area);
    out.scrollBody = scrollBody ? { cls: window.__plC(scrollBody), rect: window.__plR(scrollBody),
        scrollHeight: scrollBody.scrollHeight, clientHeight: scrollBody.clientHeight, scrollTop: Math.round(scrollBody.scrollTop),
        overflowY: getComputedStyle(scrollBody).overflowY } : null;
    out.composerSeat = seat ? { cls: window.__plC(seat), rect: window.__plR(seat), position: getComputedStyle(seat).position,
        backgroundImage: getComputedStyle(seat).backgroundImage === 'none' ? null : String(getComputedStyle(seat).backgroundImage).slice(0, 140),
        zIndex: getComputedStyle(seat).zIndex } : null;
    out.composerOverlay = !!document.querySelector('[data-conversation-composer-overlay]');
    out.vars = window.__plVars(area);
    out.content = content ? { cls: window.__plC(content), rect: window.__plR(content) } : null;
    if (content) {
        var cs = getComputedStyle(content);
        out.contentStyle = {
            display: cs.display, flex: cs.flex, maxWidth: cs.maxWidth, width: cs.width, marginLeft: cs.marginLeft,
            marginRight: cs.marginRight, paddingLeft: cs.paddingLeft, paddingRight: cs.paddingRight,
            overflowX: cs.overflowX, overflowY: cs.overflowY, containerType: cs.containerType,
            alignSelf: cs.alignSelf, boxSizing: cs.boxSizing, height: cs.height,
        };
        var cr = content.getBoundingClientRect();
        var ar = area ? area.getBoundingClientRect() : null;
        out.gaps = ar ? { left: Math.round(cr.left - ar.left), right: Math.round(ar.right - cr.right),
            centered: Math.abs((cr.left - ar.left) - (ar.right - cr.right)) <= 2 } : null;
        out.selfScrolls = content.scrollHeight > content.clientHeight + 2 &&
            /auto|scroll/.test(getComputedStyle(content).overflowY);
        out.contentScroll = { scrollHeight: content.scrollHeight, clientHeight: content.clientHeight,
            scrollTop: Math.round(content.scrollTop), overflowY: getComputedStyle(content).overflowY };
        out.firstText = window.__plFirstText(content);
    }
    var scroller = content ? window.__plFindScroller(content) : null;
    if (scroller === content) scroller = null;
    out.scroller = scroller ? { cls: window.__plC(scroller), tag: scroller.tagName.toLowerCase(),
        overflowY: getComputedStyle(scroller).overflowY, rect: window.__plR(scroller),
        scrollHeight: scroller.scrollHeight, clientHeight: scroller.clientHeight,
        scrollTop: Math.round(scroller.scrollTop), isViewArea: scroller === area } : null;
    if (!out.scroller && scrollBody) {
        out.scroller = { cls: window.__plC(scrollBody), tag: scrollBody.tagName.toLowerCase(),
            overflowY: getComputedStyle(scrollBody).overflowY, rect: window.__plR(scrollBody),
            scrollHeight: scrollBody.scrollHeight, clientHeight: scrollBody.clientHeight,
            scrollTop: Math.round(scrollBody.scrollTop), isViewArea: false, fallback: 'scrollBody' };
    }
    out.sticky = window.__plSticky(area || document.body);
    // 参照组：宿主自己的内容列（比 viewArea 窄、且左右留白不为 0 的块）
    var narrow = [];
    if (area) {
        var ar2 = area.getBoundingClientRect();
        var all = area.querySelectorAll('*');
        for (var i = 0; i < all.length && narrow.length < 6; i++) {
            var el = all[i];
            if (el.children.length === 0) continue;
            var r = el.getBoundingClientRect();
            if (r.width < 320 || r.width > ar2.width - 16) continue;
            var lg = Math.round(r.left - ar2.left);
            var rg = Math.round(ar2.right - r.right);
            if (lg < 4 && rg < 4) continue;
            var ecs = getComputedStyle(el);
            narrow.push({ cls: window.__plC(el), width: Math.round(r.width), leftGap: lg, rightGap: rg,
                maxWidth: ecs.maxWidth, marginLeft: ecs.marginLeft, marginRight: ecs.marginRight,
                paddingLeft: ecs.paddingLeft, paddingRight: ecs.paddingRight,
                centered: Math.abs(lg - rg) <= 2 });
        }
    }
    out.narrowColumns = narrow;
    out.doc = { scrollHeight: document.documentElement.scrollHeight, innerHeight: window.innerHeight, innerWidth: window.innerWidth };
    return out;
};
`;

// ── 主流程 ──────────────────────────────────────────────────────────

const findings = {
    bootOk: null,
    exceptions: [],
    consoleErrors: [],
    session: {},
    widths: {},
    scroll: {},
    verdict: { reproduced: null, reasons: [], mechanism: [] },
    shots: [],
};

async function main() {
    mkdirSync(OUT, { recursive: true });
    const startedAt = new Date().toISOString();
    const pluginClient = path.join(PLUGIN, 'lib', 'client.js');
    findings.meta = {
        startedAt, pluginClient, pluginClientSha: sha256(pluginClient),
        pluginClientMtime: existsSync(pluginClient) ? new Date(statMtime(pluginClient)).toISOString() : null,
        cwd: ROOT, widths: WIDTHS, height: HEIGHT,
    };

    let dshChild = null;
    let chrome = null;
    let ws = null;
    try {
        log('① 启动隔离实例 dsh web --port 0 --no-open');
        const started = await startDsh();
        dshChild = started.child;
        findings.meta.url = started.url.replace(/token=[^&]*/, 'token=***');
        findings.meta.reusedInstance = started.reused === true;
        log(started.reused === true ? `   复用已在跑的实例：${findings.meta.url}` : `   地址：${findings.meta.url}`);

        log('② 启动无头 Chrome + CDP');
        const chromeProfile = path.join(OUT, 'chrome-profile');
        if (existsSync(chromeProfile)) rmSync(chromeProfile, { recursive: true, force: true });
        chrome = spawn(CHROME, [
            '--headless=new', `--remote-debugging-port=${PORT}`,
            '--no-first-run', '--no-default-browser-check', '--disable-gpu',
            `--window-size=${WIDTHS[0]},${HEIGHT}`,
            '--user-data-dir=' + chromeProfile, 'about:blank',
        ], { stdio: 'ignore' });

        const wsUrl = await cdpConnect();
        ws = new WebSocket(wsUrl, { perMessageDeflate: false, maxPayload: 256 * 1024 * 1024 });
        await new Promise((res, rej) => { ws.on('open', res); ws.on('error', rej); });
        const cdp = new Cdp(ws);

        const { targetId } = await cdp.send('Target.createTarget', { url: 'about:blank' });
        const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: true });

        cdp.on((msg) => {
            if (msg.method === 'Runtime.exceptionThrown') {
                findings.exceptions.push(msg.params.exceptionDetails?.exception?.description ?? msg.params.exceptionDetails?.text ?? '?');
            }
            if (msg.method === 'Runtime.consoleAPICalled' && ['error', 'warning'].includes(msg.params.type)) {
                findings.consoleErrors.push(msg.params.type + ': ' + msg.params.args.map((a) => a.value ?? a.description ?? a.type).join(' '));
            }
        });

        await cdp.send('Runtime.enable', {}, sessionId);
        await cdp.send('Page.enable', {}, sessionId);
        await cdp.send('Network.enable', {}, sessionId);
        await cdp.send('Emulation.setDeviceMetricsOverride', { width: WIDTHS[0], height: HEIGHT, deviceScaleFactor: 1, mobile: false }, sessionId);

        log('③ 打开页面（新标签页）');
        await cdp.send('Page.navigate', { url: started.url }, sessionId);
        await sleep(9000);

        const evaluate = async (expression) => {
            const r = await cdp.send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true }, sessionId);
            if (r.exceptionDetails) throw new Error(`页面异常：${r.exceptionDetails.text} ${r.exceptionDetails.exception?.description ?? ''}`);
            return r.result?.value;
        };
        // 别的插件会弹自己的对话框（本轮实测：一块「峰谷时段新增进度条」的设置卡
        // 浮在会话内容中间），它会把量测截图挡掉大半。截图前只点**弹窗容器里面**的
        // 关闭类按钮（×、关闭、稍后、取消），别的一概不碰 —— 免得误触应用里的按钮。
        const dismissOverlays = async () => {
            try {
                const closed = await evaluate(`(function(){
                    var closed = 0;
                    var boxes = Array.prototype.slice.call(document.querySelectorAll('[role="dialog"], [role="alertdialog"], [class*="dialog"], [class*="Dialog"], [class*="modal"], [class*="Modal"]'));
                    boxes.forEach(function (box) {
                        if (box.offsetParent === null) return;
                        Array.prototype.slice.call(box.querySelectorAll('button, [role="button"]')).forEach(function (b) {
                            var text = (b.textContent || '').trim();
                            var label = (b.getAttribute('aria-label') || '') + ' ' + text;
                            if (/^[×✕x]$/i.test(text) || /关闭|稍后|取消|Close|Cancel|Dismiss/i.test(label)) { b.click(); closed += 1; }
                        });
                    });
                    return closed;
                })()`);
                findings.meta.dismissedOverlays = (findings.meta.dismissedOverlays || 0) + (closed > 0 ? closed : 0);
                if (closed > 0) await sleep(150);
                return closed;
            } catch (error) {
                findings.meta.dismissOverlayError = String(error && error.message ? error.message : error);
                return -1;
            }
        };
        const shot = async (name) => {
            await dismissOverlays();
            const r = await cdp.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false }, sessionId);
            const p = path.join(OUT, name + '.png');
            writeFileSync(p, Buffer.from(r.data, 'base64'));
            findings.shots.push(path.relative(OUT, p));
            return path.relative(OUT, p);
        };

        await evaluate(PAGE_HELPERS);
        findings.bootOk = await evaluate('/did not activate|Failed to load plugins/i.test(document.body.innerText) === false');
        await shot('00-boot');

        // ── 开会话（照抄 web-verify 的逐行尝试法：共用同一份会话史，第一行可能是
        //    正在写日志的那条，点开仍是 hero）──────────────────────────
        log('④ 打开一条有内容的会话');
        const READ_CONV_TABS = `Array.from(document.querySelectorAll('[data-conversation-tabs] [role="tab"]')).map(function(e){ return (e.innerText||'').trim(); })`;
        const step = { before: await evaluate(READ_CONV_TABS), attempts: [] };
        if (step.before.length === 0) {
            let candidates = [];
            let waited = 0;
            const readRows = `Array.from(document.querySelectorAll('[role="treeitem"][aria-selected]')).map(function(e){ return { text: (e.innerText||'').trim().slice(0,30), sel: e.getAttribute('aria-selected') }; })`;
            while (waited < 30000) {
                candidates = await evaluate(readRows);
                if (Array.isArray(candidates) && candidates.length > 0) break;
                await sleep(1500); waited += 1500;
            }
            step.candidates = candidates;
            step.waitedMs = waited;
            const unselected = (candidates || []).filter((c) => c.sel === 'false');
            const order = unselected.length > 0 ? unselected : (candidates || []);
            for (let index = 0; index < Math.min(order.length, 6); index += 1) {
                step.attempts.push({ how: `session-row#${index}`, result: await evaluate(`(function(){
                    var rows = Array.from(document.querySelectorAll('[role="treeitem"][aria-selected]'));
                    var unsel = rows.filter(function(e){ return e.getAttribute('aria-selected') === 'false'; });
                    var list = unsel.length > 0 ? unsel : rows;
                    var row = list[${index}] || null;
                    if (!row) return null;
                    row.click();
                    return { sessions: rows.length, text: (row.innerText||'').trim().slice(0,30) };
                })()`) });
                await sleep(7000);
                if ((await evaluate(READ_CONV_TABS)).length > 0) break;
            }
            if ((await evaluate(READ_CONV_TABS)).length === 0) {
                step.attempts.push({ how: 'new-session-button', result: await evaluate(`(function(){
                    var named = Array.from(document.querySelectorAll('button')).filter(function(e){ return (e.getAttribute('aria-label')||'').trim() === '新建会话'; });
                    var hit = named.find(function(e){ return (e.innerText||'').includes('新会话'); }) || named[named.length-1] || null;
                    if (!hit) return null;
                    hit.click();
                    return { candidates: named.length };
                })()`) });
                await sleep(7000);
            }
        }
        step.after = await evaluate(READ_CONV_TABS);
        findings.session = step;
        log(`   候选会话行：${JSON.stringify((step.candidates || []).map((c) => c.text))}`);
        log(`   会话页签：${JSON.stringify(step.after)}`);
        if (step.after.length === 0) throw new Error('没打开任何会话：会话页签行为空，面板测不到（不是面板的错）');

        const clickTab = async (text) => {
            const ok = await evaluate(`(function(){
                var exact = function(e){ return (e.innerText||'').trim() === ${JSON.stringify(text)}; };
                var row = Array.from(document.querySelectorAll('[data-conversation-tabs] [role="tab"]'));
                var hit = row.find(exact);
                if (!hit) return false;
                hit.click();
                return true;
            })()`);
            await sleep(1600);
            return ok;
        };
        const tabLabels = await evaluate(`Array.from(document.querySelectorAll('[data-conversation-tabs] [role="tab"]')).map(function(e){ return (e.innerText||'').trim(); })`);
        findings.session.tabLabels = tabLabels;
        log(`   页签行：${JSON.stringify(tabLabels)}`);

        // ── A/B：三种宽度 × 六个页签 ─────────────────────────────────
        log('⑤ A/B 量测：宽度 × 页签');
        for (const width of WIDTHS) {
            await cdp.send('Emulation.setDeviceMetricsOverride', { width, height: HEIGHT, deviceScaleFactor: 1, mobile: false }, sessionId);
            await sleep(700);
            findings.widths[String(width)] = {};
            for (const tab of TABS) {
                const clicked = await clickTab(tab.text);
                if (!clicked) { findings.widths[String(width)][tab.text] = { missing: true }; continue; }
                if (tab.root) {
                    for (let i = 0; i < 12; i += 1) {
                        if (await evaluate(`!!document.querySelector(${JSON.stringify(tab.root)})`)) break;
                        await sleep(400);
                    }
                    // 预览：选一份文件看右列（工作区根有 .md/.pptx/pdf 样例）
                    if (tab.text === '预览') {
                        const picked = await evaluate(`(function(){
                            var sel = document.querySelector('[data-preview-file]');
                            if (!sel) return null;
                            var setter = Object.getOwnPropertyDescriptor(window.HTMLSelectElement.prototype, 'value').set;
                            var opt = Array.from(sel.options).find(function(o){ return /pptx|pdf|md/i.test(o.textContent || ''); }) || sel.options[0];
                            if (!opt) return null;
                            setter.call(sel, opt.value);
                            sel.dispatchEvent(new Event('change', { bubbles: true }));
                            return { options: sel.options.length, picked: opt.textContent };
                        })()`);
                        findings.widths[String(width)][tab.text + ':picked'] = picked;
                        await sleep(4500);
                    }
                }
                const probe = await evaluate(`window.__plProbe(${JSON.stringify(tab.text)})`);
                findings.widths[String(width)][tab.text] = probe;
                await shot(`w${width}-${tab.slug}`);
                const gapInfo = probe.gaps ? `gap L/R=${probe.gaps.left}/${probe.gaps.right}` : 'gap=?';
                const cw = probe.content ? `content=${probe.content.rect.width}px maxW=${probe.contentStyle?.maxWidth}` : 'content=null';
                log(`   ${width}px ${tab.text}: viewArea=${probe.viewArea?.width}×${probe.viewArea?.height} ${cw} ${gapInfo} scroller=${probe.scroller?.cls ?? 'null'}(${probe.scroller?.scrollHeight ?? '-'}/${probe.scroller?.clientHeight ?? '-'}) overlay=${probe.composerOverlay} sticky=${probe.sticky.length}`);
            }
        }
        await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: HEIGHT, deviceScaleFactor: 1, mobile: false }, sessionId);
        await sleep(700);

        // ── C：拉到底闪烁 ────────────────────────────────────────────
        log('⑥ C 量测：滚到底（layout-shift + 采样 + 三张图）');
        for (const tab of TABS) {
            if (!['对话', '上下文', '记忆', '预览'].includes(tab.text)) continue;
            await clickTab(tab.text);
            if (tab.root) {
                for (let i = 0; i < 12; i += 1) {
                    if (await evaluate(`!!document.querySelector(${JSON.stringify(tab.root)})`)) break;
                    await sleep(400);
                }
            }
            findings.scroll[tab.text] = await measureScroll(cdp, sessionId, evaluate, shot, tab);
            const r = findings.scroll[tab.text];
            log(`   ${tab.text}: scroller=${r.scroller?.cls ?? 'null'} max=${r.max} shifts=${r.shiftCount} (sum=${r.shiftValue}) scrollHeight变化=${JSON.stringify(r.scrollHeightValues)} seatTop变化=${JSON.stringify(r.seatTopValues)} 尾部被盖=${r.tailUnderSeat}px`);
        }

        // ── 判定 ────────────────────────────────────────────────────
        findings.verdict = judge(findings);
        findings.error = null;
        return findings.verdict;
    } finally {
        try { if (ws) ws.close(); } catch { /* ignore */ }
        if (chrome) try { killTree(chrome); } catch { /* ignore */ }
        if (dshChild) try { killTree(dshChild); } catch { /* ignore */ }
    }
}

function statMtime(file) {
    try { return statSync(file).mtimeMs; } catch { return 0; }
}

/**
 * 一套脚本化滚动：先把滚动容器 scrollTop 设 0，装 PerformanceObserver，
 * 再用定时器 16ms 一帧把 scrollTop 从 0 推到 max（约 1.8s），每 ~150ms 采一次
 * 关键 rect；然后在 max-8 / max-2 / max 各停 150ms 截一张图。
 */
async function measureScroll(cdp, sessionId, evaluate, shot, tab) {
    const result = { tab: tab.text, scroller: null, max: null, samples: [], shifts: [], shots: [], notes: [] };
    const setup = await evaluate(`(function(){
        var content = document.querySelector('[data-memory-panel]') || document.querySelector('[data-preview-panel]');
        var area = document.querySelector('[class*="viewArea"]');
        var body = document.querySelector('[class*="scrollBody"]');
        var sc = content ? window.__plFindScroller(content) : null;
        if (!sc && body && body.scrollHeight > body.clientHeight + 2) sc = body;
        window.__plScroll = { scroller: sc, entries: [], samples: [], max: 0 };
        if (!sc) return { error: '找不到滚动容器' };
        try {
            window.__plScroll.po = new PerformanceObserver(function (list) {
                var items = list.getEntries();
                for (var i = 0; i < items.length; i++) {
                    var e = items[i];
                    var sources = (e.sources || []).map(function (s) {
                        return { node: s.node ? (s.node.nodeName + '.' + String(s.node.className || '').slice(0, 60)) : null };
                    });
                    window.__plScroll.entries.push({ value: Math.round(e.value * 1000000) / 1000000, hadRecentInput: e.hadRecentInput, startTime: Math.round(e.startTime), sources: sources });
                }
            });
            window.__plScroll.po.observe({ type: 'layout-shift', buffered: true });
        } catch (err) { window.__plScroll.observerError = String(err && err.message ? err.message : err); }
        sc.style.scrollBehavior = 'auto';
        sc.scrollTop = 0;
        return { cls: window.__plC(sc), tag: sc.tagName.toLowerCase(), overflowY: getComputedStyle(sc).overflowY,
                 scrollHeight: sc.scrollHeight, clientHeight: sc.clientHeight, max: sc.scrollHeight - sc.clientHeight,
                 rect: window.__plR(sc), observerError: window.__plScroll.observerError || null };
    })()`);
    if (setup && setup.error) { result.notes.push(setup.error); return result; }
    result.scroller = setup;
    const sampleOnce = `(function(){
        var S = window.__plScroll, el = S.scroller;
        var area = document.querySelector('[class*="viewArea"]');
        var seat = document.querySelector('[class*="composerSeat"]');
        var head = document.querySelector('[data-memory-head="sticky"]') || document.querySelector('[data-preview-head="sticky"]');
        var content = document.querySelector('[data-memory-panel]') || document.querySelector('[data-preview-panel]');
        var tail = content && content.lastElementChild ? content.lastElementChild.getBoundingClientRect() : null;
        var rec = {
            t: Math.round(performance.now()),
            scrollTop: Math.round(el.scrollTop),
            scrollHeight: el.scrollHeight,
            clientHeight: el.clientHeight,
            areaHeight: area ? Math.round(area.getBoundingClientRect().height) : null,
            areaTop: area ? Math.round(area.getBoundingClientRect().top) : null,
            seatTop: seat ? Math.round(seat.getBoundingClientRect().top) : null,
            seatHeight: seat ? Math.round(seat.getBoundingClientRect().height) : null,
            headTop: head ? Math.round(head.getBoundingClientRect().top) : null,
            headHeight: head ? Math.round(head.getBoundingClientRect().height) : null,
            overlay: !!document.querySelector('[data-conversation-composer-overlay]'),
            tailBottom: tail ? Math.round(tail.bottom) : null,
            firstTextX: (function(){ var f = content ? window.__plFirstText(content) : null; return f ? f.x : null; })()
        };
        S.samples.push(rec);
        return rec;
    })()`;

    // 装 CLS 之后先把页面稳定一下（避免把「点页签」本身的位移算进来）
    await evaluate(`new Promise(function (r) { requestAnimationFrame(function () { requestAnimationFrame(r); }); })`);
    result.atTop = await evaluate(sampleOnce);

    const anim = await evaluate(`new Promise(function (resolve) {
        var S = window.__plScroll, el = S.scroller;
        var max = Math.max(0, el.scrollHeight - el.clientHeight);
        S.max = max;
        var t0 = performance.now();
        var dur = 1800;
        var lastSample = -1000;
        var tick = 0;
        function sample() {
            var area = document.querySelector('[class*="viewArea"]');
            var seat = document.querySelector('[class*="composerSeat"]');
            var head = document.querySelector('[data-memory-head="sticky"]') || document.querySelector('[data-preview-head="sticky"]');
            S.samples.push({
                t: Math.round(performance.now() - t0),
                scrollTop: Math.round(el.scrollTop),
                scrollHeight: el.scrollHeight,
                clientHeight: el.clientHeight,
                areaHeight: area ? Math.round(area.getBoundingClientRect().height) : null,
                seatTop: seat ? Math.round(seat.getBoundingClientRect().top) : null,
                seatHeight: seat ? Math.round(seat.getBoundingClientRect().height) : null,
                headTop: head ? Math.round(head.getBoundingClientRect().top) : null,
                overlay: !!document.querySelector('[data-conversation-composer-overlay]'),
                tick: tick
            });
        }
        var timer = setInterval(function () {
            tick += 1;
            var now = performance.now();
            var p = Math.min(1, (now - t0) / dur);
            el.scrollTop = Math.round(max * p);
            if (now - lastSample >= 150) { lastSample = now; sample(); }
            if (p >= 1) { clearInterval(timer); sample(); resolve({ max: max, ticks: tick, elapsed: Math.round(now - t0) }); }
        }, 16);
    })`);
    result.max = anim?.max ?? null;
    result.anim = anim;

    // 三个位置各截一张：max-8 / max-2 / max
    for (const [name, offset] of [['max-8', 8], ['max-2', 2], ['max', 0]]) {
        const pos = await evaluate(`(function(){
            var S = window.__plScroll, el = S.scroller;
            var target = Math.max(0, S.max - ${offset});
            el.scrollTop = target;
            return target;
        })()`);
        await sleep(180);
        const sample = await evaluate(sampleOnce);
        const file = await shot(`scroll-${tab.slug}-${name}`);
        result.shots.push({ at: name, target: pos, file, sample });
    }
    // 再正反抖两下，看有没有重排（用户「拉到底」时的实际动作）
    await evaluate(`new Promise(function (resolve) {
        var S = window.__plScroll, el = S.scroller;
        var i = 0;
        var targets = [S.max, S.max - 24, S.max, S.max - 12, S.max];
        var timer = setInterval(function () {
            el.scrollTop = targets[i];
            i += 1;
            if (i >= targets.length) { clearInterval(timer); resolve(true); }
        }, 120);
    })`);
    await sleep(200);
    result.afterJiggle = await evaluate(sampleOnce);

    result.samples = await evaluate(`window.__plScroll.samples`);
    result.shifts = await evaluate(`window.__plScroll.entries`);
    result.observerError = await evaluate(`window.__plScroll.observerError || null`);

    const scrollHeights = [...new Set(result.samples.map((s) => s.scrollHeight))];
    const clientHeights = [...new Set(result.samples.map((s) => s.clientHeight))];
    const areaHeights = [...new Set(result.samples.map((s) => s.areaHeight))];
    const seatTops = [...new Set(result.samples.map((s) => s.seatTop))];
    result.scrollHeightValues = scrollHeights;
    result.clientHeightValues = clientHeights;
    result.areaHeightValues = areaHeights;
    result.seatTopValues = seatTops;
    result.overlayValues = [...new Set(result.samples.map((s) => s.overlay))];
    result.overlayToggle = result.overlayValues.length > 1;
    result.shiftCount = result.shifts.length;
    result.shiftValue = Math.round(result.shifts.reduce((a, s) => a + (s.hadRecentInput ? 0 : s.value), 0) * 10000) / 10000;
    result.shiftValueAll = Math.round(result.shifts.reduce((a, s) => a + s.value, 0) * 10000) / 10000;
    const bottomSample = result.samples[result.samples.length - 1];
    result.tailUnderSeat = bottomSample && bottomSample.seatTop !== null && result.shots.length > 0
        ? bottomSample.tailBottom - bottomSample.seatTop : null;
    return result;
}

/** 判定：任一条成立即「复现到问题」。
 *
 * 注意这是**量测工具**不是守门断言：即便复现了也不是「代码坏了」，而是「读数在这里」。
 * 所以 main 正常跑完一律 exit 0，结论写进报告；只有真跑不下去（起不了实例 / CDP 断了）才 exit 2。
 *
 * 口径（第七十一轮第二次实跑后收紧）：宿主外壳自己就会在滚动时抖动 ——
 * 「上下文」「编队」两个参照页签的 layout-shift 合计与我们的面板几乎同量级
 * （实测 0.6004 / 0.7089，条目逐条相同，全落在 footArea / sessionRow 上）。
 * 所以这里不再用「合计 > 0.001」这种绝对阈值（那会把外壳算成面板的错），
 * 改成**和参照页签比**：只有比参照多出 0.05 以上，或者高度真的变了 4px 以上，
 * 或者内容尾落进吸底块，才算复现。量不到滚动容器的页签只记线索，不进判定。
 */
function judge(data) {
    const reasons = [];
    const mechanism = [];
    const entries = Object.entries(data.scroll || {});
    const REF_TABS = ['上下文', '编队'];
    const measured = (r) => r && Array.isArray(r.samples) && r.samples.length > 0;
    const refTab = REF_TABS.find((name) => measured(data.scroll && data.scroll[name])) || null;
    const refShift = refTab ? data.scroll[refTab].shiftValue : 0;
    const reference = refTab ? { tab: refTab, shiftValue: refShift } : null;
    const round4 = (n) => Math.round(n * 10000) / 10000;
    const spread = (values) => (Array.isArray(values) && values.length > 0 ? Math.max(...values) - Math.min(...values) : 0);
    for (const [tab, r] of entries) {
        if (!measured(r)) {
            mechanism.push(`${tab}: 没量到滚动容器（${(r.notes || []).join(';') || '没有可滚的 max'}）——不参与判定`);
            continue;
        }
        if (REF_TABS.includes(tab)) continue;
        const delta = typeof r.shiftValue === 'number' ? round4(r.shiftValue - refShift) : null;
        if (delta !== null && delta > 0.05) {
            reasons.push(`${tab}: layout-shift 合计 ${r.shiftValue}，比参照「${refTab}」的 ${refShift} 多 ${delta}（> 0.05 才算面板自己的位移）`);
        } else if (delta !== null) {
            mechanism.push(`${tab}: layout-shift ${r.shiftValue} ≈ 参照「${refTab}」${refShift}（差 ${delta}）→ 判为宿主外壳位移`);
        }
        if (spread(r.scrollHeightValues) > 4) reasons.push(`${tab}: 滚动中 scrollHeight 变化 > 4px ${JSON.stringify(r.scrollHeightValues)}`);
        if (spread(r.clientHeightValues) > 4) reasons.push(`${tab}: 滚动中 clientHeight 变化 > 4px ${JSON.stringify(r.clientHeightValues)}`);
        if (spread(r.areaHeightValues) > 4) reasons.push(`${tab}: 滚动中 viewArea 高度变化 > 4px ${JSON.stringify(r.areaHeightValues)}`);
        if (spread(r.seatTopValues) > 4) reasons.push(`${tab}: 滚动中 composerSeat top 变化 > 4px ${JSON.stringify(r.seatTopValues)}`);
        if (r.overlayToggle) reasons.push(`${tab}: 滚动中 composer overlay 出现/消失`);
        if (typeof r.tailUnderSeat === 'number' && r.tailUnderSeat > 2) reasons.push(`${tab}: 滚到底时内容尾落进吸底块 ${r.tailUnderSeat}px`);
        if (r.scroller && r.max === 0) mechanism.push(`${tab}: 滚动容器 max=0（内容没溢出，这一格不会闪；要更长内容才看得到）`);
        if (typeof r.tailUnderSeat === 'number' && r.tailUnderSeat < 0) mechanism.push(`${tab}: 滚到底时内容尾在吸底块上方 ${-r.tailUnderSeat}px（没被盖）`);
    }
    return { reproduced: reasons.length > 0, reference, reasons, mechanism };
}

// ── 收尾 ────────────────────────────────────────────────────────────

const startedMs = Date.now();
main().then((verdict) => {
    findings.meta.finishedAt = new Date().toISOString();
    findings.meta.elapsedMs = Date.now() - startedMs;
    if (findings.meta.pluginClient) findings.meta.pluginClientShaAfter = sha256(findings.meta.pluginClient);
    const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const file = path.join(OUT, `report-${stamp}.json`);
    writeFileSync(file, JSON.stringify(findings, null, 2));
    log('\n── 结论 ──');
    log(`报告：${file}`);
    log(`截图 ${findings.shots.length} 张：${OUT}`);
    log(`复现：${verdict.reproduced ? '是' : '否'}`);
    for (const reason of verdict.reasons) log(`  · ${reason}`);
    log(`插件 client.js sha：${findings.meta.pluginClientSha} → ${findings.meta.pluginClientShaAfter}${findings.meta.pluginClientSha === findings.meta.pluginClientShaAfter ? '（量测期间未变）' : '（量测期间被改过！读数可能跨版本）'}`);
    if (findings.exceptions.length > 0) log(`页面异常 ${findings.exceptions.length} 条：${findings.exceptions[0].slice(0, 200)}`);
    if (findings.consoleErrors.length > 0) log(`控制台 ${findings.consoleErrors.length} 条：${findings.consoleErrors.slice(0, 3).join(' | ')}`);
    // 量测工具：跑完就是成功（读数与结论都在报告里），复现与否用 exit 0 之外的信号表达
    // 会让 `run-suites` 把「量到了」判成红 —— 之前就是这样，一个可选套件天天报红。
    process.exit(0);
}).catch((error) => {
    findings.meta.finishedAt = new Date().toISOString();
    findings.error = String(error && error.stack ? error.stack : error);
    try {
        const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        writeFileSync(path.join(OUT, `report-failed-${stamp}.json`), JSON.stringify(findings, null, 2));
    } catch { /* ignore */ }
    log(`\n量测失败：${error && error.message ? error.message : error}`);
    if (findings.session.after) log(`当时会话页签：${JSON.stringify(findings.session.after)}`);
    if (findings.exceptions.length > 0) log(`页面异常：${findings.exceptions.slice(0, 3).join(' | ').slice(0, 800)}`);
    process.exit(2);
});
