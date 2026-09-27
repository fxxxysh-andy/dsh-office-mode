/**
 * 活体验证：起一个隔离的 DSH Web 实例 + 无头 Chrome，用 CDP 打开页面，
 * 取「记忆面板到底长什么样」的运行时事实（而不是只看单测）。
 *
 * 手法来自开发期的专题笔记，三个要点都照做：
 *   1. `dsh web --port 0 --no-open`：不碰用户的 3080；
 *   2. 必须用 CDP 的 Target.createTarget 开**新标签页**，复用标签页会把上一次
 *      导航的缓冲事件重放进来；
 *   3. 不用 `--screenshot` 一次性开关（对常驻 WebSocket 的 SPA 会挂住），
 *      改用 Page.captureScreenshot。
 *
 * 本轮额外加一招：用 CDP 的 Fetch 域**拦截** `/office-memory/snapshot` 并回一份
 * 合成快照。隔离实例里没有任何工具跑过，端点本来只会回「还没有见过任何工作目录」，
 * 面板就是空态 —— 那样验不到关系图、容量条、时间线。拦截之后，跑的是真实 React
 * 渲染器 + 真实 DOM，喂的是完整数据，等于把「数据到位时长什么样」也钉住了。
 *
 * 用法：node test/web-verify.mjs
 *
 * 产物（截图、web-verify.json、日志、临时 Chrome profile）都写进仓库根的
 * `.office/diag/` —— 与其它中间产物同一棵隐藏树，不进版本库。
 */
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { mkdirSync, writeFileSync, readFileSync, existsSync, openSync, closeSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { dshBinPath } from './host-modules.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '../../..');
const OUT = path.resolve(ROOT, '.office', 'diag');
// 宿主 Web 入口：从宿主的解析位置找（见 host-modules.mjs），不写死某台机器的安装路径。
const DSH_BIN = dshBinPath();
if (DSH_BIN === undefined) {
    console.log('找不到宿主入口 lib/bin.js（DSH_BIN 与宿主安装位置都没给），跳过活体验证（如实报告，不当成通过）。');
    process.exit(2);
}
const PROFILE = path.join(process.env.USERPROFILE, '.dsh', 'profiles', 'web');
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const PORT = 9333;

const require = createRequire(path.join(PROFILE, 'node_modules', 'x.js'));
const WebSocket = require('ws');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const log = (...a) => console.log(...a);
const killTree = (child) => { try { spawn('taskkill', ['/pid', String(child.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* ignore */ } };

/** 合成快照：字段与 src/view.js 的冻结契约一一对应（含 bytes / hotBytes / projectBytes）。 */
function makeSnapshot() {
    const at = '2026-09-24T09:12:00.000Z';
    // 台账补到 25 条（> LIST_PAGE 的 20）：面板的「再显示 N 条」只有真的超过一页
    // 才会出现，少于 20 条时那条分页栏按设计根本不画。
    const ledger = [
        { id: 'L-1', at, path: '季度汇报.pptx', format: 'ppt', theme: 'business', purpose: '季度汇报', outline: ['封面', '要点', '表格'], origin: '工作区' },
        { id: 'L-2', at: '2026-09-24T08:00:00.000Z', path: '季度数据.xlsx', format: 'excel', theme: 'plain', purpose: '数据明细', outline: ['表格'], origin: '工作区' },
        { id: 'L-3', at: '2026-09-23T08:00:00.000Z', path: '季度回顾.docx', format: 'word', theme: 'plain', purpose: '文字稿', outline: ['H1 概述'], origin: '工作区' },
        { id: 'L-4', at: '2026-09-23T07:00:00.000Z', path: '论文.pdf', format: 'pdf', theme: '', purpose: '论文编译', outline: [], origin: '工作区' },
    ];
    for (let index = 5; index <= 25; index += 1) {
        const day = String(1 + (index % 22)).padStart(2, '0');
        ledger.push({
            id: 'L-' + index, at: '2026-08-' + day + 'T06:00:00.000Z',
            path: '历史产物-' + index + '.docx', format: 'word', theme: 'plain',
            purpose: '第 ' + index + ' 份历史交付物', outline: ['H1'], origin: '工作区',
        });
    }
    return {
        ok: true,
        generatedAt: at,
        cwd: ROOT,
        label: 'temp1',
        query: '',
        workspaces: [{ cwd: ROOT, label: 'temp1', seenAt: at }],
        config: {
            scope: 'workspace', userScope: 'memory',
            layers: { hot: true, ledger: true, archive: true },
            links: true, autoCapture: true,
            quality: { policy: 'strict-v1', lowScoreThreshold: 0.2, highScoreThreshold: 0.6, candidateMultiplier: 4, maxMediumResults: 2, maxUnknownResults: 1 },
            quota: { recallPerTurn: 1, recallRefinePerTurn: 1, relatedPerTurn: 1 },
            limits: { userLimitBytes: 4096, projectLimitBytes: 10240, ledgerLimit: 500, archiveKeep: 60 },
        },
        stores: [{ id: 'workspace', dir: '.office/memory', hot: 4, ledger: 25, archive: 30, archiveFiles: 1, links: 3, bytes: 24576, files: 9 }],
        // 两个口径（见 src/view.js 的注释）：hot/ledger/archive/... 是「列表里现在有几条」，
        // ledgerTotal/archiveFiles/archiveItems 是容量条用的未过滤总数。
        // 这里**故意让归档的条目数（30）与摘要文件数（1）不同** —— 面板曾经拿 30 配上限 60，
        // 显示成「30 / 60 个」像快满了，而实际只用了 1 个文件。
        counts: {
            hot: 4, ledger: 25, archive: 30, links: 3, entities: 3,
            ledgerTotal: 25, archiveFiles: 1, archiveItems: 30,
            hotBytes: 812, projectBytes: 1830,
        },
        hot: [
            // m-1 故意写成长文本：面板对长条目给「展开」按钮，展开要去掉正文限高。
            { id: 'm-1', target: 'user', importance: 'critical', origin: '工作区', updatedAt: at, entities: [], tags: ['写作'], content: '偏好：书面化表达，避免说教与第二人称。' + '这条故意写得很长，用来验证「展开」按钮会去掉正文的限高。'.repeat(6) },
            { id: 'm-2', target: 'project', importance: 'critical', origin: '工作区', updatedAt: at, entities: ['系统A'], tags: ['ooxml'], content: '约定：改 OOXML 必须读回原包、只改命中片段、原样写回。' },
            { id: 'm-3', target: 'project', importance: 'normal', origin: '全局', updatedAt: at, entities: ['系统A', '系统B'], tags: [], content: '环境：本机 TeX Live 2022，latexmk -xelatex 可直接编译。' },
            { id: 'm-4', target: 'user', importance: 'normal', origin: '全局', updatedAt: at, entities: [], tags: ['界面'], content: '偏好：面板里的写入动作只生成指令，不直接落盘。' },
        ],
        ledger,
        archive: [
            { id: 'A-1', month: '2026-09', kind: 'insight', at, origin: '工作区', importance: 'normal', category: '环境', entities: ['系统B'], tags: ['tex'], text: '长期记忆一条：TeX 引擎探测顺序 fitz > poppler > mupdf。' },
            { id: 'A-2', month: '2026-08', kind: 'insight', at: '2026-08-20T08:00:00.000Z', origin: '工作区', importance: 'low', category: '工具', entities: [], tags: [], text: '旧的 mnemon 记录：缓存 TTL 默认 12 小时。' },
        ],
        links: [
            { id: 'K-1', sourceId: 'm-1', targetId: 'L-1', kind: 'related', note: '写作偏好影响汇报稿', at, origin: '工作区' },
            { id: 'K-2', sourceId: 'm-2', targetId: 'A-1', kind: 'refines', note: '同一条 OOXML 约定的细化', at, origin: '工作区' },
            { id: 'K-3', sourceId: 'L-2', targetId: 'L-1', kind: 'derives', note: '数据来自明细表', at, origin: '工作区' },
        ],
        entities: [
            { name: '系统A', count: 2, refs: [{ id: 'm-2', layer: 'hot', origin: '工作区', text: '约定：改 OOXML 必须读回原包…' }, { id: 'm-3', layer: 'hot', origin: '全局', text: '环境：本机 TeX Live 2022…' }] },
            { name: '系统B', count: 2, refs: [{ id: 'm-3', layer: 'hot', origin: '全局', text: '环境：本机 TeX Live 2022…' }, { id: 'A-1', layer: 'archive', origin: '工作区', text: 'TeX 引擎探测顺序…' }] },
            { name: '季度汇报', count: 1, refs: [{ id: 'L-1', layer: 'ledger', origin: '工作区', text: '季度汇报.pptx' }] },
        ],
    };
}

async function startDsh() {
    const out = path.join(OUT, 'web-out.log');
    const err = path.join(OUT, 'web-err.log');
    const fdOut = openSync(out, 'w');
    const fdErr = openSync(err, 'w');
    const child = spawn(process.execPath, [DSH_BIN, 'web', '--port', '0', '--no-open'], {
        cwd: ROOT, stdio: ['ignore', fdOut, fdErr], env: { ...process.env, DSH_NO_OPEN: '1' },
    });
    closeSync(fdOut);
    closeSync(fdErr);
    for (let i = 0; i < 200; i += 1) {
        await sleep(500);
        if (!existsSync(out)) continue;
        const m = readFileSync(out, 'utf8').match(/https?:\/\/127\.0\.0\.1:\d+\/[^\s]*token=[A-Za-z0-9._-]+/);
        if (m) return { child, url: m[0] };
    }
    throw new Error(`没能解析出带 token 的地址：\n${readFileSync(out, 'utf8').slice(-2000)}`);
}

async function cdpConnect() {
    for (let i = 0; i < 40; i += 1) {
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

const findings = { bootOk: null, exceptions: [], consoleErrors: [], failedRequests: [], intercepted: 0, dom: {}, screenshots: [] };

async function main() {
    mkdirSync(OUT, { recursive: true });
    const snapshot = makeSnapshot();
    const snapshotBody = Buffer.from(JSON.stringify(snapshot), 'utf8').toString('base64');
    log('① 启动隔离实例 dsh web --port 0 --no-open');
    let dshChild = null;
    let chrome = null;
    try {
        const started = await startDsh();
        dshChild = started.child;
        const url = started.url;
        log(`   地址：${url.replace(/token=[^&]*/, 'token=***')}`);

        log('② 启动无头 Chrome + CDP');
        chrome = spawn(CHROME, [
            '--headless=new', `--remote-debugging-port=${PORT}`,
            '--no-first-run', '--no-default-browser-check', '--disable-gpu', '--window-size=1280,900',
            '--user-data-dir=' + path.join(OUT, 'chrome-profile'), 'about:blank',
        ], { stdio: 'ignore' });

        const wsUrl = await cdpConnect();
        const ws = new WebSocket(wsUrl, { perMessageDeflate: false, maxPayload: 256 * 1024 * 1024 });
        await new Promise((res, rej) => { ws.on('open', res); ws.on('error', rej); });
        const cdp = new Cdp(ws);

        // 新标签页：复用标签页会把上一轮导航的缓冲事件重放进来
        const { targetId } = await cdp.send('Target.createTarget', { url: 'about:blank' });
        const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: true });

        cdp.on((msg) => {
            if (msg.method === 'Runtime.exceptionThrown') {
                findings.exceptions.push(msg.params.exceptionDetails?.exception?.description ?? msg.params.exceptionDetails?.text ?? '?');
            }
            if (msg.method === 'Runtime.consoleAPICalled' && ['error', 'warning'].includes(msg.params.type)) {
                findings.consoleErrors.push(msg.params.type + ': ' + msg.params.args.map((a) => a.value ?? a.description ?? a.type).join(' '));
            }
            if (msg.method === 'Network.loadingFailed') {
                findings.failedRequests.push(`${msg.params.type} ${msg.params.errorText}`);
            }
            if (msg.method === 'Network.requestWillBeSent' && /office-memory\/snapshot/.test(msg.params.request.url)) {
                findings.snapshotRequests = (findings.snapshotRequests ?? 0) + 1;
                findings.snapshotUrl = msg.params.request.url;
            }
            // 拦截记忆端点：喂一份完整合成快照，让面板渲染出带数据的形态
            if (msg.method === 'Fetch.requestPaused') {
                const rid = msg.params.requestId;
                if (/\/office-memory\/snapshot/.test(msg.params.request.url)) {
                    findings.intercepted += 1;
                    cdp.send('Fetch.fulfillRequest', {
                        requestId: rid, responseCode: 200,
                        responseHeaders: [{ name: 'content-type', value: 'application/json; charset=utf-8' }],
                        body: snapshotBody,
                    }, msg.sessionId ?? sessionId).catch(() => { /* ignore */ });
                } else {
                    cdp.send('Fetch.continueRequest', { requestId: rid }, msg.sessionId ?? sessionId).catch(() => { /* ignore */ });
                }
            }
        });

        await cdp.send('Runtime.enable', {}, sessionId);
        await cdp.send('Log.enable', {}, sessionId);
        await cdp.send('Page.enable', {}, sessionId);
        await cdp.send('Network.enable', {}, sessionId);
        await cdp.send('Fetch.enable', { patterns: [{ urlPattern: '*office-memory/snapshot*' }] }, sessionId);

        log('③ 打开页面（新标签页）');
        await cdp.send('Page.navigate', { url }, sessionId);
        await sleep(9000);

        const evaluate = async (expression) => {
            const r = await cdp.send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true }, sessionId);
            return r.result?.value;
        };
        const shot = async (name) => {
            const r = await cdp.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true }, sessionId);
            const p = path.join(OUT, name);
            writeFileSync(p, Buffer.from(r.data, 'base64'));
            findings.screenshots.push(p);
            log(`   截图：${p}`);
        };

        findings.dom.bodyText = await evaluate('document.body.innerText.slice(0, 3000)');
        findings.dom.hasFailedBanner = await evaluate('/did not activate|Failed to load plugins/i.test(document.body.innerText)');
        findings.bootOk = findings.dom.hasFailedBanner === false;

        findings.dom.clickables = await evaluate(`Array.from(document.querySelectorAll('button,[role="button"],a')).map(e => ({
            aria: e.getAttribute('aria-label'), title: e.getAttribute('title'), text: (e.innerText||'').trim().slice(0,24)
        })).slice(0, 60)`);

        const clicked = await evaluate(`(() => {
            const all = Array.from(document.querySelectorAll('button,[role="button"],a,li,div'));
            // 侧栏图标 rail：优先 aria-label 恰好是「记忆」的那个。
            // 不能只按 /记忆/ 模糊匹配 —— 会话标题里也可能带「记忆」两个字。
            const exact = all.find(e => (e.getAttribute('aria-label') || '').trim() === '记忆');
            const byText = all.filter(e => (e.innerText || '').trim() === '记忆')
                .sort((a, b) => (a.innerText || '').length - (b.innerText || '').length)[0];
            const hit = exact ?? byText;
            if (!hit) return null;
            hit.click();
            return { tag: hit.tagName.toLowerCase(), aria: hit.getAttribute('aria-label'), matched: exact ? 'aria-label' : 'text' };
        })()`);
        findings.dom.clickedMemory = clicked;
        await sleep(5000);

        const snapshotDom = async () => {
            findings.dom.tabs = await evaluate(`Array.from(document.querySelectorAll('[role="tab"]')).map(e => (e.innerText||'').trim())`);
            findings.dom.graphPresent = await evaluate('!!document.querySelector(\'svg[data-layout]\')');
            findings.dom.graphMeta = await evaluate(`(() => { const s = document.querySelector('svg[data-layout]'); return s ? { nodes: s.getAttribute('data-nodes'), edges: s.getAttribute('data-edges'), layout: s.getAttribute('data-layout'), density: s.getAttribute('data-density'), role: s.getAttribute('role'), label: s.getAttribute('aria-label') } : null; })()`);
            findings.dom.svgShapes = await evaluate(`(() => { const s = document.querySelector('svg[data-layout]'); if (!s) return null; const q = (t) => s.querySelectorAll(t).length; return { circle: q('circle'), rect: q('rect'), polygon: q('polygon'), path: q('path'), text: q('text'), line: q('line') }; })()`);
            findings.dom.panelText = await evaluate(`(() => { const el = Array.from(document.querySelectorAll('*')).find(e => /记忆 ·/.test(e.innerText||'') && (e.innerText||'').length < 4000); return el ? el.innerText.slice(0, 2500) : null; })()`);
            // 面板样式表落没落地：伪类 / 关键帧 / 媒体查询都在那个 <style> 节点里。
            findings.dom.styleNode = await evaluate(`(() => {
                const styles = Array.from(document.querySelectorAll('style')).filter(s => (s.textContent||'').includes('.om-card'));
                return { count: styles.length, hasKeyframes: styles.some(s => s.textContent.includes('@keyframes om-spin')), hasReducedMotion: styles.some(s => s.textContent.includes('prefers-reduced-motion')) };
            })()`);
            // 令牌是不是真的在运行时解析出了值（不是「源码里写了」而已）。
            findings.dom.tokens = await evaluate(`(() => {
                const names = ['--dsw-alias-bg-layer-1','--dsw-alias-bg-layer-2','--dsw-alias-label-primary','--dsw-alias-label-secondary','--dsw-alias-label-tertiary','--dsw-alias-border-l1','--dsw-alias-border-l2','--dsw-alias-state-business-primary','--dsw-alias-state-success-primary','--dsw-alias-state-warn-primary','--dsw-elevation-soft','--dsw-font-family','--dsw-specific-input-major','--dsw-alias-scrollbar-bg-l2'];
                const cs = getComputedStyle(document.body);
                const out = {};
                for (const name of names) out[name] = (cs.getPropertyValue(name)||'').trim().slice(0, 48);
                return out;
            })()`);
            findings.dom.panelBox = await evaluate(`(() => {
                const el = document.querySelector('.om-mem');
                if (!el) return null;
                const r = el.getBoundingClientRect();
                return { width: Math.round(r.width), height: Math.round(r.height), cards: el.querySelectorAll('.om-card').length, inspector: !!el.querySelector('[data-graph-detail]') };
            })()`);
            // 容量口径：归档摘要要按**摘要文件数**报（1 / 60），不是条目数（30 / 60）。
            findings.dom.gauges = await evaluate(`(() => Array.from(document.querySelectorAll('[data-gauge]')).map((el) => ({
                key: el.getAttribute('data-gauge'), percent: el.getAttribute('data-percent'),
                text: (el.innerText||'').replace(/\\s+/g,' ').trim().slice(0, 60),
            })))()`);
            findings.dom.storeChips = await evaluate(`(() => { const el = document.querySelector('[data-store]'); return el ? (el.innerText||'').replace(/\\s+/g,' ').trim().slice(0, 120) : null; })()`);
        };

        await snapshotDom();
        await shot('memory-panel-hot.png');

        // 关系页签
        await evaluate(`(() => { const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => /关系/.test(e.innerText||'')); if (t) { t.click(); return true; } return false; })()`);
        await sleep(2000);
        await snapshotDom();
        await shot('memory-panel-graph.png');

        // 实体页签
        await evaluate(`(() => { const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => /实体/.test(e.innerText||'')); if (t) { t.click(); return true; } return false; })()`);
        await sleep(2000);
        await snapshotDom();
        await shot('memory-panel-entities.png');

        // 关系页签 + 聚焦一个节点：验「悬停/键盘聚焦 → 右侧详情栏」这条交互
        await evaluate(`(() => { const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => /关系/.test(e.innerText||'')); if (t) t.click(); return true; })()`);
        await sleep(1500);
        findings.dom.focusedNode = await evaluate(`(() => {
            const node = document.querySelector('g[data-node]');
            if (!node) return null;
            node.focus();
            return node.getAttribute('data-node');
        })()`);
        await sleep(800);
        findings.dom.inspectorAfterFocus = await evaluate(`(() => { const el = document.querySelector('[data-graph-detail]'); return el ? { key: el.getAttribute('data-graph-detail'), text: (el.innerText||'').trim().slice(0, 200) } : null; })()`);
        await shot('memory-panel-graph-focus.png');

        // 深色主题：宿主靠 body[data-ds-dark-theme] 切令牌，面板应当整体跟着变
        findings.dom.darkTokens = await evaluate(`(() => {
            document.body.setAttribute('data-ds-dark-theme', '');
            const cs = getComputedStyle(document.body);
            return { layer1: (cs.getPropertyValue('--dsw-alias-bg-layer-1')||'').trim(), label: (cs.getPropertyValue('--dsw-alias-label-primary')||'').trim() };
        })()`);
        await sleep(900);
        await shot('memory-panel-dark.png');

        // 宽屏：两栏（图谱 + 详情 / 容量 + 存储域）应当并排；窄屏塌成一列。
        await evaluate(`document.body.removeAttribute('data-ds-dark-theme')`);
        await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1680, height: 950, deviceScaleFactor: 1, mobile: false }, sessionId);
        await sleep(900);
        findings.dom.wideLayout = await evaluate(`(() => {
            const area = document.querySelector('[data-graph-detail]');
            const wrap = area ? area.parentElement : null;
            const cols = wrap ? getComputedStyle(wrap).gridTemplateColumns : null;
            const grid = document.querySelector('.om-grid-dash');
            return { graphColumns: cols, panelColumns: grid ? getComputedStyle(grid).gridTemplateColumns : null, tileColumns: (() => { const t = document.querySelector('[data-metrics]'); return t ? getComputedStyle(t).gridTemplateColumns : null; })() };
        })()`);
        await shot('memory-panel-wide.png');
        await cdp.send('Emulation.clearDeviceMetricsOverride', {}, sessionId);
        await sleep(400);
        findings.dom.narrowLayout = await evaluate(`(() => {
            const area = document.querySelector('[data-graph-detail]');
            const wrap = area ? area.parentElement : null;
            const root = document.querySelector('.om-mem');
            const styleEl = Array.from(document.querySelectorAll('style')).find(s => (s.textContent||'').includes('.om-card'));
            return {
                graphColumns: wrap ? getComputedStyle(wrap).gridTemplateColumns : null,
                containerType: root ? getComputedStyle(root).containerType : null,
                rootWidth: root ? Math.round(root.getBoundingClientRect().width) : null,
                supportsContainer: CSS.supports('container-type', 'inline-size'),
                styleText: styleEl ? styleEl.textContent.length : 0,
            };
        })()`);

        // ── 第十三轮：写入指令 / 分页 / 完整内容 / 视图变换 / 减少动效 ──
        //
        // 这一组必须用**真实输入事件**验，不能只看 DOM：滚轮缩放靠的是一个自己挂的
        // 非被动监听（React 的 onWheel 是被动的，preventDefault 会被丢掉并留下
        // console 警告），「页面有没有跟着滚」只有真的发一次 wheel 才看得见；
        // 拖拽靠 pointerdown/move/up 的真实序列，单测里那些 handler 是直接调用的。
        const rectOf = (selector) => `(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return null; el.scrollIntoView({block:'center'}); const r = el.getBoundingClientRect(); return { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }; })()`;

        const toGraphTab = async () => {
            await evaluate(`(() => { const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => /关系/.test(e.innerText||'')); if (t) t.click(); return true; })()`);
            await sleep(1500);
        };
        const toLedgerTab = async () => {
            await evaluate(`(() => { const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => /台账/.test(e.innerText||'')); if (t) t.click(); return true; })()`);
            await sleep(1500);
        };
        const graphState = () => evaluate(`(() => { const s = document.querySelector('svg[data-layout]'); if (!s) return null; const view = s.querySelector('[data-graph-view]'); return { layout: s.getAttribute('data-layout'), zoom: s.getAttribute('data-zoom'), pan: s.getAttribute('data-pan'), dragging: s.getAttribute('data-dragging'), view: view ? view.getAttribute('transform') : null, scrollY: Math.round(window.scrollY) }; })()`);

        await toGraphTab();
        findings.dom.graphBefore = await graphState();
        const canvas = await evaluate(rectOf('svg[data-layout]'));
        findings.dom.graphRect = canvas;
        if (canvas) {
            const cx = canvas.x + Math.round(canvas.width / 2);
            const cy = canvas.y + Math.round(canvas.height / 2);
            // 第二十轮起滚轮缩放要**按住 Ctrl**（普通滚轮留给面板滚动）：
            // 先发一次不带修饰键的，验「页面/面板真的跟着滚了、画布没缩放」；
            // 再发一次带 Ctrl 的，验缩放仍然生效。
            findings.dom.plainWheelBefore = await graphState();
            findings.dom.panelScrollBeforeWheel = await evaluate(`(() => { const p = document.querySelector('[data-memory-panel="scroll"]'); return p ? Math.round(p.scrollTop) : null; })()`);
            await cdp.send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: cx, y: cy, deltaX: 0, deltaY: 220 }, sessionId);
            await sleep(700);
            findings.dom.panelScrollAfterPlainWheel = await evaluate(`(() => { const p = document.querySelector('[data-memory-panel="scroll"]'); return p ? Math.round(p.scrollTop) : null; })()`);
            findings.dom.graphAfterPlainWheel = await graphState();
            await cdp.send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: cx, y: cy, deltaX: 0, deltaY: -120, modifiers: 2 }, sessionId);
            await sleep(700);
            findings.dom.graphAfterWheel = await graphState();

            // 拖背景平移：从画布左下角（节点够不到的地方）往右下拖
            const px = canvas.x + 8;
            const py = canvas.y + canvas.height - 8;
            await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: px, y: py, button: 'left', clickCount: 1, buttons: 1 }, sessionId);
            await sleep(200);
            await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: px + 40, y: py - 26, button: 'left', buttons: 1 }, sessionId);
            await sleep(200);
            findings.dom.graphWhilePanning = await graphState();
            await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: px + 40, y: py - 26, button: 'left', clickCount: 1, buttons: 0 }, sessionId);
            await sleep(500);
            findings.dom.graphAfterPan = await graphState();

            // 拖节点：取第一个节点的屏幕中心，往左上拖
            const nodeRect = await evaluate(rectOf('g[data-node]'));
            const nodeBefore = await evaluate(`(() => { const g = document.querySelector('g[data-node]'); return g ? { key: g.getAttribute('data-node'), transform: g.getAttribute('transform') } : null; })()`);
            if (nodeRect) {
                const nx = nodeRect.x + Math.round(nodeRect.width / 2);
                const ny = nodeRect.y + Math.round(nodeRect.height / 2);
                await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: nx, y: ny, button: 'left', clickCount: 1, buttons: 1 }, sessionId);
                await sleep(200);
                await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: nx - 30, y: ny - 18, button: 'left', buttons: 1 }, sessionId);
                await sleep(200);
                await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: nx - 30, y: ny - 18, button: 'left', clickCount: 1, buttons: 0 }, sessionId);
                await sleep(500);
            }
            findings.dom.graphNodeDrag = {
                before: nodeBefore,
                after: await evaluate(`(() => { const g = document.querySelector('g[data-node]'); return g ? { key: g.getAttribute('data-node'), transform: g.getAttribute('transform') } : null; })()`),
            };

            // 布局动作：均匀重置要真的换一张坐标表
            findings.dom.graphLayoutClick = await evaluate(`(() => { const b = document.querySelector('[data-graph-layout="uniform"]'); if (!b) return null; b.click(); return true; })()`);
            await sleep(700);
            findings.dom.graphAfterLayout = await graphState();
            await shot('memory-panel-graph-view.png');
        }

        // 分页：台账合成快照有 25 条，默认只该铺 20 条
        await toLedgerTab();
        findings.dom.paginationBefore = await evaluate(`(() => ({
            cards: document.querySelectorAll('[data-item-key]').length,
            showMore: (() => { const b = document.querySelector('[data-show-more]'); return b ? (b.innerText||'').trim() : null; })(),
            note: (() => { const n = document.querySelector('[data-list-more]'); return n ? (n.innerText||'').replace(/\\s+/g,' ').trim().slice(0, 80) : null; })(),
        }))()`);
        findings.dom.paginationClick = await evaluate(`(() => { const b = document.querySelector('[data-show-more]'); if (!b) return null; b.click(); return true; })()`);
        await sleep(900);
        findings.dom.paginationAfter = await evaluate(`(() => ({
            cards: document.querySelectorAll('[data-item-key]').length,
            showMore: !!document.querySelector('[data-show-more]'),
            collapse: !!document.querySelector('[data-collapse-list]'),
            nativeTitles: Array.from(document.querySelectorAll('[data-item-key]')).filter(e => e.hasAttribute('title')).length,
        }))()`);
        await shot('memory-panel-pagination.png');

        // 完整内容：长条目的「展开」去掉限高（合成快照里 m-1 是长文本）
        await evaluate(`(() => { const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => /热记忆/.test(e.innerText||'')); if (t) t.click(); return true; })()`);
        await sleep(1200);
        findings.dom.expandBefore = await evaluate(`(() => { const b = document.querySelector('[data-item-expand]'); const body = document.querySelector('[data-item-body]'); if (!b || !body) return null; const before = getComputedStyle(body).maxHeight; b.click(); return { label: (b.innerText||'').trim(), maxHeightBefore: before }; })()`);
        await sleep(700);
        findings.dom.expandAfter = await evaluate(`(() => { const body = document.querySelector('[data-item-body]'); const b = document.querySelector('[data-item-expand]'); return { maxHeight: body ? getComputedStyle(body).maxHeight : null, label: b ? (b.innerText||'').trim() : null, open: body ? body.getAttribute('data-item-open') : null }; })()`);

        // 详情条：把鼠标移到卡片上，完整内容要出现在面板底部的 ITEM DETAIL 里
        const cardRect = await evaluate(rectOf('[data-item-key]'));
        if (cardRect) {
            await cdp.send('Input.dispatchMouseEvent', {
                type: 'mouseMoved',
                x: cardRect.x + Math.round(cardRect.width / 2),
                y: cardRect.y + Math.round(cardRect.height / 2),
            }, sessionId);
            await sleep(600);
        }
        findings.dom.itemDetail = await evaluate(`(() => { const el = document.querySelector('[data-item-detail]'); return el ? { key: el.getAttribute('data-item-detail'), text: (el.innerText||'').slice(0, 220) } : null; })()`);

        // 写入指令：点一个条目动作，待办区必须出现一段 office_memory 调用。
        // **不能在同一次 evaluate 里点完就读** —— React 的状态更新是异步的，
        // 同一个同步块里读到的还是旧 DOM（这一条第一次写就踩了：pending 一直是 null，
        // 看着像功能没做，其实是探针读早了）。
        findings.dom.writeClicked = await evaluate(`(() => { const b = document.querySelector('[data-memory-action="remove"]'); if (!b) return null; b.click(); return (b.innerText||'').trim(); })()`);
        await sleep(900);
        findings.dom.writeInstruction = await evaluate(`(() => {
            const card = document.querySelector('[data-memory-pending]');
            const pre = document.querySelector('[data-memory-instruction]');
            return {
                pending: card ? card.getAttribute('data-memory-pending') : null,
                text: pre ? (pre.innerText||'').slice(0, 300) : null,
                copy: !!document.querySelector('[data-memory-copy]'),
                jump: !!document.querySelector('[data-memory-jump]'),
                dismiss: !!document.querySelector('[data-memory-dismiss]'),
            };
        })()`);
        await sleep(400);
        // 截图前把待办卡滚进视野：列表已经滚到下面了，不滚的话截到的还是列表，
        // 而这一张要看的正是顶部那张待办指令卡。
        await evaluate(`(() => { const el = document.querySelector('[data-memory-pending]'); if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' }); return !!el; })()`);
        await sleep(700);
        await shot('memory-panel-write.png');

        // 减少动效：把 prefers-reduced-motion 打开，列表容器的 scroll-behavior 必须变成 auto
        findings.dom.scrollBehavior = {};
        findings.dom.scrollBehavior.normal = await evaluate(`(() => { const el = document.querySelector('.om-scroll'); return el ? getComputedStyle(el).scrollBehavior : null; })()`);
        await cdp.send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'reduce' }] }, sessionId);
        await sleep(600);
        findings.dom.scrollBehavior.reduced = await evaluate(`(() => { const el = document.querySelector('.om-scroll'); return el ? getComputedStyle(el).scrollBehavior : null; })()`);
        findings.dom.scrollBehavior.matches = await evaluate(`window.matchMedia('(prefers-reduced-motion: reduce)').matches`);
        await cdp.send('Emulation.setEmulatedMedia', { features: [] }, sessionId);
        await sleep(400);
        findings.dom.scrollBehavior.cleared = await evaluate(`(() => { const el = document.querySelector('.om-scroll'); return el ? getComputedStyle(el).scrollBehavior : null; })()`);

        // ── 第二十轮：记忆面板的滚动契约 + 设置页（真实布局 / 真实字号）──
        log('\n── 第二十轮：面板滚动与设置页 ──');
        findings.dom.scrollContract = await evaluate(`(() => {
            const panel = document.querySelector('[data-memory-panel="scroll"]');
            if (!panel) return null;
            const list = document.querySelector('[data-memory-list]');
            const center = panel.closest('[class*="centerCol"]');
            const last = panel.lastElementChild;
            // 量之前先关掉平滑滚动：.om-scroll 带 scroll-behavior:smooth，赋值 scrollTop
            // 之后立刻读会读到旧值（动画还没跑完）—— 量的是「能不能滚到底」，
            // 不是「动画多快」，所以这里临时切成 auto。
            const previousBehavior = panel.style.scrollBehavior;
            panel.style.scrollBehavior = 'auto';
            const before = panel.scrollTop;
            panel.scrollTop = panel.scrollHeight;
            const after = panel.scrollTop;
            const panelRect = panel.getBoundingClientRect();
            const lastRect = last ? last.getBoundingClientRect() : null;
            const visible = lastRect ? (lastRect.bottom <= panelRect.bottom + 1 && lastRect.top >= panelRect.top - 1) : null;
            // 头部吸顶：滚到底之后它的顶边仍然贴着面板顶边（不然搜索框会滑出视野）。
            // 面板有 padding-top，sticky 贴的是滚动视口上沿，所以留 8px 容差。
            const head = panel.querySelector('[data-memory-head="sticky"]');
            const headDelta = head ? Math.round(head.getBoundingClientRect().top - panelRect.top) : null;
            const headStuck = headDelta === null ? null : Math.abs(headDelta) <= 8;
            panel.scrollTop = before;
            panel.style.scrollBehavior = previousBehavior;
            return {
                overflowY: getComputedStyle(panel).overflowY,
                clientHeight: panel.clientHeight,
                scrollHeight: panel.scrollHeight,
                centerHeight: center ? center.clientHeight : null,
                fillsCenter: center ? Math.abs(panel.clientHeight - center.clientHeight) <= 1 : null,
                scrollable: panel.scrollHeight > panel.clientHeight,
                scrolledToBottom: after > before,
                tailVisibleAfterScroll: visible,
                headStuckAtBottom: headStuck,
                headTopDeltaAtBottom: headDelta,
                listMaxHeight: list ? getComputedStyle(list).maxHeight : null,
                listOverflowY: list ? getComputedStyle(list).overflowY : null,
            };
        })()`);
        log(`滚动契约：${JSON.stringify(findings.dom.scrollContract)}`);
        await shot('memory-panel-scroll.png');

        // 设置页：打开设置对话框 → 办公模式 → 量字号与控件
        findings.dom.settings = {};
        findings.dom.settings.opened = await evaluate(`(() => {
            const button = Array.from(document.querySelectorAll('button')).find((b) => (b.innerText || '').trim() === '设置');
            if (!button) return null;
            button.click();
            return true;
        })()`);
        await sleep(1200);
        findings.dom.settings.navs = await evaluate(`(() => Array.from(document.querySelectorAll('button,[role="tab"],[role="button"]')).map((e) => (e.innerText || '').trim()).filter((t) => t !== '' && t.length <= 24).slice(0, 60))()`);
        findings.dom.settings.officeNav = await evaluate(`(() => {
            const item = Array.from(document.querySelectorAll('button,[role="tab"],[role="button"]')).find((e) => (e.innerText || '').trim() === '办公模式');
            if (!item) return false;
            item.click();
            return true;
        })()`);
        await sleep(1000);
        findings.dom.settings.sectionText = await evaluate(`(() => {
            const panel = document.querySelector('[data-fold="search-advanced"]');
            const root = panel ? panel.closest('div') : null;
            return root ? (root.innerText || '').slice(0, 1500) : (document.body.innerText || '').slice(0, 1500);
        })()`);
        findings.dom.settings.controls = await evaluate(`(() => {
            const pick = (label) => Array.from(document.querySelectorAll('div')).find((e) => (e.innerText || '').trim().startsWith(label));
            const fonts = {};
            for (const label of ['执行引擎', '联网通道', 'auto 尝试顺序', '网页预处理', '短行阈值']) {
                const row = pick(label);
                if (!row) continue;
                const input = row.querySelector('input,select');
                const title = row.firstElementChild && row.firstElementChild.firstElementChild;
                fonts[label] = {
                    title: title ? getComputedStyle(title).fontSize : null,
                    note: title && title.nextElementSibling ? getComputedStyle(title.nextElementSibling).fontSize : null,
                    control: input ? getComputedStyle(input).fontSize : null,
                };
            }
            return {
                fonts,
                providerSelectOptions: (() => {
                    const sels = Array.from(document.querySelectorAll('select'));
                    const sel = sels.find((s) => Array.from(s.options).some((o) => o.value === 'bocha'));
                    return sel ? Array.from(sel.options).map((o) => o.value) : null;
                })(),
                providerOrderChips: (() => {
                    const el = document.querySelector('[data-provider-order]');
                    return el ? el.getAttribute('data-provider-order') : null;
                })(),
                channelKeyInputs: document.querySelectorAll('[data-channel-key]').length,
                toolsAllButtons: document.querySelectorAll('[data-tools-all],[data-tools-none]').length,
            };
        })()`);
        // 折叠块：点开之后要真的铺开（React 状态更新是异步的，读之前让一帧）。
        findings.dom.settings.foldBefore = await evaluate(`(() => { const b = document.querySelector('[data-fold-body="search-advanced"]'); return b ? getComputedStyle(b).display : null; })()`);
        findings.dom.settings.foldClick = await evaluate(`(() => { const t = document.querySelector('[data-fold-toggle="search-advanced"]'); if (!t) return false; t.click(); return true; })()`);
        await sleep(700);
        findings.dom.settings.foldAfter = await evaluate(`(() => { const b = document.querySelector('[data-fold-body="search-advanced"]'); return b ? getComputedStyle(b).display : null; })()`);
        findings.dom.settings.foldBodyRows = await evaluate(`(() => {
            const body = document.querySelector('[data-fold-body="search-advanced"]');
            if (!body) return null;
            return Array.from(body.querySelectorAll('input,select')).slice(0, 8).map((el) => ({ tag: el.tagName.toLowerCase(), type: el.type || null, value: String(el.value ?? '').slice(0, 40) }));
        })()`);
        log(`设置页控件：${JSON.stringify(findings.dom.settings.controls)}`);
        await shot('settings-office.png');
        ws.close();
    } finally {
        killTree(chrome);
        killTree(dshChild);
    }

    const outFile = path.join(OUT, 'web-verify.json');
    writeFileSync(outFile, JSON.stringify(findings, null, 2), 'utf8');
    log('\n===== 结论 =====');
    log(`bootOk（无 "did not activate" 横幅）：${findings.bootOk}`);
    log(`未捕获异常：${findings.exceptions.length} 条`); findings.exceptions.slice(0, 4).forEach((e) => log('   ! ' + String(e).split('\n')[0]));
    log(`控制台 error/warning：${findings.consoleErrors.length} 条`); findings.consoleErrors.slice(0, 4).forEach((e) => log('   ! ' + e.slice(0, 140)));
    log(`拦截到的记忆端点请求：${findings.intercepted} 次（Network 观察到 ${findings.snapshotRequests ?? 0} 次：${findings.snapshotUrl ?? '-'}）`);
    log(`侧栏「记忆」入口：${JSON.stringify(findings.dom.clickedMemory)}`);
    log(`页签：${JSON.stringify(findings.dom.tabs)}`);
    log(`关系图：${JSON.stringify(findings.dom.graphMeta)}`);
    log(`图元：${JSON.stringify(findings.dom.svgShapes)}`);
    log(`样式表节点：${JSON.stringify(findings.dom.styleNode)}`);
    log(`面板盒：${JSON.stringify(findings.dom.panelBox)}`);
    log(`容量条：${JSON.stringify(findings.dom.gauges)}`);
    log(`存储域 chips：${JSON.stringify(findings.dom.storeChips)}`);
    log(`聚焦节点：${JSON.stringify(findings.dom.focusedNode)} → ${JSON.stringify(findings.dom.inspectorAfterFocus)}`);
    log(`令牌解析：${JSON.stringify(findings.dom.tokens)}`);
    log(`深色令牌：${JSON.stringify(findings.dom.darkTokens)}`);
    log(`宽屏栅格：${JSON.stringify(findings.dom.wideLayout)}`);
    log(`窄屏栅格：${JSON.stringify(findings.dom.narrowLayout)}`);
    log('');
    log('── 第十三轮：真实输入事件 ──');
    log(`关系图（初始）：${JSON.stringify(findings.dom.graphBefore)}`);
    log(`滚轮放大后：${JSON.stringify(findings.dom.graphAfterWheel)}`);
    log(`拖背景（按住时 / 松手后）：${JSON.stringify(findings.dom.graphWhilePanning)} / ${JSON.stringify(findings.dom.graphAfterPan)}`);
    log(`拖节点：${JSON.stringify(findings.dom.graphNodeDrag)}`);
    log(`切「均匀重置」后：${JSON.stringify(findings.dom.graphAfterLayout)}`);
    log(`分页（点之前）：${JSON.stringify(findings.dom.paginationBefore)}`);
    log(`分页（点之后）：${JSON.stringify(findings.dom.paginationAfter)}`);
    log(`展开长条目：${JSON.stringify(findings.dom.expandBefore)} → ${JSON.stringify(findings.dom.expandAfter)}`);
    log(`条目详情条（悬停）：${JSON.stringify(findings.dom.itemDetail)}`);
    log(`写入指令（点了「${findings.dom.writeClicked}」）：${JSON.stringify(findings.dom.writeInstruction)}`);
    log(`减少动效下的 scroll-behavior：${JSON.stringify(findings.dom.scrollBehavior)}`);
    log('── 第二十轮：滚动契约与设置页 ──');
    log(`普通滚轮：面板 scrollTop ${findings.dom.panelScrollBeforeWheel} → ${findings.dom.panelScrollAfterPlainWheel}；画布 zoom ${findings.dom.graphAfterPlainWheel && findings.dom.graphAfterPlainWheel.zoom}`);
    log(`Ctrl + 滚轮：画布 zoom ${findings.dom.graphAfterWheel && findings.dom.graphAfterWheel.zoom}`);
    log(`面板滚动契约：${JSON.stringify(findings.dom.scrollContract)}`);
    log(`设置页：打开 ${findings.dom.settings && findings.dom.settings.opened} / 办公模式 ${findings.dom.settings && findings.dom.settings.officeNav}`);
    log(`设置页字号与控件：${JSON.stringify(findings.dom.settings && findings.dom.settings.controls)}`);
    log(`折叠块：${findings.dom.settings && findings.dom.settings.foldBefore} →（点 ${findings.dom.settings && findings.dom.settings.foldClick}）→ ${findings.dom.settings && findings.dom.settings.foldAfter}`);
    log(`折叠块里的控件：${JSON.stringify(findings.dom.settings && findings.dom.settings.foldBodyRows)}`);
    log(`面板正文：\n${String(findings.dom.panelText).slice(0, 1200)}`);
    log(`详情写到 ${outFile}`);
}

main().catch((err) => { console.error('探针失败：', err); process.exit(1); });
