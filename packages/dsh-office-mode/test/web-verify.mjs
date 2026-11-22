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
 * `.office/runtime/diag/` —— 与其它中间产物同一棵隐藏树，不进版本库。
 */
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { mkdirSync, writeFileSync, readFileSync, existsSync, openSync, closeSync, rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { dshBinPath } from './host-modules.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '../../..');
const OUT = path.resolve(ROOT, '.office', 'runtime', 'diag');
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
        { id: 'L-1', at, path: '季度汇报.pptx', format: 'ppt', theme: 'business', purpose: '季度汇报', outline: ['封面', '要点', '表格'], origin: 'workspace' },
        { id: 'L-2', at: '2026-09-24T08:00:00.000Z', path: '季度数据.xlsx', format: 'excel', theme: 'plain', purpose: '数据明细', outline: ['表格'], origin: 'workspace' },
        { id: 'L-3', at: '2026-09-23T08:00:00.000Z', path: '季度回顾.docx', format: 'word', theme: 'plain', purpose: '文字稿', outline: ['H1 概述'], origin: 'workspace' },
        { id: 'L-4', at: '2026-09-23T07:00:00.000Z', path: '论文.pdf', format: 'pdf', theme: '', purpose: '论文编译', outline: [], origin: 'workspace' },
    ];
    for (let index = 5; index <= 25; index += 1) {
        const day = String(1 + (index % 22)).padStart(2, '0');
        ledger.push({
            id: 'L-' + index, at: '2026-08-' + day + 'T06:00:00.000Z',
            path: '历史产物-' + index + '.docx', format: 'word', theme: 'plain',
            purpose: '第 ' + index + ' 份历史交付物', outline: ['H1'], origin: 'workspace',
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
            autoCapture: true,
            quality: { policy: 'strict-v1', lowScoreThreshold: 0.2, highScoreThreshold: 0.6, candidateMultiplier: 4, maxMediumResults: 2, maxUnknownResults: 1 },
            quota: { recallPerTurn: 1, recallRefinePerTurn: 1, relatedPerTurn: 1 },
            limits: { userLimitBytes: 4096, projectLimitBytes: 10240, ledgerLimit: 500, archiveKeep: 60 },
        },
        stores: [{ id: 'workspace', dir: '.office/memory', hot: 4, ledger: 25, archive: 30, archiveFiles: 1, bytes: 24576, files: 9 }],
        // 第八十八轮：知识库是单库，不再挂在任何记忆根上。
        kbStore: { id: 'kb', dir: '.office/kb', docs: 1, chunks: 2, bytes: 4096, tiers: { user: 0, verified: 1, unverified: 0 } },
        // 两个口径（见 src/view.js 的注释）：hot/ledger/archive/... 是「列表里现在有几条」，
        // ledgerTotal/archiveFiles/archiveItems 是容量条用的未过滤总数。
        // 这里**故意让归档的条目数（30）与摘要文件数（1）不同** —— 面板曾经拿 30 配上限 60，
        // 显示成「30 / 60 个」像快满了，而实际只用了 1 个文件。
        counts: {
            hot: 4, ledger: 25, archive: 30,
            ledgerTotal: 25, archiveFiles: 1, archiveItems: 30,
            kb: 1, kbTotal: 1, kbChunks: 2, kbBytes: 4096, kbTruncated: false,
            hotBytes: 812, projectBytes: 1830,
        },
        // 第八十二轮：来源分段控件上的三档数字；知识库文档放全局，点「全局」之后
        // 还能看到一份，KB 页签不会立刻掉进空态。
        originCounts: {
            hot: { all: 4, global: 2, workspace: 2 },
            ledger: { all: 25, global: 0, workspace: 25 },
            archive: { all: 30, global: 0, workspace: 30 },
        },
        kbSearch: null,
        kbSearchError: '',
        kbDoc: null,
        kbDocError: '',
        kbChunk: null,
        kbChunkError: '',
        kbFolder: '',
        kbRecursive: true,
        kbTree: {
            path: '', name: '全部文档', docs: 1, chunks: 2, bytes: 4096,
            children: [
                { path: '参考', name: '参考', docs: 1, chunks: 2, bytes: 4096, children: [] },
                // 会话范围回归用的合成目录：picker 只读目录名，不读这些目录里的文档。
                // 「来源/方法」同时验证子目录也会出现在同一个选择器里。
                {
                    path: '来源', name: '来源', docs: 1, chunks: 2, bytes: 4096,
                    children: [{ path: '来源/方法', name: '方法', docs: 1, chunks: 2, bytes: 4096, children: [] }],
                },
                { path: '归档', name: '归档', docs: 1, chunks: 2, bytes: 4096, children: [] },
            ],
        },
        kb: [{
            id: 'kb:8f14e45f', path: 'docs/方法论.md', folder: '参考', title: '方法论笔记', hash: '8f14e45f',
            bytes: 4096, chars: 2048, chunks: 2, tier: 'verified', at,
            origin: 'kb',
        }],
        hot: [
            // m-1 故意写成长文本：面板对长条目给「展开」按钮，展开要去掉正文限高。
            { id: 'm-1', target: 'user', importance: 'critical', origin: 'workspace', updatedAt: at, entities: [], tags: ['写作'], content: '偏好：书面化表达，避免说教与第二人称。' + '这条故意写得很长，用来验证「展开」按钮会去掉正文的限高。'.repeat(6) },
            { id: 'm-2', target: 'project', importance: 'critical', origin: 'workspace', updatedAt: at, entities: ['系统A'], tags: ['ooxml'], content: '约定：改 OOXML 必须读回原包、只改命中片段、原样写回。' },
            { id: 'm-3', target: 'project', importance: 'normal', origin: 'global', updatedAt: at, entities: ['系统A', '系统B'], tags: [], content: '环境：本机 TeX Live 2022，latexmk -xelatex 可直接编译。' },
            { id: 'm-4', target: 'user', importance: 'normal', origin: 'global', updatedAt: at, entities: [], tags: ['界面'], content: '偏好：面板里的写入动作只生成指令，不直接落盘。' },
        ],
        ledger,
        archive: [
            { id: 'A-1', month: '2026-09', kind: 'insight', at, origin: 'workspace', importance: 'normal', category: '环境', entities: ['系统B'], tags: ['tex'], text: '长期记忆一条：TeX 引擎探测顺序 fitz > poppler > mupdf。' },
            { id: 'A-2', month: '2026-08', kind: 'insight', at: '2026-08-20T08:00:00.000Z', origin: 'workspace', importance: 'low', category: '工具', entities: [], tags: [], text: '旧的 mnemon 记录：缓存 TTL 默认 12 小时。' },
        ],
    };
}

function safeLogTail(file, limit) {
    try {
        return readFileSync(file, 'utf8').slice(-limit).replace(/token=[^\s&]+/g, 'token=***');
    } catch {
        return '';
    }
}

async function startDsh() {
    const out = path.join(OUT, 'web-out.log');
    const err = path.join(OUT, 'web-err.log');
    const fdOut = openSync(out, 'w');
    const fdErr = openSync(err, 'w');
    const startedAt = Date.now();
    const child = spawn(process.execPath, [DSH_BIN, 'web', '--port', '0', '--no-open'], {
        cwd: ROOT, stdio: ['ignore', fdOut, fdErr], env: { ...process.env, DSH_NO_OPEN: '1' },
    });
    closeSync(fdOut);
    closeSync(fdErr);
    let childExit = null;
    child.once('error', (error) => { childExit = childExit ?? { error }; });
    child.once('exit', (code, signal) => { childExit = childExit ?? { code, signal }; });
    const exitDetail = () => childExit?.error
        ? `spawn error: ${childExit.error.message}`
        : `exit code ${childExit?.code ?? 'null'}, signal ${childExit?.signal ?? 'null'}`;
    // 冷启动需要等宿主 loader；实测 market boot 在第 304 秒才开始。
    // 只放宽测试宿主启动预算，不延长浏览器断言等待。
    for (let i = 0; i < 1200; i += 1) {
        await sleep(500);
        if (childExit) {
            throw new Error(`dsh web 在打印地址前退出（${exitDetail()}）。\n--- web-err.log 尾部 ---\n${safeLogTail(err, 4000) || '(空)'}\n--- web-out.log 尾部 ---\n${safeLogTail(out, 2000) || '(空)'}`);
        }
        if (!existsSync(out)) continue;
        const m = readFileSync(out, 'utf8').match(/https?:\/\/127\.0\.0\.1:\d+\/[^\s]*token=[A-Za-z0-9._-]+/);
        if (m) return { child, url: m[0] };
    }
    const elapsed = Math.round((Date.now() - startedAt) / 1000);
    const status = childExit ? exitDetail() : `still running after ${elapsed}s (pid ${child.pid})`;
    killTree(child);
    throw new Error(`没能解析出带 token 的地址（${status}）。\n--- web-err.log 尾部 ---\n${safeLogTail(err, 4000) || '(空)'}\n--- web-out.log 尾部 ---\n${safeLogTail(out, 2000) || '(空)'}`);
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

const findings = {
    bootOk: null, exceptions: [], consoleErrors: [], failedRequests: [], intercepted: 0,
    galleryRequests: [], kbLlmRequests: [], dom: {}, screenshots: [],
    // 会话知识库范围回归：只记命令文本/回执，不记真实 session id，也不转发任何命令。
    scopeCommands: [], scopeReplies: [], scopeRefused: [], scopeRequestShapes: [], scopeCommandHits: 0,
};

function isRecord(value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function exactKeys(value, expected) {
    if (!isRecord(value)) return false;
    const keys = Object.keys(value).sort();
    return keys.length === expected.length && keys.every((key, index) => key === expected[index]);
}

function scopeEnvelopeProblem(envelope) {
    if (!isRecord(envelope)) return 'envelope-not-record';
    if (envelope.type !== 'client-request') return 'type';
    if (envelope.method !== 'commands/execute') return 'method';
    if (typeof envelope.rpcId !== 'string' || envelope.rpcId.length === 0) return 'rpcId';
    if (!isRecord(envelope.payload) || !exactKeys(envelope.payload, ['args'])) return 'payload';
    const args = envelope.payload.args;
    if (!isRecord(args) || !exactKeys(args, ['agentId', 'line', 'submittedAttachments'])) return 'args';
    if (typeof args.agentId !== 'string' || args.agentId.length === 0) return 'agentId';
    if (typeof args.line !== 'string' || args.line.length === 0) return 'line';
    if (!Array.isArray(args.submittedAttachments)) return 'submittedAttachments';
    if (args.submittedAttachments.length !== 0) return 'submittedAttachments-not-empty';
    return '';
}

function scopeRequestShape(request, endpoint, postDataInfo, envelope, problem) {
    const payload = isRecord(envelope) && isRecord(envelope.payload) ? envelope.payload : null;
    const args = payload !== null && isRecord(payload.args) ? payload.args : null;
    const keys = (value) => isRecord(value) ? Object.keys(value).sort() : [];
    let pathname = '';
    try { pathname = new URL(String(request?.url || '')).pathname; } catch { pathname = ''; }
    return {
        endpoint: endpoint,
        method: typeof request?.method === 'string' ? request.method : '',
        path: pathname,
        postDataSource: postDataInfo.source,
        postDataPresent: postDataInfo.text.length > 0,
        postDataErrorCode: postDataInfo.errorCode || null,
        envelopeType: isRecord(envelope) && typeof envelope.type === 'string' ? envelope.type : null,
        envelopeMethod: isRecord(envelope) && typeof envelope.method === 'string' ? envelope.method : null,
        rpcIdIsString: isRecord(envelope) && typeof envelope.rpcId === 'string',
        payloadKeys: keys(payload),
        argsKeys: keys(args),
        agentIdIsString: args !== null && typeof args.agentId === 'string',
        lineIsString: args !== null && typeof args.line === 'string',
        lineLength: args !== null && typeof args.line === 'string' ? args.line.length : null,
        submittedAttachmentsIsArray: args !== null && Array.isArray(args.submittedAttachments),
        submittedAttachmentsLength: args !== null && Array.isArray(args.submittedAttachments) ? args.submittedAttachments.length : null,
        problem: problem || null,
    };
}

async function readFetchPostData(cdp, msg) {
    const request = msg.params.request || {};
    if (typeof request.postData === 'string' && request.postData.length > 0) {
        return { text: request.postData, source: 'fetch.request.postData' };
    }
    const networkId = msg.params.networkId;
    if (typeof networkId !== 'string' || networkId.length === 0) {
        return { text: typeof request.postData === 'string' ? request.postData : '', source: 'unavailable' };
    }
    try {
        const body = await cdp.send('Network.getRequestPostData', { requestId: networkId }, msg.sessionId);
        if (typeof body?.postData === 'string') {
            return { text: body.postData, source: 'network.getRequestPostData' };
        }
        return { text: '', source: 'network.getRequestPostData', errorCode: 'missing-postData' };
    } catch (error) {
        return { text: '', source: 'network.getRequestPostData', errorCode: typeof error?.code === 'string' ? error.code : 'error' };
    }
}

async function main() {
    mkdirSync(OUT, { recursive: true });
    const snapshot = makeSnapshot();
    const snapshotBody = Buffer.from(JSON.stringify(snapshot), 'utf8').toString('base64');
    // 第八十二轮：拦截端点要按请求参数回不同的快照，否则「点全局」只是按钮变色，
    // 数据还是全量 —— 截图上看起来像筛选坏了。这里按 origin 过滤，按 kbq 给命中，
    // 按 kbdoc / kbchunk 给下钻数据。
    const kbHitFixture = {
        chunkId: 'kb:8f14e45f:1', docId: 'kb:8f14e45f', path: 'docs/方法论.md', title: '方法论笔记',
        tier: 'verified', origin: 'kb', n: 1, score: 1.23, scoreText: 1.1, scoreHead: 0.2,
        header: '# 方法论', spanText: '字符 1–18', chars: 18, preview: '复核要能复跑。',
    };
    const kbDocFixture = {
        docId: 'kb:8f14e45f', path: 'docs/方法论.md', title: '方法论笔记', tier: 'verified', at: '2026-09-24T09:12:00.000Z',
        origin: 'kb', selectedChunk: '', chunks_total: 2, chunks_missing: 0, total: 2,
        chunks: [
            { id: 'kb:8f14e45f:1', n: 1, chars: 7, span: '0-7', spanText: '字符 1–7', header: '# 方法论', outline: ['方法论'] },
            { id: 'kb:8f14e45f:2', n: 2, chars: 9, span: '8-16', spanText: '字符 8–16', header: '## 复核', outline: ['方法论', '复核'] },
        ],
    };
    const kbChunkFixture = {
        id: 'kb:8f14e45f:1', n: 1, text: '复核要能复跑。', spanText: '字符 1–7', header: '# 方法论',
        chars: 7, docId: 'kb:8f14e45f', path: 'docs/方法论.md', title: '方法论笔记', tier: 'verified', origin: 'kb',
    };
    const snapshotBodyFor = (requestUrl) => {
        let params = null;
        try { params = new URL(requestUrl).searchParams; } catch { /* ignore */ }
        const origin = params === null ? '' : (params.get('origin') || '');
        const kbq = params === null ? '' : (params.get('kbq') || '');
        const kbdoc = params === null ? '' : (params.get('kbdoc') || '');
        const kbchunk = params === null ? '' : (params.get('kbchunk') || '');
        const kbfolder = params === null ? '' : (params.get('kbfolder') || '');
        const kbrecursive = params === null ? true : params.get('kbrecursive') !== '0';
        if (origin === '' && kbq === '' && kbdoc === '' && kbchunk === '' && kbfolder === '') return snapshotBody;
        const copy = JSON.parse(JSON.stringify(snapshot));
        if (origin === 'global' || origin === 'workspace') {
            // 知识库是单库（第八十八轮）：origin 只筛热记忆 / 台账 / 归档。
            for (const layer of ['hot', 'ledger', 'archive']) {
                if (Array.isArray(copy[layer])) copy[layer] = copy[layer].filter((row) => row.origin === origin);
            }
            copy.origin = origin;
            copy.counts = Object.assign({}, copy.counts, {
                hot: copy.hot.length, ledger: copy.ledger.length,
                archive: copy.archive.length,
            });
        }
        if (kbfolder !== '') {
            copy.kb = copy.kb.filter((row) => kbrecursive
                ? row.folder === kbfolder || row.folder.startsWith(kbfolder + '/')
                : row.folder === kbfolder);
            copy.counts = Object.assign({}, copy.counts, { kb: copy.kb.length });
        }
        if (kbfolder === '参考') {
            // 合成 23 篇，确保第二页只能通过后续请求获取，不复用首屏 DOM。
            const exemplar = snapshot.kb[0];
            const documents = [exemplar, ...Array.from({ length: 22 }, (_, i) => ({
                ...exemplar, id: 'kb:fixture-' + (i + 2),
                path: 'docs/分页样例-' + (i + 2) + '.md', title: '分页样例-' + (i + 2),
            }))];
            const limit = Number(params.get('kblimit')) || 20;
            const offset = Number(params.get('kboffset')) || 0;
            copy.kb = documents.slice(offset, offset + limit);
            copy.kbPage = { offset, limit, total: documents.length,
                pageCount: Math.ceil(documents.length / limit), page: Math.floor(offset / limit) + 1,
                hasPrev: offset > 0, hasNext: offset + limit < documents.length };
            copy.counts = { ...copy.counts, kb: documents.length, kbTotal: documents.length,
                kbTruncated: copy.kbPage.hasNext };
        }
        copy.kbFolder = kbfolder;
        copy.kbRecursive = kbrecursive;
        if (kbq !== '') {
            copy.kbSearch = { hits: [kbHitFixture], matched: 1, truncated: false, reason: '' };
        }
        if (kbdoc !== '') copy.kbDoc = kbDocFixture;
        if (kbchunk !== '') copy.kbChunk = kbChunkFixture;
        return Buffer.from(JSON.stringify(copy), 'utf8').toString('base64');
    };

    /*
     * 会话知识库范围回归（只读夹具）：
     *
     * 真实 picker 走 remote.commands.execute -> POST /api/commands/execute。这里用 CDP
     * Fetch 在浏览器网络层把这个 POST 截下来，按下面的虚构状态直接回执；任何不认识的
     * commands 路径都 failRequest，绝不 continueRequest。于是命令不会到达宿主的
     * commands service，更不会碰真实共享会话的写入句柄。
     *
     * 只实现 picker 真正会发的三条命令；其他命令记进 scopeRefused，测试据此判红。
     */
    const scopeFixture = {
        mode: 'unset',
        folder: '',
        calls: [],
        replies: [],
        refused: [],
    };
    const scopeFixtureTree = ['参考', '来源', '来源/方法', '归档'];
    // 只接受 picker 真的会发的完整命令行；其余命令一律拒绝，且诊断里不保存原文。
    const scopeAllowedLines = new Set(['/kb-scope', '/kb-scope clear', ...scopeFixtureTree.map((folder) => '/kb-scope set ' + folder)]);
    const scopeStatusText = () => scopeFixture.mode === 'all'
        ? '知识库范围：全部知识库。'
        : scopeFixture.mode === 'folder'
            ? '知识库范围：「' + scopeFixture.folder + '」（含子目录）。'
            : '知识库范围：未设置。这个会话还没有设置知识库范围，所以不能查知识库（不会默认放开全库）。';
    const scopeCommandReply = (line) => {
        const text = String(line || '').trim();
        if (!scopeAllowedLines.has(text)) {
            scopeFixture.refused.push('unknown scope command');
            findings.scopeRefused.push('unknown scope command');
            return { kind: 'error', text: 'fixture 不接受这条命令，已拒绝发出。' };
        }
        scopeFixture.calls.push(text);
        if (text === '/kb-scope') {
            return { kind: 'success', text: scopeStatusText() };
        }
        if (text === '/kb-scope clear') {
            scopeFixture.mode = 'all';
            scopeFixture.folder = '';
            return { kind: 'success', text: '知识库范围：已清空（全部知识库）。' };
        }
        const folder = text.slice('/kb-scope set '.length).trim();
        scopeFixture.mode = 'folder';
        scopeFixture.folder = folder;
        return { kind: 'success', text: '知识库范围：已设为「' + folder + '」（含子目录）。' };
    };
    const scopeEnvelopeBody = (rpcId, reply, index) => Buffer.from(JSON.stringify({
        type: 'server-response',
        rpcId,
        result: {
            ok: true,
            value: {
                commandId: 'cmd-fixture-' + index,
                result: { kind: reply.kind, text: reply.text },
            },
        },
    }), 'utf8').toString('base64');

    // 图库回归：假数据 + 假端点。只证明「点页签 → 请求 → 渲染 → 切回」，
    // 不碰真实图库目录，也不会上传任何文件。
    const GALLERY_ID_A = '0123456789abcdef'.repeat(4);
    const GALLERY_ID_B = 'fedcba9876543210'.repeat(4);
    const galleryStatusFixture = {
        ok: true, root: '.office/kb/gallery', version: 1,
        counts: { items: 2, withThumb: 0, withAesthetic: 0 },
        thumbnail: { configured: true, lastError: '' },
        lexical: { available: true, note: 'SQLite FTS5 可用。' },
        textEmbedding: { mode: 'local-hash', remote: false, model: '', dimensions: 64, note: '默认本地哈希向量；远端文本向量需显式允许。' },
        imageEmbedding: { provider: '', model: '', configured: false, probed: false, dim: 0, note: '未配置图文对齐端点。' },
        vision: { channel: '', provider: '', model: '', configured: false, note: '未配置视觉模型。' },
        uploadPolicy: '默认不联网；描述 / 探针 / 图片向量化 / 以图搜图都需显式允许。',
    };
    const galleryListFixture = {
        ok: true, total: 2, offset: 0, limit: 200,
        items: [
            { id: GALLERY_ID_A, bytes: 2048, mediaType: 'image/png', width: 320, height: 200, addedAt: '2026-09-24T09:12:00.000Z', thumb: null, category: '界面参考', favorite: true, note: '', tags: ['假数据', '界面'], manualTags: ['假数据', '界面'], aesthetic: null },
            { id: GALLERY_ID_B, bytes: 1024, mediaType: 'image/jpeg', width: 160, height: 120, addedAt: '2026-09-23T08:00:00.000Z', thumb: null, category: '', favorite: false, note: '', tags: [], manualTags: [], aesthetic: null },
        ],
    };
    log('① 启动隔离实例 dsh web --port 0 --no-open');
    let dshChild = null;
    let chrome = null;
    try {
        const started = await startDsh();
        dshChild = started.child;
        const url = started.url;
        log(`   地址：${url.replace(/token=[^&]*/, 'token=***')}`);

        log('② 启动无头 Chrome + CDP');
        // 每次清一次临时 Chrome profile。**不清的话这一套会长期红**，而且红得看不出来：
        // 复用的 profile 里攒下一堆 127.0.0.1 的 cookie 之后，DSH 的 web 入口会直接回
        // `HTTP ERROR 431`（请求头字段过大），页面上只有「该网页无法正常运作」——
        // 于是所有 DOM 断言都成了 null，看起来像「记忆面板坏了」。第六十二轮实测：
        // 删掉这份 profile 立刻恢复（settings 段落全绿、面板段也全绿）。
        // 它只是个一次性目录（.office/runtime/diag/chrome-profile），删掉会自动重建。
        const chromeProfile = path.join(OUT, 'chrome-profile');
        if (existsSync(chromeProfile)) rmSync(chromeProfile, { recursive: true, force: true });
        chrome = spawn(CHROME, [
            '--headless=new', `--remote-debugging-port=${PORT}`,
            '--no-first-run', '--no-default-browser-check', '--disable-gpu', '--window-size=1280,900',
            '--user-data-dir=' + chromeProfile, 'about:blank',
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
                // 第六十四轮：把**每一次**请求的 URL 都留下 —— 「默认根＝当前会话的工作目录」
                // 是个关于**第一次**请求的判据（第一次就必须带 cwd），只看最后一次看不出来。
                findings.snapshotUrls = (findings.snapshotUrls ?? []).concat([msg.params.request.url]);
            }
            // 第九十五轮：知识库大模型的宿主清单是**只读** GET，这里不拦截，只记账：
            // 「off 通道不发这条请求」与「切到 host 后只发一次、重试才再发」都靠它。
            // 探针自己也 GET 同一接口做对照；带 x-office-verify-probe: 1 的是探针，
            // 不能算成 UI 的请求（否则日志会把 2 次对照读当成 2 次回退）。
            if (msg.method === 'Network.requestWillBeSent') {
                const request = msg.params.request || {};
                const headers = request.headers || {};
                const probeKey = Object.keys(headers).find((key) => key.toLowerCase() === 'x-office-verify-probe');
                const fromProbe = probeKey !== undefined && String(headers[probeKey]) === '1';
                const url = String(request.url || '');
                if (/office-memory\/llm-models/.test(url) && !fromProbe) {
                    findings.kbLlmRequests.push({ at: Date.now() });
                }
                // 本版本的 session.modelCatalog 走 HTTP POST（同源的 /api/session/modelCatalog），
                // 不是 WebSocket；只盯 webSocketFrameSent 会永远数不到它。
                if (request.method === 'POST' && url !== '') {
                    let pathname = '';
                    try {
                        pathname = new URL(url).pathname;
                    } catch (error) {
                        pathname = '';
                    }
                    if (/\/api\/session\/modelCatalog\/?$/.test(pathname)) {
                        findings.kbLlmRpcHttp = (findings.kbLlmRpcHttp ?? 0) + 1;
                        findings.kbLlmRpcFrames = (findings.kbLlmRpcFrames ?? 0) + 1;
                    }
                }
            }
            // legacy / 其他 carrier 仍可能是 WebSocket 文本帧；保留旧计数，和 HTTP POST
            // 分开记，重试时就能看出这次到底走了哪种 transport。
            if (msg.method === 'Network.webSocketFrameSent' && typeof msg.params.response?.payloadData === 'string') {
                if (msg.params.response.payloadData.indexOf('modelCatalog') !== -1) {
                    findings.kbLlmRpcWs = (findings.kbLlmRpcWs ?? 0) + 1;
                    findings.kbLlmRpcFrames = (findings.kbLlmRpcFrames ?? 0) + 1;
                }
            }
            // 拦截记忆端点：喂一份完整合成快照，让面板渲染出带数据的形态
            if (msg.method === 'Fetch.requestPaused') {
                const rid = msg.params.requestId;
                const request = msg.params.request;
                const requestUrl = String(request.url || '');
                let commandPath = '';
                try { commandPath = new URL(requestUrl).pathname; } catch { commandPath = ''; }
                const commandMatch = /^\/api\/commands\/([^/]+)$/.exec(commandPath);
                if (commandMatch !== null) {
                    // 命令只走这里：execute 按虚构状态应答；list 回空；其他路径一律
                    // failRequest。任何分支都不 continueRequest —— 命令不会到宿主。
                    const endpoint = commandMatch[1];
                    if (endpoint === 'execute') {
                        findings.scopeCommandHits += 1;
                        void (async () => {
                            const postDataInfo = await readFetchPostData(cdp, msg);
                            let envelope = null;
                            try { envelope = JSON.parse(postDataInfo.text); } catch { envelope = null; }
                            const problem = scopeEnvelopeProblem(envelope) || (request.method === 'POST' ? '' : 'request-method');
                            findings.scopeRequestShapes.push(scopeRequestShape(request, endpoint, postDataInfo, envelope, problem));
                            if (problem !== '') {
                                findings.scopeRefused.push('malformed commands/execute POST (' + problem + ')');
                                await cdp.send('Fetch.failRequest', { requestId: rid, errorReason: 'BlockedByClient' }, msg.sessionId ?? sessionId);
                                return;
                            }
                            const line = envelope.payload.args.line;
                            const reply = scopeCommandReply(line);
                            // 只记录允许列表里的命令文本；未知命令只留通用拒绝原因，
                            // 避免把请求体或任意命令原文写进诊断。
                            if (scopeAllowedLines.has(line)) {
                                findings.scopeReplies.push({ line: line, kind: reply.kind, text: reply.text });
                            }
                            await cdp.send('Fetch.fulfillRequest', {
                                requestId: rid, responseCode: 200,
                                responseHeaders: [{ name: 'content-type', value: 'application/json; charset=utf-8' }],
                                body: scopeEnvelopeBody(envelope.rpcId, reply, findings.scopeReplies.length),
                            }, msg.sessionId ?? sessionId);
                        })().catch(() => {
                            findings.scopeRefused.push('commands/execute interceptor error');
                            cdp.send('Fetch.failRequest', { requestId: rid, errorReason: 'BlockedByClient' }, msg.sessionId ?? sessionId).catch(() => { /* ignore */ });
                        });
                        return;
                    }
                    if (endpoint === 'list') {
                        let envelope = null;
                        try { envelope = JSON.parse(String(request.postData || '')); } catch { envelope = null; }
                        if (envelope === null || envelope.type !== 'client-request'
                            || typeof envelope.rpcId !== 'string') {
                            findings.scopeRefused.push('malformed commands/list POST');
                            cdp.send('Fetch.failRequest', { requestId: rid, errorReason: 'BlockedByClient' }, msg.sessionId ?? sessionId).catch(() => { /* ignore */ });
                            return;
                        }
                        cdp.send('Fetch.fulfillRequest', {
                            requestId: rid, responseCode: 200,
                            responseHeaders: [{ name: 'content-type', value: 'application/json; charset=utf-8' }],
                            body: Buffer.from(JSON.stringify({
                                type: 'server-response', rpcId: envelope.rpcId,
                                result: { ok: true, value: [] },
                            }), 'utf8').toString('base64'),
                        }, msg.sessionId ?? sessionId).catch(() => { /* ignore */ });
                        return;
                    }
                    findings.scopeRefused.push('unknown commands carrier');
                    cdp.send('Fetch.failRequest', { requestId: rid, errorReason: 'BlockedByClient' }, msg.sessionId ?? sessionId).catch(() => { /* ignore */ });
                    return;
                }
                if (/\/office-memory\/snapshot/.test(msg.params.request.url)) {
                    findings.intercepted += 1;
                    cdp.send('Fetch.fulfillRequest', {
                        requestId: rid, responseCode: 200,
                        responseHeaders: [{ name: 'content-type', value: 'application/json; charset=utf-8' }],
                        body: snapshotBodyFor(msg.params.request.url),
                    }, msg.sessionId ?? sessionId).catch(() => { /* ignore */ });
                } else if (/\/office-memory\/gallery(\?|$)/.test(msg.params.request.url)) {
                    // 图库回归的假端点：状态与清单都回合成数据，不读真实图库目录。
                    // 只记录方法 / 视图；任何 POST（上传、手工标签）都会被记下来并判失败。
                    const request = msg.params.request;
                    let view = '';
                    try { view = new URL(request.url).searchParams.get('view') || ''; } catch { /* ignore */ }
                    // 只留方法 / 视图 / 有没有 body：不把本地路径或请求体写进产物。
                    findings.galleryRequests.push({
                        method: request.method,
                        view: view || '(none)',
                        hasPostData: Boolean(request.postData),
                    });
                    const payload = view === 'status'
                        ? JSON.stringify(galleryStatusFixture)
                        : JSON.stringify(galleryListFixture);
                    cdp.send('Fetch.fulfillRequest', {
                        requestId: rid, responseCode: 200,
                        responseHeaders: [{ name: 'content-type', value: 'application/json; charset=utf-8' }],
                        body: Buffer.from(payload, 'utf8').toString('base64'),
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
        await cdp.send('Fetch.enable', { patterns: [
            { urlPattern: '*office-memory/snapshot*' },
            { urlPattern: '*office-memory/gallery*' },
            // 范围选择器的命令载体：必须在任何点击之前就挂上。两个模式重叠时
            // handler 自己按 endpoint 分派；未识别的 commands 路径一律 failRequest。
            { urlPattern: '*api/commands/execute*' },
            { urlPattern: '*api/commands/*' },
        ] }, sessionId);

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
        /** 失败时先把原因记进 findings：finally 里落盘的 JSON 要能说明今天为什么红。 */
        const fail = (message) => {
            findings.error = message;
            throw new Error(message);
        };

        findings.dom.bodyText = await evaluate('document.body.innerText.slice(0, 3000)');
        findings.dom.hasFailedBanner = await evaluate('/did not activate|Failed to load plugins/i.test(document.body.innerText)');
        findings.bootOk = findings.dom.hasFailedBanner === false;

        findings.dom.clickables = await evaluate(`Array.from(document.querySelectorAll('button,[role="button"],a')).map(e => ({
            aria: e.getAttribute('aria-label'), title: e.getAttribute('title'), text: (e.innerText||'').trim().slice(0,24)
        })).slice(0, 60)`);

        // 第六十二轮（新-22）：记忆面板从侧栏图标（root 作用域）迁到会话页签行
        // （conversation.view，**session 作用域**）—— 会话页签行只在**真的有会话**时
        // 渲染（hero / 草稿态没有那一行，宿主自己的「上下文」「编队」也一样；本轮活体
        // 实测：草稿态只有 composer，点了侧栏一条会话之后才出现
        // `div[data-conversation-tabs][role=tablist]` 与 五个 `button[role=tab]`）。
        // 所以先把会话开起来，并把这两步照实记进 findings —— 失败时要能分清
        // 「面板没迁对」和「压根没有会话」。
        //
        // 两个坑（都是活体实测出来的）：
        //   ① 「新建会话」按钮有**两个**：侧栏顶部那个品牌 logo 也带同一个 aria-label，
        //      点它没有任何反应 —— 要取带快捷键文本（「新会话 Ctrl+Alt+N」）的那一个；
        //   ② **空白的新会话也没有会话页签行**（hero / 草稿态只有 composer）：点完
        //      「新建会话」会多出一条 aria-selected=true 的「新会话」行，中心区仍是
        //      hero。这是宿主自己的口径，不是本轮引入的 —— 「上下文」「编队」在草稿态
        //      同样不存在。所以这里要挑一条**有内容**的会话行：会话行带 aria-selected，
        //      工作区分组行带 aria-expanded（用这个把两类行分开）。
        const READ_CONV_TABS = `Array.from(document.querySelectorAll('[data-conversation-tabs] [role="tab"]')).map(e => (e.innerText||'').trim())`;
        const sessionStep = { before: await evaluate(READ_CONV_TABS), attempts: [] };
        if (sessionStep.before.length === 0) {
            // 第六十三轮修正：**逐行试**，不要只点第一行。
            //
            // 第六十二轮写的是「点第一条 unselected 的会话行」，当时能过；第六十三轮
            // 复核时连续两次红：会话列表最上面那条**正是当前正在写日志的那一条**
            // （隔离实例与真 GUI 共用同一份会话史），点它之后中心区仍是 hero、页签行为空，
            // 于是后面的面板探针全部拿到 null —— 看起来像「记忆面板坏了」，其实是
            // 「没打开任何会话」。这里改成按顺序逐行试（每行给 7 秒），开出来就停；
            // 一条都开不出来才退到「新建会话」（那是已知的空会话形态）。
            // 每一条尝试都记进 findings，失败时能分清「面板没迁对」与「压根没开会话」。
            sessionStep.candidates = await evaluate(`(() => {
                const rows = Array.from(document.querySelectorAll('[role="treeitem"][aria-selected]'));
                const unselected = rows.filter(e => e.getAttribute('aria-selected') === 'false');
                return (unselected.length > 0 ? unselected : rows).map(e => (e.innerText || '').trim().slice(0, 30));
            })()`);
            // 第六十三轮第二个修正：侧栏的会话树**渲染得比 9 秒慢**（它要等一次会话快照）。
            // 实测三次连红时这里数出来是 0 条 —— 而同一台机器上晚一秒再看就有 14 个
            // treeitem。所以先轮询等它出来（最多 30 秒），再逐行试；不轮询的话，
            // 这一套会变成「跟机器负载比」，而不是「跟界面比」。
            let waitMs = 0;
            while (!Array.isArray(sessionStep.candidates) || sessionStep.candidates.length === 0) {
                if (waitMs >= 30000) break;
                await sleep(1500);
                waitMs += 1500;
                sessionStep.candidates = await evaluate(`(() => {
                    const rows = Array.from(document.querySelectorAll('[role="treeitem"][aria-selected]'));
                    const unselected = rows.filter(e => e.getAttribute('aria-selected') === 'false');
                    return (unselected.length > 0 ? unselected : rows).map(e => (e.innerText || '').trim().slice(0, 30));
                })()`);
            }
            sessionStep.waitedMs = waitMs;
            const tries = Array.isArray(sessionStep.candidates) ? Math.min(sessionStep.candidates.length, 6) : 0;
            for (let index = 0; index < tries; index += 1) {
                sessionStep.attempts.push({ how: `session-row#${index}`, result: await evaluate(`(() => {
                    const rows = Array.from(document.querySelectorAll('[role="treeitem"][aria-selected]'));
                    const unselected = rows.filter(e => e.getAttribute('aria-selected') === 'false');
                    const list = unselected.length > 0 ? unselected : rows;
                    const row = list[${index}] ?? null;
                    if (!row) return null;
                    row.click();
                    return { sessions: rows.length, text: (row.innerText || '').trim().slice(0, 30) };
                })()`) });
                await sleep(7000);
                if ((await evaluate(READ_CONV_TABS)).length > 0) break;
            }
            if ((await evaluate(READ_CONV_TABS)).length === 0) {
                sessionStep.attempts.push({ how: 'new-session-button', result: await evaluate(`(() => {
                    const named = Array.from(document.querySelectorAll('button'))
                        .filter(e => (e.getAttribute('aria-label') || '').trim() === '新建会话');
                    const hit = named.find(e => (e.innerText || '').includes('新会话')) ?? named[named.length - 1] ?? null;
                    if (!hit) return null;
                    hit.click();
                    return { candidates: named.length, cls: String(hit.className).slice(0, 60) };
                })()`) });
                await sleep(7000);
            }
        }
        sessionStep.after = await evaluate(READ_CONV_TABS);
        findings.dom.sessionStep = sessionStep;
        log(`   候选会话行：${JSON.stringify(sessionStep.candidates ?? null)}`);
        log(`   尝试：${sessionStep.attempts.map((item) => `${item.how}=${JSON.stringify(item.result)}`).join(' / ')}`);
        log(`   会话页签：${JSON.stringify(sessionStep.after)}`);

        // 这个隔离 server 与用户 GUI 共用 profile session store；真实 slash command 会拿
        // 会话写句柄。下面的范围回归会点真实 picker 并发 /kb-scope，但 commands/execute
        // 已在 CDP Fetch 层按上面的 pattern 全量截获：每个分支都 fulfill/failRequest，
        // 从不 continueRequest，命令不会到达宿主 commands 服务，也不会碰用户会话。
        findings.dom.kbScopePicker = await evaluate(`(() => {
            const button = document.querySelector('[data-kb-scope-toggle]');
            return { button: !!button, expanded: button?.getAttribute('aria-expanded') ?? null,
                panel: !!document.querySelector('[data-kb-scope-panel]') };
        })()`);
        const scopePicker = findings.dom.kbScopePicker;
        if (!scopePicker.button || scopePicker.expanded !== 'false' || scopePicker.panel) {
            fail(`聊天输入框左侧的范围入口未以收起状态挂载：${JSON.stringify(scopePicker)}`);
        }
        log('知识库会话范围：已挂载到真实宿主聊天框（不在共享会话上执行命令）');

        /*
         * 会话知识库范围的浏览器回归（第十六轮）：
         *
         * 前面的挂载检查只证明按钮在；这里把**真实 DOM 点击 -> remote.commands.execute
         * -> CDP Fetch 截获 -> 虚构 /kb-scope 回执 -> 真实 UI 回读**整条链路走一遍。
         * 所有命令都在网络层被 fulfill/failRequest，没有一条到达宿主 commands service；
         * 下面的 scopeFixture.calls 就是“实际发出过什么”的唯一证据。
         *
         * 点击用 CDP Input.dispatchMouseEvent，并在点之前用 elementFromPoint 确认
         * 坐标真的命中了目标（不是被面板/吸底输入框盖住），避免“坐标点空也算过”。
         */
        const clickScopeSelector = async (selector, label) => {
            const point = await evaluate('(() => {' +
                ' const el = document.querySelector(' + JSON.stringify(selector) + ');' +
                ' if (el === null) return { ok: false, reason: "not-found" };' +
                ' el.scrollIntoView({ block: "center", inline: "nearest" });' +
                ' const rect = el.getBoundingClientRect();' +
                ' const x = rect.left + rect.width / 2;' +
                ' const y = rect.top + rect.height / 2;' +
                ' const hit = document.elementFromPoint(x, y);' +
                ' return {' +
                '  ok: true, x: x, y: y, disabled: el.disabled === true,' +
                '  inViewport: rect.left >= 0 && rect.top >= 0 && rect.right <= innerWidth && rect.bottom <= innerHeight,' +
                '  hitTag: hit === null ? null : hit.tagName,' +
                '  hitText: hit === null ? "" : String(hit.innerText || "").trim().slice(0, 40),' +
                '  covered: hit === null || !(hit === el || el.contains(hit))' +
                ' };' +
                '})()');
            if (!point || point.ok !== true) fail('范围回归：找不到 ' + label + '（' + selector + '）');
            if (point.disabled) fail('范围回归：' + label + ' 处于禁用态，无法点击');
            if (point.covered) fail('范围回归：' + label + ' 被别的元素盖住（hit=' + JSON.stringify(point.hitTag) + ' / ' + JSON.stringify(point.hitText) + '）');
            if (!point.inViewport) fail('范围回归：' + label + ' 不在视口内（x=' + Math.round(point.x) + ', y=' + Math.round(point.y) + '）');
            await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: point.x, y: point.y, button: 'none' }, sessionId);
            await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'left', clickCount: 1, buttons: 1 }, sessionId);
            await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'left', clickCount: 1, buttons: 0 }, sessionId);
            await sleep(150);
            return point;
        };
        const waitScope = async (expression, label, timeoutMs = 10000) => {
            const deadline = Date.now() + timeoutMs;
            while (Date.now() < deadline) {
                if (await evaluate(expression)) return;
                await sleep(100);
            }
            const state = await evaluate('(() => {' +
                ' const t = document.querySelector("[data-kb-scope-toggle]");' +
                ' return {' +
                '  state: t === null ? null : t.getAttribute("data-kb-scope-state"),' +
                '  blocked: t === null ? null : t.getAttribute("data-kb-scope-blocked"),' +
                '  label: t === null ? null : (t.innerText || "").trim(),' +
                '  panel: document.querySelector("[data-kb-scope-panel]") !== null,' +
                '  loading: document.querySelector("[data-kb-scope-folders-loading]") !== null,' +
                '  error: document.querySelector("[data-kb-scope-error]") === null ? null : document.querySelector("[data-kb-scope-error]").innerText' +
                ' };' +
                '})()');
            fail('范围回归：等不到 ' + label + '；当前 ' + JSON.stringify(state));
        };

        findings.dom.kbScopeCommand = { initial: null, final: null, calls: [], replies: [] };
        await clickScopeSelector('[data-kb-scope-toggle]', '知识库范围入口');
        await waitScope('(() => {' +
            ' const t = document.querySelector("[data-kb-scope-toggle]");' +
            ' return t !== null && document.querySelector("[data-kb-scope-panel]") !== null' +
            '  && document.querySelector("[data-kb-scope-folders-loading]") === null' +
            '  && t.getAttribute("data-kb-scope-state") === "unset";' +
            '})()', '未设置状态与目录加载完成');
        const scopeInitial = await evaluate('(() => {' +
            ' const t = document.querySelector("[data-kb-scope-toggle]");' +
            ' const n = document.querySelector("[data-kb-scope-notice]");' +
            ' return {' +
            '  state: t.getAttribute("data-kb-scope-state"),' +
            '  blocked: t.getAttribute("data-kb-scope-blocked"),' +
            '  label: (t.innerText || "").trim(),' +
            '  notice: n === null ? "" : (n.innerText || "").trim()' +
            ' };' +
            '})()');
        if (scopeInitial.state !== 'unset' || scopeInitial.blocked !== 'true') {
            fail('范围回归：未设置状态没有按禁止检索展示：' + JSON.stringify(scopeInitial));
        }
        if (scopeInitial.label.indexOf('未选择') === -1 || scopeInitial.notice.indexOf('禁止检索') === -1) {
            fail('范围回归：未设置时入口/提示没有说清“禁止检索”：' + JSON.stringify(scopeInitial));
        }
        findings.dom.kbScopeCommand.initial = scopeInitial;
        // Switch only the isolated browser DOM's host theme attribute (not the saved profile).
        // Sample computed colors before sending the synthetic selection, restore immediately.
        const scopeContrast = await evaluate(`(() => {
            const body = document.body;
            const panel = document.querySelector('[data-kb-scope-panel]');
            const row = panel?.querySelector('[data-kb-scope-folder]');
            if (!row) return null;
            const attr = 'data-ds-dark-theme';
            const previous = body.getAttribute(attr);
            const channels = (color) => (color.match(/[0-9.]+/g) || []).map(Number);
            const lum = (values) => values.slice(0, 3)
                .map(v => v / 255).map(v => v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)
                .reduce((sum, v, i) => sum + v * [0.2126, 0.7152, 0.0722][i], 0);
            const sample = (dark) => {
                if (dark) body.setAttribute(attr, ''); else body.removeAttribute(attr);
                const p = getComputedStyle(panel);
                const foreground = getComputedStyle(row).color;
                const background = p.backgroundColor;
                const fg = channels(foreground);
                const bg = channels(background);
                const levels = [lum(fg), lum(bg)].sort((a, b) => a - b);
                return { foreground, background, opaque: bg.length === 3 || bg[3] === 1,
                    contrast: Number(((levels[1] + 0.05) / (levels[0] + 0.05)).toFixed(2)) };
            };
            try { return { light: sample(false), dark: sample(true) }; }
            finally { if (previous === null) body.removeAttribute(attr); else body.setAttribute(attr, previous); }
        })()`);
        findings.dom.kbScopeContrast = scopeContrast;
        if (!scopeContrast || !['light', 'dark'].every(key =>
            scopeContrast[key]?.opaque && scopeContrast[key].contrast >= 4.5)
            || scopeContrast.light.background === scopeContrast.dark.background) {
            fail('范围回归：亮/暗主题目录文字对比度不足：' + JSON.stringify(scopeContrast));
        }
        log('知识库范围主题对比（浏览器计算）：' + JSON.stringify(scopeContrast));

        await waitScope('(() => {' +
            " const r = document.querySelector(" + JSON.stringify('[data-kb-scope-folder="来源"]') + ");" +
            ' return r !== null && r.disabled !== true;' +
            '})()', '「来源」目录行可点');
        await clickScopeSelector('[data-kb-scope-folder="来源"]', '「来源」目录行');
        await waitScope('(() => {' +
            ' const t = document.querySelector("[data-kb-scope-toggle]");' +
            ' return t !== null && t.getAttribute("data-kb-scope-state") === "folder"' +
            '  && t.getAttribute("aria-expanded") === "false"' +
            '  && document.querySelector("[data-kb-scope-panel]") === null;' +
            '})()', '选择「来源」后的真实回读');
        const scopeFinal = await evaluate('(() => {' +
            ' const t = document.querySelector("[data-kb-scope-toggle]");' +
            ' return {' +
            '  state: t.getAttribute("data-kb-scope-state"),' +
            '  blocked: t.getAttribute("data-kb-scope-blocked"),' +
            '  label: (t.innerText || "").trim(),' +
            '  expanded: t.getAttribute("aria-expanded"),' +
            '  panel: document.querySelector("[data-kb-scope-panel]") !== null' +
            ' };' +
            '})()');
        if (scopeFinal.state !== 'folder' || scopeFinal.blocked !== 'false'
            || scopeFinal.label.indexOf('来源') === -1 || scopeFinal.expanded !== 'false' || scopeFinal.panel) {
            fail('范围回归：选择「来源」后入口没有回到已选状态：' + JSON.stringify(scopeFinal));
        }
        findings.dom.kbScopeCommand.final = scopeFinal;

        const scopeCalls = scopeFixture.calls.slice();
        findings.scopeCommands = scopeCalls;
        findings.scopeReplies = scopeFixture.replies.slice();
        findings.dom.kbScopeCommand.calls = scopeCalls;
        findings.dom.kbScopeCommand.replies = scopeFixture.replies.slice();
        findings.dom.kbScopeCommand.hits = findings.scopeCommandHits;
        if (findings.scopeCommandHits !== scopeCalls.length) {
            fail('范围回归：截获的 commands/execute 次数与夹具记录不一致（' + findings.scopeCommandHits + ' vs ' + scopeCalls.length + '）');
        }
        if (findings.scopeRefused.length > 0 || scopeFixture.refused.length > 0) {
            fail('范围回归：出现了未识别的命令载体/命令，已按 fail-closed 拒绝：' + JSON.stringify(findings.scopeRefused.concat(scopeFixture.refused)));
        }
        if (scopeCalls[0] !== '/kb-scope' || !scopeCalls.includes('/kb-scope set 来源')
            || scopeCalls[scopeCalls.length - 1] !== '/kb-scope') {
            fail('范围回归：picker 发出的命令序列不对（应为首读状态 -> set 来源 -> 回读状态）：' + JSON.stringify(scopeCalls));
        }
        if (scopeFixture.mode !== 'folder' || scopeFixture.folder !== '来源') {
            fail('范围回归：虚构范围没有落到「来源」：' + JSON.stringify({ mode: scopeFixture.mode, folder: scopeFixture.folder }));
        }
        log('知识库会话范围：真实 DOM 点击 + 虚构 /kb-scope 回执通过（' + JSON.stringify(scopeCalls) + '；未转发任何命令）');

        const clicked = await evaluate(`(() => {
            // 第六十二轮（新-22）起「记忆」是**会话页签行**里的一格（排在「上下文」
            // 「编队」右侧，conversation.view 的 order 40），不再是侧栏图标 ——
            // 所以先按宿主给会话页签行打的 data-conversation-tabs 定位，再退回
            // 「文本 / aria-label 恰好是记忆」的通用找法。不能只按 /记忆/ 模糊匹配：
            // 会话标题里也可能带「记忆」两个字（面板自己的六个子页签文本里也带）。
            const exactText = (e) => (e.innerText || '').trim() === '记忆';
            const conv = Array.from(document.querySelectorAll('[data-conversation-tabs] [role="tab"]'));
            const convHit = conv.find(exactText);
            const tabHit = Array.from(document.querySelectorAll('[role="tab"]')).find(exactText);
            const exact = Array.from(document.querySelectorAll('button,[role="button"],a,li,div'))
                .find(e => (e.getAttribute('aria-label') || '').trim() === '记忆');
            const byText = Array.from(document.querySelectorAll('button,[role="button"],a,li,div')).filter(exactText)
                .sort((a, b) => (a.innerText || '').length - (b.innerText || '').length)[0];
            const hit = convHit ?? tabHit ?? exact ?? byText;
            if (!hit) return null;
            hit.click();
            return {
                tag: hit.tagName.toLowerCase(),
                aria: hit.getAttribute('aria-label'),
                role: hit.getAttribute('role'),
                matched: convHit ? 'conversation-tab' : (tabHit ? 'tab' : (exact ? 'aria-label' : 'text')),
            };
        })()`);
        findings.dom.clickedMemory = clicked;
        await sleep(5000);

        const snapshotDom = async () => {
            findings.dom.tabs = await evaluate(`Array.from(document.querySelectorAll('[role="tab"]')).map(e => (e.innerText||'').trim())`);
            // 会话页签行（第六十二轮新-22 的落点）：这一排里必须能读到「记忆」，
            // 且在「上下文」「编队」之后。
            findings.dom.conversationTabs = await evaluate(READ_CONV_TABS);
            findings.dom.panelText = await evaluate(`(() => { const el = Array.from(document.querySelectorAll('*')).find(e => /记忆 ·/.test(e.innerText||'') && (e.innerText||'').length < 4000); return el ? el.innerText.slice(0, 2500) : null; })()`);
            // 第九十一轮信息架构：一级是记忆 / 知识库 / 图库；热记忆 / 台账是记忆内部的二级层级。
            findings.dom.ia = await evaluate(`(() => ({
                active: (() => { const el = document.querySelector('[data-tab][aria-selected="true"]'); return el ? el.getAttribute('data-tab') : null; })(),
                primary: Array.from(document.querySelectorAll('[data-tab]')).map((e) => e.getAttribute('data-tab')),
                layers: Array.from(document.querySelectorAll('[data-memory-layers] [data-layer]')).map((e) => e.getAttribute('data-layer')),
            }))()`);
            if (JSON.stringify(findings.dom.ia.primary) !== JSON.stringify(['memory', 'kb', 'gallery'])) {
                throw new Error('记忆面板信息架构不对：一级页签应为记忆 / 知识库 / 图库，实际 ' + JSON.stringify(findings.dom.ia.primary));
            }
            // 二级层级只在「记忆」一级下存在；切到知识库页再跑 snapshotDom 时读到空数组是对的。
            if (findings.dom.ia.active === 'memory'
                && JSON.stringify(findings.dom.ia.layers) !== JSON.stringify(['preference', 'project', 'ledger'])) {
                throw new Error('记忆面板信息架构不对：二级层级应为个人偏好 / 项目记忆 / 台账，实际 ' + JSON.stringify(findings.dom.ia.layers));
            }
            if (findings.dom.ia.active !== 'memory' && findings.dom.ia.layers.length > 0) {
                throw new Error('记忆面板信息架构不对：非记忆页不应显示二级层级，实际 ' + JSON.stringify(findings.dom.ia.layers));
            }
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
                return { width: Math.round(r.width), height: Math.round(r.height), cards: el.querySelectorAll('.om-card').length };
            })()`);
            // 容量口径：归档摘要要按**摘要文件数**报（1 / 60），不是条目数（30 / 60）。
            findings.dom.gauges = await evaluate(`(() => Array.from(document.querySelectorAll('[data-gauge]')).map((el) => ({
                key: el.getAttribute('data-gauge'), percent: el.getAttribute('data-percent'),
                text: (el.innerText||'').replace(/\\s+/g,' ').trim().slice(0, 60),
            })))()`);
            findings.dom.storeChips = await evaluate(`(() => { const el = document.querySelector('[data-store]'); return el ? (el.innerText||'').replace(/\\s+/g,' ').trim().slice(0, 120) : null; })()`);
            findings.dom.kbStoreCard = await evaluate(`(() => { const el = document.querySelector('[data-store="kb"]'); return el ? (el.innerText||'').replace(/\\s+/g,' ').trim().slice(0, 160) : null; })()`);
            if (findings.dom.kbStoreCard === null) throw new Error('统计区：找不到单库的 [data-store="kb"] 卡片');
        };

        await snapshotDom();
        await shot('memory-panel-hot.png');

        // 深色主题：宿主靠 body[data-ds-dark-theme] 切令牌，面板应当整体跟着变
        findings.dom.darkTokens = await evaluate(`(() => {
            document.body.setAttribute('data-ds-dark-theme', '');
            const cs = getComputedStyle(document.body);
            return { layer1: (cs.getPropertyValue('--dsw-alias-bg-layer-1')||'').trim(), label: (cs.getPropertyValue('--dsw-alias-label-primary')||'').trim() };
        })()`);
        await sleep(900);
        await shot('memory-panel-dark.png');

        // 宽屏：容量 + 存储域并排；窄屏塌成一列。
        await evaluate(`document.body.removeAttribute('data-ds-dark-theme')`);
        await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1680, height: 950, deviceScaleFactor: 1, mobile: false }, sessionId);
        await sleep(900);
        findings.dom.wideLayout = await evaluate(`(() => {
            const grid = document.querySelector('.om-grid-dash');
            return { panelColumns: grid ? getComputedStyle(grid).gridTemplateColumns : null, tileColumns: (() => { const t = document.querySelector('[data-metrics]'); return t ? getComputedStyle(t).gridTemplateColumns : null; })() };
        })()`);
        await shot('memory-panel-wide.png');
        await cdp.send('Emulation.clearDeviceMetricsOverride', {}, sessionId);
        await sleep(400);
        findings.dom.narrowLayout = await evaluate(`(() => {
            const root = document.querySelector('.om-mem');
            const styleEl = Array.from(document.querySelectorAll('style')).find(s => (s.textContent||'').includes('.om-card'));
            return {
                containerType: root ? getComputedStyle(root).containerType : null,
                rootWidth: root ? Math.round(root.getBoundingClientRect().width) : null,
                supportsContainer: CSS.supports('container-type', 'inline-size'),
                styleText: styleEl ? styleEl.textContent.length : 0,
            };
        })()`);

        // ── 第十三轮：写入指令 / 分页 / 完整内容 / 减少动效 ──
        const rectOf = (selector) => `(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return null; el.scrollIntoView({block:'center'}); const r = el.getBoundingClientRect(); return { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }; })()`;

        const toLedgerTab = async () => {
            // 台账是「记忆」视图下的二级层级。
            const clicked = await evaluate(`(() => {
                const layer = document.querySelector('[data-memory-layers] [data-layer="ledger"]');
                if (layer) { layer.click(); return 'layer'; }
                const memory = document.querySelector('[data-tab="memory"]');
                if (memory) { memory.click(); return 'memory'; }
                return 'missing';
            })()`);
            if (clicked === 'memory') {
                await sleep(600);
                await evaluate(`(() => { const t = document.querySelector('[data-memory-layers] [data-layer="ledger"]'); if (t) t.click(); return !!t; })()`);
            }
            await sleep(1500);
        };

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
            // 谁是**真正的**滚动容器：面板自己不再是（第六十二轮修正），要找上游第一个
            // overflow-y 可滚、且真的溢出了的祖先 —— 实测是宿主的 .…_scrollBody。
            let scroller = null;
            for (let n = panel.parentElement; n && n !== document.body; n = n.parentElement) {
                const oy = getComputedStyle(n).overflowY;
                if ((oy === 'auto' || oy === 'scroll' || oy === 'overlay') && n.scrollHeight > n.clientHeight) { scroller = n; break; }
            }
            // 量之前先关掉平滑滚动：赋值 scrollTop 之后立刻读会读到旧值（动画还没跑完）
            // —— 量的是「能不能滚到底」，不是「动画多快」，所以临时切成 auto。
            const previous = scroller ? scroller.style.scrollBehavior : null;
            if (scroller) scroller.style.scrollBehavior = 'auto';
            const before = scroller ? scroller.scrollTop : null;
            if (scroller) scroller.scrollTop = scroller.scrollHeight;
            const after = scroller ? scroller.scrollTop : null;
            const last = panel.lastElementChild;
            const panelRect = panel.getBoundingClientRect();
            const lastRect = last ? last.getBoundingClientRect() : null;
            const visible = lastRect ? (lastRect.bottom <= panelRect.bottom + 2) : null;
            // 头部吸顶：滚到底之后它的顶边仍然贴着**滚动容器**的顶边（本层不再是滚动容器）。
            const head = panel.querySelector('[data-memory-nav="sticky"]');
            const scrollerTop = scroller ? scroller.getBoundingClientRect().top : null;
            const headDelta = head && scrollerTop !== null ? Math.round(head.getBoundingClientRect().top - scrollerTop) : null;
            if (scroller) {
                // 还原留到截图之后：滚到底的状态就是仪表盘证据（截图在 evaluate 外做）。
                window.__omScrollReset = () => {
                    scroller.scrollTop = before;
                    scroller.style.scrollBehavior = previous;
                };
            }
            return {
                panelOverflowY: getComputedStyle(panel).overflowY,
                panelOverflowX: getComputedStyle(panel).overflowX,
                panelClientHeight: panel.clientHeight,
                panelScrollHeight: panel.scrollHeight,
                panelIsOwnScroller: panel.scrollHeight > panel.clientHeight && /auto|scroll/.test(getComputedStyle(panel).overflowY),
                scrollerClass: scroller ? String(scroller.className).slice(0, 60) : null,
                scrollerClientHeight: scroller ? scroller.clientHeight : null,
                scrollerScrollHeight: scroller ? scroller.scrollHeight : null,
                scrollerBefore: before,
                scrollerAfter: after,
                scrolledToBottom: before !== null && after > before,
                tailVisibleAfterScroll: visible,
                headStickDelta: headDelta,
                headStuckToScroller: headDelta === null ? null : Math.abs(headDelta) <= 12,
                listMaxHeight: list ? getComputedStyle(list).maxHeight : null,
                listOverflowY: list ? getComputedStyle(list).overflowY : null,
            };
        })()`);
        log(`滚动契约：${JSON.stringify(findings.dom.scrollContract)}`);
        // 仪表盘证据：滚到统计区顶部（让开吸顶头），截一张「指标网格」再还原滚动位置。
        findings.dom.dashboardCapture = await evaluate(`(() => {
            const panel = document.querySelector('[data-memory-panel="scroll"]');
            const dash = panel ? panel.querySelector('[data-memory-dashboard="true"]') : null;
            if (!panel || !dash) return null;
            let scroller = null;
            for (let n = panel.parentElement; n && n !== document.body; n = n.parentElement) {
                const oy = getComputedStyle(n).overflowY;
                if ((oy === 'auto' || oy === 'scroll' || oy === 'overlay') && n.scrollHeight > n.clientHeight) { scroller = n; break; }
            }
            if (!scroller) return null;
            const head = panel.querySelector('[data-memory-nav="sticky"]');
            const headHeight = head ? head.getBoundingClientRect().height : 0;
            const dashTop = dash.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop;
            scroller.scrollTop = Math.max(0, dashTop - headHeight - 8);
            return { headHeight: Math.round(headHeight), dashTop: Math.round(dashTop), scrollTop: Math.round(scroller.scrollTop) };
        })()`);
        if (!findings.dom.dashboardCapture) throw new Error('仪表盘：找不到 [data-memory-dashboard] 或宿主滚动容器');
        log(`仪表盘定位：${JSON.stringify(findings.dom.dashboardCapture)}`);
        await new Promise((resolve) => setTimeout(resolve, 300));
        await shot('memory-panel-dashboard.png');
        await evaluate(`(() => { if (typeof window.__omScrollReset === 'function') { window.__omScrollReset(); delete window.__omScrollReset; } return true; })()`);
        // ── 硬判据（第六十二轮补）──
        // 这一条是**产品行为**，不是观测值：记忆面板滚不动就是坏的。
        // 第六十二轮初版就是这样坏的（面板高＝内容高、滚轮链又被
        // `overscroll-behavior:contain` 断在这一层），而当时这一整套探针
        // **只打印不断言**，所以 42/42 全绿也没人发现 —— 现在直接判死。
        if (!findings.dom.scrollContract) throw new Error('滚动契约：找不到面板 [data-memory-panel="scroll"]');
        if (findings.dom.scrollContract.panelIsOwnScroller) {
            throw new Error('滚动契约：面板又变回自己的滚动容器了（第六十二轮定的口径是交给宿主的会话滚动容器）');
        }
        if (findings.dom.scrollContract.scrollerClass === null) {
            throw new Error('滚动契约：面板上游找不到真正能滚的容器');
        }
        if (findings.dom.scrollContract.scrolledToBottom !== true) {
            throw new Error(`滚动契约：面板滚不动（${findings.dom.scrollContract.scrollerClass} 的 scrollTop 没有变化）`);
        }
        // ── 硬判据（第六十四轮补）：默认根 = 当前会话的工作目录 ──
        // 用户报的问题就是「面板默认打开了别的工作区」。所以**第一次**请求就必须带 cwd
        // （宿主的会话上下文给了当前会话的目录）；只看最后一次请求会把「先开错再换对」
        // 读成正常 —— 那正是这个 bug 的形态。
        const snapshotUrls = Array.isArray(findings.snapshotUrls) ? findings.snapshotUrls : [];
        log(`记忆端点请求：${snapshotUrls.length} 次，第一次 ${snapshotUrls[0] ?? '-'}`);
        if (snapshotUrls.length === 0) throw new Error('默认根：这一次运行没有观察到任何 /office-memory/snapshot 请求');
        if (!/[?&]cwd=/.test(snapshotUrls[0])) {
            throw new Error(`默认根：第一次请求没带 cwd（会开成宿主挑的那个工作区）：${snapshotUrls[0]}`);
        }
        await shot('memory-panel-scroll.png');
        // 上一段点过「忘记」，会留一张待发送指令卡；后面的来源 / 知识库截图先收掉它，
        // 免得证据图上永远压着一张上一段操作留下的卡片。
        await evaluate(`(() => { const b = document.querySelector('[data-memory-dismiss]'); if (b) { b.click(); return true; } return false; })()`);
        await sleep(500);

        // ── 第八十九轮：个人偏好 / 项目记忆 / 台账三档 ──
        // 单测能证明「渲染出了按钮」，只有真浏览器能证明点了之后请求、样式与
        // 布局都还在。这里顺带留下 UI 证据图。
        await evaluate(`(() => { const t = document.querySelector('[data-tab="memory"]'); if (t) t.click(); return !!t; })()`);
        await sleep(1200);
        findings.dom.layers = await evaluate(`(() => {
            const row = document.querySelector('[data-memory-layers="true"]');
            if (!row) return null;
            return {
                visible: !!(row.offsetWidth || row.offsetHeight),
                buttons: Array.from(row.querySelectorAll('[data-layer]')).map((b) => ({
                    layer: b.getAttribute('data-layer'),
                    text: (b.innerText || '').replace(/\s+/g, ' ').trim(),
                    selected: b.getAttribute('aria-selected') === 'true',
                })),
                origin: !!document.querySelector('[data-memory-origin]'),
            };
        })()`);
        log(`记忆三档：${JSON.stringify(findings.dom.layers)}`);
        await shot('memory-panel-layers.png');
        if (!findings.dom.layers || findings.dom.layers.buttons.length !== 3) {
            throw new Error('记忆三档：找不到「个人偏好 / 项目记忆 / 台账」三个层级');
        }
        if (findings.dom.layers.origin) {
            throw new Error('记忆三档：来源筛选应已删除');
        }
        await evaluate(`(() => { const b = document.querySelector('[data-layer="project"]'); if (b) { b.click(); return true; } return false; })()`);
        await sleep(1500);
        // 第八十九轮：快速记一条默认收起成一行，点开才是输入卡。
        findings.dom.quickCollapsed = await evaluate(`(() => {
            const open = document.querySelector('[data-memory-quick-open]');
            const composer = document.querySelector('[data-memory-composer]');
            return { hasOpen: !!open, composer: !!composer };
        })()`);
        await evaluate(`(() => { const b = document.querySelector('[data-memory-quick-open]'); if (b) { b.click(); return true; } return false; })()`);
        await sleep(700);
        findings.dom.layerProject = await evaluate(`(() => ({
            selected: (() => { const el = document.querySelector('[data-layer][aria-selected="true"]'); return el ? el.getAttribute('data-layer') : null; })(),
            cards: Array.from(document.querySelectorAll('[data-item-key]')).map((e) => (e.innerText||'').replace(/\s+/g,' ').trim().slice(0, 90)).slice(0, 6),
            composer: !!document.querySelector('[data-memory-composer]'),
        }))()`);
        log(`项目记忆档：${JSON.stringify(findings.dom.layerProject)}`);
        log(`快速记一条（收起态）：${JSON.stringify(findings.dom.quickCollapsed)}`);
        await shot('memory-panel-project.png');
        if (!findings.dom.quickCollapsed || !findings.dom.quickCollapsed.hasOpen || findings.dom.quickCollapsed.composer) {
            throw new Error('记忆三档：快速记一条应默认收起成一行，点开才出输入卡');
        }
        if (findings.dom.layerProject.selected !== 'project') {
            throw new Error('记忆三档：点「项目记忆」没有切换过去');
        }
        if (!findings.dom.layerProject.composer) {
            throw new Error('记忆三档：项目记忆档没有直接写入的输入卡');
        }

        // 第八十七轮：「+ 记一条」直接打开草稿，不再生成待发送指令。
        // 这里只开不存：写一条真记忆会污染工作区，保存链路由 test/client.mjs 守。
        // 第99轮加探针：「输入卡不见」先分清是「根本没点开」「被隐藏的第二份面板」还是
        // 「宿主把整块重挂载、面板 state 回到初值」。只读 DOM，不写任何记忆。
        findings.dom.composeProbe = await evaluate(`(() => {
            const visible = (el) => !!(el && (el.offsetWidth || el.offsetHeight || (el.getClientRects && el.getClientRects().length > 0)));
            const panels = Array.from(document.querySelectorAll('[data-memory-panel="scroll"]'));
            const shownPanels = panels.filter(visible);
            const root = shownPanels.length > 0 ? shownPanels[shownPanels.length - 1] : (panels[0] || null);
            window.__omProbe = root;
            // 只点**可见且已连接**的那一份：宿主重挂载后旧节点会离开文档，隐藏副本不该被点。
            const candidates = Array.from(document.querySelectorAll('[data-memory-compose]'));
            const b = candidates.find(visible) || null;
            const clickedVisible = !!b && b.isConnected;
            if (b) b.click();
            return {
                before: {
                    panels: panels.length, shownPanels: shownPanels.length,
                    composers: document.querySelectorAll('[data-memory-composer]').length,
                    composeButtons: document.querySelectorAll('[data-memory-compose]').length,
                    quickOpenButtons: document.querySelectorAll('[data-memory-quick-open]').length,
                    selectedLayer: root ? (() => { const el = root.querySelector('[data-layer][aria-selected="true"]'); return el ? el.getAttribute('data-layer') : null; })() : null,
                },
                clickedVisible: clickedVisible,
                text: b ? (b.innerText || '').replace(/\\s+/g, ' ').trim() : null,
            };
        })()`);
        findings.dom.newMemory = findings.dom.composeProbe ? findings.dom.composeProbe.text : null;
        await sleep(900);
        findings.dom.composeProbeAfter = await evaluate(`(() => {
            const visible = (el) => !!(el && (el.offsetWidth || el.offsetHeight || (el.getClientRects && el.getClientRects().length > 0)));
            const panels = Array.from(document.querySelectorAll('[data-memory-panel="scroll"]'));
            const shownPanels = panels.filter(visible);
            const cur = shownPanels.length > 0 ? shownPanels[shownPanels.length - 1] : (panels[0] || null);
            const node = window.__omProbe || null;
            return {
                panels: panels.length, shownPanels: shownPanels.length,
                composers: document.querySelectorAll('[data-memory-composer]').length,
                composeButtons: document.querySelectorAll('[data-memory-compose]').length,
                quickOpenButtons: document.querySelectorAll('[data-memory-quick-open]').length,
                selectedLayer: cur ? (() => { const el = cur.querySelector('[data-layer][aria-selected="true"]'); return el ? el.getAttribute('data-layer') : null; })() : null,
                sameNode: !!(node && node === cur),
                oldConnected: node ? node.isConnected : null,
                newConnected: cur ? cur.isConnected : null,
            };
        })()`);
        // 宿主重挂载的可靠信号：点前记下的面板根已离开文档、当前可见根换成了新节点。
        // 只有这个信号成立才做一次有界重试；真回归时第二次点依然打不开，下面的断言照旧失败。
        const composeRemounted = findings.dom.composeProbe?.clickedVisible === true
            && findings.dom.composeProbeAfter.sameNode === false
            && findings.dom.composeProbeAfter.oldConnected === false
            && findings.dom.composeProbeAfter.newConnected === true;
        if (!findings.dom.composeProbeAfter.composers && composeRemounted) {
            log('compose：点开之后宿主把记忆面板整块重挂载了（旧根离开文档、换成新根），按一次有界重试');
            await evaluate(`(() => {
                const ok = (el) => !!(el && el.isConnected && ((el.offsetWidth || el.offsetHeight) > 0));
                const b = Array.from(document.querySelectorAll('[data-memory-compose]')).find(ok);
                if (!b) return false;
                b.click();
                return true;
            })()`);
            await sleep(900);
            findings.dom.composeProbeAfter = await evaluate(`(() => {
                const visible = (el) => !!(el && el.isConnected && (el.offsetWidth || el.offsetHeight || (el.getClientRects && el.getClientRects().length > 0)));
                const panels = Array.from(document.querySelectorAll('[data-memory-panel="scroll"]'));
                const shownPanels = panels.filter(visible);
                const cur = shownPanels.length > 0 ? shownPanels[shownPanels.length - 1] : (panels[0] || null);
                return {
                    retried: true,
                    panels: panels.length, shownPanels: shownPanels.length,
                    composers: document.querySelectorAll('[data-memory-composer]').length,
                    selectedLayer: cur ? (() => { const el = cur.querySelector('[data-layer][aria-selected="true"]'); return el ? el.getAttribute('data-layer') : null; })() : null,
                };
            })()`);
        }
        log('compose 探针（点前/点后）：' + JSON.stringify({ probe: findings.dom.composeProbe, after: findings.dom.composeProbeAfter, remounted: composeRemounted }));
        findings.dom.newMemoryEditor = await evaluate(`(() => {
            // 断言也只看可见的那一份：隐藏副本或旧节点里的输入卡不算数。
            const visibleRoot = (() => {
                const ok = (el) => !!(el && el.isConnected && (el.offsetWidth || el.offsetHeight || (el.getClientRects && el.getClientRects().length > 0)));
                const panels = Array.from(document.querySelectorAll('[data-memory-panel="scroll"]')).filter(ok);
                return panels.length > 0 ? panels[panels.length - 1] : document;
            })();
            const composer = visibleRoot.querySelector('[data-memory-composer]');
            const input = visibleRoot.querySelector('[data-memory-quick-text]');
            return {
                open: !!composer && !!input,
                title: composer ? (composer.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 80) : null,
                hasText: !!input,
                focused: input !== null && document.activeElement === input,
                duplicate: !!document.querySelector('[data-memory-editor="new"]'),
                pending: !!document.querySelector('[data-memory-pending]'),
            };
        })()`);
        log(`新增记忆入口：${JSON.stringify(findings.dom.newMemoryEditor)}`);
        await shot('memory-panel-new-memory.png');
        if (!findings.dom.newMemoryEditor || !findings.dom.newMemoryEditor.open) {
            throw new Error('新增记忆：页内「直接记一条」输入卡不见了');
        }
        if (findings.dom.newMemoryEditor.duplicate) {
            throw new Error('新增记忆：又冒出了第二张草稿卡，同一屏只能有一个新增入口');
        }
        if (!findings.dom.newMemoryEditor.focused) {
            throw new Error('新增记忆：点「+ 记一条」没有把焦点送进输入框');
        }
        if (findings.dom.newMemoryEditor.pending) {
            throw new Error('新增记忆：仍在生成待发送指令卡');
        }
        // 知识库页签：文档清单 / 块检索两种看法与空态都要在真实页面上成立。
        await evaluate(`(() => { const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => /知识库/.test(e.innerText||'')); if (t) { t.click(); return true; } return false; })()`);
        await sleep(1500);
        findings.dom.kbPanel = await evaluate(`(() => {
            const panel = document.querySelector('[data-kb-panel="true"]');
            if (!panel) return null;
            return {
                text: (panel.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 320),
                docs: panel.querySelectorAll('[data-kb-open]').length,
                empty: !!panel.querySelector('[data-kb-empty]'),
                // 单库之后知识库页不该再出现来源三档。
                originRow: !!document.querySelector('[data-memory-origin]'),
                modes: Array.from(panel.querySelectorAll('[data-kb-mode]')).map((b) => b.getAttribute('data-kb-mode') + ':' + (b.getAttribute('aria-pressed') === 'true')),
            };
        })()`);
        log(`知识库面板：${JSON.stringify(findings.dom.kbPanel)}`);
        await shot('memory-panel-kb.png');
        if (!findings.dom.kbPanel) throw new Error('知识库页签：找不到 [data-kb-panel]');
        if (findings.dom.kbPanel.originRow) throw new Error('知识库已是单库，页签上不该再有「全部 / 全局 / 工作区」来源筛选');
        // 第八十四轮：知识库是文件夹树，点一层要真的换范围（请求带 kbfolder），
        // 而不是只在本地把树高亮。
        findings.dom.kbTree = await evaluate(`(() => ({
            folders: Array.from(document.querySelectorAll('[data-kb-folder]')).map((e) => ({
                path: e.getAttribute('data-kb-folder'),
                text: (e.innerText || '').replace(/\s+/g, ' ').trim(),
            })),
            crumbs: Array.from(document.querySelectorAll('[data-kb-crumb]')).map((e) => e.getAttribute('data-kb-crumb')),
        }))()`);
        log(`知识库文件夹树：${JSON.stringify(findings.dom.kbTree)}`);
        if (!findings.dom.kbTree || !findings.dom.kbTree.folders.some((row) => row.path === '参考')) {
            throw new Error('知识库文件夹树：找不到「参考」这一层');
        }
        await evaluate(`(() => { const b = document.querySelector('[data-kb-folder="参考"]'); if (b) { b.click(); return true; } return false; })()`);
        await sleep(1500);
        findings.dom.kbFolder = await evaluate(`(() => ({
            url: location.search,
            crumbs: Array.from(document.querySelectorAll('[data-kb-crumb]')).map((e) => e.getAttribute('data-kb-crumb')),
            docs: Array.from(document.querySelectorAll('[data-kb-open]')).length,
            title: (document.querySelector('[data-kb-panel]') ? document.querySelector('[data-kb-panel]').innerText : '').replace(/\s+/g, ' ').trim().slice(0, 200),
        }))()`);
        findings.dom.kbFolder.request = findings.snapshotUrl ?? '';
        log(`知识库文件夹范围：${JSON.stringify(findings.dom.kbFolder)}`);
        if (!findings.dom.kbFolder.request.includes('kbfolder=')) throw new Error('知识库文件夹树：点了「参考」后快照请求没带 kbfolder');
        if (!findings.dom.kbFolder.crumbs.includes('参考')) throw new Error('知识库文件夹树：点了「参考」面包屑没跟上');
        await shot('memory-panel-kb-folder.png');
        const readKbPage = () => evaluate(`(() => {
            const root = document.querySelector('[data-kb-panel]');
            if (!root) return null;
            return { docs: Array.from(root.querySelectorAll('[data-kb-open]'), e => e.getAttribute('data-kb-open')),
                next: root.querySelector('[data-kb-page-next]')?.disabled,
                prev: root.querySelector('[data-kb-page-prev]')?.disabled,
                add: root.querySelectorAll('[data-kb-add-files]').length };
        })()`);
        const awaitKbPage = async (count) => {
            let page;
            for (let attempt = 0; attempt < 35; attempt += 1) {
                page = await readKbPage();
                if (page?.docs.length === count) return page;
                await sleep(100);
            }
            throw new Error(`知识库翻页：预期 ${count} 篇，实际 ${page?.docs.length}`);
        };
        const firstPage = await awaitKbPage(20);
        if (!firstPage.prev || firstPage.next || firstPage.add !== 1) throw new Error('知识库首页分页按钮或统一添加入口状态错误');
        await evaluate(`document.querySelector('[data-kb-page-next]').click()`);
        const secondPage = await awaitKbPage(3);
        const requestPage = new URL(findings.snapshotUrl).searchParams;
        if (requestPage.get('kboffset') !== '20' || requestPage.get('kbfolder') !== '参考'
            || secondPage.prev || !secondPage.next || secondPage.docs.some(id => firstPage.docs.includes(id))) {
            throw new Error('知识库第二页未重新取数、丢失当前文件夹或与首页重复');
        }
        await evaluate(`document.querySelector('[data-kb-page-prev]').click()`);
        const restoredPage = await awaitKbPage(20);
        if (restoredPage.docs.join(',') !== firstPage.docs.join(',')) throw new Error('知识库返回首页未恢复原顺序');
        findings.dom.kbPagination = { pageSizes: [firstPage.docs.length, secondPage.docs.length], folder: requestPage.get('kbfolder') };
        log(`知识库跨页：${JSON.stringify(findings.dom.kbPagination)}`);
        await evaluate(`(() => { const b = document.querySelector('[data-kb-mode="search"]'); if (b) { b.click(); return true; } return false; })()`);
        await sleep(1200);
        findings.dom.kbSearch = await evaluate(`(() => {
            const input = document.querySelector('[data-kb-search-input]');
            return input ? { placeholder: input.getAttribute('placeholder'), width: Math.round(input.getBoundingClientRect().width) } : null;
        })()`);
        log(`知识库块检索：${JSON.stringify(findings.dom.kbSearch)}`);
        await shot('memory-panel-kb-search.png');
        if (!findings.dom.kbSearch) throw new Error('知识库块检索：找不到检索输入框');
        // 真的输一个词并回车：只有活体才能证明受控输入 → 请求参数 → 命中渲染这条链没断。
        await evaluate(`(() => {
            const input = document.querySelector('[data-kb-search-input]');
            if (!input) return false;
            const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            setter.call(input, '复核');
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
            return true;
        })()`);
        await sleep(1500);
        findings.dom.kbHits = await evaluate(`(() => ({
            hits: Array.from(document.querySelectorAll('[data-kb-hit]')).map((e) => (e.innerText||'').replace(/\s+/g,' ').trim().slice(0, 100)),
            text: (document.querySelector('[data-kb-search]') ? document.querySelector('[data-kb-search]').innerText : '').replace(/\s+/g,' ').trim().slice(0, 240),
        }))()`);
        log(`知识库检索命中：${JSON.stringify(findings.dom.kbHits)}`);
        await shot('memory-panel-kb-hits.png');
        if (!Array.isArray(findings.dom.kbHits.hits) || findings.dom.kbHits.hits.length === 0) {
            throw new Error('知识库检索：输入「复核」回车后没有任何 [data-kb-hit]');
        }
        await evaluate(`(() => { const hit = document.querySelector('[data-kb-hit]'); if (hit) { hit.click(); return true; } return false; })()`);
        await sleep(1500);
        findings.dom.kbChunk = await evaluate(`(() => {
            const el = document.querySelector('[data-kb-chunk-text]');
            return el ? (el.innerText||'').replace(/\s+/g,' ').trim().slice(0, 160) : null;
        })()`);
        log(`知识块正文：${JSON.stringify(findings.dom.kbChunk)}`);
        await shot('memory-panel-kb-chunk.png');
        if (findings.dom.kbChunk === null) throw new Error('知识库检索：点命中后读不到块正文 [data-kb-chunk-text]');
        // ── 图库页签：真实点击 → 真实请求（假端点）→ 内容 → 切回知识库 ──
        // 这里只读假清单，不点上传、不填标签，也不发任何 POST。
        const galleryRequestsBefore = findings.galleryRequests.length;
        findings.dom.gallery = { requestsBefore: galleryRequestsBefore };
        findings.dom.gallery.clicked = await evaluate(`(() => {
            const t = document.querySelector('[data-tab="gallery"]');
            if (!t) return false;
            t.click();
            return true;
        })()`);
        await sleep(1800);
        findings.dom.gallery.panel = await evaluate(`(() => {
            const panel = document.querySelector('[data-gallery-panel="true"]');
            if (!panel) return null;
            return {
                active: (() => { const el = document.querySelector('[data-tab][aria-selected="true"]'); return el ? el.getAttribute('data-tab') : null; })(),
                items: panel.querySelectorAll('[data-gallery-item]').length,
                status: (() => { const el = panel.querySelector('[data-gallery-status]'); return el ? (el.innerText || '').replace(/\\s+/g, ' ').trim() : null; })(),
                capabilities: panel.querySelectorAll('[data-gallery-capability]').length,
                uploadInput: !!panel.querySelector('[data-gallery-upload-input]'),
                error: (() => { const el = panel.querySelector('[data-gallery-error]'); return el ? (el.innerText || '').replace(/\\s+/g, ' ').trim() : null; })(),
                empty: !!panel.querySelector('[data-gallery-empty]'),
            };
        })()`);
        findings.dom.gallery.requests = findings.galleryRequests.slice(galleryRequestsBefore);
        log(`图库页签：${JSON.stringify(findings.dom.gallery.panel)}；请求 ${JSON.stringify(findings.dom.gallery.requests)}`);
        if (findings.dom.gallery.clicked !== true) throw new Error('图库页签：找不到 [data-tab="gallery"]');
        if (findings.dom.gallery.panel === null) throw new Error('图库页签：点了之后没有 [data-gallery-panel]');
        if (findings.dom.gallery.panel.active !== 'gallery') throw new Error('图库页签：点开后 aria-selected 不在 gallery');
        if (findings.dom.gallery.panel.error !== null) throw new Error('图库页签：渲染报了错：' + findings.dom.gallery.panel.error);
        if (findings.dom.gallery.panel.items !== 2) throw new Error('图库页签：假清单有 2 张，页面渲染了 ' + findings.dom.gallery.panel.items + ' 张');
        if (!findings.dom.gallery.panel.status || !findings.dom.gallery.panel.status.includes('2 / 2')) {
            throw new Error('图库页签：状态栏没读到 2 / 2 张：' + String(findings.dom.gallery.panel.status));
        }
        if (findings.dom.gallery.panel.capabilities !== 5) throw new Error('图库页签：能力区应有 5 项，实际 ' + findings.dom.gallery.panel.capabilities);
        if (findings.dom.gallery.panel.uploadInput !== true) throw new Error('图库页签：上传入口不见了');
        if (!findings.dom.gallery.requests.some((r) => r.view === 'status')) throw new Error('图库页签：没有发出 view=status 请求');
        if (!findings.dom.gallery.requests.some((r) => r.view === 'list')) throw new Error('图库页签：没有发出 view=list 请求');
        if (findings.dom.gallery.requests.some((r) => r.method !== 'GET')) {
            throw new Error('图库页签：回归只该读，不应发 POST：' + JSON.stringify(findings.dom.gallery.requests));
        }
        findings.dom.gallery.listClicked = await evaluate(`(() => { const b = document.querySelector('[data-gallery-view="list"]'); if (!b) return false; b.click(); return true; })()`);
        await sleep(500);
        findings.dom.gallery.list = await evaluate(`(() => ({
            list: !!document.querySelector('[data-gallery-list]'),
            items: document.querySelectorAll('[data-gallery-item]').length,
        }))()`);
        log(`图库列表视图：${JSON.stringify(findings.dom.gallery.list)}`);
        if (findings.dom.gallery.listClicked !== true || !findings.dom.gallery.list.list || findings.dom.gallery.list.items !== 2) {
            throw new Error('图库页签：切列表视图后没有 2 行：' + JSON.stringify(findings.dom.gallery.list));
        }
        await evaluate(`(() => { const t = document.querySelector('[data-tab="kb"]'); if (t) t.click(); return !!t; })()`);
        await sleep(1200);
        findings.dom.gallery.back = await evaluate(`(() => ({
            active: (() => { const el = document.querySelector('[data-tab][aria-selected="true"]'); return el ? el.getAttribute('data-tab') : null; })(),
            kbPanel: !!document.querySelector('[data-kb-panel="true"]'),
            galleryPanel: !!document.querySelector('[data-gallery-panel="true"]'),
        }))()`);
        log(`切回知识库：${JSON.stringify(findings.dom.gallery.back)}`);
        if (findings.dom.gallery.back.active !== 'kb' || !findings.dom.gallery.back.kbPanel || findings.dom.gallery.back.galleryPanel) {
            throw new Error('图库页签：切回知识库失败：' + JSON.stringify(findings.dom.gallery.back));
        }
        // 预览面板：1258 宽度下页签行放不下第八格，先临时扩到 1680 再点；测完还原。
        // 下拉必须是实底（深色主题下原先是半透明灰玻璃）。
        await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1680, height: 950, deviceScaleFactor: 1, mobile: false }, sessionId);
        await sleep(900);
        findings.dom.previewTabClicked = await evaluate(`(() => {
            const t = Array.from(document.querySelectorAll('[role="tab"]')).find(e => (e.innerText||'').trim().startsWith('预览'));
            if (!t) return false;
            t.click();
            return true;
        })()`);
        await sleep(3200);
        findings.dom.previewPanel = await evaluate(`document.querySelector('[data-preview-head]') !== null`);
        findings.dom.previewSelect = await evaluate(`(() => {
            const s = document.querySelector('[data-preview-file]');
            if (!s) return null;
            const cs = getComputedStyle(s);
            return { background: cs.backgroundColor, color: cs.color, colorScheme: cs.colorScheme, options: s.options.length };
        })()`);
        log(`预览页签点击：${findings.dom.previewTabClicked}；面板：${findings.dom.previewPanel}；下拉配色：${JSON.stringify(findings.dom.previewSelect)}`);
        await shot('preview-panel.png');
        if (findings.dom.previewTabClicked !== true) throw new Error('预览面板：1680 宽度下页签行仍找不到「预览」');
        if (findings.dom.previewPanel !== true) throw new Error('预览面板：页签点了但没有 [data-preview-head]');
        if (findings.dom.previewSelect !== null) {
            const bg = findings.dom.previewSelect.background;
            const alpha = bg.startsWith('rgba(') ? Number(bg.split(',')[3].replace(')', '')) : 1;
            if (!(alpha >= 0.99)) throw new Error('预览下拉底色还是半透明：' + bg);
        } else {
            log('   预览面板没有产物，跳过下拉配色断言');
        }
        await cdp.send('Emulation.clearDeviceMetricsOverride', {}, sessionId);

        // 设置页：打开设置对话框 → 办公模式 → 量字号与控件
        //
        // 第七十七轮：设置入口从「文字按钮」变成了**图标按钮**（DSH 0.2.2-rc.3 实测：
        // 宿主 settings 座位的 trigger 只带 aria-label="设置" + aria-haspopup="dialog"，
        // 侧栏里看不到「设置」这两个字）。旧探针只按 innerText 找，找不到还继续跑，
        // 最后在 round35.cleared.groups 上抛 TypeError —— 真正的失败原因被空值引用
        // 盖住了。现在按 文本 / aria-label / aria-haspopup 三路找，找不到就 dump
        // 候选按钮并立刻失败。
        findings.dom.settings = {};
        findings.dom.settings.trigger = await evaluate(`(() => {
            const buttons = Array.from(document.querySelectorAll('button,[role="button"]'));
            const label = (el) => {
                const aria = (el.getAttribute('aria-label') || '').trim();
                return aria !== '' ? aria : (el.innerText || '').replace(/\\s+/g, ' ').trim();
            };
            let button = buttons.find((b) => label(b) === '设置');
            if (!button) {
                button = buttons.find((b) => b.getAttribute('aria-haspopup') === 'dialog'
                    && ((b.getAttribute('aria-label') || '').includes('设置')
                        || (b.getAttribute('title') || '').includes('设置')));
            }
            if (button) {
                button.click();
                return { found: true, via: 'direct', aria: button.getAttribute('aria-label') || null, text: (button.innerText || '').trim().slice(0, 16) || null };
            }
            // DSH 0.2.2-rc.3：settings 座位的 launcher 被账号菜单接管，页面上没有独立的
            // 「设置」按钮；设置项在账号菜单（底部头像按钮，aria-haspopup=menu）里。
            // 注意：按钮可能在「收起侧栏」的隐藏分支里，JS click 对隐藏节点没反应 ——
            // 所以只取有尺寸的那一个，并把坐标交给 CDP 真点。
            const visible = (b) => {
                const r = b.getBoundingClientRect();
                return r.width > 0 && r.height > 0;
            };
            const account = buttons.find((b) => b.querySelector('[data-avatar-status]') && visible(b))
                || buttons.filter((b) => b.getAttribute('aria-haspopup') === 'menu' && visible(b)).pop();
            if (account) {
                const rect = account.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                const hit = document.elementFromPoint(x, y);
                return {
                    found: false,
                    needsMenu: true,
                    accountAria: account.getAttribute('aria-label') || null,
                    expanded: account.getAttribute('aria-expanded'),
                    rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height },
                    hit: hit === null ? null : {
                        tag: hit.tagName,
                        aria: hit.getAttribute('aria-label') || null,
                        text: (hit.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 24) || null,
                    },
                    x: x,
                    y: y,
                };
            }
            return {
                found: false,
                candidates: buttons.slice(0, 80).map((b) => ({
                    aria: b.getAttribute('aria-label') || null,
                    title: b.getAttribute('title') || null,
                    text: (b.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 16) || null,
                    haspopup: b.getAttribute('aria-haspopup') || null,
                    visible: visible(b),
                })).filter((item) => item.aria || item.title || item.text),
            };
        })()`);
        if (findings.dom.settings.trigger.needsMenu === true) {
            const point = findings.dom.settings.trigger;
            log(`   设置入口不在页面里，点账号菜单（${point.accountAria} @ ${Math.round(point.x)},${Math.round(point.y)}）`);
            await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'left', clickCount: 1, buttons: 1 }, sessionId);
            await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'left', clickCount: 1, buttons: 0 }, sessionId);
            await sleep(700);
            await shot('settings-account-menu.png');
            findings.dom.settings.menuState = await evaluate(`(() => {
                const hits = (selector) => document.querySelectorAll(selector).length;
                return {
                    menuCount: hits('[role="menu"]'),
                    itemCount: hits('[role="menuitem"]'),
                    shadowHosts: Array.from(document.querySelectorAll('*')).filter((e) => e.shadowRoot).length,
                    expanded: (() => {
                        const b = Array.from(document.querySelectorAll('button')).find((e) => e.getAttribute('aria-label') === '账号菜单');
                        return b ? b.getAttribute('aria-expanded') : null;
                    })(),
                    textHasSettings: (document.body.innerText || '').includes('设置'),
                };
            })()`);
            findings.dom.settings.menu = await evaluate(`(() => {
                const items = Array.from(document.querySelectorAll('[role="menuitem"]'));
                // 菜单项的 innerText 带快捷键后缀（"设置 Ctrl + Alt + ,"），不能用全等；
                // 去掉空白与快捷键提示后按前缀匹配。
                const norm = (e) => ((e.innerText || '').replace(/\\s+/g, ' ').trim());
                const target = items.find((e) => {
                    const text = norm(e);
                    return text === '设置' || text.startsWith('设置 ') || text.startsWith('设置,')
                        || (e.getAttribute('aria-label') || '').trim() === '设置';
                });
                if (!target) {
                    return {
                        clicked: false,
                        count: items.length,
                        items: items.map((e) => ((e.getAttribute('aria-label') || '') + '|' + (e.innerText || '')).replace(/\\s+/g, ' ').trim()).filter((t) => t !== '|').slice(0, 40),
                    };
                }
                const rect = target.getBoundingClientRect();
                return { clicked: true, x: rect.left + rect.width / 2, y: rect.top + rect.height / 2, label: norm(target).slice(0, 24) };
            })()`);
            if (findings.dom.settings.menu.clicked === true) {
                const menuPoint = findings.dom.settings.menu;
                await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: menuPoint.x, y: menuPoint.y, button: 'left', clickCount: 1, buttons: 1 }, sessionId);
                await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: menuPoint.x, y: menuPoint.y, button: 'left', clickCount: 1, buttons: 0 }, sessionId);
            } else {
                // 菜单没开（宿主改版 / 菜单在 portal 外）：用宿主的设置快捷键兜底
                // （web:windows = Ctrl+Alt+Comma；settings.open 命令）。
                log('   账号菜单没开出设置项，改用设置快捷键 Ctrl+Alt+,');
                await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', modifiers: 3, key: ',', code: 'Comma', windowsVirtualKeyCode: 188, nativeVirtualKeyCode: 188 }, sessionId);
                await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', modifiers: 3, key: ',', code: 'Comma', windowsVirtualKeyCode: 188, nativeVirtualKeyCode: 188 }, sessionId);
                findings.dom.settings.shortcut = true;
            }
            await sleep(1000);
        }
        // 对话框是否真的打开，以「设置页里出现办公模式页签」为准：只记录“点击/快捷键已派发”
        // 会把菜单没点中、快捷键被吞记成打开（第七十七轮实测：菜单已开、快捷键无效，
        // 旧口径仍写 opened=true，直到页签断言才红）。轮询也顺便点页签。
        let officeNav = { found: false };
        for (let attempt = 0; attempt < 20; attempt += 1) {
            officeNav = await evaluate(`(() => {
                const items = Array.from(document.querySelectorAll('button,[role="tab"],[role="button"]'));
                const item = items.find((e) => {
                    const text = (e.innerText || '').replace(/\\s+/g, ' ').trim();
                    const aria = (e.getAttribute('aria-label') || '').trim();
                    return text === '办公模式' || aria === '办公模式';
                });
                if (!item) {
                    return {
                        found: false,
                        dialogs: document.querySelectorAll('[role="dialog"]').length,
                        menuItems: document.querySelectorAll('[role="menuitem"]').length,
                    };
                }
                item.click();
                return { found: true };
            })()`);
            if (officeNav.found) break;
            await sleep(250);
        }
        findings.dom.settings.officeNav = officeNav.found;
        findings.dom.settings.opened = officeNav.found;
        findings.dom.settings.dialogState = officeNav;
        if (!officeNav.found) {
            const visibleTabs = await evaluate(`(() => Array.from(document.querySelectorAll('button,[role="tab"],[role="button"]')).map((e) => ((e.getAttribute('aria-label') || '') + '|' + (e.innerText || '')).replace(/\\s+/g, ' ').trim()).filter((t) => t !== '|').slice(0, 80))()`);
            fail(`设置页没打开（dialogs=${officeNav.dialogs}, menuItems=${officeNav.menuItems}）。触发：${JSON.stringify(findings.dom.settings.trigger)}${findings.dom.settings.menu === undefined ? '' : `；菜单：${JSON.stringify(findings.dom.settings.menu)}`}；可见页签：${JSON.stringify(visibleTabs)}`);
        }
        await sleep(1000);
        // 设置对话框打开后，办公模式页要真的挂上来：轮询到 [data-locator="office"]
        // 出现为止（慢机器上 1s 不够）；等不到就把可见页签打出来，别让后面的空值
        // 引用替真正的失败顶包。
        let officeReady = false;
        for (let attempt = 0; attempt < 20; attempt += 1) {
            officeReady = await evaluate(`document.querySelector('[data-locator="office"]') !== null`);
            if (officeReady) break;
            await sleep(250);
        }
        findings.dom.settings.officeReady = officeReady;
        if (!officeReady) {
            const visibleTabs = await evaluate(`(() => Array.from(document.querySelectorAll('button,[role="tab"],[role="button"]')).map((e) => ((e.getAttribute('aria-label') || '') + '|' + (e.innerText || '')).replace(/\\s+/g, ' ').trim()).filter((t) => t !== '|').slice(0, 80))()`);
            fail(`设置页：办公模式页没挂载（officeNav=${findings.dom.settings.officeNav}）。可见页签：${JSON.stringify(visibleTabs)}`);
        }
        // 第八十七轮：办公模式页不再放一张没有分组头的「记忆」指路裸卡。
        findings.dom.settings.locatorNotes = await evaluate('document.querySelectorAll("[data-locator-note]").length');
        if (findings.dom.settings.locatorNotes !== 0) {
            fail(`设置页：办公模式页仍有 ${findings.dom.settings.locatorNotes} 张 data-locator-note 裸卡`);
        }
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

        // ── 第三十五轮：设置页的分组折叠与「定位」 ──
        //
        // 这一段量三件事，都是单测量不到的：
        //   1) 默认收起是不是**真的**收起（组体内的控件 offsetParent 为 null、
        //      组体 display:none），而不是只写了个属性；
        //   2) 「定位」框在真实 React 受控输入下过滤得对不对 —— 用原生 value setter
        //      + input 事件，而不是直接改 .value（那样 React 的 onChange 不触发）；
        //   3) 定位条 sticky 有没有生效（设置页展开后这一页很长，输入框得跟着留）。
        log('\n── 第三十五轮：设置页分组折叠与定位 ──');
        // 以定位条所在的那个 section 为范围：设置对话框里还有别的插件页面。
        const OFFICE_ROOT = `document.querySelector('[data-locator="office"]').parentElement`;
        const readOffice = () => evaluate(`(() => {
            const root = ${OFFICE_ROOT};
            if (!root) return null;
            const bar = root.querySelector('[data-locator="office"]');
            return {
                groups: Array.from(root.querySelectorAll('[data-group]')).map((el) => {
                    const body = el.querySelector('[data-group-body]');
                    return {
                        id: el.getAttribute('data-group'),
                        open: el.getAttribute('data-group-open'),
                        head: (() => { const b = el.querySelector('[data-group-toggle]'); return b ? (b.innerText || '').replace(/\\s+/g, ' ').trim() : null; })(),
                        bodyDisplay: body ? getComputedStyle(body).display : null,
                        controls: body ? body.querySelectorAll('input,select,textarea').length : 0,
                        visibleControls: body ? Array.from(body.querySelectorAll('input,select,textarea')).filter((el) => el.offsetParent !== null).length : 0,
                    };
                }),
                stats: bar && bar.querySelector('[data-locator-stats]') ? bar.querySelector('[data-locator-stats]').innerText.trim() : null,
                storage: (() => { try { return window.localStorage.getItem('dsh-office-mode.groups.office'); } catch (e) { return 'storage-unavailable'; } })(),
            };
        })()`);
        /** 受控输入：必须走原生 setter + input 事件，直接改 .value 不会触发 React 的 onChange。 */
        const typeLocator = (page, text) => evaluate(`(() => {
            const input = document.querySelector('[data-locator-input="${page}"]');
            if (!input) return null;
            const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            setter.call(input, ${JSON.stringify(text)});
            input.dispatchEvent(new Event('input', { bubbles: true }));
            return input.value;
        })()`);

        findings.dom.round35 = {};
        findings.dom.round35.collapsed = await readOffice();
        if (findings.dom.round35.collapsed === null) {
            fail('设置页：办公模式页读不到分组（[data-locator="office"] 又不见了）');
        }
        if (findings.dom.round35.collapsed.groups.length === 0) {
            fail('设置页：办公模式页一个分组都没渲染出来（定位条在、[data-group] 不在）');
        }
        log(`默认态：${JSON.stringify(findings.dom.round35.collapsed)}`);

        // 点开「检索编排」，量展开真的发生（组体不再 display:none、控件可见）
        findings.dom.round35.expandClick = await evaluate(`(() => {
            const t = document.querySelector('[data-group-toggle="search"]');
            if (!t) return false;
            t.click();
            return true;
        })()`);
        await sleep(700);
        findings.dom.round35.afterExpand = await readOffice();

        // 全部展开 → 量 sticky（页面被撑长之后）
        await evaluate(`(() => { const b = document.querySelector('[data-groups-expand="office"]'); if (b) b.click(); return !!b; })()`);
        await sleep(900);
        findings.dom.round35.allExpanded = await readOffice();
        findings.dom.round35.sticky = await evaluate(`(() => {
            const bar = document.querySelector('[data-locator="office"]');
            if (!bar) return null;
            let scroller = bar.parentElement;
            while (scroller && scroller !== document.body) {
                const style = getComputedStyle(scroller);
                if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && scroller.scrollHeight > scroller.clientHeight + 4) break;
                scroller = scroller.parentElement;
            }
            if (!scroller || scroller === document.body) return { found: false, position: getComputedStyle(bar).position };
            const topBefore = Math.round(bar.getBoundingClientRect().top - scroller.getBoundingClientRect().top);
            const previous = scroller.scrollTop;
            scroller.scrollTop = 320;
            const scrolled = scroller.scrollTop;
            const topAfter = Math.round(bar.getBoundingClientRect().top - scroller.getBoundingClientRect().top);
            scroller.scrollTop = previous;
            return { found: true, position: getComputedStyle(bar).position, topBefore, topAfter, scrolled,
                     stuck: Math.abs(topAfter - topBefore) <= 2 };
        })()`);
        log(`sticky：${JSON.stringify(findings.dom.round35.sticky)}`);
        await shot('settings-office-groups.png');

        // 定位：一个词命中一行
        findings.dom.round35.typed = await typeLocator('office', '出口代理');
        await sleep(700);
        findings.dom.round35.located = await evaluate(`(() => {
            const root = ${OFFICE_ROOT};
            if (!root) return null;
            return {
                groups: Array.from(root.querySelectorAll('[data-group]')).map((el) => el.getAttribute('data-group')),
                rows: Array.from(root.querySelectorAll('[data-group-body]')).map((body) => Array.from(body.querySelectorAll('input,select')).length),
                hasProxyRow: (root.innerText || '').includes('出口代理'),
                hasChannelRow: (root.innerText || '').includes('联网通道'),
                stats: (() => { const el = root.querySelector('[data-locator-stats]'); return el ? el.innerText.trim() : null; })(),
            };
        })()`);
        log(`定位「出口代理」：${JSON.stringify(findings.dom.round35.located)}`);

        // 定位：命中折在 Fold 里的高级参数 → 折叠块要自动展开
        await typeLocator('office', 'search.providers.deepseek-api.timeoutMs');
        await sleep(700);
        findings.dom.round35.locatedInFold = await evaluate(`(() => {
            const root = ${OFFICE_ROOT};
            const body = root ? root.querySelector('[data-fold-body="search-advanced"]') : null;
            return {
                stats: (() => { const el = root.querySelector('[data-locator-stats]'); return el ? el.innerText.trim() : null; })(),
                foldBodyDisplay: body ? getComputedStyle(body).display : null,
                foldToggleExpanded: (() => { const t = root.querySelector('[data-fold-toggle="search-advanced"]'); return t ? t.getAttribute('aria-expanded') : null; })(),
                hasProxyRow: (root.innerText || '').includes('出口代理'),
            };
        })()`);
        log(`定位高级参数：${JSON.stringify(findings.dom.round35.locatedInFold)}`);
        await shot('settings-office-locate.png');

        // 定位：没有命中时的提示；清空后回到全量 + 全部收起
        await typeLocator('office', 'zzz-没有这一项');
        await sleep(600);
        findings.dom.round35.noHit = await evaluate(`(() => {
            const root = ${OFFICE_ROOT};
            const groups = Array.from(root.querySelectorAll('[data-group]')).map((el) => el.getAttribute('data-group'));
            const el = root.querySelector('[data-locator-stats]');
            return { groups, stats: el ? el.innerText.trim() : null };
        })()`);
        await typeLocator('office', '');
        await sleep(600);
        await evaluate(`(() => { const b = document.querySelector('[data-groups-collapse="office"]'); if (b) b.click(); return !!b; })()`);
        await sleep(700);
        findings.dom.round35.cleared = await readOffice();
        log(`没有命中：${JSON.stringify(findings.dom.round35.noHit)}`);
        if (findings.dom.round35.cleared === null) {
            fail('设置页：清空定位后读不到办公模式设置页（[data-locator="office"] 不见了）');
        }
        log(`清空 + 全部收起后：${JSON.stringify({ groups: findings.dom.round35.cleared.groups, stats: findings.dom.round35.cleared.stats, storage: findings.dom.round35.cleared.storage })}`);

        // OCR 设置只查入口和空密码框，不编辑，不触发宿主凭据读取或写入。
        await evaluate(`(() => {
            const group = document.querySelector('[data-group="ocr"]');
            if (group?.getAttribute('data-group-open') === 'false') group.querySelector('[data-group-toggle="ocr"]')?.click();
        })()`);
        await sleep(500);
        findings.dom.round35.ocrSettings = await evaluate(`(() => {
            const group = document.querySelector('[data-group="ocr"]');
            const key = group?.querySelector('[data-ocr-key]');
            const reuse = group?.querySelector('[data-ocr-reuse-official]');
            return { open: group?.getAttribute('data-group-open') === 'true',
                password: key?.type === 'password', empty: key?.value === '',
                reuseToggle: reuse?.type === 'checkbox',
                officialReuse: (group?.innerText || '').includes('deepseek-official') };
        })()`);
        const ocrSettings = findings.dom.round35.ocrSettings;
        if (!ocrSettings?.open || !ocrSettings.password || !ocrSettings.empty || !ocrSettings.reuseToggle || !ocrSettings.officialReuse) {
            fail('设置页：OCR 分组不可打开、密码框未保持空白，或缺少官方凭据复用开关与说明');
        }
        log('文件识别（OCR）设置：折叠区可打开，密码框空白，复用开关存在，只读检查未写设置');

        // 记忆系统页：同一套骨架（定位框按页分开）
        await evaluate(`(() => {
            const item = Array.from(document.querySelectorAll('button,[role="tab"],[role="button"]')).find((e) => (e.innerText || '').trim() === '记忆系统');
            if (!item) return false;
            item.click();
            return true;
        })()`);
        await sleep(1000);
        findings.dom.round35.memoryPage = await evaluate(`(() => {
            const bar = document.querySelector('[data-locator="memory"]');
            if (!bar) return null;
            const root = bar.parentElement;
            return {
                groups: Array.from(root.querySelectorAll('[data-group]')).map((el) => el.getAttribute('data-group')),
                open: Array.from(root.querySelectorAll('[data-group]')).map((el) => el.getAttribute('data-group-open')),
                stats: bar.querySelector('[data-locator-stats]') ? bar.querySelector('[data-locator-stats]').innerText.trim() : null,
                storage: (() => { try { return window.localStorage.getItem('dsh-office-mode.groups.memory'); } catch (e) { return 'storage-unavailable'; } })(),
            };
        })()`);
        if (findings.dom.round35.memoryPage === null) {
            fail('设置页：记忆系统页读不到定位条（[data-locator="memory"] 不见了）');
        }
        // 第八十八轮：知识库只有一个库，目录走资源管理器接口，模型随目录下载。
        findings.dom.round35.kbSettings = await evaluate(`(() => ({
            dirInput: !!document.querySelector('[data-kb-dir-input]'),
            pickButton: (() => { const b = document.querySelector('[data-kb-dir-pick]'); return b ? (b.innerText || '').trim() : null; })(),
            models: Array.from(document.querySelectorAll('[data-model-kind]')).map((el) => ({
                kind: el.getAttribute('data-model-kind'),
                state: el.getAttribute('data-model-state'),
                download: !!el.querySelector('[data-model-download]'),
            })),
        }))()`);
        if (!findings.dom.round35.kbSettings.dirInput || findings.dom.round35.kbSettings.pickButton === null) {
            fail('设置页：知识库目录没有资源管理器选择入口（[data-kb-dir-input] / [data-kb-dir-pick]）');
        }
        if (findings.dom.round35.kbSettings.models.length !== 2) {
            fail(`设置页：应有 Qwen3-Embedding / Qwen3-Reranker 两张模型卡，实际 ${findings.dom.round35.kbSettings.models.length}`);
        }
        log(`知识库目录与模型：${JSON.stringify(findings.dom.round35.kbSettings)}`);
        // 第八十九轮：目录选择不再依赖宿主 uiWorkspace（Web 端常常没有），
        // 点「选择目录」必须真的弹出插件自己的目录浏览器并列出子目录。
        await evaluate(`(() => { const b = document.querySelector('[data-kb-dir-pick]'); if (b) { b.click(); return true; } return false; })()`);
        await sleep(1200);
        findings.dom.round35.dirBrowser = await evaluate(`(() => {
            const box = document.querySelector('[data-dir-browser="true"]');
            if (!box) return null;
            return {
                open: true,
                path: (() => { const el = box.querySelector('[data-dir-path]'); return el ? el.value : null; })(),
                entries: Array.from(box.querySelectorAll('[data-dir-entry]')).slice(0, 8).map((e) => (e.innerText || '').replace(/\s+/g, ' ').trim()),
                roots: Array.from(box.querySelectorAll('[data-dir-root]')).map((e) => (e.innerText || '').trim()),
                error: (() => { const el = box.querySelector('[data-dir-error]'); return el ? el.innerText.trim() : null; })(),
            };
        })()`);
        log(`目录浏览器：${JSON.stringify(findings.dom.round35.dirBrowser)}`);
        if (findings.dom.round35.dirBrowser === null) {
            fail('设置页：点「选择目录」没有弹出目录浏览器（[data-dir-browser]）');
        }
        await shot('settings-kb-dir-browser.png');
        await evaluate(`(() => {
            const box = document.querySelector('[data-dir-browser="true"]');
            if (!box) return false;
            const close = Array.from(box.querySelectorAll('button')).find((b) => (b.innerText || '').trim() === '关闭');
            if (close) close.click();
            return !!close;
        })()`);
        await sleep(400);
        log(`记忆系统页：${JSON.stringify(findings.dom.round35.memoryPage)}`);
        await shot('settings-memory-groups.png');

        // ── 第九十五轮：知识库大模型（kb-llm）的通道与宿主清单 ──
        // 只读回归：off / openai 通道不该为打开这一页联网；host 通道优先走
        // session.modelCatalog RPC、旧宿主才退到 GET /office-memory/llm-models
        // （都不带 Key、不调模型），拿到什么就如实画什么（provider 下拉 / 空态 /
        // 宿主错误原文 + 重试）。这里**不切通道**：设置写入会落进用户 profile 的
        // 配置，与本文件「只开不存」的既有约定一致（选择与持久化归 test/client.mjs
        // 的假 RPC 守）。所以 host 分支只在用户自己已经选了 host 时才真正跑到。
        log('\n── 第九十五轮：知识库大模型的通道与宿主清单 ──');
        findings.dom.round95 = {};
        const kbLlmBefore = findings.kbLlmRequests.length;
        const MEMORY_ROOT = `document.querySelector('[data-locator="memory"]').parentElement`;
        findings.dom.round95.entry = await evaluate(`(() => {
            const root = ${MEMORY_ROOT};
            if (!root) return null;
            const group = root.querySelector('[data-group="kb-llm"]');
            if (!group) return null;
            const wasOpen = group.getAttribute('data-group-open');
            const toggle = root.querySelector('[data-group-toggle="kb-llm"]');
            if (wasOpen !== 'true' && toggle) toggle.click();
            return { found: true, wasOpen: wasOpen };
        })()`);
        if (!findings.dom.round95.entry) {
            fail('设置页：记忆系统页找不到「知识库大模型（可选）」分组（[data-group="kb-llm"]）');
        }
        await sleep(700);
        /** 读 kb-llm 分组当前的真实 DOM；provider / 模型名只留在内存里做比对，不写进产物。 */
        const readKbLlm = () => evaluate(`(() => {
            const root = ${MEMORY_ROOT};
            const body = root ? root.querySelector('[data-group-body="kb-llm"]') : null;
            if (!body) return null;
            const group = body.closest('[data-group]');
            const selects = Array.from(body.querySelectorAll('select'));
            const hasOption = (sel, value) => Array.from(sel.options).some((o) => o.value === value);
            const channel = selects.find((s) => hasOption(s, 'off') && hasOption(s, 'host') && hasOption(s, 'openai')) || null;
            const providerSelect = selects.find((s) => Array.from(s.options).some((o) => o.textContent.trim() === '请选择 provider')) || null;
            const modelSelect = selects.find((s) => Array.from(s.options).some((o) => o.textContent.trim() === '请选择模型')) || null;
            const pick = (name) => Array.from(body.querySelectorAll('[data-llm-manual]')).find((el) => el.getAttribute('data-llm-manual') === name) || null;
            const textOf = (selector) => { const el = body.querySelector(selector); return el ? String(el.innerText || '').trim() : null; };
            const visible = (el) => !!el && el.offsetParent !== null;
            const manual = (name) => { const el = pick(name); return el === null ? null : { visible: visible(el), placeholder: el.getAttribute('placeholder'), value: el.value }; };
            const optionsOf = (el) => el === null ? null : Array.from(el.options).map((o) => ({ value: o.value, label: o.textContent.trim() }));
            return {
                open: group ? group.getAttribute('data-group-open') : null,
                bodyDisplay: getComputedStyle(body).display,
                text: String(body.innerText || '').slice(0, 1600),
                channel: channel === null ? null : { visible: visible(channel), value: channel.value, options: optionsOf(channel) },
                manualProvider: manual('memory.llm.provider'),
                manualModel: manual('memory.llm.model'),
                providerSelect: providerSelect === null ? null : { visible: visible(providerSelect), value: providerSelect.value, options: optionsOf(providerSelect) },
                modelSelect: modelSelect === null ? null : { visible: visible(modelSelect), value: modelSelect.value, options: optionsOf(modelSelect) },
                loading: textOf('[data-llm-catalog-status="loading"]'),
                empty: textOf('[data-llm-catalog-empty]'),
                error: textOf('[data-llm-catalog-error]'),
                retry: !!body.querySelector('[data-llm-catalog-retry]'),
                savedProvider: textOf('[data-llm-saved-provider]'),
                savedModel: textOf('[data-llm-saved-model]'),
            };
        })()`);
        /** host 通道 provider 下拉的取值；除占位与已存值外，剩下的才是清单项。 */
        const providerOptionValues = (state) => state === null || state.providerSelect === null
            ? null : state.providerSelect.options.map((o) => o.value);
        const providerExtraOptions = (state) => {
            const values = providerOptionValues(state);
            if (values === null) return null;
            const saved = state.providerSelect.value;
            return values.filter((value) => value !== '' && value !== saved);
        };
        /** 产物里只留形态与计数：provider / 模型 id 是宿主配置，不进 diag JSON。 */
        const kbLlmShape = (state) => state === null ? null : ({
            open: state.open,
            bodyDisplay: state.bodyDisplay,
            channel: state.channel,
            manualProvider: state.manualProvider === null ? null : { visible: state.manualProvider.visible, placeholder: state.manualProvider.placeholder, hasValue: state.manualProvider.value !== '' },
            manualModel: state.manualModel === null ? null : { visible: state.manualModel.visible, placeholder: state.manualModel.placeholder, hasValue: state.manualModel.value !== '' },
            providerSelect: state.providerSelect === null ? null : { visible: state.providerSelect.visible, optionCount: state.providerSelect.options.length, hasValue: state.providerSelect.value !== '' },
            modelSelect: state.modelSelect === null ? null : { visible: state.modelSelect.visible, optionCount: state.modelSelect.options.length, firstLabel: state.modelSelect.options.length > 0 ? state.modelSelect.options[0].label : null },
            loading: state.loading,
            empty: state.empty,
            error: state.error === null ? null : state.error.split('\n')[0].slice(0, 80),
            errorPresent: state.error !== null,
            retry: state.retry,
            savedProvider: state.savedProvider !== null,
            savedModel: state.savedModel !== null,
        });
        let kbLlm = await readKbLlm();
        if (!kbLlm) fail('设置页：kb-llm 分组没有渲染出组体（[data-group-body="kb-llm"]）');
        findings.dom.round95.initial = kbLlmShape(kbLlm);
        if (kbLlm.open !== 'true' || kbLlm.bodyDisplay === 'none') {
            fail(`设置页：kb-llm 分组没展开（open=${kbLlm.open}, display=${kbLlm.bodyDisplay}）`);
        }
        if (kbLlm.channel === null) fail('设置页：kb-llm 分组里找不到「通道」下拉（off / host / openai 三档）');
        if (kbLlm.channel.visible !== true) fail('设置页：kb-llm 的「通道」下拉不可见（分组展开了但控件没露出来）');
        const channelValues = kbLlm.channel.options.map((o) => o.value).join(',');
        const channelLabels = kbLlm.channel.options.map((o) => o.label).join('|');
        if (channelValues !== 'off,host,openai') fail(`设置页：通道下拉的取值不是 off/host/openai：${channelValues}`);
        if (channelLabels !== '关（默认）|宿主已配置模型|OpenAI 兼容端点') {
            fail(`设置页：通道下拉的标签变了（默认关 / 宿主 / OpenAI 兼容）：${channelLabels}`);
        }
        if (kbLlm.text.indexOf('默认关：不联网、不花钱') === -1) {
            fail('设置页：kb-llm 分组丢了「默认关不联网、不花钱」这句（隐私与计费提示必须留着）');
        }
        const kbLlmMode = kbLlm.channel.value;
        findings.dom.round95.mode = kbLlmMode;
        if (kbLlmMode !== 'host') {
            // off / openai：provider 与模型都是手填；只为打开设置页就拉一次清单是白联网。
            if (kbLlm.manualProvider === null || kbLlm.manualProvider.visible !== true || kbLlm.manualProvider.placeholder !== 'provider id') {
                fail(`设置页：${kbLlmMode} 通道的 provider 手填框不对：${JSON.stringify(kbLlm.manualProvider)}`);
            }
            if (kbLlm.manualModel === null || kbLlm.manualModel.visible !== true || kbLlm.manualModel.placeholder !== '模型名') {
                fail(`设置页：${kbLlmMode} 通道的模型手填框不对：${JSON.stringify(kbLlm.manualModel)}`);
            }
            if (kbLlm.loading !== null || kbLlm.empty !== null || kbLlm.error !== null) {
                fail(`设置页：${kbLlmMode} 通道不该出现宿主清单的状态块：${JSON.stringify({ loading: kbLlm.loading, empty: kbLlm.empty, error: kbLlm.error })}`);
            }
            await sleep(1200);
            if (findings.kbLlmRequests.length !== kbLlmBefore) {
                fail(`设置页：${kbLlmMode} 通道展开 kb-llm 分组却去拉了宿主模型清单（${findings.kbLlmRequests.length - kbLlmBefore} 次）`);
            }
            log(`kb-llm：通道 ${kbLlmMode}，不联网；provider / 模型都手填（0 次清单请求）`);
        } else {
            // host：provider 下拉在 host 通道始终渲染（占位 +（可选）已存值 + 成功时清单项），
            // 不能拿它当结算信号；loading 消失并隔一拍确认仍为空，才算这次取数落定。
            let settled = kbLlm;
            for (let attempt = 0; attempt < 40; attempt += 1) {
                if (settled.loading !== null) {
                    await sleep(250);
                    settled = await readKbLlm();
                    continue;
                }
                await sleep(300);
                const confirm = await readKbLlm();
                settled = confirm === null ? settled : confirm;
                if (confirm !== null && confirm.loading === null) break;
            }
            findings.dom.round95.host = kbLlmShape(settled);
            const requests = findings.kbLlmRequests.length - kbLlmBefore;
            if (settled.loading !== null) fail('设置页：host 通道的宿主清单一直停在「正在读取…」');
            if (settled.error !== null) {
                if (settled.retry !== true) fail('设置页：宿主清单读失败却没有「重试」按钮');
                // 产品实际：失败时下拉仍在，只剩占位 +（可选）已存值；只验「没有清单项」，不验不存在。
                if (settled.providerSelect === null) fail('设置页：host 通道读清单失败却没画 provider 下拉（产品约定：始终保留下拉与已存值）');
                if (settled.providerSelect.options.length > 2) fail(`设置页：读失败时 provider 下拉项太多（${settled.providerSelect.options.length} 个）`);
                if (providerExtraOptions(settled).length !== 0) fail(`设置页：读失败时 provider 下拉里不该有清单项（多出 ${providerExtraOptions(settled).length} 个）`);
                if (settled.savedProvider !== null || settled.savedModel !== null) fail('设置页：清单没读到时不该标「已保存的值不在清单里」');
                const before = findings.kbLlmRequests.length;
                const rpcBefore = findings.kbLlmRpcFrames ?? 0;
                const httpBefore = findings.kbLlmRpcHttp ?? 0;
                const wsBefore = findings.kbLlmRpcWs ?? 0;
                const clicked = await evaluate(`(() => {
                    const root = ${MEMORY_ROOT};
                    const body = root ? root.querySelector('[data-group-body="kb-llm"]') : null;
                    const button = body ? body.querySelector('[data-llm-catalog-retry]') : null;
                    if (!button || button.disabled) return false;
                    button.click();
                    return true;
                })()`);
                if (!clicked) fail('设置页：无法点击宿主清单重试按钮');
                let again = null;
                let requestDelta = 0;
                let rpcDelta = 0;
                let httpDelta = 0;
                let wsDelta = 0;
                for (let attempt = 0; attempt < 60; attempt += 1) {
                    await sleep(200);
                    again = await readKbLlm();
                    requestDelta = findings.kbLlmRequests.length - before;
                    rpcDelta = (findings.kbLlmRpcFrames ?? 0) - rpcBefore;
                    httpDelta = (findings.kbLlmRpcHttp ?? 0) - httpBefore;
                    wsDelta = (findings.kbLlmRpcWs ?? 0) - wsBefore;
                    if (again !== null && again.loading === null && (requestDelta > 0 || rpcDelta > 0)) break;
                }
                if (requestDelta <= 0 && rpcDelta <= 0) {
                    fail('设置页：点击重试后没有新的清单 HTTP 请求或 modelCatalog RPC 帧');
                }
                if (!again) fail('设置页：重试后读不到 kb-llm 分组');
                // 终态：错误 / 空态 / 下拉里占位之外还有一项（清单项或已存值），三者之一都是正常。
                if (again.empty === null && again.error === null
                    && (again.providerSelect === null || again.providerSelect.options.length <= 1)) {
                    fail('设置页：重试后清单既没成功，也没给出空态或错误');
                }
                findings.dom.round95.hostRetry = { requests: findings.kbLlmRequests.length - kbLlmBefore, shape: kbLlmShape(again) };
                log(`kb-llm：host 通道读清单失败，如实显示宿主原文并给了重试（UI 清单 HTTP ${requests} → ${findings.kbLlmRequests.length - kbLlmBefore} 次，新增 RPC ${rpcDelta} 次：HTTP POST ${httpDelta} / WS ${wsDelta}）`);
            } else if (settled.empty !== null) {
                // 产品实际：空清单时下拉仍在，只剩占位 +（可选）已存值。
                if (settled.providerSelect === null) fail('设置页：宿主清单为空却没画 provider 下拉（产品约定：始终保留下拉与已存值）');
                if (settled.providerSelect.options.length > 2) fail(`设置页：空清单时 provider 下拉项太多（${settled.providerSelect.options.length} 个）`);
                if (providerExtraOptions(settled).length !== 0) fail(`设置页：空清单时 provider 下拉里不该有清单项（多出 ${providerExtraOptions(settled).length} 个）`);
                if (settled.empty.indexOf('没有可用的 provider') === -1) fail(`设置页：空清单的提示变了：${settled.empty}`);
                if (settled.savedProvider !== null || settled.savedModel !== null) fail('设置页：清单为空时不该标「已保存的值不在清单里」');
                log(`kb-llm：host 通道拉到空清单，如实显示空态（${requests} 次清单请求）`);
            } else {
                // 有清单：RPC 会过滤「没有模型的 provider」，与旧 HTTP 清单不能作等集合比较。
                // 只验两件事：下拉里的清单项都来自 HTTP 成功 body 的 provider id（已存值例外）；
                // 且没有 RPC partial failures 时，至少一个「有模型」的 HTTP provider 真的出现。
                if (settled.providerSelect.visible !== true) fail('设置页：host 通道拿到清单却没画出 provider 下拉');
                const savedProviderValue = settled.providerSelect.value;
                const shownAll = settled.providerSelect.options.map((o) => o.value).filter((value) => value !== '');
                const shown = shownAll.filter((value) => value !== savedProviderValue);
                const live = await evaluate(`fetch('/office-memory/llm-models', { headers: { accept: 'application/json', 'x-office-verify-probe': '1' } }).then((r) => r.json()).then((b) => ({ ok: b && b.ok === true, providers: b && Array.isArray(b.providers) ? b.providers.map((p) => ({ id: String((p && (p.provider || p.id)) || ''), models: p && Array.isArray(p.models) ? p.models.length : 0 })).filter((p) => p.id !== '') : [] })).catch(() => ({ ok: false, providers: [] }))`);
                const partial = await evaluate(`(() => { const root = ${MEMORY_ROOT}; const el = root ? root.querySelector('[data-llm-catalog-failures]') : null; return el ? String(el.innerText || '').trim().slice(0, 160) : null; })()`);
                let compared = null;
                let fabricated = 0;
                if (live && live.ok === true && live.providers.length > 0) {
                    const known = live.providers.map((p) => p.id);
                    const made = shown.filter((id) => known.indexOf(id) === -1);
                    fabricated = made.length;
                    if (made.length !== 0) fail(`设置页：provider 下拉里出现了 HTTP 清单没有的项（${made.length} 个）`);
                    const withModels = live.providers.filter((p) => p.models > 0).map((p) => p.id);
                    if (withModels.length > 0 && partial === null) {
                        const hit = withModels.filter((id) => shownAll.indexOf(id) !== -1);
                        if (hit.length === 0) fail(`设置页：HTTP 有 ${withModels.length} 个「有模型」的 provider，下拉里一个都没出现（RPC 只过滤没模型的，有模型的不该消失）`);
                        compared = hit.length;
                    } else {
                        compared = 0;
                    }
                }
                if (settled.modelSelect !== null) {
                    const first = settled.modelSelect.options[0];
                    if (first.value !== '' || first.label !== '请选择模型') fail(`设置页：模型下拉的第一项不是占位：${JSON.stringify(first)}`);
                }
                findings.dom.round95.hostCatalog = { shown: shown.length, compared: compared, fabricated: fabricated, partial: partial !== null, requests: requests };
                log(compared === null
                    ? `kb-llm：host 通道下拉 ${shown.length} 个清单项；HTTP 旁路不可用或清单为空，渠道不可比，已如实跳过核对`
                    : `kb-llm：host 通道下拉 ${shown.length} 个清单项；与 HTTP 清单做包含性核对，${compared} 个有模型的项对上（未做等集合比较）` + (partial === null ? '' : '；RPC partial failures 在，只验无编造项'));
            }
        }
        // 宿主清单接口本身也读一次（只读 GET、不带 Key）：本轮不切通道，所以界面停在
        // 用户当前配置的 off；这条旁路证明「隔离实例里这个目录到底有没有模型」是如实的。
        const routeProbe = await evaluate(`fetch('/office-memory/llm-models', { headers: { accept: 'application/json', 'x-office-verify-probe': '1' } }).then((r) => r.json().then((b) => ({ status: r.status, body: b }))).catch((e) => ({ status: 0, fetchError: String(e && e.message || e) }))`);
        const probeShape = routeProbe === null || typeof routeProbe !== 'object' ? null : {
            status: routeProbe.status,
            ok: !!(routeProbe.body && routeProbe.body.ok === true),
            providers: routeProbe.body && Array.isArray(routeProbe.body.providers) ? routeProbe.body.providers.length : null,
            error: routeProbe.body && typeof routeProbe.body.error === 'string' ? routeProbe.body.error.slice(0, 80) : (routeProbe.fetchError ? String(routeProbe.fetchError).slice(0, 80) : null),
        };
        findings.dom.round95.routeProbe = probeShape;
        if (probeShape === null) fail('设置页：旁路读宿主清单接口失败（没有返回）');
        if (probeShape.ok === false && probeShape.error === null) fail('设置页：宿主清单接口返回 ok:false 却没给 error 原文');
        log(`kb-llm：宿主清单接口 HTTP ${probeShape.status}，${probeShape.ok ? '成功，' + probeShape.providers + ' 个 provider' : '如实报错：' + probeShape.error}`);
        await evaluate(`(() => { const root = ${MEMORY_ROOT}; const b = root ? root.querySelector('[data-group="kb-llm"]') : null; if (b && b.scrollIntoView) b.scrollIntoView({ block: 'center' }); return b !== null; })()`);
        await sleep(250);
        await shot('settings-kb-llm.png');

        ws.close();
    } finally {
        killTree(chrome);
        killTree(dshChild);
        // 失败也要留下现场：JSON 里记录已测得的事实与 findings.error —— 否则
        // .office/runtime/diag/web-verify.json 还停在上一次的成功结果，读的人会把旧绿
        // 当成今天通过（第七十七轮设置入口漂移时就是这样）。
        try {
            writeFileSync(path.join(OUT, 'web-verify.json'), JSON.stringify(findings, null, 2), 'utf8');
        } catch (error) { /* findings 还没建起来就算了 */ }
    }

    const outFile = path.join(OUT, 'web-verify.json');
    writeFileSync(outFile, JSON.stringify(findings, null, 2), 'utf8');
    log('\n===== 结论 =====');
    log(`bootOk（无 "did not activate" 横幅）：${findings.bootOk}`);
    log(`未捕获异常：${findings.exceptions.length} 条`); findings.exceptions.slice(0, 4).forEach((e) => log('   ! ' + String(e).split('\n')[0]));
    log(`控制台 error/warning：${findings.consoleErrors.length} 条`); findings.consoleErrors.slice(0, 4).forEach((e) => log('   ! ' + e.slice(0, 140)));
    log(`拦截到的记忆端点请求：${findings.intercepted} 次（Network 观察到 ${findings.snapshotRequests ?? 0} 次：${findings.snapshotUrl ?? '-'}）`);
    log(`会话页签行的「记忆」入口：${JSON.stringify(findings.dom.clickedMemory)}`);
    log(`开会话这一步：${JSON.stringify(findings.dom.sessionStep)}`);
    log(`会话页签行（上下文 / 编队 / 记忆 在这一排）：${JSON.stringify(findings.dom.conversationTabs)}`);
    log(`页签：${JSON.stringify(findings.dom.tabs)}`);
    log(`样式表节点：${JSON.stringify(findings.dom.styleNode)}`);
    log(`面板盒：${JSON.stringify(findings.dom.panelBox)}`);
    log(`容量条：${JSON.stringify(findings.dom.gauges)}`);
    log(`存储域 chips：${JSON.stringify(findings.dom.storeChips)}`);
    log(`令牌解析：${JSON.stringify(findings.dom.tokens)}`);
    log(`深色令牌：${JSON.stringify(findings.dom.darkTokens)}`);
    log(`宽屏栅格：${JSON.stringify(findings.dom.wideLayout)}`);
    log(`窄屏栅格：${JSON.stringify(findings.dom.narrowLayout)}`);
    log('');
    log('── 第十三轮：真实输入事件 ──');
    log(`分页（点之前）：${JSON.stringify(findings.dom.paginationBefore)}`);
    log(`分页（点之后）：${JSON.stringify(findings.dom.paginationAfter)}`);
    log(`展开长条目：${JSON.stringify(findings.dom.expandBefore)} → ${JSON.stringify(findings.dom.expandAfter)}`);
    log(`条目详情条（悬停）：${JSON.stringify(findings.dom.itemDetail)}`);
    log(`写入指令（点了「${findings.dom.writeClicked}」）：${JSON.stringify(findings.dom.writeInstruction)}`);
    log(`减少动效下的 scroll-behavior：${JSON.stringify(findings.dom.scrollBehavior)}`);
    log('── 第二十轮：滚动契约与设置页 ──');
    log(`面板滚动契约：${JSON.stringify(findings.dom.scrollContract)}`);
    log(`设置页：打开 ${findings.dom.settings && findings.dom.settings.opened} / 办公模式 ${findings.dom.settings && findings.dom.settings.officeNav}`);
    log(`设置页字号与控件：${JSON.stringify(findings.dom.settings && findings.dom.settings.controls)}`);
    log(`折叠块：${findings.dom.settings && findings.dom.settings.foldBefore} →（点 ${findings.dom.settings && findings.dom.settings.foldClick}）→ ${findings.dom.settings && findings.dom.settings.foldAfter}`);
    log(`折叠块里的控件：${JSON.stringify(findings.dom.settings && findings.dom.settings.foldBodyRows)}`);
    log('── 第三十五轮：分组折叠与定位 ──');
    const r35 = findings.dom.round35 || {};
    log(`默认态（每组：开合 / 组体 display / 控件数 / 可见控件数）：${JSON.stringify(r35.collapsed && r35.collapsed.groups)}`);
    log(`定位状态行：${JSON.stringify(r35.collapsed && r35.collapsed.stats)}`);
    log(`点开「检索编排」后：${JSON.stringify(r35.afterExpand && r35.afterExpand.groups)}`);
    log(`全部展开后：${JSON.stringify(r35.allExpanded && r35.allExpanded.groups)}`);
    log(`定位条 sticky：${JSON.stringify(r35.sticky)}`);
    log(`定位「出口代理」：${JSON.stringify(r35.located)}`);
    log(`定位折在 Fold 里的高级参数：${JSON.stringify(r35.locatedInFold)}`);
    log(`没有命中：${JSON.stringify(r35.noHit)}`);
    log(`清空 + 全部收起后：${JSON.stringify(r35.cleared && r35.cleared.groups)}（localStorage：${JSON.stringify(r35.cleared && r35.cleared.storage)}）`);
    log(`记忆系统页：${JSON.stringify(r35.memoryPage)}`);
    log(`kb-llm（第九十五轮）：通道 ${findings.dom.round95.mode}；UI 清单 HTTP 请求 ${findings.kbLlmRequests.length} 次（探针 GET 不计）；RPC ${findings.kbLlmRpcFrames ?? 0} 次（HTTP POST ${findings.kbLlmRpcHttp ?? 0} / WS ${findings.kbLlmRpcWs ?? 0}；0 HTTP 且走了 RPC 是正常路径，不是失败）；接口旁路 ${JSON.stringify(findings.dom.round95.routeProbe)}；初始形态 ${JSON.stringify(findings.dom.round95.initial)}`);
    log(`面板正文：\n${String(findings.dom.panelText).slice(0, 1200)}`);
    log(`详情写到 ${outFile}`);
}

main().catch((err) => { console.error('探针失败：', err); process.exit(1); });
