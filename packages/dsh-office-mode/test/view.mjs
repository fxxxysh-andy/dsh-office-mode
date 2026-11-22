/**
 * 记忆浏览端点的测试：只读路由 + 信任边界。
 *
 * 这一份测的重点不是「面板长什么样」，而是**端点不该给什么**：
 *   - 只接受 GET，其它方法一律 405；
 *   - 只服务「本进程真的跑过办公工具的工作目录」，未知路径 404 而不是照读；
 *   - 记忆关掉时不泄漏内容；
 *   - 快照内容与引擎口径一致（热记忆 / 台账 / 归档 / 关系 / 实体 / 计数）。
 *
 * 跑法：node test/view.mjs
 */
import { existsSync, readdirSync } from 'node:fs';
import { mkdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import { createMemory } from '../src/memory.js';
import { pdfEngines } from '../src/pdf.js';
import { buildSnapshot, EXPORT_VIEW_PATH, knownWorkspaces, noteWorkspace, registerMemoryView, resetWorkspaces, MEMORY_MUTATE_PATH, MEMORY_VIEW_PATH, MODEL_STATUS_PATH, MODEL_DOWNLOAD_PATH, FS_LIST_PATH, KB_UPLOAD_PATH, KB_ASSET_PATH, OCR_PATH, LLM_MODELS_PATH, GALLERY_PATH } from '../src/view.js';

// 测试中间产物统一落在仓库根的 .office/runtime/tmp/<suite>/ 下（不再用包内的 test/.tmp）。
const TMP = fileURLToPath(new URL('../../../.office/runtime/tmp/view/', import.meta.url));

const results = [];
async function check(name, fn) {
    try {
        // 用例可以 return 一句说明（比如「本机没装 PyMuPDF，只校验了如实报错」），
        // 打印时会跟在 PASS 后面，避免把没跑到的分支当成跑过了。
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

let caseId = 0;
async function freshRoot(memoryConfig = {}) {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    await rm(root, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    // 知识库是单库（第八十八轮）：夹具要把 kbDir 也指到本 case 自己的目录，
    // 否则所有 case 共用 globalDir 的兄弟 kb/，上一个用例的文档会漏到下一个。
    const memory = {
        globalDir: join(TMP, `global-${caseId}`),
        kbDir: join(TMP, `kb-${caseId}`),
        ...memoryConfig,
    };
    return { root, memory };
}

/** 造一个假的 req/res，把响应体收下来。 */
function makeRes() {
    return {
        status: 0,
        headers: null,
        body: '',
        chunks: [],
        flushed: false,
        writeHead(status, headers) {
            this.status = status;
            this.headers = headers;
        },
        flushHeaders() {
            this.flushed = true;
        },
        write(chunk) {
            this.chunks.push(String(chunk));
        },
        end(body) {
            if (body !== undefined && body !== null) {
                // 二进制图道路由直接把 Buffer 交给 end()，不能转成字符串。
                if (this.chunks.length === 0) {
                    this.body = body;
                    return;
                }
                this.chunks.push(String(body));
            }
            this.body = this.chunks.join('');
        },
    };
}

function makeReq({ method = 'POST', url = MEMORY_MUTATE_PATH, headers = {}, body = '' } = {}) {
    const handlers = new Map();
    return {
        method, url, headers,
        on(event, fn) {
            handlers.set(event, fn);
            // 与真实流同一顺序：先 data、再 end。
            setTimeout(() => {
                if (event === 'data') fn(Buffer.from(body));
                else if (event === 'end') fn();
            }, 0);
        },
        destroy() {},
    };
}

async function callRoute(handler, { method = 'GET', url = MEMORY_VIEW_PATH } = {}) {
    const res = makeRes();
    await handler({ method, url }, res);
    let json = null;
    try {
        json = JSON.parse(res.body);
    } catch {
        json = null;
    }
    return { status: res.status, headers: res.headers, json, raw: res.body };
}

/**
 * 一个插件上下文替身。
 *
 * `services` 里的东西按名字返回（webServer 之外还认 workspaceRegistry 与
 * sessionPersistence —— 第三十一轮起端点用它们当工作目录的信任来源）；
 * 没给的按「宿主没有这个服务」处理，返回 undefined。
 */
function fakeCtx(services = {}) {
    const routes = [];
    return {
        routes,
        get(name) {
            if (name === 'webServer') {
                return {
                    register(route) {
                        routes.push(route);
                        return () => {
                            const at = routes.indexOf(route);
                            if (at >= 0) routes.splice(at, 1);
                        };
                    },
                };
            }
            return services[name];
        },
    };
}

/** 宿主工作区登记表的替身：`list()` 返回 [{path, title}]。 */
function fakeRegistry(items) {
    return { list: () => items };
}

/** 会话存储的替身：`list()` 返回 [{header: {cwd, createdAt}}]。 */
function fakeSessions(items) {
    return { list: async () => items };
}

await rm(TMP, { recursive: true, force: true });

// ── 快照内容 ────────────────────────────────────────────────────────────────

await check('快照：三层内容、条目实体与计数都按引擎口径给出来（图谱字段已下线）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    const a = await instance.mutate({ action: 'add', target: 'user', content: '用户偏好：汇报用 16:9。', entities: ['汇报模板'] });
    const b = await instance.mutate({ action: 'add', target: 'project', content: '项目约定：主题用 business。' });
    await instance.log([{ path: '季度汇报.pptx', format: 'ppt', theme: 'business', purpose: '季度主线' }]);
    await instance.appendArchiveBlock({ reason: 'mnemon-migration', items: [{ id: 'i-1', kind: 'mnemon', content: '迁来的长期记忆', entities: ['实体甲'] }] });

    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.ok, true);
    assert.equal(snap.cwd, root);
    assert.equal(snap.label, root.split(/[\\/]/).filter(Boolean).pop());
    assert.equal(snap.hot.length, 2);
    assert.equal(snap.ledger.length, 1);
    assert.ok(snap.archive.length >= 1);
    assert.equal(snap.links, undefined, '关系字段已从快照删除');
    assert.equal(snap.entities, undefined, '实体视图字段已从快照删除');
    assert.equal(snap.counts.hot, 2);
    assert.equal(snap.counts.links, undefined);

    const user = snap.hot.find((item) => item.target === 'user');
    assert.deepEqual(user.entities, ['汇报模板'], '实体要随条目出来');
    assert.equal(user.origin, 'workspace');

    assert.deepEqual(snap.hot[0].entities, ['汇报模板'], '条目自带的 entities 字段仍保留（不是图谱实体视图）');

    // 配置回显：面板要按这些值解释自己的显示。
    assert.equal(snap.config.scope, 'workspace');
    assert.equal(snap.config.quality.policy, 'strict-v1');
    assert.equal(snap.config.layers.hot, true);
    assert.deepEqual(snap.stores.map((store) => store.id), ['workspace']);

    // 体积字段（冻结契约）：面板要显示每个根的占用与两层正文的体积。
    assert.equal(snap.stores.length, 1);
    const store = snap.stores[0];
    assert.equal(typeof store.bytes, 'number', 'stores[].bytes 必须是数字');
    assert.equal(typeof store.files, 'number', 'stores[].files 必须是数字');
    assert.ok(store.files >= 3, `至少有 memory.json / USER.md / ledger.jsonl，实际 ${store.files}`);
    assert.ok(store.bytes > 0, `有内容的记忆库体积应大于 0，实际 ${store.bytes}`);

    const userBytes = Buffer.byteLength('用户偏好：汇报用 16:9。', 'utf8') + Buffer.byteLength('\n§\n', 'utf8');
    const projectBytes = Buffer.byteLength('项目约定：主题用 business。', 'utf8') + Buffer.byteLength('\n§\n', 'utf8');
    assert.equal(snap.counts.hotBytes, userBytes, 'counts.hotBytes 只算 target:user 的正文');
    assert.equal(snap.counts.projectBytes, projectBytes, 'counts.projectBytes 只算 target:project 的正文');

    // 容量口径与「列表里有几条」是两个量（面板的容量条用前者）：
    //   ledgerTotal  —— 未过滤、未被浏览上限截断的台账总数；
    //   archiveFiles —— 归档**摘要文件**数（archiveKeep 管的是它，一个月一个）；
    //   archiveItems —— 归档条目数（与 archive.length 同源，但不受过滤影响）。
    assert.equal(snap.counts.ledgerTotal, 1, 'ledgerTotal 是台账总数');
    assert.equal(snap.counts.archiveFiles, 1, '一条 appendArchiveBlock 只产生 1 个摘要文件');
    assert.equal(snap.counts.archiveItems, 1, 'archiveItems 是归档条目数');
    assert.equal(snap.stores[0].archiveFiles, 1, 'stores[] 也要给摘要文件数');
});

await check('快照：容量口径不跟着搜索框走（q 只影响列表条数）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.log([
        { path: '季度汇报.pptx', format: 'ppt' },
        { path: '年会策划.pptx', format: 'ppt' },
        { path: '无关的.docx', format: 'word' },
    ]);
    await instance.appendArchiveBlock({ reason: 'mnemon-migration', items: [
        { id: 'i-1', kind: 'mnemon', content: '归档里的季度结论' },
        { id: 'i-2', kind: 'mnemon', content: '另一条不相干的旧记录' },
    ] });

    const all = await buildSnapshot({ root, memory });
    assert.equal(all.counts.ledger, 3);
    assert.equal(all.counts.ledgerTotal, 3);
    assert.equal(all.counts.archive, 2);
    assert.equal(all.counts.archiveFiles, 1);

    const filtered = await buildSnapshot({ root, memory, query: '季度' });
    assert.equal(filtered.counts.ledger, 1, '列表条数跟着过滤走（这是对的：页签与指标卡要它）');
    assert.equal(filtered.counts.archive, 1);
    assert.equal(filtered.counts.ledgerTotal, 3, '台账容量口径不跟着过滤缩水');
    assert.equal(filtered.counts.archiveItems, 2, '归档条目数不跟着过滤缩水');
    assert.equal(filtered.counts.archiveFiles, 1, '摘要文件数更不该跟着过滤变');
});

await check('快照：空目录的体积字段是 0，不是 undefined', async () => {
    const { root, memory } = await freshRoot();
    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.ok, true);
    assert.equal(snap.hot.length, 0);
    assert.equal(snap.stores.length, 1);
    assert.equal(snap.stores[0].bytes, 0, '目录还没建出来时体积是 0');
    assert.equal(snap.stores[0].files, 0, '目录还没建出来时文件数是 0');
    assert.equal(snap.counts.hotBytes, 0, '没有条目时 hotBytes 是 0 而不是 undefined');
    assert.equal(snap.counts.projectBytes, 0, '没有条目时 projectBytes 是 0 而不是 undefined');
});

await check('快照：知识库文档清单（kb 一期）与计数、逐根计数都给出来', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    // 入库评估闸（P1-4b）：`tier:'verified'` 的文档要自带 ≥2 个独立来源。
    await writeFile(join(root, '来源笔记.md'), '# 方法论\n\n真实 Office 才算证据。\n\n## 第二段\n\n复核要能复跑。\n\n来源：https://news-a.example.com/report、https://news-b.example.org/report\n', 'utf8');
    const ingested = await instance.kbIngest({ path: '来源笔记.md', tier: 'verified' });
    assert.equal(ingested.chunks > 0, true, '夹具要先真的切出块');

    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.counts.kb, 1, '知识库页签的条数 = 清单里有几篇');
    assert.equal(snap.counts.kbTotal, 1, 'kbTotal 是未过滤总数');
    assert.equal(snap.counts.kbChunks, ingested.chunks, 'kbChunks 是块总数');
    assert.equal(snap.counts.kbBytes > 0, true, 'kbBytes 是文档总字节');
    assert.equal(snap.counts.kbTruncated, false, '没到单次上限就不算截断');
    assert.equal(snap.kb.length, 1);

    const doc = snap.kb[0];
    assert.equal(doc.path, '来源笔记.md');
    assert.equal(doc.tier, 'verified');
    assert.equal(doc.chunks, ingested.chunks);
    assert.equal(doc.title.includes('方法论'), true, '标题从正文一级标题取');
    assert.equal(doc.id.startsWith('kb:'), true, '文档 id 是 kb: 形');
    assert.equal(doc.origin, 'kb', '知识库是单库，origin 固定为 kb');
    assert.equal(typeof doc.at, 'string');

    // 单库计数：面板的存储卡按 sb.kbStore 画「知识库 N 篇 / M 块」。
    assert.equal(snap.kbStore.docs, 1, 'kbStore.docs 要有');
    assert.equal(snap.kbStore.chunks, ingested.chunks, 'kbStore.chunks 要有');
    assert.equal(snap.kbStore.tiers.verified, 1, '档位分布要有');

    // 过滤：与其它层同口径（按路径 / 标题 / 档位匹配）。
    const hit = await buildSnapshot({ root, memory, query: '方法论' });
    assert.equal(hit.counts.kb, 1, 'kb 也参与 q 过滤');
    const miss = await buildSnapshot({ root, memory, query: '不存在的词' });
    assert.equal(miss.counts.kb, 0, '过滤后列表条数跟着变');
    assert.equal(miss.counts.kbTotal, 1, 'kbTotal 不跟着过滤缩水');
});

await check('快照：知识库块检索与块清单可下钻（第八十二轮）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await writeFile(join(root, '来源笔记.md'), '# 方法论\n\n真实 Office 才算证据。\n\n## 第二段\n\n复核要能复跑。\n', 'utf8');
    const ingested = await instance.kbIngest({ path: '来源笔记.md', tier: 'user' });
    assert.ok(ingested.chunks > 0, '夹具要先真的切出块');

    const searched = await buildSnapshot({ root, memory, kbQuery: '复核' });
    assert.equal(searched.ok, true);
    assert.ok(searched.kbSearch !== null, '点名 kbq 时要有检索结果');
    assert.equal(searched.kbSearch.hits.length > 0, true, '「复核」应命中正文');
    const hit = searched.kbSearch.hits[0];
    assert.equal(hit.path, '来源笔记.md');
    assert.equal(hit.docId.startsWith('kb:'), true, '命中要带文档 id');
    assert.equal(hit.chunkId.startsWith('kb:'), true, '命中要带块 id');
    assert.match(hit.preview, /复核/, '预览要是真正文的前缀');

    const doc = await buildSnapshot({ root, memory, kbDoc: ingested.docId });
    assert.equal(doc.ok, true);
    assert.ok(doc.kbDoc !== null, '点名 kbdoc 时要有块清单');
    assert.equal(doc.kbDoc.docId, ingested.docId);
    assert.equal(doc.kbDoc.chunks.length, ingested.chunks);
    assert.equal(doc.kbDoc.chunks_total, ingested.chunks);
    assert.equal(doc.kbDoc.chunks[0].id.startsWith('kb:'), true);
    assert.equal(typeof doc.kbDoc.chunks[0].chars, 'number');
    assert.match(doc.kbDoc.chunks[0].spanText, /^字符 \d+–\d+$/);

    const chunk = await buildSnapshot({ root, memory, kbChunk: doc.kbDoc.chunks[0].id });
    assert.equal(chunk.ok, true);
    assert.ok(chunk.kbChunk !== null, '点名 kbchunk 时要有块正文');
    assert.equal(chunk.kbChunk.id, doc.kbDoc.chunks[0].id);
    assert.equal(chunk.kbChunk.docId, ingested.docId);
    assert.ok(chunk.kbChunk.text.length > 0, '块正文不能是空的');
});

await check('快照：知识库文件夹树与当前文件夹范围（第八十四轮）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await writeFile(join(root, '合同.md'), '# 合同\n\n付款周期 30 天。\n', 'utf8');
    await writeFile(join(root, '制度.md'), '# 制度\n\n报销要两张凭证。\n', 'utf8');
    await writeFile(join(root, '根文档.md'), '# 根文档\n\n放在根上的一份说明。\n', 'utf8');
    await instance.kbIngest({ path: '合同.md', folder: '参考/合同' });
    await instance.kbIngest({ path: '制度.md', folder: '参考/制度' });
    await instance.kbIngest({ path: '根文档.md' });

    const rootSnap = await buildSnapshot({ root, memory });
    assert.equal(rootSnap.kbFolder, '');
    assert.equal(rootSnap.kbRecursive, false, '默认只列当前层');
    assert.equal(rootSnap.counts.kb, 1, '根目录只列直属文档');
    assert.equal(rootSnap.kbTree.docs, 3, '根节点计数含全部子树');
    const ref = rootSnap.kbTree.children.find((node) => node.path === '参考');
    assert.ok(ref, '树里要有「参考」这一层');
    assert.equal(ref.docs, 2);
    assert.deepEqual(ref.children.map((node) => node.name).sort(), ['制度', '合同'].sort());

    const scoped = await buildSnapshot({ root, memory, kbFolder: '参考' });
    assert.equal(scoped.kbFolder, '参考');
    assert.equal(scoped.kbRecursive, false, '默认只列当前层');
    assert.equal(scoped.counts.kb, 0, '「参考」下没有直接文档，子目录默认不展开');
    assert.equal(scoped.counts.kbTotal, 3, 'kbTotal 是整库总数，不跟当前层缩水');
    const subtree = await buildSnapshot({ root, memory, kbFolder: '参考', kbRecursive: true });
    assert.equal(subtree.counts.kb, 2, '显式递归才列子树');
    assert.equal(subtree.counts.kbTotal, 3);
    assert.ok(subtree.kb.every((doc) => doc.folder.startsWith('参考/')));
    // 左树仍是全量：换层不该把树自己剪掉。
    assert.equal(scoped.kbTree.docs, 3);

    const shallow = await buildSnapshot({ root, memory, kbFolder: '参考', kbRecursive: false });
    assert.equal(shallow.counts.kb, 0, '仅当前层时「参考」下没有直接文档');

    const searched = await buildSnapshot({ root, memory, kbFolder: '参考', kbQuery: '凭证' });
    assert.equal(searched.kbSearch.hits.length, 0, '默认不递归：当前层没有块');
    assert.equal(searched.kbSearch.folder, '参考');
    const searchedDeep = await buildSnapshot({ root, memory, kbFolder: '参考', kbQuery: '凭证', kbRecursive: true });
    assert.equal(searchedDeep.kbSearch.hits.length, 1, '显式递归才检索子树');
    assert.equal((await buildSnapshot({ root, memory, kbQuery: '凭证' })).kbSearch.hits.length, 0, '根目录默认只查直接文档');
    const miss = await buildSnapshot({ root, memory, kbFolder: '参考', kbQuery: '根上' });
    assert.equal(miss.kbSearch.hits.length, 0, '检索范围跟着当前文件夹收窄');
});

await check('快照：来源筛选把全局与工作区分开，计数仍是全量（第八十二轮）', async () => {
    const { root, memory } = await freshRoot({ scope: 'both', userScope: 'global' });
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'user', content: '全局偏好：先给结论。' });
    await instance.mutate({ action: 'add', target: 'project', content: '工作区约定：稿子放 docs。' });
    await instance.log([{ path: '本地台账.pptx', format: 'ppt' }]);

    const all = await buildSnapshot({ root, memory });
    assert.equal(all.hot.length, 2);
    assert.deepEqual(all.originCounts.hot, { all: 2, global: 1, workspace: 1 });
    assert.equal(all.origin, '', '不带 origin 时是全部');

    const global = await buildSnapshot({ root, memory, origin: 'global' });
    assert.equal(global.hot.length, 1, '只留全局条目');
    assert.equal(global.hot[0].origin, 'global');
    assert.equal(global.counts.hot, 1, '页签计数跟着筛选走');
    assert.deepEqual(global.originCounts.hot, { all: 2, global: 1, workspace: 1 }, '三档计数不随筛选缩水');
    assert.equal(global.ledger.length, 0, '台账是工作区的，全局视图里没有');

    const workspace = await buildSnapshot({ root, memory, origin: 'workspace' });
    assert.equal(workspace.hot.length, 1);
    assert.equal(workspace.hot[0].origin, 'workspace');
    assert.equal(workspace.ledger.length, 1);
});

await check('快照：没有知识库时 kb 是空数组、计数为 0（不是 undefined）', async () => {
    const { root, memory } = await freshRoot();
    const snap = await buildSnapshot({ root, memory });
    assert.deepEqual(snap.kb, [], '.office/kb 不存在时给空数组');
    assert.equal(snap.counts.kb, 0);
    assert.equal(snap.counts.kbTotal, 0);
    assert.equal(snap.counts.kbChunks, 0);
    assert.equal(snap.counts.kbBytes, 0);
    assert.equal(snap.kbStore.docs, 0, 'kbStore.docs 也要给 0 而不是 null');
});

await check('快照：知识库清单是服务端真分页（窗口 / 越界回拉 / 全局排序）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    // 清单接口的硬上限是 20 篇（KB_LIST_LIMIT）。端点声明的上限必须与它一致，
    // 否则面板会以为「要更多却只回来 20 篇」是丢了东西。
    for (let index = 0; index < 23; index += 1) {
        const rel = `来源/第${String(index).padStart(2, '0')}篇.md`;
        const absolute = join(root, '来源', `第${String(index).padStart(2, '0')}篇.md`);
        await mkdir(join(root, '来源'), { recursive: true });
        await writeFile(absolute, `# 第 ${index} 篇\n\n这是第 ${index} 篇的正文，用来把清单推到上限之外。\n`, 'utf8');
        await instance.kbIngest({ path: rel, tier: 'user' });
    }

    // 真分页（第九十四轮）：默认第一页 20 篇，kbPage 报窗口，kbTruncated 只表示「还有下一页」。
    const snap = await buildSnapshot({ root, memory });
    assert.equal(snap.counts.kbTotal, 23, 'kbTotal 是真实总数');
    assert.equal(snap.kb.length, 20, '第一页给了 20 篇');
    assert.deepEqual(snap.kbPage, { offset: 0, limit: 20, total: 23, pageCount: 2, page: 1, hasPrev: false, hasNext: true });
    assert.equal(snap.counts.kb, 23, '页签条数是过滤后的总数，不是本页条数');
    assert.equal(snap.counts.kbTruncated, true, '还有下一页');
    assert.deepEqual(snap.kbSort, { key: 'name', dir: 'asc' }, '默认按名称升序');
    assert.equal(snap.kbStore.docs, 23, '单库计数仍是真实总数');
    assert.equal(snap.kbStore.chunks, snap.counts.kbChunks, '单库与合计同源（都从 manifest 现算）');
    assert.equal(snap.kbStore.bytes, snap.counts.kbBytes, '单库与合计同源');

    // 第二页：3 篇、不重不漏、hasNext=false。
    const page2 = await buildSnapshot({ root, memory, kbOffset: 20 });
    assert.equal(page2.kb.length, 3, '第二页 3 篇');
    assert.deepEqual(page2.kbPage, { offset: 20, limit: 20, total: 23, pageCount: 2, page: 2, hasPrev: true, hasNext: false });
    const ids = new Set([...snap.kb, ...page2.kb].map((row) => row.id));
    assert.equal(ids.size, 23, '两页拼起来不重不漏');
    assert.deepEqual(page2.kb.map((row) => row.path), ['来源/第20篇.md', '来源/第21篇.md', '来源/第22篇.md'], '跨页顺序连续');

    // 越界 offset 回拉到末页行首（第 2 页删掉唯一一条时的场景），并把生效 offset 回报给客户端。
    const past = await buildSnapshot({ root, memory, kbOffset: 999 });
    assert.equal(past.kbPage.offset, 20, '越界 offset 回拉到末页行首');
    assert.equal(past.kbPage.page, 2);
    assert.equal(past.kb.length, 3, '回拉后仍有内容，不是空页');

    // 排序是全局稳定序：翻完所有页拼起来 = 一次要全量的顺序。
    const full = await buildSnapshot({ root, memory, kbLimit: 100 });
    assert.equal(full.kbPage.limit, 100, 'kbLimit 在白名单内才生效');
    assert.deepEqual(full.kb.map((row) => row.id), [...snap.kb, ...page2.kb].map((row) => row.id), '分页与全量顺序一致');
    const reversed = await buildSnapshot({ root, memory, kbLimit: 100, kbSortDir: 'desc' });
    assert.deepEqual(reversed.kb.map((row) => row.id), full.kb.map((row) => row.id).reverse(), 'desc 是同一序的反向');
    const bogus = await buildSnapshot({ root, memory, kbLimit: 100, kbSort: 'hack', kbSortDir: 'sideways' });
    assert.deepEqual(bogus.kbSort, { key: 'name', dir: 'asc' }, '拼错的排序回落到默认而不是空序');

    // 文件夹树与合计不随当前页变：左侧导航不能用「这一页的行」算。
    assert.deepEqual(page2.kbTree, snap.kbTree, 'kbTree 不随 offset 变');
    assert.equal(page2.kbStore.docs, 23, 'kbStore 不随 offset 变');

    // **先过滤、再排序、再分页**：搜一篇落在第一页之外的文档，必须搜得到。
    // 早期实现是先把 20 篇截出来再过滤，于是这一条会得到 0 条（复核 P2-1）。
    // 注意查询词不能带空格：带空格会被拆成「与」关系的多个词，`第` / `篇` 命中全部 23 篇。
    const oldest = await buildSnapshot({ root, memory, query: '第00篇' });
    assert.equal(oldest.counts.kb, 1, '搜索要覆盖全部文档，不是只覆盖第一页');
    assert.equal(oldest.kb.length, 1);
    assert.equal(oldest.kb[0].path, '来源/第00篇.md');
    assert.equal(oldest.counts.kbTotal, 23, 'kbTotal 不跟着过滤缩水');
    assert.equal(oldest.kbPage.total, 1, '分页总量按过滤后算');
});

await check('快照：摘要块真分页（总数 / 窗口 / 越界回拉）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    const rel = '摘要/分页.md';
    await mkdir(join(root, '摘要'), { recursive: true });
    await writeFile(join(root, rel), [
        '# 分页测试', '', '第一段正文，用来切第一块。', '', '## 第二段', '', '第二段正文，用来切第二块。', '', '## 第三段', '', '第三段正文，用来切第三块。',
    ].join('\n'), 'utf8');
    const doc = await instance.kbIngest({ path: rel });
    const read = await instance.kbRead({ id: doc.docId });
    assert.ok(read.chunks.length >= 3);
    for (let index = 0; index < 3; index += 1) {
        await instance.kbSummarize({ title: '摘要 ' + index, text: '第 ' + index + ' 条摘要正文。', chunkIds: [read.chunks[index].id] });
    }
    assert.equal(await instance.kbSummariesCount(), 3);

    const page1 = await buildSnapshot({ root, memory, kbSummaries: true, kbSummariesOffset: 0, kbSummariesLimit: 2 });
    assert.equal(page1.kbSummaries.length, 2);
    assert.deepEqual(page1.kbSummariesPage, { offset: 0, limit: 2, total: 3, pageCount: 2, page: 1, hasPrev: false, hasNext: true });
    const page2 = await buildSnapshot({ root, memory, kbSummaries: true, kbSummariesOffset: 2, kbSummariesLimit: 2 });
    assert.equal(page2.kbSummaries.length, 1);
    assert.deepEqual(page2.kbSummariesPage, { offset: 2, limit: 2, total: 3, pageCount: 2, page: 2, hasPrev: true, hasNext: false });
    const ids = new Set([...page1.kbSummaries, ...page2.kbSummaries].map((row) => row.id));
    assert.equal(ids.size, 3, '两页拼起来不重不漏');

    const past = await buildSnapshot({ root, memory, kbSummaries: true, kbSummariesOffset: 999, kbSummariesLimit: 2 });
    assert.equal(past.kbSummariesPage.offset, 2, '越界 offset 回拉到末页行首');
    assert.equal(past.kbSummaries.length, 1);
});

await check('快照：q 过滤覆盖四段（热记忆 / 台账 / 归档 / 知识库）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    const a = await instance.mutate({ action: 'add', target: 'project', content: '热记忆里的一条：季度汇报用 16:9。' });
    const b = await instance.mutate({ action: 'add', target: 'project', content: '另一条不相干的项目约定。' });
    await instance.log([
        { path: '季度汇报.pptx', format: 'ppt' },
        { path: '年会策划.pptx', format: 'ppt' },
    ]);
    await instance.appendArchiveBlock({ reason: 'mnemon-migration', items: [{ id: 'i-1', kind: 'mnemon', content: '归档里的季度结论' }] });

    const snap = await buildSnapshot({ root, memory, query: '季度' });
    assert.equal(snap.query, '季度');
    assert.equal(snap.hot.length, 1, '热记忆也要参与过滤');
    assert.match(snap.hot[0].content, /季度/);
    assert.equal(snap.ledger.length, 1);
    assert.equal(snap.ledger[0].path, '季度汇报.pptx');
    assert.equal(snap.archive.length, 1, '归档也要参与过滤');

    const all = await buildSnapshot({ root, memory });
    assert.equal(all.hot.length, 2, '不带 q 时不过滤');
    assert.equal(all.ledger.length, 2);
});

// ── 路由与信任边界 ──────────────────────────────────────────────────────────

await check('路由：组合里没有 webServer 时安静地不注册', async () => {
    const dispose = registerMemoryView({ get: () => undefined }, { getMemory: () => ({}) });
    assert.equal(dispose, undefined);
});

await check('路由：注册快照 / 导出 / 写入 / 模型两条，且都可释放', async () => {
    const ctx = fakeCtx();
    const dispose = registerMemoryView(ctx, { getMemory: () => ({}) });
    // 第七十一轮起是两条（快照 / 导出）；第八十五轮加入写入口（面板的就地编辑）；
    // 第八十八轮加入模型状态与下载两条；第八十九轮加入目录浏览、知识库上传与图片资产三条；
    // 第九十二轮加入宿主模型列表与图库两条。
    assert.equal(ctx.routes.length, 11);
    assert.deepEqual(ctx.routes.map((route) => route.path), [MEMORY_VIEW_PATH, EXPORT_VIEW_PATH, MEMORY_MUTATE_PATH, MODEL_STATUS_PATH, MODEL_DOWNLOAD_PATH, FS_LIST_PATH, KB_UPLOAD_PATH, LLM_MODELS_PATH, KB_ASSET_PATH, OCR_PATH, GALLERY_PATH]);
    assert.ok(ctx.routes.every((route) => route.kind === 'exact'));
    assert.equal(typeof dispose, 'function');
    dispose();
    assert.equal(ctx.routes.length, 0, '释放后十一条路由都应被摘掉');
});

await check('目录浏览路由：只列目录、绝对路径、可回到上一级（第八十九轮）', async () => {
    const { root } = await freshRoot();
    await mkdir(join(root, 'A 目录'), { recursive: true });
    await mkdir(join(root, 'B 目录'), { recursive: true });
    await writeFile(join(root, 'note.md'), '文件不该出现在目录选择里。', 'utf8');
    noteWorkspace(root);
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({}) });
    const route = ctx.routes.find((item) => item.path === FS_LIST_PATH);
    assert.ok(route, '目录浏览要注册');
    const res = makeRes();
    await route.handler({ method: 'GET', url: FS_LIST_PATH + '?path=' + encodeURIComponent(root) }, res);
    assert.equal(res.status, 200);
    const body = JSON.parse(res.body);
    assert.deepEqual(body.entries.map((item) => item.name), ['A 目录', 'B 目录', 'note.md'], '目录在前、文件在后');
    assert.equal(body.entries[2].kind, 'file', '文件要标 kind:file，资源管理器才画得对');
    assert.ok(body.entries[2].size >= 0, '文件要带字节数');
    assert.equal(body.path, root);
    assert.ok(body.roots.length >= 1, '要给出盘符 / 根目录，方便从任意位置开始');
    assert.equal(body.parent, join(root, '..') === root ? '' : body.parent, '父级要能回到上一级');
    assert.ok(body.parent !== '', '非根目录要有上一级');

    const nested = makeRes();
    await route.handler({ method: 'GET', url: FS_LIST_PATH + '?path=' + encodeURIComponent(join(root, 'A 目录')) }, nested);
    assert.deepEqual(JSON.parse(nested.body).entries, [], '空目录返回空清单');

    const relative = makeRes();
    await route.handler({ method: 'GET', url: FS_LIST_PATH + '?path=' + encodeURIComponent('相对目录') }, relative);
    assert.equal(relative.status, 400, '相对路径要拒绝');
});

await check('知识库上传路由：文件进来只留切片与图片，原文不留（第八十九轮）', async () => {
    const { root } = await freshRoot();
    noteWorkspace(root);
    const memoryConfig = { kbDir: join(root, 'kb') };
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memoryConfig });
    const route = ctx.routes.find((item) => item.path === KB_UPLOAD_PATH);
    assert.ok(route, '上传入口要注册');
    const text = '上传的知识库文档。切片以后只保留正文块，不复制源文件。';
    const uploadTmpBefore = readdirSync(tmpdir()).filter((name) => name.startsWith('dsh-kb-upload-'));
    const res = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ name: '笔记.md', folder: '参考资料', data: Buffer.from(text, 'utf8').toString('base64') }),
    }), res);
    assert.equal(res.status, 200, res.body);
    const body = JSON.parse(res.body);
    assert.equal(body.ok, true);
    assert.equal(body.result.path, 'uploads/参考资料/笔记.md', 'manifest 记的是虚拟文件夹里的逻辑路径');
    assert.equal(body.result.path.includes('dsh-kb-upload'), false, '临时目录名不能漏进 manifest');
    assert.ok(body.result.chunks > 0, '要切出块');
    assert.equal(existsSync(join(root, 'uploads', '参考资料', '笔记.md')), false, '原文不能复制进工作区');
    assert.equal(existsSync(join(root, '.office', 'runtime', 'kb-upload')), false,
        '入库临时文件不落在工作区（私有 OS 临时目录）');
    const uploadTmpAfter = readdirSync(tmpdir()).filter((name) => name.startsWith('dsh-kb-upload-'));
    assert.deepEqual(uploadTmpAfter, uploadTmpBefore, '入库结束后私有临时目录要删掉');

    const memory = createMemory({ root, memory: memoryConfig });
    const list = await memory.kbList({ folder: '参考资料' });
    assert.ok(list.items.some((item) => item.path === 'uploads/参考资料/笔记.md'), '清单里能看到上传的文档');
    const hit = await memory.kbSearch({ query: '切片 正文块' });
    assert.ok(hit.hits.some((item) => item.path === 'uploads/参考资料/笔记.md'), '上传后能被检索到');

    // 内部临时根只认 OS 临时目录：即便有人给 kbIngest 塞一个绝对路径，也不能拿它当抽取源。
    await assert.rejects(
        () => memory.kbIngest({ path: 'x.txt', displayPath: 'uploads/x.txt', sourceRoot: join(root, 'kb') }),
        /OS 临时目录/,
        '内部临时根不能指向任意绝对路径',
    );

    // 坏文件：错误里只能是用户给的文件名与可读原因，不能出现 randomUUID 临时名或私有
    // 临时目录（这两种信息对用户没意义，还会泄漏本机路径）。
    const badRes = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ name: '坏文件.pdf', data: Buffer.from('这不是一份 PDF', 'utf8').toString('base64') }),
    }), badRes);
    const badUpload = JSON.parse(badRes.body);
    assert.equal(badUpload.ok, false, '坏 PDF 要失败');
    assert.ok(String(badUpload.error).includes('坏文件.pdf'),
        '错误要保留用户看得懂的文件名与原因：' + String(badUpload.error));
    assert.equal(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}/i.test(String(badUpload.error)), false,
        '不能出现临时 UUID：' + String(badUpload.error));
    assert.equal(String(badUpload.error).includes('dsh-kb-upload'), false,
        '不能出现私有临时目录名：' + String(badUpload.error));
});

await check('知识库上传路由：stream=true 先回 NDJSON 进度，再给最终结果（第一百一十轮）', async () => {
    const { root } = await freshRoot();
    noteWorkspace(root);
    const memoryConfig = { kbDir: join(root, 'kb') };
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memoryConfig });
    const route = ctx.routes.find((item) => item.path === KB_UPLOAD_PATH);
    assert.ok(route, '上传入口要注册');
    const text = '流式进度：这一段正文会先抽取、再切片，最后一次性回结果。';
    const res = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
            name: '进度.md',
            folder: '进度',
            data: Buffer.from(text, 'utf8').toString('base64'),
            stream: true,
        }),
    }), res);
    assert.equal(res.status, 200, res.body);
    assert.match(String(res.headers?.['content-type'] ?? ''), /application\/x-ndjson/);
    assert.equal(res.flushed, true, '流式响应要尽早把头刷出去');
    const events = res.body.trim().split('\n').filter(Boolean).map((line) => JSON.parse(line));
    assert.ok(events.length >= 3, '至少要有抽取、切片与结果三个事件：' + res.body);
    assert.equal(events[0].type, 'progress');
    assert.equal(events[0].phase, 'extract');
    assert.ok(events.some((event) => event.type === 'progress' && event.phase === 'chunks'),
        '切片阶段也要有进度');
    const last = events[events.length - 1];
    assert.equal(last.type, 'result');
    assert.equal(last.ok, true);
    assert.equal(last.action, 'kb-upload');
    assert.ok(last.result.chunks > 0, '结果里仍要带切片数');
    assert.equal(res.body.includes('dsh-kb-upload'), false, '进度流不能漏出临时目录名');
});

await check('知识库上传路由：正文通道（text + ocr）写回源文件文档，并收掉旧 .ocr.md 笔记', async () => {
    const { root } = await freshRoot();
    noteWorkspace(root);
    const memoryConfig = { kbDir: join(root, 'kb') };
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memoryConfig });
    const route = ctx.routes.find((item) => item.path === KB_UPLOAD_PATH);
    // ① 先用二进制通道造出「识别前」的图片文档与第 91 轮那种旁挂笔记。
    const imageRes = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ name: '扫描件.png', folder: '资料', data: Buffer.from('png-bytes', 'utf8').toString('base64') }),
    }), imageRes);
    assert.equal(JSON.parse(imageRes.body).ok, true, imageRes.body);
    const noteRes = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ name: '扫描件.ocr.md', folder: '资料', data: Buffer.from('旧的识别正文。', 'utf8').toString('base64') }),
    }), noteRes);
    assert.equal(JSON.parse(noteRes.body).ok, true, noteRes.body);
    const memory = createMemory({ root, memory: memoryConfig });
    assert.equal((await memory.kbList({ folder: '资料' })).total, 2, '识别前：图片文档 + 旁挂笔记');

    // ② OCR 正文通道：不重传二进制，只交正文；落点仍是源文件自己的文档。
    const res = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ name: '扫描件.png', folder: '资料', text: Buffer.from('新的识别正文，讲的是营收。', 'utf8').toString('base64'), ocr: true }),
    }), res);
    assert.equal(res.status, 200, res.body);
    const body = JSON.parse(res.body);
    assert.equal(body.ok, true, res.body);
    assert.equal(body.result.ocr, true);
    assert.equal(body.result.kind, 'ocr', '正文通道的 kind 要标成 ocr');
    assert.equal(body.result.path, 'uploads/资料/扫描件.png');
    assert.equal(body.result.removedOcrNote, 'uploads/资料/扫描件.ocr.md', '回执要点名被收掉的旧笔记');
    const list = await memory.kbList({ folder: '资料' });
    assert.equal(list.total, 1, '一份文件只留一份正文');
    assert.equal(list.items[0].path, 'uploads/资料/扫描件.png');
    assert.ok(list.items[0].ocr && list.items[0].ocr.chars > 0, '清单里要能看出这是识别结果');
    const read = await memory.kbRead({ path: 'uploads/资料/扫描件.png' });
    assert.ok(read.chunks.some((chunk) => chunk.text.includes('新的识别正文')));
    await assert.rejects(() => memory.kbRead({ path: 'uploads/资料/扫描件.ocr.md' }), /kb 里没有/);

    // ③ 空正文与「正文 + preview」都要拒绝：预览只能看二进制，正文已经是识别后的结果。
    const emptyRes = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ name: '扫描件.png', folder: '资料', text: Buffer.from('   ', 'utf8').toString('base64'), ocr: true }),
    }), emptyRes);
    assert.equal(emptyRes.status, 400, emptyRes.body);
    const previewTextRes = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ preview: true, name: '扫描件.png', text: Buffer.from('新的识别正文', 'utf8').toString('base64') }),
    }), previewTextRes);
    assert.equal(previewTextRes.status, 400, previewTextRes.body);
    assert.equal((await memory.kbList({ folder: '资料' })).total, 1, '被拒的请求不该动库');
});


/** 造一份每页带大段 ASCII 文本的最小 PDF（用于触发 8 MB 拆分路径）。 */
function makeBigPdf(pageTexts) {
    const pageCount = pageTexts.length;
    const fontId = 3 + pageCount;
    const contentStart = fontId + 1;
    const infoId = contentStart + pageCount;
    const objects = [];
    objects.push('<< /Type /Catalog /Pages 2 0 R >>');
    objects.push('<< /Type /Pages /Kids [' + pageTexts.map((_text, index) => (3 + index) + ' 0 R').join(' ') + '] /Count ' + pageCount + ' >>');
    for (let index = 0; index < pageCount; index += 1) {
        objects.push('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ' + fontId + ' 0 R >> >> /Contents ' + (contentStart + index) + ' 0 R >>');
    }
    objects.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    for (let index = 0; index < pageCount; index += 1) {
        const body = 'BT /F1 12 Tf 72 720 Td (' + pageTexts[index] + ') Tj ET';
        objects.push('<< /Length ' + body.length + ' >>\nstream\n' + body + '\nendstream');
    }
    objects.push('<< /Title (Big Test) /Producer (office-mode-test) >>');
    let pdf = '%PDF-1.4\n';
    const offsets = [];
    objects.forEach((body, index) => {
        offsets.push(pdf.length);
        pdf += (index + 1) + ' 0 obj\n' + body + '\nendobj\n';
    });
    const xrefStart = pdf.length;
    pdf += 'xref\n0 ' + (objects.length + 1) + '\n0000000000 65535 f \n';
    for (const offset of offsets) pdf += String(offset).padStart(10, '0') + ' 00000 n \n';
    pdf += 'trailer\n<< /Size ' + (objects.length + 1) + ' /Root 1 0 R /Info ' + infoId + ' 0 R >>\nstartxref\n' + xrefStart + '\n%%EOF\n';
    return Buffer.from(pdf, 'latin1');
}

await check('知识库预览：只抽取不入库、不回路径，坏 PDF 与不支持的格式分开说（第九十四轮）', async () => {
    const { root } = await freshRoot();
    noteWorkspace(root);
    const memoryConfig = { kbDir: join(root, 'kb') };
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memoryConfig });
    const route = ctx.routes.find((item) => item.path === KB_UPLOAD_PATH);
    const text = '预览用的文档正文。这条内容在预览阶段不应该变成知识库里的块。';
    const res = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ preview: true, name: '预览.md', folder: '资料', data: Buffer.from(text, 'utf8').toString('base64') }),
    }), res);
    assert.equal(res.status, 200, res.body);
    const body = JSON.parse(res.body);
    assert.equal(body.ok, true);
    assert.equal(body.action, 'kb-preview');
    const preview = body.result;
    assert.equal(preview.kind, 'text');
    assert.ok(preview.chars > 0, '要报字符数');
    assert.ok(preview.chunks > 0, '要报预计块数');
    assert.match(preview.head, /预览用的文档正文/, '要给出正文前段供人确认');
    assert.equal(preview.truncated, false);
    assert.equal('path' in preview, false, '预览响应不带任何路径');
    assert.equal(preview.name, '预览.md', '只回用户自己的文件名');

    const memory = createMemory({ root, memory: memoryConfig });
    assert.equal((await memory.kbList({})).items.length, 0, '预览不能入库');
    assert.equal(existsSync(join(root, 'uploads')), false, '预览不能把文件复制进工作区');
    assert.equal(existsSync(join(root, '.office', 'runtime', 'kb-preview')), false,
        '预览临时文件不落在工作区（私有 OS 临时目录）');

    const bad = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ preview: true, name: '坏.pdf', data: Buffer.from('这不是一份 PDF', 'utf8').toString('base64') }),
    }), bad);
    const badBody = JSON.parse(bad.body);
    assert.equal(badBody.ok, true);
    assert.equal(badBody.result.kind, 'broken', '读不出来的 PDF 是 broken，不是 scan（扫描件另有提示与出路）');
    assert.match(badBody.result.note, /损坏|密码/);

    const unsupported = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ preview: true, name: '脚本.exe', data: Buffer.from('MZ', 'utf8').toString('base64') }),
    }), unsupported);
    const unsupportedBody = JSON.parse(unsupported.body);
    assert.equal(unsupportedBody.result.kind, 'unsupported', '不支持的格式要直说');
    assert.match(unsupportedBody.result.note, /不能直接入库/);
    assert.equal(unsupportedBody.result.ocrEligible, false, '.exe 不在 OCR 白名单里，不能识别');

    // 第九十六轮：8 MB 是「抽正文」的上限，不是「判定类型」的上限。直接调引擎
    // （HTTP 层有 12 MiB JSON 体上限，大文件本来就不该走上传通道）。
    const tmpBefore = readdirSync(tmpdir()).filter((name) => name.startsWith('dsh-kb-preview-'));
    const big = await memory.kbPreview({ name: '大扫描.png', data: Buffer.alloc(8 * 1024 * 1024 + 1, 1) });
    assert.equal(big.kind, 'unknown', '超过 8 MB 不能断言是扫描件，只能报「未判定文本层」');
    assert.equal(big.ocrEligible, true, 'OCR 可识别的类型仍要留出识别入口');
    assert.match(big.note, /超过本地预览上限|没有判定文本层/);
    assert.equal(big.chars, 0, '没有抽取就不能编造字符数');
    assert.equal('path' in big, false);

    // 文本类守 kbIngest 的 2 MB 口径：预览就拦，别让卡上出现点了才失败的「入库」。
    const midText = await memory.kbPreview({ name: '中等文本.txt', data: Buffer.alloc(3 * 1024 * 1024, 65) });
    assert.equal(midText.kind, 'too_large', '2–8 MB 的文本预览阶段就拒绝');
    assert.match(midText.note, /2 MB/);
    assert.equal(midText.ocrEligible, false, '文本没有识别通道，不能给 OCR 按钮');

    const bigText = await memory.kbPreview({ name: '大文本.txt', data: Buffer.alloc(8 * 1024 * 1024 + 1, 65) });
    assert.equal(bigText.kind, 'too_large', '文本超过 8 MB 且不能 OCR：直接说上限');
    assert.match(bigText.note, /2 MB|单文档上限/);

    const huge = await memory.kbPreview({ name: '超大.png', data: Buffer.alloc(24 * 1024 * 1024 + 1, 1) });
    assert.equal(huge.kind, 'too_large', '超过 OCR 内容通道的 24 MB 也要拒绝');
    assert.equal(huge.ocrEligible, false);

    const doc = await memory.kbPreview({ name: '旧文档.doc', data: Buffer.from('DOC', 'utf8') });
    assert.equal(doc.kind, 'unsupported', '.doc 不能直接入库');
    assert.equal(doc.ocrEligible, true, '但 OCR 通道认它，要给出识别入口');

    const tmpAfter = readdirSync(tmpdir()).filter((name) => name.startsWith('dsh-kb-preview-'));
    assert.deepEqual(tmpAfter, tmpBefore, '预览的私有临时目录要清掉，不留在 OS 临时目录里');
});


await check('知识库预览：8–24 MB 的 PDF 自动按页拆分，逐份可入库（第一百一十轮）', async () => {
    const { root } = await freshRoot();
    noteWorkspace(root);
    const memoryConfig = { kbDir: join(root, 'kb') };
    const memory = createMemory({ root, memory: memoryConfig });
    const chunk = 'A'.repeat(3 * 1024 * 1024);
    const pdf = Buffer.concat([
        makeBigPdf([chunk, chunk]),
        Buffer.from('\n%' + 'x'.repeat(3 * 1024 * 1024) + '\n', 'latin1'),
    ]);
    assert.ok(pdf.length > 8 * 1024 * 1024, '夹具必须超过 8 MB 才走拆分分支：' + pdf.length);
    const before = readdirSync(tmpdir()).filter((name) => name.startsWith('dsh-kb-split-'));
    // 拆分只走 PyMuPDF（pip install pymupdf）。CI 是无外部依赖的纯 Node 环境，没有它时
    // 走不到拆分分支——这时仍要钉住「拆不了必须如实说、不假装」，但不假装自己测过拆分。
    const engines = await pdfEngines();
    if (!engines.render.includes('fitz')) {
        const fallback = await memory.kbPreview({ name: '大文档.pdf', data: pdf, splitTargetBytes: 1 });
        assert.equal(fallback.kind, 'unknown', '拆不了就给 unknown，不能编一份拆好的结果');
        assert.match(fallback.note, /自动拆分也没成功/);
        assert.match(fallback.note, /PyMuPDF/);
        assert.equal(fallback.parts, undefined, '没拆出来就不能带 parts');
        const left = readdirSync(tmpdir()).filter((name) => name.startsWith('dsh-kb-split-'));
        assert.deepEqual(left, before, '拆分失败也要收掉临时目录');
        return '本机没有 PyMuPDF，拆分分支未测（只校验如实报错）';
    }
    // 目标 1 字节：每页都必须单独成一份，避免 PyMuPDF 的 deflate 把两页压进同一份。
    const preview = await memory.kbPreview({ name: '大文档.pdf', data: pdf, splitTargetBytes: 1 });
    assert.equal(preview.kind, 'split', preview.note);
    assert.equal(preview.parts.length, 2, '两页各约 3 MB，应拆成两份');
    assert.equal(preview.parts[0].pageFrom, 1);
    assert.equal(preview.parts[1].pageTo, 2);
    assert.match(preview.parts[0].name, /第1页/);
    assert.ok(preview.parts.every((part) => part.kind === 'text'), JSON.stringify(preview.parts.map((part) => part.kind)));
    assert.ok(preview.parts.every((part) => part.chars > 0 && part.chunks > 0));
    assert.ok(preview.parts.every((part) => typeof part.data === 'string' && part.data.length > 0), '每份要带回可直接入库的 base64');
    assert.ok(preview.parts.every((part) => part.bytes <= 8 * 1024 * 1024), '每份必须落在 8 MB 抽取上限内');

    const saved = await memory.kbUpload({
        name: preview.parts[0].name,
        folder: '资料',
        data: Buffer.from(preview.parts[0].data, 'base64'),
    });
    assert.equal(saved.action, 'kb-ingest');
    assert.ok(saved.chunks > 0, '拆出来的分片必须能正常入库');
    const listed = await memory.kbList({ folder: '资料' });
    assert.equal(listed.total, 1);

    const after = readdirSync(tmpdir()).filter((name) => name.startsWith('dsh-kb-split-'));
    assert.deepEqual(after, before, '拆分临时目录要清掉');
});

await check('知识库上传：超过 33 MiB 的 JSON 请求体先回结构化 413，不再让浏览器只看到 socket 被掐（第九十八轮）', async () => {
    const { root } = await freshRoot();
    noteWorkspace(root);
    const memoryConfig = { kbDir: join(root, 'kb') };
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memoryConfig });
    const route = ctx.routes.find((item) => item.path === KB_UPLOAD_PATH);
    assert.ok(route, '知识库上传入口要注册');

    // 事故形态：只含 pad 字段、略超 33 MiB 的 JSON。以前服务端一超限就 destroy，
    // 浏览器侧只剩 fetch failed；现在必须先读完（或有界读）再回 413 + 结果。
    // 上限随大 PDF 预览从 12 MiB 提到 33 MiB（24 MiB 原文的 base64 + 1 MiB 余量）。
    const body = JSON.stringify({ pad: 'x'.repeat(33 * 1024 * 1024 + 1024) });
    const res = makeRes();
    await route.handler(makeReq({
        url: KB_UPLOAD_PATH,
        headers: { 'content-type': 'application/json' },
        body,
    }), res);

    assert.equal(res.status, 413, '超限要 413，不是 400，也不是直接断 socket');
    const json = JSON.parse(res.body);
    assert.equal(json.ok, false);
    assert.equal(json.code, 'body_too_large');
    assert.match(json.error, /34603008/, '回执里要写清这个端点的上限');
    assert.equal(json.retryable, false);
    assert.ok(res.body.length < 4096, '413 回执必须有界，不能把大 body 回显出去');
    assert.equal(/\bclose\b/i.test(String(res.headers.connection || '')), false,
        '有界关闭靠 req end/定时器，不靠 Connection: close 头');
});

await check('知识库图片路由：只服务 assets/ 下的图片，挡住越界（第八十九轮）', async () => {
    const { root } = await freshRoot();
    noteWorkspace(root);
    const memoryConfig = { kbDir: join(root, 'kb') };
    const assetDir = join(root, 'kb', 'assets', 'abc');
    await mkdir(assetDir, { recursive: true });
    // 完整 1×1 PNG：魔数嗅探要读到 IHDR 里的宽高，只有 8 字节签名会被判成不支持
    // （415），所以夹具必须是真图片而不是魔数前缀。
    const pixels = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
    await writeFile(join(assetDir, '1-pic.png'), pixels);
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memoryConfig });
    const route = ctx.routes.find((item) => item.path === KB_ASSET_PATH);
    assert.ok(route, '图片资产入口要注册');
    const res = makeRes();
    await route.handler({ method: 'GET', url: KB_ASSET_PATH + '?path=assets/abc/1-pic.png' }, res);
    assert.equal(res.status, 200);
    assert.equal(res.headers['content-type'], 'image/png');
    assert.deepEqual(res.body, pixels, '要原样返回图片字节');

    const outside = makeRes();
    await route.handler({ method: 'GET', url: KB_ASSET_PATH + '?path=assets/../memory.json' }, outside);
    assert.equal(outside.status, 400, '不能借图片端点读 assets 之外的文件');
    const missing = makeRes();
    await route.handler({ method: 'GET', url: KB_ASSET_PATH + '?path=assets/abc/none.png' }, missing);
    assert.equal(missing.status, 404);
});

await check('写入路由：POST 走 memory-replace，拒绝 GET 与非 JSON（第八十五轮）', async () => {
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    const added = await instance.mutate({ action: 'add', target: 'user', content: '旧的正文。' });
    noteWorkspace(root);
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => instance });
    const route = ctx.routes.find((item) => item.path === MEMORY_MUTATE_PATH);
    assert.ok(route, '写入口要注册');

    const getRes = makeRes();
    await route.handler({ method: 'GET', url: MEMORY_MUTATE_PATH, headers: {} }, getRes);
    assert.equal(getRes.status, 405, '只读方法要挡住');

    const formRes = makeRes();
    await route.handler(makeReq({
        url: MEMORY_MUTATE_PATH,
        headers: { 'content-type': 'application/x-www-form-urlencoded' },
        body: 'action=memory-replace',
    }), formRes);
    assert.equal(formRes.status, 415, '跨站表单能发的类型要挡住');

    const okRes = makeRes();
    await route.handler(makeReq({
        url: MEMORY_MUTATE_PATH,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ action: 'memory-replace', target: 'user', id: added.entry.id, content: '新的正文。' }),
    }), okRes);
    assert.equal(okRes.status, 200);
    assert.equal(JSON.parse(okRes.body).ok, true);
    const read = await instance.read({ layer: 'hot', query: '新的正文' });
    assert.ok(read.hot.items.some((item) => item.text === '新的正文。'), '写入要真的落盘，不是只回 200');
});

await check('路由：未知工作目录一律 404，不会照路径去读盘', async () => {
    resetWorkspaces();
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({}) });
    const handler = ctx.routes[0].handler;

    const empty = await callRoute(handler);
    assert.equal(empty.json.ok, false);
    assert.match(empty.json.error, /还没有见过任何工作目录/);

    const unknown = await callRoute(handler, { url: `${MEMORY_VIEW_PATH}?cwd=C:/Windows` });
    assert.equal(unknown.status, 404, '没见过的路径必须是 404');
    assert.match(unknown.json.error, /不在已知列表里/);
    assert.equal(unknown.json.hot, undefined, '拒绝时不能带任何内容');
});

// ── 第三十一轮：重启之后、以及还没跑过办公工具的工作区 ──────────────────────
//
// 老实现的信任来源只有「本进程真跑过工具的工作目录」（进程内），于是重启 DSH
// 之后面板只能显示「还没有见过任何工作目录」—— 记忆文件明明在盘上。
// 现在补两个来源：宿主工作区登记表（持久、跨重启）与会话史（兜底）。

await check('重启后：工作区登记表里的目录不用先跑工具也能显示记忆', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'project', content: '重启前写下的约定。' });
    await instance.log([{ path: '季度汇报.pptx', format: 'ppt', purpose: '季度主线' }]);

    const ctx = fakeCtx({ workspaceRegistry: fakeRegistry([{ path: root, title: '季度项目' }]) });
    registerMemoryView(ctx, { getMemory: () => memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.status, 200, '登记过的目录不该再 404');
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 1);
    assert.equal(snap.json.ledger.length, 1);
    assert.equal(snap.json.label, '季度项目', '面板用登记表的标题当标签');
    assert.equal(snap.json.workspaces.length, 1);
    assert.equal(snap.json.workspaces[0].source, 'workspace');
    assert.equal(snap.json.workspaces[0].label, '季度项目');
});

await check('未使用的工作区：没有记忆目录时给空态，不是 404 也不是报错', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const ctx = fakeCtx({ workspaceRegistry: fakeRegistry([{ path: root, title: '还没开工的项目' }]) });
    registerMemoryView(ctx, { getMemory: () => memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.status, 200);
    assert.equal(snap.json.ok, true, '工作区里还没有记忆也该给一份空快照');
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 0);
    assert.equal(snap.json.ledger.length, 0);
    assert.equal(snap.json.counts.ledgerTotal, 0);
    assert.equal(snap.json.stores[0].bytes, 0);
});

await check('重启后默认根：先开有记忆、且最近在干活的那个工作区', async () => {
    resetWorkspaces();
    const older = await freshRoot();
    const newer = await freshRoot();
    await createMemory({ root: older.root, memory: older.memory }).mutate({ action: 'add', target: 'project', content: '旧项目的记忆。' });
    // 两个记忆目录的 mtime 必须真的分开：没有工作区时间戳时默认根的判据会退到它，
    // 同毫秒内写完两个目录时这条断言就变成掷骰子。
    await new Promise((resolve) => setTimeout(resolve, 25));
    await createMemory({ root: newer.root, memory: newer.memory }).mutate({ action: 'add', target: 'project', content: '新项目的记忆。' });
    // 登记表顺序是「用户自己排的」：旧项目在前。面板不该照搬它，而该挑有内容的那个。
    const ctx = fakeCtx({
        workspaceRegistry: fakeRegistry([
            { path: older.root, title: '旧项目' },
            { path: newer.root, title: '新项目' },
        ]),
    });
    registerMemoryView(ctx, { getMemory: () => newer.memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.workspaces.length, 2, '两个工作区都要列出来给面板切换');
    // 都没有工作区时间戳时，取记忆库文件最近动过的那个（这里后写的 newer 更新）
    assert.equal(snap.json.cwd, newer.root);
    assert.equal(snap.json.label, '新项目');

    // 显式指定时按指定来
    const byCwd = await callRoute(ctx.routes[0].handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(older.root)}` });
    assert.equal(byCwd.json.ok, true);
    assert.equal(byCwd.json.cwd, older.root);
    assert.match(byCwd.json.hot[0].content, /旧项目/);
});

await check('重启后默认根：工作区记录的最近变动时刻优先于记忆文件的 mtime', async () => {
    resetWorkspaces();
    const busy = await freshRoot();
    const idle = await freshRoot();
    await createMemory({ root: idle.root, memory: idle.memory }).mutate({ action: 'add', target: 'project', content: '刚写过记忆但没人在干活的项目。' });
    await new Promise((resolve) => setTimeout(resolve, 25));
    await createMemory({ root: busy.root, memory: busy.memory }).mutate({ action: 'add', target: 'project', content: '最近在干活的项目。' });
    // 故意把 idle 的记忆改到更晚，再把 busy 的工作区时间戳设成更晚：
    // 「最近在哪儿干活」该赢过「哪个记忆文件最近被碰过」。
    const ctx = fakeCtx({
        workspaceRegistry: fakeRegistry([
            { path: idle.root, title: '闲项目', updatedAt: '2026-01-01T00:00:00.000Z' },
            { path: busy.root, title: '忙项目', updatedAt: '2026-09-28T00:00:00.000Z' },
        ]),
    });
    registerMemoryView(ctx, { getMemory: () => busy.memory });
    const snap = await callRoute(ctx.routes[0].handler);
    assert.equal(snap.json.cwd, busy.root, '工作区时间戳更晚的那个才是「最近在干活」的');
    assert.equal(snap.json.label, '忙项目');
});

await check('工作区登记表为空时：会话史里的工作目录兜底（不登记成项目也能显示）', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    await createMemory({ root, memory }).mutate({ action: 'add', target: 'project', content: '会话史兜底的记忆。' });
    const ctx = fakeCtx({
        sessionPersistence: fakeSessions([
            { header: { cwd: root, createdAt: 1_700_000_000_000 } },
            { header: { cwd: '', createdAt: 1_700_000_000_001 } },
            { header: { createdAt: 1_700_000_000_002 } },
        ]),
    });
    registerMemoryView(ctx, { getMemory: () => memory });
    const snap = await callRoute(ctx.routes[0].handler);

    assert.equal(snap.status, 200);
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 1);
    assert.equal(snap.json.workspaces[0].source, 'session');

    // 登记表里没有、会话史里有：显式点名时也认
    const byCwd = await callRoute(ctx.routes[0].handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(root)}` });
    assert.equal(byCwd.json.ok, true);
    assert.equal(byCwd.json.cwd, root);
});

await check('工作区登记表里有别的项目时：仍然不认登记表之外的路径', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const ctx = fakeCtx({ workspaceRegistry: fakeRegistry([{ path: root, title: '已知项目' }]) });
    registerMemoryView(ctx, { getMemory: () => memory });
    const res = await callRoute(ctx.routes[0].handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(join(TMP, 'case-unknown'))}` });
    assert.equal(res.status, 404, '登记表之外的路径必须还是 404');
    assert.equal(res.json.hot, undefined);
    assert.equal(res.json.workspaces.length, 1, '拒绝时仍把已知工作区带回给面板');
});

await check('宿主服务坏掉时降级：登记表/会话史抛错不把端点打挂', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const ctx = fakeCtx({
        workspaceRegistry: { list: () => { throw new Error('存储坏了'); } },
        sessionPersistence: { list: async () => { throw new Error('会话存储坏了'); } },
    });
    registerMemoryView(ctx, { getMemory: () => memory });
    const res = await callRoute(ctx.routes[0].handler);
    assert.equal(res.status, 200, '拿不到宿主事实时退回老行为（空态），不该 500');
    assert.equal(res.json.ok, false);
    assert.match(res.json.error, /还没有见过任何工作目录/);

    // 本进程跑过工具的根不受影响
    noteWorkspace(root);
    const snap = await callRoute(ctx.routes[0].handler);
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
});

await check('路由：只服务 noteWorkspace 记下的根', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'project', content: '这个目录里的一条约定。' });
    noteWorkspace(root);
    assert.deepEqual(knownWorkspaces().map((item) => item.cwd), [root]);

    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memory });
    const handler = ctx.routes[0].handler;

    const snap = await callRoute(handler);
    assert.equal(snap.status, 200);
    assert.equal(snap.json.ok, true);
    assert.equal(snap.json.cwd, root);
    assert.equal(snap.json.hot.length, 1);
    assert.equal(snap.headers['cache-control'], 'no-store', '记忆内容不该被缓存');
    assert.equal(snap.json.workspaces.length, 1);

    const byCwd = await callRoute(handler, { url: `${MEMORY_VIEW_PATH}?cwd=${encodeURIComponent(root)}` });
    assert.equal(byCwd.json.ok, true);
    assert.equal(byCwd.json.cwd, root);
});

await check('路由：非 GET 方法一律 405（这个端点只读）', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    noteWorkspace(root);
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => memory });
    const handler = ctx.routes[0].handler;
    for (const method of ['POST', 'PUT', 'DELETE', 'PATCH']) {
        const res = await callRoute(handler, { method });
        assert.equal(res.status, 405, `${method} 应被拒绝`);
        assert.match(res.json.error, /只读/);
    }
    const head = await callRoute(handler, { method: 'HEAD' });
    assert.equal(head.status, 200, 'HEAD 允许');
});

await check('路由：记忆关掉时不返回任何内容', async () => {
    resetWorkspaces();
    const { root, memory } = await freshRoot();
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'project', content: '不该被读出来的内容。' });
    noteWorkspace(root);

    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({ ...memory, enabled: false }) });
    const res = await callRoute(ctx.routes[0].handler);
    assert.equal(res.json.ok, false);
    assert.match(res.json.error, /关掉/);
    assert.equal(res.json.hot, undefined);
    assert.doesNotMatch(res.raw, /不该被读出来的内容/);
});

await check('路由：内部出错时给出可读错误，而不是把异常抛给服务器', async () => {
    resetWorkspaces();
    noteWorkspace(join(TMP, 'case-ghost'));
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({}) });
    // 目录不存在时引擎照常给出空结果；这里要确认的是「不会 500 崩掉」。
    const res = await callRoute(ctx.routes[0].handler);
    assert.ok(res.status === 200 || res.status === 500);
    if (res.status === 500) assert.match(res.json.error, /读取记忆失败/);
    else assert.equal(res.json.ok, true);
});

// ── 第七十一轮：记忆导出（单文件迁移） ─────────────────────────────────────
//
// 用户要的是「方便迁移」：agent 记忆没有公认的公有协议，所以导出成**单文件**。
// 两条通道都要真的是单文件、能被另一个目录的引擎读回去 —— 只测状态码等于没测。

/** 起一个「有两条热记忆 + 一条台账」的根，返回导出端点。 */
async function exportCase(memoryConfig = {}) {
    resetWorkspaces();
    const { root, memory } = await freshRoot(memoryConfig);
    const instance = createMemory({ root, memory });
    await instance.mutate({ action: 'add', target: 'user', content: '用户偏好：导出要能搬走整份记忆。' });
    await instance.mutate({ action: 'add', target: 'project', content: '项目约定：这份目录用 business 主题。' });
    await instance.log([{ path: '季度汇报.pptx', format: 'ppt', theme: 'business', purpose: '季度主线' }]);
    noteWorkspace(root);
    const ctx = fakeCtx();
    registerMemoryView(ctx, { getMemory: () => ({ ...memory, ...memoryConfig }) });
    const route = ctx.routes.find((item) => item.path === EXPORT_VIEW_PATH);
    return { root, memory, instance, handler: route.handler };
}

await check('导出：JSON 是能直接 importPack 的版本化单文件包（真的搬得走）', async () => {
    const { root, handler } = await exportCase();
    const res = await callRoute(handler, { url: `${EXPORT_VIEW_PATH}?cwd=${encodeURIComponent(root)}` });
    assert.equal(res.status, 200);
    assert.match(res.headers['content-type'], /^application\/json/);
    assert.equal(res.headers['cache-control'], 'no-store');
    assert.match(res.headers['content-disposition'], /^attachment; filename="dsh-office-memory-[^"]*\.json"$/,
        '浏览器要靠这个头直接存成文件');
    const pack = res.json;
    assert.equal(pack.format, 'dsh-office-memory-pack', '导出必须带格式标记，导入侧才认得');
    assert.equal(pack.version, 1);

    // 搬到一个全新的目录：这才是「方便迁移」的定义。
    const other = await freshRoot();
    const target = createMemory({ root: other.root, memory: other.memory });
    const report = await target.importPack(pack);
    assert.ok(report.entries >= 2, `导入应带回热记忆条目（实际 ${report.entries}）`);
    const browse = await target.browse();
    const contents = browse.hot.map((entry) => entry.content).join('\n');
    assert.match(contents, /导出要能搬走整份记忆/);
    assert.match(contents, /这份目录用 business 主题/);
    assert.equal(browse.ledger.length, 1, '台账也随包走');
});

await check('导出：Markdown 版是 CommonMark 可读版（无私有 front matter）', async () => {
    const { root, handler } = await exportCase();
    const res = await callRoute(handler, { url: `${EXPORT_VIEW_PATH}?cwd=${encodeURIComponent(root)}&format=md` });
    assert.equal(res.status, 200);
    assert.match(res.headers['content-type'], /^text\/markdown/);
    assert.match(res.headers['content-disposition'], /\.md"$/);
    assert.match(res.raw, /^# 办公记忆导出/, '一级标题说明这份文件是什么');
    assert.match(res.raw, /## 用户偏好/);
    assert.match(res.raw, /## 项目与环境/);
    assert.match(res.raw, /## 台账/);
    assert.match(res.raw, /导出要能搬走整份记忆/, '内容要真的写进去');
    assert.doesNotMatch(res.raw.split(/\r?\n/)[0], /^---$/,
        '不放 YAML front matter：那是 Obsidian 的私有约定，不是 CommonMark（段间的 --- 是合法分隔线）');
});

await check('导出：format 白名单 400、非 GET 405 带 Allow、未知目录 404', async () => {
    const { root, handler } = await exportCase();
    const cwd = encodeURIComponent(root);

    const badFormat = await callRoute(handler, { url: `${EXPORT_VIEW_PATH}?cwd=${cwd}&format=xml` });
    assert.equal(badFormat.status, 400);
    assert.match(badFormat.json.error, /format 只支持 json 或 md/);
    assert.match(badFormat.json.hint, /不传 format 就是 json/, '错误要带下一步');

    const post = await callRoute(handler, { method: 'POST', url: `${EXPORT_VIEW_PATH}?cwd=${cwd}` });
    assert.equal(post.status, 405);
    assert.equal(post.headers.allow, 'GET, HEAD');

    const unknown = await callRoute(handler, { url: `${EXPORT_VIEW_PATH}?cwd=C:/Windows` });
    assert.equal(unknown.status, 404, '导出能拿到的东西比快照多，边界更不能放松');
    assert.equal(unknown.json.content, undefined);
});

await check('导出：HEAD 只回头带长度不带正文；记忆关掉时 403', async () => {
    const { root, handler } = await exportCase();
    const cwd = encodeURIComponent(root);
    const get = await callRoute(handler, { url: `${EXPORT_VIEW_PATH}?cwd=${cwd}` });
    const head = await callRoute(handler, { method: 'HEAD', url: `${EXPORT_VIEW_PATH}?cwd=${cwd}` });
    assert.equal(head.status, 200);
    assert.equal(head.raw, '', 'HEAD 不该带正文');
    assert.equal(head.headers['content-length'], get.headers['content-length'], '长度与 GET 一致，客户端才知道文件多大');
    assert.match(head.headers['content-disposition'], /\.json"$/);

    const offCase = await exportCase({ enabled: false });
    const off = await callRoute(offCase.handler, { url: `${EXPORT_VIEW_PATH}?cwd=${encodeURIComponent(offCase.root)}` });
    assert.equal(off.status, 403, '记忆关掉时导出必须明确拒绝，不能给一份空文件');
    assert.match(off.json.error, /关掉/);
    assert.doesNotMatch(off.raw, /导出要能搬走整份记忆/);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? (item.note ? `  —— ${item.note}` : '') : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`view: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
