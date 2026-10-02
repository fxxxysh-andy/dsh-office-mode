/**
 * 三层记忆的测试：热记忆（偏好与约定）/ 台账（交付物登记）/ 归档（下沉的旧条目）。
 *
 * 这一份测的是**记忆自身的契约**，与 office_help / office_run 的接线：
 *   - 存储：Markdown 是投影、json/jsonl 是真源；容量满了向下沉而不是丢；
 *   - 写入：add / replace / remove 的命中规则（唯一命中）与非法输入；
 *   - 读取：三层共用一份预算、三重上限、去重、截断标记；
 *   - 生命周期：台账滚动、归档保留上限、并发写不丢条目；
 *   - 接线：office_memory 工具的增删改查、office_help 的投影、office_run 的自动台账。
 *
 * 与 mnemon 的差别在 memory.js 顶部写明了（无回合配额、无远端 provider、按工作目录）。
 *
 * 跑法：node test/memory.mjs
 */
import { existsSync, readFileSync } from 'node:fs';
import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import {
    ARCHIVE_VOLUME_MAX_ITEMS,
    createMemory,
    globalMemoryDir,
    memoryExists,
    memoryPaths,
    memoryStats,
    renderDigest,
    renderMutation,
    renderRead,
    renderSunk,
    renderRelated,
    resolveTimeWindow,
} from '../src/memory.js';
import { migrateMnemon, readMnemonInsights, renderMigration } from '../src/migrate.js';
import { projectDigest, projectionStateOf, REFRESH_EVERY, resetProjection, sessionIdOf, signalOf } from '../src/projection.js';
import { buildTools } from '../src/tools.js';

// 测试中间产物统一落在仓库根的 .office/tmp/<suite>/ 下（不再用包内的 test/.tmp）。
const TMP = fileURLToPath(new URL('../../../.office/tmp/memory/', import.meta.url));

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 每个用例一个干净的工作目录：记忆是按工作目录存的，串了目录就串了结论。 */
let caseId = 0;
async function freshRoot(memoryConfig = {}) {
    caseId += 1;
    const root = join(TMP, `case-${caseId}`);
    await rm(root, { recursive: true, force: true });
    await mkdir(root, { recursive: true });
    return { root, memory: createMemory({ root, memory: memoryConfig }) };
}

const execAt = (root) => ({ agent: { session: { header: { cwd: root } } } });

await rm(TMP, { recursive: true, force: true });

// ── 存储与投影 ──────────────────────────────────────────────────────────────

await check('空目录：三层都读得到，且不报错', async () => {
    const { root, memory } = await freshRoot();
    const value = await memory.read({});
    assert.deepEqual(value.hot.items, []);
    assert.equal(value.ledger.total, 0);
    assert.equal(value.archive.items_total, 0);
    assert.equal(value.truncated, false);
    assert.equal(memoryExists(root), false, '只读不该顺手建目录');
    assert.equal((await memoryStats(root)).exists, false);
});

await check('Markdown 是投影：写条目后 USER.md / MEMORY.md 同步生成，且写明是投影', async () => {
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'user', content: '偏好：表达书面化，避免第二人称。', importance: 'critical' });
    await memory.mutate({ action: 'add', target: 'project', content: '季度汇报统一用 business 主题。' });
    const paths = memoryPaths(root, {});
    const user = await readFile(paths.user, 'utf8');
    const project = await readFile(paths.project, 'utf8');
    assert.ok(user.includes('偏好：表达书面化'), 'USER.md 应有用户偏好');
    assert.ok(!user.includes('business'), 'USER.md 不该混进项目条目');
    assert.ok(project.includes('business'), 'MEMORY.md 应有项目条目');
    assert.ok(!project.includes('书面化'), 'MEMORY.md 不该混进用户条目');
    assert.ok(user.includes('投影') && project.includes('投影'), '两份投影都要写明自己是生成物');
    assert.ok(/用量 \d+ \/ \d+ 字节/.test(user), '投影要带用量行');
    // 真源是 memory.json：投影可以被覆盖重建，条目不会丢。
    const store = JSON.parse(await readFile(paths.memory, 'utf8'));
    assert.equal(store.entries.length, 2);
    assert.equal(store.version, 1);
    assert.match(store.contentRevision, /^[0-9a-f]{40}$/);
});

await check('contentRevision 只由内容决定：同一份条目算出来一样，改了就变', async () => {
    const { root, memory } = await freshRoot();
    const first = await memory.mutate({ action: 'add', target: 'project', content: '约定 A' });
    const revisionA = first.contentRevision;
    const snapshot = await memory.snapshot();
    assert.equal(snapshot.contentRevision, revisionA);
    const second = await memory.mutate({ action: 'add', target: 'project', content: '约定 B' });
    assert.notEqual(second.contentRevision, revisionA);
});

await check('两个指纹分家：contentRevision 由内容算，writeRevision 由「写没写」算（24-17）', async () => {
    // 第二十四轮 P2-5 / 第二十六轮 24-17：两个指纹**都叫过 revision**，算法与用途不同。
    // 改名之后必须钉住三件事：新名字在、旧名字不在、两者真的会在同一份数据上分岔。
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '约定：季度汇报用 business。' });
    const snapshot = await memory.snapshot();
    const digest = await memory.digest();
    assert.match(snapshot.contentRevision, /^[0-9a-f]{40}$/);
    assert.match(digest.writeRevision, /^[0-9a-f]{40}$/);
    assert.equal(snapshot.revision, undefined, '内容指纹不该还叫 revision：投影侧的 writeRevision 与它同词两义');
    assert.equal(digest.revision, undefined, '写指纹不该还叫 revision');
    assert.notEqual(snapshot.contentRevision, digest.writeRevision, '两个指纹的算法本来就不同，值不该恰好相等');
    // 「写了一次、内容一字没改」：writeRevision 要变，contentRevision 不许变。
    await memory.mutate({ action: 'replace', target: 'project', oldText: '季度汇报用 business', content: '约定：季度汇报用 business。' });
    const after = await memory.digest();
    assert.equal((await memory.snapshot()).contentRevision, snapshot.contentRevision, '内容没改，contentRevision 就不该变');
    assert.notEqual(after.writeRevision, digest.writeRevision, '写操作发生了，writeRevision 必须变');
});

// ── 写入协议 ────────────────────────────────────────────────────────────────

await check('replace：唯一命中才改，改后用量与投影同步', async () => {
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '季度汇报统一用 business 主题。' });
    const value = await memory.mutate({
        action: 'replace',
        target: 'project',
        content: '季度与年度汇报统一用 business 主题。',
        oldText: '季度汇报统一用',
        importance: 'critical',
    });
    assert.equal(value.previous.content, '季度汇报统一用 business 主题。');
    assert.equal(value.entry.content, '季度与年度汇报统一用 business 主题。');
    assert.equal(value.entry.importance, 'critical');
    assert.equal(value.usage.project.entryCount, 1);
    const project = await readFile(memoryPaths(root, {}).project, 'utf8');
    assert.ok(project.includes('季度与年度汇报'));
});

await check('replace / remove：命中零条或多条都拒绝，并给出可读原因', async () => {
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '约定：甲方文档用宋体。' });
    await memory.mutate({ action: 'add', target: 'project', content: '约定：乙方文档用黑体。' });
    await assert.rejects(() => memory.mutate({ action: 'replace', content: 'X', oldText: '不存在的片段' }), /没有哪条记忆包含/);
    await assert.rejects(() => memory.mutate({ action: 'remove', oldText: '约定：' }), /命中 2 条/);
    await assert.rejects(() => memory.mutate({ action: 'replace', content: 'X', oldText: '约定：' }), /命中 2 条/);
    const snapshot = await memory.snapshot();
    assert.equal(snapshot.entries.length, 2, '被拒绝的写入不该改动存储');
});

await check('remove：唯一命中即删，投影里消失', async () => {
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'user', content: '忌：正文用第二人称。' });
    await memory.mutate({ action: 'add', target: 'user', content: '偏好：标题用短句。' });
    const value = await memory.mutate({ action: 'remove', target: 'user', oldText: '第二人称' });
    assert.equal(value.action, 'remove');
    assert.equal(value.entryCount, 1);
    assert.equal(value.evicted, false, '主动删除不该被回执成「被容量维持下沉」');
    const user = await readFile(memoryPaths(root, {}).user, 'utf8');
    assert.ok(!user.includes('第二人称'));
    assert.ok(user.includes('标题用短句'));
});

await check('非法写入：动作 / 目标 / 空内容 / 缺 oldText / 超大单条都明确报错', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 1024 });
    await assert.rejects(() => memory.mutate({ action: 'delete', content: 'x' }), /只支持 add \/ replace \/ remove/);
    await assert.rejects(() => memory.mutate({ action: 'add', target: 'global', content: 'x' }), /只支持 user \/ project/);
    await assert.rejects(() => memory.mutate({ action: 'add', content: '   ' }), /不能为空/);
    await assert.rejects(() => memory.mutate({ action: 'replace', content: 'x' }), /需要 oldText/);
    await assert.rejects(() => memory.mutate({ action: 'add', content: '填'.repeat(600) }), /容量上限还大/);
});

await check('未知重要度按 normal 处理，不报错', async () => {
    const { memory } = await freshRoot();
    const value = await memory.mutate({ action: 'add', target: 'project', content: '随便一条', importance: '超高' });
    assert.equal(value.entry.importance, 'normal');
});

// ── 容量维持与归档 ──────────────────────────────────────────────────────────

await check('容量下沉：超出上限时最旧 / 最不重要的条目进归档，用量回到上限内', async () => {
    const { root, memory } = await freshRoot({ projectLimitBytes: 300, archiveKeep: 10 });
    await memory.mutate({ action: 'add', target: 'project', content: `重要约定：${'重'.repeat(40)}`, importance: 'critical' });
    for (let i = 0; i < 6; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `旧条目 ${i}：${'旧'.repeat(30)}`, importance: 'low' });
    }
    const snapshot = await memory.snapshot();
    assert.ok(snapshot.targets.project.used <= 300, `用量应回到上限内，实际 ${snapshot.targets.project.used}`);
    assert.ok(snapshot.entries.some((entry) => entry.content.includes('重要约定')), 'critical 条目应被保住');
    // 指路条目：告诉读投影的人内容去哪儿了。
    assert.ok(snapshot.entries.some((entry) => entry.id === 'pointer:archive'), '应留下归档指路条目');
    const archive = await memory.read({ layer: 'archive', query: '旧条目' });
    assert.ok(archive.archive.items.length > 0, '下沉的条目应能在归档里检索到');
});

await check('下沉不会丢内容：归档文件与索引同时更新', async () => {
    const { root, memory } = await freshRoot({ projectLimitBytes: 400, archiveKeep: 10 });
    await memory.mutate({ action: 'add', target: 'project', content: `唯一标记词：${'甲'.repeat(60)}`, importance: 'critical' });
    for (let i = 0; i < 4; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `会下沉 ${i}：${'乙'.repeat(40)}`, importance: 'low' });
    }
    const paths = memoryPaths(root, {});
    const index = JSON.parse(await readFile(paths.index, 'utf8'));
    assert.ok(index.digests.length >= 1, '应有归档摘要');
    const month = index.digests[0].month;
    const file = await readFile(join(paths.archive, `${month}.md`), 'utf8');
    assert.ok(file.includes('热记忆下沉'), '归档文件要写明下沉原因');
    assert.ok(file.includes('会下沉'), '归档文件里要有下沉的内容');
    // 检索得到 + 文件里看得到，两处必须一致。
    const found = await memory.read({ layer: 'archive', query: '会下沉' });
    assert.ok(found.archive.items.length >= 1);
});

await check('指路条目只留一条：多次下沉不堆积', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 400, archiveKeep: 10 });
    for (let i = 0; i < 12; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `条目 ${i}：${'丙'.repeat(20)}`, importance: 'low' });
    }
    const snapshot = await memory.snapshot();
    const pointers = snapshot.entries.filter((entry) => entry.id === 'pointer:archive');
    assert.equal(pointers.length, 1, `指路条目应只有一条，实际 ${pointers.length}`);
    assert.ok(snapshot.targets.project.used <= 400);
});

await check('单条就超过上限且无可下沉时抛错，而不是静默截断', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 40 });
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: '这一个条目本身就超过四十字节的上限了。' }),
        /没有可下沉的条目|容量上限还大|超出上限/,
    );
});

await check('容量维持按重要度次序下沉：先走更旧的 low，不动 critical', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 500, archiveKeep: 10 });
    const low = (name) => ({ action: 'add', target: 'project', content: `可选项${name}：${'可'.repeat(30)}`, importance: 'low' });
    await memory.mutate({ action: 'add', target: 'project', content: `必须遵守：${'要'.repeat(30)}`, importance: 'critical' });
    await memory.mutate(low('一'));
    await memory.mutate(low('二'));
    await memory.mutate(low('三'));
    const added = await memory.mutate(low('四'));
    const snapshot = await memory.snapshot();
    const contents = snapshot.entries.map((entry) => entry.content);
    assert.ok(contents.some((text) => text.includes('必须遵守')), 'critical 条目必须留下');
    assert.ok(!contents.some((text) => text.includes('可选项一')), '最旧的 low 先下沉');
    assert.ok(!contents.some((text) => text.includes('可选项二')), '第二旧的 low 接着下沉');
    assert.ok(contents.some((text) => text.includes('可选项三')), '较新的 low 应当留下');
    assert.equal(added.evicted, false, '刚写的这条不该被下沉');
    assert.ok(snapshot.targets.project.used <= 500);
});

await check('容量实在装不下时，如实回执「刚写的这条也被下沉了」', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 300, archiveKeep: 10 });
    await memory.mutate({ action: 'add', target: 'project', content: `必须遵守：${'要'.repeat(30)}`, importance: 'critical' });
    await memory.mutate({ action: 'add', target: 'project', content: `旧的：${'旧'.repeat(30)}`, importance: 'low' });
    const added = await memory.mutate({ action: 'add', target: 'project', content: `刚写的：${'新'.repeat(30)}`, importance: 'low' });
    assert.equal(added.evicted, true, '这个容量下刚写的 low 确实留不住');
    const text = renderMutation(added);
    assert.ok(text.includes('刚写的这条也下沉了'), '回执必须说清楚，不能让模型以为记下了');
    const snapshot = await memory.snapshot();
    assert.ok(snapshot.entries.some((entry) => entry.content.includes('必须遵守')));
});

await check('下沉可见（新-10）：回执点名被下沉的 id，sunk 列最近几次，归档读侧带原因', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 500, archiveKeep: 10 });
    let last = null;
    for (let i = 0; i < 6; i += 1) {
        last = await memory.mutate({ action: 'add', target: 'project', content: `会下沉 ${i}：${'沉'.repeat(40)}`, importance: 'low' });
    }
    const maintenance = (last.maintenance ?? []).find((item) => item.target === 'project' && item.moved > 0);
    assert.ok(maintenance !== undefined, '这个容量下应当发生过下沉');
    assert.equal(maintenance.movedPreview.length, maintenance.moved, 'movedPreview 要与下沉条数一一对应');
    // ① 回执点名：只报「下沉了 2 条」时，热层忘了哪几条不可见（新-10 的病灶）
    const receipt = renderMutation(last);
    for (const entry of maintenance.movedPreview) {
        assert.ok(receipt.includes(entry.id), `回执要点名被下沉的条目 ${entry.id}`);
        assert.ok(entry.preview.length > 0, '预览不该为空');
    }
    assert.ok(receipt.includes("action: 'sunk'"), '回执要指到「最近下沉」这个读法');

    // ② sunk：按时间倒序列出下沉块，带 id 与预览
    const sunk = await memory.sunk({ limit: 3 });
    assert.ok(sunk.totalItems >= maintenance.moved, `sunk 要数出下沉过的条目，实际 ${sunk.totalItems}`);
    assert.ok(sunk.blocks.length >= 1, 'sunk 要给最近的下沉块');
    assert.ok(sunk.blocks[0].items.length >= 1, '块里要有条目');
    const sunkText = renderSunk(sunk);
    assert.ok(sunkText.includes(maintenance.movedPreview[0].id), 'sunk 回执里要有 id');
    assert.ok(sunkText.includes("layer: 'archive'"), 'sunk 要指到取回正文的那条路');
    assert.ok(!sunkText.includes('沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉沉'), 'sunk 只给预览，不搬正文');

    // ③ 归档读侧带原因：从「它在归档里」升级成「它为什么在归档里」
    const archived = await memory.read({ layer: 'archive', query: '会下沉' });
    assert.ok(archived.archive.items.length >= 1, '下沉的条目要能在归档里检索到');
    assert.ok(archived.archive.items.every((item) => item.reason === 'hot-overflow'), '归档条目要带 reason');
    assert.ok(renderRead(archived).includes('热记忆下沉'), '归档读侧要写明原因');

    // ④ 指路条目是投影里唯一看得见「热层少了东西」的一行：要报条数与清单入口。
    //    它自己也是热记忆的一条，所以只许短 —— 长一个字就少一个字给真实条目。
    const snapshot = await memory.snapshot();
    const pointer = snapshot.entries.find((entry) => entry.id === 'pointer:archive');
    assert.ok(pointer !== undefined, '应留下归档指路条目');
    assert.ok(/（\d+ 条，清单见 action:'sunk'）/.test(pointer.content),
        `指路条目要带条数与清单入口：${pointer.content}`);
    assert.ok(Buffer.byteLength(pointer.content, 'utf8') <= 200, '指路条目要有字节上界（它占的是热层容量）');
});

await check('写入期长度门（新-11）：硬上限拒绝并给压缩模板，建议线只提醒，改短不啰嗦', async () => {
    const { memory } = await freshRoot({ projectLimitBytes: 20_000, entryLimitBytes: 300, entryHintBytes: 120 });
    // ① 硬上限：拒绝，且报错里要说清「怎么办」（不然模型只会原样重试）
    await assert.rejects(
        () => memory.mutate({ action: 'add', target: 'project', content: `太长了：${'长'.repeat(200)}` }),
        (error) => /超过单条上限 300 字节/.test(error.message)
            && /kb-ingest/.test(error.message)
            && /memory\.entryLimitBytes/.test(error.message),
    );
    // 边界：正好等于上限要放行（判据是「大于」，不是「大于等于」）
    const exact = await memory.mutate({ action: 'add', target: 'project', content: 'x'.repeat(300) });
    assert.ok(exact.entry.id !== undefined, '正好等于上限应放行');

    // ② 建议线：只提醒不拒绝
    const hinted = await memory.mutate({ action: 'add', target: 'project', content: `这条在建议线之上：${'字'.repeat(60)}` });
    assert.ok(hinted.advisory !== null, '过建议线要给 advisory');
    assert.equal(hinted.advisory.hint, 120, 'advisory 要带上建议线本身');
    assert.ok(renderMutation(hinted).includes('篇幅提醒'), '回执要提醒');
    assert.ok(renderMutation(hinted).includes('kb'), '提醒里要给出长内容的去处');

    // ③ 改短不再提醒（否则每次精简都要被念一遍）
    const shortened = await memory.mutate({ action: 'replace', id: hinted.entry.id, content: '压短了。' });
    assert.equal(shortened.advisory, null, '改短不该再提醒');

    // ④ 0 = 关掉两道门：有人就是要把长篇结论放热记忆里
    const off = await freshRoot({ projectLimitBytes: 20_000, entryLimitBytes: 0, entryHintBytes: 0 });
    const big = await off.memory.mutate({ action: 'add', target: 'project', content: 'z'.repeat(4000) });
    assert.equal(big.advisory, null, '关掉后既不拒绝也不提醒');
    assert.equal(big.entry.content.length, 4000, '关掉后长条目要真的写进去');

    // ⑤ 硬上限拦住的是**写入**，不改存量：老库里已有的长条目照样读得出来（不回溯清理）
    const legacy = await freshRoot({ projectLimitBytes: 20_000, entryLimitBytes: 300 });
    await assert.rejects(() => legacy.memory.mutate({ action: 'add', target: 'project', content: 'q'.repeat(400) }));
    const snapshot = await legacy.memory.snapshot();
    assert.equal(snapshot.entries.length, 0, '被拒的条目不该留下任何痕迹');
});


await check('台账：登记后可读，字段完整', async () => {
    const { root, memory } = await freshRoot();
    const value = await memory.log([{
        path: '季度汇报.pptx',
        format: 'ppt',
        theme: 'business',
        bytes: 22585,
        purpose: '向管理层汇报 Q1',
        outline: ['封面', '整体情况', '分月数据'],
        stats: { slides: 12 },
    }]);
    assert.deepEqual([value.added, value.kept, value.rolled], [1, 1, 0]);
    const ledger = await memory.read({ layer: 'ledger' });
    assert.equal(ledger.ledger.total, 1);
    assert.equal(ledger.ledger.items[0].path, '季度汇报.pptx');
    assert.ok(ledger.ledger.items[0].text.includes('向管理层汇报 Q1'));
    // jsonl 是一行一条：追加式，坏一行不影响其它行。
    const lines = (await readFile(memoryPaths(root, {}).ledger, 'utf8')).trim().split('\n');
    assert.equal(lines.length, 1);
    assert.equal(JSON.parse(lines[0]).path, '季度汇报.pptx');
});

await check('台账空白记录被丢弃（不给空路径建索引）', async () => {
    const { memory } = await freshRoot();
    const value = await memory.log([{ path: '   ' }, { path: '' }]);
    assert.equal(value.added, 0);
});

await check('台账滚动：超过条数上限时最旧的滚进归档，文件里不再重复', async () => {
    const { root, memory } = await freshRoot({ ledgerLimit: 3, archiveKeep: 10 });
    for (let i = 0; i < 6; i += 1) {
        await memory.log([{ path: `产出-${i}.docx`, format: 'word', at: `2026-0${1 + i}-01T00:00:00.000Z` }]);
    }
    const ledger = await memory.read({ layer: 'ledger' });
    assert.equal(ledger.ledger.total, 3, `台账应只留 3 条，实际 ${ledger.ledger.total}`);
    const lines = (await readFile(memoryPaths(root, {}).ledger, 'utf8')).trim().split('\n');
    assert.equal(lines.length, 3, 'jsonl 里也只该有 3 行');
    const paths = lines.map((line) => JSON.parse(line).path);
    assert.deepEqual(paths, ['产出-3.docx', '产出-4.docx', '产出-5.docx']);
    const archive = await memory.read({ layer: 'archive', query: '产出-0' });
    assert.ok(archive.archive.items.length >= 1, '滚动掉的旧记录要能在归档里检索到');
});

await check('归档保留上限：超出时最旧的摘要连文件一起删，索引同步', async () => {
    const { root, memory } = await freshRoot({ ledgerLimit: 1, archiveKeep: 2 });
    for (let i = 0; i < 4; i += 1) {
        await memory.log([{ path: `第${i}月.docx`, format: 'word', at: `2026-0${1 + i}-01T00:00:00.000Z` }]);
    }
    const paths = memoryPaths(root, {});
    const index = JSON.parse(await readFile(paths.index, 'utf8'));
    const months = index.digests.map((digest) => digest.month);
    assert.equal(months.length, 2, `归档应只留 2 个摘要，实际 ${months.join('、')}`);
    // 四次登记里，第一次还不满一条上限不会滚动，所以产生的摘要是 01/02/03 三个月，
    // 保留最新的两个；被挤掉的那个摘要连文件一起删。
    assert.deepEqual(months, ['2026-02', '2026-03']);
    assert.equal(existsSync(join(paths.archive, '2026-01.md')), false, '被挤掉的摘要文件要删掉');
    const archive = await memory.read({ layer: 'archive' });
    assert.equal(archive.archive.digests, 2);
});

// ── 有界读取 ────────────────────────────────────────────────────────────────

await check('读取有界：条数上限之外只报总数，并标 truncated', async () => {
    const { memory } = await freshRoot();
    for (let i = 0; i < 20; i += 1) {
        await memory.log([{ path: `文档-${i}.docx`, format: 'word', at: `2026-05-${String(i + 1).padStart(2, '0')}T00:00:00.000Z` }]);
    }
    const value = await memory.read({ layer: 'ledger', limit: 3 });
    assert.equal(value.ledger.total, 20, '总数照实报');
    assert.equal(value.ledger.items.length, 3, '一次只给 3 条');
    assert.equal(value.truncated, true, '被截断要明确标出来');
    const text = renderRead(value);
    assert.ok(text.includes('共 20 条'));
});

await check('读取有界：单条过长被截断，条目本身不被改写', async () => {
    const { root, memory } = await freshRoot();
    const long = `长条目：${'字'.repeat(900)}`;
    await memory.mutate({ action: 'add', target: 'project', content: long });
    const value = await memory.read({ layer: 'hot' });
    assert.equal(value.hot.items.length, 1);
    assert.ok(value.hot.items[0].text.length < long.length, '返回给模型的文本应被截断');
    assert.equal(value.hot.items[0].truncated, true);
    // 存储里的原文不动。
    const snapshot = await memory.snapshot();
    assert.equal(snapshot.entries[0].content, long);
    assert.equal((await memory.read({ layer: 'hot' })).truncated, true);
});

await check('检索：多个关键词命中越多排越前，未命中不返回', async () => {
    const { memory } = await freshRoot();
    await memory.log([{ path: '季度汇报.pptx', format: 'ppt', purpose: '季度业绩汇报', outline: ['封面'] }]);
    await memory.log([{ path: '年度汇报.pptx', format: 'ppt', purpose: '年度业绩汇报', outline: ['封面'] }]);
    await memory.log([{ path: '会议纪要.docx', format: 'word', purpose: '周会记录' }]);
    const value = await memory.read({ layer: 'ledger', query: '季度 汇报' });
    // 命中规则是「命中任意一个词就进候选，命中越多排越前」：两个词都命中的那份排第一，
    // 只命中「汇报」的紧随其后，一个词都不命中的不返回。
    assert.equal(value.ledger.matched, 2);
    assert.equal(value.ledger.items[0].path, '季度汇报.pptx');
    assert.ok(!value.ledger.items.some((item) => item.path === '会议纪要.docx'));
    const none = await memory.read({ layer: 'ledger', query: '不存在的主题词' });
    assert.equal(none.ledger.items.length, 0);
});

await check('读取去重：同内容的条目只出现一次', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '重复的约定内容。' });
    await memory.mutate({ action: 'add', target: 'project', content: '重复的约定内容。' });
    const value = await memory.read({ layer: 'hot' });
    assert.equal(value.hot.items.length, 1, '摘要里同样的内容只给一次');
});

await check('三层共用一份读取预算（layer:all 不会把三层各自灌满）', async () => {
    const { memory } = await freshRoot();
    for (let i = 0; i < 12; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `分层条目 ${i}：${'层'.repeat(60)}` });
        await memory.log([{ path: `分层产出-${i}.docx`, format: 'word', purpose: `第 ${i} 份${'产'.repeat(60)}` }]);
    }
    const value = await memory.read({ layer: 'all' });
    const totalChars = [...(value.hot?.items ?? []), ...(value.ledger?.items ?? []), ...(value.archive?.items ?? [])]
        .reduce((sum, item) => sum + String(item.text ?? '').length, 0);
    assert.ok(totalChars <= 4000, `三层合计字符应在共享预算内，实际 ${totalChars}`);
});

// ── 并发与统计 ──────────────────────────────────────────────────────────────

await check('并发写入串行化：十条同时 add 一条都不丢', async () => {
    const { memory } = await freshRoot();
    await Promise.all(Array.from({ length: 10 }, (_, i) => memory.mutate({ action: 'add', target: 'project', content: `并发条目 ${i}` })));
    const snapshot = await memory.snapshot();
    assert.equal(snapshot.entries.length, 10);
});

await check('memoryStats 给出目录、体积与文件数', async () => {
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'user', content: '偏好：短句。' });
    await memory.log([{ path: 'a.docx', format: 'word' }]);
    const stats = await memoryStats(root, {});
    assert.equal(stats.exists, true);
    assert.ok(stats.files >= 3, `至少有 memory.json / USER.md / ledger.jsonl，实际 ${stats.files}`);
    assert.ok(stats.bytes > 0);
    assert.equal(memoryExists(root), true);
});

await check('全局层：新路径不存在而旧的 .office-memory 还在时退回旧路径', async () => {
    const home = join(TMP, 'dsh-home-fallback');
    await rm(home, { recursive: true, force: true });
    await mkdir(join(home, '.office-memory'), { recursive: true });
    const saved = process.env.DSH_HOME;
    process.env.DSH_HOME = home;
    try {
        assert.equal(globalMemoryDir({}), join(home, '.office-memory'),
            '旧目录还在、新目录还没建时，不能让用户已有的全局记忆变成孤儿');
        // 新目录一出现就只认新路径：不存在的旧目录不再参与判断。
        await mkdir(join(home, '.office', 'memory'), { recursive: true });
        assert.equal(globalMemoryDir({}), join(home, '.office', 'memory'));
        // 显式配置永远优先，两种写法都不做兜底。
        assert.equal(globalMemoryDir({ globalDir: 'custom-mem' }), join(homedir(), 'custom-mem'));
        assert.equal(globalMemoryDir({ globalDir: join(home, 'abs-mem') }), join(home, 'abs-mem'));
    } finally {
        if (saved === undefined) delete process.env.DSH_HOME;
        else process.env.DSH_HOME = saved;
    }
});

// ── 渲染 ────────────────────────────────────────────────────────────────────

await check('renderDigest：空态给「怎么记」，非空态给偏好、项目与最近台账', async () => {
    const { memory } = await freshRoot();
    const empty = renderDigest(await memory.digest());
    assert.ok(empty.includes('还没有内容'));
    assert.ok(empty.includes("target:'user'") || empty.includes("target: 'user'"));

    await memory.mutate({ action: 'add', target: 'user', content: '偏好：书面化。' });
    await memory.mutate({ action: 'add', target: 'project', content: '约定：用 business 主题。' });
    await memory.log([{ path: '季度汇报.pptx', format: 'ppt', theme: 'business' }]);
    const filled = renderDigest(await memory.digest({ recentLedger: 2 }));
    assert.ok(filled.includes('偏好：书面化'));
    assert.ok(filled.includes('约定：用 business 主题'));
    assert.ok(filled.includes('季度汇报.pptx'));
    assert.ok(filled.includes('office_memory'));
});

await check('renderDigest：给「未变化」形态时省掉正文，但保留用量与台账', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'user', content: '偏好：书面化。' });
    // 造一份**像真实热记忆那样有体积**的内容（本机实测项目层约 10 KB）：
    // 省略的收益与正文体积成正比 —— 只有两句话时，省略形态的说明行反而更长。
    // 那个方向也是对的（小到一定程度就不该省），但断言要按真实量级来立。
    for (let index = 0; index < 12; index += 1) {
        await memory.mutate({
            action: 'add', target: 'project',
            content: `约定 ${index}：`
                + '这条约定用来把热记忆的正文撑到真实量级，好让「省略正文」的收益能被断言成大小关系而不是感觉。'.repeat(6),
        });
    }
    await memory.log([{ path: '季度汇报.pptx', format: 'ppt', theme: 'business' }]);
    const digest = await memory.digest({ recentLedger: 2 });

    const full = renderDigest(digest, { context: 'help' });
    const omitted = renderDigest(digest, { context: 'help', hot: 'unchanged' });
    assert.ok(full.length > 2500, `夹具本身要有真实量级：${full.length}`);
    assert.ok(omitted.length * 4 < full.length, `省略形态应当明显更短：${omitted.length} vs ${full.length}`);
    assert.ok(!omitted.includes('偏好：书面化'), '省略形态不带热记忆正文');
    assert.ok(!omitted.includes('约定 3：'), '省略形态不带热记忆正文');
    assert.ok(omitted.includes('与上一次投影相同'), '要说清为什么没有正文');
    assert.ok(omitted.includes(String(digest.writeRevision).slice(0, 8)), '要带上 revision，省略才可核对');
    assert.ok(omitted.includes("office_memory({ action: 'read', layer: 'hot' })"), '要给读全文的入口');
    // 用量与台账照旧（ledger 形态默认 'full'；第三十六轮起台账也能折叠，见下一组用例）。
    assert.ok(omitted.includes('用户偏好 1 条'), '用量行要保留');
    assert.ok(omitted.includes('项目与环境 12 条'), '用量行要保留');
    assert.ok(omitted.includes('季度汇报.pptx'), '台账部分在 ledger:\'full\' 时照旧给');
    assert.ok(omitted.includes('主动记录'), '主动记录指引照旧');
});

// ── 投影：只在变化时贴 ──────────────────────────────────────────────────────

await check('投影：同一会话 writeRevision 不变时只给一行，变了又贴全文', async () => {
    resetProjection();
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '约定：主题用 business。' });
    const digest = await memory.digest();

    const first = projectDigest(digest, { sessionId: 's-1', context: 'help' });
    assert.equal(first.mode, 'full', '会话里第一次投影必须贴全文');
    assert.equal(first.reason, 'changed');
    assert.ok(first.text.includes('约定：主题用 business'));

    const second = projectDigest(digest, { sessionId: 's-1', context: 'help' });
    assert.equal(second.mode, 'unchanged');
    assert.ok(!second.text.includes('约定：主题用 business'), '没变就不重复贴正文');
    assert.ok(second.text.includes(String(digest.writeRevision).slice(0, 8)));

    // 热记忆变了（写了一次 → writeRevision 变）→ 又贴全文
    await memory.mutate({ action: 'replace', target: 'project', oldText: '主题用 business', content: '约定：主题用 business，标题 28pt。' });
    const changed = projectDigest(await memory.digest(), { sessionId: 's-1', context: 'help' });
    assert.equal(changed.mode, 'full');
    assert.equal(changed.reason, 'changed');
    assert.ok(changed.text.includes('标题 28pt'));
});

await check('投影：会话之间互不影响，拿不到会话时每次都贴全文', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '一条约定。' });
    const digest = await memory.digest();

    assert.equal(projectDigest(digest, { sessionId: 's-a' }).mode, 'full');
    // 另一个会话没看过，不能因为 s-a 贴过就跟着省略 —— 那才是真的丢信息。
    assert.equal(projectDigest(digest, { sessionId: 's-b' }).mode, 'full');
    assert.equal(projectDigest(digest, { sessionId: 's-a' }).mode, 'unchanged');

    // 没有会话身份（测试里的 execAt 就没有 id）：不假装「刚贴过」。
    for (let index = 0; index < 3; index += 1) {
        const out = projectDigest(digest, { sessionId: '' });
        assert.equal(out.mode, 'full');
        assert.equal(out.reason, 'no-session');
    }
    // exec.agent.session 整个缺失也一样。
    assert.equal(sessionIdOf(undefined), '');
    assert.equal(sessionIdOf({}), '');
    assert.equal(sessionIdOf({ agent: { session: { header: { cwd: 'x' } } } }), '', '只有 cwd 没有 id 也算拿不到');
    assert.equal(sessionIdOf({ agent: { session: { id: 's-z' } } }), 's-z');
});

await check('投影：连续省略到 REFRESH_EVERY 次会强制重贴一次（防宿主压缩裁掉）', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '一条约定。' });
    const digest = await memory.digest();

    assert.equal(projectDigest(digest, { sessionId: 's-r' }).reason, 'changed');
    const reasons = [];
    for (let index = 0; index < REFRESH_EVERY; index += 1) {
        reasons.push(projectDigest(digest, { sessionId: 's-r' }).reason);
    }
    // 前 REFRESH_EVERY-1 次省略，第 REFRESH_EVERY 次强制重贴，然后计数归零重新开始。
    assert.deepEqual(reasons.slice(0, REFRESH_EVERY - 1), Array(REFRESH_EVERY - 1).fill('unchanged'));
    assert.equal(reasons[REFRESH_EVERY - 1], 'refresh');
    assert.equal(projectionStateOf('s-r').since, 0, '重贴之后计数归零');
    assert.equal(projectDigest(digest, { sessionId: 's-r' }).reason, 'unchanged');
});

await check('投影：空记忆不做省略（那段本来就小，而且讲的正是「怎么记」）', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    const digest = await memory.digest();
    assert.equal(digest.empty, true);
    for (let index = 0; index < 3; index += 1) {
        const out = projectDigest(digest, { sessionId: 's-empty' });
        assert.equal(out.mode, 'full');
        assert.equal(out.reason, 'empty');
        assert.ok(out.text.includes('还没有内容'));
    }
});

await check('投影：会话状态是有界的，不会随长跑进程无限增长', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '一条约定。' });
    const digest = await memory.digest();
    for (let index = 0; index < 200; index += 1) projectDigest(digest, { sessionId: 's-' + index });
    // 最旧的已经被淘汰：再问它一次会当作「没看过」→ 贴全文。
    assert.equal(projectionStateOf('s-0'), null, '超出上限的会话状态要被淘汰');
    assert.notEqual(projectionStateOf('s-199'), null, '最近的会话状态还在');
    assert.equal(projectDigest(digest, { sessionId: 's-199' }).mode, 'unchanged');
});

await check('投影：office_help 第二次调用不再重复贴热记忆正文', async () => {
    resetProjection();
    const { root } = await freshRoot();
    // 夹具要落在 word 的话题词上：office_help({topic:'word'}) 的投影按话题筛，
    // 与话题无关的条目本来就该折叠（那是第三十六轮的按需投影，不是丢失）。
    await createMemory({ root }).mutate({ action: 'add', target: 'project', content: '约定：word 文档的投影不要重复贴。' });
    await createMemory({ root }).mutate({ action: 'add', target: 'project', content: '约定：检索来源给 URL。' });
    const tools = buildTools(resolveConfig({ memory: {} }));
    const help = tools.find((item) => item.name === 'office_help');
    // 带 id 的会话：投影状态按它记。
    const exec = { agent: { session: { id: 's-help', header: { cwd: root } } } };

    const first = await help.execute({ topic: 'word' }, exec);
    assert.ok(first.text.includes('约定：word 文档的投影不要重复贴'), '第一次要带正文');
    assert.ok(!first.text.includes('约定：检索来源给 URL'), '与话题无关的条目第一次就折叠');
    assert.ok(first.text.includes('另有 1 条与「word」无关'), '折叠行要报条数');
    const second = await help.execute({ topic: 'word' }, exec);
    assert.ok(!second.text.includes('约定：word 文档的投影不要重复贴'), '第二次不该再贴一遍正文');
    assert.ok(second.text.includes('与上一次投影相同'), '第二次给省略形态');
    assert.ok(second.text.includes('office.word'), '文档本身照旧完整返回');

    // 换一个会话：又要贴全文（它没看过）。
    const other = await help.execute({ topic: 'word' }, { agent: { session: { id: 's-help-2', header: { cwd: root } } } });
    assert.ok(other.text.includes('约定：word 文档的投影不要重复贴'));
});

// ── 按需投影（第三十六轮）：需要用什么才给什么 ─────────────────────────────

await check('signalOf：topic 与 script 都能给出信号，认不出就返回 null（不猜）', async () => {
    assert.ok(signalOf({ topic: 'ppt' }).terms.includes('ppt'));
    assert.ok(signalOf({ topic: 'ppt' }).terms.includes('幻灯'), '词表是中文正文的词，不只是 API 名');
    assert.equal(signalOf({ topic: '' }), null, '空话题没有信号');
    assert.equal(signalOf({ topic: 'guide' }), null, '总览话题不筛');
    assert.equal(signalOf({ topic: '不存在的话题' }), null);

    const fromScript = signalOf({ script: "const d = office.ppt.create(); d.save('a.pptx'); return 'ok';" });
    assert.equal(fromScript.key, 'ppt', '脚本里的 office.ppt 与 .pptx 都是信号');
    const mixed = signalOf({ script: "office.pdf.text('x.pdf'); const w = office.word.create({});" });
    assert.equal(mixed.key, 'pdf+word', '多格式取并集，键要稳定（排序）');
    assert.equal(signalOf({ script: 'return 1 + 1;' }), null, '没有 office.* 与扩展名的脚本不猜');
});

await check('按需投影：带信号只贴命中条目，无关条目折叠但留入口与用量', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    // 夹具要像真实热记忆那样有体积（本机实测单条 300–1500 B）：折叠行是固定开销，
    // 条目太短时「省下的正文」会被折叠行抵掉，大小断言必须按真实量级立。
    const filler = '这条约定用来把正文撑到真实量级，让「折叠无关条目」的收益能被断言成大小关系而不是感觉。'.repeat(6);
    await memory.mutate({ action: 'add', target: 'project', content: `约定：PPT 版面用 16:9。${filler}` });
    await memory.mutate({ action: 'add', target: 'project', content: `约定：检索来源必须给 URL。${filler}` });
    await memory.mutate({ action: 'add', target: 'project', content: `约定：检索来源要可复核。${filler}` });
    await memory.mutate({ action: 'add', target: 'project', content: `约定：检索写完要按渠道核对。${filler}` });
    const digest = await memory.digest();

    const full = renderDigest(digest, { context: 'help' });
    const filtered = renderDigest(digest, {
        context: 'help',
        signal: { label: 'ppt', terms: signalOf({ topic: 'ppt' }).terms },
    });
    assert.ok(full.includes('检索来源必须给 URL'));
    assert.ok(filtered.includes('PPT 版面用 16:9'), '命中的条目要贴正文');
    assert.ok(!filtered.includes('检索来源必须给 URL'), '无关条目不贴正文');
    assert.ok(filtered.includes('另有 3 条与「ppt」无关'), '折叠行要报条数');
    assert.ok(filtered.includes("office_memory({ action: 'read', layer: 'hot' })"), '折叠行要留读取入口');
    assert.ok(filtered.includes('项目与环境 4 条'), '用量行给全量，不是「这次给了几条」');
    assert.ok(filtered.length * 2 < full.length, `信号过滤后的投影必须明显短于全文：${filtered.length} vs ${full.length}`);
});

await check('按需投影：一条都不命中时如实说 0 条，不静默装满也不装空', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'user', content: '偏好：书面化。' });
    const digest = await memory.digest();
    const none = renderDigest(digest, {
        context: 'help',
        signal: { label: 'tex', terms: signalOf({ topic: 'tex' }).terms },
    });
    assert.ok(none.includes('与「tex」相关的条目 0 条'), '要说清没有命中');
    assert.ok(none.includes("office_memory({ action: 'read', layer: 'hot' })"), '仍然留读取入口');
    assert.ok(none.includes('用户偏好 1 条'), '用量行照旧');
});

await check('按需投影：换话题只贴「这次还没贴过」的条目，A→B→A 不再重付（P0-2）', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '约定：PPT 版面 16:9。' });
    await memory.mutate({ action: 'add', target: 'project', content: '约定：检索来源给 URL。' });
    const digest = await memory.digest();
    const ppt = signalOf({ topic: 'ppt' });
    const search = signalOf({ topic: 'search' });

    const first = projectDigest(digest, { sessionId: 's-sig', context: 'help', signal: ppt });
    assert.equal(first.reason, 'changed');
    assert.ok(first.text.includes('PPT 版面 16:9'), '本话题命中的条目要贴');
    assert.ok(!first.text.includes('检索来源给 URL'), '另一话题的条目这次折叠');
    const again = projectDigest(digest, { sessionId: 's-sig', context: 'help', signal: ppt });
    assert.equal(again.mode, 'unchanged', '同一话题、writeRevision 没变就省略');

    // 换话题：新话题下还有没贴过的条目 → 贴；这才是「换话题要重贴」的真实含义
    const switched = projectDigest(digest, { sessionId: 's-sig', context: 'help', signal: search });
    assert.equal(switched.mode, 'full', '换到没贴过的话题要贴正文');
    assert.equal(switched.reason, 'changed');
    assert.ok(switched.text.includes('检索来源给 URL'), '新话题的条目这次要贴');

    // A→B→A：切回 A 时 A 已经贴过，不再重付一份（P0-2 的节流：状态记的是「贴过哪些条目」，
    // 不是「上一次是什么话题」）
    const back = projectDigest(digest, { sessionId: 's-sig', context: 'help', signal: ppt });
    assert.equal(back.mode, 'unchanged', '切回已贴过的话题不再重贴');
    assert.ok(!back.text.includes('PPT 版面 16:9'), '已贴过的正文不重复贴');
});

await check('按需投影：心跳重贴不筛（完整性兜底，把被折叠的条目带回来）', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '约定：PPT 版面 16:9。' });
    await memory.mutate({ action: 'add', target: 'project', content: '约定：检索来源给 URL。' });
    const digest = await memory.digest();
    const ppt = signalOf({ topic: 'ppt' });

    const first = projectDigest(digest, { sessionId: 's-heart', signal: ppt });
    assert.ok(!first.text.includes('检索来源给 URL'), '信号过滤后无关条目折叠');

    let last = first;
    for (let index = 0; index < REFRESH_EVERY; index += 1) {
        last = projectDigest(digest, { sessionId: 's-heart', signal: ppt });
    }
    assert.equal(last.reason, 'refresh');
    assert.ok(last.text.includes('检索来源给 URL'), '心跳重贴必须不筛：先前折叠的条目要能回到上下文');
});

await check('台账投影：指纹没变就折叠，登记新产物后才重贴', async () => {
    resetProjection();
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '约定：一条。' });
    const first = projectDigest(await memory.digest(), { sessionId: 's-led' });
    assert.ok(!first.text.includes('【台账】'), '没有台账记录时不渲染台账段');

    await memory.log([{ path: 'a.docx', format: 'word', purpose: '第一份' }]);
    const second = projectDigest(await memory.digest(), { sessionId: 's-led' });
    assert.equal(second.reason, 'ledger', '台账指纹变了要贴（热记忆没变也贴）');
    assert.ok(second.text.includes('a.docx'), '新登记的台账行要贴出来');

    const third = projectDigest(await memory.digest(), { sessionId: 's-led' });
    assert.equal(third.mode, 'unchanged', '指纹没变，台账与热记忆都折叠');
    assert.ok(!third.text.includes('a.docx'), '台账行不重复贴');
    assert.ok(third.text.includes('与上次投影相同'), '折叠要说清并给入口');
    assert.ok(third.text.includes("office_memory({ action: 'read' })"), '读全文的入口在');
});

await check('renderDigest：台账省略形态给一行说明与读取入口，固定尾巴不再重复', async () => {
    const { memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'user', content: '偏好：书面化。' });
    await memory.log([{ path: '季度汇报.pptx', format: 'ppt', theme: 'business' }]);
    const digest = await memory.digest({ recentLedger: 2 });

    const both = renderDigest(digest, { context: 'help', hot: 'unchanged', ledger: 'unchanged' });
    assert.ok(!both.includes('季度汇报.pptx'), '台账行不贴');
    assert.ok(both.includes('最近 1 条与上次投影相同'), '要说清台账为什么没有正文');
    assert.ok(both.includes("office_memory({ action: 'read', layer: 'ledger' })"), '要给台账的读取入口');
    assert.ok(!both.includes('主动记录'), '热记忆与台账都省略时，固定尾巴不再重复');
    assert.ok(!both.includes('完整读取与检索'), '两条尾巴一起省');
});

await check('renderMutation：给回执，不回显整份文件', async () => {
    const { memory } = await freshRoot();
    const value = await memory.mutate({ action: 'add', target: 'project', content: '一条约定' });
    const text = renderMutation(value);
    assert.ok(text.includes('记忆已更新'));
    assert.ok(text.includes('新增'));
    assert.ok(!text.includes('# 办公记忆'), '不该把投影整份贴回来');
});

// ── 与工具面的接线 ──────────────────────────────────────────────────────────

await check('office_memory 工具：增 → 查 → 改 → 删全链路', async () => {
    const { root } = await freshRoot();
    const tools = buildTools(resolveConfig({ memory: {} }));
    const tool = tools.find((item) => item.name === 'office_memory');
    assert.ok(tool, '默认工具面要有 office_memory');
    const exec = execAt(root);

    let out = await tool.execute({ action: 'add', target: 'user', content: '偏好：汇报先给结论。', importance: 'critical' }, exec);
    assert.ok(out.ok && out.text.includes('记忆已更新'), out.text);
    out = await tool.execute({ action: 'read', layer: 'hot' }, exec);
    assert.ok(out.text.includes('汇报先给结论'));
    out = await tool.execute({ action: 'replace', target: 'user', content: '偏好：汇报先给结论，再给数据。', oldText: '先给结论。' }, exec);
    assert.ok(out.text.includes('改后'), out.text);
    out = await tool.execute({ action: 'remove', target: 'user', oldText: '先给结论' }, exec);
    assert.ok(out.text.includes('移除'), out.text);
    out = await tool.execute({ action: 'read' }, exec);
    assert.ok(!out.text.includes('先给结论'));
});

await check('office_memory 工具：手工登记台账与查询', async () => {
    const { root } = await freshRoot();
    const tool = buildTools(resolveConfig({ memory: {} })).find((item) => item.name === 'office_memory');
    const exec = execAt(root);
    let out = await tool.execute({ action: 'log', path: '用户给的参考稿.pptx', format: 'ppt', purpose: '版式参考', source: ['需求邮件'] }, exec);
    assert.ok(out.text.includes('台账已登记 1 条'), out.text);
    out = await tool.execute({ action: 'read', layer: 'ledger', query: '参考稿' }, exec);
    assert.ok(out.text.includes('用户给的参考稿.pptx'), out.text);
});

await check('记忆关闭时工具给出可读提示，而不是静默失败', async () => {
    const { root } = await freshRoot();
    const tool = buildTools(resolveConfig({ memory: { enabled: false } })).find((item) => item.name === 'office_memory');
    const out = await tool.execute({ action: 'read' }, execAt(root));
    assert.equal(out.ok, false);
    assert.ok(out.text.includes('记忆已关闭'));
});

await check('office_help 末尾带热记忆投影（动笔前那一次调用）', async () => {
    const { root } = await freshRoot();
    const config = resolveConfig({ memory: {} });
    const tools = buildTools(config);
    const help = tools.find((item) => item.name === 'office_help');
    const exec = execAt(root);

    let out = await help.execute({ topic: 'ppt' }, exec);
    assert.ok(out.text.includes('📒 记忆（还没有内容'), '空态也要告诉模型怎么记');
    assert.ok(out.text.includes('office.ppt'), '原文档不能被记忆段落顶掉');

    await createMemory({ root, memory: config.memory }).mutate({ action: 'add', target: 'project', content: '本项目 PPT 一律 16:9。' });
    out = await help.execute({ topic: 'ppt' }, exec);
    assert.ok(out.text.includes('本项目 PPT 一律 16:9'), '动笔前应看到项目约定');
    assert.ok(out.text.endsWith('office_memory') || out.text.includes('office_memory'), '投影后要给出读取入口');
});

await check('记忆目录不可用时，office_help 仍然照常返回文档', async () => {
    const { root } = await freshRoot();
    // 把记忆目录占成文件：目录建不出来，记忆读盘必然失败。
    // 记忆目录现在是 .office/memory，所以 .office 得先是个目录，再占住 memory 这一层。
    await mkdir(join(root, '.office'), { recursive: true });
    await writeFile(join(root, '.office', 'memory'), 'not a directory', 'utf8');
    const help = buildTools(resolveConfig({ memory: {} })).find((item) => item.name === 'office_help');
    const out = await help.execute({ topic: 'files' }, execAt(root));
    assert.ok(out.ok);
    assert.ok(out.text.includes('office.files.read'), '记忆失败不该影响用法查询');
});

await check('office_run 自动登记台账：真实产物、真实统计', async () => {
    const { root } = await freshRoot();
    const config = resolveConfig({ memory: {} });
    const run = buildTools(config).find((item) => item.name === 'office_run');
    const exec = execAt(root);
    const out = await run.execute({
        script: "const w = office.word.create({ title: '记忆探针' }); w.title('记忆探针').heading('一', 1).para('正文一段。').save('记忆探针.docx'); return 'ok';",
        purpose: '验证自动台账',
    }, exec);
    assert.equal(out.ok, true, JSON.stringify(out.error ?? {}));
    assert.equal(out.memory.enabled, true);
    assert.equal(out.memory.logged, 1, `应登记一条，实际 ${JSON.stringify(out.memory)}`);
    assert.equal(out.memory.kept, 1);

    const ledger = await createMemory({ root, memory: config.memory }).read({ layer: 'ledger' });
    assert.equal(ledger.ledger.total, 1);
    const record = JSON.parse((await readFile(memoryPaths(root, {}).ledger, 'utf8')).trim());
    assert.equal(record.path, '记忆探针.docx');
    assert.equal(record.format, 'word');
    assert.equal(record.purpose, '验证自动台账');
    assert.ok(record.outline.length > 0, '台账要带结构摘要');
    assert.ok(record.stats !== null, '台账要带复检统计');
    // 反馈里要顺带把记忆投影带回来（这次调用之后该看到的东西）。
    assert.ok(out.memory.digest === undefined || typeof out.memory.digest === 'string');
});

await check('office_run 自动登记：source 记这次读过的输入文件，告警折进 note', async () => {
    // 第十九轮 P1-4 / 第二十四轮 P1-3：台账的检索键排在 outline + source + note 上，
    // 而自动登记当初一个 source 都不传 —— 「找上次那份东西」退化成按文件名回忆。
    const { root } = await freshRoot();
    await writeFile(join(root, '素材.txt'), '季度数据：一二三。', 'utf8');
    await writeFile(join(root, '无关.txt'), '这份没被读过。', 'utf8');
    const config = resolveConfig({ memory: {} });
    const run = buildTools(config).find((item) => item.name === 'office_run');
    const out = await run.execute({
        script: "const text = office.files.read('素材.txt');"
            + " const w = office.word.create({ title: '来源探针' }); w.title('来源探针').para(text).save('来源探针.docx');"
            + " return 'ok';",
        purpose: '验证 source 与 note',
    }, execAt(root));
    assert.equal(out.ok, true, JSON.stringify(out.error ?? {}));
    const records = (await readFile(memoryPaths(root, {}).ledger, 'utf8')).trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(records.length, 1);
    assert.deepEqual(records[0].source, ['素材.txt'], '只记真的读过内容的那一份');
    assert.ok(!records[0].source.includes('无关.txt'), '没读过的文件不许当来源');
    assert.ok(!records[0].source.includes('来源探针.docx'), '自己写出来的文件不算自己的来源');

    // 告警折进 note：用「插了目录域但一个 H1–H3 都没有」这条确定性告警（word.js:1922）。
    const warned = await run.execute({
        script: "const w = office.word.create({ title: '告警探针' }); w.para('正文。').toc().save('告警探针.docx'); return 'ok';",
        purpose: '验证告警进 note',
    }, execAt(root));
    assert.equal(warned.ok, true, JSON.stringify(warned.error ?? {}));
    const after = (await readFile(memoryPaths(root, {}).ledger, 'utf8')).trim().split('\n').map((line) => JSON.parse(line));
    const probe = after.find((record) => record.path === '告警探针.docx');
    assert.ok(probe, '第二次也要登记一条');
    assert.ok(String(probe.note).includes('目录'), `告警要折进 note，实际「${probe.note}」`);
    assert.equal(probe.source.length, 0, '这次没读任何输入文件，source 就是空的（如实，不是漏记）');
    // note 参与检索：用告警里的词能查到这一条。
    const hit = await createMemory({ root, memory: config.memory }).read({ layer: 'ledger', query: '目录域' });
    assert.equal(hit.ledger.items.length, 1, 'note 要能被检索命中');
});

await check('office_run 只登记 Office 产物：普通文本文件不进台账', async () => {
    const { root } = await freshRoot();
    const config = resolveConfig({ memory: {} });
    const run = buildTools(config).find((item) => item.name === 'office_run');
    const out = await run.execute({ script: "await office.files.write('笔记.md', '# 标题\\n'); return 'ok';" }, execAt(root));
    assert.equal(out.ok, true, JSON.stringify(out.error ?? {}));
    assert.equal(out.otherFiles.length, 1);
    assert.equal(out.memory.logged, 0);
    assert.equal((await createMemory({ root, memory: config.memory }).read({ layer: 'ledger' })).ledger.total, 0);
});

await check('关掉 autoLedger 后 office_run 不再登记，但记忆投影仍在', async () => {
    const { root } = await freshRoot();
    const config = resolveConfig({ memory: { autoLedger: false } });
    await createMemory({ root, memory: config.memory }).mutate({ action: 'add', target: 'project', content: '既有约定：word 标题从简。' });
    const run = buildTools(config).find((item) => item.name === 'office_run');
    const out = await run.execute({
        script: "const w = office.word.create({ title: 'x' }); w.title('x').save('x.docx'); return 'ok';",
    }, execAt(root));
    assert.equal(out.ok, true, JSON.stringify(out.error ?? {}));
    assert.equal(out.memory.logged, 0);
    assert.equal((await createMemory({ root, memory: config.memory }).read({ layer: 'ledger' })).ledger.total, 0);
    assert.ok(String(out.memory.digest ?? '').includes('既有约定'), '投影仍要带回来');
});

await check('记忆目录被占成文件时，office_run 照常给结论，只报台账写入失败', async () => {
    const { root } = await freshRoot();
    await mkdir(join(root, '.office'), { recursive: true });
    await writeFile(join(root, '.office', 'memory'), 'not a directory', 'utf8');
    const config = resolveConfig({ memory: {} });
    const run = buildTools(config).find((item) => item.name === 'office_run');
    const out = await run.execute({
        script: "const w = office.word.create({ title: 'y' }); w.title('y').save('y.docx'); return 'ok';",
    }, execAt(root));
    assert.equal(out.ok, true, '文件写出来了，结论就该是成功');
    assert.equal(out.files.length, 1);
    assert.equal(typeof out.memory.error, 'string');
    const { renderRun } = await import('../src/tools.js');
    assert.ok(renderRun(out).includes('台账写入失败'), '反馈里要说明记忆没写成');
});

await check('配置里的记忆目录会被工具与引擎用上（不是只认默认目录）', async () => {
    const { root } = await freshRoot();
    const config = resolveConfig({ memory: { dir: '.my-office-memory' } });
    const tool = buildTools(config).find((item) => item.name === 'office_memory');
    await tool.execute({ action: 'add', target: 'user', content: '偏好：用短标题。' }, execAt(root));
    assert.equal(existsSync(join(root, '.my-office-memory', 'memory.json')), true);
    assert.equal(existsSync(join(root, '.office', 'memory')), false);
});

// ── mnemon 迁移 ─────────────────────────────────────────────────────────────
//
// 迁移的读取端（runtime 存储 + 长期记忆库）用真实目录形状，长期记忆的读取器
// 由测试注入：一来不必依赖本机有 node:sqlite，二来能把「数据库读坏了」这种
// 情况稳定地造出来。真读数那一节单独用 node:sqlite 造一个真库（见文末）。

/** 造一棵 mnemon 数据树。db 是占位文件：存在性判断用它，内容由注入的读取器给。 */
async function fakeMnemon(root, { workspaceEntries = [], documents = [] } = {}) {
    const ws = join(root, '.mnemon');
    await mkdir(join(ws, 'runtime'), { recursive: true });
    await writeFile(join(ws, 'runtime', 'memories.json'), JSON.stringify({ version: 1, entries: workspaceEntries }), 'utf8');
    await mkdir(join(ws, 'data', 'default'), { recursive: true });
    await writeFile(join(ws, 'data', 'default', 'mnemon.db'), '', 'utf8');
    await writeFile(join(ws, 'data', '.dsh-memory-bodies.json'), JSON.stringify({ version: 1, bodies: [{ id: 'default', name: '测试记忆体' }] }), 'utf8');
    await mkdir(join(ws, 'documents'), { recursive: true });
    await writeFile(join(ws, 'documents', 'index.json'), JSON.stringify({ version: 1, documents }), 'utf8');
    return ws;
}

/** 造一个全局 mnemon 根（runtimeUserScope = "global" 时 USER.md 落在这里）。 */
async function fakeGlobalMnemon(root, entries) {
    const dir = join(root, 'global-mnemon');
    await mkdir(join(dir, 'runtime'), { recursive: true });
    await writeFile(join(dir, 'runtime', 'memories.json'), JSON.stringify({ version: 1, entries }), 'utf8');
    return dir;
}

const runtimeEntry = (content, target = 'memory', importance = 'normal') => ({
    content, target, importance, created_at: '2026-09-01T00:00:00.000Z', updated_at: '2026-09-01T00:00:00.000Z',
});

const insightRow = (id, content, extra = {}) => ({
    id, content, category: 'fact', importance: 4, tags: '["office","测试"]', entities: '["dsh-office-mode"]',
    source: 'agent', stored_at: '2026-09-20T00:00:00.000Z', deleted_at: null, ...extra,
});

/**
 * 默认把全局根指到「不存在的地方」。
 *
 * 不这么做的话，本机真实的 `~/.mnemon`（里面有用户的全局偏好）会漏进每个用例，
 * 断言就会变成「跟这台机器当前的状态比」而不是「跟用例造的数据比」——
 * 这类测试换台机器就红。要测全局来源的用例自己传 globalSource。
 */
const migrate = (options = {}) => migrateMnemon({ globalSource: join(options.root, 'no-global-mnemon'), ...options });

await check('迁移：runtime 热记忆按 target 归位，长期记忆进归档', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, {
        workspaceEntries: [runtimeEntry('项目约定：季度汇报用 business 主题。'), runtimeEntry('用户偏好：书面化表达。', 'user', 'critical')],
    });
    const global = await fakeGlobalMnemon(root, [runtimeEntry('全局用户偏好：习题用十级难度。', 'user')]);
    const receipt = await migrate({
        root,
        globalSource: global,
        readInsights: async () => ({
            insights: [insightRow('i-1', '第一条长期记忆的内容。')],
            edges: [
                { source_id: 'i-1', target_id: 'i-2', edge_type: 'temporal' },
                { source_id: 'i-1', target_id: 'i-9', edge_type: 'entity' },
                { source_id: 'i-8', target_id: 'i-1', edge_type: 'causal' },
            ],
        }),
    });

    assert.equal(receipt.ok, true);
    assert.equal(receipt.hot.project.added, 1);
    assert.equal(receipt.hot.user.added, 2, '工作区 1 条 + 全局 1 条');
    assert.equal(receipt.hot.globalCopied, 1);
    assert.equal(receipt.archive.added, 1);

    const memory = createMemory({ root, memory: {} });
    const snapshot = await memory.snapshot();
    const user = snapshot.entries.filter((entry) => entry.target === 'user').map((entry) => entry.content);
    assert.ok(user.some((text) => text.includes('书面化表达')));
    assert.ok(user.some((text) => text.includes('十级难度')));
    assert.ok(snapshot.entries.some((entry) => entry.target === 'project' && entry.content.includes('business 主题')));
    // 归档条目带上了来源信息与关系边计数（图跳转降级为计数）。
    const archived = (await memory.archiveItems()).find((item) => item.mnemonId === 'i-1');
    assert.ok(archived, '长期记忆应在归档里');
    assert.equal(archived.category, 'fact');
    assert.equal(archived.importance, 4);
    assert.deepEqual(archived.tags, ['office', '测试']);
    assert.deepEqual(archived.entities, ['dsh-office-mode']);
    assert.equal(archived.links.total, 3);
    assert.equal(archived.links.temporal, 1);
    assert.equal(archived.memoryBodyName, '测试记忆体');
    assert.equal(archived.source, 'agent');
});

await check('迁移幂等：同样的数据跑两遍不产生第二份', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, { workspaceEntries: [runtimeEntry('一条项目约定。')] });
    const readInsights = async () => ({ insights: [insightRow('i-1', '一条长期记忆。')], edges: [] });
    const first = await migrate({ root, readInsights });
    assert.equal(first.hot.project.added, 1);
    assert.equal(first.archive.added, 1);

    const second = await migrate({ root, readInsights });
    assert.equal(second.hot.project.added, 0, '第二次不该再加热记忆');
    assert.equal(second.hot.project.skipped, 1);
    assert.equal(second.archive.added, 0, '第二次不该再写归档');
    assert.equal(second.archive.skipped, 1);

    const memory = createMemory({ root, memory: {} });
    assert.equal((await memory.snapshot()).entries.length, 1);
    assert.equal((await memory.archiveItems()).filter((item) => item.mnemonId === 'i-1').length, 1);
});

await check('迁移：已软删的长期记忆不迁，只计数', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, {});
    const receipt = await migrate({
        root,
        readInsights: async () => ({
            insights: [insightRow('i-1', '活的。'), insightRow('i-2', '删掉的。', { deleted_at: '2026-09-21T00:00:00.000Z' })],
            edges: [],
        }),
    });
    assert.equal(receipt.archive.added, 1);
    assert.equal(receipt.archive.deleted, 1);
    const memory = createMemory({ root, memory: {} });
    assert.deepEqual((await memory.archiveItems()).map((item) => item.mnemonId), ['i-1']);
});

await check('迁移：长期记忆读取失败只报警告，runtime 照迁', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, { workspaceEntries: [runtimeEntry('runtime 里的约定。')] });
    const receipt = await migrate({
        root,
        readInsights: async () => { throw new Error('node:sqlite 不存在'); },
    });
    assert.equal(receipt.hot.project.added, 1, 'runtime 部分不受影响');
    assert.equal(receipt.archive.added, 0);
    assert.equal(receipt.archive.failed.length, 1);
    assert.ok(receipt.warnings.some((line) => line.includes('node:sqlite 不存在')));
    assert.ok(renderMigration(receipt).includes('读取失败'), '回执要说清长期记忆没迁成');
});

await check('迁移 dryRun：只统计，不落盘、不建目录', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, { workspaceEntries: [runtimeEntry('约定。')] });
    const receipt = await migrate({
        root,
        dryRun: true,
        readInsights: async () => ({ insights: [insightRow('i-1', '长期记忆。')], edges: [] }),
    });
    assert.equal(receipt.dryRun, true);
    assert.equal(receipt.hot.project.added, 1);
    assert.equal(receipt.archive.added, 1);
    assert.equal(receipt.archive.file, null);
    assert.equal(memoryExists(root), false, '试运行不该建记忆目录');
    assert.ok(renderMigration(receipt).includes('试运行'));
});

await check('迁移：Documents 有内容时明确说「没有对应层」，不静默丢', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, { documents: [{ id: 'd-1', title: '一份项目档案' }] });
    const receipt = await migrate({ root, readInsights: async () => ({ insights: [], edges: [] }) });
    assert.equal(receipt.sources.documents.total, 1);
    assert.ok(receipt.warnings.some((line) => line.includes('没有对应层')));
    assert.ok(renderMigration(receipt).includes('没有对应层'));
});

await check('迁移：源目录不存在时什么都不做，也不报错', async () => {
    const { root } = await freshRoot();
    const receipt = await migrate({ root, readInsights: async () => ({ insights: [], edges: [] }) });
    assert.equal(receipt.sources.workspace.exists, false);
    assert.equal(receipt.hot.project.added, 0);
    assert.equal(receipt.archive.added, 0);
    assert.equal(memoryExists(root), false);
});

await check('office_memory 工具：migrate 动作（含 dryRun）端到端', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, { workspaceEntries: [runtimeEntry('工具侧迁移的约定。')] });
    const tool = buildTools(resolveConfig({ memory: {} })).find((item) => item.name === 'office_memory');
    const exec = execAt(root);

    let out = await tool.execute({ action: 'migrate', dryRun: true }, exec);
    assert.ok(out.ok && out.text.includes('试运行'), out.text);
    assert.equal(memoryExists(root), false, 'dryRun 不落盘');

    out = await tool.execute({ action: 'migrate' }, exec);
    assert.ok(out.text.includes('mnemon → 办公记忆'), out.text);
    assert.equal(memoryExists(root), true);
    out = await tool.execute({ action: 'read', layer: 'hot' }, exec);
    assert.ok(out.text.includes('工具侧迁移的约定'));
});

await check('迁移：用 node:sqlite 读真实长期记忆库（本机没有就跳过）', async () => {
    let sqlite;
    try {
        sqlite = await import('node:sqlite');
    } catch {
        console.log('      （本机 Node 没有 node:sqlite，跳过真库读取这一项）');
        return;
    }
    const { root } = await freshRoot();
    const ws = await fakeMnemon(root, { workspaceEntries: [runtimeEntry('真库旁边的约定。')] });
    const dbPath = join(ws, 'data', 'default', 'mnemon.db');
    const db = new sqlite.DatabaseSync(dbPath);
    db.exec('CREATE TABLE insights (id TEXT PRIMARY KEY, content TEXT, category TEXT, importance INTEGER, tags TEXT, entities TEXT, source TEXT, stored_at TEXT, deleted_at TEXT)');
    db.exec('CREATE TABLE edges (source_id TEXT, target_id TEXT, edge_type TEXT, weight REAL, metadata TEXT, created_at TEXT)');
    const insert = db.prepare('INSERT INTO insights VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)');
    insert.run('real-1', '真库里的长期记忆。', 'insight', 5, '["sqlite"]', '["mnemon.db"]', 'agent', '2026-09-22T00:00:00.000Z', null);
    insert.run('real-2', '这条已软删。', 'fact', 4, '[]', '[]', 'agent', '2026-09-22T00:00:00.000Z', '2026-09-23T00:00:00.000Z');
    db.prepare('INSERT INTO edges VALUES (?, ?, ?, ?, ?, ?)').run('real-1', 'real-2', 'semantic', 0.8, null, '2026-09-22T00:00:00.000Z');
    db.close();

    const receipt = await migrate({ root, readInsights: readMnemonInsights });
    assert.equal(receipt.archive.added, 1, `真库应迁出 1 条，实际 ${JSON.stringify(receipt.archive)}`);
    assert.equal(receipt.archive.deleted, 1);
    const memory = createMemory({ root, memory: {} });
    const archived = (await memory.archiveItems()).find((item) => item.mnemonId === 'real-1');
    assert.equal(archived.category, 'insight');
    assert.equal(archived.importance, 5);
    assert.equal(archived.links.semantic, 1, '关系边要按类型计数');
});

await rm(TMP, { recursive: true, force: true });

await check('插件启动时注册一段静态记忆说明（提示层接住 mnemon 的位置）', async () => {
    const { apply, MEMORY_PROMPT_HINT } = await import('../src/index.js');
    const run = (config) => {
        const sections = [];
        const tools = [];
        const ctx = {
            tools: { register: (definition) => { tools.push(definition.name); return () => {}; } },
            get: (key) => (key === 'systemPrompt'
                ? { section: (value) => { sections.push(value); return () => {}; } }
                : undefined),
            inject: () => {},
            on: () => {},
            logger: {},
        };
        apply(ctx, config);
        return { sections, tools };
    };

    const on = run({ memory: { promptHint: true } });
    assert.ok(on.tools.includes('office_memory'), '记忆工具要注册');
    const hint = on.sections.find((section) => section.name === 'office-mode-memory');
    assert.ok(hint, '应注册记忆说明段');
    assert.equal(hint.text, MEMORY_PROMPT_HINT);
    assert.ok(hint.text.includes('office_memory'), '说明里要给出工具名');
    assert.ok(hint.text.includes('.office/memory'), '说明里要给出记忆目录');

    const off = run({ memory: { promptHint: false } });
    assert.ok(!off.sections.some((section) => section.name === 'office-mode-memory'), '关掉后不注册');

    const disabled = run({ memory: { enabled: false, promptHint: true } });
    assert.ok(!disabled.sections.some((section) => section.name === 'office-mode-memory'), '记忆整层关掉后也不注册');
});

// ── 第二轮：把 mnemon 并进来的能力 ──────────────────────────────────────────
//
// 覆盖：层拓扑开关、跨项目层（全局根）、图关系（含下沉后关系不断）、
// 召回质量分档、每回合配额、备份与迁移 Pack、实体视图、浏览用原始读取。

/** 带隔离全局根的用例目录：绝不碰本机真实的 $DSH_HOME/.office/memory。 */
async function freshScoped(memoryConfig = {}) {
    const { root, memory } = await freshRoot();
    const globalDir = join(TMP, `global-${caseId}`);
    await rm(globalDir, { recursive: true, force: true });
    const scoped = { ...memoryConfig, globalDir };
    return { root, globalDir, memory: createMemory({ root, memory: scoped }), memoryConfig: scoped };
}

/** 造一个带回合日志的假 exec：配额要能读出 turn 号才生效。 */
function execWithTurns(root, turns = [1]) {
    const events = turns.map((turn) => ({ type: 'turn/start', data: { turn } }));
    return {
        agent: {
            session: {
                id: `s-${caseId}`,
                header: { cwd: root },
                ownEvents: () => events,
            },
        },
    };
}

await check('层拓扑：关掉热记忆后读不到也写不进，重新打开内容还在', async () => {
    const { root, memoryConfig } = await freshScoped();
    await createMemory({ root, memory: memoryConfig }).mutate({ action: 'add', target: 'project', content: '关层前写下的约定。' });

    const off = createMemory({ root, memory: { ...memoryConfig, layers: { hot: false, ledger: true, archive: true } } });
    const value = await off.read({ layer: 'hot' });
    assert.equal(value.hot, undefined, '热记忆层关掉后不该返回 hot');
    await assert.rejects(() => off.mutate({ action: 'add', target: 'project', content: '不该写进去。' }), /热记忆层已在设置里关掉/);

    const on = createMemory({ root, memory: memoryConfig });
    const back = await on.read({ layer: 'hot' });
    assert.equal(back.hot.items.length, 1, '关层不删数据：重新打开还在');
    assert.equal(back.hot.items[0].text, '关层前写下的约定。');
});

await check('层拓扑：关掉台账层后不再登记，但已有台账仍可读', async () => {
    const { root, memoryConfig } = await freshScoped();
    await createMemory({ root, memory: memoryConfig }).log([{ path: 'a.pptx', format: 'ppt' }]);
    const off = createMemory({ root, memory: { ...memoryConfig, layers: { hot: true, ledger: false, archive: true } } });
    const logged = await off.log([{ path: 'b.pptx', format: 'ppt' }]);
    assert.equal(logged.disabled, true, '台账层关掉要明确回执，而不是静默成功');
    const value = await off.read({ layer: 'ledger' });
    assert.equal(value.ledger, undefined, '台账层关掉后 read 不该返回 ledger');
});

await check('跨根改删：命中在另一个根里的条目时，改的必须真的是那一条', async () => {
    // 这一条是第二十七轮代码审查挖出来的**既有缺陷**（不是本轮引入的，但本轮把
    // 「按 id 跨根改删」写成了可用能力，所以必须钉住）：跨根命中时若回头再 loadStore
    // 一次，matched 与写回的数组不是同一批对象 —— remove 的 indexOf 得 -1（splice(-1,1)
    // 会删掉最后一条），replace 只改到临时对象、saveStore 写回时内容没变（静默假回执）。
    const { root, memoryConfig } = await freshScoped({ scope: 'both' });
    const globalOnly = createMemory({ root, memory: { ...memoryConfig, scope: 'global' } });
    const first = await globalOnly.mutate({ action: 'add', target: 'project', content: '全局根里的第一条（A）。' });
    const second = await globalOnly.mutate({ action: 'add', target: 'project', content: '全局根里的第二条（B）。' });

    const both = createMemory({ root, memory: memoryConfig });
    // 「当前根里没有」这件事要用只读工作区的那一份看：scope:'both' 的 read 会把两个根合并。
    const workspaceOnly = createMemory({ root, memory: { ...memoryConfig, scope: 'workspace' } });
    assert.equal((await workspaceOnly.read({ layer: 'hot' })).hot.items.length, 0, '工作区根里本来是空的');

    // 按 oldText 跨根改：内容必须真的落盘
    const replaced = await both.mutate({ action: 'replace', oldText: '全局根里的第一条（A）', content: '全局根里的第一条（A，已改）。' });
    assert.equal(replaced.entry.id, first.entry.id);
    const afterReplace = (await globalOnly.read({ layer: 'hot' })).hot.items;
    assert.ok(afterReplace.some((item) => item.id === first.entry.id && item.text.includes('已改')), '跨根 replace 必须真的写回，不能只给回执');

    // 按 id 跨根删：删的必须是那一条，不能变成「最后一条」
    const removed = await both.mutate({ action: 'remove', id: second.entry.id });
    assert.equal(removed.previous.id, second.entry.id);
    const left = (await globalOnly.read({ layer: 'hot' })).hot.items.map((item) => item.id).sort();
    assert.deepEqual(left, [first.entry.id], `只该删掉 B，实际剩下 ${JSON.stringify(left)}`);
});

await check('跨项目层：userScope=global 的用户偏好落全局根，换工作目录仍读得到', async () => {
    const globalDir = join(TMP, `global-shared-${caseId + 1}`);
    await rm(globalDir, { recursive: true, force: true });
    const config = { scope: 'both', userScope: 'global', globalDir };

    const a = join(TMP, `case-${caseId + 1}-a`);
    const b = join(TMP, `case-${caseId + 1}-b`);
    await rm(a, { recursive: true, force: true });
    await rm(b, { recursive: true, force: true });
    await mkdir(a, { recursive: true });
    await mkdir(b, { recursive: true });

    const written = await createMemory({ root: a, memory: config }).mutate({ action: 'add', target: 'user', content: '用户要求：所有汇报一律 16:9。' });
    assert.equal(written.origin, 'global', 'userScope=global 时用户偏好应写进全局根');
    assert.ok(existsSync(join(globalDir, 'memory.json')), '全局根里应有 memory.json');
    assert.ok(!existsSync(join(a, '.office', 'memory', 'memory.json')), '这条不该同时写进工作区根');

    const other = await createMemory({ root: b, memory: config }).read({ layer: 'hot' });
    const hit = other.hot.items.find((item) => item.text.includes('16:9'));
    assert.ok(hit, '换一个工作目录也该读到这条全局偏好');
    assert.equal(hit.origin, 'global');
});

await check('跨项目层：scope=workspace 时全局根不参与，行为与单根一致', async () => {
    const { root, globalDir, memoryConfig } = await freshScoped();
    await createMemory({ root, memory: memoryConfig }).mutate({ action: 'add', target: 'project', content: '只属于这个项目的约定。' });
    const value = await createMemory({ root, memory: memoryConfig }).read({ layer: 'hot' });
    assert.deepEqual(value.origins, ['workspace'], '默认只开工作区根');
    assert.equal(value.hot.items[0].origin, 'workspace');
    assert.ok(!existsSync(globalDir), '默认范围不该去建全局目录');
});

await check('图关系：建关系后从两端都能走到对面，重复建不产生第二条', async () => {
    const { root, memory } = await freshRoot();
    const one = await memory.mutate({ action: 'add', target: 'project', content: '决策：用 business 主题。' });
    const two = await memory.mutate({ action: 'add', target: 'project', content: '理由：客户品牌色是蓝色。' });

    const made = await memory.link({ sourceId: one.entry.id, targetId: two.entry.id, kind: 'supports', note: '配色依据' });
    assert.equal(made.added, true);
    assert.equal(made.link.kind, 'supports');

    const again = await memory.link({ sourceId: one.entry.id, targetId: two.entry.id, kind: 'supports' });
    assert.equal(again.added, false, '同一条关系重复建应幂等');

    const fromOne = await memory.related({ id: one.entry.id });
    assert.equal(fromOne.nodes.length, 1);
    assert.equal(fromOne.nodes[0].id, two.entry.id, '从起点应走到终点');
    assert.equal(fromOne.nodes[0].layer, 'hot');

    const fromTwo = await memory.related({ id: two.entry.id });
    assert.equal(fromTwo.nodes[0].id, one.entry.id, '关系是双向的：从终点也应走回起点');
});

await check('图关系：指向不存在的 id 报错，不建悬空边', async () => {
    const { root, memory } = await freshRoot();
    const one = await memory.mutate({ action: 'add', target: 'project', content: '真实存在的一条。' });
    await assert.rejects(() => memory.link({ sourceId: one.entry.id, targetId: 'm-不存在' }), /找不到 id 为/);
    await assert.rejects(() => memory.link({ sourceId: one.entry.id, targetId: one.entry.id }), /不能把一条记忆连到它自己/);
    assert.equal((await memory.related({ id: one.entry.id })).edges.length, 0);
});

await check('图关系：热记忆下沉进归档后关系仍走得通（id 必须带过去）', async () => {
    const { root, memory } = await freshRoot({ userLimitBytes: 512, projectLimitBytes: 1024 });
    const anchor = await memory.mutate({ action: 'add', target: 'project', content: '锚点：这份稿子的主题约定。', importance: 'critical' });
    const doomed = await memory.mutate({ action: 'add', target: 'project', content: '要被挤下去的一条次要记录。', importance: 'low' });
    await memory.link({ sourceId: anchor.entry.id, targetId: doomed.entry.id, kind: 'refines' });

    // 一直塞到容量维持真的把 low 那条顶下去为止：断言不依赖对字节数的精确估算。
    let archived = [];
    for (let i = 0; i < 25; i += 1) {
        await memory.mutate({ action: 'add', target: 'project', content: `填充条目 ${i}：占位内容，用来把项目侧容量顶到上限以上，让最次要的那条下沉。`, importance: 'normal' });
        archived = (await memory.archiveItems()).map((item) => item.id);
        if (archived.includes(doomed.entry.id)) break;
    }
    assert.ok(archived.includes(doomed.entry.id), '次要条目应已下沉到归档，且归档里保留原 id');

    const walked = await memory.related({ id: anchor.entry.id });
    assert.ok(walked.nodes.some((node) => node.id === doomed.entry.id), '下沉之后关系不该断');
    assert.equal(walked.nodes.find((node) => node.id === doomed.entry.id).layer, 'archive');
});

await check('图关系：关掉 links 后 link / related 都明确报错', async () => {
    const { root, memoryConfig } = await freshScoped({ links: false });
    const memory = createMemory({ root, memory: memoryConfig });
    await assert.rejects(() => memory.link({ sourceId: 'a', targetId: 'b' }), /图关系已在设置里关掉/);
    await assert.rejects(() => memory.related({ id: 'a' }), /图关系已在设置里关掉/);
});

await check('实体视图：只认条目显式声明的 entities，不做抽取', async () => {
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'project', content: '系统 A 的接口约定。', entities: ['系统 A', '接口组'] });
    await memory.mutate({ action: 'add', target: 'project', content: '系统 A 的部署方式。', entities: ['系统 A'] });
    await memory.mutate({ action: 'add', target: 'project', content: '这条没有声明任何实体。' });

    const all = await memory.entities({});
    const names = all.items.map((item) => item.name);
    assert.deepEqual(names.sort(), ['接口组', '系统 A'], '只有声明过的名字才出现');
    assert.equal(all.items.find((item) => item.name === '系统 A').count, 2);
    assert.equal(all.items.find((item) => item.name === '系统 A').refs[0].layer, 'hot');

    const filtered = await memory.entities({ query: '接口' });
    assert.deepEqual(filtered.items.map((item) => item.name), ['接口组']);
});

await check('召回质量 strict-v1：低相关结果被丢掉并报出条数，policy=off 时不丢', async () => {
    const { root, memoryConfig } = await freshScoped();
    const memory = createMemory({ root, memory: memoryConfig });
    await memory.log([
        { path: '季度汇报.pptx', format: 'ppt', purpose: '季度汇报主线' },
        { path: '季度数据.xlsx', format: 'excel', purpose: '季度汇报数据底稿' },
        { path: '随手记.md', format: 'md', purpose: '季度里的一条随手记录' },
        { path: '年会策划.pptx', format: 'ppt', purpose: '年会策划' },
    ]);
    // 五个词只命中一个 → 相关度 0.2，低于 0.25 的下限，落进「未知档」。
    // 名额**显式配成 2**：这条用例测的是分档机械本身，不该因为默认值调整而变红
    //（默认名额在第二十四轮从 2 放宽到 4，另有专门的默认值用例守着）。
    const capped = createMemory({ root, memory: { ...memoryConfig, recallQuality: { ...(memoryConfig.recallQuality ?? {}), maxUnknownResults: 2 } } });
    const strict = await capped.read({ layer: 'ledger', query: '季度 甲词 乙词 丙词 丁词' });
    assert.equal(strict.quality, 'strict-v1');
    assert.ok(strict.ledger.matched >= 3, '命中 3 条候选');
    assert.equal(strict.ledger.items.length, 2, 'strict-v1 只收「未知档」的名额（配成 2 条）');
    assert.ok(strict.ledger.qualityDropped > 0, '被丢掉多少条要报出来');

    // 名额 0 = 一条都不要（mnemon 的原义）。
    const none = createMemory({ root, memory: { ...memoryConfig, recallQuality: { ...(memoryConfig.recallQuality ?? {}), maxUnknownResults: 0 } } });
    const zero = await none.read({ layer: 'ledger', query: '季度 甲词 乙词 丙词 丁词' });
    assert.equal(zero.ledger.items.length, 0, '未知档名额 0 = 一条都不要');

    const off = createMemory({ root, memory: { ...memoryConfig, recallQuality: { ...memoryConfig.recallQuality, policy: 'off' } } });
    const loose = await off.read({ layer: 'ledger', query: '季度 甲词 乙词 丙词 丁词' });
    assert.equal(loose.quality, 'off');
    assert.equal(loose.ledger.qualityDropped, 0, 'off 时不丢结果');
    assert.ok(loose.ledger.items.length > strict.ledger.items.length, 'off 应给出比 strict-v1 更多的结果');
});

await check('召回质量：阈值写反时整组退回默认（不留下自相矛盾的区间）', async () => {
    const { root } = await freshScoped();
    const broken = resolveConfig({ memory: { recallQuality: { lowScoreThreshold: 0.9, highScoreThreshold: 0.1 } } }).memory.recallQuality;
    assert.equal(broken.lowScoreThreshold, 0.25);
    assert.equal(broken.highScoreThreshold, 0.6);
    void root;
});

// 第二十四轮放宽的三个名额与候选倍数。钉在这里是为了让「默认值被改动」变成一件显式的事：
// 依据是实证的（RMM：Oracle 上限 90.2 而实际 70.4，Top-M 5→10 值 +4.6 Recall@5，reranker
// 只值 +1.4；RAGChecker：换检索器值 +2.4 F1，名额 k 5→20 值 +1.7）—— 上限由召回决定。
await check('召回质量：第二十四轮放宽的默认名额与候选倍数（阈值不动）', async () => {
    const q = resolveConfig({ memory: {} }).memory.recallQuality;
    assert.equal(q.candidateMultiplier, 5, '候选池倍数 3 → 5');
    assert.equal(q.maxMediumResults, 6, '中档名额 4 → 6');
    assert.equal(q.maxUnknownResults, 4, '未知档名额 2 → 4');
    assert.equal(q.lowScoreThreshold, 0.25, '阈值不动：放宽的是名额，不是判定');
    assert.equal(q.highScoreThreshold, 0.6, '阈值不动');
    assert.equal(q.policy, 'strict-v1', '策略不动');
});

// 第二十四轮 P0：中文查询必须切词。改前（整串当一个 term）16 条中文查询里 8 条零结果；
// 这里钉住「词序颠倒」与「合称 vs 全称」这两类最典型的失败，它们是最容易回归的地方。
await check('中文切词：词序颠倒与合称查询都要能命中（bigram）', async () => {
    const { root, memory } = await freshRoot();
    await memory.log([
        { path: '面板重做.md', format: 'md', purpose: '记忆面板视觉重做' },
        { path: '音视频.md', format: 'md', purpose: '音频与视频的内容提取' },
        { path: '知识库.md', format: 'md', purpose: '记忆系统结构优化与知识库结合' },
    ]);

    // 整串命中：改造前也能过，留作基线。
    const exact = await memory.read({ layer: 'ledger', query: '记忆面板' });
    assert.ok(exact.ledger.items.some((item) => item.path === '面板重做.md'), '整串命中要过');

    // 词序颠倒：改造前 0 条（`面板记忆` 不是任何条目的子串）。
    const swapped = await memory.read({ layer: 'ledger', query: '面板记忆' });
    assert.ok(swapped.ledger.matched > 0, '词序颠倒不该零结果');
    assert.ok(swapped.ledger.items.some((item) => item.path === '面板重做.md'), '词序颠倒要命中');

    // 合称 vs 全称：正文写「音频与视频」，查询写「音视频」。
    const abbr = await memory.read({ layer: 'ledger', query: '音视频' });
    assert.ok(abbr.ledger.items.some((item) => item.path === '音视频.md'), '合称要命中全称');

    // 换词序的第二个例子。
    const swapped2 = await memory.read({ layer: 'ledger', query: '结合知识库' });
    assert.ok(swapped2.ledger.items.some((item) => item.path === '知识库.md'), '「结合知识库」要命中「知识库结合」');

    // 中英混排：CJK 段切 bigram，拉丁段整词。
    await memory.log([{ path: '绘图.py', format: 'md', purpose: 'office.python 计算与绘图' }]);
    const mixed = await memory.read({ layer: 'ledger', query: 'python 绘图' });
    assert.ok(mixed.ledger.items.some((item) => item.path === '绘图.py'), '中英混排要命中');
    const glued = await memory.read({ layer: 'ledger', query: 'python绘图' });
    assert.ok(glued.ledger.items.some((item) => item.path === '绘图.py'), '中英黏在一起也要命中（分段处理）');

    // 单字 CJK 仍按整串（没有二元组可切，回落）。
    const single = await memory.read({ layer: 'ledger', query: '绘' });
    assert.ok(single.ledger.items.some((item) => item.path === '绘图.py'), '单字回落成整串匹配');

    // 完全不相关的查询仍要零结果，且要说清是「没命中」而不是「空」。
    // 注意选的词：`量子计算` 会通过二元组 `计算` 命中上面那条「计算与绘图」—— 那是
    // bigram 的**正常部分重合**，不是误命中。要用一个连二元组都不重合的查询来测零结果。
    const miss = await memory.read({ layer: 'ledger', query: '量子纠缠' });
    assert.equal(miss.ledger.matched, 0, '不相关查询不该凭空命中');
    const { renderRead } = await import('../src/memory.js');
    assert.ok(renderRead(miss).includes('没有命中的条目'), '零结果要说清是没命中');
});

await check('每回合配额：第一次初查、第二次细化、第三次被拒；换回合后重新计数', async () => {
    const { root } = await freshScoped();
    const { createTurnQuota } = await import('../src/quota.js');
    // 显式给上限：默认值来自 config.js，这里要测的是计数逻辑本身。
    const quota = createTurnQuota({ limits: { recallPerTurn: 1, recallRefinePerTurn: 1, relatedPerTurn: 1 } });
    const exec = execWithTurns(root, [1]);

    const first = quota.take('recallPerTurn', exec);
    assert.equal(first.allowed, true);
    assert.equal(first.enforced, true);
    assert.equal(first.turn, 1);

    const second = quota.take('recallPerTurn', exec);
    assert.equal(second.allowed, false, '初查名额用完');
    const refine = quota.take('recallRefinePerTurn', exec);
    assert.equal(refine.allowed, true, '第二次走细化名额');

    const third = quota.take('recallPerTurn', exec);
    const thirdRefine = quota.take('recallRefinePerTurn', exec);
    assert.equal(third.allowed || thirdRefine.allowed, false, '两个名额都用完就该拒绝');

    // 同一个会话进入下一回合：事件日志里出现新的 turn/start。
    const nextTurn = execWithTurns(root, [1, 2]);
    const fresh = quota.take('recallPerTurn', nextTurn);
    assert.equal(fresh.allowed, true, '新回合应重新计数');
    assert.equal(fresh.turn, 2);
});

await check('每回合配额：拿不到回合身份时放行并如实说没在管', async () => {
    const { createTurnQuota } = await import('../src/quota.js');
    const quota = createTurnQuota({ limits: { recallPerTurn: 1 } });
    const decision = quota.take('recallPerTurn', {});
    assert.equal(decision.allowed, true);
    assert.equal(decision.enforced, false, '没有会话对象时不能假装在管配额');
    assert.equal(decision.turn, null);
    assert.match(decision.reason, /拿不到回合身份/);
});

await check('每回合配额：0 表示不限制', async () => {
    const { createTurnQuota } = await import('../src/quota.js');
    const quota = createTurnQuota({ limits: { recallPerTurn: 0 } });
    const exec = { agent: { session: { id: 's-zero', ownEvents: () => [{ type: 'turn/start', data: { turn: 1 } }] } } };
    for (let i = 0; i < 5; i += 1) {
        assert.equal(quota.take('recallPerTurn', exec).allowed, true, '0 = 不限制');
    }
});

await check('Pack 导出与导入：整包搬走并幂等合并，重复导入不产生第二份', async () => {
    const { root, memory } = await freshRoot();
    const one = await memory.mutate({ action: 'add', target: 'project', content: '要搬走的项目约定。' });
    const two = await memory.mutate({ action: 'add', target: 'project', content: '第二条，用来建关系。' });
    await memory.link({ sourceId: one.entry.id, targetId: two.entry.id, kind: 'related' });
    await memory.log([{ path: '产物.pptx', format: 'ppt', purpose: '导出用例' }]);
    await memory.appendArchiveBlock({ reason: 'mnemon-migration', items: [{ id: 'i-9', kind: 'mnemon', content: '迁来的一条长期记忆', entities: ['实体甲'] }] });

    const pack = await memory.exportPack();
    assert.equal(pack.format, 'dsh-office-memory-pack');
    assert.equal(pack.stores.length, 1);
    assert.equal(pack.stores[0].entries.length, 2);
    assert.equal(pack.stores[0].links.length, 1);
    assert.ok(pack.stores[0].archiveFiles['2026-09'] !== undefined || Object.keys(pack.stores[0].archiveFiles).length === 1, '归档 Markdown 全文应随包带走');

    const { root: other, memory: fresh } = await freshRoot();
    const report = await fresh.importPack(pack);
    assert.equal(report.entries, 2);
    assert.equal(report.links, 1);
    assert.equal(report.ledger, 1);
    assert.ok(report.archiveBlocks >= 1);

    const read = await fresh.read({ layer: 'all' });
    assert.equal(read.hot.items.length, 2);
    assert.equal(read.ledger.total, 1);
    assert.ok(read.archive.items_total >= 1);

    const second = await fresh.importPack(pack);
    assert.equal(second.entries, 0, '重复导入不应新增热记忆');
    assert.equal(second.skippedEntries, 2);
    assert.equal(second.links, 0, '重复导入不应新增关系');
    assert.equal(second.ledger, 0);
    // 归档块也要幂等：老库（mnemon 迁移那批）的块没有 id，只有「时刻+原因+条数」这个
    // 兜底键 —— 写入时若给它发随机 id，第二次导入就对不上、归档会翻倍
    // （第四十轮现场探针实测：重复导入多出 18 条）。
    assert.equal(second.archiveBlocks, 0, '重复导入不应新增归档块');
    assert.equal((await fresh.read({ layer: 'archive' })).archive.items_total, read.archive.items_total, '重复导入不应让归档条目变多');
    assert.equal((await fresh.read({ layer: 'hot' })).hot.items.length, 2, '导入是只增不改，不覆盖现内容');
    void other;
});

await check('Pack 导入：不是本插件的包要明确拒绝', async () => {
    const { memory } = await freshRoot();
    await assert.rejects(() => memory.importPack({ format: 'something-else' }), /这不是办公记忆 Pack/);
    await assert.rejects(() => memory.importPack(null), /需要一份 Pack 对象/);
});

await check('browse：给面板的是原始记录（带 origin），不过召回质量', async () => {
    const { root, memory } = await freshRoot();
    await memory.mutate({ action: 'add', target: 'user', content: '一条用户偏好。' });
    await memory.log([{ path: 'a.pptx', format: 'ppt' }, { path: 'b.xlsx', format: 'excel' }]);
    const value = await memory.browse({});
    assert.equal(value.hot.length, 1);
    assert.equal(value.ledger.length, 2);
    assert.equal(value.hot[0].origin, 'workspace');
    assert.equal(value.ledger[0].origin, 'workspace');
    assert.deepEqual(value.links, []);
});

await check('status：给出各根的计数与层开关', async () => {
    const { root, memoryConfig } = await freshScoped({ scope: 'both', userScope: 'global' });
    const memory = createMemory({ root, memory: memoryConfig });
    await memory.mutate({ action: 'add', target: 'user', content: '全局的用户偏好。' });
    await memory.mutate({ action: 'add', target: 'project', content: '工作区的项目约定。' });
    const value = await memory.status();
    assert.equal(value.scope, 'both');
    assert.equal(value.userScope, 'global');
    assert.deepEqual(value.stores.map((store) => store.id), ['global', 'workspace']);
    assert.equal(value.stores.find((store) => store.id === 'global').hot, 1);
    assert.equal(value.stores.find((store) => store.id === 'workspace').hot, 1);
});

await check('office_memory 工具：link / related / entities / status / export / import 端到端', async () => {
    const { root } = await freshRoot();
    const config = resolveConfig({ memory: { globalDir: join(TMP, 'global-tool') } });
    const tools = buildTools(config, {});
    const find = (name) => tools.find((tool) => tool.name === name);
    const exec = execAt(root);

    const a = await find('office_memory').execute({ action: 'add', target: 'project', content: '甲：模板用 A。', entities: ['模板 A'] }, exec);
    const b = await find('office_memory').execute({ action: 'add', target: 'project', content: '乙：模板用 A 的原因。' }, exec);
    assert.match(a.text, /记忆已更新/);

    const ids = (await createMemory({ root, memory: config.memory }).read({ layer: 'hot' })).hot.items.map((item) => item.id);
    assert.equal(ids.length, 2);

    const linked = await find('office_memory').execute({ action: 'link', sourceId: ids[0], targetId: ids[1], kind: 'supports' }, exec);
    assert.match(linked.text, /已建立关系/);

    const walked = await find('office_memory').execute({ action: 'related', id: ids[0] }, exec);
    assert.match(walked.text, /关系遍历/);
    assert.ok(walked.text.includes(ids[1]), '遍历结果里应出现对面那条');

    const entities = await find('office_memory').execute({ action: 'entities' }, exec);
    assert.match(entities.text, /实体视图/);
    assert.ok(entities.text.includes('模板 A'));

    const status = await find('office_memory').execute({ action: 'status' }, exec);
    assert.match(status.text, /记忆状态/);

    const exported = await find('office_memory').execute({ action: 'export', packPath: 'pack.json' }, exec);
    assert.match(exported.text, /记忆已导出/);
    assert.ok(existsSync(join(root, 'pack.json')));

    const imported = await find('office_memory').execute({ action: 'import', packPath: 'pack.json' }, exec);
    assert.match(imported.text, /记忆导入完成/);
    assert.match(imported.text, /热记忆 新增 0 条/);

    // 配额：一个回合里第一次带 query 的检索走初查名额，第二次走细化名额，第三次被挡。
    const quotaConfig = resolveConfig({ memory: { quota: { recallPerTurn: 1, recallRefinePerTurn: 1, relatedPerTurn: 1 } } });
    const quotaTools = buildTools(quotaConfig, {});
    const quotaExec = execWithTurns(root, [7]);
    const readAction = (args) => quotaTools.find((tool) => tool.name === 'office_memory').execute(args, quotaExec);
    assert.equal((await readAction({ action: 'read', query: '甲' })).ok, true, '第一次检索应放行');
    assert.equal((await readAction({ action: 'read', query: '乙' })).ok, true, '第二次走细化名额');
    const third = await readAction({ action: 'read', query: '丙' });
    assert.equal(third.ok, false, '第三个检索应被拒绝');
    assert.match(third.text, /配额已用完/);
    const browse = await readAction({ action: 'read' });
    assert.equal(browse.ok, true, '不带 query 的浏览不占配额');

    // 抛错的检索不吃名额（第四十五轮复核 P3 的同类）：时间词解析不了会当场抛错，那次没有
    // 产生任何结果，所以「初查 + 细化」两次机会应当都还在（不退回的话第 2 次就没了）。
    //
    // 事件数组**要比上面那次长**：`turnKeyOf` 按会话缓存「事件数 → 当时的 turn 号」，
    // 同样长度的新数组会让它复用上一个回合的缓存（那就变成「本回合配额已用完」，而不是抛错）。
    const refundExec = execWithTurns(root, [7, 8]);
    const refundRead = (args) => quotaTools.find((tool) => tool.name === 'office_memory').execute(args, refundExec);
    await assert.rejects(() => refundRead({ action: 'read', query: '甲', since: '不是时间' }), /看不懂的时间/);
    assert.equal((await refundRead({ action: 'read', query: '甲' })).ok, true, '抛错那次把名额吃掉了（初查）');
    assert.equal((await refundRead({ action: 'read', query: '乙' })).ok, true, '抛错那次把名额吃掉了（细化）');
    assert.equal((await refundRead({ action: 'read', query: '丙' })).ok, false, '退回不等于不限制');
});

await check('主动记录：digest 里带上「当场记一条」的指引，关掉后不再出现', async () => {
    const { root, memoryConfig } = await freshScoped({ autoCapture: true });
    const on = await createMemory({ root, memory: memoryConfig }).digest();
    const text = renderDigest(on, { context: 'help' });
    assert.match(text, /主动记录/);

    const { root: root2, memoryConfig: offConfig } = await freshScoped({ autoCapture: false });
    const off = await createMemory({ root: root2, memory: offConfig }).digest();
    assert.doesNotMatch(renderDigest(off, { context: 'help' }), /主动记录/);
});

// ── 按 id 寻址（第十九轮 P0-1 / P1-3，第二十四轮 P1-3 / P1-4，第二十七轮落地） ──
//
// 两条主责链的硬前置：
//   · 归档条目的唯一入口 `entryIndex` 只认 `id` —— 迁移条目当初只写了 `mnemonId`，
//     18 条历史长期记忆在 link / related 里等于不存在；
//   · `link` / `related` / `unlink` 全都要 id，而 read 当初只在归档行印 id，
//     热记忆与台账行不印 —— 第八轮交付的图关系有一半入口是暗的。

await check('按 id 寻址：read 三层都印 id，link / related / replace / remove 都能直接用它', async () => {
    const { root } = await freshRoot();
    const memory = createMemory({ root, memory: {} });
    const project = await memory.mutate({ action: 'add', target: 'project', content: '约定：季度汇报统一用 business 主题。' });
    const user = await memory.mutate({ action: 'add', target: 'user', content: '偏好：表达书面化。', importance: 'critical' });
    await memory.log([{ path: '季度回顾.docx', format: 'word', purpose: '给管理层看' }]);
    const ledgerId = JSON.parse((await readFile(memoryPaths(root, {}).ledger, 'utf8')).trim().split('\n')[0]).id;
    await memory.appendArchiveBlock({
        reason: 'test',
        items: [{ id: 'a-probe', kind: 'hot', origin: 'project', importance: 'normal', content: '一条归档条目。', at: '2026-09-01T00:00:00.000Z' }],
    });

    const value = await memory.read({ layer: 'all' });
    const text = renderRead(value);
    for (const id of [project.entry.id, user.entry.id, ledgerId, 'a-probe']) {
        assert.ok(text.includes(id), `read 的输出里没有 id ${id}\n${text}`);
    }
    // 形状断言必须先确认「这一层真的有条目」：层被预算饿死时列表是空的，
    // for-of 会静默全过（第二十六轮实测过「台账恒 0 条」，这种回归不能靠空循环兜住）。
    assert.ok(value.hot.items.length >= 2, `热记忆应至少 2 条，实际 ${value.hot.items.length}`);
    assert.ok(value.ledger.items.length >= 1, `台账应至少 1 条，实际 ${value.ledger.items.length}`);
    // 「行首 `- ` 之后第二段就是 id」这条形状要稳住：模型是按位置抄的，形状变了它就抄错。
    for (const item of value.hot.items) assert.ok(text.includes(`- [${item.importance}] ${item.id} `), `热记忆行缺 id：${item.id}`);
    for (const item of value.ledger.items) assert.ok(text.includes(`- ${item.at.slice(0, 10)} ${item.id} `), `台账行缺 id：${item.id}`);

    // 拿 read 出来的两个 id 直接建关系，再走一跳 —— 不必再猜原文或翻归档。
    const linked = await memory.link({ sourceId: project.entry.id, targetId: ledgerId, kind: 'derives' });
    assert.equal(linked.added, true);
    const related = await memory.related({ id: project.entry.id });
    assert.deepEqual(related.nodes.map((node) => node.id), [ledgerId]);

    // 按 id 改 / 删：oldText 那条老路照旧，不必给 id。
    const replaced = await memory.mutate({ action: 'replace', id: project.entry.id, content: '约定：季度汇报统一用 business 主题（双周更新）。' });
    assert.equal(replaced.entry.id, project.entry.id);
    assert.ok(replaced.entry.content.includes('双周更新'));
    assert.ok(renderMutation(replaced).includes(project.entry.id), '写回执也要印 id，刚记的这条才能马上拿去 link');
    const removed = await memory.mutate({ action: 'remove', id: user.entry.id });
    assert.equal(removed.previous.id, user.entry.id);
    assert.equal((await memory.snapshot()).entries.some((entry) => entry.id === user.entry.id), false, '按 id 删要真的删掉');

    // 两个地址都不给、或 id 不存在时，给的是可读错误，不是静默新建 / 改错条。
    await assert.rejects(() => memory.mutate({ action: 'replace', content: 'x' }), /需要 oldText 或 id/);
    await assert.rejects(() => memory.mutate({ action: 'remove', id: 'm-不存在' }), /没有 id 为/);
    await assert.rejects(() => memory.mutate({ action: 'remove', id: ledgerId }), /没有 id 为/, '台账条目不在热记忆里，按 id 也改不动');
});

await check('归档补 id：迁移条目可 link / related，人可读的摘要里也有边数', async () => {
    const { root } = await freshRoot();
    await fakeMnemon(root, { workspaceEntries: [runtimeEntry('项目约定：季度汇报用 business 主题。')] });
    const receipt = await migrate({
        root,
        readInsights: async () => ({
            insights: [insightRow('i-1', '第一条长期记忆的内容。')],
            edges: [
                { source_id: 'i-1', target_id: 'i-2', edge_type: 'temporal' },
                { source_id: 'i-1', target_id: 'i-9', edge_type: 'entity' },
                { source_id: 'i-8', target_id: 'i-1', edge_type: 'causal' },
            ],
        }),
    });
    assert.equal(receipt.archive.added, 1);

    const memory = createMemory({ root, memory: {} });
    const paths = memoryPaths(root, {});
    const archived = (await memory.archiveItems()).find((item) => item.mnemonId === 'i-1');
    assert.equal(archived.id, 'mnemon:i-1', '归档条目的稳定 id 就是 mnemon:<原始 id>');
    assert.equal(archived.linkedCount, 3, '边数要有渲染用的那个字段名');

    // 摘要（人能打开的那份 Markdown）里关系边数不能再是空的。
    const index = JSON.parse(await readFile(paths.index, 'utf8'));
    const md = await readFile(join(paths.archive, `${index.digests[0].month}.md`), 'utf8');
    assert.ok(md.includes('关系边 3 条'), `摘要要写出边数\n${md}`);

    // 迁移条目现在能当关系的两端：这正是 18 条历史记忆当初做不到的事。
    const hot = await memory.mutate({ action: 'add', target: 'project', content: '再记一条项目约定。' });
    assert.equal((await memory.link({ sourceId: 'mnemon:i-1', targetId: hot.entry.id })).added, true);
    const related = await memory.related({ id: 'mnemon:i-1' });
    assert.ok(related.nodes.some((node) => node.id === hot.entry.id), '从迁移条目出发要能走一跳');
});

await check('归档补 id 的历史兜底：老索引只写 mnemonId 时，读入即补，不必重迁', async () => {
    // 真实库就是这样：58 条归档里 18 条只有 mnemonId（第二十六轮实测）。
    const { root } = await freshRoot();
    const paths = memoryPaths(root, {});
    await mkdir(paths.archive, { recursive: true });
    await writeFile(paths.index, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-01T00:00:00.000Z',
        digests: [{
            id: 'a-2026-09',
            month: '2026-09',
            at: '2026-09-01T00:00:00.000Z',
            file: 'archive/2026-09.md',
            blocks: [{
                id: 'b-legacy',
                at: '2026-09-01T00:00:00.000Z',
                reason: 'mnemon-migration',
                items: [{
                    kind: 'mnemon',
                    mnemonId: 'i-9',
                    content: '老库里的一条长期记忆。',
                    links: { total: 2, temporal: 1, entity: 1, causal: 0, semantic: 0 },
                }],
            }],
        }],
    }, null, 2), 'utf8');

    const memory = createMemory({ root, memory: {} });
    const item = (await memory.archiveItems()).find((row) => row.mnemonId === 'i-9');
    assert.equal(item.id, 'mnemon:i-9', '读入时就该补出 id');
    assert.equal(item.linkedCount, 2, 'links.total 要兜底成 linkedCount');

    const hot = await memory.mutate({ action: 'add', target: 'project', content: '一条热记忆。' });
    assert.equal((await memory.link({ sourceId: 'mnemon:i-9', targetId: hot.entry.id })).added, true, 'entryIndex 不该再跳过它');
    assert.ok((await memory.related({ id: 'mnemon:i-9' })).nodes.some((node) => node.id === hot.entry.id));

    // 下一次写入把补出来的值固化进**结构化的那份正文**（老库不必重迁一次）。
    // 索引里只留指针、不再内联正文 —— 这正是第一梯队 24-9 要的那条去冗余。
    await memory.appendArchiveBlock({ reason: 'test', items: [{ id: 'x-probe', kind: 'hot', content: '新的。' }], month: '2026-09' });
    const saved = JSON.parse(await readFile(paths.index, 'utf8'));
    const savedDigest = saved.digests.find((row) => row.month === '2026-09');
    assert.ok(savedDigest.itemFile.endsWith('.jsonl'), '索引要写明结构化正文的文件');
    assert.equal(savedDigest.blocks.some((block) => 'items' in block), false, '索引里只留指针，不再内联正文');
    assert.equal(savedDigest.blocks[0].count, 1, '指针要写着这一块有几条');
    assert.deepEqual(savedDigest.blocks[0].ids, ['mnemon:i-9'], '指针要挂着条目 id（liveIds 只读索引）');
    const itemFile = await readFile(join(paths.archive, '2026-09.jsonl'), 'utf8');
    const persisted = itemFile.split('\n').filter((line) => line.trim() !== '')
        .map((line) => JSON.parse(line)).flatMap((block) => block.items).find((row) => row.mnemonId === 'i-9');
    assert.equal(persisted.id, 'mnemon:i-9');
    assert.equal(persisted.linkedCount, 2);
});

await check('归档补 id 的兜底会连人可读的摘要一起修：写别的月份也重写老月份的 md', async () => {
    // 只修 index.json 是不够的（第二十七轮审查发现）：老月份的 .md 只有被重写才会
    // 带上「关系边 N 条」，而 appendArchive 原先只写「这次被触碰到的月份」。
    const { root } = await freshRoot();
    const paths = memoryPaths(root, {});
    await mkdir(paths.archive, { recursive: true });
    await writeFile(paths.index, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-01T00:00:00.000Z',
        digests: [{
            id: 'a-2026-09', month: '2026-09', at: '2026-09-01T00:00:00.000Z', file: 'archive/2026-09.md',
            blocks: [{
                id: 'b-legacy', at: '2026-09-01T00:00:00.000Z', reason: 'mnemon-migration',
                items: [{ kind: 'mnemon', mnemonId: 'i-7', content: '老月份里的一条长期记忆。', links: { total: 4, temporal: 1, entity: 1, causal: 1, semantic: 1 } }],
            }],
        }],
    }, null, 2), 'utf8');
    // 先写一份「还没有边数」的旧 md（就是真实库里的样子）
    await writeFile(join(paths.archive, '2026-09.md'), '# 归档 2026-09\n\n## 2026-09-01 · mnemon 迁移（1 条）\n\n- [fact] 老月份里的一条长期记忆。\n', 'utf8');

    const memory = createMemory({ root, memory: {} });
    // 写一个**别的月份**：老月份没被触碰，但它的摘要仍然要被修好
    await memory.appendArchiveBlock({ reason: 'test', items: [{ id: 'y-probe', kind: 'hot', content: '十月的。', at: '2026-10-01T00:00:00.000Z' }], month: '2026-10' });
    const md = await readFile(join(paths.archive, '2026-09.md'), 'utf8');
    assert.ok(md.includes('关系边 4 条'), `老月份的摘要要补上边数\n${md}`);

    // 再写一次同样的内容：md 不该被无谓地重写（渲染一致就不落盘）
    const before = (await readFile(join(paths.archive, '2026-09.md'), 'utf8'));
    await memory.appendArchiveBlock({ reason: 'test', items: [{ id: 'z-probe', kind: 'hot', content: '十月的第二条。', at: '2026-10-02T00:00:00.000Z' }], month: '2026-10' });
    assert.equal(await readFile(join(paths.archive, '2026-09.md'), 'utf8'), before, '内容一致时不该重写');
});

// ── 第二十八轮：删除的引用完整性与时间窗 ────────────────────────────────────

await check('删除引用完整性：边同事务清掉、墓碑写上、related 不再穿过已删节点', async () => {
    const { memory } = await freshRoot();
    const a = await memory.mutate({ action: 'add', target: 'project', content: '甲：链路的一端。' });
    const b = await memory.mutate({ action: 'add', target: 'project', content: '乙：中间那条，等下要删。' });
    const c = await memory.mutate({ action: 'add', target: 'project', content: '丙：只有穿过乙才够得到。' });
    await memory.link({ sourceId: a.entry.id, targetId: b.entry.id });
    await memory.link({ sourceId: b.entry.id, targetId: c.entry.id });
    const before = await memory.related({ id: a.entry.id, depth: 2 });
    assert.deepEqual(before.nodes.map((node) => node.id).sort(), [b.entry.id, c.entry.id].sort(), '删之前两跳应当看得见丙');

    const removed = await memory.mutate({ action: 'remove', id: b.entry.id });
    assert.equal(removed.action, 'remove');
    assert.equal(removed.prunedLinks, 2, '落在乙身上的两条边都要清掉');
    assert.equal(removed.tombStoned, true, '删除要留墓碑');
    assert.equal((await memory.danglingRefs()).length, 0, '不变量：悬空引用必须为 0');

    // 幽灵桥：乙被删之后，丙不该再出现在 related(甲) 里
    const after = await memory.related({ id: a.entry.id, depth: 2 });
    assert.deepEqual(after.nodes, [], '不能穿过已删节点继续遍历');
    assert.equal(after.danglingCount, 0);

    // 墓碑：删掉的 id 不能再建关系，也不能再遍历，且错误信息要说清是「已被删除」
    await assert.rejects(() => memory.link({ sourceId: a.entry.id, targetId: b.entry.id }), /已经被删除/);
    await assert.rejects(() => memory.related({ id: b.entry.id }), /已经被删除/);

    const status = await memory.status();
    assert.equal(status.dangling, 0);
    assert.equal(status.tombstones, 1);
});

await check('悬空边：related 如实报出且不穿过它（老数据 / 外部改过文件的兜底）', async () => {
    const { root, memory } = await freshRoot();
    const a = await memory.mutate({ action: 'add', target: 'project', content: '甲。' });
    const c = await memory.mutate({ action: 'add', target: 'project', content: '丙。' });
    const paths = memoryPaths(root, {});
    // 直接造两条边：甲—幽灵、幽灵—丙。这就是「删了条目忘了清边」的老数据形状。
    await mkdir(paths.dir, { recursive: true });
    const linkLine = (id, sourceId, targetId) => JSON.stringify({ id, sourceId, targetId, kind: 'related', note: '', at: '2026-09-01T00:00:00.000Z' });
    await writeFile(paths.links, `${linkLine('K-1', a.entry.id, 'm-ghost')}\n${linkLine('K-2', 'm-ghost', c.entry.id)}\n`, 'utf8');

    const value = await memory.related({ id: a.entry.id, depth: 2 });
    assert.deepEqual(value.nodes, [], '幽灵节点不能当桥走到丙');
    assert.equal(value.danglingCount, 1, '遍历中撞到的悬空对端只报一次');
    assert.equal(value.dangling[0].to, 'm-ghost');
    // status 数的是**全库的悬空边条数**（两条边都有一端指着幽灵），与遍历撞到的数不是同一个量
    assert.equal((await memory.status()).dangling, 2);
    // 渲染里必须出现悬空提示与处理办法
    const text = renderRelated(value);
    assert.ok(text.includes('悬空边'), text);

    await memory.unlink({ id: 'K-1' });
    assert.equal((await memory.danglingRefs()).length, 1, 'K-2 还指着幽灵');
    await memory.unlink({ id: 'K-2' });
    assert.equal((await memory.danglingRefs()).length, 0);
});

await check('Pack 与墓碑：删掉的条目不会被旧 Pack 导入复活', async () => {
    const { memory } = await freshRoot();
    const keep = await memory.mutate({ action: 'add', target: 'project', content: '留着的那条。' });
    const gone = await memory.mutate({ action: 'add', target: 'project', content: '等下删掉的那条。' });
    await memory.link({ sourceId: keep.entry.id, targetId: gone.entry.id });
    const pack = await memory.exportPack({});
    assert.equal(pack.stores[0].tombstones.length, 0, '导出时还没有墓碑');

    await memory.mutate({ action: 'remove', id: gone.entry.id });
    const report = await memory.importPack(JSON.parse(JSON.stringify(pack)));
    assert.ok(report.skippedTombstoned >= 1, `落在墓碑上的条目与边都不该搬回来：${JSON.stringify(report)}`);
    const hot = await memory.read({ layer: 'hot' });
    assert.ok(!hot.hot.items.some((item) => item.id === gone.entry.id), '删掉的条目不能复活');
    assert.ok(hot.hot.items.some((item) => item.id === keep.entry.id), '没被删的条目照旧');
    assert.equal((await memory.danglingRefs()).length, 0, '导入也不能造出悬空边');

    // 墓碑随 Pack 走：另一份库导入同一份包时，才知道这条被删过
    const pack2 = await memory.exportPack({});
    assert.equal(pack2.stores[0].tombstones.length, 1, '导出会把墓碑带上');
    const other = await freshRoot();
    const otherReport = await other.memory.importPack(JSON.parse(JSON.stringify(pack2)));
    assert.equal(otherReport.tombstones, 1, 'Pack 里的墓碑要一起搬');
    assert.equal(otherReport.skippedTombstoned, 0);
    const otherHot = await other.memory.read({ layer: 'hot' });
    assert.equal(otherHot.hot.total, 1, '只搬没被删的那条');
    assert.equal((await other.memory.status()).tombstones, 1);
});

await check('时间窗：三层都按时刻过滤，窗外与窗内分得清（不谎报「空」）', async () => {
    const { root } = await freshRoot();
    const paths = memoryPaths(root, {});
    await mkdir(paths.archive, { recursive: true });
    const entry = (id, content, updatedAt) => ({ id, target: 'project', content, importance: 'normal', entities: [], tags: [], createdAt: updatedAt, updatedAt });
    await writeFile(paths.memory, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-01T00:00:00.000Z',
        entries: [
            entry('m-old', '八月记的。', '2026-08-01T10:00:00.000Z'),
            entry('m-mid', '九月中旬记的。', '2026-09-15T10:00:00.000Z'),
            entry('m-new', '九月底记的。', '2026-09-25T10:00:00.000Z'),
        ],
    }, null, 2), 'utf8');
    await writeFile(paths.ledger, [
        JSON.stringify({ id: 'L-old', at: '2026-08-02T10:00:00.000Z', path: '旧.docx', format: 'word' }),
        JSON.stringify({ id: 'L-new', at: '2026-09-20T10:00:00.000Z', path: '新.docx', format: 'word' }),
    ].join('\n') + '\n', 'utf8');
    await writeFile(paths.index, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-26T00:00:00.000Z',
        digests: [{
            id: 'a-2026-09', month: '2026-09', at: '2026-09-26T00:00:00.000Z', file: 'archive/2026-09.md',
            blocks: [{
                id: 'b-1', at: '2026-09-26T00:00:00.000Z', reason: 'mnemon-migration',
                items: [
                    { id: 'mnemon:x', kind: 'mnemon', content: '八月的归档条目。', at: '2026-08-05T00:00:00.000Z', category: 'fact' },
                    { id: 'mnemon:y', kind: 'mnemon', content: '九月的归档条目。', at: '2026-09-22T00:00:00.000Z', category: 'fact' },
                ],
            }],
        }],
    }, null, 2), 'utf8');

    const memory = createMemory({ root, memory: {} });
    assert.equal((await memory.read({ layer: 'hot' })).window, undefined, '不带时间窗时结果里不该多出 window');

    const windowed = await memory.read({ layer: 'all', since: '2026-09-01', until: '2026-09-30' });
    assert.equal(windowed.window.sinceInput, '2026-09-01');
    // 同重要度内是 **recency 弱先验**（新的在前，第二十四轮 24-4）：所以九月底那条排在
    // 九月中那条前面。八月那条要被挡在窗外。
    assert.deepEqual(windowed.hot.items.map((item) => item.id), ['m-new', 'm-mid'], '八月那条要被挡在窗外');
    assert.equal(windowed.hot.total, 2);
    assert.equal(windowed.hot.windowFiltered, 1);
    assert.equal(windowed.ledger.total, 1);
    assert.equal(windowed.ledger.items[0].id, 'L-new');
    assert.equal(windowed.archive.items_total, 1);
    assert.equal(windowed.archive.items[0].id, 'mnemon:y');
    const text = renderRead(windowed);
    assert.ok(text.includes('时间窗：2026-09-01 ~ 2026-09-30'), text);

    // 窗外：三层都是 0 条，但要说清「都在窗外」，不能说成「（空）」
    const outside = await memory.read({ layer: 'all', since: '2026-07-01', until: '2026-07-31' });
    assert.equal(outside.hot.total, 0);
    assert.equal(outside.ledger.total, 0);
    assert.equal(outside.archive.items_total, 0);
    const outsideText = renderRead(outside);
    assert.ok(outsideText.includes('热记忆3 条在时间窗之外'), outsideText);
    assert.ok(outsideText.includes('（2 条在时间窗之外（可以放宽 since / until））'), outsideText);
});

await check('时间词解析：固定的 now 下逐词可对，且解析不出来会报错（不静默忽略）', async () => {
    const now = new Date(2026, 8, 27, 15, 30, 0);
    const win = (options) => resolveTimeWindow({ ...options, now });
    const day = (value) => new Date(value);

    const month = win({ since: '本月' });
    assert.equal(day(month.since).getDate(), 1, '本月从 1 号开始');
    assert.equal(day(month.since).getHours(), 0);
    assert.equal(month.until, null, '只给 since 时 until 是 null（不设上界）');

    const lastWeek = win({ since: '上周', until: '上周' });
    assert.equal(day(lastWeek.since).getDay(), 1, '上周从周一开始');
    assert.equal(day(lastWeek.until).getDay(), 0, '上周到周日结束');
    const thisWeek = win({ since: '本周' });
    assert.ok(thisWeek.since > lastWeek.until, '本周严格晚于上周');

    const yesterday = win({ since: '昨天' });
    assert.equal(yesterday.since, new Date(2026, 8, 26, 0, 0, 0, 0).getTime());

    // 第四十七轮（复核 F8）：新导出的词表里**新增**的那几个字也要钉值 ——
    // 只验「能解析」的话，把「前天」写成 +2 天也照样全绿。
    const beforeYesterday = win({ since: '前天' });
    assert.equal(beforeYesterday.since, new Date(2026, 8, 25, 0, 0, 0, 0).getTime(), '前天是 -2 天');
    const tomorrow = win({ until: '明天' });
    assert.equal(tomorrow.until, new Date(2026, 8, 28, 23, 59, 59, 999).getTime(), '明天是 +1 天的整天');
    const lastMonth = win({ since: '上个月' });
    assert.equal(day(lastMonth.since).getMonth(), 7, '上个月是 8 月');
    assert.equal(day(lastMonth.since).getDate(), 1, '上个月从 1 号开始');
    // 同义写法（提示里不列，但词表要认）。
    assert.equal(win({ since: 'today' }).since, win({ since: '今天' }).since);
    assert.equal(win({ since: '这周' }).since, win({ since: '本周' }).since);

    // 原型上的名字不能被当成时间词（复核 F15）：`constructor` / `__proto__` 要落到
    // 「看不懂的时间」，而不是 TypeError。
    assert.throws(() => win({ since: 'constructor' }), /看不懂的时间/);
    assert.throws(() => win({ since: '__proto__' }), /看不懂的时间/);

    const days = win({ since: '最近7天' });
    assert.equal(days.since, new Date(2026, 8, 21, 0, 0, 0, 0).getTime(), '最近 7 天含今天，从 6 天前算起');

    const exact = win({ since: '2026-09-27T10:30' });
    assert.equal(day(exact.since).getHours(), 10);
    assert.equal(day(exact.since).getMinutes(), 30);

    const untilDate = win({ until: '2026-09-27' });
    assert.equal(day(untilDate.until).getHours(), 23, '只给日期时 until 取当天最后一刻（闭区间）');

    assert.throws(() => win({ since: '前天早上' }), /看不懂的时间/);
    assert.throws(() => win({ since: '2026-09-30', until: '2026-09-01' }), /时间窗是空的/);
});

await check('office_memory 工具：since / until 真的接到底层，解析失败如实报错', async () => {
    const { root } = await freshRoot();
    const config = resolveConfig({ memory: { globalDir: join(TMP, 'global-window') } });
    const tool = buildTools(config, {}).find((item) => item.name === 'office_memory');
    const exec = execAt(root);
    await tool.execute({ action: 'add', target: 'project', content: '今天记的：模板用 A。' }, exec);

    const hit = await tool.execute({ action: 'read', layer: 'hot', since: '今天' }, exec);
    assert.equal(hit.ok, true);
    assert.match(hit.text, /时间窗：今天/, '工具面要把时间窗回执透出来');
    assert.match(hit.text, /模板用 A/);

    const miss = await tool.execute({ action: 'read', layer: 'hot', since: '昨天', until: '昨天' }, exec);
    assert.equal(miss.ok, true);
    assert.match(miss.text, /在时间窗之外/);

    // 解析不出来：与记忆的其它参数错误同一口径 —— 抛错（宿主会把消息带给模型），
    // **不能**退化成「没有命中的条目」。
    await assert.rejects(() => tool.execute({ action: 'read', layer: 'hot', since: '前天早上' }, exec), /看不懂的时间/);
});

// ── 第二十八轮（独立审查后补）：审查挖出的覆盖空洞 ──────────────────────────

await check('归档封顶删月：清边 + 墓碑 + 正文真删 + 悬空 0（此前这条路径零覆盖）', async () => {
    const { memory } = await freshRoot({ archiveKeep: 1 });
    const keep = await memory.mutate({ action: 'add', target: 'project', content: '乙：还活着的那条。' });
    // 只存在于归档里的 id（模拟 mnemon 迁来的历史长期记忆）
    await memory.appendArchiveBlock({
        reason: 'mnemon-migration',
        month: '2026-07',
        items: [{ id: 'mnemon:old', kind: 'mnemon', content: '七月的老条目。', at: '2026-07-01T00:00:00.000Z' }],
    });
    const linked = await memory.link({ sourceId: 'mnemon:old', targetId: keep.entry.id });
    assert.equal(linked.added, true);
    assert.equal((await memory.danglingRefs()).length, 0);

    // 再写一个月 → archiveKeep=1 把 2026-07 挤掉：边清掉、墓碑写上、正文文件删掉
    await memory.appendArchiveBlock({
        reason: 'test',
        month: '2026-08',
        items: [{ id: 'x-august', kind: 'hot', content: '八月的，留下。', at: '2026-08-01T00:00:00.000Z' }],
    });
    const status = await memory.status();
    assert.equal(status.dangling, 0, '删月之后不能留下悬空边');
    assert.equal(status.tombstones, 1, '删月要留一条墓碑');
    assert.equal((await memory.read({ layer: 'archive' })).archive.items_total, 1, '只剩八月那一块');
    assert.equal(existsSync(join(memory.paths.archive, '2026-07.md')), false, '封顶删月要删掉正文文件（删除不是假的）');
    assert.equal(existsSync(join(memory.paths.archive, '2026-07.jsonl')), false,
        '结构化正文也要一起删：只删 .md 而把 .jsonl 留着，正文还读得到，「删最旧卷」就是假删除');
    assert.equal(existsSync(join(memory.paths.archive, '2026-08.md')), true);
    assert.equal(existsSync(join(memory.paths.archive, '2026-08.jsonl')), true, '留下的那卷两份投影都要在');
    // 归档**参与召回**（layer:'archive' 与 'all' 都会给归档条目）—— 所以「删掉」必须在
    // 召回面上也成立：不是「这次没命中」，是内容真的不在库里了（24-16 要的那句声明）。
    // 口径说明：`items_total` 是**过筛前**落在时间窗里的条数，命中数看 `matched`。
    const ghost = await memory.read({ layer: 'archive', query: '老条目' });
    assert.equal(ghost.archive.matched, 0, '被删掉的卷不能再被召回（命中 0）');
    assert.equal(ghost.archive.items.length, 0);
    assert.equal(ghost.archive.items_total, 1, '剩下的那一条仍在库里，只是不命中这个词');
    const leftOver = (await readdir(memory.paths.archive)).filter((name) => name.startsWith('2026-07'));
    assert.deepEqual(leftOver, [], '2026-07 不该在归档目录里留下任何文件');
    // 被删的那条不可寻址、不可建关系，而且报错说清是删除
    await assert.rejects(() => memory.link({ sourceId: 'mnemon:old', targetId: keep.entry.id }), /已经被删除/);
});

// ── 第四十轮：归档分卷（14-1）与索引去冗余（24-9） ──────────────────────────

await check('归档分卷：单月超上限开下一卷，检索 / 保留 / 删除都按卷走（14-1）', async () => {
    const { root, memory } = await freshRoot({ archiveKeep: 10 });
    const paths = memoryPaths(root, {});
    // 一块本身超上限（台账一次滚动可能带几百条）**不切开**：块是去重与导入的单位。
    await memory.appendArchiveBlock({
        reason: 'ledger-rollover',
        month: '2026-09',
        items: Array.from({ length: ARCHIVE_VOLUME_MAX_ITEMS + 1 }, (_, i) => ({ id: `big-${i}`, kind: 'hot', content: `九月第一批第 ${i} 条。` })),
    });
    assert.equal(existsSync(join(paths.archive, '2026-09.md')), true, '第 1 卷用月份当文件名');
    assert.equal(existsSync(join(paths.archive, '2026-09-2.md')), false, '单块超上限时不切开');
    // 再来一块：当前卷已满，这一块落到第 2 卷。
    await memory.appendArchiveBlock({
        reason: 'test',
        month: '2026-09',
        items: [{ id: 'v2-1', kind: 'hot', content: '第二卷的那条。', at: '2026-09-20T00:00:00.000Z' }],
    });
    assert.equal(existsSync(join(paths.archive, '2026-09-2.md')), true, '第 2 卷的 Markdown 要开出来');
    assert.equal(existsSync(join(paths.archive, '2026-09-2.jsonl')), true, '第 2 卷的结构化正文要一起开');
    const index = JSON.parse(await readFile(paths.index, 'utf8'));
    assert.deepEqual(index.digests.map((digest) => digest.volume), [1, 2], '索引里两卷各有卷号');
    assert.deepEqual(index.digests.map((digest) => digest.itemFile), ['archive/2026-09.jsonl', 'archive/2026-09-2.jsonl']);
    const md2 = await readFile(join(paths.archive, '2026-09-2.md'), 'utf8');
    assert.ok(md2.includes('第 2 卷'), `第 2 卷的标题要写出卷号，否则两个文件看起来一模一样\n${md2}`);
    const all = await memory.read({ layer: 'archive' });
    assert.equal(all.archive.digests, 2, '两卷都算摘要文件（archiveKeep 管的就是这个数）');
    assert.equal(all.archive.items_total, ARCHIVE_VOLUME_MAX_ITEMS + 2, '一条不丢');
    assert.equal(all.archive.damagedVolumes, 0);
    // 两块分别落在两卷里，检索都要能命中（分卷不是分家）。命中数看 `matched`，
    // `items_total` 是过筛前落在窗内的条数（202 条都在库里）。
    assert.ok((await memory.read({ layer: 'archive', query: '第一批第 7 条' })).archive.matched >= 1, '第 1 卷里的条目要检索得到');
    assert.equal((await memory.read({ layer: 'archive', query: '第二卷的那条' })).archive.matched, 1, '第 2 卷里的条目也要检索得到');

    // 容量管的是**卷**：挤到只剩最新一卷时，同月的第 1 卷先走，且两份投影一起走。
    const tight = createMemory({ root, memory: { archiveKeep: 1 } });
    await tight.appendArchiveBlock({ reason: 'test', month: '2026-10', items: [{ id: 'oct-1', kind: 'hot', content: '十月那条。' }] });
    const left = (await readdir(paths.archive)).filter((name) => name !== 'index.json').sort();
    assert.deepEqual(left, ['2026-10.jsonl', '2026-10.md'], `archiveKeep=1 只该留下最新那一卷，实际 ${left.join('、')}`);
    assert.equal((await memory.read({ layer: 'archive' })).archive.items_total, 1);
});

await check('归档索引去冗余：条目数与块数不变时，索引不随正文长度膨胀（24-9）', async () => {
    // 只比「结构 + 计数」：随机 id 与时刻天然不同，不参与比较。
    const shapeOf = async (root) => {
        const raw = JSON.parse(await readFile(memoryPaths(root, {}).index, 'utf8'));
        return JSON.stringify({
            digests: raw.digests.map((digest) => ({
                month: digest.month,
                volume: digest.volume,
                file: digest.file,
                itemFile: digest.itemFile,
                blocks: digest.blocks.map((block) => ({ reason: block.reason, count: block.count, ids: block.ids })),
            })),
        });
    };
    const short = await freshRoot({ archiveKeep: 10 });
    await short.memory.appendArchiveBlock({ reason: 'test', month: '2026-09', items: [{ id: 'same-1', kind: 'hot', content: '短。' }] });
    const long = await freshRoot({ archiveKeep: 10 });
    await long.memory.appendArchiveBlock({ reason: 'test', month: '2026-09', items: [{ id: 'same-1', kind: 'hot', content: '长。'.repeat(3000) }] });

    assert.equal(await shapeOf(long.root), await shapeOf(short.root),
        '同样的条目数下索引形状必须完全一样 —— 索引不该随正文长度变');
    const longItems = await readFile(join(memoryPaths(long.root, {}).archive, '2026-09.jsonl'), 'utf8');
    const shortItems = await readFile(join(memoryPaths(short.root, {}).archive, '2026-09.jsonl'), 'utf8');
    assert.ok(longItems.length > shortItems.length + 3000, '正文长度差要落在结构化的那份正文里');
    const indexText = await readFile(memoryPaths(short.root, {}).index, 'utf8');
    assert.equal(indexText.includes('短。'), false, '索引里不该出现正文（否则上面那条只是碰巧相等）');
    // 去冗余不是丢内容：正文读得到，检索照样命中。
    const found = await long.memory.read({ layer: 'archive', query: '长' });
    assert.equal(found.archive.items_total, 1);
    assert.ok(found.archive.items[0].text.startsWith('长。'), '检索回执给的还是正文开头');
});

// ── 第四十轮（独立验证挖出的反例）：半状态、受损卷与去重键 ────────────────────

await check('归档半状态：索引还是老格式、.jsonl 已经多了一块时，以 .jsonl 为准（不丢块）', async () => {
    // 这正是「.jsonl 已落、索引还没落」的崩溃窗口，也是升级后第一次写入前的半状态。
    const { root } = await freshRoot();
    const paths = memoryPaths(root, {});
    await mkdir(paths.archive, { recursive: true });
    await writeFile(paths.index, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-01T00:00:00.000Z',
        digests: [{
            id: 'a-2026-09', month: '2026-09', at: '2026-09-01T00:00:00.000Z', file: 'archive/2026-09.md',
            blocks: [{
                id: 'b-legacy', at: '2026-09-01T00:00:00.000Z', reason: 'mnemon-migration',
                items: [{ id: 'legacy-1', kind: 'hot', content: '老那份。' }],
            }],
        }],
    }, null, 2), 'utf8');
    await writeFile(join(paths.archive, '2026-09.jsonl'), [
        JSON.stringify({ id: 'b-legacy', at: '2026-09-01T00:00:00.000Z', reason: 'mnemon-migration', items: [{ id: 'legacy-1', kind: 'hot', content: '老那份。' }] }),
        JSON.stringify({ id: 'b-crash', at: '2026-09-02T00:00:00.000Z', reason: 'hot-overflow', items: [{ id: 'crash-1', kind: 'hot', content: '崩在两步之间的那块。' }] }),
        '',
    ].join('\n'), 'utf8');

    const memory = createMemory({ root, memory: {} });
    // 写**另一个月**：这一步会重写所有卷的 .jsonl —— 丢内容的时刻就在这儿。
    await memory.appendArchiveBlock({ reason: 'test', month: '2026-10', items: [{ id: 'oct-1', kind: 'hot', content: '十月的。' }] });
    const items = await memory.archiveItems();
    assert.ok(items.some((item) => item.id === 'crash-1'), 'jsonl 里比索引多的那块不能被丢掉');
    assert.ok(items.some((item) => item.id === 'legacy-1'), '索引里那块也要还在');
    assert.ok((await readFile(join(paths.archive, '2026-09.jsonl'), 'utf8')).includes('crash-1'), '重写后的 .jsonl 里也要有它');
    assert.equal((await memory.read({ layer: 'archive' })).archive.items_total, 3);
    assert.equal((await memory.read({ layer: 'archive' })).archive.damagedVolumes, 0, '这是索引落后，不是正文受损');
});

await check('归档受损卷被隔离：删掉 .jsonl 之后，后续写入不会把 .md 与人读正文一起抹平', async () => {
    const { memory } = await freshRoot({ archiveKeep: 10 });
    const keep = await memory.mutate({ action: 'add', target: 'project', content: '乙：活着的那条。' });
    await memory.appendArchiveBlock({ reason: 'mnemon-migration', month: '2026-07', items: [{ id: 'mnemon:lost', kind: 'mnemon', content: '七月的老条目。' }] });
    await memory.link({ sourceId: 'mnemon:lost', targetId: keep.entry.id });
    const mdBefore = await readFile(join(memory.paths.archive, '2026-07.md'), 'utf8');
    // 模拟「盘上的正文文件被人删了」：索引还列着这一卷
    await rm(join(memory.paths.archive, '2026-07.jsonl'), { force: true });
    assert.equal((await memory.read({ layer: 'archive' })).archive.damagedVolumes, 1, '读要如实报受损');

    // 写一个**别的月份**：以前这一步会拿指针块重写正文，把受损卷的 .jsonl 造成空文件、
    // 把唯一还留着内容的那份 .md 抹成空摘要，damage 也随之洗成 0。
    await memory.appendArchiveBlock({ reason: 'test', month: '2026-08', items: [{ id: 'aug-1', kind: 'hot', content: '八月的。' }] });
    assert.equal(await readFile(join(memory.paths.archive, '2026-07.md'), 'utf8'), mdBefore, '受损卷的人读投影不该被重写');
    assert.equal(existsSync(join(memory.paths.archive, '2026-07.jsonl')), false, '不该凭空造一个空正文文件');
    assert.equal((await memory.read({ layer: 'archive' })).archive.damagedVolumes, 1, '受损状态不能被洗白成 0');
    assert.equal((await memory.danglingRefs()).length, 0, '指向受损卷条目的边不该变成悬空边');
    assert.equal((await memory.status()).stores[0].archiveDamaged, 1, 'status 也要报得出这个根有一卷受损');
    // 新内容照样写得进去：受损的那一卷被隔离，接下来的块开新卷
    assert.equal(existsSync(join(memory.paths.archive, '2026-08.jsonl')), true);
});

await check('Pack 去重：块里含已被墓碑挡掉的条目时，重复导入仍不产生第二份', async () => {
    const { memory } = await freshRoot();
    const dead = await memory.mutate({ action: 'add', target: 'project', content: '乙：等下删掉，Pack 里还带着它。' });
    await memory.mutate({ action: 'remove', id: dead.entry.id });
    const pack = {
        format: 'dsh-office-memory-pack',
        version: 1,
        exportedAt: '2026-09-06T00:00:00.000Z',
        stores: [{
            id: 'workspace',
            entries: [],
            ledger: [],
            links: [],
            tombstones: [],
            archiveFiles: {},
            archiveIndex: {
                version: 1,
                updatedAt: '2026-09-06T00:00:00.000Z',
                digests: [{
                    id: 'a-2026-09', month: '2026-09', at: '2026-09-06T00:00:00.000Z', file: 'archive/2026-09.md',
                    blocks: [{
                        at: '2026-09-06T00:00:00.000Z',
                        reason: 'ledger-rollover',
                        items: [
                            { id: dead.entry.id, kind: 'hot', content: '乙的归档副本。' },
                            { id: 'survivor-1', kind: 'hot', content: '幸存者。' },
                        ],
                    }],
                }],
            },
        }],
    };
    const first = await memory.importPack(JSON.parse(JSON.stringify(pack)));
    assert.equal(first.archiveBlocks, 1, '一块要搬进来');
    // 关键：块的 id 必须按**入包声明的两条**算，不能按墓碑过滤后的一条算 ——
    // 否则第二次导入算出的键与库里那条不同，同一条内容会被搬两份。
    const second = await memory.importPack(JSON.parse(JSON.stringify(pack)));
    assert.equal(second.archiveBlocks, 0, '含被墓碑挡掉的条目时也要幂等');
    const items = await memory.archiveItems();
    assert.equal(items.filter((item) => item.id === 'survivor-1').length, 1, '幸存者只该有一份');
    assert.equal(items.some((item) => item.id === dead.entry.id), false, '被墓碑挡掉的条目不该搬回来');
});

await check('Pack 去重：同月两块形状键相同但内容不同时，两块都要在（别静默丢一块）', async () => {
    const { memory } = await freshRoot();
    // 同 at / 同 reason / 同条数：老实现给它们算出同一个「时刻+原因+条数」键，
    // 第二块会被去重逻辑判成「已存在」而**静默丢掉**。内容寻址的 id 让它们各自身份不同。
    const shape = (content) => ({ at: '2026-09-06T00:00:00.000Z', reason: 'ledger-rollover', items: [{ id: `item-${content}`, kind: 'hot', content }] });
    const pack = {
        format: 'dsh-office-memory-pack',
        version: 1,
        exportedAt: '2026-09-06T00:00:00.000Z',
        stores: [{
            id: 'workspace',
            entries: [],
            ledger: [],
            links: [],
            tombstones: [],
            archiveFiles: {},
            archiveIndex: { version: 1, updatedAt: '2026-09-06T00:00:00.000Z', digests: [{ id: 'a-2026-09', month: '2026-09', at: '2026-09-06T00:00:00.000Z', file: 'archive/2026-09.md', blocks: [shape('甲'), shape('乙')] }] },
        }],
    };
    const report = await memory.importPack(JSON.parse(JSON.stringify(pack)));
    assert.equal(report.archiveBlocks, 2, '内容不同的两块都要入库');
    const items = await memory.archiveItems();
    assert.ok(items.some((item) => item.content === '甲') && items.some((item) => item.content === '乙'), '两块内容都要在');
    const second = await memory.importPack(JSON.parse(JSON.stringify(pack)));
    assert.equal(second.archiveBlocks, 0, '再导一次仍然幂等');
    assert.equal((await memory.archiveItems()).length, 2);
});

await check('归档受损的两个口径：老格式不算受损（P3），0 字节正文算受损（P4）', async () => {
    // ① 老格式（正文还内联在索引里、还没有 .jsonl）：读得出来，status 不该报受损。
    //    不排除这一条的话，每个还没迁移过的库都会被报成「有一卷受损」——
    //    真实库就是这样（read 说 88/88 可读，status 却说受损 1 卷）。
    const { root: legacyRoot, memory: legacy } = await freshRoot();
    const legacyPaths = memoryPaths(legacyRoot, {});
    await mkdir(legacyPaths.archive, { recursive: true });
    await writeFile(legacyPaths.index, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-01T00:00:00.000Z',
        digests: [{
            id: 'a-2026-09', month: '2026-09', at: '2026-09-01T00:00:00.000Z', file: 'archive/2026-09.md',
            blocks: [{ at: '2026-09-01T00:00:00.000Z', reason: 'mnemon-migration', items: [{ id: 'legacy-1', kind: 'hot', content: '老格式那条。' }] }],
        }],
    }, null, 2), 'utf8');
    const legacyRead = await legacy.read({ layer: 'archive' });
    assert.equal(legacyRead.archive.items_total, 1, '老格式照样读得出来');
    assert.equal(legacyRead.archive.damagedVolumes, 0, '读侧不该把它当受损');
    assert.equal((await legacy.status()).stores[0].archiveDamaged, 0, 'status 也不该 —— 它只是还没固化');

    // ② 指针格式 + 0 字节（被截断）的正文：两个口径都要算受损，别一个报一个不报。
    const { memory } = await freshRoot();
    await memory.appendArchiveBlock({ reason: 'test', month: '2026-09', items: [{ id: 'x-1', kind: 'hot', content: '一条。' }] });
    await writeFile(join(memory.paths.archive, '2026-09.jsonl'), '', 'utf8');
    assert.equal((await memory.read({ layer: 'archive' })).archive.damagedVolumes, 1, 'read 要报');
    assert.equal((await memory.status()).stores[0].archiveDamaged, 1, 'status 也要报');
});

await check('墓碑不锁死活着的 id：同 id 重新活过来之后还能建关系（alive wins）', async () => {
    const { memory } = await freshRoot();
    const x = await memory.mutate({ action: 'add', target: 'project', content: '甲：等下删掉再让它回来。' });
    const y = await memory.mutate({ action: 'add', target: 'project', content: '乙：一直活着。' });
    await memory.link({ sourceId: x.entry.id, targetId: y.entry.id });
    await memory.mutate({ action: 'remove', id: x.entry.id });
    assert.equal((await memory.status()).tombstones, 1);
    await assert.rejects(() => memory.link({ sourceId: x.entry.id, targetId: y.entry.id }), /已经被删除/);

    // 让同一个 id 重新活起来（迁移重跑 / 归档块合并会发生这件事）
    await memory.appendArchiveBlock({
        reason: 'mnemon-migration',
        month: '2026-08',
        items: [{ id: x.entry.id, kind: 'mnemon', content: '同 id 又回来了。', at: '2026-08-01T00:00:00.000Z' }],
    });
    assert.equal((await memory.liveIds()).has(x.entry.id), true, '它现在活着');
    const again = await memory.link({ sourceId: x.entry.id, targetId: y.entry.id });
    assert.equal(again.added, true, '活着的条目不该被陈旧墓碑永久锁死');
    assert.equal((await memory.danglingRefs()).length, 0);
});

await check('Pack 导入：指着不存在条目的边不搬进来，自带墓碑的 Pack 也不会先搬边', async () => {
    const { memory } = await freshRoot();
    const keep = await memory.mutate({ action: 'add', target: 'project', content: '留下的一条。' });
    const packOf = (stores) => ({ format: 'dsh-office-memory-pack', version: 1, exportedAt: '2026-09-27T00:00:00.000Z', stores });

    // ① 老 Pack 的悬空边（本轮之前 remove 不清边，所以这是常态）
    const stale = await memory.importPack(packOf([{
        id: 'workspace',
        entries: [],
        ledger: [],
        links: [{ id: 'K-x', sourceId: keep.entry.id, targetId: 'm-nope', kind: 'related', note: '', at: '2026-09-01T00:00:00.000Z' }],
        archiveIndex: { digests: [] },
        tombstones: [],
    }]));
    assert.equal(stale.links, 0, '悬空边不该被搬进来');
    assert.equal(stale.skippedDangling, 1);
    assert.equal((await memory.danglingRefs()).length, 0, '干净库不能继承旧库的悬空引用');

    // ② 自带墓碑的 Pack：墓碑先装，指向被它判死的 id 的边也要被挡
    const selfDead = await memory.importPack(packOf([{
        id: 'workspace',
        entries: [{ id: 'm-revive', target: 'project', content: '这条是活的。', importance: 'normal', createdAt: '2026-09-01T00:00:00.000Z', updatedAt: '2026-09-01T00:00:00.000Z' }],
        ledger: [],
        links: [{ id: 'K-y', sourceId: 'm-gone', targetId: 'm-revive', kind: 'related', note: '', at: '2026-09-01T00:00:00.000Z' }],
        archiveIndex: { digests: [] },
        tombstones: [{ id: 'm-gone', at: '2026-09-01T00:00:00.000Z', reason: 'remove' }],
    }]));
    assert.equal(selfDead.links, 0);
    assert.ok(selfDead.skippedTombstoned >= 1);
    assert.equal(selfDead.tombstones, 1, 'Pack 自带的墓碑要装进来');
    assert.equal((await memory.danglingRefs()).length, 0);
});

await check('时间窗：迁移条目（只有 storedAt）在窗内可见，且「没有时间戳」与「在窗外」分开报', async () => {
    const { root } = await freshRoot();
    const paths = memoryPaths(root, {});
    await mkdir(paths.archive, { recursive: true });
    await writeFile(paths.index, JSON.stringify({
        version: 1,
        updatedAt: '2026-09-01T00:00:00.000Z',
        digests: [{
            id: 'a-2026-08', month: '2026-08', at: '2026-08-05T00:00:00.000Z', file: 'archive/2026-08.md',
            blocks: [{
                id: 'b-1', at: '2026-08-05T00:00:00.000Z', reason: 'mnemon-migration',
                items: [
                    { id: 'mnemon:dated', kind: 'mnemon', content: '只有 storedAt 的历史记忆。', storedAt: '2026-08-05T00:00:00.000Z', category: 'fact' },
                    { id: 'mnemon:noTime', kind: 'mnemon', content: '连时间戳都没有的老条目。', category: 'fact' },
                ],
            }],
        }],
    }, null, 2), 'utf8');
    const memory = createMemory({ root, memory: {} });

    const hit = await memory.read({ layer: 'archive', since: '2026-08-01', until: '2026-08-31' });
    assert.equal(hit.archive.items_total, 1, '只有 storedAt 的迁移条目必须能被时间窗读到');
    assert.equal(hit.archive.items[0].id, 'mnemon:dated');
    assert.equal(hit.archive.windowUnstamped, 1);
    assert.equal(hit.archive.windowFiltered, 0);

    // 窗口完全对不上时：要说清「没有时间戳」，**不能**建议放宽窗口（放宽永远无效）
    const miss = await memory.read({ layer: 'archive', since: '2027-01-01' });
    assert.equal(miss.archive.items_total, 0);
    const missText = renderRead(miss);
    assert.ok(missText.includes('1 条没有时间戳（带时间窗时排除，放宽窗口也读不到）'), missText);
    assert.ok(missText.includes('1 条在时间窗之外'), missText);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`memory: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
