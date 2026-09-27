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
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import {
    createMemory,
    globalMemoryDir,
    memoryExists,
    memoryPaths,
    memoryStats,
    renderDigest,
    renderMutation,
    renderRead,
} from '../src/memory.js';
import { migrateMnemon, readMnemonInsights, renderMigration } from '../src/migrate.js';
import { projectDigest, projectionStateOf, REFRESH_EVERY, resetProjection, sessionIdOf } from '../src/projection.js';
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
    assert.match(store.revision, /^[0-9a-f]{40}$/);
});

await check('revision 只由内容决定：同一份条目算出来一样，改了就变', async () => {
    const { root, memory } = await freshRoot();
    const first = await memory.mutate({ action: 'add', target: 'project', content: '约定 A' });
    const revisionA = first.revision;
    const snapshot = await memory.snapshot();
    assert.equal(snapshot.revision, revisionA);
    const second = await memory.mutate({ action: 'add', target: 'project', content: '约定 B' });
    assert.notEqual(second.revision, revisionA);
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

// ── 台账 ────────────────────────────────────────────────────────────────────

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
    assert.ok(omitted.includes(String(digest.revision).slice(0, 8)), '要带上 revision，省略才可核对');
    assert.ok(omitted.includes("office_memory({ action: 'read', layer: 'hot' })"), '要给读全文的入口');
    // 用量与台账照旧：台账不体现在 revision 里，省掉它会让 office_run 之后看不到新登记。
    assert.ok(omitted.includes('用户偏好 1 条'), '用量行要保留');
    assert.ok(omitted.includes('项目与环境 12 条'), '用量行要保留');
    assert.ok(omitted.includes('季度汇报.pptx'), '台账部分照旧每次都给');
    assert.ok(omitted.includes('主动记录'), '主动记录指引照旧');
});

// ── 投影：只在变化时贴 ──────────────────────────────────────────────────────

await check('投影：同一会话 revision 不变时只给一行，变了又贴全文', async () => {
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
    assert.ok(second.text.includes(String(digest.revision).slice(0, 8)));

    // 热记忆变了（内容改了 → revision 变）→ 又贴全文
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
    await createMemory({ root }).mutate({ action: 'add', target: 'project', content: '约定：投影不要重复贴。' });
    const tools = buildTools(resolveConfig({ memory: {} }));
    const help = tools.find((item) => item.name === 'office_help');
    // 带 id 的会话：投影状态按它记。
    const exec = { agent: { session: { id: 's-help', header: { cwd: root } } } };

    const first = await help.execute({ topic: 'word' }, exec);
    assert.ok(first.text.includes('约定：投影不要重复贴'), '第一次要带正文');
    const second = await help.execute({ topic: 'word' }, exec);
    assert.ok(!second.text.includes('约定：投影不要重复贴'), '第二次不该再贴一遍正文');
    assert.ok(second.text.includes('与上一次投影相同'), '第二次给省略形态');
    assert.ok(second.text.includes('office.word'), '文档本身照旧完整返回');

    // 换一个会话：又要贴全文（它没看过）。
    const other = await help.execute({ topic: 'word' }, { agent: { session: { id: 's-help-2', header: { cwd: root } } } });
    assert.ok(other.text.includes('约定：投影不要重复贴'));
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

    let out = await help.execute({ topic: 'files' }, exec);
    assert.ok(out.text.includes('📒 记忆（还没有内容'), '空态也要告诉模型怎么记');
    assert.ok(out.text.includes('office.files.read'), '原文档不能被记忆段落顶掉');

    await createMemory({ root, memory: config.memory }).mutate({ action: 'add', target: 'project', content: '本项目 PPT 一律 16:9。' });
    out = await help.execute({ topic: 'files' }, exec);
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
    await createMemory({ root, memory: config.memory }).mutate({ action: 'add', target: 'project', content: '既有约定。' });
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
    // 五个词只命中一个 → 相关度 0.2，低于 0.25 的下限，落进「未知档」（默认只收 2 条）。
    const strict = await memory.read({ layer: 'ledger', query: '季度 甲词 乙词 丙词 丁词' });
    assert.equal(strict.quality, 'strict-v1');
    assert.ok(strict.ledger.matched >= 3, '命中 3 条候选');
    assert.equal(strict.ledger.items.length, 2, 'strict-v1 只收「未知档」的名额（默认 2 条）');
    assert.ok(strict.ledger.qualityDropped > 0, '被丢掉多少条要报出来');

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

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`memory: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
