/**
 * 实际 Memory/KB 集成回归（不是只测路径 helper）。
 *
 * 覆盖三件事：
 *   1. 两个同名工作区共用 absolute memoryDir 时，新归档按真实工作区分开；
 *   2. 没有可核实归属的旧档保持逐字节不变，且不进入任何工作区的 KB 投影
 *      （kbRows / kbList / kbRead / kbSearch 都找不到），只在 status 里留
 *      一条不含文件名/正文/绝对路径的警告；
 *   3. 有一条显式 workspace-owner.json 时，只有认领者能读旧档并保留旧 KB
 *      路径；全局层在 scope:both / scope:global 与配置翻转后保持同一落点。
 *
 * 跑法（在已应用 archive-memory-integration.patch 的副本里）：
 *   node test/memory-archive-integration.mjs
 */
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';

import { createMemory, globalMemoryPaths } from '../src/memory.js';
import { globalNamespaceMarkerPath } from '../src/memory-archive-scope.js';

const TMP = mkdtempSync(join(tmpdir(), 'dsh-archive-integration-'));
const rootA = join(TMP, 'a', 'proj');
const rootB = join(TMP, 'b', 'proj');
const shared = join(TMP, 'shared-memory');
const kbDir = join(TMP, 'kb-single');
const globalDir = join(TMP, 'global-memory');
const alias = join(TMP, 'alias-proj');
const legacyMonth = '2026-09';
const legacyPage = 'ZZZ-LEGACY-UNIQUE-42 未知旧档正文，不得被任何工作区看到。';
const legacyGlobalPage = 'ZZZ-LEGACY-GLOBAL-42 旧全局档正文，不得被命名空间全局冒充。';
const legacyBody = join(shared, 'archive', legacyMonth + '.md');
const legacyIndex = join(shared, 'archive', 'index.json');
const legacyIndexJson = {
    version: 1,
    updatedAt: '2026-09-01T00:00:00.000Z',
    digests: [{
        id: 'a-legacy-1',
        month: legacyMonth,
        volume: 1,
        at: '2026-09-01T00:00:00.000Z',
        file: 'archive/' + legacyMonth + '.md',
        itemFile: 'archive/' + legacyMonth + '.jsonl',
        blocks: [{
            id: 'block-legacy-1',
            at: '2026-09-01T00:00:00.000Z',
            reason: 'external',
            items: [{ id: 'legacy-1', kind: 'hot', origin: 'workspace', importance: 'normal', content: legacyPage }],
        }],
    }],
};
const config = { dir: shared, kbDir, globalDir };
const pass = [];
function ok(name, condition, detail = '') {
    assert.ok(condition, name + (detail ? ' :: ' + detail : ''));
    pass.push(name);
}
function bytes(path) {
    return readFileSync(path);
}
async function readProbe(fn) {
    try {
        return await fn();
    } catch (error) {
        return { error: String(error?.message ?? error) };
    }
}

try {
    mkdirSync(rootA, { recursive: true });
    mkdirSync(rootB, { recursive: true });
    mkdirSync(join(shared, 'archive'), { recursive: true });
    writeFileSync(legacyIndex, JSON.stringify(legacyIndexJson, null, 2) + '\n', 'utf8');
    writeFileSync(legacyBody, '# 旧归档\n\n- ' + legacyPage + '\n', 'utf8');
    const beforeIndex = bytes(legacyIndex);
    const beforeBody = bytes(legacyBody);

    // junction 是 Windows 上无需管理员权限的目录别名；环境不支持时直接失败，避免静默跳过。
    symlinkSync(rootA, alias, 'junction');
    const A = createMemory({ root: rootA, memory: config });
    const B = createMemory({ root: rootB, memory: config });
    const AL = createMemory({ root: alias, memory: config });

    ok('两个同名工作区共用 absolute memoryDir：新写目录不同', A.paths.archive !== B.paths.archive, A.paths.archive + ' vs ' + B.paths.archive);
    ok('新写落在 archive/workspaces 命名空间', A.paths.archive.startsWith(join(shared, 'archive', 'workspaces')) && B.paths.archive.startsWith(join(shared, 'archive', 'workspaces')));
    ok('别名（junction）归一到同一物理工作区：归档目录相同', AL.paths.archive === A.paths.archive, AL.paths.archive + ' vs ' + A.paths.archive);
    ok('判定与创建记忆不改旧档', bytes(legacyIndex).equals(beforeIndex) && bytes(legacyBody).equals(beforeBody));

    const sameItem = { id: 'same-1', kind: 'hot', origin: 'workspace', importance: 'normal', content: '同名工作区里同 id 同内容的新归档。' };
    await A.appendArchiveBlock({ items: [sameItem], month: legacyMonth });
    await B.appendArchiveBlock({ items: [sameItem], month: legacyMonth });

    const aIndex = join(A.paths.archive, 'index.json');
    const bIndex = join(B.paths.archive, 'index.json');
    ok('A 的新档只落在 A 的命名空间', existsSync(aIndex));
    ok('B 的新档只落在 B 的命名空间', existsSync(bIndex));
    ok('旧档在新写前后逐字节不变', bytes(legacyIndex).equals(beforeIndex) && bytes(legacyBody).equals(beforeBody));

    const aBrowse = await A.browse();
    const bBrowse = await B.browse();
    ok('A 的归档读侧只含自己的新档', aBrowse.archive.filter((item) => item.id === 'same-1').length === 1);
    ok('B 的归档读侧只含自己的新档', bBrowse.archive.filter((item) => item.id === 'same-1').length === 1);
    ok('未归属旧档不进入任一本工作区的归档读侧', !aBrowse.archive.some((item) => item.origin === 'legacy') && !bBrowse.archive.some((item) => item.origin === 'legacy'));
    const aOwnIndex = JSON.parse(readFileSync(aIndex, 'utf8'));
    ok('命名空间索引里的 file 仍是相对 archive 的旧写法', aOwnIndex.digests[0].file.startsWith('archive/'), aOwnIndex.digests[0].file);
    ok('命名空间正文按 basename 落在本命名空间', existsSync(join(A.paths.archive, basename(aOwnIndex.digests[0].file))), aOwnIndex.digests[0].file);

    const rel = '普通文档.md';
    writeFileSync(join(rootA, rel), '# 普通文档\n\n知识库单库里的普通文档，保留。\n', 'utf8');
    await A.kbIngest({ path: rel, tier: 'user' });

    const aRows = await A.kbRows();
    const bRows = await B.kbRows();
    const aOwn = aRows.filter((row) => row.synthetic === true && String(row.folder).startsWith('归档/工作区/'));
    const bOwn = bRows.filter((row) => row.synthetic === true && String(row.folder).startsWith('归档/工作区/'));
    ok('A 有一个自己的归档投影', aOwn.length === 1, JSON.stringify(aRows.filter((row) => row.synthetic)));
    ok('B 有一个自己的归档投影', bOwn.length === 1);
    ok('同 id 同内容的归档虚拟路径与 docHash 不撞', aOwn[0].folder !== bOwn[0].folder && aOwn[0].hash !== bOwn[0].hash, JSON.stringify([aOwn[0].folder, bOwn[0].folder, aOwn[0].hash, bOwn[0].hash]));
    ok('未知旧档不在任一 KB 行里', !JSON.stringify([aRows, bRows]).includes('ZZZ-LEGACY-UNIQUE-42'));
    ok('知识库单库：A 入库的普通文档 B 也看得到', bRows.some((row) => row.synthetic !== true && String(row.path).includes('普通文档')));
    ok('KB 投影不泄露共享目录绝对路径', !JSON.stringify([aRows, bRows]).includes(shared), shared);

    const aScoped = await A.kbList({ folder: aOwn[0].folder });
    const bScoped = await B.kbList({ folder: aOwn[0].folder });
    ok('按 A 的归档文件夹读取只给 A', aScoped.items.length === 1 && aScoped.items[0].id === aOwn[0].id);
    ok('B 按 A 的归档文件夹读取为空（命名空间隔离）', bScoped.items.length === 0);

    const aUnattributed = await A.kbList({ folder: '归档/未归属' });
    const bUnattributed = await B.kbList({ folder: '归档/未归属' });
    ok('A 的未归属文件夹为空', aUnattributed.total === 0 && aUnattributed.items.length === 0);
    ok('B 的未归属文件夹为空', bUnattributed.total === 0 && bUnattributed.items.length === 0);
    const aSearch = await A.kbSearch({ query: 'ZZZ-LEGACY-UNIQUE-42' });
    const bSearch = await B.kbSearch({ query: 'ZZZ-LEGACY-UNIQUE-42' });
    ok('A 的 KB 检索找不到未知旧档', aSearch.matched === 0 && aSearch.hits.length === 0);
    ok('B 的 KB 检索找不到未知旧档', bSearch.matched === 0 && bSearch.hits.length === 0);
    const aRead = await readProbe(() => A.kbRead({ path: '归档/未归属/2026-09/legacy.md' }));
    const bRead = await readProbe(() => B.kbRead({ path: '归档/未归属/2026-09/legacy.md' }));
    ok('A 的 KB 按未归属路径读不到未知旧档', !JSON.stringify(aRead).includes('ZZZ-LEGACY-UNIQUE-42'));
    ok('B 的 KB 按未归属路径读不到未知旧档', !JSON.stringify(bRead).includes('ZZZ-LEGACY-UNIQUE-42'));

    const aStatus = await A.status();
    const aScopeStatus = aStatus.stores.find((store) => store.id === 'workspace')?.archiveScope;
    ok('A 的状态里有归档隔离警告', aScopeStatus?.legacy === 'unattributed' && aScopeStatus.warnings.some((text) => text.includes('无法核实归属')), JSON.stringify(aScopeStatus));
    ok('状态警告不含正文/文件名/绝对路径', !JSON.stringify(aScopeStatus).includes('ZZZ-LEGACY-UNIQUE-42') && !JSON.stringify(aScopeStatus).includes(shared), JSON.stringify(aScopeStatus));
    ok('状态警告明确保留上限已跳过', aScopeStatus?.retentionSafe === false && aScopeStatus.warnings.some((text) => text.includes('保留上限')), JSON.stringify(aScopeStatus));

    // ── 共用 archive 的封顶删除必须跳过，不能拿共享 links/tombstones 去赌 ──────
    const tightRoot = join(TMP, 'tight', 'proj');
    const tightShared = join(TMP, 'tight-shared-memory');
    mkdirSync(tightRoot, { recursive: true });
    mkdirSync(join(tightShared, 'archive'), { recursive: true });
    const Tight = createMemory({ root: tightRoot, memory: { dir: tightShared, kbDir: join(TMP, 'tight-kb'), globalDir, archiveKeep: 1 } });
    await Tight.appendArchiveBlock({ items: [{ id: 'tight-1', kind: 'hot', origin: 'workspace', importance: 'normal', content: '保留上限测试一。' }], month: '2026-01' });
    await Tight.appendArchiveBlock({ items: [{ id: 'tight-2', kind: 'hot', origin: 'workspace', importance: 'normal', content: '保留上限测试二。' }], month: '2026-02' });
    const tightIndex = JSON.parse(readFileSync(join(Tight.paths.archive, 'index.json'), 'utf8'));
    ok('shared archiveKeep=1 时两卷都保留（跳过删除）', tightIndex.digests.length === 2, JSON.stringify(tightIndex.digests.map((row) => row.month)));
    ok('跳过后两卷正文都还在', tightIndex.digests.every((row) => existsSync(join(Tight.paths.archive, basename(row.file))) && existsSync(join(Tight.paths.archive, basename(row.itemFile)))));
    const tightStatus = (await Tight.status()).stores.find((store) => store.id === 'workspace');
    ok('跳过后 status 明确报保留上限已跳过', tightStatus.archiveScope.retentionSafe === false && tightStatus.archiveScope.warnings.some((text) => text.includes('保留上限')));

    // ── 显式归属：只有认领者能读旧档，并保留旧的 KB 路径 / docHash ────────────
    writeFileSync(join(shared, 'archive', 'workspace-owner.json'), JSON.stringify({ version: 1, workspace: { root: rootA } }) + '\n', 'utf8');
    const AClaim = createMemory({ root: rootA, memory: config });
    const BClaim = createMemory({ root: rootB, memory: config });
    ok('认领者沿用旧 archive 路径', AClaim.paths.archive === join(shared, 'archive'), AClaim.paths.archive);
    ok('未认领者仍写自己的命名空间', BClaim.paths.archive.startsWith(join(shared, 'archive', 'workspaces')), BClaim.paths.archive);
    const aClaimBrowse = await AClaim.browse();
    const bClaimBrowse = await BClaim.browse();
    ok('认领者的归档读侧能看到旧档', aClaimBrowse.archive.some((item) => item.id === 'legacy-1' && String(item.text ?? item.content).includes('ZZZ-LEGACY-UNIQUE-42')));
    ok('未认领者的归档读侧看不到旧档', !bClaimBrowse.archive.some((item) => item.id === 'legacy-1') && !JSON.stringify(bClaimBrowse).includes('ZZZ-LEGACY-UNIQUE-42'));
    const aClaimRows = await AClaim.kbRows();
    const bClaimRows = await BClaim.kbRows();
    ok('认领者的 KB 旧档保留旧路径 归档/YYYY-MM', aClaimRows.some((row) => row.synthetic === true && row.folder === '归档/' + legacyMonth && String(row.text).includes('ZZZ-LEGACY-UNIQUE-42')));
    ok('未认领者的 KB 里没有旧档', !JSON.stringify(bClaimRows).includes('ZZZ-LEGACY-UNIQUE-42'));
    ok('认领读写后旧档仍逐字节不变', bytes(legacyIndex).equals(beforeIndex) && bytes(legacyBody).equals(beforeBody));

    // ── 全局层：A/B / scope:both / scope:global / 配置翻转都必须同一落点 ──────
    mkdirSync(join(globalDir, 'archive'), { recursive: true });
    writeFileSync(join(globalDir, 'archive', 'index.json'), JSON.stringify({
        version: 1,
        updatedAt: '2026-09-01T00:00:00.000Z',
        digests: [{
            id: 'a-global-legacy',
            month: legacyMonth,
            volume: 1,
            at: '2026-09-01T00:00:00.000Z',
            file: 'archive/' + legacyMonth + '.md',
            itemFile: 'archive/' + legacyMonth + '.jsonl',
            blocks: [{ id: 'block-global-legacy', at: '2026-09-01T00:00:00.000Z', reason: 'external', items: [{ id: 'global-legacy-1', kind: 'hot', origin: 'global', importance: 'normal', content: legacyGlobalPage }] }],
        }],
    }, null, 2) + '\n', 'utf8');
    writeFileSync(join(globalDir, 'archive', legacyMonth + '.md'), '# 旧全局归档\n\n- ' + legacyGlobalPage + '\n', 'utf8');

    const configBothA = { ...config, dir: globalDir, globalDir, scope: 'both' };
    const configBothB = { ...config, dir: globalDir, globalDir, scope: 'both' };
    const AGlobal = createMemory({ root: rootA, memory: { ...config, dir: globalDir, globalDir, scope: 'global' } });
    const BGlobal = createMemory({ root: rootB, memory: { ...config, dir: globalDir, globalDir, scope: 'global' } });
    const ABoth = createMemory({ root: rootA, memory: configBothA });
    const BBoth = createMemory({ root: rootB, memory: configBothB });
    const expectedGlobal = join(globalDir, 'archive', 'global');
    ok('scope:global A 的全局归档落在 archive/global', AGlobal.paths.archive === expectedGlobal, AGlobal.paths.archive);
    ok('scope:global B 的全局归档与 A 同落点', BGlobal.paths.archive === expectedGlobal, BGlobal.paths.archive);
    ok('scope:both A 的全局路径与 scope:global 一致', globalMemoryPaths(configBothA).archive === expectedGlobal);
    ok('scope:both B 的全局路径与 scope:global 一致', globalMemoryPaths(configBothB).archive === expectedGlobal);
    ok('A/B 的 scope:both 全局落点一致', globalMemoryPaths(configBothA).archive === globalMemoryPaths(configBothB).archive);

    const globalItem = { id: 'global-new-1', kind: 'hot', origin: 'global', importance: 'normal', content: '命名空间全局层的新条目。' };
    await AGlobal.appendArchiveBlock({ items: [globalItem], month: legacyMonth });
    const bGlobalBrowse = await BGlobal.browse();
    ok('B 的 scope:global 能读到 A 写的全局归档', bGlobalBrowse.archive.some((item) => item.id === 'global-new-1'));
    ok('命名空间全局不投影旧的未归属全局档', !JSON.stringify(bGlobalBrowse).includes('ZZZ-LEGACY-GLOBAL-42'));
    const aGlobalRows = await AGlobal.kbRows();
    const bGlobalRows = await BGlobal.kbRows();
    ok('命名空间全局的 KB 路径是 归档/全局/YYYY-MM', aGlobalRows.some((row) => row.synthetic === true && row.folder === '归档/全局/' + legacyMonth));
    ok('命名空间全局的 KB 不含旧全局档', !JSON.stringify([aGlobalRows, bGlobalRows]).includes('ZZZ-LEGACY-GLOBAL-42'));
    ok('scope:both 的全局 browses 也看不到旧全局档', !JSON.stringify(await ABoth.browse()).includes('ZZZ-LEGACY-GLOBAL-42') && !JSON.stringify(await BBoth.browse()).includes('ZZZ-LEGACY-GLOBAL-42'));

    // ── 默认目录碰撞：memory.dir 未配置时，global 也必须把工作区 archive 当碰撞证据 ──
    const defRoot = join(TMP, 'default', 'proj');
    const defGlobal = join(defRoot, '.office', 'memory');
    const defMonth = '2026-11';
    const defPage = 'ZZZ-DEFAULT-LEGACY-42 默认目录碰撞下的旧档正文，不得被全局层读到。';
    mkdirSync(join(defGlobal, 'archive'), { recursive: true });
    writeFileSync(join(defGlobal, 'archive', 'index.json'), JSON.stringify({
        version: 1,
        updatedAt: '2026-11-01T00:00:00.000Z',
        digests: [{
            id: 'a-default-legacy',
            month: defMonth,
            volume: 1,
            at: '2026-11-01T00:00:00.000Z',
            file: 'archive/' + defMonth + '.md',
            itemFile: 'archive/' + defMonth + '.jsonl',
            blocks: [{ id: 'block-default-legacy', at: '2026-11-01T00:00:00.000Z', reason: 'external', items: [{ id: 'default-legacy-1', kind: 'hot', origin: 'global', importance: 'normal', content: defPage }] }],
        }],
    }, null, 2) + '\n', 'utf8');
    writeFileSync(join(defGlobal, 'archive', defMonth + '.md'), '# 默认碰撞旧档\n\n- ' + defPage + '\n', 'utf8');
    const Def = createMemory({ root: defRoot, memory: { globalDir: defGlobal, kbDir: join(TMP, 'def-kb'), scope: 'global' } });
    const defMarker = globalNamespaceMarkerPath(join(defGlobal, 'archive'));
    ok('默认目录碰撞：全局归档改落 archive/global', Def.paths.archive === join(defGlobal, 'archive', 'global'), Def.paths.archive);
    ok('默认目录碰撞：createMemory 不写全局标记', !existsSync(defMarker), defMarker);
    ok('默认目录碰撞：全局读侧看不到未归属旧档', !JSON.stringify(await Def.browse()).includes('ZZZ-DEFAULT-LEGACY-42'));
    ok('默认目录碰撞：全局 KB 看不到未归属旧档', !JSON.stringify(await Def.kbRows()).includes('ZZZ-DEFAULT-LEGACY-42'));
    ok('默认目录碰撞：全局导出包看不到未归属旧档', !JSON.stringify(await Def.exportPack()).includes('ZZZ-DEFAULT-LEGACY-42'));
    const defScope = (await Def.status()).stores.find((store) => store.id === 'global')?.archiveScope;
    ok('默认目录碰撞：状态判定为 global-shared + unattributed', defScope?.mode === 'global-shared' && defScope?.legacy === 'unattributed', JSON.stringify(defScope));
    ok('默认目录碰撞：状态保留上限标记为不安全', defScope?.retentionSafe === false);
    ok('默认目录碰撞：读/状态之后仍无标记', !existsSync(defMarker));
    await Def.appendArchiveBlock({ items: [{ id: 'def-new-1', kind: 'hot', origin: 'global', importance: 'normal', content: '默认碰撞下真实写入成功才落标记。' }], month: defMonth });
    const defMarkerJson = JSON.parse(readFileSync(defMarker, 'utf8'));
    ok('默认目录碰撞：真实写成功后标记 legacy=unattributed', defMarkerJson.legacy === 'unattributed' && defMarkerJson.shared === true, JSON.stringify(defMarkerJson));
    ok('默认目录碰撞：标记出现后旧档仍不被全局读到', !JSON.stringify(await Def.browse()).includes('ZZZ-DEFAULT-LEGACY-42'));

    // ── 相对目录碰撞：memory.dir 相对路径解析后与 globalDir 重合，结论必须一致 ──
    const relRoot = join(TMP, 'relative', 'proj');
    const relDir = join(relRoot, '.office', 'mem2');
    const relMonth = '2026-12';
    const relPage = 'ZZZ-RELATIVE-LEGACY-42 相对目录碰撞下的旧档正文，不得被全局层读到。';
    mkdirSync(join(relDir, 'archive'), { recursive: true });
    writeFileSync(join(relDir, 'archive', 'index.json'), JSON.stringify({
        version: 1,
        updatedAt: '2026-12-01T00:00:00.000Z',
        digests: [{
            id: 'a-relative-legacy',
            month: relMonth,
            volume: 1,
            at: '2026-12-01T00:00:00.000Z',
            file: 'archive/' + relMonth + '.md',
            itemFile: 'archive/' + relMonth + '.jsonl',
            blocks: [{ id: 'block-relative-legacy', at: '2026-12-01T00:00:00.000Z', reason: 'external', items: [{ id: 'relative-legacy-1', kind: 'hot', origin: 'global', importance: 'normal', content: relPage }] }],
        }],
    }, null, 2) + '\n', 'utf8');
    writeFileSync(join(relDir, 'archive', relMonth + '.md'), '# 相对碰撞旧档\n\n- ' + relPage + '\n', 'utf8');
    const Rel = createMemory({ root: relRoot, memory: { dir: '.office/mem2', globalDir: relDir, kbDir: join(TMP, 'rel-kb'), scope: 'global' } });
    const relMarker = globalNamespaceMarkerPath(join(relDir, 'archive'));
    ok('相对目录碰撞：全局归档改落 archive/global', Rel.paths.archive === join(relDir, 'archive', 'global'), Rel.paths.archive);
    ok('相对目录碰撞：createMemory 不写全局标记', !existsSync(relMarker));
    ok('相对目录碰撞：全局读/KB/导出都不泄露旧档', !JSON.stringify([await Rel.browse(), await Rel.kbRows(), await Rel.exportPack()]).includes('ZZZ-RELATIVE-LEGACY-42'));

    // ── 独立真实全局：不与工作区碰撞时仍按旧路径读取，且不写标记 ─────────────
    const indepRoot = join(TMP, 'independent', 'proj');
    const indepGlobal = join(TMP, 'independent-global');
    const indepMonth = '2026-08';
    const indepPage = 'ZZZ-INDEPENDENT-GLOBAL-42 独立全局层的正常旧档。';
    mkdirSync(indepRoot, { recursive: true });
    mkdirSync(join(indepGlobal, 'archive'), { recursive: true });
    writeFileSync(join(indepGlobal, 'archive', 'index.json'), JSON.stringify({
        version: 1,
        updatedAt: '2026-08-01T00:00:00.000Z',
        digests: [{
            id: 'a-independent-global',
            month: indepMonth,
            volume: 1,
            at: '2026-08-01T00:00:00.000Z',
            file: 'archive/' + indepMonth + '.md',
            itemFile: 'archive/' + indepMonth + '.jsonl',
            blocks: [{ id: 'block-independent-global', at: '2026-08-01T00:00:00.000Z', reason: 'external', items: [{ id: 'independent-global-1', kind: 'hot', origin: 'global', importance: 'normal', content: indepPage }] }],
        }],
    }, null, 2) + '\n', 'utf8');
    writeFileSync(join(indepGlobal, 'archive', indepMonth + '.md'), '# 独立全局旧档\n\n- ' + indepPage + '\n', 'utf8');
    const Indep = createMemory({ root: indepRoot, memory: { globalDir: indepGlobal, kbDir: join(TMP, 'indep-kb'), scope: 'global' } });
    ok('独立真实全局：仍使用旧的 archive 路径', Indep.paths.archive === join(indepGlobal, 'archive'), Indep.paths.archive);
    ok('独立真实全局：旧档可正常读取', JSON.stringify(await Indep.browse()).includes('ZZZ-INDEPENDENT-GLOBAL-42'));
    ok('独立真实全局：不写命名空间标记', !existsSync(globalNamespaceMarkerPath(join(indepGlobal, 'archive'))));

    // ── 外部 globalDir 指向带工作区归属标记的 archive：不能冒充真正的全局旧档 ──
    const foreignRoot = join(TMP, 'foreign-a', 'proj');
    const foreignMemory = join(foreignRoot, '.office', 'memory');
    const foreignArchive = join(foreignMemory, 'archive');
    const foreignMonth = '2026-07';
    const foreignPage = 'ZZZ-FOREIGN-OWNER-42 别的默认工作区旧档，不能被外部 globalDir 当成全局旧档。';
    mkdirSync(foreignArchive, { recursive: true });
    writeFileSync(join(foreignArchive, 'index.json'), JSON.stringify({
        version: 1,
        updatedAt: '2026-07-01T00:00:00.000Z',
        digests: [{
            id: 'a-foreign-owner',
            month: foreignMonth,
            volume: 1,
            at: '2026-07-01T00:00:00.000Z',
            file: 'archive/' + foreignMonth + '.md',
            itemFile: 'archive/' + foreignMonth + '.jsonl',
            blocks: [{ id: 'block-foreign-owner', at: '2026-07-01T00:00:00.000Z', reason: 'external', items: [{ id: 'foreign-owner-1', kind: 'hot', origin: 'workspace', importance: 'normal', content: foreignPage }] }],
        }],
    }, null, 2), 'utf8');
    writeFileSync(join(foreignArchive, foreignMonth + '.md'), '# 外部工作区旧档  - ' + foreignPage, 'utf8');
    writeFileSync(join(foreignArchive, 'workspace-owner.json'), JSON.stringify({ version: 1, workspace: { root: foreignRoot } }), 'utf8');
    const foreignBRoot = join(TMP, 'foreign-b', 'proj');
    mkdirSync(foreignBRoot, { recursive: true });
    const BForeign = createMemory({ root: foreignBRoot, memory: { globalDir: foreignMemory, kbDir: join(TMP, 'foreign-b-kb'), scope: 'global' } });
    ok('外部 globalDir 有工作区归属标记：全局归档改落 archive/global', BForeign.paths.archive === join(foreignArchive, 'global'), BForeign.paths.archive);
    ok('外部 globalDir 有工作区归属标记：createMemory 不写全局标记', !existsSync(globalNamespaceMarkerPath(foreignArchive)));
    ok('外部 globalDir 有工作区归属标记：全局读/KB/导出都不泄露旧档', !JSON.stringify([await BForeign.browse(), await BForeign.kbRows(), await BForeign.exportPack()]).includes('ZZZ-FOREIGN-OWNER-42'));
    const foreignScope = (await BForeign.status()).stores.find((store) => store.id === 'global')?.archiveScope;
    ok('外部 globalDir 有工作区归属标记：判为 global-shared + unattributed', foreignScope?.mode === 'global-shared' && foreignScope?.legacy === 'unattributed' && foreignScope?.retentionSafe === false, JSON.stringify(foreignScope));
    const AForeign = createMemory({ root: foreignRoot, memory: { globalDir: join(TMP, 'foreign-a-global'), kbDir: join(TMP, 'foreign-a-kb'), scope: 'workspace' } });
    ok('工作区自己仍按归属标记读旧档', AForeign.paths.archive === foreignArchive && JSON.stringify(await AForeign.browse()).includes('ZZZ-FOREIGN-OWNER-42'), AForeign.paths.archive);

    // ── 显式 globalDir 指向未标记的工作区 memory：保留旧路径/读取，但跳过删除 ──
    const ambRoot = join(TMP, 'ambiguous-global-a', 'proj');
    const ambMemory = join(ambRoot, '.office', 'memory');
    const ambArchive = join(ambMemory, 'archive');
    const ambOldMonth = '2026-06';
    const ambOldPage = 'ZZZ-AMBIGUOUS-OLD-42 未标记工作区的旧卷，不能被外部 globalDir 的保留上限删掉。';
    mkdirSync(ambArchive, { recursive: true });
    writeFileSync(join(ambArchive, 'index.json'), JSON.stringify({
        version: 1,
        updatedAt: '2026-06-01T00:00:00.000Z',
        digests: [{
            id: 'a-ambiguous-old',
            month: ambOldMonth,
            volume: 1,
            at: '2026-06-01T00:00:00.000Z',
            file: 'archive/' + ambOldMonth + '.md',
            itemFile: 'archive/' + ambOldMonth + '.jsonl',
            blocks: [{ id: 'block-ambiguous-old', at: '2026-06-01T00:00:00.000Z', reason: 'external', items: [{ id: 'ambiguous-old-1', kind: 'hot', origin: 'workspace', importance: 'normal', content: ambOldPage }] }],
        }],
    }, null, 2), 'utf8');
    writeFileSync(join(ambArchive, ambOldMonth + '.md'), '# 未标记工作区旧卷  - ' + ambOldPage, 'utf8');
    const ambBRoot = join(TMP, 'ambiguous-global-b', 'proj');
    mkdirSync(ambBRoot, { recursive: true });
    const Amb = createMemory({ root: ambBRoot, memory: { globalDir: ambMemory, kbDir: join(TMP, 'ambiguous-b-kb'), scope: 'global', archiveKeep: 1 } });
    const ambMarker = globalNamespaceMarkerPath(ambArchive);
    ok('显式 globalDir 未标记碰撞：沿用旧 archive 路径', Amb.paths.archive === ambArchive, Amb.paths.archive);
    ok('显式 globalDir 未标记碰撞：旧档仍可读', JSON.stringify(await Amb.browse()).includes('ZZZ-AMBIGUOUS-OLD-42'));
    ok('显式 globalDir 未标记碰撞：读/创建不写标记', !existsSync(ambMarker));
    await Amb.appendArchiveBlock({ items: [{ id: 'amb-new-1', kind: 'hot', origin: 'global', importance: 'normal', content: '显式全局一。' }], month: '2027-04' });
    await Amb.appendArchiveBlock({ items: [{ id: 'amb-new-2', kind: 'hot', origin: 'global', importance: 'normal', content: '显式全局二。' }], month: '2027-05' });
    const ambIndex = JSON.parse(readFileSync(join(ambArchive, 'index.json'), 'utf8'));
    ok('显式 globalDir 未标记碰撞：archiveKeep=1 不删旧卷', ambIndex.digests.length === 3 && ambIndex.digests.some((row) => row.month === ambOldMonth), JSON.stringify(ambIndex.digests.map((row) => row.month)));
    ok('显式 globalDir 未标记碰撞：旧卷正文与读取都还在', existsSync(join(ambArchive, ambOldMonth + '.md')) && JSON.stringify(await Amb.browse()).includes('ZZZ-AMBIGUOUS-OLD-42'));
    const ambScope = (await Amb.status()).stores.find((store) => store.id === 'global')?.archiveScope;
    ok('显式 globalDir 未标记碰撞：status 跳过保留上限并说明原因', ambScope?.retentionSafe === false && ambScope?.warnings.some((text) => text.includes('显式配置') && text.includes('保留上限')), JSON.stringify(ambScope));
    ok('显式 globalDir 未标记碰撞：写入后仍不写错误归属标记', !existsSync(ambMarker));

    // ── 默认专用全局路径：没有显式 globalDir，保留上限仍正常清理 ─────────────
    const dedicatedHome = join(TMP, 'dedicated-dsh-home');
    const previousDshHome = process.env.DSH_HOME;
    process.env.DSH_HOME = dedicatedHome;
    try {
        const dedRoot = join(TMP, 'dedicated', 'proj');
        mkdirSync(dedRoot, { recursive: true });
        const Ded = createMemory({ root: dedRoot, memory: { kbDir: join(TMP, 'dedicated-kb'), scope: 'global', archiveKeep: 1 } });
        const dedArchive = join(dedicatedHome, '.office', 'memory', 'archive');
        ok('默认专用全局：仍用默认全局 archive 路径', Ded.paths.archive === dedArchive, Ded.paths.archive);
        await Ded.appendArchiveBlock({ items: [{ id: 'ded-1', kind: 'hot', origin: 'global', importance: 'normal', content: '默认全局一。' }], month: '2027-06' });
        await Ded.appendArchiveBlock({ items: [{ id: 'ded-2', kind: 'hot', origin: 'global', importance: 'normal', content: '默认全局二。' }], month: '2027-07' });
        const dedIndex = JSON.parse(readFileSync(join(dedArchive, 'index.json'), 'utf8'));
        ok('默认专用全局：archiveKeep=1 仍会清理旧卷', dedIndex.digests.length === 1 && dedIndex.digests[0].month === '2027-07', JSON.stringify(dedIndex.digests.map((row) => row.month)));
        ok('默认专用全局：被淘汰的旧卷正文已删除', !existsSync(join(dedArchive, '2027-06.md')));
        const dedScope = (await Ded.status()).stores.find((store) => store.id === 'global')?.archiveScope;
        ok('默认专用全局：status 保留上限仍为安全', dedScope?.retentionSafe === true && dedScope?.warnings.length === 0, JSON.stringify(dedScope));
    } finally {
        if (previousDshHome === undefined) delete process.env.DSH_HOME;
        else process.env.DSH_HOME = previousDshHome;
    }

    // ── 标记只由成功的真实归档写入建立；空写/失败写都不建立 ─────────────────
    const writeGuardRoot = join(TMP, 'guard', 'proj');
    const writeGuardGlobal = join(writeGuardRoot, '.office', 'memory');
    mkdirSync(writeGuardRoot, { recursive: true });
    const Guard = createMemory({ root: writeGuardRoot, memory: { globalDir: writeGuardGlobal, kbDir: join(TMP, 'guard-kb'), scope: 'global' } });
    const guardMarker = globalNamespaceMarkerPath(join(writeGuardGlobal, 'archive'));
    ok('碰撞但尚未写归档：没有标记', !existsSync(guardMarker));
    const noop = await Guard.appendArchiveBlock({ items: [], month: '2027-01' });
    ok('空写不建立标记', noop.added === 0 && !existsSync(guardMarker));
    await Guard.appendArchiveBlock({ items: [{ id: 'guard-1', kind: 'hot', origin: 'global', importance: 'normal', content: '真实写成功才落标记。' }], month: '2027-01' });
    ok('真实归档写成功后标记出现', existsSync(guardMarker), guardMarker);
    const guardScope = JSON.parse(readFileSync(guardMarker, 'utf8'));
    ok('标记不把未知历史洗成 global', guardScope.version === 1 && guardScope.kind === 'global-archive' && guardScope.legacy !== 'global' && guardScope.shared === true, JSON.stringify(guardScope));

    // ── 碰撞 + archiveKeep：命名空间全局也必须跳过删除，并明确报警 ─────────
    const tightDefRoot = join(TMP, 'tight-default', 'proj');
    const tightDefGlobal = join(tightDefRoot, '.office', 'memory');
    mkdirSync(tightDefRoot, { recursive: true });
    const TightDef = createMemory({ root: tightDefRoot, memory: { globalDir: tightDefGlobal, kbDir: join(TMP, 'tight-def-kb'), scope: 'global', archiveKeep: 1 } });
    await TightDef.appendArchiveBlock({ items: [{ id: 'tight-def-1', kind: 'hot', origin: 'global', importance: 'normal', content: '碰撞保留上限一。' }], month: '2027-02' });
    await TightDef.appendArchiveBlock({ items: [{ id: 'tight-def-2', kind: 'hot', origin: 'global', importance: 'normal', content: '碰撞保留上限二。' }], month: '2027-03' });
    const tightDefIndex = JSON.parse(readFileSync(join(TightDef.paths.archive, 'index.json'), 'utf8'));
    ok('默认碰撞 + archiveKeep=1：两卷都保留', tightDefIndex.digests.length === 2, JSON.stringify(tightDefIndex.digests.map((row) => row.month)));
    const tightDefScope = (await TightDef.status()).stores.find((store) => store.id === 'global')?.archiveScope;
    ok('默认碰撞 + archiveKeep=1：status 报保留上限已跳过', tightDefScope?.retentionSafe === false && tightDefScope.warnings.some((text) => text.includes('保留上限')), JSON.stringify(tightDefScope));



    // 标记落盘后，配置翻转（memory.dir 改回默认/另一个目录）仍沿用 archive/global。
    const flipped = globalMemoryPaths({ globalDir, kbDir });
    ok('标记落盘后配置翻转仍用 archive/global', flipped.archive === expectedGlobal, flipped.archive);
    ok('配置翻转后 scope:both 仍与 scope:global 同落点', globalMemoryPaths({ ...config, dir: join(TMP, 'other'), globalDir, scope: 'both' }).archive === expectedGlobal);

    console.log('PASS ' + pass.length + ' integration assertions');
    console.log('A archive folder: ' + aOwn[0].folder);
    console.log('B archive folder: ' + bOwn[0].folder);
} finally {
    rmSync(TMP, { recursive: true, force: true });
}
