import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import {
    ARCHIVE_GLOBAL_MARKER,
    ARCHIVE_OWNER_FILE,
    archiveDocHash,
    canonicalWorkspacePath,
    detectArchiveCollisions,
    globalNamespaceMarkerPath,
    inspectArchiveScope,
    isPathInside,
    kbArchiveFolder,
    legacyOwnerOf,
    readGlobalNamespaceMarker,
    samePhysicalPath,
    withArchiveScope,
} from '../src/memory-archive-scope.js';

const TMP = mkdtempSync(join(tmpdir(), 'dsh-archive-scope-'));
test.after(() => rmSync(TMP, { recursive: true, force: true }));

function rootOf(name) {
    const root = join(TMP, name, 'proj');
    mkdirSync(root, { recursive: true });
    return root;
}

function archiveAt(dir) {
    const archive = join(dir, 'archive');
    mkdirSync(archive, { recursive: true });
    writeFileSync(join(archive, 'index.json'), '{"version":1,"digests":[]}\n');
    writeFileSync(join(archive, '2026-09.md'), '# 旧归档\n');
    return archive;
}

function snapshot(dir) {
    if (!existsSync(dir)) return [];
    return readdirSync(dir).sort();
}

function shasum(text) {
    return createHash('sha1').update(text).digest('hex');
}

test('真实工作区身份按物理路径算：同名不同路径不碰撞，别名归一', () => {
    const a = rootOf('a');
    const b = rootOf('b');
    const ca = canonicalWorkspacePath(a);
    const cb = canonicalWorkspacePath(b);
    const alias = canonicalWorkspacePath(join(a, '..', 'proj'));
    assert.equal(ca.ok, true);
    assert.equal(cb.ok, true);
    assert.notEqual(ca.key, cb.key, '两个 proj 必须是两个身份');
    assert.equal(ca.label, cb.label, '显示名可以相同');
    assert.notEqual(ca.display, cb.display, '显示名要带 shortKey 区分');
    assert.equal(alias.key, ca.key, '.. 写法必须归一到同一身份');
    assert.equal(alias.identity, ca.identity);
    assert.equal(canonicalWorkspacePath('').ok, false);
});

test('默认 .office/memory 在真实工作区内：旧档可归属，沿用原路径与旧 KB 路径', () => {
    const root = rootOf('owned');
    archiveAt(join(root, '.office', 'memory'));
    const scope = inspectArchiveScope({ root, store: 'workspace' });
    assert.equal(scope.mode, 'workspace-owned');
    assert.equal(scope.writeNamespaced, false);
    assert.equal(scope.writeDir, scope.legacyDir);
    assert.equal(scope.legacyAttribution, 'own');
    assert.equal(scope.legacyReadable, true);
    assert.equal(scope.canWrite, true);
    assert.equal(scope.sharedDir, false);
    assert.equal(scope.retentionSafe, true);
    assert.equal(scope.reliability, 'realpath');
    assert.equal(kbArchiveFolder(scope, '2026-09'), '归档/2026-09', 'owned 根沿用旧的归档页路径');
    assert.equal(archiveDocHash(scope, 'item-1'), 'a' + shasum('workspace:item-1').slice(0, 12), 'owned 根沿用旧的 docHash');
});

test('工作区外绝对目录：判为共用，新写命名空间，旧档不归任何库', () => {
    const a = rootOf('shared-a');
    const b = rootOf('shared-b');
    const shared = join(TMP, 'shared-memory');
    const archive = archiveAt(shared);
    const before = snapshot(archive);
    const memory = { dir: shared };
    const sa = inspectArchiveScope({ root: a, memory });
    const sb = inspectArchiveScope({ root: b, memory });
    assert.equal(sa.mode, 'shared');
    assert.equal(sb.mode, 'shared');
    assert.equal(sa.legacyAttribution, 'unattributed');
    assert.equal(sa.legacyReadable, false, '没有归属标记的旧档不能被当成本工作区的');
    assert.equal(sa.writeNamespaced, true);
    assert.equal(sa.sharedDir, true);
    assert.equal(sa.retentionSafe, false, '共用 archive 不能安全执行封顶删除');
    assert.notEqual(sa.writeDir, sb.writeDir, '两个工作区的新写目录必须不同');
    assert.equal(sa.writeDir, join(archive, 'workspaces', sa.workspaceKey));
    assert.equal(sa.legacyDir, archive, '旧档原位不动');
    assert.deepEqual(snapshot(archive), before, '判定本身不得创建/移动/删除任何文件');
});

test('知识库虚拟路径：共用场景带真实工作区身份，owned/global 保留旧路径', () => {
    const a = rootOf('kb-a');
    const b = rootOf('kb-b');
    const shared = join(TMP, 'kb-shared');
    archiveAt(shared);
    const sa = inspectArchiveScope({ root: a, memory: { dir: shared } });
    const sb = inspectArchiveScope({ root: b, memory: { dir: shared } });
    assert.notEqual(kbArchiveFolder(sa, '2026-09'), kbArchiveFolder(sb, '2026-09'));
    assert.match(kbArchiveFolder(sa, '2026-09'), /^归档\/工作区\/.+ \[[0-9a-f]{8}\]\/2026-09$/);
    const key = 'same-item-id';
    assert.notEqual(archiveDocHash(sa, key), archiveDocHash(sb, key));
    assert.notEqual(archiveDocHash(sa, key), 'a' + shasum('workspace:' + key).slice(0, 12), '共用场景不能用旧的无身份 hash');

    const globalRaw = inspectArchiveScope({ store: 'global', memoryDir: join(TMP, 'global-memory') });
    assert.equal(globalRaw.writeNamespaced, false);
    assert.equal(kbArchiveFolder(globalRaw, '2026-09'), '归档/2026-09');
    assert.equal(archiveDocHash(globalRaw, key), 'a' + shasum('global:' + key).slice(0, 12));

    const globalShared = inspectArchiveScope({
        store: 'global',
        memoryDir: join(TMP, 'global-shared'),
        otherArchiveDirs: [join(TMP, 'global-shared', 'archive')],
    });
    assert.equal(globalShared.writeNamespaced, true);
    assert.equal(kbArchiveFolder(globalShared, '2026-09'), '归档/全局/2026-09');
    assert.notEqual(archiveDocHash(globalShared, key), archiveDocHash(globalRaw, key), '命名空间全局不能复用旧的 global:key');
});

test('显式 workspace-owner.json 可以让共用目录归一个工作区；另一个工作区仍按共用处理', () => {
    const a = rootOf('claim-a');
    const b = rootOf('claim-b');
    const shared = join(TMP, 'claimed-memory');
    const archive = archiveAt(shared);
    writeFileSync(join(archive, ARCHIVE_OWNER_FILE), JSON.stringify({ version: 1, workspace: { root: a } }));
    const sa = inspectArchiveScope({ root: a, memory: { dir: shared } });
    const sb = inspectArchiveScope({ root: b, memory: { dir: shared } });
    assert.equal(sa.mode, 'claimed');
    assert.equal(sa.writeDir, archive, '有明确归属标记才可以沿用旧路径');
    assert.equal(sa.legacyAttribution, 'own');
    assert.equal(sa.legacyReadable, true);
    assert.equal(sa.sharedDir, true);
    assert.equal(sa.retentionSafe, false, '即使认领，共用目录仍不能安全判断兄弟命名空间的条目');
    assert.equal(sb.mode, 'shared');
    assert.equal(sb.legacyAttribution, 'other');
    assert.equal(sb.legacyReadable, false);
    assert.equal(sb.writeNamespaced, true);
    assert.equal(sb.retentionSafe, false);
    const owner = legacyOwnerOf(archive);
    assert.equal(owner.valid, true);
    assert.equal(owner.identity, sa.workspace.identity);
});

test('只有命名空间目录时不算旧档：不会报虚假的「未归属」', () => {
    const root = rootOf('namespace-only');
    const shared = join(TMP, 'namespace-only-memory');
    mkdirSync(join(shared, 'archive', 'workspaces', 'w-deadbeef'), { recursive: true });
    const scope = inspectArchiveScope({ root, memory: { dir: shared } });
    assert.equal(scope.legacyAttribution, 'none');
    assert.equal(scope.legacyReadable, false);
    assert.equal(scope.sharedDir, true);
});

test('归属标记坏了时按没有标记处理，不据此猜归属', () => {
    const root = rootOf('bad-owner');
    const shared = join(TMP, 'bad-owner-memory');
    const archive = archiveAt(shared);
    writeFileSync(join(archive, ARCHIVE_OWNER_FILE), '{not json');
    const scope = inspectArchiveScope({ root, memory: { dir: shared } });
    assert.equal(scope.mode, 'shared');
    assert.equal(scope.legacyAttribution, 'unattributed');
    assert.equal(scope.legacyReadable, false);
    assert.equal(legacyOwnerOf(archive).valid, false);
});

test('工作区内自定义目录仍是 owned；用 .. 逃出工作区则判为 shared', () => {
    const root = rootOf('custom');
    const inside = inspectArchiveScope({ root, memory: { dir: '.office/mem2' } });
    assert.equal(inside.mode, 'workspace-owned');
    assert.equal(inside.writeNamespaced, false);
    const outside = inspectArchiveScope({ root, memory: { dir: '../../shared-memory' } });
    assert.equal(outside.mode, 'shared');
    assert.equal(outside.writeNamespaced, true);
    assert.equal(isPathInside(root, join(root, '..', 'proj-evil')), false);
});

test('全局层默认独立；与工作区共用物理 archive 时双方都命名空间化', () => {
    const root = rootOf('global-collision');
    const memoryDir = join(root, '.office', 'memory');
    archiveAt(memoryDir);
    archiveAt(join(TMP, 'global-memory'));
    const globalOnly = inspectArchiveScope({ store: 'global', memoryDir: join(TMP, 'global-memory') });
    assert.equal(globalOnly.mode, 'global');
    assert.equal(globalOnly.writeNamespaced, false);
    assert.equal(globalOnly.legacyAttribution, 'global');
    assert.equal(globalOnly.retentionSafe, true);
    const workspace = inspectArchiveScope({ root, memoryDir, globalDir: memoryDir });
    assert.equal(workspace.mode, 'collision');
    assert.equal(workspace.legacyAttribution, 'unattributed');
    assert.equal(workspace.retentionSafe, false);
    assert.equal(workspace.writeDir, join(memoryDir, 'archive', 'workspaces', workspace.workspaceKey));
    const global = inspectArchiveScope({ store: 'global', memory: { dir: memoryDir }, memoryDir, otherArchiveDirs: [join(memoryDir, 'archive')] });
    assert.equal(global.mode, 'global-shared');
    assert.equal(global.writeDir, join(memoryDir, 'archive', 'global'));
    assert.equal(global.legacyAttribution, 'unattributed');
    assert.equal(global.retentionSafe, false);
});

test('全局指向带工作区归属标记的 archive：不当成真正的全局旧档', () => {
    const aRoot = rootOf('foreign-workspace-a');
    const aMemory = join(aRoot, '.office', 'memory');
    const aArchive = archiveAt(aMemory);
    writeFileSync(join(aArchive, ARCHIVE_OWNER_FILE), JSON.stringify({ version: 1, workspace: { root: aRoot } }), 'utf8');
    const foreign = inspectArchiveScope({ store: 'global', memoryDir: aMemory });
    assert.equal(legacyOwnerOf(aArchive).valid, true);
    assert.equal(foreign.writeNamespaced, true, '有工作区归属标记时不能覆盖旧 archive');
    assert.equal(foreign.mode, 'global-shared');
    assert.equal(foreign.legacyAttribution, 'unattributed');
    assert.equal(foreign.legacyReadable, false, '旧档属于工作区，不能归全局');
    assert.equal(foreign.retentionSafe, false);
    assert.equal(foreign.writeDir, join(aArchive, 'global'));
    const independent = join(TMP, 'independent-global-no-owner');
    archiveAt(independent);
    const raw = inspectArchiveScope({ store: 'global', memoryDir: independent });
    assert.equal(raw.writeNamespaced, false, '没有归属标记的独立全局层仍走旧路径');
    assert.equal(raw.legacyAttribution, 'global');
    assert.equal(raw.legacyReadable, true);
});

test('显式 globalDir 无归属证据：保留旧路径与读取，但跳过保留上限', () => {
    const dir = join(TMP, 'explicit-global-raw');
    archiveAt(dir);
    const explicit = inspectArchiveScope({ store: 'global', memory: { globalDir: dir }, memoryDir: dir });
    assert.equal(explicit.writeNamespaced, false);
    assert.equal(explicit.legacyAttribution, 'global');
    assert.equal(explicit.legacyReadable, true);
    assert.equal(explicit.retentionSafe, false);
    assert.equal(explicit.retentionReason, 'explicit-global-dir-unverified');
    const defaultPath = inspectArchiveScope({ store: 'global', memoryDir: dir });
    assert.equal(defaultPath.writeNamespaced, false);
    assert.equal(defaultPath.retentionSafe, true, '默认全局路径仍按原 archiveKeep 清理');
    assert.equal(defaultPath.retentionReason, '');
});

test('全局命名空间标记一旦存在就保持：配置不再碰撞也不会切回根 archive', () => {
    const shared = join(TMP, 'global-marker');
    const archive = archiveAt(shared);
    const marker = globalNamespaceMarkerPath(archive);
    mkdirSync(join(archive, 'global'), { recursive: true });
    writeFileSync(marker, JSON.stringify({ version: 1, kind: 'global-archive', legacy: 'unattributed', shared: true, reason: 'global-shared' }) + '\n');
    const scope = inspectArchiveScope({ store: 'global', memoryDir: shared });
    assert.equal(scope.writeNamespaced, true, '标记存在时必须继续走 archive/global');
    assert.equal(scope.writeDir, join(archive, 'global'));
    assert.equal(scope.mode, 'global-shared');
    assert.equal(scope.legacyAttribution, 'unattributed');
    assert.equal(scope.retentionSafe, false, '标记记录过共享，保守跳过封顶删除');
    assert.equal(readGlobalNamespaceMarker(archive).present, true);
    assert.equal(ARCHIVE_GLOBAL_MARKER, 'scope.json');
});

test('多工作区碰撞报告：shared-archive / same-label / duplicate-workspace 分清', () => {
    const a = rootOf('collision-a');
    const alias = join(a, '..', 'proj');
    const b = rootOf('collision-b');
    const shared = join(TMP, 'collision-shared');
    archiveAt(shared);
    const report = detectArchiveCollisions({
        roots: [a, b, alias],
        memory: { dir: shared },
        globalDir: join(TMP, 'collision-global'),
    });
    assert.equal(report.workspaces.length, 3);
    const kinds = report.collisions.map((item) => item.kind);
    assert.ok(kinds.includes('shared-archive'));
    assert.ok(kinds.includes('same-label'));
    assert.ok(kinds.includes('duplicate-workspace'));
    const sharedCollision = report.collisions.find((item) => item.kind === 'shared-archive');
    assert.equal(sharedCollision.workspaceKeys.length, 3);
    assert.equal(sharedCollision.legacyAttribution, 'unattributed');
    const duplicate = report.collisions.find((item) => item.kind === 'duplicate-workspace');
    assert.equal(duplicate.roots.length, 2);
});

test('未知工作区拒绝写入，旧档不读取也不冒充本库', () => {
    const shared = join(TMP, 'unknown-memory');
    archiveAt(shared);
    const scope = inspectArchiveScope({ root: '', memory: { dir: shared } });
    assert.equal(scope.mode, 'unknown');
    assert.equal(scope.canWrite, false);
    assert.equal(scope.writeNamespaced, true);
    assert.equal(scope.writeDir, join(shared, 'archive', 'workspaces', 'unknown'));
    assert.equal(scope.legacyAttribution, 'unattributed');
    assert.equal(scope.legacyReadable, false);
    assert.equal(scope.retentionSafe, false);
    assert.equal(scope.workspace.ok, false);
});

test('withArchiveScope 是纯函数：只换写路径，不改原对象、不碰盘', () => {
    const root = rootOf('with-scope');
    const shared = join(TMP, 'with-scope-memory');
    archiveAt(shared);
    const base = {
        dir: shared,
        archive: join(shared, 'archive'),
        index: join(shared, 'archive', 'index.json'),
        memory: join(shared, 'memory.json'),
    };
    const before = snapshot(join(shared, 'archive'));
    const scope = inspectArchiveScope({ root, memory: { dir: shared } });
    const next = withArchiveScope(base, scope);
    assert.equal(base.archive, join(shared, 'archive'));
    assert.equal(next.archive, scope.writeDir);
    assert.equal(next.index, join(scope.writeDir, 'index.json'));
    assert.equal(next.legacyArchive, join(shared, 'archive'));
    assert.equal(next.archiveScope, scope);
    assert.deepEqual(snapshot(join(shared, 'archive')), before);
});

test('junction / symlink 别名归一到同一工作区（环境不允许时跳过）', (t) => {
    const real = rootOf('link-real');
    const link = join(TMP, 'link-alias');
    try {
        symlinkSync(real, link, process.platform === 'win32' ? 'junction' : 'dir');
    } catch {
        t.skip('当前环境不允许创建目录链接');
        return;
    }
    const a = canonicalWorkspacePath(real);
    const b = canonicalWorkspacePath(link);
    assert.equal(a.key, b.key);
    assert.equal(samePhysicalPath(real, link), true);
});
