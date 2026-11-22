/**
 * 空虚拟文件夹 × 会话范围：memory.kbTree 合并与 /kb-scope 回归。
 *
 * 本轮要求（第二二轮补充）：
 *   - 显式创建的空文件夹必须出现在 Memory.kbTree 里（index.js 的 folderExists 就靠它），
 *     否则 /kb-scope set 会把刚建好的空目录当成「不存在」拒绝；
 *   - 有会话范围时**不**合并显式目录：范围外的目录名同样是泄漏；工具面的默认
 *     recursive=true 语义不变；
 *   - 只写 <kbDir>/folders.json，不建真实目录、不动 manifest。
 *
 * 跑法：node test/kb-folders-scope.mjs
 */
import { mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { createMemory } from '../src/memory.js';
import { createKbFolder, readKbFolders } from '../src/kb-folders.js';
import { createScopeStore, registerKbScopeCommand, SCOPE_DIR } from '../src/kb-session-scope.js';

const TMP = fileURLToPath(new URL('../../../.office/runtime/tmp/kb-folders-scope/', import.meta.url));
process.env.DSH_HOME = join(TMP, 'home');

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

let caseId = 0;
async function freshRoot() {
    caseId += 1;
    const root = join(TMP, 'case-' + caseId);
    await rm(root, { recursive: true, force: true });
    await mkdir(join(root, 'docs'), { recursive: true });
    const home = join(root, 'home');
    await mkdir(home, { recursive: true });
    const kbDir = join(root, '.office', 'kb');
    const config = { memory: { kbDir } };
    return {
        root,
        home,
        kbDir,
        config,
        scopeStore: createScopeStore({ root: join(home, SCOPE_DIR) }),
        memory: createMemory({ root, memory: config.memory }),
    };
}

function findNode(node, wanted) {
    if (node === null || typeof node !== 'object') return null;
    if (String(node.path ?? '') === wanted) return node;
    for (const child of Array.isArray(node.children) ? node.children : []) {
        const found = findNode(child, wanted);
        if (found !== null) return found;
    }
    return null;
}

function treePaths(node) {
    const out = [];
    const walk = (item) => {
        if (item === null || typeof item !== 'object') return;
        out.push(String(item.path ?? ''));
        for (const child of Array.isArray(item.children) ? item.children : []) walk(child);
    };
    walk(node);
    return out;
}

/** 与 src/index.js 的 folderExists 同口径：在完整树里按路径找节点。 */
function folderExistsFrom(memory) {
    return async (_session, folder) => {
        const tree = await memory.kbTree({});
        return findNode(tree.root ?? tree, folder) !== null;
    };
}

function scopeCtx(home, onRegister) {
    const profileContext = { name: 'test', dir: home, home, cwd: home };
    const commands = { register: (definition) => { onRegister(definition); return () => {}; } };
    return { get: (name) => (name === 'profileContext' ? profileContext : name === 'commands' ? commands : null) };
}

function execOf(id, root) {
    return { agent: { session: { id, header: { id, cwd: root } } } };
}

await check('空虚拟目录进入 Memory.kbTree；有会话范围时不合并（不泄漏范围外目录）', async () => {
    const { root, config, memory } = await freshRoot();
    const doc = join(root, 'docs', '合同.md');
    await writeFile(doc, '# 合同\n\n付款条款见附件。\n', 'utf8');
    await memory.kbIngest({ path: doc, folder: '资料/合同', displayPath: 'docs/合同.md' });
    await createKbFolder(config.memory.kbDir, '项目/空');

    const tree = await memory.kbTree({});
    const node = findNode(tree.root, '项目/空');
    assert.ok(node !== null, '未设置范围时 kbTree 要含显式创建的空目录');
    assert.equal(node.virtual, true);
    assert.equal(node.docs, 0);
    assert.equal(node.chunks, 0);

    // 工具面默认 recursive=true 的既有语义不变：资料 子树里那篇文档仍要算进来。
    const list = await memory.kbList({ folder: '资料' });
    assert.equal(list.total, 1);
    assert.equal(list.recursive, true);
    // 空目录本身没有任何文档。
    assert.equal((await memory.kbList({ folder: '项目/空' })).total, 0);

    // 会话范围（scope != null）：树只由可见行现算，不并显式目录。
    const scopedTree = await memory.kbTree({ scope: { mode: 'folder', folder: '资料' } });
    const paths = treePaths(scopedTree.root);
    assert.ok(!paths.includes('项目/空'), '范围外的显式目录不能出现在范围树里');
    assert.ok(paths.includes('资料/合同'), '范围内的目录照常出现');
});

await check('/kb-scope set 接受刚创建的空目录；不存在的仍然拒绝', async () => {
    const { root, home, config, scopeStore, memory } = await freshRoot();
    const doc = join(root, 'docs', 'a.md');
    await writeFile(doc, '# a\n\n正文。\n', 'utf8');
    await memory.kbIngest({ path: doc, folder: '资料', displayPath: 'docs/a.md' });
    await createKbFolder(config.memory.kbDir, '项目/空');

    let definition = null;
    registerKbScopeCommand(scopeCtx(home, (def) => { definition = def; }), {
        folderExists: folderExistsFrom(memory),
    });
    assert.ok(definition !== null, '/kb-scope 命令要能注册');
    const call = (rawInput) => definition.handler({
        rawInput,
        agent: { session: { id: 's1', header: { id: 's1', cwd: root } } },
    });

    const accepted = await call('set 项目/空');
    assert.equal(accepted.kind, 'success', accepted.text);
    const stored = await scopeStore.read('s1');
    assert.equal(stored.status, 'ok');
    assert.equal(stored.record.mode, 'folder');
    assert.equal(stored.record.folder, '项目/空');

    const rejected = await call('set 项目/不存在');
    assert.equal(rejected.kind, 'error');
    assert.match(rejected.text, /没有「项目\/不存在」这个虚拟文件夹/);
    assert.equal((await scopeStore.read('s1')).record.folder, '项目/空', '拒绝时不能改掉原范围');

    // 选一个空目录后，检索范围合法但结果为空（不是报错）。
    const scoped = await memory.kbList({ scope: { mode: 'folder', folder: '项目/空' } });
    assert.equal(scoped.total, 0);
});

await check('只落 folders.json：同路径幂等、不建真实目录、不改 manifest', async () => {
    const { root, kbDir, config, memory } = await freshRoot();
    const doc = join(root, 'docs', 'b.md');
    await writeFile(doc, '# b\n\n正文。\n', 'utf8');
    await memory.kbIngest({ path: doc, folder: '资料', displayPath: 'docs/b.md' });

    const manifestBefore = await readFile(join(kbDir, 'manifest.jsonl'), 'utf8');
    const first = await createKbFolder(config.memory.kbDir, '项目/空');
    assert.equal(first.created, true);
    const second = await createKbFolder(config.memory.kbDir, '项目/空');
    assert.equal(second.created, false, '同路径重复创建要幂等');
    assert.equal(second.total, 1);

    const stored = await readKbFolders(config.memory.kbDir);
    assert.deepEqual(stored, ['项目/空']);
    const manifestAfter = await readFile(join(kbDir, 'manifest.jsonl'), 'utf8');
    assert.equal(manifestAfter, manifestBefore, '虚拟目录不能写 manifest');
    let realDir = true;
    try {
        await stat(join(kbDir, '项目'));
    } catch {
        realDir = false;
    }
    assert.equal(realDir, false, '虚拟目录不能落到真实文件系统');
});

for (const result of results) {
    if (result.ok) console.log('PASS ' + result.name);
    else console.log('FAIL ' + result.name + '\n' + (result.error && result.error.stack ? result.error.stack : String(result.error)));
}
const failed = results.filter((item) => !item.ok).length;
console.log('kb-folders-scope: ' + (results.length - failed) + '/' + results.length + ' 通过');
if (failed > 0) process.exitCode = 1;
