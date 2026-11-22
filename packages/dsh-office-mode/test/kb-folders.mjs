/**
 * 显式空文件夹（folders.json）的独立单测。
 *
 * 这些用例不碰面板、不碰 manifest：只钉住「路径校验 / 幂等落盘 / 与文档推导树合并」
 * 三件事。跑法：node test/kb-folders.mjs
 *
 * 放在 test/ 顶层而不是塞进 view.mjs：这一段是纯函数 + 一个 JSON 文件，
 * 不该为了测它去起 HTTP 路由或造知识库索引。
 */

import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { promisify } from 'node:util';
import {
    KB_FOLDER_MAX_COUNT,
    KB_FOLDER_MAX_DEPTH,
    KB_FOLDER_SEGMENT_MAX,
    KB_FOLDERS_FILE,
    createKbFolder,
    mergeKbFolders,
    readKbFolders,
    validateKbFolderPath,
} from '../src/kb-folders.js';

const here = dirname(fileURLToPath(import.meta.url));
const TMP = join(here, '..', '.office', 'runtime', 'tmp', 'kb-folders');
let caseId = 0;
const results = [];

function check(name, fn) {
    return Promise.resolve().then(fn).then(() => {
        results.push({ name, ok: true });
    }, (error) => {
        results.push({ name, ok: false, error });
    });
}

async function freshDir() {
    caseId += 1;
    const dir = join(TMP, `case-${caseId}`, 'kb');
    await rm(join(TMP, `case-${caseId}`), { recursive: true, force: true });
    return dir;
}

await check('校验：合法的相对路径按段收成 / 分隔；反斜杠与首尾空格被规范掉', () => {
    assert.equal(validateKbFolderPath('参考资料'), '参考资料');
    assert.equal(validateKbFolderPath('  项目/甲  '), '项目/甲');
    assert.equal(validateKbFolderPath('项目\\甲'), '项目/甲');
    assert.equal(validateKbFolderPath('a/b/c'), 'a/b/c');
});

await check('校验：空名、绝对路径、盘符、空段、. / ..、控制字符、超长都被拒绝', () => {
    const rejects = [
        ['', /不能为空/],
        ['   ', /不能为空/],
        ['/a', /相对路径/],
        ['a/', /空目录名/],
        ['/a/b', /相对路径/],
        ['a//b', /空目录名/],
        ['C:/x', /盘符/],
        ['c:\\x', /盘符/],
        ['a/../b', /不能有 \. 或 \.\./],
        ['a/.', /不能有 \. 或 \.\./],
        ['../x', /不能有 \. 或 \.\./],
        ['a /b', /首尾不能有空格/],
        ['a/b /c', /首尾不能有空格/],
        ['a:b', /盘符/],
        ['ab:c', /不能包含/],
        ['a|b', /不能包含/],
        ['a\u0001b', /不能包含/],
        ['a.', /不要以点结尾/],
        ['x'.repeat(KB_FOLDER_SEGMENT_MAX + 1), new RegExp(`不能超过 ${KB_FOLDER_SEGMENT_MAX}`)],
        [Array.from({ length: KB_FOLDER_MAX_DEPTH + 1 }, (_, i) => `d${i}`).join('/'), /最多/],
    ];
    for (const [value, pattern] of rejects) {
        assert.throws(() => validateKbFolderPath(value), pattern, `应拒绝：${JSON.stringify(value)}`);
    }
});

await check('创建：空目录进 folders.json，不在真实磁盘 mkdir，也不碰 manifest', async () => {
    const dir = await freshDir();
    const before = existsSync(dir) ? await readdir(dir) : [];
    const created = await createKbFolder(dir, '参考资料/甲');
    assert.equal(created.action, 'kb-folder-create');
    assert.equal(created.path, '参考资料/甲');
    assert.equal(created.created, true);
    assert.equal(created.total, 1);
    assert.equal(existsSync(join(dir, '参考资料')), false, '不能真的建目录');
    assert.equal(existsSync(join(dir, '参考资料', '甲')), false, '不能真的建目录');
    assert.equal(existsSync(join(dir, 'manifest.jsonl')), false, '不能写 manifest');
    const entries = await readdir(dir);
    assert.deepEqual(entries, [KB_FOLDERS_FILE], `库目录里只应多出清单文件：${entries.join(',')}`);
    assert.deepEqual(await readKbFolders(dir), ['参考资料/甲']);
    const raw = JSON.parse(await readFile(join(dir, KB_FOLDERS_FILE), 'utf8'));
    assert.equal(raw.version, 1);
    assert.deepEqual(raw.folders, ['参考资料/甲']);
    assert.ok(before.every((name) => entries.includes(name) || name === KB_FOLDERS_FILE));
});

await check('创建：同路径再次创建幂等（created:false、文件字节不变）', async () => {
    const dir = await freshDir();
    await createKbFolder(dir, '参考资料');
    const first = await readFile(join(dir, KB_FOLDERS_FILE), 'utf8');
    const second = await createKbFolder(dir, '参考资料');
    assert.equal(second.created, false);
    assert.equal(second.total, 1);
    assert.equal(await readFile(join(dir, KB_FOLDERS_FILE), 'utf8'), first, '幂等时不应重写文件');
    const third = await createKbFolder(dir, '  参考资料  ');
    assert.equal(third.path, '参考资料', '校验会 trim，仍是同一条');
    assert.equal(third.created, false);
});

await check('创建：多条排序、上限、失败不留临时文件', async () => {
    const dir = await freshDir();
    await createKbFolder(dir, '乙');
    await createKbFolder(dir, '甲/二');
    const folders = await readKbFolders(dir);
    assert.deepEqual(folders, [...folders].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN', { numeric: true })));
    assert.deepEqual(new Set(folders), new Set(['乙', '甲/二']));
    const entries = await readdir(dir);
    assert.equal(entries.some((name) => name.includes('.tmp')), false, '不能留下临时文件');
});

await check('损坏的 folders.json：展示退回空树，但创建拒绝覆盖、原文件字节不变', async () => {
    const dir = await freshDir();
    await mkdir(dir, { recursive: true });
    const file = join(dir, KB_FOLDERS_FILE);
    for (const broken of ['这不是 JSON', '{"version":1,"folders":"不是数组"}', '{"version":1,"folders":["好",42]}']) {
        await writeFile(file, broken, 'utf8');
        const before = await readFile(file);
        await assert.rejects(createKbFolder(dir, '新目录'), (error) => error.code === 'KB_FOLDERS_CORRUPT', '坏清单必须拒绝写入：' + broken);
        assert.deepEqual(await readFile(file), before, '拒绝写入时原文件一个字节都不能变');
        assert.equal((await readdir(dir)).some((name) => name.includes('.tmp')), false, '失败不能留下临时文件');
        await assert.rejects(readKbFolders(dir, { strict: true }), (error) => error.code === 'KB_FOLDERS_CORRUPT');
    }
});

await check('并发：同一进程同时建不同目录，一条都不能丢；同路径只写一次', async () => {
    const dir = await freshDir();
    const names = ['甲', '乙', '丙', '丁', '戊'];
    const settled = await Promise.all(names.map((name) => createKbFolder(dir, name)));
    assert.equal(settled.filter((item) => item.created === true).length, names.length);
    assert.deepEqual(new Set(await readKbFolders(dir)), new Set(names));
    const same = await Promise.all(Array.from({ length: 8 }, () => createKbFolder(dir, '已存在')));
    assert.equal(same.filter((item) => item.created === true).length, 1, '同路径并发只应有一个 created:true');
    assert.equal((await readKbFolders(dir)).filter((path) => path === '已存在').length, 1);
    const entries = await readdir(dir);
    assert.equal(entries.some((name) => name.includes('.tmp')), false);
});

await check('并发：多个进程同时建不同目录，跨进程锁下一条都不能丢', async () => {
    const dir = await freshDir();
    const worker = join(TMP, 'worker-' + caseId + '.mjs');
    const moduleUrl = pathToFileURL(join(here, '..', 'src', 'kb-folders.js')).href;
    await writeFile(worker, 'import { createKbFolder } from ' + JSON.stringify(moduleUrl) + ';\n'
        + 'const [dir, path] = process.argv.slice(2);\n'
        + 'await createKbFolder(dir, path);\n', 'utf8');
    const run = promisify(execFile);
    const paths = ['进程一', '进程二', '进程三', '进程四'];
    await Promise.all(paths.map((path) => run(process.execPath, [worker, dir, path], { windowsHide: true })));
    assert.deepEqual(new Set(await readKbFolders(dir)), new Set(paths));
    const entries = await readdir(dir);
    assert.equal(entries.some((name) => name.includes('.tmp')), false, '临时文件必须清干净');
});

await check('读取：坏 JSON 退回空数组，单条坏记录被跳过，重复去重', async () => {
    const dir = await freshDir();
    await mkdir(dir, { recursive: true });
    await writeFile(join(dir, KB_FOLDERS_FILE), '这不是 JSON', 'utf8');
    assert.deepEqual(await readKbFolders(dir), []);
    await writeFile(join(dir, KB_FOLDERS_FILE), JSON.stringify({
        version: 1,
        folders: ['好', '../坏', 42, '', '好', ' 好 ', 'a/b'],
    }), 'utf8');
    assert.deepEqual(await readKbFolders(dir), ['a/b', '好'].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN', { numeric: true })));
    await writeFile(join(dir, KB_FOLDERS_FILE), JSON.stringify(['旧数组格式', 'a']), 'utf8');
    assert.deepEqual(await readKbFolders(dir), ['a', '旧数组格式'].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN', { numeric: true })));
    await rm(dir, { recursive: true, force: true });
    assert.deepEqual(await readKbFolders(dir), [], '目录不存在时是空清单而不是抛错');
});

await check('合并：只补空节点，文档推导出的计数原样保留', () => {
    const root = {
        path: '', name: '全部文档', docs: 2, chunks: 3, bytes: 10,
        children: [{
            path: '参考', name: '参考', docs: 2, chunks: 3, bytes: 10,
            children: [{ path: '参考/己', name: '己', docs: 1, chunks: 2, bytes: 5, children: [] }],
        }],
    };
    const merged = mergeKbFolders(root, ['参考/新建', '空/更深', '参考']);
    assert.equal(merged.docs, 2, '根计数不能被显式目录改动');
    const ref = merged.children.find((node) => node.path === '参考');
    assert.equal(ref.docs, 2);
    assert.equal(ref.chunks, 3);
    assert.equal(ref.bytes, 10);
    const existing = ref.children.find((node) => node.path === '参考/己');
    assert.equal(existing.docs, 1, '已有子节点计数不变');
    const created = ref.children.find((node) => node.path === '参考/新建');
    assert.ok(created, '缺的子节点要补出来');
    assert.equal(created.docs, 0);
    assert.equal(created.chunks, 0);
    assert.equal(created.bytes, 0);
    assert.equal(created.virtual, true);
    const empty = merged.children.find((node) => node.path === '空');
    assert.ok(empty && empty.virtual === true);
    const deeper = empty.children.find((node) => node.path === '空/更深');
    assert.ok(deeper, '多层路径的中间层与末层都要补');
    assert.equal(deeper.docs, 0);
    // 子节点按名字排序，面包屑与左树顺序一致。
    const names = merged.children.map((node) => node.name);
    assert.deepEqual(names, [...names].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN', { numeric: true })));
});

await check('合并：重复 / 非法路径不制造重复节点', () => {
    const root = { path: '', name: '全部文档', docs: 0, chunks: 0, bytes: 0, children: [] };
    mergeKbFolders(root, ['甲', '甲', '../坏', '甲/乙', '甲/乙']);
    assert.equal(root.children.filter((node) => node.path === '甲').length, 1);
    const jia = root.children.find((node) => node.path === '甲');
    assert.equal(jia.children.filter((node) => node.path === '甲/乙').length, 1);
    assert.equal(root.children.some((node) => node.path.includes('..')), false);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `\n      ${item.error?.message ?? item.error}`}`);
}
console.log(`kb-folders: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
