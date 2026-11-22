/**
 * 清理 .office/runtime/ 下的运行期文件（第八十四轮）。
 *
 * 用法：
 *   node scripts/clean-office.mjs --older-than 7 --dry-run
 *   node scripts/clean-office.mjs --older-than 7
 *   node scripts/clean-office.mjs --all
 *
 * 默认只碰 cache / tmp / diag / labeling；memory、kb、backups、search 永远不删。
 * --all 会清空这四个目录（仍保留目录本身）。这个脚本不做「回收站」，--all 前先 --dry-run。
 */
import { existsSync } from 'node:fs';
import { readdir, rm, rmdir, stat } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const argv = process.argv.slice(2);
const dryRun = argv.includes('--dry-run');
const all = argv.includes('--all');
const olderIndex = argv.indexOf('--older-than');
const days = olderIndex >= 0 && Number.isFinite(Number(argv[olderIndex + 1])) ? Number(argv[olderIndex + 1]) : 7;
const rootIndex = argv.indexOf('--root');
const ROOT = resolve(rootIndex >= 0 && argv[rootIndex + 1] ? argv[rootIndex + 1] : resolve(HERE, '../../..'));
const RUNTIME = join(ROOT, '.office', 'runtime');
const TARGETS = ['cache', 'tmp', 'diag', 'labeling'];
const cutoff = Date.now() - days * 24 * 60 * 60 * 1000;

if (!all && days < 0) {
    console.error('--older-than 不能是负数');
    process.exit(1);
}

async function walk(dir, rel = '') {
    const rows = [];
    for (const name of await readdir(dir).catch(() => [])) {
        const target = join(dir, name);
        const info = await stat(target).catch(() => null);
        if (info === null) continue;
        if (info.isDirectory()) {
            rows.push(...await walk(target, rel + name + '/'));
        } else {
            rows.push({ path: target, rel: rel + name, size: info.size, mtimeMs: info.mtimeMs });
        }
    }
    return rows;
}

const report = { dryRun, all, days, root: ROOT, targets: [], removedFiles: 0, removedBytes: 0, keptFiles: 0 };
for (const name of TARGETS) {
    const dir = join(RUNTIME, name);
    if (!existsSync(dir)) continue;
    const files = await walk(dir);
    const doomed = all ? files : files.filter((file) => file.mtimeMs < cutoff);
    const bytes = doomed.reduce((sum, file) => sum + file.size, 0);
    report.targets.push({ name, files: files.length, doomed: doomed.length, bytes, human: (bytes / 1024 / 1024).toFixed(1) + ' MB' });
    report.removedFiles += doomed.length;
    report.removedBytes += bytes;
    report.keptFiles += files.length - doomed.length;
    if (dryRun) continue;
    for (const file of doomed) await rm(file.path, { force: true });
    // 自底向上收掉空目录，只收这棵子树里的。
    const prune = async (current) => {
        for (const child of await readdir(current).catch(() => [])) {
            const target = join(current, child);
            const info = await stat(target).catch(() => null);
            if (info?.isDirectory()) await prune(target);
        }
        if (current !== dir) await rmdir(current).catch(() => {});
    };
    await prune(dir);
}

console.log(JSON.stringify(report, null, 2));
if (!dryRun) console.log(`\n清理完成：删除 ${report.removedFiles} 个文件，释放 ${(report.removedBytes / 1024 / 1024).toFixed(1)} MB。`);
