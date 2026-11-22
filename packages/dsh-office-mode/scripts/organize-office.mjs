/**
 * 把旧版 .office 里的运行期目录迁进 .office/runtime/（第八十四轮）。
 *
 * 用法：
 *   node scripts/organize-office.mjs --dry-run
 *   node scripts/organize-office.mjs
 *   node scripts/organize-office.mjs --root D:/某个工作区
 *
 * 只搬运行期目录：cache / tmp / diag / labeling。memory、kb、backups、search 不动。
 * 目标已存在时按目录合并，同名文件不覆盖（冲突会逐条列为 skipped）。
 */
import { existsSync } from 'node:fs';
import { mkdir, readdir, rename, rmdir, stat } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const argv = process.argv.slice(2);
const dryRun = argv.includes('--dry-run');
const rootIndex = argv.indexOf('--root');
const ROOT = resolve(rootIndex >= 0 && argv[rootIndex + 1] ? argv[rootIndex + 1] : resolve(HERE, '../../..'));
const OFFICE = join(ROOT, '.office');
const MOVES = [
    ['cache', 'runtime/cache'],
    ['tmp', 'runtime/tmp'],
    ['diag', 'runtime/diag'],
    ['labeling', 'runtime/labeling'],
];

async function sizeOf(target) {
    let total = 0;
    const stack = [target];
    while (stack.length > 0) {
        const current = stack.pop();
        let info;
        try { info = await stat(current); } catch { continue; }
        if (info.isDirectory()) {
            for (const name of await readdir(current).catch(() => [])) stack.push(join(current, name));
        } else {
            total += info.size;
        }
    }
    return total;
}

function human(bytes) {
    if (bytes >= 1024 * 1024) return (bytes / 1024 / 1024).toFixed(1) + ' MB';
    if (bytes >= 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return bytes + ' B';
}

const report = { dryRun, root: ROOT, moves: [], skipped: [] };

async function mergeTree(source, target, rel) {
    await mkdir(target, { recursive: true });
    for (const name of await readdir(source)) {
        const from = join(source, name);
        const to = join(target, name);
        const info = await stat(from);
        if (!existsSync(to)) {
            if (!dryRun) await rename(from, to);
            report.moves.push({ from: rel + '/' + name, to: rel.replace(/^cache|^tmp|^diag|^labeling/, 'runtime/$&') + '/' + name });
            continue;
        }
        const existing = await stat(to);
        if (info.isDirectory() && existing.isDirectory()) {
            await mergeTree(from, to, rel + '/' + name);
            if (!dryRun) await rmdir(from).catch(() => {});
        } else {
            report.skipped.push(rel + '/' + name);
        }
    }
}

for (const [legacyName, runtimeName] of MOVES) {
    const from = join(OFFICE, legacyName);
    const to = join(OFFICE, runtimeName);
    if (!existsSync(from)) continue;
    const bytes = await sizeOf(from);
    report.moves.push({ from: legacyName, to: runtimeName, bytes, human: human(bytes) });
    if (!dryRun) await mergeTree(from, to, legacyName);
    if (!dryRun && existsSync(from)) {
        const left = await readdir(from);
        if (left.length === 0) await rmdir(from).catch(() => {});
    }
}

console.log(JSON.stringify(report, null, 2));
if (!dryRun) {
    console.log(report.moves.length === 0 && report.skipped.length === 0
        ? '\n.office 已经是标准布局。'
        : `\n迁移完成：${report.moves.length} 项，冲突跳过 ${report.skipped.length} 项。`);
}
