#!/usr/bin/env node
/**
 * 把 mnemon 的记忆迁进办公记忆（命令行入口）。
 *
 * 为什么要有命令行入口：设置页只能读写配置，没有执行通道；而迁移需要在
 * **插件新版本还没被 DSH 加载**（改了源码必须重启）时就能做一次。命令行直接
 * import 源码，不受进程内模块缓存影响。
 *
 * 跑法：
 *   node scripts/migrate-mnemon.mjs                 # 迁当前工作目录
 *   node scripts/migrate-mnemon.mjs --dry-run       # 先看一眼会迁什么
 *   node scripts/migrate-mnemon.mjs --dir "D:\某项目" --source .mnemon
 *   node scripts/migrate-mnemon.mjs --memory-dir .office/memory --global "C:\Users\me\.mnemon"
 */
import { resolve } from 'node:path';
import process from 'node:process';
import { resolveConfig } from '../src/config.js';
import { migrateMnemon, renderMigration } from '../src/migrate.js';

function parseArgs(argv) {
    const options = { dir: process.cwd(), source: undefined, global: undefined, memoryDir: undefined, dryRun: false };
    for (let i = 0; i < argv.length; i += 1) {
        const arg = argv[i];
        const next = argv[i + 1];
        if (arg === '--dir' && next !== undefined) { options.dir = resolve(next); i += 1; continue; }
        if (arg === '--source' && next !== undefined) { options.source = next; i += 1; continue; }
        if (arg === '--global' && next !== undefined) { options.global = next; i += 1; continue; }
        if (arg === '--memory-dir' && next !== undefined) { options.memoryDir = next; i += 1; continue; }
        if (arg === '--dry-run') { options.dryRun = true; continue; }
        if (arg === '--help' || arg === '-h') { options.help = true; continue; }
        throw new Error(`不认识的参数：${arg}`);
    }
    return options;
}

const options = parseArgs(process.argv.slice(2));
if (options.help === true) {
    console.log(`用法：node scripts/migrate-mnemon.mjs [--dir <工作目录>] [--source <mnemon 数据根>] [--global <全局根>] [--memory-dir <记忆目录>] [--dry-run]`);
    process.exit(0);
}

const config = resolveConfig(options.memoryDir === undefined ? {} : { memory: { dir: options.memoryDir } });
const receipt = await migrateMnemon({
    root: options.dir,
    memory: config.memory,
    source: options.source,
    globalSource: options.global,
    dryRun: options.dryRun,
});
console.log(renderMigration(receipt));
console.log('');
console.log(JSON.stringify({
    workspace: receipt.sources.workspace,
    global: receipt.sources.global,
    databases: receipt.sources.databases,
    hot: receipt.hot,
    archive: { added: receipt.archive.added, skipped: receipt.archive.skipped, deleted: receipt.archive.deleted, file: receipt.archive.file, failed: receipt.archive.failed },
}, null, 2));
process.exit(receipt.archive.failed.length > 0 ? 1 : 0);
