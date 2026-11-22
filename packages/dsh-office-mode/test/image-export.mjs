#!/usr/bin/env node
/**
 * office.ppt.images 导出专项自测（第七十二轮 S-05）。
 *
 * 钉住三件事：
 *   1. 显式 out 的每个文件都经 env.writeFile，出现在统一写盘清单里（office_run 的产物清单靠它）；
 *   2. 目标已存在时默认跳过并在 skipped 里说明原因，overwrite:true 才重写；
 *   3. 默认缓存目录按「词干 + 包内容哈希」分桶：同 basename 不同源文档不会互相覆盖，
 *      同一份稿子重复导出命中已有文件（existing:true）而不是重写。
 *
 * 跑法：node test/image-export.mjs（加 --keep 保留 .office/runtime/tmp/image-export/）。
 * 本套件不联网。
 */
import { deflateSync } from 'node:zlib';
import { mkdirSync, rmSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { create, images } from '../src/formats/ppt.js';
import { createEnv } from '../src/engine/kit.js';
import { crc32 } from '../src/engine/zip.js';
import { resolveTheme } from '../src/engine/theme.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const TMP = resolve(HERE, '..', '..', '..', '.office', 'runtime', 'tmp', 'image-export');
const KEEP = process.argv.includes('--keep');

const failures = [];
let checks = 0;

function check(condition, label, detail = '') {
    checks += 1;
    if (!condition) failures.push(detail ? `${label} —— ${detail}` : label);
}

function checkEqual(actual, expected, label) {
    check(actual === expected, label, `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`);
}

function checkIncludes(haystack, needle, label) {
    const hit = Array.isArray(haystack) ? haystack.includes(needle) : String(haystack).includes(needle);
    check(hit, label, `没有找到 ${JSON.stringify(needle)}`);
}

/** 手搓真 PNG（渐变 + tint 让两张图字节不同）：零依赖，也顺带验证魔数。 */
function makePng(width, height, tint) {
    const raw = Buffer.alloc((width * 3 + 1) * height);
    for (let y = 0; y < height; y += 1) {
        const rowAt = y * (width * 3 + 1);
        raw[rowAt] = 0;
        for (let x = 0; x < width; x += 1) {
            const at = rowAt + 1 + x * 3;
            raw[at] = Math.round((x / width) * 200) + tint;
            raw[at + 1] = Math.round((y / height) * 180) + 50;
            raw[at + 2] = 180;
        }
    }
    const chunk = (type, data) => {
        const head = Buffer.alloc(8);
        head.writeUInt32BE(data.length, 0);
        head.write(type, 4, 'latin1');
        const tail = Buffer.alloc(4);
        tail.writeUInt32BE(crc32(Buffer.concat([head.subarray(4), data])), 0);
        return Buffer.concat([head, data, tail]);
    };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw, { level: 9 })),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

// ---------------------------------------------------------------- 样例准备

rmSync(TMP, { recursive: true, force: true });
mkdirSync(TMP, { recursive: true });
const env = createEnv({ root: TMP, themeResolver: resolveTheme });
const pngA = makePng(120, 80, 0);
const pngB = makePng(140, 90, 60);
env.writeFile('assets/a.png', pngA);
env.writeFile('assets/b.png', pngB);

/** 建一份只有封面 + 一页图片的稿子，路径刻意用同名 deck.pptx。 */
function build(relPath, imagePath, title) {
    const deck = create({ title, theme: 'plain', path: relPath, size: '16:9' }, env);
    deck.cover({ title });
    deck.image({ path: imagePath, title: '图片页', caption: '测试图片' });
    return deck.save();
}

const reportA = build('a/deck.pptx', 'assets/a.png', '样例 A');
const reportB = build('b/deck.pptx', 'assets/b.png', '样例 B');
checkEqual(reportA.path, 'a/deck.pptx', '样例 A 落盘路径');
checkEqual(reportB.path, 'b/deck.pptx', '样例 B 落盘路径');

// ---------------------------------------------------------------- 1. 显式 out
{
    const first = images(reportA.path, { out: 'out-explicit' }, env);
    checkEqual(first.ok, true, 'images() 的 ok');
    checkEqual(first.count, 1, '第一次导出应写出 1 张');
    checkEqual(first.files.length, 1, 'files 与 count 一致');
    checkEqual(first.dir, 'out-explicit', 'dir 是解析后的相对路径');
    checkEqual(first.overwrite, false, '默认 overwrite=false');
    checkEqual(first.skipped.length, 0, `第一次不应跳过：${JSON.stringify(first.skipped)}`);
    const exported = first.files[0].path;
    check(exported.startsWith('out-explicit/'), `导出路径应落在 out 里：${exported}`);
    check(env.writes.some((entry) => entry.path === exported), '导出物必须进 env.writes（统一产物清单）');
    check(Buffer.compare(env.readFile(exported), pngA) === 0, '导出字节必须与源图逐字节一致');
    check(first.files[0].bytes === pngA.length, 'files[0].bytes 是真实字节数');

    const second = images(reportA.path, { out: 'out-explicit' }, env);
    checkEqual(second.count, 0, '目标已存在时默认不覆盖');
    checkEqual(second.files.length, 0, '未写出的文件不进 files');
    check(second.skipped.some((item) => /overwrite/.test(String(item.reason))), '跳过原因要给出覆盖办法');
    check(Buffer.compare(env.readFile(exported), pngA) === 0, '跳过时不得改写原文件');

    const before = env.writes.filter((entry) => entry.path === exported).length;
    const third = images(reportA.path, { out: 'out-explicit', overwrite: true }, env);
    checkEqual(third.count, 1, 'overwrite 时应重写 1 张');
    checkEqual(third.overwrite, true, '返回值要能看出这次是覆盖');
    check(env.writes.filter((entry) => entry.path === exported).length > before, 'overwrite 必须再走一次 env.writeFile');
}

// ---------------------------------------------------------------- 2. 相对 out 不得逃逸
{
    let message = '';
    try {
        images(reportA.path, { out: '../escape' }, env);
    } catch (error) {
        message = String(error.message);
    }
    checkIncludes(message, '工作目录', '相对 out 逃出工作目录要当场拒绝');
}

// ---------------------------------------------------------------- 3. 默认缓存：同名不同源分桶
{
    const base = resolve(TMP, 'cache');
    const cache = {
        artifacts: 0,
        hits: 0,
        ensureDir(name) {
            const dir = resolve(base, name);
            mkdirSync(dir, { recursive: true });
            return dir;
        },
        noteArtifact() {
            cache.artifacts += 1;
        },
        noteHit() {
            cache.hits += 1;
        },
    };

    const cachedA = images(reportA.path, {}, env, { cache });
    const cachedB = images(reportB.path, {}, env, { cache });
    check(cachedA.dir !== cachedB.dir, `同 basename 不同源文档必须分桶：${cachedA.dir} / ${cachedB.dir}`);
    checkIncludes(cachedA.dir, 'deck-', '缓存桶名要带词干');
    checkEqual(cachedA.count, 1, '缓存路径也能导出');
    checkEqual(cachedB.count, 1, '第二份同名稿子也能导出');
    check(Buffer.compare(env.readFile(cachedA.files[0].path), pngA) === 0, '缓存路径的 A 图字节一致');
    check(Buffer.compare(env.readFile(cachedB.files[0].path), pngB) === 0, '缓存路径的 B 图字节一致且不被 A 覆盖');
    checkEqual(cache.artifacts, 2, '两次真实写盘记两次 artifact');

    const again = images(reportA.path, {}, env, { cache });
    checkEqual(again.count, 1, '同一份稿子重复导出仍要返回文件');
    checkEqual(again.files[0].path, cachedA.files[0].path, '命中缓存时路径不变');
    checkEqual(again.files[0].existing, true, '命中缓存要标 existing');
    checkEqual(cache.artifacts, 2, '命中缓存不应再写盘');
    checkEqual(cache.hits, 1, '命中缓存要记一次 hit');
}

// ---------------------------------------------------------------- 4. 没有可写目录
{
    let message = '';
    try {
        images(reportA.path, {}, env);
    } catch (error) {
        message = String(error.message);
    }
    checkIncludes(message, 'out', '没有 out / 缓存时要把下一步说清楚');
}

// ---------------------------------------------------------------- 结果

if (failures.length > 0) {
    console.error(`\nFAIL image-export：${failures.length} 项不通过（共 ${checks} 项断言）`);
    for (const line of failures) console.error(`  ✗ ${line}`);
    if (!KEEP) rmSync(TMP, { recursive: true, force: true });
    process.exit(1);
}

console.log(`\nPASS image-export checks=${checks}`);
if (KEEP) {
    console.log(`产物保留在 ${TMP}`);
} else {
    rmSync(TMP, { recursive: true, force: true });
    console.log(`临时产物已清理：${TMP}`);
}
