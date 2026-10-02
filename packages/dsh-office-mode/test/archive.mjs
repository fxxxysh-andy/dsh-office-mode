/**
 * 归档与会话日志读取（office.archive，历史遗留 `23-5`）的测试。
 *
 * 这条账说的是「会话日志 zip / zstd 在办公模式里读不出」。测试分三层：
 *   1. 判型与五条读法：ZIP / gzip / zstd / 纯文本 / 不认识的二进制；
 *   2. 有界性与诚实性：翻页、条目上限、二进制正文不假装成文本、
 *      `office.files.read()` 读到压缩包时给的是可执行的错（不是一段乱码）；
 *   3. **真实形状**：工作区里的会话导出（`session*.zip`，里面是
 *      `session.v4.jsonl` + `media/sha256:*`）—— 存在就按真包测，不存在就 SKIP。
 *
 * 跑法：node test/archive.mjs
 */
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { randomBytes } from 'node:crypto';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { gzipSync, zstdCompressSync } from 'node:zlib';

import { createEnv } from '../src/engine/kit.js';
import { zip } from '../src/engine/zip.js';
import { archiveExtract, archiveFind, archiveHintForBytes, archiveInfo, archiveList, archiveText, sniffArchive } from '../src/archive.js';
import { resolveConfig } from '../src/config.js';
import { executeRun } from '../src/run.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        if (error?.skip === true) results.push({ name, ok: true, note: `SKIP ${error.message}` });
        else results.push({ name, ok: false, error });
    }
}
function skip(message) {
    const error = new Error(message);
    error.skip = true;
    throw error;
}

const root = mkdtempSync(join(tmpdir(), 'office-archive-'));
const env = createEnv({ root });
const encoder = new TextEncoder();

/** 一份「像会话导出」的包：一条 jsonl + 一条二进制媒体。 */
const jsonl = [
    '{"type":"session","version":4,"id":"session-test","cwd":"D:/demo"}',
    '{"type":"user","text":"把季度报表做出来"}',
    '{"type":"assistant","text":"我先用 office_run 建一份 xlsx"}',
    '{"type":"tool","name":"office_run","ok":true}',
    '{"type":"assistant","text":"office_run 写出 明细.xlsx"}',
    '{"type":"assistant","text":"office_run(1) 返回 ok"}',
].join('\n');
const fakePng = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]);
const subJsonl = '{"type":"session","version":4,"id":"sub"}\n{"type":"assistant","text":"office_run"}\n';
const zipPath = join(root, 'session.v4.zip');
writeFileSync(zipPath, zip([
    { name: 'session.v4.jsonl', data: jsonl },
    { name: 'subagents/abc/session.v4.jsonl', data: subJsonl },
    { name: 'media/sha256:deadbeef.png', data: fakePng },
]));
/** 期望的命中行数：直接在夹具文本上数，不写死数字（夹具改了就跟着变）。 */
const expectedOfficeRunLines = [jsonl, subJsonl]
    .reduce((sum, text) => sum + text.split('\n').filter((line) => /office_run/i.test(line)).length, 0);
writeFileSync(join(root, 'notes.jsonl.gz'), gzipSync(encoder.encode(jsonl)));
writeFileSync(join(root, 'notes.jsonl.zst'), zstdCompressSync(encoder.encode(jsonl)));
writeFileSync(join(root, 'plain.jsonl'), jsonl);

await check('sniffArchive：按魔数判型（ZIP / gzip / zstd / 文本 / 二进制）', async () => {
    assert.equal(sniffArchive(new Uint8Array(readFileSync(zipPath))), 'zip');
    assert.equal(sniffArchive(new Uint8Array(readFileSync(join(root, 'notes.jsonl.gz')))), 'gzip');
    assert.equal(sniffArchive(new Uint8Array(readFileSync(join(root, 'notes.jsonl.zst')))), 'zstd');
    assert.equal(sniffArchive(encoder.encode(jsonl)), 'plain');
    assert.equal(sniffArchive(fakePng), undefined, 'PNG 这类二进制不该被判成文本或归档');
    assert.equal(sniffArchive(new Uint8Array(0)), undefined);
    return '5 种输入';
});

await check('archive.info：条目数、解压后体积、像文本的条目数；纯文本给出替代路径', async () => {
    const info = archiveInfo(zipPath, env);
    assert.equal(info.kind, 'zip');
    assert.equal(info.entries, 3);
    assert.equal(info.textEntries, 2);
    // 用 UTF-8 字节数而不是 JS 的 .length：条目是按字节存的，中文一个字占 3 字节。
    assert.equal(
        info.uncompressedBytes,
        Buffer.byteLength(jsonl, 'utf8') + Buffer.byteLength(subJsonl, 'utf8') + fakePng.length,
        '解压后体积是三条之和',
    );
    assert.ok(info.sample.includes('session.v4.jsonl'));
    const plain = archiveInfo('plain.jsonl', env);
    assert.equal(plain.kind, 'plain');
    assert.match(plain.note, /office\.files\.read/);
    const bad = join(root, 'blob.bin');
    writeFileSync(bad, fakePng);
    assert.equal(archiveInfo('blob.bin', env).kind, 'binary');
    return `${info.entries} 条 / ${info.uncompressedBytes} B`;
});

await check('archive.list：列全部条目，带字节数与压缩方式', async () => {
    const list = archiveList(zipPath, env);
    assert.equal(list.count, 3);
    assert.equal(list.truncated, false);
    assert.deepEqual(list.entries.map((item) => item.name), ['session.v4.jsonl', 'subagents/abc/session.v4.jsonl', 'media/sha256:deadbeef.png']);
    assert.ok(list.entries.every((item) => item.bytes > 0));
    const limited = archiveList(zipPath, env, { limit: 2 });
    assert.equal(limited.entries.length, 2);
    assert.equal(limited.truncated, true);
    return '3 条';
});

await check('archive.text：默认挑像文本的那条；gzip / zstd 只有一条正文', async () => {
    const text = archiveText(zipPath, env);
    assert.equal(text.entry, 'session.v4.jsonl');
    assert.equal(text.chars, jsonl.length);
    assert.equal(text.truncated, false);
    assert.equal(text.text, jsonl);
    for (const name of ['notes.jsonl.gz', 'notes.jsonl.zst']) {
        const item = archiveText(name, env);
        assert.equal(item.kind, name.endsWith('.gz') ? 'gzip' : 'zstd');
        assert.equal(item.text, jsonl, `${name} 解出来应当与原文一致`);
    }
    return 'zip / gzip / zstd';
});

await check('archive.text：支持按条目取、按后缀模糊匹配、offset 翻页', async () => {
    const single = archiveText(zipPath, env, { entry: 'subagents/abc/session.v4.jsonl' });
    assert.match(single.text, /"id":"sub"/);
    const fuzzy = archiveText(zipPath, env, { entry: 'abc/session.v4.jsonl' });
    assert.match(fuzzy.text, /"id":"sub"/, '名字太长时允许用后缀定位');
    const mid = archiveText(zipPath, env, { entry: 'abc/session' });
    assert.match(mid.text, /"id":"sub"/, '也允许用名字里的一段定位');
    assert.throws(() => archiveText(zipPath, env, { entry: 'nope.jsonl' }), /没有「nope\.jsonl」/);
    const page = archiveText(zipPath, env, { offset: 10, maxChars: 20 });
    assert.equal(page.text, jsonl.slice(10, 30));
    assert.equal(page.truncated, true);
    assert.equal(page.nextOffset, 30);
    return 'entry / 后缀 / 翻页';
});

await check('archive.text：二进制条目如实说「这是二进制」，不假装成文本', async () => {
    const item = archiveText(zipPath, env, { entry: 'media/sha256:deadbeef.png' });
    assert.equal(item.text, '');
    assert.match(item.note, /二进制/);
    return 'PNG 被认出来';
});

await check('archive.find：按行检索（正则 / 字面量回退 / context / 上限）', async () => {
    const hits = archiveFind(zipPath, env, { pattern: 'office_run' });
    assert.equal(hits.hits, expectedOfficeRunLines, '两个 jsonl 条目都要被扫到');
    assert.ok(hits.matches.every((item) => typeof item.line === 'number'));
    assert.equal(hits.matches[0].entry, 'session.v4.jsonl');
    assert.equal(hits.matches[0].line, 3, '第一条命中在第 3 行');
    // 普通文本里的正则元字符要能当字面量用（`office_run(1` 不是合法正则，
    // 退回字面量匹配后应当命中那一行）。
    const literal = archiveFind('plain.jsonl', env, { pattern: 'office_run(1' });
    assert.equal(literal.hits, 1);
    const context = archiveFind(zipPath, env, { pattern: '季度报表', context: 1, entry: 'session.v4.jsonl' });
    assert.equal(context.hits, 1);
    assert.equal(context.matches[0].before.length, 1);
    const capped = archiveFind('plain.jsonl', env, { pattern: 'e', maxHits: 2 });
    assert.equal(capped.hits, 2);
    assert.equal(capped.truncated, true);
    const single = archiveFind(zipPath, env, { pattern: 'office_run', entry: 'subagents/abc/session.v4.jsonl' });
    assert.equal(single.hits, 1, '限定 entry 时只扫那一条');
    // 二进制条目与超大条目都不该被整条解码（只扫描文本、跳过并如实报出来）
    const skipped = archiveFind(zipPath, env, { pattern: 'PNG' });
    assert.equal(skipped.hits, 0, 'PNG 那条是二进制，不该被当文本扫');
    // 二进制条目是**静默跳过**（不当成「超大」报出来）：这条契约要钉住，
    // 否则「跳过了什么」既没有断言、也没有回归保护。
    assert.equal(skipped.scannedEntries, 2, '只扫两个文本条目');
    assert.ok(!(skipped.skippedEntries ?? []).some((item) => /\.png$/.test(item.entry)), '二进制不进 skippedEntries');
    // 显式点名一条二进制时，如实报「二进制（不是文本）」而不是拿它当文本扫
    const named = archiveFind(zipPath, env, { pattern: 'PNG', entry: 'media/sha256:deadbeef.png' });
    assert.equal(named.hits, 0);
    assert.equal(named.scannedEntries, 0);
    assert.equal(named.skippedEntries?.[0]?.reason, '二进制（不是文本）');
    return `${hits.hits} 处命中`;
});

await check('archive.extract：把一条正文落到工作目录，默认不覆盖', async () => {
    const extracted = archiveExtract(zipPath, env, { entry: 'session.v4.jsonl', to: 'logs/session.jsonl' });
    assert.equal(extracted.to, 'logs/session.jsonl');
    assert.equal(readFileSync(join(root, 'logs/session.jsonl'), 'utf8'), jsonl);
    assert.equal(extracted.lines, 6);
    assert.throws(() => archiveExtract(zipPath, env, { entry: 'session.v4.jsonl', to: 'logs/session.jsonl' }), /已存在/);
    const again = archiveExtract(zipPath, env, { entry: 'session.v4.jsonl', to: 'logs/session.jsonl', overwrite: true });
    assert.equal(again.bytes, extracted.bytes);
    // 默认落名由包名 + 条目名拼出来，中文路径也能用。
    const auto = archiveExtract('notes.jsonl.zst', env);
    assert.equal(auto.to, 'notes.jsonl.txt');
    return auto.to;
});

await check('archiveHintForBytes：文本不算归档、ZIP 认出条目数', async () => {
    assert.equal(archiveHintForBytes(encoder.encode(jsonl)), undefined);
    assert.match(archiveHintForBytes(new Uint8Array(readFileSync(zipPath))), /zip 归档（3 个条目）/);
    assert.match(archiveHintForBytes(new Uint8Array(readFileSync(join(root, 'notes.jsonl.zst')))), /zstd 归档/);
    return '提示文案一致';
});

await check('炸弹包：按中央目录的**声明值**在解压前挡住（不是解完再判）', async () => {
    // 手改中央目录里的「解压后体积」为 300 MB（压缩炸弹就是靠这种声明值吃内存）。
    // 关键在**解压之前**判：解完再判时内存已经被吃掉了。
    const bytes = zip([{ name: 'bomb.txt', data: 'x'.repeat(4096) }]);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let patched = false;
    for (let i = 0; i + 4 <= bytes.length; i += 1) {
        if (view.getUint32(i, true) === 0x02014b50) {
            view.setUint32(i + 24, 300 * 1024 * 1024, true);
            patched = true;
            break;
        }
    }
    assert.equal(patched, true, '没找到中央目录项，夹具没生效');
    writeFileSync(join(root, 'bomb.zip'), bytes);
    const rssBefore = process.memoryUsage().rss;
    assert.throws(() => archiveInfo('bomb.zip', env), /office\.archive：ZIP 解析失败.*超过上限/s);
    assert.throws(() => archiveText('bomb.zip', env), /超过上限/);
    const rssAfter = process.memoryUsage().rss;
    assert.ok(rssAfter - rssBefore < 64 * 1024 * 1024,
        `解压前就挡住了，内存不该涨这么多：${Math.round((rssAfter - rssBefore) / 1024 / 1024)} MB`);
    // files.read 那条提示只数条目、不解压，所以也不该被炸弹拖爆。
    const hint = archiveHintForBytes(new Uint8Array(readFileSync(join(root, 'bomb.zip'))));
    assert.match(hint, /zip 归档（1 个条目）/);
    return `RSS 变化 ${Math.round((rssAfter - rssBefore) / 1024 / 1024)} MB`;
});

await check('坏包与坏 gzip：报的是 office.archive 的口径，不是 zlib 的内部消息', async () => {
    const corruptZip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5, 6, 7, 8]);
    writeFileSync(join(root, 'broken.zip'), corruptZip);
    assert.throws(() => archiveInfo('broken.zip', env), /office\.archive：ZIP 解析失败/);
    const badGzip = new Uint8Array([0x1f, 0x8b, 0x08, 0x00, 0, 0, 0, 0, 0, 0, 9, 9, 9]);
    writeFileSync(join(root, 'broken.gz'), badGzip);
    assert.throws(() => archiveText('broken.gz', env), /office\.archive：gzip 解压失败/);
    return 'zip / gzip 都是本模块口径';
});

await check('压缩方式如实报：stored 与 deflate 分得开', async () => {
    // 真随机字节压不小 → zip() 会以 stored 存；长重复文本 → deflate。
    // 这一条钉住「method 不是写死的 deflate」。
    const random = randomBytes(4096);
    writeFileSync(join(root, 'methods.zip'), zip([
        { name: 'stored.bin', data: new Uint8Array(random) },
        { name: 'deflated.txt', data: 'a'.repeat(8192) },
    ]));
    const list = archiveList('methods.zip', env);
    const methods = Object.fromEntries(list.entries.map((item) => [item.name, item.method]));
    assert.equal(methods['stored.bin'], 'stored', `随机字节应当以 stored 存：${JSON.stringify(methods)}`);
    assert.equal(methods['deflated.txt'], 'deflate');
    return JSON.stringify(methods);
});

await check('zstd 不可用时明确报「本机不支持」，不退化成乱码', async () => {
    // 这里不去动 node:zlib，只钉住错误文案的存在：真缺函数时 archiveText 会抛这句。
    const source = readFileSync(new URL('../src/archive.js', import.meta.url), 'utf8');
    assert.match(source, /本机 Node 不支持 zstd/);
    assert.match(source, /zstdDecompressSync/);
    return '文案在源码里';
});

// ── 经 office_run 的 SDK 面（脚本里看得到的那一层） ─────────────────────────

const runScript = (script, extra = {}) => executeRun({ script }, { agent: { session: { header: { cwd: root } } } }, resolveConfig(extra));

await check('SDK 面：office.archive 五个动作在 office_run 里可用', async () => {
    const script = `
        const info = office.archive.info('session.v4.zip');
        const list = office.archive.list('session.v4.zip', { limit: 2 });
        const text = office.archive.text('session.v4.zip', { entry: 'session.v4.jsonl', maxChars: 40 });
        const hits = office.archive.find('session.v4.zip', { pattern: 'office_run', maxHits: 2 });
        const out = office.archive.extract('notes.jsonl.gz', { to: 'logs/from-gz.jsonl' });
        return { info, list, text, hits, out };`;
    const result = await runScript(script);
    assert.equal(result.ok, true, result.error?.message);
    const value = result.returned;
    assert.equal(value.info.kind, 'zip');
    assert.equal(value.list.count, 3);
    assert.equal(value.text.returned, 40);
    assert.equal(value.text.truncated, true);
    assert.ok(value.hits.hits > 0);
    assert.equal(value.out.to, 'logs/from-gz.jsonl');
    return '5 个动作';
});

await check('诚实性：office.files.read 读到压缩包时给出可执行的错，而不是乱码', async () => {
    const zipped = await runScript(`return office.files.read('session.v4.zip');`);
    assert.equal(zipped.ok, false);
    assert.match(zipped.error.message, /zip 归档（3 个条目）/);
    assert.match(zipped.error.message, /office\.archive/);
    const gz = await runScript(`return office.files.read('notes.jsonl.gz');`);
    assert.equal(gz.ok, false);
    assert.match(gz.error.message, /gzip 归档/);
    // 普通文本照常读得出来（不要把正常路径也堵掉）。
    const plain = await runScript(`return office.files.read('plain.jsonl');`);
    assert.equal(plain.ok, true, plain.error?.message);
    assert.equal(plain.returned, jsonl);
    // edit / template 会整份写回：按乱码读再写回等于把压缩包毁掉，所以它们走同一个入口。
    const before = readFileSync(zipPath);
    for (const script of [
        `return office.files.edit('session.v4.zip', [['a', 'b']]);`,
        `return office.files.template('session.v4.zip', { a: 'b' });`,
    ]) {
        const result = await runScript(script);
        assert.equal(result.ok, false, '改压缩包必须报错');
        assert.match(result.error.message, /office\.archive/);
    }
    assert.ok(before.equals(readFileSync(zipPath)), '报错之后压缩包必须一个字节都没变');
    return '压缩包报错 / 文本照常 / 不改包';
});

await check('office_help：archive 话题在（默认层指路 + 全文层给流程）', async () => {
    const { buildHelp } = await import('../src/docs.js');
    const brief = await buildHelp('archive');
    assert.match(brief.text, /office\.archive\.info/);
    assert.match(brief.text, /office\.files\.read/);
    const full = await buildHelp('archive', { detail: true });
    assert.match(full.text, /office\.archive\.find/);
    assert.match(full.text, /session10\.zip/);
    const index = await buildHelp('');
    assert.match(index.text, /archive：zip \/ gz \/ zst/);
    return '默认层 + 全文层';
});

// ── 真实形状：工作区里的会话导出 ────────────────────────────────────────────

await check('真包：工作区里的 session*.zip 能读出主 jsonl 与子代理 jsonl', async () => {
    // 仓库根（test/ → dsh-office-mode → packages → 根）。会话导出没被跟踪，但常驻在工作区。
    const sessionDir = fileURLToPath(new URL('../../../', import.meta.url));
    const candidates = ['session10.zip', 'session.zip', 'session2.zip']
        .map((name) => join(sessionDir, name))
        .filter((file) => existsSync(file));
    if (candidates.length === 0) skip('工作区里没有 session*.zip（那是没被跟踪的会话导出）');
    const picked = candidates[0];
    const info = archiveInfo(picked, env);
    assert.equal(info.kind, 'zip');
    assert.ok(info.entries >= 1);
    const list = archiveList(picked, env);
    assert.ok(list.entries.some((item) => /session\.v\d+\.jsonl$/.test(item.name)), '包里应当有 session.v*.jsonl');
    const text = archiveText(picked, env, { maxChars: 2000 });
    assert.match(text.text, /"type":"session"/);
    assert.match(text.text, /"version":\d+/);
    const hits = archiveFind(picked, env, { pattern: 'office_run', maxHits: 5 });
    return `${picked.split(/[\\/]/).pop()}：${info.entries} 条 / 主稿 ${text.chars} 字符 / office_run 命中 ${hits.hits}`;
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`archive: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
