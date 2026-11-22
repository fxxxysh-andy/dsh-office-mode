/**
 * 文件通道（二进制）与跨调用脚本复用（第五十一轮：`新-8` + `新-7` + `新-13`）。
 *
 * 三条账说的是同一类事 —— **调用方必须自觉**：
 *   1. `新-8`：`office.files` 只认文本（`read` 强制 utf8、`write` 把非字符串 JSON 化、
 *      也没有整份复制）。一份 PNG 从读进来到写回去会在中间被毁掉，而调用方看不出来。
 *   2. `新-7`：跨调用没有复用机制 —— 200 行的构建器只能每次内联进工具参数
 *      （沙箱 `codeGeneration:{strings:false}` 让 `import` / `eval` 都不可用）。
 *   3. `新-13`：文档里的数字与实现漂移（回显上限写 4000、实为 1200；README 教了
 *      一个不存在的拷贝动作）。
 *
 * 这一套按四层钉住：
 *   1. 二进制往返逐字节相等（含 vm realm 里造的 Uint8Array —— 跨 realm 的坑就在这里）；
 *   2. 走错路的报错要**指路**（二进制 → readBytes / copy；PDF → office.pdf）；
 *   3. copy / move 的三条闸与记账语义（同文件拒绝、目标已存在要 overwrite、源不存在要指路）；
 *   4. scriptFile + args 的复用与错误行号，以及「文档里的数字从实现现取」这条自同步。
 *
 * 跑法：node test/files-channel.mjs
 */
import { existsSync, linkSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import assert from 'node:assert/strict';

import { resolveConfig, RETURN_ECHO_CHARS } from '../src/config.js';
import { buildHelp } from '../src/docs.js';
import { createMemory, memoryPaths } from '../src/memory.js';
import { executeRun } from '../src/run.js';
import { testEnv } from '../src/sdk.js';
import { renderRun } from '../src/tools.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const root = mkdtempSync(join(tmpdir(), 'office-files-'));
const config = resolveConfig({});
const exec = { agent: { session: { header: { cwd: root } } } };
const runScript = (args) => executeRun(args, exec, config);
const run = (script, extra = {}) => runScript({ script, ...extra });
const read = (name) => readFileSync(join(root, name));

// 夹具：一份**真二进制**（PNG 魔数 + 空字节 + 0xFF，utf8 解出来必然是乱码），
// 与一份最小 PDF（前 8 KB 里带空字节，且开头就是 %PDF-）。
const PNG_BYTES = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x7f, 0xff, 0xfe, 0x01]);
writeFileSync(join(root, 'logo.png'), PNG_BYTES);
writeFileSync(join(root, 'paper.pdf'), Buffer.concat([Buffer.from('%PDF-1.4\n'), Buffer.from([0x00, 0x01]), Buffer.from('\n%%EOF')]));

// ── 1. 二进制往返 ──────────────────────────────────────────────────────────

await check('office.files：文本读写往返不变，bytes 按 UTF-8 报', async () => {
    writeFileSync(join(root, 'notes.txt'), '第一行\n');
    const result = await run(`
        const text = office.files.read('notes.txt');
        office.files.write('notes.txt', text + '第二行');
        return office.files.stat('notes.txt').bytes;
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned, Buffer.byteLength('第一行\n第二行', 'utf8'));
    assert.equal(read('notes.txt').toString('utf8'), '第一行\n第二行');
});

await check('office.files：readBytes → write 逐字节相等（含空字节与 0xFF）', async () => {
    const result = await run(`
        const bytes = office.files.readBytes('logo.png');
        const written = office.files.write('logo-copy.png', bytes);
        return { length: bytes.length, early: bytes[0], zero: bytes[8], written };
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.length, PNG_BYTES.length);
    assert.equal(result.returned.early, 0x89, 'readBytes 给的是原始字节，不是解出来的字符');
    assert.equal(result.returned.zero, 0x00);
    assert.deepEqual(read('logo-copy.png'), PNG_BYTES, '二进制复制必须逐字节相等');
    assert.equal(result.returned.written.bytes, PNG_BYTES.length);
});

await check('office.files：脚本里 new Uint8Array 造的字节也按字节写（跨 realm 不被 JSON 化）', async () => {
    // 关键在 realm：脚本里的 Uint8Array 是 vm realm 的实例，宿主侧
    // `instanceof Uint8Array` 为假。早先 write 一律 JSON.stringify，写出来的是
    // {"0":1,...} 这种文本 —— 一份「看起来写成功了」的坏文件。
    const result = await run(`
        const bytes = new Uint8Array([0, 1, 2, 255]);
        office.files.write('made.bin', bytes);
        return office.files.stat('made.bin').bytes;
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned, 4);
    assert.deepEqual([...read('made.bin')], [0, 1, 2, 255]);
});

await check('office.files：readBytes 的返回值也能整份写回（Buffer 视图）', async () => {
    const result = await run(`
        const raw = office.files.readBytes('made.bin');
        const back = office.files.write('made-back.bin', raw);
        return back.bytes;
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.deepEqual(read('made-back.bin'), Buffer.from([0, 1, 2, 255]));
});

// ── 2. 走错路的报错要指路 ──────────────────────────────────────────────────

await check('office.files.read：读到二进制直接报错并指到 readBytes / copy，不给乱码', async () => {
    const result = await run(`return office.files.read('logo.png');`);
    assert.equal(result.ok, false, '按文本读二进制必须失败，而不是返回一串替换字符');
    assert.match(result.error.message, /二进制文件/);
    assert.match(result.error.message, /readBytes/);
    assert.match(result.error.message, /office\.files\.copy/);
    // 报错要带上是**哪一个**文件（批量改文件时这是唯一能定位的信息）。
    assert.match(result.error.message, /logo\.png/);
});

await check('office.files.read：PDF 指到 office.pdf（指错路比不指更糟）', async () => {
    const result = await run(`return office.files.read('paper.pdf');`);
    assert.equal(result.ok, false);
    assert.match(result.error.message, /office\.pdf\.info \/ text \/ pages/);
});

await check('office.files.read：不含空字节的二进制靠魔数接住（纯 ASCII 的 PDF / 只截到签名的 PNG）', async () => {
    // 独立复核实测：空字节判据在 245 个二进制夹具里漏 2 个，而「一份前 8 KB 恰好没有
    // 0x00 的合法 PDF」在真实世界完全可能出现 —— 只靠空字节就会把 PDF 当文本返回原文。
    writeFileSync(join(root, 'ascii.pdf'), Buffer.from('%PDF-1.4\n1 0 obj\n<< /Type /Catalog >>\nendobj\ntrailer\n%%EOF\n'));
    writeFileSync(join(root, 'sig-only.png'), Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
    const pdf = await run(`return office.files.read('ascii.pdf');`);
    assert.equal(pdf.ok, false, '纯 ASCII 的 PDF 也必须被认出来');
    assert.match(pdf.error.message, /开头就是二进制魔数/);
    assert.match(pdf.error.message, /office\.pdf\.info \/ text \/ pages/);
    const png = await run(`return office.files.read('sig-only.png');`);
    assert.equal(png.ok, false, '只有 8 字节签名的 PNG 也要被认出来');
    assert.match(png.error.message, /readBytes/);
    // 正向控制：普通文本不许被误伤 —— 特别是**以两字节魔数开头的文本**
    // （「BM 格式说明」「MZ 是…」），那是魔数判据最容易误判的形状，所以只收四字节签名。
    writeFileSync(join(root, 'notes-magic.txt'), 'BM 与 MZ 都是两字节魔数，%PDF- 是 PDF 的。\n');
    const text = await run(`return office.files.read('notes-magic.txt');`);
    assert.equal(text.ok, true, text.error?.message);
    assert.match(text.returned, /%PDF-/);
});

await check('office.files.read：压缩包仍指到 office.archive（老行为不变）', async () => {
    // 用 gzip 魔数而不是 zip：这条提示对 zip 会去数条目（合法的 zip 才数得出来），
    // 而这里要验的是「认出来了就指路」，与包解不解得开无关。
    writeFileSync(join(root, 'log.gz'), Buffer.concat([Buffer.from([0x1f, 0x8b, 0x08, 0x00]), Buffer.alloc(16)]));
    const result = await run(`return office.files.read('log.gz');`);
    assert.equal(result.ok, false);
    assert.match(result.error.message, /office\.archive/);
});

// ── 3. copy / move 的三条闸与记账 ──────────────────────────────────────────

await check('office.files.copy：整份复制逐字节相等，并报出来源路径', async () => {
    const result = await run(`
        const r = office.files.copy('logo.png', 'deep/dir/logo2.png');
        return r;
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.path, 'deep/dir/logo2.png');
    assert.equal(result.returned.from, 'logo.png');
    assert.equal(result.returned.bytes, PNG_BYTES.length);
    assert.deepEqual(read('deep/dir/logo2.png'), PNG_BYTES);
    assert.equal(result.otherFiles.some((file) => file.path === 'deep/dir/logo2.png'), true, '复制出来的产物也要进产物清单');
});

await check('office.files.move：源消失、目标逐字节相等', async () => {
    writeFileSync(join(root, 'to-move.png'), PNG_BYTES);
    const result = await run(`
        const r = office.files.move('to-move.png', 'moved/renamed.png');
        return r;
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.from, 'to-move.png');
    assert.equal(existsSync(join(root, 'to-move.png')), false, '移动之后源路径不该还在');
    assert.deepEqual(read('moved/renamed.png'), PNG_BYTES);
});

await check('office.files.copy / move：目标已存在要显式 overwrite:true', async () => {
    const refused = await run(`return office.files.copy('logo.png', 'logo-copy.png');`);
    assert.equal(refused.ok, false);
    assert.match(refused.error.message, /已存在/);
    assert.match(refused.error.message, /overwrite:true/);
    // 拒绝时**不能**动到已有文件（安静覆盖一份交付物是数据丢失）。
    assert.deepEqual(read('logo-copy.png'), PNG_BYTES);

    const allowed = await run(`
        office.files.write('logo-copy.png', 'not a png anymore');
        const r = office.files.copy('logo.png', 'logo-copy.png', { overwrite: true });
        return r.bytes;
    `);
    assert.equal(allowed.ok, true, allowed.error?.message);
    assert.deepEqual(read('logo-copy.png'), PNG_BYTES);

    const moved = await run(`
        office.files.move('moved/renamed.png', 'logo.png');
        return 'unreachable';
    `);
    assert.equal(moved.ok, false, 'move 到已存在且没给 overwrite 时也要拒绝');
    assert.match(moved.error.message, /已存在/);
    assert.deepEqual(read('logo.png'), PNG_BYTES);
});

await check('office.files.copy / move：同一个文件被拒绝（移动会自毁）', async () => {
    for (const action of ['copy', 'move']) {
        const result = await run(`return office.files.${action}('logo.png', './logo.png', { overwrite: true });`);
        assert.equal(result.ok, false, `${action} 到自身必须被拒绝`);
        assert.match(result.error.message, /源与目标是同一个文件/);
    }
    assert.deepEqual(read('logo.png'), PNG_BYTES, '拒绝之后源文件必须完好');
});

await check('office.files.copy / move：硬链接 / 符号链接别名也是「同一个文件」', async () => {
    // 只比路径字符串会漏掉别名，而 move 到自身的实现是「写目标 + 删源」——
    // 别名下就是自毁（独立复核打穿的一条）。判定交给内核层的 realpath + inode。
    writeFileSync(join(root, 'alias-src.bin'), PNG_BYTES);
    const aliases = [];
    try {
        linkSync(join(root, 'alias-src.bin'), join(root, 'alias-hard.bin'));
        aliases.push('alias-hard.bin');
    } catch {
        // 某些文件系统不支持硬链接：这一层跳过，别把不支持当成失败。
    }
    try {
        symlinkSync(join(root, 'alias-src.bin'), join(root, 'alias-sym.bin'));
        aliases.push('alias-sym.bin');
    } catch {
        // Windows 上创建符号链接需要权限；拿不到就只验硬链接。
    }
    assert.ok(aliases.length > 0, '夹具至少要造出一个别名');
    for (const alias of aliases) {
        for (const action of ['copy', 'move']) {
            const result = await run(`return office.files.${action}('alias-src.bin', '${alias}', { overwrite: true });`);
            assert.equal(result.ok, false, `${action} 到别名 ${alias} 必须被拒绝`);
            assert.match(result.error.message, /源与目标是同一个文件/);
        }
        assert.deepEqual(read('alias-src.bin'), PNG_BYTES, `拒绝之后源文件必须完好（${alias}）`);
    }
});

await check('内核 sameFile：realpath / inode 三层判据，且拿不到 inode 时不误判', () => {
    const env = testEnv(join(root, 'same-file'));
    mkdirSync(join(root, 'same-file'), { recursive: true });
    writeFileSync(join(root, 'same-file', 'a.txt'), 'x');
    writeFileSync(join(root, 'same-file', 'b.txt'), 'y');
    assert.equal(env.sameFile('a.txt', 'a.txt'), true);
    assert.equal(env.sameFile('a.txt', './a.txt'), true);
    assert.equal(env.sameFile('a.txt', 'b.txt'), false);
    assert.equal(env.sameFile('a.txt', '还不存在.txt'), false, '目标不存在时不能误判成同一个文件');
});

await check('office.files.move：只改大小写的改名在 Windows 上给「先移到中间名」的出路', async () => {
    writeFileSync(join(root, 'case-name.md'), '# x');
    const result = await run(`return office.files.move('case-name.md', 'CASE-NAME.md');`);
    assert.equal(result.ok, false);
    assert.match(result.error.message, /源与目标是同一个文件/);
    if (process.platform === 'win32') {
        assert.match(result.error.message, /先 move 到一个中间名/, 'Windows 上要说清怎么改名');
    }
    assert.equal(read('case-name.md').toString('utf8'), '# x');
});

await check('office.files.copy：内核动作的原始 Node 报错被包成人话（带动作名与下一步）', async () => {
    // 目标路径的父级是一个**文件**：原始报错是 `EEXIST: file already exists, mkdir '<绝对路径>'`，
    // 里面是系统措辞与本机绝对路径，模型照着做不了。
    writeFileSync(join(root, 'plain.txt'), 'x');
    const result = await run(`return office.files.copy('logo.png', 'plain.txt/child.png', { overwrite: true });`);
    assert.equal(result.ok, false);
    assert.match(result.error.message, /office\.files\.copy：复制 logo\.png → plain\.txt\/child\.png 失败/);
    assert.doesNotMatch(result.error.message, /[A-Z]:\\/, '报错里不该出现本机绝对路径');
});

await check('office.files.copy / move：源不存在 / 少给路径都给可读的报错', async () => {
    const missing = await run(`return office.files.copy('nope.png', 'x.png');`);
    assert.equal(missing.ok, false);
    assert.match(missing.error.message, /源文件不存在/);
    assert.match(missing.error.message, /office\.files\.list/);

    const noTarget = await run(`return office.files.move('logo.png', '  ');`);
    assert.equal(noTarget.ok, false);
    assert.match(noTarget.error.message, /要给 from 与 to 两个路径/);
});

await check('office.files.copy / move：目录不能当源也不能当目标（报错要指到具体的那一侧）', async () => {
    mkdirSync(join(root, 'a-dir'), { recursive: true });
    const asSource = await run(`return office.files.copy('a-dir', 'x.png');`);
    assert.equal(asSource.ok, false);
    assert.match(asSource.error.message, /a-dir 是一个目录/);

    const asTarget = await run(`return office.files.move('logo.png', 'a-dir');`);
    assert.equal(asTarget.ok, false);
    assert.match(asTarget.error.message, /a-dir 是一个目录，不能当目标文件/);
    assert.deepEqual(read('logo.png'), PNG_BYTES, '拒绝之后源文件必须完好');
});

await check('office.files.copy：同一文件按平台的大小写口径判定（Windows 不区分大小写）', async () => {
    if (process.platform !== 'win32') {
        return '非 Windows：路径大小写敏感，这一条不适用';
    }
    const result = await run(`return office.files.copy('logo.png', 'LOGO.PNG', { overwrite: true });`);
    assert.equal(result.ok, false, 'Windows 上 LOGO.PNG 与 logo.png 是同一个文件，必须拒绝');
    assert.match(result.error.message, /源与目标是同一个文件/);
    assert.deepEqual(read('logo.png'), PNG_BYTES);
});

await check('内核记账：copy 把源记成读过的输入，move 不记（源路径已不存在）', () => {
    const env = testEnv(join(root, 'tracking'));
    mkdirSync(dirname(join(root, 'tracking', 'a.bin')), { recursive: true });
    writeFileSync(join(root, 'tracking', 'a.bin'), PNG_BYTES);
    env.copyFile('a.bin', 'b.bin');
    assert.equal(env.reads.some((item) => item.path === 'a.bin'), true, '复制：源文件是这份产物的来源');
    assert.equal(env.writes.some((item) => item.path === 'b.bin'), true);
    env.moveFile('b.bin', 'c.bin');
    assert.equal(env.reads.some((item) => item.path === 'b.bin'), false, '移动：源路径移动后已不存在，记进 source 会指向查不到的文件');
    assert.equal(env.writes.some((item) => item.path === 'c.bin'), true);
});

/** 台账原文（JSONL）里的那条记录：`memory.read` 的投影不带 source，要验 source 得读原文。 */
function ledgerRecord(path) {
    const text = readFileSync(memoryPaths(root, config.memory).ledger, 'utf8');
    const records = text.trim().split('\n').filter((line) => line.trim() !== '').map((line) => JSON.parse(line));
    return records.find((record) => record.path === path);
}

await check('全链路：复制一份 docx，副本照样被复检并进台账', async () => {
    const origin = await run(`
        const wb = office.word.create({ path: '原件.docx', title: '原件' });
        wb.para('内容一行。');
        wb.save();
    `);
    assert.equal(origin.ok, true, origin.error?.message);

    // 第二次调用：复核一份**上一次**写出的 docx。这样「原件」才是这次真正的输入
    // （同一次调用里写出来的文件会被 source 排除 —— 原地 revise 读自己属于同义反复）。
    const result = await run(`office.files.copy('原件.docx', '副本.docx');`);
    assert.equal(result.ok, true, result.error?.message);
    assert.deepEqual(result.files.map((file) => file.path), ['副本.docx']);
    const copy = result.files[0];
    assert.equal(copy.format, 'word', '副本按扩展名走同一个格式复检器');
    assert.notEqual(copy.stats, null, '副本要有真实统计，不能只是「写了一个文件」');
    assert.equal(result.memory.logged, 1, '副本登记一条台账');

    const record = ledgerRecord('副本.docx');
    assert.ok(record !== undefined, '副本要在台账里');
    assert.match(JSON.stringify(record.source), /原件\.docx/, 'source 要指向被复制的那一份');

    // 投影侧也要能看到它（检索串里带 source）。
    const memory = createMemory({ root, memory: config.memory });
    const found = await memory.read({ layer: 'ledger', query: '原件' });
    assert.equal(found.ledger.items.some((item) => item.path === '副本.docx'), true, '按来源文件检索要能找到副本');
});

// ── 4. scriptFile + args（跨调用复用）与文档自同步 ─────────────────────────

await check('office_run(scriptFile)：读工作目录里的 .js 跑，args 以全局 args 传进去', async () => {
    writeFileSync(join(root, 'build.js'), [
        "office.files.write('report-' + args.name + '.md', '# ' + args.name);",
        'return { name: args.name, hasOffice: typeof office, hasArgs: typeof args };',
    ].join('\n'), 'utf8');
    const result = await runScript({ scriptFile: 'build.js', args: { name: 'q1' }, purpose: '复用脚本' });
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.name, 'q1');
    assert.equal(result.returned.hasOffice, 'object');
    assert.equal(result.scriptFile, 'build.js');
    assert.equal(read('report-q1.md').toString('utf8'), '# q1');
    assert.match(renderRun(result), /脚本来源：build\.js/);

    // 同一份脚本换个参数再来一次 —— 这就是「跨调用复用」的全部意思。
    const second = await runScript({ scriptFile: 'build.js', args: { name: 'q2' } });
    assert.equal(second.ok, true, second.error?.message);
    assert.equal(read('report-q2.md').toString('utf8'), '# q2');
});

await check('office_run(scriptFile)：错误行号是那份脚本文件里的行号', async () => {
    writeFileSync(join(root, 'broken.js'), [
        'const ok = 1;',
        '// 注释一行',
        'office.assert(ok === 2, "这一行该失败");',
    ].join('\n'), 'utf8');
    const result = await runScript({ scriptFile: 'broken.js' });
    assert.equal(result.ok, false);
    assert.equal(result.error.line, 3, `期望第 3 行，实际 ${String(result.error.line)}`);
    assert.equal(result.scriptFile, 'broken.js');
    const text = renderRun(result);
    assert.match(text, /office_run 执行失败（broken\.js）/);
    assert.match(text, /脚本第 3 行/);
});

await check('office_run(args)：非对象 / 循环引用都落一条警告，不再静默变成 {}', async () => {
    // 静默变 {} 会让脚本按 args.x 取到 undefined 却毫无信号（独立复核的一条）。
    const array = await runScript({ script: 'return Object.keys(args).length;', args: [1, 2] });
    assert.equal(array.ok, true, array.error?.message);
    assert.equal(array.returned, 0);
    assert.ok(array.warnings.some((line) => /args 不是对象（数组）/.test(line)), `要落警告：${JSON.stringify(array.warnings)}`);
    const string = await runScript({ script: 'return Object.keys(args).length;', args: 'x' });
    assert.ok(string.warnings.some((line) => /args 不是对象（string）/.test(line)));
    const circular = {};
    circular.self = circular;
    const bad = await runScript({ script: 'return Object.keys(args).length;', args: circular });
    assert.equal(bad.ok, true, '不可序列化的 args 不该把这次调用打挂');
    assert.ok(bad.warnings.some((line) => /args 无法序列化/.test(line)), `要落警告：${JSON.stringify(bad.warnings)}`);
    // 正向控制：正常对象不许有这些警告。
    const good = await runScript({ script: 'return args.a;', args: { a: 7 } });
    assert.equal(good.returned, 7);
    assert.equal(good.warnings.some((line) => /args/.test(line)), false);
});

await check('office_run(args)：注入的是 vm realm 的对象，拿不到宿主 process 的把手', async () => {
    // `codeGeneration:{strings:false}` 只挡 vm 自己那套 eval / new Function；
    // 直接把宿主对象塞进沙箱等于留下 `args.constructor.constructor` 这条构造器链。
    const result = await runScript({
        script: "try { args.constructor.constructor('return process')().version; return 'escaped'; }"
            + " catch (error) { return 'blocked:' + error.name; }",
        args: { a: 1 },
    });
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned, 'blocked:EvalError', `args 必须在 vm realm 里重建：${String(result.returned)}`);
});

await check('office_run：超时时说清「无法定位行号」，不留下一个空位置', async () => {
    const result = await executeRun({ script: 'while (true) {}' }, exec, resolveConfig({ scriptTimeoutMs: 1_000 }));
    assert.equal(result.ok, false);
    assert.equal(result.error.line, null);
    const text = renderRun(result);
    assert.match(text, /脚本超时被中断，无法定位到具体行/);
    assert.doesNotMatch(text, /第 null 行/);
});

await check('office.files.write：content 是 undefined 时给可读的报错（不是宿主内部的 Buffer 报错）', async () => {
    const result = await run(`office.files.write('u.txt', undefined); return 1;`);
    assert.equal(result.ok, false);
    assert.match(result.error.message, /content 不能是 undefined/);
    assert.doesNotMatch(result.error.message, /first argument must be of type/);
    assert.equal(existsSync(join(root, 'u.txt')), false, '拒绝时不写半截文件');
});

await check('office.files.write：写不出文本的形状（函数 / Symbol / toJSON 返回 undefined）也报可读错误', async () => {
    // 独立复核打穿的一条：`undefined` 守卫只挡了实参本身，`JSON.stringify` 结果为
    // `undefined` 的三种形状仍然直通宿主内部的 `Buffer.from` 报错。
    for (const expression of ['() => 1', "Symbol('x')", '{ toJSON() { return undefined; } }']) {
        const result = await run(`office.files.write('shape.txt', ${expression}); return 1;`);
        assert.equal(result.ok, false, `${expression} 必须被拒绝`);
        assert.match(result.error.message, /content 写不出文本/);
        assert.doesNotMatch(result.error.message, /first argument must be of type/);
    }
    assert.equal(existsSync(join(root, 'shape.txt')), false, '拒绝时不写半截文件');
});

await check('office.files.write：Map / Set / BigInt / 循环引用都给可读报错（不许静默写成 {}）', async () => {
    const map = await run(`office.files.write('m.json', new Map([['i', new Uint8Array([0, 255])]])); return 1;`);
    assert.equal(map.ok, false);
    assert.match(map.error.message, /content 是 Map/);
    assert.match(map.error.message, /Array\.from/);
    const set = await run(`office.files.write('s.json', new Set([1])); return 1;`);
    assert.equal(set.ok, false);
    assert.match(set.error.message, /content 是 Set/);
    const bigint = await run(`office.files.write('b.json', 1n); return 1;`);
    assert.equal(bigint.ok, false);
    assert.match(bigint.error.message, /无法写成 JSON/);
    const circular = await run(`const a = {}; a.self = a; office.files.write('c.json', a); return 1;`);
    assert.equal(circular.ok, false);
    assert.match(circular.error.message, /无法写成 JSON/);
    for (const name of ['m.json', 's.json', 'b.json', 'c.json']) {
        assert.equal(existsSync(join(root, name)), false, `${name} 不该被写出来`);
    }
});

await check('office.files.readBytes → write：子区间 / DataView / ArrayBuffer 三种形状都按字节写', async () => {
    const result = await run(`
        const raw = office.files.readBytes('logo.png');
        office.files.write('slice.bin', raw.subarray(1, 3));
        const view = new DataView(new ArrayBuffer(2));
        view.setUint8(0, 65); view.setUint8(1, 66);
        office.files.write('view.bin', view);
        office.files.write('buffer.bin', new Uint8Array([7, 8]).buffer);
        return [office.files.stat('slice.bin').bytes, office.files.stat('view.bin').bytes, office.files.stat('buffer.bin').bytes];
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.deepEqual(result.returned, [2, 2, 2]);
    assert.deepEqual([...read('slice.bin')], [...PNG_BYTES.subarray(1, 3)]);
    assert.deepEqual([...read('view.bin')], [65, 66]);
    assert.deepEqual([...read('buffer.bin')], [7, 8]);
});

await check('office.files.write：嵌在对象里的二进制要响亮失败，不许静默变成数字键', async () => {
    // 独立复核打穿的一条：顶层不是字节时整份走 JSON.stringify，`{img: Uint8Array}` 会变成
    // `{"img":{"0":0,"1":255}}` —— 落盘「成功」、没有告警，而文档写着「字节不会被 JSON 化」。
    const nested = await run(`office.files.write('nested.json', { img: new Uint8Array([0, 255]) }); return 1;`);
    assert.equal(nested.ok, false, '嵌套二进制必须报错，而不是写出一份数字键的坏文件');
    assert.match(nested.error.message, /img 字段/);
    assert.match(nested.error.message, /二进制/);
    assert.equal(existsSync(join(root, 'nested.json')), false, '拒绝时不写半截文件');

    const shared = await run(`office.files.write('shared.json', { buf: new SharedArrayBuffer(4) }); return 1;`);
    assert.equal(shared.ok, false, 'SharedArrayBuffer 同样不许静默变 {}');
    assert.match(shared.error.message, /SharedArrayBuffer/);

    // 顶层字节仍然照写（这条是正向控制：上面的拒绝不能把正路也堵了）。
    const top = await run(`office.files.write('top.bin', new Uint8Array([3, 4])); return office.files.stat('top.bin').bytes;`);
    assert.equal(top.ok, true, top.error?.message);
    assert.deepEqual([...read('top.bin')], [3, 4]);
});

await check('office_run(scriptFile)：语法错误的行号也拿得到（编译期栈帧只有「文件:行」）', async () => {
    // 运行期错误的栈帧是 `文件:行:列`，编译期（SyntaxError）只有 `文件:行`。
    // 只认带列号的那种写法会让「脚本文件里写错一个字符」退化成「脚本第 null 行」。
    writeFileSync(join(root, 'syntax.js'), [
        'const a = 1;',
        'const b = ;',
        'return b;',
    ].join('\n'), 'utf8');
    const result = await runScript({ scriptFile: 'syntax.js' });
    assert.equal(result.ok, false);
    assert.equal(result.error.name, 'SyntaxError');
    assert.equal(result.error.line, 2, `语法错误也要给行号，实际 ${String(result.error.line)}`);
    assert.match(renderRun(result), /脚本第 2 行/);
    assert.doesNotMatch(renderRun(result), /第 null 行/);
});

await check('office_run(scriptFile)：指向二进制文件时给可读的报错，不是一句 Invalid token', async () => {
    writeFileSync(join(root, 'not-a-script.js'), Buffer.from([0x00, 0x01, 0x02, 0xff]));
    const result = await runScript({ scriptFile: 'not-a-script.js' }).catch((error) => ({ ok: false, error }));
    const message = result.error?.message ?? '';
    assert.equal(result.ok, false);
    assert.match(message, /看着是二进制文件/);
    assert.doesNotMatch(message, /Invalid or unexpected token/);
});

await check('office_run：script 与 scriptFile 的三种坏组合都给可读的报错', async () => {
    await assert.rejects(() => runScript({ script: 'return 1;', scriptFile: 'build.js' }), /只能给一个/);
    await assert.rejects(() => runScript({}), /script 或 scriptFile 必须给一个/);
    await assert.rejects(() => runScript({ scriptFile: '不存在的.js' }), /读不到脚本文件/);
    assert.match(
        await runScript({ scriptFile: '不存在的.js' }).catch((error) => error.message),
        /office\.files\.write/,
        '读不到脚本时要说清怎么落一份脚本',
    );
    writeFileSync(join(root, 'empty.js'), '   \n');
    await assert.rejects(() => runScript({ scriptFile: 'empty.js' }), /是空的/);
    writeFileSync(join(root, 'huge.js'), `return '${'x'.repeat(2000)}';`);
    await assert.rejects(
        () => executeRun({ scriptFile: 'huge.js' }, exec, resolveConfig({ maxScriptChars: 1000 })),
        /脚本文件 huge\.js 太长/,
    );
});

await check('office_run：脚本文件本身记进台账的 source', async () => {
    await assert.rejects(() => runScript({ scriptFile: 'make-doc.js' }), /读不到脚本文件/);
    writeFileSync(join(root, 'make-doc.js'), [
        "const wb = office.word.create({ path: '由脚本构建.docx', title: '由脚本构建' });",
        "wb.para('内容。');",
        'wb.save();',
    ].join('\n'), 'utf8');
    const built = await runScript({ scriptFile: 'make-doc.js' });
    assert.equal(built.ok, true, built.error?.message);
    const record = ledgerRecord('由脚本构建.docx');
    assert.ok(record !== undefined, '产物要在台账里');
    assert.match(JSON.stringify(record.source), /make-doc\.js/,
        'source 要能回答「这份产物是拿哪段脚本构建的」');
});

await check('office_run：返回二进制时回显成字节数，不铺开成几十个数字键', async () => {
    const result = await run(`return { data: office.files.readBytes('logo.png'), n: 1 };`);
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.data, `[二进制 ${PNG_BYTES.length} 字节]`);
    // 「回显必须仍是无损 JSON」不是一句空话：真出问题的是**宿主对象漏进返回值**
    // （Buffer / Uint8Array 在 JSON 化后变成数字键）。所以这里逐层扫，不许有二进制视图。
    const walk = (value, path) => {
        if (ArrayBuffer.isView(value)) assert.fail(`${path} 漏了二进制视图进返回值`);
        if (value === null || typeof value !== 'object') return;
        for (const [key, item] of Object.entries(value)) walk(item, `${path}.${key}`);
    };
    walk(result, 'result');
    assert.deepEqual(JSON.parse(JSON.stringify(result)), result, '回显必须仍是无损 JSON');
    assert.match(renderRun(result), new RegExp(`\\[二进制 ${PNG_BYTES.length} 字节\\]`));
});

await check('回显上限自同步：文档里的数字从实现现取，不是手写的 4000', async () => {
    // 新-13 的第一处漂移：docs.js 两处手写「超过 4000 字符会被截断」，而实现是 1200。
    // 现在两边都引用 config.js 的 RETURN_ECHO_CHARS，这一条同时守行为与文案。
    const compact = (await buildHelp('run')).text;
    const full = (await buildHelp('run', { detail: true })).text;
    assert.match(compact, new RegExp(`超过 ${RETURN_ECHO_CHARS} 字符会截断`), '默认层要写实现里的真值');
    assert.match(full, new RegExp(`超过 ${RETURN_ECHO_CHARS} 字符会被截断`), '全文层要写实现里的真值');
    assert.doesNotMatch(`${compact}\n${full}`, /超过 4000 字符/, '手写的 4000 是这一条账的原始症状');

    // 行为侧：回显真的按这个数字截断（改常量而文档不动，上面两条就红）。
    const result = await run(`return 'x'.repeat(${RETURN_ECHO_CHARS + 200});`);
    assert.equal(result.ok, true, result.error?.message);
    const echoed = /脚本返回：([\s\S]*)/.exec(renderRun(result))?.[1] ?? '';
    // 回显的是 JSON 字符串，开头那个引号也占一位，所以 x 的个数比上限少 1。
    const shown = echoed.split('…')[0].replace(/"/g, '').length;
    assert.equal(shown, RETURN_ECHO_CHARS - 1, `回显应当正好截到 ${RETURN_ECHO_CHARS} 个字符（含 JSON 引号）`);
});

await check('新-13 第二处：README 说的拷贝动作真的存在（readBytes / copy / move）', async () => {
    const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
    for (const action of ['office.files.readBytes', 'office.files.copy', 'office.files.move']) {
        assert.ok(readme.includes(action), `README 里要写清 ${action} 的用法（旧账里的漂移就是「教了一个不存在的动作」）`);
    }
    // 反向：数字只许有一份来源。爬注释那种断言换个措辞就通过（复核指出过），
    // 所以这里查的是**手写数字本身**：docs.js 里不许再出现硬编码的「超过 4000 字符」。
    const docs = readFileSync(new URL('../src/docs.js', import.meta.url), 'utf8');
    assert.ok(!/超过\s*4000/.test(docs), 'docs.js 里不许再手写「超过 4000」——那个数字要从 RETURN_ECHO_CHARS 现取');
    assert.ok(/RETURN_ECHO_CHARS/.test(docs), 'docs.js 要真的引用 config.js 的常量');
});

// ── 汇总 ──────────────────────────────────────────────────────────────────

rmSync(root, { recursive: true, force: true });

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok && item.note !== undefined ? `  → ${item.note}` : ''}`);
    if (!item.ok) console.log(`       | ${item.error?.message ?? item.error}`);
}
console.log(`files-channel: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);
