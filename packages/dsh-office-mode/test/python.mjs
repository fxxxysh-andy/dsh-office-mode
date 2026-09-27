/**
 * Python 计算与绘图通道（office.python）的测试。
 *
 * 钉住的是四件事：
 *   1. 面在：office.python.run / file / check 挂在 SDK 上，设置能关掉它；
 *   2. 真跑得起来：解释器探测、stdout / stderr、退出码、超时、回溯行号不漂移；
 *   3. 编码与后端：中文输出不乱码（Windows 默认 GBK 会当场 UnicodeEncodeError）、
 *      matplotlib 无显示器也能出图；
 *   4. 产物归置：只报这一次跑出来的文件、只落缓存目录、且被算进 office_run
 *      的产物统计（否则「只画了一张图」会被误报成漏了 save()）。
 *
 * 解释器与科学包是外部依赖：本机没有 Python 时相关断言标记 SKIP 并如实说明，
 * 不当成通过（与 test/pdf.mjs 对渲染引擎的处理一致）。
 *
 * 跑法：node test/python.mjs
 */
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import { probePython, pythonCheck, resetPythonProbe } from '../src/python.js';
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

/** 把一段 Python 代码塞进 office_run 脚本里跑。 */
function runScript(runRoot, code, options = {}, configOverrides = {}) {
    const script = [
        `const result = await office.python.run(${JSON.stringify(code)}, ${JSON.stringify(options)});`,
        'return result;',
    ].join('\n');
    return executeRun({ script }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig(configOverrides));
}

/**
 * 清理测试目录。
 *
 * 不能直接 rmSync：Windows 上被 SIGKILL 掉的子进程可能还占着目录（超时用例就是
 * 这种情况），清理失败是环境现象，不该让断言变成红。
 */
function cleanup(dir) {
    try {
        rmSync(dir, { recursive: true, force: true });
    } catch {
        // 目录被占住时留给系统临时目录清理
    }
}

const python = await probePython('');
const hasPython = python !== undefined;
const hasMatplotlib = hasPython && python.packages?.matplotlib !== null && python.packages?.matplotlib !== undefined;
const note = hasPython ? `解释器 ${python.bin}（Python ${python.version}）` : '本机没有可用的 Python 解释器';

// ── 配置与探测 ────────────────────────────────────────────────────────────

await check('config：Python 通道默认开启，解释器自动探测、产物落 python/out', async () => {
    const config = resolveConfig({});
    assert.equal(config.python.enabled, true);
    assert.equal(config.python.bin, '');
    assert.equal(config.python.outDir, 'python/out');
    assert.equal(config.python.timeoutMs, 120_000);
    const off = resolveConfig({ python: { enabled: false, timeoutMs: 5_000 } });
    assert.equal(off.python.enabled, false);
    assert.equal(off.python.timeoutMs, 5_000);
});

await check('probePython：找得到解释器并读出版本', async () => {
    if (!hasPython) skip(note);
    assert.match(python.version, /^\d+\.\d+/);
    return note;
});

await check('probePython：不存在的解释器返回 undefined（不假装可用）', async () => {
    resetPythonProbe();
    const missing = await probePython('office-mode-no-such-python-xyz');
    resetPythonProbe();
    assert.equal(missing, undefined);
});

await check('office.python.check：包清单与解释器一次问清', async () => {
    if (!hasPython) skip(note);
    const info = await pythonCheck({ config: resolveConfig({}) });
    assert.equal(info.available, true);
    assert.equal(typeof info.version, 'string');
    assert.equal(typeof info.packages, 'object');
    assert.equal(Array.isArray(info.missingPackages), true);
    if (hasMatplotlib) assert.match(String(info.packages.matplotlib), /^\d+/);
    return Object.keys(info.packages).join(' / ') || '（没有装常用科学包）';
});

// ── 真跑 ──────────────────────────────────────────────────────────────────

await check('office.python：纯计算能跑，返回值与 stdout 都拿得到', async () => {
    if (!hasPython) skip(note);
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-calc-'));
    const code = [
        'total = sum(i * i for i in range(10))',
        'print("平方和：", total)',
        'print("done")',
    ].join('\n');
    const result = await runScript(runRoot, code);
    assert.equal(result.ok, true, JSON.stringify(result.error));
    assert.equal(result.returned.ok, true, result.returned.stderr);
    assert.equal(result.returned.code, 0);
    assert.match(result.returned.stdout, /平方和： 285/);
    assert.equal(result.returned.files.length, 0);
    cleanup(runRoot);
});

await check('office.python：中文 stdout 不乱码（Windows 默认 GBK 会当场炸）', async () => {
    if (!hasPython) skip(note);
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-utf8-'));
    const result = await runScript(runRoot, 'print("中文输出：", "—·—", 42)');
    assert.equal(result.returned.ok, true, result.returned.stderr);
    assert.match(result.returned.stdout, /中文输出： —·— 42/);
    assert.equal(result.returned.stderr, '');
    cleanup(runRoot);
});

await check('office.python：脚本报错时 ok:false，回溯行号就是代码里的行号', async () => {
    if (!hasPython) skip(note);
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-err-'));
    const code = [
        'value = 1',
        'other = 0',
        'print(value / other)',
    ].join('\n');
    const result = await runScript(runRoot, code);
    const returned = result.returned;
    assert.equal(result.ok, true, 'office_run 本身不该失败：脚本报错是脚本的事');
    assert.equal(returned.ok, false);
    assert.notEqual(returned.code, 0);
    assert.match(returned.stderr, /ZeroDivisionError/);
    // 行号不漂移：驱动脚本用 runpy，不在用户代码前面拼任何前导。
    assert.match(returned.stderr, /line 3/);
    cleanup(runRoot);
});

await check('office.python：超时会杀掉进程并如实报 timedOut', async () => {
    if (!hasPython) skip(note);
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-timeout-'));
    const result = await runScript(runRoot, 'import time\ntime.sleep(30)', { timeoutMs: 2_000 });
    const returned = result.returned;
    assert.equal(returned.ok, false);
    assert.equal(returned.timedOut, true);
    assert.match(String(returned.error ?? ''), /没有结束/);
    cleanup(runRoot);
});

await check('office.python：超长 stdout 会截断，完整内容留在日志文件里', async () => {
    if (!hasPython) skip(note);
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-long-'));
    const result = await runScript(runRoot, 'print("x" * 9000)');
    const returned = result.returned;
    assert.equal(returned.ok, true, returned.stderr);
    assert.equal(returned.stdoutTruncated, true);
    assert.match(returned.stdout, /已截断/);
    const full = readFileSync(join(runRoot, returned.logFiles.stdout), 'utf8');
    assert.equal(full.trim().length, 9000);
    cleanup(runRoot);
});

await check('office.python.file：跑工作目录里的 .py，args 进 sys.argv', async () => {
    if (!hasPython) skip(note);
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-file-'));
    writeFileSync(join(runRoot, 'report.py'), [
        'import sys',
        'print("参数：", sys.argv[1:])',
    ].join('\n'), 'utf8');
    const result = await executeRun({
        script: 'return await office.python.file("report.py", { args: ["2026", "三季度"] });',
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    assert.equal(result.returned.ok, true, result.returned.stderr);
    assert.match(result.returned.stdout, /参数： \['2026', '三季度'\]/);
    cleanup(runRoot);
});

// ── 绘图与产物 ────────────────────────────────────────────────────────────

await check('office.python：matplotlib 出图，产物落缓存并标成 image', async () => {
    if (!hasPython) skip(note);
    if (!hasMatplotlib) skip('本机解释器里没有 matplotlib');
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-plot-'));
    const code = [
        'import matplotlib.pyplot as plt',
        'picked = use_cjk_font()',
        'fig, ax = plt.subplots(figsize=(6, 3.6), dpi=120)',
        'ax.plot([1, 2, 3], [4, 5, 9], marker="o")',
        'ax.set_title("月度趋势")',
        'fig.savefig(OUT_DIR + "/trend.png", bbox_inches="tight")',
        'print("font:", picked)',
    ].join('\n');
    const result = await executeRun({
        script: [
            `const run = await office.python.run(${JSON.stringify(code)});`,
            'return { ...run, artifacts: office.cache.stats().artifacts };',
        ].join('\n'),
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    const returned = result.returned;
    assert.equal(returned.ok, true, returned.stderr);
    assert.match(returned.stdout, /font: (Microsoft YaHei|SimHei|Noto Sans CJK SC|SimSun|None)/);
    assert.equal(returned.files.length, 1, JSON.stringify(returned.files));
    const file = returned.files[0];
    assert.equal(file.kind, 'image');
    assert.match(file.path, /^\.office\/cache\/python\/out\/trend\.png$/);
    // 真图，不是 0 字节的壳：PNG 魔数 + 合理体积。
    const absolute = join(runRoot, file.path);
    const bytes = readFileSync(absolute);
    assert.equal(bytes.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
    assert.equal(statSync(absolute).size > 2_000, true, `图太小：${file.bytes} 字节`);
    // 产物必须记进缓存统计，否则 office_run 会报「没有写出任何文件」。
    assert.equal(returned.artifacts, 1, JSON.stringify(result.cache));
    assert.equal(result.warnings.some((line) => line.includes('没有写出任何文件')), false, JSON.stringify(result.warnings));
    cleanup(runRoot);
});

await check('office.python：产物只报这一次跑出来的（上一次的图不算）', async () => {
    if (!hasPython) skip(note);
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-diff-'));
    const first = await runScript(runRoot, 'open(OUT_DIR + "/a.txt", "w", encoding="utf-8").write("a")');
    assert.deepEqual(first.returned.files.map((file) => file.path), ['.office/cache/python/out/a.txt']);
    const second = await runScript(runRoot, 'open(OUT_DIR + "/b.txt", "w", encoding="utf-8").write("b")');
    assert.deepEqual(second.returned.files.map((file) => file.path), ['.office/cache/python/out/b.txt']);
    cleanup(runRoot);
});

// ── 关闭与边界 ────────────────────────────────────────────────────────────

await check('设置：关掉 Python 通道后调用报错，且说明在哪儿打开', async () => {
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-off-'));
    const result = await runScript(runRoot, 'print(1)', {}, { python: { enabled: false } });
    assert.equal(result.ok, false);
    assert.match(String(result.error?.message ?? ''), /已在本插件的设置里关闭/);
    cleanup(runRoot);
});

await check('office.python：空代码与过长的代码都明确拒绝', async () => {
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-guard-'));
    const empty = await runScript(runRoot, '   ');
    assert.equal(empty.ok, false);
    assert.match(String(empty.error?.message ?? ''), /code 不能为空/);
    const huge = await runScript(runRoot, 'x = 1\n'.repeat(120_000), {}, { maxScriptChars: 1_000_000 });
    assert.equal(huge.ok, false);
    assert.match(String(huge.error?.message ?? ''), /代码太长/);
    cleanup(runRoot);
});

await check('office_run：office.python 挂在 SDK 上（面不能静默丢）', async () => {
    const runRoot = mkdtempSync(join(tmpdir(), 'office-py-sdk-'));
    const result = await executeRun({
        script: 'return { run: typeof office.python.run, file: typeof office.python.file, check: typeof office.python.check };',
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    assert.deepEqual(result.returned, { run: 'function', file: 'function', check: 'function' });
    cleanup(runRoot);
});

// ── 文档 ──────────────────────────────────────────────────────────────────

await check('office_help：python 话题与索引都提到 OUT_DIR / savefig', async () => {
    const { buildHelp } = await import('../src/docs.js');
    const python = await buildHelp('python');
    assert.equal(python.topic, 'python');
    assert.match(python.text, /office\.python\.run/);
    assert.match(python.text, /OUT_DIR/);
    assert.match(python.text, /savefig/);
    assert.match(python.text, /use_cjk_font/);
    const index = await buildHelp('');
    assert.match(index.text, /python：Python 计算与绘图/);
    const run = await buildHelp('run');
    assert.match(run.text, /office\.python/);
    const guide = await buildHelp('guide');
    assert.match(guide.text, /office\.python\.run/);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const noteText = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${noteText}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`python: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
