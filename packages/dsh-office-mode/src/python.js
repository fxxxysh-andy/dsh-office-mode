/**
 * Python 计算与绘图通道（office.python）。
 *
 * 办公模式的工具面是「极简」的：命令行、后台任务、子代理这些编程入口一律关掉。
 * 代价是**科学计算与绘图没有出口** —— 一张要嵌进报告的折线图，只能靠手工画
 * 形状，或者让用户自己算完再放进来。这一层补的就是那一小块：只跑 Python、
 * 只跑一次、产物只落缓存目录，不打开 shell，也不给脚本任何命令行入口。
 *
 * 三件事：
 *   office.python.check()            本机有没有 Python、装了哪些科学包（一次探测）
 *   office.python.run(code, opts)    跑一段代码；stdout/stderr 与产物文件都带回来
 *   office.python.file(path, opts)   跑工作目录里已有的 .py（长脚本先写文件再跑）
 *
 * 几条刻意的设计：
 *   - **不用管道**：受限沙箱下 Node 的 `stdio:'pipe'` 会 `spawn EPERM`，
 *     所以复用 pdf.js 的 runPdfProcess（stdout/stderr 重定向到文件句柄）。
 *   - **UTF-8 强制**：中文 Windows 的默认代码页是 GBK，重定向后的 stdout
 *     按 GBK 编码会当场 UnicodeEncodeError。子进程统一带 PYTHONUTF8=1 与
 *     PYTHONIOENCODING=utf-8。
 *   - **MPLBACKEND=Agg**：没有显示器的环境里 matplotlib 默认后端会报错，
 *     这里直接把后端钉成 Agg，脚本里不必自己写 matplotlib.use()。
 *   - **产物只认缓存目录**：脚本写出的图/表按修改时间收集，落在 .office/cache
 *     下（跨调用保留、按 TTL 清理）。要当交付物的文本仍走 office.files.write。
 *   - **行号不漂移**：脚本原样落盘、用 runpy 执行，异常回溯里的行号就是用户
 *     写的行号 —— 不在前面拼任何前导代码。
 *
 * @module dsh-office-mode/python
 */
import { existsSync, mkdirSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { displayPath } from './engine/kit.js';
// 复用 PDF 那一套「不用管道」的进程执行器：受限沙箱下 stdio:'pipe' 会 EPERM，
// 而这里跑的是完全相同的形态（外部解释器 + 重定向到文件句柄）。
// 与 tex.js 的做法一致：一个实现，两处调用。
import { findPdfExecutable as findExecutable, runPdfProcess as runProcess } from './pdf.js';

/** 产物目录（缓存目录内，相对路径）。 */
export const DEFAULT_PYTHON_OUT_DIR = 'python/out';
/** 单次运行的默认超时：科学计算与绘图比读一份 PDF 慢，给两分钟。 */
export const DEFAULT_PYTHON_TIMEOUT_MS = 120_000;
/** 探测（找解释器、列科学包）的超时：探测卡住不该让整个调用陪等。 */
const PROBE_TIMEOUT_MS = 30_000;
/** 返回值里 stdout / stderr 的字符上限；完整内容留在日志文件里。 */
export const PYTHON_STDOUT_CHARS = 8_000;
export const PYTHON_STDERR_CHARS = 4_000;
/** 单段代码的字符上限（与 office_run 的脚本上限同一量级）。 */
export const PYTHON_MAX_CODE_CHARS = 400_000;

/** 解释器候选：显式配置优先，然后按常见命令名找。 */
const PYTHON_NAMES = ['python', 'python3', 'py'];

/**
 * 探测脚本：解释器版本 + 常用科学包。
 *
 * 一次进程拿到全部信息（每次 import 都要几百毫秒，逐个探太慢）。
 * 包缺失不是错误 —— 只报它不在，让模型据此决定用不用。
 */
const PROBE_SCRIPT = [
    'import importlib, json, sys',
    "names = ['numpy', 'scipy', 'pandas', 'matplotlib', 'sympy', 'PIL', 'sklearn', 'statsmodels']",
    'packages = {}',
    'for name in names:',
    '    try:',
    '        module = importlib.import_module(name)',
    '        packages[name] = str(getattr(module, "__version__", "?"))',
    '    except Exception:',
    '        packages[name] = None',
    'print(json.dumps({"version": sys.version.split()[0], "executable": sys.executable, "packages": packages}))',
].join('\n');

/**
 * 驱动脚本：把用户代码当 `__main__` 跑，并塞三个名字进它的全局命名空间。
 *
 * 为什么用 runpy 而不是在代码前面拼几行：拼前导会让所有回溯的行号偏移，
 * 而「脚本第几行出错」正是模型改代码唯一可靠的线索。
 * OUT_DIR / WORK_DIR 是给脚本用的路径；use_cjk_font() 只在被调用时才 import
 * matplotlib，所以纯计算脚本不必为它付导入代价。
 */
const DRIVER_SCRIPT = [
    'import os, runpy, sys',
    'script, outdir, workdir = sys.argv[1], sys.argv[2], sys.argv[3]',
    'sys.argv = [script] + sys.argv[4:]',
    'os.environ.setdefault("OFFICE_OUT_DIR", outdir)',
    'os.environ.setdefault("OFFICE_WORK_DIR", workdir)',
    'if workdir and workdir not in sys.path:',
    '    sys.path.insert(0, workdir)',
    '',
    'def use_cjk_font(name="Microsoft YaHei"):',
    '    """让 matplotlib 的中文标题与负号正常显示；不调用就不碰任何 rcParams。"""',
    '    import matplotlib',
    '    from matplotlib import font_manager',
    '    wanted = [name, "Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "SimSun"]',
    '    available = {font.name for font in font_manager.fontManager.ttflist}',
    '    picked = [item for item in wanted if item in available]',
    '    if not picked:',
    '        return None',
    '    matplotlib.rcParams["font.sans-serif"] = picked + ["DejaVu Sans"]',
    '    matplotlib.rcParams["axes.unicode_minus"] = False',
    '    return picked[0]',
    '',
    'runpy.run_path(script, run_name="__main__", init_globals={',
    '    "OUT_DIR": outdir,',
    '    "WORK_DIR": workdir,',
    '    "use_cjk_font": use_cjk_font,',
    '})',
].join('\n');

/** 图片类产物：能被 read_image 读、也能被 Word / PPT 直接嵌入。 */
const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg']);
/** 数据类产物：下一步多半要被读回来或者写进表格。 */
const DATA_EXTENSIONS = new Set(['.csv', '.tsv', '.json', '.txt', '.md', '.xlsx', '.xls', '.parquet', '.npy', '.npz']);

/** 子进程环境：UTF-8 + Agg 后端，其余继承宿主。 */
function pythonEnv(extra) {
    return {
        ...process.env,
        PYTHONUTF8: '1',
        PYTHONIOENCODING: 'utf-8',
        PYTHONUNBUFFERED: '1',
        MPLBACKEND: 'Agg',
        ...extra,
    };
}

/** 在一个临时目录里跑一次并拿回 stdout / stderr（探测用，不落工作区）。 */
async function captureRun(command, args, options = {}) {
    const stamp = `${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const outPath = join(tmpdir(), `office-python-${stamp}.out.txt`);
    const errPath = join(tmpdir(), `office-python-${stamp}.err.txt`);
    try {
        const result = await runProcess(command, args, {
            outPath,
            errPath,
            timeoutMs: options.timeoutMs ?? PROBE_TIMEOUT_MS,
            cwd: options.cwd,
            env: options.env,
        });
        return {
            code: result.code,
            timedOut: result.timedOut === true,
            error: result.error,
            out: typeof result.out === 'string' ? result.out : '',
            err: typeof result.err === 'string' ? result.err : '',
        };
    } finally {
        for (const file of [outPath, errPath]) {
            try {
                rmSync(file, { force: true });
            } catch {
                // 临时文件删不掉不影响结论
            }
        }
    }
}

/** 进程内缓存：同一个解释器只探一次。key 是「显式配置的路径或命令名」，空串表示自动。 */
const interpreterCache = new Map();

/**
 * 找一个真能用的 Python 解释器。
 *
 * 为什么要真跑一次：Windows 上 `python` 常常是应用商店的占位程序（跑起来弹商店、
 * 退出码非 0），PATH 里存在不等于能用。所以候选逐个试，第一个能打印版本号的才算数。
 *
 * @param {string} requested 配置里指定的解释器（绝对路径或命令名），留空自动探测
 * @returns {Promise<{bin: string, version: string, executable: string} | undefined>}
 */
export async function probePython(requested = '') {
    const key = String(requested ?? '');
    if (interpreterCache.has(key)) return interpreterCache.get(key);

    const candidates = [];
    if (key !== '') {
        // 显式配置：先按原样用（绝对路径或命令名），找不到再退回自动探测。
        candidates.push(key);
    } else {
        for (const name of PYTHON_NAMES) {
            const found = findExecutable([name]);
            if (found !== undefined) candidates.push(found);
        }
    }

    let resolved;
    for (const candidate of candidates) {
        const probe = await captureRun(candidate, ['-c', PROBE_SCRIPT], { timeoutMs: PROBE_TIMEOUT_MS });
        if (probe.code !== 0) continue;
        let parsed;
        try {
            parsed = JSON.parse(probe.out.trim().split('\n').pop() ?? '');
        } catch {
            continue;
        }
        if (typeof parsed?.version !== 'string') continue;
        resolved = {
            bin: candidate,
            version: parsed.version,
            executable: typeof parsed.executable === 'string' ? parsed.executable : candidate,
            packages: parsed.packages ?? {},
        };
        break;
    }
    interpreterCache.set(key, resolved);
    return resolved;
}

/** 探测结果按需刷新（设置里改了 python.bin 之后，下一次调用重新探）。 */
export function resetPythonProbe() {
    interpreterCache.clear();
}

/** 从配置里取 Python 通道的参数（配置缺失时用默认值，绝不因为没配置就不可用）。 */
function pythonSettings(config) {
    const raw = config?.python !== null && typeof config?.python === 'object' ? config.python : {};
    const timeout = Number(raw.timeoutMs);
    return {
        enabled: raw.enabled !== false,
        bin: typeof raw.bin === 'string' ? raw.bin.trim() : '',
        timeoutMs: Number.isFinite(timeout) && timeout > 0 ? timeout : DEFAULT_PYTHON_TIMEOUT_MS,
        outDir: typeof raw.outDir === 'string' && raw.outDir.trim() !== '' ? raw.outDir.trim() : DEFAULT_PYTHON_OUT_DIR,
    };
}

/** 解释器没找到时的报错：把「怎么让它能用」一次说清楚。 */
function missingPythonError(settings) {
    const where = settings.bin === '' ? 'PATH 里的 python / python3 / py' : `配置里的解释器「${settings.bin}」`;
    return new Error(
        `没有找到可用的 Python 解释器（找过 ${where}）。`
        + '装一个 Python 3 并确保 python 在 PATH 里即可，插件会自动探测；'
        + '也可以在本插件的设置里把「Python 解释器」填成绝对路径（如 C:\\Python313\\python.exe）。'
        + '注意 Windows 上应用商店的 python 占位程序不算可用。',
    );
}

/** 产物分类：图片 / 数据 / 其它。 */
function kindOf(filePath) {
    const at = filePath.lastIndexOf('.');
    const ext = at === -1 ? '' : filePath.slice(at).toLowerCase();
    if (IMAGE_EXTENSIONS.has(ext)) return 'image';
    if (DATA_EXTENSIONS.has(ext)) return 'data';
    return 'other';
}

/**
 * 跑之前先给产物目录拍一张快照（名字 → 体积:修改时间）。
 *
 * 为什么要快照而不是「只看修改时间晚于开始时刻」：文件系统时间戳的精度按卷不同
 * （网络盘、FAT 系能粗到 1-2 秒），只按时间比会把上一次刚写的图当成这一次的产物。
 * 前后对比则是精确的：新增的文件、或同名但体积/时间变了的文件，才算这一次跑出来的。
 */
function snapshotDir(dir) {
    const snapshot = new Map();
    const walk = (current, prefix) => {
        let entries = [];
        try {
            entries = readdirSync(current, { withFileTypes: true });
        } catch {
            return;
        }
        for (const entry of entries) {
            const full = join(current, entry.name);
            const name = prefix === '' ? entry.name : `${prefix}/${entry.name}`;
            if (entry.isDirectory()) {
                walk(full, name);
                continue;
            }
            try {
                const info = statSync(full);
                snapshot.set(name, `${info.size}:${Math.round(info.mtimeMs)}`);
            } catch {
                // 读不到的文件不参与对比
            }
        }
    };
    walk(dir, '');
    return snapshot;
}

/** 对比快照，收出这一次真正写出来的文件。 */
function diffOutputs(dir, before, limit = 200) {
    const out = [];
    const walk = (current, prefix) => {
        let entries = [];
        try {
            entries = readdirSync(current, { withFileTypes: true });
        } catch {
            return;
        }
        for (const entry of entries) {
            const full = join(current, entry.name);
            const name = prefix === '' ? entry.name : `${prefix}/${entry.name}`;
            if (entry.isDirectory()) {
                walk(full, name);
                continue;
            }
            let info;
            try {
                info = statSync(full);
            } catch {
                continue;
            }
            const mark = `${info.size}:${Math.round(info.mtimeMs)}`;
            if (before.get(name) === mark) continue;
            out.push({ name, absolute: full, bytes: info.size, modifiedMs: info.mtimeMs });
        }
    };
    walk(dir, '');
    return out.sort((a, b) => a.modifiedMs - b.modifiedMs).slice(0, limit);
}

/** 超长输出截断：完整内容在日志文件里，返回值里只留头部并标明截断。 */
function clampText(text, max) {
    if (typeof text !== 'string') return { text: '', truncated: false };
    if (text.length <= max) return { text, truncated: false };
    return { text: `${text.slice(0, max)}\n…（输出过长已截断）`, truncated: true };
}

/**
 * 跑一段 Python。
 *
 * @param {{code?: string, scriptPath?: string, options?: object, env: object, cache?: object, config?: object}} input
 * @returns {Promise<object>} 运行结果（脚本自己报错时也正常返回，`ok:false`）
 */
async function executePython(input) {
    const { env, cache, config } = input;
    const options = input.options !== null && typeof input.options === 'object' ? input.options : {};
    const settings = pythonSettings(config);
    if (!settings.enabled) {
        throw new Error('Python 通道已在本插件的设置里关闭（「Python 计算与绘图」→ 启用）。');
    }

    const root = resolve(env?.root ?? process.cwd());
    const code = typeof input.code === 'string' ? input.code : undefined;
    if (code !== undefined) {
        if (code.trim() === '') throw new Error('office.python.run：code 不能为空。');
        if (code.length > PYTHON_MAX_CODE_CHARS) {
            throw new Error(`office.python.run：代码太长（${code.length} 字符，上限 ${PYTHON_MAX_CODE_CHARS}）。请拆成多次或写进 .py 文件后用 office.python.file 跑。`);
        }
    }

    const python = await probePython(settings.bin);
    if (python === undefined) throw missingPythonError(settings);

    // 脚本与日志都进缓存目录（工作目录里不留任何临时文件）。
    const stamp = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
    const rawName = typeof options.name === 'string' && options.name.trim() !== '' ? options.name.trim() : `run-${stamp}`;
    const name = rawName.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-').replace(/^\.+/, '').slice(0, 80) || `run-${stamp}`;

    let scriptPath;
    if (code !== undefined) {
        scriptPath = cache !== undefined && cache !== null && typeof cache.path === 'function'
            ? cache.path(`python/${name}.py`)
            : join(tmpdir(), `office-python-${name}-${process.pid}.py`);
        writeFileSync(scriptPath, code, 'utf8');
    } else {
        const wanted = String(input.scriptPath ?? '').trim();
        if (wanted === '') throw new Error('office.python.file：要给出 .py 文件的路径。');
        scriptPath = resolve(root, wanted);
        if (!existsSync(scriptPath)) throw new Error(`office.python.file：找不到脚本 ${displayPath(root, scriptPath)}`);
    }

    const outDirOption = typeof options.outDir === 'string' && options.outDir.trim() !== '' ? options.outDir.trim() : settings.outDir;
    const outDir = cache !== undefined && cache !== null && typeof cache.ensureDir === 'function'
        ? cache.ensureDir(outDirOption)
        : join(tmpdir(), `office-python-out-${process.pid}`);
    mkdirSync(outDir, { recursive: true });

    const logBase = cache !== undefined && cache !== null && typeof cache.path === 'function'
        ? cache.path(`python/${name}`)
        : join(tmpdir(), `office-python-${name}-${process.pid}`);
    const outPath = `${logBase}.out.txt`;
    const errPath = `${logBase}.err.txt`;

    const timeoutRaw = Number(options.timeoutMs);
    const timeoutMs = Number.isFinite(timeoutRaw) && timeoutRaw > 0
        ? Math.min(900_000, Math.max(1_000, timeoutRaw))
        : settings.timeoutMs;
    const args = Array.isArray(options.args) ? options.args.map((item) => String(item)) : [];

    const started = Date.now();
    const before = snapshotDir(outDir);
    const result = await runProcess(python.bin, ['-c', DRIVER_SCRIPT, scriptPath, outDir, root, ...args], {
        outPath,
        errPath,
        cwd: root,
        timeoutMs,
        env: pythonEnv({ OFFICE_OUT_DIR: outDir, OFFICE_WORK_DIR: root }),
    });

    const outputs = diffOutputs(outDir, before);
    // 产物记进缓存统计：office_run 的「这次没写出任何文件」告警据此判断，
    // 否则「只画一张图」会被误报成漏了 save()。
    if (cache !== undefined && cache !== null && typeof cache.noteArtifact === 'function') {
        for (const _file of outputs) cache.noteArtifact();
    }

    const stdout = clampText(result.out, PYTHON_STDOUT_CHARS);
    const stderr = clampText(result.err, PYTHON_STDERR_CHARS);
    const rel = (absolute) => displayPath(root, absolute);

    return {
        ok: result.code === 0,
        code: result.code,
        timedOut: result.timedOut === true,
        error: typeof result.error === 'string' ? result.error : null,
        interpreter: python.bin,
        version: python.version,
        script: rel(scriptPath),
        outDir: rel(outDir),
        stdout: stdout.text,
        stderr: stderr.text,
        stdoutTruncated: stdout.truncated,
        stderrTruncated: stderr.truncated,
        logFiles: { stdout: rel(outPath), stderr: rel(errPath) },
        files: outputs.map((file) => ({
            path: rel(file.absolute),
            bytes: file.bytes,
            kind: kindOf(file.name),
        })),
        elapsedMs: Date.now() - started,
    };
}

/**
 * office.python.run(code, options)：跑一段代码。
 *
 * options: { name?, outDir?, timeoutMs?, args? }
 *   - name   这次运行的名字（脚本与日志的文件名，便于事后回看）
 *   - outDir 产物目录（缓存目录内的相对路径），默认 python/out
 *   - args   传给脚本的命令行参数（脚本里读 sys.argv[1:]）
 */
export function pythonRun(code, options, env, cache, config) {
    return executePython({ code, options, env, cache, config });
}

/** office.python.file(path, options)：跑工作目录里已有的 .py 文件。 */
export function pythonFile(filePath, options, env, cache, config) {
    const target = typeof filePath === 'string' ? filePath.trim() : '';
    if (target === '') throw new Error('office.python.file：要给出 .py 文件的路径。');
    return executePython({ scriptPath: target, options, env, cache, config });
}

/**
 * office.python.check()：本机 Python 能用到什么程度。
 *
 * 只探测、不执行任何用户代码。模型在写代码之前问一句，就能知道该不该用 numpy、
 * 该不该画图，而不是写完再撞 ModuleNotFoundError。
 */
export async function pythonCheck(options = {}) {
    const settings = pythonSettings(options.config);
    const python = await probePython(settings.bin);
    if (python === undefined) {
        return {
            available: false,
            enabled: settings.enabled,
            timeoutMs: settings.timeoutMs,
            outDir: settings.outDir,
            packages: {},
            hint: missingPythonError(settings).message,
        };
    }
    const installed = Object.entries(python.packages ?? {}).filter(([, version]) => version !== null && version !== undefined);
    const missing = Object.entries(python.packages ?? {}).filter(([, version]) => version === null || version === undefined).map(([name]) => name);
    return {
        available: true,
        enabled: settings.enabled,
        interpreter: python.bin,
        version: python.version,
        executable: python.executable,
        packages: Object.fromEntries(installed),
        missingPackages: missing,
        plotBackend: 'Agg（无显示器也能出图）',
        outDir: settings.outDir,
        timeoutMs: settings.timeoutMs,
        hint: installed.length === 0
            ? '这个解释器里没有任何常用科学包（numpy / matplotlib / pandas…）。装：python -m pip install numpy matplotlib'
            : `可用：${installed.map(([name, version]) => `${name} ${version}`).join(' / ')}`,
    };
}

export { PROBE_SCRIPT as PYTHON_PROBE_SCRIPT, DRIVER_SCRIPT as PYTHON_DRIVER_SCRIPT };
