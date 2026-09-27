/**
 * 批量执行引擎：一次工具调用跑完一整段脚本。
 *
 * 这里是「降低调用频率」的落点：模型把要做的所有事写进一个脚本，一次跑完；
 * 跑完后再把每个写出的 Office 文件重新打开解析一遍，把真实结构、体积、
 * 排版风险一起回给模型 —— 反馈来自复检，不是脚本的自述。
 *
 * @module dsh-office-mode/run
 */
import vm from 'node:vm';
import { createCache } from './engine/cache.js';
import { createEnv } from './engine/kit.js';
import { resolveTheme } from './engine/theme.js';
import { createMemory } from './memory.js';
import { projectDigest, sessionIdOf } from './projection.js';
import { formatByExtension, formatIds, loadFormat } from './registry.js';
import { buildSdk, isOfficeFile } from './sdk.js';

/** 包装脚本时插入的行数：`(async () => {` 占一行。 */
const WRAPPER_LINES = 1;

function safeJson(value, depth = 0) {
    if (value === undefined) return null;
    if (value === null) return null;
    const type = typeof value;
    if (type === 'string' || type === 'number' || type === 'boolean') {
        return type === 'number' && !Number.isFinite(value) ? String(value) : value;
    }
    if (type === 'bigint') return `${String(value)}n`;
    if (type === 'function') return '[function]';
    if (type === 'symbol') return String(value);
    if (depth >= 4) return '[深度截断]';
    if (Array.isArray(value)) {
        const head = Array.from(value.slice(0, 40), (item) => safeJson(item, depth + 1));
        if (value.length > 40) head.push(`…共 ${value.length} 项`);
        return head;
    }
    if (value instanceof Error) return { name: value.name, message: value.message };
    const out = {};
    let keys = 0;
    for (const key of Object.keys(value)) {
        if (keys >= 40) {
            out['…'] = '键过多已截断';
            break;
        }
        try {
            out[key] = safeJson(value[key], depth + 1);
        } catch {
            out[key] = '[无法序列化]';
        }
        keys += 1;
    }
    return out;
}

/** 把任意值收敛成可无损 JSON 的值：`undefined` 一律变 `null`。 */
function jsonSafe(value) {
    if (value === undefined) return null;
    if (value === null) return null;
    const type = typeof value;
    if (type === 'function' || type === 'symbol') return null;
    if (type !== 'object') return value;
    // 必须用 Array.from 而不是 value.map()：脚本跑在 vm 里，vm 数组的
    // map() 返回的仍是 vm realm 的数组，宿主侧 deepStrictEqual 会因为
    // 原型不同而判定不相等。Array.from 走的是宿主 Array，出来就是宿主数组。
    if (Array.isArray(value)) return Array.from(value, (item) => jsonSafe(item));
    if (value instanceof Error) return { name: value.name, message: value.message };
    const out = {};
    for (const [key, item] of Object.entries(value)) out[key] = jsonSafe(item);
    return out;
}

/**
 * 脚本抛出的错误可能是 vm realm 的实例：`error instanceof Error` 在宿主侧为假
 * （原型不同 realm），一旦按 `String(error)` 重新包装，行号与调用栈就全丢了，
 * 模型只能看到「脚本第 null 行」。所以按「有 stack 字符串」来判断，
 * 保留原始 name / message / stack。
 */
function isErrorLike(value) {
    if (value instanceof Error) return true;
    return value !== null && typeof value === 'object' && typeof value.stack === 'string';
}

/** 把 v8 的错误位置从包装后的行号换算回脚本自身的行号。 */
function cleanError(error) {
    const raw = isErrorLike(error) ? error : new Error(String(error));
    const stack = typeof raw.stack === 'string' ? raw.stack : '';
    const match = /office-script\.js:(\d+):(\d+)/.exec(stack);
    const line = match ? Math.max(1, Number.parseInt(match[1], 10) - WRAPPER_LINES) : undefined;
    const frames = stack
        .split('\n')
        .filter((frame) => frame.trim() !== '' && !frame.includes('node:vm'))
        .slice(0, 4)
        .map((frame) => frame.trim());
    return {
        name: typeof raw.name === 'string' && raw.name !== '' ? raw.name : 'Error',
        message: typeof raw.message === 'string' ? raw.message : String(raw),
        line,
        column: match ? Number.parseInt(match[2], 10) : undefined,
        stack: frames,
    };
}

/** 在受限上下文里执行脚本。没有 process / require / fs / 网络 / 定时器。 */
function runInVm(source, office, timeoutMs) {
    const sandbox = {
        office,
        console: {
            log: office.log,
            info: office.log,
            debug: office.log,
            warn: office.warn,
            error: office.warn,
        },
    };
    const context = vm.createContext(sandbox, {
        name: 'office-mode',
        codeGeneration: { strings: false, wasm: false },
    });
    const wrapped = `(async () => {\n${source}\n})()`;
    const compiled = new vm.Script(wrapped, { filename: 'office-script.js' });
    return compiled.runInContext(context, { timeout: timeoutMs });
}

/** 写出的文件按路径去重（后写覆盖先写），保留最后一次的体积。 */
function collectWrites(writes) {
    const byPath = new Map();
    for (const write of writes) byPath.set(write.path, write);
    return [...byPath.values()];
}

/**
 * 复检一个写出的文件：用对应格式模块重新解析，拿到真实 stats/outline/warnings。
 * 任何解析失败都不影响其它文件，只记成该文件的一条 note。
 */
async function inspectFile(write, env) {
    const entry = formatByExtension(write.path);
    const record = {
        path: write.path,
        bytes: write.bytes,
        format: entry?.id ?? null,
        stats: null,
        outline: [],
        warnings: [],
        note: null,
        theme: null,
    };
    if (entry === undefined) return record;
    const loaded = await loadFormat(entry.id);
    if (loaded.module?.read === undefined) {
        record.note = `无法复检：${loaded.error?.message ?? '模块缺少 read()'}`;
        return record;
    }
    try {
        const report = loaded.module.read(write.path, env);
        record.stats = report?.stats ?? null;
        record.outline = Array.isArray(report?.outline) ? report.outline : [];
        record.warnings = Array.isArray(report?.warnings) ? report.warnings : [];
        record.theme = report?.theme;
        record.bytes = report?.bytes ?? write.bytes;
    } catch (error) {
        record.note = `复检失败：${error instanceof Error ? error.message : String(error)}`;
    }
    return record;
}

/**
 * 执行一次 office_run。
 * @param {unknown} rawArgs 工具参数
 * @param {{agent?: {session?: {header?: {cwd?: string}}}, signal?: AbortSignal}} exec 工具执行上下文
 * @param {ReturnType<import('./config.js').resolveConfig>} config
 */
export async function executeRun(rawArgs, exec, config) {
    const args = rawArgs !== null && typeof rawArgs === 'object' ? rawArgs : {};
    const script = typeof args.script === 'string' ? args.script : '';
    if (script.trim() === '') throw new Error('office_run：script 不能为空。先用 office_help 查写法。');
    if (script.length > config.maxScriptChars) {
        throw new Error(`office_run：script 太长（${script.length} 字符，上限 ${config.maxScriptChars}）。请拆成多次或减少注释。`);
    }

    const cwd = exec?.agent?.session?.header?.cwd ?? process.cwd();
    const cache = createCache({ root: cwd, dir: config.cacheDir });
    // 缓存**不再**每次调用都清空：清了就没有跨调用复用，渲染好的 PDF 页面图
    // 下一步 read_image 就读不到（2026-09-23 的实测问题）。这里只收走过期条目，
    // 需要立刻清空时用 office.cache.clear() 或 office_run({ keepCache: false })。
    const pruned = cache.prune({
        maxAgeMs: config.cacheTtlMinutes > 0 ? config.cacheTtlMinutes * 60_000 : 0,
        maxBytes: config.cacheMaxBytes,
    });
    // 主题解析器把「没传主题」兜到配置的默认主题，格式模块因此永远拿得到主题。
    // config 也一起给进去：LaTeX 编译要用它（引擎、超时、模板目录），
    // 而格式模块的 create/read/edit 拿不到插件上下文，只能从 env 取。
    const env = createEnv({
        root: cwd,
        themeResolver: (id) => resolveTheme(id ?? config.defaultTheme),
        config,
    });

    const sdk = await buildSdk({ env, cache, config });
    const started = Date.now();
    let returned;
    let failure;

    try {
        exec?.signal?.throwIfAborted();
        returned = await runInVm(script, sdk.office, config.scriptTimeoutMs);
    } catch (error) {
        failure = cleanError(error);
    }

    const writes = collectWrites(env.writes);
    const files = [];
    const otherFiles = [];
    for (const write of writes) {
        if (isOfficeFile(write.path)) files.push(await inspectFile(write, env));
        else otherFiles.push({ path: write.path, bytes: write.bytes });
    }

    // clearedAfter 表示「这次调用把缓存清空了」；keepCache 默认跟随配置（默认保留），
    // 只有显式 keepCache:false 或配置里关了保留才会清。清空时目录本身一起删掉。
    const keepAfter = args.keepCache !== undefined ? args.keepCache === true : config.keepCache;
    let clearedAfter = false;
    if (!keepAfter) {
        cache.clear();
        clearedAfter = true;
    }
    const kept = clearedAfter ? [] : cache.list().filter((entry) => !entry.dir);

    const warnings = [...env.warnings];
    for (const file of files) {
        for (const warning of file.warnings) warnings.push(`${file.path}：${warning}`);
    }
    if (sdk.loadIssues.length > 0) warnings.push(`部分格式不可用：${sdk.loadIssues.join('；')}`);
    // 「没写出文件」这条警告只在真的什么都没产出时说：PDF 渲染（artifacts）与
    // 命中复用的页面图（hits）都算产出，否则「只读一份 PDF」会被误报成漏了 save()。
    if (failure === undefined && writes.length === 0 && cache.stats.artifacts === 0 && cache.stats.hits === 0) {
        warnings.push('这次调用没有写出任何文件；如果本意是产出文档，检查脚本里是否真的调用了 save()。');
    }

    // 工具输出会被宿主按「无损 JSON」校验，`undefined` 不是合法 JSON。
    // 统一在这里收口，之后往结果里加字段就不必再惦记这件事。
    //
    // purpose 与 error 是可选字段，**不要写 null 顶替**：宿主会按 schema 校验返回值。
    // 2026-09-21 实测更严：输出 schema 里声明过的字段只要缺键，整次调用就会被判成
    // 「返回了非法输出」——成功路径不写 error 报 "value.error must be an object"，
    // 不传 purpose 报 "value.purpose must be a string"。因此在 tools.js 里这两个字段
    // 已经不声明类型，这里也就没有「必须补一个空对象」的负担：没有就不写。
    const result = {
        ok: failure === undefined,
        files,
        otherFiles,
        returned: safeJson(returned),
        logs: env.logs,
        notes: env.notes,
        warnings,
        cache: {
            dir: cache.rel,
            kept: kept.length,
            keptBytes: kept.reduce((sum, entry) => sum + entry.bytes, 0),
            files: kept.slice(0, 20).map((entry) => entry.name),
            pruned: pruned.removed.slice(0, 20),
            prunedBytes: pruned.bytes,
            hits: cache.stats.hits,
            hitBytes: cache.stats.hitBytes,
            ttlMinutes: config.cacheTtlMinutes,
            clearedAfter,
        },
        formats: [...formatIds(), 'pdf'],
        elapsedMs: Date.now() - started,
    };
    if (typeof args.purpose === 'string' && args.purpose.trim() !== '') result.purpose = args.purpose.trim();
    if (failure !== undefined) result.error = failure;

    // ── 记忆：把这次写出的文件登记进台账，并把热记忆带回反馈 ──────────────
    //
    // 台账是三层记忆的第二层。产物一落盘就自动记一条，模型不必为此多调一次
    // 工具 —— 这是「自动捕获」在办公场景里的落点：办公_run 的结果本身就是最
    // 可靠的事实来源（真实路径、真实统计），比让模型复述一遍准确得多。
    //
    // 写记忆失败绝不能把这次执行的结论弄坏：文件已经在盘上了，记忆只是附加
    // 信息，所以只把失败原因写进反馈（renderRun 会印出来）。
    if (config.memory?.enabled !== false) {
        try {
            const memory = createMemory({ root: cwd, memory: config.memory });
            // autoLedger 只管「要不要自动登记」：关掉仍然把热记忆与最近台账带回反馈，
            // 否则关掉自动登记等于顺手把「动笔前看记忆」这件事也一起关了。
            const autoLedger = config.memory?.autoLedger !== false;
            const digest = await memory.digest({ recentLedger: 2 });
            const logged = autoLedger && files.length > 0
                ? await memory.log(files.map((file) => ({
                    path: file.path,
                    format: file.format,
                    theme: file.theme,
                    bytes: file.bytes,
                    purpose: result.purpose,
                    outline: file.outline,
                    stats: file.stats,
                    note: file.note,
                })))
                : { added: 0, kept: digest.ledgerTotal, rolled: 0 };
            result.memory = {
                enabled: true,
                autoLedger,
                logged: logged.added,
                kept: logged.kept,
                rolled: logged.rolled,
                // 投影只在热记忆变化时贴全文（没变就给一行 + 台账照旧）——
                // office_run 是同一个会话里会被反复调的入口，详见 projection.js。
                //
                // 第十七轮再把「没变」那一档压成一行：office_run 的职责是把
                // **这次写出了什么**说清楚，而热记忆的正文在动笔前的 office_help
                // 上已经贴过（同一份热记忆，revision 没变就等于已经在上下文里）。
                // 原来每次 office_run 都重贴一段骨架（标题 + 两条尾巴 + 用量行 ≈
                // 1.1 KB），一轮里调三次就是 3.3 KB，收益为零。
                ...(digest.empty === false
                    ? (() => {
                        const projected = projectDigest(digest, { sessionId: sessionIdOf(exec), context: 'run' });
                        return {
                            digest: projected.mode === 'full'
                                ? projected.text
                                : '📒 热记忆与上次投影相同（要看：office_memory({ action: \'read\', layer: \'hot\' })）。',
                            digestMode: projected.mode,
                        };
                    })()
                    : {}),
            };
        } catch (error) {
            result.memory = {
                enabled: true,
                autoLedger: config.memory?.autoLedger !== false,
                logged: 0,
                kept: 0,
                rolled: 0,
                error: error instanceof Error ? error.message : String(error),
            };
        }
    }
    return jsonSafe(result);
}
