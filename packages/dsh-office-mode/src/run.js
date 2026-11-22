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
import { createEnv, displayPath } from './engine/kit.js';
import { resolveTheme } from './engine/theme.js';
import { createMemory } from './memory.js';
import { projectDigest, sessionIdOf, signalOf } from './projection.js';
import { formatByExtension, formatIds, loadFormat } from './registry.js';
import { buildSdk, isOfficeFile } from './sdk.js';

/** 包装脚本时插入的行数：`(async () => {` 占一行。 */
const WRAPPER_LINES = 1;

/** 内联脚本在 vm 里的文件名（错误行号按它换算）。 */
const INLINE_SCRIPT_NAME = 'office-script.js';

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
    // 二进制（office.files.readBytes 的返回值）：按索引展开会变成几十个数字键，
    // 对模型零信息。只报字节数。（`ArrayBuffer.isView` 跨 realm 也认。）
    if (ArrayBuffer.isView(value)) return `[二进制 ${value.byteLength} 字节]`;
    if (Object.prototype.toString.call(value) === '[object ArrayBuffer]') return `[二进制 ${value.byteLength} 字节]`;
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
    if (ArrayBuffer.isView(value)) return `[二进制 ${value.byteLength} 字节]`;
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
function cleanError(error, filename = INLINE_SCRIPT_NAME) {
    const raw = isErrorLike(error) ? error : new Error(String(error));
    const stack = typeof raw.stack === 'string' ? raw.stack : '';
    const escaped = filename.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // 列号是**可选**的：语法错误（编译期）的栈帧只有 `文件:行`，运行期错误才有 `文件:行:列`。
    // 只认带列号的那种写法会让「脚本文件里写错一个字符」这类最常见的失败拿不到行号
    // （回执只剩 `脚本第 null 行` 那一类退化，2026-10-04 实测）。
    const match = new RegExp(`${escaped}:(\\d+)(?::(\\d+))?`).exec(stack);
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
        column: match && match[2] !== undefined ? Number.parseInt(match[2], 10) : undefined,
        stack: frames,
    };
}

/**
 * 在受限上下文里执行脚本。没有 process / require / fs / 网络 / 定时器。
 *
 * `options.args` 会成为脚本里的全局 `args`（`新-7`：跨调用复用同一段脚本时，
 * 参数从外面给，脚本本身不必改）；`options.filename` 决定错误行号按哪个文件名换算。
 */
function runInVm(source, office, timeoutMs, options = {}) {
    const sandbox = {
        office,
        args: {},
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
    // `args` 必须是 **vm realm 自己的对象**：直接塞宿主对象等于给脚本一个
    // `args.constructor.constructor('return process')()` 的把手 —— `codeGeneration:{strings:false}`
    // 只挡 vm 自己那套 `eval` / `new Function`，挡不住宿主对象的构造器链（2026-10-04 实测）。
    // 走一遍 JSON 往返最省事，而 `args` 本来就只承诺 JSON 值（调用方已先验过可序列化）。
    if (options.args !== undefined) {
        try {
            vm.runInContext(`args = (${JSON.stringify(options.args)})`, context);
        } catch {
            // 极端形状（循环引用等）在调用方已经记过警告；这里保持空对象，不让它变成激活失败。
            vm.runInContext('args = {}', context);
        }
    }
    const wrapped = `(async () => {\n${source}\n})()`;
    const compiled = new vm.Script(wrapped, { filename: options.filename ?? INLINE_SCRIPT_NAME });
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
 * 台账 `note` 里的告警摘要：最多列 4 条，多的只报条数。
 *
 * 告警是**逐项 push** 的（一个 500 处公式错误的 xlsx 能带出几百条），而 note 会进
 * `ledgerSearchText` 参与每次检索的匹配 —— 不封顶就等于给台账塞一条几十 KB 的键。
 * memory.log 那一侧还有一道 400 字符的兜底（手工 log 也走那条）。
 */
function warningsNoteOf(warnings) {
    const list = Array.isArray(warnings) ? warnings : [];
    if (list.length === 0) return '';
    const head = list.slice(0, 4).join('；');
    return list.length > 4 ? `${head}；…共 ${list.length} 条告警` : head;
}

/**
 * 执行一次 office_run。
 * @param {unknown} rawArgs 工具参数
 * @param {{agent?: {session?: {header?: {cwd?: string}}}, signal?: AbortSignal}} exec 工具执行上下文
 * @param {ReturnType<import('./config.js').resolveConfig>} config
 */
export async function executeRun(rawArgs, exec, config) {
    const args = rawArgs !== null && typeof rawArgs === 'object' ? rawArgs : {};
    const inline = typeof args.script === 'string' ? args.script : '';
    const scriptFile = typeof args.scriptFile === 'string' ? args.scriptFile.trim() : '';
    // 两种来源都给了是**歧义**，不猜：猜错的那次调用会跑出一份与用户预期无关的结果，
    // 而两边都写了的调用方本来就没想清楚。（`新-7` 的入口就是 scriptFile。）
    if (inline.trim() !== '' && scriptFile !== '') {
        throw new Error('office_run：script 与 scriptFile 只能给一个 —— scriptFile 是「同一段脚本跨调用复用」的入口。');
    }
    if (inline.trim() === '' && scriptFile === '') {
        throw new Error('office_run：script 或 scriptFile 必须给一个。先用 office_help 查写法。');
    }
    if (inline.length > config.maxScriptChars) {
        throw new Error(`office_run：script 太长（${inline.length} 字符，上限 ${config.maxScriptChars}）。请拆成多次或减少注释。`);
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

    // 脚本来源：内联那段，或从工作目录读一份 .js（`新-7` 的「持久化脚本、复用」）。
    // 走 env.readText 而不是自己读盘：这样脚本文件也会被登记成「这次读过什么」，
    // 台账的 source 因此能回答「这份产物是拿哪段脚本构建的」。
    let script = inline;
    let scriptLabel;
    if (scriptFile !== '') {
        let text;
        try {
            text = env.readText(scriptFile);
        } catch (error) {
            throw new Error(`office_run：读不到脚本文件 ${scriptFile}（${error.message}）。`
                + '先用 office.files.write 把脚本落成 .js，或改用 script 内联。');
        }
        if (text.trim() === '') throw new Error(`office_run：脚本文件 ${scriptFile} 是空的。`);
        // 二进制当脚本读会得到一句 `Invalid or unexpected token`（光标停在一堆乱码里），
        // 模型据此只能猜。UTF-8 里的空字节解码后就是 U+0000，用它当判据与 office.files.read 同口径。
        if (text.includes('\u0000')) {
            throw new Error(`office_run：脚本文件 ${scriptFile} 看着是二进制文件，不是 JavaScript 源码；`
                + '要跑脚本请给它一份 .js 文本（office.files.write 落盘）。');
        }
        if (text.length > config.maxScriptChars) {
            throw new Error(`office_run：脚本文件 ${scriptFile} 太长（${text.length} 字符，上限 ${config.maxScriptChars}）。请拆成多次或减少注释。`);
        }
        script = text;
        scriptLabel = displayPath(cwd, env.resolve(scriptFile));
    }
    // 传给脚本的参数（全局 `args`）：只在是**可序列化的普通对象**时采用，其它形状
    // 按「没给」处理并落一条警告 —— 静默变成 `{}` 会让脚本取到 undefined 却毫无信号
    // （第五十一轮独立复核的一条）。
    let scriptArgs = {};
    if (args.args !== undefined && args.args !== null) {
        const value = args.args;
        if (typeof value !== 'object' || Array.isArray(value)) {
            env.warn(`args 不是对象（${Array.isArray(value) ? '数组' : typeof value}），已按 {} 处理；要传参数请给一个对象。`);
        } else {
            try {
                JSON.stringify(value);
                scriptArgs = value;
            } catch (error) {
                env.warn(`args 无法序列化（${error instanceof Error ? error.message : String(error)}），已按 {} 处理。`);
            }
        }
    }

    const sdk = await buildSdk({ env, cache, config });
    const started = Date.now();
    let returned;
    let failure;

    try {
        exec?.signal?.throwIfAborted();
        returned = await runInVm(script, sdk.office, config.scriptTimeoutMs, {
            filename: scriptLabel ?? INLINE_SCRIPT_NAME,
            args: scriptArgs,
        });
    } catch (error) {
        failure = cleanError(error, scriptLabel ?? INLINE_SCRIPT_NAME);
    }

    const writes = collectWrites(env.writes);
    const files = [];
    const otherFiles = [];
    for (const write of writes) {
        if (isOfficeFile(write.path)) files.push(await inspectFile(write, env));
        else otherFiles.push({ path: write.path, bytes: write.bytes });
    }

    // 台账的 `source`（第十九轮 P1-4 / 第二十四轮 P1-3）：本次脚本**真的读过**、
    // 而且不是这次自己写出来的那些输入文件。
    //
    // 两条口径都不是随手定的：
    //   · 只算读过的内容（env.reads）——`exists` / `stat` / `list` 只是探路，不是依据；
    //   · 自己写出来的文件不算来源 —— 原地 revise 会先读同一份产物，记进去是同义反复
    //     （「这份 pptx 的来源是这份 pptx」），只会把台账的检索键灌满噪声。
    // 上限 8 条与 memory.log 的 source 截断一致。
    //
    // 网络来源（第三十八轮起 env.citations：图片的图库页 / 许可页）排在文件之后：
    // 它们同样是「这份产物依据了什么」，但不如本地输入常用，超出上限时先丢它们。
    const writtenPaths = new Set(writes.map((write) => write.path));
    const sourcePaths = [];
    const seenSource = new Set();
    for (const read of env.reads) {
        if (writtenPaths.has(read.path) || seenSource.has(read.path)) continue;
        seenSource.add(read.path);
        sourcePaths.push(read.path);
        if (sourcePaths.length >= 8) break;
    }
    for (const citation of env.citations ?? []) {
        if (sourcePaths.length >= 8) break;
        if (seenSource.has(citation.url)) continue;
        seenSource.add(citation.url);
        const license = citation.license === '' ? '' : `（${citation.license}）`;
        sourcePaths.push(`${citation.url}${license}`);
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
    // 脚本来自文件时把它报出来：失败时模型才知道该去改哪一份（行号是**那份文件**里的行号）。
    if (scriptLabel !== undefined) result.scriptFile = scriptLabel;
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
                    // 复检没给 note 时，把这次文件自己的告警折进 note：`ledgerSearchText`
                    // 的排序指望 outline + source + note 三块，而告警（如「正文溢出」）
                    // 正是「以后再找这份东西」时最可能被想起的词。有界，见 warningsNoteOf。
                    note: file.note ?? warningsNoteOf(file.warnings),
                    source: sourcePaths,
                })))
                : { added: 0, kept: digest.ledgerTotal, rolled: 0 };
            result.memory = {
                enabled: true,
                autoLedger,
                logged: logged.added,
                kept: logged.kept,
                rolled: logged.rolled,
                // 投影只在热记忆变化、话题信号切换或台账指纹变化时贴正文（没变就一行）——
                // office_run 是同一个会话里会被反复调的入口，详见 projection.js。
                //
                // 第十七轮把「没变」那一档压成一行：office_run 的职责是把
                // **这次写出了什么**说清楚，而热记忆的正文在动笔前的 office_help
                // 上已经贴过（同一份热记忆，writeRevision 没变就等于已经在上下文里）。
                // 原来每次 office_run 都重贴一段骨架（标题 + 两条尾巴 + 用量行 ≈
                // 1.1 KB），一轮里调三次就是 3.3 KB，收益为零。
                //
                // 注意时序：这里的 digest 在 log **之前**取，所以本次刚登记的台账
                // 行不在这份反馈里（登记数在 memory.logged 上）；下一次调用会看到
                // 台账指纹变了、把最近几行贴出来。
                ...(digest.empty === false
                    ? (() => {
                        // 信号来自脚本本身：里面写的 office.ppt… / .pptx 就是这次在做的事，
                        // 投影据此只给相关条目（详见 projection.js 的 signalOf）。
                        const projected = projectDigest(digest, {
                            sessionId: sessionIdOf(exec),
                            context: 'run',
                            signal: signalOf({ script }),
                            // 投影预算（第四十八轮 P0-2）：0 / 未配置 = 不限
                            budget: config?.memory?.projectionBudgetBytes,
                        });
                        // 未变化就不带 digest 字段：反馈里不再为「和上次一样」付一行 token。
                        return projected.mode === 'full'
                            ? { digest: projected.text, digestMode: projected.mode }
                            : {};
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
