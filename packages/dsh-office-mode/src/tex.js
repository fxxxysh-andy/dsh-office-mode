/**
 * LaTeX 工具链：探测本机 TeX 发行版、定位 thuthesis 模板、编译学位论文并解析日志。
 *
 * 为什么单独做一层：Word / Excel / PPT 的产物由插件自己写字节（零依赖、完全可控），
 * 而 PDF 是 TeX 发行版算出来的 —— 插件既不能也不必替它排版。所以这一层的职责是
 * 「把机器上已有的 TeX 用起来，并且把它说的话翻译回模型能懂的话」：
 *
 *   - 探测：xelatex / lualatex / latexmk / bibtex / biber / kpsewhich 在不在；
 *     一个都没有时明确报错并说清装什么，不假装支持（与 pdf.js 的取向一致）。
 *   - 定位：thuthesis 模板在哪（用户配置 → 工作目录里的 thuthesis* → TeX Live 自带）。
 *     工作目录里的那一份优先，因为论文模板的版本必须与源文件对得上。
 *   - 编译：latexmk 优先（它自己会算 bibtex 与重复编译次数），没有 latexmk 才退化成
 *     手工跑 xelatex ×2 → bibtex → xelatex ×2。
 *   - 解析：把 .log 里的 `文件:行号: 错误`、`! LaTeX Error`、Overfull/Underfull
 *     计数、页数与体积抽成结构化字段。模型只看这份报告就该知道编译成没成、错在哪。
 *
 * 跑的姿势沿用 pdf.js 的实测结论：**全程不用管道**，stdout / stderr 重定向到普通
 * 文件句柄（受限沙箱下 Node 的 stdio:'pipe' 会 spawn EPERM），因此复用 runPdfProcess。
 * 另外必须带 cwd —— latexmk 的中间产物与输出都按当前目录算，不切目录会把 .aux
 * 撒到工作目录根上。
 *
 * @module dsh-office-mode/tex
 */
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync, statSync, unlinkSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { displayPath } from './engine/kit.js';
import { findPdfExecutable, runPdfProcess } from './pdf.js';

/** thuthesis 只支持 xelatex / lualatex（ctex + fontspec），pdflatex 碰不了中文字体，不在候选里。 */
export const TEX_ENGINE_IDS = ['xelatex', 'lualatex'];
/** 编译兜底超时：学位论文三四遍编译 + bibtex，慢机器上几分钟是正常的，但不能没有上限。 */
export const DEFAULT_TEX_TIMEOUT_MS = 300_000;
/** 模板环境变量：给部署层一个不改配置就能指路的入口。 */
export const TEX_TEMPLATE_ENV = 'DSH_OFFICE_TEX_TEMPLATE';

/**
 * 编译一个 thuthesis 项目所需的模板文件（与版本绑定，必须成套拷贝）。
 * 只带 .cls 会少封面校徽（thu-fig-logo.pdf / thu-text-logo.pdf 由宏包直接读），
 * 只带 .cls 与校徽又会少参考文献样式（bibtex 找不到 thuthesis-numeric.bst）。
 */
export const SUPPORT_PATTERNS = [
    /^thuthesis\.cls$/i,
    /^thuthesis-[\w-]+\.(bst|bbx|cbx)$/i,
    /^thuthesis\.(bst|bbx|cbx)$/i,
    /^thu-(fig|text)-logo\.pdf$/i,
];

/** 手工清理时要删的中间产物扩展名（latexmk 缺席时的兜底）。 */
const BUILD_EXTENSIONS = ['.aux', '.log', '.out', '.toc', '.lof', '.lot', '.bbl', '.blg', '.idx', '.ind', '.ilg', '.xdv', '.fls', '.fdb_latexmk', '.synctex.gz', '.bcf', '.run.xml', '.nav', '.snm', '.nlo', '.nls', '.thm'];

let probeResult;

/** 探测本机可用的 TeX 程序（进程内只探一次，`refresh` 可强制重探）。 */
export function probeTexEngines(options = {}) {
    if (probeResult !== undefined && options.refresh !== true) return probeResult;
    probeResult = {
        xelatex: findPdfExecutable(['xelatex']),
        lualatex: findPdfExecutable(['lualatex']),
        pdflatex: findPdfExecutable(['pdflatex']),
        latexmk: findPdfExecutable(['latexmk']),
        bibtex: findPdfExecutable(['bibtex']),
        biber: findPdfExecutable(['biber']),
        kpsewhich: findPdfExecutable(['kpsewhich']),
        xdvipdfmx: findPdfExecutable(['xdvipdfmx']),
    };
    return probeResult;
}

/**
 * 本机 LaTeX 能力一览（给 office.tex.engines() 与报错信息用）。
 * `available` 里是能直接用来编译的入口：latexmk（推荐）/ xelatex / lualatex。
 */
export async function texEngines() {
    const probe = probeTexEngines();
    const available = [];
    if (probe.latexmk !== undefined && (probe.xelatex !== undefined || probe.lualatex !== undefined)) available.push('latexmk');
    if (probe.xelatex !== undefined) available.push('xelatex');
    if (probe.lualatex !== undefined) available.push('lualatex');
    const paths = {};
    for (const [name, value] of Object.entries(probe)) if (value !== undefined) paths[name] = value;
    return {
        available,
        tools: Object.keys(paths),
        paths,
        bibliography: [probe.bibtex !== undefined ? 'bibtex' : null, probe.biber !== undefined ? 'biber' : null].filter((item) => item !== null),
        hint: available.length > 0
            ? '编译走 latexmk（自动决定 bibtex 与重复编译次数）；只装了 xelatex 时退化成手工多遍编译。'
            : '本机没有 TeX 发行版：装 TeX Live（或 MiKTeX）后 office.tex.compile 才能出 PDF；只有源文件不需要它。',
    };
}

/**
 * 找一个「成套」的 thuthesis 模板目录。优先级：
 *   1. 显式给的目录（配置 texTemplateDir / create 的 templateDir）
 *   2. 环境变量 DSH_OFFICE_TEX_TEMPLATE
 *   3. 会话工作目录下叫 thuthesis* 的目录（深浅两层）—— 工作区里放着模板时这一条命中
 *   4. TeX Live 自带（kpsewhich thuthesis.cls）
 *
 * `complete` 表示这一份能不能整套拷走：TeX Live 自带的 .cls 与 .bst 分在两个树里，
 * 拷不全，所以那种来源只用来「就地编译」，不拷贝。
 */
export function locateTemplateDir(env, explicit) {
    const root = (env !== null && typeof env === 'object' && typeof env.root === 'string') ? env.root : process.cwd();
    const candidates = [];
    const push = (dir, source) => {
        if (typeof dir !== 'string' || dir.trim() === '') return;
        const absolute = resolve(root, dir.trim());
        if (candidates.some((item) => item.dir === absolute)) return;
        candidates.push({ dir: absolute, source });
    };

    push(explicit, 'options');
    push(process.env[TEX_TEMPLATE_ENV], 'env');

    // 工作目录里找 thuthesis<版本>/ 这种目录名；thuthesis.cls 必须在里面。
    try {
        for (const entry of readdirSync(root, { withFileTypes: true })) {
            if (!entry.isDirectory() || !/^thuthesis/i.test(entry.name)) continue;
            push(join(root, entry.name), 'workspace');
            try {
                for (const inner of readdirSync(join(root, entry.name), { withFileTypes: true })) {
                    if (inner.isDirectory()) push(join(root, entry.name, inner.name), 'workspace');
                }
            } catch {
                // 读不了就跳过
            }
        }
    } catch {
        // 工作目录读不了（权限）时继续走下一条
    }

    for (const candidate of candidates) {
        if (!existsSync(join(candidate.dir, 'thuthesis.cls'))) continue;
        return { ...candidate, complete: isCompleteTemplate(candidate.dir) };
    }

    // TeX Live 自带：kpsewhich 找得到就说明能就地编译（但拷不全，所以 complete 恒为 false）。
    const probe = probeTexEngines();
    if (probe.kpsewhich !== undefined) {
        const found = runKpsewhich(probe.kpsewhich, ['thuthesis.cls']);
        if (found !== undefined) {
            return { dir: dirname(found), source: 'texlive', complete: false, cls: found };
        }
    }
    return undefined;
}

/** 目录里有没有成套的运行期文件（.cls + 两个校徽 + 至少一个参考文献样式）。 */
export function isCompleteTemplate(dir) {
    let entries = [];
    try {
        entries = readdirSync(dir);
    } catch {
        return false;
    }
    const has = (pattern) => entries.some((name) => pattern.test(name));
    return has(/^thuthesis\.cls$/i)
        && has(/^thu-fig-logo\.pdf$/i)
        && has(/^thu-text-logo\.pdf$/i)
        && has(/^thuthesis-[\w-]+\.bst$/i);
}

/**
 * 同步跑一次 kpsewhich（只在探测模板时用，输出很短）。
 * 这里刻意用 spawnSync 而不是 runPdfProcess：locateTemplateDir 是同步函数
 * （create / read 这类路径上被调用），拿不到 Promise；kpsewhich 只输出一行路径，
 * 不存在管道刷屏的问题。
 */
function runKpsewhich(command, args) {
    try {
        const result = spawnSync(command, args, { encoding: 'utf8', timeout: 20_000, windowsHide: true });
        const line = String(result.stdout ?? '').split(/\r?\n/).map((item) => item.trim()).filter((item) => item !== '')[0];
        if (result.status === 0 && line !== undefined && existsSync(line)) return line;
    } catch {
        // 探测失败就当没有
    }
    return undefined;
}

/**
 * 把模板里成套的运行期文件拷进项目目录，让论文项目自包含（换机器、换 TeX 版本都能编）。
 * 返回 `{copied: [{name, bytes}], missing: [name]}`；调用方负责把它写进反馈。
 */
export function copySupport(templateDir, projectDir, env) {
    const copied = [];
    const missing = [];
    let entries = [];
    try {
        entries = readdirSync(templateDir, { withFileTypes: true });
    } catch (error) {
        return { copied, missing, error: `模板目录读不了：${error.message}` };
    }
    const wanted = entries.filter((entry) => entry.isFile() && SUPPORT_PATTERNS.some((pattern) => pattern.test(entry.name)));
    const essential = ['thuthesis.cls', 'thu-fig-logo.pdf', 'thu-text-logo.pdf'];
    for (const name of essential) {
        if (!wanted.some((entry) => entry.name.toLowerCase() === name.toLowerCase())) missing.push(name);
    }
    for (const entry of wanted.sort((a, b) => a.name.localeCompare(b.name))) {
        const bytes = readFileSync(join(templateDir, entry.name));
        // 支持文件是二进制且与源文件同目录：写盘走 env.writeFile，产物清单里才看得到它们。
        const written = env.writeFile(join(projectDir, entry.name), bytes);
        copied.push({ name: entry.name, path: written.path, bytes: written.bytes });
    }
    return { copied, missing };
}

/** LaTeX 特殊字符转义：模型的正文里出现 % & _ # $ 时用它，别让一句正常的话把编译搞挂。 */
export function escapeLatex(text) {
    return String(text ?? '')
        .replace(/\\/g, '\\textbackslash{}')
        .replace(/([&%$#_{}])/g, '\\$1')
        .replace(/~/g, '\\textasciitilde{}')
        .replace(/\^/g, '\\textasciicircum{}');
}

/* ── 编译 ───────────────────────────────────────────────────────────────── */

/** 解析后的错误项：`文件:行号: 说明`（-file-line-error 的格式，最实用）。 */
function parseLogErrors(log) {
    const errors = [];
    const seen = new Set();
    const push = (item) => {
        const key = `${item.file ?? ''}:${item.line ?? ''}:${item.message}`;
        if (seen.has(key)) return;
        seen.add(key);
        if (errors.length < 25) errors.push(item);
    };
    const lines = String(log ?? '').split(/\r?\n/);
    for (let index = 0; index < lines.length; index += 1) {
        const line = lines[index];
        const located = /^(\S+?\.(?:tex|sty|cls|def|cfg|bbx|cbx|bib)):(\d+):\s*(.+)$/.exec(line.trim());
        if (located !== null) {
            // xelatex 会把同一处错误重复很多遍（每一遍编译一次），这里已经去重。
            if (/^LaTeX Warning|^Package .* Warning/.test(located[3])) continue;
            push({ file: located[1].replace(/^\.\//, ''), line: Number.parseInt(located[2], 10), message: located[3].trim() });
            continue;
        }
        const bang = /^!\s*(.+)$/.exec(line.trim());
        if (bang !== null) {
            // 没有 -file-line-error 时只能看 `l.<行号>` 那一行。
            let source;
            for (let ahead = index + 1; ahead < Math.min(index + 8, lines.length); ahead += 1) {
                const lLine = /^l\.(\d+)\s*(.*)$/.exec(lines[ahead]);
                if (lLine !== null) {
                    source = { line: Number.parseInt(lLine[1], 10), context: lLine[2].trim().slice(0, 80) };
                    break;
                }
            }
            push({ file: null, line: source?.line ?? null, message: bang[1].replace(/\s+$/, ''), context: source?.context });
        }
    }
    return errors;
}

/** 日志里的排版提示：Overfull/Underfull 计数 + 缺字计数。 */
function parseLogQuality(log) {
    const text = String(log ?? '');
    const count = (pattern) => (text.match(pattern) ?? []).length;
    const pages = [];
    const overfullLines = [];
    for (const line of text.split(/\r?\n/)) {
        if (/^Overfull \\[hv]box/.test(line)) {
            overfullLines.push(line.trim().slice(0, 160));
            continue;
        }
        const out = /Output written on .*?\((\d+) pages?, (\d+) bytes\)/.exec(line);
        if (out !== null) pages.push({ pages: Number.parseInt(out[1], 10), bytes: Number.parseInt(out[2], 10), line: line.trim() });
    }
    return {
        overfull: count(/^Overfull \\[hv]box/gm),
        underfull: count(/^Underfull \\[hv]box/gm),
        missingChars: count(/^Missing character:/gm),
        overfullSamples: overfullLines.slice(0, 5),
        outputs: pages,
    };
}

/** 编译日志尾部（给反馈用；完整日志在项目目录里的 .log 文件，可以再读）。 */
function logTail(log, maxLines = 18) {
    const lines = String(log ?? '').split(/\r?\n/).map((line) => line.replace(/\s+$/, '')).filter((line) => line !== '');
    return lines.slice(-maxLines).join('\n');
}

/** 挑编译入口：auto = 有 latexmk 就用 latexmk，否则退回能用的 TeX 引擎。 */
function pickEngine(requested) {
    const wanted = typeof requested === 'string' && requested.trim() !== '' ? requested.trim().toLowerCase() : 'auto';
    const probe = probeTexEngines();
    if (wanted === 'auto') {
        if (probe.latexmk !== undefined) return probe.lualatex !== undefined && probe.xelatex === undefined ? 'lualatex' : 'latexmk';
        if (probe.xelatex !== undefined) return 'xelatex';
        if (probe.lualatex !== undefined) return 'lualatex';
        return 'auto';
    }
    return wanted;
}

/** 把 `engine` 翻译成命令行。`tex` 是真正的 TeX 引擎（latexmk 时用 -xelatex / -lualatex）。 */
function buildCommand(engineId, texEngine, mainName) {
    const probe = probeTexEngines();
    if (engineId === 'latexmk') {
        const flag = texEngine === 'lualatex' ? '-lualatex' : '-xelatex';
        return {
            command: probe.latexmk,
            args: [flag, '-interaction=nonstopmode', '-file-line-error', '-synctex=1', mainName],
        };
    }
    const command = engineId === 'lualatex' ? probe.lualatex : probe.xelatex;
    return { command, args: ['-interaction=nonstopmode', '-file-line-error', '-synctex=1', mainName] };
}

/**
 * 手工编译（本机没有 latexmk 时的兜底）：
 *   xelatex → (bibtex，若源文件里有 \bibliography) → xelatex → xelatex
 * 与 latexmk 的差别只是「多跑的遍数由固定规则决定」，不确定时多跑一遍不会错。
 */
async function runManualPasses(context) {
    const { command, args, projectDir, mainName, timeoutMs, capture, wantsBib } = context;
    const runs = [];
    const base = args.slice(0, -1);
    const runOnce = async (label) => {
        const result = await runPdfProcess(command, [...base, mainName], {
            ...capture(`build-${label}`),
            cwd: projectDir,
            timeoutMs,
        });
        runs.push({ step: label, exitCode: result.code, error: result.error });
        return result;
    };
    await runOnce('1');
    if (wantsBib) {
        const probe = probeTexEngines();
        const aux = join(projectDir, mainName.replace(/\.tex$/i, '.aux'));
        if (probe.bibtex !== undefined && existsSync(aux)) {
            const result = await runPdfProcess(probe.bibtex, [basename(aux)], {
                ...capture('bibtex'),
                cwd: projectDir,
                timeoutMs: Math.min(timeoutMs, 120_000),
            });
            runs.push({ step: 'bibtex', exitCode: result.code, error: result.error });
        }
    }
    await runOnce('2');
    const last = await runOnce('3');
    return { runs, last };
}

/**
 * 编译一个 .tex 主文件出 PDF。
 *
 * @param {string} mainPath 主文件（相对 env.root 或绝对）；必须含 \documentclass
 * @param {object} env 文件环境（engine/kit 的 createEnv；env.config 里带插件配置）
 * @param {{engine?: string, timeoutMs?: number, clean?: 'none'|'aux'|'all', cache?: object, support?: string, templateDir?: string}} [options]
 */
export async function compileTex(mainPath, env, options = {}) {
    const config = (env !== null && typeof env === 'object' && env.config !== null && typeof env.config === 'object') ? env.config : {};
    const absolute = env.resolve(mainPath);
    if (!existsSync(absolute)) throw new Error(`找不到 LaTeX 主文件：${mainPath}`);
    const mainName = basename(absolute);
    const projectDir = dirname(absolute);
    const source = readFileSync(absolute, 'utf8');
    if (!/\\documentclass/.test(source)) {
        throw new Error(`${mainPath} 不是主文件（里面没有 \\documentclass）。主文件指含文档类声明的那一个；按章分文件的论文里，data/chap01.tex 这种是片段，要编译整个项目请传主文件。`);
    }

    const engines = await texEngines();
    if (engines.available.length === 0) {
        throw new Error(
            '本机没有可用的 TeX 发行版（找不到 latexmk / xelatex / lualatex），无法编译出 PDF。'
            + '装 TeX Live 或 MiKTeX 之后不用改配置，插件会自动探测到；只要源文件不需要 TeX。',
        );
    }

    // 编译前先确保模板成套：优先项目目录里已有的，其次从能找到的模板目录拷，
    // 再次只能靠 TeX Live 自带的 thuthesis（版本可能与源文件不一致，给出提示）。
    const support = ensureSupport(env, projectDir, options);

    const engineId = pickEngine(options.engine ?? config.texEngine ?? 'auto');
    if (engineId === 'auto' || !engines.available.includes(engineId)) {
        const usable = engines.available[0];
        if (usable === undefined) throw new Error('本机没有可用的 TeX 编译入口。');
        if (engineId !== 'auto') {
            env.note(`指定的引擎 ${engineId} 不可用，改用 ${usable}（可用：${engines.available.join(' / ')}）。`);
        }
        return compileTexWith(usable === 'latexmk' ? 'latexmk' : usable, usable, env, { absolute, mainName, projectDir, source, options, config, support });
    }
    // latexmk 用哪个 TeX 引擎：配置里点了 lualatex 就用 lualatex，否则 xelatex。
    const texEngine = options.texEngine === 'lualatex' || config.texEngine === 'lualatex' ? 'lualatex' : 'xelatex';
    return compileTexWith(engineId, texEngine, env, { absolute, mainName, projectDir, source, options, config, support });
}

/** 编译主体：命令行已定，跑完解析日志与产物。 */
async function compileTexWith(engineId, texEngine, env, context) {
    const { absolute, mainName, projectDir, source, options, config, support } = context;
    const timeoutMs = clampTimeout(options.timeoutMs ?? config.texTimeoutMs);
    const stamp = `${process.pid}-${Date.now().toString(36)}`;
    const capture = (label) => ({
        outPath: join(projectDir, `.tex-${label}.${stamp}.out.txt`),
        errPath: join(projectDir, `.tex-${label}.${stamp}.err.txt`),
    });
    const cleanFiles = [];
    const started = Date.now();

    let runs = [];
    let result;
    const { command, args } = buildCommand(engineId, texEngine, mainName);
    if (command === undefined) {
        throw new Error(`本机没有 ${engineId}，无法编译。先跑 office.tex.engines() 看可用入口。`);
    }

    if (engineId === 'latexmk') {
        const paths = capture('build');
        cleanFiles.push(paths.outPath, paths.errPath);
        result = await runPdfProcess(command, args, { ...paths, cwd: projectDir, timeoutMs });
        runs = [{ step: 'latexmk', exitCode: result.code, error: result.error }];
    } else {
        const manual = await runManualPasses({
            command,
            args,
            projectDir,
            mainName,
            timeoutMs,
            capture,
            wantsBib: /\\bibliography\{/.test(source) || /\\addbibresource\{/.test(source),
        });
        runs = manual.runs;
        result = manual.last;
    }

    // 每次编译都会往项目目录写一堆临时文件；不清理会越积越多，也会污染交付目录。
    for (const file of cleanFiles) {
        try {
            unlinkSync(file);
        } catch {
            // 删不掉就留着
        }
    }

    const logPath = join(projectDir, mainName.replace(/\.tex$/i, '.log'));
    let log = '';
    try {
        log = readFileSync(logPath, 'utf8');
    } catch {
        log = '';
    }

    const errors = parseLogErrors(log);
    const quality = parseLogQuality(log);
    const pdfAbsolute = join(projectDir, mainName.replace(/\.tex$/i, '.pdf'));
    const hasPdf = existsSync(pdfAbsolute);
    const pdfBytes = hasPdf ? statSync(pdfAbsolute).size : 0;
    // 页数与体积优先信日志里的 "Output written on ..."：那是 TeX 自己报的数，
    // 不必再为一次编译去拉一个 PDF 解析器。
    const output = quality.outputs.length > 0 ? quality.outputs[quality.outputs.length - 1] : undefined;
    const ok = result.code === 0 && hasPdf && pdfBytes > 0;
    const elapsedMs = Date.now() - started;

    if (ok && options.clean !== 'none') await cleanBuild(env, projectDir, mainName, options.clean, engineId);

    const report = {
        ok,
        engine: engineId,
        texEngine: engineId === 'latexmk' ? texEngine : engineId,
        main: displayPath(env.root, absolute),
        projectDir: displayPath(env.root, projectDir),
        command: [basename(String(command)), ...args].join(' '),
        exitCode: result.code,
        timedOut: result.timedOut === true,
        elapsedMs,
        runs,
        pdf: hasPdf
            ? { path: displayPath(env.root, pdfAbsolute), bytes: pdfBytes, pages: output?.pages ?? null }
            : null,
        log: {
            path: existsSync(logPath) ? displayPath(env.root, logPath) : null,
            errors,
            overfull: quality.overfull,
            underfull: quality.underfull,
            missingChars: quality.missingChars,
            overfullSamples: quality.overfullSamples,
            tail: logTail(log),
        },
        support,
        pages: output?.pages ?? null,
    };

    if (ok) {
        report.hint = '编译成功。要看排版效果就用 office.pdf.pages() 渲染成图片再 read_image 逐页看；'
            + `${report.log.overfull} 处 Overfull 与 ${report.log.underfull} 处 Underfull 是排版提示，不一定是错误。`;
    } else if (result.timedOut === true) {
        report.hint = `编译超过 ${Math.round(timeoutMs / 1000)} 秒被中断。可以调大 office.tex.compile 的 timeoutMs，或先用 clearAux 清掉中间产物再试。`;
    } else {
        const first = errors[0];
        report.hint = first !== undefined
            ? `编译失败（退出码 ${String(result.code)}）。第一处：${first.file === null ? '' : `${first.file}:${String(first.line)} `}${first.message}`
            : '编译失败。看 log.errors 与 log.tail，完整日志在 log.path。';
    }
    return report;
}

/** 编译前保证模板可用：项目目录已有 → 拷一份 → 只剩 TeX Live 自带。 */
function ensureSupport(env, projectDir, options) {
    const wanted = typeof options.support === 'string' && options.support !== '' ? options.support : 'auto';
    const local = existsSync(join(projectDir, 'thuthesis.cls'));
    if (local) return { mode: 'project', copied: [], note: '项目目录里已有 thuthesis 模板文件。' };
    if (wanted === 'system' || wanted === 'none') {
        return { mode: 'system', copied: [], note: '按设置不拷贝模板：依赖 TeX Live 自带的 thuthesis。' };
    }
    const template = locateTemplateDir(env, options.templateDir ?? env.config?.texTemplateDir);
    if (template === undefined) {
        return { mode: 'missing', copied: [], note: '没找到成套的 thuthesis 模板目录，也没有 TeX Live 自带的版本；编译会因找不到 thuthesis.cls 失败。' };
    }
    if (template.complete !== true) {
        return { mode: 'texlive', copied: [], note: `用 TeX Live 自带的 thuthesis（${template.dir}）；它与源文件可能不是同一个版本。` };
    }
    const support = copySupport(template.dir, displayPath(env.root, projectDir), env);
    return {
        mode: 'copied',
        source: displayPath(env.root, template.dir),
        copied: support.copied.map((item) => item.name),
        missing: support.missing,
        note: `已从 ${displayPath(env.root, template.dir)} 拷贝模板文件，项目自包含（换机器、换 TeX 版本都能编）。`,
    };
}

function clampTimeout(value) {
    const n = Number(value);
    if (!Number.isFinite(n) || n <= 0) return DEFAULT_TEX_TIMEOUT_MS;
    return Math.min(1_800_000, Math.max(30_000, Math.trunc(n)));
}

/**
 * 清理中间产物。
 * latexmk 在就交给它（-c 清中间文件保留 PDF，-C 连 PDF 一起清），否则按扩展名手工删 ——
 * 手工删按「主文件名开头 + 已知扩展名」双条件匹配，绝不误伤同目录的其它文件。
 */
export async function cleanBuild(env, projectDir, mainName, mode = 'aux', engineId = 'latexmk') {
    const probe = probeTexEngines();
    const stem = mainName.replace(/\.tex$/i, '');
    if (probe.latexmk !== undefined && engineId === 'latexmk') {
        const outPath = join(projectDir, '.tex-clean.out.txt');
        const errPath = join(projectDir, '.tex-clean.err.txt');
        await runPdfProcess(probe.latexmk, [mode === 'all' ? '-C' : '-c'], { outPath, errPath, cwd: projectDir, timeoutMs: 60_000 });
        for (const name of [outPath, errPath]) {
            try {
                unlinkSync(name);
            } catch {
                // 无所谓
            }
        }
        return { cleaned: true, mode };
    }
    let removed = 0;
    let entries = [];
    try {
        entries = readdirSync(projectDir);
    } catch {
        return { cleaned: false, mode };
    }
    for (const name of entries) {
        if (!name.startsWith(stem)) continue;
        const lower = name.toLowerCase();
        if (!BUILD_EXTENSIONS.some((ext) => lower.endsWith(ext))) continue;
        if (mode === 'aux' && lower.endsWith('.pdf')) continue;
        try {
            unlinkSync(join(projectDir, name));
            removed += 1;
        } catch {
            // 删不掉就留着，不影响交付
        }
    }
    return { cleaned: true, mode, removed };
}
