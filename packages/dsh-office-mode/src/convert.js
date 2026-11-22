/**
 * 文档互转与统一出 PDF（office.convert）。
 *
 * 为什么需要它：办公模式一直能**造**三件套（office.word / excel / ppt）与 LaTeX 论文
 * （office.tex），也能**读**它们，但「手上的这份换个格式」一直只能重新生成一遍 ——
 * 别人发来一份 .md 要交 .docx、论文源要一份 .md 摘要、写完的 .md 要一份 PDF，
 * 这些都没有入口。这里把「格式 × 格式」补成一张闭合的表。
 *
 * 一张矩阵（源 → 目标）：
 *   md   → docx / tex / pdf        pandoc；pdf 走 LibreOffice 或 xelatex
 *   docx → md / tex / pdf          pandoc（md / tex）；pdf 走宿主 LibreOffice 引擎
 *   tex  → md / docx / pdf         pandoc（md / docx）；pdf 走本机 TeX 发行版
 *   其余（xlsx / pptx / odt / html / rst / txt…）→ pdf（或文本格式，pandoc 能读的都行）
 *
 * 四条口径：
 *   1. **引擎在运行期探测**：pandoc / TeX / LibreOffice 都不写进依赖，探不到就照实报
 *      缺什么、下一步装什么或换哪条路（`office.convert.check()`）；一条路不通时自动
 *      换下一条候选路，换过要说一声（env.note）。
 *   2. **产物先落缓存、再经 env.writeFile 写进工作目录**：只有走 env 写出的文件才会
 *      出现在 office_run 的反馈与台账里。LaTeX 那条路是例外中的例外（TeX 自己把 PDF
 *      写在主文件旁边），所以写回时用 `force` 让账目一定记上。
 *   3. **不覆盖**：目标已存在且内容不同时要显式 `overwrite:true`；内容一致就不重写
 *      （与 office.preview.pdf / office.image.fetch / office.archive.extract 同一口径）。
 *      `.tex → pdf` 的**默认落点**多一道：那里的 PDF 是 TeX 自己写出来的，先说清
 *      「那里已经有一份 PDF」并要求 `overwrite:true`，不静默重写。
 *   4. **中文字体是硬要求**：pandoc 的 LaTeX 路线在本机没有中文字体时会**静默丢字**
 *      （退出码仍是 0，PDF 里只剩空白），所以这里先探一个中文字体名，源里真有中文
 *      （整份扫，含二进制源一律按「有」处理）而探不到就宁可报错也不假装成功；
 *      stderr 里的 `Missing character` 会数出来进 warnings。
 *
 * @module dsh-office-mode/convert
 */
import { createHash } from 'node:crypto';
import { closeSync, existsSync, mkdirSync, openSync, readFileSync, readdirSync, readSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, delimiter, dirname, isAbsolute, join, resolve } from 'node:path';

import { assertInsideRoot, displayPath } from './engine/kit.js';
import { findPdfExecutable, runPdfProcess } from './pdf.js';
import { previewKitAvailable, previewPdf } from './preview.js';
import { compileTex, probeTexEngines } from './tex.js';

/** 目标格式 id（`office.convert.file` 的 `to` 取值）。 */
export const CONVERT_TARGETS = ['md', 'docx', 'tex', 'pdf'];

/** 目标格式 → 扩展名。 */
const TARGET_EXT = { md: '.md', docx: '.docx', tex: '.tex', pdf: '.pdf' };

/** `to` 的别名（带不带点都认）。 */
const TARGET_ALIASES = {
    md: 'md', markdown: 'md', gfm: 'md',
    docx: 'docx', word: 'docx',
    tex: 'tex', latex: 'tex',
    pdf: 'pdf',
};

/** `engine` 的别名。`auto` 之外的取值只在「出 PDF」时有意义。 */
const ENGINE_ALIASES = {
    auto: 'auto', pandoc: 'pandoc',
    libreoffice: 'libreoffice', lo: 'libreoffice', office: 'libreoffice',
    tex: 'tex', latex: 'tex', xelatex: 'tex', lualatex: 'tex', tectonic: 'tex',
};

/** 认识的源格式（按扩展名）。
 *   reader —— pandoc 的读入器名；没写就是 pandoc 读不了（只能原样出 PDF 或报错）
 *   pdf    —— 出 PDF 的候选路线，按优先级：
 *             tex  本机 TeX 发行版编译（源就是 .tex，唯一路线）
 *             direct 宿主 LibreOffice 引擎直接读（docx / xlsx / pptx / odt）
 *             docx  先用 pandoc 转成 docx，再交给 LibreOffice（纯文本源）
 *             latex pandoc 直接调 xelatex 出 PDF（LibreOffice 不可用时的退路）
 *             copy  源本来就是 PDF，原样复用
 *   media  —— 转成文本格式（md / tex）时要不要把内嵌图片抽成工作目录里的文件
 */
const SOURCES = [
    { id: 'md', exts: ['.md', '.markdown'], reader: 'markdown', pdf: ['docx', 'latex'], media: false },
    { id: 'docx', exts: ['.docx'], reader: 'docx', pdf: ['direct', 'latex'], media: true },
    // 老 Word（.doc）pandoc 读不了，但宿主 LibreOffice 引擎读得了 —— 只给出 PDF 这一条路。
    { id: 'doc', exts: ['.doc'], reader: undefined, pdf: ['direct'], media: false },
    { id: 'tex', exts: ['.tex', '.latex'], reader: 'latex', pdf: ['tex'], media: true },
    { id: 'pdf', exts: ['.pdf'], reader: undefined, pdf: ['copy'], media: false },
    { id: 'odt', exts: ['.odt'], reader: 'odt', pdf: ['direct', 'latex'], media: true },
    // OpenDocument 的表格与演示（.ods / .odp）：pandoc 读不了，同样只剩 LibreOffice 那条路。
    { id: 'odf', exts: ['.ods', '.odp'], reader: undefined, pdf: ['direct'], media: false },
    { id: 'xlsx', exts: ['.xlsx', '.xls'], reader: undefined, pdf: ['direct'], media: false },
    { id: 'pptx', exts: ['.pptx', '.ppt'], reader: undefined, pdf: ['direct'], media: false },
    { id: 'html', exts: ['.html', '.htm'], reader: 'html', pdf: ['docx', 'latex'], media: false },
    { id: 'rst', exts: ['.rst'], reader: 'rst', pdf: ['docx', 'latex'], media: false },
    { id: 'txt', exts: ['.txt'], reader: 'markdown', pdf: ['docx', 'latex'], media: false },
];

/** pandoc 可执行文件的环境变量（部署层不改配置就能指路）。 */
export const PANDOC_ENV = 'DSH_OFFICE_PANDOC';
/** 中文字体名的环境变量（点名一个本机字体，跳过自动探测）。 */
export const CJK_FONT_ENV = 'DSH_OFFICE_CJK_FONT';
/** 一次转换的兜底超时。LaTeX 编译是重活，但不能没有上限。 */
export const DEFAULT_CONVERT_TIMEOUT_MS = 300_000;
/** pandoc 版本探测的超时（只读一行版本号）。 */
const PROBE_TIMEOUT_MS = 20_000;
/** 一次转换最多往反馈里带几条引擎告警。 */
const MAX_WARNINGS = 6;
/** CJK 判据：汉字 / 假名 / 谚文（用于决定要不要挂中文字体）。 */
const CJK_RE = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uac00-\ud7af]/;
/** 中文字体候选：先看文件在不在，再把**字体名**交给 xeCJK。 */
const CJK_FONTS = [
    { name: 'Microsoft YaHei', file: 'C:\\Windows\\Fonts\\msyh.ttc' },
    { name: 'SimSun', file: 'C:\\Windows\\Fonts\\simsun.ttc' },
    { name: 'SimHei', file: 'C:\\Windows\\Fonts\\simhei.ttf' },
    { name: 'Noto Sans CJK SC', file: '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc' },
    { name: 'Noto Sans CJK SC', file: '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc' },
    { name: 'Source Han Sans SC', file: '/usr/share/fonts/opentype/adobe-source-han-sans/SourceHanSansSC-Regular.otf' },
    { name: 'PingFang SC', file: '/System/Library/Fonts/PingFang.ttc' },
];

let probeResult;

/** 源格式查表（认不出返回 undefined，调用方负责报错）。 */
function sourceEntry(filePath) {
    const at = String(filePath ?? '').lastIndexOf('.');
    const ext = at === -1 ? '' : String(filePath).slice(at).toLowerCase();
    return SOURCES.find((entry) => entry.exts.includes(ext));
}

/** 源格式 id 清单（给 check 与报错用）。 */
export function convertSourceIds() {
    return SOURCES.map((entry) => entry.id);
}

/** 认识的扩展名清单（报错里给用户看的；内部 id 里有 `odf` 这种磁盘上不存在的写法）。 */
export function convertExtensions() {
    return SOURCES.flatMap((entry) => entry.exts);
}

/**
 * pandoc 的位置：① 环境变量点名的（可给路径或命令名）② PATH 与常见安装目录。
 * 与 pdf.js 的 findExecutable 同一套「装了就能探到、不用改配置」的口径。
 */
function findPandoc() {
    const explicit = typeof process.env[PANDOC_ENV] === 'string' ? process.env[PANDOC_ENV].trim() : '';
    if (explicit !== '') {
        if (isAbsolute(explicit) || explicit.includes('/') || explicit.includes('\\')) {
            if (existsSync(explicit)) return explicit;
        }
        const named = findPdfExecutable([explicit]);
        if (named !== undefined) return named;
    }
    const fromPath = findPdfExecutable(['pandoc']);
    if (fromPath !== undefined) return fromPath;
    const dirs = [];
    const push = (dir) => {
        if (typeof dir === 'string' && dir !== '' && existsSync(dir)) dirs.push(dir);
    };
    if (process.platform === 'win32') {
        push(join(process.env.ProgramFiles ?? 'C:\\Program Files', 'Pandoc'));
        push(join(process.env['ProgramFiles(x86)'] ?? 'C:\\Program Files (x86)', 'Pandoc'));
        push(join(process.env.LOCALAPPDATA ?? '', 'Pandoc'));
        push(join(process.env.USERPROFILE ?? '', 'scoop', 'shims'));
        push(join(process.env.ProgramData ?? '', 'chocolatey', 'bin'));
    } else {
        for (const dir of ['/usr/bin', '/usr/local/bin', '/opt/homebrew/bin', '/snap/bin']) push(dir);
    }
    for (const dir of dirs) {
        for (const name of ['pandoc.exe', 'pandoc']) {
            const candidate = join(dir, name);
            if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
        }
    }
    return undefined;
}

/** 点名的中文字体名，其次按候选表看文件在不在。 */
function findCjkFont() {
    const named = typeof process.env[CJK_FONT_ENV] === 'string' ? process.env[CJK_FONT_ENV].trim() : '';
    if (named !== '') return named;
    for (const font of CJK_FONTS) if (existsSync(font.file)) return font.name;
    return undefined;
}

/**
 * 探测转换引擎（进程内只探一次；`refresh:true` 强制重探 —— 测试与设置改动后用）。
 *
 * @returns {{pandoc?: string, cjkFont?: string}}
 */
export function probeConvertEngines(options = {}) {
    if (probeResult !== undefined && options.refresh !== true) return probeResult;
    probeResult = {
        pandoc: findPandoc(),
        cjkFont: findCjkFont(),
        version: undefined,
    };
    return probeResult;
}

/** 测试用：清掉探测缓存（改了环境变量之后重新探）。 */
export function resetConvertProbe() {
    probeResult = undefined;
}

/** pandoc 版本（探测一次记住；只用来在 check 里报出实况）。 */
async function pandocVersion(probe) {
    if (probe.pandoc === undefined) return undefined;
    if (probe.version !== undefined) return probe.version;
    // 名字里带 pid 之外还要带一个随机段：同一进程里两次 check 并发跑时，只用 pid 会互相覆盖。
    const stamp = `${process.pid}-${Math.random().toString(36).slice(2, 8)}`;
    const outPath = join(tmpdir(), `office-pandoc-${stamp}.out.txt`);
    const errPath = join(tmpdir(), `office-pandoc-${stamp}.err.txt`);
    const result = await runPdfProcess(probe.pandoc, ['--version'], { outPath, errPath, timeoutMs: PROBE_TIMEOUT_MS });
    probe.version = null;
    if (result.code === 0) {
        const first = String(result.out ?? '').split(/\r?\n/)[0].trim();
        if (first !== '') probe.version = first;
    }
    try {
        rmSync(outPath, { force: true });
        rmSync(errPath, { force: true });
    } catch {
        // 临时文件删不掉不影响结论
    }
    return probe.version ?? undefined;
}

/** 引擎实况（一次转换里反复用到的四个事实）。 */
function engineState() {
    const probe = probeConvertEngines();
    const tex = probeTexEngines();
    const latexName = tex.xelatex !== undefined ? 'xelatex' : (tex.lualatex !== undefined ? 'lualatex' : undefined);
    return {
        probe,
        pandoc: probe.pandoc,
        cjkFont: probe.cjkFont,
        latexName,
        kit: previewKitAvailable() === true,
    };
}

/** 路线在不在（与会话里真正能不能跑是两件事：这里只看入口探没探到）。 */
function routeUsable(route, state) {
    if (route === 'copy') return true;
    if (route === 'tex') return state.latexName !== undefined;
    if (route === 'direct') return state.kit;
    if (route === 'docx') return state.kit && state.pandoc !== undefined;
    if (route === 'latex') return state.pandoc !== undefined && state.latexName !== undefined;
    return false;
}

/** 路线不可用时缺什么（照实写，并指到能装/能查的地方）。 */
function routeMissing(route) {
    if (route === 'tex') return '本机没有 TeX 发行版（探不到 xelatex / lualatex，装 TeX Live 或 MiKTeX）';
    if (route === 'direct') return '宿主随部署安装的 LibreOffice 引擎解析不到（@deepseek-ai/libreoffice-kit；非 DSH 环境会缺它）';
    if (route === 'docx') return 'pandoc 与 LibreOffice 引擎要同时在（探不到 ' + [PANDOC_ENV, '@deepseek-ai/libreoffice-kit'].join(' / ') + ' 之一）';
    if (route === 'latex') return 'pandoc 与本机 TeX 发行版要同时在（pandoc 探不到，或没有 xelatex / lualatex）';
    return '路线不可用';
}

/** `engine` 参数归一（认不出的当 auto，与「认不出的值收敛」同一口径）。 */
function normalizeEngine(value) {
    const key = String(value ?? '').trim().toLowerCase();
    return ENGINE_ALIASES[key] ?? 'auto';
}

/** `to` 归一：认不出就报错（这是必填项，收敛成兜底值只会把产物写到错的格式上）。 */
function normalizeTarget(value) {
    const key = String(value ?? '').trim().toLowerCase().replace(/^\./, '');
    const id = TARGET_ALIASES[key];
    if (id === undefined) {
        throw new Error(`office.convert.file：to 要写成 ${CONVERT_TARGETS.join(' / ')} 之一（也认 markdown / word / latex 与带点的写法）；`
            + `收到的是 ${JSON.stringify(value ?? null)}。`);
    }
    return id;
}

/** 源文件路径解析（相对工作目录或绝对；不存在时报出可执行的下一步）。 */
function resolveInput(filePath, env) {
    const wanted = typeof filePath === 'string' ? filePath.trim() : '';
    if (wanted === '') throw new Error('office.convert.file：要给一个文件路径（md / docx / tex / pdf / xlsx / pptx…）。');
    const absolute = env.resolve(wanted);
    if (!existsSync(absolute)) {
        throw new Error(`office.convert.file：找不到 ${wanted}。用 office.files.list 看工作目录里有什么。`);
    }
    if (!statSync(absolute).isFile()) {
        throw new Error(`office.convert.file：${wanted} 是目录，不是文件；这个入口一次转一份文件（多份就在脚本里循环）。`);
    }
    return absolute;
}

/** 目标路径：默认与源同目录同名（换扩展名）；`out` 覆盖。 */
function resolveTarget(absolute, target, options, env) {
    const out = typeof options.out === 'string' && options.out.trim() !== '' ? options.out.trim() : '';
    const stem = basename(absolute).replace(/\.[^.]+$/, '');
    const fallback = join(dirname(absolute), `${stem}${TARGET_EXT[target]}`);
    const candidate = out === '' ? displayPath(env.root, fallback) : out;
    const resolved = assertInsideRoot(candidate, env, 'office.convert.file');
    // 「目标就是源」要按**同一个文件**判，不能只比路径字符串：硬链接 / 符号链接 /
    // Windows 的大小写别名都能绕过字符串比较，而这道闸的契约是无条件的
    // （第五十一轮给 office.files.move 建同一套判据时踩过；这里直接用 env.sameFile）。
    const samePath = resolved === absolute;
    const sameFile = typeof env.sameFile === 'function' && env.sameFile(absolute, resolved) === true;
    if (samePath || sameFile) {
        throw new Error(`office.convert.file：目标就是源文件自己（${candidate}）—— 那会把原文件换成另一种格式。`
            + '换一个 out，或换一个 to。');
    }
    if (existsSync(resolved) && statSync(resolved).isDirectory()) {
        // 不拦的话下一站是 emit() 里的 readFileSync(目录)，报出来是 EISDIR 这种读不懂的错。
        throw new Error(`office.convert.file：目标 ${candidate} 是一个目录，不是文件；换一个 out。`);
    }
    return resolved;
}

function clampTimeout(value) {
    const n = Number(value);
    if (!Number.isFinite(n) || n <= 0) return DEFAULT_CONVERT_TIMEOUT_MS;
    return Math.min(1_800_000, Math.max(30_000, Math.trunc(n)));
}

/**
 * `.tex` 编译后的中间产物清理：只认 `'none'` 与 `'aux'`，其余（含 `'all'`）一律按 `'aux'`。
 *
 * 为什么不让 `'all'` 透传：latexmk 的 `-C` 与手工分支的 `all` 都会**连 PDF 一起删**，
 * 而 PDF 正是这次转换的交付物 —— 透传等于「先毁产物再报一个读不懂的 ENOENT」。
 */
function cleanMode(value, env) {
    const wanted = typeof value === 'string' ? value.trim().toLowerCase() : '';
    if (wanted === 'none') return 'none';
    if (wanted === 'all') {
        env.note('clean:\'all\' 在转换里按 \'aux\' 处理 —— 它会连 PDF 一起删掉，而 PDF 正是这次的交付物。');
    }
    return 'aux';
}

/** 缓存目录下的工作名（同名文件不同路径不会互相踩）。 */
function workKey(absolute) {
    const stem = basename(absolute).replace(/\.[^.]+$/, '').replace(/[^\w.-]+/g, '_').slice(0, 40);
    const hash = createHash('sha1').update(absolute).digest('hex').slice(0, 8);
    return `${stem}-${hash}`;
}

/**
 * 把字节写到目标位置。
 *
 * 三条口径：内容一致 → 不重写（不改 mtime、不重复登记）；已存在且内容不同 →
 * 必须显式 `overwrite:true`；一律经 `env.writeFile` 写 —— 只有走 env 写出的文件
 * 才会出现在 office_run 的反馈与台账里。`force` 给「引擎自己已经把文件写在那儿」
 * 的情形用（LaTeX 编译）：那时文件一定一致，但账目仍然要记上。
 */
function emit(env, absolute, bytes, options = {}) {
    const existedBefore = options.existedBefore === true || existsSync(absolute);
    if (existedBefore && options.force !== true) {
        let current;
        try {
            current = readFileSync(absolute);
        } catch {
            current = undefined;
        }
        if (current !== undefined && current.length === bytes.length && current.equals(bytes)) {
            return { path: displayPath(env.root, absolute), bytes: bytes.length, reused: true, wrote: false };
        }
        if (options.overwrite !== true) {
            throw new Error(`office.convert.file：${displayPath(env.root, absolute)} 已存在且内容不同；`
                + '换一个 out，或显式传 overwrite:true（覆盖别人的文件要明确说一声）。'
                + '若这是同一份源重转（.tex 编译与 LibreOffice 渲染出来的产物里带生成时间戳，字节不会完全相同），'
                + '确认要重写就传 overwrite:true。');
        }
    }
    const written = env.writeFile(absolute, bytes);
    return {
        path: written.path,
        bytes: written.bytes,
        reused: false,
        wrote: true,
        ...(existedBefore ? { overwrote: true } : {}),
    };
}

/** pandoc 的告警行（去重、封顶；只留引擎自己说的问题）。 */
function pandocWarnings(text) {
    const out = [];
    for (const line of String(text ?? '').split(/\r?\n/)) {
        const trimmed = line.trim();
        if (trimmed === '' || (!/^\[(WARNING|ERROR)\]/.test(trimmed) && !/^pandoc:/.test(trimmed))) continue;
        if (!out.includes(trimmed)) out.push(trimmed);
        if (out.length >= MAX_WARNINGS) break;
    }
    return out;
}

/** stderr 里数「缺字」的次数：pandoc + xelatex 走这条路时退出码仍是 0，只能靠它判。 */
function missingGlyphs(text) {
    const hits = String(text ?? '').match(/Missing character/g);
    return hits === null ? 0 : hits.length;
}

/**
 * 跑一次 pandoc：stdout / stderr 一律落文件（受限沙箱下不能用管道，见 pdf.js 的 runProcess）。
 *
 * `SOURCE_DATE_EPOCH` 设为**源文件的修改时间**：pandoc 写 docx 时会往 docProps/core.xml
 * 里塞当前时间，同一份源连转两次的字节就不一样（实测确认），于是「内容一致就不重写」
 * 这条复用判据永远失效、重复转换还会撞上覆盖闸。钉住它之后同一份源必得同一份产物，
 * LaTeX 引擎（xetex / pdftex）同样认这个变量，PDF 也顺手可复现。
 */
async function runPandoc(state, args, context) {
    const dir = context.cache.path(`convert/${context.label}`);
    mkdirSync(dir, { recursive: true });
    const epoch = sourceEpoch(context.sourcePath);
    const result = await runPdfProcess(state.pandoc, args, {
        outPath: join(dir, 'pandoc.out.txt'),
        errPath: join(dir, 'pandoc.err.txt'),
        cwd: context.cwd,
        timeoutMs: context.timeoutMs,
        env: epoch === undefined ? undefined : { ...process.env, SOURCE_DATE_EPOCH: String(epoch) },
    });
    return { result, warnings: pandocWarnings(result.err), missing: missingGlyphs(result.err) };
}

/** 源文件的修改时间（秒）；读不到就交给引擎自己决定时间戳。 */
function sourceEpoch(absolute) {
    try {
        return Math.max(0, Math.floor(statSync(absolute).mtimeMs / 1000));
    } catch {
        return undefined;
    }
}

/** pandoc 失败时的可读报错（带退出码、引擎原话的一小段与命令）。 */
function pandocFailure(args, run, env) {
    const { result } = run;
    const detail = String(result.error ?? result.err ?? '').replace(/\s+/g, ' ').trim().slice(0, 300);
    const timedOut = result.timedOut === true ? '：超时被中断（可以调大 timeoutMs）' : '';
    // 命令里的绝对路径一律折成相对工作目录的写法：错误文本会进会话，本机目录不该跟着走。
    const shown = args.map((arg) => (typeof arg === 'string' && arg !== '' ? displayPath(env.root, arg) : arg));
    return `office.convert.file：pandoc 没转成功（退出码 ${String(result.code)}）${timedOut}`
        + (detail === '' ? '' : `：${detail}`)
        + `。命令：pandoc ${shown.join(' ')}`;
}

/**
 * 中文字体：点名优先；源里没有 CJK 就不挂。
 *
 * 三条口径：
 *   - 显式给了 `font` 就**一定挂上**（用户点名要它，不该被自动判据丢掉）；
 *   - 二进制源（docx / odt）读不出正文，一律按「有中文」处理（装字体无害）；
 *   - 文本源**整份扫**（见 hasCjk）—— 只看前若干字符会让大文件末尾的中文漏过去，
 *     那时 `.md → tex` 会静默产出一份缺字的 .tex（pandoc 自己不报）。
 */
function cjkVars(source, absolute, state, options) {
    const forced = typeof options.font === 'string' ? options.font.trim() : '';
    if (forced !== '') return { vars: ['-V', `CJKmainfont=${forced}`], needed: true, font: forced };
    const font = state.cjkFont;
    if (font === undefined || font === '') return { vars: [], needed: false };
    const needed = source.media === true || hasCjk(absolute) !== false;
    return { vars: ['-V', `CJKmainfont=${font}`], needed, font };
}

/** 读得出正文的文本片段（大小有界：只给 needsThesis 这类整份判据用）。 */
function safeText(absolute) {
    try {
        return readFileSync(absolute, 'utf8').slice(0, 200_000);
    } catch {
        return undefined;
    }
}

/**
 * 整个文件里有没有 CJK：分块扫，命中即停（内存有界）。
 *
 * 分块要**重叠 3 字节**：UTF-8 的汉字是 3 字节，正好卡在块边界上时会被截成半个字符，
 * 单块解码看不见它 —— 那正是「大文件末尾的中文被判没有」的成因。
 *
 * @returns {boolean|undefined} 读不了时 undefined（调用方按「可能有」处理）
 */
function hasCjk(absolute) {
    let fd;
    try {
        fd = openSync(absolute, 'r');
        const chunk = Buffer.alloc(1 << 20);
        let offset = 0;
        for (;;) {
            const read = readSync(fd, chunk, 0, chunk.length, offset);
            if (read <= 0) return false;
            if (CJK_RE.test(chunk.subarray(0, read).toString('utf8'))) return true;
            if (read < chunk.length) return false;
            offset += read - 3; // 回退 3 字节，避免把跨块的汉字截断
        }
    } catch {
        return undefined;
    } finally {
        if (fd !== undefined) {
            try {
                closeSync(fd);
            } catch {
                // 关不上不影响结论
            }
        }
    }
}

/** 这份 .tex 是不是 thuthesis 论文（决定编译前要不要拷那套模板）。 */
function needsThesis(absolute) {
    const text = safeText(absolute);
    if (text === undefined) return true;
    return /\\documentclass\s*(\[[^\]]*\])?\s*\{\s*thuthesis\s*\}/.test(text);
}

/** 写文本格式（md / docx / tex）的 pandoc 参数。 */
function writerArgs(target, options, context) {
    if (target === 'md') {
        // 表格要**管道表**（GitHub / 编辑器都认），公式要 `$...$`：pandoc 的 markdown writer
        // 默认会挑网格表，所以这里把另外三种表扩展关掉；不关就会得到一张 ASCII 网格
        // （第五十三轮实测：同一份 docx 用默认参数出来的是 grid table）。
        return ['-t', 'markdown+pipe_tables-grid_tables-multiline_tables-simple_tables+tex_math_dollars', '--wrap=none'];
    }
    if (target === 'docx') {
        const args = [];
        const reference = typeof options.reference === 'string' ? options.reference.trim() : '';
        if (reference !== '') {
            const absolute = assertInsideRoot(reference, context.env, 'office.convert.file');
            if (!existsSync(absolute)) {
                throw new Error(`office.convert.file：reference（参考样式文档）找不到：${reference}。`
                    + '它要是一份本机的 .docx（Word 里做好样式另存为即可）。');
            }
            args.push('--reference-doc=' + absolute);
        }
        return args;
    }
    const args = ['-t', 'latex'];
    // 默认出**能直接编译**的整份：md → tex 之后紧接着 tex → pdf 是常用链路，
    // 少了导言区的片段编不了。要往自己的论文里并，传 fragment:true 拿正文片段。
    if (options.fragment !== true) args.push('-s');
    const cjk = cjkVars(context.source, context.absolute, context.state, options);
    if (cjk.needed) args.push(...cjk.vars);
    return args;
}

/**
 * 媒体落点的**两段式**：先规划（只读盘），主文件写出去之后再搬。
 *
 * 一段式（先搬图再写主文件）在目标冲突被拒时会留下孤儿：图已经躺进 `<目标名>.media/`，
 * 而主文件一个字节没写 —— 交付目录里多出一批没人引用的文件。所以把冲突检查提到主文件
 * 写出**之前**（`assertMediaWritable`），写出之后再搬（那时必然可行）。
 */
function planMedia(workDir, targetAbsolute) {
    const mediaDir = join(workDir, 'media');
    const empty = { entries: [], folder: undefined, mediaDir };
    if (!existsSync(mediaDir)) return empty;
    const outDir = dirname(targetAbsolute);
    const stem = basename(targetAbsolute).replace(/\.[^.]+$/, '');
    const folder = `${stem}.media`;
    const entries = [];
    const walk = (dir) => {
        for (const entry of readdirSync(dir, { withFileTypes: true })) {
            const from = join(dir, entry.name);
            if (entry.isDirectory()) {
                walk(from);
                continue;
            }
            if (!entry.isFile()) continue;
            const rel = from.slice(mediaDir.length + 1);
            entries.push({ from, to: join(outDir, folder, rel) });
        }
    };
    walk(mediaDir);
    return { entries, folder, mediaDir };
}

/** 主文件写出前先查媒体目标：内容不同又没给 overwrite 就当场拒绝（不留孤儿）。 */
function assertMediaWritable(env, planned, overwrite) {
    for (const entry of planned.entries) {
        if (!existsSync(entry.to)) continue;
        const current = readFileSync(entry.to);
        const next = readFileSync(entry.from);
        if (current.length === next.length && current.equals(next)) continue;
        if (overwrite !== true) {
            throw new Error(`office.convert.file：${displayPath(env.root, entry.to)} 已存在且内容不同；`
                + '换一个 out，或显式传 overwrite:true（覆盖别人的文件要明确说一声）。');
        }
    }
}

/** 主文件写完之后搬图（冲突已经查过，这里只会命中复用或按 overwrite 覆盖）。 */
function writeMedia(env, planned, overwrite) {
    const files = [];
    for (const entry of planned.entries) {
        mkdirSync(dirname(entry.to), { recursive: true });
        const written = emit(env, entry.to, readFileSync(entry.from), { overwrite });
        files.push(written.path);
    }
    return files;
}

/**
 * pandoc 写出的 md / tex 里指向媒体目录的路径：把绝对路径改写成工作目录里的相对路径。
 *
 * 分隔符要**两种都认**：pandoc 交出来的路径是混的（`…\img-1a2b/media/rId20.png` ——
 * 目录用反斜杠、文件名前那一段用斜杠），照字符串替换会整条漏掉（第五十三轮实测）。
 */
function rewriteMedia(text, mediaDir, folder) {
    const parts = mediaDir.split(/[\\/]+/).filter((part) => part !== '').map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    if (parts.length === 0) return text;
    // 容许 `file:///` 前缀（有些 pandoc 版本/平台会写成 URL）。
    const pattern = new RegExp(`(?:file:(?:/{1,3}))?${parts.join('[\\\\/]+')}`, 'g');
    // 用**函数**替换：folder 取自目标文件名，直接当替换串时 `$&` / `$1` 这类会被当成模式。
    return text.replace(pattern, () => folder);
}

/**
 * 文本格式之间的互转（pandoc 一条路）。
 *
 * @param {object} context 见 {@link convertFile} 里的组装
 */
async function convertText(context) {
    const { source, target, absolute, targetAbsolute, options, env, cache, state, timeoutMs, overwrite } = context;
    if (source.reader === undefined) {
        throw new Error(`office.convert.file：${source.id} → ${target} 这条路上没有引擎 —— pandoc 读不了 ${source.id}。`
            + (source.id === 'pdf'
                ? '要 PDF 里的文字用 office.pdf.text（带页码区间）；要图片用 office.pdf.pages 再 read_image。'
                : `能从 ${source.id} 转出的是：pdf（走宿主 LibreOffice 引擎）。`));
    }
    if (state.pandoc === undefined) {
        throw new Error(`office.convert.file：${source.id} → ${target} 要 pandoc（本机探不到）。`
            + `装一个：winget install --id JohnMacFarlane.Pandoc（或从 pandoc.org 下安装包）；`
            + `便携版/自定义位置用环境变量 ${PANDOC_ENV} 指路，装了会自动探到。`);
    }

    const work = cache.path(`convert/${workKey(absolute)}`);
    mkdirSync(work, { recursive: true });
    const tmpOut = join(work, `out${TARGET_EXT[target]}`);
    // 显式给 `-f <reader>`：别把「源是什么」交给扩展名推断（推断表一变就会静默换读入器）。
    const args = ['-f', source.reader, absolute, '-o', tmpOut, ...writerArgs(target, options, { ...context, env })];
    // 内嵌图片：只有二进制源（docx / odt）且目标是文本格式时才抽，抽到目标同级的 .media/ 里。
    const wantMedia = options.media !== false && source.media === true && (target === 'md' || target === 'tex');
    // `--extract-media=DIR` 会把图片放进 `DIR/media/…`，所以这里给**工作目录本身**：
    // 少一层 media 嵌套，抽到工作目录时也就不会出现 `<目标名>.media/media/…`。
    if (wantMedia) args.push('--extract-media=' + work);

    const run = await runPandoc(state, args, { cache, cwd: dirname(absolute), timeoutMs, label: workKey(absolute), sourcePath: absolute });
    if (run.result.code !== 0 || !existsSync(tmpOut)) throw new Error(pandocFailure(args, run, env));

    let bytes = readFileSync(tmpOut);
    const steps = [`pandoc：${basename(absolute)} → ${target}（缓存）`];
    // media:false 时图片引用会指向一个不存在的 media/ 目录 —— 说一声，别让人拿到断链的 md。
    if (!wantMedia && (target === 'md' || target === 'tex')
        && /(?:^|[(/"])media\//m.test(bytes.toString('utf8'))) {
        env.warn('源里的内嵌图片没有抽出来（media:false）：转出的文本里 media/ 引用指向不存在的目录；'
            + '要连图一起带走就别传 media:false。');
    }
    // 源里有中文却挂不上字体时要说一声：这份 .tex 编译出来中文会丢，而 pandoc 自己不报。
    if (target === 'tex') {
        const cjk = cjkVars(source, absolute, state, options);
        if (cjk.needed && cjk.vars.length === 0) {
            env.warn('源里有中文但本机没探到中文字体：这份 .tex 没挂 CJKmainfont，编译时中文会被丢掉。'
                + `装一款中文字体，或用 font 选项 / ${CJK_FONT_ENV} 点名一款本机字体名。`);
        }
    }
    let media = { files: [], folder: undefined };
    const planned = wantMedia ? planMedia(work, targetAbsolute) : { entries: [], folder: undefined };
    if (planned.entries.length > 0) {
        // 冲突要在写主文件之前查：顺序反了就会「主文件被拒、图已经落地」。
        assertMediaWritable(env, planned, overwrite);
        bytes = Buffer.from(rewriteMedia(bytes.toString('utf8'), planned.mediaDir, planned.folder), 'utf8');
    }

    const emitted = emit(env, targetAbsolute, bytes, { overwrite });
    if (planned.entries.length > 0) {
        media = { files: writeMedia(env, planned, overwrite), folder: planned.folder };
        steps.push(`抽出 ${planned.entries.length} 个内嵌图片 → ${planned.folder}/`);
    }
    const warnings = [...run.warnings];
    for (const warning of warnings) env.warn(`pandoc：${warning}`);
    const hint = media.files.length > 0
        ? `图片抽到了 ${media.folder}/（已登记为产物）；改 md 时连图一起带走，别只复制 .md。`
        : (emitted.reused ? '目标文件已经是这次转换的结果（内容一致），没有重写。' : '转换完成，可以直接交付。');
    return {
        ok: true,
        from: source.id,
        to: target,
        ...emitted,
        engine: 'pandoc',
        steps,
        ...(media.files.length > 0 ? { media: media.files } : {}),
        ...(warnings.length > 0 ? { warnings } : {}),
        hint,
    };
}

/** 出 PDF 时按 `engine` 收缩候选路线。 */
function pdfCandidates(source, engine) {
    // PDF 源只有「原样复用 / 换名留一份」一条路：engine 说什么都改不了它，
    // 所以这里单独回一句，别把人指到另一条同样走不通的路上。
    if (source.id === 'pdf') {
        if (engine === 'auto') return ['copy'];
        throw new Error('office.convert.file：PDF 源只有「原样复用 / 换名留一份」一条路（engine 不传即可）—— '
            + '换格式要先从 PDF 里取内容：文字用 office.pdf.text，画面用 office.pdf.pages 再 read_image。');
    }
    if (engine === 'auto') return source.pdf;
    if (engine === 'libreoffice') {
        for (const route of ['direct', 'docx']) if (source.pdf.includes(route)) return [route];
        throw new Error(`office.convert.file：${source.id} → pdf 走不了 LibreOffice —— 它读不了 ${source.id} 源。`
            + (source.id === 'tex' ? '论文源要用本机 TeX 编译（engine 不传就是这条）。' : ''));
    }
    if (source.id === 'tex') return ['tex'];
    if (source.pdf.includes('latex')) return ['latex'];
    throw new Error(`office.convert.file：${source.id} → pdf 走不了 TeX 路线 —— pandoc 读不了 ${source.id}。`
        + '改用 engine:\'libreoffice\'（宿主 LibreOffice 引擎直接读）。');
}

/**
 * 出 PDF（唯一的收口）：.tex 走本机 TeX，office 三件套走宿主 LibreOffice 引擎，
 * 纯文本源两条路（先转 docx 再渲染 / pandoc 直接调 xelatex）。
 */
async function convertPdf(context) {
    const { source, absolute, targetAbsolute, options, env, cache, state, timeoutMs, overwrite } = context;
    const engine = normalizeEngine(options.engine);
    const candidates = pdfCandidates(source, engine);
    const route = candidates.find((item) => routeUsable(item, state));
    if (route === undefined) {
        throw new Error(`office.convert.file：${source.id} → pdf 现在跑不了 —— ${routeMissing(candidates[0])}。`
            + 'office.convert.check() 会报出本机有哪几条路。');
    }
    if (route !== candidates[0]) env.note(`首选的 PDF 路线不可用（${routeMissing(candidates[0])}），改用 ${route}。`);

    if (route === 'copy') {
        const bytes = readFileSync(absolute);
        const emitted = emit(env, targetAbsolute, bytes, { overwrite });
        return {
            ok: true, from: source.id, to: 'pdf', ...emitted, engine: 'copy',
            steps: [`原样复用：${basename(absolute)}`],
            hint: emitted.reused ? '目标已经是这份 PDF（内容一致），没有重写。' : '这份 PDF 就是源文件本身（换名/换目录的副本）。',
        };
    }

    if (route === 'tex') {
        // TeX 会把 PDF 写在主文件旁边，而默认落点就是它 —— 也就是说「目标已经有一份 PDF」这件
        // 事必须**在编译之前**拦住：编译一跑就把那份 PDF 覆盖了，事后只能补一条 warning。
        // 目标位置由 TeX 决定（主文件同名），所以这里先按同一个规则算出它。
        const landingPath = resolve(env.root, join(dirname(absolute), `${basename(absolute).replace(/\.tex$/i, '')}.pdf`));
        const targetExisted = existsSync(targetAbsolute);
        if (targetExisted && overwrite !== true && targetAbsolute === landingPath) {
            throw new Error(`office.convert.file：${displayPath(env.root, targetAbsolute)} 已存在，`
                + '而 .tex 出 PDF 的默认落点就是 TeX 写 PDF 的位置 —— 编译会直接覆盖它。'
                + '确认要重编就传 overwrite:true；想两份都留就传 out。');
        }
        const report = await compileTex(absolute, env, {
            engine: typeof options.texEngine === 'string' && options.texEngine !== '' ? options.texEngine : undefined,
            timeoutMs,
            // 只认 'none' / 'aux'：'all' 在 latexmk 里是 `-C`，**连 PDF 一起删**（手工分支同样），
            // 而 PDF 正是这次的交付物；透传下去会先毁产物、再让后面的 readFileSync 抛裸 ENOENT。
            clean: cleanMode(options.clean, env),
            // 只有 thuthesis 论文项目才要那套模板（.cls 与两个校徽）；普通的 article /
            // ctexart 文档被顺手拷进一整套模板文件是污染，所以按 \documentclass 判一判。
            support: needsThesis(absolute) ? undefined : 'none',
        });
        if (report.ok !== true || report.pdf === null) {
            const first = report.log?.errors?.[0];
            const detail = first === undefined ? '' : `第一处：${first.file === null ? '' : `${first.file}:${String(first.line)} `}${first.message}`;
            throw new Error(`office.convert.file：LaTeX 编译没成功（引擎 ${report.texEngine}，退出码 ${String(report.exitCode)}）。`
                + `${detail === '' ? '' : detail + ' '}完整日志在 ${report.log?.path ?? '项目目录里的 .log'}。`
                + '（论文项目的排版问题看 office.tex.read 的 warnings，比读日志快。）');
        }
        const produced = resolve(env.root, report.pdf.path);
        // 清完中间产物之后再确认一次产物还在：万一哪个清理规则把它删了，这里要给人话，不是 ENOENT。
        if (!existsSync(produced)) {
            throw new Error(`office.convert.file：编译报告说成功，但产物不在（${report.pdf.path}）——`
                + '它可能被清理规则删掉了。不要传 clean，或换一个 out 再试。');
        }
        const landing = produced === targetAbsolute;
        const emitted = emit(env, targetAbsolute, readFileSync(produced), landing
            ? { force: true, existedBefore: targetExisted }
            : { overwrite });
        const warnings = [];
        if (report.log.missingChars > 0) warnings.push(`有 ${report.log.missingChars} 处字符在本机字体里找不到（PDF 里会缺字）`);
        if (landing && targetExisted) {
            // 走到这里说明 overwrite:true 是显式给的（否则上面那道闸已经拦下）—— 仍然说清做了什么。
            warnings.push(`${displayPath(env.root, targetAbsolute)} 原本有一份 PDF，已按 overwrite:true 重写。`);
        }
        for (const warning of warnings) env.warn(warning);
        const exception = landing
            ? '（.tex 路线的默认落点就是 TeX 自己写 PDF 的位置：这次一律登记为产物，reused 不适用；要给到别处就传 out。）'
            : '';
        return {
            ok: true, from: source.id, to: 'pdf', ...emitted,
            engine: report.texEngine,
            pages: report.pdf.pages,
            steps: [`${report.engine}：${basename(absolute)} → ${basename(produced)}`],
            ...(warnings.length > 0 ? { warnings } : {}),
            ...(report.log.overfull > 0 || report.log.underfull > 0
                ? { layout: { overfull: report.log.overfull, underfull: report.log.underfull } }
                : {}),
            hint: `编译成功${exception}要看排版效果就用 office.pdf.pages() 渲染成图片再 read_image 逐页看。`,
        };
    }

    if (route === 'direct') {
        const staged = await previewPdf(absolute, {}, env, cache);
        const bytes = readFileSync(resolve(env.root, staged.path));
        const emitted = emit(env, targetAbsolute, bytes, { overwrite });
        return {
            ok: true, from: source.id, to: 'pdf', ...emitted, engine: 'libreoffice',
            steps: [`LibreOffice：${basename(absolute)} → PDF`],
            hint: emitted.reused ? '目标已经是这份 PDF（内容一致），没有重写。' : '转换完成（渲染引擎与 office.preview 是同一份）。',
        };
    }

    // 纯文本源：route === 'docx' 时先 pandoc 出一份 docx 交给 LibreOffice；
    // route === 'latex' 时 pandoc 直接把 xelatex 当 PDF 引擎用。
    if (state.pandoc === undefined) {
        throw new Error(`office.convert.file：${source.id} → pdf 的这条路要 pandoc（本机探不到）；`
            + `装 pandoc 或改用别的源格式。环境变量 ${PANDOC_ENV} 可以指路。`);
    }
    const work = cache.path(`convert/${workKey(absolute)}`);
    mkdirSync(work, { recursive: true });
    const label = workKey(absolute);
    const stem = basename(absolute).replace(/\.[^.]+$/, '');

    if (route === 'docx') {
        const tmpDocx = join(work, `${stem}.docx`);
        const args = [absolute, '-o', tmpDocx];
        const run = await runPandoc(state, args, { cache, cwd: dirname(absolute), timeoutMs, label, sourcePath: absolute });
        if (run.result.code !== 0 || !existsSync(tmpDocx)) throw new Error(pandocFailure(args, run, env));
        const staged = await previewPdf(tmpDocx, {}, env, cache);
        const bytes = readFileSync(resolve(env.root, staged.path));
        const emitted = emit(env, targetAbsolute, bytes, { overwrite });
        for (const warning of run.warnings) env.warn(`pandoc：${warning}`);
        return {
            ok: true, from: source.id, to: 'pdf', ...emitted, engine: 'pandoc + libreoffice',
            steps: [`pandoc：${basename(absolute)} → docx（缓存）`, 'LibreOffice：docx → PDF'],
            ...(run.warnings.length > 0 ? { warnings: run.warnings } : {}),
            hint: emitted.reused ? '目标已经是这份 PDF（内容一致），没有重写。' : '转换完成（中间那份 docx 在缓存里，没进工作目录）。',
        };
    }

    const tmpPdf = join(work, `${stem}.pdf`);
    const cjk = cjkVars(source, absolute, state, options);
    if (cjk.needed && cjk.vars.length === 0) {
        throw new Error('office.convert.file：这条路线要用 LaTeX 出 PDF，但本机没探到中文字体；'
            + 'pandoc 找不到字体时会**静默丢字**（退出码仍是 0）。'
            + `装一款中文字体（Windows 自带微软雅黑，Linux 装 fonts-noto-cjk），或用 ${CJK_FONT_ENV} / font 选项点名一款本机字体名。`);
    }
    const args = [absolute, '-o', tmpPdf, `--pdf-engine=${state.latexName}`];
    if (cjk.needed) args.push(...cjk.vars);
    const run = await runPandoc(state, args, { cache, cwd: dirname(absolute), timeoutMs, label, sourcePath: absolute });
    if (run.result.code !== 0 || !existsSync(tmpPdf)) throw new Error(pandocFailure(args, run, env));
    const warnings = [...run.warnings];
    if (run.missing > 0) {
        warnings.push(`有 ${run.missing} 处字符在本机字体里找不到（PDF 里会缺字；换 ${CJK_FONT_ENV} 点名的字体试试）`);
    }
    for (const warning of warnings) env.warn(warning);
    const emitted = emit(env, targetAbsolute, readFileSync(tmpPdf), { overwrite });
    return {
        ok: true, from: source.id, to: 'pdf', ...emitted, engine: `pandoc + ${state.latexName}`,
        steps: [`pandoc：${basename(absolute)} → PDF（${state.latexName}）`],
        ...(warnings.length > 0 ? { warnings } : {}),
        hint: emitted.reused ? '目标已经是这份 PDF（内容一致），没有重写。' : '转换完成（走本机 TeX 排版，没经 LibreOffice）。',
    };
}

/**
 * 转一份文件。
 *
 * @param {string} filePath 源文件（相对工作目录或绝对）
 * @param {{to?: string, out?: string, overwrite?: boolean, engine?: string, timeoutMs?: number,
 *   media?: boolean, fragment?: boolean, reference?: string, font?: string, clean?: string,
 *   texEngine?: string}} [options]
 *   - `to`（必填）：md / docx / tex / pdf（别名 markdown / word / latex）
 *   - `out`：目标路径；不传就与源同目录同名换扩展名
 *   - `overwrite`：目标已存在且内容不同时才会用上（默认拒绝覆盖）
 *   - `engine`：只对出 PDF 有意义 —— auto（默认）/ libreoffice / tex（= 用本机 TeX 排版）
 *   - `media`：docx / odt → md / tex 时要不要把内嵌图片抽出来（默认抽；false 只出文本）
 *   - `fragment`：md → tex 时只要正文片段（默认出能直接编译的整份）
 *   - `reference`：md → docx 时的参考样式文档（本机 .docx）
 *   - `font`：LaTeX 路线用的中文字体名（默认自动探）
 *   - `texEngine`：`.tex` 出 PDF 时指定编译入口（latexmk / xelatex / lualatex）
 *   - `clean`：`.tex` 出 PDF 后清中间产物 —— `'aux'`（默认）或 `'none'`；`'all'` 按 `'aux'` 处理
 *     （它会连 PDF 一起删，而 PDF 是交付物）
 *   - `timeoutMs`：单次引擎调用的超时（默认 300 秒，范围 30 秒 - 30 分钟）
 * @param {object} env 文件环境（office_run 注入）
 * @param {object} cache 缓存目录
 */
export async function convertFile(filePath, options = {}, env, cache) {
    const raw = options !== null && typeof options === 'object' ? options : {};
    const absolute = resolveInput(filePath, env);
    const source = sourceEntry(absolute);
    if (source === undefined) {
        const at = absolute.lastIndexOf('.');
        const ext = at === -1 ? '（没有扩展名）' : absolute.slice(at);
        throw new Error(`office.convert.file：还不认识 ${ext} 这个源格式。`
            + `认识的扩展名：${convertExtensions().join(' / ')}（各类源能转出什么看 office.convert.check() 的矩阵）。`
            + '别的格式先用 office.word / excel / ppt / tex 生成，或 office.files.read 取文本。');
    }
    const target = normalizeTarget(raw.to);
    // engine 认不出时按 auto 走，但要说一声（静默收敛会让人以为参数生效了）。
    const engineRaw = typeof raw.engine === 'string' ? raw.engine.trim() : '';
    if (engineRaw !== '' && ENGINE_ALIASES[engineRaw.toLowerCase()] === undefined) {
        env.note(`认不出 engine:${engineRaw}，已按 auto 处理（可用：auto / libreoffice / tex）。`);
    }
    const state = engineState();
    const context = {
        source,
        target,
        absolute,
        targetAbsolute: resolveTarget(absolute, target, raw, env),
        options: raw,
        env,
        cache,
        state,
        timeoutMs: clampTimeout(raw.timeoutMs),
        overwrite: raw.overwrite === true,
    };
    if (target === 'pdf') return convertPdf(context);
    // 文本格式只有 pandoc 一条路：给了别的 engine 就说清楚，别装作听懂了。
    const engine = normalizeEngine(raw.engine);
    if (engine !== 'auto' && engine !== 'pandoc') {
        env.note(`md / docx / tex 之间的转换只有 pandoc 一条路，engine:${raw.engine} 已按 auto 处理。`);
    }
    return convertText({ ...context, engine });
}

/**
 * 本机转换能力实况（给 office.convert.check 与报错用）。
 *
 * 两条纪律与 preview.check 一致：**只报探到的入口**（能不能真跑由各自的调用报），
 * 并且把「缺什么、下一步怎么办」写进 hint，而不是只说一句不可用。
 */
export async function convertCheck() {
    const state = engineState();
    const version = await pandocVersion(state.probe);
    const matrix = SOURCES.map((source) => ({
        source: source.id,
        // 只列真能走的路，两条过滤：
        //   · 自己转自己（`to` 与源格式相同、又没给 out 时本来就该拒绝）；例外是 PDF 源 ——
        //     「换名/换目录留一份」是个真实需求；
        //   · pandoc 读不了的源（pdf / xlsx / pptx）没有文本目标的入口，只剩出 PDF。
        targets: CONVERT_TARGETS.filter((target) => {
            if (target === source.id) return source.id === 'pdf' && target === 'pdf';
            if (target !== 'pdf' && source.reader === undefined) return false;
            return true;
        }).map((target) => {
            if (target !== 'pdf') {
                return {
                    target,
                    engine: 'pandoc',
                    available: state.pandoc !== undefined,
                    ...(state.pandoc === undefined ? { reason: `pandoc 探不到（${PANDOC_ENV} 可指路）` } : {}),
                };
            }
            const candidates = source.pdf;
            const route = candidates.find((item) => routeUsable(item, state));
            return {
                target,
                engine: route ?? candidates[0],
                available: route !== undefined,
                ...(route === undefined ? { reason: routeMissing(candidates[0]) } : {}),
            };
        }),
    }));
    const missing = [];
    if (state.pandoc === undefined) missing.push(`pandoc（${PANDOC_ENV} 可指路）`);
    if (state.latexName === undefined) missing.push('TeX 发行版（xelatex / lualatex）');
    if (!state.kit) missing.push('宿主 LibreOffice 引擎');
    return {
        available: matrix.some((row) => row.targets.some((item) => item.available)),
        pandoc: state.pandoc === undefined ? null : { path: state.pandoc, version: version ?? null },
        latex: state.latexName ?? null,
        libreoffice: state.kit,
        cjkFont: state.cjkFont ?? null,
        targets: CONVERT_TARGETS,
        sources: convertSourceIds(),
        matrix,
        hint: missing.length === 0
            ? '三种引擎都在：文本互转走 pandoc，.tex 出 PDF 走 TeX，office 三件套出 PDF 走 LibreOffice。'
            : `还缺：${missing.join('；')}。缺的引擎只影响它那几条路，其余路线照常可用；`
                + '装了不用改配置（自动探测），便携版用环境变量指路。',
    };
}



