/**
 * PDF 读取：先问「里面有什么」，再决定怎么读。
 *
 * 为什么单独做一层：PDF 有两种完全不同的来源。
 *   - 文本型（导出的讲义、论文）：抽文字最省事，一次拿到全部内容。
 *   - 图像型（扫描件、手写笔记、电子纸导出）：里面根本没有字体与文本层，
 *     抽出来是空的 —— 只能把页面**渲染成图片**，再用模型的读图能力看。
 * 2026-09-22 的实测里，办公会话面对一份 ONYX 电子纸导出的 5 页手写笔记，
 * 探了 20 步 API 也没找到出口，就是因为只想着「读文件」，而那份 PDF 里
 * 一个字都没有（fonts=0 / images=10）。所以这里把两条路都做成现成的：
 *   office.pdf.info(path)   → 页数、有没有文本层、生产者、页面尺寸
 *   office.pdf.text(path)   → 抽文本（可写进文件，避免长文本挤占上下文）
 *   office.pdf.pages(path)  → 逐页渲染成 PNG/JPEG，交给 read_image
 *
 * 渲染引擎不打包进插件（保持零依赖），按机器上现有的探测：
 *   fitz（PyMuPDF）/ pdftoppm / pdftocairo（poppler）/ mutool / gswin64c。
 * 一个都没有时如实报错并给出装哪个的建议，不假装支持。
 *
 * 渲染结果按内容键落在缓存目录里：同一份 PDF、同样的 DPI 再渲染一次就是
 * **缓存命中**，不重算（这也是「每次 office_run 都把缓存清掉」必须修掉的原因）。
 *
 * @module dsh-office-mode/pdf
 */
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { closeSync, existsSync, mkdirSync, openSync, readFileSync, readdirSync, rmSync, statSync, unlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, delimiter, dirname, join } from 'node:path';
import { displayPath } from './engine/kit.js';

/** 渲染引擎 id（探测顺序即默认优先级：实测 PyMuPDF 比 poppler 快一倍以上）。 */
export const PDF_ENGINE_IDS = ['fitz', 'pdftoppm', 'pdftocairo', 'mutool', 'gs'];

/** 进程兜底超时：渲染是重活，但要有个上限，不能把会话挂住。 */
const PROCESS_TIMEOUT_MS = 180_000;
const PROBE_TIMEOUT_MS = 20_000;

/** 常见安装位置：PATH 里没有时再找这些（TeX Live 自带 poppler）。 */
function extraBinDirs() {
    const dirs = [];
    const push = (dir) => {
        if (typeof dir === 'string' && dir !== '' && existsSync(dir)) dirs.push(dir);
    };
    if (typeof process.env.DSH_OFFICE_PDF_BIN === 'string') {
        for (const dir of process.env.DSH_OFFICE_PDF_BIN.split(delimiter)) push(dir);
    }
    for (const root of ['C:\\texlive', 'C:\\Program Files', 'C:\\Program Files (x86)']) {
        if (!existsSync(root)) continue;
        let entries = [];
        try {
            entries = readdirSync(root, { withFileTypes: true });
        } catch {
            continue;
        }
        for (const entry of entries) {
            if (!entry.isDirectory()) continue;
            const name = entry.name.toLowerCase();
            if (name.startsWith('texlive')) push(join(root, entry.name, 'bin', 'win32'));
            if (name.startsWith('poppler')) push(join(root, entry.name, 'bin'));
            if (name.startsWith('gs') || name.startsWith('ghostscript')) {
                let inner = [];
                try {
                    inner = readdirSync(join(root, entry.name), { withFileTypes: true });
                } catch {
                    inner = [];
                }
                for (const sub of inner) if (sub.isDirectory()) push(join(root, entry.name, sub.name, 'bin'));
            }
            if (name.startsWith('imagemagick')) push(join(root, entry.name));
        }
    }
    return dirs;
}

/** 在 PATH 与常见安装位置里找一个可执行文件。 */
function findExecutable(names) {
    const dirs = [
        ...String(process.env.PATH ?? '').split(delimiter).filter((dir) => dir !== ''),
        ...extraBinDirs(),
    ];
    const exts = process.platform === 'win32'
        ? String(process.env.PATHEXT ?? '.EXE;.CMD;.BAT').split(';').filter((ext) => ext !== '')
        : [''];
    for (const name of names) {
        for (const dir of dirs) {
            for (const ext of exts) {
                const candidate = join(dir, `${name}${ext}`);
                try {
                    if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
                } catch {
                    // 目录不可读就跳过
                }
            }
        }
    }
    return undefined;
}

/**
 * 跑一个外部进程，**全程不用管道**。
 *
 * 受限沙箱下 Node 的 `stdio: 'pipe'` 会 `spawn EPERM`（命名管道被禁，2026-09-23
 * 实测复现），而「把 stdout / stderr 重定向到文件」走的是普通文件句柄，两种模式下
 * 都可用。所以这里统一：stdout 与 stderr 都写文件，退出码照常拿得到，报错时还能
 * 从 stderr 文件里读到引擎自己说的话。
 */
function runProcess(command, args, options = {}) {
    return new Promise((resolvePromise) => {
        const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : PROCESS_TIMEOUT_MS;
        const { outPath, errPath } = options;
        let outFd;
        let errFd;
        try {
            if (outPath !== undefined) outFd = openSync(outPath, 'w');
            if (errPath !== undefined) errFd = openSync(errPath, 'w');
        } catch (error) {
            if (outFd !== undefined) closeSync(outFd);
            if (errFd !== undefined) closeSync(errFd);
            resolvePromise({ code: null, error: `无法打开输出文件：${error.message}` });
            return;
        }
        const closeAll = () => {
            for (const fd of [outFd, errFd]) {
                if (fd === undefined) continue;
                try {
                    closeSync(fd);
                } catch {
                    // 已经关掉了
                }
            }
        };
        let child;
        try {
            child = spawn(command, args, {
                stdio: ['ignore', outFd ?? 'ignore', errFd ?? 'ignore'],
                windowsHide: true,
                // cwd 只有 LaTeX 编译用得上（latexmk 的中间产物与输出都按当前目录算）；
                // 不传就是宿主进程的工作目录，与原来完全一致。
                cwd: typeof options.cwd === 'string' && options.cwd !== '' ? options.cwd : undefined,
                // env 只有 Python 通道用得上（PYTHONUTF8 / PYTHONIOENCODING / MPLBACKEND）；
                // 不传就完整继承宿主环境，与原来完全一致。
                env: options.env !== null && typeof options.env === 'object' ? options.env : undefined,
            });
        } catch (error) {
            closeAll();
            resolvePromise({ code: null, error: `无法启动 ${basename(String(command))}：${error.message}` });
            return;
        }
        let settled = false;
        const finish = (value) => {
            if (settled) return;
            settled = true;
            clearTimeout(timer);
            closeAll();
            resolvePromise(value);
        };
        const timer = setTimeout(() => {
            try {
                child.kill('SIGKILL');
            } catch {
                // 已经退出了
            }
            finish({ code: null, timedOut: true, error: `${basename(String(command))} 超过 ${Math.round(timeoutMs / 1000)} 秒没有结束` });
        }, timeoutMs);
        child.on('error', (error) => finish({ code: null, error: `无法启动 ${basename(String(command))}：${error.message}` }));
        child.on('close', (code) => {
            let out = '';
            let err = '';
            try {
                if (outPath !== undefined) out = readFileSync(outPath, 'utf8');
            } catch {
                out = '';
            }
            try {
                if (errPath !== undefined) err = readFileSync(errPath, 'utf8');
            } catch {
                err = '';
            }
            finish({ code, out, err });
        });
    });
}

let probeResult;

/** 探测本机可用的 PDF 引擎（结果缓存，进程内只探一次）。 */
export function probePdfEngines(options = {}) {
    if (probeResult !== undefined && options.refresh !== true) return probeResult;
    probeResult = {
        bin: {
            pdftoppm: findExecutable(['pdftoppm']),
            pdftocairo: findExecutable(['pdftocairo']),
            mutool: findExecutable(['mutool']),
            gswin64c: findExecutable(['gswin64c', 'gs']),
            pdfinfo: findExecutable(['pdfinfo']),
            pdftotext: findExecutable(['pdftotext']),
            pdffonts: findExecutable(['pdffonts']),
            python: findExecutable(['python', 'python3', 'py']),
        },
        hasFitz: undefined,
    };
    return probeResult;
}

/** fitz 要靠一次真实导入来判断（装了 python 但没装 PyMuPDF 的情况很常见）。 */
async function hasFitz() {
    const probe = probePdfEngines();
    if (probe.hasFitz !== undefined) return probe.hasFitz;
    if (probe.bin.python === undefined) {
        probe.hasFitz = false;
        return false;
    }
    const result = await runProcess(probe.bin.python, ['-c', 'import fitz'], { timeoutMs: PROBE_TIMEOUT_MS });
    probe.hasFitz = result.code === 0;
    return probe.hasFitz;
}

/** 引擎可用性一览（给 office.pdf.engines() 与报错信息用）。 */
export async function pdfEngines() {
    const probe = probePdfEngines();
    const fitzOk = await hasFitz();
    const render = [];
    if (fitzOk) render.push('fitz');
    for (const id of ['pdftoppm', 'pdftocairo', 'mutool']) if (probe.bin[id] !== undefined) render.push(id);
    if (probe.bin.gswin64c !== undefined) render.push('gs');
    const text = [];
    if (probe.bin.pdftotext !== undefined) text.push('pdftotext');
    if (fitzOk) text.push('fitz');
    const info = [];
    if (probe.bin.pdfinfo !== undefined) info.push('pdfinfo');
    if (fitzOk) info.push('fitz');
    info.push('light');
    const fonts = [];
    if (probe.bin.pdffonts !== undefined) fonts.push('pdffonts');
    return {
        render,
        text,
        info,
        fonts,
        paths: Object.fromEntries(Object.entries(probe.bin).filter(([, value]) => value !== undefined)),
    };
}

/** 挑一个能渲染的引擎；`requested` 不是 auto 时必须是它。 */
async function pickRenderEngine(requested) {
    const available = (await pdfEngines()).render;
    const wanted = typeof requested === 'string' && requested !== '' ? requested : 'auto';
    if (wanted !== 'auto') {
        const id = wanted === 'gswin64c' ? 'gs' : wanted;
        if (!available.includes(id)) {
            throw new Error(`指定的 PDF 渲染引擎 ${wanted} 在本机不可用；可用：${available.length > 0 ? available.join(' / ') : '（无）'}`);
        }
        return id;
    }
    for (const id of PDF_ENGINE_IDS) if (available.includes(id)) return id;
    throw new Error(
        '本机没有可用的 PDF 渲染引擎，无法把页面变成图片。装一个即可（任选）：'
        + 'poppler（pdftoppm）/ Python 的 PyMuPDF（pip install pymupdf）/ MuPDF（mutool）/ Ghostscript。'
        + '装了之后不用改配置，插件会自动探测到。',
    );
}

/** 临时文件位置：优先放缓存目录，没给缓存就放系统临时目录。 */
function tempFile(cache, name) {
    if (cache !== undefined && cache !== null && typeof cache.path === 'function') {
        return cache.path(`pdf/${name}`);
    }
    return join(tmpdir(), `office-pdf-${process.pid}-${name}`);
}

/** PyMuPDF 的渲染脚本：路径走 argv，脚本本身是常量，不拼字符串。 */
const FITZ_RENDER_SCRIPT = [
    'import sys',
    'import fitz',
    'src, outdir, prefix, first, last, dpi, fmt = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]), int(sys.argv[5]), int(sys.argv[6]), sys.argv[7]',
    'doc = fitz.open(src)',
    'if doc.needs_pass:',
    "    sys.stderr.write('encrypted')",
    '    raise SystemExit(3)',
    'last = min(last, doc.page_count)',
    'ext = "jpg" if fmt == "jpeg" else "png"',
    'for n in range(first, last + 1):',
    '    pix = doc[n - 1].get_pixmap(dpi=dpi, alpha=False)',
    '    data = pix.tobytes("jpeg" if fmt == "jpeg" else "png")',
    '    with open("%s/%s-%d.%s" % (outdir, prefix, n, ext), "wb") as fh:',
    '        fh.write(data)',
].join('\n');

/** 各引擎的渲染命令行（输出统一成 `<prefix>-<页码>.<ext>`）。 */
function renderCommand(engineId, context) {
    const probe = probePdfEngines();
    const { src, outDir, prefix, first, last, dpi, format } = context;
    const ext = format === 'jpeg' ? 'jpg' : 'png';
    if (engineId === 'fitz') {
        return {
            command: probe.bin.python,
            args: ['-c', FITZ_RENDER_SCRIPT, src, outDir, prefix, String(first), String(last), String(dpi), format],
        };
    }
    if (engineId === 'pdftoppm') {
        return {
            command: probe.bin.pdftoppm,
            args: [format === 'jpeg' ? '-jpeg' : '-png', '-r', String(dpi), '-f', String(first), '-l', String(last), src, join(outDir, prefix)],
        };
    }
    if (engineId === 'pdftocairo') {
        return {
            command: probe.bin.pdftocairo,
            args: [format === 'jpeg' ? '-jpeg' : '-png', '-r', String(dpi), '-f', String(first), '-l', String(last), src, join(outDir, prefix)],
        };
    }
    if (engineId === 'mutool') {
        return {
            command: probe.bin.mutool,
            args: ['draw', '-q', '-r', String(dpi), '-o', join(outDir, `${prefix}-%d.${ext}`), src, `${first}-${last}`],
        };
    }
    if (engineId === 'gs') {
        return {
            command: probe.bin.gswin64c,
            args: [
                '-q', '-dNOPAUSE', '-dBATCH', '-dSAFER',
                `-sDEVICE=${format === 'jpeg' ? 'jpeg' : 'png16m'}`,
                `-r${dpi}`,
                `-dFirstPage=${first}`,
                `-dLastPage=${last}`,
                `-sOutputFile=${join(outDir, `${prefix}-%d.${ext}`)}`,
                src,
            ],
        };
    }
    throw new Error(`未知的 PDF 渲染引擎：${engineId}`);
}

/* ── PDF 元信息（轻量解析，不依赖外部工具） ─────────────────────────────── */

/** 解一个 PDF 字符串字面量或十六进制串（够用就好，不追求覆盖全部转义）。 */
function decodePdfString(raw) {
    const text = String(raw ?? '');
    if (text.startsWith('<') && text.endsWith('>')) {
        const hex = text.slice(1, -1).replace(/[^0-9A-Fa-f]/g, '');
        const bytes = Buffer.from(hex.length % 2 === 0 ? hex : `${hex}0`, 'hex');
        if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
            // UTF-16BE：Node 只有 utf16le，先把字节对换过来（copy 一份，swap16 是原地改）。
            const body = Buffer.from(bytes.subarray(2, bytes.length - ((bytes.length - 2) % 2)));
            body.swap16();
            return body.toString('utf16le').replace(/\u0000/g, '');
        }
        return bytes.toString('latin1');
    }
    const body = text.startsWith('(') && text.endsWith(')') ? text.slice(1, -1) : text;
    const decoded = body
        .replace(/\\([nrtbf()\\])/g, (_all, ch) => ({ n: '\n', r: '\r', t: '\t', b: '\b', f: '\f' }[ch] ?? ch))
        .replace(/\\([0-7]{1,3})/g, (_all, oct) => String.fromCharCode(Number.parseInt(oct, 8)));
    // 字面量串里也常见 UTF-16BE（前两个字节是 BOM）：PDF 允许括号串直接放 UTF-16 字节。
    if (decoded.length >= 2 && decoded.charCodeAt(0) === 0xfe && decoded.charCodeAt(1) === 0xff) {
        const buffer = Buffer.alloc(decoded.length);
        for (let index = 0; index < decoded.length; index += 1) buffer[index] = decoded.charCodeAt(index) & 0xff;
        const payload = Buffer.from(buffer.subarray(2, buffer.length - ((buffer.length - 2) % 2)));
        payload.swap16();
        return payload.toString('utf16le').replace(/\u0000/g, '');
    }
    return decoded;
}

/** 从原始字节里找 `/(Title|Producer|Creator|Author) (…)`，取第一次出现。 */
function scanPdfStrings(raw) {
    const out = {};
    for (const key of ['Title', 'Producer', 'Creator', 'Author']) {
        const pattern = new RegExp(`/${key}\\s*(\\((?:\\\\.|[^\\\\()])*\\)|<[0-9A-Fa-f\\s]*>)`);
        const match = pattern.exec(raw);
        if (match === null) continue;
        const value = decodePdfString(match[1]).trim();
        if (value !== '') out[key.toLowerCase()] = value;
    }
    return out;
}

/** 轻量解析：页数（退化时用）、字体/图像计数、是否加密、页面尺寸。 */
function lightParse(bytes) {
    const raw = bytes.toString('latin1');
    const count = (re) => (raw.match(re) ?? []).length;
    const counts = [...raw.matchAll(/\/Count\s+(\d+)/g)].map((match) => Number.parseInt(match[1], 10));
    const mediaBox = /\/MediaBox\s*\[\s*([\d.+-]+)\s+([\d.+-]+)\s+([\d.+-]+)\s+([\d.+-]+)\s*\]/.exec(raw);
    let pageSizeCm = null;
    if (mediaBox !== null) {
        const width = Number.parseFloat(mediaBox[3]) - Number.parseFloat(mediaBox[1]);
        const height = Number.parseFloat(mediaBox[4]) - Number.parseFloat(mediaBox[2]);
        if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
            pageSizeCm = { width: Number((width / 28.3465).toFixed(2)), height: Number((height / 28.3465).toFixed(2)) };
        }
    }
    return {
        pagesByCount: counts.length > 0 ? Math.max(...counts) : 0,
        pagesByDict: count(/\/Type\s*\/Page(?![s])/g),
        fonts: count(/\/Type\s*\/Font/g),
        images: count(/\/Subtype\s*\/Image/g),
        encrypted: /\/Encrypt\b/.test(raw),
        pageSizeCm,
        meta: scanPdfStrings(raw),
    };
}

/** pdfinfo 的 `Key: value` 输出 → 对象（只取 ASCII 可靠的字段）。 */
function parsePdfinfo(text) {
    const out = {};
    for (const line of String(text ?? '').split(/\r?\n/)) {
        const at = line.indexOf(':');
        if (at <= 0) continue;
        const key = line.slice(0, at).trim().toLowerCase();
        const value = line.slice(at + 1).trim();
        if (value === '') continue;
        out[key] = value;
    }
    return out;
}

/**
 * pdffonts 的输出 → 字体名列表。
 * 表头之后是 `名字 类型 编码 …` 的定宽表；分隔线之前的内容都要丢掉。
 */
function parsePdffonts(text) {
    const fonts = [];
    let started = false;
    for (const line of String(text ?? '').split(/\r?\n/)) {
        if (/^\s*-{5,}/.test(line)) {
            started = true;
            continue;
        }
        if (!started || line.trim() === '') continue;
        const name = line.trim().split(/\s{2,}/)[0];
        if (name !== '') fonts.push(name);
    }
    return fonts;
}

/** fitz 的一行信息：`pages needs_pass producer creator title`（制表符分隔）。 */
const FITZ_INFO_SCRIPT = [
    'import sys',
    'import fitz',
    'doc = fitz.open(sys.argv[1])',
    'meta = doc.metadata or {}',
    'print("\\t".join([str(doc.page_count), "1" if doc.needs_pass else "0", (meta.get("producer") or ""), (meta.get("creator") or ""), (meta.get("title") or "")]))',
].join('\n');

/**
 * 这份 PDF 里有什么。
 *
 * @param {string} filePath 相对工作目录或绝对路径
 * @param {object} env 文件环境（来自 engine/kit）
 * @param {{bytes?: Buffer, cache?: object}} [options] bytes 已读进来的字节；cache 用来放临时文件
 */
export async function pdfInfo(filePath, env, options = {}) {
    const absolute = env.resolve(filePath);
    if (!existsSync(absolute)) throw new Error(`找不到文件：${filePath}`);
    const bytes = options.bytes ?? readFileSync(absolute);
    if (!(bytes.length >= 5 && bytes.subarray(0, 5).toString('latin1') === '%PDF-')) {
        throw new Error(`不是 PDF 文件（缺少 %PDF- 头）：${filePath}`);
    }
    const light = lightParse(bytes);
    const engines = await pdfEngines();
    const probe = probePdfEngines();

    let pages = 0;
    let pagesSource = 'light';
    let producer = light.meta.producer;
    let creator = light.meta.creator;
    let title = light.meta.title;
    let encrypted = light.encrypted;

    if (engines.info.includes('pdfinfo')) {
        const outPath = tempFile(options.cache, `info-${Date.now().toString(36)}.txt`);
        mkdirSync(dirname(outPath), { recursive: true });
        const result = await runProcess(probe.bin.pdfinfo, [absolute], { outPath, timeoutMs: PROBE_TIMEOUT_MS });
        if (result.code === 0) {
            const fields = parsePdfinfo(result.out);
            const parsedPages = Number.parseInt(fields.pages ?? '', 10);
            if (Number.isFinite(parsedPages) && parsedPages > 0) {
                pages = parsedPages;
                pagesSource = 'pdfinfo';
            }
            if (fields.encrypted === 'yes') encrypted = true;
            // pdfinfo 的中文标题会按控制台代码页输出成乱码，所以元信息仍以轻量解析为准。
            if (producer === undefined && fields.producer !== undefined) producer = fields.producer;
            if (creator === undefined && fields.creator !== undefined) creator = fields.creator;
        }
        try {
            unlinkSync(outPath);
        } catch {
            // 清不掉也无所谓，prune 会收走
        }
    }

    if (pages === 0 && engines.info.includes('fitz')) {
        const outPath = tempFile(options.cache, `info-fitz-${Date.now().toString(36)}.txt`);
        mkdirSync(dirname(outPath), { recursive: true });
        const result = await runProcess(probe.bin.python, ['-c', FITZ_INFO_SCRIPT, absolute], { outPath, timeoutMs: PROBE_TIMEOUT_MS });
        if (result.code === 0) {
            const [rawPages, rawPass, rawProducer, rawCreator, rawTitle] = String(result.out).trim().split('\t');
            const parsedPages = Number.parseInt(rawPages ?? '', 10);
            if (Number.isFinite(parsedPages) && parsedPages > 0) {
                pages = parsedPages;
                pagesSource = 'fitz';
            }
            if (rawPass === '1') encrypted = true;
            if (producer === undefined && rawProducer !== undefined && rawProducer !== '') producer = rawProducer;
            if (creator === undefined && rawCreator !== undefined && rawCreator !== '') creator = rawCreator;
            if (title === undefined && rawTitle !== undefined && rawTitle !== '') title = rawTitle;
        }
        try {
            unlinkSync(outPath);
        } catch {
            // 同上
        }
    }

    if (pages === 0) {
        pages = light.pagesByCount > 0 ? light.pagesByCount : light.pagesByDict;
        pagesSource = 'light';
    }

    // 有没有文本层：先看轻量解析。TeX 这类工具会把所有对象压进对象流
    // （/ObjStm），`/Type /Font` 在原始字节里根本不出现 —— 于是 xelatex 编出来的
    // 纯文本 PDF 会被误判成扫描件，模型就会去渲染图片而不是抽文字。
    // 所以轻量解析数到 0 时，再用 pdffonts 判一次（探测不到就维持原判）。
    let fonts = light.fonts;
    let fontNames = [];
    if (fonts === 0 && engines.fonts.includes('pdffonts')) {
        const outPath = tempFile(options.cache, `fonts-${Date.now().toString(36)}.txt`);
        mkdirSync(dirname(outPath), { recursive: true });
        const result = await runProcess(probe.bin.pdffonts, [absolute], { outPath, timeoutMs: PROBE_TIMEOUT_MS });
        if (result.code === 0) {
            fontNames = parsePdffonts(result.out);
            fonts = fontNames.length;
        }
        try {
            unlinkSync(outPath);
        } catch {
            // 清不掉也无所谓，prune 会收走
        }
    }
    const hasTextLayer = fonts > 0;
    return {
        path: displayPath(env.root, absolute),
        bytes: bytes.length,
        pages,
        pagesSource,
        pageSizeCm: light.pageSizeCm,
        title: title ?? null,
        producer: producer ?? null,
        creator: creator ?? null,
        encrypted,
        hasTextLayer,
        counts: { fonts, images: light.images },
        fonts: fontNames.slice(0, 12),
        hint: hasTextLayer
            ? '有字体信息：先 office.pdf.text() 抽文字'
            : '无文本层（扫描/手写）：用 office.pdf.pages() 渲染成图片，再 read_image 看',
    };
}

/* ── 抽文本 ─────────────────────────────────────────────────────────────── */

const FITZ_TEXT_SCRIPT = [
    'import sys',
    'import fitz',
    'src, out, first, last = sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4])',
    'doc = fitz.open(src)',
    'last = min(last, doc.page_count)',
    'parts = []',
    'for n in range(first, last + 1):',
    '    parts.append(doc[n - 1].get_text("text"))',
    'with open(out, "w", encoding="utf-8") as fh:',
    '    fh.write("\\n".join(parts))',
].join('\n');

/**
 * 抽文本。文本型 PDF 走这条路；图像型会返回空（这时要改走 pages）。
 *
 * @param {string} filePath
 * @param {{from?: number, to?: number, out?: string}} [options] out 给出时把文本写进文件
 * @param {object} env
 * @param {object} [cache]
 */
export async function pdfText(filePath, options = {}, env, cache) {
    const absolute = env.resolve(filePath);
    if (!existsSync(absolute)) throw new Error(`找不到文件：${filePath}`);
    const info = await pdfInfo(filePath, env, { cache });
    const from = Math.max(1, Math.trunc(Number(options.from) || 1));
    const to = Math.min(info.pages > 0 ? info.pages : from, Math.trunc(Number(options.to) || info.pages || from));
    if (to < from) throw new Error(`页码范围不合法：from=${from} to=${to}`);
    if (info.encrypted) throw new Error('这份 PDF 加密了，抽不出文本；请先解密。');

    const engines = await pdfEngines();
    const probe = probePdfEngines();
    const tmpPath = tempFile(cache, `text-${Date.now().toString(36)}.txt`);
    mkdirSync(dirname(tmpPath), { recursive: true });
    let engine = null;
    let result;
    if (engines.text.includes('pdftotext')) {
        result = await runProcess(probe.bin.pdftotext, ['-enc', 'UTF-8', '-f', String(from), '-l', String(to), absolute, tmpPath], { timeoutMs: PROCESS_TIMEOUT_MS });
        if (result.code === 0) engine = 'pdftotext';
    }
    if (engine === null && engines.text.includes('fitz')) {
        result = await runProcess(probe.bin.python, ['-c', FITZ_TEXT_SCRIPT, absolute, tmpPath, String(from), String(to)], { timeoutMs: PROCESS_TIMEOUT_MS });
        if (result.code === 0) engine = 'fitz';
    }
    if (engine === null) {
        // 两个回退都要落到「未知原因」：`result?.err` 为空时 split(...)[0] 是空串，
        // 而空串不是 nullish —— 原先那句 `?? '未知原因'` 永远不生效，消息里会出现
        // 一个悬空的「：」（审查 P2-5）。
        const stderr = String(result?.err ?? '').trim().split(/\r?\n/)[0];
        const reason = result?.error || stderr || '未知原因';
        throw new Error(`抽文本失败（可用引擎：${engines.text.join(' / ') || '无'}）：${reason}`);
    }

    let text = '';
    try {
        text = readFileSync(tmpPath, 'utf8');
    } finally {
        try {
            unlinkSync(tmpPath);
        } catch {
            // 清不掉也无所谓
        }
    }
    const chars = text.replace(/\s+/g, '').length;
    const emptyHint = '这份 PDF 没有文本层（扫描/手写）：改用 office.pdf.pages() 渲染成图片，再 read_image 看。';
    if (typeof options.out === 'string' && options.out.trim() !== '') {
        const written = env.writeFile(options.out, text);
        const value = { engine, from, to, pages: to - from + 1, chars, path: written.path, head: text.slice(0, 200) };
        if (chars === 0) value.hint = emptyHint;
        return value;
    }
    const value = { engine, from, to, pages: to - from + 1, chars, text };
    value.hint = chars === 0
        ? emptyHint
        : '长文本不要靠返回值带回（工具反馈会截断）：传 out 参数写进文件，再用 read 读。';
    return value;
}

/* ── 渲染成图片（跨调用可命中缓存） ─────────────────────────────────────── */

/**
 * 把指定页渲染成图片，落在缓存目录里，返回可直接交给 read_image 的路径。
 *
 * 同一份 PDF + 同样的 DPI / 格式 / 引擎 → 内容键相同 → 已有文件直接复用，
 * 这就是「缓存命中」。所以缓存不能被每次调用清空。
 *
 * @param {string} filePath
 * @param {{from?: number, to?: number, dpi?: number, format?: 'png'|'jpeg'}} [options]
 * @param {object} env
 * @param {object} cache engine/cache 实例
 * @param {object} config resolved 配置（pdfDpi / pdfMaxPages / pdfEngine）
 */
export async function pdfPages(filePath, options = {}, env, cache, config = {}) {
    const absolute = env.resolve(filePath);
    if (!existsSync(absolute)) throw new Error(`找不到文件：${filePath}`);
    const bytes = readFileSync(absolute);
    const info = await pdfInfo(filePath, env, { bytes, cache });
    if (info.pages <= 0) throw new Error(`读不出页数，无法渲染：${filePath}`);
    if (info.encrypted) throw new Error('这份 PDF 加密了，渲染会失败；请先解密。');

    const maxPages = Number(config.pdfMaxPages) > 0 ? Number(config.pdfMaxPages) : 20;
    const dpi = Math.min(400, Math.max(36, Math.trunc(Number(options.dpi) || Number(config.pdfDpi) || 120)));
    const format = String(options.format ?? 'png').toLowerCase() === 'jpeg' ? 'jpeg' : 'png';
    const from = Math.max(1, Math.trunc(Number(options.from) || 1));
    const to = Math.min(info.pages, Math.trunc(Number(options.to) || info.pages));
    if (to < from) throw new Error(`页码范围不合法：from=${from} to=${to}（共 ${info.pages} 页）`);
    const wanted = to - from + 1;
    if (wanted > maxPages) {
        throw new Error(
            `一次最多渲染 ${maxPages} 页，这次要 ${wanted} 页（${from}-${to}，共 ${info.pages} 页）。`
            + `请分批传 from/to（例如 from: ${from}, to: ${Math.min(to, from + maxPages - 1)}）。`,
        );
    }

    const engine = await pickRenderEngine(config.pdfEngine);
    const key = createHash('sha1').update(bytes).digest('hex').slice(0, 12);
    const stem = (basename(absolute).replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]/g, '-').slice(0, 24) || 'doc');
    // 内容键 = 文件内容 + dpi + 格式 + 引擎：同样的参数才会命中，换引擎会重渲染
    // （不同引擎的光栅化结果不完全一样，混用会把「换引擎看看」变成假象）。
    const subdir = `pdf/${stem}-${key}-${dpi}-${format}-${engine}`;
    const outDir = cache.ensureDir(subdir);
    const ext = format === 'jpeg' ? 'jpg' : 'png';
    const width = String(info.pages).length;

    const files = [];
    const missing = [];
    for (let page = from; page <= to; page += 1) {
        const name = `${stem}-p${String(page).padStart(width, '0')}.${ext}`;
        const target = join(outDir, name);
        let size = 0;
        try {
            size = statSync(target).size;
        } catch {
            size = 0;
        }
        if (size > 0) {
            cache.noteHit(size);
            files.push({ page, path: displayPath(env.root, target), bytes: size, reused: true });
            continue;
        }
        missing.push(page);
    }

    let rendered = 0;
    if (missing.length > 0) {
        const firstMissing = missing[0];
        const lastMissing = missing[missing.length - 1];
        const prefix = `_tmp-${process.pid}-${Date.now().toString(36)}`;
        const errPath = join(outDir, `${prefix}.err.txt`);
        const { command, args } = renderCommand(engine, { src: absolute, outDir, prefix, first: firstMissing, last: lastMissing, dpi, format });
        const result = await runProcess(command, args, { errPath, timeoutMs: PROCESS_TIMEOUT_MS });
        const produced = [];
        try {
            for (const entry of readdirSync(outDir, { withFileTypes: true })) {
                if (!entry.isFile()) continue;
                const match = new RegExp(`^${prefix}-(\\d+)\\.(png|jpg|jpeg)$`).exec(entry.name);
                if (match === null) continue;
                produced.push({ page: Number.parseInt(match[1], 10), name: entry.name });
            }
        } catch {
            // 目录读不了就当作没产出
        }
        if (result.code !== 0 && produced.length === 0) {
            let reason = String(result.err ?? '').trim().split(/\r?\n/).filter((line) => line !== '').slice(-2).join(' ');
            if (reason === '') reason = result.error ?? `退出码 ${String(result.code)}`;
            throw new Error(`渲染失败（引擎 ${engine}）：${reason}`);
        }
        for (const item of produced) {
            const targetName = `${stem}-p${String(item.page).padStart(width, '0')}.${ext}`;
            const target = join(outDir, targetName);
            try {
                const data = readFileSync(join(outDir, item.name));
                writeFileSync(target, data);
                rmSync(join(outDir, item.name), { force: true });
            } catch (error) {
                throw new Error(`渲染结果落盘失败：${error.message}`);
            }
            cache.noteArtifact();
            rendered += 1;
            files.push({ page: item.page, path: displayPath(env.root, target), bytes: statSync(target).size, reused: false });
        }
        try {
            rmSync(errPath, { force: true });
        } catch {
            // 无所谓
        }
    }

    files.sort((a, b) => a.page - b.page);
    return {
        engine,
        dpi,
        format,
        from,
        to,
        pages: info.pages,
        dir: displayPath(env.root, outDir),
        files,
        reused: files.filter((file) => file.reused).length,
        rendered,
        bytes: files.reduce((sum, file) => sum + file.bytes, 0),
        hint: '图片在缓存目录里，直接 read_image({ file_path }) 逐页看；同样的页码再渲染会命中缓存。',
    };
}

/** 一份 PDF 是不是「图片型」（没有文本层）——反馈里给建议用。 */
export function isImageOnlyPdf(info) {
    return info !== null && typeof info === 'object' && info.hasTextLayer !== true;
}

export { runProcess as runPdfProcess, findExecutable as findPdfExecutable };
