/**
 * 归档与会话日志的读取（office.archive）。
 *
 * 为什么单独做一层：会话导出（`session*.zip`、`*.jsonl.zst`）与各种日志包
 * 在办公模式里以前是**读不出**的 —— `office.files.read()` 按 UTF-8 读一份 ZIP
 * 只会得到一段乱码，模型既不知道那是压缩包，也不知道该换哪条路。第二十三轮
 * 记下的 `23-5` 说的就是这件事。
 *
 * 三条口径：
 *   1. **按魔数判型，不按扩展名**。`.zip` 里的内容可能是纯文本，反过来一份没有
 *      扩展名的文件可能是 zstd —— 扩展名只用来兜底。
 *   2. **零第三方依赖**。ZIP 复用引擎里已有的 `unzip()`，gzip / zstd 走 `node:zlib`
 *      （zstd 需要 Node 22.15+；本机 Node 没有那个函数时明确报「本机不支持」，
 *      不退化成乱码）。
 *   3. **返回有界**。会话日志动辄几 MB、上万行，`text()` 默认只回一段并支持
 *      `offset`，要在里面找东西用 `find()`；要把整份落到盘上用 `extract()`。
 *
 * @module dsh-office-mode/archive
 */
import { gunzipSync, zstdDecompressSync } from 'node:zlib';
import { ZIP_MAX_ENTRY_BYTES, crc32, unzipMeta, zipEntryCount } from './engine/zip.js';
import { assertInsideRoot } from './engine/kit.js';

/** 判出来的容器类型。 */
export const ARCHIVE_KINDS = ['zip', 'gzip', 'zstd', 'plain'];

const MAGIC_ZIP = [0x50, 0x4b, 0x03, 0x04];
const MAGIC_ZIP_EMPTY = [0x50, 0x4b, 0x05, 0x06];
const MAGIC_GZIP = [0x1f, 0x8b];
const MAGIC_ZSTD = [0x28, 0xb5, 0x2f, 0xfd];

function startsWith(bytes, magic) {
    if (bytes.length < magic.length) return false;
    for (let i = 0; i < magic.length; i += 1) {
        if (bytes[i] !== magic[i]) return false;
    }
    return true;
}

/**
 * 按魔数判容器类型；认不出来时按「是不是像文本」分成 `plain` 或 `undefined`。
 * @param {Uint8Array} bytes
 * @returns {'zip'|'gzip'|'zstd'|'plain'|undefined}
 */
export function sniffArchive(bytes) {
    if (!(bytes instanceof Uint8Array) || bytes.length === 0) return undefined;
    if (startsWith(bytes, MAGIC_ZIP) || startsWith(bytes, MAGIC_ZIP_EMPTY)) return 'zip';
    if (startsWith(bytes, MAGIC_GZIP)) return 'gzip';
    if (startsWith(bytes, MAGIC_ZSTD)) return 'zstd';
    // 文本判定：前 4KB 里没有 NUL、且没有明显的二进制控制字符。
    const head = bytes.subarray(0, 4096);
    let control = 0;
    for (let i = 0; i < head.length; i += 1) {
        const b = head[i];
        if (b === 0) return undefined;
        // 允许 \t \n \r；其余 C0 控制字符按二进制算。
        if (b < 0x09 || (b > 0x0d && b < 0x20)) control += 1;
    }
    return control / head.length > 0.05 ? undefined : 'plain';
}

const decoder = new TextDecoder('utf-8', { fatal: false });

function decodeText(bytes) {
    return decoder.decode(bytes);
}

/** 一段文本里有没有 NUL —— 用来判断「这份内容当文本读是否合理」。 */
function looksBinary(bytes) {
    const head = bytes.subarray(0, 4096);
    for (let i = 0; i < head.length; i += 1) {
        if (head[i] === 0) return true;
    }
    return false;
}

function hasZstd() {
    return typeof zstdDecompressSync === 'function';
}

/**
 * 单条正文的字节上限（与 zip.js 的同一口径）：挡「一读就把内存撑爆」的极端包。
 * ZIP 在 `unzipMeta` 里按中央目录的**声明值**先挡；gzip / zstd 用 zlib 的
 * `maxOutputLength` 挡 —— 三种容器都在解压前/解压中封顶，不是解完再判。
 */
const MAX_ENTRY_BYTES = ZIP_MAX_ENTRY_BYTES;

/** 多帧 zstd 的检测：把几个 zstd 流拼起来时，第二个帧头出现在正文里。 */
function countZstdFrames(bytes) {
    let count = 0;
    for (let i = 0; i + 4 <= bytes.length; i += 1) {
        if (bytes[i] === 0x28 && bytes[i + 1] === 0xb5 && bytes[i + 2] === 0x2f && bytes[i + 3] === 0xfd) count += 1;
    }
    return count;
}

/**
 * 建一个「按名字取正文」的读取器。ZIP 走中央目录，gzip / zstd 只有一条正文。
 * @returns {{kind: string, names: string[], list: () => Array<{name: string, bytes: number, method: string}>, at: (name: string) => Uint8Array|undefined, note?: string}}
 */
function openContainer(kind, bytes) {
    if (kind === 'zip') {
        let parsed;
        try {
            parsed = unzipMeta(bytes, { maxEntryBytes: MAX_ENTRY_BYTES });
        } catch (error) {
            // zip.js 的消息是给解析器用的（「找不到中央目录结尾」）；裹上这一层的口径，
            // 调用方才知道是谁在读、以及坏了还是超限了。
            throw new Error(`office.archive：ZIP 解析失败（${error instanceof Error ? error.message : String(error)}）`);
        }
        const { files, entries } = parsed;
        const byName = new Map(entries.map((entry) => [entry.name, entry]));
        return {
            kind,
            names: [...files.keys()],
            // 压缩方式如实报：stored（0）与 deflate（8）是两回事，写死 'deflate' 等于报错。
            list: () => [...files.keys()].map((name) => ({
                name,
                bytes: files.get(name).length,
                method: byName.get(name)?.method === 0 ? 'stored' : 'deflate',
            })),
            at: (name) => files.get(name),
        };
    }
    if (kind === 'gzip') {
        let data;
        try {
            data = new Uint8Array(gunzipSync(bytes, { maxOutputLength: MAX_ENTRY_BYTES }));
        } catch (error) {
            // 与 zstd 那条一样：把 zlib 的内部消息裹上这一层的上下文，
            // 别让调用方看到一句「invalid stored block lengths」而不知道是谁在报。
            throw new Error(`office.archive：gzip 解压失败（${error instanceof Error ? error.message : String(error)}）。`
                + '这份包可能被截断或损坏；也可以用外部工具解开再读。');
        }
        return {
            kind,
            names: [''],
            list: () => [{ name: '', bytes: data.length, method: 'gzip' }],
            at: () => data,
        };
    }
    if (kind === 'zstd') {
        if (!hasZstd()) {
            throw new Error('office.archive：本机 Node 不支持 zstd（zlib.zstdDecompressSync 不存在，需要 Node 22.15 以上）。'
                + '可以先用外部工具解压成文本，再用 office.files.read 读。');
        }
        let data;
        try {
            data = new Uint8Array(zstdDecompressSync(bytes, { maxOutputLength: MAX_ENTRY_BYTES }));
        } catch (error) {
            throw new Error(`office.archive：zstd 解压失败（${error instanceof Error ? error.message : String(error)}）。`
                + '这份包可能被截断、损坏，或由多个帧拼成（Node 只解第一帧）。');
        }
        // 多帧拼接时 Node 只解第一帧，而且**不报错** —— 静默少读一半比报错更糟，
        // 所以这里扫一遍帧头，把「这份文件里有 N 帧、只解了第一帧」如实说出来。
        const frames = countZstdFrames(bytes);
        return {
            kind,
            names: [''],
            list: () => [{ name: '', bytes: data.length, method: 'zstd' }],
            at: () => data,
            ...(frames > 1
                ? { note: `这份 zstd 里数到 ${frames} 个帧头，Node 只解第一帧（${data.length} 字节）；后面的帧没有读进来。` }
                : {}),
        };
    }
    return {
        kind: 'plain',
        names: [''],
        list: () => [{ name: '', bytes: bytes.length, method: 'none' }],
        at: () => bytes,
    };
}

function readBytes(filePath, env) {
    const absolute = env.resolve(filePath);
    if (!env.exists(absolute)) throw new Error(`office.archive：找不到文件 ${filePath}`);
    return env.readFile(absolute);
}

/**
 * 这份归档里有什么。
 * @param {string} filePath 相对工作目录或绝对路径
 * @param {object} env 引擎环境
 * @param {{limit?: number}} [options]
 */
export function archiveInfo(filePath, env, options = {}) {
    const bytes = readBytes(filePath, env);
    const kind = sniffArchive(bytes);
    const limit = Math.max(1, Math.min(200, Math.trunc(Number(options.limit) || 50)));
    if (kind === undefined) {
        return {
            path: filePath,
            bytes: bytes.length,
            kind: 'binary',
            entries: 0,
            note: '认不出来这是什么容器（不是 ZIP / gzip / zstd，也不像文本）：按二进制处理，本插件不解析它。',
        };
    }
    const container = openContainer(kind, bytes);
    const list = container.list();
    const total = list.reduce((sum, item) => sum + item.bytes, 0);
    const textEntries = list.filter((item) => /\.(jsonl?|txt|md|log|csv|tsv|xml|yaml|yml)$/i.test(item.name));
    return {
        path: filePath,
        bytes: bytes.length,
        kind,
        entries: list.length,
        uncompressedBytes: total,
        ratio: bytes.length > 0 ? Number((total / bytes.length).toFixed(2)) : 0,
        textEntries: textEntries.length,
        sample: list.slice(0, Math.min(limit, 12)).map((item) => item.name === '' ? '(正文)' : item.name),
        ...(container.note === undefined ? {} : { note: container.note }),
        ...(kind === 'plain'
            ? { note: '这是一份纯文本，直接 office.files.read 就行；用 archive.find 可以在里面按行找。' }
            : {}),
    };
}

/**
 * 列出归档里的条目。ZIP 给全部条目；gzip / zstd / 纯文本只有一条。
 */
export function archiveList(filePath, env, options = {}) {
    const bytes = readBytes(filePath, env);
    const kind = sniffArchive(bytes);
    if (kind === undefined) throw new Error(`office.archive：${filePath} 不是可识别的归档（ZIP / gzip / zstd / 文本）。`);
    const container = openContainer(kind, bytes);
    const limit = Math.max(1, Math.min(5000, Math.trunc(Number(options.limit) || 500)));
    const all = container.list();
    return {
        path: filePath,
        kind,
        count: all.length,
        truncated: all.length > limit,
        entries: all.slice(0, limit).map((item) => ({
            name: item.name === '' ? '(正文)' : item.name,
            bytes: item.bytes,
            method: item.method,
        })),
        ...(container.note === undefined ? {} : { note: container.note }),
    };
}

/**
 * 挑一条正文：显式 `entry` 优先；否则挑第一个像文本的条目；再否则挑最大的一条。
 * 返回 `undefined` 表示包是空的。
 */
function pickEntry(container, wanted) {
    const names = container.names;
    if (names.length === 0) return undefined;
    if (typeof wanted === 'string' && wanted !== '' && wanted !== '(正文)') {
        if (container.at(wanted) !== undefined) return wanted;
        // 允许用后缀 / 包含匹配：会话包里的名字带 uuid，模型未必抄得全。
        const hit = names.find((name) => name.endsWith(wanted) || name.includes(wanted));
        if (hit !== undefined) return hit;
        throw new Error(`office.archive：这个包里没有「${wanted}」；用 office.archive.list 看有哪些条目。`);
    }
    const textual = names.find((name) => /\.(jsonl?|txt|md|log|csv|tsv)$/i.test(name));
    if (textual !== undefined) return textual;
    return names.reduce((best, name) => (container.at(name).length > container.at(best).length ? name : best), names[0]);
}

/**
 * 取一段正文。默认从第 0 个字符起，最多 `maxChars`。
 * @param {{entry?: string, offset?: number, maxChars?: number}} [options]
 */
export function archiveText(filePath, env, options = {}) {
    const bytes = readBytes(filePath, env);
    const kind = sniffArchive(bytes);
    if (kind === undefined) throw new Error(`office.archive：${filePath} 不是可识别的归档（ZIP / gzip / zstd / 文本）。`);
    const container = openContainer(kind, bytes);
    const entry = pickEntry(container, options.entry);
    if (entry === undefined) return { path: filePath, kind, entry: null, text: '', chars: 0, truncated: false, note: '这个包里没有条目。' };
    const raw = container.at(entry);
    if (raw.length > MAX_ENTRY_BYTES) {
        // 正常路径下走不到这里：ZIP 在 unzipMeta 里按声明值就挡了、gzip/zstd 由
        // maxOutputLength 挡。留着是兜底（比如将来换了解压实现）。
        throw new Error(`office.archive：条目「${entry}」有 ${raw.length} 字节，超过单条上限；请用 archive.find 按行检索。`);
    }
    if (looksBinary(raw)) {
        return {
            path: filePath,
            kind,
            entry: entry === '' ? '(正文)' : entry,
            text: '',
            chars: 0,
            bytes: raw.length,
            truncated: false,
            note: '这条是二进制（媒体文件），不是文本；要取出来用 office.archive.extract。',
        };
    }
    const text = decodeText(raw);
    const offset = Math.max(0, Math.trunc(Number(options.offset) || 0));
    // maxChars 是**上限**：只挡「一次带回十万字符」，不把调用方要的小值悄悄放大
    // （要 40 个字符就给 40 个 —— 悄悄返 200 会让「翻页」这种用法算错步长）。
    const requested = Number(options.maxChars);
    const maxChars = Number.isFinite(requested) && requested > 0
        ? Math.min(100_000, Math.trunc(requested))
        : 8000;
    const slice = text.slice(offset, offset + maxChars);
    const truncated = offset + slice.length < text.length;
    return {
        path: filePath,
        kind,
        entry: entry === '' ? '(正文)' : entry,
        chars: text.length,
        offset,
        returned: slice.length,
        truncated,
        nextOffset: truncated ? offset + slice.length : null,
        text: slice,
    };
}

/**
 * 在归档的正文里按行检索（会话日志最常用的动作）。
 * @param {{entry?: string, pattern: string, ignoreCase?: boolean, maxHits?: number, maxLineChars?: number, context?: number}} options
 */
export function archiveFind(filePath, env, options = {}) {
    const pattern = typeof options.pattern === 'string' ? options.pattern : '';
    if (pattern === '') throw new Error('office.archive：find 需要 pattern（正则或普通字符串）。');
    let re;
    try {
        re = new RegExp(pattern, options.ignoreCase === false ? '' : 'i');
    } catch (error) {
        // 用户给的多半是普通文本（含 `(` `.` 之类），退回字面量匹配而不是报错。
        const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        re = new RegExp(escaped, options.ignoreCase === false ? '' : 'i');
    }
    const bytes = readBytes(filePath, env);
    const kind = sniffArchive(bytes);
    if (kind === undefined) throw new Error(`office.archive：${filePath} 不是可识别的归档（ZIP / gzip / zstd / 文本）。`);
    const container = openContainer(kind, bytes);
    const wanted = options.entry;
    // 逐条扫的候选：二进制跳过；超过单条上限的**不整条解码**（一条 1 GB 的日志
    // 读进内存会直接把宿主拖死），并把它记进 skippedEntries 如实报出来。
    // 显式指定 entry 时**同样**过这道闸：指定了一条媒体文件不等于可以拿它当文本扫。
    const skipped = [];
    const accept = (name) => {
        const data = container.at(name);
        if (data === undefined) return false;
        if (looksBinary(data)) {
            if (typeof wanted === 'string' && wanted !== '') {
                skipped.push({ entry: name === '' ? '(正文)' : name, bytes: data.length, reason: '二进制（不是文本）' });
            }
            return false;
        }
        if (data.length > MAX_ENTRY_BYTES) {
            skipped.push({ entry: name === '' ? '(正文)' : name, bytes: data.length, reason: '超过单条上限' });
            return false;
        }
        return true;
    };
    const targets = typeof wanted === 'string' && wanted !== ''
        ? [pickEntry(container, wanted)].filter((name) => name !== undefined && accept(name))
        : container.names.filter(accept);
    const maxHits = Math.max(1, Math.min(500, Math.trunc(Number(options.maxHits) || 40)));
    const requestedLine = Number(options.maxLineChars);
    const maxLineChars = Number.isFinite(requestedLine) && requestedLine > 0
        ? Math.min(4000, Math.trunc(requestedLine))
        : 400;
    const context = Math.max(0, Math.min(5, Math.trunc(Number(options.context) || 0)));
    const hits = [];
    let scannedLines = 0;
    let scannedEntries = 0;
    let truncated = false;
    for (const name of targets) {
        if (name === undefined) continue;
        scannedEntries += 1;
        const text = decodeText(container.at(name));
        const lines = text.split('\n');
        for (let i = 0; i < lines.length; i += 1) {
            scannedLines += 1;
            if (!re.test(lines[i])) continue;
            if (hits.length >= maxHits) {
                truncated = true;
                break;
            }
            const hit = {
                entry: name === '' ? '(正文)' : name,
                line: i + 1,
                text: lines[i].length > maxLineChars ? `${lines[i].slice(0, maxLineChars)}…` : lines[i],
            };
            if (context > 0) {
                hit.before = lines.slice(Math.max(0, i - context), i).map((line) => (line.length > maxLineChars ? `${line.slice(0, maxLineChars)}…` : line));
                hit.after = lines.slice(i + 1, i + 1 + context).map((line) => (line.length > maxLineChars ? `${line.slice(0, maxLineChars)}…` : line));
            }
            hits.push(hit);
        }
        if (truncated) break;
    }
    return {
        path: filePath,
        kind,
        pattern,
        scannedEntries,
        scannedLines,
        ...(skipped.length === 0 ? {} : { skippedEntries: skipped }),
        hits: hits.length,
        truncated: truncated || undefined,
        matches: hits,
    };
}

/**
 * 把一条正文落到工作目录（会话日志「先导出为文本再读」那条路）。
 * @param {{entry?: string, to?: string, overwrite?: boolean}} [options]
 */
export function archiveExtract(filePath, env, options = {}) {
    const bytes = readBytes(filePath, env);
    const kind = sniffArchive(bytes);
    if (kind === undefined) throw new Error(`office.archive：${filePath} 不是可识别的归档（ZIP / gzip / zstd / 文本）。`);
    const container = openContainer(kind, bytes);
    const entry = pickEntry(container, options.entry);
    if (entry === undefined) throw new Error('office.archive：这个包里没有条目。');
    const raw = container.at(entry);
    const base = String(filePath).replace(/^.*[\\/]/, '').replace(/\.(zip|gz|zst|zstd)$/i, '');
    const fallback = entry === '' ? `${base || 'archive'}.txt` : `${base}-${entry.replace(/[\\/]/g, '_')}`;
    const to = typeof options.to === 'string' && options.to.trim() !== '' ? options.to.trim() : fallback;
    assertInsideRoot(to, env, 'office.archive.extract');
    if (options.overwrite !== true && env.exists(to)) {
        throw new Error(`office.archive：${to} 已存在；换一个 to，或显式传 overwrite:true。`);
    }
    const written = env.writeFile(to, raw);
    return {
        path: filePath,
        kind,
        entry: entry === '' ? '(正文)' : entry,
        to: written.path,
        bytes: written.bytes,
        lines: looksBinary(raw) ? null : decodeText(raw).split('\n').length,
        crc32: crc32(raw),
    };
}

/**
 * 给 `office.files.read()` 用的诚实提示：这份内容其实是归档，按文本读只会得到乱码。
 * 判定不了（或就是普通文本）时返回 `undefined`，调用方照常读。
 *
 * 收在这里的第二层理由：判型与提示文案只有一处，`files.read` 与 `archive.*`
 * 不会对「这到底是不是压缩包」给出两套说法。
 * @param {Uint8Array} bytes
 */
export function archiveHintForBytes(bytes) {
    const kind = sniffArchive(bytes);
    if (kind === undefined || kind === 'plain') return undefined;
    let entries = 1;
    if (kind === 'zip') {
        // 只数条目、不解压：这条提示在**每次** files.read 一个 zip/xlsx/pptx 时都会走到，
        // 为了一句「有几个条目」把整包 inflate 出来是不可接受的（炸弹包会当场吃内存）。
        entries = zipEntryCount(bytes);
    }
    return `这是一份 ${kind} 归档（${entries} 个条目），按文本读只会得到乱码。`
        + '换 office.archive：info 看有哪些条目、list 列条目、text 取一段、find 按行检索、extract 落到工作目录。';
}
