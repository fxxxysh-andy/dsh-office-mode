/**
 * 知识库来源抽取：把工作目录里的一份文件变成「文本 + 图片资产」。
 *
 * 为什么单独一层、而不是复用 formats/word.js 的 read：那边返回的是**摘要**
 * （标题、段落数、表格数……），够生成文档回执，但知识库要的是**逐字原文**——
 * 引用必须能落回字符区间。所以这里只做抽取：ZIP 走 engine/zip，OOXML 直接扫
 * XML，PDF 交给 pdf.js 的既有引擎探测，全程零第三方依赖。
 *
 * 三种约束：
 *   1. **只读、不落盘**。图片以 Buffer 交回调用方，由 kb 决定资产落在哪。
 *   2. **边界在入口处**：resolve 后必须仍在 root 内；单文件默认不超过
 *      KB_EXTRACT_MAX_BYTES，超了报错，不静默截断。
 *   3. **不猜**。xlsx 不做公式计算（只取缓存值），OOXML 里不认识的媒体格式
 *      跳过并记 warning，不把 .emf 之类硬塞成 png。
 *
 * @module dsh-office-mode/kb-extract
 */
import { existsSync, readFileSync, statSync } from 'node:fs';
import { basename, extname, resolve, sep } from 'node:path';
import { unzip } from './engine/zip.js';
import { pdfText } from './pdf.js';

/** 单文件抽取上限（字节）。再大的文件不进知识库：入库是复制正文，不是搬文件。 */
export const KB_EXTRACT_MAX_BYTES = 8 * 1024 * 1024;

/** 按文本处理的扩展名（小写、不带点）。 */
const TEXT_EXTS = new Set([
    'txt', 'md', 'markdown', 'json', 'jsonl', 'csv', 'tsv', 'log', 'xml',
    'html', 'htm', 'css', 'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'py',
    'java', 'go', 'rs', 'c', 'h', 'cpp', 'hpp', 'sh', 'ps1', 'yml', 'yaml',
    'toml', 'ini', 'sql',
]);

/** 按二进制文档处理的扩展名。 */
const DOC_EXTS = new Set(['pdf', 'docx', 'pptx', 'xlsx']);

/** 图片扩展名 → MIME。key 是小写扩展名（不带点）。 */
const IMAGE_MIME = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    webp: 'image/webp',
    gif: 'image/gif',
};

const UTF8 = new TextDecoder('utf-8');

function extOf(filePath) {
    return extname(String(filePath ?? '')).slice(1).toLowerCase();
}

/**
 * 这个路径能不能作为知识库来源抽取。
 * 只看扩展名（调用方给的是相对工作目录的路径，不是 URL）。
 */
export function isExtractableKbPath(filePath) {
    if (typeof filePath !== 'string' || filePath.trim() === '') return false;
    const ext = extOf(filePath);
    if (ext === '') return false;
    return TEXT_EXTS.has(ext) || DOC_EXTS.has(ext) || Object.prototype.hasOwnProperty.call(IMAGE_MIME, ext);
}

function kindOfExt(ext) {
    if (Object.prototype.hasOwnProperty.call(IMAGE_MIME, ext)) return 'image';
    if (DOC_EXTS.has(ext)) return ext;
    return 'text';
}

function normalizedLimit(value) {
    if (value === undefined || value === null || value === '') return KB_EXTRACT_MAX_BYTES;
    const n = typeof value === 'number' ? value : Number.parseInt(String(value), 10);
    if (!Number.isFinite(n) || n <= 0) throw new Error('maxBytes 必须是正整数（收到「' + String(value) + '」）。');
    return Math.trunc(n);
}

function zipBaseName(name) {
    const at = name.lastIndexOf('/');
    return at === -1 ? name : name.slice(at + 1);
}

/* ── XML 小工具 ─────────────────────────────────────────────────────────── */

/** 反转义五个预定义实体与数字实体（与 engine/xml.js 同一口径，但本模块不依赖它）。 */
function decodeXmlEntities(text) {
    if (typeof text !== 'string' || text.indexOf('&') === -1) return text;
    return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (whole, body) => {
        if (body[0] === '#') {
            const hex = body[1] === 'x' || body[1] === 'X';
            const code = Number.parseInt(hex ? body.slice(2) : body.slice(1), hex ? 16 : 10);
            if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return whole;
            try {
                return String.fromCodePoint(code);
            } catch {
                return whole;
            }
        }
        switch (body) {
            case 'amp': return '&';
            case 'lt': return '<';
            case 'gt': return '>';
            case 'quot': return '"';
            case 'apos': return "'";
            default: return whole;
        }
    });
}

function localName(name) {
    const at = name.indexOf(':');
    return at === -1 ? name : name.slice(at + 1);
}

/** 找开始标签的 '>'，跳过引号内的 '>'（属性值里可以出现实体外的 '>'）。 */
function findTagEnd(xml, from) {
    let quote = '';
    for (let i = from; i < xml.length; i += 1) {
        const ch = xml[i];
        if (quote !== '') {
            if (ch === quote) quote = '';
        } else if (ch === '"' || ch === "'") {
            quote = ch;
        } else if (ch === '>') {
            return i;
        }
    }
    return -1;
}

function parseAttrs(source) {
    const attrs = {};
    const re = /([^\s=/]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let match;
    while ((match = re.exec(source)) !== null) {
        attrs[match[1]] = decodeXmlEntities(match[3] !== undefined ? match[3] : match[4]);
    }
    return attrs;
}

/**
 * 极小的 XML 事件扫描器。只覆盖 OOXML 用到的语法：声明、注释、CDATA、
 * DOCTYPE、自闭合标签；不做命名空间解析（按 localName 比较）。
 *
 * handlers: { start(name, attrs, selfClosing), end(name), text(raw) }
 */
function scanXml(xml, handlers) {
    const len = xml.length;
    let at = 0;
    while (at < len) {
        const lt = xml.indexOf('<', at);
        if (lt === -1) {
            if (handlers.text) handlers.text(xml.slice(at));
            break;
        }
        if (lt > at && handlers.text) handlers.text(xml.slice(at, lt));
        if (xml.startsWith('<!--', lt)) {
            const end = xml.indexOf('-->', lt + 4);
            at = end === -1 ? len : end + 3;
            continue;
        }
        if (xml.startsWith('<![CDATA[', lt)) {
            const end = xml.indexOf(']]>', lt + 9);
            const body = xml.slice(lt + 9, end === -1 ? len : end);
            if (handlers.text) handlers.text(body);
            at = end === -1 ? len : end + 3;
            continue;
        }
        if (xml.startsWith('<?', lt)) {
            const end = xml.indexOf('?>', lt + 2);
            at = end === -1 ? len : end + 2;
            continue;
        }
        if (xml.startsWith('<!', lt)) {
            const end = xml.indexOf('>', lt + 2);
            at = end === -1 ? len : end + 1;
            continue;
        }
        const gt = findTagEnd(xml, lt + 1);
        if (gt === -1) {
            if (handlers.text) handlers.text(xml.slice(lt));
            break;
        }
        const raw = xml.slice(lt + 1, gt).trim();
        if (raw === '') {
            at = gt + 1;
            continue;
        }
        if (raw[0] === '/') {
            const name = raw.slice(1).trim();
            if (name !== '' && handlers.end) handlers.end(name);
            at = gt + 1;
            continue;
        }
        const selfClosing = raw.endsWith('/');
        const body = selfClosing ? raw.slice(0, -1) : raw;
        const space = body.search(/\s/);
        const name = space === -1 ? body : body.slice(0, space);
        if (name === '') {
            at = gt + 1;
            continue;
        }
        const attrs = space === -1 ? {} : parseAttrs(body.slice(space));
        if (handlers.start) handlers.start(name, attrs, selfClosing);
        if (selfClosing && handlers.end) handlers.end(name);
        at = gt + 1;
    }
}

/* ── OOXML：DOCX ────────────────────────────────────────────────────────── */

/** 正文段落 + 表格（行内单元格用 \t，行间用 \n）。表格外的段落按出现顺序成行。 */
function extractDocxBody(xml) {
    const lines = [];
    let tableDepth = 0;
    let row = null;
    let cell = null;
    let inParagraph = false;
    let paragraph = [];
    let capture = null;

    scanXml(xml, {
        start(name) {
            const local = localName(name);
            if (local === 'tbl') {
                tableDepth += 1;
                return;
            }
            if (local === 'tr') {
                if (tableDepth === 1) row = [];
                return;
            }
            if (local === 'tc') {
                if (tableDepth === 1) cell = [];
                return;
            }
            if (local === 'p') {
                // 单元格里也可能有多个段落：也收，最后按空格拼成一格。
                if (tableDepth <= 1) {
                    inParagraph = true;
                    paragraph = [];
                }
                return;
            }
            if (local === 't') {
                if (inParagraph) capture = [];
                return;
            }
            if (local === 'tab') {
                if (inParagraph) paragraph.push('\t');
                return;
            }
            if (local === 'br' || local === 'cr') {
                if (inParagraph) paragraph.push('\n');
            }
        },
        end(name) {
            const local = localName(name);
            if (local === 't') {
                if (capture !== null) {
                    if (inParagraph) paragraph.push(decodeXmlEntities(capture.join('')));
                    capture = null;
                }
                return;
            }
            if (local === 'p') {
                if (!inParagraph) return;
                const value = paragraph.join('');
                if (tableDepth === 0) lines.push(value);
                else if (tableDepth === 1 && cell !== null) cell.push(value);
                inParagraph = false;
                paragraph = [];
                return;
            }
            if (local === 'tc') {
                if (tableDepth === 1) {
                    if (row !== null) row.push(cell === null ? '' : cell.join(' ').trim());
                    cell = null;
                }
                return;
            }
            if (local === 'tr') {
                if (tableDepth === 1) {
                    const cells = row === null ? [] : row;
                    // 末尾空单元格不带出一串悬空 \t；行首空单元格保留（列位有意义）。
                    while (cells.length > 0 && cells[cells.length - 1] === '') cells.pop();
                    lines.push(cells.join('\t'));
                    row = null;
                }
                return;
            }
            if (local === 'tbl') {
                tableDepth = Math.max(0, tableDepth - 1);
            }
        },
        text(value) {
            if (capture !== null) capture.push(value);
        },
    });
    return lines.join('\n');
}

function collectMedia(files, prefix, warnings) {
    const images = [];
    for (const [name, data] of files) {
        if (!name.startsWith(prefix)) continue;
        const base = zipBaseName(name);
        if (base === '' || base.endsWith('/')) continue;
        const ext = extname(base).slice(1).toLowerCase();
        const mime = IMAGE_MIME[ext];
        if (!mime) {
            warnings.push('「' + name + '」不是支持的图片格式（png/jpeg/webp/gif），未收进图片资产。');
            continue;
        }
        images.push({ name: base, ext, mime, data: Buffer.from(data) });
    }
    return images;
}

function docxParts(bytes) {
    const files = unzip(bytes);
    const documentXml = files.get('word/document.xml');
    if (documentXml === undefined) throw new Error('DOCX 包里没有 word/document.xml（不是有效的 .docx？）');
    const warnings = [];
    return {
        text: extractDocxBody(UTF8.decode(documentXml)),
        images: collectMedia(files, 'word/media/', warnings),
        warnings,
    };
}

/** 从 DOCX 的 ZIP 字节里抽正文与图片（单测入口；extractKbSource 也走这条）。 */
export function extractDocxText(buffer) {
    const { text, images } = docxParts(Buffer.from(buffer));
    return { text, images };
}

/* ── OOXML：PPTX ────────────────────────────────────────────────────────── */

/** 一张幻灯片的文本：段落内 run 直接相接，段落之间换行；空段落丢掉。 */
function extractPptxSlideText(xml) {
    const paragraphs = [];
    let current = null;
    let synthetic = false;
    let capture = null;
    scanXml(xml, {
        start(name) {
            const local = localName(name);
            if (local === 'p') {
                current = [];
                synthetic = false;
                return;
            }
            if (local === 't') {
                // 正常的 a:t 都在 a:p 里；这里给「没有 a:p 的最小样例」留一条路，
                // 免得只因少一层段落就把整页文本丢掉。
                if (current === null) {
                    current = [];
                    synthetic = true;
                }
                capture = [];
                return;
            }
            if (local === 'br') {
                if (current !== null) current.push('\n');
            }
        },
        end(name) {
            const local = localName(name);
            if (local === 't') {
                if (capture !== null) {
                    if (current !== null) current.push(decodeXmlEntities(capture.join('')));
                    capture = null;
                }
                if (synthetic && current !== null) {
                    const value = current.join('');
                    if (value.trim() !== '') paragraphs.push(value);
                    current = null;
                    synthetic = false;
                }
                return;
            }
            if (local === 'p') {
                if (current === null) return;
                const value = current.join('');
                if (value.trim() !== '') paragraphs.push(value);
                current = null;
                synthetic = false;
            }
        },
        text(value) {
            if (capture !== null) capture.push(value);
        },
    });
    return paragraphs.join('\n');
}

function slideNumber(name) {
    const match = /^ppt\/slides\/slide(\d+)\.xml$/i.exec(name);
    return match ? Number(match[1]) : 0;
}

function pptxParts(bytes) {
    const files = unzip(bytes);
    const slides = [...files.keys()]
        .map((name) => ({ name, n: slideNumber(name) }))
        .filter((entry) => entry.n > 0)
        .sort((left, right) => left.n - right.n);
    const text = slides.map((slide) => extractPptxSlideText(UTF8.decode(files.get(slide.name)))).join('\n\n');
    const warnings = [];
    return {
        text,
        images: collectMedia(files, 'ppt/media/', warnings),
        warnings,
    };
}

/** 从 PPTX 的 ZIP 字节里抽正文与图片（页序按 slide 后的数字，不按字典序）。 */
export function extractPptxText(buffer) {
    const { text, images } = pptxParts(Buffer.from(buffer));
    return { text, images };
}

/* ── OOXML：XLSX ────────────────────────────────────────────────────────── */

/** 共享字符串表：si 下所有 t（含富文本 run）拼接，忽略 rPh 注音。 */
function parseSharedStrings(xml) {
    if (xml === undefined) return [];
    const values = [];
    let current = null;
    let phoneticDepth = 0;
    let capture = null;
    scanXml(xml, {
        start(name) {
            const local = localName(name);
            if (local === 'si') {
                current = [];
                phoneticDepth = 0;
                return;
            }
            if (local === 'rPh') {
                phoneticDepth += 1;
                return;
            }
            if (local === 't') {
                if (current !== null && phoneticDepth === 0) capture = [];
            }
        },
        end(name) {
            const local = localName(name);
            if (local === 't') {
                if (capture !== null) {
                    if (current !== null) current.push(decodeXmlEntities(capture.join('')));
                    capture = null;
                }
                return;
            }
            if (local === 'rPh') {
                phoneticDepth = Math.max(0, phoneticDepth - 1);
                return;
            }
            if (local === 'si') {
                if (current !== null) values.push(current.join(''));
                current = null;
                phoneticDepth = 0;
            }
        },
        text(value) {
            if (capture !== null) capture.push(value);
        },
    });
    return values;
}

/** A→0、Z→25、AA→26。解析不出列号时返回 -1。 */
function columnIndex(ref) {
    if (typeof ref !== 'string') return -1;
    const match = /^([A-Za-z]+)/.exec(ref);
    if (match === null) return -1;
    let value = 0;
    for (const ch of match[1].toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
    return value - 1;
}

function sheetRowLine(cells) {
    if (cells.length === 0) return '';
    let width = 0;
    for (const cell of cells) width = Math.max(width, cell.col + 1);
    const values = new Array(width).fill('');
    for (const cell of cells) values[cell.col] = cell.value;
    while (values.length > 0 && values[values.length - 1] === '') values.pop();
    return values.join('\t');
}

/** 一个 sheet：row 为行，c 为单元格；支持 inlineStr / shared string / 数字。 */
function extractXlsxSheet(xml, sharedStrings, warnings, sheetName) {
    const rows = [];
    let row = null;
    let cell = null;
    let capture = null;
    let captureKind = '';

    const closeCell = () => {
        if (cell === null) return;
        const value = cell.inline ? cell.inlineText.join('') : cell.value;
        let resolved = value;
        if (!cell.inline && cell.type === 's') {
            const index = Number.parseInt(value, 10);
            if (!Number.isInteger(index) || index < 0 || index >= sharedStrings.length) {
                warnings.push('「' + sheetName + '」的 ' + String(cell.ref ?? '某个单元格')
                    + ' 引用了不存在的共享字符串 #' + value + '，按空值处理。');
                resolved = '';
            } else {
                resolved = sharedStrings[index];
            }
        }
        const col = columnIndex(cell.ref);
        row.cells.push({ col: col === -1 ? row.cells.length : col, value: resolved });
        cell = null;
        capture = null;
        captureKind = '';
    };

    scanXml(xml, {
        start(name, attrs) {
            const local = localName(name);
            if (local === 'row') {
                row = { cells: [] };
                return;
            }
            if (local === 'c') {
                if (row === null) row = { cells: [] };
                cell = {
                    ref: typeof attrs.r === 'string' ? attrs.r : null,
                    type: typeof attrs.t === 'string' ? attrs.t : '',
                    value: '',
                    inline: false,
                    inlineText: [],
                };
                return;
            }
            if (cell === null) return;
            if (local === 'is') {
                cell.inline = true;
                return;
            }
            if (local === 'v') {
                capture = [];
                captureKind = 'v';
                return;
            }
            if (local === 't' && cell.inline) {
                capture = [];
                captureKind = 't';
            }
        },
        end(name) {
            const local = localName(name);
            if (local === 'v' && capture !== null && captureKind === 'v') {
                cell.value = decodeXmlEntities(capture.join(''));
                capture = null;
                captureKind = '';
                return;
            }
            if (local === 't' && capture !== null && captureKind === 't') {
                cell.inlineText.push(decodeXmlEntities(capture.join('')));
                capture = null;
                captureKind = '';
                return;
            }
            if (local === 'c') {
                closeCell();
                return;
            }
            if (local === 'row') {
                if (row !== null && row.cells.length > 0) rows.push(sheetRowLine(row.cells));
                row = null;
            }
        },
        text(value) {
            if (capture !== null) capture.push(value);
        },
    });
    return rows.join('\n');
}

function sheetNumber(name) {
    const match = /^xl\/worksheets\/sheet(\d+)\.xml$/i.exec(name);
    return match ? Number(match[1]) : 0;
}

function xlsxParts(bytes) {
    const files = unzip(bytes);
    const sharedXml = files.get('xl/sharedStrings.xml');
    const sharedStrings = parseSharedStrings(sharedXml === undefined ? undefined : UTF8.decode(sharedXml));
    const sheets = [...files.keys()]
        .map((name) => ({ name, n: sheetNumber(name) }))
        .filter((entry) => entry.n > 0)
        .sort((left, right) => left.n - right.n);
    const warnings = [];
    const text = sheets
        .map((sheet) => extractXlsxSheet(UTF8.decode(files.get(sheet.name)), sharedStrings, warnings, sheet.name))
        .filter((value) => value !== '')
        .join('\n\n');
    return {
        text,
        images: collectMedia(files, 'xl/media/', warnings),
        warnings,
    };
}

/** 从 XLSX 的 ZIP 字节里抽正文与图片（不计算公式，只取 v 里的缓存值）。 */
export function extractXlsxText(buffer) {
    const { text, images } = xlsxParts(Buffer.from(buffer));
    return { text, images };
}

/* ── 入口 ───────────────────────────────────────────────────────────────── */

function decodeTextFile(buffer) {
    if (buffer.indexOf(0) !== -1) return null;
    return UTF8.decode(buffer);
}

/**
 * 抽取一份知识库来源。
 *
 * @param {object} [options]
 * @param {string} options.root 工作区根（绝对或相对当前进程）
 * @param {string} options.path 相对 root 的文件路径
 * @param {object} [options.env] 办公模式运行环境（PDF 必需，pdfText 用它 resolve）
 * @param {object} [options.cache] 引擎缓存（透传给 PDF）
 * @param {number} [options.maxBytes] 单文件上限，默认 KB_EXTRACT_MAX_BYTES
 * @returns {Promise<{ok: true, kind: string, title: string, text: string, images: Array<{name: string, ext: string, mime: string, data: Buffer}>, warnings: string[], bytes: number}>}
 */
export async function extractKbSource({ root, path: relPath, env, cache, maxBytes } = {}) {
    if (typeof root !== 'string' || root.trim() === '') throw new Error('抽取需要工作目录 root。');
    if (typeof relPath !== 'string' || relPath.trim() === '') throw new Error('抽取需要相对路径 path。');

    const limit = normalizedLimit(maxBytes);
    const rootAbs = resolve(root);
    const absolute = resolve(rootAbs, relPath);
    const rootPrefix = rootAbs.endsWith(sep) ? rootAbs : rootAbs + sep;
    // 只接受 root 内的路径：../x、别的盘符、绝对路径都要在这里挡住。
    if (absolute !== rootAbs && !absolute.startsWith(rootPrefix)) {
        throw new Error('路径越界：「' + relPath + '」不在工作目录内，拒绝抽取。');
    }

    if (!isExtractableKbPath(relPath)) {
        throw new Error('「' + relPath + '」不是可抽取的知识库来源（支持文本、PDF、Word/PPT/Excel 与 png/jpeg/webp/gif 图片）。');
    }
    const ext = extOf(relPath);
    const kind = kindOfExt(ext);

    // PDF 先判 env：没有 env 时 pdfText 连 resolve 都做不了，报错要指向真正缺的东西。
    if (kind === 'pdf' && !env) throw new Error('抽 PDF 需要 env（「' + relPath + '」）。');
    if (!existsSync(absolute)) throw new Error('找不到文件：「' + relPath + '」。');
    const stat = statSync(absolute);
    if (!stat.isFile()) throw new Error('「' + relPath + '」不是文件，拒绝抽取。');
    if (stat.size > limit) {
        throw new Error('「' + relPath + '」有 ' + stat.size + ' 字节，超过单文件上限 ' + limit
            + ' 字节（KB_EXTRACT_MAX_BYTES），拒绝抽取。');
    }

    if (kind === 'pdf') {
        let result;
        try {
            result = await pdfText(relPath, {}, env, cache);
        } catch (error) {
            // 损坏 / 加密 / 根本不是 PDF：与「没有文本层」是两回事 —— 前者该让人重导文件，
            // 后者该让人去跑 OCR。文案保持原样，只加 code 让上层能分开提示。
            const reason = String(error?.message ?? error);
            const wrapped = new Error('「' + relPath + '」解析失败（pdf）：' + reason);
            // 三种「读不出来」要分开：加密（先去解密）、缺 PDF 引擎（去装引擎）、真损坏
            // （换一份副本）。混成一句「解析失败」会让人对着一个完好的文件反复重导。
            wrapped.code = /可用引擎/.test(reason)
                ? 'kb_engine_missing'
                : (/加密/.test(reason) ? 'kb_encrypted' : 'kb_parse_failed');
            throw wrapped;
        }
        const text = typeof result?.text === 'string' ? result.text : '';
        if (text.trim() === '') {
            const empty = new Error('「' + relPath + '」没有可抽取文本（PDF 没有文本层？）。');
            empty.code = 'kb_no_text_layer';
            throw empty;
        }
        return {
            ok: true,
            kind,
            title: basename(absolute),
            text,
            images: [],
            warnings: [],
            bytes: stat.size,
        };
    }

    const buffer = readFileSync(absolute);

    if (kind === 'image') {
        const mime = IMAGE_MIME[ext];
        const name = basename(absolute);
        return {
            ok: true,
            kind,
            title: name,
            text: '[图片：' + name + ']',
            images: [{ name, ext, mime, data: buffer }],
            warnings: [],
            bytes: stat.size,
        };
    }

    if (kind === 'text') {
        const text = decodeTextFile(buffer);
        if (text === null) throw new Error('「' + relPath + '」含 NUL 字节，像是二进制文件，不能按文本抽取。');
        if (text.trim() === '') throw new Error('「' + relPath + '」没有可抽取文本（空文件或全是空白）。');
        return {
            ok: true,
            kind,
            title: basename(absolute),
            text,
            images: [],
            warnings: [],
            bytes: stat.size,
        };
    }

    let parts;
    try {
        if (kind === 'docx') parts = docxParts(buffer);
        else if (kind === 'pptx') parts = pptxParts(buffer);
        else parts = xlsxParts(buffer);
    } catch (error) {
        throw new Error('「' + relPath + '」解析失败（' + kind + '）：' + (error?.message ?? error));
    }

    if (parts.text.trim() === '') throw new Error('「' + relPath + '」没有可抽取文本（' + kind + ' 里没有文字）。');
    return {
        ok: true,
        kind,
        title: basename(absolute),
        text: parts.text,
        images: parts.images,
        warnings: parts.warnings,
        bytes: stat.size,
    };
}
