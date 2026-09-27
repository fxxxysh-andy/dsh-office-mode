/**
 * PPT 就地修订引擎（`office.ppt.readSlides` / `office.ppt.revise`）。
 *
 * 为什么单独一个模块：`ppt.js` 是「从零生成」，这里是「读回已有包、在原 XML 上定点改」。
 * 两者共用的只有单位换算与枚举收敛，共用面很小；放一起会让本来就 5000 行的 ppt.js
 * 更难读，而这一块的难点（关系表、形状定位、rPr 保序）也自成一套。
 *
 * 三条硬约束（都来自 docs/ooxml-pitfalls.md 踩过的坑）：
 *
 * 1. **绝不重建 slide XML。** 已有页面里有母版元素、装饰形状、超链接关系、备注关系，
 *    重建只能重建出「我们以为有的东西」，用户文件里其余部分会静默消失。
 *    所以这里全部走「字符串定位 → 定点替换」：只改命中的那段。
 * 2. **改 xfrm 之外的东西一律不动 xfrm。** `move` / `resize` 只替换 `a:off` / `a:ext`
 *    的属性值，不重写 `a:xfrm` 元素本身 —— 里面还挂着 `rot` 与 `flipH`/`flipV`，
 *    丢掉它们等于把旋转过的形状掰正。
 * 3. **rPr 里没有的属性写 `a:rPr` 的开标签，有的只改属性值。** `a:rPr` 的子元素顺序
 *    由 schema 定死（ln → solidFill → effectLst → latin → ea → cs → hlinkClick …），
 *    新增一个 `a:solidFill` 必须插在 `a:latin` 之前；已有的 `solidFill` 又必须先删掉，
 *    否则同一个 `rPr` 里出现两个填充，PowerPoint 判包损坏。
 *
 * @module dsh-office-mode/formats/ppt-revise
 */
import { createHash } from 'node:crypto';
import { unzip, zip } from '../engine/zip.js';
import { parseXml, descendants, children as xmlChildren, textOf } from '../engine/xml.js';
import { EMU_PER_CM, asArray, asNumber, asString, clampNumber } from '../engine/kit.js';

/** 幻灯片尺寸兜底（16:9）。presentation.xml 读不到时按它算「页外」判断。 */
const DEFAULT_SLIDE = { cx: 12192000, cy: 6858000 };

const RT_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const RT_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const RT_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** 96 DPI 下 1 像素 = 9525 EMU（与 ppt.js 的 EMU_PER_PX 同源）。 */
const EMU_PER_PX = 9525;

/** 插图支持的适配方式，与 ppt.js 的 IMAGE_FITS 同源。 */
const IMAGE_FITS = ['contain', 'cover', 'natural'];

/** a:srcRect 是千分之一百分比（100000 = 100%），给满 100000 会被判非法，封顶 99999。 */
const SRC_RECT_MAX = 99999;

/** 媒体部件的目录前缀：抽取、编号顺延、内容类型判断都以它为准。 */
const MEDIA_PREFIX = 'ppt/media/';

/** 与 ppt.js 的 TEXT_ALIGN 同源；「改一个已有形状」同样必须收敛到合法枚举。 */
const TEXT_ALIGN = {
    l: 'l', left: 'l', ctr: 'ctr', center: 'ctr', centre: 'ctr', r: 'r', right: 'r',
    just: 'just', justify: 'just', justified: 'just', dist: 'dist', distributed: 'dist',
};

const TEXT_ANCHOR = {
    t: 't', top: 't', ctr: 'ctr', center: 'ctr', centre: 'ctr', middle: 'ctr',
    b: 'b', bottom: 'b', just: 'just', dist: 'dist',
};

/** 形状类型 → 形状名列表。名字取自 ppt.js 里各版式实际写的 `p:cNvPr name=`。 */
const SHAPE_KINDS = {
    title: ['Title'],
    subtitle: ['Subtitle'],
    kicker: ['Kicker'],
    presenter: ['Presenter'],
    date: ['Date'],
    statement: ['Statement', 'Statement Sub'],
    quote: ['Quote Text', 'Quote By', 'Quote Mark', 'Quote Bar'],
};

/** 段落与 run 的正则。只用它们做「定位」与「计数」，不拿来做整体解析。 */
const PARA_RE = /<a:p>[\s\S]*?<\/a:p>|<a:p\/>/g;
const RUN_RE = /<a:r>[\s\S]*?<\/a:r>/g;
const RPR_RE = /<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/;

// ───────────────────────────────────────────────────────────────────────────
// 基础工具
// ───────────────────────────────────────────────────────────────────────────

function toHex(value, fallback = '000000') {
    const text = asString(value, '').trim().replace(/^#/, '').replace(/[^0-9a-fA-F]/g, '');
    if (text.length === 3) return text.split('').map((ch) => ch + ch).join('').toUpperCase();
    if (text.length === 6) return text.toUpperCase();
    return fallback;
}

function cmToEmu(value) {
    return Math.round(asNumber(value, 0) * EMU_PER_CM);
}

function emuToCm(value) {
    return Math.round((asNumber(value, 0) / EMU_PER_CM) * 100) / 100;
}

function escapeAttr(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/** 把 v 写进一段 XML 的属性：既改已有的，也在没有时补上。 */
function setAttr(tag, name, value) {
    const re = new RegExp(`(\\s${name}=")[^"]*(")`);
    if (re.test(tag)) return tag.replace(re, `$1${value}$2`);
    return tag.replace(/^(<\w+(?::\w+)?)/, `$1 ${name}="${value}"`);
}

function readAttr(tag, name) {
    const match = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
    return match === null ? undefined : match[1];
}

/** 解码 XML 实体，用于把形状名 / 文本还原成可比较的字符串。 */
function decodeEntities(text) {
    return String(text ?? '').replace(/&(?:#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole) => {
        if (whole.startsWith('&#x') || whole.startsWith('&#X')) {
            const code = Number.parseInt(whole.slice(3, -1), 16);
            return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        if (whole.startsWith('&#')) {
            const code = Number.parseInt(whole.slice(2, -1), 10);
            return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        switch (whole) {
            case '&amp;': return '&';
            case '&lt;': return '<';
            case '&gt;': return '>';
            case '&quot;': return '"';
            case '&apos;': return "'";
            default: return whole;
        }
    });
}

/** 把纯文本包成 `<a:p>`（保留 rPr 模板，因此改写后字体字号不被改掉）。 */
function paragraphFromText(text, rPrTemplate, pPr) {
    const lines = String(text ?? '').split('\n');
    const head = pPr === undefined ? '<a:pPr><a:buNone/></a:pPr>' : pPr;
    return lines.map((line) => {
        const tag = rPrTemplate === undefined
            ? '<a:rPr lang="zh-CN" altLang="en-US" sz="1800" dirty="0"/>'
            : rPrTemplate;
        return `<a:p>${head}<a:r>${tag}<a:t xml:space="preserve">${escapeAttr(line)}</a:t></a:r></a:p>`;
    }).join('');
}

/**
 * 把 `text` 写成「只有一个 run 的段落组」，rPr 沿用模板。
 * `a:t` 一定带 `xml:space="preserve"`：带前导 / 尾随空格的文本去掉它会被渲染器吃掉空格。
 */
function replaceParagraphText(block, text) {
    const rPr = RPR_RE.exec(block)?.[0];
    const pPr = /<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/.exec(block)?.[0];
    return paragraphFromText(text, rPr, pPr);
}

/**
 * 整段重写一个 `p:txBody`，只换掉里面的 `a:p`，**保留外壳**。
 *
 * 为什么不能用 replaceParagraphText 直接盖在 txBody 上：那样会把
 * `<p:txBody>`、`<a:bodyPr>`、`<a:lstStyle/>` 一起替换掉，产出
 * `<p:spPr>…</p:spPr><a:p>…</a:p></p:sp>` 这种没有 txBody 的坏形状。
 * 它在本引擎自己的解析器与 LibreOffice 下都照常渲染（读回时报「没有文本」），
 * 真实 PowerPoint 则会判包损坏 —— 正是 docs/ooxml-pitfalls.md 那一类最难发现的坑。
 * 2026-09-22 实测：改一次封面标题就踩中，`readSlides()` 随即报 Title「txBody false」。
 *
 * `a:bodyPr` 与 `a:lstStyle` 必须原样保留：前者带着 anchor / 内边距 / autofit，
 * 丢了会让文本框的垂直对齐与自动缩放行为一起变。
 */
function replaceTextBodyParagraphs(txBody, text) {
    const open = /^<p:txBody\b[^>]*>/.exec(txBody);
    const close = /<\/p:txBody>$/.exec(txBody);
    if (open === null || close === null) {
        // 自闭合或残缺的 txBody：交回调用方，不要在这里猜一个外壳出来
        return undefined;
    }
    const inner = txBody.slice(open[0].length, txBody.length - close[0].length);
    // bodyPr 与 lstStyle 是 txBody 的固定前缀，段落一律排在它们之后
    const prefix = /^(?:\s*<a:bodyPr\b[^>]*\/>|\s*<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>)?\s*(?:<a:lstStyle\b[^>]*\/>)?/.exec(inner)?.[0] ?? '';
    return `${open[0]}${prefix}${replaceParagraphText(inner, text)}</p:txBody>`;
}

// ───────────────────────────────────────────────────────────────────────────
// 包级读取
// ───────────────────────────────────────────────────────────────────────────

/** 解析 `.rels`：返回 `Map<Id, {target, type, mode}>`。 */
function parseRels(text) {
    const out = new Map();
    if (text === undefined) return out;
    const root = parseXml(text).children[0];
    for (const rel of xmlChildren(root, 'Relationship')) {
        out.set(rel.attrs.Id, {
            target: rel.attrs.Target,
            type: rel.attrs.Type,
            mode: rel.attrs.TargetMode,
        });
    }
    return out;
}

/** 把关系 Target 解析成包内条目名。 */
function resolveTarget(fromPart, target) {
    const value = asString(target, '');
    if (value === '') return undefined;
    if (value.startsWith('/')) return value.slice(1);
    const base = fromPart.slice(0, fromPart.lastIndexOf('/'));
    const parts = `${base}/${value}`.split('/');
    const stack = [];
    for (const part of parts) {
        if (part === '' || part === '.') continue;
        if (part === '..') stack.pop();
        else stack.push(part);
    }
    return stack.join('/');
}

function relsPartOf(partName) {
    const at = partName.lastIndexOf('/');
    return `${partName.slice(0, at)}/_rels/${partName.slice(at + 1)}.rels`;
}

/** 读一个部件的文本；部件不在包里就是 undefined（调用方自己决定这是不是异常）。 */
function partText(files, decoder, partName) {
    const data = files.get(partName);
    return data === undefined ? undefined : decoder.decode(data);
}

/** 按 `p:sldIdLst` 的顺序列出幻灯片部件（放映顺序，不是文件名顺序）。 */
function orderedSlideParts(files, decoder) {
    const presentation = files.get('ppt/presentation.xml');
    if (presentation === undefined) throw new Error('pptx: 缺少 ppt/presentation.xml（不是有效的演示文稿）');
    const presRoot = parseXml(decoder.decode(presentation)).children[0];
    const rels = parseRels(decoder.decode(files.get('ppt/_rels/presentation.xml.rels') ?? new Uint8Array()));
    const idList = descendants(presRoot, 'p:sldIdLst')[0];
    const refs = idList === undefined ? [] : xmlChildren(idList, 'p:sldId');
    const ordered = refs
        .map((node) => rels.get(node.attrs['r:id'])?.target)
        .filter((target) => typeof target === 'string')
        .map((target) => resolveTarget('ppt/presentation.xml', target));
    if (ordered.length > 0) return ordered;
    return [...files.keys()]
        .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
        .sort((a, b) => Number.parseInt(a.replace(/\D+/g, ''), 10) - Number.parseInt(b.replace(/\D+/g, ''), 10));
}

// ───────────────────────────────────────────────────────────────────────────
// 包级读取：内容类型、媒体部件
// ───────────────────────────────────────────────────────────────────────────

/**
 * 图片格式与像素尺寸：只认 PNG / JPEG / GIF。
 *
 * 与 `ppt.js` 的 `imageInfo()` **同源实现**，但这里刻意复制一份而不是静态引入
 * `ppt.js`：两个模块本来就互相引用（ppt.js 只在函数里动态 import 本模块），
 * 再加一条静态回边就成了真正的循环依赖 —— 本仓库为「模块实例陈旧导致新增导出
 * 看不见」吃过一次苦，循环依赖会让这类问题更难定位。代价是改探测器要改两处，
 * 所以两边的函数名与返回形状必须保持一致。
 */
function imageInfo(raw) {
    const buffer = Buffer.isBuffer(raw) ? raw : Buffer.from(raw);
    if (buffer.length > 24 && buffer.readUInt32BE(0) === 0x89504e47) {
        return {
            ext: 'png', contentType: 'image/png',
            width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20),
        };
    }
    if (buffer.length > 4 && buffer[0] === 0xff && buffer[1] === 0xd8) {
        let at = 2;
        while (at + 9 < buffer.length) {
            if (buffer[at] !== 0xff) {
                at += 1;
                continue;
            }
            const marker = buffer[at + 1];
            const length = buffer.readUInt16BE(at + 2);
            const isSof = marker >= 0xc0 && marker <= 0xcf
                && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
            if (isSof) {
                return {
                    ext: 'jpeg', contentType: 'image/jpeg',
                    height: buffer.readUInt16BE(at + 5), width: buffer.readUInt16BE(at + 7),
                };
            }
            at += 2 + length;
        }
        return { ext: 'jpeg', contentType: 'image/jpeg', width: 1600, height: 900, guessed: true };
    }
    if (buffer.length > 10 && buffer.toString('ascii', 0, 3) === 'GIF') {
        return {
            ext: 'gif', contentType: 'image/gif',
            width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8),
        };
    }
    return undefined;
}

/** `[Content_Types].xml` → 扩展名默认表 + 逐部件覆盖表（与 ppt.js 同源）。 */
function parseContentTypes(text) {
    const defaults = new Map();
    const overrides = new Map();
    if (typeof text !== 'string') return { defaults, overrides };
    const root = parseXml(text).children[0];
    if (root === undefined) return { defaults, overrides };
    for (const node of xmlChildren(root)) {
        if (node.name === 'Default') {
            const ext = asString(node.attrs.Extension, '').trim().toLowerCase();
            const type = asString(node.attrs.ContentType, '').trim();
            if (ext !== '' && type !== '') defaults.set(ext, type);
        } else if (node.name === 'Override') {
            const part = asString(node.attrs.PartName, '').replace(/^\//, '');
            const type = asString(node.attrs.ContentType, '').trim();
            if (part !== '' && type !== '') overrides.set(part, type);
        }
    }
    return { defaults, overrides };
}

/** 包内部件的内容类型：先看 Override，再按扩展名看 Default；都没有就是 undefined。 */
function contentTypeOf(contentTypes, partName) {
    const exact = contentTypes.overrides.get(partName);
    if (exact !== undefined) return exact;
    const at = partName.lastIndexOf('.');
    if (at === -1) return undefined;
    return contentTypes.defaults.get(partName.slice(at + 1).toLowerCase());
}

/** `ppt/media/image12.jpeg` → `{index: 12, ext: 'jpeg'}`；不是这个命名就返回 undefined。 */
function mediaPartIndexOf(partName) {
    const match = /^ppt\/media\/image(\d+)\.([A-Za-z0-9]+)$/.exec(partName);
    return match === null
        ? undefined
        : { index: Number.parseInt(match[1], 10), ext: match[2].toLowerCase() };
}

/** 从 `fromPart` 指向 `toPart` 的相对 Target（`ppt/slides/slide1.xml` → `../media/image1.png`）。 */
function relativeTarget(fromPart, toPart) {
    const from = fromPart.split('/').slice(0, -1);
    const to = toPart.split('/');
    let common = 0;
    while (common < from.length && common < to.length - 1 && from[common] === to[common]) common += 1;
    return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/');
}

/**
 * 包内文本的可写叠加层。
 *
 * `read()` 只解码不登记，`write()` 才进写回集合 —— revise 的既有约定是
 * 「没改过的部件一个字节都不动」，把读过的部件也塞进写回集合会破坏这条约定
 * （即使重新编码后的字节通常一样，也不该赌）。
 */
function createTextLayer(files, decoder) {
    const changed = new Map();
    return {
        read(partName) {
            if (changed.has(partName)) return changed.get(partName);
            const data = files.get(partName);
            return data === undefined ? undefined : decoder.decode(data);
        },
        write(partName, text) {
            changed.set(partName, text);
        },
        entries() {
            return changed;
        },
    };
}

// ───────────────────────────────────────────────────────────────────────────
// 形状定位（字符串层）
// ───────────────────────────────────────────────────────────────────────────

/**
 * 找出一个 slide XML 里形状的可改切片。
 *
 * 为什么要逐层做配对扫描而不是一把正则：`p:sp` / `p:pic` / `p:graphicFrame` 都可以嵌套
 * （组合形状、图形框里的表格），而 `p:txBody` 在 `p:sp` 与 `p:graphicFrame` 里都会出现。
 * 配对扫描能保证「截出来的区间一定是一个完整元素」，不会把父亲的一个子元素切一半。
 *
 * @returns {Array<{kind: string, start: number, end: number, tag: string, id: number|undefined, name: string, text: string}>}
 */
function sliceShapes(xml) {
    const out = [];
    const re = /<(\/?)(p:sp|p:pic|p:graphicFrame|p:grpSp|p:cxnSp)\b([^>]*?)(\/?)>/g;
    const stack = [];
    let match;
    while ((match = re.exec(xml)) !== null) {
        const [whole, closing, tag, attrs, selfClose] = match;
        if (closing === '/') {
            const open = stack.pop();
            // 只有最外层才登记：组合形状内部的小形状不该被单独改，否则会破坏组合的位置语义
            if (open !== undefined && stack.length === 0) out.push(finishShape(xml, open, match.index + whole.length));
            continue;
        }
        if (selfClose === '/') {
            if (stack.length === 0) out.push(finishShape(xml, { tag, attrs, start: match.index }, match.index + whole.length));
            continue;
        }
        stack.push({ tag, attrs, start: match.index });
    }
    return out.filter((shape) => shape !== undefined).sort((a, b) => a.start - b.start);
}

function finishShape(xml, open, end) {
    const block = xml.slice(open.start, end);
    const cNvPr = /<p:cNvPr\b[^>]*>/.exec(block)?.[0];
    const rawName = cNvPr === undefined ? '' : readAttr(cNvPr, 'name');
    return {
        kind: open.tag.slice(2),
        start: open.start,
        end,
        block,
        id: cNvPr === undefined ? undefined : Number.parseInt(readAttr(cNvPr, 'id') ?? '', 10),
        name: decodeEntities(rawName ?? ''),
        text: normalizeText(block),
    };
}

/** 形状里的纯文本：把 `a:t` 逐个取出来拼，段落之间不加分隔（对齐 read() 的口径）。 */
function normalizeText(block) {
    const root = parseXml(block).children[0];
    if (root === undefined) return '';
    return descendants(root, 'a:t').map((node) => node.text).join('').replace(/\s+/g, ' ').trim();
}

/** 找形状在 XML 里的 `a:xfrm`（`p:spPr` 下的那个，不是 `p:xfrm`）。 */
function findXfrm(xml, shape) {
    const within = xml.slice(shape.start, shape.end);
    const match = /<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>|<a:xfrm\b[^>]*\/>/.exec(within);
    if (match === null) return undefined;
    const offset = shape.start + match.index;
    const off = /<a:off\b[^>]*\/>/.exec(match[0])?.[0];
    const ext = /<a:ext\b[^>]*\/>/.exec(match[0])?.[0];
    return {
        start: offset,
        end: offset + match[0].length,
        block: match[0],
        off,
        ext,
        x: off === undefined ? undefined : Number.parseInt(readAttr(off, 'x') ?? '', 10),
        y: off === undefined ? undefined : Number.parseInt(readAttr(off, 'y') ?? '', 10),
        w: ext === undefined ? undefined : Number.parseInt(readAttr(ext, 'cx') ?? '', 10),
        h: ext === undefined ? undefined : Number.parseInt(readAttr(ext, 'cy') ?? '', 10),
    };
}

function findTxBody(xml, shape) {
    const within = xml.slice(shape.start, shape.end);
    const match = /<p:txBody>[\s\S]*?<\/p:txBody>/.exec(within);
    if (match === null) return undefined;
    return { start: shape.start + match.index, end: shape.start + match.index + match[0].length, block: match[0] };
}

/** 形状选择器 → 命中判定。选择器可以是 id / 名字 / 「第 N 页的那个」三类。 */
function makeMatcher(selector) {
    if (typeof selector === 'number' || (typeof selector === 'string' && /^\d+$/.test(selector.trim()))) {
        const id = Number.parseInt(String(selector).trim(), 10);
        return { label: `#${id}`, test: (shape) => shape.id === id };
    }
    const text = asString(selector, '').trim();
    if (text === '') return undefined;
    const kind = SHAPE_KINDS[text.toLowerCase()];
    if (kind !== undefined) {
        return { label: text, test: (shape) => kind.includes(shape.name) };
    }
    return { label: text, test: (shape) => shape.name === text };
}

// ───────────────────────────────────────────────────────────────────────────
// 改样式
// ───────────────────────────────────────────────────────────────────────────

const RPR_CHILD_ORDER = ['a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill',
    'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill',
    'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'];

/** 一个 `a:rPr` 里已有某个子元素吗。 */
function hasChild(rPr, name) {
    const re = new RegExp(`<${name}\\b`);
    return re.test(rPr);
}

function removeChild(rPr, name) {
    const re = new RegExp(`<${name}\\b[^>]*/>|<${name}\\b[^>]*>[\\s\\S]*?</${name}>`, 'g');
    return rPr.replace(re, '');
}

/** 按 schema 顺序把一段子元素插进 `a:rPr`。 */
function insertChild(rPr, name, childXml) {
    const openMatch = /^<a:rPr\b[^>]*?(\/?)>/.exec(rPr);
    if (openMatch === null) return rPr;
    if (openMatch[1] === '/') {
        // 自闭合形式：展开成开闭标签对，再把子元素放进去
        const open = openMatch[0].replace(/\/>$/, '>');
        return `${open}${childXml}</a:rPr>`;
    }
    const rank = RPR_CHILD_ORDER.indexOf(name);
    const bodyStart = openMatch[0].length;
    const body = rPr.slice(bodyStart, rPr.lastIndexOf('</a:rPr>'));
    const childRe = /<(a:[A-Za-z]+)\b/g;
    let insertAt = body.length;
    let scan;
    while ((scan = childRe.exec(body)) !== null) {
        const other = RPR_CHILD_ORDER.indexOf(scan[1]);
        if (other === -1 || rank === -1) continue;
        if (other > rank) { insertAt = scan.index; break; }
    }
    const open = rPr.slice(0, bodyStart);
    const rest = body.slice(insertAt);
    const head = body.slice(0, insertAt);
    return `${open}${head}${childXml}${rest}</a:rPr>`;
}

/**
 * 改一个 `a:rPr`。
 *
 * `patch` 的键：sizePt / bold / italic / underline / color / font / spacing。
 * 值为 `undefined` 表示这一项不动；`bold:false` 表示显式去掉加粗（不是「不动」）。
 */
function patchRPr(rPr, patch) {
    let out = rPr;
    const openMatch = /^<a:rPr\b[^>]*?(\/?)>/.exec(out);
    if (openMatch === null) return out;

    const attrs = {};
    if (patch.sizePt !== undefined) attrs.sz = String(Math.max(100, Math.round(asNumber(patch.sizePt, 18) * 100)));
    if (patch.spacing !== undefined) attrs.spc = String(Math.round(asNumber(patch.spacing, 0)));
    if (patch.bold !== undefined) attrs.b = patch.bold ? '1' : '0';
    if (patch.italic !== undefined) attrs.i = patch.italic ? '1' : '0';
    if (patch.underline !== undefined) attrs.u = patch.underline ? 'sng' : 'none';

    // 属性改动要落在开标签上；自闭合的 rPr 后面还可能被 insertChild 展开，
    // 所以先统一展开成开闭标签对，再往下走，避免两处各改一次开标签。
    if (openMatch[1] === '/') out = out.replace(/^<a:rPr\b[^>]*\/>/, (tag) => `${tag.replace(/\/>$/, '>')}</a:rPr>`);
    let open = /^<a:rPr\b[^>]*>/.exec(out)[0];
    for (const [name, value] of Object.entries(attrs)) open = setAttr(open, name, value);
    out = open + out.slice(/^<a:rPr\b[^>]*>/.exec(out)[0].length);

    if (patch.color !== undefined) {
        out = hasChild(out, 'a:solidFill') ? removeChild(out, 'a:solidFill') : out;
        out = insertChild(out, 'a:solidFill', `<a:solidFill><a:srgbClr val="${toHex(patch.color)}"/></a:solidFill>`);
    }
    if (patch.font !== undefined) {
        for (const tag of ['a:latin', 'a:ea', 'a:cs']) out = hasChild(out, tag) ? removeChild(out, tag) : out;
        const face = escapeAttr(asString(patch.font));
        out = insertChild(out, 'a:cs', `<a:cs typeface="${face}"/>`);
        out = insertChild(out, 'a:ea', `<a:ea typeface="${face}"/>`);
        out = insertChild(out, 'a:latin', `<a:latin typeface="${face}"/>`);
    }
    return out;
}

/** 该不该按段内子串改：`text` 给的是子串时只改命中的 run，其余段落不动。 */
function applyStyleToBlock(block, patch, target) {
    // 段落级属性：a:pPr 上的对齐与段间距
    let out = block;
    if (patch.align !== undefined || patch.lineSpacing !== undefined) {
        const pPrMatch = /<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/.exec(out);
        if (pPrMatch !== null) {
            let pPr = pPrMatch[0];
            if (/<a:pPr\b[^>]*\/>/.test(pPr)) pPr = `${pPr.replace(/\/>$/, '>')}</a:pPr>`;
            let open = /^<a:pPr\b[^>]*>/.exec(pPr)[0];
            if (patch.align !== undefined) {
                const align = TEXT_ALIGN[asString(patch.align, '').trim().toLowerCase()];
                if (align !== undefined) open = setAttr(open, 'algn', align);
            }
            pPr = open + pPr.slice(/^<a:pPr\b[^>]*>/.exec(pPr)[0].length);
            if (patch.lineSpacing !== undefined) {
                const lnSpc = `<a:lnSpc><a:spcPct val="${Math.round(Math.max(0.1, asNumber(patch.lineSpacing, 1.2)) * 100000)}"/></a:lnSpc>`;
                pPr = /<a:lnSpc\b[\s\S]*?<\/a:lnSpc>/.test(pPr)
                    ? pPr.replace(/<a:lnSpc\b[\s\S]*?<\/a:lnSpc>/, lnSpc)
                    : pPr.replace(/^(<a:pPr\b[^>]*>)/, `$1${lnSpc}`);
            }
            out = out.slice(0, pPrMatch.index) + pPr + out.slice(pPrMatch.index + pPrMatch[0].length);
        }
    }

    // run 级属性：整段命中，或按子串只改命中的 run
    const runPatch = { ...patch };
    delete runPatch.align;
    delete runPatch.lineSpacing;
    const hasRunPatch = Object.values(runPatch).some((value) => value !== undefined);
    if (!hasRunPatch) return out;

    if (target === undefined) {
        return out.replace(RUN_RE, (run) => {
            const rPr = RPR_RE.exec(run)?.[0];
            if (rPr === undefined) return run;
            return run.replace(rPr, patchRPr(rPr, runPatch));
        });
    }
    // 子串模式：run 里的 a:t 命中才改这个 run
    return out.replace(RUN_RE, (run) => {
        const tMatch = /<a:t[^>]*>([\s\S]*?)<\/a:t>/.exec(run);
        if (tMatch === null || !decodeEntities(tMatch[1]).includes(target)) return run;
        const rPr = RPR_RE.exec(run)?.[0];
        if (rPr === undefined) return run;
        return run.replace(rPr, patchRPr(rPr, runPatch));
    });
}

// ───────────────────────────────────────────────────────────────────────────
// 对外：读页面结构
// ───────────────────────────────────────────────────────────────────────────

/**
 * 打开一个已有 .pptx，列出每一页的每个形状（含 id、类型、名字、几何、文字）。
 *
 * 这是「微调」的第一步：改之前先看清楚有什么、在第几页、长什么样。
 *
 * `kind === 'pic'` 的形状额外带 `image`：
 * `{relId, part, linked, contentType, bytes, width, height, crop}` —— 部件名与 relId
 * 是接着去 `office.ppt.images()` 把这张图抽出来看的入口；外链图片 `linked: true`，
 * `part` 是链出去的目标、`bytes` 为 null。图片形状的 `editable` 一律是 false。
 */
export function readSlides(path, env) {
    const bytes = env.readFile(path);
    const files = unzip(bytes);
    const decoder = new TextDecoder('utf-8');
    const parts = orderedSlideParts(files, decoder);
    const layoutNames = layoutNameByPart(files, decoder);
    const contentTypes = parseContentTypes(partText(files, decoder, '[Content_Types].xml'));
    const pages = parts.map((partName, index) => {
        const data = files.get(partName);
        if (data === undefined) {
            return { index: index + 1, part: partName, layout: '', shapes: [], missing: true };
        }
        const xml = decoder.decode(data);
        const rels = parseRels(partText(files, decoder, relsPartOf(partName)));
        const shapes = sliceShapes(xml).map((shape) => {
            const xfrm = findXfrm(xml, shape);
            const txBody = findTxBody(xml, shape);
            const paragraphs = txBody === undefined ? [] : (txBody.block.match(PARA_RE) ?? []);
            const info = {
                id: shape.id,
                kind: shape.kind,
                name: shape.name,
                text: shape.text,
                paragraphs: paragraphs.length,
                x: xfrm?.x === undefined ? null : emuToCm(xfrm.x),
                y: xfrm?.y === undefined ? null : emuToCm(xfrm.y),
                w: xfrm?.w === undefined ? null : emuToCm(xfrm.w),
                h: xfrm?.h === undefined ? null : emuToCm(xfrm.h),
                editable: txBody !== undefined && xfrm !== undefined,
            };
            // 图片：把「这张图到底是包里的哪个部件、多大、有没有裁切」一并报出来。
            // 只有部件名与 relId 才能让调用方接着去 office.ppt.images() 把它抽出来看。
            if (shape.kind === 'pic') {
                const image = describePicture(files, partName, rels, contentTypes, shape);
                if (image !== undefined) info.image = image;
            }
            return info;
        }).filter((shape) => shape.name !== '' || shape.kind !== 'sp');
        return {
            index: index + 1,
            part: partName,
            layout: layoutNames.get(partName) ?? '',
            shapes,
        };
    });
    return {
        ok: true,
        format: 'pptx',
        path: asString(path),
        bytes: bytes.length,
        pages,
        counts: {
            pages: pages.length,
            shapes: pages.reduce((sum, page) => sum + page.shapes.length, 0),
            textShapes: pages.reduce((sum, page) => sum + page.shapes.filter((shape) => shape.text !== '').length, 0),
        },
    };
}

/**
 * 一个 `p:pic` 背后的图片事实：部件名、关系 id、内容类型、字节数、像素尺寸、裁切量。
 *
 * 外链图片（`a:blip/@r:link`，或关系写了 `TargetMode="External"`）没有包内部件，
 * 这里照实报 `linked: true` 并把 `part` 留成链出去的目标 —— 编一个包内部件名
 * 会让调用方拿着一个读不出来的路径去 read_image。
 */
function describePicture(files, slidePart, rels, contentTypes, shape) {
    const blip = /<a:blip\b[^>]*>/.exec(shape.block)?.[0];
    if (blip === undefined) return undefined;
    const embed = readAttr(blip, 'r:embed');
    const link = readAttr(blip, 'r:link');
    const relId = embed ?? link;
    if (relId === undefined) return undefined;
    const rel = rels.get(relId);
    const linked = link !== undefined || rel?.mode === 'External';
    const part = rel === undefined
        ? undefined
        : (linked ? asString(rel.target) : resolveTarget(slidePart, rel.target));
    const data = linked || part === undefined ? undefined : files.get(part);
    const info = data === undefined ? undefined : imageInfo(Buffer.from(data));
    const srcRect = /<a:srcRect\b[^>]*\/>/.exec(shape.block)?.[0];
    const crop = { l: 0, t: 0, r: 0, b: 0 };
    if (srcRect !== undefined) {
        for (const key of ['l', 't', 'r', 'b']) {
            crop[key] = Number.parseInt(readAttr(srcRect, key) ?? '0', 10) || 0;
        }
    }
    return {
        relId,
        part: part ?? null,
        linked,
        contentType: data === undefined ? null : (contentTypeOf(contentTypes, part) ?? null),
        bytes: data === undefined ? null : data.length,
        width: info?.width ?? null,
        height: info?.height ?? null,
        crop,
    };
}

/** 每一页挂的版式名（拿来做「这是封面页」这类判断）。 */
function layoutNameByPart(files, decoder) {
    const out = new Map();
    for (const name of files.keys()) {
        if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
        const rels = parseRels(files.get(relsPartOf(name)) === undefined ? undefined : decoder.decode(files.get(relsPartOf(name))));
        for (const rel of rels.values()) {
            if (rel.type !== RT_LAYOUT) continue;
            const layoutPart = resolveTarget(name, rel.target);
            const xml = layoutPart === undefined ? undefined : files.get(layoutPart);
            if (xml === undefined) continue;
            const cSld = descendants(parseXml(decoder.decode(xml)).children[0], 'p:cSld')[0];
            out.set(name, cSld?.attrs?.name ?? '');
        }
    }
    return out;
}

// ───────────────────────────────────────────────────────────────────────────
// 对外：修订
// ───────────────────────────────────────────────────────────────────────────

/**
 * 就地修订一个已有 .pptx。
 *
 * 支持的操作（`ops` 数组，逐条执行，互不影响；一条失败只进 skipped）：
 *
 * | 写法 | 作用 |
 * | --- | --- |
 * | `{ slide, shape, setText: '新文字' }` | 整段重写（rPr 保留，字体字号不变；`\n` 拆成多段） |
 * | `{ slide, shape, replace: [['旧','新']] }` | 只替换段内命中子串，段落结构不动 |
 * | `{ slide, shape, style: { sizePt, bold, color, font, align, lineSpacing, text } }` | 改字体、字号、颜色、对齐；`text` 给子串则只改命中的 run |
 * | `{ slide, shape, move: { x, y } }` | 改位置（厘米，原点页面左上角；给 dx/dy 则相对移动） |
 * | `{ slide, shape, resize: { w, h } }` | 改尺寸（厘米；给 dw/dh 则相对缩放） |
 * | `{ slide, shape, fit: true }` | 让文本框高度贴合内容（按当前字号估算） |
 * | `{ slide, addImage: { path, x, y, wCm, hCm, fit, name, alt } }` | 往该页插一张图（包级操作，不需要 `shape`） |
 *
 * `slide` / `shape` 是选择器：slide 是页码（1 起）或 `'2-5'` / `[2,3]`，
 * shape 是 id / 形状名（`'Card 1 Body'`）/ 类型名（`title` / `subtitle` / …）。
 *
 * 插图与改文字可以放在同一批 ops 里：插图那条走 `insertPictureIntoPage()`（要同时动
 * slide XML、rels、[Content_Types].xml 与媒体部件），改文字那条走 `applyOp()`（只动一个形状的切片）。
 */
export function revise(path, ops, env) {
    const list = asArray(ops);
    const bytes = env.readFile(path);
    const files = unzip(bytes);
    const decoder = new TextDecoder('utf-8');
    const encoder = new TextEncoder();
    const parts = orderedSlideParts(files, decoder);
    const size = slideSize(files, decoder);
    const warnings = [];
    const applied = [];
    const skipped = [];

    /** 每一页的 XML 在内存里改，最后按页一次性写回。 */
    const byPart = new Map();

    const xmlOf = (partName) => {
        if (!byPart.has(partName)) byPart.set(partName, decoder.decode(files.get(partName)));
        return byPart.get(partName);
    };

    // 插图要改的部件不止 slide XML（还有 slide rels 与 [Content_Types].xml），
    // 走独立的一层文本叠加：只有真正写过的部件才进写回集合。
    const layer = createTextLayer(files, decoder);
    let pictureState;
    const pictureStateOf = () => {
        pictureState ??= {
            ...createPictureState({ files, layer, parts, env, warnings, size }),
            layer,
            parts,
        };
        return pictureState;
    };

    list.forEach((rawOp, position) => {
        const where = `ops[${position}]`;
        const op = rawOp !== null && typeof rawOp === 'object' ? rawOp : null;
        if (op === null) {
            skipped.push({ op: where, reason: '操作不是对象' });
            return;
        }
        const pageNumbers = resolvePages(op.slide, parts.length);
        if (pageNumbers.length === 0) {
            skipped.push({ op: where, reason: `slide 选择器 ${JSON.stringify(op.slide)} 没匹配到任何页（共 ${parts.length} 页）` });
            return;
        }
        // 插图是包级操作（slide XML + rels + [Content_Types].xml + 新媒体部件），
        // 落在不了 applyOp() 的「单个形状 XML 切片」模型里，所以在这里单独分流。
        // 它不需要 shape 选择器：新图的形状 id 由该页最大 id 顺延。
        if (op.addImage !== undefined) {
            const spec = op.addImage !== null && typeof op.addImage === 'object' ? op.addImage : {};
            if (op.shape !== undefined) {
                // 静默忽略用户写的东西比报一句更糟：他以为自己在改某个形状
                warnings.push(`${where}：addImage 不需要 shape 选择器，已忽略 shape=${JSON.stringify(op.shape)}`);
            }
            for (const pageNumber of pageNumbers) {
                const result = insertPictureIntoPage(pictureStateOf(), pageNumber, spec);
                if (result.ok !== true) {
                    skipped.push({ op: where, page: pageNumber, reason: result.reason });
                    continue;
                }
                applied.push({ op: 'insertImage', page: pageNumber, ...result.detail });
            }
            return;
        }
        const matcher = makeMatcher(op.shape);
        if (matcher === undefined) {
            skipped.push({ op: where, reason: 'shape 必须给 id、形状名或类型名（title/subtitle/…）' });
            return;
        }

        let hits = 0;
        let candidates = 0;
        for (const pageNumber of pageNumbers) {
            const partName = parts[pageNumber - 1];
            let xml = xmlOf(partName);
            const shapes = sliceShapes(xml).filter((shape) => matcher.test(shape));
            if (shapes.length === 0) continue;
            candidates += shapes.length;
            if (shapes.length > 1) {
                warnings.push(`第 ${pageNumber} 页有 ${shapes.length} 个形状匹配 ${matcher.label}，已全部处理；要只改一个请用形状 id`);
            }
            // 从后往前改：前面的替换会移动后面形状的下标
            for (const shape of shapes.slice().sort((a, b) => b.start - a.start)) {
                const result = applyOp(xml, shape, op, size, pageNumber);
                if (result.ok !== true) {
                    skipped.push({ op: where, page: pageNumber, shape: shape.name || `#${shape.id}`, reason: result.reason });
                    continue;
                }
                for (const message of result.warnings ?? []) warnings.push(`第 ${pageNumber} 页 ${shape.name || `#${shape.id}`}：${message}`);
                if (result.xml !== undefined) xml = result.xml;
                hits += 1;
                applied.push({
                    op: result.label,
                    page: pageNumber,
                    shape: shape.name || `#${shape.id}`,
                    id: shape.id,
                    ...(result.detail ?? {}),
                });
            }
            byPart.set(partName, xml);
        }
        // 形状找到了却一条都没落地时，不要再叠一句「找不到形状」——
        // 真正的原因（没有 p:txBody / 没有 a:xfrm / 找不到子串）已经在上面那条 skipped 里。
        if (hits === 0 && candidates === 0) {
            const onPages = pageNumbers.length === parts.length ? '整篇' : `第 ${pageNumbers.join('、')} 页`;
            skipped.push({ op: where, reason: `${onPages}找不到形状 ${matcher.label}；先用 office.ppt.readSlides(path) 看形状名与 id` });
        }
    });

    for (const [partName, xml] of byPart) files.set(partName, encoder.encode(xml));
    // 插图改过的 slide XML / rels / [Content_Types].xml 同样按文本写回
    for (const [partName, text] of layer.entries()) files.set(partName, encoder.encode(text));

    if (applied.length === 0) {
        return {
            ok: false,
            format: 'pptx',
            path: asString(path),
            saved: false,
            applied,
            skipped,
            warnings: [...warnings, '没有任何操作生效，文件未改写'],
        };
    }

    const written = env.writeFile(path, zip(files));
    return {
        ok: true,
        format: 'pptx',
        path: written.path,
        saved: true,
        bytes: written.bytes,
        applied,
        skipped,
        warnings,
    };
}

function slideSize(files, decoder) {
    const data = files.get('ppt/presentation.xml');
    if (data === undefined) return DEFAULT_SLIDE;
    const node = descendants(parseXml(decoder.decode(data)).children[0], 'p:sldSz')[0];
    const cx = Number.parseInt(node?.attrs?.cx ?? '', 10);
    const cy = Number.parseInt(node?.attrs?.cy ?? '', 10);
    return {
        cx: Number.isFinite(cx) && cx > 0 ? cx : DEFAULT_SLIDE.cx,
        cy: Number.isFinite(cy) && cy > 0 ? cy : DEFAULT_SLIDE.cy,
    };
}

/** `slide` 选择器：1 / '2' / '2-5' / '3,7' / [2,3] / 省略（= 全篇）。 */
function resolvePages(selector, total) {
    if (selector === undefined || selector === null || selector === '') {
        return Array.from({ length: total }, (_, index) => index + 1);
    }
    if (Array.isArray(selector)) {
        return [...new Set(selector.map((item) => Number.parseInt(String(item), 10))
            .filter((n) => Number.isInteger(n) && n >= 1 && n <= total))].sort((a, b) => a - b);
    }
    const text = String(selector).trim();
    if (/^\d+$/.test(text)) {
        const n = Number.parseInt(text, 10);
        return n >= 1 && n <= total ? [n] : [];
    }
    const out = new Set();
    for (const piece of text.split(/[,，\s]+/).filter((item) => item !== '')) {
        const range = /^(\d+)\s*[-~—]\s*(\d+)$/.exec(piece);
        if (range !== null) {
            const from = Math.max(1, Number.parseInt(range[1], 10));
            const to = Math.min(total, Number.parseInt(range[2], 10));
            for (let n = from; n <= to; n += 1) out.add(n);
            continue;
        }
        const n = Number.parseInt(piece, 10);
        if (Number.isInteger(n) && n >= 1 && n <= total) out.add(n);
    }
    return [...out].sort((a, b) => a - b);
}

// ───────────────────────────────────────────────────────────────────────────
// 对外：往已有包里插图
// ───────────────────────────────────────────────────────────────────────────

/**
 * 插图的状态机：媒体编号、内容类型声明、每页关系 id 的自增都集中在这里。
 *
 * 为什么不能复用 `applyOp()`：那个函数只在**一个形状的 XML 切片**里做定点替换，
 * 而插图要同时动四处 —— slide XML（追加 `p:pic`）、slide rels（加一条关系）、
 * `[Content_Types].xml`（补扩展名 Default）、以及包里新增一个媒体部件。
 * 三处文本与一个二进制部件必须一起成功，否则会写出一个「有图框但没图」的坏包。
 */
function createPictureState(input) {
    const { files, layer, parts, env, warnings, size } = input;
    const ctPart = '[Content_Types].xml';
    const ctText = layer.read(ctPart);
    const ct = ctText === undefined ? undefined : parseContentTypes(ctText);
    const usedNames = new Set([...files.keys()].filter((name) => name.startsWith(MEDIA_PREFIX)));
    let maxIndex = 0;
    for (const name of usedNames) {
        const parsed = mediaPartIndexOf(name);
        if (parsed !== undefined) maxIndex = Math.max(maxIndex, parsed.index);
    }
    // 同一份字节只登记一个部件：同一张图插到几页里，包里不该出现几份副本
    const byDigest = new Map();

    /** 扩展名没有 Default 时补一条。已有 Override 的包也补：Default 覆盖的是别的同名部件。 */
    function declareExtension(ext, contentType) {
        if (ct.defaults.has(ext)) return;
        ct.defaults.set(ext, contentType);
        const node = `<Default Extension="${escapeAttr(ext)}" ContentType="${escapeAttr(contentType)}"/>`;
        const current = layer.read(ctPart) ?? '';
        const at = current.indexOf('<Override');
        if (at !== -1) {
            layer.write(ctPart, `${current.slice(0, at)}${node}${current.slice(at)}`);
            return;
        }
        // 没有 Override 时插在 </Types> 之前；连 </Types> 都没有就只好追加（源文件本身已经坏了）
        const closed = current.replace(/<\/Types>\s*$/, `${node}</Types>`);
        layer.write(ctPart, closed === current ? `${current}${node}` : closed);
    }

    return {
        warnings,
        size,
        /** 把一张图登记成媒体部件（字节原样进包），返回 `{ok, partName, info}`。 */
        addMedia(buffer, info) {
            if (ct === undefined) {
                return { ok: false, reason: '包里没有 [Content_Types].xml，无法声明图片的内容类型（这个包本身已经不合规）' };
            }
            const digest = `${info.ext}:${createHash('sha1').update(buffer).digest('hex')}`;
            const known = byDigest.get(digest);
            if (known !== undefined) return { ok: true, ...known };
            // 编号从包里已有的 imageN 顺延；名字被占用就往后找，绝不覆盖已有部件
            do { maxIndex += 1; } while (usedNames.has(`${MEDIA_PREFIX}image${maxIndex}.${info.ext}`));
            const partName = `${MEDIA_PREFIX}image${maxIndex}.${info.ext}`;
            usedNames.add(partName);
            files.set(partName, buffer);
            byDigest.set(digest, { partName, info });
            declareExtension(info.ext, info.contentType);
            return { ok: true, partName, info };
        },
        /**
         * 在某一页的 rels 里加一条指向该媒体部件的关系，返回新关系 id。
         * 返回 undefined 表示这一页的 rels 结构异常 —— 宁可失败也不能留下悬空的 r:embed。
         */
        addImageRel(pageNumber, partName) {
            const slidePart = parts[pageNumber - 1];
            const relsPart = relsPartOf(slidePart);
            const text = layer.read(relsPart)
                ?? `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}"></Relationships>`;
            if (!text.includes('</Relationships>')) return undefined;
            let maxId = 0;
            const used = new Set();
            for (const match of text.matchAll(/<Relationship\b[^>]*\sId="([^"]+)"/g)) {
                used.add(match[1]);
                const numbered = /^rId(\d+)$/.exec(match[1]);
                if (numbered !== null) maxId = Math.max(maxId, Number.parseInt(numbered[1], 10));
            }
            let relId;
            do { maxId += 1; relId = `rId${maxId}`; } while (used.has(relId));
            const node = `<Relationship Id="${relId}" Type="${RT_IMAGE}"`
                + ` Target="${escapeAttr(relativeTarget(slidePart, partName))}"/>`;
            const closed = text.replace(/<\/Relationships>\s*$/, `${node}</Relationships>`);
            layer.write(relsPart, closed === text
                ? text.replace('</Relationships>', `${node}</Relationships>`)
                : closed);
            return relId;
        },
        env,
    };
}

/** 该页下一个可用的 `p:cNvPr/@id`：同一页里 id 必须唯一，取最大值 + 1。 */
function nextShapeId(xml) {
    let max = 0;
    for (const match of xml.matchAll(/<p:cNvPr\b[^>]*\sid="(\d+)"/g)) {
        const value = Number.parseInt(match[1], 10);
        if (Number.isFinite(value)) max = Math.max(max, value);
    }
    return max + 1;
}

/**
 * 把新形状追加到 `p:spTree` 的末尾（文档顺序 = z 序，最后画的在最上面）。
 * `p:spTree` 允许以 `p:extLst` 收尾，而 `p:extLst` 必须排在最后，所以新形状要插在它之前。
 */
function appendToSpTree(xml, node) {
    const closeAt = xml.lastIndexOf('</p:spTree>');
    if (closeAt === -1) return undefined;
    const head = xml.slice(0, closeAt);
    const ext = /<p:extLst\b[\s\S]*?<\/p:extLst>\s*$/.exec(head);
    const at = ext === null ? closeAt : ext.index;
    return `${xml.slice(0, at)}${node}${xml.slice(at)}`;
}

/** cover 的裁切量：与 ppt.js 的 cropOf 同一套算法（差 0.5% 以内算作 0，封顶 99999）。 */
function cropOf(drawnW, drawnH, areaW, areaH) {
    const percent = (part, whole) => {
        if (!(whole > 0)) return 0;
        const value = Math.round((part / whole) * 100000);
        return value <= 500 ? 0 : Math.min(SRC_RECT_MAX, value);
    };
    return {
        l: percent((drawnW - areaW) / 2, drawnW),
        t: percent((drawnH - areaH) / 2, drawnH),
        r: percent((drawnW - areaW) / 2, drawnW),
        b: percent((drawnH - areaH) / 2, drawnH),
    };
}

function srcRectXml(crop) {
    if (crop === undefined || (crop.l === 0 && crop.t === 0 && crop.r === 0 && crop.b === 0)) return '';
    return `<a:srcRect l="${crop.l}" t="${crop.t}" r="${crop.r}" b="${crop.b}"/>`;
}

/** fit 合法性检查：非法值不抛，记一条 warning 后回落 contain（与 ppt.js 同口径）。 */
function imageFitOf(raw, warnings) {
    const value = asString(raw, '').trim().toLowerCase();
    if (value === '') return 'contain';
    if (IMAGE_FITS.includes(value)) return value;
    warnings.push(`未知的图片适配方式「${asString(raw)}」，已按 contain 处理（可用：contain / cover / natural）`);
    return 'contain';
}

/**
 * 一张图在页面上的落点。
 * - contain：等比缩放进 wCm×hCm 的框并居中；
 * - cover：铺满整个框，多出来的边转成 a:srcRect 裁掉；
 * - natural：按 96 DPI 把原始像素折算成 EMU，不缩放（此时框只当原点用）。
 * x/y 单位厘米、原点页面左上角，与 deck.shape / revise.move 完全一致。
 */
function pictureGeometry(spec, info, fit, warnings) {
    const naturalW = Math.max(1, info.width * EMU_PER_PX);
    const naturalH = Math.max(1, info.height * EMU_PER_PX);
    const aspect = naturalW / naturalH;
    const x = cmToEmu(spec.x);
    const y = cmToEmu(spec.y);
    let boxW = asNumber(spec.wCm, 0) > 0 ? cmToEmu(spec.wCm) : 0;
    let boxH = asNumber(spec.hCm, 0) > 0 ? cmToEmu(spec.hCm) : 0;
    if (boxW > 0 && boxH <= 0) boxH = Math.round(boxW / aspect);
    if (boxH > 0 && boxW <= 0) boxW = Math.round(boxH * aspect);
    if (boxW <= 0 && boxH <= 0) {
        boxW = naturalW;
        boxH = naturalH;
    }
    if (fit === 'natural') {
        if (asNumber(spec.wCm, 0) > 0 || asNumber(spec.hCm, 0) > 0) {
            warnings.push('fit 为 natural 时按原图像素尺寸放置，wCm/hCm 只当原点用、不参与缩放');
        }
        return { x, y, w: naturalW, h: naturalH, crop: undefined };
    }
    if (fit === 'cover') {
        const scale = Math.max(boxW / naturalW, boxH / naturalH);
        return { x, y, w: boxW, h: boxH, crop: cropOf(naturalW * scale, naturalH * scale, boxW, boxH) };
    }
    const scale = Math.min(boxW / naturalW, boxH / naturalH);
    const w = Math.max(9525, Math.round(naturalW * scale));
    const h = Math.max(9525, Math.round(naturalH * scale));
    return { x: x + Math.round((boxW - w) / 2), y: y + Math.round((boxH - h) / 2), w, h, crop: undefined };
}

/**
 * 往一页里插一张图，返回 `{ok, detail}` 或 `{ok:false, reason}`。
 *
 * 顺序有讲究：先把「这一页能不能改」全部检查完，再登记媒体与关系。
 * 反过来做的话，一次失败的插图会在包里留下一个没人引用的媒体部件与一条悬空关系。
 */
function insertPictureIntoPage(state, pageNumber, spec) {
    const slidePart = state.parts[pageNumber - 1];
    if (slidePart === undefined) return { ok: false, reason: `第 ${pageNumber} 页不存在` };
    const sourcePath = asString(spec.path);
    if (sourcePath === '') return { ok: false, reason: '插图需要 path（图片文件路径）' };
    let buffer;
    try {
        buffer = Buffer.from(state.env.readFile(sourcePath));
    } catch {
        return { ok: false, reason: `找不到图片文件：${sourcePath}` };
    }
    const info = imageInfo(buffer);
    if (info === undefined) {
        return { ok: false, reason: `图片格式不支持（只认 PNG/JPEG/GIF）：${sourcePath}` };
    }
    const xml = state.layer.read(slidePart);
    if (xml === undefined) return { ok: false, reason: `第 ${pageNumber} 页的部件不在包里：${slidePart}` };
    if (!xml.includes('</p:spTree>')) {
        return { ok: false, reason: `第 ${pageNumber} 页没有 p:spTree，无法插图（部件结构异常）` };
    }
    const media = state.addMedia(buffer, info);
    if (media.ok !== true) return { ok: false, reason: media.reason };
    const relId = state.addImageRel(pageNumber, media.partName);
    if (relId === undefined) {
        return { ok: false, reason: `第 ${pageNumber} 页的 rels 部件结构异常（没有 </Relationships>），无法登记图片关系` };
    }
    const fit = imageFitOf(spec.fit, state.warnings);
    const box = pictureGeometry(spec, info, fit, state.warnings);
    const id = nextShapeId(xml);
    const name = asString(spec.name) === '' ? `Picture ${id}` : asString(spec.name);
    const alt = asString(spec.alt);
    const picture = `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${escapeAttr(name)}"`
        + `${alt === '' ? '' : ` descr="${escapeAttr(alt)}"`}/>`
        + '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        + `<p:blipFill><a:blip r:embed="${relId}"/>${srcRectXml(box.crop)}`
        + '<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        + '<p:spPr><a:xfrm>'
        + `<a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/>`
        + `<a:ext cx="${Math.round(box.w)}" cy="${Math.round(box.h)}"/></a:xfrm>`
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'
        + '</p:spPr></p:pic>';
    const next = appendToSpTree(xml, picture);
    if (next === undefined) {
        return { ok: false, reason: `第 ${pageNumber} 页没有 p:spTree，无法插图（部件结构异常）` };
    }
    state.layer.write(slidePart, next);
    // 超页只在 warnings 里报（与 revise 的 move/resize 同口径）：图放得出去、文件也照样合法，
    // 但「插完看不见」是调用方最需要提前知道的事 —— natural 不缩放时最容易踩到。
    if (box.x < 0 || box.y < 0 || box.x + box.w > state.size.cx || box.y + box.h > state.size.cy) {
        state.warnings.push(`第 ${pageNumber} 页插入的图片超出页面（x ${emuToCm(box.x)}cm, y ${emuToCm(box.y)}cm,`
            + ` ${emuToCm(box.w)}×${emuToCm(box.h)}cm，页面只有 ${emuToCm(state.size.cx)}×${emuToCm(state.size.cy)}cm）`);
    }
    return {
        ok: true,
        detail: {
            id,
            part: media.partName,
            relId,
            name,
            fit,
            path: sourcePath,
            x: emuToCm(box.x),
            y: emuToCm(box.y),
            w: emuToCm(box.w),
            h: emuToCm(box.h),
        },
    };
}

/**
 * 往已有 .pptx 里插图（就地改包，既有形状一个字节都不动）。
 *
 * `items` 可传单个对象或数组：`{slide, path, x, y, wCm, hCm, fit, name, alt}`。
 * 缺文件、格式不支持、页不存在都只进 `skipped`（带原因），一条失败不影响其它条；
 * 一张图都没插进去时不落盘、`ok:false` —— 与 revise 同一条「没改动就不写盘」的约定。
 */
export function insertImages(path, items, env) {
    const list = Array.isArray(items)
        ? items
        : (items === null || items === undefined ? [] : [items]);
    const bytes = env.readFile(path);
    const files = unzip(bytes);
    const decoder = new TextDecoder('utf-8');
    const encoder = new TextEncoder();
    const parts = orderedSlideParts(files, decoder);
    const layer = createTextLayer(files, decoder);
    const warnings = [];
    const applied = [];
    const skipped = [];
    const state = {
        ...createPictureState({ files, layer, parts, env, warnings, size: slideSize(files, decoder) }),
        layer,
        parts,
    };

    list.forEach((raw, position) => {
        const where = `items[${position}]`;
        const spec = raw !== null && typeof raw === 'object' ? raw : null;
        if (spec === null) {
            skipped.push({ op: where, reason: 'item 不是对象' });
            return;
        }
        const pageNumbers = resolvePages(spec.slide, parts.length);
        if (pageNumbers.length === 0) {
            skipped.push({
                op: where,
                reason: `slide 选择器 ${JSON.stringify(spec.slide)} 没匹配到任何页（共 ${parts.length} 页）`,
            });
            return;
        }
        for (const pageNumber of pageNumbers) {
            const result = insertPictureIntoPage(state, pageNumber, spec);
            if (result.ok !== true) {
                skipped.push({ op: where, page: pageNumber, reason: result.reason });
                continue;
            }
            applied.push({ op: 'insertImage', page: pageNumber, ...result.detail });
        }
    });

    for (const [partName, text] of layer.entries()) files.set(partName, encoder.encode(text));

    if (applied.length === 0) {
        return {
            ok: false,
            format: 'pptx',
            path: asString(path),
            saved: false,
            applied,
            skipped,
            warnings: [...warnings, '没有任何图片插入，文件未改写'],
        };
    }
    const written = env.writeFile(path, zip(files));
    return {
        ok: true,
        format: 'pptx',
        path: written.path,
        saved: true,
        bytes: written.bytes,
        applied,
        skipped,
        warnings,
    };
}

/** 执行一条操作。返回 `{ok, xml?, label, detail?, warnings?, reason?}`。 */
function applyOp(xml, shape, op, size, page) {
    const before = xml.slice(shape.start, shape.end);
    let block = before;

    // shape 在 sliceShapes 里是「相对整页 XML」的坐标；进入本函数后 block 是它的切片，
    // 所有定位都改用相对坐标（从 0 开始），这样两次 findTxBody 不会互相打架。
    const local = { kind: shape.kind, start: 0, end: block.length, block, name: shape.name, id: shape.id };
    const txBodyAtStart = findTxBody(block, local);
    const xfrmAtStart = findXfrm(block, local);
    const needsText = op.setText !== undefined || op.replace !== undefined || op.style !== undefined || op.fit === true;
    if (needsText && txBodyAtStart === undefined) {
        return {
            ok: false,
            reason: '这个形状没有文本（没有 p:txBody），改不了文字/字体；'
                + '图片、表格、组合形状的位置要改请用 move/resize，或改用 deck.shape 画一个文本框',
        };
    }
    if ((op.move !== undefined || op.resize !== undefined) && xfrmAtStart === undefined) {
        return { ok: false, reason: '这个形状没有 a:xfrm，本引擎无法改它的几何' };
    }

    // 1) style（可在 setText 之前，也可能只改样式）
    if (op.style !== undefined) {
        const style = op.style !== null && typeof op.style === 'object' ? op.style : {};
        const patch = normalizeStylePatch(style);
        if (Object.keys(patch).length === 0) return { ok: false, reason: 'style 里没有任何可识别的项（sizePt/bold/italic/underline/color/font/align/lineSpacing/text）' };
        const target = patch.text === undefined ? undefined : asString(patch.text);
        delete patch.text;
        const next = applyStyleToBlock(txBodyAtStart.block, patch, target);
        if (next === txBodyAtStart.block) return { ok: false, reason: target === undefined ? '样式与现状相同，没有产生改动' : `形状文本里找不到「${target}」` };
        block = splice(block, txBodyAtStart.start, txBodyAtStart.end, next);
    }

    // 2) setText / replace：都作用在当前的 txBody 上
    if (op.setText !== undefined || op.replace !== undefined) {
        const txBody = findTxBody(block, local);
        if (txBody === undefined) return { ok: false, reason: '这个形状没有文本（没有 p:txBody）' };
        let next = txBody.block;
        if (op.setText !== undefined) {
            // 换段落但保留 p:txBody / a:bodyPr / a:lstStyle 外壳，见 replaceTextBodyParagraphs
            const rewritten = replaceTextBodyParagraphs(next, asString(op.setText));
            if (rewritten === undefined) return { ok: false, reason: '这个形状的 p:txBody 结构异常（不是 <p:txBody>…</p:txBody>），拒绝改写' };
            next = rewritten;
        }
        for (const pair of normalizePairs(op.replace)) {
            if (pair.find === '') return { ok: false, reason: 'replace 的 find 不能为空' };
            // 只在 a:t 文本里替换：直接对整段 XML 做替换会误伤属性值
            const swapped = replaceInTextNodes(next, pair.find, pair.replace, pair.all !== false);
            if (swapped.count === 0) return { ok: false, reason: `形状文本里找不到「${pair.find}」` };
            next = swapped.xml;
        }
        block = splice(block, txBody.start, txBody.end, next);
    }

    // 3) fit：按内容重算文本框高度
    if (op.fit === true) {
        const txBody = findTxBody(block, local);
        const xfrm = findXfrm(block, local);
        if (txBody === undefined || xfrm?.off === undefined || xfrm.ext === undefined) {
            return { ok: false, reason: '这个形状没有文本或没有 a:xfrm，无法按内容调整高度' };
        }
        const height = estimateTextHeight(txBody.block, xfrm.w);
        const nextExt = setAttr(xfrm.ext, 'cy', String(height));
        block = splice(block, xfrm.start, xfrm.end, xfrm.block.replace(xfrm.ext, nextExt));
    }

    // 4) move / resize：只改 a:off / a:ext 的属性值
    if (op.move !== undefined || op.resize !== undefined) {
        const xfrm = findXfrm(block, local);
        if (xfrm?.off === undefined || xfrm.ext === undefined) {
            return { ok: false, reason: '这个形状没有 a:xfrm（图片、表格、组合形状的几何在别处），改不了位置' };
        }
        let nextOff = xfrm.off;
        let nextExt = xfrm.ext;
        if (op.move !== undefined) {
            const move = op.move !== null && typeof op.move === 'object' ? op.move : {};
            if (move.x !== undefined && move.y !== undefined) {
                nextOff = setAttr(setAttr(nextOff, 'x', String(cmToEmu(move.x))), 'y', String(cmToEmu(move.y)));
            } else {
                const dx = cmToEmu(move.dx ?? 0);
                const dy = cmToEmu(move.dy ?? 0);
                if (dx === 0 && dy === 0) return { ok: false, reason: 'move 需要 x/y 或非零的 dx/dy' };
                nextOff = setAttr(setAttr(nextOff, 'x', String((xfrm.x ?? 0) + dx)), 'y', String((xfrm.y ?? 0) + dy));
            }
        }
        if (op.resize !== undefined) {
            const resize = op.resize !== null && typeof op.resize === 'object' ? op.resize : {};
            let w = xfrm.w ?? 0;
            let h = xfrm.h ?? 0;
            if (resize.w !== undefined || resize.h !== undefined) {
                w = resize.w === undefined ? w : cmToEmu(resize.w);
                h = resize.h === undefined ? h : cmToEmu(resize.h);
            } else {
                const dw = cmToEmu(resize.dw ?? 0);
                const dh = cmToEmu(resize.dh ?? 0);
                if (dw === 0 && dh === 0) return { ok: false, reason: 'resize 需要 w/h 或非零的 dw/dh' };
                w += dw;
                h += dh;
            }
            nextExt = setAttr(setAttr(nextExt, 'cx', String(Math.max(1, w))), 'cy', String(Math.max(1, h)));
        }
        block = splice(block, xfrm.start, xfrm.end, xfrm.block.replace(xfrm.off, nextOff).replace(xfrm.ext, nextExt));
    }

    if (block === before) return { ok: false, reason: '这一条操作没有产生任何改动' };

    const xfrm = findXfrm(block, local);
    const overflow = [];
    if (xfrm !== undefined && xfrm.off !== undefined && xfrm.ext !== undefined) {
        const outside = xfrm.x < 0 || xfrm.y < 0 || xfrm.x + xfrm.w > size.cx || xfrm.y + xfrm.h > size.cy;
        if (outside) {
            overflow.push(`改完后形状超出页面（x ${emuToCm(xfrm.x)}cm, y ${emuToCm(xfrm.y)}cm, ${emuToCm(xfrm.w)}×${emuToCm(xfrm.h)}cm）`);
        }
    }
    return {
        ok: true,
        xml: xml.slice(0, shape.start) + block + xml.slice(shape.end),
        label: labelOf(op),
        detail: {
            ...(xfrm === undefined ? {} : { x: emuToCm(xfrm.x), y: emuToCm(xfrm.y), w: emuToCm(xfrm.w), h: emuToCm(xfrm.h) }),
            text: normalizeText(block).slice(0, 40),
        },
        warnings: overflow,
    };
}

function splice(source, from, to, replacement) {
    return source.slice(0, from) + replacement + source.slice(to);
}

function labelOf(op) {
    if (op.setText !== undefined) return 'setText';
    if (op.replace !== undefined) return 'replace';
    if (op.style !== undefined) return 'style';
    if (op.move !== undefined && op.resize !== undefined) return 'move+resize';
    if (op.move !== undefined) return 'move';
    if (op.resize !== undefined) return 'resize';
    if (op.fit === true) return 'fit';
    return 'op';
}

function normalizeStylePatch(style) {
    const patch = {};
    if (style.sizePt !== undefined || style.size !== undefined) patch.sizePt = clampNumber(style.sizePt ?? style.size, 1, 400, undefined);
    if (style.bold !== undefined) patch.bold = style.bold === true || style.bold === 'true';
    if (style.italic !== undefined) patch.italic = style.italic === true || style.italic === 'true';
    if (style.underline !== undefined) patch.underline = style.underline === true || style.underline === 'true';
    if (style.color !== undefined) patch.color = toHex(style.color);
    if (style.font !== undefined) patch.font = asString(style.font);
    if (style.spacing !== undefined) patch.spacing = asNumber(style.spacing, 0);
    if (style.align !== undefined) patch.align = asString(style.align);
    if (style.lineSpacing !== undefined) patch.lineSpacing = asNumber(style.lineSpacing, 1.2);
    if (style.text !== undefined) patch.text = asString(style.text);
    for (const key of Object.keys(patch)) {
        if (patch[key] === undefined || patch[key] === '') delete patch[key];
    }
    return patch;
}

/** `replace` 的三种写法都收：`[['a','b']]` / `[{find, replace}]` / `{find, replace}`。 */
function normalizePairs(value) {
    if (Array.isArray(value)) {
        return value.map((item) => (Array.isArray(item)
            ? { find: asString(item[0]), replace: asString(item[1]), all: true }
            : { find: asString(item?.find), replace: asString(item?.replace), all: item?.all !== false }));
    }
    if (value !== null && typeof value === 'object') {
        return [{ find: asString(value.find), replace: asString(value.replace), all: value.all !== false }];
    }
    return [];
}

/** 只在 `a:t` 文本节点里做字面替换 —— 直接改整段 XML 会把属性值里的同名子串也换掉。 */
function replaceInTextNodes(block, find, replace, all) {
    let remaining = all ? Infinity : 1;
    let count = 0;
    const xml = block.replace(/<a:t(\s[^>]*)?>([\s\S]*?)<\/a:t>/g, (whole, attrs, body) => {
        if (remaining <= 0) return whole;
        const text = decodeEntities(body);
        if (!text.includes(find)) return whole;
        const times = all ? text.split(find).length - 1 : 1;
        const used = Math.min(times, remaining);
        remaining -= used;
        count += used;
        const next = all ? text.split(find).join(replace) : text.replace(find, replace);
        return `<a:t xml:space="preserve">${escapeAttr(next)}</a:t>`;
    });
    return { xml, count };
}

/**
 * 按当前字号与文本框宽度估算内容高度（EMU）。
 * 只为 `fit: true` 服务，用的是与生成端同一套 estimateEm，所以两边口径一致。
 */
function estimateTextHeight(txBody, widthEmu) {
    const paragraphs = txBody.match(PARA_RE) ?? [];
    const widthPt = Math.max(20, widthEmu / 12700);
    let totalPt = 0;
    let maxSize = 11;
    for (const para of paragraphs) {
        const tMatch = /<a:t[^>]*>([\s\S]*?)<\/a:t>/.exec(para);
        const text = tMatch === null ? '' : decodeEntities(tMatch[1]);
        const rPr = RPR_RE.exec(para)?.[0];
        const size = Math.max(6, (Number.parseInt(readAttr(rPr ?? '', 'sz') ?? '1800', 10) || 1800) / 100);
        maxSize = Math.max(maxSize, size);
        const lnSpc = /<a:lnSpc>[\s\S]*?<a:spcPct val="(\d+)"\/>/.exec(para);
        const spacing = lnSpc === null ? 1.2 : Math.max(0.1, Number.parseInt(lnSpc[1], 10) / 100000);
        // 与生成端同一套估算：全角字符按 1 em，西文按字面宽度近似
        let em = 0;
        for (const ch of text) em += /[\u1100-\u115F\u2E80-\u303E\u3041-\u33FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFF60\uFFE0-\uFFE6]/.test(ch)
            ? 1
            : (ch === ' ' ? 0.28 : (ch >= 'A' && ch <= 'Z' ? 0.66 : (ch >= '0' && ch <= '9' ? 0.56 : 0.54)));
        const lines = Math.max(1, Math.ceil((em * size) / widthPt));
        totalPt += lines * size * spacing;
    }
    // 上下留一点内边距，乘 1.06 避免正好贴着边
    return Math.max(1, Math.round((totalPt + maxSize * 0.4) * 12700 * 1.06));
}
