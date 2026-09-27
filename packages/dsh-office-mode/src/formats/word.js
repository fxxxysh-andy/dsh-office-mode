/**
 * Word（.docx）格式引擎。
 *
 * 为什么排版直接拼 XML 字符串，而不是先建 DOM 再序列化：
 * OOXML 的 schema 对子元素顺序极敏感（`w:pPr` 里 `w:spacing` 必须在 `w:ind` 之前、
 * `w:tcPr` 里 `w:tcW` 必须在 `w:shd` 之前），顺序错一个 Word 就报「内容有问题」。
 * 手写顺序 + 注释说明「为什么放在这」比「先堆进对象字典、再靠排序函数猜」更可控，
 * 出问题时也能一眼定位到具体那一行。
 *
 * 为什么 stats/outline/warnings 由「渲染完再解析自己产出的 document.xml」得到：
 * create 与 read 因此走完全同一条代码路径，两边数字不可能对不上；
 * 代价只是 render 时多解析一次几 KB 的 XML。
 *
 * @module dsh-office-mode/formats/word
 */
import { unzip, zip } from '../engine/zip.js';
import { attr, children, descendant, descendants, escapeXml, parseXml, textOf } from '../engine/xml.js';
import {
    asArray, asNumber, asString, classifyValue, displayPath, EMU_PER_CM, EMU_PER_PT, ensureExtension,
    estimateLines, normalizeColumns, normalizeRows, ptToHalfPoints, ptToTwips, slugify,
    TWIPS_PER_CM,
} from '../engine/kit.js';
import { DEFAULT_THEME_ID, THEMES, resolveTheme, shadeOf, toHex } from '../engine/theme.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** 数学命名空间：公式（OMML）的元素全部挂在它下面，document.xml 根元素必须声明。 */
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CP_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC_NS = 'http://purl.org/dc/elements/1.1/';
const DCTERMS_NS = 'http://purl.org/dc/terms/';
const XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance';
const EP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

const REL_TYPE = {
    officeDocument: `${R_NS}/officeDocument`,
    core: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    app: `${R_NS}/extended-properties`,
    styles: `${R_NS}/styles`,
    numbering: `${R_NS}/numbering`,
    settings: `${R_NS}/settings`,
    footer: `${R_NS}/footer`,
    header: `${R_NS}/header`,
    image: `${R_NS}/image`,
};

const CT_PART = {
    document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
    settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
    header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    core: 'application/vnd.openxmlformats-package.core-properties+xml',
    app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
};

/** A4 / A3 / Letter / Legal / B5（横向时 w/h 互换）。单位 twips。 */
const PAGE_SIZES = {
    a4: { w: 11906, h: 16838, label: 'A4' },
    a3: { w: 16838, h: 23811, label: 'A3' },
    letter: { w: 12240, h: 15840, label: 'Letter' },
    legal: { w: 12240, h: 20160, label: 'Legal' },
    b5: { w: 9979, h: 14173, label: 'B5' },
};

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** 可以在 pStyle / setStyle 里出现的样式 id 白名单，防止 edit 写出 Word 不认识的样式。 */
const PARAGRAPH_STYLES = [
    'Normal', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3',
    'Quote', 'ListParagraph', 'Code', 'Caption', 'TableText', 'Formula',
];

const BODY_SIZE_PT = 10.5;

/**
 * 段落样式的「真实排版参数」。
 * 这里是 styles.xml 的生成源，也是 analyze() 估算行高与页数的唯一依据：
 * 同一份数字喂给两边，报告里的 pagesEstimate 才不是拍脑袋写死的。
 */
const STYLE_INFO = {
    Normal: { kind: 'body', size: BODY_SIZE_PT, line: 360, before: 0, after: 120, align: 'both' },
    Title: { kind: 'title', size: 26, line: 280, before: 360, after: 120, align: 'left' },
    Subtitle: { kind: 'subtitle', size: 13, line: 300, before: 0, after: 360, align: 'left' },
    Heading1: { kind: 'heading', level: 1, size: 18, line: 280, before: 400, after: 160, keepNext: true },
    Heading2: { kind: 'heading', level: 2, size: 15, line: 280, before: 280, after: 120, keepNext: true },
    Heading3: { kind: 'heading', level: 3, size: 12.5, line: 280, before: 200, after: 100, keepNext: true },
    Quote: { kind: 'quote', size: BODY_SIZE_PT, line: 320, before: 120, after: 200, indentLeft: 480, indentRight: 240 },
    ListParagraph: { kind: 'list', size: BODY_SIZE_PT, line: 300, before: 0, after: 60, indentLeft: 420, hanging: 240 },
    Code: { kind: 'code', size: 9, line: 240, before: 120, after: 120, indentLeft: 240, indentRight: 240 },
    Caption: { kind: 'caption', size: 8.5, line: 240, before: 60, after: 200, align: 'center' },
    TableText: { kind: 'tableText', size: 10, line: 240, before: 20, after: 20 },
    // 公式段落单独一个样式：默认 11pt、居中、留白比正文略大（公式比汉字高，紧了会撞行）
    Formula: { kind: 'formula', size: 11, line: 300, before: 120, after: 120, align: 'center' },
};

/**
 * 目录条目的排版参数。目录用 Word 自己的内置样式 id（TOC1–TOC3）：
 * 一是真实 Word 生成的目录就是这么写的，二是 `read()` 靠 `w:pStyle` 把这些段落
 * 认出来、不再当成正文段落统计。缩进逐级 +420 twips（约 0.74cm）。
 */
const TOC_INFO = [
    { level: 1, size: 11, indentTw: 0, bold: true },
    { level: 2, size: 10.5, indentTw: 420, bold: false },
    { level: 3, size: 10, indentTw: 840, bold: false },
];
const TOC_STYLE_LEVEL = { TOC1: 1, TOC2: 2, TOC3: 3 };
/** 标题上的书签名前缀：目录里的超链接靠它跳转（w:hyperlink w:anchor）。 */
const TOC_ANCHOR_PREFIX = '_Toc_DSH_';

// ---------------------------------------------------------------- 小工具

const be32 = (bytes, at) => (((bytes[at] << 24) | (bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]) >>> 0);
const be16 = (bytes, at) => ((bytes[at] << 8) | bytes[at + 1]);

/** 实体反转义：xml.js 只导出 escapeXml，替换操作需要把 `&amp;` 还原成 `&` 才能比对文本。 */
const ENTITY_MAP = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
function decodeXmlText(text) {
    if (!text.includes('&')) return text;
    return text.replace(/&(#[xX]?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body) => {
        if (body[0] === '#') {
            const hex = body[1] === 'x' || body[1] === 'X';
            const code = Number.parseInt(hex ? body.slice(2) : body.slice(1), hex ? 16 : 10);
            return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        return ENTITY_MAP[body] ?? ENTITY_MAP[body.toLowerCase()] ?? whole;
    });
}

/** 只认 PNG/JPEG 的魔数，顺便读原始像素尺寸 —— 图片大小要按真实比例算，不能靠文件名猜。 */
function sniffImage(bytes) {
    if (bytes.length > 24 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
        return { kind: 'png', ext: 'png', width: be32(bytes, 16), height: be32(bytes, 20) };
    }
    if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
        let at = 2;
        while (at + 9 < bytes.length) {
            if (bytes[at] !== 0xff) { at += 1; continue; }
            const marker = bytes[at + 1];
            // SOF0..SOF15 里除 DHT/JPG/DAC 之外都带尺寸
            if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
                return { kind: 'jpeg', ext: 'jpeg', width: be16(bytes, at + 7), height: be16(bytes, at + 5) };
            }
            if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { at += 2; continue; }
            const size = be16(bytes, at + 2);
            if (size < 2) break;
            at += 2 + size;
        }
        return { kind: 'jpeg', ext: 'jpeg', width: 0, height: 0 };
    }
    return null;
}

function rFontsXml(font, ctx) {
    const cn = typeof font === 'string' ? font : asString(font?.cn, ctx.fonts.cn);
    const en = typeof font === 'string' ? font : asString(font?.en, ctx.fonts.en);
    return `<w:rFonts w:ascii="${escapeXml(en)}" w:hAnsi="${escapeXml(en)}" w:eastAsia="${escapeXml(cn)}" w:cs="${escapeXml(en)}"/>`;
}

function borderXml(name, { sz, color, space = 0, val = 'single' }) {
    // 接受的写法统一成裸边名（bottom / w:bottom 都行）：
    // 早期这里直接拼 `<w:${name}>`，调用方传 'w:bottom' 就会写出 `<w:w:bottom>`，
    // 这种名字连 XML 解析器都过不去（'：' 不能出现在名字里），Word 必然报损坏。
    const side = asString(name).replace(/^w:/, '');
    return `<w:${side} w:val="${val}" w:sz="${sz}" w:space="${space}" w:color="${color}"/>`;
}

function alignOf(value, fallback) {
    const text = asString(value).toLowerCase();
    const map = { left: 'left', start: 'left', center: 'center', centre: 'center', right: 'right', end: 'right', justify: 'both', both: 'both', distribute: 'distribute' };
    return map[text] ?? fallback;
}

/** 规范化 `opts.indent`：数字按厘米，对象按厘米逐边。 */
function indentTwips(indent, base) {
    if (indent === undefined || indent === null) return base;
    if (typeof indent === 'number') return { left: Math.round(indent * TWIPS_PER_CM) };
    const out = {};
    for (const key of ['left', 'right', 'firstLine', 'hanging']) {
        if (indent[key] !== undefined) out[key] = Math.round(asNumber(indent[key], 0) * TWIPS_PER_CM);
    }
    return out;
}

function indentAttr(indent) {
    if (!indent) return '';
    const bits = [];
    for (const [key, value] of Object.entries(indent)) {
        if (Number.isFinite(value) && value !== 0) bits.push(`${key === 'firstLine' ? 'w:firstLine' : `w:${key}`}="${Math.round(value)}"`);
    }
    return bits.join(' ');
}

// ---------------------------------------------------------------- 主题 / 页面

/**
 * 深色主题落到白纸上必须换墨：`tech` 的 text 是浅灰（E5E7EB），
 * 直接拿去写正文就是一页看不见的字。这里只替换文字/底色系，主色保持，
 * 于是「科技深色」在 Word 里表现为「青蓝主色的浅色文档」。
 */
function paperInk(theme) {
    const colors = theme.colors;
    if (!theme.dark) return { ...colors, table: { ...theme.table } };
    return {
        ...colors,
        text: '1F2430',
        muted: '5A6472',
        subtle: '8A94A3',
        bg: 'FFFFFF',
        surface: 'F1F5F9',
        border: 'CBD5E1',
        onPrimary: 'FFFFFF',
        table: {
            headerFill: colors.primary,
            headerText: 'FFFFFF',
            zebraFill: 'F1F5F9',
            border: 'CBD5E1',
            accentFill: 'E0F2FE',
        },
    };
}

function pickTheme(id, env) {
    const picked = typeof env?.theme === 'function' ? env.theme(id) : resolveTheme(id);
    const safe = picked && picked.theme ? picked : resolveTheme(id);
    return safe;
}

function resolvePage(spec = {}) {
    const raw = spec.page;
    let key = 'a4';
    let size = PAGE_SIZES.a4;
    if (typeof raw === 'string' && PAGE_SIZES[raw.trim().toLowerCase()]) {
        key = raw.trim().toLowerCase();
        size = PAGE_SIZES[key];
    } else if (raw && typeof raw === 'object' && Number.isFinite(asNumber(raw.widthCm, Number.NaN))) {
        size = {
            w: Math.round(asNumber(raw.widthCm) * TWIPS_PER_CM),
            h: Math.round(asNumber(raw.heightCm, 29.7) * TWIPS_PER_CM),
            label: '自定义',
        };
        key = 'custom';
    }
    const landscape = /land|横向|横排/i.test(asString(spec.orientation, 'portrait'));
    const width = landscape ? size.h : size.w;
    const height = landscape ? size.w : size.h;

    const m = spec.margins;
    let margins = { top: 2.4, bottom: 2.4, left: 2.6, right: 2.6 };
    if (typeof m === 'number') margins = { top: m, bottom: m, left: m, right: m };
    else if (m && typeof m === 'object') {
        margins = {
            top: asNumber(m.top, margins.top),
            bottom: asNumber(m.bottom, margins.bottom),
            left: asNumber(m.left, margins.left),
            right: asNumber(m.right, margins.right),
        };
    }
    const tw = (value) => Math.round(value * TWIPS_PER_CM);
    return {
        label: `${size.label} ${landscape ? '横向' : '纵向'}`,
        landscape,
        widthTw: width,
        heightTw: height,
        margin: { top: tw(margins.top), bottom: tw(margins.bottom), left: tw(margins.left), right: tw(margins.right) },
        headerTw: Math.round(Math.min(1.5, margins.top * 0.6) * TWIPS_PER_CM),
        footerTw: Math.round(Math.min(1.4, margins.bottom * 0.6) * TWIPS_PER_CM),
    };
}

/** 从已存在的 sectPr 反推页面参数：read()/edit() 必须尊重原文件的版心，不能假设 A4。 */
function pageFromXml(body) {
    const sectPr = children(body, 'w:sectPr').pop();
    if (!sectPr) return { ...resolvePage({}), fromFile: false };
    const pgSz = children(sectPr, 'w:pgSz')[0];
    const pgMar = children(sectPr, 'w:pgMar')[0];
    const widthTw = asNumber(attr(pgSz, 'w:w'), 11906);
    const heightTw = asNumber(attr(pgSz, 'w:h'), 16838);
    const margin = {
        top: asNumber(attr(pgMar, 'w:top'), 1361),
        bottom: asNumber(attr(pgMar, 'w:bottom'), 1361),
        left: asNumber(attr(pgMar, 'w:left'), 1474),
        right: asNumber(attr(pgMar, 'w:right'), 1474),
    };
    const known = Object.values(PAGE_SIZES).find(
        (size) => (size.w === widthTw && size.h === heightTw) || (size.w === heightTw && size.h === widthTw),
    );
    return {
        label: known ? `${known.label} ${heightTw >= widthTw ? '纵向' : '横向'}` : `自定义 ${(widthTw / TWIPS_PER_CM).toFixed(1)}×${(heightTw / TWIPS_PER_CM).toFixed(1)}cm`,
        landscape: widthTw > heightTw,
        widthTw,
        heightTw,
        margin,
        headerTw: asNumber(attr(pgMar, 'w:header'), 851),
        footerTw: asNumber(attr(pgMar, 'w:footer'), 794),
        fromFile: true,
    };
}

/** 用 Heading1 的颜色反查主题 id：styles.xml 是包内唯一记录了主色的地方。 */
function detectTheme(stylesXml) {
    if (typeof stylesXml !== 'string') return DEFAULT_THEME_ID;
    for (const id of ['Heading1', 'Title']) {
        const match = new RegExp(`w:styleId="${id}"[\\s\\S]{0,600}?<w:color w:val="([0-9A-Fa-f]{6})"`).exec(stylesXml);
        if (!match) continue;
        const hex = match[1].toUpperCase();
        const hit = THEMES.find((theme) => theme.colors.primary.toUpperCase() === hex);
        if (hit) return hit.id;
    }
    return DEFAULT_THEME_ID;
}

// ---------------------------------------------------------------- 样式 / 部件

function spacingOf(info, opts = {}) {
    const before = opts.spaceBefore !== undefined ? ptToTwips(asNumber(opts.spaceBefore, 0)) : info.before;
    const after = opts.spaceAfter !== undefined ? ptToTwips(asNumber(opts.spaceAfter, 0)) : info.after;
    let line = info.line;
    let lineRule = 'auto';
    if (opts.lineSpacing !== undefined) {
        const value = asNumber(opts.lineSpacing, 0);
        // 小于等于 3 当成倍数（1.5 倍行距），大于 3 当成磅值
        if (value > 0 && value <= 3) { line = Math.round(value * 240); lineRule = 'auto'; }
        else if (value > 3) { line = ptToTwips(value); lineRule = 'atLeast'; }
    }
    return { before, after, line, lineRule };
}

function pPrXml(o = {}) {
    const bits = [];
    if (o.styleId) bits.push(`<w:pStyle w:val="${o.styleId}"/>`);
    if (o.keepNext) bits.push('<w:keepNext/>');
    if (o.keepLines) bits.push('<w:keepLines/>');
    if (o.numId) bits.push(`<w:numPr><w:ilvl w:val="${o.ilvl ?? 0}"/><w:numId w:val="${o.numId}"/></w:numPr>`);
    if (o.borders) bits.push(`<w:pBdr>${o.borders}</w:pBdr>`);
    if (o.fill) bits.push(`<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>`);
    // w:tabs 在 CT_PPrBase 里排在 w:shd 之后、w:spacing 之前；带编号的公式就靠
    // 「居中制表位 + 右制表位」把公式和右侧编号摆到同一行上
    if (o.tabs) bits.push(`<w:tabs>${o.tabs}</w:tabs>`);
    if (o.spacing) bits.push(o.spacing);
    if (o.indent) bits.push(`<w:ind ${o.indent}/>`);
    if (o.contextualSpacing) bits.push('<w:contextualSpacing/>');
    if (o.align) bits.push(`<w:jc w:val="${o.align}"/>`);
    if (o.outlineLvl !== undefined) bits.push(`<w:outlineLvl w:val="${o.outlineLvl}"/>`);
    if (!bits.length) return '';
    return `<w:pPr>${bits.join('')}</w:pPr>`;
}

function rPrXml(o = {}, ctx) {
    const bits = [];
    if (o.font) bits.push(rFontsXml(o.font, ctx));
    if (o.bold) bits.push('<w:b/><w:bCs/>');
    if (o.italic) bits.push('<w:i/><w:iCs/>');
    if (o.color) bits.push(`<w:color w:val="${toHex(o.color, ctx.ink.text)}"/>`);
    if (o.size) bits.push(`<w:sz w:val="${ptToHalfPoints(asNumber(o.size, BODY_SIZE_PT))}"/><w:szCs w:val="${ptToHalfPoints(asNumber(o.size, BODY_SIZE_PT))}"/>`);
    if (o.underline) bits.push('<w:u w:val="single"/>');
    return bits.length ? `<w:rPr>${bits.join('')}</w:rPr>` : '';
}

/** 文本里的换行在 OOXML 里是 `<w:br/>`，只能落在 run 内部。 */
function runTextXml(text) {
    return String(text ?? '')
        .split('\n')
        .map((piece, index) => `${index > 0 ? '<w:br/>' : ''}<w:t xml:space="preserve">${escapeXml(piece)}</w:t>`)
        .join('');
}

function runXml(text, o, ctx) {
    return `<w:r>${rPrXml(o, ctx)}${runTextXml(text)}</w:r>`;
}

/** 段落内容既接受纯字符串，也接受 run 数组 `[{text,bold,...}]`。 */
function normalizeRuns(value) {
    if (Array.isArray(value)) {
        return value.map((item) => (typeof item === 'string' ? { text: item } : { ...item, text: asString(item?.text ?? item?.value) }));
    }
    return [{ text: asString(value) }];
}

/** 行内公式：run 里的 `{math:'x^2'}` 转成 <m:oMath>，与普通 w:r 平级排在同一段里。 */
const runMathXml = (run, opts, ctx) => mathToOmml(asString(run.math), ctx, {
    size: asNumber(run.size, asNumber(opts?.size, 0)) || 0,
    color: asString(run.color, '') || undefined,
});

function runsXml(runs, opts, ctx) {
    return runs.map((run) => (run.math !== undefined
        ? runMathXml(run, opts, ctx)
        : runXml(run.text, { ...opts, ...run }, ctx))).join('');
}

/** 段落纯文本（用于报表，不参与渲染细节）。 */
const runsText = (runs) => runs.map((run) => asString(run.text)).join('');

function styleXml({ id, name, type = 'paragraph', def = false, basedOn = 'Normal', next = 'Normal', uiPriority, qFormat, pPr = '', rPr = '', tblPr = '' }) {
    const bits = [`<w:style w:type="${type}"${def ? ' w:default="1"' : ''} w:styleId="${escapeXml(id)}">`,
        `<w:name w:val="${escapeXml(name)}"/>`];
    if (basedOn && id !== 'Normal') bits.push(`<w:basedOn w:val="${basedOn}"/>`);
    if (next && type === 'paragraph') bits.push(`<w:next w:val="${escapeXml(next)}"/>`);
    if (uiPriority !== undefined) bits.push(`<w:uiPriority w:val="${uiPriority}"/>`);
    if (qFormat) bits.push('<w:qFormat/>');
    bits.push(pPr, rPr, tblPr, '</w:style>');
    return bits.join('');
}

function styleSpacingXml(info, opts = {}) {
    const spacing = spacingOf(info, opts);
    return `<w:spacing w:before="${spacing.before}" w:after="${spacing.after}" w:line="${spacing.line}" w:lineRule="${spacing.lineRule}"/>`;
}

function stylesXml(theme, ink, fonts) {
    const en = fonts.en;
    const cn = fonts.cn;
    const titleEn = fonts.titleEn;
    const titleCn = fonts.titleCn;
    const info = STYLE_INFO;
    const rpr = (size, color, extra = '') => `<w:rPr>${extra}<w:color w:val="${color}"/><w:sz w:val="${ptToHalfPoints(size)}"/><w:szCs w:val="${ptToHalfPoints(size)}"/></w:rPr>`;

    const docDefaults = '<w:docDefaults>'
        + `<w:rPrDefault><w:rPr>${rFontsXml({ cn, en }, { fonts })}<w:kern w:val="2"/><w:sz w:val="${ptToHalfPoints(BODY_SIZE_PT)}"/><w:szCs w:val="${ptToHalfPoints(BODY_SIZE_PT)}"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>`
        + `<w:pPrDefault>${pPrXml({ spacing: styleSpacingXml(info.Normal), align: 'both' })}</w:pPrDefault>`
        + '</w:docDefaults>';

    // 样式里的 pPr 一律走 pPrXml：CT_PPrBase 的子元素顺序是 schema 强制的
    // （pBdr/shd 在 spacing 之前、contextualSpacing 在 ind 之后、jc 在最后），
    // 手写字符串迟早会在某个样式上写反，Word 就直接报「内容有问题」。
    const list = [
        styleXml({
            id: 'Normal', name: 'Normal', def: true, next: 'Normal', uiPriority: 0,
            pPr: pPrXml({ spacing: styleSpacingXml(info.Normal), align: 'both' }),
            rPr: `<w:rPr>${rFontsXml({ cn, en }, { fonts })}<w:color w:val="${ink.text}"/><w:sz w:val="${ptToHalfPoints(BODY_SIZE_PT)}"/><w:szCs w:val="${ptToHalfPoints(BODY_SIZE_PT)}"/></w:rPr>`,
        }),
        styleXml({
            id: 'Title', name: 'Title', next: 'Subtitle', uiPriority: 10, qFormat: true,
            pPr: pPrXml({
                keepNext: true,
                borders: borderXml('bottom', { sz: 12, color: ink.primary, space: 8 }),
                spacing: styleSpacingXml(info.Title),
                align: 'left',
            }),
            // 字距 w:spacing 在 CT_RPr 里排在 w:color 之后，写前面同样是顺序错误
            rPr: `<w:rPr>${rFontsXml({ cn: titleCn, en: titleEn }, { fonts })}<w:b/><w:bCs/><w:color w:val="${ink.primary}"/><w:spacing w:val="20"/><w:sz w:val="${ptToHalfPoints(info.Title.size)}"/><w:szCs w:val="${ptToHalfPoints(info.Title.size)}"/></w:rPr>`,
        }),
        styleXml({
            id: 'Subtitle', name: 'Subtitle', next: 'Normal', uiPriority: 11, qFormat: true,
            pPr: pPrXml({ spacing: styleSpacingXml(info.Subtitle), align: 'left' }),
            rPr: rpr(info.Subtitle.size, ink.muted, rFontsXml({ cn, en }, { fonts })),
        }),
    ];

    // 标题层级靠 before/after 递进留白：H1 400/160、H2 280/120、H3 200/100（twips）
    for (const level of [1, 2, 3]) {
        const key = `Heading${level}`;
        const style = info[key];
        const color = level === 1 ? ink.primary : (level === 2 ? ink.primaryDark : ink.secondary);
        list.push(styleXml({
            id: key,
            name: `heading ${level}`,
            next: 'Normal',
            uiPriority: 9,
            qFormat: true,
            pPr: pPrXml({
                keepNext: true,
                keepLines: true,
                borders: level === 1 ? borderXml('bottom', { sz: 6, color: ink.border, space: 6 }) : '',
                spacing: styleSpacingXml(style),
                outlineLvl: level - 1,
            }),
            rPr: `<w:rPr>${rFontsXml({ cn: titleCn, en: titleEn }, { fonts })}<w:b/><w:bCs/><w:color w:val="${color}"/><w:sz w:val="${ptToHalfPoints(style.size)}"/><w:szCs w:val="${ptToHalfPoints(style.size)}"/></w:rPr>`,
        }));
    }

    // 目录条目样式（TOC1–TOC3）：只放字体、间距与逐级缩进；右对齐点线制表位
    // 由渲染时的直接 pPr 给出 —— 那里才知道版心宽度，样式里写死会随页面设置漂移。
    for (const toc of TOC_INFO) {
        list.push(styleXml({
            id: `TOC${toc.level}`,
            name: `toc ${toc.level}`,
            next: 'Normal',
            uiPriority: 39,
            qFormat: true,
            pPr: pPrXml({
                spacing: '<w:spacing w:before="0" w:after="40" w:line="300" w:lineRule="auto"/>',
                indent: `w:left="${toc.indentTw}" w:right="0"`,
                align: 'left',
            }),
            rPr: `<w:rPr>${rFontsXml({ cn, en }, { fonts })}${toc.bold ? '<w:b/><w:bCs/>' : ''}<w:color w:val="${toc.level === 1 ? ink.text : ink.muted}"/><w:sz w:val="${ptToHalfPoints(toc.size)}"/><w:szCs w:val="${ptToHalfPoints(toc.size)}"/></w:rPr>`,
        }));
    }

    list.push(styleXml({
        id: 'Quote', name: 'Quote', next: 'Normal', uiPriority: 29, qFormat: true,
        pPr: pPrXml({
            borders: borderXml('left', { sz: 18, color: ink.accent, space: 8 }),
            // 素色主题的 surface 是白色，shadeOf 会把它变成「不填充」：
            // 引用块只留左侧竖线，不再给整段加底色
            fill: shadeOf(ink.surface),
            spacing: styleSpacingXml(info.Quote),
            indent: `w:left="${info.Quote.indentLeft}" w:right="${info.Quote.indentRight}"`,
            align: 'left',
        }),
        rPr: rpr(info.Quote.size, ink.muted),
    }));
    list.push(styleXml({
        id: 'ListParagraph', name: 'List Paragraph', next: 'ListParagraph', uiPriority: 34, qFormat: true,
        // 刻意不在样式里写 w:ind：段落级 ind 会覆盖 numbering 的缩进，
        // 所以缩进按层级逐段显式给出（见 renderBlock 的 list 分支）。
        pPr: pPrXml({ contextualSpacing: true, spacing: styleSpacingXml(info.ListParagraph), align: 'both' }),
        rPr: rpr(info.ListParagraph.size, ink.text),
    }));
    list.push(styleXml({
        id: 'Code', name: 'Code', next: 'Normal', uiPriority: 35,
        pPr: pPrXml({
            keepLines: true,
            borders: ['top', 'left', 'bottom', 'right'].map((side) => borderXml(side, { sz: 4, color: ink.border, space: 4 })).join(''),
            fill: shadeOf(ink.surface),
            spacing: styleSpacingXml(info.Code),
            indent: `w:left="${info.Code.indentLeft}" w:right="${info.Code.indentRight}"`,
            align: 'left',
        }),
        rPr: `<w:rPr>${rFontsXml({ cn, en: 'Consolas' }, { fonts })}<w:color w:val="${ink.text}"/><w:sz w:val="${ptToHalfPoints(info.Code.size)}"/><w:szCs w:val="${ptToHalfPoints(info.Code.size)}"/></w:rPr>`,
    }));
    list.push(styleXml({
        id: 'Caption', name: 'caption', next: 'Normal', uiPriority: 36, qFormat: true,
        pPr: pPrXml({ spacing: styleSpacingXml(info.Caption), align: 'center' }),
        rPr: rpr(info.Caption.size, ink.muted),
    }));
    list.push(styleXml({
        id: 'TableText', name: 'Table Text', next: 'TableText', uiPriority: 40,
        pPr: pPrXml({ spacing: styleSpacingXml(info.TableText), align: 'left' }),
        rPr: rpr(info.TableText.size, ink.text),
    }));
    list.push(styleXml({
        id: 'Formula', name: 'Formula', next: 'Normal', uiPriority: 38,
        // keepLines 让「公式 + 编号」这一行不会被拆到两页；对齐仍由每段的 w:jc 决定
        pPr: pPrXml({ keepLines: true, spacing: styleSpacingXml(info.Formula), align: 'center' }),
        rPr: rpr(info.Formula.size, ink.text),
    }));
    const tightSpacing = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
    list.push(styleXml({
        id: 'Header', name: 'header', basedOn: 'Normal', next: 'Header',
        pPr: pPrXml({
            borders: borderXml('bottom', { sz: 4, color: ink.border, space: 2 }),
            spacing: tightSpacing,
            align: 'right',
        }),
        rPr: rpr(9, ink.muted),
    }));
    list.push(styleXml({
        id: 'Footer', name: 'footer', basedOn: 'Normal', next: 'Footer',
        pPr: pPrXml({ spacing: tightSpacing, align: 'center' }),
        rPr: rpr(9, ink.muted),
    }));

    // TableGrid 是表格样式（w:type="table"），不是段落样式 —— 段落引用它会变成非法引用。
    list.push(styleXml({
        id: 'TableGrid', name: 'Table Grid', type: 'table', basedOn: '', next: '',
        tblPr: '<w:tblPr>'
            + `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => borderXml(side, { sz: 4, color: ink.border, space: 0 })).join('')}</w:tblBorders>`
            + '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
            + '</w:tblPr>',
    }));

    return `${XML_HEAD}<w:styles xmlns:w="${W_NS}" xmlns:r="${R_NS}">${docDefaults}${list.join('')}</w:styles>`;
}

function numberingXml(fonts) {
    const bullets = ['•', '○', '▪'];
    const bodyCn = fonts.cn;
    const bodyEn = fonts.en;
    const bulletFonts = `<w:rFonts w:ascii="${escapeXml(bodyEn)}" w:hAnsi="${escapeXml(bodyEn)}" w:eastAsia="${escapeXml(bodyCn)}" w:hint="eastAsia"/>`;
    const parts = [];
    for (let level = 0; level < 3; level += 1) {
        const left = 420 + level * 420;
        parts.push(`<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>`
            + `<w:lvlText w:val="${escapeXml(bullets[level])}"/><w:lvlJc w:val="left"/>`
            + `<w:pPr><w:ind w:left="${left}" w:hanging="240"/></w:pPr>`
            + `<w:rPr>${bulletFonts}</w:rPr></w:lvl>`);
    }
    const orderedParts = [];
    for (let level = 0; level < 3; level += 1) {
        const left = 420 + level * 420;
        const text = ['%1.', '%1.%2.', '%1.%2.%3.'][level];
        orderedParts.push(`<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>`
            + `<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>`
            + `<w:pPr><w:ind w:left="${left}" w:hanging="240"/></w:pPr>`
            + `<w:rPr>${bulletFonts}</w:rPr></w:lvl>`);
    }
    return `${XML_HEAD}<w:numbering xmlns:w="${W_NS}">`
        + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${parts.join('')}</w:abstractNum>`
        + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${orderedParts.join('')}</w:abstractNum>`
        + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        + '</w:numbering>';
}

function settingsXml() {
    // 这里刻意**不写** <w:updateFields w:val="true"/>。
    //
    // 只要文档带「打开时更新域」这类标记，Word 打开时就会弹
    // 「该文档包含的域可能引用了其他文件。是否更新该文档中的这些域？」。
    // 真实 Word（cscript + COM，DisplayAlerts 全开）实测：settings 里的
    // w:updateFields 与目录域 begin fldChar 上的 w:dirty 各自都能单独触发该弹窗，
    // 去掉任意一处仍然弹，两处都去掉才不弹。触发条件是「打开时更新域」这个意图本身，
    // 与文档里放的是 PAGE 域还是 TOC 域无关。
    //
    // 代价是 Word 不再自动刷新目录页码，所以目录改成把条目与页码**预渲染**进域的
    // 缓存结果（见 renderToc）：打开即见完整目录，按 F9 仍可让 Word 重算成精确值。
    //
    // 元素顺序受 CT_Settings 约束：defaultTabStop 在 compat 之前。
    return `${XML_HEAD}<w:settings xmlns:w="${W_NS}">`
        + '<w:zoom w:percent="100"/>'
        + '<w:defaultTabStop w:val="420"/>'
        + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
        + '</w:settings>';
}

function coreXml({ title, author, subject, created }) {
    const stamp = created;
    return `${XML_HEAD}<cp:coreProperties xmlns:cp="${CP_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:xsi="${XSI_NS}">`
        + `<dc:title>${escapeXml(title)}</dc:title>`
        + `<dc:subject>${escapeXml(subject)}</dc:subject>`
        + `<dc:creator>${escapeXml(author)}</dc:creator>`
        + `<cp:lastModifiedBy>${escapeXml(author)}</cp:lastModifiedBy>`
        + '<cp:revision>1</cp:revision>'
        + `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>`
        + `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>`
        + '</cp:coreProperties>';
}

function appXml({ title }) {
    return `${XML_HEAD}<Properties xmlns="${EP_NS}" xmlns:vt="${VT_NS}">`
        + '<Application>DSH Office Mode</Application>'
        + '<DocSecurity>0</DocSecurity>'
        + '<ScaleCrop>false</ScaleCrop>'
        + '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>标题</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>'
        + `<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>${escapeXml(title)}</vt:lpstr></vt:vector></TitlesOfParts>`
        + '<Company></Company>'
        + '<LinksUpToDate>false</LinksUpToDate>'
        + '<SharedDoc>false</SharedDoc>'
        + '<HyperlinksChanged>false</HyperlinksChanged>'
        + '<AppVersion>16.0000</AppVersion>'
        + '</Properties>';
}

function contentTypesXml({ hasFooter, hasHeader, media }) {
    const overrides = [
        ['/word/document.xml', CT_PART.document],
        ['/word/styles.xml', CT_PART.styles],
        ['/word/numbering.xml', CT_PART.numbering],
        ['/word/settings.xml', CT_PART.settings],
        ['/docProps/core.xml', CT_PART.core],
        ['/docProps/app.xml', CT_PART.app],
    ];
    if (hasFooter) overrides.push(['/word/footer1.xml', CT_PART.footer]);
    if (hasHeader) overrides.push(['/word/header1.xml', CT_PART.header]);
    const defaults = [['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']];
    const exts = new Set(media.map((item) => item.ext));
    if (exts.has('png')) defaults.push(['png', 'image/png']);
    if (exts.has('jpeg')) defaults.push(['jpeg', 'image/jpeg']);
    return `${XML_HEAD}<Types xmlns="${CT_NS}">`
        + defaults.map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('')
        + overrides.map(([name, type]) => `<Override PartName="${name}" ContentType="${type}"/>`).join('')
        + '</Types>';
}

function rootRelsXml() {
    return `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}">`
        + `<Relationship Id="rId1" Type="${REL_TYPE.officeDocument}" Target="word/document.xml"/>`
        + `<Relationship Id="rId2" Type="${REL_TYPE.core}" Target="docProps/core.xml"/>`
        + `<Relationship Id="rId3" Type="${REL_TYPE.app}" Target="docProps/app.xml"/>`
        + '</Relationships>';
}

// ---------------------------------------------------------------- OMML 公式

/**
 * Word 不认识 LaTeX：公式必须以 OMML（ECMA-376 Part 1 §22.1，math 命名空间）
 * 写进 document.xml。这里是一个手写的递归下降解析器 + OMML 序列化器，只覆盖
 * 办公文档真正常用的 LaTeX 子集，不引任何第三方依赖。
 *
 * 两条贯穿始终的原则：
 *   1. OMML 的子元素顺序被 schema 钉死（m:f 是 num→den、m:rad 是 deg→e、
 *      m:nary 是 sub→sup→e、m:r 是 rPr→w:rPr→m:t），所以每个结构都由一个函数
 *      整段产出，不把子元素零散拼在别处 —— 顺序错一个 Word 就报「内容有问题」；
 *   2. 解析不了的东西一律退化成纯文本 + 一条 warning，绝不抛异常：
 *      一段写坏的公式不该让整份文档生成失败。
 */

/** 希腊字母（小写 + 大写）。OMML 里希腊字母是正体，所以走 m:sty="p"。 */
const MATH_GREEK = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε',
    zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
    lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ',
    sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ',
    chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
    Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

/** 运算符 / 关系符 / 集合符 / 箭头等。 */
const MATH_SYMBOLS = {
    times: '×', cdot: '·', div: '÷', pm: '±', mp: '∓', ast: '∗', star: '⋆',
    le: '≤', leq: '≤', ge: '≥', geq: '≥', neq: '≠', ne: '≠', equiv: '≡',
    approx: '≈', sim: '∼', simeq: '≃', propto: '∝', ll: '≪', gg: '≫',
    infty: '∞', partial: '∂', nabla: '∇', square: '□', triangle: '△',
    in: '∈', notin: '∉', ni: '∋', subset: '⊂', subseteq: '⊆', supset: '⊃',
    supseteq: '⊇', cup: '∪', cap: '∩', setminus: '∖', emptyset: '∅', varnothing: '∅',
    to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐',
    leftrightarrow: '↔', Leftrightarrow: '⇔', mapsto: '↦', implies: '⟹', iff: '⟺',
    forall: '∀', exists: '∃', nexists: '∄', angle: '∠', degree: '°', circ: '∘',
    cdots: '⋯', ldots: '…', dots: '…', vdots: '⋮', ddots: '⋱', prime: '′',
    aleph: 'ℵ', hbar: 'ℏ', ell: 'ℓ', Re: 'ℜ', Im: 'ℑ', wp: '℘',
    oplus: '⊕', otimes: '⊗', perp: '⊥', parallel: '∥', therefore: '∴', because: '∵',
};

/** 反斜杠 + 单个非字母字符：LaTeX 的空格、转义与换行。 */
const MATH_ESCAPE_CHARS = {
    ',': '\u2009', // 细空格
    ';': '\u2005', // 中空格
    ':': '\u2005',
    '!': '',       // 负空格：OMML 没有负间距，直接吞掉
    ' ': ' ',
    '{': '{', '}': '}', '%': '%', '$': '$', '#': '#', '&': '&', '_': '_', '|': '‖',
};
/** 反斜杠 + 单词：显式空格命令。 */
const MATH_SPACE_WORDS = { quad: '\u2003', qquad: '\u2003\u2003', thinspace: '\u2009', medspace: '\u2005', enspace: '\u2002' };
/** 大运算符：m:nary 的 chr 与上下限位置（求和类用 undOvr，积分类用 subSup）。 */
const MATH_NARY = {
    sum: { chr: '∑', limLoc: 'undOvr' },
    prod: { chr: '∏', limLoc: 'undOvr' },
    coprod: { chr: '∐', limLoc: 'undOvr' },
    bigcup: { chr: '⋃', limLoc: 'undOvr' },
    bigcap: { chr: '⋂', limLoc: 'undOvr' },
    int: { chr: '∫', limLoc: 'subSup' },
    oint: { chr: '∮', limLoc: 'subSup' },
    iint: { chr: '∬', limLoc: 'subSup' },
    iiint: { chr: '∭', limLoc: 'subSup' },
};
/** 函数名一律正体；带「正下方极限」的那些用 m:limLow 而不是右侧下标。 */
const MATH_FUNCTIONS = new Set([
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'log', 'ln', 'lg', 'exp', 'lim', 'max', 'min',
    'sup', 'inf', 'det', 'dim', 'gcd', 'deg', 'arg', 'ker', 'hom', 'Pr', 'mod', 'bmod',
]);
const MATH_LIMIT_FUNCTIONS = new Set([
    'lim', 'max', 'min', 'sup', 'inf', 'det', 'gcd', 'arg', 'ker', 'hom', 'dim', 'Pr',
]);
/** 定界符命令（\left\langle … \right\rangle）。 */
const MATH_DELIM_WORDS = {
    langle: '⟨', rangle: '⟩', lfloor: '⌊', rfloor: '⌋', lceil: '⌈', rceil: '⌉',
    vert: '|', Vert: '‖', lvert: '|', rvert: '|', lVert: '‖', rVert: '‖', backslash: '\\',
};
/** 节点里所有「装子节点」的键：重排样式（\mathrm / \mathbf）与 nary 收编都要遍历它们。 */
const MATH_CHILD_KEYS = ['children', 'num', 'den', 'deg', 'base', 'sub', 'sup', 'body', 'arg', 'lim'];
/** 含中文的文本要走 w:rFonts 的 eastAsia，否则公式里的中文会回落到默认宋体。 */
const CJK_RE = /[\u2e80-\u9fff\u3040-\u30ff\uff00-\uffef]/;

/**
 * 查符号表。只认自有属性：`\constructor` / `\toString` 这类命令会命中
 * Object.prototype 上的成员，不挡掉的话会把一个函数当成符号塞进公式里。
 */
const tableGet = (table, key) => (Object.prototype.hasOwnProperty.call(table, key) ? table[key] : undefined);

/** 一个数学 run：m:rPr（nor/sty）+ 可选 w:rPr + m:t，顺序不许颠倒。 */
function mRunXml(text, opts = {}, ctx) {
    const value = String(text ?? '');
    if (value === '') return '';
    const rpr = [];
    if (opts.nor) rpr.push('<m:nor/>');
    if (opts.sty) rpr.push(`<m:sty m:val="${opts.sty}"/>`);
    // 只有中文或显式字号才写 w:rPr：给每个数学 run 都挂 rFonts/sz 会把公式 XML 撑大好几倍
    const useFont = opts.nor || CJK_RE.test(value);
    const wrpr = rPrXml({
        font: useFont ? { cn: ctx.fonts.cn, en: ctx.fonts.en } : undefined,
        size: opts.size || undefined,
        color: opts.color || undefined,
    }, ctx);
    return `<m:r>${rpr.length ? `<m:rPr>${rpr.join('')}</m:rPr>` : ''}${wrpr}<m:t xml:space="preserve">${escapeXml(value)}</m:t></m:r>`;
}

function ommlNodes(nodes, ctx, opts) {
    return (nodes ?? []).map((node) => ommlXml(node, ctx, opts)).join('');
}

/** 节点 → OMML。每个 case 的元素顺序都按 CT_* 的 sequence 写死。 */
function ommlXml(node, ctx, opts) {
    const inner = (list) => ommlNodes(list, ctx, opts);
    const runOpts = { sty: node.sty, size: opts.size, color: opts.color };
    switch (node.type) {
        case 'row':
            return inner(node.children);
        case 'r':
            return mRunXml(node.text, runOpts, ctx);
        case 'text':
            return mRunXml(node.text, { nor: true, size: opts.size, color: opts.color }, ctx);
        case 'frac':
            return `<m:f><m:num>${inner(node.num)}</m:num><m:den>${inner(node.den)}</m:den></m:f>`;
        case 'rad': {
            const deg = node.deg ?? [];
            return `<m:rad><m:radPr><m:degHide m:val="${deg.length ? 0 : 1}"/></m:radPr>`
                + `<m:deg>${inner(deg)}</m:deg><m:e>${inner(node.base)}</m:e></m:rad>`;
        }
        case 'sup':
            return `<m:sSup><m:e>${inner(node.base)}</m:e><m:sup>${inner(node.sup)}</m:sup></m:sSup>`;
        case 'sub':
            return `<m:sSub><m:e>${inner(node.base)}</m:e><m:sub>${inner(node.sub)}</m:sub></m:sSub>`;
        case 'subsup':
            return `<m:sSubSup><m:e>${inner(node.base)}</m:e><m:sub>${inner(node.sub)}</m:sub>`
                + `<m:sup>${inner(node.sup)}</m:sup></m:sSubSup>`;
        case 'nary': {
            const sub = node.sub ?? [];
            const sup = node.sup ?? [];
            return '<m:nary><m:naryPr>'
                + `<m:chr m:val="${escapeXml(node.chr)}"/>`
                + `<m:limLoc m:val="${node.limLoc}"/>`
                + `<m:subHide m:val="${sub.length ? 0 : 1}"/><m:supHide m:val="${sup.length ? 0 : 1}"/>`
                + `</m:naryPr><m:sub>${inner(sub)}</m:sub><m:sup>${inner(sup)}</m:sup>`
                + `<m:e>${inner(node.body ?? [])}</m:e></m:nary>`;
        }
        case 'delim':
            return '<m:d><m:dPr>'
                + `<m:begChr m:val="${escapeXml(node.beg)}"/><m:endChr m:val="${escapeXml(node.end)}"/>`
                + `</m:dPr><m:e>${inner(node.children)}</m:e></m:d>`;
        case 'func':
            // 函数名一律正体：数学区默认是斜体，sin/lim 斜着写就不是函数名了
            return `<m:func><m:fName>${mRunXml(node.name, { ...runOpts, sty: 'p' }, ctx)}</m:fName>`
                + `<m:e>${inner(node.arg ?? [])}</m:e></m:func>`;
        case 'funcName':
            return mRunXml(node.name, { ...runOpts, sty: 'p' }, ctx);
        case 'limLow':
            return `<m:limLow><m:e>${inner(node.base)}</m:e><m:lim>${inner(node.lim)}</m:lim></m:limLow>`;
        default:
            return '';
    }
}

/**
 * LaTeX 子集 → OMML 的 m:oMath 片段。
 * 解析过程中的告警直接写进 ctx.warnings（同一条消息在一份文档里只出现一次）。
 */
function mathToOmml(latex, ctx, opts = {}) {
    const source = String(latex ?? '').replace(/\r\n?/g, ' ');
    const seen = new Set();
    const size = asNumber(opts.size, 0);
    let at = 0;

    const warn = (key, message) => {
        // 同一个命令只报一次：一段公式里 \foo 出现十次也只需要提醒一次
        if (seen.has(key) || ctx.warnings.includes(message)) return;
        seen.add(key);
        ctx.warnings.push(message);
    };
    const eof = () => at >= source.length;
    const peek = (offset = 0) => source[at + offset] ?? '';
    const isLetter = (ch) => (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');
    const skipSpaces = () => { while (!eof() && /\s/.test(peek())) at += 1; };

    /** 一串同类字符：字母连写算一个斜体变量（dx、abc），数字与括号一律正体。 */
    function plainRun() {
        const italic = isLetter(peek());
        const start = at;
        while (!eof()) {
            const ch = peek();
            // 收尾符与 ^ _ { } 必须断开，否则 x^2} 这种会把 '2}' 当成一整串
            if (ch === '\\' || ch === '{' || ch === '}' || ch === ']' || ch === ')' || ch === '^' || ch === '_' || /\s/.test(ch)) break;
            if (isLetter(ch) !== italic) break;
            at += 1;
        }
        return { type: 'r', text: source.slice(start, at), sty: italic ? '' : 'p' };
    }

    /** ^ / _ 的单个 token 参数：LaTeX 里 x^ab 的指数只是 a。 */
    function singleToken() {
        const ch = peek();
        at += 1;
        return [{ type: 'r', text: ch, sty: isLetter(ch) ? '' : 'p' }];
    }

    /** ^ / _ 的参数：{…} 分组或单个 token。 */
    function scriptArg() {
        skipSpaces();
        if (eof()) return [];
        if (peek() === '{') { at += 1; return expression('}'); }
        if (peek() === '\\') { const node = command(); return Array.isArray(node) ? node : (node ? [node] : []); }
        return singleToken();
    }

    /** 命令的必需参数（\frac/\sqrt）；缺失时给空参数并告警，绝不抛。 */
    function groupArg(name) {
        skipSpaces();
        if (eof()) {
            warn(`missing:${name}`, `公式里的 \\${name} 缺少参数，已按空参数处理`);
            return [];
        }
        if (peek() === '{') { at += 1; return expression('}'); }
        if (peek() === '\\') { const node = command(); return Array.isArray(node) ? node : (node ? [node] : []); }
        return singleToken();
    }

    /** \text{…} / \mathrm{…} 这类原样文本：只按花括号配平，不解析内部命令。 */
    function rawGroup(name) {
        skipSpaces();
        if (peek() !== '{') {
            warn(`missing:${name}`, `公式里的 \\${name} 缺少 {…} 参数，已按单个 token 处理`);
            return eof() ? '' : source[at++];
        }
        at += 1;
        let depth = 0;
        let out = '';
        while (!eof()) {
            const ch = peek();
            if (ch === '{') depth += 1;
            else if (ch === '}') {
                if (depth === 0) { at += 1; break; }
                depth -= 1;
            }
            out += ch;
            at += 1;
        }
        return out;
    }

    /** 定界符：. 表示隐形，\{ \} \langle 之类按表翻译。 */
    function delimChar() {
        if (eof()) return '';
        const ch = peek();
        if (ch === '.') { at += 1; return ''; }
        if (ch !== '\\') { at += 1; return ch; }
        at += 1;
        if (!isLetter(peek())) { const single = peek(); at += 1; return tableGet(MATH_ESCAPE_CHARS, single) ?? single; }
        const start = at;
        while (!eof() && isLetter(peek())) at += 1;
        const word = source.slice(start, at);
        return tableGet(MATH_DELIM_WORDS, word) ?? word;
    }

    const atRight = () => peek() === '\\' && /^right(?![A-Za-z])/.test(source.slice(at + 1));
    function rightDelim() {
        if (!atRight()) {
            warn('missing-right', '公式里的 \\left 没有配对的 \\right，已按缺少右定界符处理');
            return '';
        }
        at += '\\right'.length;
        skipSpaces();
        return delimChar();
    }

    function leftRight() {
        skipSpaces();
        const beg = delimChar();
        const children = expression('', true);
        return { type: 'delim', beg, end: rightDelim(), children };
    }

    /** 普通括号：\sin(x) 里的参数要连括号一起进 m:e，否则视觉效果会丢括号。 */
    function bracketGroup(open) {
        const close = open === '(' ? ')' : ']';
        at += 1;
        return [{ type: 'delim', beg: open, end: close, children: expression(close) }];
    }

    /** 函数的参数：单个原子、{…}、\cmd 或括号组；没有可吃的东西就给空 m:e。 */
    function functionArg() {
        skipSpaces();
        if (eof()) return [];
        const ch = peek();
        if (ch === '}' || ch === ']' || ch === ')' || ch === '&' || ch === '^' || ch === '_') return [];
        if (ch === '{') { at += 1; return expression('}'); }
        if (ch === '\\') { const node = command(); return Array.isArray(node) ? node : (node ? [node] : []); }
        if (ch === '(' || ch === '[') return bracketGroup(ch);
        return singleToken();
    }

    /** 递归重排样式：\mathrm 全部正体，\mathbf 全部加粗。 */
    function restyle(nodes, sty) {
        for (const node of nodes) {
            if (!node || typeof node !== 'object') continue;
            if (node.type === 'r') { node.sty = sty; continue; }
            if (node.type === 'text') continue;
            for (const key of MATH_CHILD_KEYS) {
                if (Array.isArray(node[key])) restyle(node[key], sty);
            }
        }
        return nodes;
    }

    function attach(nodes, mark, script) {
        const base = nodes.length ? nodes[nodes.length - 1] : null;
        const slot = mark === '^' ? 'sup' : 'sub';
        if (base && base.type === 'nary') { base[slot] = script; return; }
        if (base && base.type === 'funcName' && mark === '_') {
            // \lim_{x \to 0} 用 m:limLow（极限写在正下方），这才是 Word 里极限的存法
            nodes[nodes.length - 1] = { type: 'limLow', base: [base], lim: script };
            return;
        }
        if (!base) {
            // 公式以 ^ 或 _ 开头时补一个空基底：m:sSub/m:sSup 没有 m:e 会被判成非法结构
            nodes.push({ type: mark === '^' ? 'sup' : 'sub', base: [], [slot]: script });
            return;
        }
        if (base.type === 'sub' && mark === '^') { nodes[nodes.length - 1] = { type: 'subsup', base: base.base, sub: base.sub, sup: script }; return; }
        if (base.type === 'sup' && mark === '_') { nodes[nodes.length - 1] = { type: 'subsup', base: base.base, sub: script, sup: base.sup }; return; }
        if (base.type === 'subsup' || base.type === 'sub' || base.type === 'sup') { base[slot] = script; return; }
        nodes[nodes.length - 1] = { type: mark === '^' ? 'sup' : 'sub', base: [base], [slot]: script };
    }

    function command() {
        at += 1; // 吃掉反斜杠
        if (eof()) { warn('trailing-backslash', '公式以孤立的反斜杠结尾，已忽略'); return null; }
        const ch = peek();
        if (!isLetter(ch)) {
            at += 1;
            if (ch === '\\') return null; // \\ 是换行，公式里当软换行直接吞掉
            const escape = tableGet(MATH_ESCAPE_CHARS, ch);
            if (escape !== undefined) return escape === '' ? null : { type: 'r', text: escape, sty: 'p' };
            warn(`unknown:\\${ch}`, `公式里有未识别的命令 \\${ch}，已按纯文本原样输出`);
            return { type: 'r', text: `\\${ch}`, sty: 'p' };
        }
        const start = at;
        while (!eof() && isLetter(peek())) at += 1;
        return named(source.slice(start, at));
    }

    function named(name) {
        // 符号一律正体：数学区默认斜体，α/×/∈ 斜着写是错的
        const symbol = (text) => ({ type: 'r', text, sty: 'p' });
        const greek = tableGet(MATH_GREEK, name);
        if (greek !== undefined) return symbol(greek);
        const known = tableGet(MATH_SYMBOLS, name);
        if (known !== undefined) return symbol(known);
        const space = tableGet(MATH_SPACE_WORDS, name);
        if (space !== undefined) return symbol(space);
        const nary = tableGet(MATH_NARY, name);
        if (nary) return { type: 'nary', chr: nary.chr, limLoc: nary.limLoc, sub: [], sup: [], body: [] };
        if (name === 'frac') return { type: 'frac', num: groupArg('frac'), den: groupArg('frac') };
        if (name === 'sqrt') {
            skipSpaces();
            let deg = [];
            if (peek() === '[') { at += 1; deg = expression(']'); }
            return { type: 'rad', deg, base: groupArg('sqrt') };
        }
        if (name === 'text') return { type: 'text', text: rawGroup('text') };
        if (name === 'mathrm') return { type: 'row', children: restyle(groupArg('mathrm'), 'p') };
        if (name === 'mathbf') return { type: 'row', children: restyle(groupArg('mathbf'), 'b') };
        if (name === 'left') return leftRight();
        if (name === 'right') return { type: 'r', text: delimChar(), sty: 'p' };
        if (MATH_FUNCTIONS.has(name)) {
            if (MATH_LIMIT_FUNCTIONS.has(name)) return { type: 'funcName', name };
            return { type: 'func', name, arg: functionArg() };
        }
        warn(`unknown:\\${name}`, `公式里有未识别的命令 \\${name}，已按纯文本原样输出`);
        return { type: 'r', text: `\\${name}`, sty: 'p' };
    }

    /**
     * 表达式循环。stop 是收尾符（} ] )），stopAtRight 表示遇到 \right 就交给上游。
     * 多余的收尾符与 & 对齐点当噪声丢弃 —— 模型写出来的公式经常多一个括号。
     */
    function expression(stop = '', stopAtRight = false) {
        const nodes = [];
        while (!eof()) {
            if (stopAtRight && atRight()) break;
            const ch = peek();
            if (ch === '}' || ch === ']' || ch === ')') { at += 1; if (ch === stop) break; continue; }
            if (ch === '{') { at += 1; nodes.push({ type: 'row', children: expression('}') }); continue; }
            if (ch === '&' || /\s/.test(ch)) { at += 1; continue; }
            if (ch === '~') { at += 1; nodes.push({ type: 'r', text: '\u00a0', sty: 'p' }); continue; }
            if (ch === '^' || ch === '_') { at += 1; attach(nodes, ch, scriptArg()); continue; }
            if (ch === '\\') {
                const node = command();
                if (Array.isArray(node)) nodes.push(...node);
                else if (node) nodes.push(node);
                continue;
            }
            nodes.push(plainRun());
        }
        return nodes;
    }

    /**
     * m:nary 在 OMML 里必须有一个被作用的 m:e，而 LaTeX 的 \sum 并不「吃掉」后面的
     * 内容。折中成「收编紧随其后的一个原子」：\sum_{i=1}^n a_i 得到 Σ 带 a_i，
     * 而 \int_0^1 x dx = \frac12 里的等号右边不会被卷进积分号里。
     */
    function absorbNary(nodes) {
        for (let index = 0; index < nodes.length; index += 1) {
            const node = nodes[index];
            if (node.type === 'nary' && node.body.length === 0 && index + 1 < nodes.length) {
                node.body = [nodes[index + 1]];
                nodes.splice(index + 1, 1);
            }
            for (const key of MATH_CHILD_KEYS) {
                if (Array.isArray(node[key])) absorbNary(node[key]);
            }
        }
    }

    const nodes = expression('');
    absorbNary(nodes);
    return `<m:oMath>${ommlNodes(nodes, ctx, { size, color: opts.color })}</m:oMath>`;
}

/** 公式段落：m:oMathPara 包住 m:oMath；带编号时用居中+右制表位把编号摆到行尾。 */
function renderFormula(block, ctx) {
    const info = STYLE_INFO.Formula;
    const spacing = `<w:spacing w:before="${info.before}" w:after="${info.after}" w:line="${info.line}" w:lineRule="auto"/>`;
    const para = `<m:oMathPara><m:oMathParaPr><m:jc m:val="${block.align}"/></m:oMathParaPr>${block.omml}</m:oMathPara>`;
    if (!block.number) {
        return `<w:p>${pPrXml({ styleId: 'Formula', spacing, align: block.align })}${para}</w:p>`;
    }
    const right = Math.max(1, Math.round(ctx.contentTw));
    const center = Math.max(1, Math.round(ctx.contentTw / 2));
    const tabs = block.align === 'left'
        ? `<w:tab w:val="right" w:pos="${right}"/>`
        : `<w:tab w:val="center" w:pos="${center}"/><w:tab w:val="right" w:pos="${right}"/>`;
    const lead = block.align === 'left' ? '' : '<w:r><w:tab/></w:r>';
    const numberRun = `<w:r><w:tab/>${runTextXml(block.number)}</w:r>`;
    return `<w:p>${pPrXml({ styleId: 'Formula', tabs, spacing, align: 'left' })}${lead}${para}${numberRun}</w:p>`;
}

/** 公式预览：给 outline 用，取 OMML 里所有 m:t 拼起来（不含编号 run）。 */
function mathPreview(paragraph) {
    const text = descendants(paragraph, 'm:t').map((node) => textOf(node)).join('').replace(/\s+/g, ' ').trim();
    if (text === '') return '(空)';
    const chars = [...text];
    return chars.length > 24 ? `${chars.slice(0, 24).join('')}…` : text;
}

// ---------------------------------------------------------------- 内容渲染

function resolveWidths(raw, columns, contentTw, warn) {
    const count = columns.length;
    const equal = () => Array.from({ length: count }, () => Math.floor(contentTw / count));
    let list = null;
    if (Array.isArray(raw) && raw.length > 0) {
        const nums = raw.map((value) => asNumber(value, 0));
        if (nums.every((value) => value > 0)) {
            const sum = nums.reduce((a, b) => a + b, 0);
            // 合计 ≤ 40 当成厘米；否则当成相对权重 —— 两种写法在实践里都用得到
            list = sum <= 40
                ? nums.map((value) => Math.round(value * TWIPS_PER_CM))
                : nums.map((value) => Math.round((contentTw * value) / sum));
        }
    }
    if (!list) {
        const widths = columns.map((column) => column.width);
        if (widths.some((value) => value > 0)) {
            list = widths.map((value) => Math.round((value > 0 ? value : 3) * TWIPS_PER_CM));
        }
    }
    if (!list || list.length !== count || list.some((value) => !Number.isFinite(value) || value <= 0)) list = equal();
    const total = list.reduce((a, b) => a + b, 0);
    if (total > contentTw) {
        warn(`表格列宽合计 ${(total / TWIPS_PER_CM).toFixed(2)}cm 超出版心 ${(contentTw / TWIPS_PER_CM).toFixed(2)}cm，已按比例压缩`);
        const scaled = list.map((value) => Math.floor((value * contentTw) / total));
        scaled[scaled.length - 1] += contentTw - scaled.reduce((a, b) => a + b, 0);
        list = scaled;
    }
    return list;
}

/** 单元格归一：字符串与 {text, colspan, rowspan, align, bold, italic, fill, color, size} 都收。 */
function normalizeCell(raw) {
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
        return {
            text: asString(raw.text ?? raw.value),
            colspan: Math.max(1, Math.round(asNumber(raw.colspan, 1))),
            rowspan: Math.max(1, Math.round(asNumber(raw.rowspan, 1))),
            align: asString(raw.align, ''),
            bold: !!raw.bold,
            italic: !!raw.italic,
            fill: asString(raw.fill, ''),
            color: asString(raw.color, ''),
            size: asNumber(raw.size, 0),
        };
    }
    return {
        text: asString(raw), colspan: 1, rowspan: 1, align: '',
        bold: false, italic: false, fill: '', color: '', size: 0,
    };
}

/**
 * 每行的补齐规则分两种情况：
 *  - 整表没有合并：沿用 kit.normalizeRows（按单元格个数补齐/截断），行为与以前一字不差；
 *  - 有合并：不能按单元格个数补齐 —— rowspan 下方的行本来就会少几个格，补出来的空格
 *    会挤到被上方占用的列上。这时补齐交给 layoutTableRows 按网格做。
 */
function normalizeTableRows(rows, columnCount) {
    const source = asArray(rows);
    const list = source.map((row) => (Array.isArray(row) ? row.slice() : [row]));
    const merged = list.some((cells) => cells.some((cell) => {
        const parsed = normalizeCell(cell);
        return parsed.colspan > 1 || parsed.rowspan > 1;
    }));
    if (!merged) return source.map((row) => normalizeRows([row], columnCount)[0]);
    return list;
}

/**
 * 把带合并的行摊到网格上，再按网格顺序输出每行的格子。
 * 为什么要过一层网格：rowspan 的占位格（<w:vMerge/>）必须落在被合并的那一列上，
 * 光看「这一行有几个单元格」是算不出位置的。
 * 越界的 colspan/rowspan 一律裁剪并写 warning —— 宁可少一格，也不要写出 Word 打不开的结构。
 */
function layoutTableRows(rows, columnCount, warn) {
    const grid = rows.map(() => new Array(columnCount).fill(null));
    for (let r = 0; r < rows.length; r += 1) {
        let col = 0;
        for (const raw of rows[r]) {
            while (col < columnCount && grid[r][col]) col += 1;
            if (col >= columnCount) {
                warn(`表格第 ${r + 1} 行的单元格超出 ${columnCount} 列，多出的已丢弃`);
                break;
            }
            const cell = normalizeCell(raw);
            let colspan = cell.colspan;
            if (col + colspan > columnCount) {
                warn(`表格第 ${r + 1} 行第 ${col + 1} 格的 colspan=${cell.colspan} 越界（该行只剩 ${columnCount - col} 列），已裁剪为 ${columnCount - col}`);
                colspan = columnCount - col;
            }
            let rowspan = cell.rowspan;
            if (r + rowspan > rows.length) {
                warn(`表格第 ${r + 1} 行第 ${col + 1} 格的 rowspan=${cell.rowspan} 越界（表格只有 ${rows.length} 行），已裁剪为 ${rows.length - r}`);
                rowspan = rows.length - r;
            }
            // 下方已经被别的合并占用时不能再压上去：裁到冲突行之前，结构才不会打架
            for (let rr = r + 1; rr < r + rowspan; rr += 1) {
                let blocked = false;
                for (let cc = col; cc < col + colspan; cc += 1) if (grid[rr][cc]) blocked = true;
                if (blocked) {
                    warn(`表格第 ${r + 1} 行第 ${col + 1} 格的 rowspan 与下方已有合并冲突，已裁剪为 ${rr - r}`);
                    rowspan = Math.max(1, rr - r);
                    break;
                }
            }
            const node = { ...cell, colspan, rowspan, row: r, col };
            for (let rr = r; rr < r + rowspan; rr += 1) {
                for (let cc = col; cc < col + colspan; cc += 1) grid[rr][cc] = node;
            }
            col += colspan;
        }
        // 行尾补齐到整列（跳过被上方 rowspan 占用的格子）：Word 的固定列宽表里
        // 行尾缺格会让这一行整体短一截，补出来的空格子照常参与斑马纹与边框
        while (col < columnCount) {
            if (grid[r][col]) { col += 1; continue; }
            grid[r][col] = { ...normalizeCell(''), colspan: 1, rowspan: 1, row: r, col };
            col += 1;
        }
    }
    return rows.map((row, r) => {
        const out = [];
        let col = 0;
        while (col < columnCount) {
            const owner = grid[r][col];
            if (!owner) { out.push({ kind: 'empty', col, span: 1 }); col += 1; continue; }
            out.push(owner.row === r ? { kind: 'cell', cell: owner } : { kind: 'merge', cell: owner });
            col += owner.colspan;
        }
        return out;
    });
}

function renderTable(block, ctx) {
    const columns = normalizeColumns(block.columns);
    const rows = normalizeTableRows(block.rows, columns.length);
    const warn = (message) => ctx.warnings.push(message);
    const widths = resolveWidths(block.widths, columns, ctx.contentTw, warn);
    const totalTw = widths.reduce((a, b) => a + b, 0);
    const mode = ['plain', 'striped', 'accent'].includes(asString(block.style, 'striped')) ? asString(block.style, 'striped') : 'striped';
    const ink = ctx.ink;
    // 默认重复表头：跨页的长表格没有表头，翻页后一列都认不出来
    const headerRepeat = block.headerRepeat !== false;
    const tblAlign = alignOf(block.align, '');
    const gridRows = layoutTableRows(rows, columns.length, warn);
    // 跨列格的宽度是它吃掉的各列之和，否则 Word 的固定列宽和 tblGrid 会对不上
    const spanWidth = (col, span) => {
        let total = 0;
        for (let i = 0; i < span; i += 1) total += widths[col + i] ?? 0;
        return total;
    };

    const tableAlign = tblAlign && tblAlign !== 'both' && tblAlign !== 'distribute' ? `<w:jc w:val="${tblAlign}"/>` : '';
    const tblPr = '<w:tblPr>'
        + '<w:tblStyle w:val="TableGrid"/>'
        + `<w:tblW w:w="${totalTw}" w:type="dxa"/>`
        + tableAlign
        + `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => borderXml(side, { sz: 4, color: ink.border })).join('')}</w:tblBorders>`
        + '<w:tblLayout w:type="fixed"/>'
        + '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
        + '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
        + '</w:tblPr>';
    const grid = `<w:tblGrid>${widths.map((value) => `<w:gridCol w:w="${value}"/>`).join('')}</w:tblGrid>`;

    // 表头填充同样来自主题：素色主题没有底色，只靠加粗与一条下框线区分表头
    const headFill = shadeOf(ink.table.headerFill);
    const headCells = columns.map((column, index) => {
        const cellPr = `<w:tcPr><w:tcW w:w="${widths[index]}" w:type="dxa"/>`
            + `<w:tcBorders>${borderXml('bottom', { sz: 12, color: ink.primaryDark })}</w:tcBorders>`
            + `${headFill ? `<w:shd w:val="clear" w:color="auto" w:fill="${headFill}"/>` : ''}<w:vAlign w:val="center"/></w:tcPr>`;
        const para = pPrXml({ styleId: 'TableText', align: 'center' })
            + runXml(column.title, { bold: true, color: ink.table.headerText, size: STYLE_INFO.TableText.size }, ctx);
        return `<w:tc>${cellPr}<w:p>${para}</w:p></w:tc>`;
    }).join('');
    const headRow = `<w:tr><w:trPr><w:trHeight w:val="397" w:hRule="atLeast"/>${headerRepeat ? '<w:tblHeader/>' : ''}</w:trPr>${headCells}</w:tr>`;

    const emptyCell = (col, span) => `<w:tc><w:tcPr><w:tcW w:w="${spanWidth(col, span)}" w:type="dxa"/>`
        + `${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}<w:vAlign w:val="center"/></w:tcPr>`
        + `<w:p>${pPrXml({ styleId: 'TableText', align: 'left' })}</w:p></w:tc>`;

    const bodyRows = gridRows.map((slots, rowIndex) => {
        const zebra = mode === 'striped' && rowIndex % 2 === 1;
        const cells = slots.map((slot) => {
            if (slot.kind === 'empty') return emptyCell(slot.col, slot.span);
            const cell = slot.cell;
            // 占位格：上方 rowspan 的延续。宽度与起始格一致，Word 才把这一列连起来
            if (slot.kind === 'merge') {
                return `<w:tc><w:tcPr><w:tcW w:w="${spanWidth(cell.col, cell.colspan)}" w:type="dxa"/>`
                    + `${cell.colspan > 1 ? `<w:gridSpan w:val="${cell.colspan}"/>` : ''}<w:vMerge/>`
                    + `<w:vAlign w:val="center"/></w:tcPr>`
                    + `<w:p>${pPrXml({ styleId: 'TableText', align: 'left' })}</w:p></w:tc>`;
            }
            const column = columns[cell.col] ?? columns[0] ?? { align: '' };
            const autoAlign = (() => {
                const type = classifyValue(cell.text);
                if (['number', 'percent', 'currency'].includes(type)) return 'right';
                if (type === 'date') return 'center';
                return 'left';
            })();
            const align = alignOf(cell.align || column.align, autoAlign);
            // 底色一律来自主题或单元格显式 fill，且都过一遍 shadeOf：
            // 素色主题的 fill 是空串，于是这里根本不写 <w:shd>，公式表格才不会有白底
            const explicitFill = shadeOf(cell.fill);
            const themeFill = mode === 'accent' && cell.col === 0 ? ink.table.accentFill : (zebra ? ink.table.zebraFill : '');
            const fill = explicitFill || shadeOf(themeFill);
            const cellPr = `<w:tcPr><w:tcW w:w="${spanWidth(cell.col, cell.colspan)}" w:type="dxa"/>`
                + `${cell.colspan > 1 ? `<w:gridSpan w:val="${cell.colspan}"/>` : ''}`
                + `${cell.rowspan > 1 ? '<w:vMerge w:val="restart"/>' : ''}`
                + `${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}<w:vAlign w:val="center"/></w:tcPr>`;
            const para = pPrXml({ styleId: 'TableText', align })
                + runXml(cell.text, {
                    bold: cell.bold || (mode === 'accent' && cell.col === 0),
                    italic: cell.italic,
                    color: cell.color || undefined,
                    size: cell.size > 0 ? cell.size : undefined,
                }, ctx);
            return `<w:tc>${cellPr}<w:p>${para}</w:p></w:tc>`;
        }).join('');
        return `<w:tr><w:trPr><w:trHeight w:val="340" w:hRule="atLeast"/></w:trPr>${cells}</w:tr>`;
    }).join('');

    // 题注用 Caption 样式段落：above（中文表格惯例，表题在上）或 below，靠 keepNext 让它跟住表格
    const caption = asString(block.caption);
    const captionBelow = asString(block.captionPosition, 'above').toLowerCase() === 'below';
    const captionPara = caption === '' ? '' : `<w:p>${pPrXml({ styleId: 'Caption', keepNext: !captionBelow })}`
        + `${runXml(caption, { size: STYLE_INFO.Caption.size, color: ink.muted }, ctx)}</w:p>`;

    // 表格后补一个 4pt 空段：Word 里紧贴的表格与正文之间没有任何间距，视觉上会糊在一起
    const tail = `<w:p>${pPrXml({ spacing: '<w:spacing w:before="0" w:after="80" w:line="240" w:lineRule="exact"/>' })}</w:p>`;
    const tableXml = `<w:tbl>${tblPr}${grid}${headRow}${bodyRows}</w:tbl>`;
    return `${captionBelow ? '' : captionPara}${tableXml}${captionBelow ? captionPara : ''}${tail}`;
}

function renderImage(block, ctx) {
    const { widthCm, caption, alt } = block;
    const cx = Math.round(widthCm * EMU_PER_CM);
    const cy = block.heightCm ? Math.round(block.heightCm * EMU_PER_CM) : Math.round(cx * 0.6);
    const relId = block.relId;
    const drawingId = block.drawingId;
    const name = `图片 ${drawingId}`;
    const drawing = '<w:r><w:drawing>'
        + '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        + `<wp:extent cx="${cx}" cy="${cy}"/>`
        + '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        + `<wp:docPr id="${drawingId}" name="${escapeXml(name)}" descr="${escapeXml(alt ?? caption ?? '')}"/>`
        + '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        + `<a:graphic><a:graphicData uri="${PIC_NS}">`
        + `<pic:pic><pic:nvPicPr><pic:cNvPr id="${drawingId}" name="${escapeXml(block.fileName)}"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
        + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const imagePara = `<w:p>${pPrXml({ align: 'center', spacing: '<w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/>' })}${drawing}</w:p>`;
    if (!caption) return imagePara;
    const captionPara = `<w:p>${pPrXml({ styleId: 'Caption' })}${runXml(caption, { size: STYLE_INFO.Caption.size, color: ctx.ink.muted }, ctx)}</w:p>`;
    return imagePara + captionPara;
}

function pageFieldRuns(template, ctx, opts = {}) {
    // 约定：模板里的 X 是当前页码、Y 是总页数；没有占位符就输出静态文字
    const pieces = String(template).split(/([XY])/);
    return pieces.map((piece) => {
        if (piece === 'X') return fieldRunXml(' PAGE ', '1', opts, ctx);
        if (piece === 'Y') return fieldRunXml(' NUMPAGES ', '1', opts, ctx);
        if (piece === '') return '';
        return runXml(piece, opts, ctx);
    }).join('');
}

function fieldRunXml(instr, cached, opts, ctx) {
    const rp = rPrXml(opts, ctx);
    return `<w:r>${rp}<w:fldChar w:fldCharType="begin"/></w:r>`
        + `<w:r>${rp}<w:instrText xml:space="preserve">${escapeXml(instr)}</w:instrText></w:r>`
        + `<w:r>${rp}<w:fldChar w:fldCharType="separate"/></w:r>`
        + `<w:r>${rp}<w:t>${escapeXml(cached)}</w:t></w:r>`
        + `<w:r>${rp}<w:fldChar w:fldCharType="end"/></w:r>`;
}

function footerXml(ctx) {
    const runs = pageFieldRuns(ctx.footerTemplate, ctx, {});
    return `${XML_HEAD}<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}">`
        + `<w:p>${pPrXml({ styleId: 'Footer' })}${runs}</w:p>`
        + '</w:ftr>';
}

function headerXml(ctx) {
    const runs = String(ctx.headerTemplate).split('\n').map((line) => runXml(line, {}, ctx)).join('<w:br/>');
    return `${XML_HEAD}<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}">`
        + `<w:p>${pPrXml({ styleId: 'Header' })}${runs}</w:p>`
        + '</w:hdr>';
}

/** 制表符必须落在 run 内部（`<w:tab/>`），写成 `<w:t>\t</w:t>` Word 不认。 */
const tabRunXml = (opts, ctx) => `<w:r>${rPrXml(opts, ctx)}<w:tab/></w:r>`;

/**
 * 渲染前把目录的条目、锚点与页码定好。
 *
 * 目录里的超链接锚点必须与标题上的书签名同名，页码又依赖分页估算，所以
 * `create()` 分两遍渲染：第一遍用占位页码，`analyzeDocument` 顺手报出每个标题
 * 落在第几页；第二遍（`ctx.tocPages` 有值时）换成真实页码。两遍的目录条目完全
 * 相同，因此目录自身占的高度不变，第一遍估出的页码在第二遍里仍然成立。
 *
 * 幂等：两遍都从这里走，标题上的书签与目录条目不会重复累加。
 */
function prepareToc(ctx) {
    const tocBlock = ctx.blocks.find((block) => block.type === 'toc');
    if (!tocBlock) return;
    const headings = ctx.blocks.filter((block) => block.type === 'heading');
    headings.forEach((block, index) => {
        block.anchor = `${TOC_ANCHOR_PREFIX}${index + 1}`;
        block.bookmarkId = index + 1;
    });
    const pages = ctx.tocPages ?? [];
    tocBlock.entries = headings.map((block, index) => ({
        level: Math.max(1, Math.min(3, block.level)),
        text: runsText(block.runs),
        anchor: block.anchor,
        page: pages[index] ?? 1,
    }));
}

/**
 * 目录：域的缓存结果里**预渲染**好条目与页码。
 *
 * 为什么不用 `w:dirty` + settings 里的 `w:updateFields` 让 Word 自动刷新：
 * 那两种标记都会让 Word 打开时弹「该文档包含的域可能引用了其他文件。是否更新
 * 该文档中的这些域？」（真实 Word 实测，去掉一处仍弹）。预渲染换来的是「打开即见、
 * 零弹窗」，代价是页码是引擎的分页估算值，按 F9 可以让 Word 重算成精确值。
 *
 * 域的 begin/separate 放在第一条目段落、end 放在最后一条目段落 —— 真实 Word
 * 生成的目录就是这个形状，条目各占一段才能逐级缩进。
 */
function renderToc(block, ctx) {
    const entries = block.entries ?? [];
    // 右对齐 + 点线前导符：页码贴版心右边缘，条目文字与页码之间自动补点
    const tabs = `<w:tab w:val="right" w:leader="dot" w:pos="${Math.round(ctx.contentTw)}"/>`;
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        + `<w:r><w:instrText xml:space="preserve">TOC \\o "1-3" \\h \\z \\u</w:instrText></w:r>`
        + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    if (entries.length === 0) {
        // 没有 H1–H3：域的壳保持完整，只给一行说明，避免目录区一片空白。
        // 这一段**不带 TOC1 样式**，否则 read() 会把它当成一条目录条目来数。
        return `<w:p>${pPrXml({ tabs, spacing: '<w:spacing w:before="0" w:after="240" w:line="300" w:lineRule="auto"/>', align: 'left' })}${begin}`
            + runXml('（文档里还没有 H1–H3 标题，按 F9 可刷新目录）', { size: 9, color: ctx.ink.muted }, ctx)
            + `${end}</w:p>`;
    }
    const last = entries.length - 1;
    return entries.map((entry, index) => {
        const toc = TOC_INFO[entry.level - 1];
        const opts = { size: toc.size, bold: toc.bold, color: toc.level === 1 ? ctx.ink.text : ctx.ink.muted };
        const label = `<w:hyperlink w:anchor="${escapeXml(entry.anchor)}">${runXml(entry.text, opts, ctx)}</w:hyperlink>`;
        const page = runXml(String(entry.page), { size: toc.size, color: ctx.ink.muted }, ctx);
        return `<w:p>${pPrXml({ styleId: `TOC${toc.level}`, tabs, align: 'left' })}`
            + `${index === 0 ? begin : ''}${label}${tabRunXml({ size: toc.size }, ctx)}${page}`
            + `${index === last ? end : ''}</w:p>`;
    }).join('');
}

function renderBlock(block, ctx) {
    switch (block.type) {
        case 'title': {
            const para = `<w:p>${pPrXml({ styleId: 'Title' })}${runsXml(block.runs, {}, ctx)}</w:p>`;
            if (!block.subtitle) return para;
            return para + `<w:p>${pPrXml({ styleId: 'Subtitle' })}${runXml(block.subtitle, {}, ctx)}</w:p>`;
        }
        case 'heading': {
            // 书签包住标题文字：目录里的超链接（w:hyperlink w:anchor）靠它跳转
            const bookmark = block.anchor
                ? `<w:bookmarkStart w:id="${block.bookmarkId}" w:name="${escapeXml(block.anchor)}"/>`
                : '';
            const bookmarkEnd = block.anchor ? `<w:bookmarkEnd w:id="${block.bookmarkId}"/>` : '';
            return `<w:p>${pPrXml({ styleId: `Heading${block.level}` })}${bookmark}${runsXml(block.runs, block.opts ?? {}, ctx)}${bookmarkEnd}</w:p>`;
        }
        case 'para': {
            const info = STYLE_INFO.Normal;
            const spacing = spacingOf(info, block.opts ?? {});
            const indent = indentAttr(indentTwips(block.opts?.indent, null));
            return `<w:p>${pPrXml({
                styleId: 'Normal',
                spacing: `<w:spacing w:before="${spacing.before}" w:after="${spacing.after}" w:line="${spacing.line}" w:lineRule="${spacing.lineRule}"/>`,
                indent,
                align: alignOf(block.opts?.align, info.align),
            })}${runsXml(block.runs, block.opts ?? {}, ctx)}</w:p>`;
        }
        case 'list': {
            const numId = block.ordered ? 2 : 1;
            const level = Math.max(0, Math.min(2, asNumber(block.opts?.level, 0)));
            const info = STYLE_INFO.ListParagraph;
            const left = 420 + level * 420;
            const spacing = spacingOf(info, block.opts ?? {});
            return block.items.map((runs, index) => `<w:p>${pPrXml({
                styleId: 'ListParagraph',
                numId,
                ilvl: level,
                spacing: `<w:spacing w:before="${spacing.before}" w:after="${index === block.items.length - 1 ? 160 : spacing.after}" w:line="${spacing.line}" w:lineRule="${spacing.lineRule}"/>`,
                indent: `w:left="${left}" w:hanging="240"`,
                align: 'both',
            })}${runsXml(runs, block.opts ?? {}, ctx)}</w:p>`).join('');
        }
        case 'quote': {
            const info = STYLE_INFO.Quote;
            const spacing = spacingOf(info, block.opts ?? {});
            return `<w:p>${pPrXml({
                styleId: 'Quote',
                spacing: `<w:spacing w:before="${spacing.before}" w:after="${spacing.after}" w:line="${spacing.line}" w:lineRule="${spacing.lineRule}"/>`,
                indent: `w:left="${info.indentLeft}" w:right="${info.indentRight}"`,
                align: 'left',
            })}${runsXml(block.runs, block.opts ?? {}, ctx)}</w:p>`;
        }
        case 'code': {
            const info = STYLE_INFO.Code;
            return `<w:p>${pPrXml({
                styleId: 'Code',
                spacing: `<w:spacing w:before="${info.before}" w:after="${info.after}" w:line="${info.line}" w:lineRule="auto"/>`,
                indent: `w:left="${info.indentLeft}" w:right="${info.indentRight}"`,
                align: 'left',
            })}${runXml(block.text, { size: info.size }, ctx)}</w:p>`;
        }
        case 'table':
            return renderTable(block, ctx);
        case 'formula':
            return renderFormula(block, ctx);
        case 'image':
            return renderImage(block, ctx);
        case 'pageBreak':
            return `<w:p>${pPrXml({ spacing: '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' })}<w:r><w:br w:type="page"/></w:r></w:p>`;
        case 'spacer': {
            const line = Math.max(1, Math.round(asNumber(block.cm, 0.4) * TWIPS_PER_CM));
            return `<w:p>${pPrXml({ spacing: `<w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="exact"/>` })}</w:p>`;
        }
        case 'toc':
            return renderToc(block, ctx);
        default:
            return '';
    }
}

function sectPrXml(ctx) {
    const page = ctx.page;
    const refs = [];
    if (ctx.relHeader) refs.push(`<w:headerReference w:type="default" r:id="${ctx.relHeader}"/>`);
    if (ctx.relFooter) refs.push(`<w:footerReference w:type="default" r:id="${ctx.relFooter}"/>`);
    const orient = page.landscape ? ' w:orient="landscape"' : '';
    return `<w:sectPr>${refs.join('')}`
        + `<w:pgSz w:w="${page.widthTw}" w:h="${page.heightTw}"${orient}/>`
        + `<w:pgMar w:top="${page.margin.top}" w:right="${page.margin.right}" w:bottom="${page.margin.bottom}" w:left="${page.margin.left}" w:header="${page.headerTw}" w:footer="${page.footerTw}" w:gutter="0"/>`
        + '<w:cols w:space="425"/>'
        + '</w:sectPr>';
}

function documentXml(ctx) {
    prepareToc(ctx);
    const body = ctx.blocks.map((block) => renderBlock(block, ctx)).join('');
    // xmlns:m 始终声明：公式（m:oMath / m:oMathPara）没有它 Word 会判整份文档非法，
    // 而多声明一个没用到的命名空间是无害的
    return `${XML_HEAD}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:m="${M_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">`
        + `<w:body>${body}${sectPrXml(ctx)}</w:body>`
        + '</w:document>';
}

function documentRelsXml(ctx) {
    const rels = [
        `<Relationship Id="rId1" Type="${REL_TYPE.styles}" Target="styles.xml"/>`,
        `<Relationship Id="rId2" Type="${REL_TYPE.numbering}" Target="numbering.xml"/>`,
        `<Relationship Id="rId3" Type="${REL_TYPE.settings}" Target="settings.xml"/>`,
    ];
    if (ctx.relFooter) rels.push(`<Relationship Id="${ctx.relFooter}" Type="${REL_TYPE.footer}" Target="footer1.xml"/>`);
    if (ctx.relHeader) rels.push(`<Relationship Id="${ctx.relHeader}" Type="${REL_TYPE.header}" Target="header1.xml"/>`);
    for (const item of ctx.media) {
        rels.push(`<Relationship Id="${item.relId}" Type="${REL_TYPE.image}" Target="media/${item.fileName}"/>`);
    }
    return `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}">${rels.join('')}</Relationships>`;
}

function buildParts(ctx) {
    const entries = [];
    entries.push({ name: '[Content_Types].xml', data: contentTypesXml({ hasFooter: !!ctx.relFooter, hasHeader: !!ctx.relHeader, media: ctx.media }) });
    entries.push({ name: '_rels/.rels', data: rootRelsXml() });
    entries.push({ name: 'docProps/core.xml', data: coreXml(ctx) });
    entries.push({ name: 'docProps/app.xml', data: appXml(ctx) });
    entries.push({ name: 'word/document.xml', data: documentXml(ctx) });
    entries.push({ name: 'word/_rels/document.xml.rels', data: documentRelsXml(ctx) });
    entries.push({ name: 'word/styles.xml', data: stylesXml(ctx.theme, ctx.ink, ctx.fonts) });
    entries.push({ name: 'word/numbering.xml', data: numberingXml(ctx.fonts) });
    entries.push({ name: 'word/settings.xml', data: settingsXml() });
    if (ctx.relFooter) entries.push({ name: 'word/footer1.xml', data: footerXml(ctx) });
    if (ctx.relHeader) entries.push({ name: 'word/header1.xml', data: headerXml(ctx) });
    for (const item of ctx.media) entries.push({ name: `word/media/${item.fileName}`, data: item.data });
    return entries;
}

// ---------------------------------------------------------------- 解析 / 报表

function paragraphChunks(paragraph) {
    // 按文档顺序收集文本片段与换行：换行决定代码块的行数和行宽估算，
    // 光靠 descendants(w:t) 会把 <w:br/> 这段结构信息丢掉。
    const chunks = [];
    const walk = (node) => {
        for (const child of node.children) {
            if (child.name === 'w:t') chunks.push({ kind: 'text', value: textOf(child) });
            else if (child.name === 'w:br') chunks.push({ kind: 'break', page: attr(child, 'w:type') === 'page' });
            else if (child.name === 'w:tab') chunks.push({ kind: 'text', value: '\t' });
            else walk(child);
        }
    };
    walk(paragraph);
    return chunks;
}

function paragraphSize(paragraph, info) {
    for (const run of descendants(paragraph, 'w:r')) {
        const rPr = children(run, 'w:rPr')[0];
        const sz = rPr ? children(rPr, 'w:sz')[0] : undefined;
        const value = asNumber(attr(sz, 'w:val'), 0);
        if (value > 0) return value / 2;
    }
    return info.size;
}

function cellText(cell) {
    return descendants(cell, 'w:t').map((node) => textOf(node)).join('');
}

/**
 * 解析 document.xml，产出 stats/outline/warnings。
 * create 与 read 都调它，因此两边的数字天然一致；分析本身与主题无关
 * （主题只决定颜色，不影响行长与页数），所以不接 theme 参数。
 */
function analyzeDocument(documentXml) {
    const doc = parseXml(documentXml);
    const root = doc.children.find((node) => node.name === 'w:document') ?? doc.children[0];
    const body = root ? (root.children.find((node) => node.name === 'w:body') ?? descendant(root, 'w:body')) : undefined;
    if (!body) throw new Error('word: document.xml 里没有 w:body，不是可读的 Word 文档');
    const page = pageFromXml(body);
    const contentTw = page.widthTw - page.margin.left - page.margin.right;
    const contentPt = contentTw / 20;

    const stats = {
        paragraphs: 0, bodyParagraphs: 0, headings: 0, lists: 0, listItems: 0,
        quotes: 0, codeBlocks: 0, tables: 0, images: 0, words: 0, chars: 0,
        formulas: 0, pagesEstimate: 1, pageSize: page.label, contentWidthCm: Number((contentTw / TWIPS_PER_CM).toFixed(2)),
    };
    const outline = [];
    const warnings = [];
    const flow = [];
    let headingLevels = [];
    // 目录：条目（TOC1–TOC3 段落）与标题页码都要在分页之后才知道最终值，
    // 所以先收集，等分页算完再回填 outline 那一行。
    const tocEntries = [];
    const headingPages = [];
    let headingSeq = 0;
    let tocSeen = false;
    let tocDirty = false;
    let tocOutlineAt = -1;

    const textWidth = (indentTw) => Math.max(40, contentPt - indentTw / 20);

    for (const node of body.children) {
        if (node.name === 'w:p') {
            stats.paragraphs += 1;
            const pPr = children(node, 'w:pPr')[0];
            const styleNode = pPr ? children(pPr, 'w:pStyle')[0] : undefined;
            const styleId = asString(attr(styleNode, 'w:val'), 'Normal');
            const info = STYLE_INFO[styleId] ?? STYLE_INFO.Normal;
            const numPr = pPr ? children(pPr, 'w:numPr')[0] : undefined;
            const numId = numPr ? asNumber(attr(children(numPr, 'w:numId')[0], 'w:val'), 0) : 0;
            const chunks = paragraphChunks(node);
            const text = chunks.filter((chunk) => chunk.kind === 'text').map((chunk) => chunk.value).join('');
            // 公式按 m:oMath 的个数数，块级（m:oMathPara）与行内共用同一套统计
            const mathNodes = descendants(node, 'm:oMath');
            if (mathNodes.length) stats.formulas += mathNodes.length;
            const blockMath = !!descendant(node, 'm:oMathPara');
            const pageBreak = chunks.some((chunk) => chunk.kind === 'break' && chunk.page);
            const breaks = chunks.filter((chunk) => chunk.kind === 'break' && !chunk.page).length;
            const instr = descendants(node, 'w:instrText').map((child) => textOf(child)).join(' ');
            const indNode = pPr ? children(pPr, 'w:ind')[0] : undefined;
            const indentLeft = indNode ? asNumber(attr(indNode, 'w:left'), info.indentLeft ?? 0) : (info.indentLeft ?? 0);
            const spacingNode = pPr ? children(pPr, 'w:spacing')[0] : undefined;
            const lineAttr = spacingNode ? asNumber(attr(spacingNode, 'w:line'), 0) : 0;
            const lineRule = spacingNode ? asString(attr(spacingNode, 'w:lineRule'), 'auto') : 'auto';
            const before = spacingNode && attr(spacingNode, 'w:before') !== undefined ? asNumber(attr(spacingNode, 'w:before'), info.before) : info.before;
            const after = spacingNode && attr(spacingNode, 'w:after') !== undefined ? asNumber(attr(spacingNode, 'w:after'), info.after) : info.after;
            const size = paragraphSize(node, info);
            const factor = lineRule === 'auto' && lineAttr > 0 ? lineAttr / 240 : 1.3;

            const base = before / 20 + after / 20;
            // 行内公式在文本流里占宽度但不在 text 里，按每个公式约两个字宽预留，行数才不会低估
            const wrapLines = (extraIndent = 0, reserved = mathNodes.length * size * 2) => Math.max(
                1,
                estimateLines(text, Math.max(40, textWidth(indentLeft + extraIndent) - reserved), size),
                breaks + 1,
            );

            if (pageBreak) {
                outline.push('分页');
                flow.push({ kind: 'pageBreak', heightPt: 0 });
                if (text === '' && !descendants(node, 'a:blip').length) continue;
            }
            if (instr.includes('TOC')) {
                tocSeen = true;
                // w:dirty 是「打开时更新域」的另一种写法，和 settings 里的
                // w:updateFields 一样会让 Word 弹更新提示。老产物带它，新产物不带。
                tocDirty = descendants(node, 'w:fldChar').some((child) => asString(attr(child, 'w:dirty')) === 'true');
                tocOutlineAt = outline.length;
                outline.push('目录');
            }
            const tocLevel = TOC_STYLE_LEVEL[styleId] ?? 0;
            if (tocLevel > 0) {
                // 目录条目：真实 Word 生成的目录也用 TOC1–TOC3 这些样式 id。
                // 它们既不是正文段落也不算标题，只按行占高度。
                tocEntries.push({ level: tocLevel, text });
                flow.push({ kind: 'tocEntry', heightPt: wrapLines() * size * factor + base });
                continue;
            }
            if (instr.includes('TOC')) {
                // 目录域的壳（文档里没有 H1–H3 时它自己带一行说明）
                flow.push({ kind: 'toc', heightPt: wrapLines() * size * factor + base });
                continue;
            }
            if (styleId === 'Title') {
                outline.push(`封面 ${text}`);
                flow.push({ kind: 'title', heightPt: wrapLines() * size * factor + base });
                continue;
            }
            if (info.kind === 'heading') {
                stats.headings += 1;
                const level = info.level;
                outline.push(`H${level} ${text}`);
                if (headingLevels.length && level - headingLevels[headingLevels.length - 1] > 1) {
                    warnings.push(`标题层级跳跃：H${headingLevels[headingLevels.length - 1]} 之后直接出现 H${level}（“${text}”），建议补一级标题`);
                }
                headingLevels.push(level);
                if ([...text].length > 40) warnings.push(`标题“${text.slice(0, 18)}…”共 ${[...text].length} 字，可能折行，建议控制在 40 字内`);
                // headingIndex 让分页那一步能把「第几个标题落在第几页」记下来，
                // 目录的预渲染页码就是从这里来的
                flow.push({ kind: 'heading', heightPt: wrapLines() * size * factor + base, headingIndex: headingSeq });
                headingSeq += 1;
                continue;
            }
            // 图片段落的文本是空的，必须在「空段落」判定之前拦下来，
            // 否则一张插图只会被算成一行 12pt 的空白
            const extent = descendant(node, 'wp:extent');
            if (extent) {
                const cxPt = asNumber(attr(extent, 'cx'), 0) / EMU_PER_PT;
                const cyPt = asNumber(attr(extent, 'cy'), 0) / EMU_PER_PT;
                outline.push(`图片 ${(cxPt / 28.3465).toFixed(1)}×${(cyPt / 28.3465).toFixed(1)}cm`);
                flow.push({ kind: 'image', heightPt: cyPt + base });
                continue;
            }
            if (text === '' && lineRule === 'exact' && lineAttr > 0) {
                outline.push(`留白 ${(lineAttr / TWIPS_PER_CM).toFixed(2)}cm`);
                flow.push({ kind: 'spacer', heightPt: lineAttr / 20 });
                continue;
            }
            // 公式段落（块级 m:oMathPara，或整段只有行内公式）要在「空段落」判定之前拦下来，
            // 否则一行公式只会被算成 12pt 的空白。分行数按分数/根式/大运算符的个数估，
            // 这些结构自带上下两层，比单行高。
            if (mathNodes.length && (blockMath || text === '')) {
                const tall = descendants(node, 'm:f').length + descendants(node, 'm:rad').length + descendants(node, 'm:nary').length;
                const lines = Math.min(3, 1 + Math.min(2, tall));
                outline.push(`${blockMath ? '公式' : '公式（行内）'} ${mathPreview(node)}`);
                flow.push({ kind: 'formula', heightPt: lines * size * 1.6 + base });
                continue;
            }
            if (info.kind === 'code') {
                stats.codeBlocks += 1;
                const codeLines = breaks + 1;
                outline.push(`代码 ${codeLines} 行`);
                // 代码不做折行：单行过宽就是真溢出，必须报出来
                let longest = 0;
                let current = 0;
                for (const chunk of chunks) {
                    if (chunk.kind === 'break') { current = 0; continue; }
                    if (chunk.kind !== 'text') continue;
                    current += chunk.value.length;
                    longest = Math.max(longest, current);
                }
                const limitChars = Math.floor(textWidth(indentLeft) / (size * 0.55));
                if (longest > limitChars) warnings.push(`代码块有 ${longest} 字符的长行，超出 ${Math.round(textWidth(indentLeft))}pt 版心，建议换行或缩短`);
                flow.push({ kind: 'code', heightPt: codeLines * size * factor + base });
                continue;
            }
            if (info.kind === 'quote') {
                stats.quotes += 1;
                outline.push(`引用 ${text}`);
                flow.push({ kind: 'quote', heightPt: wrapLines(info.indentRight ?? 0) * size * factor + base });
                continue;
            }
            if (info.kind === 'caption') {
                outline.push(`题注 ${text}`);
                flow.push({ kind: 'caption', heightPt: wrapLines() * size * factor + base });
                continue;
            }
            if (info.kind === 'subtitle') {
                outline.push(`副标题 ${text}`);
                flow.push({ kind: 'subtitle', heightPt: wrapLines() * size * factor + base });
                continue;
            }
            if (numId > 0 || info.kind === 'list') {
                stats.listItems += 1;
                if (!flow.length || flow[flow.length - 1].kind !== 'list') stats.lists += 1;
                outline.push(`列表 ${text}`);
                flow.push({ kind: 'list', heightPt: wrapLines() * size * factor + base });
                continue;
            }
            if (text === '') {
                flow.push({ kind: 'empty', heightPt: Math.max(1, size * factor + base) });
                continue;
            }
            stats.bodyParagraphs += 1;
            outline.push(`正文 ${[...text].length} 字${mathNodes.length ? `（含 ${mathNodes.length} 个行内公式）` : ''}`);
            if ([...text].length > 500) {
                warnings.push(`第 ${stats.bodyParagraphs} 段正文 ${[...text].length} 字未分段（超过 500 字），建议拆成多段`);
            }
            flow.push({ kind: 'body', heightPt: wrapLines() * size * factor + base });
            continue;
        }
        if (node.name === 'w:tbl') {
            stats.tables += 1;
            const grid = children(node, 'w:tblGrid')[0];
            const cols = grid ? children(grid, 'w:gridCol').map((col) => asNumber(attr(col, 'w:w'), 0)) : [];
            const rows = children(node, 'w:tr');
            const totalTw = cols.reduce((a, b) => a + b, 0);
            outline.push(`表格 ${rows.length}×${cols.length}`);
            if (totalTw > contentTw) {
                warnings.push(`第 ${stats.tables} 个表格列宽合计 ${(totalTw / TWIPS_PER_CM).toFixed(2)}cm 超出版心 ${(contentTw / TWIPS_PER_CM).toFixed(2)}cm，Word 会挤压列宽`);
            }
            if (cols.length > 8) {
                warnings.push(`第 ${stats.tables} 个表格有 ${cols.length} 列，${page.label} 版心下每列约 ${(contentTw / TWIPS_PER_CM / cols.length).toFixed(2)}cm，过窄`);
            }
            let height = 0;
            rows.forEach((row) => {
                const cells = children(row, 'w:tc');
                const isHeader = !!descendant(row, 'w:tblHeader');
                let rowLines = 1;
                // 合并单元格的宽度是它吃掉的各列之和，按单元格序号取宽度会越数越偏
                let colIndex = 0;
                cells.forEach((cell) => {
                    const cellPr = children(cell, 'w:tcPr')[0];
                    const span = Math.max(1, asNumber(attr(cellPr ? children(cellPr, 'w:gridSpan')[0] : undefined, 'w:val'), 1));
                    let widthTw = 0;
                    for (let i = 0; i < span; i += 1) widthTw += cols[colIndex + i] ?? 1200;
                    const widthPt = Math.max(24, widthTw / 20 - 10);
                    rowLines = Math.max(rowLines, estimateLines(cellText(cell), widthPt, STYLE_INFO.TableText.size));
                    colIndex += span;
                });
                height += Math.max(isHeader ? 22 : 20, rowLines * STYLE_INFO.TableText.size * 1.35 + 6);
            });
            flow.push({ kind: 'table', heightPt: height + 4 });
            continue;
        }
    }

    stats.images = descendants(body, 'a:blip').length;
    const allText = descendants(body, 'w:t').map((node) => textOf(node)).join('');
    stats.chars = [...allText.replace(/\s/g, '')].length;
    const cjk = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uac00-\ud7a3]/g;
    const cjkCount = (allText.match(cjk) ?? []).length;
    const latinCount = (allText.replace(cjk, ' ').match(/[A-Za-z0-9]+(?:['\-_.][A-Za-z0-9]+)*/g) ?? []).length;
    stats.words = cjkCount + latinCount;

    const usable = (page.heightTw - page.margin.top - page.margin.bottom) / 20;
    let pages = 1;
    let cursor = 0;
    for (const item of flow) {
        if (item.kind === 'pageBreak') { pages += 1; cursor = 0; continue; }
        const height = Math.max(0, item.heightPt);
        if (cursor > 0 && cursor + height > usable) { pages += 1; cursor = 0; }
        // 目录的预渲染页码来源：标题落在第几页
        if (item.headingIndex !== undefined) headingPages[item.headingIndex] = pages;
        cursor += height;
    }
    stats.pagesEstimate = pages;

    if (tocOutlineAt >= 0) {
        outline[tocOutlineAt] = tocEntries.length
            ? `目录（${tocDirty ? '打开时自动更新' : '预渲染'} ${tocEntries.length} 条${tocDirty ? '' : '，按 F9 可刷新'}）`
            : '目录（文档里没有 H1–H3 标题，只有占位提示）';
    }

    if (stats.paragraphs === 0 && stats.tables === 0) warnings.push('文档正文为空，只有节属性');
    if (!stats.headings && tocSeen) warnings.push('插入了目录域，但文档里没有任何 H1–H3 标题，目录会更新为空');
    if (stats.headings && headingLevels[0] !== 1) warnings.push(`文档第一个标题是 H${headingLevels[0]}，建议以 H1 开头以保持层级完整`);

    return { stats, outline, warnings, page, contentTw, headingPages, tocEntries };
}

// ---------------------------------------------------------------- builder

/** 表格块的统一形状：create() 的 builder.table 与 edit() 的 append 共用一份默认值。 */
function tableBlock(spec = {}) {
    return {
        type: 'table',
        columns: spec.columns,
        rows: spec.rows,
        widths: spec.widths,
        style: spec.style,
        caption: spec.caption ? asString(spec.caption) : '',
        // 表题默认在表格上方（中文排版惯例），要放下方显式写 captionPosition:'below'
        captionPosition: asString(spec.captionPosition, 'above').toLowerCase() === 'below' ? 'below' : 'above',
        headerRepeat: spec.headerRepeat !== false,
        align: asString(spec.align, ''),
    };
}

/** 统一的包内上下文：所有部件生成都读它，保证 rels/媒体/页面参数只有一份真相。 */
function createContext({ theme, page, footer, header, title, author, subject, created }) {
    const fonts = theme.fonts;
    const ctx = {
        theme,
        ink: paperInk(theme),
        fonts,
        page,
        title: title ?? '未命名文档',
        author: author ?? 'DSH Office Mode',
        subject: subject ?? '',
        created: created ?? new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'),
        blocks: [],
        media: [],
        warnings: [],
        relFooter: null,
        relHeader: null,
        contentTw: page.widthTw - page.margin.left - page.margin.right,
        drawingSeq: 0,
        relSeq: 3,
        footerTemplate: footer ?? null,
        headerTemplate: header ?? null,
    };
    if (ctx.footerTemplate) { ctx.relSeq += 1; ctx.relFooter = `rId${ctx.relSeq}`; }
    if (ctx.headerTemplate) { ctx.relSeq += 1; ctx.relHeader = `rId${ctx.relSeq}`; }
    return ctx;
}

export function create(spec = {}, env = {}) {
    if (typeof env.writeFile !== 'function') throw new Error('word.create: 需要 createEnv() 提供的 env（含 writeFile）');
    const options = spec ?? {};
    const picked = pickTheme(options.theme, env);
    const theme = picked.theme;
    if (picked.fellBack) {
        env.note?.(`未知主题 “${picked.requested}”，已回落到 ${theme.id}；可用：${THEMES.map((item) => item.id).join(' / ')}`);
    }
    if (theme.dark) env.note?.(`主题 ${theme.id} 是深色主题，Word 落在白纸上：主色保留，文字色已换成深墨以保证可读性`);
    const page = resolvePage(options);
    const ctx = createContext({
        theme,
        page,
        footer: options.footer === undefined ? '第 X 页' : (options.footer === false || options.footer === null || options.footer === '' ? null : asString(options.footer)),
        header: options.header ? asString(options.header) : null,
        title: options.title,
        author: options.author,
        subject: options.subject,
        created: options.created,
    });

    let cache = null;

    const builder = {
        /** 封面式大标题；`opts.subtitle` 会额外落一行副标题。 */
        title(text, opts = {}) {
            ctx.blocks.push({ type: 'title', runs: normalizeRuns(text), subtitle: opts.subtitle ? asString(opts.subtitle) : '' });
            return this;
        },
        heading(text, level = 1, opts = {}) {
            const safe = Math.max(1, Math.min(3, Math.round(asNumber(level, 1))));
            ctx.blocks.push({ type: 'heading', level: safe, runs: normalizeRuns(text), opts });
            return this;
        },
        para(text, opts = {}) {
            ctx.blocks.push({ type: 'para', runs: normalizeRuns(text), opts });
            return this;
        },
        bullets(items, opts = {}) {
            ctx.blocks.push({ type: 'list', ordered: false, items: asArray(items).map(normalizeRuns), opts });
            return this;
        },
        steps(items, opts = {}) {
            ctx.blocks.push({ type: 'list', ordered: true, items: asArray(items).map(normalizeRuns), opts });
            return this;
        },
        quote(text, opts = {}) {
            ctx.blocks.push({ type: 'quote', runs: normalizeRuns(text), opts });
            return this;
        },
        code(text) {
            const normalized = asString(text).replace(/\t/g, '    ').replace(/\r\n?/g, '\n');
            ctx.blocks.push({ type: 'code', text: normalized });
            return this;
        },
        table(spec2 = {}) {
            ctx.blocks.push(tableBlock(spec2));
            return this;
        },
        /**
         * 块级公式。opts: {align:'center'|'left', number:'（1）', size}。
         * LaTeX 在这里就解析成 OMML 存进 block：render() 必须保持纯函数，
         * 不能在渲染时再解析一次（那会让两次 render() 的告警数不一致）。
         */
        formula(latex, opts = {}) {
            const source = asString(latex);
            if (source.trim() === '') {
                ctx.warnings.push('formula(): LaTeX 为空，已跳过该公式');
                return this;
            }
            const size = asNumber(opts.size, 0);
            ctx.blocks.push({
                type: 'formula',
                align: asString(opts.align, 'center').toLowerCase() === 'left' ? 'left' : 'center',
                number: opts.number === undefined || opts.number === null || opts.number === false ? '' : asString(opts.number),
                // 字号在解析时就烘进 OMML 的 run 里，渲染阶段只负责摆位置
                omml: mathToOmml(source, ctx, { size: size > 0 ? size : 0 }),
            });
            return this;
        },
        image(imagePath, opts = {}) {
            const target = asString(imagePath);
            if (target === '') { ctx.warnings.push('image(): 路径为空，已跳过'); return this; }
            let bytes;
            try {
                bytes = env.readFile(target);
            } catch {
                ctx.warnings.push(`图片 ${target} 读不到，已跳过（Word 里不会出现这张图）`);
                return this;
            }
            const info = sniffImage(bytes);
            if (!info) {
                ctx.warnings.push(`图片 ${target} 不是 PNG/JPEG，OOXML 只嵌入这两种，已跳过`);
                return this;
            }
            if (!info.width || !info.height) ctx.warnings.push(`无法读取 ${target} 的像素尺寸，按 4:3 估算显示尺寸`);
            const natural = info.width && info.height ? (info.width / 96) * 2.54 : 12;
            const naturalHeight = info.width && info.height ? (info.height / 96) * 2.54 : 9;
            const contentCm = ctx.contentTw / TWIPS_PER_CM;
            let widthCm = asNumber(opts.widthCm, 0);
            if (!(widthCm > 0)) widthCm = Math.min(natural, contentCm);
            if (widthCm > contentCm) {
                ctx.warnings.push(`图片 ${target} 宽度 ${widthCm.toFixed(2)}cm 超出版心 ${contentCm.toFixed(2)}cm，已等比缩到版心宽度`);
                widthCm = contentCm;
            }
            const heightCm = natural > 0 ? (widthCm * naturalHeight) / natural : widthCm * 0.75;
            ctx.drawingSeq += 1;
            ctx.relSeq += 1;
            const fileName = `image${ctx.media.length + 1}.${info.ext}`;
            const relId = `rId${ctx.relSeq}`;
            ctx.media.push({ fileName, ext: info.ext, data: bytes, relId });
            ctx.blocks.push({
                type: 'image',
                fileName,
                relId,
                drawingId: ctx.drawingSeq,
                widthCm,
                heightCm,
                caption: opts.caption ? asString(opts.caption) : '',
                alt: opts.alt ? asString(opts.alt) : asString(opts.caption, target),
            });
            env.note?.(`嵌入图片 ${target} → word/media/${fileName}（${widthCm.toFixed(2)}×${heightCm.toFixed(2)}cm）`);
            return this;
        },
        pageBreak() {
            ctx.blocks.push({ type: 'pageBreak' });
            return this;
        },
        spacer(cm = 0.4) {
            ctx.blocks.push({ type: 'spacer', cm: asNumber(cm, 0.4) });
            return this;
        },
        toc() {
            ctx.blocks.push({ type: 'toc' });
            return this;
        },
        /** 纯函数：只产出字节，不碰磁盘。 */
        render() {
            if (!cache) cache = renderPackage(ctx);
            return cache.bytes;
        },
        save(target) {
            const bytes = this.render();
            const name = ensureExtension(
                asString(target) || asString(options.path) || `${slugify(options.title ?? 'word-document')}.docx`,
                '.docx',
            );
            const entry = env.writeFile(name, bytes);
            return {
                ok: true,
                format: 'docx',
                path: entry.path,
                bytes: entry.bytes,
                theme: theme.id,
                stats: cache.stats,
                outline: cache.outline,
                warnings: cache.warnings,
            };
        },
    };

    function renderPackage(context) {
        // 目录页码要用真实的分页估算，所以带目录时跑两遍：
        //   第一遍：目录条目已就位但页码是占位值 → analyzeDocument 报出每个标题落在第几页；
        //   第二遍：把估算页码写进目录域的缓存结果。
        // 两遍的条目与文字完全相同，目录自身占的高度不变，因此第一遍估出的页码在
        // 第二遍里仍然成立（页码数字的宽度差异可以忽略）。
        let entries = buildParts(context);
        let documentPart = entries.find((entry) => entry.name === 'word/document.xml').data;
        let analysis = analyzeDocument(documentPart);
        if (context.blocks.some((block) => block.type === 'toc')) {
            context.tocPages = analysis.headingPages ?? [];
            entries = buildParts(context);
            documentPart = entries.find((entry) => entry.name === 'word/document.xml').data;
            analysis = analyzeDocument(documentPart);
        }
        const bytes = zip(entries);
        const warnings = [...context.warnings, ...analysis.warnings];
        return { bytes, stats: analysis.stats, outline: analysis.outline, warnings };
    }

    return builder;
}

// ---------------------------------------------------------------- read

function decodePart(parts, name) {
    const data = parts.get(name);
    if (!data) return undefined;
    return new TextDecoder('utf-8').decode(data);
}

export function read(path, env = {}) {
    const absolute = env.resolve ? env.resolve(path) : path;
    const display = env.root ? displayPath(env.root, absolute) : path;
    if (env.exists && !env.exists(path)) throw new Error(`word.read: 文件不存在：${display}`);
    const buffer = env.readFile(path);
    const parts = unzip(buffer);
    const documentXml = decodePart(parts, 'word/document.xml');
    if (documentXml === undefined) throw new Error(`word.read: ${display} 里没有 word/document.xml，不是 .docx`);
    const themeId = detectTheme(decodePart(parts, 'word/styles.xml'));
    const analysis = analyzeDocument(documentXml);
    return {
        ok: true,
        format: 'docx',
        path: display,
        bytes: env.stat ? (env.stat(path)?.size ?? buffer.length) : buffer.length,
        theme: themeId,
        stats: analysis.stats,
        outline: analysis.outline,
        warnings: analysis.warnings,
    };
}

// ---------------------------------------------------------------- edit

const PARA_RE = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
const TEXT_RE = /<w:t(?:\s[^>]*)?\/>|<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g;

/** 把 `<w:t>` 的原标签还原成带新文本的标签（保留原属性，例如 xml:space）。 */
function rewriteTextTag(original, text) {
    const open = original.slice(0, original.indexOf('>') + 1);
    const match = /^<w:t((?:\s[^>]*?)?)\s*\/?>$/.exec(open);
    let attrs = match ? match[1] : '';
    if (!/xml:space/.test(attrs)) attrs += ' xml:space="preserve"';
    return `<w:t${attrs}>${escapeXml(text)}</w:t>`;
}

function replaceInXml(xml, find, replace, all) {
    let hits = 0;
    const out = xml.replace(PARA_RE, (block) => {
        if (all === false && hits > 0) return block;
        const records = [];
        let cursor = 0;
        TEXT_RE.lastIndex = 0;
        let match;
        while ((match = TEXT_RE.exec(block)) !== null) {
            const raw = match[0];
            const inner = /\/>$/.test(raw) ? '' : raw.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '');
            const text = decodeXmlText(inner);
            records.push({ index: match.index, raw, text, start: cursor, end: cursor + text.length });
            cursor += text.length;
        }
        if (!records.length) return block;
        let whole = records.map((record) => record.text).join('');
        if (!whole.includes(find)) return block;

        const edits = [];
        if (all) {
            let from = 0;
            while (from <= whole.length) {
                const at = whole.indexOf(find, from);
                if (at === -1) break;
                edits.push({ at, end: at + find.length });
                from = at + Math.max(1, find.length);
            }
        } else {
            const at = whole.indexOf(find);
            if (at !== -1) edits.push({ at, end: at + find.length });
        }
        if (!edits.length) return block;

        for (const edit of edits.reverse()) {
            // 从后往前替换，前面的偏移量才不会失效
            const firstRun = records.findIndex((record) => edit.at >= record.start && edit.at < record.end);
            const lastRun = records.findIndex((record) => edit.end - 1 >= record.start && edit.end - 1 < record.end);
            if (firstRun === -1 || lastRun === -1) continue;
            if (firstRun === lastRun) {
                const record = records[firstRun];
                const local = edit.at - record.start;
                const localEnd = edit.end - record.start;
                record.text = record.text.slice(0, local) + replace + record.text.slice(localEnd);
            } else {
                const head = records[firstRun];
                const tail = records[lastRun];
                head.text = head.text.slice(0, edit.at - head.start) + replace;
                tail.text = tail.text.slice(edit.end - tail.start);
                for (let i = firstRun + 1; i < lastRun; i += 1) records[i].text = '';
            }
            whole = records.map((record) => record.text).join('');
            hits += 1;
        }

        let rebuilt = '';
        let at = 0;
        for (const record of records) {
            // at 必须推进到「原标签的结尾」，而不是下一个标签的起点：
            // 段落在最后一个 w:t 之后还有 </w:r></w:p>，用 block.length 会把它整段吃掉
            rebuilt += block.slice(at, record.index);
            rebuilt += rewriteTextTag(record.raw, record.text);
            at = record.index + record.raw.length;
        }
        rebuilt += block.slice(at);
        return rebuilt;
    });
    return { xml: out, hits };
}

/** 从原始 XML 片段里取段落文本：find/replace 与 setStyle 都要按「解码后的真实文字」匹配。 */
function paragraphText(block) {
    const out = [];
    TEXT_RE.lastIndex = 0;
    let match;
    while ((match = TEXT_RE.exec(block)) !== null) {
        const raw = match[0];
        const inner = /\/>$/.test(raw) ? '' : raw.slice(raw.indexOf('>') + 1, raw.lastIndexOf('</w:t>'));
        out.push(decodeXmlText(inner));
    }
    return out.join('');
}

function setParagraphStyle(block, styleId) {
    if (/<w:pPr(?:\s[^>]*)?>/.test(block)) {
        if (/<w:pStyle\s[^>]*\/>/.test(block)) {
            return block.replace(/<w:pStyle\s[^>]*\/>/, `<w:pStyle w:val="${styleId}"/>`);
        }
        return block.replace(/(<w:pPr(?:\s[^>]*)?>)/, `$1<w:pStyle w:val="${styleId}"/>`);
    }
    return block.replace(/(<w:p(?:\s[^>]*)?>)/, `$1<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>`);
}

function insertIntoBody(documentXml, xml) {
    // 新内容必须插在节属性之前，否则 sectPr 被挤到中间会导致页面设置丢失
    const sectAt = documentXml.lastIndexOf('<w:sectPr');
    if (sectAt > -1) return `${documentXml.slice(0, sectAt)}${xml}${documentXml.slice(sectAt)}`;
    return documentXml.replace('</w:body>', `${xml}</w:body>`);
}

function appendBlocks(ops, page, theme, hasNumbering) {
    const ctx = createContext({ theme, page });
    const warnings = [];
    for (const op of ops) {
        const spec2 = op.append ?? {};
        const type = asString(spec2.type, 'para');
        if (type === 'para') ctx.blocks.push({ type: 'para', runs: normalizeRuns(spec2.text), opts: spec2.opts ?? {} });
        else if (type === 'heading') ctx.blocks.push({ type: 'heading', level: Math.max(1, Math.min(3, asNumber(spec2.level, 1))), runs: normalizeRuns(spec2.text), opts: spec2.opts ?? {} });
        else if (type === 'bullets') ctx.blocks.push({ type: 'list', ordered: !!spec2.ordered, items: asArray(spec2.items).map(normalizeRuns), opts: spec2.opts ?? {} });
        else if (type === 'table') ctx.blocks.push(tableBlock(spec2));
        else if (type === 'formula') {
            const source = asString(spec2.latex);
            if (source.trim() === '') return { xml: '', ok: false, reason: 'append.latex 为空，无法追加公式' };
            const size = asNumber(spec2.opts?.size ?? spec2.size, 0);
            ctx.blocks.push({
                type: 'formula',
                align: asString(spec2.opts?.align ?? spec2.align, 'center').toLowerCase() === 'left' ? 'left' : 'center',
                number: spec2.opts?.number ?? spec2.number ?? '',
                omml: mathToOmml(source, ctx, { size: size > 0 ? size : 0 }),
            });
        } else return { xml: '', ok: false, reason: `append.type 只支持 para/heading/bullets/table/formula，收到 ${type}` };
    }
    let xml = ctx.blocks.map((block) => renderBlock(block, ctx)).join('');
    // 公式解析产生的告警也要出现在 edit 报告里，否则「公式里有个不认识的命令」会被静默吞掉
    warnings.push(...ctx.warnings);
    if (!hasNumbering && /<w:numPr>/.test(xml)) {
        // 包里没有 numbering.xml 时，列表退化成手动符号，至少内容不丢
        xml = xml.replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, '');
        warnings.push('目标文档没有 word/numbering.xml，追加的列表已退化为普通段落（符号丢失）');
    }
    return { xml, ok: true, warnings };
}

export function edit(path, ops, env = {}) {
    const absolute = env.resolve ? env.resolve(path) : path;
    const display = env.root ? displayPath(env.root, absolute) : path;
    if (env.exists && !env.exists(path)) throw new Error(`word.edit: 文件不存在：${display}`);
    const parts = unzip(env.readFile(path));
    const documentXml = decodePart(parts, 'word/document.xml');
    if (documentXml === undefined) throw new Error(`word.edit: ${display} 里没有 word/document.xml，不是 .docx`);

    const themeId = detectTheme(decodePart(parts, 'word/styles.xml'));
    const { theme } = resolveTheme(themeId);
    const body = descendant(parseXml(documentXml), 'w:body');
    const page = pageFromXml(body ?? { children: [] });

    let xml = documentXml;
    const applied = [];
    const skipped = [];
    const changes = [];

    for (const [index, op] of asArray(ops).entries()) {
        const where = `ops[${index}]`;
        if (!op || typeof op !== 'object') { skipped.push({ op: where, reason: '操作不是对象' }); continue; }

        if (op.find !== undefined) {
            const find = asString(op.find);
            const replace = asString(op.replace);
            if (find === '') { skipped.push({ op: where, reason: 'find 为空字符串' }); continue; }
            const result = replaceInXml(xml, find, replace, op.all !== false);
            xml = result.xml;
            if (!result.hits) { skipped.push({ op: where, reason: `文档里找不到 “${find}”` }); continue; }
            applied.push({ op: 'find/replace', find, replace, count: result.hits, all: op.all !== false });
            changes.push({ find, replace, count: result.hits });
            continue;
        }

        if (op.setStyle) {
            const styleId = asString(op.setStyle.style);
            if (!PARAGRAPH_STYLES.includes(styleId)) {
                skipped.push({ op: where, reason: `style 必须是 ${PARAGRAPH_STYLES.join('/')} 之一，收到 “${styleId}”` });
                continue;
            }
            const matcher = op.setStyle.match;
            let test;
            if (matcher instanceof RegExp) {
                // 去掉 g/y：带状态的 test() 会在段落之间来回翻转，导致行为不可预测
                const re = new RegExp(matcher.source, matcher.flags.replace(/[gy]/g, ''));
                test = (text) => re.test(text);
            } else if (typeof matcher === 'string' && matcher !== '') {
                const re = new RegExp(matcher);
                test = (text) => re.test(text);
            } else { skipped.push({ op: where, reason: 'setStyle.match 必须是正则或字符串' }); continue; }
            let count = 0;
            xml = xml.replace(PARA_RE, (block) => {
                const text = paragraphText(block);
                if (!text || !test(text)) return block;
                const from = /<w:pStyle\s[^>]*w:val="([^"]*)"/.exec(block)?.[1] ?? 'Normal';
                count += 1;
                changes.push({ setStyle: `${from}→${styleId}`, text: text.slice(0, 40) });
                return setParagraphStyle(block, styleId);
            });
            if (!count) { skipped.push({ op: where, reason: 'setStyle.match 没有匹配到任何段落' }); continue; }
            applied.push({ op: 'setStyle', style: styleId, count });
            continue;
        }

        if (op.append) {
            const result = appendBlocks([op], page, theme, parts.has('word/numbering.xml'));
            if (!result.ok) { skipped.push({ op: where, reason: result.reason }); continue; }
            xml = insertIntoBody(xml, result.xml);
            for (const message of result.warnings) skipped.push({ op: where, reason: message });
            applied.push({ op: 'append', type: asString(op.append.type, 'para') });
            changes.push({ append: asString(op.append.type, 'para'), text: asString(op.append.text ?? (asArray(op.append.items)[0] ?? '')).slice(0, 40) });
            continue;
        }

        if (op.setFooter !== undefined) {
            const template = asString(op.setFooter);
            if (template === '') { skipped.push({ op: where, reason: 'setFooter 传空字符串没有意义；要显示空白页脚请传一个空格' }); continue; }
            const ctx = createContext({ theme, page, footer: template });
            const relsXml = decodePart(parts, 'word/_rels/document.xml.rels') ?? '';
            if (relsXml === '') { skipped.push({ op: where, reason: '包里没有 word/_rels/document.xml.rels，无法挂载页脚部件' }); continue; }
            // 关系 id 必须在包内唯一：已有 footer1.xml 的关系就复用它的 id，否则顺延一个新的
            const existing = /<Relationship Id="([^"]+)"[^>]*Target="footer1\.xml"[^>]*\/>/.exec(relsXml);
            let relId = existing?.[1];
            let nextRels = relsXml;
            if (!relId) {
                relId = `rId${ctx.relSeq}`;
                while (new RegExp(`Id="${relId}"`).test(relsXml)) {
                    ctx.relSeq += 1;
                    relId = `rId${ctx.relSeq}`;
                }
                nextRels = relsXml.replace('</Relationships>', `<Relationship Id="${relId}" Type="${REL_TYPE.footer}" Target="footer1.xml"/></Relationships>`);
                parts.set('word/_rels/document.xml.rels', new TextEncoder().encode(nextRels));
            }
            ctx.relFooter = relId;
            parts.set('word/footer1.xml', new TextEncoder().encode(footerXml(ctx)));

            const contentTypes = decodePart(parts, '[Content_Types].xml') ?? '';
            if (contentTypes && !/footer1\.xml/.test(contentTypes)) {
                const nextTypes = contentTypes.replace('</Types>', `<Override PartName="/word/footer1.xml" ContentType="${CT_PART.footer}"/></Types>`);
                parts.set('[Content_Types].xml', new TextEncoder().encode(nextTypes));
            }
            xml = xml.replace(/<w:footerReference[^>]*\/>/g, '');
            if (/<w:sectPr/.test(xml)) {
                xml = xml.replace(/<w:sectPr(\s[^>]*)?>/, (whole) => `${whole}<w:footerReference w:type="default" r:id="${relId}"/>`);
            }
            applied.push({ op: 'setFooter', text: template, relId });
            changes.push({ setFooter: template });
            continue;
        }

        skipped.push({ op: where, reason: `未知操作，支持的键：find/replace、setStyle、append、setFooter（收到 ${Object.keys(op).join(',')}）` });
    }

    // 追加了公式就必须保证根元素声明了 m 前缀：外部来的 .docx 多半没声明过，
    // 少了它 Word 会把整份文档判成非法（而不是只丢公式）
    if (/<m:oMath/.test(xml) && !/xmlns:m=/.test(xml)) {
        xml = xml.replace(/<w:document(?=[\s>])/, `<w:document xmlns:m="${M_NS}"`);
    }

    parts.set('word/document.xml', new TextEncoder().encode(xml));
    const bytes = zip(parts);
    const entry = env.writeFile(path, bytes);
    const analysis = analyzeDocument(xml);

    return {
        ok: true,
        format: 'docx',
        path: entry.path,
        bytes: entry.bytes,
        theme: themeId,
        stats: analysis.stats,
        outline: analysis.outline,
        warnings: analysis.warnings,
        applied,
        skipped,
        changes,
    };
}

// ---------------------------------------------------------------- meta

export const meta = {
    id: 'word',
    name: 'Word 文档',
    ext: '.docx',
    summary: '生成/读取/批量改写 .docx：封面、标题层级、列表、引用、代码块、表格（合并单元格/题注）、图片、OMML 数学公式、页码页脚、目录域（条目与页码预渲染，打开零弹窗）；默认素色（不加背景底纹）',
    // office_help 的默认层用这一份（提示词预算；全文层留给 detail:true）。
    // 签名与参数名必须与 methods 一致 —— 分层省的是解释与举例，不是接口。
    brief: {
        create: [
            'create({title, theme, path, page, orientation, margins, header, footer, author, subject}) → builder',
            '  builder.title(text, {subtitle}) / builder.heading(text, level=1, opts) / builder.para(text, opts)',
            '  builder.bullets(items, opts) / builder.steps(items, opts) / builder.quote(text, opts) / builder.code(text)',
            '  builder.table({columns, rows, widths, style, caption, captionPosition, headerRepeat, align})',
            "    columns: ['季度', {title:'营收', width:5, align:'right'}]",
            "    rows 单元格: '文本' 或 {text, colspan, rowspan, align, bold, italic, fill, color, size}",
            '  builder.para([{math:\'x^2\'}]) 行内公式 / builder.formula(latex, {align, number, size}) 块级公式',
            '  builder.image(path, {widthCm, caption, alt})   内嵌 PNG/JPEG，超版心自动等比缩放',
            '  builder.pageBreak() / builder.spacer(cm) / builder.toc()   目录（H1–H3 条目与页码预渲染进 TOC 域）',
            '  opts: {bold, italic, underline, color, size, font, align, indent, spaceBefore, spaceAfter, lineSpacing}',
            '  builder.render() → Uint8Array / builder.save(path?) → report',
        ],
        read: ['read(path, env) → report（stats/outline/warnings 来自真实解析；stats.formulas 为 m:oMath 个数）'],
        edit: [
            'edit(path, ops, env) → report + {applied, skipped, changes}，一次调用只写一次盘',
            "  [{find, replace, all}]  跨 run 文本替换（all 默认 true）",
            "  [{setStyle: {match: /^旧标题/, style: 'Heading2'}}] / [{setFooter: '第 X 页 / 共 Y 页'}]",
            "  [{append: {type: 'para'|'heading'|'bullets'|'table'|'formula', ...}}]",
        ],
    },
    methods: {
        create: [
            'create({title, theme, path, page, orientation, margins, header, footer, author, subject}) → builder',
            '  builder.title(text, {subtitle})              封面式大标题（主题色 + 下边框）',
            '  builder.heading(text, level=1, opts)         标题 1–3 级，带大纲级别（导航窗格/目录可用）',
            '  builder.para(text, opts)                     正文；text 可为 string 或 [{text,bold,italic,color,size}]',
            '  builder.para([{math:\'x^2\'}])               行内公式与普通 run 混排（LaTeX 子集 → OMML）',
            '  builder.bullets(items, opts)                 无序列表（1.5 倍行距，自动缩进分级）',
            '  builder.steps(items, opts)                   有序列表（1. / 1.1. / 1.1.1.）',
            '  builder.quote(text, opts)                    引用块（左侧强调色竖线；主题有底色时才有浅底）',
            '  builder.code(text)                           等宽代码块（按 \\n 拆分，不折行）',
            '  builder.table({columns, rows, widths, style, caption, captionPosition, headerRepeat, align})',
            '    columns: [\'季度\', {title:\'营收\', width:5, align:\'right\'}]',
            '    rows 单元格: \'文本\' 或 {text, colspan, rowspan, align, bold, italic, fill, color, size}',
            '    caption 用 Caption 题注段落（默认表上方，captionPosition:\'below\' 放下方）；headerRepeat:false 不重复表头',
            '  builder.formula(latex, {align:\'center\'|\'left\', number:\'（1）\', size})  块级公式（OMML）',
            '    LaTeX 子集: ^ _ \\frac \\sqrt[n] \\left( \\right) \\sum \\prod \\int \\oint \\iint \\sin \\lim \\alpha \\times \\text{中文} \\mathrm \\mathbf \\, \\quad 等',
            '  builder.image(path, {widthCm, caption, alt}) 内嵌 PNG/JPEG，超版心自动等比缩放',
            '  builder.pageBreak() / builder.spacer(cm) / builder.toc()',
            '    toc() 生成 TOC 域，并把 H1–H3 的条目与页码**预渲染**进域的缓存结果：打开即见完整目录、',
            '    条目可点击跳转（标题上有书签），且不带任何「打开时更新域」标记 —— 那类标记会让 Word 弹',
            '    「该文档包含的域可能引用了其他文件」。页码来自引擎的分页估算，按 F9 可让 Word 重算成精确值。',
            '  opts: {bold, italic, underline, color, size, font, align, indent, spaceBefore, spaceAfter, lineSpacing}',
            '  builder.render() → Uint8Array（纯函数，不写盘）；builder.save(path?) → report',
        ],
        read: ['read(path, env) → report（stats/outline/warnings 全部来自真实解析 word/document.xml；stats.formulas 为 m:oMath 个数）'],
        edit: [
            'edit(path, ops, env) → report + {applied, skipped, changes}，一次调用只写一次盘',
            '  [{find, replace, all=true}]                   跨 run 文本替换（默认替换全部命中，all:false 只替换第一处）',
            "  [{setStyle: {match: /^旧标题/, style: 'Heading2'}}]",
            "  [{append: {type: 'para'|'heading'|'bullets'|'table'|'formula', ...}}]",
            "  [{append: {type: 'formula', latex: '\\\\frac{a}{b}', number: '（1）'}}]",
            "  [{setFooter: '第 X 页 / 共 Y 页'}]              X→PAGE 域，Y→NUMPAGES 域",
        ],
    },
};
