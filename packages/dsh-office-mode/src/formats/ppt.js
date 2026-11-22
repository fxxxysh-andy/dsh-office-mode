/**
 * PPT 格式引擎（.pptx / PresentationML）。
 *
 * 设计取舍（为什么这样做）：
 * 1. 所有内容都用独立文本框（`p:sp` + `p:cNvSpPr txBox="1"`），一个占位符都不用。
 *    `p:ph` 要在 slide / slideLayout / slideMaster 三处把 type 与 idx 配成一套，
 *    任何一处错位 PowerPoint 就报「需要修复」；独立文本框没有继承链，几何、字体、
 *    颜色、项目符号全部自带，代价是形状多几个，换来的是零继承风险。
 * 2. 版式只作为「归属」存在：8 个 slideLayout 各自只有 spTree 骨架 + clrMapOvr。
 *    这样做符合 OOXML 语义（layout 允许没有占位符），也避免 layout 与 slide 之间的
 *    几何互相打架。
 * 3. 表格不写 `a:tableStyleId`。引用一个包里不存在的样式 guid，PowerPoint 会走修复；
 *    改成逐格写 `a:tcPr` 的填充与边框，表格外观完全自持，也不再多一个 tableStyles 部件。
 * 4. 溢出风险用 kit.wrapText / estimateEm 在写盘前算：先按字号阶梯找出「放得下」的
 *    字号，找不到就带着真实数字（需要多少磅、可用多少磅）进 warnings。报告里的每个
 *    数字都来自这次估算，`a:normAutofit` 只是留给 PowerPoint 编辑时的兜底。
 * 5. theme1.xml 完整写足 clrScheme(12) / fontScheme / fmtScheme(fill,ln,effect,bg 各 3 项)。
 *    PowerPoint 的样式索引按 1-based 取用，fmtScheme 少一项就判定包损坏，不能省。
 * 6. 图片的尺寸 / 裁切 / 对齐全部由 layoutImage 一处算出来，渲染、stats、outline、
 *    warnings 四个出口读同一个结果：报告与画面不一致比没有报告更糟。
 *
 * @module dsh-office-mode/formats/ppt
 */
import { createHash } from 'node:crypto';
import { zip, unzip } from '../engine/zip.js';
import { parseXml, descendants, children as xmlChildren, textOf, normalizeSpace, decodeXmlEntities } from '../engine/xml.js';
import {
    EMU_PER_PT, cmToEmu, estimateEm, wrapText, asArray, asNumber, asString,
    normalizeColumns, cellType, createEnv, assertInsideRoot, displayPath, ensureExtension, slugify,
    basename, join,
} from '../engine/kit.js';
import { DEFAULT_THEME_ID, resolveTheme, toHex } from '../engine/theme.js';
import { chartSpaceXml, normalizeChartSpec, presentationChartFrame } from './chart.js';
import { latexToMath, looksLikeLatex, looksLikePlainFormula, mathPlain } from './ppt-math.js';

// ───────────────────────────────────────────────────────────────────────────
// 常量：OOXML 里写死的固定部分
// ───────────────────────────────────────────────────────────────────────────

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const NS_VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const NS_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const NS_DC = 'http://purl.org/dc/elements/1.1/';
const NS_DCTERMS = 'http://purl.org/dc/terms/';
const NS_DCMITYPE = 'http://purl.org/dc/dcmitype/';
const NS_XSI = 'http://www.w3.org/2001/XMLSchema-instance';
// 公式（a14:m）：markup-compatibility 的 AlternateContent 包装 + DrawingML 2010 的 a14
// + Office Math 的 m 三个命名空间都内联声明在节点上（与 PowerPoint 自己写的文件一致）
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const NS_MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

/** 关系类型。PPT 的关系类型比 Word/Excel 多，写错一个就是「需要修复」。 */
const RT = {
    officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
    slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
    slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
    notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
    theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
    image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    core: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    app: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
};

const CT = {
    presentation: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
    theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
    chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    core: 'application/vnd.openxmlformats-package.core-properties+xml',
    app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
};

/**
 * 8 种版式。`type` 用 ST_SlideLayoutType 里的保守取值（title/tx/tbl/secHead/titleOnly/blank），
 * 只作语义标注：本引擎不用占位符，PowerPoint 不会因为这个属性改变渲染。
 */
const LAYOUTS = [
    { key: 'cover', name: '封面', type: 'title' },
    { key: 'section', name: '章节', type: 'secHead' },
    { key: 'bullets', name: '要点', type: 'tx' },
    { key: 'table', name: '表格', type: 'tbl' },
    { key: 'quote', name: '引用', type: 'tx' },
    { key: 'statement', name: '陈述', type: 'titleOnly' },
    { key: 'image', name: '图片', type: 'picTx' },
    { key: 'closing', name: '结尾', type: 'titleOnly' },
];

const LAYOUT_KEYS = LAYOUTS.map((item) => item.key);

/**
 * 版式助手：一次调用生成一整页的常用汇报页（卡片网格、编号流程、对比双栏、
 * KPI 大数字、图文、时间线、图标网格、目录、图片墙、路线图、矩阵）。
 *
 * 为什么它们要自成一类 kind，而不是「往自由绘制页里塞形状」：
 * 1. stats.layouts / outline 里必须能看出这一页用的是哪种助手
 *    （「P3 cards：4 张卡片」），笼统报成 shape 等于没报；
 * 2. 助手的 items 通道与自由绘制页共用，所以 deck.cards() 之后接着
 *    deck.banner()/deck.shape() 会续在同一页 —— 参考稿式的整页就是
 *    「助手排版 + 元素级微调」拼出来的。
 *
 * 名单顺序 = read() 的识别优先级，见 HELPER_MARKERS。
 */
const HELPER_KINDS = [
    'cards', 'steps', 'compare', 'kpi', 'imageText', 'timeline', 'iconGrid',
    'agenda', 'gallery', 'roadmap', 'matrix',
];

/**
 * 允许在同一页继续落「叠加层 / 配图」的版式（第五十二轮 `新-6`）。
 *
 * 从前只有自由绘制页与助手页开放画布：`deck.bullets()` 之后写 `deck.banner()`
 * 会另起一页空白页，`deck.cards()` 之后写 `deck.image()` 也只能另起一页 ——
 * 「版式 + 配图」这种最常见的成稿方式反而要手工摆形状。
 * 现在内容型版式一律开放：叠加层画在版式之后（图层顺序 = 调用顺序）。
 * 封面 / 章节页 / 结束页仍然不开放 —— 它们是整页构图，塞东西进去只会压标题。
 */
const CANVAS_HOST_KINDS = new Set([
    'shape', ...HELPER_KINDS,
    'bullets', 'table', 'quote', 'statement', 'image', 'images', 'chart',
]);

/**
 * 反过来：不开放画布的版式（整页构图，塞东西进去只会压标题）。
 * read() 用它判断「基础版式页之后的自由绘制页」这条告警还成不成立。
 */
const CANVAS_CLOSED_KINDS = ['cover', 'section', 'closing'];

/**
 * 画布配图与版式共存时的默认分栏：图占版心右侧这一比例，宿主版式用剩下的左栏。
 * 只在调用方没给 `x/y/w/h` 时生效；给了显式摆位就完全按调用方给的画。
 */
const CANVAS_IMAGE_SPLIT_RATIO = 0.42;
/** 分栏时版式与配图之间的间隙（厘米）。 */
const CANVAS_IMAGE_SPLIT_GAP_CM = 0.6;

/** 画布尺寸（EMU）。16:9 与 4:3 是 SPEC 指定的两档。 */
const SIZES = {
    '16:9': { cx: 12192000, cy: 6858000, type: 'screen16x9' },
    '4:3': { cx: 9144000, cy: 6858000, type: 'screen4x3' },
};

/** clrMap 的 12 个槽位必须全给，缺一个 PowerPoint 就按修复处理。 */
const CLR_MAP_ATTRS = [
    'bg1="lt1"', 'tx1="dk1"', 'bg2="lt2"', 'tx2="dk2"',
    'accent1="accent1"', 'accent2="accent2"', 'accent3="accent3"',
    'accent4="accent4"', 'accent5="accent5"', 'accent6="accent6"',
    'hlink="hlink"', 'folHlink="folHlink"',
].join(' ');

/** 项目符号的层级表：缩进、悬挂、字符、相对正文字号的缩减量。 */
const LEVELS = [
    { marL: cmToEmu(0.52), indent: -cmToEmu(0.52), bullet: '•', drop: 0, gapPt: 10 },
    { marL: cmToEmu(1.30), indent: -cmToEmu(0.46), bullet: '–', drop: 3, gapPt: 6 },
    { marL: cmToEmu(2.06), indent: -cmToEmu(0.44), bullet: '·', drop: 5, gapPt: 5 },
    { marL: cmToEmu(2.72), indent: -cmToEmu(0.44), bullet: '·', drop: 6, gapPt: 4 },
];

const LINE_SPACING_BODY = 1.28;
const TABLE_MAR_H = cmToEmu(0.16);
const TABLE_MAR_V = cmToEmu(0.09);

/**
 * 图片版式的三档适配方式：
 * - contain：等比缩放进可用区域，留白（默认，也是本引擎一直以来的行为）；
 * - cover：铺满可用区域，多出来的边用 a:srcRect 裁掉（真正的「填充」）；
 * - natural：按 96 DPI 把原始像素折算成厘米，只在超出可用区域时才缩。
 * 非法值不抛异常，由 imageFit() 记一条 warning 后回落到 contain。
 */
const IMAGE_FITS = ['contain', 'cover', 'natural'];

/** 96 DPI 下 1 像素 = 9525 EMU，这是「原始像素 → 厘米」的唯一换算依据。 */
const EMU_PER_PX = 9525;

/** a:srcRect 是千分之一百分比（100000 = 100%），给满 100000 会被判非法，封顶 99999。 */
const SRC_RECT_MAX = 99999;

/**
 * 可自由绘制的预设几何（ST_ShapeType 的取值）。
 * 名单刻意只收「在 PowerPoint 里一定存在」的 preset：`prst` 写错一个字母，
 * PowerPoint 会走修复流程，而自研解析器与 LibreOffice 都看不出异常。
 * `line` 也在这里：deck.line 用 `prstGeom prst="line"` + flipH/flipV 表达任意方向的直线。
 */
const PRST_GEOMS = new Set([
    'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'hexagon',
    'chevron', 'pentagon', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'bentArrow',
    'curvedRightArrow', 'blockArc', 'donut', 'pie', 'teardrop', 'cloud', 'star5', 'heart',
    'plaque', 'frame', 'halfFrame', 'corner', 'line',
    'flowChartProcess', 'flowChartDecision', 'flowChartTerminator', 'flowChartData',
    'roundedRectCallout', 'ovalCallout',
]);

/**
 * 口语名 → ST_ShapeType 真名。
 * 这三个名字是从「形状长什么样」起的别名，DrawingML 里并不存在同名的 prst：
 * 直接写 `prst="flowChartData"` 这类值，自研解析器与 LibreOffice 都会照画，
 * 但真实 PowerPoint 打不开文件（实测：走修复流程且不给错误说明）。
 * 因此别名一律映射到语义最接近的合法值，画出来的仍然是用户想要的那个形状。
 */
const PRST_ALIASES = {
    flowChartData: 'flowChartInputOutput',
    roundedRectCallout: 'wedgeRoundRectCallout',
    ovalCallout: 'wedgeEllipseCallout',
};

/** a:prstDash 的合法取值。只收常用四种，其余一律 warning 后回落 solid。 */
const DASH_VALUES = ['solid', 'dash', 'sysDot', 'lgDash'];

/** a:headEnd / a:tailEnd 的 type 取值（none 表示不画箭头）。 */
const LINE_END_TYPES = ['none', 'triangle', 'arrow', 'oval', 'stealth', 'diamond'];

/** 角标位置：logo 与页码都用这一套语义。 */
const CORNERS = ['top-right', 'top-left', 'bottom-left', 'bottom-right'];

/** 装饰条的四种走向。 */
const ACCENT_BARS = ['left-top', 'left', 'bottom', 'top'];

/** master 默认在封面与结尾页不加页眉/页脚/页码（背景、logo、装饰条照加）。 */
const DEFAULT_SKIP_LAYOUTS = ['cover', 'closing'];

/**
 * 自由绘制页挂到哪一个 slideLayout 部件上。
 * 版式部件一共只有 8 个（SPEC 把 8 个版式名写死，且 master 必须给 8 条 sldLayoutId），
 * 所以自由页复用「陈述」版式的部件：本引擎的版式只是归属，几何全部自带，挂哪个都不影响渲染。
 */
const KIND_LAYOUT_KEY = { shape: 'statement' };

/** 渐变 stop 的数量上限：畸形输入不至于写出一段巨型 XML。 */
const MAX_GRADIENT_STOPS = 16;

// ───────────────────────────────────────────────────────────────────────────
// 通用工具
// ───────────────────────────────────────────────────────────────────────────

function colorMix(a, b, t) {
    const A = toHex(a);
    const B = toHex(b);
    const part = (at) => {
        const left = Number.parseInt(A.slice(at, at + 2), 16);
        const right = Number.parseInt(B.slice(at, at + 2), 16);
        return Math.round(left * (1 - t) + right * t).toString(16).padStart(2, '0');
    };
    return `${part(0)}${part(2)}${part(4)}`.toUpperCase();
}

function emuToCm(value) {
    return Math.round((value / 360000) * 100) / 100;
}

/**
 * 主题里的空串底色 → 回退色。
 * 素色主题（默认）的 `table.*Fill` 是空串，直接写进 `srgbClr val=""` 会被
 * `toHex` 当成非法值变成 000000：表头就成黑底黑字。回退到「与背景同色」，
 * 观感上等于不填充，XML 又始终合法。
 */
function solidOr(value, fallback) {
    return toHex(value, '') === '' ? fallback : value;
}

/**
 * 0~1 的比例钳制。新 API 里的 alpha / dim / pos 都是 0~1，
 * 非法值不抛异常，只夹回区间 —— 与「非法输入一律 warning」的约定一致。
 */
function clamp01(value) {
    const n = asNumber(value, 0);
    return Math.min(1, Math.max(0, n));
}

/** 角度归一到 [0,360)：a:lin 的 ang 与 a:outerShdw 的 dir 都是 60000/度。 */
function normalizeAngle(value) {
    const deg = asNumber(value, 0);
    return ((deg % 360) + 360) % 360;
}

/**
 * 「没有填充」的两种写法：空串、以及纯白 FFFFFF。
 * 素色主题（plain）的底色就是 FFFFFF，如果把它当成真实填充写进 XML，
 * 每张自由页都会凭空多一层白底，等于主题偷偷改变了用户的画布。
 */
function isBlankFill(value) {
    const hex = toHex(value, '');
    return hex === '' || hex === 'FFFFFF';
}

/** 渐变的一个色标：pos/alpha 都是 0~1。 */
function gradientStopXml(stop) {
    const alpha = stop === null || stop === undefined || stop.alpha === undefined || stop.alpha === null
        ? ''
        : `<a:alpha val="${Math.round(clamp01(stop.alpha) * 100000)}"/>`;
    return `<a:srgbClr val="${toHex(stop?.color)}">${alpha}</a:srgbClr>`;
}

/**
 * 线性渐变填充。
 * stops 少于 2 个构不成渐变：这里回落到第一个色标的纯色（再退到 noFill），
 * warning 由调用方（校验层）记，序列化层只保证「写出来的 XML 一定合法」。
 */
function gradientXml(fill) {
    const stops = asArray(fill.stops).slice(0, MAX_GRADIENT_STOPS);
    if (stops.length < 2) {
        const first = stops[0];
        return first === undefined ? '<a:noFill/>' : fillXml({ color: first.color });
    }
    const list = stops.map((stop) => `<a:gs pos="${Math.round(clamp01(stop?.pos) * 100000)}">`
        + gradientStopXml(stop) + '</a:gs>').join('');
    return `<a:gradFill rotWithShape="1"><a:gsLst>${list}</a:gsLst>`
        + `<a:lin ang="${Math.round(normalizeAngle(fill.angle) * 60000)}" scaled="0"/></a:gradFill>`;
}

/** 一个形状的填充：字符串 = 纯色，`{color, alpha}` = 带透明度的纯色，`{type:'gradient'}` = 渐变。 */
function fillXml(fill) {
    if (!fill || fill === 'none') return '<a:noFill/>';
    if (typeof fill === 'string') {
        return `<a:solidFill><a:srgbClr val="${toHex(fill)}"/></a:solidFill>`;
    }
    if (fill.type === 'gradient') return gradientXml(fill);
    const alpha = fill.alpha ? `<a:alpha val="${Math.round(fill.alpha * 1000)}"/>` : '';
    return `<a:solidFill><a:srgbClr val="${toHex(fill.color)}">${alpha}</a:srgbClr></a:solidFill>`;
}

/** dash 取值归一：true 是旧写法的「虚线」开关，其余按字面匹配。 */
function dashOf(raw, notes) {
    if (raw === undefined || raw === null || raw === false || raw === '') return 'solid';
    if (raw === true) return 'dash';
    const value = asString(raw, '').trim();
    if (DASH_VALUES.includes(value)) return value;
    if (notes !== undefined) {
        notes.push(`未知的虚线样式「${value}」，已按 solid 处理（可用：${DASH_VALUES.join(' / ')}）`);
    }
    return 'solid';
}

/** 单个线端：空串表示这一端不画箭头。 */
function lineEndOne(type) {
    const value = asString(type, '').trim();
    if (value === '' || value === 'none') return '';
    return LINE_END_TYPES.includes(value) ? value : 'triangle';
}

/**
 * 线端箭头。`arrow.begin` 写 a:headEnd（线的起点），`arrow.end` 写 a:tailEnd（线的终点）。
 * 两个元素都必须排在 a:prstDash 之后：CT_LineProperties 的顺序是
 * fill → prstDash → join → headEnd → tailEnd。
 */
function lineEndXml(arrow) {
    if (arrow === undefined || arrow === null || typeof arrow !== 'object') return '';
    const begin = lineEndOne(arrow.begin);
    const end = lineEndOne(arrow.end);
    return (begin === '' ? '' : `<a:headEnd type="${begin}" w="med" len="med"/>`)
        + (end === '' ? '' : `<a:tailEnd type="${end}" w="med" len="med"/>`);
}

/**
 * a:ln 序列化，兼容三种写法（前两种是既有调用方在用的，行为一字不变）：
 * - 字符串：纯色细线；
 * - 旧对象 `{color, dash:true}`：宽度由调用方给，缺省 12700；
 * - 新对象 `{color, widthPt, dash, alpha, arrow}`：宽度按磅给，alpha 是 0~100 的百分数，
 *   arrow 为 `{begin, end}`（deck.line 的 'end'/'both' 也归一成这个形状）。
 */
function lineXml(line, width) {
    if (line === undefined || line === null || line === false || line === 'none') {
        return '<a:ln><a:noFill/></a:ln>';
    }
    if (typeof line === 'object') {
        const dash = dashOf(line.dash);
        const explicitW = line.widthPt === undefined || line.widthPt === null
            ? undefined
            : Math.max(0, Math.round(asNumber(line.widthPt, 1) * EMU_PER_PT));
        const w = width ?? explicitW ?? (dash === 'dash' ? 12700 : 9525);
        const alpha = line.alpha ? `<a:alpha val="${Math.round(line.alpha * 1000)}"/>` : '';
        return `<a:ln w="${w}"><a:solidFill><a:srgbClr val="${toHex(line.color)}">${alpha}</a:srgbClr></a:solidFill>`
            + `<a:prstDash val="${dash}"/>${lineEndXml(line.arrow)}</a:ln>`;
    }
    return `<a:ln w="${width ?? 9525}"><a:solidFill><a:srgbClr val="${toHex(line)}"/></a:solidFill><a:prstDash val="solid"/></a:ln>`;
}

/**
 * 外阴影。`true` 用一套保守的默认值（4pt 模糊、2pt 偏移、45°、35% 黑）。
 * 只有用户显式要 shadow 时才写 a:effectLst：素色主题不能自己给形状加阴影。
 */
function shadowXml(shadow) {
    if (shadow === undefined || shadow === null || shadow === false) return '';
    const spec = shadow === true ? {} : shadow;
    if (typeof spec !== 'object') return '';
    const blur = Math.max(0, Math.round(asNumber(spec.blurPt, 4) * EMU_PER_PT));
    const dist = Math.max(0, Math.round(asNumber(spec.distPt, 2) * EMU_PER_PT));
    const dir = Math.round(normalizeAngle(spec.dirDeg === undefined ? 45 : spec.dirDeg) * 60000);
    const alpha = Math.round(clamp01(spec.alpha === undefined ? 0.35 : spec.alpha) * 100000);
    return `<a:effectLst><a:outerShdw blurRad="${blur}" dist="${dist}" dir="${dir}" rotWithShape="0">`
        + `<a:srgbClr val="${toHex(spec.color ?? '000000')}"><a:alpha val="${alpha}"/></a:srgbClr>`
        + '</a:outerShdw></a:effectLst>';
}

/**
 * 几何调整值（a:avLst）。
 * 目前只有圆角半径用到：roundRect 的 adj 是「相对短边」的千分之一，
 * 所以 `radius` 传厘米，这里换算成比例 —— 同一个半径在大小不同的卡片上观感一致。
 */
function avLstXml(opts) {
    if (opts.adj === undefined || opts.adj === null) return '<a:avLst/>';
    return `<a:avLst><a:gd name="adj" fmla="val ${Math.max(0, Math.min(50000, Math.round(opts.adj)))}"/></a:avLst>`;
}

/** 厘米保留一位小数：报告里的尺寸数字全部走这里，免得出现 5.666666666666667。 */
function cm1(value) {
    return Math.round(value * 10) / 10;
}

/** 厘米尺寸的展示形式，outline 与 read() 共用。 */
function sizeText(widthCm, heightCm) {
    return `${cm1(widthCm)}×${cm1(heightCm)}cm`;
}

function docXml(tag, body, extraNs = '') {
    return `${XML_HEAD}<${tag} xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"${extraNs}>${body}</${tag}>`;
}

function styleRef(body) {
    return `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}">${body}</Relationships>`;
}

/** `p:spTree` 的头两个子元素必须是 nvGrpSpPr 与 grpSpPr，且 xfrm 四个值全 0。 */
function spTreeSkeleton() {
    return '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        + '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
}

function bgXml(color) {
    return `<p:bg><p:bgPr>${fillXml(color)}<a:effectLst/></p:bgPr></p:bg>`;
}

/** 版心与版式的所有纵向刻度都从这里出，改一处全局一致。 */
function makeGeometry(cx, cy) {
    const margin = cmToEmu(0.9);
    const titleBandH = cmToEmu(2.85);
    const titleY = cmToEmu(0.80);
    const titleH = cmToEmu(1.55);
    const ruleY = cmToEmu(2.50);
    const ruleH = cmToEmu(0.09);
    const bodyY = cmToEmu(3.25);
    const footerY = cy - cmToEmu(1.10);
    return {
        cx,
        cy,
        margin,
        contentW: cx - margin * 2,
        contentRight: cx - margin,
        titleBandH,
        titleY,
        titleH,
        ruleY,
        ruleH,
        bodyY,
        bodyH: Math.max(cmToEmu(3), footerY - cmToEmu(0.25) - bodyY),
        footerY,
        footerH: cmToEmu(0.72),
    };
}

function wordsOf(text) {
    const value = String(text ?? '');
    const cjk = (value.match(/[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uac00-\ud7a3]/g) ?? []).length;
    const latin = (value.replace(/[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uac00-\ud7a3]/g, ' ')
        .match(/[A-Za-z0-9][A-Za-z0-9'’\-.]*/g) ?? []).length;
    return cjk + latin;
}

/**
 * 图注 + 署名（历史遗留 `1-2`「配图无来源」）。
 *
 * PPT 的题注是页面底部那一行居中小字，没有第二个自然落点，所以署名并进同一行；
 * 有题注时用 `｜` 分隔，只有署名时就直接是署名那一行。
 * @param {string} caption
 * @param {unknown} source 通常传 office.image.fetch 返回的 `credit`
 */
function withCredit(caption, source) {
    const credit = asString(source).trim();
    const text = asString(caption).trim();
    if (credit === '') return text;
    return text === '' ? `来源：${credit}` : `${text} ｜ 来源：${credit}`;
}

/**
 * 图片格式与像素尺寸：只认 PNG / JPEG / GIF 三种 PowerPoint 原生支持的位图。
 *
 * 导出（`images()`）与生成共用这一个探测器：两边若各写一套，同一张图在
 * 「画进去」和「抽出来」时可能被认成不同格式，报告立刻变成谎话。
 * `ppt-revise.js` 里有一份同源实现（那里不能静态引入本模块，会形成循环依赖）。
 */
export function imageInfo(raw) {
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

// ───────────────────────────────────────────────────────────────────────────
// 配色
// ───────────────────────────────────────────────────────────────────────────

/**
 * 主题 → 落地调色板。
 * 中间色（封面副标题、章节大号序号、引用符号）全部由主色与底色现算，
 * 这样深色主题与浅色主题共用同一套版式代码，对比度也不会凭感觉挑。
 */
function makePalette(theme) {
    const c = theme.colors;
    const table = theme.table;
    const dark = theme.dark;
    // 封面/结束页自带底色与文字色。多数主题直接用 primaryDark + onPrimary（深底白字），
    // 素色主题没有深色，于是通过 theme.cover 显式给出「白底黑字 + 浅灰色带」。
    const cover = theme.cover ?? {};
    return {
        dark,
        bg: c.bg,
        surface: c.surface,
        primary: c.primary,
        primaryDark: c.primaryDark,
        secondary: c.secondary,
        accent: c.accent,
        text: c.text,
        muted: c.muted,
        subtle: c.subtle,
        border: c.border,
        onPrimary: c.onPrimary,
        band: c.surface,
        bandLine: c.border,
        ghost: colorMix(c.primary, c.surface, 0.84),
        coverBg: cover.bg ?? c.primaryDark,
        coverBand: cover.band ?? c.primary,
        coverText: cover.text ?? c.onPrimary,
        coverKicker: cover.kicker ?? colorMix(c.secondary, c.onPrimary, 0.45),
        onBand: cover.onBand ?? colorMix(c.primary, c.onPrimary, 0.88),
        onBandMuted: cover.onBandMuted ?? colorMix(c.primary, c.onPrimary, 0.60),
        onCoverMuted: cover.muted ?? colorMix(c.primaryDark, c.onPrimary, 0.62),
        quoteMark: colorMix(c.accent, c.surface, 0.32),
        // 表格底色/表头文字同样要过 solidOr：素色主题给的是空串（= 不填充）
        tableHeaderFill: solidOr(table.headerFill, c.surface),
        tableHeaderText: solidOr(table.headerText, c.text),
        zebra: solidOr(table.zebraFill, c.bg),
        tableBorder: table.border,
        footer: dark ? colorMix(c.bg, c.text, 0.55) : c.muted,
    };
}

// ───────────────────────────────────────────────────────────────────────────
// 文本量算：所有 stats / warnings 的数字来源
// ───────────────────────────────────────────────────────────────────────────

/**
 * runs 的纯文本（runs 为空时给空串）。段落的「文字」与「runs」两个字段要一致，
 * 否则 warnings 里报的字数、read() 复检的行数会与画面脱节。
 * math run 用公式的纯文本近似顶上 —— 量算与字数统计都从这里走。
 */
function runText(runs) {
    return asArray(runs).map((run) => asString(run?.text) || (run?.math === undefined ? '' : mathPlain(run.math))).join('');
}

/**
 * 把要点数组编译成段落对象（字号、缩进、项目符号都在这里定死）。
 * 量算与实际 XML 都走这一个函数，保证「估算的行数」就是 PowerPoint 看到的行数。
 * `runs` 原样透传：段内混排时字号阶梯只作为「段落的基值」，run 自己写的字号优先。
 */
function buildBulletParas(items, baseSize, pal) {
    return items.map((item, index) => {
        const level = Math.min(Math.max(1, asNumber(item.level, 1)), LEVELS.length);
        const spec = LEVELS[level - 1];
        return {
            text: asString(item.text),
            runs: asArray(item.runs),
            sizePt: Math.max(10, baseSize - spec.drop),
            marL: spec.marL,
            indent: spec.indent,
            bullet: spec.bullet,
            bulletColor: level === 1 ? pal.accent : pal.subtle,
            color: level === 1 ? pal.text : pal.muted,
            lineSpacing: LINE_SPACING_BODY,
            spaceBeforePt: index === 0 ? 0 : spec.gapPt,
            level,
        };
    });
}

/**
 * 段落的纯文本。
 * 段落有两种写法：老的 `{text}` 与新的 `{runs:[{text}...]}`，量算、字数统计、
 * read() 复检都只读这一个入口 —— 两套写法各算一次，报告里就会出现两个不同的行数。
 */
function paragraphText(para) {
    const runs = asArray(para?.runs);
    if (runs.length > 0) {
        return runs.map((run) => asString(run?.text) || (run?.math === undefined ? '' : mathPlain(run.math))).join('');
    }
    return asString(para?.text);
}

/**
 * 段落的有效字号：run 里最大的那个字号决定折行估算。
 * 返回 0 表示「没指定」——wrapText 会按它一直以来的规则用 11pt 兜底，
 * 这样旧调用方（段落没有 sizePt 时）的行数估算与扩展前完全一致。
 */
function paragraphSize(para) {
    let size = asNumber(para?.sizePt, 0);
    for (const run of asArray(para?.runs)) {
        size = Math.max(size, asNumber(run?.sizePt ?? run?.size, 0));
    }
    return size;
}

/** 估算一组段落需要的高度（磅）与总行数。 */
function measureParas(paras, widthEmu) {
    let neededPt = 0;
    let lines = 0;
    let maxLines = 0;
    for (const para of paras) {
        const textWidth = Math.max(20, (widthEmu - (para.marL ?? 0)) / EMU_PER_PT);
        const size = paragraphSize(para);
        const count = Math.max(1, wrapText(paragraphText(para), textWidth, size).length);
        neededPt += count * size * (para.lineSpacing ?? 1.2) + (para.spaceBeforePt ?? 0);
        lines += count;
        maxLines = Math.max(maxLines, count);
    }
    return { neededPt, lines, maxLines };
}

/**
 * 在字号阶梯里挑第一个「放得下」的字号。
 * 全部放不下时返回最小字号那一档，并让调用方带着 neededPt 去写 warning ——
 * 这就是「溢出风险用真算」的落点。
 */
function fitParas(parasOf, widthEmu, heightEmu, sizes) {
    const heightPt = heightEmu / EMU_PER_PT;
    let last;
    for (const size of sizes) {
        const paras = parasOf(size);
        const measured = measureParas(paras, widthEmu);
        last = { sizePt: size, paras, ...measured, fits: measured.neededPt <= heightPt + 0.01 };
        if (last.fits) return last;
    }
    return last;
}

function fitTitle(text, widthEmu, heightEmu, sizes = [26, 24, 22, 20]) {
    return fitParas(
        (size) => [{
            text: asString(text), sizePt: size, bold: true, color: '000000', lineSpacing: 1.08,
        }],
        widthEmu,
        heightEmu,
        sizes,
    );
}

// ───────────────────────────────────────────────────────────────────────────
// 形状序列化
// ───────────────────────────────────────────────────────────────────────────

function shapeProps(opts) {
    // rot 是 60000/度；两个 flip 只有在直线需要表达方向时才写
    const rot = opts.rotate ? ` rot="${Math.round(asNumber(opts.rotate, 0) * 60000)}"` : '';
    const flip = `${opts.flipH === true ? ' flipH="1"' : ''}${opts.flipV === true ? ' flipV="1"' : ''}`;
    const parts = [
        `<a:xfrm${rot}${flip}><a:off x="${Math.round(opts.x)}" y="${Math.round(opts.y)}"/>`
        + `<a:ext cx="${Math.round(Math.max(1, opts.w))}" cy="${Math.round(Math.max(1, opts.h))}"/></a:xfrm>`,
        `<a:prstGeom prst="${opts.prst ?? 'rect'}">${avLstXml(opts)}</a:prstGeom>`,
        fillXml(opts.fill),
        lineXml(opts.line, opts.lineW),
        // a:effectLst 必须排在填充与描边之后（CT_ShapeProperties 的顺序）
        shadowXml(opts.shadow),
    ];
    return `<p:spPr>${parts.join('')}</p:spPr>`;
}

/**
 * 单条 run 的序列化（`a:r` + `a:rPr` + `a:t`）。
 *
 * 取值顺序是「run 自己的 → 段落基值（base）→ 兜底」，这样两种写法都能落到同一段代码：
 * - 旧写法 `runXml(para, font)`：base 缺省就是 run 自己，产出的 XML 与扩展前一字不差；
 * - 新写法 `runXml(run, font, para)`：段内混排时，没写的属性自动继承段落基值。
 *
 * `spacing` 与旧的 `spc` 是同一个东西（a:rPr 的 spc，单位 1/100 磅），两个名字都收。
 * run 带 `math` 时改走公式通道（a14:m 的 AlternateContent，见 mathRunXml）。
 */
function runXml(run, font, base = run, scope = {}) {
    if (run?.math !== undefined && run?.math !== null && String(run.math) !== '') {
        return mathRunXml(run, font, base, scope);
    }
    // run 的字号两个名字都收：段落对象里叫 sizePt，用户写 runs 时习惯写 size
    const size = Math.round(asNumber(run.sizePt ?? run.size ?? base.sizePt, 18) * 100);
    const bold = run.bold ?? base.bold;
    const italic = run.italic ?? base.italic;
    const underline = run.underline ?? base.underline;
    const spc = run.spacing ?? run.spc ?? base.spc;
    const attrs = [
        'lang="zh-CN"', 'altLang="en-US"', `sz="${size}"`,
        bold ? 'b="1"' : '', italic ? 'i="1"' : '', underline ? 'u="sng"' : '',
        spc ? `spc="${Math.round(asNumber(spc, 0))}"` : '', 'dirty="0"',
    ].filter((item) => item !== '').join(' ');
    const fill = `<a:solidFill><a:srgbClr val="${toHex(run.color ?? base.color ?? '000000')}"/></a:solidFill>`;
    // 单个 run / 段落基值可以换字体（字符串 = 中英文同一张脸；{en,cn} = 分开指定；
    // 图标字形就靠这个：Segoe MDL2 Assets 与正文完全不同）
    const rawFace = run.font ?? base.font;
    const face = fontFaceOf(rawFace, font);
    const faces = `<a:latin typeface="${escapeAttr(face.en)}"/>`
        + `<a:ea typeface="${escapeAttr(face.cn)}"/><a:cs typeface="${escapeAttr(face.en)}"/>`;
    // a:hlinkClick 排在 latin/ea/cs 之后（CT_TextCharacterProperties 的顺序）
    const relId = run.linkRelId ?? base.linkRelId;
    const link = relId === undefined || relId === '' ? '' : `<a:hlinkClick r:id="${escapeAttr(relId)}"/>`;
    // 短公式兜底（第五十八轮 新-23）：走到这里的都是**真·文本** run（math run 在上面就返回了），
    // 所以「看着像公式的短文本」能当场报出来，位置精确到「哪一页、哪个形状、第几个 run」。
    // 计数不在这里做：sink.push 按最终 XML 数一遍（同一个 run 走两次渲染也只算一条）。
    // 这里**不做本地去重**（逐条 push）：同一个形状里三条相同的要点是三个位置，
    // 而 compile 的 `warn()` 本来就会把逐字相同的告警并成一条（引擎既有口径），
    // 所以「重复文本只报一条告警」发生在那一层，`stats.plainFormulas` 照实报条数。
    if (scope.notes !== undefined && looksLikePlainFormula(run.text)) {
        const page = asNumber(scope.page, 0);
        const shape = asString(scope.shape);
        const runIndex = asNumber(scope.runIndex, 0);
        const where = `第 ${page} 页${shape === '' ? '' : `形状「${shape}」`}`
            + `${runIndex > 0 ? ` 第 ${runIndex} 个 run ` : ''}`;
        scope.notes.push(`${where}看着是公式却写成了普通文本（「${asString(run.text).trim().slice(0, 32)}」）：`
            + "短公式同样要走 math 通道，写成 runs:[{math:'r=qr_0'}]；确实要当文字用可以忽略这条");
    }
    return `<a:r><a:rPr ${attrs}>${fill}${faces}${link}</a:rPr><a:t>${escapeAttr(asString(run.text))}</a:t></a:r>`;
}

/**
 * 字体声明收敛：run/段落写的 font（字符串或 {en,cn}）→ {en,cn}；没写就用全篇默认。
 * 字体名是自由值（PowerPoint 缺字体时自己回落），这里只做转义与空值处理，不校验枚举。
 */
function fontFaceOf(raw, fallback) {
    const base = fallback !== null && typeof fallback === 'object'
        ? { en: asString(fallback.en, ''), cn: asString(fallback.cn, '') }
        : { en: '', cn: '' };
    if (raw === undefined || raw === null || raw === '') return base;
    if (typeof raw === 'object') {
        return {
            en: asString(raw.en, base.en),
            cn: asString(raw.cn, base.cn),
        };
    }
    return { en: asString(raw), cn: asString(raw) };
}

/**
 * 公式 run：a14:m 的 AlternateContent。
 *
 * - Choice 给 PowerPoint 2010+ / WPS / LibreOffice：原生 OMML，可在 PowerPoint 里继续编辑；
 * - Fallback 给不认 a14 的阅读器：公式的纯文本近似（斜体），内容不丢。
 * 段落里只有这一条 run 时按「展示式」排（oMathPara，独立成行居中），否则行内混排。
 * 解析失败不抛异常：按原样排成普通文字 run，并把原因写进 warnings。
 */
function mathRunXml(run, font, base, scope) {
    const latex = String(run.math);
    const size = Math.round(asNumber(run.sizePt ?? run.size ?? base.sizePt, 18) * 100);
    const color = toHex(run.color ?? base.color ?? '000000');
    const notes = asArray(scope.notes);
    const page = asNumber(scope.page, 0);
    // 数学 run 的字号/颜色写在 m:r 内嵌的 a:rPr 上；字体不写 —— 公式区由 Cambria Math 接管，
    // 显式写正文字体反而会破坏数学排版
    const rPr = `<a:rPr lang="en-US" altLang="zh-CN" sz="${size}" i="1" dirty="0">`
        + `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr>`;
    const parsed = latexToMath(latex, { rPr });
    if (!parsed.ok) {
        // 定位到「哪一页、哪个形状、第几个 math run」：只有页码的话，一页里几个公式
        // 同时失败时分不清是哪一条（第四十八轮 P0-1）。
        const shape = asString(scope.shape);
        const runIndex = asNumber(scope.runIndex, 0);
        const where = `第 ${page} 页${shape === '' ? '' : `形状「${shape}」`}`
            + `${runIndex > 0 ? ` 第 ${runIndex} 个 math run ` : ''}`;
        if (scope.notes !== undefined) {
            notes.push(`${where}公式无法解析（${parsed.error}）：${latex.slice(0, 40)} 已按原样排成普通文字`);
        }
        // 结构化的失败记录：stats.formulaErrors 与回执都从这里来（报告比文字更好用）
        if (Array.isArray(scope.formulaErrors)) {
            scope.formulaErrors.push({
                page, shape: shape === '' ? null : shape,
                runIndex: runIndex > 0 ? runIndex : null,
                latex: latex.slice(0, 80),
                error: parsed.error,
            });
        }
        return `<a:r><a:rPr lang="en-US" altLang="zh-CN" sz="${size}" i="1" dirty="0">`
            + `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr>`
            + `<a:t>${escapeAttr(latex)}</a:t></a:r>`;
    }
    const spaceRun = `<m:r>${rPr}<m:t> </m:t></m:r>`;
    const fallback = `<a:r><a:rPr lang="en-US" altLang="zh-CN" sz="${size}" i="1" dirty="0">`
        + `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr>`
        + `<a:t>${escapeAttr(parsed.plain)}</a:t></a:r>`;
    let choice;
    if (scope.mathDisplay === true) {
        const alignMap = { l: 'left', ctr: 'center', r: 'right', just: 'center', dist: 'center' };
        const jc = alignMap[textAlignOf(scope.align, 'ctr')] ?? 'centerGroup';
        const jcPr = jc === 'centerGroup' ? '' : `<m:oMathParaPr><m:jc m:val="${jc}"/></m:oMathParaPr>`;
        const mathList = parsed.blocks.map((block) => `<m:oMath xmlns:m="${NS_MATH}">${block}</m:oMath>`).join('');
        choice = `<m:oMathPara xmlns:m="${NS_MATH}">${jcPr}${mathList}</m:oMathPara>`;
    } else {
        const inner = parsed.blocks.join(spaceRun);
        choice = `<m:oMath xmlns:m="${NS_MATH}">${inner}</m:oMath>`;
    }
    return `<mc:AlternateContent xmlns:mc="${NS_MC}">`
        + `<mc:Choice xmlns:a14="${NS_A14}" Requires="a14"><a14:m>${choice}</a14:m></mc:Choice>`
        + `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
}

function escapeAttr(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/**
 * 段落序列化。
 * 子元素顺序由 schema 定死：lnSpc → spcBef → spcAft → buClr → buFont → buChar/buNone，
 * 顺序错了 PowerPoint 同样判定包损坏，所以这里不用「想到哪写到哪」的拼装方式。
 */
/** 文本对齐收敛：left/center/right 这类口语值 → ST_TextAlignType（l/ctr/r/just/dist…）。 */
const TEXT_ALIGN = {
    l: 'l', left: 'l', ctr: 'ctr', center: 'ctr', centre: 'ctr', r: 'r', right: 'r',
    just: 'just', justify: 'just', justified: 'just', dist: 'dist', distributed: 'dist',
    justlow: 'justLow', thaidist: 'thaiDist',
};

/** 垂直锚点收敛：top/middle/bottom → ST_TextAnchoringType（t/ctr/b/just/dist）。 */
const TEXT_ANCHOR = {
    t: 't', top: 't', ctr: 'ctr', center: 'ctr', centre: 'ctr', middle: 'ctr',
    b: 'b', bottom: 'b', just: 'just', dist: 'dist',
};

/**
 * 枚举属性必须收敛后再落 XML：`algn="left"` / `anchor="middle"` 都是非法取值，
 * 自研解析器与 LibreOffice 都照常渲染，真实 PowerPoint 却拒绝打开整个文件
 * （docs/ooxml-pitfalls.md 第 6 条那一类）。未知值一律回落默认值，绝不写非法属性。
 */
function textAlignOf(value, fallback) {
    if (value === undefined || value === null || value === '') return fallback;
    return TEXT_ALIGN[String(value).trim().toLowerCase()] ?? fallback;
}

function textAnchorOf(value, fallback) {
    if (value === undefined || value === null || value === '') return fallback;
    return TEXT_ANCHOR[String(value).trim().toLowerCase()] ?? fallback;
}

/**
 * 段落序列化。
 * 子元素顺序由 schema 定死：lnSpc → spcBef → spcAft → buClr → buFont → buChar/buNone，
 * 顺序错了 PowerPoint 同样判定包损坏，所以这里不用「想到哪写到哪」的拼装方式。
 * scope（notes/page/公式展示式判定）从 textBodyXml 的 opts 一路传进来 —— 公式解析
 * 失败的 warning 要带上页码，而这一层拿不到 ctx，只能由上层注入。
 */
function paragraphXml(para, font, scope = {}) {
    const attrs = [];
    const align = textAlignOf(para.align);
    if (align) attrs.push(`algn="${align}"`);
    if (para.marL) attrs.push(`marL="${Math.round(para.marL)}"`);
    if (para.indent) attrs.push(`indent="${Math.round(para.indent)}"`);
    const kids = [`<a:lnSpc><a:spcPct val="${Math.round((para.lineSpacing ?? 1.2) * 100000)}"/></a:lnSpc>`];
    if (para.spaceBeforePt) kids.push(`<a:spcBef><a:spcPts val="${Math.round(para.spaceBeforePt * 100)}"/></a:spcBef>`);
    if (para.spaceAfterPt) kids.push(`<a:spcAft><a:spcPts val="${Math.round(para.spaceAfterPt * 100)}"/></a:spcAft>`);
    if (para.bullet) {
        if (para.bulletColor) kids.push(`<a:buClr><a:srgbClr val="${toHex(para.bulletColor)}"/></a:buClr>`);
        kids.push('<a:buFont typeface="Arial"/>');
        kids.push(`<a:buChar char="${escapeAttr(para.bullet)}"/>`);
    } else {
        kids.push('<a:buNone/>');
    }
    const size = Math.round((paragraphSize(para) || 18) * 100);
    // 段内混排：runs 非空就逐条写 a:r；否则维持「一段一个 run」的旧行为
    const runs = asArray(para.runs).filter((run) => run !== null && typeof run === 'object');
    let body;
    if (runs.length > 0) {
        // 整段只有一条 math run 时按展示式排（独立成行的居中大公式）；
        // 与文字混排（哪怕夹着空 text run）一律行内 —— 与 PowerPoint 的两种插入方式对应
        const mathOnly = runs.length === 1
            && runs[0].math !== undefined && String(runs[0].math) !== ''
            && asString(runs[0].text) === '';
        const runScope = { ...scope, mathDisplay: mathOnly, align: para.align };
        // run 序号（段内第几个 run）一路传下去：公式失败的告警要能指到具体那一条
        body = runs.map((run, index) => runXml(run, font, para, { ...runScope, runIndex: index + 1 })).join('');
    } else if (para.text === undefined || para.text === '') {
        body = `<a:endParaRPr lang="zh-CN" sz="${size}"/>`;
    } else {
        // 旧写法（一段一个 run）也要带 scope：短公式告警与公式失败告警都靠它定位到
        // 「哪一页、哪个形状、第几个 run」；不传的话这一档会静默（参数含义与原地一致）。
        body = runXml(para, font, para, { ...scope, runIndex: 1 });
    }
    return `<a:p><a:pPr${attrs.length ? ` ${attrs.join(' ')}` : ''}>${kids.join('')}</a:pPr>${body}</a:p>`;
}

function textBodyXml(paras, font, opts = {}) {
    const anchor = textAnchorOf(opts.anchor, 't');
    const insets = opts.insets ?? { l: 0, t: 0, r: 0, b: 0 };
    const autofit = opts.autofit === false ? '<a:noAutofit/>' : '<a:normAutofit/>';
    const bodyPr = `<a:bodyPr wrap="square" lIns="${insets.l ?? 0}" tIns="${insets.t ?? 0}"`
        + ` rIns="${insets.r ?? 0}" bIns="${insets.b ?? 0}" anchor="${anchor}">${autofit}</a:bodyPr>`;
    const scope = { notes: opts.notes, page: opts.page, shape: opts.shape, formulaErrors: opts.formulaErrors };
    const list = paras.map((para) => paragraphXml(para, font, scope)).join('');
    return `<p:txBody>${bodyPr}<a:lstStyle/>${list === '' ? '<a:p><a:pPr><a:buNone/></a:pPr><a:endParaRPr lang="zh-CN"/></a:p>' : list}</p:txBody>`;
}

// ─── 短公式兜底（第五十八轮 新-23） ────────────────────────────────────────
//
// 为什么写侧与读侧各扫一遍：写侧（run 通道）能给出「哪一页、哪个形状、第几个 run」的
// 精确位置，读侧（a:t 节点）覆盖别人写的老稿子与第三方文件。
// 两处都**必须剔除 mc:Fallback** —— 那是公式成功解析后留给不认 a14 的阅读器的纯文本近似
// （`E = mc²`），把它当成「写成文字的公式」会得到一条永远甩不掉的假告警。
const PPT_MC_FALLBACK = /<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g;
const PPT_A_T = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g;

/** 从一段 shape / slide XML 字符串里抽出真·普通文本（剔除 mc:Fallback，反转义实体）。 */
function plainTextsOfXml(xml) {
    const out = [];
    for (const match of String(xml ?? '').replace(PPT_MC_FALLBACK, '').matchAll(PPT_A_T)) {
        out.push(decodeXmlEntities(match[1]));
    }
    return out;
}

/**
 * 从**解析后**的一页 XML 里收集普通文本节点：同样跳过 mc:Fallback 子树。
 * 解析树没有父指针，所以判据只能是「走到 Fallback 就不往下走」，不能事后按祖先过滤。
 */
function plainTextNodesOfTree(node, out = []) {
    for (const child of node?.children ?? []) {
        if (child.name === 'mc:Fallback') continue;
        if (child.name === 'a:t') out.push(child);
        else plainTextNodesOfTree(child, out);
    }
    return out;
}

/** 滑动形状收集器：负责 id 自增、rel 自增、文本面积记账。 */
function createSink() {
    return {
        shapes: [],
        rels: [],
        id: 1,
        contentArea: 0,
        words: 0,
        textBoxes: 0,
        tableCount: 0,
        imageCount: 0,
        // 原生公式数（a14:m）：报告里要能看出「这页有没有公式、有几个」
        formulas: 0,
        // 看着像公式却写成了普通文本的条数（第五十八轮 新-23）。三个公式口径是三件事，
        // 处置办法完全不同，合在一起等于都答不清：formulas = 写成了公式；
        // formulaErrors = 想写成公式但**解析失败**（改命令）；plainFormulas = **压根没走**
        // 公式通道（改成 {math:…}）。
        plainFormulas: 0,
        plainFormulaSamples: [],
        nextId() {
            this.id += 1;
            return this.id;
        },
        nextRelId() {
            return `rId${this.rels.length + 2}`;
        },
        push(xml) {
            this.formulas += (xml.match(/<a14:m[ >]/g) ?? []).length;
            const plain = plainTextsOfXml(xml).filter((text) => looksLikePlainFormula(text));
            this.plainFormulas += plain.length;
            for (const text of plain) {
                if (this.plainFormulaSamples.length < 8) this.plainFormulaSamples.push(text.trim());
            }
            this.shapes.push(xml);
        },
    };
}

function addRect(sink, opts) {
    const id = sink.nextId();
    const txBody = textBodyXml([], null, { autofit: false });
    sink.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeAttr(opts.name ?? `Decoration ${id}`)}"/>`
        + `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>${shapeProps(opts)}${txBody}</p:sp>`);
    return id;
}

function addText(sink, ctx, opts) {
    // opts.title: 标题类文本框（封面/章节/内容页/结尾的标题）用主题的「标题字体对」，
    // 与 fontScheme 的 majorFont 对齐；正文与其它元素用 minorFont（ctx.font）
    const font = opts.title === true ? ctx.titleFont : ctx.font;
    const paras = asArray(opts.paras);
    const id = sink.nextId();
    sink.textBoxes += 1;
    for (const para of paras) sink.words += wordsOf(paragraphText(para));
    const filled = paras.some((para) => paragraphText(para) !== '');
    if (filled) sink.contentArea += opts.w * opts.h;
    const name = opts.name ?? `TextBox ${id}`;
    const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeAttr(name)}"/>`
        + `<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>${shapeProps(opts)}`
        + textBodyXml(paras, font, {
            ...opts,
            notes: ctx.notes,
            page: ctx.page,
            shape: name,
            formulaErrors: ctx.formulaErrors,
        }) + '</p:sp>';
    sink.push(xml);
    return { id, lines: measureParas(paras, opts.w).lines };
}

/** 一条线的两个端点（直线形状）；越界与重合都只记 warning。 */
function lineGeometry(spec, ctx, notes) {
    const from = asArray(spec.from);
    const to = asArray(spec.to);
    if (from.length < 2 || to.length < 2) {
        notes.push(`第 ${ctx.page} 页 line 需要 from:[x,y] 与 to:[x,y]（厘米），已按 (0,0)→(0,0) 处理`);
    }
    const x1 = cmToEmu(asNumber(from[0], 0));
    const y1 = cmToEmu(asNumber(from[1], 0));
    const x2 = cmToEmu(asNumber(to[0], 0));
    const y2 = cmToEmu(asNumber(to[1], 0));
    const dx = x2 - x1;
    const dy = y2 - y1;
    if (dx === 0 && dy === 0) {
        notes.push(`第 ${ctx.page} 页 line 的起点与终点重合（${emuToCm(x1)}, ${emuToCm(y1)}），该线不可见`);
    }
    // 任意方向的直线用「左上角定位的矩形 + flipH/flipV」表达：这是 DrawingML 表达斜线的唯一办法
    return {
        x: Math.min(x1, x2),
        y: Math.min(y1, y2),
        w: Math.abs(dx),
        h: Math.abs(dy),
        flipH: dx < 0,
        flipV: dy < 0,
    };
}

/** 图标字形：接受 Unicode 码点（数字）、`U+E72C` / `0xE72C` / `&#xE72C;` 写法或单个字符。 */
function glyphText(raw, ctx, notes) {
    if (typeof raw === 'number') {
        if (Number.isInteger(raw) && raw >= 0 && raw <= 0x10ffff) return String.fromCodePoint(raw);
        notes.push(`第 ${ctx.page} 页 icon 的 glyph 码点 ${raw} 非法，已忽略该图标`);
        return '';
    }
    const text = asString(raw, '').trim();
    if (text === '') {
        notes.push(`第 ${ctx.page} 页 icon 缺少 glyph，已忽略该图标`);
        return '';
    }
    const matched = /^(?:U\+|0x|&#x?)([0-9a-fA-F]+);?$/.exec(text);
    if (matched !== null) {
        const code = Number.parseInt(matched[1], 16);
        if (code >= 0 && code <= 0x10ffff) return String.fromCodePoint(code);
        notes.push(`第 ${ctx.page} 页 icon 的 glyph 码点 ${text} 非法，已忽略该图标`);
        return '';
    }
    return text;
}

/**
 * 图标字形。
 * 只用「无填充无边框的矩形 + 居中文字」实现：DrawingML 里没有单独的图标元素，
 * 而字体图标（Segoe MDL2 Assets 之类）本来就是一个字形，写成形状+文字最稳。
 * 字号按「字形大约占满外框的 72%」折算，这样 sizeCm 才是用户能预期的视觉边长。
 */
function addIconShape(sink, ctx, spec) {
    const notes = ctx.notes;
    const glyph = glyphText(spec.glyph, ctx, notes);
    if (glyph === '') return undefined;
    const sizeCm = asNumber(spec.sizeCm, 1);
    const size = cmToEmu(sizeCm > 0 ? sizeCm : 1);
    const left = cmToEmu(asNumber(spec.x, 0));
    const top = cmToEmu(asNumber(spec.y, 0));
    const font = asString(spec.font, 'Segoe MDL2 Assets');
    const sizePt = Math.max(6, Math.round((size / EMU_PER_PT) * 0.72));
    const id = sink.nextId();
    const paras = [{
        sizePt,
        color: spec.color ?? ctx.pal.text,
        align: 'ctr',
        lineSpacing: 1,
        runs: [{ text: glyph, font, sizePt }],
    }];
    const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeAttr(spec.name ?? `Icon ${id}`)}"/>`
        + `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
        + shapeProps({ x: left, y: top, w: size, h: size, prst: 'rect' })
        + textBodyXml(paras, ctx.font, { anchor: 'ctr', autofit: false })
        + '</p:sp>';
    sink.push(xml);
    sink.words += wordsOf(glyph);
    return size;
}

/**
 * 自由形状。
 * 与版式里的装饰矩形（addRect）不同，这里的几何、填充、描边、旋转、圆角、阴影、
 * 形状内文字全部来自用户输入，因此校验与 warning 也都在这里收口：
 * 未知 preset 回落 rect、非法 dash 回落 solid、缺宽高只记 warning 不抛。
 */
function addAutoShape(sink, ctx, spec) {
    const notes = ctx.notes;
    const presetRaw = asString(spec.preset, 'rect').trim();
    let preset = 'rect';
    // kind 是给 stats / outline 用的名字（用户输入的那个），preset 是真正写进 prstGeom 的值。
    // 两者只在「口语名别名」这一种情况下不同：别名是合法输入，统计仍按用户写的名字算。
    let kind = 'rect';
    if (presetRaw !== '') {
        if (PRST_GEOMS.has(presetRaw)) {
            kind = presetRaw;
            preset = PRST_ALIASES[presetRaw] ?? presetRaw;
        } else {
            notes.push(`第 ${ctx.page} 页形状的 preset「${presetRaw}」不认识，已回落 rect`);
        }
    }
    const x = cmToEmu(asNumber(spec.x, 0));
    const y = cmToEmu(asNumber(spec.y, 0));
    const w = cmToEmu(asNumber(spec.w, 0));
    const h = cmToEmu(asNumber(spec.h, 0));
    if (!(w > 0) || !(h > 0)) {
        notes.push(`第 ${ctx.page} 页形状（${preset}）缺少宽高，尺寸按 0 处理，画出来不可见`);
    }
    const fill = shapeFillOf(spec.fill, ctx);
    const line = shapeLineOf(spec.line, ctx);
    let adj;
    if (spec.radius !== undefined && spec.radius !== null && spec.radius !== '') {
        if (preset === 'roundRect') {
            const radius = cmToEmu(Math.max(0, asNumber(spec.radius, 0)));
            const short = Math.max(1, Math.min(w > 0 ? w : 1, h > 0 ? h : 1));
            adj = (radius / short) * 100000;
        } else {
            notes.push(`第 ${ctx.page} 页只有 roundRect 支持 radius，已忽略（preset=${presetRaw}）`);
        }
    }
    const id = sink.nextId();
    const text = shapeTextXml(spec, ctx);
    for (const para of text.paras) sink.words += wordsOf(paragraphText(para));
    sink.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeAttr(spec.name ?? `Shape ${id}`)}"/>`
        // 自绘形状不用 txBox="1"：那会把形状降级成文本框，形状的填充/轮廓语义就没了
        + `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
        + shapeProps({
            x, y, w, h, prst: preset, fill, line, adj,
            rotate: spec.rotate, shadow: spec.shadow,
        })
        + text.body + '</p:sp>');
    return { preset: kind, x, y, w, h };
}

/** 形状内文字：text / textOpts.runs / textOpts.paras 三种写法都归一成段落数组。 */
function shapeTextXml(spec, ctx) {
    const opts = spec.textOpts !== undefined && spec.textOpts !== null && typeof spec.textOpts === 'object'
        ? spec.textOpts
        : {};
    const base = {
        sizePt: asNumber(opts.size, 18),
        bold: opts.bold === true,
        color: opts.color ?? ctx.pal.text,
        align: asString(opts.align, 'ctr'),
        lineSpacing: asNumber(opts.lineSpacing, 1.2),
        font: opts.font,
    };
    const paras = asArray(opts.paras);
    const runs = asArray(opts.runs);
    const list = paras.length > 0
        ? paras.map((para) => ({ ...base, ...(para ?? {}) }))
        : (runs.length > 0 ? [{ ...base, runs }] : [{ ...base, text: asString(spec.text) }]);
    const margin = opts.marginCm === undefined || opts.marginCm === null
        ? 0
        : cmToEmu(Math.max(0, asNumber(opts.marginCm, 0)));
    const insets = margin === 0 ? { l: 0, t: 0, r: 0, b: 0 } : { l: margin, t: margin, r: margin, b: margin };
    return {
        paras: list,
        body: textBodyXml(list, ctx.font, {
            anchor: asString(opts.anchor, 'ctr'),
            insets,
            notes: ctx.notes,
            page: ctx.page,
            // 形状名进 scope：形状里的公式失败要能指到「哪个形状」
            shape: asString(spec.name),
            formulaErrors: ctx.formulaErrors,
            // 形状里文字一律 noAutofit：autofit 会让 PowerPoint 自己缩字号，
            // 那报告里算出来的字号就与实际显示的不一致了
            autofit: false,
        }),
    };
}

/** 线条形状（deck.line）：prstGeom prst="line" + flipH/flipV + 两端箭头。 */
function addLineShape(sink, ctx, spec) {
    const notes = ctx.notes;
    const geo = lineGeometry(spec, ctx, notes);
    const line = shapeLineOf({
        color: spec.color,
        widthPt: spec.widthPt,
        dash: spec.dash,
        alpha: spec.alpha,
        arrow: lineArrowOf(spec.arrow, ctx),
    }, ctx);
    const id = sink.nextId();
    sink.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeAttr(spec.name ?? `Line ${id}`)}"/>`
        + `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
        + shapeProps({
            x: geo.x, y: geo.y, w: geo.w, h: geo.h, prst: 'line',
            flipH: geo.flipH, flipV: geo.flipV, line,
        })
        + textBodyXml([], ctx.font, { autofit: false }) + '</p:sp>');
    return geo;
}

/** deck.line 的 arrow 简写（'none' / 'end' / 'both'）→ a:headEnd / a:tailEnd 的对象。 */
function lineArrowOf(raw, ctx) {
    if (raw === undefined || raw === null || raw === false || raw === '' || raw === 'none') return undefined;
    if (typeof raw === 'object') {
        return {
            begin: raw.begin === undefined ? undefined : lineEndName(raw.begin),
            end: raw.end === undefined ? undefined : lineEndName(raw.end),
        };
    }
    const value = asString(raw, '').trim();
    if (value === 'end') return { end: 'triangle' };
    if (value === 'both') return { begin: 'triangle', end: 'triangle' };
    ctx.notes.push(`第 ${ctx.page} 页 line 的 arrow「${value}」不认识，已按 none 处理（可用：none / end / both）`);
    return undefined;
}

function lineEndName(raw) {
    const value = asString(raw, '').trim();
    return LINE_END_TYPES.includes(value) ? value : 'triangle';
}

/** 形状填充的校验与归一：字符串 / {color,alpha} / {type:'gradient'} / 'none'。 */
function shapeFillOf(raw, ctx) {
    if (raw === undefined || raw === null || raw === false || raw === 'none') return undefined;
    if (typeof raw === 'string') return isBlankFill(raw) ? undefined : raw;
    if (typeof raw !== 'object') return undefined;
    if (raw.type === 'gradient') {
        const stops = asArray(raw.stops);
        if (stops.length < 2) {
            ctx.notes.push(`第 ${ctx.page} 页渐变的 stops 只有 ${stops.length} 个，少于 2 个，已回落纯色`);
            return stops.length === 0 ? undefined : { color: stops[0]?.color };
        }
        return {
            type: 'gradient',
            stops: stops.map((stop) => ({
                pos: clamp01(stop?.pos),
                color: stop?.color,
                alpha: stop?.alpha,
            })),
            angle: asNumber(raw.angle, 0),
        };
    }
    if (isBlankFill(raw.color)) return undefined;
    // 新 API 的 alpha 是 0~1；fillXml 内部（含既有版式代码）沿用的是 0~100 的百分数
    return { color: raw.color, alpha: raw.alpha === undefined ? undefined : clamp01(raw.alpha) * 100 };
}

/** 形状描边的校验与归一：字符串 / {color,widthPt,dash,alpha,arrow} / 'none'。 */
function shapeLineOf(raw, ctx) {
    if (raw === undefined || raw === null || raw === false || raw === 'none') return undefined;
    if (typeof raw === 'string') return raw;
    if (typeof raw !== 'object') return undefined;
    const dash = dashOf(raw.dash, ctx.notes);
    return {
        color: raw.color ?? ctx.pal.subtle,
        widthPt: raw.widthPt,
        dash,
        alpha: raw.alpha === undefined ? undefined : clamp01(raw.alpha) * 100,
        arrow: lineArrowOf(raw.arrow, ctx),
    };
}

/** 单元格内容 → 纯文本（对象写法取 text 或 runs，数组由 paragraphText 兜底）。 */
function cellText(cell) {
    if (cell !== null && typeof cell === 'object' && !Array.isArray(cell)) {
        const runs = asArray(cell.runs);
        if (runs.length > 0) {
            return runs.map((run) => asString(run?.text) || (run?.math === undefined ? '' : mathPlain(run.math))).join('');
        }
        return asString(cell.text);
    }
    return asString(cell);
}

/** 单元格定义归一：字符串视为「只有文字」，对象取合并/样式/链接字段。 */
function normalizeCell(raw) {
    if (raw === null || raw === undefined || typeof raw !== 'object' || Array.isArray(raw)) {
        return {
            text: asString(raw), runs: [], colSpan: 1, rowSpan: 1,
            fill: undefined, color: undefined, bold: undefined, align: '', link: '',
        };
    }
    return {
        text: cellText(raw),
        runs: asArray(raw.runs),
        colSpan: Math.max(1, Math.floor(asNumber(raw.colSpan, 1))),
        rowSpan: Math.max(1, Math.floor(asNumber(raw.rowSpan, 1))),
        fill: raw.fill,
        color: raw.color,
        bold: raw.bold,
        align: asString(raw.align, ''),
        link: asString(raw.link ?? raw.href, ''),
    };
}

/**
 * 把 rows 摊平成「列数固定」的网格，并把合并信息记在 origin 单元格上。
 *
 * 为什么要有这一层：合并之后「一行的原始单元格个数」与「网格列数」不再相等，
 * 列宽、行高、XML 三个出口都必须读同一张网格，否则表头折行与行高会各算一套。
 *
 * 约定（与 PowerPoint 的行为一致）：
 * - 被 rowSpan 占住的列位在下一行不再接收原始值，原始值顺延到下一个空列；
 * - 覆盖格写 hMerge/vMerge，两个方向同时覆盖的对角格两个都写 1（实测 PowerPoint 亦如此）；
 * - 越界只裁剪 + warning，不抛异常。
 */
function buildTableGrid(rawRows, cols, page, notes) {
    const n = cols.length;
    const rowCount = rawRows.length;
    const grid = Array.from({ length: rowCount }, () => new Array(n).fill(null));
    rawRows.forEach((rawRow, r) => {
        const entries = Array.isArray(rawRow) ? rawRow.slice() : [rawRow];
        let c = 0;
        let dropped = 0;
        for (const entry of entries) {
            while (c < n && grid[r][c] !== null) c += 1;
            if (c >= n) {
                dropped += 1;
                continue;
            }
            const cell = normalizeCell(entry);
            const colSpan = Math.min(cell.colSpan, n - c);
            const rowSpan = Math.min(cell.rowSpan, rowCount - r);
            if (colSpan !== cell.colSpan) {
                notes.push(`第 ${page} 页表格第 ${r + 1} 行第 ${c + 1} 列的 colSpan=${cell.colSpan} 超出 `
                    + `${n} 列，已裁剪为 ${colSpan}`);
            }
            if (rowSpan !== cell.rowSpan) {
                notes.push(`第 ${page} 页表格第 ${r + 1} 行第 ${c + 1} 列的 rowSpan=${cell.rowSpan} 超出 `
                    + `${rowCount} 行，已裁剪为 ${rowSpan}`);
            }
            const origin = { ...cell, kind: 'origin', row: r, col: c, colSpan, rowSpan };
            grid[r][c] = origin;
            for (let dr = 0; dr < rowSpan; dr += 1) {
                for (let dc = 0; dc < colSpan; dc += 1) {
                    if (dr === 0 && dc === 0) continue;
                    if (grid[r + dr][c + dc] !== null) {
                        // 两个合并区域叠在一起：多出来的格子丢掉，但要让用户知道
                        notes.push(`第 ${page} 页表格第 ${r + dr + 1} 行第 ${c + dc + 1} 列被两个合并区域覆盖，已保留先出现的那个`);
                        continue;
                    }
                    grid[r + dr][c + dc] = { kind: 'covered', h: dc > 0, v: dr > 0 };
                }
            }
            c += colSpan;
        }
        if (dropped > 0) {
            notes.push(`第 ${page} 页表格第 ${r + 1} 行有 ${dropped} 个单元格超出行宽 ${n} 列，已丢弃`);
        }
        for (let col = 0; col < n; col += 1) {
            if (grid[r][col] !== null) continue;
            grid[r][col] = {
                kind: 'origin', row: r, col, colSpan: 1, rowSpan: 1,
                text: '', runs: [], fill: undefined, color: undefined, bold: undefined, align: '', link: '',
            };
        }
    });
    return grid;
}

/**
 * 表格的样式选项：headerFill / zebra / border / cellPadCm / rowHeightCm / firstRowBold。
 * 全部缺省时与扩展前的观感一致（表头用主题色、隔行浅底、只画横线）。
 */
function tableOptions(spec) {
    const pad = spec.cellPadCm === undefined || spec.cellPadCm === null
        ? undefined
        : cmToEmu(Math.max(0, asNumber(spec.cellPadCm, 0)));
    let zebra;
    if (spec.zebra === undefined || spec.zebra === null) zebra = true;
    else if (spec.zebra === false || spec.zebra === 'none' || spec.zebra === '') zebra = false;
    else if (spec.zebra === true) zebra = true;
    else if (typeof spec.zebra === 'string') zebra = isBlankFill(spec.zebra) ? false : spec.zebra;
    else if (typeof spec.zebra === 'object') zebra = spec.zebra.fill ?? spec.zebra.color ?? true;
    else zebra = true;
    return {
        headerFill: spec.headerFill,
        zebra,
        border: spec.border,
        insets: pad === undefined ? undefined : { l: pad, t: pad, r: pad, b: pad },
        rowHeight: spec.rowHeightCm === undefined || spec.rowHeightCm === null
            ? 0
            : cmToEmu(Math.max(0, asNumber(spec.rowHeightCm, 0))),
        firstRowBold: spec.firstRowBold !== false,
    };
}

/** 单元格四边描边：border 给定时四边全画，否则维持「只画底部横线」的旧观感。 */
function cellBorderLines(opts) {
    if (opts.borders !== undefined && opts.borders !== null && opts.borders !== 'none') {
        const spec = typeof opts.borders === 'string' ? { color: opts.borders } : opts.borders;
        const color = spec.color ?? opts.pal?.tableBorder ?? 'D9D9D9';
        const width = spec.widthPt === undefined || spec.widthPt === null
            ? 9525
            : Math.max(0, Math.round(asNumber(spec.widthPt, 0.75) * EMU_PER_PT));
        const dash = dashOf(spec.dash, opts.notes);
        const one = (tag) => `<a:${tag} w="${width}"><a:solidFill><a:srgbClr val="${toHex(color)}"/></a:solidFill>`
            + `<a:prstDash val="${dash}"/></a:${tag}>`;
        return `${one('lnL')}${one('lnR')}${one('lnT')}${one('lnB')}`;
    }
    const lines = ['<a:lnL w="0"><a:noFill/></a:lnL>', '<a:lnR w="0"><a:noFill/></a:lnR>', '<a:lnT w="0"><a:noFill/></a:lnT>'];
    if (opts.bottomLine) {
        lines.push(`<a:lnB w="${opts.bottomLine.width}"><a:solidFill>`
            + `<a:srgbClr val="${toHex(opts.bottomLine.color)}"/></a:solidFill><a:prstDash val="solid"/></a:lnB>`);
    } else {
        lines.push('<a:lnB w="0"><a:noFill/></a:lnB>');
    }
    return lines.join('');
}

/**
 * 一个单元格（a:tc）。
 * 默认路径（字符串内容 + 无 link + 无 runs）产出的 XML 与扩展前完全相同，
 * 因为表格外观是全项目最容易回归的地方，任何改动都必须是「只对用到的字段生效」。
 */
function tableCellXml(cell, opts) {
    const font = opts.font;
    const paras = asArray(opts.paras).length > 0 ? asArray(opts.paras) : [{
        text: opts.text,
        // 段内混排：runs 为空时 paragraphXml 会退回「一段一个 run」的老路径
        runs: asArray(opts.runs),
        sizePt: opts.sizePt,
        bold: opts.bold,
        color: opts.color,
        align: opts.align,
        lineSpacing: 1.22,
        linkRelId: opts.linkRelId,
        underline: opts.underline,
    }];
    const insets = opts.insets ?? { l: TABLE_MAR_H, t: TABLE_MAR_V, r: TABLE_MAR_H, b: TABLE_MAR_V };
    const cellAnchor = textAnchorOf(opts.anchor, 'ctr');
    const bodyPr = `<a:bodyPr wrap="square" lIns="${insets.l}" tIns="${insets.t}"`
        + ` rIns="${insets.r}" bIns="${insets.b}" anchor="${cellAnchor}"><a:noAutofit/></a:bodyPr>`;
    const txBody = `<a:txBody>${bodyPr}<a:lstStyle/>${paras.map((para) => paragraphXml(para, font, {
        notes: opts.notes,
        page: opts.page,
        shape: opts.shape,
        formulaErrors: opts.formulaErrors,
    })).join('')}</a:txBody>`;
    const span = `${opts.colSpan > 1 ? ` gridSpan="${opts.colSpan}"` : ''}${opts.rowSpan > 1 ? ` rowSpan="${opts.rowSpan}"` : ''}`;
    const tcPr = `<a:tcPr marL="${insets.l}" marR="${insets.r}" marT="${insets.t}" marB="${insets.b}"`
        + ` anchor="${cellAnchor}">${cellBorderLines(opts)}${fillXml(opts.fill)}</a:tcPr>`;
    return `<a:tc${span}>${txBody}${tcPr}</a:tc>`;
}

/**
 * 被合并覆盖的格子：hMerge / vMerge 是 a:tc 的属性（不是 a:tcPr 的，实测 PowerPoint 输出）。
 * 内容必须为空 —— 留着原文会让同一段文字被数两遍。
 */
function mergedCellXml(cell) {
    const flags = `${cell.h ? ' hMerge="1"' : ''}${cell.v ? ' vMerge="1"' : ''}`;
    return `<a:tc${flags}><a:txBody><a:bodyPr/><a:lstStyle/>`
        + '<a:p><a:pPr><a:buNone/></a:pPr><a:endParaRPr lang="zh-CN"/></a:p></a:txBody><a:tcPr/></a:tc>';
}

// ───────────────────────────────────────────────────────────────────────────
// 表格量算
// ───────────────────────────────────────────────────────────────────────────

/**
 * 列宽：显式 width 当权重用；没给就按「表头 + 各 origin 格」的估算字宽分配。
 * 只按等分会让中文表头频繁折行，这一层估算直接决定表格好不好看。
 * 覆盖格（被合并吃掉的格子）没有内容，不参与权重。
 */
function tableColumnWidths(cols, grid, contentW) {
    const weights = cols.map((col, index) => {
        if (col.width > 0) return col.width;
        let weight = estimateEm(col.title) + 2.5;
        for (const row of grid) {
            const cell = row[index];
            if (cell === null || cell.kind !== 'origin') continue;
            weight = Math.max(weight, estimateEm(cell.text) + 2);
        }
        return Math.min(26, Math.max(4.5, weight));
    });
    const total = weights.reduce((sum, item) => sum + item, 0) || 1;
    const minCol = cmToEmu(1.15);
    let used = 0;
    return weights.map((weight, index) => {
        if (index === weights.length - 1) return Math.max(cmToEmu(0.7), contentW - used);
        const width = Math.max(minCol, Math.min(contentW - used - minCol, Math.round((contentW * weight) / total)));
        used += width;
        return width;
    });
}

/** 一行里「有内容、要给高度」的格子：宽度按跨列求和。 */
function rowCells(row, widths) {
    const cells = [];
    row.forEach((cell, index) => {
        if (cell === null || cell.kind !== 'origin') return;
        let width = 0;
        for (let k = 0; k < cell.colSpan; k += 1) width += widths[index + k] ?? 0;
        cells.push({ text: cell.text, width });
    });
    return cells;
}

function tableRowHeight(cells, sizePt, isHeader) {
    let lines = 1;
    for (const cell of cells) {
        const textWidth = Math.max(12, (cell.width - TABLE_MAR_H * 2) / EMU_PER_PT);
        lines = Math.max(lines, wrapText(asString(cell.text), textWidth, sizePt).length);
    }
    const value = Math.round(lines * sizePt * 1.22 * EMU_PER_PT + TABLE_MAR_V * 2);
    return Math.max(value, isHeader ? cmToEmu(1.02) : cmToEmu(0.74));
}

/** 整张表的字号 + 行高：先挑字号，再在超版心时按比例压行高（并如实报告）。 */
function layoutTable(spec, g, pal, page, notes) {
    const cols = normalizeColumns(spec.columns);
    if (cols.length === 0) return undefined;
    const rawRows = asArray(spec.rows);
    const grid = buildTableGrid(rawRows, cols, page, notes);
    const options = tableOptions(spec);
    const widths = tableColumnWidths(cols, grid, g.contentW);
    const headerCells = cols.map((col, index) => ({ text: col.title, width: widths[index] }));
    const bodyCells = grid.map((row) => rowCells(row, widths));
    const sizes = [12, 11.5, 11, 10.5, 10, 9.5, 9];
    const floor = (heights) => heights.map((height) => Math.max(height, options.rowHeight));
    const heightsOf = (size) => floor([
        tableRowHeight(headerCells, size + 0.5, true),
        ...bodyCells.map((cells) => tableRowHeight(cells, size, false)),
    ]);
    let chosen = sizes[sizes.length - 1];
    let heights = heightsOf(chosen);
    let total = heights.reduce((sum, item) => sum + item, 0);
    for (const size of sizes) {
        chosen = size;
        heights = heightsOf(size);
        total = heights.reduce((sum, item) => sum + item, 0);
        if (total <= g.bodyH) break;
    }
    let compressed = false;
    let needed = total;
    if (total > g.bodyH) {
        // 最小字号仍然超版心：按比例压行高，但保留文字所需的最小高度，压完照样报出来
        needed = total;
        const ratio = g.bodyH / total;
        heights = heights.map((height, index) => {
            const floorH = tableRowHeight(
                index === 0 ? headerCells : bodyCells[index - 1],
                chosen + (index === 0 ? 0.5 : 0),
                index === 0,
            ) * 0.72;
            return Math.max(Math.round(height * ratio), Math.round(floorH));
        });
        total = heights.reduce((sum, item) => sum + item, 0);
        compressed = true;
    }
    return { cols, rows: grid, widths, heights, sizePt: chosen, total, needed, compressed, options };
}

// ───────────────────────────────────────────────────────────────────────────
// 页面骨架：内容页共用的标题带 / 页脚
// ───────────────────────────────────────────────────────────────────────────

function addFooter(sink, ctx, label, options = {}) {
    const { g, pal } = ctx;
    // 配了母版就把页脚/页码的责任整体交给母版：两套页码同时画出来是最常见的事故
    if (ctx.master !== undefined) return;
    const color = options.color ?? pal.footer;
    const lineColor = options.lineColor ?? pal.bandLine;
    // 页脚是页面级装饰：分栏页（新-6）上仍横贯整幅版心，不跟着收窄
    const fullW = g.fullContentW ?? g.contentW;
    const fullRight = g.fullContentRight ?? g.contentRight;
    addRect(sink, {
        x: g.margin, y: g.footerY - cmToEmu(0.20), w: fullW, h: 9525,
        fill: lineColor, name: 'Footer Rule',
    });
    addText(sink, ctx, {
        x: g.margin, y: g.footerY, w: Math.round(fullW * 0.7), h: g.footerH,
        anchor: 't', name: 'Footer Text', autofit: false,
        paras: [{ text: label, sizePt: 10, color, lineSpacing: 1 }],
    });
    addText(sink, ctx, {
        x: fullRight - Math.round(fullW * 0.3), y: g.footerY,
        w: Math.round(fullW * 0.3), h: g.footerH, anchor: 't', name: 'Page Number', autofit: false,
        paras: [{ text: `${ctx.page} / ${ctx.total}`, sizePt: 10, color, align: 'r', lineSpacing: 1 }],
    });
}

/** 内容页顶部：浅色带 + 标题 + 强调短线。标题层级靠色带与色块建立，不靠占位符。 */
function addContentChrome(sink, ctx, title, options = {}) {
    const { g, pal } = ctx;
    addRect(sink, { x: 0, y: 0, w: g.cx, h: g.titleBandH, fill: pal.band, name: 'Title Band' });
    addRect(sink, {
        x: 0, y: g.titleBandH - cmToEmu(0.04), w: g.cx, h: cmToEmu(0.04),
        fill: pal.bandLine, name: 'Title Band Rule',
    });
    // 标题同样横贯整幅版心：分栏页上标题缩到半幅会显得整页没对齐（新-6）
    const boxW = (g.fullContentW ?? g.contentW) - cmToEmu(1.4);
    const fit = fitTitle(title, boxW, g.titleH);
    addText(sink, ctx, {
        x: g.margin, y: g.titleY, w: boxW, h: g.titleH, anchor: 'ctr',
        name: 'Title', autofit: false, defaultColor: pal.primary, title: true,
        paras: [{
            text: asString(title), sizePt: fit.sizePt, bold: true,
            color: pal.primary, lineSpacing: 1.08, align: 'l',
        }],
    });
    addRect(sink, {
        x: g.margin, y: g.ruleY, w: cmToEmu(2.4), h: g.ruleH,
        fill: pal.accent, name: 'Title Accent',
    });
    if (options.footer !== false) addFooter(sink, ctx, options.footerLabel ?? ctx.docTitle);
    return fit;
}

// ───────────────────────────────────────────────────────────────────────────
// 8 种版式的渲染
// ───────────────────────────────────────────────────────────────────────────

function renderCover(sink, ctx, data) {
    const { g, pal } = ctx;
    const title = asString(data.title ?? ctx.docTitle);
    const kicker = asString(data.kicker);
    const presenter = asString(data.presenter ?? ctx.author);
    const date = asString(data.date ?? ctx.date);
    const titleY = Math.round(g.cy * 0.40);
    const titleW = g.contentW - cmToEmu(1.6);
    const runs = asArray(data.runs);
    const fit = fitParas(
        (size) => [{
            text: title, runs, sizePt: size, bold: true, color: pal.coverText, lineSpacing: 1.12,
        }],
        titleW,
        cmToEmu(3.6),
        [40, 36, 32, 28, 26],
    );
    addRect(sink, {
        x: g.cx - cmToEmu(3.6), y: -cmToEmu(2.2), w: cmToEmu(9.6), h: cmToEmu(9.6),
        fill: { color: pal.coverText, alpha: 7 }, prst: 'ellipse', name: 'Cover Glow',
    });
    addRect(sink, {
        x: 0, y: Math.round(g.cy * 0.66), w: g.cx, h: g.cy - Math.round(g.cy * 0.66),
        fill: pal.coverBand, name: 'Cover Band',
    });
    addRect(sink, {
        x: g.margin, y: Math.round(g.cy * 0.30), w: cmToEmu(3.2), h: cmToEmu(0.14),
        fill: pal.accent, name: 'Cover Accent',
    });
    if (kicker !== '') {
        addText(sink, ctx, {
            x: g.margin, y: Math.round(g.cy * 0.325), w: titleW, h: cmToEmu(0.9),
            anchor: 'ctr', name: 'Kicker', autofit: false,
            paras: [{ text: kicker, sizePt: 13, bold: true, color: pal.coverKicker, spc: 300, lineSpacing: 1 }],
        });
    }
    addText(sink, ctx, {
        x: g.margin, y: titleY, w: titleW, h: cmToEmu(3.6), anchor: 't', name: 'Title', autofit: false, title: true,
        paras: [{
            text: title, runs, sizePt: fit.sizePt, bold: true, color: pal.coverText, lineSpacing: 1.12,
        }],
    });
    const subtitleY = titleY + Math.round(fit.neededPt * EMU_PER_PT) + cmToEmu(0.30);
    if (asString(data.subtitle) !== '') {
        addText(sink, ctx, {
            x: g.margin, y: subtitleY, w: titleW, h: cmToEmu(1.4), anchor: 't', name: 'Subtitle', autofit: false,
            paras: [{ text: asString(data.subtitle), sizePt: 18, color: pal.onCoverMuted, lineSpacing: 1.3 }],
        });
    }
    if (presenter !== '') {
        addText(sink, ctx, {
            // 底边留在 cy-2.1cm：日期框从 cy-1.75cm 起，两者必须留出间隙，
            // 否则「团队 + 日期」会叠在一起（长团队名尤其明显）
            x: g.margin, y: g.cy - cmToEmu(3.0), w: g.contentW, h: cmToEmu(0.9),
            anchor: 'b', name: 'Presenter', autofit: false,
            paras: [{ text: presenter, sizePt: 13, bold: true, color: pal.onBand, lineSpacing: 1 }],
        });
    }
    if (date !== '') {
        addText(sink, ctx, {
            x: g.margin, y: g.cy - cmToEmu(1.75), w: g.contentW, h: cmToEmu(0.8),
            anchor: 't', name: 'Date', autofit: false,
            paras: [{ text: date, sizePt: 11.5, color: pal.onBandMuted, lineSpacing: 1 }],
        });
    }
    return { title, subtitle: asString(data.subtitle), fitted: fit };
}

function renderSection(sink, ctx, data) {
    const { g, pal } = ctx;
    const title = asString(data.title);
    const titleW = g.contentW - cmToEmu(1.4);
    const runs = asArray(data.runs);
    const fit = fitParas(
        (size) => [{ text: title, runs, sizePt: size, bold: true, color: pal.primary, lineSpacing: 1.1 }],
        titleW,
        cmToEmu(2.2),
        [34, 30, 26, 24, 22],
    );
    addRect(sink, { x: 0, y: 0, w: cmToEmu(0.42), h: g.cy, fill: pal.primary, name: 'Section Spine' });
    addText(sink, ctx, {
        x: g.margin, y: Math.round(g.cy * 0.24), w: cmToEmu(9), h: cmToEmu(4.0),
        anchor: 'b', name: 'Section Number', autofit: false,
        paras: [{
            text: String(ctx.sectionIndex).padStart(2, '0'), sizePt: 88, bold: true,
            color: pal.ghost, lineSpacing: 0.95,
        }],
    });
    addRect(sink, {
        x: g.margin, y: Math.round(g.cy * 0.545), w: cmToEmu(2.0), h: cmToEmu(0.11),
        fill: pal.accent, name: 'Section Accent',
    });
    addText(sink, ctx, {
        x: g.margin, y: Math.round(g.cy * 0.575), w: titleW, h: cmToEmu(2.2),
        anchor: 't', name: 'Title', autofit: false, title: true,
        paras: [{ text: title, runs, sizePt: fit.sizePt, bold: true, color: pal.primary, lineSpacing: 1.1 }],
    });
    if (asString(data.subtitle) !== '') {
        addText(sink, ctx, {
            x: g.margin, y: Math.round(g.cy * 0.755), w: titleW, h: cmToEmu(1.2),
            anchor: 't', name: 'Subtitle', autofit: false,
            paras: [{ text: asString(data.subtitle), sizePt: 15, color: pal.muted, lineSpacing: 1.3 }],
        });
    }
    addFooter(sink, ctx, ctx.docTitle);
    return { title, subtitle: asString(data.subtitle) };
}

function renderBullets(sink, ctx, data) {
    const { g, pal } = ctx;
    const items = asArray(data.items).map((item) => (typeof item === 'object' && item !== null
        ? {
            // runs 必须一起带过来：只取 text 会把段内混排丢在这一层
            text: asString(item.text ?? item.title) || runText(item.runs),
            runs: asArray(item.runs),
            level: asNumber(item.level, 1),
        }
        : { text: asString(item), level: 1 }));
    const chrome = addContentChrome(sink, ctx, data.title);
    const columnCount = Math.max(1, Math.min(2, Math.floor(asNumber(data.columns, 1))));
    const columns = [];
    if (columnCount === 2 && items.length >= 4) {
        let cut = Math.ceil(items.length / 2);
        // 分栏时不让二级要点成为新栏的第一条：它没有上级，视觉上会突然缩进
        while (cut > 1 && cut < items.length && asNumber(items[cut].level, 1) > 1) cut -= 1;
        columns.push(items.slice(0, cut), items.slice(cut));
    } else {
        columns.push(items);
    }
    const gutter = cmToEmu(0.9);
    const columnW = Math.round((g.contentW - gutter * (columns.length - 1)) / columns.length);
    const sizes = [18, 17, 16, 15, 14, 13];
    const fitted = columns.map((columnItems) => fitParas(
        (size) => buildBulletParas(columnItems, size, pal),
        columnW,
        g.bodyH,
        sizes,
    ));
    const size = Math.min(...fitted.map((item) => item.sizePt));
    columns.forEach((columnItems, index) => {
        const paras = buildBulletParas(columnItems, size, pal);
        const measured = measureParas(paras, columnW);
        addText(sink, ctx, {
            x: g.margin + index * (columnW + gutter), y: g.bodyY, w: columnW, h: g.bodyH,
            anchor: 't', name: `Bullets ${index + 1}`,
            paras,
        });
        fitted[index] = { ...fitted[index], paras, ...measured, sizePt: size, fits: measured.neededPt <= g.bodyH / EMU_PER_PT };
    });
    if (columns.length === 2) {
        addRect(sink, {
            x: g.margin + columnW + Math.round(gutter / 2) - 4763, y: g.bodyY,
            w: 9525, h: Math.round(g.bodyH * 0.92), fill: pal.bandLine, name: 'Column Divider',
        });
    }
    return {
        title: asString(data.title),
        items,
        columns: columns,
        fitted,
        sizePt: size,
        lines: fitted.reduce((sum, item) => sum + item.lines, 0),
        titleFit: chrome,
        contentW: g.contentW,
    };
}

/**
 * 原生图表页（历史遗留 `11-6`）。
 *
 * 画的是一个 `p:graphicFrame`，内容 `<c:chart r:id>` 指向 `ppt/charts/chartN.xml`
 * —— 与 Excel 那边共用同一份 `c:chartSpace`（见 formats/chart.js）。数据内联在图表里，
 * 所以幻灯片打开就能画；PowerPoint 里它是一个**图表对象**（可改类型、可编辑数据）。
 *
 * 编号：`ctx.chartSpecs` 是这次 compile 的图表清单，第 n 个就是 chartN.xml，
 * 关系目标按同一编号算 —— 一处编号，两处引用，不会错位。
 */
function renderChart(sink, ctx, data) {
    const { g, pal } = ctx;
    addContentChrome(sink, ctx, data.title);
    ctx.chartSpecs.push(data.spec);
    const chartIndex = ctx.chartSpecs.length;
    const relId = sink.nextRelId();
    sink.rels.push({ id: relId, type: RT.chart, target: `../charts/chart${chartIndex}.xml` });
    // 图表占满版心：标题条由 addContentChrome 画在顶部，图从 bodyY 起、铺到 bodyH。
    sink.push(presentationChartFrame({
        id: sink.nextId(),
        name: `图表 ${chartIndex}`,
        x: g.margin,
        y: g.bodyY,
        cx: g.contentW,
        cy: Math.max(cmToEmu(4), g.bodyH),
        chartRelId: relId,
    }));
    sink.chartCount = (sink.chartCount ?? 0) + 1;
    return {
        title: asString(data.title),
        chart: chartIndex,
        type: data.spec.type,
        series: data.spec.series.length,
        kinds: { chart: 1 },
    };
}

function renderTable(sink, ctx, data) {
    const { g, pal } = ctx;
    addContentChrome(sink, ctx, data.title);
    const layout = layoutTable(data, g, pal, ctx.page, ctx.notes);
    if (layout === undefined) {
        return { title: asString(data.title), layout: undefined, columns: 0, rows: 0, merges: 0 };
    }
    const { cols, rows, widths, heights, sizePt, total, compressed, options } = layout;
    const aligns = cols.map(() => 'l');
    const headerFill = isBlankFill(options.headerFill) ? pal.tableHeaderFill : options.headerFill;
    const borders = options.border === 'none' ? undefined : options.border;
    const header = cols.map((col) => tableCellXml(col.title, {
        font: ctx.font,
        text: col.title,
        sizePt: sizePt + 0.5,
        bold: options.firstRowBold,
        color: pal.tableHeaderText,
        align: 'l',
        fill: headerFill,
        insets: options.insets,
        borders,
        notes: ctx.notes,
        page: ctx.page,
        shape: asString(data.title, '表格'),
        formulaErrors: ctx.formulaErrors,
        pal,
        bottomLine: { color: pal.accent, width: 25400 },
    })).join('');
    let merges = 0;
    const body = rows.map((row, rowIndex) => {
        const zebraFill = options.zebra === true
            ? (rowIndex % 2 === 1 ? pal.zebra : pal.bg)
            : (options.zebra === false ? pal.bg : options.zebra);
        const cells = row.map((cell) => {
            if (cell.kind === 'covered') return mergedCellXml(cell);
            const type = cellType(cols[cell.col].type, cell.text);
            const numeric = type === 'number' || type === 'currency' || type === 'percent';
            const colAlign = asString(cell.align, '') !== '' ? cell.align : cols[cell.col].align;
            const align = colAlign === 'center' || colAlign === 'ctr' ? 'ctr'
                : colAlign === 'right' || colAlign === 'r' || (colAlign === '' && numeric) ? 'r' : 'l';
            aligns[cell.col] = aligns[cell.col] === 'r' || align === 'r' ? 'r' : 'l';
            if (cell.colSpan > 1 || cell.rowSpan > 1) merges += 1;
            const link = asString(cell.link, '').trim();
            const linkRelId = link === '' ? undefined : addHyperlink(sink, link, ctx, cell.row + 1, cell.col + 1);
            return tableCellXml(cell, {
                font: ctx.font,
                text: cell.text,
                // runs 原样透传：颜色/字号由 run 自己或段落基值决定，这里再注入一次会盖掉链接色
                runs: asArray(cell.runs),
                sizePt,
                bold: cell.bold === undefined ? false : cell.bold === true,
                // 带链接的单元格如果没有显式颜色，用主题的 hlink 色，否则「像链接」这件事只能靠下划线
                color: cell.color ?? (linkRelId === undefined ? pal.text : pal.secondary),
                align,
                fill: isBlankFill(cell.fill) ? zebraFill : cell.fill,
                insets: options.insets,
                borders,
                notes: ctx.notes,
                page: ctx.page,
                shape: asString(data.title, '表格'),
                formulaErrors: ctx.formulaErrors,
                pal,
                linkRelId,
                underline: linkRelId === undefined ? undefined : true,
                colSpan: cell.colSpan,
                rowSpan: cell.rowSpan,
                bottomLine: rowIndex === rows.length - 1
                    ? { color: pal.tableBorder, width: 12700 }
                    : { color: pal.tableBorder, width: 9525 },
            });
        }).join('');
        return `<a:tr h="${heights[rowIndex + 1]}">${cells}</a:tr>`;
    }).join('');
    const id = sink.nextId();
    sink.tableCount += 1;
    sink.contentArea += g.contentW * total;
    for (const col of cols) sink.words += wordsOf(col.title);
    for (const row of rows) {
        for (const cell of row) if (cell.kind === 'origin') sink.words += wordsOf(cell.text);
    }
    const grid = widths.map((width) => `<a:gridCol w="${width}"/>`).join('');
    sink.push(`<p:graphicFrame><p:nvGraphicFramePr>`
        + `<p:cNvPr id="${id}" name="Table ${sink.tableCount}"/>`
        + `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>`
        // p:nvPr 在 PresentationML 里是必填（minOccurs=1），漏掉 PowerPoint 直接拒绝打开整个文件
        + `<p:nvPr/></p:nvGraphicFramePr>`
        + `<p:xfrm><a:off x="${g.margin}" y="${g.bodyY}"/><a:ext cx="${g.contentW}" cy="${total}"/></p:xfrm>`
        + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">`
        + `<a:tbl><a:tblPr firstRow="1" bandRow="1"/>`
        + `<a:tblGrid>${grid}</a:tblGrid>`
        + `<a:tr h="${heights[0]}">${header}</a:tr>${body}`
        + `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
    return {
        title: asString(data.title),
        layout: { ...layout, aligns },
        columns: cols.length,
        rows: rows.length,
        compressed,
        merges,
    };
}

/** 表格里的超链接：外部链接必须在关系上写 TargetMode="External"，否则 PowerPoint 会去找包内部件。 */
function addHyperlink(sink, url, ctx, row, col) {
    if (!/^(https?:\/\/|mailto:|#)/i.test(url)) {
        ctx.notes.push(`第 ${ctx.page} 页表格第 ${row} 行第 ${col} 列的超链接「${url}」不像 URL，仍按要求写入`);
    }
    const id = sink.nextRelId();
    sink.rels.push({ id, type: RT.hyperlink, target: url, mode: 'External' });
    return id;
}

function renderQuote(sink, ctx, data) {
    const { g, pal } = ctx;
    const text = asString(data.text);
    const by = asString(data.by);
    const runs = asArray(data.runs);
    const textW = g.contentW - cmToEmu(0.9);
    const textY = Math.round(g.cy * 0.33);
    const textH = Math.round(g.cy * 0.36);
    const fit = fitParas(
        (size) => [{
            text, runs, sizePt: size, bold: false, color: pal.text, lineSpacing: 1.36, align: 'l',
        }],
        textW,
        textH,
        [28, 26, 24, 22, 20, 18],
    );
    addText(sink, ctx, {
        x: g.margin + cmToEmu(0.55), y: Math.round(g.cy * 0.15), w: cmToEmu(3.4), h: cmToEmu(2.6),
        anchor: 'b', name: 'Quote Mark', autofit: false,
        paras: [{ text: '“', sizePt: 110, bold: true, color: pal.quoteMark, lineSpacing: 1 }],
    });
    addRect(sink, {
        x: g.margin, y: textY, w: cmToEmu(0.12), h: textH, fill: pal.accent, name: 'Quote Bar',
    });
    addText(sink, ctx, {
        x: g.margin + cmToEmu(0.72), y: textY, w: textW, h: textH, anchor: 'ctr',
        name: 'Quote Text', autofit: false,
        paras: [{ text, runs, sizePt: fit.sizePt, color: pal.text, lineSpacing: 1.36 }],
    });
    if (by !== '') {
        addText(sink, ctx, {
            x: g.margin + cmToEmu(0.72), y: Math.round(g.cy * 0.715), w: textW, h: cmToEmu(1.1),
            anchor: 't', name: 'Quote By', autofit: false,
            paras: [{ text: `—— ${by}`, sizePt: 13, color: pal.muted, align: 'r', lineSpacing: 1 }],
        });
    }
    addFooter(sink, ctx, ctx.docTitle);
    return { title: '', text, by, fitted: fit };
}

function renderStatement(sink, ctx, data) {
    const { g, pal } = ctx;
    const text = asString(data.text);
    const sub = asString(data.sub);
    const runs = asArray(data.runs);
    // 宣言文字默认 onPrimary（白），那是给深色底准备的；浅色主题的陈述页没有深色底，
    // 白字会直接隐形 —— 深浅按下页有没有深底选色，展示式公式页（statement + math run）
    // 在浅色主题上同样靠这一条才看得见
    const statementColor = ctx.theme.dark === true ? pal.onPrimary : pal.text;
    const subColor = ctx.theme.dark === true ? colorMix(pal.primary, pal.onPrimary, 0.72) : pal.muted;
    const fit = fitParas(
        (size) => [{ text, runs, sizePt: size, bold: true, color: statementColor, lineSpacing: 1.2, align: 'ctr' }],
        g.contentW - cmToEmu(1.2),
        Math.round(g.cy * 0.34),
        [36, 32, 28, 24, 20],
    );
    addRect(sink, {
        x: Math.round(g.cx / 2 - cmToEmu(1.2)), y: Math.round(g.cy * 0.235),
        w: cmToEmu(2.4), h: cmToEmu(0.13), fill: pal.accent, name: 'Statement Accent',
    });
    addText(sink, ctx, {
        x: g.margin, y: Math.round(g.cy * 0.29), w: g.contentW, h: Math.round(g.cy * 0.36),
        anchor: 'ctr', name: 'Statement', autofit: false,
        paras: [{ text, runs, sizePt: fit.sizePt, bold: true, color: statementColor, lineSpacing: 1.2, align: 'ctr' }],
    });
    if (sub !== '') {
        addText(sink, ctx, {
            x: g.margin, y: Math.round(g.cy * 0.68), w: g.contentW, h: cmToEmu(1.2),
            anchor: 't', name: 'Statement Sub', autofit: false,
            paras: [{ text: sub, sizePt: 16, color: subColor, align: 'ctr', lineSpacing: 1.3 }],
        });
    }
    addFooter(sink, ctx, ctx.docTitle, { color: pal.onBandMuted, lineColor: { color: pal.onBandMuted, dash: true } });
    return { title: '', text, sub, fitted: fit };
}

// ───────────────────────────────────────────────────────────────────────────
// 图片版式：尺寸计算（contain / cover / natural）、裁切、对齐、多图网格
// ───────────────────────────────────────────────────────────────────────────

/**
 * 图片尺寸与位置的全部决定都收敛在这里，因为 stats / outline / warnings 三个出口
 * 必须报同一个数字：一旦渲染和报告各算一套，报告立刻变成谎话。
 * 非法输入（未知 fit、负数尺寸、非数字）在这里只能落成 warning 或回落值，不许抛。
 *
 * @returns {{ok:boolean, warn?:string, w?:number, h?:number, x?:number, y?:number,
 *            crop?:{l:number,t:number,r:number,b:number}, downscaled?:boolean}}
 */
function layoutImage(info, spec, frame) {
    const fullBleed = spec.fullBleed === true;
    // 铺满整页时不认页边距，也不给标题区留位置：背景图要的是整张画布。
    // 满页用 cx/cy，版心用 contentW/h —— 两种调用方给的是不同的几何字段，这里统一取
    const areaW = fullBleed ? (frame.cx ?? frame.contentW) : frame.contentW;
    const areaH = fullBleed ? (frame.cy ?? frame.h) : frame.h;
    if (!(areaW > 0) || !(areaH > 0)) return { ok: false, warn: '可用区域为 0，图片未放置' };

    const naturalW = info.width * EMU_PER_PX;
    const naturalH = info.height * EMU_PER_PX;
    if (!(naturalW > 0) || !(naturalH > 0)) return { ok: false, warn: '图片像素尺寸非法（宽或高为 0），已画占位框' };
    const aspect = naturalW / naturalH;

    // 显式尺寸只给一个时按比例推另一个；两个都给就完全按用户说的来
    let explicitW = spec.widthCm > 0 ? cmToEmu(spec.widthCm) : 0;
    let explicitH = spec.heightCm > 0 ? cmToEmu(spec.heightCm) : 0;
    if (explicitW > 0 && explicitH <= 0) explicitH = Math.round(explicitW / aspect);
    if (explicitH > 0 && explicitW <= 0) explicitW = Math.round(explicitH * aspect);

    let warning;
    if (explicitW > 0 || explicitH > 0) {
        // 显式尺寸是用户的明确意图，因此不再按 fit 重新摆布，只保证它落在可用区域里
        if (explicitW > areaW || explicitH > areaH) {
            const shrink = Math.min(areaW / explicitW, areaH / explicitH);
            warning = `图片指定尺寸 ${sizeText(explicitW / 360000, explicitH / 360000)} 超出可用区域 `
                + `${sizeText(areaW / 360000, areaH / 360000)}，已整体缩小`;
            [explicitW, explicitH] = [Math.round(explicitW * shrink), Math.round(explicitH * shrink)];
        }
        const x = fullBleed ? 0 : alignX(frame, explicitW, spec.align);
        const y = fullBleed ? 0 : Math.round(frame.y + (areaH - explicitH) / 2);
        return { ok: true, w: explicitW, h: explicitH, x, y, warn: warning };
    }

    // 满页背景默认按 cover：背景图的语义就是「铺满整页」，留白会露出底色，那不是背景图。
    // 想留白仍然可以显式写 fit:'contain'，这里只改默认值，不改用户明确给的选择
    if (spec.fit === 'cover' || (fullBleed && spec.fit !== 'contain')) {
        // 铺满：把「等比放大到刚好盖住区域」的那个矩形记下来，超出的部分交给 a:srcRect 裁
        const s = Math.max(areaW / naturalW, areaH / naturalH);
        const dw = naturalW * s;
        const dh = naturalH * s;
        const crop = cropOf(dw, dh, areaW, areaH);
        const x = fullBleed ? 0 : Math.round(frame.x + (areaW - areaW) / 2);
        const y = fullBleed ? 0 : Math.round(frame.y + (areaH - areaH) / 2);
        return { ok: true, w: areaW, h: areaH, x, y, crop, downscaled: dw > naturalW };
    }

    // contain 与 natural 的缩放规则完全一致：两者都「只缩不放」，差别只在 natural 的
    // 基准是 96 DPI 折算出的原始厘米数。这里用同一个 s，报告里的尺寸才只有一个来源。
    const s = Math.min(areaW / naturalW, areaH / naturalH);
    // 缩放整条链路用同一个 s，w/h 才不会因为两次独立取整而破坏纵横比
    const w = Math.max(9525, Math.round(naturalW * s));
    const h = Math.max(9525, Math.round(naturalH * s));
    const x = fullBleed ? 0 : alignX(frame, w, spec.align);
    const y = fullBleed ? 0 : Math.round(frame.y + (areaH - h) / 2);
    return { ok: true, w, h, x, y, downscaled: s < 1 };
}

/** 图片宽度小于可用区域时按 align 决定水平位置；纵向一律居中（版心上下留白对称最稳）。 */
function alignX(frame, w, align) {
    if (align === 'left') return frame.x;
    if (align === 'right') return Math.round(frame.x + frame.contentW - w);
    return Math.round(frame.x + (frame.contentW - w) / 2);
}

/**
 * cover 的裁切量：在「等比放大后的矩形」里取居中的一块区域，剩下四边转成 a:srcRect
 * 的千分之一百分比。两个细节：
 * - 差值小于 0.5% 时算作 0：整页背景图与页面纵横比常常只差千分之几，写一个 14/100000
 *   的裁切既看不出来，又会平白触发「被裁切」告警。
 * - 刻意不把 srcRect 写满 100000（那等于把图裁没了），封顶 99999。
 */
function cropOf(drawnW, drawnH, areaW, areaH) {
    const percent = (part, whole) => {
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

/** 图片的展示尺寸信息：outline 与 stats 都要报，口径只留这一处。 */
function imageLabel(info) {
    if (info.missing > 0 && info.images === 0) return '图片缺失';
    if (info.images === 0) return '图片格式不支持';
    return `${info.fit ?? 'contain'}｜${sizeText(info.widthCm, info.heightCm)}`
        + `${info.cropped > 0 ? '（有裁切）' : ''}`;
}

/** fit 合法性检查：非法值不能抛，记一条 warning 后回落 contain。 */
function imageFit(raw, notes = []) {
    const value = asString(raw, '').trim().toLowerCase();
    if (value === '') return 'contain';
    if (IMAGE_FITS.includes(value)) return value;
    notes.push(`未知的图片适配方式「${asString(raw)}」，已按 contain 处理（可用：contain / cover / natural）`);
    return 'contain';
}

function imageAlign(raw, notes = []) {
    const value = asString(raw, '').trim().toLowerCase();
    if (value === '' || value === 'left' || value === 'center' || value === 'right') {
        return value === '' ? 'center' : value;
    }
    notes.push(`未知的图片对齐方式「${asString(raw)}」，已按 center 处理（可用：left / center / right）`);
    return 'center';
}

/** 媒体条目是否是「占位」：missing（没有文件）或 unsupported（不是 PNG/JPEG/GIF）。 */
function placeholderOf(media) {
    return media !== undefined && media !== null
        && (media.missing === true || media.unsupported === true);
}

/** 缺图/坏图时不参与缩放计算，直接占满可用区域，让虚线框明显可见。 */
function placeholderBox(x, y, w, h, media) {
    return { x, y, w, h, crop: undefined, info: media };
}

/** 图片在区域内的排版结果 → cNvPr 之外的形状 XML；missing 时画虚线占位框。 */
function placeheldXml(sink, ctx, info, box) {
    const { pal } = ctx;
    const id = sink.nextId();
    // 路径挂在 info 上：box 只是几何，缺图时把「缺的是哪张」写进占位框才有排查价值
    const path = asString(info?.path);
    const text = info?.missing === true
        ? `图片缺失：${path}`
        : `图片格式不支持（只认 PNG/JPEG/GIF）：${path}`;
    sink.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Missing Image"/>`
        + `<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
        + shapeProps({
            x: box.x, y: box.y, w: box.w, h: box.h,
            fill: pal.surface, line: { color: pal.subtle, dash: true }, lineW: 12700,
        })
        + textBodyXml([{ text, sizePt: 14, color: pal.muted, align: 'ctr', lineSpacing: 1.3 }],
            ctx.font, { anchor: 'ctr', autofit: false })
        + '</p:sp>');
    return id;
}

/**
 * 画一张图（或一个缺图占位框）。frame 为 true 时给图片加细边框 —— 只对 contain/natural
 * 有意义：cover 的图片已经铺满整个框，边框只会被裁掉一半。
 * counts=false 用于满页背景：它占满整页、不算「一张版心内的图」，否则 images 会被重复计数。
 */
function addPicture(sink, ctx, box, counts = true) {
    const { pal } = ctx;
    if (counts) {
        sink.imageCount += 1;
        sink.contentArea += box.w * box.h;
    }
    if (placeholderOf(box.info)) {
        return placeheldXml(sink, ctx, box.info, box);
    }
    const info = box.info;
    const id = sink.nextId();
    const rel = { id: sink.nextRelId(), type: RT.image, target: `../media/${info.partName}` };
    sink.rels.push(rel);
    const covered = box.crop !== undefined && box.crop.l > 0;
    const line = box.frame === true && !covered ? { line: pal.border, lineW: 9525 } : {};
    const name = box.name ?? (box.fullBleed === true ? `Picture Full Bleed ${id}` : `Picture ${sink.imageCount}`);
    // alt 写进 cNvPr 的 descr：这是 PowerPoint「设置替代文字」用的同一个字段，
    // 图片墙的语义信息（alt）因此不丢，且不需要新增 rel 或额外形状
    const altAttr = asString(box.alt) === '' ? '' : ` descr="${escapeAttr(asString(box.alt))}"`;
    sink.push(`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"${altAttr}/>`
        + `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
        + `<p:blipFill><a:blip r:embed="${rel.id}"/>${srcRectXml(box.crop)}`
        + '<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        + shapeProps({ x: box.x, y: box.y, w: box.w, h: box.h, ...line })
        + '</p:pic>');
    return id;
}

/** 图片页共用的收尾：题注、DPI 复算、报告字段。 */
function finishImagePage(sink, ctx, data, opts) {
    const { g } = ctx;
    const caption = asString(data.caption);
    const boxes = opts.boxes;
    const missing = boxes.filter((box) => box.info !== undefined && box.info.missing === true).length;
    // 被裁切 = 真的发生了 cover 裁切；只写「cover」但没裁不报，否则每张铺满图都会误告警
    const cropped = boxes.filter((box) => box.crop !== undefined
        && (box.crop.l > 0 || box.crop.t > 0)).length;
    const area = boxes.length === 0
        ? { x: g.margin, y: g.bodyY, w: g.contentW, h: g.bodyH }
        : {
            x: Math.min(...boxes.map((box) => box.x)),
            y: Math.min(...boxes.map((box) => box.y)),
            w: Math.max(...boxes.map((box) => box.x + box.w)) - Math.min(...boxes.map((box) => box.x)),
            h: Math.max(...boxes.map((box) => box.y + box.h)) - Math.min(...boxes.map((box) => box.y)),
        };
    let dpi = 0;
    if (area.w > 0 && area.h > 0) {
        // DPI 按「图片原图像素 / 它在纸面上的物理尺寸」算：这里的尺寸就是刚才算出来的那个，
        // 所以报告里的 DPI 与 PowerPoint 真正渲染出来的清晰度是同一个数
        let total = 0;
        let counted = 0;
        for (const box of boxes) {
            if (placeholderOf(box.info)) continue;
            total += box.info.width / ((box.w / 360000) / 2.54);
            counted += 1;
        }
        dpi = counted === 0 ? 0 : Math.round(total / counted);
    }
    if (opts.captionAt !== null && caption !== '') {
        addText(sink, ctx, {
            x: g.margin, y: opts.captionAt, w: g.contentW, h: cmToEmu(1.0),
            anchor: 't', name: 'Caption', autofit: false,
            paras: [{ text: caption, sizePt: 11, color: opts.captionColor, align: 'ctr', lineSpacing: 1.2 }],
        });
    }
    return {
        title: asString(data.title),
        caption,
        missing,
        // 缺的是哪张：warning 里要写清楚，不然「有 1 张图片缺失」等于没说
        missingPaths: boxes
            .filter((box) => box.info !== undefined && box.info.missing === true)
            .map((box) => asString(box.info.path)),
        placeheld: boxes.filter((box) => box.info !== undefined && box.info.unsupported === true).length,
        fullBleed: opts.fullBleed === true,
        fit: opts.fit ?? 'contain',
        cropped,
        widthCm: cm1(area.w / 360000),
        heightCm: cm1(area.h / 360000),
        grid: opts.grid,
        columns: opts.columns ?? 1,
        items: boxes.length,
        dpi,
    };
}

/** 单图版式：contain / cover / natural + 显式尺寸 + 对齐 + 边框 + 满页背景图。 */
function renderImage(sink, ctx, data) {
    const { g, pal } = ctx;
    const spec = data.spec ?? {};
    // 用户有没有明确写 fit：满页背景在「没写」时默认 cover（背景图必须铺满），
    // 写了 contain 就尊重用户想留白的选择
    const fitGiven = asString(spec.fit, '').trim() !== '';
    const fit = imageFit(spec.fit, ctx.notes);
    const align = imageAlign(spec.align, ctx.notes);
    const fullBleed = spec.fullBleed === true;
    const frame = spec.frame === true;
    // 非法尺寸只记警告、当没给：NaN / 负数都不能让「只给宽度」的等比推算拿到脏输入。
    // 必须区分「没给」（undefined/null，合法）与「给了个负数」（非法）
    const gaveW = spec.widthCm !== undefined && spec.widthCm !== null && spec.widthCm !== '';
    const gaveH = spec.heightCm !== undefined && spec.heightCm !== null && spec.heightCm !== '';
    if ((gaveW && !(asNumber(spec.widthCm, Number.NaN) >= 0))
        || (gaveH && !(asNumber(spec.heightCm, Number.NaN) >= 0))) {
        ctx.notes.push('图片的 widthCm / heightCm 必须是大于 0 的数字，已忽略该尺寸');
    }
    const media = ctx.media;
    const boxSpec = {
        // 满页背景没写 fit 时按 cover 算：留白会露出底色，那就不叫背景图了
        fit: fullBleed && !fitGiven ? 'cover' : fit,
        align,
        widthCm: gaveW ? Math.max(0, asNumber(spec.widthCm, 0)) : 0,
        heightCm: gaveH ? Math.max(0, asNumber(spec.heightCm, 0)) : 0,
        fullBleed,
    };
    // 缺图时占位框要带上原始路径，否则报告里只剩一句「图片缺失」而不知缺的是哪张
    const missingBox = (x, y, w, h) => placeholderBox(x, y, w, h, {
        ...(media ?? {}),
        path: asString(data.path),
    });

    if (fullBleed) {
        // 满页背景：铺满整页、不要标题带与页脚；标题/题注改成压在图上、带半透明底衬的浮层，
        // 否则深色主题的浅字压在浅色照片上会完全读不出来
        const box = placeholderOf(media)
            ? missingBox(0, 0, g.cx, g.cy)
            : mediaBox(media, boxSpec, { x: 0, contentW: g.cx, y: 0, h: g.cy }, ctx);
        if (asString(data.title) !== '') {
            addRect(sink, {
                x: 0, y: 0, w: g.cx, h: cmToEmu(1.7),
                fill: { color: pal.surface, alpha: 72 }, name: 'Full Bleed Title Band',
            });
            addText(sink, ctx, {
                x: g.margin, y: cmToEmu(0.42), w: g.contentW, h: cmToEmu(0.9),
                anchor: 't', name: 'Title', autofit: false,
                paras: [{ text: asString(data.title), sizePt: 20, bold: true, color: pal.primary, lineSpacing: 1.08 }],
            });
        }
        if (asString(data.caption) !== '') {
            addRect(sink, {
                x: 0, y: g.cy - cmToEmu(1.3), w: g.cx, h: cmToEmu(1.3),
                fill: { color: pal.surface, alpha: 72 }, name: 'Full Bleed Caption Band',
            });
            addText(sink, ctx, {
                x: g.margin, y: g.cy - cmToEmu(1.12), w: g.contentW, h: cmToEmu(0.8),
                anchor: 't', name: 'Caption', autofit: false,
                paras: [{ text: asString(data.caption), sizePt: 11, color: pal.text, align: 'ctr', lineSpacing: 1.2 }],
            });
        }
        // counts=false：满页背景是「背景」，不参与 stats.images 的图片张数，也不占版心面积，
        // 否则每张背景图都会连带触发「图片过多 / 超过版心」的误告警
        addPicture(sink, ctx, { ...box, frame: false, fullBleed: true }, false);
        return finishImagePage(sink, ctx, data, {
            boxes: [box], captionAt: null, captionColor: pal.muted,
            fullBleed: true, fit: boxSpec.fit, grid: false, columns: 1,
        });
    }

    const caption = asString(data.caption);
    const captionH = caption === '' ? 0 : cmToEmu(1.0);
    const areaH = Math.max(cmToEmu(3), g.bodyH - captionH - cmToEmu(0.25));
    addContentChrome(sink, ctx, data.title);
    const box = placeholderOf(media)
        ? missingBox(g.margin, g.bodyY, g.contentW, areaH)
        : mediaBox(media, boxSpec, { x: g.margin, contentW: g.contentW, y: g.bodyY, h: areaH }, ctx);
    addPicture(sink, ctx, { ...box, frame });
    return finishImagePage(sink, ctx, data, {
        boxes: [box],
        captionAt: box.y + box.h + cmToEmu(0.18),
        captionColor: pal.muted,
        fullBleed: false,
        fit: boxSpec.fit,
        grid: false,
        columns: 1,
    });
}

/**
 * 把 layoutImage 的纯计算接到真实媒体条目上：拿到「放哪儿、多大、裁多少、原图是谁」。
 * 单独抽出来是因为单图版式与多图网格共用同一套尺寸规则，两边各算一套迟早会不一致。
 * 算不出几何时退回占位框，并把原因写进 warnings —— 静默退化会让报告与画面对不上。
 */
function mediaBox(media, spec, frame, ctx) {
    const box = layoutImage(media, spec, frame);
    if (box.ok !== true) {
        ctx.notes.push(`${asString(media?.path)}：${asString(box.warn, '图片尺寸算不出来')}`);
        return placeholderBox(frame.x, frame.y, frame.contentW ?? frame.cx, frame.h ?? frame.cy, media);
    }
    return { x: box.x, y: box.y, w: box.w, h: box.h, crop: box.crop, info: media };
}

/** 多图网格：等分格宽、逐格等比放入并居中；缺图格子照画占位框，不影响其它格子。 */
function renderImages(sink, ctx, data) {
    const { g, pal } = ctx;
    const items = asArray(data.items);
    const columns = asNumber(data.columns, 2) === 3 ? 3 : 2;
    const gapCm = Math.min(2, Math.max(0, asNumber(data.gapCm, 0.4)));
    const fit = imageFit(data.fit, ctx.notes);
    addContentChrome(sink, ctx, data.title);

    const cells = items.map((item) => ({ item, media: ctx.mediaOf(item.path) }));
    const cellW = Math.round((g.contentW - cmToEmu(gapCm) * (columns - 1)) / columns);
    const rowsCount = Math.max(1, Math.ceil(cells.length / columns));
    // 每行给题注留固定一条：格子高度按「可用高度 / 行数」均分，一行里所有图共用同一个展示框，
    // 视觉上才是「网格」而不是参差不齐的拼贴
    const rowH = Math.round((g.bodyH - cmToEmu(0.25)) / rowsCount);
    const cellH = Math.max(cmToEmu(1.2), rowH - cmToEmu(1.05));
    const boxes = [];
    cells.forEach((cell, index) => {
        const row = Math.floor(index / columns);
        const col = index % columns;
        const x = g.margin + col * (cellW + cmToEmu(gapCm));
        const y = g.bodyY + row * rowH;
        const spec = { fit, align: 'center', widthCm: 0, heightCm: 0, fullBleed: false };
        const box = placeholderOf(cell.media)
            ? placeholderBox(x, y, cellW, cellH, { ...cell.media, path: asString(cell.item.path) })
            : mediaBox(cell.media, spec, { x, contentW: cellW, y, h: cellH }, ctx);
        addPicture(sink, ctx, { ...box, frame: false });
        boxes.push(box);
        if (asString(cell.item.caption) !== '') {
            addText(sink, ctx, {
                x, y: box.y + box.h + cmToEmu(0.12), w: cellW, h: cmToEmu(0.8),
                anchor: 't', name: 'Cell Caption', autofit: false,
                paras: [{ text: asString(cell.item.caption), sizePt: 10, color: pal.muted, align: 'ctr', lineSpacing: 1.15 }],
            });
        }
    });
    return finishImagePage(sink, ctx, data, {
        boxes, captionAt: null, captionColor: pal.muted,
        fullBleed: false, fit, grid: true, columns,
    });
}

function renderClosing(sink, ctx, data) {
    const { g, pal } = ctx;
    const title = asString(data.title ?? '谢谢') || runText(data.runs);
    const subtitle = asString(data.subtitle);
    const runs = asArray(data.runs);
    addRect(sink, {
        x: 0, y: g.cy - cmToEmu(1.1), w: g.cx, h: cmToEmu(1.1), fill: pal.coverBand, name: 'Closing Band',
    });
    const fit = fitParas(
        (size) => [{ text: title, runs, sizePt: size, bold: true, color: pal.coverText, align: 'ctr', lineSpacing: 1.15 }],
        g.contentW - cmToEmu(1.2),
        Math.round(g.cy * 0.18),
        [40, 36, 32, 28],
    );
    addText(sink, ctx, {
        x: g.margin, y: Math.round(g.cy * 0.34), w: g.contentW, h: Math.round(g.cy * 0.18),
        anchor: 'ctr', name: 'Title', autofit: false, title: true,
        paras: [{ text: title, runs, sizePt: fit.sizePt, bold: true, color: pal.coverText, align: 'ctr', lineSpacing: 1.15 }],
    });
    addRect(sink, {
        x: Math.round(g.cx / 2 - cmToEmu(1.2)), y: Math.round(g.cy * 0.565),
        w: cmToEmu(2.4), h: cmToEmu(0.11), fill: pal.accent, name: 'Closing Accent',
    });
    if (subtitle !== '') {
        addText(sink, ctx, {
            x: g.margin, y: Math.round(g.cy * 0.615), w: g.contentW, h: cmToEmu(1.3),
            anchor: 't', name: 'Subtitle', autofit: false,
            paras: [{ text: subtitle, sizePt: 15, color: pal.onCoverMuted, align: 'ctr', lineSpacing: 1.35 }],
        });
    }
    if (ctx.author !== '') {
        addText(sink, ctx, {
            x: g.margin, y: g.cy - cmToEmu(1.0), w: g.contentW, h: cmToEmu(0.75),
            anchor: 'ctr', name: 'Author', autofit: false,
            paras: [{ text: ctx.author, sizePt: 11.5, color: pal.onBandMuted, align: 'ctr', lineSpacing: 1 }],
        });
    }
    return { title, subtitle };
}

// ───────────────────────────────────────────────────────────────────────────
// 版式助手：卡片网格 / 编号流程 / 对比双栏 / KPI / 图文 / 时间线 / 图标网格
// ───────────────────────────────────────────────────────────────────────────

/**
 * 版心（厘米）。
 * 助手页的对外单位是厘米，所以整套布局都先在厘米上算完再统一换算成 EMU：
 * 一次换算只有一次取整，报告与画面读到的就是同一个数，也不会出现 5.666666666666667。
 */
function contentBox(ctx) {
    const { g } = ctx;
    return {
        x: emuToCm(g.margin),
        y: emuToCm(g.bodyY),
        w: emuToCm(g.contentW),
        h: emuToCm(g.bodyH),
        pageW: emuToCm(g.cx),
        pageH: emuToCm(g.cy),
    };
}

/**
 * 一页的几何计划（第五十二轮 `新-6`：版式与配图共存）。
 *
 * 页面上有「没给摆位的画布配图」时按分栏处理：图占版心右侧 `ratio`（默认 0.42），
 * 宿主版式收进剩下的左栏。**只对版式助手页生效**（全部助手的几何都经 `contentBox(ctx)`
 * 取宽度，所以收窄 `ctx.g` 一处就整页生效）；基础版式页与自由绘制页上的配图都要求显式摆位，
 * 没有「版心该收多窄」这个问题可算（见下面的早退）。
 *
 * 自由绘制页（kind='shape'）尤其不能分栏：那一页的形状由调用方自己摆位，
 * 没有「版心」这个概念可收。
 *
 * @param {object} g 版心几何（EMU）
 * @param {object} source 这一页的 slide 数据
 * @returns {{geom: object, rects: Array}} geom 是本页实际用的几何（含 fullContentW），
 *   rects 按 canvas 下标给出各张配图的矩形（厘米），没分栏的位置是 undefined
 */
function planPageGeometry(g, source) {
    // fullContentW / fullContentRight：页面级装饰（标题带、页脚、母版）要的整幅宽度。
    // 分栏只收版心，标题与页脚仍横贯页面 —— 否则半个页面的标题带看着像画错了
    const geom = { ...g, fullContentW: g.contentW, fullContentRight: g.contentRight };
    const rects = [];
    const canvas = asArray(source.canvas);
    // 分栏只在版式助手上发生：基础版式页与自由绘制页上的配图都要求显式摆位，
    // 没有「版心该收多窄」这个问题可算
    if (!HELPER_KINDS.includes(source.kind)) return { geom, rects };
    const at = canvas.findIndex((entry) => entry?.type === 'image' && entry?.spec?.split === true);
    if (at < 0) return { geom, rects };
    const ratio = Math.min(0.7, Math.max(0.2, asNumber(canvas[at].spec.ratio, CANVAS_IMAGE_SPLIT_RATIO)));
    const gap = CANVAS_IMAGE_SPLIT_GAP_CM;
    const boxW = emuToCm(g.contentW);
    const leftW = Math.max(2, boxW * (1 - ratio) - gap / 2);
    geom.contentW = cmToEmu(leftW);
    geom.contentRight = g.margin + geom.contentW;
    rects[at] = {
        x: emuToCm(g.margin) + boxW * (1 - ratio) + gap / 2,
        y: emuToCm(g.bodyY),
        w: Math.max(1, boxW * ratio - gap / 2),
        h: emuToCm(g.bodyH),
    };
    return { geom, rects };
}

/** 计数表自增（stats.shapeKinds 与各助手的条目统计共用）。 */
function bump(map, key, by = 1) {
    map[key] = (map[key] ?? 0) + by;
}

/**
 * 彩色底上的文字色。
 * 序号圆、KPI 卡片这类地方要「在色块上写字」，浅色底上写白字会直接消失，
 * 所以按亮度阈值二选一，而不是一律写白。
 */
function onColorText(value) {
    const hex = toHex(value);
    const parts = [0, 2, 4].map((at) => Number.parseInt(hex.slice(at, at + 2), 16));
    const luma = (0.299 * parts[0] + 0.587 * parts[1] + 0.114 * parts[2]) / 255;
    return luma > 0.62 ? '1F2937' : 'FFFFFF';
}

/**
 * 助手页的卡片底色。
 * 没给就用主题 surface；素色主题的 surface 正是 FFFFFF，而 FFFFFF 按本引擎的
 * 约定等于「不填充」—— 直接写进去卡片会变透明。所以只在这种情况下回落到
 * 一档极浅灰：主题不自己加底，助手又能真的画出一张卡片。
 * 用户显式给的 cardFill 照写（FFFFFF/空串仍然是「不填充」）。
 */
function helperCardFill(raw, ctx) {
    if (raw !== undefined && raw !== null && raw !== false) return shapeFillOf(raw, ctx);
    return { color: isBlankFill(ctx.pal.surface) ? 'F8FAFC' : ctx.pal.surface };
}

/** 助手页的卡片描边：默认一条 0.75pt 的细边，'none' 或 false 表示不描边。 */
function helperCardLine(raw, ctx) {
    if (raw === false || raw === 'none') return undefined;
    if (raw === undefined || raw === null) return { color: ctx.pal.border, widthPt: 0.75 };
    return shapeLineOf(raw, ctx);
}

/**
 * 一页里的默认强调色：在主色 / 副色 / 强调色之间轮转。
 * 多张卡片才不至于全是同一种颜色，用户也不必逐个指定 accent。
 */
function cycleAccent(pal, index) {
    const list = [pal.primary, pal.secondary, pal.accent, pal.primaryDark].filter((value) => !isBlankFill(value));
    return list.length === 0 ? pal.text : list[index % list.length];
}

/**
 * 列数归一。签名里承诺的合法值由 allowed 给出，非法值只记 warning 后回落默认值，
 * 绝不抛 —— 与「非法输入一律 warning」的约定一致。
 */
function helperColumns(raw, fallback, ctx, allowed) {
    if (raw === undefined || raw === null || raw === '') return fallback;
    const value = Math.round(asNumber(raw, fallback));
    if (!allowed.includes(value)) {
        ctx.notes.push(`第 ${ctx.page} 页只支持 ${allowed.join(' / ')} 列，已按 ${fallback} 列排版（收到 ${asString(raw)}）`);
        return fallback;
    }
    return value;
}

/**
 * 越界告警。
 * 负坐标在本引擎里是刻意允许的「出血到页边」，所以只报右/下越出页面，
 * 以及整块完全落在页面之外这两种真正看不见的情况。
 */
function warnOffPage(ctx, what, x, y, w, h) {
    const box = contentBox(ctx);
    if (x + w > box.pageW + 0.05 || y + h > box.pageH + 0.05 || x > box.pageW || y > box.pageH) {
        ctx.notes.push(`第 ${ctx.page} 页 ${what} 超出页面：右 ${cm1(x + w)}cm、下 ${cm1(y + h)}cm，`
            + `页面只有 ${cm1(box.pageW)}×${cm1(box.pageH)}cm`);
    }
}

/**
 * 助手页的 items 归一：字符串条目补成对象，字段交给各渲染函数用 asString/asArray 兜底。
 * 只做这一层，不再给每个助手写一遍 map。
 */
function helperItems(raw) {
    return asArray(raw).map((item) => (item !== null && typeof item === 'object' ? item : { title: asString(item) }));
}

/** 卡片正文的基准字号：内容自适应高度时按它量，放不下仍走原来的字号阶梯。 */
const CARD_BODY_PT = 11;
/** KPI 卡片的最高高度（厘米）：内容就是一个大数字 + 一行标签，撑满整页会变成四张空卡。 */
const KPI_CARD_MAX_CM = 4.6;
/** 卡片头部到正文之间的固定占位（厘米）：pad 0.3 + headH 0.95 + 细线 0.13 + 间距 0.18 + 下边距 0.3。 */
const CARD_CHROME_CM = 1.86;

/**
 * 一张助手卡片的「内容需要多高」（厘米）。
 *
 * 卡片高度过去只由版心与行数决定（cardGrid 均分），内容只有一两行时整张卡
 * 大半是空的 —— 2026-09-22 实测：42 页里 34 页正文区一半以上没字，
 * 卡片页平均七成空白。这里按固定头部 + 正文在基准字号下真实需要的高度算下限，
 * 让卡片跟着内容走，整块再在版心内垂直居中。
 *
 * @param {{body?: unknown, runs?: unknown}} item 卡片数据
 * @param {number} cardW 卡片宽度（厘米）
 * @param {{muted: string}} pal 主题色板
 */
function cardNeedCm(item, cardW, pal) {
    const runs = asArray(item.runs);
    const bodyText = asString(item.body) || runText(runs);
    if (bodyText === '') return CARD_CHROME_CM;
    const fit = fitParas(
        (size) => [{ text: bodyText, runs, sizePt: size, color: pal.muted, lineSpacing: 1.3, align: 'l' }],
        cmToEmu(Math.max(1, cardW - 0.6)),
        cmToEmu(200),   // 只量「需要多高」，不给上限
        [CARD_BODY_PT],
    );
    return CARD_CHROME_CM + emuToCm(fit.neededPt * EMU_PER_PT);
}

/** 图标网格每格的固定留白：上 0.3 + 图标与标签间隙 0.18 + 标签与说明间隙 0.08 + 下 0.25。 */
const ICON_CELL_CHROME_CM = 0.81;
/** 图标网格格高的下限（厘米）：低于此值图标与标签会挤在一起，读者失去层次。 */
const ICON_CELL_MIN_CM = 2.0;
/** 图标网格说明文字的基准字号：内容自适应高度时按它量，放不下仍走原来的字号阶梯。 */
const ICON_CELL_BODY_PT = 11;
/** 时间线节点的基准字号（与字号阶梯的首项一致）：内容自适应高度时按它量。 */
const TIMELINE_TEXT_PT = 12;
/** 时间线相邻节点之间的最小间距（厘米）：再挤下去节点圆与文字会叠在一起。 */
const TIMELINE_MIN_SLOT_CM = 1.6;
/** 时间线一个节点的文字块之外要留的余量（厘米）：上下各 0.2 强，节点圆 0.44 高。 */
const TIMELINE_SLOT_PAD_CM = 0.45;

/**
 * 一个图标网格格子「内容需要多高」（厘米）。
 *
 * 与卡片同源的问题（历史遗留 `3-2`）：格高过去只由版心与行数决定，
 * 三列两行只有图标与标签时每格仍分到 7cm，整块上下各空出一大片。
 * 这里按「固定头部 + 说明在基准字号下真实需要的高度」算下限，
 * 让格子跟着内容走，整块再在版心内垂直居中。
 *
 * @param {{body?: unknown}} item 格子数据
 * @param {number} cellW 格宽（厘米）
 * @param {{muted: string}} pal 主题色板
 * @param {{iconSize: number, labelH: number}} opts 图标边长与标签高度（厘米）
 */
function iconCellNeedCm(item, cellW, pal, opts) {
    const body = asString(item.body);
    const chrome = ICON_CELL_CHROME_CM + opts.iconSize + opts.labelH;
    if (body === '') return chrome;
    const fit = fitParas(
        (size) => [{ text: body, sizePt: size, color: pal.muted, align: 'ctr', lineSpacing: 1.3 }],
        cmToEmu(Math.max(1, cellW - 0.3)),
        cmToEmu(200),   // 只量「需要多高」，不给上限
        [ICON_CELL_BODY_PT],
    );
    return chrome + emuToCm(fit.neededPt * EMU_PER_PT);
}

/**
 * 卡片式助手页的落定几何：每张卡在同一列宽 / 行高网格里。
 *
 * preferH 是内容算出来的需要高度：比均分高度小就按内容收，
 * 返回的 offsetY 让整块网格在版心内垂直居中（不传 preferH 时行为与从前一致）。
 */
function cardGrid(box, count, columns, gap, preferH) {
    const rows = Math.max(1, Math.ceil(count / columns));
    const cardW = (box.w - gap * (columns - 1)) / columns;
    const fillH = (box.h - gap * (rows - 1)) / rows;
    const cardH = preferH === undefined ? fillH : Math.max(1.8, Math.min(fillH, preferH));
    const usedH = cardH * rows + gap * (rows - 1);
    return {
        rows,
        cardW,
        cardH,
        fillH,
        usedH,
        offsetY: Math.max(0, (box.h - usedH) / 2),
    };
}

/** 某一行第一张卡的 x：最后一行不满时整行居中，不再靠左留一个洞。 */
function rowStartX(box, grid, row, count, columns, gap) {
    const inRow = Math.min(columns, count - row * columns);
    const rowW = inRow * grid.cardW + (inRow - 1) * gap;
    return box.x + Math.max(0, (box.w - rowW) / 2);
}

/** 助手页里的一张卡：圆角矩形（几何与填充由调用方按厘米给出）。 */
function addCardBox(sink, ctx, spec) {
    addAutoShape(sink, ctx, {
        preset: 'roundRect',
        name: spec.name,
        x: spec.x,
        y: spec.y,
        w: spec.w,
        h: spec.h,
        fill: spec.fill,
        line: spec.line,
        radius: spec.radius,
        shadow: spec.shadow === true,
    });
}

/**
 * 卡片头部：图标 / 序号圆 + 标题 + 强调细线。
 * 三种装饰的位置都从同一个 headH 推出来，卡片标题的基线才能对齐。
 */
function addCardHead(sink, ctx, opts) {
    const headH = opts.headH ?? 0.95;
    const pad = opts.pad ?? 0.3;
    // 卡片头部自己画的形状数（图标 + 序号圆），由调用方并进该页的 stats.shapeKinds
    const info = { shapes: 0, icons: 0 };
    const item = opts.item ?? {};
    const accent = opts.accent;
    let cursor = opts.x + pad;
    if (item.icon !== undefined && item.icon !== null && item.icon !== '' && opts.iconSize > 0) {
        const size = opts.iconSize;
        const drawn = addIconShape(sink, ctx, {
            glyph: item.icon,
            x: cursor,
            y: opts.y + pad + (headH - size) / 2,
            sizeCm: size,
            color: accent,
            name: `${opts.name} Icon`,
        });
        if (drawn !== undefined) {
            info.shapes += 1;
            info.icons += 1;
            bump(opts.kinds, 'icon');
        }
        cursor += size + 0.22;
    }
    if (opts.numbered === true) {
        const d = 0.78;
        addAutoShape(sink, ctx, {
            preset: 'ellipse',
            name: `${opts.name} Badge`,
            x: cursor,
            y: opts.y + pad + (headH - d) / 2,
            w: d,
            h: d,
            fill: accent,
            text: String(opts.index + 1),
            textOpts: { size: 11, bold: true, color: onColorText(accent), align: 'ctr', anchor: 'ctr' },
        });
        info.shapes += 1;
        bump(opts.kinds, 'ellipse');
        cursor += d + 0.16;
    }
    info.titleX = cursor;
    info.titleW = Math.max(1, opts.x + opts.w - pad - cursor);
    // 强调细线：贴在头部下方，给卡片一个不依赖颜色的视觉锚点
    addRect(sink, {
        x: cmToEmu(opts.x + pad),
        y: cmToEmu(opts.y + pad + headH + 0.06),
        w: cmToEmu(Math.min(1.6, Math.max(0.6, info.titleW))),
        h: cmToEmu(0.07),
        fill: accent,
        name: `${opts.name} Rule`,
    });
    info.bodyY = opts.y + pad + headH + 0.06 + 0.07 + 0.18;
    info.bodyH = opts.y + opts.h - pad - info.bodyY;
    return info;
}

/**
 * 卡片网格（deck.cards）。
 * columns 支持 1~4（默认 2）；每张卡的正文按字号阶梯找「放得下」的那一档，
 * 找不到就带着真实数字进 warnings —— 与版式页共用同一套量算。
 */
function renderCards(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const items = helperItems(data.items);
    const columns = helperColumns(data.columns, 2, ctx, [1, 2, 3, 4]);
    const gap = Math.max(0, asNumber(data.gapCm, 0.5));
    const radius = Math.max(0, asNumber(data.radius, 0.22));
    const numbered = data.numbered === true;
    const iconSize = Math.max(0, asNumber(data.iconSizeCm, 0.9));
    const fill = helperCardFill(data.cardFill, ctx);
    const line = helperCardLine(data.cardLine, ctx);
    // 卡片高度跟着内容走：先按列宽量出需要多高，再让网格按它收高并垂直居中
    const measureW = (box.w - gap * (columns - 1)) / columns;
    const preferH = items.length === 0
        ? undefined
        : Math.max(...items.map((item) => cardNeedCm(item, measureW, pal)));
    const grid = cardGrid(box, Math.max(1, items.length), columns, gap, preferH);
    const info = {
        title: asString(data.title), count: items.length, columns, rows: grid.rows,
        numbered, iconCount: 0, cards: [], kinds: {}, shapes: 0,
        cardW: grid.cardW, cardH: grid.cardH, minCardH: 1.8,
    };
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 cards 没有任何卡片，该页只有标题`);
        return info;
    }
    // 卡片数量 × 网格容量：行高不足 1.8cm 时卡片会被压成一条，读者失去层次
    const maxRows = Math.max(1, Math.floor((box.h + gap) / (info.minCardH + gap)));
    const overflow = grid.rows > maxRows;
    if (overflow) {
        // 网格本身溢出时不再逐张报「正文放不下」：40 张卡就是 40 条重复告警，
        // 真正可操作的那一条（数量超出列数×行数）会被埋掉
        ctx.notes.push(`第 ${ctx.page} 页 cards：${items.length} 张卡片按 ${columns} 列排成 ${grid.rows} 行，`
            + `每张只有 ${cm1(grid.cardH)}cm 高，超出可用高度（最多 ${columns}×${maxRows} = ${columns * maxRows} 张），`
            + '建议减少卡片或分页');
    }
    const sizes = [13, 12, 11, 10.5, 10, 9.5, 9];
    items.forEach((item, index) => {
        const col = index % columns;
        const row = Math.floor(index / columns);
        const x = rowStartX(box, grid, row, items.length, columns, gap) + col * (grid.cardW + gap);
        const y = box.y + grid.offsetY + row * (grid.cardH + gap);
        const accent = toHex(asString(item.accent, '') === '' ? cycleAccent(pal, index) : item.accent);
        const name = `Card ${index + 1}`;
        addCardBox(sink, ctx, {
            name, x, y, w: grid.cardW, h: grid.cardH,
            fill, line, radius, shadow: data.shadow === true,
        });
        info.shapes += 1;
        bump(info.kinds, 'roundRect');
        const head = addCardHead(sink, ctx, {
            name, item, accent, index, x, y, w: grid.cardW, h: grid.cardH,
            headH: 0.95, pad: 0.3, iconSize, numbered, kinds: info.kinds,
        });
        info.shapes += head.shapes;
        info.iconCount += head.icons;
        const titleFit = fitTitle(
            asString(item.title), cmToEmu(head.titleW), cmToEmu(0.95), [15, 14, 13, 12, 11],
        );
        if (!titleFit.fits && !overflow) {
            ctx.notes.push(`第 ${ctx.page} 页 cards 第 ${index + 1} 张卡片标题过长，已压到最小 ${titleFit.sizePt}pt`);
        }
        addText(sink, ctx, {
            x: cmToEmu(head.titleX), y: cmToEmu(y + 0.3), w: cmToEmu(head.titleW), h: cmToEmu(0.95),
            anchor: 'ctr', name: `${name} Title`, autofit: false,
            paras: [{
                text: asString(item.title), sizePt: titleFit.sizePt, bold: true, color: pal.text,
                lineSpacing: 1.1, align: 'l',
            }],
        });
        const runs = asArray(item.runs);
        const bodyText = asString(item.body) || runText(runs);
        const bodyH = Math.max(0.3, head.bodyH);
        const bodyFit = fitParas(
            (size) => [{
                text: bodyText, runs, sizePt: size, color: pal.muted, lineSpacing: 1.3, align: 'l',
            }],
            cmToEmu(Math.max(1, grid.cardW - 0.6)),
            cmToEmu(bodyH),
            sizes,
        );
        if (bodyText !== '' && !bodyFit.fits && !overflow) {
            ctx.notes.push(`第 ${ctx.page} 页 cards 第 ${index + 1} 张卡片正文放不下，`
                + `压缩到最小 ${bodyFit.sizePt}pt 仍需要 ${Math.ceil(bodyFit.neededPt)}pt 高，`
                + `可用只有 ${Math.floor(cmToEmu(bodyH) / EMU_PER_PT)}pt`);
        }
        addText(sink, ctx, {
            x: cmToEmu(x + 0.3), y: cmToEmu(head.bodyY), w: cmToEmu(Math.max(1, grid.cardW - 0.6)),
            h: cmToEmu(bodyH), anchor: 't', name: `${name} Body`, autofit: false,
            paras: [{
                text: bodyText, runs, sizePt: bodyFit.sizePt, color: pal.muted, lineSpacing: 1.3, align: 'l',
            }],
        });
        info.cards.push({ x, y, w: grid.cardW, h: grid.cardH, name, title: asString(item.title) });
    });
    return info;
}

/**
 * 编号流程（deck.steps）。
 * direction 'horizontal'（默认）| 'vertical'；numbered 默认 true（画序号圆），
 * arrows 默认 true（步骤之间用 deck.line 的 a:tailEnd 三角形箭头连起来）。
 */
function renderSteps(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const items = helperItems(data.items);
    const raw = asString(data.direction, 'horizontal').trim().toLowerCase();
    let direction = 'horizontal';
    if (raw === 'vertical') direction = 'vertical';
    else if (raw !== '' && raw !== 'horizontal') {
        ctx.notes.push(`第 ${ctx.page} 页 steps 的 direction「${asString(data.direction)}」不认识，`
            + '已按 horizontal 处理（可用：horizontal / vertical）');
    }
    const numbered = data.numbered !== false;
    const arrows = data.arrows !== false;
    const fill = helperCardFill(data.fill, ctx);
    const line = helperCardLine(data.line, ctx);
    const info = {
        title: asString(data.title), count: items.length, direction, numbered, arrows,
        kinds: {}, shapes: 0, arrowsDrawn: 0, slots: [],
    };
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 steps 没有任何步骤，该页只有标题`);
        return info;
    }
    const gap = Math.max(0.2, asNumber(data.gapCm, direction === 'horizontal' ? 0.9 : 0.8));
    const pad = 0.3;
    const headH = 0.95;
    if (direction === 'horizontal') {
        const stepW = (box.w - gap * (items.length - 1)) / items.length;
        // 高度跟着内容走：原来固定 5.2cm，正文只有一两行时卡片里大半是空的
        const needH = Math.max(...items.map((item) => cardNeedCm(item, stepW, pal)));
        const stepH = Math.max(2.6, Math.min(box.h, needH));
        const y = box.y + (box.h - stepH) / 2;
        if (stepW < 3.2) {
            ctx.notes.push(`第 ${ctx.page} 页 steps：${items.length} 步横向排布每步只有 ${cm1(stepW)}cm 宽，`
                + `建议减到 ${Math.max(1, Math.floor((box.w + gap) / (3.2 + gap)))} 步以内或改成 vertical`);
        }
        items.forEach((item, index) => {
            const x = box.x + index * (stepW + gap);
            const name = `Step Card ${index + 1}`;
            addCardBox(sink, ctx, { name, x, y, w: stepW, h: stepH, fill, line, radius: 0.22 });
            info.shapes += 1;
            bump(info.kinds, 'roundRect');
            const head = addCardHead(sink, ctx, {
                name, item, accent: cycleAccent(pal, index), index, x, y, w: stepW, h: stepH,
                headH, pad, iconSize: 0, numbered, kinds: info.kinds,
            });
            info.shapes += head.shapes;
            addText(sink, ctx, {
                x: cmToEmu(head.titleX), y: cmToEmu(y + pad), w: cmToEmu(head.titleW), h: cmToEmu(headH),
                anchor: 'ctr', name: `${name} Title`, autofit: false,
                paras: [{
                    text: asString(item.title), sizePt: 14, bold: true, color: pal.text,
                    lineSpacing: 1.1, align: 'l',
                }],
            });
            const bodyH = Math.max(0.3, head.bodyH);
            const fit = fitParas(
                (size) => [{
                    text: asString(item.body), sizePt: size, color: pal.muted, lineSpacing: 1.3, align: 'l',
                }],
                cmToEmu(Math.max(1, stepW - pad * 2)),
                cmToEmu(bodyH),
                [12, 11, 10.5, 10, 9.5, 9],
            );
            if (asString(item.body) !== '' && !fit.fits) {
                ctx.notes.push(`第 ${ctx.page} 页 steps 第 ${index + 1} 步正文放不下，已压到最小 ${fit.sizePt}pt`);
            }
            addText(sink, ctx, {
                x: cmToEmu(x + pad), y: cmToEmu(head.bodyY), w: cmToEmu(Math.max(1, stepW - pad * 2)),
                h: cmToEmu(bodyH), anchor: 't', name: `${name} Body`, autofit: false,
                paras: [{
                    text: asString(item.body), sizePt: fit.sizePt, color: pal.muted, lineSpacing: 1.3, align: 'l',
                }],
            });
            info.slots.push({ x, y, w: stepW, h: stepH });
            if (arrows && index < items.length - 1) {
                addLineShape(sink, ctx, {
                    from: [x + stepW + 0.1, y + stepH / 2],
                    to: [x + stepW + gap - 0.1, y + stepH / 2],
                    color: pal.subtle, widthPt: 1.25, arrow: 'end',
                    name: `Step Arrow ${index + 1}`,
                });
                info.shapes += 1;
                info.arrowsDrawn += 1;
                bump(info.kinds, 'line');
            }
        });
        return info;
    }
    const arrowH = Math.max(0.5, asNumber(data.gapCm, 0.7));
    const stepH = (box.h - arrowH * (items.length - 1)) / items.length;
    if (stepH < 1.5) {
        ctx.notes.push(`第 ${ctx.page} 页 steps：${items.length} 步纵向排布每步只有 ${cm1(stepH)}cm 高，`
            + `建议减到 ${Math.max(1, Math.floor((box.h + arrowH) / (1.5 + arrowH)))} 步以内或分页`);
    }
    items.forEach((item, index) => {
        const y = box.y + index * (stepH + arrowH);
        const name = `Step Card ${index + 1}`;
        addCardBox(sink, ctx, { name, x: box.x, y, w: box.w, h: stepH, fill, line, radius: 0.22 });
        info.shapes += 1;
        bump(info.kinds, 'roundRect');
        const head = addCardHead(sink, ctx, {
            name, item, accent: cycleAccent(pal, index), index, x: box.x, y, w: box.w, h: stepH,
            headH, pad, iconSize: 0, numbered, kinds: info.kinds,
        });
        info.shapes += head.shapes;
        addText(sink, ctx, {
            x: cmToEmu(head.titleX), y: cmToEmu(y + pad), w: cmToEmu(head.titleW), h: cmToEmu(headH),
            anchor: 'ctr', name: `${name} Title`, autofit: false,
            paras: [{
                text: asString(item.title), sizePt: 14, bold: true, color: pal.text,
                lineSpacing: 1.1, align: 'l',
            }],
        });
        const bodyH = Math.max(0.3, head.bodyH);
        const fit = fitParas(
            (size) => [{
                text: asString(item.body), sizePt: size, color: pal.muted, lineSpacing: 1.3, align: 'l',
            }],
            cmToEmu(Math.max(1, box.w - pad * 2)),
            cmToEmu(bodyH),
            [12, 11, 10.5, 10, 9.5, 9],
        );
        if (asString(item.body) !== '' && !fit.fits) {
            ctx.notes.push(`第 ${ctx.page} 页 steps 第 ${index + 1} 步正文放不下，已压到最小 ${fit.sizePt}pt`);
        }
        addText(sink, ctx, {
            x: cmToEmu(box.x + pad), y: cmToEmu(head.bodyY), w: cmToEmu(Math.max(1, box.w - pad * 2)),
            h: cmToEmu(bodyH), anchor: 't', name: `${name} Body`, autofit: false,
            paras: [{
                text: asString(item.body), sizePt: fit.sizePt, color: pal.muted, lineSpacing: 1.3, align: 'l',
            }],
        });
        info.slots.push({ x: box.x, y, w: box.w, h: stepH });
        if (arrows && index < items.length - 1) {
            addLineShape(sink, ctx, {
                from: [box.x + box.w / 2, y + stepH + 0.1],
                to: [box.x + box.w / 2, y + stepH + arrowH - 0.1],
                color: pal.subtle, widthPt: 1.25, arrow: 'end',
                name: `Step Arrow ${index + 1}`,
            });
            info.shapes += 1;
            info.arrowsDrawn += 1;
            bump(info.kinds, 'line');
        }
    });
    return info;
}

/** 对比双栏（deck.compare）：左右各一张卡 + （默认）中间一条竖分隔线。 */
function renderCompare(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const gap = Math.max(0.3, asNumber(data.gapCm, 0.7));
    const dividerRaw = data.divider;
    const showDivider = dividerRaw !== false && dividerRaw !== 'none' && dividerRaw !== null;
    const dividerColor = typeof dividerRaw === 'string' && dividerRaw !== '' && dividerRaw !== 'none'
        ? toHex(dividerRaw)
        : pal.border;
    const fill = helperCardFill(data.cardFill, ctx);
    const line = helperCardLine(data.cardLine, ctx);
    const colW = (box.w - gap) / 2;
    const info = {
        title: asString(data.title), count: 2, columns: 2, divider: showDivider,
        kinds: {}, shapes: 0, sides: [],
    };
    const sides = [
        { key: 'left', spec: data.left, x: box.x },
        { key: 'right', spec: data.right, x: box.x + colW + gap },
    ];
    const pad = 0.3;
    // 两栏卡片高度跟着条目走，整块在版心内垂直居中（原来一律撑满 box.h）
    const measureCol = (spec) => {
        const list = helperItems(spec !== null && typeof spec === 'object' ? spec.items : undefined);
        if (list.length === 0) return 1.58 + pad;
        const fit = fitParas(
            (size) => list.map((item, index) => ({
                text: asString(item.text ?? item.title), runs: asArray(item.runs), sizePt: size,
                color: pal.text, marL: cmToEmu(0.42), indent: -cmToEmu(0.42), bullet: '•',
                bulletColor: pal.text, lineSpacing: 1.3, spaceBeforePt: index === 0 ? 0 : 8,
            })),
            cmToEmu(colW - pad * 2),
            cmToEmu(200),
            [12],
        );
        return 1.58 + emuToCm(fit.neededPt * EMU_PER_PT) + pad;
    };
    const needH = Math.max(measureCol(data.left), measureCol(data.right));
    const cardH = Math.max(2.6, Math.min(box.h, needH));
    const cardY = box.y + Math.max(0, (box.h - cardH) / 2);
    for (const side of sides) {
        const spec = side.spec !== null && typeof side.spec === 'object' ? side.spec : {};
        const items = helperItems(spec.items);
        const accent = toHex(asString(spec.accent, '') === '' ? cycleAccent(pal, side.key === 'left' ? 0 : 1) : spec.accent);
        const name = `Compare Column ${side.key === 'left' ? 'Left' : 'Right'}`;
        addCardBox(sink, ctx, {
            name, x: side.x, y: cardY, w: colW, h: cardH, fill, line, radius: 0.22,
        });
        info.shapes += 1;
        bump(info.kinds, 'roundRect');
        // 标题用强调色 + 一条强调细线：不画整条色带，色带是方角的，压在圆角卡上会露边
        const title = asString(spec.title);
        const titleFit = fitTitle(title, cmToEmu(colW - pad * 2), cmToEmu(0.95), [16, 15, 14, 13, 12]);
        addText(sink, ctx, {
            x: cmToEmu(side.x + pad), y: cmToEmu(cardY + 0.28), w: cmToEmu(colW - pad * 2), h: cmToEmu(0.95),
            anchor: 'ctr', name: `${name} Title`, autofit: false,
            paras: [{ text: title, sizePt: titleFit.sizePt, bold: true, color: accent, lineSpacing: 1.1, align: 'l' }],
        });
        addRect(sink, {
            x: cmToEmu(side.x + pad), y: cmToEmu(cardY + 1.3), w: cmToEmu(1.8), h: cmToEmu(0.07),
            fill: accent, name: `${name} Rule`,
        });
        const bodyY = cardY + 1.58;
        const bodyH = cardH - 1.58 - pad;
        const parasOf = (size) => items.map((item, index) => ({
            text: asString(item.text ?? item.title),
            runs: asArray(item.runs),
            sizePt: size,
            color: pal.text,
            marL: cmToEmu(0.42),
            indent: -cmToEmu(0.42),
            bullet: '•',
            bulletColor: accent,
            lineSpacing: 1.3,
            spaceBeforePt: index === 0 ? 0 : 8,
        }));
        const fit = fitParas(parasOf, cmToEmu(colW - pad * 2), cmToEmu(Math.max(0.3, bodyH)), [14, 13, 12, 11, 10.5, 10]);
        if (items.length > 0 && !fit.fits) {
            ctx.notes.push(`第 ${ctx.page} 页 compare ${side.key === 'left' ? '左' : '右'}栏 ${items.length} 条要点放不下，`
                + `已压到最小 ${fit.sizePt}pt，建议精简条目或缩短文字`);
        }
        addText(sink, ctx, {
            x: cmToEmu(side.x + pad), y: cmToEmu(bodyY), w: cmToEmu(colW - pad * 2),
            h: cmToEmu(Math.max(0.3, bodyH)), anchor: 't', name: `${name} Items`, autofit: false,
            paras: fit.paras,
        });
        info.sides.push({ key: side.key, title, items: items.length, accent, x: side.x, w: colW, fits: fit.fits });
    }
    if (showDivider) {
        addLineShape(sink, ctx, {
            from: [box.x + colW + gap / 2, cardY + 0.1],
            to: [box.x + colW + gap / 2, cardY + cardH - 0.1],
            color: dividerColor, widthPt: 1,
            name: 'Compare Divider',
        });
        info.shapes += 1;
        info.dividerX = box.x + colW + gap / 2;
        bump(info.kinds, 'line');
    }
    return info;
}

/** KPI 大数字（deck.kpi）：大号数字 + 单位 + 小标签，一行或两行铺开。 */
function renderKpi(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const items = helperItems(data.items);
    const fallback = Math.min(4, Math.max(1, items.length));
    const columns = helperColumns(data.columns, fallback, ctx, [1, 2, 3, 4]);
    const gap = Math.max(0, asNumber(data.gapCm, 0.5));
    const fill = helperCardFill(data.cardFill, ctx);
    const line = helperCardLine(data.cardLine, ctx);
    // 收高 + 垂直居中：与 cards 同一套处理（KPI 内容固定，用一个上限即可）
    const grid = cardGrid(box, Math.max(1, items.length), columns, gap, KPI_CARD_MAX_CM);
    const info = {
        title: asString(data.title), count: items.length, columns, rows: grid.rows,
        kinds: {}, shapes: 0, values: [],
    };
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 kpi 没有任何指标，该页只有标题`);
        return info;
    }
    const pad = 0.3;
    items.forEach((item, index) => {
        const col = index % columns;
        const row = Math.floor(index / columns);
        const x = rowStartX(box, grid, row, items.length, columns, gap) + col * (grid.cardW + gap);
        const y = box.y + grid.offsetY + row * (grid.cardH + gap);
        const color = toHex(asString(item.color, '') === '' ? cycleAccent(pal, index) : item.color);
        const name = `KPI Card ${index + 1}`;
        addCardBox(sink, ctx, { name, x, y, w: grid.cardW, h: grid.cardH, fill, line, radius: 0.22 });
        info.shapes += 1;
        bump(info.kinds, 'roundRect');
        addRect(sink, {
            x: cmToEmu(x + grid.cardW / 2 - 0.9), y: cmToEmu(y + pad), w: cmToEmu(1.8), h: cmToEmu(0.07),
            fill: color, name: `${name} Rule`,
        });
        const value = asString(item.value);
        const unit = asString(item.unit);
        const valueH = Math.max(1.2, grid.cardH * 0.52);
        const fit = fitParas(
            (size) => [{
                sizePt: size, bold: true, color, lineSpacing: 1, align: 'ctr',
                runs: unit === ''
                    ? [{ text: value, sizePt: size, bold: true, color }]
                    : [
                        { text: value, sizePt: size, bold: true, color },
                        { text: unit, sizePt: Math.max(12, Math.round(size * 0.42)), bold: false, color: pal.muted },
                    ],
            }],
            cmToEmu(grid.cardW - pad * 2),
            cmToEmu(valueH),
            [54, 48, 42, 36, 30, 26, 22, 18],
        );
        if (fit.sizePt < 24) {
            ctx.notes.push(`第 ${ctx.page} 页 kpi 第 ${index + 1} 个数字「${value}」过长，已压到 ${fit.sizePt}pt`);
        }
        if (!fit.fits) {
            ctx.notes.push(`第 ${ctx.page} 页 kpi 第 ${index + 1} 个数字放不下，可用高度只有 ${cm1(valueH)}cm`);
        }
        addText(sink, ctx, {
            x: cmToEmu(x + pad), y: cmToEmu(y + pad + 0.18), w: cmToEmu(grid.cardW - pad * 2),
            h: cmToEmu(valueH), anchor: 'b', name: `${name} Value`, autofit: false,
            paras: fit.paras,
        });
        const label = asString(item.label);
        const labelFit = fitTitle(label, cmToEmu(grid.cardW - pad * 2), cmToEmu(0.8), [14, 13, 12, 11]);
        addText(sink, ctx, {
            x: cmToEmu(x + pad), y: cmToEmu(y + grid.cardH - pad - 0.8), w: cmToEmu(grid.cardW - pad * 2),
            h: cmToEmu(0.8), anchor: 'ctr', name: `${name} Label`, autofit: false,
            paras: [{ text: label, sizePt: labelFit.sizePt, color: pal.muted, align: 'ctr', lineSpacing: 1.2 }],
        });
        info.values.push({ value, unit, label, sizePt: fit.sizePt, x, y, w: grid.cardW, h: grid.cardH });
    });
    return info;
}

/**
 * 图 + 文（deck.imageText）。
 * side 'left'（默认）| 'right' 决定图片在哪一侧，ratio 是图片占版心宽的比例（默认 0.46）。
 * 文本侧永远画一块浅底面板：文字直接落在白底上会与图片失去分界。
 */
function renderImageText(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const image = data.image !== null && typeof data.image === 'object' ? data.image : {};
    const path = asString(image.path ?? data.path);
    const rawSide = asString(data.side, 'left').trim().toLowerCase();
    let side = 'left';
    if (rawSide === 'right') side = 'right';
    else if (rawSide !== '' && rawSide !== 'left') {
        ctx.notes.push(`第 ${ctx.page} 页 imageText 的 side「${asString(data.side)}」不认识，`
            + '已按 left 处理（可用：left / right）');
    }
    const ratioRaw = asNumber(data.ratio, 0.46);
    const ratio = Math.min(0.8, Math.max(0.2, ratioRaw));
    if (ratioRaw !== ratio) {
        ctx.notes.push(`第 ${ctx.page} 页 imageText 的 ratio ${ratioRaw} 超出 0.2~0.8，已按 ${ratio} 处理`);
    }
    const gap = 0.7;
    const imgW = box.w * ratio;
    const textW = box.w - imgW - gap;
    const imgX = side === 'left' ? box.x : box.x + textW + gap;
    const textX = side === 'left' ? box.x + imgW + gap : box.x;
    const caption = asString(image.caption);
    const captionH = caption === '' ? 0 : 0.95;
    const areaH = Math.max(1.5, box.h - captionH - 0.15);
    const fit = imageFit(image.fit, ctx.notes);
    const media = ctx.mediaOf(path);
    if (path === '') ctx.notes.push(`第 ${ctx.page} 页 imageText 缺少图片 path，已画成占位框`);
    const spec = { fit, align: 'center', widthCm: 0, heightCm: 0, fullBleed: false };
    const picBox = placeholderOf(media)
        ? placeholderBox(cmToEmu(imgX), cmToEmu(box.y), cmToEmu(imgW), cmToEmu(areaH), { ...(media ?? {}), path })
        : mediaBox(media, spec, {
            x: cmToEmu(imgX), contentW: cmToEmu(imgW), y: cmToEmu(box.y), h: cmToEmu(areaH),
        }, ctx);
    addPicture(sink, ctx, { ...picBox, name: 'ImageText Picture' });
    if (caption !== '') {
        addText(sink, ctx, {
            x: cmToEmu(imgX), y: picBox.y + picBox.h + cmToEmu(0.12), w: cmToEmu(imgW), h: cmToEmu(0.8),
            anchor: 't', name: 'ImageText Caption', autofit: false,
            paras: [{ text: caption, sizePt: 10.5, color: pal.muted, align: 'ctr', lineSpacing: 1.2 }],
        });
    }
    // 文本侧：浅底面板 + 要点（没有 items 就写整段文字）
    addAutoShape(sink, ctx, {
        preset: 'roundRect', name: 'ImageText Panel',
        x: textX, y: box.y, w: textW, h: box.h,
        fill: helperCardFill(data.cardFill, ctx), line: helperCardLine(data.cardLine, ctx), radius: 0.22,
    });
    const pad = 0.32;
    const items = helperItems(data.items);
    const innerW = Math.max(1, textW - pad * 2);
    const innerH = Math.max(0.5, box.h - pad * 2);
    let paras;
    let sizePt;
    if (items.length > 0) {
        const parasOf = (size) => items.map((item, index) => ({
            text: asString(item.text ?? item.title),
            runs: asArray(item.runs),
            sizePt: size,
            color: pal.text,
            marL: cmToEmu(0.42),
            indent: -cmToEmu(0.42),
            bullet: '•',
            bulletColor: pal.accent,
            lineSpacing: 1.32,
            spaceBeforePt: index === 0 ? 0 : 8,
        }));
        const fitted = fitParas(parasOf, cmToEmu(innerW), cmToEmu(innerH), [15, 14, 13, 12, 11, 10.5]);
        paras = fitted.paras;
        sizePt = fitted.sizePt;
        if (!fitted.fits) {
            ctx.notes.push(`第 ${ctx.page} 页 imageText 的要点放不下，已压到最小 ${fitted.sizePt}pt，建议减少条目`);
        }
    } else {
        const text = asString(data.text);
        const fitted = fitParas(
            (size) => [{ text, sizePt: size, color: pal.text, lineSpacing: 1.38, align: 'l' }],
            cmToEmu(innerW),
            cmToEmu(innerH),
            [15, 14, 13, 12, 11, 10.5],
        );
        paras = fitted.paras;
        sizePt = fitted.sizePt;
        if (text !== '' && !fitted.fits) {
            ctx.notes.push(`第 ${ctx.page} 页 imageText 的文字放不下，已压到最小 ${fitted.sizePt}pt`);
        }
    }
    addText(sink, ctx, {
        x: cmToEmu(textX + pad), y: cmToEmu(box.y + pad), w: cmToEmu(innerW), h: cmToEmu(innerH),
        anchor: 't', name: 'ImageText Text', autofit: false, paras,
    });
    const base = finishImagePage(sink, ctx, { title: asString(data.title), caption }, {
        boxes: [picBox], captionAt: null, captionColor: pal.muted,
        fullBleed: false, fit, grid: false, columns: 1,
    });
    return {
        ...base,
        side, ratio, items: items.length, sizePt, text: items.length > 0 ? '' : asString(data.text),
        count: items.length > 0 ? items.length : (asString(data.text) === '' ? 0 : 1),
        shapes: 1, kinds: { roundRect: 1 },
    };
}

/**
 * 时间线（deck.timeline）：一条轴线 + 等距节点圆 + 上下交错的时间/标题/正文。
 * 交错是为了让相邻节点在纵向上不打架：横向 n 个节点平分版心宽，每格已经很窄了。
 */
/** 时间线一个节点的三段文字（时间 / 标题 / 正文）：量高度与真正落盘共用同一份。 */
function timelineParasOf(item, size, pal) {
    return [
        { text: asString(item.time), sizePt: Math.max(9, size - 1), bold: true, color: pal.accent, lineSpacing: 1.15, align: 'l' },
        { text: asString(item.title), sizePt: size + 1, bold: true, color: pal.text, lineSpacing: 1.15, align: 'l', spaceBeforePt: 3 },
        { text: asString(item.body), sizePt: Math.max(9, size - 1), color: pal.muted, lineSpacing: 1.3, align: 'l', spaceBeforePt: 3 },
    ];
}

/**
 * 纵向时间线：一个节点在基准字号下真正需要多高（厘米）。
 *
 * 与卡片 / 图标网格同源的问题（历史遗留 `3-2`）：节点间距过去一律 = 版心高 / 节点数，
 * 两个节点时会拉出一根 14cm 长的轴线、中间一大片空。这里按内容量算间距，
 * 量出来的总高不超过版心时整块在版心内垂直居中，超过才退回「贯穿版心」。
 *
 * @param {{time?: unknown, title?: unknown, body?: unknown}} item 节点数据
 * @param {number} blockW 文字块宽度（厘米）
 * @param {{accent: string, text: string, muted: string}} pal 主题色板
 */
function timelineTextNeedCm(item, blockW, pal) {
    const fit = fitParas(
        (size) => timelineParasOf(item, size, pal),
        cmToEmu(Math.max(1, blockW)),
        cmToEmu(200),   // 只量「需要多高」，不给上限
        [TIMELINE_TEXT_PT],
    );
    return emuToCm(fit.neededPt * EMU_PER_PT);
}

function renderTimeline(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const items = helperItems(data.items);
    const raw = asString(data.direction, 'horizontal').trim().toLowerCase();
    let direction = 'horizontal';
    if (raw === 'vertical') direction = 'vertical';
    else if (raw !== '' && raw !== 'horizontal') {
        ctx.notes.push(`第 ${ctx.page} 页 timeline 的 direction「${asString(data.direction)}」不认识，`
            + '已按 horizontal 处理（可用：horizontal / vertical）');
    }
    const info = {
        title: asString(data.title), count: items.length, direction,
        kinds: {}, shapes: 0, nodes: [], slot: 0,
    };
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 timeline 没有任何节点，该页只有标题`);
        return info;
    }
    const d = 0.44;
    // 三段文字的样式只有一处定义：量高度与落盘都走 timelineParasOf
    const parasOf = (item, size) => timelineParasOf(item, size, pal);
    if (direction === 'horizontal') {
        const slot = box.w / items.length;
        info.slot = slot;
        const midY = box.y + box.h / 2;
        addLineShape(sink, ctx, {
            from: [box.x, midY], to: [box.x + box.w, midY],
            color: pal.border, widthPt: 1.5, name: 'Timeline Axis',
        });
        info.shapes += 1;
        bump(info.kinds, 'line');
        // 节点太密时只报一条：每格宽度不够是同一个原因，逐节点重复报没有信息量
        const tight = slot < 3.2;
        if (tight) {
            ctx.notes.push(`第 ${ctx.page} 页 timeline：${items.length} 个节点横向排布每格只有 ${cm1(slot)}cm 宽，`
                + `建议减到 ${Math.max(1, Math.floor(box.w / 3.2))} 个以内或分页`);
        }
        items.forEach((item, index) => {
            const nodeX = box.x + slot * (index + 0.5);
            addAutoShape(sink, ctx, {
                preset: 'ellipse', name: `Timeline Node ${index + 1}`,
                x: nodeX - d / 2, y: midY - d / 2, w: d, h: d,
                fill: cycleAccent(pal, index), line: { color: pal.bg, widthPt: 1.5 },
            });
            info.shapes += 1;
            bump(info.kinds, 'ellipse');
            const above = index % 2 === 0;
            const blockW = Math.min(slot - 0.3, 6.4);
            const blockX = Math.min(Math.max(nodeX - blockW / 2, box.x), box.x + box.w - blockW);
            const blockH = Math.max(0.6, (above ? midY - 0.35 - box.y : box.y + box.h - midY - 0.35));
            const blockY = above ? box.y : midY + 0.35;
            const fit = fitParas(
                (size) => parasOf(item, size),
                cmToEmu(blockW),
                cmToEmu(blockH),
                [12, 11, 10, 9.5, 9],
            );
            if (!fit.fits && !tight) {
                ctx.notes.push(`第 ${ctx.page} 页 timeline 第 ${index + 1} 个节点文字放不下，`
                    + `已压到最小 ${fit.sizePt}pt，建议精简该节点的正文`);
            }
            addText(sink, ctx, {
                x: cmToEmu(blockX), y: cmToEmu(blockY), w: cmToEmu(blockW), h: cmToEmu(blockH),
                anchor: above ? 'b' : 't', name: `Timeline Text ${index + 1}`, autofit: false,
                paras: fit.paras,
            });
            info.nodes.push({ index, x: nodeX, y: midY, above, sizePt: fit.sizePt, block: { x: blockX, y: blockY, w: blockW, h: blockH } });
        });
        return info;
    }
    const axisX = box.x + 2.4;
    const blockX = axisX + 0.8;
    const blockW = Math.max(1, box.x + box.w - blockX);
    // 轴线长度跟着内容走（3-2）：不超过版心时整块在版心内垂直居中，超过才贯穿版心。
    // 槽高取**各节点需要的最大值**（与 cardGrid 同一口径），不是均值 ——
    // 均值会让「一个重节点 + 一个轻节点」里的重节点拿不到自己需要的高度：
    // 复核实测 500 字节点会从 12pt 无告警退化成 9pt + 「放不下」（旧公式反而够用）。
    // 取最大值后：内容轻时整块收高居中，内容重到装不下时退回版心均分（与旧行为一致，只会更宽裕）。
    const needs = items.map((item) => Math.max(
        TIMELINE_MIN_SLOT_CM,
        timelineTextNeedCm(item, blockW, pal) + TIMELINE_SLOT_PAD_CM,
    ));
    const slot = Math.min(box.h / items.length, Math.max(...needs));
    const usedH = slot * items.length;
    const startY = box.y + (box.h - usedH) / 2;
    info.slot = slot;
    info.usedH = usedH;
    info.startY = startY;
    addLineShape(sink, ctx, {
        from: [axisX, startY], to: [axisX, startY + usedH],
        color: pal.border, widthPt: 1.5, name: 'Timeline Axis',
    });
    info.shapes += 1;
    bump(info.kinds, 'line');
    const tight = slot < 1.8;
    if (tight) {
        ctx.notes.push(`第 ${ctx.page} 页 timeline：${items.length} 个节点纵向排布每格只有 ${cm1(slot)}cm 高，建议分页`);
    }
    items.forEach((item, index) => {
        const nodeY = startY + slot * (index + 0.5);
        addAutoShape(sink, ctx, {
            preset: 'ellipse', name: `Timeline Node ${index + 1}`,
            x: axisX - d / 2, y: nodeY - d / 2, w: d, h: d,
            fill: cycleAccent(pal, index), line: { color: pal.bg, widthPt: 1.5 },
        });
        info.shapes += 1;
        bump(info.kinds, 'ellipse');
        const blockH = Math.max(0.6, slot - 0.2);
        const blockY = nodeY - blockH / 2;
        const fit = fitParas(
            (size) => parasOf(item, size),
            cmToEmu(blockW),
            cmToEmu(blockH),
            [12, 11, 10, 9.5, 9],
        );
        if (!fit.fits && !tight) {
            ctx.notes.push(`第 ${ctx.page} 页 timeline 第 ${index + 1} 个节点文字放不下，`
                + `已压到最小 ${fit.sizePt}pt，建议精简该节点的正文`);
        }
        addText(sink, ctx, {
            x: cmToEmu(blockX), y: cmToEmu(blockY), w: cmToEmu(blockW), h: cmToEmu(blockH),
            anchor: 'ctr', name: `Timeline Text ${index + 1}`, autofit: false, paras: fit.paras,
        });
        info.nodes.push({ index, x: axisX, y: nodeY, above: false, sizePt: fit.sizePt, block: { x: blockX, y: blockY, w: blockW, h: blockH } });
    });
    return info;
}

/**
 * 图标网格（deck.iconGrid）。
 * icon 走 deck.icon 的字形通道（默认字体 Segoe MDL2 Assets），
 * 一格 = 图标 + 标签 + 说明，横竖都居中。
 */
function renderIconGrid(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const items = helperItems(data.items);
    const columns = helperColumns(data.columns, 3, ctx, [1, 2, 3, 4, 5, 6]);
    // 间隙先钳一次：gapCm 没有上限时（既有边界，分栏让窄版心更容易命中）格宽会被压到 0，
    // 标签与说明的文本框宽度变成负数、被 shapeProps 钳到 1 EMU —— 文字直接看不见且不告警。
    // 留够每格 1.5cm 的可用宽度，超出就钳回来并如实报一句。
    const wantedGap = Math.max(0.2, asNumber(data.gapCm, 0.6));
    const gapMax = columns > 1 ? Math.max(0.2, (box.w - columns * 1.5) / (columns - 1)) : wantedGap;
    const gap = Math.min(wantedGap, gapMax);
    if (gap < wantedGap) {
        ctx.notes.push(`第 ${ctx.page} 页 iconGrid 的 gapCm ${cm1(wantedGap)}cm 太大`
            + `（${columns} 列排进 ${cm1(box.w)}cm 的版心会把格子压没），已按 ${cm1(gap)}cm 处理`);
    }
    const iconSize = Math.max(0.4, asNumber(data.iconSizeCm, 1.1));
    const rows = Math.max(1, Math.ceil(items.length / columns));
    const cellW = (box.w - gap * (columns - 1)) / columns;
    const labelH = 0.8;
    // 格高跟着内容走（3-2）：先按格宽量出每格需要多高，再让整块按它收高并垂直居中。
    // 量不出来（没有格子）或量出来比均分高度还高时仍按均分高度，行为与从前一致。
    const fillH = (box.h - gap * (rows - 1)) / rows;
    const needH = items.length === 0
        ? undefined
        : Math.max(...items.map((item) => iconCellNeedCm(item, cellW, pal, { iconSize, labelH })));
    // 内容下限只用来「抬」太矮的格子，绝不能超过均分高度 ——
    // 否则格子会互相压住、整块溢出页面（与 cardGrid 同一个口径：上限永远是 fillH）
    const cellH = needH === undefined ? fillH : Math.min(fillH, Math.max(ICON_CELL_MIN_CM, needH));
    const usedH = cellH * rows + gap * (rows - 1);
    const offsetY = Math.max(0, (box.h - usedH) / 2);
    const info = {
        title: asString(data.title), count: items.length, columns, rows,
        kinds: {}, shapes: 0, icons: 0, cellW, cellH, fillH, offsetY,
    };
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 iconGrid 没有任何图标，该页只有标题`);
        return info;
    }
    const dense = cellH < ICON_CELL_MIN_CM;
    if (dense) {
        ctx.notes.push(`第 ${ctx.page} 页 iconGrid：${items.length} 个图标排成 ${columns}×${rows}，`
            + `每格只有 ${cm1(cellH)}cm 高，建议减少图标或分页`);
    }
    items.forEach((item, index) => {
        const col = index % columns;
        const row = Math.floor(index / columns);
        const x = box.x + col * (cellW + gap);
        const y = box.y + offsetY + row * (cellH + gap);
        const name = `IconGrid ${index + 1}`;
        const accent = cycleAccent(pal, index);
        const drawn = addIconShape(sink, ctx, {
            glyph: item.icon,
            x: x + (cellW - iconSize) / 2,
            y: y + 0.3,
            sizeCm: iconSize,
            color: accent,
            name: `${name} Icon`,
        });
        if (drawn !== undefined) {
            info.icons += 1;
            info.shapes += 1;
            bump(info.kinds, 'icon');
        }
        const labelY = y + 0.3 + iconSize + 0.18;
        const label = asString(item.label ?? item.title);
        const labelFit = fitTitle(label, cmToEmu(cellW - 0.3), cmToEmu(labelH), [14, 13, 12, 11, 10]);
        addText(sink, ctx, {
            x: cmToEmu(x + 0.15), y: cmToEmu(labelY), w: cmToEmu(cellW - 0.3), h: cmToEmu(labelH),
            anchor: 'ctr', name: `${name} Label`, autofit: false,
            paras: [{ text: label, sizePt: labelFit.sizePt, bold: true, color: pal.text, align: 'ctr', lineSpacing: 1.15 }],
        });
        const bodyY = labelY + labelH + 0.08;
        const bodyH = Math.max(0.3, y + cellH - 0.25 - bodyY);
        const body = asString(item.body);
        const fit = fitParas(
            (size) => [{ text: body, sizePt: size, color: pal.muted, align: 'ctr', lineSpacing: 1.3 }],
            cmToEmu(cellW - 0.3),
            cmToEmu(bodyH),
            [11, 10.5, 10, 9.5, 9],
        );
        if (body !== '' && !fit.fits && !dense) {
            ctx.notes.push(`第 ${ctx.page} 页 iconGrid 第 ${index + 1} 格说明放不下，已压到最小 ${fit.sizePt}pt`);
        }
        addText(sink, ctx, {
            x: cmToEmu(x + 0.15), y: cmToEmu(bodyY), w: cmToEmu(cellW - 0.3), h: cmToEmu(bodyH),
            anchor: 't', name: `${name} Body`, autofit: false,
            paras: [{ text: body, sizePt: fit.sizePt, color: pal.muted, align: 'ctr', lineSpacing: 1.3 }],
        });
    });
    return info;
}

/**
 * 目录页（deck.agenda）：编号条目纵向排列，当前章节高亮。
 * 每行 = 圆角条 + 序号圆（numbered:false 时改画左侧强调竖条）+ 标题（+ 可选说明）。
 * active 的那一行整行换成强调色、文字取反色；只认第一条 active —— 多条高亮等于没有高亮。
 */
function renderAgenda(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const items = helperItems(data.items).map((item) => ({
        text: asString(item.text ?? item.title),
        runs: asArray(item.runs),
        desc: asString(item.desc ?? item.body),
        active: item.active === true,
    }));
    // 顶层 active（1 起的序号）是给「只标当前章节」这种最常见写法的快捷方式
    if (!items.some((item) => item.active) && Number.isInteger(data.active) && data.active > 0
        && data.active <= items.length) {
        items[data.active - 1].active = true;
    }
    const numbered = data.numbered !== false;
    const gap = Math.max(0.1, asNumber(data.gapCm, 0.28));
    const rows = Math.max(1, items.length);
    const fillH = (box.h - gap * (rows - 1)) / rows;
    // 行高上限 2.4cm：条目少时不要拉成一排空荡荡的色带，整块垂直居中
    const rowH = Math.min(2.4, fillH);
    const usedH = rowH * rows + gap * (rows - 1);
    const offsetY = Math.max(0, (box.h - usedH) / 2);
    const activeAt = items.findIndex((item) => item.active);
    const info = {
        title: asString(data.title), count: items.length, rows, numbered, rowH,
        active: activeAt < 0 ? 0 : activeAt + 1, kinds: {}, shapes: 0,
    };
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 agenda 没有任何条目，该页只有标题`);
        return info;
    }
    if (rows > 10 || rowH < 1.05) {
        ctx.notes.push(`第 ${ctx.page} 页 agenda：${items.length} 条按版心高 ${cm1(box.h)}cm 排成 ${rows} 行，`
            + `每行只有 ${cm1(rowH)}cm 高，建议精简到 10 条以内或拆成两页`);
    }
    if (activeAt < 0) {
        // 目录没有「当前章节」是合法状态，但要让调用方知道没有行被高亮
        ctx.notes.push(`第 ${ctx.page} 页 agenda 没有标记 active 的条目，全部按普通条目排版`);
    }
    items.forEach((item, index) => {
        const y = box.y + offsetY + index * (rowH + gap);
        const active = index === activeAt;
        const accent = active ? pal.accent : cycleAccent(pal, index);
        const name = `Agenda Row ${index + 1}`;
        addCardBox(sink, ctx, {
            name,
            x: box.x, y, w: box.w, h: rowH,
            fill: active ? accent : helperCardFill(data.cardFill, ctx),
            line: active ? accent : helperCardLine(data.cardLine, ctx),
            radius: 0.16,
        });
        info.shapes += 1;
        bump(info.kinds, 'roundRect');
        const pad = 0.42;
        let textX = box.x + pad;
        if (numbered) {
            const d = Math.max(0.5, Math.min(0.82, rowH - 0.46));
            addAutoShape(sink, ctx, {
                preset: 'ellipse', name: `${name} Badge`,
                x: box.x + pad, y: y + (rowH - d) / 2, w: d, h: d,
                // 素色主题的 surface 是 FFFFFF，而本引擎把 FFFFFF 当「不填充」：
                // 直接写会把圆整个变透明、只剩 595959 的序号浮在同样 595959 的
                // 高亮行上（实测肉眼不可见）。active 行改回落到一档极浅灰。
                fill: active ? (isBlankFill(pal.surface) ? 'F8FAFC' : pal.surface) : accent,
                text: String(index + 1),
                textOpts: {
                    size: Math.max(9, Math.round(d * 15)), bold: true,
                    color: active ? accent : onColorText(accent), align: 'ctr', anchor: 'ctr',
                },
            });
            info.shapes += 1;
            bump(info.kinds, 'ellipse');
            textX = box.x + pad + d + 0.34;
        } else {
            // 不编号时用左侧一条竖色条建立节奏：既省掉序号圆，也保住「一条条」的视觉
            addRect(sink, {
                name: `${name} Mark`,
                x: cmToEmu(box.x + pad), y: cmToEmu(y + 0.18),
                w: cmToEmu(0.1), h: cmToEmu(Math.max(0.3, rowH - 0.36)),
                fill: accent,
            });
        }
        const textW = Math.max(2, box.w - (textX - box.x) - 0.5);
        const textH = Math.max(0.4, rowH - (item.desc === '' ? 0.32 : 0.26));
        const color = active ? onColorText(accent) : pal.text;
        const descColor = active ? onColorText(accent) : pal.muted;
        const fit = fitParas(
            (size) => {
                const paras = [{
                    text: item.text, runs: item.runs, sizePt: size, bold: true,
                    color, lineSpacing: 1.12,
                }];
                if (item.desc !== '') {
                    paras.push({
                        text: item.desc, sizePt: Math.max(9, Math.round(size * 0.62)),
                        color: descColor, lineSpacing: 1.2, spaceBeforePt: 4,
                    });
                }
                return paras;
            },
            cmToEmu(textW),
            cmToEmu(textH),
            [22, 20, 19, 18, 17, 16, 15, 14, 13],
        );
        addText(sink, ctx, {
            x: cmToEmu(textX), y: cmToEmu(y + 0.11), w: cmToEmu(textW), h: cmToEmu(textH),
            anchor: 'ctr', name: `Agenda Text ${index + 1}`, autofit: false, paras: fit.paras,
        });
    });
    return info;
}

/**
 * 图片墙（deck.gallery）：等格图片网 + 图注 + 角标。
 * 图片一律走 imageText 那套（imageFit / ctx.mediaOf / placeholderBox / mediaBox），
 * 所以相对路径、缺图占位、DPI 复算与 deck.image 是同一口径，这里不另写解析。
 */
function renderGallery(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const items = asArray(data.items).map((entry) => (
        entry !== null && typeof entry === 'object' ? entry : { path: asString(entry) }
    ));
    const columns = helperColumns(data.columns, 2, ctx, [1, 2, 3]);
    // gap 与 iconGrid 同口径：没有上限时 3 列窄版心会把格子压到 0 宽，先钳一次并如实报
    const wantedGap = Math.max(0.2, asNumber(data.gapCm, 0.5));
    const gapMax = columns > 1 ? Math.max(0.2, (box.w - columns * 3) / (columns - 1)) : wantedGap;
    const gap = Math.min(wantedGap, gapMax);
    if (gap < wantedGap) {
        ctx.notes.push(`第 ${ctx.page} 页 gallery 的 gapCm ${cm1(wantedGap)}cm 太大`
            + `（${columns} 列排进 ${cm1(box.w)}cm 的版心会把格子压没），已按 ${cm1(gap)}cm 处理`);
    }
    const ratioRaw = asNumber(data.ratio, 0.62);
    const ratio = Math.min(1.5, Math.max(0.2, ratioRaw));
    if (ratio !== ratioRaw) {
        ctx.notes.push(`第 ${ctx.page} 页 gallery 的 ratio ${ratioRaw} 超出 0.2~1.5，已按 ${ratio} 处理`);
    }
    const fit = imageFit(data.fit, ctx.notes);
    const rows = Math.max(1, Math.ceil(items.length / columns));
    const cellW = (box.w - gap * (columns - 1)) / columns;
    const fillH = (box.h - gap * (rows - 1)) / rows;
    const hasCaption = items.some((item) => asString(item.caption ?? item.alt) !== '');
    const captionH = hasCaption ? 0.72 : 0;
    // 图框高 = 格宽 × ratio（宽高比指的是高/宽），但不能吃穿均分行高
    const imgH = Math.max(1.0, Math.min(cellW * ratio, fillH - captionH - 0.28));
    const cellH = imgH + captionH + 0.28;
    const usedH = cellH * rows + gap * (rows - 1);
    const offsetY = Math.max(0, (box.h - usedH) / 2);
    const info = {
        title: asString(data.title), count: items.length, items: items.length,
        columns, rows, ratio, fit, cellW, imgH, cellH, images: 0, kinds: {}, shapes: 0,
    };
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 gallery 没有任何图片，该页只有标题`);
        return info;
    }
    if (cellH > fillH + 0.01 || items.length > 12) {
        ctx.notes.push(`第 ${ctx.page} 页 gallery：${items.length} 张图按 ${columns} 列排成 ${rows} 行，`
            + `每格只有 ${cm1(fillH)}cm 高，建议减少图片、调小 ratio 或分页`);
    }
    const boxes = [];
    const noPath = [];
    items.forEach((item, index) => {
        const col = index % columns;
        const row = Math.floor(index / columns);
        const x = box.x + col * (cellW + gap);
        const y = box.y + offsetY + row * (cellH + gap);
        const name = `Gallery Cell ${index + 1}`;
        // image 允许写成 {path, alt} 对象（与 deck.image 同形），也认 path / src 简写
        const imageObj = item.image !== null && typeof item.image === 'object' ? item.image : {};
        const path = asString(imageObj.path ?? item.path ?? item.src ?? item.image);
        const caption = withCredit(asString(item.caption ?? item.alt), item.source);
        const badge = asString(item.badge);
        const accent = cycleAccent(pal, index);
        addCardBox(sink, ctx, {
            name, x, y, w: cellW, h: cellH,
            fill: helperCardFill(data.cardFill, ctx), line: helperCardLine(data.cardLine, ctx), radius: 0.16,
        });
        info.shapes += 1;
        bump(info.kinds, 'roundRect');
        if (path === '') noPath.push(index + 1);
        const media = ctx.mediaOf(path);
        const frameX = x + 0.14;
        const frameY = y + 0.14;
        const frameW = Math.max(1, cellW - 0.28);
        const picBox = placeholderOf(media)
            ? placeholderBox(cmToEmu(frameX), cmToEmu(frameY), cmToEmu(frameW), cmToEmu(imgH), { ...(media ?? {}), path })
            : mediaBox(media, { fit, align: 'center', widthCm: 0, heightCm: 0, fullBleed: false }, {
                x: cmToEmu(frameX), contentW: cmToEmu(frameW), y: cmToEmu(frameY), h: cmToEmu(imgH),
            }, ctx);
        // alt 写进图片的替代文字（a:cNvPr descr）：这是 PowerPoint 里「设置替代文字」的同一个位置
        addPicture(sink, ctx, { ...picBox, name: `${name} Picture`, alt: asString(item.alt ?? imageObj.alt) });
        boxes.push(picBox);
        if (badge !== '') {
            const badgeW = Math.min(Math.max(1.0, cellW - 0.6), Math.max(1.1, 0.3 * badge.length + 0.5));
            addAutoShape(sink, ctx, {
                preset: 'roundRect', name: `${name} Badge`,
                x: x + cellW - badgeW - 0.24, y: y + 0.26, w: badgeW, h: 0.48,
                fill: accent, radius: 0.18,
                text: badge,
                textOpts: { size: 10, bold: true, color: onColorText(accent), align: 'ctr', anchor: 'ctr' },
            });
            info.shapes += 1;
            bump(info.kinds, 'roundRect');
        }
        if (caption !== '') {
            addText(sink, ctx, {
                x: cmToEmu(x + 0.16), y: cmToEmu(y + 0.14 + imgH + 0.05),
                w: cmToEmu(Math.max(1, cellW - 0.32)), h: cmToEmu(Math.max(0.3, captionH - 0.06)),
                anchor: 't', name: `${name} Caption`, autofit: false,
                paras: [{ text: caption, sizePt: 10, color: pal.muted, align: 'ctr', lineSpacing: 1.2 }],
            });
        }
    });
    info.images = boxes.length;
    if (noPath.length > 0) {
        ctx.notes.push(`第 ${ctx.page} 页 gallery 有 ${noPath.length} 张图片缺少 path（第 ${noPath.join('、')} 张），`
            + '已画成占位框');
    }
    const base = finishImagePage(sink, ctx, { title: asString(data.title), caption: '' }, {
        boxes, captionAt: null, captionColor: pal.muted, fullBleed: false, fit, grid: true, columns,
    });
    return { ...base, ...info };
}

/**
 * 阶段路线图（deck.roadmap）：横向 3~5 个阶段卡，每卡 = 强调条 + 阶段名 + 要点 + 可选完成度条。
 * 阶段多过 5 个时按网格换行（不丢阶段、不抛），并报一条「每列会被压窄」的告警。
 */
function renderRoadmap(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const badProgress = [];
    const phases = asArray(data.phases).map((entry, index) => {
        const phase = entry !== null && typeof entry === 'object' ? entry : { label: asString(entry) };
        const raw = phase.progress === undefined || phase.progress === null ? undefined : asNumber(phase.progress, 0);
        if (raw !== undefined && (raw < 0 || raw > 1)) badProgress.push(index + 1);
        return {
            label: asString(phase.label ?? phase.title),
            items: asArray(phase.items).map((item) => asString(item)),
            progress: raw === undefined ? undefined : Math.min(1, Math.max(0, raw)),
            accent: asString(phase.accent),
        };
    });
    const total = phases.length;
    const columns = Math.max(1, Math.min(5, total));
    const rows = Math.max(1, Math.ceil(total / columns));
    const gap = Math.max(0.25, asNumber(data.gapCm, 0.45));
    const cellW = (box.w - gap * (columns - 1)) / columns;
    const cellH = (box.h - gap * (rows - 1)) / rows;
    const info = {
        title: asString(data.title), count: total, columns, rows, phases: phases.map((phase) => phase.label),
        kinds: {}, shapes: 0,
    };
    if (total === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 roadmap 没有任何阶段，该页只有标题`);
        return info;
    }
    if (badProgress.length > 0) {
        ctx.notes.push(`第 ${ctx.page} 页 roadmap 有 ${badProgress.length} 个阶段的 progress 超出 0~1`
            + `（第 ${badProgress.join('、')} 个），已钳到边界`);
    }
    if (total > 5) {
        ctx.notes.push(`第 ${ctx.page} 页 roadmap：${total} 个阶段排成 ${columns} 列 ${rows} 行，`
            + `每列只有 ${cm1(cellW)}cm 宽，建议精简到 3~5 个阶段或拆成多页`);
    }
    if (cellH < 3) {
        ctx.notes.push(`第 ${ctx.page} 页 roadmap：${total} 个阶段每格只有 ${cm1(cellH)}cm 高，`
            + '建议减少每阶段要点或拆成两页');
    }
    let overflowed = 0;
    phases.forEach((phase, index) => {
        const col = index % columns;
        const row = Math.floor(index / columns);
        const x = box.x + col * (cellW + gap);
        const y = box.y + row * (cellH + gap);
        const name = `Roadmap Phase ${index + 1}`;
        const accent = phase.accent === '' ? cycleAccent(pal, index) : phase.accent;
        addCardBox(sink, ctx, {
            name, x, y, w: cellW, h: cellH,
            fill: helperCardFill(data.cardFill, ctx), line: helperCardLine(data.cardLine, ctx), radius: 0.2,
        });
        info.shapes += 1;
        bump(info.kinds, 'roundRect');
        // 顶部强调条：不依赖颜色也能看出「一阶段一卡」的分段
        addRect(sink, {
            name: `${name} Strip`,
            x: cmToEmu(x + 0.28), y: cmToEmu(y + 0.3),
            w: cmToEmu(Math.max(0.5, cellW - 0.56)), h: cmToEmu(0.1), fill: accent,
        });
        const badgeD = 0.72;
        addAutoShape(sink, ctx, {
            preset: 'ellipse', name: `${name} Badge`,
            x: x + 0.28, y: y + 0.56, w: badgeD, h: badgeD,
            fill: accent, text: String(index + 1),
            textOpts: { size: 11, bold: true, color: onColorText(accent), align: 'ctr', anchor: 'ctr' },
        });
        info.shapes += 1;
        bump(info.kinds, 'ellipse');
        const labelFit = fitTitle(phase.label, cmToEmu(Math.max(1, cellW - 1.6)), cmToEmu(0.9), [19, 17, 16, 15, 14, 13]);
        addText(sink, ctx, {
            x: cmToEmu(x + 1.16), y: cmToEmu(y + 0.58), w: cmToEmu(Math.max(1, cellW - 1.44)), h: cmToEmu(0.9),
            anchor: 'ctr', name: `${name} Label`, autofit: false,
            paras: [{ text: phase.label, sizePt: labelFit.sizePt, bold: true, color: pal.text, lineSpacing: 1.1 }],
        });
        const progressH = phase.progress === undefined ? 0 : 0.95;
        const bodyY = y + 1.56;
        const bodyH = Math.max(0.4, y + cellH - 0.32 - progressH - bodyY);
        const bullets = phase.items;
        if (bullets.length > 0) {
            const fit = fitParas(
                (size) => bullets.map((text) => ({
                    text, sizePt: size, color: pal.muted, lineSpacing: 1.28,
                    marL: cmToEmu(0.38), indent: -cmToEmu(0.38),
                    bullet: '•', bulletColor: accent, spaceBeforePt: 3,
                })),
                cmToEmu(Math.max(1, cellW - 0.72)),
                cmToEmu(bodyH),
                [13, 12, 11, 10.5, 10, 9.5, 9],
            );
            if (!fit.fits) overflowed += 1;
            addText(sink, ctx, {
                x: cmToEmu(x + 0.34), y: cmToEmu(bodyY), w: cmToEmu(Math.max(1, cellW - 0.72)), h: cmToEmu(bodyH),
                anchor: 't', name: `${name} Items`, autofit: false, paras: fit.paras,
            });
        }
        if (phase.progress !== undefined) {
            const trackW = Math.max(1, cellW - 1.66);
            const trackY = y + cellH - 0.78;
            const trackX = x + 0.34;
            addRect(sink, {
                name: `${name} Progress Track`,
                x: cmToEmu(trackX), y: cmToEmu(trackY), w: cmToEmu(trackW), h: cmToEmu(0.24),
                fill: pal.band,
            });
            if (phase.progress > 0) {
                addRect(sink, {
                    name: `${name} Progress Fill`,
                    x: cmToEmu(trackX), y: cmToEmu(trackY),
                    w: cmToEmu(trackW * phase.progress), h: cmToEmu(0.24), fill: accent,
                });
            }
            addText(sink, ctx, {
                x: cmToEmu(trackX + trackW + 0.16), y: cmToEmu(trackY - 0.13), w: cmToEmu(0.95), h: cmToEmu(0.5),
                anchor: 'ctr', name: `${name} Progress Text`, autofit: false,
                paras: [{
                    text: `${Math.round(phase.progress * 100)}%`, sizePt: 10, bold: true,
                    color: accent, align: 'r', lineSpacing: 1,
                }],
            });
        }
        if (rows === 1 && index < total - 1) {
            addLineShape(sink, ctx, {
                from: [x + cellW + 0.06, y + cellH / 2],
                to: [x + cellW + gap - 0.06, y + cellH / 2],
                color: pal.subtle, widthPt: 1.25, arrow: 'end',
                name: `Roadmap Arrow ${index + 1}`,
            });
            info.shapes += 1;
            bump(info.kinds, 'line');
        }
    });
    if (overflowed > 0) {
        ctx.notes.push(`第 ${ctx.page} 页 roadmap 有 ${overflowed} 个阶段的要点放不下，已压到最小 9pt，`
            + '建议精简每阶段要点');
    }
    return info;
}

/**
 * 2×2 矩阵（deck.matrix）：四象限 + 两条中线 + 轴标签 + 归一化坐标点。
 * x/y 是 0~1（原点左下），越界与缺坐标都只钳制 + 汇总一条告警，不抛。
 */
function renderMatrix(sink, ctx, data) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    addContentChrome(sink, ctx, data.title);
    const axisOf = (raw, fallback) => {
        const value = raw !== null && typeof raw === 'object' ? raw : {};
        return { low: asString(value.low, fallback[0]), high: asString(value.high, fallback[1]) };
    };
    const xAxis = axisOf(data.xAxis, ['低', '高']);
    const yAxis = axisOf(data.yAxis, ['低', '高']);
    const items = helperItems(data.items).map((item) => ({
        x: item.x === undefined || item.x === null ? undefined : asNumber(item.x, 0),
        y: item.y === undefined || item.y === null ? undefined : asNumber(item.y, 0),
        text: asString(item.text ?? item.title),
        desc: asString(item.desc ?? item.body),
    }));
    const gutter = 2.1;
    const plot = {
        x: box.x + gutter, y: box.y + 0.15,
        w: box.w - gutter - 0.25, h: box.h - 0.8 - 0.15,
    };
    const midX = plot.x + plot.w / 2;
    const midY = plot.y + plot.h / 2;
    const info = {
        title: asString(data.title), count: items.length, points: items.length,
        xAxis, yAxis, quadrants: 4, kinds: {}, shapes: 0,
    };
    // 四个象限：交替两档极浅底，点落在哪一格一眼能看出来（不靠颜色也靠中线分界）
    const fills = [
        { color: pal.band, alpha: 55 },
        undefined,
        undefined,
        { color: pal.band, alpha: 35 },
    ];
    const quads = [
        { x: plot.x, y: plot.y },
        { x: midX, y: plot.y },
        { x: plot.x, y: midY },
        { x: midX, y: midY },
    ];
    quads.forEach((quad, index) => {
        addRect(sink, {
            name: `Matrix Quadrant ${index + 1}`,
            x: cmToEmu(quad.x), y: cmToEmu(quad.y),
            w: cmToEmu(plot.w / 2), h: cmToEmu(plot.h / 2),
            fill: fills[index],
        });
    });
    addLineShape(sink, ctx, {
        from: [midX, plot.y], to: [midX, plot.y + plot.h],
        color: pal.border, widthPt: 1, name: 'Matrix Axis Vertical',
    });
    addLineShape(sink, ctx, {
        from: [plot.x, midY], to: [plot.x + plot.w, midY],
        color: pal.border, widthPt: 1, name: 'Matrix Axis Horizontal',
    });
    info.shapes += 2;
    bump(info.kinds, 'line');
    addAutoShape(sink, ctx, {
        preset: 'rect', name: 'Matrix Frame',
        x: plot.x, y: plot.y, w: plot.w, h: plot.h,
        fill: 'none', line: { color: pal.border, widthPt: 1 },
    });
    // 轴标签：x 轴在画面下缘（low 左 / high 右），y 轴在左侧留白（low 底 / high 顶）
    addText(sink, ctx, {
        x: cmToEmu(plot.x), y: cmToEmu(plot.y + plot.h + 0.06), w: cmToEmu(plot.w / 2 - 0.2), h: cmToEmu(0.55),
        anchor: 'ctr', name: 'Matrix XAxis Low', autofit: false,
        paras: [{ text: xAxis.low, sizePt: 10.5, color: pal.muted, align: 'l', lineSpacing: 1 }],
    });
    addText(sink, ctx, {
        x: cmToEmu(midX + 0.2), y: cmToEmu(plot.y + plot.h + 0.06), w: cmToEmu(plot.w / 2 - 0.3), h: cmToEmu(0.55),
        anchor: 'ctr', name: 'Matrix XAxis High', autofit: false,
        paras: [{ text: xAxis.high, sizePt: 10.5, color: pal.muted, align: 'r', lineSpacing: 1 }],
    });
    addText(sink, ctx, {
        x: cmToEmu(box.x), y: cmToEmu(plot.y - 0.08), w: cmToEmu(gutter - 0.25), h: cmToEmu(0.55),
        anchor: 'ctr', name: 'Matrix YAxis High', autofit: false,
        paras: [{ text: yAxis.high, sizePt: 10.5, color: pal.muted, align: 'r', lineSpacing: 1 }],
    });
    addText(sink, ctx, {
        x: cmToEmu(box.x), y: cmToEmu(plot.y + plot.h - 0.47), w: cmToEmu(gutter - 0.25), h: cmToEmu(0.55),
        anchor: 'ctr', name: 'Matrix YAxis Low', autofit: false,
        paras: [{ text: yAxis.low, sizePt: 10.5, color: pal.muted, align: 'r', lineSpacing: 1 }],
    });
    if (items.length === 0) {
        ctx.notes.push(`第 ${ctx.page} 页 matrix 没有任何坐标点，该页只有坐标轴`);
        return info;
    }
    const badCoord = [];
    const outOfRange = [];
    if (items.length > 12) {
        ctx.notes.push(`第 ${ctx.page} 页 matrix 有 ${items.length} 个坐标点，标签容易互相压住，建议减到 12 个以内`);
    }
    const dotD = 0.34;
    const clamps = { min: 0, max: 1 };
    const clamp01 = (value) => Math.min(clamps.max, Math.max(clamps.min, value));
    items.forEach((item, index) => {
        let { x, y } = item;
        let clamped = false;
        if (x === undefined || y === undefined) {
            badCoord.push(index + 1);
            x = 0;
            y = 0;
            clamped = true;
        } else {
            if (!(x >= 0 && x <= 1) || !(y >= 0 && y <= 1)) clamped = true;
            x = clamp01(x);
            y = clamp01(y);
        }
        if (clamped && !badCoord.includes(index + 1)) outOfRange.push(index + 1);
        const px = plot.x + x * plot.w;
        const py = plot.y + (1 - y) * plot.h;
        const accent = cycleAccent(pal, index);
        const name = `Matrix Point ${index + 1}`;
        addAutoShape(sink, ctx, {
            preset: 'ellipse', name,
            x: px - dotD / 2, y: py - dotD / 2, w: dotD, h: dotD,
            fill: accent, line: { color: 'FFFFFF', widthPt: 1.25 },
        });
        info.shapes += 1;
        bump(info.kinds, 'ellipse');
        const text = item.text;
        const desc = item.desc;
        if (text === '' && desc === '') return;
        const labelW = Math.min(6.4, plot.w * 0.32);
        const labelH = desc === '' ? 0.55 : 1.02;
        // 靠右的点把标签放到点的左侧、靠上的点往下放：标签不越出绘图区
        let labelX = px + 0.3;
        if (labelX + labelW > plot.x + plot.w) labelX = px - 0.3 - labelW;
        let labelY = py - labelH / 2;
        labelY = Math.min(plot.y + plot.h - labelH, Math.max(plot.y, labelY));
        const paras = [{ text, sizePt: desc === '' ? 11.5 : 12, bold: true, color: pal.text, lineSpacing: 1.15 }];
        if (desc !== '') {
            paras.push({ text: desc, sizePt: 9.5, color: pal.muted, lineSpacing: 1.2, spaceBeforePt: 2 });
        }
        addText(sink, ctx, {
            x: cmToEmu(labelX), y: cmToEmu(labelY), w: cmToEmu(labelW), h: cmToEmu(labelH),
            anchor: 'ctr', name: `${name} Text`, autofit: false, paras,
        });
    });
    if (badCoord.length > 0) {
        ctx.notes.push(`第 ${ctx.page} 页 matrix 有 ${badCoord.length} 个点缺少 x/y 坐标`
            + `（第 ${badCoord.join('、')} 个），已按左下角 (0, 0) 处理`);
    }
    if (outOfRange.length > 0) {
        ctx.notes.push(`第 ${ctx.page} 页 matrix 有 ${outOfRange.length} 个点的坐标超出 0~1`
            + `（第 ${outOfRange.join('、')} 个），已钳到边界`);
    }
    return info;
}

/** 助手 kind → 渲染函数。compile() 的派发与 outline/stats 用的是同一个名单。 */
const HELPER_RENDERERS = {
    cards: renderCards,
    steps: renderSteps,
    compare: renderCompare,
    kpi: renderKpi,
    imageText: renderImageText,
    timeline: renderTimeline,
    iconGrid: renderIconGrid,
    agenda: renderAgenda,
    gallery: renderGallery,
    roadmap: renderRoadmap,
    matrix: renderMatrix,
};

// ───────────────────────────────────────────────────────────────────────────
// 元素级助手：提示词面板 / 横幅
// ───────────────────────────────────────────────────────────────────────────

/**
 * panel：浅底圆角面板 + 小字号长文本（参考稿的「提示词示例」就是这种面板）。
 * 与 deck.shape 同级 —— 画在当前页，不新起页；h 缺省时按真实文字量算出来，
 * 面板高度才不会固定成一个既不贴合内容又容易压到页脚的值。
 */
function addPanel(sink, ctx, spec) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    const pad = 0.28;
    const title = asString(spec.title);
    const runs = asArray(spec.runs);
    const text = asString(spec.text) || runText(runs);
    const alignRaw = asString(spec.align, 'l');
    const align = ['l', 'ctr', 'r'].includes(alignRaw) ? alignRaw : 'l';
    if (align !== alignRaw && alignRaw !== '') {
        ctx.notes.push(`第 ${ctx.page} 页 panel 的 align「${alignRaw}」不认识，已按 l 处理（可用：l / ctr / r）`);
    }
    const sizePt = Math.min(20, Math.max(6, asNumber(spec.sizePt, 11)));
    const gaveW = spec.w !== undefined && spec.w !== null;
    const gaveH = spec.h !== undefined && spec.h !== null;
    if ((gaveW && !(asNumber(spec.w, 0) > 0)) || (gaveH && !(asNumber(spec.h, 0) > 0))) {
        ctx.notes.push(`第 ${ctx.page} 页 panel 的 w / h 必须是大于 0 的数字，已按最小值处理`);
    }
    const w = Math.max(1, gaveW ? asNumber(spec.w, box.w) : box.w);
    const innerW = Math.max(0.5, w - pad * 2);
    const titleH = title === '' ? 0 : 0.72;
    const titleParas = title === ''
        ? []
        : [{ text: title, sizePt: Math.min(14, sizePt + 1.5), bold: true, color: pal.text, lineSpacing: 1.2, align }];
    const bodyParas = [{
        text, runs, sizePt, color: pal.text, lineSpacing: 1.32, align,
    }];
    const measured = measureParas(bodyParas, cmToEmu(innerW));
    const neededCm = emuToCm(measured.neededPt * EMU_PER_PT);
    const defaultH = neededCm + pad * 2 + titleH;
    const h = Math.max(0.5, gaveH ? asNumber(spec.h, defaultH) : defaultH);
    const x = asNumber(spec.x, box.x);
    const y = asNumber(spec.y, box.y);
    const radius = Math.max(0, asNumber(spec.radius, 0.18));
    warnOffPage(ctx, 'panel', x, y, w, h);
    // 给了 h 就按 h 放；文字按字号阶梯往下压，压到最小仍放不下才进 warnings
    const availH = Math.max(0.2, h - pad * 2 - titleH);
    const fit = fitParas(
        (size) => [{ text, runs, sizePt: size, color: pal.text, lineSpacing: 1.32, align }],
        cmToEmu(innerW),
        cmToEmu(availH),
        [sizePt, sizePt - 0.5, sizePt - 1, sizePt - 1.5, sizePt - 2, sizePt - 2.5]
            .map((value) => Math.max(6, value)),
    );
    // 空面板（既没有标题也没有正文）没东西可放，不该报「放不下」
    if ((text !== '' || title !== '') && !fit.fits) {
        ctx.notes.push(`第 ${ctx.page} 页 panel 的文字放不下：估算需要 ${Math.ceil(fit.neededPt)}pt 高，`
            + `可用只有 ${Math.floor(cmToEmu(availH) / EMU_PER_PT)}pt，已压到最小 ${fit.sizePt}pt`);
    }
    addAutoShape(sink, ctx, {
        preset: 'roundRect',
        name: asString(spec.name, 'Panel'),
        x, y, w, h,
        fill: helperCardFill(spec.fill, ctx),
        line: helperCardLine(spec.line, ctx),
        radius,
        shadow: spec.shadow === true,
    });
    let cursor = y + pad;
    if (title !== '') {
        addText(sink, ctx, {
            x: cmToEmu(x + pad), y: cmToEmu(cursor), w: cmToEmu(innerW), h: cmToEmu(titleH),
            anchor: 'ctr', name: 'Panel Title', autofit: false, paras: titleParas,
        });
        cursor += titleH;
    }
    addText(sink, ctx, {
        x: cmToEmu(x + pad), y: cmToEmu(cursor), w: cmToEmu(innerW), h: cmToEmu(availH),
        anchor: 't', name: 'Panel Text', autofit: false, paras: fit.paras,
    });
    return { x, y, w, h, sizePt: fit.sizePt };
}

/**
 * banner：横幅 / 色带。
 * fill 支持纯色，gradient 支持渐变（复用 shape 的 a:gradFill：stops + angle）；
 * 文字色默认按底色亮度自动选黑/白，浅底上写白字这种事不该由用户来发现。
 */
function addBanner(sink, ctx, spec) {
    const { pal } = ctx;
    const box = contentBox(ctx);
    const x = asNumber(spec.x, box.x);
    const y = asNumber(spec.y, box.y);
    const w = Math.max(0.5, spec.w === undefined || spec.w === null ? box.w : asNumber(spec.w, box.w));
    const h = Math.max(0.4, spec.h === undefined || spec.h === null ? 1.4 : asNumber(spec.h, 1.4));
    let fill;
    if (spec.gradient !== undefined && spec.gradient !== null && typeof spec.gradient === 'object') {
        const stops = asArray(spec.gradient.stops);
        if (stops.length < 2) {
            ctx.notes.push(`第 ${ctx.page} 页 banner 渐变的 stops 只有 ${stops.length} 个，少于 2 个，已回落纯色`);
            fill = stops.length === 0 ? { color: pal.primary } : { color: stops[0]?.color };
        } else {
            fill = {
                type: 'gradient',
                stops: stops.map((stop) => ({ pos: clamp01(stop?.pos), color: stop?.color, alpha: stop?.alpha })),
                angle: asNumber(spec.gradient.angle, 0),
            };
        }
    } else if (spec.fill !== undefined && spec.fill !== null && spec.fill !== false) {
        fill = shapeFillOf(spec.fill, ctx);
    } else {
        fill = { color: pal.primary };
    }
    const baseColor = fill === undefined
        ? 'FFFFFF'
        : (fill.type === 'gradient' ? fill.stops?.[0]?.color : fill.color);
    const color = spec.color === undefined || spec.color === null || spec.color === ''
        ? onColorText(baseColor)
        : toHex(spec.color);
    const text = asString(spec.text);
    const sub = asString(spec.sub);
    const sizePt = Math.min(40, Math.max(8, asNumber(spec.sizePt, 18)));
    warnOffPage(ctx, 'banner', x, y, w, h);
    addAutoShape(sink, ctx, {
        preset: 'rect', name: asString(spec.name, 'Banner'),
        x, y, w, h, fill, line: 'none',
    });
    const paras = [];
    if (text !== '') {
        const fit = fitParas(
            (size) => [{ text, sizePt: size, bold: true, color, align: 'ctr', lineSpacing: 1.1 }],
            cmToEmu(Math.max(1, w - 0.6)),
            cmToEmu(sub === '' ? h : h * 0.62),
            [sizePt, sizePt - 1, sizePt - 2, sizePt - 3, sizePt - 4].map((value) => Math.max(8, value)),
        );
        if (!fit.fits) {
            ctx.notes.push(`第 ${ctx.page} 页 banner 的文字放不下，已压到最小 ${fit.sizePt}pt`);
        }
        paras.push(fit.paras[0]);
    }
    if (sub !== '') {
        paras.push({ text: sub, sizePt: Math.max(9, Math.round(sizePt * 0.66)), color, align: 'ctr', lineSpacing: 1.2, spaceBeforePt: 3 });
    }
    addText(sink, ctx, {
        x: cmToEmu(x + 0.3), y: cmToEmu(y), w: cmToEmu(Math.max(1, w - 0.6)), h: cmToEmu(h),
        anchor: 'ctr', name: 'Banner Text', autofit: false,
        paras: paras.length === 0 ? [{ text: '', sizePt, color, align: 'ctr' }] : paras,
    });
    return { x, y, w, h };
}

// ───────────────────────────────────────────────────────────────────────────
// 自由绘制页：形状 / 直线 / 图标 / 面板 / 横幅
// ───────────────────────────────────────────────────────────────────────────

/**
 * 画布配图（第五十二轮 `新-6`）：把一张图放进「宿主页」的一栏里，与版式共存。
 *
 * 矩形来源三选一（优先级从高到低）：
 * 1. 调用方给的 `x/y/w/h`（厘米）—— 完全按给的画，越界照旧进 warnings；
 * 2. `plan` —— 分栏算出来的那一栏（见 planPageGeometry）；
 * 3. 版心整块 —— 自由绘制页上没给摆位时的兜底。
 *
 * 与「单图版式」（renderImage）的分工：这里不画标题带、不接管整页，只画图 + 题注，
 * 所以它能叠在 cards / bullets 这些版式页上。图片照旧进 stats.images（与 read() 数 p:pic 口径一致）。
 *
 * 与单图页同一套尺寸规则：`widthCm` / `heightCm` 照收（只给一个时按原图纵横比推另一个，
 * 超出可用区域整体缩小并告警），`fit` 的裁切与 DPI 也都报回生成端 —— 否则
 * `save()` 与 `read()` 对同一份文件会给出两个数（复核实测过）。
 *
 * @returns {{x: number, y: number, w: number, h: number, caption: string, missing: boolean,
 *   unsupported: boolean, cropped: boolean, dpi: number, widthCm: number, heightCm: number,
 *   path: string}}
 */
function addCanvasImage(sink, ctx, spec, plan) {
    const { pal } = ctx;
    // 默认框用**整幅版心**：分栏页上的版心已经收窄成左栏，拿它当默认会让
    // 「只给 x」的配图变成一条左栏宽的带子（复核实测会盖住卡片）
    const box = contentBox(ctx);
    const full = {
        x: box.x,
        y: box.y,
        w: emuToCm(ctx.g.fullContentW ?? ctx.g.contentW),
        h: box.h,
    };
    const given = (key) => spec[key] !== undefined && spec[key] !== null;
    const explicit = ['x', 'y', 'w', 'h'].some(given);
    let x;
    let y;
    let w;
    let h;
    if (explicit) {
        x = asNumber(spec.x, full.x);
        y = asNumber(spec.y, full.y);
        w = Math.max(0.5, asNumber(spec.w, full.w));
        h = Math.max(0.5, asNumber(spec.h, full.h));
    } else if (plan !== undefined) {
        ({ x, y, w, h } = plan);
    } else {
        ({ x, y, w, h } = full);
    }
    warnOffPage(ctx, 'image', x, y, w, h);
    // 与版式同页的配图没有标题带（标题带是整页版式的构件）：给了 title 就如实说清去哪了，
    // 静默丢掉一个调用方明确写下的字段是最难排查的一类「没生效」
    if (asString(spec.title) !== '') {
        ctx.notes.push(`第 ${ctx.page} 页的配图与版式同页，image 的 title「${asString(spec.title)}」`
            + '只在一张图独占一页时生效，已忽略；要与版式同页又要标题，请改用 deck.imageText 或页内形状');
    }
    const caption = asString(spec.caption);
    // 题注高度按栏高取一档，最高 1.1cm：窄栏里用固定 1cm 会把图挤得太小
    const captionH = caption === '' ? 0 : Math.min(1.1, Math.max(0.6, h * 0.16));
    const areaH = Math.max(0.6, h - captionH - (captionH === 0 ? 0 : 0.18));
    const media = ctx.mediaOf(asString(spec.path));
    const frame = { x: cmToEmu(x), y: cmToEmu(y), contentW: cmToEmu(w), h: cmToEmu(areaH) };
    const boxSpec = {
        fit: imageFit(spec.fit, ctx.notes),
        align: imageAlign(spec.align, ctx.notes),
        // 显式尺寸照收（只给一个时 layoutImage 按纵横比推另一个）
        widthCm: asNumber(spec.widthCm, 0),
        heightCm: asNumber(spec.heightCm, 0),
        fullBleed: false,
    };
    const drawn = placeholderOf(media)
        ? placeholderBox(frame.x, frame.y, frame.contentW, frame.h, media)
        : mediaBox(media, boxSpec, frame, ctx);
    // 名字带序号：同页两张配图重名会让 read() 侧按名取几何的探针只看得到第一张
    addPicture(sink, ctx, {
        ...drawn,
        frame: spec.frame === true,
        name: `Canvas Picture ${sink.imageCount + 1}`,
    });
    if (caption !== '') {
        addText(sink, ctx, {
            x: frame.x, y: cmToEmu(y + h - captionH), w: frame.contentW, h: cmToEmu(captionH),
            anchor: 't', name: 'Caption', autofit: false,
            paras: [{ text: caption, sizePt: 10, color: pal.muted, align: 'ctr', lineSpacing: 1.2 }],
        });
    }
    // 裁切与 DPI 与单图页同口径（finishImagePage）：生成端不报，read() 一读就会报，
    // 同一份文件两个数就是报告在说谎
    const cropped = drawn.crop !== undefined && (drawn.crop.l > 0 || drawn.crop.t > 0);
    const dpi = placeholderOf(media) || drawn.w <= 0
        ? 0
        : Math.round(media.width / ((drawn.w / 360000) / 2.54));
    return {
        x, y, w, h, caption,
        path: asString(spec.path),
        missing: media?.missing === true,
        unsupported: media?.unsupported === true,
        cropped,
        dpi,
        widthCm: cm1(drawn.w / 360000),
        heightCm: cm1(drawn.h / 360000),
    };
}

/**
 * 自由绘制页的画布条目。
 * 一页里形状的先后顺序就是图层顺序（先写的在下），因此不排序、不合并，
 * 完全按调用顺序落盘 —— 用户能预期「后画的盖在前面」。
 * 助手页复用这同一个函数画它们的叠加层（cards 之后接着 banner 就是走这里），
 * 基础版式页（bullets / table / …）现在也走这里（第五十二轮 `新-6`）。
 */
function renderCanvasItems(sink, ctx, items) {
    const kinds = {};
    let icons = 0;
    let shapes = 0;
    let images = 0;
    let missing = 0;
    let placeheld = 0;
    let cropped = 0;
    let dpi = 0;
    let widthCm = 0;
    let heightCm = 0;
    const missingPaths = [];
    const plan = asArray(ctx.canvasImageRects);
    for (const [index, entry] of items.entries()) {
        const spec = entry?.spec ?? {};
        if (entry.type === 'icon') {
            if (addIconShape(sink, ctx, spec) !== undefined) icons += 1;
            bump(kinds, 'icon');
        } else if (entry.type === 'line') {
            addLineShape(sink, ctx, spec);
            shapes += 1;
            bump(kinds, 'line');
        } else if (entry.type === 'panel') {
            addPanel(sink, ctx, spec);
            shapes += 1;
            bump(kinds, 'panel');
        } else if (entry.type === 'banner') {
            addBanner(sink, ctx, spec);
            shapes += 1;
            bump(kinds, 'banner');
        } else if (entry.type === 'image') {
            const drawn = addCanvasImage(sink, ctx, spec, plan[index]);
            images += 1;
            shapes += 1;
            if (drawn.missing) {
                missing += 1;
                missingPaths.push(drawn.path);
            }
            if (drawn.unsupported) placeheld += 1;
            if (drawn.cropped) cropped += 1;
            // DPI 取最悲观的一张（要报就报最该被看见的那张）
            if (drawn.dpi > 0 && (dpi === 0 || drawn.dpi < dpi)) {
                dpi = drawn.dpi;
                widthCm = drawn.widthCm;
                heightCm = drawn.heightCm;
            }
        } else {
            const drawn = addAutoShape(sink, ctx, spec);
            shapes += 1;
            bump(kinds, drawn.preset);
        }
    }
    // 字段名与单图页（finishImagePage）对齐：collectWarnings 那一段才能两条路共用
    return {
        items: items.length, shapes, icons, images, kinds,
        missing, missingPaths, placeheld, cropped, dpi, widthCm, heightCm,
    };
}

/** 自由绘制页（deck.shape / line / icon / panel / banner）。 */
function renderShapePage(sink, ctx, data) {
    // 叠加层单独走 data.canvas：bullets / table 这些版式的 data.items 是自己的内容，
    // 两者同名会让「要点页的 items 被当成自由形状画一遍」
    const items = asArray(data.canvas);
    const drawn = renderCanvasItems(sink, ctx, items);
    return {
        title: asString(data.title ?? items[0]?.spec?.text),
        ...drawn,
    };
}

/**
 * 助手页 / 版式页的叠加层：把 items 通道画出来的形状并进这一页自己的计数。
 *
 * 图片那几项（missing / placeheld / cropped / dpi / widthCm / heightCm）按「与宿主页自己的
 * 图合并」处理，而不是覆盖：`imageText` 这一页本来就带一套（finishImagePage 给的），
 * 直接覆盖会把宿主自己的图从告警里抹掉。
 */
function mergeCanvasInfo(info, extra) {
    const kinds = { ...(info.kinds ?? {}) };
    for (const [key, count] of Object.entries(extra.kinds)) bump(kinds, key, count);
    const merged = {
        ...info,
        extra,
        shapes: (info.shapes ?? 0) + extra.shapes,
        icons: (info.icons ?? 0) + extra.icons,
        // 画布配图（新-6）：叠在版式页上的图也要能被 outline / warnings 报出来
        images: (info.images ?? 0) + (extra.images ?? 0),
    };
    if ((extra.images ?? 0) === 0) return { ...merged, kinds };
    merged.missing = (info.missing ?? 0) + (extra.missing ?? 0);
    merged.missingPaths = [...asArray(info.missingPaths), ...asArray(extra.missingPaths)];
    merged.placeheld = (info.placeheld ?? 0) + (extra.placeheld ?? 0);
    merged.cropped = (info.cropped ?? 0) + (extra.cropped ?? 0);
    const own = asNumber(info.dpi, 0);
    const added = asNumber(extra.dpi, 0);
    const candidates = [own, added].filter((value) => value > 0);
    merged.dpi = candidates.length === 0 ? 0 : Math.min(...candidates);
    // 尺寸跟着报出的那一张走（自己更悲观时保留自己的）
    if (added > 0 && (own === 0 || added <= own)) {
        merged.widthCm = extra.widthCm;
        merged.heightCm = extra.heightCm;
    }
    return { ...merged, kinds };
}

// ───────────────────────────────────────────────────────────────────────────
// 页面母版：背景 / logo / 页眉 / 页脚 / 页码 / 装饰条
// ───────────────────────────────────────────────────────────────────────────

/**
 * 背景归一。
 * 三种写法：`{image}`、`{color}`、`{gradient}`，或字符串（纯色）/ 'none'。
 * 空串与 FFFFFF 一律当成「不加背景」：素色主题的底色就是白，把白当成真实填充
 * 会让每一页都多一层本不存在的底（与 Word 引擎里的 shadeOf 是同一个理由）。
 */
function normalizeBackground(raw, ctx, where) {
    if (raw === undefined || raw === null || raw === false) return undefined;
    if (typeof raw === 'string') {
        if (raw.trim().toLowerCase() === 'none' || isBlankFill(raw)) return undefined;
        return { kind: 'fill', fill: raw };
    }
    if (typeof raw !== 'object') return undefined;
    if (raw.image !== undefined && raw.image !== null) {
        const path = asString(raw.image.path);
        if (path === '') {
            ctx.notes.push(`${where}背景图缺少 path，已忽略背景`);
            return undefined;
        }
        return {
            kind: 'image',
            path,
            fit: raw.image.fit === 'contain' ? 'contain' : imageFit(raw.image.fit, ctx.notes),
            dim: clamp01(raw.image.dim),
            dimColor: raw.image.dimColor ?? '000000',
        };
    }
    if (raw.gradient !== undefined && raw.gradient !== null) {
        const stops = asArray(raw.gradient.stops);
        if (stops.length < 2) {
            ctx.notes.push(`${where}背景渐变的 stops 只有 ${stops.length} 个，少于 2 个，已回落纯色`);
            if (stops.length === 0) return undefined;
            return { kind: 'fill', fill: stops[0]?.color };
        }
        return {
            kind: 'fill',
            fill: {
                type: 'gradient',
                stops: stops.map((stop) => ({ pos: clamp01(stop?.pos), color: stop?.color, alpha: stop?.alpha })),
                angle: asNumber(raw.gradient.angle, 0),
            },
        };
    }
    if (raw.color !== undefined) return isBlankFill(raw.color) ? undefined : { kind: 'fill', fill: raw.color };
    return undefined;
}

/**
 * 铺设背景层，返回要写进 `p:bg` 的填充（纯色/渐变走 p:bg，图片走一张铺满的 p:pic）。
 * 为什么图片不用 blipFill 写 p:bg：本引擎的满页背景图一直是用 p:pic 表达的，
 * 两条路径都留着会让「背景图被裁切」「满页背景」这些报告字段多出一套口径。
 */
function applyBackground(sink, ctx, bg) {
    if (bg === undefined) return undefined;
    if (bg.kind === 'fill') return bg.fill;
    const { g } = ctx;
    const media = ctx.mediaOf(bg.path);
    const spec = { fit: bg.fit, align: 'center', widthCm: 0, heightCm: 0, fullBleed: false };
    const frame = { x: 0, contentW: g.cx, y: 0, h: g.cy };
    const box = placeholderOf(media)
        ? placeholderBox(0, 0, g.cx, g.cy, media)
        : mediaBox(media, spec, frame, ctx);
    // counts=false：背景不算「一张配图」，也不占版心面积；名字保留 Picture Full Bleed，
    // 这样 read() 能把「满页背景」这一事实读回来
    addPicture(sink, ctx, { ...box, frame: false, fullBleed: true }, false);
    if (bg.dim > 0) {
        // 压暗 = 一层半透明纯色矩形。画在图片之上、内容之下，参考稿的「照片 + 红色遮罩」就是这个
        addRect(sink, {
            x: 0, y: 0, w: g.cx, h: g.cy,
            fill: { color: bg.dimColor, alpha: bg.dim * 100 },
            name: 'Background Dim',
        });
    }
    return undefined;
}

/** 母版归一：非法枚举只 warning + 回落，绝不抛。 */
function normalizeMaster(raw, ctx) {
    if (raw === undefined || raw === null || typeof raw !== 'object') return undefined;
    const master = {
        background: normalizeBackground(raw.background, ctx, '母版'),
        logo: undefined,
        header: undefined,
        footer: undefined,
        pageNumber: undefined,
        accents: [],
        skipLayouts: DEFAULT_SKIP_LAYOUTS.slice(),
    };
    if (raw.logo !== undefined && raw.logo !== null && typeof raw.logo === 'object') {
        const path = asString(raw.logo.path);
        if (path === '') {
            ctx.notes.push('母版 logo 缺少 path，已忽略 logo');
        } else {
            const corner = asString(raw.logo.corner, 'top-right').trim();
            if (!CORNERS.includes(corner)) {
                ctx.notes.push(`母版 logo 的 corner「${corner}」不认识，已按 top-right 处理（可用：${CORNERS.join(' / ')}）`);
            }
            master.logo = {
                path,
                corner: CORNERS.includes(corner) ? corner : 'top-right',
                widthCm: asNumber(raw.logo.widthCm, 1.6),
                marginCm: asNumber(raw.logo.marginCm, 0.5),
            };
        }
    }
    if (raw.header !== undefined && raw.header !== null) {
        const spec = typeof raw.header === 'string' ? { text: raw.header } : raw.header;
        const text = asString(spec.text);
        if (text !== '') {
            master.header = {
                text,
                size: asNumber(spec.size, 18),
                bold: spec.bold !== false,
                color: spec.color,
                align: asString(spec.align, 'l'),
                x: spec.x,
                y: spec.y,
                w: spec.w,
                font: spec.font,
            };
        }
    }
    if (raw.footer !== undefined && raw.footer !== null) {
        const spec = typeof raw.footer === 'string' ? { text: raw.footer } : raw.footer;
        const text = asString(spec.text);
        if (text !== '') {
            master.footer = {
                text,
                size: asNumber(spec.size, 10),
                color: spec.color,
                align: asString(spec.align, 'l'),
                font: spec.font,
            };
        }
    }
    if (raw.pageNumber !== undefined && raw.pageNumber !== null && typeof raw.pageNumber === 'object') {
        const corner = asString(raw.pageNumber.corner, 'bottom-right').trim();
        if (!CORNERS.includes(corner)) {
            ctx.notes.push(`母版页码的 corner「${corner}」不认识，已按 bottom-right 处理（可用：${CORNERS.join(' / ')}）`);
        }
        master.pageNumber = {
            show: raw.pageNumber.show !== false,
            corner: CORNERS.includes(corner) ? corner : 'bottom-right',
            from: Math.max(1, Math.floor(asNumber(raw.pageNumber.from, 2))),
            format: asString(raw.pageNumber.format, '{n}'),
            size: asNumber(raw.pageNumber.size, 10),
            color: raw.pageNumber.color,
            font: raw.pageNumber.font,
        };
    }
    // accent 既支持单个对象也支持数组：写一条装饰条不该强迫用户套一层数组
    const accentList = Array.isArray(raw.accent)
        ? raw.accent
        : (raw.accent !== undefined && raw.accent !== null && typeof raw.accent === 'object' ? [raw.accent] : []);
    for (const item of accentList) {
        if (item === null || item === undefined || typeof item !== 'object') continue;
        const bar = asString(item.bar, 'left-top').trim();
        if (!ACCENT_BARS.includes(bar)) {
            ctx.notes.push(`母版装饰条的 bar「${bar}」不认识，已忽略该条（可用：${ACCENT_BARS.join(' / ')}）`);
            continue;
        }
        master.accents.push({
            bar,
            color: item.color,
            sizeCm: Math.max(0, asNumber(item.sizeCm, 0.4)),
            lengthCm: item.lengthCm === undefined || item.lengthCm === null ? undefined : asNumber(item.lengthCm, 0),
        });
    }
    if (raw.skipLayouts !== undefined && raw.skipLayouts !== null) {
        const list = asArray(raw.skipLayouts).map((item) => asString(item).trim()).filter((item) => item !== '');
        const known = [...LAYOUT_KEYS, 'shape', ...HELPER_KINDS];
        for (const item of list) {
            if (!known.includes(item)) ctx.notes.push(`母版 skipLayouts 里的「${item}」不是版式名，已忽略`);
        }
        master.skipLayouts = list.filter((item) => known.includes(item));
    }
    return master;
}

/** 母版引用的图片先过一遍媒体表：缺图/格式不支持要在 warnings 里说清是哪张。 */
function preloadMasterMedia(ctx, master, warn) {
    if (master === undefined) return;
    const entries = [];
    if (master.background !== undefined && master.background.kind === 'image') {
        entries.push(['母版背景图', master.background.path]);
    }
    if (master.logo !== undefined) entries.push(['母版 logo', master.logo.path]);
    for (const [what, path] of entries) {
        const media = ctx.mediaOf(path);
        if (media === undefined || media === null) continue;
        if (media.missing === true) warn(`${what} 缺失：${path}，该图不会被画出来`);
        else if (media.unsupported === true) warn(`${what} 格式不支持（只认 PNG/JPEG/GIF）：${path}，已画成占位框`);
    }
}

/** 角标盒子：四个角共用一套margin，角落语义只在这一处翻译成坐标。 */
function cornerBox(ctx, corner, w, h, margin) {
    const { g } = ctx;
    const left = corner === 'top-left' || corner === 'bottom-left';
    const top = corner === 'top-left' || corner === 'top-right';
    return {
        x: left ? margin : g.cx - margin - w,
        y: top ? margin : g.cy - margin - h,
        w,
        h,
    };
}

/** logo：按 widthCm 等比缩放（原图纵横比优先），落在指定角。 */
function addMasterLogo(sink, ctx, logo) {
    const media = ctx.mediaOf(logo.path);
    const w = Math.max(9525, cmToEmu(logo.widthCm));
    const ratio = media !== undefined && media !== null && media.width > 0 && media.height > 0
        ? media.height / media.width
        : 1;
    const h = Math.max(9525, Math.round(w * ratio));
    const box = cornerBox(ctx, logo.corner, w, h, cmToEmu(logo.marginCm));
    // logo 计入 stats.images：read() 会把每个 p:pic 都数进去，两边口径保持一致才对得上
    addPicture(sink, ctx, { ...box, info: media, frame: false, name: 'Master Logo' });
}

function addMasterHeader(sink, ctx, header) {
    const { g } = ctx;
    // 母版是页面级装饰：分栏页（新-6）上仍按整幅版心摆
    const fullW = g.fullContentW ?? g.contentW;
    const x = header.x === undefined ? g.margin : cmToEmu(asNumber(header.x, 0.9));
    const y = header.y === undefined ? cmToEmu(0.32) : cmToEmu(asNumber(header.y, 0.32));
    const w = header.w === undefined ? fullW : cmToEmu(asNumber(header.w, 0));
    addText(sink, ctx, {
        x, y, w: w > 0 ? w : fullW, h: cmToEmu(0.95), anchor: 't', name: 'Master Header', autofit: false,
        paras: [{
            text: header.text, sizePt: header.size, bold: header.bold,
            color: header.color ?? ctx.pal.muted, align: header.align, lineSpacing: 1, font: header.font,
        }],
    });
}

function addMasterFooter(sink, ctx, footer) {
    const { g } = ctx;
    addText(sink, ctx, {
        x: g.margin, y: g.footerY + cmToEmu(0.1), w: Math.round((g.fullContentW ?? g.contentW) * 0.72),
        h: g.footerH,
        anchor: 't', name: 'Master Footer', autofit: false,
        paras: [{
            text: footer.text, sizePt: footer.size,
            color: footer.color ?? ctx.pal.footer, align: footer.align, lineSpacing: 1, font: footer.font,
        }],
    });
}

/** 页码：真实页序号（1 基）。format 支持 {n} 与 {total}。 */
function addMasterPageNumber(sink, ctx, spec) {
    const text = asString(spec.format, '{n}')
        .replace(/\{n\}/g, String(ctx.page))
        .replace(/\{total\}/g, String(ctx.total));
    const box = cornerBox(ctx, spec.corner, cmToEmu(4.0), cmToEmu(0.8), cmToEmu(0.6));
    addText(sink, ctx, {
        x: box.x, y: box.y, w: box.w, h: box.h, anchor: 'ctr', name: 'Master Page Number', autofit: false,
        paras: [{
            text, sizePt: spec.size, color: spec.color ?? ctx.pal.footer,
            align: spec.corner.endsWith('right') ? 'r' : 'l', lineSpacing: 1, font: spec.font,
        }],
    });
}

/** 装饰条：贴边的纯色矩形，四种走向。sizeCm 是厚度，lengthCm 是长度（缺省铺满）。 */
function addAccentBar(sink, ctx, bar, index) {
    const { g, pal } = ctx;
    const size = cmToEmu(bar.sizeCm);
    const length = bar.lengthCm === undefined ? undefined : cmToEmu(bar.lengthCm);
    const color = bar.color ?? pal.accent;
    const name = `Master Accent ${index + 1}`;
    if (bar.bar === 'top' || bar.bar === 'bottom') {
        addRect(sink, {
            x: 0, y: bar.bar === 'top' ? 0 : g.cy - size,
            w: length ?? g.cx, h: size, fill: color, name,
        });
        return;
    }
    addRect(sink, {
        x: 0, y: 0, w: size,
        h: length ?? (bar.bar === 'left-top' ? cmToEmu(4.0) : g.cy),
        fill: color, name,
    });
}

/**
 * 母版叠加层。
 * 顺序即图层：logo → 页眉 → 页脚 → 页码 → 装饰条，全部在内容之上；
 * 背景层由 compile() 在内容之前铺好。
 * skipLayouts（默认封面与结尾）不加页眉/页脚/页码：这两类页的版式自带大标题，
 * 再压一行页眉会打架；背景、logo、装饰条照加。
 */
function masterOverlays(sink, ctx, source, master) {
    if (master === undefined) return;
    const skip = master.skipLayouts.includes(source.kind);
    if (master.logo !== undefined) addMasterLogo(sink, ctx, master.logo);
    if (!skip) {
        if (master.header !== undefined) addMasterHeader(sink, ctx, master.header);
        if (master.footer !== undefined) addMasterFooter(sink, ctx, master.footer);
        const page = master.pageNumber;
        if (page !== undefined && page.show === true && ctx.page >= page.from) {
            addMasterPageNumber(sink, ctx, page);
        }
    }
    master.accents.forEach((bar, index) => addAccentBar(sink, ctx, bar, index));
}

// ───────────────────────────────────────────────────────────────────────────
// 包级 XML
// ───────────────────────────────────────────────────────────────────────────

function titleStyleXml(font, color) {
    return '<p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="2600" b="1" lang="zh-CN" altLang="en-US">'
        + `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`
        + `<a:latin typeface="${escapeAttr(font.en)}"/><a:ea typeface="${escapeAttr(font.cn)}"/>`
        + `<a:cs typeface="${escapeAttr(font.en)}"/></a:defRPr></a:lvl1pPr></p:titleStyle>`;
}

function bodyStyleXml(font, pal) {
    const sizes = [1800, 1500, 1300, 1200, 1100];
    const levels = sizes.map((size, index) => {
        const spec = LEVELS[Math.min(index, LEVELS.length - 1)];
        return `<a:lvl${index + 1}pPr marL="${spec.marL}" indent="${spec.indent}">`
            + `<a:defRPr sz="${size}" lang="zh-CN" altLang="en-US">`
            + `<a:solidFill><a:srgbClr val="${index === 0 ? pal.text : pal.muted}"/></a:solidFill>`
            + `<a:latin typeface="${escapeAttr(font.en)}"/><a:ea typeface="${escapeAttr(font.cn)}"/>`
            + `<a:cs typeface="${escapeAttr(font.en)}"/></a:defRPr></a:lvl${index + 1}pPr>`;
    }).join('');
    return `<p:bodyStyle>${levels}</p:bodyStyle>`;
}

function otherStyleXml(font, pal) {
    const defRPr = `<a:defRPr sz="1200" lang="zh-CN" altLang="en-US">`
        + `<a:solidFill><a:srgbClr val="${pal.text}"/></a:solidFill>`
        + `<a:latin typeface="${escapeAttr(font.en)}"/><a:ea typeface="${escapeAttr(font.cn)}"/>`
        + `<a:cs typeface="${escapeAttr(font.en)}"/></a:defRPr>`;
    return '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1200" lang="zh-CN" altLang="en-US">'
        + `<a:solidFill><a:srgbClr val="${pal.text}"/></a:solidFill>`
        + `<a:latin typeface="${escapeAttr(font.en)}"/><a:ea typeface="${escapeAttr(font.cn)}"/>`
        + `<a:cs typeface="${escapeAttr(font.en)}"/></a:defRPr></a:lvl1pPr>`
        + `<a:defPPr>${defRPr}</a:defPPr></p:otherStyle>`;
}

function slideMasterXml(ctx) {
    const { font, pal } = ctx;
    const layoutIds = LAYOUTS.map((item, index) => `<p:sldLayoutId id="${2147483649 + index}" r:id="rId${index + 1}"/>`).join('');
    return docXml('p:sldMaster', `
<p:cSld>${bgXml(pal.bg)}<p:spTree>${spTreeSkeleton()}</p:spTree></p:cSld>
<p:clrMap ${CLR_MAP_ATTRS}/>
<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst>
<p:txStyles>${titleStyleXml(font, pal.primary)}${bodyStyleXml(font, pal)}${otherStyleXml(font, pal)}</p:txStyles>`.trim());
}

function slideLayoutXml(ctx, layout) {
    return docXml('p:sldLayout', `
<p:cSld name="${escapeAttr(layout.name)}"><p:spTree>${spTreeSkeleton()}</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`.trim(), ` preserve="1" type="${layout.type}"`);
}

function themeXml(ctx) {
    const { font, pal, theme } = ctx;
    const c = theme.colors;
    const scheme = [
        ['dk1', c.text], ['lt1', c.bg], ['dk2', c.primaryDark], ['lt2', c.surface],
        ['accent1', c.primary], ['accent2', c.secondary], ['accent3', c.accent],
        ['accent4', c.subtle], ['accent5', c.border], ['accent6', c.muted],
        ['hlink', c.secondary], ['folHlink', c.subtle],
    ].map(([slot, value]) => `<a:${slot}><a:srgbClr val="${toHex(value)}"/></a:${slot}>`).join('');
    const fontEntry = (script, face) => `<a:font script="${script}" typeface="${escapeAttr(face)}"/>`;
    const fontCollection = (latin, ea, extra) => `<a:latin typeface="${escapeAttr(latin)}"/>`
        + `<a:ea typeface="${escapeAttr(ea)}"/><a:cs typeface=""/>${extra}`;
    const cjk = fontEntry('Hans', font.cn) + fontEntry('Hant', font.cn);

    const gradA = `<a:gradFill rotWithShape="1"><a:gsLst>`
        + '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
        + '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
        + '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>'
        + '</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>';
    const gradB = `<a:gradFill rotWithShape="1"><a:gsLst>`
        + '<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>'
        + '<a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>'
        + '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>'
        + '</a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>';
    const line = (w, extra) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">`
        + `<a:solidFill><a:schemeClr val="phClr">${extra}</a:schemeClr></a:solidFill>`
        + '<a:prstDash val="solid"/></a:ln>';
    const shadow = (blur, dist, alpha) => '<a:effectLst><a:outerShdw'
        + ` blurRad="${blur}" dist="${dist}" dir="5400000" rotWithShape="0">`
        + `<a:srgbClr val="000000"><a:alpha val="${alpha}"/></a:srgbClr></a:outerShdw></a:effectLst>`;

    // fmtScheme 的四个列表都必须是 3 项：PowerPoint 按 1-based 索引取用，少一项就判包损坏
    const fmtScheme = `<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>${gradA}${gradB}</a:fillStyleLst>
<a:lnStyleLst>${line(9525, '<a:shade val="95000"/><a:satMod val="105000"/>')}${line(25400, '')}${line(38100, '')}</a:lnStyleLst>
<a:effectStyleLst><a:effectStyle>${shadow(40000, 20000, 38000)}</a:effectStyle><a:effectStyle>${shadow(40000, 23000, 35000)}</a:effectStyle><a:effectStyle>${shadow(40000, 23000, 35000)}</a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>${gradB}</a:bgFillStyleLst>
</a:fmtScheme>`;

    return `${XML_HEAD}<a:theme xmlns:a="${NS_A}" name="DSH Office">`
        + `<a:themeElements><a:clrScheme name="DSH ${theme.name}">${scheme}</a:clrScheme>`
        + '<a:fontScheme name="DSH Office">'
        + `<a:majorFont>${fontCollection(font.titleEn, font.titleCn, cjk)}</a:majorFont>`
        + `<a:minorFont>${fontCollection(font.en, font.cn, cjk)}</a:minorFont>`
        + '</a:fontScheme>'
        + fmtScheme
        + '</a:themeElements>'
        + '<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
}

function defaultTextStyleXml(font, pal) {
    const sizes = [1800, 1600, 1400, 1200, 1100];
    const levels = sizes.map((size, index) => `<a:lvl${index + 1}pPr><a:defRPr sz="${size}" lang="zh-CN" altLang="en-US">`
        + `<a:solidFill><a:srgbClr val="${pal.text}"/></a:solidFill>`
        + `<a:latin typeface="${escapeAttr(font.en)}"/><a:ea typeface="${escapeAttr(font.cn)}"/>`
        + `<a:cs typeface="${escapeAttr(font.en)}"/></a:defRPr></a:lvl${index + 1}pPr>`).join('');
    return `<p:defaultTextStyle>${levels}</p:defaultTextStyle>`;
}

function presentationXml(ctx, slideCount, size, notesRelId) {
    const masterId = '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>';
    // notesMaster 必须在这里登记（并有一条 presentation→notesMaster 关系），否则 PowerPoint
    // 认得包里的 notesSlide 部件，却不把备注显示到备注页上：不报错，只是静默丢掉。
    // 位置也有讲究：实测放在 sldIdLst 之后才被接受，放在 sldMasterIdLst 之后会被判为坏文件。
    const notesMasterId = notesRelId === undefined
        ? ''
        : `<p:notesMasterIdLst><p:notesMasterId id="2147483649" r:id="${notesRelId}"/></p:notesMasterIdLst>\n`;
    const slideIds = Array.from({ length: slideCount }, (_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 3}"/>`).join('');
    return docXml('p:presentation', `
${masterId}
<p:sldIdLst>${slideIds}</p:sldIdLst>
${notesMasterId}<p:sldSz cx="${size.cx}" cy="${size.cy}" type="${size.type}"/>
<p:notesSz cx="6858000" cy="9144000"/>
${defaultTextStyleXml(ctx.font, ctx.pal)}`.trim());
}

function notesMasterXml(ctx) {
    const ph = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/>'
        + '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
        + '<p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm>'
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        + '<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>'
        + '<a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>';
    const notesStyle = `<p:notesStyle><a:lvl1pPr><a:defRPr sz="1200" lang="zh-CN" altLang="en-US">`
        + `<a:solidFill><a:srgbClr val="${ctx.pal.text}"/></a:solidFill>`
        + `<a:latin typeface="${escapeAttr(ctx.font.en)}"/><a:ea typeface="${escapeAttr(ctx.font.cn)}"/>`
        + `<a:cs typeface="${escapeAttr(ctx.font.en)}"/></a:defRPr></a:lvl1pPr></p:notesStyle>`;
    return docXml('p:notesMaster', `
<p:cSld>${bgXml(ctx.pal.bg)}<p:spTree>${spTreeSkeleton()}${ph}</p:spTree></p:cSld>
<p:clrMap ${CLR_MAP_ATTRS}/>
${notesStyle}`.trim());
}

function notesSlideXml(ctx, notes) {
    const font = ctx.font;
    const paras = String(notes).split('\n');
    const body = paras.map((line, index) => `<a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="115000"/></a:lnSpc>`
        + `<a:buNone/>${index === 0 ? '' : '<a:spcBef><a:spcPts val="400"/></a:spcBef>'}</a:pPr>`
        + (line === ''
            ? '<a:endParaRPr lang="zh-CN" sz="1200"/>'
            : `<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="1200" dirty="0">`
            + `<a:solidFill><a:srgbClr val="${ctx.pal.text}"/></a:solidFill>`
            + `<a:latin typeface="${escapeAttr(font.en)}"/><a:ea typeface="${escapeAttr(font.cn)}"/>`
            + `<a:cs typeface="${escapeAttr(font.en)}"/></a:rPr><a:t>${escapeAttr(line)}</a:t></a:r>`)
        + '</a:p>').join('');
    const shape = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/>'
        + '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
        + '<p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm>'
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        + `<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>`
        + `<a:lstStyle/>${body}</p:txBody></p:sp>`;
    return docXml('p:notes', `
<p:cSld><p:spTree>${spTreeSkeleton()}${shape}</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`.trim());
}

function slideXml(ctx, built) {
    const shapes = built.sink.shapes.join('');
    const bg = built.bg ? bgXml(built.bg) : '';
    return docXml('p:sld', `
<p:cSld>${bg}<p:spTree>${spTreeSkeleton()}${shapes}</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`.trim());
}

function coreXml(state) {
    const stamp = state.stamp;
    return `${XML_HEAD}<cp:coreProperties xmlns:cp="${NS_CP}" xmlns:dc="${NS_DC}"`
        + ` xmlns:dcterms="${NS_DCTERMS}" xmlns:dcmitype="${NS_DCMITYPE}" xmlns:xsi="${NS_XSI}">`
        + `<dc:title>${escapeAttr(state.title)}</dc:title>`
        + `<dc:creator>${escapeAttr(state.author)}</dc:creator>`
        + `<cp:lastModifiedBy>${escapeAttr(state.author)}</cp:lastModifiedBy>`
        // category 用来记住主题 id：读回别人的文件时没法从配色反推主题名，自己写的文件可以
        + `<cp:category>${escapeAttr(state.themeId)}</cp:category>`
        + `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>`
        + `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>`
        + '</cp:coreProperties>';
}

function appXml(state, titles, notesCount) {
    const vector = titles.map((title) => `<vt:lpstr>${escapeAttr(title || '幻灯片')}</vt:lpstr>`).join('');
    return `${XML_HEAD}<Properties xmlns="${NS_EP}" xmlns:vt="${NS_VT}">`
        + '<Application>DSH Office Mode</Application>'
        + '<PresentationFormat>宽屏</PresentationFormat>'
        + `<Slides>${titles.length}</Slides>`
        + `<Notes>${notesCount}</Notes>`
        + '<HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop>'
        + '<HeadingPairs><vt:vector size="2" baseType="variant">'
        + '<vt:variant><vt:lpstr>幻灯片标题</vt:lpstr></vt:variant>'
        + `<vt:variant><vt:i4>${titles.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>`
        + `<TitlesOfParts><vt:vector size="${titles.length}" baseType="lpstr">${vector}</vt:vector></TitlesOfParts>`
        + '<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
        + '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>';
}

function contentTypesXml(parts) {
    const defaults = new Map([
        ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
        ['xml', 'application/xml'],
    ]);
    const overrides = [];
    for (const part of parts) {
        if (part.name.endsWith('.rels')) continue;
        if (part.contentType) {
            overrides.push(`<Override PartName="/${part.name}" ContentType="${part.contentType}"/>`);
        }
        if (part.mediaType) {
            const ext = part.name.slice(part.name.lastIndexOf('.') + 1).toLowerCase();
            defaults.set(ext, part.mediaType);
        }
    }
    const defaultXml = [...defaults]
        .map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('');
    return `${XML_HEAD}<Types xmlns="${NS_CT}">${defaultXml}${overrides.join('')}</Types>`;
}

// ───────────────────────────────────────────────────────────────────────────
// 编译：内容模型 → OOXML 包
// ───────────────────────────────────────────────────────────────────────────

function compile(state) {
    const theme = state.theme;
    const pal = makePalette(theme);
    const size = SIZES[state.sizeId];
    const g = makeGeometry(size.cx, size.cy);
    // 全篇字体：主题的字体对 + create({font}) 的覆盖（字符串 = 四个槽全换；
    // 对象 = 只换写到的槽）。titleFont 是「标题字体对」，与 fontScheme 的 majorFont 对齐。
    const font = {
        cn: state.fontOverride.cn ?? theme.fonts.cn,
        en: state.fontOverride.en ?? theme.fonts.en,
        titleCn: state.fontOverride.titleCn ?? theme.fonts.titleCn,
        titleEn: state.fontOverride.titleEn ?? theme.fonts.titleEn,
    };
    const titleFont = { en: font.titleEn ?? font.en, cn: font.titleCn ?? font.cn };
    const warnings = [...state.inputWarnings];
    const warn = (message) => {
        if (!warnings.includes(message)) warnings.push(message);
    };

    const ctx = {
        theme, pal, g, font, titleFont,
        docTitle: state.title, author: state.author, date: state.date,
        page: 0, total: state.slides.length, sectionIndex: 0,
        // 版式渲染时产生的「非致命输入问题」都进这里，最后统一并入 warnings：
        // 非法 fit / 负数尺寸之类的输入错误不该让整次生成失败
        notes: [],
        // 公式解析失败的结构化记录（第四十八轮 P0-1）：stats.formulaErrors 与 read() 侧同名同口径。
        // 回落到普通文字 run 的公式只有这里能证明它「本来是个公式」
        formulaErrors: [],
        // 原生图表部件清单（历史遗留 11-6）：渲染时按页顺序 push，
        // 第 n 个就是 ppt/charts/chartN.xml，幻灯片的 rels 目标按同一编号算。
        chartSpecs: [],
    };

    // 媒体部件按「绝对路径」去重：同一张图在几页里复用同一个 imageN
    const media = new Map();
    const mediaParts = [];
    const takeMedia = (filePath) => {
        const absolute = state.env.resolve(filePath);
        if (media.has(absolute)) return media.get(absolute);
        let entry;
        try {
            const buffer = Buffer.from(state.env.readFile(filePath));
            const info = imageInfo(buffer);
            if (info === undefined) {
                // 文件在、但不是 PNG/JPEG/GIF：交给版式画占位框，报告里说清是「格式不支持」
                warn(`图片格式不支持（只认 PNG/JPEG/GIF）：${filePath}，该页已画成占位框`);
                entry = { unsupported: true, path: asString(filePath) };
            } else {
                const partName = `image${mediaParts.length + 1}.${info.ext}`;
                entry = {
                    ...info,
                    partName,
                    bytes: buffer,
                    bytesLength: buffer.length,
                    path: asString(filePath),
                    mediaType: info.contentType,
                };
                mediaParts.push(entry);
            }
        } catch {
            entry = { missing: true, path: asString(filePath) };
        }
        media.set(absolute, entry);
        return entry;
    };
    ctx.mediaOf = takeMedia;

    // 母版在这里归一（而不是在 deck.master 里）：这样所有 warning 与页面渲染共用同一批
    // ctx.notes，报告里不会出现「输入阶段报了、页面阶段又报一遍」的重复条目
    ctx.total = state.slides.length;
    const master = normalizeMaster(state.master, ctx);
    ctx.master = master;
    preloadMasterMedia(ctx, master, warn);

    const built = [];
    state.slides.forEach((source, index) => {
        const sink = createSink();
        ctx.page = index + 1;
        if (source.kind === 'section') ctx.sectionIndex += 1;
        // 本页几何：有「没给摆位的画布配图」时按分栏收窄版心（新-6）。
        // 逐页算、逐页恢复 —— 下一页从完整版心重新开始
        const plan = planPageGeometry(g, source);
        ctx.g = plan.geom;
        ctx.canvasImageRects = plan.rects;
        // 背景在最底层：本页显式给的优先，否则用母版
        const pageBackground = source.background !== undefined
            ? normalizeBackground(source.background, ctx, `第 ${ctx.page} 页`)
            : (master === undefined ? undefined : master.background);
        // 封面与结束页的标题画在整页上，底色必须是它们自己的 coverBg：
        // 沿用正文的浅色底时，浅色主题的 onPrimary（白字）会落在白底上直接隐形。
        // 显式给的页背景或母版背景优先，这里只补默认值。
        const coverDefault = (source.kind === 'cover' || source.kind === 'closing') && !isBlankFill(ctx.pal.coverBg)
            ? { kind: 'fill', fill: ctx.pal.coverBg }
            : undefined;
        const bg = applyBackground(sink, ctx, pageBackground ?? coverDefault);
        let info;
        if (source.kind === 'cover') {
            info = renderCover(sink, ctx, source);
        } else if (source.kind === 'section') {
            info = renderSection(sink, ctx, source);
        } else if (source.kind === 'bullets') {
            info = renderBullets(sink, ctx, source);
        } else if (source.kind === 'table') {
            info = renderTable(sink, ctx, source);
        } else if (source.kind === 'quote') {
            info = renderQuote(sink, ctx, source);
        } else if (source.kind === 'statement') {
            info = renderStatement(sink, ctx, source);
        } else if (source.kind === 'image') {
            ctx.media = takeMedia(source.path);
            info = renderImage(sink, ctx, source);
            info.media = ctx.media;
            ctx.media = undefined;
        } else if (source.kind === 'images') {
            info = renderImages(sink, ctx, source);
        } else if (source.kind === 'chart') {
            info = renderChart(sink, ctx, source);
        } else if (source.kind === 'shape') {
            info = renderShapePage(sink, ctx, source);
        } else if (HELPER_RENDERERS[source.kind] !== undefined) {
            info = HELPER_RENDERERS[source.kind](sink, ctx, source);
        } else {
            info = renderClosing(sink, ctx, source);
        }
        // 助手页与自由绘制页共用 canvas 通道：助手排版画完之后，叠在上面的
        // banner / panel / shape 再画一遍 —— 一页参考稿就是这两层拼出来的。
        // 自由绘制页（kind='shape'）自己已经把 canvas 画完了，这里不能重复画。
        if (source.kind !== 'shape' && asArray(source.canvas).length > 0) {
            info = mergeCanvasInfo(info, renderCanvasItems(sink, ctx, source.canvas));
        }
        // 母版叠加层在最上层
        masterOverlays(sink, ctx, source, master);
        const layoutIndex = LAYOUT_KEYS.indexOf(layoutPartKeyOf(source.kind));
        const rels = [{
            id: 'rId1',
            type: RT.slideLayout,
            target: `../slideLayouts/slideLayout${layoutIndex + 1}.xml`,
        }, ...sink.rels];
        const report = {
            source, kind: source.kind, info, sink, rels, bg,
            pageBackground, notes: asString(source.notes),
        };
        built.push(report);
        collectWarnings(report, ctx, warn);
    });

    const notes = built.filter((item) => item.notes !== '');
    // notesMaster 的关系挂在所有 slide 关系之后，id 顺延，避免打乱 p:sldIdLst 里已有的 rId
    const notesRelId = notes.length === 0 ? undefined : `rId${built.length + 3}`;
    const parts = [];
    const slideParts = built.map((item, index) => ({
        name: `ppt/slides/slide${index + 1}.xml`,
        contentType: CT.slide,
        data: slideXml({ ...ctx, pal, font }, item),
    }));
    parts.push({
        name: '[Content_Types].xml',
        data: contentTypesXml([
            { name: 'ppt/presentation.xml', contentType: CT.presentation },
            { name: 'ppt/slideMasters/slideMaster1.xml', contentType: CT.master },
            { name: 'ppt/theme/theme1.xml', contentType: CT.theme },
            { name: 'docProps/core.xml', contentType: CT.core },
            { name: 'docProps/app.xml', contentType: CT.app },
            ...LAYOUTS.map((item, index) => ({
                name: `ppt/slideLayouts/slideLayout${index + 1}.xml`, contentType: CT.layout,
            })),
            ...slideParts,
            ...(notes.length === 0 ? [] : [
                { name: 'ppt/notesMasters/notesMaster1.xml', contentType: CT.notesMaster },
                ...notes.map((item, index) => ({ name: `ppt/notesSlides/notesSlide${index + 1}.xml`, contentType: CT.notesSlide })),
            ]),
            ...mediaParts.map((item) => ({ name: `ppt/media/${item.partName}`, mediaType: item.mediaType })),
            // 原生图表：每个图表部件一条 Override（PPT 的图表部件与 Excel 同格式）
            ...ctx.chartSpecs.map((_spec, index) => ({ name: `ppt/charts/chart${index + 1}.xml`, contentType: CT.chart })),
        ]),
    });
    parts.push({
        name: '_rels/.rels',
        data: styleRef(
            `<Relationship Id="rId1" Type="${RT.officeDocument}" Target="ppt/presentation.xml"/>`
            + `<Relationship Id="rId2" Type="${RT.core}" Target="docProps/core.xml"/>`
            + `<Relationship Id="rId3" Type="${RT.app}" Target="docProps/app.xml"/>`,
        ),
    });
    parts.push({ name: 'docProps/core.xml', contentType: CT.core, data: coreXml(state) });
    parts.push({
        name: 'docProps/app.xml',
        contentType: CT.app,
        data: appXml(state, built.map((item) => outlineTitle(item)), notes.length),
    });
    parts.push({
        name: 'ppt/presentation.xml',
        contentType: CT.presentation,
        data: presentationXml({ ...ctx, pal, font }, built.length, size, notesRelId),
    });
    parts.push({
        name: 'ppt/_rels/presentation.xml.rels',
        data: styleRef(
            `<Relationship Id="rId1" Type="${RT.slideMaster}" Target="slideMasters/slideMaster1.xml"/>`
            + `<Relationship Id="rId2" Type="${RT.theme}" Target="theme/theme1.xml"/>`
            + slideParts.map((part, index) => `<Relationship Id="rId${index + 3}" Type="${RT.slide}"`
                + ` Target="${part.name.replace('ppt/', '')}"/>`).join('')
            + (notesRelId === undefined ? '' : `<Relationship Id="${notesRelId}" Type="${RT.notesMaster}"`
                + ' Target="notesMasters/notesMaster1.xml"/>'),
        ),
    });
    parts.push({ name: 'ppt/theme/theme1.xml', contentType: CT.theme, data: themeXml({ ...ctx, pal, font }) });
    parts.push({ name: 'ppt/slideMasters/slideMaster1.xml', contentType: CT.master, data: slideMasterXml({ pal, font }) });
    parts.push({
        name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
        data: styleRef(LAYOUTS.map((item, index) => `<Relationship Id="rId${index + 1}" Type="${RT.slideLayout}"`
            + ` Target="../slideLayouts/slideLayout${index + 1}.xml"/>`).join('')
            + `<Relationship Id="rId${LAYOUTS.length + 1}" Type="${RT.theme}" Target="../theme/theme1.xml"/>`),
    });
    LAYOUTS.forEach((layout, index) => {
        parts.push({
            name: `ppt/slideLayouts/slideLayout${index + 1}.xml`,
            contentType: CT.layout,
            data: slideLayoutXml(ctx, layout),
        });
        parts.push({
            name: `ppt/slideLayouts/_rels/slideLayout${index + 1}.xml.rels`,
            data: styleRef(`<Relationship Id="rId1" Type="${RT.slideMaster}" Target="../slideMasters/slideMaster1.xml"/>`),
        });
    });
    parts.push(...slideParts);
    built.forEach((item, index) => {
        const rels = [...item.rels];
        const notesIndex = notes.indexOf(item);
        if (notesIndex !== -1) {
            rels.push({
                id: `rId${rels.length + 1}`,
                type: RT.notesSlide,
                target: `../notesSlides/notesSlide${notesIndex + 1}.xml`,
            });
        }
        parts.push({
            name: `ppt/slides/_rels/slide${index + 1}.xml.rels`,
            // 超链接是外部关系，必须写 TargetMode="External"，否则 PowerPoint 会在包内找不到目标
            data: styleRef(rels.map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"`
                + `${rel.mode === undefined ? '' : ` TargetMode="${rel.mode}"`}/>`).join('')),
        });
    });
    if (notes.length > 0) {
        parts.push({ name: 'ppt/notesMasters/notesMaster1.xml', contentType: CT.notesMaster, data: notesMasterXml({ pal, font }) });
        parts.push({
            name: 'ppt/notesMasters/_rels/notesMaster1.xml.rels',
            data: styleRef(`<Relationship Id="rId1" Type="${RT.theme}" Target="../theme/theme1.xml"/>`),
        });
        notes.forEach((item, index) => {
            const slideIndex = built.indexOf(item);
            parts.push({
                name: `ppt/notesSlides/notesSlide${index + 1}.xml`,
                contentType: CT.notesSlide,
                data: notesSlideXml({ pal, font }, item.notes),
            });
            parts.push({
                name: `ppt/notesSlides/_rels/notesSlide${index + 1}.xml.rels`,
                data: styleRef(
                    `<Relationship Id="rId1" Type="${RT.slide}" Target="../slides/slide${slideIndex + 1}.xml"/>`
                    + `<Relationship Id="rId2" Type="${RT.notesMaster}" Target="../notesMasters/notesMaster1.xml"/>`,
                ),
            });
        });
    }
    for (const item of mediaParts) {
        parts.push({ name: `ppt/media/${item.partName}`, mediaType: item.mediaType, data: item.bytes });
    }
    // 图表部件（历史遗留 11-6）：与 Excel 共用 chartSpaceXml，编号与 slide rels 一致。
    ctx.chartSpecs.forEach((spec, index) => {
        parts.push({ name: `ppt/charts/chart${index + 1}.xml`, contentType: CT.chart, data: chartSpaceXml(spec) });
    });

    const bytes = zip(parts);
    const layouts = {};
    for (const key of [...LAYOUT_KEYS, 'shape', ...HELPER_KINDS]) layouts[key] = 0;
    for (const item of built) {
        const key = layoutKeyOf(item.kind);
        layouts[key] = (layouts[key] ?? 0) + 1;
    }
    // 版式渲染过程中攒下的输入问题（非法 fit、负数尺寸…）与状态机产生的警告合并去重
    for (const message of ctx.notes) warn(message);
    // 构建期警告双写 env.warn（第四十八轮 P0-1）：office_run 顶层的 warnings 只由
    // env.warnings + 逐文件的 read() 复检警告拼成，不并进来模型就只能看到一半风险
    // —— excel.js 一直是这么做的，ppt.js 与 word.js 缺的就是这一步，于是「公式解析失败」
    // 这类构建期问题在批量脚本里彻底静默。compile() 会被 invalidate() 反复调用，
    // 所以按 state 去重：同一份稿子重编译多少次都只报一遍。
    if (state.env !== undefined && typeof state.env.warn === 'function') {
        state.reportedWarnings ??= new Set();
        for (const message of warnings) {
            if (state.reportedWarnings.has(message)) continue;
            state.reportedWarnings.add(message);
            state.env.warn(`ppt：${message}`);
        }
    }
    const shapeKinds = {};
    for (const item of built) {
        for (const [key, count] of Object.entries(item.info?.kinds ?? {})) {
            shapeKinds[key] = (shapeKinds[key] ?? 0) + count;
        }
    }
    const stats = {
        slides: built.length,
        layouts,
        shapes: built.reduce((sum, item) => sum + item.sink.shapes.length, 0),
        // 自绘形状的细分（preset → 个数）：deck.shape / line / icon / panel / banner 与
        // 版式助手自己画的卡片都算，版式自带的装饰矩形（标题带、页脚线…）不进这个表，
        // 否则「我画了几个形状」这个问题的答案会被色带淹没
        shapeKinds,
        textBoxes: built.reduce((sum, item) => sum + item.sink.textBoxes, 0),
        tables: built.reduce((sum, item) => sum + item.sink.tableCount, 0),
        // 原生图表数（历史遗留 11-6），read() 侧同名同口径
        charts: ctx.chartSpecs.length,
        // 合并区域数（gridSpan/rowSpan > 1 的 origin 格），read() 侧同名同口径
        mergedCells: built.reduce((sum, item) => sum + (item.info?.merges ?? 0), 0),
        // images = 页面上真实放置的图片数（含缺图占位框），imageParts = 去重后的媒体部件数。
        // 两者分开报：同一张图在几页复用会得到 images > imageParts，模型据此判断体积
        images: built.reduce((sum, item) => sum + item.sink.imageCount, 0),
        imageParts: mediaParts.length,
        cropped: built.reduce((sum, item) => sum + (item.info?.cropped ?? 0), 0),
        lowDpi: built.filter((item) => (item.info?.dpi ?? 0) > 0 && item.info.dpi < 96).length,
        notes: notes.length,
        words: built.reduce((sum, item) => sum + item.sink.words, 0),
        // 原生公式数（a14:m 计数），read() 侧同名同口径
        formulas: built.reduce((sum, item) => sum + item.sink.formulas, 0),
        // 解析失败、已回落成普通文字的公式数（第四十八轮 P0-1）。与 formulas 分开报：
        // 「一共有几个公式」与「有几个没写成公式」是两个问题，合在一起等于都答不清楚
        formulaErrors: ctx.formulaErrors.length,
        formulaErrorList: ctx.formulaErrors.slice(0, 8),
        // 看着像公式却写成普通文本的短公式数（第五十八轮 新-23）。read() 侧同名同口径，
        // 与 formulaErrors 也分开：那一条是「解析失败」（改命令），这一条是「没进通道」（改成 math）
        plainFormulas: built.reduce((sum, item) => sum + item.sink.plainFormulas, 0),
        plainFormulaList: built.flatMap((item, index) => item.sink.plainFormulaSamples
            .map((text) => ({ page: index + 1, text }))).slice(0, 8),
        mediaBytes: mediaParts.reduce((sum, item) => sum + item.bytesLength, 0),
        bytes: bytes.length,
    };
    const outline = built.map((item, index) => outlineLine(item, index + 1));
    return { bytes, parts, stats, outline, warnings, built };
}

/** 页面 kind → 逻辑版式 key（stats.layouts 与 outline 用，自由页自成一类）。 */
function layoutKeyOf(kind) {
    return kind === 'images' ? 'image' : kind;
}

/**
 * 页面 kind → 实际挂载的 slideLayout 部件 key。
 * 版式部件只有 8 个（SPEC 写死），自由绘制页只能借用其中一个；
 * 它只决定 slide 的 rId1 指向哪个布局部件，几何全部自带，挂哪个都不影响渲染。
 */
function layoutPartKeyOf(kind) {
    if (kind === 'images') return 'image';
    // 图表借用「陈述」版式（titleOnly）：几何全部自带，挂哪个布局部件都不影响渲染。
    if (kind === 'chart') return 'statement';
    return LAYOUT_KEYS.includes(kind) ? kind : (KIND_LAYOUT_KEY[kind] ?? 'statement');
}

/** 每页一行摘要，给模型观察结构用。 */
function outlineTitle(item) {
    const info = item.info ?? {};
    if (item.kind === 'quote') return asString(info.text).slice(0, 20);
    if (item.kind === 'statement') return asString(info.text).slice(0, 20);
    return asString(info.title);
}

/**
 * 背景/母版的中文摘要，outline 的两处出口共用。
 * 只写「用户看得见东西」：背景图、背景色、logo、页眉、页脚、页码、装饰条条数。
 */
function backgroundLabel(bg) {
    if (bg === undefined) return '';
    return bg.kind === 'image' ? '背景图' : '背景色';
}

function masterLabel(item) {
    const master = item.master;
    if (master === undefined) return '';
    const bits = [];
    const bg = item.pageBackground;
    if (bg !== undefined) bits.push(backgroundLabel(bg));
    if (master.logo !== undefined) bits.push('logo');
    const skip = master.skipLayouts.includes(item.kind);
    if (!skip) {
        if (master.header !== undefined) bits.push('页眉');
        if (master.footer !== undefined) bits.push('页脚');
        if (master.pageNumber !== undefined && master.pageNumber.show === true) bits.push('页码');
    }
    if (master.accents.length > 0) bits.push(`${master.accents.length} 条装饰`);
    return bits.join(' + ');
}

/**
 * 助手页的一行摘要。
 * 生成端（outlineLine）与 read() 端共用这一个函数：两边的版式名与计数口径只有一处，
 * `P3 cards：4 张卡片` 才不会变成「生成说 4 张、读回说 1 个文本框」。
 */
function helperOutline(kind, info) {
    const title = asString(info.title);
    const count = Math.max(0, Math.floor(asNumber(info.count, 0)));
    if (kind === 'cards') return `${title}｜${count} 张卡片`;
    if (kind === 'steps') return `${title}｜${count} 步`;
    if (kind === 'compare') return `${title}｜双栏对比`;
    if (kind === 'kpi') return `${title}｜${count} 个指标`;
    if (kind === 'imageText') return `${title}｜图 + 文`;
    if (kind === 'timeline') return `${title}｜${count} 个节点`;
    if (kind === 'iconGrid') return `${title}｜${count} 个图标`;
    if (kind === 'agenda') return `${title}｜${count} 条`;
    if (kind === 'gallery') return `${title}｜${count} 张图片`;
    if (kind === 'roadmap') return `${title}｜${count} 个阶段`;
    return `${title}｜${count} 个坐标点`; // matrix
}

function outlineLine(item, page) {
    const info = item.info ?? {};
    const head = `P${page} ${item.kind}`;
    const tail = masterLabel(item);
    const suffix = tail === '' ? '' : `｜母版：${tail}`;
    if (HELPER_KINDS.includes(item.kind)) {
        return `${head}：${helperOutline(item.kind, info)}${suffix}`;
    }
    if (item.kind === 'shape') {
        const bits = [];
        if (backgroundLabel(item.pageBackground) !== '') bits.push(backgroundLabel(item.pageBackground));
        bits.push(`${info.shapes} 个形状`);
        if (info.icons > 0) bits.push(`${info.icons} 个图标`);
        return `${head}：${bits.join(' + ')}${suffix}`;
    }
    if (item.kind === 'cover') {
        const bits = [asString(info.title), asString(info.subtitle)].filter((text) => text !== '');
        return `${head}：${bits.join('｜')}${suffix}`;
    }
    if (item.kind === 'section') {
        return `${head}：${asString(info.title)}${asString(info.subtitle) === '' ? '' : `｜${asString(info.subtitle)}`}${suffix}`;
    }
    if (item.kind === 'bullets') {
        return `${head}：${asString(info.title)}｜${info.items.length} 条${info.columns.length > 1 ? '（双栏）' : ''}${suffix}`;
    }
    if (item.kind === 'table') {
        return `${head}：${asString(info.title)}｜${info.rows} 行 × ${info.columns} 列`
            + `${info.merges > 0 ? `（含 ${info.merges} 处合并）` : ''}${suffix}`;
    }
    if (item.kind === 'quote') {
        return `${head}：${asString(info.text).slice(0, 24)}${asString(info.by) === '' ? '' : `｜— ${asString(info.by)}`}${suffix}`;
    }
    if (item.kind === 'statement') {
        return `${head}：${asString(info.text).slice(0, 24)}${asString(info.sub) === '' ? '' : `｜${asString(info.sub)}`}${suffix}`;
    }
    if (item.kind === 'chart') {
        return `${head}：${asString(info.title)}｜${info.type} 图，${info.series} 个系列（原生图表）${suffix}`;
    }
    if (item.kind === 'image' || item.kind === 'images') {
        const head2 = info.grid === true
            ? `${head}（${info.columns} 列 × ${info.items} 图）`
            : head;
        return `${head2}：${asString(info.title)}｜${imageLabel(info)}`
            + `${asString(info.caption) === '' ? '' : `｜${asString(info.caption)}`}${suffix}`;
    }
    return `${head}：${asString(info.title)}${asString(info.subtitle) === '' ? '' : `｜${asString(info.subtitle)}`}${suffix}`;
}

/** 排版风险：全部来自上面量算出来的真实数字。 */
function collectWarnings(item, ctx, warn) {
    const page = ctx.page;
    const info = item.info ?? {};
    if (item.kind === 'bullets') {
        const items = info.items;
        const overflow = info.fitted.some((column) => !column.fits);
        if (items.length > 7) {
            warn(`第 ${page} 页要点 ${items.length} 条，超过 7 条，建议拆成两页`);
        }
        items.forEach((entry, index) => {
            const length = [...entry.text].length;
            if (entry.level > 2) {
                warn(`第 ${page} 页第 ${index + 1} 条要点是 ${entry.level} 级，层级超过 2 级会削弱主次`);
            }
            if (length > 60 && overflow) {
                warn(`第 ${page} 页第 ${index + 1} 条要点 ${length} 字，压到最小 ${info.sizePt}pt 仍然放不下，建议拆条`);
            } else if (length > 60 && info.sizePt < 16) {
                warn(`第 ${page} 页第 ${index + 1} 条要点 ${length} 字，已压到 ${info.sizePt}pt 才放得下`);
            } else if (length > 120) {
                warn(`第 ${page} 页第 ${index + 1} 条要点 ${length} 字，单条过长，建议拆成 2 条`);
            }
        });
        for (const column of info.fitted) {
            if (!column.fits) {
                warn(`第 ${page} 页要点放不下：估算需要 ${Math.ceil(column.neededPt)}pt 高，版心只有 `
                    + `${Math.floor(ctx.g.bodyH / EMU_PER_PT)}pt，已压到最小 ${column.sizePt}pt`);
            }
        }
    }
    if (item.kind === 'table') {
        const layout = info.layout;
        if (layout !== undefined) {
            if (info.columns > 6) {
                warn(`第 ${page} 页表格 ${info.columns} 列，列宽平均不足 ${emuToCm(ctx.g.contentW / info.columns)}cm，建议精简列`);
            }
            if (info.rows > 12) {
                warn(`第 ${page} 页表格 ${info.rows} 行，信息密度偏高，建议拆表或分页`);
            }
            if (layout.compressed) {
                warn(`第 ${page} 页表格按 ${layout.sizePt}pt 排版需要 ${emuToCm(layout.needed)}cm，`
                    + `超过可用 ${emuToCm(ctx.g.bodyH)}cm，已压缩行高到 ${emuToCm(layout.total)}cm`);
            }
        }
    }
    // imageText 复用同一套图片口径：缺图 / 格式不支持 / 被裁切 / DPI 不足都从 info 上读
    // 图片类告警：单图 / 多图 / 图文页走各自的 info，版式页上的**画布配图**走 mergeCanvasInfo
    // 并进来的同一套字段（第五十二轮 新-6）—— 两条路共用一段，save() 与 read() 才不会各报一个数
    if (item.kind === 'image' || item.kind === 'images' || item.kind === 'imageText'
        || (info.images ?? 0) > 0) {
        if (info.missing > 0) {
            const paths = asArray(info.missingPaths).filter((item) => item !== '').join('、');
            warn(`第 ${page} 页有 ${info.missing} 张图片缺失，已画占位框${paths === '' ? '' : `：${paths}`}`);
        }
        if (info.placeheld > 0) {
            warn(`第 ${page} 页有 ${info.placeheld} 张图片格式不支持（只认 PNG/JPEG/GIF），已画占位框`);
        }
        if (info.cropped > 0) {
            // cover 的裁切是「填满」的代价，但用户必须知道有内容被切掉了；
            // 满页背景被裁时说清是背景图，避免与版心内图片的裁切混为一谈
            warn(`第 ${page} 页有 ${info.cropped} 张图片按 cover 铺满`
                + `${info.fullBleed === true ? '（含满页背景图）' : ''}，超出部分已被裁切（左右或上下各切一刀）`);
        }
        if (info.missing === 0 && info.placeheld === 0 && info.dpi > 0 && info.dpi < 96) {
            warn(`第 ${page} 页图片被放到 ${sizeText(info.widthCm, info.heightCm)}（约 ${info.dpi} DPI），低于 96 DPI 会明显发虚`);
        }
        if (info.grid === true && info.items > 6) {
            warn(`第 ${page} 页放了 ${info.items} 张图，${info.columns} 列网格每格约 `
                + `${cm1(ctx.g.contentW / 360000 / info.columns)}cm 宽，建议拆成两页`);
        }
    }
    if (item.kind === 'quote' && !info.fitted?.fits) {
        warn(`第 ${page} 页引用文字放不下：估算需要 ${Math.ceil(info.fitted?.neededPt ?? 0)}pt，版心只有 ${Math.floor((ctx.g.cy * 0.36) / EMU_PER_PT)}pt`);
    }
    if (item.kind === 'statement' && !info.fitted?.fits) {
        warn(`第 ${page} 页宣言文字放不下，已压到最小字号 ${info.fitted?.sizePt}pt`);
    }
    if (item.kind === 'cover' && !info.fitted?.fits) {
        warn(`第 1 页封面标题过长，已压到最小字号 ${info.fitted?.sizePt}pt`);
    }
    // 容量按**整幅版心**算：分栏页（新-6）的 ctx.g 已收窄成左栏，拿它当分母会把
    // 「一页正常内容 + 一张右栏配图」判成压到页脚（复核实测过）
    const capacityW = ctx.g.fullContentW ?? ctx.g.contentW;
    const contentArea = capacityW * (ctx.g.footerY - ctx.g.titleY);
    if (item.sink.contentArea > (ctx.g.cx * ctx.g.cy) * 0.95) {
        warn(`第 ${page} 页文本框总面积占页面 ${Math.round((item.sink.contentArea / (ctx.g.cx * ctx.g.cy)) * 100)}%，几乎没有留白`);
    } else if (item.sink.contentArea > contentArea) {
        // 1 cm² = 1.296e11 EMU²（360000²）—— 这里原来写 /10000/100，报出的数大了 129,600 倍
        const toCm2 = (value) => Math.round(value / 1.296e11 * 100) / 100;
        warn(`第 ${page} 页文本框总面积 ${toCm2(item.sink.contentArea)}cm² 超过版心 `
            + `${toCm2(contentArea)}cm²，可能压到页脚`);
    }
}

// ───────────────────────────────────────────────────────────────────────────
// 对外接口
// ───────────────────────────────────────────────────────────────────────────

export const meta = {
    id: 'ppt',
    name: 'PPT 演示文稿',
    ext: '.pptx',
    summary: '生成可被 PowerPoint 直接打开的 .pptx：封面/章节/要点/表格/引用/陈述/图片/结尾 8 种版式，'
        + '图片支持 contain/cover/natural 三种适配、显式尺寸、裁切、左右对齐、边框、满页背景图与多图网格；'
        + '另有「版式助手」：deck.cards / steps / compare / kpi / imageText / timeline / iconGrid / agenda / gallery / roadmap / matrix'
        + ' 一次调用成形一整页（卡片网格、编号流程、对比双栏、KPI 大数字、图文、时间线、图标网格、目录、图片墙、路线图、2×2 矩阵），'
        + 'deck.panel / deck.banner 则在当前页叠加提示词面板与横幅/色带，可与助手同页拼成参考稿式整页；'
        + '还有「自由绘制」能力：deck.shape / deck.line / deck.icon 可在同一页任意摆放 33 种预设形状、'
        + '渐变/透明填充、虚线箭头、旋转圆角阴影与形状内文字，deck.master 给全篇加背景图（可压暗）、'
        + 'logo、页眉页脚与页码，表格支持 gridSpan/rowSpan 合并、单元格样式与超链接，'
        + '任意段落可用 runs 做段内多色混排；runs 里 {math:\'…\'} 写 LaTeX 公式'
        + '（转原生 OMML，PowerPoint 里可继续编辑），字体可全篇换（create({font})）'
        + '也可逐段逐 run 换；自带主题配色与真实排版风险报告',
    // office_help 的默认层用这一份（缓存提示词预算：全文那份 10 KB 只留给 detail:true）。
    // 签名与参数名必须与下面的全文层保持一致 —— 分层省的是解释与举例，不是接口。
    brief: {
        create: [
            'create({title, theme, path, size, author, date}, env) → deck',
            '  deck.cover/section/closing({title, subtitle, kicker, presenter, date, runs})',
            '  deck.bullets({title, items, columns})   items: string 或 {text, level, runs}',
            '  deck.table({title, columns, rows, headerFill, zebra, border, cellPadCm, rowHeightCm, firstRowBold})',
            '  deck.quote({text, by, runs}) / deck.statement({text, sub, runs})',
            '  deck.image({path, title, caption, source, fit, widthCm, heightCm, align, frame, fullBleed, x, y, w, h, ratio})',
            '    fit: contain（默认）/ cover（裁切）/ natural（96 DPI 原尺寸）',
            '    接在版式页后面 = 与版式同页的配图（助手页上不给 x/y/w/h 时图占右侧 ratio 默认 0.42、版式收到左栏；',
            '    一页只自动分栏一次，第二张不给坐标的配图另起一页）',
            '  deck.images(items, {title, columns, gapCm, fit})   items: [{path, caption, source}]',
            '  deck.chart({type, title, categories, series:[{name, values}]})  原生图表页（bar/column/line/pie/area/scatter）',
            '  ── 版式助手（各新起一页，之后可继续叠元素级助手）──',
            '  deck.cards({title, items, columns, cardFill, cardLine, radius, gapCm, numbered, iconSizeCm, shadow})',
            '  deck.steps({title, items, direction, numbered, arrows, gapCm, fill, line})',
            '  deck.compare({title, left, right, divider, gapCm, cardFill, cardLine})',
            '  deck.kpi({title, items, columns, gapCm, cardFill, cardLine, shadow})   items: [{value,label,unit,color}]',
            '  deck.imageText({title, image, items, text, side, ratio, cardFill, cardLine})',
            '  deck.timeline({title, items, direction, gapCm})   items: [{time,title,body}]',
            '  deck.iconGrid({title, items, columns, gapCm, iconSizeCm})   items: [{icon,label,body}]',
            '  deck.agenda({title, items, numbered, active})   items: [{text, desc, active}]  目录/议程，当前节高亮',
            '  deck.gallery({title, items, columns, ratio, fit})   items: [{path, caption, badge, alt}]  图片墙',
            '  deck.roadmap({title, phases})   phases: [{label, items, progress, accent}]  阶段路线图（progress 0~1）',
            '  deck.matrix({title, xAxis, yAxis, items})   items: [{x,y,text,desc}]  x/y 归一化 0~1  2×2 象限',
            '  ── 元素级助手（画在当前页，与 deck.shape 同级）──',
            '  deck.panel({title, text, runs, x, y, w, h, fill, line, radius, sizePt, align, shadow})',
            '  deck.banner({text, sub, x, y, w, h, fill, gradient, color, sizePt, name})',
            '  deck.shape({preset, x, y, w, h, fill, line, rotate, radius, shadow, text, textOpts})',
            '    x/y/w/h 厘米、原点左上角；preset 见全文层（33 种，未知值回落 rect 并 warning）',
            '  deck.line({from:[x,y], to:[x,y], color, widthPt, dash, alpha, arrow})',
            '  deck.icon({glyph, x, y, sizeCm, color, font})   glyph 如 0xE72C / \'U+E72C\'',
            '  deck.master({background, logo, header, footer, pageNumber, accent, skipLayouts})',
            '  deck.background(spec) 只改当前页 / deck.page() 另起一页 / deck.notes(text)',
            '  deck.render() → Uint8Array / deck.save(path?) → report',
            "  runs: [{text, bold, italic, underline, color, size, font, spacing}]；{math:'E=mc^2'} 插入公式",
            "    短公式（r=qr_0、F=ma、E=mc^2）同样走 math 通道，别写成 r=q·r0 这样的纯文本",
            '    （LaTeX 子集，详见全文层）；整段只有一条 math run 时按展示式（独立成行）排版',
            "  create({font:'楷体'}) 全篇换字体，或 {font:{body, bodyEn, title, titleEn}} 分槽覆盖；",
            '    段落/run 的 font 收字符串或 {en, cn}，逐处覆盖',
            '  ── 三个约定 ──',
            '  1) preset 口语名会映射：flowChartData→flowChartInputOutput、roundedRectCallout→wedgeRoundRectCallout、',
            '     ovalCallout→wedgeEllipseCallout（写原名真实 PowerPoint 打不开）',
            '  2) cover/section/closing 这 3 种版式**不开放自由画布**（整页构图）：',
            '     接在它们后面的 shape/line/icon/panel/banner 会另起一页。bullets/table/quote/statement/image 与',
            '     版式助手页都能续画；deck.image() 接在**版式助手页**后面 = 与版式同页的配图（见全文层）',
            '  3) 填充值 \'FFFFFF\' 与空串 = 不填充（纯白卡片会变透明，要白卡用 F8FAFC 这类极浅灰）',
        ],
        read: [
            'read(path, env) → report：stats（slides/layouts/shapes/tables/images/words…）、',
            '  pages[]（每页 index/layout/title/shapes/images/words/notes…）、outline、warnings',
            '  stats.media[]：包里每个 ppt/media 部件一条',
        ],
        revise: [
            'readSlides(path) → {ok, pages[]}：逐页每个形状 {id, kind, name, text, x, y, w, h, editable}',
            'revise(path, ops) → {ok, saved, applied[], skipped[], warnings[]}   就地改已有 .pptx',
            "  一条 op：{slide, shape, ...动作}；slide: 1 起页码 | '2-5' | '2,4' | [2,3] | 省略=全篇",
            '  shape: 形状 id（数字）| 形状名 | 类型名 title/subtitle/kicker/presenter/date/statement/quote',
            "  动作（可叠加，顺序 style → setText/replace → fit → move/resize）：",
            "    setText: '新文字' / replace: [['旧','新']] / style: {sizePt, bold, color, font, align, text:'子串'}",
            '    move: {x,y} 或 {dx,dy} / resize: {w,h} 或 {dw,dh} / fit: true（厘米，原点左上角）',
            '  一条 op 失败只进 skipped；全都没生效时不落盘（ok:false）。不能增删整页、换版式、改表格结构。',
        ],
        images: [
            'images(path, {out?}) → {count, dir, files[], skipped[], reused}   抽包内嵌图，路径可直接 read_image',
            'insertImages(path, items) → {saved, applied[], skipped[], warnings[]}   往已有稿子里插图',
            '  items: [{slide, path, x, y, wCm, hCm, fit, name, alt}]；fit: contain（默认）/ cover / natural',
            '  也可写成 revise 的一条 op：{slide, addImage:{path, x, y, wCm, hCm, fit}}',
        ],
    },
    methods: {
        create: [
            'create({title, theme, path, size, author, date}, env) → deck',
            "  deck.cover({title, subtitle, kicker, presenter, date, runs})",
            '  deck.section({title, subtitle, runs})',
            '  deck.bullets({title, items, columns})        items: string 或 {text, level, runs}（level 从 1 起，最多建议 2 级）',
            '  deck.table({title, columns, rows, headerFill, zebra, border, cellPadCm, rowHeightCm, firstRowBold})',
            '    columns: string[] 或 [{title, width, type, align}]',
            '    rows 的格子：string 或 {text, runs, colSpan, rowSpan, fill, color, bold, align, link}',
            '    colSpan/rowSpan 写 a:tc 的 gridSpan/rowSpan，被覆盖的格子写 hMerge/vMerge；越界裁剪并 warning',
            '    link 写真正的 a:hlinkClick（关系 TargetMode="External"），未指定颜色时用主题 hlink 色 + 下划线',
            '    headerFill 表头底色 / zebra 隔行底色（true|false|色值）/ border 四边描边 / cellPadCm 单元格内边距',
            '    / rowHeightCm 最小行高 / firstRowBold 表头加粗（默认 true）',
            '  deck.quote({text, by, runs}) / deck.statement({text, sub, runs})',
            '  deck.image({path, title, caption, source, fit, widthCm, heightCm, align, frame, fullBleed, x, y, w, h, ratio})',
            "    fit: 'contain'（默认，等比放进版心）| 'cover'（铺满版心，超出部分用 a:srcRect 裁切）",
            "         | 'natural'（按 96 DPI 折算原始像素尺寸，超出才缩）",
            '    widthCm / heightCm 只给一个时另一个按原图纵横比推算；align: left|center|right（默认 center）',
            '    source 是署名（通常传 office.image.fetch 的 credit）：并进题注那一行（`题注 ｜ 来源：…`）',
            '    frame: true 加细边框；fullBleed: true 铺满整页（忽略页边距与标题区），可当背景图',
            '    ── 与版式共存（第五十二轮 新-6）──',
            '    deck.image() 与版式同页只有两条路（画布开着 ≠ 一定同页）：',
            '      ① 落在**版式助手页**上、不给 x/y/w/h → 分栏：图占版心右侧（ratio 默认 0.42），宿主版式收到左栏；',
            '         一页只自动分栏一次 —— 第二张不给坐标的配图没有第二栏可分，会另起一页（给了坐标仍可同页）；',
            '      ② 给了 x/y/w/h（厘米）→ 按给的摆位画在当前页，越界照旧 warning；',
            '      其余情况（bullets / shape 之后不给坐标）照旧另起一页 —— 要独立成页先 deck.page()，',
            '      或干脆不给坐标（fullBleed: true 永远独立成页）。',
            '    图与题注都并进这一页：stats.images、read().pages[i].images 都会数到它（缺图占位框只算生成端）。',
            '  deck.images(items, {title, columns, gapCm, fit})',
            '    items: [{path, caption, source}]，columns 支持 2 或 3，gapCm 默认 0.4，逐格等比放入并居中',
            '  deck.chart({type, title, categories, series, legend, labels, stacked, gapWidth, colors}) → 原生图表页',
            '    type: bar / column（默认）/ line / pie / area / scatter；series: [{name, values, x?}]（散点图用 x）',
            '    categories 与每个系列的 values 必须等长（不等长当场报错，不会画一半）',
            '    legend: bottom（默认）/ right / left / top / none；labels: true 显示数据标签；stacked: true 堆叠',
            '    这是真正的 DrawingML 图表部件（ppt/charts/chartN.xml）：在 PowerPoint 里是可选中、可改类型、',
            '    可「编辑数据」的图表对象；数据以 numLit / strLit 内联在图表里，所以打开即画、不需要刷新。',
            '  deck.closing({title, subtitle, runs})',
            '  ── 版式助手（每次调用新起一页，自动叠加母版与背景；之后可继续 deck.panel/banner/shape 叠在同一页）──',
            '  deck.cards({title, items, columns, cardFill, cardLine, radius, gapCm, numbered, iconSizeCm, shadow})',
            '    items: [{title, body, accent, icon, runs}]；columns 默认 2（支持 1/2/3/4），其余值 warning 后回落',
            '    runs 是卡片正文的段内混排（body 为空时用它）；body 与 runs 都为空则只画标题',
            '    每张卡 = 圆角矩形（细边框，shadow:true 加阴影）+ 可选 icon（deck.icon 字形）或 numbered 序号圆 + 标题 + 正文',
            '    cardFill 默认取主题 surface；素色主题的 surface 是 FFFFFF（= 不填充），此时回落 F8FAFC，卡片才不会透明',
            '    accent 是卡片强调色（细线 + 序号圆），不给就按主色/副色/强调色轮转；iconSizeCm 默认 0.9',
            '    卡片数量超出 列数×行数（行高不足 1.8cm）时进 warnings',
            '  deck.steps({title, items, direction, numbered, arrows, gapCm, fill, line})',
            "    items: [{title, body}]；direction 'horizontal'（默认）| 'vertical'，其它值 warning 后回落",
            '    numbered 默认 true（画序号圆），arrows 默认 true（步骤之间用 a:tailEnd 三角箭头连起来）',
            '  deck.compare({title, left, right, divider, gapCm, cardFill, cardLine})',
            '    left/right: {title, items, accent}；左右各一张卡，divider 默认画中间竖分隔线（false / \'none\' 关闭，',
            '    传色值则用该颜色）；两侧标题用 accent 色 + 强调细线，不画整条色带（方角压在圆角卡上会露边）',
            '  deck.kpi({title, items, columns, gapCm, cardFill, cardLine, shadow})',
            '    items: [{value, label, unit, color}]；大号数字（54~18pt 阶梯自适应）+ 小标签，unit 作为小号 run 跟在数字后',
            '    columns 默认 min(items.length, 4)（支持 1~4）；数字压到 24pt 以下时进 warnings',
            '  deck.imageText({title, image, items, text, side, ratio, cardFill, cardLine})',
            "    image: {path, fit, caption, source}；side 'left'（默认）| 'right'；ratio 图片占版心宽的比例（默认 0.46，钳制在 0.2~0.8）",
            '    另一侧写要点（items）或整段文字（text）；缺图照画虚线占位框并 warning，图片口径与 deck.image 一致',
            '  deck.timeline({title, items, direction, gapCm})',
            "    items: [{time, title, body}]；direction 'horizontal'（默认）| 'vertical'",
            '    横向 = 一条轴线 + 等距节点圆 + 上下交错的时间/标题/正文（相邻节点纵向不打架）',
            '  deck.iconGrid({title, items, columns, gapCm, iconSizeCm})',
            '    items: [{icon, label, body}]；icon 走 deck.icon 的字形（默认 Segoe MDL2 Assets）',
            '    columns 默认 3（支持 1~6）；每格 = 图标 + 标签 + 说明，全部居中',
            '  deck.agenda({title, items, numbered, active, gapCm, cardFill, cardLine})',
            '    items: [{text, desc, active, runs}]（也认 title/body 两个别名）；numbered 默认 true（序号圆），false 改画左侧色条',
            '    active 标「当前章节」：条目级 active:true 或顶层 active: 页码N（1 起）都行，只高亮第一条；没有则全部普通样式并进 warnings',
            '    用于汇报目录 / 议程页：一行一节、当前节反色高亮；超过 10 条或行高不足 1.05cm 进 warnings',
            '    正例：deck.agenda({title:"目录", items:[{text:"背景",desc:"行业现状"},{text:"方案",active:true},{text:"验证"}]})',
            '  deck.gallery({title, items, columns, gapCm, ratio, fit, cardFill, cardLine})',
            '    items: [{path, caption, badge, alt, source}]；path 相对工作目录，缺图/坏图照画虚线占位框并 warning',
            '    columns 默认 2（支持 1~3）；ratio 是图框高/宽（默认 0.62，钳制 0.2~1.5，1.0 = 正方形）',
            '    badge 是图片角上的小圆角色标（如 "实拍" / "示意图"）；alt 写进图片的替代文字（cNvPr descr）',
            '    用于实拍照片墙 / 作品集：等格网 + 图注，图片解析与 deck.image 同一套（fit 同 contain/cover/natural）',
            '    正例：deck.gallery({title:"现场照片", columns:3, items:[{path:"assets/a.png",caption:"入口",badge:"实拍"},{path:"assets/b.png",caption:"中庭"}]})',
            '  deck.roadmap({title, phases, gapCm, cardFill, cardLine})',
            '    phases: [{label, items, progress, accent}]；items 是 string[]（阶段要点），progress 是 0~1 的完成度（画进度条 + 百分比）',
            '    横排阶段卡（1~5 列，列间自动加箭头）；phase 数超过 5 个按网格换行并 warning，progress 越界钳制后警告',
            '    用于季度路线图 / 实施计划：每阶段一卡，要点 + 完成度一眼看清',
            '    正例：deck.roadmap({title:"实施路线", phases:[{label:"调研",items:["访谈 12 家客户","竞品拆解"],progress:1},{label:"试点",items:["两个场景上线"],progress:0.4},{label:"推广",items:["全量灰度"]}]})',
            '  deck.matrix({title, xAxis, yAxis, items})',
            '    xAxis/yAxis: {low, high}（两端的轴名）；items: [{x, y, text, desc}]，x/y 是 0~1 的归一化坐标（原点左下角）',
            '    用于 2×2 象限分析（如「重要-紧急」「价值-难度」）：四象限浅底 + 交叉中线 + 坐标点与标签',
            '    x/y 缺省按 (0,0) 处理并汇总警告，越界钳到 0~1；超过 12 个点提示标签易互相压住',
            '    正例：deck.matrix({title:"任务优先级", xAxis:{low:"不紧急",high:"紧急"}, yAxis:{low:"低影响",high:"高影响"}, items:[{x:0.82,y:0.8,text:"修复线上故障",desc:"本周"},{x:0.2,y:0.7,text:"规划架构",desc:"下周"}]})',
            '  ── 元素级助手（画在当前页，不新起页，与 deck.shape 同级）──',
            '  deck.panel({title, text, runs, x, y, w, h, fill, line, radius, sizePt, align, shadow})',
            '    浅底圆角面板 + 小字号长文本（参考稿的「提示词示例」）；h 缺省时按真实文字量算，越出页面进 warnings',
            '  deck.banner({text, sub, x, y, w, h, fill, gradient, color, sizePt, name})',
            '    横幅/色带；gradient: {stops:[{pos, color, alpha}], angle} 复用 shape 的 a:gradFill，',
            '    color 缺省时按底色亮度自动选黑/白；x/y/w/h 缺省 = 版心整宽、高 1.4cm',
            '  deck.shape({preset, x, y, w, h, fill, line, rotate, radius, shadow, text, textOpts})',
            '    x/y/w/h 单位厘米、原点在页面左上角，允许负值（色块出血到页边）',
            '    preset: rect/roundRect/ellipse/triangle/rtTriangle/diamond/hexagon/chevron/pentagon/',
            '      rightArrow/leftArrow/upArrow/downArrow/bentArrow/curvedRightArrow/blockArc/donut/pie/',
            '      teardrop/cloud/star5/heart/plaque/frame/halfFrame/corner/flowChartProcess/',
            '      flowChartDecision/flowChartTerminator/flowChartData/roundedRectCallout/ovalCallout',
            '      未知 preset 回落 rect 并 warning；flowChartData/roundedRectCallout/ovalCallout 是口语名，',
            '      实际写进 prstGeom 的是 flowChartInputOutput/wedgeRoundRectCallout/wedgeEllipseCallout',
            '      （DrawingML 里没有同名 prst，写原名会让真实 PowerPoint 打不开文件）',
            "    fill: 'RRGGBB' | {color, alpha:0~1} | {type:'gradient', stops:[{pos:0~1, color, alpha}], angle:度} | 'none'（默认不填充）",
            '      空串与 FFFFFF 一律视为不填充；stops 少于 2 个回落纯色并 warning',
            "    line: 'RRGGBB' | {color, widthPt, dash:'solid'|'dash'|'sysDot'|'lgDash', alpha:0~1,",
            "          arrow:{begin, end}} | 'none'（默认无边框）；未知 dash 回落 solid 并 warning",
            '    rotate 度（a:xfrm rot，60000/度）；radius 圆角半径（厘米，只对 roundRect 生效，写 a:avLst adj）',
            '    shadow: true 或 {blurPt, distPt, dirDeg, color, alpha:0~1} → a:effectLst/a:outerShdw',
            '    text / textOpts:{size, bold, color, align:l|ctr|r, anchor:t|ctr|b, font, lineSpacing, marginCm, runs, paras}',
            '  连续调用 shape/line/icon 会画在同一页；调用 deck.page() 或任意其它版式后另起一页',
            '  deck.line({from:[x,y], to:[x,y], color, widthPt, dash, alpha, arrow})',
            "    arrow: 'none'|'end'|'both'（映射到 a:headEnd/a:tailEnd）或 {begin, end}",
            '  deck.icon({glyph, x, y, sizeCm, color, font})',
            "    glyph 是 Unicode 码点（如 0xE72C）、'U+E72C' 或单个字符；font 默认 'Segoe MDL2 Assets'",
            '  deck.master({background, logo, header, footer, pageNumber, accent, skipLayouts})',
            "    background: {image:{path, fit:'cover'|'contain', dim:0~1, dimColor}} | {color} | {gradient:{stops, angle}} | 'none'",
            "    logo: {path, corner:top-right|top-left|bottom-left|bottom-right, widthCm=1.6, marginCm=0.5}",
            '    header: 文字 或 {text, size=18, bold=true, color, align, x, y, w, font}',
            '    footer: 文字 或 {text, size=10, color, align}',
            "    pageNumber: {show=true, corner=bottom-right, from=2, format='{n}'（支持 {total}）, size=10, color}",
            "    accent: {bar:left-top|left|bottom|top, color, sizeCm=0.4, lengthCm} 或这类对象的数组",
            "    skipLayouts 默认 ['cover','closing']：这两类页不加页眉/页脚/页码（背景、logo、装饰条照加）",
            '    母版对全篇生效，背景在最底层、logo/页眉/页码/装饰条在最上层；配了母版后版式自带的页脚与页码不再画',
            '  ── 三个必须知道的约定 ──',
            '  1) preset 口语名映射：flowChartData→flowChartInputOutput、roundedRectCallout→wedgeRoundRectCallout、',
            '     ovalCallout→wedgeEllipseCallout；DrawingML 里没有这三个同名 prst，写原名真实 PowerPoint 会打不开文件',
            '     且不给错误说明（只报「需要修复」），引擎已静默映射到语义最接近的合法形状',
            '  2) deck.page()：连续 shape/line/icon/panel/banner 画在同一页；要另起一页就用 deck.page() 或任意版式方法，',
            '     也可以先用版式助手（deck.cards 等）再往上叠元素级助手 —— 它们共用同一页的自由画布',
            '  2b) 反过来说：cover / section / closing 这 3 种整页构图版式不开放自由画布，',
            '     接在它们后面的 deck.shape / line / icon / panel / banner 会另起一页空白页',
            '     （导出后表现为「没有标题文本」的页）。bullets / table / quote / statement / image 与',
            '     版式助手页都开放：接在它们后面的**元素级助手**落在同一页。',
            '     deck.image() 是另一条规则：只有「接在版式助手页后面且不给坐标」（自动分栏）或',
            '     「给了 x/y/w/h」才同页，其余情况另起一页 —— 见上面 deck.image 那一段。',
            '     要装饰封面这类页只有两条路：deck.master()（全篇统一）或版式自带参数'
            + '（cover 的 kicker/subtitle/runs、bullets 的 columns、quote 的 by 等）。',
            '     只想要整页背景用 deck.background()（它只改当前页，不会新起页）。',
            '     2026-09-22 实测：把形状接到 cover/section/bullets 后面装饰，41 页的稿子被撑到 67 页、再到 80 页；',
            '     第五十二轮起 bullets 这一类不再另起页，接在后面的元素级助手直接续画。',
            '  3) FFFFFF 与空串 = 不填充：shape/panel/banner 的 fill、卡片底色、背景与表格底色都遵循这一条，',
            '     纯白卡片会变透明；要一张看得见的白卡请用 F8FAFC 这类极浅灰（助手默认值已这样做）',
            '  deck.background(spec)                          只改当前页（写法同母版 background）',
            '  deck.page()                                    结束当前自由画布，下一次 shape/line/icon 另起一页',
            '  deck.notes(text)                               给上一页加备注（可多行）',
            '  deck.render() → Uint8Array（纯函数，不写盘）/ deck.save(path?) → report',
            '  runs: [{text, bold, italic, underline, color, size, font, spacing}]',
            '    spacing 是字距（a:rPr spc，1/100 磅）；runs 可用于段落对象、形状文字、表格单元格与 bullets 条目；',
            '    带自动字号适配的版式（cover/section/quote/statement/bullets）按 runs 里最大字号做溢出估算',
            '  ── 公式（原生 OMML，PowerPoint/WPS/LibreOffice 里都是可编辑公式）──',
            "  run 写 {math:'E=mc^2'}（可与文字 run 混排；size/color 照 run 的规则取）。",
            "    **短公式同样走这条通道**：r=qr_0 要写 {math:'r=qr_0'}，不要写成 r=q·r0 这样的纯文本 ——",
            '    短公式最容易下意识写成文字，而纯文本的公式在 PowerPoint 里不是公式也不能再编辑；',
            '    这样的文本会进 warnings 与 stats.plainFormulas（写侧与读侧同名同口径；',
            '    PPT 按文本节点计，Word 按段落与表格单元格计）',
            '    子集：\\frac \\binom \\sqrt[n]{}、上下标与撇号、\\sum \\prod \\int 等大运算符带上下限',
            '    （\\limits/\\nolimits 改位置）、\\lim \\max 等极限、\\sin \\log 等函数名、\\text \\mathrm \\mathbf \\mathbb',
            '    等样式、\\left\\right 定界符（含 \\lfloor \\langle 与空定界符 .）、\\hat \\vec \\overline \\overrightarrow',
            '    等重音、矩阵族 matrix/pmatrix/bmatrix/vmatrix/Bmatrix、cases、aligned（& 分列、\\\\ 换行）、',
            '    希腊字母与约 150 个常用符号、间距 \\, \\; \\quad 与注释 %',
            '    整段只有一条 math run 时自动按「展示式」排（oMathPara 独立成行，jc 跟段落对齐）；',
            '    解析失败不炸整份文件：按原样排成文字 run，并在 warnings 里报原因与位置',
            '    推荐：行内公式放 bullets/table/cards 的 runs 里；展示式公式用 statement({runs:[{math}]})',
            '    或 shape 的 textOpts.paras / panel —— 一页一条关键等式最常见',
            '  ── 字体（三级可覆盖）──',
            "  ① 全篇：create({font:'楷体'}) 四槽全换；或 {font:{body:'微软雅黑', bodyEn:'Calibri',",
            "     title:'黑体', titleEn:'Georgia'}} 只换写到的槽（内部名 cn/en/titleCn/titleEn 也认）。",
            '     字体名是自由值不校验（机器缺字体时 PowerPoint 自己回落），空值忽略、超 64 字符 warning；',
            '     标题槽作用于封面/章节/内容页/结尾的标题，正文槽作用于其它一切文字',
            '  ② 段落：para.font = 字符串或 {en, cn}（shape 的 textOpts.font、master 的 header.font 走这里）',
            '  ③ run：run.font = 字符串或 {en, cn}（图标字形就靠它指定 Segoe MDL2 Assets）',
            '    字体优先级 run → 段落 → 全篇 → 主题；数学 run 不写字体（公式区由 Cambria Math 接管）',
        ],
        read: [
            'read(path, env) → report（stats/outline/warnings 全部来自真实解析，不是脚本自述）',
            '  report 的键：ok / format / path / bytes / theme / size / stats / pages / outline / warnings',
            '  stats: slides + layouts（各版式页数）+ shapes/autoShapes/textBoxes/tables/images/words/formulas/bytes 等',
            '  stats.media: [{part, contentType, bytes, width, height, slides[]}] —— 包里每个 ppt/media 部件一条',
            '    （与 stats.imageParts 同一批条目；slides 是引用它的页码，空数组表示没有页面引用它）',
            '  pages[]: 每页 {index, layout, title, shapes, autoShapes, tables, tableRows, tableColumns,',
            '    mergedCells, images, cropped, fullBleed, dpi, imageSizes, textBoxes, lines, words, formulas, plainFormulas, notes, helperCount}',
            '  outline: ["P1 cover：标题", ...]；warnings: 排版风险（例如「第 N 页没有标题文本」）',
            '  注意：逐页数组叫 pages，不是 slides（slides 只在 stats 里是页数）。',
        ],
        // 审阅与微调：这两条是「生成之后」的入口，与 create/read 并列在 office.ppt 上。
        revise: [
            'readSlides(path) → {ok, format, path, bytes, pages[], counts}',
            '  逐页列出每个形状：{id, kind, name, text, paragraphs, x, y, w, h, editable}',
            '  x/y/w/h 单位厘米、原点在页面左上角；editable 表示这个形状能不能改文字/字体',
            '  改之前先看这个：拿到 id 与当前几何，再决定动哪一个。',
            '',
            'revise(path, ops) → {ok, path, saved, bytes, applied[], skipped[], warnings[]}',
            '  就地改已有 .pptx：只替换命中的那段 XML，其余部件与形状一个字节都不动。',
            '  一条 op 的写法：{ slide, shape, ...动作 }',
            '',
            '  slide 选择器：页码（1 起）| \'2-5\' | \'2,4\' | [2,3] | 省略 = 全篇',
            "  shape 选择器：形状 id（数字）| 形状名（如 'Card 1 Body'）| 类型名 title/subtitle/kicker/presenter/date/statement/quote",
            '',
            '  动作（同一条 op 里可以叠加，按 style → setText/replace → fit → move/resize 的顺序执行）：',
            "    setText: '新文字'          整段重写。保留原字体字号（沿用原 a:rPr 模板）；'\\n' 拆成多段；",
            '                               p:txBody / a:bodyPr / a:lstStyle 外壳原样保留',
            "    replace: [['旧','新'], ...]  只替换段内命中的子串，段落结构不动；也可写 {find, replace, all}",
            '    style: { sizePt, bold, italic, underline, color, font, align, lineSpacing }',
            '      改字体、字号、颜色、对齐；再加 text: \'子串\' 时只改命中的那个 run（段内混排时用）',
            '    move:   { x, y } 绝对定位，或 { dx, dy } 相对移动（厘米，原点页面左上角）',
            '    resize: { w, h } 或 { dw, dh }（厘米）',
            '    fit:    true                按当前字号与文本框宽度重算高度，让框贴合内容',
            '',
            '  行为约定：',
            '  - 一条 op 失败只进 skipped（带原因），不影响同一批里的其它 op；',
            '  - 同一页命中多个同名形状时会全部处理，并在 warnings 里提示；要只改一个请用 id；',
            '  - 全部 op 都没生效时不落盘，ok:false，避免写出一个内容没变的文件还报成功；',
            '  - 改完形状超出页面会在 warnings 里报出厘米坐标；',
            '  - 单位、原点与 deck.shape 完全一致，readSlides 读回来的几何可以直接填进 move/resize。',
            '',
            '  典型用法：先 read(path) 看版式与 warnings，再 readSlides(path) 取 id 与几何，',
            '  然后一句话改一批：',
            "    office.ppt.revise('汇报.pptx', [",
            "      { slide: 1, shape: 'title', setText: '2026 年度汇报' },",
            "      { slide: '2-4', shape: 'title', style: { sizePt: 30, color: '1F4E79' } },",
            "      { slide: 5, shape: 'Card 2 Body', replace: [['待定', '已确认']] },",
            "      { slide: 5, shape: 'Card 2', move: { x: 2, y: 6 }, resize: { w: 8 } },",
            '    ]);',
            '  不适用的场景：增删整页、换版式、改表格结构 —— 那些请重新生成或直接用 PowerPoint。',
        ],
        // 图片的两个方向：抽出来看（images）与插进去（insertImages）。
        // 生成侧的 deck.image 只能落在「这一趟正在造的页」上（第五十二轮起还能落在刚写过的
        // 版式页上，见新-6），改一份**别人给的**稿子要插图只能靠这两条。
        images: [
            'images(path, {out?}) → {ok, format, path, bytes, count, dir, files[], skipped[], reused, hint}',
            '  把演示文稿里**嵌着**的图片抽出来，落盘后返回可直接交给 read_image 的路径。',
            '  files[]: {part, name, ext, contentType, bytes, width, height, slides[], relIds[], path, reused}',
            '    part 是包内部件名（如 ppt/media/image1.png）；slides 是它出现在第几页；',
            '    reused=true 表示同一张图被多处引用（只落一个文件，slides/relIds 记下全部引用）',
            '  skipped[]: {part, slide?, relId?, reason} —— 外链图片（a:blip/@r:link，字节不在包里）与',
            '    内容类型未知（[Content_Types].xml 里既无 Override 也无对应 Default）的部件都记在这里，不猜不编。',
            '  不传 out 时落插件缓存目录，工作目录不会多出中间文件；抽完直接 read_image 看。',
            '',
            'insertImages(path, items) → {ok, format, path, saved, bytes, applied[], skipped[], warnings[]}',
            '  往已有 .pptx 里插图：就地改包，既有形状一个字节都不动。',
            '  items 可传单个对象或数组：[{slide, path, x, y, wCm, hCm, fit, name, alt}]',
            "    slide 选择器与 revise 一致（1 起页码 / '2-5' / '2,4' / [2,3]；省略 = 全篇）",
            '    x / y 单位厘米、原点页面左上角（与 deck.shape、revise.move 完全一致）',
            "    fit: 'contain'（默认，等比放进 wCm×hCm 的框并居中）| 'cover'（铺满框，超出部分写 a:srcRect 裁掉）",
            "         | 'natural'（按 96 DPI 折算原始像素尺寸，不缩放）",
            '    wCm / hCm 只给一个时另一个按原图纵横比推算；两个都不给 = 原图 96 DPI 的尺寸',
            '    name 是形状名（readSlides 里按它找得到）；alt 写 p:cNvPr/@descr（无障碍替代文字）',
            '  新图顺延 ppt/media/imageN.<ext> 编号，[Content_Types].xml 的 Default 与 slide rels 同步补齐；',
            '  新关系 id = 该页 rels 里最大数字 id + 1，新形状 id = 该页最大 cNvPr id + 1，p:pic 追加在 p:spTree 末尾。',
            '  缺文件或格式不支持（只认 PNG/JPEG/GIF）时进 skipped，且不落盘（saved:false）。',
            '  也可以写成 revise 里的一条 op：{slide, addImage:{path, x, y, wCm, hCm, fit}}，与改文字同一批做完。',
        ],
    },
};

/**
 * 建新演示文稿。env 缺省时用内存环境兜底，方便单测。
 * @param {object} spec
 * @param {object} env
 */
export function create(spec, env) {
    const context = env ?? createEnv({ root: process.cwd() });
    if (typeof context.theme !== 'function') {
        context.theme = (id) => resolveTheme(id);
    }
    return makeDeck(spec, context);
}

/**
 * 就地修订能力（`office.ppt.readSlides` / `office.ppt.revise` / `office.ppt.insertImages`）
 * 与图片导出（`office.ppt.images`）。
 *
 * 修订与插图实现在 `ppt-revise.js`，这里用**动态 import 挂成惰性工厂**而不是在模块顶部静态
 * 引入：两个模块会互相引用（revise 要用 ppt 的枚举口径，ppt 要暴露 revise），
 * 静态引入会形成循环依赖 —— 本仓库已经踩过一次「进程内模块实例陈旧导致新增导出
 * 看不见」的坑（见 README 的已知边界），循环依赖会让它更难诊断。
 * 工厂返回函数，SDK 侧在 wrap() 里调用一次就把这些方法挂到 office.ppt 上。
 *
 * 工厂的第二个参数是插件上下文 `{cache, config}`：`images()` 要往里写导出的图，
 * 缓存目录的位置是配置项，不能硬编码。`readSlides` / `revise` 用不到，忽略即可。
 */
export const api = {
    readSlides: (env) => {
        let impl;
        return async (path) => {
            impl ??= await import('./ppt-revise.js');
            return impl.readSlides(path, env);
        };
    },
    revise: (env) => {
        let impl;
        return async (path, ops) => {
            impl ??= await import('./ppt-revise.js');
            return impl.revise(path, ops, env);
        };
    },
    images: (env, extra) => (path, options) => images(path, options, env, extra),
    insertImages: (env) => {
        let impl;
        return async (path, items) => {
            impl ??= await import('./ppt-revise.js');
            return impl.insertImages(path, items, env);
        };
    },
};

function normalizeSize(value) {
    const text = asString(value, '16:9').trim().toLowerCase().replace(/\s/g, '');
    if (SIZES[text] !== undefined) return text;
    if (text === 'wide' || text === '169' || text === '16x9') return '16:9';
    if (text === 'standard' || text === '43' || text === '4x3') return '4:3';
    return '16:9';
}

/**
 * create({font}) 的归一：全篇字体覆盖。
 * 字符串 = 四个槽（中文/西文正文、中文/西文标题）全换；
 * 对象 = body/bodyEn/title/titleEn（内部名 cn/en/titleCn/titleEn 也认），只换写到的槽。
 * 字体名不校验枚举（PowerPoint 缺字体时自己回落），但空值忽略、超长报 warning。
 */
function normalizeFontOverride(raw, push) {
    const invalid = 'font 只收字符串或 {body, bodyEn, title, titleEn} 对象，已忽略';
    if (raw === undefined || raw === null) return {};
    const clean = (value, label) => {
        const face = asString(value).trim();
        if (face === '') return undefined;
        if (face.length > 64) {
            push(`字体 ${label} 超过 64 个字符，已忽略`);
            return undefined;
        }
        return face;
    };
    if (typeof raw === 'string') {
        const face = clean(raw, 'font');
        if (face === undefined) return {};
        return { cn: face, en: face, titleCn: face, titleEn: face };
    }
    if (typeof raw !== 'object' || Array.isArray(raw)) {
        push(invalid);
        return {};
    }
    const out = {};
    const cn = clean(raw.body ?? raw.cn, 'body');
    const en = clean(raw.bodyEn ?? raw.en, 'bodyEn');
    const titleCn = clean(raw.title ?? raw.titleCn, 'title');
    const titleEn = clean(raw.titleEn, 'titleEn');
    if (cn !== undefined) out.cn = cn;
    if (en !== undefined) out.en = en;
    if (titleCn !== undefined) out.titleCn = titleCn;
    if (titleEn !== undefined) out.titleEn = titleEn;
    return out;
}

function makeDeck(spec, env) {
    const input = spec !== null && typeof spec === 'object' ? spec : {};
    const resolved = env.theme(input.theme);
    const theme = resolved.theme;
    const title = asString(input.title, '未命名演示文稿');
    const sizeId = normalizeSize(input.size);
    // 全篇字体覆盖先归一（warning 要进 state.inputWarnings，所以先建数组再建 state）
    const inputWarnings = [];
    const fontOverride = normalizeFontOverride(input.font, (message) => inputWarnings.push(message));
    const state = {
        env,
        theme,
        themeId: theme.id,
        title,
        author: asString(input.author, ''),
        date: asString(input.date, new Date().toISOString().slice(0, 10)),
        sizeId,
        fontOverride,
        path: input.path === undefined || input.path === null || String(input.path) === ''
            ? ensureExtension(slugify(title, 'presentation'), '.pptx')
            : ensureExtension(String(input.path), '.pptx'),
        slides: [],
        inputWarnings,
        stamp: new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'),
    };
    if (resolved.fellBack) {
        state.inputWarnings.push(`主题 ${asString(resolved.requested)} 不存在，已回落到「${theme.name}」`);
        env.warn(`ppt：主题 ${asString(resolved.requested)} 不存在，已回落到「${theme.name}」`);
    }
    if (asString(input.size, sizeId) !== sizeId) {
        env.warn(`ppt：尺寸 ${asString(input.size)} 不支持（只有 16:9 / 4:3），已用 ${sizeId}`);
    }
    const defaultPath = state.path;
    let dirty = true;
    let cache = null;

    const invalidate = () => {
        dirty = true;
        cache = null;
    };
    // 当前页是否还接着收「叠加层 / 配图」：连续画形状或给版式加配图时续在同一页
    let canvasOpen = false;
    /**
     * 叠加层还能落在哪一页：`undefined` 表示要另起一页。
     * 判据只有两条 —— 画布还开着、且最近这一页是内容型版式（见 CANVAS_HOST_KINDS）。
     */
    const canvasHost = () => {
        const last = state.slides[state.slides.length - 1];
        return canvasOpen && last !== undefined && CANVAS_HOST_KINDS.has(last.kind) ? last : undefined;
    };
    const push = (kind, data) => {
        state.slides.push({ ...data, kind, canvas: [], notes: '' });
        // 落到内容型版式（bullets / table / …）时画布保持打开：接着画的
        // shape / banner / panel / image 会续在同一页；封面 / 章节页 / 结束页则关掉。
        // 满页背景图那一页也关掉 —— 它是整页构图，再叠一张配图只会压在图上看不清
        canvasOpen = CANVAS_HOST_KINDS.has(kind) && !(kind === 'image' && data?.spec?.fullBleed === true);
        invalidate();
        return deck;
    };

    /**
     * 助手页的落点：新起一页，但画布保持打开。
     * 为什么保持打开：助手页与自由绘制页共用同一条 items 通道，
     * 所以 deck.cards() 之后接着 deck.banner()/deck.panel()/deck.shape() 会落在同一页 ——
     * 「一页参考稿」就是助手排版 + 元素级微调拼出来的，拆成两页就不是那张稿子了。
     */
    const pushHelper = (kind, data) => {
        state.slides.push({ ...data, kind, canvas: [], notes: '' });
        canvasOpen = true;
        invalidate();
        return deck;
    };

    /**
     * 自由绘制页 / 叠加层的落点。
     * 连续调用 shape / line / icon / panel / banner 会画在同一页（参考稿那种
     * 「一页里 4 张卡片 + 一条横幅」就是这样拼出来的）；一旦中间插了别的版式，
     * 或显式调了 deck.page()，下一次就开新页。
     */
    const pushCanvas = (type, spec) => {
        const host = canvasHost();
        if (host !== undefined) {
            host.canvas.push({ type, spec });
        } else {
            state.slides.push({ kind: 'shape', canvas: [{ type, spec }], notes: '' });
            canvasOpen = true;
        }
        invalidate();
        return deck;
    };
    const lastSlide = () => state.slides[state.slides.length - 1];
    const compileOnce = () => {
        if (dirty || cache === null) {
            cache = compile(state);
            dirty = false;
        }
        return cache;
    };
    const report = (written) => {
        const built = compileOnce();
        const path = written?.path ?? displayPath(env.root, env.resolve(state.path));
        return {
            ok: true,
            format: 'pptx',
            path,
            bytes: built.bytes.length,
            theme: theme.id,
            stats: built.stats,
            outline: built.outline,
            warnings: built.warnings,
        };
    };

    const deck = {
        format: 'pptx',
        meta,
        theme: theme.id,
        size: sizeId,
        path: defaultPath,

        cover(data) {
            const input2 = data ?? {};
            return push('cover', {
                title: asString(input2.title, state.title) || runText(input2.runs),
                runs: asArray(input2.runs),
                subtitle: asString(input2.subtitle, ''),
                kicker: asString(input2.kicker, ''),
                presenter: asString(input2.presenter, state.author),
                date: asString(input2.date, state.date),
            });
        },
        section(data) {
            const input2 = data ?? {};
            return push('section', {
                title: asString(input2.title) || runText(input2.runs),
                runs: asArray(input2.runs),
                subtitle: asString(input2.subtitle, ''),
            });
        },
        bullets(data) {
            const input2 = data ?? {};
            const items = asArray(input2.items).map((item) => {
                if (item !== null && typeof item === 'object') {
                    const runs = asArray(item.runs);
                    return {
                        // runs 与 text 并存时以 runs 为准：两者一致性由这里保证，后面只读 text
                        text: asString(item.text ?? item.title) || runText(runs),
                        level: asNumber(item.level, 1),
                        runs,
                    };
                }
                return { text: asString(item), level: 1 };
            });
            return push('bullets', {
                title: asString(input2.title),
                items,
                columns: Math.max(1, Math.min(2, Math.floor(asNumber(input2.columns, 1)))),
            });
        },
        table(data) {
            const input2 = data ?? {};
            const columns = asArray(input2.columns);
            const rows = asArray(input2.rows);
            if (columns.length === 0) {
                state.inputWarnings.push(`表格「${asString(input2.title)}」没有列定义，该页只显示标题`);
                env.warn('ppt.table：columns 为空，表格未生成');
            }
            // 表格样式选项原样透传，非法值统一在渲染时落成 warning
            return push('table', {
                title: asString(input2.title),
                columns,
                rows,
                headerFill: input2.headerFill,
                zebra: input2.zebra,
                border: input2.border,
                cellPadCm: input2.cellPadCm,
                rowHeightCm: input2.rowHeightCm,
                firstRowBold: input2.firstRowBold,
            });
        },
        quote(data) {
            const input2 = data ?? {};
            return push('quote', {
                text: asString(input2.text) || runText(input2.runs),
                runs: asArray(input2.runs),
                by: asString(input2.by, ''),
            });
        },
        statement(data) {
            const input2 = data ?? {};
            return push('statement', {
                text: asString(input2.text) || runText(input2.runs),
                runs: asArray(input2.runs),
                sub: asString(input2.sub, ''),
            });
        },
        /**
         * 单图（第五十二轮 `新-6`：与版式共存）。
         *
         * 三种落点，按当前页的状态自动选：
         * 1. 画布开着（刚写过 cards / bullets / shape…）且没给 `fullBleed` → **落当前页**，
         *    成为这一页的配图。没给 `x/y/w/h` 时按「版式 + 配图」分栏：图占版心右侧
         *    （`ratio` 默认 0.42），宿主版式自动收到左栏；给了 `x/y/w/h` 就完全按给的画。
         * 2. 要独立成页：先 `deck.page()` 再调，或第一次调用就是它。
         * 3. `fullBleed: true` 永远是独立的一页（满页背景不该跟版式同页）。
         *
         * `x/y/w/h` 与 `ratio` 只在落当前页时有意义；独立成页时仍走原来的
         * contain / cover / natural + align 那一套。
         */
        image(data) {
            const input2 = data ?? {};
            const file = asString(input2.path);
            // 来源/署名（历史遗留 1-2）：PPT 的图注就是图片下方那一行居中文本，
            // 署名并进同一行（用 ｜ 分隔）—— 页面里没有第二个自然落点，
            // 而「配图无来源」正是要修的事。传 office.image.fetch 返回的 credit。
            const caption = withCredit(asString(input2.caption, ''), input2.source);
            const spec = {
                // 这里刻意只做「原样透传」：非法 fit / 负数尺寸统一在渲染时落成 warning，
                // 这样 save() 之前随时改 spec 都不会让状态变得不可预期
                fit: asString(input2.fit, ''),
                widthCm: input2.widthCm,
                heightCm: input2.heightCm,
                align: asString(input2.align, ''),
                frame: input2.frame === true,
                fullBleed: input2.fullBleed === true,
            };
            // 显式摆位（厘米）：只有四个都缺省才算「没给」，交给分栏规则
            const placed = { x: input2.x, y: input2.y, w: input2.w, h: input2.h, ratio: input2.ratio };
            const given = ['x', 'y', 'w', 'h'].filter((key) => placed[key] !== undefined && placed[key] !== null);
            const host = canvasHost();
            // 与版式同页的两条路（第五十二轮 新-6）：
            // ① 落在**版式助手页**上 —— 助手页只差配图，这是「版式 + 配图」的主路径；
            //    但一页只自动分栏一次：第二张不给坐标的配图没有第二栏可分，照旧另起一页
            //    （否则它会落进已经收窄的版心、压住助手内容）；
            // ② 落在任意开着画布的页上、且调用方给了 x/y/w/h —— 显式摆位就是明确要求同页。
            // 其余情况（bullets 之后直接 image()、自由绘制页上不给摆位）照旧另起一页：
            // 「每次调用一页」是多数稿子的写法，不能因为画布开着就被悄悄改掉。
            const hostHasImage = host !== undefined
                && asArray(host.canvas).some((entry) => entry?.type === 'image');
            const joinHost = host !== undefined
                && (given.length > 0 || (HELPER_KINDS.includes(host.kind) && !hostHasImage));
            // 缺图警告要报**真的落点页码**：与版式同页时是当前这一页，另起一页时才是下一页
            if (file === '' || !env.exists(file)) {
                const page = input2.fullBleed !== true && joinHost ? state.slides.length : state.slides.length + 1;
                state.inputWarnings.push(`第 ${page} 页图片不存在：${file}`);
                env.warn(`ppt.image：找不到图片 ${file}`);
            }
            if (input2.fullBleed !== true && joinHost) {
                host.canvas.push({
                    type: 'image',
                    spec: {
                        ...placed, ...spec, caption, path: file, title: asString(input2.title),
                        source: asString(input2.source, ''), split: given.length === 0,
                    },
                });
                invalidate();
                return deck;
            }
            return push('image', {
                path: file,
                title: asString(input2.title),
                caption,
                source: asString(input2.source, ''),
                spec,
            });
        },
        /**
         * 原生图表页（历史遗留 `11-6`）：真正的 DrawingML 图表部件，不是位图。
         *
         * 与「Python 画图再插进来」的区别：这是 PowerPoint 里的**图表对象** ——
         * 能选中、能改图表类型、能「编辑数据」（数据以 numLit/strLit 内联在图表里，
         * 所以打开即画，不需要刷新）。要改数据走 Office 的「编辑数据」。
         *
         * @param {{type?: 'bar'|'column'|'line'|'pie'|'area'|'scatter', title?: string,
         *   categories?: Array<string|number>, series: Array<{name?: string, values: number[], x?: number[]}>,
         *   legend?: 'bottom'|'right'|'left'|'top'|'none'|false, labels?: boolean, stacked?: boolean,
         *   gapWidth?: number, colors?: string[]}} data
         */
        chart(data) {
            const input2 = data ?? {};
            // 当场校验：参数错了要在这一行报出来，而不是等到 save() 之后
            const spec = normalizeChartSpec(input2);
            return push('chart', { title: asString(input2.title), spec });
        },
        /** 一页多图：等分网格，逐格等比放入。items 为 [{path, caption}]。 */
        images(data, opts) {
            const input2 = data ?? {};
            const options = opts ?? {};
            // 允许 deck.images([...]) 这种省略 opts 的写法
            const items = asArray(input2.items === undefined && Array.isArray(input2) ? input2 : input2.items)
                .map((item) => {
                    if (item === null || typeof item !== 'object') {
                        return { path: asString(item), caption: '' };
                    }
                    return {
                        path: asString(item.path ?? item.src ?? item.file),
                        caption: withCredit(asString(item.caption, ''), item.source),
                        source: asString(item.source, ''),
                    };
                });
            items.forEach((item, index) => {
                if (item.path === '' || !env.exists(item.path)) {
                    state.inputWarnings.push(`第 ${state.slides.length + 1} 页第 ${index + 1} 张图片不存在：${item.path}`);
                    env.warn(`ppt.images：找不到图片 ${item.path}`);
                }
            });
            const columns = asNumber(options.columns, 2) === 3 ? 3 : 2;
            if (options.columns !== undefined && ![2, 3].includes(asNumber(options.columns, 0))) {
                state.inputWarnings.push(`图片网格只支持 2 列或 3 列，已按 ${columns} 列排版`);
            }
            return push('images', {
                title: asString(options.title ?? input2.title),
                items,
                columns,
                gapCm: options.gapCm === undefined ? 0.4 : asNumber(options.gapCm, 0.4),
                fit: asString(options.fit, ''),
            });
        },
        closing(data) {
            const input2 = data ?? {};
            return push('closing', {
                title: asString(input2.title, '谢谢') || runText(input2.runs),
                runs: asArray(input2.runs),
                subtitle: asString(input2.subtitle, ''),
            });
        },
        /** 备注挂在「上一页」上：调用顺序即页面顺序，不需要另传页号。 */
        notes(text) {
            const last = state.slides[state.slides.length - 1];
            if (last === undefined) {
                env.warn('ppt.notes：还没有任何页面，备注已忽略');
                return deck;
            }
            last.notes = asString(text);
            invalidate();
            return deck;
        },
        /**
         * 自由形状（一页可放多个）。
         * 单位是厘米、原点在页面左上角，允许负值 —— 参考稿里的色块经常出血到页边。
         * 几何与填色在渲染时才校验，非法输入只变 warning，不影响整次生成。
         */
        shape(data) {
            return pushCanvas('shape', data ?? {});
        },
        /** 直线：from/to 是 [x,y]（厘米），arrow 支持 none / end / both。 */
        line(data) {
            return pushCanvas('line', data ?? {});
        },
        /** 图标字形：glyph 是码点或单字符，font 默认 Segoe MDL2 Assets。 */
        icon(data) {
            return pushCanvas('icon', data ?? {});
        },
        /** 提示词面板：浅底圆角块 + 小字号长文本（画在当前页）。 */
        panel(data) {
            return pushCanvas('panel', data ?? {});
        },
        /** 横幅 / 色带：纯色或渐变（画在当前页）。 */
        banner(data) {
            return pushCanvas('banner', data ?? {});
        },
        /** 卡片网格：一次调用一整页，每张卡 = 圆角矩形 + 可选图标/序号 + 标题 + 正文。 */
        cards(data) {
            const input2 = data ?? {};
            return pushHelper('cards', {
                title: asString(input2.title),
                items: helperItems(input2.items),
                columns: input2.columns,
                cardFill: input2.cardFill,
                cardLine: input2.cardLine,
                radius: input2.radius,
                gapCm: input2.gapCm,
                numbered: input2.numbered === true,
                iconSizeCm: input2.iconSizeCm,
                shadow: input2.shadow === true,
            });
        },
        /** 编号流程：横向（默认）或纵向，步骤之间用三角箭头连起来。 */
        steps(data) {
            const input2 = data ?? {};
            return pushHelper('steps', {
                title: asString(input2.title),
                items: helperItems(input2.items),
                direction: input2.direction === undefined ? 'horizontal' : asString(input2.direction),
                numbered: input2.numbered,
                arrows: input2.arrows,
                gapCm: input2.gapCm,
                fill: input2.fill,
                line: input2.line,
            });
        },
        /** 对比双栏：左右各一张卡，中间默认一条竖分隔线。 */
        compare(data) {
            const input2 = data ?? {};
            return pushHelper('compare', {
                title: asString(input2.title),
                left: input2.left,
                right: input2.right,
                divider: input2.divider,
                gapCm: input2.gapCm,
                cardFill: input2.cardFill,
                cardLine: input2.cardLine,
            });
        },
        /** KPI 大数字：大号数字 + 单位 + 小标签。 */
        kpi(data) {
            const input2 = data ?? {};
            return pushHelper('kpi', {
                title: asString(input2.title),
                items: helperItems(input2.items),
                columns: input2.columns,
                gapCm: input2.gapCm,
                cardFill: input2.cardFill,
                cardLine: input2.cardLine,
                shadow: input2.shadow === true,
            });
        },
        /** 图 + 文：一侧放图（含题注），另一侧写要点或整段文字。 */
        imageText(data) {
            const input2 = data ?? {};
            const image = input2.image !== null && typeof input2.image === 'object' ? input2.image : {};
            const path = asString(image.path ?? input2.path);
            if (path === '' || !env.exists(path)) {
                state.inputWarnings.push(`第 ${state.slides.length + 1} 页 imageText 图片不存在：${path}`);
                env.warn(`ppt.imageText：找不到图片 ${path}`);
            }
            return pushHelper('imageText', {
                title: asString(input2.title),
                image: {
                    path,
                    fit: asString(image.fit, ''),
                    caption: withCredit(asString(image.caption, ''), image.source),
                },
                items: helperItems(input2.items),
                text: asString(input2.text),
                side: input2.side === undefined ? 'left' : asString(input2.side),
                ratio: input2.ratio,
                cardFill: input2.cardFill,
                cardLine: input2.cardLine,
            });
        },
        /** 时间线：一条轴线 + 等距节点圆 + 上下交错的时间/标题/正文。 */
        timeline(data) {
            const input2 = data ?? {};
            return pushHelper('timeline', {
                title: asString(input2.title),
                items: helperItems(input2.items),
                direction: input2.direction === undefined ? 'horizontal' : asString(input2.direction),
                gapCm: input2.gapCm,
            });
        },
        /** 图标网格：icon 走 deck.icon 的字形，每格 = 图标 + 标签 + 说明。 */
        iconGrid(data) {
            const input2 = data ?? {};
            return pushHelper('iconGrid', {
                title: asString(input2.title),
                items: helperItems(input2.items),
                columns: input2.columns,
                gapCm: input2.gapCm,
                iconSizeCm: input2.iconSizeCm,
            });
        },
        /** 目录页：编号条目纵向排列，active 的那条高亮（只认第一条 active）。 */
        agenda(data) {
            const input2 = data ?? {};
            return pushHelper('agenda', {
                title: asString(input2.title),
                items: helperItems(input2.items),
                numbered: input2.numbered,
                active: input2.active,
                gapCm: input2.gapCm,
                cardFill: input2.cardFill,
                cardLine: input2.cardLine,
            });
        },
        /** 图片墙：等格图片网 + 图注 + 角标，图片走 deck.image 的同一套解析与占位。 */
        gallery(data) {
            const input2 = data ?? {};
            return pushHelper('gallery', {
                title: asString(input2.title),
                items: asArray(input2.items).map((item) => (
                    item !== null && typeof item === 'object'
                        ? { ...item }
                        : { path: asString(item) }
                )),
                columns: input2.columns,
                gapCm: input2.gapCm,
                ratio: input2.ratio,
                fit: input2.fit,
                cardFill: input2.cardFill,
                cardLine: input2.cardLine,
            });
        },
        /** 路线图：横向 3~5 个阶段卡，每卡要点 + 可选完成度。 */
        roadmap(data) {
            const input2 = data ?? {};
            return pushHelper('roadmap', {
                title: asString(input2.title),
                phases: asArray(input2.phases).map((phase) => (
                    phase !== null && typeof phase === 'object' ? { ...phase } : { label: asString(phase) }
                )),
                gapCm: input2.gapCm,
                cardFill: input2.cardFill,
                cardLine: input2.cardLine,
            });
        },
        /** 2×2 矩阵：x/y 归一化坐标点 + 轴标签。 */
        matrix(data) {
            const input2 = data ?? {};
            return pushHelper('matrix', {
                title: asString(input2.title),
                xAxis: input2.xAxis,
                yAxis: input2.yAxis,
                items: helperItems(input2.items),
            });
        },
        /** 结束当前自由画布：下一次 shape/line/icon 会另起一页。 */
        page() {
            canvasOpen = false;
            return deck;
        },
        /** 页面母版（全篇生效）：背景、logo、页眉、页脚、页码、装饰条。多次调用按字段合并。 */
        master(data) {
            const spec = data !== null && typeof data === 'object' ? data : {};
            state.master = { ...(state.master ?? {}), ...spec };
            invalidate();
            return deck;
        },
        /** 只改当前页的背景（写法与母版 background 一致）。 */
        background(spec) {
            const last = lastSlide();
            if (last === undefined) {
                state.inputWarnings.push('deck.background：还没有任何页面，背景已忽略');
                env.warn('ppt.background：还没有任何页面，背景已忽略');
                return deck;
            }
            last.background = spec;
            invalidate();
            return deck;
        },

        /** 纯函数：只按当前内容重新编译一个包，不碰磁盘。 */
        render() {
            return compileOnce().bytes;
        },

        save(path) {
            const target = path === undefined || path === null || path === ''
                ? state.path
                : ensureExtension(String(path), '.pptx');
            state.path = target;
            deck.path = target;
            const bytes = deck.render();
            const written = env.writeFile(target, bytes);
            return report(written);
        },

        /** 给测试与调试用：直接看编译出来的结构，不落盘。 */
        inspect() {
            const built = compileOnce();
            return {
                stats: built.stats,
                outline: built.outline,
                warnings: built.warnings,
                parts: built.parts.map((part) => part.name),
            };
        },
    };
    return deck;
}

// ───────────────────────────────────────────────────────────────────────────
// 读取
// ───────────────────────────────────────────────────────────────────────────

/** 版式中文名 → 语义 key；同时兼容按 layout 的 type 反推。 */
const LAYOUT_NAME_MAP = LAYOUTS.reduce((map, item) => {
    map[item.name] = item.key;
    return map;
}, {});
const LAYOUT_TYPE_MAP = {
    title: 'cover', secHead: 'section', tbl: 'table', picTx: 'image', pic: 'image',
    tx: 'bullets', twoColTx: 'bullets', titleOnly: 'statement', blank: 'statement',
    obj: 'bullets', txAndObj: 'bullets', objAndTx: 'bullets',
};

/** 版式自带的装饰性文本框（标题带、页脚、页码、题注、母版叠加层…），读回时不计入正文行数。 */
const CHROME_SHAPES = new Set([
    'Title', 'Kicker', 'Presenter', 'Date', 'Footer Text', 'Page Number',
    'Caption', 'Cell Caption', 'Quote Mark', 'Section Number', 'Author',
    'Master Header', 'Master Footer', 'Master Page Number',
]);

/**
 * 助手页的「页面标记」。
 * 版式部件一共只有 8 个（SPEC 把 8 个版式名写死），助手页只能借用「陈述」，
 * 所以 read() 没法从 slideLayout 的名字分辨它们，只能靠「只有这个助手才会给形状起的名字」。
 * detect 用来认页面，count 用正则数条目 —— 两者必须分开：卡片的「Card 1」身上还挂着
 * 「Card 1 Rule / Card 1 Title / Card 1 Body」，用前缀数会数出 4 倍。
 */
const HELPER_MARKERS = [
    { kind: 'cards', detect: 'Card ', count: /^Card \d+$/ },
    { kind: 'steps', detect: 'Step Card ', count: /^Step Card \d+$/ },
    { kind: 'compare', detect: 'Compare ', count: /^Compare Column / },
    { kind: 'kpi', detect: 'KPI Card ', count: /^KPI Card \d+$/ },
    { kind: 'imageText', detect: 'ImageText ', count: /^ImageText Panel$/ },
    { kind: 'timeline', detect: 'Timeline Node ', count: /^Timeline Node \d+$/ },
    { kind: 'iconGrid', detect: 'IconGrid ', count: /^IconGrid \d+ Label$/ },
    { kind: 'agenda', detect: 'Agenda Row ', count: /^Agenda Row \d+$/ },
    { kind: 'gallery', detect: 'Gallery Cell ', count: /^Gallery Cell \d+$/ },
    { kind: 'roadmap', detect: 'Roadmap Phase ', count: /^Roadmap Phase \d+$/ },
    { kind: 'matrix', detect: 'Matrix Point ', count: /^Matrix Point \d+$/ },
];

/** 从形状名里认出版式助手页；返回 undefined 表示这不是助手页（普通版式页或自由绘制页）。 */
function helperMarkerOf(names) {
    for (const item of HELPER_MARKERS) {
        if (names.some((name) => name.startsWith(item.detect))) {
            return { kind: item.kind, count: names.filter((name) => item.count.test(name)).length };
        }
    }
    return undefined;
}

function decodePart(decoder, files, name) {
    const data = files.get(name);
    return data === undefined ? undefined : decoder.decode(data);
}

function parseRels(text) {
    const map = new Map();
    if (typeof text !== 'string') return map;
    const root = parseXml(text).children.find((node) => node.name === 'Relationships');
    if (root === undefined) return map;
    for (const node of root.children) {
        if (node.name !== 'Relationship') continue;
        map.set(node.attrs.Id, { target: node.attrs.Target, type: node.attrs.Type, mode: node.attrs.TargetMode });
    }
    return map;
}

/** rels 里的 Target 相对「声明它的部件」解析。 */
function resolveTarget(partName, target) {
    if (typeof target !== 'string' || target === '') return undefined;
    if (target.startsWith('/')) return target.slice(1);
    const segments = partName.split('/').slice(0, -1);
    for (const piece of target.split('/')) {
        if (piece === '.' || piece === '') continue;
        if (piece === '..') segments.pop();
        else segments.push(piece);
    }
    return segments.join('/');
}

/**
 * `[Content_Types].xml` → 扩展名默认表 + 逐部件覆盖表。
 *
 * 图片的内容类型必须**两张表都查**：本引擎自己产出的包走
 * `<Default Extension="png">`，别的工具产出的包却可能只给某个具体部件写
 * `<Override PartName="/ppt/media/image1.png">`。只看 Default 会把这类包里的图
 * 判成「类型未知」，只看 Override 又会漏掉绝大多数包。
 */
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

/** 媒体部件的稳定排序：image1、image2、…、image10 按数字排，不是按字典序。 */
function compareMediaParts(a, b) {
    const left = mediaPartIndexOf(a);
    const right = mediaPartIndexOf(b);
    if (left !== undefined && right !== undefined && left.index !== right.index) {
        return left.index - right.index;
    }
    return a < b ? -1 : (a > b ? 1 : 0);
}

/** 一个部件名 → 它自己的 `.rels` 部件名（`ppt/slides/slide1.xml` → `…/_rels/slide1.xml.rels`）。 */
function relsPartOf(partName) {
    const at = partName.lastIndexOf('/');
    return `${partName.slice(0, at)}/_rels/${partName.slice(at + 1)}.rels`;
}

function shapeText(node) {
    const texts = [];
    for (const run of descendants(node, 'a:t')) texts.push(textOf(run));
    return texts.join('');
}

function maxFontSize(node) {
    let size = 0;
    for (const node2 of descendants(node, 'a:rPr')) {
        const value = Number.parseInt(node2.attrs.sz ?? '0', 10);
        if (Number.isFinite(value)) size = Math.max(size, value);
    }
    for (const node2 of descendants(node, 'a:defRPr')) {
        const value = Number.parseInt(node2.attrs.sz ?? '0', 10);
        if (Number.isFinite(value)) size = Math.max(size, value);
    }
    return size;
}

/** 解析一页：形状清单 + 标题 + 正文行数 + 表格/图片统计。 */
function parseSlide(root) {
    const spTree = descendants(root, 'p:spTree')[0];
    const shapes = [];
    let tables = 0;
    let tableRows = 0;
    let tableCols = 0;
    let charts = 0;
    let images = 0;
    let croppedImages = 0;
    let fullBleed = false;
    let bodyLines = 0;
    // 自绘形状（cNvSpPr 不带 txBox="1" 的 p:sp）：outline 里「N 个形状」的来源
    let autoShapes = 0;
    // 合并区域数：gridSpan/rowSpan > 1 的 origin 格。覆盖格（hMerge/vMerge）不重复计数
    let mergedCells = 0;
    let hasNamedStatement = false;
    // 所有形状名：助手页的识别与条目计数都从这里来（见 HELPER_MARKERS）
    const names = [];
    if (spTree !== undefined) {
        for (const node of xmlChildren(spTree)) {
            if (node.name === 'p:sp') {
                const cNvPr = descendants(node, 'p:cNvPr')[0];
                const label = cNvPr?.attrs?.name ?? node.attrs.name ?? '';
                if (label !== '') names.push(label);
                // 段落级拆分：PPT 的「行数」其实就是段落数 + 每个段落的折行数，
                // 报告里要区分「几段」与「几行」，所以两样都留
                const paragraphs = descendants(node, 'a:p')
                    .map((para) => normalizeSpace(shapeText(para)))
                    .filter((text) => text !== '');
                const text = paragraphs.join(' ');
                const size = maxFontSize(node);
                const isBox = descendants(node, 'p:cNvSpPr').some((item) => item.attrs.txBox === '1');
                if (!isBox) autoShapes += 1;
                if (label === 'Statement') hasNamedStatement = true;
                if (text !== '') {
                    shapes.push({ kind: 'text', name: label, text, size, box: isBox, paragraphs });
                }
            } else if (node.name === 'p:graphicFrame') {
                const table = descendants(node, 'a:tbl')[0];
                // 没有 a:tbl 的 graphicFrame 就是图表框架（历史遗留 11-6）：以前这里
                // 直接把它丢掉 —— 报告里既没有表格也没有图表，模型会以为这页是空的。
                const chart = descendants(node, 'c:chart')[0];
                if (table !== undefined) {
                    const rows = descendants(table, 'a:tr');
                    const cols = descendants(table, 'a:gridCol');
                    tables += 1;
                    tableRows = Math.max(tableRows, rows.length);
                    tableCols = Math.max(tableCols, cols.length);
                    for (const cell of descendants(table, 'a:tc')) {
                        const span = Number.parseInt(cell.attrs.gridSpan ?? '1', 10);
                        const rowSpan = Number.parseInt(cell.attrs.rowSpan ?? '1', 10);
                        if (span > 1 || rowSpan > 1) mergedCells += 1;
                    }
                    const header = rows[0] === undefined ? '' : normalizeSpace(shapeText(rows[0])).slice(0, 40);
                    shapes.push({ kind: 'table', name: '表格', rows: rows.length, columns: cols.length, text: header });
                } else if (chart !== undefined) {
                    charts += 1;
                    const cNvPr = descendants(node, 'p:cNvPr')[0];
                    shapes.push({ kind: 'chart', name: cNvPr?.attrs?.name ?? '图表', relId: chart.attrs['r:id'] ?? '' });
                }
            } else if (node.name === 'p:pic') {
                images += 1;
                const cNvPr = descendants(node, 'p:cNvPr')[0];
                const label = cNvPr?.attrs?.name ?? '图片';
                // 图片的真实几何只能从 p:pic 自己身上读：p:spPr 的 a:xfrm 是绝对 EMU，
                // 用它反推「这张图在纸面上多大、DPI 够不够」，而不是复述生成时的意图
                const xfrm = descendants(node, 'a:xfrm')[0];
                const ext = descendants(xfrm, 'a:ext')[0];
                const w = Number.parseInt(ext?.attrs?.cx ?? '0', 10);
                const h = Number.parseInt(ext?.attrs?.cy ?? '0', 10);
                const srcRect = descendants(node, 'a:srcRect')[0];
                const crop = srcRect === undefined ? 0 : ['l', 't', 'r', 'b']
                    .reduce((sum, key) => sum + Number.parseInt(srcRect.attrs[key] ?? '0', 10), 0);
                if (crop > 0) croppedImages += 1;
                if (label.startsWith('Picture Full Bleed')) fullBleed = true;
                shapes.push({
                    kind: 'image',
                    name: label,
                    text: '',
                    widthCm: w > 0 ? cm1(w / 360000) : 0,
                    heightCm: h > 0 ? cm1(h / 360000) : 0,
                    crop,
                });
            }
        }
    }
    const textShapes = shapes.filter((item) => item.kind === 'text');
    // 助手页识别：形状名里出现「只有这个助手才起的名字」就是它的页面。
    // 这一层必须在下面挑标题之前算完，因为助手页的标题文本框同样叫 Title
    const helper = helperMarkerOf(names);
    // 图片页可能没有名字叫 Title 的文本框：把 Picture / Caption 这类装饰名排除掉，
    // 剩下的最大字号文本框才是真标题
    const titleShape = textShapes.find((item) => item.name === 'Title')
        ?? textShapes.slice().sort((a, b) => b.size - a.size)[0];
    const title = titleShape?.text ?? '';
    // 页眉页脚与装饰性文字不算「正文」：否则报告里的行数会被页码带偏
    const bodyShapes = textShapes.filter((item) => item !== titleShape && !CHROME_SHAPES.has(item.name));
    bodyLines = bodyShapes.reduce((sum, item) => sum + (item.paragraphs?.length ?? 0), 0);
    return {
        shapes, title, tables, tableRows, tableCols, charts, images, croppedImages,
        fullBleed, bodyLines, textShapes, autoShapes, mergedCells, hasNamedStatement,
        helper: helper?.kind, helperCount: helper?.count ?? 0,
        imageShapes: shapes.filter((item) => item.kind === 'image'),
    };
}

/**
 * 读已有文件，返回报告。
 * 版式类型优先看 layout 的中文名，其次看 layout 的 type，最后按内容猜 ——
 * 这样自己生成的文件与别人的文件都能描述清楚。
 *
 * `options.density`（第六十九轮）：对 `bullets` 内容页做**内容密度自检**，把「太空」
 * 与「太密」写进 `warnings`。默认关：测试与内部复用只想验结构时不该被内容口径干扰，
 * 工具面（`office.ppt.read`）由 SDK 传 `{ density: true }` 打开。
 */
export function read(path, env, options = {}) {
    const context = env ?? createEnv({ root: process.cwd() });
    const bytes = context.readFile(path);
    const files = unzip(bytes);
    const decoder = new TextDecoder('utf-8');
    const presentation = decodePart(decoder, files, 'ppt/presentation.xml');
    if (presentation === undefined) throw new Error('pptx: 缺少 ppt/presentation.xml（不是有效的演示文稿）');
    const presRoot = parseXml(presentation).children[0];
    const presRels = parseRels(decodePart(decoder, files, 'ppt/_rels/presentation.xml.rels'));
    const sizeNode = descendants(presRoot, 'p:sldSz')[0];
    const sizeCx = Number.parseInt(sizeNode?.attrs?.cx ?? '0', 10);
    const sizeCy = Number.parseInt(sizeNode?.attrs?.cy ?? '0', 10);
    const slideIdList = descendants(presRoot, 'p:sldIdLst')[0];
    const slideRefs = slideIdList === undefined ? [] : xmlChildren(slideIdList, 'p:sldId');
    const warnings = [];
    const pages = [];
    const layoutCount = {};
    for (const key of [...LAYOUT_KEYS, 'shape', ...HELPER_KINDS]) layoutCount[key] = 0;
    let totalShapes = 0;
    let totalNotes = 0;
    let totalTables = 0;
    let totalCharts = 0;
    let totalImages = 0;
    let totalWords = 0;
    let totalCropped = 0;
    let totalLowDpi = 0;
    let totalAutoShapes = 0;
    let totalMerged = 0;
    let totalFormulas = 0;
    let totalFormulaErrors = 0;
    // 看着像公式却写成普通文本的短公式数（第五十八轮 新-23）
    let totalPlainFormulas = 0;
    // 图片页的展示尺寸与 DPI 从 slide 关系里的媒体部件反推：rels 指向哪个 imageN，
    // 就按那个部件的真实像素数算清晰度，报告里的 DPI 才不是凭空猜的
    const mediaByPart = new Map();
    const readImageInfo = (name) => {
        if (mediaByPart.has(name)) return mediaByPart.get(name);
        const data = files.get(name);
        const info = data === undefined ? undefined : imageInfo(Buffer.from(data));
        mediaByPart.set(name, info);
        return info;
    };
    // 媒体部件的「被谁引用」：rel 循环里顺手记下页码，最后与包内全部 ppt/media 部件合并。
    // 只报计数（imageParts）看不出包里到底是哪几张图、分别在第几页，改稿时无处下手。
    const contentTypes = parseContentTypes(decodePart(decoder, files, '[Content_Types].xml'));
    const mediaSlides = new Map();

    const ordered = slideRefs.length > 0
        ? slideRefs.map((node) => presRels.get(node.attrs['r:id'])?.target)
        : [...files.keys()].filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort(
            (a, b) => Number.parseInt(a.replace(/\D+/g, ''), 10) - Number.parseInt(b.replace(/\D+/g, ''), 10),
        );

    ordered.forEach((target, index) => {
        const partName = resolveTarget('ppt/presentation.xml', target) ?? `ppt/slides/slide${index + 1}.xml`;
        const xml = decodePart(decoder, files, partName);
        if (xml === undefined) {
            warnings.push(`第 ${index + 1} 页部件缺失：${partName}`);
            return;
        }
        const root = parseXml(xml).children[0];
        const parsed = parseSlide(root);
        const relsPart = `${partName.slice(0, partName.lastIndexOf('/'))}/_rels/${partName.slice(partName.lastIndexOf('/') + 1)}.rels`;
        const rels = parseRels(decodePart(decoder, files, relsPart));
        let layout;
        let notes = '';
        const imageInfos = [];
        for (const rel of rels.values()) {
            const resolved = resolveTarget(partName, rel.target);
            if (rel.type === RT.slideLayout) {
                const layoutXml = decodePart(decoder, files, resolved);
                const layoutRoot = layoutXml === undefined ? undefined : parseXml(layoutXml).children[0];
                const cSld = layoutRoot === undefined ? undefined : xmlChildren(layoutRoot, 'p:cSld')[0];
                const name = cSld?.attrs?.name ?? '';
                layout = LAYOUT_NAME_MAP[name]
                    ?? LAYOUT_TYPE_MAP[layoutRoot?.attrs?.type ?? '']
                    ?? undefined;
            } else if (rel.type === RT.notesSlide) {
                const notesXml = decodePart(decoder, files, resolved);
                if (notesXml !== undefined) notes = normalizeSpace(shapeText(parseXml(notesXml).children[0]));
            } else if (rel.type === RT.image) {
                imageInfos.push(resolved === undefined ? undefined : readImageInfo(resolved));
                if (resolved !== undefined && resolved.startsWith('ppt/media/')) {
                    if (!mediaSlides.has(resolved)) mediaSlides.set(resolved, new Set());
                    mediaSlides.get(resolved).add(index + 1);
                }
            }
        }
        if (layout === undefined) {
            // 内容兜底：有表格认表格、有图表认图表、有图认图片，首尾页分别按封面/结尾处理
            if (parsed.tables > 0) layout = 'table';
            else if (parsed.charts > 0) layout = 'chart';
            else if (parsed.images > 0) layout = 'image';
            else if (index === 0) layout = 'cover';
            else if (index === ordered.length - 1) layout = 'closing';
            else layout = 'bullets';
        } else if (layout === 'statement' && parsed.autoShapes > 0 && !parsed.hasNamedStatement) {
            // 自由绘制页、版式助手页与图表页借的都是「陈述」版式部件。页面上没有 Statement
            // 文本框、却有自绘形状时：有图表就是图表页，名字能认出助手就报助手，否则才是自由页。
            // 不这样认的话，read() 会把复刻的参考稿页报成「陈述」，图表页更是会被报成自由页。
            layout = parsed.charts > 0 ? 'chart' : (parsed.helper ?? 'shape');
        }
        layoutCount[layout] = (layoutCount[layout] ?? 0) + 1;
        totalShapes += parsed.shapes.length;
        totalTables += parsed.tables;
        totalCharts += parsed.charts;
        totalImages += parsed.images;
        totalCropped += parsed.croppedImages;
        totalAutoShapes += parsed.autoShapes;
        totalMerged += parsed.mergedCells;
        if (notes !== '') totalNotes += 1;
        // 图片 DPI：只按「能对上媒体部件的图」算，媒体部件缺失（比如被别的工具改坏）时跳过，
        // 不编造一个 DPI 出来
        let pageDpi = 0;
        const dpiList = [];
        parsed.imageShapes.forEach((shape, at) => {
            const info = imageInfos[at];
            if (info === undefined || !(shape.widthCm > 0) || !(info.width > 0)) return;
            dpiList.push(Math.round(info.width / (shape.widthCm / 2.54)));
        });
        if (dpiList.length > 0) {
            pageDpi = Math.round(dpiList.reduce((sum, value) => sum + value, 0) / dpiList.length);
        }
        if (pageDpi > 0 && pageDpi < 96) {
            totalLowDpi += 1;
            warnings.push(`第 ${index + 1} 页图片按当前尺寸只有约 ${pageDpi} DPI，低于 96 DPI 会明显发虚`);
        }
        if (parsed.croppedImages > 0) {
            warnings.push(`第 ${index + 1} 页有 ${parsed.croppedImages} 张图片带 a:srcRect 裁切，部分内容不显示`);
        }
        // 字数按整页所有文本节点算：这样与生成端的 words 口径一致（含表格单元格）
        const textNodes = descendants(root, 'a:t');
        const words = textNodes.reduce((sum, node) => sum + wordsOf(textOf(node)), 0);
        totalWords += words;
        // 公式按 a14:m 计数：自己写的与 PowerPoint 写的文件都是这个包装
        const formulas = (xml.match(/<a14:m[ >]/g) ?? []).length;
        totalFormulas += formulas;
        // 兜底扫描（第四十八轮 P0-1）：解析失败的公式在写入时回落成普通 run，原文进 a:t，
        // 之后再读就只是普通文本 —— 只看 a14:m 的话，这些页会被报成「没有公式、也没有问题」。
        // 判据在 ppt-math 的 looksLikeLatex（要求命中支持域里的命令，Windows 路径不会误报）。
        const strayLatex = textNodes.map((node) => textOf(node)).filter((text) => looksLikeLatex(text));
        if (strayLatex.length > 0) {
            totalFormulaErrors += strayLatex.length;
            warnings.push(`第 ${index + 1} 页有 ${strayLatex.length} 处疑似未解析的 LaTeX 文本`
                + `（如「${strayLatex[0].slice(0, 32)}」）：写入时公式解析失败会回落成普通文字，`
                + '这几处要按 runs:[{ math }] 重写或改写成纯文本');
        }
        // 短公式兜底（第五十八轮 新-23）：**压根没进 math 通道**的短公式（`r=q·r0`）。
        // 上面那条判据要求文本里有 `\命令`，对这一档完全看不见。扫描跳过 mc:Fallback ——
        // 那里是成功公式的纯文本近似（`E = mc²`），算进来就是一条永远甩不掉的假告警。
        const plainFormulaTexts = plainTextNodesOfTree(root)
            .map((node) => textOf(node))
            .filter((text) => looksLikePlainFormula(text));
        totalPlainFormulas += plainFormulaTexts.length;
        if (plainFormulaTexts.length > 0) {
            warnings.push(`第 ${index + 1} 页有 ${plainFormulaTexts.length} 处看着像公式的普通文本`
                + `（如「${plainFormulaTexts[0].trim().slice(0, 32)}」）：短公式同样要走 math 通道，`
                + "写成 runs:[{math:'r=qr_0'}]；确实要当文字用可以忽略这条");
        }
        const page = {
            index: index + 1,
            layout,
            title: parsed.title,
            shapes: parsed.shapes.length,
            autoShapes: parsed.autoShapes,
            tables: parsed.tables,
            tableRows: parsed.tableRows,
            tableColumns: parsed.tableCols,
            mergedCells: parsed.mergedCells,
            charts: parsed.charts,
            images: parsed.images,
            cropped: parsed.croppedImages,
            fullBleed: parsed.fullBleed,
            dpi: pageDpi,
            imageSizes: parsed.imageShapes.map((shape) => sizeText(shape.widthCm, shape.heightCm)),
            textBoxes: parsed.textShapes.length,
            lines: parsed.bodyLines,
            words,
            formulas,
            plainFormulas: plainFormulaTexts.length,
            notes: notes.slice(0, 120),
            helperCount: parsed.helperCount,
        };
        pages.push(page);
        if (page.layout === 'bullets' && page.textBoxes > 1) {
            const box = parsed.textShapes.find((item) => item.name.startsWith('Bullets'));
            const count = box?.paragraphs?.length ?? 0;
            if (count > 7) warnings.push(`第 ${page.index} 页要点 ${count} 条，超过 7 条`);
        }
        // 内容密度自检（第六十九轮）：`bullets` 是承载正文的主力版式，最常见的两种失手
        // 是「一页三个词」（信息量不够、用户要问「为什么这么空」）和「塞满整页」
        // （13pt 往下压、读不动）。阈值按用户偏好定：内容页 120–200 字为目标，
        // 低于 80 字算太空、高于 400 字算太密；只报 bullets，表格/KPI 的「少」通常是设计。
        if (options.density === true && page.layout === 'bullets') {
            if (page.words < 80) {
                // 第七十七轮实测：真实稿里被点名的都是「3 条短要点」页（40 字 / 33 字）。
                // 按用户偏好（每条 15–40 字）这确实是太空，阈值不该放宽；但提示要说清
                // 「平均每行多少字」，用户才知道是补句子还是整页改提纲。
                const perLine = page.lines > 0 ? Math.round(page.words / page.lines) : null;
                const detail = perLine === null ? '' : `（页内正文 ${page.lines} 行，平均 ${perLine} 字/行）`;
                warnings.push(`第 ${page.index} 页只有 ${page.words} 字${detail}：内容页一般每页 120–200 字、`
                    + '每条 15–40 字，建议补数据、对比、归因或责任人；确认要「提纲 / 极简」形态可忽略这条。');
            } else if (page.words > 400) {
                warnings.push(`第 ${page.index} 页 ${page.words} 字偏多：13pt 已是下限，建议拆成两页`
                    + '或删到 200 字以内，并把可表格化的内容改成 deck.table。');
            }
        }
        const prevPage = pages[pages.length - 2];
        // 自由绘制页天生没有标题（标题是版式的构件），只要它不是第 1 页就不再念这一句 ——
        // 「bullets(); deck.page(); shape()」是本轮起推荐的显式换页写法，
        // 再报一句「没有标题文本」只是无法执行的噪音（分栏那三类版式的组告警照旧见下）
        const freeDrawPage = page.layout === 'shape' && prevPage !== undefined;
        if (page.title === '' && page.layout !== 'quote' && page.layout !== 'statement' && !freeDrawPage) {
            warnings.push(`第 ${page.index} 页没有标题文本`);
        }
    });

    // 封面 / 章节页 / 结束页后面紧跟的自由绘制页：这三类版式是整页构图、不开放画布，
    // 接在它们后面的 shape/panel/banner 只会另起一页。这条规则最容易被误解
    // （2026-09-22 实测：调用方为了给版式页加装饰，把 41 页的稿子撑到 67 页、再到 80 页），
    // 所以合成一条带页码的提示，而不是每页各报一次「没有标题文本」。
    // 第五十二轮（新-6）起内容型版式与助手页都开放画布，这条告警只剩这三类版式会命中。
    const strayCanvasPages = pages.filter((page, index) => {
        const prev = pages[index - 1];
        return page.layout === 'shape' && prev !== undefined
            && CANVAS_CLOSED_KINDS.includes(prev.layout);
    });
    if (strayCanvasPages.length > 0) {
        const shown = strayCanvasPages.slice(0, 6).map((page) => `P${page.index}`).join('、');
        warnings.push(`共 ${strayCanvasPages.length} 页是「封面 / 章节页 / 结束页之后的自由绘制页」（${shown}`
            + `${strayCanvasPages.length > 6 ? ' 等' : ''}）：cover/section/closing `
            + '是整页构图、不开放自由画布，接在它们后面的 deck.shape / line / icon / panel / banner 会另起一页。'
            + '要接着画请先 deck.page() 显式换页，整篇统一装饰用 deck.master()，只改当前页背景用 deck.background()。');
    }

    const themeFromCategory = (() => {
        const core = decodePart(decoder, files, 'docProps/core.xml');
        if (core === undefined) return undefined;
        const matched = /<cp:category>([^<]*)<\/cp:category>/.exec(core);
        return matched === null ? undefined : matched[1];
    })();
    const theme = themeFromCategory !== undefined && resolveTheme(themeFromCategory).theme.id === themeFromCategory
        ? themeFromCategory
        : DEFAULT_THEME_ID;
    const outline = pages.map((page) => {
        // 助手页：版式名与计数都走 helperOutline，与生成端 outlines 一字不差
        if (HELPER_KINDS.includes(page.layout)) {
            return `P${page.index} ${page.layout}：${helperOutline(page.layout, {
                title: page.title, count: page.helperCount,
            })}`;
        }
        if (page.layout === 'table') {
            return `P${page.index} table：${page.title}｜${page.tableRows} 行 × ${page.tableColumns} 列`
                + `${page.mergedCells > 0 ? `（含 ${page.mergedCells} 处合并）` : ''}`;
        }
        if (page.layout === 'image') {
            const sizes = page.imageSizes.join('、');
            return `P${page.index} image：${page.title}｜${page.images} 张图`
                + `${sizes === '' ? '' : `｜${sizes}`}${page.cropped > 0 ? '（有裁切）' : ''}`
                + `${page.fullBleed ? '（满页背景）' : ''}`;
        }
        if (page.layout === 'shape') {
            return `P${page.index} shape：${page.autoShapes} 个形状`
                + `${page.fullBleed ? ' + 背景图' : ''}${page.textBoxes > 0 ? `｜${page.textBoxes} 个文本框` : ''}`;
        }
        if (page.layout === 'bullets') return `P${page.index} bullets：${page.title}｜${page.textBoxes} 个文本框`;
        return `P${page.index} ${page.layout}：${page.title}`;
    });
    const unique = [...new Set(warnings)];
    for (const message of unique) context.warn(message);
    // 包内全部媒体部件（含没有被任何页面引用的：母版/版式里的图也算在包里），
    // 与 imageParts 是同一批条目 —— 两个数字必须一致，否则报告自相矛盾
    const media = [...files.keys()]
        .filter((name) => name.startsWith('ppt/media/'))
        .sort(compareMediaParts)
        .map((part) => {
            const data = files.get(part);
            const info = readImageInfo(part);
            return {
                part,
                contentType: contentTypeOf(contentTypes, part) ?? null,
                bytes: data === undefined ? 0 : data.length,
                width: info?.width ?? null,
                height: info?.height ?? null,
                slides: [...(mediaSlides.get(part) ?? [])].sort((a, b) => a - b),
            };
        });
    return {
        ok: true,
        format: 'pptx',
        path: asString(path),
        bytes: bytes.length,
        theme,
        size: sizeCx === 9144000 ? '4:3' : '16:9',
        stats: {
            slides: pages.length,
            layouts: layoutCount,
            shapes: totalShapes,
            // 自绘形状与合并区域：与 create 侧的 stats 同名同口径，两个方向都报同一件事
            autoShapes: totalAutoShapes,
            mergedCells: totalMerged,
            textBoxes: pages.reduce((sum, page) => sum + page.textBoxes, 0),
            tables: totalTables,
            // 原生图表（历史遗留 11-6）：生成端与 read() 侧同一口径
            charts: totalCharts,
            images: totalImages,
            cropped: totalCropped,
            lowDpi: totalLowDpi,
            imageParts: media.length,
            // 逐张报清楚：部件名、内容类型、体积、像素尺寸、出现在第几页
            media,
            notes: totalNotes,
            words: totalWords,
            // 原生公式数（a14:m 计数）：生成端与 read() 侧同一口径
            formulas: totalFormulas,
            // 疑似未解析的 LaTeX 文本数（兜底扫描，见上面的注释）：生成端与 read() 侧同一口径
            formulaErrors: totalFormulaErrors,
            // 看着像公式的普通文本数（第五十八轮 新-23）：与 formulaErrors 分开，见 createSink 的注释
            plainFormulas: totalPlainFormulas,
            bytes: bytes.length,
        },
        pages,
        outline,
        warnings: unique,
    };
}

/**
 * 把演示文稿里**嵌着**的图片抽出来，返回可直接交给 `read_image` 的路径。
 *
 * 为什么默认落缓存目录而不是工作目录：抽图是「下一步要看它」的中间动作，不是交付物，
 * 而「工作目录里不出现临时文件」是本仓库的一条硬规则（见 engine/cache.js 的说明）。
 * 目录一律走注入进来的 cache 对象，绝不硬编码 `.office/runtime/cache` 这类路径 ——
 * 缓存目录的位置是插件配置项，写死一处就会在改配置时静默失效。
 *
 * 三件事照实报，不糊过去：
 * - 同一个部件被多处引用只落一个文件，`slides` / `relIds` 记下全部引用（第二处起 `reused: true`）；
 * - 用 `a:blip/@r:link` 链到外部的图片，字节根本不在包里，进 `skipped`，绝不编造字节；
 * - 内容类型既不在 `<Default Extension>` 也不在 `<Override PartName>` 里的部件同样进
 *   `skipped` —— 猜一个 MIME 写进返回值，等于把「读不出来」伪装成「读出来了」。
 *
 * @param {string} path .pptx 路径
 * @param {{out?: string}} [options] out 给出时写那个目录，否则写插件缓存目录
 * @param {object} env 文件环境
 * @param {{cache?: object}} [extra] 插件上下文（缓存目录从这里来）
 */
export function images(path, options, env, extra) {
    const context = env ?? createEnv({ root: process.cwd() });
    const opts = options !== null && typeof options === 'object' ? options : {};
    const cache = extra !== undefined && extra !== null ? extra.cache : undefined;
    const bytes = context.readFile(path);
    const files = unzip(bytes);
    const decoder = new TextDecoder('utf-8');
    const presentation = decodePart(decoder, files, 'ppt/presentation.xml');
    if (presentation === undefined) throw new Error('pptx: 缺少 ppt/presentation.xml（不是有效的演示文稿）');

    const contentTypes = parseContentTypes(decodePart(decoder, files, '[Content_Types].xml'));
    const presRoot = parseXml(presentation).children[0];
    const presRels = parseRels(decodePart(decoder, files, 'ppt/_rels/presentation.xml.rels'));
    const idList = descendants(presRoot, 'p:sldIdLst')[0];
    const refs = idList === undefined ? [] : xmlChildren(idList, 'p:sldId');
    const ordered = refs.length > 0
        ? refs.map((node) => presRels.get(node.attrs['r:id'])?.target)
        : [...files.keys()].filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort(
            (a, b) => Number.parseInt(a.replace(/\D+/g, ''), 10) - Number.parseInt(b.replace(/\D+/g, ''), 10),
        );

    const skipped = [];
    const found = new Map();
    const payloads = new Map();
    const reported = new Set();
    // 每个部件记下「哪一页的哪条关系引用过它」：同一处引用重复出现不算复用，
    // 换一页或换一条关系才算（rId 只在单页的 rels 命名空间里唯一，不能单独当键）
    const refKeys = new Map();
    const skipOnce = (key, entry) => {
        if (reported.has(key)) return;
        reported.add(key);
        skipped.push(entry);
    };
    const remember = (part, pageNumber, relId) => {
        const key = `${pageNumber ?? ''}#${relId ?? ''}`;
        let entry = found.get(part);
        if (entry === undefined) {
            const data = files.get(part);
            const info = imageInfo(Buffer.from(data));
            entry = {
                part,
                name: '',
                ext: part.slice(part.lastIndexOf('.') + 1).toLowerCase(),
                contentType: contentTypeOf(contentTypes, part),
                bytes: data.length,
                width: info?.width ?? null,
                height: info?.height ?? null,
                slides: [],
                relIds: [],
                path: '',
                reused: false,
            };
            found.set(part, entry);
            payloads.set(part, Buffer.from(data));
            refKeys.set(part, []);
        } else if (!refKeys.get(part).includes(key)) {
            entry.reused = true;
        }
        if (!refKeys.get(part).includes(key)) refKeys.get(part).push(key);
        if (pageNumber !== undefined && !entry.slides.includes(pageNumber)) entry.slides.push(pageNumber);
        if (relId !== undefined && !entry.relIds.includes(relId)) entry.relIds.push(relId);
    };

    ordered.forEach((target, index) => {
        const pageNumber = index + 1;
        const partName = resolveTarget('ppt/presentation.xml', target) ?? `ppt/slides/slide${pageNumber}.xml`;
        const xml = decodePart(decoder, files, partName);
        if (xml === undefined) return;
        const rels = parseRels(decodePart(decoder, files, relsPartOf(partName)));
        // 这一页里用 r:link 引用的关系 id：图片是「链出去的」，包里没有字节
        const linked = new Set();
        for (const match of xml.matchAll(/<a:blip\b[^>]*\sr:link="([^"]+)"/g)) linked.add(match[1]);
        for (const [relId, rel] of rels) {
            if (rel.type !== RT.image) continue;
            const external = linked.has(relId) || rel.mode === 'External';
            const resolved = external ? undefined : resolveTarget(partName, rel.target);
            if (external) {
                skipOnce(`${partName}#${relId}`, {
                    part: asString(rel.target),
                    slide: pageNumber,
                    relId,
                    reason: '图片是外部链接（a:blip/@r:link），字节不在包里，抽不出来；需要图片请让作者把它嵌入',
                });
                continue;
            }
            if (resolved === undefined || !files.has(resolved)) {
                skipOnce(`${partName}#${relId}`, {
                    part: asString(rel.target),
                    slide: pageNumber,
                    relId,
                    reason: `关系指向的部件不在包里：${asString(rel.target)}`,
                });
                continue;
            }
            if (contentTypeOf(contentTypes, resolved) === undefined) {
                skipOnce(`ct:${resolved}`, {
                    part: resolved,
                    slide: pageNumber,
                    relId,
                    reason: '[Content_Types].xml 里既没有这个部件的 Override，也没有对应扩展名的 Default，内容类型未知',
                });
                continue;
            }
            remember(resolved, pageNumber, relId);
        }
    });

    // 包里有、但没有任何页面引用的媒体（母版/版式里的图、或改坏了的残留）：
    // 一并抽出来，slides 为空数组。否则 count 会比 read() 的 imageParts 少，看着像漏了。
    for (const part of [...files.keys()].filter((name) => name.startsWith('ppt/media/')).sort(compareMediaParts)) {
        if (found.has(part)) continue;
        if (contentTypeOf(contentTypes, part) === undefined) {
            skipOnce(`ct:${part}`, {
                part,
                reason: '[Content_Types].xml 里既没有这个部件的 Override，也没有对应扩展名的 Default，内容类型未知',
            });
            continue;
        }
        remember(part, undefined, undefined);
    }

    // 输出目录：显式 out 优先，否则落缓存。缓存目录由插件注入，这里不写死任何路径。
    // 缓存桶带上包内容哈希（S-05）：从前只按 basename 前 24 个字符分目录，
    // 两份同名但内容不同的稿子会落进同一个目录，后一次导出直接覆盖前一次，
    // 文件清单里还看不出覆盖发生过。
    const stem = (basename(asString(path)).replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]/g, '-').slice(0, 24) || 'deck');
    const digest = createHash('sha256').update(bytes).digest('hex').slice(0, 10);
    let outDir;
    let intoCache = false;
    if (typeof opts.out === 'string' && opts.out.trim() !== '') {
        // 与 office.image.fetch / archive.extract 同口径：相对路径不得逃出工作目录，
        // 绝对路径视为调用方显式选择。
        outDir = assertInsideRoot(opts.out.trim(), context, 'office.ppt.images');
    } else if (cache !== undefined && cache !== null && typeof cache.ensureDir === 'function') {
        outDir = cache.ensureDir(`ppt-images/${stem}-${digest}`);
        intoCache = true;
    } else {
        throw new Error('office.ppt.images：没有可写的目录 —— 请给 { out }，或在 office_run 里调用（缓存目录会自动注入）');
    }

    const list = [];
    // 覆盖必须是调用方明说的决定：显式 out 指向用户已有文件时，默认跳过而不是悄悄改写；
    // 缓存目录同理 —— 同一份稿子重复导出是常态，第二次直接复用已导出的文件即可。
    const overwrite = opts.overwrite === true;
    for (const entry of found.values()) {
        // 部件名带上演示文稿的词干：缓存目录是多份文档共用的，只用 image1.png 会互相覆盖
        entry.name = `${stem}-${entry.part.slice('ppt/media/'.length)}`;
        const target = join(outDir, entry.name);
        if (!overwrite && context.exists(target)) {
            if (intoCache) {
                // 缓存桶按包内容哈希分桶，命中说明这份稿子已经导出过：文件字节相同，
                // 直接复用，不重写；同时把路径放回 files，调用方仍能 read_image。
                entry.path = displayPath(context.root, target);
                entry.slides.sort((a, b) => a - b);
                entry.existing = true;
                if (typeof cache.noteHit === 'function') cache.noteHit(entry.bytes);
                list.push(entry);
                continue;
            }
            skipped.push({
                part: entry.part,
                reason: '目标文件已存在，未覆盖（需要覆盖请在 options 里传 overwrite: true）',
                path: displayPath(context.root, target),
            });
            continue;
        }
        // 落盘统一走 env.writeFile：父目录自动创建、写进 writes 台账，
        // office_run 的统一结果清单与产物复检才会包含这些导出的图。
        context.writeFile(target, payloads.get(entry.part));
        if (intoCache && typeof cache.noteArtifact === 'function') cache.noteArtifact();
        entry.path = displayPath(context.root, target);
        entry.slides.sort((a, b) => a - b);
        list.push(entry);
    }

    return {
        ok: true,
        format: 'pptx',
        path: asString(path),
        bytes: bytes.length,
        count: list.length,
        dir: displayPath(context.root, outDir),
        files: list,
        skipped,
        overwrite,
        reused: list.filter((item) => item.reused).length,
        hint: '图片已导出，直接 read_image({ file_path }) 逐张看；files[].slides 说明它出现在第几页，'
            + 'slides 为空表示包里没有页面引用它（母版/版式里的图或残留部件）。'
            + '抽的是「嵌在包里的」图片：r:link 链到外部的图没有字节，已在 skipped 里说明。'
            + '显式 out 里已有同名文件时默认跳过（skipped 里能看到路径），确实要覆盖请传 overwrite: true；'
            + '默认缓存目录按包内容哈希分桶，同一份稿子重复导出会复用已有文件（existing: true）。',
    };
}

/**
 * PPT 不做编辑：PowerPoint 的排版是逐页形状的绝对定位，改一处文本就要重排该页所有形状，
 * 「批量替换文字」在 PPT 上的用户预期本来就低。保留一个显式的不支持反馈，
 * 避免调用方拿到 `undefined is not a function` 这种没有信息量的错误。
 */
export function edit(path, ops, env) {
    const context = env ?? createEnv({ root: process.cwd() });
    const requested = asArray(ops).length;
    const reason = 'PowerPoint 版式是逐页绝对定位，本引擎只提供 create / read；改文字请重新生成或直接编辑 .pptx';
    context.warn(`ppt.edit 不支持：${reason}`);
    return {
        ok: false,
        format: 'pptx',
        path: asString(path),
        applied: [],
        skipped: Array.from({ length: requested }, (_, index) => ({ at: index, reason })),
        warnings: [reason],
    };
}

