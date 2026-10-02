/**
 * Excel 格式引擎（.xlsx / SpreadsheetML）。
 *
 * 设计取舍（为什么这样做）：
 * 1. 字符串一律 inlineStr。sharedStrings 会多出一个部件、多一处索引错位的风险，
 *    而办公场景的表通常只有几百行，省下的体积不值这个失败面。
 * 2. 不写 `xl/theme/theme1.xml`。OOXML schema 里 theme 是可选部件，Excel 会回落到
 *    内置主题；主题的颜色与字体由本模块自己渲染成 styles.xml 里的具体值。
 * 3. 布局走「虚拟落笔 → 文本估宽」两道工序。所有 stats / warnings 都来自这两道
 *    真实计算，而不是写死的模板文字。
 * 4. 字符串、百分比、货币、日期写成真正的数字 + numFmt，这样 Excel 里还能排序、
 *    求和、画图，而不是一堆「看起来像数字的文本」。
 *
 * @module dsh-office-mode/formats/excel
 */
import { zip, unzip, unzipText } from '../engine/zip.js';
import { parseXml, descendants, children as xmlChildren, textOf } from '../engine/xml.js';
import {
    asArray, asString, classifyValue, createEnv, ensureExtension, estimateEm,
    normalizeColumns, normalizeRows, slugify, toNumber,
} from '../engine/kit.js';
import { isAbsolute, relative, sep } from 'node:path';
import { DEFAULT_THEME_ID, resolveTheme, shadeOf, toHex } from '../engine/theme.js';
import {
    CHART_CONTENT_TYPE, DRAWING_CONTENT_TYPE as CHART_DRAWING_CONTENT_TYPE,
    REL_TYPE_CHART as CHART_REL_CHART, REL_TYPE_DRAWING as CHART_REL_DRAWING,
    chartSpaceXml, describeChartSpace, normalizeChartSpec, spreadsheetDrawingXml,
} from './chart.js';

// ───────────────────────────────────────────────────────────────────────────
// 常量：OOXML 里写死的固定部分
// ───────────────────────────────────────────────────────────────────────────

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_DOC_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

/** 自定义数字格式 id 的起点：164 以下被内置格式占用。 */
const CUSTOM_FMT_BASE = 164;

/**
 * 内置 numFmtId → 格式串。
 * Excel 把一部分常用格式编进了编号（0–49 段），这些不必在 styles.xml 里声明；
 * 反过来读别人的文件时，要把编号翻回格式串才能判断「这格是不是日期」。
 */
const BUILTIN_NUMFMTS = new Map([
    [0, 'General'], [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'],
    [9, '0%'], [10, '0.00%'], [11, '0.00E+00'], [12, '# ?/?'], [13, '# ??/??'],
    [14, 'm/d/yyyy'], [15, 'd-mmm-yy'], [16, 'd-mmm'], [17, 'mmm-yy'],
    [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'], [20, 'h:mm'], [21, 'h:mm:ss'],
    [22, 'm/d/yyyy h:mm'], [37, '#,##0 ;(#,##0)'], [38, '#,##0 ;[Red](#,##0)'],
    [39, '#,##0.00;(#,##0.00)'], [40, '#,##0.00;[Red](#,##0.00)'],
    [45, 'mm:ss'], [46, '[h]:mm:ss'], [47, 'mmss.0'], [48, '##0.0E+0'], [49, '@'],
]);

/** 常用格式串 → 内置编号，命中时不必新增 numFmt 部件。 */
const BUILTIN_BY_CODE = new Map([
    ['0', 1], ['0.00', 2], ['#,##0', 3], ['#,##0.00', 4], ['0%', 9], ['0.00%', 10], ['@', 49],
]);

/** 单元格里的字面错误值。出现在 inlineStr 里说明是手工粘进来的文本，不是真公式结果。 */
const ERROR_LITERALS = new Set([
    '#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NULL!', '#NUM!', '#GETTING_DATA',
]);

/** 每种「真类型」的默认数字格式。调用方没指定时才兜底。 */
const DEFAULT_FORMAT = {
    int: '#,##0',
    float: '#,##0.00',
    percent: '0.0%',
    currency: '¥#,##0.00',
    date: 'yyyy-mm-dd',
    datetime: 'yyyy-mm-dd hh:mm',
};

// ───────────────────────────────────────────────────────────────────────────
// 地址与数值工具
// ───────────────────────────────────────────────────────────────────────────

/** 0 → A，25 → Z，26 → AA … */
function colLetters(index) {
    let n = Math.max(0, Math.floor(index));
    let out = '';
    do {
        out = String.fromCharCode(65 + (n % 26)) + out;
        n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return out;
}

/** 'A' → 0，'AA' → 26。非法输入返回 -1。 */
function lettersToCol(letters) {
    let n = 0;
    for (const ch of String(letters ?? '').toUpperCase()) {
        const code = ch.charCodeAt(0);
        if (code < 65 || code > 90) return -1;
        n = n * 26 + (code - 64);
    }
    return n - 1;
}

/** 'C3' → {col:2,row:2}（行列都归零）。非法地址返回 undefined。 */
function parseRef(ref) {
    const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(ref ?? '').trim());
    if (!match) return undefined;
    const col = lettersToCol(match[1]);
    if (col < 0) return undefined;
    return { col, row: Number.parseInt(match[2], 10) - 1 };
}

/** 归零行列 → 'C3'。 */
function buildRef(col, row) {
    return `${colLetters(col)}${row + 1}`;
}

/** 'A1:D9' → {minCol,minRow,maxCol,maxRow}；单个 'B2' 也接受。非法返回 undefined。 */
function parseRange(range) {
    const text = String(range ?? '').trim();
    if (text === '') return undefined;
    const [head, tail = head] = text.split(':');
    const a = parseRef(head);
    const b = parseRef(tail);
    if (!a || !b) return undefined;
    return {
        minCol: Math.min(a.col, b.col),
        maxCol: Math.max(a.col, b.col),
        minRow: Math.min(a.row, b.row),
        maxRow: Math.max(a.row, b.row),
    };
}

/**
 * 显示宽度：全角字符折算成 2 个西文字符。
 * 直接用 estimateEm 会把「中文比西文宽一倍」这件事抹平，列宽就会估窄。
 */
function visualWidth(text) {
    let wide = 0;
    let total = 0;
    for (const ch of String(text ?? '')) {
        total += 1;
        if (estimateEm(ch) >= 0.95) wide += 1;
    }
    return (total - wide) + wide * 2;
}

/** 1900 日期系统的序列号。Excel 保留了「1900 是闰年」的 Lotus 兼容 bug，只能用魔法基数换算。 */
function dateSerial(date) {
    const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
    const serial = Math.round((utc - Date.UTC(1899, 11, 30)) / 86400000);
    const fraction = (date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds()) / 86400;
    return serial + fraction;
}

/** 解析「2024-3-5」「2024/3/5」「2024年3月5日」以及可选的时间。 */
function parseDateText(text) {
    const match = /^(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})日?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
        .exec(String(text ?? '').trim());
    if (!match) return undefined;
    const date = new Date(
        Number(match[1]), Number(match[2]) - 1, Number(match[3]),
        Number(match[4] ?? 0), Number(match[5] ?? 0), Number(match[6] ?? 0),
    );
    return Number.isNaN(date.getTime()) ? undefined : date;
}

const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
const round2 = (value) => Math.round(Number(value) * 100) / 100;

function escapeText(value) {
    return String(value ?? '').replace(/[&<>"']/g, (ch) => (
        ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '"' ? '&quot;' : '&apos;'
    ));
}

function decodeEntities(value) {
    return String(value)
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
        .replace(/&#(\d+);/g, (_m, code) => String.fromCodePoint(Number(code)))
        .replace(/&amp;/g, '&');
}

// ───────────────────────────────────────────────────────────────────────────
// 样式簿：字体 / 填充 / 边框 / numFmt 去重后分配下标
// ───────────────────────────────────────────────────────────────────────────

/** 边框预设：表格好看的关键是「细、淡、只描需要的边」。 */
const BORDER_PRESETS = {
    thin: { all: { style: 'thin' } },
    medium: { all: { style: 'medium' } },
    header: { bottom: { style: 'medium' } },
    top: { top: { style: 'thin' } },
    bottom: { bottom: { style: 'thin' } },
};

const ALIGN_H = {
    left: 'left', right: 'right', center: 'center', centre: 'center',
    justify: 'justify', distributed: 'distributed', fill: 'fill',
};
const ALIGN_V = { top: 'top', center: 'center', centre: 'center', bottom: 'bottom', middle: 'center' };

/** 把 {align,valign,wrap,indent} 收敛成 xf 的去重键与渲染参数。 */
function normalizeAlign(spec) {
    if (!spec) return 'none';
    const h = ALIGN_H[asString(spec.align).trim().toLowerCase()] ?? '';
    const v = ALIGN_V[asString(spec.valign).trim().toLowerCase()] ?? '';
    const wrap = !!spec.wrap;
    const indent = Number(spec.indent) > 0 ? Math.min(250, Math.floor(Number(spec.indent))) : 0;
    if (!h && !v && !wrap && !indent) return 'none';
    return JSON.stringify({ h, v, wrap, indent });
}

/** 主题色令牌 → ARGB，让调用方能写 `color:'primary'` 而不必记十六进制。 */
function createPalette(theme) {
    const palette = {};
    for (const [key, value] of Object.entries(theme.colors ?? {})) palette[key] = toHex(value);
    palette.white = 'FFFFFF';
    palette.black = '000000';
    return palette;
}

/** 颜色值 → ARGB。6 位 hex 必须补 FF 前缀，否则 Excel 会当成透明。 */
function argb(value, palette, fallback) {
    if (value === undefined || value === null || value === '') return fallback;
    const text = String(value).trim();
    if (/^[0-9a-fA-F]{6}$/.test(text)) return `FF${text.toUpperCase()}`;
    if (/^[0-9a-fA-F]{8}$/.test(text)) return text.toUpperCase();
    if (palette[text] !== undefined) return `FF${palette[text]}`;
    return fallback;
}

class StyleBook {
    constructor(theme) {
        this.palette = createPalette(theme);
        this.colors = theme.colors ?? {};
        this.table = theme.table ?? {};
        // 三张去重表必须先于默认项建好：默认项也走同一条注册路径，顺序错了会读到 undefined
        this.fonts = [];
        this.fontIndex = new Map();
        this.fills = [];
        this.fillIndex = new Map();
        this.borders = [];
        this.borderIndex = new Map();
        // fills 的前两位被 schema 钉死：0 = none，1 = gray125（Excel 要求保留）
        this.registerFill('none');
        this.registerFill('gray125');
        this.registerBorder('');
        this.font({ name: theme.fonts?.en ?? 'Calibri', size: 11 });
        this.numFmts = new Map();   // id → code
        this.nextFmtId = CUSTOM_FMT_BASE;
        this.xfs = [];
        this.xfIndex = new Map();
    }

    font(spec) {
        const name = asString(spec.name, 'Calibri');
        const size = Number(spec.size) > 0 ? Number(spec.size) : 11;
        const key = `${spec.bold ? 'b' : ''}${spec.italic ? 'i' : ''}|${name}|${size}|${spec.color ?? ''}`;
        const hit = this.fontIndex.get(key);
        if (hit !== undefined) return hit;
        this.fonts.push({ bold: !!spec.bold, italic: !!spec.italic, name, size, color: spec.color || undefined });
        this.fontIndex.set(key, this.fonts.length - 1);
        return this.fonts.length - 1;
    }

    fill(spec) {
        if (!spec || typeof spec !== 'object' || (!spec.argb && !spec.pattern)) return 0;
        return this.registerFill(asString(spec.pattern, 'solid'), spec.argb);
    }

    registerFill(pattern, argbValue) {
        const key = `${pattern}|${argbValue ?? ''}`;
        const hit = this.fillIndex.get(key);
        if (hit !== undefined) return hit;
        this.fills.push({ pattern, argb: argbValue });
        this.fillIndex.set(key, this.fills.length - 1);
        return this.fills.length - 1;
    }

    border(spec) {
        if (spec === undefined || spec === null || spec === '') return 0;
        const preset = typeof spec === 'string' ? spec : asString(spec.preset, 'thin');
        const color = typeof spec === 'object' ? spec.color : undefined;
        return this.registerBorder(preset, color);
    }

    registerBorder(preset, color) {
        const sides = BORDER_PRESETS[String(preset)] ?? BORDER_PRESETS.thin;
        const key = `${preset}|${color ?? ''}`;
        const hit = this.borderIndex.get(key);
        if (hit !== undefined) return hit;
        const entry = {};
        for (const side of ['left', 'right', 'top', 'bottom']) {
            if (sides.all) entry[side] = { style: sides.all.style, color };
            else if (sides[side]) entry[side] = { style: sides[side].style, color };
        }
        this.borders.push(entry);
        this.borderIndex.set(key, this.borders.length - 1);
        return this.borders.length - 1;
    }

    /** 注册 numFmt 并返回 id；命中内置格式时直接返回内置编号。 */
    numFmt(code) {
        const text = asString(code).trim();
        if (text === '' || text.toLowerCase() === 'general') return 0;
        const builtin = BUILTIN_BY_CODE.get(text);
        if (builtin !== undefined) return builtin;
        for (const [id, existing] of this.numFmts) {
            if (existing === text) return id;
        }
        const id = this.nextFmtId;
        this.nextFmtId += 1;
        this.numFmts.set(id, text);
        return id;
    }

    /** 取（或新建）cellXfs 下标。单元格 `s` 属性写的就是它。 */
    xf(spec = {}) {
        const fontId = spec.fontId ?? 0;
        const fillId = spec.fillId ?? 0;
        const borderId = spec.borderId ?? 0;
        const numFmtId = spec.numFmtId ?? 0;
        const align = normalizeAlign(spec.align);
        const key = `${fontId}|${fillId}|${borderId}|${numFmtId}|${align}`;
        const hit = this.xfIndex.get(key);
        if (hit !== undefined) return hit;
        this.xfs.push({ fontId, fillId, borderId, numFmtId, align: align === 'none' ? undefined : align });
        this.xfIndex.set(key, this.xfs.length - 1);
        return this.xfs.length - 1;
    }
}

/** 判断某个 numFmt 是否属于日期/时间（读文件时判断序列号语义用）。 */
function isDateFormatCode(code) {
    if (!code) return false;
    const cleaned = String(code)
        .replace(/\[[^\]]*\]/g, '')
        .replace(/"[^"]*"/g, '')
        .replace(/\\./g, '');
    return /[ymdhs]/i.test(cleaned) && !/^[#0.,%\s]*$/.test(cleaned);
}

// ───────────────────────────────────────────────────────────────────────────
// 值编译：写盘与估宽必须看到同一份结论
// ───────────────────────────────────────────────────────────────────────────

function makeCell(col, value) {
    return {
        col,
        value: value === undefined ? '' : value,
        styleId: undefined,
        explicitFmt: '',
        formula: undefined,
        style: { align: '', valign: '', wrap: false, indent: 0 },
    };
}

/** 值 → 类型提示。显式 type 优先，否则用 kit 的猜测。 */
function detectKind(value, hint) {
    if (typeof value === 'number') return 'number';
    if (value instanceof Date) return 'date';
    if (typeof value === 'boolean') return 'bool';
    if (hint && hint !== 'auto') return hint;
    return classifyValue(value);
}

/**
 * 把原始值编译成「写入 XML 需要的一切」。
 * 日期/百分比/货币都变成数字，numFmt 留给 defaultFormat 决定。
 */
function compileValue(raw, hint) {
    if (raw === null || raw === undefined) return { kind: 'text', text: '' };
    if (raw instanceof Date) return { kind: 'date', number: dateSerial(raw) };
    if (typeof raw === 'number') {
        return Number.isFinite(raw) ? { kind: 'number', number: raw } : { kind: 'text', text: '' };
    }
    if (typeof raw === 'boolean') return { kind: 'bool', bool: raw };

    const text = asString(raw);
    if (text === '') return { kind: 'text', text: '' };
    const kind = detectKind(raw, hint);

    if (kind === 'number' || kind === 'currency') {
        const n = toNumber(text);
        if (Number.isFinite(n)) return { kind: 'number', number: n, money: kind === 'currency' };
        return { kind: 'text', text };
    }
    if (kind === 'percent') {
        // 只有「带 % 号的文本」才需要除以 100；数值型百分比按原样存
        const n = toNumber(text);
        if (Number.isFinite(n)) return { kind: 'number', number: n, percent: /%/.test(text) };
        return { kind: 'text', text };
    }
    if (kind === 'date') {
        const date = parseDateText(text);
        if (date) return { kind: 'date', number: dateSerial(date) };
        return { kind: 'text', text };
    }
    return { kind: 'text', text };
}

/** 没有显式格式时，按值类型给默认 numFmt。 */
function defaultFormat(cell) {
    if (cell.explicitFmt) return cell.explicitFmt;
    const compiled = cell.compiled;
    if (!compiled) return '';
    // 带缓存值的公式格（统计块）：公式本身没有「值」可供判型，就用缓存值判。
    // 少了这一步，统计块会退化成「常规」格式，千分位与小数位全丢
    if (cell.formula && Number.isFinite(cell.cachedValue)) {
        return Number.isInteger(cell.cachedValue) ? DEFAULT_FORMAT.int : DEFAULT_FORMAT.float;
    }
    if (compiled.kind === 'number') {
        if (compiled.money) return DEFAULT_FORMAT.currency;
        if (compiled.percent) return DEFAULT_FORMAT.percent;
        if (!Number.isFinite(compiled.number)) return '';
        // 整数用千分位、小数保留两位 —— 财务表最不容易出错的默认观感
        return Number.isInteger(compiled.number) ? DEFAULT_FORMAT.int : DEFAULT_FORMAT.float;
    }
    if (compiled.kind === 'date') {
        return Number.isInteger(compiled.number) ? DEFAULT_FORMAT.date : DEFAULT_FORMAT.datetime;
    }
    return '';
}

/** 单元格显示文本（估宽用）。数字按格式暗示的千分位/小数位还原。 */
function displayText(cell) {
    const compiled = cell.compiled;
    if (!compiled) return asString(cell.value);
    // 公式格自己没有值（compiled 是空文本）。带缓存值时就按缓存数字展示与估宽 ——
    // 统计块靠这一条才能拿到合适的列宽；没缓存值仍旧返回空串，与老行为一致
    if (cell.formula) {
        return Number.isFinite(cell.cachedValue) ? formatNumberText(cell.cachedValue, cell.explicitFmt ?? '') : '';
    }
    if (compiled.kind === 'text') return compiled.text;
    if (compiled.kind === 'bool') return compiled.bool ? 'TRUE' : 'FALSE';
    return formatNumberText(compiled.number, cell.explicitFmt ?? '');
}

/** 数字 → 按 numFmt 暗示的千分位/小数位还原的显示文本（估宽用，不求精确）。 */
function formatNumberText(n, fmt) {
    if (/0\.00/.test(fmt)) {
        return /#,##0/.test(fmt)
            ? Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
            : Number(n).toFixed(2);
    }
    if (/%/.test(fmt)) return `${round2(n * 100)}%`;
    if (/[ymd]/.test(fmt) && !/^[#0.,%\s]*$/.test(fmt)) return '2000-00-00';
    if (Number.isInteger(n)) return /#,##0/.test(fmt) ? Number(n).toLocaleString('en-US') : String(n);
    return String(round2(n));
}

/** 类型 → 默认对齐。数字右、文本左、表头居中，这是 Excel 用户的肌肉记忆。 */
function alignForType(column, value) {
    if (column.align) return column.align;
    const kind = column.type === 'auto' ? classifyValue(value) : column.type;
    if (kind === 'number' || kind === 'percent' || kind === 'currency' || kind === 'date') return 'right';
    return 'left';
}

// ───────────────────────────────────────────────────────────────────────────
// 统计：函数表 + 取数口径
// ───────────────────────────────────────────────────────────────────────────

/**
 * 支持的统计量。三条约束决定了这张表的结构：
 * 1. `formula` 必须是 Excel 里真实存在的函数名（要写进 `<f>`），不能自造；
 * 2. `calc` 是缓存值（`<v>`）的唯一来源 —— 数字由本模块从内存模型真算出来，
 *    绝不写死，也不靠阅读器替我们算；
 * 3. `min` 是「几个数才算得出来」的下限：标准差/方差只有一个数时 Excel 会给出
 *    #DIV/0!，那种格子宁可留空 + 记 warning，也不要往文件里塞一个错误值。
 */
const STAT_FUNCS = {
    count: { label: '计数', formula: 'COUNT', min: 0, calc: (nums) => nums.length },
    counta: { label: '非空计数', formula: 'COUNTA', min: 0, calc: (_nums, nonEmpty) => nonEmpty },
    sum: { label: '求和', formula: 'SUM', min: 0, calc: (nums) => nums.reduce((a, b) => a + b, 0) },
    average: { label: '平均值', formula: 'AVERAGE', min: 1, calc: (nums) => nums.reduce((a, b) => a + b, 0) / nums.length },
    median: {
        label: '中位数', formula: 'MEDIAN', min: 1,
        calc: (nums) => {
            const sorted = [...nums].sort((a, b) => a - b);
            const mid = Math.floor(sorted.length / 2);
            return sorted.length % 2 === 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
        },
    },
    // 用 reduce 而不是 Math.max(...nums)：展开大数组会压爆调用栈，区间长度不可控
    max: { label: '最大值', formula: 'MAX', min: 1, calc: (nums) => nums.reduce((a, b) => (b > a ? b : a), nums[0]) },
    min: { label: '最小值', formula: 'MIN', min: 1, calc: (nums) => nums.reduce((a, b) => (b < a ? b : a), nums[0]) },
    stdev: { label: '标准差', formula: 'STDEV', min: 2, calc: (nums) => Math.sqrt(sampleVariance(nums)) },
    var: { label: '方差', formula: 'VAR', min: 2, calc: (nums) => sampleVariance(nums) },
    product: { label: '乘积', formula: 'PRODUCT', min: 1, calc: (nums) => nums.reduce((a, b) => a * b, 1) },
};

/** 样本方差：Excel 的 VAR / STDEV 都是 n−1 分母（总体版本 VARP/STDEVP 不在支持范围内）。 */
function sampleVariance(nums) {
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return nums.reduce((sum, n) => sum + (n - mean) ** 2, 0) / (nums.length - 1);
}

/** summary() 只支持有条件版本的老函数。 */
const SUMMARY_FUNCS = new Set(['sum', 'count', 'average']);

/** 统计项别名 → 规范 id。中英文都收，模型少查一次文档。 */
const STAT_ALIASES = new Map([
    ['count', 'count'], ['计数', 'count'],
    ['counta', 'counta'], ['非空计数', 'counta'],
    ['sum', 'sum'], ['total', 'sum'], ['求和', 'sum'],
    ['average', 'average'], ['avg', 'average'], ['mean', 'average'], ['平均值', 'average'], ['均值', 'average'],
    ['median', 'median'], ['中位数', 'median'],
    ['max', 'max'], ['maximum', 'max'], ['最大值', 'max'],
    ['min', 'min'], ['minimum', 'min'], ['最小值', 'min'],
    ['stdev', 'stdev'], ['stddev', 'stdev'], ['std', 'stdev'], ['标准差', 'stdev'],
    ['var', 'var'], ['variance', 'var'], ['方差', 'var'],
    ['product', 'product'], ['乘积', 'product'],
]);

/** 统计项 id（含别名）→ 规范 id；不认识返回 undefined。 */
function statFuncOf(id) {
    const key = asString(id).trim().toLowerCase();
    return STAT_ALIASES.get(key);
}

/** 公式缓存值：抹掉浮点尾数噪声（0.1+0.2 → 0.3），整数与大数保持原样。 */
function cleanNumber(value) {
    if (!Number.isFinite(value)) return undefined;
    if (Number.isInteger(value)) return value;
    const rounded = Number(value.toPrecision(12));
    return Number.isFinite(rounded) ? rounded : value;
}

/**
 * 单元格是否为空（COUNTA 口径）。
 * 公式格一律算非空：Excel 的 COUNTA 也把公式格计入，哪怕它算出来是空串。
 */
function cellIsEmpty(cell) {
    if (!cell) return true;
    if (cell.formula) return false;
    const compiled = cell.compiled ?? compileValue(cell.value, cell.typeHint ?? 'auto');
    if (compiled.kind === 'text') return compiled.text === '';
    if (compiled.kind === 'number' || compiled.kind === 'date') return !Number.isFinite(compiled.number);
    return false;
}

/**
 * 单元格的数值。数字与日期算数；文本、布尔、空、公式格都不算 ——
 * 这正是 Excel 里 COUNT / SUM 对区域参数的口径（区域里的文本数字不会被自动转换）。
 */
function numericValueOf(cell) {
    if (!cell || cell.formula) return Number.NaN;
    const compiled = cell.compiled ?? compileValue(cell.value, cell.typeHint ?? 'auto');
    if (compiled.kind !== 'number' && compiled.kind !== 'date') return Number.NaN;
    return Number.isFinite(compiled.number) ? compiled.number : Number.NaN;
}

/** 区间 → 'B5:B8'；单格就是 'B5'。 */
function rangeTextOf(range) {
    if (range.minCol === range.maxCol && range.minRow === range.maxRow) return buildRef(range.minCol, range.minRow);
    return `${buildRef(range.minCol, range.minRow)}:${buildRef(range.maxCol, range.maxRow)}`;
}

/** SUMIF/COUNTIF/AVERAGEIF 的判据：数字不带引号，文本加引号并把内部引号翻倍。 */
function criteriaOf(cell, text) {
    if (cell) {
        const compiled = cell.compiled ?? compileValue(cell.value, cell.typeHint ?? 'auto');
        if ((compiled.kind === 'number' || compiled.kind === 'date') && Number.isFinite(compiled.number)) {
            return String(compiled.number);
        }
    }
    return `"${String(text).replace(/"/g, '""')}"`;
}

/**
 * 分组的条件公式。
 * 注意 COUNTIF 只有两个参数（区间 + 判据）：写成 COUNTIF(key, criteria, value)
 * 是语法错误，Excel 打开时会弹「发现不可读取的内容」并要求修复。
 */
function conditionalFormula(id, keyRange, criteria, valueRange) {
    if (id === 'count') return `COUNTIF(${keyRange},${criteria})`;
    return `${STAT_FUNCS[id].formula}IF(${keyRange},${criteria},${valueRange})`;
}

/**
 * 宽松取整数：数字与数字字符串都收，非法值返回 undefined。
 * 单独写一个而不是直接 Number(x)，是因为「非法输入只记 warning、绝不抛异常」
 * 是这套 API 的硬约束 —— Number(Symbol()) 之类会直接抛。
 */
function looseInteger(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? Math.floor(value) : undefined;
    if (value === undefined || value === null || value === '') return undefined;
    const n = Number.parseFloat(String(value));
    return Number.isFinite(n) ? Math.floor(n) : undefined;
}

/** 非法输入要进 warning，得先把原始值变成一句能读的文字。 */
function describeSource(source) {
    if (source === undefined || source === null) return '';
    if (typeof source === 'object') {
        try {
            return JSON.stringify(source);
        } catch {
            return '[object]';
        }
    }
    return String(source);
}

// ───────────────────────────────────────────────────────────────────────────
// Sheet：一个「虚拟落笔 + 样式 + 布局」的表格
// ───────────────────────────────────────────────────────────────────────────

class Sheet {
    constructor(workbook, name, opts = {}) {
        this.wb = workbook;
        this.styles = workbook.styles;
        this.name = name;
        this.opts = opts;
        this.rows = [];
        this.cursor = 0;                      // 下一个可用行（0 基）
        // 显式列宽不能叫 widths：那会盖住同名的 builder 方法
        this.colWidths = new Map();
        this.merges = [];
        this.freezeAt = undefined;
        this.filterRange = undefined;
        this.widthHints = new Map();          // 列 → 人工宽度
        this.formulas = [];
        this.layout = {
            titleRange: undefined, headerRow: undefined,
            dataFrom: undefined, dataTo: undefined, totalRow: undefined,
            tableFrom: undefined, tableTo: undefined, columns: 0,
            // 统计块/汇总块的位置记在这里：outline 与布局标记都要靠它，
            // 报告里才能出现「统计块」这几个字，而不是让模型去猜那几行是什么
            statsBlocks: [], summaryBlocks: [],
        };
        this.columnCount = 0;
        this.headerCount = 0;
        // 这张表上的原生图表（历史遗留 11-6）：每项 { spec, at }，序列化时变成
        // xl/charts/chartN.xml + xl/drawings/drawingN.xml + 两个 rels。
        this.charts = [];
    }

    rowIndex(index) {
        while (this.rows.length <= index) this.rows.push([]);
        return this.rows[index];
    }

    getCell(ref) {
        const at = parseRef(ref);
        if (!at) return undefined;
        return this.rows[at.row]?.[at.col];
    }

    /** 落笔一个值（或空格），样式与格式先记在单元格上，渲染前才分配 xf 下标。 */
    place(ref, value, opts = {}) {
        const at = parseRef(ref);
        if (!at) {
            this.wb.warn(`excel: 单元格地址「${String(ref)}」不合法，已忽略`);
            return undefined;
        }
        const cells = this.rowIndex(at.row);
        const cell = cells[at.col] ?? makeCell(at.col, '');
        cells[at.col] = cell;
        cell.value = value === undefined ? '' : value;
        cell.compiled = undefined;
        this.applyStyle(cell, opts);
        if (opts.format) cell.explicitFmt = asString(opts.format);
        if (opts.width) this.widthHints.set(at.col, Math.max(this.widthHints.get(at.col) ?? 0, Number(opts.width) || 0));
        this.cursor = Math.max(this.cursor, at.row + 1);
        this.columnCount = Math.max(this.columnCount, at.col + 1);
        return cell;
    }

    applyStyle(cell, opts) {
        cell.opts = { ...(cell.opts ?? {}), ...opts };
        cell.style = {
            align: asString(opts.align, cell.style?.align ?? ''),
            valign: asString(opts.valign, cell.style?.valign ?? ''),
            wrap: opts.wrap !== undefined ? !!opts.wrap : !!cell.style?.wrap,
            indent: Number(opts.indent) > 0 ? Math.min(250, Math.floor(Number(opts.indent))) : (cell.style?.indent ?? 0),
        };
        cell.styleId = undefined; // 样式变了，缓存的 xf 下标作废
    }

    // ── builder 方法 ──────────────────────────────────────────────────────

    /** 大标题行：合并 + 加粗。底色由主题给（素色主题为空 = 不填充）。 */
    title(text, opts = {}) {
        const row = Number.isFinite(opts.row) ? Math.floor(opts.row) : this.cursor;
        const span = Math.max(1, Math.floor(Number(opts.span) || Math.max(this.columnCount, 1)));
        const titleFill = shadeOf(this.styles.table.titleFill);
        const cells = this.rowIndex(row);
        cells.length = Math.max(cells.length, span);
        for (let col = 0; col < span; col += 1) {
            const cell = cells[col] ?? makeCell(col, '');
            cells[col] = cell;
            // 只有首格承载文字与加粗，其余格只留底色，免得合并后出现意外加粗
            cell.value = col === 0 ? text : '';
            cell.compiled = undefined;
            cell.opts = {
                ...(cell.opts ?? {}), ...opts,
                bold: true, size: Number(opts.size) || 16,
                fill: opts.fill ?? titleFill, color: opts.color ?? 'primaryDark',
                align: 'left', valign: 'center', border: opts.border,
            };
            cell.style = { align: 'left', valign: 'center', wrap: false, indent: 0 };
            cell.styleId = undefined;
            cell.isTitle = col === 0;
        }
        if (span > 1) this.mergeRange(`${buildRef(0, row)}:${buildRef(span - 1, row)}`);
        this.rowHeight(row, Number(opts.height) || 30);
        this.layout.titleRange = { row, span };
        this.columnCount = Math.max(this.columnCount, span);
        this.cursor = Math.max(this.cursor, row + 1);
        return this;
    }

    /** 灰字说明行：只在 A 列，不参与列宽估算。 */
    note(text, opts = {}) {
        const row = Number.isFinite(opts.row) ? Math.floor(opts.row) : this.cursor;
        const cells = this.rowIndex(row);
        const cell = cells[0] ?? makeCell(0, '');
        cells[0] = cell;
        cell.value = text;
        cell.compiled = undefined;
        cell.opts = { ...(cell.opts ?? {}), color: opts.color ?? 'muted', size: Number(opts.size) || 9, italic: opts.italic !== false };
        cell.style = { align: 'left', valign: 'center', wrap: false, indent: 0 };
        cell.styleId = undefined;
        cell.isNote = true;
        this.rowHeight(row, 16);
        this.cursor = Math.max(this.cursor, row + 1);
        return this;
    }

    /** 只有表头没有数据。 */
    header(columns, opts = {}) {
        return this.table({ columns, rows: [], headerOnly: true, ...opts });
    }

    /**
     * 主数据表：表头 + 数据行 + 合计行 + 冻结 + 自动筛选。
     * 底色全部来自主题：素色主题（默认）四个 Fill 都是空串，于是只剩边框与加粗，
     * 也就是 Excel 里普通表格的网格形态。totalRow 为 true 时对数字列生成 SUM 公式。
     */
    table(spec = {}) {
        const columns = normalizeColumns(spec.columns ?? []);
        if (columns.length === 0) {
            this.wb.warn(`excel: 表「${this.name}」的 table() 没有列定义，已跳过`);
            return this;
        }
        const rows = normalizeRows(spec.rows ?? [], columns.length);
        const startRow = Number.isFinite(spec.startRow) ? Math.floor(spec.startRow) : this.cursor;
        const t = this.styles.table;
        // 斑马纹默认跟随主题：主题没给 zebraFill（素色）就不画；显式 zebra: true 仍照办
        const zebra = spec.zebra === undefined ? shadeOf(t.zebraFill) !== '' : spec.zebra === true;
        const totalRow = spec.totalRow === true ? {}
            : (spec.totalRow && typeof spec.totalRow === 'object' ? spec.totalRow : undefined);
        const border = spec.border ?? 'thin';

        // ── 表头：主题给了主色底就用主色底 + 白字，素色主题只有加粗与边框
        for (let col = 0; col < columns.length; col += 1) {
            const column = columns[col];
            const cell = this.place(buildRef(col, startRow), column.title, {
                bold: true,
                size: 11,
                color: t.headerText,
                fill: shadeOf(t.headerFill),
                // 表头一律居中：同一张表里文本列表头左、数字列表头中会显得没对齐；
                // 数据行才按类型分左右（见 alignForType）
                align: column.align || 'center',
                valign: 'center',
                border,
                width: column.width,
            });
            if (cell) {
                cell.isHeader = true;
                if (column.format) cell.explicitFmt = column.format;
            }
        }
        this.rowHeight(startRow, Number(spec.headerHeight) || 26);

        // ── 数据行：主题有斑马纹底色时交替填充，数字右对齐
        for (let r = 0; r < rows.length; r += 1) {
            const row = startRow + 1 + r;
            const stripe = zebra && r % 2 === 1;
            for (let col = 0; col < columns.length; col += 1) {
                const column = columns[col];
                const cell = this.place(buildRef(col, row), rows[r][col], {
                    fill: stripe ? shadeOf(t.zebraFill) : undefined,
                    align: column.align || alignForType(column, rows[r][col]),
                    valign: 'center',
                    border,
                    width: column.width,
                });
                if (cell) {
                    cell.isData = true;
                    cell.typeHint = column.type;
                    if (column.format) cell.explicitFmt = column.format;
                }
            }
            this.rowHeight(row, Number(spec.height) || 19);
        }

        // ── 合计行：主题有强调色就上底色，素色主题只加粗 + SUM 公式
        const dataFrom = startRow + 1;
        const dataTo = startRow + rows.length;
        if (totalRow) {
            const at = dataTo + 1;
            const sumColumns = totalRow.columns
                ? new Set(asArray(totalRow.columns).map((c) => (typeof c === 'number' ? c : lettersToCol(String(c)))))
                : new Set(columns.map((c, i) => (['number', 'currency', 'percent'].includes(c.type) ? i : -1)).filter((i) => i >= 0));
            for (let col = 0; col < columns.length; col += 1) {
                const ref = buildRef(col, at);
                const base = {
                    bold: true, fill: shadeOf(t.accentFill), color: 'primaryDark',
                    valign: 'center', border, width: columns[col].width,
                };
                if (col === 0) {
                    const cell = this.place(ref, asString(totalRow.label, '合计'), { ...base, align: 'left' });
                    if (cell) cell.isTotal = true;
                } else if (sumColumns.has(col) && rows.length > 0) {
                    const cell = this.place(ref, '', { ...base, align: 'right' });
                    if (cell) {
                        cell.isTotal = true;
                        cell.typeHint = columns[col].type;
                        this.setFormula(cell, `SUM(${buildRef(col, dataFrom)}:${buildRef(col, dataTo)})`);
                        if (columns[col].format) cell.explicitFmt = columns[col].format;
                    }
                } else {
                    const values = asArray(totalRow.values);
                    const cell = this.place(ref, values[col] ?? '', { ...base, align: columns[col].align || 'center' });
                    if (cell) cell.isTotal = true;
                }
            }
            this.rowHeight(at, 22);
            this.layout.totalRow = at;
        }

        this.layout.headerRow = startRow;
        this.layout.tableFrom = startRow;
        this.layout.tableTo = totalRow ? dataTo + 1 : dataTo;
        // 记下这张表有几列：同一行里表头右边可能还有别的文字（例如说明格），
        // read() 按列名找列时必须只取前 N 格，否则会混进不相干的文字
        this.layout.columns = columns.length;
        if (rows.length > 0) {
            this.layout.dataFrom = dataFrom;
            this.layout.dataTo = dataTo;
        }
        this.columnCount = Math.max(this.columnCount, columns.length);
        this.headerCount = Math.max(this.headerCount, columns.length);
        this.cursor = Math.max(this.cursor, totalRow ? dataTo + 2 : startRow + rows.length + 1);

        if (spec.freeze === true || spec.freeze === 'header') {
            this.freeze(startRow + 1, Number(spec.freezeCol) || 0);
        } else if (Array.isArray(spec.freeze)) {
            this.freeze(spec.freeze[0], spec.freeze[1]);
        } else if (spec.freeze && typeof spec.freeze === 'object') {
            this.freeze(spec.freeze.row, spec.freeze.col);
        }
        if (spec.autofilter === true && rows.length > 0) {
            this.autofilter(`${buildRef(0, startRow)}:${buildRef(columns.length - 1, totalRow ? dataTo + 1 : dataTo)}`);
        }
        return this;
    }

    /** 追加单行。`at` 缺省时接着游标写。 */
    row(values, opts = {}) {
        const list = asArray(values);
        const at = Number.isFinite(opts.at) ? Math.floor(opts.at) : this.cursor;
        for (let col = 0; col < list.length; col += 1) {
            const cell = this.place(buildRef(col, at), list[col], opts);
            if (cell && opts.format) cell.explicitFmt = opts.format;
        }
        if (opts.height) this.rowHeight(at, opts.height);
        this.cursor = Math.max(this.cursor, at + 1);
        return this;
    }

    /** 指定地址写一格。 */
    cell(ref, value, opts = {}) {
        const cell = this.place(ref, value, opts);
        if (cell && opts.format) cell.explicitFmt = opts.format;
        return this;
    }

    /** 公式：`formula` 不含前导 `=`。空公式串直接忽略，免得生成 #NAME?。 */
    formula(ref, formula) {
        const text = asString(formula).trim().replace(/^=/, '');
        const at = parseRef(ref);
        if (text === '' || !at) {
            this.wb.warn(`excel: ${this.name}!${String(ref)} 的公式${text === '' ? '为空' : '地址不合法'}，已忽略`);
            return this;
        }
        const cells = this.rowIndex(at.row);
        const cell = cells[at.col] ?? makeCell(at.col, '');
        cells[at.col] = cell;
        cell.styleId = undefined;
        this.setFormula(cell, text);
        this.cursor = Math.max(this.cursor, at.row + 1);
        this.columnCount = Math.max(this.columnCount, at.col + 1);
        return this;
    }

    setFormula(cell, formula, cachedValue) {
        const text = asString(formula).trim().replace(/^=/, '');
        if (text === '') return cell;
        cell.formula = text;
        // 缓存值只在调用方真的算出来时才带上（统计块走这条路）；手写表达式引擎算不了，
        // 传 undefined 就等于「没有缓存值」，照旧只写 <f>，行为与老版本一致
        cell.cachedValue = Number.isFinite(cachedValue) ? cleanNumber(cachedValue) : undefined;
        cell.value = '';
        cell.compiled = undefined;
        cell.styleId = undefined;
        const row = this.rows.findIndex((cells) => cells && cells.includes(cell));
        if (row >= 0) this.formulas.push({ ref: `${this.name}!${buildRef(cell.col, row)}`, formula: text });
        return cell;
    }

    // ── 统计块与分组汇总 ──────────────────────────────────────────────────

    /**
     * 统计块：`sheet.stats('B5:B8', {at, label, funcs, layout, format, header})`。
     *
     * 每个数值格同时写 `<f>` 真公式（COUNT/SUM/AVERAGE…）与 `<v>` 缓存值（本模块
     * 从内存模型真算出来的）。为什么两样都要：公式格在「不重算的阅读器」里是一片
     * 空白，而统计数字本身就是结论，看不见等于没写；保留 `<f>` 则让 Excel 里能点开
     * 看算法、源数据改了也跟着变（workbook.xml 的 fullCalcOnLoad 保证打开时重算）。
     *
     * 源区间一个数字都没有时，只写标题与项名，统计格留空并记一条 warning ——
     * 写 0 会让人误以为「真的是 0」，写错误值更是坏文件。
     */
    stats(source, opts = {}) {
        // 传了 null / 字符串这类「非对象」也要能跑：默认参数只接住 undefined
        const options = opts && typeof opts === 'object' ? opts : {};
        const at = this.blockRow(options.at);
        const parsed = this.statsSource(source);
        if (!parsed) {
            this.wb.warn(`excel: stats() 的源区间「${describeSource(source)}」不合法，已忽略`);
            return this;
        }
        const funcs = this.pickFuncs(options.funcs, 'stats', ['count', 'sum', 'average', 'max', 'min']);
        if (funcs.length === 0) return this;
        const layout = this.blockLayout(options.layout);
        const label = asString(options.label, '统计') || '统计';
        const scan = this.scanRange(parsed);
        const sourceText = rangeTextOf(parsed);
        if (scan.nums.length === 0) {
            this.wb.warn(`excel: 统计块源区间「${sourceText}」里没有数字，统计格已留空`);
        }
        // 每个统计量的缓存值都从同一份扫描结果算出来，公式里引用的也是同一个区间：
        // 缓存值与公式结果对不上是最难查的错，这里保证两者同源
        const funcValue = (item) => {
            if (scan.nums.length === 0) return undefined;
            if (scan.nums.length < item.fn.min) {
                this.wb.warn(`excel: 统计项「${item.label}」需要至少 ${item.fn.min} 个数字，`
                    + `${this.name}!${sourceText} 只有 ${scan.nums.length} 个，已留空`);
                return undefined;
            }
            const value = item.fn.calc(scan.nums, scan.nonEmpty);
            return Number.isFinite(value) ? value : undefined;
        };

        // 标题只加粗 + 主题强调色底：素色主题的 accentFill 是空串，shadeOf 后不写填充
        const head = { bold: true, fill: shadeOf(this.styles.table.accentFill), color: 'primaryDark', valign: 'center' };
        let lastRow = at;
        let cols = 2;
        if (layout === 'columns') {
            // 横排：表头行（标题 + 各统计项名）/ 数值行。表头可以关掉，只留数值行
            const header = options.header !== false;
            if (header) {
                this.place(buildRef(0, at), label, head);
                funcs.forEach((item, index) => this.place(buildRef(1 + index, at), item.label, { ...head, align: 'center' }));
                lastRow = at;
            }
            const row = header ? at + 1 : at;
            // 首列写源区间：横排时它就是这一块统计的行标题，也让读者知道数字来自哪里
            this.place(buildRef(0, row), sourceText, { color: 'muted', size: 9, align: 'left', valign: 'center' });
            funcs.forEach((item, index) => {
                const value = funcValue(item);
                if (value !== undefined) {
                    this.statCell(buildRef(1 + index, row), `${item.fn.formula}(${sourceText})`, value, item, options);
                }
            });
            lastRow = row;
            cols = 1 + funcs.length;
        } else {
            // 竖排（默认）：左列统计项名、右列数值
            this.place(buildRef(0, at), label, head);
            funcs.forEach((item, index) => {
                const row = at + 1 + index;
                this.place(buildRef(0, row), item.label, { align: 'left', valign: 'center' });
                const value = funcValue(item);
                if (value !== undefined) {
                    this.statCell(buildRef(1, row), `${item.fn.formula}(${sourceText})`, value, item, options);
                }
                lastRow = row;
            });
        }
        this.layout.statsBlocks.push({ at, rows: lastRow - at + 1, cols, source: sourceText, funcs: funcs.length });
        this.cursor = Math.max(this.cursor, lastRow + 1);
        this.columnCount = Math.max(this.columnCount, cols);
        return this;
    }

    /**
     * 分组汇总（轻量透视）：按 key 的**首次出现顺序**输出唯一值，每行一个
     * SUMIF / COUNTIF / AVERAGEIF 公式 + 真算出来的缓存值，可选补一行合计。
     *
     * 只用这三个有条件版本的老函数：MAXIFS/MINIFS 这类新函数在文件里必须写成
     * `_xlfn.MAXIFS`，前缀写错 Excel 直接显示 #NAME?，代价远大于收益。
     */
    summary(opts = {}) {
        const options = opts && typeof opts === 'object' ? opts : {};
        const at = this.blockRow(options.at);
        const { dataFrom, dataTo } = this.layout;
        if (dataFrom === undefined || dataTo === undefined) {
            this.wb.warn(`excel: 表「${this.name}」还没有 table() 数据区，summary() 已跳过`);
            return this;
        }
        const keyCol = this.resolveColumn(options.key);
        const valueCol = options.value === undefined ? keyCol : this.resolveColumn(options.value);
        if (keyCol === undefined || valueCol === undefined) {
            const bad = keyCol === undefined ? `key「${describeSource(options.key)}」` : `value「${describeSource(options.value)}」`;
            this.wb.warn(`excel: summary() 找不到 ${bad} 对应的列（可用列字母、列号或表头标题），已忽略`);
            return this;
        }
        const funcs = this.pickFuncs(options.funcs, 'summary', ['sum', 'count', 'average']);
        if (funcs.length === 0) return this;

        // 分组：Map 保持插入顺序，正好就是「首次出现顺序」
        const order = [];
        const groups = new Map();
        let skipped = 0;
        for (let row = dataFrom; row <= dataTo; row += 1) {
            const keyCell = this.rows[row]?.[keyCol];
            const keyText = this.displayOf(keyCell).trim();
            if (keyText === '') { skipped += 1; continue; }
            let group = groups.get(keyText);
            if (!group) {
                group = { key: keyText, criteria: criteriaOf(keyCell, keyText), nums: [], nonEmpty: 0 };
                groups.set(keyText, group);
                order.push(group);
            }
            const valueCell = this.rows[row]?.[valueCol];
            if (!cellIsEmpty(valueCell)) group.nonEmpty += 1;
            const n = numericValueOf(valueCell);
            if (Number.isFinite(n)) group.nums.push(n);
        }
        if (order.length === 0) {
            this.wb.warn(`excel: summary() 在 ${this.name} 的数据区里没有非空 key，已跳过`);
            return this;
        }
        if (skipped > 0) {
            // 空 key 不分组：SUMIF 的判据写成 "" 时各阅读器对空格的匹配行为不完全一致，
            // 与其赌它，不如跳过并在报告里说清「合计行按整个数值区间计算」
            this.wb.warn(`excel: summary() 跳过了 ${skipped} 行 key 为空的记录（不计入分组；合计行按整个数值区间计算）`);
        }
        const keyRange = `${buildRef(keyCol, dataFrom)}:${buildRef(keyCol, dataTo)}`;
        const valueRange = `${buildRef(valueCol, dataFrom)}:${buildRef(valueCol, dataTo)}`;
        const totalScan = this.scanRange({ minCol: valueCol, maxCol: valueCol, minRow: dataFrom, maxRow: dataTo });
        const head = { bold: true, fill: shadeOf(this.styles.table.accentFill), color: 'primaryDark', valign: 'center' };
        const label = asString(options.label, '汇总') || '汇总';
        this.place(buildRef(0, at), label, head);
        funcs.forEach((item, index) => this.place(buildRef(1 + index, at), item.label, { ...head, align: 'center' }));

        let lastRow = at;
        order.forEach((group, index) => {
            const row = at + 1 + index;
            this.place(buildRef(0, row), group.key, { align: 'left', valign: 'center' });
            funcs.forEach((item, fi) => {
                // 该组在数值列上一个数都没有（例如整组都是文字）时留空：
                // AVERAGEIF 会算成 #DIV/0!，SUMIF/COUNTIF 会算成 0，两种都跟「没有数据」打架
                if (group.nums.length < item.fn.min) return;
                const value = item.fn.calc(group.nums, group.nonEmpty);
                if (!Number.isFinite(value)) return;
                this.statCell(buildRef(1 + fi, row), conditionalFormula(item.id, keyRange, group.criteria, valueRange), value, item, options);
            });
            lastRow = row;
        });
        if (options.total === true) {
            const row = lastRow + 1;
            this.place(buildRef(0, row), '合计', { bold: true, align: 'left', valign: 'center' });
            funcs.forEach((item, fi) => {
                if (totalScan.nums.length < item.fn.min) return;
                const value = item.fn.calc(totalScan.nums, totalScan.nonEmpty);
                if (!Number.isFinite(value)) return;
                this.statCell(buildRef(1 + fi, row), `${item.fn.formula}(${valueRange})`, value, item, options);
            });
            lastRow = row;
        }
        this.layout.summaryBlocks.push({
            at, rows: lastRow - at + 1, cols: 1 + funcs.length,
            groups: order.length, key: keyCol, value: valueCol,
        });
        this.cursor = Math.max(this.cursor, lastRow + 1);
        this.columnCount = Math.max(this.columnCount, 1 + funcs.length);
        return this;
    }

    /** 块的起始行：缺省接在游标后面；非法值记 warning 后退回游标（非法输入不抛异常）。 */
    blockRow(value) {
        if (value === undefined || value === null || value === '') return this.cursor;
        const n = looseInteger(value);
        if (n === undefined || n < 0 || n > 1048575) {
            this.wb.warn(`excel: 块的起始行「${describeSource(value)}」不合法，已改用当前游标（第 ${this.cursor + 1} 行）`);
            return this.cursor;
        }
        return n;
    }

    /** 统计块的排布：rows（两列竖排）/ columns（表头行 + 数值行横排）。非法值回落 rows。 */
    blockLayout(value) {
        const text = asString(value, 'rows').trim().toLowerCase();
        if (text === '' || text === 'rows') return 'rows';
        if (text === 'columns' || text === 'cols') return 'columns';
        this.wb.warn(`excel: 统计块的 layout「${String(value)}」无法识别（可用 rows / columns），已按 rows 处理`);
        return 'rows';
    }

    /**
     * 归一化统计项。元素可以是 'average'，也可以是 {id, label, format}。
     * 未知 id 记 warning 后跳过，不抛异常 —— 调用方少写对一个函数名，
     * 不该让整份文档失败。
     */
    pickFuncs(list, scope, fallback) {
        const wanted = asArray(list);
        const source = wanted.length > 0 ? wanted : fallback;
        const out = [];
        for (const raw of source) {
            const spec = typeof raw === 'string' ? { id: raw } : (raw ?? {});
            const id = statFuncOf(spec.id);
            if (!id) {
                this.wb.warn(`excel: ${scope}() 里的统计项「${describeSource(spec.id)}」不认识，已跳过`);
                continue;
            }
            if (scope === 'summary' && !SUMMARY_FUNCS.has(id)) {
                this.wb.warn(`excel: summary() 不支持统计项「${id}」（条件版只有 SUMIF/COUNTIF/AVERAGEIF），已跳过`);
                continue;
            }
            out.push({
                id,
                fn: STAT_FUNCS[id],
                label: asString(spec.label, '') || STAT_FUNCS[id].label,
                format: asString(spec.format, ''),
            });
        }
        return out;
    }

    /** 统计块的数值格：公式 + 缓存值 + 数字格式一次写齐（顺序不能反，setFormula 会重置值）。 */
    statCell(ref, formula, value, item, opts) {
        const cell = this.place(ref, '', { align: 'right', valign: 'center' });
        if (!cell) return undefined;
        this.setFormula(cell, formula, value);
        // 格式优先级：元素级 format > 块级 format > 按缓存值是不是整数自适应。
        // 自适应的意义：计数/求和给 #,##0，平均值带小数就给 #,##0.00，不必调用方逐个指定
        cell.explicitFmt = item.format || asString(opts.format) || (Number.isInteger(value) ? DEFAULT_FORMAT.int : DEFAULT_FORMAT.float);
        return cell;
    }

    /**
     * 把「列字母 / 列号 / 列标题」统一成 0 基列号。
     * 列号按 **1 基** 解释：{column: 2} 与 'B' 是同一个意思（API 文档写明的等价关系）。
     */
    resolveColumn(spec) {
        const asNum = (n) => (Number.isFinite(n) && n >= 1 && n <= 16384 ? Math.floor(n) - 1 : undefined);
        if (typeof spec === 'number') return asNum(spec);
        const text = asString(spec).trim();
        if (text === '') return undefined;
        if (/^[A-Za-z]{1,3}$/.test(text)) {
            const col = lettersToCol(text);
            // 字母合法不代表列号合法：'ZZZ' 能算出 18278，但 Excel 最后一列是 XFD(16384)，
            // 越界的列号写进公式只会得到 #REF!，不如在这里就判为「找不到」
            return col >= 0 && col <= 16383 ? col : undefined;
        }
        if (/^\d{1,5}$/.test(text)) return asNum(Number(text));
        // 剩下的当列标题：只在最近一次 table() 的表头行里找（那才是列名的权威来源，
        // 标题行与说明行里的文字不该被当成列名）
        const headerRow = this.layout.headerRow;
        if (headerRow === undefined) return undefined;
        const cells = this.rows[headerRow] ?? [];
        const limit = this.layout.columns > 0 ? this.layout.columns : cells.length;
        for (let col = 0; col < limit; col += 1) {
            if (asString(cells[col]?.value).trim() === text) return col;
        }
        return undefined;
    }

    /**
     * 源区间的几种写法 → {minCol, maxCol, minRow, maxRow}：
     *   stats('B5:B8')                       / stats({range: 'B5:B8'})
     *   stats({column: 'B'|2, from: 5, to: 8})   // 列号与行号都是 1 基
     *   stats('B') / stats({column: 'B'})         // 行范围缺省取最近一次 table() 的数据区
     */
    statsSource(source) {
        if (typeof source === 'string' || typeof source === 'number') {
            const text = asString(source).trim();
            if (text === '') return undefined;
            // 只给列（字母或列号）时补上数据区行范围，省掉调用方一次查坐标
            if (/^[A-Za-z]{1,3}$/.test(text) || /^\d{1,4}$/.test(text)) {
                const col = this.resolveColumn(text);
                if (col !== undefined) return this.columnSpan(col);
            }
            return parseRange(text);
        }
        if (!source || typeof source !== 'object') return undefined;
        if (source.range !== undefined) return parseRange(source.range);
        const col = this.resolveColumn(source.column);
        if (col === undefined) return undefined;
        if (source.from !== undefined || source.to !== undefined) {
            const from = looseInteger(source.from);
            const to = source.to === undefined ? from : looseInteger(source.to);
            if (from === undefined || to === undefined) return undefined;
            // 行号是 1 基（{from: 5} 就是第 5 行），且允许写反
            const minRow = Math.max(0, Math.min(from, to) - 1);
            const maxRow = Math.max(0, Math.max(from, to) - 1);
            return { minCol: col, maxCol: col, minRow, maxRow };
        }
        return this.columnSpan(col);
    }

    /** 一列 + 最近一次 table() 的数据行 → 区间；没调用过 table() 时返回 undefined。 */
    columnSpan(col) {
        const { dataFrom, dataTo } = this.layout;
        if (dataFrom === undefined || dataTo === undefined) return undefined;
        return { minCol: col, maxCol: col, minRow: dataFrom, maxRow: dataTo };
    }

    /** 扫一遍区间，取出数字与非空格数：所有统计量的唯一数据来源。 */
    scanRange(range) {
        const nums = [];
        let nonEmpty = 0;
        for (let row = range.minRow; row <= range.maxRow; row += 1) {
            for (let col = range.minCol; col <= range.maxCol; col += 1) {
                const cell = this.rows[row]?.[col];
                if (cellIsEmpty(cell)) continue;
                nonEmpty += 1;
                const n = numericValueOf(cell);
                if (Number.isFinite(n)) nums.push(n);
            }
        }
        return { nums, nonEmpty };
    }

    /** 一格的显示文本（分组 key、估宽、告警同源，避免三处各算一遍）。 */
    displayOf(cell) {
        if (!cell) return '';
        if (!cell.compiled) cell.compiled = compileValue(cell.value, cell.opts?.type ?? cell.typeHint ?? 'auto');
        return displayText(cell);
    }

    merge(range) {
        const parsed = parseRange(range);
        if (!parsed) {
            this.wb.warn(`excel: 合并区域「${String(range)}」不合法，已忽略`);
            return this;
        }
        if (parsed.minRow === parsed.maxRow && parsed.minCol === parsed.maxCol) return this;
        return this.mergeRange(`${buildRef(parsed.minCol, parsed.minRow)}:${buildRef(parsed.maxCol, parsed.maxRow)}`);
    }

    mergeRange(range) {
        if (!this.merges.includes(range)) this.merges.push(range);
        return this;
    }

    widths(map) {
        for (const [key, value] of Object.entries(map ?? {})) {
            const col = lettersToCol(key);
            const width = Number(value);
            if (col < 0 || !Number.isFinite(width) || width <= 0) {
                this.wb.warn(`excel: 列宽「${key}: ${String(value)}」不合法，已忽略`);
                continue;
            }
            this.colWidths.set(col, clamp(width, 2, 255));
        }
        return this;
    }

    /** 冻结窗格。(0,0) 等于不冻结，直接忽略。 */
    freeze(row, col = 0) {
        const r = Number.isFinite(Number(row)) ? clamp(Math.floor(Number(row)), 0, 1048575) : 0;
        const c = Number.isFinite(Number(col)) ? clamp(Math.floor(Number(col)), 0, 16383) : 0;
        if (r === 0 && c === 0) return this;
        this.freezeAt = { row: r, col: c };
        return this;
    }

    autofilter(range) {
        let text = asString(range).trim().toUpperCase();
        if (text === '') {
            const from = this.layout.tableFrom ?? 0;
            const to = this.layout.tableTo ?? Math.max(0, this.cursor - 1);
            text = `${buildRef(0, from)}:${buildRef(Math.max(0, this.headerCount - 1), to)}`;
        }
        if (!parseRange(text)) {
            this.wb.warn(`excel: 自动筛选区域「${text}」不合法，已忽略`);
            return this;
        }
        this.filterRange = text;
        return this;
    }

    /** 给区域套数字格式。空格也会被创建出来，这样格式确实落在区域上。 */
    numberFormat(range, fmt) {
        const parsed = parseRange(range);
        const code = asString(fmt).trim();
        if (!parsed || code === '') {
            this.wb.warn(`excel: numberFormat 的${parsed ? '格式串为空' : `区域「${String(range)}」不合法`}，已忽略`);
            return this;
        }
        for (let row = parsed.minRow; row <= parsed.maxRow; row += 1) {
            for (let col = parsed.minCol; col <= parsed.maxCol; col += 1) {
                const cell = this.rows[row]?.[col];
                if (cell) cell.explicitFmt = code;
                else {
                    const created = this.place(buildRef(col, row), '', {});
                    if (created) created.explicitFmt = code;
                }
            }
        }
        return this;
    }

    /** 行高：非法值直接丢弃，免得生成 Excel 打不开的 ht。 */
    rowHeight(row, height) {
        const value = Number(height);
        if (!Number.isFinite(value) || value <= 0) return this;
        this.rowIndex(row).rowHeight = clamp(value, 4, 409);
        return this;
    }

    /**
     * 在这张表上放一个**原生图表**（历史遗留 `11-6`）。
     *
     * 与「Python 画一张图再嵌进来」的区别：这是真正的 DrawingML 图表部件
     * （`xl/charts/chartN.xml`），在 Excel 里是图表对象 —— 能选中、能改标题、
     * 能改图表类型。数据以 `c:numLit` / `c:strLit` 内联在图表里（不引用单元格、
     * 也不嵌工作簿），所以打开即画；要改数据得走 Office 的「编辑数据」。
     *
     * @param {{type?: 'bar'|'column'|'line'|'pie'|'area'|'scatter', title?: string,
     *   categories?: Array<string|number>, series: Array<{name?: string, values: number[], x?: number[]}>,
     *   legend?: 'bottom'|'right'|'left'|'top'|'none'|false, labels?: boolean, stacked?: boolean,
     *   gapWidth?: number, colors?: string[],
     *   at?: {col?: number, row?: number, toCol?: number, toRow?: number}}} spec
     */
    chart(spec) {
        // 立刻校验：图表参数错了要在这一行报出来（带着系列序号与长度差），
        // 而不是等到 save() 之后才在 warnings 里出现一句「图表没生成」。
        const normalized = normalizeChartSpec(spec ?? {});
        const at = (spec ?? {}).at ?? {};
        const from = { col: Math.max(0, Math.trunc(Number(at.col) || 0)), row: Math.max(0, Math.trunc(Number(at.row) || 0)) };
        const to = {
            col: Math.max(from.col + 3, Math.trunc(Number(at.toCol) || (from.col + 8))),
            row: Math.max(from.row + 6, Math.trunc(Number(at.toRow) || (from.row + 15))),
        };
        this.charts.push({ spec: normalized, from, to });
        return this;
    }

    // ── 排版计算 ──────────────────────────────────────────────────────────

    maxRow() {
        return this.rows.length - 1;
    }

    maxCol() {
        let max = -1;
        for (const cells of this.rows) {
            if (!cells) continue;
            for (let i = cells.length - 1; i >= 0; i -= 1) {
                if (cells[i]) { max = Math.max(max, i); break; }
            }
        }
        return max;
    }

    /** 每格：编译值 + 分配 xf。渲染前统一做一次，保证 XML 与 warnings 同源。 */
    prepare() {
        let cells = 0;
        let formulas = 0;
        for (let row = 0; row < this.rows.length; row += 1) {
            const cellsOfRow = this.rows[row];
            if (!cellsOfRow) continue;
            for (let col = 0; col < cellsOfRow.length; col += 1) {
                const cell = cellsOfRow[col];
                if (!cell) continue;
                cells += 1;
                if (cell.formula) formulas += 1;
                if (!cell.compiled) cell.compiled = compileValue(cell.value, cell.opts?.type ?? cell.typeHint ?? 'auto');
                // 合计行的 SUM 没有值，defaultFormat 猜不出类型会退回常规格式，
                // 结果「合计」显示成 12345 而数据行是 12,345。这里沿用同列上一行的格式，
                // 保证合计与它求和的那些格子口径一致（数据行先于合计行处理）。
                if (!cell.explicitFmt && cell.formula) {
                    const above = this.rows[row - 1]?.[col];
                    if (above?.explicitFmt) cell.explicitFmt = above.explicitFmt;
                }
                if (!cell.explicitFmt) cell.explicitFmt = defaultFormat(cell);
                this.assignStyleId(cell);
            }
        }
        return { cells, formulas };
    }

    assignStyleId(cell) {
        if (cell.styleId !== undefined) return cell.styleId;
        const opts = cell.opts ?? {};
        const s = this.styles;
        const isTitle = !!cell.isTitle;
        const size = Number(opts.size) || (isTitle ? 16 : 11);
        const fontId = s.font({
            name: opts.font ?? (isTitle ? this.wb.theme.fonts?.titleEn : this.wb.theme.fonts?.en),
            size,
            bold: !!opts.bold,
            italic: !!opts.italic,
            color: argb(opts.color, s.palette, s.colors.text),
        });
        const fillId = s.fill(opts.fill ? { argb: argb(opts.fill, s.palette) } : undefined);
        const borderId = s.border(opts.border);
        const numFmtId = s.numFmt(cell.explicitFmt);
        const styleId = s.xf({
            fontId, fillId, borderId, numFmtId,
            align: { align: opts.align, valign: opts.valign, wrap: opts.wrap, indent: opts.indent },
        });
        cell.styleId = styleId;
        return styleId;
    }

    /** 单列内容估宽：表头与数据都算，中文为主的列再放宽。 */
    estimateColumn(col) {
        let maxWidth = 0;
        let seen = false;
        for (const cells of this.rows) {
            const cell = cells?.[col];
            if (!cell || cell.isTitle || cell.isNote) continue; // 标题与说明行不参与列宽
            // 开了自动换行的格子靠折行消化宽度，不该把整列撑开：
            // 否则一句长备注就能把 A 列顶到 80 格上限。warnings 里也是同一口径。
            if (cell.style?.wrap) continue;
            if (!cell.compiled) continue;
            const width = visualWidth(displayText(cell));
            if (width > maxWidth) maxWidth = width;
            seen = true;
        }
        if (!seen) return undefined;
        // 中文列：字符格子数低估了字体实际渲染宽度，按契约 ×1.2 放宽
        const factor = this.columnCjkRatio(col) > 0.5 ? 1.2 : 1.0;
        return clamp(maxWidth * factor + 2.5, 6, 80);
    }

    columnCjkRatio(col) {
        let cjk = 0;
        let total = 0;
        for (const cells of this.rows) {
            const cell = cells?.[col];
            if (!cell?.compiled || cell.isTitle || cell.isNote) continue;
            if (cell.style?.wrap) continue; // 与 estimateColumn 保持一致，换行格不参与
            for (const ch of displayText(cell)) {
                if (/\s/.test(ch)) continue;
                total += 1;
                if (estimateEm(ch) >= 0.95) cjk += 1;
            }
        }
        return total === 0 ? 0 : cjk / total;
    }

    // ── 序列化 ────────────────────────────────────────────────────────────

    toXml() {
        const maxRow = Math.max(this.cursor - 1, this.maxRow(), this.layout.tableTo ?? 0);
        const maxCol = Math.max(this.maxCol(), this.headerCount - 1, this.columnCount - 1);
        const range = maxRow < 0 || maxCol < 0 ? 'A1' : `A1:${buildRef(maxCol, maxRow)}`;
        const parts = [XML_HEAD, `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">`];
        // 子元素顺序被 schema 钉死：dimension → sheetViews → sheetFormatPr → cols → sheetData → autoFilter → mergeCells
        parts.push(`<dimension ref="${range}"/>`);
        parts.push(this.sheetViewsXml());
        parts.push('<sheetFormatPr defaultRowHeight="15" defaultColWidth="9.140625"/>');
        parts.push(this.colsXml(maxCol));
        parts.push(`<sheetData>${this.sheetDataXml()}</sheetData>`);
        if (this.filterRange) parts.push(`<autoFilter ref="${this.filterRange}"/>`);
        if (this.merges.length > 0) {
            parts.push(`<mergeCells count="${this.merges.length}">${this.merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`);
        }
        // 图表锚在这张表上：`<drawing>` 在 CT_Worksheet 的顺序里紧跟 mergeCells 之后，
        // r:id 指向本表自己的 rels（`xl/worksheets/_rels/sheetN.xml.rels` 里的 rId1）。
        if (this.charts.length > 0) parts.push('<drawing r:id="rId1"/>');
        parts.push('</worksheet>');
        return parts.join('');
    }

    sheetViewsXml() {
        // tabSelected 是 sheetView 的**属性**，不是子元素。
        // 写成 <tabSelected val="1"/> 时 LibreOffice 会忽略、照常渲染，
        // 但真实 Excel 判定为非法内容并拒绝打开整个工作簿 —— 外部校验抓到的就是这条。
        const tab = this.wb.firstSheetName === this.name ? ' tabSelected="1"' : '';
        if (!this.freezeAt) return `<sheetViews><sheetView workbookViewId="0"${tab}/></sheetViews>`;
        const { row, col } = this.freezeAt;
        const topLeft = buildRef(col, row);
        // activePane 必须与两个 split 对应，否则 Excel 会自行「修复」文件
        const activePane = row > 0 && col > 0 ? 'bottomRight' : row > 0 ? 'bottomLeft' : 'topRight';
        const pane = `<pane${col > 0 ? ` xSplit="${col}"` : ''}${row > 0 ? ` ySplit="${row}"` : ''}`
            + ` topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>`;
        const selection = `<selection pane="${activePane}" activeCell="${topLeft}" sqref="${topLeft}"/>`;
        return `<sheetViews><sheetView workbookViewId="0"${tab}>${pane}${selection}</sheetView></sheetViews>`;
    }

    /**
     * 列宽：优先显式宽度，否则按内容估宽。
     * Excel 的 width 单位是「默认字体下的字符数」，中文正好占两格，所以估宽用
     * visualWidth 而不是 em，再加固定内边距。
     *
     * 一个 col 都写不出来时整块省略，**绝不写 `<cols></cols>`**：CT_Cols 里
     * `col` 的 minOccurs=1，空元素是非法内容，真实 Excel 会因此拒绝打开整个工作簿
     * （LibreOffice 与自研解析器都照常渲染，所以只能靠 COM 校验抓到）。
     */
    colsXml(maxCol) {
        if (maxCol < 0) return '';
        const widths = [];
        for (let col = 0; col <= maxCol; col += 1) {
            let width = this.colWidths.get(col) ?? 0;
            if (!width) width = this.estimateColumn(col) ?? 0;
            if (!width) width = this.widthHints.get(col) ?? 0;
            widths.push(width);
        }
        if (widths.every((w) => !w)) return '';
        const out = [];
        let start = 0;
        while (start <= maxCol) {
            const width = widths[start];
            let end = start;
            while (end + 1 <= maxCol && widths[end + 1] === width) end += 1;
            if (width) out.push(`<col min="${start + 1}" max="${end + 1}" width="${round2(width)}" customWidth="1"/>`);
            start = end + 1;
        }
        return out.length === 0 ? '' : `<cols>${out.join('')}</cols>`;
    }

    sheetDataXml() {
        const out = [];
        for (let row = 0; row < this.rows.length; row += 1) {
            const cells = this.rows[row];
            if (!cells) continue;
            const body = [];
            for (let col = 0; col < cells.length; col += 1) {
                const cell = cells[col];
                if (!cell) continue;
                const xml = cellXml(cell, buildRef(col, row));
                if (xml) body.push(xml);
            }
            const height = cells.rowHeight;
            if (body.length === 0 && !height) continue;
            const ht = height ? ` ht="${round2(height)}" customHeight="1"` : '';
            out.push(`<row r="${row + 1}"${ht}>${body.join('')}</row>`);
        }
        return out.join('');
    }
}

/** 生成一个 `<c>`。空值不写 t，免得 Excel 抱怨类型与内容不匹配。 */
function cellXml(cell, ref) {
    const compiled = cell.compiled ?? { kind: 'text', text: '' };
    const sAttr = cell.styleId ? ` s="${cell.styleId}"` : '';
    if (cell.formula) {
        // 手写表达式没有缓存值，只写 <f>，交给 fullCalcOnLoad 重算；
        // 统计块（count/sum/average…）由本模块真算出数字，顺带写 <v>，
        // 这样任何不重算的阅读器也能看见结论。<f> 在前 <v> 在后是 schema 顺序
        const cached = Number.isFinite(cell.cachedValue) ? `<v>${cleanNumber(cell.cachedValue)}</v>` : '';
        return `<c r="${ref}"${sAttr}><f>${escapeText(cell.formula)}</f>${cached}</c>`;
    }
    if (compiled.kind === 'number' || compiled.kind === 'date') {
        if (!Number.isFinite(compiled.number)) return '';
        return `<c r="${ref}"${sAttr} t="n"><v>${compiled.number}</v></c>`;
    }
    if (compiled.kind === 'bool') return `<c r="${ref}"${sAttr} t="b"><v>${compiled.bool ? 1 : 0}</v></c>`;
    if (compiled.text === '') return '';
    const space = /^\s|\s$/.test(compiled.text) ? ' xml:space="preserve"' : '';
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t${space}>${escapeText(compiled.text)}</t></is></c>`;
}

// ───────────────────────────────────────────────────────────────────────────
// styles.xml 与包内其它部件
// ───────────────────────────────────────────────────────────────────────────

function buildStylesXml(s) {
    const numFmts = [...s.numFmts.entries()].sort((a, b) => a[0] - b[0]);
    const parts = [XML_HEAD, `<styleSheet xmlns="${NS_MAIN}">`];
    if (numFmts.length > 0) {
        parts.push(`<numFmts count="${numFmts.length}">`
            + numFmts.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${escapeText(code)}"/>`).join('')
            + '</numFmts>');
    }
    parts.push(`<fonts count="${s.fonts.length}">${s.fonts.map(fontXml).join('')}</fonts>`);
    parts.push(`<fills count="${s.fills.length}">${s.fills.map(fillXml).join('')}</fills>`);
    parts.push(`<borders count="${s.borders.length}">${s.borders.map(borderXml).join('')}</borders>`);
    parts.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
    parts.push(`<cellXfs count="${s.xfs.length}">${s.xfs.map(xfXml).join('')}</cellXfs>`);
    parts.push('<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>');
    parts.push('<dxfs count="0"/>');
    parts.push('<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>');
    parts.push('</styleSheet>');
    return parts.join('');
}

/**
 * x:color@rgb 要的是 8 位 ARGB；主题里的颜色是 6 位 hex。
 * 少写不透明前缀时 Excel 会把它当成另一种颜色（甚至直接判文件有问题），
 * 所以统一在这里补齐，而不是逐个调用点记得手写 FF。
 */
function argb8(value) {
    const hex = String(value ?? '').replace(/^#/, '').trim();
    if (/^[0-9a-fA-F]{8}$/.test(hex)) return hex.toUpperCase();
    if (/^[0-9a-fA-F]{6}$/.test(hex)) return `FF${hex.toUpperCase()}`;
    return 'FF000000';
}

function fontXml(font) {
    const bits = ['<font>'];
    if (font.bold) bits.push('<b/>');
    if (font.italic) bits.push('<i/>');
    bits.push(`<sz val="${font.size}"/>`);
    if (font.color) bits.push(`<color rgb="${argb8(font.color)}"/>`);
    // 中西文都用同一个字体名：只给 latin 时中文会回落到宋体，观感立刻掉一档
    bits.push(`<name val="${escapeText(font.name)}"/>`);
    bits.push('<family val="2"/><scheme val="minor"/>');
    bits.push('</font>');
    return bits.join('');
}

function fillXml(fill) {
    if (fill.pattern === 'solid' && fill.argb) {
        return `<fill><patternFill patternType="solid"><fgColor rgb="${fill.argb}"/><bgColor indexed="64"/></patternFill></fill>`;
    }
    if (fill.pattern === 'gray125') return '<fill><patternFill patternType="gray125"/></fill>';
    return '<fill><patternFill patternType="none"/></fill>';
}

function borderXml(border) {
    const side = (name) => {
        const spec = border[name];
        if (!spec) return `<${name}/>`;
        const color = spec.color ? `<color rgb="${argb8(spec.color)}"/>` : '<color indexed="64"/>';
        return `<${name} style="${spec.style}">${color}</${name}>`;
    };
    return `<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border>`;
}

function xfXml(xf) {
    const attrs = [
        `numFmtId="${xf.numFmtId}"`,
        `fontId="${xf.fontId}"`,
        `fillId="${xf.fillId}"`,
        `borderId="${xf.borderId}"`,
        'xfId="0"',
    ];
    if (xf.numFmtId) attrs.push('applyNumberFormat="1"');
    if (xf.fontId) attrs.push('applyFont="1"');
    if (xf.fillId) attrs.push('applyFill="1"');
    if (xf.borderId) attrs.push('applyBorder="1"');
    if (xf.align) {
        // 对齐必须落在 <alignment> 子元素上：只写 applyAlignment 属性 Excel 不认
        attrs.push('applyAlignment="1"');
        const align = JSON.parse(xf.align);
        const inner = [];
        if (align.h) inner.push(`horizontal="${align.h}"`);
        if (align.v) inner.push(`vertical="${align.v}"`);
        if (align.wrap) inner.push('wrapText="1"');
        if (align.indent) inner.push(`indent="${align.indent}"`);
        return `<xf ${attrs.join(' ')}><alignment ${inner.join(' ')}/></xf>`;
    }
    return `<xf ${attrs.join(' ')}/>`;
}

function contentTypesXml(sheetCount, chartParts = {}) {
    const overrides = ['/xl/workbook.xml', '/xl/styles.xml'];
    for (let i = 1; i <= sheetCount; i += 1) overrides.push(`/xl/worksheets/sheet${i}.xml`);
    // 图表与绘图部件（历史遗留 11-6）：每个都要有 Override，否则真实 Excel 判「不可读取的内容」。
    const drawings = Math.max(0, Math.trunc(Number(chartParts.drawings) || 0));
    const charts = Math.max(0, Math.trunc(Number(chartParts.charts) || 0));
    for (let i = 1; i <= drawings; i += 1) overrides.push(`/xl/drawings/drawing${i}.xml`);
    for (let i = 1; i <= charts; i += 1) overrides.push(`/xl/charts/chart${i}.xml`);
    overrides.push('/docProps/core.xml', '/docProps/app.xml', '/docProps/custom.xml');
    const body = overrides.map((part) => {
        const type = part === '/xl/workbook.xml'
            ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
            : part === '/xl/styles.xml'
                ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
                : part.startsWith('/xl/worksheets/')
                    ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
                    : part.startsWith('/xl/drawings/')
                        ? CHART_DRAWING_CONTENT_TYPE
                        : part.startsWith('/xl/charts/')
                            ? CHART_CONTENT_TYPE
                            : part.endsWith('custom.xml')
                                ? 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
                                : part.endsWith('core.xml')
                                    ? 'application/vnd.openxmlformats-package.core-properties+xml'
                                    : 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
        return `<Override PartName="${part}" ContentType="${type}"/>`;
    }).join('');
    return `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + `${body}</Types>`;
}

/** 一张工作表自己的 rels（只放绘图）：`xl/worksheets/_rels/sheetN.xml.rels`。 */
function sheetRelsXml(drawingTarget) {
    return `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}">`
        + `<Relationship Id="rId1" Type="${CHART_REL_DRAWING}" Target="${drawingTarget}"/>`
        + '</Relationships>';
}

/** 一张绘图的 rels：`xl/drawings/_rels/drawingN.xml.rels`（每个图表一条关系）。 */
function drawingRelsXml(chartTargets) {
    const rels = chartTargets.map((target, index) => (
        `<Relationship Id="rId${index + 1}" Type="${CHART_REL_CHART}" Target="${target}"/>`
    )).join('');
    return `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}">${rels}</Relationships>`;
}

function rootRelsXml() {
    return `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}">`
        + `<Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/>`
        + `<Relationship Id="rId2" Type="${NS_REL}/metadata/core-properties" Target="docProps/core.xml"/>`
        + `<Relationship Id="rId3" Type="${NS_REL}/extended-properties" Target="docProps/app.xml"/>`
        + `<Relationship Id="rId4" Type="${NS_REL}/custom-properties" Target="docProps/custom.xml"/>`
        + '</Relationships>';
}

/**
 * workbook.xml 的 calcPr。
 *
 * fullCalcOnLoad="1" 是「公式缓存值可以不写」的前提，也是「缓存值万一是旧的」
 * 的保险：Excel 打开时一定重算一次。calcId="0" 表示不声明具体算表版本。
 *
 * 位置被 schema 钉死：sheets → … → definedNames → **calcPr** → … → extLst。
 * 这里由 workbookXml 单点生成，edit() 也是整份重写 workbook.xml，
 * 所以包里永远只有这一个 calcPr，不会出现两个。
 */
function calcPrXml() {
    return '<calcPr calcId="0" fullCalcOnLoad="1"/>';
}

function workbookXml(sheets) {
    const list = sheets.map((sheet, index) => (
        `<sheet name="${escapeText(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
    )).join('');
    return `${XML_HEAD}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">`
        + '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27231"/>'
        + '<workbookPr defaultThemeVersion="124226"/>'
        + '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="22000" windowHeight="12000"/></bookViews>'
        + `<sheets>${list}</sheets>`
        + calcPrXml()
        + '</workbook>';
}

function workbookRelsXml(sheetCount) {
    const rels = [];
    for (let i = 1; i <= sheetCount; i += 1) {
        rels.push(`<Relationship Id="rId${i}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i}.xml"/>`);
    }
    rels.push(`<Relationship Id="rId${sheetCount + 1}" Type="${NS_REL}/styles" Target="styles.xml"/>`);
    return `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}">${rels.join('')}</Relationships>`;
}

function coreXml(spec) {
    const iso = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    return `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"`
        + ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
        + ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        + `<dc:title>${escapeText(spec.title)}</dc:title>`
        + `<dc:creator>${escapeText(spec.author)}</dc:creator>`
        + `<cp:lastModifiedBy>${escapeText(spec.author)}</cp:lastModifiedBy>`
        + `<dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created>`
        + `<dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified>`
        + '</cp:coreProperties>';
}

/**
 * 布局元数据（表头行/数据区/合计行）的去处。
 *
 * 早先把它当成一个裸的 <vt:lpstr> 挂在 docProps/app.xml 的 <Properties> 下面，
 * 想省掉一个部件：LibreOffice 会忽略这个多余元素照常渲染，但扩展属性的 schema
 * 不允许 Properties 直接带 vt:lpstr 子元素，真实 Excel 因此拒绝打开整个工作簿。
 * 现在放进标准自定义属性部件 docProps/custom.xml —— 那里本来就是给自定义字段用的。
 */
function customXml(layoutPayload) {
    const marker = `DSHEXCEL:${JSON.stringify(layoutPayload ?? {})}`;
    return `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"`
        + ` xmlns:vt="${NS_DOC_PROPS}">`
        + '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="DSHEXCEL">'
        + `<vt:lpstr>${escapeText(marker)}</vt:lpstr>`
        + '</property></Properties>';
}

function appXml(sheets, layoutPayload) {
    const titles = sheets.map((sheet) => `<vt:lpstr>${escapeText(sheet.name)}</vt:lpstr>`).join('');
    return `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"`
        + ` xmlns:vt="${NS_DOC_PROPS}">`
        + '<Application>DSH Office Mode</Application>'
        + '<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
        + '<HeadingPairs><vt:vector size="2" baseType="variant">'
        + '<vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant>'
        + `<vt:variant><vt:i4>${Math.max(1, sheets.length)}</vt:i4></vt:variant>`
        + '</vt:vector></HeadingPairs>'
        + `<TitlesOfParts><vt:vector size="${Math.max(1, sheets.length)}" baseType="lpstr">${titles || '<vt:lpstr>Sheet1</vt:lpstr>'}</vt:vector></TitlesOfParts>`
        + '<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
        + '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>'
        + '</Properties>';
}

/** 从 custom.xml（旧版本是 app.xml）里取回布局标记。 */
function parseAppLayout(text) {
    if (!text || !text.includes('DSHEXCEL:')) return undefined;
    const from = text.indexOf('DSHEXCEL:') + 'DSHEXCEL:'.length;
    const end = text.indexOf('</', from);
    try {
        return JSON.parse(decodeEntities(text.slice(from, end === -1 ? undefined : end)));
    } catch {
        return undefined;
    }
}

// ───────────────────────────────────────────────────────────────────────────
// Workbook builder
// ───────────────────────────────────────────────────────────────────────────

class Workbook {
    constructor(spec, env) {
        this.spec = spec ?? {};
        this.env = env;
        // 报告里的 warnings 必须是「这一次构建自己产生的」。直接读 env.warnings 会把
        // 同一个 env 上别的 workbook 的警告也算进来，所以这里单独收一份；
        // 同时照旧透传 env.warn，反馈流的行为不变。
        this.ownWarnings = [];
        // 主题优先走 env.theme（契约 §2 说它返回 {theme, fellBack}），
        // 没注入时退回 resolveTheme，两条路的回落语义保持一致。
        const themed = typeof env.theme === 'function' ? env.theme(this.spec.theme) : resolveTheme(this.spec.theme);
        const resolved = themed?.theme ? themed : resolveTheme(this.spec.theme);
        this.theme = resolved.theme;
        if (resolved.fellBack) {
            this.warn(`主题「${String(resolved.requested ?? this.spec.theme)}」不存在，已回落到 ${this.theme.id}（${this.theme.name}）`);
        }
        this.styles = new StyleBook(this.theme);
        this.sheets = [];
        this.usedNames = new Set();
        this.author = asString(this.spec.author, 'DSH Office');
        this.title = asString(this.spec.title, '');
        this.firstSheetName = '';
        this.parts = 0;
    }

    sheet(name, opts = {}) {
        const base = asString(name).trim() === '' ? 'Sheet' : asString(name).trim().slice(0, 31);
        let finalName = base;
        let n = 1;
        while (this.usedNames.has(finalName) && n < 1000) {
            // Excel 不允许重名（且不区分大小写）；自动加序号比报错省模型一次往返
            n += 1;
            finalName = `${base}${n}`.slice(0, 31);
        }
        this.usedNames.add(finalName);
        const sheet = new Sheet(this, finalName, opts);
        this.sheets.push(sheet);
        if (this.sheets.length === 1) this.firstSheetName = finalName;
        return sheet;
    }

    /** 编译一遍所有单元格并汇总真实数字。写盘与报告共用同一份结论。 */
    describe() {
        const perSheet = [];
        let cells = 0;
        let formulas = 0;
        let merges = 0;
        let rows = 0;
        let columns = 0;
        for (const sheet of this.sheets) {
            const info = sheet.prepare();
            cells += info.cells;
            formulas += info.formulas;
            merges += sheet.merges.length;
            const maxRow = Math.max(sheet.maxRow(), sheet.cursor - 1);
            const maxCol = sheet.maxCol();
            rows += Math.max(0, maxRow + 1);
            columns = Math.max(columns, maxCol + 1);
            perSheet.push({
                name: sheet.name,
                rows: Math.max(0, maxRow + 1),
                columns: maxCol + 1,
                cells: info.cells,
                formulas: info.formulas,
                merges: sheet.merges.length,
                range: maxRow < 0 || maxCol < 0 ? 'A1' : `A1:${buildRef(maxCol, maxRow)}`,
                layout: sheet.layout,
                freeze: sheet.freezeAt,
                filter: sheet.filterRange,
            });
        }
        return { perSheet, cells, formulas, merges, rows, columns };
    }

    /** 布局标记：写进 custom.xml，read() 靠它精确还原 outline。 */
    layoutPayload(stats) {
        return {
            v: 1,
            theme: this.theme.id,
            sheets: stats.perSheet.map((sheet) => ({
                name: sheet.name,
                range: sheet.range,
                rows: sheet.rows,
                columns: sheet.columns,
                titleRange: sheet.layout.titleRange,
                headerRow: sheet.layout.headerRow,
                headerColumns: sheet.layout.columns,
                dataFrom: sheet.layout.dataFrom,
                dataTo: sheet.layout.dataTo,
                totalRow: sheet.layout.totalRow,
                // 统计块/汇总块的位置也要跟着走：read() 的 outline 只能靠这份标记，
                // 不能靠「猜第几行长得像统计」
                statsBlocks: (sheet.layout.statsBlocks ?? []).map((block) => ({
                    at: block.at, rows: block.rows, cols: block.cols, source: block.source, funcs: block.funcs,
                })),
                summaryBlocks: (sheet.layout.summaryBlocks ?? []).map((block) => ({
                    at: block.at, rows: block.rows, cols: block.cols, groups: block.groups,
                })),
            })),
        };
    }

    buildEntries(stats) {
        // 图表部件先编号：一张表最多一张 drawing，drawing 里按顺序挂它自己的图表。
        // 编号从 1 起、全局连续（xl/charts/chart1..N、xl/drawings/drawing1..M），
        // [Content_Types].xml 的 Override 与 rels 的 Target 都用同一套编号。
        const sheetCharts = this.sheets.map((sheet) => sheet.charts ?? []);
        const drawingCount = sheetCharts.filter((list) => list.length > 0).length;
        const chartCount = sheetCharts.reduce((sum, list) => sum + list.length, 0);
        const entries = [
            // [Content_Types].xml 必须排在包首：某些解析器按顺序判断包类型
            { name: '[Content_Types].xml', data: contentTypesXml(this.sheets.length, { drawings: drawingCount, charts: chartCount }) },
            { name: '_rels/.rels', data: rootRelsXml() },
            { name: 'xl/workbook.xml', data: workbookXml(this.sheets) },
            { name: 'xl/_rels/workbook.xml.rels', data: workbookRelsXml(this.sheets.length) },
            { name: 'xl/styles.xml', data: buildStylesXml(this.styles) },
        ];
        let drawingIndex = 0;
        let chartIndex = 0;
        this.sheets.forEach((sheet, index) => {
            entries.push({ name: `xl/worksheets/sheet${index + 1}.xml`, data: sheet.toXml() });
            const charts = sheet.charts ?? [];
            if (charts.length === 0) return;
            drawingIndex += 1;
            const chartTargets = [];
            const frames = [];
            for (const chart of charts) {
                chartIndex += 1;
                entries.push({ name: `xl/charts/chart${chartIndex}.xml`, data: chartSpaceXml(chart.spec) });
                chartTargets.push(`../charts/chart${chartIndex}.xml`);
                frames.push({
                    chartRelId: `rId${frames.length + 1}`,
                    name: `图表 ${chartIndex}`,
                    id: frames.length + 2,
                    from: chart.from,
                    to: chart.to,
                });
            }
            entries.push({
                name: `xl/drawings/drawing${drawingIndex}.xml`,
                // 一张表上有几个图表就有几个锚点（只锚第一个时 Excel 只显示那一个）
                data: spreadsheetDrawingXml({ frames }),
            });
            entries.push({ name: `xl/drawings/_rels/drawing${drawingIndex}.xml.rels`, data: drawingRelsXml(chartTargets) });
            entries.push({ name: `xl/worksheets/_rels/sheet${index + 1}.xml.rels`, data: sheetRelsXml(`../drawings/drawing${drawingIndex}.xml`) });
        });
        entries.push({ name: 'docProps/core.xml', data: coreXml({ title: this.title, author: this.author }) });
        entries.push({ name: 'docProps/app.xml', data: appXml(this.sheets, this.layoutPayload(stats)) });
        entries.push({ name: 'docProps/custom.xml', data: customXml(this.layoutPayload(stats)) });
        this.parts = entries.length;
        return entries;
    }

    /** 纯函数：不写盘，只出字节。 */
    render() {
        const stats = this.describe();
        return zip(this.buildEntries(stats));
    }

    targetPath(path) {
        const raw = asString(path, this.spec.path) || `${slugify(this.title || 'workbook')}.xlsx`;
        return ensureExtension(raw, '.xlsx');
    }

    /** 记一条构建期警告：既进报告，也透传给 env，反馈流照旧。 */
    warn(message) {
        const text = String(message);
        this.ownWarnings.push(text);
        this.env.warn(text);
        return this;
    }

    /**
     * 报告里的 warnings = 构建期警告（非法输入、主题回落）+ 排版警告。
     * 构建期警告走的是 env.warn，不并进来模型就只能看到一半风险。
     */
    mergeWarnings(layoutWarnings) {
        const seen = new Set();
        const out = [];
        for (const message of [...this.ownWarnings, ...layoutWarnings]) {
            if (seen.has(message)) continue;
            seen.add(message);
            out.push(message);
        }
        for (const message of layoutWarnings) this.env.warn(message);
        return out;
    }

    save(path) {
        const stats = this.describe();
        const bytes = zip(this.buildEntries(stats));
        const written = this.env.writeFile(this.targetPath(path), bytes);
        const warnings = this.mergeWarnings(this.collectWarnings());
        return {
            ok: true,
            format: 'xlsx',
            path: written.path,
            bytes: written.bytes,
            theme: this.theme.id,
            stats: {
                sheets: stats.perSheet.length,
                rows: stats.rows,
                columns: stats.columns,
                cells: stats.cells,
                formulas: stats.formulas,
                merges: stats.merges,
                bytes: written.bytes,
                // 统计块/汇总块的个数：模型只看 stats 就知道这次有没有真的生成统计
                statsBlocks: stats.perSheet.reduce((n, sheet) => n + (sheet.layout.statsBlocks?.length ?? 0), 0),
                summaryBlocks: stats.perSheet.reduce((n, sheet) => n + (sheet.layout.summaryBlocks?.length ?? 0), 0),
                perSheet: stats.perSheet.map((sheet) => ({
                    name: sheet.name, rows: sheet.rows, columns: sheet.columns,
                    range: sheet.range, formulas: sheet.formulas, merges: sheet.merges,
                })),
                fonts: this.styles.fonts.length,
                fills: this.styles.fills.length,
                numberFormats: this.styles.numFmts.size,
            },
            outline: this.outlineOf(stats),
            warnings,
            parts: this.parts,
        };
    }

    outlineOf(stats) {
        return stats.perSheet.map((sheet) => {
            const layout = sheet.layout ?? {};
            const bits = [];
            if (layout.titleRange) bits.push('标题');
            if (layout.headerRow !== undefined) {
                bits.push(`表头在第 ${layout.headerRow + 1} 行`);
                if (layout.dataFrom !== undefined) {
                    bits.push(`${layout.dataTo - layout.dataFrom + 1} 行数据（${layout.dataFrom + 1}–${layout.dataTo + 1}）`);
                }
                if (layout.totalRow !== undefined) bits.push(`合计在第 ${layout.totalRow + 1} 行`);
            }
            bits.push(...describeBlocks(layout));
            if (sheet.freeze) bits.push('冻结窗格');
            if (sheet.filter) bits.push(`自动筛选 ${sheet.filter}`);
            if (sheet.cells === 0) bits.push('空表');
            return `【${sheet.name}】${sheet.range} ${sheet.rows} 行 × ${sheet.columns} 列${bits.length ? '：' + bits.join('，') : ''}`;
        });
    }

    /** 所有排版风险都在这里算出来（调用前必须已经 describe()，值已编译）。 */
    collectWarnings() {
        const messages = [];
        for (const sheet of this.sheets) {
            const maxRow = Math.max(sheet.maxRow(), sheet.cursor - 1);
            const maxCol = sheet.maxCol();
            for (let row = 0; row <= maxRow; row += 1) {
                const cells = sheet.rows[row];
                if (!cells) continue;
                for (let col = 0; col <= maxCol; col += 1) {
                    const cell = cells[col];
                    if (!cell?.compiled) continue;
                    const ref = `${sheet.name}!${buildRef(col, row)}`;
                    // ① 字面错误值：多半是粘贴断链公式留下的
                    if (cell.compiled.kind === 'text') {
                        const text = cell.compiled.text.trim();
                        if (text && ERROR_LITERALS.has(text.toUpperCase())) {
                            messages.push(`公式错误字面量：${ref} 的值是 ${text}，Excel 里会显示成错误`);
                        }
                    }
                    // ② 公式引用了不存在的表名
                    if (cell.formula) {
                        for (const name of formulaSheetRefs(cell.formula)) {
                            if (!sheet.wb.usedNames.has(name)) {
                                messages.push(`公式引用了不存在的表名：${ref} → 「${name}」`);
                            }
                        }
                    }
                    // ③ 单格文字比列宽长：会溢出到邻格或被截断
                    if (cell.isTitle || cell.isNote || cell.formula || cell.style?.wrap) continue;
                    if (cell.compiled.kind !== 'text' && cell.compiled.kind !== 'bool') continue;
                    const width = sheet.colWidths.get(col) ?? sheet.estimateColumn(col);
                    if (!width) continue;
                    const need = visualWidth(displayText(cell));
                    // 阈值用「内容宽 + 1.5 格」：固定余量比比例阈值稳，不会对临界值过敏
                    if (need > width + 1.5) {
                        messages.push(`列宽不足：${ref} 内容约 ${Math.round(need)} 格，列宽只有 ${round2(width)} 格（${cell.compiled.text}）`);
                    }
                }
            }
            // ④ 合并区域压住数据 / 互相重叠
            const parsed = sheet.merges.map((range) => ({ range, area: parseRange(range) })).filter((item) => item.area);
            for (let i = 0; i < parsed.length; i += 1) {
                const a = parsed[i];
                const isTitle = sheet.layout?.titleRange && a.area.minRow === sheet.layout.titleRange.row && a.area.maxRow === a.area.minRow;
                if (!isTitle && intersectsData(sheet, a.area)) {
                    messages.push(`合并区域压住数据：${sheet.name}!${a.range} 与数据单元格重叠，只有左上角的值会保留`);
                }
                for (let k = i + 1; k < parsed.length; k += 1) {
                    if (rangesOverlap(a.area, parsed[k].area)) {
                        messages.push(`合并区域互相重叠：${sheet.name}!${a.range} 与 ${parsed[k].range}`);
                    }
                }
            }
        }
        const seen = new Set();
        return messages.filter((message) => (seen.has(message) ? false : (seen.add(message), true)));
    }
}

/** 公式里出现的表名引用。支持 Sheet1!A1、'My Sheet'!A1、[Book]Sheet!A1。 */
function formulaSheetRefs(formula) {
    const names = new Set();
    const text = asString(formula);
    const quoted = /'((?:[^']|'')+)'!/.exec(text);
    if (quoted) names.add(quoted[1].replace(/''/g, "'").replace(/^\[[^\]]*\]/, ''));
    const re = /(?:^|[^A-Za-z0-9_'"[\]$:.!])(?:\[[^\]\r\n]*\])?([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_.\u4e00-\u9fff]*)!/g;
    let match;
    while ((match = re.exec(text)) !== null) names.add(match[1]);
    return [...names];
}

/**
 * 合并区域是否压住了数据。
 *
 * 注意左上角也要看：`merge('B2:C2')` 在 B2 已经有值时，Excel 的合并语义是
 * 「只保留左上角的值」—— 也就是说这里不会丢数据。真正会丢的是**被并进来的那些格**，
 * 所以左上角是允许有值的；只有当合并区里除左上角以外还有内容（含公式）时才报警。
 */
function intersectsData(sheet, area) {
    for (let row = area.minRow; row <= area.maxRow; row += 1) {
        for (let col = area.minCol; col <= area.maxCol; col += 1) {
            if (row === area.minRow && col === area.minCol) continue;
            const cell = sheet.rows[row]?.[col];
            if (!cell?.compiled) continue;
            // 公式格的值要打开文件才算，compiled 是空文本，但它确实是数据：
            // 「把合计标签跟 SUM 格并成一格」正是最典型的吃掉数据，必须报出来
            if (cell.formula) return true;
            if (cell.compiled.kind === 'text' && cell.compiled.text === '') continue;
            if (cell.compiled.kind === 'number' && !Number.isFinite(cell.compiled.number)) continue;
            return true;
        }
    }
    return false;
}

function rangesOverlap(a, b) {
    return a.minRow <= b.maxRow && b.minRow <= a.maxRow && a.minCol <= b.maxCol && b.minCol <= a.maxCol;
}

/**
 * 统计块 / 汇总块的 outline 片段。
 * 写报告与读文件两条路共用同一句描述，模型不会遇到「同一张表两种说法」。
 */
function describeBlocks(layout) {
    const bits = [];
    for (const block of layout?.statsBlocks ?? []) {
        bits.push(`统计块 ${block.rows} 行 × ${block.cols} 列（第 ${block.at + 1} 行，源 ${block.source}）`);
    }
    for (const block of layout?.summaryBlocks ?? []) {
        bits.push(`汇总块 ${block.groups} 组（第 ${block.at + 1} 行）`);
    }
    return bits;
}

/**
 * 报告里的 path 一律是「相对 root 的正斜杠路径」。
 * 调用方可能传绝对路径（测试、批量脚本），这里统一折算，报告格式才不会漂移。
 */
function reportPath(context, path) {
    const raw = asString(path);
    if (raw === '') return raw;
    const absolute = context.resolve(raw);
    const rel = relative(context.root, absolute);
    if (rel === '' || rel.startsWith('..') || isAbsolute(rel)) return raw;
    return rel.split(sep).join('/');
}

// ───────────────────────────────────────────────────────────────────────────
// 读取：把包解回内存模型
// ───────────────────────────────────────────────────────────────────────────

/** 相对关系目标 → 包内绝对部件名（rels 的 Target 是相对 part 所在目录的）。 */
function resolvePart(partName, target) {
    if (!target) return undefined;
    const clean = String(target).replace(/\\/g, '/');
    if (clean.startsWith('/')) return clean.slice(1);
    const stack = partName.split('/').slice(0, -1);
    for (const seg of clean.split('/')) {
        if (seg === '' || seg === '.') continue;
        if (seg === '..') stack.pop();
        else stack.push(seg);
    }
    return stack.join('/');
}

function emptySheet() {
    return {
        rows: [], merges: [], formats: new Map(), formulas: [],
        freeze: undefined, filter: undefined, widths: new Map(),
        rowCount: 0, colCount: 0,
    };
}

/** 去掉尾部的空串（行样例与表头列名都用它，两边口径才一致）。 */
function trimTrailingEmpty(values) {
    const out = values.slice();
    while (out.length > 0 && out[out.length - 1] === '') out.pop();
    return out;
}

function parseRels(text) {
    const map = new Map();
    if (!text) return map;
    for (const node of descendants(parseXml(text).children[0], 'Relationship')) {
        map.set(node.attrs.Id ?? '', node.attrs.Target ?? '');
    }
    return map;
}

function parseStyles(text) {
    const index = { numFmts: new Map(), xfs: [], fonts: 0, fills: 0, borders: 0 };
    if (!text) return index;
    const root = parseXml(text).children[0];
    for (const node of descendants(root, 'numFmt')) {
        const id = Number(node.attrs.numFmtId);
        if (Number.isFinite(id)) index.numFmts.set(id, node.attrs.formatCode ?? '');
    }
    const cellXfs = descendants(root, 'cellXfs')[0];
    if (cellXfs) {
        for (const xf of xmlChildren(cellXfs, 'xf')) {
            index.xfs.push({
                numFmtId: Number(xf.attrs.numFmtId) || 0,
                fontId: Number(xf.attrs.fontId) || 0,
                fillId: Number(xf.attrs.fillId) || 0,
                borderId: Number(xf.attrs.borderId) || 0,
            });
        }
    }
    index.fonts = xmlChildren(descendants(root, 'fonts')[0]).length;
    index.fills = xmlChildren(descendants(root, 'fills')[0]).length;
    index.borders = xmlChildren(descendants(root, 'borders')[0]).length;
    return index;
}

function parseSharedStrings(text) {
    if (!text) return [];
    return descendants(parseXml(text).children[0], 'si').map((si) => textOf(si));
}

/** 读一个 `<c>`：内联字符串 / 共享字符串 / 数字 / 布尔 / 公式。 */
function readCellValue(cellNode, sharedStrings) {
    const type = cellNode.attrs.t;
    const formulaNode = descendants(cellNode, 'f')[0];
    const formula = formulaNode ? textOf(formulaNode).trim() : undefined;
    const vNode = xmlChildren(cellNode, 'v')[0];
    const raw = vNode ? textOf(vNode) : '';
    // 带缓存的公式格（统计块就是这种）：把数值一并带回去。edit() 会整份重写工作表，
    // 不带上它，「不重算也能看到数字」这件事就会在编辑后悄悄丢掉
    const cached = formula && raw !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined;
    const value = readCellPayload(cellNode, type, formula, raw, sharedStrings);
    return cached === undefined ? value : { ...value, cached };
}

/** 按 `t` 判类型取值。与 readCellValue 拆开，只为让缓存值那条支路保持单一出口。 */
function readCellPayload(cellNode, type, formula, raw, sharedStrings) {
    if (type === 'inlineStr') {
        const text = descendants(cellNode, 'is').map((node) => textOf(node)).join('');
        return { text, raw: text, kind: 'text', formula };
    }
    if (type === 's') return { text: sharedStrings[Number(raw)] ?? '', raw, kind: 'text', formula };
    if (type === 'b') return { text: raw === '1' ? 'TRUE' : 'FALSE', raw, kind: 'bool', formula };
    if (type === 'str') return { text: raw, raw, kind: 'text', formula };
    if (type === 'e') return { text: raw, raw, kind: 'error', formula };
    if (raw === '') return { text: '', raw: '', kind: 'empty', formula };
    const n = Number(raw);
    if (!Number.isFinite(n)) return { text: raw, raw, kind: 'text', formula };
    return { text: String(n), raw, kind: 'number', number: n, formula };
}

/** 解析一张 sheet：单元格、合并、列宽、冻结、筛选。 */
function parseSheet(text, styleIndex, sharedStrings) {
    const result = emptySheet();
    const root = parseXml(text).children[0];
    let maxRow = -1;
    let maxCol = -1;
    const sheetData = descendants(root, 'sheetData')[0];
    if (sheetData) {
        for (const rowNode of xmlChildren(sheetData, 'row')) {
            const rowIndex = Number(rowNode.attrs.r) - 1;
            if (!Number.isFinite(rowIndex) || rowIndex < 0) continue;
            maxRow = Math.max(maxRow, rowIndex);
            const cells = result.rows[rowIndex] ?? [];
            result.rows[rowIndex] = cells;
            for (const cellNode of xmlChildren(rowNode, 'c')) {
                const at = parseRef(cellNode.attrs.r);
                if (!at) continue;
                maxCol = Math.max(maxCol, at.col);
                const value = readCellValue(cellNode, sharedStrings);
                const styleId = Number(cellNode.attrs.s) || 0;
                const numFmtId = styleIndex.xfs[styleId]?.numFmtId ?? 0;
                cells[at.col] = {
                    col: at.col, value: value.text, formula: value.formula, raw: value.raw,
                    kind: value.kind, numFmtId, styleId, number: value.number,
                    cachedValue: value.cached,
                };
                if (value.formula) result.formulas.push({ ref: buildRef(at.col, rowIndex), formula: value.formula });
                const fmt = styleIndex.numFmts.get(numFmtId) ?? BUILTIN_NUMFMTS.get(numFmtId);
                if (fmt && fmt !== 'General' && fmt !== '@') result.formats.set(`${rowIndex}:${at.col}`, fmt);
            }
        }
    }
    for (const node of descendants(root, 'mergeCell')) {
        if (node.attrs.ref) result.merges.push(node.attrs.ref);
    }
    for (const node of descendants(root, 'col')) {
        const min = Number(node.attrs.min) - 1;
        const max = Number(node.attrs.max) - 1;
        const width = Number(node.attrs.width);
        if (!Number.isFinite(width) || min < 0) continue;
        // 上限 200 列：异常文件不该让宽度表无限膨胀
        for (let col = min; col <= max && col < min + 200; col += 1) result.widths.set(col, width);
    }
    const pane = descendants(root, 'pane')[0];
    if (pane) result.freeze = { row: Number(pane.attrs.ySplit) || 0, col: Number(pane.attrs.xSplit) || 0 };
    const filter = descendants(root, 'autoFilter')[0];
    if (filter?.attrs.ref) result.filter = filter.attrs.ref;
    result.rowCount = maxRow + 1;
    result.colCount = maxCol + 1;
    return result;
}

/** 只读地装载一个 xlsx：解析部件、样式索引、每张表的单元格。 */
function loadPackage(bytes) {
    const files = unzip(bytes);
    const textParts = unzipText(bytes);
    const decoder = new TextDecoder('utf-8');
    const readPart = (name) => {
        if (textParts.has(name)) return textParts.get(name);
        const data = files.get(name);
        return data ? decoder.decode(data) : undefined;
    };
    const rels = parseRels(readPart('xl/_rels/workbook.xml.rels') ?? '');
    const workbookPart = readPart('xl/workbook.xml');
    if (!workbookPart) throw new Error('xlsx: 缺少 xl/workbook.xml（不是有效的工作簿）');
    const root = parseXml(workbookPart).children[0];
    const styleIndex = parseStyles(readPart('xl/styles.xml') ?? '');
    const sharedStrings = parseSharedStrings(readPart('xl/sharedStrings.xml') ?? '');
    // 布局标记的新家在 custom.xml；仍然兜底读一次 app.xml，好让旧文件也能读回来。
    const layoutPayload = parseAppLayout(readPart('docProps/custom.xml') ?? '')
        ?? parseAppLayout(readPart('docProps/app.xml') ?? '');
    const themeId = asString(layoutPayload?.theme, '');

    const sheets = [];
    descendants(root, 'sheet').forEach((node, index) => {
        const name = node.attrs.name ?? `Sheet${index + 1}`;
        const part = resolvePart('xl/workbook.xml', rels.get(node.attrs['r:id'] ?? ''))
            ?? `xl/worksheets/sheet${index + 1}.xml`;
        const xml = readPart(part);
        sheets.push({
            name, part, xml,
            ...(xml ? parseSheet(xml, styleIndex, sharedStrings) : emptySheet()),
            // 这张表上有没有绘图（图表）：就地编辑时要把 <drawing r:id> 原样带回去，
            // 否则图表部件还在包里、却没有任何东西引用它 —— 打开看不到图。
            drawing: drawingRefOf(xml),
            layout: layoutPayload?.sheets?.[index],
        });
    });
    if (sheets.length === 0) {
        // 没有 <sheets> 的包只能按约定路径兜底，总比返回空报告好
        [...files.keys()].filter((key) => /^xl\/worksheets\/sheet\d+\.xml$/.test(key)).sort()
            .forEach((part, index) => {
                const xml = readPart(part);
                sheets.push({ name: `Sheet${index + 1}`, part, xml, ...parseSheet(xml, styleIndex, sharedStrings), drawing: drawingRefOf(xml) });
            });
    }
    return { files, sheets, styleIndex, sharedStrings, layoutPayload, themeId };
}

/** 工作表根节点上的 `<drawing r:id>`（没有就 undefined）。 */
function drawingRefOf(xml) {
    if (typeof xml !== 'string' || xml === '') return undefined;
    try {
        const root = parseXml(xml).children[0];
        const node = descendants(root, 'drawing')[0];
        const id = node?.attrs?.['r:id'];
        return typeof id === 'string' && id !== '' ? id : undefined;
    } catch {
        return undefined;
    }
}

// ───────────────────────────────────────────────────────────────────────────
// 对外 API
// ───────────────────────────────────────────────────────────────────────────

export const meta = {
    id: 'excel',
    name: 'Excel 表格',
    ext: '.xlsx',
    summary: '生成带标题、表头、合计行、统计块（求和/平均/中位数…真公式 + 缓存值）与分组汇总（SUMIF/COUNTIF/AVERAGEIF）的 .xlsx 表格，可读回、可批量编辑；默认素色网格（无背景填充）',
    // office_help 的默认层用这一份（提示词预算；全文层留给 detail:true）。
    // 签名与参数名必须与 methods 一致 —— 分层省的是解释与举例，不是接口。
    brief: {
        create: [
            'create({theme, path, author, title}, env) → wb',
            '  wb.sheet(name, opts?) → sheet（重名自动加序号）',
            '  sheet.title(text, {span, size, height, fill}) / sheet.note(text)',
            '  sheet.header(columns) / sheet.table({columns, rows, startRow, totalRow, zebra, freeze, autofilter})',
            '  sheet.row(values, {bold, fill, color, height, align, width})',
            '  sheet.cell(ref, value, {bold, italic, fill, color, align, format, size, border, wrap, width})',
            '  sheet.formula(ref, "SUM(B2:B9)")',
            '  sheet.stats(source, {at, label, funcs, layout, format, header})   统计块（真公式 + 缓存值）',
            '    funcs: count/counta/sum/average/median/max/min/stdev/var/product',
            '  sheet.summary({key, value, at, funcs, label, total, format})      分组汇总',
            '  sheet.merge(range) / widths({A:12}) / freeze(row, col) / autofilter(range?) / numberFormat(range, fmt)',
            '  sheet.chart({type, title, categories, series:[{name, values}]})  原生图表',
            '    type: bar / column（默认）/ line / pie / area / scatter；数据内联在图表部件里',
            '  wb.render() → Uint8Array / wb.save(path?) → report',
        ],
        read: ['read(path, env) → report（perSheet: name/range/rows 前 5 行/formulas）'],
        edit: [
            'edit(path, ops, env) → report（applied / skipped）',
            '  [{sheet, cell, value}] / [{sheet, row:{at, values}}] / [{sheet, formula:{ref, formula}}]',
            '  [{sheet, find, replace}] / [{sheet, addSheet}] / [{sheet, numberFormat:{range, format}}]',
        ],
    },
    methods: {
        create: [
            'create({theme, path, author, title}, env) → wb',
            '  wb.sheet(name, opts?) → sheet（重名自动加序号）',
            '  sheet.title(text, {span, size, height, fill})',
            '  sheet.note(text)',
            '  sheet.header(columns) / sheet.table({columns, rows, startRow, totalRow, zebra, freeze, autofilter})',
            '    zebra 缺省跟随主题：素色主题不画斑马纹，传 zebra: true 可强制要',
            '  sheet.row(values, {bold, fill, color, height, align, width})',
            '  sheet.cell(ref, value, {bold, italic, fill, color, align, format, size, border, wrap, width})',
            '  sheet.formula(ref, "SUM(B2:B9)")',
            '  sheet.stats(source, {at, label, funcs, layout, format, header}) → 统计块',
            '    source: "B5:B8" / {range:"B5:B8"} / {column:"B"|2, from:5, to:8} / "B"（取 table() 数据区）',
            '    funcs 默认 count/sum/average/max/min，支持 count/counta/sum/average/median/max/min/stdev/var/product',
            '    元素可写成 {id:"average", label:"均值", format:"0.00"}；layout: "rows"（默认两列）| "columns"（表头行+数值行）',
            '    每个数值格同时写真公式与真算出来的缓存值，不重算的阅读器也能看到数字',
            '  sheet.summary({key, value, at, funcs, label, total, format}) → 分组汇总（轻量透视）',
            '    key/value 可用列字母、1 基列号或表头标题；数据区缺省取最近一次 table()；total: true 补合计行',
            '  sheet.merge(range) / sheet.widths({A:12, B:24}) / sheet.freeze(row, col) / sheet.autofilter(range?)',
            '  sheet.numberFormat(range, "#,##0.00")',
            '  sheet.chart({type, title, categories, series, legend, labels, stacked, gapWidth, colors, at}) → 原生图表',
            '    type: bar / column（默认）/ line / pie / area / scatter；series: [{name, values, x?}]（散点图用 x）',
            '    categories 与每个系列的 values 必须等长（不等长当场报错，不会画一半）',
            '    legend: bottom（默认）/ right / left / top / none；labels: true 显示数据标签；stacked: true 堆叠',
            '    at: {col, row, toCol, toRow} 锚在哪块单元格上（默认 A1 起，8 列 × 15 行）',
            '    数据以 numLit / strLit **内联**在图表部件里：打开即画；改数据要在 Excel 里「编辑数据」',
            '  wb.render() → Uint8Array / wb.save(path?) → report',
        ],
        read: ['read(path, env) → report（perSheet: name/range/rows 前 5 行/formulas）'],
        edit: [
            'edit(path, ops, env) → report（applied / skipped）',
            '  [{sheet, cell, value}] / [{sheet, row:{at, values}}] / [{sheet, formula:{ref, formula}}]',
            '  [{sheet, find, replace}] / [{sheet, addSheet}] / [{sheet, numberFormat:{range, format}}]',
        ],
    },
};

/**
 * 建新工作簿。env 缺省时用内存环境，方便单测与预览。
 * @param {object} spec {theme, path, author, title}
 * @param {object} env 见 SPEC §2
 */
export function create(spec, env) {
    const context = env ?? createEnv({ root: process.cwd() });
    if (typeof context.theme !== 'function') {
        // env.theme 由 createEnv 的 themeResolver 注入；缺省时自己补一个，调用方不必先装配
        context.theme = (id) => resolveTheme(id);
    }
    return new Workbook(spec, context);
}

/**
 * 读已有工作簿。返回与 save() 同构的报告，另带 perSheet 明细。
 * @param {string} path 相对 env.root
 * @param {object} env
 */
export function read(path, env) {
    const context = env ?? createEnv({ root: process.cwd() });
    const bytes = context.readFile(path);
    const loaded = loadPackage(bytes);
    const relPath = reportPath(context, path);
    const allNames = new Set(loaded.sheets.map((sheet) => sheet.name));
    const perSheet = [];
    const warnings = [];
    let totalCells = 0;
    let totalFormulas = 0;
    let maxColumns = 0;

    for (const sheet of loaded.sheets) {
        let cells = 0;
        const sample = [];
        for (let row = 0; row < sheet.rowCount; row += 1) {
            const cellsOfRow = sheet.rows[row] ?? [];
            for (const cell of cellsOfRow) {
                if (!cell) continue;
                // 只算「有内容」的格：空 inlineStr 与纯样式格不算，否则 cells 会虚高
                if (!cell.formula && (cell.kind === 'empty' || String(cell.value ?? '') === '')) continue;
                cells += 1;
                if (cell.formula) {
                    totalFormulas += 1;
                    for (const name of formulaSheetRefs(cell.formula)) {
                        if (!allNames.has(name)) {
                            warnings.push(`公式引用了不存在的表名：${sheet.name}!${buildRef(cell.col, row)} → 「${name}」`);
                        }
                    }
                } else if (cell.kind === 'error') {
                    warnings.push(`单元格是错误值：${sheet.name}!${buildRef(cell.col, row)} = ${cell.raw}`);
                } else if (cell.kind === 'text' && ERROR_LITERALS.has(String(cell.value).trim().toUpperCase())) {
                    warnings.push(`公式错误字面量：${sheet.name}!${buildRef(cell.col, row)} = ${cell.value}`);
                }
            }
            if (row < 5) {
                const values = [];
                for (let col = 0; col < Math.min(sheet.colCount, 40); col += 1) {
                    values.push(sheet.rows[row]?.[col]?.value ?? '');
                }
                sample.push(trimTrailingEmpty(values));
            }
        }
        totalCells += cells;
        maxColumns = Math.max(maxColumns, sheet.colCount);
        const range = sheet.rowCount === 0 || sheet.colCount === 0
            ? 'A1'
            : `A1:${buildRef(sheet.colCount - 1, sheet.rowCount - 1)}`;

        // 列宽不足：拿文件里真实写下的 col width 对比真实内容宽度。
        //
        // 只报「真的会被截断」的格子：Excel 在该格右侧为空时会直接把文字溢出显示，
        // 观感正常（典型例子是跨越整张表的说明行）。把那种情况也报出来会变成噪声，
        // 而噪声会让模型不再相信 warnings。
        const mergeSpan = new Map();
        for (const ref of sheet.merges) {
            const [start] = String(ref).split(':');
            if (start) mergeSpan.set(start.toUpperCase(), parseRange(ref));
        }
        for (let col = 0; col < sheet.colCount; col += 1) {
            const width = sheet.widths.get(col);
            if (!width) continue;
            let worst;
            for (let row = 0; row < sheet.rowCount; row += 1) {
                const cell = sheet.rows[row]?.[col];
                if (!cell || cell.formula) continue;
                if (cell.kind !== 'text' && cell.kind !== 'bool') continue;
                if (String(cell.value).includes('\n')) continue;

                const merged = mergeSpan.get(buildRef(col, row).toUpperCase());
                const span = merged ? merged.maxCol - merged.minCol + 1 : 1;
                let effective = width;
                if (merged) {
                    effective = 0;
                    for (let c = merged.minCol; c <= merged.maxCol; c += 1) effective += sheet.widths.get(c) ?? 8.43;
                }
                const neighbour = sheet.rows[row]?.[col + span];
                const rightOccupied = neighbour !== undefined && neighbour !== null && displayText(neighbour) !== '';
                // 右邻被占用一定截断；右邻已经超出已用范围时也没有溢出空间可用，
                // 仍然按「这列太窄」报出来 —— 这是用户最常遇到的观感问题。
                if (!rightOccupied && col + span < sheet.colCount) continue;

                const need = visualWidth(cell.value);
                if (!worst || need > worst.need) worst = { need, row, width: effective, text: cell.value };
            }
            if (worst && worst.need > worst.width + 1.5) {
                warnings.push(`列宽不足：${sheet.name}!${buildRef(col, worst.row)} 内容约 ${Math.round(worst.need)} 格，列宽只有 ${round2(worst.width)} 格（${worst.text}）`);
            }
        }

        perSheet.push({
            name: sheet.name,
            range,
            rows: sheet.rowCount,
            columns: sheet.colCount,
            cells,
            formulas: sheet.formulas.length,
            merges: sheet.merges.length,
            sample,
            // 表头在第几行由 app.xml 的布局标记决定：写文件的模块知道，
            // 读文件的模块不该瞎猜第 0 行（标题行）当表头，否则按列名找列永远找不到。
            // 列数同样按布局取前 N 格：表头右边可能还挂着说明文字，不能当成列名。
            headers: (() => {
                const rowIndex = Number.isInteger(sheet.layout?.headerRow) ? sheet.layout.headerRow : 0;
                const cells = sheet.rows[rowIndex] ?? [];
                const width = Number.isInteger(sheet.layout?.headerColumns) && sheet.layout.headerColumns > 0
                    ? sheet.layout.headerColumns
                    : cells.length;
                return trimTrailingEmpty(cells.slice(0, width).map((cell) => cell?.value ?? ''));
            })(),
            freeze: sheet.freeze,
            filter: sheet.filter,
            layout: sheet.layout,
        });
    }

    const theme = loaded.themeId && resolveTheme(loaded.themeId).theme.id === loaded.themeId ? loaded.themeId : DEFAULT_THEME_ID;
    // 原生图表（历史遗留 11-6）：读的时候要把它们认出来 —— 以前图表部件在包里，
    // 但 read 只看 worksheet / styles / strings，报告里一个字都没有，模型会以为没图表。
    const chartParts = [...loaded.files.keys()].filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name)).sort();
    const charts = chartParts.map((name) => {
        const text = loaded.files.get(name) ? new TextDecoder('utf-8').decode(loaded.files.get(name)) : '';
        return { part: name, ...describeChartSpace(text) };
    });
    const outline = perSheet.map((sheet) => {
        const layout = sheet.layout ?? {};
        const bits = [];
        if (layout.titleRange) bits.push('标题');
        if (layout.headerRow !== undefined) {
            bits.push(`表头在第 ${layout.headerRow + 1} 行`);
            if (layout.dataFrom !== undefined) bits.push(`${layout.dataTo - layout.dataFrom + 1} 行数据`);
            if (layout.totalRow !== undefined) bits.push(`合计在第 ${layout.totalRow + 1} 行`);
        }
        bits.push(...describeBlocks(layout));
        if (sheet.freeze) bits.push('冻结窗格');
        if (sheet.filter) bits.push(`自动筛选 ${sheet.filter}`);
        if (sheet.cells === 0) bits.push('空表');
        return `【${sheet.name}】${sheet.range} ${sheet.rows} 行 × ${sheet.columns} 列${bits.length ? '：' + bits.join('，') : ''}`;
    });
    for (const chart of charts) {
        outline.push(`【图表】${chart.part}：${chart.type}${chart.title === '' ? '' : `「${chart.title}」`}，${chart.series} 个系列`);
    }

    const unique = [...new Set(warnings)];
    for (const message of unique) context.warn(message);
    return {
        ok: true,
        format: 'xlsx',
        path: relPath,
        bytes: bytes.length,
        theme,
        stats: {
            sheets: perSheet.length,
            rows: perSheet.reduce((sum, sheet) => sum + sheet.rows, 0),
            columns: maxColumns,
            cells: totalCells,
            formulas: totalFormulas,
            merges: perSheet.reduce((sum, sheet) => sum + sheet.merges, 0),
            bytes: bytes.length,
            charts: charts.length,
            statsBlocks: perSheet.reduce((n, sheet) => n + (sheet.layout?.statsBlocks?.length ?? 0), 0),
            summaryBlocks: perSheet.reduce((n, sheet) => n + (sheet.layout?.summaryBlocks?.length ?? 0), 0),
            perSheet: perSheet.map((sheet) => ({
                name: sheet.name, rows: sheet.rows, columns: sheet.columns, range: sheet.range,
                cells: sheet.cells, formulas: sheet.formulas, merges: sheet.merges,
            })),
        },
        perSheet,
        outline,
        warnings: unique,
        ...(charts.length > 0 ? { charts } : {}),
    };
}

/**
 * 批量编辑：一次调用内所有操作作用在同一份内存模型上，最后只写一次盘。
 * 未知或失败的操作进 skipped 并带原因，绝不抛。
 * @param {string} path
 * @param {Array<object>} ops
 * @param {object} env
 */
export function edit(path, ops, env) {
    const context = env ?? createEnv({ root: process.cwd() });
    const bytes = context.readFile(path);
    const loaded = loadPackage(bytes);
    const applied = [];
    const skipped = [];
    const changes = [];
    const allNames = new Set(loaded.sheets.map((sheet) => sheet.name));

    const findSheet = (name) => {
        const wanted = asString(name, '').trim();
        if (wanted === '') return loaded.sheets.length === 1 ? loaded.sheets[0] : undefined;
        const lower = wanted.toLowerCase();
        return loaded.sheets.find((sheet) => sheet.name.toLowerCase() === lower);
    };

    /** 写一格（或写公式）。新格默认沿用 xf 0（常规样式）。 */
    const touch = (sheet, ref, value, extra = {}) => {
        const at = parseRef(ref);
        if (!at) return undefined;
        const cells = sheet.rows[at.row] ?? [];
        sheet.rows[at.row] = cells;
        const cell = cells[at.col] ?? { col: at.col, value: '', kind: 'text', styleId: 0 };
        cells[at.col] = cell;
        if (extra.formula) {
            cell.formula = extra.formula;
            cell.value = '';
            cell.kind = 'empty';
            // 换了公式，旧的缓存值必然过期，必须丢掉（fullCalcOnLoad 会重算）
            cell.cachedValue = undefined;
        } else {
            cell.formula = undefined;
            const numeric = typeof value === 'number';
            cell.value = value === undefined || value === null ? '' : (numeric ? value : String(value));
            cell.kind = numeric ? 'number' : 'text';
            if (extra.format) cell.format = extra.format;
        }
        sheet.rowCount = Math.max(sheet.rowCount, at.row + 1);
        sheet.colCount = Math.max(sheet.colCount, at.col + 1);
        return { cell, at };
    };

    for (const raw of asArray(ops)) {
        const op = raw ?? {};
        const opName = op.addSheet !== undefined ? 'addSheet'
            : Object.keys(op).find((key) => key !== 'sheet' && key !== 'note') ?? 'unknown';
        try {
            // ── 加表：addSheet 给字符串就是表名；给 true（契约里的 `{sheet, addSheet: name}` 简写）
            //    时用同一个 op 的 sheet 字段当表名 —— 把 true 直接当名字会建出一张叫「true」的表，
            //    而且随后 {sheet:'追加表'} 的写入就找不到表了
            if (op.addSheet !== undefined) {
                const asked = op.addSheet;
                let wanted = typeof asked === 'string' ? asked.trim()
                    : (asked === true ? asString(op.sheet, '').trim() : '');
                if (wanted === '') {
                    let n = loaded.sheets.length + 1;
                    while (allNames.has(`Sheet${n}`)) n += 1;
                    wanted = `Sheet${n}`;
                }
                if (allNames.has(wanted)) {
                    skipped.push({ op: 'addSheet', reason: `工作表「${wanted}」已存在` });
                    continue;
                }
                loaded.sheets.push({ ...emptySheet(), name: wanted, isNew: true });
                allNames.add(wanted);
                applied.push({ op: 'addSheet', sheet: wanted });
                changes.push(`新增工作表「${wanted}」`);
                continue;
            }

            const sheet = findSheet(op.sheet);
            if (!sheet) {
                skipped.push({
                    op: opName,
                    reason: op.sheet === undefined
                        ? '文件里有多个工作表，请用 sheet 指定目标'
                        : `找不到工作表「${String(op.sheet)}」`,
                });
                continue;
            }

            if (op.cell !== undefined) {
                const at = parseRef(op.cell);
                if (!at) {
                    skipped.push({ op: 'cell', reason: `单元格地址「${String(op.cell)}」不合法` });
                    continue;
                }
                const ref = buildRef(at.col, at.row);
                const before = sheet.rows[at.row]?.[at.col]?.raw;
                touch(sheet, ref, op.value);
                applied.push({ op: 'cell', sheet: sheet.name, cell: ref, value: String(op.value ?? '') });
                changes.push(`${sheet.name}!${ref}: ${before === undefined ? '(空)' : before} → ${String(op.value ?? '')}`);
                continue;
            }

            if (op.row && typeof op.row === 'object') {
                const at = Number.isFinite(op.row.at) ? Math.floor(op.row.at) - 1 : sheet.rowCount;
                const values = asArray(op.row.values);
                if (at < 0) {
                    skipped.push({ op: 'row', reason: `行号「${String(op.row.at)}」不合法（从 1 开始）` });
                    continue;
                }
                values.forEach((value, index) => touch(sheet, buildRef(index, at), value));
                applied.push({ op: 'row', sheet: sheet.name, at: at + 1, cells: values.length });
                changes.push(`${sheet.name} 第 ${at + 1} 行写入 ${values.length} 个值`);
                continue;
            }

            if (op.formula && typeof op.formula === 'object') {
                const ref = asString(op.formula.ref, '').trim();
                const formula = asString(op.formula.formula, '').trim().replace(/^=/, '');
                if (!parseRef(ref) || formula === '') {
                    skipped.push({ op: 'formula', reason: 'formula.ref 或 formula.formula 不合法' });
                    continue;
                }
                touch(sheet, ref, '', { formula });
                for (const name of formulaSheetRefs(formula)) {
                    if (!allNames.has(name)) context.warn(`公式引用了不存在的表名：${sheet.name}!${ref} → 「${name}」`);
                }
                applied.push({ op: 'formula', sheet: sheet.name, cell: ref, formula });
                changes.push(`${sheet.name}!${ref} = ${formula}`);
                continue;
            }

            if (op.find !== undefined) {
                const needle = asString(op.find);
                const replacement = asString(op.replace);
                if (needle === '') {
                    skipped.push({ op: 'find', reason: 'find 为空字符串' });
                    continue;
                }
                let hits = 0;
                for (const cells of sheet.rows) {
                    if (!cells) continue;
                    for (const cell of cells) {
                        if (!cell || cell.formula) continue;
                        if (typeof cell.value !== 'string' || !cell.value.includes(needle)) continue;
                        cell.value = cell.value.split(needle).join(replacement);
                        cell.kind = 'text';
                        hits += 1;
                    }
                }
                if (hits === 0) skipped.push({ op: 'find', sheet: sheet.name, reason: `没有找到「${needle}」` });
                else {
                    applied.push({ op: 'find', sheet: sheet.name, find: needle, replace: replacement, hits });
                    changes.push(`${sheet.name}: 替换「${needle}」→「${replacement}」共 ${hits} 处`);
                }
                continue;
            }

            if (op.numberFormat && typeof op.numberFormat === 'object') {
                const range = parseRange(op.numberFormat.range);
                const fmt = asString(op.numberFormat.format, '').trim();
                if (!range || fmt === '') {
                    skipped.push({ op: 'numberFormat', reason: 'numberFormat.range / .format 不合法' });
                    continue;
                }
                let hits = 0;
                for (let row = range.minRow; row <= range.maxRow; row += 1) {
                    for (let col = range.minCol; col <= range.maxCol; col += 1) {
                        const cell = sheet.rows[row]?.[col];
                        if (!cell) continue;
                        cell.format = fmt;
                        hits += 1;
                    }
                }
                if (hits === 0) skipped.push({ op: 'numberFormat', sheet: sheet.name, reason: '区域内没有单元格' });
                else {
                    applied.push({ op: 'numberFormat', sheet: sheet.name, range: op.numberFormat.range, format: fmt, cells: hits });
                    changes.push(`${sheet.name} ${op.numberFormat.range} 套用格式 ${fmt}（${hits} 格）`);
                }
                continue;
            }

            skipped.push({ op: opName, reason: '不支持的操作（可用：cell / row / formula / find / numberFormat / addSheet）' });
        } catch (error) {
            skipped.push({ op: opName, reason: `执行出错：${error?.message ?? String(error)}` });
        }
    }

    // ── 重新打包：以原包为基础，只替换被改到的部件（styles.xml 等原样搬运）
    const carried = [];
    for (const [name, data] of loaded.files) {
        if (name === '[Content_Types].xml') continue;
        if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue;
        if (name === 'xl/workbook.xml' || name === 'xl/_rels/workbook.xml.rels' || name === 'docProps/app.xml') continue;
        // custom.xml 里是布局标记，表结构可能已经变了，必须跟着重算而不是原样搬运。
        if (name === 'docProps/custom.xml') continue;
        carried.push({ name, data });
    }
    // 图表与绘图部件原样搬运（上面的循环没跳过它们），所以重新生成的 [Content_Types].xml
    // 必须照样声明它们 —— 少一条 Override，真实 Excel 会判「发现不可读取的内容」。
    const carriedNames = [...loaded.files.keys()];
    const chartPartCount = carriedNames.filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name)).length;
    const drawingPartCount = carriedNames.filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name)).length;
    const outBytes = zip([
        { name: '[Content_Types].xml', data: contentTypesXml(loaded.sheets.length, { drawings: drawingPartCount, charts: chartPartCount }) },
        ...carried,
        { name: 'xl/workbook.xml', data: workbookXml(loaded.sheets) },
        { name: 'xl/_rels/workbook.xml.rels', data: workbookRelsXml(loaded.sheets.length) },
        ...loaded.sheets.map((sheet, index) => ({ name: `xl/worksheets/sheet${index + 1}.xml`, data: serializeEditedSheet(sheet) })),
        { name: 'docProps/app.xml', data: appXml(loaded.sheets, loaded.layoutPayload) },
        { name: 'docProps/custom.xml', data: customXml(loaded.layoutPayload) },
    ]);
    const written = context.writeFile(path, outBytes);
    const report = read(reportPath(context, path), context);
    const chartNotes = [];
    if (chartPartCount > 0) {
        chartNotes.push(`工作簿里有 ${chartPartCount} 个原生图表：就地编辑原样保留它们（部件与锚点都不动），`
            + '但不会重算图表里的数据 —— 改了数据区之后要在 Excel 里「编辑数据」或重新生成图表。');
    }
    return {
        ok: true,
        format: 'xlsx',
        path: written.path,
        bytes: outBytes.length,
        theme: report.theme,
        applied,
        skipped,
        changes,
        stats: report.stats,
        outline: report.outline,
        warnings: [...report.warnings, ...chartNotes],
    };
}

/** 把内存模型重新序列化成 sheet XML（保留原样式下标、列宽、合并、冻结、筛选）。 */
function serializeEditedSheet(sheet) {
    const parts = [XML_HEAD, `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">`];
    const rows = sheet.rowCount ?? sheet.rows.length;
    const cols = sheet.colCount ?? 0;
    parts.push(`<dimension ref="${rows === 0 || cols === 0 ? 'A1' : `A1:${buildRef(cols - 1, rows - 1)}`}"/>`);
    if (sheet.freeze && (sheet.freeze.row > 0 || sheet.freeze.col > 0)) {
        const topLeft = buildRef(sheet.freeze.col, sheet.freeze.row);
        const activePane = sheet.freeze.row > 0 && sheet.freeze.col > 0 ? 'bottomRight' : sheet.freeze.row > 0 ? 'bottomLeft' : 'topRight';
        parts.push('<sheetViews><sheetView workbookViewId="0">'
            + `<pane${sheet.freeze.col > 0 ? ` xSplit="${sheet.freeze.col}"` : ''}${sheet.freeze.row > 0 ? ` ySplit="${sheet.freeze.row}"` : ''}`
            + ` topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>`
            + `<selection pane="${activePane}" activeCell="${topLeft}" sqref="${topLeft}"/>`
            + '</sheetView></sheetViews>');
    } else {
        parts.push('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
    }
    parts.push('<sheetFormatPr defaultRowHeight="15" defaultColWidth="9.140625"/>');
    if (sheet.widths && sheet.widths.size > 0) {
        const colsXml = [...sheet.widths.entries()].sort((a, b) => a[0] - b[0])
            .map(([col, width]) => `<col min="${col + 1}" max="${col + 1}" width="${round2(width)}" customWidth="1"/>`).join('');
        parts.push(`<cols>${colsXml}</cols>`);
    }
    const body = [];
    for (let row = 0; row < rows; row += 1) {
        const cells = sheet.rows[row];
        if (!cells) continue;
        const cellXmlParts = [];
        for (let col = 0; col < cells.length; col += 1) {
            const cell = cells[col];
            if (!cell) continue;
            cellXmlParts.push(editedCellXml(cell, buildRef(col, row)));
        }
        if (cellXmlParts.length === 0) continue;
        body.push(`<row r="${row + 1}">${cellXmlParts.join('')}</row>`);
    }
    parts.push(`<sheetData>${body.join('')}</sheetData>`);
    if (sheet.filter) parts.push(`<autoFilter ref="${sheet.filter}"/>`);
    if (sheet.merges.length > 0) {
        parts.push(`<mergeCells count="${sheet.merges.length}">${sheet.merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`);
    }
    // 原表锚过绘图（图表）就把引用带回去：部件本来就原样搬运，少了这一行图表就不显示。
    if (typeof sheet.drawing === 'string' && sheet.drawing !== '') parts.push(`<drawing r:id="${sheet.drawing}"/>`);
    parts.push('</worksheet>');
    return parts.join('');
}

/** 编辑路径的单元格序列化：改过值就按新值决定类型，没改过的保持原样。 */
function editedCellXml(cell, ref) {
    const sAttr = cell.styleId ? ` s="${cell.styleId}"` : '';
    if (cell.formula) {
        // 没被改动的公式格要把原缓存值搬回去：统计块里的数字是结论，
        // 编辑别的格子不该让它们变成空白
        const cached = Number.isFinite(cell.cachedValue) ? `<v>${cleanNumber(cell.cachedValue)}</v>` : '';
        return `<c r="${ref}"${sAttr}><f>${escapeText(cell.formula)}</f>${cached}</c>`;
    }
    const text = asString(cell.value);
    if (text === '') return '';
    if (typeof cell.value === 'number' && Number.isFinite(cell.value)) {
        return `<c r="${ref}"${sAttr} t="n"><v>${cell.value}</v></c>`;
    }
    if (cell.kind === 'number' && Number.isFinite(Number(text))) {
        return `<c r="${ref}"${sAttr} t="n"><v>${Number(text)}</v></c>`;
    }
    const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t${space}>${escapeText(text)}</t></is></c>`;
}

export default { meta, create, read, edit };
