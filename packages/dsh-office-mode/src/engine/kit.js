/**
 * 引擎工具箱：单位换算、文本宽度估算、数据规范化、文件环境。
 *
 * 三种格式共用这里的换算与度量。文本宽度估算是「排版反馈」的基础：
 * 它让生成结果在写盘前就能报出「这段文字大概会溢出」，而不是等用户打开文件才发现。
 *
 * @module dsh-office-mode/engine/kit
 */
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path';

export const EMU_PER_PT = 12700;
export const EMU_PER_INCH = 914400;
export const EMU_PER_CM = 360000;
export const TWIPS_PER_PT = 20;
export const TWIPS_PER_CM = 567;
export const PT_PER_CM = 28.3464567;
export const PT_PER_INCH = 72;

export const cm = (value) => Math.round(value * TWIPS_PER_CM);
export const cmToEmu = (value) => Math.round(value * EMU_PER_CM);
export const ptToEmu = (value) => Math.round(value * EMU_PER_PT);
export const inchToEmu = (value) => Math.round(value * EMU_PER_INCH);
export const ptToTwips = (value) => Math.round(value * TWIPS_PER_PT);
export const ptToHalfPoints = (value) => Math.round(value * 2);

/** 字符宽度权重（相对字号）。中日韩全角按 1 个 em，西文按字面宽度近似。 */
const NARROW = new Set('iljI.,;:!\'|`[](){}'.split(''));
const WIDE = new Set('mwMW@%'.split(''));

const CJK = /[\u1100-\u115F\u2E80-\u303E\u3041-\u33FF\u3400-\u4DBF\u4E00-\u9FFF\uA000-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFF60\uFFE0-\uFFE6]/;

/** 估算一段文本占多少个 em（全角字符 1.0）。 */
export function estimateEm(text) {
    let total = 0;
    for (const ch of String(text ?? '')) {
        if (ch === '\n') continue;
        if (CJK.test(ch)) total += 1;
        else if (ch === ' ') total += 0.28;
        else if (ch === '\t') total += 1.1;
        else if (NARROW.has(ch)) total += 0.28;
        else if (WIDE.has(ch)) total += 0.9;
        else if (ch >= 'A' && ch <= 'Z') total += 0.66;
        else if (ch >= '0' && ch <= '9') total += 0.56;
        else total += 0.54;
    }
    return total;
}

/** 估算文本宽度（磅）。 */
export function textWidthPt(text, sizePt) {
    return estimateEm(text) * (Number(sizePt) || 11);
}

/**
 * 按宽度折行（贪心，正确处理中文逐字断行与西文按词断行）。
 * @returns {string[]}
 */
export function wrapText(text, maxWidthPt, sizePt) {
    const size = Number(sizePt) || 11;
    const limit = Math.max(size * 0.6, Number(maxWidthPt) || 0);
    const out = [];
    for (const rawLine of String(text ?? '').split('\n')) {
        if (rawLine === '') {
            out.push('');
            continue;
        }
        let line = '';
        let width = 0;
        const tokens = rawLine.match(/[A-Za-z0-9@%._#+\-/]+|\s+|[^\s]/g) ?? [];
        for (const token of tokens) {
            const tokenWidth = estimateEm(token) * size;
            if (width + tokenWidth > limit && line.trim() !== '') {
                out.push(line.replace(/\s+$/, ''));
                line = token.trimStart();
                width = estimateEm(line) * size;
            } else {
                line += token;
                width += tokenWidth;
            }
        }
        out.push(line.replace(/\s+$/, ''));
    }
    return out;
}

/** 估算折行后的行数。 */
export function estimateLines(text, maxWidthPt, sizePt) {
    return wrapText(text, maxWidthPt, sizePt).length;
}

export function asString(value, fallback = '') {
    if (typeof value === 'string') return value;
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    return fallback;
}

export function asNumber(value, fallback = 0) {
    const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
    return Number.isFinite(n) ? n : fallback;
}

export function asBool(value, fallback = false) {
    return typeof value === 'boolean' ? value : fallback;
}

export function asArray(value) {
    return Array.isArray(value) ? value : [];
}

export function clampNumber(value, min, max, fallback) {
    const n = asNumber(value, Number.NaN);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

/** 文件名安全化：保留中英文，去掉路径分隔符与非法字符。 */
export function slugify(value, fallback = 'office-doc') {
    const text = asString(value).trim()
        .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-')
        .replace(/\s+/g, '-')
        .replace(/^-+|-+$/g, '');
    return text === '' || text === '.' || text === '..' ? fallback : text.slice(0, 80);
}

/** 补上扩展名（缺省或写错时）。 */
export function ensureExtension(filePath, ext) {
    const wanted = ext.startsWith('.') ? ext : `.${ext}`;
    const current = extname(filePath).toLowerCase();
    return current === wanted.toLowerCase() ? filePath : `${filePath}${wanted}`;
}

/**
 * 规范化表格列定义。
 * 接受 `['名称','金额']` 或 `[{ title:'金额', width:16, type:'number', format:'#,##0' }]`。
 */
export function normalizeColumns(columns) {
    return asArray(columns).map((raw, index) => {
        if (typeof raw === 'string' || typeof raw === 'number') {
            return { key: `c${index}`, title: String(raw), width: 0, type: 'auto', format: '', align: '' };
        }
        const item = raw ?? {};
        const type = ['auto', 'text', 'number', 'currency', 'percent', 'date'].includes(item.type) ? item.type : 'auto';
        return {
            key: asString(item.key, `c${index}`),
            title: asString(item.title ?? item.name ?? item.label ?? item.key, `列${index + 1}`),
            width: asNumber(item.width, 0),
            type,
            format: asString(item.format, ''),
            align: asString(item.align, ''),
        };
    });
}

/** 规范化表格行：不足补空，超出截断，保证列数一致。 */
export function normalizeRows(rows, width) {
    return asArray(rows).map((row) => {
        const cells = Array.isArray(row) ? row.slice() : [row];
        while (cells.length < width) cells.push('');
        return cells.slice(0, width);
    });
}

/** 数字/日期猜测，供 Excel 与 PPT 表格自动右对齐。 */
export function classifyValue(value) {
    if (typeof value === 'number') return 'number';
    if (typeof value === 'boolean') return 'text';
    // 脚本在 vm 里构造的 Date 不满足宿主侧的 instanceof，用标签判断。
    if (Object.prototype.toString.call(value) === '[object Date]') return 'date';
    const text = asString(value).trim();
    if (text === '') return 'text';
    if (/^-?[\d,]+(\.\d+)?%$/.test(text)) return 'percent';
    if (/^-?[\d,]+(\.\d+)?$/.test(text)) return 'number';
    if (/^-?[¥$€£]\s?[\d,]+(\.\d+)?$/.test(text)) return 'currency';
    if (/^\d{4}[-/年]\d{1,2}([-/月]\d{1,2}日?)?$/.test(text)) return 'date';
    return 'text';
}

/** 把列类型 + 单元格值推进为具体类型。 */
export function cellType(columnType, value) {
    if (columnType && columnType !== 'auto') return columnType;
    return classifyValue(value);
}

/** 数字字面量：去掉千分位、货币符号、百分号。 */
export function toNumber(value) {
    const text = asString(value).trim();
    if (text === '') return Number.NaN;
    const percent = text.endsWith('%');
    const cleaned = text.replace(/[¥$€£,\s]/g, '').replace(/%$/, '');
    const n = Number.parseFloat(cleaned);
    if (!Number.isFinite(n)) return Number.NaN;
    return percent ? n / 100 : n;
}

/** 简单的递增 id 工厂。 */
export function makeCounter(prefix) {
    let n = 0;
    return () => {
        n += 1;
        return `${prefix}${n}`;
    };
}

/** 相对工作目录的展示路径（写进反馈里给模型看）。 */
export function displayPath(root, absolute) {
    const rel = relative(root, absolute);
    if (rel === '' || rel.startsWith('..') || isAbsolute(rel)) return absolute;
    return rel.split(sep).join('/');
}

/**
 * 构造文件环境。格式模块只通过它读写，因此可以在测试里替换成内存实现。
 */
export function createEnv(options = {}) {
    const root = resolve(options.root ?? process.cwd());
    const themeResolver = options.themeResolver;
    const writes = [];
    const notes = [];
    const warnings = [];
    const logs = [];

    const resolvePath = (filePath) => (isAbsolute(filePath) ? resolve(filePath) : resolve(root, filePath));

    const env = {
        root,
        resolve: resolvePath,
        theme: themeResolver,
        // 插件配置（settings 页与组合层收敛后的那一份）。格式模块拿不到插件上下文，
        // 需要配置的地方（LaTeX 编译的引擎/超时/模板目录）从这里取；单测里可以为空。
        config: options.config !== null && typeof options.config === 'object' ? options.config : {},
        writes,
        notes,
        warnings,
        logs,
        writeFile(filePath, data) {
            const absolute = resolvePath(filePath);
            mkdirSync(dirname(absolute), { recursive: true });
            const bytes = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
            writeFileSync(absolute, bytes);
            const entry = { path: displayPath(root, absolute), absolute, bytes: bytes.length };
            writes.push(entry);
            return entry;
        },
        readFile(filePath) {
            return readFileSync(resolvePath(filePath));
        },
        readText(filePath) {
            return readFileSync(resolvePath(filePath), 'utf8');
        },
        exists(filePath) {
            return existsSync(resolvePath(filePath));
        },
        stat(filePath) {
            try {
                return statSync(resolvePath(filePath));
            } catch {
                return undefined;
            }
        },
        list(dirPath = '.') {
            const absolute = resolvePath(dirPath);
            if (!existsSync(absolute)) return [];
            return readdirSync(absolute, { withFileTypes: true }).map((entry) => ({
                name: entry.name,
                dir: entry.isDirectory(),
                bytes: entry.isFile() ? statSync(join(absolute, entry.name)).size : 0,
            }));
        },
        remove(filePath) {
            const absolute = resolvePath(filePath);
            if (!existsSync(absolute)) return false;
            rmSync(absolute, { recursive: true, force: true });
            return true;
        },
        log(message) {
            logs.push(String(message));
        },
        warn(message) {
            warnings.push(String(message));
        },
        note(message) {
            notes.push(String(message));
        },
    };
    return env;
}

export { basename, dirname, extname, join, relative, resolve };
