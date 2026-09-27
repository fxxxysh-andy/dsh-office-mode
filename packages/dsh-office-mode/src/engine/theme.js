/**
 * 视觉主题表（排版与审美的最小来源）。
 *
 * 一个主题 = 字体对 + 调色板 + 表格样式 + 版式留白。三种格式共用同一份数据：
 * Word 取 half-point / twips，Excel 取 ARGB，PPT 取 EMU 与 hex。
 * 扩展一个主题 = 往 THEMES 里加一条，不需要改任何渲染代码。
 *
 * 两条硬约定：
 * 1. 默认主题是 `plain`（素色网格）：不加任何背景填充，Excel 保持普通表格的
 *    网格形态、Word 保持白纸黑字。要配色必须显式传 `theme: '<id>'`。
 * 2. `table.*Fill` 写成空串表示「不填充」。渲染代码必须先判断再写 XML ——
 *    空串不是颜色，落成 `<fgColor rgb=""/>` / `<w:shd w:fill=""/>` 都是坏文件。
 *
 * @module dsh-office-mode/engine/theme
 */

/** 中文办公场景下安全可用的中文字体。 */
const CN_SANS = '微软雅黑';
const CN_SERIF = '宋体';
const CN_HEI = '黑体';
/** Word / Excel 的东亚默认字体，素色主题用它与「什么都没加」的观感一致。 */
const CN_BODY = '等线';

/**
 * @typedef {object} OfficeTheme
 * @property {string} id
 * @property {string} name       中文名，给用户看的
 * @property {string} mood       一句话气质描述
 * @property {string} bestFor    适用场景
 * @property {boolean} dark      是否深色底
 * @property {{cn:string, en:string, titleCn:string, titleEn:string}} fonts
 * @property {Record<string,string>} colors  6 位 hex，不带 #
 * @property {{titleFill:string, headerFill:string, headerText:string, zebraFill:string, border:string, accentFill:string}} table
 *   `*Fill` 为空串 = 不填充；`headerText` 为空串 = 用正文色。
 * @property {{bg?:string, band?:string, text?:string, kicker?:string, muted?:string, onBand?:string, onBandMuted?:string}} [cover]
 *   封面/结束页的底色与文字色覆盖项（PPT 用）。缺省时按 primaryDark + onPrimary 推导；
 *   素色主题没有深色可用，于是显式给出「白底黑字 + 浅灰色带」。
 */

/** @type {OfficeTheme[]} */
export const THEMES = [
    {
        id: 'plain',
        name: '素色网格',
        mood: '白底黑字、灰色细边，不加任何背景填充',
        bestFor: '默认输出、需要自己排版的初稿、要粘进别的模板的内容',
        dark: false,
        fonts: { cn: CN_BODY, en: 'Calibri', titleCn: CN_BODY, titleEn: 'Calibri' },
        colors: {
            primary: '000000',
            primaryDark: '000000',
            secondary: '404040',
            accent: '595959',
            text: '000000',
            muted: '595959',
            subtle: '808080',
            bg: 'FFFFFF',
            surface: 'FFFFFF',
            border: 'BFBFBF',
            onPrimary: 'FFFFFF',
        },
        // 空串 = 不填充。标题行、表头、斑马纹、合计行都只靠加粗与边框区分。
        table: {
            titleFill: '', headerFill: '', headerText: '', zebraFill: '', border: 'BFBFBF', accentFill: '',
        },
        // 素色主题不做深色封面：白底黑字 + 浅灰色带，避免「白字落白底」
        cover: {
            bg: 'FFFFFF', band: 'F2F2F2', text: '000000',
            kicker: '595959', muted: '595959', onBand: '000000', onBandMuted: '595959',
        },
    },
    {
        id: 'business',
        name: '商务蓝',
        mood: '稳重、克制，深浅两级蓝',
        bestFor: '汇报、方案、对外文档',
        dark: false,
        fonts: { cn: CN_SANS, en: 'Calibri', titleCn: CN_SANS, titleEn: 'Calibri' },
        colors: {
            primary: '1F4E79',
            primaryDark: '16365C',
            secondary: '2E75B6',
            accent: 'C55A11',
            text: '262626',
            muted: '595959',
            subtle: '8496B0',
            bg: 'FFFFFF',
            surface: 'F2F6FB',
            border: 'BFCEDF',
            onPrimary: 'FFFFFF',
        },
        table: {
            titleFill: 'F2F6FB', headerFill: '1F4E79', headerText: 'FFFFFF',
            zebraFill: 'F2F6FB', border: 'BFCEDF', accentFill: 'DEEAF6',
        },
    },
    {
        id: 'minimal',
        name: '简约灰',
        mood: '黑白灰，留白多，几乎不用彩色',
        bestFor: '内部备忘、审阅稿、需要打印的黑白文档',
        dark: false,
        fonts: { cn: CN_SANS, en: 'Calibri', titleCn: CN_HEI, titleEn: 'Calibri' },
        colors: {
            primary: '333333',
            primaryDark: '1A1A1A',
            secondary: '6B6B6B',
            accent: '8C8C8C',
            text: '262626',
            muted: '737373',
            subtle: 'A6A6A6',
            bg: 'FFFFFF',
            surface: 'F5F5F5',
            border: 'D9D9D9',
            onPrimary: 'FFFFFF',
        },
        table: {
            titleFill: 'F5F5F5', headerFill: '404040', headerText: 'FFFFFF',
            zebraFill: 'F5F5F5', border: 'D9D9D9', accentFill: 'EDEDED',
        },
    },
    {
        id: 'warm',
        name: '暖阳橙',
        mood: '暖色、亲和，阅读压力低',
        bestFor: '培训材料、宣传文案、面向大众的说明',
        dark: false,
        fonts: { cn: CN_SANS, en: 'Segoe UI', titleCn: CN_SANS, titleEn: 'Segoe UI' },
        colors: {
            primary: 'B45309',
            primaryDark: '8A3E06',
            secondary: 'D97706',
            accent: '0F766E',
            text: '33302B',
            muted: '6B6459',
            subtle: 'C2A98A',
            bg: 'FFFFFF',
            surface: 'FFF7ED',
            border: 'F0D9BC',
            onPrimary: 'FFFFFF',
        },
        table: {
            titleFill: 'FFF7ED', headerFill: 'B45309', headerText: 'FFFFFF',
            zebraFill: 'FFF7ED', border: 'F0D9BC', accentFill: 'FDE9D2',
        },
    },
    {
        id: 'forest',
        name: '森林绿',
        mood: '安静、可信，绿色主色',
        bestFor: '研究报告、年度总结、可持续主题',
        dark: false,
        fonts: { cn: CN_SANS, en: 'Calibri', titleCn: CN_SANS, titleEn: 'Calibri' },
        colors: {
            primary: '1F5C3A',
            primaryDark: '14402A',
            secondary: '3E8E5A',
            accent: 'B7791F',
            text: '22302A',
            muted: '55655C',
            subtle: '9DB8A8',
            bg: 'FFFFFF',
            surface: 'F1F7F3',
            border: 'C6DCCE',
            onPrimary: 'FFFFFF',
        },
        table: {
            titleFill: 'F1F7F3', headerFill: '1F5C3A', headerText: 'FFFFFF',
            zebraFill: 'F1F7F3', border: 'C6DCCE', accentFill: 'DFEEE5',
        },
    },
    {
        id: 'tech',
        name: '科技深色',
        mood: '深色底、青色点缀，屏幕观感强',
        bestFor: '产品发布、技术分享、投影演示',
        dark: true,
        fonts: { cn: CN_SANS, en: 'Segoe UI', titleCn: CN_SANS, titleEn: 'Segoe UI' },
        colors: {
            primary: '0E7490',
            primaryDark: '075985',
            secondary: '22D3EE',
            accent: 'F59E0B',
            text: 'E5E7EB',
            muted: 'A1A8B3',
            subtle: '64748B',
            bg: '111827',
            surface: '1B2432',
            border: '334155',
            onPrimary: 'F8FAFC',
        },
        table: {
            titleFill: '1B2432', headerFill: '0E7490', headerText: 'F8FAFC',
            zebraFill: '1B2432', border: '334155', accentFill: '164E63',
        },
    },
    {
        id: 'academic',
        name: '学术宋',
        mood: '宋体正文、单色标题，接近正式出版物',
        bestFor: '论文、研究报告、正式函件',
        dark: false,
        fonts: { cn: CN_SERIF, en: 'Times New Roman', titleCn: CN_HEI, titleEn: 'Times New Roman' },
        colors: {
            primary: '1A1A1A',
            primaryDark: '000000',
            secondary: '444444',
            accent: '8B1E1E',
            text: '1A1A1A',
            muted: '555555',
            subtle: '999999',
            bg: 'FFFFFF',
            surface: 'F7F7F7',
            border: 'CCCCCC',
            onPrimary: 'FFFFFF',
        },
        table: {
            titleFill: 'F7F7F7', headerFill: '333333', headerText: 'FFFFFF',
            zebraFill: 'F7F7F7', border: 'CCCCCC', accentFill: 'EDEDED',
        },
    },
];

export const DEFAULT_THEME_ID = 'plain';

/** 规范化 hex：接受 `#rgb` / `#rrggbb` / `rrggbb`，输出大写 6 位。 */
export function toHex(value, fallback = '000000') {
    if (typeof value !== 'string') return fallback;
    let text = value.trim().replace(/^#/, '');
    if (/^[0-9a-fA-F]{3}$/.test(text)) text = text.split('').map((ch) => ch + ch).join('');
    return /^[0-9a-fA-F]{6}$/.test(text) ? text.toUpperCase() : fallback;
}

/**
 * 底色归一：空串、白色都当成「不填充」，返回大写 hex 或空串。
 *
 * Word 的 `w:shd` 写 `FFFFFF` 等于给段落/单元格加了一层白色背景，与素色默认
 * 冲突；所以渲染前一律过一遍这里，只有真正有色值时才写填充元素。
 */
export function shadeOf(value) {
    const hex = toHex(value, '');
    return hex === '' || hex === 'FFFFFF' ? '' : hex;
}

/**
 * 取主题。未知 id 会回落到默认主题，并把可用 id 列表带回去，
 * 这样模型一次调用就能自我纠正，不必再发一次询问。
 */
export function resolveTheme(id) {
    if (typeof id === 'string' && id.trim() !== '') {
        const wanted = id.trim().toLowerCase();
        const hit = THEMES.find((theme) => theme.id === wanted);
        if (hit) return { theme: hit, fellBack: false };
        return { theme: THEMES.find((theme) => theme.id === DEFAULT_THEME_ID), fellBack: true, requested: id };
    }
    return { theme: THEMES.find((theme) => theme.id === DEFAULT_THEME_ID), fellBack: false };
}

/** 主题清单（给 office_help / 反馈用）。 */
export function themeCatalog() {
    return THEMES.map((theme) => ({
        id: theme.id,
        name: theme.name,
        mood: theme.mood,
        bestFor: theme.bestFor,
        dark: theme.dark,
        primary: theme.colors.primary,
    }));
}

/** 供帮助文档渲染的主题表（纯文本，避免模型再解析 JSON）。 */
export function renderThemeTable() {
    return themeCatalog()
        .map((theme) => {
            const mark = theme.id === DEFAULT_THEME_ID ? '（默认，无填充）' : '';
            return `- ${theme.id}（${theme.name}）${mark}：${theme.mood}；适合${theme.bestFor}${theme.dark ? '；深色' : ''}`;
        })
        .join('\n');
}
