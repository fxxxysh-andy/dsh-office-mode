/**
 * 能力注册表：把「办公模式支持哪些格式」收敛成一张表。
 *
 * 扩展一种新格式（比如 Markdown 报告、CSV、Visio）只需要在这里加一行，
 * 它自动获得：office.<id> 的脚本入口、office_help 的索引与详情、以及
 * 生成后的自动复检。工具声明本身不增长，符合「开始时只声明工具、具体内容
 * 在需要时再添加」。
 *
 * 用动态 import：某个格式模块坏掉时只影响它自己那一次调用，不会让整个插件挂掉。
 *
 * @module dsh-office-mode/registry
 */

/** 格式扩展示例（新增能力时改这一张表）。 */
const FORMATS = [
    {
        id: 'word',
        kind: 'document',
        extensions: ['.docx'],
        loader: () => import('./formats/word.js'),
    },
    {
        id: 'excel',
        kind: 'workbook',
        extensions: ['.xlsx'],
        loader: () => import('./formats/excel.js'),
    },
    {
        id: 'ppt',
        kind: 'deck',
        extensions: ['.pptx'],
        loader: () => import('./formats/ppt.js'),
    },
    {
        // LaTeX 是唯一「产物由外部程序算出来」的格式：源文件插件自己写，
        // PDF 交给本机 TeX 发行版（见 src/tex.js）。扩展名注册成 .tex，
        // 生成后同样会被 read() 自动复检一遍。
        id: 'tex',
        kind: 'project',
        extensions: ['.tex'],
        loader: () => import('./formats/tex.js'),
    },
];

const byId = new Map(FORMATS.map((entry) => [entry.id, entry]));
const byExtension = new Map();
for (const entry of FORMATS) {
    for (const ext of entry.extensions) byExtension.set(ext, entry);
}

/** 全部格式 id。 */
export function formatIds() {
    return FORMATS.map((entry) => entry.id);
}

/** 按 id 取注册项（不加载模块）。 */
export function formatEntry(id) {
    return byId.get(String(id ?? '').trim().toLowerCase());
}

/** 按扩展名找格式（用于生成后的自动复检）。 */
export function formatByExtension(filePath) {
    const at = String(filePath ?? '').lastIndexOf('.');
    if (at === -1) return undefined;
    return byExtension.get(String(filePath).slice(at).toLowerCase());
}

/** 加载模块；失败返回 `{ error }` 而不抛，调用方自己决定怎么反馈。 */
export async function loadFormat(id) {
    const entry = formatEntry(id);
    if (entry === undefined) return { error: new Error(`未知的格式：${id}；可用：${formatIds().join(' / ')}`) };
    try {
        const module = await entry.loader();
        return { entry, module };
    } catch (error) {
        return { entry, error };
    }
}

/** 逐个加载全部格式，返回 `[{entry, module?, error?}]`，不抛。 */
export async function loadAllFormats() {
    return Promise.all(FORMATS.map(async (entry) => loadFormat(entry.id)));
}
