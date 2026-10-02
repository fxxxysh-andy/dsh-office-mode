/**
 * 能力地图：本会话「实际有什么」由插件现算，persona 不再写死（第十八轮 P2-4 / 第二十三轮 23-2）。
 *
 * 为什么要单独一个模块：
 *   1. **persona 写死的能力清单会与运行期事实错配**（23-2）。工具清单那半已经不需要复述
 *      —— 插件注册的工具面本身就是实际清单（设置页关掉的工具不会出现，宿主每次请求
 *      重发的就是它）。工具面**表达不了**的是「引擎在不在、派工走哪条路」，而这两种
 *      错配最贵：persona 说有 ffmpeg 而本机没装，模型就会承诺一份转写稿。
 *   2. 办公 preset 的 persona 是 `complete: true` 的**唯一**提示段，插件注册的其它
 *      section（guide / 记忆说明）在办公会话里会被「完整段」规则丢掉。能进这份提示词的
 *      插件通道只有 `systemPrompt.variable()` 的 `{{变量}}` 插值，所以这段地图以变量的
 *      形式注入，见 `src/index.js` 的 `registerCapabilityMap`。
 *
 * 两条纪律：
 *   - **只报探测到的入口**：这里全是同步探测（PATH / 常见安装目录 / existsSync），
 *     查得到才算有；「能不能真跑」由各自的 check 报（python 在 Windows 上可能是应用
 *     商店的占位程序，fitz 装了没有要导入一次才知道）。地图末尾写明这一条。
 *   - **探测结果按进程缓存**：变量在每一次装配时求值，不能每次都翻一遍盘；
 *     设置改动时由 `rebuild` 带 `refresh: true` 重探。
 *
 * @module dsh-office-mode/capabilities
 */
import { existsSync } from 'node:fs';
import { basename } from 'node:path';

import { probeAvSync } from './av.js';
import { findPdfExecutable, probePdfEngines } from './pdf.js';
import { previewKitAvailable } from './preview.js';
import { probeTexEngines } from './tex.js';

/**
 * persona 里引用的变量名（`{{office_capabilities}}`）。
 *
 * 名字必须匹配 `[a-z][a-z0-9_]*`（宿主 dsh-system-prompt 的判据）。
 * `test/preset-check.mjs` 拿这个常量去核对 preset 里引用的就是它 —— 两边改名必须同时改。
 */
export const CAPABILITY_VARIABLE = 'office_capabilities';

/**
 * 注入文本的字节预算（第十八轮 P2-4 的「先减后加」：persona 的 16,300 B 预算没放宽，
 * 这一份是从 persona 的「工具」一节里省出来的额度里花的）。
 *
 * 第四十七轮复核实测最坏情况 **489 B**（引擎满 + 音视频「部分缺件」+ 派工走长的那一支 +
 * 解释器文件名顶到 16 字符上限），所以预算收到 520 —— 留 31 B 余量，涨一点就红。
 * 原来的 640 要涨 30% 才报警，等于没守（复核 瑕疵 1）。
 */
export const CAPABILITY_MAP_BUDGET_BYTES = 520;

/** PDF 渲染引擎候选（与 pdf.js 的 PDF_ENGINE_IDS 同一批，只是这里按可执行文件在不在判）。 */
const PDF_RENDER_BINS = [
    ['pdftoppm', 'pdftoppm'],
    ['pdftocairo', 'pdftocairo'],
    ['mutool', 'mutool'],
    ['gswin64c', 'gs'],
];

/** 地图里每类引擎最多列几个（剩下的收成「等」，见 renderCapabilityMap 的字节纪律）。 */
const MAP_LIST_LIMIT = 3;

/** Python 解释器候选名（与 python.js 的 PYTHON_NAMES 同一批）。 */
const PYTHON_NAMES = ['python', 'python3', 'py'];

let snapshotCache;

/**
 * 同步探测本机能力（结果按进程缓存；`refresh: true` 重新探）。
 *
 * @param {object} [options] `config` = 运行期配置；`refresh` = 强制重探
 * @returns {object} 能力快照（纯数据，交给 renderCapabilityMap 渲染）
 */
export function syncCapabilities(options = {}) {
    if (snapshotCache !== undefined && options.refresh !== true) return snapshotCache;
    const config = options.config ?? {};
    const refresh = options.refresh === true ? { refresh: true } : {};

    const pdf = probePdfEngines(refresh);
    const tex = probeTexEngines(refresh);

    const render = PDF_RENDER_BINS.filter(([bin]) => pdf.bin[bin] !== undefined).map(([, label]) => label);
    const text = pdf.bin.pdftotext !== undefined ? ['pdftotext'] : [];
    // 与 tex.js 的 texEngines 同一口径：latexmk 只有在还有 xelatex / lualatex 时才算可用
    // （latexmk 自己是个调度器，没有引擎它也编不出东西）。
    const latex = [];
    if (tex.latexmk !== undefined && (tex.xelatex !== undefined || tex.lualatex !== undefined)) latex.push('latexmk');
    if (tex.xelatex !== undefined) latex.push('xelatex');
    if (tex.lualatex !== undefined) latex.push('lualatex');

    const pythonRequested = trimText(config?.python?.bin);
    // 设置里允许填**命令名**（`python`）也可以填路径：路径按 existsSync 判，
    // 命令名交给 PATH 查找 —— 只判 existsSync 会把 `python` 这种填法误报成「没有」。
    const python = pythonRequested !== ''
        ? (existsSync(pythonRequested) ? pythonRequested : findPdfExecutable([pythonRequested]))
        : findPdfExecutable(PYTHON_NAMES);

    snapshotCache = {
        pdf: { render, text },
        tex: latex,
        preview: previewKitAvailable() === true,
        python: python ?? null,
        // 音视频那半交给 av 通道自己的同步判据（probeAvSync）：地图与 office.av.check
        // 必须说同一件事 —— 只看 ffmpeg 会漏掉 ffprobe / 模型 / VAD / 运行时，
        // 而按 PATH 找 ffmpeg 又会漏掉 av.js 的常见安装目录（复核实测过这两个方向）。
        av: probeAvSync(config),
    };
    return snapshotCache;
}

/** 测试用：清掉快照缓存（改了环境变量或配置之后重探）。 */
export function resetCapabilities() {
    snapshotCache = undefined;
}

/** 字符串收敛（非字符串 / 空串一律当没填，与 config.js 的 text() 同一口径）。 */
function trimText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

/**
 * 把能力快照渲染成一段给模型看的文本（纯函数：同样输入必然同样输出，便于钉住字节）。
 *
 * 字节纪律：这一段会跟着 persona 进每一次请求的前缀，所以四行是**封顶**的形状 ——
 * 引擎清单每类最多三个（多了收成「等」），并且有 `CAPABILITY_MAP_BUDGET_BYTES` 的断言。
 *
 * @param {object|undefined} snapshot `syncCapabilities()` 的结果（undefined = 探测失败）
 * @param {object} [options] `subagents` = 当前组合里有没有子代理服务
 * @returns {string} 注入 persona 的那几行
 */
export function renderCapabilityMap(snapshot, options = {}) {
    if (!isProbed(snapshot)) {
        return '本会话能力：这一轮没探到（按 office_help / 各自 check 的输出为准，不要凭记忆承诺）。';
    }
    return [
        '本会话能力（插件现算，工具以本会话工具面为准）：',
        '- 引擎：' + engineSentence(snapshot),
        '- 检索派工：' + (options.subagents === true
            ? '有 subagents 服务，office_search_dispatch 走子代理。'
            : '没有 subagents 服务，office_search_dispatch 走插件内置检索通道（回执会写明）。'),
        '（只列探到的入口；能不能真跑看各自的 check。）',
    ].join('\n');
}

/**
 * 这一份是不是「探过」的快照。
 *
 * 只看 `typeof === 'object'` 会把 `{}` 当成「全都探不到」——那是**假陈述**：
 * 没探过与探到「没有」是两件事（前者要说「没探到」，后者才能逐项报缺什么）。
 */
function isProbed(snapshot) {
    return snapshot !== null && typeof snapshot === 'object'
        && (snapshot.pdf !== undefined || snapshot.tex !== undefined || snapshot.av !== undefined);
}

/** 引擎那半句：一段段可选的短语，全空时照实说「没探到」。 */
function engineSentence(snapshot) {
    const parts = [];
    const render = listOf(snapshot.pdf?.render);
    const text = listOf(snapshot.pdf?.text);
    if (render.length > 0) parts.push('PDF 渲染 ' + shorten(render));
    // 没有 CLI 引擎时**不能**断言「要装 PyMuPDF」：PyMuPDF 是 Python 包，同步探测看不到它
    // （要 import 一次才知道）。照实说「CLI 引擎没探到、PyMuPDF 去问 office.pdf.engines」。
    else if (hasPython(snapshot)) parts.push('PDF 渲染 未探到 CLI 引擎（PyMuPDF 由 office.pdf.engines 报）');
    else parts.push('PDF 渲染 无（装 poppler / PyMuPDF / MuPDF / Ghostscript 任一）');
    if (text.length > 0) parts.push('PDF 文字 ' + shorten(text));
    const tex = listOf(snapshot.tex);
    if (tex.length > 0) parts.push('LaTeX ' + shorten(tex));
    if (snapshot.preview === true) parts.push('预览 LibreOffice');
    if (hasPython(snapshot)) parts.push('Python ' + pythonLabel(snapshot.python));
    // 音视频那半用 av 通道自己的同步判据（`probeAvSync`）：**五样齐**才算可用
    // （ffmpeg / ffprobe / SenseVoice 权重 + 词表 / Silero VAD / sherpa-onnx 运行时），
    // 缺哪样由 office.av.check 报；关掉了就照实说关掉了。逐个列缺件会让字节跟着分支涨，
    // 而「缺什么」本来就该问 check。
    const av = snapshot.av ?? {};
    if (av.enabled === false) parts.push('音视频 已关闭（设置页 → 音频与视频）');
    else if (av.available === true) parts.push('音视频 ffmpeg + SenseVoice 模型');
    else if ([av.ffmpeg, av.ffprobe, av.model, av.vad, av.runtime].some((flag) => flag === true)) {
        parts.push('音视频 部分缺件（用 office.av.check 看缺什么）');
    }
    return parts.join('；') + '。';
}

/** 列表收敛：非数组当空。 */
function listOf(value) {
    return Array.isArray(value) ? value.filter((item) => typeof item === 'string' && item !== '') : [];
}

/** 每类最多列三个（地图封顶；再多的收成「等」）。 */
function shorten(items) {
    return items.length > MAP_LIST_LIMIT
        ? items.slice(0, MAP_LIST_LIMIT).join(' / ') + ' 等'
        : items.join(' / ');
}

function hasPython(snapshot) {
    return typeof snapshot.python === 'string' && snapshot.python !== '';
}

/** 解释器只报**文件名**（绝对路径又长又含本机目录，对模型没有用），并且截断到 16 字符
 * —— 地图有字节上界，未截断的文件名会让「最坏情况」没有上界。 */
function pythonLabel(path) {
    const name = basename(String(path)) || '在';
    return name.length > 16 ? name.slice(0, 15) + '…' : name;
}
