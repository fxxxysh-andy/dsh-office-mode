/**
 * 参考稿级验收样例：按 DeepSeek 内部研讨稿（86 页 / 16:9）的页面骨架复刻一整套演示稿。
 *
 * 它不跑断言，只产出可打开的 pptx 供人眼与真实 PowerPoint 复核：
 *   封面（满页照片 + 压遮罩 + 校徽）→ 目录（2×2 大号编号）→ 章节页 → 卡片网格 →
 *   对比双栏 → 编号流程 → 合并单元格表格 → 图+文 → KPI → 时间线 → 图标网格 → 结束页。
 *
 * 素材由脚本即时生成到 <仓库根>/.office/tmp/ref-assets/，不往交付目录塞资源文件。
 * 用法：node test/make-reference-deck.mjs [输出目录]
 */
import { mkdirSync, statSync, writeFileSync } from 'node:fs';
import { deflateSync } from 'node:zlib';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { resolveConfig } from '../src/config.js';
import { executeRun } from '../src/run.js';
import { renderRun } from '../src/tools.js';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = resolve(process.argv[2] ?? resolve(here, '办公模式示例'));
mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------- 素材

const assetDir = fileURLToPath(new URL('../../../.office/tmp/ref-assets/', import.meta.url));
mkdirSync(assetDir, { recursive: true });

/** 生成 RGBA PNG。painter(x, y, w, h) 返回 [r,g,b,a]。 */
function makePng(width, height, painter) {
    const table = Array.from({ length: 256 }, (_, n) => {
        let c = n;
        for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        return c >>> 0;
    });
    const crc32 = (buf) => {
        let c = 0xffffffff;
        for (const byte of buf) c = table[(c ^ byte) & 0xff] ^ (c >>> 8);
        return (c ^ 0xffffffff) >>> 0;
    };
    const chunk = (type, data) => {
        const len = Buffer.alloc(4);
        len.writeUInt32BE(data.length, 0);
        const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
        const crc = Buffer.alloc(4);
        crc.writeUInt32BE(crc32(body), 0);
        return Buffer.concat([len, body, crc]);
    };
    const stride = width * 4 + 1;
    const raw = Buffer.alloc(stride * height);
    for (let y = 0; y < height; y += 1) {
        const rowStart = y * stride;
        raw[rowStart] = 0;
        for (let x = 0; x < width; x += 1) {
            const [r, g, b, a] = painter(x, y, width, height);
            const at = rowStart + 1 + x * 4;
            raw[at] = r; raw[at + 1] = g; raw[at + 2] = b; raw[at + 3] = a === undefined ? 255 : a;
        }
    }
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8; ihdr[9] = 6;
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw, { level: 9 })),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

/** 封面背景：斜向暖红渐变 + 底部建筑剪影 + 一轮浅色月亮。 */
const bgCover = resolve(assetDir, '封面背景.png');
writeFileSync(bgCover, makePng(1600, 900, (x, y, w, h) => {
    const t = (x / w) * 0.55 + (y / h) * 0.45;
    let r = 118 + 78 * t;
    let g = 26 + 34 * t;
    let b = 26 + 34 * t;
    if (y > h * 0.60) {
        const shade = 0.42 + 0.22 * Math.abs(Math.sin((x / w) * Math.PI * 3.1));
        r *= shade; g *= shade; b *= shade;
    }
    const dx = x - w * 0.79;
    const dy = y - h * 0.27;
    if (Math.sqrt(dx * dx + dy * dy) < h * 0.055) { r = 246; g = 236; b = 214; }
    const grain = ((x * 7 + y * 13) % 11) - 5;
    return [Math.max(0, Math.min(255, r + grain)), Math.max(0, Math.min(255, g + grain)), Math.max(0, Math.min(255, b + grain))];
}));

/** 内容页背景：近白纹理（参考稿每页都有一层纸纹）。 */
const bgContent = resolve(assetDir, '内容背景.png');
writeFileSync(bgContent, makePng(1600, 900, (x, y, w, h) => {
    const wave = Math.sin((x / w) * Math.PI * 5) * Math.sin((y / h) * Math.PI * 4) * 3;
    const edge = 250 - Math.abs((x / w) - 0.5) * 8;
    const v = Math.round(edge + wave);
    return [v, v - 2, v - 5];
}));

/** 校徽：深红圆底 + 白色圆环 + 中心留白。 */
const logo = resolve(assetDir, '校徽.png');
writeFileSync(logo, makePng(256, 256, (x, y, w) => {
    const dx = x - w / 2;
    const dy = y - w / 2;
    const d = Math.sqrt(dx * dx + dy * dy) / (w / 2);
    if (d > 0.97) return [0, 0, 0, 0];
    if (d > 0.80) return [139, 30, 30, 255];
    if (d > 0.70) return [255, 255, 255, 255];
    if (d > 0.26) return [139, 30, 30, 255];
    return [255, 255, 255, 255];
}));

/** 截图：浅色卡片 + 柱状图（模拟一张真实界面截图）。 */
function chartPng(file, bars, accent) {
    writeFileSync(file, makePng(1200, 700, (x, y, w, h) => {
        let r = 250; let g = 250; let b = 252;
        if (y < 64) { r = accent[0]; g = accent[1]; b = accent[2]; }
        if (x % 60 === 0 || y % 60 === 0) { r -= 8; g -= 8; b -= 6; }
        for (const bar of bars) {
            const top = h - 90 - bar.h;
            if (x >= bar.x && x < bar.x + bar.w && y >= top && y < h - 90) {
                const k = (h - 90 - y) / bar.h;
                r = Math.round(accent[0] + (255 - accent[0]) * 0.35 * (1 - k));
                g = Math.round(accent[1] + (255 - accent[1]) * 0.35 * (1 - k));
                b = Math.round(accent[2] + (255 - accent[2]) * 0.35 * (1 - k));
            }
        }
        if (y > h - 90) { r = 232; g = 232; b = 236; }
        return [r, g, b];
    }));
}
const shotSchedule = resolve(assetDir, '日程截图.png');
chartPng(shotSchedule, [
    { x: 120, w: 90, h: 240 }, { x: 280, w: 90, h: 360 }, { x: 440, w: 90, h: 180 },
    { x: 600, w: 90, h: 420 }, { x: 760, w: 90, h: 300 }, { x: 920, w: 90, h: 480 },
], [139, 30, 30]);
const shotChannel = resolve(assetDir, '渠道截图.png');
chartPng(shotChannel, [
    { x: 140, w: 110, h: 300 }, { x: 360, w: 110, h: 430 }, { x: 580, w: 110, h: 210 },
    { x: 800, w: 110, h: 360 }, { x: 1000, w: 110, h: 150 },
], [31, 78, 121]);

// ---------------------------------------------------------------- 脚本

const ICON = {
    globe: '\\uE774', page: '\\uE7C3', doc: '\\uE8A5', home: '\\uE80F',
    gear: '\\uE713', search: '\\uE721', edit: '\\uE70F', save: '\\uE74E',
};

const SCRIPT = `
const out = '.';
const deck = office.ppt.create({
    title: '提示词工程和落地场景',
    theme: 'academic',
    path: '参考稿复刻.pptx',
    author: 'AI 肖睿团队',
});

// 全篇母版：内容页纸纹背景 + 右上校徽 + 页眉 + 页码 + 左上装饰条
deck.master({
    background: { image: { path: ${JSON.stringify(bgContent)}, fit: 'cover', dim: 0.05, dimColor: 'FFFFFF' } },
    logo: { path: ${JSON.stringify(logo)}, corner: 'top-right', widthCm: 1.5, marginCm: 0.45 },
    header: { text: '提示词工程和落地场景', size: 15, color: '8B1E1E', align: 'left' },
    pageNumber: { show: true, corner: 'bottom-right', from: 2, size: 10, color: '8B1E1E' },
    accent: [{ bar: 'left-top', color: '8B1E1E', sizeCm: 0.12, lengthCm: 4.4 }],
});

// P1 封面：满页照片 + 红色压遮罩 + 大标题
deck.cover({
    title: '提示词工程和落地场景',
    subtitle: 'DeepSeek 内部研讨系列',
    kicker: '内部研讨系列',
    presenter: 'AI 肖睿团队（韩露、吴寒、孙萍、李娜、刘誉）',
    date: '2026 年 3 月',
});
deck.background({ image: { path: ${JSON.stringify(bgCover)}, fit: 'cover', dim: 0.52, dimColor: '8B1E1E' } });

// P2 目录：2×2 大号编号
deck.kpi({
    title: '目录',
    columns: 2,
    items: [
        { value: '01', label: 'DeepSeek 火爆的原因分析' },
        { value: '02', label: '直接使用 DeepSeek 的三种方法' },
        { value: '03', label: 'DeepSeek 提示词技巧' },
        { value: '04', label: 'DeepSeek 常见应用场景' },
    ],
});

// P3 章节页
deck.section({ title: 'DeepSeek 火爆的原因分析', subtitle: 'Analysis of the Reasons for DeepSeek\\u2019s Popularity' });

// P4 卡片网格：能力突破 / 开源 / 低成本 / 国产化
deck.cards({
    title: '为什么火：能力突破、开源、低成本、国产化',
    columns: 2,
    items: [
        { title: '基础能力', body: '进入推理模型阶段，综合性能跻身全球第一梯队。', icon: '${ICON.page}', accent: '8B1E1E' },
        { title: '开源开放', body: '训练代码、数据清洗工具与微调框架全量开源。', icon: '${ICON.doc}', accent: '1F4E79' },
        { title: '超低成本', body: '训练成本仅 557 万美元，推理成本降低 83%。', icon: '${ICON.save}', accent: '1F5C3A' },
        { title: '国产自主', body: '把代际差距从 3–5 年缩短到 3–5 个月。', icon: '${ICON.home}', accent: 'B45309' },
    ],
});

// P5 对比双栏：传统写法 vs DeepSeek 写法
deck.compare({
    title: '提示词技巧：真诚 + 直接',
    left: {
        title: '传统写法',
        accent: '808080',
        items: ['你现在是一个新能源汽车的市场研究分析师……', '请按周报的格式帮我完成并进行润色。', '不少于 500 字。'],
    },
    right: {
        title: 'DeepSeek 写法',
        accent: '8B1E1E',
        items: ['帮我把这份报告包装一下。', '我要写成周报给老板看。', '老板很看重数据。'],
    },
});

// P6 编号流程：任务 → 背景 → 目标 → 负面限定
deck.steps({
    title: '通用公式',
    items: [
        { title: '任务', body: '做什么' },
        { title: '背景', body: '给谁用' },
        { title: '目标', body: '期望效果' },
        { title: '负面限定', body: '担心的问题' },
    ],
});

// P7 合并单元格表格
deck.table({
    title: '蒸馏模型与满血版',
    columns: ['模型', '基座模型', '类型'],
    headerFill: '8B1E1E',
    rows: [
        ['DeepSeek-R1-Distill-Qwen-1.5B', 'Qwen2.5-Math-1.5B', { text: '蒸馏模型', rowSpan: 4 }],
        ['DeepSeek-R1-Distill-Qwen-7B', 'Qwen2.5-Math-7B'],
        ['DeepSeek-R1-Distill-Llama-8B', 'Llama-3.1-8B'],
        ['DeepSeek-R1-Distill-Qwen-32B', 'Qwen2.5-32B'],
        [{ text: 'DeepSeek-R1-671B', bold: true }, { text: 'DeepSeek-V3-Base', bold: true }, { text: '满血版', bold: true }],
    ],
});

// P8 图 + 文
deck.imageText({
    title: '通用公式的输出：日程安排清晰明确',
    side: 'right',
    image: { path: ${JSON.stringify(shotSchedule)}, fit: 'contain', caption: '由模型生成的 20 天行程' },
    items: [
        '贴心准备清单，细致到每日 5000 步',
        'Day1–5 东京：浅草寺、晴空塔、台场',
        'Day6–7 箱根：温泉与富士山',
        'Day8–13 京都：清水寺、岚山、伏见稻荷',
        'Day14–17 大阪：环球影城、道顿堀',
    ],
});

// P9 KPI 大数字
deck.kpi({
    title: '关键数字',
    items: [
        { value: '108', unit: '%', label: '季度达成率' },
        { value: '32', label: '新增客户' },
        { value: '17', unit: '天', label: '平均交付周期' },
        { value: '83', unit: '%', label: '推理成本降幅' },
    ],
});

// P10 时间线
deck.timeline({
    title: '落地路线',
    items: [
        { time: '第一步', title: '明确需求', body: '把要做的事写成一句话' },
        { time: '第二步', title: '设计提示词', body: '任务 + 背景 + 目标 + 限定' },
        { time: '第三步', title: '生成初稿', body: '大纲 → 成稿 → 复核' },
        { time: '第四步', title: '细节调整', body: '选主题、改内容、控风险' },
    ],
});

// P11 图标网格
deck.iconGrid({
    title: '直接使用 DeepSeek 的三种方法',
    columns: 3,
    items: [
        { icon: '${ICON.globe}', label: '官方网页与 APP', body: 'chat.deepseek.com' },
        { icon: '${ICON.gear}', label: '第三方通道', body: '硅基流动 / 火山引擎 / 阿里百炼' },
        { icon: '${ICON.home}', label: '私有化部署', body: 'Ollama / vLLM' },
    ],
});

// P12 结束页：满页照片 + 白色压遮罩 + 居中大字
deck.closing({ title: '感谢各位老师同学批评与指导', subtitle: 'AI 肖睿团队' });
deck.background({ image: { path: ${JSON.stringify(bgCover)}, fit: 'cover', dim: 0.72, dimColor: 'FFFFFF' } });

deck.save(out + '/参考稿复刻.pptx');
office.log('参考稿复刻：12 页');
return { slides: 12, theme: 'academic' };
`;

const result = await executeRun(
    { script: SCRIPT, purpose: '复刻参考稿页面骨架' },
    { agent: { session: { header: { cwd: outDir } } } },
    resolveConfig({}),
);

console.log(renderRun(result));
console.log('');
if (result.ok !== true) process.exit(1);
const file = resolve(outDir, '参考稿复刻.pptx');
console.log(`产物：${file}  ${statSync(file).size} 字节`);
