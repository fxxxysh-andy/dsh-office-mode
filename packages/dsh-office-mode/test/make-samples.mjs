/**
 * 生成一套可以直接打开的办公模式样例（Word / Excel / PPT）。
 *
 * 和端到端测试不同，它不跑断言，只把产物落到指定目录，用来肉眼验收排版与主题。
 * 默认生成两组：不带主题的「素色网格」与显式 business 主题的彩色版 ——
 * 前者是现在的默认观感（无背景填充），后者用来看主题模板还在不在。
 *
 * 样例刻意覆盖本轮新增的能力：Excel 统计块与分组汇总、Word 公式与合并单元格表格、
 * PPT 一页多图。图片素材由脚本即时生成到 <仓库根>/.office/tmp/sample-assets/，
 * 不往交付目录里塞资源文件。
 *
 * 用法：node test/make-samples.mjs [输出目录]       默认 ./办公模式示例
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

const assetDir = fileURLToPath(new URL('../../../.office/tmp/sample-assets/', import.meta.url));
mkdirSync(assetDir, { recursive: true });

/** 生成一张纯色/柱状 PNG（真魔数，不是占位文件）。 */
function makePng(width, height, bars) {
    const crcTable = Array.from({ length: 256 }, (_, n) => {
        let c = n;
        for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        return c >>> 0;
    });
    const crc32 = (buf) => {
        let c = 0xffffffff;
        for (const byte of buf) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
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
    const raw = Buffer.alloc((width * 3 + 1) * height);
    for (let y = 0; y < height; y += 1) {
        const rowStart = y * (width * 3 + 1);
        raw[rowStart] = 0;
        for (let x = 0; x < width; x += 1) {
            // 背景：浅灰网格；柱：主题蓝渐变，用来模拟一张真实截图
            let r = 246; let g = 248; let b = 251;
            if (x % 40 === 0 || y % 40 === 0) { r = 226; g = 232; b = 240; }
            for (const bar of bars) {
                if (x >= bar.x && x < bar.x + bar.w && y >= height - bar.h && y < height - 30) {
                    const t = (height - y) / bar.h;
                    r = Math.round(31 + 60 * t); g = Math.round(78 + 90 * t); b = Math.round(121 + 80 * t);
                }
            }
            const at = rowStart + 1 + x * 3;
            raw[at] = r; raw[at + 1] = g; raw[at + 2] = b;
        }
    }
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8; ihdr[9] = 2;
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw, { level: 9 })),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

const assetChannel = resolve(assetDir, '渠道分布.png');
const assetTrend = resolve(assetDir, '月度趋势.png');
writeFileSync(assetChannel, makePng(960, 540, [
    { x: 120, w: 90, h: 260 }, { x: 300, w: 90, h: 380 }, { x: 480, w: 90, h: 180 },
    { x: 660, w: 90, h: 300 }, { x: 800, w: 90, h: 140 },
]));
writeFileSync(assetTrend, makePng(960, 540, [
    { x: 100, w: 70, h: 120 }, { x: 220, w: 70, h: 200 }, { x: 340, w: 70, h: 170 },
    { x: 460, w: 70, h: 280 }, { x: 580, w: 70, h: 240 }, { x: 700, w: 70, h: 360 },
]));

// ---------------------------------------------------------------- 样例脚本

/** 同一套内容跑两遍：默认（plain 素色）与显式 business。 */
const SETS = [
    { suffix: '', theme: null, label: '默认（plain 素色网格，无背景填充）' },
    { suffix: '-商务蓝', theme: 'business', label: 'business 主题（表头底色 + 斑马纹）' },
];

/** 达成率公式：JSON.stringify 负责把反斜杠原样写进生成的脚本。 */
const FORMULA = '\\text{达成率} = \\frac{完成（万元）}{目标（万元）} \\times 100\\%';

function scriptFor({ suffix, theme }) {
    const themeLine = theme === null ? '' : `const theme = '${theme}';`;
    const themed = theme === null ? '' : ', theme';
    return `
const out = '.';
${themeLine}

const w = office.word.create({ title: '2026 年第一季度经营回顾'${themed}, author: '运营中心' });
w.title('2026 年第一季度经营回顾')
 .para('编制：运营中心　　数据截止：2026 年 3 月 31 日')
 .heading('一、整体情况', 1)
 .para('本季度整体达成率 108%，超额完成目标。三条业务线均实现正增长，企业服务线增长最快。')
 .bullets(['营收 3,860 万元，同比增长 18.4%', '新增客户 32 家，其中付费客户 21 家', '平均交付周期由 21 天缩短到 17 天'])
 .heading('二、分月数据', 1)
 .para('达成率口径如下：')
 .formula(${JSON.stringify(FORMULA)}, { number: '（1）' })
 .table({
     columns: ['月份', { title: '目标（万元）', type: 'number' }, { title: '完成（万元）', type: 'number' }, '达成率'],
     rows: [
         ['1 月', 1200, 1302, '108.5%'],
         ['2 月', 1100, 1144, '104.0%'],
         ['3 月', 1300, 1414, '108.8%'],
         [{ text: '合计', colspan: 2, align: 'right', bold: true }, 3860, '107.2%'],
     ],
     caption: '表 1 分月达成情况',
 })
 .heading('三、下季度计划', 1)
 .steps(['把企业服务线的交付模板标准化', '六月前完成客户成功团队扩编', '试点按季度结算的付费方式'])
 .quote('把交付做成标准件，增长才不依赖人手。')
 .save(out + '/季度回顾${suffix}.docx');

const x = office.excel.create({${themed === '' ? '' : ` theme: '${theme}',`} title: '2026 年第一季度分月明细' });
const s = x.sheet('分月明细');
s.title('2026 年第一季度分月明细', { span: 5 })
 .table({
     columns: ['月份', '目标（万元）', '完成（万元）', '差额（万元）', '达成率'],
     rows: [['1 月', 1200, 1302, '', '108.5%'], ['2 月', 1100, 1144, '', '104.0%'], ['3 月', 1300, 1414, '', '108.8%']],
     totalRow: true,
     freeze: true,
 })
 .note('差额 = 完成 − 目标；达成率 = 完成 ÷ 目标。');
// 统计块：真公式 + 缓存值，任何阅读器不重算也能看到数字
s.stats('C3:C5', { label: '完成额统计', layout: 'columns' });

// 分组汇总：按部门求和/计数/平均，替代轻量透视
const g = x.sheet('部门汇总');
g.title('部门投入汇总', { span: 3 })
 .table({
     columns: ['部门', { title: '金额（万元）', type: 'number' }, '备注'],
     rows: [
         ['市场部', 860, '含活动投放'],
         ['交付部', 1240, ''],
         ['研发部', 1520, ''],
         ['市场部', 320, '补充预算'],
         ['交付部', 180, ''],
     ],
 })
 .summary({ key: '部门', value: '金额（万元）', total: true, label: '按部门汇总' });
x.save(out + '/分月明细${suffix}.xlsx');

const p = office.ppt.create({ title: '2026 Q1 经营回顾'${themed} });
p.cover({ title: '2026 年第一季度经营回顾', subtitle: '运营中心 · 2026 年 4 月' })
 .bullets({ title: '本季度三件事', items: ['整体达成率 108%，超额完成', '新增付费客户 21 家', '交付周期缩短 4 天'] })
 .table({
     title: '分月达成情况',
     columns: ['月份', '目标', '完成', '达成率'],
     rows: [['1 月', '1,200', '1,302', '108.5%'], ['2 月', '1,100', '1,144', '104.0%'], ['3 月', '1,300', '1,414', '108.8%']],
 })
 .images([
     { path: ${JSON.stringify(assetChannel)}, caption: '渠道分布' },
     { path: ${JSON.stringify(assetTrend)}, caption: '月度趋势' },
 ], { title: '数据截图', columns: 2 })
 .quote({ text: '把交付做成标准件，增长才不依赖人手。' })
 .section({ title: '下季度计划' })
 .bullets({ title: '三件要事', items: ['交付模板标准化', '客户成功团队扩编', '试点季度结算'] })
 .closing({ title: '谢谢', subtitle: '运营中心' })
 .save(out + '/经营回顾${suffix}.pptx');

office.log('三件套已生成：${theme ?? 'plain'}');
return { theme: '${theme ?? 'plain'}', formats: office.formats() };
`;
}

// ---------------------------------------------------------------- 执行

let failed = false;
for (const set of SETS) {
    const result = await executeRun(
        { script: scriptFor(set), purpose: `生成办公模式样例（${set.label}）` },
        { agent: { session: { header: { cwd: outDir } } } },
        resolveConfig({}),
    );
    console.log(`── ${set.label}`);
    console.log(renderRun(result));
    console.log('');
    if (result.ok !== true) failed = true;
}

console.log(`样例目录：${outDir}`);
for (const set of SETS) {
    for (const name of ['季度回顾', '分月明细', '经营回顾']) {
        const ext = { 季度回顾: '.docx', 分月明细: '.xlsx', 经营回顾: '.pptx' }[name];
        const file = resolve(outDir, `${name}${set.suffix}${ext}`);
        console.log(`  ${name}${set.suffix}${ext}  ${statSync(file).size} 字节`);
    }
}
if (failed) process.exit(1);
