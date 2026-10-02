/**
 * 原生图表的**真实 Office 校验**（历史遗留 `11-6`）：用本机的 Excel / PowerPoint
 * 打开生成的 .xlsx / .pptx，并读出它们认到几个图表。
 *
 * 为什么值得单独一步：自研解析器与 LibreOffice 都宽容。本轮就抓到过一次 ——
 * 图例位置写成 `bottom`（ST_LegendPos 的合法值是 b/tr/l/r/t）时，自研解析器、
 * LibreOffice 都照常渲染，**真实 Excel 直接判文件坏**；同一份包里换成 Excel 自己写的
 * `chart1.xml` 立刻通过。只有真 Office 能证明「这份图表是能打开的图表」。
 *
 * 产物写在**系统临时目录**而不是工作区：实测本机通过 IDispatch 打开工作区路径下的
 * 文件时 Excel 会报「不能访问文件」（其它路径都正常），属于环境限制，不是文件问题。
 * 所以这里生成到 tmpdir 再校验，报告也写在那里。
 *
 * 跑法：node test/validate-chart-com.mjs
 * 环境：Windows + 装了 Microsoft Office + cscript 可用。缺任一项就 SKIP（exit 0），
 *       并把缺什么写清楚 —— 不把「跑不了」当成「通过」。
 */
import { existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { createEnv } from '../src/engine/kit.js';
import { resolveTheme } from '../src/engine/theme.js';
import { create as createExcel } from '../src/formats/excel.js';
import { create as createPpt } from '../src/formats/ppt.js';
import { runPdfProcess } from '../src/pdf.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        if (error?.skip === true) results.push({ name, ok: true, note: `SKIP ${error.message}` });
        else results.push({ name, ok: false, error });
    }
}
function skip(message) {
    const error = new Error(message);
    error.skip = true;
    throw error;
}

const root = join(tmpdir(), `office-chart-com-${process.pid}`);
rmSync(root, { recursive: true, force: true });
mkdirSync(root, { recursive: true });
const env = createEnv({ root, themeResolver: (id) => resolveTheme(id ?? 'business') });

const SERIES = { categories: ['Q1', 'Q2', 'Q3'], series: [{ name: '营收', values: [1280, 1530, 1410] }] };

/** 生成六种图表类型的 xlsx（每种一个文件，便于逐个定位失败）。 */
function makeWorkbooks() {
    for (const type of ['column', 'bar', 'line', 'pie', 'area']) {
        const wb = createExcel({ path: `chart-${type}.xlsx`, title: `图表 ${type}` }, env);
        const sheet = wb.sheet('明细');
        sheet.table({ columns: ['季度', '营收'], rows: [['Q1', 1280], ['Q2', 1530], ['Q3', 1410]] });
        sheet.chart({ type, title: `分季度营收（${type}）`, ...SERIES, labels: true });
        wb.save();
    }
    const scatter = createExcel({ path: 'chart-scatter.xlsx', title: '散点' }, env);
    const s = scatter.sheet('散点');
    s.table({ columns: ['x', 'y'], rows: [[1, 2], [2, 4], [3, 9]] });
    s.chart({ type: 'scatter', title: '样本散点', series: [{ name: '样本', x: [1, 2, 3], values: [2, 4, 9] }] });
    scatter.save();
    // 两个图表挤在同一张表上（多图表编号 + 多 drawing rel）
    const multi = createExcel({ path: 'chart-multi.xlsx', title: '多图' }, env);
    const m = multi.sheet('多图');
    m.table({ columns: ['季度', '营收'], rows: [['Q1', 1280], ['Q2', 1530], ['Q3', 1410]] });
    m.chart({ type: 'column', title: '柱形', ...SERIES, at: { col: 4, row: 1 } });
    m.chart({ type: 'line', title: '折线', ...SERIES, at: { col: 4, row: 18 } });
    multi.save();
}

function makeDecks() {
    const deck = createPpt({ path: 'chart-deck.pptx', title: '原生图表' }, env);
    deck.cover({ title: '原生图表' });
    deck.chart({ type: 'column', title: '分季度营收', ...SERIES, labels: true });
    deck.chart({ type: 'pie', title: '渠道占比', categories: ['线上', '线下'], series: [{ name: '占比', values: [62, 38] }] });
    deck.bullets({ title: '小结', items: ['图表是可编辑对象', '数据内联在图表里'] });
    deck.save();
}

/** 跑一次 cscript validate-com.vbs，返回报告文本。 */
async function runComValidator() {
    const report = join(root, 'com-report.txt');
    // vbs 的路径按**本文件的位置**算，不按 cwd：按 cwd 算时从工作区根跑会找不到脚本，
    // cscript 退出 1、报告不生成，而这一步会被记成 SKIP —— 一个「0 份 checked 也全绿」
    // 的假绿（复核抓到的）。
    const vbs = fileURLToPath(new URL('./validate-com.vbs', import.meta.url));
    const run = await runPdfProcess('cscript.exe', ['//nologo', vbs, root, report], {
        outPath: join(root, 'cscript.out.txt'),
        errPath: join(root, 'cscript.err.txt'),
        timeoutMs: 600_000,
        cwd: process.cwd(),
    });
    if (!existsSync(report)) {
        throw Object.assign(new Error(`cscript 没写出报告（code=${run.code}）：${(run.err ?? '').slice(-200)}`), { skip: true });
    }
    const { readFileSync } = await import('node:fs');
    const text = readFileSync(report, 'utf8');
    writeFileSync(join(root, 'com-report.copy.txt'), text);
    return { text, code: run.code };
}

let report;
await check('生成图表产物（六种类型 + 一张表两个图表 + 一份图表 PPT）', async () => {
    makeWorkbooks();
    makeDecks();
    const files = ['chart-column.xlsx', 'chart-bar.xlsx', 'chart-line.xlsx', 'chart-pie.xlsx',
        'chart-area.xlsx', 'chart-scatter.xlsx', 'chart-multi.xlsx', 'chart-deck.pptx'];
    for (const name of files) {
        if (!existsSync(join(root, name))) throw new Error(`${name} 没生成出来`);
    }
    return `${files.length} 份`;
});

await check('真实 Office：Excel 打开六个图表工作簿并认出图表数', async () => {
    const run = await runComValidator();
    report = run.text;
    if (/com-validate: 0 passed \/ 0 failed/.test(report)) skip('cscript 没有校验任何文件（可能没装 Office）');
    const lines = report.split(/\r?\n/).filter((line) => line.startsWith('OK') || line.startsWith('FAIL'));
    if (lines.length === 0) skip(`报告里没有逐文件结论：${report.slice(0, 120)}`);
    const failures = lines.filter((line) => line.startsWith('FAIL'));
    assert.equal(failures.length, 0, `真实 Office 打不开这些文件：\n${failures.join('\n')}`);
    for (const line of lines.filter((item) => item.includes('.xlsx'))) {
        assert.match(line, /charts=[1-9]/, `Excel 没认出图表：${line}`);
    }
    return lines.filter((item) => item.includes('.xlsx')).length + ' 份工作簿';
});

await check('真实 Office：PowerPoint 打开图表 PPT 并认出 chartShapes', async () => {
    if (report === undefined) skip('上一步没跑成');
    const line = report.split(/\r?\n/).find((item) => item.includes('.pptx'));
    assert.ok(line !== undefined, `报告里没有 pptx 的结论：${report.slice(0, 160)}`);
    assert.ok(line.startsWith('OK'), `PowerPoint 打不开：${line}`);
    assert.match(line, /chartShapes=[1-9]/, `PowerPoint 没认出图表形状：${line}`);
    return line.trim();
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (report !== undefined) for (const line of report.split(/\r?\n/).filter((item) => item.trim() !== '')) console.log(`  | ${line}`);
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`validate-chart-com: ${results.length - failed.length}/${results.length} 通过（产物在 ${root}）`);
process.exit(failed.length === 0 ? 0 : 1);
