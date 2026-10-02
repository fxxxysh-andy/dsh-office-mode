/**
 * 原生图表部件（历史遗留 `11-6`）的测试。
 *
 * 这条账说的是「Excel / PPT 没有原生图表，Python 画图只是替代路径」。本轮的落点是
 * 真正的 DrawingML 图表部件：Excel 的 `xl/charts/chartN.xml` + `xl/drawings/drawingN.xml`，
 * PPT 的 `ppt/charts/chartN.xml` + `p:graphicFrame`。两份共用同一份 `c:chartSpace`
 * （见 src/formats/chart.js）。
 *
 * 四段：
 *   1. 规格校验：类型 / 系列 / 长度不一致都要**当场报错**，不画一半；
 *   2. 部件结构：schema 顺序、枚举取值（`legendPos` 只能是 b/tr/l/r/t —— 写 `bottom`
 *      时自研解析器与 LibreOffice 都照常，**真实 Excel 判文件坏**，这条是回归点）、
 *      饼图无坐标轴、堆叠出 overlap；
 *   3. Excel 全链路：部件 / rels / [Content_Types] / 工作表锚点 / 读回 / 就地编辑保留；
 *   4. PPT 全链路：部件 / 幻灯片 rels / graphicFrame / 读回（stats.charts 与 layout）。
 *
 * 真 Office 校验（Excel / PowerPoint 打开并报出图表数）在 `validate-chart-com.mjs`，
 * 那一步要本机装了 Office，按可选套件跑。
 *
 * 跑法：node test/chart.mjs
 */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { createEnv } from '../src/engine/kit.js';
import { unzipText } from '../src/engine/zip.js';
import { resolveTheme } from '../src/engine/theme.js';
import { CHART_TYPES, chartSpaceXml, describeChartSpace, normalizeChartSpec, spreadsheetDrawingXml } from '../src/formats/chart.js';
import { create as createExcel, edit as editExcel, read as readExcel } from '../src/formats/excel.js';
import { create as createPpt, read as readPpt } from '../src/formats/ppt.js';

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

const root = mkdtempSync(join(tmpdir(), 'office-chart-'));
const env = createEnv({ root, themeResolver: (id) => resolveTheme(id ?? 'business') });

const BASE = {
    type: 'column',
    title: '分季度营收',
    categories: ['Q1', 'Q2', 'Q3'],
    series: [{ name: '营收', values: [1280, 1530, 1410] }],
};

/** 元素出现顺序：schema 的 sequence 顺序错了真实 Office 会判坏。 */
function orderOf(xml, needles) {
    return needles.map((needle) => {
        const at = xml.indexOf(needle);
        assert.ok(at >= 0, `chart XML 里缺 ${needle}`);
        return at;
    });
}
function assertOrder(xml, needles, label) {
    const positions = orderOf(xml, needles);
    for (let i = 1; i < positions.length; i += 1) {
        assert.ok(positions[i - 1] < positions[i], `${label} 顺序不对：${needles[i - 1]} 应在 ${needles[i]} 之前`);
    }
}

// ── 1. 规格校验 ─────────────────────────────────────────────────────────────

await check('normalizeChartSpec：类型 / 系列 / 长度不一致当场报错', async () => {
    assert.throws(() => normalizeChartSpec({ series: [] }), /至少要有一个系列/);
    assert.throws(() => normalizeChartSpec({ type: 'radar', series: [{ values: [1] }] }), /type 只支持 bar \/ column/);
    assert.throws(() => normalizeChartSpec({ series: [{ values: [1, 'x'] }] }), /有非数字/);
    assert.throws(
        () => normalizeChartSpec({ categories: ['a', 'b'], series: [{ values: [1, 2, 3] }] }),
        /数量必须一致/,
    );
    assert.throws(() => normalizeChartSpec({ type: 'scatter', series: [{ values: [1, 2], x: [1] }] }), /x 与 y 必须一一对应/);
    assert.throws(() => normalizeChartSpec({ ...BASE, legend: 'botom' }), /legend 只支持/);
    // 不给 categories 时按**最长**系列定宽（提示里说的「按最长系列推」要真的做到）
    assert.throws(() => normalizeChartSpec({ series: [{ values: [1] }, { values: [1, 2, 3] }] }), /按最长系列/);
    assert.equal(normalizeChartSpec({ series: [{ values: [1, 2, 3] }, { values: [3, 2, 1] }] }).categories.length, 3);
    const ok = normalizeChartSpec(BASE);
    assert.equal(ok.type, 'column');
    assert.equal(ok.legend, 'bottom');
    assert.equal(ok.gapWidth, 150);
    assert.equal(ok.palette.length, 8);
    // 不给 categories 时按最长系列补 1..n
    assert.deepEqual(normalizeChartSpec({ series: [{ values: [5, 6] }] }).categories, ['1', '2']);
    // legend:false 与 'none' 等价
    assert.equal(normalizeChartSpec({ ...BASE, legend: false }).legend, 'none');
    return CHART_TYPES.join('/');
});

// ── 2. 部件结构 ─────────────────────────────────────────────────────────────

await check('chartSpaceXml：schema 顺序与必需元素', async () => {
    const xml = chartSpaceXml(BASE);
    assert.ok(xml.startsWith('<?xml'), '要有 XML 声明');
    assert.match(xml, /<c:chartSpace xmlns:c="[^"]+" xmlns:a="[^"]+" xmlns:r="[^"]+">/);
    assertOrder(xml, ['<c:title>', '<c:autoTitleDeleted', '<c:plotArea>', '<c:legend>', '<c:plotVisOnly'], 'CT_Chart');
    assertOrder(xml, ['<c:layout/>', '<c:barChart>', '<c:catAx>', '<c:valAx>'], 'CT_PlotArea');
    assertOrder(xml, ['<c:barDir', '<c:grouping', '<c:varyColors', '<c:ser>', '<c:gapWidth', '<c:axId'], 'CT_BarChart');
    assertOrder(xml, ['<c:idx val="0"/>', '<c:order val="0"/>', '<c:tx><c:v>', '<c:cat>', '<c:val>'], 'CT_BarSer');
    assertOrder(xml, ['<c:chart>', '<c:printSettings>'], 'CT_ChartSpace');
    // 数据内联：不引用单元格、不嵌工作簿
    assert.match(xml, /<c:strLit><c:ptCount val="3"\/>/);
    assert.match(xml, /<c:numLit><c:ptCount val="3"\/>/);
    assert.ok(!xml.includes('<c:numRef>') && !xml.includes('externalData'), '内联数据不该出现引用或外部工作簿');
    return `${xml.length} 字符`;
});

await check('回归点：legendPos 只能是 ST_LegendPos 的枚举值（写 bottom 会被真实 Excel 拒收）', async () => {
    const allowed = new Set(['b', 'tr', 'l', 'r', 't']);
    for (const legend of ['bottom', 'right', 'left', 'top']) {
        const value = /<c:legendPos val="([^"]+)"\/>/.exec(chartSpaceXml({ ...BASE, legend }))?.[1];
        assert.ok(allowed.has(value), `legend=${legend} 出来的是 ${value}，不在 ${[...allowed].join('/')} 里`);
    }
    assert.equal(/<c:legendPos val="([^"]+)"\/>/.exec(chartSpaceXml({ ...BASE, legend: 'right' }))[1], 'r');
    assert.ok(!chartSpaceXml({ ...BASE, legend: 'none' }).includes('<c:legend>'), 'legend:none 不该有图例');
    return 'b/tr/l/r/t';
});

await check('各图表类型：饼图无坐标轴、堆叠出 overlap、标签出 dLbls、散点用 xVal/yVal', async () => {
    const pie = chartSpaceXml({ ...BASE, type: 'pie' });
    assert.match(pie, /<c:pieChart><c:varyColors val="1"\/>/);
    assert.ok(!pie.includes('<c:catAx>') && !pie.includes('<c:valAx>'), '饼图不该有坐标轴');
    assert.ok(!pie.includes('<c:axId'), '饼图不该有 axId');

    const stacked = chartSpaceXml({ ...BASE, stacked: true });
    assert.match(stacked, /<c:grouping val="stacked"\/>/);
    assert.match(stacked, /<c:overlap val="100"\/>/);

    const labelled = chartSpaceXml({ ...BASE, labels: true });
    assert.match(labelled, /<c:dLbls><c:showLegendKey val="0"\/><c:showVal val="1"\/>/);
    // CT_*Ser 的 sequence：dLbls 必须在 cat / val **之前**（放后面真实 Excel 也认，
    // 但那是 schema 违规，严格校验器会拒收）。
    assertOrder(labelled, ['<c:tx><c:v>', '<c:dLbls>', '<c:cat>', '<c:val>'], 'CT_BarSer + dLbls');
    const scatterLabelled = chartSpaceXml({ type: 'scatter', labels: true, series: [{ name: '样本', x: [1, 2], values: [2, 4] }] });
    assertOrder(scatterLabelled, ['<c:marker>', '<c:dLbls>', '<c:xVal>', '<c:yVal>'], 'CT_ScatterSer + dLbls');

    const bar = chartSpaceXml({ ...BASE, type: 'bar' });
    assert.match(bar, /<c:barDir val="bar"\/>/, 'bar 是横向条形');

    const line = chartSpaceXml({ ...BASE, type: 'line' });
    assertOrder(line, ['<c:lineChart>', '<c:marker val="1"/>', '<c:catAx>'], 'CT_LineChart');

    const area = chartSpaceXml({ ...BASE, type: 'area' });
    assert.match(area, /<c:areaChart><c:grouping val="standard"\/>/);

    const scatter = chartSpaceXml({ type: 'scatter', series: [{ name: '样本', x: [1, 2, 3], values: [2, 4, 9] }] });
    assert.match(scatter, /<c:scatterStyle val="lineMarker"\/>/);
    assert.match(scatter, /<c:xVal><c:numLit>/);
    assert.match(scatter, /<c:yVal><c:numLit>/);
    assert.ok(!scatter.includes('<c:cat>'), '散点图没有类别轴');

    // 两个系列才会用到调色板的第二格；认不出来的颜色被丢掉（不写进 XML）
    const colored = chartSpaceXml({
        ...BASE,
        series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }],
        colors: ['#FF0000', '00ff00', '不是颜色'],
    });
    assert.match(colored, /<a:srgbClr val="FF0000"\/>/);
    assert.match(colored, /<a:srgbClr val="00FF00"\/>/);
    assert.ok(!colored.includes('不是颜色'));
    return '7 种结构';
});

await check('describeChartSpace：读回类型 / 标题 / 系列数（col 报 column）', async () => {
    assert.deepEqual(describeChartSpace(chartSpaceXml(BASE)), { type: 'column', title: '分季度营收', series: 1 });
    assert.equal(describeChartSpace(chartSpaceXml({ ...BASE, type: 'bar' })).type, 'bar');
    assert.equal(describeChartSpace(chartSpaceXml({ ...BASE, type: 'line' })).type, 'line');
    assert.equal(describeChartSpace(chartSpaceXml({ ...BASE, type: 'pie' })).type, 'pie');
    const two = chartSpaceXml({ ...BASE, series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }] });
    assert.equal(describeChartSpace(two).series, 2);
    assert.equal(describeChartSpace('<c:chartSpace/>').type, 'unknown');
    // 标题里的实体与撇号要走 XML 解析器（手写替换会漏 &apos;、还会二次解码 &amp;lt;）
    const tricky = chartSpaceXml({ ...BASE, title: "It's a 50% & more" });
    assert.equal(describeChartSpace(tricky).title, "It's a 50% & more");
    // 字面量实体文本要**原样**读回来（旧的手写替换会把它二次解码成 `中'文`）
    const entity = describeChartSpace(chartSpaceXml({ ...BASE, title: '中&#39;文' }));
    assert.equal(entity.title, '中&#39;文');
    // 多段标题（a:r 多个）要全拼起来，不能只取第一段
    const multiRun = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>'
        + '<a:p><a:r><a:t>前半</a:t></a:r><a:r><a:t>后半</a:t></a:r></a:p>'
        + '</c:rich></c:tx></c:title><c:plotArea><c:layout/><c:pieChart><c:ser/></c:pieChart></c:plotArea></c:chart></c:chartSpace>';
    assert.equal(describeChartSpace(multiRun).title, '前半后半');
    return '5 种';
});

await check('数值精度：不进小数位也不把小数截掉', async () => {
    const xml = chartSpaceXml({
        type: 'line',
        title: '精度',
        series: [{ name: 's', values: [123456.789012345, 1e-7, -0.1 + 0.2, 1e21] }],
    });
    assert.match(xml, /<c:v>123456\.789012345<\/c:v>/, '长小数不该被截成 6 位');
    assert.match(xml, /<c:v>1e-7<\/c:v>/, '极小值不该被写成 0');
    assert.ok(!/<c:v>0<\/c:v>/.test(xml), '没有值应该被悄悄写成 0');
    assert.match(xml, /<c:v>1e\+21<\/c:v>/, '大数用指数形式（合法的 xsd:double）');
    return '4 个值';
});

await check('spreadsheetDrawingXml：frames 为空 / 缺 chartRelId 时当场报错', async () => {
    assert.throws(() => spreadsheetDrawingXml({ frames: [] }), /frames 不能是空数组/);
    assert.throws(() => spreadsheetDrawingXml({ frames: [{ name: 'x' }] }), /每个 frame 都要给 chartRelId/);
    assert.throws(() => spreadsheetDrawingXml({}), /每个 frame 都要给 chartRelId/, '单个锚点缺 rId 同样拦下');
    return '2 类错误';
});

await check('spreadsheetDrawingXml：twoCellAnchor + graphicFrame + c:chart r:id', async () => {
    const xml = spreadsheetDrawingXml({ chartRelId: 'rId1', name: '图表 1', id: 2, from: { col: 4, row: 1 }, to: { col: 11, row: 15 } });
    assertOrder(xml, ['<xdr:from>', '<xdr:to>', '<xdr:graphicFrame', '<xdr:clientData/>'], 'CT_TwoCellAnchor');
    assert.match(xml, /<xdr:col>4<\/xdr:col>/);
    assert.match(xml, /<xdr:row>15<\/xdr:row>/);
    assert.match(xml, /uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart"/);
    assert.match(xml, /<c:chart [^>]*r:id="rId1"\/>/);
    return '锚点 + 框架';
});

// ── 3. Excel 全链路 ─────────────────────────────────────────────────────────

await check('Excel：图表部件 / rels / Content_Types / 工作表锚点齐全', async () => {
    const wb = createExcel({ path: '图表.xlsx', title: '季度营收' }, env);
    const sheet = wb.sheet('明细');
    sheet.table({ columns: ['季度', '营收'], rows: [['Q1', 1280], ['Q2', 1530], ['Q3', 1410]] });
    sheet.chart({ ...BASE, at: { col: 4, row: 1, toCol: 11, toRow: 15 } });
    const report = wb.save();
    assert.equal(report.ok, true);
    assert.deepEqual(report.warnings, [], '正常参数不该有告警');

    const parts = unzipText(readFileSync(join(root, '图表.xlsx')));
    for (const name of ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml',
        'xl/drawings/_rels/drawing1.xml.rels', 'xl/worksheets/_rels/sheet1.xml.rels']) {
        assert.ok(parts.has(name), `缺部件 ${name}`);
    }
    const sheetXml = parts.get('xl/worksheets/sheet1.xml');
    assert.match(sheetXml, /<drawing r:id="rId1"\/><\/worksheet>/, '工作表末尾要有 <drawing>');
    assert.ok(sheetXml.indexOf('<mergeCells') < sheetXml.indexOf('<drawing') || !sheetXml.includes('<mergeCells'),
        '<drawing> 必须在 mergeCells 之后（CT_Worksheet 顺序）');
    assert.match(parts.get('xl/worksheets/_rels/sheet1.xml.rels'), /Type="[^"]*\/drawing" Target="\.\.\/drawings\/drawing1\.xml"/);
    assert.match(parts.get('xl/drawings/_rels/drawing1.xml.rels'), /Type="[^"]*\/chart" Target="\.\.\/charts\/chart1\.xml"/);
    const ct = parts.get('[Content_Types].xml');
    assert.match(ct, /<Override PartName="\/xl\/charts\/chart1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"\/>/);
    assert.match(ct, /<Override PartName="\/xl\/drawings\/drawing1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawing\+xml"\/>/);
    assert.match(parts.get('xl/drawings/drawing1.xml'), /<xdr:col>4<\/xdr:col>/);
    return `parts=${[...parts.keys()].length}`;
});

await check('Excel：多表多图编号连续，读回认得图表', async () => {
    const wb = createExcel({ path: '多图.xlsx', title: '多图' }, env);
    const one = wb.sheet('一');
    one.table({ columns: ['项', '值'], rows: [['A', 1], ['B', 2]] });
    one.chart({ type: 'column', categories: ['A', 'B'], series: [{ name: '值', values: [1, 2] }] });
    one.chart({ type: 'pie', categories: ['A', 'B'], series: [{ name: '值', values: [1, 2] }] });
    const two = wb.sheet('二');
    two.table({ columns: ['项', '值'], rows: [['X', 3], ['Y', 4]] });
    two.chart({ type: 'line', categories: ['X', 'Y'], series: [{ name: '值', values: [3, 4] }] });
    wb.save();

    const parts = unzipText(readFileSync(join(root, '多图.xlsx')));
    const charts = [...parts.keys()].filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name)).sort();
    const drawings = [...parts.keys()].filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name)).sort();
    assert.deepEqual(charts, ['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml']);
    assert.deepEqual(drawings, ['xl/drawings/drawing1.xml', 'xl/drawings/drawing2.xml']);
    assert.ok(parts.has('xl/worksheets/_rels/sheet1.xml.rels') && parts.has('xl/worksheets/_rels/sheet2.xml.rels'));
    assert.match(parts.get('xl/drawings/_rels/drawing1.xml.rels'), /chart1\.xml/);
    assert.match(parts.get('xl/drawings/_rels/drawing1.xml.rels'), /chart2\.xml/);
    assert.match(parts.get('xl/drawings/_rels/drawing2.xml.rels'), /chart3\.xml/);
    // 一张表两个图表 → 绘图里必须有两个锚点（只锚一个时 Excel 只显示一个）
    const drawing1 = parts.get('xl/drawings/drawing1.xml');
    assert.equal((drawing1.match(/<xdr:twoCellAnchor>/g) ?? []).length, 2, '同表两个图表要有两个锚点');
    assert.match(drawing1, /r:id="rId1"/);
    assert.match(drawing1, /r:id="rId2"/);

    const back = readExcel(join(root, '多图.xlsx'), env);
    assert.equal(back.stats.charts, 3);
    assert.deepEqual(back.charts.map((item) => item.type), ['column', 'pie', 'line']);
    assert.ok(back.outline.some((line) => line.includes('【图表】xl/charts/chart1.xml')), `outline 缺图表行：${back.outline.join(' / ')}`);
    return `${back.stats.charts} 个图表`;
});

await check('Excel：就地编辑保留图表（部件、锚点与 Content_Types 都不丢）并如实提醒', async () => {
    const edited = editExcel(join(root, '多图.xlsx'), [{ sheet: '一', cell: 'B2', value: 99 }], env);
    assert.equal(edited.ok, true);
    assert.ok(edited.warnings.some((line) => /原生图表/.test(line)), `编辑后要提醒图表不重算数据：${JSON.stringify(edited.warnings)}`);
    const parts = unzipText(readFileSync(join(root, '多图.xlsx')));
    assert.ok(parts.has('xl/charts/chart1.xml'));
    assert.match(parts.get('xl/worksheets/sheet1.xml'), /<drawing r:id="rId1"\/>/);
    assert.match(parts.get('[Content_Types].xml'), /\/xl\/charts\/chart1\.xml/);
    assert.match(parts.get('[Content_Types].xml'), /\/xl\/drawings\/drawing2\.xml/);
    const back = readExcel(join(root, '多图.xlsx'), env);
    assert.equal(back.stats.charts, 3, '编辑后图表数不能变');
    assert.equal(back.perSheet[0].range.includes('A1'), true);
    return '编辑后仍有 3 个图表';
});

await check('Excel：参数错误在 chart() 那一行就报出来（不写半个图表）', async () => {
    const wb = createExcel({ path: '坏图.xlsx', title: '坏图' }, env);
    const sheet = wb.sheet('明细');
    assert.throws(() => sheet.chart({ type: 'column', categories: ['a'], series: [{ values: [1, 2] }] }), /数量必须一致/);
    assert.throws(() => sheet.chart({ type: 'nope', series: [{ values: [1] }] }), /type 只支持/);
    assert.equal(sheet.charts.length, 0, '报错的那次不该留下半个图表');
    sheet.table({ columns: ['项', '值'], rows: [['A', 1]] });
    wb.save();
    const parts = unzipText(readFileSync(join(root, '坏图.xlsx')));
    assert.ok(![...parts.keys()].some((name) => name.startsWith('xl/charts/')), '没有合法图表时不该有图表部件');
    assert.ok(!parts.get('xl/worksheets/sheet1.xml').includes('<drawing'));
    return '0 个图表部件';
});

// ── 4. PPT 全链路 ───────────────────────────────────────────────────────────

await check('PPT：图表部件 / 幻灯片 rels / graphicFrame / Content_Types 齐全', async () => {
    const deck = createPpt({ path: '图表.pptx', title: '原生图表' }, env);
    deck.cover({ title: '原生图表' });
    deck.chart({ ...BASE, labels: true });
    deck.chart({ type: 'pie', title: '渠道占比', categories: ['线上', '线下'], series: [{ name: '占比', values: [62, 38] }] });
    const report = deck.save();
    assert.equal(report.ok, true);
    assert.equal(report.stats.charts, 2);
    assert.ok(report.outline.some((line) => /P2 chart：分季度营收｜column 图，1 个系列/.test(line)), report.outline.join(' / '));

    const parts = unzipText(readFileSync(join(root, '图表.pptx')));
    assert.ok(parts.has('ppt/charts/chart1.xml') && parts.has('ppt/charts/chart2.xml'));
    assert.match(parts.get('[Content_Types].xml'), /<Override PartName="\/ppt\/charts\/chart1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"\/>/);
    assert.match(parts.get('ppt/slides/_rels/slide2.xml.rels'), /Type="[^"]*\/chart" Target="\.\.\/charts\/chart1\.xml"/);
    assert.match(parts.get('ppt/slides/_rels/slide3.xml.rels'), /Type="[^"]*\/chart" Target="\.\.\/charts\/chart2\.xml"/);
    const slide2 = parts.get('ppt/slides/slide2.xml');
    assert.match(slide2, /<p:graphicFrame>/);
    assert.match(slide2, /uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart"/);
    assert.match(slide2, /<c:chart [^>]*r:id="rId2"\/>/);
    assert.ok(!slide2.includes('<a:tbl>'), '图表页不该有表格');
    return `parts=${[...parts.keys()].length}`;
});

await check('PPT：读回认得图表页（stats.charts / layout / outline）', async () => {
    const back = readPpt(join(root, '图表.pptx'), env);
    assert.equal(back.stats.charts, 2);
    assert.equal(back.pages[1].charts, 1);
    assert.equal(back.pages[1].layout, 'chart', `第 2 页版式应为 chart，实际 ${back.pages[1].layout}`);
    assert.equal(back.pages[2].layout, 'chart');
    assert.equal(back.stats.layouts.chart, 2);
    assert.equal(back.pages[1].title, '分季度营收');
    assert.ok(back.outline.some((line) => /P2 chart：分季度营收/.test(line)), back.outline.join(' / '));
    return `${back.stats.charts} 个图表页`;
});

await check('PPT：参数错误在 deck.chart() 那一行报出来', async () => {
    const deck = createPpt({ path: '坏图.pptx', title: '坏图' }, env);
    assert.throws(() => deck.chart({ type: 'column', series: [] }), /至少要有一个系列/);
    assert.throws(() => deck.chart({ categories: ['a', 'b'], series: [{ values: [1] }] }), /数量必须一致/);
    deck.cover({ title: '坏图' });
    const report = deck.save();
    assert.equal(report.stats.charts, 0);
    assert.ok(![...unzipText(readFileSync(join(root, '坏图.pptx'))).keys()].some((name) => name.startsWith('ppt/charts/')));
    return '0 个图表部件';
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`chart: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
