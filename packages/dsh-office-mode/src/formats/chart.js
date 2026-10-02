/**
 * 原生图表部件（历史遗留 `11-6`）：Excel 的 `c:chartSpace` 与 PPT 的 `c:chart` 框架。
 *
 * 为什么值得单开一个模块：`11-6` 说的是「Excel / PPT 没有原生图表，Python 画图只是替代
 * 路径」—— Python 画出来是一张**位图**，改不了数据、点不开数据表、在 Office 里也不是
 * 图表对象。这里生成的是真正的 DrawingML 图表部件：数据以 `c:numLit` / `c:strLit`
 * **内联**在图表里（不引用单元格、也不嵌工作簿），所以：
 *   - 打开就能画（不需要刷新、不需要外部数据）；
 *   - 数据是图表自己的缓存，改数据要在 Office 里「编辑数据」（会提示新建工作簿）——
 *     这是内联数据的固有边界，写进文档而不是假装它可编辑。
 *
 * 两种宿主共用同一份 `c:chartSpace`（Excel 的图表部件与 PPT 的图表部件格式一致），
 * 差别只在**外面那层**：Excel 是 `xdr:wsDr` 里锚一个 `xdr:graphicFrame`，PPT 是
 * 幻灯片里的 `p:graphicFrame`。两个生成器都在这里，免得两处各写一遍。
 *
 * 结构顺序按 schema 的 sequence 摆（CT_Chart / CT_PlotArea / 各 chart 组 / 各 ser），
 * 顺序错了真实 Office 会判文件坏 —— 自研解析器与 LibreOffice 都看不出来
 * （见 docs/ooxml-pitfalls.md）。
 *
 * @module dsh-office-mode/formats/chart
 */
import { descendants, escapeXml, parseXml, textOf } from '../engine/xml.js';

/** 支持的图表类型（对外名字 → 内部族）。 */
export const CHART_TYPES = Object.freeze(['bar', 'column', 'line', 'pie', 'area', 'scatter']);

/** 关系类型与内容类型：Excel 与 PPT 用的是同一套 URI。 */
export const REL_TYPE_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
export const REL_TYPE_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
export const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';

export const NS_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const NS_DRAWINGML = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const NS_RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NS_SPREADSHEET_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** 默认配色：与 Office 默认主题的强调色一致，不传 palette 时用它。 */
export const DEFAULT_CHART_PALETTE = Object.freeze([
    '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E',
]);

function asText(value) {
    return typeof value === 'string' ? value : (value === undefined || value === null ? '' : String(value));
}

function asArray(value) {
    return Array.isArray(value) ? value : [];
}

function chartError(message) {
    const error = new Error(`office 图表：${message}`);
    error.code = 'OFFICE_CHART_BAD_SPEC';
    return error;
}

/** 十六进制颜色规范化（`#4472C4` → `4472C4`）；认不出来返回空串。 */
function normalizeColor(value) {
    const text = asText(value).trim().replace(/^#/, '').toUpperCase();
    return /^[0-9A-F]{6}$/.test(text) ? text : '';
}

/**
 * 规范化并校验图表规格。**错了就抛**：图表画出来一半、或者静默少一个系列，
 * 都比一条可执行的错更糟（模型看不到图，只会以为「做了」）。
 *
 * @param {object} raw
 * @param {{type?: string, title?: string, categories?: Array<string|number>,
 *   series?: Array<{name?: string, values?: number[], x?: number[]}>,
 *   legend?: 'bottom'|'right'|'left'|'top'|'none'|boolean, labels?: boolean,
 *   stacked?: boolean, gapWidth?: number, palette?: string[], colors?: string[]}} raw
 */
export function normalizeChartSpec(raw) {
    const input = raw !== null && typeof raw === 'object' ? raw : {};
    const type = asText(input.type ?? 'column').trim().toLowerCase();
    if (!CHART_TYPES.includes(type)) {
        throw chartError(`type 只支持 ${CHART_TYPES.join(' / ')}，收到「${type || '(空)'}」。`
            + '（bar = 横向条形，column = 纵向柱形；想要饼图用 pie、折线用 line、面积用 area、散点用 scatter）');
    }
    const series = asArray(input.series).map((item, index) => {
        const spec = item !== null && typeof item === 'object' ? item : { values: item };
        const values = asArray(spec.values ?? spec.data ?? spec.y).map((value) => Number(value));
        if (values.length === 0) throw chartError(`第 ${index + 1} 个系列没有 values。`);
        if (values.some((value) => !Number.isFinite(value))) {
            throw chartError(`第 ${index + 1} 个系列的 values 里有非数字（${JSON.stringify(spec.values)}）。`);
        }
        const rawX = asArray(spec.x);
        const x = rawX.map((value) => Number(value));
        if (rawX.length > 0 && x.length !== values.length) {
            throw chartError(`第 ${index + 1} 个系列给了 ${x.length} 个 x、${values.length} 个 values ——`
                + '散点图的 x 与 y 必须一一对应（长度相同）。');
        }
        if (x.length > 0 && x.some((value) => !Number.isFinite(value))) {
            throw chartError(`第 ${index + 1} 个系列的 x 里有非数字（散点图用 x + values）。`);
        }
        return {
            name: asText(spec.name ?? spec.label ?? spec.title) || `系列 ${index + 1}`,
            values,
            x: x.length === values.length && x.length > 0 ? x : undefined,
        };
    });
    if (series.length === 0) throw chartError('至少要有一个系列（series: [{name, values}]）。');
    const categories = asArray(input.categories ?? input.labels ?? input.cats).map((item) => asText(item));
    // 不给 categories 时按**最长**系列补 1..n：按第一个系列取会把「第二个系列更长」
    // 变成一条自相矛盾的错（提示里写着「按最长系列推」，代码却做不到）。
    const width = categories.length > 0
        ? categories.length
        : series.reduce((max, item) => Math.max(max, item.values.length), 0);
    for (const [index, item] of series.entries()) {
        if (item.values.length !== width) {
            throw chartError(`第 ${index + 1} 个系列有 ${item.values.length} 个值，`
                + `而这次要画 ${width} 个点（categories ${categories.length > 0 ? `给了 ${categories.length} 个` : '没给，按最长系列'}）`
                + ' —— 数量必须一致。');
        }
    }
    const legendInput = input.legend;
    const legendText = typeof legendInput === 'string' ? legendInput.trim().toLowerCase() : '';
    if (typeof legendInput === 'string' && legendText !== '' && !['bottom', 'right', 'left', 'top', 'none'].includes(legendText)) {
        // 拼错的位置不该静默落回默认值（写 `botom` 却得到底部图例，等于把配置错误藏起来）。
        throw chartError(`legend 只支持 bottom / right / left / top / none（或 false），收到「${legendInput}」。`);
    }
    const legend = legendInput === false || legendText === 'none'
        ? 'none'
        : (['bottom', 'right', 'left', 'top'].includes(legendText) ? legendText : 'bottom');
    const palette = asArray(input.palette ?? input.colors).map(normalizeColor).filter((value) => value !== '');
    return {
        type,
        title: asText(input.title),
        categories: categories.length > 0 ? categories : Array.from({ length: width }, (_item, index) => String(index + 1)),
        series,
        legend,
        labels: input.labels === true || input.dataLabels === true,
        stacked: input.stacked === true,
        gapWidth: Number.isFinite(Number(input.gapWidth)) ? Math.min(500, Math.max(0, Math.trunc(Number(input.gapWidth)))) : 150,
        palette: palette.length > 0 ? palette : [...DEFAULT_CHART_PALETTE],
    };
}

/**
 * 图例位置：对外的可读名字 → ST_LegendPos 的枚举值。
 *
 * 枚举只有 `b` / `tr` / `l` / `r` / `t` —— 写成 `bottom` 这类「可读值」时，
 * 自研解析器与 LibreOffice 都照常渲染，**真实 Excel 直接判文件坏**
 * （本轮 COM 校验抓到的就是这个：同一份包里换成 Excel 自己写的 chart1.xml 立刻通过）。
 */
const LEGEND_POS = Object.freeze({ bottom: 'b', right: 'r', left: 'l', top: 't' });

/**
 * 数字 → c:v 里的文本。
 *
 * 用 JS 的最短往返表示（`String(value)`）：它不会把 `123456.789012345` 悄悄截成
 * `123456.789012`（曾经按 6 位小数取整，等于改了调用方给的数据），也很小概率给出
 * 指数形式（`1e+21`）—— 而指数形式是合法的 `xsd:double` 词法，Excel 认。
 */
function numText(value) {
    return String(value);
}

/** 系列标题：内联字符串，不引用单元格。 */
function serTxXml(name) {
    return `<c:tx><c:v>${escapeXml(name)}</c:v></c:tx>`;
}

function strLitXml(categories) {
    const points = categories.map((item, index) => `<c:pt idx="${index}"><c:v>${escapeXml(item)}</c:v></c:pt>`).join('');
    return `<c:cat><c:strLit><c:ptCount val="${categories.length}"/>${points}</c:strLit></c:cat>`;
}

function numLitXml(values, tag = 'c:val') {
    const points = values.map((item, index) => `<c:pt idx="${index}"><c:v>${numText(item)}</c:v></c:pt>`).join('');
    return `<${tag}><c:numLit><c:ptCount val="${values.length}"/>${points}</c:numLit></${tag}>`;
}

function spPrXml(color) {
    return color === '' ? '' : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
}

function dLblsXml(enabled) {
    return enabled ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
}

/** 标题（富文本）——CT_Chart 的第一个可选元素。 */
function titleXml(title) {
    if (title === '') return '';
    return '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>'
        + `<a:p><a:pPr><a:defRPr sz="1200" b="0"/></a:pPr><a:r><a:rPr lang="zh-CN" sz="1200" b="0"/><a:t>${escapeXml(title)}</a:t></a:r></a:p>`
        + '</c:rich></c:tx><c:overlay val="0"/></c:title>';
}

/** 坐标轴：catAx + valAx（散点图是两条 valAx）。 */
function axesXml(idCat, idVal, { scatter = false } = {}) {
    const scaling = '<c:scaling><c:orientation val="minMax"/></c:scaling>';
    const xAxis = scatter
        ? `<c:valAx><c:axId val="${idCat}"/>${scaling}<c:delete val="0"/><c:axPos val="b"/><c:majorGridlines/><c:crossAx val="${idVal}"/></c:valAx>`
        : `<c:catAx><c:axId val="${idCat}"/>${scaling}<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${idVal}"/></c:catAx>`;
    const yAxis = `<c:valAx><c:axId val="${idVal}"/>${scaling}<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="${idCat}"/></c:valAx>`;
    return `${xAxis}${yAxis}`;
}

/**
 * 一个系列的 XML。散点图的 x/y 走 `c:xVal` / `c:yVal`（且只有它没有 cat/val）。
 *
 * 顺序按 CT_*Ser 的 sequence：idx, order, tx, spPr, **dLbls**, cat, val
 * （散点是 idx, order, tx, spPr, marker, **dLbls**, xVal, yVal）。
 * `c:dLbls` 必须在 cat/val **之前** —— 放在后面时真实 Excel 仍能打开并显示标签，
 * 但那是 schema 违规，严格校验器与第三方 OOXML 消费者会拒收。
 */
function serXml(type, series, index, palette, labels) {
    const color = spPrXml(palette[index % palette.length]);
    const head = `<c:idx val="${index}"/><c:order val="${index}"/>${serTxXml(series.name)}`;
    if (type === 'scatter') {
        const x = series.x ?? series.values.map((_value, at) => at + 1);
        return `<c:ser>${head}${color}<c:marker><c:symbol val="circle"/></c:marker>`
            + `${dLblsXml(labels)}${numLitXml(x, 'c:xVal')}${numLitXml(series.values, 'c:yVal')}</c:ser>`;
    }
    return `<c:ser>${head}${color}${dLblsXml(labels)}${strLitXml(series.categories)}${numLitXml(series.values)}</c:ser>`;
}

/**
 * `c:chartSpace` 部件（Excel 的 `xl/charts/chartN.xml` 与 PPT 的 `ppt/charts/chartN.xml`
 * 用同一份）。
 *
 * @param {object} raw 图表规格（见 normalizeChartSpec）
 * @returns {string}
 */
export function chartSpaceXml(raw) {
    const spec = normalizeChartSpec(raw);
    const idCat = 111111111;
    const idVal = 222222222;
    const axisIds = (spec.type === 'pie' ? '' : `<c:axId val="${idCat}"/><c:axId val="${idVal}"/>`);
    const seriesXml = spec.series
        .map((series, index) => serXml(spec.type, { ...series, categories: spec.categories }, index, spec.palette, spec.labels))
        .join('');

    let group = '';
    if (spec.type === 'bar' || spec.type === 'column') {
        group = '<c:barChart>'
            + `<c:barDir val="${spec.type === 'bar' ? 'bar' : 'col'}"/>`
            + `<c:grouping val="${spec.stacked ? 'stacked' : 'clustered'}"/>`
            + '<c:varyColors val="0"/>'
            + seriesXml
            + `<c:gapWidth val="${spec.gapWidth}"/>`
            + (spec.stacked ? '<c:overlap val="100"/>' : '')
            + axisIds
            + '</c:barChart>';
    } else if (spec.type === 'line') {
        group = '<c:lineChart>'
            + `<c:grouping val="${spec.stacked ? 'stacked' : 'standard'}"/>`
            + '<c:varyColors val="0"/>'
            + seriesXml
            + '<c:marker val="1"/>'
            + axisIds
            + '</c:lineChart>';
    } else if (spec.type === 'area') {
        group = '<c:areaChart>'
            + `<c:grouping val="${spec.stacked ? 'stacked' : 'standard'}"/>`
            + '<c:varyColors val="0"/>'
            + seriesXml
            + axisIds
            + '</c:areaChart>';
    } else if (spec.type === 'pie') {
        group = '<c:pieChart><c:varyColors val="1"/>' + seriesXml + '<c:firstSliceAng val="0"/></c:pieChart>';
    } else if (spec.type === 'scatter') {
        group = '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>'
            + seriesXml + axisIds + '</c:scatterChart>';
    }

    const legend = spec.legend === 'none' ? '' : `<c:legend><c:legendPos val="${LEGEND_POS[spec.legend] ?? 'b'}"/><c:overlay val="0"/></c:legend>`;
    const plotArea = `<c:plotArea><c:layout/>${group}${spec.type === 'pie' ? '' : axesXml(idCat, idVal, { scatter: spec.type === 'scatter' })}</c:plotArea>`;

    return XML_HEAD
        + `<c:chartSpace xmlns:c="${NS_CHART}" xmlns:a="${NS_DRAWINGML}" xmlns:r="${NS_RELATIONSHIPS}">`
        + `<c:chart>${titleXml(spec.title)}<c:autoTitleDeleted val="${spec.title === '' ? 1 : 0}"/>${plotArea}${legend}<c:plotVisOnly val="1"/></c:chart>`
        + '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>'
        + '</c:chartSpace>';
}

/**
 * Excel 的绘图容器（`xl/drawings/drawingN.xml`）：把图表锚在一个单元格区间上。
 *
 * 用 twoCellAnchor：随行列缩放（图会跟着表格走），这也是 Excel 自己插图表时的默认。
 *
 * **一张表上有几个图表就要有几个锚点**：绘图 rels 里声明了 N 条 chart 关系、
 * drawing 里却只锚一个，Excel 只显示那一个（本轮实测：同表两个图表时
 * `ChartObjects.Count` 报 1 —— 第二个图表部件在包里、没有任何东西引用它）。
 *
 * @param {{frames?: Array<object>, chartRelId?: string, name?: string, id?: number,
 *   from?: {col: number, row: number}, to?: {col: number, row: number}}} options
 *   `frames` 给多个锚点；不给时按单个锚点处理（单图表的老写法仍然可用）。
 */
export function spreadsheetDrawingXml(options) {
    const frames = Array.isArray(options?.frames) ? options.frames : [options ?? {}];
    // 空 frames 是调用方的 bug（一张绘图里什么都没有）：给一句可执行的错，
    // 而不是产出一个 r:id="undefined" 的坏部件。
    if (frames.length === 0) throw chartError('spreadsheetDrawingXml：frames 不能是空数组（一张绘图至少要锚一个图表）。');
    for (const frame of frames) {
        if (asText(frame?.chartRelId) === '') throw chartError('spreadsheetDrawingXml：每个 frame 都要给 chartRelId。');
    }
    const body = frames.map((frame, index) => twoCellAnchorXml(frame, index)).join('');
    return XML_HEAD
        + `<xdr:wsDr xmlns:xdr="${NS_SPREADSHEET_DRAWING}" xmlns:a="${NS_DRAWINGML}">`
        + body
        + '</xdr:wsDr>';
}

/** 一个 twoCellAnchor（图表锚在 from → to 这块单元格区域上）。 */
function twoCellAnchorXml(frame, index) {
    const from = frame?.from ?? { col: 0, row: 0 };
    const to = frame?.to ?? { col: 8, row: 15 };
    const id = Number.isFinite(Number(frame?.id)) ? Math.trunc(Number(frame.id)) : index + 2;
    const name = asText(frame?.name) || `图表 ${index + 1}`;
    const anchor = (side) => `<xdr:${side}><xdr:col>${Math.max(0, Math.trunc(side === 'from' ? from.col : to.col))}</xdr:col><xdr:colOff>0</xdr:colOff>`
        + `<xdr:row>${Math.max(0, Math.trunc(side === 'from' ? from.row : to.row))}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${side}>`;
    return '<xdr:twoCellAnchor>'
        + anchor('from') + anchor('to')
        + '<xdr:graphicFrame macro="">'
        + `<xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeXml(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
        + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        + `<a:graphic><a:graphicData uri="${NS_CHART}">`
        + `<c:chart xmlns:c="${NS_CHART}" xmlns:r="${NS_RELATIONSHIPS}" r:id="${escapeXml(asText(frame.chartRelId))}"/>`
        + '</a:graphicData></a:graphic></xdr:graphicFrame>'
        + '<xdr:clientData/>'
        + '</xdr:twoCellAnchor>';
}

/**
 * PPT 的图表框架（`p:graphicFrame` + `c:chart`）。
 *
 * 与表格那个 graphicFrame 同一骨架，只是 `graphicData/@uri` 换成图表、
 * 内容换成 `<c:chart r:id>`。尺寸单位是 EMU。
 *
 * @param {{id: number, name?: string, x: number, y: number, cx: number, cy: number, chartRelId: string}} options
 */
export function presentationChartFrame(options) {
    const name = asText(options?.name) || '图表';
    return '<p:graphicFrame><p:nvGraphicFramePr>'
        + `<p:cNvPr id="${Math.trunc(Number(options.id) || 2)}" name="${escapeXml(name)}"/>`
        + '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
        + '<p:nvPr/></p:nvGraphicFramePr>'
        + `<p:xfrm><a:off x="${Math.round(Number(options.x) || 0)}" y="${Math.round(Number(options.y) || 0)}"/>`
        + `<a:ext cx="${Math.round(Number(options.cx) || 0)}" cy="${Math.round(Number(options.cy) || 0)}"/></p:xfrm>`
        + `<a:graphic><a:graphicData uri="${NS_CHART}">`
        + `<c:chart xmlns:c="${NS_CHART}" xmlns:r="${NS_RELATIONSHIPS}" r:id="${escapeXml(asText(options.chartRelId))}"/>`
        + '</a:graphicData></a:graphic></p:graphicFrame>';
}

/**
 * 从一份 `c:chartSpace` 里读出「这是什么图、标题是什么、几个系列」。
 * 读取侧（Excel / PPT 的 read）共用它 —— 只做几何与结构识别，不重画。
 */
export function describeChartSpace(xml) {
    const text = asText(xml);
    const found = [
        ['barChart', 'bar'], ['lineChart', 'line'], ['pieChart', 'pie'],
        ['areaChart', 'area'], ['scatterChart', 'scatter'], ['doughnutChart', 'doughnut'],
    ].find(([tag]) => text.includes(`<c:${tag}>`));
    // 标题用现成的 XML 解析器读（engine/xml.js）：手写实体替换会漏 `&apos;`、还会
    // 把字面量 `&amp;lt;` 二次解码成 `<`；解析器还认得数字实体。
    let title = '';
    try {
        const root = parseXml(text).children[0];
        const titleNode = descendants(root, 'c:title')[0];
        if (titleNode !== undefined) {
            // 标题可能是多段（a:r 多个）：全部拼起来，别只取第一段。
            title = descendants(titleNode, 'a:t').map((node) => textOf(node)).join('');
        }
    } catch {
        title = '';
    }
    const seriesCount = (text.match(/<c:ser>/g) ?? []).length;
    const barDir = /<c:barDir val="(bar|col)"\/>/.exec(text)?.[1];
    // barChart 有两种方向：barDir=col 是纵向柱形（对外的 column），barDir=bar 是横向条形。
    const type = found === undefined
        ? 'unknown'
        : (found[0] === 'barChart' ? (barDir === 'bar' ? 'bar' : 'column') : found[1]);
    return { type, title, series: seriesCount };
}
