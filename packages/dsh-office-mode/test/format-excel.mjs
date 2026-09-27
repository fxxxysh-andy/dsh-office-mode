/**
 * Excel 格式引擎自测（SPEC §5）。
 *
 * 跑法：`node test/format-excel.mjs`（加 `--keep` 则保留 .office/tmp/excel/ 里的产物，
 * 供随后的 COM 校验用；默认跑完即清理，不留垃圾文件）。
 *
 * 期望值都由测试自己按同一口径算出来（或写死的外部常数），不抄引擎的输出 ——
 * 抄一遍只能证明代码稳定，证明不了它正确。
 */
import { copyFileSync, readFileSync, rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { createEnv, estimateEm } from '../src/engine/kit.js';
import { unzip, unzipText } from '../src/engine/zip.js';
import { attr, children, descendants, parseXml, textOf } from '../src/engine/xml.js';
import { DEFAULT_THEME_ID, resolveTheme } from '../src/engine/theme.js';
import { create, edit, read } from '../src/formats/excel.js';

const KEEP = process.argv.includes('--keep');
// 测试中间产物统一落在**仓库根**的 .office/tmp/<suite>/ 下（不再用包内的 test/.tmp）。
// 只占 excel/ 这个子目录：word / ppt 的自测同时在跑，
// 清理时也只能删这一个子目录，绝不能 rmSync 整个 .office/tmp（会误伤同伴、也可能撞上占用）。
const ROOT = fileURLToPath(new URL('../../../.office/tmp/excel/', import.meta.url));
const SAMPLE = 'excel-full.xlsx';
const EDITED = 'excel-edited.xlsx';

const failures = [];
let checks = 0;

function check(condition, label, detail = '') {
    checks += 1;
    if (!condition) failures.push(detail ? `${label} —— ${detail}` : label);
}

function checkEqual(actual, expected, label) {
    check(actual === expected, label, `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`);
}

function checkNear(actual, expected, tolerance, label) {
    check(Math.abs(actual - expected) <= tolerance, label, `期望约 ${expected}（±${tolerance}），实际 ${actual}`);
}

function checkIncludes(haystack, needle, label) {
    const hit = Array.isArray(haystack) ? haystack.includes(needle) : String(haystack).includes(needle);
    check(hit, label, `没有找到 ${JSON.stringify(needle)}`);
}

// ---------------------------------------------------------------- 测试素材

const COLUMNS = [
    { title: '产品名称', key: 'name', type: 'text' },
    { title: '销量', key: 'qty', type: 'number' },
    { title: '单价（元）', key: 'price', type: 'currency' },
    { title: '毛利率', key: 'margin', type: 'percent', format: '0.0%' },
    { title: '首次上架', key: 'since', type: 'date' },
];

const PRODUCTS = [
    ['工业路由器 R200', 320, 1280, 0.125, '2024-01-15'],
    ['边缘网关 G5', 96, 2580, 0.082, '2024-02-03'],
    ['温控传感器 T9', 1240, 168, 0.236, '2024-02-18'],
    ['远程运维服务', 42, 6800, 0.451, '2024-03-02'],
];

/** 故意写超长的备注：它会把 A 列撑宽，同时用来验证「合并区压住数据」的告警。 */
const LONG_NOTE = '备注：本表由办公模式自动生成，含标题行、表头、斑马纹、合计行与冻结窗格；数据口径为含税销售额，单位人民币元。';

/** 样例的固定版式（0 基）：标题 0、说明 1、制表 2、表头 3、数据 4–7、合计 8、备注 10。 */
const LAYOUT = { title: 0, note: 1, byline: 2, header: 3, dataFrom: 4, dataTo: 7, total: 8, footnote: 10 };
/** 列号（0 基）+ 行号（0 基）→ 'C3'。 */
const ref = (col, row) => `${String.fromCharCode(65 + col)}${row + 1}`;

/** 日期 → 1900 日期系统的序列号（测试自己算）。 */
const serialOf = (year, month, day) => Math.round((Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);

/** 与引擎同口径的宽度估算：全角折 2 格。 */
function visualWidth(text) {
    let total = 0;
    let wide = 0;
    for (const ch of String(text)) {
        total += 1;
        if (estimateEm(ch) >= 0.95) wide += 1;
    }
    return (total - wide) + wide * 2;
}

function buildSample(env) {
    const wb = create({
        theme: 'business',
        path: SAMPLE,
        title: '2024 年 Q1 月度经营分析',
        author: 'DSH Office',
    }, env);

    const report = wb.sheet('经营明细');
    report.title('2024 年 Q1 月度经营分析', { span: 5 });                                    // 第 1 行
    report.note('数据口径：含税销售额，单位人民币元；毛利率 =（售价 − 成本）/ 售价。');         // 第 2 行
    report.row(['制表：财务部', '', '版本 v1.2', '', '导出日期 2024-04-01'], { height: 18, color: 'muted' }); // 第 3 行
    report.table({                                                                            // 表头第 4 行，数据 5–8，合计第 9 行
        columns: COLUMNS,
        rows: PRODUCTS,
        totalRow: true,
        zebra: true,
        freeze: true,
        autofilter: true,
    });
    report.numberFormat('B5:B8', '#,##0');
    report.numberFormat('C5:C8', '¥#,##0.00');
    report.formula('F3', 'SUM(C5:C8)');
    report.cell('F4', '合计口径见 F3', { italic: true, color: 'muted', size: 9 });
    report.widths({ F: 16 });
    report.merge('A11:E11');
    report.cell('A11', LONG_NOTE, { wrap: true, size: 9, color: 'muted' }); // 合并区里的长文本靠自动换行，不报列宽不足
    report.formula('F5', 'SUM(不存在的表!A1:A9)');                // 故意错表名
    report.widths({ G: 3 });                                      // 故意压窄
    report.cell('G2', '这一列故意很窄，用来触发列宽不足的警告');

    const notes = wb.sheet('说明');
    notes.header(['项', '内容', '备注']);                  // 表头第 1 行
    notes.row(['版本', 'v1.2', '草案']);                   // 第 2 行
    notes.row(['生成器', 'dsh-office-mode / excel.js']);   // 第 3 行

    wb.sheet('说明');                                      // 同名第二次 → 自动变成「说明2」

    // 边界：非法输入必须变成 warning，不能抛
    report.cell('NOT-A-REF', '非法地址');
    report.merge('!!');
    report.widths({ ZZ: -5 });
    notes.numberFormat('X', 123);
    notes.autofilter('bad range');
    // 合并 B2:C2 会把 C2 里的「草案」吃掉（B2 在左上角，按合并语义保留）→ 必须报出来
    notes.merge('B2:C2');
    return wb;
}

// ---------------------------------------------------------------- 统计样例

/**
 * 统计样例的固定版式（0 基）：表头 0、数据 1–4 行。
 * 数据里故意混了一个文字销量（华南 =「缺货」），用来验证「非数字跳过」。
 */
const STATS = 'excel-stats.xlsx';
const STATS_EDITED = 'excel-stats-edited.xlsx';
const STATS_LAYOUT = { header: 0, dataFrom: 1, dataTo: 4 };

/** 统计表的期望数字，全部由测试自己算（华南那行不是数字）。 */
const STAT_NUMS = [10, 20, 30];
const statSum = STAT_NUMS.reduce((a, b) => a + b, 0);

/**
 * 造统计能力样例。刻意用素色默认主题：统计块的标题/表头只应加粗，
 * 不能因为新增能力就往默认产出里塞背景填充。
 */
function buildStatsSample(env) {
    const wb = create({ path: STATS, title: '统计能力自测' }, env);
    const sheet = wb.sheet('统计');
    sheet.table({
        columns: [
            { title: '区域', key: 'region', type: 'text' },
            { title: '销量', key: 'qty', type: 'number' },
            { title: '金额', key: 'amount', type: 'currency' },
        ],
        rows: [['华东', 10, 100], ['华北', 20, 200], ['华东', 30, 300], ['华南', '缺货', 400]],
    });
    // 竖排默认布局：块标题 + 一行一个统计项
    sheet.stats('B2:B5', { at: 6 });
    // 横排：表头行 + 数值行；列号按 1 基，2 就是 B
    sheet.stats({ column: 2, from: 2, to: 5 },
        { at: 13, layout: 'columns', label: '销量统计', funcs: ['count', 'sum'], format: '#,##0.00' });
    // header:false 只写数值行；{range} 是第三种源区间写法；标准差必然带小数
    sheet.stats({ range: 'B2:B3' }, { at: 16, layout: 'columns', header: false, funcs: ['stdev'] });
    // 元素级 label / format 覆盖块级；counta 走「非空格数」而不是数字个数
    sheet.stats('B2:B5', {
        at: 18,
        funcs: [{ id: 'average', label: '均值', format: '0.00' }, 'median', 'stdev', 'var', 'product', 'counta'],
    });
    sheet.stats('A2:A5', { at: 27, label: '没有数字' });      // 源区间全是文字 → 留空 + warning
    sheet.stats('NOT-A-RANGE', { at: 34 });                  // 非法区间 → 只出 warning
    sheet.stats('B2:B5', { at: 35, funcs: ['bogus'] });      // 未知统计项 → 只出 warning
    sheet.summary({ at: 37, key: '区域', value: '销量', total: true });
    sheet.stats('B', { at: 44, funcs: ['sum'] });            // 只给列：行范围取 table() 数据区
    sheet.summary({ at: 50, key: '不存在的列' });             // 列找不到 → 只出 warning
    sheet.stats({ column: 'ZZZ' }, { at: 52 });              // 列字母非法 → 只出 warning
    const blank = wb.sheet('无数据区');
    blank.note('这张表从未调用 table()');
    blank.summary({ at: 1, key: 'A', value: 'B' });          // 没有数据区 → 跳过
    blank.stats('B2:B5', { at: -3 });                        // at 非法 → 回落游标；源区间又没数字
    return wb;
}

// ---------------------------------------------------------------- 包解析小工具
/** rels 文件 → 它所属的 part（用于按目录解析相对 Target）。 */
function ownerOfRels(relsPath) {
    // ZIP 里的路径永远是正斜杠，先归一化再匹配（Windows 上 sep 是反斜杠，直接用会漏匹配）
    const posix = String(relsPath).replace(/\\/g, '/');
    // 包根的 _rels/.rels 不属于任何 part，owner 记为空串；它的 Target 一律是包内绝对路径
    if (posix === '_rels/.rels') return '';
    const match = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(posix);
    return match ? `${match[1]}/${match[2]}` : undefined;
}

function resolveTarget(owner, target) {
    const clean = String(target).replace(/\\/g, '/');
    if (clean.startsWith('/')) return clean.slice(1);
    const base = owner === '' ? [] : owner.split('/').slice(0, -1);
    for (const seg of clean.split('/')) {
        if (seg === '' || seg === '.') continue;
        if (seg === '..') base.pop();
        else base.push(seg);
    }
    return base.join('/');
}

/** 把 sheet XML 解成 ref → {t, v, s, f}，用来验证真正写进字节里的东西。 */
function sheetCells(xml) {
    const out = new Map();
    for (const cell of descendants(parseXml(xml).children[0], 'c')) {
        const vNode = descendants(cell, 'v')[0];
        const fNode = descendants(cell, 'f')[0];
        const isNode = descendants(cell, 'is')[0];
        out.set(attr(cell, 'r', ''), {
            t: attr(cell, 't', ''),
            s: attr(cell, 's', '0'),
            v: vNode ? textOf(vNode) : (isNode ? textOf(isNode) : ''),
            f: fNode ? textOf(fNode) : null,
        });
    }
    return out;
}

// ---------------------------------------------------------------- 主流程

async function main() {
    rmSync(ROOT, { recursive: true, force: true });
    const env = createEnv({ root: ROOT, themeResolver: (id) => resolveTheme(id) });

    // ── 1. 造样例（覆盖全部 builder 方法） ──────────────────────────────
    const wb = buildSample(env);

    // 未知主题 → 回落 + warning
    const wbBad = create({ theme: '不存在的主题', path: 'bad-theme.xlsx' }, env);
    wbBad.sheet('Sheet1').row(['x']);
    const badReport = wbBad.save();

    // ── 2. render() 出字节 → unzip() 回读 ──────────────────────────────
    const bytes = wb.render();
    check(bytes instanceof Uint8Array, 'render() 返回 Uint8Array');
    const parts = unzip(bytes);
    const texts = unzipText(bytes);
    const partNames = [...parts.keys()].sort();
    // 3 张表 + Content_Types + .rels + workbook + workbook.rels + styles + core + app = 10
    checkEqual(partNames.length, 11, '部件数量（3 表 + 固定 8 件，含存放布局标记的 docProps/custom.xml）');
    for (const required of [
        '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
        'xl/styles.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml',
        'xl/worksheets/sheet3.xml', 'docProps/core.xml', 'docProps/app.xml',
    ]) {
        checkIncludes(partNames, required, `包内含 ${required}`);
    }
    checkEqual(partNames[0], '[Content_Types].xml', '[Content_Types].xml 排在包首');
    for (const [name, data] of parts) check(data.length > 0, `部件非空：${name}`);

    // ── 3. 解析每个 xml/rels 部件，确认根元素与关键节点 ────────────────
    const roots = new Map();
    for (const [name, text] of texts) {
        if (!/\.(xml|rels)$/.test(name)) continue;
        const root = parseXml(text).children[0];
        check(!!root, `可解析：${name}`);
        roots.set(name, root);
    }
    checkEqual(roots.get('[Content_Types].xml')?.name, 'Types', '[Content_Types].xml 根元素是 Types');
    checkEqual(roots.get('_rels/.rels')?.name, 'Relationships', '_rels/.rels 根元素是 Relationships');
    checkEqual(roots.get('xl/workbook.xml')?.name, 'workbook', 'workbook.xml 根元素是 workbook');
    checkEqual(roots.get('xl/styles.xml')?.name, 'styleSheet', 'styles.xml 根元素是 styleSheet');
    checkEqual(roots.get('xl/worksheets/sheet1.xml')?.name, 'worksheet', 'sheet1 根元素是 worksheet');
    checkEqual(roots.get('docProps/core.xml')?.name, 'cp:coreProperties', 'core.xml 根元素正确');
    checkEqual(roots.get('docProps/app.xml')?.name, 'Properties', 'app.xml 根元素正确');

    // 字符串全用 inlineStr：不建 sharedStrings
    check(!partNames.includes('xl/sharedStrings.xml'), '没有 sharedStrings 部件');
    checkIncludes(texts.get('xl/worksheets/sheet1.xml'), 't="inlineStr"', 'sheet 用 inlineStr 写字符串');
    checkIncludes(texts.get('xl/workbook.xml'), 'fullCalcOnLoad="1"', 'workbook 开了打开时重算');

    // 工作表子元素顺序被 schema 钉死
    const sheet1 = roots.get('xl/worksheets/sheet1.xml');
    const childNames = sheet1.children.map((node) => node.name);
    const order = ['sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'autoFilter', 'mergeCells'];
    const indexes = order.filter((name) => childNames.includes(name)).map((name) => childNames.indexOf(name));
    check(indexes.every((value, index) => index === 0 || value > indexes[index - 1]),
        'sheet 子元素顺序符合 schema', childNames.join(' → '));
    // 空表不能写 <cols></cols>：CT_Cols 里 col 的 minOccurs=1，空的 cols 是非法内容。
    // 这条是真实 Excel 校验抓到的 —— 曾经因为第 3 张空表带着 <cols></cols>，
    // 整个工作簿被 Excel 判为「发现不可读取的内容」，而 LibreOffice 与自研解析器都照常渲染。
    check(!texts.get('xl/worksheets/sheet3.xml').includes('<cols></cols>'), '空表不写空的 <cols></cols>');
    check(![...texts.entries()].some(([name, text]) => name.startsWith('xl/worksheets/') && text.includes('<cols></cols>')),
        '所有工作表都不写空的 cols 元素');
    // pane 是 sheetViews/sheetView 的子元素，不是 worksheet 的直接子元素
    const sheetView = descendants(sheet1, 'sheetView')[0];
    const pane = descendants(sheetView, 'pane')[0];
    checkEqual(attr(pane, 'state', ''), 'frozen', '冻结窗格写进了 sheetViews/pane');
    checkEqual(attr(pane, 'ySplit', ''), '4', '冻结行数 = 表头行号 + 1');
    checkEqual(attr(pane, 'activePane', ''), 'bottomLeft', '只冻行时 activePane 是 bottomLeft');

    // CT_SheetView 只允许 pane / selection / pivotSelection / extLst 四个子元素。
    // 曾经把 tabSelected 写成子元素：LibreOffice 忽略它照常渲染，真实 Excel 直接
    // 拒绝整个工作簿。这条断言防止同类「多写了一个子元素」的回归。
    const allowedViewChildren = new Set(['pane', 'selection', 'pivotSelection', 'extLst']);
    const strayViewChildren = children(sheetView)
        .map((node) => node.name)
        .filter((name) => !allowedViewChildren.has(name));
    checkEqual(strayViewChildren.join(','), '', 'sheetView 里没有非法子元素（tabSelected 只能是属性）');
    checkEqual(attr(sheetView, 'tabSelected', ''), '1', '首表的 tabSelected 是 sheetView 的属性');

    // ── 4. rels 目标必须都在包里 ───────────────────────────────────────
    let relTargets = 0;
    for (const [name, root] of roots) {
        if (!name.endsWith('.rels')) continue;
        const owner = ownerOfRels(name);
        check(owner !== undefined, `可推断 rels 所属 part：${name}`);
        for (const rel of children(root, 'Relationship')) {
            const id = attr(rel, 'Id', '');
            const target = attr(rel, 'Target', '');
            relTargets += 1;
            checkEqual(attr(rel, 'TargetMode', 'Internal'), 'Internal', `${name} 的 ${id} 是内部关系`);
            const part = resolveTarget(owner, target);
            check(parts.has(part), `关系目标存在：${name} ${id} → ${target}`,
                `解析为 ${part}，包内有 ${partNames.join(', ')}`);
        }
    }
    checkEqual(relTargets, 8, '关系总数（根 4 = officeDocument/core/app/custom + workbook 3 表 + 1 样式）');

    // 反向检查：Content_Types 覆盖了每个 xml 部件
    const overrides = children(roots.get('[Content_Types].xml'), 'Override')
        .map((node) => attr(node, 'PartName', '').replace(/^\//, ''));
    for (const name of partNames) {
        if (!name.endsWith('.xml') || name === '[Content_Types].xml') continue;
        checkIncludes(overrides, name, `Content_Types 声明了 ${name}`);
    }

    // ── 5. save() 落盘 + read() 核对 stats ─────────────────────────────
    const saved = wb.save();
    checkEqual(saved.ok, true, 'save() ok');
    checkEqual(saved.format, 'xlsx', 'format 是 xlsx');
    checkEqual(saved.path, SAMPLE, 'save 用了 spec.path');
    checkEqual(saved.theme, 'business', '主题是 business');
    check(env.exists(SAMPLE), '文件确实落盘了');
    checkEqual(env.stat(SAMPLE).size, saved.bytes, '报告里的 bytes 与磁盘一致');
    check(saved.bytes < 60 * 1024, '文件体积合理（< 60KB）', `实际 ${saved.bytes}`);
    checkEqual(saved.parts, 11, 'save 报告里的部件数');

    const back = read(SAMPLE, env);
    const first = back.perSheet[0];
    checkEqual(back.ok, true, 'read() ok');
    checkEqual(back.stats.sheets, 3, 'read: 工作表数');
    checkEqual(back.theme, 'business', 'read: 主题从 app.xml 读回');
    checkEqual(first.range, 'A1:G11', 'read: 第一张表的范围');
    checkEqual(first.rows, 11, 'read: 第一张表行数');
    checkEqual(first.columns, 7, 'read: 第一张表列数');
    // 3 条合计 SUM + F3 求和 + F5 故意错表名 = 5
    checkEqual(back.stats.formulas, 5, 'read: 公式数');
    // 标题 A1:E1 + A11:E11 备注 = 2（合计行不合并，靠强调色区分）；说明表另有 1 个
    checkEqual(back.stats.merges, 3, 'read: 合并区域数');
    checkEqual(first.cells, 39, 'read: 第一张表非空单元格数（空 inlineStr 不计）');
    checkEqual(back.stats.cells, 47, 'read: 全簿非空单元格数');
    check(back.stats.bytes === saved.bytes, 'read: bytes 与文件大小一致');

    // 头 5 行样例：数据必须真的在里面
    checkEqual(first.sample.length, 5, 'read: 给出前 5 行样例');
    checkIncludes(first.sample[LAYOUT.title].join('|'), '2024 年 Q1 月度经营分析', 'read: 第 1 行是标题');
    checkIncludes(first.sample[LAYOUT.note].join('|'), '数据口径', 'read: 第 2 行是说明');
    checkIncludes(first.sample[LAYOUT.header].join('|'), '产品名称', 'read: 表头在第 4 行');
    checkIncludes(first.sample[LAYOUT.header].join('|'), '毛利率', 'read: 表头内容正确');

    // outline：每张表一行，带真实布局
    checkEqual(back.outline.length, 3, 'outline 每张表一行');
    checkIncludes(back.outline[0], '【经营明细】A1:G11', 'outline[0] 带表名与范围');
    checkIncludes(back.outline[0], '表头在第 4 行', 'outline[0] 标出表头行');
    checkIncludes(back.outline[0], '4 行数据', 'outline[0] 标出数据行数');
    checkIncludes(back.outline[0], '合计在第 9 行', 'outline[0] 标出合计行');
    checkIncludes(back.outline[0], '冻结窗格', 'outline[0] 标出冻结窗格');
    checkIncludes(back.outline[2], '空表', 'outline[2] 标出空表');

    // 布局元数据必须与设计一致（read 的 outline 靠它，别让它悄悄漂移）
    checkEqual(first.layout.headerRow, LAYOUT.header, '布局：表头行号');
    checkEqual(first.layout.dataFrom, LAYOUT.dataFrom, '布局：数据起始行号');
    checkEqual(first.layout.dataTo, LAYOUT.dataTo, '布局：数据结束行号');
    checkEqual(first.layout.totalRow, LAYOUT.total, '布局：合计行号');

    // read() 的 headers 必须取自真正的表头行（不是第 1 行的标题）：否则
    // 「按列名找列」永远找不到，模型只能靠猜列号。
    // 表头行里还散落着 F4 那句「合计口径见 F3」，所以前 5 格按列定义逐一对上即可。
    const headerTexts = first.headers.slice(0, COLUMNS.length);
    checkEqual(headerTexts.join('|'), COLUMNS.map((column) => column.title).join('|'), 'read: headers 取自表头行');
    checkEqual(first.headers.indexOf('首次上架'), 4, 'read: 能按表头名定位「首次上架」列');
    check(first.headers[0] !== '2024 年 Q1 月度经营分析', 'read: headers 不是第 1 行的标题');

    // 值与格式：日期/百分比/货币必须是真数字 + 正确 numFmt
    const live = sheetCells(texts.get('xl/worksheets/sheet1.xml'));
    const stylesRoot = roots.get('xl/styles.xml');
    const xfs = descendants(stylesRoot, 'cellXfs')[0];
    const numFmts = new Map(descendants(stylesRoot, 'numFmt')
        .map((node) => [Number(attr(node, 'numFmtId')), attr(node, 'formatCode', '')]));
    // numFmtId < 164 是 Excel 内置编号，不在 styles.xml 里声明，必须自己带一张表；
    // 只查 numFmts 会把「千分位」误判成「没有格式」
    const BUILTIN = new Map([[0, 'General'], [3, '#,##0'], [4, '#,##0.00'], [9, '0%'], [10, '0.00%'], [49, '@']]);
    const xfOf = (address) => children(xfs, 'xf')[Number(live.get(address).s)];
    const fmtOf = (address) => {
        const id = Number(attr(xfOf(address), 'numFmtId', '0'));
        return numFmts.get(id) ?? BUILTIN.get(id) ?? '';
    };
    const alignOf = (address) => attr(descendants(xfOf(address), 'alignment')[0], 'horizontal', '');

    checkEqual(live.get(ref(1, LAYOUT.dataFrom)).t, 'n', '销量列是数字类型');
    checkEqual(Number(live.get(ref(1, LAYOUT.dataFrom)).v), 320, '销量 320 原样写入');
    checkEqual(Number(live.get(ref(2, LAYOUT.dataFrom)).v), 1280, '货币 1280 存成数字');
    checkNear(Number(live.get(ref(3, LAYOUT.dataFrom)).v), 0.125, 1e-9, '百分比 12.5% 存成 0.125');
    checkEqual(Number(live.get(ref(4, LAYOUT.dataFrom)).v), serialOf(2024, 1, 15), '日期存成 2024-01-15 的序列号');

    checkEqual(fmtOf(ref(1, LAYOUT.dataFrom)), '#,##0', '整数列用千分位格式');
    checkEqual(fmtOf(ref(2, LAYOUT.dataFrom)), '¥#,##0.00', '货币列用货币格式');
    checkEqual(fmtOf(ref(3, LAYOUT.dataFrom)), '0.0%', '百分比列用 0.0%');
    checkEqual(fmtOf(ref(4, LAYOUT.dataFrom)), 'yyyy-mm-dd', '日期列用日期格式');
    checkEqual(fmtOf(ref(1, LAYOUT.total)), '#,##0', '合计行继承数字格式');
    checkEqual(live.get(ref(1, LAYOUT.total)).f,
        `SUM(B${LAYOUT.dataFrom + 1}:B${LAYOUT.total})`, '合计行是 SUM 公式');
    checkEqual(live.get(ref(1, LAYOUT.total)).v, '', '公式不写缓存值（靠 fullCalcOnLoad 重算）');

    // 表头一律居中（同一张表里文本列表头左、数字列表头中会显得没对齐），数据行才按类型分左右
    checkEqual(alignOf(ref(0, LAYOUT.header)), 'center', '文本表头也居中');
    checkEqual(alignOf(ref(1, LAYOUT.header)), 'center', '数字表头居中');
    checkEqual(alignOf(ref(1, LAYOUT.dataFrom)), 'right', '数字右对齐');
    checkEqual(alignOf(ref(0, LAYOUT.dataFrom)), 'left', '文本左对齐');
    checkEqual(alignOf(ref(0, LAYOUT.total)), 'left', '合计行标签左对齐');

    // 列宽：中文列必须比数字列宽，且中文列 = 估宽 × 1.2 + 内边距（超长内容按上限 80 截住）
    const cols = new Map(descendants(sheet1, 'col')
        .map((node) => [Number(attr(node, 'min')) - 1, Number(attr(node, 'width'))]));
    const nameWidth = cols.get(0) ?? 0;
    // 换行的格（A11 备注）不参与估宽，所以 A 列由表头与产品名决定
    const widestName = Math.max(
        visualWidth('产品名称'),
        ...PRODUCTS.map((row) => visualWidth(row[0])),
    );
    checkNear(nameWidth, Math.min(widestName * 1.2 + 2.5, 80), 0.02, '中文列宽 = 估宽 ×1.2 + 内边距');
    const qtyWidth = cols.get(1) ?? 0;
    check(qtyWidth > 0 && qtyWidth < nameWidth, '纯数字列比中文列窄', `数字列 ${qtyWidth}，中文列 ${nameWidth}`);
    checkEqual(cols.get(6) ?? 0, 3, '显式列宽 G=3 被写入');

    // ── warnings 必须来自真实计算 ──────────────────────────────────────
    const warnings = saved.warnings;
    const narrow = warnings.filter((text) => text.startsWith('列宽不足'));
    check(narrow.length > 0, '报出了列宽不足');
    checkIncludes(narrow[0], 'G2', '列宽不足指向故意压窄的 G2');
    checkIncludes(narrow[0], '列宽只有 3', '列宽不足带上了真实列宽');
    check(warnings.some((text) => text.includes('不存在的表名') && text.includes('不存在的表')), '报出了不存在的表名');
    check(warnings.some((text) => text.includes('合并区域压住数据') && text.includes('说明!B2:C2')), '报出了合并区压住数据');
    check(warnings.some((text) => text.includes('不合法')), '非法输入变成了 warning 而不是抛异常');
    // 故意压窄的 G 列不该把别的列也拖进警告
    check(!narrow.some((text) => /!B[0-9]/.test(text)), '数字列没有被误报列宽不足');
    check(!narrow.some((text) => /!A[0-9]/.test(text)), 'A 列（估宽自动放足）没有被误报');
    // A11 在合并区里且开了自动换行 → 不该报列宽不足
    check(!narrow.some((text) => /!A11/.test(text)), '合并区里的换行文本没有被误报列宽不足');
    check(badReport.warnings.some((text) => text.includes('主题') && text.includes('回落到')), '未知主题回落后有 warning');
    checkEqual(read('bad-theme.xlsx', env).theme, DEFAULT_THEME_ID, '未知主题落盘后是默认主题');

    // 写入模型：给人工复核数值口径用
    env.writeFile('excel-model.json', `${JSON.stringify({
        file: SAMPLE,
        bytes: saved.bytes,
        expected: {
            sheets: 3,
            firstSheetRange: first.range,
            cells: back.stats.cells,
            formulas: back.stats.formulas,
            merges: back.stats.merges,
            dateSerial: serialOf(2024, 1, 15),
            percentStored: Number(live.get(ref(3, LAYOUT.dataFrom)).v),
            numberFormats: {
                qty: fmtOf(ref(1, LAYOUT.dataFrom)), price: fmtOf(ref(2, LAYOUT.dataFrom)),
                margin: fmtOf(ref(3, LAYOUT.dataFrom)), since: fmtOf(ref(4, LAYOUT.dataFrom)),
            },
        },
        cells: [ref(0, LAYOUT.title), ref(0, LAYOUT.header), ref(1, LAYOUT.dataFrom),
            ref(4, LAYOUT.dataFrom), ref(0, LAYOUT.total)]
            .map((address) => ({ ref: address, v: live.get(address)?.v ?? null, f: live.get(address)?.f ?? null })),
    }, null, 2)}\n`);

    // ── 5b. edit()：一次调用做完所有修改，只写一次盘 ───────────────────
    // edit() 就地改 path 指向的文件，所以先拷一份，原文件留作对照。
    copyFileSync(env.resolve(SAMPLE), env.resolve(EDITED));
    const edited = edit(EDITED, [
        { sheet: '经营明细', cell: 'B5', value: 999 },
        { sheet: '说明', row: { at: 4, values: ['责任人', '张三'] } },
        { sheet: '说明', find: 'v1.2', replace: 'v2.0' },
        { sheet: '经营明细', formula: { ref: 'F6', formula: '=COUNTA(A5:A8)' } },
        { sheet: '经营明细', numberFormat: { range: 'B9:D9', format: '#,##0' } },
        { sheet: '追加表', addSheet: true },
        { sheet: '追加表', cell: 'A1', value: '新表内容' },
        { sheet: '不存在的表', cell: 'A1', value: 'x' },
        { sheet: '说明', cell: 'NOT-A-REF', value: 'x' },
        { sheet: '说明', find: '找不到的字符串', replace: 'x' },
        { sheet: '说明', bogusOp: true },
    ], env);
    checkEqual(edited.ok, true, 'edit() ok');
    checkEqual(edited.path, EDITED, 'edit: 报告路径是相对 root 的正斜杠形式');
    checkEqual(edited.applied.length, 7, 'edit: 应用了 7 个操作');
    checkEqual(edited.skipped.length, 4, 'edit: 跳过了 4 个操作');
    check(edited.skipped.every((item) => typeof item.reason === 'string' && item.reason !== ''), 'edit: 每个 skip 都有原因');
    checkIncludes(edited.skipped.map((item) => item.op), 'bogusOp', 'edit: 未知操作名进了 skipped');
    checkEqual(edited.changes.length, 7, 'edit: changes 与 applied 一一对应');

    const after = read(EDITED, env);
    const afterFirst = after.perSheet[0];
    checkEqual(after.stats.sheets, 4, 'edit: 新增表后共 4 张');
    const editedTexts = unzipText(readFileSync(env.resolve(EDITED)));
    checkIncludes(editedTexts.get('xl/worksheets/sheet2.xml'), 'v2.0', 'edit: find/replace 生效');
    checkIncludes(editedTexts.get('xl/worksheets/sheet2.xml'), '责任人', 'edit: 整行写入生效');
    checkIncludes(editedTexts.get('xl/worksheets/sheet1.xml'), 'COUNTA(A5:A8)', 'edit: 写入公式生效');
    checkIncludes(editedTexts.get('xl/worksheets/sheet1.xml'), '999', 'edit: 单元格改值生效');
    check(editedTexts.has('xl/styles.xml'), 'edit: styles.xml 仍在包里');
    checkEqual((editedTexts.get('xl/_rels/workbook.xml.rels').match(/<Relationship /g) ?? []).length,
        5, 'edit: 关系数量随表数更新');
    checkEqual(afterFirst.rows, 11, 'edit: 原有表结构保持');
    checkEqual(afterFirst.range, 'A1:G11', 'edit: 原有表范围保持');
    checkEqual(afterFirst.freeze?.row, 4, 'edit: 冻结窗格被保留');
    checkEqual(afterFirst.filter, 'A4:E9', 'edit: 自动筛选被保留');
    checkEqual(afterFirst.cells, 40, 'edit: 改值后非空格数（新增 F6 公式）');
    // 新表要能用：再编辑一次，验证包结构仍然自洽
    const second = edit(EDITED, [{ sheet: '追加表', cell: 'B2', value: 'ok' }], env);
    checkEqual(second.applied.length, 1, 'edit: 二次编辑仍然成功');
    checkEqual(read(EDITED, env).stats.sheets, 4, 'edit: 二次编辑后表数不变');

    // ── 5c. 统计块 / 分组汇总（本轮新增能力） ──────────────────────────
    const statsWb = buildStatsSample(env);
    const statsReport = statsWb.save();
    checkEqual(statsReport.path, STATS, 'stats: save 用了 spec.path');
    checkEqual(statsReport.theme, DEFAULT_THEME_ID, 'stats: 不传主题就是素色默认主题');
    checkEqual(statsReport.stats.statsBlocks, 7, 'stats: 记下了 7 个统计块（含源区间没数字的那块）');
    checkEqual(statsReport.stats.summaryBlocks, 1, 'stats: 记下了 1 个汇总块');
    // 统计 14 + 分组 8 + 合计 3 + 只给列的那块 1 = 26
    checkEqual(statsReport.stats.formulas, 26, 'stats: 公式总数（统计 + 汇总 + 合计）');
    check(statsReport.warnings.some((text) => text.includes('统计块') && text.includes('没有数字')), 'stats: 源区间没数字时给了 warning');
    check(statsReport.warnings.some((text) => text.includes('源区间') && text.includes('不合法')), 'stats: 非法源区间变成 warning');
    check(statsReport.warnings.some((text) => text.includes('bogus')), 'stats: 未知统计项变成 warning');
    check(statsReport.warnings.some((text) => text.includes('还没有 table() 数据区')), 'stats: 没有数据区时 summary 给 warning 并跳过');
    check(statsReport.warnings.some((text) => text.includes('不存在的列')), 'stats: 找不到的列标题变成 warning');
    check(statsReport.warnings.some((text) => text.includes('起始行')), 'stats: 非法 at 回落游标并给 warning');
    check(statsReport.warnings.some((text) => text.includes('ZZZ')), 'stats: 超出 Excel 列范围的列号变成 warning');

    // outline 里必须出现「统计块」：只靠报告就该知道表里有一块统计
    checkIncludes(statsReport.outline[0], '统计块', 'stats: outline 标出统计块');
    checkIncludes(statsReport.outline[0], '汇总块', 'stats: outline 标出汇总块');

    const statsTexts = unzipText(env.readFile(STATS));
    const statsXml = statsTexts.get('xl/worksheets/sheet1.xml');
    const statsLive = sheetCells(statsXml);
    const at = (col, row) => statsLive.get(ref(col, row)) ?? {};

    // ① 竖排统计块：项名左、数值右；公式与缓存值必须一一对应
    checkEqual(at(0, 6).v, '统计', 'stats: 块标题写在起始行');
    checkEqual(at(0, 7).v, '计数', 'stats: 默认第一项是计数');
    checkEqual(at(1, 7).f, `COUNT(B2:B5)`, 'stats: 计数写的是真公式');
    checkEqual(Number(at(1, 7).v), STAT_NUMS.length, 'stats: 计数缓存值 = 数字个数（文字那行不算）');
    checkEqual(at(1, 7).t, '', 'stats: 带数字缓存值的公式格不写 t（默认就是数字）');
    checkEqual(at(1, 8).f, 'SUM(B2:B5)', 'stats: 求和公式');
    checkEqual(Number(at(1, 8).v), statSum, 'stats: 求和缓存值是真算出来的');
    checkEqual(at(1, 9).f, 'AVERAGE(B2:B5)', 'stats: 平均值公式');
    checkEqual(Number(at(1, 9).v), statSum / STAT_NUMS.length, 'stats: 平均值缓存值');
    checkEqual(at(1, 10).f, `MAX(B2:B5)`, 'stats: 最大值公式');
    checkEqual(Number(at(1, 10).v), 30, 'stats: 最大值缓存值');
    checkEqual(at(1, 11).f, 'MIN(B2:B5)', 'stats: 最小值公式');
    checkEqual(Number(at(1, 11).v), 10, 'stats: 最小值缓存值');

    // ② layout:'columns'：表头行 + 数值行，首列写源区间
    checkEqual(at(0, 13).v, '销量统计', 'stats: 横排块的标题是自定义 label');
    checkEqual(at(1, 13).v, '计数', 'stats: 横排表头写统计项名');
    checkEqual(at(2, 13).v, '求和', 'stats: 横排表头第二列');
    checkEqual(at(0, 14).v, 'B2:B5', 'stats: 横排数值行首列写源区间');
    checkEqual(at(1, 14).f, 'COUNT(B2:B5)', 'stats: 横排也写公式');
    checkEqual(Number(at(2, 14).v), statSum, 'stats: 横排也写缓存值');
    // {column:2, from:2, to:5} 与 'B2:B5' 必须解析成同一个区间
    checkEqual(at(2, 14).f, 'SUM(B2:B5)', 'stats: {column,from,to} 解析出 B2:B5');

    // ③ header:false 只写数值行；标准差算出来必然带小数
    check(statsLive.get(ref(0, 15)) === undefined, 'stats: header:false 时不写表头行');
    checkEqual(at(0, 16).v, 'B2:B3', 'stats: {range} 写法同样可用');
    checkEqual(at(1, 16).f, 'STDEV(B2:B3)', 'stats: 标准差公式');
    checkNear(Number(at(1, 16).v), Math.sqrt(50), 1e-9, 'stats: 标准差缓存值（样本口径 n−1）');
    checkEqual(at(1, 16).v, String(Number(Math.sqrt(50).toPrecision(12))), 'stats: 缓存值抹掉了浮点尾数噪声');

    // ④ 元素级 label / format 与其余统计量
    checkEqual(at(0, 19).v, '均值', 'stats: 元素级 label 覆盖默认中文名');
    checkEqual(Number(at(1, 19).v), statSum / STAT_NUMS.length, 'stats: 均值缓存值');
    checkEqual(at(1, 20).f, 'MEDIAN(B2:B5)', 'stats: 中位数公式');
    checkEqual(Number(at(1, 20).v), 20, 'stats: 中位数缓存值');
    checkEqual(Number(at(1, 21).v), 10, 'stats: 标准差缓存值');
    checkEqual(at(1, 22).f, 'VAR(B2:B5)', 'stats: 方差公式');
    checkEqual(Number(at(1, 22).v), 100, 'stats: 方差缓存值');
    checkEqual(at(1, 23).f, 'PRODUCT(B2:B5)', 'stats: 乘积公式');
    checkEqual(Number(at(1, 23).v), 6000, 'stats: 乘积缓存值');
    checkEqual(at(1, 24).f, 'COUNTA(B2:B5)', 'stats: 非空计数公式');
    checkEqual(Number(at(1, 24).v), 4, 'stats: 非空计数值含文字格（4 行都有内容）');

    // ⑤ 源区间没有数字：只写项名，统计格真的不写
    checkEqual(at(0, 27).v, '没有数字', 'stats: 空源区间仍然写块标题');
    checkEqual(at(0, 28).v, '计数', 'stats: 空源区间仍然写项名');
    check(statsLive.get(ref(1, 28)) === undefined, 'stats: 空源区间的统计格留空（不写 0 也不写公式）');
    check(statsLive.get(ref(0, 34)) === undefined, 'stats: 非法源区间一行都不写');
    check(statsLive.get(ref(0, 35)) === undefined, 'stats: 统计项全不认识时整块跳过');

    // ⑥ summary：唯一值按首次出现顺序、条件公式 + 缓存值、合计行
    checkEqual(at(0, 37).v, '汇总', 'summary: 块标题');
    checkEqual(at(1, 37).v, '求和', 'summary: 表头列出统计项');
    checkEqual(at(3, 37).v, '平均值', 'summary: 默认 funcs 是 sum/count/average');
    checkEqual(at(0, 38).v, '华东', 'summary: 第一个唯一值是首次出现的「华东」');
    checkEqual(at(0, 39).v, '华北', 'summary: 第二个唯一值按首现顺序是「华北」');
    checkEqual(at(0, 40).v, '华南', 'summary: 第三个唯一值「华南」（销量是文字）');
    checkEqual(at(1, 38).f, 'SUMIF(A2:A5,"华东",B2:B5)', 'summary: 分组求和公式带判据');
    checkEqual(Number(at(1, 38).v), 40, 'summary: 华东求和缓存值');
    checkEqual(at(2, 38).f, 'COUNTIF(A2:A5,"华东")', 'summary: 计数公式只有两个参数（COUNTIF 没有第三参）');
    checkEqual(Number(at(2, 38).v), 2, 'summary: 华东计数缓存值');
    checkEqual(at(3, 38).f, 'AVERAGEIF(A2:A5,"华东",B2:B5)', 'summary: 分组平均公式');
    checkEqual(Number(at(3, 38).v), 20, 'summary: 华东平均缓存值');
    checkEqual(Number(at(1, 40).v), 0, 'summary: 整组没有数字时求和缓存值是 0');
    checkEqual(Number(at(2, 40).v), 0, 'summary: 整组没有数字时计数缓存值是 0');
    check(statsLive.get(ref(3, 40)) === undefined, 'summary: 没有数字可平均时留空（AVERAGEIF 会算成 #DIV/0!）');
    checkEqual(at(0, 41).v, '合计', 'summary: total:true 补一行合计');
    checkEqual(at(1, 41).f, 'SUM(B2:B5)', 'summary: 合计行按整个数值区间求和');
    checkEqual(Number(at(1, 41).v), statSum, 'summary: 合计求和缓存值');
    checkEqual(at(2, 41).f, 'COUNT(B2:B5)', 'summary: 合计行计数公式');
    checkEqual(Number(at(3, 41).v), statSum / STAT_NUMS.length, 'summary: 合计平均值缓存值');

    // ⑦ 只给列名时行范围取 table() 的数据区
    checkEqual(at(0, 44).v, '统计', 'stats: 只给列也能成块');
    checkEqual(at(1, 45).f, 'SUM(B2:B5)', 'stats: 只给列时行范围 = table() 数据区');
    checkEqual(at(1, 45).f, at(1, 8).f, 'stats: 只给列与显式区间解析出同一个区间');

    // ⑧ 数字格式：块级 → 元素级 → 自适应整数
    const statsStyles = parseXml(statsTexts.get('xl/styles.xml')).children[0];
    const statsXfs = children(descendants(statsStyles, 'cellXfs')[0], 'xf');
    const statsNumFmts = new Map(descendants(statsStyles, 'numFmt')
        .map((node) => [Number(attr(node, 'numFmtId')), attr(node, 'formatCode', '')]));
    const STATS_BUILTIN = new Map([[0, 'General'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00']]);
    const statsFmtOf = (address) => {
        const xf = statsXfs[Number(statsLive.get(address)?.s ?? 0)];
        const id = Number(attr(xf, 'numFmtId', '0'));
        return statsNumFmts.get(id) ?? STATS_BUILTIN.get(id) ?? '';
    };
    checkEqual(statsFmtOf(ref(1, 7)), '#,##0', 'stats: 整数缓存值自适应千分位');
    checkEqual(statsFmtOf(ref(1, 14)), '#,##0.00', 'stats: 块级 format 落到统计格');
    checkEqual(statsFmtOf(ref(1, 19)), '0.00', 'stats: 元素级 format 覆盖块级');
    checkEqual(statsFmtOf(ref(1, 16)), '#,##0.00', 'stats: 带小数的缓存值自适应两位小数');

    // ⑨ 默认素色主题：统计块不得引入任何背景填充
    check(!/patternType="solid"/.test(statsTexts.get('xl/styles.xml')), 'stats: 素色主题下统计块没有纯色填充');
    check(!/fgColor/.test(statsTexts.get('xl/styles.xml')), 'stats: 素色主题下统计块不写 fgColor');
    check(!/applyFill/.test(statsTexts.get('xl/styles.xml')), 'stats: 素色主题下没有单元格应用填充');

    // ⑩ calcPr 只能有一个，且落在 schema 允许的位置（sheets 之后）
    const statsWorkbookXml = statsTexts.get('xl/workbook.xml');
    checkEqual((statsWorkbookXml.match(/<calcPr/g) ?? []).length, 1, 'stats: workbook.xml 里 calcPr 恰好一个');
    check(statsWorkbookXml.indexOf('<calcPr') > statsWorkbookXml.indexOf('</sheets>'), 'stats: calcPr 排在 sheets 之后');
    const statsWorkbook = parseXml(statsWorkbookXml).children[0];
    const statsWorkbookChildren = statsWorkbook.children.map((node) => node.name);
    checkEqual(statsWorkbookChildren[statsWorkbookChildren.indexOf('calcPr') - 1], 'sheets', 'stats: calcPr 紧跟在 sheets 之后');
    checkEqual(attr(descendants(statsWorkbook, 'calcPr')[0], 'fullCalcOnLoad', ''), '1', 'stats: calcPr 开了打开时重算');
    checkEqual(attr(descendants(statsWorkbook, 'calcPr')[0], 'calcId', ''), '0', 'stats: calcPr 的 calcId 是 0');

    // ⑪ read() 回读：布局标记里的统计块要跟着走
    const statsBack = read(STATS, env);
    checkEqual(statsBack.stats.statsBlocks, 7, 'stats: read 回读统计块数量');
    checkEqual(statsBack.stats.summaryBlocks, 1, 'stats: read 回读汇总块数量');
    checkEqual(statsBack.stats.formulas, 26, 'stats: read 回读公式数');
    checkIncludes(statsBack.outline[0], '统计块 6 行 × 2 列', 'stats: read 的 outline 由布局标记还原统计块');
    checkIncludes(statsBack.outline[1], '统计块', 'stats: 第二张表的统计块也在 outline 里');
    checkEqual(statsBack.perSheet[0].layout.statsBlocks.length, 6, 'stats: 第一张表 6 个统计块');
    checkEqual(statsBack.perSheet[0].layout.summaryBlocks.length, 1, 'stats: 第一张表 1 个汇总块');

    // ⑫ edit() 重写工作表时不能把统计块的缓存值弄丢，也不能写出第二个 calcPr
    copyFileSync(env.resolve(STATS), env.resolve(STATS_EDITED));
    const statsEdited = edit(STATS_EDITED, [{ sheet: '统计', cell: 'A1', value: '区域2' }], env);
    checkEqual(statsEdited.ok, true, 'stats: edit() ok');
    const statsEditedTexts = unzipText(readFileSync(env.resolve(STATS_EDITED)));
    checkIncludes(statsEditedTexts.get('xl/worksheets/sheet1.xml'), '<f>COUNT(B2:B5)</f><v>3</v>', 'stats: edit 后统计块的公式与缓存值都在');
    checkIncludes(statsEditedTexts.get('xl/worksheets/sheet1.xml'), '区域2', 'stats: edit 的改动生效');
    checkEqual((statsEditedTexts.get('xl/workbook.xml').match(/<calcPr/g) ?? []).length, 1, 'stats: edit 后 calcPr 仍然只有一个');
    checkEqual(read(STATS_EDITED, env).stats.statsBlocks, 7, 'stats: edit 后统计块标记仍然可读');

    // ⑬ 非法输入只该变成 warning：非对象 opts / Symbol / 未知列 / 非法行号都不能抛
    let probeError = '';
    try {
        const probeWb = create({ path: 'excel-stats-probe.xlsx', title: '非法输入探针' }, env);
        const probe = probeWb.sheet('探针');
        probe.table({ columns: [{ title: '甲' }, { title: '乙' }], rows: [['a', 1]] });
        probe.stats('B2', null);
        probe.stats({ column: 'B' }, 'not-an-object');
        probe.stats('B2', { at: Symbol('行号'), funcs: [Symbol('项')] });
        probe.stats({ column: 0 }, { at: 10 });
        probe.summary(null);
        probe.summary({ at: {}, key: Symbol('列') });
        probe.summary({ key: 'B', value: '乙' });
        const probeReport = probeWb.save();
        checkEqual(probeReport.ok, true, 'stats: 非法输入探针仍然能落盘');
    } catch (error) {
        probeError = error?.message ?? String(error);
    }
    checkEqual(probeError, '', 'stats/summary 遇到非法输入只记 warning，不抛异常');

    // ── 6. 打印结果 ────────────────────────────────────────────────────
    console.log('--- 生成报告 ---');
    console.log(`path=${saved.path} bytes=${saved.bytes} theme=${saved.theme} parts=${saved.parts}`);
    console.log(`stats=${JSON.stringify(saved.stats)}`);
    console.log('outline:');
    for (const line of saved.outline) console.log(`  ${line}`);
    console.log('warnings:');
    for (const line of saved.warnings) console.log(`  - ${line}`);
    console.log('--- read() 回读 ---');
    console.log(`stats=${JSON.stringify(back.stats)}`);
    console.log('outline:');
    for (const line of back.outline) console.log(`  ${line}`);
    console.log('--- edit() ---');
    console.log(`applied=${edited.applied.length} skipped=${edited.skipped.length}`);
    for (const item of edited.changes) console.log(`  ✎ ${item}`);
    for (const item of edited.skipped) console.log(`  ⤫ ${item.op}: ${item.reason}`);
    console.log('--- 统计块 / 汇总 ---');
    console.log(`path=${statsReport.path} bytes=${statsReport.bytes} stats=${JSON.stringify(statsReport.stats)}`);
    console.log('outline:');
    for (const line of statsReport.outline) console.log(`  ${line}`);
    console.log('warnings:');
    for (const line of statsReport.warnings) console.log(`  - ${line}`);

    if (!KEEP) rmSync(ROOT, { recursive: true, force: true });
    check(KEEP || !env.exists(SAMPLE), '测试产物已清理（--keep 时保留给 COM 校验）');

    console.log('');
    if (failures.length > 0) {
        console.error(`FAIL excel（${failures.length}/${checks} 项不符）`);
        for (const message of failures.slice(0, 40)) console.error(`  ✗ ${message}`);
        process.exitCode = 1;
        return;
    }
    console.log(`PASS excel bytes=${saved.bytes} parts=${partNames.length} warnings=${saved.warnings.length} checks=${checks}`);
}

await main();
