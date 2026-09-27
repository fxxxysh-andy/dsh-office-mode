/**
 * Word 格式引擎自测（SPEC §5）。
 *
 * 跑法：`node test/format-word.mjs`（加 `--keep` 则保留 .office/tmp/word/ 里的产物，
 * 供随后的 COM 校验用；默认跑完即清理，不留垃圾文件）。
 *
 * 产物统一落在仓库根的 .office/tmp/word/ 这个子目录里：COM 校验按目录扫 .docx，
 * 混着别的引擎的产物会把「这一轮改的东西」和旧文件搅在一起。
 *
 * 这里刻意生成的样例覆盖全部 builder 方法 + 全部 edit 操作，
 * 并且用「跨 run 文本替换」「超长段落」「标题层级跳跃」「图片超宽」「合并单元格越界」
 * 「公式里的未知命令」这几个真实存在坑的场景去逼 warnings 现身 ——
 * 只会报警报不出问题的测试等于没测。
 */
import { deflateSync } from 'node:zlib';
import { mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { createEnv } from '../src/engine/kit.js';
import { crc32, unzip, zip } from '../src/engine/zip.js';
import { attr, children, descendant, descendants, parseXml, textOf } from '../src/engine/xml.js';
import { resolveTheme } from '../src/engine/theme.js';
import { create, edit, read } from '../src/formats/word.js';

const KEEP = process.argv.includes('--keep');
const ROOT = fileURLToPath(new URL('../../../.office/tmp/word/', import.meta.url));
const SAMPLE = 'word-full.docx';

const failures = [];
let checks = 0;

function check(condition, label, detail = '') {
    checks += 1;
    if (!condition) failures.push(detail ? `${label} —— ${detail}` : label);
}

function checkEqual(actual, expected, label) {
    check(actual === expected, label, `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`);
}

// ---------------------------------------------------------------- 测试素材

/** 手搓一张真 PNG（240×160 渐变）：不引第三方库，也顺带验证引擎的魔数嗅探。 */
function makePng(width, height) {
    const raw = Buffer.alloc((width * 3 + 1) * height);
    for (let y = 0; y < height; y += 1) {
        const rowAt = y * (width * 3 + 1);
        raw[rowAt] = 0; // filter: None
        for (let x = 0; x < width; x += 1) {
            const at = rowAt + 1 + x * 3;
            raw[at] = Math.round((x / width) * 200) + 30;
            raw[at + 1] = Math.round((y / height) * 180) + 50;
            raw[at + 2] = 180;
        }
    }
    const chunk = (type, data) => {
        const head = Buffer.alloc(8);
        head.writeUInt32BE(data.length, 0);
        head.write(type, 4, 'latin1');
        const tail = Buffer.alloc(4);
        tail.writeUInt32BE(crc32(Buffer.concat([head.subarray(4), data])), 0);
        return Buffer.concat([head, data, tail]);
    };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;  // bit depth
    ihdr[9] = 2;  // truecolor
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw, { level: 9 })),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

const LONG_PARA = '这是一段刻意写得很长的正文，用来验证引擎能不能真的算出行数并给出「该分段了」的提示。'.repeat(14);

function buildSample(env) {
    const builder = create({
        title: '2026 年度产品与交付说明',
        theme: 'business',
        path: SAMPLE,
        page: 'A4',
        orientation: 'portrait',
        margins: { top: 2.4, bottom: 2.4, left: 2.6, right: 2.6 },
        header: '2026 年度产品与交付说明｜内部资料',
        footer: '第 X 页 / 共 Y 页',
        author: '办公模式测试',
        subject: '格式引擎自测样例',
    }, env);

    builder
        .title('2026 年度产品与交付说明', { subtitle: '产品线、交付节奏与风险提示' })
        .heading('公司简介', 1)
        .para('本公司专注于把复杂的办公文档生成过程压缩成一次函数调用，让模型不必逐字拼 XML。')
        .para([
            { text: '本段用于验证多 run：', bold: true },
            { text: '加粗提示', bold: true, color: 'C55A11' },
            { text: '、普通正文、' },
            { text: '斜体强调', italic: true },
            { text: ' 与 ' },
            { text: '下划线', underline: true },
            { text: ' 混排在同一段落里的效果。' },
        ])
        .heading('产品线', 2)
        .bullets(['办公模式 Word 引擎：样式齐全，中文不回落到宋体', 'Excel 引擎：inlineStr 免 sharedStrings', 'PPT 引擎：全文本框绕开占位符索引'])
        .heading('旗舰型号', 3)
        .steps(['读取需求并选择主题', '生成文档结构', '落盘并给出可核对的报告'])
        .quote('排版反馈必须来自真实解析：模型只靠报告就该知道文件里到底有什么。')
        .code('const report = word.create({ theme: "business" }, env).title("标题").save();\nconsole.log(report.stats.pagesEstimate);')
        .heading('数据概览', 2)
        .table({
            columns: ['季度', '营收（万元）', '同比', '负责人'],
            rows: [
                // 单元格既可以是字符串，也可以是对象：这里顺带验证「显式 fill」与主题斑马纹共存
                ['Q1', { text: '1,280', fill: 'FFF2CC' }, '+12%', '张三'],
                ['Q2', '1,530', '+18%', '李四'],
                ['Q3', '1,410', '+7%', '王五'],
                ['Q4', '1,760', '+25%', '赵六'],
            ],
            widths: [3, 5, 3.5, 4.3],
        })
        .image('assets/sample.png', { caption: '图 1 由脚本即时生成的 PNG（真魔数，非占位）' })
        .image('assets/sample.png', { widthCm: 30, caption: '图 2 故意给到 30cm，用来逼出「超出版心」告警' })
        .pageBreak()
        .heading('实施计划', 1)
        .para('附录部分用分页符另起一页，验证 pagesEstimate 会跟着分页符变化。')
        .spacer(0.5)
        .heading('里程碑', 3)
        .steps(['第 1 周：跑通 COM 校验', '第 2 周：补齐 Excel/PPT 引擎'])
        .toc()
        .para(LONG_PARA)
        .para([{ text: '跨 run' }, { text: '替换测试' }])
        .para('旧标题：待改样式')
        // 公式：一条块级（无编号）+ 一段行内混排，主样例也要覆盖到
        .formula('\\frac{a+b}{c}')
        .para([{ text: '行内公式混排：' }, { math: 'E = mc^2' }, { text: ' 与普通文本同段。' }]);

    return builder;
}

/**
 * 公式样例：把 LaTeX 子集里每一类结构都跑一遍。
 * 单独出一个文件，是因为 COM 校验「Word 是否接受 OMML」这件事必须有一份
 * 只包含公式的干净产物，混在长文档里出问题不好定位。
 */
function buildFormulaSample(env) {
    const builder = create({ title: '公式与数学排版', path: 'word-formula.docx', footer: false }, env);
    builder
        .heading('公式与数学排版', 1)
        .para('下面每一条都是一个独立的公式段落，用来验证 OMML 结构与 Word 的接受度。')
        .formula('E = mc^2', { size: 14 })
        .formula('\\frac{a+b}{c}', { number: '（1）' })
        .formula('\\sqrt[3]{x+1}', { align: 'left', number: '(2)' })
        .formula('\\sqrt{\\frac{x}{y}}')
        .formula('\\sum_{i=1}^{n} a_i')
        .formula('\\prod_{k=1}^{m} \\frac{1}{k}')
        .formula('\\int_0^1 x \\, dx = \\frac{1}{2}')
        .formula('\\oint_C F \\cdot dr')
        .formula('\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1')
        .formula('\\left( \\frac{1}{2} \\right)^2 + \\left[ 1 + 2 \\right]')
        .formula('x_{i}^{2} \\approx \\alpha \\times \\beta \\le \\gamma \\neq \\infty')
        .formula('\\partial \\nabla \\in \\subset \\subseteq \\cup \\cap \\to \\rightarrow \\Rightarrow')
        .formula('\\forall \\exists \\angle \\degree \\cdots \\ldots \\prime')
        .formula('\\text{当} x \\ge 0 \\text{时} \\quad \\mathrm{d}x + \\mathbf{v}')
        .formula('a \\foobar b \\constructor c \\foobar d')
        .formula('a \\; b \\: c \\quad d \\qquad e \\! f \\, g')
        .para([{ text: '勾股定理：' }, { math: 'a^2 + b^2 = c^2' }, { text: '，当 ' }, { math: 'x \\ge 0' }, { text: ' 时成立。' }])
        // 整段只有行内公式：outline 要单独报一行「公式（行内）」
        .para([{ math: '\\sum_{k=1}^{n} k' }]);
    return builder;
}

/**
 * 表格样例：合并单元格、题注、表头重复、列对齐、越界裁剪。
 * 用默认（素色）主题，顺便盯住「默认主题不得引入背景填充」这条硬约束。
 */
function buildTableSample(env) {
    const builder = create({ title: '表格与合并单元格', path: 'word-table.docx', footer: false }, env);
    builder
        .heading('表格与合并单元格', 1)
        .table({
            columns: ['指标', { title: '本期', width: 3.5, align: 'right' }, { title: '上期', width: 3.5, align: 'center' }, { title: '备注', width: 5 }],
            caption: '表 1 合并单元格与列对齐（题注在表格上方）',
            captionPosition: 'above',
            align: 'center',
            rows: [
                [{ text: '营收（万元）', colspan: 2 }, { text: '1,410', align: 'right' }, '含税'],
                [{ text: '华东区', rowspan: 2 }, { text: '820', bold: true }, '760', '环比 +8%'],
                ['590', '560', '环比 +5%'],
                [{ text: '越界的 colspan', colspan: 9 }, '多余的单元格'],
            ],
        })
        .table({
            columns: ['项目', '状态'],
            rows: [['公式', '已支持'], ['合并单元格', '已支持']],
            caption: '表 2 表题在表格下方，且不重复表头',
            captionPosition: 'below',
            headerRepeat: false,
        });
    return builder;
}

// ---------------------------------------------------------------- 包结构校验

function normZipPath(path) {
    const out = [];
    for (const piece of path.split('/')) {
        if (piece === '' || piece === '.') continue;
        if (piece === '..') out.pop();
        else out.push(piece);
    }
    return out.join('/');
}

function checkPackage(bytes) {
    const parts = unzip(bytes);
    const decoder = new TextDecoder('utf-8');
    const text = new Map([...parts].map(([name, data]) => [name, decoder.decode(data)]));

    const required = [
        '[Content_Types].xml', '_rels/.rels', 'docProps/core.xml', 'docProps/app.xml',
        'word/document.xml', 'word/_rels/document.xml.rels', 'word/styles.xml',
        'word/numbering.xml', 'word/settings.xml', 'word/footer1.xml', 'word/header1.xml',
    ];
    for (const name of required) check(parts.has(name), `缺少必需部件 ${name}`);

    // 每个 xml/rels 都必须能被解析，根元素名字要对得上
    const roots = {
        '[Content_Types].xml': 'Types',
        '_rels/.rels': 'Relationships',
        'docProps/core.xml': 'cp:coreProperties',
        'docProps/app.xml': 'Properties',
        'word/document.xml': 'w:document',
        'word/_rels/document.xml.rels': 'Relationships',
        'word/styles.xml': 'w:styles',
        'word/numbering.xml': 'w:numbering',
        'word/settings.xml': 'w:settings',
        'word/footer1.xml': 'w:ftr',
        'word/header1.xml': 'w:hdr',
    };
    for (const [name, want] of Object.entries(roots)) {
        if (!text.has(name)) continue;
        let root;
        try {
            root = parseXml(text.get(name)).children[0];
        } catch (error) {
            check(false, `${name} 解析失败`, String(error));
            continue;
        }
        check(root && root.name === want, `${name} 根元素应为 ${want}`, `实际 ${root?.name}`);
    }

    // 包首必须是 [Content_Types].xml，OOXML 阅读器依赖这个顺序
    checkEqual([...parts.keys()][0], '[Content_Types].xml', 'zip 首个条目');

    // 关系目标必须都在包里：rels 的基目录不是「它的父目录」，
    // 而是「去掉 _rels/<name>.rels 之后剩下的部分」（_rels/.rels 的基目录就是包根）
    const relBaseDir = (name) => {
        const marker = name.lastIndexOf('_rels/');
        if (marker === -1) return '';
        return name.slice(0, marker).replace(/\/$/, '');
    };
    for (const name of parts.keys()) {
        if (!name.endsWith('.rels')) continue;
        const root = parseXml(text.get(name)).children[0];
        const dir = relBaseDir(name);
        let count = 0;
        for (const rel of children(root, 'Relationship')) {
            count += 1;
            if (attr(rel, 'TargetMode') === 'External') continue;
            const target = attr(rel, 'Target');
            const resolved = target.startsWith('/')
                ? target.slice(1)
                : normZipPath(dir === '' ? target : `${dir}/${target}`);
            check(parts.has(resolved), `${name} 的关系目标 ${target} 不在包内`, `解析为 ${resolved}`);
        }
        check(count > 0, `${name} 里没有任何 Relationship`);
    }

    // 每个 part 都必须被 Content_Types 覆盖（Override 或按扩展名的 Default），否则 Word 报「内容有问题」
    const ctRoot = parseXml(text.get('[Content_Types].xml')).children[0];
    const overrides = new Set(children(ctRoot, 'Override').map((node) => attr(node, 'PartName').toLowerCase()));
    const defaults = new Set(children(ctRoot, 'Default').map((node) => attr(node, 'Extension').toLowerCase()));
    for (const name of parts.keys()) {
        if (overrides.has(`/${name.toLowerCase()}`)) continue;
        const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
        check(defaults.has(ext), `部件 ${name} 没有 Content_Types 覆盖`);
    }

    // document.xml 的关键节点
    const doc = parseXml(text.get('word/document.xml'));
    const body = descendants(doc, 'w:body')[0];
    check(!!body, 'document.xml 缺少 w:body');
    checkEqual(descendants(body, 'w:sectPr').length, 1, 'w:sectPr 数量');
    const pgSz = descendants(body, 'w:pgSz')[0];
    const pgMar = descendants(body, 'w:pgMar')[0];
    checkEqual(attr(pgSz, 'w:w'), '11906', 'A4 宽应为 11906 twips');
    checkEqual(attr(pgSz, 'w:h'), '16838', 'A4 高应为 16838 twips');
    checkEqual(attr(pgMar, 'w:left'), '1474', '左边距应为 2.6cm（1474 twips）');
    checkEqual(descendants(body, 'w:footerReference').length, 1, 'footerReference 数量');
    checkEqual(descendants(body, 'w:headerReference').length, 1, 'headerReference 数量');

    const usedStyles = descendants(body, 'w:pStyle').map((node) => attr(node, 'w:val'));
    for (const want of ['Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Quote', 'ListParagraph', 'Code', 'TableText', 'Formula']) {
        check(usedStyles.includes(want), `正文里没有用到样式 ${want}`);
    }
    check(descendants(body, 'w:numPr').length > 0, '列表段落缺少 w:numPr');

    // 公式：主样例里一条块级 + 一段行内混排
    checkEqual(descendants(body, 'm:oMathPara').length, 1, '主样例的块级公式数');
    checkEqual(descendants(body, 'm:oMath').length, 2, '主样例的公式总数（1 块级 + 1 行内）');
    const inlinePara = descendants(body, 'w:p').find((node) => descendants(node, 'm:oMath').length && !descendant(node, 'm:oMathPara'));
    check(!!inlinePara, '找不到行内公式段落');
    // 行内公式必须与普通 run 平级混排，而不是被塞进某个 w:r 里
    check(children(inlinePara, 'm:oMath').length === 1 && children(inlinePara, 'w:r').length >= 2, '行内公式没有与普通 run 平排在同一段里');
    check(/xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math"/.test(text.get('word/document.xml')), 'document.xml 根元素缺少 xmlns:m 声明');

    const tables = descendants(body, 'w:tbl');
    checkEqual(tables.length, 1, '样例表格数量');
    checkEqual(descendants(tables[0], 'w:gridCol').length, 4, '表格列数（w:gridCol）');
    checkEqual(descendants(tables[0], 'w:tblHeader').length, 1, '表头重复标记 w:tblHeader');
    const fills = descendants(tables[0], 'w:shd').map((node) => attr(node, 'w:fill'));
    check(fills.includes('1F4E79'), '表头没有主题色填充', `实际填充 ${[...new Set(fills)].join(',')}`);
    check(fills.includes('F2F6FB'), '缺少斑马纹填充', `实际填充 ${[...new Set(fills)].join(',')}`);
    // 单元格对象里的显式 fill 必须写进 w:shd（空串/白色仍然什么都不写，见 shadeOf）
    check(fills.includes('FFF2CC'), '单元格显式 fill 未生效', `实际填充 ${[...new Set(fills)].join(',')}`);
    checkEqual(descendants(tables[0], 'w:tblBorders').length, 1, '表格边框 w:tblBorders');
    const cellWidths = descendants(tables[0], 'w:tcW').map((node) => Number(attr(node, 'w:w')));
    const gridSum = descendants(tables[0], 'w:gridCol').reduce((sum, node) => sum + Number(attr(node, 'w:w')), 0);
    checkLess(gridSum, 8958 + 1, '表格列宽合计不应超出版心（8958 twips）');
    check(cellWidths.length === 20, '每个单元格都要有 w:tcW', `实际 ${cellWidths.length}`);

    checkEqual(descendants(body, 'a:blip').length, 2, '嵌入图片数量（a:blip）');
    const media = [...parts.keys()].filter((name) => name.startsWith('word/media/'));
    checkEqual(media.length, 2, 'word/media 文件数');
    for (const name of media) {
        const data = parts.get(name);
        check(data[0] === 0x89 && data[1] === 0x50, `${name} 不是合法 PNG 头`);
    }

    const instr = descendants(body, 'w:instrText').map((node) => textOf(node)).join('|');
    check(instr.includes('TOC'), '缺少 TOC 域');

    // styles.xml：内置名必须写成 Word 认的英文内置名，否则标题不会被识别成标题
    const stylesRoot = parseXml(text.get('word/styles.xml')).children[0];
    const styleNodes = children(stylesRoot, 'w:style');
    const ids = styleNodes.map((node) => attr(node, 'w:styleId'));
    for (const want of ['Normal', 'Title', 'Heading1', 'Heading2', 'Heading3', 'Quote', 'ListParagraph', 'Code', 'TableGrid', 'Formula']) {
        check(ids.includes(want), `styles.xml 缺少 w:styleId=${want}`);
    }
    const styleNames = styleNodes.map((node) => attr(children(node, 'w:name')[0], 'w:val'));
    check(styleNames.includes('heading 1') && styleNames.includes('heading 3'), '标题样式缺少内置名（heading 1/3）');
    const defaultFonts = descendants(children(stylesRoot, 'w:docDefaults')[0], 'w:rFonts')[0];
    check(!!attr(defaultFonts, 'w:ascii') && !!attr(defaultFonts, 'w:eastAsia'), 'docDefaults 必须同时给 ascii 与 eastAsia 字体');
    const tableStyle = styleNodes.find((node) => attr(node, 'w:styleId') === 'TableGrid');
    checkEqual(attr(tableStyle, 'w:type'), 'table', 'TableGrid 必须是表格样式（w:type="table"）');
    const heading1 = styleNodes.find((node) => attr(node, 'w:styleId') === 'Heading1');
    checkEqual(attr(descendants(heading1, 'w:outlineLvl')[0], 'w:val'), '0', 'Heading1 需要 w:outlineLvl=0（导航窗格/目录依赖它）');

    // numbering.xml
    const numbering = parseXml(text.get('word/numbering.xml')).children[0];
    check(children(numbering, 'w:abstractNum').length >= 2, 'numbering.xml 至少要有无序+有序两个 abstractNum');
    check(children(numbering, 'w:num').length >= 2, 'numbering.xml 至少要有两个 w:num');

    // 页脚页码域
    const footer = parseXml(text.get('word/footer1.xml')).children[0];
    const instrs = descendants(footer, 'w:instrText').map((node) => textOf(node)).join('|');
    check(instrs.includes('PAGE'), '页脚缺少 PAGE 域');
    check(instrs.includes('NUMPAGES'), '页脚缺少 NUMPAGES 域');
    check(descendants(footer, 'w:fldChar').length >= 4, '页脚域的 fldChar 不完整');

    checkEqual(descendants(parseXml(text.get('word/settings.xml')).children[0], 'w:updateFields').length, 0,
        'settings.xml 不能带 w:updateFields（会让 Word 打开时弹「是否更新域」）');

    // 目录必须「打开即见」而且不带任何「打开时更新域」标记：
    // w:updateFields 与域上的 w:dirty 各自都能单独触发 Word 的更新域弹窗。
    const tocParas = descendants(body, 'w:p').filter((node) => {
        const pPr = children(node, 'w:pPr')[0];
        const styleId = pPr ? attr(children(pPr, 'w:pStyle')[0], 'w:val') : '';
        return /^TOC[123]$/.test(styleId ?? '');
    });
    check(tocParas.length >= 3, '目录条目段落（TOC1–TOC3）数量', `实际 ${tocParas.length}`);
    check(descendants(body, 'w:fldChar').every((node) => attr(node, 'w:dirty') === undefined),
        '任何 fldChar 都不能带 w:dirty');
    const tocFieldParas = descendants(body, 'w:p').filter((node) => descendants(node, 'w:instrText').some((t) => textOf(t).includes('TOC')));
    check(tocFieldParas.length === 1, 'TOC 域只能有一个开始段落', `实际 ${tocFieldParas.length}`);
    const tocHyperlinks = descendants(body, 'w:hyperlink');
    check(tocHyperlinks.length === tocParas.length, '每个目录条目都要有 w:hyperlink 跳转',
        `目录条目 ${tocParas.length}，超链接 ${tocHyperlinks.length}`);
    const anchors = tocHyperlinks.map((node) => attr(node, 'w:anchor'));
    const bookmarks = descendants(body, 'w:bookmarkStart').map((node) => attr(node, 'w:name'));
    for (const anchor of anchors) {
        check(bookmarks.includes(anchor), `目录锚点 ${anchor} 没有对应的 w:bookmarkStart`);
    }
    checkEqual(descendants(body, 'w:bookmarkStart').length, descendants(body, 'w:bookmarkEnd').length, '书签开始/结束数量');
    // 页码必须真的算出来了（不是占位 1）：主样例有分页，目录里至少出现一个大于 1 的页码
    const tocPages = tocParas.map((node) => descendants(node, 'w:t').map((t) => textOf(t)).join(''))
        .map((line) => Number((line.match(/(\d+)\s*$/) ?? [])[1]));
    check(tocPages.some((n) => Number.isFinite(n) && n > 1), '目录页码应是分页估算值，不能全是第 1 页',
        tocPages.join(','));
    // 右对齐点线制表位：页码贴版心右边缘
    // 段落里有两个 w:tab：pPr 里的制表位定义（带 w:pos）与正文里的制表符（不带）
    const tocTabs = descendants(tocParas[0], 'w:tab');
    const tabStop = tocTabs.find((node) => attr(node, 'w:pos') !== undefined);
    const inlineTab = tocTabs.find((node) => attr(node, 'w:pos') === undefined);
    check(tabStop && attr(tabStop, 'w:val') === 'right' && attr(tabStop, 'w:leader') === 'dot',
        '目录段落缺少右对齐点线制表位', `tabs=${JSON.stringify(tocTabs.map((n) => n.attrs))}`);
    check(!!inlineTab, '目录条目里缺少正文制表符（条目与页码之间要有点线）');
    const tocStyles = styleNodes.map((node) => attr(node, 'w:styleId'));
    for (const want of ['TOC1', 'TOC2', 'TOC3']) {
        check(tocStyles.includes(want), `styles.xml 缺少目录样式 ${want}`);
    }

    return { parts, text };
}

function checkLess(actual, limit, label) {
    check(actual < limit, label, `实际 ${actual}，上限 ${limit}`);
}

/**
 * OOXML 的子元素顺序是 schema 强制的：顺序写反 Word 不一定立刻报错，
 * 但经常会以「内容有问题，是否恢复」收场。这里按 CT_* 的定义逐容器校验，
 * 顺便把「未知子元素」也当失败 —— 那通常是元素名拼错。
 */

/** 数学区里可以出现的元素（EG_OMathElements 的子集，只列我们真的会产出的）。 */
const MATH_ELEMENTS = ['m:r', 'm:f', 'm:rad', 'm:sSup', 'm:sSub', 'm:sSubSup', 'm:nary', 'm:d', 'm:func', 'm:limLow'];
/** CT_OMathArg（m:e / m:num / m:sub / m:sup …）的内容：可选的 m:argPr + 数学元素。 */
const MATH_ARG = ['m:argPr', ...MATH_ELEMENTS];

const CHILD_ORDER = {
    'w:p': ['w:pPr', 'w:r'],
    'w:rPr': [
        'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike',
        'w:outline', 'w:shadow', 'w:noProof', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position',
        'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:bdr', 'w:shd', 'w:vertAlign', 'w:rtl', 'w:lang',
    ],
    'w:pPr': [
        'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:widowControl', 'w:numPr',
        'w:pBdr', 'w:shd', 'w:tabs', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:jc',
        'w:textDirection', 'w:outlineLvl', 'w:rPr',
    ],
    'w:tabs': ['w:tab'],
    'w:tbl': ['w:tblPr', 'w:tblGrid', 'w:tr'],
    'w:tblPr': ['w:tblStyle', 'w:tblW', 'w:jc', 'w:tblBorders', 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook'],
    'w:tr': ['w:trPr', 'w:tc'],
    'w:trPr': ['w:trHeight', 'w:tblHeader'],
    'w:tc': ['w:tcPr', 'w:p'],
    'w:tcPr': ['w:tcW', 'w:gridSpan', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:tcMar', 'w:vAlign'],
    'w:sectPr': ['w:headerReference', 'w:footerReference', 'w:pgSz', 'w:pgMar', 'w:cols', 'w:docGrid'],
    'w:style': ['w:name', 'w:basedOn', 'w:next', 'w:link', 'w:uiPriority', 'w:qFormat', 'w:pPr', 'w:rPr', 'w:tblPr'],
    'w:styles': ['w:docDefaults', 'w:latentStyles', 'w:style'],
    'w:lvl': ['w:start', 'w:numFmt', 'w:lvlText', 'w:lvlJc', 'w:pPr', 'w:rPr'],
    'w:abstractNum': ['w:nsid', 'w:multiLevelType', 'w:name', 'w:lvl'],
    'w:numbering': ['w:abstractNum', 'w:num'],
    'w:document': ['w:body'],
    'w:settings': ['w:zoom', 'w:defaultTabStop', 'w:compat'],
    'w:ftr': ['w:p', 'w:tbl'],
    'w:hdr': ['w:p', 'w:tbl'],
    'w:inline': ['w:extent', 'w:effectExtent', 'w:docPr', 'w:cNvGraphicFramePr', 'w:graphic'],
    'wp:inline': ['wp:extent', 'wp:effectExtent', 'wp:docPr', 'wp:cNvGraphicFramePr', 'a:graphic'],
    'pic:pic': ['pic:nvPicPr', 'pic:blipFill', 'pic:spPr'],
    'pic:nvPicPr': ['pic:cNvPr', 'pic:cNvPicPr'],
    'pic:blipFill': ['a:blip', 'a:stretch'],
    'pic:spPr': ['a:xfrm', 'a:prstGeom'],
    'Types': ['Default', 'Override'],
    // ---- OMML（ECMA-376 Part 1 §22.1）：数学元素的顺序同样是 schema 强制的，
    // m:f 是 num→den、m:nary 是 sub→sup→e、m:r 是 rPr→w:rPr→m:t，写反 Word 直接判损坏
    'm:oMathPara': ['m:oMathParaPr', 'm:oMath'],
    'm:oMathParaPr': ['m:jc'],
    'm:oMath': MATH_ELEMENTS,
    'm:r': ['m:rPr', 'w:rPr', 'm:t'],
    'm:rPr': ['m:lit', 'm:nor', 'm:scr', 'm:sty', 'm:brk', 'm:aln'],
    'm:f': ['m:fPr', 'm:num', 'm:den'],
    'm:rad': ['m:radPr', 'm:deg', 'm:e'],
    'm:radPr': ['m:degHide', 'm:ctrlPr'],
    'm:sSup': ['m:sSupPr', 'm:e', 'm:sup'],
    'm:sSub': ['m:sSubPr', 'm:e', 'm:sub'],
    'm:sSubSup': ['m:sSubSupPr', 'm:e', 'm:sub', 'm:sup'],
    'm:nary': ['m:naryPr', 'm:sub', 'm:sup', 'm:e'],
    'm:naryPr': ['m:chr', 'm:limLoc', 'm:grow', 'm:subHide', 'm:supHide', 'm:ctrlPr'],
    'm:d': ['m:dPr', 'm:e'],
    'm:dPr': ['m:begChr', 'm:sepChr', 'm:endChr', 'm:grow', 'm:shp', 'm:ctrlPr'],
    'm:func': ['m:funcPr', 'm:fName', 'm:e'],
    'm:limLow': ['m:limLowPr', 'm:e', 'm:lim'],
    'm:e': MATH_ARG, 'm:num': MATH_ARG, 'm:den': MATH_ARG, 'm:deg': MATH_ARG,
    'm:sub': MATH_ARG, 'm:sup': MATH_ARG, 'm:lim': MATH_ARG, 'm:fName': MATH_ARG,
};

/**
 * 这些容器的子元素在 schema 里是「重复 choice」——顺序自由，但名字仍必须在白名单里。
 * w:p 是典型：pPr 之后可以任意交替出现 w:r 与 m:oMathPara（带编号的公式段落
 * 正是「tab run + oMathPara + tab+编号 run」这种交替），强行排序会误报。
 * 数学区同理：EG_OMathElements 本身就是一个重复 choice，m:r 与 m:f 可以任意穿插。
 */
const FREE_CHILDREN = {
    'w:p': [
        'w:pPr', 'w:r', 'm:oMathPara', 'm:oMath', 'w:hyperlink', 'w:bookmarkStart', 'w:bookmarkEnd',
        'w:proofErr', 'w:fldSimple', 'w:sdt', 'w:ins', 'w:del', 'w:commentRangeStart', 'w:commentRangeEnd',
    ],
    'm:oMath': MATH_ELEMENTS,
    'm:e': MATH_ARG, 'm:num': MATH_ARG, 'm:den': MATH_ARG, 'm:deg': MATH_ARG,
    'm:sub': MATH_ARG, 'm:sup': MATH_ARG, 'm:lim': MATH_ARG, 'm:fName': MATH_ARG,
};

/** 这些子元素若出现，必须排在所有兄弟之前（schema 要求 minOccurs 位置固定）。 */
const FIRST_CHILD_ONLY = new Set(['w:pPr', 'm:argPr']);

function fail(message) {
    checks += 1;
    failures.push(message);
}

/**
 * 名字合法性：XML 里元素/属性名最多只能有一个冒号，且前缀必须在本部件声明过。
 * 这两条是「容忍式解析器看不出来、但 Word 直接判损坏」的典型漏洞：
 * 例如 `<w:w:bottom>` 这种名字连 System.Xml 都会拒绝加载。
 */
function checkNames(text) {
    for (const [partName, xml] of text) {
        if (!/\.(xml|rels)$/.test(partName)) continue;
        const declared = new Set(['xml', 'xmlns']);
        for (const match of xml.matchAll(/xmlns:([A-Za-z_][\w.-]*)\s*=/g)) declared.add(match[1]);
        const used = new Map();
        for (const match of xml.matchAll(/<([A-Za-z_][\w.:-]*)/g)) used.set(match[1], '元素');
        for (const match of xml.matchAll(/\s([A-Za-z_][\w.:-]*)\s*=\s*"/g)) {
            if (!used.has(match[1])) used.set(match[1], '属性');
        }
        for (const [name, kind] of used) {
            if (name.includes(':')) {
                const colons = (name.match(/:/g) ?? []).length;
                check(colons === 1, `${partName}: ${kind}名 ${name} 含多个冒号（前缀被重复拼接）`);
                const prefix = name.slice(0, name.indexOf(':'));
                check(declared.has(prefix), `${partName}: ${kind} ${name} 的前缀 ${prefix} 没有声明 xmlns:${prefix}`);
            }
        }
    }
}

function checkSchemaOrder(text, sink = failures) {
    let visited = 0;
    const report = (message) => {
        if (sink === failures) checks += 1;
        sink.push(message);
    };
    for (const [partName, xml] of text) {
        if (!/\.(xml|rels)$/.test(partName)) continue;
        const root = parseXml(xml).children[0];
        const walk = (node) => {
            const order = CHILD_ORDER[node.name];
            const free = FREE_CHILDREN[node.name];
            if (free) {
                // 顺序自由，但名字必须合法；w:pPr / m:argPr 仍然必须排在最前面
                visited += 1;
                let seenContent = false;
                for (const child of node.children) {
                    if (!free.includes(child.name)) {
                        report(`${partName}: ${node.name} 里出现未知子元素 ${child.name}`);
                        continue;
                    }
                    if (FIRST_CHILD_ONLY.has(child.name)) {
                        if (seenContent) report(`${partName}: ${node.name} 的 ${child.name} 必须排在所有内容之前`);
                    } else seenContent = true;
                }
            } else if (order) {
                visited += 1;
                let last = -1;
                for (const child of node.children) {
                    const at = order.indexOf(child.name);
                    if (at === -1) {
                        report(`${partName}: ${node.name} 里出现未知子元素 ${child.name}`);
                        continue;
                    }
                    if (at < last) {
                        report(`${partName}: ${node.name} 子元素顺序错误 —— ${child.name} 出现在 ${order[last]} 之后`);
                    }
                    last = Math.max(last, at);
                }
            }
            for (const child of node.children) walk(child);
        };
        walk(root);
        if (partName === 'word/document.xml' && sink === failures) {
            const body = descendants(root, 'w:body')[0];
            const last = body?.children[body.children.length - 1];
            check(last?.name === 'w:sectPr', 'w:body 的最后一个子元素必须是 w:sectPr', `实际 ${last?.name}`);
        }
    }
    // 顺序校验本身也要自证：如果 CHILD_ORDER 一个都没命中，这套断言就是空的
    if (sink === failures) check(visited > 50, '子元素顺序校验覆盖的元素太少，检查表可能失效', `实际命中 ${visited} 个容器`);
}

// ---------------------------------------------------------------- 主流程

mkdirSync(ROOT, { recursive: true });
const env = createEnv({ root: ROOT, themeResolver: resolveTheme });
env.writeFile('assets/sample.png', makePng(240, 160));

const builder = buildSample(env);
const bytes = builder.render();
check(bytes instanceof Uint8Array && bytes.length > 1000, 'render() 应返回非空 Uint8Array');
check(bytes[0] === 0x50 && bytes[1] === 0x4b, 'render() 结果不是 ZIP（缺 PK 头）');
checkEqual(builder.render().length, bytes.length, 'render() 必须是纯函数（两次调用长度相同）');

const { parts, text } = checkPackage(bytes);
checkEqual(parts.size, 13, '部件总数（11 个固定部件 + 2 张图）');
checkSchemaOrder(text);
checkNames(text);

// 校验器自证：喂一份子元素顺序写反的片段，必须被抓出来（否则上面那套断言等于没跑）
{
    const sink = [];
    checkSchemaOrder(new Map([['fake.xml', '<?xml version="1.0"?><w:pPr><w:jc w:val="both"/><w:spacing w:line="240"/></w:pPr>']]), sink);
    check(sink.length > 0, '顺序校验器自证：应能识别 <w:jc> 排在 <w:spacing> 之前的非法顺序');
}

const report = builder.save();

console.log('— create —');
console.log(JSON.stringify({
    path: report.path, bytes: report.bytes, theme: report.theme, stats: report.stats,
}, null, 2));
console.log(`warnings(${report.warnings.length}):`);
for (const line of report.warnings) console.log(`  ⚠ ${line}`);
console.log(`outline(${report.outline.length}):`);
for (const line of report.outline) console.log(`  ${line}`);

checkEqual(report.ok, true, 'report.ok');
checkEqual(report.format, 'docx', 'report.format');
checkEqual(report.path, SAMPLE, 'report.path');
checkEqual(report.bytes, bytes.length, 'report.bytes 应等于 render() 长度');
checkEqual(report.theme, 'business', 'report.theme');

// save(显式路径)：COM 校验用的干净产物走这条路径
const draft = builder.save('word-draft.docx');
checkEqual(draft.path, 'word-draft.docx', 'save(显式路径) 的 report.path');
checkEqual(draft.bytes, bytes.length, '同一 builder 二次保存的字节数应一致');
check(env.exists('word-draft.docx'), 'save(显式路径) 没有落盘');

// stats 必须来自真实解析：逐条核对可数的东西
checkEqual(report.stats.headings, 6, 'stats.headings');
checkEqual(report.stats.tables, 1, 'stats.tables');
checkEqual(report.stats.images, 2, 'stats.images');
checkEqual(report.stats.listItems, 3 + 3 + 2, 'stats.listItems（无序 3 + 有序 3 + 附录 2 步）');
checkEqual(report.stats.quotes, 1, 'stats.quotes');
checkEqual(report.stats.codeBlocks, 1, 'stats.codeBlocks');
checkEqual(report.stats.formulas, 2, 'stats.formulas（1 条块级 + 1 个行内）');
check(report.stats.pagesEstimate >= 2, '有分页符的长文档 pagesEstimate 应 ≥ 2', `实际 ${report.stats.pagesEstimate}`);
check(report.stats.words > 200, 'stats.words 应统计到正文+表格文本', `实际 ${report.stats.words}`);
checkEqual(report.stats.contentWidthCm, 15.8, 'A4 版心宽（cm）');
checkEqual(report.stats.pageSize, 'A4 纵向', 'stats.pageSize');

// warnings 必须真的被触发（长段落 / 标题跳跃 / 图片超宽）
const warningText = report.warnings.join('\n');
check(/未分段（超过 500 字）/.test(warningText), '缺少「正文过长」告警', warningText);
check(/标题层级跳跃/.test(warningText), '缺少「标题层级跳跃」告警', warningText);
check(/超出版心/.test(warningText), '缺少「图片超出版心」告警', warningText);

// outline 结构
const outlineText = report.outline.join('\n');
check(outlineText.includes('H1 公司简介'), 'outline 缺少 H1 公司简介', outlineText);
check(outlineText.includes('表格 5×4'), 'outline 缺少表格 5×4', outlineText);
check(/目录（预渲染 \d+ 条，按 F9 可刷新）/.test(outlineText), 'outline 缺少预渲染目录行', outlineText);
check(outlineText.includes('分页'), 'outline 缺少分页行');
check(outlineText.includes('留白 0.50cm'), 'outline 缺少留白行', outlineText);
check(outlineText.includes('代码 2 行'), 'outline 缺少代码行数', outlineText);
check(/图片 \d+\.\d×\d+\.\dcm/.test(outlineText), 'outline 缺少图片尺寸行', outlineText);
check(outlineText.includes('封面 2026 年度产品与交付说明'), 'outline 缺少封面行', outlineText);

// read() 与 create() 必须走同一条代码路径，stats/outline 应完全一致
const reread = read(SAMPLE, env);
checkEqual(JSON.stringify(reread.stats), JSON.stringify(report.stats), 'read().stats 应与 create().stats 一致');
checkEqual(JSON.stringify(reread.outline), JSON.stringify(report.outline), 'read().outline 应与 create().outline 一致');
checkEqual(reread.theme, 'business', 'read() 反查主题（靠 Heading1 颜色）');
checkEqual(reread.bytes, bytes.length, 'read().bytes');
checkEqual(reread.format, 'docx', 'read().format');

// ---------------------------------------------------------------- edit（原地改，一次写盘）

const editReport = edit(SAMPLE, [
    { find: 'run替换', replace: 'RUN-已合并', all: false },
    { find: '旗舰型号', replace: '旗舰型号（已改名）' },
    { find: '这段文字不存在', replace: 'x' },
    { setStyle: { match: /^旧标题/, style: 'Heading2' } },
    { append: { type: 'para', text: '追加正文：由 edit() 写入。' } },
    { append: { type: 'heading', text: '附录 A', level: 2 } },
    { append: { type: 'bullets', items: ['追加要点一', '追加要点二'] } },
    { append: { type: 'table', columns: ['项目', '状态'], rows: [['COM 校验', '待跑'], ['自测', '通过']] } },
    { append: { type: 'video', text: '不支持的类型' } },
    { setFooter: '第 X 页 / 共 Y 页' },
    { totallyUnknownOp: true },
], env);

console.log('— edit —');
console.log(JSON.stringify({
    path: editReport.path, bytes: editReport.bytes, theme: editReport.theme,
    stats: editReport.stats, applied: editReport.applied, skipped: editReport.skipped,
    changes: editReport.changes,
}, null, 2));

checkEqual(editReport.ok, true, 'edit().ok');
checkEqual(editReport.applied.length, 8, 'applied 操作数');
checkEqual(editReport.skipped.length, 3, 'skipped 操作数（find 未命中 / append.type 非法 / 未知操作）');
check(editReport.skipped.some((item) => /找不到/.test(item.reason)), 'skipped 应说明 find 未命中');
check(editReport.skipped.some((item) => /append\.type/.test(item.reason)), 'skipped 应说明 append.type 非法');
check(editReport.skipped.some((item) => /未知操作/.test(item.reason)), 'skipped 应说明未知操作');
checkEqual(editReport.stats.headings, 8, 'setStyle + append heading 后标题数应为 8');
checkEqual(editReport.stats.tables, 2, 'append 表格后应有两个表格');
checkEqual(editReport.stats.listItems, 8 + 2, 'append 列表后 listItems 应 +2');

const editedParts = unzip(env.readFile(SAMPLE));
const editedDoc = new TextDecoder().decode(editedParts.get('word/document.xml'));
check(editedDoc.includes('RUN-已合并'), '跨 run 替换未生效');
check(!/<w:t[^>]*>替换测试<\/w:t>/.test(editedDoc), '跨 run 替换后仍残留被替换文本');
// 跨 run 命中时只改写命中的两个 run，前后缀必须原样保留：
// 「跨 run」+「替换测试」中替换 “run替换” 后应得到两个 run「跨 RUN-已合并」「测试」
check(editedDoc.includes('<w:t xml:space="preserve">跨 RUN-已合并</w:t>'), '跨 run 替换的前缀没保留', editedDoc.slice(editedDoc.indexOf('RUN-已合并') - 200, editedDoc.indexOf('RUN-已合并') + 80));
check(editedDoc.includes('<w:t xml:space="preserve">测试</w:t>'), '跨 run 替换的后缀没保留');
check(editedDoc.includes('旗舰型号（已改名）'), '单 run 替换未生效');
check(editedDoc.includes('附录 A'), 'append heading 未生效');
check(editedDoc.includes('追加正文：由 edit() 写入。'), 'append para 未生效');
checkEqual((editedDoc.match(/<w:footerReference/g) ?? []).length, 1, 'edit 后应恰好保留一个 footerReference');

const footerText = new TextDecoder().decode(editedParts.get('word/footer1.xml'));
check(footerText.includes('NUMPAGES'), 'setFooter 未写入 NUMPAGES 域');
const editedRels = new TextDecoder().decode(editedParts.get('word/_rels/document.xml.rels'));
checkEqual((editedRels.match(/Target="footer1\.xml"/g) ?? []).length, 1, 'edit 后 footer 关系不应重复');
checkEqual((editedRels.match(/Id="rId4"/g) ?? []).length, 1, 'edit 后关系 id 不应重复');

// 改完再读一遍：edit 的结果必须同样可解析、可核对
const editedRead = read(SAMPLE, env);
console.log('— read(edited) —');
console.log(JSON.stringify({ theme: editedRead.theme, stats: editedRead.stats, warnings: editedRead.warnings }, null, 2));
console.log(`outline(${editedRead.outline.length}):`);
for (const line of editedRead.outline) console.log(`  ${line}`);

check(editedRead.outline.includes('H2 旧标题：待改样式'), 'setStyle 未把段落改成 Heading2', editedRead.outline.join('|'));
check(editedRead.outline.includes('H2 附录 A'), 'append 的标题未出现');
// 正文段落在 outline 里只有字数，所以「内容确实写进去了」要回到文档 XML 上确认
check(editedRead.outline.includes('正文 11 字'), '跨 run 替换后的段落实长不对（应为「跨 RUN-已合并测试」11 字）', editedRead.outline.join('|'));
check(editedRead.outline.includes('表格 3×2'), 'append 的表格未出现', editedRead.outline.join('|'));

// ---------------------------------------------------------------- 公式（OMML）

/** 段落/单元格的纯文本。 */
const textOfNode = (node) => descendants(node, 'w:t').map((child) => textOf(child)).join('');
/** 数学元素里的文字（m:t），用来核对解析出来的内容。 */
const mathText = (node) => descendants(node, 'm:t').map((child) => textOf(child)).join('');
/** 某个容器的第一个直接子元素。 */
const first = (node, name) => children(node, name)[0];
/** 单元格的 tcPr。 */
const cellPr = (cell) => first(cell, 'w:tcPr');
/** 按文字找单元格。 */
const cellWithText = (table, value) => descendants(table, 'w:tc').find((cell) => textOfNode(cell) === value);

const formulaBuilder = buildFormulaSample(env);
const formulaBytes = formulaBuilder.render();
const formulaReport = formulaBuilder.save();
const formulaParts = unzip(formulaBytes);
const formulaText = new TextDecoder().decode(formulaParts.get('word/document.xml'));
const formulaRoot = parseXml(formulaText).children[0];
const formulaBody = descendants(formulaRoot, 'w:body')[0];
const mathParas = descendants(formulaBody, 'm:oMathPara');
const mathOmaths = mathParas.map((node) => first(node, 'm:oMath'));

console.log('— 公式样例 —');
console.log(JSON.stringify({
    path: formulaReport.path, stats: formulaReport.stats, warnings: formulaReport.warnings,
}, null, 2));
console.log(`outline(${formulaReport.outline.length}):`);
for (const line of formulaReport.outline) console.log(`  ${line}`);

// OMML 的元素顺序是这一轮的硬指标之一：把顺序校验器直接喂给公式样例
checkSchemaOrder(new Map([['word/document.xml', formulaText]]));
checkNames(new Map([['word/document.xml', formulaText]]));

checkEqual(formulaReport.stats.formulas, 19, '公式样例的 stats.formulas（16 条块级 + 3 个行内）');
checkEqual(mathParas.length, 16, '块级公式段落数（m:oMathPara）');
checkEqual(descendants(formulaBody, 'm:oMath').length, 19, 'm:oMath 总数');
check(descendants(formulaBody, 'm:r').length > 20, '公式里应该有大量 m:r');

// 1) 上标：E = mc^2 → m:sSup(e=mc, sup=2)；size:14 要写进 run 的 w:sz（14pt = 28 half-points）
{
    const sup = first(mathOmaths[0], 'm:sSup');
    check(!!sup, 'E=mc^2 没有生成 m:sSup');
    checkEqual(mathText(first(sup, 'm:e')), 'mc', 'm:sSup 的 m:e');
    checkEqual(mathText(first(sup, 'm:sup')), '2', 'm:sSup 的 m:sup');
    const sizedRun = descendants(mathOmaths[0], 'm:r')[0];
    checkEqual(attr(first(first(sizedRun, 'w:rPr'), 'w:sz'), 'w:val'), '28', 'formula 的 size 选项没有写进 run');
}
// 2) 分数 + 右侧编号：m:f(num,den) + 居中/右制表位 + oMathParaPr 居中
{
    const omath = mathOmaths[1];
    const frac = first(omath, 'm:f');
    check(!!frac, '\\frac 没有生成 m:f');
    checkEqual(mathText(first(frac, 'm:num')), 'a+b', 'm:f 的 m:num');
    checkEqual(mathText(first(frac, 'm:den')), 'c', 'm:f 的 m:den');
    checkEqual(children(frac, 'm:num').length + children(frac, 'm:den').length, 2, 'm:f 只能有 num 与 den');
    const para = descendants(formulaBody, 'w:p').find((node) => descendant(node, 'm:oMathPara') === mathParas[1]);
    const tabs = children(first(para, 'w:pPr'), 'w:tabs')[0];
    check(!!tabs, '带编号的公式段落缺少 w:tabs');
    checkEqual(children(tabs, 'w:tab').map((node) => attr(node, 'w:val')).join(','), 'center,right', '制表位类型（居中 + 右）');
    checkEqual(attr(first(para, 'w:pPr') && first(first(para, 'w:pPr'), 'w:jc'), 'w:val'), 'left', '带编号的公式段落应左对齐（靠制表位排版）');
    checkEqual(attr(first(mathParas[1], 'm:oMathParaPr') && first(first(mathParas[1], 'm:oMathParaPr'), 'm:jc'), 'm:val'), 'center', 'oMathParaPr 的 m:jc');
    checkEqual(textOfNode(para), '（1）', '公式编号文本');
    check(!!descendants(para, 'w:tab').length, '带编号的公式段落里没有 w:tab');
}
// 3) 带次数的根式：m:rad(degHide=0, deg=3, e=x+1)；align:'left' + 编号时只留右制表位
{
    const rad = first(mathOmaths[2], 'm:rad');
    check(!!rad, '\\sqrt[n]{} 没有生成 m:rad');
    checkEqual(attr(first(first(rad, 'm:radPr'), 'm:degHide'), 'm:val'), '0', '带次数的根式 degHide 应为 0');
    checkEqual(mathText(first(rad, 'm:deg')), '3', 'm:rad 的 m:deg');
    checkEqual(mathText(first(rad, 'm:e')), 'x+1', 'm:rad 的 m:e');
    checkEqual(attr(first(mathParas[2], 'm:oMathParaPr') && first(first(mathParas[2], 'm:oMathParaPr'), 'm:jc'), 'm:val'), 'left', 'align:"left" 的 oMathParaPr');
    const leftPara = descendants(formulaBody, 'w:p').find((node) => descendant(node, 'm:oMathPara') === mathParas[2]);
    checkEqual(children(first(leftPara, 'w:pPr'), 'w:tabs')[0] && children(children(first(leftPara, 'w:pPr'), 'w:tabs')[0], 'w:tab').map((node) => attr(node, 'w:val')).join(','), 'right', 'align:"left" + 编号时只应有右制表位');
    checkEqual(children(leftPara).map((node) => node.name).join(','), 'w:pPr,m:oMathPara,w:r', '左对齐公式不应有前置制表 run');
    checkEqual(textOfNode(leftPara), '(2)', '左对齐公式的编号文本');
}
// 4) 无次数的根式：degHide=1 且 m:deg 为空，被开方数里嵌套 m:f
{
    const rad = first(mathOmaths[3], 'm:rad');
    checkEqual(attr(first(first(rad, 'm:radPr'), 'm:degHide'), 'm:val'), '1', '平方根 degHide 应为 1');
    checkEqual(mathText(first(rad, 'm:deg')), '', '平方根的 m:deg 应为空');
    check(!!first(first(rad, 'm:e'), 'm:f'), '根号里的分数没有嵌套 m:f');
}
// 5) 大运算符：m:nary 的顺序必须是 naryPr→sub→sup→e
{
    const nary = first(mathOmaths[4], 'm:nary');
    check(!!nary, '\\sum 没有生成 m:nary');
    checkEqual(children(nary).map((node) => node.name).join(','), 'm:naryPr,m:sub,m:sup,m:e', 'm:nary 子元素顺序');
    checkEqual(attr(first(first(nary, 'm:naryPr'), 'm:chr'), 'm:val'), '∑', 'm:nary 的 m:chr');
    checkEqual(attr(first(first(nary, 'm:naryPr'), 'm:limLoc'), 'm:val'), 'undOvr', '求和上下限位置 limLoc');
    checkEqual(mathText(first(nary, 'm:sub')), 'i=1', 'm:nary 的下限');
    checkEqual(mathText(first(nary, 'm:sup')), 'n', 'm:nary 的上限');
    checkEqual(mathText(first(nary, 'm:e')), 'ai', 'm:nary 的作用对象（a_i）');
    check(!!first(first(nary, 'm:e'), 'm:sSub'), 'm:nary 的 e 里没有嵌套 m:sSub');
}
// 6) 积分：chr=∫、上下限放在角标位置（subSup）
{
    const prod = first(mathOmaths[5], 'm:nary');
    checkEqual(attr(first(first(prod, 'm:naryPr'), 'm:chr'), 'm:val'), '∏', '\\prod 的 m:chr');
    check(!!first(first(prod, 'm:e'), 'm:f'), '\\prod 的作用对象应是分数');
    const integral = first(mathOmaths[6], 'm:nary');
    checkEqual(attr(first(first(integral, 'm:naryPr'), 'm:chr'), 'm:val'), '∫', '\\int 的 m:chr');
    checkEqual(attr(first(first(integral, 'm:naryPr'), 'm:limLoc'), 'm:val'), 'subSup', '积分的上下限应在角标位置');
    checkEqual(mathText(first(integral, 'm:sub')), '0', '积分的下限');
    checkEqual(mathText(first(integral, 'm:sup')), '1', '积分的上限');
    checkEqual(attr(first(first(first(mathOmaths[7], 'm:nary'), 'm:naryPr'), 'm:chr'), 'm:val'), '∮', '\\oint 的 m:chr');
    checkEqual(attr(first(first(first(mathOmaths[7], 'm:nary'), 'm:naryPr'), 'm:supHide'), 'm:val'), '1', '没有上限时 supHide 应为 1');
}
// 7) 极限与函数名：\lim 用 m:limLow，\sin 用 m:func（fName 正体）
{
    const lim = first(mathOmaths[8], 'm:limLow');
    check(!!lim, '\\lim 没有生成 m:limLow');
    checkEqual(children(lim).map((node) => node.name).join(','), 'm:e,m:lim', 'm:limLow 子元素顺序');
    checkEqual(mathText(first(lim, 'm:e')), 'lim', 'm:limLow 的 m:e');
    checkEqual(mathText(first(lim, 'm:lim')), 'x→0', 'm:limLow 的极限条件');
    const func = descendant(mathOmaths[8], 'm:func');
    check(!!func, '\\sin 没有生成 m:func');
    checkEqual(mathText(first(func, 'm:fName')), 'sin', 'm:func 的函数名');
    checkEqual(attr(first(first(first(func, 'm:fName'), 'm:r'), 'm:rPr') && first(first(first(first(func, 'm:fName'), 'm:r'), 'm:rPr'), 'm:sty'), 'm:val'), 'p', '函数名必须是正体（m:sty="p"）');
}
// 8) 定界符：\left(…\right) 生成 m:d，外面还能挂上标
{
    const sup = first(mathOmaths[9], 'm:sSup');
    const delim = first(first(sup, 'm:e'), 'm:d');
    check(!!delim, '\\left(\\right) 没有生成 m:d');
    checkEqual(attr(first(first(delim, 'm:dPr'), 'm:begChr'), 'm:val'), '(', 'm:d 的左定界符');
    checkEqual(attr(first(first(delim, 'm:dPr'), 'm:endChr'), 'm:val'), ')', 'm:d 的右定界符');
    check(!!first(first(delim, 'm:e'), 'm:f'), '括号里应是分数');
    const brackets = descendants(mathOmaths[9], 'm:d').filter((node) => attr(first(first(node, 'm:dPr'), 'm:begChr'), 'm:val') === '[');
    checkEqual(brackets.length, 1, '\\left[\\right] 没有生成方括号 m:d');
}
// 9) 上下标同体：x_{i}^{2} → m:sSubSup(e,sub,sup)
{
    const subsup = first(mathOmaths[10], 'm:sSubSup');
    check(!!subsup, 'x_{i}^{2} 没有生成 m:sSubSup');
    checkEqual(children(subsup).map((node) => node.name).join(','), 'm:e,m:sub,m:sup', 'm:sSubSup 子元素顺序');
    checkEqual(mathText(first(subsup, 'm:e')), 'x', 'm:sSubSup 的基底');
    checkEqual(mathText(first(subsup, 'm:sub')), 'i', 'm:sSubSup 的下标');
    checkEqual(mathText(first(subsup, 'm:sup')), '2', 'm:sSubSup 的上标');
}
// 10) 符号表与希腊字母
{
    const symbols = mathText(mathOmaths[10]);
    for (const want of ['α', '×', 'β', '≤', 'γ', '≠', '∞', '≈']) {
        check(symbols.includes(want), `符号 ${want} 没有输出`, symbols);
    }
    const more = mathText(mathOmaths[11]);
    for (const want of ['∂', '∇', '∈', '⊂', '⊆', '∪', '∩', '→', '⇒']) {
        check(more.includes(want), `符号 ${want} 没有输出`, more);
    }
    const rest = mathText(mathOmaths[12]);
    for (const want of ['∀', '∃', '∠', '°', '⋯', '…', '′']) {
        check(rest.includes(want), `符号 ${want} 没有输出`, rest);
    }
}
// 11) \text 中文走 m:nor + w:rFonts(eastAsia)；\mathrm 正体、\mathbf 加粗
{
    const nor = descendants(mathOmaths[13], 'm:nor');
    checkEqual(nor.length, 2, '\\text{…} 应生成 m:nor');
    const textRun = descendants(mathOmaths[13], 'm:r').find((node) => descendant(node, 'm:nor'));
    const textFonts = first(first(textRun, 'w:rPr'), 'w:rFonts');
    checkEqual(attr(textFonts, 'w:eastAsia'), '等线', '\\text 里的中文必须给 eastAsia 字体');
    const styles = descendants(mathOmaths[13], 'm:sty').map((node) => attr(node, 'm:val'));
    check(styles.includes('p'), '\\mathrm{} 没有把内容改成正体');
    check(styles.includes('b'), '\\mathbf{} 没有把内容改成加粗');
}
// 12) 未知命令：原样输出 + 只告警一次；符号表查找不能命中 Object.prototype 上的成员
{
    const text = mathText(mathOmaths[14]);
    checkEqual(text, 'a\\foobarb\\constructorc\\foobard', '未知命令应原样输出');
    checkEqual(formulaReport.warnings.filter((line) => line.includes('\\foobar')).length, 1, '未知命令应按命令去重，只报一次');
    checkEqual(formulaReport.warnings.filter((line) => line.includes('\\constructor')).length, 1, '\\constructor 应被当成未知命令，而不是命中原型链上的成员');
}
// 12b) 显式空格：\, \; \: \quad \qquad 各自对应一个不可见空格字符，\! 是负空格（吞掉）
{
    checkEqual(mathText(mathOmaths[15]), 'a\u2005b\u2005c\u2003d\u2003\u2003ef\u2009g', 'LaTeX 空格命令没有映射成对应的空格字符');
}
// 13) 行内公式混排
{
    const inlineParas = descendants(formulaBody, 'w:p').filter((node) => descendants(node, 'm:oMath').length && !descendant(node, 'm:oMathPara'));
    checkEqual(inlineParas.length, 2, '含行内公式的段落数（一段混排 + 一段纯公式）');
    const names = children(inlineParas[0]).map((node) => node.name).join(',');
    checkEqual(names, 'w:pPr,w:r,m:oMath,w:r,m:oMath,w:r', '行内公式必须与普通 run 交替排列');
    checkEqual(children(inlineParas[1]).map((node) => node.name).join(','), 'w:pPr,m:oMath', '只有行内公式的段落应直接挂 m:oMath');
    check(formulaReport.outline.some((line) => /正文 \d+ 字（含 2 个行内公式）/.test(line)), 'outline 没有标出行内公式数', formulaReport.outline.join('|'));
    check(formulaReport.outline.includes('公式 E=mc2'), 'outline 缺少块级公式行', formulaReport.outline.join('|'));
    check(formulaReport.outline.includes('公式（行内） k=1nk'), 'outline 缺少行内公式行', formulaReport.outline.join('|'));
}
// read() 必须能从已有文件里数出公式，且不因 m: 元素崩掉
{
    const formulaRead = read('word-formula.docx', env);
    checkEqual(formulaRead.stats.formulas, formulaReport.stats.formulas, 'read() 的 stats.formulas');
    checkEqual(JSON.stringify(formulaRead.outline), JSON.stringify(formulaReport.outline), 'read() 的 outline 应与 create() 一致');
}

// ---------------------------------------------------------------- 表格（合并单元格 / 题注）

const tableBuilder = buildTableSample(env);
const tableBytes = tableBuilder.render();
const tableReport = tableBuilder.save();
const tableParts = unzip(tableBytes);
const tableText = new TextDecoder().decode(tableParts.get('word/document.xml'));
const tableRoot = parseXml(tableText).children[0];
const tableBody = descendants(tableRoot, 'w:body')[0];
const sampleTables = descendants(tableBody, 'w:tbl');

console.log('— 表格样例 —');
console.log(JSON.stringify({ path: tableReport.path, stats: tableReport.stats, warnings: tableReport.warnings }, null, 2));
for (const line of tableReport.outline) console.log(`  ${line}`);

checkSchemaOrder(new Map([['word/document.xml', tableText]]));
checkEqual(sampleTables.length, 2, '表格样例的表格数');
checkEqual(tableReport.stats.tables, 2, 'stats.tables');
// 默认主题 plain：合并单元格、题注、表头都不许引入任何底纹
check(!/<w:shd/.test(tableText), '默认主题的表格样例里出现了 w:shd 底纹');
check(!/w:fill=/.test(tableText), '默认主题的表格样例里出现了填充色');

// 表格 1：4 列 + 居中 + 题注在上
{
    const table = sampleTables[0];
    checkEqual(attr(first(first(table, 'w:tblPr'), 'w:jc'), 'w:val'), 'center', '表格 align:"center" 没有写 w:jc');
    checkEqual(descendants(table, 'w:gridCol').length, 4, '表格 1 的列数');
    checkEqual(descendants(table, 'w:tr').length, 5, '表格 1 的行数（1 表头 + 4 数据）');
    checkEqual(descendants(table, 'w:tblHeader').length, 1, '表格 1 默认应重复表头');
    checkEqual(attr(first(first(table, 'w:tblPr'), 'w:tblW'), 'w:type'), 'dxa', '表格宽度类型');

    // colspan → w:gridSpan，且宽度是各列之和
    const spanCell = cellWithText(table, '营收（万元）');
    check(!!spanCell, '找不到跨列单元格');
    checkEqual(attr(first(cellPr(spanCell), 'w:gridSpan'), 'w:val'), '2', 'colspan=2 应写成 w:gridSpan');
    const grid = descendants(table, 'w:gridCol').map((node) => Number(attr(node, 'w:w')));
    checkEqual(Number(attr(first(cellPr(spanCell), 'w:tcW'), 'w:w')), grid[0] + grid[1], '跨列单元格的宽度应为两列之和');

    // rowspan → 起始格 vMerge="restart"，下一行同列补 <w:vMerge/> 占位格
    const startCell = cellWithText(table, '华东区');
    checkEqual(attr(first(cellPr(startCell), 'w:vMerge'), 'w:val'), 'restart', 'rowspan 起始格应为 w:vMerge w:val="restart"');
    const rows = children(table, 'w:tr');
    const placeholder = children(rows[3], 'w:tc')[0];
    check(!!first(cellPr(placeholder), 'w:vMerge'), 'rowspan 的下一行缺少 w:vMerge 占位格');
    checkEqual(attr(first(cellPr(placeholder), 'w:vMerge'), 'w:val'), undefined, '占位格必须是 <w:vMerge/>（不带 val）');
    checkEqual(attr(first(cellPr(placeholder), 'w:tcW'), 'w:w'), attr(first(cellPr(startCell), 'w:tcW'), 'w:w'), '占位格的宽度应与起始格一致');
    checkEqual(textOfNode(placeholder), '', '占位格必须是空的');
    // 占位格必须落在被合并的那一列上：它后面紧跟的应是原本第 2 列的「590」
    checkEqual(textOfNode(children(rows[3], 'w:tc')[1]), '590', '占位格后的第一个单元格应是原第 2 列');

    // 列对齐：第 2 列 right、第 3 列 center
    checkEqual(attr(descendants(cellWithText(table, '590'), 'w:jc')[0], 'w:val'), 'right', '列的 align:"right" 没有生效');
    checkEqual(attr(descendants(cellWithText(table, '560'), 'w:jc')[0], 'w:val'), 'center', '列的 align:"center" 没有生效');
    // 单元格显式 align 覆盖列对齐
    checkEqual(attr(descendants(cellWithText(table, '1,410'), 'w:jc')[0], 'w:val'), 'right', '单元格 align 没有生效');

    // 越界裁剪：colspan=9 被裁到 4 列，多出来的单元格被丢弃并写 warning
    const clipped = cellWithText(table, '越界的 colspan');
    checkEqual(attr(first(cellPr(clipped), 'w:gridSpan'), 'w:val'), '4', '越界的 colspan 应被裁剪到剩余列数');
    checkEqual(children(rows[4], 'w:tc').length, 1, '越界行多出来的单元格应被丢弃');
    const tableWarnings = tableReport.warnings.join('\n');
    check(/colspan=9 越界/.test(tableWarnings), '缺少 colspan 越界告警', tableWarnings);
    check(/单元格超出 4 列/.test(tableWarnings), '缺少单元格超出列数告警', tableWarnings);
}

// 表格 2：题注在下 + 不重复表头
{
    const table = sampleTables[1];
    checkEqual(descendants(table, 'w:tblHeader').length, 0, 'headerRepeat:false 时不应写 w:tblHeader');
}

// 题注位置：表 1 在 <w:tbl> 之前，表 2 在 </w:tbl> 之后，且都用 Caption 样式
{
    const firstTableAt = tableText.indexOf('<w:tbl>');
    const lastTableEnd = tableText.lastIndexOf('</w:tbl>');
    check(tableText.indexOf('表 1 合并单元格与列对齐') > 0 && tableText.indexOf('表 1 合并单元格与列对齐') < firstTableAt, 'captionPosition:"above" 的题注应排在表格之前');
    check(tableText.indexOf('表 2 表题在表格下方') > lastTableEnd, 'captionPosition:"below" 的题注应排在表格之后');
    const captionParas = descendants(tableBody, 'w:p').filter((node) => attr(first(first(node, 'w:pPr'), 'w:pStyle'), 'w:val') === 'Caption');
    checkEqual(captionParas.length, 2, '题注段落数（都用 Caption 样式）');
    const outlineText2 = tableReport.outline.join('|');
    check(outlineText2.includes('题注 表 1 合并单元格与列对齐（题注在表格上方）'), 'outline 缺少表 1 题注', outlineText2);
    check(outlineText2.includes('题注 表 2 表题在表格下方，且不重复表头'), 'outline 缺少表 2 题注', outlineText2);
    const outlineLines = tableReport.outline;
    check(outlineLines.findIndex((line) => line.startsWith('题注 表 1')) < outlineLines.findIndex((line) => line.startsWith('表格 ')), 'outline 里表 1 的题注应排在表格行之前');
    check(outlineLines.findIndex((line) => line.startsWith('题注 表 2')) > outlineLines.findIndex((line) => line.startsWith('表格 3×2')), 'outline 里表 2 的题注应排在表格行之后');
}

// ---------------------------------------------------------------- edit 追加公式 / 合并表格

const appendReport = edit('word-table.docx', [
    { append: { type: 'formula', latex: '\\frac{\\partial f}{\\partial x}', number: '（2）' } },
    { append: { type: 'table', columns: ['甲', '乙'], rows: [[{ text: '合并', colspan: 2 }]], caption: '表 3 追加的表格', headerRepeat: false } },
    { append: { type: 'formula', latex: '' } },
], env);

console.log('— edit(追加公式) —');
console.log(JSON.stringify({ applied: appendReport.applied, skipped: appendReport.skipped, stats: appendReport.stats }, null, 2));

checkEqual(appendReport.applied.length, 2, '追加公式/表格的 applied 数');
checkEqual(appendReport.skipped.length, 1, '空 latex 应记入 skipped');
check(/append\.latex 为空/.test(appendReport.skipped[0].reason), 'skipped 应说明 latex 为空');
checkEqual(appendReport.stats.formulas, 1, '追加公式后的 stats.formulas');
checkEqual(appendReport.stats.tables, 3, '追加表格后的 stats.tables');
{
    const appendedDoc = new TextDecoder().decode(unzip(env.readFile('word-table.docx')).get('word/document.xml'));
    check(/<m:oMathPara>/.test(appendedDoc), '追加的公式没有写进文档');
    check(/<m:f>/.test(appendedDoc), '追加的公式里没有分数结构');
    check(appendReport.outline.includes('题注 表 3 追加的表格'), '追加表格的题注没有出现在 outline 里');
}

// 外部来的 .docx 常常没有声明 xmlns:m：edit() 追加公式时必须自己补上，否则整份文档非法
{
    const raw = unzip(env.readFile('word-table.docx'));
    const stripped = new TextDecoder().decode(raw.get('word/document.xml')).replace(/\sxmlns:m="[^"]*"/, '');
    check(!/xmlns:m=/.test(stripped), '样例文档本来应该没有 xmlns:m 才测得到补齐逻辑');
    raw.set('word/document.xml', new TextEncoder().encode(stripped));
    env.writeFile('word-external.docx', zip(raw));
    const externalReport = edit('word-external.docx', [{ append: { type: 'formula', latex: 'x^2' } }], env);
    checkEqual(externalReport.applied.length, 1, '外部文档追加公式的 applied 数');
    const externalDoc = new TextDecoder().decode(unzip(env.readFile('word-external.docx')).get('word/document.xml'));
    check(/xmlns:m=/.test(externalDoc), 'edit() 没有给外部文档补上 xmlns:m 声明');
    check(/<m:oMath>/.test(externalDoc), '外部文档里没有写入 m:oMath');
    checkEqual(externalReport.stats.formulas, appendReport.stats.formulas + 1, '外部文档追加公式后的公式数');
}

// ---------------------------------------------------------------- 结果

if (!KEEP) {
    // .office/tmp 是与 Excel/PPT 引擎共用的目录，只能删自己产出的文件，
    // 绝不能整个 rmSync —— 那会把别的引擎正在跑的产物一起清掉
    for (const name of [SAMPLE, 'word-draft.docx', 'word-formula.docx', 'word-table.docx', 'word-external.docx']) env.remove(name);
    env.remove('assets/sample.png');
    if (env.list('assets').length === 0) env.remove('assets');
    console.log('已清理本次产物（保留 .office/tmp 下其他文件）');
} else {
    console.log(`保留产物：${ROOT}`);
}

if (failures.length) {
    console.error(`\nFAIL word —— ${failures.length}/${checks} 项断言未通过：`);
    for (const item of failures) console.error(`  ✗ ${item}`);
    process.exitCode = 1;
} else {
    console.log(`\nPASS word bytes=${bytes.length} parts=${parts.size} warnings=${report.warnings.length} checks=${checks}`);
}
