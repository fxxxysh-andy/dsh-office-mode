# 格式模块契约（Word / Excel / PPT / LaTeX）

本文件是 `src/formats/` 下四个格式引擎的唯一接口约定。实现模块前先读它，
写完先跑自己的测试再交付。**除本文件与 `src/engine/` 里的现有文件外，
不要修改任何共享文件**；需要共享文件新增能力时，在交付说明里提出，不要自己改。

## 0. 共同前提

- ESM，零第三方依赖。可用 `node:fs` / `node:path` / `node:zlib` 等标准库。
- 只通过 `src/engine/zip.js`、`src/engine/xml.js`、`src/engine/kit.js`、
  `src/engine/theme.js` 里的能力读写 OOXML。先读这四个文件的实际导出。
- 生成结果必须能被真实 Microsoft Office 打开（交付前用 COM 验证一次）。
- 代码注释用中文，风格与 `src/engine/*.js` 一致：解释「为什么」，不复述代码。
- 不做垃圾文件：测试产物写到 `test/.tmp/`，验证完自己删掉。

## 1. 模块导出

三个模块各自导出完全相同的形状（`src/formats/<id>.js`）：

```js
export const meta = {
    id: 'word',                       // office.<id>
    name: 'Word 文档',                 // 中文名
    ext: '.docx',
    summary: '一句话说明能做什么',       // 进 office_help 索引
    methods: {                        // 进 office_help 详情（字符串数组）
        create: ['create(spec) → builder', '  builder.title(text) ...'],
    },
};

/**
 * 建新文档。env 见 §2。
 * 返回 builder：所有链式方法返回 this，另有：
 *   render()            → Uint8Array        （纯函数，不写盘）
 *   save(path?)         → report            （写盘并返回报告，path 省略时用 spec.path）
 */
export function create(spec, env) { ... }

/** 读已有文件，返回报告（见 §3）。path 相对 env.root。 */
export function read(path, env) { ... }

/**
 * 批量编辑已有文件。ops 是操作数组，一次调用做完所有修改后只写一次盘。
 * 返回 report（含 applied / skipped 明细）。
 */
export function edit(path, ops, env) { ... }

/**
 * 可选：格式额外的方法（不导出就等于没有）。
 * 每个值是「拿 env、返回一个函数」的工厂，SDK 的 wrap() 会把结果挂到
 * office.<id> 上，与 create / read / edit 并列。
 *
 * 为什么用工厂而不是直接给函数：这些方法需要 env，而 env 在每次 office_run
 * 里才存在；工厂让格式模块不必知道「当前的 env 是哪一个」。
 * 惰性 import 由格式模块自己决定（ppt.js 就这么做，避免与 ppt-revise.js 循环依赖）。
 *
 * 现有实现：ppt.js 用它挂 readSlides / revise。
 */
export const api = {
    readSlides: (env) => (path) => { ... },
    revise: (env) => (path, ops) => { ... },
};
```

**挂载面必须有断言。** 新方法挂不上时插件照常加载、其它格式毫发无损，
脚本只会拿到 `undefined is not a function` —— 这种静默缺失防不住，
只能靠测试钉住（见 `test/smoke.mjs` 的「office.ppt 暴露 readSlides / revise」）。

## 2. env（由 `createEnv` 提供，已实现）

```
env.root            会话工作目录（绝对路径）
env.resolve(p)      相对 → 绝对
env.theme(id)       → { theme, fellBack, requested? }，见 engine/theme.js
env.writeFile(p, data)  → { path, absolute, bytes }   // 自动建目录，相对 root
env.readFile(p)     → Buffer
env.readText(p)     → string
env.exists(p)       → boolean
env.stat(p)         → fs.Stats | undefined
env.list(p)         → [{name,dir,bytes}]
env.remove(p)       → boolean
env.log(msg) env.warn(msg) env.note(msg)   // 进反馈
```

**不要**直接 `fs.writeFileSync`。所有落盘走 `env.writeFile`，这样反馈里的
文件清单、体积、缓存统计才准确。

## 3. 报告结构（三种格式统一）

```js
{
  ok: true,
  format: 'docx' | 'xlsx' | 'pptx',
  path: '报告.docx',            // 相对 root，正斜杠
  bytes: 12345,
  theme: 'plain',               // 默认 plain（素色网格：Excel/Word 都不加背景填充）
  stats: { /* 该格式的准确数字 */ },
  outline: [ /* 结构摘要，每项一行字符串，给模型观察用 */ ],
  warnings: [ /* 排版风险，如「第 3 页文字可能溢出」 */ ],
}
```

`stats` / `outline` / `warnings` 是「准确反馈」的核心：模型只靠这份报告就应该
知道文件里到底有什么、哪里可能不好看，而不必再读一遍文件。

LaTeX（`tex.js`）沿用同一形状，另加 `kind: 'main' | 'fragment'` 与 `files` / `support`
等字段；`theme` 恒为 `null`（版式由模板决定，与视觉主题无关）。

## 4. 各自的 API 与 OOXML 要求

### 4.1 Word（`src/formats/word.js`，.docx）

**builder 方法**（`spec` 支持 `{title, theme, path, page, orientation, margins, header, footer, author, subject}`）：

| 方法 | 说明 |
| --- | --- |
| `title(text, opts)` | 封面式大标题 |
| `heading(text, level=1, opts)` | 标题 1–3 级 |
| `para(text, opts)` | 正文；`text` 可为 string 或 run 数组 `[{text,bold,italic,color,size,math}]`，带 `math` 的 run 渲染成行内 OMML 公式 |
| `bullets(items, opts)` / `steps(items, opts)` | 无序 / 有序列表，`items` 为 string 或 run 数组 |
| `quote(text, opts)` / `code(text)` | 引用 / 等宽代码块 |
| `formula(latex, opts)` | 块级数学公式（OMML），`opts: {align:'center'\|'left', number, size}`；`number` 为右侧编号（制表位实现） |
| `table({columns, rows, widths, style, caption, captionPosition, headerRepeat, align})` | `columns` 支持 `{title, width, align}`；单元格可为 string 或 `{text, colspan, rowspan, align, bold, italic, fill, color, size}`；`caption` 写题注段落 |
| `image(path, {widthCm, caption})` | 嵌入图片（png/jpeg），带题注 |
| `pageBreak()` / `spacer(cm)` / `toc()` | 分页 / 竖向留白 / 目录字段 |
| `save(path?)` / `render()` | 见 §1 |

`opts` 支持：`{bold, italic, underline, color, size, font, align, indent, spaceBefore, spaceAfter, lineSpacing}`。

**包内必需部件**（`[Content_Types].xml` 必须为每个 part 写 Override）：
`[Content_Types].xml`、`_rels/.rels`、`word/document.xml`、
`word/_rels/document.xml.rels`、`word/styles.xml`、`word/numbering.xml`、
`word/settings.xml`、`word/footer1.xml`（有页脚时）、`docProps/core.xml`、`docProps/app.xml`。

要点：
- 段落样式靠 `w:pStyle`，`styles.xml` 里 `w:styleId` 用 `Normal/Title/Heading1..3/Quote/ListParagraph/Code/TableGrid`，
  并且 `w:name w:val` 必须写内置名（`heading 1` 等），否则 Word 不认成标题。
- 字体必须同时给 `w:ascii`/`w:hAnsi`（西文）与 `w:eastAsia`（中文），否则中文回落到默认宋体。
- 项目符号：`numbering.xml` 里 `w:abstractNum` + `w:num`，段落用 `<w:numPr><w:ilvl/><w:numId/></w:numPr>`。
- 表格：`w:tblGrid` 列宽 + 每格 `w:tcW`（dxa/twips）+ `w:tblHeader`，表格边框写进 `w:tblBorders`。
  底色一律来自主题：`table.headerFill` / `table.zebraFill` 非空时才写 `w:shd w:fill`，
  默认主题 `plain` 两者都是空串，于是只靠加粗与边框区分（**不要写 `w:fill="FFFFFF"`**，
  那等于给单元格加了一层白色背景）。
- 页面：A4 = 11906×16838 twips，页边距用 `w:pgMar`；横向时交换 `w:pgSz` 的 w/h 并加 `w:orient="landscape"`。
- 页脚页码用 `PAGE` 域（`w:fldChar begin` / `w:instrText` / `end`）。
- **不写任何「打开时更新域」标记**：`word/settings.xml` 里不出现 `<w:updateFields>`，
  域字符上也不出现 `w:dirty`。两者各自都会让 Word 弹「该文档包含的域可能引用了其他文件」。
  目录（`builder.toc()`）改成把 H1–H3 的条目与页码**预渲染**进 TOC 域的缓存结果
  （`renderToc`），条目各占一段、用 Word 内置样式 id `TOC1`–`TOC3`，段落里带右对齐点线
  制表位；标题上用 `w:bookmarkStart/End` 给锚点，目录条目用 `w:hyperlink w:anchor` 跳转。
  页码来自引擎的分页估算，所以 `create()` 带目录时渲染两遍（第一遍估页码、第二遍写进目录）。
  详见 [docs/ooxml-pitfalls.md](docs/ooxml-pitfalls.md) 第 8 条。
- 中文正文默认 10.5pt（`w:sz` = 21 half-points），标题 16–22pt。

`stats` 至少含：`paragraphs, headings, tables, images, words, pagesEstimate`。
`warnings` 至少覆盖：正文段落过长（>500 字未分段）、表格列宽合计超出版心、
标题层级跳跃（H3 直接跟在 H1 后）。

**read(path, env)** 解析 `word/document.xml`，`outline` 每段形如
`"H1 公司简介"`、`"正文 120 字"`、`"表格 4×6"`；`stats` 同上。

**edit(path, ops, env)** 支持的操作（未知操作记入 `skipped` 并给原因，不要抛）：
```js
[{ find: '旧文字', replace: '新文字', all: true }]      // 跨 run 的文本替换
[{ setStyle: { match: /^旧标题/, style: 'Heading2' } }]
[{ append: { type: 'para', text: '追加内容' } }]        // 支持 para/heading/bullets/table
[{ setFooter: '第 X 页' }]                              // 可选
```
一次调用内所有操作作用在同一个包上，最后只写一次盘。返回 `{applied, skipped, changes}`。

### 4.2 Excel（`src/formats/excel.js`，.xlsx）

**workbook builder**：`create({theme, path, author, title}, env)` → `wb`
- `wb.sheet(name, opts?)` → sheet builder（同名取第二次时自动加序号）
- `wb.save(path?)` / `wb.render()`

**sheet builder 方法**：

| 方法 | 说明 |
| --- | --- |
| `title(text, {span})` | 大标题行，合并单元格 |
| `note(text)` | 灰字说明行 |
| `header(columns)` / `table({columns, rows, startRow, totalRow, zebra, freeze})` | 主数据表 |
| `row(values, opts)` | 单行追加，`opts: {bold, fill, color, height, align}` |
| `cell(ref, value, opts)` | 指定地址；`opts: {bold, italic, fill, color, align, format, size, border}` |
| `formula(ref, formula)` | 公式，`formula` 不含前导 `=`（只写 `<f>`，由 Excel 打开时重算） |
| `stats(source, opts)` | 统计块：`source` 为 `'B5:B8'` / `{column, from, to}`；`opts: {at, label, funcs, layout:'rows'\|'columns', format, header}`。数值格写**真公式 + 缓存值** |
| `summary(opts)` | 分组汇总（轻量透视）：`opts: {key, value, at, funcs, label, total, format}`，写 `SUMIF/COUNTIF/AVERAGEIF` + 缓存值 |
| `merge(range)` / `widths({A:12, B:24})` / `freeze(row, col)` / `autofilter(range?)` |
| `numberFormat(range, fmt)` | 如 `#,##0.00`、`0.0%`、`yyyy-mm-dd` |

**包内必需部件**：`[Content_Types].xml`、`_rels/.rels`、`xl/workbook.xml`、
`xl/_rels/workbook.xml.rels`、`xl/worksheets/sheetN.xml`、`xl/styles.xml`、
`docProps/core.xml`、`docProps/app.xml`。
字符串一律用 `t="inlineStr"` + `<is><t>`（不建 sharedStrings，少一个部件少一处错）。
`workbook.xml` 里加 `<calcPr calcId="0" fullCalcOnLoad="1"/>`，公式才会在 Excel 打开时重算。

要点：
- `styles.xml` 的 `cellXfs` 下标就是单元格的 `s` 属性；fills 的 0/1 固定为 `none`/`gray125`。
- 颜色写 `rgb="FFRRGGBB"`（ARGB，必须带 FF 不透明前缀）。
- 列宽 `width` 单位是「字符数」，中文列要按 `估宽 × 1.2` 放宽；`<cols><col min max width customWidth="1"/>`。
- 行高 `ht` 单位磅；`<row r="1" ht="22" customHeight="1">`。
- 元素顺序（schema 强制）：`sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData, autoFilter, mergeCells`。
- 冻结窗格：`<pane xSplit ySplit topLeftCell activePane state="frozen"/>`，`activePane` 要与 split 对应。
- 数字单元 `t="n"` + `<v>`；日期用序列号（1900 日期系统）+ 日期 numFmt。

`stats` 至少含：`sheets, rows, columns, cells, formulas, merges, bytes`，每张表给 `{name, rows, columns, range}`。
`warnings` 至少覆盖：列宽明显不够（中文表头被截断）、合并区域与数据区域重叠、
公式引用了不存在的表名、`#REF!`/`#DIV/0!` 之类的字面量错误。

**read(path, env)**：逐表返回 `{name, range, rows: 前 5 行样例, formulas: [...]}`。
**edit(path, ops, env)**：支持 `{sheet, cell, value}`、`{sheet, row:{at, values}}`、
`{sheet, find, replace}`、`{sheet, addSheet: name}`。同样一次写盘。

### 4.3 PPT（`src/formats/ppt.js`，.pptx）

**deck builder**：`create({title, theme, path, size='16:9', author}, env)` → `deck`
- `deck.cover({title, subtitle, kicker, presenter, date})`
- `deck.section({title, subtitle})`
- `deck.bullets({title, items, columns=1})`，`items` 为 string 或 `{text, level}`
- `deck.table({title, columns, rows, headerFill, zebra, border, cellPadCm, rowHeightCm, firstRowBold})`
  单元格为 string 或 `{text, runs, colSpan, rowSpan, fill, color, bold, align, link}`；
  合并写 `a:tc` 的 `gridSpan`/`rowSpan` 与覆盖格的 `hMerge="1"`/`vMerge="1"`（见 docs/ooxml-pitfalls.md）
- `deck.quote({text, by})` / `deck.statement({text, sub})`
- `deck.image({path, title, caption, fit, widthCm, heightCm, align, frame, fullBleed})`
  `fit`: `contain`（默认，等比放进内容区）/ `cover`（铺满内容区，超出用 `a:srcRect` 裁切）/ `natural`（按 96 DPI 原始像素折算）；
  `align`: `left|center|right`；`frame` 加细边框；`fullBleed` 铺满整页（可当背景图）
- `deck.images(items, {title, columns=2, gapCm, fit})` 一页多图网格，`items: [{path, caption}]`
- **绘制层**：`deck.shape({preset, x, y, w, h, fill, line, rotate, radius, shadow, text, textOpts})`
  （32 个 preset；`fill` 支持纯色 / `{color, alpha}` / `{type:'gradient', stops, angle}` / `'none'`；
  `line` 支持 `dash` 与 `arrow:{begin,end}`）、`deck.line({from, to, …})`、`deck.icon({glyph, x, y, sizeCm, color, font})`；
  连续调用画在同一页，`deck.page()` 另起一页
- **版式助手**（每次调用成一整页，自动吃母版叠加层）：
  `deck.cards({title, items, columns, cardFill, cardLine, radius, gapCm, numbered, iconSizeCm})`、
  `deck.steps({title, items, direction, numbered, arrows})`、`deck.compare({title, left, right, divider})`、
  `deck.kpi({title, items, columns})`、`deck.imageText({image, title, items, text, side, ratio})`、
  `deck.timeline({title, items, direction})`、`deck.iconGrid({title, items, columns})`
- **元素级助手**（画在当前页，与 `deck.shape` 同级）：`deck.panel({title, text, runs, x, y, w, h, fill, line, radius, sizePt})`、
  `deck.banner({text, sub, x, y, w, h, fill, gradient, color, sizePt})`
  助手页与自由页共用同一条画布通道：`cards()` 后接 `banner()/panel()/shape()` 落在同一页，`deck.page()` 才换页
- **母版层**：`deck.master({background, logo, header, footer, pageNumber, accent, skipLayouts})`、
  `deck.background(spec)`（只改当前页）；图层顺序 = 背景 → 内容 → logo/页眉/页脚/页码/装饰条
- **富文本**：任意段落支持 `runs:[{text, bold, italic, underline, color, size, font, spacing}]`
- `deck.closing({title, subtitle})`
- `deck.notes(text)` 给上一页加备注
- `deck.save(path?)` / `deck.render()`

尺寸：16:9 = `12192000 × 6858000` EMU；4:3 = `9144000 × 6858000`。
页边距统一 0.9cm 左右。

**图片的读与写**（`office.ppt.images` / `insertImages` / `readSlides`）：

- `images(path, {out}, env, extra)` → `{ok, format, path, bytes, count, dir, files[], skipped[], reused, hint}`；
  `files[] = {part, name, ext, contentType, bytes, width, height, slides[], relIds[], path, reused}`。
  `out` 省略时写进注入的 cache（`ppt-images/<词干>/`），与 `office.pdf.pages` 同一套口径
  （中间产物不进工作目录，返回路径直接交给 `read_image`）。内容类型同时查
  `<Default Extension>` 与 `<Override PartName>`，两者都没有就进 `skipped`，不猜 MIME。
  `a:blip` 上是 `r:link`（外链）而不是 `r:embed` 的，也只进 `skipped`，不下载不编造字节。
- `insertImages(path, items, env)` → `{ok, format, path, saved, bytes, applied[], skipped[], warnings[]}`；
  `items: [{slide, path, x, y, wCm, hCm, fit, name, alt}]`，`fit ∈ contain|cover|natural`，
  坐标单位厘米、原点左上（与 `deck.shape` / `revise.move` 同一套）。无改动不写盘
  （`ok:false, saved:false`），与 `revise()` 同一约定。
- `revise(path, ops)` 新增 op `{slide, addImage: {…}}`：在 op 循环里于 shape matcher
  之前分流到包级插入器（`applyOp()` 只改单个形状的 XML 切片，做不到包级改动）。
- 插入是**改包不是重建包**：追加 image 关系（id = 该页最大数字 id + 1，Target 为
  `../media/<part>`）、仅在缺 Default 时补 `<Default Extension="png|jpeg|gif" …/>`、
  `ppt/media/imageN.<ext>` 从包内已有编号往后顺延、`p:pic` 追加为 `p:spTree` 最后一个
  形状（末尾有 `p:extLst` 则插在它之前）。`p:nvPicPr` 里 `<p:nvPr/>` 必填，
  `p:spPr` 子元素顺序 xfrm → prstGeom → 填充 → ln（见 docs/ooxml-pitfalls.md 第 4、8 条）。
  EMU = 厘米 × 360000；`natural` 用 `px × 9525`（96 DPI）。
- 读路径：`read()` 的 `stats.media` 给出 `{part, contentType, bytes, width, height, slides[]}`；
  `readSlides()` 的 `p:pic` 形状新增 `image: {relId, part, linked, contentType, bytes, width, height, crop}`。
  **`p:pic` 的 `editable` 仍是 `false`**（图片没有文本可改），但几何可以 move / resize ——
  它自带 `p:spPr/a:xfrm`。
- `src/sdk.js` 的 `wrap()` 现在把 `{cache, config}` 作为第二参传给 `module.api` 的工厂
  （`api[name] = fn(env, extra)`）；tex/word/excel 的工厂忽略第二参，行为不变。

**包内必需部件**：
`[Content_Types].xml`（slide/layout/master/theme/notesSlide 都要 Override）、
`_rels/.rels`、`ppt/presentation.xml` + rels、
`ppt/slideMasters/slideMaster1.xml` + rels、
`ppt/slideLayouts/slideLayoutN.xml` + rels（每种版式一个，至少 6 个）、
`ppt/theme/theme1.xml`、`ppt/slides/slideN.xml` + rels、
有备注时 `ppt/notesSlides/notesSlideN.xml` + rels。

要点（PowerPoint 比 Word/Excel 严格得多，逐条照做）：
- `p:presentation` 子元素顺序：`sldMasterIdLst, sldIdLst, sldSz, notesSz, defaultTextStyle`。
- 关系 id 不能重复；`p:sldIdLst` 里 `p:sldId id` 从 256 起递增，`r:id` 指到 slide。
- `p:sldMasterId id` 固定 `2147483648`，`p:sldLayoutId id` 从 `2147483649` 起。
- 每个 slide/layout 的 `p:spTree` 前两个子元素必须是
  `p:nvGrpSpPr` 和 `p:grpSpPr`（grpSpPr 里的 xfrm 四个值都给 0）。
- 版式必须有 `p:clrMapOvr`→`a:masterClrMapping`；master 必须有 `p:clrMap`（12 个属性全给）。
- master 必须有 `p:txStyles`（`titleStyle`/`bodyStyle`/`otherStyle` 三块都要在）。
- `theme1.xml` 必须有 `a:clrScheme`(12 色)、`a:fontScheme`(major/minor 各 latin/ea/cs)、
  `a:fmtScheme`(fillStyleLst 3 项、lnStyleLst 3 项、effectStyleLst 3 项、bgFillStyleLst 3 项)。
  缺一项 PowerPoint 就报「需要修复」。
- `a:prstGeom` 的 `prst` 必须是 ST_ShapeType 的合法取值：自造名字自研解析器与 LibreOffice 都能渲染，
  真实 PowerPoint 却打不开且不报原因（三个已踩到的口语名映射见 docs/ooxml-pitfalls.md）。
- 封面与结束页的底色取主题的 `cover.bg`（缺省 = `primaryDark`）、文字取 `cover.text`（缺省 = `onPrimary`）：
  浅色主题若不覆盖，白字会落在白底上直接隐形。素色主题显式给出「白底黑字 + 浅灰色带」。
- 文本字号 `sz` 是百分之一磅（18pt → `sz="1800"`）。
- 简单形状不用占位符也可以：直接用 `p:sp` + `p:cNvSpPr txBox="1"` + `a:prstGeom prst="rect"`，
  这样能完全绕开 `p:ph` 的索引配套问题。**推荐全部用独立文本框**。
- 中文字体写 `a:ea typeface="微软雅黑"`，西文写 `a:latin`，并给 `lang="zh-CN"`。
- 走 `a:buChar`/`a:buAutoNum` 给项目符号，`a:spcBef`/`a:spcAft` 控行距。

`stats` 至少含：`slides, layouts:{cover:1,...}, shapes, notes, bytes`。
`outline` 每页一行：`"P1 cover：标题｜副标题"`、`"P3 bullets：4 条"`。
`warnings` 至少覆盖：单页要点超过 7 条、单条要点文字超过约 60 字而版心放不下、
一页里文本框面积超出版心 95%、项目符号层级 > 2。

**read(path, env)**：返回每页的版式类型、标题文本、正文行数与备注摘要。

### 4.4 LaTeX（`src/formats/tex.js`，.tex）

唯一「产物由外部程序算出来」的格式：源文件插件自己写，PDF 交给本机 TeX 发行版
（`src/tex.js`）。它也不是「一个文件」而是「一个项目」，所以三条接口的语义与
前三种格式不同 —— 这里逐条写清楚，改之前先读。

**create(spec, env)**：一次写出一整个可编译项目并返回报告。
`kind: 'project'`，`files` 列出**内容文件**（主文件 / thusetup / data/*.tex / refs.bib）
与**模板运行期文件**（thuthesis.cls、thu-fig-logo.pdf、thu-text-logo.pdf、
thuthesis-*.bst/bbx/cbx），`support` 说明模板从哪来。

- 主文件路径由 `spec.path` 定（默认 `thesis/thesis.tex`），**文件名必须是 ASCII**：
  中文文件名在 TeX 日志、aux 与 bibtex 里会变乱码；中文标题写 `title`。
- 正文按章分文件（`data/chap01.tex` …），主文件按 `layout.include` 只 `\input`
  实际存在的那些（本科生不写 committee / comments / resolution）。
- `blocks` 是内容模型：字符串 = 一段正文；对象有 `{p} {items,ordered} {equation,label}
  {align} {figure} {table} {code} {raw}`。文本按 **LaTeX 源码**处理（转义是调用方的事，
  需要纯文本时用 `office.tex.escape()`）；图与表在没给 label 时自动补 `fig:N-k` / `tab:N-k`。
- `\thusetup{}` 块里不能出现空行；关键词不超过 5 个；这些都要在 `warnings` 里点出来。

**read(path, env)**：同步。主文件走**项目级**解析（顺 `\input` 递归，深度上限 3），
片段文件（`data/chap01.tex`、`thusetup.tex` 这种）只解析自己并返回 `kind: 'fragment'` ——
不能拿整篇论文的标准去挑片段的毛病（否则 `office_run` 的自动复检会刷一屏假警告）。
`\appendix` 之后的 `\input` 视为附录：它的 `\chapter` 记进 `stats.appendixChapters`，
outline 里写「附录 …」，不占正文章号。

`stats` 至少含：`files, chapters, appendixChapters, sections, subsections, figures,
tables, equations, citations, citationKeys, references, words, degree, degreeType,
language, fontset, inputFiles`。
`warnings` 至少覆盖：图片文件不存在、`\input` 指向的文件不存在、`\cite` 的 key 在
`.bib` 里找不到、`.bib` 有条目却一处没引用、`\ref` 的 label 不存在、图/表缺 `\caption`
或 `\label`、单段超过 800 字、缺摘要 / 致谢 / 目录、标题超过 30 字。

**edit(path, ops, env)**：跨文件编辑，一次调用把所有改动落盘，返回
`{applied, skipped, changes, files}`。支持的操作：

```js
[{ find: '旧文字', replace: '新文字', all: true, file: 'data/chap01.tex' }]  // file 省略 = 全项目
[{ set: { title: '新标题', date: '2026-07-01' }, file: 'thusetup.tex' }]      // 改 \thusetup 字段
[{ appendTo: { file: 'data/chap02.tex', text: '追加一段' } }]
[{ appendChapter: { title: '实验', sections: [{ title: '设置', blocks: [...] }] } }]  // 新章 + 自动插 \input
[{ addReference: { key, type, author, title, year } }]                       // 同 key 不重复加
```

两条必须遵守的规矩（都踩过）：
1. **找文件一律按后缀匹配**（`data/chap02.tex` ↔ `thesis/data/chap02.tex`）。
   用 Map 的键直接查会查不到，然后「新建」一份只含追加内容的同名文件，把原文件覆盖掉。
2. **追加不是覆盖**：`.bib` 不在 `\input` 收集范围里，`addReference` 必须先把已有的
   `.bib` 读出来再追加，否则整份参考文献会被一条新条目替换掉。

**api（挂在 `office.tex` 上）**：`compile(path, {timeoutMs, clean, engine, support})`、
`engines()`、`template(dir?)`、`escape(text)`、`clearAux(path, mode)`。

编译只走 `src/tex.js`：探测（latexmk / xelatex / lualatex / bibtex / biber / kpsewhich）、
模板定位（配置 → 环境变量 → 工作目录里的 `thuthesis*` → TeX Live 自带）、
执行（`latexmk -xelatex`，没 latexmk 才手工多遍）、日志解析（`-file-line-error` 的
`文件:行号: 说明`、Overfull/Underfull 计数、`Output written on … (N pages)`）。
**不用管道**：stdout / stderr 重定向到文件句柄，并带 `cwd`（见 §2 与 src/pdf.js 的实测结论）。

`compile` 报告：`{ok, engine, main, projectDir, command, exitCode, timedOut, elapsedMs,
runs, pdf: {path, bytes, pages} | null, log: {path, errors, overfull, underfull,
missingChars, overfullSamples, tail}, support, hint, diagnosis}`。
`ok` 的判据是「PDF 真的出现了 且 解析不到 LaTeX 错误」——TeX 有时退出码非 0 但产物完好。

## 5. 自测要求（交付前必须跑通）

每个模块写一个 `test/format-<id>.mjs`，用 `node test/format-<id>.mjs` 可运行，内容：
1. 用 `createEnv({root: <test/.tmp>})` 造一个覆盖全部 builder 方法的样例；
2. `render()` 出字节 → `unzip()` 回读；
3. 用 `parseXml()` 解析每个 `.xml`/`.rels` part，确认根元素与关键节点存在；
4. 检查所有关系目标（rels 的 `Target`）在包内都存在（相对路径要按 part 目录解析）；
5. `read()` 回去，核对 `stats` 与预期一致；
6. 打印 `PASS <id> bytes=… parts=… warnings=…`，失败时打印具体差异并以非零码退出。

再跑一次真实 Office 校验（可选但强烈建议，Windows 上装了 Office）：

```powershell
$app = New-Object -ComObject PowerPoint.Application   # 或 Excel.Application / Word.Application
$doc = $app.Presentations.Open("D:\path\file.pptx", $true, $false, $false)
"$($doc.Slides.Count)"; $doc.Close(); $app.Quit()
```

COM 可能弹窗卡住：给命令 90 秒超时，超时就 `Stop-Process -Name POWERPNT,WINWORD,EXCEL -Force`
并如实报告「COM 校验未完成」，不要假装通过。

## 6. 交付说明格式

完成后回复一段话，包含：模块文件路径、`node test/format-<id>.mjs` 的实际输出、
COM 校验结果（通过 / 未完成+原因）、以及你为了绕开 OOXML 坑做的取舍。
