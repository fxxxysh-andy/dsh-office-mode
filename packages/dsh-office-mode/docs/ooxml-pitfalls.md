# OOXML 踩坑记录

这里记的都是**项目内自测全绿、真实办公软件却拒绝打开**的问题。它们的共同
特征是：自研解析器和 LibreOffice 都很宽容，只有 Microsoft Office 会判文件坏掉。
所以每次改这三个引擎，跑完单元测试后一定要再过一遍真实 Office 校验（见 README）。

## Excel（.xlsx）

### 1. `tabSelected` 是属性，不是子元素

```xml
<!-- 错：Office 判整个工作簿非法，LibreOffice 忽略后照常渲染 -->
<sheetView workbookViewId="0"><tabSelected val="1"/></sheetView>

<!-- 对 -->
<sheetView workbookViewId="0" tabSelected="1"/>
```

`CT_SheetView` 只允许 `pane` / `selection` / `pivotSelection` / `extLst` 四个子元素。
自测已经加了「sheetView 里没有非法子元素」的断言防回归。

### 2. 不要把自定义数据挂在 `docProps/app.xml` 的 `<Properties>` 下

扩展属性（extended-properties）的 schema 不允许 `Properties` 直接带 `vt:lpstr`
子元素。我们曾经为了省一个部件，把布局元数据当成裸的 `<vt:lpstr>` 塞在那里 ——
LibreOffice 忽略它照常出 PDF，Excel 直接拒绝打开。

自定义数据的正确去处是 `docProps/custom.xml`（标准自定义属性部件），并且要在
`[Content_Types].xml` 与 `_rels/.rels` 各注册一次：

```
ContentType: application/vnd.openxmlformats-officedocument.custom-properties+xml
RelType:     .../officeDocument/2006/relationships/custom-properties
```

### 3. `<color rgb="...">` 要 8 位 ARGB

主题里给的是 6 位 hex（`262626`）。少写不透明前缀时 Excel 会把它当另一种颜色，
统一用 `argb8()` 补齐成 `FF262626`。

## PowerPoint（.pptx）

### 4. `p:graphicFrame` 的 `p:nvGraphicFramePr` 里 `p:nvPr` 必填

PresentationML 里 `p:nvPr` 的 `minOccurs=1`，和 DrawingML 不同。漏掉它
PowerPoint 直接「无法打开该文件」。

### 5. 备注需要 notesMaster，否则静默丢弃

要让备注真的显示，必须写 `p:notesMasterIdLst` 并加一条
presentation → notesMaster 的关系；而且它必须排在 `p:sldIdLst` **之后**。
放错位置会被判坏文件（注意：是「静默丢失」而不是报错，所以很容易漏）。

### 6. `a:prstGeom` 的 preset 必须是 ST_ShapeType 的合法取值

自造名字时自研解析器与 LibreOffice 都照常渲染，**真实 PowerPoint 却打不开整个文件，
而且 COM 的 `Err.Description` 是空的**，只能靠「逐个 preset 单文件二分」定位。踩到的三个：

| 口语名（会打不开） | 合法取值 |
| --- | --- |
| `flowChartData` | `flowChartInputOutput` |
| `roundedRectCallout` | `wedgeRoundRectCallout` |
| `ovalCallout` | `wedgeEllipseCallout` |

引擎现在把口语名静默映射到真名（32 个 preset 已逐个用真实 PowerPoint 验证）。

同一类问题还有**所有枚举型属性**：`a:pPr/@algn` 只认 `l/ctr/r/just/justLow/dist/thaiDist`，
`a:bodyPr/@anchor` 只认 `t/ctr/b/just/dist`。曾经把用户写的 `align:'left'`、
`anchor:'middle'` 直接落成 `algn="left"` / `anchor="middle"`，结果同样是「LibreOffice 正常、
PowerPoint 拒绝打开且不给原因」。现在 `paragraphXml` / `textBodyXml` / 表格单元格三处
统一过 `textAlignOf()` / `textAnchorOf()` 收敛表，未知值回落默认值，绝不写非法枚举
（自测里有专门的回归断言）。

### 7. 表格合并属性挂在 `a:tc` 上，不在 `a:tcPr` 上

```xml
<!-- 对：gridSpan/rowSpan/hMerge/vMerge 都是 a:tc 的属性 -->
<a:tc gridSpan="2" rowSpan="2">…</a:tc>
<a:tc hMerge="1">…</a:tc>   <!-- 被横向覆盖的格 -->
```

用真实 PowerPoint 生成 2×2 合并与单行两列合并的 pptx 作 ground truth 核对过：
2×2 的对角格同时写 `hMerge="1" vMerge="1"`。**被覆盖的格必须是空 `a:txBody`**，
否则同一段文字会被数两遍。

### 8. `a:spPr` 与 `a:ln` 的子元素顺序

`a:spPr`：`a:xfrm` → `a:prstGeom`(含 `a:avLst`) → 填充 → `a:ln` → `a:effectLst`。
`a:ln`：填充 → `a:prstDash` → `a:headEnd` → `a:tailEnd`。顺序写反同样会被判坏文件。

### 9. 外部超链接的关系必须写 `TargetMode="External"`

`a:hlinkClick r:id` 指向的关系若漏了 `TargetMode="External"`，
PowerPoint 会去包内找目标，找不到就报错。

## Word（.docx）

### 6. `w:pPr` 子元素顺序被 schema 钉死

`CT_PPrBase` 要求 `w:pBdr`/`w:shd` 在 `w:spacing` 之前、`contextualSpacing`
在 `w:ind` 之后。手写字符串极易写反，所以 `word.js` 里所有 `w:pPr`/`w:rPr`
都收敛到 `pPrXml`/`rPrXml` 生成。

### 7. 标签名不要拼出双冒号

`borderXml` 曾经写成 `` `<w:${name}>` `` 而调用方传 `'w:bottom'`，产出
`<w:w:bottom>`。任何 XML 解析器都会直接拒绝这个文档。

### 8. 「打开时更新域」的两种写法都会让 Word 弹窗

Word 打开带域（PAGE / NUMPAGES / TOC …）的文档时，只要文档里有「打开时更新域」
的意图，就会弹：

> 该文档包含的域可能引用了其他文件。是否更新该文档中的这些域？

这条弹窗有两种独立触发源，**去掉任一处仍然弹**（本机真实 Word 实测）：

| 写法 | 位置 |
| --- | --- |
| `<w:updateFields w:val="true"/>` | `word/settings.xml` |
| `<w:fldChar w:fldCharType="begin" w:dirty="true"/>` | 具体域的 begin 字符 |

触发条件是「打开时更新域」这个意图本身，**与文档里放的是 PAGE 域还是 TOC 域无关**
（一个只有页脚页码域的文档，只要带 `w:updateFields` 就会弹）。

所以 `word.js` 两者都不写，目录改成把条目与页码**预渲染**进域的缓存结果
（`renderToc`）：打开即见完整目录、条目可点击跳转，按 F9 仍可让 Word 重算。
代价是页码来自引擎的分页估算（`create()` 带目录时渲染两遍：第一遍估页码、
第二遍写进目录），而不是 Word 的精确排版。

复现与判定手法：`cscript` 打开文件时把 `DisplayAlerts` 设成 `-1`（全开），
**若 `Documents.Open` 在超时内不返回，就是被模态弹窗挡住了**；对照组去掉标记后
同一份文档秒开。只看 `DisplayAlerts = 0`（全关）会漏掉这条 —— 弹窗被抑制，
但更新行为与页码仍然不对。

## 引擎内部：不判坏文件、但会静默出错的地方

下面这些不会让 Office 拒绝打开，只会让产出「看着对、其实不对」，所以单独立一节。

### 10. 同一个字段名在不同版式里含义不同

自由绘制页的叠加层最初复用了 `items` 字段，结果 `bullets({items:[…]})` 的要点数组被当成
自由形状又画了一遍（主样例立刻多出两条「形状缺少宽高」告警）。叠加层必须走自己的字段
（现在叫 `canvas`）。

### 11. 助手页在 read() 里只能靠形状名识别

包内版式部件只有 8 个，助手页都借用「陈述」版式，slideLayout 的名字分辨不出来。数条目必须
用精确正则（`^Card \d+$`）：用前缀匹配会把 `Card 1 Rule/Title/Body` 也算进去，4 张卡被数成 21 张。

### 12. 主题 surface 与「FFFFFF = 不填充」的耦合

`plain` 主题的 `colors.surface` 就是 `FFFFFF`。助手若直接拿它当卡片底色，卡片会按既有约定
变成透明 —— 默认值必须在助手层回落（现在用 `F8FAFC`），不能只信主题。

### 13. 有色块的地方都要算文字对比度

白字落在浅色卡片上会直接消失。序号圆、KPI 数字、横幅的文字色现在按底色亮度自动选黑/白。

### 14. 逐条告警会淹没页级告警

40 张卡会产出 40 条「正文放不下」，把唯一可操作的那条埋掉。溢出/过密时只报页级一条。

## 校验方法

### 真实 Office（本机）

PowerShell 的 `New-Object -ComObject` 在本机取不到成员（类型库加载失败），
用 `cscript` + VBScript 走纯 IDispatch：

```powershell
cscript //nologo test\validate-com.vbs "<含 Office 文件的目录>" "test\.tmp\com-report.txt"
```

两个坑：
- `Workbooks.Open(path, 0, True)` 这种「位置参数传可选参数」的写法会让 Excel
  回一句「不能取得类 Workbooks 的 Open 属性」，看起来像文件坏了，其实不是。
  用单个参数 `Workbooks.Open(path)` 就好（已用对照实验确认：打开不存在的路径
  会正常报 1004 文件错误，`Workbooks.Add` 正常）。
- `Presentations.Open(..., WithWindow:=False)` 时，凡带图片的文件都会
  「could not open the file」（Office 自带模板、python-pptx 官方输出同样失败），
  必须 `WithWindow:=True`。

### 第二个引擎（部署自带的 LibreOfficeKit）

```powershell
node test/validate-libreoffice.mjs --dir test\.tmp\e2e
```

它启动的转换工作进程需要管道 stdio，在受限沙箱下会 `spawn EPERM`；
这一项要在放宽的文件权限下跑，否则会误报失败。
