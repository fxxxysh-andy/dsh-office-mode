/**
 * 按需文档（office_help 的内容来源）。
 *
 * 工具声明只有两行，写法和细节都放在这里，模型需要哪一种格式再取哪一种。
 * 格式部分直接从模块的 `meta` 生成，所以新增一种格式不需要同步改文档。
 *
 * 文档分两层（第十七轮）—— 这一层是**注入预算**的落点：
 *
 *   - 默认（compact）：只给「怎么写才不错」的最小事实集 —— 操作名与签名的
 *     一览、致命的边界（没有文本层的 PDF、office.memory 不在
 *     脚本里…）。一次投影进上下文之后，这份文档会留在**每一个后续请求的前缀**
 *     里（对话前缀只增不减），所以它越薄越好。
 *   - `detail: true`：例子、为什么这样做、配置项与迁移说明等长尾。需要时再取，
 *     不必为了问一句 API 形状就把整篇散文塞进前缀。
 *
 * 分层的判据只有一条：**模型照着默认这一层就能写出对的调用**。为此签名与参数
 * 名一个都不省 —— 省的是解释与举例，不是接口。
 *
 * 真实量级（本机实测，UTF-8 字节；**第七十三轮重测**，上一版这几个数字停在第六十轮、
 * 而写作话题与 guide 已在那之后扩过）：默认层与全文层相比，memory 3674→28248、
 * ppt 7223→25490、search 4001→12944、settings 2028→4878、tex 4068→6357、
 * av 2610→6238、run 1965→5169、convert 1834→5076、word 2292→4697、index 2304→4774、
 * guide 2266→9404；全文层超过 3 KB 的还有 writing（5865）、image（4925）、
 * excel（3228）与 python（3180），其余话题都在 3 KB 以下。（这一段没有套件守门 ——
 * 第六十轮复核就是靠现测抓到末句写错了一处：初稿把 excel / image 漏掉了，改完记得
 * 再量一遍再落笔；第七十三轮同样按现测重写。）
 *
 * @module dsh-office-mode/docs
 */
import { OFFICE_GUIDE } from './guide.js';
import { DEFAULT_PDF_MAX_PAGES, RETURN_ECHO_CHARS } from './config.js';
import { PYTHON_STDOUT_CHARS } from './python.js';
import { renderThemeTable, themeCatalog } from './engine/theme.js';
import { loadAllFormats, loadFormat, formatIds } from './registry.js';

const BASE_TOPICS = {
    run: `office_run 的用法

参数
- script（与 scriptFile 二选一）：一段 JavaScript。可以直接写多步，最后 return 一个值（会回显）。
- scriptFile（与 script 二选一）：读工作目录里的一份 .js 当脚本跑 —— 长脚本先
  office.files.write('build.js', code) 落盘，之后每次调用复用同一段，改动只改那份文件；
  错误行号是那份文件里的行号，它也会记进台账的 source。配 args 按参数复用。
- args（选填）：传给脚本的参数，脚本里以全局 args 读到。
- purpose（选填）：一句话说明这次要做什么，只用于显示。
- keepCache（选填）：false 时这次调用结束就清空缓存目录；不传则按设置里的
  「保留中间产物」（默认保留）。中间产物默认跨调用留着，才会命中缓存。

脚本里能用什么
- office.word / office.excel / office.ppt —— 造文档、读文档、改文档
- office.tex —— LaTeX 学位论文：生成 thuthesis 源文件项目、编辑、编译出 PDF
  （见 office_help({topic:'tex'})）
- office.pdf —— 读 PDF：info / text / pages（见 office_help({topic:'pdf'})）
- office.python —— 科学计算与绘图：run / file / check（见 office_help({topic:'python'})）
- office.av —— 音频与视频的内容提取：check / info / transcribe / frames / extract
  （见 office_help({topic:'av'})）
- office.preview / office.image / office.writing / office.archive —— 渲染 / 配图 / 写作规划 /
  读压缩包（见各自话题）
- office.convert —— 互转与统一出 PDF（见 convert）
- office.files —— 读写改文件；批量替换只写一次盘
- office.cache —— 中间产物目录（可命中复用）
- office.theme —— 主题清单
- office.log / office.warn / office.note —— 追加到反馈里
- office.assert(条件, 说明) —— 不满足就中断并报错

没有的东西：process、require、fs、网络、计时器。
记忆也不在脚本里：没有 office.memory 这个入口。脚本跑完、文件复检之后，
office_run 会把写出的文档自动登记进台账；要主动记偏好或查记忆，用
office_memory 工具（见 office_help({topic:'memory'})）。
路径一律相对当前工作目录；也可以用绝对路径。
能调外部程序的地方只有六处，都由插件自己探测、脚本里给不了命令：
PDF 渲染（poppler / PyMuPDF / Ghostscript 等）、LaTeX 编译（latexmk / xelatex / lualatex）、
Python（python / python3 / py）、ffmpeg（office.av）、预览与转换（LibreOffice 随部署装）、
pandoc（office.convert）。

反馈里会有什么
- 每个写出的文件都会被自动重新打开检查一遍，附上真实统计和结构摘要。
- warnings 是排版风险提示，不是错误；按提示调整通常能明显变好看。
- 脚本返回的 JSON 超过 ${RETURN_ECHO_CHARS} 字符会被截断；长内容写进文件再读，不要靠返回值带。

三个真实用例（第十八轮 P1-2 / 18-12：这三个形状是从真实会话里重建出来的，照它们
改参数就够；每行参数的完整含义见对应话题的 detail:true。**默认层不带例子**）

【例 1】Word 报告（.docx）
    const wb = office.word.create({ path: '季度汇报.docx', title: '2026 年 Q1 汇报' });
    wb.heading('一、整体情况', 1);
    wb.para('本季度营收 1,280 万元，同比增长 12%。');
    wb.bullets(['华东区贡献 46%', '新签客户 18 家']);
    wb.table({
        columns: ['季度', { title: '营收', width: 5, align: 'right' }],
        rows: [['Q1', 1280], ['Q2', 1530]],
        caption: '分季度营收',
    });
    const report = wb.save();          // 返回值就是复检结果：stats / outline / warnings

【例 2】Excel 表（.xlsx）
    const wb = office.excel.create({ path: '明细.xlsx', title: '2026 年 Q1 明细' });
    const sheet = wb.sheet('明细');
    sheet.title('2026 年 Q1 明细', { span: 4 });
    sheet.table({
        columns: [{ title: '季度', type: 'text' }, { title: '营收', type: 'number' }],
        rows: [['Q1', 1280], ['Q2', 1530]],
        totalRow: true, freeze: true,
    });
    sheet.stats('B', { funcs: ['sum', 'average'], format: '#,##0.00' });  // 只给列 = 取表格数据区
    wb.save();

【例 3】Python 画图，再把图嵌进文档（office.python + office.word / office.ppt）
    const code = [
        'import matplotlib.pyplot as plt',
        'use_cjk_font()',
        'fig, ax = plt.subplots(figsize=(6, 3.6), dpi=160)',
        "ax.plot(['Q1', 'Q2'], [1280, 1530], marker='o')",
        "ax.set_title('季度营收'); ax.set_ylabel('万元')",
        "fig.savefig(OUT_DIR + '/trend.png', bbox_inches='tight')",
    ].join('\\n');
    const py = await office.python.run(code);            // 图必须落在 OUT_DIR 才会出现在 files 里
    const png = py.files.find((file) => file.kind === 'image');
    wb.image(png.path, { widthCm: 12, caption: '季度趋势' });   // PPT 用 deck.image({ path: png.path })
    // 长脚本先 office.files.write('analysis.py', code)，再 office.python.file('analysis.py', { args: [...] })`,

    pdf: `PDF 的读法（office.pdf）

PDF 分两类，读法完全不同，先问一句再动手：

- office.pdf.info(path) —— 这份 PDF 里有什么
    返回 pages / pageSizeCm / title / producer / encrypted /
    hasTextLayer / counts{fonts,images} / fonts（字体名，探测得到时） / hint。
    hasTextLayer 为 false 就是扫描件或手写笔记（里面没有字，只有图），
    别去抽文字，直接走 pages。

- office.pdf.text(path, { from?, to?, out? }) —— 抽文字（文本型 PDF）
    out 给出时把文本写进那个文件，返回 { path, chars }；不给就把文本
    放在返回值里（长文本会被工具反馈截断，超过一屏就一定要给 out）。
    chars 为 0 说明这份 PDF 没有文本层，改走下面一条。

- office.pdf.pages(path, { from?, to?, dpi?, format? }) —— 渲染成图片
    返回 { engine, dir, files: [{ page, path, bytes, reused }], reused, rendered }。
    拿到 path 之后**用 read_image({ file_path }) 逐页看** —— 手写笔记、
    扫描件、图表都靠这一步进上下文。这是读图能力，不是 OCR，认得出字。
    dpi 默认 120（手写建议 120-200，越大越慢）；format 默认 png，也可 jpeg。
    单次最多渲染 ${DEFAULT_PDF_MAX_PAGES} 页，超了要分批传 from/to。
    结果按「文件内容 + dpi + 格式 + 引擎」缓存在 .office/runtime/cache 里：同样的页码再渲染
    就是 reused（缓存命中），不重算。

- office.pdf.engines() —— 本机探测到的引擎（渲染 / 抽文本 / 页数）
    渲染用其中一个：fitz（PyMuPDF）/ pdftoppm / pdftocairo / mutool / gswin64c。
    一个都没有时会明确报错并告诉你装什么，不会假装支持。

典型流程（把 PDF 变成 Markdown / 讲义 / PPT）
1. const info = await office.pdf.info('讲义.pdf');   // 先看是哪种
2. 有文本层：await office.pdf.text('讲义.pdf', { out: 'raw.md' })，再 read 那个文件。
   没有文本层：const r = await office.pdf.pages('讲义.pdf', { dpi: 150 });
   然后对 r.files 逐页 read_image。
3. 照着内容写目标文件（.md / .docx / .pptx）。
4. 交付前把「原 PDF 里的内容」与「你写出来的东西」对一遍，缺页或漏字要说明。`,

    python: `Python 计算与绘图（office.python）

办公模式里没有命令行，但留了一小块 Python：只做科学计算与绘图，跑完把
stdout / stderr 与产出的文件一起带回来。装包、改环境、开 shell 都不在能力内。

先问一句：await office.python.check()
    返回 available / interpreter / version / packages（装了哪些包与版本）
    / missingPackages / outDir。有 numpy、scipy、pandas、matplotlib 就用它们；
    没装的包不要 import —— 会当场 ModuleNotFoundError，白跑一次。

跑一段代码：await office.python.run(code, { name?, outDir?, timeoutMs?, args? })
    返回 { ok, code, stdout, stderr, files, script, outDir, logFiles, elapsedMs }。
    - ok = 退出码为 0。脚本自己报错时 ok:false，stderr 里是完整回溯，
      行号就是代码里的行号，直接照着那一行改。
    - files 是这次跑出来的产物，图片标 kind:'image'，路径可以直接交给
      Word 的 builder.image(path, { widthCm, caption }) 或 builder.images(items),
      也可以用 PPT 的 deck.image({ path }) / deck.images(items)。
    - stdout 超过 ${PYTHON_STDOUT_CHARS} 字符会截断，完整输出在 logFiles.stdout 里。
    - 产物默认落在缓存目录的 python/out（即 .office/runtime/cache/python/out）。
    - args 会变成脚本的 sys.argv[1:]（参数化跑同一段分析用）。

脚本里多出来的三个名字
    OUT_DIR   产物目录的绝对路径。图要往这里写，才会出现在 files 里
    WORK_DIR  当前工作目录的绝对路径（读输入数据用）
    use_cjk_font()  让 matplotlib 的中文标题与负号正常显示；不调用就不动任何配置

画图（一定要 savefig 到 OUT_DIR，不要 plt.show —— 没有显示器，后端是 Agg）
    import matplotlib.pyplot as plt
    use_cjk_font()
    fig, ax = plt.subplots(figsize=(6, 3.6), dpi=160)
    ax.plot([1, 2, 3], [4, 5, 9], marker='o')
    ax.set_title('月度趋势'); ax.set_ylabel('万元')
    fig.savefig(OUT_DIR + '/trend.png', bbox_inches='tight')

长脚本先写文件再跑
    office.files.write('analysis.py', code)            // 或 office.cache.write
    await office.python.file('analysis.py', { args: ['2026'] })

几条边界（照实说，不要试）
- **要图表对象就别画 PNG**：Excel 的 sheet.chart({type, title, categories, series})
  与 PPT 的 deck.chart({...}) 生成原生图表部件（可改类型、可编辑数据）；
  Python 画 PNG 适合示意图与 Office 画不出来的形状。Excel 要放真实位图时用
  sheet.image(path, {at, widthCm, heightCm, alt})（PNG / JPEG 原生图片部件）；
  Word / PPT 也可以直接嵌 Python 画出来的 PNG。
- 数据要交付就写成 CSV 放 OUT_DIR；要当交付物的文本仍走 office.files.write ——
  Python 直接写工作目录的文件不进台账，也不会被复检。
- 没有网络（pip install 装不了包）、没有 stdin（input() 会立刻 EOFError）、
  没有命令行（脚本里给不了别的命令）。
- 跑一次的成本是一次进程启动（几百毫秒）加你自己的计算，别把 office_run
  拆成很多次小 Python 调用。`,

    files: `office.files 的用法

- office.files.read(path)              读文本（二进制会当场报错并指路，不给乱码）
- office.files.readBytes(path)         读原始字节（二进制安全，返回的字节视图可直接交给 write / copy）
- office.files.write(path, 文本或字节)  整份覆盖写（**顶层**字节按字节写；嵌在对象里的字节、Map / Set、函数与 undefined 一律被拒绝，不静默写成坏文件）
- office.files.copy(from, to, {overwrite?})   整份复制（二进制安全，走内核复制、不经过脚本内存）
- office.files.move(from, to, {overwrite?})   移动 / 改名（同一份文件换个位置或名字）
- office.files.edit(path, edits)       批量字面替换，只写一次盘
      写法：[[旧, 新], [旧, 新]]  或  [{ find, replace, all }]
      没命中的 find 会出现在结果 missing 里，不会静默跳过
- office.files.template(path, {k: v})  把 {{k}} 一次性替换掉
- office.files.exists / list / remove / stat

改 Markdown 或长文本就用 edit/template：一次调用改完所有位置，
不要一条一条改。
二进制（PNG / PDF / 压缩包 / Office 文档）不能用 read 读：**前 8 KB 里有空字节**、
或**开头命中已知二进制魔数**（%PDF- / PNG / JPEG / GIF / ELF）就会直接报错
并指路 —— 要字节用 readBytes，要整份搬用 copy / move，要看内容用对应的 office.* 读取器
（PDF → office.pdf，docx/xlsx/pptx → 对应格式的 read）。
压缩包（zip / gz / zst，含会话导出）不能用 read 读：用 office.archive，
见 office_help({topic:'archive'})。read 遇到压缩包会直接报错并指到那里，
不会给你一段乱码。
copy / move 的目标已存在时要显式传 overwrite:true（与 preview.pdf / image.fetch
同一条口径）；复制会把源文件登记成这次读过的输入，移动不会（移动后源路径已不存在）。
两个动作只处理**单个文件**：源或目标是目录会直接报错（不递归、不跟随链接）；
源与目标指向同一个文件（含硬链接 / 符号链接别名）也会拒绝。
路径按工作目录解析，相对路径里的 ../ 可以出目录（与 write / remove 一致），绝对路径也允许。
office.archive 的判型按**魔数**（ZIP / gzip / zstd / 纯文本），不看扩展名；
zstd 要 Node 22.15 以上；返回全部有界（text 可翻页、find 有命中上限）。`,

    cache: `office.cache 的用法

- office.cache.dir()            缓存目录（相对工作目录）
- office.cache.write(n, 数据)    写中间文件
- office.cache.list()           当前内容（递归，带 modifiedMs）
- office.cache.stats()          体积与命中统计（跨调用复用了多少次）
- office.cache.clear()          立刻清空

缓存**跨调用保留**（默认），所以里面的东西可以下次接着用：
- 渲染好的 PDF 页面图放在这里，下一步 read_image 直接读；
- 同一份输入重复渲染会命中已有文件，不重算。

清理规则：每次 office_run 开始时只收走超过 TTL（默认 12 小时）没碰过的条目；
office_run({ keepCache: false }) 或 office.cache.clear() 才会整目录清空。
缓存目录是工作目录下的隐藏目录，交付物不要放这里 —— 交付物写工作目录。`,

    archive: `归档与会话日志的读法（office.archive）

会话导出（session*.zip）、单独压缩的 jsonl（.gz / .zst）**不能用 office.files.read 读** ——
那是压缩包，按文本读只会得到乱码；office.files.read 认出来会直接报错并指到这里。

- office.archive.info(path)                    这份包里有什么（kind / 条目数 / 解压后体积）
- office.archive.list(path, {limit})           逐条列条目（名字、字节数、压缩方式）
- office.archive.text(path, {entry, offset, maxChars})   取一段正文（默认 8000 字符，可翻页）
- office.archive.find(path, {pattern, entry, maxHits, context})   按行正则检索（会话日志最常用）
- office.archive.extract(path, {entry, to})    把一条正文落到工作目录，再 office.files.read

判型按**魔数**（ZIP / gzip / zstd / 纯文本），不看扩展名；zstd 要 Node 22.15 以上，
本机没有那个函数会明确报「本机不支持」。所有返回都有界；日志动辄几 MB，别整份往返回值里带。
解压**在分配之前**就按声明值封顶（ZIP 看中央目录、gzip / zstd 用 maxOutputLength）——
一个几百 KB 的压缩炸弹不会把内存吃掉，只会得到一句「超过单条上限」。

典型流程（读一次会话导出）：
    const info = office.archive.info('session10.zip');      // kind: zip，6 条
    const hits = office.archive.find('session10.zip', { entry: 'session.v4.jsonl', pattern: 'office_run' });
    const head = office.archive.text('session10.zip', { entry: 'session.v4.jsonl', maxChars: 4000 });`,

    preview: `渲染预览（office.preview）—— 把自家写出的文档渲染成页面图 / PDF

office.word.read / excel.read / ppt.read 报的是**结构**（统计、大纲、排版风险），
回答不了「看起来对不对」。要看画面就渲染成图，再交给 read_image。

- await office.preview.check()
    → { available, engine, backend, version, formats, defaultDpi, maxPages, hint }
    引擎是**宿主随部署安装的那份 LibreOffice**（@deepseek-ai/libreoffice-kit），
    不是要另装的插件。available:false 时 hint 里写了缺什么，其余能力不受影响。
- await office.preview.render(path, { pages?, sheet?, range?, dpi?, maxPages? })
    → { ok, kind:'images', dir, dpi, pageCount, count, images:[{index,page,width,height,bytes,path}],
        missingFonts, reused, hint }
    - path 支持 .docx / .xlsx / .pptx（另有 .doc/.xls/.ppt/.odt/.ods/.odp/.pdf）。
    - pages：'all'（默认）或页码数组（从 1 开始），如 [1,3]；xlsx 用 sheet（工作表名）
      配 range（A1 区域），**pages 与 sheet 不能同时给**。
    - dpi 默认 110（范围 24-600）；maxPages 默认 12、上限 30 —— 文件比 maxPages 多时
      只渲前 maxPages 页并在 hint 里说明（不会因为页数多就失败）；要看别的页传 pages。
    - 拿到 images[].path 之后**逐张 read_image**。图落在 .office/runtime/cache/preview/ 下，
      是中间产物；同一份文件（内容未变）+ 同样的参数会命中已有图，reused:true 不重渲染
      （命中的返回形状与新渲染一致：pageCount / backend / missingFonts 都在）。
    - missingFonts 是「文档声明了但本机没有的字体」：渲染出来会有替换字体，不是文件坏了。
- await office.preview.pdf(path, { to?, overwrite? })
    → { ok, kind:'pdf', path, bytes, reused }
    不传 to 只放缓存目录；传 to 写到工作目录（交付物，会出现在反馈的 otherFiles 里）。
    to 不能是输入文件自己（会把原文档覆盖成 PDF）；目标已存在且内容不同时要显式传
    overwrite:true；相对路径必须在工作目录内。

何时用：**按需**。渲染一次 2-5 秒（本机实测），所以只在「要确认版式 / 交付前看一眼」
时调，别每写一份都调。典型：写完 pptx → office.preview.render('汇报.pptx') →
逐张 read_image → 发现某页文字溢出 → office.ppt.revise 就地改 → 再 render 一次确认。`,

    convert: `文档互转与统一出 PDF（office.convert）—— tex / md / word 互相转，认识的文档格式都能出 PDF

三件引擎各管一段，都在运行期探测（office.convert.check() 报实况）：
- pandoc：md / docx / tex 之间的互相转换（也认 odt / html / rst / txt 源）；
- 本机 TeX 发行版：.tex → PDF（latexmk / xelatex）；
- 宿主随部署安装的 LibreOffice 引擎：docx / xlsx / pptx → PDF（老格式 .doc / .xls / .ppt
  与 OpenDocument 的 .ods / .odp 同样走它，那几种 pandoc 读不了、只给出 PDF）。

矩阵（源 → 可转的目标）
- md   → docx / tex / pdf
- docx → md / tex / pdf
- tex  → md / docx / pdf
- html / rst / txt / odt → md / docx / tex / pdf
- pdf / xlsx / pptx / doc / odf → pdf（pdf 是原样留一份，其余只有出 PDF 这一条）
pandoc 读不了 PDF 的正文：要 PDF 里的文字用 office.pdf.text，要画面用 office.pdf.pages。

- await office.convert.check()
    → { available, pandoc:{path,version}|null, latex, libreoffice, cjkFont,
        sources, targets, matrix:[{source, targets:[{target, engine, available, reason?}]}], hint }
    只报探到的入口；缺哪样只影响它那几条路（hint 里写了缺什么、怎么装）。
- await office.convert.file(path, { to, out?, overwrite?, engine?, timeoutMs?,
                                   media?, fragment?, reference?, font?, texEngine?, clean? })
    → { ok, from, to, path, bytes, engine, steps, reused, warnings?, media?, hint }
    - to（必填）：md / docx / tex / pdf（也认 markdown / word / latex 与带点的写法）。
    - out：目标路径；不传就写到**源文件同目录同名**换扩展名。
    - engine：只对出 PDF 有意义 —— auto（默认）/ libreoffice（宿主引擎）/ tex（本机 TeX 排版）。
      matrix[].engine 是**路线名**（direct / docx / latex / tex / copy），不是 file() 的 engine 取值。
    - media：docx / odt → md / tex 时把内嵌图片抽成 \`<目标名>.media/\`（默认抽；false 只出文本）。
    - fragment：md → tex 只要正文片段（默认出**能直接编译**的整份，导言区与中文字体都配好）。
    - reference：md → docx 的参考样式文档（一份本机 .docx，Word 里做好样式另存即可）。
    - font：LaTeX 路线用的中文字体名；默认自动探（Windows 微软雅黑 / Linux Noto CJK / macOS 苹方）。
    - texEngine：.tex 出 PDF 时指定编译入口（latexmk / xelatex / lualatex）；
      clean 只管中间产物（'aux' 默认 / 'none'）—— 传 'all' 按 'aux' 处理，它会把 PDF 一起删掉。

出 PDF 的路线（engine 不传时按源格式挑）
- .tex → 本机 TeX 编译（只有 \\documentclass{thuthesis} 的论文项目才拷那套模板）。
- .docx / .xlsx / .pptx 以及老格式 .doc / .xls / .ppt、OpenDocument 的 .odt / .ods / .odp
  → 宿主 LibreOffice 引擎直接渲染（与 office.preview 同一份引擎）。
- .md / .html / .rst / .txt → 默认先 pandoc 成 docx 再交给 LibreOffice（中文字体不用配）；
  LibreOffice 不在时自动退到 pandoc --pdf-engine=xelatex —— 那条退路上，**源里有中文
  （或读不出正文的二进制源）而探不到中文字体时直接报错不做**：pandoc 找不到字体会静默
  丢字（退出码仍是 0，PDF 里只剩空白）。源里没有中文时不需要字体，照常出。

四条边界
- 覆盖：目标已存在且内容不同时要显式 overwrite:true；内容一致就不重写（reused:true）。
  **例外**：\`.tex → pdf\` 的默认落点就是 TeX 自己写 PDF 的位置（源文件旁边同名）——
  那里**已经有一份 PDF 时，不传 overwrite:true 会被当场拒绝**（编译会直接覆盖它，事后再报
  就晚了）；传了就重写并把这次产物登记上，\`reused\` 恒为 false。想两份都留就传 out。
  .tex 编译与 LibreOffice 产物里带生成时间戳，同一份源重转后字节仍会不同，重写要显式说一声。
- 复用：pandoc 的 SOURCE_DATE_EPOCH 钉在**源文件的修改时间**（mtime）上，所以「同一份文件
  （没改过、mtime 也没变）重复转换得到同一份产物」；git checkout / 从别处复制会把 mtime 换掉，
  那时产物字节不同，第二次会被覆盖闸拦下 —— 确认要重转就传 overwrite:true。
- 版式：tex → md 只保证正文与公式，thuthesis 的自定义命令、版式与参考文献样式都会丢；
  要留版式就留 .tex / .pdf，别拿 md 当中转。
- 交给谁：转换只换容器，不改内容判断 —— 要按目标格式的规矩重排版（封面、目录、版式助手）
  仍然用 office.word / excel / ppt / tex 的 create。

引擎缺失：pandoc 从 pandoc.org 装（winget install JohnMacFarlane.Pandoc）；TeX 装
TeX Live / MiKTeX；LibreOffice 由 DSH 随部署安装。便携版用环境变量 DSH_OFFICE_PANDOC
（可执行文件或命令名）与 DSH_OFFICE_CJK_FONT（字体名）指路，装了不用改配置。`,

    image: `配图（office.image）—— 图库检索 / 网页选图 / 抽文档内嵌图 / 元信息 / 署名

office.python 只能画数据图；照片 / 素材要么用户自己放进工作目录，要么从这里取。
两条免 Key 的公开图库通道：维基共享资源（commons，署名信息最全）与 Openverse
（聚合多个 CC 图库）。要 Key 的图库不在默认清单里 —— 取一张图不该先配 Key。
除了图库，还有三条路：看本地图片的元信息（info）、从网页列候选（fromPage）、
把 docx / pptx / xlsx 里已经嵌好的图抽出来复用（extract）。

- office.image.channels()
    → [{ id, label, needsKey, note }]；可用 id：commons / openverse
- await office.image.search(query, { channel?, limit?, width? })
    → { query, channel, count, results:[{index, title, url, fetchUrl, thumbUrl, pageUrl, width, height,
        mime, license, licenseUrl, author, channel, insertable, insertableInWord, insertableInPpt,
        credit, attribution}] }
    - channel 不给就是 auto（按上面的顺序试）；limit 默认 8（上限 40）；
      width 是缩略图宽度（默认 1024）—— 大图先看缩略图，选定了再 fetch 原图。
    - credit 是一行署名（如「维基共享资源 · 作者 X · CC BY-SA 4.0」），直接传给文档的
      source 选项；attribution 是多行完整署名（含来源页与许可链接），写进报告或脚注。
    - insertable 指的是**能不能进 Word**（PNG/JPEG）；GIF 只在 PPT 里能用，所以另给
      insertableInWord / insertableInPpt 两个布尔量，别拿一个值当两处用。
      fetchUrl 是「取哪一份」：原图不是位图（SVG / TIFF）时它指向图库的 PNG 缩略图 ——
      fetch 认的是 fetchUrl，所以矢量图也能拿到可嵌的位图。
    - 接口在 200 里回错误信封（如 error.code = badvalue）时，错误里会转述那个原因，
      不会说成「没有结果」。
- await office.image.fetch(候选或URL, { to?, maxBytes?, overwrite? })
    → { path, bytes, mime, width, height, sha256, source:{url,pageUrl,title,license,licenseUrl,author},
        credit, attribution, insertable, insertableInWord, insertableInPpt }
    - 默认落到 assets/<名称>-<内容哈希前6位>.<ext>；**先校验魔数**，返回 HTML 说明页时
      明确报「不是图片」，不当成下载成功。相对路径的 to 必须在工作目录内（写成 ../ 会被拒绝），
      目标已存在且不同要显式传 overwrite:true。
    - 取回来的依据（来源页 + 许可）会进这次调用的引用列表 → office_run 写进台账 source。
- office.image.info(path)
    → { ok:true, path, bytes, mime, ext, width, height, aspectRatio, sha256,
        insertableInWord, insertableInPpt }；读不到 / 不是位图时回 { ok:false, path, reason }
    - 不联网、不落盘。只按魔数认 PNG / JPEG / GIF / WebP（WebP 三种块头都读宽高）；
      SVG / TIFF / BMP 这些不做元信息。嵌之前先看这一条：Word 只吃 PNG / JPEG。
- await office.image.fromPage(url, { limit?, minWidth? })
    → { ok, url, finalUrl, title, count, results:[候选], skipped:[{url?, reason}], hint }
    - 候选与 search 同构（可直接交给 image.fetch），另带 alt 与 role（og = 社交卡片主图）。
      顺序 og:image / twitter:image → <img>；相对地址按 <base href> 解析，srcset 取最大一份；
      data: / blob: / 非 http(s) 记进 skipped，按解析后的 URL 去重。
    - 只解析服务端返回的 HTML：纯 JS 渲染的页面可能 0 条；网页图的许可通常未知，
      正式发布前自己核对来源（这里只给「图片在哪个页面上」，不给 license）。
- office.image.extract(path, { out?, overwrite? })
    → { ok, format, path, dir, count, files:[{path, part, bytes, mime, width, height, sha256}],
        skipped, hint }
    - 抽 docx / pptx / xlsx 包里的 word/media、ppt/media、xl/media；默认落
      assets/<stem>-media/，文件名带内容哈希前 8 位 —— 同 basename 的不同文档不会互覆。
    - 每张都经统一文件环境落盘（进这次调用的产物清单）；目标已存在默认跳过，
      要重写传 overwrite:true。EMF / WMF / SVG / TIFF 抽得出来，但 Word / PPT 不认，
      要先转成 PNG。

配图必须带来源（正式文档尤其）：拿到 credit 之后嵌进文档时传 source ——
    const picked = (await office.image.search('风力发电机 照片')).results[0];
    const img = await office.image.fetch(picked, { to: 'assets/wind.jpg' });
    wb.image(img.path, { widthCm: 12, caption: '陆上风电场', source: img.credit });
Word 会在图注下面加一行小字「来源：…」；PPT 并进题注那一行。直接给 URL 取图也行，
但那时署名只能自己写（结果里的 license / author 是空的）。`,

    writing: `写作规划与自检（office.writing）—— 生成前分注意力，生成后按同一张表查

这不是提示词，也不是结构模板。它把「哪一节该多写、哪一节该收住」算成一张表：
按每节的内容难度与读者收益分配字数，难点 / 跃迁点多给笔墨，背景与常识压到一两句；
写完再用同一张表查句长、套话、锚点、让步与重点节是否写透。

- office.writing.plan(input, options)
    input：string（一句话 brief）或 { title, genre, audience, goal, targetChars, throughLine, sections }
    sections 每项：string（节名）或 { heading, kind, difficulty, value, purpose, mustCover,
      moves, register, figurative, rhythm, avoid }
    → { version, brief, strategy, sections:[{ id, heading, kind, difficulty, value, weight,
        share, targetChars, depth, depthLabel, registerLabel, figurativeLabel, purpose,
        readerQuestion, mustCover, moves, register, figurative, rhythm, avoid }], checks,
        totalChars, markdown }
    - genre：academic / explainer / report / memo / essay / howto / general（也认论文、科普、
      汇报、提案、评论、指南等中文别名）；不给就是 general。
    - kind：context / problem / method / theory / evidence / analysis / leap / implication /
      example / howto / limitation / summary / faq。给节名时也会按关键词猜（「机制」→theory，
      「局限」→limitation），猜错就显式传 kind。
    - difficulty / value 都是 0–3，越大越该多写。权重 = base × (1 + difficulty×0.16) ×
      (1 + value×0.10)；单节上限按节数（1 节 100% / 2 节 65% / 3 节 50% / 4 节及以上 40%），
      目标字数会精确加总到 targetChars。
    - depth：share ≥ 0.25 是 dwell（写透），≥ 0.15 是 develop（展开），≥ 0.08 是 touch
      （点到），否则 compress（收住）。sections 缺省时按 genre 给一套可改的骨架。
    - checks 是写前确认：读者要同时拿住几个新概念（>2 就拆步）、这一节改变读者哪个判断
      （不改变就删）、全文最长的一节是不是最难的那点；另有目标 / 主线 / 权重 / 节数提醒。
    - markdown 是可直接 office.log 的版本（含篇幅表、全局质感、逐节写法与写前确认）。
- office.writing.audit(text, planOrOptions)
    → { ok, score, band, chars, signals, sections, warnings, suggestions, note }
    - 传 plan 时按节名对表：重点节低于目标 68% 报 underdeveloped，收住节超过 1.7 倍报
      overexpanded；各节实际篇幅仍平均报 uniform-allocation；最长节不是最难/最值钱的节报
      longest-not-hardest。
    - 不传 plan 也能查：句长过匀 / 段长持平 / 套话密度 / 抽象名词过多 / 锚点不足 / 比喻过密 /
      过度对冲 / 通篇无让步 / 每节总结 / 重复句首。
    - score 只是启发式加权（≥82 自然、≥68 尚可、否则模板感重），**不是「像人程度」的证明**；
      改稿看 warnings 与 suggestions，不看分数。
- office.writing.deai(text, options)
    options：{ mode: 'report' | 'safe', register: 'platform' | 'official' | 'academic' |
      'personal' | 'general'（默认，也认中文）, styleRef?: string }
    → { ok, mode, register, chars, flags, stats, styleRef, notes, note,
        ...(mode:'safe' ? { text, edits } : {}) }
    - 只报可解释的句式：A1 对举 / A2 提示语冒号 / M1 元话语起手式；每条带 spans、count、
      density、confidence、reason 与 falsePositiveHint。**没有 0–100 分，也没有「AI 含量%」**，
      不推断作者；人类也会写 A1/A2。
    - register 调口径：platform 与 general 全高置信；official 把 A2 降为提示、A1 中置信；
      academic 把 A1 降为中、A2 降为提示；personal 把 A1 降为提示。F1 破折号 / A4 长定语 /
      A5 顿号 / B1 抽象 / C2 连接词只进 stats，不进默认 flags。
    - styleRef 传一段人类参考稿：比较句长、段长、对举/提示语/锚点/套话/问句密度，
      报「与参考稿的偏离」，不报 AI 概率；参考稿不足 200 字只作方向。
    - mode:'safe' 只做机械删改：删「先说结论：」「说白了：」「一句话：」等元话语前缀；
      「原因不复杂：」改成「原因是」，「这个原因：」改成「这个原因，」；A1 只报不改，
      事实、数据与论证不动。改完必须人复核，不要直接当终稿。
- office.writing.signals(text)：只给表面量（句长 / 段长 / 套话 / 锚点 / 比喻 / 让步 / 问句），
  不判分；调阈值或自己接评测时用。
- office.writing.render(plan)：把规划对象重新渲染成 Markdown。

一条完整用法（Word / Markdown 都适用；脚本里只规划，正文自己写）：
    const plan = office.writing.plan({
        title: '检索噪声为什么会拖垮答案',
        genre: 'academic',
        audience: '做过 RAG、但不熟悉检索噪声的工程师',
        goal: '让读者相信扩库后的正确率下降主要来自噪声，而不是模型变弱',
        targetChars: 1200,
        sections: [
            { heading: '反差的来处', kind: 'context', difficulty: 1, value: 1.5 },
            { heading: '两条机制假设', kind: 'theory', difficulty: 2.8, value: 3 },
            { heading: '可证伪的预测与实验', kind: 'evidence', difficulty: 2.4, value: 3 },
            { heading: '边界', kind: 'limitation', difficulty: 1.8, value: 2 },
        ],
    });
    office.log(plan.markdown);       // 篇幅表与逐节写法进反馈，模型照着写
    // …生成正文…
    const report = office.writing.audit(text, plan);
    if (report.warnings.length > 0) office.log(JSON.stringify(report.warnings, null, 2));

写作口径（与 guide 一致）：
- 学术 / 困难 / 跃迁节多花笔墨：机制链条、替代解释、反例、可证伪预测；说明性内容可以用
  日常比喻、短句与放松语气，但每个比喻要说明「哪里像、哪里不像」，一段最多一个。
- 每 150 字至少一个可核对锚点（数字带口径 / 日期 / 专名 / 原话 / 具体场景）；没有就写「推测」。
- 每个核心主张配一条最强反对或竞争解释；不要用「有人认为」设稻草人。
- 连接词只在关系可能被误读时用；「首先/其次/最后」全文只许出现一组。
- 语言强度跟证据强度对齐：相关性写「与……相关（n=…）」，未验证机制写「一种解释是」，外推
  写「在……条件下」。
- 背景与常识压到 1–2 句；省下的字数给新定义、机制、跃迁点与边界。
- 不编造个人经历、对话或数据；不为了「像人」加错别字、同义词替换或伪口语。

大模型评测：仓库的 scripts/writing-eval.mjs 做 baseline 与 planned 两臂 A/B（A 学术难点 /
B 说明文 / C 汇报），自动匿名化、换序控位置偏差、限长、记录 audit 与反模板清单；没有
WRITING_EVAL_API_KEY / DEEPSEEK_API_KEY 时打印 SKIP，不伪造结果。--mock 可离线跑通流程。
这不是「人类化改写器」：不做同义替换、不改写正文，也不以检测器分数为目标。`,

    settings: `设置界面（Settings > 插件 > 办公模式 / 记忆系统）

插件带两页设置（同一个命名空间 dsh-office-mode），不用改 YAML。

「记忆系统」页（mnemon 关掉之后由它接住那个位置）
- enabled           总开关；关掉后记忆工具、记忆提示段与自动台账一起停
- dir               记忆目录，默认 .office/memory（相对会话工作目录）
- autoLedger        office_run 写出文档时是否自动登记台账，默认开
- promptHint        是否往系统提示里加一段静态说明（记忆在哪、怎么用，不含内容）；
                    办公模式的 persona 是 complete 的，那一段在那里会被丢掉
- userLimitBytes    热记忆里「用户偏好」的容量上限（字节），默认 4096
- projectLimitBytes 热记忆里「项目与环境」的容量上限（字节），默认 10240
- ledgerLimit       台账最多保留多少条，默认 500，超出的最旧记录滚进归档
- archiveKeep       归档最多保留多少卷摘要（默认 60）：一个月一个卷，条数超过 200 的
                    月份会开下一卷（2026-09.md → 2026-09-2.md）。超出时最旧的卷被删
- 迁移命令也写在这一页上（面板只读写配置，没有执行通道）。

「办公模式」页
工具开关 —— 关掉的工具不出现在模型面前，也不占每次请求的 schema 开销。
- office_help / office_run / office_memory
- office_search_run / office_search_dispatch / office_parse_findings
开关即时生效：改完立刻重建工具面，不用重启。

工具面守卫（第五十四轮 新-15①）
- guard.hideDevToolsInOffice  默认**关闭**。preset 里那段「关掉编程类工具」的
  disabled:true 行运行期不生效（那些能力挂在全局层，preset 的禁用行只作用于
  它自己那一层），这一项打开后由插件在**办公会话**里真正掩掉它们：命令行（bash /
  pwsh）、后台任务（job_*）、workflow / ralph、子代理及其控制面、待办、技能、
  目标、插件管理、任务看板。代价写清：模型**自己**起子代理、命令行、后台任务、
  待办与技能入口都没了；检索三步走的派工**不受影响**（它用的是宿主 subagents 服务，
  不是被掩掉的那个工具）。**即时生效**：已经在跑的办公会话也会立刻重算掩码
  （那些入口当场消失，不必重开）。

办公模式插件（第六十三轮 新-21）
- 这一组列的是**办公 preset 自己声明的插件行**（读的是 preset 文件本身，不是当前会话的
  设置）：每行一个开关，改的是那一行的 disabled。宿主侧同时同步 profile 上已有的覆盖行，
  写前自动落一份备份，页面上的「还原上一次备份」把两个文件一起写回去。
- **生效范围是新建的办公会话**：preset 在新建会话时才读，已经在跑的会话不受影响 ——
  用户问「关了怎么没变化」时先看这一条。
- 列表是实时的：外部（git / 手改）动了 preset 文件，点一次「刷新」就会跟着变。

检索编排
- maxParallel     同时铺开几个检索子代理，1-8，默认 4
- resultLimit     解析结果最多列几条，1-40，默认 40
- maxChannels     一次提纲最多几个渠道，1-12，默认 12
- outputDir       提纲与结果文件目录，默认 .office/search
- engine          auto（默认，先派子代理、跑不了用内置）/ subagent / builtin
- provider        用哪条联网通道：auto（按 providerOrder 依次试）/ 固定某一条
- providerOrder   auto 的尝试顺序
- providers.<id>  每条通道各自的 apiKey / apiKeyEnv / baseURL / model / timeoutMs
- preprocess      网页预处理：mode（article | plain | off）/ dropBoilerplate /
                  dedupeLines / minLineChars / keepTitle
- fetch           取正文与 PDF 的上限（与用哪条通道无关）：maxResults / fetchPages /
                  fetchTimeoutMs / maxBytes / maxChars / maxRedirects /
                  pdfMaxBytes / pdfMaxPages

文档与缓存
- defaultTheme / scriptTimeoutMs / maxScriptChars / cacheDir / keepCache
- cacheTtlMinutes   中间产物多少分钟没被碰过就清掉，默认 720（12 小时），0 = 不按时间清
- pdfDpi / pdfMaxPages / pdfEngine   PDF 渲染的默认分辨率、单次页数上限、引擎
- texEngine / texTimeoutMs / texTemplateDir   LaTeX 编译入口（auto/latexmk/xelatex/lualatex）、
  单次编译超时、thuthesis 模板目录（留空自动探测：工作目录里的 thuthesis* → TeX Live 自带）

所有数值参数都带上下界：设置页调到越界会被 schema 挡下，运行期还会再收敛
一次，所以配置写坏了也不会让插件崩。宿主没有设置能力时这两页自动不出现，
插件其余功能不受影响。`,

    memory: `办公记忆（office_memory）—— 三层，跨会话保留

记忆存在会话工作目录下的 .office/memory/ 里。三层不是三种存储，而是三种
**生命周期**：热记忆常驻、台账自动登记、归档只读下沉。

1) 热记忆（layer:'hot'）—— 用户偏好与项目约定
   - 真源是 memory.json；USER.md（用户偏好）与 MEMORY.md（项目与环境）是
     它的**投影**，由插件生成。改记忆只能走 office_memory，直接编辑 Markdown
     会在下次写入时被覆盖。
   - 容量：用户偏好默认 4096 字节、项目与环境默认 10240 字节。装不下时按
     「重要度低、更旧」的顺序把条目下沉到归档，并在 MEMORY.md 里留一条指路
     条目；不会静默丢掉内容。**单条也有上限**（默认 2800 字节，设置页可调，
     0 = 不限）：超过会被拒绝，要求压成一两句、细节放进 source 指的输入文件或
     先 kb-ingest 入库；另有一条建议线（默认 1200 字节），超过只在回执里提醒。
   - 想知道热层忘了哪几条：office_memory({ action:'sunk' }) 按时间列最近几次下沉
     （id + 预览），正文用 read 的 layer:'archive' 取回。下沉的回执本身也会点名
     被下沉的 id —— 只报条数时没人知道丢的是不是要紧的那条。
   - 动笔前照它办。office_help 的返回值末尾本来就会带一份热记忆投影，
     office_run 的反馈里也会带，所以平常不必特意 read。
   - 投影按「本次在做什么」筛：只贴相关条目，其余折叠成一行（带条数与读取入口）。
     还有两道节流：同一条正文在一个会话里只贴一次（话题来回切不重复贴），以及
     条目正文的**投影预算**（设置页 → 记忆 → 投影预算，默认 4096 字节；0 = 不限）。
     超出预算的条目一样不丢：折叠行会报「超出本次投影预算」的条数，要看全文用
     office_memory 的 read（layer:'hot'）。

2) 台账（layer:'ledger'）—— 每份交付物一条
   - office_run 每写出一份 Office 文档就自动登记：路径、格式、主题、字节、
     复检统计、结构摘要、这次的目的、以及**这次脚本经 office.files 读过的输入文件**
     （source）、复检告警（note）。模型不用为此多调一次工具。
   - source 只记同一段脚本里真的读过内容的文件（office.files.read / readBytes →
     readFile / readText）；office.files.copy 也算 —— 复制出来的那一份，来源就是被复制的那一份。
     这次自己写出来的那些要排除：原地 revise 读的是同一份产物，记进去是同义反复；
     office.files.move 同样排除源路径（移动之后它已经不存在了）。
     **不经 office.files 的读取不在其中**：office.pdf 直接读盘、
     office.av 走 ffmpeg、office.python 在工作目录里跑，这三条路拿到的输入不会进
     source —— 要追溯就把依据写进 purpose。脚本本身也走这条路：用 scriptFile 跑时，
     那份 .js 会记成 source（「这份产物是拿哪段脚本构建的」）。
   - 其它情况（依据是口述、检索结果或用户给的临时事实）source 就是空的 ——
     那是如实的「没有文件来源」，不是漏记。
   - 找「上次那份东西」用它：office_memory({ action:'read', layer:'ledger',
     query:'季度汇报' })。多个关键词用空格分开，命中越多排越前。
   - 条数超过 ledgerLimit（默认 500）时，最旧的一批滚成月度归档摘要。

3) 归档（layer:'archive'）—— 只读的长期层
   - 装的是两类东西：热记忆装不下时下沉的旧条目、台账滚动下来的旧记录。
     按月份归卷，一卷两个投影文件：
       .office/memory/archive/YYYY-MM.md     人读的 Markdown（可以 grep、可以直接打开）
       .office/memory/archive/YYYY-MM.jsonl  结构化的那份正文（块 + 条目，一块一行）
     一个月条数超过 200 会开下一卷（YYYY-MM-2.md / .jsonl），标题里写明「第 N 卷」，
     检索与展示仍按月份算。索引 archive/index.json 只存指针（卷、块、条数、条目 id），
     **不内联正文** —— 否则它会比正文文件还大，而只要 id 的调用也得整份解析它。
   - 检索：office_memory({ action:'read', layer:'archive', query:'...' })。
     每行会标出它**为什么**在归档里（热记忆下沉 / 台账滚动 / mnemon 迁移 / Pack 导入）——
     「在归档里」不等于「是被容量挤下来的」，两者要能分开看。
   - 归档本身也有上限（archiveKeep，默认 60 卷）：超出时最旧的卷被删掉 ——
     它已经是压缩过的老内容，这是唯一会真正丢东西的地方。
   - **归档仍然参与召回**（layer:'archive' 与 layer:'all' 都会给归档条目），所以
     「封顶删最旧卷」是一次**真删除**：那一卷的两个投影文件（.md 与 .jsonl）都删掉、
     落在那些条目上的关系边同事务清掉、并留一条墓碑（id + 时刻 + 原因）。
     删除之后它就召不回来了。只删 .md 而把 .jsonl 留着是**假删除**，不要那样改。

读回执每行去掉行首的「- 」之后**第二段**就是条目 id（热记忆紧跟「[重要度]」、
台账紧跟日期、归档紧跟「[月份]」；形如 m-… / L-… / mnemon:… / b-…）：
replace / remove 可以直接给 id 代替 oldText（旧图关系的 link/unlink/related 已下线）。
要引用某条记忆就先把 id 抄下来。

**分区读取（第五十四轮 新-12）**：read 给 tags / entities 就是按它**分区** ——
硬过滤，只出**显式声明过**这些标签 / 实体的条目（正文里出现不算）。这是「一个工作目录
一份记忆、里面混装多主题」时唯一能**只读某一类**的入口；不给分区 = 老行为（全给）。

- 三层各自的归类字段不同：热记忆与归档条目看写条目时声明的 tags / entities；
  台账行看 theme（登记时的主题 id 就是它的归类键）。台账**没有实体** —— 拿
  entities 分区时它一行都不会出现，回执里如实报「分区外的条目」，不装作「台账里没有」。
- 同一个维度里给多个词是**合取**（每一个都要声明过），不是「任一命中」：分区是一把
  钥匙，不是一串检索词 —— tags:['发布','模板'] 问的是同时属于这两个分区的那一批。
- 分区与 query 是两件事，都排在**打分之前**：分区是过滤器（硬隔离），query 也是**过滤器**
  （一个词都没命中的条目不出现，与台账 / 归档同口径）。**热记忆的排序不随 query 变** ——
  仍是「重要度永远是主键 → 新在前」：热记忆装的是**约束**，一条 low 的笔记命中两个词不该
  压过一条 critical 的「永远不要…」；同族两版里也不该让旧版靠命中数冒到当前版前面。
  与那两层的第二处区别是热记忆**不套** strict-v1 的分档丢弃：它条目少、每条都短，
  命中了一个词就不该因为占比低被丢掉。要不带过滤地看全量，去掉 query。
- 分区对 layer:'all' 也生效：三层各自过滤，回执首行印出这次的分区口径与隔掉的条数。
  非数组 / 数组里有非字符串会被**拒绝**；全是空白的分区键（tags:['']）同样被拒绝 ——
  它不等于「没给分区」（静默当没给会让回执连分区行都不印）。要不过滤就传空数组 []。
  （同 since / until 的口径：过滤器不许静默丢弃。第五十七轮顺手删掉了这里重复的一遍。）

4) 知识库（kb，第四十一轮）—— 外部原文的块与字符区间，存在 .office/kb/
   - 它装的**不是结论而是来源原文**：一份文档一行（manifest.jsonl），正文按内容哈希
     切块存 chunks/<docHash>/<n>.json。每块存原文 span（字符区间，source.slice(span)
     与块正文逐字相等）、标题路径、以及块头「标题路径 · 来源 · 日期」——
     块头是 Anthropic Contextual Retrieval 的零依赖版本（官方数字全部来自「给块前置
     50—100 token 专属上下文」）。切块是**结构切块**（Markdown 标题分层 + 段落聚合，
     块大小按来源分档），不做语义切块；默认词法，配了 embedding 才走向量混合。
   - 读法有两条：**检索**与**显式挑读**；浏览走文件夹树。
       office_memory({ action:'kb-tree' })                          整棵树（每层的文档数）
       office_memory({ action:'kb-search', query:'切块 结构', folder:'参考' }) 在子树里检索
       office_memory({ action:'kb-ingest', path:'资料.md', tier:'user', folder:'参考/合同' }) 入库并归到虚拟文件夹
       office_memory({ action:'kb-list', folder:'参考', recursive:false }) 只看当前层
       office_memory({ action:'kb-read', id:'<文档 id 或块 id>' })   读块（逐字引用走这条）
       office_memory({ action:'kb-drop', id 或 path, scope:'global' }) 只删全局根那一份
       office_memory({ action:'kb-update', id, title?, folder?, tier? }) 改文档信息（不重切块）
       office_memory({ action:'kb-edit', id, text })                 改文本源正文并重新入库
       office_memory({ action:'kb-link', sourceId, targetId, kind:'supports', note }) 建块间关系
       office_memory({ action:'kb-related', chunkId })               看相邻块（双向）
       office_memory({ action:'kb-summarize', title, text, chunkIds:['<块 id>', …] }) 多块合成摘要块
       office_memory({ action:'kb-summaries' })                     摘要块清单
   - **虚拟文件夹**（第八十四轮）：folder 只是 manifest 上的一段相对路径（如「参考/合同」），
     不复制磁盘目录；空文件夹不显示。folder 同时是**检索范围**：recursive 默认 true（整棵
     子树），false 只看当前层；根 '' 表示整个库。非法路径（绝对路径 / 盘符 / ..）当场报错。
   - **按次指定根**（第八十轮）：kb-ingest / kb-drop 可带 scope:'workspace' | 'global'，
     只动那一个根（同路径可以两份并存：一份全局、一份工作区）；根没在设置里打开就报错，
     不会悄悄写到别处。省略 scope 时：有全局根就写全局（知识库语义上是这台机器的资料库），
     否则写工作区；显式 scope:'workspace' 仍可写工作区。
     全局根的 kb 在全局记忆目录的兄弟目录（默认 $DSH_HOME/.office/kb）。kb-list / kb-search /
     kb-ingest / kb-drop 的回执会在全局那份上标「（全局）」；未标的是工作区根。
   - **归档并入知识库**：面板不再有独立归档页签；归档条目以只读虚拟文档出现在「归档/YYYY-MM」
     文件夹（id kb:<hash>，固定一块），不复制、不迁移，也删不掉（kb-drop 碰不到）。
   - 入库是显式的（没有后台扫描）：同路径同内容再入库不写任何字节；同路径新内容是**换代**
     （旧块删掉、边同事务清掉、旧块 id 留墓碑）。删除与换代都保持「悬空引用 == 0」这条不变量。
   - 检索默认是**词法**的：中文按字符二元组（bigram）展开、拉丁词与数字整词保留，排序用
     BM25（k1=1.2 / b=0.75），正文与块头两路各自归一化后加权（0.75 / 0.25）。
     显式配 embedding（远端 /embeddings、本机推理服务或 mode:'local'）后升级成 SQLite
     混合检索：词法 + 向量余弦按 hybrid 权重融合。向量先走 ANN（纯 JS HNSW）取候选，
     再按库里的原始向量精确重算；索引缺失 / 损坏 / 超规模时退回精确扫描。selector.mode:'llm'
     时再让一个独立模型从候选里挑编号，主上下文只拿选中的块。**任一段失败都退回上一步**，
     不阻断检索。**auto 无端点时仍是词法**，所以语义改写仍可能召不回 —— 换更贴近原文的词。
   - **零重叠就弃答**：没有一块的词表与查询沾边时，回执明说「弃答」，不会把「最像的那块」
     端出来当命中。混合检索的向量单独命中要求余弦 ≥ 0.25，本地向量必须显式开启。
     分数是**同一批命中里的相对次序**（每路都归一化过），不是置信度。
   - **索引**：manifest 是唯一真源；kb.sqlite 只是可重建的索引（docs / chunks / FTS /
     embeddings / chunk_links / summaries）。拿不到 node:sqlite 时自动回退 manifest + 词法，
     入库 / 读取 / 删除不受影响。全局库目录由设置 memory.kbDir 指定，默认 $DSH_HOME/.office/kb。
   - **多格式入库**：文本直读；PDF / DOCX / PPTX / XLSX 抽文本；图片落 assets/ 并把
     ![名](assets/…) 写进正文，图片块可检索。单文件上限 8 MB；PDF 只抽文本层，不抽内嵌图。
   - 命中给块 id / 路径 / 块号 / 字符区间 / 块头 / 预览（每块开头若干字符）。要逐字引用
     就把块 id 给 kb-read 读整块 —— 预览只是让你决定要不要读。检索有界：单次最多扫
     2000 块、最多给 8 条命中（limit 可收小，硬上限 20），被截断时回执会说明。
   - 次数占**每回合配额**里的「知识块检索」一档（默认 2，0 = 不限制）—— 词法检索最容易
     被「没查到就换个词再查」拖着走。tier / path 可以当筛选（只看某一档或某一份）。
   - **块间关系与摘要**（第八十五轮）：kb-link 写 chunk_links，两端都必须在索引里；
     kb-related 走 1–2 跳邻居（双向）；kb-summarize 把多块合成一条独立摘要块
     （summary_of 记录来源块），摘要块照常进检索，来源块仍保留。关系不会自动连边，
     摘要默认仍是手写；开了下面的 AI 增强才由模型自动生成。
    - **AI 建库**（可选，默认关）：设置 memory.llm 里选 mode（host = 宿主已配置模型，
      openai = 任意 OpenAI 兼容 /chat/completions 端点）与 provider/model（Key 可走环境变量）。
      kb-ingest({ ai:true, aiModel? }) 才走模型：chunking 让模型只给「哪些小节并成一块」的
      分组，正文仍按原文偏移 source.slice 切出，不改写、不删字；summary 生成摘要另存
      summaries 表，不替代原文块，检索命中摘要会带来源块 id。aiEstimate:true 只预估
      批数 / 调用数与预算，不调用模型；超预算直接拒绝；单次有 maxCallsPerRun，到顶返回
      paused + stateKey，aiResume:true 续跑。任一次调用失败 / 超时都按确定性切块回退
      （degraded 计数），不会让入库失败在半路。
    - **千万字长篇**：单文件上限 2 MB（约 60 万汉字），先按卷 / 章切成多个文件入库；
      全局词法检索单次最多扫 2000 块，几十卷的书会截断后半（bounded=true 会说明）。
      精准召回靠「先定位再检索」：folder 收窄到卷 / 章、path 锁定某一份，再在范围内
      查关键词 / 人物 / 事件；AI 摘要可给每卷一段导语帮助选卷。不保证 100% 召回。
   - **来源三档是硬规则，不是权重**：user（用户放进工作目录的文档）/ verified（≥2 个
     独立来源确认过）/ unverified（单源未核实）。unverified 可以进 kb 的待核区，但
     **永远不能进热记忆** —— add 时声明 tier:'unverified'、或引用一块 unverified 的
     知识块，都会被当场拒绝。依据：投毒率 < 0.1% 就能把攻击成功率推到 80% 以上，
     而「按来源类别排除」有效、「加性可信度权重」与完全无防御不可区分（p = 0.80）；
     读时过滤还要付 4.4 个准确率点、误隔离 33.6% 的合法记忆 —— 所以设防在写入期。
     第二家独立来源确认之后用 kb-ingest({ tier:'verified' }) **升档**（正文一字不改）
     —— 硬规则不留死路；不带 tier 的重复入库不改档。
   - 引用（三根线）：把块 id 写进 add / log 的 source（形如 kb:<内容哈希>:<块号>），
     它们会被**另存**成引用字段 kbRefs（记忆的 kbRefs、台账的 kbRefs）—— 于是
     「这条结论依据哪块来源」「哪份交付物用过这块来源」都能反着查；link 也可以直接
     指向块（kind:'supports' / 'contradicts'）。引用块时 id 必须在 kb 里找得到，
     否则报错（编不出来的引用等于一句无法核对的话）。
   - 冲突边的三态：link({ kind:'contradicts', conflict, state })。类别取三分
     （context-memory / inter-context / intra-memory），状态取三分
     （unresolved / prefer-source / prefer-target）。**保留双方、显式承认未决**，
     不做自动裁决（模型对冲突的行为自相矛盾，且系统性偏好更早的证据）；同一条冲突边
     再调一次可以改状态，related 的回执里会带着「冲突：未决 / 偏向哪一端」。

5) 统一入口路由（第五十七轮 19-8 的 P2-8）—— 一条 recall 决定要不要下探 kb
   office_memory({ action:'recall', query:'…' })：先查记忆三层，再按**现算的门控**决定
   要不要下探知识库，并在回执首行自报这次的路由。门控是纯判据（memory.js 的
   gateVerdictOf），四档按顺序判：
     empty   命中 0 条 —— 记忆里没有这件事；
     partial 查询词没被命中条目的词表全覆盖 —— 有一半问题记忆答不了；
     weak    词全覆盖、但没有一条约束级 —— 有相关内容，却没有「必须 / 永远」可依；
     enough  词全覆盖 + 至少一条 critical —— 停在记忆层，不再花 kb 配额。
   除 enough 之外都下探 kb；下探被「知识块检索」配额挡住时，回执如实写「未下探」并给出
   那条真的能走的路（kb-read / kb-list / 调设置）—— **不把「被挡住」说成「不必查」**。
   三项原料（命中数、查询词覆盖、最高重要度）都在回执里印出来，判据可以逐次核对。
   两条要记住的边界：① **覆盖是三层命中的并集**，不是「有一条命中覆盖了整句查询」——
   所以 enough 只表示「词都对上了、而且有约束级可依」，不保证一条就答得全；
   ② **约束级看的是热记忆层与归档里的 hot 条目**（台账行没有重要度这个字段），
   所以热层被挤空、critical 沉到归档之后，门控照样认得出它。
   配额上与 read 带 query **共用同一条界**（初查 / 细化各一次）：换一个入口不该换来额外的
   检索次数；只有真的下探了才再吃一档 kb 配额。
   什么时候仍该直接用 read / kb-search：只想要记忆层（或只要台账 / 归档、要看现状）时用
   read；已经知道要查来源原文、也不怕多花一次配额时用 kb-search。recall 是「不知道该查
   哪边」时的默认入口，不是它们的替代品。

写入（action: add / replace / remove）
- add     新记一条：content 要一两句、自包含，下次不看上下文也读得懂。
- replace 改一条：oldText 给一段能**唯一**命中原文的片段，或直接给 read 行首的 id；
          命中多条会被拒绝（宁可让你补长，也不要改错）。
- remove  删一条：同上要求唯一命中。只在用户明确要求、或确有证据说明作废时删。
- target  user = 用户是谁、偏好、要求、忌讳；project = 项目约定、环境事实、
          工具坑、可复用的做法。选错层会让「项目的事实」被当成「用户的偏好」。
- scope   选填 workspace / global：这一条只写进指定根，不按设置走。省略时 user 在
          userScope=global 下落全局、其余按 scope；显式指定时替换 / 删除也**只在那一个根里找**，
          不跨根兜底（找不到就报错，不静默跑到另一个根）。
- importance  critical（明确的必须 / 永远不要）/ normal / low（只对一段时间有用）。
- entities / tags  选填。entities 填了才能按它分区读取（实体是声明的，不猜；
  第九十一轮起不再有实体视图）；tags 参与归档检索。

只记什么 / 不记什么
- 记：用户说过的偏好、纠正、要求「记住」的事；稳定的项目约定与环境事实；
  踩过一次、下次不想再踩的坑。
- 不记：检索结果原文（那在结果文件里）、一次性进度、自己的推测与复述、
  从记忆里刚读出来的东西（读不是写）。
- 记完给一句说明就行，不要把条目整段念给用户听；记忆是静默生效的。

手工登记（action:'log'）
自动登记只认 office_run 写出、且注册表认得出来的格式（docx / xlsx / pptx / tex）。
其余文件 —— **PDF**（office.pdf 只读，tex 编译出的 PDF 也不经写盘记账）、.md 笔记、
.js 脚本、抽出来的图、用户放进工作目录的外部文件 —— **不会**进台账，
需要它们出现在「上次那份东西」里就得手工补：
office_memory({ action: 'log', path:'参考稿.pptx', format:'ppt',
purpose:'作为版式参考', source:['9-15文以载道.md'] })。

记忆范围与层开关（设置页可改）
- 范围 scope：workspace = 只跟当前工作目录（一个项目一份，默认）；global = 所有
  项目共用一份；both = 两者并存 —— 读的时候**全局在前、工作区在后**，因为具体
  项目的约定比通用偏好更该被最后看到。来自全局根的行标「（全局）」，未标的是工作区根。
- 全局层默认在 $DSH_HOME/.office/memory（globalDir 可改；相对路径相对用户主目录，
  不相对工作目录 —— 全局层跟着工作目录走就没有意义了）；它的 kb 在兄弟目录
  $DSH_HOME/.office/kb。
- userScope：memory = 「用户偏好」跟着上面的范围；global = 始终落全局层，
  这样换一个工作目录不必把「用户是谁」重记一遍。
- 按次指定根（第八十轮）：add / replace / remove / kb-ingest / kb-drop 都可带
  scope:'workspace' | 'global'，只动那一个根；根没开就报错。它不改变读的可见范围：
  读仍按 scope 合并各根。典型用法是把跨项目通用的偏好写进全局、项目流程留在工作区：
  office_memory({ action:'add', target:'user', content:'…', scope:'global' })。
- layers.hot / ledger / archive：每层独立开关。关掉只是不再读 / 不再写，
  **已有数据不会被删除**，重新打开就回来。

图关系已下线（第九十一轮）
- 旧的 link / unlink / related / entities 动作、面板图谱页与实体视图都已移除；这些数据不再参与浏览或召回。
- 历史 links.jsonl 与旧 Pack 里的关系数据**不删除**：export / import 继续兼容，删除条目时内部仍会清掉落在它上面的旧边并留墓碑，避免悬空引用污染后来的导入。
- 条目自带的 entities 字段仍可作为分区读取的硬过滤条件，但不再有「实体视图」。
- 新需求请用知识库块关系（kb-link / kb-related），它有自己的限流与面板。

召回质量（strict-v1）
- 带 query 的检索会分档：按「命中了查询里几成的**词**」算相关度（**不是**向量余弦
  相似度），达到高阈值的直接采纳，中档与未知档各有一个名额上限，低于下限的丢掉。
- **中文按字切二元组**（「记忆面板」→ 记忆 / 忆面 / 面板），所以词序颠倒（「面板记忆」）
  与合称（查询「音视频」、正文写「音频与视频」）都能命中；英文与数字仍按整词。
  这条不这么做的话中文会退化成「整串命中才有结果」的二值开关。
- 先取「条数上限 × candidateMultiplier」个候选再分档：直接截断的话，低分结果会把
  高分结果挤出候选池。被丢掉多少条会在结果里报出来。
- 想让它别丢结果：把设置页的 policy 改成 off。

读取的字符额度（layer:all 与单层的区别）
- 单层读取（layer:'hot' / 'ledger' / 'archive'）独占 4,000 字额度，**要看全部就用单层**。
- layer:'all'（默认）三层**各有各的额度**（热记忆 4,000 / 台账 2,000 / 归档 1,600 字），
  谁也不吃谁的。以前是三层抢一份 4,000 字，结果热记忆正文一满就把台账与归档挤成 0 条。
- 任何一层被截断时，结果里会明说「给了几条 / 额度多少字 / 看全部用哪个 layer」。
  看到「（命中 N 条，但本层 X 字额度被更长的条目占满，一条都没给）」就是额度问题，
  不是没有数据 —— 换成单层读取再来一次。

时间窗（read 的 since / until）
- 按时间缩小搜索空间：{ action:'read', layer:'ledger', since:'上周' }、
  { action:'read', layer:'all', query:'汇报', since:'2026-09-01', until:'2026-09-27' }。
- 写法：2026-09-27、2026/09/27、2026-09-27T10:30（按**本地**时区解析）、
  带 Z / 偏移量的 ISO 时刻（按绝对时刻）；时间词：今天 / 昨天 / 前天 / 本周 / 上周 /
  本月 / 上月 / 今年 / 去年 / 最近7天 / 最近3个月 / 最近48小时（也认 this week、last month）。
- 只给日期时是**闭区间**：since 取当天 00:00:00，until 取当天 23:59:59.999。
- 过滤字段：热记忆用 updatedAt（改动时刻），台账与归档用 at —— 归档里 mnemon 迁来的
  历史条目只有 storedAt，会自动回退到它，所以时间窗读得到它们。
- 解析不出来**直接报错**，不会静默忽略 —— 静默忽略会把「那段时间没有」说成结论，
  而它只是过滤器没生效。时分秒越界（T25:00）也算解析不出来，不靠进位蒙过去。
  结果里会印出解析成了哪两个时刻（本地时刻）。
- 带时间窗时，**没有时间戳的条目一律排除**，而且回执会把它与「在时间窗之外」分开说
  （「N 条没有时间戳（带时间窗时排除，放宽窗口也读不到）」）—— 对没有时间戳的条目建议
  「放宽窗口」是错的，那会把人引到一个永远找不到的方向。
  热记忆的条目时间戳由插件写入时确定；只在手工往 memory.json 里塞了一条没有 updatedAt
  的条目时，才会在读入时补成当时时刻，那种条目在时间窗里等于「刚刚」。
- links.jsonl / tombstones.jsonl 读不出来时（权限、被换成目录等），相关操作会**中止并
  报错**，不会当成空文件：当成空文件会让下一次写入把整份关系图覆盖掉、还会把「删了条目
  没清边」漏报成 0。
- 读取的次序：**重要度永远是主键**（critical → normal → low），同重要度内按 **recency
  弱先验**排序 —— updatedAt 每日衰减 0.995，也就是「新的在前」（第二十四轮 24-4）。
  它只定序、不参与取舍：critical 不会被一条新的 normal 挤下去。同一事实的两版并存时，
  后写的那版就是当前事实，所以它排在前（这是量到增益才落的那一项：改造前 latest@1 是 0/4，
  落地后 4/4）。时间戳解析不出来的条目排在**同重要度的最后**，不会被当成最新。
- 没有 lastAccessAt（上次取用）：现在的次序用的是**写入时刻**（updatedAt），不是取用时刻；
  记录取用时刻会**让 read 变成写操作**，而「读不写盘」是被面板、只读探针与真库对照共同
  依赖的性质。将来若要拿它做别的（比如「很久没被用过的先下沉」），那是另一条链，得先量。

每回合配额
- 一个回合里第一次带 query 的记忆检索、后续的换词细化、**知识库（kb）检索**、知识块
  关系遍历（kb-related）各有次数上限（默认 1 / 1 / 2 / 1）。回合号是从会话日志的 turn/start 事件读出来的。
- kb 检索单独一档：它的语料是几百块长原文、代价与记忆层不同，共用名额会让「查一次记忆」
  把「查一次来源」的机会一起吃掉。改它去设置页「记忆系统 → 每回合配额 → 知识块检索」。
- 用完会**明确拒绝并说明原因**（kb 那一档会直接给出「改走 kb-read」这条路），不要以为是
  检索坏了然后换个工具重复同样的检索。不带 query 的 read 是看现状、不是查资料，不占配额。0 = 不限制。

备份与迁移（action: export / import）
- export 把热记忆 + 台账 + 归档（索引 + 两投影全文：.md 与 .jsonl）+ 历史关系数据 + 墓碑
  整包写成一个 JSON 单文件（旧图关系已下线，但数据随包保留，导入不会丢）：{ action:'export', packPath?:'...' }（不传就写到记忆目录下的
  pack-<时间戳>.json）。
- import 读回来：{ action:'import', packPath:'...' }。口径是**只增不改**，按 id
  幂等：热记忆 / 台账 / 关系同 id（或同 source+target+kind）已存在就跳过，不覆盖
  现内容；归档按月合并、按块 id 去重（块没有 id 的老库用「时刻+原因+条数」当键，
  落盘时就把它当 id，所以重复导入同一份包不会产生第二份）。
  本地墓碑优先：落在墓碑上的条目与边不会被搬回来（旧 Pack 复活已删内容是被明确
  拒绝的）。**边还要逐条校验两端真的存在** —— 老 Pack 里带悬空边是常态（本轮之前
  remove 不清边），照搬会让一个干净的库继承旧库的悬空引用；跳过的条数会在回执里报出来。
- 这份 Pack 含私有记忆，不要当公开文件传。

浏览面板
- 设置对话框左侧「记忆系统」是配置页；会话页签行里的「记忆」是**浏览**页（排在
  「上下文」「编队」的右侧，要开着会话才看得见）：记忆（热记忆 / 台账 / 归档）与知识库
  两个一级页签，带搜索与工作目录切换。它走一个只读的本地端点
  （/office-memory/snapshot），且只服务「本进程真的跑过办公工具的工作目录」。

边界（与 mnemon 的差别，说清楚免得误期待）
- 没有远端记忆体、没有多 Provider 后端：全部落在工作目录与全局层的普通文件里。
- 没有空闲审查（后台自动复盘）：主动记录是一句**指引**，不是自主记录器。
- 记忆不会被缓存的 TTL 清掉：它不是中间产物。要清只能删记忆目录。

从 mnemon 迁移（action:'migrate'）
mnemon 关掉之后，原来存在它那里的记忆可以整批搬过来：
- 先看会迁什么：office_memory({ action: 'migrate', dryRun: true })；
- 再执行：office_memory({ action: 'migrate' })。
映射是降级映射，不是等量搬运：
- runtime 热记忆（工作区 .mnemon/runtime + 全局 ~/.mnemon/runtime）→ 热记忆：
  target=user 的进 USER.md，target=memory 的进 MEMORY.md；全局那份会被**复制**
  进本工作目录（换项目要各迁一次，记忆本来就是按工作目录的）。
- 长期记忆（Memory Spaces 的 insights）→ 归档：内容、类别、重要度、标签、实体
  与「有多少条关系边」的计数都随条目保留；**图边本身不迁**（mnemon 的图在它自己的库里，
  本插件的图关系也已下线），计数只作历史记录。
  已软删（deleted_at 非空）的不迁。
- Documents（项目档案）→ 办公记忆没有对应层：只报告数量并提醒原件仍在原地。
- 幂等：热记忆按内容去重、归档按 mnemon 的原始 id 去重，重复跑不会产生第二份；
  只做 add，不 replace / remove，不会动你后来改过的条目。
- 读长期记忆库要用 Node 自带的 node:sqlite（Node 22.5+）。拿不到时只迁 runtime
  部分并明确报警，不会整次失败。
命令行入口（改完源码不用重启就能迁移）：
  node scripts/migrate-mnemon.mjs --dry-run
  node scripts/migrate-mnemon.mjs`,

    search: `检索与结果解析（office_web_search / office_web_fetch / office_search_run / office_search_dispatch / office_parse_findings）

办公模式自带联网检索（第六十五轮起默认顺序 deepseek-api → seam → 免 Key 抓取）
- 查一个事实点：office_web_search({ queries: [...] })；打开一个页面：office_web_fetch({ url })。
  默认先走 **DeepSeek Search API 直连**（一次请求直接拿结果，不跑辅助模型；结果自带摘要），
  再退到宿主的 web 服务，最后走插件的免 Key 抓取（DuckDuckGo HTML 结果页 / 自建 SearXNG）。
  都不依赖宿主的联网工具，也不走 Tavily 这类要 Key 的三方 API。来源直接回到对话，不落盘。
- 要让材料落盘可复核、或要按渠道覆盖，走 office_search_run / 渠道覆盖检索。通道有多条，
  按顺序自动挑或者由你点名：
  deepseek-api=DeepSeek Search API 直连（默认第一，要 DEEPSEEK_API_KEY）/
  seam=宿主 web 服务 / duckduckgo=免 Key 抓取 / searxng=自建实例（免 Key）/
  openai=OpenAI 兼容 / tavily / brave / bocha=博查（中文长尾好）/ exa / serper=Google 结果代理。
  auto（默认）按 providerOrder 依次试，第一条成功的就用它；三方检索 API 不在默认顺序里，
  要用就显式点名（office_search_run({ provider:'bocha' })）或在设置页加回 providerOrder。
- 出口要走代理：设置页「检索编排 → 出口代理」填 http://127.0.0.1:7897 这样的地址。
  优先装成**只影响本插件**的分派器（改设置当场生效）；拿不到 undici 的部署退回
  「整个进程」的环境变量那条路（那种情况下清空要重启 DSH，代理值也会进 process.env）。
  两条路都让回环地址直连（自建 SearXNG 走 127.0.0.1 不受影响）。
  宿主接缝那一侧不归它管：要让宿主的联网也走代理，得在启动 DSH 前设 HTTPS_PROXY。
  office_help({topic:'search'}) 的末尾会给一行运行期实况。
- 只查一两个事实点：office_search_run({ queries: [...] })。一轮直查，
  来源与摘录写进 .office/search/<slug>/run.md，聊天里只回一份紧凑清单。
- **站点优先**（第三十四轮）：设置页「检索编排 → 站点清单」按类型维护一批站点
  （学术 / 图书 / 代码 / 自定义；内置 arXiv、Google 学术、知网、CrossRef、Open Library、
  GitHub、Stack Overflow 等，影子图书馆类只进目录、默认关闭）。开启后检索先把查询
  **限定到清单站点**，命中的来源排在前面并按类型汇总；限定一无所获（被墙 / 没收录）时
  退回不限定来源的泛搜，并在反馈里点名哪些站点空手。按次点名：
  office_web_search({ queries: [...], sites: 'academic' }) 或 sites: ['arxiv.org','github.com']；
  sites: false = 本次不限定。清单里的空手站点不等于「这个题目没有资料」。
- 取回来的网页**先过预处理**再进上下文（去脚本样式、去导航与样板、挑主容器、丢样板行、去重），
  结果文件与反馈里会给出「几页、原文多少字符 → 洗完多少、丢了多少行」。判错时用
  preprocess:'off' / 'plain' 退回原样（设置页里的 search.preprocess.mode 是默认值）。
- 来源本身是 **PDF**（题库、期刊、机构站点上很常见）时，正文交给 office.pdf 抽文本，
  不做网页预处理，反馈与结果文件里会点明「PDF 来源 N 条」。体积超上限（默认 24 MB）是
  直接放弃而不是截断 —— 截断的 PDF 抽不出文本；抽不出文本的多半是扫描件，那种要
  用 office.pdf.pages() 渲染成图再看。标着 PDF 而内容其实是 HTML 错误页时按
  「不支持的类型」报（不会误导成去装引擎）；notice 里的页数是**抽到了几页**，不是总页数。
- 要按渠道逐项覆盖（下一条的分流规则）：走下面两步走。

两步走：派工与解析，顺序不要颠倒
1. office_search_dispatch({ topic, type?, audience? })
   给 topic 就自动按内容类型出提纲（brief.md 落盘，渠道分流规则与固定要求都在
   文件里，可审计可复用）并逐渠道执行；briefPath 可复用上一轮的提纲重跑或改渠道，
   outputPaths 不给就按渠道数自动生成。
   type 可选 hotspot（热点）/ knowledge（知识）/ manual（手册）/ mixed；
   省略时按主题文字自动判断。组合里有子代理服务时派子代理（每个渠道一个，
   各自用 office_web_search / office_web_fetch 抓网页）；没有时自动改用
   插件的内置检索通道，在进程内跑同样的渠道，结果文件同一种格式。
   材料都只落盘，主上下文不会被搜索结果灌满。
2. office_parse_findings({ paths, type?, topic?, draft? })
   把结果文件读成摘要：按渠道归类、算来源覆盖度、标出只有单一来源背书
   与没带 URL 的条目。摘要是几十行，不是整页原文。
   给了 draft（产物正文，md / txt）还会做**句级引用审计** —— 见下面「引用纪律」。

引用纪律（第五十七轮 19-8 的 P2-9）
- 产物写完之后，把**正文**（Markdown / 纯文本）当 draft 传进来，插件逐句对账：
  ① 声称外部事实的句子有没有给出处；
  ② 正文里写下的每个地址，在**本次材料**里对不对得上（材料就是刚落盘的那几份结果
     文件里的**全部**地址，含写在散文 / 表格里的那些；对不上就是编的）。
  没过时工具结果 ok:false，并**带句号**点名要改的那几句与哪条地址（最多列 8 条）。
- 事实句的判据是**显式的、可以逐条核对的**五条，一项都不做语义判断：
    · 年份 —— 必须带「年」字（2024 年）：光看四位数字会把 1920×1080 判成年份；
    · 百分比 —— 数字后面跟 % 或 ％；
    · 带**外部**量词的数字 —— 钱 / 人 / 倍 / 吨 / 公里 / 平方米 / ℃；刻意不收
      「个 / 项 / 条 / 页 / 份」，那多半是在说这份产物自己的形状；
    · 句首归因 —— 据 / 根据 / 依据 + 16 字内出现来源性名词（通报 / 数据 / 用户…）
      （光有「按 Enter 继续」「根据判据」不算：那是在讲方法）；
    · 句中归因 / 统计词 —— 表明 / 报道 / 统计 / 宣布 / 公布 / 同比 / 环比 / 占比 /
      首次，以及「数据 / 报告 / 通报… + 显示 / 表明 / 称」这种带主语的引述。
  整行**不判**的三类（有意为之，属边界）：Markdown 标题、分隔线（三个连字符）、引用块
  （行首 >）与表格行（行首 |）—— 把带数字的事实放进引用块或表格，这次审计就不看它；
  另外长度不到 6 字的片段不判。
- 句界：句末标点（。！？!?）才算，**分号不算**（「一是…；二是…。（来源：X）」是最常见
  的写法，按分号切会把前两个分句判成没出处）；句号**后面**的「（来源：…）」会并回上一句，
  但**绝不跨标题 / 分隔线**并（跨过去会让上一节那句白捡一个出处）。围栏代码块与行内
  代码里的地址是示例，先剔掉再判。
- 指标口径（ALCE 三项里哪几项算得出，必须写清，不许含糊）：
    · 召回 = 事实句里带引用的比例（ALCE 原定义要蕴含判断，这里只是它的**词法代理**）；
    · 精确率 = 正文里的地址**逐字**对上材料的比例。同站不同页只算「站点级对上」，
      单独报数、**不计入分子**；同一路径不同参数（?id=1 / ?id=2）算**对不上**
      （那可能是另一个页面）；kb 块 id 只核形状，出处标记与编号引用**不进分母**；
    · 正确性 = **算不出**，恒为 null。它要蕴含判断与事实真值 —— 不许拿上面两项冒充。
- 三态，不许混说：**通过**（真有事实句、地址也逐字对上）/ **未过**（ok:false）/
  **没判出可核的东西**（没有可判的事实句、或正文里一个地址都没有 —— 这不是「引用都齐了」，
  也不会印成「通过」）。
- 边界照实说：Office 三件套 / PDF / ODF 的正文回读不了（read() 只给结构与统计），
  传这些文件会被当场拒绝并指三条路（传写文档用的 Markdown 源 / office.convert 转成 md /
  PDF 先 office.pdf.text 抽文本）。draft 给了空串或非字符串同样当场报错 —— 静默当
  「没给」会让调用方以为审过了。

渠道分流规则（固定，不靠临场发挥）
- 热点事件：权威媒体（定事实骨架）+ 各社交平台（现场与争议）；两边都要。
- 知识类：百科拿定义与体系；需要更深再下沉到文献与一手资料。
- 手册类：只认官方平台的参考文档；博客与问答只能用来看懂报错。
- 任何类型都先做一轮不限定来源的泛搜，用来找准关键词与分歧点。
- 平台渠道用 site: 限定（如 site:v2ex.com 查询），结果要标注「经搜索引擎间接取得」；
  抓取通道没有独立的时间窗参数：时间词（最新 / 本月 / 2026-07）写进查询，
  抓回来的旧文按页面日期自己筛掉。

为什么这样做
- 茧房效应的来源是只用一个渠道就把结论定下来。规则写死在插件里，
  每次检索都会被要求覆盖多个互不相关的来源。
- 结果落盘：一次检索的原始材料不进主上下文，主会话留着余地做排版与措辞；
  文件留在工作目录里也可以复核与追查出处。
- 内置通道只搬来源与摘录，不会自己换角度追查，也不判「这条说法该不该信」；
  要更强的覆盖就开子代理（组合里有 subagents 服务）或按提纲手工补查。

子代理的工具面（极简，只有这五个）
- 检索：office_web_search（插件的免 Key 抓取搜索）
- 取正文：office_web_fetch（PDF 由 office.pdf 抽文本）
- 读：read（文本）、read_image（图表 / 截图 / 扫描件）
- 写：write（唯一交付方式）
刻意不给：edit / glob / grep（只写自己的结果文件，不用改文件或找文件）、
office_help / office_run / office_memory（不生成文档）、bash / pwsh（不执行命令）、
子代理与任务板（不派活）、宿主的 web_search / web_fetch（preset 不声明 tool-web）。
裁剪靠 toolFilter.allow 白名单（子代理默认继承父方组装），
用白名单而不是黑名单：以后新装的工具也不会悄悄漏给子代理。

内置检索怎么配（设置页「办公模式 → 检索编排」）
- search.provider     auto（默认，按顺序自动挑）/ 固定成某一条通道 id
- search.providerOrder  auto 的尝试顺序（默认 deepseek-api → seam → duckduckgo → searxng；
                       三方 API 不进默认顺序）
- search.providers.<id> 每条通道各自的 apiKey / apiKeyEnv / baseURL / model / timeoutMs
- search.preprocess   mode（article | plain | off）/ dropBoilerplate / dedupeLines /
                       minLineChars / keepTitle
- search.fetch        取正文与 PDF 的上限（与通道无关）：maxResults / fetchPages /
                       fetchTimeoutMs / maxBytes / maxChars / maxRedirects /
                       pdfMaxBytes 与 pdfMaxPages（取正文撞上 PDF 时的体积与页数上限）
- search.proxy        出口代理（留空 = 直连）。填了让本插件自己发起的联网走它；
                      只认 http / https，本地地址不经代理
- search.engine       auto（默认，先子代理、跑不了用内置）/ subagent / builtin

取正文的成败判定：**不能只看状态码**
- 4xx/5xx 直接判失败（403 被挡、404 页面不存在、5xx 服务端错误），不把响应体当正文。
- 200 也要过一遍内容级哨兵：最终地址是区域封锁页、正文是 Cloudflare 拦截页 / 人机校验页 /
  登录同意墙 / 验证码、或者清洗后几乎没字 —— 都判失败。
- 判失败时抛出的错误会带**分类**，反馈里按分类计数（同因合并，不逐条重复那句话）：
  config 配置缺失 / network 网络出口不可达 / blocked 目标站拒绝 / empty 没拿到结果。
  被目标站拒绝的那几条要**换一个来源**，重试同一个地址只会再失败一次；
  网络类的才值得换时间重试。
- 一条来源取正文失败**不会**让整轮直查或整条渠道失败 —— 它只影响那一条的摘录。

够了的标准（摘要里会直接提示）
- 每个结论都要带来源 URL；只有单一来源的会标 [单一来源]。
- 「★ 仍缺渠道」= 该类型下必须覆盖的渠道还没搜到，要补一轮再写文档。
- 「待核 N 组」= 这些说法只有一个站点背书，写进正文时要照实标注。`,

    theme: `视觉主题（模板）

${renderThemeTable()}

默认是 plain（素色网格）：不传 theme 时 Excel 保持普通表格的网格形态、
Word 保持白纸黑字，都不加任何背景填充；要配色必须显式选主题。
用法：create({ theme: 'business' })；不传就是 plain。
三件套共用同一套主题，一份 Word、一份 Excel、一份 PPT 看起来是一套的。`,

    av: `音频与视频的内容提取（office.av）

会议录音、访谈视频、手机语音条这类**只有声音**的输入，和视频里**只有画面**的信息，
都从这里进上下文。两件事：

- **声音 → 文字**：本机 SenseVoice 离线转写（语音输入下载的那份 ONNX 模型），
  按块解码、Silero VAD 切句，每句带绝对时间戳；语种可 auto（中 / 粤 / 英 / 日 / 韩）。
  块与块之间**不是完全切开的**：相邻块留一小段重叠，切在句子中间的块会把下一块的
  起点回退到那句话的话头 —— 一句话不会被切成两半，也不会重复出稿。
- **画面 → 图片**：视频按时间点抽帧成 JPEG / PNG，路径交给 read_image 看 —— 与
  PDF 图像型那条路一样，是读图能力，不是 OCR。

先问一句：await office.av.check()
    返回 available / missing / ffmpeg / ffprobe / sensevoice（模型、VAD、运行时、
    精度、线程、语言）/ chunkSeconds / overlapSeconds / maxSeconds / maxFrames / hint。
    **四样能力缺任何一样都直接报错**（ffmpeg、ffprobe、模型、sherpa-onnx 运行时）——
    不静默降级、不假装转写出了空文字。hint 里写了缺什么、怎么补。

看清楚是什么：await office.av.info(path)
    → { kind:'audio'|'video', container, durationSeconds, hasAudio, hasVideo,
        audio:[{codec,sampleRate,channels}], video:[{codec,width,height,fps,frames}], hint }
    视频里没有音频轨、音频里没有视频轨，这里一眼就能看出来，别盲调后面的方法。

转写：await office.av.transcribe(path, { language?, chunkSeconds?, overlapSeconds?, maxSeconds?, out?, words? })
    → { text, segments:[{index,start,end,clock,text,lang,emotion,event,words?}], audioSeconds,
        speechSeconds, inferenceSeconds, realtimeFactor, chunks, rolls, model,
        detectedLanguages, wordTimings?, audio:{path,reused} }
    - segments 是逐句结果：时间戳可直接用来做字幕、回听定位与逐字稿；
      lang / emotion / event 是 SenseVoice 的富文本输出（<|zh|> / <|NEUTRAL|> / <|Speech|>）。
    - out 给出时把带时间戳的 Markdown 逐字稿写进那个文件（**长稿一定要给 out**：
      工具反馈里的脚本返回值会被截断，写到文件里再 read 才拿得到全文）。
    - chunkSeconds 默认 120 秒（官方那条链一次 131 秒封顶，这里留了余量）；
      overlapSeconds 默认 0.5 秒，是相邻两块的重叠（上限半块）—— 块不再完全切开，
      跨在切点上的短音能在同一块里解完；跨得更深的长句由边界回退兜住。
      rolls 是这次转写里边界回退的次数，0 表示每块都正好切在句子之间。
      maxSeconds 是单次时长上限（默认 3600 秒），超过就报错要求先切段。
    - 解码出来的规范 WAV 按内容键缓存在 .office/runtime/cache/av/audio 下：同一份文件
      再转一次是 audio.reused=true，不重解码。
    - words:true 才给词级时间：每句多一个 words:[{text,start,end}]（中日韩按**字**、
      西文按**词**，标点并进前一个词），顶层多一个 wordTimings 计数。
      **它不是模型直接给的**：SenseVoice 通道不返回词级时间，这里是拿逐 token 的
      起始秒（tokens + timestamps）现算的，起点=该词第一个 token 的起点、
      终点=下一个 token 的起点。默认关：逐字稿用句子时间戳够用，开了体积明显变大
      （实测 28 句的 165.8 秒音频，句级 8.8 KB → 词级 25.6 KB；推理耗时不变）。

抽帧：await office.av.frames(path, { at?, every?, count?, from?, to?, format?, width?, outDir? })
    → { dir, count, format, files:[{index, at, clock, path, bytes}], hint }
    - 时间点优先级：at（明确给秒，数组或单值）> every（定步长秒）> count（默认 6 张均匀铺开）。
    - 拿到 path 之后**逐张 read_image({ file_path })**。图上有字也认得出来，那是读图能力。
    - 单次上限 maxFrames（默认 12 张）；4K 视频建议给 width（如 1280），图小、读得快。
    - 抽出来的帧落在 .office/runtime/cache/av/frames/<文件名>/ 下，是中间产物，不进交付物。

一次做完：await office.av.extract(path, { ...转写参数, ...抽帧参数, out?, frames?, transcript? })
    → { info, transcript, frames, notes, next }
    视频一条调用就能同时拿到逐字稿与若干画面；frames:false 只转写、
    transcript:false 只抽帧、frames:数字 表示抽几张。这是会议纪要与课程录播的入口。

典型流程
1. 会议视频 → 纪要：office.av.extract('会议.mp4', { language:'zh', out:'逐字稿.md', count:6 })
   → read 逐字稿.md → 用 office.word 写纪要（引用逐字稿里的时间点）。
2. 语音条 / 录音 → 文字：office.av.transcribe('录音.m4a', { out:'录音.txt' }) → read 那个文件。
3. 视频里的板书 / 幻灯片：office.av.frames('课程.mp4', { every:60, width:1280 })
   → 逐张 read_image，把看到的写进文档（图上文字靠读图，不保证逐字准确）。

边界（照实说，不要越界承诺）
- 只做识别，**不做翻译**；不认说话人（多人的录音里分不出谁说的）；不做流式（一次一段完整音频）。
- 不做降噪与增强：底噪大、多人重叠的录音识别质量会明显下降。
- 切句是 VAD 启发式：句子的边界与标点由模型给，不保证与人的断句一致。
- 不支持的输入形状由 ffmpeg 兜底转码；ffmpeg 自己都不认的容器/编码会明确报错。
- 长音频靠分块 + VAD。块之间留了重叠、切在句子中间时下一块会回退到话头，所以
  **一句话不会被切点切成两半、也不会重复出稿**；但块与块之间仍然没有共享上下文
  （上一块的词不会成为下一块的解码提示），专业名词跨在切点上时前后用词可能不一致。
- 设置页只列常用项；VAD 阈值、单句上限、词级时间戳（words，见上面转写那一条）、缓存子目录
  （audioDir / framesDir）、模型精度与线程的细项在配置文件里改。`,
};

/**
 * theme 的默认层：只报可用 id 与挑选口径，省掉那一张逐色的表
 * （表是给人看的，模型只需要知道有哪几个 id）。
 */
const COMPACT_THEME = `视觉主题（默认 plain 素色网格，无背景填充）

可用主题：${themeCatalog().map((theme) => theme.id).join(' / ')}。
用法：create({ theme: 'business' })。不传 theme 就是 plain：Excel 保持普通表格网格、
Word 白纸黑字，都不加背景填充 —— 用户明确要配色时才显式选主题
（正式汇报 business / 内部稿 minimal / 论文函件 academic / 演示 tech 或 warm）。
三件套共用同一套主题，一份 Word、一份 Excel、一份 PPT 看起来是一套的。
每个主题的配色表：office_help({ topic: 'theme', detail: true })。`;

/**
 * 默认层的文档（`detail` 不给时的形态）。
 *
 * 只写「照着它就能写出对的调用」—— 操作名、签名、必需的边界。例子、为什么、
 * 配置明细与迁移说明留在全文层（BASE_TOPICS / 模块 meta / OFFICE_GUIDE）。
 *
 * 硬约束：这里出现的接口名与参数名必须与全文层一致；测试里有体积预算与
 * 「关键接口不许丢」两类断言钉住这件事（test/injection-budget.mjs）。
 */
const COMPACT_TOPICS = {
    theme: COMPACT_THEME,
    run: `office_run 的用法（批处理脚本）

- script（与 scriptFile 二选一）：一段 JavaScript，可以写多步，最后 return 一个值（会回显）。
- scriptFile（与 script 二选一）：读工作目录里的一份 .js 当脚本跑 —— 长脚本先
  office.files.write('build.js', code) 落盘，之后每次调用复用同一段，不必重复内联。
- args（选填）：传给脚本的参数，脚本里以全局 args 读到。
- purpose（选填）：一句话说明这次做什么。
- keepCache（选填）：false 结束即清空缓存目录（默认保留，保留才能跨调用复用）。

脚本里能用：office.word / office.excel / office.ppt（造、读、改文档）、office.tex
（学位论文，见 tex）、office.pdf（见 pdf）、office.convert（md / docx / tex 互转、
认识的文档格式出 PDF，见 convert）、office.python（见 python）、office.av
（音频转写与视频抽帧，见 av）、office.preview（把自家文档渲染成页面图 / PDF，
见 preview）、office.image（取图与用图，见 image）、office.writing（写作规划与自检，见 writing）、
office.archive（读 zip / gz / zst，见 archive）、office.files、office.cache、office.theme、
office.log / warn / note、office.assert(条件, 说明)。

没有：process / require / fs / 网络 / 计时器；也没有 office.memory（记忆走
office_memory 工具）。路径相对当前工作目录。能调外部程序的只有 PDF 渲染、
LaTeX 编译、Python、ffmpeg（office.av）、文档渲染预览（office.preview）、
文档转换（office.convert）六处，都由插件自己探测，脚本里给不了命令。
反馈：每个写出的文件都会被重新打开复检（真实统计 + warnings）；脚本返回的 JSON
超过 ${RETURN_ECHO_CHARS} 字符会截断，长内容写进文件再读。
三个真实用例（docx / xlsx / Python 绘图）在全文层：office_help({ topic: 'run', detail: true })。`,
    pdf: `PDF 的读法（office.pdf）—— 先 info 判断类型，再决定怎么读

- await office.pdf.info(path) → { pages, pageSizeCm, hasTextLayer, counts, hint }
    hasTextLayer:false = 扫描件 / 手写笔记，别抽文字，直接 pages。
- await office.pdf.text(path, { from?, to?, out? }) → { path, chars }
    文本型用它；out 给出时写进文件（长文本必须给 out，返回值会被截断）。
- await office.pdf.pages(path, { from?, to?, dpi?, format? })
    → { engine, dir, files: [{ page, path, bytes, reused }] }
    图像型用它，每页渲染成图；拿到 path 后逐页 read_image —— 那是读图能力，
    认得出手写字，不是 OCR。dpi 默认 120（手写 120-200）。
    单次最多 ${DEFAULT_PDF_MAX_PAGES} 页，超了分批传 from/to；同样页码再渲染是 reused。

一句话流程：info → 文本型 text(out) 再 read；图像型 pages 再逐页 read_image。
细节（引擎探测、返回字段全表、典型流程）用 office_help({ topic: 'pdf', detail: true })。`,
    python: `Python 计算与绘图（office.python）—— 办公模式里唯一的编程出口

- await office.python.check() → { available, interpreter, version, packages, outDir }
    先问一句本机装了哪些包；没装的不要 import（ModuleNotFoundError 白跑一次）。
- await office.python.run(code, { name?, outDir?, timeoutMs?, args? })
    → { ok, code, stdout, stderr, files, script, outDir, logFiles, elapsedMs }
    ok = 退出码为 0；报错时 stderr 是完整回溯，行号就是代码里的行号。
    files 是这次跑出来的产物，图片标 kind:'image'，路径可直接交给
    Word 的 builder.image(path, …) / builder.images(items)，或 PPT 的
    deck.image({ path }) / deck.images(items)。
    stdout 超过 ${PYTHON_STDOUT_CHARS} 字符会截断，完整输出在 logFiles.stdout。
- await office.python.file('analysis.py', { args? })   长脚本先落成 .py 再跑。

脚本里多出三个名字：OUT_DIR（产物目录，图必须 savefig 到这里才会出现在 files 里）、
WORK_DIR（当前工作目录）、use_cjk_font()（matplotlib 中文与负号）。
一定要 savefig，不要 plt.show（没有显示器，后端是 Agg）。
没有网络（pip install 装不了包）、没有 stdin、没有命令行。
细节（返回字段、边界清单、绘图示例）用 office_help({ topic: 'python', detail: true })。`,
    av: `音频与视频的内容提取（office.av）—— 声音转文字、视频抽帧

先问一句：await office.av.check() → { available, missing, ffmpeg, ffprobe,
    sensevoice, chunkSeconds, overlapSeconds, maxSeconds, maxFrames, hint }。
    ffmpeg / ffprobe / SenseVoice 模型 / sherpa-onnx 运行时缺任何一样都**直接报错**，
    不静默降级；hint 里写了缺什么、怎么补。推理由插件的子进程跑（在脚本沙箱之外，
    不占宿主事件循环），所以脚本里只管 await。

- await office.av.info(path) → { kind, container, durationSeconds, hasAudio, hasVideo,
    audio:[{codec,sampleRate,channels}], video:[{codec,width,height,fps,frames}] }
- await office.av.transcribe(path, { language?, chunkSeconds?, overlapSeconds?, maxSeconds?, out?, words? })
    → { text, segments:[{start,end,clock,text,lang,emotion,event,words?}], audioSeconds,
        speechSeconds, inferenceSeconds, realtimeFactor, chunks, rolls, detectedLanguages,
        wordTimings?, audio:{path,reused} }
    language: auto（默认）/ zh / en / yue / ja / ko；chunkSeconds 默认 120；
    overlapSeconds 默认 0.5（相邻块的重叠，块不再完全切开：跨在切点上的短音能在
    同一块里解完，跨得更深的长句由边界回退兜住，不丢字也不重复）；
    out 给出时把带时间戳的 Markdown 逐字稿写进文件（**长稿必须给 out**，
    脚本返回值会被反馈截断）；同一份文件再转是 audio.reused=true。
    words:true 才给词级时间（segments[].words:[{text,start,end}]，中日韩按字、西文按词）——
    SenseVoice 通道不返回词级时间，这是拿逐 token 起始秒现算的；默认关（体积）。
- await office.av.frames(path, { at?, every?, count?, from?, to?, format?, width? })
    → { dir, count, files:[{at, clock, path, bytes}] }
    时间点：at > every > count（默认 6 张均匀铺开）；单次上限 maxFrames（默认 12）。
    每张 path 交给 read_image —— 图上文字是读图能力认出来的，不是 OCR。
- await office.av.extract(path, { ...转写参数, ...抽帧参数, out?, frames?, transcript? })
    → { info, transcript, frames, notes, next } —— 会议视频一条调用拿到逐字稿 + 画面。

边界：只识别不翻译；不认说话人；不流式；不做降噪；切句是 VAD 启发式；
块之间没有共享上下文（切点前后用词可能不一致，但一句话不会被切成两半）。
细节（返回字段全表、典型流程、配置项）用 office_help({ topic: 'av', detail: true })。`,
    files: `office.files 的用法（改 Markdown 或长文本就用它，一次改完所有位置）

- office.files.read(path)               读文本（二进制会报错并指路，不给乱码；判据是空字节或已知魔数）
- office.files.readBytes(path)          读原始字节（二进制安全，交给 write / copy 用）
- office.files.write(path, 文本或字节)   整份覆盖写（顶层字节按字节写；嵌在对象里的字节、Map / Set、函数与 undefined 会被拒绝）
- office.files.copy / move(from, to, {overwrite})   整份复制 / 移动（二进制安全；
      目标已存在要显式 overwrite:true）
- office.files.edit(path, edits)        批量字面替换，只写一次盘
      写法：[[旧, 新], [旧, 新]] 或 [{ find, replace, all }]；没命中的 find 在结果
      missing 里报出来，不会静默跳过。
- office.files.template(path, {k: v})   把 {{k}} 一次性替换掉
- office.files.exists / list / remove / stat
- office.archive.info / list / text / find / extract   读 zip / gz / zst（压缩包不能用 read）`,
    archive: `归档与会话日志的读法（office.archive）—— 压缩包不能用 office.files.read 读

- office.archive.info(path)                  这份包里有什么（kind / 条目数 / 解压后体积）
- office.archive.list(path, {limit})          逐条列条目（名字 / 字节数 / 压缩方式）
- office.archive.text(path, {entry, offset, maxChars})   取一段正文（可翻页）
- office.archive.find(path, {pattern, entry, maxHits, context})   按行检索（日志最常用）
- office.archive.extract(path, {entry, to})   把一条正文落到工作目录，再 office.files.read

判型按魔数（ZIP / gzip / zstd / 纯文本），不看扩展名；zstd 要 Node 22.15 以上，
本机没有会明确报「本机不支持」。返回全部有界，别整份往返回值里带。
典型：info 看有什么 → find 找行 → text 取一段 → 要全文就 extract 落盘再 read。
细节（默认挑哪一条、二进制条目怎么报、翻页与上限）用 office_help({ topic:'archive', detail:true })。`,
    preview: `渲染预览（office.preview）—— 把自家 docx / xlsx / pptx 渲染成页面图或 PDF

read 报的是结构（统计 / 大纲 / 排版风险），看不了画面；要看画面就渲染再 read_image。

- await office.preview.check() → { available, engine, backend, version, formats, hint }
    引擎是宿主随部署安装的 LibreOffice（不是要另装的插件）；available:false 时 hint 写了缺什么。
- await office.preview.render(path, { pages?, sheet?, range?, dpi?, maxPages? })
    → { dir, dpi, pageCount, count, images:[{index,page,width,height,bytes,path}], missingFonts, reused }
    pages 默认 'all'，也可给页码数组（从 1 开始）；xlsx 用 sheet（工作表名）配 range（A1 区域），
    pages 与 sheet 不能同时给。dpi 默认 110，maxPages 默认 12（上限 30）。
    **拿到 images[].path 逐张 read_image**；图落 .office/runtime/cache/preview/，内容与参数没变就 reused。
- await office.preview.pdf(path, { to? }) → { path, bytes, reused }
    不传 to 只放缓存；传 to 写到工作目录（交付物，出现在反馈的 otherFiles 里）。

渲染一次 2-5 秒，**按需调**（写完想确认版式、交付前看一眼），别每写一份都调。
细节（页选择、缺字体、缓存键、支持的全部输入格式）用 office_help({ topic:'preview', detail:true })。`,
    convert: `文档互转与统一出 PDF（office.convert）—— md / docx / tex 互相转，认识的文档格式都能出 PDF

三件引擎各管一段（都在运行期探测）：pandoc（互转）/ 本机 TeX（.tex 出 PDF）/
宿主 LibreOffice（Office 三件套与老格式 .doc/.xls/.ppt/.ods/.odp 出 PDF）。

- await office.convert.check()
    → { available, pandoc, latex, libreoffice, cjkFont, sources, targets, matrix, hint }
    只报探到的入口；缺哪样只影响它那几条路。
- await office.convert.file(path, { to, out?, overwrite?, engine?, media?, fragment?, reference?, font? })
    → { ok, from, to, path, bytes, engine, steps, reused, warnings?, media?, hint }
    to：md / docx / tex / pdf（也认 markdown / word / latex 与带点的写法）。
    默认写到源文件同目录同名换扩展名；out 指定别的路径。
    docx / odt → md / tex 默认把内嵌图片抽成 \`<目标名>.media/\`（media:false 只出文本）。
    md → tex 默认出能直接编译的整份（fragment:true 只要正文片段）。
    engine 只对出 PDF 有意义：auto / libreoffice（宿主引擎）/ tex（本机 TeX 排版）。
    .md 出 PDF 默认「先转 docx 再渲染」，LibreOffice 不在时自动退到 pandoc + xelatex。

目标已存在且内容不同要显式 overwrite:true；.tex 编译与 LibreOffice 产物带生成时间戳，
同一份源重转也不会字节相同。.tex 出 PDF 的默认落点由 TeX 自己写：那里**已经有一份 PDF** 时，
不传 overwrite:true 会被拒（传了就重写并登记，reused 恒为 false；想两份都留就传 out）。
PDF 正文读不出来：要文字用 office.pdf.text，要画面用 pages。

细节（矩阵、路线、中文字体、边界与引擎环境变量）用 office_help({ topic: 'convert', detail: true })。`,
    image: `配图（office.image）—— 图库检索 / 网页选图 / 抽内嵌图 / 元信息 / 署名

两条免 Key 的图库通道：commons（维基共享资源，署名信息最全）/ openverse（CC 图库聚合）。

- office.image.channels() → 通道清单（要 Key 的图库不在里面：取一张图不该先配 Key）
- await office.image.search(query, { channel?, limit?, width? })
    → { channel, count, results:[{title, url, fetchUrl, thumbUrl, pageUrl, width, height, mime, license,
        author, insertable, insertableInWord, insertableInPpt, credit, attribution}] }
    credit 是一行署名（直接传给文档的 source）；insertable 指 Word（PNG/JPEG），
    GIF 只在 PPT 里能用（insertableInPpt）；fetchUrl 是「取哪一份」（SVG / TIFF 原图时指向 PNG 缩略图）。
- await office.image.fetch(候选或URL, { to?, maxBytes?, overwrite? })
    → { path, bytes, mime, width, height, source, credit, attribution }
    默认落 assets/；先校验魔数（返回 HTML 说明页会明确报「不是图片」）；来源页与许可
    会进台账的 source。
- office.image.info(path) → { ok, path, bytes, mime, ext, width, height, aspectRatio, sha256,
    insertableInWord, insertableInPpt }；读不到 / 不是位图回 { ok:false, reason }，不抛。
    只按魔数认 PNG / JPEG / GIF / WebP；嵌 Word 前先确认 insertableInWord。
- await office.image.fromPage(url, { limit?, minWidth? })
    → { ok, url, finalUrl, title, count, results, skipped, hint }。候选与 search 同构，
      直接交给 fetch；og:image 优先，srcset / 懒加载属性也认；网页图许可未知，发布前自己核对。
- office.image.extract(docx|pptx|xlsx, { out?, overwrite? })
    → { ok, format, path, dir, count, files, skipped, hint }。抽包内 media，默认
      assets/<stem>-media/，文件名带内容哈希；已存在默认跳过，要重写传 overwrite:true。

配图必须带来源：拿到 credit 后嵌进文档时传 source（Word 在图注下加一行小字，
PPT 并进题注）。示例与细节用 office_help({ topic:'image', detail:true })。`,

    writing: `写作规划与自检（office.writing）—— 生成前分注意力，生成后按同一张表查

- office.writing.plan(input, options) → { brief, strategy, sections, checks, totalChars, markdown }
    input：string 或 { title, genre, audience, goal, targetChars, sections }；
    genre：academic / explainer / report / memo / essay / howto / general；
    section：{ heading, kind, difficulty, value, ... }，缺省按 genre 给可改骨架。
    每节按难度 / 读者收益分配 targetChars 与 depth（dwell 写透 / develop 展开 / touch 点到 /
    compress 收住）：难点、机制、跃迁多给，背景与常识 1–2 句。plan.markdown 直接 office.log。
- office.writing.audit(text, plan?) → { score, band, warnings, suggestions, signals }
    传 plan：重点节没写透 / 收住节超长 / 各节仍平均 / 最长节不是最难，都会报；
    不传也查：句长过匀、段长持平、套话、锚点不足、比喻过密、过度对冲、无让步、每节总结。
    score 是启发式，不是「像人程度」的证明；改稿看 warnings 与 suggestions。
- office.writing.deai(text, { mode, register, styleRef }) → { flags, stats, styleRef, edits? }
    只报 A1 对举 / A2 提示语冒号 / M1 元话语与文体偏离，不给分；register 调口径；
    safe 只删提示语前缀（「先说结论：」「原因不复杂：」），A1 只报不改。
- office.writing.signals(text)：只看表面量；office.writing.render(plan)：重渲染规划。
口径：每 150 字至少一个可核对锚点；核心主张配最强反例；连接词能删就删；学术段少比喻，
说明文可用日常比喻但每个都要标「哪里不像」。大模型 A/B 用 scripts/writing-eval.mjs
（无 Key 会 SKIP，--mock 可离线跑通）。细节、例子与 kind 全表用
office_help({ topic: 'writing', detail: true })。`,

    cache: `office.cache 的用法（中间产物目录，跨调用保留）

- office.cache.dir() / write(n, 数据) / list() / stats() / clear()
清理规则：每次 office_run 开始时只收走超过 TTL（默认 12 小时）没碰过的条目；
office_run({ keepCache:false }) 或 office.cache.clear() 才整目录清空。
渲染好的 PDF 页面图放这里，下一步 read_image 直接读；交付物写工作目录，不要放这里。`,
    settings: `设置界面（Settings > 插件 > 办公模式 / 记忆系统），同一个命名空间 dsh-office-mode

「办公模式」页
- 工具开关：office_help / office_run / office_memory / office_search_run /
  office_search_dispatch / office_parse_findings
  关掉的工具不出现在模型面前，也不占每次请求的 schema 开销；改完即时生效。
- 办公模式插件：列的是办公 preset 自己声明的插件行（读 preset 文件本身），开关直接回写
  那一行的 disabled（preset 真源 + profile 侧已有的覆盖行），写前自动落备份。
  **生效范围是新建的办公会话**，已经在跑的会话不受影响。
- 检索编排：maxParallel 默认 4 / resultLimit 默认 40 / maxChannels /
  outputDir 默认 .office/search / engine（auto|subagent|builtin）/
  provider（用哪条联网通道）/ providerOrder / providers.<id>（各通道的 Key 与端点）/
  preprocess（网页预处理强度）/ fetch.*（取正文与 PDF 的上限）
- 文档与缓存：defaultTheme / scriptTimeoutMs / maxScriptChars / cacheDir / keepCache /
  cacheTtlMinutes 默认 720 / pdfDpi / pdfMaxPages / pdfEngine /
  texEngine / texTimeoutMs / texTemplateDir
- 音频与视频：enabled / ffmpegPath / ffprobePath / modelDir / language /
  chunkSeconds 默认 120 / overlapSeconds 默认 0.5 / maxSeconds / timeoutMs /
  precision / threads / frames / maxFrames / frameFormat / frameWidth

「记忆系统」页
- enabled / dir（默认 .office/memory）/ autoLedger / promptHint /
  userLimitBytes 默认 4096 / projectLimitBytes 默认 10240 / ledgerLimit 默认 500 /
  archiveKeep 默认 60（归档**卷**数：一个月一个卷，超 200 条的月份开下一卷）。

所有数值参数都有上下界：设置页挡一道、运行期再收敛一次，写坏配置不会让插件崩。
宿主没有设置能力时这两页自动不出现，其余功能不受影响。
完整清单（每个参数的含义与默认值）用 office_help({ topic:'settings', detail:true })。`,
    memory: `办公记忆（office_memory）—— 三层，跨会话保留在会话工作目录的 .office/memory/

- 热记忆（layer:'hot'）：用户偏好与项目约定，动笔前照它办。真源是 memory.json；
  USER.md / MEMORY.md 是它的投影，只能经 office_memory 改（直接编辑会被覆盖）。
- 台账（layer:'ledger'）：office_run 每写出一份 Office 文档自动登记一条
  （路径 / 格式 / 主题 / 字节 / 复检统计 / 用途 / office.files 读过的输入文件 source），
  不用手工补。.md / .js / 图片 / PDF 这些不会自动登记，要进台账得手工
  office_memory({ action:'log', path, format, purpose, source? })。
- 归档（layer:'archive'，只读）：热记忆装不下时下沉的旧条目与滚动下来的旧台账，
  按月份装订成 .office/memory/archive/YYYY-MM.md。

读：office_memory({ action:'read', layer:'hot'|'ledger'|'archive'|'all', query? })。
    每行去掉行首「- 」之后的第二段就是条目 id（m-… / L-… / mnemon:…）：
    replace / remove / kb-link / kb-related 可以直接给 id。
    带 query 是检索（按命中词占比分档，占每回合配额）；不带 query 是看现状。
    给 tags / entities 就是**分区**（硬过滤：只出**声明过**这些标签 / 实体的条目，
    台账按 theme 算；三层各自过滤；entities 只是条目字段，不再有实体视图）。细节见全文层。
找资料（记忆 + 知识库）用 office_memory({ action:'recall', query })：统一入口 —— 插件先查记忆层，
    按「命中数 / 查询词覆盖 / 最高重要度」判够不够答，不够才下探 kb，并在回执首行自报路由。
    它与 read 带 query 共用同一条每回合配额；真的下探了才再吃一档「知识块检索」配额，
    被挡住时如实写「未下探」并给出真的能走的那条路。
写：add / replace / remove（target: user = 用户偏好与要求，project = 项目约定与环境）。
    replace 与 remove 给 oldText（一段能**唯一**命中原文的片段）或直接给 id，命中多条会被拒绝。
    importance：critical（明确的必须 / 永远不要）/ normal / low。
    scope：workspace / global，只动指定根（跨项目通用偏好写 global，项目流程留 workspace）。
只记：用户说过的偏好、纠正、要求「记住」的事；稳定的项目约定与环境事实；踩过的坑。
不记：检索结果原文、一次性进度、自己的推测与复述、刚从记忆里读出来的东西。
记完一句说明即可，不要把条目整段念给用户听。
知识库 kb（.office/kb/）装的是外部原文的块 + 字符区间，不是结论；找来源用 kb-search
（**词法**检索：中文 bigram + BM25，语义改写召不回；零重叠就弃答 —— 查不到就换更贴近原文的
词，别换工具重复检索），逐字引用走 kb-read（有界），入库 kb-ingest（tier: user / verified /
unverified，单源 unverified 永不进热记忆；scope 可指定 workspace / global），删除 kb-drop
（也可带 scope），清单 kb-list。
把块 id 写进 add、log 的 source 就是「依据哪块来源」。kb 检索单独占一档每回合配额（默认 2）。
其他动作：log（手工登记外部产物）/ status / export·import（Pack 备份，只增不改，
    旧图关系数据随包兼容保留）/ migrate（从 mnemon 迁移）。
记忆按工作目录存（一个项目一份）—— 运行时的记忆目录默认为会话工作目录下的
.office/memory（设置页可改）。跨项目共享有两层：设置页把 scope 改成 both / global
（常开全局根），或单条写 scope:'global'；全局根默认在 $DSH_HOME/.office/memory。
细节（范围与层开关、召回质量、每回合配额、面板、与 mnemon 的差别、迁移映射）用
office_help({ topic:'memory', detail: true })。`,
    search: `检索（office_web_search / office_web_fetch / office_search_run / 三步走）

办公模式自带网页抓取通道（免 Key，preset 不声明宿主的 tool-web）：
查一个事实点 office_web_search({ queries })，打开一个页面 office_web_fetch({ url })，
都不落盘、不走 Tavily 这类要 Key 的三方 API；
直查与渠道覆盖：auto 按顺序自动挑（deepseek-api 直连 → 宿主 web 服务 → 免 Key 抓取），
也可以用 provider 参数点名（bocha 中文好 / serper 就是 Google …，三方 API 不在默认顺序里）。
取回来的网页先过预处理（去导航与样板），preprocess 参数可按次改强度；来源是 PDF 时
正文交给 office.pdf 抽文本。出口要走代理在设置页「检索编排 → 出口代理」填。

- 只查一两个事实点、要落盘：office_search_run({ queries: [...] }) —— 一轮直查，来源与摘录
  落 .office/search/，聊天里只回紧凑清单；写进文档前按 URL 核对。
- 要按渠道逐项覆盖：两步走，顺序不要颠倒
  1. office_search_dispatch({ topic, type?, audience? }) 自动出提纲（brief.md 落盘，可审计
     可复用）并逐渠道执行；briefPath 可复用上一轮提纲，outputPaths 不给就自动生成。
     type：hotspot（热点）/ knowledge（知识）/ manual（手册）/ mixed（判不准）
     有子代理服务就派子代理（用 office_web_search / office_web_fetch 抓网页）；否则自动走
     内置通道，并在反馈里说一次原因（那是装配期就知道的事实，不是错误 —— 不用绕路去派
     spawn_teammate）。结果文件格式一样。材料只落盘，不进主上下文。
  2. office_parse_findings({ paths, type?, topic?, draft? }) 读成摘要（按渠道归类、算来源
     覆盖度、标出单一来源与没带 URL 的条目）；给了 draft（产物正文 md / txt）还会做
     **句级引用审计**：事实句没出处、或地址对不上本次材料，都当场点名且 ok:false。
- 派工反馈说「N 个渠道空手而归 / 一条来源都没拿到」时，那是**通道没通**，不是「这个题目
  没有资料」；摘要里的「未找到」不算结论（「到手结论 0 条 + 空手渠道 N 个」就是这个形状）。
  先 office_search_run 直查一轮确认哪条通道是活的，或在设置页配出口代理，再重新派工。
- 站点优先（第三十四轮）：设置页按类型维护站点清单（学术 / 图书 / 代码），检索先限定到
  清单站点、命中排前并按类型汇总；限定全空时退回泛搜并点名空手站点。按次点名用
  sites: 'academic' / ['arxiv.org'] / false。
- 分流规则（固定）：热点 = 权威媒体 + 社交平台；知识 = 百科，需要更深下沉到文献；
  手册 = 只认官方文档；**任何类型都先泛搜一轮**找准关键词与分歧。
- 平台渠道用 site: 限定（如 site:v2ex.com 查询）；时间词写进查询，旧文按日期筛掉。
- 茧房效应就来自只用一个渠道定结论 —— 规则写死在插件里，不要跳过渠道覆盖。
- 通道报错时看它给的**分类**，四类修法不同：**配置缺失**（去设置页配 Key 与端点）；
  **网络出口不可达**（连不上/超时/5xx —— 换时间、换出口，或在设置页配出口代理）；
  **目标站拒绝**（401/403/429/451、区域封锁页、人机校验页、页面不存在 —— **换一个来源，
  别重试同一个地址**）；**没拿到结果**（换关键词或换角度）。取正文失败会按分类计数报出来。
- **HTTP 200 也可能是失败**：区域封锁页、拦截页、登录墙常常连状态码都是 200。
  插件已做内容级哨兵，被判失败的那一条不要当正文用，也不要拿它当「这个站点没有内容」。
细节（通道实现、子代理白名单、内置检索参数、够了的标准）用
office_help({ topic:'search', detail: true })。`,
};

/** 默认层的索引短语（全文层那份在 indexText() 里）。 */
const COMPACT_INDEX = `办公模式能力索引（默认只给一行摘要；要细节传 detail:true）

- word：Word 文档（.docx）—— 生成 / 读取 / 批量改写：封面、标题层级、列表、引用、表格、图片与多图网格（builder.images）、公式、页码、目录
- excel：Excel 工作簿（.xlsx）—— 表格、公式、统计、条件格式、原生图表（sheet.chart）、原生图片（sheet.image，PNG/JPEG）
- ppt：演示文稿（.pptx）—— 版式助手、图片、生成后审阅与就地微调
- tex：LaTeX 学位论文（thuthesis）—— 生成源码并编译 PDF
- pdf：PDF（.pdf，只读）—— 先 info 判断文本层，文本型抽文字、图像型渲染成页图再 read_image
- python：Python 计算与绘图—— 科学计算、画图，图直接嵌进 Word / PPT
- av：音频与视频的内容提取—— 录音 / 视频转文字（带时间戳），视频按时间点抽帧给 read_image
- run：office_run 批处理脚本怎么用
- files：文件读写与批量替换
- archive：zip / gz / zst 归档与会话导出怎么读（office.archive）
- preview：把自家 docx / xlsx / pptx 渲染成页面图或 PDF（office.preview）
- convert：文档互转与统一出 PDF（office.convert）—— md / docx / tex 互转，tex / Office 三件套与老格式 / OpenDocument 都能出 PDF
- image：配图（office.image）—— 元信息 / 免 Key 图库检索 / 网页选图 / 抽文档内嵌图 / 取图 / 署名（credit）
- writing：写作规划与自检（office.writing）—— 按难度 / 读者收益分篇幅、逐节写法、句长 / 套话 / 锚点 / 让步自检
- cache：中间产物目录与清理规则
- memory：三层记忆（热记忆 / 台账 / 归档）+ 知识库 kb（入库 / 检索 / 按块读）怎么读写、记什么不记什么
- search：联网检索（内置通道 / 三步走 / 渠道分流规则 / 结果解析）
- settings：设置界面（工具开关与参数）
- theme：视觉主题（默认 plain 素色，无背景填充）
- guide：说话与做事的方式（第一次进入办公模式建议读一遍）

查详情：office_help({ topic: 'word' }) 这样传话题名；要看例子与全部解释再加 detail:true。
传错话题时这份索引会原样退回，一次调用就能自我纠正。`;

/** 默认层的 guide（全文层是 OFFICE_GUIDE，约 9.4 KB —— 第六十轮实测 9,373 B）。 */
const COMPACT_GUIDE = `办公模式的工作方式（默认层；完整版 office_help({ topic:'guide', detail:true })）

说话：直接说结果，先说做了什么、文件在哪；短句，少术语；不寒暄、不复述用户的话；
内容多就建文件，聊天只留结论与位置。
动手：一次工具调用完成一批操作（office_run 一次写完），不要多次试探；
改已有文件先读后一次写回；交付物写工作目录，中间产物写 .office/runtime/cache。
文档：先 office_help 查写法，再用 office_run 一次写成；不传主题就是 plain（素色无底纹），
要配色显式选主题；生成后看反馈里的 stats / outline / warnings，有 warning 就按提示调。
写作：动笔前 office.writing.plan（按难度/收益分篇幅），重点给足、铺垫收住；篇幅表是草稿，
总长以它为准，写完用 office.writing.audit 自查；去 AI 味用
office.writing.deai（register/styleRef 只报，safe 只删提示语）。
PPT 多走一步：生成 → office.ppt.read 审阅（版式与内容量是否匹配）→ office.ppt.revise
就地微调（改文字 / 字号 / 位置）→ 再 read 复检一次。一轮就够，不要整篇重生成。
PDF：先 office.pdf.info 看有没有文本层，文本型抽文字、图像型渲染成页图再 read_image。
Python：算数、统计、画图用 office.python.run（先 office.python.check() 看装了哪些包）；
长脚本先 office.files.write 落成 .py 再 office.python.file 跑。要图表对象用原生图表
（Excel 的 sheet.chart / PPT 的 deck.chart），Python 画的 PNG 只当示意图。
音频与视频：先 office.av.check() 看 ffmpeg 与本机 SenseVoice 在不在，再用
office.av.extract（视频：逐字稿 + 抽帧）或 office.av.transcribe（只要文字）；
长逐字稿给 out 落成文件再 read，抽出来的帧逐张 read_image。
检索：一两个事实点走 office_search_run；要渠道覆盖走 brief → dispatch → parse_findings。
记忆：热记忆动笔前照办、台账由 office_run 自动记、归档按需检索；office_help 末尾会带
一份热记忆投影，所以平常不用特意 read；用户说明偏好或纠正你时当场记一条。
文件路径原样引用，方便用户打开。`;

/** 帮助索引。 */
async function indexText(detail = false) {
    if (!detail) return COMPACT_INDEX;
    const loaded = await loadAllFormats();
    const lines = loaded.map((item) => {
        if (item.module?.meta) {
            return `- ${item.entry.id}：${item.module.meta.name}（${item.module.meta.ext}）—— ${item.module.meta.summary}`;
        }
        return `- ${item.entry.id}：暂不可用（${item.error?.message ?? '加载失败'}）`;
    });
    return `办公模式能力索引

格式
${lines.join('\n')}
- pdf：PDF（.pdf，只读）—— 读 PDF：先 info 看有没有文本层，文本型抽文字、
  图像型（扫描 / 手写）渲染成图片再用 read_image 看

其他话题
- guide：说话与做事的方式（每次进入办公模式建议读一遍）
- run：批处理脚本怎么用
- python：Python 计算与绘图（科学计算、画图，图直接嵌进 Word / PPT）
- av：音频与视频的内容提取（录音 / 视频 → 带时间戳的文字；视频 → 抽帧给 read_image）
- files：文件读写与批量替换
- archive：zip / gz / zst 归档与会话导出的读法（info 看有什么 / list 列条目 /
  text 取一段 / find 按行检索 / extract 落盘；判型按魔数不看扩展名，
  office.files.read 读到压缩包会直接报错并指到这里）
- preview：把自家 docx / xlsx / pptx 渲染成页面图（交给 read_image 看版式）
  或转成 PDF（office.preview；引擎是宿主随部署安装的 LibreOffice，按需调用）
- convert：文档互转与统一出 PDF（office.convert）—— md / docx / tex 互相转；
  tex / md / Office 三件套与老格式 / OpenDocument 都能出 PDF（认识的格式见 check 的矩阵）；
  三条引擎（pandoc / TeX / LibreOffice）各自探测，缺哪样只影响它那几条路
- image：配图（office.image）—— 元信息（尺寸 / 哈希 / 可嵌入性）、免 Key 图库
  （维基共享资源 / Openverse）检索、网页选图（og:image / srcset）、抽 docx / pptx / xlsx
  内嵌图、取图（校验魔数、落 assets/）、一行署名 credit（嵌文档时传 source，
  配图必须能追到来源）
- writing：写作规划与自检（office.writing）—— 按每节难度与读者收益分配字数与深度
  （重点 / 机制 / 跃迁多写，背景收住），逐节给写法、语言与比喻边界；写完查句长、套话、
  锚点、让步与重点节是否写透。plan.markdown 可直接 office.log；不是结构模板，也不是检测器
- cache：中间产物与清理规则（跨调用保留、可命中复用）
- memory：三层办公记忆（热记忆 / 台账 / 归档）与知识库 kb 怎么读写、记什么不记什么，
  以及记忆范围、来源三档、召回质量、每回合配额、备份迁移与浏览面板
- theme：视觉主题（默认 plain 素色，无背景填充）
- search：联网检索（内置通道 / 三步走 / 渠道分流规则 / 结果解析）
- settings：设置界面（工具开关与参数调节）

查详情：office_help({ topic: 'word' }) 这样传话题名。`;
}

/**
 * office_help 能回答的话题清单 —— **按真源现取**，不手写。
 *
 * 为什么要有它：这份清单以前手写了两遍（「没有这个话题」的提示里一遍、`office_help`
 * 工具参数描述里一遍），两处都会漂。第四十七轮实测：提示里 `tex` 出现两次，
 * 而 `memory` / `search` / `settings` / `archive` / `preview` / `image` 一个都没列
 * —— 尽管它们都能查。现在两边都从这里取，新话题只要进了 `BASE_TOPICS` /
 * `COMPACT_TOPICS` / 格式注册表就自动出现。
 *
 * 顺序由 {@link HELP_TOPIC_ORDER} 定（给模型看的清单要好读），没登记的新话题追加在末尾，
 * 不会被静默漏掉。
 *
 * @returns {string[]} 话题 id
 */
export function availableHelpTopics() {
    const real = new Set([...Object.keys(BASE_TOPICS), ...Object.keys(COMPACT_TOPICS), ...formatIds(), 'guide']);
    const ordered = HELP_TOPIC_ORDER.filter((id) => real.has(id));
    const rest = [...real].filter((id) => !HELP_TOPIC_ORDER.includes(id)).sort();
    return [...ordered, ...rest];
}

/** 话题清单的展示顺序（不在这张表里的话题会追加到末尾）。 */
const HELP_TOPIC_ORDER = [
    'guide', 'run',
    'word', 'excel', 'ppt', 'tex', 'pdf', 'python', 'av',
    'preview', 'convert', 'image', 'writing', 'archive',
    'files', 'cache', 'theme',
    'memory', 'search', 'settings',
];

/**
 * 取的帮助文本。未知话题不抛错，直接把可用话题列回去，
 * 模型一次调用就能自我纠正。
 *
 * @param {string} topic 话题名（省略 = 索引）
 * @param {{detail?: boolean}} [options] `detail:true` 给全文层（例子、为什么、
 *   配置明细与迁移说明）；默认只给「怎么写才不错」的那一层 —— 文档会留在
 *   后续每个请求的前缀里，默认层薄一点、缓存命中率与 token 都受益。
 */
export async function buildHelp(topic, { detail = false } = {}) {
    const wanted = String(topic ?? '').trim().toLowerCase();
    if (wanted === '') return { topic: 'index', text: await indexText(detail) };
    if (wanted === 'guide') return { topic: 'guide', text: detail ? OFFICE_GUIDE : COMPACT_GUIDE };
    if (BASE_TOPICS[wanted] !== undefined) {
        const full = BASE_TOPICS[wanted];
        return { topic: wanted, text: detail ? full : (COMPACT_TOPICS[wanted] ?? full) };
    }

    // tex 的长文档由格式模块自己维护（meta.guide）：项目结构、内容块写法、
    // 编译参数与常见错误都在那一份里，文档不再抄一遍 —— 抄一遍就会漂移。
    // 默认层用同一模块的 meta.guideBrief（没有就退回全文）。
    if (wanted === 'tex') {
        const loaded = await loadFormat('tex');
        const meta = loaded.module?.meta;
        if (loaded.error === undefined && typeof meta?.guide === 'string') {
            const brief = typeof meta.guideBrief === 'string' ? meta.guideBrief : meta.guide;
            return { topic: 'tex', text: detail ? meta.guide : brief };
        }
    }

    const loaded = await loadFormat(wanted);
    if (loaded.error !== undefined) {
        return {
            topic: 'index',
            text: `没有「${wanted}」这个话题。可用：${availableHelpTopics().join(' / ')}。\n\n${await indexText(detail)}`,
        };
    }
    const meta = loaded.module.meta;
    const header = `${meta.name}（office.${meta.id}，${meta.ext}）\n\n${meta.summary}`;
    if (!detail) {
        // 格式话题的默认层：方法签名一览（模块自带的 brief，没有就退回 methods）
        // + 往哪儿要全文。brief 由格式模块自己维护，接口改了这里跟着改。
        return {
            topic: wanted,
            text: `${header}\n\n${methodSignatureLines(meta.brief ?? meta.methods)}\n\n`
                + `主题：默认 plain（素色、无底纹）；要配色 office_help({ topic:'theme' })。\n`
                + `例子、字段全表与注意事项：office_help({ topic:'${wanted}', detail:true })。`,
        };
    }
    const methodLines = Object.entries(meta.methods ?? {})
        .map(([group, list]) => [`【${group}】`, ...list.map((line) => `  ${line}`)].join('\n'))
        .join('\n\n');
    return {
        topic: wanted,
        text: `${header}

${methodLines}

主题当前可用：${themeCatalog().map((theme) => theme.id).join(' / ')}
约定：默认不传主题就是 plain（素色网格、无背景填充）；要配色再
office_help({topic:'theme'}) 挑一个，然后 office_run 一次写完。`,
    };
}

/** 格式话题默认层的方法一览：分组标题 + 签名行。 */
function methodSignatureLines(source) {
    const groups = Object.entries(source ?? {});
    if (groups.length === 0) return '（这个格式没有可调的方法）';
    return groups
        .map(([group, list]) => [`【${group}】`, ...list.map((line) => `  ${line}`)].join('\n'))
        .join('\n\n');
}
