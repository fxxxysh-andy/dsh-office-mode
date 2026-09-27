# dsh-office-mode（办公模式）

给 DeepSeek Harness 的办公插件：**七个工具声明**，一次调用干完一批活，
自带 Word / Excel / PPT 的生成、读取、编辑、排版与主题，零第三方依赖；
另有 LaTeX 学位论文（thuthesis 模板）的源文件生成与 PDF 编译；
带一条按内容类型分流的检索链路，**联网通道也自带**（宿主 web 服务优先、拿不到就自己发 HTTP，
所以 preset 里没有 `web_search` 也能查资料），结果落盘、主上下文只收摘要；
还带一套按工作目录存的**三层记忆**（热记忆 / 台账 / 归档，参考 mnemon 的分层做法）。

## 它解决什么

| 问题 | 做法 |
| --- | --- |
| 工具太多、每次都要挑 | 只注册 `office_help` / `office_run` / `office_memory` 三个常用工具 |
| 写法和细节占满上下文 | 用 `office_help` 按需取，不查不给 |
| 改一处文字要来回好几次 | `office_run` 里一次批量替换，只写一次盘 |
| 生成的文档"能用但难看" | 主题模板 + 排版测量，写盘前就报溢出风险 |
| 默认就被上了底色 | 不传主题 = `plain` 素色网格：Excel 是普通表格的网格形态，Word 是白纸黑字，零背景填充 |
| 不知道到底生成了什么 | 写出的文件自动重新打开解析，回真实统计与结构 |
| 满目录临时文件 | 中间产物只进 `.office/cache/`，跨调用保留、按 12 小时自动清理 |
| 用户纠正过一次，下次又犯 | 三层记忆：偏好与约定（热记忆）、交付物登记（台账）、下沉的旧条目（归档），跨会话留在 `.office/memory/` |
| 忘了"上次那份季度汇报"叫什么、用什么主题 | `office_run` 每写出一份文档就自动登记台账；`office_memory` 按关键词检索 |
| PDF 读不了（扫描件 / 手写笔记根本没有文本层） | `office.pdf`：先 `info` 判断类型，文本型抽文字，图像型渲染成页面图交给 `read_image` |
| 学位论文的版式要求严，手工排版过不了审查 | `office.tex`：按 thuthesis 模板生成整篇项目（主文件 + thusetup + data/ 分章 + refs.bib + 模板文件）并调用本机 TeX 编出 PDF —— 版式交给模板，源文件的结构、字段、引用对不对由插件保证 |
| 检索只用一个来源，结论被单一渠道钉死 | 按内容类型分流：热点走权威媒体 + 社交平台，知识走百科后下沉文献，手册只认官方文档；每类都先泛搜 |
| preset 里没有 `web_search`，检索整条不可用 | 插件自带联网通道：优先宿主的 `ctx.web`，拿不到就自己发 HTTP（DeepSeek 原生 web_search + 自带取正文）；`office_search_dispatch` 在子代理跑不了时自动改用它 |
| 搜索结果把上下文灌满 | 材料一律写进文件，主上下文只解析出几十行摘要（内置通道同样如此） |
| 同一件事被当成多个来源 | 跨源核对：按事实合簇，只有单一站点背书的照实标出来 |

## 安装

本仓库有两件东西：插件本身在 `packages/dsh-office-mode`，「办公模式」这份 preset 在
`presets/office`。下面两条命令都在**仓库根目录**下执行。

```powershell
# 1) 插件（能力）：装进当前 profile，所有模式都能用
#    用 plugin_manager 装该目录，或在仓库根目录下：
dsh plugin --profile web add ./packages/dsh-office-mode

# 2) 模式（说话方式 + 精简工具目录）：presets/office 是一个**声明式 preset bundle**
#    （patch 里一行 @deepseek-ai/dsh-agent-preset 声明），同样用 plugin_manager 装：
#      install_bundle → ./presets/office
#    注意：0.1.7 起 DSH **不再读** ~/.dsh/.agent-presets/ 下的遗留目录式预设
#    （preset.yml + agent.cordis.yml），所以不要再往那里复制 —— 那条路会让
#    「办公模式」从选择器里静默消失。

```

装完插件后，在任意会话里 `office_help` 就能用；选择「办公模式」启动的会话
才会吃到 persona 和精简工具目录。

办公模式与检索子代理两套组装都**关掉了 mnemon（记忆）**：整套 mnemon 组件挂在
一个 group 行 `mnemon-bundle` 下，只关组行不够——组内每个组件行仍会被 loader
逐条解析，所以组行与组内 9 个组件行都要逐个关。`node test/preset-check.mjs`
会拿 profile 的真实行清单核对这件事。

关掉 mnemon 不等于办公场景没有记忆：本插件自带一套**按工作目录**的三层记忆
（见下文「三层记忆」）。两者的分工是清楚的——mnemon 是跨项目的持久记忆体与
图关系检索，本插件管的是「这份稿子的约定」与「这个目录里产出过什么」，
按目录隔离反而更准（不同项目的约定经常互相矛盾）。

## 七个工具

**`office_help({ topic })`** —— 按需文档。`topic` 取 `word` / `excel` / `ppt` / `tex` / `pdf` /
`python` / `theme` / `files` / `cache` / `memory` / `run` / `search` / `settings` / `guide`；省略返回索引。
它的返回值**末尾会附一段记忆投影**（热记忆 + 最近台账），所以模型动笔前那一次
调用就能看到约定，不必额外读一次记忆。

**`office_memory({ action, ... })`** —— 三层记忆的读写与检索，见下文「三层记忆」。

**`office_search_run({ queries, ... })`** —— 内置检索直查：给 1–5 条查询，
来源与摘录写进 `.office/search/<slug>/run.md`，聊天里只回一份清单。见下文「检索」。

**`office_search_brief` / `office_search_dispatch` / `office_parse_findings`**
—— 按渠道的检索三步（出提纲 → 执行 → 解析回摘要），见下文「检索」一节。

**`office_run({ script, purpose?, keepCache? })`** —— 一次调用执行一整段脚本。

```js
// 一次调用：三件套 + 批量改 Markdown，全部做完
const w = office.word.create({ theme: 'business' });
w.title('2026 年第一季度业绩汇报')
 .heading('一、整体情况', 1)
 .para('本季度整体达成率 108%，超额完成目标。')
 .table({ columns: ['月份', '目标', '完成', '达成率'],
          rows: [['1 月', 1200, 1302, '108.5%'], ['2 月', 1100, 1144, '104%']] })
 .save('季度汇报.docx');

const x = office.excel.create({ theme: 'business' });
const s = x.sheet('明细');
s.title('季度明细').table({ columns: ['月份', '目标', '完成'], rows: [...], totalRow: true });
x.save('季度数据.xlsx');

const p = office.ppt.create({ theme: 'business' });
p.cover({ title: '季度汇报', subtitle: '2026 Q1' })
 .bullets({ title: '本季度三件事', items: ['达成率 108%', '新增客户 32 家', '交付周期缩短 4 天'] })
 .closing();
p.save('季度汇报.pptx');

office.files.edit('notes.md', [['旧结论', '新结论'], ['TODO', '已完成']]);
return { done: 3 };
```

执行环境里**没有** `process` / `require` / `fs` / 网络 / 定时器；碰文件只能走
`office.files`，所以每个产物都被记账。

## 音频与视频：把声音变成文字，把画面变成图

会议录音、访谈、手机语音条这类**只有声音**的输入，此前在办公模式里没有任何出口；
视频里**只有画面**的信息（板书、幻灯片、现场照片）也一样。`office.av` 补的就是它，
默认按「设备上本来就有 ffmpeg 与 SenseVoice」的口径工作：

```js
await office.av.check();                     // ffmpeg / ffprobe / 模型 / 运行时在不在

// 会议视频：一条调用拿到逐字稿 + 若干画面
const r = await office.av.extract('季度会.mp4', { language: 'zh', out: '逐字稿.md', count: 6 });
// r.transcript.segments: [{ start, end, clock, text, lang, emotion, event }]
// r.frames.files:        [{ at, clock, path }]  → 逐张 read_image 看

await office.av.transcribe('采访.m4a', { out: '采访.txt' });   // 只要文字
await office.av.frames('课程.mp4', { every: 60, width: 1280 }); // 只要画面
```

- **声音 → 文字**：本机 SenseVoice 离线识别（语音输入下载的那份 ONNX 模型），
  默认 int8；解码交给 ffmpeg，按块（默认 120 秒）跑 Silero VAD 切句，逐句给绝对
  时间戳，语种 `auto` / `zh` / `yue` / `en` / `ja` / `ko`。识别结果里还带
  `lang` / `emotion` / `event`（`<|zh|>` / `<|NEUTRAL|>` / `<|Speech|>` 这类富文本
  标签）—— 这是上游那条语音输入链丢掉的部分。
- **画面 → 图片**：`office.av.frames` 按 `at`（明确秒）> `every`（定步长）> `count`
  （默认 6 张均匀铺开）抽帧成 JPEG / PNG，路径交给 `read_image` 看。图上文字是
  读图能力认出来的，不是 OCR。
- **四样能力缺一不可**：ffmpeg、ffprobe、SenseVoice 模型、sherpa-onnx 运行时。
  缺任何一样都在调用处**明确报错**并说清怎么补（`office.av.check()` 一次报全），
  不静默降级、不假装转写出了一段空文字。配了路径就以配置为准（不再退回 PATH）。
- **规范 WAV 的坑**：SenseVoice 只吃 16 kHz 单声道 PCM16（`fmt` 块 16 字节、
  `data` 落在偏移 36）。SAPI 的默认输出（fmt 18）与 ffmpeg 的默认封装（带 LIST 块）
  都会被拒 —— 插件用 `-map_metadata -1 -fflags +bitexact` 解码后**再自己验一遍**，
  不合格就地规范化（`inspectWav` / `canonicalizeWav`）。
- **长音频**：单次上限默认 3600 秒（`av.maxSeconds`），单块 120 秒
  （`av.chunkSeconds`）。分块解决的是单块大小，不改变时长上限；超了就明确报错。
- **只识别，不越界**：不翻译、不认说话人、不做流式与降噪；分块之间没有跨块上下文，
  边界上的句子可能被切成两句。

## Python：算与画（办公模式里唯一的编程出口）

办公模式的工具面是刻意做小的：命令行、后台任务、子代理这些编程入口一律关掉。
代价是**科学计算与绘图没有出口** —— 一张要嵌进报告的折线图，只能手工画形状，
或者让用户自己算完再放进来。`office.python` 补的就是这一小块：

```js
const info = await office.python.check();   // 本机有没有 Python、装了哪些包
const py = await office.python.run(`
import numpy as np
import matplotlib.pyplot as plt
use_cjk_font()                              // 中文标题不变成方框
months = np.arange(1, 7)
revenue = np.array([182.4, 203.1, 197.8, 241.5, 268.2, 255.9])
fig, ax = plt.subplots(figsize=(7.2, 4.0), dpi=160)
ax.plot(months, revenue, marker='o')
ax.set_title('上半年营业收入（万元）')
fig.savefig(OUT_DIR + '/revenue.png', bbox_inches='tight')   // 图写进 OUT_DIR
print('均值 %.2f 万元' % revenue.mean())
`, { name: 'revenue-chart' });

const chart = py.files.find((file) => file.kind === 'image');   // 产物路径直接可用
office.ppt.create({ theme: 'business' })
    .image({ path: chart.path, title: '上半年营业收入' })
    .save('收入回顾.pptx');
```

- **只跑 Python**：解释器由插件探测（`python` / `python3` / `py`），脚本由插件落盘后
  执行；脚本里给不了别的命令，也开不了 shell。
- **产物只落缓存目录**：默认 `.office/cache/python/out`，按修改时间对比收出「这一次
  跑出来的文件」（上一次的图不会被误报成这一次的产物），并记进 `office_run` 的产物统计。
- **返回值**：`{ ok, code, stdout, stderr, files, script, outDir, logFiles, elapsedMs }`。
  脚本自己报错时 `ok:false`、`stderr` 里是完整回溯，**行号就是代码里的行号**
  （用 `runpy` 执行，不在代码前面拼前导）。
- **三个注入名**：`OUT_DIR`（产物目录绝对路径）、`WORK_DIR`（工作目录绝对路径）、
  `use_cjk_font()`（只在被调用时才 import matplotlib，纯计算不付这个代价）。
- **UTF-8 与无显示器**：子进程带 `PYTHONUTF8=1` / `PYTHONIOENCODING=utf-8`
  （中文 Windows 默认 GBK，重定向后的 stdout 会当场 `UnicodeEncodeError`），
  `MPLBACKEND=Agg` 让 matplotlib 在没有显示器的环境里也能出图。
- **长脚本**：先 `office.files.write('analysis.py', code)`，再
  `office.python.file('analysis.py', { args: ['2026'] })`（`args` 进 `sys.argv`）。
- **没有的东西**：`pip install`（不联网装包）、stdin（`input()` 立刻 `EOFError`）、
  命令行。缺包请让用户自己装。

## 读 PDF：先问一句，再决定抽文字还是渲染成图

PDF 分两类，读法完全不同 —— 2026-09-22 的实测就是在这上面卡住的：一份 ONYX
电子纸导出的手写笔记（13 MB / 5 页）里 `fonts=0`、`images=10`，**一个字都没有**，
而当时的办公模式只会「读文件」，于是探了 20 步 API 也没找到出口。

现在 `office.pdf` 把两条路都摆好：

| 调用 | 做什么 |
| --- | --- |
| `office.pdf.info(path)` | 页数 / 页面尺寸 / 标题 / 生产者 / 是否加密 / **有没有文本层** / 字体与图像计数 |
| `office.pdf.text(path, { from?, to?, out? })` | 抽文字（文本型）；`out` 给出时写进文件，避免长文本挤占上下文 |
| `office.pdf.pages(path, { from?, to?, dpi?, format? })` | 逐页渲染成 PNG/JPEG，返回可直接交给 `read_image` 的路径 |
| `office.pdf.engines()` | 本机探测到的渲染 / 抽文本 / 页数引擎 |

```js
const info = await office.pdf.info('线性代数/9-22线性代数笔记.pdf');
// { pages: 5, hasTextLayer: false, producer: 'NeoPdf', counts: { fonts: 0, images: 10 }, … }

const text = await office.pdf.text('讲义.pdf', { out: 'raw.md' });   // 文本型
const r = await office.pdf.pages('线性代数/9-22线性代数笔记.pdf', { from: 1, to: 2, dpi: 150 });
// r.files = [{ page: 1, path: '.office/cache/pdf/…-p1.png', bytes: …, reused: false }, …]
// 然后对每个 path 调 read_image —— 手写、扫描件、图表都靠这一步进上下文。
```

**引擎不打包进插件**（保持零依赖），按机器上现有的探测：
`fitz`（PyMuPDF，实测最快：5 页手写 1.3 s）→ `pdftoppm` / `pdftocairo`（poppler）→
`mutool` → `gswin64c`（Ghostscript）。一个都没有时 `pages()` 会明确报错并说清装哪个，
不假装支持。抽文本用 `pdftotext` 或 `fitz`；页数与元信息优先 `pdfinfo`，都探测不到时
退回插件自己的轻量解析（读 `%PDF-` 头、`/Count`、`/Info` 里的字符串，连中文 UTF-16BE
标题都能解出来），所以 `info()` 在任何机器上都可用。

**渲染结果按内容键缓存**：`sha1(文件内容) + dpi + 格式 + 引擎` 决定目录，同样的页码再渲染
直接命中已有文件（`reused`），不重算。这也是下面那条缓存改动存在的原因。

## PPT 图片：插进去，也抽出来

演示文稿里最常见的两种「图片需求」是反的：一种是**把图插进已有稿子**（改版、补素材），
另一种是**把稿子里的图抽出来**（复用、核对、交给 `read_image` 看）。两条路现在都有：

| 调用 | 做什么 |
| --- | --- |
| `office.ppt.images(path, { out? })` | 抽出包内嵌图。`out` 省略时写进 `.office/cache/ppt-images/<词干>/`，返回可直接交给 `read_image` 的路径 |
| `office.ppt.insertImages(path, items)` | 把图片插进**已有的**演示文稿（不是重新生成整篇） |
| `revise(path, [{ slide, addImage: {...} }])` | 与改文字同一批做：一次调用里既改标题又插图 |
| `read(path).stats.media` | 读路径上先看清包里有哪些图（部件名 / 内容类型 / 字节 / 像素 / 在哪几页） |
| `readSlides(path).pages[].shapes[].image` | 每个 `p:pic` 形状带上 `{relId, part, contentType, bytes, width, height, crop}` |

```js
// 抽出稿子里的图，交给 read_image 看
const r = office.ppt.images('季度汇报.pptx');
// r.files = [{ part: 'ppt/media/image1.png', path: '.office/cache/ppt-images/季度汇报/image1.png',
//              bytes: 18254, width: 800, height: 450, slides: [2], reused: false }, …]

// 往第 3 页插一张图：contain = 等比放进 8×5cm 的框并居中
office.ppt.insertImages('季度汇报.pptx', [
  { slide: 3, path: '素材/渠道分布.png', x: 2, y: 6, wCm: 8, hCm: 5, fit: 'contain', alt: '渠道分布' },
]);

// 同一批里既改文字又插图
office.ppt.revise('季度汇报.pptx', [
  { slide: 1, shape: 'Subtitle', setText: '2026 Q1 · 定稿' },
  { slide: 3, addImage: { path: '素材/月度趋势.png', x: 18, y: 6, wCm: 10, fit: 'natural' } },
]);
```

**插入的语义**（`items[]` 的字段）：

| 字段 | 说明 |
| --- | --- |
| `slide` | 页码（1 起）；省略 = 全篇 —— 注意那会把同一张图插到每一页 |
| `path` | 图片路径，只认 **PNG / JPEG / GIF**（与生成端同一口径，别的格式进 `skipped` 并说明） |
| `x` / `y` | 左上角坐标，单位**厘米**，与 `deck.shape` / `revise.move` 同一套坐标 |
| `wCm` / `hCm` | 目标框尺寸；只给一个时按原图纵横比推另一个；都不给 = 原图尺寸 |
| `fit` | `contain`（等比放进框并居中，默认）/ `cover`（铺满框、超出部分用 `a:srcRect` 裁掉）/ `natural`（按 96 DPI 的原始像素，不缩放） |
| `name` / `alt` | 形状名与替代文字（写进 `p:cNvPr/@name`、`@descr`） |

插入是**改包而不是重建包**：读回原 pptx，只在命中处动手 —— 追加一条 image 关系、
必要时补 `<Default Extension="png" …/>`、把 `ppt/media/imageN.<ext>` 从已有编号往后顺延、
把 `p:pic` 追加为该页 `p:spTree` 的最后一个形状。因此母版、备注、超链接关系、
其它页的既有内容都不会被丢掉（这条有字节级回归断言与真实 PowerPoint 校验兜着）。

**已知的两条**：外链图（`a:blip` 上是 `r:link` 而不是 `r:embed`）只报进 `skipped`，
不下载也不编造字节；`fit:'natural'` 不缩放，图比版心大就只给 warning，不拦。

## LaTeX 学位论文：生成源文件，再编出 PDF

`office.tex` 管的是清华学位论文（thuthesis 模板）。它和另外三种格式的区别在于：
产物不是一个文件，而是**一整个能直接编译的项目**。

| 调用 | 做什么 |
| --- | --- |
| `office.tex.create({...})` | 一次写出主文件 + `thusetup.tex` + `data/` 分章 + `ref/refs.bib` + 模板运行期文件 |
| `office.tex.read(path)` | 顺 `\input` 解析整个项目：章节结构、图表公式计数、引用与文献对不对得上、图片与 `\input` 指向的文件在不在 |
| `office.tex.edit(path, ops)` | 跨文件替换、改 `\thusetup` 字段、追加章节、加文献条目，一次调用全部落地 |
| `await office.tex.compile(path)` | 调本机 TeX 发行版编出 PDF：`{ok, pdf:{path,bytes,pages}, log:{errors,tail,path}, diagnosis}` |
| `await office.tex.engines()` | 本机探测到的编译入口（latexmk / xelatex / lualatex）与 bibtex / biber |
| `office.tex.escape(text)` | 把普通文本里的 `% & _ # $` 转义成 LaTeX 能编译的形式 |
| `await office.tex.clearAux(path)` | 清中间产物（`.aux` / `.log` / `.synctex.gz`），保留 PDF 与源文件 |

```js
// 一次调用：写出整篇论文并编出 PDF
office.tex.create({
  path: 'thesis/thesis.tex',         // 文件名用 ASCII；中文标题写 title
  degree: 'master',                  // doctor | master | bachelor | postdoc
  title: '基于××的××研究', titleEn: 'Research on ...',
  author: '张三', supervisor: '李四, 教授',
  department: '计算机科学与技术系', discipline: '计算机科学与技术',
  degreeCategory: '工学硕士', degreeCategoryEn: 'Master of Science',
  abstract: { zh: ['摘要…'], en: ['Abstract…'], keywordsZh: ['关键词1', '关键词2'], keywordsEn: ['kw1', 'kw2'] },
  chapters: [
    { title: '绪论', sections: [
      { title: '研究背景', blocks: ['正文…', { items: ['要点一', '要点二'] }, { equation: 'E = mc^2' }] },
    ] },
  ],
  references: [{ key: 'zhang2024', type: 'article', author: '张三 and 李四', title: '题名', journal: '某学报', year: '2024', pages: '1--7' }],
  acknowledgements: ['感谢…'],
});
const r = await office.tex.compile('thesis/thesis.tex');
return r.ok ? r.pdf : r.log.errors;
```

**排版不是插件做的。** 字体、行距、页边距、题注格式、参考文献著录格式都由
`thuthesis.cls` 决定：插件保证「结构、字段、文件、引用」都对，让模板去排版。
手工拼版式（`\vspace`、`\fontsize`）反而过不了格式审查。内容块文本按 LaTeX
源码处理，普通文本先过一遍 `office.tex.escape()`。

**模板从哪来**：按「配置 `texTemplateDir` → 环境变量 `DSH_OFFICE_TEX_TEMPLATE` →
工作目录里的 `thuthesis*` 目录 → TeX Live 自带的版本」依次找。找到成套的就整套
拷进项目（项目自包含，换机器、换 TeX 版本都能编）；只有 TeX Live 自带版本时就用
它编译，并在反馈里说明版本可能不同。

**编译**优先走 `latexmk`（它自己会算 bibtex 与重复编译次数），没有 latexmk 时退化
成 `xelatex ×2 → bibtex → xelatex`。日志里 `-file-line-error` 格式的真实错误会被
解析成 `{file, line, message}`；Overfull / Underfull 只算排版提示，不当错误。
命令行全程不用管道（受限沙箱下 `stdio:'pipe'` 会 `spawn EPERM`），stdout / stderr
重定向到普通文件句柄 —— 与 `pdf.js` 同一套跑法。

**没有 TeX 发行版也能用**：`create` / `read` / `edit` 照常（源文件本身就是交付物），
只有 `compile` 会明确报错并说清装什么。

写完的项目里的 `.tex` 与另外三件套一样会被自动复检一遍：`read()` 会报出
「图片不存在 / `\input` 指向的文件不存在 / `\cite` 的 key 在 refs.bib 里找不到 /
`\ref` 的 label 不存在 / 图缺题注 / 段落太长 / 缺摘要或致谢」这类会卡住编译或
不符合学位论文要求的问题。附录的 `\chapter` 不占正文章号，报告里单独计数。

## 三层记忆：热记忆 / 台账 / 归档

记忆存在**会话工作目录**下的 `.office/memory/`。三层不是三种存储，而是三种
生命周期——什么时候被读到、什么时候被写、装满了往哪儿去。

| 层 | 装什么 | 谁写 | 怎么读 | 装满了 |
| --- | --- | --- | --- | --- |
| 热记忆 | 用户偏好、项目约定（小、常驻） | `office_memory` 的 `add` / `replace` / `remove` | `office_help` 与 `office_run` 的反馈里自动带投影；也可 `layer:'hot'` | 按「重要度低、更旧」下沉到归档，并在 `MEMORY.md` 留一条指路条目 |
| 台账 | 每份交付物的登记：路径、格式、主题、复检统计、结构摘要、目的 | `office_run` 写盘后**自动**登记（也可 `action:'log'` 手工补） | `office_memory({ action:'read', layer:'ledger', query })` | 超过条数上限时，最旧的一批滚成月度归档摘要 |
| 归档 | 下沉的旧热记忆 + 滚动的旧台账（只读） | 由上面两层触发 | `office_memory({ action:'read', layer:'archive', query })` | **摘要文件**数有上限（一个月一个文件），最旧的摘要连文件一起删（唯一真正会丢东西的地方） |

**容量口径要注意两处别配对错了**（面板的容量条按这个来）：

- **归档的上限 `archiveKeep` 管的是「摘要文件数」**（一个月一个），不是条目数。
  一个月里沉 30 条仍然只占 1 个文件，所以是 `1 / 60 个` 而不是 `30 / 60 个`。
  条目数另算：页签「归档（30）」与指标卡报的是它。
- **容量条用的是未过滤、未截断的总数**：搜过之后列表里只剩 3 条，容量条仍然报
  `25 / 500` —— 它回答的是「离上限还有多远」，不该跟着搜索框缩水。
  快照里这两组数分开给：`counts.hot/ledger/archive/...`（列表口径）与
  `counts.ledgerTotal/archiveFiles/archiveItems`（容量口径）。

**投影只在热记忆变化时贴全文**：`office_help` 与 `office_run` 的反馈尾巴上挂着
热记忆投影，而这两个工具在同一个会话里会被反复调用。热记忆没变时只给一行
（条数 + `revision` + 怎么读全文），台账最近几条照旧 —— 本机实测
**11714 字节 → 1442 字节（8.1x）**，一轮里调 10 次办公工具从 117 KB 降到 24 KB。
三条边界写在 `src/projection.js` 顶部：① 每 12 次强制重贴一次全文（防宿主的
工具结果压缩把早先那份裁掉；这是**有界的近似**，残余风险已在文件里写明）；
② 台账部分每次都贴（它不体现在热记忆的 `revision` 里，省掉会看不到刚登记的那条）；
③ 拿不到会话身份时不假装「刚贴过」，每次都贴全文。

目录结构：

```
.office/memory/
  memory.json       热记忆的真源（USER.md / MEMORY.md 都是它的投影）
  USER.md           投影：用户偏好（默认上限 4096 字节）
  MEMORY.md         投影：项目与环境（默认上限 10240 字节）
  ledger.jsonl      台账，一行一条（默认最多 500 条）
  links.jsonl       图关系，一行一条（双向：从任一端都走得到对面）
  archive/YYYY-MM.md  月度归档摘要（可打开、可 grep）
  archive/index.json  归档索引（检索用；与上面的文件同一次写入更新）
```

开了「两者并存」范围后，全局层是同样一套文件，默认在 `$DSH_HOME/.office/memory`。

三条与 mnemon 一致的原则：

1. **Markdown 是投影，不是存储。** 直接编辑 `USER.md` / `MEMORY.md` 会在下次
   写入时被覆盖 —— 改记忆只能走 `office_memory`。
2. **容量满了向下沉，不静默丢。** 下沉的条目进归档且可检索，`MEMORY.md` 里留
   指路条目；连指路都放不下的极端容量下，才只保内容。单条本身就超限则直接报错，
   不截断内容。
3. **读是有界的。** 一次 `read` 有「条数 / 单条字符 / 总字符」三重上限，`layer:'all'`
   时三层共用一份预算（和 mnemon 的「一个证据信封」同一个目的：别把排版与措辞的
   余地从上下文里挤走）。被截断时明确带 `truncated`。
4. **投影是有界的，而且只在变化时重复。** 挂在每次办公工具调用上的那段投影，
   热记忆没变时只给一行（见上一节的实测数字）—— 常驻不等于每次都复制一份。

写入协议（与 mnemon 的 `add` / `replace` / `remove` 同形）：

- `add` 用于**新的独立事实**；`replace` 需要 `oldText` 能**唯一**命中（命中多条会
  被拒绝——宁可让你补长，也不要改错）；`remove` 只在用户明确要求或确有证据时用。
- `target`：`user` = 用户是谁、偏好与要求；`project` = 项目约定、环境事实、工具坑。
- 只记稳定的事实。检索结果原文、一次性进度、助手自己的推测都不记。

### 记忆范围与层开关

- `scope`：`workspace`（默认，一个项目一份）/ `global`（所有项目共用一份）/
  `both`（两者并存）。读的时候**全局在前、工作区在后**——具体项目的约定比通用
  偏好更该被最后看到。每条结果都带 `origin`（全局 / 工作区）。
- `globalDir`：全局层目录，默认 `$DSH_HOME/.office/memory`。相对路径相对**用户主
  目录**，不相对工作目录：全局层跟着工作目录走就没有意义了。
- `userScope`：`memory`（默认，跟着上面的范围）/ `global`（「用户偏好」始终落全局
  层，换工作目录不必把「用户是谁」重记一遍）。
- `layers.hot` / `.ledger` / `.archive`：每层独立开关。关掉是**不再读、不再写**，
  不是只读；已有数据不会被删除，重新打开就回来。

### 图关系与实体

- `link` 给两条记忆建一条**双向**类型化关系。`kind`：`related`（默认）/ `refines` /
  `supersedes` / `contradicts` / `supports` / `derives`。
- 两个 id 都必须**真实存在**，否则报错。允许悬空边会让关系图慢慢烂掉，而悬空边在
  遍历时表现为「静默少一条」——比直接报错难查得多。
- 热记忆下沉进归档时**保留条目 id**，所以关系不会因为下沉而断（这条是第一版漏掉
  的：id 一丢，「下沉」等于「关系断掉」）。
- `related` 默认只走**一跳**（`depth` 可到 3）：图一旦放开很容易把整个记忆库拉进
  上下文。
- `entities` 的实体来自条目**显式声明**的 `entities` 数组，不做抽取。猜出来的实体
  比没有实体更误导。

### 召回质量与每回合配额

两件不同的事：**召回质量**管「一次查询返回什么」，**配额**管「一个回合能查几次」。

- 召回质量 `strict-v1` 按「命中了查询里几成的词」算相关度分档——这是**命中占比**，
  不是向量余弦相似度（本插件没有向量）。达到高阈值直接采纳，中档与未知档各有名额
  上限，低于下限的丢掉，丢了多少条会报出来。
- 分档前先取「条数上限 × `candidateMultiplier`」个候选：直接截断的话，低分结果会
  把高分结果挤出候选池。
- 配额：一个回合里第一次带 `query` 的检索、后续换词细化、图关系遍历各有上限
  （默认 1 / 1 / 1）。回合号从会话日志的 `turn/start` 事件读出来。用完会**明确
  拒绝并说明原因**。不带 `query` 的 `read` 是看现状、不是查资料，不占配额。
- 两者都能在设置页关掉（`policy: off`、配额设 0）。

### 备份与迁移（Pack）

`office_memory({ action:'export', packPath? })` 把热记忆 + 台账 + 归档（索引与
Markdown 全文）+ 关系整包写成一个 JSON 单文件；`action:'import'` 读回来。
口径是**只增不改、按 id 幂等**：同 id 已存在就跳过，不覆盖现内容；归档按月合并、
按块去重。所以重复导入同一份包不会产生第二份。这份 Pack 含私有记忆，不要当公开
文件传。

不打包成 zip 是有意的：本插件零依赖，读回来这条路上任何一个压缩库都会变成新的依赖。

### 浏览面板

设置对话框左侧「记忆系统」是**配置**页；侧栏的「记忆」入口是**浏览**页：热记忆 /
台账 / 归档 / 关系 / 实体五个页签，带搜索与工作目录切换。数据来自本插件注册的只读
端点 `GET /office-memory/snapshot`。

面板不只是列表 —— 每一层都尽量给出**形状**，因为「记忆里有什么」用数字与图形比用
文字更容易一眼看清（做法参考 mnemon 的图谱页，但零依赖、不引入任何图表库）：

| 位置 | 可视化 |
| --- | --- |
| 标题下 | 五张指标卡：热记忆 / 台账 / 归档 / 关系 / 实体 的条数、层色点与说明（数字用等宽字） |
| 容量 | 台账条数对 `ledgerLimit`、归档摘要对 `archiveKeep`、热记忆与项目记忆的**字节占用**对各自上限；4px 轨道 + 语义色 + 百分比 |
| 存储域 | 每个根一张卡片：目录、四层条数（带层色点）、可读字节数、文件数 |
| 关系 / 实体 | **手写 SVG 力导向关系图** + 右侧 `GRAPH INSPECTOR`：节点 = 三层条目 + 实体（按层区分圆 / 方 / 菱），边 = `links[]`（按关系类型上色）+ 实体引用；悬停高亮邻居、点击选中（虚线环）、方向键微调、滚轮缩放、拖背景平移、拖节点挪位置、「自然铺开 / 均匀重置」两个布局动作、图例页脚与详情栏 |
| 实体 | 频次条（宽度 = `count / maxCount`）+ 来源层 chip |
| 台账 / 归档 | 时间线：台账按天、归档按月，柱高按峰值折算 |
| 热记忆 | 按 `target` 分组（用户偏好 / 项目与环境），左侧 3px 色带按 `importance` 上色 |
| 列表 | 每页 20 条，底部给「再显示 N 条」与「已显示 20 / 25，还有 5 条」；全显示后换「收起」 |
| 条目完整内容 | 卡片上的「展开 / 收起」（键盘可达）+ 面板底部的 `ITEM DETAIL` 详情条；**不用原生 `title`**（系统气泡不可样式化、键盘触发不了、位置也由浏览器决定） |

**面板怎么「写」**（第十三轮）：条目上有 `改写 / 忘记 / 改用途 / 删关系 / 找相关`
这些动作，标题栏有「记一条」。点一下**不会落盘** —— 面板把一条精确的
`office_memory({...})` 调用生成出来，放进顶部的待办区，用户看过之后「复制指令」
（或「切到会话」）再在会话里发送，真正的写入由 `office_memory` 完成。
这样做的三条理由：

1. 记忆一旦写进去就跨会话影响后面所有回合，该由用户先看见「要记什么」再点头 ——
   与回复旁的「存入记忆」同一个口径；
2. 浏览器半侧**没有任何写入通道**：打包的客户端入口只拿到 `require`，既没有
   `host.call`（那是动态 cordis 包的），客户端服务目录里也没有 composer /
   `inputActions`（只有 layout / locale / sessions / slots / theme / timer /
   uiWorkspace / workspaces），`main` slot 的 standardProps 同样不含 `inputActions`；
3. 唯一的现成通道是只读端点，放宽它的信任边界等于把本地 HTTP 端点变成可写面，不做。

归档层刻意**不给**写入动作（它由热记忆下沉与台账滚动产生），面板上给一句说明而不是
一个点了会报错的按钮。

**视觉语言对齐 mnemon**（第十二轮）：令牌先收敛成一层（表面 `bg-layer-1` / 下沉底
`bg-module-platform` / 文字四级 `label-primary…caption` / 描边 `border-l1`·`l2` /
四个状态色 / `elevation-soft`），面板样式只引用它 —— 原来满屏 `opacity: 0.6` 的
层级写法全部撤掉。几个从 mnemon 抄来的签名手法：卡片侧色带用 `inset 3px 0 0`
（不改盒模型、不被圆角裁切）、同色 5% 渐变卡片底、状态点光晕、画布中心 6% 径向光晕、
4px 容量条 + 填充 `border-radius: inherit`、空态用虚线框 + 圆环 glyph。

需要伪类 / 关键帧 / 容器查询的地方（悬停、`:focus-visible`、加载转圈、
`prefers-reduced-motion`、窄面板塌列）走面板树里的**一个 `<style>` 节点**：
它是 React 树里的普通节点而不是模块顶层副作用（测试与 SSR 环境没有 `document`），
内容由常量拼成、没有插值，类名一律 `om-` 前缀。其余样式仍然是内联对象，
主题令牌用宿主真实存在的 `--dsw-alias-*`（另一套 `dsh` 前缀的变量在 DSH 里
一个都没定义，写了等于永远走兜底值）。空快照与出错时照样渲染骨架，不会崩、
不会出现 `NaN`。

> 一条只有实测能发现的结论：**浅色主题下 `bg-layer-1/2/3` 全是 `#fff`**，
> 拿 `bg-layer-2` 当「嵌套表面」等于白叠白；嵌套底要用 `bg-module-platform`
> （浅色 `#f5f6f7` / 深色 `#353638`）。探针 `test/web-verify.mjs` 会把面板用到的
> 令牌逐个读回运行时值（`findings.dom.tokens`），换令牌时先看那张表。

**滚动：面板根是唯一的滚动容器**（第二十轮修）。宿主主面板那一格是
`.pI_x6G_centerCol{display:flex;flex-direction:column;overflow:hidden}` —— 它**不给**
`main` 槽位任何滚动能力，而槽位锚点的 style 是 `display:contents`（不产生盒子），
所以面板根元素就是那个 flex 项。第二十轮之前面板没有高度也没有 `overflow`：内容比视口
高时被中心列裁掉，只有列表内部那 520px 能滚，指标卡 / 容量条 / `ITEM DETAIL` / 页脚
**全都在首屏之外够不着**。现在：

- 面板根 `height:100%` + `min-height:0` + `overflow-y:auto` + `overscroll-behavior:contain`；
- **列表不再自己滚**（去掉内层 520px），消除双滚动条；「再显示 N 条」滚的是面板根；
- 头部 `position:sticky` 吸顶（底色取 `--dsw-alias-bg-base`），搜索框与「刷新 / 记一条」
  不会滑出视野；
- **图谱缩放要按住 Ctrl / ⌘ / Alt + 滚轮**（或视图条 ±）：原先普通滚轮一律
  `preventDefault` + 缩放，画布占了列表上面一大块，鼠标停在图上就滚不动页面 ——
  普通滚轮现在留给滚动。

真实浏览器实测（`test/web-verify.mjs`）：`clientHeight 802` 正好等于中心列高度、
`scrollHeight 1257`、滚到底后页脚可见、吸顶头贴顶差 4px；普通滚轮面板 `scrollTop`
从 460 到 397 而画布 zoom 仍 100，Ctrl + 滚轮 zoom 100 → 115。

该端点的信任边界（`src/view.js` 顶部有完整说明）：

- **只读**：只接受 `GET` / `HEAD`，其余方法一律 405，没有任何写入路径。
  「面板上能改记忆」不等于「端点能写」：面板只生成 `office_memory` 指令文本，
  写入仍走会话里的工具（见上一节）。第十三轮刻意**没有**加写端点。
- **只服务见过的根**：`?cwd=` 必须命中「本进程真的跑过办公工具的工作目录」，否则
  404。没有这一条，一个本地 HTTP 端点就成了「传任意路径读任意目录」的数据外泄面。
- 组合里没有 `webServer`（例如 headless）时端点不注册，面板显示可读提示。
- 快照里 `stores[]` 带 `bytes` / `files`、`counts` 带 `hotBytes` / `projectBytes`，
  容量条与存储域卡片靠它们；字段缺失时那两块静默不画，不影响其余部分。

### 与 mnemon 的有意差别

- **没有远端记忆体、没有多 Provider 后端**：全部落在工作目录与全局层的普通文件里。
- **没有空闲审查**（后台自动复盘）：`autoCapture` 是一句**指引**，不是自主记录器。
- **实体不抽取**、**关系不自动生成**：都要求显式声明，理由见上。

为什么投影挂在工具结果上而不是提示段落上：办公模式的 persona 是 `complete: true`
的（唯一的系统提示），插件注册的提示段落会在组装时被丢掉；即便在别的模式里能注册，
section 的 `text` 也只接受同步函数，而记忆要读盘。所以「动笔前看到记忆」固定在
`office_help` 的返回值与 `office_run` 的反馈里。

在**别的模式**里（persona 不是 complete 的）另有一段静态说明 `promptHint`：只讲
记忆在哪、怎么用，不含任何记忆内容。它是 mnemon 关掉之后顶替「每轮热记忆提醒」的
那一段；办公模式下会被 persona 丢掉，所以不会重复。

### 从 mnemon 迁移

mnemon（`dsh-mnemon` bundle）在本 profile 里已经关闭；它的三层与本插件的三层不是
一一对应，所以迁移是**降级映射**：

| mnemon | → 办公记忆 | 说明 |
| --- | --- | --- |
| runtime `USER.md`（全局 `~/.mnemon`） | 热记忆 `user` | 全局那份会被**复制**进本工作目录：换项目要各迁一次 |
| runtime `MEMORY.md`（工作目录 `.mnemon`） | 热记忆 `project` | 逐条 add，保留重要度 |
| Memory Spaces 的 insights | 归档 | 内容、类别、重要度、标签、实体与「关系边数量」都保留；图跳转降级为计数；已软删的不迁 |
| Documents（项目档案） | 没有对应层 | 只报告数量，原件留在原地 |

两条入口（幂等，只增不改，重复跑不会产生第二份）：

```powershell
# 命令行（改完源码不必重启就能迁移；也可 --dir 指定别的工作目录）
node scripts/migrate-mnemon.mjs --dry-run      # 先看会迁什么
node scripts/migrate-mnemon.mjs

# 会话里（插件已加载时）
#   office_memory({ action: 'migrate', dryRun: true })
#   office_memory({ action: 'migrate' })
```

读长期记忆库用 Node 自带的 `node:sqlite`（Node 22.5+）；拿不到时只迁 runtime 部分
并明确报警，不会整次失败。`.mnemon/` 原数据**不删** —— 回溯、核对与重迁都靠它。

## 检索：插件自带联网通道

六件套里的后四个工具是一条检索链路（外加一个直查入口）。它存在的理由是两件事：
**别让结论被单一来源钉死**，以及**别让搜索结果把主上下文灌满**。

而这条链路要能跑起来，先得有个「能查」的通道 —— 这就是**多路联网通道**
（`src/web.js` 编排 + `src/web-providers.js` 十条通道 + `src/web-preprocess.js` 预处理）。

### 为什么需要自带通道

「办公模式」是 agent preset，preset 的 `plugins` 列表就是那个会话的**全部**插件行，
里面没有 `@deepseek-ai/dsh-tool-web`。所以办公会话的工具面里根本没有
`web_search` / `advanced_search` / `platform_search` / `web_fetch`：
第二步派工时给子代理圈的那七个工具名一个都不存在，子代理一启动就报

```
tools.restrict() names unknown global tools "web_search", "advanced_search", …
```

整条链路 0/N 全失败（2026-09-25 的真实会话就是这样：检索提纲要得出来，
派工 7/7 全红）。所以本插件自己带了一条联网通道，不依赖别的插件。

### 十条通道（第二十轮起不再只有 DeepSeek 一条路）

| id | 通道 | 要 Key | 备注 |
| --- | --- | --- | --- |
| `seam` | 宿主 `ctx.web` | 不要 | 首选：宿主自带公网地址校验 + **地址钉死在连接上**、同源跳转、体积与超时上限、代理路由 |
| `anthropic` | Anthropic 兼容 + 原生 `web_search` | 要 | 默认 DeepSeek 官方（`deepseek-v4-flash`，Key 取 `DEEPSEEK_API_KEY`）；换 `baseURL` 可指别的兼容端点 |
| `openai` | OpenAI 兼容 `chat/completions` | 要 | 用 `web_search_options` 触发联网 |
| `tavily` | Tavily Search API | 要 | 结果自带摘要片段 |
| `brave` | Brave Search API | 要 | 独立索引 |
| `bocha` | 博查 BochaAI | 要 | **中文长尾覆盖好、国内直连** |
| `exa` | Exa | 要 | 语义检索（找观点 / 论文） |
| `serper` | Serper | 要 | 就是 Google 那一页的有机结果 |
| `searxng` | 自建 SearXNG | 不要（要地址） | 自己的实例 |
| `duckduckgo` | DuckDuckGo HTML | 不要 | 零配置兜底（抓结果页解析） |

**怎么选**：默认 `search.provider: auto` —— 按 `search.providerOrder` 依次试，
**第一条成功的就用它**。顺序里没配 Key 的通道会当场失败（不发请求），所以顺序长也不拖时间。
想固定一条就写 `provider: bocha`；想按次点名就用
`office_search_run({ queries: [...], provider: 'bocha' })`。
全部失败时错误里会把**每条通道的原因**分别写出来（配置缺失 / 网络不可达 / 被挡 / 没结果），
而不是压成一句「查不到」。

Key 的解析顺序：配置里的字面量 → 宿主凭据服务 → 进程环境变量 →
`$DSH_HOME/.credentials.yaml`。

```yaml
search:
  provider: auto
  providerOrder: [seam, anthropic, openai, tavily, brave, bocha, exa, serper, searxng, duckduckgo]
  providers:                  # 每条通道只留它用得上的键
    anthropic: { apiKey: '', apiKeyEnv: DEEPSEEK_API_KEY, baseURL: https://api.deepseek.com/anthropic/v1, model: deepseek-v4-flash, timeoutMs: 60000 }
    tavily:    { apiKey: '', apiKeyEnv: TAVILY_API_KEY, timeoutMs: 30000 }
    searxng:   { baseURL: http://127.0.0.1:8080, timeoutMs: 30000 }
    duckduckgo: { baseURL: https://html.duckduckgo.com/html, timeoutMs: 30000 }
    # seam 不吃参数
  builtin: …                  # 老入口：与 providers.anthropic 是同一件事（同名键以 providers.anthropic 为准）
```

### 取回来的网页先过预处理（第二十轮）

`src/web-preprocess.js` 是一条**确定性脚本管线**（无模型、无网络、无依赖）：

```
去注释 → 去 script/style/svg/form 等非内容标签 → 去 nav/header/footer/aside
      → 去 class/id 命中 nav|menu|sidebar|footer|cookie|consent|advert|share|related|subscribe… 的样板区
      → 挑主容器（article / main / role=main / id|class 带 article|post|content|entry|main，取文本最长的）
      → 转纯文本（不带链接标记 —— URL 属于来源清单，不属于正文）
      → 行级清洗（丢样板行、丢「短且不像句子」的导航行、重复行去重、压空行）
```

三种模式：`article`（默认）/ `plain`（只做行级清洗）/ `off`（原样，等于回到第十八轮）。
报告会进结果文件与反馈：`正文预处理：1 页（article 模式）：558 → 117 字符（去掉 79%；丢样板行 1、去重 1）`。
按次改强度：`office_search_run({ queries: [...], preprocess: 'off' })`。

**哨兵读清洗前的文本**：清洗会把拦截页的样板一起删掉，拿清洗后的文本判会把
「目标站拒绝」误判成「正文几乎为空」，而这两类失败的修法完全不同。

### 成败怎么判：内容级哨兵 + 四类失败（第十八轮）

**状态码不足以判断一次取正文成没成。** 2026-09-26 本机实测：直连一个被墙站点的
文档页拿到 **HTTP 200**，内容却是 447 KB 的「本区域不可用」—— 只看 `response.ok`
就会把封锁页当成正文喂给模型，**比直接报错更糟，因为它看起来是成功的**。
所以每次取正文都过一遍 `inspectFetchedPage`，判定顺序按可靠性排：

| 序 | 信号 | 判成 |
| --- | --- | --- |
| 1 | `4xx`（401/403/404/429/451…） | 目标站拒绝（4xx 连正文都不读，读完也是浪费） |
| 2 | `5xx` | 网络出口不可达（服务端错误） |
| 3 | 最终地址是区域封锁页 / 人机校验页 | 目标站拒绝 |
| 4 | 正文命中拦截页、区域封锁、登录同意墙、验证码特征（**仅在正文 < 4000 字符时扫**，避免长文里的巧合短语误判） | 目标站拒绝 |
| 5 | 清洗后 < 20 字符 | 没拿到结果 |
| 6 | 清洗后 < 200 字符 | 通过，但附一句「正文偏短」提醒 |

**四类失败分开报**（`WEB_FAILURE_KINDS`），因为修法完全不同：

| 分类 | 含义 | 该怎么办 |
| --- | --- | --- |
| `config` | 配置缺失 | 去设置页配 Key 与端点 |
| `network` | 网络出口不可达 | 连不上/超时/5xx —— 换时间或换出口再试 |
| `blocked` | 目标站拒绝 | 401/403/429/451、区域封锁、人机校验、页面不存在 —— **换一个来源，别重试同一个地址** |
| `empty` | 没拿到结果 | 换关键词或换角度 |

反馈里按分类**计数**而不是逐条重复原因（同因合并）：`取正文失败 3 条（目标站拒绝 2、
网络出口不可达 1）`。一条来源取正文失败不会让整轮直查或整条渠道失败，但它一定会被写进
结果文件与反馈 —— 在此之前这里是 `catch {}` **静默吞掉**的，于是「被墙 / 被挡 / 页面没了」
在模型眼里等价于「这条来源没有摘录」。`test/web.mjs` 有一条漂移守卫：新加错误码却忘了
登记分类会当场红（否则新码静默落进 `other`，「四类分开报」就退化了）。

### 直查：`office_search_run({ queries, out?, maxResults?, fetchPages?, provider?, preprocess? })`

只查一两个事实点时不必走提纲：给 1–5 条查询，插件去查、把来源与摘录写进
`.office/search/<slug>/run.md`，聊天里只回一份紧凑清单（标题 + URL）。
`provider` 按次点名一条通道（不传就按设置页的顺序自动挑），`preprocess` 按次改预处理强度。
清单是线索不是结论，写进文档前按 URL 核对。

### 第一步：`office_search_brief({ topic, type?, audience? })`

按内容类型给出该走哪些渠道，并把提纲写成文件（`.office/search/<主题>/brief.md`）。
`type` 取 `hotspot` / `knowledge` / `manual` / `mixed`；省略时按主题文字自动判断。

**渠道分流规则**（写在 `src/search-routes.js`，是数据不是提示词）：

| 内容类型 | 主渠道 | 辅渠道 |
| --- | --- | --- |
| 热点事件 | 权威媒体（定事实骨架）、社交平台（现场与争议） | 视频评论、国际讨论、原文核实 |
| 知识类 | 百科（定义与体系） | 问答、文献与一手资料、数字核实 |
| 手册类 | 官方参考文档、官方仓库（版本） | 官方包说明；社区只能作参考 |
| 混合/未指明 | 先泛搜判类型，再按类走 | — |

三条硬规则：

1. **先泛搜**。每个类型的第一条渠道都是不限定来源的一轮搜索，用来找准关键词、
   发现问题与该领域的实际叫法——这一步是防茧房的关键，不能跳。
2. **★ 渠道必须覆盖**。提纲里标 ★ 的是该类型的必查渠道，缺了会在解析摘要里
   被点出来。
3. **平台直连失败要有兜底**。`platform_search` 依赖各平台公开接口，某些网络环境
   会拦住（本机实测 wikipedia / v2ex / reddit 直连失败，bing / github /
   stackoverflow / npm / hn / bilibili 正常）。所以每个平台渠道都配了
   `fallbackQueries`：直连失败改用搜索引擎搜该平台内容，结果必须标注
   「经搜索引擎间接取得」。

### 第二步：`office_search_dispatch({ briefPath, outputPaths })`

按提纲逐渠道执行，`outputPaths` 必须与渠道数一一对应。两条通道：

- **子代理通道**（默认优先）：组合里有子代理、**且工具面里真有那四个联网工具**时，
  **每个渠道一个子代理**，各自去调 `web_search` / `advanced_search` / `platform_search` /
  `web_fetch`，把材料写进指定的结果文件。子代理只回一句「已写入 <路径>」，原文不进主上下文。
  通过 `ctx.get('subagents').start('spawn', …)` 起一次性子代理（检索有明确终点，
  等它写完再继续比后台可继续子代理好收口）；并发用简单闸门限制，默认最多 4 个。
- **内置通道**（自动回退）：没有 `subagents` 服务、或组合里缺那几个检索工具时，插件用内置检索
  在进程内跑同样的渠道 —— 每个渠道按提纲的查询去查、去重、必要时打开页面取一段正文当摘录，
  写成**同样格式**的结果文件，所以第三步不用改。

**「组合里缺检索工具」是装配期就知道的事实，不是运行时意外**（第十八轮 P1-11）。
办公 preset 是一份完整的组合，里面没有 `@deepseek-ai/dsh-tool-web`，所以
`tools.restrict({ allow: CHANNEL_TOOLS })` **必然**被拒。在此之前这件事是靠**每个渠道各撞一次
失败**才发现的，那串原始报错还会贴在**每一个渠道**的反馈行上 —— 模型据此写出过「本工作区会话
没有联网检索工具」的 `[critical]` 记忆（session6），下一轮工具修好之后那条记忆就是错的。

现在的做法：`missingChannelTools(ctx)` 经 **`ctx.tools.schemas(scope)`** 读一次可见的工具名
（`tools.restrict()` 的报错本来就是拿 `restrictableNames` 比出来的，同一个来源），三态判定：

| 返回 | 含义 | 行为 |
| --- | --- | --- |
| `[]` | 齐备 | 照旧派子代理 |
| 非空数组 | 白名单一定会被整批拒绝 | **不试**，直接用内置通道 |
| `undefined` | 问不出来（没有 tools 服务 / 没有 `schemas` / 读不到名字 / `schemas` 抛错） | **保持老行为**照旧试一次 |

**绝不用「不知道」去推断「没有」**：误判会让本来能派子代理的部署静默退化成只走内置通道。
缺口由 `renderDispatch` 在**底部说一次**「这个组合里没有联网检索工具（缺 …）：本次全部走内置
通道。这不是错误，也不用绕路去派 `spawn_teammate`」；探不出来时撞上该报错的兜底会把长报错收成
一句事实（**保留缺失的工具名**、丢掉 `known global tools` 清单）并按会话记一次（有界 64、
拿不到会话身份不记）。

`search.engine` 可以钉死用哪条：`auto`（默认）/ `subagent` / `builtin`。

**服务必须经 `ctx.get('subagents')` 取，不能直接读 `ctx.subagents`**：cordis 的 ctx
是个代理，读一个没写进 `inject` 的服务属性会**当场抛错**
（`cannot get property "subagents" without inject`），而不是返回 undefined。这个错
让第二步从上线起就没在真实会话里跑通过（单测用的是普通对象替身，恰好没复现宿主这条
行为）；2026-09-25 修复，复现与守门见 `test/subagent-seam.mjs`。同理，`'subagents'`
也不该写进模块级 `inject`——那是必需依赖，缺这个服务的精简部署会让整个插件起不来。
两条通道都不行时（例如既没有子代理、又没有 Key），派工才会失败，并**明确指向
Agent Teams**：按提纲逐渠道 `spawn_teammate`，写完再用 `office_parse_findings` 读回。
同样地，`ctx.get('web')` 也走 `ctx.get`，取不到就退到自带 HTTP，而不是让插件起不来。

### 第三步：`office_parse_findings({ paths, type?, topic? })`

把结果文件读成紧凑摘要。做四件事：

- **按渠道归类**结论，并算出提纲里哪些 ★ 渠道还没覆盖（「★ 仍缺渠道」）。
- **跨源核对**：把讲同一件事的结论合簇（并查集 + 字面重合度），看这一簇背后
  有几个互不相关的站点。这正是茧房效应最容易骗过人的地方——搜出来一堆结果，
  其实全是同一篇稿子的转载。
- **标出单一来源**与**完全没带 URL** 的条目。
- **URL 从正文里摘掉**，结论与出处分开存放，避免长链被抄进文档。

### 子代理用什么工具

子代理**默认加入父方的组装**（也就是办公模式那一套，含 `office_run` 等），
所以派工时用 `toolFilter.allow` 把它裁成**极简面**——只有七个工具：

```js
toolFilter: { allow: [
  // 检索
  'web_search', 'advanced_search', 'platform_search',
  // 取正文
  'web_fetch',
  // 读（含图片：图表 / 截图 / 扫描件）
  'read', 'read_image',
  // 写（唯一交付方式）
  'write',
] }
```

| 给了 | 为什么 |
| --- | --- |
| `web_search` / `advanced_search` / `platform_search` | 三个渠道族，分流规则的落点 |
| `web_fetch` | 打开具体页面拿原文，不只依赖搜索摘要 |
| `read` / `read_image` | 读文本与图片；`read_image` 让图表、截图、扫描件里的信息可读 |
| `write` | 把结果写盘，子代理唯一的交付方式 |

| 刻意不给 | 理由 |
| --- | --- |
| `edit` / `glob` / `grep` | 只写自己那个已知路径的结果文件，不需要改文件或满目录找文件 |
| `office_*` | 检索子代理不生成文档 |
| `bash` / `pwsh` | 不执行命令 |
| 子代理与任务板 | 不派活给更下一层 |
| `present` / `todo` / `goal` / `skill` | 与「查资料、写文件」无关 |

用**白名单而不是黑名单**：以后 profile 里新装了什么工具，也不会悄悄出现在
检索子代理手上。`test/search.mjs` 里有一条测试把这个七项列表逐字钉住，
谁想加工具都会先看到它失败。

检索子代理**不是**一个可选择的模式：它由 `office_search_dispatch` 在派工时
临时创建，工具面由上面的 `toolFilter` 白名单决定。Agent 预设列表里只有
「办公模式」一项，不会多出一个「检索子代理」。

## 设置界面

装好插件后，设置对话框左侧会多出**两个**一级导航项，和「Wallpaper Engine」同一
层级、同样的交互形态（带图标的入口，右侧是可操作的控件）：**「记忆系统」**与
**「办公模式」**。不需要手改 YAML。

- **记忆系统**：三层记忆的开关、容量与归档规则，以及从 mnemon 迁移的命令。
  mnemon 关掉之前，这个位置是它提供的记忆页面；现在由本插件接住，导航里不会
  出现两个「记忆系统」。
- **办公模式**：工具开关、检索编排、文档与缓存；记忆那一组只留一句指路
  （同一份设置在两个面板里各渲染一遍，容易让人以为改的是两个东西）。

这两页共用同一个设置命名空间 `dsh-office-mode`，由插件的**浏览器半侧**渲染：
`lib/client.js` 通过 `dsh.client` 声明被宿主加载，注册进 `settings.section` 槽位
（两个 section id：`dsh-office-memory` 与 `dsh-office-mode`）。读写走
`ctx.configForms`（0.1.7 起由 dsh-client-ui-settings 提供的基础服务：按插件行 id
取一份 ConfigForm，写入带回 revision 栅栏），服务端 schema 仍是
唯一真源——前端只负责展示与写回；两页共用一套读写接线（`SettingsPanel`），
只读提示与错误显示因此不会各自漂移。

### 字号与「操作简易」（第二十轮）

字号**不再写死**，改用宿主的类型刻度 `--dsw-font-*`（`xxxs-11` / `xxs-12` / `xs-13` /
`s-14` / `base-16`）：行标题 14px、说明 12px、控件 13px 就是它们解析出来的值
（真实浏览器实测）。三个分量（`-font-size` / `-line-height` / `-font-weight`）各自取令牌、
各自带字面兜底 —— `font:` 简写一旦整条无效，字号会跟着一起丢，分着写就不会。
次级文字一律用颜色令牌（`label-secondary/tertiary/caption`）而不是 `opacity`。

操作上做了四件事：

1. **「联网通道」下拉**（`auto` + 十条通道）+ **一把粘 Key 的输入框**
   （写 `search.providers.<id>.apiKey`）—— 新能力两步就能用起来；
2. **长参数收进折叠块**（通道端点/模型/超时 + 内置检索细项）：默认收起但**内容照渲染**
   （`display:none`），切换不重建子树、草稿不丢，单测也仍找得到控件；
3. **「auto 尝试顺序」用可点标签加减**（没有拖拽排序也能调）；
4. **工具开关加「全部打开 / 全部关闭」**。

> 界面必须覆盖服务端 schema 的**每个参数**是一条硬守卫（`test/client.mjs` 逐字核对源码里的
> 写回路径）。新增的通道参数用一张**字面量表**（`CHANNEL_TABLE`）同时承担「界面数据」与
> 「守卫凭据」两个角色：拼出来的路径（`"search.providers." + id + …`）在源码里搜不到，
> 等于把守卫废掉。

### 工具开关

| 开关 | 关掉后 |
| --- | --- |
| `office_help` | 模型不再能按需查写法（不建议关，会显著增加写错 API 的概率；记忆投影也挂在这个返回值上） |
| `office_run` | 不注册批量执行工具，办公模式退化成纯检索/写作 |
| `office_memory` | 不注册记忆工具（`office_run` 的自动台账仍按下面的 `autoLedger` 走） |
| `office_search_run` | 不注册内置检索直查，只能走提纲三步 |
| `office_search_brief` | 不出检索提纲 |
| `office_search_dispatch` | 不用子代理/内置通道，只能按提纲手工检索 |
| `office_parse_findings` | 不解析结果文件，需要自己读文件 |

**关掉的工具不会出现在模型面前，也不占每次请求的 schema 开销。** 开关是
**即时生效**的：改完立刻重建工具面，不用重启宿主。

### 记忆（在「记忆系统」页）

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `enabled` | 开 | 总开关：关掉后记忆工具、提示段与自动台账一起停 |
| `dir` | `.office/memory` | 记忆目录（相对会话工作目录）；一个项目一份记忆 |
| `autoLedger` | 开 | `office_run` 写出文档时自动登记台账；关掉仍保留记忆投影 |
| `promptHint` | 开 | 往系统提示加一段静态说明（不含记忆内容）；办公模式下会被 persona 丢掉 |
| `userLimitBytes` | 4096 | 热记忆里「用户偏好」的容量上限，超出时下沉到归档 |
| `projectLimitBytes` | 10240 | 热记忆里「项目与环境」的容量上限，同上 |
| `ledgerLimit` | 500 | 台账最多保留多少条，超出的最旧记录滚成月度归档摘要 |
| `archiveKeep` | 60 | 归档最多保留多少个摘要文件（归档自身也是有界的） |
| `layers.hot` / `.ledger` / `.archive` | 全开 | 层拓扑开关。关掉是「不再读、不再写」，不删数据 |
| `scope` | `workspace` | 记忆范围：`workspace` / `global` / `both`（跨项目层） |
| `globalDir` | 空 | 全局层目录；留空 = `$DSH_HOME/.office/memory`（相对路径相对用户主目录） |
| `userScope` | `memory` | `global` 时「用户偏好」始终落全局层 |
| `links` | 开 | 图关系：`link` / `related` / 实体视图 |
| `autoCapture` | 开 | 主动记录**指引**（不是自主记录器） |
| `recallQuality.policy` | `strict-v1` | `strict-v1` 分档过滤 / `off` 不丢结果 |
| `recallQuality.lowScoreThreshold` | 0.25 | 低于它算「未知」档（分数是命中词占比，不是余弦相似度） |
| `recallQuality.highScoreThreshold` | 0.6 | 达到它直接采纳；必须大于下限，写反了整组退回默认 |
| `recallQuality.candidateMultiplier` | 3 | 先取「条数上限 × 倍数」个候选再分档 |
| `recallQuality.maxMediumResults` | 4 | 「中档」最多采纳几条 |
| `recallQuality.maxUnknownResults` | 2 | 「未知档」最多采纳几条（0 = 一条不要） |
| `quota.recallPerTurn` | 1 | 一个回合里第一次带 `query` 的检索（0 = 不限制） |
| `quota.recallRefinePerTurn` | 1 | 同一回合里后续的换词细化 |
| `quota.relatedPerTurn` | 1 | 一个回合里的图关系遍历 |

### 后台任务 Agent 的模型路由

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `mode` | `inherit` | `inherit` 跟随主会话；`fixed` 用下面指定的模型 |
| `provider` | 空 | `fixed` 时的 provider 路由名；留空 = 只覆盖 `model` |
| `model` | 空 | `fixed` 时的模型 id；留空按 `inherit` 处理 |

只影响检索三步走的**派工子代理**，不影响主对话。覆盖需要子代理后端声明
`SubagentCapabilities.agentOptions`；不支持时 `start` 会**响亮地拒绝**，不会静默
退回继承。

### 检索编排

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `engine` | `auto` | `auto` = 先派子代理、跑不了改用内置检索；`subagent` = 只用子代理；`builtin` = 只用内置检索 |
| `subagentTools` | 七项极简白名单 | 检索子代理能用哪些工具。加得越多子代理越容易跑偏 |
| `maxParallel` | 4 | 同时铺开几个检索子代理（1–8），渠道多时别一次全开 |
| `resultLimit` | 40 | 解析结果最多列几条结论（1–40） |
| `maxChannels` | 12 | 一次提纲最多几个渠道（1–12） |
| `outputDir` | `.office/search` | 提纲与结果文件目录 |
| `requireCrossSource` | 开 | 单一来源背书的结论会被点名要求补检索 |
| `fallbackOnPlatformError` | 开 | 平台直连失败时改用网页搜索兜底 |
| `provider` | `auto` | 用哪条联网通道：`auto` 按 `providerOrder` 依次试；也可以固定成某一条 id（见上表） |
| `providerOrder` | 十条通道的默认顺序 | `auto` 的尝试顺序（设置页里点标签加/移） |

`search.providers.<id>` 是每条通道各自的参数（**键集就是它真正用得上的那几个**）：

| 通道 | 参数 |
| --- | --- |
| `seam` | 无（空对象） |
| `anthropic` / `openai` | `apiKey` / `apiKeyEnv` / `baseURL` / `model` / `timeoutMs` |
| `tavily` / `brave` / `bocha` / `exa` / `serper` | `apiKey` / `apiKeyEnv` / `timeoutMs` |
| `searxng` / `duckduckgo` | `baseURL` / `timeoutMs` |

`search.preprocess.*` 是网页预处理（`src/web-preprocess.js`）：

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `mode` | `article` | `article` 去导航与样板 + 挑主容器 / `plain` 只做行级清洗 / `off` 原样返回 |
| `dropBoilerplate` | 开 | 丢掉 cookie 条 / 登录注册 / 订阅分享 / 版权声明这类样板行 |
| `dedupeLines` | 开 | 同一条重复出现的行只留第一次 |
| `minLineChars` | 12 | 短于它、又不像句子的行当导航丢掉（0 = 关掉这条，清单型页面建议关） |
| `keepTitle` | 开 | 抽出标题 / 作者 / 时间放进预处理报告（不进正文） |

`search.builtin.*` 是 Anthropic 兼容通道的细项（**与 `search.providers.anthropic` 是同一件事
的两个入口，同名键以 `providers.anthropic` 为准**；老配置只写 `builtin` 也照旧生效）：

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `apiKeyEnv` | `DEEPSEEK_API_KEY` | 取 Key 的环境变量名；先问宿主凭据服务，再看环境变量，最后看 `$DSH_HOME/.credentials.yaml` |
| `apiKey` | 空 | 字面量 Key（一般留空，用上面的环境变量名） |
| `baseURL` | `https://api.deepseek.com/anthropic/v1` | Anthropic 兼容 Messages 接口的 base（末尾拼 `/messages`） |
| `model` | `deepseek-v4-flash` | 检索用模型，要支持原生 `web_search` |
| `maxUses` | 5 | 一次检索最多调用几次原生 `web_search` |
| `maxTokens` | 2048 | 一次检索的生成上限（只要清单，不要长文） |
| `maxResults` | 8 | 每个渠道 / 每条直查查询最多几条来源 |
| `fetchPages` | 2 | 每个渠道 / 每条查询最多打开几页取正文摘录（0 = 只用搜索片段） |
| `searchTimeoutMs` | 60000 | 单次检索超时 |
| `fetchTimeoutMs` | 20000 | 单页取正文超时 |
| `maxBytes` | 2097152 | 单页体积上限（超出即截断） |
| `maxChars` | 20000 | 单页字符上限 |
| `maxRedirects` | 3 | 取正文最多跟随几次**同源**跳转（跨站跳转一律不跟） |

环境变量覆盖：`DSH_OFFICE_SEARCH_BASE_URL` / `DSH_OFFICE_SEARCH_MODEL` /
`DSH_OFFICE_SEARCH_API_KEY_ENV`，端点也认宿主的 `DEEPSEEK_SEARCH_BASE_URL`。

### 文档与缓存

`defaultTheme`（下拉选主题）、`scriptTimeoutMs`、`maxScriptChars`、`cacheDir`、
`keepCache`。

### Python 计算与绘图

`enabled`（默认开；关掉后 `office.python.*` 明确报错并说明在哪儿打开）、`bin`
（解释器，留空自动探测，也可填绝对路径）、`timeoutMs`（默认 120000，单次运行上限）、
`outDir`（产物目录，相对缓存目录，默认 `python/out`）。

### 音频与视频

`enabled`（默认开）、`ffmpegPath` / `ffprobePath`（留空自动探测；**填了就以它为准**，
路径不存在即报缺件）、`modelDir`（SenseVoice 模型目录，留空 = 语音输入下载的那份）、
`language`（`auto` / `zh` / `yue` / `en` / `ja` / `ko`）、`chunkSeconds`（默认 120）、
`maxSeconds`（默认 3600）、`timeoutMs`（默认 300000）、`precision`（`int8` / `fp32`）、
`threads`、`frames` / `maxFrames`（默认 6 / 12）、`frameFormat`（`jpg` / `png`）、
`frameWidth`（0 = 原分辨率）。

配置文件里还能调几个不进设置页的细项：`vadModel`（Silero VAD 路径）、`runtimePath`
（sherpa-onnx 运行时，留空从宿主解析）、`vadThreshold` / `minSpeechSeconds` /
`minSilenceSeconds` / `maxSegmentSeconds`（VAD 切句参数）、`audioDir` / `framesDir`
（缓存的子目录）、`words`（是否带词级时间戳）。

### 越界调不坏插件

所有数值参数都在 schema 里带上下界，设置页调到越界会被挡下；运行期还会再
收敛一次（`resolveConfig`），所以即使有人直接改配置文件写了 `maxParallel: 999`，
插件也只会用 8，不会崩。

### 没有设置能力的部署

宿主没有 `settings` 服务、或解析不到 `schemastery` 时，本插件的 `Config` 退回
「无 schema」（cordis 会把 config 原样透传），设置页不注册，插件其余功能完全
不受影响——**空配置永远可用**这条约束不变。

> 关于 `Config` 与 `schemastery`：它不是本插件的依赖，由宿主提供。插件是 `link:`
> 进 profile 的（源码在工作区），直接 `import` 裸包名会 `ERR_MODULE_NOT_FOUND`，
> 所以按「profile 根 → 宿主安装位置 → 宿主入口所在目录」的锚点依次**同步**
> `require`——`export const Config` 必须模块求值时就已经有值，用不了异步 import。

> **可编辑字段必须标成 volatile**：0.1.7 起 dsh-settings 只接受 volatile 路径的
> 写入，而 volatile 字段由宿主就地更新、不重挂插件。schema 里由
> `markLeavesVolatile` 整片置位（只标叶子、容器不标，否则会撞上「volatile 字段
> 不能嵌套」的校验）；宿主半侧每次重建前用 `unwrapVolatile` 把引用拆成纯值，
> 并监听 `settings/document-updated` 重建工具面（旧版的 `settings.onChange`
> 回调契约已经没有了）。

> 关于浏览器半侧：`lib/client.js` 是**手写的单文件 bundle**（`window.__ModuleLoader__.load`
> + `require("react")`），不经过打包步骤——宿主只要求这个文件存在。它只 `require`
> 平台基座里的 `react`，没有别的外部请求，所以 `dsh.client.external` 留空。
> `inject` 只声明 `slots`：`configForms` 是可选服务，拿不到时这一页显示可读提示
> 而不是白屏；把可选服务写进 `inject` 会让缺服务时整个 `apply` 都不跑，
> 连记忆面板一起消失。

## 默认长什么样

不传 `theme` 时用 `plain`（素色网格）：

- **Excel**：保留普通表格的网格形态——细边框、加粗表头、右对齐数字，
  **不写任何背景填充**（标题行、表头、斑马纹、合计行的 Fill 全为空）。
- **Word**：白纸黑字，标题靠字号与层级留白区分，表格只有边框，
  **不写 `w:shd` 底纹**；引用块只留左侧竖线，代码块只留边框。
- 字体回到 Word / Excel 的默认组合（等线 + Calibri），不额外指定排版字体。

需要配色时显式选主题（`create({ theme: 'business' })`），主题模板会照旧
给出表头底色、斑马纹与强调色。主题数据里的约定是：**空串 = 不填充**，
渲染前一律判断，绝不把空值写成 `rgb=""` / `w:fill=""` 这类坏 XML。

## 轻量办公替代清单

目标不是「生成一个像文档的文件」，而是「一句话能说清的办公活，不必打开 Office 手工做」。
常用动作与写法：

| 要做的事 | 怎么写 |
| --- | --- |
| 算数：统计 / 拟合 / 模拟 / 单位换算 | `office.python.run('import numpy as np …')`（先 `office.python.check()` 看装了哪些包） |
| 画图：折线 / 柱状 / 散点 / 饼图 | Python 里 `fig.savefig(OUT_DIR + '/chart.png')`，产物路径交给 `builder.image()` / `deck.image()` |
| 长分析脚本 + 命令行参数 | `office.files.write('analysis.py', code)` 后 `office.python.file('analysis.py', { args: ['2026'] })` |
| 会议录音 / 访谈 / 语音条 → 带时间戳的文字 | `office.av.transcribe('录音.m4a', { language:'zh', out:'逐字稿.md' })`，再读那个文件 |
| 视频 → 逐字稿 + 关键画面 | `office.av.extract('会议.mp4', { out:'逐字稿.md', count:6 })`，画面逐张 `read_image` |
| 视频里的板书 / 幻灯片截帧 | `office.av.frames('课程.mp4', { every: 60, width: 1280 })` |
| 先确认本机能不能做音频 / 视频 | `office.av.check()`（ffmpeg / ffprobe / 模型 / 运行时四项一次报全） |
| 一列数字求和 / 平均 / 最大最小 / 标准差 / 中位数 | `sheet.stats('C5:C28')` —— 写成真公式（AVERAGE 等）并带缓存值 |
| 按某列分组汇总（轻量透视） | `sheet.summary({key:'部门', value:'金额', funcs:['sum','count','average'], total:true})` |
| 手写公式与数字格式 | `sheet.formula('F3', 'SUM(C5:C28)')`、`sheet.numberFormat('C5:C28', '¥#,##0.00')` |
| 数据表带合计行 | `sheet.table({..., totalRow: true})`（数字列自动 SUM） |
| 论文 / 试卷里的数学公式 | `builder.formula('\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}', {number:'（1）'})` |
| 行内公式 | `builder.para([{text:'由'}, {math:'E=mc^2'}, {text:'可得'}])` |
| 合并单元格的表格 | `builder.table({columns:[...], rows:[[{text:'合计', colspan:2}, '1200']]})` |
| 表格题注 | `builder.table({..., caption:'表 1 分月达成情况', captionPosition:'above'})` |
| 单张图片（裁切 / 对齐 / 边框） | `deck.image({path:'chart.png', fit:'cover', align:'center', frame:true})` |
| 一页多图 | `deck.images([{path:'a.png'}, {path:'b.png'}], {columns:2})` |
| 满页背景图 | `deck.image({path:'cover.png', fullBleed:true})` |
| 自绘卡片 / 色块 / 箭头 / 横幅 | `deck.shape({preset:'roundRect', x, y, w, h, fill:{type:'gradient', stops:[…]}, line:{dash:'dash'}, shadow:true})`、`deck.line({from:[1,2], to:[8,2], arrow:'end'})` |
| 全篇统一的背景图 / 校徽 / 页眉 / 页码 | `deck.master({background:{image:{path:'bg.png', dim:0.35}}, logo:{path:'logo.png'}, header:'2026 年度汇报', pageNumber:{show:true}})` |
| 段内多色文字混排 | 段落写成 `runs:[{text:'重点', bold:true, color:'C00000'}, {text:'其余'}]`（形状内文字与表格单元格同样支持） |
| 合并单元格的表格 | `deck.table({columns, rows:[[{text:'合计', colSpan:2}, 3860, '107.2%']]})` |
| 超链接单元格 | `deck.table({rows:[[{text:'官网', link:'https://example.com'}]]})` |
| 卡片网格 / 编号流程 / 对比双栏 / KPI | `deck.cards({title, items:[{title, body, icon}], columns:2})`、`deck.steps({items:[…]})`、`deck.compare({left, right})`、`deck.kpi({items:[{value, label, unit}]})` |
| 图+文 / 时间线 / 图标网格 | `deck.imageText({image:{path}, items:[…]})`、`deck.timeline({items:[{time, title, body}]})`、`deck.iconGrid({items:[{icon, label}]})` |
| 提示词面板 / 横幅色带 | `deck.panel({title, text})`、`deck.banner({text, gradient:{stops:[…], angle:0}})` |

**生成之后：审阅与微调**（`office.ppt.readSlides` / `office.ppt.revise`）

| 要做的事 | 怎么写 |
| --- | --- |
| 看清每一页有什么形状、什么几何 | `office.ppt.readSlides('汇报.pptx')` → 逐页 `{id, kind, name, text, paragraphs, x, y, w, h, editable}`（厘米） |
| 改一处文字（保留原字体字号） | `office.ppt.revise(path, [{slide:1, shape:'title', setText:'2026 年度汇报'}])` |
| 只换段内某个词 | `{slide:5, shape:'Card 2 Body', replace:[['待定','已确认']]}` |
| 改字号 / 字体 / 颜色 / 对齐 | `{slide:'2-4', shape:'title', style:{sizePt:30, color:'1F4E79', font:'微软雅黑', align:'left'}}` |
| 只改段内命中的那截文字 | `{slide:1, shape:'Bullets 1', style:{sizePt:24, text:'重点'}}`（段内多色混排时用） |
| 挪位置 / 改尺寸 | `{slide:5, shape:'Card 2', move:{x:2, y:6}, resize:{w:8}}`（也支持相对量 `{dx,dy}` / `{dw,dh}`） |
| 让文本框贴合内容 | `{slide:1, shape:'Bullets 1', fit:true}` |
| 一次批量改多页多处 | 把若干 op 放进同一个数组，一次调用全部落地 |

`slide` 选择器收页码（1 起）、`'2-5'`、`'2,4'`、`[2,3]`，省略即全篇；
`shape` 收形状 id、形状名（`'Card 1 Body'`）或类型名（`title`/`subtitle`/…）。
**就地改而不是重生成**：只替换命中的那段 XML，同一页其余形状、母版元素、
备注与超链接关系一个字节都不动（这一条有断言钉住，见 `test/format-ppt-revise.mjs`）。
一条 op 失败只进 `skipped` 并带上原因，不影响同批的其它 op；全部没生效则不落盘。

Excel 的统计块与分组汇总写的是**真公式 + 缓存值**：Excel 打开时重算，
不重算的阅读器也能直接看到数字。

整篇观感的验收样例见 `办公模式示例\参考稿复刻.pptx`：12 页，依次是封面（满页照片 + 压遮罩 +
校徽）、目录（2×2 大号编号）、章节页、卡片网格、对比双栏、编号流程、合并单元格表格、图+文、
KPI、时间线、图标网格与结束页，由 `node test/make-reference-deck.mjs` 生成。

## 反馈长什么样

`office_run` 结束后，每个写出的 Office 文件都会被重新打开解析一遍：

```
✅ office_run 完成（412 ms）：做季度汇报三件套
文件（3）
1) 季度汇报.docx  [word · business]  18.4 KB
   段落 24｜标题 5｜表格 2｜约 3 页
   结构：H1 2026 年第一季度业绩汇报 / H1 一、整体情况 / 表格 4×5 …
   ⚠ 第 3 页表格列宽合计超出版心约 6%
2) 季度数据.xlsx  [excel · business]  9.1 KB
   sheets 1｜rows 14｜列宽充足
3) 季度汇报.pptx  [ppt · business]  31.7 KB
   15 页｜cover 1｜bullets 9｜table 2｜closing 1
   ⚠ P6 要点 9 条，超过 7 条，建议拆页
缓存：.office/cache 保留 1 个中间文件（20 B）
   720 分钟没碰过会在下次调用开始时清掉；要立刻清空用 office.cache.clear()
```

统计来自真实解析，不是脚本自述 —— 脚本写错了什么，这里会露出来。

## 可迭代 / 可扩展

```
src/
  index.js        插件入口（注册工具；injectGuide 时注入提示段落）
  config.js       配置校验（空配置永远可用）
  tools.js        六个工具定义 + 反馈渲染
  run.js          批量执行引擎（vm 沙箱 + 结果复检 + 缓存生命周期 + 自动台账）
  memory.js       三层记忆（热记忆 / 台账 / 归档）：存储、写入协议、容量下沉、有界读取
  sdk.js          脚本可见的 office 对象
  pdf.js          PDF：引擎探测 / info / text / pages（渲染结果进缓存，可命中）
  tex.js          LaTeX：TeX 发行版与模板探测 / 编译 / 日志解析（不打包 TeX）
  docs.js         按需文档（格式部分从模块 meta 自动生成）
  guide.js        办公模式工作约定（唯一一份）
  registry.js     格式表 —— 加格式只改这里
  engine/         zip / xml / kit / theme / cache（共享底座）
  formats/        word.js / excel.js / ppt.js / ppt-revise.js / tex.js
```

`ppt.js` 负责「从零生成」，`ppt-revise.js` 负责「读回已有包、在原 XML 上定点改」。
两者刻意分文件：共用的只有单位换算与枚举收敛，而 revise 的难点（关系表、形状定位、
`a:rPr` 子元素保序）自成一套。`office.ppt` 上把它们拼成一个对象 ——
`ppt.js` 导出 `api`（惰性工厂），`sdk.js` 的 `wrap()` 把工厂结果挂上去。
挂载面有断言钉住（`smoke.mjs` 的「office.ppt 暴露 readSlides / revise」），
因为这种「静默少一个方法」不会让插件加载失败，只会让脚本拿到
`undefined is not a function`。

- **加一种格式**：写 `src/formats/<id>.js`（契约见 `SPEC-formats.md`），
  在 `registry.js` 的 `FORMATS` 里加一行。它自动获得脚本入口、help 文档、
  生成后复检，工具声明不变。格式若要暴露额外方法，导出 `api`（见 `ppt.js`）。
- **加一个主题**：往 `engine/theme.js` 的 `THEMES` 里加一条，三种格式同时生效。
- **改说话方式**：改 `presets/office` 的 persona，或 `src/guide.js`。

## 配置（都可选）

```yaml
- id: dsh-office-mode
  name: dsh-office-mode
  config:
    cacheDir: .office/cache      # 中间产物目录，相对会话工作目录
    keepCache: true              # 调用结束后保留中间产物（默认；false = 老行为，每次清空）
    cacheTtlMinutes: 720         # 中间产物多少分钟没被碰过就清掉（0 = 不按时间清）
    scriptTimeoutMs: 60000       # 单次脚本超时
    defaultTheme: plain          # plain/minimal/business/warm/forest/tech/academic
    pdfDpi: 120                  # PDF 渲染默认分辨率
    pdfMaxPages: 20              # 单次 office.pdf.pages() 页数上限
    pdfEngine: auto              # auto/fitz/pdftoppm/pdftocairo/mutool/gs
    texEngine: auto              # LaTeX 编译入口：auto/latexmk/xelatex/lualatex
    texTimeoutMs: 300000         # 单次 LaTeX 编译超时（学位论文要跑好几遍）
    texTemplateDir: ''           # thuthesis 模板目录；留空则自动探测
    injectGuide: false           # true 时往系统提示里注入办公模式约定
    python:                      # Python 计算与绘图（办公模式里唯一的编程出口）
      enabled: true              # 关掉后 office.python.* 报错，不删任何东西
      bin: ''                    # 解释器：留空自动探测 python / python3 / py，也可填绝对路径
      timeoutMs: 120000          # 单次运行超时
      outDir: python/out         # 产物目录（相对缓存目录）：脚本里拿 OUT_DIR 就是它
    memory:                      # 三层记忆（见上文「三层记忆」一节）
      enabled: true              # 总开关
      dir: .office/memory        # 记忆目录，相对会话工作目录
      autoLedger: true           # office_run 写出文档时自动登记台账
      userLimitBytes: 4096       # 热记忆「用户偏好」容量上限
      projectLimitBytes: 10240   # 热记忆「项目与环境」容量上限
      ledgerLimit: 500           # 台账条数上限，超出的最旧记录滚进归档
      archiveKeep: 60            # 归档摘要文件数上限
```

环境变量兜底：`DSH_OFFICE_CACHE_DIR`、`DSH_OFFICE_THEME`、`DSH_OFFICE_PDF_ENGINE`、
`DSH_OFFICE_TEX_ENGINE`、`DSH_OFFICE_TEX_TEMPLATE`、`DSH_OFFICE_MEMORY_DIR`、
`DSH_OFFICE_PYTHON`（解释器）、`DSH_OFFICE_PYTHON_OUT`（产物目录）、
`DSH_OFFICE_FFMPEG`（ffmpeg 路径）、`DSH_OFFICE_SENSEVOICE_DIR`（模型目录）、
`DSH_OFFICE_PDF_BIN`（额外搜索引擎的目录，分号分隔）。

## 垃圾文件规则

办公插件在**工作区里产生的全部中间产物都收在一个隐藏目录 `.office/` 下**，
工作目录根上不会再多出第二个插件目录：

| 目录 | 装什么 | 生命周期 |
| --- | --- | --- |
| `.office/cache/` | `office.cache.write` 的中间文件、PDF 渲染页图、抽出来的 PPT 图片、Python 产物（图 / 数据 / 脚本与日志）、音频转写解码出来的规范 WAV 与视频抽帧图（`av/audio`、`av/frames`） | 跨调用保留，按 TTL（默认 12 小时）与容量（默认 512 MB）清理 |
| `.office/search/` | 检索三步走的提纲与结果文件 | 跟着那次检索走，可随时删 |
| `.office/memory/` | 三层记忆（热记忆 / 台账 / 归档） | **不是缓存**：跨会话保留，不按 TTL 清理 |
| `.office/tmp/` | 测试中间产物（每套件一个子目录） | 每次跑测试重新生成 |
| `.office/diag/` | 浏览器取证脚本与截图 | 排查用，可随时删 |
| `.office/backups/` | 源码快照备份 | 可随时删 |

- **缓存跨调用保留**（默认）：调用开始时只清掉超过 TTL（默认 12 小时）没碰过的
  条目，结束时保留 —— 这样下一步还能接着用（最典型的是渲染好的 PDF 页面图要
  交给 `read_image`），同一份输入重复渲染才会命中缓存。
- 需要清空时：`office_run({ keepCache: false })` 或脚本里 `office.cache.clear()`；
  容量超过上限（默认 512 MB）时从最旧的开始删。
- 记忆记的是用户的偏好与这个目录的交付史，跨会话保留、不按 TTL 清理，要清只能显式
  删目录（`enabled:false` 只是停用）。全局层（工作区之外）在 `$DSH_HOME/.office/memory`；
  新路径不存在而旧的 `$DSH_HOME/.office-memory` 还在时会退回读旧的，不会把用户级记忆弄丢。
- 交付物必须写在当前工作目录，反馈里会给出相对路径。

> 2026-09-23 之前是「开始清一次、结束再清一次」，等于缓存永远不可能命中：
> 渲染好的页面图在第一次调用结束时就没了，`read_image` 只能读到「文件不存在」。
> 这是那一轮修掉的第二个问题，回归断言见 `test/pdf.mjs` 与 `test/smoke.mjs`。
>
> 2026-09-24 之前这些产物是散着的（`.office-cache` / `.office-search` / `.office-memory`
> 三个并列目录），本轮收进 `.office/` 一棵树。旧目录的数据已迁入（`.office-memory` 与
> `.backups` 走 `git mv`，历史跟着走）；`.gitignore` 只放行
> `.office/memory/{memory.json,USER.md,MEMORY.md}` 三个文件。

## 测试

```powershell
node test/smoke.mjs          # 引擎与管线（zip/xml/度量/缓存/脚本沙箱/错误行号/输出 JSON 安全）
node test/memory.mjs         # 三层记忆：存储与投影 / 写入协议 / 容量下沉 / 台账滚动 / 有界读取 / mnemon 迁移
                             #   第二轮：层拓扑 / 跨项目层 / 图关系（含下沉后不断） / 召回质量 / 每回合配额 / Pack / 实体 / 与工具接线
node test/view.mjs           # 记忆浏览端点：快照口径 / 只读（405） / 只服务见过的根（404） / 关掉时不泄漏
node test/pdf.mjs            # PDF：轻量解析 / 抽文本 / 渲染 / 缓存命中 / 缓存 TTL 与容量
node test/python.mjs         # Python 通道：解释器探测 / 中文 stdout 不乱码 / 回溯行号 / 超时 / 出图与产物归置 / 关掉后报错
node test/av.mjs             # 音频与视频：规范 WAV 判据与规范化 / 分块与抽帧时间点 / 缺件报错 /
                             #   真跑 ffmpeg 造的视频（info → transcribe → frames → extract / 解码缓存命中）
node test/search.mjs         # 检索：渠道分流 / 提纲 / 子代理任务书与白名单 / 内置通道回退与直查 / 解析与跨源核对（离线）
node test/web.mjs            # 内置联网通道：接缝优先与回落 / 响应映射 / HTML 收文本 / SSRF 分类 / 截断（离线）
node test/web-live.mjs       # 内置联网通道的真联网自测（**要联网**，不在默认套件里；没有 Key 时跳过）
node test/subagent-seam.mjs  # 在真实 cordis 上下文里验证 subagents 是经 ctx.get 取到的（含缺服务时的降级说明）
node test/settings.mjs       # 设置 schema：默认值 / 上下界 / volatile 叶子 / 与工具开关一致
node test/client.mjs         # 浏览器半侧：设置页渲染与写回 / 记忆面板 / 令牌与主题
node test/preset-check.mjs   # 用 harness 自己的规则校验本地 preset 形状
node test/injection-budget.mjs  # 提示词注入预算（第十七轮）：office_help 分层比例与关键接口不许丢 /
                             #   固定长句一个会话只说一次 / office_run 不回灌模型自己写的 outline
node test/e2e.mjs            # 端到端：走 office_run 真实路径产出三件套并复检（含自动台账）
node test/format-word.mjs    # 各格式模块自测（Word 1032 / Excel 275 / PPT 1291 项断言）
node test/format-excel.mjs
node test/format-ppt.mjs
node test/format-ppt-revise.mjs  # PPT 就地微调：读结构 / 改文字样式位置 / 改别人的包 / 插图与抽图（153 项）
node test/format-tex.mjs     # LaTeX 论文：生成项目 / 项目级复检 / 编辑 / 真实编译出 PDF（--no-compile 只测前三条）
node test/plain-default.mjs  # 默认素色网格：不传主题时 Excel/Word 没有任何背景填充

# 浏览器半侧的活体验证（第十轮新增，可复用）：起隔离实例 + 无头 Chrome，用 CDP 打开页面，
# 并用 Fetch 域拦截 /office-memory/snapshot 回一份合成快照，于是能验到「数据到位时的面板」
# 而不是只有空态。产物（截图 / web-verify.json）落在 .office/diag/ 下。
# 第十二轮扩了三处：令牌解析表（面板用到的 --dsw-* 逐个读运行时值）、聚焦交互
# （focus 一个节点 → 右侧详情栏）、深浅色与宽窄栅格（含容器查询是否真的塌成一列）。
# 第十三轮再扩一组**真实输入事件**：Input.dispatchMouseEvent 发滚轮 / 拖拽（缩放、平移、
# 拖节点都只有真事件才验得到，顺带证明页面没跟着滚）、点「再显示 N 条」、点「展开」、
# 点条目动作看待办指令、以及 Emulation.setEmulatedMedia 打开 prefers-reduced-motion
# 后读 .om-scroll 的计算样式。
# 注意：React 的状态更新是异步的 —— 点完不能在同一次 Runtime.evaluate 里就读 DOM。
node test/web-verify.mjs

# 测试中间产物统一落在仓库根的 .office/tmp/<suite>/ 下（不再用包内的 test/.tmp）
# 生成两组样例（默认素色 + business）与参考稿级整篇样例，并巡检底色/底纹
node test/make-samples.mjs "办公模式示例"
node test/make-reference-deck.mjs "办公模式示例"   # 12 页：封面/目录/章节/卡片/对比/流程/合并表格/图+文/KPI/时间线/图标/结束
node test/inspect-fills.mjs "办公模式示例\分月明细.xlsx" "办公模式示例\季度回顾.docx"

# 再用两个真实办公软件各验一遍（单元测试全绿也可能被 Office 判为坏文件）
cscript //nologo test\validate-com.vbs "..\..\.office\tmp\e2e" "..\..\.office\diag\com-report.txt"
node test\validate-libreoffice.mjs --dir "..\..\.office\tmp\e2e"

# Word 校验失败时先跑这个：逐步报错码，区分「Word 起不来」「Word 拒绝这个文件」
# 与「COM 被残留进程占住」。六步全 0 就说明 Word 与文件都没问题。
cscript //nologo test\probe-word-com.vbs "<绝对路径>\某文件.docx" "..\..\.office\diag\word-probe.txt"
```

`validate-com.vbs` 用本机安装的 Microsoft Office 打开产物；
`validate-libreoffice.mjs` 用部署自带的 LibreOfficeKit 把产物转成 PDF。
**两个都要过**：LibreOffice 宽容，Office 严格，只有 Office 判坏的那类问题
（非法子元素、缺必填元素、属性写成元素）才是真正会让用户打不开文件的 bug。
踩过的坑记在 [docs/ooxml-pitfalls.md](docs/ooxml-pitfalls.md)。

**校验失败先怀疑环境，再怀疑代码。** 残留的 `WINWORD`/`EXCEL` 进程会占住 COM，
表现为「Word 未能引发事件」并挂住 —— 2026-09-22 就因此把「本机 Word COM 校验
不可用」误写成了已知限制，清掉进程后同一份文件立刻通过。排查用
`probe-word-com.vbs`（它还提醒一件事：cscript 的工作目录是 `system32`，
传相对路径会让 Word 报「找不到文件」，看着像文档坏了其实只是路径问题）。

## 已知边界

- **LaTeX 编译依赖外部 TeX 发行版**（TeX Live / MiKTeX 的 latexmk 或 xelatex）。
  插件不打包 TeX：探测不到时 `office.tex.compile` 会明确报错，`create` / `read` /
  `edit` 不受影响（源文件本身就是可交付的产物）。thuthesis 需要 fontspec，
  所以只支持 xelatex / lualatex，`pdflatex` 不在候选里。
- **LaTeX 的版式由模板决定，插件不调版式**：`\vspace` / `\fontsize` 这类手工调法
  不生成、也不建议写；字体库（`fontset`）要跟着平台走（Windows 用 `windows`）。
- 论文项目里的主文件名始终是 ASCII（中文文件名在 TeX 日志、aux 与 bibtex 里会变乱码）；
  中文标题写在 `title` 字段里。附录的 `\chapter` 不占正文章号，`read()` 单独计数。
- 本科生（`degree=bachelor`）不需要 committee / comments / resolution，
  `create()` 会跳过并说明；关键词超过 5 个、`\thusetup{}` 里出现空行都会被提示。
- `office.pdf.info()` 的「有没有文本层」以前只按原始字节里的 `/Type /Font` 判，
  而 TeX 会把对象压进对象流 —— xelatex 编出来的纯文本 PDF 会被误判成扫描件。
  现在数到 0 时再用 `pdffonts` 复核一次（探测不到就维持原判），返回里也给出字体名。
- **PDF 只读**：`office.pdf` 能看能渲染，但不能生成 / 合并 / 拆分 PDF，也改不了原文件。
- **PDF 渲染依赖外部引擎**（fitz / poppler / MuPDF / Ghostscript 之一）。插件不打包
  引擎，一个都探测不到时 `office.pdf.pages()` 会明确报错；`office.pdf.info()` 不依赖
  外部程序，任何机器上都能用。单次渲染页数上限默认 20（`pdfMaxPages`），超了要分批 ——
  渲染是重活，一次几百页会把会话卡住。
- PDF 的「有没有文本层」是按 `/Type /Font` 判的**启发式**：有字体不等于一定有可抽的
  文字（比如文字被转成了曲线）。`office.pdf.text()` 返回 `chars: 0` 时照它说的走渲染。
- 加密的 PDF 会如实报 `encrypted: true` 并拒绝渲染 / 抽文本，不做解密。
- **Python 通道依赖本机的 Python 解释器**：插件不打包 Python，探测不到时
  `office.python.*` 会明确报错并给出装法与「在设置里填绝对路径」两条出路；
  Windows 上应用商店的 `python` 占位程序不算可用（探测会真跑一次并读版本号，
  占位程序过不了这一关）。装了哪些科学包要先用 `office.python.check()` 问清楚，
  缺包时插件**不会**去装（不联网、也不替用户改环境）。
- **Python 只跑、不给命令**：没有 shell、没有 stdin（`input()` 立刻 `EOFError`）、
  没有网络。脚本写工作目录里的文件不会被复检、也不进台账 —— 交付物请走
  `office.files.write`，图与中间数据放 `OUT_DIR`（缓存目录内）。
- 单次运行有超时（默认 120 秒，`python.timeoutMs`）：超时会杀掉进程并如实报
  `timedOut: true`，不留半截结论。stdout 超过 8000 字符只回前一段，完整内容在
  `logFiles.stdout` 指向的文件里。
- **音频与视频依赖四样外部件**：ffmpeg、ffprobe、SenseVoice 模型（默认取语音输入
  下载的那份）、`sherpa-onnx-node` 运行时（随语音输入装在宿主的 node_modules 里，
  按宿主安装位置解析）。缺任何一样都在调用处明确报错并给出补法，**没有静默降级**。
  推理跑在插件的子进程里，不占宿主事件循环。
- **SenseVoice 只吃规范 16 kHz 单声道 PCM16 WAV**（`fmt` 16 字节、`data` 在偏移 36）：
  SAPI 的默认输出与 ffmpeg 的默认封装都要先过一遍规范化（插件自己验、自己修）。
  能听到的格式由 ffmpeg 决定；ffmpeg 也不认的容器会明确报「解码失败」。
- **只识别，不翻译**；不认说话人（多人录音分不出谁说的）；不做流式；不做降噪与增强；
  切句是 Silero VAD 启发式，标点与断句由模型给；分块之间没有跨块上下文，边界上的
  句子可能被切成两句。识别不出内容时不编造 —— `segments` 为空就是空。
- 单次转写时长上限默认 3600 秒（`av.maxSeconds`），单块 120 秒（`av.chunkSeconds`）：
  分块解决单块大小，**不改变时长上限**，超了先切段。抽帧单次上限默认 12 张
  （`av.maxFrames`），4K 视频建议给 `width`（图小、读得快）。
- 抽出来的帧是**中间产物**（`.office/cache/av/frames/`），不进交付物；要留在工作目录
  就自己在脚本里 `office.files` 拷过去，或者直接把帧当图片嵌进 Word / PPT。
- **Excel 没有图表能力**（Excel chart 尚未实现）：图一律用 Python 画成 PNG，
  再嵌进 Word（`builder.image`）或 PPT（`deck.image` / `deck.images` / `insertImages`）。
- PPT 的**微调**是「改形状」，不是「改结构」：`office.ppt.revise` 能改文字、字体字号颜色、
  位置尺寸与文本框高度，也能**插入图片**（`addImage` / `office.ppt.insertImages`），
  但**不能增删整页、换版式、改表格结构** —— 那几件事请重新生成，或直接在 PowerPoint 里做。
- 形状几何要求带 `a:xfrm`：**图片（`p:pic`）是可以 move / resize 的**（它自带
  `p:spPr/a:xfrm`，`readSlides` 会把 `editable` 报成 `false` 只是因为图片没有文本可改，
  不代表几何改不了）。组合形状（`p:grpSp`）与表格 / 图表（`p:graphicFrame`）的几何不在
  形状自己的 `spPr` 上，改它们会进 `skipped` 并说明原因，不会假装成功。
- 同一页命中多个同名形状时全部处理并在 `warnings` 里提示；要精确到某一个请用 `readSlides` 拿到的 id。
- 图表（Excel chart、PPT chart）尚未实现，计划按 `registry.js` 加模块的方式后续补。
- PPT 自绘形状支持 32 个 preset；`flowChartData` / `roundedRectCallout` / `ovalCallout` 是口语名，
  会被映射到 DrawingML 的真名（写真名会让真实 PowerPoint 打不开文件，见 docs/ooxml-pitfalls.md）。
  连续调用 `deck.shape/line/icon` 画在同一页，用 `deck.page()` 或任一版式方法另起一页。
- 与 Excel/Word 一致：`fill` 为空串或 `FFFFFF` 视为**不填充**，纯白卡片会透明；
  要白卡片用 `F8FAFC` 这类极浅灰。
- 封面与结束页使用主题的封面底色（彩色主题是 `primaryDark` 深底白字，素色主题是白底黑字 +
  浅灰色带），可用主题里的 `cover` 覆盖项改。
- 版式助手页与自由绘制页共用一条画布通道：`deck.cards()` 之后调用 `deck.banner()/panel()/shape()`
  会续在同一页，`deck.page()` 或任一版式方法才另起一页。
- `deck.panel()` 与 `deck.banner()` 的缺省位置都在版心左上，同页使用时要显式给坐标；
  空的助手页（如 `cards({items:[]})`）在 `read()` 里会被认成自由绘制页。
- Word 公式支持 LaTeX 子集（分数、根式、上下标、大运算符、定界符、常用符号与希腊字母）；
  矩阵、多行对齐（`align` 环境）、化学式等尚未实现，遇到不认识的命令会原样保留文字并给 warning。
- Excel 的 `stats()` / `summary()` 写的是「公式 + 缓存值」：缓存值取自内存模型里已有的数字，
  公式单元格不参与统计；`sheet.formula()` 手写的表达式引擎无法求值，只写公式不写缓存值。
- Word 的目录（TOC）是域字段，但**条目与页码是预渲染好的**：打开即见完整目录、条目可点击
  跳转，不必先按 F9。页码来自引擎的分页估算（不是 Word 的精确排版），按 F9 可以让 Word
  重算成精确值。之所以不靠「打开时自动更新」：那两种标记（settings 里的 `w:updateFields`、
  域上的 `w:dirty`）都会让 Word 弹「该文档包含的域可能引用了其他文件。是否更新该文档中的
  这些域？」—— 本机真实 Word 实测，去掉任一处仍弹，两处都去掉才不弹。回归断言见
  `test/format-word.mjs`，复现手法见 [docs/ooxml-pitfalls.md](docs/ooxml-pitfalls.md) 第 8 条。
- 记忆面板（侧栏「记忆」）**本身不写记忆**：条目上的动作只生成一条精确的
  `office_memory({...})` 指令文本（复制 / 切到会话，再由用户发送），
  归档层连动作都不给。召回质量的分数是**命中词占比**，不是语义相似度 ——
  同义改写的查询命中不了。
- 图谱的缩放 / 平移 / 拖节点只改视图，不改布局；「均匀重置」是两圈同心圆的
  整齐摆法（不是力导向的另一种解），节点一多仍然受 48 个的上限约束。
- 列表分页是**对已经取回来的那一份**做渐进式展开（每页 20 条），端点上各层
  仍有 200 / 200 / 500 / 60 的硬上限，所以翻到底也看不到被端点截掉的部分。
- 三件套的排版基于测量估算（字号 × 字符宽度），不是逐像素排版引擎；
  `warnings` 是风险提示，不是精确结论。
- Excel 的 `edit()` 里 `numberFormat` 只改内存模型、没有落盘到 `xl/styles.xml`，
  会在 `skipped` 里如实报告，不会假装成功。
- 插件以 `link:` 方式装进 profile（指向本仓库目录）；移动或删除本目录会让
  已安装的插件失效。
- **改完源码必须重载 DSH 进程。** profile 按 `link:` 引入的是 ES 模块，进程启动后模块实例
  就固定在内存里：之后新增的具名导出，只有「后来才被首次加载」的模块才看得见。
  2026-09-21 实测：先给 `engine/theme.js` 加 `shadeOf`，再让 `formats/word.js` 引用它，
  结果正在跑的进程里 theme.js 还是旧实例 —— `office_help` 报
  `does not provide an export named 'shadeOf'`（word / excel 不可用，ppt 因为不引用它而照常），
  而全新进程里三个格式都能装载。改完源码请重启 DSH，再用 `office_help({ topic: 'word' })`
  确认格式装配成功。
- **`office_run` 的输出 schema 不要声明可选字段的类型。** 宿主按 `output.schema` 逐项校验，
  且声明过的字段缺键即判非法：成功路径不写 `error` 会被报成
  `"value.error" must be an object`，不传 `purpose` 会被报成 `"value.purpose" must be a string`，
  整次调用被判成「返回了非法输出」——脚本明明跑成功、文件也写出来了，反馈却拿不到。
  可选字段交给 `additionalProperties: true` 兜，别在 `properties` 里写类型。
  回归断言见 `test/smoke.mjs` 的「返回值满足宿主对声明字段的校验（声明即必填）」。