/**
 * 按需文档（office_help 的内容来源）。
 *
 * 工具声明只有两行，写法和细节都放在这里，模型需要哪一种格式再取哪一种。
 * 格式部分直接从模块的 `meta` 生成，所以新增一种格式不需要同步改文档。
 *
 * 文档分两层（第十七轮）—— 这一层是**注入预算**的落点：
 *
 *   - 默认（compact）：只给「怎么写才不错」的最小事实集 —— 操作名与签名的
 *     一览、致命的边界（没有文本层的 PDF、Excel 没有图表、office.memory 不在
 *     脚本里…）。一次投影进上下文之后，这份文档会留在**每一个后续请求的前缀**
 *     里（对话前缀只增不减），所以它越薄越好。
 *   - `detail: true`：例子、为什么这样做、配置项与迁移说明等长尾。需要时再取，
 *     不必为了问一句 API 形状就把整篇散文塞进前缀。
 *
 * 分层的判据只有一条：**模型照着默认这一层就能写出对的调用**。为此签名与参数
 * 名一个都不省 —— 省的是解释与举例，不是接口。
 *
 * 真实量级（本机实测，UTF-8 字节）：默认层与全文层相比，ppt 17906→约 2200、
 * memory 9311→约 2000、settings 3137→约 900、search 4225→约 1300、
 * pdf 2275→约 900、python 2784→约 1000、index 3083→约 1300、
 * guide 6676→约 700。
 *
 * @module dsh-office-mode/docs
 */
import { OFFICE_GUIDE } from './guide.js';
import { DEFAULT_PDF_MAX_PAGES } from './config.js';
import { PYTHON_STDOUT_CHARS } from './python.js';
import { renderThemeTable, themeCatalog } from './engine/theme.js';
import { loadAllFormats, loadFormat, formatIds } from './registry.js';

const BASE_TOPICS = {
    run: `office_run 的用法

参数
- script（必填）：一段 JavaScript。可以直接写多步，最后 return 一个值（会回显）。
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
能调外部程序的地方只有四处，都由插件自己探测、脚本里给不了命令：
- PDF 渲染（office.pdf.pages，探测 poppler / PyMuPDF / Ghostscript 等）；
- LaTeX 编译（office.tex.compile，探测 latexmk / xelatex / lualatex）；
- Python（office.python，探测 python / python3 / py）；
- ffmpeg（office.av，探测 ffmpeg / ffprobe；推理由插件的子进程跑本机 SenseVoice）。

反馈里会有什么
- 每个写出的文件都会被自动重新打开检查一遍，附上真实统计和结构摘要。
- warnings 是排版风险提示，不是错误；按提示调整通常能明显变好看。
- 脚本返回的 JSON 超过 4000 字符会被截断；长内容写进文件再读，不要靠返回值带。`,

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
    结果按「文件内容 + dpi + 格式 + 引擎」缓存在 .office/cache 里：同样的页码再渲染
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
      Word 的 builder.image(path, { widthCm, caption }) 或 PPT 的 deck.image(path)。
    - stdout 超过 ${PYTHON_STDOUT_CHARS} 字符会截断，完整输出在 logFiles.stdout 里。
    - 产物默认落在缓存目录的 python/out（即 .office/cache/python/out）。
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
- Excel 没有图表能力：图一律用 Python 画成 PNG，再嵌进 Word 或 PPT。
- 数据要交付就写成 CSV 放 OUT_DIR；要当交付物的文本仍走 office.files.write ——
  Python 直接写工作目录的文件不进台账，也不会被复检。
- 没有网络（pip install 装不了包）、没有 stdin（input() 会立刻 EOFError）、
  没有命令行（脚本里给不了别的命令）。
- 跑一次的成本是一次进程启动（几百毫秒）加你自己的计算，别把 office_run
  拆成很多次小 Python 调用。`,

    files: `office.files 的用法

- office.files.read(path)              读文本
- office.files.write(path, text)       整份覆盖写
- office.files.edit(path, edits)       批量字面替换，只写一次盘
      写法：[[旧, 新], [旧, 新]]  或  [{ find, replace, all }]
      没命中的 find 会出现在结果 missing 里，不会静默跳过
- office.files.template(path, {k: v})  把 {{k}} 一次性替换掉
- office.files.exists / list / remove / stat

改 Markdown 或长文本就用 edit/template：一次调用改完所有位置，
不要一条一条改。`,

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
- archiveKeep       归档最多保留多少个摘要文件，默认 60
- 迁移命令也写在这一页上（面板只读写配置，没有执行通道）。

「办公模式」页
工具开关 —— 关掉的工具不出现在模型面前，也不占每次请求的 schema 开销。
- office_help / office_run / office_memory
- office_search_run / office_search_brief / office_search_dispatch / office_parse_findings
开关即时生效：改完立刻重建工具面，不用重启。

检索编排
- subagentTools   检索子代理的工具白名单（默认极简七项）
- maxParallel     同时铺开几个检索子代理，1-8，默认 4
- resultLimit     解析结果最多列几条，1-40，默认 40
- maxChannels     一次提纲最多几个渠道，1-12，默认 12
- outputDir       提纲与结果文件目录，默认 .office/search
- requireCrossSource     单一来源的结论是否点名要求补检索，默认开
- fallbackOnPlatformError 平台直连失败时是否改用网页搜索兜底，默认开
- engine          auto（默认，先派子代理、跑不了用内置）/ subagent / builtin
- provider        用哪条联网通道：auto（按 providerOrder 依次试）/ 固定某一条
- providerOrder   auto 的尝试顺序
- providers.<id>  每条通道各自的 apiKey / apiKeyEnv / baseURL / model / timeoutMs
- preprocess      网页预处理：mode（article | plain | off）/ dropBoilerplate /
                  dedupeLines / minLineChars / keepTitle
- builtin         Anthropic 兼容通道（也是 providers.anthropic 的老入口）：
                  apiKeyEnv / apiKey / baseURL / model / maxUses / maxTokens /
                  maxResults / fetchPages / searchTimeoutMs / fetchTimeoutMs /
                  maxBytes / maxChars / maxRedirects

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
     条目；不会静默丢掉内容。
   - 动笔前照它办。office_help 的返回值末尾本来就会带一份热记忆投影，
     office_run 的反馈里也会带，所以平常不必特意 read。

2) 台账（layer:'ledger'）—— 每份交付物一条
   - office_run 每写出一份 Office 文档就自动登记：路径、格式、主题、字节、
     复检统计、结构摘要、这次的目的。模型不用为此多调一次工具。
   - 找「上次那份东西」用它：office_memory({ action:'read', layer:'ledger',
     query:'季度汇报' })。多个关键词用空格分开，命中越多排越前。
   - 条数超过 ledgerLimit（默认 500）时，最旧的一批滚成月度归档摘要。

3) 归档（layer:'archive'）—— 只读的长期层
   - 装的是两类东西：热记忆装不下时下沉的旧条目、台账滚动下来的旧记录。
     按月份归到一个 Markdown 摘要文件（.office/memory/archive/YYYY-MM.md），
     可以 grep、可以直接打开。
   - 检索：office_memory({ action:'read', layer:'archive', query:'...' })。
   - 归档本身也有上限（archiveKeep，默认 60 个摘要文件）：超出时最旧的摘要
     被删掉 —— 它已经是压缩过的老内容，这是唯一会真正丢东西的地方。

写入（action: add / replace / remove）
- add     新记一条：content 要一两句、自包含，下次不看上下文也读得懂。
- replace 改一条：oldText 给一段能**唯一**命中原文的片段；命中多条会被拒绝
          （宁可让你补长，也不要改错）。
- remove  删一条：同上要求唯一命中。只在用户明确要求、或确有证据说明作废时删。
- target  user = 用户是谁、偏好、要求、忌讳；project = 项目约定、环境事实、
          工具坑、可复用的做法。选错层会让「项目的事实」被当成「用户的偏好」。
- importance  critical（明确的必须 / 永远不要）/ normal / low（只对一段时间有用）。
- entities / tags  选填。entities 填了才会出现在实体视图里（实体是声明的，不猜）；
  tags 参与归档检索。

只记什么 / 不记什么
- 记：用户说过的偏好、纠正、要求「记住」的事；稳定的项目约定与环境事实；
  踩过一次、下次不想再踩的坑。
- 不记：检索结果原文（那在结果文件里）、一次性进度、自己的推测与复述、
  从记忆里刚读出来的东西（读不是写）。
- 记完给一句说明就行，不要把条目整段念给用户听；记忆是静默生效的。

手工登记（action:'log'）
office_run 写出的 Office 文档已经自动登记，正常不用管。只有它覆盖不到的
产物（例如用户放进工作目录的外部文件）才手工补一条：office_memory({ action:
'log', path:'参考稿.pptx', format:'ppt', purpose:'作为版式参考' })。

记忆范围与层开关（设置页可改）
- 范围 scope：workspace = 只跟当前工作目录（一个项目一份，默认）；global = 所有
  项目共用一份；both = 两者并存 —— 读的时候**全局在前、工作区在后**，因为具体
  项目的约定比通用偏好更该被最后看到。每条结果都带 origin（全局 / 工作区）。
- 全局层默认在 $DSH_HOME/.office/memory（globalDir 可改；相对路径相对用户主目录，
  不相对工作目录 —— 全局层跟着工作目录走就没有意义了）。
- userScope：memory = 「用户偏好」跟着上面的范围；global = 始终落全局层，
  这样换一个工作目录不必把「用户是谁」重记一遍。
- layers.hot / ledger / archive：每层独立开关。关掉只是不再读 / 不再写，
  **已有数据不会被删除**，重新打开就回来。

图关系（action: link / unlink / related / entities）
- link 给两条记忆建一条**双向**类型化关系：{ action:'link', sourceId, targetId,
  kind }。kind：related（默认）/ refines / supersedes / contradicts / supports /
  derives。两个 id 都必须真实存在，否则报错 —— 允许悬空边会让关系图慢慢烂掉，
  而悬空边在遍历时表现为「静默少一条」，比直接报错难查得多。
- id 从哪儿来：read 的结果里热记忆与归档条目都带 id（热记忆下沉进归档时
  **保留 id**，所以关系不会因为下沉而断）。
- related 沿关系找相邻条目：{ action:'related', id, depth?:1-3, kind? }。
  默认只走一跳 —— 图一旦放开很容易把整个记忆库拉进上下文。
- entities 看实体视图：实体来自条目**显式声明**的 entities 数组（add / replace
  时可以传），不做抽取。猜出来的实体比没有实体更误导。

召回质量（strict-v1）
- 带 query 的检索会分档：按「命中了查询里几成的词」算相关度（**不是**向量余弦
  相似度），达到高阈值的直接采纳，中档与未知档各有一个名额上限，低于下限的丢掉。
- 先取「条数上限 × candidateMultiplier」个候选再分档：直接截断的话，低分结果会把
  高分结果挤出候选池。被丢掉多少条会在结果里报出来。
- 想让它别丢结果：把设置页的 policy 改成 off。

每回合配额
- 一个回合里第一次带 query 的检索、后续的换词细化、图关系遍历各有次数上限
  （默认 1 / 1 / 1）。回合号是从会话日志的 turn/start 事件读出来的。
- 用完会**明确拒绝并说明原因**，不要以为是检索坏了然后换个工具重复同样的检索。
  不带 query 的 read 是看现状、不是查资料，不占配额。0 = 不限制。

备份与迁移（action: export / import）
- export 把热记忆 + 台账 + 归档（索引与 Markdown 全文）+ 关系整包写成一个 JSON
  单文件：{ action:'export', packPath?:'...' }（不传就写到记忆目录下的
  pack-<时间戳>.json）。
- import 读回来：{ action:'import', packPath:'...' }。口径是**只增不改**，按 id
  幂等：热记忆 / 台账 / 关系同 id（或同 source+target+kind）已存在就跳过，不覆盖
  现内容；归档按月合并、按块 id 去重。所以重复导入同一份包不会产生第二份。
- 这份 Pack 含私有记忆，不要当公开文件传。

浏览面板
- 设置对话框左侧「记忆系统」是配置页；侧栏的「记忆」入口是**浏览**页：热记忆 /
  台账 / 归档 / 关系 / 实体五个页签，带搜索与工作目录切换。它走一个只读的本地
  端点（/office-memory/snapshot），且只服务「本进程真的跑过办公工具的工作目录」。

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
  与「有多少条关系边」都随条目保留。**图边本身不迁**（mnemon 的图在它自己的库里），
  只保留计数；迁完之后可以用 action:'link' 按需要重建关系。
  已软删（deleted_at 非空）的不迁。
- Documents（项目档案）→ 办公记忆没有对应层：只报告数量并提醒原件仍在原地。
- 幂等：热记忆按内容去重、归档按 mnemon 的原始 id 去重，重复跑不会产生第二份；
  只做 add，不 replace / remove，不会动你后来改过的条目。
- 读长期记忆库要用 Node 自带的 node:sqlite（Node 22.5+）。拿不到时只迁 runtime
  部分并明确报警，不会整次失败。
命令行入口（改完源码不用重启就能迁移）：
  node scripts/migrate-mnemon.mjs --dry-run
  node scripts/migrate-mnemon.mjs`,

    search: `检索与结果解析（office_search_run / office_search_brief / office_search_dispatch / office_parse_findings）

办公模式自带联网检索（不用装别的插件）
- preset 里没有 web_search / web_fetch 这类工具，所以要查资料一律走本插件
  自己的联网通道。通道有多条，按顺序自动挑或者由你点名：
  seam=宿主 web 服务（首选）/ anthropic=Anthropic 兼容 + 原生 web_search（默认 DeepSeek 官方）/
  openai=OpenAI 兼容 / tavily / brave / bocha=博查（中文长尾好）/ exa / serper=Google 结果代理 /
  searxng=自建实例 / duckduckgo=免 Key 兜底。
  auto（默认）按 providerOrder 依次试，第一条成功的就用它；没配 Key 的通道当场失败（不联网），
  所以顺序长也不拖时间。office_search_run({ provider:'bocha' }) 可以按次点名。
- 只查一两个事实点：office_search_run({ queries: [...] })。一轮直查，
  来源与摘录写进 .office/search/<slug>/run.md，聊天里只回一份紧凑清单。
- 取回来的网页**先过预处理**再进上下文（去脚本样式、去导航与样板、挑主容器、丢样板行、去重），
  结果文件与反馈里会给出「几页、原文多少字符 → 洗完多少、丢了多少行」。判错时用
  preprocess:'off' / 'plain' 退回原样（设置页里的 search.preprocess.mode 是默认值）。
- 要按渠道逐项覆盖（下一条的分流规则）：走下面三步走。

三步走，顺序不要颠倒
1. office_search_brief({ topic, type?, audience? })
   出提纲：按内容类型决定该搜哪些渠道，并把提纲写成文件。
   type 可选 hotspot（热点）/ knowledge（知识）/ manual（手册）/ mixed；
   省略时按主题文字自动判断。
2. office_search_dispatch({ briefPath, outputPaths })
   按提纲逐渠道执行。组合里有子代理与联网检索工具时派子代理（每个渠道一个，
   各自调 web_search / advanced_search / platform_search）；没有时自动改用
   插件的内置检索通道，在进程内跑同样的渠道。两条通道写出的结果文件同一种
   格式，所以第三步一样用。
   材料都只落盘，主上下文不会被搜索结果灌满。
3. office_parse_findings({ paths, type?, topic? })
   把结果文件读成摘要：按渠道归类、算来源覆盖度、标出只有单一来源背书
   与没带 URL 的条目。摘要是几十行，不是整页原文。

渠道分流规则（固定，不靠临场发挥）
- 热点事件：权威媒体（定事实骨架）+ 各社交平台（现场与争议）；两边都要。
- 知识类：百科拿定义与体系；需要更深再下沉到文献与一手资料。
- 手册类：只认官方平台的参考文档；博客与问答只能用来看懂报错。
- 任何类型都先做一轮不限定来源的泛搜，用来找准关键词与分歧点。

为什么这样做
- 茧房效应的来源是只用一个渠道就把结论定下来。规则写死在插件里，
  每次检索都会被要求覆盖多个互不相关的来源。
- 结果落盘：一次检索的原始材料不进主上下文，主会话留着余地做排版与措辞；
  文件留在工作目录里也可以复核与追查出处。
- 内置通道只搬来源与摘录，不会自己换角度追查，也不判「这条说法该不该信」；
  要更强的覆盖就开子代理（组合里装 web 工具）或按提纲手工补查。

子代理的工具面（极简，只有这七个）
- 检索：web_search / advanced_search / platform_search
- 取正文：web_fetch
- 读：read（文本）、read_image（图表 / 截图 / 扫描件）
- 写：write（唯一交付方式）
刻意不给：edit / glob / grep（只写自己的结果文件，不用改文件或找文件）、
office_*（不生成文档）、bash / pwsh（不执行命令）、子代理与任务板（不派活）。
裁剪靠 toolFilter.allow 白名单（子代理默认继承父方组装），
用白名单而不是黑名单：以后新装的工具也不会悄悄漏给子代理。

内置检索怎么配（设置页「办公模式 → 检索编排」）
- search.provider     auto（默认，按顺序自动挑）/ 固定成某一条通道 id
- search.providerOrder  auto 的尝试顺序（接缝在最前、零配置通道在最后）
- search.providers.<id> 每条通道各自的 apiKey / apiKeyEnv / baseURL / model / timeoutMs。
                       anthropic 与老的 search.builtin 是同一件事的两个入口，
                       同名键以 providers.anthropic 为准
- search.preprocess   mode（article | plain | off）/ dropBoilerplate / dedupeLines /
                       minLineChars / keepTitle
- search.builtin      Anthropic 通道的细项：apiKeyEnv（默认 DEEPSEEK_API_KEY，
                       先问宿主凭据服务、再看环境变量、最后看 $DSH_HOME/.credentials.yaml）/
                       baseURL / model（要支持原生 web_search）/ maxUses / maxTokens /
                       maxResults / fetchPages / 各类超时与体积上限
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
- **画面 → 图片**：视频按时间点抽帧成 JPEG / PNG，路径交给 read_image 看 —— 与
  PDF 图像型那条路一样，是读图能力，不是 OCR。

先问一句：await office.av.check()
    返回 available / missing / ffmpeg / ffprobe / sensevoice（模型、VAD、运行时、
    精度、线程、语言）/ chunkSeconds / maxSeconds / maxFrames / hint。
    **四样能力缺任何一样都直接报错**（ffmpeg、ffprobe、模型、sherpa-onnx 运行时）——
    不静默降级、不假装转写出了空文字。hint 里写了缺什么、怎么补。

看清楚是什么：await office.av.info(path)
    → { kind:'audio'|'video', container, durationSeconds, hasAudio, hasVideo,
        audio:[{codec,sampleRate,channels}], video:[{codec,width,height,fps,frames}], hint }
    视频里没有音频轨、音频里没有视频轨，这里一眼就能看出来，别盲调后面的方法。

转写：await office.av.transcribe(path, { language?, chunkSeconds?, maxSeconds?, out?, words? })
    → { text, segments:[{index,start,end,clock,text,lang,emotion,event}], audioSeconds,
        speechSeconds, inferenceSeconds, realtimeFactor, chunks, model,
        detectedLanguages, audio:{path,reused} }
    - segments 是逐句结果：时间戳可直接用来做字幕、回听定位与逐字稿；
      lang / emotion / event 是 SenseVoice 的富文本输出（<|zh|> / <|NEUTRAL|> / <|Speech|>）。
    - out 给出时把带时间戳的 Markdown 逐字稿写进那个文件（**长稿一定要给 out**：
      工具反馈里的脚本返回值会被截断，写到文件里再 read 才拿得到全文）。
    - chunkSeconds 默认 120 秒（官方那条链一次 131 秒封顶，这里留了余量）；
      maxSeconds 是单次时长上限（默认 3600 秒），超过就报错要求先切段。
    - 解码出来的规范 WAV 按内容键缓存在 .office/cache/av/audio 下：同一份文件
      再转一次是 audio.reused=true，不重解码。

抽帧：await office.av.frames(path, { at?, every?, count?, from?, to?, format?, width?, outDir? })
    → { dir, count, format, files:[{index, at, clock, path, bytes}], hint }
    - 时间点优先级：at（明确给秒，数组或单值）> every（定步长秒）> count（默认 6 张均匀铺开）。
    - 拿到 path 之后**逐张 read_image({ file_path })**。图上有字也认得出来，那是读图能力。
    - 单次上限 maxFrames（默认 12 张）；4K 视频建议给 width（如 1280），图小、读得快。
    - 抽出来的帧落在 .office/cache/av/frames/<文件名>/ 下，是中间产物，不进交付物。

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
- 长音频靠分块 + VAD，**没有**跨块上下文：块边界处的一句话可能被切成两句。
- 设置页只列常用项；VAD 阈值、单句上限、词级时间戳（words）、缓存子目录
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

- script（必填）：一段 JavaScript，可以写多步，最后 return 一个值（会回显）。
- purpose（选填）：一句话说明这次做什么。
- keepCache（选填）：false 结束即清空缓存目录（默认保留，保留才能跨调用复用）。

脚本里能用：office.word / office.excel / office.ppt（造、读、改文档）、office.tex
（学位论文，见 tex）、office.pdf（见 pdf）、office.python（见 python）、office.av
（音频转写与视频抽帧，见 av）、office.files、office.cache、office.theme、
office.log / warn / note、office.assert(条件, 说明)。

没有：process / require / fs / 网络 / 计时器；也没有 office.memory（记忆走
office_memory 工具）。路径相对当前工作目录。能调外部程序的只有 PDF 渲染、
LaTeX 编译、Python、ffmpeg（office.av）四处，都由插件自己探测，脚本里给不了命令。
反馈：每个写出的文件都会被重新打开复检（真实统计 + warnings）；脚本返回的 JSON
超过 4000 字符会截断，长内容写进文件再读。`,
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
    Word 的 builder.image(path, …) 或 PPT 的 deck.image(path)。
    stdout 超过 ${PYTHON_STDOUT_CHARS} 字符会截断，完整输出在 logFiles.stdout。
- await office.python.file('analysis.py', { args? })   长脚本先落成 .py 再跑。

脚本里多出三个名字：OUT_DIR（产物目录，图必须 savefig 到这里才会出现在 files 里）、
WORK_DIR（当前工作目录）、use_cjk_font()（matplotlib 中文与负号）。
一定要 savefig，不要 plt.show（没有显示器，后端是 Agg）。
没有网络（pip install 装不了包）、没有 stdin、没有命令行。
细节（返回字段、边界清单、绘图示例）用 office_help({ topic: 'python', detail: true })。`,
    av: `音频与视频的内容提取（office.av）—— 声音转文字、视频抽帧

先问一句：await office.av.check() → { available, missing, ffmpeg, ffprobe,
    sensevoice, chunkSeconds, maxSeconds, maxFrames, hint }。
    ffmpeg / ffprobe / SenseVoice 模型 / sherpa-onnx 运行时缺任何一样都**直接报错**，
    不静默降级；hint 里写了缺什么、怎么补。

- await office.av.info(path) → { kind, container, durationSeconds, hasAudio, hasVideo,
    audio:[{codec,sampleRate,channels}], video:[{codec,width,height,fps,frames}] }
- await office.av.transcribe(path, { language?, chunkSeconds?, maxSeconds?, out?, words? })
    → { text, segments:[{start,end,clock,text,lang,emotion,event}], audioSeconds,
        speechSeconds, inferenceSeconds, realtimeFactor, chunks, detectedLanguages,
        audio:{path,reused} }
    language: auto（默认）/ zh / en / yue / ja / ko；chunkSeconds 默认 120；
    out 给出时把带时间戳的 Markdown 逐字稿写进文件（**长稿必须给 out**，
    脚本返回值会被反馈截断）；同一份文件再转是 audio.reused=true。
- await office.av.frames(path, { at?, every?, count?, from?, to?, format?, width? })
    → { dir, count, files:[{at, clock, path, bytes}] }
    时间点：at > every > count（默认 6 张均匀铺开）；单次上限 maxFrames（默认 12）。
    每张 path 交给 read_image —— 图上文字是读图能力认出来的，不是 OCR。
- await office.av.extract(path, { ...转写参数, ...抽帧参数, out?, frames?, transcript? })
    → { info, transcript, frames, notes, next } —— 会议视频一条调用拿到逐字稿 + 画面。

边界：只识别不翻译；不认说话人；不流式；不做降噪；切句是 VAD 启发式；
分块之间没有跨块上下文（边界上的句子可能被切成两句）。
细节（返回字段全表、典型流程、配置项）用 office_help({ topic: 'av', detail: true })。`,
    files: `office.files 的用法（改 Markdown 或长文本就用它，一次改完所有位置）

- office.files.read(path)               读文本
- office.files.write(path, text)        整份覆盖写
- office.files.edit(path, edits)        批量字面替换，只写一次盘
      写法：[[旧, 新], [旧, 新]] 或 [{ find, replace, all }]；没命中的 find 在结果
      missing 里报出来，不会静默跳过。
- office.files.template(path, {k: v})   把 {{k}} 一次性替换掉
- office.files.exists / list / remove / stat`,
    cache: `office.cache 的用法（中间产物目录，跨调用保留）

- office.cache.dir() / write(n, 数据) / list() / stats() / clear()
清理规则：每次 office_run 开始时只收走超过 TTL（默认 12 小时）没碰过的条目；
office_run({ keepCache:false }) 或 office.cache.clear() 才整目录清空。
渲染好的 PDF 页面图放这里，下一步 read_image 直接读；交付物写工作目录，不要放这里。`,
    settings: `设置界面（Settings > 插件 > 办公模式 / 记忆系统），同一个命名空间 dsh-office-mode

「办公模式」页
- 工具开关：office_help / office_run / office_memory / office_search_run /
  office_search_brief / office_search_dispatch / office_parse_findings
  关掉的工具不出现在模型面前，也不占每次请求的 schema 开销；改完即时生效。
- 检索编排：subagentTools（子代理白名单）/ maxParallel 默认 4 / resultLimit 默认 40 /
  maxChannels / outputDir 默认 .office/search / requireCrossSource /
  fallbackOnPlatformError / engine（auto|subagent|builtin）/
  provider（用哪条联网通道）/ providerOrder / providers.<id>（各通道的 Key 与端点）/
  preprocess（网页预处理强度）/ builtin.*（Anthropic 兼容通道的细项）
- 文档与缓存：defaultTheme / scriptTimeoutMs / maxScriptChars / cacheDir / keepCache /
  cacheTtlMinutes 默认 720 / pdfDpi / pdfMaxPages / pdfEngine /
  texEngine / texTimeoutMs / texTemplateDir

「记忆系统」页
- enabled / dir（默认 .office/memory）/ autoLedger / promptHint /
  userLimitBytes 默认 4096 / projectLimitBytes 默认 10240 / ledgerLimit 默认 500 /
  archiveKeep 默认 60（归档摘要文件数）。

所有数值参数都有上下界：设置页挡一道、运行期再收敛一次，写坏配置不会让插件崩。
宿主没有设置能力时这两页自动不出现，其余功能不受影响。
完整清单（每个参数的含义与默认值）用 office_help({ topic:'settings', detail:true })。`,
    memory: `办公记忆（office_memory）—— 三层，跨会话保留在会话工作目录的 .office/memory/

- 热记忆（layer:'hot'）：用户偏好与项目约定，动笔前照它办。真源是 memory.json；
  USER.md / MEMORY.md 是它的投影，只能经 office_memory 改（直接编辑会被覆盖）。
- 台账（layer:'ledger'）：office_run 每写出一份 Office 文档自动登记一条
  （路径 / 格式 / 主题 / 字节 / 复检统计 / 用途），不用手工补。
- 归档（layer:'archive'，只读）：热记忆装不下时下沉的旧条目与滚动下来的旧台账，
  按月份装订成 .office/memory/archive/YYYY-MM.md。

读：office_memory({ action:'read', layer:'hot'|'ledger'|'archive'|'all', query? })。
    带 query 是检索（按命中词占比分档，占每回合配额）；不带 query 是看现状。
写：add / replace / remove（target: user = 用户偏好与要求，project = 项目约定与环境）。
    replace 与 remove 的 oldText 要给一段能**唯一**命中原文的片段，命中多条会被拒绝。
    importance：critical（明确的必须 / 永远不要）/ normal / low。
只记：用户说过的偏好、纠正、要求「记住」的事；稳定的项目约定与环境事实；踩过的坑。
不记：检索结果原文、一次性进度、自己的推测与复述、刚从记忆里读出来的东西。
记完一句说明即可，不要把条目整段念给用户听。
其他动作：log（手工登记外部产物）/ link·unlink·related（图关系）/ entities（实体视图）/
    status / export·import（Pack 备份，只增不改）/ migrate（从 mnemon 迁移）。
记忆按工作目录存（一个项目一份）—— 运行时的记忆目录默认为会话工作目录下的
.office/memory（设置页可改）。跨项目共享在设置页把 scope 改成 both / global。
细节（范围与层开关、召回质量、每回合配额、面板、与 mnemon 的差别、迁移映射）用
office_help({ topic:'memory', detail: true })。`,
    search: `检索（office_search_run / office_search_brief / office_search_dispatch / office_parse_findings）

办公模式自带联网检索通道（preset 里没有 web_search / web_fetch 这类工具，查资料一律走它）：
多条通道，auto 按顺序自动挑（宿主 web 服务 → Anthropic 兼容 → 三方检索 API → 免 Key 兜底），
也可以用 provider 参数点名（bocha 中文好 / serper 就是 Google / duckduckgo 免 Key …）。
取回来的网页先过预处理（去导航与样板），preprocess 参数可按次改强度。

- 只查一两个事实点：office_search_run({ queries: [...] }) —— 一轮直查，来源与摘录
  落 .office/search/，聊天里只回紧凑清单；写进文档前按 URL 核对。
- 要按渠道逐项覆盖：三步走，顺序不要颠倒
  1. office_search_brief({ topic, type?, audience? }) 出提纲并落盘
     type：hotspot（热点）/ knowledge（知识）/ manual（手册）/ mixed（判不准）
  2. office_search_dispatch({ briefPath, outputPaths }) 逐渠道执行
     有子代理**且工具面里真有那四个联网工具**才派子代理；否则自动走内置通道，并在反馈里
     说一次缺哪些工具（那是装配期就知道的事实，不是错误 —— 不用绕路去派 spawn_teammate）。
     结果文件格式一样。材料只落盘，不进主上下文。
  3. office_parse_findings({ paths, type?, topic? }) 读成摘要（按渠道归类、算来源
     覆盖度、标出单一来源与没带 URL 的条目）
- 分流规则（固定）：热点 = 权威媒体 + 社交平台；知识 = 百科，需要更深下沉到文献；
  手册 = 只认官方文档；**任何类型都先泛搜一轮**找准关键词与分歧。
- 茧房效应就来自只用一个渠道定结论 —— 规则写死在插件里，不要跳过渠道覆盖。
- 通道报错时看它给的**分类**，四类修法不同：**配置缺失**（去设置页配 Key 与端点）；
  **网络出口不可达**（连不上/超时/5xx —— 换时间或换出口再试）；**目标站拒绝**
  （401/403/429/451、区域封锁页、人机校验页、页面不存在 —— **换一个来源，别重试同一个
  地址**）；**没拿到结果**（换关键词或换角度）。取正文失败会按分类计数报出来。
- **HTTP 200 也可能是失败**：区域封锁页、拦截页、登录墙常常连状态码都是 200。
  插件已做内容级哨兵，被判失败的那一条不要当正文用，也不要拿它当「这个站点没有内容」。
细节（通道实现、子代理白名单、内置检索参数、够了的标准）用
office_help({ topic:'search', detail: true })。`,
};

/** 默认层的索引短语（全文层那份在 indexText() 里）。 */
const COMPACT_INDEX = `办公模式能力索引（默认只给一行摘要；要细节传 detail:true）

- word：Word 文档（.docx）—— 生成 / 读取 / 批量改写：封面、标题层级、列表、引用、表格、图片、公式、页码、目录
- excel：Excel 工作簿（.xlsx）—— 表格、公式、统计、条件格式（没有图表能力）
- ppt：演示文稿（.pptx）—— 版式助手、图片、生成后审阅与就地微调
- tex：LaTeX 学位论文（thuthesis）—— 生成源码并编译 PDF
- pdf：PDF（.pdf，只读）—— 先 info 判断文本层，文本型抽文字、图像型渲染成页图再 read_image
- python：Python 计算与绘图—— 科学计算、画图，图直接嵌进 Word / PPT
- av：音频与视频的内容提取—— 录音 / 视频转文字（带时间戳），视频按时间点抽帧给 read_image
- run：office_run 批处理脚本怎么用
- files：文件读写与批量替换
- cache：中间产物目录与清理规则
- memory：三层记忆（热记忆 / 台账 / 归档）怎么读写、记什么不记什么
- search：联网检索（内置通道 / 三步走 / 渠道分流规则 / 结果解析）
- settings：设置界面（工具开关与参数）
- theme：视觉主题（默认 plain 素色，无背景填充）
- guide：说话与做事的方式（第一次进入办公模式建议读一遍）

查详情：office_help({ topic: 'word' }) 这样传话题名；要看例子与全部解释再加 detail:true。
传错话题时这份索引会原样退回，一次调用就能自我纠正。`;

/** 默认层的 guide（全文层是 OFFICE_GUIDE，约 6.7 KB）。 */
const COMPACT_GUIDE = `办公模式的工作方式（默认层；完整版 office_help({ topic:'guide', detail:true })）

说话：直接说结果，先说做了什么、文件在哪；短句，少术语；不寒暄、不复述用户的话；
内容一多就单独建文件，聊天里只留结论与文件位置。
动手：一次工具调用完成一批操作（office_run 一次写完），不要分很多次试探；
改已有文件先读回来、一次改完写回；交付物写工作目录，中间产物写 .office/cache。
文档：先 office_help 查写法，再用 office_run 一次写成；不传主题就是 plain（素色无底纹），
要配色显式选主题；生成后看反馈里的 stats / outline / warnings，有 warning 就按提示调。
PPT 多走一步：生成 → office.ppt.read 审阅（版式与内容量是否匹配）→ office.ppt.revise
就地微调（改文字 / 字号 / 位置）→ 再 read 复检一次。一轮就够，不要整篇重生成。
PDF：先 office.pdf.info 看有没有文本层，文本型抽文字、图像型渲染成页图再 read_image。
Python：算数、统计、画图用 office.python.run（先 office.python.check() 看装了哪些包）；
长脚本先 office.files.write 落成 .py 再 office.python.file 跑；Excel 没有图表能力。
音频与视频：先 office.av.check() 看 ffmpeg 与本机 SenseVoice 在不在，再用
office.av.extract（视频：逐字稿 + 抽帧）或 office.av.transcribe（只要文字）；
长逐字稿给 out 落成文件再 read，抽出来的帧逐张 read_image。
检索：一两个事实点走 office_search_run；要渠道覆盖走 brief → dispatch → parse_findings。
记忆：热记忆动笔前照办、台账由 office_run 自动记、归档按需检索；office_help 末尾会带
一份热记忆投影，所以平常不用特意 read；用户说明偏好或纠正你时当场记一条。
报告时给的文件路径要原样引用，方便用户直接打开。`;

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
- tex：LaTeX 学位论文（thuthesis）怎么生成与编译（office.tex）
- pdf：PDF 怎么读（文本型抽文字 / 图像型渲染成图片再读图）
- python：Python 计算与绘图（科学计算、画图，图直接嵌进 Word / PPT）
- av：音频与视频的内容提取（录音 / 视频 → 带时间戳的文字；视频 → 抽帧给 read_image）
- files：文件读写与批量替换
- cache：中间产物与清理规则（跨调用保留、可命中复用）
- memory：三层办公记忆（热记忆 / 台账 / 归档）怎么读写、记什么不记什么，
  以及记忆范围、图关系、召回质量、每回合配额、备份迁移与浏览面板
- theme：视觉主题（默认 plain 素色，无背景填充）
- search：联网检索（内置通道 / 三步走 / 渠道分流规则 / 结果解析）
- settings：设置界面（工具开关与参数调节）

查详情：office_help({ topic: 'word' }) 这样传话题名。`;
}

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
            text: `没有「${wanted}」这个话题。可用：guide / run / tex / pdf / python / av / files / cache / theme / ${formatIds().join(' / ')}。\n\n${await indexText(detail)}`,
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