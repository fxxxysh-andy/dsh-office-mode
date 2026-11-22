# dsh-office-mode（办公模式）

给 DeepSeek Harness 的办公插件：**九个工具声明**，一次调用干完一批活，
自带 Word / Excel / PPT 的生成、读取、编辑、排版与主题，零第三方依赖；
另有 LaTeX 学位论文（thuthesis 模板）的源文件生成与 PDF 编译，
以及文档互转与统一出 PDF（`office.convert`：md / docx / tex 互相转，认识的文档格式都能出 PDF）；
带一条按内容类型分流的检索链路，**联网通道也自带**（免 Key 的网页抓取优先：
抓 DuckDuckGo HTML 解析，自建 SearXNG 随后，宿主 web 服务兜底；
可选出口代理，取正文撞上 PDF 交给 `office.pdf` 抽文本），结果落盘、主上下文只收摘要；
办公 preset **不声明**宿主的 `tool-web`（第三十轮起）：资料搜索只走插件自己的抓取通道；
还带一套按工作目录存的**三层记忆**（热记忆 / 台账 / 归档，参考 mnemon 的分层做法）
与一棵**知识库**（`kb`：外部原文的结构化切块，块带块头与字符区间，按来源三档设防；
manifest 是真源，SQLite 提供 FTS / 向量 / 块边 / 摘要索引）。

## 它解决什么

| 问题 | 做法 |
| --- | --- |
| 工具太多、每次都要挑 | 只注册 `office_help` / `office_run` / `office_memory` / `office_web_search` / `office_web_fetch` 五个常用工具 |
| 写法和细节占满上下文 | 用 `office_help` 按需取，不查不给 |
| 改一处文字要来回好几次 | `office_run` 里一次批量替换，只写一次盘 |
| 生成的文档"能用但难看" | 主题模板 + 排版测量，写盘前就报溢出风险 |
| 默认就被上了底色 | 不传主题 = `plain` 素色网格：Excel 是普通表格的网格形态，Word 是白纸黑字，零背景填充 |
| 不知道到底生成了什么 | 写出的文件自动重新打开解析，回真实统计与结构 |
| 满目录临时文件 | 中间产物只进 `.office/runtime/cache/`，跨调用保留、按 12 小时自动清理 |
| 用户纠正过一次，下次又犯 | 三层记忆：偏好与约定（热记忆）、交付物登记（台账）、下沉的旧条目（归档），跨会话留在 `.office/memory/` |
| 忘了"上次那份季度汇报"叫什么、用什么主题 | `office_run` 每写出一份文档就自动登记台账；`office_memory` 按关键词检索 |
| 从文档或检索结果里拿到的原文，下次还要重新读一遍 | 知识库 `kb`：显式入库成块（带块头与字符区间），`kb-search` 走 SQLite 词法 BM25（中文 bigram，零重叠弃答）；配了 embedding 时升级成向量 + 词法混合检索，需要时再由独立大模型筛选候选（只回编号，不重排）。按需有界整份读；引用落到 `kb:<哈希>:<块号>`，台账与记忆都能反着查 |
| 单来源的说法被当成事实记下来 | 写入期硬规则：来源分三档（`user` / `verified` / `unverified`），**单源未核实的内容永远进不了热记忆**（引用它的知识块同样被拒） |
| PDF 读不了（扫描件 / 手写笔记根本没有文本层） | `office.pdf`：先 `info` 判断类型，文本型抽文字，图像型渲染成页面图交给 `read_image` |
| 学位论文的版式要求严，手工排版过不了审查 | `office.tex`：按 thuthesis 模板生成整篇项目（主文件 + thusetup + data/ 分章 + refs.bib + 模板文件）并调用本机 TeX 编出 PDF —— 版式交给模板，源文件的结构、字段、引用对不对由插件保证 |
| 手上的文档要换一种格式（tex / md / word），或不管什么格式都要一份 PDF | `office.convert`：md ↔ docx ↔ tex 互转（pandoc），`.tex` 走本机 TeX、Office 三件套走宿主 LibreOffice、纯文本源两条路自动挑；引擎在运行期探测，缺哪样只影响它那几条路 |
| 检索只用一个来源，结论被单一渠道钉死 | 按内容类型分流：热点走权威媒体 + 社交平台，知识走百科后下沉文献，手册只认官方文档；每类都先泛搜 |
| 检索要配 Key、被挡就没辙 | 插件自带联网通道：免 Key 的网页抓取在前（DuckDuckGo HTML 解析 → 自建 SearXNG），宿主 web 服务兜底；可配出口代理，来源是 PDF 时交给 `office.pdf` 抽文本；`office_search_dispatch` 在子代理跑不了时自动改用它 |
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

两份东西**要一起装**：办公模式的 persona 是 `complete: true` 的唯一系统提示段，
插件注册的提示段落（guide、记忆说明）在这个模式里进不去，所以 persona 用
`{{office_capabilities}}` 这个变量接收插件现算的「本会话能力」（引擎入口、检索派工走哪条路）。
宿主只认注册过的变量，因此**缺插件时装配会明确报
`unknown prompt variable "{{office_capabilities}}"`**（不是静默少一段）——
两个包本来就成对使用，这条把它写死成硬依赖。变量名两边各写一处：
插件在 `packages/dsh-office-mode/src/capabilities.js` 的 `CAPABILITY_VARIABLE`，
preset 在 persona 正文里；`node test/preset-check.mjs` 会双向核对（插件侧改名 →
preset-check 红；preset 引用了没人注册的变量 → 同样红），`node test/capabilities.mjs`
则负责「插件确实注册了这一个变量」那一半。**别把插件停用而留着 preset**：
那会让办公会话每一轮都在装配时失败（preset-check 有一条断言守这个启用态）。

办公模式与检索子代理两套组装都**不吃 mnemon（记忆）**：整套 mnemon 组件已在 profile
里停用（不再出现在 `dsh.profile.bundles` 里），preset 自第三十轮起也不再为它写
禁用行（拦一个没挂载的东西是死配置）。`node test/preset-check.mjs` 仍会拿 profile
的真实行清单核对 mnemon 的生效状态——谁把 bundle 重新启用而没补禁用行，那套测试会当场红。

关掉 mnemon 不等于办公场景没有记忆：本插件自带一套**按工作目录**的三层记忆
（见下文「三层记忆」）。两者的分工是清楚的——mnemon 是跨项目的持久记忆体与
图关系检索，本插件管的是「这份稿子的约定」与「这个目录里产出过什么」，
按目录隔离反而更准（不同项目的约定经常互相矛盾）。

## 九个工具

**`office_help({ topic })`** —— 按需文档。`topic` 取 `guide` / `run` / `word` / `excel` / `ppt` /
`tex` / `pdf` / `python` / `av` / `preview` / `image` / `archive` / `files` / `cache` / `theme` /
`memory` / `search` / `settings`；省略返回索引。这份清单从真源现取
（`docs.js` 的 `availableHelpTopics()`），工具描述与「没有这个话题」的回执都跟着它走，
所以新话题不会漏在文档里（第四十七轮 18-23：清单以前手写两遍，一处重复了 `tex`、
另一处少了四个话题）。
它的返回值**末尾会附一段记忆投影**（热记忆 + 最近台账；没变化时不重复贴），所以模型
动笔前那一次调用就能看到约定，不必额外读一次记忆。

**`office_memory({ action, ... })`** —— 三层记忆的读写与检索，见下文「三层记忆」。

**`office_web_search({ queries })`** —— 查一个事实点（网页抓取通道，免 Key）：
给 1–5 条查询，抓 DuckDuckGo HTML 结果页解析（配了自建 SearXNG 也会用），
来源清单直接回到对话，不落盘、不起子代理。

**`office_web_fetch({ url })`** —— 打开一个页面取正文（同一条抓取通道）：
自带安全抓取（公网地址校验、只跟同源跳转、限长限时），取回的正文先过预处理；
来源是 PDF 时交给 `office.pdf` 抽文本。

**`office_search_run({ queries, ... })`** —— 内置检索直查：给 1–5 条查询，
来源与摘录写进 `.office/search/<slug>/run.md`，聊天里只回一份清单。见下文「检索」。

**`office_search_dispatch` / `office_parse_findings`**
—— 按渠道的覆盖检索（出提纲并执行 → 解析回摘要），见下文「检索」一节。

**`office_run({ script | scriptFile, args?, purpose?, keepCache? })`** —— 一次调用执行一整段脚本。
`script` 与 `scriptFile` 二选一：长脚本先 `office.files.write('build.js', code)` 落成一份 `.js`，
之后每次 `office_run({ scriptFile: 'build.js', args: { month: '3 月' } })` 复用同一段构建器
（脚本里以全局 `args` 读到它）—— 不必把几百行内联进工具参数，改动也只改那一份文件。
错误里的行号是那份脚本文件里的行号；脚本文件本身会记进台账的 `source`。

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

### 文件通道：文本走 `read`，二进制走 `readBytes` / `copy` / `move`（第五十一轮）

`office.files` 以前只认文本：`read` 强制 utf8（读一张 PNG 得到一串替换字符）、
`write` 把非字符串 JSON 化、也没有整份复制。现在两条路分开了：

| 调用 | 做什么 |
| --- | --- |
| `office.files.read(path)` | 读文本。遇到压缩包（zip / gz / zst）或二进制**直接报错并指路**，不再返回乱码：判据是前 8 KB 里有空字节、或开头命中已知二进制魔数（`%PDF-` / PNG / JPEG / GIF / ELF；两字节的 BM / MZ 不收，真文件由空字节那条接住） |
| `office.files.readBytes(path)` | 读原始字节（二进制安全）。返回值可直接交给 `write` / `copy` |
| `office.files.write(path, 文本或字节)` | 整份覆盖写；**顶层**字节按字节写，不会被 JSON 化。嵌在对象 / 数组里的字节、`Map` / `Set`、函数与 `undefined` 一律**拒绝** —— JSON 化会把它们变成数字键或 `{}`，属于静默坏文件 |
| `office.files.copy(from, to, { overwrite? })` | 整份复制（内核复制，不经过脚本内存）。目标已存在要显式 `overwrite:true` |
| `office.files.move(from, to, { overwrite? })` | 移动 / 改名。跨设备时自动退回「复制 + 删源」 |

三条口径：① 源与目标相同时**拒绝**（移动会自毁；Windows 上按不区分大小写比较，硬链接与
符号链接这类别名由内核层的 `realpath` + inode 判定接住）；
② 目标已存在且没给 `overwrite:true` 时报错并给出换名建议 —— 与 `preview.pdf` /
`office.image.fetch` / `archive.extract` 是同一口径；③ 复制会把源文件登记成这次读过的输入
（台账的 `source`），移动不会（移动之后源路径已经不存在）。
两个动作只处理**单个文件**（源或目标是目录会直接报错，不递归、不跟随链接）。
路径按工作目录解析，相对路径里的 `../` 可以出目录（与 `write` / `remove` 一致），绝对路径也允许。
`office.files.read` 的报错文本对 PDF 会直接指到 `office.pdf.info / text / pages`，
对 docx / xlsx / pptx 指到对应格式的 `read`。

### 极简模式里不注入这九个工具（第五十轮）

DSH 内置的「极简模式」（preset id `minimal`）只留一个 shell 与一段人格，
九个 `office_*` 工具与它无关：本插件在这些 agent 上装一层
`ctx.tools.restrict({ deny })` 掩码，工具从 schema 展示、名字解析到执行一起消失
（不只是提示词省几个字节）。办公模式、标准模式、PTC 模式、创造模式都不受影响 ——
在 ptc 会话里直接调 `office_run` 的那条路子照旧。

**为什么不是 preset 里写一行 `disabled: true`**：宿主 `dsh-tools` 的工具注册表按
作用域分层，agent 可见的工具 = 全局层 + 作用域链上每一层。本插件是 profile 层的
一行，工具注册进**全局层**，每个 preset 都继承得到；preset 里写
`- id: dsh-office-mode / disabled: true` 只关掉该 preset 自己那一份（它本来也没有），
关不掉全局层那份。实测（第五十轮）：ptc preset 里 `tool-workflow` / `workflow-ptc` /
`tool-ralph` / `tool-plugin-manager` 四行都是 `disabled: true`，而 profile 层把这四行
打开之后，ptc 会话的实际工具面里 `workflow` / `ralph` / `plugin_manager` 一个不少。
按 agent 隐藏继承来的工具，唯一可用的机制就是 `restrict()`。

判据只看 preset id，**读不到就不隐藏**（没有 `agentPresets` 服务的单会话部署照旧
可用）；设置页改了工具开关会重算掩码，已经在跑的极简会话也跟着补齐。实现与边界见
`src/preset-guard.js`，回归见 `test/preset-guard.mjs`（18 条）。

### 办公模式的工具面守卫：把 preset 那些「声明要关」的行真正落地（第五十四轮）

办公 preset 里有一整段 `disabled: true` 的行（命令行、后台任务、工作流、子代理、
待办、技能、目标、插件管理、任务看板）。上面那条作用域事实说明**它们运行期不生效**
—— 那些能力挂在 profile 层的全局层里，preset 的禁用行只作用于它自己那一层。

第五十四轮把「要不要真的关掉」做成了一个**设置项**：
「办公模式 → 办公模式的工具面守卫 → 藏掉编程类工具」，**默认关闭**。打开之后，
办公会话里由插件用同一套 `restrict()` 机制逐个掩掉那批宿主工具（清单是
`src/preset-guard.js` 的 `OFFICE_DEV_TOOLS`，与 preset 的禁用行一一对应，
`test/preset-check.mjs` 钉住两处不脱节）。

两点必须知道：① 它是**削减能力** —— 模型自己起子代理、命令行、后台任务、待办与
技能入口都会消失，所以默认不隐藏，由用户明确打开。**但渠道覆盖检索的派工不受影响**：
`office_search_dispatch` 用的是宿主的 `subagents` **服务**（`ctx.get('subagents')` +
`subagents.start()`），而掩码只挡**工具注册表** —— 服务与工具是两回事（服务由
`dsh-subagent` 提供，`subagent` 那几个工具由 `dsh-tool-subagent` 注册）；Agent Teams 的
`spawn_teammate` / `team_task_*` 也不在掩码清单里（预设没声明要关它们）。
② 它是**即时生效**的：改完广播一过来，已经在跑的办公会话就会重算掩码（那些入口当场
消失，不必重开），关掉开关同样立刻撤下。宿主工具逐个名字装掩码，装之前先用
`tools.get(name)` 筛一遍「本机注册在全局层的名字」，本机没有的跳过并在日志里留痕
（不同机器装的能力不同，「以为藏了、其实没藏」这种失败必须看得见）。

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
  （`av.chunkSeconds`），相邻块重叠 0.5 秒（`av.overlapSeconds`，上限半块）。
  分块解决的是单块大小，不改变时长上限；超了就明确报错。
- **切点不会切开一句话**：块不是完全切开的 —— 窗口比自有区间多出那段重叠，
  跨在切点上的短音在同一块里解完；跨得更深、在窗口末尾被切断的句子由**边界回退**
  兜住：这一块不出稿，把下一块的起点回退到那句话的话头，让它整句在下一块里解一遍
  （`resolveChunkBoundary`）。同时用「已出稿水位线」保证重叠区不会被两块各报一次 ——
  不丢字、也不重复。`transcribe` 回执里的 `rolls` 就是这次回退了几次。
- **只识别，不越界**：不翻译、不认说话人、不做流式与降噪；块与块之间仍不共享上下文
  （上一块的词不会成为下一块的解码提示，切点前后的用词可能不一致）。

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
- **产物只落缓存目录**：默认 `.office/runtime/cache/python/out`，按修改时间对比收出「这一次
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
// r.files = [{ page: 1, path: '.office/runtime/cache/pdf/…-p1.png', bytes: …, reused: false }, …]
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

## 文档互转与统一出 PDF：`office.convert`（第五十三轮）

此前办公模式能**造**三件套与 LaTeX 论文、也能**读**它们，但「手上这份换个格式」没有入口：
别人发来一份 `.md` 要交 `.docx`、论文源要一份 `.md` 摘要、写完的 `.md` 要一份 PDF，
都只能重新生成一遍。`office.convert` 把「格式 × 格式」补成一张闭合的表。

三件引擎各管一段，都在运行期探测（`office.convert.check()` 报实况）：

| 引擎 | 管什么 | 探不到时 |
| --- | --- | --- |
| pandoc | md / docx / tex 之间的互转（也认 odt / html / rst / txt 源） | 文本互转那条路不可用，其余照常 |
| 本机 TeX 发行版（latexmk / xelatex） | `.tex` → PDF | `.tex` 出 PDF 不可用 |
| 宿主随部署安装的 LibreOffice（`@deepseek-ai/libreoffice-kit`） | docx / xlsx / pptx → PDF；老格式 `.doc` / `.xls` / `.ppt` 与 OpenDocument 的 `.ods` / `.odp` 也由它出 PDF（那几种 pandoc 读不了） | Office 三件套与老格式出 PDF 不可用 |

| 调用 | 做什么 |
| --- | --- |
| `office.convert.check()` | 三件引擎的实况 + 矩阵（每类源能转出什么、走哪条路、缺什么） |
| `office.convert.file(path, { to, out?, overwrite?, engine?, media?, fragment?, reference?, font? })` | 转一份文件；`to` = `md` / `docx` / `tex` / `pdf`（别名 `markdown` / `word` / `latex`） |

```js
// 别人发来的 md 要交 Word，再顺手出一份 PDF
await office.convert.file('会议纪要.md', { to: 'docx' });                 // 会议纪要.docx
await office.convert.file('会议纪要.md', { to: 'pdf', out: '纪要-交付.pdf' });

// Word 转回 md 时图一起带走（抽到 会议纪要.media/），论文源编成 PDF（本机 TeX）
await office.convert.file('会议纪要.docx', { to: 'md' });
await office.convert.file('thesis/thesis.tex', { to: 'pdf' });
```

几条口径：

- **默认落点**：与源文件同目录同名、只换扩展名；`out` 指定别的路径。目标已存在且内容不同时
  必须显式 `overwrite:true`，内容一致就不重写（`reused:true`）。
  **例外**：`.tex` 出 PDF 的默认落点就是 TeX 自己写 PDF 的位置（源文件旁边同名）——
  那里**已经有一份 PDF 时，不传 `overwrite:true` 会被当场拒绝**（编译会直接覆盖它，事后再报
  就晚了）；传了就重写并把这次产物登记上，`reused` 恒为 false。想两份都留就传 `out`。
- **同一份文件重复转换得到同一份产物**：pandoc 的 `SOURCE_DATE_EPOCH` 钉在源文件的
  **修改时间**（mtime）上（否则 docx 里的 `dcterms:created` 每次都不同，「内容一致就不重写」
  永远失效）。所以严格说这条承诺的前提是「文件没改过、mtime 也没变」—— `git checkout`
  或从别处复制会换掉 mtime，那时第二次转换会被覆盖闸拦下，确认重转传 `overwrite:true`。
  `.tex` 编译（本机 TeX）与 LibreOffice 渲染的产物带生成时间戳，不受这一条保护。
- **中文**：`.md` 出 PDF 默认「先转 docx 再交给 LibreOffice」，不依赖字体配置；LibreOffice
  不在时退到 `pandoc --pdf-engine=xelatex` —— 那条退路上，**源里有中文（或读不出正文的
  二进制源）而探不到中文字体时报错不做**：pandoc 找不到字体会静默丢字（退出码仍是 0，
  PDF 里只剩空白）。`font` 选项或 `DSH_OFFICE_CJK_FONT` 可点名字体。
- **md → tex 默认出能直接编译的整份**（导言区与中文字体都配好）；往自己的论文里并时传
  `fragment:true` 拿正文片段。`.tex` 出 PDF 时 `clean` 只管中间产物（`'aux'` 默认 / `'none'`）——
  传 `'all'` 按 `'aux'` 处理，因为它会连 PDF 一起删掉。
- **边界**：pandoc 读不了 PDF 的正文（要文字用 `office.pdf.text`，要画面用 `office.pdf.pages`）；
  `tex → md` 只保证正文与公式，thuthesis 的自定义命令、版式与参考文献样式都会丢。
- **转换只换容器**：要按目标格式的规矩重排版（封面、目录、版式助手），仍然用
  `office.word` / `excel` / `ppt` / `tex` 的 `create`。

引擎缺失时装一个即可，不用改配置：
`winget install JohnMacFarlane.Pandoc`（或 pandoc.org）、TeX Live / MiKTeX；
LibreOffice 由 DSH 随部署安装。便携版用 `DSH_OFFICE_PANDOC`（可执行文件或命令名）
与 `DSH_OFFICE_CJK_FONT`（字体名）指路。

## PPT 图片：插进去，也抽出来

演示文稿里最常见的两种「图片需求」是反的：一种是**把图插进已有稿子**（改版、补素材），
另一种是**把稿子里的图抽出来**（复用、核对、交给 `read_image` 看）。两条路现在都有：

| 调用 | 做什么 |
| --- | --- |
| `office.ppt.images(path, { out?, overwrite? })` | 抽出包内嵌图。`out` 省略时写进 `.office/runtime/cache/ppt-images/<词干>-<内容哈希前 10 位>/`（同 basename 的不同稿子不会互覆）；显式 `out` 默认不覆盖已有文件，要重写传 `overwrite:true`。返回可直接交给 `read_image` 的路径 |
| `office.ppt.insertImages(path, items)` | 把图片插进**已有的**演示文稿（不是重新生成整篇） |
| `revise(path, [{ slide, addImage: {...} }])` | 与改文字同一批做：一次调用里既改标题又插图 |
| `read(path).stats.media` | 读路径上先看清包里有哪些图（部件名 / 内容类型 / 字节 / 像素 / 在哪几页） |
| `readSlides(path).pages[].shapes[].image` | 每个 `p:pic` 形状带上 `{relId, part, contentType, bytes, width, height, crop}` |

```js
// 抽出稿子里的图，交给 read_image 看
const r = office.ppt.images('季度汇报.pptx');
// r.files = [{ part: 'ppt/media/image1.png',
//              path: '.office/runtime/cache/ppt-images/季度汇报-a1b2c3d4e5/image1.png',
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

## 图片：取、看、插、抽（第七十二轮）

配图一条链从「拿到图」到「嵌进文档」。`office.image` 管获取，三件套管使用：

| 调用 | 做什么 |
| --- | --- |
| `office.image.info(path)` | 看本地图片：体积 / 魔数类型 / 宽高 / 长宽比 / `sha256` / 能不能嵌 Word、PPT。读不到就回 `{ok:false, reason}`，不抛 |
| `await office.image.fromPage(url, {limit?, minWidth?})` | 列一个网页里的图片候选：og:image / twitter:image 优先，`<base href>`、srcset 与懒加载属性都认，按解析后 URL 去重；候选可直接交给 `fetch`。网页图的许可未知，发布前自己核对 |
| `office.image.extract(docx\|pptx\|xlsx, {out?, overwrite?})` | 抽包内 `word/media` / `ppt/media` / `xl/media`，默认落 `assets/<词干>-media/`，文件名带内容哈希前 8 位；每张经统一文件环境落盘、进产物清单；已存在默认跳过 |
| `await office.image.search(query, …)` / `fetch(候选)` | 免 Key 图库检索（维基共享资源 / Openverse）与取图；候选带原图字节数，原图超上限时自动改用缩略图；`credit` 传给文档的 `source` |

怎么用：

- **Word**：`builder.image(path, {widthCm, caption, source})` 单图；
  `builder.images(items, {columns: 2|3, gapCm, widthCm})` 多图无框网格 —— 每格图 + 题注 + 来源小字，
  末行补格；缺图 / 非 PNG·JPEG 留「图片不可用：…」占位，整篇继续。
- **Excel**：`sheet.image(path, {at, widthCm, heightCm, alt, name})` 原生图片部件
  （PNG / JPEG，oneCellAnchor），与 `sheet.chart` 共存，同图重复插入只存一份 media。
- **PPT**：`deck.image` / `deck.images`，以及往既有稿里插的 `insertImages`（见上一节）。

```js
const meta = office.image.info('素材/封面.png');
const page = await office.image.fromPage('https://example.com/report');
const picked = page.results[0];
const img = await office.image.fetch(picked, { to: 'assets/report.png' });
office.image.extract('旧稿.docx', { out: 'assets/旧稿-media' });
```

取图边界：候选的 `fetchUrl` 默认优先原图；原图字节超过取图上限（默认 8 MiB，
  调用方最多可放到 32 MiB）时改指缩略图。`fetch` 拿候选对象取图撞上体积上限，
  会用同一上限回退缩略图再试一次，其他错误不重试。返回的 `source.url` 是实际取到的地址。

边界：`info` 只按魔数认 PNG / JPEG / GIF / WebP；`fromPage` 只解析服务端返回的 HTML
（纯 JS 渲染的页面可能 0 条）；`extract` 抽得出 EMF / WMF / SVG / TIFF，但 Word / PPT 不认，
要先转成 PNG。

## 产物预览面板：在会话页签里看每一页的实际排版（第六十八轮）

生成完的 pptx / docx / xlsx / pdf 不用调工具也能看：会话页签行最右多一格「预览」
（`conversation.view` 的 `office-preview`，order 50，与「记忆」并列、同属 session 作用域）。

- **列文件**：扫描当前会话工作目录（深度 2，跳过 `.git` / `node_modules` / 隐藏目录，
  上限 20 条，按 mtime 倒序），不读文件内容；
- **看大纲**：pptx 走 `office.ppt.read()` 的 `stats / outline / warnings` —— 左列每页带
  版式徽标（`cover` / `cards` / `agenda` / `gallery` …）、标题、风险角标；底部风险清单
  点一条直接跳到那一页；
- **看页面**：右侧大图来自 `office.preview` 同一套引擎（LibreOffice），按页现渲染，
  dpi 24–300 可调（默认 110）；图只进 `.office/runtime/cache/preview/`，同一份文件同一 dpi
  命中缓存。渲染一次 2–5 秒，所以**切页 / 换 dpi 才取图**，打开面板只取第 1 页。

三条只读路由（`src/preview-view.js`，`GET` / `HEAD`，其余 405）：

```
GET /office-preview/files?cwd=<绝对路径>
GET /office-preview/outline?cwd=&path=<cwd 内相对路径>
GET /office-preview/page?cwd=&path=&page=1&dpi=110
```

信任边界与记忆面板**同一套**（`?cwd=` 必须命中宿主认得的工作目录，`path` 只收相对路径、
必须落在 cwd 内，越界一律 404）；错误分档：参数/越界 400、越权/穿越/不存在 404、
引擎不可用 503、渲染失败 502，每条都带「下一步怎么做」。非 pptx 的页数不预检，
由引擎报；引擎没装时清单与大纲照常可用，只有取图降级。

面板没有「改」的能力（只读）：要改文字/几何仍走 `office.ppt.revise` / `office.word.update`，
改完在面板里点「刷新」看新的一页。本轮只在测试替身里验过交互与数据链（`test/client.mjs`、
`test/preview-view.mjs`），真实浏览器像素与「重启 DSH 后页签是否出现」尚未活体验证 ——
要重启 DSH 对新建会话生效。

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
| 热记忆 | 用户偏好、项目约定（小、常驻） | `office_memory` 的 `add` / `replace` / `remove` | `office_help` 与 `office_run` 的反馈里自动带投影；也可 `layer:'hot'`。读的次序是「**重要度 → 最近写入**」（`recency` 弱先验：`updatedAt` 每日衰减 0.995） | 按「重要度低、更旧」下沉到归档，并在 `MEMORY.md` 留一条指路条目（回执点名 id，`action:'sunk'` 看清单）；单条超过上限（默认 2800 字节）直接拒绝写入 |
| 台账 | 每份交付物的登记：路径、格式、主题、复检统计、结构摘要、目的 | `office_run` 写盘后**自动**登记（也可 `action:'log'` 手工补） | `office_memory({ action:'read', layer:'ledger', query })` | 超过条数上限时，最旧的一批滚成月度归档摘要 |
| 归档 | 下沉的旧热记忆 + 滚动的旧台账（只读） | 由上面两层触发 | `office_memory({ action:'read', layer:'archive', query })` | **摘要文件**数有上限（一个月一个文件），最旧的摘要连文件一起删（唯一真正会丢东西的地方） |

**开场注入与回执减税（第七十六轮）**：写记忆时额外生成 `.office/memory/HOT.md`
（只放 `user` 条目，最多 4 条 × 70 字、总 1,100 B），persona 的「记忆」节用
`{{office_hot_memory}}` 在**新会话第一轮**就注入这段；项目条目仍按话题投影。写操作回执里
正文截 120 字，「用量」只在 used/limit ≥ 0.95 时印；`office_help` / `office_run` 的投影
没变化时返回空串，不再贴「与上次相同」的一行。

**读侧分区：一个工作目录一份记忆，怎么只看其中一类（第五十四轮 新-12）**：
三层记忆装在一个库里的代价是「多个主题混在一起」，而读取从前只有关键词打分一条路
（打分是软的：正文里出现那个词就算命中）。现在 `read` 支持**分区**：给 `tags` /
`entities` 就是**硬过滤**，只出**显式声明过**这些标签 / 实体的条目。

```js
// 只读「发布流程」这个分区：没声明过这个标签的条目一条都不会出现
office_memory({ action:'read', layer:'all', tags:['发布流程'] })
office_memory({ action:'read', layer:'hot', entities:['dsh-office-mode'] })
office_memory({ action:'read', layer:'ledger', tags:['季度汇报'] })   // 台账按 theme 判
```

三层各自的归类字段：热记忆与归档条目看写条目时声明的 `tags` / `entities`，台账行看
`theme`（且台账没有实体）。同一个维度里给多个词是**合取**（每一个都要声明过）——
分区是一把钥匙，不是一串检索词。回执首部会印出这次的分区口径与隔掉的条数，
所以「分区里没有」与「记忆里没有」不会混成一句话。

带 `query` 时热记忆也**只过滤、不重排**：一个词都没命中的条目不出现（与台账 / 归档
同口径），但排序仍是「**重要度 → 新在前**」。理由是可核实的两条：热记忆装的是**约束**，
一条 `low` 的笔记命中两个词不该压过一条 `critical` 的「永远不要…」（`limit:1` 时
那条约束会直接消失）；同族两版里也不该让**旧版**靠命中数冒到当前版前面（第四十二轮
量到增益才落地的 `recency` 先验）。台账与归档是另一回事 —— 它们是大规模检索面，
带 `query` 时按「命中强度 → 时间」排，并套 strict-v1 分档丢弃。

空白的分区键（`tags:['']`）会被**拒绝**：它不等于「没给分区」——静默当没给，回执里
连分区行都不印，读的人会把「没过滤」当成「这一类只有这些」。要不过滤就传空数组 `[]`。

**可见性与权限：这套记忆没有「用户身份」**（第五十四轮 19-7）：可见性是
「哪个根 × 哪个目标 × 哪个分区 × 哪个来源档」四个已有维度的乘积，加上三个执行点
（自动投影 / 显式读 / 写入期质量门）上的判据。模型全文见
`docs/memory-visibility-model.md`；`office_memory({ action:'status' })` 会把当前生效的
口径印出来（读得到哪几个根、来源档分布、以及**没做**的那些：多用户身份 / ACL /
授权撤销 / 任何可写端点）。

**容量口径要注意两处别配对错了**（面板的容量条按这个来）：

- **归档的上限 `archiveKeep` 管的是「摘要文件数」**（一个月一个），不是条目数。
  一个月里沉 30 条仍然只占 1 个文件，所以是 `1 / 60 个` 而不是 `30 / 60 个`。
  条目数另算：页签「归档（30）」与指标卡报的是它。
- **容量条用的是未过滤、未截断的总数**：搜过之后列表里只剩 3 条，容量条仍然报
  `25 / 500` —— 它回答的是「离上限还有多远」，不该跟着搜索框缩水。
  快照里这两组数分开给：`counts.hot/ledger/archive/...`（列表口径）与
  `counts.ledgerTotal/archiveFiles/archiveItems`（容量口径）。

**投影贴什么、贴几次（三只时钟 + 两道节流）**：`office_help` 与 `office_run` 的反馈尾巴上
挂着热记忆投影，而这两个工具在同一个会话里会被反复调用。

- **热记忆**：`writeRevision` 变了、第一次见这个会话、或心跳（每 12 次未变化一次，
  完整性兜底）→ 贴正文；贴的时候**按本次话题筛**（`office_help.topic` / `office_run.script`），
  无关条目折叠成一行（带条数与读取入口）。
- **节流一：本会话已贴过的条目不再重贴。** 状态记的是「哪些条目真的贴过」，不是
  「上一次是什么话题」—— 话题 A→B→A 时第二次 A 不再重付一份全文（第四十八轮 P0-2）。
- **节流二：投影预算**（`projectionBudgetBytes`，默认 4096 字节，0 = 不限）限的是
  条目正文那一块，按「命中强度 → 重要度 → 写入时刻」排序先给最相关的。超出的条目不丢：
  折叠行报「超出本次投影预算」的条数，并留 `office_memory({ action: 'read', layer: 'hot' })`。
  给了预算时首行会报这次的实际字节（成本可见）。
- **台账近况**：用自己的 `ledgerRevision`（最近几条的指纹）判断，变了就贴，没变折叠成一行。
- **拿不到会话身份**时不假装「刚贴过」，每次都贴全文。

本机实测（真实记忆库副本、一轮 20 次调用）：**改造前 76,656 B → 改造后 32,389 B（−58%）**，
全文形态 15,444 B → 预算 4,096 B 时 6,302 B。三条边界写在 `src/projection.js` 顶部：
① 每 12 次强制重贴一次全文（防宿主的工具结果压缩把早先那份裁掉；这是**有界的近似**，
残余风险已在文件里写明）；② 折叠一律带条数与读取入口；③ 用量行照旧给全量（容量是
「记忆库多大」，不是「这次给了多少」）。

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
2. **容量满了向下沉，不静默丢；下沉点名。** 下沉的条目进归档且可检索，`MEMORY.md`
   里留指路条目（带条数与 `action:'sunk'` 的入口）；连指路都放不下的极端容量下，才只保内容
   —— 而且**先让指路条目让位，不动 critical**（指路是 `low`，职责是服务条目）。
   写入回执会逐条点名被下沉的 id 与 60 字预览，`office_memory({ action:'sunk' })`
   按时间列最近几次；归档读侧每行标明原因（热记忆下沉 / 台账滚动 / …）。
   单条本身就超过**单条上限**（默认 2800 字节）则直接拒绝并给压缩提示，不截断内容；
   超过建议线（默认 1200 字节）只在回执里提醒。
3. **读是有界的。** 一次 `read` 有「条数 / 单条字符 / 总字符」三重上限；`layer:'all'`
   时**三层各有各的字符额度**（4000 / 2000 / 1600，第二十四轮起的口径）—— 共用一份
   预算会让热记忆把后两层饿死，实测台账与归档恒返回 0 条。被截断时明确带 `truncated`。
4. **投影是有界的，而且只贴该贴的。** 挂在每次办公工具调用上的那段投影，热记忆没变时
   只给一行、本次话题下已贴过的条目不再重贴、条目正文还有投影预算（默认 4096 字节）——
   常驻不等于每次都复制一份（见上一节的实测数字）。

## 知识库（kb）：外部原文的块与字符区间

记忆层装的是「结论与登记」，kb 装的是**来源原文**：一份文档一行清单，正文按内容哈希
切成块，每块带**字符区间**（`span`）与**块头**（标题路径 · 来源 · 日期）。全机**只有一座**
知识库（第八十八轮），默认 `.office/kb/`（`$DSH_HOME/.office/kb`）；设置 `memory.kbDir`
可把它放到别的目录，下载的本地模型也落在库目录下的 `models/`。

```
.office/kb/
  manifest.jsonl              一文档一行：路径、folder、内容哈希、字节、块数、来源档、入库时间
  chunks/<docHash>/<n>.json   一块一文件：正文、span（字符区间）、标题路径、块头、来源
  index.json                  计数与体积（投影；读不到就按 manifest 现算）
  assets/<assetKey>/<n>-<名>  多格式入库抽出的图片资产
  kb.sqlite                   docs / chunks / FTS / embeddings / chunk_links / summaries
```

**manifest 是唯一真源，SQLite 只是可重建的索引**（第八十五轮）：同步按 origin 分区增量做，
hash 未变只更新元数据；拿不到 `node:sqlite` 时自动退回 manifest + 词法 BM25，入库 / 读取 /
删除不受影响。

入库支持文本、PDF、DOCX、PPTX、XLSX 与图片：文本直读，Office 解 OOXML 取文字，PDF 走
`pdfText()`，图片落 `assets/` 并在正文写 `![名](assets/…)`。最终只存**文本块与图片资产**，
不保留原始二进制；单文件抽取上限 8 MB。8–24 MB 的 PDF 在预览时会自动按页拆成小份，逐份入库（见下文）。

动作都挂在 `office_memory` 上：

| 动作 | 参数 | 说明 |
| --- | --- | --- |
| `kb-ingest` | `path`、`tier?` | 把工作目录里的一份文本按块入库。**同路径同内容 = 幂等**（一个字节不写）；**同路径新内容 = 换代**（旧块删掉、指向旧块的边同事务清掉、旧块 id 留墓碑） |
| `kb-search` | `query`、`limit?`、`tier?`、`path?` | **词法检索**（第四十五轮接线）：中文按字符二元组展开，`text`（正文）与 `head`（块头）两路 BM25 各自归一化后加权（0.75 / 0.25）；命中给块 id、`span`、块头与预览；**词表零重叠即弃答**（不会端出「最像的那块」） |
| `kb-list` | `limit?` | 清单：文档 id、路径、标题、块数、来源档、三档分布 |
| `kb-read` | `id` 或 `path` | 文档 id 给整份（额度 4000 字符 / 最多 12 块，截断时明说还有几块没给）；块 id 只给那一块 |
| `kb-drop` | `id` 或 `path` | 删掉一份文档：正文删掉、边同事务清掉、留墓碑（旧 Pack 不会再把它搬回来） |
| `kb-tree` | 无 | 虚拟文件夹树：每层的文档数、块数与字节；空文件夹不显示 |
| `kb-update` | `id`、`title?`、`folder?`、`tier?` | 改文档信息（标题 / 文件夹 / 来源档）；不重切块 |
| `kb-edit` | `id`、`text` | 改**文本型**源文件正文，然后按原路径重新入库；PDF / DOCX 等抽取文档没有可写源 |
| `kb-link` | `sourceId`、`targetId`、`kind?`、`note?` | 建块间关系；两端都必须在索引里 |
| `kb-related` | `chunkId`、`depth?` | 走 1–2 跳邻居，关系双向 |
| `kb-summarize` | `title`、`text`、`chunkIds` | 多块合成一条摘要块；摘要独立参与检索 |
| `kb-summaries` | `limit?` | 摘要块清单 |

面板里摘要块排在文档清单之后，默认折叠，每页 10 条（可选 10/20/50）。分页由服务端
算窗口，底部显示「第 a–b 条 / 共 N 条」并带上一页/下一页；页码越界会回拉到末页。
点标题才展开正文预览。块检索在面板的
「块检索」模式里**始终递归当前文件夹的子树**（在「全部文档」下就是整棵库），浏览
文档时仍按「仅当前层 / 含子文件夹」开关走 —— 两者不是同一个口径。

顶部共享搜索框按当前看法路由（第一百一十一轮）：文档看法发 `q`，只过滤标题与路径，
仍受「仅当前层 / 含子文件夹」开关约束；块检索看法发 `kbq`，始终递归，且不叠加文档
过滤。占位文字会写明「找文档」还是「找块」。清空按钮按当前看法清对应的一套状态，
切换页签不把上一页的搜索词带过来。

无命中时服务端的机器码会译成人话（`no-overlap` → 「查询词与块内容没有重叠」，
`no-docs` → 「这个范围内没有文档」），并写出这次检索词；命中卡片可用 Tab 聚焦，
Enter 或空格打开正文。文档清单的空态也带主动作：搜索无结果「清空搜索」、空文件夹
「全部文档」、整库为空「＋ 添加文件」。大图标视图的文档磁贴聚焦后 Enter 直接打开，
和列表行一致，不再只能双击。

知识库是**单库**（第八十八轮）：`kb-ingest` / `kb-drop` 的 `scope` 只为旧调用面保留，不再选根；
要按项目分档用 `folder`。

**面板上传（第八十九轮）**：知识库页的「＋上传文件」与拖放不经过会话，直接 POST
`/office-memory/kb-upload`（JSON，请求体上限 33 MiB，`data` 为 base64）；服务端把内容写到
`.office/runtime/kb-upload/` 临时文件，`kbIngest` 抽取、切块后立刻删掉临时文件。manifest 里
记的逻辑路径是 `uploads/<虚拟文件夹>/<文件名>`，**不指向磁盘原文件**；同名同内容幂等，
同名新内容换代。上传件只保留切片与图片（`assets/` 下的嵌入图），因此 `kb-edit` 对
`uploads/` 文档明确拒绝改正文：要改就删掉重传。设置页「选择目录」走
`/office-memory/fs-list` 自己列目录，不依赖宿主桌面桥；选完目录会等宿主确认写入后才显示
「已保存」，失败就在这一行显示原因，不再假成功。模型下载默认 **魔搭 ModelScope →
Hugging Face → hf-mirror** 顺序试端点（国内直连优先），`HF_ENDPOINT` / `HF_MIRRORS`
可覆盖后两个；清单与文件 URL 都由各自来源的接口形状决定，失败会列出每个来源的原因。

**大 PDF 自动拆分（第一百一十轮）**：8–24 MB 的 PDF 不再直接判「太大」。预览时服务端
用 PyMuPDF 按页重新封装成每份约 4 MB 的完整 PDF（单页仍超限时单独成一份并标出来），
每份一张待入库卡；入库仍走原有 8 MB 抽取上限。本机没有 PyMuPDF 时不会假装拆开，预览卡
照实写「自动拆分也没成功」并保留「识别文字」入口。超过 24 MB 的文件仍在预检阶段拒绝：
那是整份 JSON 上传的上限，不是拆分能力本身。

模型入库（勾选「用模型切片 / 摘要」并确认预估后）走 NDJSON 进度流：服务端每完成一箱
就写一条 `progress`，最后写 `result`；面板据此画进度条，显示「模型 a/b 批 · 已调用 N 次」，
有降级或截断也一并列出。旧服务端回普通 JSON，或环境没有可读的 NDJSON 流时，自动退回整段 JSON，行为不变。

三条口径值得记住：

- **切块是结构切块**（Markdown 标题分层 + 段落聚合，超长段落按可断点硬切），确定性、
  可复跑；**`source.slice(span.start, span.end)` 与块正文逐字相等**，这是「引用可回溯」
  的全部含义。不做语义切块；嵌入是可选的检索升级，见下一条。
- **块头是零依赖的上下文前缀**（标题路径 · 来源 · 日期）：官方数字里「给块前置
  50—100 token 专属上下文」值 −35% / −49% / −67% 的检索失败率，而它不需要任何模型调用。
- **默认检索是词法的，嵌入是可选升级**（第四十五轮；第八十八轮补上下载与单库）：中文按
  字符二元组（bigram）展开、拉丁词与数字整词保留，排序用 BM25（k1 = 1.2 / b = 0.75），正文与
  块头两路各自归一化后加权。配了远端 OpenAI 兼容 `/embeddings`、或启用本地 `bge-small-zh-v1.5`
  时升级成向量 + 词法混合检索；`selector.mode: "llm"` 可让独立模型从候选里挑编号，只做粗筛、
  不重排。本地权重可在设置页下载进知识库目录，启用后按需拉起本机推理进程（空闲退出），
  下载器本身只负责把文件放对位置。下载端点是
  `huggingface.co` 失败后自动换 `hf-mirror.com`（实测后者 metadata 971ms 返回），
  `HF_ENDPOINT` 指定单端点、`HF_MIRRORS` 改兜底列表。所以边界要照实说：
  **语义改写召不回**（查询与正文用词完全对不上时给不出结果）—— 换一个更贴近原文的词再试，
  别换工具重复检索。候选打分器是 `src/kb.js` 里的纯函数，探针与 action 面量的是**同一份**实现
  （`test/memory-probes.mjs` 的 C 组：自测集 15 条查询 hit@3 15/15、朴素整串 5/15）。
- 检索**有界**：一次最多扫 2000 块、最多给 8 条命中（`limit` 硬上限 20）；命中只给预览
  （每块前 280 字符），要逐字引用就走 `kb-read`（它有自己的额度）。次数占每回合配额的
  「知识块检索」一档（默认 2，0 = 不限制）—— 词法检索最容易被「没查到就换个词再查」拖着走。

**来源三档是硬规则，不是权重**：`user`（用户放进工作目录的文档）/ `verified`（≥2 个独立
来源确认过）/ `unverified`（单源检索结果）。`unverified` 可以进 kb 的**待核区**（读它的
回执会标「待核」），但**永远不能进热记忆** —— `add` 时声明 `tier:'unverified'`、或引用一块
`unverified` 的知识块，都会被当场拒绝。依据是实证的：投毒率 < 0.1% 就能把攻击成功率推到
80% 以上，而「按来源类别排除」有效、「加性可信度权重」与完全无防御不可区分；读时过滤还要
付准确率并误隔离合法记忆，所以设防放在**写入期**。第二家独立来源确认之后，用
`kb-ingest({ tier:'verified' })` 把同一份内容**升档**（正文一字不改）即可 —— 硬规则不留死路。

**入库评估闸：`verified` 从「一句话」变成「数得出来」（第五十四轮 19-8 的 P1-4b）**：
`tier` 从前是**纯声明** —— `kb-ingest({ tier:'verified' })` 只要拼写正确就过，而 verified
正是「能进热记忆当依据」的那一档。现在入库时**现算来源证据面**：`verified` 要求正文里
有 **≥2 个独立主机名**（同一站点的两个页面算一个来源；去掉开头的 `www.` 再比；
单标签的「主机名」如 `https://a`、`https://占位` 不算），数不出来就当场拒绝，并给出
三条去处（承认没核实 → `unverified`；材料是用户给的 → `user`；把第二份独立来源的地址
写进正文）。

三条口径要一起记住：

- **它是一道形状门，不是真伪门**：挡得住「声明两份独立来源、正文里一份都没有」这种自相
  矛盾、以及明显的占位符，挡不住刻意伪造（不同子域仍算两个来源，没有公共后缀表）。
  所以回执与文档说的是「**可核对的地址**」，不是「证据确凿」。
- **幂等重入不判门**：内容与档位都没变的再次入库照旧放行 —— 门是向前的，不追历史，
  库里那些在门之前入库的 verified 文档不会被追溯拦下。
- **台账侧的 ⚠ 是读时算的**，不是登记时的快照：块从 unverified 升成 verified 之后警告
  就消失，块被降档或删掉之后警告才出现。`log` 从不拒绝登记（交付物真实存在），
  依据的强弱读台账那一行时看得见（「⚠ 依据里 N 块待核 / N 块在 kb 里找不到」）。

**引用三条线**：把块 id 写进 `add` / `log` 的 `source`（形如 `kb:<内容哈希>:<块号>`），
它们会另存成引用字段（记忆的 `kbRefs`、台账的 `kbRefs`），于是「这条结论依据哪块来源」
「哪份交付物用过这块来源」都能反着查；`link` 也能直接指向块（`kind:'supports'` /
`'contradicts'`）。冲突边还带三分**类别**（`context-memory` / `inter-context` /
`intra-memory`）与三分**状态**（`unresolved` / `prefer-source` / `prefer-target`）——
**保留双方、显式承认未决**，不做自动裁决，同一条边可以改状态。

Pack（`export` / `import`）会把 kb 的清单与块正文一起搬走；导入按文档 id 幂等，
被墓碑判死的文档不会复活，**半份文档**（清单有、正文缺块）一律不入库。**导入也要过门**：
逐块自校验（块号、块 id、正文哈希三者对上才收）、带 `unverified` 引用的热记忆条目不入库、
档位冲突时取**更保守**的一侧。块文件里只放与这一块有关的东西（正文 / span / 标题路径 /
自校验哈希）—— 来源路径、入库时刻、来源档是**文档级**的，从清单现取，所以两份内容相同的
文档共用同一批块也不会互相覆盖。

**边界**：来源档是文件里的声明。插件保证「两处说法冲突时取最保守、导入时逐块自校验」，
**不承诺防伪造** —— 谁能改盘上的文件，就能改包里的字段；这正是「什么进 kb、什么进热记忆
必须显式」这条信任边界的另一面。

## 统一入口路由：`recall`（第五十七轮 19-8 的 P2-8）

记忆层与知识库是两条语料完全不同的检索路径，从前「先查哪边、什么时候下沉」全靠临场决定。
`office_memory({ action:'recall', query })` 把这件事收成**一条入口 + 一个可核对的门控**：

```
📒 统一入口（recall）：发布流程
- 路由：记忆层命中 1 条｜查询词覆盖 3/3｜最高重要度 critical
- 判定：记忆层够答 → **未下探知识库**
```

门控是纯判据（`src/memory.js` 的 `gateVerdictOf`），四档按顺序判，判据只看两项
**现算得出来**的事实 —— 命中条目的词表覆盖了查询词的几成、命中里最高的重要度是什么：

| 档 | 判据 | 处置 |
| --- | --- | --- |
| `empty` | 命中 0 条 | 下探 kb |
| `partial` | 查询词没被命中条目的词表全覆盖 | 下探 kb |
| `weak` | 词全覆盖、但没有一条约束级（`critical`） | 下探 kb |
| `enough` | 词全覆盖 + 至少一条 `critical` | 停在记忆层 |

两条纪律值得单记：

- **`enough` 是门控结论，不是「记忆能答对」**：它的意思是「不必再下沉到来源原文」。
  所以判据里没有语义判断 —— 同样的输入永远给同样的分档，测试钉得住它（`memory-recall`
  的 D 组）；代价是会偏保守（宁可多下探一次）。另外**覆盖是三层命中的并集**，不是
  「有一条命中覆盖了整句查询」，而且 `critical` 这一项看的是**热记忆层与归档里的 hot 条目**
  （台账行没有重要度字段）—— 所以热层被挤空、`critical` 沉到归档之后，门控照样认得出它
  （旧写法只认热记忆层，会一边说「没有约束级」一边在正文里印出 `critical`）。
- **换入口不换来额外的检索次数**：`recall` 与 `read` 带 `query` 共用同一条每回合配额
  （初查 / 细化各一次），只有**真的下探了**才再吃一档「知识块检索」配额。被配额挡住时
  回执如实写「未下探」并给出那条真的能走的路（`kb-read` / `kb-list` / 调设置）——
  不把「被挡住」说成「不必查」。落点与反向验证见 `test/kb.mjs` 的三条探针。

## 引用纪律：句级引用审计（第五十七轮 19-8 的 P2-9）

第十九轮附录把这条记成「落盘断言强制句级来源引用，评测套 ALCE 三指标」。落地前先得
回答一句绕不开的话：**三指标里哪几项在离线、无模型、零依赖的环境里算得出来？**

| ALCE 指标 | 原定义 | 本项目能算的那一层 |
| --- | --- | --- |
| citation recall | 每个陈述都被引用的段落**支撑**（蕴含判断） | **词法代理**：声称外部事实的句子有没有给出处（句级引用率） |
| citation precision | 每条引用都**支撑**对应陈述（蕴含判断） | **可对账的那半**：正文里写下的地址，在本次材料里对不对得上 |
| correctness | 回答是否正确（要事实真值与蕴含判断） | **算不出** —— 恒为 `null`，不许拿上面两项冒充 |

用法：产物落盘之后，把**正文**（Markdown / 纯文本）当 `draft` 传给解析步：

```js
office_parse_findings({ paths: ['.office/search/某主题/01-泛搜.md'], draft: 'docs/正文.md' })
```

没过时工具结果 `ok:false`（这是**断言**，不是提示），并**带句号**点名要改的是哪几句、
哪条地址对不上；摘要本身照常给全，所以一轮就能照着改（最多列 8 条，其余按句号在同一份
正文里定位）。四件要照实记住的事：

- **事实句判据是显式词表**（年份 / 百分比 / 带**外部**量词的数字 / 归因主语），不是
  语义模型 —— 判据能被人逐条核对，也会漏、也会误；回执里的结论只有「按这份词表看来」这一层。
  刻意不收「个 / 项 / 条 / 页」这类量词（那多半是在说这份产物自己的形状），年份也要求带
  「年」字（否则 `1920×1080` 会被判成年份）—— 误判的代价是硬失败，所以两个方向不等价，
  一律偏向「宁可漏、不可误」。
- **句界与合并**：只有句末标点算句界，**分号不算**（「一是…；二是…。（来源：X）」是最常见
  的写法）；句号**后面**的「（来源：…）」并回上一句，但**绝不跨标题 / 分隔线**并 —— 跨过去
  会让上一节那句白捡一个出处，而回执会说「通过」。围栏代码块与行内代码里的地址是示例，先剔掉。
- **不进分母的东西**：kb 块 id 只核形状（本审计没有库的入口）、`（来源：…）` 标记与
  `[1]` 编号引用只算「标了出处」；同站不同页算「站点级对上」，单独报数而**不计入**逐字
  对账的分子；`?id=1` 与 `?id=2` 算两条不同地址。把算不了的东西塞进分母，等于用一个漂亮
  的比率掩盖没算。
- **三态不许混说**：`通过` / `未过`（`ok:false`）/ **`没判出可核的东西`**（没有可判的事实句、
  或正文里一个地址都没有 —— 这不是「引用都齐了」，也不会印成「通过」）。
- **边界**：Office 三件套 / PDF / ODF 的正文回读不了（`read()` 只给结构与统计），传这些
  文件会被当场拒绝并给三条路：传写文档用的 Markdown 源、`office.convert.file({ to:'md' })`
  转成 Markdown、PDF 先 `office.pdf.text` 抽文本。

写入协议（与 mnemon 的 `add` / `replace` / `remove` 同形）：
- `add` 用于**新的独立事实**；`replace` 需要 `oldText` 能**唯一**命中（命中多条会
  被拒绝——宁可让你补长，也不要改错）；`remove` 只在用户明确要求或确有证据时用。
- `target`：`user` = 用户是谁、偏好与要求；`project` = 项目约定、环境事实、工具坑。
- `scope`（选填）：`workspace` / `global`，这一条只写进指定根；显式指定时
  `replace` / `remove` 也只在该根里找，不跨根兜底。
- 只记稳定的事实。检索结果原文、一次性进度、助手自己的推测都不记。

### 记忆范围与层开关

- `scope`：`workspace`（默认，一个项目一份）/ `global`（所有项目共用一份）/
  `both`（两者并存）。读的时候**全局在前、工作区在后**——具体项目的约定比通用
  偏好更该被最后看到。来自全局根的行标「（全局）」，未标的是工作区根。
- `globalDir`：全局层目录，默认 `$DSH_HOME/.office/memory`。相对路径相对**用户主
  目录**，不相对工作目录：全局层跟着工作目录走就没有意义了。它的 kb 是兄弟目录
  （默认 `$DSH_HOME/.office/kb`）——kb 与记忆目录由 `kbDirOf()` 按同一父目录定位。
- 按次指定根：`add` / `replace` / `remove` / `kb-ingest` / `kb-drop` 可带
  `scope: 'workspace'|'global'`，只动那一个根；根没开就报错，不会退到默认根。
  它不改变读的可见范围。跨项目通用的偏好写全局、项目流程留在工作区时用这条。
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
- 配额：一个回合里第一次带 `query` 的记忆检索、后续换词细化、知识库（`kb`）检索、图关系
  遍历各有上限（默认 1 / 1 / 2 / 1）。回合号从会话日志的 `turn/start` 事件读出来。用完会
  **明确拒绝并说明原因**（`kb` 那一档直接给出「改走 `kb-read`」这条路）。不带 `query` 的
  `read` 是看现状、不是查资料，不占配额。
- 两者都能在设置页关掉（`policy: off`、配额设 0）。

### 备份与迁移（Pack）

`office_memory({ action:'export', packPath? })` 把热记忆 + 台账 + 归档（索引与
Markdown 全文）+ 关系整包写成一个 JSON 单文件；`action:'import'` 读回来。
口径是**只增不改、按 id 幂等**：同 id 已存在就跳过，不覆盖现内容；归档按月合并、
按块去重。所以重复导入同一份包不会产生第二份。这份 Pack 含私有记忆，不要当公开
文件传。

不打包成 zip 是有意的：本插件零依赖，读回来这条路上任何一个压缩库都会变成新的依赖。

### 浏览面板

设置对话框左侧「办公与记忆」是**配置**页；会话页签行里的「记忆」是**浏览**页（它排在
「上下文」与「编队」的右侧 —— 第六十二轮从侧栏图标迁过来）：热记忆 /
台账 / 归档 / 关系 / 实体五个页签，带搜索与工作目录切换。数据来自本插件注册的只读
端点 `GET /office-memory/snapshot`。

页签属于**会话**（`conversation.view`，与「上下文」「编队」同一个槽位），所以要先打开
一个会话才看得到它；这是这次迁移的已知代价，换来的是它正好落在用户点名的那个位置。

端点只服务**宿主认得的工作目录**，三个来源依次是：本进程里真跑过办公工具的工作目录、
宿主工作区登记表里的项目目录（`ctx.workspaceRegistry`，跨重启仍在）、会话日志里出现过
的目录（`ctx.sessionPersistence`，只在前面都没命中时查一次）。所以**重启 DSH 之后、
或者一个还没跑过办公工具的工作区里，面板照样能显示 `.office/memory` 里的内容** ——
不用先随便调一次办公工具把它「喂」进来。请求里点名的工作目录不在三个来源里时仍然
404，端点不会照着任意路径去读盘。

**默认打开哪一个工作区**（第六十四轮）：**当前会话的工作目录**。这一格是 `conversation.view`
（session 作用域），宿主会把 `sessionId` 与 `useSessions` 一起发给视图组件，面板就用
dsh 自己的那套取法 `useSessions((s) => s.byId[sessionId]?.cwd)` 把默认根定下来，
并且**第一次请求就带上它**（不会先开成别的工作区再换）。拿不到会话上下文的老宿主 /
空会话退回老行为：不带 `cwd`，由宿主在三个来源里挑一个（优先「本进程刚跑过工具的那个」）。
之后要换目录用面板标题栏的工作目录选择器，选择器里也列出了全部候选根。

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

**面板怎么「写」**（第八十九轮扩到知识库）：条目上有「编辑」——热记忆可改正文 / 归到 /
重要度；知识库文档可改标题 / 文件夹 / 来源档，文本型文档还能「改正文」（保存后重新切块）；
从本地上传的 `uploads/` 文档没有源文件，只能删掉重传。知识库还多了两条直接写入：
文档列表与文档抬头都有「删除」（先出确认卡，删的是**整篇**，连同块与关系），块抬头有
「加关系」（目标块、关系类型、备注）与相邻块旁的「解除」。这些保存走
`POST /office-memory/mutate`，由服务端调用 `memory.js` 的同一套方法落盘，成功后重新拉快照。

记忆条目的「改写 / 忘记 / 删关系」仍保留**生成指令卡**的旧路径：需要用户先看见「要记什么」
再点头的场合，用指令卡比直接写更稳。知识库的删除与块关系则是直接写入；两条路并存，面板
不强制一种。

页内「直接记一条」默认收起成一行，点开才是完整输入卡：常驻表单会把列表挤到首屏之外，
而它多数时候只是备着。标题栏的「+ 记一条」会把这一行展开并把焦点送进输入框。

归档层刻意**不给**写入动作（它由热记忆下沉与台账滚动产生），面板上给一句说明而不是
一个点了会报错的按钮。台账 / 关系 / 实体页签也仍只读。

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

**宽度**（第八十九轮）：面板不再固定跟聊天列一样窄 —— 上限取
`min(1380px, calc(var(--dsh-chat-content-width, 680px) + 560px))`，默认宿主下约 **1240px**
（旧版约 728px），左右留 `clamp(18px, 2.6vw, 40px)`；仍引用宿主变量，窄屏与宿主改列宽时
会一起收。

**滚动：面板根是唯一的滚动容器**（第二十轮修）。宿主主面板那一格是
`.pI_x6G_centerCol{display:flex;flex-direction:column;overflow:hidden}` —— 它**不给**
`main` 槽位任何滚动能力，而槽位锚点的 style 是 `display:contents`（不产生盒子），
所以面板根元素就是那个 flex 项。第二十轮之前面板没有高度也没有 `overflow`：内容比视口
高时被中心列裁掉，只有列表内部那 520px 能滚，指标卡 / 容量条 / `ITEM DETAIL` / 页脚
**全都在首屏之外够不着**。现在：

- 面板根 `height:100%` + `min-height:0` + `overflow-y:auto` + `overscroll-behavior:contain`；
- **列表不再自己滚**（去掉内层 520px），消除双滚动条；「再显示 N 条」滚的是面板根；
- **标题与搜索不再吸顶**（第八十九轮）：整块标题会随内容滚走，只有页签与来源筛选这一条
  导航吸顶（底色取 `--dsw-alias-bg-base`），避免半屏一直被标题占着；
- **图谱缩放要按住 Ctrl / ⌘ / Alt + 滚轮**（或视图条 ±）：原先普通滚轮一律
  `preventDefault` + 缩放，画布占了列表上面一大块，鼠标停在图上就滚不动页面 ——
  普通滚轮现在留给滚动。

真实浏览器实测（`test/web-verify.mjs`）：`clientHeight 802` 正好等于中心列高度、
`scrollHeight 1257`、滚到底后页脚可见、吸顶导航条贴顶差 0；普通滚轮面板 `scrollTop`
从 460 到 397 而画布 zoom 仍 100，Ctrl + 滚轮 zoom 100 → 115。

该端点的信任边界（`src/view.js` 顶部有完整说明）：

- **读端点仍只读**：快照 / 导出只接受 `GET` / `HEAD`，其余方法 405；写走独立的
  `POST /office-memory/mutate`，不与读混在一条路径上。
- **写端点的边界**：只接受 POST，`content-type` 必须是 `application/json`（跨站表单
  发不出这个类型，挡掉 CSRF）；body 上限 512 KB；`?cwd=` 必须命中宿主认得的根；
  记忆总开关关掉时 403。写入调用 `memory.js` 的同一套方法，不另起一套实现。
- **只服务见过的根**：`?cwd=` 必须命中「本进程真的跑过办公工具的工作目录」，否则
  404。没有这一条，一个本地 HTTP 端点就成了「传任意路径读任意目录」的数据外泄面。
- 组合里没有 `webServer`（例如 headless）时端点不注册，面板显示可读提示。
- 快照里 `stores[]` 带 `bytes` / `files`、`counts` 带 `hotBytes` / `projectBytes`，
  容量条与存储域卡片靠它们；字段缺失时那两块静默不画，不影响其余部分。

### 图库（第一百零六至一百一十一轮）

图库和热记忆、台账、归档、知识库并列在会话页签「记忆」里。左侧除了「全部图片」、
「收藏」与「分类」，还有一棵**文件夹**树；文件夹不是标签的别名，而是图片的存放位置。

**版式（第一百零九轮）**

图库不再默认铺缩略图网格，主列改成知识库同款资源管理器：

- 左栏是「图库 · 文件夹」树：根「全部图片」带总数，子文件夹可逐级折叠；
  折叠只影响这一栏，不改变当前浏览层。当前文件夹行才出现改名与删除。
- 主列默认「详细信息」：第一列缩略图，随后是名称、类型、数量、大小、加入时间；
  名称、大小、加入时间可点列头排序。文件夹行排在图片前，点行进入下一层，
  类型列写「文件夹」、数量列写总数；「大图标」时文件夹显示为磁贴。
- 命令行：← 后退、↑ 上一级、刷新、新建文件夹（下方内联输入条，Esc 取消）、
  详细信息 / 大图标、排序与方向、文字搜索、允许内容发给模型、上传图片。
- 状态栏：当前层的文件夹数、图片数与总数、已选图片名、当前位置、
  截断提示与这次检索用到的融合信号。切换文件夹会清空检索与选中。

**文件夹模型**

- 一个文件夹就是一段路径（`A/B/C`）。规则：最多 8 层、每段 ≤40 字、整条 ≤200 字；
  拒绝 `.` / `..` 与空段。旧图库没有 `folders` 字段时，按图片的 `folder` 反推目录，不搬文件。
- 清单里同时记显式文件夹与图片的 `folder`：**空文件夹也能存在**，父目录先于子目录
  自动补齐。路径、祖先与图片归类都写同一份 `manifest.json`，沿用原有的写锁与
  跨进程锁，不另开一份索引。
- 计数分 `direct`（直属图片）与 `total`（含后代）。点一个文件夹只显示**直属**图片；
  子文件夹里的图不会混进来。空态会说明这一点，并给「全部图片」的回退入口。
- 新建默认建在当前文件夹下；改名只改当前这一段，后代路径与图片的 `folder` 一起跟着改；
  删除只允许空目录（`total === 0`），否则报「文件夹里还有 N 张图（直属 M 张）…」。
  删除空父目录时，它下面已空的子目录一起收掉。
- 图片的文件夹在详情里改：输入路径或从已有路径里选，保存走 `POST manual` 的
  `patch.folder`。上传会把新图直接归到当前浏览的文件夹；移动、上传、移除之后
  左侧树与计数会重新拉一次状态。
- 面包屑按路径逐段可点；「全部图片」清掉文件夹筛选。顶部搜索是全局的，会退出
  当前文件夹，避免「搜不到明明存在的图」。
- 文件夹只是组织维度，**不拼进图片的检索文本**：名称、描述、标签、分类与细节
  仍决定文字相关性。
- **检索只列真命中（第一百一十一轮）**：有查询词或查询图时，融合分 ≤0 的图片不再
  作为结果返回；零分是「所有信号都没命中」，负的图文余弦是「不相似」。无命中返回
  空列表与 `reason: "no-match"`，面板显示空态和「清空搜索」，不再把整库画成结果。
  只用标签 / 分类 / 收藏做硬过滤时仍返回过滤后的全部图片。
- 文字搜索与以图搜图默认按融合分从高到低排；清掉检索后回到按加入时间。空库或空
  文件夹的空态直接给「＋ 上传图片」，不用先去工具条找。

**名称、描述与细节（第一百零八轮）**

- 名称 ≤120 字，上传图片时默认取文件名；不填时列表回落显示 id 前 8 位。名称参与
  检索，也决定「名称」排序。
- 描述 ≤600 字，是用户可改写的覆盖：留空时显示视觉模型生成的总结，模型原文仍
  保留；填了就以改写为准，检索文本不再重复模型总结。
- 重新生成视觉描述会清空用户改写，让新总结直接显示出来，避免旧文案一直挡着。
- 「细节」是自由备注（≤2000 字），和标签、分类一样跟着图片走，不参与摘要生成。
- 「保存细节」一次提交名称、描述、标签、分类、文件夹、收藏与细节；服务端的
  manual patch 接受 name 与 description，add 也接受。

**视觉描述**

默认 `off`：只保存图片与手工标签，不生成描述。配置入口是
「设置 → 办公与记忆 → 图库 → 图库视觉描述」：

| 通道 | 需要填 | 行为 |
| --- | --- | --- |
| `off` | 无 | 只留一句「未启用」，不铺输入框 |
| `host` | provider + model | 用宿主已配置的多模态模型；宿主没声明图片能力时明确失败 |
| `openai` | baseURL + model（+ Key / Key 环境变量） | 把缩略图发到兼容端点，按 JSON 解析后写审美描述 |

未配置时点「生成视觉描述」不再只说「没有配置视觉模型」，而是直接指向
「设置 → 办公与记忆 → 图库 → 视觉描述」；能力面板里的同一项也带这条路径。
已经配置好但面板还停在保存前的旧状态时，点「生成视觉描述」会先向服务端重取一次
能力状态；不再拿旧缓存直接判「未配置」。服务端确认后仍回 `unconfigured` 的，
回执同样带这条设置路径。
图文对齐是可选的第二组（`none` / Voyage / Cohere / 兼容端点），只管图片向量，
关掉不影响标签、分类与视觉描述。

**接口与工具**

- `GET /office-memory/gallery` 的 `list` / `search` 都接受 `folder`（只匹配直属）；
  `status` 返回 `folders[]`（`path/name/parent/depth/direct/total`）与
  `counts.withFolder` / `counts.folders`。
- `POST /office-memory/gallery` 新增 `folder-create`、`folder-rename`、
  `folder-remove` 三个 action；`add` 与 `manual` 接受 `folder`、`name`、
  `description`（`manual` 也可放进 `patch`）。
- `office_memory` 的 gallery `op` 新增 `folder-list` / `folder-create` /
  `folder-rename` / `folder-remove`，`list` / `search` / `add` / `tag` 接受
  `folder`；`add` / `tag` 另接受 `name` 与 `description`。

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

两条入口（幂等，只增不改，重复跑不会产生第二份；设置页不再放这个一次性入口）：

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

这条通道诞生于「办公模式」preset 的 `plugins` 列表里**没有** `@deepseek-ai/dsh-tool-web`
的那几轮：那时办公会话的工具面里没有 `web_search` / `advanced_search` /
`platform_search` / `web_fetch`，派工给子代理圈的工具名一个都不存在，
子代理一启动就报

```
tools.restrict() names unknown global tools "web_search", "advanced_search", …
```

整条链路 0/N 全失败（2026-09-25 的真实会话就是这样：检索提纲要得出来，派工 7/7 全红）。
所以本插件自己带了一条联网通道，不依赖别的插件。

第二十九轮办公 preset 曾把 `tool-web` 加回来；**第三十轮起又收掉了**（按用户要求：
资料搜索不走宿主工具、不走 Tavily 这类要 Key 的三方 API，回到「子代理 + 网页抓取」
的最初设想）。现在子代理的检索与取正文改用插件自己的两个工具
`office_web_search` / `office_web_fetch`（默认先走 DeepSeek Search API 直连，再退宿主接缝与免 Key 抓取）；
这条通道同时还是落盘取证（office_search_run / 渠道覆盖检索）的唯一出口，
在没装子代理服务的精简部署里也一样能跑。

### 十条通道（第二十轮起不再只有 DeepSeek 一条路）

| id | 通道 | 要 Key | 备注 |
| --- | --- | --- | --- |
| `duckduckgo` | DuckDuckGo HTML | 不要 | **第三手**：零配置抓取（抓结果页解析）；出口被墙时配 `search.proxy` |
| `searxng` | 自建 SearXNG | 不要（要地址） | **第四手**：自己的实例，免 Key |
| `seam` | 宿主 `ctx.web` | 不要 | **第二手**：宿主自带公网地址校验 + **地址钉死在连接上**、同源跳转、体积与超时上限、代理路由 |
| `deepseek-api` | DeepSeek Search API（直连） | 要 | **默认第一**：`POST /v1/tools/search`，一次请求直接拿结果、不跑辅助模型，结果自带摘要；Key 取 `DEEPSEEK_API_KEY`（`baseURL` 默认 `https://api.deepseek.com/v1/tools`） |
| `openai` | OpenAI 兼容 `chat/completions` | 要 | 用 `web_search_options` 触发联网 |
| `tavily` | Tavily Search API | 要 | 结果自带摘要片段 |
| `brave` | Brave Search API | 要 | 独立索引 |
| `bocha` | 博查 BochaAI | 要 | **中文长尾覆盖好、国内直连** |
| `exa` | Exa | 要 | 语义检索（找观点 / 论文） |
| `serper` | Serper | 要 | 就是 Google 那一页的有机结果 |

**怎么选**：默认 `search.provider: auto` —— 按 `search.providerOrder` 依次试，
**第一条成功的就用它**。第六十五轮起默认顺序是 `[deepseek-api, seam, duckduckgo, searxng]`：
先走新 Search API 直连（快、不跑辅助模型），宿主接缝第二手，免 Key 抓取兜底；
**其他三方检索 API 不进默认顺序**——要用就显式点名（`provider: bocha`），
或把它们加回设置页的 providerOrder。想按次点名就用
`office_search_run({ queries: [...], provider: 'bocha' })`。
全部失败时错误里会把**每条通道的原因**分别写出来（配置缺失 / 网络不可达 / 被挡 / 没结果），
而不是压成一句「查不到」。

Key 的解析顺序：配置里的字面量 → 宿主凭据服务 → 进程环境变量 →
`$DSH_HOME/.credentials.yaml`。

```yaml
search:
  provider: auto
  providerOrder: [deepseek-api, seam, duckduckgo, searxng]   # 第六十五轮起：直连 API 在前
  providers:                  # 每条通道只留它用得上的键
    deepseek-api: { apiKey: '', apiKeyEnv: DEEPSEEK_API_KEY, baseURL: https://api.deepseek.com/v1/tools, timeoutMs: 30000 }
    tavily:    { apiKey: '', apiKeyEnv: TAVILY_API_KEY, timeoutMs: 30000 }
    searxng:   { baseURL: http://127.0.0.1:8080, timeoutMs: 30000 }
    duckduckgo: { baseURL: https://html.duckduckgo.com/html, timeoutMs: 30000 }
    # seam 不吃参数
  fetch: { maxResults: 8, fetchPages: 2 }   # 取正文与 PDF 的上限（与通道无关）
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

### 来源是 PDF 时交给 `office.pdf`（第二十九轮）

检索回来的来源里 PDF 很常见（学生题库、期刊、机构站点），此前这一律报
`不支持的内容类型：application/pdf`，等于把一整类可用的材料挡在门外。现在：

- `content-type: application/pdf`（或缺席 / `application/octet-stream` 且 URL 以 `.pdf` 结尾）
  → 下载到系统临时目录 → 交给 `office.pdf` 的 `pdfText` 抽文本 → 正文进 `content`、
  页数与抽取引擎进 `pdf`，`kind` 是 `'pdf'`，**不做网页预处理**（PDF 文本没有导航可洗）；
- 体积上限（`pdfMaxBytes`，默认 24 MB）是**放弃线**：超了直接报
  `OFFICE_WEB_PDF_TOO_BIG` 并提示换来源或先下载到本地 —— 截断的 PDF 抽不出文本；
- 抽不出文本（扫描件 / 手写）报 `OFFICE_WEB_EMPTY`，并把 `office.pdf` 那句提示带出去
  （渲染成页面图再 `read_image`）；没有抽取引擎（`pdftotext` / `PyMuPDF`）报
  `OFFICE_WEB_PDF_ENGINE`（归 `config`：装引擎，不是换来源）；
- 直查的结果文件与反馈里会多一行 `PDF 来源：N 条（正文由 office.pdf 抽文本，没做网页预处理）`。
- 落盘前先嗅 `%PDF-` 魔数：标着 PDF（或链接以 `.pdf` 结尾）而内容其实是 HTML 错误页时，
  报 `OFFICE_WEB_UNSUPPORTED_TYPE`（换来源），不会误导成「去装引擎」。
- `pdf.pages` 是**抽到了几页**（受 `pdfMaxPages` 夹），不是这份 PDF 的总页数 —— 长文档抽到
  上限时 notice 里的「N 页」要按这个理解。

`httpFetch` 的测试注入点是 `createWebAccess(ctx, options, { fetch, lookup, pdf })` ——
`pdf` 是一个 `{ text(path, opts) }` 替身，单测因此不必真的起 `pdftotext`。

### 出口代理：`search.proxy`（第二十九轮）

见上文「检索编排」里的 `search.proxy` 表：优先插件私有的 undici 分派器（只影响本插件、
改设置当场生效），拿不到 `undici` 才退回进程级环境变量那条路（整个进程走代理、只能装不能清）。
`src/web-proxy.js` 只做三件事：校验地址、装/换代理、报当前状态；`web.js` 的
`defaultNetwork.fetch` 负责把私有分派器通过 `dispatcher` 选项挂到这一次请求上。

### 直查：`office_search_run({ queries, out?, maxResults?, fetchPages?, provider?, preprocess? })`

只查一两个事实点时不必走提纲：给 1–5 条查询，插件去查、把来源与摘录写进
`.office/search/<slug>/run.md`，聊天里只回一份紧凑清单（标题 + URL）。
`provider` 按次点名一条通道（不传就按设置页的顺序自动挑），`preprocess` 按次改预处理强度。
清单是线索不是结论，写进文档前按 URL 核对。

### 渠道覆盖：`office_search_dispatch({ topic, type?, audience?, briefPath?, outputPaths? })`

给 `topic` 就自动按内容类型出提纲并落盘（`.office/search/<主题>/brief.md`，可审计、
可复用），随后逐渠道执行 —— 出提纲与执行两段在第七十轮合并成了一个工具。
`type` 取 `hotspot` / `knowledge` / `manual` / `mixed`；省略时按主题文字自动判断。
`briefPath` 复用上一轮落盘的提纲（改渠道或重跑时用），`outputPaths` 不给就按渠道数自动生成。

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

#### 执行通道

按提纲逐渠道执行（自己给 `outputPaths` 时必须与渠道数一一对应）。两条通道：

- **子代理通道**（默认优先）：组合里有子代理、**且工具面里真有那几个联网工具**时，
  **每个渠道一个子代理**，各自去调 `office_web_search` / `office_web_fetch`，
  把材料写进指定的结果文件。子代理只回一句「已写入 <路径>」，原文不进主上下文。
  通过 `ctx.get('subagents').start('spawn', …)` 起一次性子代理（检索有明确终点，
  等它写完再继续比后台可继续子代理好收口）；并发用简单闸门限制，默认最多 4 个。
- **内置通道**（自动回退）：没有 `subagents` 服务、或组合里缺那几个检索工具时，插件用内置检索
  在进程内跑同样的渠道 —— 每个渠道按提纲的查询去查、去重、必要时打开页面取一段正文当摘录，
  写成**同样格式**的结果文件，所以解析步不用改。

**「组合里缺检索工具」是装配期就知道的事实，不是运行时意外**（第十八轮 P1-11）。
第二十九轮之前的办公 preset 就是这种情况（`plugins` 列表里没有
`@deepseek-ai/dsh-tool-web`，而那时子代理白名单圈的是宿主那套 web 工具名），
`tools.restrict({ allow: CHANNEL_TOOLS })` **必然**被拒。
在此之前这件事是靠**每个渠道各撞一次失败**才发现的，那串原始报错还会贴在**每一个渠道**
的反馈行上 —— 模型据此写出过「本工作区会话没有联网检索工具」的 `[critical]` 记忆（session6），
下一轮工具修好之后那条记忆就是错的。判据是**运行时**读一次工具面；第三十轮起
子代理白名单换成了插件自己的抓取工具（`office_web_search` / `office_web_fetch`），
只要本插件在组合里，它们就在工具面上——探针自动认出「齐备」，无需改代码。

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

### 解析：`office_parse_findings({ paths, type?, topic?, draft? })`

把结果文件读成紧凑摘要。做五件事：

- **按渠道归类**结论，并算出提纲里哪些 ★ 渠道还没覆盖（「★ 仍缺渠道」）。
- **跨源核对**：把讲同一件事的结论合簇（并查集 + 字面重合度），看这一簇背后
  有几个互不相关的站点。这正是茧房效应最容易骗过人的地方——搜出来一堆结果，
  其实全是同一篇稿子的转载。
- **标出单一来源**与**完全没带 URL** 的条目。
- **URL 从正文里摘掉**，结论与出处分开存放，避免长链被抄进文档。
- 给了 `draft`（产物正文）时**做句级引用审计**：事实句没出处、或地址对不上本次
  材料，都当场点名且 `ok:false`（第五十七轮，见上面的「引用纪律」一节）。

### 子代理用什么工具

子代理**默认加入父方的组装**（也就是办公模式那一套，含 `office_run` 等），
所以派工时用 `toolFilter.allow` 把它裁成**极简面**——只有五个工具
（第三十轮起检索与取正文走插件自己的抓取工具，不用宿主的 web 工具族）：

```js
toolFilter: { allow: [
  // 检索（插件的免 Key 网页抓取通道）
  'office_web_search',
  // 取正文（PDF 由 office.pdf 抽文本）
  'office_web_fetch',
  // 读（含图片：图表 / 截图 / 扫描件）
  'read', 'read_image',
  // 写（唯一交付方式）
  'write',
] }
```

| 给了 | 为什么 |
| --- | --- |
| `office_web_search` | 免 Key 的网页搜索（抓 DuckDuckGo HTML 解析），渠道分流规则的落点；平台渠道用 `site:` 限定 |
| `office_web_fetch` | 打开具体页面拿原文，不只依赖搜索摘要 |
| `read` / `read_image` | 读文本与图片；`read_image` 让图表、截图、扫描件里的信息可读 |
| `write` | 把结果写盘，子代理唯一的交付方式 |

| 刻意不给 | 理由 |
| --- | --- |
| `edit` / `glob` / `grep` | 只写自己那个已知路径的结果文件，不需要改文件或满目录找文件 |
| `office_help` / `office_run` / `office_memory` | 检索子代理不生成文档、不读写记忆 |
| `bash` / `pwsh` | 不执行命令 |
| 子代理与任务板 | 不派活给更下一层 |
| 宿主的 `web_search` / `web_fetch` 等 | preset 不声明 tool-web：检索只走插件的抓取通道 |
| `present` / `todo` / `goal` / `skill` | 与「查资料、写文件」无关 |

用**白名单而不是黑名单**：以后 profile 里新装了什么工具，也不会悄悄出现在
检索子代理手上。`test/search.mjs` 里有一条测试把这个五项列表逐字钉住，
谁想加工具都会先看到它失败。

检索子代理**不是**一个可选择的模式：它由 `office_search_dispatch` 在派工时
临时创建，工具面由上面的 `toolFilter` 白名单决定。Agent 预设列表里只有
「办公模式」一项，不会多出一个「检索子代理」。

## 设置界面

装好插件后，设置对话框左侧只多出**一个**一级导航项「办公与记忆」，和
「Wallpaper Engine」同一层级、同样的交互形态。原来的「记忆系统」与「办公模式」
两页已合并；记忆设置不再占两个入口。一次性迁移不再占设置页，命令与工具调用仍
保留在 `office_help({ topic: 'memory' })` 的文档里。

### 八个页签（第一百零五轮，第一百零六轮加图库）

一屏只放 2–4 组，不用在几十行里找一格：

| 页签 | 分组 |
| --- | --- |
| 总览 | 模式 |
| 办公 | 插件、工具、守卫 |
| 检索 | 检索、子代理 |
| 文档 | 文档与缓存、OCR、Python、AV |
| 图库 | 图库视觉描述、图文对齐 |
| 记忆 | 记忆开关、三层记忆、作用范围、提示 |
| 知识库 | 知识库、知识库大模型 |
| 召回 | 容量、召回质量、每回合配额 |

页签选择记在 `localStorage` 的 `dsh-office-mode.settings-tab.v1`；下次打开回到
上次那一页，坏值退回总览。键盘走标准 tablist：← / → 循环切换，Home / End 跳两端；
只有当前页签在 Tab 顺序里，切换后焦点跟着走。

### 模式：一次改一组超参

总览是模式卡，不是指标卡。四张内置模式：

| 模式 | 向量通道 | 本地推理 | 筛选 | ANN / efSearch | 词法 / 向量 / 候选 | 召回分档 | 每回合配额 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 均衡 | auto | 关 | off | 开 / 64 | 0.55 / 0.45 / 30 | strict-v1 / ×5 / 6 / 4 | 1,1,2,1 |
| 轻量 | off | 关 | off | 关 / 64 | 1 / 0 / 20 | strict-v1 / ×3 / 3 / 1 | 1,1,1,1 |
| 本地语义 | auto | 开 | off | 开 / 96 | 0.5 / 0.5 / 50 | strict-v1 / ×5 / 6 / 4 | 2,2,3,1 |
| 深检索 | auto | 关 | off | 开 / 128 | 0.4 / 0.6 / 50 | strict-v1 / ×5 / 8 / 5 | 2,2,4,2 |

一张模式覆盖 16 项：向量通道、本地推理开关、筛选、ANN 开关与 efSearch、三格融合
参数、召回分档四项、每回合配额四项。端点、Key、模型目录不在其中 —— **切换模式
不会覆盖凭据或路径**。点卡片立即写回，状态行报「当前：X · 检索：Y」，Y 是按当前
设置现算的实际通道。

「保存当前参数」把当前 16 项存成一个命名模式（`localStorage` 键
`dsh-office-mode.hyper-presets.v1`）：最多 12 个、名字 ≤20 字，同名覆盖，内置名
不允许覆盖；自定义模式可以删除。

### 分组的展开 / 收起与「定位」

每个页签的分组默认展开（一页只有 2–4 组），点标题收起 / 展开；开合状态按页签
分别记在 `dsh-office-mode.groups.office`、`…groups.memory` 与 `…groups.gallery`，换页签互不影响。
收起只是 `display:none`（内容照渲染），输入框里的草稿不丢。

顶上那条框是**「定位」，不是搜索内容**：输入名称或配置项（「代理」「latex」
「阈值」「search.providers.deepseek-api.baseURL」都可以），命中的行只留它自己、
命中所在的分组自动展开、折在「高级参数」里的命中也会把折叠块展开，下面一行报
「匹配 N 项。」。多个词按**与**匹配（空格分开），大小写不敏感；分组标题或说明
本身命中时整组保留（搜「音频与视频」给整组）。清空即回到全量；一条都没命中时
明确说「没有匹配项。」。

合并页只注册一个设置命名空间 `dsh-office-mode`，由插件的**浏览器半侧**渲染：
`lib/client.js` 通过 `dsh.client` 声明被宿主加载，注册进 `settings.section` 槽位。
读写走 `ctx.configForms`（0.1.7 起由 dsh-client-ui-settings 提供的基础服务：按插件
行 id 取一份 ConfigForm，写入带回 revision 栅栏），服务端 schema 仍是唯一真源 ——
前端只负责展示与写回；各页签共用一套读写接线（`SettingsPanel`），只读提示与错误
显示因此不会各自漂移。

### 文字

设置项说明只讲「这一项做什么」与「关掉后的直接后果」，不写「其余功能不受影响」
这类废话。例如 OCR 的启用说明是「关掉后「识别文件」提交会报错，提示到设置里
打开。」，不再补一句其余功能照常。

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

### 办公模式插件：把 preset 里的插件行做成开关（第六十三轮 新-21）

「办公模式」页的第一组叫**「办公模式插件」**：它列的是**办公 preset 自己声明的插件行**
（`presets/office/cordis.patch.yml` 里 `config.plugins` 那一串，21 行 —— persona、工具行、
禁用行都在内），每行一个开关。这一组读的是**磁盘上的 preset 文件本身**（不是当前会话的
设置），所以：

| 事实 | 说明 |
| --- | --- |
| 开关改什么 | 只改那一行的 `disabled:`，其余字节原样（注释、顺序都不动）。YAML 不能整份重排 —— 重排会把设计记载一起改掉，diff 也就没人看得懂 |
| 写哪两处 | ① preset 真源 `presets/office/cordis.patch.yml`（profile 里的 `@local/dsh-office-preset` 是 `link:` 指向它，「安装目录」与「工作区」是同一份文件）；② profile 补丁 `~/.dsh/profiles/<profile>/cordis.patch.yml` 里**已经存在**的那个 id 的覆盖行 —— profile 层在 bundle 层之后应用，不一起改的话，preset 里写的 `disabled` 会被盖掉 |
| 不做什么 | **绝不新建** profile 覆盖行：凭空给一行加覆盖会把「办公模式不要这个能力」变成「所有模式都不要」，那不是按开关的本意（界面上的行说明会写清「profile 侧有没有覆盖行」） |
| 生效范围 | **新建的办公会话**。preset 在新建会话时读取，已经在跑的会话不受影响 —— 这正是需求里那句「对已经启用的对话无影响」 |
| 退路 | 每次真正写盘前先落一份快照（两份文件的原文）到 `$DSH_HOME/.office/preset-backups/`，**快照写不成就不动文件**；页面上的**「还原上一次备份」**把两个文件一起写回。写入是「临时文件 + 原子改名」，第二处失败时第一处回滚 |
| 为什么会失败得有原因 | 未知行 id / 非法值 / 非 JSON 请求体当场拒绝并给下一步；列表与 preset 脱节时提示「刷新一次再试」 |

列表是**实时**的：用 git 或手改动了 preset 文件，点一次「刷新」就会跟着变（界面上那行
「写到：…」会把两个真实路径摆出来，方便直接去改文件）。

> 已知边界：这一组只能改 preset 自己声明的行。profile 层那些「本来就在全局层」的能力
> （命令行、后台任务、子代理……）不在这里的列表里 —— 它们要真被掩掉，走的是
> 「办公模式的工具面守卫」（`guard.hideDevToolsInOffice`，见上一节）。

### 工具开关

| 开关 | 关掉后 |
| --- | --- |
| `office_help` | 模型不再能按需查写法（不建议关，会显著增加写错 API 的概率；记忆投影也挂在这个返回值上） |
| `office_run` | 不注册批量执行工具，办公模式退化成纯检索/写作 |
| `office_memory` | 不注册记忆工具（`office_run` 的自动台账仍按下面的 `autoLedger` 走） |
| `office_search_run` | 不注册内置检索直查，只能走渠道覆盖检索 |
| `office_search_dispatch` | 不出提纲也不派工，只能直查或手工检索 |
| `office_parse_findings` | 不解析结果文件，需要自己读文件 |

**关掉的工具不会出现在模型面前，也不占每次请求的 schema 开销。** 开关是
**即时生效**的：改完立刻重建工具面，不用重启宿主。

### 记忆（在「办公与记忆 → 记忆」页签）

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `enabled` | 开 | 总开关：关掉后记忆工具、提示段与自动台账一起停 |
| `dir` | `.office/memory` | 记忆目录（相对会话工作目录）；一个项目一份记忆 |
| `autoLedger` | 开 | `office_run` 写出文档时自动登记台账；关掉仍保留记忆投影 |
| `promptHint` | 开 | 往系统提示加一段静态说明（不含记忆内容）；办公模式下会被 persona 丢掉 |
| `userLimitBytes` | 4096 | 热记忆里「用户偏好」的容量上限，超出时下沉到归档 |
| `projectLimitBytes` | 10240 | 热记忆里「项目与环境」的容量上限，同上 |
| `projectionBudgetBytes` | 4096 | 一次投影里条目正文的字节上限（0 = 不限）。超出的条目不丢：折叠行报条数与读取入口 |
| `entryLimitBytes` | 2800 | 单条记忆的字节上限（0 = 不限）。超过就拒绝写入并给压缩提示：热记忆是动笔前照办的清单，不是文档仓库 |
| `entryHintBytes` | 1200 | 单条记忆的建议线（0 = 关掉）。超过只在写入回执里提醒，不拒绝 |
| `ledgerLimit` | 500 | 台账最多保留多少条，超出的最旧记录滚成月度归档摘要 |
| `archiveKeep` | 60 | 归档最多保留多少个摘要文件（归档自身也是有界的） |
| `layers.hot` / `.ledger` / `.archive` | 全开 | 层拓扑开关。关掉是「不再读、不再写」，不删数据 |
| `scope` | `workspace` | 记忆范围：`workspace` / `global` / `both`（跨项目层） |
| `globalDir` | 空 | 全局层目录；留空 = `$DSH_HOME/.office/memory`（相对路径相对用户主目录）。kb 在其兄弟目录 |
| `kbDir` | 空 | 知识库目录（单库，文档 / SQLite 索引 / 下载的模型都在这里）；留空 = `$DSH_HOME/.office/kb`（相对路径相对用户主目录） |
| `embedding.mode` | `auto` | `auto` = 本地推理优先、其次远端，都没有就纯词法；`remote` = 只走远端；`local` = 本地推理，没开才用本地哈希；`off` = 只用词法 |
| `embedding.baseURL` / `.model` / `.apiKeyEnv` | 空 | OpenAI 兼容 `/embeddings` 端点、模型、Key 环境变量；端点/模型不齐时该通道不生效 |
| `embedding.dimensions` / `.batchSize` / `.timeoutMs` | 0 / 16 / 15000 | 维度（0 = 模型默认）/ 批量 / 超时（失败或超时退回词法） |
| `embedding.local.enabled` | 关 | 按需拉起本机 Python 推理服务；第一次检索才启动，空闲自动退出 |
| `embedding.local.command` / `.modelDir` | 空 | 解释器（留空依次找 python / py / python3）/ 模型目录（留空 = 知识库目录下的 bge-small-zh-v1.5） |
| `embedding.local.idleTimeoutMs` / `.startupTimeoutMs` | 120000 / 60000 | 空闲多久退出 / 启动多久没就绪算失败 |
| `embedding.local.maxRssMb` / `.threads` / `.maxLength` | 600 / 2 / 256 | 峰值内存硬上限 / 推理线程 / 单块最长 token |
| `selector.mode` | `off` | `off` 只按融合分返回；`llm` 把前 N 条候选的编号 + 块头发给独立模型，只让它回编号 |
| `selector.baseURL` / `.model` / `.apiKeyEnv` / `.timeoutMs` | 空 / 空 / 空 / 12000 | 筛选模型端点、模型、Key 环境变量、超时；端点与模型必填，留空不筛选 |
| `selector.maxCandidates` / `.maxSelected` / `.headChars` | 12 / 5 / 300 | 送进筛选的候选数 / 最多保留几条 / 每块只发前多少字 |
| `ann.enabled` | 开 | 纯 JS HNSW 索引；缺失/损坏自动重建，超规模退回精确扫描 |
| `ann.efSearch` / `.m` / `.efConstruction` / `.maxBuild` | 64 / 16 / 200 / 20000 | 检索探测宽度 / 每节点邻居数 / 建图宽度 / 超过多少块不建图 |
| `hybrid.lexicalWeight` / `.vectorWeight` / `.candidateLimit` | 0.55 / 0.45 / 30 | 融合权重与每路候选池 |
| `userScope` | `memory` | `global` 时「用户偏好」始终落全局层 |
| `links` | 开 | 图关系：`link` / `related` / 实体视图 |
| `autoCapture` | 开 | 主动记录**指引**（不是自主记录器） |
| `recallQuality.policy` | `strict-v1` | `strict-v1` 分档过滤 / `off` 不丢结果 |
| `recallQuality.lowScoreThreshold` | 0.25 | 低于它算「未知」档（分数是命中词占比，不是余弦相似度） |
| `recallQuality.highScoreThreshold` | 0.6 | 达到它直接采纳；必须大于下限，写反了整组退回默认 |
| `recallQuality.candidateMultiplier` | 5 | 先取「条数上限 × 倍数」个候选再分档 |
| `recallQuality.maxMediumResults` | 6 | 「中档」最多采纳几条 |
| `recallQuality.maxUnknownResults` | 4 | 「未知档」最多采纳几条（0 = 一条不要） |
| `quota.recallPerTurn` | 1 | 一个回合里第一次带 `query` 的检索（0 = 不限制） |
| `quota.recallRefinePerTurn` | 1 | 同一回合里后续的换词细化 |
| `quota.kbSearchPerTurn` | 2 | 一个回合里的知识库（`kb`）检索次数（词法检索单独一档） |
| `quota.relatedPerTurn` | 1 | 一个回合里的图关系遍历 |

### 图库（在「办公与记忆 → 图库」页签）

两组都默认关；关着只显示一句说明，不渲染其余输入框。

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `gallery.vision.channel` | `off` | `off` 不生成描述；`host` 用宿主已配置的多模态模型；`openai` 用兼容端点 |
| `gallery.vision.provider` / `.model` | 空 | host 通道必填：宿主里的 provider 与模型 id |
| `gallery.vision.baseURL` / `.model` | 空 | openai 通道必填：端点（到 `/v1`）与视觉模型 id |
| `gallery.vision.apiKeyEnv` / `.apiKey` | 空 | Key 环境变量 / 直接填 Key；两者都没有时明确报未配置 |
| `gallery.vision.maxTokens` / `.timeoutMs` | 900 / 60000 | 单次描述的输出上限与超时（超时按失败回执，不写占位描述） |
| `gallery.vision.language` | `zh` | 只影响模型写出的描述文字 |
| `gallery.imageEmbedding.provider` | `none` | `none` / `voyage` / `cohere` / `vllm`（OpenAI 兼容多模态端点） |
| `gallery.imageEmbedding.baseURL` / `.model` | 空 | vllm 必填；官方通道留空用默认端点/模型 |
| `gallery.imageEmbedding.apiKeyEnv` / `.apiKey` | 空 | 同上；自建端点通常不需要 |
| `gallery.imageEmbedding.dimensions` / `.batchSize` / `.timeoutMs` | 0 / 4 / 30000 | 期望维度（0 = 模型默认）/ 每批图片数 / 超时 |

### 后台任务 Agent 的模型路由

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `mode` | `inherit` | `inherit` 跟随主会话；`fixed` 用下面指定的模型 |
| `provider` | 空 | `fixed` 时的 provider 路由名；留空 = 只覆盖 `model` |
| `model` | 空 | `fixed` 时的模型 id；留空按 `inherit` 处理 |

只影响渠道覆盖检索的**派工子代理**，不影响主对话。覆盖需要子代理后端声明
`SubagentCapabilities.agentOptions`；不支持时 `start` 会**响亮地拒绝**，不会静默
退回继承。

### 检索编排

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `engine` | `auto` | `auto` = 先派子代理、跑不了改用内置检索；`subagent` = 只用子代理；`builtin` = 只用内置检索 |
| `maxParallel` | 4 | 同时铺开几个检索子代理（1–8），渠道多时别一次全开 |
| `resultLimit` | 40 | 解析结果最多列几条结论（1–40） |
| `maxChannels` | 12 | 一次提纲最多几个渠道（1–12） |
| `outputDir` | `.office/search` | 提纲与结果文件目录 |
| `provider` | `auto` | 用哪条联网通道：`auto` 按 `providerOrder` 依次试；也可以固定成某一条 id（见上表） |
| `providerOrder` | `deepseek-api → seam → duckduckgo → searxng` | `auto` 的尝试顺序（设置页里点标签加/移） |

`search.providers.<id>` 是每条通道各自的参数（**键集就是它真正用得上的那几个**）：

| 通道 | 参数 |
| --- | --- |
| `seam` | 无（空对象） |
| `deepseek-api` / `openai` | `apiKey` / `apiKeyEnv` / `baseURL` / `model` / `timeoutMs` |
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

`search.fetch.*` 是**取正文与 PDF 的上限**（与用哪条检索通道无关；第六十五轮从旧 `search.builtin`
里拆出来，只保留「取正文」这一半）：

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `maxResults` | 8 | 每个渠道 / 每条直查查询最多几条来源 |
| `fetchPages` | 2 | 每个渠道 / 每条查询最多打开几页取正文摘录（0 = 只用搜索片段） |
| `fetchTimeoutMs` | 20000 | 单页取正文超时 |
| `maxBytes` | 2097152 | 单页体积上限（超出即截断） |
| `maxChars` | 20000 | 单页字符上限 |
| `maxRedirects` | 3 | 取正文最多跟随几次**同源**跳转（跨站跳转一律不跟） |
| `pdfMaxBytes` | 25165824（24 MB） | 取正文撞上 PDF 时的体积上限。**超上限是放弃而不是截断**（截断的 PDF 抽不出文本）：换来源，或先下载到本地用 `office.pdf` 读 |
| `pdfMaxPages` | 30 | PDF 抽文本最多抽前几页，用来圈住抽取耗时 |

`search.proxy` 是**出口代理**（第二十九轮新增，`src/web-proxy.js`）。留空 = 联网按本进程原有
出口直连；填了（形如 `http://127.0.0.1:7897`，只认 http / https）让**本插件自己发起**的请求走它：

| 路径 | 触发条件 | 作用范围与代价 |
| --- | --- | --- |
| 私有分派器（首选） | 能从宿主安装位置取到 `undici`（DSH 自带） | 只有本插件的联网走代理；宿主 web 服务与别的插件不受影响；改值 / 清空**当场生效**；回环地址（自建 SearXNG 的 `127.0.0.1:8080`）由**请求侧**绕开 —— 私有分派器不吃 `NO_PROXY`，`web.js` 的 `defaultNetwork.fetch` 对回环地址不挂 `dispatcher` |
| 进程级环境变量（兜底） | 取不到 `undici`，且运行时有 `http.setGlobalProxyFromEnv()`（Node 24+） | 写 `HTTP_PROXY` / `HTTPS_PROXY` 并装全局 dispatcher：**整个进程**都走代理，且**只能装不能清**——清空设置要重启 DSH 才回直连；`NO_PROXY` 自动补回环地址。两个副作用要知道：① 它会**顶掉**宿主启动时按环境变量装好的那份全局 dispatcher；② 代理地址（含 `user:pass@`）会进 `process.env`，宿主拉起的子进程（bash / pwsh / python）会继承 |
| 不生效（报错） | 两条都不满足，或地址不是合法 http(s) | 返回 `ok:false` 与原因（**不回显凭据**，非法地址会打码），工具反馈照实说，不假装设置生效了 |

**可配的代理不覆盖接缝**：默认顺序里 `deepseek-api` 排在最前、`seam`（宿主 `ctx.web`）第二，
取正文也先试 `seam.fetch` —— 所以配了代理时，宿主接缝与直连 API 那两条路径**可能不经过它**，
只有免 Key 抓取通道与「接缝失败后的兜底」才走代理。要让宿主那一侧也走代理，得在**启动 DSH 前**
设 `HTTPS_PROXY`（宿主 `@deepseek-ai/dsh-http-proxy` 在 boot 时按环境变量装全局 dispatcher）。

`office_help({ topic: 'search' })` 的末尾会给一行**运行期实况**（配了什么、走的是哪条路径、
装成没装），设置页改了立刻能看见。

环境变量覆盖：`DSH_OFFICE_SEARCH_BASE_URL` / `DSH_OFFICE_SEARCH_MODEL` /
`DSH_OFFICE_SEARCH_API_KEY_ENV`，端点也认宿主的 `DEEPSEEK_SEARCH_BASE_URL`。

### 文档与缓存

`defaultTheme`（下拉选主题）、`scriptTimeoutMs`、`maxScriptChars`、`cacheDir`、
`keepCache`。

### Python 计算与绘图

设置页「办公模式 → Python 计算与绘图」一组里能改全部四项（第三十八轮补的界面；
在此之前只能改配置文件）：

`enabled`（默认开；关掉后 `office.python.*` 明确报错并说明在哪儿打开）、`bin`
（解释器，留空自动探测，也可填绝对路径）、`timeoutMs`（默认 120000，单次运行上限）、
`outDir`（产物目录，相对缓存目录，默认 `python/out`）。

### 音频与视频

`enabled`（默认开）、`ffmpegPath` / `ffprobePath`（留空自动探测；**填了就以它为准**，
路径不存在即报缺件）、`modelDir`（SenseVoice 模型目录，留空 = 语音输入下载的那份）、
`language`（`auto` / `zh` / `yue` / `en` / `ja` / `ko`）、`chunkSeconds`（默认 120）、
`overlapSeconds`（相邻块的重叠秒数，默认 0.5，上限半块）、
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
| 目录 / 图片墙 / 阶段路线 / 四象限（第六十八轮） | `deck.agenda({title, items:[{text, desc, active}]})`（编号目录，当前章高亮）、`deck.gallery({items:[{path, caption, badge}], columns:2, ratio:0.62})`（等格图片墙）、`deck.roadmap({phases:[{label, items:[…], progress:0.6}]})`（1–5 阶段横排）、`deck.matrix({xAxis:{low:'低',high:'高'}, yAxis:{low,high}, items:[{x:0.2, y:0.8, text}]})`（2×2 象限，坐标 0–1、原点左下） |
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

## 写作：先规划，再落笔（第七十三轮）

用户这一轮的要求有两层：**文字生成要更像人**，以及**写作规划不能只有结构、要合适** —— 偏学术、困难、跃迁性的内容多花笔墨，说明性内容可以用日常比喻和放松的语句。还有一个硬要求：**能力要外带，不能只靠提示词，并且可以进行大模型测试**。

落点是三块：`src/writing.js` 的确定性规划 / 自检、SDK 的 `office.writing` 五个入口、`scripts/writing-eval.mjs` 的 A/B 评测。没有新增 MCP 工具 —— 工具面已经只剩几百字节余量，而写作是「动笔前查一次」的能力，适合走 `office_help` + SDK。

### 规划：按难度与读者收益分注意力

`office.writing.plan(input, options)` 返回的不是六个小标题，而是一张注意力分配表：

- `input`：一句话 brief，或 `{ title, genre, audience, goal, targetChars, throughLine, sections }`。
- `genre`：academic / explainer / report / memo / essay / howto / general（也认论文、科普、汇报、提案、评论、指南等中文别名）。
- `sections` 每项：节名字符串，或 `{ heading, kind, difficulty, value, purpose, mustCover, moves, register, figurative, rhythm, avoid }`；缺省时按 genre 给一套可改骨架。
- `kind`：context / problem / method / theory / evidence / analysis / leap / implication / example / howto / limitation / summary / faq。节名会按关键词猜（「机制」→ theory，「局限」→ limitation），猜错就显式传。
- 权重：`base × (1 + difficulty×0.16) × (1 + value×0.10)`，`difficulty` / `value` 都是 0–3；单节上限按节数（1 节 100% / 2 节 65% / 3 节 50% / 4 节及以上 40%），目标字数精确加总到 `targetChars`。
- 深度：`share ≥ 0.25` 是 dwell（写透），`≥ 0.15` 是 develop（展开），`≥ 0.08` 是 touch（点到），否则 compress（收住）。

每一节拿到：目标字数、深度档、目的、读者问题、必须写到、四段动作（先 / 展开 / 转折 / 落点）、语言登记（准确克制 / 平实 / 放松 / 张弛交替）、比喻边界与节奏。`plan.checks` 是写前确认，固定三问：

1. 读者要同时拿住几个新概念？超过 2 个就拆步、给例子。
2. 这一节 / 这段话改变读者哪个判断？不改变就删，或压成一句。
3. 全文最长的一节是不是最难的那点？若最长的是背景 / 常识，规划失败。

`plan.markdown` 可以直接 `office.log(plan.markdown)`，把篇幅表和逐节写法带回上下文。规划本身是草稿，不是交付物；用户要大纲时才给。

### 自检：可解释的信号，不是 AI 检测器

`office.writing.audit(text, plan?)` 做两件事：

- **不传 plan**：句长过匀 / 段长持平 / 套话密度 / 抽象名词过多 / 锚点不足 / 比喻过密 / 过度对冲 / 通篇无让步 / 每节总结 / 重复句首。
- **传 plan**：重点节低于目标 68% 报 `underdeveloped`，收住节超过 1.7 倍报 `overexpanded`，各节实际篇幅仍平均报 `uniform-allocation`，最长节不是最难 / 最值钱的节报 `longest-not-hardest`。

每条 warning 都带具体数字与改稿动作（例如「「机制」至少扩到 864 字：补机制链条、一个具体例子，或一个反例」）。`score` 只是启发式加权（≥82 自然、≥68 尚可、否则模板感重），**不是「像人程度」的证明**；改稿看 warnings 与 suggestions。

`office.writing.signals(text)` 只给表面量（句长 / 段长 / 套话 / 锚点 / 比喻 / 让步 / 问句），不判分；调阈值或接自己的评测时用。`office.writing.render(plan)` 把规划对象重新渲染成 Markdown。

### 去 AI 味：只报可解释的句式（第七十五轮）

`office.writing.deai(text, { mode, register, styleRef })` 是第七十四轮标定的落点：不给「AI 含量」分数，只列可解释的句式与参考稿偏离。

- `flags`：A1 对举（不是……而是）、A2 提示语冒号（原因不复杂：）、M1 元话语起手式（先说结论：说白了：）；每条带 spans / count / density / confidence / reason / falsePositiveHint。
- `register`：platform / official / academic / personal / general（也认中文）。official / academic 把 A2 降为提示，personal 把 A1 降为提示；F1 破折号、A4 长定语、A5 顿号、B1 抽象、C2 连接词只进 stats。
- `styleRef`：传一段人类参考稿，报句长、段长、对举 / 提示语 / 锚点 / 套话密度与参考稿的偏离，不报 AI 概率；参考稿不足 200 字只作方向。
- `mode:'safe'`：只删「先说结论：」等元话语前缀，把「原因不复杂：」改成「原因是」、「这个原因：」改成「这个原因，」；A1 只报不改，事实与论证不动。

标定结果（24 条盲标，3 体裁 × 2 长度 × 4 条件）也写进了 `docs/办公模式第七十四轮-标定结果与规则修正-20261107.md`：clean 组 4.33 > heavy 3.17 > 人类 1.33，说明「避开套话但写具体」不自动像人；A1/A2 是仅有的高精度信号，破折号 / 长定语 / 顿号在人类稿上也会命中。**它不是检测器，也不推断作者。**

### 写作口径：研究里能落地的那几条

调研按「机制 + 写法 + 来源」整理了十条，落进 plan / audit 的主要是：

- **篇幅随难度分配**：新定义、反直觉结论、机制与跃迁点给足；背景与常识压到 1–2 句。平均用力本身就是没做难度判断。
- **句长节奏**：一段内至少一次 ≤10 字短句落结论，或 ≥45 字长句装条件；连续三句字数差 <5 就重写。
- **具体与可核验**：每 150 字至少一个锚点（数字带口径 / 日期 / 专名 / 原话 / 具体场景）；把「显著提升」换成「0.62→0.71（n=1,204）」。
- **让步与反例**：每个核心主张配一条最强反对或竞争解释，并写清它在什么条件下成立；不要用「有人认为」设稻草人。
- **段落单一功能 + 已知→未知**：段首给主张，段内先旧信息后新信息，一段只推进一个信息差。
- **连接词有语义负荷**：靠语序、关键词重复、代词回指能接上就删；「首先 / 其次 / 最后」全文只许出现一组。
- **不确定性校准**：相关性写「与……相关（n=…）」，未验证机制写「一种解释是」，外推写「在……条件下」。
- **比喻的使用条件**：说明文可以用日常比喻，但每个都要说明「哪里像、哪里不像」，一段最多一个；学术论证段不用装饰性比喻。
- **删冗余，不每节总结**：只复述的总结等于承认前文没讲清；总结句必须带新条件、代价或下一步。

**明确不做**：不做「人类化改写器」（同义词替换、加错别字、插伪口语），不把「绕过检测器」当目标。检测器在跨域、改写与非母语写作上系统性失效（TOEFL 作文 61% 误判、14 个商用工具无一可靠），把「像人」简化成躲检测既学不到好写作，也不合规范。

### 大模型 A/B 测试

`scripts/writing-eval.mjs` 做 baseline 与 planned 两臂：

- 三个任务：A 学术论文难点（机制 / 跃迁，目标约 550 字）、B 面向普通读者的说明文（允许 1–2 个日常比喻，约 475 字）、C 工作汇报（结论 + 依据 + 风险，约 350 字）；三题都按目标 ±10% 限长。
- 两臂：baseline 普通写作提示；planned 先调 `writingPlan(...)`，把 `plan.markdown` 放进提示，再让模型写。
- 控制偏差：同一任务 × 每臂 × runs 次，两次运行交换臂顺序；按任务目标字数 ±10% 限长；样本按编号匿名，`key.json` 单独保存臂名。
- 记录：每份产出的 `writingAudit`（score / band / warnings / chars）、研究里的 10 条反模板 0/1 清单、长度与截断标记；`report.md` 汇总。
- 无 `WRITING_EVAL_API_KEY` 或 `DEEPSEEK_API_KEY` 时打印 `SKIP` 并 exit 0，不伪造结果；`--mock` 完全离线，`--self-test` 跑通流程。

评测不以检测器分数为目标；人工盲评用三维 1–5 分（人类自然度 / 规划合理性 / 内容可信度），每个锚点先看样例再打分，具体设计见交付档与调研记录。

### 第六十轮的篇幅表还在

第六十轮那条「动笔前先列篇幅表」仍然是底线，只是从「写在提示里的一句话」升级成可计算的 plan + 可回查的 audit：分几节、每节写什么、多少字、依据，重点给足、铺垫收住，用户给的总长以它为准，写完对表自查。

`guide` 的两层都写了这条：默认层一句话（先规划、重点给足、总长以用户要求为准、写完 audit 自查 —— 2,266 B），全文层给表的四要素、正反例与两条处置 —— 9,404 B；`injection-budget` 钉住「两层都有、字节有界」，`preset-check` 反向钉住「persona 里没有」。为什么不在 persona：它是 `complete: true` 的唯一提示段，按最坏能力地图渲染后 13,951 / 16,300 B（余量 2,349 B），而这是「动笔前查一次就够」的写法约定，不是每次请求都必须阻止的失败模式。

能机械验的与不能机械验的仍分开：`node test/injection-budget.mjs` 验指引、分层、正反例与字节；`node test/writing.mjs` 验规划分配、字段、自检信号与 SDK 挂载；模型照不照做仍只能靠 A/B 与盲评，不能由一次生成下结论。

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
缓存：.office/runtime/cache 保留 1 个中间文件（20 B）
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
  memory.js       三层记忆（热记忆 / 台账 / 归档）与 kb 读写：存储、写入协议、容量下沉、有界读取
  kb.js           知识库切块：结构切块 / span / 块头 / 词法 BM25
  kb-db.js        SQLite 索引：FTS / 向量 / 块边 / 摘要（拿不到 node:sqlite 时降级）
  kb-embed.js     embedding 通道（远端 OpenAI 兼容 / 本地进程 / 哈希兜底）
  kb-selector.js  候选筛选：只发编号与截断块头，只收编号
  kb-model.js     本地 bge-small-zh 嵌入模型下载器（只落盘，不做推理）
  kb-extract.js   多格式抽取：PDF / DOCX / PPTX / XLSX / 图片 → 文本块 + 图片资产
  sdk.js          脚本可见的 office 对象
  pdf.js          PDF：引擎探测 / info / text / pages（渲染结果进缓存，可命中）
  convert.js      文档互转与统一出 PDF：pandoc / TeX / LibreOffice 三条路的探测与调度（不打包引擎）
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
    cacheDir: .office/runtime/cache      # 中间产物目录，相对会话工作目录
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
      projectionBudgetBytes: 4096 # 一次投影里条目正文的字节上限（0 = 不限）
      entryLimitBytes: 2800      # 单条上限：超过就拒绝写入（0 = 不限）
      entryHintBytes: 1200       # 单条建议线：超过只在回执里提醒（0 = 关掉）
      ledgerLimit: 500           # 台账条数上限，超出的最旧记录滚进归档
      archiveKeep: 60            # 归档摘要文件数上限
```

环境变量兜底：`DSH_OFFICE_CACHE_DIR`、`DSH_OFFICE_THEME`、`DSH_OFFICE_PDF_ENGINE`、
`DSH_OFFICE_TEX_ENGINE`、`DSH_OFFICE_TEX_TEMPLATE`、`DSH_OFFICE_MEMORY_DIR`、
`DSH_OFFICE_PYTHON`（解释器）、`DSH_OFFICE_PYTHON_OUT`（产物目录）、
`DSH_OFFICE_FFMPEG`（ffmpeg 路径）、`DSH_OFFICE_SENSEVOICE_DIR`（模型目录）、
`DSH_OFFICE_PDF_BIN`（额外搜索引擎的目录，分号分隔）、
`DSH_OFFICE_PANDOC`（pandoc 可执行文件或命令名）、`DSH_OFFICE_CJK_FONT`（LaTeX 路线用的中文字体名）。

## 垃圾文件规则

办公插件在**工作区里产生的全部中间产物都收在一个隐藏目录 `.office/` 下**，
工作目录根上不会再多出第二个插件目录：

根目录只放长期数据，运行期文件统一在 `.office/runtime/` 下（`node scripts/organize-office.mjs`
可把旧版散落的目录迁进去；`node scripts/clean-office.mjs --older-than 7` 清运行期旧文件）：

| 目录 | 装什么 | 生命周期 |
| --- | --- | --- |
| `.office/memory/` | 三层记忆（热记忆 / 台账 / 归档） | **不是缓存**：跨会话保留，不按 TTL 清理 |
| `.office/kb/` | 知识库（manifest + 块文件） | 长期数据，不按 TTL 清理 |
| `.office/backups/` | 备份 | 长期数据 |
| `.office/search/` | 渠道覆盖检索的提纲与结果文件 | 跟着那次检索走，可随时删 |
| `.office/runtime/cache/` | `office.cache.write` 的中间文件、PDF 渲染页图、抽出来的 PPT 图片、Python 产物（图 / 数据 / 脚本与日志）、音频转写解码出来的规范 WAV 与视频抽帧图（`av/audio`、`av/frames`） | 跨调用保留，按 TTL（默认 12 小时）与容量（默认 512 MB）清理；也可整目录删 |
| `.office/runtime/tmp/` | 测试中间产物（每套件一个子目录） | 每次跑测试重新生成 |
| `.office/runtime/diag/` | 浏览器取证脚本与截图 | 排查用，可随时删 |
| `.office/runtime/labeling/` | 标注夹具 | 可随时删 |

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
node test/memory-recall.mjs  # 召回自测：三层额度与渲染 / 中文 hit@3 标注集 / 读侧分区（硬过滤与合取）
                             #   第五十七轮 D 组：统一入口路由的「记忆够答吗」门控四档（纯判据 + read 现算）
node test/citation.mjs       # 引用纪律（第五十七轮）：事实句词表 / 句号后的出处并回上一句 /
                             #   句级引用率与「地址对得上材料」两项算得出的指标 / correctness 照实标 null /
                             #   落盘断言（office_parse_findings 的 draft 没过时 ok:false 并逐句点名）
node test/kb.mjs             # 知识库（第四十一轮）：结构切块与 span 等式 / 幂等与换代 / 有界读取 / 删除的引用完整性 /
                             #   来源三档硬规则 / kbRefs 反查 / 冲突三态 / Pack 往返 / 跨根同名路径 /
                             #   词法检索（第四十五轮）：命中给块 id 与 span、弃答、筛选、有界、配额 /
                             #   换代与删除的两段式提交协议（第五十六轮）：manifest 才是提交点、崩在提交前/后
                             #   各留下什么样的盘面、意图日志 pending.json 的重放与丢弃、跨根清边不许丢写 /
                             #   第八十五轮：虚拟文件夹 / 编辑与块关系 / 混合检索标记 / 多格式图片入库
node test/kb-db.mjs          # SQLite 索引（第八十五轮）：schema / 增量同步 / 中文 FTS / 向量 / 两跳邻居 / 摘要
node test/kb-extract.mjs     # 多格式抽取（第八十五轮）：文本 / PDF / DOCX / PPTX / XLSX / 图片与大小上限
node test/view.mjs           # 记忆浏览端点：快照口径 / 读端点只读（405）/ 写端点 POST 落盘 / 只服务见过的根（404）
node test/pdf.mjs            # PDF：轻量解析 / 抽文本 / 渲染 / 缓存命中 / 缓存 TTL 与容量
node test/convert.mjs        # 文档互转与统一出 PDF（第五十三轮）：引擎探针与矩阵 / md↔docx↔tex 回环 /
                             #   内嵌图片抽取 / 幂等与覆盖闸 / 四条出 PDF 路线（含老格式与 OpenDocument）/
                             #   负例与 office_run 全链路（缺引擎的路线 SKIP 并如实说明）
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
                             #   固定长句一个会话只说一次 / office_run 不回灌模型自己写的 outline /
                             #   guide 两层的「先定篇幅」与字节上限（第六十轮 新-24）
node test/writing.mjs        # 写作规划、自检与去 AI 味（第七十三/七十五轮）：分配公式 / 深度阈值 / 目标字数精确加总 /
                             #   套话·句长·锚点·让步信号 / 对表查重点节与平均用力 / deai 的 register/styleRef/safe / SDK 五个入口（81 项）
node test/writing-eval.mjs   # 大模型写作 A/B（第七十三轮）：--mock 全流程 / 无 Key 时 SKIP / 127.0.0.1 假端点真 HTTP /
                             #   样本匿名、换序、限长、--judge 降级（122 项，不联网）
node test/e2e.mjs            # 端到端：走 office_run 真实路径产出三件套并复检（含自动台账）
node test/format-word.mjs    # 各格式模块自测（Word 2515 / Excel 322 / PPT 1500 项断言）
node test/format-excel.mjs
node test/format-ppt.mjs
node test/format-ppt-revise.mjs  # PPT 就地微调：读结构 / 改文字样式位置 / 改别人的包 / 插图与抽图（153 项）
node test/format-tex.mjs     # LaTeX 论文：生成项目 / 项目级复检 / 编辑 / 真实编译出 PDF（--no-compile 只测前三条）
node test/image-export.mjs   # PPT 图片导出（第七十二轮 S-05）：统一文件环境 / 显式 out 覆盖控制 / 缓存按内容哈希分桶（33 项）
node test/plain-default.mjs  # 默认素色网格：不传主题时 Excel/Word 没有任何背景填充
node test/ci-guard.mjs       # CI / 发布守门的名单核对（第五十六轮）：两份工作流只许调 --ci、
                             #   CI 档必须含核心格式回归且不含要本机环境/要出网的套件
node test/measure.mjs        # 测量口径（第十八轮 / 第四十七轮）：pass^k 外壳、等字节基线、
                             #   多格式扰动回归；第六十轮补一条「对照实验的证据固定」
                             #   （test/fixtures/r60-writing-plan/ 的产物哈希与逐节字数）
node test/rebuild.mjs        # 设置重建事务与显式降级态（第五十九轮 S-04）：装/拆分离、
                             #   失败回滚、连回滚也失败时的降级态、重入守卫

# 一条命令跑全部（51 套，逐套件 exit code + 末尾「N/M 套件 exit=0」；快的在前、要联网的在后）
node scripts/run-suites.mjs            # 也可以只跑其中几套：node scripts/run-suites.mjs kb memory client

# CI 与发布守门跑的那一批（当前 51 套；纯 Node、零依赖、不联网、不要本机引擎；名单只有一处真源：
# scripts/ci-suites.mjs —— 两份工作流都调这一条，改名单只改那一个文件）
node scripts/run-suites.mjs --ci

# 浏览器半侧的活体验证（第十轮新增，可复用）：起隔离实例 + 无头 Chrome，用 CDP 打开页面，
# 并用 Fetch 域拦截 /office-memory/snapshot 回一份合成快照，于是能验到「数据到位时的面板」
# 而不是只有空态。产物（截图 / web-verify.json）落在 .office/runtime/diag/ 下。
# 第十二轮扩了三处：令牌解析表（面板用到的 --dsw-* 逐个读运行时值）、聚焦交互
# （focus 一个节点 → 右侧详情栏）、深浅色与宽窄栅格（含容器查询是否真的塌成一列）。
# 第十三轮再扩一组**真实输入事件**：Input.dispatchMouseEvent 发滚轮 / 拖拽（缩放、平移、
# 拖节点都只有真事件才验得到，顺带证明页面没跟着滚）、点「再显示 N 条」、点「展开」、
# 点条目动作看待办指令、以及 Emulation.setEmulatedMedia 打开 prefers-reduced-motion
# 后读 .om-scroll 的计算样式。
# 注意：React 的状态更新是异步的 —— 点完不能在同一次 Runtime.evaluate 里就读 DOM。
node test/web-verify.mjs

# 测试中间产物统一落在仓库根的 .office/runtime/tmp/<suite>/ 下（不再用包内的 test/.tmp）
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
  切句是 Silero VAD 启发式，标点与断句由模型给；块与块之间不共享上下文（切点前后的
  用词可能不一致）。识别不出内容时不编造 —— `segments` 为空就是空。
- 单次转写时长上限默认 3600 秒（`av.maxSeconds`），单块 120 秒（`av.chunkSeconds`）、
  块间重叠 0.5 秒（`av.overlapSeconds`）：分块解决单块大小，**不改变时长上限**，超了
  先切段。切点不会把一句话切成两半 —— 重叠盖住切点附近的短音，更深的跨界句子由边界
  回退整句重解（`rolls` 记录回退次数），重叠区不会被两块各报一次。抽帧单次上限默认
  12 张（`av.maxFrames`），4K 视频建议给 `width`（图小、读得快）。
- 抽出来的帧是**中间产物**（`.office/runtime/cache/av/frames/`），不进交付物；要留在工作目录
  就用 `office.files.copy(帧路径, '工作目录里的名字')` 搬过去，或者直接把帧当图片嵌进 Word / PPT。
- **原生图表**（`11-6` 已交付）：Excel 用 `sheet.chart({type, title, categories, series, ...})`、
  PPT 用 `deck.chart({...})` 生成真正的 DrawingML 图表部件（`xl/charts/chartN.xml` /
  `ppt/charts/chartN.xml`）—— 在 Office 里是可选中、可改类型、可「编辑数据」的图表对象。
  数据以 `c:numLit` / `c:strLit` **内联**在图表里（不引用单元格、不嵌工作簿），所以打开即画；
  要改数据走 Office 的「编辑数据」。读回时 `stats.charts` 与 outline 都会报出来。
  Python 画 PNG 仍然有用：示意图、以及 Office 画不出来的形状。
- PPT 的**微调**是「改形状」，不是「改结构」：`office.ppt.revise` 能改文字、字体字号颜色、
  位置尺寸与文本框高度，也能**插入图片**（`addImage` / `office.ppt.insertImages`），
  但**不能增删整页、换版式、改表格结构** —— 那几件事请重新生成，或直接在 PowerPoint 里做。
- 形状几何要求带 `a:xfrm`：**图片（`p:pic`）是可以 move / resize 的**（它自带
  `p:spPr/a:xfrm`，`readSlides` 会把 `editable` 报成 `false` 只是因为图片没有文本可改，
  不代表几何改不了）。组合形状（`p:grpSp`）与表格 / 图表（`p:graphicFrame`）的几何不在
  形状自己的 `spPr` 上，改它们会进 `skipped` 并说明原因，不会假装成功。
- 同一页命中多个同名形状时全部处理并在 `warnings` 里提示；要精确到某一个请用 `readSlides` 拿到的 id。
- 原生图表的两条边界：① 数据是内联的（改数据要在 Office 里「编辑数据」，脚本里改不了
  已有图表的数值）；② 就地编辑（`office.excel.edit`）会**原样保留**图表部件与锚点，
  但不会重算图表数据 —— 编辑结果里会有一条 warning 说明这件事。
- PPT 自绘形状支持 32 个 preset；`flowChartData` / `roundedRectCallout` / `ovalCallout` 是口语名，
  会被映射到 DrawingML 的真名（写真名会让真实 PowerPoint 打不开文件，见 docs/ooxml-pitfalls.md）。
  连续调用 `deck.shape/line/icon` 画在同一页，用 `deck.page()` 或任一版式方法另起一页。
- 与 Excel/Word 一致：`fill` 为空串或 `FFFFFF` 视为**不填充**，纯白卡片会透明；
  要白卡片用 `F8FAFC` 这类极浅灰。
- 封面与结束页使用主题的封面底色（彩色主题是 `primaryDark` 深底白字，素色主题是白底黑字 +
  浅灰色带），可用主题里的 `cover` 覆盖项改。
- 画布通道（第五十二轮起）：版式助手页、自由绘制页与内容版式页（`bullets` / `table` / `quote` /
  `statement` / `image` / `images` / `chart`）共用同一条通道 —— `deck.bullets()` 之后调用
  `deck.banner()/panel()/shape()` 会续在同一页，只有 `cover` / `section` / `closing` 这三类
  整页构图版式不开放，接在它们后面的元素级助手会另起一页（`read()` 会报出这类页）。
- `deck.image()` 能与版式同页（第五十二轮）：接在**版式助手页**后面时不给坐标就自动分栏 ——
  图占版心右侧（`ratio` 默认 0.42），助手排版收到左栏；给了 `x/y/w/h`（厘米）就按给的摆位画。
  **一页只自动分栏一次**：同一页上第二张不给坐标的配图会另起一页（否则会落进收窄后的版心、
  压住助手内容）。接在其它版式页后面时，只有给了 `x/y/w/h` 才会同页；不给坐标照旧
  「一次调用一页」（`deck.page()` 也能显式换页），`fullBleed: true` 永远独占一页。
  同页配图没有标题带：给了 `title` 会在 `warnings` 里说明已忽略；`widthCm` / `heightCm` 照收，
  裁切与低 DPI 也照报（与单图页同一套尺寸规则，`save()` 与 `read()` 报同一个数）。
  图与题注计入 `stats.images` 与 `read().pages[i].images`，**缺图占位框例外**
  （它是 `p:sp` 不是 `p:pic`，只有生成端数得到）。
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
- 记忆面板（会话页签行里的「记忆」）**本身不写记忆**：条目上的动作只生成一条精确的
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