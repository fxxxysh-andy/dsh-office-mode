/**
 * 办公模式设置页（浏览器半侧）。
 *
 * 这一页注册成设置对话框里的一级导航项（settings.section），和「记忆系统」
 * 「Wallpaper Engine」同一层级、同样的交互形态：左边一个带图标的入口，
 * 右边是可操作的控件（开关、下拉、数字输入）。
 *
 * 为什么不只用服务端的 settings schema：schema 只描述「有哪些字段」，
 * 渲染成什么样、分组怎么排、什么时候灰掉，都需要前端说了算。
 * 服务端 schema 仍然是唯一真源 —— 这一页只做展示与写回，不自己存值。
 *
 * 读写都走 ctx.configForms.get(插件行 id)（dsh-client-ui-settings 提供的设置
 * 传输，0.1.7 起取代了 ctx.settingsScope）：读来自共享 describe mirror，写带回
 * revision 栅栏，所以多个页签同时改也不会互相覆盖。写入只允许落在 volatile
 * 字段上 —— 服务端 schema 里已经整片标好（见 src/settings.js）。
 *
 * 第三十五轮起，两页的骨架都是**可折叠的分组**（默认全部收起，进页面先看到一份
 * 分组目录），顶上一条「定位」框按名称 / 配置项过滤。细节见下面「分组折叠与
 * 定位」那一段与 SettingsGroup / SettingsBody 的说明。
 */
window.__ModuleLoader__.load({
    id: "dsh-office-mode",
    factory: (require) => {
        var module = { exports: {} };
        var exports = module.exports;
        Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });

        const React = require("react");
        const h = React.createElement;

        const NAMESPACE = "dsh-office-mode";
        /**
         * 记忆系统面板的 section id。
         *
         * 与命名空间分开：命名空间是「设置存哪儿」（一个），section id 是「导航里
         * 哪一格」（两格）。复用命名空间当第二个 section 的 id 会把两页塞进同一格。
         */
        const MEMORY_SECTION_ID = "dsh-office-memory";

        /** 工具开关的显示文案。顺序即界面顺序。 */
        const TOOL_ROWS = [
            { key: "office_help", label: "office_help", note: "按需查 Word / Excel / PPT 的写法与主题。关掉会明显增加写错 API 的概率" },
            { key: "office_run", label: "office_run", note: "批量执行：生成或修改三件套，以及批量改文本文件" },
            { key: "office_memory", label: "office_memory", note: "三层记忆：读写热记忆（偏好与约定）、检索台账与归档" },
            { key: "office_web_search", label: "office_web_search", note: "查一个事实点（网页抓取通道，免 Key）：来源直接回到对话，不落盘" },
            { key: "office_web_fetch", label: "office_web_fetch", note: "打开一个页面取正文（网页抓取通道）；PDF 由 office.pdf 抽文本" },
            { key: "office_search_run", label: "office_search_run", note: "内置检索直查：给几条查询，插件的联网通道去找来源并落盘" },
            { key: "office_search_brief", label: "office_search_brief", note: "检索第一步：按内容类型出检索提纲" },
            { key: "office_search_dispatch", label: "office_search_dispatch", note: "检索执行：有子代理就派子代理，没有就用内置检索逐渠道跑" },
            { key: "office_parse_findings", label: "office_parse_findings", note: "检索第三步：把结果文件读成摘要" },
        ];

        /** 子代理可用工具。键名必须与服务端白名单一致。 */
        const SUBAGENT_TOOLS = [
            { id: "office_web_search", label: "抓取搜索", note: "插件的免 Key 网页搜索，渠道分流的基础" },
            { id: "office_web_fetch", label: "抓取正文", note: "打开具体页面拿原文；PDF 由 office.pdf 抽文本" },
            { id: "read", label: "读文本", note: "读结果文件与已抓取内容" },
            { id: "read_image", label: "读图片", note: "图表 / 截图 / 扫描件" },
            { id: "write", label: "写文件", note: "子代理唯一的交付方式，建议保持开启" },
        ];

        /**
         * 站点清单的类型表与服务端 src/site-catalog.js 的 SITE_TYPES 同一份。
         *
         * 浏览器产物是单文件（factory 只拿到 require），服务端模块不在它的作用域里，
         * 所以只能镜像；漂移由 test/client.mjs 逐项核对，改一边就要改另一边。
         */
        const SITE_TYPES = [
            { id: "academic", name: "学术", note: "预印本、引用索引、文献库与文献检索" },
            { id: "book", name: "图书", note: "开放书库、书目与电子书" },
            { id: "code", name: "代码", note: "代码仓库、问答与包索引" },
            { id: "custom", name: "自定义", note: "自己加的站点（只按域名限定）" },
        ];

        /**
         * 内置目录：与服务端 BUILTIN_SITES 逐条一致（字段名与顺序都不能变，有测试核对）。
         *
         * 这一份同时是三样东西：`search.sites.entries` 的 schema 默认值（没配过的用户
         * 看到的就是它）、「恢复内置默认」写回去的内容，以及拿不到设置值时的兜底显示。
         * 顺序即默认优先级顺序；影子图书馆类只进目录、默认关闭（要用就在行上打开，
         * 或检索时按域名点名）。
         */
        const SITE_CATALOG = [
            // ── 学术 ──
            { type: "academic", domain: "arxiv.org", label: "arXiv", note: "预印本，理工科一手材料", enabled: true },
            { type: "academic", domain: "scholar.google.com", label: "Google 学术", note: "覆盖广，通常要代理才通", enabled: true },
            { type: "academic", domain: "cnki.net", label: "中国知网", note: "中文期刊与学位论文（多为摘要页）", enabled: true },
            { type: "academic", domain: "crossref.org", label: "CrossRef", note: "DOI 与引用元数据，核实出处最省事", enabled: true },
            { type: "academic", domain: "semanticscholar.org", label: "Semantic Scholar", note: "语义检索 + 引用关系", enabled: true },
            { type: "academic", domain: "pubmed.ncbi.nlm.nih.gov", label: "PubMed", note: "生物医学文献库", enabled: true },
            { type: "academic", domain: "researchgate.net", label: "ResearchGate", note: "作者自存的全文，质量参差", enabled: true },
            { type: "academic", domain: "sciencedirect.com", label: "ScienceDirect", note: "付费墙，通常只有摘要", enabled: false },
            { type: "academic", domain: "springer.com", label: "Springer", note: "付费墙，通常只有摘要", enabled: false },
            { type: "academic", domain: "ieee.org", label: "IEEE Xplore", note: "付费墙，通常只有摘要", enabled: false },
            { type: "academic", domain: "sci-hub.se", label: "Sci-Hub", note: "第三方镜像（影子图书馆），默认关闭；可用性与版权状态自行判断", enabled: false },

            // ── 图书 ──
            { type: "book", domain: "openlibrary.org", label: "Open Library", note: "书目与借阅入口", enabled: true },
            { type: "book", domain: "gutenberg.org", label: "Project Gutenberg", note: "公版电子书全文", enabled: true },
            { type: "book", domain: "book.douban.com", label: "豆瓣读书", note: "中文书目、目录与书评", enabled: true },
            { type: "book", domain: "standardebooks.org", label: "Standard Ebooks", note: "校订过的公版电子书", enabled: true },
            { type: "book", domain: "libgen.is", label: "Library Genesis", note: "影子图书馆，默认关闭；镜像域名常变", enabled: false },
            { type: "book", domain: "z-lib.io", label: "Z-Library", note: "影子图书馆，默认关闭；镜像域名常变", enabled: false },
            { type: "book", domain: "annas-archive.org", label: "Anna's Archive", note: "影子图书馆聚合，默认关闭", enabled: false },

            // ── 代码 ──
            { type: "code", domain: "github.com", label: "GitHub", note: "仓库、Issue 与讨论", enabled: true },
            { type: "code", domain: "stackoverflow.com", label: "Stack Overflow", note: "问答，常有可复现的答案", enabled: true },
            { type: "code", domain: "gitee.com", label: "Gitee", note: "国内仓库，直连更稳", enabled: true },
            { type: "code", domain: "developer.mozilla.org", label: "MDN", note: "Web 平台参考文档", enabled: true },
            { type: "code", domain: "pypi.org", label: "PyPI", note: "Python 包与文档入口", enabled: true },
            { type: "code", domain: "npmjs.com", label: "npm", note: "Node 包与文档入口", enabled: true },
        ];

        /** 单次最多限定几个站点，与服务端 SITE_PRIORITY_LIMITS.maxPerCall 同一条区间。 */
        const SITE_MAX_PER_CALL = [1, 8];

        /**
         * 域名的规范化：去掉协议、路径与 `www.` 前缀，转小写（与服务端
         * normalizeSiteEntry 的口径一致）。查重必须排在规范化之后 ——
         * `WWW.Arxiv.org/abs` 与 `arxiv.org` 是同一条。
         */
        function normalizeSiteDomain(raw) {
            return String(raw === undefined || raw === null ? "" : raw).trim().toLowerCase()
                .replace(/^https?:\/\//, "")
                .replace(/\/.*$/, "")
                .replace(/^www\./, "");
        }

        /**
         * 合法域名只认 `a.b` 形状（与服务端 site-catalog.js 的 DOMAIN_PATTERN 同一条）。
         * 不合法就不写入：宁可少一条，也不把坏域名塞进限定查询。
         */
        const SITE_DOMAIN_PATTERN = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/i;

        /** 主题下拉。 */
        // 顺序与服务端 engine/theme.js 的 THEMES 一致（有测试逐项核对）。
        const THEMES = [
            { id: "plain", label: "plain —— 素色网格（默认，不加背景填充）" },
            { id: "business", label: "business —— 商务蓝" },
            { id: "minimal", label: "minimal —— 简约灰" },
            { id: "warm", label: "warm —— 暖阳橙" },
            { id: "forest", label: "forest —— 森林绿" },
            { id: "tech", label: "tech —— 科技深色" },
            { id: "academic", label: "academic —— 学术宋" },
        ];

        /** 记忆范围下拉。值必须与服务端 MEMORY_SCOPES 一致（有测试逐项核对）。 */
        const MEMORY_SCOPE_OPTIONS = [
            { value: "workspace", label: "workspace —— 只跟当前工作目录（一个项目一份）" },
            { value: "global", label: "global —— 所有项目共用一份" },
            { value: "both", label: "both —— 两者并存（全局偏好在前、项目约定在后）" },
        ];

        /** 「用户偏好」落哪一层。值必须与服务端 USER_SCOPES 一致。 */
        const USER_SCOPE_OPTIONS = [
            { value: "memory", label: "memory —— 跟着上面的记忆范围" },
            { value: "global", label: "global —— 始终落全局层（换工作目录不必重记）" },
        ];

        /** 召回质量策略。值必须与服务端 schema 的 union 一致。 */
        const RECALL_POLICY_OPTIONS = [
            { value: "strict-v1", label: "strict-v1 —— 按相关度分档，低相关的丢掉" },
            { value: "off", label: "off —— 只按关键词排序，不丢结果" },
        ];

        /** 后台任务 Agent 的模型路由方式。 */
        const SUBAGENT_MODEL_MODE_OPTIONS = [
            { value: "inherit", label: "inherit —— 跟随主会话" },
            { value: "fixed", label: "fixed —— 用下面指定的模型" },
        ];

        /** 检索的执行引擎。 */
        const SEARCH_ENGINE_OPTIONS = [
            { value: "auto", label: "auto —— 先派子代理，跑不了自动改用内置检索" },
            { value: "subagent", label: "subagent —— 只用子代理" },
            { value: "builtin", label: "builtin —— 只用内置检索" },
        ];

        /** 音频 / 视频转写的语言提示（与服务端 AV_LANGUAGES 同一份取值）。 */
        const AV_LANGUAGE_OPTIONS = [
            { value: "auto", label: "auto —— 让模型自己判" },
            { value: "zh", label: "中文 zh" },
            { value: "yue", label: "粤语 yue" },
            { value: "en", label: "英文 en" },
            { value: "ja", label: "日文 ja" },
            { value: "ko", label: "韩文 ko" },
        ];

        /**
         * 内置检索（office 自己的联网通道）的可编辑参数。
         *
         * 写回路径写成**字面量**：设置面板的「界面覆盖服务端每个参数」那条测试按
         * 源码里的点号路径逐项核对，拼出来的路径在源码里根本不存在。
         */
        const BUILTIN_ROWS = [
            { key: "apiKeyEnv", path: "search.builtin.apiKeyEnv", label: "Key 环境变量名", note: "先问宿主凭据服务，再看环境变量，最后看 $DSH_HOME/.credentials.yaml", kind: "text" },
            { key: "apiKey", path: "search.builtin.apiKey", label: "Key 字面量", note: "一般留空，用上面的环境变量名；填了以它为准", kind: "password" },
            { key: "baseURL", path: "search.builtin.baseURL", label: "检索端点", note: "Anthropic 兼容 Messages 接口的 base（末尾会拼 /messages），默认 DeepSeek 官方", kind: "text" },
            { key: "model", path: "search.builtin.model", label: "检索模型", note: "要支持原生 web_search 工具", kind: "text" },
            { key: "maxResults", path: "search.builtin.maxResults", label: "每次来源上限", note: "每个渠道 / 每条直查查询最多几条来源（1-20）", kind: "number", min: 1, max: 20, step: 1 },
            { key: "fetchPages", path: "search.builtin.fetchPages", label: "取正文页数", note: "每个渠道 / 每条查询最多打开几页取摘录（0-5）；0 = 只用搜索返回的片段", kind: "number", min: 0, max: 5, step: 1 },
            { key: "maxUses", path: "search.builtin.maxUses", label: "web_search 次数", note: "一次检索最多让模型调用几次原生 web_search（1-20）", kind: "number", min: 1, max: 20, step: 1 },
            { key: "maxTokens", path: "search.builtin.maxTokens", label: "生成 token 上限", note: "一次检索最多生成多少 token（只要来源清单，不需要长文）", kind: "number", min: 256, max: 16384, step: 256 },
            { key: "searchTimeoutMs", path: "search.builtin.searchTimeoutMs", label: "检索超时", note: "单次检索的超时（毫秒）", kind: "number", min: 5000, max: 300000, step: 5000, unit: "ms" },
            { key: "fetchTimeoutMs", path: "search.builtin.fetchTimeoutMs", label: "取正文超时", note: "单页取正文的超时（毫秒）", kind: "number", min: 5000, max: 300000, step: 5000, unit: "ms" },
            { key: "maxBytes", path: "search.builtin.maxBytes", label: "单页体积上限", note: "单页正文的体积上限（字节），超出即截断", kind: "number", min: 32768, max: 33554432, step: 32768, unit: "B" },
            { key: "maxChars", path: "search.builtin.maxChars", label: "单页字符上限", note: "单页正文的字符上限，超出即截断", kind: "number", min: 1000, max: 400000, step: 1000, unit: "字" },
            { key: "maxRedirects", path: "search.builtin.maxRedirects", label: "跳转次数上限", note: "取正文最多跟随几次同源跳转（跨站跳转一律不跟）", kind: "number", min: 0, max: 10, step: 1, unit: "次" },
            { key: "pdfMaxBytes", path: "search.builtin.pdfMaxBytes", label: "PDF 体积上限", note: "取正文撞上 PDF 时的体积上限（字节）。超了直接放弃（截断的 PDF 抽不出文本）：换来源，或先下载到本地用 office.pdf 读", kind: "number", min: 1048576, max: 134217728, step: 1048576, unit: "B" },
            { key: "pdfMaxPages", path: "search.builtin.pdfMaxPages", label: "PDF 页数上限", note: "PDF 抽文本最多抽前几页，用来圈住抽取耗时（1-500）", kind: "number", min: 1, max: 500, step: 1, unit: "页" },
        ];

        /**
         * 记忆层拓扑的三行。
         *
         * 写回路径写成**字面量**（而不是 "memory.layers." + key 拼出来）：设置面板
         * 的「界面覆盖服务端每个参数」那条测试是按源码里的点号路径逐项核对的，
         * 拼出来的路径在源码里根本不存在，等于把一个开关漏在了界面之外。
         */
        const LAYER_ROWS = [
            { key: "hot", path: "memory.layers.hot", label: "热记忆层", note: "用户偏好与项目约定，常驻并随 office_help / office_run 投影出来" },
            { key: "ledger", path: "memory.layers.ledger", label: "台账层", note: "每份交付物一条自动登记。关掉后 office_run 不再登记，手工 log 也停" },
            { key: "archive", path: "memory.layers.archive", label: "归档层", note: "下沉的旧条目与滚动的旧台账。关掉后检索不到归档，但文件仍在盘上" },
        ];

        /** 记忆浏览面板的六个 Tab；key 对应快照里的数组字段名。 */
        const BROWSER_TABS = [
            { key: "hot", label: "热记忆" },
            { key: "ledger", label: "台账" },
            { key: "archive", label: "归档" },
            // 第四十二轮：知识库（kb 一期）的文档清单。面板只读：入库 / 读取都由
            // 会话里的 office_memory 执行，这里给的是 manifest 行（路径 / 块数 / 档位）。
            { key: "kb", label: "知识库" },
            { key: "links", label: "关系" },
            { key: "entities", label: "实体" },
        ];

        /**
         * 列表每页条数（「再显示 N 条」的步长）。
         *
         * 端点上各层的上限是 200 / 200 / 500 / 60（`src/view.js` 的 VIEW_LIMITS），
         * 所以渐进式加载只发生在**已经取回来的那一份**里，不需要给端点加 offset
         * 参数 —— 少一个参数就少一条要守的信任边界。20 这个数取的是「一屏多一点」：
         * 热记忆通常只有几条，台账与关系才需要翻页。
         */
        const LIST_PAGE = 20;

        /**
         * 记忆浏览面板的入口 id。
         *
         * 这个字符串被用在**两个** slot 上：sidebar.panellist 的 id 与 main 的 key。
         * 侧栏按钮就是按「id === key」去点亮主面板那一格的，写成两个名字会得到
         * 一个点不亮的面板，所以只能有一份定义。
         */
        const MEMORY_PANEL_ID = "office-memory";

        /**
         * 记忆浏览面板的数据端点（宿主侧注册的只读 HTTP 路由）。
         *
         * 为什么要有这条通道：浏览器半侧只有配置读写（configForms）与远程调用，读不到
         * 工作目录里的文件；记忆内容必须由宿主提供。用**同源相对路径**而不是
         * 绝对 URL：相对路径自然命中当前部署，也不会把请求发到别的源上去。
         */
        const MEMORY_VIEW_PATH = "/office-memory/snapshot";

        // ── 样式 ──────────────────────────────────────────────────────────
        //
        // 全部用内联样式 + 宿主主题变量，不引入 CSS 文件：
        // 客户端 bundle 是单文件产物，外部样式表没法跟着走；
        // 用 var(--…) 能自动跟随浅色/深色主题。
        //
        // ── 记忆面板的令牌层 ──
        //
        // 这一层是「像 mnemon 一样」的落点。先把 mnemon 的浏览器半侧 bundle 数了一遍：
        // 它出现频率最高的东西不是装饰，而是**宿主设计系统的原生令牌** ——
        // border-l2(45) / label-primary(38) / label-tertiary(36) / label-secondary(32) /
        // state-business-primary(25) / border-l1(18) / interactive-bg-hover(18)，
        // 表面用 bg-layer-1/2/3，输入框用 specific-input-major。
        // 反过来看本面板原来的写法：满屏 `opacity: 0.6` 的文字 + 一律拿
        // interactive-bg-hover 当底色 —— 层级全靠透明度硬撑，深色主题下尤其糊，
        // 也没有「表面 / 文字 / 描边」的分工。所以这里先把令牌收敛成一层 M，
        // 再让 S 里的面板样式只引用 M。
        //
        // 两条硬约束（有测试逐条钉住）：
        //   1. 只能用宿主真实存在的 --dsw-*（另一套 dsh 前缀的变量在宿主里一个都不存在，写了等于永远走兜底）；
        //   2. 每个令牌都带字面兜底 —— 令牌缺失时不能变成透明边框或没背景。
        // 兜底值一律取「深浅两色主题下都退得过去」的半透明中性色。
        const M = {
            /** 抬升表面（卡片）。 */
            surface: "var(--dsw-alias-bg-layer-1, rgba(255,255,255,0.72))",
            /**
             * 下沉底：轨道、嵌套小卡、chip、分段控件轨道。
             *
             * 这里刻意**不用** bg-layer-2：浅色主题里 layer-1/2/3 都是 #fff
             * （运行时探针实测），拿它当嵌套底等于白叠白，什么都看不出来。
             * bg-module-platform 是 #f5f6f7 / 深色 #353638，两个主题都分得开。
             */
            sunken: "var(--dsw-alias-bg-module-platform, rgba(128,128,128,0.1))",
            /** 输入框底（宿主给输入框用的就是它）。 */
            input: "var(--dsw-specific-input-major, rgba(255,255,255,0.6))",
            hover: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))",
            /** 描边：细（卡片轮廓，配合 elevation 的 0.5px 描边）/ 强（chip、输入框）。 */
            stroke: "var(--dsw-alias-border-l1, rgba(128,128,128,0.22))",
            strokeStrong: "var(--dsw-alias-border-l2, rgba(128,128,128,0.3))",
            /** 文字四级：主 / 次 / 三级 / 说明。不再用 opacity 造层级。 */
            text: "var(--dsw-alias-label-primary, inherit)",
            textSecondary: "var(--dsw-alias-label-secondary, rgba(128,128,128,0.92))",
            textTertiary: "var(--dsw-alias-label-tertiary, rgba(128,128,128,0.78))",
            caption: "var(--dsw-alias-label-caption, rgba(128,128,128,0.62))",
            /** 强调与状态色。 */
            accent: "var(--dsw-alias-brand-primary, #4d6bfe)",
            business: "var(--dsw-alias-state-business-primary, #5a8cff)",
            success: "var(--dsw-alias-state-success-primary, #22a879)",
            warn: "var(--dsw-alias-state-warn-primary, #c08a2e)",
            error: "var(--dsw-alias-state-error-primary, #e5484d)",
            /** 高度：soft 给卡片，lift 给悬停 / 选中。 */
            soft: "var(--dsw-elevation-soft, 0 1px 2px rgba(0,0,0,0.05))",
            lift: "var(--dsw-elevation-panel, 0 2px 6px rgba(0,0,0,0.06))",
            radiusSm: "8px",
            radiusMd: "12px",
            radiusLg: "16px",
            pill: "999px",
            font: "var(--dsw-font-family, inherit)",
            mono: "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace",
        };
        /** 数字一律等宽：容量条与计数并排时才不会左右跳。 */
        const NUM = { fontVariantNumeric: "tabular-nums" };
        // ── 宿主类型刻度 ──
        //
        // 第二十轮把设置页的字号从「写死的 11/12/12.5/13.5px」换成宿主的刻度
        // （`--dsw-font-*`，宿主 UI 自己用的就是这套）：
        //   xxxs-11 = 11px/14px   xxs-12 = 12px/18px   xs-13 = 13px/20px
        //   s-14    = 14px/22px   base-16 = 16px/24px
        // 为什么不是「直接把数字调大」：宿主将来改刻度（或用户在系统层改字号）时，
        // 面板会跟着走；写死的像素不会。三个分量各自取令牌、各自带字面兜底 ——
        // `font:` 简写整条无效时字号会一起丢，分着写就不会。
        const TYPE = (token, size, lineHeight, weight) => ({
            fontSize: "var(--dsw-font-" + token + "-font-size, " + size + ")",
            lineHeight: "var(--dsw-font-" + token + "-line-height, " + lineHeight + ")",
            fontWeight: "var(--dsw-font-" + token + "-font-weight, " + weight + ")",
        });
        const T = {
            /** 分组标题（14px 中黑）。 */
            heading: TYPE("s-strong-14", "14px", "22px", 500),
            /** 一级正文（14px）。 */
            body: TYPE("s-14", "14px", "22px", 400),
            /** 控件与强调标签（13px 中黑）。 */
            strong: TYPE("xs-strong-13", "13px", "20px", 500),
            /** 控件与正文（13px）。 */
            label: TYPE("xs-13", "13px", "20px", 400),
            /** 说明文字（12px）。 */
            note: TYPE("xxs-12", "12px", "18px", 400),
            /** 说明里要强调的那几个词（12px 中黑）。 */
            noteStrong: TYPE("xxs-strong-12", "12px", "18px", 500),
            /** 最小一级：chip、徽标、代码块（11px 在宿主里是最小档）。 */
            tiny: TYPE("xxxs-11", "11px", "14px", 400),
            /** 页面主标题（16px）。 */
            title: TYPE("base-16", "16px", "24px", 400),
        };
        /** 次级文字一律用颜色令牌，不用 opacity 造层级（第十二轮定的规矩）。 */
        const dim = (type, color) => Object.assign({}, type, { color: color ?? M.textTertiary });
        const S = {
            section: { padding: "2px 0" },
            group: { marginBottom: "26px" },
            groupTitle: Object.assign({ margin: "0 0 4px" }, T.heading),
            groupNote: Object.assign({ margin: "0 0 12px" }, dim(T.note)),
            row: {
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: "16px",
                padding: "10px 0",
                borderTop: "1px solid var(--dsw-alias-border-l2, rgba(128,128,128,0.18))"
            },
            rowFirst: { borderTop: "none" },
            // 控件需要整行宽度时用这一套：标题在上、控件在下，避免横向挤压。
            rowStacked: { display: "flex", flexDirection: "column", alignItems: "stretch", gap: "8px" },
            controlStacked: { flex: "0 0 auto", display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap", justifyContent: "flex-start" },
            rowTextStacked: { minWidth: 0 },
            // minWidth 必须留一个下限：控件那一列是 flex: 0 0 auto（不能被压缩），
            // 若这里允许缩到 0，宽控件（例如七个可勾选标签）会把标题挤成一字一行。
            rowText: { flex: "1 1 240px", minWidth: "180px" },
            rowLabel: Object.assign({ marginBottom: "2px" }, T.body),
            rowNote: Object.assign({}, dim(T.note)),
            control: { flex: "0 1 auto", display: "flex", alignItems: "center", gap: "8px", justifyContent: "flex-end", flexWrap: "wrap" },
            input: Object.assign({
                width: "92px",
                padding: "6px 10px",
                borderRadius: "8px",
                border: "1px solid var(--dsw-alias-border-l1, rgba(128,128,128,0.3))",
                background: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))",
                color: "inherit",
                textAlign: "right"
            }, T.label),
            select: Object.assign({
                padding: "6px 10px",
                borderRadius: "8px",
                border: "1px solid var(--dsw-alias-border-l1, rgba(128,128,128,0.3))",
                background: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))",
                color: "inherit",
                minWidth: "220px"
            }, T.label),
            unit: Object.assign({ minWidth: "20px" }, dim(T.note)),
            // 命令片段：等宽、可选、可换行（迁移那一组要给人复制走）。
            code: Object.assign({
                display: "block",
                fontFamily: M.mono,
                padding: "6px 10px",
                borderRadius: "8px",
                background: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.12))",
                border: "1px solid var(--dsw-alias-border-l2, rgba(128,128,128,0.2))",
                userSelect: "text",
                wordBreak: "break-all"
            }, T.label),
            chips: { display: "flex", flexWrap: "wrap", gap: "8px", marginTop: "4px" },
            chip: Object.assign({
                display: "flex",
                alignItems: "center",
                gap: "6px",
                padding: "6px 10px",
                borderRadius: "999px",
                border: "1px solid var(--dsw-alias-border-l1, rgba(128,128,128,0.3))",
                cursor: "pointer",
                userSelect: "none"
            }, T.label),
            chipOn: {
                background: "var(--dsw-alias-state-business-primary, rgba(90,140,255,0.22))",
                borderColor: "var(--dsw-alias-state-business-primary, rgba(90,140,255,0.5))"
            },
            chipOff: { color: M.textTertiary, borderStyle: "dashed" },
            status: Object.assign({ marginBottom: "14px" }, dim(T.note)),
            error: Object.assign({ marginBottom: "12px" }, T.note, { color: "var(--dsw-alias-state-error-primary, #e5484d)" }),
            button: Object.assign({
                padding: "6px 12px",
                borderRadius: "8px",
                border: "1px solid var(--dsw-alias-border-l1, rgba(128,128,128,0.3))",
                background: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))",
                color: "inherit",
                cursor: "pointer"
            }, T.label),
            // 分组的折叠头：整行可点，右边一个 caret。默认收起的长参数组用它，
            // 「一眼能看完的页」比「一屏全是输入框」好操作（第二十轮）。
            foldHead: {
                display: "flex",
                alignItems: "center",
                gap: "8px",
                width: "100%",
                padding: "6px 0",
                background: "none",
                border: "none",
                color: "inherit",
                cursor: "pointer",
                textAlign: "left"
            },
            foldCaret: Object.assign({ color: M.textTertiary, flex: "0 0 auto" }, T.note),
            foldCount: Object.assign({ color: M.caption }, T.tiny),
            // ── 分组折叠与定位（第三十五轮）──
            // 分组头：整行可点，左边 caret + 组名，右边「N 项 / 展开」。
            // 底下一条细线把它和组体分开，收起时一排组名本身就是目录。
            groupHead: {
                display: "flex",
                alignItems: "center",
                gap: "8px",
                width: "100%",
                padding: "6px 0",
                background: "none",
                border: "none",
                borderBottom: "1px solid var(--dsw-alias-border-l2, rgba(128,128,128,0.18))",
                color: "inherit",
                cursor: "pointer",
                textAlign: "left"
            },
            groupHeadTitle: T.heading,
            groupHeadCount: Object.assign({ color: M.caption }, T.tiny),
            groupHeadHint: Object.assign({ marginLeft: "auto" }, dim(T.tiny)),
            // 定位条吸顶：分组展开后这一页很长，输入框滑出视野就没法连着换词定位。
            // 底色取宿主的 base 表面色（与记忆面板的吸顶头同一套写法）。
            locator: {
                position: "sticky",
                top: 0,
                zIndex: 2,
                padding: "8px 0 10px",
                marginBottom: "8px",
                background: "var(--dsw-alias-bg-base, #fff)",
                borderBottom: "1px solid var(--dsw-alias-border-l2, rgba(128,128,128,0.18))"
            },
            locatorRow: { display: "flex", flexWrap: "wrap", gap: "8px", alignItems: "center" },
            locatorInput: Object.assign({
                flex: "1 1 260px",
                minWidth: "200px",
                padding: "6px 10px",
                borderRadius: "8px",
                border: "1px solid var(--dsw-alias-border-l1, rgba(128,128,128,0.3))",
                background: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))",
                color: "inherit"
            }, T.label),
            locatorStats: Object.assign({ marginTop: "6px" }, dim(T.note)),
            quickBar: { display: "flex", flexWrap: "wrap", gap: "8px", alignItems: "center", marginBottom: "10px" },
            callout: {
                padding: "10px 12px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.stroke,
                background: M.sunken,
                marginBottom: "12px"
            },
            // 回合记忆条：只在本回合真的调过记忆 / 检索工具时才画，所以样式做轻，
            // 一行、小字、低对比 —— 它是一条旁注，不该和回合尾的操作按钮抢注意力。
            turnBar: {
                fontSize: "12px",
                opacity: 0.65,
                padding: "6px 10px",
                borderRadius: "8px",
                border: "1px solid var(--dsw-alias-border-l2, rgba(128,128,128,0.2))",
                background: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.08))",
                marginBottom: "8px",
                lineHeight: 1.5
            },
            // ── 记忆浏览面板 ──
            //
            // 版式：头部（图标 + 标题 + 目录 + 工具）→ 分段控件（六个页签）→ 指标卡 →
            // 两栏（容量 / 存储域）→ 随页签出现的可视化 → 条目列表 → 页脚元信息。
            //
            // 宽度收在 1120px：这一格在宽屏下会拉满，一行路径能拖到屏幕另一头，
            // 读起来很累；内容（路径、正文）仍然自己滚动 + 强制断行。
            browser: {
                padding: "4px 0 28px",
                maxWidth: "1120px",
                color: M.text,
                fontFamily: M.font,
                // 容器查询的锚点：.om-grid-wide 的两栏在面板自己变窄时塌成一列，
                // 而不是等整个窗口变窄（见 PANEL_CSS）。
                containerType: "inline-size",
                // ── 面板自己就是滚动容器（第二十轮修的滚动问题）──
                //
                // 为什么必须自己滚：主面板那一格的宿主容器是
                // `.pI_x6G_centerCol{display:flex;flex-direction:column;overflow:hidden}`
                // —— 它**不给** main 槽位任何滚动能力，而槽位的 div 是
                // `display:contents`（不产生盒子），所以本面板的根元素就是那个
                // flex 项。上一版没有高度、没有 overflow：内容比视口高时被中心列
                // 直接裁掉，只有列表内部那 520px 能滚 —— 「指标卡 / 容量条 / 条目
                // 详情 / 页脚」全都在首屏之外够不着（第十三轮已把它记成已知边界）。
                //
                // 修法就是标准姿势：撑满可用高度 + 自己滚动。`height:100%` 对
                // flex 父容器里高度已定的项是可靠的；`minHeight:0` 与
                // `flex:"1 1 auto"` 一起写是为了不依赖 slot 包装层的实现细节。
                height: "100%",
                maxHeight: "100%",
                minHeight: 0,
                flex: "1 1 auto",
                boxSizing: "border-box",
                overflowY: "auto",
                overflowX: "hidden",
                // 滚到底不要把滚动链传给宿主（宿主那一格不滚，链过去只会让整页跳）。
                overscrollBehavior: "contain"
            },
            // 头部随面板一起滚会带来一个副作用：搜索框与「刷新 / 记一条」滑出视野。
            // 所以头部**吸顶**（sticky 在面板这个滚动容器里），底色取宿主的 base
            // 表面色 —— 面板本来就画在这个底色上，吸顶块不会看出边界。
            browserHead: {
                display: "flex", alignItems: "flex-start", gap: "12px", flexWrap: "wrap",
                position: "sticky", top: 0, zIndex: 2,
                padding: "4px 0 10px", marginBottom: "10px",
                background: "var(--dsw-alias-bg-base, #fff)"
            },
            browserHeadIcon: {
                width: "30px",
                height: "30px",
                flex: "0 0 auto",
                borderRadius: "10px",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                background: M.sunken,
                border: "1px solid " + M.stroke,
                color: M.accent
            },
            browserTitleBlock: { flex: "1 1 260px", minWidth: "200px" },
            browserTitle: { fontSize: "16px", fontWeight: 600, lineHeight: 1.3, letterSpacing: "0.01em" },
            browserSubtitle: { fontSize: "12px", color: M.textTertiary, marginTop: "3px", lineHeight: 1.5, overflowWrap: "anywhere" },
            browserTools: { display: "flex", alignItems: "center", gap: "8px", marginLeft: "auto", flexWrap: "wrap" },
            // 搜索框：图标绝对定位在左边，输入框自己留出 30px 内边距 ——
            // 这样 type / placeholder / onKeyDown 都还是原来那个 input（有测试按它们找）。
            searchWrap: { position: "relative", display: "inline-flex", alignItems: "center" },
            searchIcon: { position: "absolute", left: "9px", display: "flex", color: M.caption, pointerEvents: "none" },
            searchInput: {
                width: "230px",
                padding: "7px 10px 7px 30px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.strokeStrong,
                background: M.input,
                color: M.text,
                font: "inherit",
                fontSize: "12.5px",
                outline: "none"
            },
            ghostButton: {
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                padding: "7px 12px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.strokeStrong,
                background: M.surface,
                color: M.textSecondary,
                font: "inherit",
                fontSize: "12.5px",
                cursor: "pointer",
                boxShadow: M.soft
            },
            // 分段控件：一条浅底轨道 + 选中项浮起。比原来那种「整颗实心蓝药丸」
            // 安静得多，六个页签并排时也不会把注意力全抢走。
            tabsTrack: {
                display: "flex",
                flexWrap: "wrap",
                gap: "2px",
                marginBottom: "14px",
                padding: "3px",
                borderRadius: M.radiusMd,
                background: M.sunken,
                border: "1px solid " + M.stroke,
                width: "fit-content",
                maxWidth: "100%"
            },
            browserTab: {
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                padding: "6px 12px",
                borderRadius: "9px",
                border: "1px solid transparent",
                cursor: "pointer",
                fontSize: "12.5px",
                color: M.textTertiary,
                userSelect: "none"
            },
            browserTabOn: {
                background: M.surface,
                borderColor: M.stroke,
                color: M.text,
                fontWeight: 600,
                boxShadow: M.soft
            },
            browserTabHover: { color: M.textSecondary, background: M.hover },
            browserTabDot: { width: "7px", height: "7px", borderRadius: M.pill, flex: "0 0 auto" },
            // 列表**不再自己滚动**（第二十轮）：面板根元素已经是唯一的滚动容器，
            // 再套一层 520px 的内层滚动会出现「滚到中间卡住、要换一个滚动条接着滚」
            // 的双滚动条，而且条目详情与页脚还被压在下面看不见。
            // 分页（每页 20 条 + 再显示 N 条）仍然管「DOM 里有几条」。
            browserList: { display: "flex", flexDirection: "column", gap: "8px" },
            browserItem: {
                padding: "10px 12px",
                borderRadius: M.radiusMd,
                border: "1px solid " + M.stroke,
                background: M.surface,
                boxShadow: M.soft,
                fontSize: "12.5px",
                lineHeight: 1.6,
                overflowWrap: "anywhere",
                wordBreak: "break-word"
            },
            browserItemHover: { borderColor: M.strokeStrong, boxShadow: M.lift },
            browserMeta: { display: "flex", flexWrap: "wrap", gap: "6px", marginBottom: "6px", fontSize: "11px" },
            badge: {
                padding: "1px 8px",
                borderRadius: M.pill,
                border: "1px solid " + M.strokeStrong,
                background: M.sunken,
                color: M.textTertiary,
                whiteSpace: "nowrap"
            },
            // 正文限高 + pre-wrap：换行照原样保留，但不会把一条记忆拉成一屏。
            // 限高只是**折叠**不是截断：卡片上的「展开」把 maxHeight 去掉，完整内容
            // 就地铺开。原来这里挂的是原生 title，系统气泡既不可样式化、也没法键盘
            // 触发，而且它出现的位置由浏览器决定、常常盖住相邻条目 —— 现在完整内容
            // 只有两个出口：点「展开」（键盘可达），或看面板底部那条条目详情。
            browserBody: { whiteSpace: "pre-wrap", maxHeight: "9em", overflow: "hidden", color: M.textSecondary },
            browserBodyOpen: { whiteSpace: "pre-wrap", color: M.textSecondary },
            /** 条目底部那排小动作：展开 / 改 / 忘 / 删关系 …。 */
            itemActions: { display: "flex", flexWrap: "wrap", gap: "6px", marginTop: "8px" },
            itemAction: {
                display: "inline-flex",
                alignItems: "center",
                gap: "4px",
                padding: "3px 9px",
                borderRadius: M.pill,
                border: "1px solid " + M.strokeStrong,
                background: M.sunken,
                color: M.textTertiary,
                font: "inherit",
                fontSize: "11px",
                cursor: "pointer",
                whiteSpace: "nowrap"
            },
            /** 「再显示 N 条」那一条：按钮在左，剩余条数在右。 */
            listMore: { display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap", marginTop: "10px", paddingTop: "10px", borderTop: "1px solid " + M.stroke },
            listMoreNote: { fontSize: "11px", color: M.caption, fontFamily: M.mono, ...NUM },
            /** 受监督的写入：面板只**生成指令**，不落盘。 */
            pending: {
                marginBottom: "14px",
                padding: "12px 14px",
                borderRadius: M.radiusMd,
                border: "1px solid color-mix(in srgb, " + M.business + " 42%, " + M.stroke + ")",
                background: tintedSurface(M.business),
                boxShadow: M.soft
            },
            pendingHead: { display: "flex", alignItems: "baseline", gap: "8px", flexWrap: "wrap", marginBottom: "8px" },
            pendingTitle: { fontSize: "12.5px", fontWeight: 600, color: M.text },
            pendingNote: { fontSize: "11px", color: M.textTertiary, marginLeft: "auto" },
            pendingCode: {
                margin: "0 0 10px",
                padding: "9px 11px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.strokeStrong,
                background: M.input,
                color: M.text,
                fontFamily: M.mono,
                fontSize: "11.5px",
                lineHeight: 1.6,
                whiteSpace: "pre-wrap",
                overflowWrap: "anywhere",
                maxHeight: "16em",
                overflow: "auto"
            },
            pendingActions: { display: "flex", flexWrap: "wrap", gap: "8px", alignItems: "center" },
            pendingHint: { fontSize: "11px", color: M.caption, lineHeight: 1.6, marginTop: "8px" },
            /** 条目详情条：与图谱右侧 INSPECTOR 同一套路（列表里不做浮动气泡）。 */
            itemDetail: {
                marginTop: "12px",
                padding: "11px 13px",
                borderRadius: M.radiusMd,
                border: "1px solid " + M.stroke,
                background: M.sunken
            },
            itemDetailHead: { display: "flex", alignItems: "baseline", gap: "8px", flexWrap: "wrap", marginBottom: "6px" },
            itemDetailKey: { fontSize: "9px", letterSpacing: "0.12em", color: M.caption, fontFamily: M.mono, textTransform: "uppercase" },
            itemDetailTitle: { fontSize: "12px", fontWeight: 600, color: M.textSecondary, marginLeft: "auto", fontFamily: M.mono },
            itemDetailBody: { whiteSpace: "pre-wrap", color: M.textSecondary, fontSize: "12px", lineHeight: 1.65, overflowWrap: "anywhere", maxHeight: "14em", overflow: "auto" },
            itemDetailHint: { fontSize: "11px", color: M.caption, marginTop: "6px", lineHeight: 1.6 },

            // ── 可视化：指标条 / 容量条 / 存储卡 / 时间线 / 关系图 / 实体条 ──
            //
            // 这里全部用内联样式而不是运行时注入 CSS：bundle 在测试与 SSR 环境里
            // 是没有 document 的（test/client.mjs 用 new Function 把它跑起来，
            // 只给 window / require / console / fetch）。一旦在模块顶层碰 document，
            // 整个 bundle 会当场抛错、所有用例一起红。
            //
            // 需要伪类与关键帧的地方（悬停、焦点、转圈、窄屏塌列）走面板里的
            // <style> 节点（见 PANEL_CSS）—— 那是 React 树里的普通节点，不是
            // 模块顶层副作用，两个环境里都安全。
            tiles: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: "10px", marginBottom: "14px" },
            tile: {
                display: "flex",
                flexDirection: "column",
                gap: "6px",
                padding: "12px 14px",
                borderRadius: M.radiusMd,
                border: "1px solid " + M.stroke,
                background: M.surface,
                boxShadow: M.soft
            },
            tileHead: { display: "flex", alignItems: "center", gap: "7px" },
            tileDot: { width: "8px", height: "8px", borderRadius: "3px", flex: "0 0 auto" },
            tileLabel: { fontSize: "12px", color: M.textTertiary, fontWeight: 500 },
            // 指标数字用等宽字：并排五枚时数字对得齐，也带出 mnemon 那点技术感。
            tileValue: { fontSize: "23px", fontWeight: 650, lineHeight: 1.1, letterSpacing: "-0.01em", fontFamily: M.mono, color: M.text },
            tileNote: { fontSize: "11px", color: M.caption, lineHeight: 1.4 },

            panel: {
                marginBottom: "12px",
                padding: "14px 16px",
                borderRadius: M.radiusLg,
                border: "1px solid " + M.stroke,
                background: M.surface,
                boxShadow: M.soft
            },
            panelTitle: { display: "flex", alignItems: "center", gap: "8px", fontSize: "13px", fontWeight: 600, color: M.text, marginBottom: "10px" },
            panelTitleText: { flex: "0 0 auto" },
            panelNote: { fontSize: "11px", fontWeight: 400, color: M.caption, marginLeft: "auto", fontFamily: M.mono },
            panelFoot: { fontSize: "11px", color: M.caption, marginTop: "10px", lineHeight: 1.5 },
            /** 两栏仪表区：宽屏并排、窄屏塌成一列（塌列靠 PANEL_CSS 的媒体查询）。
             *  alignItems 用 start：两块卡片各自贴住自己的内容高度，右边那块
             *  不会因为左边高而被拉成一大片空白。 */
            panelGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: "12px", marginBottom: "0", alignItems: "start" },

            gauge: { marginBottom: "12px" },
            gaugeHead: { display: "flex", alignItems: "baseline", gap: "8px", fontSize: "12px", marginBottom: "6px" },
            gaugeName: { color: M.textSecondary },
            gaugeValue: { marginLeft: "auto", color: M.textTertiary, fontFamily: M.mono, ...NUM },
            gaugePercent: { color: M.caption, minWidth: "46px", textAlign: "right", fontFamily: M.mono, ...NUM },
            // 4px 轨道 + 继承圆角的填充（mnemon 的容量条口径）：细到像一条刻度，
            // 但右端不会露出直角。
            gaugeTrack: {
                position: "relative",
                height: "4px",
                borderRadius: M.pill,
                background: M.sunken,
                overflow: "hidden"
            },
            gaugeFill: { height: "100%", borderRadius: "inherit", background: M.business },
            storeCards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: "10px" },
            storeCard: {
                padding: "12px",
                borderRadius: M.radiusMd,
                border: "1px solid " + M.stroke,
                background: M.sunken,
                fontSize: "12px",
                lineHeight: 1.55
            },
            storeHead: { display: "flex", alignItems: "center", gap: "8px", marginBottom: "9px" },
            storeDir: {
                fontFamily: M.mono,
                fontSize: "11.5px",
                color: M.textSecondary,
                overflowWrap: "anywhere"
            },
            storeRow: { display: "flex", flexWrap: "wrap", gap: "6px", alignItems: "center" },

            timeline: { display: "flex", alignItems: "flex-end", gap: "6px", minHeight: "84px", paddingBottom: "6px", borderBottom: "1px solid " + M.stroke },
            timelineCol: { display: "flex", flexDirection: "column", alignItems: "center", gap: "4px", minWidth: "22px", flex: "1 1 auto" },
            timelineBar: { width: "100%", maxWidth: "24px", minHeight: "3px", borderRadius: "6px 6px 2px 2px", background: M.success },
            timelineCount: { fontSize: "10px", color: M.caption, fontFamily: M.mono, ...NUM },
            timelineLabel: { fontSize: "10px", color: M.caption, writingMode: "vertical-rl", maxHeight: "56px", overflow: "hidden" },

            graphCanvas: {
                display: "block",
                width: "100%",
                height: "auto",
                borderRadius: M.radiusMd
            },
            // 画布底：中心一团 6% 强调色的径向光晕 —— mnemon 的「高级感」有一半
            // 来自这一行；网格再叠在上面，细到只当纹理。
            graphCanvasWrap: {
                minWidth: 0,
                borderRadius: M.radiusMd,
                border: "1px solid " + M.stroke,
                background: "radial-gradient(circle at 50% 48%, color-mix(in srgb, " + M.business + " 7%, transparent), transparent 47%), " + M.sunken,
                overflow: "hidden"
            },
            graphArea: { display: "grid", gridTemplateColumns: "minmax(0,1fr) minmax(240px,270px)", gap: "12px", alignItems: "stretch" },
            graphToolbar: { display: "flex", flexWrap: "wrap", gap: "12px", alignItems: "center", minHeight: "38px", padding: "0 13px", borderTop: "1px solid " + M.stroke, fontSize: "10px", color: M.textTertiary, fontFamily: M.mono },
            /**
             * 画布上方的视图条：缩放三件套 + 两个布局动作。
             *
             * mnemon 的画布是 `cursor: grab` 直接拖、滚轮缩放、并给「自然铺开 /
             * 均匀重置」两个布局动作；这里对齐这三样，但**视图变换与布局是两件事**：
             * 缩放/平移只改一个父级 `<g>` 的 transform，节点的 translate 仍然是
             * 布局坐标 —— 于是拖动视图不会让布局重算，方向键微调的 12px 也仍然
             * 是布局单位（缩放后视觉上会跟着放大，这是对的）。
             */
            graphControls: { display: "flex", flexWrap: "wrap", gap: "6px", alignItems: "center", minHeight: "40px", padding: "0 12px", borderBottom: "1px solid " + M.stroke, background: M.surface },
            controlLabel: { fontSize: "10px", letterSpacing: "0.1em", color: M.caption, fontFamily: M.mono, textTransform: "uppercase", marginRight: "2px" },
            viewButton: {
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                minWidth: "26px",
                padding: "4px 9px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.strokeStrong,
                background: M.sunken,
                color: M.textSecondary,
                font: "inherit",
                fontSize: "11.5px",
                cursor: "pointer",
                whiteSpace: "nowrap"
            },
            viewButtonOn: { borderColor: M.business, background: tintedSurface(M.business), color: M.text, fontWeight: 600 },
            zoomLabel: { minWidth: "44px", textAlign: "center", fontSize: "11px", color: M.textTertiary, fontFamily: M.mono, ...NUM },
            controlSpacer: { flex: "1 1 auto" },
            legendItem: { display: "inline-flex", alignItems: "center", gap: "6px" },
            graphInspector: {
                minWidth: 0,
                padding: "14px",
                borderRadius: M.radiusMd,
                border: "1px solid " + M.stroke,
                background: M.sunken,
                display: "flex",
                flexDirection: "column",
                gap: "8px"
            },
            inspectorKicker: { fontSize: "9px", letterSpacing: "0.12em", color: M.caption, fontFamily: M.mono, textTransform: "uppercase" },
            inspectorTitle: { fontSize: "13.5px", fontWeight: 600, color: M.text, lineHeight: 1.6, overflowWrap: "anywhere" },
            inspectorMeta: { display: "flex", flexDirection: "column", gap: "6px", marginTop: "2px" },
            inspectorMetaRow: { display: "flex", alignItems: "baseline", gap: "8px", paddingTop: "6px", borderTop: "1px solid " + M.stroke, fontSize: "11.5px", color: M.textSecondary },
            inspectorMetaKey: { minWidth: "42px", color: M.caption, fontSize: "10px", fontFamily: M.mono, letterSpacing: "0.08em" },
            inspectorHint: { fontSize: "11.5px", color: M.textTertiary, lineHeight: 1.6, marginTop: "auto" },
            inspectorEmpty: { display: "flex", flexDirection: "column", gap: "8px", alignItems: "flex-start" },
            // 空态不画插画：一个圆圈 + 径向光晕 + 等宽字符（mnemon 的空态 glyph 就是这么做的）。
            inspectorEmptyGlyph: {
                width: "44px",
                height: "44px",
                borderRadius: M.pill,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                color: M.business,
                fontSize: "20px",
                fontFamily: M.mono,
                border: "1px solid color-mix(in srgb, " + M.business + " 35%, " + M.stroke + ")",
                background: "radial-gradient(circle, color-mix(in srgb, " + M.business + " 12%, transparent), transparent 65%)"
            },
            emptyInline: { padding: "18px 14px", borderRadius: M.radiusMd, border: "1px dashed " + M.strokeStrong, color: M.textTertiary, fontSize: "12px", textAlign: "center" },
            emptyState: {
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                gap: "10px",
                padding: "28px 24px",
                borderRadius: M.radiusLg,
                border: "1px dashed " + M.strokeStrong,
                background: "color-mix(in srgb, " + M.surface + " 60%, transparent)"
            },
            emptyGlyph: {
                width: "56px",
                height: "56px",
                borderRadius: M.pill,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                color: M.business,
                fontSize: "24px",
                fontFamily: M.mono,
                border: "1px solid color-mix(in srgb, " + M.business + " 35%, " + M.stroke + ")",
                background: "radial-gradient(circle, color-mix(in srgb, " + M.business + " 12%, transparent), transparent 65%)"
            },
            emptyTitle: { fontSize: "13px", fontWeight: 600, color: M.textSecondary },
            emptyHint: { fontSize: "11.5px", color: M.caption, lineHeight: 1.6, textAlign: "center", maxWidth: "420px" },
            loadingState: { display: "flex", alignItems: "center", gap: "10px", padding: "20px 16px", borderRadius: M.radiusLg, border: "1px solid " + M.stroke, color: M.textTertiary, fontSize: "12px" },
            spinner: {
                width: "13px",
                height: "13px",
                borderRadius: "50%",
                border: "1.5px solid color-mix(in srgb, " + M.business + " 24%, transparent)",
                borderTopColor: M.business,
                flex: "0 0 auto"
            },

            entityRow: { display: "grid", gridTemplateColumns: "minmax(72px, 26%) minmax(0,1fr) auto", alignItems: "center", gap: "10px", fontSize: "12px", marginBottom: "4px" },
            entityName: { color: M.textSecondary, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
            entityTrack: {
                height: "6px",
                borderRadius: M.pill,
                background: M.sunken,
                overflow: "hidden"
            },
            entityFill: { height: "100%", borderRadius: "inherit", background: M.success },
            entityCount: { color: M.textTertiary, fontFamily: M.mono, fontSize: "11.5px", ...NUM },
            entityBlock: { marginBottom: "12px" },
            layerChips: { display: "flex", flexWrap: "wrap", gap: "5px", margin: "6px 0 0 0" },
            layerChip: {
                display: "inline-flex",
                alignItems: "center",
                gap: "5px",
                padding: "2px 8px",
                borderRadius: M.pill,
                fontSize: "11px",
                color: M.textTertiary,
                border: "1px solid " + M.strokeStrong,
                background: M.sunken
            },
            layerChipDot: { width: "6px", height: "6px", borderRadius: M.pill, flex: "0 0 auto" },
            hotGroupTitle: { display: "flex", alignItems: "center", gap: "8px", fontSize: "12px", fontWeight: 600, color: M.textSecondary, margin: "6px 0 8px" },

            footerMeta: { display: "flex", flexWrap: "wrap", gap: "6px", marginTop: "16px", paddingTop: "12px", borderTop: "1px solid " + M.stroke },
            footerChip: {
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                padding: "3px 9px",
                borderRadius: M.pill,
                border: "1px solid " + M.strokeStrong,
                background: M.sunken,
                fontSize: "11px",
                color: M.textTertiary
            },
            footerKey: { color: M.caption, fontFamily: M.mono, fontSize: "10px", letterSpacing: "0.06em" },

            // ── 站点清单编辑器（检索编排组内）──
            //
            // 一行 = 启用开关 + 域名 + 显示名 + 说明 + 三个动作按钮。窄屏下整行换行
            // （flexWrap），不把域名输入框压成看不清的宽度；域名与显示名是短输入框，
            // 说明给宽一点。组头（类型名 + 一句说明）比行本身低一级。
            siteList: { display: "flex", flexDirection: "column", gap: "2px", marginTop: "4px" },
            siteGroupHead: Object.assign({ display: "flex", alignItems: "baseline", gap: "8px", marginTop: "10px" }, T.strong),
            siteGroupNote: Object.assign({ marginBottom: "2px" }, dim(T.note)),
            siteRow: {
                display: "flex",
                alignItems: "center",
                gap: "8px",
                flexWrap: "wrap",
                padding: "6px 0",
                borderTop: "1px solid var(--dsw-alias-border-l2, rgba(128,128,128,0.14))"
            },
            siteCell: { display: "flex", alignItems: "center", gap: "6px", flex: "0 0 auto" },
            siteField: Object.assign({
                padding: "5px 8px",
                borderRadius: "8px",
                border: "1px solid var(--dsw-alias-border-l1, rgba(128,128,128,0.3))",
                background: "var(--dsw-alias-interactive-bg-hover, rgba(128,128,128,0.1))",
                color: "inherit",
                textAlign: "left"
            }, T.label),
            siteAdd: { display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap", marginTop: "10px" },
            siteHint: Object.assign({ marginTop: "8px" }, dim(T.note)),
        };

        // ── 小组件 ────────────────────────────────────────────────────────

        /** 一行：左边标题+说明，右边控件。 */
        function Row(props) {
            // stacked：控件较宽（例如七个可勾选标签）时改成上下排，
            // 否则 flex 会把左侧标题挤成一字一行。默认仍是左右排。
            const base = props.stacked ? S.rowStacked : S.row;
            const style = props.first ? Object.assign({}, base, S.rowFirst) : base;
            const controlStyle = props.stacked ? S.controlStacked : S.control;
            const text = h("div", { style: props.stacked ? S.rowTextStacked : S.rowText },
                h("div", { style: S.rowLabel }, props.label),
                props.note ? h("div", { style: S.rowNote }, props.note) : null,
            );
            return h("div", { style: style }, text, h("div", { style: controlStyle }, props.children));
        }

        /** 开关。 */
        function Toggle(props) {
            return h("input", {
                type: "checkbox",
                checked: props.checked === true,
                disabled: props.disabled === true,
                onChange: (event) => props.onChange(event.target.checked),
                style: { width: "18px", height: "18px", cursor: props.disabled ? "default" : "pointer" },
            });
        }

        /**
         * 把小数吸附到步长刻度上，并按步长的小数位数收敛结果。
         *
         * 不做这一步会留下 0.30000000000000004 这种浮点噪声：它会被写进设置，
         * 下次读回来又不等于 0.3，界面看着像「改了但没生效」。
         */
        function roundToStep(value, step, min) {
            const base = Number.isFinite(min) ? min : 0;
            const snapped = base + Math.round((value - base) / step) * step;
            const decimals = (String(step).split(".")[1] || "").length;
            const factor = Math.pow(10, decimals);
            return Math.round(snapped * factor) / factor;
        }

        /**
         * 数字输入：受控 + 失焦提交。
         *
         * 不在每次按键时写回：写回是异步的（要过 revision 栅栏），
         * 边打字边提交会把「6」这种中间态也存进去，还会和输入框抢焦点。
         * 所以本地先存草稿，失焦或回车才提交。
         *
         * 提交时**一律吸附到步长网格**（第二十四轮 新-3 收口）：
         *   小数步长（0.05 的阈值）吸附到刻度，否则 0.35 会被 Math.trunc 变成 0 ——
         *   阈值调到 0 就等于把整档过滤关掉，用户却看不出来。
         *   整数步长**也**吸附，而不是取整：`min` 不是 0、或 `step` 不是 1 的字段
         *   （容量上限是 min=512 / step=512、台账条数是 min=10 / step=10）停在网格外时
         *   会被 schema 拒掉，而界面看着像「改了但没生效」。
         *   `roundToStep` 对整数步长天然给出整数（step 的小数位数是 0，收敛因子是 1）。
         */
        function NumberField(props) {
            const [draft, setDraft] = React.useState(String(props.value));
            React.useEffect(() => { setDraft(String(props.value)); }, [props.value]);
            const commit = () => {
                const n = Number(draft);
                if (!Number.isFinite(n)) { setDraft(String(props.value)); return; }
                const step = props.step || 1;
                const bounded = Math.min(props.max, Math.max(props.min, n));
                const clamped = roundToStep(bounded, step, props.min);
                setDraft(String(clamped));
                if (clamped !== props.value) props.onChange(clamped);
            };
            return h("input", {
                type: "number",
                value: draft,
                min: props.min,
                max: props.max,
                step: props.step || 1,
                disabled: props.disabled === true,
                onChange: (event) => setDraft(event.target.value),
                onBlur: commit,
                onKeyDown: (event) => { if (event.key === "Enter") commit(); },
                style: S.input,
            });
        }

        /** 下拉。 */
        function Select(props) {
            return h("select", {
                value: props.value,
                disabled: props.disabled === true,
                onChange: (event) => props.onChange(event.target.value),
                style: S.select,
            }, props.options.map((option) =>
                h("option", { key: option.value, value: option.value }, option.label),
            ));
        }

        /** 一行可勾选的小标签（子代理工具多选用）。 */
        function Chip(props) {
            const on = props.on === true;
            const style = on
                ? Object.assign({}, S.chip, S.chipOn)
                : Object.assign({}, S.chip, S.chipOff);
            return h("div", {
                style: style,
                title: props.note || "",
                onClick: () => props.onToggle(!on),
                role: "checkbox",
                "aria-checked": on,
            },
                h("span", null, on ? "✓" : "＋"),
                h("span", null, props.label),
            );
        }

        // ── 主面板 ────────────────────────────────────────────────────────

        /**
         * 读取命名空间并订阅变化。
         *
         * 数据来自 ctx.configForms（dsh-client-ui-settings 提供的基础服务）；
         * 拿不到就退化成「只读提示」，而不是让整页崩掉 —— 插件在别人的部署里跑，
         * 少一个可选服务不该变成白屏。
         */
        function useScope(scope) {
            const [snap, setSnap] = React.useState(null);
            React.useEffect(() => {
                if (!scope) return undefined;
                const push = () => setSnap(scope.getSnapshot());
                push();
                const off = scope.subscribe(push);
                return () => {
                    if (typeof off === "function") off();
                };
            }, [scope]);
            return snap;
        }

        /**
         * 把宿主的 ConfigForm 适配成本页用惯的 scope 形状。
         *
         * 0.1.7 起设置传输只剩一条：基础设置插件的 configForms 服务，
         * 按「宿主插件行 id」取一份 ConfigForm —— 旧版的 ctx.settingsScope
         * 已经整个没有了（它的服务名在 0.1.7 里搜不到）。
         *
         * ConfigForm 的截面与旧的 scope 几乎一一对应：
         *   { status: 'loading'|'ready'|'unavailable', value, writable, ... }
         * 只有写入不同：旧的 scope.set 收点分路径，新的 ConfigForm 只有
         * mutate([{ op: 'set', path: [...], value }])。
         *
         * 写入被宿主拒绝时 mutate 返回 false 而不是抛错，这里翻成抛错，
         * 好让面板上那套「写入失败：…」的提示照旧显示出来。
         */
        function formToScope(form) {
            return {
                getSnapshot() {
                    const snap = form.getSnapshot();
                    return { status: snap.status, value: snap.value, writable: snap.writable };
                },
                subscribe(listener) {
                    return form.subscribe(listener);
                },
                set(field, value) {
                    const path = String(field).split(".").filter((part) => part !== "");
                    return Promise.resolve(form.mutate([{ op: "set", path: path, value: value }])).then((accepted) => {
                        if (accepted === false) throw new Error("宿主拒绝了这次写入");
                        return accepted;
                    });
                },
            };
        }

        /**
         * 面板外壳：读命名空间、写回、错误与三种兜底状态都在这里，一处一份。
         *
         * 「办公模式」与「记忆系统」是两个一级导航项，但共用同一个设置命名空间 ——
         * 读写接线只该有一套，否则两页的只读提示、错误显示会各自漂移。
         * 渲染函数走具名的 render 属性（而不是 children）：宿主注入的是 React，
         * 单子元素在 children 里是一个数组，取值方式容易记错，具名更直白。
         */
        function SettingsPanel(props) {
            const scope = props.scope;
            const snap = useScope(scope);
            const [writeError, setWriteError] = React.useState(null);

            const write = (field, value) => {
                if (!scope) return;
                setWriteError(null);
                Promise.resolve(scope.set(field, value)).catch((e) => {
                    setWriteError("写入失败：" + (e && e.message ? e.message : String(e)));
                });
            };

            if (!scope) {
                return h("div", { style: S.section }, h("div", { style: S.error },
                    "当前部署没有 configForms 服务，设置页无法读写。"));
            }
            if (snap === null) {
                return h("div", { style: S.section }, h("div", { style: S.status }, "正在读取设置…"));
            }
            if (snap.status === "unavailable") {
                return h("div", { style: S.section },
                    h("div", { style: S.error }, "宿主没有暴露 " + NAMESPACE + " 命名空间，设置页无法读写。"),
                );
            }
            const value = snap.value;
            if (value === undefined) {
                return h("div", { style: S.section }, h("div", { style: S.status }, "正在读取设置…"));
            }
            return props.render({ value: value, disabled: snap.writable === false, write: write, writeError: writeError });
        }

        // ── 联网通道表（浏览器半侧自己的一份）─────────────────────────────
        //
        // 为什么重复一份而不是从服务端 import：浏览器产物是单文件、factory 只拿到
        // require，服务端模块不在它的作用域里。
        //
        // `fields` 里是**写回路径的字面量**：test/client.mjs 逐字核对「schema 里的
        // 每个参数，界面都得有一行」。拼出来的路径（"search.providers." + id + …）
        // 在源码里搜不到，等于把那条守卫废掉 —— 所以宁可写成表。
        // 每条的 fields 与服务端 schema 同形：接缝不吃参数，检索 API 只要 Key，
        // 兼容端点还要端点与模型，自建实例只要地址。
        const CHANNEL_TABLE = [
            {
                id: "seam", label: "宿主 web 服务（接缝）", needsKey: false,
                note: "兜底通道（第三十轮起排在免 Key 抓取之后）：宿主自带公网地址校验、地址钉死、体积与超时上限。没有参数要填（search.providers.seam 是空对象）。",
                fields: [],
            },
            {
                id: "anthropic", label: "Anthropic 兼容 + 原生 web_search", needsKey: true, keyEnv: "DEEPSEEK_API_KEY",
                note: "默认指 DeepSeek 官方端点（模型要支持原生 web_search）；填别的兼容端点也能用。",
                fields: [
                    ["search.providers.anthropic.apiKey", "Key", "password"],
                    ["search.providers.anthropic.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.anthropic.baseURL", "端点", "text"],
                    ["search.providers.anthropic.model", "模型", "text"],
                    ["search.providers.anthropic.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "openai", label: "OpenAI 兼容（chat/completions）", needsKey: true, keyEnv: "OPENAI_API_KEY",
                note: "用 web_search_options 触发联网；端点不支持时工具会明确报「没触发检索」，不会静默返回空。",
                fields: [
                    ["search.providers.openai.apiKey", "Key", "password"],
                    ["search.providers.openai.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.openai.baseURL", "端点", "text"],
                    ["search.providers.openai.model", "模型", "text"],
                    ["search.providers.openai.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "tavily", label: "Tavily Search API", needsKey: true, keyEnv: "TAVILY_API_KEY",
                note: "专给 Agent 用的检索 API，结果自带摘要片段，不需要再逐页取正文。",
                fields: [
                    ["search.providers.tavily.apiKey", "Key", "password"],
                    ["search.providers.tavily.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.tavily.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "brave", label: "Brave Search API", needsKey: true, keyEnv: "BRAVE_API_KEY",
                note: "独立索引；免费档有频率限制，被限流会报「目标站拒绝」。",
                fields: [
                    ["search.providers.brave.apiKey", "Key", "password"],
                    ["search.providers.brave.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.brave.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "bocha", label: "博查 BochaAI（中文长尾好）", needsKey: true, keyEnv: "BOCHA_API_KEY",
                note: "中文覆盖好、国内直连 —— 中文主题建议优先选它。",
                fields: [
                    ["search.providers.bocha.apiKey", "Key", "password"],
                    ["search.providers.bocha.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.bocha.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "exa", label: "Exa（语义检索）", needsKey: true, keyEnv: "EXA_API_KEY",
                note: "按语义找页面（观点、论文、资料页），不是按关键词找官网。",
                fields: [
                    ["search.providers.exa.apiKey", "Key", "password"],
                    ["search.providers.exa.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.exa.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "serper", label: "Serper（Google 结果代理）", needsKey: true, keyEnv: "SERPER_API_KEY",
                note: "拿到的就是 Google 那一页的有机结果；要「按 Google 的口径」时选它。",
                fields: [
                    ["search.providers.serper.apiKey", "Key", "password"],
                    ["search.providers.serper.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.serper.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "searxng", label: "自建 SearXNG", needsKey: false,
                note: "自己的实例，不用外键；要填实例地址（含协议与端口），实例必须允许 JSON 输出。",
                fields: [
                    ["search.providers.searxng.baseURL", "实例地址", "text"],
                    ["search.providers.searxng.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
            {
                id: "duckduckgo", label: "DuckDuckGo（免 Key 兜底）", needsKey: false,
                note: "不需要任何配置的兜底通道：抓 HTML 结果页再解析。被人机校验拦住时换别的通道。",
                fields: [
                    ["search.providers.duckduckgo.baseURL", "结果页地址", "text"],
                    ["search.providers.duckduckgo.timeoutMs", "超时（毫秒）", "number"],
                ],
            },
        ];
        /** 通道 id → 表项。 */
        const CHANNEL_BY_ID = {};
        for (const item of CHANNEL_TABLE) CHANNEL_BY_ID[item.id] = item;

        /**
         * 折叠块：把「常用的几行」和「高级参数」分开。
         *
         * 默认收起，但**内容照渲染**（只是 display:none）—— 一是切换时不重建子树、
         * 草稿不丢；二是单测仍能按 props 找到里面的控件。改一个参数要翻半屏输入框，
         * 是这一页「操作不够简易」的主因。
         */
        function Fold(props) {
            // searching：定位过滤命中折叠块里的行时由 filterRows 置上 —— 命中却看不见
            // 等于没定位到，所以定位期间折叠块一律展开（定位条清空后回到用户自己的状态）。
            const open = props.open === true || props.searching === true;
            return h("div", { style: props.style, "data-fold": props.name },
                h("button", {
                    type: "button",
                    className: "om-btn",
                    style: S.foldHead,
                    "data-fold-toggle": props.name,
                    "aria-expanded": open,
                    onClick: () => props.onToggle(!open),
                },
                    h("span", { style: S.foldCaret }, open ? "▾" : "▸"),
                    h("span", { style: T.strong }, props.label),
                    props.count === undefined ? null : h("span", { style: S.foldCount }, props.count),
                    h("span", { style: S.groupHeadHint }, open ? "收起" : "展开"),
                ),
                h("div", { style: open ? undefined : { display: "none" }, "data-fold-body": props.name }, props.children),
            );
        }

        // ── 分组折叠与「定位」（第三十五轮）───────────────────────────────
        //
        // 设置项长到近百条以后，「改某一格」要么记得它在第几屏，要么一屏一屏翻。
        // 这一层给两个入口：
        //   - 每个分组可收起。默认**全部收起**，进页面先看到一份分组目录
        //     （每组写「几项」），点开哪一组才铺哪一组的控件；
        //   - 顶部一个关键词框。命中的行只留它自己、命中所在的分组自动展开、
        //     命中数写在框下面 —— 输入「代理」「latex」「阈值」直接落到那一行。
        //
        // 过滤在**元素树**上做，不是 DOM 查询：bundle 里没有 DOM（也拿不到宿主
        // 元素的引用），而「哪一行命中」本来就该由这一行自己的 label / note 决定。
        // 三条口径：
        //   1) 比对的是每行的标题与说明。说明里常写着配置路径
        //      （例如 search.providers.anthropic.baseURL），所以按路径也能定位；
        //   2) 空格分开的多个词是**与**关系（中英混排都按这种切法，中文单串照样命中）；
        //   3) 分组标题或说明本身命中时整组保留 —— 搜「音频」要看到那一组全部，
        //      而不是只留下标题命中的那一行。
        //
        // 「命中几项」是**跨分组**的数，只有拿着整页子元素的那一层算得出来，
        // 所以由 SettingsBody 统一数、统一改造分组的 props；各页的分组元素仍写在
        // 原地（标签、说明、控件一行都不搬），SettingsBody 只负责数、过滤、发开合状态。

        /** 关键词切词：空格分开、转小写、去空。 */
        function searchTerms(query) {
            return String(query === undefined || query === null ? "" : query)
                .toLowerCase()
                .split(/\s+/)
                .filter((part) => part !== "");
        }

        /** 「与」匹配：每个词都要出现在文本里（terms 为空一律不命中）。 */
        function matchesTerms(text, terms) {
            if (terms.length === 0) return false;
            const hay = String(text === undefined || text === null ? "" : text).toLowerCase();
            for (const term of terms) {
                if (hay.indexOf(term) === -1) return false;
            }
            return true;
        }

        /** 一行是否命中：标题 + 说明（说明里常带设置路径）。 */
        function rowMatches(node, terms) {
            const props = node.props || {};
            return matchesTerms(String(props.label || "") + " " + String(props.note || ""), terms);
        }

        /** 一个分组里有几条可调项（Row 的个数）：折叠头右边写出来，找组时先看规模。 */
        function countRows(node) {
            if (node === null || node === undefined || typeof node !== "object") return 0;
            if (Array.isArray(node)) {
                let total = 0;
                for (const item of node) total += countRows(item);
                return total;
            }
            let total = node.type === Row ? 1 : 0;
            if (node.props && node.props.children !== undefined) total += countRows(node.props.children);
            return total;
        }

        /** 命中了几行（口径与 filterRows 完全一致，两处必须同源）。 */
        function countMatchedRows(node, terms) {
            if (node === null || node === undefined || typeof node !== "object") return 0;
            if (Array.isArray(node)) {
                let total = 0;
                for (const item of node) total += countMatchedRows(item, terms);
                return total;
            }
            if (node.type === Row) return rowMatches(node, terms) ? 1 : 0;
            return countMatchedRows(node.props ? node.props.children : undefined, terms);
        }

        /**
         * 只留下命中的行。
         *
         * 不命中的行换成 null；包着它们的容器跟着去掉，容器里一条都不剩时整个容器
         * 返回 null（分组据此判断「这一组不用画」）。折在 Fold 里的高级参数命中时
         * 把 Fold 置成 searching —— 否则「定位到了」却看不见。
         */
        function filterRows(node, terms) {
            if (node === null || node === undefined || typeof node === "boolean") return null;
            if (Array.isArray(node)) {
                const kept = [];
                for (const item of node) {
                    const next = filterRows(item, terms);
                    if (next !== null) kept.push(next);
                }
                return kept.length === 0 ? null : kept;
            }
            if (typeof node !== "object") return null;
            if (node.type === Row) return rowMatches(node, terms) ? node : null;
            const children = filterRows(node.props ? node.props.children : undefined, terms);
            if (children === null) return null;
            const props = Object.assign({}, node.props, { children: children });
            if (node.type === Fold) props.searching = true;
            return Object.assign({}, node, { props: props });
        }

        /** 一个元素里的可读文本（标题 / 说明 / 字符串子节点）：给「指路说明」做定位比对。 */
        function collectText(node) {
            if (node === null || node === undefined) return "";
            if (Array.isArray(node)) return node.map(collectText).join(" ");
            if (typeof node === "string" || typeof node === "number") return String(node);
            if (typeof node !== "object") return "";
            const props = node.props || {};
            const own = [props.label, props.note, props.title]
                .filter((item) => typeof item === "string").join(" ");
            return (own + " " + collectText(props.children)).trim();
        }

        /** 复制一个元素并覆盖几个 props（元素不可变，改造只能复制）。 */
        function withProps(node, patch) {
            return Object.assign({}, node, { props: Object.assign({}, node.props, patch) });
        }

        /**
         * 拆开一个分组的子元素：开头那两个「组名 / 说明」块归头部，其余是要收进组体的行。
         *
         * 认的是**样式对象的同一性**（S.groupTitle / S.groupNote 是共享常量），
         * 所以各页的写法完全不用改 —— 仍然照原来的顺序把标题、说明、行作为子元素写下去。
         */
        function splitGroupChildren(children) {
            const list = Array.isArray(children) ? children.slice()
                : (children === undefined || children === null ? [] : [children]);
            const head = [];
            let at = 0;
            while (at < list.length) {
                const item = list[at];
                if (item === null || item === undefined || item === false) { at += 1; continue; }
                const style = item !== null && typeof item === "object" && item.props ? item.props.style : undefined;
                if (style !== S.groupTitle && style !== S.groupNote) break;
                head.push(item);
                at += 1;
            }
            const titleItem = head.find((item) => item.props.style === S.groupTitle);
            const noteItem = head.find((item) => item.props.style === S.groupNote);
            return {
                head: head,
                body: list.slice(at),
                title: titleItem === undefined ? "" : collectText(titleItem.props.children),
                note: noteItem === undefined ? "" : collectText(noteItem.props.children),
            };
        }

        /**
         * 一个可折叠的设置分组。
         *
         * 与 Fold 的分工：Fold 收的是「同一组里的高级参数」；分组收的是整页的骨架。
         * 两者都走「内容照渲染、只是 display:none」——草稿不丢、切换不重建子树，
         * 也让「界面覆盖服务端每个参数」的核对不必先展开每一组。
         */
        function SettingsGroup(props) {
            const open = props.open === true;
            const parts = splitGroupChildren(props.children);
            return h("div", { style: S.group, "data-group": props.id, "data-group-open": open ? "true" : "false" },
                h("button", {
                    type: "button",
                    className: "om-btn",
                    style: S.groupHead,
                    "data-group-toggle": props.id,
                    "aria-expanded": open,
                    onClick: () => props.onToggle(!open),
                },
                    h("span", { style: S.foldCaret }, open ? "▾" : "▸"),
                    h("span", { style: S.groupHeadTitle }, parts.title),
                    props.count === undefined ? null : h("span", { style: S.groupHeadCount }, props.count),
                    h("span", { style: S.groupHeadHint }, open ? "收起" : "展开"),
                ),
                h("div", {
                    style: open ? undefined : { display: "none" },
                    "data-group-body": props.id,
                },
                    parts.head.map((item) => (item.props.style === S.groupTitle ? null : item)),
                    parts.body,
                ),
            );
        }

        /**
         * 分组开合的持久化（键按页分开：办公模式页 / 记忆系统页各记各的）。
         *
         * 默认全部收起；用户点过之后按他的选择记住 —— 下次打开设置页仍是那样。
         * localStorage 拿不到（隐私模式 / 宿主禁掉 / 测试环境）就退回进程内存：
         * 功能照旧，只是不跨刷新。
         */
        const GROUP_STORE_PREFIX = "dsh-office-mode.groups.";
        const GROUP_STORE_MEMORY = {};

        function groupStorage() {
            try {
                const store = typeof window !== "undefined" && window ? window.localStorage : undefined;
                return store && typeof store.getItem === "function" && typeof store.setItem === "function"
                    ? store : null;
            } catch (error) {
                return null;
            }
        }

        function loadGroupState(page) {
            if (GROUP_STORE_MEMORY[page] === undefined) {
                let parsed = null;
                const store = groupStorage();
                if (store !== null) {
                    try {
                        const raw = store.getItem(GROUP_STORE_PREFIX + page);
                        if (raw !== null && raw !== undefined && raw !== "") parsed = JSON.parse(raw);
                    } catch (error) {
                        parsed = null;
                    }
                }
                GROUP_STORE_MEMORY[page] = parsed !== null && typeof parsed === "object" ? parsed : {};
            }
            return GROUP_STORE_MEMORY[page];
        }

        function saveGroupState(page, state) {
            GROUP_STORE_MEMORY[page] = state;
            const store = groupStorage();
            if (store === null) return;
            try {
                store.setItem(GROUP_STORE_PREFIX + page, JSON.stringify(state));
            } catch (error) {
                // 存不下就算了：内存里那份仍然管这一次会话。
            }
        }

        /**
         * 一页的分组开合 + 定位状态（一个设置页一套）。
         *
         * 默认**全部收起**（用户口径：进页面先看到分组目录，找得到再展开）；
         * 点过之后记进 localStorage，下次打开设置页还是那样。
         */
        function useGroupView(page, ids) {
            const [state, setState] = React.useState(loadGroupState(page));
            const [query, setQuery] = React.useState("");
            const terms = searchTerms(query);
            const persist = (next) => {
                setState(next);
                saveGroupState(page, next);
            };
            return {
                page: page,
                query: query,
                setQuery: setQuery,
                terms: terms,
                searching: terms.length > 0,
                isOpen: (id) => state[id] === true,
                toggle: (id, open) => persist(Object.assign({}, state, { [id]: open })),
                setAll: (open) => {
                    const next = {};
                    for (const id of ids) next[id] = open;
                    persist(next);
                },
            };
        }

        /**
         * 定位条：关键词框 + 全部展开 / 全部收起 + 一行状态。
         *
         * 不写「搜索」两个字：这一页搜的是**功能在哪一格**，不是搜内容 ——
         * 用「定位」与记忆面板的搜索框（搜记忆内容）区分开。
         */
        function LocatorBar(props) {
            const view = props.view;
            return h("div", { style: S.locator, "data-locator": view.page },
                h("div", { style: S.locatorRow },
                    h("input", {
                        type: "search",
                        value: view.query,
                        placeholder: "定位功能：输入名称或配置项（如「代理」「latex」「阈值」）",
                        "data-locator-input": view.page,
                        onChange: (event) => view.setQuery(event.target.value),
                        style: S.locatorInput,
                    }),
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-groups-expand": view.page,
                        onClick: () => view.setAll(true),
                    }, "全部展开"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-groups-collapse": view.page,
                        onClick: () => view.setAll(false),
                    }, "全部收起"),
                ),
                h("div", { style: S.locatorStats, "data-locator-stats": view.page },
                    view.searching
                        ? (props.hits === 0
                            ? "没有匹配的项。换个说法，或点「全部展开」自己找。"
                            : "命中 " + props.hits + " 项（共 " + props.total + " 项设置）："
                                + "只留下命中的行，命中的分组已展开。")
                        : "共 " + props.total + " 项设置，默认收起；点分组标题展开，或在这里按名称定位。"),
            );
        }

        /**
         * 设置页外壳：定位条 + 各页写好的那一串子元素。
         *
         * 对每个 SettingsGroup 做三件事：数「几项」、按定位结果过滤、把开合状态发下去；
         * 分组以外的子元素原样透传 —— 只有标了 data-locator-note 的「指路说明」
         * （办公模式页那条「记忆在左侧面板」）会跟着定位一起隐藏。
         */
        function SettingsBody(props) {
            const view = useGroupView(props.page, props.groups);
            const list = Array.isArray(props.children)
                ? props.children
                : (props.children === undefined || props.children === null ? [] : [props.children]);
            let total = 0;
            let hits = 0;
            const prepared = [];
            for (const node of list) {
                if (node === null || node === undefined || node === false) continue;
                const isNote = node.props !== undefined && node.props !== null
                    && node.props["data-locator-note"] !== undefined;
                if (node.type !== SettingsGroup) {
                    if (view.searching && isNote) {
                        if (!matchesTerms(collectText(node), view.terms)) continue;
                        hits += 1;
                    }
                    prepared.push(node);
                    continue;
                }
                const parts = splitGroupChildren(node.props.children);
                const count = countRows(parts.body);
                total += count;
                if (!view.searching) {
                    prepared.push(withProps(node, {
                        open: view.isOpen(node.props.id),
                        onToggle: (next) => view.toggle(node.props.id, next),
                        count: count + " 项",
                    }));
                    continue;
                }
                const titleHit = matchesTerms(parts.title + " " + parts.note, view.terms);
                const matched = titleHit ? count : countMatchedRows(parts.body, view.terms);
                hits += matched;
                if (matched === 0) continue;
                const filtered = titleHit ? parts.body : filterRows(parts.body, view.terms);
                prepared.push(withProps(node, {
                    open: true,
                    onToggle: (next) => view.toggle(node.props.id, next),
                    count: matched + " / " + count + " 项",
                    children: parts.head.concat(filtered === null ? [] : filtered),
                }));
            }
            return h("div", { style: S.section },
                h(LocatorBar, { view: view, hits: hits, total: total }),
                prepared,
            );
        }

        /** 办公模式页的分组（顺序即页面顺序）：用于「全部展开 / 全部收起」。 */
        const OFFICE_GROUP_IDS = ["tools", "search", "subagents", "documents", "python", "av"];
        /** 记忆系统页的分组。 */
        const MEMORY_GROUP_IDS = ["memory", "layers", "scope", "graph", "capacity", "recall", "quota", "migrate"];

        /**
         * 站点清单的可视化编辑器。
         *
         * 写回一律是**整条数组**：`search.sites.entries` 是一个数组设置，服务端按数组
         * 整体校验（见 src/site-catalog.js 的 normalizeSiteEntry），没有「只改第 n 条」
         * 这种协议。所以任何一处改动都经 commit()，免得几个入口各写一份数组。
         *
         * 编辑框用受控草稿 + 失焦提交（与「出口代理」「通道 Key」同一写法）：写回是异步的
         * （要过 revision 栅栏），每键入一次写一次会把「a」「ar」「arx」这些中间态也存进去，
         * 还会和输入框抢焦点。域名要先规范化再查重 —— 重复或不是 `a.b` 形状时只报错、不写入。
         *
         * 两处「空」的语义分开（与服务端 effectiveSiteEntries 一致）：
         *   - `entries === undefined`（拿不到这一项：schema 还没接线 / 宿主没发默认值）
         *     按内置目录显示，而不是画一个空清单；
         *   - `entries === []`（用户把行删光了）就按空清单显示，并且**不**兜回内置目录 ——
         *     否则「删空」这个动作在界面上等于没生效。
         */
        function SiteCatalogEditor(props) {
            const disabled = props.disabled === true;
            const entries = Array.isArray(props.entries)
                ? props.entries
                : SITE_CATALOG.map((item) => Object.assign({}, item));
            /** 新增表单的四个草稿（不进设置，点了「添加」才写回）。 */
            const [draft, setDraft] = React.useState({ type: "custom", domain: "", label: "", note: "" });
            /** 一行错误提示：非法域名 / 重复域名。写入成功即清掉。 */
            const [error, setError] = React.useState(null);
            const patchDraft = (patch) => setDraft((prev) => Object.assign({}, prev, patch));

            /** 内置目录的深拷贝：写回前必须复制，否则会污染常量。 */
            const builtinCopy = () => SITE_CATALOG.map((item) => Object.assign({}, item));

            /** 整条数组写回。 */
            const commit = (next) => {
                setError(null);
                props.write("search.sites.entries", next);
            };
            const replaceAt = (index, patch) => commit(
                entries.map((item, at) => (at === index ? Object.assign({}, item, patch) : item)),
            );
            /** 与数组里相邻的一项交换（跨类型分组时也是整条数组的相邻项）。 */
            const moveAt = (index, delta) => {
                const to = index + delta;
                if (to < 0 || to >= entries.length) return;
                const next = entries.slice();
                const held = next[index];
                next[index] = next[to];
                next[to] = held;
                commit(next);
            };
            /** 改域名：规范化 → 合法性 → 查重，任一条不过就只报错、不动设置。 */
            const renameDomain = (index, raw) => {
                const domain = normalizeSiteDomain(raw);
                if (domain === normalizeSiteDomain(entries[index].domain)) return;
                if (!SITE_DOMAIN_PATTERN.test(domain)) {
                    setError("域名「" + domain + "」不合法：只认 a.b 形状（不带协议、路径与端口）。");
                    return;
                }
                if (entries.some((item, at) => at !== index && normalizeSiteDomain(item.domain) === domain)) {
                    setError("域名「" + domain + "」已经在清单里。");
                    return;
                }
                replaceAt(index, { domain });
            };
            /** 新增：追加到数组末尾（数组顺序就是优先顺序）。 */
            const addEntry = () => {
                const domain = normalizeSiteDomain(draft.domain);
                if (!SITE_DOMAIN_PATTERN.test(domain)) {
                    setError("域名「" + domain + "」不合法：只认 a.b 形状（不带协议、路径与端口）。");
                    return;
                }
                if (entries.some((item) => normalizeSiteDomain(item.domain) === domain)) {
                    setError("域名「" + domain + "」已经在清单里。");
                    return;
                }
                const type = SITE_TYPES.some((item) => item.id === draft.type) ? draft.type : "custom";
                const label = String(draft.label || "").trim();
                commit(entries.concat([{
                    type,
                    domain,
                    label: label === "" ? domain : label,
                    note: String(draft.note || "").trim(),
                    enabled: true,
                }]));
                setDraft({ type, domain: "", label: "", note: "" });
            };

            // 按类型分组，顺序跟着 SITE_TYPES。不认识的类型归到最后一组（自定义）显示：
            // 既不把行藏起来，也不在写回时动它的 type 字段。
            const groups = [];
            for (const type of SITE_TYPES) {
                const items = [];
                for (let index = 0; index < entries.length; index += 1) {
                    const entry = entries[index] || {};
                    const belong = SITE_TYPES.some((candidate) => candidate.id === entry.type)
                        ? entry.type
                        : SITE_TYPES[SITE_TYPES.length - 1].id;
                    if (belong === type.id) items.push({ entry, index });
                }
                if (items.length > 0) groups.push({ type, items });
            }
            const enabledCount = entries.filter((item) => (item || {}).enabled !== false).length;

            const rowFor = (entry, index) => h("div", {
                key: String(entry.domain) + "#" + index,
                style: S.siteRow,
                "data-site-row": entry.domain,
            },
                h("span", { style: S.siteCell, "data-site-toggle": entry.domain },
                    h(Toggle, {
                        checked: entry.enabled !== false,
                        disabled: disabled,
                        onChange: (next) => replaceAt(index, { enabled: next }),
                    }),
                ),
                h("input", {
                    type: "text",
                    defaultValue: String(entry.domain === undefined ? "" : entry.domain),
                    disabled: disabled,
                    "data-site-domain": entry.domain,
                    "aria-label": "域名",
                    onBlur: (event) => renameDomain(index, event.target.value),
                    onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                    style: Object.assign({}, S.siteField, { width: "150px" }),
                }),
                h("input", {
                    type: "text",
                    defaultValue: String(entry.label === undefined ? "" : entry.label),
                    disabled: disabled,
                    placeholder: "显示名",
                    "data-site-label": entry.domain,
                    onBlur: (event) => {
                        const next = String(event.target.value || "").trim();
                        if (next !== String(entry.label === undefined ? "" : entry.label)) replaceAt(index, { label: next });
                    },
                    onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                    style: Object.assign({}, S.siteField, { width: "130px" }),
                }),
                h("input", {
                    type: "text",
                    defaultValue: String(entry.note === undefined ? "" : entry.note),
                    disabled: disabled,
                    placeholder: "一句话说明（为什么收它 / 要注意什么）",
                    "data-site-note": entry.domain,
                    onBlur: (event) => {
                        const next = String(event.target.value || "").trim();
                        if (next !== String(entry.note === undefined ? "" : entry.note)) replaceAt(index, { note: next });
                    },
                    onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                    style: Object.assign({}, S.siteField, { width: "260px" }),
                }),
                h("span", { style: S.siteCell },
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-site-up": entry.domain,
                        disabled: disabled || index === 0,
                        onClick: () => moveAt(index, -1),
                    }, "上移"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-site-down": entry.domain,
                        disabled: disabled || index === entries.length - 1,
                        onClick: () => moveAt(index, 1),
                    }, "下移"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-site-remove": entry.domain,
                        disabled: disabled,
                        onClick: () => commit(entries.filter((item, at) => at !== index)),
                    }, "删除"),
                ),
            );

            return h("div", { style: S.siteList, "data-site-editor": "true" },
                h("div", { style: S.quickBar },
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-site-reset": "true",
                        disabled: disabled,
                        onClick: () => commit(builtinCopy()),
                    }, "恢复内置默认"),
                    h("span", { style: S.status }, "清单 " + entries.length + " 条，启用 " + enabledCount + " 条"),
                ),
                error === null ? null : h("div", { style: S.error, "data-site-error": "true" }, error),
                entries.length === 0
                    ? h("div", { style: S.quickBar, "data-site-empty": "true" },
                        h("span", { style: S.rowNote }, "当前清单为空：这次不限定任何站点。"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.button,
                            "data-site-load": "true",
                            disabled: disabled,
                            onClick: () => commit(builtinCopy()),
                        }, "载入内置目录"),
                    )
                    : null,
                groups.map((group) => h("div", { key: group.type.id, "data-site-group": group.type.id },
                    h("div", { style: S.siteGroupHead }, group.type.name),
                    h("div", { style: S.siteGroupNote }, group.type.note),
                    h("div", { style: S.siteList }, group.items.map((item) => rowFor(item.entry, item.index))),
                )),
                h("div", { style: S.siteAdd, "data-site-add-form": "true" },
                    h("span", { style: S.siteCell, "data-site-add-type": "true" },
                        h(Select, {
                            value: draft.type,
                            disabled: disabled,
                            options: SITE_TYPES.map((type) => ({ value: type.id, label: type.name })),
                            onChange: (next) => patchDraft({ type: next }),
                        }),
                    ),
                    h("input", {
                        type: "text",
                        value: draft.domain,
                        disabled: disabled,
                        placeholder: "域名（例 example.com）",
                        "data-site-add-domain": "true",
                        onChange: (event) => patchDraft({ domain: event.target.value }),
                        onKeyDown: (event) => { if (event.key === "Enter") addEntry(); },
                        style: Object.assign({}, S.siteField, { width: "180px" }),
                    }),
                    h("input", {
                        type: "text",
                        value: draft.label,
                        disabled: disabled,
                        placeholder: "显示名（留空用域名）",
                        "data-site-add-label": "true",
                        onChange: (event) => patchDraft({ label: event.target.value }),
                        onKeyDown: (event) => { if (event.key === "Enter") addEntry(); },
                        style: Object.assign({}, S.siteField, { width: "160px" }),
                    }),
                    h("input", {
                        type: "text",
                        value: draft.note,
                        disabled: disabled,
                        placeholder: "一句话说明",
                        "data-site-add-note": "true",
                        onChange: (event) => patchDraft({ note: event.target.value }),
                        onKeyDown: (event) => { if (event.key === "Enter") addEntry(); },
                        style: Object.assign({}, S.siteField, { width: "220px" }),
                    }),
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-site-add": "true",
                        disabled: disabled,
                        onClick: addEntry,
                    }, "添加"),
                ),
                h("div", { style: S.siteHint },
                    "列表顺序就是优先顺序。域名只按域名限定（site:域名）；说明只在这里显示，不进检索请求。"),
            );
        }

        function OfficeSettingsSection(props) {
            /** 只放「界面本地状态」：折叠开关与正在编辑哪条通道的参数（都不写设置）。 */
            const [ui, setUi] = React.useState({ paramsChannel: "anthropic", advancedSearch: false, advancedDocs: false, toolsHint: null });
            const patchUi = (patch) => setUi(Object.assign({}, ui, patch));
            return h(SettingsPanel, { scope: props.scope, render: ({ value, disabled, write, writeError }) => {
            const tools = value.tools || {};
            const search = value.search || {};
            const builtin = search.builtin || {};
            const providers = search.providers || {};
            const preprocess = search.preprocess || {};
            // 站点清单：`entries` 只在真的拿到数组时才算「有值」（undefined → 内置目录，
            // [] → 空清单），这个区分在 SiteCatalogEditor 里处理。
            const sites = search.sites || {};
            const documents = value.documents || {};
            const python = value.python || {};
            const av = value.av || {};
            const subagentModel = value.subagentModel || {};
            const subagentTools = Array.isArray(search.subagentTools) ? search.subagentTools : [];
            const hasWrite = subagentTools.indexOf("write") !== -1;
            // 正在编辑参数的通道：默认跟着「联网通道」走，但允许单独切换 ——
            // 想先把 Tavily 的 Key 填好、但把通道留着 auto 的人不该被迫改通道。
            const paramsChannel = CHANNEL_BY_ID[ui.paramsChannel] !== undefined
                ? ui.paramsChannel
                : (search.provider && CHANNEL_BY_ID[search.provider] !== undefined ? search.provider : "anthropic");
            const paramsMeta = CHANNEL_BY_ID[paramsChannel];
            const paramsValue = providers[paramsChannel] || {};
            const order = Array.isArray(search.providerOrder) && search.providerOrder.length > 0
                ? search.providerOrder
                : CHANNEL_TABLE.map((item) => item.id);
            const enabledChannels = CHANNEL_TABLE.filter((item) => {
                const entry = providers[item.id] || {};
                if (item.needsKey !== true) return true;
                return String(entry.apiKey || "") !== "" || String(entry.apiKeyEnv || "") !== "";
            });

            return h(SettingsBody, { page: "office", groups: OFFICE_GROUP_IDS },
                writeError !== null ? h("div", { key: "error", style: S.error }, writeError) : null,
                disabled
                    ? h("div", { key: "readonly", style: S.status }, "当前连接以只读方式同步设置（memory 模式），改动不会持久化。")
                    : null,

                // ── 工具开关 ──
                h(SettingsGroup, { key: "tools", id: "tools" },
                    h("div", { style: S.groupTitle }, "工具开关"),
                    h("div", { style: S.groupNote },
                        "关掉的工具不出现在模型面前，也不占每次请求的 schema 开销。改动即时生效，不用重启。"),
                    h("div", { style: S.quickBar },
                        h("button", {
                            type: "button", className: "om-btn", style: S.button, "data-tools-all": "true",
                            disabled: disabled,
                            onClick: () => {
                                for (const row of TOOL_ROWS) write("tools." + row.key, true);
                                patchUi({ toolsHint: "已全部打开。" });
                            },
                        }, "全部打开"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.button, "data-tools-none": "true",
                            disabled: disabled,
                            onClick: () => {
                                for (const row of TOOL_ROWS) write("tools." + row.key, false);
                                patchUi({ toolsHint: "已全部关闭（模型面前不再有 office_* 工具）。" });
                            },
                        }, "全部关闭"),
                        ui.toolsHint === null ? null : h("span", { style: S.status }, ui.toolsHint),
                    ),
                    TOOL_ROWS.map((row, index) => h(Row, {
                        key: row.key,
                        label: row.label,
                        note: row.note,
                        first: index === 0,
                    }, h(Toggle, {
                        checked: tools[row.key] !== false,
                        disabled: disabled,
                        onChange: (next) => write("tools." + row.key, next),
                    }))),
                ),

                // ── 记忆 ──
                //
                // 记忆有自己的面板（「记忆系统」），这里只留一句指路：
                // 同一个命名空间在两处渲染同一份设置容易让人以为改的是两个东西。
                h("div", { key: "memory", "data-locator-note": "memory", style: S.group },
                    h("div", { style: S.groupTitle }, "记忆"),
                    h("div", { style: S.groupNote },
                        "三层记忆（热记忆 / 台账 / 归档）的开关、容量与迁移说明在左侧「记忆系统」面板里。"),
                ),
                // ── 检索编排 ──
                h(SettingsGroup, { key: "search", id: "search" },
                    h("div", { style: S.groupTitle }, "检索编排"),
                    h("div", { style: S.groupNote },
                        "控制检索子代理能做什么、一次铺开多少。渠道多时不要一次全开，容易被限流。"
                        + "第三十轮起资料搜索走插件的网页抓取通道（office_web_search / office_web_fetch，免 Key）；"
                        + "三方检索 API 不在默认通道顺序里，要显式加回。"),
                    h(Row, { label: "执行引擎", note: "auto = 先派子代理，跑不了自动改用内置检索；builtin = 只用内置检索；subagent = 只用子代理", first: true },
                        h(Select, {
                            value: search.engine === undefined ? "auto" : search.engine,
                            disabled: disabled,
                            options: SEARCH_ENGINE_OPTIONS,
                            onChange: (next) => write("search.engine", next),
                        }),
                    ),
                    // ── 联网通道（第二十轮）──
                    //
                    // 这一组是「不再只有 DeepSeek 一条路」的落点：选一条、填 Key，
                    // 两步就能用起来。参数细节收在下面的折叠块里。
                    h(Row, {
                        label: "联网通道",
                        note: "auto = 按「尝试顺序」逐条试，第一条成功的就用它；固定成某一条时只走那一条。"
                            + "免 Key 的 DuckDuckGo / 自建 SearXNG 适合还没配 Key 的机器。",
                        stacked: true,
                    },
                        h(Select, {
                            value: search.provider === undefined ? "auto" : search.provider,
                            disabled: disabled,
                            options: [{ value: "auto", label: "auto（自动挑，推荐）" }].concat(
                                CHANNEL_TABLE.map((item) => ({ value: item.id, label: item.label })),
                            ),
                            onChange: (next) => write("search.provider", next),
                        }),
                        h("div", { style: S.rowNote },
                            search.provider === undefined || search.provider === "auto"
                                ? "按顺序试；已配置的通道：" + enabledChannels.map((item) => item.label).join("、")
                                : "只用这一条。"),
                    ),
                    h(Row, {
                        label: "出口代理",
                        note: "留空 = 联网按本进程原有出口直连。填了（形如 http://127.0.0.1:7897，只认 http / https）"
                            + "让本插件自己发起的请求走它：装了私有分派器时改值当场生效；拿不到 undici 的部署会退回"
                            + "「整个进程」的环境变量那条路，那种情况下清空要重启 DSH 才回直连。",
                    },
                        h("input", {
                            type: "text",
                            defaultValue: search.proxy === undefined ? "" : search.proxy,
                            disabled: disabled,
                            placeholder: "http://127.0.0.1:7897",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== search.proxy) write("search.proxy", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "填哪条通道的参数", note: "下面那一组参数只针对选中的通道（改这里不影响「联网通道」的选择）" },
                        h(Select, {
                            value: paramsChannel,
                            disabled: disabled,
                            options: CHANNEL_TABLE.map((item) => ({ value: item.id, label: item.label })),
                            onChange: (next) => patchUi({ paramsChannel: next }),
                        }),
                    ),
                    h("div", { style: S.callout },
                        h("div", { style: S.rowNote }, paramsMeta.note),
                        paramsMeta.fields.length === 0
                            ? h("div", { style: S.rowNote }, "这条通道没有参数要填 —— 它在就用。")
                            : h(Row, {
                                label: paramsMeta.needsKey ? "通道 Key" : "通道参数",
                                note: paramsMeta.needsKey
                                    ? "可以直接粘 Key（写进这条通道的 Key 字面量）；也可以只填环境变量名，让插件去取。"
                                    : "这一条通道唯一的必填项是地址。",
                                first: true,
                            },
                                h("input", {
                                    type: "password",
                                    defaultValue: String(paramsValue.apiKey || ""),
                                    disabled: disabled || !paramsMeta.needsKey,
                                    placeholder: paramsMeta.needsKey ? "粘 Key（留空则用环境变量）" : "不需要 Key",
                                    "data-channel-key": paramsChannel,
                                    onBlur: (event) => {
                                        const next = String(event.target.value || "").trim();
                                        if (next !== String(paramsValue.apiKey || "")) write("search.providers." + paramsChannel + ".apiKey", next);
                                    },
                                    onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                    style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                                }),
                            ),
                    ),
                    h(Fold, {
                        name: "search-advanced",
                        label: "高级参数（通道端点 / 模型 / 超时，内置检索细项）",
                        count: "改了才生效",
                        open: ui.advancedSearch,
                        onToggle: (next) => patchUi({ advancedSearch: next }),
                    },
                        h("div", { style: S.groupNote },
                            "通道参数只影响它自己那条通道。anthropic 的这几项与「内置检索」是同一件事的两个入口，"
                            + "同名键以通道参数为准。"),
                        paramsMeta.fields.map(([path, label, kind], index) => {
                            const key = path.split(".").pop();
                            const current = paramsValue[key];
                            return h(Row, { key: path, label: label, note: path, first: index === 0 },
                                kind === "number"
                                    ? h(NumberField, {
                                        value: Number.isFinite(current) ? current : 30_000,
                                        min: 5000, max: 300000, step: 5000, disabled: disabled,
                                        onChange: (next) => write(path, next),
                                    })
                                    : h("input", {
                                        type: kind === "password" ? "password" : "text",
                                        defaultValue: current === undefined ? "" : String(current),
                                        disabled: disabled,
                                        onBlur: (event) => {
                                            const next = String(event.target.value || "").trim();
                                            if (next !== String(current === undefined ? "" : current)) write(path, next);
                                        },
                                        onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                        style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                                    }),
                            );
                        }),
                        // ── 内置检索（anthropic 通道的细项）──
                        h("div", { style: Object.assign({}, S.groupTitle, { marginTop: "16px" }) }, "内置检索（Anthropic 兼容通道的细项）"),
                        h("div", { style: S.groupNote },
                            "maxUses / maxTokens 决定模型在一次检索里能调几次原生 web_search；maxResults 与 fetchPages "
                            + "决定每个渠道最多要几条来源、最多打开几页取正文。取正文的体积、超时与跳转上限也在这里。"),
                        ...BUILTIN_ROWS.map((row, index) => h(Row, {
                            key: row.path,
                            label: row.label,
                            note: row.note,
                            first: index === 0,
                        },
                            row.kind === "number"
                                ? h(NumberField, {
                                    value: builtin[row.key] === undefined ? row.min : builtin[row.key],
                                    min: row.min, max: row.max, step: row.step, disabled: disabled,
                                    onChange: (next) => write(row.path, next),
                                })
                                : h("input", {
                                    type: row.kind === "password" ? "password" : "text",
                                    defaultValue: builtin[row.key] === undefined ? "" : builtin[row.key],
                                    disabled: disabled,
                                    onBlur: (event) => {
                                        const next = String(event.target.value || "").trim();
                                        if (next !== builtin[row.key]) write(row.path, next);
                                    },
                                    onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                    style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                                }),
                            row.kind === "number" && row.unit
                                ? h("span", { style: S.unit }, row.unit)
                                : null,
                        )),
                    ),
                    // ── auto 的尝试顺序（点标签加/移）──
                    h(Row, {
                        label: "auto 尝试顺序",
                        note: "点一下把通道加进 / 移出顺序。sequence 越靠前越先试；没配 Key 的通道当场失败（不发请求），所以顺序长也不拖时间。",
                        stacked: true,
                    },
                        h("div", { style: S.chips, "data-provider-order": order.join(",") }, CHANNEL_TABLE.map((item) => {
                            const at = order.indexOf(item.id);
                            return h(Chip, {
                                key: item.id,
                                label: item.label + (at === -1 ? "" : " · " + (at + 1)),
                                note: "点一下" + (at === -1 ? "加进" : "移出") + "尝试顺序",
                                on: at !== -1,
                                onToggle: (next) => {
                                    const list = order.slice();
                                    const index = list.indexOf(item.id);
                                    if (next && index === -1) list.push(item.id);
                                    if (!next && index !== -1) list.splice(index, 1);
                                    write("search.providerOrder", list.length === 0 ? ["seam"] : list);
                                },
                            });
                        })),
                    ),
                    // ── 网页预处理（第二十轮）──
                    h(Row, {
                        label: "网页预处理",
                        note: "取回来的页面先清洗再进上下文：去脚本样式与导航骨架、挑主容器、丢样板行、去重。"
                            + "off = 回到原样（清单型页面建议 off 或 plain）。",
                        first: true,
                    },
                        h(Select, {
                            value: preprocess.mode === undefined ? "article" : preprocess.mode,
                            disabled: disabled,
                            options: [
                                { value: "article", label: "article（去导航与样板，推荐）" },
                                { value: "plain", label: "plain（只做行级清洗）" },
                                { value: "off", label: "off（原样返回）" },
                            ],
                            onChange: (next) => write("search.preprocess.mode", next),
                        }),
                    ),
                    h(Row, { label: "丢掉样板行", note: "cookie 条 / 登录注册 / 订阅分享 / 版权声明这类成行的样板" },
                        h(Toggle, {
                            checked: preprocess.dropBoilerplate !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.preprocess.dropBoilerplate", next),
                        }),
                    ),
                    h(Row, { label: "重复行去重", note: "同一条重复出现的行只留第一次（导航条收成文本后最常撞这条）" },
                        h(Toggle, {
                            checked: preprocess.dedupeLines !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.preprocess.dedupeLines", next),
                        }),
                    ),
                    h(Row, { label: "短行阈值", note: "短于这个长度、又不像句子的行当导航丢掉。0 = 关掉这条（清单型页面建议关）" },
                        h(NumberField, {
                            value: preprocess.minLineChars === undefined ? 12 : preprocess.minLineChars,
                            min: 0, max: 80, step: 1, disabled: disabled,
                            onChange: (next) => write("search.preprocess.minLineChars", next),
                        }),
                        h("span", { style: S.unit }, "字"),
                    ),
                    h(Row, { label: "抽出标题与出处", note: "把标题 / 作者 / 发布时间放进预处理报告（不进正文）" },
                        h(Toggle, {
                            checked: preprocess.keepTitle !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.preprocess.keepTitle", next),
                        }),
                    ),
                    // ── 站点清单（第三十四轮：优先检索的可视化编辑）──
                    //
                    // 三段开关 + 一张清单。清单本体在 SiteCatalogEditor 里，写回一律是
                    // 整条数组（search.sites.entries），没有逐条改的协议。
                    h("div", { style: Object.assign({}, S.groupTitle, { marginTop: "18px" }) }, "站点清单"),
                    h("div", { style: S.groupNote },
                        "先把查询限定到清单里的站点（site:域名）去问：命中的来源排在前面，并按类型汇总。"
                        + "清单顺序就是优先顺序；一条都没有时这一次不做限定。"
                        + "限定轮一无所获时按下面的开关决定退不退回泛搜 —— 不因为清单查不到就当成「没有资料」。"),
                    h(Row, {
                        label: "站点优先检索",
                        note: "把查询限定到清单里的站点去问，命中的来源排前面并按类型汇总。关掉则所有检索都不限定来源。",
                        first: true,
                    },
                        h(Toggle, {
                            checked: sites.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.sites.enabled", next),
                        }),
                    ),
                    h(Row, {
                        label: "退回泛搜",
                        note: "这些站点被墙、没收录或站点改版时，退回不限定来源的泛搜，并把哪些站点空手照实写出来。"
                            + "关掉则限定轮没有结果就是没有结果（清单被墙时容易误判成「查不到」）。",
                    },
                        h(Toggle, {
                            checked: sites.fallback !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.sites.fallback", next),
                        }),
                    ),
                    h(Row, {
                        label: "单次最多限定",
                        note: "一次检索最多限定几个站点，按清单顺序自上而下取（1-8）。限定越多越慢，也越容易被目标站限流。",
                    },
                        h(NumberField, {
                            value: Number.isFinite(sites.maxPerCall) ? sites.maxPerCall : SITE_MAX_PER_CALL[0],
                            min: SITE_MAX_PER_CALL[0], max: SITE_MAX_PER_CALL[1], step: 1, disabled: disabled,
                            onChange: (next) => write("search.sites.maxPerCall", next),
                        }),
                        h("span", { style: S.unit }, "个"),
                    ),
                    h(Row, {
                        label: "清单内容",
                        note: "按类型分组；行上可以直接改域名、显示名与说明（失焦或回车才写回）。"
                            + "域名只按域名限定，说明只在这里显示、不进检索请求。",
                        stacked: true,
                    },
                        h(SiteCatalogEditor, {
                            entries: sites.entries,
                            disabled: disabled,
                            write: write,
                        }),
                    ),
                    h(Row, { label: "子代理可用工具", note: "加得越多，子代理越容易跑偏去干检索以外的事", stacked: true },
                        h("div", { style: S.chips }, SUBAGENT_TOOLS.map((tool) => h(Chip, {
                            key: tool.id,
                            label: tool.label,
                            note: tool.note,
                            on: subagentTools.indexOf(tool.id) !== -1,
                            onToggle: (next) => {
                                const set = subagentTools.slice();
                                const at = set.indexOf(tool.id);
                                if (next && at === -1) set.push(tool.id);
                                if (!next && at !== -1) set.splice(at, 1);
                                write("search.subagentTools", set);
                            },
                        }))),
                    ),
                    !hasWrite
                        ? h("div", { style: S.error }, "警告：子代理没有 write 工具，检索结果将无法落盘。")
                        : null,
                    h(Row, { label: "并发子代理数", note: "同时铺开几个检索子代理（1-8）" },
                        h(NumberField, {
                            value: search.maxParallel, min: 1, max: 8, step: 1, disabled: disabled,
                            onChange: (next) => write("search.maxParallel", next),
                        }),
                        h("span", { style: S.unit }, "个"),
                    ),
                    h(Row, { label: "结果条数上限", note: "解析结果时最多列出多少条结论（1-40）" },
                        h(NumberField, {
                            value: search.resultLimit, min: 1, max: 40, step: 1, disabled: disabled,
                            onChange: (next) => write("search.resultLimit", next),
                        }),
                        h("span", { style: S.unit }, "条"),
                    ),
                    h(Row, { label: "渠道数上限", note: "一次提纲最多列几个渠道（1-12）" },
                        h(NumberField, {
                            value: search.maxChannels, min: 1, max: 12, step: 1, disabled: disabled,
                            onChange: (next) => write("search.maxChannels", next),
                        }),
                        h("span", { style: S.unit }, "个"),
                    ),
                    h(Row, { label: "结果目录", note: "检索提纲与结果文件的存放位置（相对会话工作目录）" },
                        h("input", {
                            type: "text",
                            defaultValue: search.outputDir,
                            disabled: disabled,
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== "" && next !== search.outputDir) write("search.outputDir", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "跨源核对提醒", note: "开启后，只有单一来源背书的结论会被点名要求补检索" },
                        h(Toggle, {
                            checked: search.requireCrossSource !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.requireCrossSource", next),
                        }),
                    ),
                    h(Row, { label: "平台失败时兜底", note: "平台直连被网络拦住时，自动改用网页搜索搜该平台内容" },
                        h(Toggle, {
                            checked: search.fallbackOnPlatformError !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.fallbackOnPlatformError", next),
                        }),
                    ),
                    // 内置检索的细项搬进了上面的「高级参数」折叠块 —— 它其实就是
                    // anthropic 通道的参数，放在两处会让人以为改的是两个东西。
                ),

                // ── 后台任务 Agent 的模型路由 ──
                //
                // 本插件的后台任务只有一处：检索三步走的派工子代理。这一组只影响
                // 它们，不动主对话 —— 两件事混在一个下拉里会让人以为改了主模型。
                h(SettingsGroup, { key: "subagents", id: "subagents" },
                    h("div", { style: S.groupTitle }, "子代理模型"),
                    h("div", { style: S.groupNote },
                        "检索子代理用哪个模型。只影响后台派工，不影响主对话。"),
                    h(Row, { label: "模型路由", note: "inherit = 跟随主会话；fixed = 用下面指定的模型", first: true },
                        h(Select, {
                            value: subagentModel.mode === undefined ? "inherit" : subagentModel.mode,
                            disabled: disabled,
                            options: SUBAGENT_MODEL_MODE_OPTIONS,
                            onChange: (next) => write("subagentModel.mode", next),
                        }),
                    ),
                    h(Row, { label: "Provider 路由名", note: "留空 = 只覆盖 model，仍留在主会话那条路由上" },
                        h("input", {
                            type: "text",
                            defaultValue: subagentModel.provider,
                            disabled: disabled,
                            placeholder: "留空 = 不覆盖",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== subagentModel.provider) write("subagentModel.provider", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "模型 id", note: "留空按 inherit 处理。写错不会弄崩插件：派工失败会明确报错" },
                        h("input", {
                            type: "text",
                            defaultValue: subagentModel.model,
                            disabled: disabled,
                            placeholder: "留空 = 不覆盖",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== subagentModel.model) write("subagentModel.model", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "240px", textAlign: "left" }),
                        }),
                    ),
                ),

                // ── 文档与缓存 ──
                h(SettingsGroup, { key: "documents", id: "documents" },
                    h("div", { style: S.groupTitle }, "文档与缓存"),
                    h("div", { style: S.groupNote },
                        "默认视觉主题、批处理的资源上限，以及中间产物与 PDF 渲染的规则。"),
                    h(Row, { label: "默认主题", note: "plain 是素色网格：Excel 只留细边框、Word 白纸黑字，不加任何背景填充", first: true },
                        h(Select, {
                            value: documents.defaultTheme,
                            disabled: disabled,
                            options: THEMES.map((theme) => ({ value: theme.id, label: theme.label })),
                            onChange: (next) => write("documents.defaultTheme", next),
                        }),
                    ),
                    h(Row, { label: "脚本超时", note: "单次 office_run 脚本的超时。批量生成很多文件时可以调大" },
                        h(NumberField, {
                            value: documents.scriptTimeoutMs, min: 1000, max: 600000, step: 1000, disabled: disabled,
                            onChange: (next) => write("documents.scriptTimeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "ms"),
                    ),
                    h(Row, { label: "脚本字符上限", note: "单次脚本的字符上限。脚本很大时调大，但过大会拖慢错误定位" },
                        h(NumberField, {
                            value: documents.maxScriptChars, min: 1000, max: 4000000, step: 1000, disabled: disabled,
                            onChange: (next) => write("documents.maxScriptChars", next),
                        }),
                        h("span", { style: S.unit }, "字"),
                    ),
                    h(Row, { label: "缓存目录", note: "中间产物目录。跨调用保留，按下面的存活时间自动清理" },
                        h("input", {
                            type: "text",
                            defaultValue: documents.cacheDir,
                            disabled: disabled,
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== "" && next !== documents.cacheDir) write("documents.cacheDir", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "保留中间产物", note: "调用结束后保留缓存。保留才能跨调用复用（例如渲染好的 PDF 页面图）；关掉则每次调用结束就清空" },
                        h(Toggle, {
                            checked: documents.keepCache !== false,
                            disabled: disabled,
                            onChange: (next) => write("documents.keepCache", next),
                        }),
                    ),
                    h(Row, { label: "缓存存活时间", note: "中间产物多少分钟没被碰过就清掉（每次调用开始时清理）。0 = 不按时间清理" },
                        h(NumberField, {
                            value: documents.cacheTtlMinutes === undefined ? 720 : documents.cacheTtlMinutes,
                            min: 0, max: 43200, step: 10, disabled: disabled,
                            onChange: (next) => write("documents.cacheTtlMinutes", next),
                        }),
                        h("span", { style: S.unit }, "分钟"),
                    ),

                    // ── PDF 渲染 ──
                    h(Row, { label: "PDF 渲染分辨率", note: "PDF 渲染成图片时的默认 DPI。手写笔记建议 120-200；越大越慢、图片越大" },
                        h(NumberField, {
                            value: documents.pdfDpi === undefined ? 120 : documents.pdfDpi,
                            min: 36, max: 400, step: 1, disabled: disabled,
                            onChange: (next) => write("documents.pdfDpi", next),
                        }),
                        h("span", { style: S.unit }, "dpi"),
                    ),
                    h(Row, { label: "PDF 单次页数上限", note: "单次 office.pdf.pages() 最多渲染多少页。渲染是重活，防止一次卡住整个脚本" },
                        h(NumberField, {
                            value: documents.pdfMaxPages === undefined ? 20 : documents.pdfMaxPages,
                            min: 1, max: 500, step: 1, disabled: disabled,
                            onChange: (next) => write("documents.pdfMaxPages", next),
                        }),
                        h("span", { style: S.unit }, "页"),
                    ),
                    h(Row, { label: "PDF 渲染引擎", note: "auto = 自动探测（PyMuPDF / poppler / MuPDF / Ghostscript）。一个都没有时 office.pdf.pages() 会明确报错" },
                        h(Select, {
                            value: documents.pdfEngine === undefined ? "auto" : documents.pdfEngine,
                            disabled: disabled,
                            options: [
                                { value: "auto", label: "自动探测（推荐）" },
                                { value: "fitz", label: "PyMuPDF（fitz）" },
                                { value: "pdftoppm", label: "poppler pdftoppm" },
                                { value: "pdftocairo", label: "poppler pdftocairo" },
                                { value: "mutool", label: "MuPDF mutool" },
                                { value: "gs", label: "Ghostscript" },
                            ],
                            onChange: (next) => write("documents.pdfEngine", next),
                        }),
                    ),

                    // ── LaTeX 论文 ──
                    h(Row, { label: "LaTeX 编译入口", note: "auto = 有 latexmk 就用它（自动处理 bibtex 与重复编译），否则用 xelatex。thuthesis 需要 fontspec，pdflatex 不可用" },
                        h(Select, {
                            value: documents.texEngine === undefined ? "auto" : documents.texEngine,
                            disabled: disabled,
                            options: [
                                { value: "auto", label: "自动探测（推荐）" },
                                { value: "latexmk", label: "latexmk" },
                                { value: "xelatex", label: "xelatex" },
                                { value: "lualatex", label: "lualatex" },
                            ],
                            onChange: (next) => write("documents.texEngine", next),
                        }),
                    ),
                    h(Row, { label: "LaTeX 编译超时", note: "单次编译的超时。学位论文要跑三四遍 xelatex 加 bibtex，目录长的大论文要调大" },
                        h(NumberField, {
                            value: documents.texTimeoutMs === undefined ? 300000 : documents.texTimeoutMs,
                            min: 30000, max: 1800000, step: 10000, disabled: disabled,
                            onChange: (next) => write("documents.texTimeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "毫秒"),
                    ),
                    h(Row, { label: "thuthesis 模板目录", note: "含 thuthesis.cls 与两个校徽的目录。留空自动探测：工作目录里的 thuthesis* 目录，其次 TeX Live 自带的版本" },
                        h("input", {
                            type: "text",
                            defaultValue: documents.texTemplateDir,
                            disabled: disabled,
                            placeholder: "留空 = 自动探测",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== documents.texTemplateDir) write("documents.texTemplateDir", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                        }),
                    ),
                ),

                // ── Python 计算与绘图 ──
                //
                // 服务端 `python` 组（src/settings.js）从第十一轮起就存在，但界面一直没有
                // 这一组 —— 四个参数只能改配置文件。第三十八轮补上：模式与其它组一致
                // （Row + 控件 + 失焦写回），四项全部走 `python.*` 点号路径。
                //
                // 三项留空的语义都要说清（与 src/python.js 的探测一致）：
                //   解释器留空 = 自动探测 PATH 里的 python / python3 / py；
                //   产物目录留空 = 缓存目录下的 python/out；
                //   单项超时不设「留空」，因为它有确定默认值（120 秒）。
                h(SettingsGroup, { key: "python", id: "python" },
                    h("div", { style: S.groupTitle }, "Python 计算与绘图"),
                    h("div", { style: S.groupNote },
                        "办公模式里唯一的编程出口：脚本里用 office.python.run / file / check 做科学计算与绘图，"
                        + "产物落缓存目录。命令行仍然是关掉的，插件也不替用户装包。"),
                    h(Row, { label: "启用", note: "关掉后 office.python 不可用（报错说明在设置里关了），其余能力不受影响", first: true },
                        h(Toggle, {
                            checked: python.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("python.enabled", next),
                        }),
                    ),
                    h(Row, { label: "解释器路径", note: "留空自动探测 PATH 里的 python / python3 / py；也可以填绝对路径。应用商店的占位程序不算可用" },
                        h("input", {
                            type: "text",
                            defaultValue: python.bin || "",
                            disabled: disabled,
                            placeholder: "留空 = 自动探测",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== python.bin) write("python.bin", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "300px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "运行超时", note: "单次 Python 运行的超时（毫秒）。跑长计算或大批量绘图时调大；office_run 的脚本超时管不到 await，Python 的有效上限就是这一项" },
                        h(NumberField, {
                            value: python.timeoutMs === undefined ? 120000 : python.timeoutMs,
                            min: 5000, max: 900000, step: 5000, disabled: disabled,
                            onChange: (next) => write("python.timeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "毫秒"),
                    ),
                    h(Row, { label: "产物目录", note: "Python 产物（图 / 表 / 数据）的目录，相对缓存目录；脚本里拿 OUT_DIR 就是它的绝对路径。留空 = python/out" },
                        h("input", {
                            type: "text",
                            defaultValue: python.outDir || "",
                            disabled: disabled,
                            placeholder: "留空 = python/out",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== python.outDir) write("python.outDir", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                        }),
                    ),
                ),

                // ── 音频与视频 ──
                //
                // 四样能力（ffmpeg / ffprobe / SenseVoice 模型 / sherpa-onnx 运行时）
                // 都由插件按这里的路径自动探测；缺任何一样时 office.av 会明确报错，
                // 报错文案与 office.av.check() 的 hint 是同一份。
                h(SettingsGroup, { key: "av", id: "av" },
                    h("div", { style: S.groupTitle }, "音频与视频"),
                    h("div", { style: S.groupNote },
                        "录音 / 视频转文字（本机 SenseVoice 离线识别，带逐句时间戳），视频按时间点抽帧交给读图。"
                        + "四样能力缺任何一样都会明确报错：ffmpeg、ffprobe、模型、sherpa-onnx 运行时。"),
                    h(Row, { label: "启用", note: "关掉后 office.av 不可用（报错说明在设置里关了），其余能力不受影响", first: true },
                        h(Toggle, {
                            checked: av.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("av.enabled", next),
                        }),
                    ),
                    h(Row, { label: "ffmpeg 路径", note: "留空自动探测：PATH，其次是 C:\\app\\ffmpeg\\bin 这类常见目录" },
                        h("input", {
                            type: "text",
                            defaultValue: av.ffmpegPath || "",
                            disabled: disabled,
                            placeholder: "留空 = 自动探测",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== av.ffmpegPath) write("av.ffmpegPath", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "ffprobe 路径", note: "留空 = 先按 ffmpeg 同目录找，再走 PATH（ffprobe 与 ffmpeg 一般装在一起）" },
                        h("input", {
                            type: "text",
                            defaultValue: av.ffprobePath || "",
                            disabled: disabled,
                            placeholder: "留空 = 自动探测",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== av.ffprobePath) write("av.ffprobePath", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "SenseVoice 模型目录", note: "含 model.int8.onnx 与 tokens.txt。留空 = 语音输入下载的那份（$DSH_HOME/speech-to-text/sensevoice/models/sensevoice-onnx）" },
                        h("input", {
                            type: "text",
                            defaultValue: av.modelDir || "",
                            disabled: disabled,
                            placeholder: "留空 = 语音输入下载的那份",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== av.modelDir) write("av.modelDir", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "300px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "转写语言", note: "auto = 让模型自己判（中 / 粤 / 英 / 日 / 韩）；固定成某一种时混语种录音可能被强行翻成那一种" },
                        h(Select, {
                            value: av.language === undefined ? "auto" : av.language,
                            disabled: disabled,
                            options: AV_LANGUAGE_OPTIONS,
                            onChange: (next) => write("av.language", next),
                        }),
                    ),
                    h(Row, { label: "单块秒数", note: "模型一次吃的音频长度。官方那条链 131 秒封顶，这里默认 120 秒留余量" },
                        h(NumberField, {
                            value: av.chunkSeconds === undefined ? 120 : av.chunkSeconds,
                            min: 10, max: 600, step: 10, disabled: disabled,
                            onChange: (next) => write("av.chunkSeconds", next),
                        }),
                        h("span", { style: S.unit }, "秒"),
                    ),
                    h(Row, { label: "块间重叠", note: "相邻两块的重叠秒数（默认 0.5）。块不再完全切开：跨在切点上的短音在同一块里解完；跨得更深的长句由边界回退兜住，不丢字也不重复。上限是半块" },
                        h(NumberField, {
                            value: av.overlapSeconds === undefined ? 0.5 : av.overlapSeconds,
                            min: 0, max: 30, step: 0.5, disabled: disabled,
                            onChange: (next) => write("av.overlapSeconds", next),
                        }),
                        h("span", { style: S.unit }, "秒"),
                    ),
                    h(Row, { label: "单次时长上限", note: "超过就报错并要求先切段（内存保护）。切块只解决单块大小，不改变这条上限" },
                        h(NumberField, {
                            value: av.maxSeconds === undefined ? 3600 : av.maxSeconds,
                            min: 10, max: 21600, step: 10, disabled: disabled,
                            onChange: (next) => write("av.maxSeconds", next),
                        }),
                        h("span", { style: S.unit }, "秒"),
                    ),
                    h(Row, { label: "处理超时", note: "单次解码 / 抽帧 / 转写的超时。长会议录音调大" },
                        h(NumberField, {
                            value: av.timeoutMs === undefined ? 300000 : av.timeoutMs,
                            min: 5000, max: 1800000, step: 5000, disabled: disabled,
                            onChange: (next) => write("av.timeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "毫秒"),
                    ),
                    h(Row, { label: "模型精度", note: "int8 约 228 MB（默认，够快够准）；fp32 约 894 MB，本机没有实测收益" },
                        h(Select, {
                            value: av.precision === undefined ? "int8" : av.precision,
                            disabled: disabled,
                            options: [
                                { value: "int8", label: "int8（默认）" },
                                { value: "fp32", label: "fp32（更慢更大）" },
                            ],
                            onChange: (next) => write("av.precision", next),
                        }),
                    ),
                    h(Row, { label: "推理线程", note: "默认 2：再往上收益递减，而且会和宿主抢 CPU" },
                        h(NumberField, {
                            value: av.threads === undefined ? 2 : av.threads,
                            min: 1, max: 16, step: 1, disabled: disabled,
                            onChange: (next) => write("av.threads", next),
                        }),
                        h("span", { style: S.unit }, "线程"),
                    ),
                    h(Row, { label: "默认抽帧张数", note: "不给 at / every / count 时视频抽几张。抽出来的图交给 read_image" },
                        h(NumberField, {
                            value: av.frames === undefined ? 6 : av.frames,
                            min: 1, max: 60, step: 1, disabled: disabled,
                            onChange: (next) => write("av.frames", next),
                        }),
                        h("span", { style: S.unit }, "张"),
                    ),
                    h(Row, { label: "单次抽帧上限", note: "防止一次抽几百张把缓存与上下文撑爆" },
                        h(NumberField, {
                            value: av.maxFrames === undefined ? 12 : av.maxFrames,
                            min: 1, max: 60, step: 1, disabled: disabled,
                            onChange: (next) => write("av.maxFrames", next),
                        }),
                        h("span", { style: S.unit }, "张"),
                    ),
                    h(Row, { label: "抽帧格式", note: "jpg 体积小（默认，适合实拍）；png 无损，图上文字更清楚" },
                        h(Select, {
                            value: av.frameFormat === undefined ? "jpg" : av.frameFormat,
                            disabled: disabled,
                            options: [
                                { value: "jpg", label: "jpg（默认）" },
                                { value: "png", label: "png（无损）" },
                            ],
                            onChange: (next) => write("av.frameFormat", next),
                        }),
                    ),
                    h(Row, { label: "抽帧宽度", note: "缩放后的宽度（像素）。0 = 原分辨率；4K 视频建议填 1280 之类，图小、读起来也快" },
                        h(NumberField, {
                            value: av.frameWidth === undefined ? 0 : av.frameWidth,
                            min: 0, max: 4096, step: 2, disabled: disabled,
                            onChange: (next) => write("av.frameWidth", next),
                        }),
                        h("span", { style: S.unit }, "px"),
                    ),
                ),

                h("div", { key: "tail", style: S.groupNote },
                    "数值超出范围会被宿主拒绝并自动回读；改动即时生效。"),
            );
            } });
        }

        // ── 记忆系统面板 ──────────────────────────────────────────────────
        //
        // 独立的一级导航项：原来那个「记忆系统」入口由 mnemon 提供，本插件现在自带
        // 三层记忆，就由这一页接住。面板只做三件事：说清三层是什么、调节范围
        // （开关 / 目录 / 容量 / 提示），以及说清怎么把 mnemon 的记忆迁过来。
        //
        // 迁移**没有按钮**：面板只有配置读写通道（configForms），没有执行通道；
        // 与其做一个点了没反应的按钮，不如给确切的命令与工具调用。
        function MemorySettingsSection(props) {
            return h(SettingsPanel, { scope: props.scope, render: ({ value, disabled, write, writeError }) => {
            const memory = value.memory || {};
            // 三个嵌套组各自取出来：它们不是扁平键，写回也必须用点号路径
            // （write("memory.layers.hot", …)），所以读的时候也要走同一层。
            const layers = memory.layers || {};
            const recallQuality = memory.recallQuality || {};
            const quota = memory.quota || {};
            const layerNote = "三层各管一段时间尺度：热记忆常驻（用户偏好与项目约定），台账按次登记（每份交付物一条），归档只读（下沉的旧条目与滚动的旧台账）。记忆不是缓存，不会被 TTL 清掉。";

            return h(SettingsBody, { page: "memory", groups: MEMORY_GROUP_IDS },
                writeError !== null ? h("div", { key: "error", style: S.error }, writeError) : null,
                disabled
                    ? h("div", { key: "readonly", style: S.status }, "当前连接以只读方式同步设置（memory 模式），改动不会持久化。")
                    : null,

                h(SettingsGroup, { key: "memory", id: "memory" },
                    h("div", { style: S.groupTitle }, "记忆系统"),
                    h("div", { style: S.groupNote }, layerNote),
                    h(Row, { label: "开启记忆", note: "关掉后 office_memory 工具、记忆提示段与 office_run 的自动台账一起停", first: true },
                        h(Toggle, {
                            checked: memory.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("memory.enabled", next),
                        }),
                    ),
                    h(Row, { label: "记忆目录", note: "相对会话工作目录。一个项目一份记忆：换目录就是换一套记忆" },
                        h("input", {
                            type: "text",
                            defaultValue: memory.dir,
                            disabled: disabled,
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== "" && next !== memory.dir) write("memory.dir", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "自动登记台账", note: "office_run 每写出一份文档就自动记一条（路径、格式、主题、复检统计）。关掉后只能手工登记" },
                        h(Toggle, {
                            checked: memory.autoLedger !== false,
                            disabled: disabled,
                            onChange: (next) => write("memory.autoLedger", next),
                        }),
                    ),
                    h(Row, { label: "写入系统提示", note: "加一段静态说明（记忆在哪、怎么用，不含记忆内容）。办公模式的 persona 是完整的，那一段在那里会被丢掉" },
                        h(Toggle, {
                            checked: memory.promptHint !== false,
                            disabled: disabled,
                            onChange: (next) => write("memory.promptHint", next),
                        }),
                    ),
                ),

                // ── 记忆层拓扑 ──
                //
                // 每层一个独立开关：关掉只是不再读 / 不再写，数据留在盘上。
                // 这是「先别用」而不是「删掉」，所以不给一个总开关一了百了。
                h(SettingsGroup, { key: "layers", id: "layers" },
                    h("div", { style: S.groupTitle }, "记忆层拓扑"),
                    h("div", { style: S.groupNote },
                        "每层独立开关。关掉只是不再读 / 不再写，已有数据不会被删除，重新打开就回来。"),
                    LAYER_ROWS.map((row, index) => h(Row, {
                        key: row.key,
                        label: row.label,
                        note: row.note,
                        first: index === 0,
                    }, h(Toggle, {
                        checked: layers[row.key] !== false,
                        disabled: disabled,
                        onChange: (next) => write(row.path, next),
                    }))),
                ),

                // ── 记忆范围（跨项目层） ──
                h(SettingsGroup, { key: "scope", id: "scope" },
                    h("div", { style: S.groupTitle }, "记忆范围"),
                    h("div", { style: S.groupNote },
                        "决定记忆是「一个项目一份」还是「所有项目共用一份」。跨项目层解决的是换工作目录就要重记一遍偏好这件事。"),
                    h(Row, { label: "记忆放哪儿", note: "workspace = 只跟当前工作目录；global = 所有项目共用；both = 两者并存", first: true },
                        h(Select, {
                            value: memory.scope === undefined ? "workspace" : memory.scope,
                            disabled: disabled,
                            options: MEMORY_SCOPE_OPTIONS,
                            onChange: (next) => write("memory.scope", next),
                        }),
                    ),
                    h(Row, { label: "全局层目录", note: "留空 = $DSH_HOME/.office/memory。相对路径相对用户主目录 —— 全局层不该跟着工作目录搬家" },
                        h("input", {
                            type: "text",
                            defaultValue: memory.globalDir,
                            disabled: disabled,
                            placeholder: "留空 = $DSH_HOME/.office/memory",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== memory.globalDir) write("memory.globalDir", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "240px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "用户偏好范围", note: "memory = 跟着上面的范围；global = 始终落全局层，换工作目录不必重记一遍" },
                        h(Select, {
                            value: memory.userScope === undefined ? "memory" : memory.userScope,
                            disabled: disabled,
                            options: USER_SCOPE_OPTIONS,
                            onChange: (next) => write("memory.userScope", next),
                        }),
                    ),
                ),

                // ── 图关系与主动记录 ──
                h(SettingsGroup, { key: "graph", id: "graph" },
                    h("div", { style: S.groupTitle }, "图关系与主动记录"),
                    h("div", { style: S.groupNote },
                        "两个可选的记忆增强，对应 mnemon 的 link / auto-capture。都只改行为，不动已有数据。"),
                    h(Row, { label: "图关系", note: "允许给条目建双向类型化关系并沿关系检索（office_memory 的 link / related 与实体视图）。关掉不影响已有关系文件", first: true },
                        h(Toggle, {
                            checked: memory.links !== false,
                            disabled: disabled,
                            onChange: (next) => write("memory.links", next),
                        }),
                    ),
                    h(Row, { label: "主动记录", note: "在工具反馈里提示「用户这轮说出的偏好与纠正当场记一条」，不必等他说「记住」。这是一句指引，不是自主记录器" },
                        h(Toggle, {
                            checked: memory.autoCapture !== false,
                            disabled: disabled,
                            onChange: (next) => write("memory.autoCapture", next),
                        }),
                    ),
                ),

                h(SettingsGroup, { key: "capacity", id: "capacity" },
                    h("div", { style: S.groupTitle }, "容量与归档"),
                    h("div", { style: S.groupNote },
                        "容量满了向下沉，不静默丢：热记忆装不下时最旧、最不重要的条目进归档，并在 MEMORY.md 留一条指路条目；台账超出条数上限时最旧的滚成月度归档摘要。"),
                    h(Row, { label: "用户偏好上限", note: "热记忆里「用户偏好」的容量，默认 4096 字节", first: true },
                        h(NumberField, {
                            value: memory.userLimitBytes === undefined ? 4096 : memory.userLimitBytes,
                            min: 512, max: 65536, step: 512, disabled: disabled,
                            onChange: (next) => write("memory.userLimitBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "项目记忆上限", note: "热记忆里「项目与环境」的容量，默认 10240 字节" },
                        h(NumberField, {
                            value: memory.projectLimitBytes === undefined ? 10240 : memory.projectLimitBytes,
                            min: 1024, max: 262144, step: 1024, disabled: disabled,
                            onChange: (next) => write("memory.projectLimitBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "投影预算", note: "一次投影里热记忆条目正文的字节上限（0 = 不限），默认 4096。超出的条目不丢：折叠行报条数与读取入口" },
                        h(NumberField, {
                            value: memory.projectionBudgetBytes === undefined ? 4096 : memory.projectionBudgetBytes,
                            min: 0, max: 131072, step: 1024, disabled: disabled,
                            onChange: (next) => write("memory.projectionBudgetBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "单条上限", note: "一条热记忆最多多少字节（0 = 不限），默认 2800。超过就拒绝写入并提示压缩：热记忆是动笔前照办的清单，不是文档仓库" },
                        h(NumberField, {
                            value: memory.entryLimitBytes === undefined ? 2800 : memory.entryLimitBytes,
                            min: 0, max: 131072, step: 100, disabled: disabled,
                            onChange: (next) => write("memory.entryLimitBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "单条建议线", note: "超过这个字节数只在写入回执里提醒（不拒绝），默认 1200；0 = 关掉提醒" },
                        h(NumberField, {
                            value: memory.entryHintBytes === undefined ? 1200 : memory.entryHintBytes,
                            min: 0, max: 131072, step: 100, disabled: disabled,
                            onChange: (next) => write("memory.entryHintBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "台账条数上限", note: "台账最多保留多少条，超出的最旧记录滚成月度归档摘要" },
                        h(NumberField, {
                            value: memory.ledgerLimit === undefined ? 500 : memory.ledgerLimit,
                            min: 10, max: 10000, step: 10, disabled: disabled,
                            onChange: (next) => write("memory.ledgerLimit", next),
                        }),
                        h("span", { style: S.unit }, "条"),
                    ),
                    h(Row, { label: "归档保留数", note: "归档最多保留多少个摘要文件。归档也有界：最旧的摘要是唯一会被真正删掉的东西" },
                        h(NumberField, {
                            value: memory.archiveKeep === undefined ? 60 : memory.archiveKeep,
                            min: 1, max: 1000, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.archiveKeep", next),
                        }),
                        h("span", { style: S.unit }, "个"),
                    ),
                ),

                // ── 召回质量 ──
                //
                // 管的是「一次检索返回什么」。注意这里的分档分数是「命中了查询里
                // 几成的词」，不是向量余弦相似度 —— 不写清楚，用户会按向量检索的
                // 直觉把阈值调到 0.8，然后一条都查不出来。
                h(SettingsGroup, { key: "recall", id: "recall" },
                    h("div", { style: S.groupTitle }, "召回质量"),
                    h("div", { style: S.groupNote },
                        "一次检索返回什么：按相关度分档、低相关的丢掉。分数是「命中了查询里几成的词」，不是向量相似度。"),
                    h(Row, { label: "分档策略", note: "strict-v1 = 按相关度分档并丢弃低分；off = 只按关键词排序，不丢结果", first: true },
                        h(Select, {
                            value: recallQuality.policy === undefined ? "strict-v1" : recallQuality.policy,
                            disabled: disabled,
                            options: RECALL_POLICY_OPTIONS,
                            onChange: (next) => write("memory.recallQuality.policy", next),
                        }),
                    ),
                    h(Row, { label: "低分阈值", note: "命中词占比低于它算「未知」档（0-1，步长 0.05）" },
                        h(NumberField, {
                            value: recallQuality.lowScoreThreshold === undefined ? 0.25 : recallQuality.lowScoreThreshold,
                            min: 0, max: 1, step: 0.05, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.lowScoreThreshold", next),
                        }),
                    ),
                    h(Row, { label: "高分阈值", note: "达到它算「高分」档，直接采纳。必须大于低分阈值，写反了整组会退回默认" },
                        h(NumberField, {
                            value: recallQuality.highScoreThreshold === undefined ? 0.6 : recallQuality.highScoreThreshold,
                            min: 0, max: 1, step: 0.05, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.highScoreThreshold", next),
                        }),
                    ),
                    h(Row, { label: "候选倍数", note: "先取「条数上限 × 这个倍数」个候选再分档，防止低分结果把高分结果挤出候选池" },
                        h(NumberField, {
                            value: recallQuality.candidateMultiplier === undefined ? 5 : recallQuality.candidateMultiplier,
                            min: 1, max: 10, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.candidateMultiplier", next),
                        }),
                        h("span", { style: S.unit }, "倍"),
                    ),
                    h(Row, { label: "中档最多几条", note: "「中档」最多采纳几条（0-40）" },
                        h(NumberField, {
                            value: recallQuality.maxMediumResults === undefined ? 6 : recallQuality.maxMediumResults,
                            min: 0, max: 40, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.maxMediumResults", next),
                        }),
                        h("span", { style: S.unit }, "条"),
                    ),
                    h(Row, { label: "未知档最多几条", note: "「未知档」最多采纳几条（0 = 一条都不要）" },
                        h(NumberField, {
                            value: recallQuality.maxUnknownResults === undefined ? 4 : recallQuality.maxUnknownResults,
                            min: 0, max: 40, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.maxUnknownResults", next),
                        }),
                        h("span", { style: S.unit }, "条"),
                    ),
                ),

                // ── 每回合配额 ──
                //
                // 管的是「一个回合能查几次」。与上一组的分工必须写在说明里：两组
                // 都调「检索」，不写清楚会被当成同一个东西。
                h(SettingsGroup, { key: "quota", id: "quota" },
                    h("div", { style: S.groupTitle }, "每回合配额"),
                    h("div", { style: S.groupNote },
                        "防止「没查到就换个词再查」把上下文填满。0 = 该类不限制。回合号从会话日志的 turn/start 读出来。"),
                    h(Row, { label: "首次检索", note: "一个回合里第一次带关键词的记忆检索（0 = 不限制）", first: true },
                        h(NumberField, {
                            value: quota.recallPerTurn === undefined ? 1 : quota.recallPerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.recallPerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                    h(Row, { label: "换词细化", note: "同一回合里后续的换词细化次数。用完就拒绝，并说明原因" },
                        h(NumberField, {
                            value: quota.recallRefinePerTurn === undefined ? 1 : quota.recallRefinePerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.recallRefinePerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                    h(Row, { label: "知识块检索", note: "一个回合里的 kb（来源原文）检索次数。与记忆检索分开算：语料与代价都不同" },
                        h(NumberField, {
                            value: quota.kbSearchPerTurn === undefined ? 2 : quota.kbSearchPerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.kbSearchPerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                    h(Row, { label: "图关系遍历", note: "一个回合里沿关系跳转的次数" },
                        h(NumberField, {
                            value: quota.relatedPerTurn === undefined ? 1 : quota.relatedPerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.relatedPerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                ),

                h(SettingsGroup, { key: "migrate", id: "migrate" },
                    h("div", { style: S.groupTitle }, "从 mnemon 迁移"),
                    h("div", { style: S.groupNote },
                        "把 mnemon 的 runtime 热记忆（含全局 USER.md）迁进本目录的热记忆，长期记忆（insights）连同类别、重要度、标签、实体与关系边数迁进归档。幂等、只增不改，重复跑不会产生第二份。"),
                    h(Row, { label: "先看一眼", note: "--dry-run 只统计不落盘", first: true, stacked: true },
                        h("code", { style: S.code }, "node scripts/migrate-mnemon.mjs --dry-run"),
                    ),
                    h(Row, { label: "执行迁移", note: "也可以在会话里让助手调用 office_memory({ action: 'migrate' })", stacked: true },
                        h("code", { style: S.code }, "node scripts/migrate-mnemon.mjs"),
                    ),
                    h(Row, { label: "检索迁进来的长期记忆", note: "归档层按关键词检索；热记忆会随 office_help 一起投影出来", stacked: true },
                        h("code", { style: S.code }, "office_memory({ action: 'read', layer: 'archive', query: '季度汇报' })"),
                    ),
                ),

                h("div", { key: "tail", style: S.groupNote },
                    "本面板只读写设置；记忆内容本身由助手用 office_memory 读取与检索（三层记忆的完整说明见 office_help({ topic: 'memory' })）。"),
            );
            } });
        }

        // ── 导航图标 ──────────────────────────────────────────────────────
        //
        // 设置对话框的导航项要一个图标，和「记忆系统」「Wallpaper Engine」并列。
        // 用内联 SVG：客户端 bundle 是单文件产物，外链图标没法跟着走。
        // 图案是一张纸 + 一支笔，对应「办公文档」。
        function OfficeIcon() {
            return h("svg", {
                width: 16, height: 16, viewBox: "0 0 24 24", fill: "none",
                stroke: "currentColor", strokeWidth: 1.8,
                strokeLinecap: "round", strokeLinejoin: "round",
                "aria-hidden": "true",
            },
                h("path", { d: "M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h7" }),
                h("path", { d: "M14 3v5h5" }),
                h("path", { d: "M14 3l5 5" }),
                h("path", { d: "M18.5 13.5l2.6 2.6a1.4 1.4 0 0 1 0 2l-1.6 1.6a1.4 1.4 0 0 1-2 0L14.9 17" }),
                h("path", { d: "M14.9 17l-2.4 1 .9-2.4 4.1-4.1 2.5 2.5z" }),
            );
        }

        /** 导航项：图标 + 名称。 */
        function OfficeNavLabel() {
            return h("span", { style: { display: "flex", alignItems: "center", gap: "10px" } },
                h(OfficeIcon, null),
                h("span", null, "办公模式"),
            );
        }

        /**
         * 记忆系统的图标：三层叠起来的盘子（热记忆 / 台账 / 归档）。
         * 与办公模式那张「纸 + 笔」区分开，免得两个导航项看起来是同一个东西。
         */
        function MemoryIcon() {
            return h("svg", {
                width: 16, height: 16, viewBox: "0 0 24 24", fill: "none",
                stroke: "currentColor", strokeWidth: 1.8,
                strokeLinecap: "round", strokeLinejoin: "round",
                "aria-hidden": "true",
            },
                h("ellipse", { cx: 12, cy: 6, rx: 7, ry: 3 }),
                h("path", { d: "M5 6v6c0 1.66 3.13 3 7 3s7-1.34 7-3V6" }),
                h("path", { d: "M5 12v6c0 1.66 3.13 3 7 3s7-1.34 7-3v-6" }),
            );
        }

        /** 导航项：图标 + 名称。 */
        function MemoryNavLabel() {
            return h("span", { style: { display: "flex", alignItems: "center", gap: "10px" } },
                h(MemoryIcon, null),
                h("span", null, "记忆系统"),
            );
        }

        // ── 记忆浏览面板 ──────────────────────────────────────────────────
        //
        // 与「记忆系统」设置页是两件事：那一页调的是**配置**，这一格看的是**内容**。
        // 数据走宿主注册的只读路由（MEMORY_VIEW_PATH），因为配置通道给不了工作目录
        // 里的文件内容。

        /**
         * 侧栏「记忆」入口的图标。
         *
         * owner props 是 { size, active }：size 必须用上（折叠态与展开态会给不同
         * 尺寸），active 只影响配色 —— 用 currentColor 让宿主的高亮态自然生效，
         * 不去猜它用什么颜色。
         */
        function MemoryPanelIcon(props) {
            const size = Number.isFinite(props && props.size) ? props.size : 16;
            return h("svg", {
                width: size, height: size, viewBox: "0 0 16 16", fill: "none",
                stroke: "currentColor", strokeWidth: 1.4,
                strokeLinecap: "round", strokeLinejoin: "round",
                "aria-hidden": "true",
            },
                h("ellipse", { cx: 8, cy: 4, rx: 5, ry: 2 }),
                h("path", { d: "M3 4v4c0 1.1 2.24 2 5 2s5-.9 5-2V4" }),
                h("path", { d: "M3 8v4c0 1.1 2.24 2 5 2s5-.9 5-2V8" }),
            );
        }

        /** 条目上的一枚小徽标（来源 / 重要度 / 格式 / 月份 …）。 */
        function Badges(props) {
            const list = (props.items || []).filter((text) => text !== "" && text !== null && text !== undefined);
            if (list.length === 0) return null;
            return h("div", { style: S.browserMeta },
                list.map((text, index) => h("span", { key: index, style: S.badge }, String(text))),
            );
        }

        /** 来源徽标：记忆可能是工作区层的，也可能是跨项目的全局层。 */
        function originLabel(origin) {
            return origin === "global" ? "全局" : "工作区";
        }

        /**
         * 卡片底：极浅的同色渐变，向右 34% 处淡出到表面色。
         *
         * 这是 mnemon 的卡片签名手法之一 —— 有颜色但不脏，也不需要多一条边框。
         * color-mix 的两个参数都自带字面兜底，所以令牌缺失时这一行仍然是合法声明
         * （不合法的话整个 background 会被丢掉，卡片会变透明）。
         */
        function tintedSurface(color) {
            return "linear-gradient(90deg, color-mix(in srgb, " + color + " 6%, " + M.surface + "), " + M.surface + " 34%)";
        }

        /** 数值夹取（缩放倍率、视图偏移都要用）。 */
        function clampNumber(value, min, max) {
            const number = Number(value);
            if (!Number.isFinite(number)) return min;
            return Math.min(max, Math.max(min, number));
        }

        /**
         * 系统是不是要求「减少动效」。
         *
         * 光靠 CSS 的 `@media (prefers-reduced-motion)` 管不住 JS 发起的平滑滚动
         * （`scrollTo({behavior:'smooth'})` 不看 CSS 的 `scroll-behavior`），
         * 所以那一条在代码里也要读一次。读不到 matchMedia 就当「不要求」——
         * 测试与 SSR 环境里没有它，那不是异常。
         */
        function prefersReducedMotion() {
            try {
                const view = typeof window !== "undefined" && window ? window : null;
                if (!view || typeof view.matchMedia !== "function") return false;
                const query = view.matchMedia("(prefers-reduced-motion: reduce)");
                return query !== null && query !== undefined && query.matches === true;
            } catch (error) {
                return false;
            }
        }

        /**
         * 把一段文本放进剪贴板。
         *
         * 读 `window.navigator` 而不是裸 `navigator`：bundle 的包装函数只拿到
         * window / require / console / fetch 四个名字，裸全局在测试与受限环境里
         * 可能根本不存在。拿不到就如实回 false，界面据此提示「手动选中复制」，
         * 而不是画一个点了没反应的按钮。
         */
        function copyText(text) {
            const value = String(text === undefined || text === null ? "" : text);
            try {
                const view = typeof window !== "undefined" && window ? window : null;
                const nav = view && view.navigator ? view.navigator : null;
                const clip = nav && nav.clipboard ? nav.clipboard : null;
                if (!clip || typeof clip.writeText !== "function") return Promise.resolve(false);
                return Promise.resolve(clip.writeText(value)).then(() => true, () => false);
            } catch (error) {
                return Promise.resolve(false);
            }
        }

        /** 一条记忆的「唯一命中片段」：空白压平，太长就截断（oldText 只要唯一即可）。 */
        function oldTextOf(item) {
            const text = String(item && item.content ? item.content : "").replace(/\s+/g, " ").trim();
            return text.length <= 40 ? text : text.slice(0, 40);
        }

        /** 生成一句可读的 office_memory 调用；参数顺序固定，便于肉眼核对。 */
        function memoryCall(action, args) {
            const parts = ["action: " + JSON.stringify(action)];
            for (const pair of args) {
                if (pair[1] === undefined || pair[1] === null || pair[1] === "") continue;
                parts.push(pair[0] + ": " + JSON.stringify(pair[1]));
            }
            return "office_memory({ " + parts.join(", ") + " })";
        }

        /**
         * 条目上的写入动作。
         *
         * **面板不落盘**：点一下只是把一条精确的 `office_memory` 调用生成出来，
         * 放进面板顶部的待办区，由用户确认、复制、在会话里发送。三条理由：
         *   1. 记忆一旦写进去就跨会话影响后面所有回合，该由用户先看见「要记什么」
         *      再点头 —— 与回复旁的「存入记忆」同一个口径；
         *   2. 浏览器半侧**没有任何写入通道**：打包的客户端入口只拿到 require，
         *      既没有 `host.call`（那是动态 cordis 包的），客户端服务目录里也没有
         *      composer / inputActions（只有 layout / locale / sessions / slots /
         *      theme / timer / uiWorkspace / workspaces），`main` slot 的
         *      standardProps 同样不含 inputActions —— 面板既改不了输入框草稿、
         *      也提交不了消息；
         *   3. 唯一的现成通道是只读端点，放宽它的信任边界等于把本地 HTTP 端点变成
         *      可写面，不做。
         * 所以这里只生成**指令文本**：唯一命中校验、下沉、落盘都还在 office_memory 里。
         */
        const ITEM_ACTIONS = {
            hot: [
                {
                    id: "replace", label: "改写", hint: "把这条热记忆改成新的说法",
                    build: (item) => ({
                        title: "改写这条热记忆",
                        text: [
                            "用 office_memory 改掉这条热记忆（oldText 唯一命中即可，命中多条会被拒绝）：",
                            memoryCall("replace", [
                                ["target", item.target === "user" ? "user" : "project"],
                                ["oldText", oldTextOf(item)],
                                ["content", "<改成什么，把这里替换掉>"],
                            ]),
                        ].join("\n"),
                    }),
                },
                {
                    id: "remove", label: "忘记", hint: "删掉这条热记忆",
                    build: (item) => ({
                        title: "删掉这条热记忆",
                        text: [
                            "用 office_memory 删掉这条热记忆，删完把删掉的那条原文贴回来：",
                            memoryCall("remove", [
                                ["target", item.target === "user" ? "user" : "project"],
                                ["oldText", oldTextOf(item)],
                            ]),
                        ].join("\n"),
                    }),
                },
            ],
            ledger: [
                {
                    id: "log", label: "改用途", hint: "重新登记这份产物的用途",
                    build: (item) => ({
                        title: "改这条台账的用途",
                        text: [
                            "用 office_memory 重新登记这份交付物（同 path 会更新那一条）：",
                            memoryCall("log", [
                                ["path", item.path],
                                ["format", item.format],
                                ["theme", item.theme],
                                ["purpose", "<新用途，把这里替换掉>"],
                            ]),
                        ].join("\n"),
                    }),
                },
            ],
            links: [
                {
                    id: "unlink", label: "删关系", hint: "删掉这条关系",
                    build: (item) => ({
                        title: "删掉这条关系",
                        text: [
                            "用 office_memory 删掉这条关系：",
                            memoryCall("unlink", [["id", item.id]]),
                        ].join("\n"),
                    }),
                },
            ],
            entities: [
                {
                    id: "related", label: "找相关", hint: "沿关系看这个实体牵到的条目",
                    build: (item) => {
                        const refs = Array.isArray(item.refs) ? item.refs : [];
                        const first = refs.length > 0 ? refs[0] : null;
                        return {
                            title: "看这个实体牵到的条目",
                            text: [
                                "用 office_memory 看「" + String(item.name || "") + "」牵到的条目"
                                + "（实体自己没有 id，先用它引用到的第一条起跳，depth 默认 1）：",
                                memoryCall("related", [["id", first ? first.id : "<条目 id>"], ["depth", 1]]),
                            ].join("\n"),
                        };
                    },
                },
            ],
        };

        /** 归档是只读层：给它一句说明，而不是一个点了会报错的动作。 */
        const ARCHIVE_READONLY_NOTE = "归档是只读层：条目由热记忆下沉与台账滚动产生，面板不提供写入动作。";

        /** 「记一条」：面板只给骨架，内容由用户写。 */
        function addInstruction() {
            return {
                title: "记一条新记忆",
                text: [
                    "用 office_memory 记一条（target / importance 你自己判断，只记稳定的偏好、约定与事实）：",
                    memoryCall("add", [
                        ["target", "project"],
                        ["content", "<要记的内容：一两句、自包含>"],
                    ]),
                ].join("\n"),
            };
        }

        /**
         * 一条条目铺开时的完整内容（按 Tab 取不同字段）。
         *
         * 列表卡片里的正文是**折叠**的（`maxHeight: 9em`），完整内容有两条出口：
         * 卡片上的「展开」，以及面板底部那条详情。两者都不用原生 `title` ——
         * 系统气泡不可样式化、键盘触发不了，位置也由浏览器决定。
         */
        function detailTextOf(item, tab) {
            const entry = item || {};
            if (tab === "hot") return String(entry.content || "");
            if (tab === "ledger") {
                const lines = [String(entry.path || "")];
                if (entry.purpose) lines.push("用途：" + entry.purpose);
                if (entry.format) lines.push("格式：" + entry.format + (entry.theme ? " · 主题 " + entry.theme : ""));
                if (Array.isArray(entry.outline) && entry.outline.length > 0) lines.push("大纲：" + entry.outline.join(" / "));
                if (entry.at) lines.push("登记：" + String(entry.at).slice(0, 10));
                return lines.filter((line) => line !== "").join("\n");
            }
            if (tab === "archive") return String(entry.text || "");
            if (tab === "kb") {
                // 文档级字段都摆出来：档位决定它能不能被当成依据（写入期质量门按它判），
                // 块数与文档 id 是会话侧 kb-read 要用的两个入口。
                const lines = [String(entry.path || "")];
                if (entry.title) lines.push("标题：" + entry.title);
                lines.push("知识块：" + String(entry.chunks || 0) + " 块 ｜ " + String(entry.chars || 0) + " 字 ｜ " + String(entry.bytes || 0) + " 字节");
                if (entry.tier) lines.push("来源档：" + entry.tier);
                if (entry.at) lines.push("入库：" + String(entry.at).slice(0, 19).replace("T", " "));
                lines.push("文档 id：" + String(entry.id || ""));
                return lines.join("\n");
            }
            if (tab === "links") {
                return String(entry.sourceId || "") + " —[" + String(entry.kind || "related") + "]→ " + String(entry.targetId || "")
                    + (entry.note ? "\n" + entry.note : "");
            }
            const refs = Array.isArray(entry.refs) ? entry.refs : [];
            return refs
                .map((ref) => (ref && ref.layer ? "[" + ref.layer + "] " : "") + String((ref && (ref.text || ref.id)) || ""))
                .join("\n");
        }

        /** 条目 key：与列表渲染用的是同一套（id → name → 序号）。 */
        function itemKeyOf(item, index) {
            const entry = item || {};
            return entry.id || entry.name || index;
        }

        /** 内容够长才需要「展开」：短条目折叠与展开没差别，不该多一个按钮。 */
        function isLongItem(item, tab) {
            const text = detailTextOf(item, tab);
            return text.length > 120 || text.split("\n").length > 6;
        }

        /**
         * 把分组后的条目裁到前 limit 条（跨组连续计数）。
         *
         * 分组标题上的条数用**整组的总数**（不是裁完剩下的）：否则「再显示 N 条」
         * 会让标题里的数字越点越小，读起来像条目被删了。
         */
        function paginateGroups(groups, limit) {
            let left = limit;
            const out = [];
            for (const group of groups) {
                if (left <= 0) break;
                const items = group.items.slice(0, left);
                left -= items.length;
                if (items.length > 0) out.push(Object.assign({}, group, { items: items }));
            }
            return out;
        }

        /** 把一条条目渲染成卡片。不同 Tab 的字段不一样，所以按 Tab 分支。 */
        function BrowserItem(props) {
            const item = props.item || {};
            const tab = props.tab;
            const key = itemKeyOf(item, props.index);
            const hover = props.hovered === true;
            const open = props.expanded === true;
            // 悬停只改描边、阴影与底色浓度，不动几何 —— 列表密度高，位移会抖。
            const card = hover ? Object.assign({}, S.browserItem, S.browserItemHover) : S.browserItem;
            const handlers = {
                onMouseEnter: () => { if (typeof props.onHover === "function") props.onHover(key); },
                onMouseLeave: () => { if (typeof props.onHover === "function") props.onHover(null); },
                // 键盘聚焦走同一套状态：底部详情条据此跟着走，纯键盘用户也能读到全文。
                onFocus: () => { if (typeof props.onHover === "function") props.onHover(key); },
                onBlur: () => { if (typeof props.onHover === "function") props.onHover(null); },
            };
            // 「展开」是**折叠**的逆操作，不是提示：正文默认限高 9em，点开就地铺满。
            // 原来这里靠原生 title 兜底，现在两个出口都不用系统气泡。
            const body = (text) => h("div", {
                style: open ? S.browserBodyOpen : S.browserBody,
                "data-item-body": String(key),
                "data-item-open": open ? "true" : undefined,
            }, String(text === undefined || text === null ? "" : text));

            const actions = [];
            if (isLongItem(item, tab)) {
                actions.push(h("button", {
                    key: "expand", type: "button", className: "om-act",
                    style: S.itemAction, "data-item-expand": String(key),
                    "aria-expanded": open,
                    onClick: () => { if (typeof props.onToggle === "function") props.onToggle(key); },
                }, open ? "收起" : "展开"));
            }
            for (const action of ITEM_ACTIONS[tab] || []) {
                actions.push(h("button", {
                    key: action.id, type: "button", className: "om-act",
                    style: S.itemAction, "data-memory-action": action.id,
                    "data-memory-action-item": String(key),
                    "aria-label": action.hint,
                    onClick: () => { if (typeof props.onAction === "function") props.onAction(action, item); },
                }, action.label));
            }
            const actionRow = actions.length > 0
                ? h("div", { style: S.itemActions, "data-item-actions": String(key) }, actions)
                : null;
            const shell = (style, children) => h("div", Object.assign({
                key: key, style: style, "data-item-key": String(key),
            }, handlers), children);

            if (tab === "hot") {
                // 重要度落在左边那条竖线上：critical / normal / low 一眼分得开，
                // 不必先在徽标里读一遍字。用 inset 阴影而不是 border-left ——
                // 前者不改盒模型、也不会被圆角裁掉。
                const importance = item.importance === "critical" || item.importance === "low" ? item.importance : "normal";
                const tone = IMPORTANCE_COLORS[importance] || IMPORTANCE_COLORS.normal;
                const style = Object.assign({}, card, {
                    boxShadow: "inset 3px 0 0 " + tone + ", " + (hover ? M.lift : M.soft),
                    background: tintedSurface(tone),
                });
                return h("div", Object.assign({
                    key: key, style: style, "data-importance": importance, "data-item-key": String(key),
                }, handlers),
                    h(Badges, { items: [originLabel(item.origin), item.importance || "normal", item.target === "user" ? "用户偏好" : "项目与环境"] }),
                    body(item.content),
                    actionRow,
                );
            }
            if (tab === "ledger") {
                return shell(card, [
                    h(Badges, { key: "badges", items: [originLabel(item.origin), item.format, item.theme, String(item.at || "").slice(0, 10)] }),
                    body(String(item.path || "") + (item.purpose ? " —— " + item.purpose : "")),
                    actionRow,
                ]);
            }
            if (tab === "archive") {
                return shell(card, [
                    h(Badges, { key: "badges", items: [originLabel(item.origin), item.kind, item.month, item.importance, item.category] }),
                    body(item.text),
                    actionRow,
                ]);
            }
            if (tab === "kb") {
                // 一条文档：档位与块数是「能不能拿它当依据」的两个关键字段
                // （单源未核实的档位会被写入期质量门拦下），所以放在徽标里。
                const size = humanBytes(item.bytes);
                return shell(card, [
                    h(Badges, { key: "badges", items: [
                        originLabel(item.origin), item.tier,
                        String(item.chunks || 0) + " 块", size, String(item.at || "").slice(0, 10),
                    ] }),
                    body(String(item.path || "") + (item.title ? " —— " + item.title : "")),
                    actionRow,
                ]);
            }
            if (tab === "links") {
                return shell(card, [
                    h(Badges, { key: "badges", items: [originLabel(item.origin), String(item.at || "").slice(0, 10)] }),
                    body(String(item.sourceId || "") + " —[" + String(item.kind || "related") + "]→ " + String(item.targetId || "")
                        + (item.note ? "（" + item.note + "）" : "")),
                    actionRow,
                ]);
            }
            // entities：一个实体名 + 出现次数 + 指向它的若干条目（展开后给全部）。
            const refs = Array.isArray(item.refs) ? item.refs : [];
            const shown = open ? refs : refs.slice(0, 5);
            return shell(card, [
                h(Badges, { key: "badges", items: [String(item.name || ""), "出现 " + String(item.count || 0) + " 次"] }),
                body(shown.map((ref) => String((ref && (ref.text || ref.id)) || "")).join("\n")),
                actionRow,
            ]);
        }

        // ── 记忆浏览面板：可视化 ──────────────────────────────────────────
        //
        // 参考 mnemon 的记忆图谱：手写 SVG + 自写的力导向布局，不引任何图形库。
        // 三条硬约束决定了这里的写法：
        //   1. bundle 在测试与 SSR 里**没有 document**（test/client.mjs 用
        //      new Function('window','require','console','fetch') 跑它），所以
        //      样式一律是内联对象，绝不在模块顶层注入 CSS；
        //   2. 只允许 require("react")，所以布局必须自己算；
        //   3. 这是浏览器面板里的同步渲染，布局必须有界 —— 节点数与迭代次数
        //      都封顶，并且按节点集签名缓存，重渲染不重算。

        /** 组合图上的画布与布局参数。 */
        const GRAPH_WIDTH = 720;
        const GRAPH_HEIGHT = 380;
        const GRAPH_MARGIN = 44;
        /**
         * 节点实际允许贴到多边。
         *
         * 比 GRAPH_MARGIN 再往里收：标签是**居中**画的，一个 14 字的中文标签
         * 向两侧各伸约 45px，正好等于原来的边距 —— 于是贴边的节点标签会被画布裁掉
         * （第十轮截图里 "季度回顾.docx" 就被切了一半）。测试按 GRAPH_MARGIN 核对
         * 节点不越界，往里收不违反它。
         */
        const GRAPH_LABEL_MARGIN = 78;
        /** 节点上限：再多就只画有关系的那部分 + 先来的条目，保证 O(n²) 的迭代不失控。 */
        const GRAPH_MAX_NODES = 48;
        /** 迭代次数：mnemon 是 180（930×520 的大画布），这里画布更小，90 次足够收敛。 */
        const GRAPH_ITERATIONS = 90;
        /** 视图缩放的上下限：低于 0.5 读不出标签，高于 3 只剩几个节点。 */
        const GRAPH_MIN_SCALE = 0.5;
        const GRAPH_MAX_SCALE = 3;
        /** 滚轮 / 按钮每一格的缩放倍率（1.15 ≈ 每 5 格翻一倍）。 */
        const GRAPH_ZOOM_STEP = 1.15;
        /** 方向键微调的步长（布局坐标系，不随缩放变）。 */
        const GRAPH_NUDGE = 12;

        /** 每层的形状与配色：圆 = 热记忆，圆角方 = 台账，方 = 归档，菱形 = 实体。 */
        const LAYER_STYLE = {
            hot: { shape: "circle", color: "#5a8cff", label: "热记忆" },
            ledger: { shape: "rect", color: "#22a879", label: "台账" },
            archive: { shape: "square", color: "#c08a2e", label: "归档" },
            entity: { shape: "diamond", color: "#a06bd6", label: "实体" },
        };
        const LAYER_FALLBACK = { shape: "circle", color: "#8a8f98", label: "条目" };

        /** 关系类型的配色。未知类型走兜底灰，不猜。 */
        const EDGE_COLORS = {
            related: "#708199",
            refines: "#2f9e78",
            supersedes: "#c08a2e",
            contradicts: "#e5484d",
            supports: "#5a8cff",
            derives: "#a06bd6",
            entity: "#a06bd6",
        };

        /** 重要度对应的左竖线颜色。 */
        const IMPORTANCE_COLORS = {
            critical: "var(--dsw-alias-state-error-primary, #e5484d)",
            normal: "var(--dsw-alias-state-business-primary, #5a8cff)",
            low: "var(--dsw-alias-border-l3, rgba(128,128,128,0.45))",
        };

        /** 热记忆按 target 分两组：这两组是「记给谁」的分界，不是重要度分界。 */
        const HOT_GROUPS = [
            { key: "user", label: "用户偏好" },
            { key: "project", label: "项目与环境" },
        ];

        /** 指标条与图例上的层顺序（固定，保证渲染确定性）。 */
        const LAYER_ORDER = ["hot", "ledger", "archive", "entity"];

        /** 指标条顺序：与 Tab 的对应关系固定。 */
        const METRIC_TILES = [
            { key: "hot", label: "热记忆", color: "#5a8cff", note: "常驻的偏好与约定" },
            { key: "ledger", label: "台账", color: "#22a879", note: "交付物自动登记" },
            { key: "archive", label: "归档", color: "#c08a2e", note: "下沉的旧条目" },
            { key: "kb", label: "知识库", color: "#3d9bbf", note: "入库的文档与知识块" },
            { key: "links", label: "关系", color: "#708199", note: "类型化的关系边" },
            { key: "entities", label: "实体", color: "#a06bd6", note: "被声明的实体" },
        ];

        /** 指标卡与页签共用一套配色：按 key 查 METRIC_TILES，查不到走兜底灰。 */
        function metricColor(key) {
            const found = METRIC_TILES.find((tile) => tile.key === key);
            return found === undefined ? "#8a8f98" : found.color;
        }

        /**
         * 快照里的计数 / 字节 / 条数一律当**可选**字段读。
         *
         * stores[].bytes / files 与 counts.hotBytes / projectBytes 是后加的，
         * 老版本快照没有它们。用 Number.isFinite 兜住，缺了就不画那一项 ——
         * 不能让面板出现 NaN 或 undefined。
         */
        function finite(value) {
            return Number.isFinite(value) ? value : null;
        }

        /** 计数文本：缺字段给一个破折号，不谎报 0。 */
        function countText(value) {
            const number = finite(value);
            return number === null ? "—" : String(number);
        }

        /** 人读的字节数。缺失或非法返回 null（而不是 "NaN B"）。 */
        function humanBytes(value) {
            const bytes = finite(value);
            if (bytes === null) return null;
            if (bytes < 1024) return bytes + " B";
            const units = ["KB", "MB", "GB", "TB"];
            let size = bytes / 1024;
            let index = 0;
            while (size >= 1024 && index < units.length - 1) {
                size = size / 1024;
                index += 1;
            }
            return (size >= 10 ? Math.round(size) : Math.round(size * 10) / 10) + " " + units[index];
        }

        /**
         * 容量条的百分比（保留一位小数）。
         *
         * 1 / 500 直接取整会变成 0%，看起来像「压根没有条」；保留一位小数才是
         * 0.2%。填充条另有 minWidth，所以极小的比例在界面上仍然看得见。
         * 上限缺失或为 0 时返回 null（这一项不画），而不是除零得 Infinity。
         */
        function percentOf(value, limit) {
            const used = finite(value);
            const total = finite(limit);
            if (used === null || total === null || total <= 0) return null;
            return Math.min(100, Math.round((used / total) * 1000) / 10);
        }

        /** 容量条的行数据：缺上限或上限为 0 的项直接不出现。 */
        function capacityRows(counts, limits) {
            const source = counts || {};
            const bound = limits || {};
            const rows = [];
            const push = (key, name, value, limit, unit, color) => {
                const percent = percentOf(value, limit);
                if (percent === null) return;
                rows.push({ key: key, name: name, value: finite(value), limit: finite(limit), unit: unit, percent: percent, color: color });
            };
            // 每条一个语义色：台账=主色、归档=警示、热记忆=强调、项目记忆=成功。
            // 不给阈值变色 —— mnemon 的容量条也是恒定色，轨道永远是浅底。
            //
            // 两个口径必须与「列表里有几条」分开（见 src/view.js 里 counts 的注释）：
            //   台账用 ledgerTotal —— 未过滤、未被浏览上限截断的总数；
            //   归档用 archiveFiles —— 上限 archiveKeep 管的是**摘要文件数**（一个月
            //   一个），拿条目数配它会显示成「30 / 60 个」这种假满。
            push("ledger", "台账条数", source.ledgerTotal, bound.ledgerLimit, "条", M.business);
            push("archive", "归档摘要文件", source.archiveFiles, bound.archiveKeep, "个", M.warn);
            push("hotBytes", "热记忆占用", source.hotBytes, bound.userLimitBytes, "字节", M.accent);
            push("projectBytes", "项目记忆占用", source.projectBytes, bound.projectLimitBytes, "字节", M.success);
            return rows;
        }

        /** 按某个键把条目分桶；键为空的条目直接丢掉。返回按键升序的桶。 */
        function bucketize(items, keyOf) {
            const order = [];
            const counts = new Map();
            for (const item of items) {
                const key = keyOf(item);
                if (typeof key !== "string" || key === "") continue;
                if (!counts.has(key)) order.push(key);
                counts.set(key, (counts.get(key) || 0) + 1);
            }
            order.sort();
            return order.map((key) => ({ key: key, label: key, count: counts.get(key) }));
        }

        /** 台账按天（at 的前 10 位，即 YYYY-MM-DD）。 */
        function ledgerDay(item) {
            const at = item && typeof item.at === "string" ? item.at.slice(0, 10) : "";
            return /^\d{4}-\d{2}-\d{2}$/.test(at) ? at : "";
        }

        /** 归档按月：优先 month 字段，其次从 at 推。 */
        function archiveMonth(item) {
            const month = item && typeof item.month === "string" ? item.month : "";
            if (/^\d{4}-\d{2}$/.test(month)) return month;
            const at = item && typeof item.at === "string" ? item.at.slice(0, 7) : "";
            return /^\d{4}-\d{2}$/.test(at) ? at : "";
        }

        /** 热记忆分组：按 target 分成两组，认不出来的进「其它」。 */
        function groupHot(items) {
            const buckets = [];
            for (const group of HOT_GROUPS) buckets.push({ key: group.key, label: group.label, items: [] });
            const other = { key: "other", label: "其它", items: [] };
            for (const item of items) {
                const target = item && typeof item.target === "string" ? item.target : "";
                let placed = false;
                for (const bucket of buckets) {
                    if (bucket.key === target) { bucket.items.push(item); placed = true; break; }
                }
                if (!placed) other.items.push(item);
            }
            const visible = [];
            for (const bucket of buckets) if (bucket.items.length > 0) visible.push(bucket);
            if (other.items.length > 0) visible.push(other);
            return visible;
        }

        /** 层的显示名与颜色；未知层走兜底。 */
        function layerStyle(layer) {
            const key = String(layer || "");
            if (key === "hot" || key === "ledger" || key === "archive" || key === "entity") return LAYER_STYLE[key];
            return LAYER_FALLBACK;
        }

        /** 关系类型配色；未知类型走兜底灰（hasOwnProperty：不认原型链上的键）。 */
        function edgeColor(kind) {
            const key = String(kind || "");
            if (Object.prototype.hasOwnProperty.call(EDGE_COLORS, key)) return EDGE_COLORS[key];
            return "#8a8f98";
        }

        /** 图上的节点标签：压掉换行、限长，空值给兜底名。 */
        function graphLabel(value, fallback, max) {
            const text = String(value === undefined || value === null ? "" : value).replace(/\s+/g, " ").trim();
            if (text === "") return fallback;
            return text.length > max ? text.slice(0, max - 1) + "…" : text;
        }

        /** FNV-1a：布局的种子。同一个节点集必须落回同一张图，不能用随机数。 */
        function fnvHash(text) {
            let result = 2166136261;
            const value = String(text);
            for (let index = 0; index < value.length; index += 1) {
                result = Math.imul(result ^ value.charCodeAt(index), 16777619);
            }
            return result >>> 0;
        }

        function clampGraphPoint(x, y) {
            return {
                x: Math.min(GRAPH_WIDTH - GRAPH_LABEL_MARGIN, Math.max(GRAPH_LABEL_MARGIN, x)),
                y: Math.min(GRAPH_HEIGHT - GRAPH_LABEL_MARGIN + 20, Math.max(GRAPH_MARGIN, y)),
            };
        }

        function round1(value) {
            return Math.round(value * 10) / 10;
        }

        /**
         * 三位小数。
         *
         * 视图变换用它而不是 round1：缩放步长是 1.15，四舍五入到一位小数会变成
         * 1.2 —— 每按一次「+」实际放大的倍率和读数（115%）就对不上了，而且
         * 倍率会被这个取整吞掉（1.15 → 1.2 → 1.4…）。
         */
        function round3(value) {
            return Math.round(value * 1000) / 1000;
        }

        /**
         * 把快照摊成一张图。
         *
         * 节点 = 热记忆 + 台账 + 归档 + 实体；边 = links[]（按 kind 上色）与
         * entities[].refs[]（实体 → 条目）。links 里存的是裸 id（不带层），
         * 所以先建一张 rawId → 节点 key 的索引再解析；解析不出来的边直接丢，
         * 不画悬空的线。
         */
        function buildGraphModel(data) {
            const source = data || {};
            const hot = Array.isArray(source.hot) ? source.hot : [];
            const ledger = Array.isArray(source.ledger) ? source.ledger : [];
            const archive = Array.isArray(source.archive) ? source.archive : [];
            const links = Array.isArray(source.links) ? source.links : [];
            const entities = Array.isArray(source.entities) ? source.entities : [];

            const nodes = [];
            const nodeKeys = new Set();
            const byRawId = new Map();
            const add = (key, rawId, layer, label, extra) => {
                if (nodeKeys.has(key)) return;
                nodeKeys.add(key);
                nodes.push(Object.assign({ key: key, layer: layer, label: label }, extra || {}));
                if (typeof rawId === "string" && rawId !== "") {
                    const list = byRawId.get(rawId);
                    if (list === undefined) byRawId.set(rawId, [key]);
                    else list.push(key);
                }
            };

            hot.forEach((item, index) => {
                const entry = item || {};
                add("hot:" + (entry.id ? entry.id : index), entry.id, "hot",
                    graphLabel(entry.content, "热记忆 " + (index + 1), 14), { importance: entry.importance });
            });
            ledger.forEach((item, index) => {
                const entry = item || {};
                add("ledger:" + (entry.id ? entry.id : index), entry.id, "ledger",
                    graphLabel(entry.path, "台账 " + (index + 1), 14), {});
            });
            archive.forEach((item, index) => {
                const entry = item || {};
                add("archive:" + (entry.id ? entry.id : index), entry.id, "archive",
                    graphLabel(entry.text, "归档 " + (index + 1), 14), {});
            });
            entities.forEach((item, index) => {
                const entry = item || {};
                add("entity:" + index + ":" + String(entry.name || ""), "", "entity",
                    graphLabel(entry.name, "实体 " + (index + 1), 14), { count: finite(entry.count) });
            });

            const resolve = (id) => {
                const key = String(id === undefined || id === null ? "" : id);
                if (key === "") return null;
                const list = byRawId.get(key);
                if (list === undefined || list.length === 0) return null;
                return list[0];
            };

            const edges = [];
            const edgeKeys = new Set();
            const addEdge = (from, to, kind) => {
                if (from === null || to === null || from === to) return;
                const key = from + ">" + to + ":" + kind;
                if (edgeKeys.has(key)) return;
                edgeKeys.add(key);
                edges.push({ key: key, source: from, target: to, kind: kind, color: edgeColor(kind) });
            };

            links.forEach((item) => {
                if (!item) return;
                const kind = typeof item.kind === "string" && item.kind !== "" ? item.kind : "related";
                addEdge(resolve(item.sourceId), resolve(item.targetId), kind);
            });
            entities.forEach((item, index) => {
                const entry = item || {};
                if (!Array.isArray(entry.refs)) return;
                const entityKey = "entity:" + index + ":" + String(entry.name || "");
                entry.refs.forEach((ref) => {
                    if (!ref) return;
                    const layer = typeof ref.layer === "string" && ref.layer !== "" ? ref.layer : "hot";
                    const direct = layer + ":" + String(ref.id === undefined || ref.id === null ? "" : ref.id);
                    addEdge(entityKey, nodeKeys.has(direct) ? direct : resolve(ref.id), "entity");
                });
            });

            // 节点封顶：先保有关系的那批（否则图会退化成一堆孤点），再按原顺序补齐。
            let kept = nodes;
            if (nodes.length > GRAPH_MAX_NODES) {
                const connected = new Set();
                for (const edge of edges) { connected.add(edge.source); connected.add(edge.target); }
                const head = [];
                const tail = [];
                for (const node of nodes) {
                    if (connected.has(node.key)) head.push(node); else tail.push(node);
                }
                kept = head.concat(tail).slice(0, GRAPH_MAX_NODES);
            }
            const keptKeys = new Set();
            for (const node of kept) keptKeys.add(node.key);
            const visibleEdges = edges.filter((edge) => keptKeys.has(edge.source) && keptKeys.has(edge.target));
            return { nodes: kept, edges: visibleEdges };
        }

        /**
         * 力导向布局：斥力 18e3/d²、按边类型定弹簧长度、逐轮冷却。
         *
         * 迭代次数与节点数都是有界的（48 × 90），并且布局对同一个节点集是纯函数 ——
         * 所以结果能安全地按签名缓存，重渲染（悬停 / 选中 / 键盘微调）不会重算。
         */
        function computeGraphLayout(nodes, edges) {
            const positions = new Map();
            const velocities = new Map();
            const centerX = GRAPH_WIDTH / 2;
            const centerY = GRAPH_HEIGHT / 2;

            // 种子化初值：黄金角螺旋 + FNV 抖动。同一个 key 每次都落同一个位置，
            // 快照刷新（内容不变）时图不会整体跳一下。
            nodes.forEach((node, index) => {
                const seed = fnvHash(node.key);
                const angle = index * 2.399963 + ((seed % 37) / 37) * 0.4;
                const radius = nodes.length === 1 ? 0 : 34 + Math.sqrt(index + 1) * 32;
                positions.set(node.key, clampGraphPoint(centerX + Math.cos(angle) * radius, centerY + Math.sin(angle) * radius));
                velocities.set(node.key, { x: 0, y: 0 });
            });

            const sparseScale = nodes.length <= 3 ? 2 : nodes.length <= 8 ? 1.45 : 1;
            for (let iteration = 0; iteration < GRAPH_ITERATIONS; iteration += 1) {
                const cooling = 1 - iteration / (GRAPH_ITERATIONS * 1.2);

                for (let leftIndex = 0; leftIndex < nodes.length; leftIndex += 1) {
                    const leftPosition = positions.get(nodes[leftIndex].key);
                    const leftVelocity = velocities.get(nodes[leftIndex].key);
                    for (let rightIndex = leftIndex + 1; rightIndex < nodes.length; rightIndex += 1) {
                        const rightPosition = positions.get(nodes[rightIndex].key);
                        const rightVelocity = velocities.get(nodes[rightIndex].key);
                        let dx = leftPosition.x - rightPosition.x;
                        let dy = leftPosition.y - rightPosition.y;
                        if (dx === 0 && dy === 0) {
                            dx = (fnvHash(nodes[leftIndex].key) % 13) - 6 || 1;
                            dy = (fnvHash(nodes[rightIndex].key) % 11) - 5 || -1;
                        }
                        const distanceSquared = Math.max(100, dx * dx + dy * dy);
                        const distance = Math.sqrt(distanceSquared);
                        const force = Math.min(9, 18e3 / distanceSquared) * cooling
                            + (distance < 66 ? (66 - distance) * 0.08 : 0);
                        const forceX = (dx / distance) * force;
                        const forceY = (dy / distance) * force;
                        leftVelocity.x += forceX;
                        leftVelocity.y += forceY;
                        rightVelocity.x -= forceX;
                        rightVelocity.y -= forceY;
                    }
                }

                for (const edge of edges) {
                    const from = positions.get(edge.source);
                    const to = positions.get(edge.target);
                    if (from === undefined || to === undefined) continue;
                    const fromVelocity = velocities.get(edge.source);
                    const toVelocity = velocities.get(edge.target);
                    const dx = to.x - from.x;
                    const dy = to.y - from.y;
                    const distance = Math.max(1, Math.sqrt(dx * dx + dy * dy));
                    // 实体 → 条目这条边要短一些：实体是条目的附属，不该飘到对岸去。
                    const rest = (edge.kind === "entity" ? 86 : 112) * sparseScale;
                    const spring = (distance - rest) * 0.018 * cooling;
                    const forceX = (dx / distance) * spring;
                    const forceY = (dy / distance) * spring;
                    fromVelocity.x += forceX;
                    fromVelocity.y += forceY;
                    toVelocity.x -= forceX;
                    toVelocity.y -= forceY;
                }

                for (const node of nodes) {
                    const position = positions.get(node.key);
                    const velocity = velocities.get(node.key);
                    // 向心力：让整张图收在画布中间，否则孤立节点会贴在边上。
                    velocity.x = Math.max(-12, Math.min(12, (velocity.x + (centerX - position.x) * 0.0016) * 0.76));
                    velocity.y = Math.max(-12, Math.min(12, (velocity.y + (centerY - position.y) * 0.0016) * 0.76));
                    positions.set(node.key, clampGraphPoint(position.x + velocity.x, position.y + velocity.y));
                }
            }
            return positions;
        }

        /**
         * 「均匀重置」布局：把节点按层序排到两圈同心圆上。
         *
         * 力导向布局每次微调都会连带挪动邻居，读「谁和谁有关系」时容易越看越乱；
         * mnemon 因此给了一个「均匀重置」的动作，把节点摊成一张整齐的图。
         * 这里同样只做几何，不做任何语义判断：内圈 / 外圈交错摆，避免相邻标签叠住。
         */
        function computeUniformLayout(nodes) {
            const positions = new Map();
            const count = nodes.length;
            const centerX = GRAPH_WIDTH / 2;
            const centerY = GRAPH_HEIGHT / 2;
            if (count === 0) return positions;
            if (count === 1) {
                positions.set(nodes[0].key, clampGraphPoint(centerX, centerY));
                return positions;
            }
            const ring = Math.min(GRAPH_WIDTH, GRAPH_HEIGHT) / 2 - GRAPH_LABEL_MARGIN;
            const ordered = nodes.slice().sort((left, right) => {
                const diff = LAYER_ORDER.indexOf(left.layer) - LAYER_ORDER.indexOf(right.layer);
                if (diff !== 0) return diff;
                return String(left.key) < String(right.key) ? -1 : String(left.key) > String(right.key) ? 1 : 0;
            });
            ordered.forEach((node, index) => {
                const angle = (index / count) * Math.PI * 2 - Math.PI / 2;
                const radius = ring * (index % 2 === 0 ? 1 : 0.62);
                positions.set(node.key, clampGraphPoint(
                    centerX + Math.cos(angle) * radius,
                    centerY + Math.sin(angle) * radius,
                ));
            });
            return positions;
        }

        /** 布局缓存：按「节点集 + 边集 + 布局模式」签名命中。上限 8 张，超出整片清掉。 */
        const GRAPH_LAYOUT_CACHE = new Map();

        function graphLayout(nodes, edges, mode) {
            const parts = [];
            for (const node of nodes) parts.push(node.key);
            parts.push("|");
            for (const edge of edges) parts.push(edge.source + ">" + edge.target + ":" + edge.kind);
            // 模式进签名：换布局不能命中上一张缓存，否则「均匀重置」按下去没反应。
            parts.push("|" + (mode === "uniform" ? "uniform" : "force"));
            const signature = parts.join(",");
            const cached = GRAPH_LAYOUT_CACHE.get(signature);
            if (cached !== undefined) return cached;
            const computed = mode === "uniform" ? computeUniformLayout(nodes) : computeGraphLayout(nodes, edges);
            if (GRAPH_LAYOUT_CACHE.size >= 8) GRAPH_LAYOUT_CACHE.clear();
            GRAPH_LAYOUT_CACHE.set(signature, computed);
            return computed;
        }

        /** 节点半径：实体按出现次数略放大，其余等大。 */
        function graphRadius(node) {
            if (node.layer === "entity") {
                const count = finite(node.count);
                return 10 + Math.min(6, count === null ? 0 : Math.sqrt(Math.max(1, count)) * 2);
            }
            return 9;
        }

        /** 按层形状造一个图形元素（圆 / 圆角方 / 方 / 菱形）。 */
        function graphShape(shape, radius, props) {
            if (shape === "diamond") {
                const reach = radius * 1.25;
                return h("polygon", Object.assign({
                    points: "0," + (-reach) + " " + reach + ",0 0," + reach + " " + (-reach) + ",0",
                }, props));
            }
            if (shape === "circle") return h("circle", Object.assign({ r: radius }, props));
            return h("rect", Object.assign({
                x: -radius, y: -radius, width: radius * 2, height: radius * 2,
                rx: shape === "square" ? 1.5 : 5,
            }, props));
        }

        /**
         * 关系图本体：SVG + 自写布局 + 视图变换。
         *
         * 交互分两层，刻意分开：
         *   - **视图层**（缩放 / 平移 / 拖节点）：缩放与平移只改一个父级 `<g>` 的
         *     transform，节点的 translate 仍然是布局坐标 —— 于是拖视图不会触发
         *     90 轮迭代重算，方向键微调的 12px 也仍然是布局单位。滚轮以**光标**为
         *     锚点缩放（mnemon 的画布手感），背景 `cursor: grab` 直接拖。
         *   - **布局层**（自然铺开 / 均匀重置）：换一张坐标表，视图变换不动。
         *   - 悬停高亮邻居、点击选中、Enter / 空格切换、方向键微调都不变。
         *
         * 滚轮为什么要自己挂原生监听：React 的 `onWheel` 挂在 root 上而且是**被动**
         * 监听，在它里面 `preventDefault()` 会被浏览器丢掉、并留下一条 console 警告
         * （活体探针要求零 error / warning）。所以 React 那一侧只做缩放，元素上另挂
         * 一个 `{passive:false}` 的监听专门 `preventDefault()`，两个监听各做一件事，
         * 不会缩放两次。
         */
        function MemoryGraph(props) {
            const nodes = Array.isArray(props.nodes) ? props.nodes : [];
            const edges = Array.isArray(props.edges) ? props.edges : [];
            const [hovered, setHovered] = React.useState(null);
            const [selected, setSelected] = React.useState(null);
            const [offsets, setOffsets] = React.useState({});
            const [layout, setLayout] = React.useState("force");
            const [view, setView] = React.useState({ k: 1, tx: 0, ty: 0 });
            const [dragging, setDragging] = React.useState(null);
            /**
             * 拖拽事实放 ref 而不是 state：pointermove 是高频事件，闭包里的 state
             * 会落后一帧，累积起来就是「拖了 100px 只挪了 80px」那种漂移。
             */
            const dragRef = React.useRef(null);
            /** 这一轮 pointerdown → click 之间有没有真的移动过（移动过就不算点击）。 */
            const movedRef = React.useRef(false);
            const surfaceRef = React.useRef(null);

            const base = graphLayout(nodes, edges, layout);

            const neighbours = new Map();
            for (const node of nodes) neighbours.set(node.key, new Set());
            for (const edge of edges) {
                const from = neighbours.get(edge.source);
                const to = neighbours.get(edge.target);
                if (from !== undefined) from.add(edge.target);
                if (to !== undefined) to.add(edge.source);
            }
            const focus = selected !== null ? selected : hovered;
            const focusSet = focus !== null && neighbours.has(focus) ? neighbours.get(focus) : null;

            const points = new Map();
            for (const node of nodes) {
                const point = base.get(node.key) || { x: GRAPH_WIDTH / 2, y: GRAPH_HEIGHT / 2 };
                const offset = offsets[node.key] || { dx: 0, dy: 0 };
                points.set(node.key, clampGraphPoint(point.x + offset.dx, point.y + offset.dy));
            }

            const offsetBy = (key, dx, dy) => {
                setOffsets((current) => {
                    const next = Object.assign({}, current);
                    const previous = next[key] || { dx: 0, dy: 0 };
                    next[key] = { dx: previous.dx + dx, dy: previous.dy + dy };
                    return next;
                });
            };
            const toggleSelect = (key) => {
                setSelected((current) => (current === key ? null : key));
            };

            /** 以 anchor（视图坐标系里的点）为锚点缩放：光标底下那个点在缩放前后不动。 */
            const zoomAt = (factor, anchor) => {
                setView((current) => {
                    const next = clampNumber(current.k * factor, GRAPH_MIN_SCALE, GRAPH_MAX_SCALE);
                    if (next === current.k) return current;
                    const ratio = next / current.k;
                    const ax = anchor && Number.isFinite(anchor.x) ? anchor.x : GRAPH_WIDTH / 2;
                    const ay = anchor && Number.isFinite(anchor.y) ? anchor.y : GRAPH_HEIGHT / 2;
                    return {
                        k: next,
                        tx: ax - (ax - current.tx) * ratio,
                        ty: ay - (ay - current.ty) * ratio,
                    };
                });
            };
            /**
             * 事件坐标 → 视图坐标。
             *
             * 量不到 getBoundingClientRect（测试 / SSR）时退回画布中心：缩放仍然可用，
             * 只是锚点从光标变成中心 —— 不因为量不到就把整条路降级掉。
             */
            const viewAnchorOf = (event) => {
                const target = event && event.currentTarget ? event.currentTarget : null;
                const rect = target && typeof target.getBoundingClientRect === "function" ? target.getBoundingClientRect() : null;
                if (rect && rect.width > 0 && rect.height > 0) {
                    return {
                        x: (Number(event.clientX) - rect.left) * (GRAPH_WIDTH / rect.width),
                        y: (Number(event.clientY) - rect.top) * (GRAPH_HEIGHT / rect.height),
                    };
                }
                return { x: GRAPH_WIDTH / 2, y: GRAPH_HEIGHT / 2 };
            };
            /**
             * 滚轮缩放：**只在按住修饰键时**接管（Ctrl / ⌘ / Alt）。
             *
             * 第二十轮改的这条：面板改成了整页一个滚动容器以后，画布占了列表上面的
             * 一大块，原先「滚轮一律缩放」等于把这一大块变成滚动死区 —— 鼠标停在
             * 图上就滚不动页面（用户报的「记忆面板滚动有问题」有一半来自这里）。
             * 缩放本身没丢：按住 Ctrl/⌘ 滚轮，或者用视图条上的 ± 按钮。
             */
            const zoomModifier = (event) => Boolean(event && (event.ctrlKey || event.metaKey || event.altKey));
            const onWheel = (event) => {
                if (!zoomModifier(event)) return;
                const delta = Number(event && event.deltaY);
                if (!Number.isFinite(delta) || delta === 0) return;
                zoomAt(delta < 0 ? GRAPH_ZOOM_STEP : 1 / GRAPH_ZOOM_STEP, viewAnchorOf(event));
            };

            const startDrag = (event, kind, key) => {
                if (event && typeof event.button === "number" && event.button !== 0) return;
                // 节点在平移面之上：不掐断冒泡的话，拖节点会同时拖动整张画布。
                if (kind === "node" && event && typeof event.stopPropagation === "function") event.stopPropagation();
                movedRef.current = false;
                dragRef.current = {
                    kind: kind,
                    key: key === undefined ? null : key,
                    x: Number(event && event.clientX) || 0,
                    y: Number(event && event.clientY) || 0,
                };
                setDragging(kind);
                const target = event && event.currentTarget ? event.currentTarget : null;
                if (target && typeof target.setPointerCapture === "function" && event && event.pointerId !== undefined) {
                    try { target.setPointerCapture(event.pointerId); } catch (error) { /* 不支持就算了 */ }
                }
            };
            const moveDrag = (event) => {
                const drag = dragRef.current;
                if (drag === null) return;
                const x = Number(event && event.clientX) || 0;
                const y = Number(event && event.clientY) || 0;
                const dx = x - drag.x;
                const dy = y - drag.y;
                if (dx === 0 && dy === 0) return;
                dragRef.current = Object.assign({}, drag, { x: x, y: y });
                if (Math.abs(dx) + Math.abs(dy) >= 1) movedRef.current = true;
                if (drag.kind === "node") {
                    // 屏幕位移 → 布局位移：画布放大两倍时，手指走 10px 只该挪 5 单位。
                    const scale = view.k > 0 ? view.k : 1;
                    offsetBy(drag.key, dx / scale, dy / scale);
                    return;
                }
                setView((current) => ({ k: current.k, tx: current.tx + dx, ty: current.ty + dy }));
            };
            const endDrag = (event) => {
                if (dragRef.current === null) return;
                dragRef.current = null;
                setDragging(null);
                const target = event && event.currentTarget ? event.currentTarget : null;
                if (target && typeof target.releasePointerCapture === "function" && event && event.pointerId !== undefined) {
                    try { target.releasePointerCapture(event.pointerId); } catch (error) { /* 不支持就算了 */ }
                }
            };
            /** 布局动作：换坐标表，并清掉手动偏移（否则新布局会带着旧位移）。 */
            const applyLayout = (mode) => {
                setLayout(mode);
                setOffsets({});
            };
            const resetView = () => setView({ k: 1, tx: 0, ty: 0 });
            const onNodeClick = (key) => {
                // 拖过就不算点击：否则「把节点拖到别处」会顺手把它选中 / 取消选中。
                if (movedRef.current) { movedRef.current = false; return; }
                toggleSelect(key);
            };
            const onNodeKeyDown = (event, key) => {
                const pressed = event && typeof event.key === "string" ? event.key : "";
                if (pressed === "Enter" || pressed === " " || pressed === "Spacebar") {
                    if (event && typeof event.preventDefault === "function") event.preventDefault();
                    toggleSelect(key);
                    return;
                }
                const moves = {
                    ArrowLeft: [-GRAPH_NUDGE, 0], ArrowRight: [GRAPH_NUDGE, 0],
                    ArrowUp: [0, -GRAPH_NUDGE], ArrowDown: [0, GRAPH_NUDGE],
                };
                const move = Object.prototype.hasOwnProperty.call(moves, pressed) ? moves[pressed] : null;
                if (move === null) return;
                if (event && typeof event.preventDefault === "function") event.preventDefault();
                offsetBy(key, move[0], move[1]);
            };

            // 元素上另挂一个**非被动**滚轮监听：只负责「按住修饰键时 preventDefault」
            // （缩放时别让页面跟着滚）。不按修饰键的滚轮交给浏览器 —— 那是面板滚动，
            // 不该被画布吃掉。缩放仍由 React 的 onWheel 做，两个监听各做一件事。
            React.useEffect(() => {
                const surface = surfaceRef.current;
                if (!surface || typeof surface.addEventListener !== "function") return undefined;
                const block = (event) => {
                    if (!event || (!event.ctrlKey && !event.metaKey && !event.altKey)) return;
                    if (typeof event.preventDefault === "function") event.preventDefault();
                };
                surface.addEventListener("wheel", block, { passive: false });
                return () => {
                    if (typeof surface.removeEventListener === "function") surface.removeEventListener("wheel", block);
                };
            }, []);

            const dense = nodes.length > 24;
            const kinds = [];
            for (const edge of edges) if (kinds.indexOf(edge.kind) === -1) kinds.push(edge.kind);

            const svg = h("svg", {
                viewBox: "0 0 " + GRAPH_WIDTH + " " + GRAPH_HEIGHT,
                role: "img",
                "aria-label": "记忆关系图：共 " + nodes.length + " 个节点、" + edges.length + " 条关系。按住 Ctrl 或 ⌘ 滚轮缩放（也可用下面的加减按钮），拖动背景平移，拖动节点调整位置。",
                "data-graph": "memory",
                "data-layout": layout,
                "data-zoom": String(Math.round(view.k * 100)),
                "data-pan": round1(view.tx) + "," + round1(view.ty),
                "data-dragging": dragging === null ? undefined : dragging,
                "data-density": dense ? "dense" : "normal",
                "data-nodes": nodes.length,
                "data-edges": edges.length,
                ref: surfaceRef,
                onWheel: onWheel,
                // 拖拽的 move / up 挂在**画布**上，不挂在被拖的那个元素上：
                // 指针捕获之后事件仍会冒泡到画布，挂一处就能同时服务「拖背景」与
                // 「拖节点」两种拖拽，也不会因为两个元素各挂一份而把位移算两遍。
                onPointerMove: moveDrag,
                onPointerUp: endDrag,
                onPointerLeave: endDrag,
                style: S.graphCanvas,
            },
                h("defs", null,
                    // 网格线按 mnemon 的口径：stroke-width .6 + opacity .5 —— 细到几乎只是
                    // 一层纹理，不该和节点抢注意力。
                    h("pattern", {
                        id: "office-memory-grid", width: 28, height: 28, patternUnits: "userSpaceOnUse",
                    }, h("path", {
                        d: "M28 0H0V28", fill: "none",
                        stroke: "var(--dsw-alias-border-l1, rgba(128,128,128,0.2))",
                        strokeWidth: 0.6, opacity: 0.5,
                    })),
                ),
                // 平移面：整块画布大小的透明矩形，垫在所有内容**下面**（节点在上层，
                // 所以拖节点不会变成拖背景）。它必须在视图组外面 —— 放进组里的话
                // 平移之后它自己也跟着挪，边角就会露出拖不动的死区。
                h("rect", {
                    key: "pan", x: 0, y: 0, width: GRAPH_WIDTH, height: GRAPH_HEIGHT,
                    fill: "transparent", pointerEvents: "all",
                    "data-graph-pan": "true",
                    onPointerDown: (event) => startDrag(event, "pan", null),
                    style: { cursor: dragging === "pan" ? "grabbing" : "grab", touchAction: "none" },
                }),
                // 视图变换只落在这个组上：节点的 translate 仍是布局坐标，所以
                // 缩放 / 平移不重算布局，单测里读到的节点坐标也不受视图影响。
                h("g", {
                    key: "view",
                    "data-graph-view": "true",
                    transform: "translate(" + round3(view.tx) + " " + round3(view.ty) + ") scale(" + round3(view.k) + ")",
                },
                    h("rect", {
                        key: "grid", x: 0, y: 0, width: GRAPH_WIDTH, height: GRAPH_HEIGHT,
                        fill: "url(#office-memory-grid)", pointerEvents: "none", "data-graph-grid": "true",
                    }),
                    edges.map((edge) => {
                        const from = points.get(edge.source);
                        const to = points.get(edge.target);
                        if (from === undefined || to === undefined) return null;
                        const dx = to.x - from.x;
                        const dy = to.y - from.y;
                        const distance = Math.max(1, Math.sqrt(dx * dx + dy * dy));
                        const bend = Math.min(26, distance * 0.18);
                        const controlX = (from.x + to.x) / 2 - (dy / distance) * bend;
                        const controlY = (from.y + to.y) / 2 + (dx / distance) * bend;
                        const related = focus !== null && (edge.source === focus || edge.target === focus);
                        const dim = focus !== null && !related;
                        // 层次靠透明度而不是靠加粗（mnemon 的边就是这么分的）：
                        // 底噪 .4，实体边最实，聚焦到某个节点时只把相关的几条提亮。
                        const base = edge.kind === "entity" ? 0.72 : 0.4;
                        return h("path", {
                            key: edge.key,
                            className: "om-edge",
                            d: "M " + round1(from.x) + " " + round1(from.y)
                                + " Q " + round1(controlX) + " " + round1(controlY)
                                + " " + round1(to.x) + " " + round1(to.y),
                            fill: "none",
                            stroke: edge.color,
                            strokeWidth: related ? 2.2 : (edge.kind === "entity" ? 1.45 : 1.1),
                            opacity: dim ? 0.12 : (related ? 0.92 : base),
                            vectorEffect: "non-scaling-stroke",
                            "data-edge": edge.kind,
                            "data-source": edge.source,
                            "data-target": edge.target,
                        });
                    }),
                    nodes.map((node, index) => {
                        const point = points.get(node.key);
                        const palette = layerStyle(node.layer);
                        const radius = graphRadius(node);
                        const isFocus = focus === node.key;
                        const isNeighbour = focusSet !== null && focusSet.has(node.key);
                        const isSelected = selected === node.key;
                        const dim = focus !== null && !isFocus && !isNeighbour;
                        // 高密度时每三个节点隐一个标签：mnemon 用的是同一招。
                        const showLabel = !dense || index % 3 === 0 || isFocus;
                        // 贴边的节点把标签往回收：居中的标签会被画布裁掉（见 GRAPH_LABEL_MARGIN）。
                        const anchor = point.x > GRAPH_WIDTH - GRAPH_LABEL_MARGIN - 10 ? "end"
                            : point.x < GRAPH_LABEL_MARGIN + 10 ? "start" : "middle";
                        const parts = [
                            // 三层结构（halo / core / 焦点环）：halo 的浓度跟着状态走，
                            // 几何完全不变 —— 悬停时不抖。
                            graphShape(palette.shape, radius * (isFocus ? 2.2 : 1.75), {
                                key: "halo", className: "om-halo", fill: palette.color,
                                opacity: isFocus ? 0.34 : isNeighbour ? 0.22 : 0.16,
                                "data-shape": "halo",
                            }),
                            graphShape(palette.shape, radius * (isFocus ? 1.25 : 1), {
                                key: "core", fill: palette.color,
                                stroke: M.surface, strokeWidth: 1.5,
                                "data-shape": "core",
                            }),
                        ];
                        if (isSelected) {
                            parts.push(graphShape(palette.shape, radius * 2.7, {
                                key: "ring", fill: "none", stroke: palette.color, strokeWidth: 1.2,
                                strokeDasharray: "3 3", opacity: 0.9, "data-shape": "ring",
                            }));
                        }
                        if (showLabel) {
                            parts.push(h("text", {
                                key: "label", className: "om-label", x: 0, y: round1(radius + 14),
                                textAnchor: anchor, fontSize: dense ? 10 : 11,
                                fill: isFocus ? M.text : M.textSecondary,
                                fontWeight: isFocus ? 600 : 400,
                                // 描一圈表面色：标签压在边和网格上仍然读得清（mnemon 的
                                // nodeLabel 也是这个思路）。
                                stroke: M.surface, strokeWidth: 3, strokeLinejoin: "round",
                                paintOrder: "stroke",
                                opacity: dim ? 0.24 : 0.95,
                                "data-label": node.key,
                            }, node.label));
                        }
                        const held = dragRef.current !== null && dragRef.current.kind === "node" && dragRef.current.key === node.key;
                        return h("g", {
                            key: node.key,
                            className: "om-node",
                            transform: "translate(" + round1(point.x) + " " + round1(point.y) + ")",
                            "data-node": node.key,
                            "data-layer": node.layer,
                            "data-focused": isFocus ? "true" : undefined,
                            "data-selected": isSelected ? "true" : undefined,
                            "data-node-dragging": held && dragging === "node" ? "true" : undefined,
                            tabIndex: 0,
                            role: "button",
                            "aria-label": node.label + "（" + palette.label + "）",
                            onMouseEnter: () => setHovered(node.key),
                            onMouseLeave: () => setHovered((current) => (current === node.key ? null : current)),
                            onFocus: () => setHovered(node.key),
                            onBlur: () => setHovered((current) => (current === node.key ? null : current)),
                            onPointerDown: (event) => startDrag(event, "node", node.key),
                            onClick: () => onNodeClick(node.key),
                            onKeyDown: (event) => onNodeKeyDown(event, node.key),
                            style: { cursor: "pointer", outline: "none", opacity: dim ? 0.26 : 1 },
                        }, parts);
                    }),
                ),
            );

            const legend = h("div", { style: S.graphToolbar, "data-graph-legend": "true" },
                LAYER_ORDER.map((layer) => h("span", { key: layer, className: "om-legend", style: S.legendItem },
                    h("span", {
                        style: {
                            display: "inline-block", width: "9px", height: "9px", background: layerStyle(layer).color,
                            borderRadius: layer === "entity" ? "2px" : "999px",
                            transform: layer === "entity" ? "rotate(45deg)" : "none",
                            boxShadow: "0 0 0 3px color-mix(in srgb, " + layerStyle(layer).color + " 12%, transparent)",
                        },
                    }),
                    layerStyle(layer).label,
                )),
                kinds.map((kind) => h("span", { key: "edge:" + kind, className: "om-legend", style: S.legendItem },
                    h("span", { style: { display: "inline-block", width: "14px", height: "2px", borderRadius: "2px", background: edgeColor(kind) } }),
                    kind,
                )),
            );

            const focusNode = focus === null ? null : nodes.find((node) => node.key === focus) || null;
            const neighbourCount = focusSet === null ? 0 : focusSet.size;
            const inspector = h("div", {
                style: S.graphInspector,
                role: "status",
                "data-graph-detail": focusNode === null ? "none" : focusNode.key,
            },
                h("div", { style: S.inspectorKicker }, "GRAPH INSPECTOR"),
                focusNode === null
                    ? h("div", { style: S.inspectorEmpty },
                        h("div", { style: S.inspectorEmptyGlyph }, "◎"),
                        h("div", { style: S.inspectorTitle }, "悬停或选中一个节点"),
                        h("div", { style: S.inspectorHint }, "高亮它的邻居，详情显示在这一栏。滚轮缩放、拖背景平移、拖节点挪位置；方向键微调，Enter / 空格选中。"),
                    )
                    : h("div", null,
                        h("div", { style: S.inspectorTitle }, focusNode.label),
                        h("div", { style: S.inspectorMeta },
                            h("span", { style: S.inspectorMetaRow }, h("span", { style: S.inspectorMetaKey }, "层"), layerStyle(focusNode.layer).label),
                            h("span", { style: S.inspectorMetaRow }, h("span", { style: S.inspectorMetaKey }, "邻居"), String(neighbourCount) + " 个"),
                            h("span", { style: S.inspectorMetaRow }, h("span", { style: S.inspectorMetaKey }, "状态"), selected === focusNode.key ? "已选中" : "预览"),
                        ),
                        h("div", { style: S.inspectorHint },
                            focusNode.label + " ｜ " + layerStyle(focusNode.layer).label
                            + " ｜ 邻居 " + neighbourCount + " 个"
                            + (selected === focusNode.key ? " ｜ 已选中" : "")),
                    ),
            );

            // 视图条：缩放三件套 + 两个布局动作。放在画布**上方**（不是塞进图例那条
            // 页脚）：缩放是常用动作，藏在页脚里等于没有。
            const viewButton = (key, on, attrs, label, onClick) => h("button", Object.assign({
                key: key, type: "button", className: "om-view",
                style: on ? Object.assign({}, S.viewButton, S.viewButtonOn) : S.viewButton,
                onClick: onClick,
            }, attrs), label);
            const controls = h("div", { style: S.graphControls, "data-graph-controls": "true" },
                h("span", { key: "view-label", style: S.controlLabel }, "视图"),
                viewButton("zoom-out", false, { "data-graph-zoom-out": "true", "aria-label": "缩小记忆关系图" }, "−",
                    () => zoomAt(1 / GRAPH_ZOOM_STEP, null)),
                h("span", { key: "zoom", style: S.zoomLabel, "data-graph-zoom": "true" }, Math.round(view.k * 100) + "%"),
                viewButton("zoom-in", false, { "data-graph-zoom-in": "true", "aria-label": "放大记忆关系图" }, "+",
                    () => zoomAt(GRAPH_ZOOM_STEP, null)),
                viewButton("reset-view", false, { "data-graph-reset-view": "true", "aria-label": "重置视图（缩放与平移）" }, "重置视图", resetView),
                h("span", { key: "spacer", style: S.controlSpacer }),
                h("span", { key: "layout-label", style: S.controlLabel }, "布局"),
                viewButton("layout-force", layout === "force",
                    { "data-graph-layout": "force", "aria-pressed": layout === "force" }, "自然铺开",
                    () => applyLayout("force")),
                viewButton("layout-uniform", layout === "uniform",
                    { "data-graph-layout": "uniform", "aria-pressed": layout === "uniform" }, "均匀重置",
                    () => applyLayout("uniform")),
            );

            return h("div", { className: "om-grid-graph", style: S.graphArea },
                h("div", { style: S.graphCanvasWrap }, controls, svg, legend),
                inspector,
            );
        }

        /** 关系图的容器：图本身没什么可画时也给一句说明，而不是留一片空白。 */
        function GraphPanel(props) {
            const model = props.model;
            if (model.nodes.length === 0) {
                return h("div", { style: S.panel, "data-panel": "graph" },
                    h("div", { style: S.panelTitle },
                        h("span", { style: S.panelTitleText }, "关系图"),
                        h("span", { style: S.panelNote }, "0 个节点"),
                    ),
                    h("div", { style: S.emptyInline }, "这一层还没有节点或关系。"),
                );
            }
            return h("div", { style: S.panel, "data-panel": "graph" },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "关系图"),
                    h("span", { style: S.panelNote }, model.nodes.length + " 个节点 ｜ " + model.edges.length + " 条关系"),
                ),
                h(MemoryGraph, { nodes: model.nodes, edges: model.edges }),
            );
        }

        /** 标题下的六枚指标卡（mnemon 健康条那一路的读法：点 + 名称 + 数字）。 */
        function MetricTiles(props) {
            const counts = props.counts || {};
            return h("div", { style: S.tiles, "data-metrics": "memory" }, METRIC_TILES.map((tile) => {
                const value = finite(counts[tile.key]);
                return h("div", { key: tile.key, className: "om-card", style: S.tile, "data-metric": tile.key, title: tile.note },
                    h("div", { style: S.tileHead },
                        // 状态点带一圈同色光晕：8px 的点因此有了「亮着」的感觉。
                        h("span", {
                            style: Object.assign({}, S.tileDot, {
                                background: tile.color,
                                boxShadow: "0 0 0 3px color-mix(in srgb, " + tile.color + " 14%, transparent)",
                            }),
                        }),
                        h("span", { style: S.tileLabel }, tile.label),
                    ),
                    h("div", { style: S.tileValue, "data-metric-value": tile.key }, value === null ? "—" : String(value)),
                    h("div", { style: S.tileNote }, tile.note),
                );
            }));
        }

        /** 容量条：n / 上限，宽度按百分比。缺上限的项不画。 */
        function CapacityGauges(props) {
            const rows = Array.isArray(props.rows) ? props.rows : [];
            if (rows.length === 0) return null;
            return h("div", { style: S.panel, "data-panel": "capacity" },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "容量"),
                    h("span", { style: S.panelNote }, "满了向下沉"),
                ),
                rows.map((row) => h("div", { key: row.key, style: S.gauge, "data-gauge": row.key, "data-percent": row.percent },
                    h("div", { style: S.gaugeHead },
                        h("span", { style: S.gaugeName }, row.name),
                        h("span", { style: S.gaugeValue, "data-gauge-text": row.key },
                            countText(row.value) + " / " + countText(row.limit) + (row.unit ? " " + row.unit : "")),
                        h("span", { style: S.gaugePercent, "data-gauge-percent": row.key }, row.percent + "%"),
                    ),
                    h("div", { style: S.gaugeTrack },
                        h("div", {
                            "data-gauge-fill": row.key,
                            style: Object.assign({}, S.gaugeFill, {
                                width: row.percent + "%",
                                minWidth: row.percent > 0 ? "2px" : "0",
                                background: row.color || M.business,
                            }),
                        }),
                    ),
                )),
                h("div", { style: S.panelFoot }, "满了向下沉：台账滚成月度归档，热记忆沉进归档"),
            );
        }

        /** 存储域卡片：一个 stores[] 条目一张，含四层条数与可选的字节数。 */
        function StoreCards(props) {
            const stores = Array.isArray(props.stores) ? props.stores : [];
            if (stores.length === 0) return null;
            return h("div", { style: S.panel, "data-panel": "stores" },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "存储域"),
                    h("span", { style: S.panelNote }, stores.length + " 个目录"),
                ),
                h("div", { style: S.storeCards }, stores.map((store, index) => {
                    const entry = store || {};
                    const bytes = humanBytes(entry.bytes);
                    const files = finite(entry.files);
                    const chip = (key, text, layer) => h("span", { key: key, style: S.layerChip },
                        layer ? h("span", { style: Object.assign({}, S.layerChipDot, { background: layerStyle(layer).color }) }) : null,
                        text,
                    );
                    const chips = [
                        chip("hot", "热 " + countText(entry.hot), "hot"),
                        chip("ledger", "台账 " + countText(entry.ledger), "ledger"),
                        chip("archive", "归档 " + countText(entry.archive), "archive"),
                        chip("links", "关系 " + countText(entry.links), null),
                    ];
                    // 归档的「条数」与「摘要文件数」是两个量（一个月一个文件）：只写
                    // 「归档 30」看不出它只占 1 个文件，容量条那边就容易读成快满了。
                    const archiveFiles = finite(entry.archiveFiles);
                    if (archiveFiles !== null && archiveFiles > 0) {
                        chips.push(chip("archiveFiles", "摘要 " + countText(archiveFiles) + " 个", null));
                    }
                    // 知识库：文档数与块数一起给 —— 只有文档数看不出「一份文档切了多少块」，
                    // 而块数才是占用与检索面的那个量。
                    const kb = entry.kb && typeof entry.kb === "object" ? entry.kb : null;
                    const kbDocs = kb === null ? null : finite(kb.docs);
                    if (kbDocs !== null && kbDocs > 0) {
                        chips.push(chip("kb", "知识库 " + countText(kbDocs) + " 篇 / " + countText(kb.chunks) + " 块", null));
                    }
                    if (bytes !== null) chips.push(chip("bytes", bytes, null));
                    if (files !== null) chips.push(chip("files", files + " 个文件", null));
                    return h("div", {
                        key: entry.id ? String(entry.id) : String(index),
                        style: S.storeCard,
                        "data-store": entry.id ? String(entry.id) : String(index),
                    },
                        h("div", { style: S.storeHead },
                            h("span", { style: S.storeDir }, String(entry.dir || "（未知目录）")),
                        ),
                        h("div", { style: S.storeRow }, chips),
                    );
                })),
            );
        }

        /** 时间线：一条一条小柱子，高度按该时段的条数。 */
        function TimelineBars(props) {
            const buckets = Array.isArray(props.buckets) ? props.buckets : [];
            if (buckets.length === 0) return null;
            let max = 1;
            for (const bucket of buckets) max = Math.max(max, bucket.count);
            return h("div", { style: S.panel, "data-panel": props.name, "data-buckets": buckets.length },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, props.title),
                    h("span", { style: S.panelNote }, buckets.length + " 个时段 ｜ 峰值 " + max + " 条"),
                ),
                h("div", { style: S.timeline }, buckets.map((bucket) => h("div", {
                    key: bucket.key,
                    style: S.timelineCol,
                    "data-bucket": bucket.key,
                    "data-count": bucket.count,
                    title: bucket.label + "：" + bucket.count + " 条",
                },
                    h("span", { style: S.timelineCount }, String(bucket.count)),
                    h("div", {
                        style: Object.assign({}, S.timelineBar, {
                            height: Math.max(2, Math.round((bucket.count / max) * 44)) + "px",
                            background: props.color,
                            boxShadow: "0 0 0 3px color-mix(in srgb, " + props.color + " 10%, transparent)",
                        }),
                    }),
                    h("div", { style: S.timelineLabel }, bucket.label),
                ))),
            );
        }

        /** 实体频次条 + 每个实体指向的层 chip。 */
        function EntityBars(props) {
            const entities = Array.isArray(props.entities) ? props.entities : [];
            if (entities.length === 0) return null;
            let max = 1;
            for (const item of entities) {
                const count = finite(item && item.count);
                if (count !== null) max = Math.max(max, count);
            }
            return h("div", { style: S.panel, "data-panel": "entities" },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "实体频次"),
                    h("span", { style: S.panelNote }, "条长 = 出现次数 / 最高 " + max + " 次"),
                ),
                entities.slice(0, 20).map((item, index) => {
                    const entry = item || {};
                    const name = String(entry.name || "");
                    const count = finite(entry.count);
                    const percent = Math.min(100, Math.round(((count === null ? 0 : count) / max) * 1000) / 10);
                    const refs = Array.isArray(entry.refs) ? entry.refs : [];
                    // 每个实体指向的层各一枚 chip；同一层出现多次就带上次次
                    // （六条 refs 全是热记忆时画六枚一模一样的 chip 只是噪声）。
                    const chips = [];
                    const chipAt = new Map();
                    for (const ref of refs) {
                        const layer = ref && typeof ref.layer === "string" && ref.layer !== "" ? ref.layer : "";
                        if (layer === "") continue;
                        const at = chipAt.get(layer);
                        if (at === undefined) {
                            chipAt.set(layer, chips.length);
                            chips.push({ layer: layer, times: 1 });
                        } else {
                            chips[at].times += 1;
                        }
                    }
                    return h("div", { key: name || index, style: S.entityBlock, "data-entity": name, "data-percent": percent },
                        h("div", { style: S.entityRow },
                            h("span", { style: S.entityName, title: name }, name),
                            h("div", { style: S.entityTrack },
                                h("div", {
                                    "data-entity-fill": name,
                                    style: Object.assign({}, S.entityFill, {
                                        width: percent + "%",
                                        background: "linear-gradient(90deg, color-mix(in srgb, " + M.success + " 55%, transparent), " + M.success + ")",
                                    }),
                                }),
                            ),
                            h("span", { style: S.entityCount }, "出现 " + (count === null ? 0 : count) + " 次"),
                        ),
                        h("div", { style: S.layerChips }, chips.map((chip) =>
                            h("span", { key: chip.layer, style: S.layerChip, "data-entity-layer": chip.layer, "data-entity-refs": chip.times },
                                h("span", { style: Object.assign({}, S.layerChipDot, { background: layerStyle(chip.layer).color }) }),
                                layerStyle(chip.layer).label + (chip.times > 1 ? " ×" + chip.times : "")),
                        )),
                    );
                }),
            );
        }

        /**
         * 面板的样式表：只放内联样式写不了的东西 —— 伪类（悬停 / 焦点）、
         * 关键帧（加载转圈）、媒体查询（窄屏塌列、减少动效）。
         *
         * 为什么可以这么写：它是一段**固定字符串**，渲染成面板树里的一个 <style>
         * 节点。React 允许在树里渲染 style 元素，所以这既不需要在模块顶层碰
         * document（测试与 SSR 环境里没有 document，模块顶层碰它就整包抛错），
         * 也没有任何插值 —— 字符串由常量拼成，没有注入面。
         *
         * 约束（与内联样式同一套）：类名一律 om- 前缀，避免撞宿主；令牌只用
         * 宿主真实存在的 --dsw-*，并且每个都带字面兜底。
         */
        const PANEL_CSS = [
            ".om-card{transition:border-color .16s ease,box-shadow .16s ease,background-color .16s ease}",
            ".om-card:hover{border-color:" + M.strokeStrong + ";box-shadow:" + M.lift + "}",
            ".om-tab{transition:color .16s ease,background-color .16s ease}",
            ".om-tab:hover{color:" + M.textSecondary + ";background:" + M.hover + "}",
            ".om-btn{transition:color .16s ease,border-color .16s ease}",
            ".om-btn:hover{color:" + M.text + ";border-color:" + M.strokeStrong + "}",
            ".om-btn:active{transform:translateY(1px)}",
            ".om-input:focus{border-color:color-mix(in srgb," + M.business + " 55%," + M.stroke + ")}",
            ".om-card:focus-visible,.om-tab:focus-visible,.om-btn:focus-visible,.om-act:focus-visible,.om-view:focus-visible{outline:2px solid color-mix(in srgb," + M.business + " 58%,transparent);outline-offset:2px}",
            ".om-act{transition:color .16s ease,border-color .16s ease,background-color .16s ease}",
            ".om-act:hover{color:" + M.text + ";border-color:" + M.strokeStrong + ";background:" + M.hover + "}",
            ".om-view{transition:color .16s ease,border-color .16s ease,background-color .16s ease}",
            ".om-view:hover{color:" + M.text + ";border-color:" + M.strokeStrong + "}",
            ".om-node{transition:opacity .16s ease}",
            ".om-node:focus-visible .om-halo{opacity:.34}",
            ".om-halo,.om-edge{transition:opacity .16s ease}",
            ".om-legend{transition:color .16s ease}",
            ".om-legend:hover{color:" + M.textSecondary + "}",
            // 列表是封顶滚动容器：点「再显示 N 条」之后滚到底部去看新露出来的那几条。
            // 平滑滚动写在 CSS 里，就是为了让下面那条 reduced-motion 规则**管得住它**
            // —— 媒体查询只压 transition / animation 的话，滚动照样是平滑的。
            ".om-scroll{scroll-behavior:smooth}",
            ".om-scroll::-webkit-scrollbar{width:8px;height:8px}",
            ".om-scroll::-webkit-scrollbar-thumb{border-radius:999px;background:var(--dsw-alias-scrollbar-bg-l2," + M.sunken + ")}",
            ".om-scroll::-webkit-scrollbar-thumb:hover{background:var(--dsw-alias-scrollbar-hover-l2," + M.hover + ")}",
            ".om-spin{animation:om-spin .72s linear infinite}",
            "@keyframes om-spin{to{transform:rotate(360deg)}}",
            // 窄屏塌成一列：只有「图谱 + 详情」这一处需要（详情栏固定 240-270px，
            // 画布再被压就太小了）。容量 + 存储域那两栏用 auto-fit 自己就会塌，
            // 不需要额外的规则 —— 一刀切会让 978px 的面板白丢一栏。
            // 这里用的是**容器查询**而不是视口媒体查询：面板的实际宽度跟窗口宽度
            // 不是一回事（同一窗口下侧栏折叠与否能差 300px），按视口判断会出现
            // 「窗口 1280 但面板只有 660px，两栏挤在一起」这种结果。
            // 容器查询需要根元素声明 container-type（见 S.browser），宿主本身也在用
            // 容器查询，所以支持没问题；媒体查询留着当老浏览器的兜底。
            "@media (max-width:1080px){.om-grid-graph{grid-template-columns:minmax(0,1fr) !important}}",
            "@container (max-width:1080px){.om-grid-graph{grid-template-columns:minmax(0,1fr) !important}}",
            // 减少动效：三样都要压住 —— transition / animation / **scroll-behavior**。
            // 只压前两样时，「再显示 N 条」这类会滚动的动作仍然平滑移动，对前庭敏感的
            // 用户来说恰恰是幅度最大的那一类动效。scroll-behavior 用 !important 覆盖
            // 到整棵子树，因为宿主也可能在上层设过平滑滚动。
            "@media (prefers-reduced-motion:reduce){.om-mem,.om-mem *{transition-duration:.01ms !important;animation-duration:.01ms !important;animation-iteration-count:1 !important;scroll-behavior:auto !important}}",
        ].join("\n");

        /** 搜索框里的小放大镜（内联 SVG：bundle 是单文件产物，外链图标跟不走）。 */
        function SearchGlyph() {
            return h("svg", {
                width: 13, height: 13, viewBox: "0 0 16 16", fill: "none",
                stroke: "currentColor", strokeWidth: 1.5, strokeLinecap: "round",
                "aria-hidden": "true",
            },
                h("circle", { cx: 7, cy: 7, r: 4.5 }),
                h("path", { d: "M10.5 10.5 14 14" }),
            );
        }

        /**
         * 记忆浏览面板。
         *
         * 为什么把取数写成显式的 load()、而不是 useEffect 的依赖数组：本文件是
         * 手写的单文件产物，测试里那套 React 替身对 effect 只跑一次、依赖变化不
         * 重跑。把「取数」抽成一个函数，挂载时调一次，刷新 / 换目录 / 提交搜索时
         * 直接再调 —— 行为在真实 React 与替身里一致，也就测得到。
         */
        function MemoryBrowserPanel(props) {
            /** 客户端 layout 服务（可选）：待办指令上的「切到会话」用它切主面板。 */
            const layout = props && props.layout ? props.layout : null;
            const [state, setState] = React.useState({
                loading: true, error: null, data: null,
                tab: "hot", cwd: "", query: "", applied: "",
                // 悬停态放在面板这一层：卡片与页签都只是它的渲染结果，
                // 于是「悬停」不会给每个条目各建一份状态（也便于统一收敛）。
                hoverTab: null, hoverItem: null,
                // 「再显示 N 条」当前放到第几条。换页签 / 换目录 / 搜索都回到第一页，
                // 否则在 200 条的台账里翻到第 10 页，切回来会一口气铺 200 条。
                visible: LIST_PAGE,
                // 展开的条目 key（同一时刻只展开一条）。
                expandedItem: null,
                // 待办指令 { id, title, text }：面板**不落盘**，只把它交给用户确认。
                // copied 三态：null = 还没点，true = 复制成功，false = 这个环境不给剪贴板。
                pending: null, copied: null,
            });
            // 异步回来时要拿「最新的」state：闭包里那份是发起请求那一刻的旧值。
            const latest = React.useRef(state);
            latest.current = state;

            const patchState = (patch) => setState(Object.assign({}, latest.current, patch));

            const load = (patch) => {
                const base = Object.assign({}, latest.current, patch || {});
                // patch 必须落进 state，不能只用在这次请求的 URL 上：
                // 不落的话，搜索完再点刷新 / 换工作目录时 applied 又变回空，
                // 用户看到的过滤词被悄悄丢掉（第一次写就踩了这个坑）。
                patchState(Object.assign({}, patch || {}, { loading: true, error: null }));
                const params = [];
                if (base.cwd) params.push("cwd=" + encodeURIComponent(base.cwd));
                if (base.applied) params.push("q=" + encodeURIComponent(base.applied));
                const url = MEMORY_VIEW_PATH + (params.length > 0 ? "?" + params.join("&") : "");

                // fetch 只在浏览器里有。这里必须自己兜住：抛出去会让整格主面板白屏，
                // 而「没有 fetch」在测试与受限环境里是正常状态，不是异常。
                if (typeof fetch !== "function") {
                    patchState({ loading: false, error: "当前环境没有 fetch，无法读取记忆内容（这一格只在浏览器里工作）。" });
                    return;
                }
                let pending;
                try {
                    pending = fetch(url, { headers: { accept: "application/json" } });
                } catch (error) {
                    patchState({ loading: false, error: "读取记忆失败：" + (error && error.message ? error.message : String(error)) });
                    return;
                }
                // 端点失败时 HTTP 可能是 200 / 404 / 500，三种都要当数据看：
                // 只要 body 里带 error 就照实显示，不靠状态码判断。
                Promise.resolve(pending).then(
                    (response) => Promise.resolve(response && typeof response.json === "function" ? response.json() : null)
                        .then((body) => ({ response: response, body: body }), () => ({ response: response, body: null })),
                ).then((result) => {
                    const body = result.body;
                    const status = result.response && result.response.status ? result.response.status : "?";
                    if (body === null || typeof body !== "object") {
                        patchState({ loading: false, error: "记忆端点没有返回 JSON（HTTP " + status + "）。" });
                        return;
                    }
                    if (body.ok !== true) {
                        // 失败响应仍会带回 workspaces：留着，用户才好换一个目录再试。
                        patchState({
                            loading: false,
                            error: String(body.error || "读取记忆失败。"),
                            data: Array.isArray(body.workspaces) && body.workspaces.length > 0
                                ? { workspaces: body.workspaces }
                                : latest.current.data,
                        });
                        return;
                    }
                    const next = { loading: false, error: null, data: body };
                    // 第一次拿到数据时把 cwd 补上，后续请求才会带上它。
                    if (!latest.current.cwd && body.cwd) next.cwd = body.cwd;
                    patchState(next);
                }, (error) => {
                    patchState({ loading: false, error: "读取记忆失败：" + (error && error.message ? error.message : String(error)) });
                });
            };

            React.useEffect(() => { load(); }, []);

            const data = state.data || {};
            const hasData = state.data !== null && typeof state.data === "object";
            const counts = data.counts || {};
            const config = data.config;
            const limits = config && config.limits ? config.limits : {};
            const workspaces = Array.isArray(data.workspaces) ? data.workspaces : [];
            const stores = Array.isArray(data.stores) ? data.stores : [];
            const hot = Array.isArray(data.hot) ? data.hot : [];
            const ledger = Array.isArray(data.ledger) ? data.ledger : [];
            const archive = Array.isArray(data.archive) ? data.archive : [];
            const entities = Array.isArray(data.entities) ? data.entities : [];
            const items = Array.isArray(data[state.tab]) ? data[state.tab] : [];

            const head = h("div", { style: S.browserHead, "data-memory-head": "sticky" },
                h("div", { style: S.browserHeadIcon }, h(MemoryPanelIcon, { size: 16 })),
                h("div", { style: S.browserTitleBlock },
                    h("div", { style: S.browserTitle }, "记忆" + (data.label ? " · " + data.label : "")),
                    h("div", { style: S.browserSubtitle },
                        workspaces.length > 1
                            ? "浏览 · 数据来自各工作目录的 .office/memory；写入由会话里的 office_memory 执行"
                            : (data.cwd || "浏览 · 数据来自 .office/memory；写入由会话里的 office_memory 执行")),
                ),
                h("div", { style: S.browserTools },
                    // 只有一个工作目录时不必给下拉：多一个只有一项的控件只是噪声。
                    workspaces.length > 1
                        ? h(Select, {
                            value: state.cwd || data.cwd || "",
                            disabled: false,
                            options: workspaces.map((item) => ({ value: item.cwd, label: item.label || item.cwd })),
                            onChange: (next) => load({ cwd: next }),
                        })
                        : null,
                    h("div", { style: S.searchWrap },
                        h("span", { style: S.searchIcon }, h(SearchGlyph, null)),
                        h("input", {
                            type: "text",
                            className: "om-input",
                            defaultValue: state.query,
                            placeholder: "过滤台账 / 归档 / 实体，回车检索",
                            onKeyDown: (event) => {
                                if (event.key !== "Enter") return;
                                const next = String(event.target.value || "").trim();
                                load({ applied: next, query: next });
                            },
                            style: S.searchInput,
                        }),
                    ),
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-memory-compose": "true",
                        onClick: () => patchState({ pending: Object.assign({ id: "add" }, addInstruction()), copied: null }),
                    }, "记一条"),
                    h("button", { type: "button", className: "om-btn", style: S.ghostButton, onClick: () => load() }, "刷新"),
                ),
            );

            const tabs = h("div", { style: S.tabsTrack, role: "tablist" }, BROWSER_TABS.map((tab) => {
                const on = state.tab === tab.key;
                const hovered = state.hoverTab === tab.key && !on;
                const style = on ? Object.assign({}, S.browserTab, S.browserTabOn)
                    : hovered ? Object.assign({}, S.browserTab, S.browserTabHover) : S.browserTab;
                const tone = metricColor(tab.key);
                return h("div", {
                    key: tab.key,
                    className: "om-tab",
                    style: style,
                    role: "tab",
                    "aria-selected": on,
                    "data-tab": tab.key,
                    onClick: () => patchState({ tab: tab.key, visible: LIST_PAGE, expandedItem: null, hoverItem: null }),
                    onMouseEnter: () => patchState({ hoverTab: tab.key }),
                    onMouseLeave: () => patchState({ hoverTab: null }),
                },
                    h("span", { style: Object.assign({}, S.browserTabDot, {
                        background: tone,
                        // 选中的页签那颗点带一圈光晕：不靠加粗文字也能看出「我在这一格」。
                        boxShadow: on ? "0 0 0 3px color-mix(in srgb, " + tone + " 16%, transparent)" : "none",
                    }) }),
                    // 缺字段时页签与指标块用**同一个口径**：破折号，而不是 0。
                    // 「这一层是空的」与「这个端点没告诉我」是两件事（复核 P3-1）。
                    tab.label + "（" + (finite(counts[tab.key]) === null ? "—" : counts[tab.key]) + "）",
                );
            }));

            // 热记忆按 target 分组：这一层里「记给谁」比时间顺序重要得多。
            const hotGroups = state.tab === "hot" ? groupHot(items) : null;
            const hotTotals = new Map();
            if (hotGroups !== null) for (const group of hotGroups) hotTotals.set(group.key, group.items.length);

            // ── 「再显示 N 条」 ──
            //
            // 端点上各层已经有上限（200 / 200 / 500 / 60），所以这里只是**渐进式展开
            // 已经取回来的那一份**，不需要给端点加 offset —— 少一个参数就少一条要守
            // 的信任边界。分页决定「DOM 里有几条」，面板那一个滚动容器决定「一屏看几条」。
            const total = items.length;
            const shown = Math.max(0, Math.min(total, Math.max(LIST_PAGE, state.visible)));
            const hidden = Math.max(0, total - shown);
            // 滚动容器是**面板根元素**（listRef 现在指向它，不再指向列表本身）。
            const listRef = React.useRef(null);
            const showMore = () => {
                patchState({ visible: Math.min(total, shown + LIST_PAGE) });
                // 新露出来的那几条在列表底部：把它滚过去，用户才看得到「刚才多了什么」。
                // 平滑与否要看系统设置 —— `scrollTo({behavior:'smooth'})` 不读 CSS 的
                // `scroll-behavior`，所以 reduced-motion 必须在代码里再读一次。
                const el = listRef.current;
                if (el && typeof el.scrollTo === "function") {
                    try {
                        el.scrollTo({ top: el.scrollHeight, behavior: prefersReducedMotion() ? "auto" : "smooth" });
                    } catch (error) { /* 老浏览器不认 options 就算了 */ }
                }
            };

            /** 每个条目拿到的那组回调：悬停 / 展开 / 写入动作都汇到面板这一层。 */
            const itemProps = (item, index, tab) => {
                const key = itemKeyOf(item, index);
                return {
                    key: key, item: item, tab: tab, index: index,
                    hovered: state.hoverItem === key,
                    expanded: state.expandedItem === key,
                    onHover: (next) => patchState({ hoverItem: next }),
                    onToggle: (target) => patchState({ expandedItem: state.expandedItem === target ? null : target }),
                    onAction: (action, entry) => {
                        const built = action.build(entry);
                        patchState({ pending: { id: action.id, title: built.title, text: built.text }, copied: null });
                    },
                };
            };

            const list = hotGroups !== null
                ? h("div", { style: S.browserList, "data-memory-list": "hot" },
                    paginateGroups(hotGroups, shown).map((group) => h("div", { key: group.key, "data-hot-group": group.key },
                        h("div", { style: S.hotGroupTitle },
                            // 标题与条数必须是**同一个字符串**：测试按「用户偏好（1）」
                            // 整串核对，拆成两个 span 就对不上了。条数取**整组的总数**
                            // 而不是当前显示了几条 —— 否则「再显示 N 条」会让标题里的
                            // 数字越点越小，读起来像条目被删了。
                            h("span", { style: S.panelTitleText }, group.label + "（" + hotTotals.get(group.key) + "）"),
                        ),
                        group.items.map((item, index) => h(BrowserItem, itemProps(item, index, "hot"))),
                    )))
                : h("div", { style: S.browserList, "data-memory-list": state.tab },
                    items.slice(0, shown).map((item, index) => h(BrowserItem, itemProps(item, index, state.tab))));

            const moreBar = state.loading || total <= LIST_PAGE
                ? null
                : h("div", { style: S.listMore, "data-list-more": "true" },
                    hidden > 0
                        ? h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton, "data-show-more": "true",
                            onClick: showMore,
                        }, "再显示 " + Math.min(LIST_PAGE, hidden) + " 条")
                        : null,
                    shown > LIST_PAGE
                        ? h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton, "data-collapse-list": "true",
                            onClick: () => patchState({ visible: LIST_PAGE, expandedItem: null }),
                        }, "收起")
                        : null,
                    h("span", { style: S.listMoreNote },
                        "已显示 " + shown + " / " + total + (hidden > 0 ? "，还有 " + hidden + " 条" : "")),
                );

            // 条目详情条：与图谱右侧的 GRAPH INSPECTOR 同一套路。
            // 列表里刻意**不做浮动气泡**：列表是封顶滚动容器（overflow:auto），
            // 绝对定位的气泡在第一行会被容器上沿裁掉、最后一行会被下沿裁掉 ——
            // 这也是 mnemon 把信息放进固定详情栏而不是气泡的原因。
            const detailIndex = state.hoverItem === null
                ? -1
                : items.findIndex((item, index) => String(itemKeyOf(item, index)) === String(state.hoverItem));
            const detailItem = detailIndex >= 0 ? items[detailIndex] : null;
            const detailBar = detailItem === null ? null : h("div", {
                style: S.itemDetail, role: "status",
                "data-item-detail": String(itemKeyOf(detailItem, detailIndex)),
            },
                h("div", { style: S.itemDetailHead },
                    h("span", { style: S.itemDetailKey }, "ITEM DETAIL"),
                    h("span", { style: S.itemDetailTitle }, String(itemKeyOf(detailItem, detailIndex))),
                ),
                h("div", { style: S.itemDetailBody }, detailTextOf(detailItem, state.tab)),
                h("div", { style: S.itemDetailHint }, "完整内容。卡片上的「展开」可以就地铺开 —— 列表项不再用原生 title 提示。"),
            );

            // 归档页签的一条实证：把**真实总数**（counts.archiveItems，未过滤、未截断）
            // 与「这一页拿回来多少」摆在一起。端点上归档上限是 200 条，条目超过它时
            // 列表里那 200 条看起来像全部 —— 说清楚差多少，用户才知道该不该去会话里检索。
            // 面板有意**不做远端翻页**：那要给端点加 offset 参数，多一条要守的信任边界，
            // 而归档的全文检索本来就有 office_memory 的 read（带 since/until）。
            const archiveTotal = finite(counts.archiveItems);
            const archiveTruncated = archiveTotal !== null && archiveTotal > total;
            const archiveNote = state.tab === "archive" && total > 0
                ? h("div", { style: S.pendingHint, "data-archive-note": "true" },
                    ARCHIVE_READONLY_NOTE
                    + (archiveTotal === null ? "" : " 归档共 " + archiveTotal + " 条。")
                    + (archiveTruncated
                        ? "这一页只给了 " + total + " 条（浏览入口的单次上限），要看更旧的用 office_memory 的 read 检索。"
                        : ""))
                : null;

            // 知识库页签的同一件事：清单上限是 20 篇，文档再多也只列到 20 —— 把真实总数与
            // 块数摆出来，别让 20 看起来像全部。块正文不在这一页：要读哪一块，用会话里的
            // kb-read（面板只给文档 id 与块数）。
            //
            // 措辞不写「最新的 N 篇」：`at` 撞在同一毫秒时排序是稳定排序的产物，说「最新」
            // 会是一句撑不住的承诺（复核 P2-4）。说「这一页列了 N 篇」，并把总数放在旁边。
            const kbTotal = finite(counts.kbTotal);
            const kbTruncated = counts.kbTruncated === true || (kbTotal !== null && kbTotal > total);
            const kbNote = state.tab === "kb" && total > 0
                ? h("div", { style: S.pendingHint, "data-kb-note": "true" },
                    "知识库是只读页：入库、检索与读取都由会话里的 office_memory 执行"
                    + "（kb-ingest / kb-search / kb-read）。"
                    + (kbTotal === null ? "" : " 库里共 " + kbTotal + " 篇、"
                        + countText(counts.kbChunks) + " 块、"
                        + (humanBytes(counts.kbBytes) === null ? "—" : humanBytes(counts.kbBytes)) + "。")
                    + (kbTruncated
                        ? "这一页列了 " + total + " 篇（清单接口的单次上限），要看全部用 kb-list。"
                        : ""))
                : null;

            const body = state.loading
                ? h("div", { style: S.loadingState },
                    h("span", { className: "om-spin", style: S.spinner }),
                    "正在读取记忆…")
                : items.length === 0
                    ? h("div", { style: S.emptyState },
                        h("div", { style: S.emptyGlyph }, "◌"),
                        h("div", { style: S.emptyTitle }, "这一层还没有内容。"),
                        h("div", { style: S.emptyHint }, "用 office_memory({ action: 'add' }) 记一条偏好或约定；台账会在 office_run 写出第一份交付物时自动登记。"))
                    : h("div", null, archiveNote, kbNote, list, moreBar, detailBar);

            // ── 可视化块 ──
            //
            // 只在真的拿到过数据时画：错误页上不摆一排 0，那会让人以为记忆是空的。
            // 每一块自己判断有没有内容（返回 null 就什么都不占），所以这里不做
            // 逐块的条件判断，免得条件与组件里的判断各自漂移。
            const dashboards = [];
            const visuals = [];
            if (hasData) {
                visuals.push(h(MetricTiles, { key: "tiles", counts: counts }));
                const rows = capacityRows(counts, limits);
                if (rows.length > 0) dashboards.push(h(CapacityGauges, { key: "capacity", rows: rows }));
                if (stores.length > 0) dashboards.push(h(StoreCards, { key: "stores", stores: stores }));
                if (state.tab === "ledger") {
                    visuals.push(h(TimelineBars, {
                        key: "ledger-timeline", name: "ledger-timeline", title: "台账按天分布",
                        color: "#22a879", buckets: bucketize(ledger, ledgerDay),
                    }));
                }
                if (state.tab === "archive") {
                    visuals.push(h(TimelineBars, {
                        key: "archive-months", name: "archive-months", title: "归档按月分布",
                        color: "#c08a2e", buckets: bucketize(archive, archiveMonth),
                    }));
                }
                if (state.tab === "entities") {
                    visuals.push(h(EntityBars, { key: "entity-bars", entities: entities }));
                }
                // 关系图在「关系」和「实体」两格都画：两个 Tab 看的是同一张图的两个切面。
                if (state.tab === "links" || state.tab === "entities") {
                    visuals.push(h(GraphPanel, { key: "graph", model: buildGraphModel(data) }));
                }
            }

            const layersOn = config && config.layers
                ? Object.keys(config.layers).filter((key) => config.layers[key]).join("、")
                : "";
            const footerChip = (key, label, value) => h("span", { key: key, style: S.footerChip },
                h("span", { style: S.footerKey }, label), value);

            // 待办指令：面板**只生成指令文本**，不落盘。
            // 点条目上的「改写 / 忘记 / 删关系」或标题栏的「记一条」都会到这里；
            // 用户确认之后复制、切到会话发送，真正的写入由 office_memory 完成。
            // 这与回复旁的「存入记忆」是同一个口径：记忆会跨会话影响后面所有回合，
            // 该由用户先看见「要记什么」再点头。
            const pendingCard = state.pending === null ? null : h("div", {
                style: S.pending, "data-memory-pending": String(state.pending.id),
            },
                h("div", { style: S.pendingHead },
                    h("span", { style: S.pendingTitle }, "待发送指令 · " + state.pending.title),
                    h("span", { style: S.pendingNote }, "面板不写记忆，确认后由会话执行"),
                ),
                h("pre", { style: S.pendingCode, "data-memory-instruction": "true" }, state.pending.text),
                h("div", { style: S.pendingActions },
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton, "data-memory-copy": "true",
                        onClick: () => copyText(state.pending.text).then((ok) => patchState({ copied: ok })),
                    }, state.copied === true ? "已复制" : "复制指令"),
                    layout && typeof layout.selectPanel === "function"
                        ? h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton, "data-memory-jump": "true",
                            onClick: () => {
                                // 拿不到 layout 服务时这个按钮根本不画（见下），所以这里只管切。
                                try { layout.selectPanel("conversation"); } catch (error) { /* 面板切不动就算了 */ }
                            },
                        }, "切到会话")
                        : null,
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton, "data-memory-dismiss": "true",
                        onClick: () => patchState({ pending: null, copied: null }),
                    }, "关闭"),
                ),
                h("div", { style: S.pendingHint },
                    state.copied === true
                        ? "已复制。切到会话粘进输入框发送即可，office_memory 会把改完的回执带回来。"
                        : state.copied === false
                            ? "这个环境不给剪贴板权限：请手动选中上面那段文本复制。"
                            : "复制后粘到输入框发送。唯一命中校验、下沉与落盘都还在 office_memory 里，面板只负责把指令写准。"),
            );

            // 根元素就是**唯一的滚动容器**（见 S.browser 的注释）：om-scroll 把滚动条
            // 样式与 reduced-motion 下的 scroll-behavior 一起管住，ref 给「再显示 N 条」用。
            return h("div", { className: "om-mem om-scroll", style: S.browser, ref: listRef, "data-memory-panel": "scroll" },
                // 伪类 / 关键帧 / 媒体查询在这里；内容是常量，没有注入面。
                h("style", { key: "panel-css", dangerouslySetInnerHTML: { __html: PANEL_CSS } }),
                head,
                // 出错时列表照样显示（可能是上一次的数据）：只把错误顶在上面，
                // 而不是用错误页把已有内容盖掉。
                state.error !== null ? h("div", { style: S.error }, state.error) : null,
                pendingCard,
                tabs,
                visuals.slice(0, 1),
                dashboards.length > 0 ? h("div", { className: "om-grid-dash", style: S.panelGrid }, dashboards) : null,
                visuals.slice(1),
                body,
                config
                    ? h("div", { style: S.footerMeta },
                        footerChip("scope", "范围", String(config.scope || "—")),
                        footerChip("user", "用户偏好", String(config.userScope || "—")),
                        footerChip("layers", "开启的层", layersOn || "（无）"),
                        footerChip("links", "图关系", config.links ? "开" : "关"),
                        footerChip("capture", "主动记录", config.autoCapture ? "开" : "关"),
                    )
                    : null,
            );
        }

        // ── 会话内两处 ────────────────────────────────────────────────────

        /** 记忆相关的工具名前缀：本回合的「记忆活动」按这个筛。 */
        const MEMORY_TOOL = "office_memory";
        const SEARCH_TOOL_PREFIX = "office_search_";

        /** 从 tool 调用的 argsRaw 里读 action；读不出来（流式未完成 / 非法 JSON）返回空串。 */
        function actionOf(argsRaw) {
            try {
                const parsed = JSON.parse(String(argsRaw));
                return parsed && typeof parsed.action === "string" ? parsed.action : "";
            } catch {
                return "";
            }
        }

        /**
         * 统计某一回合的记忆活动：召回 / 沉淀 / 检索各几次。
         *
         * 数据来源是 standard props 的 useTrajectory 快照，两条通道合起来才完整：
         *   - `eventNodes` 里的 `tool-result` 是**已完成**的调用：它自带 seq，
         *     而「seq → 回合」的对应在 `eventLocations`（Map<seq, ConversationLocation>，
         *     location.kind 为 turn / step 时 location.turn.turn 就是回合号）。
         *     tool-result 自身没有 turn 字段，只能这样换算。
         *   - `runningCalls` 是**进行中**的调用，自带 name 与 turn，直接算。
         *     只看前者会让「正在查资料」这一回合什么都不显示。
         *
         * 这两处形状都取自 dsh-client-ui-conversation / dsh-client-ui-trajectory 的
         * 公开 .d.ts（records.d.ts 的 ToolResultNode / RunningToolCall，
         * conversation.d.ts 的 ConversationLocation），不是猜的。
         */
        function countTurnMemoryActivity(snapshot, turn) {
            const stats = { recall: 0, write: 0, search: 0, other: 0, total: 0 };
            const bump = (name, argsRaw) => {
                if (typeof name !== "string") return;
                if (name === MEMORY_TOOL) {
                    const action = actionOf(argsRaw);
                    if (action === "read") stats.recall += 1;
                    else if (action === "add" || action === "replace" || action === "remove"
                        || action === "link" || action === "unlink") stats.write += 1;
                    else stats.other += 1;
                    stats.total += 1;
                    return;
                }
                if (name.indexOf(SEARCH_TOOL_PREFIX) === 0) {
                    stats.search += 1;
                    stats.total += 1;
                }
            };

            const nodes = snapshot && Array.isArray(snapshot.eventNodes) ? snapshot.eventNodes : [];
            const locations = snapshot ? snapshot.eventLocations : undefined;
            for (const node of nodes) {
                if (!node || node.kind !== "tool-result" || !node.call) continue;
                const name = node.call.name;
                if (typeof name !== "string" || name.indexOf("office_") !== 0) continue;
                const where = locations && typeof locations.get === "function" ? locations.get(node.seq) : undefined;
                const nodeTurn = where && where.turn && typeof where.turn.turn === "number" ? where.turn.turn : null;
                if (nodeTurn !== turn) continue;
                bump(name, node.call.argsRaw);
            }

            const running = snapshot && Array.isArray(snapshot.runningCalls) ? snapshot.runningCalls : [];
            for (const call of running) {
                if (!call || call.turn !== turn) continue;
                bump(call.name, call.argsRaw);
            }
            return stats;
        }

        /**
         * 回合记忆条：显示这一回合的记忆活动（召回 / 沉淀 / 检索各几次）。
         *
         * 两处「读不到就不画」的取舍，都是刻意的：
         *   - 拿不到 `useTrajectory`（测试替身 / 部署里没挂 Trajectory）：返回 null。
         *   - 本回合一次记忆 / 检索工具都没调：返回 null。slot 的 catalog 明确写了
         *     「entries without content return null」——画一条「本回合 0 次」是在
         *     每个回合尾部加一行噪声，而它并不提供任何信息。
         *
         * 关于 hooks 的写法：`useTrajectory` 是宿主注入的 standard prop，对一个已挂载的
         * slot 条目来说是常量，所以「它存在才调用」不会让同一个实例的 hook 顺序发生变化。
         * 选择器返回的是快照里已有的引用（快照对象本身），不是每次新建的对象 ——
         * 这是 SnapshotSelectorHook 能收敛的前提（返回新对象会重渲染不收敛）。
         */
        function TurnMemoryBar(props) {
            const useTrajectory = props ? props.useTrajectory : undefined;
            const snapshot = typeof useTrajectory === "function" ? useTrajectory((value) => value) : null;
            const turn = props && props.turn && typeof props.turn.turn === "number" ? props.turn.turn : null;
            if (turn === null || snapshot === null || snapshot === undefined) return null;
            const stats = countTurnMemoryActivity(snapshot, turn);
            if (stats.total === 0) return null;
            const parts = [];
            if (stats.recall > 0) parts.push("召回 " + stats.recall);
            if (stats.write > 0) parts.push("沉淀 " + stats.write);
            if (stats.search > 0) parts.push("检索 " + stats.search);
            if (stats.other > 0) parts.push("其它 " + stats.other);
            return h("div", { style: S.turnBar, title: "本回合的办公记忆与检索活动" },
                "📒 本回合记忆活动：" + parts.join(" ｜ "));
        }

        /** 放进输入框的那句话。写死在代码里，免得和界面文案各自漂移。 */
        const SAVE_TO_MEMORY_PROMPT = "用 office_memory 把上一条回复里值得长期保留的事实记下来（target / importance 你自己判断，只记稳定的偏好、约定与事实）";

        /**
         * 定稿回复旁的「存入记忆」。
         *
         * 为什么只 setDraft 不 submit：这是**受监督**的入口。记忆一旦写进去就会
         * 跨会话影响后面所有回合，用户该先看见这句话、确认之后再发；直接提交等于
         * 让一次点击替用户决定了「记什么」。
         *
         * 拿不到输入通道时返回 null —— 画一个点了没反应的按钮比不画更糟。
         */
        function SaveToMemoryAction(props) {
            const actions = props && props.inputActions;
            if (!actions || typeof actions.setDraft !== "function") return null;
            return h("button", {
                type: "button",
                style: S.button,
                title: "把这句话放进输入框，看过再发",
                onClick: () => actions.setDraft(SAVE_TO_MEMORY_PROMPT),
            }, "存入记忆");
        }

        // ── 注册 ──────────────────────────────────────────────────────────

        /**
         * 必需服务。
         *
         * settingsScope 这条路在 0.1.7 已经不存在了：现在按「宿主插件行 id」
         * 去 ctx.configForms 取一份 ConfigForm。**只声明 slots** ——
         * configForms 是可选服务，拿不到时设置页退化成只读提示就行；
         * 若把它列进 inject，缺服务时整个 apply 都不会跑，记忆面板也跟着消失
         * （0.1.7 这次更新踩的正是这个坑：注入列表里留着一个已经没有的服务名）。
         */
        const inject = ["slots"];

        function apply(ctx) {
            if (!ctx || !ctx.slots) return;

            // ConfigForm 只能用 ctx.inject 取，**不能**直接读 ctx.configForms：
            // 客户端上下文里访问一个没写进 inject 的服务属性会当场抛错，整个 entry
            // 因此变成 failed（实测报错：web boot: 1 entry did not activate /
            // dsh-office-mode: failed，而宿主侧一切正常）。
            // 反过来把 configForms 写进 inject 也是个坑：缺服务时 apply 根本不跑，
            // 记忆面板会跟着一起消失。ctx.inject 的回调只在服务就绪时执行 ——
            // 两条路都避开了，设置页拿不到服务时退化成只读提示。
            let scope = null;
            if (typeof ctx.inject === "function") {
                try {
                    ctx.inject(["configForms"], (sctx) => {
                        try {
                            scope = formToScope(sctx.configForms.get(NAMESPACE));
                        } catch (error) {
                            scope = null;
                        }
                    });
                } catch (error) {
                    scope = null;
                }
            }

            // layout 服务同样是**可选**的，理由和 configForms 一样：写进 inject 的话，
            // 一个没有 layout 的部署会让 apply 根本不跑、记忆面板跟着消失。
            // 拿不到时面板照常渲染，只是待办指令上少一个「切到会话」按钮
            // （复制 + 手动切换仍然可用），比画一个点了没反应的按钮好。
            let layout = null;
            if (typeof ctx.inject === "function") {
                try {
                    ctx.inject(["layout"], (sctx) => {
                        layout = sctx && sctx.layout ? sctx.layout : null;
                    });
                } catch (error) {
                    layout = null;
                }
            }

            // 两个一级设置导航项，共用同一个命名空间（SettingsPanel 是同一套读写接线）：
            //   - 记忆系统（order 519）：三层记忆的开关、容量与迁移说明。原来这个位置
            //     是 mnemon 提供的「记忆系统」，mnemon 关掉后由本插件接住；
            //   - 办公模式（order 520）：工具开关、检索编排、文档与缓存。
            ctx.slots.inject("settings.section", () =>
                ctx.slots.register(
                    {
                        name: "settings.section",
                        id: MEMORY_SECTION_ID,
                        order: 519,
                        label: () => h(MemoryNavLabel, null),
                    },
                    () => h(MemorySettingsSection, { scope: scope }),
                ),
            );

            ctx.slots.inject("settings.section", () =>
                ctx.slots.register(
                    {
                        name: "settings.section",
                        id: NAMESPACE,
                        order: 520,
                        label: () => h(OfficeNavLabel, null),
                    },
                    () => h(OfficeSettingsSection, { scope: scope }),
                ),
            );

            // ── 记忆浏览面板：侧栏一个入口 + 主面板一格 ──
            //
            // 两处的标识必须是**同一个字符串**（MEMORY_PANEL_ID）：sidebar.panellist
            // 的 id 就是 main 的 key，侧栏按钮靠「id === key」找到并点亮主面板那一格。
            // 写成两个名字会得到一个点不亮的面板，而两边都不会报错。
            //
            // 作用域提醒：这两个 slot 都是 root 作用域（整框架一份），不是 session
            // 作用域 —— 所以注册在 apply 里、跟着插件生命周期走。
            ctx.slots.inject("sidebar.panellist", () =>
                ctx.slots.register(
                    { name: "sidebar.panellist", id: MEMORY_PANEL_ID, order: 30, label: () => "记忆" },
                    MemoryPanelIcon,
                ),
            );

            ctx.slots.inject("main", () =>
                ctx.slots.register(
                    { name: "main", key: MEMORY_PANEL_ID },
                    () => h(MemoryBrowserPanel, { layout: layout }),
                ),
            );

            // ── 会话内两处（session 作用域） ──
            ctx.slots.inject("conversation.chat.turnTail", () =>
                ctx.slots.register(
                    { name: "conversation.chat.turnTail", id: "office-memory-turnbar", order: 30 },
                    TurnMemoryBar,
                ),
            );

            ctx.slots.inject("conversation.chat.assistant-actions", () =>
                ctx.slots.register(
                    { name: "conversation.chat.assistant-actions", id: "office-memory-save", order: 30 },
                    SaveToMemoryAction,
                ),
            );
        }

        exports.apply = apply;
        exports.inject = inject;
        return module.exports;
    },
});