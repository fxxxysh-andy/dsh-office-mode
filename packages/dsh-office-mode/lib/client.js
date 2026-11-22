/**
 * 办公模式设置页（浏览器半侧）。
 *
 * 这一页注册成设置对话框里的一级导航项（settings.section），和
 * 「Wallpaper Engine」同一层级、同样的交互形态：左边一个带图标的入口，
 * 右边是可操作的控件（开关、下拉、数字输入）。第一百零五轮起，办公与记忆
 * 合并成这一页，内部再用页签分屏。
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
 * 骨架是**可折叠的分组**：页签内默认展开，进页面直接看到控件，不再是
 * 一屏分组目录；顶上一条「搜索设置项」框按名称 / 配置项过滤。细节见下面
 * 「分组折叠与定位」那一段与 SettingsGroup / SettingsBody 的说明。
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

        /** 工具开关的显示文案。顺序即界面顺序。 */
        const TOOL_ROWS = [
            { key: "office_help", label: "office_help", note: "按需查 Word / Excel / PPT 的写法与主题。关掉会明显增加写错 API 的概率" },
            { key: "office_run", label: "office_run", note: "批量执行：生成或修改三件套，以及批量改文本文件" },
            { key: "office_memory", label: "office_memory", note: "三层记忆：读写热记忆（偏好与约定）、检索台账与归档" },
            { key: "office_web_search", label: "office_web_search", note: "查一个事实点（网页抓取通道，免 Key）：来源直接回到对话，不落盘" },
            { key: "office_web_fetch", label: "office_web_fetch", note: "打开一个页面取正文（网页抓取通道）；PDF 由 office.pdf 抽文本" },
            { key: "office_search_run", label: "office_search_run", note: "内置检索直查：给几条查询，插件的联网通道去找来源并落盘" },
            { key: "office_search_dispatch", label: "office_search_dispatch", note: "渠道覆盖检索：给 topic 自动出提纲并逐渠道执行" },
            { key: "office_parse_findings", label: "office_parse_findings", note: "把渠道结果文件读成摘要（派工之后那一步）" },
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

        /** 知识库向量化的通道。值必须与服务端 schema 的 union 一致。 */
        const EMBEDDING_MODE_OPTIONS = [
            { value: "auto", label: "auto —— 有端点或本地推理就用，否则只用词法" },
            { value: "remote", label: "remote —— 只用远端（失败退回本地并标记）" },
            { value: "local", label: "local —— 只用本地哈希向量（离线，非语义）" },
            { value: "off", label: "off —— 关掉向量检索，只用词法" },
        ];

        /** 候选筛选策略（第一百零三轮替代重排）：让侧路模型从候选里挑编号。 */
        const SELECTOR_MODE_OPTIONS = [
            { value: "off", label: "off —— 只用融合排序，不调模型" },
            { value: "llm", label: "llm —— 侧路模型只看编号与块头，回编号" },
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
         * 取正文与 PDF 的上限（与用哪条检索通道无关）。
         *
         * 写回路径写成**字面量**：设置面板的「界面覆盖服务端每个参数」那条测试按
         * 源码里的点号路径逐项核对，拼出来的路径在源码里根本不存在。
         */
        const FETCH_ROWS = [
            { key: "maxResults", path: "search.fetch.maxResults", label: "每次来源上限", note: "每个渠道 / 每条直查查询最多几条来源（1-20）", kind: "number", min: 1, max: 20, step: 1 },
            { key: "fetchPages", path: "search.fetch.fetchPages", label: "取正文页数", note: "每个渠道 / 每条查询最多打开几页取摘录（0-5）；0 = 只用搜索返回的片段", kind: "number", min: 0, max: 5, step: 1 },
            { key: "fetchTimeoutMs", path: "search.fetch.fetchTimeoutMs", label: "取正文超时", note: "单页取正文的超时（毫秒）", kind: "number", min: 5000, max: 300000, step: 5000, unit: "ms" },
            { key: "maxBytes", path: "search.fetch.maxBytes", label: "单页体积上限", note: "单页正文的体积上限（字节），超出即截断", kind: "number", min: 32768, max: 33554432, step: 32768, unit: "B" },
            { key: "maxChars", path: "search.fetch.maxChars", label: "单页字符上限", note: "单页正文的字符上限，超出即截断", kind: "number", min: 1000, max: 400000, step: 1000, unit: "字" },
            { key: "maxRedirects", path: "search.fetch.maxRedirects", label: "跳转次数上限", note: "取正文最多跟随几次同源跳转（跨站跳转一律不跟）", kind: "number", min: 0, max: 10, step: 1, unit: "次" },
            { key: "pdfMaxBytes", path: "search.fetch.pdfMaxBytes", label: "PDF 体积上限", note: "取正文撞上 PDF 时的体积上限（字节）。超了直接放弃（截断的 PDF 抽不出文本）：换来源，或先下载到本地用 office.pdf 读", kind: "number", min: 1048576, max: 134217728, step: 1048576, unit: "B" },
            { key: "pdfMaxPages", path: "search.fetch.pdfMaxPages", label: "PDF 页数上限", note: "PDF 抽文本最多抽前几页，用来圈住抽取耗时（1-500）", kind: "number", min: 1, max: 500, step: 1, unit: "页" },
        ];

        /**
         * 记忆层拓扑的三行。
         *
         * 写回路径写成**字面量**（而不是 "memory.layers." + key 拼出来）：设置面板
         * 的「界面覆盖服务端每个参数」那条测试是按源码里的点号路径逐项核对的，
         * 拼出来的路径在源码里根本不存在，等于把一个开关漏在了界面之外。
         */
        const LAYER_ROWS = [
            { key: "hot", path: "memory.layers.hot", label: "热记忆层", note: "用户偏好与项目约定，随 office_help / office_run 投影。" },
            { key: "ledger", path: "memory.layers.ledger", label: "台账层", note: "每份交付物自动登记；关掉后不再登记。" },
            { key: "archive", path: "memory.layers.archive", label: "归档层", note: "下沉的旧条目与旧台账；关掉后检索不到，文件仍在。" },
        ];

        /**
         * 记忆浏览面板的两个**一级**视图：记忆 / 知识库。
         *
         * 一级只回答「在看哪一类东西」，记忆内部的层级（热记忆 / 台账）降为二级
         * 分段；归档本来就已经作为知识库里的「归档/YYYY-MM」文件夹存在，不另开入口。
         * 第九十一轮按用户要求下线了旧「图谱」（图关系 + 实体视图）：「关系」只在
         * 知识库的块级 kb-link / kb-related 里保留，面板不再提供整库图谱。
         */
        const BROWSER_TABS = [
            { key: "memory", label: "记忆" },
            // 第四十二轮：知识库（kb 一期）的文档清单。面板只读：入库 / 读取都由
            // 会话里的 office_memory 执行，这里给的是 manifest 行（路径 / 块数 / 档位）。
            { key: "kb", label: "知识库" },
            // 审美参考图库：原图 / 缩略图 / 手工标签与分类 / 文字与以图搜图。
            // 图库是独立的一级视图，不复用知识库的文档行。
            { key: "gallery", label: "图库" },
        ];


        /** 当前 tab 属于哪个一级视图。 */
        function sectionOf(tab) {
            if (tab === "kb") return "kb";
            if (tab === "gallery") return "gallery";
            return "memory";
        }

        /**
         * 一致的文件 / 文件夹图标（Windows 11 资源管理器口径）。
         *
         * 不用 emoji / 字符拼图标：不同系统字体下字宽与颜色都不可控，深色主题里
         * 还会变成彩色表情。这里用内联 SVG + currentColor，尺寸与颜色都跟宿主令牌走。
         */
        function ExplorerGlyph(props) {
            const kind = String((props && props.kind) || "file");
            const size = Number((props && props.size) || 18);
            const common = {
                width: size, height: size, viewBox: "0 0 24 24", fill: "none",
                stroke: "currentColor", strokeWidth: 1.6, strokeLinecap: "round",
                strokeLinejoin: "round", "aria-hidden": "true", focusable: "false",
            };
            if (kind === "folder") {
                return h("svg", common,
                    h("path", { d: "M3 7.5A2.5 2.5 0 0 1 5.5 5h4l2 2.5H18.5A2.5 2.5 0 0 1 21 10v7.5a2.5 2.5 0 0 1-2.5 2.5h-13A2.5 2.5 0 0 1 3 17.5z" }));
            }
            if (kind === "image") {
                return h("svg", common,
                    h("rect", { x: 3.5, y: 4.5, width: 17, height: 15, rx: 2 }),
                    h("circle", { cx: 9, cy: 10, r: 1.5 }),
                    h("path", { d: "M4.5 17l4.2-4.2 3.1 3.1 2.4-2.4 5.3 5.2" }));
            }
            if (kind === "gallery") {
                return h("svg", common,
                    h("rect", { x: 3.5, y: 3.5, width: 7.5, height: 7.5, rx: 1.4 }),
                    h("rect", { x: 13, y: 3.5, width: 7.5, height: 7.5, rx: 1.4 }),
                    h("rect", { x: 3.5, y: 13, width: 7.5, height: 7.5, rx: 1.4 }),
                    h("rect", { x: 13, y: 13, width: 7.5, height: 7.5, rx: 1.4 }));
            }
            return h("svg", common,
                h("path", { d: "M6 3.5h7l5 5v12A1.5 1.5 0 0 1 16.5 22h-10A1.5 1.5 0 0 1 5 20.5v-15A1.5 1.5 0 0 1 6 3.5z" }),
                h("path", { d: "M13 3.8V9h4.7" }));
        }

        /** 块间关系类型：与 src/memory.js 的 LINK_KINDS 同一份词表（服务端会再校验）。 */
        const KB_LINK_KINDS = ["related", "refines", "supersedes", "contradicts", "supports", "derives"];

        /**
         * 列表每页条数（「再显示 N 条」的步长）。
         *
         * 端点上各层的上限是 200 / 200 / 500 / 60（`src/view.js` 的 VIEW_LIMITS），
         * 所以渐进式加载只发生在**已经取回来的那一份**里，不需要给端点加 offset
         * 参数 —— 少一个参数就少一条要守的信任边界。20 这个数取的是「一屏多一点」：
         * 热记忆通常只有几条，台账与关系才需要翻页。
         */
        const LIST_PAGE = 20;
        // 知识库每页可选条数：与服务端 kb-page.js 的 KB_PAGE_MAX(100) 对齐，
        // 这里只列合法档位，避免用户填出会被服务端 clamp 的怪数字。
        const PAGE_SIZES = [20, 50, 100];
        // 摘要块每页档位：默认 10，摘要正文长，一页多会又把页面拉长。
        const SUMMARY_PAGE_SIZES = [10, 20, 50];
        // 统一入口的本地预检（第九十六轮）：选完文件**先看类型与体积**再决定要不要
        // 读成 base64。数值与 src/kb-extract.js 的 8 MB 抽取上限、src/ocr.js 的
        // 24 MB 内容通道上限对齐；扩展名表漂移最多多走一次服务端预览，不会造成越权。
        const KB_PREVIEW_LOCAL_MAX = 8 * 1024 * 1024;
        const KB_OCR_CONTENT_MAX = 24 * 1024 * 1024;
        // 文本类入库的单文档上限（与服务端 kb.js 的 KB_MAX_DOC_BYTES 一致）：预览就拦住。
        // 文本没有「识别文字」这条退路，2–8 MB 的文本若放行到「入库」，点下去才被拒。
        const KB_MAX_TEXT_BYTES = 2 * 1024 * 1024;
        // 二进制文档 / 图片走抽取器的 8 MB；其余可抽取扩展名按文本的 2 MB 算。
        const KB_BINARY_EXTS = ["pdf", "docx", "pptx", "xlsx", "png", "jpg", "jpeg", "webp", "gif"];
        // 与 kb-extract.js 的 TEXT_EXTS / DOC_EXTS / IMAGE_MIME 一致（不带点、小写）。
        const KB_EXTRACTABLE_EXTS = [
            "txt", "md", "markdown", "json", "jsonl", "csv", "tsv", "log", "xml",
            "html", "htm", "css", "js", "mjs", "cjs", "ts", "tsx", "jsx", "py",
            "java", "go", "rs", "c", "h", "cpp", "hpp", "sh", "ps1", "yml", "yaml",
            "toml", "ini", "sql",
            "pdf", "docx", "pptx", "xlsx",
            "png", "jpg", "jpeg", "webp", "gif",
        ];
        // 与 ocr.js 的 OCR_SUPPORTED_EXTENSIONS 一致（不带点、小写）。
        const KB_OCR_EXTS = ["pdf", "doc", "docx", "ppt", "pptx", "png", "jpg", "jpeg", "webp", "gif", "bmp", "tif", "tiff"];
        const kbExtOf = (name) => {
            const text = String(name || "");
            const at = text.lastIndexOf(".");
            return at === -1 ? "" : text.slice(at + 1).toLowerCase();
        };
        const kbOcrEligibleName = (name) => KB_OCR_EXTS.indexOf(kbExtOf(name)) >= 0;

        /**
         * 记忆浏览面板在会话页签行里的 id（`conversation.view` 的一格）。
         *
         * 第六十二轮（新-22）之前，这个字符串是**两个** slot 之间的接缝：
         * `sidebar.panellist` 的 id 与 `main` 的 key，侧栏按钮按「id === key」
         * 去找要点亮的主面板那一格。迁到 `conversation.view` 之后接缝没有了
         * —— 宿主按页签自己的 id 派发这一格，不再有第二个名字要跟上。
         *
         * 邻居的顺序：`chat` 0 / `trajectory` 10 / `context` 20 / `fleet` 30 /
         * **本面板 40**。也就是说它排在「上下文」与「编队」的右侧，这就是
         * 用户要的位置；`order` 是唯一决定这件事的东西，别再按注册先后去推。
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

        /**
         * 面板的写入口（第八十五轮）：POST JSON，动作名与 src/view.js 的
         * handleMemoryMutateRoute 一一对应。它仍然经过同一套 trustedRootFor，
         * 不是「浏览器拿到文件系统」——只是把生成的指令卡换成当场可核对的一次写入。
         */
        const MEMORY_MUTATE_PATH = "/office-memory/mutate";

        /**
         * 记忆导出端点（第七十一轮）：单文件迁移。
         *
         * 为什么是 `<a download>` 直连服务端、而不是 fetch + Blob：服务端已经写了
         * `content-disposition: attachment`，浏览器原生下载不用在面板里存一份完整包
         * （归档大的工作区，导出包可能几十 MB）。面板只负责拼 URL。
         */
        const MEMORY_EXPORT_PATH = "/office-memory/export";

        /**
         * 本地模型状态 / 下载（第八十八轮）：设置页选完知识库目录后，把
         * Qwen3-Embedding-0.6B 与 Qwen3-Reranker-0.6B 拉到 `<目录>/models/`。
         * 状态轮询走 GET，下载是 POST，两者都经过宿主的 trustedRootFor。
         */
        const MODEL_STATUS_PATH = "/office-memory/model-status";
        const MODEL_DOWNLOAD_PATH = "/office-memory/model-download";

        /** 目录浏览与知识库上传（2026-11-19）：不依赖宿主桌面桥，Web 端直接可用。 */
        const FS_LIST_PATH = "/office-memory/fs-list";
        const KB_UPLOAD_PATH = "/office-memory/kb-upload";
        const KB_ASSET_PATH = "/office-memory/kb-asset";
        /** 图库统一入口：GET 浏览/检索/取图，POST 入库/手工标签/描述/向量。 */
        const GALLERY_PATH = "/office-memory/gallery";
        /** 宿主已配置模型清单（只读）：上传时可从真实模型里选，避免手填 id 猜错。 */
        const LLM_MODELS_PATH = "/office-memory/llm-models";
        /** 文件识别（OCR）：POST {path, purpose}，走 DeepSeek 官方 /v1/tools/ocr。 */
        const OCR_PATH = "/office-memory/ocr";

        /** 宿主目录选择器：由 apply 的可选 uiWorkspace 注入，拿不到时为 null。 */
        let HOST_PICK_DIRECTORY = null;

        /** 导出链接：带 cwd（缺省由服务端挑默认根）与 format（json 无损 / md 可读）。 */
        function memoryExportHref(cwd, format) {
            const params = [];
            if (cwd) params.push("cwd=" + encodeURIComponent(cwd));
            params.push("format=" + encodeURIComponent(format));
            return MEMORY_EXPORT_PATH + "?" + params.join("&");
        }

        /**
         * 「办公模式插件」那一格的两条端点（第六十三轮 新-21）。
         *
         * 与记忆端点同一条通道口径（同源相对路径、宿主侧注册）。区别是这里**会写盘**：
         * 写的是办公 preset 的声明文件（真源 + profile 侧已有的覆盖行），不是插件自己的
         * 设置文档 —— 所以它不走 configForms，而是走 src/preset-plugins.js 那两条路由。
         */
        const PRESET_PLUGINS_PATH = "/office-preset/plugins";
        const PRESET_RESTORE_PATH = "/office-preset/restore";

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
        const KB_EXPLORER_COLUMNS = "18px minmax(0, 1fr) 64px 56px 76px 92px 96px";
        // 图库详细信息：缩略图 / 名称 / 类型 / 数量 / 大小 / 加入时间。
        // 与知识库同一套「图标 + 名称 + 元数据列」读法，但不硬套文档的「档位 / 块数」。
        const GALLERY_EXPLORER_COLUMNS = "46px minmax(0, 1fr) 72px 60px 76px 92px";
        const S = {
            section: { padding: "2px 0" },
            // ── 设置页卡片化（第八十六轮）──
            // 原来每组是一条细线目录行；用户给的目标是「像仪表盘卡片那样」。
            // 组本身变成一张有表面色与圆角的卡，组头仍是整行按钮（折叠 / 定位契约不变），
            // 组体单独留内边距 —— data-group-body 的 style 仍然只在收起时是 display:none
            // （test/client.mjs 直接断言了这层 style）。
            group: {
                marginBottom: "10px",
                borderRadius: "14px",
                border: "1px solid " + M.stroke,
                background: M.surface,
                boxShadow: M.soft,
                overflow: "hidden"
            },
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
            // 原生下拉的**弹层**会跟着 select 自己的 background-color 走。这里原来用的是
            // 半透明 hover 底，深色主题下弹层变成一块灰玻璃，选项文字几乎看不见（用户
            // 截图里的预览面板文件下拉）。换成不透明的基础表面色 + 显式前景色：弹层有
            // 实底，文字在两套主题下都读得出来；设置页的 select 共用这一条，一并修好。
            select: Object.assign({
                padding: "6px 10px",
                borderRadius: "8px",
                border: "1px solid var(--dsw-alias-border-l1, rgba(128,128,128,0.3))",
                background: "var(--dsw-alias-bg-base, #fff)",
                color: "var(--dsw-alias-label-primary, #1f2328)",
                minWidth: "220px"
            }, T.label),
            unit: Object.assign({ minWidth: "20px" }, dim(T.note)),
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
            // 分组头：整行可点，左边色点 + 组名 + 计数徽标，右边开合提示。
            // 卡片头部展开时加一条底线，把表头与组体分开；收起时卡片本身就是目录项。
            groupHead: {
                display: "flex",
                alignItems: "center",
                gap: "9px",
                width: "100%",
                padding: "13px 15px",
                background: "none",
                border: "none",
                borderBottom: "1px solid transparent",
                color: "inherit",
                cursor: "pointer",
                textAlign: "left"
            },
            groupHeadOpen: {
                borderBottomColor: M.stroke,
                background: "color-mix(in srgb, " + M.sunken + " 40%, transparent)"
            },
            groupHeadTitle: T.heading,
            groupDot: {
                width: "8px", height: "8px", borderRadius: "3px",
                flex: "0 0 auto", boxShadow: "0 0 0 3px color-mix(in srgb, currentColor 12%, transparent)"
            },
            groupHeadCount: Object.assign({
                marginLeft: "2px", padding: "1px 7px", borderRadius: M.pill,
                background: M.sunken, color: M.textTertiary
            }, T.tiny),
            groupHeadHint: Object.assign({ marginLeft: "auto" }, dim(T.tiny)),
            groupBody: { padding: "2px 15px 8px" },
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
            // ── 设置页仪表盘（第八十六轮）──
            // 一张下沉底的概览卡 + 四枚抬升表面的指标卡。数字用等宽字，
            // 与记忆面板的指标块是同一套读法（用户给的目标页就是这种卡片网格）。
            settingsHero: {
                marginBottom: "16px", padding: "16px 16px 14px",
                borderRadius: M.radiusLg,
                border: "1px solid " + M.stroke,
                background: M.sunken,
                boxShadow: M.soft
            },
            heroKicker: {
                fontFamily: M.mono, fontSize: "9.5px", letterSpacing: "0.14em",
                textTransform: "uppercase", color: M.caption, marginBottom: "4px"
            },
            heroTitle: Object.assign({ margin: "0 0 2px" }, T.title),
            heroNote: Object.assign({ margin: "0 0 14px" }, dim(T.note)),
            heroFacts: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: "8px" },
            heroFact: {
                padding: "11px 12px 10px", borderRadius: "12px",
                border: "1px solid " + M.stroke,
                background: M.surface,
                boxShadow: M.soft
            },
            heroFactTop: { display: "flex", alignItems: "center", gap: "6px", marginBottom: "5px" },
            heroFactDot: { width: "6px", height: "6px", borderRadius: M.pill, flex: "0 0 auto" },
            heroFactLabel: { fontSize: "10.5px", color: M.textTertiary, letterSpacing: "0.03em" },
            heroFactValue: {
                fontFamily: M.mono, fontSize: "20px", fontWeight: 650,
                lineHeight: 1.15, letterSpacing: "-0.02em", color: M.text, ...NUM
            },
            heroFactNote: { marginTop: "3px", fontSize: "10px", lineHeight: 1.45, color: M.caption },
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
            // ── 宽度（第七十一轮重写）──
            //
            // 以前是「写死 maxWidth:1120px + 零横向内边距 + 不居中」：面板贴着宿主
            // 滚动区的左边缘，同屏对照「对话」会宽出一大截、左边距也对不上，看起来
            // 就是「这一格没排版」（用户原话：宽度有问题，可以参考其他页面）。
            //
            // 宿主给会话正文列定的宽度是 `--dsh-chat-content-width`（默认
            // clamp(680px, …, 920px)，见 `.…_body` 的变量定义），输入卡 = 该宽 + 32px、
            // 宽度把手 = ±宽/2 + 24px ⇒ 正文两侧各留 24px。这里照抄这个口径：
            // 内容列宽 = 该变量、外层左右各 24px、整体居中。变量不在时退回 680px
            // （宿主自己的默认值），所以在任何宿主里都不会比「对话」更宽。
            // 顺带保留了原来的初衷：宽屏下不被拉满，一行路径不会拖到屏幕另一头。
            browser: {
                // 宿主 composer 是浮层（高约 118–130px），面板滚到底也要给它让位，
                // 否则最后一张卡、仪表盘和底栏会永远压在输入框下面。
                padding: "4px clamp(18px, 2.6vw, 40px) 148px",
                maxWidth: "min(1380px, calc(var(--dsh-chat-content-width, 680px) + 560px))",
                margin: "0 auto",
                width: "100%",
                color: M.text,
                fontFamily: M.font,
                // 容器查询的锚点：.om-grid-wide 的两栏在面板自己变窄时塌成一列，
                // 而不是等整个窗口变窄（见 PANEL_CSS）。
                containerType: "inline-size",
                // ── 面板**不是**自己的滚动容器：滚动交给宿主的会话滚动容器 ──
                //
                // （第二十轮～第六十二轮初版这里写的是「撑满 + 自己滚」：
                // `height:100%` + `minHeight:0` + `flex:'1 1 auto'` + `overflowY:auto`
                // + `overscrollBehavior:contain`。那套写法在 `main` 那一格是对的，
                // 迁到 `conversation.view` 之后**在真实运行时是坏的**，下面是实测。）
                //
                // 坏在哪（无头 Chrome 实测，第六十二轮修正前）：
                //   · `.om-mem` 实测高 2997px ＝**内容高**（clientHeight == scrollHeight），
                //     自己一格都滚不动；
                //   · 把滚轮打在面板上，面板的 scrollTop 与宿主滚动容器
                //     `.…_scrollBody` 的 scrollTop **都是 0** —— `overscroll-behavior:
                //     contain` 把滚动链**断在面板这一层**（面板是 overflow:auto 的滚动
                //     容器、只是没有可滚的内容），于是整页都滚不动。
                //
                // 为什么 `height:100%` 不再成立 —— 宿主只在**一种**情况下给 viewArea 定高：
                //   `.…_scrollBody:has([data-conversation-composer-overlay])
                //    > [data-slot=conversation\.session] > .…_viewArea{flex:1 1 0;min-height:0;overflow:hidden}`
                // 命中时 viewArea 是 726px 定高、视图内部 `height:100%` 才有效 ——
                // 实测**只有「轨迹」**那个页签命中（它把输入框渲染成 overlay）。
                // 不命中时 `.…_root[data-phase=active] .…_viewArea{flex:1 0 auto;min-height:auto}`
                // 让 viewArea **撑到内容高**，滚动归宿主的 `.…_scrollBody{overflow-y:auto}`。
                // 实测「对话」「上下文」「编队」「记忆」四个页签的 viewArea 高度
                // 分别是 1845 / 1735 / 726 / 2997（＝内容高或视口高，都不是「定高 + 自己滚」），
                // 也就是说**会话视图的通用口径就是「内容照常铺开、由宿主那层滚」**。
                // 所以本面板跟着走：不写高度、不写纵向 overflow、不写 overscroll 阻断。
                //
                // 连带的两件事：
                //   · 横向仍然裁切（长路径 / 长 URL），但**不能用 `overflow-x:hidden`** ——
                //     那会把 overflow-y 也算成 auto，本层又变回滚动容器。`clip` 不创建
                //     滚动容器，正好只做裁切（浏览器不认这条声明时退化成不裁切，纯外观）。
                //   · 头部吸顶照旧：`position:sticky` 现在吸在**宿主滚动容器**的顶边
                //     （见 S.browserHead），搜索框与「刷新 / 记一条」一样一直在视野里。
                overflowX: "clip",
                // 滚到底的那一下不再有「跳一下」：面板里高度会变的部分（「再显示 N 条」
                // 追加条目、预览面板图片晚到）都可能在**滚动已在底部**时改变上方高度，
                // 浏览器默认的滚动锚定会替用户补 scrollTop，表现出来的就是尾部抖一下。
                // 关掉它，位置只听用户的（第七十一轮量测：滚到底后出现过 8px 回跳）。
                overflowAnchor: "none",
                boxSizing: "border-box"
            },
            // 头部随面板一起滚会带来一个副作用：搜索框与「刷新 / 记一条」滑出视野。
            // 所以头部**吸顶** —— 第六十二轮修好滚动之后，吸的是**宿主会话滚动容器**
            // （`.…_scrollBody`）的顶边（面板自己不再是滚动容器，见 S.browser）；
            // 底色取宿主的 base 表面色 —— 面板本来就画在这个底色上，吸顶块不会看出边界。
            //
            // 第七十一轮改间距的写法：原来 `marginBottom:10px` 让吸顶块**下面**留一条
            // 透明缝，滚动时内容从这条缝里穿过去（拉到底时最显眼，看着像闪）。
            // 改成 `paddingBottom:10px`：缝仍在（视觉间距不变），但它落在底色里，
            // 内容滚到吸顶块下沿就被底色盖住，不会先露一条再消失。
            // 头部 + 页签一起吸顶：只吸标题/搜索框还不够 —— 读到长台账的中间想换层，
            // 页签已经滑出去了，只能先把整页滚回顶部。两行绑在同一个 sticky 容器里，
            // 滚到哪一层都能直接换。
            browserSticky: {
                // 标题 / 搜索 / 主动作**不吸顶**：读长清单时把它们让出去，正文从视口顶边开始，
                // 一屏能多看几条。只有页签与筛选留在 sticky 里（见 browserNav），
                // 保证「换层」随时可点，而不用先滚回顶部。
                position: "relative", top: "auto", zIndex: "auto",
                background: "transparent",
                paddingBottom: "0"
            },
            browserNav: {
                position: "sticky", top: 0, zIndex: 3,
                background: "var(--dsw-alias-bg-base, #fff)",
                paddingBottom: "0",
                boxShadow: "0 1px 0 " + M.stroke + ", 0 10px 18px -18px rgba(0,0,0,0.4)"
            },
            // 记忆面板的「工作台」头部：品牌块 / 标题 / 工作区 / 主动作一行，
            // 搜索与导出动作一行，页签与来源筛选压在最下面。三行都跟着宿主滚动吸顶。
            memHead: {
                display: "grid",
                gridTemplateColumns: "auto minmax(0,1fr) auto",
                alignItems: "center",
                gap: "12px",
                padding: "14px 0 10px"
            },
            memBrand: {
                width: "38px", height: "38px", flex: "0 0 auto",
                borderRadius: "12px",
                display: "flex", alignItems: "center", justifyContent: "center",
                color: M.accent,
                background: "linear-gradient(145deg, color-mix(in srgb, " + M.accent + " 13%, " + M.surface + "), " + M.surface + ")",
                border: "1px solid color-mix(in srgb, " + M.accent + " 22%, " + M.stroke + ")",
                boxShadow: "inset 0 1px 0 rgba(255,255,255,0.6), " + M.soft
            },
            memTitleBlock: { minWidth: 0 },
            memKicker: {
                fontFamily: M.mono, fontSize: "9px", lineHeight: 1.2,
                letterSpacing: "0.14em", textTransform: "uppercase",
                color: M.caption, marginBottom: "3px"
            },
            memTitle: { fontSize: "19px", lineHeight: 1.25, fontWeight: 650, letterSpacing: "-0.015em", color: M.text },
            memSubtitle: {
                marginTop: "3px", fontSize: "11.5px", lineHeight: 1.45,
                color: M.textTertiary, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis"
            },
            memHeadActions: { display: "flex", alignItems: "center", gap: "8px", flex: "0 0 auto" },
            memPrimary: {
                display: "inline-flex", alignItems: "center", gap: "5px",
                height: "34px", padding: "0 13px",
                borderRadius: "9px", border: "1px solid " + M.accent,
                background: M.accent,
                // 反色：浅色主题的强调色是深色 → 白字；深色主题的强调色可能是浅色
                // → 必须跟着底色反过来，否则实心按钮上会是一片白字（第八十三轮截图抓到的）。
                color: "var(--dsw-alias-bg-base, #fff)",
                font: "inherit", fontSize: "12.5px", fontWeight: 600,
                cursor: "pointer", whiteSpace: "nowrap",
                boxShadow: "0 1px 2px color-mix(in srgb, " + M.accent + " 35%, transparent)"
            },
            memPlus: { fontSize: "15px", lineHeight: 1, marginTop: "-1px", fontWeight: 400 },
            memControls: { display: "flex", alignItems: "center", gap: "10px", margin: "0 0 10px" },
            memSearchWrap: {
                position: "relative", display: "flex", alignItems: "center",
                flex: "1 1 auto", minWidth: "160px"
            },
            memSearchInput: {
                width: "100%", height: "36px",
                padding: "0 32px 0 32px",
                borderRadius: "10px",
                border: "1px solid " + M.stroke,
                background: M.input,
                color: M.text, font: "inherit", fontSize: "12.5px",
                outline: "none", boxSizing: "border-box",
                boxShadow: "inset 0 1px 2px rgba(0,0,0,0.02)"
            },
            memSearchClear: {
                position: "absolute", right: "6px",
                width: "22px", height: "22px",
                display: "flex", alignItems: "center", justifyContent: "center",
                border: "none", borderRadius: "6px",
                background: "transparent", color: M.caption,
                fontSize: "15px", lineHeight: 1, cursor: "pointer"
            },
            memToolGroup: { display: "flex", alignItems: "center", gap: "1px", flex: "0 0 auto" },
            memTool: {
                display: "inline-flex", alignItems: "center", justifyContent: "center",
                height: "34px", padding: "0 9px",
                borderRadius: "8px", border: "1px solid transparent",
                background: "transparent", color: M.textTertiary,
                font: "inherit", fontSize: "11.5px", textDecoration: "none",
                cursor: "pointer", whiteSpace: "nowrap"
            },
            memTabs: {
                display: "flex", alignItems: "stretch", gap: "2px",
                borderBottom: "1px solid " + M.stroke,
                marginBottom: "0", overflowX: "auto", scrollbarWidth: "none"
            },
            memTab: {
                display: "inline-flex", alignItems: "center", gap: "7px",
                padding: "8px 10px 10px",
                border: "none", borderBottom: "2px solid transparent",
                borderRadius: "0", background: "transparent",
                color: M.textTertiary, font: "inherit", fontSize: "12.5px",
                cursor: "pointer", userSelect: "none", whiteSpace: "nowrap",
                marginBottom: "-1px"
            },
            memTabOn: { color: M.text, fontWeight: 600, background: "transparent" },
            memTabHover: { color: M.textSecondary, background: "transparent" },
            memTabDot: { width: "7px", height: "7px", borderRadius: M.pill, flex: "0 0 auto" },
            memFilters: {
                display: "flex", alignItems: "center", gap: "12px", flexWrap: "wrap",
                padding: "8px 0 9px"
            },
            memFilterGroup: { display: "inline-flex", alignItems: "center", gap: "4px", flexWrap: "wrap" },
            memFilterLabel: {
                fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase",
                color: M.caption, marginRight: "3px"
            },
            memOriginRow: {
                display: "flex", alignItems: "center", gap: "4px", flexWrap: "wrap"
            },
            quickOpen: {
                display: "flex", alignItems: "center", gap: "8px", width: "100%",
                padding: "10px 12px", marginTop: "2px",
                border: "1px dashed " + M.stroke, borderRadius: "9px",
                background: "transparent", color: M.textSecondary,
                font: "inherit", fontSize: "12px", textAlign: "left",
                cursor: "pointer"
            },
            quickOpenHint: { marginLeft: "auto", color: M.caption, fontSize: "11px" },
            memOriginLabel: { fontSize: "10.5px", letterSpacing: "0.04em", color: M.caption, marginRight: "3px" },
            memOriginSeg: {
                display: "inline-flex", alignItems: "center", gap: "5px",
                padding: "3px 9px", borderRadius: "7px",
                border: "1px solid transparent", background: "transparent",
                color: M.textTertiary, font: "inherit", fontSize: "11.5px",
                cursor: "pointer", userSelect: "none"
            },
            memOriginSegOn: { background: M.sunken, borderColor: M.stroke, color: M.text, fontWeight: 600 },
            memOriginSegCount: { fontFamily: M.mono, fontSize: "9.5px", color: M.caption, fontVariantNumeric: "tabular-nums" },
            memOriginNote: { marginLeft: "auto", fontSize: "10.5px", color: M.caption, fontFamily: M.mono },
            memSectionHead: { display: "flex", alignItems: "center", gap: "8px", margin: "4px 0 10px" },
            memSectionDot: { width: "6px", height: "6px", borderRadius: M.pill, flex: "0 0 auto" },
            memSectionTitle: { fontSize: "13.5px", fontWeight: 600, color: M.text, letterSpacing: "-0.01em" },
            memSectionNote: { marginLeft: "auto", fontSize: "10.5px", color: M.caption, fontFamily: M.mono },
            browserHead: {
                display: "flex", alignItems: "flex-start", gap: "12px", flexWrap: "wrap",
                position: "sticky", top: 0, zIndex: 2,
                padding: "4px 0 10px", marginBottom: 0,
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
            // 来源筛选：与页签同为分段控件，但刻意更轻 —— 它是一层过滤，不是第二个导航。
            // 三档的计数来自后端 originCounts（按当前过滤词算，不受来源筛选本身影响）。
            originRow: { display: "flex", alignItems: "center", gap: "4px", flexWrap: "wrap", marginTop: "2px" },
            originLabel: { fontSize: "11px", letterSpacing: "0.06em", color: M.caption, marginRight: "4px" },
            originSeg: {
                display: "inline-flex", alignItems: "center", gap: "5px",
                padding: "4px 10px", borderRadius: M.pill,
                border: "1px solid transparent", background: "transparent",
                color: M.textTertiary, font: "inherit", fontSize: "12px",
                cursor: "pointer", userSelect: "none"
            },
            originSegOn: { background: M.surface, borderColor: M.stroke, color: M.text, fontWeight: 600, boxShadow: M.soft },
            originSegCount: { fontFamily: M.mono, fontSize: "10px", color: M.caption, fontVariantNumeric: "tabular-nums" },
            // 列表**不再自己滚动**（第二十轮）：面板根元素已经是唯一的滚动容器，
            // 再套一层 520px 的内层滚动会出现「滚到中间卡住、要换一个滚动条接着滚」
            // 的双滚动条，而且条目详情与页脚还被压在下面看不见。
            // 分页（每页 20 条 + 再显示 N 条）仍然管「DOM 里有几条」。
            browserList: { display: "flex", flexDirection: "column", gap: "10px" },
            // 知识库的第二层（第八十二轮）：文档清单之外还要能检索块、看块清单、读块正文。
            // 工具条与页签分开：它只在这一格出现，且「文档 / 块检索」是两种看法而不是两层记忆。
            kbBar: {
                display: "flex", alignItems: "center", gap: "9px", flexWrap: "wrap",
                paddingBottom: "12px", marginBottom: "14px", borderBottom: "1px solid " + M.stroke
            },
            kbModeTrack: {
                display: "inline-flex", alignItems: "center", gap: "2px", padding: "3px",
                borderRadius: "9px", background: M.sunken, border: "1px solid " + M.stroke,
            },
            kbMode: {
                padding: "5px 11px", borderRadius: "9px", border: "1px solid transparent",
                background: "transparent", color: M.textTertiary, font: "inherit", fontSize: "12px",
                cursor: "pointer", userSelect: "none",
            },
            kbModeOn: { background: M.surface, borderColor: M.stroke, color: M.text, fontWeight: 600, boxShadow: M.soft },
            kbBack: {
                display: "inline-flex", alignItems: "center", gap: "5px", padding: "5px 10px",
                borderRadius: M.radiusSm, border: "1px solid " + M.strokeStrong, background: M.surface,
                color: M.textSecondary, font: "inherit", fontSize: "12px", cursor: "pointer",
            },
            kbDocHead: {
                display: "flex", alignItems: "center", gap: "12px", flexWrap: "wrap",
                padding: "12px 14px", marginBottom: "12px", borderRadius: "12px",
                background: M.surface, border: "1px solid " + M.stroke,
                boxShadow: M.soft,
            },
            kbDocTitle: { fontSize: "13px", fontWeight: 600, lineHeight: 1.45, overflowWrap: "anywhere" },
            kbDocMeta: { fontFamily: M.mono, fontSize: "10.5px", color: M.caption, lineHeight: 1.6, overflowWrap: "anywhere" },
            kbChunkCardOn: { borderColor: M.accent, boxShadow: "0 0 0 1px " + M.accent },
            kbPre: {
                margin: "8px 0 0", padding: "14px 16px", borderRadius: "12px",
                border: "1px solid " + M.stroke, background: M.sunken,
                color: M.text, fontFamily: M.mono, fontSize: "12px", lineHeight: 1.8,
                whiteSpace: "pre-wrap", wordBreak: "break-word", overflowWrap: "anywhere",
            },
            kbSearchMeta: { display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap", marginBottom: "10px", fontFamily: M.mono, fontSize: "10.5px", color: M.caption },
            // 行内编辑（第八十五轮）：草稿卡与条目卡同宽，保存前不离开列表。
            editor: {
                display: "flex", flexDirection: "column", gap: "8px",
                margin: "0 0 14px", padding: "14px 15px", borderRadius: "13px",
                border: "1px solid " + M.accent, background: M.surface,
                boxShadow: "0 0 0 1px color-mix(in srgb, " + M.accent + " 18%, transparent)",
            },
            editorTitle: { fontSize: "12.5px", fontWeight: 600, color: M.text },
            editorLabel: { display: "flex", flexDirection: "column", gap: "4px", fontSize: "11px", color: M.caption },
            editorInput: {
                width: "100%", boxSizing: "border-box", padding: "7px 10px", borderRadius: "8px",
                border: "1px solid " + M.strokeStrong, background: M.input, color: M.text,
                font: "inherit", fontSize: "12.5px", outline: "none",
            },
            editorSelect: {
                padding: "6px 9px", borderRadius: "8px", border: "1px solid " + M.strokeStrong,
                background: M.surface, color: M.text, font: "inherit", fontSize: "12px",
            },
            editorTextarea: {
                width: "100%", boxSizing: "border-box", minHeight: "84px", resize: "vertical",
                padding: "9px 11px", borderRadius: "9px", border: "1px solid " + M.strokeStrong,
                background: M.input, color: M.text, font: "inherit", fontSize: "12.5px",
                lineHeight: 1.7, outline: "none",
            },
            editorRow: { display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" },
            editorActions: { display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" },
            editorHint: { fontSize: "11.5px", color: M.caption, lineHeight: 1.6 },
            editorError: { fontSize: "11.5px", color: "var(--dsw-alias-state-error-primary, #e5484d)" },
            kbSelectBar: {
                display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap",
                margin: "0 0 10px", padding: "8px 10px", borderRadius: "10px",
                border: "1px dashed " + M.strokeStrong, background: M.sunken,
                fontFamily: M.mono, fontSize: "10.5px", color: M.caption,
            },
            kbRelated: { margin: "8px 0 0", padding: "10px 12px", borderRadius: "10px", border: "1px solid " + M.stroke, background: M.sunken },
            kbRelatedItem: { display: "block", width: "100%", textAlign: "left", border: "none", background: "transparent", color: M.textSecondary, font: "inherit", fontSize: "12px", padding: "4px 0", cursor: "pointer" },
            // 知识库的资源管理器版式（第八十四轮）：左树固定 210px，右列自适应；
            // 面板窄于 620px 时退成上下两块，树限高滚动，避免把正文挤成一条。
            kbExplorer: { display: "flex", alignItems: "flex-start", gap: "18px", marginTop: "4px" },
            // 「详细信息」式的文件清单（第九十一轮）：一列名称 + 块/大小/修改/操作，
            // 文件夹行与文件行共用同一套网格，读起来像资源管理器而不是卡片墙。
            kbExplorerList: {
                display: "flex", flexDirection: "column", marginTop: "8px",
                border: "1px solid " + M.stroke, borderRadius: "12px",
                overflow: "hidden", background: M.surface,
            },
            kbExplorerHead: {
                display: "grid", gridTemplateColumns: KB_EXPLORER_COLUMNS,
                alignItems: "center", gap: "8px", padding: "6px 10px",
                background: M.sunken, borderBottom: "1px solid " + M.stroke,
                fontFamily: M.mono, fontSize: "9.5px", letterSpacing: "0.08em",
                color: M.caption, textTransform: "uppercase",
            },
            kbExplorerRow: {
                display: "grid", gridTemplateColumns: KB_EXPLORER_COLUMNS,
                alignItems: "center", gap: "8px", padding: "6px 10px",
                borderBottom: "1px solid " + M.stroke,
                background: "transparent", font: "inherit", fontSize: "12.5px",
                color: M.text, textAlign: "left",
            },
            kbExplorerName: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: M.text, fontWeight: 500 },
            kbExplorerNameBtn: {
                border: "none", background: "transparent", padding: 0, margin: 0,
                font: "inherit", fontWeight: 500, color: M.text, textAlign: "left",
                cursor: "pointer", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
            },
            kbExplorerMeta: { fontFamily: M.mono, fontSize: "11px", color: M.caption, whiteSpace: "nowrap", ...NUM },
            kbTree: {
                flex: "0 0 210px", width: "210px", maxWidth: "100%",
                display: "flex", flexDirection: "column", gap: "1px",
                padding: "8px 6px", borderRadius: "12px",
                border: "1px solid " + M.stroke, background: M.sunken,
            },
            kbTreeHead: {
                padding: "4px 8px 8px", fontFamily: M.mono, fontSize: "9.5px",
                letterSpacing: "0.08em", color: M.caption, textTransform: "uppercase",
            },
            kbTreeRow: {
                display: "flex", alignItems: "center", gap: "7px", width: "100%",
                padding: "5px 8px", borderRadius: "7px", border: "1px solid transparent",
                background: "transparent", color: M.textSecondary, font: "inherit",
                fontSize: "12px", textAlign: "left", cursor: "pointer", overflow: "hidden",
            },
            kbTreeRowOn: { background: M.surface, borderColor: M.stroke, color: M.text, fontWeight: 600, boxShadow: M.soft },
            kbTreeGlyph: { flex: "0 0 auto", width: "12px", fontSize: "10px", color: M.caption, textAlign: "center" },
            kbTreeName: { flex: "1 1 auto", minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
            kbTreeCount: { flex: "0 0 auto", fontFamily: M.mono, fontSize: "10px", color: M.caption, fontVariantNumeric: "tabular-nums" },
            kbTreeEmpty: { padding: "8px", fontSize: "11.5px", color: M.caption },
            kbMain: { flex: "1 1 auto", minWidth: 0 },
            kbCrumb: {
                display: "flex", alignItems: "center", gap: "1px", flexWrap: "wrap",
                marginBottom: "10px", minHeight: "26px",
            },
            kbCrumbBtn: {
                padding: "3px 5px", borderRadius: "5px", border: "1px solid transparent",
                background: "transparent", color: M.textTertiary, font: "inherit",
                fontSize: "12px", cursor: "pointer", whiteSpace: "nowrap",
            },
            kbCrumbSep: { color: M.caption, fontSize: "11px", padding: "0 1px" },
            kbRecursive: {
                marginLeft: "auto", padding: "3px 8px", borderRadius: M.pill,
                border: "1px solid " + M.stroke, background: "transparent",
                color: M.caption, font: "inherit", fontSize: "10.5px", cursor: "pointer",
            },
            kbIngestBtn: {
                marginLeft: "4px", padding: "3px 9px", borderRadius: M.pill,
                border: "1px solid " + M.strokeStrong, background: M.surface,
                color: M.textSecondary, font: "inherit", fontSize: "11px", cursor: "pointer",
            },
            // ── 资源管理器（Windows 11 口径）──
            // 顶部命令行：后退 / 上一级 / 刷新 / 视图 / 排序都放这里，按钮保持紧凑。
            explorerBar: {
                display: "flex", alignItems: "center", gap: "6px", flexWrap: "wrap",
                padding: "6px 8px", marginBottom: "8px", borderRadius: "10px",
                border: "1px solid " + M.stroke, background: M.surface,
            },
            explorerTool: {
                display: "inline-flex", alignItems: "center", justifyContent: "center", gap: "5px",
                minWidth: "28px", height: "28px", padding: "0 8px", borderRadius: "7px",
                border: "1px solid transparent", background: "transparent", color: M.textSecondary,
                font: "inherit", fontSize: "12px", cursor: "pointer",
            },
            explorerToolOn: { background: M.sunken, borderColor: M.stroke, color: M.text },
            explorerToolDisabled: { opacity: 0.45, cursor: "default" },
            explorerSort: {
                height: "28px", padding: "0 6px", borderRadius: "7px",
                border: "1px solid " + M.stroke, background: M.input, color: M.text,
                font: "inherit", fontSize: "11.5px", cursor: "pointer",
            },
            explorerStatus: {
                display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap",
                marginTop: "8px", padding: "5px 9px", borderRadius: "7px",
                borderTop: "1px solid " + M.stroke, color: M.caption,
                fontFamily: M.mono, fontSize: "10.5px",
            },
            explorerGrid: {
                display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(132px, 1fr))",
                gap: "10px", marginTop: "8px",
            },
            explorerTile: {
                display: "flex", flexDirection: "column", alignItems: "stretch", gap: "6px",
                minWidth: 0, padding: "9px", borderRadius: "10px",
                border: "1px solid " + M.stroke, background: M.surface,
                color: M.text, font: "inherit", textAlign: "left", cursor: "pointer",
            },
            explorerTileOn: { borderColor: M.accent, boxShadow: "0 0 0 1px " + M.accent },
            explorerTileIcon: { display: "flex", alignItems: "center", justifyContent: "center", height: "58px", color: M.textTertiary },
            explorerTileName: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: "12px", fontWeight: 500 },
            explorerTileMeta: { fontFamily: M.mono, fontSize: "10px", color: M.caption },
            gallerySide: {
                flex: "0 0 210px", width: "210px", maxWidth: "100%",
                display: "flex", flexDirection: "column", gap: "1px",
                padding: "8px 6px", borderRadius: "12px",
                border: "1px solid " + M.stroke, background: M.sunken,
            },
            galleryNav: {
                display: "flex", alignItems: "center", gap: "7px", width: "100%",
                padding: "6px 8px", borderRadius: "7px", border: "1px solid transparent",
                background: "transparent", color: M.textSecondary, font: "inherit",
                fontSize: "12px", textAlign: "left", cursor: "pointer",
            },
            galleryNavOn: { background: M.surface, borderColor: M.stroke, color: M.text, fontWeight: 600 },
            galleryGrid: {
                display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(154px, 1fr))",
                gap: "10px", marginTop: "8px",
            },
            galleryCard: {
                display: "flex", flexDirection: "column", minWidth: 0, padding: 0,
                borderRadius: "10px", overflow: "hidden", border: "1px solid " + M.stroke,
                background: M.surface, color: M.text, font: "inherit", textAlign: "left", cursor: "pointer",
            },
            galleryCardOn: { borderColor: M.accent, boxShadow: "0 0 0 1px " + M.accent },
            galleryThumb: { display: "block", width: "100%", height: "116px", objectFit: "cover", background: M.sunken },
            galleryCardBody: { display: "flex", flexDirection: "column", gap: "3px", padding: "7px 8px" },
            galleryDetail: {
                display: "flex", gap: "14px", flexWrap: "wrap", alignItems: "flex-start",
                marginBottom: "10px", padding: "10px", borderRadius: "12px",
                border: "1px solid " + M.stroke, background: M.surface,
            },
            galleryPreview: {
                flex: "0 0 300px", maxWidth: "100%", borderRadius: "10px",
                border: "1px solid " + M.stroke, background: M.sunken,
            },
            // 图库的「列表」视图：缩略图 + 编号 + 尺寸/时间，一行一张，和资源管理器的
            // 详细信息同一种读法；列是固定的，不用 KB 的文档列（这里没有标题/块数）。
            galleryRow: {
                display: "grid", gridTemplateColumns: GALLERY_EXPLORER_COLUMNS,
                alignItems: "center", gap: "9px", padding: "5px 9px",
                borderBottom: "1px solid " + M.stroke, background: "transparent",
                color: M.text, font: "inherit", fontSize: "12px", textAlign: "left",
                width: "100%", cursor: "pointer",
            },
            galleryRowThumb: {
                display: "block", width: "42px", height: "42px", objectFit: "cover",
                borderRadius: "6px", border: "1px solid " + M.stroke, background: M.sunken,
            },
            galleryRowMeta: { fontFamily: M.mono, fontSize: "10.5px", color: M.caption, whiteSpace: "nowrap" },
            // 能力区：每行一条「通道 → 可用性 / 原因」，未配置时照实写，不画假按钮。
            galleryCap: {
                display: "grid", gridTemplateColumns: "88px minmax(0, 1fr)", gap: "8px",
                alignItems: "baseline", padding: "4px 0", borderBottom: "1px dashed " + M.stroke,
                fontSize: "11.5px", color: M.textSecondary,
            },
            galleryCapLabel: { fontFamily: M.mono, fontSize: "10px", letterSpacing: "0.06em", color: M.caption, textTransform: "uppercase" },
            // 原图查看层：盖在面板内容区上，点关闭 / 背景 / Escape 都能退出来。
            galleryOverlay: {
                position: "fixed", inset: 0, zIndex: 60, display: "flex",
                alignItems: "center", justifyContent: "center", padding: "24px",
                background: "color-mix(in srgb, #000 72%, transparent)",
            },
            galleryOverlayImg: {
                maxWidth: "92vw", maxHeight: "86vh", borderRadius: "8px",
                boxShadow: "0 18px 48px rgba(0,0,0,.45)", background: M.sunken,
            },
            // 统计区：内容优先 —— 列表先出来，量化视图放到列表之后。否则打开「记忆」
            // 先看到五张指标卡 + 容量 + 存储域，热记忆本身被推到首屏之外（活体截图里
            // 802px 高的窗口连一条都看不到）。这里用一条分隔线把「看完内容再看账」
            // 讲清楚，而不是让统计卡与条目抢同一个视觉起点。
            overview: {
                marginTop: "30px", padding: "15px 16px 16px",
                borderRadius: M.radiusLg,
                border: "1px solid " + M.stroke,
                background: "color-mix(in srgb, " + M.sunken + " 52%, transparent)"
            },
            dashKicker: {
                fontFamily: M.mono, fontSize: "9.5px", letterSpacing: "0.14em",
                textTransform: "uppercase", color: M.caption, marginBottom: "4px"
            },
            overviewTitle: {
                display: "flex", alignItems: "baseline", gap: "8px", flexWrap: "wrap",
                marginBottom: "12px",
                fontSize: "14px", fontWeight: 600,
                letterSpacing: 0, color: M.textSecondary,
                textTransform: "none"
            },
            overviewNote: { fontFamily: M.mono, fontSize: "10px", letterSpacing: 0, color: M.caption, marginLeft: "auto" },
            browserItem: {
                padding: "13px 15px 12px",
                borderRadius: "13px",
                border: "1px solid color-mix(in srgb, " + M.stroke + " 82%, transparent)",
                background: M.surface,
                boxShadow: "0 1px 1px rgba(0,0,0,0.015)",
                fontSize: "13px",
                lineHeight: 1.7,
                overflowWrap: "anywhere",
                wordBreak: "break-word"
            },
            browserItemHover: { borderColor: M.strokeStrong, boxShadow: M.lift },
            browserMeta: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "10px", marginBottom: "7px", fontSize: "10.5px" },
            badge: {
                padding: 0, borderRadius: 0, border: "1px solid transparent",
                background: "transparent", color: M.caption,
                whiteSpace: "nowrap", letterSpacing: "0.01em"
            },
            badgeOrigin: {
                padding: "1px 7px", borderRadius: M.pill,
                border: "1px solid " + M.stroke, background: M.sunken,
                color: M.textTertiary, fontSize: "10px"
            },
            badgeGlobal: {
                borderColor: "color-mix(in srgb, " + M.business + " 30%, " + M.stroke + ")",
                background: "color-mix(in srgb, " + M.business + " 8%, transparent)",
                color: M.business
            },
            badgeCritical: {
                padding: "1px 7px", borderRadius: M.pill,
                border: "1px solid color-mix(in srgb, " + M.error + " 28%, " + M.stroke + ")",
                background: "color-mix(in srgb, " + M.error + " 7%, transparent)",
                color: M.error
            },
            // 正文限高 + pre-wrap：换行照原样保留，但不会把一条记忆拉成一屏。
            // 限高只是**折叠**不是截断：卡片上的「展开」把 maxHeight 去掉，完整内容
            // 就地铺开。原来这里挂的是原生 title，系统气泡既不可样式化、也没法键盘
            // 触发，而且它出现的位置由浏览器决定、常常盖住相邻条目 —— 现在完整内容
            // 只有两个出口：点「展开」（键盘可达），或看面板底部那条条目详情。
            browserBody: {
                whiteSpace: "pre-wrap", maxHeight: "8.4em", overflow: "hidden",
                color: M.text, fontSize: "13.5px", lineHeight: 1.75
            },
            browserBodyOpen: { whiteSpace: "pre-wrap", color: M.text, fontSize: "13.5px", lineHeight: 1.75 },
            /** 条目底部那排小动作：展开 / 改 / 忘 / 删关系 …。 */
            itemActions: { display: "flex", flexWrap: "wrap", gap: "1px", marginTop: "9px", marginLeft: "-8px" },
            itemAction: {
                display: "inline-flex",
                alignItems: "center",
                gap: "4px",
                padding: "4px 8px",
                borderRadius: "6px",
                border: "1px solid transparent",
                background: "transparent",
                color: M.caption,
                font: "inherit",
                fontSize: "11.5px",
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
            /** 归档隔离提示：只在归档视图出现，纯文本，不含路径或未核实内容。 */
            archiveScopeWarning: {
                margin: "8px 0 12px",
                padding: "9px 11px",
                border: "1px solid color-mix(in srgb, " + M.warn + " 42%, " + M.stroke + ")",
                background: "color-mix(in srgb, " + M.warn + " 8%, transparent)",
                borderRadius: M.radiusSm,
                color: M.text,
                fontSize: "11.5px",
                lineHeight: 1.6,
            },
            archiveScopeWarningTitle: { color: M.warn, fontWeight: 600, marginBottom: "2px" },
            archiveScopeWarningList: { margin: 0, paddingLeft: "16px" },
            /** 条目详情条：列表里不做浮动气泡，详情固定占一条。 */
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

            // ── 可视化：指标条 / 容量条 / 存储卡 / 时间线 ──
            //
            // 这里全部用内联样式而不是运行时注入 CSS：bundle 在测试与 SSR 环境里
            // 是没有 document 的（test/client.mjs 用 new Function 把它跑起来，
            // 只给 window / require / console / fetch）。一旦在模块顶层碰 document，
            // 整个 bundle 会当场抛错、所有用例一起红。
            //
            // 需要伪类与关键帧的地方（悬停、焦点、转圈、窄屏塌列）走面板里的
            // <style> 节点（见 PANEL_CSS）—— 那是 React 树里的普通节点，不是
            // 模块顶层副作用，两个环境里都安全。
            // 指标卡改成仪表盘网格（第八十六轮）：190px 起跳，标准 728px 宽是三列
            // （六枚正好两行），窄屏自动落成两列。仍保留 .om-card 与抬升表面。
            tiles: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: "10px", marginBottom: "14px" },
            tile: {
                position: "relative",
                display: "flex",
                flexDirection: "column",
                gap: "4px",
                padding: "14px 15px 12px",
                borderRadius: M.radiusMd,
                border: "1px solid " + M.stroke,
                background: M.surface,
                boxShadow: M.soft,
                overflow: "hidden"
            },
            tileAccent: { position: "absolute", top: 0, left: 0, right: 0, height: "3px" },
            tileHead: { display: "flex", alignItems: "center", gap: "7px" },
            tileDot: { width: "7px", height: "7px", borderRadius: "3px", flex: "0 0 auto", marginLeft: "auto" },
            tileLabel: { fontSize: "11px", color: M.textTertiary, fontWeight: 500, letterSpacing: "0.04em" },
            // 指标数字用等宽字：并排三枚时数字对得齐，也带出 mnemon 那点技术感。
            tileValue: { fontSize: "28px", fontWeight: 650, lineHeight: 1.15, letterSpacing: "-0.02em", fontFamily: M.mono, color: M.text, ...NUM },
            tileNote: { fontSize: "10.5px", color: M.caption, lineHeight: 1.45 },

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
            hotGroupTitle: {
                display: "flex", alignItems: "center", gap: "8px",
                fontSize: "11px", fontWeight: 600, color: M.textTertiary,
                letterSpacing: "0.02em", margin: "2px 0 8px"
            },

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

        /**
         * 开关。
         *
         * attrs：需要给原生 input 挂 data-* 标记时用（只读检查据此确认某个开关
         * 存在，不必点击、也不去读已存的配置）。顺序上放最后，允许调用方覆盖。
         */
        function Toggle(props) {
            return h("input", Object.assign({
                type: "checkbox",
                checked: props.checked === true,
                disabled: props.disabled === true,
                onChange: (event) => props.onChange(event.target.checked),
                style: { width: "18px", height: "18px", cursor: props.disabled ? "default" : "pointer" },
            }, props.attrs || {}));
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
            // data-* 原样透传：设置页的自动化测试靠它定位下拉，不再按选项文案猜。
            const markers = {};
            for (const key of Object.keys(props)) {
                if (key.startsWith("data-")) markers[key] = props[key];
            }
            return h("select", Object.assign({
                value: props.value,
                disabled: props.disabled === true,
                onChange: (event) => props.onChange(event.target.value),
                style: S.select,
            }, markers), props.options.map((option) =>
                h("option", { key: option.value, value: option.value }, option.label),
            ));
        }

        /**
         * 目录浏览器（2026-11-19）：插件自己的服务端目录列表 + 前端弹层。
         *
         * 为什么不再只依赖 `uiWorkspace.pickDirectory()`：它在本机 Web 部署里依赖宿主的
         * 桌面桥 / 远程目录选择 RPC，用户点了没反应（宿主模板自己的目录选择 flow 才注册了
         * 那条通道）。知识库目录必须能选，所以这里用插件自己的只读 `fs-list` 端点：
         * 只列目录、不读文件内容、仍经过同一套 trustedRootFor。
         */
        function DirBrowser(props) {
            const mode = props.mode === "file" ? "file" : "dir";
            const [current, setCurrent] = React.useState(String(props.start || ""));
            const [picked, setPicked] = React.useState("");
            const [state, setState] = React.useState({ loading: true, error: "", parent: "", roots: [], entries: [] });
            const load = async (target) => {
                setState({ loading: true, error: "", parent: "", roots: [], entries: [] });
                setPicked("");
                try {
                    const response = await fetch(FS_LIST_PATH + "?cwd=" + encodeURIComponent(props.cwd || "")
                        + "&path=" + encodeURIComponent(target || ""), { headers: { accept: "application/json" } });
                    const body = await response.json();
                    if (!response.ok || body.ok !== true) throw new Error(body.error || ("HTTP " + response.status));
                    setCurrent(String(body.path || ""));
                    setState({ loading: false, error: "", parent: String(body.parent || ""), roots: Array.isArray(body.roots) ? body.roots : [], entries: Array.isArray(body.entries) ? body.entries : [] });
                } catch (error) {
                    setState({ loading: false, error: String(error && error.message ? error.message : error), parent: "", roots: [], entries: [] });
                }
            };
            React.useEffect(() => { load(String(props.start || "")); }, []);
            const overlay = {
                position: "fixed", inset: 0, zIndex: 9999, background: "rgba(0,0,0,.45)",
                display: "flex", alignItems: "center", justifyContent: "center", padding: "24px",
            };
            const panel = {
                width: "min(760px, 100%)", maxHeight: "min(680px, 86vh)", display: "flex", flexDirection: "column",
                background: "var(--dsw-alias-bg-base, #fff)", color: "var(--dsw-alias-label-primary, #1f2328)",
                border: "1px solid " + M.stroke, borderRadius: "14px", boxShadow: "0 24px 80px rgba(0,0,0,.35)", overflow: "hidden",
            };
            const rowStyle = { display: "flex", alignItems: "center", gap: "8px", padding: "8px 12px", borderBottom: "1px solid " + M.stroke };
            const dirs = state.entries.filter((entry) => entry && entry.kind !== "file");
            const files = mode === "file" ? state.entries.filter((entry) => entry && entry.kind === "file") : [];
            const rows = [];
            for (const entry of dirs) {
                rows.push(h("button", {
                    key: "d:" + entry.path, type: "button", className: "om-btn",
                    "data-dir-entry": entry.path,
                    onDoubleClick: () => load(entry.path),
                    onClick: () => load(entry.path),
                    style: {
                        display: "flex", alignItems: "center", gap: "9px", width: "100%", textAlign: "left",
                        padding: "9px 14px", border: 0, background: "transparent", cursor: "pointer",
                        color: "inherit", borderBottom: "1px solid " + M.stroke,
                    },
                }, h("span", { style: { opacity: .7 } }, "▸"), h("span", null, entry.name)));
            }
            for (const entry of files) {
                const on = picked === entry.path;
                rows.push(h("button", {
                    key: "f:" + entry.path, type: "button", className: "om-btn",
                    "data-dir-file": entry.path,
                    onClick: () => setPicked(entry.path),
                    onDoubleClick: () => { props.onPick(entry.path); props.onClose(); },
                    style: {
                        display: "flex", alignItems: "center", gap: "9px", width: "100%", textAlign: "left",
                        padding: "9px 14px", border: 0, cursor: "pointer", color: "inherit",
                        background: on ? "rgba(61,155,191,.14)" : "transparent",
                        borderBottom: "1px solid " + M.stroke,
                    },
                },
                    h("span", { style: { opacity: .7 } }, "▤"),
                    h("span", { style: { flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis" } }, entry.name),
                    h("span", { style: Object.assign({}, dim(T.note), { whiteSpace: "nowrap" }) }, humanBytes(entry.size))));
            }
            const emptyText = mode === "file"
                ? (dirs.length === 0 ? "这个目录下没有子文件夹，也没有可选的文件。" : "这个目录下没有文件，点开子文件夹继续找。")
                : "这个目录下没有子文件夹。直接点「用这个目录」即可。";
            const pickTarget = mode === "file" ? picked : current;
            return h("div", { style: overlay, "data-dir-browser": "true", "data-dir-mode": mode, onClick: (event) => { if (event.target === event.currentTarget) props.onClose(); } },
                h("div", { style: panel },
                    h("div", { style: Object.assign({}, rowStyle, { justifyContent: "space-between" }) },
                        h("div", { style: { fontWeight: 600 } }, mode === "file" ? "选择文件（只识别工作目录或知识库目录里的文件）" : "选择知识库目录"),
                        h("button", { type: "button", className: "om-btn", style: S.ghostButton, onClick: props.onClose }, "关闭")),
                    h("div", { style: rowStyle },
                        h("button", { type: "button", className: "om-btn", style: S.ghostButton, "data-dir-up": "true", disabled: state.parent === "", onClick: () => load(state.parent) }, "↑ 上一级"),
                        h("input", { type: "text", value: current, readOnly: true, "data-dir-path": "true", style: Object.assign({}, S.input, { flex: 1, minWidth: 0, textAlign: "left" }) }),
                        h("button", { type: "button", className: "om-btn", style: S.ghostButton, onClick: () => load(current) }, "刷新")),
                    state.roots.length > 1
                        ? h("div", { style: Object.assign({}, rowStyle, { flexWrap: "wrap" }) },
                            state.roots.map((root) => h("button", { key: root, type: "button", className: "om-btn", style: S.ghostButton, "data-dir-root": root, onClick: () => load(root) }, root)))
                        : null,
                    h("div", { style: { flex: 1, overflow: "auto", minHeight: "240px" } },
                        state.loading
                            ? h("div", { style: Object.assign({}, dim(T.note), { padding: "18px" }) }, "正在读取目录…")
                            : state.error
                                ? h("div", { style: { padding: "18px", color: "#d4564c" }, "data-dir-error": "true" }, state.error)
                                : rows.length === 0
                                    ? h("div", { style: Object.assign({}, dim(T.note), { padding: "18px" }) }, emptyText)
                                    : rows),
                    h("div", { style: Object.assign({}, rowStyle, { borderBottom: 0, justifyContent: "space-between" }) },
                        h("div", { style: Object.assign({}, dim(T.note), { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: "430px" }) }, pickTarget || current || "—"),
                        h("div", { style: { display: "flex", gap: "8px" } },
                            mode === "file" || HOST_PICK_DIRECTORY === null ? null : h("button", {
                                type: "button", className: "om-btn", style: S.ghostButton,
                                onClick: async () => {
                                    try {
                                        const pickedDir = await HOST_PICK_DIRECTORY();
                                        const dir = typeof pickedDir === "string" ? pickedDir : (pickedDir && typeof pickedDir.path === "string" ? pickedDir.path : "");
                                        if (dir) { props.onPick(dir); props.onClose(); }
                                    } catch (error) {
                                        setState(Object.assign({}, state, { error: "系统对话框没有选到目录：" + String(error && error.message ? error.message : error) }));
                                    }
                                },
                            }, "系统对话框"),
                            h("button", {
                                type: "button", className: "om-btn om-primary", style: S.memPrimary, "data-dir-select": "true",
                                disabled: pickTarget === "",
                                onClick: () => { if (pickTarget === "") return; props.onPick(pickTarget); props.onClose(); },
                            }, mode === "file" ? "用这个文件" : "用这个目录"))),
                ));
        }
        /**
         * 知识库目录行：文本框 + 目录浏览器。留空 = $DSH_HOME/.office/kb。
         */
        function KbDirRow(props) {
            // 同一行既管知识库目录，也管图库目录：两者都是「留空 = 默认、相对路径相对
            // 用户主目录」的路径设置，行为、保存回执与目录浏览器完全一样，不为图库另抄一份。
            const settingKey = String(props.settingKey || "memory.kbDir");
            const inputName = String(props.inputName || "data-kb-dir-input");
            const pickName = String(props.pickName || "data-kb-dir-pick");
            const blankText = String(props.blankText || "默认目录");
            const [draft, setDraft] = React.useState(String(props.value || ""));
            const [note, setNote] = React.useState("");
            const [saving, setSaving] = React.useState(false);
            const [browsing, setBrowsing] = React.useState(false);
            React.useEffect(() => { setDraft(String(props.value || "")); }, [props.value]);
            // 等 host 真正确认写入后再说「已保存」。write 失败时把错误留在这一行，
            // 而不是只写进设置页顶部的全局错误（用户可能已经滚过去了）。
            const save = (next) => {
                setSaving(true);
                setNote("正在保存…");
                return Promise.resolve(props.write(settingKey, next)).then((result) => {
                    setSaving(false);
                    if (result && result.ok === false) setNote("保存失败：" + String(result.error || "设置服务拒绝了这次写入。"));
                    else setNote("已保存：" + (next === "" ? blankText : next));
                }).catch((error) => {
                    setSaving(false);
                    setNote("保存失败：" + String(error && error.message ? error.message : error));
                });
            };
            const commit = (event) => {
                // 给测试/自动化留一条不依赖 React 状态刷新就能提交的路：事件里带了值就用它。
                const shown = event && event.target && typeof event.target.value === "string" ? event.target.value : draft;
                const next = String(shown).trim();
                if (next !== String(props.value || "")) save(next);
            };
            const pick = (dir) => {
                setDraft(dir);
                save(dir);
            };
            return h(Row, {
                label: String(props.label || "知识库目录"),
                note: String(props.note || "一个知识库，不再分全局库 / 工作区库。相对路径相对用户主目录。"),
                first: props.first !== false,
            },
                h("div", { style: { display: "flex", flexDirection: "column", gap: "5px", alignItems: "flex-end" } },
                    h("div", { style: { display: "flex", gap: "8px", alignItems: "center" } },
                        h("input", {
                            type: "text",
                            value: draft,
                            [inputName]: "true",
                            disabled: props.disabled === true || saving,
                            placeholder: String(props.placeholder || "留空 = $DSH_HOME/.office/kb"),
                            onChange: (event) => setDraft(event.target.value),
                            onBlur: commit,
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "300px", textAlign: "left" }),
                        }),
                        h("button", {
                            type: "button",
                            className: "om-btn",
                            [pickName]: "true",
                            style: Object.assign({}, S.ghostButton, { whiteSpace: "nowrap" }),
                            disabled: props.disabled === true || saving,
                            onClick: () => setBrowsing(true),
                        }, "选择目录"),
                    ),
                    note ? h("div", { style: Object.assign({}, dim(T.note), { maxWidth: "420px", textAlign: "right" }) }, note) : null,
                ),
                browsing ? h(DirBrowser, {
                    start: String(props.value || ""),
                    onPick: pick,
                    onClose: () => setBrowsing(false),
                }) : null,
            );
        }

        function humanModelBytes(value) {
            const n = Number(value) || 0;
            if (n <= 0) return "—";
            if (n >= 1024 * 1024 * 1024) return (n / (1024 * 1024 * 1024)).toFixed(2) + " GB";
            if (n >= 1024 * 1024) return (n / (1024 * 1024)).toFixed(1) + " MB";
            return Math.round(n / 1024) + " KB";
        }

        /**
         * 本地模型下载区：bge-small-zh-v1.5（本机推理用，512 维，实测峰值 RSS 约 480–515MB）。
         *
         * 下载是后台任务（大模型按 GB 计），这里只轮询状态；进度条按已下载字节 /
         * 仓库总字节算，下载器拿不到总大小时退化成「正在下载 <文件名>」。
         */
        function KbModelsPanel(props) {
            const [status, setStatus] = React.useState(null);
            const [error, setError] = React.useState("");
            const [starting, setStarting] = React.useState("");
            const load = async () => {
                try {
                    const response = await fetch(MODEL_STATUS_PATH + "?cwd=" + encodeURIComponent(props.cwd || ""), { headers: { accept: "application/json" } });
                    const body = await response.json();
                    if (!response.ok || body.ok !== true) throw new Error(body.error || ("HTTP " + response.status));
                    setStatus(body);
                    setError("");
                } catch (err) {
                    setError(String(err && err.message ? err.message : err));
                }
            };
            React.useEffect(() => {
                let alive = true;
                const tick = async () => { if (alive) await load(); };
                tick();
                // poll=false：只读一次。本地推理没开时没有下载/加载中的状态可追，
                // 每 2 秒打一次状态接口只是白耗请求。
                const timer = props.poll === false ? null : setInterval(tick, 2000);
                return () => { alive = false; if (timer !== null) clearInterval(timer); };
            }, [props.cwd, props.kbDir, props.poll]);
            const start = async (kind) => {
                setStarting(kind);
                setError("");
                try {
                    const response = await fetch(MODEL_DOWNLOAD_PATH + "?cwd=" + encodeURIComponent(props.cwd || ""), {
                        method: "POST",
                        headers: { "content-type": "application/json" },
                        body: JSON.stringify({ kind: kind }),
                    });
                    const body = await response.json();
                    if (!response.ok || body.ok !== true) throw new Error(body.error || ("HTTP " + response.status));
                    await load();
                } catch (err) {
                    setError(String(err && err.message ? err.message : err));
                } finally {
                    setStarting("");
                }
            };
            const models = status && Array.isArray(status.models) ? status.models : [];
            return h("div", { style: { padding: "12px 0 4px", borderTop: "1px solid " + M.stroke }, "data-kb-models": "true" },
                h("div", { style: { display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: "12px", marginBottom: "8px" } },
                    h("div", null,
                        h("div", { style: Object.assign({}, T.body, { fontWeight: 600 }) }, "本地模型"),
                        h("div", { style: dim(T.note) }, "下载到知识库目录下的 models/。优先魔搭 ModelScope（国内直连），失败自动退到 HF 镜像；只放文件，推理由本机 Python 服务按需启动（设置页下方「本地推理服务」）。"),
                    ),
                    h("button", { type: "button", className: "om-btn", style: S.ghostButton, onClick: load }, "刷新状态"),
                ),
                error ? h("div", { style: Object.assign({}, dim(T.note), { color: "#d4564c" }) }, error) : null,
                h("div", { style: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "10px" } },
                    models.length === 0
                        ? h("div", { style: dim(T.note) }, "正在读取模型状态…")
                        : models.map((model) => {
                            const state = String(model.state || "missing");
                            const ready = state === "ready";
                            const running = state === "running";
                            const progress = Math.max(0, Math.min(1, Number(model.progress) || 0));
                            const total = Number(model.totalBytes) || 0;
                            const done = Number(model.downloadedBytes) || 0;
                            return h("div", {
                                key: model.kind,
                                "data-model-kind": model.kind,
                                "data-model-state": state,
                                style: {
                                    border: "1px solid " + M.stroke,
                                    borderRadius: "12px",
                                    padding: "11px 12px",
                                    background: M.surface,
                                    display: "flex",
                                    flexDirection: "column",
                                    gap: "7px",
                                },
                            },
                                h("div", { style: { display: "flex", justifyContent: "space-between", gap: "8px", alignItems: "baseline" } },
                                    h("div", { style: Object.assign({}, T.body, { fontWeight: 600 }) }, model.label || model.kind),
                                    h("span", { style: dim(T.note) }, ready ? "已就绪" : (running ? "下载中" : (state === "error" ? "失败" : "未下载"))),
                                ),
                                h("div", { style: dim(T.note), "data-model-source": String(model.sourceLabel || model.source || "") },
                                    ready
                                        ? (humanModelBytes(model.bytes) + " · " + (model.files || 0) + " 个文件" + (model.sourceLabel ? " · " + model.sourceLabel : ""))
                                        : running
                                            ? (humanModelBytes(done) + " / " + humanModelBytes(total) + (model.file ? " · " + model.file : "") + (model.sourceLabel ? " · " + model.sourceLabel : ""))
                                            : "下载后可用于本地向量推理；不下载不影响词法检索，也不影响远端 embedding。"),
                                state === "missing" && status && Array.isArray(status.sources) && status.sources.length > 0
                                    ? h("div", { style: dim(T.note) }, "来源顺序：" + status.sources.map((source) => source.label).join(" → "))
                                    : null,
                                running ? h("div", { style: { height: "5px", borderRadius: "999px", background: M.stroke, overflow: "hidden" } },
                                    h("div", { style: { height: "100%", width: Math.round(progress * 100) + "%", background: "var(--dsw-alias-label-primary, #1f2328)", transition: "width .2s ease" } }),
                                ) : null,
                                model.error ? h("div", { style: Object.assign({}, dim(T.note), { color: "#d4564c" }) }, model.error) : null,
                                h("button", {
                                    type: "button",
                                    className: "om-btn",
                                    "data-model-download": model.kind,
                                    style: Object.assign({}, S.ghostButton, { alignSelf: "flex-start", marginTop: "2px" }),
                                    disabled: props.disabled === true || running || starting === model.kind,
                                    onClick: () => start(model.kind),
                                }, running ? "下载中…" : (ready ? "重新下载" : (starting === model.kind ? "启动中…" : "下载模型"))),
                            );
                        }),
                ),
            );
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
         *
         * 2026-11-19（第九十一轮）修「保存失败：写入失败：宿主拒绝了」：
         * 1) ConfigForm.mutate 只看 response.ok，失败时 recover 后 **return false**，
         *    服务端的 code 与原始 message 在 dsh-client-ui-settings 里被整个丢掉
         *    （config-form.ts:135-137）—— 面板只能显示一句没信息量的「宿主拒绝了这次写入」。
         * 2) 它还会带上 pendingRevision ?? snapshot.revision。revision 由宿主按
         *    [fiberUid, schema.toJSON(), entry.options.config] 的指纹递增（dsh-settings
         *    src/index.ts:311-315）。revision 过期时宿主按契约拒写（:394-395，
         *    settings/conflict）；**不落盘、不丢数据**，正确做法是重读后重试。
         *    （只改某个字段的 path ops 重放，不会覆盖别的字段；同字段并发写仍是后者胜。）
         *
         * 所以：拿得到 remote.settings 时，先 settings.describe() 取当前 revision，
         * 再用这个 revision 调 mutate；只有 settings/conflict 才重读 revision 后重试
         * （最多 3 次），其余拒绝不重试，把宿主 error.code / error.message 原样抛出来。
         * 读不到 revision（或没有 remote 服务）时退回 form.mutate：那一路同样带
         * revision 与队列，只是错误被宿主吞掉，提示里会点明这一点。
         *
         * 第九十五轮（现场「保存失败、只回 true/false」）：这个直连通道以前**从未
         * 生效过**。宿主把远程命名空间挂成独立服务 `remote.<ns>`
         * （dsh-api-gateway/src/client/index.ts:363-390 用 ownerCtx.plugin 异步挂载，
         * 键名见 :765-767），设置面自己 inject 的是 ['remote', 'remote.settings']
         * （dsh-client-ui-settings/src/client/index.ts:31）。本插件原来只 inject
         * ['remote']：回调在 remote 就绪那一刻跑一次，那时 settingsController 还没
         * 挂上（它排在 $mount 列表第 5 位，见 dsh-api-remotes/src/client/index.ts:199-209），
         * settingsRemote 被锁成 null，写入永远走 configForms —— 而 configForms 会把
         * 服务端的 error.code / message 丢光，只剩那句固定文案。修法见下面的
         * currentSettingsRemote() 与注册处的两段 inject。
         */
        let settingsRemote = null;

        /**
         * 宿主 remote 服务的整份引用（不是 settings 那一层）：知识库大模型清单
         * 优先走 ctx.remote.session.modelCatalog() —— 登录后可用，不需要 cwd，
         * 也不经过插件的 HTTP 路由。formToScope 仍只用 settingsRemote。
         *
         * 单独存 + 变更通知：remote 可能晚于设置页就绪（或者宿主重新注入），
         * 选择器要能按「当前」这份引用重新取数，而不是在挂载时把方法捕获死。
         */
        let hostRemote = null;
        /**
         * remote.session 是**独立服务** remote.<ns>（宿主把每个命名空间挂成扁平
         * 服务 remote.session）。严格 Cordis 下，只声明了 ['remote'] 的注入回调
         * 里读 hostRemote.session 会抛「cannot get property "remote.session"
         * without inject」；必须由声明了 remote.session 的回调把引用交到这里。
         * 命名空间卸载/重挂时，回调返回的 disposer + 重跑负责清掉旧引用。
         */
        let sessionRemote = null;
        const hostRemoteListeners = new Set();
        function notifyHostRemote() {
            for (const listener of Array.from(hostRemoteListeners)) {
                try { listener(); } catch (error) { /* 选择器自己的错误由它自己显示 */ }
            }
        }
        function setHostRemote(next) {
            const value = next || null;
            if (value === hostRemote) return;
            hostRemote = value;
            notifyHostRemote();
        }
        function setSessionRemote(next) {
            const value = next || null;
            if (value === sessionRemote) return;
            sessionRemote = value;
            notifyHostRemote();
        }
        function subscribeHostRemote(listener) {
            hostRemoteListeners.add(listener);
            return () => { hostRemoteListeners.delete(listener); };
        }

        /**
         * 当前可用的直连写入通道（每次写入现取，不缓存「没有」这个结论）。
         *
         * 为什么不能在注入回调里存一次：宿主把每个远程命名空间挂成独立服务
         * remote.<ns>（dsh-api-gateway 异步挂载），cordis 的 inject 回调在所列服务
         * 都就绪时才跑、**不会**因为事后又挂上一个嵌套服务而重跑（实测：先
         * ctx.inject(['remote']) 再挂 remote.settings，回调只跑第一次）。
         * 所以每次都从当前 remote 引用上现取 remote.settings；服务真的挂上之后，
         * 下一笔写入就能走直连。settingsRemote 只作为兜底：某些宿主实现对未挂载的
         * 嵌套服务访问会抛错，那种情况仍可用注入回调里拿到的那一份。
         */
        function currentSettingsRemote() {
            try {
                const nested = hostRemote && hostRemote.settings;
                if (nested && typeof nested.mutate === "function") return nested;
            } catch (error) {
                // 访问未挂载的嵌套服务抛错：退到注入回调里记下的那一份。
            }
            return settingsRemote !== null && typeof settingsRemote.mutate === "function" ? settingsRemote : null;
        }

        /** 一次写入最多试几次：首试 + 两次「刷新 revision 后重试」。 */
        const SET_ATTEMPTS = 3;

        /** 把宿主返回的 code / message 拼成一句给人看的话。 */
        function hostErrorText(error) {
            const code = error && typeof error.code === "string" ? error.code : "";
            const message = error && typeof error.message === "string" ? error.message : "";
            const text = message || "宿主没有给出原因";
            return code === "" ? text : code + "：" + text;
        }

        /**
         * 读一次宿主当前 revision（settings.describe 的同一个镜像）。
         *
         * 宿主契约（dsh-api-settings-controller/src/index.ts:98-103、111/147）：
         * mutate 的 expectedRevision 是「调用方读到的 revision」，undefined 才无条件写。
         * 所以这里必须先读再写，把并发保护留给宿主。
         *
         * 返回 { revision } 或 { reason }：读不到时必须说清是哪一步断的。以前
         * 静默返回 undefined，调用方只能把原因一概写成「remote.settings 不可用」——
         * 服务明明挂着的时候这是错的归因，也把真正的拒绝原因藏掉了。
         */
        async function hostRevision(remote) {
            if (remote === null || typeof remote.describe !== "function") {
                return { reason: "remote.settings 没有 describe()" };
            }
            let response;
            try {
                response = await remote.describe();
            } catch (error) {
                return { reason: "remote.settings.describe() 抛错：" + hostErrorText(error) };
            }
            if (!response || response.ok !== true) {
                const error = response && response.error ? response.error : null;
                return { reason: "remote.settings.describe() 失败：" + hostErrorText(error) };
            }
            const value = response.value || {};
            const namespaces = Array.isArray(value.namespaces) ? value.namespaces : [];
            const found = namespaces.find((item) => item && item.ns === NAMESPACE);
            if (found === undefined) return { reason: "宿主 describe 里没有命名空间 " + NAMESPACE };
            if (!Number.isInteger(found.revision)) return { reason: "宿主没给 " + NAMESPACE + " 的整数 revision" };
            return { revision: found.revision };
        }

        function formToScope(form) {
            /** 点号路径 -> 宿主 ops 用的段数组（"memory.llm.mode" -> [memory, llm, mode]）。 */
            const toPath = (field) => String(field).split(".").filter((part) => part !== "");
            /**
             * 一笔写入：一个 ops 数组 = 一次 form.mutate / remote.mutate。
             *
             * 多字段为什么必须挤进同一笔：切 provider 时 model 要跟着清空，分两次
             * 写会在两笔之间留下「新 provider + 旧 model」的坏组合，第二笔还可能
             * 撞 revision 冲突（第一笔已经把 revision 推进了）。
             */
            const writeOps = (ops) => {
                const remote = currentSettingsRemote();
                /**
                 * 退回 configForms：那条通道只回 true/false。
                 *
                 * why 是「为什么没走直连」，拼进提示里。以前这句话把原因一概写成
                 * 「remote.settings 也不可用」：服务明明挂着（只是 describe 没给出
                 * revision）时这是在误导人，也把宿主侧真正的拒绝原因藏掉了。
                 */
                const viaForm = (why) => Promise.resolve(form.mutate(ops)).then((accepted) => {
                    if (accepted === false) {
                        throw new Error("宿主拒绝了这次写入（configForms 只回 true/false，服务端原因被丢掉"
                            + (why === "" ? "" : "；" + why) + "）");
                    }
                    return accepted;
                });
                if (remote === null) return viaForm("直连通道 remote.settings 还没就绪");
                // 先读 revision 再写：这是宿主的乐观并发协议，写冲突时宿主拒写、不落盘。
                // 冲突（只有 settings/conflict）刷新 revision 后有限重试；校验/只读等拒绝
                // 一律不重试，把宿主原话带回去。
                const attempt = (remaining) => hostRevision(remote).then((probe) => {
                    if (probe.revision === undefined) return viaForm(probe.reason);
                    return Promise.resolve(remote.mutate(NAMESPACE, ops, probe.revision)).then((response) => {
                        if (response && response.ok === true) return true;
                        const error = response && response.error ? response.error : null;
                        if (error === null) {
                            throw new Error("宿主拒绝了这次写入（宿主没有给出原因）");
                        }
                        if (error.code === "settings/conflict" && remaining > 1) {
                            return attempt(remaining - 1);
                        }
                        const tried = remaining < SET_ATTEMPTS ? "（已刷新 revision 重试 " + (SET_ATTEMPTS - remaining) + " 次）" : "";
                        throw new Error("宿主拒绝了这次写入（" + hostErrorText(error) + "）" + tried);
                    });
                });
                return attempt(SET_ATTEMPTS);
            };
            return {
                getSnapshot() {
                    const snap = form.getSnapshot();
                    return { status: snap.status, value: snap.value, writable: snap.writable };
                },
                subscribe(listener) {
                    return form.subscribe(listener);
                },
                set(field, value) {
                    return writeOps([{ op: "set", path: toPath(field), value: value }]);
                },
                /**
                 * 一次 mutate 写多个字段（第九十四轮）：切 provider 要同时清空 model。
                 * entries = [[点号路径, 值], …]；空数组不发请求。
                 */
                setMany(entries) {
                    const ops = (Array.isArray(entries) ? entries : []).map((entry) => ({
                        op: "set", path: toPath(entry[0]), value: entry[1],
                    }));
                    if (ops.length === 0) return Promise.resolve(true);
                    return writeOps(ops);
                },
            };
        }

        /**
         * 面板外壳：读命名空间、写回、错误与三种兜底状态都在这里，一处一份。
         *
         * 合并后的页签共用同一个设置命名空间 —— 读写接线只该有一套，否则各页
         * 的只读提示、错误显示会各自漂移。渲染函数走具名的 render 属性
         * （而不是 children）：宿主注入的是 React，单子元素在 children 里是
         * 一个数组，取值方式容易记错，具名更直白。
         */
        function SettingsPanel(props) {
            const scope = props.scope;
            const snap = useScope(scope);
            const [writeError, setWriteError] = React.useState(null);

            // 写入返回结果而不是只发出去：调用方（如知识库目录）要等落盘成功再报「已保存」，
            // 否则写失败时面板仍显示成功，用户以为存了、下次打开又变回旧值。
            const runWrite = (run) => {
                if (!scope) return Promise.resolve({ ok: false, error: "设置服务不可用。" });
                setWriteError(null);
                return Promise.resolve(run()).then(() => ({ ok: true })).catch((e) => {
                    const message = "写入失败：" + (e && e.message ? e.message : String(e));
                    setWriteError(message);
                    return { ok: false, error: message };
                });
            };
            const write = (field, value) => runWrite(() => scope.set(field, value));
            /**
             * 一次 mutate 写多个字段（第九十四轮）：切 provider 时 model 必须跟着
             * 清空。分两次写会留下「新 provider + 旧 model」的坏组合，第二笔还可能
             * 撞 revision 冲突；所以这里优先走 scope.setMany（同一个 ops 数组）。
             * 万一 scope 是旧版（没有 setMany），退回顺序写，行为仍然可用。
             */
            const writeMany = (entries) => runWrite(() => (typeof scope.setMany === "function"
                ? scope.setMany(entries)
                : entries.reduce((chain, entry) => chain.then(() => scope.set(entry[0], entry[1])), Promise.resolve())));

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
            return props.render({ value: value, disabled: snap.writable === false, write: write, writeMany: writeMany, writeError: writeError });
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
                note: "第二手通道（默认顺序是 deepseek-api → seam → 免 Key 抓取）：宿主自带公网地址校验、地址钉死、体积与超时上限。没有参数要填（search.providers.seam 是空对象）。",
                fields: [],
            },
            {
                id: "deepseek-api", label: "DeepSeek Search API（直连）", needsKey: true, keyEnv: "DEEPSEEK_API_KEY",
                note: "直连 DeepSeek 的 Search and Fetch API：一次请求直接拿结果，不跑辅助模型；结果自带摘要（summary），可当摘录用。默认顺序排第一。",
                fields: [
                    ["search.providers.deepseek-api.apiKey", "Key", "password"],
                    ["search.providers.deepseek-api.apiKeyEnv", "Key 环境变量名", "text"],
                    ["search.providers.deepseek-api.baseURL", "端点", "text"],
                    ["search.providers.deepseek-api.timeoutMs", "超时（毫秒）", "number"],
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
         * 两种用法：
         *   - 受控：传 open + onToggle（办公模式页的「高级参数」保持原样）；
         *   - 自管：只传 defaultOpen，开合落在组件自己的 state 里 —— 设置页里
         *     一组一组的高级参数不必让外层记得每一个的开合状态。
         *
         * 默认收起，但**内容照渲染**（只是 display:none）—— 一是切换时不重建子树、
         * 草稿不丢；二是单测仍能按 props 找到里面的控件。改一个参数要翻半屏输入框，
         * 是这一页「操作不够简易」的主因。
         */
        function Fold(props) {
            const controlled = typeof props.onToggle === "function";
            const [innerOpen, setInnerOpen] = React.useState(props.defaultOpen === true);
            const open = controlled ? props.open === true : innerOpen;
            // searching：定位过滤命中折叠块里的行时由 filterRows 置上 —— 命中却看不见
            // 等于没定位到，所以定位期间折叠块一律展开（定位条清空后回到用户自己的状态）。
            const expanded = open || props.searching === true;
            return h("div", { style: props.style, "data-fold": props.name },
                h("button", {
                    type: "button",
                    className: "om-btn",
                    style: S.foldHead,
                    "data-fold-toggle": props.name,
                    "aria-expanded": expanded,
                    onClick: () => {
                        if (controlled) props.onToggle(!open);
                        else setInnerOpen(!open);
                    },
                },
                    h("span", { style: S.foldCaret }, expanded ? "▾" : "▸"),
                    h("span", { style: T.strong }, props.label),
                    props.count === undefined ? null : h("span", { style: S.foldCount }, props.count),
                    h("span", { style: S.groupHeadHint }, expanded ? "收起" : "展开"),
                ),
                h("div", { style: expanded ? undefined : { display: "none" }, "data-fold-body": props.name }, props.children),
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
        //      （例如 search.providers.deepseek-api.baseURL），所以按路径也能定位；
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
            const tone = groupToneOf(props.id);
            return h("div", { style: S.group, "data-group": props.id, "data-group-open": open ? "true" : "false" },
                h("button", {
                    type: "button",
                    className: "om-btn",
                    style: open ? Object.assign({}, S.groupHead, S.groupHeadOpen) : S.groupHead,
                    "data-group-toggle": props.id,
                    "aria-expanded": open,
                    onClick: () => props.onToggle(!open),
                },
                    h("span", {
                        style: Object.assign({}, S.foldCaret, { fontSize: "10px", width: "10px" }),
                    }, open ? "▾" : "▸"),
                    h("span", {
                        style: Object.assign({}, S.groupDot, { background: tone, color: tone }),
                        "data-group-tone": props.id,
                    }),
                    h("span", { style: S.groupHeadTitle }, parts.title),
                    props.count === undefined ? null : h("span", { style: S.groupHeadCount, "data-group-count": props.id }, props.count),
                    h("span", { style: S.groupHeadHint }, open ? "收起" : "展开"),
                ),
                h("div", {
                    style: open ? undefined : { display: "none" },
                    "data-group-body": props.id,
                },
                    // 组体单独包一层内边距：data-group-body 的 style 必须保持
                    // 「展开时 undefined / 收起时 {display:none}」，测试直接读这层。
                    h("div", { style: S.groupBody },
                        parts.head.map((item) => (item.props.style === S.groupTitle ? null : item)),
                        parts.body,
                    ),
                ),
            );
        }

        /**
         * 分组开合的持久化（键按页签分开，各页签各记各的）。
         *
         * 默认展开；用户点过之后按他的选择记住 —— 下次打开设置页仍是那样。
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
        function useGroupView(page, ids, defaultOpen) {
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
                // 合并页一屏只放 2–4 组，默认展开；用户点过「收起」之后按他的选择记。
                isOpen: (id) => (state[id] === undefined ? defaultOpen === true : state[id] === true),
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
                        placeholder: "搜索设置项",
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
                        ? (props.hits === 0 ? "没有匹配项。" : "匹配 " + props.hits + " 项。")
                        : "共 " + props.total + " 项设置"),
            );
        }

        /**
         * 设置页顶部的仪表盘（第八十六轮）。
         *
         * 四枚指标卡：总数由外壳算（与分组计数同源），其余三项由页面给
         * （工具开关 / 联网通道 / 插件，或记忆层 / 范围 / 检索方式）。
         * 定位时整块隐藏：用户已经在找具体项，不需要再看概览。
         */
        function SettingsOverview(props) {
            const facts = Array.isArray(props.facts) ? props.facts : [];
            if (facts.length === 0) return null;
            const tones = [M.business, M.success, M.warn, M.accent];
            return h("div", { style: S.settingsHero, "data-settings-overview": props.page },
                h("div", { style: S.heroKicker }, props.kicker || "OFFICE CONTROL"),
                h("div", { style: S.heroTitle }, props.title || "设置总览"),
                props.note ? h("div", { style: S.heroNote }, props.note) : null,
                h("div", { style: S.heroFacts }, facts.map((fact, index) => {
                    const tone = fact.tone || tones[index % tones.length];
                    return h("div", {
                        key: fact.key || String(index),
                        style: S.heroFact,
                        "data-settings-fact": fact.key || String(index),
                    },
                        h("div", { style: S.heroFactTop },
                            h("span", { style: Object.assign({}, S.heroFactDot, { background: tone }) }),
                            h("span", { style: S.heroFactLabel }, fact.label),
                        ),
                        h("div", {
                            style: S.heroFactValue,
                            "data-settings-fact-value": fact.key || String(index),
                        }, fact.value),
                        fact.note ? h("div", { style: S.heroFactNote }, fact.note) : null,
                    );
                })),
            );
        }

        /**
         * 设置页外壳：定位条 + 各页写好的那一串子元素。
         *
         * 对每个 SettingsGroup 做三件事：数「几项」、按定位结果过滤、把开合状态发下去；
         * 分组以外的子元素（页脚说明等）原样透传。
         */
        function SettingsBody(props) {
            const view = useGroupView(props.page, props.groups, props.defaultOpen === true);
            const visible = new Set(Array.isArray(props.groups) ? props.groups : []);
            const list = Array.isArray(props.children)
                ? props.children
                : (props.children === undefined || props.children === null ? [] : [props.children]);
            let total = 0;
            let hits = 0;
            const prepared = [];
            for (const node of list) {
                if (node === null || node === undefined || node === false) continue;
                if (node.type !== SettingsGroup) {
                    prepared.push(node);
                    continue;
                }
                if (!visible.has(node.props.id)) continue;
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
            const overview = props.overview && typeof props.overview === "object" ? props.overview : null;
            const facts = [];
            if (overview !== null) {
                facts.push({ key: "total", label: "设置项", value: total + " 项", note: "本页可调参数" });
                for (const fact of (Array.isArray(overview.facts) ? overview.facts : [])) {
                    if (facts.length >= 4) break;
                    facts.push(fact);
                }
            }
            return h("div", { style: S.section },
                // 概览只在浏览时出现：定位时用户已经在找具体项，整块让位给结果。
                !view.searching && overview !== null
                    ? h(SettingsOverview, {
                        page: props.page,
                        title: overview.title,
                        note: overview.note,
                        kicker: overview.kicker,
                        facts: facts,
                    })
                    : null,
                h(LocatorBar, { view: view, hits: hits, total: total }),
                prepared,
            );
        }

        /** 办公页签的分组（顺序即页面顺序）：用于「全部展开 / 全部收起」。 */
        const OFFICE_GROUP_IDS = ["plugins", "tools", "guard", "search", "subagents", "documents", "ocr", "python", "av"];
        /** 记忆页签的分组。 */
        const MEMORY_GROUP_IDS = ["memory", "layers", "scope", "kb", "kb-llm", "capture", "capacity", "recall", "quota"];
        /** 图库页签的分组：视觉描述与图文对齐是两件独立的能力。 */
        const GALLERY_GROUP_IDS = ["gallery", "gallery-image"];

        // ── 合并后的页签与「模式」（第一百零五轮）─────────────────────────
        //
        // 原来「办公模式」「记忆系统」是两个一级导航项，各自一屏十几张分组卡。
        // 现在合成一页，用页签把 20 个分组成八个小页面：一屏只放 2–4 组，
        // 不用在几十行里找一格。页签的顺序即用户最常走的顺序。
        const SETTINGS_TABS = [
            { id: "overview", label: "总览" },
            { id: "office", label: "办公", kind: "office", groups: ["plugins", "tools", "guard"] },
            { id: "search", label: "检索", kind: "office", groups: ["search", "subagents"] },
            { id: "docs", label: "文档", kind: "office", groups: ["documents", "ocr", "python", "av"] },
            { id: "gallery", label: "图库", kind: "gallery", groups: ["gallery", "gallery-image"] },
            { id: "memory", label: "记忆", kind: "memory", groups: ["memory", "layers", "scope", "capture"] },
            { id: "kb", label: "知识库", kind: "memory", groups: ["kb", "kb-llm"] },
            { id: "recall", label: "召回", kind: "memory", groups: ["capacity", "recall", "quota"] },
        ];

        function settingsTabOf(id) {
            for (const tab of SETTINGS_TABS) {
                if (tab.id === id) return tab;
            }
            return SETTINGS_TABS[0];
        }

        /**
         * 一套超参 = 一个模式。只收「改了就影响检索质量 / 成本」的项：
         * 向量通道、筛选、ANN、融合、召回分档与每回合配额。
         * 凭据与路径不在里面 —— 模式不该覆盖用户的 Key 或模型目录。
         */
        const HYPER_FIELDS = [
            "memory.embedding.mode",
            "memory.embedding.local.enabled",
            "memory.selector.mode",
            "memory.ann.enabled",
            "memory.ann.efSearch",
            "memory.hybrid.lexicalWeight",
            "memory.hybrid.vectorWeight",
            "memory.hybrid.candidateLimit",
            "memory.recallQuality.policy",
            "memory.recallQuality.candidateMultiplier",
            "memory.recallQuality.maxMediumResults",
            "memory.recallQuality.maxUnknownResults",
            "memory.quota.recallPerTurn",
            "memory.quota.recallRefinePerTurn",
            "memory.quota.kbSearchPerTurn",
            "memory.quota.relatedPerTurn",
        ];

        const HYPER_PRESETS = [
            {
                id: "balanced", name: "均衡",
                note: "默认；词法为主，配了向量就融合",
                values: {
                    "memory.embedding.mode": "auto",
                    "memory.embedding.local.enabled": false,
                    "memory.selector.mode": "off",
                    "memory.ann.enabled": true,
                    "memory.ann.efSearch": 64,
                    "memory.hybrid.lexicalWeight": 0.55,
                    "memory.hybrid.vectorWeight": 0.45,
                    "memory.hybrid.candidateLimit": 30,
                    "memory.recallQuality.policy": "strict-v1",
                    "memory.recallQuality.candidateMultiplier": 5,
                    "memory.recallQuality.maxMediumResults": 6,
                    "memory.recallQuality.maxUnknownResults": 4,
                    "memory.quota.recallPerTurn": 1,
                    "memory.quota.recallRefinePerTurn": 1,
                    "memory.quota.kbSearchPerTurn": 2,
                    "memory.quota.relatedPerTurn": 1,
                },
            },
            {
                id: "lite", name: "轻量",
                note: "纯词法，不建索引；最省内存",
                values: {
                    "memory.embedding.mode": "off",
                    "memory.embedding.local.enabled": false,
                    "memory.selector.mode": "off",
                    "memory.ann.enabled": false,
                    "memory.ann.efSearch": 64,
                    "memory.hybrid.lexicalWeight": 1,
                    "memory.hybrid.vectorWeight": 0,
                    "memory.hybrid.candidateLimit": 20,
                    "memory.recallQuality.policy": "strict-v1",
                    "memory.recallQuality.candidateMultiplier": 3,
                    "memory.recallQuality.maxMediumResults": 3,
                    "memory.recallQuality.maxUnknownResults": 1,
                    "memory.quota.recallPerTurn": 1,
                    "memory.quota.recallRefinePerTurn": 1,
                    "memory.quota.kbSearchPerTurn": 1,
                    "memory.quota.relatedPerTurn": 1,
                },
            },
            {
                id: "local", name: "本地语义",
                note: "开本机向量；首次约 1–2 分钟",
                values: {
                    "memory.embedding.mode": "auto",
                    "memory.embedding.local.enabled": true,
                    "memory.selector.mode": "off",
                    "memory.ann.enabled": true,
                    "memory.ann.efSearch": 96,
                    "memory.hybrid.lexicalWeight": 0.5,
                    "memory.hybrid.vectorWeight": 0.5,
                    "memory.hybrid.candidateLimit": 50,
                    "memory.recallQuality.policy": "strict-v1",
                    "memory.recallQuality.candidateMultiplier": 5,
                    "memory.recallQuality.maxMediumResults": 6,
                    "memory.recallQuality.maxUnknownResults": 4,
                    "memory.quota.recallPerTurn": 2,
                    "memory.quota.recallRefinePerTurn": 2,
                    "memory.quota.kbSearchPerTurn": 3,
                    "memory.quota.relatedPerTurn": 1,
                },
            },
            {
                id: "deep", name: "深检索",
                note: "更大候选池与召回名额；更慢",
                values: {
                    "memory.embedding.mode": "auto",
                    "memory.embedding.local.enabled": false,
                    "memory.selector.mode": "off",
                    "memory.ann.enabled": true,
                    "memory.ann.efSearch": 128,
                    "memory.hybrid.lexicalWeight": 0.4,
                    "memory.hybrid.vectorWeight": 0.6,
                    "memory.hybrid.candidateLimit": 50,
                    "memory.recallQuality.policy": "strict-v1",
                    "memory.recallQuality.candidateMultiplier": 5,
                    "memory.recallQuality.maxMediumResults": 8,
                    "memory.recallQuality.maxUnknownResults": 5,
                    "memory.quota.recallPerTurn": 2,
                    "memory.quota.recallRefinePerTurn": 2,
                    "memory.quota.kbSearchPerTurn": 4,
                    "memory.quota.relatedPerTurn": 2,
                },
            },
        ];

        function settingPathValue(root, path) {
            let cur = root;
            for (const part of String(path).split(".")) {
                if (cur === null || cur === undefined || typeof cur !== "object") return undefined;
                cur = cur[part];
            }
            return cur;
        }

        /** 当前值 + 默认值：缺字段时按「均衡」补齐，保存自定义模式不会存进 undefined。 */
        function hyperValuesOf(value) {
            const out = {};
            const fallback = HYPER_PRESETS[0].values;
            for (const path of HYPER_FIELDS) {
                const current = settingPathValue(value, path);
                out[path] = current === undefined ? fallback[path] : current;
            }
            return out;
        }

        function hyperPresetMatches(value, preset) {
            for (const path of HYPER_FIELDS) {
                const current = settingPathValue(value, path);
                if (String(current) !== String(preset.values[path])) return false;
            }
            return true;
        }

        const PRESET_STORE_KEY = "dsh-office-mode.hyper-presets.v1";
        const PRESET_STORE_MEMORY = { loaded: false, value: [] };

        function presetStorage() {
            try {
                const store = typeof window !== "undefined" && window ? window.localStorage : undefined;
                return store && typeof store.getItem === "function" && typeof store.setItem === "function" ? store : null;
            } catch (error) {
                return null;
            }
        }

        function normalizeCustomPreset(raw) {
            if (raw === null || typeof raw !== "object") return null;
            const name = String(raw.name || "").trim().slice(0, 20);
            if (name === "") return null;
            const values = {};
            for (const path of HYPER_FIELDS) {
                const value = raw.values && raw.values[path] !== undefined ? raw.values[path] : HYPER_PRESETS[0].values[path];
                values[path] = value;
            }
            return { id: "custom:" + name, name: name, values: values };
        }

        function loadCustomPresets() {
            if (PRESET_STORE_MEMORY.loaded) return PRESET_STORE_MEMORY.value.slice();
            PRESET_STORE_MEMORY.loaded = true;
            const store = presetStorage();
            if (store !== null) {
                try {
                    const raw = store.getItem(PRESET_STORE_KEY);
                    const parsed = raw ? JSON.parse(raw) : [];
                    if (Array.isArray(parsed)) {
                        PRESET_STORE_MEMORY.value = parsed.map(normalizeCustomPreset).filter((item) => item !== null).slice(0, 12);
                    }
                } catch (error) {
                    PRESET_STORE_MEMORY.value = [];
                }
            }
            return PRESET_STORE_MEMORY.value.slice();
        }

        function saveCustomPresets(list) {
            PRESET_STORE_MEMORY.loaded = true;
            PRESET_STORE_MEMORY.value = Array.isArray(list) ? list.slice(0, 12) : [];
            const store = presetStorage();
            if (store === null) return;
            try {
                store.setItem(PRESET_STORE_KEY, JSON.stringify(PRESET_STORE_MEMORY.value));
            } catch (error) {
                // 存不下就只保留内存里的这一份。
            }
        }

        const TAB_STORE_KEY = "dsh-office-mode.settings-tab.v1";
        const TAB_STORE_MEMORY = { value: "overview" };

        function loadSettingsTab() {
            const store = presetStorage();
            if (store !== null) {
                try {
                    const raw = store.getItem(TAB_STORE_KEY);
                    if (typeof raw === "string" && raw !== "") TAB_STORE_MEMORY.value = raw;
                } catch (error) {
                    // 读不到就回到总览。
                }
            }
            return settingsTabOf(TAB_STORE_MEMORY.value).id;
        }

        function saveSettingsTab(id) {
            TAB_STORE_MEMORY.value = settingsTabOf(id).id;
            const store = presetStorage();
            if (store === null) return;
            try {
                store.setItem(TAB_STORE_KEY, TAB_STORE_MEMORY.value);
            } catch (error) {
                // 同上：不影响本次切换。
            }
        }

        const TAB_STYLE = {
            bar: {
                display: "flex", alignItems: "center", gap: "4px",
                overflowX: "auto", padding: "0 0 10px", marginBottom: "12px",
                borderBottom: "1px solid " + M.stroke,
            },
            tab: Object.assign({
                flex: "0 0 auto", padding: "7px 13px", borderRadius: "9px",
                border: "1px solid transparent", background: "transparent",
                color: M.textTertiary, cursor: "pointer", whiteSpace: "nowrap",
            }, T.label),
            tabActive: {
                background: M.sunken, borderColor: M.stroke, color: M.text, fontWeight: 500,
            },
        };

        function SettingsTabs(props) {
            // 标准 tablist 键盘：左右循环、Home / End 跳到两端。只让当前页签进
            // Tab 顺序，其余用方向键到达；切换后把焦点送过去，读屏才跟得上。
            const moveTab = (event) => {
                const key = event && event.key;
                if (key !== "ArrowRight" && key !== "ArrowLeft" && key !== "Home" && key !== "End") return;
                event.preventDefault();
                const current = SETTINGS_TABS.findIndex((tab) => tab.id === props.tab);
                const last = SETTINGS_TABS.length - 1;
                const next = key === "Home" ? 0
                    : key === "End" ? last
                        : key === "ArrowRight" ? (current + 1 + SETTINGS_TABS.length) % SETTINGS_TABS.length
                            : (current - 1 + SETTINGS_TABS.length) % SETTINGS_TABS.length;
                const target = SETTINGS_TABS[next];
                if (!target || target.id === props.tab) return;
                props.onChange(target.id);
                const parent = event.currentTarget && event.currentTarget.parentNode;
                if (parent && typeof parent.querySelector === "function") {
                    const button = parent.querySelector('[data-settings-tab="' + target.id + '"]');
                    if (button && typeof button.focus === "function") button.focus();
                }
            };
            return h("div", {
                style: TAB_STYLE.bar, role: "tablist", "data-settings-tabs": "true",
            }, SETTINGS_TABS.map((tab) => h("button", {
                key: tab.id, type: "button", role: "tab",
                className: "om-btn",
                "data-settings-tab": tab.id,
                "aria-selected": props.tab === tab.id,
                tabIndex: props.tab === tab.id ? 0 : -1,
                style: props.tab === tab.id ? Object.assign({}, TAB_STYLE.tab, TAB_STYLE.tabActive) : TAB_STYLE.tab,
                onClick: () => props.onChange(tab.id),
                onKeyDown: moveTab,
            }, tab.label)));
        }

        const PRESET_STYLE = {
            home: { padding: "2px 0" },
            head: { display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: "12px", marginBottom: "6px" },
            title: Object.assign({}, T.heading),
            note: Object.assign({}, dim(T.note)),
            grid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: "8px", marginBottom: "10px" },
            card: {
                position: "relative", borderRadius: "12px", border: "1px solid " + M.stroke,
                background: M.surface, overflow: "hidden",
            },
            cardActive: { borderColor: "var(--dsw-alias-state-business-primary, #4f7cff)" },
            cardMain: {
                display: "block", width: "100%", padding: "12px 14px", textAlign: "left",
                border: "none", background: "none", color: "inherit", cursor: "pointer",
            },
            cardTop: { display: "flex", alignItems: "center", gap: "7px" },
            cardName: Object.assign({}, T.strong),
            badge: Object.assign({ marginLeft: "auto", padding: "1px 7px", borderRadius: M.pill }, T.tiny, {
                color: "var(--dsw-alias-state-business-primary, #4f7cff)",
                background: "color-mix(in srgb, var(--dsw-alias-state-business-primary, #4f7cff) 12%, transparent)",
            }),
            cardNote: Object.assign({ marginTop: "3px" }, dim(T.note)),
            remove: Object.assign({ position: "absolute", right: "5px", bottom: "4px", border: "none", background: "none", cursor: "pointer" }, dim(T.tiny)),
            saveRow: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "8px", marginTop: "10px" },
            status: Object.assign({ marginTop: "10px" }, dim(T.note)),
            hint: Object.assign({ marginTop: "6px" }, dim(T.note)),
        };

        function PresetHome(props) {
            return h(SettingsPanel, {
                scope: props.scope,
                render: (state) => h(PresetHomeBody, state),
            });
        }

        /**
         * 模式卡的内容。
         *
         * 必须是一个真组件，不能把 useState 写在 SettingsPanel 的 render 回调里：
         * 回调第一次渲染（还在读设置）不会执行，第二次才执行，钩子数量会变，
         * React 会抛「Rendered more hooks than during the previous render」，
         * 设置页整块空白。
         */
        function PresetHomeBody(props) {
            const { value, disabled, writeMany, writeError } = props;
            const [customs, setCustoms] = React.useState(loadCustomPresets);
            const [name, setName] = React.useState("");
            const [hint, setHint] = React.useState("");
            const all = HYPER_PRESETS.concat(customs);
            const active = all.find((preset) => hyperPresetMatches(value, preset)) || null;
            const memory = value.memory || {};
            const route = retrievalStatusOf(memory, memory.embedding || {}, (memory.embedding || {}).local || {}, memory.selector || {});
            const apply = (preset) => {
                if (disabled) return;
                setHint("");
                const entries = HYPER_FIELDS.map((path) => [path, preset.values[path]]);
                Promise.resolve(writeMany(entries)).then((result) => {
                    if (result && result.ok === false) setHint(result.error || "写入失败。");
                    else setHint("已切换到「" + preset.name + "」。");
                });
            };
            const save = () => {
                const clean = name.trim().slice(0, 20);
                if (clean === "") {
                    setHint("先给模式起个名字。");
                    return;
                }
                if (HYPER_PRESETS.some((preset) => preset.name === clean)) {
                    setHint("内置模式名不能覆盖。");
                    return;
                }
                const item = { id: "custom:" + clean, name: clean, values: hyperValuesOf(value) };
                const at = customs.findIndex((preset) => preset.name === clean);
                if (at === -1 && customs.length >= 12) {
                    setHint("最多保存 12 个模式。");
                    return;
                }
                const next = at === -1 ? customs.concat(item) : customs.map((preset, index) => (index === at ? item : preset));
                setCustoms(next);
                saveCustomPresets(next);
                setName("");
                setHint(at === -1 ? "已保存「" + clean + "」。" : "已更新「" + clean + "」。");
            };
            const remove = (preset) => {
                const next = customs.filter((item) => item.id !== preset.id);
                setCustoms(next);
                saveCustomPresets(next);
                setHint("已删除「" + preset.name + "」。");
            };
            return h("div", { style: PRESET_STYLE.home, "data-preset-home": "true" },
                writeError !== null ? h("div", { style: S.error }, writeError) : null,
                h("div", { style: PRESET_STYLE.head },
                    h("div", { style: PRESET_STYLE.title }, "模式"),
                    h("div", { style: PRESET_STYLE.note }, "切换即应用一组参数。"),
                ),
                h("div", { style: PRESET_STYLE.grid },
                    all.map((preset) => {
                        const selected = active !== null && active.id === preset.id;
                        const custom = preset.id.indexOf("custom:") === 0;
                        return h("div", {
                            key: preset.id,
                            style: selected ? Object.assign({}, PRESET_STYLE.card, PRESET_STYLE.cardActive) : PRESET_STYLE.card,
                            "data-preset-card": preset.id,
                            "data-preset-active": selected ? "true" : "false",
                        },
                            h("button", {
                                type: "button", className: "om-btn", style: PRESET_STYLE.cardMain,
                                "data-preset-apply": preset.id,
                                disabled: disabled,
                                onClick: () => apply(preset),
                            },
                                h("div", { style: PRESET_STYLE.cardTop },
                                    h("span", { style: PRESET_STYLE.cardName }, preset.name),
                                    selected ? h("span", { style: PRESET_STYLE.badge }, "当前") : null,
                                ),
                                h("div", { style: PRESET_STYLE.cardNote }, preset.note),
                            ),
                            custom ? h("button", {
                                type: "button", style: PRESET_STYLE.remove,
                                "data-preset-remove": preset.id,
                                onClick: () => remove(preset),
                            }, "删除") : null,
                        );
                    }),
                ),
                h("div", { style: PRESET_STYLE.saveRow },
                    h("input", {
                        type: "text", value: name, maxLength: 20,
                        placeholder: "模式名称",
                        "data-preset-name": "true",
                        disabled: disabled,
                        onChange: (event) => setName(event.target.value),
                        onKeyDown: (event) => { if (event.key === "Enter") save(); },
                        style: Object.assign({}, S.input, { width: "180px", textAlign: "left" }),
                    }),
                    h("button", {
                        type: "button", className: "om-btn", style: S.button,
                        "data-preset-save": "true", disabled: disabled, onClick: save,
                    }, "保存当前参数"),
                ),
                h("div", { style: PRESET_STYLE.status, "data-settings-current": "true" },
                    "当前：" + (active === null ? "自定义" : active.name) + " · 检索：" + route.value),
                hint === "" ? null : h("div", { style: PRESET_STYLE.hint, "data-preset-hint": "true" }, hint),
            );
        }

        /**
         * 合并页：一页入口，八个页签。总览放模式；其余页签只渲染自己那几组，
         * 分组默认展开（每页只有 2–4 组，再点一次「展开」才是多余动作）。
         */
        function MergedSettingsSection(props) {
            const initial = settingsTabOf(typeof props.initialTab === "string" ? props.initialTab : loadSettingsTab()).id;
            const [tab, setTab] = React.useState(initial);
            React.useEffect(() => {
                if (typeof props.initialTab === "string" && settingsTabOf(props.initialTab).id !== tab) {
                    setTab(settingsTabOf(props.initialTab).id);
                }
            }, [props.initialTab]);
            const change = (id) => {
                setTab(settingsTabOf(id).id);
                saveSettingsTab(id);
            };
            const current = settingsTabOf(tab);
            let content = null;
            if (current.id === "overview") {
                content = h(PresetHome, { scope: props.scope });
            } else if (current.kind === "office") {
                content = h(OfficeSettingsSection, {
                    scope: props.scope, groups: current.groups, page: "office",
                    hideOverview: true, defaultOpen: true,
                });
            } else if (current.kind === "gallery") {
                content = h(GallerySettingsSection, {
                    scope: props.scope, groups: current.groups, page: "gallery",
                });
            } else {
                content = h(MemorySettingsSection, {
                    scope: props.scope, groups: current.groups, page: "memory",
                    hideOverview: true, defaultOpen: true,
                });
            }
            return h("div", { style: S.section, "data-settings-page": "merged" },
                h(SettingsTabs, { tab: current.id, onChange: change }),
                content,
            );
        }

        /**
         * 分组卡片左侧的色点。颜色不承担语义，只给「一屏十几张卡」一点定位锚点；
         * 没有登记的 id 退中性灰，不会因为新加一组就报错。
         */
        const GROUP_TONES = {
            plugins: "#7c6cff", tools: "#4f7cff", guard: "#e5484d", search: "#22a879",
            subagents: "#a855f7", documents: "#c08a2e", ocr: "#0e9488", python: "#3b9cda", av: "#e6689b",
            memory: "#4f7cff", layers: "#22a879", scope: "#a855f7", kb: "#c08a2e", "kb-llm": "#8b5cf6",
            capture: "#3b9cda", capacity: "#e6689b", recall: "#7c6cff", quota: "#c08a2e",
            gallery: "#b7791f", "gallery-image": "#0f766e",
        };
        function groupToneOf(id) {
            const key = String(id || "");
            return Object.prototype.hasOwnProperty.call(GROUP_TONES, key) ? GROUP_TONES[key] : "#6b7280";
        }

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
                    ? h("div", { style: S.status, "data-site-empty": "true" }, "清单为空：这次不做站点限定。")
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
            );
        }

        function OfficeSettingsSection(props) {
            const pageGroups = Array.isArray(props.groups) ? props.groups : OFFICE_GROUP_IDS;
            const pageKey = typeof props.page === "string" ? props.page : "office";
            /** 只放「界面本地状态」：折叠开关与正在编辑哪条通道的参数（都不写设置）。 */
            const [ui, setUi] = React.useState({ paramsChannel: "deepseek-api", advancedSearch: false, advancedDocs: false, toolsHint: null });
            const patchUi = (patch) => setUi(Object.assign({}, ui, patch));

            /**
             * 「办公模式插件」那一格的状态（第六十三轮 新-21）。
             *
             * 与页面上其它分组不同，这一格的数据**不来自设置文档**：它读的是 preset 文件
             * （GET /office-preset/plugins），开关也是直接写那个文件。所以三件事必须说清：
             *   - `data.rows` 是**每次打开 / 点刷新时现读的磁盘状态**（外部改了文件，刷新就跟着变）；
             *   - 写入只影响**新建的办公会话**（preset 在新建会话时读取），回执里照实说；
             *   - 端点不可用时（宿主没有这条路由 / 环境没有 fetch）照实显示原因，
             *     而不是画一个点了没反应的开关。
             */
            const [preset, setPreset] = React.useState({ loading: true, error: null, data: null, busy: null, hint: null });
            const latestPreset = React.useRef(preset);
            latestPreset.current = preset;
            const patchPreset = (patch) => setPreset(Object.assign({}, latestPreset.current, patch));

            /** 统一的端点响应处理：只要 body 带 error 就照实显示，不靠状态码判断。 */
            const endpointOf = (url, init, onBody, onFail) => {
                if (typeof fetch !== "function") {
                    onFail("当前环境没有 fetch（这一格只在浏览器里工作）。");
                    return;
                }
                let pending;
                try {
                    pending = fetch(url, init);
                } catch (error) {
                    onFail(error && error.message ? error.message : String(error));
                    return;
                }
                Promise.resolve(pending).then(
                    (response) => Promise.resolve(response && typeof response.json === "function" ? response.json() : null)
                        .then((body) => ({ response: response, body: body }), () => ({ response: response, body: null })),
                ).then((result) => {
                    const body = result.body;
                    const status = result.response && result.response.status ? result.response.status : "?";
                    if (body === null || typeof body !== "object") {
                        onFail("端点没有返回 JSON（HTTP " + status + "）。");
                        return;
                    }
                    onBody(body);
                }, (error) => onFail(error && error.message ? error.message : String(error)));
            };

            /** 读一次那一格（GET）。 */
            const loadPreset = (hint) => {
                patchPreset(hint === undefined ? { loading: true, error: null } : { loading: true, error: null, hint: hint });
                endpointOf(PRESET_PLUGINS_PATH, { headers: { accept: "application/json" } }, (body) => {
                    if (body.ok !== true) {
                        patchPreset({ loading: false, error: String(body.error || "读取插件列表失败。"), data: body });
                        return;
                    }
                    patchPreset({ loading: false, error: null, data: body });
                }, (message) => patchPreset({ loading: false, error: "读取插件列表失败：" + message }));
            };

            /** 开关一行（POST）：回执带回新状态与「只影响新建会话」的说明。 */
            const setPlugin = (id, enabled) => {
                patchPreset({ busy: id, error: null, hint: null });
                endpointOf(PRESET_PLUGINS_PATH, {
                    method: "POST",
                    headers: { "content-type": "application/json", accept: "application/json" },
                    body: JSON.stringify({ id: id, enabled: enabled === true }),
                }, (body) => {
                    if (body.ok !== true) {
                        patchPreset({
                            busy: null,
                            error: String(body.error || "写入失败。"),
                            data: Array.isArray(body.rows) ? body : latestPreset.current.data,
                        });
                        return;
                    }
                    const applied = body.applied || {};
                    const wrote = applied.changed && applied.changed.profile === true ? "preset 真源 + profile 覆盖行" : "preset 真源";
                    patchPreset({
                        busy: null,
                        error: null,
                        data: body,
                        hint: "\"" + id + "\" 已" + (applied.enabled === true ? "启用" : "停用") + "（写了 " + wrote + "）。"
                            + "生效范围：" + String(body.effect || "新建的办公会话") + "；已经在跑的会话不受影响。",
                    });
                }, (message) => patchPreset({ busy: null, error: "写入失败：" + message }));
            };

            /** 还原上一次备份（POST）。 */
            const restorePreset = () => {
                patchPreset({ busy: "restore", error: null, hint: null });
                endpointOf(PRESET_RESTORE_PATH, {
                    method: "POST",
                    headers: { "content-type": "application/json", accept: "application/json" },
                    body: "{}",
                }, (body) => {
                    if (body.ok !== true) {
                        patchPreset({ busy: null, error: String(body.error || "还原失败。") });
                        return;
                    }
                    patchPreset({
                        busy: null,
                        error: null,
                        data: body,
                        hint: "已还原上一次备份（" + String(body.restoredAt || "时间未知") + "），两份文件都回去了。",
                    });
                }, (message) => patchPreset({ busy: null, error: "还原失败：" + message }));
            };

            // 打开「办公」页签才读一次：这一格没有「配置默认值」可显示，磁盘就是唯一真源。
            React.useEffect(() => {
                if (pageGroups.indexOf("plugins") !== -1) loadPreset();
            }, [pageGroups.join(",")]);

            return h(SettingsPanel, { scope: props.scope, render: ({ value, disabled, write, writeError }) => {
            const presetData = preset.data || {};
            const presetRows = Array.isArray(presetData.rows) ? presetData.rows : [];
            const presetBusy = preset.busy !== null;
            const tools = value.tools || {};
            const guard = value.guard || {};
            const search = value.search || {};
            const fetchLimits = search.fetch || {};
            const providers = search.providers || {};
            const preprocess = search.preprocess || {};
            // 站点清单：`entries` 只在真的拿到数组时才算「有值」（undefined → 内置目录，
            // [] → 空清单），这个区分在 SiteCatalogEditor 里处理。
            const sites = search.sites || {};
            const documents = value.documents || {};
            const ocr = value.ocr || {};
            const python = value.python || {};
            const av = value.av || {};
            const subagentModel = value.subagentModel || {};
            // 正在编辑参数的通道：默认跟着「联网通道」走，但允许单独切换 ——
            // 想先把 Tavily 的 Key 填好、但把通道留着 auto 的人不该被迫改通道。
            const paramsChannel = CHANNEL_BY_ID[ui.paramsChannel] !== undefined
                ? ui.paramsChannel
                : (search.provider && CHANNEL_BY_ID[search.provider] !== undefined ? search.provider : "deepseek-api");
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
            // 设置页顶部的四枚指标：设置总数由外壳算，这里给另外三项。
            // 只数布尔型工具开关，避免把带子字段的工具对象也当成「开」。
            const toolKeys = Object.keys(tools).filter((key) => typeof tools[key] === "boolean");
            const enabledTools = toolKeys.filter((key) => tools[key] === true).length;
            const enabledPlugins = presetRows.filter((row) => row.enabled === true).length;
            const officeOverview = {
                kicker: "OFFICE CONTROL",
                title: "办公模式 · 设置总览",
                note: "常用能力一屏看清；点分组标题展开逐项调整。",
                facts: [
                    { key: "tools", label: "启用工具", value: enabledTools + " / " + toolKeys.length, note: "当前打开的办公工具" },
                    { key: "channels", label: "可用通道", value: enabledChannels.length + " / " + CHANNEL_TABLE.length, note: "填了 Key 的联网通道" },
                    { key: "plugins", label: "办公插件", value: presetRows.length === 0 ? "—" : enabledPlugins + " / " + presetRows.length, note: "preset 声明的插件行" },
                ],
            };

            return h(SettingsBody, {
                page: pageKey, groups: pageGroups,
                overview: props.hideOverview === true ? null : officeOverview,
                defaultOpen: props.defaultOpen === true,
            },
                writeError !== null ? h("div", { key: "error", style: S.error }, writeError) : null,
                disabled
                    ? h("div", { key: "readonly", style: S.status }, "当前连接以只读方式同步设置（memory 模式），改动不会持久化。")
                    : null,

                // ── 办公模式插件（第六十三轮 新-21）──
                //
                // 用户原话：「在设置中增设办公模式插件列表，使得用户能主动决定启用哪些插件，
                // 并在用户开关后立刻在 preset 中改变（对已经启用的对话无影响）。」
                //
                // 这一格与其它分组的根本区别：它改的**不是插件自己的设置文档**，而是办公
                // preset 的声明文件。所以：
                //   - 数据来自磁盘（那条 GET），不来自 configForms；「刷新」就是重新读一次；
                //   - 开关 POST 出去之后，宿主会同时写 preset 真源与 profile 侧**已存在的**
                //     覆盖行（后者的来历见 src/preset-plugins.js 的头部注释），并先落一份快照；
                //   - 「对已经启用的对话无影响」是机制自带的（preset 在新建会话时读取），
                //     这里把它写成回执里的一句话，不靠读者猜。
                h(SettingsGroup, { key: "plugins", id: "plugins" },
                    h("div", { style: S.groupTitle }, "办公模式插件"),
                    h("div", { style: S.groupNote },
                        "preset 里的插件行；开关回写到那一行，新建会话生效。"),
                    h("div", { style: S.quickBar },
                        h("button", {
                            type: "button", className: "om-btn", style: S.button,
                            "data-preset-refresh": "true",
                            disabled: preset.loading === true || presetBusy,
                            onClick: () => loadPreset("已按磁盘上的最新内容刷新。"),
                        }, "刷新"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.button,
                            "data-preset-restore": "true",
                            disabled: presetBusy || presetData.backup === null || presetData.backup === undefined,
                            onClick: restorePreset,
                        }, "还原上一次备份"),
                        presetData.backup === null || presetData.backup === undefined
                            ? null
                            : h("span", { style: S.status, "data-preset-backup": "true" },
                                "备份：" + String(presetData.backup.name || "") + "（还原会把两个文件一起写回）"),
                    ),
                    preset.error !== null && preset.error !== undefined
                        ? h("div", { key: "preset-error", style: S.error, "data-preset-error": "true" }, preset.error)
                        : null,
                    preset.hint === null || preset.hint === undefined
                        ? null
                        : h("div", { key: "preset-hint", style: S.status, "data-preset-hint": "true" }, preset.hint),
                    preset.loading === true
                        ? h("div", { key: "preset-loading", style: S.status }, "正在读取插件列表…")
                        : null,
                    preset.loading !== true && presetRows.length === 0
                        ? h("div", { key: "preset-empty", style: S.status },
                            "没有读到插件行：preset 文件里没有 plugins: 列表，或者这一格读不到那份文件。")
                        : null,
                    presetRows.length > 0
                        ? h("div", { key: "preset-paths", style: S.status, "data-preset-paths": "true" },
                            "写到：" + String(presetData.presetPath || "")
                            + (presetData.profilePath === null || presetData.profilePath === undefined
                                ? "（这个部署里 profile 侧没有覆盖文件，只写真源）"
                                : " 与 " + String(presetData.profilePath)))
                        : null,
                    presetRows.map((row, index) => h(Row, {
                        key: row.id,
                        label: row.id,
                        first: index === 0,
                        note: String(row.name || "（没有 name）")
                            + "｜当前：" + (row.enabled === true ? "启用" : "停用")
                            + (row.profileOverridden === true ? "｜profile 侧有同名覆盖行，开关两处同写" : ""),
                    },
                        h("span", {
                            style: { display: "inline-flex", alignItems: "center", gap: "8px" },
                            "data-preset-plugin": row.id,
                            "data-preset-enabled": row.enabled === true ? "true" : "false",
                        },
                            h(Toggle, {
                                checked: row.enabled === true,
                                disabled: presetBusy,
                                onChange: (next) => setPlugin(row.id, next),
                            }),
                            preset.busy === row.id ? h("span", { style: S.status }, "写入中…") : null,
                        ),
                    )),
                ),

                // ── 工具开关 ──
                h(SettingsGroup, { key: "tools", id: "tools" },
                    h("div", { style: S.groupTitle }, "工具开关"),
                    h("div", { style: S.groupNote },
                        "关掉的工具不出现，也不占 schema 开销；即时生效。"),
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

                // ── preset 工具守卫（新-15①） ──
                //
                // 这一段解释的是一件反直觉的事：办公 preset 里写着「关掉编程类工具」的
                // 那些行**运行期不生效**（它们关的是 preset 自己那一层，而那些能力挂在
                // 全局层）。真正能按会话掩掉继承工具的只有插件的 restrict，所以这里给一个
                // 开关：打开之后由插件来掩。默认关着 —— 检索派工要用子代理、长任务要用
                // 后台任务，隐藏它们是削减能力。
                h(SettingsGroup, { key: "guard", id: "guard" },
                    h("div", { style: S.groupTitle }, "办公模式的工具面守卫"),
                    h("div", { style: S.groupNote },
                        "打开后由插件真正掩掉编程类工具：命令行、后台任务、子代理等入口消失。"),
                    h(Row, {
                        label: "藏掉编程类工具",
                        note: "默认关。打开后没有命令行、后台任务、子代理、待办与技能入口；检索派工不受影响。",
                        first: true,
                    },
                        h(Toggle, {
                            checked: guard.hideDevToolsInOffice === true,
                            disabled: disabled,
                            onChange: (next) => write("guard.hideDevToolsInOffice", next),
                        }),
                    ),
                ),

                // ── 检索编排 ──
                h(SettingsGroup, { key: "search", id: "search" },
                    h("div", { style: S.groupTitle }, "检索编排"),
                    h("div", { style: S.groupNote },
                        "检索子代理的引擎、通道与上限。"),
                    h(Row, { label: "执行引擎", first: true },
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
                        note: "免 Key 可用 DuckDuckGo / SearXNG。",
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
                                ? "已配置的通道：" + enabledChannels.map((item) => item.label).join("、")
                                : "只用这一条。"),
                    ),
                    h(Row, {
                        label: "出口代理",
                        note: "留空直连；只认 http / https。",
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
                    h(Row, { label: "填哪条通道的参数", note: "下面只针对选中的通道。" },
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
                                    ? "可直接粘 Key，也可只填环境变量名。"
                                    : "这条通道只需填地址。",
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
                        label: "高级参数（通道端点 / 超时，取正文与 PDF 上限）",
                        count: "改了才生效",
                        open: ui.advancedSearch,
                        onToggle: (next) => patchUi({ advancedSearch: next }),
                    },
                        h("div", { style: S.groupNote },
                            "只影响当前通道；取正文上限与通道无关。"),
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
                        // ── 取正文与 PDF 的上限（与通道无关）──
                        h("div", { style: Object.assign({}, S.groupTitle, { marginTop: "16px" }) }, "取正文与 PDF 上限"),
                        h("div", { style: S.groupNote },
                            "每个渠道的来源条数与取正文上限。"),
                        ...FETCH_ROWS.map((row, index) => h(Row, {
                            key: row.path,
                            label: row.label,
                            note: row.note,
                            first: index === 0,
                        },
                            row.kind === "number"
                                ? h(NumberField, {
                                    value: fetchLimits[row.key] === undefined ? row.min : fetchLimits[row.key],
                                    min: row.min, max: row.max, step: row.step, disabled: disabled,
                                    onChange: (next) => write(row.path, next),
                                })
                                : h("input", {
                                    type: row.kind === "password" ? "password" : "text",
                                    defaultValue: fetchLimits[row.key] === undefined ? "" : fetchLimits[row.key],
                                    disabled: disabled,
                                    onBlur: (event) => {
                                        const next = String(event.target.value || "").trim();
                                        if (next !== fetchLimits[row.key]) write(row.path, next);
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
                        note: "越靠前越先试；没配 Key 的通道会当场失败。",
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
                        note: "去脚本、导航与样板行；清单页建议 plain 或 off。",
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
                    h(Row, { label: "丢掉样板行", note: "cookie 条 / 登录注册 / 版权声明这类整行样板。" },
                        h(Toggle, {
                            checked: preprocess.dropBoilerplate !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.preprocess.dropBoilerplate", next),
                        }),
                    ),
                    h(Row, { label: "重复行去重", note: "重复出现的行只留第一次。" },
                        h(Toggle, {
                            checked: preprocess.dedupeLines !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.preprocess.dedupeLines", next),
                        }),
                    ),
                    h(Row, { label: "短行阈值", note: "短于这个长度、又不像句子就当导航丢掉；0 = 关闭。" },
                        h(NumberField, {
                            value: preprocess.minLineChars === undefined ? 12 : preprocess.minLineChars,
                            min: 0, max: 80, step: 1, disabled: disabled,
                            onChange: (next) => write("search.preprocess.minLineChars", next),
                        }),
                        h("span", { style: S.unit }, "字"),
                    ),
                    h(Row, { label: "抽出标题与出处", note: "标题 / 作者 / 时间放进预处理报告，不进正文。" },
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
                        "先把查询限定到清单里的站点，命中排前并按类型汇总。"),
                    h(Row, {
                        label: "站点优先检索",
                        note: "关掉则所有检索都不限定来源。",
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
                        note: "限定站无结果时退回泛搜。",
                    },
                        h(Toggle, {
                            checked: sites.fallback !== false,
                            disabled: disabled,
                            onChange: (next) => write("search.sites.fallback", next),
                        }),
                    ),
                    h(Row, {
                        label: "单次最多限定",
                        note: "限定越多越慢，也越容易被限流。",
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
                        note: "失焦或回车写回；说明只在这里显示。",
                        stacked: true,
                    },
                        h(SiteCatalogEditor, {
                            entries: sites.entries,
                            disabled: disabled,
                            write: write,
                        }),
                    ),
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
                    // 取正文与 PDF 的上限搬进了上面的「高级参数」折叠块 —— 它们与通道无关，
                    // 放在通道参数旁边要写清楚「换通道不影响这些」。
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
                    h(Row, { label: "Provider 路由名", note: "留空 = 只覆盖 model。" },
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
                    h(Row, { label: "模型 id", note: "留空 = 跟随；写错会在派工时明确报错。" },
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
                    h(Row, { label: "默认主题", note: "plain = 素色网格，无背景填充。", first: true },
                        h(Select, {
                            value: documents.defaultTheme,
                            disabled: disabled,
                            options: THEMES.map((theme) => ({ value: theme.id, label: theme.label })),
                            onChange: (next) => write("documents.defaultTheme", next),
                        }),
                    ),
                    h(Row, { label: "脚本超时", note: "单次 office_run 脚本的超时。" },
                        h(NumberField, {
                            value: documents.scriptTimeoutMs, min: 1000, max: 600000, step: 1000, disabled: disabled,
                            onChange: (next) => write("documents.scriptTimeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "ms"),
                    ),
                    h(Row, { label: "脚本字符上限", note: "单次脚本的字符上限。" },
                        h(NumberField, {
                            value: documents.maxScriptChars, min: 1000, max: 4000000, step: 1000, disabled: disabled,
                            onChange: (next) => write("documents.maxScriptChars", next),
                        }),
                        h("span", { style: S.unit }, "字"),
                    ),
                    h(Row, { label: "缓存目录", note: "中间产物目录，按存活时间清理。" },
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
                    h(Row, { label: "保留中间产物", note: "关掉则每次调用结束清空。" },
                        h(Toggle, {
                            checked: documents.keepCache !== false,
                            disabled: disabled,
                            onChange: (next) => write("documents.keepCache", next),
                        }),
                    ),
                    h(Row, { label: "缓存存活时间", note: "多少分钟没碰就清掉；0 = 不按时间清。" },
                        h(NumberField, {
                            value: documents.cacheTtlMinutes === undefined ? 720 : documents.cacheTtlMinutes,
                            min: 0, max: 43200, step: 10, disabled: disabled,
                            onChange: (next) => write("documents.cacheTtlMinutes", next),
                        }),
                        h("span", { style: S.unit }, "分钟"),
                    ),

                    // ── PDF 渲染 ──
                    h(Row, { label: "PDF 渲染分辨率", note: "PDF 转图的 DPI；越大越慢。" },
                        h(NumberField, {
                            value: documents.pdfDpi === undefined ? 120 : documents.pdfDpi,
                            min: 36, max: 400, step: 1, disabled: disabled,
                            onChange: (next) => write("documents.pdfDpi", next),
                        }),
                        h("span", { style: S.unit }, "dpi"),
                    ),
                    h(Row, { label: "PDF 单次页数上限", note: "单次最多渲染多少页。" },
                        h(NumberField, {
                            value: documents.pdfMaxPages === undefined ? 20 : documents.pdfMaxPages,
                            min: 1, max: 500, step: 1, disabled: disabled,
                            onChange: (next) => write("documents.pdfMaxPages", next),
                        }),
                        h("span", { style: S.unit }, "页"),
                    ),
                    h(Row, { label: "PDF 渲染引擎", note: "一个都探测不到时 office.pdf.pages() 会明确报错。" },
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
                    h(Row, { label: "LaTeX 编译入口", note: "thuthesis 需 xelatex / lualatex，不能用 pdflatex。" },
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
                    h(Row, { label: "LaTeX 编译超时", note: "学位论文要跑多遍，目录长时调大。" },
                        h(NumberField, {
                            value: documents.texTimeoutMs === undefined ? 300000 : documents.texTimeoutMs,
                            min: 30000, max: 1800000, step: 10000, disabled: disabled,
                            onChange: (next) => write("documents.texTimeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "毫秒"),
                    ),
                    h(Row, { label: "thuthesis 模板目录", note: "留空自动探测工作目录与 TeX Live。" },
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

                // ── 文件识别（OCR）──
                //
                // 服务端 `ocr` 组（src/settings.js:604-627）一直存在，界面却没有这一组：
                // 缺 Key 的报错把人指到「设置 → 办公模式 → 文件识别（OCR）」，
                // 打开设置却找不到能填 Key 的地方 —— 这一组就是那个落点。
                //
                // 写法由宿主的脱敏契约决定（dsh-api-settings-controller/lib/index.js:88-89、
                // :292-297；dsh-settings/lib/index.js:73-89）：
                //   1) role('secret') 字段在每次远程读里被物理删掉，configForms 的快照也没有
                //      secrets 旁路 —— `ocr.apiKey` 到这里永远是空/缺席。所以密码框**不从设置
                //      值回填**（回填等于把明文塞进 DOM），永远是空的写入框。
                //   2) 写入是 path op 的 set：写空串等于拿空覆盖已存的 Key。所以只有非空
                //      输入才提交；换 Key = 粘新的，清空要改配置文件（这里不提供「清除」）。
                h(SettingsGroup, { key: "ocr", id: "ocr" },
                    h("div", { style: S.groupTitle }, "文件识别（OCR）"),
                    h("div", { style: S.groupNote },
                        "扫描件 / 图片 / PDF 交给 DeepSeek OCR；官方 1.5 元 / 1000 页，Key 与对话模型凭据是两份。"),

                    h(Row, { label: "启用", note: "关掉后「识别文件」提交会报错，提示到设置里打开。", first: true },
                        h(Toggle, {
                            checked: ocr.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("ocr.enabled", next),
                        }),
                    ),
                    h(Row, { label: "服务端点", note: "带不带 /v1 都认，插件补 /v1/tools/ocr。" },
                        h("input", {
                            type: "text",
                            defaultValue: ocr.baseURL || "",
                            disabled: disabled,
                            placeholder: "留空 = https://api.deepseek.com",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== String(ocr.baseURL || "")) write("ocr.baseURL", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "300px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, {
                        label: "DeepSeek API Key",
                        note: "只写不读，留空不改；官方端点按专用 Key > 环境变量 > 复用取凭据。",
                    },
                        h("input", {
                            type: "password",
                            // 永远从空开始：服务端读回来的 secret 字段已被删掉，回填只会是
                            // 空；万一某个部署没脱敏，回填等于把明文写进 DOM。
                            defaultValue: "",
                            disabled: disabled,
                            placeholder: "粘 Key（留空 = 不改动已存的 / 用下面的环境变量）",
                            // 不用 "off"：那是「请勿自动填充」的提示，浏览器常常照样填；
                            // "new-password" 才是「这是新密码，别用已存的填」。
                            autoComplete: "new-password",
                            spellCheck: false,
                            "data-ocr-key": "true",
                            onBlur: (event) => {
                                const target = event.target;
                                const next = String(target.value || "").trim();
                                // 立刻清空本地输入：这个框是不受控的，留着值等于把 Key 留在 DOM 里，
                                // 下一次失焦还会把同一份 Key 再提交一遍。提交完即清，重进设置也是空的。
                                target.value = "";
                                // 空输入不写：set 是覆盖语义，写空会把已存的 Key 抹掉。
                                if (next !== "") write("ocr.apiKey", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "300px", textAlign: "left" }),
                        }),
                    ),
                    h(Row, { label: "Key 的环境变量名", note: "只存变量名；改完要重启 DSH。" },
                        h("input", {
                            type: "text",
                            defaultValue: ocr.apiKeyEnv || "",
                            disabled: disabled,
                            placeholder: "留空 = DEEPSEEK_API_KEY",
                            onBlur: (event) => {
                                const next = String(event.target.value || "").trim();
                                if (next !== String(ocr.apiKeyEnv || "")) write("ocr.apiKeyEnv", next);
                            },
                            onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                            style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                        }),
                    ),
                    // 复用由宿主保存的 deepseek-official 凭据：必须显式打开，默认关。
                    // 原始 /office-memory/ocr 路由没有可认证的网关对端，默认替用户借宿主凭据
                    // 等于悄悄花宿主的钱；所以「专用 Key > 环境变量 > 复用」里的最后一级
                    // 只能由用户在这行亲手打开，且后端也只在自己的官方端点上认它。
                    h(Row, {
                        label: "复用 deepseek-official",
                        // 文案只讲清能力边界，不承诺「任何网关都进不来」：来源检查不是身份认证，
                        // 本机程序与会剥掉转发头的代理都能绕过；用户要按「可信本机环境」自己判断。
                        note: "默认关，只给可信的本机环境用。仅官方端点、且专用 Key 与环境变量都没有时，才借用宿主的 deepseek-official。它只校验请求来源，不是身份认证，本机程序或代理仍可绕过。",
                    },
                        h(Toggle, {
                            checked: ocr.reuseOfficialProvider === true,
                            disabled: disabled,
                            onChange: (next) => write("ocr.reuseOfficialProvider", next),
                            // 只读检查用：确认这个开关在页面上（默认关），不必点击。
                            attrs: { "data-ocr-reuse-official": "true" },
                        }),
                    ),
                    h(Row, { label: "识别精度", note: "文字 = 只取文字；文字 + 版面 = 带段落与表格位置。" },
                        h(Select, {
                            value: ocr.purpose === undefined ? "text_recognition" : ocr.purpose,
                            disabled: disabled,
                            options: [
                                { value: "text_recognition", label: "文字" },
                                { value: "rich_layout_recognition", label: "文字 + 版面" },
                            ],
                            onChange: (next) => write("ocr.purpose", next),
                        }),
                    ),
                    h(Row, { label: "识别超时", note: "10–300 秒；官方端点 180 秒封顶。" },
                        h(NumberField, {
                            value: ocr.timeoutMs === undefined ? 180000 : ocr.timeoutMs,
                            min: 10000, max: 300000, step: 5000, disabled: disabled,
                            onChange: (next) => write("ocr.timeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "毫秒"),
                    ),
                    h("div", { style: S.callout },
                        h("div", { style: S.rowNote },
                            "官方端点（https://api.deepseek.com）：专用 Key > 环境变量 > 复用已保存的 deepseek-official 凭据"
                            + "（默认关，需显式开启），通常不必重新输入；自定义端点永远不借用这份凭据，必须填 Key 或指定环境变量。"),
                        h("div", { style: S.rowNote },
                            "Key 只交给宿主进程；保存走已登录的写入通道，不经过诊断接口、不回显，插件本身不带 Key。"),
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
                        "脚本里用 office.python 做计算与绘图，产物落缓存目录；不装包。"),
                    h(Row, { label: "启用", note: "关掉后 office.python 不可用，调用会报错并指向设置。", first: true },
                        h(Toggle, {
                            checked: python.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("python.enabled", next),
                        }),
                    ),
                    h(Row, { label: "解释器路径", note: "留空自动探测 python / python3 / py。" },
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
                    h(Row, { label: "运行超时", note: "单次运行超时；长计算调大。" },
                        h(NumberField, {
                            value: python.timeoutMs === undefined ? 120000 : python.timeoutMs,
                            min: 5000, max: 900000, step: 5000, disabled: disabled,
                            onChange: (next) => write("python.timeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "毫秒"),
                    ),
                    h(Row, { label: "产物目录", note: "相对缓存目录；留空 = python/out。" },
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
                        "本机 SenseVoice 转写与视频抽帧；缺 ffmpeg、模型或运行时报错。"),
                    h(Row, { label: "启用", note: "关掉后 office.av 不可用，调用会报错并指向设置。", first: true },
                        h(Toggle, {
                            checked: av.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("av.enabled", next),
                        }),
                    ),
                    h(Row, { label: "ffmpeg 路径", note: "留空自动探测 PATH 与常见目录。" },
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
                    h(Row, { label: "ffprobe 路径", note: "留空先找 ffmpeg 同目录，再走 PATH。" },
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
                    h(Row, { label: "SenseVoice 模型目录", note: "留空 = 语音输入下载的那份。" },
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
                    h(Row, { label: "转写语言", note: "auto 让模型判断；固定语言可能翻错混语种。" },
                        h(Select, {
                            value: av.language === undefined ? "auto" : av.language,
                            disabled: disabled,
                            options: AV_LANGUAGE_OPTIONS,
                            onChange: (next) => write("av.language", next),
                        }),
                    ),
                    h(Row, { label: "单块秒数", note: "模型一次吃的音频长度；默认 120 秒。" },
                        h(NumberField, {
                            value: av.chunkSeconds === undefined ? 120 : av.chunkSeconds,
                            min: 10, max: 600, step: 10, disabled: disabled,
                            onChange: (next) => write("av.chunkSeconds", next),
                        }),
                        h("span", { style: S.unit }, "秒"),
                    ),
                    h(Row, { label: "块间重叠", note: "相邻块重叠；默认 0.5 秒，上限半块。" },
                        h(NumberField, {
                            value: av.overlapSeconds === undefined ? 0.5 : av.overlapSeconds,
                            min: 0, max: 30, step: 0.5, disabled: disabled,
                            onChange: (next) => write("av.overlapSeconds", next),
                        }),
                        h("span", { style: S.unit }, "秒"),
                    ),
                    h(Row, { label: "单次时长上限", note: "超过要求先切段；切块不改变这条上限。" },
                        h(NumberField, {
                            value: av.maxSeconds === undefined ? 3600 : av.maxSeconds,
                            min: 10, max: 21600, step: 10, disabled: disabled,
                            onChange: (next) => write("av.maxSeconds", next),
                        }),
                        h("span", { style: S.unit }, "秒"),
                    ),
                    h(Row, { label: "处理超时", note: "解码 / 抽帧 / 转写的单次超时。" },
                        h(NumberField, {
                            value: av.timeoutMs === undefined ? 300000 : av.timeoutMs,
                            min: 5000, max: 1800000, step: 5000, disabled: disabled,
                            onChange: (next) => write("av.timeoutMs", next),
                        }),
                        h("span", { style: S.unit }, "毫秒"),
                    ),
                    h(Row, { label: "模型精度", note: "int8 约 228 MB；fp32 更大更慢。" },
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
                    h(Row, { label: "推理线程", note: "默认 2；再多收益递减。" },
                        h(NumberField, {
                            value: av.threads === undefined ? 2 : av.threads,
                            min: 1, max: 16, step: 1, disabled: disabled,
                            onChange: (next) => write("av.threads", next),
                        }),
                        h("span", { style: S.unit }, "线程"),
                    ),
                    h(Row, { label: "默认抽帧张数", note: "默认抽几张，交给 read_image。" },
                        h(NumberField, {
                            value: av.frames === undefined ? 6 : av.frames,
                            min: 1, max: 60, step: 1, disabled: disabled,
                            onChange: (next) => write("av.frames", next),
                        }),
                        h("span", { style: S.unit }, "张"),
                    ),
                    h(Row, { label: "单次抽帧上限", note: "防止一次抽太多撑爆缓存。" },
                        h(NumberField, {
                            value: av.maxFrames === undefined ? 12 : av.maxFrames,
                            min: 1, max: 60, step: 1, disabled: disabled,
                            onChange: (next) => write("av.maxFrames", next),
                        }),
                        h("span", { style: S.unit }, "张"),
                    ),
                    h(Row, { label: "抽帧格式", note: "jpg 体积小；png 文字更清楚。" },
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
                    h(Row, { label: "抽帧宽度", note: "0 = 原分辨率；4K 建议 1280。" },
                        h(NumberField, {
                            value: av.frameWidth === undefined ? 0 : av.frameWidth,
                            min: 0, max: 4096, step: 2, disabled: disabled,
                            onChange: (next) => write("av.frameWidth", next),
                        }),
                        h("span", { style: S.unit }, "px"),
                    ),
                ),

                h("div", { key: "tail", style: S.groupNote },
                    "改动即时生效；越界值会被拒绝并回读。"),
            );
            } });
        }

        /**
         * 知识库大模型的 provider / 模型选择（第九十四轮）。
         *
         * 这两行原来是纯手填：用户得先知道宿主里真实的 provider id，再猜它下面
         * 有哪些模型 id；填错只会在入库时收到一句失败回执。上传工具栏早就拉
         * /office-memory/llm-models 填下拉，设置页却一直没接这条只读清单。
         *
         * 规则：
         * - provider 与 model 从**同一份**宿主清单里选；切 provider 时在**同一次
         *   mutate** 里把 model 清空 —— 分两次写会在中间留下「新 provider + 旧
         *   model」的坏组合，第二笔还可能撞 revision 冲突。
         * - 已保存的值始终保留为选项 / 手填兜底；只有**成功拿到非空清单**时才
         *   标注它不在清单里 —— 首帧、加载中、读取失败、空表都没有可核对的东西。
         * - 清单来源：登录后优先 ctx.remote.session.modelCatalog()（宿主自己的
         *   Models 设置页用的同一个 RPC，见 dsh-client-ui-settings-models 的
         *   store.ts:268）；只有旧宿主没有这个 session 方法时才退回原来的
         *   HTTP 只读接口。RPC 拒绝 / 抛错原样显示并停下重试，尤其权限拒绝
         *   **不能**自动改走 HTTP 绕过授权；partial failures 也照实列出。
         * - 清单读取失败原样显示宿主 error（宿主已带前缀，界面不再叠一层）；重试
         *   只重发同一来源的那一次只读调用。
         * - 每次取数带请求号：host → openai → host、连点重试时只让最新响应落状态。
         * 只读宿主清单、不带 Key、不调模型，不扩大隐私面；openai 通道仍手填，
         * 因为宿主列不出远端端点上的模型。
         */
        /**
         * 宿主 remote.session 上的模型清单方法（登录后可用，不需要 cwd）。
         * 每次取数时现读 hostRemote，而不是挂载时捕获 —— remote 服务可能晚于
         * 组件就绪，捕获一次会让迟绑定永远走旧路径。
         * 缺这个方法 = 旧宿主，才退回原来的 HTTP 只读接口。
         */
        function sessionModelCatalog() {
            const session = sessionRemote;
            if (session === null || typeof session.modelCatalog !== "function") return null;
            return () => session.modelCatalog();
        }

        /**
         * 旧宿主明确的「没有这个方法」能力错误（dsh-api-gateway/src/types.ts:124-145
         * 的 TypertGatewayErrorCode，服务端 index.ts:906-908 在没有任何 Remote
         * 方法导出该 endpoint 时抛出）。**只**认这一个码：权限拒绝、网络错误、
         * gateway/internal 都绝不能借此改走 HTTP。
         *
         * 客户端会静态挂载当前版本的方法描述，所以旧宿主上 session.modelCatalog
         * 仍是函数，typeof 检查不够；这个码才是真正区分「宿主没实现」的依据。
         */
        function modelCatalogMethodUnavailable(error) {
            return error !== null && typeof error === "object"
                && error.code === "gateway/invocation-unavailable";
        }

        /** 宿主 RPC 拒绝 / 抛错时的人话；宿主没给原因也要说清楚。 */
        function modelCatalogErrorText(error) {
            const code = error && typeof error.code === "string" ? error.code : "";
            const message = error && typeof error.message === "string" ? error.message : "";
            const detail = message || "宿主没有给出原因";
            const text = code === "" ? detail : code + "：" + detail;
            return text.indexOf("读取宿主模型清单失败") === 0 ? text : "读取宿主模型清单失败：" + text;
        }

        /**
         * 把 ModelCatalog（dsh-api-session-controller/src/types.ts:110-141）投影成
         * 与旧 HTTP 接口同形的 { providers, failures }：
         * groups[].id -> provider，groups[].models[].name -> 模型显示名。
         * failures 原样保留（不吞掉部分 provider 读取失败）。
         */
        function catalogFromModelCatalog(value) {
            if (value === null || typeof value !== "object" || !Array.isArray(value.groups)) {
                throw new Error("宿主模型清单响应格式不对（modelCatalog 没有返回 groups）。");
            }
            const providers = [];
            for (const group of value.groups) {
                if (group === null || typeof group !== "object") continue;
                const id = String(group.id || "");
                if (id === "") continue;
                const models = [];
                for (const item of Array.isArray(group.models) ? group.models : []) {
                    if (item === null || typeof item !== "object") continue;
                    const modelId = String(item.id || "");
                    if (modelId === "") continue;
                    models.push({ id: modelId, label: String(item.name || modelId) });
                }
                providers.push({ provider: id, models: models });
            }
            const failures = [];
            for (const item of Array.isArray(value.failures) ? value.failures : []) {
                if (item === null || typeof item !== "object") continue;
                failures.push({
                    id: String(item.id || ""),
                    name: String(item.name || item.id || ""),
                    message: String(item.message || ""),
                });
            }
            return { providers: providers, failures: failures };
        }

        function KbLlmModelPicker(props) {
            const llm = props.llm && typeof props.llm === "object" ? props.llm : {};
            // 同一个宿主模型选择器也给图库的视觉通道用：写入前缀可换，
            // 默认仍是知识库大模型的 memory.llm，老调用方不用改。
            const prefix = typeof props.prefix === "string" && props.prefix !== "" ? props.prefix : "memory.llm";
            // 两个已知前缀保留完整字面量：源码覆盖测试逐项核对「界面上每个 schema
            // 参数都有写回路径」，动态拼接的 prefix + ".provider" 会被它当成缺失。
            const llmPath = (key) => {
                if (prefix === "gallery.vision") return key === "provider" ? "gallery.vision.provider" : "gallery.vision.model";
                if (prefix === "memory.llm") return key === "provider" ? "memory.llm.provider" : "memory.llm.model";
                return prefix + "." + key;
            };
            const mode = String(llm.mode === undefined ? "off" : llm.mode);
            const provider = String(llm.provider || "");
            const model = String(llm.model || "");
            const [catalog, setCatalog] = React.useState(null);
            const [loading, setLoading] = React.useState(false);
            const [error, setError] = React.useState("");
            const alive = React.useRef(true);
            // 每次取数领一个号；只有「仍然存活且号是最新」的响应才允许落状态。
            // host → openai → host、连点重试都会留下在途请求，先到的旧响应不能
            // 盖掉后到的清单，也不能把旧错误贴到新清单上。
            const requestId = React.useRef(0);
            const load = () => {
                requestId.current += 1;
                const id = requestId.current;
                // 新请求未决期间旧清单不能再用来判「缺失」或画选项；失败时也由
                // catch 保持 null，回到手填兜底。
                setCatalog(null);
                setLoading(true);
                setError("");
                const settle = (thenable) => Promise.resolve(thenable)
                    .then((next) => {
                        if (alive.current !== true || id !== requestId.current) return null;
                        setCatalog(next);
                        setLoading(false);
                        return next;
                    })
                    .catch((err) => {
                        if (alive.current !== true || id !== requestId.current) return null;
                        setLoading(false);
                        // 失败后不能拿上一份清单继续判「缺失」；清掉，回到手填兜底。
                        setCatalog(null);
                        setError(String(err && err.message ? err.message : err));
                        return null;
                    });
                let viaRpc = null;
                try {
                    viaRpc = sessionModelCatalog();
                } catch (err) {
                    setLoading(false);
                    setError(String(err && err.message ? err.message : err));
                    return Promise.resolve(null);
                }
                // 旧宿主 / 旧兼容端点的只读 HTTP 清单；两条路径共用同一份投影。
                const loadViaHttp = () => {
                    if (typeof fetch !== "function") {
                        setLoading(false);
                        setError("当前环境没有 fetch，读不到宿主模型清单。");
                        return Promise.resolve(null);
                    }
                    return Promise.resolve(fetch(LLM_MODELS_PATH, { headers: { accept: "application/json" } }))
                        .then((response) => Promise.resolve(response.json()).then(
                            (body) => ({ response: response, body: body }),
                            () => ({ response: response, body: null }),
                        ))
                        .then((result) => {
                            const status = result.response && result.response.status ? result.response.status : "?";
                            const body = result.body;
                            if (body === null || typeof body !== "object") {
                                throw new Error("宿主模型清单响应格式不对（HTTP " + status + "）。");
                            }
                            if (body.ok !== true) {
                                // 宿主 error 已经自带「读取宿主模型清单失败：」这类前缀，
                                // 直接显示，界面上不再叠一层同样的话。
                                throw new Error(String(body.error || ("宿主返回 HTTP " + status)));
                            }
                            return {
                                providers: Array.isArray(body.providers) ? body.providers : [],
                                failures: [],
                            };
                        });
                };
                if (viaRpc !== null) {
                    // 登录后的首选路径：不依赖 cwd、不发 HTTP。拒绝 / 抛错就原样
                    // 显示并停在重试，**不**偷偷退回 HTTP（权限拒绝尤其不能绕）。
                    return settle(Promise.resolve().then(() => viaRpc()).catch((err) => {
                        // 调用本身被拒绝 / 抛错时也要说清是「读模型清单失败」，但不改走
                        // HTTP：权限拒绝、网络错误都停在这里显示原文。
                        throw new Error(modelCatalogErrorText(err));
                    }).then((result) => {
                        if (result === null || typeof result !== "object") {
                            throw new Error("宿主模型清单响应格式不对（modelCatalog 没有返回结果）。");
                        }
                        if (result.ok !== true) {
                            // 唯一例外：旧宿主明确没有导出这个方法。此时退回插件
                            // 自己的只读 HTTP 接口（仍不带 Key、不调模型）。
                            if (modelCatalogMethodUnavailable(result.error)) return loadViaHttp();
                            throw new Error(modelCatalogErrorText(result.error));
                        }
                        return catalogFromModelCatalog(result.value);
                    }));
                }
                // 旧宿主没有 session.modelCatalog：才退回原来的只读 HTTP 接口。
                return settle(loadViaHttp());
            };
            // 只在 host 通道取清单：off / openai 用不到它，没必要为打开设置页多打
            // 一次请求。依赖 mode 是为了「从 off 切到 host」时补一次取数。
            React.useEffect(() => {
                alive.current = true;
                // 换通道先作废在途请求（host → openai → host 时旧响应可能晚到）。
                requestId.current += 1;
                if (mode === "host") {
                    load();
                } else {
                    setLoading(false);
                    setError("");
                    // 离开 host 就把清单丢掉：切回来时新请求未决期间不能拿旧表判缺失。
                    setCatalog(null);
                }
                return () => {
                    alive.current = false;
                    requestId.current += 1;
                };
            }, [mode]);
            // remote 迟绑定 / 重新注入时刷新：清单方法刚出现，下一次取数就自动
            // 走 RPC；仍然缺这个方法时 load 自己退回 HTTP。回调里用当前 mode 判断，
            // 非 host 通道不该为一次服务就绪多打请求。
            React.useEffect(() => subscribeHostRemote(() => {
                if (mode === "host") load();
            }), [mode]);
            const manualInput = (path, value, width, note) => h("input", {
                type: "text", defaultValue: value, disabled: props.disabled,
                placeholder: note,
                "data-llm-manual": path,
                onBlur: (event) => {
                    const next = String(event.target.value || "").trim();
                    if (next !== value) props.write(path, next);
                },
                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                style: Object.assign({}, S.input, { width: width, textAlign: "left" }),
            });
            // off / openai：保留手填（openai 的模型名只能手填，宿主列不出远端清单）。
            if (mode !== "host") {
                return [
                    h(Row, { key: "llm-provider", label: "provider", note: "host 通道填宿主里的 provider id；openai 通道留空" },
                        manualInput(llmPath("provider"), provider, "200px", "provider id")),
                    h(Row, { key: "llm-model", label: "模型", note: "host 通道从宿主模型列表选；openai 通道手填，如 gpt-4o-mini" },
                        manualInput(llmPath("model"), model, "220px", "模型名")),
                ];
            }
            // host：按真实清单画下拉。providers 只认非空 id；同一 provider 重复出现
            // 时后者覆盖不了前者，但宿主当前契约里不会重复，这里只做防御。
            const providers = [];
            if (catalog !== null && Array.isArray(catalog.providers)) {
                for (const item of catalog.providers) {
                    if (item === null || typeof item !== "object") continue;
                    const id = String(item.provider || item.id || "");
                    if (id === "") continue;
                    providers.push({ id: id, models: Array.isArray(item.models) ? item.models : [] });
                }
            }
            // partial failure 的 provider 只是「没读到」，不能当成「不存在」：
            // 已保存值落在失败名单里时保留原值，也不贴「不在清单里」。
            const failedProviders = new Set();
            if (catalog !== null && Array.isArray(catalog.failures)) {
                for (const item of catalog.failures) {
                    if (item === null || typeof item !== "object") continue;
                    const id = String(item.id || "");
                    if (id !== "") failedProviders.add(id);
                }
            }
            const chosen = providers.find((item) => item.id === provider) || null;
            const providerOptions = [{ value: "", label: "请选择 provider" }];
            for (const item of providers) {
                providerOptions.push({ value: item.id, label: item.id });
            }
            // 已保存的值先保留成选项，select 的 value 才总能落在 options 里。
            // 标签只写原值：清单没回来时还不能说它「不在清单里」。
            if (provider !== "" && chosen === null) {
                providerOptions.push({ value: provider, label: provider });
            }
            // 只有成功拿到、且非空的清单才谈得上「核对」；空表不报缺失。
            const catalogReady = catalog !== null && loading !== true && error === "";
            const savedProviderMissing = catalogReady && providers.length > 0
                && provider !== "" && chosen === null && !failedProviders.has(provider);
            const modelOptions = [{ value: "", label: "请选择模型" }];
            const listed = [];
            if (chosen !== null) {
                for (const item of chosen.models) {
                    if (item === null || typeof item !== "object") continue;
                    const id = String(item.id || "");
                    if (id === "") continue;
                    listed.push(id);
                    modelOptions.push({ value: id, label: String(item.label || id) });
                }
            }
            // provider 未知、或已知但没列出模型：模型退回手填。此时没有可核对的
            // 模型清单，既不标「模型缺失」，也不叠加第二条提示。
            const modelManual = provider !== "" && (chosen === null || chosen.models.length === 0);
            const savedModelMissing = catalogReady && chosen !== null && chosen.models.length > 0
                && model !== "" && listed.indexOf(model) === -1 && !failedProviders.has(provider);
            if (!modelManual && model !== "" && listed.indexOf(model) === -1) {
                modelOptions.push({ value: model, label: model });
            }
            const modelControl = modelManual
                ? manualInput(llmPath("model"), model, "240px", "模型 id")
                : h(Select, {
                    value: model,
                    disabled: props.disabled === true || provider === "",
                    options: modelOptions,
                    onChange: (next) => {
                        if (next === model) return;
                        props.write(llmPath("model"), next);
                    },
                });
            const notes = [];
            if (loading) {
                notes.push(h("div", { key: "loading", style: S.status, "data-llm-catalog-status": "loading" }, "正在读取宿主模型清单…"));
            }
            if (error !== "") {
                // error 里要么是宿主原文（已自带「读取宿主模型清单失败：」），
                // 要么是上面自己生成的完整句子；这里原样显示，不再加前缀。
                notes.push(h("div", { key: "error", style: S.error, "data-llm-catalog-error": "true" },
                    error,
                    h("button", {
                        type: "button", className: "om-btn",
                        style: Object.assign({}, S.ghostButton, { marginLeft: "8px" }),
                        "data-llm-catalog-retry": "true",
                        onClick: () => { load(); },
                    }, "重试")));
            }
            if (catalogReady && providers.length === 0) {
                notes.push(h("div", { key: "empty", style: S.status, "data-llm-catalog-empty": "true" },
                    "宿主清单里没有可用的 provider：请先在宿主里配置模型，或改用 OpenAI 兼容端点。"));
            }
            // RPC 的 partial failures：部分 provider 没读到就照实列出来，不能因为
            // 其它 provider 成功就把这几条悄悄吞掉。
            if (catalogReady && catalog.failures.length > 0) {
                const failureText = catalog.failures.map((item) => {
                    const label = item.name || item.id || "未知 provider";
                    return label + "：" + (item.message || "宿主没有给出原因");
                }).join("；");
                notes.push(h("div", { key: "failures", style: S.status, "data-llm-catalog-failures": "true" },
                    "部分 provider 的模型清单没读到：" + failureText));
            }
            if (savedProviderMissing) {
                notes.push(h("div", { key: "saved-provider", style: S.status, "data-llm-saved-provider": "true" },
                    "已保存的 provider「" + provider + "」不在宿主当前清单里：保留原值，仍按它入库。"));
            }
            if (savedModelMissing) {
                notes.push(h("div", { key: "saved-model", style: S.status, "data-llm-saved-model": "true" },
                    "已保存的模型「" + model + "」不在当前 provider 的清单里：保留原值，仍按它入库。"));
            }
            return [
                h(Row, { key: "llm-provider", label: "provider", note: "从宿主已配置的 provider 里选；切换会同时清空模型" },
                    h(Select, {
                        value: provider,
                        disabled: props.disabled === true,
                        options: providerOptions,
                        onChange: (next) => {
                            if (next === provider) return;
                            // 一次 mutate：provider 与清空后的 model 一起落盘。
                            props.writeMany([[llmPath("provider"), next], [llmPath("model"), ""]]);
                        },
                    })),
                h(Row, { key: "llm-model", label: "模型", note: modelManual
                    ? "该 provider 没有列出模型：手填模型 id"
                    : "只列当前 provider 的模型" },
                    modelControl),
                notes.length === 0 ? null : h("div", { key: "llm-notes" }, notes),
            ];
        }

        // ── 记忆系统面板 ──────────────────────────────────────────────────
        //
        // 独立的一级导航项：原来那个「记忆系统」入口由 mnemon 提供，本插件现在自带
        // 三层记忆，就由这一页接住。面板只做两件事：说清三层是什么、调节范围
        // （开关 / 目录 / 容量 / 检索）。一次性迁移不再占设置页，命令仍在 memory 文档里。

        /** 非空字符串（设置值可能缺字段 / 是数字，这里统一成文本）。 */
        function settingText(value) {
            return String(value === undefined || value === null ? "" : value).trim();
        }

        /** 远端 embedding 是否真的配齐：端点 + 模型（Key 可为空，自建端点常免鉴权）。 */
        function remoteEmbeddingReady(embedding) {
            const baseURL = settingText(embedding && embedding.baseURL);
            const model = settingText(embedding && embedding.model);
            return baseURL !== "" && model !== "";
        }

        /** 候选筛选是否真的能调：llm + 端点 + 模型，缺一项都不发请求。 */
        function candidateSelectorReady(selector) {
            return settingText(selector && selector.mode) === "llm"
                && settingText(selector && selector.baseURL) !== ""
                && settingText(selector && selector.model) !== "";
        }

        /**
         * 顶部「检索方式」：按设置值现算**实际会走哪条路**，而不是只回显下拉值。
         *
         * 服务端口径（src/memory.js hybridKbSearch）：off 纯词法；local 本地推理起不来
         * 才退本地哈希；remote 端点/模型不齐就退回词法；auto 先本地推理、再远端，
         * 两者都没有才是纯词法。以前这里只看 embedding.mode，开了本地推理仍显示
         * 「自动混合」，用户看到的和实际跑的检索不是一回事。
         */
        function retrievalStatusOf(memory, embedding, local, selector) {
            const mode = settingText(embedding.mode) || "auto";
            const localOn = local.enabled === true;
            const remoteOn = remoteEmbeddingReady(embedding);
            const selectorOn = candidateSelectorReady(selector);
            const suffix = (text) => {
                if (selectorOn) return text + "；再由筛选模型保留前几条";
                if (settingText(selector.mode) === "llm") {
                    return text + "；筛选未配齐（端点与模型必填），本轮不筛选";
                }
                return text;
            };
            if (memory.enabled === false) {
                return { route: "disabled", value: "已关闭", note: "记忆工具与知识库检索一起停用。" };
            }
            if (mode === "off") {
                return { route: "off", value: "纯词法", note: "向量化通道已关：只走 BM25 词法检索。" };
            }
            if (mode === "local") {
                if (localOn) {
                    return { route: "local", value: "本地推理", note: suffix("本地 Python 模型按需启动，空闲退出。") };
                }
                return { route: "local-hash", value: "本地哈希", note: "本地推理未开：显式 local 用零依赖哈希向量，能补字面相近，不是语义模型。" };
            }
            if (mode === "remote") {
                if (remoteOn) {
                    return { route: "remote", value: "远端向量", note: suffix("端点与模型已配齐；请求失败会退回词法并标 degraded。") };
                }
                return { route: "none", value: "纯词法", note: "显式 remote 但端点或模型没填齐，退回词法。" };
            }
            if (localOn) {
                return { route: "local", value: "本地推理", note: suffix("自动：本地推理已开启，优先用它。") };
            }
            if (remoteOn) {
                return { route: "remote", value: "远端向量", note: suffix("自动：本地推理未开，改用已配置的远端端点。") };
            }
            return { route: "none", value: "纯词法", note: "自动：本地推理未开、远端端点/模型未配齐，只走词法。" };
        }

        /**
         * 图库页签：视觉描述与图文对齐。两件事都默认关，开了才会把图片或查询发出去。
         * 以前 schema 里有这两组配置，但合并设置页时没有入口，用户只能看到
         * 「视觉描述未配置」的报错。这里把入口补上，并让 off 状态只留一句说明。
         */
        function GallerySettingsSection(props) {
            const pageGroups = Array.isArray(props.groups) ? props.groups : GALLERY_GROUP_IDS;
            const pageKey = typeof props.page === "string" ? props.page : "gallery";
            return h(SettingsPanel, { scope: props.scope, render: ({ value, disabled, write, writeMany }) => {
                const gallery = value.gallery || {};
                const vision = gallery.vision || {};
                const image = gallery.imageEmbedding || {};
                const channel = settingText(vision.channel) || "off";
                const imageProvider = settingText(image.provider) || "none";
                const text = (path, current, width, placeholder, marker) => h("input", Object.assign({
                    type: "text",
                    defaultValue: current,
                    disabled: disabled,
                    placeholder: placeholder || "",
                    onBlur: (event) => {
                        const next = String(event.target.value || "").trim();
                        if (next !== current) write(path, next);
                    },
                    onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                    style: Object.assign({}, S.input, { width: width, textAlign: "left" }),
                }, marker ? { [marker]: "true" } : {}));
                const secret = (path, current, width, placeholder, marker) => h("input", Object.assign({
                    type: "password",
                    defaultValue: current,
                    disabled: disabled,
                    placeholder: placeholder || "",
                    onBlur: (event) => {
                        const next = String(event.target.value || "").trim();
                        if (next !== current) write(path, next);
                    },
                    onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                    style: Object.assign({}, S.input, { width: width, textAlign: "left" }),
                }, marker ? { [marker]: "true" } : {}));
                return h(SettingsBody, {
                    page: pageKey,
                    groups: pageGroups,
                    overview: null,
                    defaultOpen: true,
                },
                    h(SettingsGroup, { key: "gallery", id: "gallery" },
                        h("div", { style: S.groupTitle }, "图库视觉描述"),
                        h("div", { style: S.groupNote },
                            "默认关：只保存图片与手工标签，不生成描述，也不把图片发出去。"),
                        h(Row, { label: "通道", note: "off 不生成；host 用宿主已配置的多模态模型；openai 用兼容端点。", first: true },
                            h(Select, {
                                value: channel,
                                disabled: disabled,
                                "data-gallery-vision-channel": "true",
                                options: [
                                    { value: "off", label: "关（默认）" },
                                    { value: "host", label: "宿主已配置模型" },
                                    { value: "openai", label: "OpenAI 兼容端点" },
                                ],
                                onChange: (next) => write("gallery.vision.channel", next),
                            }),
                        ),
                        channel === "off"
                            ? h("div", { style: S.callout, "data-gallery-vision-off": "true" },
                                h("div", { style: T.strong }, "未启用：生成视觉描述会提示「未配置」"),
                                h("div", { style: S.rowNote }, "把上面的通道切到 host 或 openai，再点图库详情里的「生成视觉描述」。"),
                            )
                            : h("div", { "data-gallery-vision-on": channel },
                                channel === "host"
                                    ? h(KbLlmModelPicker, {
                                        llm: { mode: "host", provider: vision.provider, model: vision.model },
                                        prefix: "gallery.vision",
                                        disabled: disabled,
                                        write: write,
                                        writeMany: writeMany,
                                    })
                                    : [
                                        h(Row, { key: "gallery-vision-endpoint", label: "端点", note: "到 /v1 这一层；填了它就会把缩略图发到这个地址。" },
                                            text("gallery.vision.baseURL", vision.baseURL, "280px", "https://api.example.com/v1", "data-gallery-vision-base-url")),
                                        h(Row, { key: "gallery-vision-model", label: "模型", note: "视觉模型 id，例如 gpt-4o-mini。" },
                                            text("gallery.vision.model", vision.model, "220px", "gpt-4o-mini", "data-gallery-vision-model")),
                                        h(Row, { key: "gallery-vision-key-env", label: "视觉 Key 环境变量", note: "取 Key 的环境变量名；和下面的 Key 都没有时明确报未配置。" },
                                            text("gallery.vision.apiKeyEnv", vision.apiKeyEnv, "220px", "VISION_API_KEY", "data-gallery-vision-key-env")),
                                        h(Row, { key: "gallery-vision-key", label: "视觉 Key", note: "一般留空用环境变量；只用于请求，不会写进回执。" },
                                            secret("gallery.vision.apiKey", vision.apiKey, "220px", "sk-…", "data-gallery-vision-key")),
                                    ],
                                channel === "off" ? null : h(Row, { label: "输出上限", note: "单次描述最多多少 token。" },
                                    h(NumberField, {
                                        value: vision.maxTokens === undefined ? 900 : vision.maxTokens,
                                        min: 200, max: 2000, step: 100, disabled: disabled,
                                        onChange: (next) => write("gallery.vision.maxTokens", next),
                                    }),
                                ),
                                channel === "off" ? null : h(Row, { label: "单次超时", note: "超时按失败回执，不写占位描述。" },
                                    h(NumberField, {
                                        value: vision.timeoutMs === undefined ? 60000 : vision.timeoutMs,
                                        min: 5000, max: 180000, step: 1000, disabled: disabled,
                                        onChange: (next) => write("gallery.vision.timeoutMs", next),
                                    }),
                                    h("span", { style: S.unit }, "毫秒"),
                                ),
                                channel === "off" ? null : h(Row, { label: "描述语言", note: "只影响模型写出的描述文字。" },
                                    h(Select, {
                                        value: vision.language === "en" ? "en" : "zh",
                                        disabled: disabled,
                                        options: [{ value: "zh", label: "中文" }, { value: "en", label: "英文" }],
                                        onChange: (next) => write("gallery.vision.language", next),
                                    }),
                                ),
                            ),
                    ),

                    h(SettingsGroup, { key: "gallery-image", id: "gallery-image" },
                        h("div", { style: S.groupTitle }, "图库图文对齐（可选）"),
                        h("div", { style: S.groupNote },
                            "把图片与文字放进同一个向量空间，用于文字搜图与以图搜图；默认不做。"),
                        h(Row, { label: "通道", note: "none 不做图片向量；voyage / cohere 用官方多模态接口；vllm 填任意兼容端点。", first: true },
                            h(Select, {
                                value: imageProvider,
                                disabled: disabled,
                                "data-gallery-image-provider": "true",
                                options: [
                                    { value: "none", label: "none（默认）" },
                                    { value: "voyage", label: "Voyage" },
                                    { value: "cohere", label: "Cohere" },
                                    { value: "vllm", label: "OpenAI 兼容多模态端点" },
                                ],
                                onChange: (next) => write("gallery.imageEmbedding.provider", next),
                            }),
                        ),
                        imageProvider === "none"
                            ? h("div", { style: S.callout, "data-gallery-image-off": "true" },
                                h("div", { style: T.strong }, "未启用：不做图片向量"),
                                h("div", { style: S.rowNote }, "文字搜图与以图搜图会跳过图文信号；标签、分类与描述仍可用。"),
                            )
                            : h("div", { "data-gallery-image-on": imageProvider },
                                h(Row, { label: "端点", note: imageProvider === "vllm" ? "必填：到 /v1 这一层。" : "留空用官方端点；自建代理可填。" },
                                    text("gallery.imageEmbedding.baseURL", image.baseURL, "280px", imageProvider === "vllm" ? "http://127.0.0.1:8000/v1" : "留空 = 官方端点", "data-gallery-image-base-url")),
                                h(Row, { label: "模型", note: imageProvider === "vllm" ? "必填：多模态 embedding 模型 id。" : "留空用官方默认模型。" },
                                    text("gallery.imageEmbedding.model", image.model, "240px", "voyage-multimodal-3.5", "data-gallery-image-model")),
                                h(Row, { label: "向量 Key 环境变量", note: "取 Key 的环境变量名；自建端点通常不需要。" },
                                    text("gallery.imageEmbedding.apiKeyEnv", image.apiKeyEnv, "220px", "VOYAGE_API_KEY", "data-gallery-image-key-env")),
                                h(Row, { label: "向量 Key", note: "一般留空用环境变量；只用于请求。" },
                                    secret("gallery.imageEmbedding.apiKey", image.apiKey, "220px", "sk-…", "data-gallery-image-key")),
                                h(Row, { label: "期望维度", note: "0 = 用模型默认；与文本向量不一致时明确报维度不匹配。" },
                                    h(NumberField, {
                                        value: image.dimensions === undefined ? 0 : image.dimensions,
                                        min: 0, max: 4096, step: 1, disabled: disabled,
                                        onChange: (next) => write("gallery.imageEmbedding.dimensions", next),
                                    }),
                                ),
                                h(Row, { label: "批大小", note: "一次请求最多几张图。" },
                                    h(NumberField, {
                                        value: image.batchSize === undefined ? 4 : image.batchSize,
                                        min: 1, max: 16, step: 1, disabled: disabled,
                                        onChange: (next) => write("gallery.imageEmbedding.batchSize", next),
                                    }),
                                ),
                                h(Row, { label: "单次超时", note: "超时按失败回执，不写占位向量。" },
                                    h(NumberField, {
                                        value: image.timeoutMs === undefined ? 30000 : image.timeoutMs,
                                        min: 5000, max: 180000, step: 1000, disabled: disabled,
                                        onChange: (next) => write("gallery.imageEmbedding.timeoutMs", next),
                                    }),
                                    h("span", { style: S.unit }, "毫秒"),
                                ),
                            ),
                    ),
                );
            } });
        }

        function MemorySettingsSection(props) {
            const pageGroups = Array.isArray(props.groups) ? props.groups : MEMORY_GROUP_IDS;
            const pageKey = typeof props.page === "string" ? props.page : "memory";
            return h(SettingsPanel, { scope: props.scope, render: ({ value, disabled, write, writeMany, writeError }) => {
            const memory = value.memory || {};
            // 三个嵌套组各自取出来：它们不是扁平键，写回也必须用点号路径
            // （write("memory.layers.hot", …)），所以读的时候也要走同一层。
            const layers = memory.layers || {};
            const recallQuality = memory.recallQuality || {};
            const quota = memory.quota || {};
            const embedding = memory.embedding || {};
            const local = embedding.local || {};
            const selector = memory.selector || {};
            const ann = memory.ann || {};
            const hybrid = memory.hybrid || {};
            const llm = memory.llm || {};
            // 通道决定哪些参数真的会被用到：没有向量化时，远端/本地/筛选/索引
            // 四组都只是「配了也不生效」，设置页要把这层关系直接画出来，
            // 而不是把三十行输入框一起铺开。
            const embeddingMode = settingText(embedding.mode) || "auto";
            const embeddingModeText = embeddingMode === "auto" ? "自动"
                : (embeddingMode === "remote" ? "远端" : (embeddingMode === "local" ? "本地" : "关闭"));
            const localOn = local.enabled === true;
            const remoteDisabled = disabled || embeddingMode === "off" || embeddingMode === "local";
            const localDisabled = disabled || embeddingMode === "off" || embeddingMode === "remote";
            const localParamsDisabled = localDisabled || !localOn;
            const selectorMode = settingText(selector.mode) || "off";
            const selectorDisabled = disabled || embeddingMode === "off";
            const selectorParamsDisabled = selectorDisabled || selectorMode !== "llm";
            const annOn = ann.enabled !== false;
            const annDisabled = disabled || embeddingMode === "off";
            const annParamsDisabled = annDisabled || !annOn;
            const llmMode = settingText(llm.mode) || "off";
            const layerNote = "热记忆常驻，台账按次登记，归档只读。";
            // 仪表盘：数字都从设置值里现算，不写死 —— 改了下拉/开关，顶部跟着变。
            const layerOn = LAYER_ROWS.filter((row) => layers[row.key] !== false).length;
            const scopeText = memory.scope === "global" ? "全局共用"
                : (memory.scope === "both" ? "全局 + 工作区" : "工作区");
            const retrieval = retrievalStatusOf(memory, embedding, local, selector);
            const memoryOverview = {
                kicker: "MEMORY CONTROL",
                title: "记忆系统 · 设置总览",
                note: "三层记忆、知识库与检索策略共用一套参数；下面是分组卡片。",
                facts: [
                    { key: "layers", label: "开启的层", value: layerOn + " / " + LAYER_ROWS.length, note: "热记忆 / 台账 / 归档" },
                    { key: "scope", label: "记忆范围", value: scopeText, note: "跨项目层的落点" },
                    { key: "retrieval", label: "检索方式", value: retrieval.value, note: retrieval.note },
                ],
            };

            return h(SettingsBody, {
                page: pageKey, groups: pageGroups,
                overview: props.hideOverview === true ? null : memoryOverview,
                defaultOpen: props.defaultOpen === true,
            },
                writeError !== null ? h("div", { key: "error", style: S.error }, writeError) : null,
                disabled
                    ? h("div", { key: "readonly", style: S.status }, "当前连接以只读方式同步设置（memory 模式），改动不会持久化。")
                    : null,

                h(SettingsGroup, { key: "memory", id: "memory" },
                    h("div", { style: S.groupTitle }, "记忆系统"),
                    h("div", { style: S.groupNote }, layerNote),
                    h(Row, { label: "开启记忆", note: "关掉后 office_memory、记忆提示段与自动台账一起停。", first: true },
                        h(Toggle, {
                            checked: memory.enabled !== false,
                            disabled: disabled,
                            onChange: (next) => write("memory.enabled", next),
                        }),
                    ),
                    h(Row, { label: "记忆目录", note: "换目录就是换一套记忆。" },
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
                    h(Row, { label: "自动登记台账", note: "office_run 写出文档时自动登记；关掉只能手工 log。" },
                        h(Toggle, {
                            checked: memory.autoLedger !== false,
                            disabled: disabled,
                            onChange: (next) => write("memory.autoLedger", next),
                        }),
                    ),
                    h(Row, { label: "写入系统提示", note: "办公模式已有 persona，这段通常会被丢掉。" },
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
                        "关掉只是不读不写，数据留在盘上。"),
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
                        "让偏好跨工作目录复用；项目事实仍落工作区。"),
                    h(Row, { label: "记忆放哪儿", note: "默认 workspace；both = 全局偏好 + 项目事实并存", first: true },
                        h(Select, {
                            value: memory.scope === undefined ? "workspace" : memory.scope,
                            disabled: disabled,
                            options: MEMORY_SCOPE_OPTIONS,
                            onChange: (next) => write("memory.scope", next),
                        }),
                    ),
                    h(Row, { label: "全局层目录", note: "留空 = $DSH_HOME/.office/memory；相对用户主目录。" },
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
                    h(Row, { label: "用户偏好范围", note: "偏好建议 global；项目事实只落工作区。" },
                        h(Select, {
                            value: memory.userScope === undefined ? "memory" : memory.userScope,
                            disabled: disabled,
                            options: USER_SCOPE_OPTIONS,
                            onChange: (next) => write("memory.userScope", next),
                        }),
                    ),
                ),

                // ── 知识库（第八十五轮） ──
                h(SettingsGroup, { key: "kb", id: "kb" },
                    h("div", { style: S.groupTitle }, "知识库"),
                    h("div", { style: S.groupNote },
                        "manifest 是真源；SQLite 与索引都可重建。"),
                    h(KbDirRow, { value: memory.kbDir, disabled: disabled, write: write }),
                    h(KbDirRow, {
                        value: memory.galleryDir,
                        disabled: disabled,
                        write: write,
                        settingKey: "memory.galleryDir",
                        inputName: "data-gallery-dir-input",
                        pickName: "data-gallery-dir-pick",
                        blankText: "知识库目录旁的 gallery",
                        label: "图库目录",
                        note: "留空 = 知识库目录旁的 gallery。",
                        placeholder: "留空 = 知识库目录旁的 gallery",
                        first: false,
                    }),
                    h(Row, { label: "向量化通道", note: "auto 本地优先，其次远端；都没有才词法。" },
                        h(Select, {
                            value: embedding.mode === undefined ? "auto" : embedding.mode,
                            disabled: disabled,
                            options: EMBEDDING_MODE_OPTIONS,
                            onChange: (next) => write("memory.embedding.mode", next),
                        }),
                    ),
                    h("div", { style: S.callout, "data-kb-route": retrieval.route },
                        h("div", { style: T.strong }, "当前实际通道：" + retrieval.value),
                        h("div", { style: S.rowNote }, retrieval.note),
                    ),

                    // 高级参数按通道收进折叠块：内容照渲染（定位与单测仍找得到），
                    // 但默认只看到「通道 + 开关」，不再一屏三十行输入框。
                    h(Fold, { name: "kb-remote-params", label: "远端向量参数（端点 / 模型 / 批大小）" },
                        remoteDisabled
                            ? h("div", { style: S.status }, "当前通道是" + embeddingModeText + "：远端参数不生效，切到 auto / remote 后可编辑。")
                            : null,
                        h(Row, { label: "embedding 端点", note: "OpenAI 兼容 /embeddings，例如 https://api.siliconflow.cn/v1。" },
                            h("input", {
                                type: "text",
                                defaultValue: embedding.baseURL,
                                disabled: remoteDisabled,
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== embedding.baseURL) write("memory.embedding.baseURL", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "embedding 模型", note: "例如 BAAI/bge-m3；auto 下留空跳过远端。" },
                            h("input", {
                                type: "text",
                                defaultValue: embedding.model,
                                disabled: remoteDisabled,
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== embedding.model) write("memory.embedding.model", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "220px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "向量 Key 环境变量", note: "留空且无字面量 Key 时按无鉴权端点处理。" },
                            h("input", {
                                type: "text",
                                defaultValue: embedding.apiKeyEnv,
                                disabled: remoteDisabled,
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== embedding.apiKeyEnv) write("memory.embedding.apiKeyEnv", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "向量维度", note: "0 = 由模型决定。" },
                            h(NumberField, {
                                value: embedding.dimensions === undefined ? 0 : embedding.dimensions,
                                min: 0, max: 4096, step: 1, disabled: remoteDisabled,
                                onChange: (next) => write("memory.embedding.dimensions", next),
                            }),
                        ),
                        h(Row, { label: "批量大小", note: "一次请求送多少个块。" },
                            h(NumberField, {
                                value: embedding.batchSize === undefined ? 16 : embedding.batchSize,
                                min: 1, max: 64, step: 1, disabled: remoteDisabled,
                                onChange: (next) => write("memory.embedding.batchSize", next),
                            }),
                        ),
                        h(Row, { label: "embedding 超时", note: "超时或报错退回词法。" },
                            h(NumberField, {
                                value: embedding.timeoutMs === undefined ? 15000 : embedding.timeoutMs,
                                min: 1000, max: 120000, step: 1000, disabled: remoteDisabled,
                                onChange: (next) => write("memory.embedding.timeoutMs", next),
                            }),
                            h("span", { style: S.unit }, "毫秒"),
                        ),
                    ),
                    // ── 本地推理（第一百零三轮） ──
                    // 默认关：首次打开要有模型目录与 Python 才能用，缺什么在回执与状态里写明。
                    h(Row, { label: "本地推理服务", note: "第一次检索才起进程，空闲自动退出。" },
                        h(Toggle, {
                            checked: localOn,
                            disabled: localDisabled,
                            onChange: (next) => write("memory.embedding.local.enabled", next),
                        }),
                    ),
                    h(Fold, { name: "kb-models", label: "本地模型文件（下载 / 状态）" },
                        h(KbModelsPanel, { kbDir: memory.kbDir, disabled: disabled, poll: localOn }),
                    ),
                    h(Fold, { name: "kb-local-params", label: "本地推理参数（模型 / 进程 / 内存）" },
                        localParamsDisabled
                            ? h("div", { style: S.status }, localOn
                                ? "当前通道是" + embeddingModeText + "：本地参数不生效，切到 auto / local 后可编辑。"
                                : "本地推理未开启：先勾选上面的「本地推理服务」再调这些参数。")
                            : null,
                        h(Row, { label: "本地模型目录", note: "留空 = <知识库目录>/models/bge-small-zh-v1.5；要有模型与 tokenizer。" },
                            h("input", {
                                type: "text",
                                defaultValue: local.modelDir,
                                disabled: localParamsDisabled,
                                placeholder: "留空 = <知识库目录>/models/bge-small-zh-v1.5",
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== local.modelDir) write("memory.embedding.local.modelDir", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "Python 命令", note: "留空自动找 python / py / python3。" },
                            h("input", {
                                type: "text",
                                defaultValue: local.command,
                                disabled: localParamsDisabled,
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== local.command) write("memory.embedding.local.command", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "240px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "空闲退出", note: "多久没检索就结束进程树。" },
                            h(NumberField, {
                                value: local.idleTimeoutMs === undefined ? 120000 : local.idleTimeoutMs,
                                min: 5000, max: 3600000, step: 1000, disabled: localParamsDisabled,
                                onChange: (next) => write("memory.embedding.local.idleTimeoutMs", next),
                            }),
                            h("span", { style: S.unit }, "毫秒"),
                        ),
                        h(Row, { label: "内存上限", note: "超过就杀进程并退回词法。" },
                            h(NumberField, {
                                value: local.maxRssMb === undefined ? 600 : local.maxRssMb,
                                min: 200, max: 4096, step: 10, disabled: localParamsDisabled,
                                onChange: (next) => write("memory.embedding.local.maxRssMb", next),
                            }),
                            h("span", { style: S.unit }, "MB"),
                        ),
                        h(Row, { label: "推理线程", note: "2 线程实测快约 6 倍。" },
                            h(NumberField, {
                                value: local.threads === undefined ? 2 : local.threads,
                                min: 1, max: 16, step: 1, disabled: localParamsDisabled,
                                onChange: (next) => write("memory.embedding.local.threads", next),
                            }),
                        ),
                        h(Row, { label: "启动超时", note: "这么久没就绪就判失败。" },
                            h(NumberField, {
                                value: local.startupTimeoutMs === undefined ? 60000 : local.startupTimeoutMs,
                                min: 5000, max: 300000, step: 1000, disabled: localParamsDisabled,
                                onChange: (next) => write("memory.embedding.local.startupTimeoutMs", next),
                            }),
                            h("span", { style: S.unit }, "毫秒"),
                        ),
                        h(Row, { label: "最大 token", note: "单块最长 token；256 是安全档。" },
                            h(NumberField, {
                                value: local.maxLength === undefined ? 256 : local.maxLength,
                                min: 64, max: 512, step: 32, disabled: localParamsDisabled,
                                onChange: (next) => write("memory.embedding.local.maxLength", next),
                            }),
                        ),
                    ),

                    // ── 候选筛选（第一百零三轮替代重排） ──
                    h(Row, { label: "候选筛选", note: "llm = 只把候选编号与块头发给侧路模型，回编号。" },
                        h(Select, {
                            value: selector.mode === undefined ? "off" : selector.mode,
                            disabled: selectorDisabled,
                            options: SELECTOR_MODE_OPTIONS,
                            onChange: (next) => write("memory.selector.mode", next),
                        }),
                    ),
                    h(Fold, { name: "kb-selector-params", label: "筛选参数（端点 / 保留条数 / 超时）" },
                        selectorDisabled
                            ? h("div", { style: S.status }, "向量化已关：没有候选可筛，切回任一向量通道后可配置。")
                            : (selectorMode !== "llm"
                                ? h("div", { style: S.status }, "候选筛选未开：先把上面的「候选筛选」切到 llm，这些参数才会发请求。")
                                : null),
                        h(Row, { label: "筛选端点", note: "OpenAI 兼容 /chat/completions；空 = 不启用。" },
                            h("input", {
                                type: "text",
                                defaultValue: selector.baseURL,
                                disabled: selectorParamsDisabled,
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== selector.baseURL) write("memory.selector.baseURL", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "筛选模型", note: "留空 = 不生效。" },
                            h("input", {
                                type: "text",
                                defaultValue: selector.model,
                                disabled: selectorParamsDisabled,
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== selector.model) write("memory.selector.model", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "220px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "筛选 Key 环境变量", note: "留空且无字面量 Key 时按无鉴权处理。" },
                            h("input", {
                                type: "text",
                                defaultValue: selector.apiKeyEnv,
                                disabled: selectorParamsDisabled,
                                onBlur: (event) => {
                                    const next = String(event.target.value || "").trim();
                                    if (next !== selector.apiKeyEnv) write("memory.selector.apiKeyEnv", next);
                                },
                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                            }),
                        ),
                        h(Row, { label: "筛选候选数", note: "送进筛选的候选条数（2–30）" },
                            h(NumberField, {
                                value: selector.maxCandidates === undefined ? 12 : selector.maxCandidates,
                                min: 2, max: 30, step: 1, disabled: selectorParamsDisabled,
                                onChange: (next) => write("memory.selector.maxCandidates", next),
                            }),
                        ),
                        h(Row, { label: "最多保留", note: "筛选后最多留下几条（1–20）；只按编号过滤，不重排" },
                            h(NumberField, {
                                value: selector.maxSelected === undefined ? 5 : selector.maxSelected,
                                min: 1, max: 20, step: 1, disabled: selectorParamsDisabled,
                                onChange: (next) => write("memory.selector.maxSelected", next),
                            }),
                        ),
                        h(Row, { label: "块头字数", note: "只发块头，不发全文。" },
                            h(NumberField, {
                                value: selector.headChars === undefined ? 300 : selector.headChars,
                                min: 80, max: 2000, step: 20, disabled: selectorParamsDisabled,
                                onChange: (next) => write("memory.selector.headChars", next),
                            }),
                        ),
                        h(Row, { label: "筛选超时", note: "超时回退融合结果。" },
                            h(NumberField, {
                                value: selector.timeoutMs === undefined ? 12000 : selector.timeoutMs,
                                min: 1000, max: 120000, step: 1000, disabled: selectorParamsDisabled,
                                onChange: (next) => write("memory.selector.timeoutMs", next),
                            }),
                            h("span", { style: S.unit }, "毫秒"),
                        ),
                    ),

                    // ── ANN 索引（第一百零三轮） ──
                    h(Row, { label: "向量索引", note: "缺失 / 损坏自动重建；超规模退回精确扫描。" },
                        h(Toggle, {
                            checked: annOn,
                            disabled: annDisabled,
                            onChange: (next) => write("memory.ann.enabled", next),
                        }),
                    ),
                    h(Fold, { name: "kb-ann-params", label: "向量索引参数（efSearch / 邻居 / 建图）" },
                        annDisabled
                            ? h("div", { style: S.status }, "向量化已关：索引不参与检索，参数不生效。")
                            : (!annOn
                                ? h("div", { style: S.status }, "向量索引已关：检索走精确扫描，这些参数不生效。")
                                : null),
                        h(Row, { label: "索引 efSearch", note: "越大越准越慢；默认 64。" },
                            h(NumberField, {
                                value: ann.efSearch === undefined ? 64 : ann.efSearch,
                                min: 8, max: 512, step: 8, disabled: annParamsDisabled,
                                onChange: (next) => write("memory.ann.efSearch", next),
                            }),
                        ),
                        h(Row, { label: "邻居数 M", note: "建图时定死，改了会重建。" },
                            h(NumberField, {
                                value: ann.m === undefined ? 16 : ann.m,
                                min: 4, max: 64, step: 4, disabled: annParamsDisabled,
                                onChange: (next) => write("memory.ann.m", next),
                            }),
                        ),
                        h(Row, { label: "建图宽度", note: "越大图越好、建得越慢。" },
                            h(NumberField, {
                                value: ann.efConstruction === undefined ? 200 : ann.efConstruction,
                                min: 16, max: 1000, step: 8, disabled: annParamsDisabled,
                                onChange: (next) => write("memory.ann.efConstruction", next),
                            }),
                        ),
                        h(Row, { label: "建图上限", note: "超过就不建图，退回精确扫描。" },
                            h(NumberField, {
                                value: ann.maxBuild === undefined ? 20000 : ann.maxBuild,
                                min: 100, max: 20000, step: 100, disabled: annParamsDisabled,
                                onChange: (next) => write("memory.ann.maxBuild", next),
                            }),
                            h("span", { style: S.unit }, "块"),
                        ),
                    ),
                    h(Fold, { name: "kb-hybrid", label: "融合权重与候选池" },
                        h(Row, { label: "词法权重", note: "BM25 / FTS 的融合权重。" },
                            h(NumberField, {
                                value: hybrid.lexicalWeight === undefined ? 0.55 : hybrid.lexicalWeight,
                                min: 0, max: 1, step: 0.05, disabled: disabled,
                                onChange: (next) => write("memory.hybrid.lexicalWeight", next),
                            }),
                        ),
                        h(Row, { label: "向量权重", note: "纯词法时不起作用。" },
                            h(NumberField, {
                                value: hybrid.vectorWeight === undefined ? 0.45 : hybrid.vectorWeight,
                                min: 0, max: 1, step: 0.05, disabled: disabled || embeddingMode === "off",
                                onChange: (next) => write("memory.hybrid.vectorWeight", next),
                            }),
                        ),
                        h(Row, { label: "候选池", note: "越大越稳也越慢。" },
                            h(NumberField, {
                                value: hybrid.candidateLimit === undefined ? 30 : hybrid.candidateLimit,
                                min: 5, max: 100, step: 5, disabled: disabled,
                                onChange: (next) => write("memory.hybrid.candidateLimit", next),
                            }),
                            h("span", { style: S.unit }, "个"),
                        ),
                    ),
                ),

                // ── 知识库大模型（可选，默认关） ──
                h(SettingsGroup, { key: "kb-llm", id: "kb-llm" },
                    h("div", { style: S.groupTitle }, "知识库大模型（可选）"),
                    h("div", { style: S.groupNote },
                        "默认关；打开后模型只决定切片边界与摘要，不改正文。"),
                    h(Row, { label: "通道", note: "off 不用模型；host 用宿主模型；openai 兼容端点。", first: true },
                        h(Select, {
                            value: llm.mode === undefined ? "off" : llm.mode,
                            disabled: disabled,
                            options: [
                                { value: "off", label: "关（默认）" },
                                { value: "host", label: "宿主已配置模型" },
                                { value: "openai", label: "OpenAI 兼容端点" },
                            ],
                            onChange: (next) => write("memory.llm.mode", next),
                        }),
                    ),
                    llmMode === "off"
                        ? h("div", { style: S.callout, "data-kb-llm-off": "true" },
                            h("div", { style: T.strong }, "未启用：知识库只用本地规则切片"),
                            h("div", { style: S.rowNote }, "不联网、不调用模型；需要模型时把上面的「通道」切到 host 或 openai。"),
                        )
                        : h("div", { "data-kb-llm-on": "true" },
                            h(Row, { label: "模型决定切片", note: "模型只决定并块；正文不改写。" },
                                h(Toggle, { checked: llm.chunking === true, disabled: disabled, onChange: (next) => write("memory.llm.chunking", next) }),
                            ),
                            h(Row, { label: "生成摘要", note: "摘要另存；命中会带原文块 id。" },
                                h(Toggle, { checked: llm.summary === true, disabled: disabled, onChange: (next) => write("memory.llm.summary", next) }),
                            ),
                            // provider / 模型：host 通道从宿主真实清单里选（第九十四轮），
                            // 切 provider 与清空 model 在同一笔 mutate 里完成。
                            h(KbLlmModelPicker, { llm: llm, disabled: disabled, write: write, writeMany: writeMany }),
                            h(Fold, { name: "kb-llm-params", label: "连接与预算（端点 / 超时 / 调用上限）" },
                                llmMode === "openai"
                                    ? [
                                        h(Row, { key: "llm-endpoint", label: "端点", note: "openai 通道填到 /v1 这一层；填了它就会把原文发到这个地址" },
                                            h("input", {
                                                type: "text", defaultValue: llm.baseURL, disabled: disabled,
                                                onBlur: (event) => {
                                                    const next = String(event.target.value || "").trim();
                                                    if (next !== llm.baseURL) write("memory.llm.baseURL", next);
                                                },
                                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                                style: Object.assign({}, S.input, { width: "260px", textAlign: "left" }),
                                            }),
                                        ),
                                        h(Row, { key: "llm-key", label: "大模型 Key", note: "一般留空用环境变量；只用于请求，不会写进回执" },
                                            h("input", {
                                                type: "password", defaultValue: llm.apiKey, disabled: disabled,
                                                onBlur: (event) => {
                                                    const next = String(event.target.value || "").trim();
                                                    if (next !== llm.apiKey) write("memory.llm.apiKey", next);
                                                },
                                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                                style: Object.assign({}, S.input, { width: "220px", textAlign: "left" }),
                                            }),
                                        ),
                                        h(Row, { key: "llm-key-env", label: "大模型 Key 环境变量", note: "取 Key 的环境变量名；两处都没有时明确报错" },
                                            h("input", {
                                                type: "text", defaultValue: llm.apiKeyEnv, disabled: disabled,
                                                onBlur: (event) => {
                                                    const next = String(event.target.value || "").trim();
                                                    if (next !== llm.apiKeyEnv) write("memory.llm.apiKeyEnv", next);
                                                },
                                                onKeyDown: (event) => { if (event.key === "Enter") event.target.blur(); },
                                                style: Object.assign({}, S.input, { width: "200px", textAlign: "left" }),
                                            }),
                                        ),
                                    ]
                                    : h("div", { style: S.status }, "host 使用宿主端点与凭据。"),
                                h(Row, { label: "单次超时", note: "超时按规则回退。" },
                                    h(NumberField, { value: llm.timeoutMs === undefined ? 30000 : llm.timeoutMs, min: 1000, max: 120000, step: 1000, disabled: disabled, onChange: (next) => write("memory.llm.timeoutMs", next) }),
                                    h("span", { style: S.unit }, "毫秒"),
                                ),
                                h(Row, { label: "单次原文上限", note: "按章分批，不整本一次塞。" },
                                    h(NumberField, { value: llm.maxInputChars === undefined ? 6000 : llm.maxInputChars, min: 1000, max: 60000, step: 500, disabled: disabled, onChange: (next) => write("memory.llm.maxInputChars", next) }),
                                    h("span", { style: S.unit }, "字符"),
                                ),
                                h(Row, { label: "输出上限", note: "单次输出 token 上限。" },
                                    h(NumberField, { value: llm.maxOutputTokens === undefined ? 1500 : llm.maxOutputTokens, min: 200, max: 8000, step: 100, disabled: disabled, onChange: (next) => write("memory.llm.maxOutputTokens", next) }),
                                ),
                                h(Row, { label: "单次最多调用", note: "到顶暂停并保留进度。" },
                                    h(NumberField, { value: llm.maxCallsPerRun === undefined ? 20 : llm.maxCallsPerRun, min: 1, max: 100, step: 1, disabled: disabled, onChange: (next) => write("memory.llm.maxCallsPerRun", next) }),
                                    h("span", { style: S.unit }, "次"),
                                ),
                                h(Row, { label: "总预算", note: "超出直接拒绝，不会先烧钱。" },
                                    h(NumberField, { value: llm.maxCalls === undefined ? 400 : llm.maxCalls, min: 1, max: 5000, step: 1, disabled: disabled, onChange: (next) => write("memory.llm.maxCalls", next) }),
                                    h("span", { style: S.unit }, "次"),
                                ),
                                h(Row, { label: "块上限", note: "超过再按小节切。" },
                                    h(NumberField, { value: llm.maxChunkChars === undefined ? 3000 : llm.maxChunkChars, min: 500, max: 10000, step: 100, disabled: disabled, onChange: (next) => write("memory.llm.maxChunkChars", next) }),
                                    h("span", { style: S.unit }, "字符"),
                                ),
                                h(Row, { label: "块下限", note: "太碎的段并进来，不跨章。" },
                                    h(NumberField, { value: llm.minChunkChars === undefined ? 400 : llm.minChunkChars, min: 100, max: 3000, step: 50, disabled: disabled, onChange: (next) => write("memory.llm.minChunkChars", next) }),
                                    h("span", { style: S.unit }, "字符"),
                                ),
                                h(Row, { label: "每批小节数", note: "防止编号太多模型数错。" },
                                    h(NumberField, { value: llm.maxUnitsPerBatch === undefined ? 60 : llm.maxUnitsPerBatch, min: 10, max: 200, step: 10, disabled: disabled, onChange: (next) => write("memory.llm.maxUnitsPerBatch", next) }),
                                    h("span", { style: S.unit }, "个"),
                                ),
                            ),
                        ),
                ),

                // ── 主动记录 ──
                h(SettingsGroup, { key: "capture", id: "capture" },
                    h("div", { style: S.groupTitle }, "主动记录"),
                    h("div", { style: S.groupNote },
                        "只在工具反馈里加一句提示。"),
                    h(Row, { label: "主动记录", note: "提示当场记一条；只是指引，不自主记录。", first: true },
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
                        "超限下沉到归档，不静默丢。"),
                    h(Row, { label: "用户偏好上限", note: "热记忆「用户偏好」容量，默认 4096 字节。", first: true },
                        h(NumberField, {
                            value: memory.userLimitBytes === undefined ? 4096 : memory.userLimitBytes,
                            min: 512, max: 65536, step: 512, disabled: disabled,
                            onChange: (next) => write("memory.userLimitBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "项目记忆上限", note: "热记忆「项目与环境」容量，默认 10240 字节。" },
                        h(NumberField, {
                            value: memory.projectLimitBytes === undefined ? 10240 : memory.projectLimitBytes,
                            min: 1024, max: 262144, step: 1024, disabled: disabled,
                            onChange: (next) => write("memory.projectLimitBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "投影预算", note: "投影正文字节上限；超出折叠但保留。" },
                        h(NumberField, {
                            value: memory.projectionBudgetBytes === undefined ? 4096 : memory.projectionBudgetBytes,
                            min: 0, max: 131072, step: 1024, disabled: disabled,
                            onChange: (next) => write("memory.projectionBudgetBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "单条上限", note: "单条上限，默认 2800；超过拒绝写入。" },
                        h(NumberField, {
                            value: memory.entryLimitBytes === undefined ? 2800 : memory.entryLimitBytes,
                            min: 0, max: 131072, step: 100, disabled: disabled,
                            onChange: (next) => write("memory.entryLimitBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "单条建议线", note: "超过只在回执里提醒，默认 1200。" },
                        h(NumberField, {
                            value: memory.entryHintBytes === undefined ? 1200 : memory.entryHintBytes,
                            min: 0, max: 131072, step: 100, disabled: disabled,
                            onChange: (next) => write("memory.entryHintBytes", next),
                        }),
                        h("span", { style: S.unit }, "字节"),
                    ),
                    h(Row, { label: "台账条数上限", note: "超出的最旧记录滚成月归档。" },
                        h(NumberField, {
                            value: memory.ledgerLimit === undefined ? 500 : memory.ledgerLimit,
                            min: 10, max: 10000, step: 10, disabled: disabled,
                            onChange: (next) => write("memory.ledgerLimit", next),
                        }),
                        h("span", { style: S.unit }, "条"),
                    ),
                    h(Row, { label: "归档保留数", note: "最旧的归档摘要会被删掉。" },
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
                        "分数是命中查询词的比例，不是向量相似度。"),
                    h(Row, { label: "分档策略", first: true },
                        h(Select, {
                            value: recallQuality.policy === undefined ? "strict-v1" : recallQuality.policy,
                            disabled: disabled,
                            options: RECALL_POLICY_OPTIONS,
                            onChange: (next) => write("memory.recallQuality.policy", next),
                        }),
                    ),
                    h(Row, { label: "低分阈值", note: "低于它算「未知」档。" },
                        h(NumberField, {
                            value: recallQuality.lowScoreThreshold === undefined ? 0.25 : recallQuality.lowScoreThreshold,
                            min: 0, max: 1, step: 0.05, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.lowScoreThreshold", next),
                        }),
                    ),
                    h(Row, { label: "高分阈值", note: "达到它算「高分」；必须大于低分阈值。" },
                        h(NumberField, {
                            value: recallQuality.highScoreThreshold === undefined ? 0.6 : recallQuality.highScoreThreshold,
                            min: 0, max: 1, step: 0.05, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.highScoreThreshold", next),
                        }),
                    ),
                    h(Row, { label: "候选倍数", note: "先放大候选池再分档。" },
                        h(NumberField, {
                            value: recallQuality.candidateMultiplier === undefined ? 5 : recallQuality.candidateMultiplier,
                            min: 1, max: 10, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.candidateMultiplier", next),
                        }),
                        h("span", { style: S.unit }, "倍"),
                    ),
                    h(Row, { label: "中档最多几条", note: "中档最多采纳几条。" },
                        h(NumberField, {
                            value: recallQuality.maxMediumResults === undefined ? 6 : recallQuality.maxMediumResults,
                            min: 0, max: 40, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.recallQuality.maxMediumResults", next),
                        }),
                        h("span", { style: S.unit }, "条"),
                    ),
                    h(Row, { label: "未知档最多几条", note: "未知档最多几条；0 = 不要。" },
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
                        "0 = 不限。"),
                    h(Row, { label: "首次检索", note: "每回合第一次记忆检索。", first: true },
                        h(NumberField, {
                            value: quota.recallPerTurn === undefined ? 1 : quota.recallPerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.recallPerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                    h(Row, { label: "换词细化", note: "同一回合的后续细化次数。" },
                        h(NumberField, {
                            value: quota.recallRefinePerTurn === undefined ? 1 : quota.recallRefinePerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.recallRefinePerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                    h(Row, { label: "知识块检索", note: "知识块检索次数，与记忆检索分开算。" },
                        h(NumberField, {
                            value: quota.kbSearchPerTurn === undefined ? 2 : quota.kbSearchPerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.kbSearchPerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                    h(Row, { label: "知识库块关系遍历", note: "块关系跳转次数（kb-related）。" },
                        h(NumberField, {
                            value: quota.relatedPerTurn === undefined ? 1 : quota.relatedPerTurn,
                            min: 0, max: 50, step: 1, disabled: disabled,
                            onChange: (next) => write("memory.quota.relatedPerTurn", next),
                        }),
                        h("span", { style: S.unit }, "次"),
                    ),
                ),

                h("div", { key: "tail", style: S.groupNote },
                    "这里只改设置；记忆内容由 office_memory 读写。"),
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

        /** 导航项：图标 + 名称。办公与记忆已合并成一页。 */
        function OfficeNavLabel() {
            return h("span", { style: { display: "flex", alignItems: "center", gap: "10px" } },
                h(OfficeIcon, null),
                h("span", null, "办公与记忆"),
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
                list.map((text, index) => {
                    const value = String(text);
                    // 来源与重要度是「这条记忆从哪来、会不会跨工作区生效」的两个关键字段，
                    // 给它们一点点颜色；格式 / 日期这类事实保持安静的灰字，不把卡片做成徽标墙。
                    const style = value === "全局" ? Object.assign({}, S.badge, S.badgeOrigin, S.badgeGlobal)
                        : value === "工作区" ? Object.assign({}, S.badge, S.badgeOrigin)
                            : value === "critical" ? Object.assign({}, S.badge, S.badgeCritical)
                                : S.badge;
                    return h("span", { key: index, className: "om-badge", style: style }, value);
                }),
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
         * 找**真正能滚的那一层**祖先（含自身之后的上游）。
         *
         * 为什么要它：面板根元素不再是滚动容器（第六十二轮修滚动，见 S.browser），
         * 滚动权在宿主的会话滚动容器（`.…_scrollBody`）上。对不可滚的盒子调
         * `scrollTo` 是个**静默的空操作** —— 「再显示 N 条」会变成「点了没反应」。
         *
         * 判据只看两件：① overflow-y 是 auto / scroll / overlay；② 真的溢出了
         * （scrollHeight > clientHeight）。拿不到 getComputedStyle（测试替身、
         * 老环境）就一路往上游走、最后返回 null，调用方自己降级。
         */
        function scrollParentOf(node) {
            let current = node ? node.parentElement : null;
            for (let depth = 0; current && depth < 40; depth += 1) {
                let overflowY = "";
                try {
                    const view = typeof window !== "undefined" && window ? window : null;
                    overflowY = view && typeof view.getComputedStyle === "function"
                        ? view.getComputedStyle(current).overflowY
                        : "";
                } catch (error) { overflowY = ""; }
                if ((overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay")
                    && current.scrollHeight > current.clientHeight) return current;
                current = current.parentElement;
            }
            return null;
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
        };

        /** 归档是只读层：给它一句说明，而不是一个点了会报错的动作。 */
        const ARCHIVE_READONLY_NOTE = "归档是只读层：条目由热记忆下沉与台账滚动产生。";

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
            // 第八十五轮：编辑是**直接写入**，与下面那些「生成指令卡」的动作分开；
            // 按钮标记单独一支 data-memory-edit，测试与用户都能分辨这两种动作。
            if (typeof props.onEdit === "function") {
                actions.push(h("button", {
                    key: "edit", type: "button", className: "om-act",
                    style: S.itemAction, "data-memory-edit": String(key),
                    "aria-label": "编辑这一条",
                    onClick: () => props.onEdit(item),
                }, "编辑"));
            }
            // 删除文档（知识库）：直接调用写入通道，不生成指令卡；确认由 deleteCard 负责。
            if (typeof props.onDelete === "function") {
                actions.push(h("button", {
                    key: "delete", type: "button", className: "om-act",
                    style: Object.assign({}, S.itemAction, { color: M.error }), "data-kb-delete": String(key),
                    "aria-label": "删除整份文档",
                    onClick: () => props.onDelete(item),
                }, "删除"));
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
                key: key, style: style, className: "om-card om-row", "data-item-key": String(key),
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
                    key: key, style: style, className: "om-card om-row",
                    "data-importance": importance, "data-item-key": String(key),
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
                // 第八十二轮：文档不是终点 —— 点「看分块」下钻到块清单，再点一块读正文。
                // 这个动作用的是只读端点，和「改写 / 忘记」那种生成待发送指令的按钮分开。
                const openDoc = typeof props.onOpenDoc === "function"
                    ? h("div", { style: S.itemActions, "data-item-actions": "kb-open-" + String(key) },
                        h("button", {
                            type: "button", className: "om-act", style: S.itemAction,
                            "data-kb-open": String(key),
                            "aria-label": "看这份文档切成了哪些块",
                            onClick: () => props.onOpenDoc(item),
                        }, "看分块"))
                    : null;
                return shell(card, [
                    h(Badges, { key: "badges", items: [
                        item.tier,
                        String(item.chunks || 0) + " 块", size, String(item.at || "").slice(0, 10),
                    ] }),
                    body(String(item.path || "") + (item.title ? " —— " + item.title : "")),
                    openDoc,
                    actionRow,
                ]);
            }
        }

        /** 标题下的六枚指标卡（mnemon 健康条那一路的读法：点 + 名称 + 数字）。 */
        function MetricTiles(props) {
            const counts = props.counts || {};
            return h("div", { style: S.tiles, "data-metrics": "memory" }, METRIC_TILES.map((tile) => {
                const value = finite(counts[tile.key]);
                return h("div", { key: tile.key, className: "om-card", style: S.tile, "data-metric": tile.key },
                    // 顶部一条 3px 的层色：六张卡并排时每一张都有自己的锚点，
                    // 不靠改变数字颜色来区分（数字要读得清）。
                    h("div", { style: Object.assign({}, S.tileAccent, { background: tile.color }) }),
                    h("div", { style: S.tileHead },
                        h("span", { style: S.tileLabel }, tile.label),
                        // 状态点带一圈同色光晕：7px 的点因此有了「亮着」的感觉。
                        h("span", {
                            style: Object.assign({}, S.tileDot, {
                                background: tile.color,
                                boxShadow: "0 0 0 3px color-mix(in srgb, " + tile.color + " 14%, transparent)",
                            }),
                        }),
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
                    h("span", { style: S.panelNote }, "用量 / 上限"),
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

        /** 存储域卡片：记忆根各一张，知识库单库一张（第八十八轮）。 */
        const LAYER_STYLE = {
            hot: { shape: "circle", color: "#5a8cff", label: "热记忆" },
            ledger: { shape: "rect", color: "#22a879", label: "台账" },
            archive: { shape: "square", color: "#c08a2e", label: "归档" },
        };
        const LAYER_FALLBACK = { shape: "circle", color: "#8a8f98", label: "条目" };

        function layerStyle(layer) {
            const key = String(layer || "");
            if (key === "hot" || key === "ledger" || key === "archive") return LAYER_STYLE[key];
            return LAYER_FALLBACK;
        }

        /** 重要度对应的左竖线颜色。 */
        const IMPORTANCE_COLORS = {
            critical: "var(--dsw-alias-state-error-primary, #e5484d)",
            normal: "var(--dsw-alias-state-business-primary, #5a8cff)",
            low: "var(--dsw-alias-border-l3, rgba(128,128,128,0.45))",
        };

        const METRIC_TILES = [
            { key: "hot", label: "热记忆", color: "#5a8cff", note: "常驻的偏好与约定" },
            { key: "ledger", label: "台账", color: "#22a879", note: "交付物自动登记" },
            { key: "archive", label: "归档", color: "#c08a2e", note: "下沉的旧条目" },
            { key: "kb", label: "知识库", color: "#3d9bbf", note: "入库的文档与知识块" },
        ];

        function metricColor(key) {
            const found = METRIC_TILES.find((tile) => tile.key === key);
            return found === undefined ? "#8a8f98" : found.color;
        }

        /** 快照里的计数 / 字节一律当可选字段读，缺了给破折号而不是 NaN。 */
        function finite(value) {
            return Number.isFinite(value) ? value : null;
        }

        function countText(value) {
            const number = finite(value);
            return number === null ? "—" : String(number);
        }

        /** 块检索的 reason 是服务端给程序看的（no-overlap / no-docs…），界面要译成人话。 */
        function kbReasonText(reason) {
            const code = String(reason || "");
            if (code === "no-overlap") return "查询词与块内容没有重叠";
            if (code === "no-docs") return "这个范围内没有文档";
            if (code === "filter-empty") return "筛选后没有文档";
            if (code === "no-chunks") return "文档还没有可检索的块";
            return "";
        }

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

        function percentOf(value, limit) {
            const used = finite(value);
            const total = finite(limit);
            if (used === null || total === null || total <= 0) return null;
            return Math.min(100, Math.round((used / total) * 1000) / 10);
        }

        function capacityRows(counts, limits) {
            const source = counts || {};
            const bound = limits || {};
            const rows = [];
            const push = (key, name, value, limit, unit, color) => {
                const percent = percentOf(value, limit);
                if (percent === null) return;
                rows.push({ key: key, name: name, value: finite(value), limit: finite(limit), unit: unit, percent: percent, color: color });
            };
            push("ledger", "台账条数", source.ledgerTotal, bound.ledgerLimit, "条", M.business);
            push("archive", "归档摘要文件", source.archiveFiles, bound.archiveKeep, "个", M.warn);
            push("hotBytes", "热记忆占用", source.hotBytes, bound.userLimitBytes, "字节", M.accent);
            push("projectBytes", "项目记忆占用", source.projectBytes, bound.projectLimitBytes, "字节", M.success);
            return rows;
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

        function StoreCards(props) {
            const stores = Array.isArray(props.stores) ? props.stores : [];
            const kbStore = props.kbStore && typeof props.kbStore === "object" ? props.kbStore : null;
            const kbBytes = kbStore === null ? null : humanBytes(kbStore.bytes);
            const kbCard = kbStore === null ? null : h("div", { style: S.storeCard, "data-store": "kb" },
                h("div", { style: S.storeHead }, h("span", { style: S.storeDir }, "知识库 · 单库")),
                h("div", { style: S.storeRow },
                    h("span", { style: S.layerChip }, "文档 " + countText(kbStore.docs) + " 篇 / " + countText(kbStore.chunks) + " 块"),
                    kbBytes === null ? null : h("span", { style: S.layerChip }, kbBytes),
                    h("span", { style: S.layerChip }, String(kbStore.dir || "（默认数据库目录）"))),
            );
            if (stores.length === 0 && kbCard === null) return null;
            return h("div", { style: S.panel, "data-panel": "stores" },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "存储域"),
                    h("span", { style: S.panelNote }, (stores.length + (kbCard === null ? 0 : 1)) + " 个域"),
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
                }).concat(kbCard === null ? [] : [kbCard])),
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
            ".om-primary:hover{filter:brightness(.96)}",
            ".om-tool:hover{color:" + M.text + " !important;background:" + M.hover + " !important}",
            ".om-search-clear:hover{color:" + M.text + " !important;background:" + M.hover + " !important}",
            ".om-tabs::-webkit-scrollbar{display:none}",
            ".om-mem-search:focus{border-color:" + M.accent + " !important;box-shadow:0 0 0 3px color-mix(in srgb, " + M.accent + " 14%, transparent) !important}",
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
            // 容量 + 存储域那两栏用 auto-fit 自己就会塌，
            // 不需要额外的规则 —— 一刀切会让 978px 的面板白丢一栏。
            // 这里用的是**容器查询**而不是视口媒体查询：面板的实际宽度跟窗口宽度
            // 不是一回事（同一窗口下侧栏折叠与否能差 300px），按视口判断会出现
            // 「窗口 1280 但面板只有 660px，两栏挤在一起」这种结果。
            // 容器查询需要根元素声明 container-type（见 S.browser），宿主本身也在用
            // 容器查询，所以支持没问题；媒体查询留着当老浏览器的兜底。
            // 知识库资源管理器：面板窄到 620px 以下时左树塌到上方；标准 1258
            // 宽的会话里面板约 728px，应当是「左树 + 右列」而不是上下堆。
            "@media (max-width:620px){.om-kb-explorer{flex-direction:column !important}.om-kb-tree{flex:0 0 auto !important;width:100% !important;max-height:180px !important;overflow:auto !important}.om-gallery-explorer .om-gallery-side{flex:0 0 auto !important;width:100% !important;max-height:160px !important;overflow:auto !important}}",
            "@container (max-width:620px){.om-kb-explorer{flex-direction:column !important}.om-kb-tree{flex:0 0 auto !important;width:100% !important;max-height:180px !important;overflow:auto !important}.om-gallery-explorer .om-gallery-side{flex:0 0 auto !important;width:100% !important;max-height:160px !important;overflow:auto !important}}",
            // 图库的卡片 / 列表行是 <button>：键盘焦点要看得见，选中态已经在行内样式里。
            ".om-gallery-explorer [data-gallery-item]:focus-visible,.om-gallery-explorer [data-gallery-subfolder]:focus-visible{outline:2px solid " + M.accent + ";outline-offset:2px}",
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
        /**
         * 「拿不到会话快照 hook」时的替身（第六十四轮）。
         *
         * 面板要同步读当前会话的工作目录，而那个 hook 只有部分部署会给
         * （`conversation.view` 的 standard props 里有 `useSessions`）。为了不写
         * 「按 props 决定调不调 hook」那种会违反 hook 规则的代码，这里备一个恒返回
         * undefined 的替身：真 hook 与替身**都只调一次**，调用次数稳定。
         */
        const useNoSessions = () => undefined;

        function MemoryBrowserPanel(props) {
            /** 客户端 layout 服务（可选）：待办指令上的「切到会话」用它切主面板。 */
            const layout = props && props.layout ? props.layout : null;
            /**
             * 宿主给的会话页签切换入口（`conversation.view` 的 owner props 之一）。
             *
             * 第六十二轮起本面板就是会话页签行里的一格，所以「切到会话」= 选中
             * `chat` 那一格。宿主实现里还会先 activateView 再写一份 viewRequest，
             * focus 给空串（这一跳没有要定位到的那条消息，各视图对空 focus 都是
             * 不定位）。两个入口都拿不到时按钮不画 —— 画一个点了没反应的按钮更糟。
             */
            const openView = props && typeof props.openView === "function" ? props.openView : null;
            const jumpToChat = openView !== null
                ? () => { try { openView("chat", ""); } catch (error) { /* 切不动就算了 */ } }
                : (layout && typeof layout.selectPanel === "function"
                    ? () => {
                        // 老宿主（本面板还挂在 main 那一格的部署）：没有 openView，
                        // 退回「切主面板到 conversation」这条老路。
                        try { layout.selectPanel("conversation"); } catch (error) { /* 面板切不动就算了 */ }
                    }
                    : null);
            /**
             * 当前会话的工作目录（第六十四轮：先把默认根定成「我现在在哪个工作区」）。
             *
             * 为什么拿得到：这一格是 `conversation.session` 下的 `conversation.view`
             * （**session 作用域**），宿主会把 `sessionId` 与 `useSessions` 一起发给视图组件 ——
             * 口径照抄 dsh 自己的 `ConversationContent`：
             * `useSessions((s) => s.byId[sessionId]?.cwd)`。
             *
             * **同步读**（不是「先取一份数据再换」）：第一次取数就把 cwd 带上，否则用户
             * 打开面板时会先看到宿主挑的那个工作区、再被换成当前工作区，闪一下。
             * 拿不到 hook 或拿不到 cwd（老宿主 / 空会话，例如单测）时留空串，退回老行为。
             */
            const useSessions = typeof props.useSessions === "function" ? props.useSessions : useNoSessions;
            const sessionCwdRaw = useSessions((snapshot) => (
                typeof props.sessionId === "string" && props.sessionId !== ""
                    ? snapshot?.byId?.[props.sessionId]?.cwd
                    : undefined
            ));
            /** `""` = 这个部署/这一刻给不出当前工作区 → 由宿主挑默认根。 */
            const bootCwd = typeof sessionCwdRaw === "string" ? sessionCwdRaw : "";
            const [state, setState] = React.useState({
                loading: true, error: null, data: null,
                tab: "hot", hotScope: "preference", cwd: "", query: "", applied: "",
                // 来源筛选（第八十二轮）："" = 全部，global / workspace 只看一个根。
                // 后端按当前层过滤列表，originCounts 给三档计数，前端只负责切换。
                origin: "",
                // 知识库页签的两种看法：文档清单（docs）与块检索（search）。
                // 检索是后端 kbSearch 的 BM25 结果，块清单走 kbDoc，正文走 kbRead。
                kbMode: "docs", kbQuery: "", appliedKb: "",
                // 文件识别（OCR）：本地目录浏览器选文件 → POST /office-memory/ocr → 结果面板。
                ocrBrowsing: false, ocr: null,
                // 知识库的虚拟文件夹（第八十四轮）：空串 = 根。左树换层只影响列出的文档
                // 与检索范围，不改变来源筛选；来源与文件夹是两个正交的维度。
                kbFolder: "", kbRecursive: false, kbClosed: {},
                // Windows 11 资源管理器式状态：排序、视图、选中行、后退栈。
                // 选中只是界面态，不写回服务端；双击才进入目录 / 打开文档。
                kbSort: "name", kbSortDir: "asc", kbView: "details", kbSelectedKey: "", kbBack: [],
                // 「新建文件夹」的行内草稿：null = 收起，字符串 = 正在输入的名字。
                // 只活在面板里，点「新建」才 POST；成功后进入新目录，失败保留草稿。
                kbNewFolder: null,
                // 新建失败的原因单独放：它必须紧贴新建条显示，不能依赖编辑器是否打开。
                kbFolderError: "",
                // 目录切换失败时的显式恢复：confirmed 记录「POST 已成功、目录确实存在」，
                // 与「内容没读回来」分开；retry 让用户不必先退回再重进。
                kbFolderNotice: "", kbFolderRetry: null,
                // 服务端真分页（第九十五轮）：offset/limit 是这一页的窗口，服务端按
                // 全量清单排序后切片，前端不再自己截断。kbPreviews 是「先预览、
                // 再显式入库」的待确认清单：正文只留在内存 ref 里，不落盘、不调模型。
                kbOffset: 0, kbLimit: LIST_PAGE, kbPreviews: [],
                // 模型增强入库：勾选才把原文发给模型；模型默认取设置里的值，
                // 用户可在真实宿主模型清单里换一个。未就绪时复选框置灰并给原因。
                kbAi: false, kbAiModel: "", kbAiCatalog: null, kbAiLoading: false,
                kbUploadNote: "",
                // 流式进度（第一百一十轮）：只描述当前这一次上传，收尾/失败都清空。
                kbProgress: null,
                // AI 入库先预估、再确认、可暂停续跑：预估结果与暂停回执都留在面板上，
                // 用户看得到「要调几次、超没超预算」，而不是点完才在回执里发现。
                kbEstimate: null, kbPaused: null,
                kbDoc: "", kbChunk: "",
                // 第八十五轮：可编辑。editor 是一份草稿（不落盘），saving/actionError
                // 只服务当前这次写入。知识库的块选择与摘要草稿也放在这一层。
                editor: null, saving: false, actionError: "",
                kbSelected: {}, kbSummary: null, kbRelated: "", kbSummaries: false,
                // 摘要块分页：与文档清单分开，默认每页 10 条。
                kbSummaryOffset: 0, kbSummaryLimit: 10,
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
                // 面板顶部的快速新增：填正文 → 选个人偏好 / 项目记忆 → 直接写入。
                // 与「记一条」草稿分开，用户不必先开编辑器才能写字。
                // 默认收起成一行：常驻的输入卡会把正文挤到首屏之外，用户要先滚过
                // 一张「现在不想用」的表单；点开后才是完整输入。
                quickText: "", quickTarget: "user", quickBusy: false, quickOpen: false,
                // 知识库删除确认（第八十九轮）：先选目标，再点一次确认；不在列表里
                // 直接删，避免误触把一份文档连同块关系一起清掉。
                deleteTarget: null,
                // 块间关系草稿：{ sourceId, targetId, kind, note }。面板不再要求用户
                // 回会话敲 office_memory(kb-link)，块详情里就能建边。
                kbLink: null,
                // 有旧数据时刷新只显示「更新中」，不把整页换成加载态。
                refreshing: false,
                // 图库（独立一级页签）：列表与状态分开取；预览 / 手工标签 / 描述
                // 都围绕当前选中图，不把图片塞进记忆快照。
                galleryLoading: true, galleryError: "", galleryData: null, galleryStatus: null,
                galleryMode: "list", galleryQuery: "", galleryApplied: "", gallerySelected: "",
                // 浏览位置的历史栈与文件夹展开状态：和知识库一样，后退只恢复路径，
                // 不缓存清单（目录内容可能刚被上传 / 移动改过）。
                galleryBack: [], galleryClosed: {},
                galleryDraft: null, galleryBusy: false, galleryUploading: "",
                galleryDescribeConfirm: false, galleryCategory: "", galleryFavorite: false,
                // 文件夹是组织维度：选中后只看该层直属图片，检索与分类仍可在其中叠加。
                galleryFolder: "", galleryFolderName: "", galleryFolderNewOpen: false,
                galleryFolderEdit: "", galleryFolderEditValue: "", galleryFolderRemoveConfirm: false,
                galleryTags: "", gallerySort: "added", gallerySortDir: "desc",
                galleryImageSearch: false, galleryAllowUpload: false,
                // 文字搜索默认只走本地词法 / 文本向量；云端图文对齐（把搜索词发给
                // 多模态端点）要用户在这里显式勾选，请求才带 allowUpload=true。
                galleryCloudSearch: false,
                // 原图查看层与删除确认：都是本层面板状态，不落盘。
                galleryPreview: false, galleryRemoveConfirm: false, galleryNotice: "",
            });
            // 模型清单请求号：RPC 与旧 HTTP 可能同时在途（例如 remote 迟绑定），
            // 只让最新一次响应落状态，避免旧清单盖掉新清单。
            const kbCatalogSeq = React.useRef(0);
            // 记忆快照请求号：输入检索词时每一键都可能发一次请求，只让**最新**
            // 那一次落状态；否则慢请求后到会把新结果盖成旧结果（闪回）。
            const kbViewSeq = React.useRef(0);
            // 预览正文只放 ref：base64 塞进 state 会让整棵面板跟着重渲染。
            const kbPreviewDataRef = React.useRef({});
            // 预览批次的请求号：一次选了多份时，后一批的选择作废前一批未完成的预览。
            const kbPreviewSeq = React.useRef(0);
            // 异步回来时要拿「最新的」state：闭包里那份是发起请求那一刻的旧值。
            const latest = React.useRef(state);
            latest.current = state;

            // patch 立刻并进 latest.current：写入成功后会先清输入框、再等一次 reload，
            // 只写 setState 的话 reload 响应会按旧 state 合并，把刚清掉的字又带回来。
            const patchState = (patch) => {
                const merged = Object.assign({}, latest.current, patch);
                latest.current = merged;
                setState(merged);
            };

            /**
             * 面板根（「再显示 N 条」与「回到顶部」共用）。
             *
             * 面板不是滚动容器（见 S.browser），真正能滚的是宿主会话滚动容器；
             * 切页签 / 搜索 / 换目录后如果不把滚动位置带回去，用户会停在上一层的
             * 底部，看见新页签的页脚或一片空白，以为内容没了。
             */
            const listRef = React.useRef(null);
            const scrollPanelTop = () => {
                const node = listRef.current;
                if (node === null || node === undefined) return;
                const el = scrollParentOf(node);
                if (!el) return;
                try {
                    const panelTop = node.getBoundingClientRect().top - el.getBoundingClientRect().top + el.scrollTop;
                    el.scrollTo({ top: Math.max(0, panelTop), behavior: "auto" });
                } catch (error) {
                    try { el.scrollTop = 0; } catch (inner) { /* 滚不动就算了 */ }
                }
            };

            const load = (patch, options) => {
                const base = Object.assign({}, latest.current, patch || {});
                const opt = options && typeof options === "object" ? options : {};
                // 目录切换先只发请求，**不先改 state.kbFolder**：GET 失败时旧目录的清单
                // 必须继续挂在旧面包屑下，不能变成「新目录 + 旧行」。成功响应才提交路径。
                const requestedFolder = patch && Object.prototype.hasOwnProperty.call(patch, "kbFolder")
                    ? String(patch.kbFolder || "")
                    : String(base.kbFolder || "");
                const previousFolder = String(latest.current.kbFolder || "");
                const folderChanged = patch !== undefined && patch !== null
                    && Object.prototype.hasOwnProperty.call(patch, "kbFolder")
                    && requestedFolder !== previousFolder;
                const confirmedFolder = typeof opt.confirmedFolder === "string" ? opt.confirmedFolder : "";
                const noticeFor = (reason) => {
                    if (!folderChanged) return "";
                    const text = String(reason || "读取失败。");
                    if (confirmedFolder !== "") {
                        return "文件夹「" + confirmedFolder + "」已创建，但读取内容失败：" + text
                            + "。它已经在知识库里，点「重试打开」或稍后刷新后再进去。";
                    }
                    return "没能打开「" + requestedFolder + "」：" + text
                        + "。仍停留在「" + (previousFolder === "" ? "全部文档" : previousFolder) + "」。";
                };
                const retryFor = () => folderChanged
                    ? {
                        folder: requestedFolder,
                        confirmed: confirmedFolder,
                        // 失败时 kbBack 还没提交；重试成功要把它一起带上，
                        // 否则后退键会丢掉「失败前所在目录」这一层历史。
                        back: Array.isArray(patch && patch.kbBack) ? patch.kbBack.slice() : null,
                    }
                    : null;
                // 并发防护：输入时每按一键就可能发一条请求，只认最后一发。
                // 旧响应后到时直接丢弃，不覆盖新结果（否则检索框会闪回旧清单）。
                const seq = ++kbViewSeq.current;
                // patch 必须落进 state，不能只用在这次请求的 URL 上：
                // 不落的话，搜索完再点刷新 / 换工作目录时 applied 又变回空，
                // 用户看到的过滤词被悄悄丢掉（第一次写就踩了这个坑）。
                //
                // 有旧数据时**不**切 loading：切页签 / 换层 / 刷新都走这里，原来每次都
                // 把整页换成「正在读取记忆…」，看起来像闪了一下白屏。保留旧内容，只在
                // 标题旁标「更新中」；第一次取数没有旧内容，才显示加载态。
                const firstLoad = latest.current.data === null;
                const initialPatch = Object.assign({}, patch || {});
                // 换文件夹或进入块检索时摘要块回到第一页：否则新范围下还停在第 3 页，
                // 服务端会把它回拉到末页，看起来像「一搜索就跳到最后一页」。
                if ((folderChanged || base.kbMode === "search")
                    && !(patch && Object.prototype.hasOwnProperty.call(patch, "kbSummaryOffset"))) {
                    initialPatch.kbSummaryOffset = 0;
                }
                if (folderChanged) {
                    // 路径与后退栈都等成功响应再提交；偏移 / 文档选择等可以立刻生效。
                    delete initialPatch.kbFolder;
                    delete initialPatch.kbBack;
                    initialPatch.kbFolderNotice = "";
                    initialPatch.kbFolderRetry = null;
                }
                patchState(Object.assign({}, initialPatch, firstLoad
                    ? { loading: true, refreshing: false, error: null }
                    : { loading: false, refreshing: true, error: null }));
                const params = [];
                if (base.cwd) params.push("cwd=" + encodeURIComponent(base.cwd));
                if (base.applied) params.push("q=" + encodeURIComponent(base.applied));
                // 来源筛选对所有层生效；空串不传，保持普通浏览的 URL 与原行为一致。
                if (base.origin) params.push("origin=" + encodeURIComponent(base.origin));
                // 知识库下钻只在知识库页签下要：别让离开这一格之后还替它算检索。
                if (base.tab === "kb") {
                    if (base.kbFolder) params.push("kbfolder=" + encodeURIComponent(base.kbFolder));
                    // 资源管理器口径：文档浏览默认只列当前层，这里显式上报 1/0。
                    // 块检索不跟着「仅当前层」走：在「全部文档」下检索要能打到子文件夹，
                    // 在某个文件夹里检索则是它的整棵子树 —— 搜索永远递归。
                    const searchRecursive = base.kbMode === "search" && base.appliedKb
                        ? true
                        : base.kbRecursive !== false;
                    params.push("kbrecursive=" + (searchRecursive ? "1" : "0"));
                    if (base.kbMode === "search" && base.appliedKb) params.push("kbq=" + encodeURIComponent(base.appliedKb));
                    if (base.kbDoc) params.push("kbdoc=" + encodeURIComponent(base.kbDoc));
                    if (base.kbChunk) params.push("kbchunk=" + encodeURIComponent(base.kbChunk));
                    if (base.kbRelated) params.push("kbrelated=" + encodeURIComponent(base.kbRelated));
                    // 摘要块是 SQLite 里的派生数据：只有知识库页签要看，别让别的页签打开库。
                    params.push("kbsummaries=1");
                    params.push("kbsummariesoffset=" + encodeURIComponent(String(Math.max(0, Number(base.kbSummaryOffset) || 0))));
                    params.push("kbsummarieslimit=" + encodeURIComponent(String(Math.max(1, Number(base.kbSummaryLimit) || 10))));
                    // 真分页：窗口与排序都由服务端算，前端只报「第几页、怎么排」。
                    params.push("kboffset=" + encodeURIComponent(String(Math.max(0, Number(base.kbOffset) || 0))));
                    params.push("kblimit=" + encodeURIComponent(String(Math.max(1, Number(base.kbLimit) || LIST_PAGE))));
                    params.push("kbsort=" + encodeURIComponent(String(base.kbSort || "name")));
                    params.push("kbsortdir=" + encodeURIComponent(String(base.kbSortDir || "asc")));
                }
                const url = MEMORY_VIEW_PATH + (params.length > 0 ? "?" + params.join("&") : "");

                // fetch 只在浏览器里有。这里必须自己兜住：抛出去会让整格主面板白屏，
                // 而「没有 fetch」在测试与受限环境里是正常状态，不是异常。
                if (typeof fetch !== "function") {
                    patchState({ loading: false, refreshing: false, error: "当前环境没有 fetch，无法读取记忆内容（这一格只在浏览器里工作）。" });
                    return;
                }
                let pending;
                try {
                    pending = fetch(url, { headers: { accept: "application/json" } });
                } catch (error) {
                    const reason = "读取记忆失败：" + (error && error.message ? error.message : String(error));
                    const failure = { loading: false, refreshing: false, error: reason };
                    if (folderChanged) {
                        failure.kbFolderNotice = noticeFor(reason);
                        failure.kbFolderRetry = retryFor();
                    }
                    patchState(failure);
                    return;
                }
                // 端点失败时 HTTP 可能是 200 / 404 / 500，三种都要当数据看：
                // 只要 body 里带 error 就照实显示，不靠状态码判断。
                Promise.resolve(pending).then(
                    (response) => Promise.resolve(response && typeof response.json === "function" ? response.json() : null)
                        .then((body) => ({ response: response, body: body }), () => ({ response: response, body: null })),
                ).then((result) => {
                    // 过期响应：晚到的旧请求不许覆盖新结果。
                    if (kbViewSeq.current !== seq) return;
                    const body = result.body;
                    const status = result.response && result.response.status ? result.response.status : "?";
                    if (body === null || typeof body !== "object") {
                        const reason = "记忆端点没有返回 JSON（HTTP " + status + "）。";
                        const failure = { loading: false, refreshing: false, error: reason };
                        if (folderChanged) {
                            failure.kbFolderNotice = noticeFor(reason);
                            failure.kbFolderRetry = retryFor();
                        }
                        patchState(failure);
                        return;
                    }
                    if (body.ok !== true) {
                        // 失败响应仍会带回 workspaces：留着，用户才好换一个目录再试。
                        const reason = String(body.error || "读取记忆失败。");
                        const failure = {
                            loading: false,
                            refreshing: false,
                            error: reason,
                            data: Array.isArray(body.workspaces) && body.workspaces.length > 0
                                ? { workspaces: body.workspaces }
                                : latest.current.data,
                        };
                        if (folderChanged) {
                            failure.kbFolderNotice = noticeFor(reason);
                            failure.kbFolderRetry = retryFor();
                        }
                        patchState(failure);
                        return;
                    }
                    // 发起这次请求时带的 patch（清编辑器、下钻目标）要在响应落地时
                    // 再生效一次：宿主那一轮渲染可能还没把 latest.current 刷新，
                    // 只按旧 state 合并会把刚清掉的 editor 又带回来。
                    const next = Object.assign({}, patch || {}, { loading: false, refreshing: false, error: null, data: body, kbFolderNotice: "", kbFolderRetry: null });
                    // 第一次拿到数据时把 cwd 补上，后续请求才会带上它。
                    if (!latest.current.cwd && body.cwd) next.cwd = body.cwd;
                    // 服务端会把越界 offset 回拉到合法页（删掉末页最后一条就是这种情况），
                    // 回写的 offset 必须跟响应一致，否则下一次「下一页」会跳过一页。
                    if (body.kbPage && typeof body.kbPage === "object" && Number.isFinite(Number(body.kbPage.offset))) {
                        next.kbOffset = Math.max(0, Number(body.kbPage.offset));
                    }
                    if (body.kbSummariesPage && typeof body.kbSummariesPage === "object" && Number.isFinite(Number(body.kbSummariesPage.offset))) {
                        next.kbSummaryOffset = Math.max(0, Number(body.kbSummariesPage.offset));
                        // 服务端会按硬上限把每页条数收口；回写生效值，下一页才不会换窗口。
                        if (Number.isFinite(Number(body.kbSummariesPage.limit))) {
                            next.kbSummaryLimit = Math.max(1, Number(body.kbSummariesPage.limit));
                        }
                    }
                    patchState(next);
                }, (error) => {
                    if (kbViewSeq.current !== seq) return;
                    const reason = "读取记忆失败：" + (error && error.message ? error.message : String(error));
                    const failure = { loading: false, refreshing: false, error: reason };
                    if (folderChanged) {
                        failure.kbFolderNotice = noticeFor(reason);
                        failure.kbFolderRetry = retryFor();
                    }
                    patchState(failure);
                });
            };

            /**
             * 图库取数：列表与检索都用同一条 GET（view=list / view=search）。
             *
             * 图库是独立一级视图，不走记忆快照：图片条目多、缩略图按需取，
             * 混进 snapshot 会让「打开记忆」也付一次图库的读盘成本。
             */
            const loadGallery = (patch) => {
                const base = Object.assign({}, latest.current, patch || {});
                const params = ["cwd=" + encodeURIComponent(String(base.cwd || bootCwd || ""))];
                // 文件夹是浏览位置，不是检索条件：list / search 都带上，切文件夹不必清空检索。
                const folderNow = String(base.galleryFolder || "");
                if (folderNow !== "") params.push("folder=" + encodeURIComponent(folderNow));
                // 以图搜图必须有「选中一张图」才有查询图；只勾了开关还没选图时
                // 退回列表，不要发一条没有 image 的 search 请求。
                const imageQuery = base.galleryImageSearch === true && String(base.gallerySelected || "") !== "";
                const searching = String(base.galleryApplied || "") !== ""
                    || String(base.galleryTags || "") !== ""
                    || String(base.galleryCategory || "") !== ""
                    || base.galleryFavorite === true
                    || imageQuery;
                if (searching) {
                    params.push("view=search", "limit=60");
                    if (String(base.galleryApplied || "") !== "") params.push("q=" + encodeURIComponent(String(base.galleryApplied)));
                    if (String(base.galleryTags || "") !== "") params.push("tags=" + encodeURIComponent(String(base.galleryTags)));
                    if (String(base.galleryCategory || "") !== "") params.push("category=" + encodeURIComponent(String(base.galleryCategory)));
                    if (base.galleryFavorite === true) params.push("favorite=1");
                    if (imageQuery) {
                        params.push("image=" + encodeURIComponent(String(base.gallerySelected)));
                        // 云端费用提示由界面上的复选框负责：没勾「允许内容发给模型」
                        // 就绝不替用户补 allowUpload=true（那等于绕过确认）。
                        // 同时带文字查询时，文字也走云端要另一个勾选（见下），两个都
                        // 允许才发这一个开关，免得顺带把搜索词发出去。
                        const textToo = String(base.galleryApplied || "") !== "";
                        if (base.galleryAllowUpload === true && (!textToo || base.galleryCloudSearch === true)) params.push("allowUpload=true");
                    }
                    // 文字检索默认只走本地通道（词法 / 标签 / 本地文本向量）。只有用户
                    // 显式勾「文字搜索走云端（含向量）」才把 allowUpload 发出去；否则后端就算
                    // 配了多模态端点也不该为一次普通搜索产生云端调用。
                    if (!imageQuery && String(base.galleryApplied || "") !== "" && base.galleryCloudSearch === true) {
                        params.push("allowUpload=true");
                    }
                } else {
                    params.push("view=list", "limit=200");
                }
                patchState(Object.assign({}, patch || {}, { galleryLoading: true, galleryError: "" }));
                if (typeof fetch !== "function") {
                    patchState({ galleryLoading: false, galleryError: "当前环境没有 fetch，无法读取图库（这一格只在浏览器里工作）。" });
                    return;
                }
                return Promise.resolve(fetch(GALLERY_PATH + "?" + params.join("&"), { headers: { accept: "application/json" } }))
                    .then((response) => Promise.resolve(response.json()).then((body) => ({ response: response, body: body }), () => ({ response: response, body: null })))
                    .then((result) => {
                        const body = result.body;
                        if (body === null || typeof body !== "object" || body.ok !== true) {
                            patchState({ galleryLoading: false, galleryError: String((body && body.error) || ("图库端点没有返回 JSON（HTTP " + (result.response && result.response.status ? result.response.status : "?") + "）。")) });
                            return;
                        }
                        patchState(Object.assign({}, patch || {}, { galleryLoading: false, galleryError: "", galleryData: body }));
                    }, (error) => {
                        patchState({ galleryLoading: false, galleryError: "读取图库失败：" + String(error && error.message ? error.message : error) });
                    });
            };

            /** 图库能力状态：未配置图文对齐 / 视觉模型时，界面必须照实说，不画假能力。 */
            const loadGalleryStatus = () => {
                if (typeof fetch !== "function") return Promise.resolve(null);
                const url = GALLERY_PATH + "?view=status&cwd=" + encodeURIComponent(String(latest.current.cwd || bootCwd || ""));
                return Promise.resolve(fetch(url, { headers: { accept: "application/json" } }))
                    .then((response) => Promise.resolve(response.json()).then((body) => body, () => null))
                    .then((body) => {
                        if (body && body.ok === true) patchState({ galleryStatus: body });
                        return body && body.ok === true ? body : null;
                    }, () => null); /* 状态读不到不挡浏览 */
            };

            /** 图库写入：add / manual / describe / vectors / remove 都走这一条 POST。 */
            const galleryAction = (action, payload, reloadPatch) => {
                if (typeof fetch !== "function") {
                    patchState({ galleryError: "当前环境没有 fetch，无法写入图库（这一格只在浏览器里工作）。" });
                    return Promise.resolve(null);
                }
                patchState({ galleryBusy: true, galleryError: "" });
                let request;
                try {
                    request = fetch(GALLERY_PATH + "?cwd=" + encodeURIComponent(String(latest.current.cwd || bootCwd || "")), {
                        method: "POST",
                        headers: { "content-type": "application/json", accept: "application/json" },
                        body: JSON.stringify(Object.assign({ action: action }, payload || {})),
                    });
                } catch (error) {
                    patchState({ galleryBusy: false, galleryError: "图库写入失败：" + String(error && error.message ? error.message : error) });
                    return Promise.resolve(null);
                }
                return Promise.resolve(request)
                    .then((response) => Promise.resolve(response.json()).then((body) => ({ response: response, body: body }), () => ({ response: response, body: null })))
                    .then((result) => {
                        const body = result.body;
                        if (body === null || typeof body !== "object" || body.ok !== true) {
                            const reason = body && body.error ? String(body.error) : ("HTTP " + (result.response && result.response.status ? result.response.status : "?"));
                            // 未配置是用户要处理的事，不是网络故障：把设置路径一起给出来。
                            throw new Error(reason + (body && body.code === "unconfigured"
                                ? " 去「设置 → 办公与记忆 → 图库 → 视觉描述」选一个通道。"
                                : ""));
                        }
                        patchState({ galleryBusy: false });
                        loadGallery(reloadPatch || {});
                        return body;
                    })
                    .catch((error) => {
                        patchState({ galleryBusy: false, galleryError: "图库写入失败：" + String(error && error.message ? error.message : error) });
                        return null;
                    });
            };

            /**
             * 宿主模型清单：真实 provider/model，用来填充上传工具栏的选择器。
             *
             * 与设置页的 KbLlmModelPicker 用同一条来源：登录后优先
             * ctx.remote.session.modelCatalog()（宿主 Models 页的同一个 RPC），
             * 不需要 cwd，也不发 HTTP。只有旧宿主没有这个方法时才退回
             * /office-memory/llm-models。RPC 拒绝 / 抛错原样显示并停在重试，
             * **不能**因为失败就偷偷改走 HTTP —— 权限拒绝尤其不能绕过。
             */
            const loadModelCatalog = (force) => {
                if (latest.current.kbAiCatalog !== null && force !== true) return Promise.resolve(latest.current.kbAiCatalog);
                kbCatalogSeq.current += 1;
                const seq = kbCatalogSeq.current;
                const stale = () => kbCatalogSeq.current !== seq;
                // 新请求未决期间旧清单不能再用来判「缺失」或画选项；同时把旧错误收掉。
                patchState({ kbAiLoading: true, kbAiCatalog: null });
                const fail = (error) => {
                    if (stale()) return null;
                    patchState({
                        kbAiLoading: false,
                        kbAiCatalog: {
                            ok: false, providers: [], failures: [],
                            error: String(error && error.message ? error.message : error),
                        },
                    });
                    return null;
                };
                const succeed = (catalog) => {
                    if (stale()) return null;
                    patchState({ kbAiLoading: false, kbAiCatalog: catalog });
                    return catalog;
                };
                let viaRpc = null;
                try {
                    viaRpc = sessionModelCatalog();
                } catch (error) {
                    return Promise.resolve(fail(error));
                }
                // 旧宿主 / 旧兼容端点的只读 HTTP 清单；两条路径共用同一份投影。
                const loadViaHttp = () => {
                    if (typeof fetch !== "function") {
                        return Promise.reject(new Error("当前环境没有 fetch，读不到宿主模型清单。"));
                    }
                    const url = LLM_MODELS_PATH + "?cwd=" + encodeURIComponent(String(latest.current.cwd || bootCwd || ""));
                    return Promise.resolve()
                        .then(() => fetch(url, { headers: { accept: "application/json" } }))
                        .then((response) => Promise.resolve(response.json()).then(
                            (body) => ({ response: response, body: body }),
                            () => ({ response: response, body: null }),
                        ))
                        .then((result) => {
                            const status = result.response && result.response.status ? result.response.status : "?";
                            const body = result.body;
                            if (body === null || typeof body !== "object") {
                                throw new Error("宿主模型清单响应格式不对（HTTP " + status + "）。");
                            }
                            if (body.ok !== true) {
                                throw new Error(String(body.error || ("宿主返回 HTTP " + status)));
                            }
                            return {
                                ok: true,
                                providers: Array.isArray(body.providers) ? body.providers : [],
                                failures: [],
                            };
                        });
                };
                if (viaRpc !== null) {
                    return Promise.resolve()
                        .then(() => viaRpc())
                        .then((result) => {
                            if (result === null || typeof result !== "object") {
                                throw new Error("宿主模型清单响应格式不对（modelCatalog 没有返回结果）。");
                            }
                            if (result.ok !== true) {
                                // 唯一例外：旧宿主明确没有导出这个方法。
                                if (modelCatalogMethodUnavailable(result.error)) return loadViaHttp();
                                throw new Error(modelCatalogErrorText(result.error));
                            }
                            const catalog = catalogFromModelCatalog(result.value);
                            return { ok: true, providers: catalog.providers, failures: catalog.failures };
                        })
                        .then(succeed, fail);
                }
                return loadViaHttp().then(succeed, fail);
            };

            const galleryImageHref = (id, kind) => {
                const owner = latest.current || {};
                const holder = owner.data && typeof owner.data === "object" ? owner.data : {};
                return GALLERY_PATH + "?view=" + (kind || "thumb")
                    + "&cwd=" + encodeURIComponent(String(owner.cwd || holder.cwd || bootCwd || ""))
                    + "&id=" + encodeURIComponent(String(id || ""));
            };

            /**
             * 选中图的编辑草稿。
             *
             * 名称是可选项（上传时自动带文件名），描述优先用用户改写，空串时回落到
             * 模型总结；标签在界面上用逗号分隔的一行编辑，写回时再切成数组
             * （manual 的 patch.tags 是整组替换）。
             */
            const galleryDraftOf = (item) => ({
                id: String(item && item.id || ""),
                name: String(item && item.name || ""),
                description: String(item && (item.description || (item.aesthetic && item.aesthetic.summary)) || ""),
                tags: (Array.isArray(item && item.tags) ? item.tags : []).join(", "),
                category: String(item && item.category || ""),
                folder: String(item && item.folder || ""),
                note: String(item && item.note || ""),
                favorite: Boolean(item && item.favorite === true),
            });
            const galleryTagList = (text) => String(text || "")
                .split(/[,，\s]+/)
                .map((part) => part.trim())
                .filter((part) => part !== "");
            const galleryShortId = (id) => String(id || "").slice(0, 8);
            /** 加入时间：只画到分钟，卡片里放不下完整 ISO。 */
            const galleryWhen = (value) => {
                const time = Date.parse(String(value || ""));
                if (!Number.isFinite(time)) return "—";
                const date = new Date(time);
                const pad = (n) => String(n).padStart(2, "0");
                return date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate())
                    + " " + pad(date.getHours()) + ":" + pad(date.getMinutes());
            };
            /**
             * 图库列表排序。
             *
             * 名称排序优先用户起的名字，没名字再按 id 前 8 位；相似度只在检索结果里
             * 有意义，没分数时排到最后而不是当成 0 分混进前面。
             */
            const galleryCompare = (left, right, sort, dir) => {
                const sign = dir === "desc" ? -1 : 1;
                if (sort === "score") {
                    const a = Number(left && left.score);
                    const b = Number(right && right.score);
                    const av = Number.isFinite(a) ? a : -1;
                    const bv = Number.isFinite(b) ? b : -1;
                    return av === bv ? 0 : (av < bv ? -sign : sign);
                }
                if (sort === "size") {
                    const a = Number(left && left.bytes) || 0;
                    const b = Number(right && right.bytes) || 0;
                    return a === b ? 0 : (a < b ? -sign : sign);
                }
                if (sort === "name") {
                    const a = String(left && left.name || left && left.id || "");
                    const b = String(right && right.name || right && right.id || "");
                    return a === b ? 0 : (a < b ? -sign : sign);
                }
                const a = Date.parse(String(left && left.addedAt || "")) || 0;
                const b = Date.parse(String(right && right.addedAt || "")) || 0;
                return a === b ? 0 : (a < b ? -sign : sign);
            };

            /**
             * 面板写入（第八十五轮）：所有编辑都走同一条 POST，成功后 reload 当前视图。
             *
             * 不在这里做「乐观更新」：写入要经过宿主的 trustedRootFor 与 memory 的
             * 容量/质量门，前端猜它会成功只会让失败时显示一份假数据。
             */
            const mutatePanel = (action, payload, reloadPatch, options) => {
                const opt = options && typeof options === "object" ? options : {};
                // 新建文件夹的错误要显示在新建条旁：即使编辑器开着，也不能被藏进编辑器
                // 的错误位。默认仍是 actionError，其它写入不受影响。
                const failureKey = typeof opt.failureKey === "string" && opt.failureKey !== "" ? opt.failureKey : "actionError";
                const fail = (message) => {
                    const next = { saving: false };
                    next[failureKey] = message;
                    if (failureKey !== "actionError") next.actionError = "";
                    patchState(next);
                };
                if (typeof fetch !== "function") {
                    fail("当前环境没有 fetch，无法写入（这一格只在浏览器里工作）。");
                    return;
                }
                const start = { saving: true, actionError: "" };
                start[failureKey] = "";
                patchState(start);
                let request;
                try {
                    request = fetch(MEMORY_MUTATE_PATH, {
                        method: "POST",
                        headers: { "content-type": "application/json", accept: "application/json" },
                        body: JSON.stringify(Object.assign({ action: action }, payload || {})),
                    });
                } catch (error) {
                    fail("写入失败：" + (error && error.message ? error.message : String(error)));
                    return;
                }
                return Promise.resolve(request).then(
                    (response) => Promise.resolve(response && typeof response.json === "function" ? response.json() : null)
                        .then((body) => ({ response: response, body: body }), () => ({ response: response, body: null })),
                ).then((result) => {
                    const body = result.body;
                    if (body === null || typeof body !== "object" || body.ok !== true) {
                        const reason = body && body.error ? String(body.error) : ("HTTP " + (result.response && result.response.status ? result.response.status : "?"));
                        throw new Error(reason);
                    }
                    // 清态要跟 reload 合并成一次 patch：load 内部按 latest.current 合并，
                    // 先 patch 再 load 会被旧 state 把 editor 又带回来（测试踩过）。
                    patchState({ saving: false });
                    load(Object.assign({ editor: null, kbSummary: null, kbSelected: {}, actionError: "" }, reloadPatch || {}), {
                        confirmedFolder: typeof opt.confirmedFolder === "string" ? opt.confirmedFolder : "",
                    });
                    return body;
                }).catch((error) => {
                    fail("写入失败：" + (error && error.message ? error.message : String(error)));
                    return null;
                });
            };

            const openHotEditor = (item) => patchState({
                actionError: "",
                editor: {
                    kind: "hot", id: item.id, content: String(item.content || ""),
                    target: item.target === "user" ? "user" : "project",
                    importance: item.importance === "critical" || item.importance === "low" ? item.importance : "normal",
                },
            });

            const openKbEditor = (item) => patchState({
                actionError: "",
                editor: {
                    kind: "kb-doc", id: String(item.id || ""), path: String(item.path || ""),
                    title: String(item.title || ""), folder: String(item.folder || ""), tier: String(item.tier || "user"),
                },
            });

            // 知识库正文（第八十五轮）：元数据之外还要能改内容。整篇回填自 kbDoc.text，
            // 保存走 kb-edit —— 它写回源文件再重切，不是直接改块。
            const openKbTextEditor = (doc) => {
                if (doc === null || typeof doc !== "object") return;
                patchState({
                    actionError: "",
                    editor: {
                        kind: "kb-text", id: String(doc.docId || ""), path: String(doc.path || ""),
                        text: String(doc.text || ""),
                    },
                });
            };

            const saveEditor = () => {
                const editor = state.editor;
                if (editor === null || state.saving) return;
                if (editor.kind === "new") {
                    const content = String(editor.content || "").trim();
                    if (content === "") { patchState({ actionError: "正文不能为空。" }); return; }
                    mutatePanel("memory-add", {
                        content: content, target: editor.target, importance: editor.importance,
                    });
                    return;
                }
                if (editor.kind === "hot") {
                    const content = String(editor.content || "").trim();
                    if (content === "") { patchState({ actionError: "正文不能为空。" }); return; }
                    mutatePanel("memory-replace", {
                        id: editor.id, content: content, target: editor.target, importance: editor.importance,
                    });
                    return;
                }
                if (editor.kind === "kb-doc") {
                    mutatePanel("kb-update", {
                        id: editor.id, title: String(editor.title || ""),
                        folder: String(editor.folder || ""), tier: String(editor.tier || ""),
                    });
                    return;
                }
                if (editor.kind === "kb-text") {
                    const text = String(editor.text || "").trim();
                    if (text === "") { patchState({ actionError: "正文不能为空。" }); return; }
                    mutatePanel("kb-edit", { id: editor.id, text: text });
                }
            };

            // 面板里直接写一条：不走「待发送指令」，POST 完 reload 就能看见。
            // 个人偏好进全局根，项目记忆只进当前工作区 —— 由 target 决定根，
            // 这里不额外传 scope（否则会把项目的写进全局）。
            const quickAdd = () => {
                if (state.quickBusy === true) return;
                const content = String(state.quickText || "").trim();
                if (content === "") { patchState({ actionError: "先写点内容，再点添加。" }); return; }
                patchState({ quickBusy: true, actionError: "" });
                Promise.resolve(mutatePanel("memory-add", {
                    content: content,
                    target: state.quickTarget === "project" ? "project" : "user",
                    importance: "normal",
                })).then((body) => {
                    if (body && body.ok === true) {
                        patchState({ quickBusy: false, quickText: "" });
                    } else {
                        patchState({ quickBusy: false });
                    }
                });
            };

            // 删除一份知识库文档（文档列表与块详情共用）：先摆出目标，再让用户确认。
            // kb-drop 的删除粒度就是**整份文档**，块 id 也指到所属文档；确认文案必须
            // 说清这一点，不能让「删这一块」看起来只删一块。
            const askDeleteDoc = (doc) => {
                if (doc === null || typeof doc !== "object") return;
                const id = String(doc.docId || doc.id || "");
                if (id === "") return;
                patchState({
                    actionError: "",
                    deleteTarget: {
                        id: id,
                        title: String(doc.title || doc.path || id),
                        path: String(doc.path || ""),
                        chunks: Number(doc.chunks_total || doc.chunks || 0) || 0,
                    },
                });
            };
            const confirmDeleteDoc = () => {
                const target = state.deleteTarget;
                if (target === null || target === undefined || state.saving === true) return;
                mutatePanel("kb-drop", { id: target.id }, {
                    deleteTarget: null, kbDoc: "", kbChunk: "", kbRelated: "", kbLink: null,
                });
            };

            // 块间关系：两端都必须是索引里的真实块 id。面板只负责把 id/类型/备注
            // 送给 kb-link，存在性与自连校验仍在服务端（同一套写入边界）。
            const saveChunkLink = () => {
                const draft = state.kbLink;
                if (draft === null || draft === undefined || state.saving === true) return;
                const targetId = String(draft.targetId || "").trim();
                if (targetId === "") { patchState({ actionError: "关系需要目标块 id。" }); return; }
                if (targetId === String(draft.sourceId || "")) { patchState({ actionError: "不能把一块连到它自己。" }); return; }
                mutatePanel("kb-link", {
                    sourceId: String(draft.sourceId || ""),
                    targetId: targetId,
                    kind: KB_LINK_KINDS.indexOf(String(draft.kind || "")) >= 0 ? String(draft.kind) : "related",
                    note: String(draft.note || "").trim(),
                }, {
                    kbLink: null,
                    kbDoc: String(state.kbDoc || ""),
                    kbChunk: String(draft.sourceId || ""),
                    kbRelated: String(draft.sourceId || ""),
                });
            };
            const removeChunkLink = (sourceId, item) => {
                const targetId = String(item && item.chunkId || "");
                if (targetId === "") return;
                mutatePanel("kb-unlink", {
                    sourceId: String(sourceId || ""),
                    targetId: targetId,
                    kind: String(item && item.kind || ""),
                }, {
                    kbDoc: String(state.kbDoc || ""),
                    kbChunk: String(sourceId || ""),
                    kbRelated: String(sourceId || ""),
                });
            };

            // 第一次取数：默认根就是**当前会话的工作目录**（拿不到就不带 cwd，由宿主挑）。
            // 只在挂载时跑一次：之后换目录是用户自己的动作（面板里的工作区选择器 / 搜索），
            // 不该被这里覆盖 —— 所以这一条的依赖数组是空的，bootCwd 只当初始值用。
            React.useEffect(() => { load(bootCwd === "" ? undefined : { cwd: bootCwd }); }, []);

            // 图库取数由一级页签的点击触发（见 tabs 的 onClick）：不用 useEffect 盯 tab，
            // 是因为「切页签」是用户动作，取数应当和这次点击同拍；依赖 effect 还会在
            // 每次重渲染时多一次判断，测试替身也只跑挂载那一次。

            // host 通道的模型选择器需要真实清单；只在知识库页且确实用 host 时取。
            // 读 latest.current 而不是闭包里的 data / state：effect 在测试替身里同步执行，
            // 闭包里此时 data 还没初始化（TDZ），真实 React 里也不该抓住旧快照。
            React.useEffect(() => {
                const now = latest.current || {};
                const cfg = now.data && now.data.config && now.data.config.kbLlm ? now.data.config.kbLlm : null;
                if (now.tab !== "kb" || cfg === null || cfg.mode !== "host") return;
                if (now.kbAiCatalog === null) loadModelCatalog();
            }, [state.tab, state.data]);

            // remote 迟绑定 / 重新注入：上传工具栏也要自动换成 RPC 清单。
            // 只刷新当前确实停在 host + 知识库页的那一份，别为一次服务就绪
            // 在别的页签上多打请求。
            React.useEffect(() => subscribeHostRemote(() => {
                const now = latest.current || {};
                const cfg = now.data && now.data.config && now.data.config.kbLlm ? now.data.config.kbLlm : null;
                if (now.tab !== "kb" || cfg === null || cfg.mode !== "host") return;
                loadModelCatalog(true);
            }), [state.tab, state.data]);

            // 点开某一块后，正文就插在那张卡片下面；把它滚进视野，别让用户自己找。
            // scrollIntoView 只在这里有 DOM 时调用，测试环境没有也不影响渲染。
            React.useEffect(() => {
                if (!state.kbChunk || state.loading) return;
                if (typeof document === "undefined") return;
                const node = document.querySelector('[data-kb-chunk-text]');
                if (node && typeof node.scrollIntoView === "function") node.scrollIntoView({ block: "nearest" });
            }, [state.kbChunk, state.loading]);

            const data = state.data || {};
            /**
             * 取数用的工作目录：当前选的（state.cwd）→ 快照回报的（data.cwd）→ 会话 cwd。
             *
             * 第九十一轮修：上传、OCR、图片引用这几处从前直接写裸 `cwd`，而组件作用域里
             * 从来没有这个名字 —— 点「上传文件」会抛 ReferenceError: cwd is not defined。
             * 单测只断言入口存在、没点过按钮，所以一直没被发现。这里收成一个变量。
             */
            const cwdFor = String(state.cwd || data.cwd || bootCwd || "");
            const hasData = state.data !== null && typeof state.data === "object";
            const counts = data.counts || {};
            const config = data.config;
            const limits = config && config.limits ? config.limits : {};
            const workspaces = Array.isArray(data.workspaces) ? data.workspaces : [];
            const stores = Array.isArray(data.stores) ? data.stores : [];
            // 归档隔离提示只在归档视图渲染：后端快照里每个 store 只带字符串数组，
            // 这里再做一次类型过滤与去重，避免同一句提示在多个根上重复出现。
            const archiveWarnings = [];
            const archiveWarningSeen = new Set();
            for (const store of stores) {
                const rows = store && Array.isArray(store.warnings) ? store.warnings : [];
                for (const row of rows) {
                    if (typeof row !== "string") continue;
                    const text = row.trim();
                    if (text === "" || archiveWarningSeen.has(text)) continue;
                    archiveWarningSeen.add(text);
                    archiveWarnings.push(text);
                }
            }
            const archiveWarningBox = archiveWarnings.length === 0
                ? null
                : h("div", {
                    style: S.archiveScopeWarning,
                    "data-archive-scope-warning": "true",
                    role: "status",
                },
                    h("div", { style: S.archiveScopeWarningTitle }, "归档隔离提示"),
                    h("ul", { style: S.archiveScopeWarningList },
                        archiveWarnings.map((text, index) => h("li", { key: "archive-warning-" + index }, text))));
            const withArchiveWarnings = (node) => state.tab === "archive" && archiveWarningBox !== null
                ? h("div", { "data-archive-warning-wrap": "true" }, archiveWarningBox, node)
                : node;
            const hot = Array.isArray(data.hot) ? data.hot : [];
            const ledger = Array.isArray(data.ledger) ? data.ledger : [];
            const archive = Array.isArray(data.archive) ? data.archive : [];
            const baseItems = Array.isArray(data[state.tab]) ? data[state.tab] : [];
            // 热记忆不再按「来源根」分，而按 target 分：user = 个人偏好（写全局），
            // 其它 = 项目记忆（写当前工作区）。台账 / 归档仍按原样。
            const items = state.tab === "hot"
                ? baseItems.filter((item) => (item.target === "user" ? "preference" : "project") === state.hotScope)
                : baseItems;

            // 每个页签先给一句「这一层是什么」，再进列表。比只靠页签文字更容易扫读，
            // 也让切换页签时看起来像换了一页，而不是同一列卡片换了数据。
            const layerKey = state.tab === "hot" ? state.hotScope : state.tab;
            const sectionMeta = {
                preference: ["个人偏好", "可跨工作区使用 · 随办公工具按需提供"],
                project: ["项目记忆", "只在这个工作区生效 · 不跨项目"],
                ledger: ["台账", "office_run 产出的交付物"],
                archive: ["归档", "只读 · 热记忆下沉与台账滚动"],
            }[layerKey] || null;
            const contentHead = hasData && state.tab !== "kb" && items.length > 0
                ? h("div", { style: S.memSectionHead, "data-memory-section": layerKey },
                    h("span", { style: Object.assign({}, S.memSectionDot, { background: metricColor(layerKey === "preference" || layerKey === "project" ? "hot" : layerKey) }) }),
                    h("span", { style: S.memSectionTitle }, sectionMeta[0]),
                    h("span", { style: S.memSectionNote }, sectionMeta[1] + " · " + items.length + " 条"))
                : null;

            const head = h("div", { style: S.memHead },
                h("div", { style: S.memBrand }, h(MemoryPanelIcon, { size: 16 })),
                h("div", { style: S.memTitleBlock },
                    h("div", { style: S.memKicker }, "OFFICE MEMORY"),
                    h("div", { style: S.memTitle }, "记忆" + (data.label ? " · " + data.label : "")),
                    h("div", { style: S.memSubtitle },
                        data.cwd || "工作区",
                        state.refreshing === true
                            ? h("span", { "data-memory-refreshing": "true", style: { marginLeft: "8px", color: M.accent } }, "正在更新…")
                            : null),
                ),
                h("div", { style: S.memHeadActions },
                    // 只有一个工作目录时不必给下拉：多一个只有一项的控件只是噪声。
                    workspaces.length > 1
                        ? h(Select, {
                            value: state.cwd || data.cwd || "",
                            disabled: false,
                            options: workspaces.map((item) => ({ value: item.cwd, label: item.label || item.cwd })),
                            onChange: (next) => { load({ cwd: next }); scrollPanelTop(); },
                        })
                        : null,
                    h("button", {
                        type: "button", className: "om-btn om-primary", style: S.memPrimary,
                        "data-memory-compose": "true",
                        onClick: () => {
                            // 新增只有一个入口：跳回记忆页并聚焦页内的「直接记一条」。
                            // 以前这里还会另开一张草稿卡，导致同一屏出现两个新增框。
                            patchState({
                                tab: "hot", visible: LIST_PAGE, expandedItem: null, pending: null, copied: null, actionError: "",
                                quickTarget: state.hotScope === "project" ? "project" : "user",
                                quickOpen: true,
                            });
                            scrollPanelTop();
                            if (typeof document !== "undefined") {
                                setTimeout(() => {
                                    const input = document.querySelector("[data-memory-quick-text]");
                                    if (input && typeof input.focus === "function") input.focus();
                                }, 0);
                            }
                        },
                    }, h("span", { style: S.memPlus }, "+"), "记一条"),
                ),
            );

            // 搜索与导出单独一行：标题行只留「在哪里」和一个主动作，控件不再挤在标题右边。
            const controls = h("div", { style: S.memControls },
                h("div", { style: S.memSearchWrap },
                    h("span", { style: S.searchIcon }, h(SearchGlyph, null)),
                    h("input", {
                        type: "text",
                        className: "om-input om-mem-search",
                        value: state.query,
                        // 顶部搜索跟着当前看法走：知识库的文档 / 块检索是两套参数，
                        // 占位文案不能只写「块」却只过滤文档标题。
                        placeholder: sectionOf(state.tab) === "kb"
                            ? (state.kbMode === "search" ? "找块：正文、标题" : "找文档：标题、路径")
                            : sectionOf(state.tab) === "gallery" ? "找图库：描述、标签、分类" : "找记忆：内容、来源、标签",
                        onChange: (event) => patchState({ query: String(event.target.value || "") }),
                        onKeyDown: (event) => {
                            if (event.key !== "Enter") return;
                            const next = String(event.target.value || "").trim();
                            const section = sectionOf(state.tab);
                            if (section === "gallery") {
                                loadGallery({
                                    galleryApplied: next, galleryQuery: next, galleryFolder: "", galleryBack: [], query: next,
                                    // 检索结果的默认顺序是相似度：按加入时间排会把最相关的一张埋到后面。
                                    gallerySort: "score", gallerySortDir: "desc",
                                });
                                scrollPanelTop();
                                return;
                            }
                            if (section === "kb" && state.kbMode === "search") {
                                // 块检索：只发 kbq，不把上一轮文档过滤的 applied 带进来。
                                load({ kbMode: "search", kbQuery: next, appliedKb: next, kbDoc: "", kbChunk: "", applied: "", query: next });
                                scrollPanelTop();
                                return;
                            }
                            // 文档 / 记忆搜索：只发 q，清掉另一套检索词，避免两套状态串台。
                            load({ applied: next, query: next, appliedKb: "", kbQuery: "" });
                            scrollPanelTop();
                        },
                        style: S.memSearchInput,
                    }),
                    (sectionOf(state.tab) === "gallery"
                        ? Boolean(state.galleryApplied)
                        : sectionOf(state.tab) === "kb" && state.kbMode === "search"
                            ? Boolean(state.appliedKb)
                            : Boolean(state.applied))
                        ? h("button", {
                            type: "button", className: "om-search-clear", style: S.memSearchClear,
                            "data-memory-search-clear": "true",
                            title: "清空检索",
                            onClick: () => {
                                const section = sectionOf(state.tab);
                                if (section === "gallery") {
                                    galleryLoadFilter({ galleryApplied: "", galleryQuery: "", galleryFolder: "", galleryBack: [], query: "" });
                                    return;
                                }
                                if (section === "kb" && state.kbMode === "search") {
                                    load({ appliedKb: "", kbQuery: "", kbDoc: "", kbChunk: "", query: "" });
                                    scrollPanelTop();
                                    return;
                                }
                                load({ applied: "", query: "" });
                                scrollPanelTop();
                            },
                        }, "×")
                        : null,
                ),
                // 导出（第七十一轮）：agent 记忆没有公认的公有协议，所以采用**单文件**。
                // JSON 是插件自己的版本化 Pack（导入侧已存在，保证无损）；Markdown 给
                // 人读、也给别的笔记工具收。两个都是原生下载，不经过面板内存。
                h("div", { style: S.memToolGroup },
                    h("a", {
                        className: "om-btn om-tool", style: S.memTool,
                        href: memoryExportHref(state.cwd || data.cwd || "", "json"),
                        download: "",
                        title: "导出为单文件 JSON（三层记忆 + 关系 + 知识库清单）：无损，可用 office_memory 的 import 在别的目录导回",
                        "data-memory-export": "json",
                    }, "导出"),
                    h("a", {
                        className: "om-btn om-tool", style: S.memTool,
                        href: memoryExportHref(state.cwd || data.cwd || "", "md"),
                        download: "",
                        title: "导出为 Markdown（CommonMark）：给人读、给笔记工具收；完整迁移请用 JSON",
                        "data-memory-export": "md",
                    }, "MD"),
                    h("span", { style: { width: "1px", height: "16px", background: M.stroke, margin: "0 4px" } }),
                    h("button", { type: "button", className: "om-btn om-tool", style: S.memTool, onClick: () => load() }, "刷新"),
                ),
            );

            // 一级视图只有两个；当前 tab 仍是具体数据层（hot / ledger / kb），
            // section 决定高亮哪一个一级按钮。
            const section = sectionOf(state.tab);
            // 一级页签只说「去哪」，不说「有几条」：记忆是条目、知识库是篇，
            // 两种单位放在同一条轴上会互相打架。数字留给二级与区块头各自回答。
            const tabs = h("div", { style: S.memTabs, className: "om-tabs", role: "tablist" }, BROWSER_TABS.map((tab) => {
                const on = section === tab.key;
                const hovered = state.hoverTab === tab.key && !on;
                const style = on
                    ? Object.assign({}, S.memTab, S.memTabOn)
                    : hovered ? Object.assign({}, S.memTab, S.memTabHover) : S.memTab;
                return h("div", {
                    key: tab.key,
                    className: "om-tab",
                    style: style,
                    role: "tab",
                    "aria-selected": on,
                    "data-tab": tab.key,
                    onClick: () => {
                        const next = tab.key === "memory" ? "hot" : tab.key;
                        patchState({
                            tab: next, visible: LIST_PAGE, expandedItem: null, hoverItem: null,
                            // 搜索框是当前页签的草稿：切页签时清掉上一页的词，避免把
                            // 「图库搜深色」带到记忆页，再回车又查一遍旧词。
                            query: "",
                            // 回到记忆页时，写入目标跟着当前档走：在「项目记忆」里不该默认写全局。
                            quickTarget: tab.key === "memory" && state.hotScope === "project" ? "project" : "user",
                            quickOpen: false,
                        });
                        // 图库是独立端点：进页签的这一刻就取状态和列表，用户不用再点「刷新」。
                        // 已有列表时保留并显示「更新中」，不把整格换成加载态。
                        if (tab.key === "gallery") {
                            loadGalleryStatus();
                            loadGallery({});
                        }
                        scrollPanelTop();
                    },
                    onMouseEnter: () => patchState({ hoverTab: tab.key }),
                    onMouseLeave: () => patchState({ hoverTab: null }),
                }, tab.label);
            }));

            // 二级不再问「从哪个根来」，而是问「这条记忆是给谁的」：
            //   个人偏好 = target:user，写全局，跨工作区生效；
            //   项目记忆 = 其它 target，只写当前工作区；
            //   台账是交付物流水，单独一档。
            // 这样「全局被项目事实填满」在界面上就不可能发生，用户也不用理解 store 的根。
            const hotUserCount = finite(counts.hotUser);
            const hotProjectCount = finite(counts.hotProject);
            const layerDefs = [
                { key: "preference", label: "个人偏好", count: hotUserCount === null ? baseItems.filter((item) => item.target === "user").length : hotUserCount },
                { key: "project", label: "项目记忆", count: hotProjectCount === null ? baseItems.filter((item) => item.target !== "user").length : hotProjectCount },
                { key: "ledger", label: "台账", count: finite(counts.ledger) },
            ];
            const layerRow = section === "memory" ? h("div", { style: S.memFilterGroup, "data-memory-layers": "true" },
                layerDefs.map((layer) => {
                    const on = layerKey === layer.key;
                    return h("button", {
                        key: layer.key,
                        type: "button",
                        className: "om-btn",
                        style: on ? Object.assign({}, S.memOriginSeg, S.memOriginSegOn) : S.memOriginSeg,
                        role: "tab",
                        "aria-selected": on,
                        "data-layer": layer.key,
                        "aria-pressed": on,
                        onClick: () => {
                            if (on) return;
                            if (layer.key === "ledger") patchState({ tab: "ledger", visible: LIST_PAGE, expandedItem: null, quickOpen: false });
                            else patchState({ tab: "hot", hotScope: layer.key, visible: LIST_PAGE, expandedItem: null, quickTarget: layer.key === "project" ? "project" : "user", quickOpen: false });
                            scrollPanelTop();
                        },
                    },
                        h("span", { style: Object.assign({}, S.memTabDot, { background: metricColor(layer.key === "ledger" ? "ledger" : layer.key === "archive" ? "archive" : "hot") }) }),
                        layer.label + "（" + (layer.count === null ? "—" : layer.count) + "）",
                    );
                }),
            ) : null;

            const filters = section === "memory" ? h("div", { style: S.memFilters, "data-memory-filters": "true" }, layerRow) : null;

            // 头部 / 搜索不吸顶（用户要「滚动不必始终保持标题」），正文一上来就顶到视口边；
            // 只有页签 + 二级筛选留在 sticky 里，读到中间仍能换层。
            const topBar = h("div", { style: S.browserSticky, "data-memory-head": "top" }, head, controls);
            const stickyNav = h("div", { style: S.browserNav, "data-memory-nav": "sticky", "data-memory-head": "nav" }, tabs, filters);

            // 第二级已经把 target 分成「个人偏好 / 项目记忆」，列表里不再套一层同名分组
            // （以前的「用户偏好 / 项目与环境」分组会让同一件事说两遍）。
            const hotGroups = null;
            const hotTotals = new Map();

            // ── 「再显示 N 条」 ──
            //
            // 端点上各层已经有上限（200 / 200 / 500 / 60），所以这里只是**渐进式展开
            // 已经取回来的那一份**，不需要给端点加 offset —— 少一个参数就少一条要守
            // 的信任边界。分页决定「DOM 里有几条」，面板那一个滚动容器决定「一屏看几条」。
            const total = items.length;
            const shown = Math.max(0, Math.min(total, Math.max(LIST_PAGE, state.visible)));
            const hidden = Math.max(0, total - shown);
            // 面板根元素（listRef 指向它）。第六十二轮修好滚动之后它**不是**滚动容器了
            // （见 S.browser），所以「再显示 N 条」不能再对它 scrollTo —— 要找**真正能滚的那一层**。
            const showMore = () => {
                patchState({ visible: Math.min(total, shown + LIST_PAGE) });
                // 新露出来的那几条在列表底部：把它滚过去，用户才看得到「刚才多了什么」。
                // 平滑与否要看系统设置 —— `scrollTo({behavior:'smooth'})` 不读 CSS 的
                // `scroll-behavior`，所以 reduced-motion 必须在代码里再读一次。
                const el = scrollParentOf(listRef.current);
                if (el) {
                    try {
                        el.scrollTo({ top: el.scrollHeight, behavior: prefersReducedMotion() ? "auto" : "smooth" });
                    } catch (error) {
                        // 老浏览器不认 options（或连 scrollTo 都没有）：直接给 scrollTop。
                        try { el.scrollTop = el.scrollHeight; } catch (inner) { /* 滚不动就算了 */ }
                    }
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
                    // 知识库文档 → 块清单（第八十二轮）：这是一次只读下钻，不是写入指令，
                    // 所以走 load 而不是 pending 卡；块清单一律从服务端现取，面板不自己切。
                    onOpenDoc: tab === "kb"
                        ? (entry) => {
                            load({ kbDoc: String(entry && entry.id ? entry.id : ""), kbChunk: "", kbMode: "docs" });
                            scrollPanelTop();
                        }
                        : undefined,
                    onEdit: tab === "hot" ? openHotEditor : (tab === "kb" ? openKbEditor : undefined),
                    onDelete: tab === "kb" ? (entry) => askDeleteDoc(entry) : undefined,
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

            const detailPool = items.map((item) => ({ item: item, tab: state.tab }));
            const detailIndex = state.hoverItem === null
                ? -1
                : detailPool.findIndex((entry, index) => String(itemKeyOf(entry.item, index)) === String(state.hoverItem));
            const detailEntry = detailIndex >= 0 ? detailPool[detailIndex] : null;
            const detailItem = detailEntry === null ? null : detailEntry.item;
            const detailBar = detailItem === null ? null : h("div", {
                style: S.itemDetail, role: "status",
                "data-item-detail": String(itemKeyOf(detailItem, detailIndex)),
            },
                h("div", { style: S.itemDetailHead },
                    h("span", { style: S.itemDetailKey }, "ITEM DETAIL"),
                    h("span", { style: S.itemDetailTitle }, String(itemKeyOf(detailItem, detailIndex))),
                ),
                h("div", { style: S.itemDetailBody }, detailTextOf(detailItem, detailEntry.tab)),
                h("div", { style: S.itemDetailHint }, "完整内容。卡片上的「展开」可以就地铺开。"),
            );

            // 草稿卡（第八十五轮；第八十七轮加直接新增）：热记忆改正文/目标/重要度，
            // 新增记忆填正文并直接 POST memory-add；知识库文档改标题/文件夹/来源档。
            // 只改元数据与正文，不在这里改块：块是内容寻址的派生数据，改正文后重切。
            const editorPanel = state.editor === null ? null : (() => {
                const editor = state.editor;
                const patchEditor = (patch) => patchState({ editor: Object.assign({}, latest.current.editor || {}, patch) });
                const creating = editor.kind === "new";
                const hot = editor.kind === "hot" || creating;
                const kbText = editor.kind === "kb-text";
                return h("div", { style: S.editor, "data-memory-editor": editor.kind },
                    h("div", { style: S.editorTitle }, creating ? "新增记忆"
                        : hot ? "编辑热记忆"
                            : kbText ? ("编辑文档正文 · " + String(editor.path || ""))
                                : ("编辑文档信息 · " + String(editor.path || ""))),
                    creating
                        ? h("div", { style: S.editorHint }, "直接写入记忆，不再生成会话指令。写一两句稳定的偏好、约定或事实。")
                        : null,
                    hot
                        ? h(React.Fragment, null,
                            h("label", { style: S.editorLabel }, "正文",
                                h("textarea", {
                                    style: S.editorTextarea, "data-memory-editor-text": "true",
                                    value: editor.content || "",
                                    placeholder: creating ? "例如：图表默认用蓝灰色系，坐标轴不写单位。" : undefined,
                                    onChange: (event) => patchEditor({ content: String(event.target.value || "") }),
                                })),
                            h("div", { style: S.editorRow },
                                h("label", { style: S.editorLabel }, "归到",
                                    h("select", {
                                        style: S.editorSelect, "data-memory-editor-target": "true",
                                        value: editor.target || "project",
                                        onChange: (event) => patchEditor({ target: String(event.target.value || "project") }),
                                    },
                                        h("option", { value: "user" }, "用户偏好"),
                                        h("option", { value: "project" }, "项目与环境"))),
                                h("label", { style: S.editorLabel }, "重要度",
                                    h("select", {
                                        style: S.editorSelect, "data-memory-editor-importance": "true",
                                        value: editor.importance || "normal",
                                        onChange: (event) => patchEditor({ importance: String(event.target.value || "normal") }),
                                    },
                                        h("option", { value: "critical" }, "critical"),
                                        h("option", { value: "normal" }, "normal"),
                                        h("option", { value: "low" }, "low")))))
                        : kbText ? h(React.Fragment, null,
                            h("label", { style: S.editorLabel }, "正文",
                                h("textarea", {
                                    style: Object.assign({}, S.editorTextarea, { minHeight: "220px" }),
                                    "data-kb-editor-text": "true",
                                    value: editor.text || "",
                                    onChange: (event) => patchEditor({ text: String(event.target.value || "") }),
                                })),
                            h("div", { style: S.editorHint }, "保存后按新正文重切块；块 id 会换代，指向旧块的边自然失效。"))
                            : h(React.Fragment, null,
                            h("label", { style: S.editorLabel }, "标题",
                                h("input", {
                                    type: "text", style: S.editorInput, "data-memory-editor-title": "true",
                                    value: editor.title || "",
                                    onChange: (event) => patchEditor({ title: String(event.target.value || "") }),
                                })),
                            h("div", { style: S.editorRow },
                                h("label", { style: S.editorLabel }, "虚拟文件夹",
                                    h("input", {
                                        type: "text", style: S.editorInput, "data-memory-editor-folder": "true",
                                        value: editor.folder || "", placeholder: "留空 = 根，例如 参考/规范",
                                        onChange: (event) => patchEditor({ folder: String(event.target.value || "") }),
                                    })),
                                h("label", { style: S.editorLabel }, "来源档",
                                    h("select", {
                                        style: S.editorSelect, "data-memory-editor-tier": "true",
                                        value: editor.tier || "user",
                                        onChange: (event) => patchEditor({ tier: String(event.target.value || "user") }),
                                    },
                                        h("option", { value: "user" }, "user"),
                                        h("option", { value: "unverified" }, "unverified"),
                                        h("option", { value: "verified" }, "verified"))))),
                    h("div", { style: S.editorActions },
                        h("button", {
                            type: "button", className: "om-btn om-primary", style: S.memPrimary,
                            "data-memory-save": "true", disabled: state.saving === true,
                            onClick: saveEditor,
                        }, state.saving === true ? "保存中…" : (creating ? "记下" : "保存")),
                        h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton,
                            "data-memory-edit-cancel": "true",
                            onClick: () => patchState({ editor: null, actionError: "" }),
                        }, "取消")),
                    state.actionError
                        ? h("div", { style: S.editorError, "data-memory-action-error": "true" }, String(state.actionError))
                        : null,
                );
            })();

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

            // 知识库页签的同一件事：清单是服务端**真分页**（第九十五轮），这一页只装
            // limit 篇；把总数与窗口摆出来，页码由响应里的 kbPage 回写。块正文不在
            // 这一页：要读哪一块，用会话里的 kb-read（面板只给文档 id 与块数）。
            //
            // 措辞不写「最新的 N 篇」：`at` 撞在同一毫秒时排序是稳定排序的产物，说「最新」
            // 会是一句撑不住的承诺（复核 P2-4）。说「这一页列了 N 篇」，并把总数放在旁边。
            // 服务端严格读 folders.json 失败时会把原因带上来；不要把它显示成「没有目录」。
            const kbFoldersError = typeof data.kbFoldersError === "string" ? data.kbFoldersError.trim() : "";
            const kbTotal = finite(counts.kbTotal);
            const kbCount = finite(counts.kb);
            // 分页窗口：服务端给的是权威值（越界时它已经回拉到合法页）；读不到时退回本地推算。
            const kbPageInfo = data.kbPage && typeof data.kbPage === "object" ? data.kbPage : null;
            const kbPageTotal = kbPageInfo !== null && Number.isFinite(Number(kbPageInfo.total))
                ? Math.max(0, Number(kbPageInfo.total))
                : (kbCount === null ? total : kbCount);
            const kbPageOffset = kbPageInfo !== null && Number.isFinite(Number(kbPageInfo.offset))
                ? Math.max(0, Number(kbPageInfo.offset))
                : Math.max(0, Number(state.kbOffset) || 0);
            const kbPageLimit = kbPageInfo !== null && Number.isFinite(Number(kbPageInfo.limit))
                ? Math.max(1, Number(kbPageInfo.limit))
                : Math.max(1, Number(state.kbLimit) || LIST_PAGE);
            const kbPageHasPrev = kbPageInfo !== null ? kbPageInfo.hasPrev === true : kbPageOffset > 0;
            const kbPageHasNext = kbPageInfo !== null
                ? kbPageInfo.hasNext === true
                : kbPageOffset + total < kbPageTotal;
            const kbNote = state.tab === "kb" && total > 0
                ? h("div", { style: S.pendingHint, "data-kb-note": "true" },
                    "文档可直接上传、改标题与文件夹；上传件只留切片与图片，不再保留原文件。"
                    + (kbTotal === null ? "" : " 库里共 " + kbTotal + " 篇、"
                        + countText(counts.kbChunks) + " 块、"
                        + (humanBytes(counts.kbBytes) === null ? "—" : humanBytes(counts.kbBytes)) + "。"))
                : null;

            // ── 知识库的第二层（第八十二轮）──
            //
            // 文档清单只是入口：用户真正要问的是「哪一块讲了这件事」。两块能力都已在
            // 服务端（kbSearch / kbDoc / kbRead），这里只把它们接成一条可点的路径：
            // 文档 → 块清单 → 块正文，或 关键词 → 命中块 → 块正文。
            const kbSearchData = data.kbSearch && typeof data.kbSearch === "object" ? data.kbSearch : null;
            const kbSearchHits = kbSearchData && Array.isArray(kbSearchData.hits) ? kbSearchData.hits : [];
            const kbDocData = data.kbDoc && typeof data.kbDoc === "object" ? data.kbDoc : null;
            const kbDocChunks = kbDocData && Array.isArray(kbDocData.chunks) ? kbDocData.chunks : [];
            const kbChunkData = data.kbChunk && typeof data.kbChunk === "object" ? data.kbChunk : null;

            // ── 虚拟文件夹（第八十四轮）──
            //
            // 知识库是全局的一棵树：文件夹只是 manifest 里的 `folder` 字段，不是磁盘目录。
            // 左树负责「现在在哪一层」，右列负责这一层的文档与检索；当前文件夹同时是
            // 检索范围（搜索框只在当前子树里查），这样「大范围」与「小范围」是同一个控件。
            // 快照里的 kbTree 直接是根节点；kb-tree 动作返回的是 { root }。两种形状都认，
            // 免得面板对着一个永远空的树画「还没有入库的文档」。
            const kbTreeRoot = data.kbTree && typeof data.kbTree === "object"
                ? (data.kbTree.root && typeof data.kbTree.root === "object" ? data.kbTree.root : data.kbTree)
                : null;
            const kbFolderNow = typeof state.kbFolder === "string" ? state.kbFolder : "";
            const kbFolderName = kbFolderNow === "" ? "全部文档" : String(kbFolderNow.split("/").pop() || kbFolderNow);
            // 折叠状态只活在面板里：文件夹的真源仍是 manifest，展开/收起不写回任何东西。
            // 默认全展开，点小三角才收起；收起时子层不进 DOM，深树不会把左栏拉成一条长带。
            const kbClosed = state.kbClosed && typeof state.kbClosed === "object" ? state.kbClosed : {};
            const kbTreeRows = [];
            const walkKbTree = (node, depth) => {
                if (node === null || typeof node !== "object") return;
                const path = typeof node.path === "string" ? node.path : "";
                const on = path === kbFolderNow;
                const children = Array.isArray(node.children) ? node.children : [];
                const closed = children.length > 0 && kbClosed[path] === true;
                const toggle = children.length === 0 ? null : (event) => {
                    // 小三角只负责展开/收起，不换当前文件夹；事件不能再冒泡到整行的 onClick。
                    if (event && typeof event.stopPropagation === "function") event.stopPropagation();
                    patchState({ kbClosed: Object.assign({}, kbClosed, { [path]: !closed }) });
                };
                kbTreeRows.push(h("button", {
                    key: path === "" ? "__root" : path,
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.kbTreeRow, { paddingLeft: String(8 + depth * 15) + "px" }, on ? S.kbTreeRowOn : {}),
                    "data-kb-folder": path === "" ? "" : path,
                    "aria-pressed": on,
                    title: path === "" ? "全部文档" : path,
                    onClick: () => {
                        if (on) return;
                        kbGoFolder(path);
                    },
                },
                    h("span", {
                        style: Object.assign({}, S.kbTreeGlyph, { display: "inline-flex", alignItems: "center", justifyContent: "center" }, toggle === null ? {} : { cursor: "pointer" }),
                        "data-kb-toggle": toggle === null ? undefined : (closed ? "closed" : "open"),
                        onClick: toggle || undefined,
                    }, toggle === null ? (path === "" ? h(ExplorerGlyph, { kind: "folder", size: 13 }) : h("span", { style: { fontSize: "10px" } }, "·")) : (closed ? "▸" : "▾")),
                    h("span", { style: S.kbTreeName }, String(node.name || (path === "" ? "全部文档" : path))),
                    h("span", { style: S.kbTreeCount }, String(finite(node.docs) === null ? 0 : node.docs)),
                ));
                if (!closed) for (const child of children) walkKbTree(child, depth + 1);
            };
            if (kbTreeRoot !== null) walkKbTree(kbTreeRoot, 0);
            const kbTreeSide = h("div", { className: "om-kb-tree", style: S.kbTree, "data-kb-tree": "true" },
                h("div", { style: S.kbTreeHead }, "知识库 · 文件夹"),
                kbTreeRows.length > 0
                    ? kbTreeRows
                    : h("div", { style: S.kbTreeEmpty }, "还没有入库的文档"),
            );
            // 资源管理器口径：当前层的子文件夹与文件（文档）放在同一张清单里，
            // 文件夹在前、文件在后；文件夹行双击/单击都进入下一层，文件行保留既有操作
            // （看分块 / 编辑 / 删除），只是从卡片换成一行一列的「详细信息」式列表。
            const kbNodeAt = (node, path) => {
                if (node === null || typeof node !== "object") return null;
                if (String(node.path || "") === path) return node;
                for (const child of Array.isArray(node.children) ? node.children : []) {
                    const found = kbNodeAt(child, path);
                    if (found !== null) return found;
                }
                return null;
            };
            const kbCurrentNode = kbTreeRoot === null ? null : kbNodeAt(kbTreeRoot, kbFolderNow);
            const kbChildNodes = kbCurrentNode !== null && Array.isArray(kbCurrentNode.children) ? kbCurrentNode.children : [];

            // 导航栈：只有用户真的换了文件夹才入栈。后退恢复的是「上一层的路径」，
            // 不缓存清单本身 —— 目录内容可能刚被上传/删除改过，回来必须重新取。
            const kbGoFolder = (next, mode) => {
                const target = String(next || "");
                if (mode !== "back" && target === kbFolderNow) return;
                const stack = Array.isArray(state.kbBack) ? state.kbBack : [];
                const nextStack = mode === "back"
                    ? stack.slice(0, Math.max(0, stack.length - 1))
                    : stack.concat([kbFolderNow]);
                const patch = {
                    kbFolder: target, kbDoc: "", kbChunk: "", visible: LIST_PAGE,
                    expandedItem: null, kbSelectedKey: "", kbBack: nextStack,
                    // 换层 / 后退都回到第一页：沿用旧 offset 会落进新目录的空白页。
                    kbOffset: 0,
                };
                if (state.kbMode === "search" && state.appliedKb) patch.kbQuery = state.kbQuery || state.appliedKb;
                load(patch);
                scrollPanelTop();
            };
            const kbParentFolder = () => {
                const at = kbFolderNow.lastIndexOf("/");
                return at < 0 ? "" : kbFolderNow.slice(0, at);
            };

            // 新建虚拟文件夹：名字只占一层，落在当前目录下；只写 <kbDir>/folders.json，
            // 不建真实目录、不写 manifest。空目录也要能被树 / picker / /kb-scope 看见，
            // 所以服务端把它并进 kb-tree；这里成功后直接进入新目录（和资源管理器一致）。
            const kbCreateFolder = () => {
                if (state.saving === true) return;
                const draft = typeof state.kbNewFolder === "string" ? state.kbNewFolder.trim() : "";
                if (draft === "") {
                    patchState({ kbFolderError: "文件夹名不能为空。", actionError: "" });
                    return;
                }
                if (draft.includes("/") || draft.includes("\\")) {
                    patchState({ kbFolderError: "一次只能建一层：名字里不要带斜杠，深一层请进去再建。", actionError: "" });
                    return;
                }
                const target = kbFolderNow === "" ? draft : kbFolderNow + "/" + draft;
                mutatePanel("kb-folder-create", { path: target }, {
                    kbNewFolder: null, kbFolder: target, kbDoc: "", kbChunk: "", kbOffset: 0,
                    kbFolderError: "", kbFolderNotice: "", kbFolderRetry: null,
                }, { failureKey: "kbFolderError", confirmedFolder: target });
                scrollPanelTop();
            };

            // 排序键只影响「当前这一层」的呈现，不写回 manifest；目录始终在文件之前，
            // 名称用中文排序（numeric 让 2 排在 10 前面）。**文件排序在服务端**：
            // 前端只报 sort/sortDir，切片与 tie-break 由 kb-page.js 统一算，
            // 否则客户端只排当前页，跨页顺序会断掉。
            const kbSortSign = state.kbSortDir === "desc" ? -1 : 1;
            const kbCompareName = (left, right) => String(left || "").localeCompare(String(right || ""), "zh-Hans-CN", { numeric: true, sensitivity: "base" });
            const kbFolderRows = kbChildNodes.slice().sort((left, right) => kbCompareName(left.name || left.path || "", right.name || right.path || "") * kbSortSign).map((node) => h("button", {
                key: "folder:" + String(node.path || ""),
                type: "button", className: "om-btn",
                "data-kb-subfolder": String(node.path || ""),
                "aria-selected": state.kbSelectedKey === "folder:" + String(node.path || ""),
                title: "进入文件夹 " + String(node.path || ""),
                style: Object.assign({}, S.kbExplorerRow, {
                    cursor: "pointer", width: "100%", textAlign: "left",
                    background: state.kbSelectedKey === "folder:" + String(node.path || "") ? M.hover : "transparent",
                }),
                onClick: () => kbGoFolder(String(node.path || "")),
            },
                h("span", { style: { width: "18px", display: "inline-flex", justifyContent: "center", color: M.warn } }, h(ExplorerGlyph, { kind: "folder", size: 16 })),
                h("span", { style: S.kbExplorerName }, String(node.name || node.path || "")),
                h("span", { style: S.kbExplorerMeta }, finite(node.docs) === null || finite(node.docs) === 0 ? "空" : String(finite(node.docs)) + " 篇"),
                h("span", { style: S.kbExplorerMeta }, ""),
                h("span", { style: S.kbExplorerMeta }, "")));
            // 直接渲染服务端已经切好的这一页：不再本地 slice，否则「第 2 页」永远是空的。
            const kbVisibleDocs = items;
            const kbOpenDoc = (doc) => {
                load({ kbDoc: String(doc.id || ""), kbChunk: "", kbMode: "docs" });
                scrollPanelTop();
            };
            const kbFileRows = kbVisibleDocs.map((doc, index) => {
                const key = itemKeyOf(doc, index);
                const name = String(doc.path || doc.title || doc.id || "").split("/").pop();
                const ext = String(name).includes(".") ? String(name).split(".").pop().toLowerCase() : "";
                const iconKind = ["png", "jpg", "jpeg", "gif", "webp"].includes(ext) ? "image" : "file";
                const selected = state.kbSelectedKey === key;
                // 资源管理器的一行：图标 + 名称（标题在上、入库路径在下）+ 档位/块/大小/修改 + 操作。
                // 单击只选中，双击或回车才进入；操作按钮自己 stopPropagation，不触发选中。
                return h("div", {
                    key: key,
                    style: Object.assign({}, S.kbExplorerRow, {
                        cursor: "pointer",
                        background: selected ? M.hover : "transparent",
                    }),
                    "data-kb-file": String(doc.id || ""), "data-item-key": key,
                    tabIndex: 0,
                    "aria-selected": selected,
                    onMouseEnter: () => patchState({ hoverItem: key }),
                    onMouseLeave: () => patchState({ hoverItem: null }),
                    onFocus: () => patchState({ hoverItem: key }),
                    onClick: () => patchState({ kbSelectedKey: key }),
                    onDoubleClick: () => kbOpenDoc(doc),
                    onKeyDown: (event) => {
                        if (event.key !== "Enter" || event.target !== event.currentTarget) return;
                        event.preventDefault();
                        kbOpenDoc(doc);
                    },
                },
                    h("span", { style: { width: "18px", display: "inline-flex", justifyContent: "center", color: M.textTertiary } }, h(ExplorerGlyph, { kind: iconKind, size: 16 })),
                    h("div", { style: { minWidth: 0, display: "flex", flexDirection: "column", gap: "1px" } },
                        h("button", {
                            type: "button", className: "om-btn",
                            "data-kb-open": String(doc.id || ""),
                            title: "看这份文档切成了哪些块",
                            style: S.kbExplorerNameBtn,
                            onClick: (event) => { if (event && event.stopPropagation) event.stopPropagation(); kbOpenDoc(doc); },
                        }, String(doc.title || name)),
                        h("span", {
                            style: { fontFamily: M.mono, fontSize: "10.5px", color: M.caption, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
                        }, String(doc.path || "")),
                    ),
                    h("span", { style: S.kbExplorerMeta }, String(doc.tier || "user")),
                    h("span", { style: S.kbExplorerMeta }, String(doc.chunks || 0) + " 块"),
                    h("span", { style: S.kbExplorerMeta }, humanBytes(doc.bytes) === null ? "—" : humanBytes(doc.bytes)),
                    h("span", { style: S.kbExplorerMeta }, String(doc.at || "").slice(0, 10)),
                    h("span", { style: { display: "flex", gap: "4px", justifyContent: "flex-end" } },
                        h("button", {
                            type: "button", className: "om-act", style: S.itemAction,
                            "data-memory-edit": String(doc.id || ""),
                            onClick: (event) => { if (event && event.stopPropagation) event.stopPropagation(); openKbEditor(doc); },
                        }, "编辑"),
                        h("button", {
                            type: "button", className: "om-act",
                            style: Object.assign({}, S.itemAction, { color: M.error }),
                            "data-kb-delete": String(doc.id || ""),
                            onClick: (event) => { if (event && event.stopPropagation) event.stopPropagation(); askDeleteDoc(doc); },
                        }, "删除")));
            });
            const kbSortHeader = (label, key) => h("button", {
                type: "button", className: "om-btn",
                style: Object.assign({}, S.kbExplorerMeta, { textAlign: "left", background: "transparent", border: "none", padding: 0, cursor: "pointer" }),
                "data-kb-sort": key,
                "aria-pressed": state.kbSort === key,
                onClick: () => patchState({ kbSort: key, kbSortDir: state.kbSort === key && state.kbSortDir === "asc" ? "desc" : "asc" }),
            }, label + (state.kbSort === key ? (state.kbSortDir === "asc" ? " ↑" : " ↓") : ""));
            const kbExplorerHead = h("div", { style: S.kbExplorerHead, "data-kb-file-head": "true" },
                h("span", { style: { width: "18px" } }, ""),
                kbSortHeader("名称", "name"),
                h("span", { style: S.kbExplorerMeta }, "档位"),
                kbSortHeader("块", "chunks"),
                kbSortHeader("大小", "size"),
                kbSortHeader("修改", "time"),
                h("span", { style: Object.assign({}, S.kbExplorerMeta, { textAlign: "right" }) }, "操作"));

            // 大图标视图与详细信息视图共用同一份排序/选中状态：换视图不丢选中。
            const kbFolderTiles = kbChildNodes.slice().sort((left, right) => kbCompareName(left.name || left.path || "", right.name || right.path || "") * kbSortSign).map((node) => {
                const key = "folder:" + String(node.path || "");
                const selected = state.kbSelectedKey === key;
                return h("button", {
                    key: key, type: "button", className: "om-btn",
                    "data-kb-subfolder": String(node.path || ""),
                    style: Object.assign({}, S.explorerTile, selected ? S.explorerTileOn : {}),
                    onClick: () => {
                        patchState({ kbSelectedKey: key });
                        kbGoFolder(String(node.path || ""));
                    },
                },
                    h("span", { style: Object.assign({}, S.explorerTileIcon, { color: M.warn }) }, h(ExplorerGlyph, { kind: "folder", size: 34 })),
                    h("span", { style: S.explorerTileName }, String(node.name || node.path || "")),
                    h("span", { style: S.explorerTileMeta }, (finite(node.docs) === null ? 0 : node.docs) + " 篇"));
            });
            const kbFileTiles = kbVisibleDocs.map((doc, index) => {
                const key = itemKeyOf(doc, index);
                const name = String(doc.path || doc.title || doc.id || "").split("/").pop();
                const selected = state.kbSelectedKey === key;
                return h("button", {
                    key: key, type: "button", className: "om-btn",
                    "data-kb-file": String(doc.id || ""), "data-item-key": key,
                    style: Object.assign({}, S.explorerTile, selected ? S.explorerTileOn : {}),
                    onClick: () => patchState({ kbSelectedKey: key }),
                    onDoubleClick: () => kbOpenDoc(doc),
                    // 磁贴也是按钮：Enter 直接打开。没有这一条时，键盘用户只能
                    // 选中磁贴，无法进入文档（单击只选中，双击才打开）。
                    onKeyDown: (event) => {
                        if (!event || event.key !== "Enter") return;
                        event.preventDefault();
                        kbOpenDoc(doc);
                    },
                },
                    h("span", { style: S.explorerTileIcon }, h(ExplorerGlyph, { kind: "file", size: 34 })),
                    h("span", { style: S.explorerTileName }, String(doc.title || name)),
                    h("span", { style: S.explorerTileMeta }, String(doc.chunks || 0) + " 块 · " + (humanBytes(doc.bytes) === null ? "—" : humanBytes(doc.bytes))));
            });
            const kbExplorerList = state.kbView === "tiles"
                ? h("div", { style: S.explorerGrid, "data-memory-list": "kb", "data-kb-files": "true", "data-kb-view": "tiles" },
                    kbFolderTiles.concat(kbFileTiles))
                : h("div", { style: S.kbExplorerList, "data-memory-list": "kb", "data-kb-files": "true", "data-kb-view": "details" },
                    kbFolderRows.length === 0 && kbFileRows.length === 0
                        ? null
                        : h(React.Fragment, null, kbExplorerHead, kbFolderRows, kbFileRows));

            const kbSelectedDoc = kbVisibleDocs.find((doc, index) => itemKeyOf(doc, index) === state.kbSelectedKey) || null;
            const kbSelectedLabel = state.kbSelectedKey === ""
                ? "未选中"
                : state.kbSelectedKey.startsWith("folder:")
                    ? "已选文件夹 · " + state.kbSelectedKey.slice(7)
                    : kbSelectedDoc === null ? "已选中 1 项" : "已选 · " + String(kbSelectedDoc.title || kbSelectedDoc.path || "");
            // 分页控件（第九十五轮）：服务端真分页后，这一栏就是唯一的「翻页」入口。
            // 文案用「第 a–b 篇 / 共 N 篇」，不写「已截断」——没有截断，只有窗口。
            const kbGoPage = (nextOffset) => {
                const target = Math.max(0, Number(nextOffset) || 0);
                if (target === kbPageOffset) return;
                load({ kbOffset: target });
            };
            const kbPageFrom = kbPageTotal === 0 ? 0 : kbPageOffset + 1;
            const kbPageTo = kbPageOffset + items.length;
            const kbStatusBar = h("div", { style: S.explorerStatus, "data-kb-status": "true" },
                h("span", null, "文件夹 " + String(kbChildNodes.length) + " · 文档 " + String(kbPageTotal)),
                h("span", null, kbSelectedLabel),
                h("span", { style: { display: "inline-flex", alignItems: "center", gap: "6px" }, "data-kb-pager": "true" },
                    h("button", {
                        type: "button", className: "om-btn", style: Object.assign({}, S.explorerTool, kbPageHasPrev ? {} : S.explorerToolDisabled),
                        "data-kb-page-prev": "true", disabled: !kbPageHasPrev,
                        title: kbPageHasPrev ? "上一页" : "已经是第一页",
                        onClick: () => kbGoPage(kbPageOffset - kbPageLimit),
                    }, "‹ 上一页"),
                    h("span", { "data-kb-page-info": "true", style: { color: M.caption } },
                        "第 " + String(kbPageFrom) + "–" + String(kbPageTo) + " 篇 / 共 " + String(kbPageTotal) + " 篇"),
                    h("button", {
                        type: "button", className: "om-btn", style: Object.assign({}, S.explorerTool, kbPageHasNext ? {} : S.explorerToolDisabled),
                        "data-kb-page-next": "true", disabled: !kbPageHasNext,
                        title: kbPageHasNext ? "下一页" : "已经是最后一页",
                        onClick: () => kbGoPage(kbPageOffset + kbPageLimit),
                    }, "下一页 ›"),
                    h("label", { style: { display: "inline-flex", alignItems: "center", gap: "4px", fontSize: "11px", color: M.caption } },
                        "每页",
                        h("select", {
                            style: S.explorerSort, "data-kb-page-size": "true",
                            value: String(kbPageLimit),
                            onChange: (event) => {
                                const size = Math.max(1, Number(event.target.value) || LIST_PAGE);
                                load({ kbLimit: size, kbOffset: 0 });
                            },
                        },
                            PAGE_SIZES.map((size) => h("option", { key: String(size), value: String(size) }, String(size)))))),
                h("span", { style: { marginLeft: "auto" } }, state.kbView === "tiles" ? "大图标" : "详细信息"));
            const kbCrumbParts = kbFolderNow === "" ? [] : kbFolderNow.split("/");
            const kbCrumb = h("div", { style: S.kbCrumb, "data-kb-crumb": kbFolderNow },
                h("button", {
                    type: "button", className: "om-btn", style: S.kbCrumbBtn,
                    "data-kb-crumb-root": "true",
                    onClick: () => { if (kbFolderNow === "") return; kbGoFolder(""); },
                }, "知识库"),
                kbCrumbParts.map((part, index) => {
                    const path = kbCrumbParts.slice(0, index + 1).join("/");
                    return h(React.Fragment, { key: path },
                        h("span", { style: S.kbCrumbSep }, "/"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbCrumbBtn,
                            "data-kb-crumb": path,
                            onClick: () => { if (path === kbFolderNow) return; kbGoFolder(path); },
                        }, part),
                    );
                }),
                h("button", {
                    type: "button", className: "om-btn", style: S.kbRecursive,
                    "data-kb-recursive": state.kbRecursive === false ? "current" : "subtree",
                    title: "切换当前文件夹的检索范围",
                    onClick: () => {
                        const next = state.kbRecursive === false;
                        load({ kbRecursive: next, kbDoc: "", kbChunk: "" });
                        scrollPanelTop();
                    },
                }, state.kbRecursive === false ? "仅当前层" : "含子文件夹"),
                // 单一入口（第九十五轮）：上传与识别合并成「＋ 添加文件」。
                // 选完先**只预览**（不落库、不调模型），在预览卡上再显式点「入库」；
                // 扫描件 / 图片要读文字时，才在预览卡上点「识别文字」——OCR 绝不自动发。
                h("button", {
                    type: "button", className: "om-btn", style: S.kbIngestBtn,
                    "data-kb-add-files": "true",
                    title: "选文件先预览：确认后再入库；扫描件要识别文字时再点识别",
                    disabled: Boolean(state.kbUploading),
                    onClick: pickFiles,
                }, state.kbUploading ? ("处理中 · " + String(state.kbUploading).slice(0, 16)) : "＋ 添加文件"),
                h("input", {
                    type: "file", multiple: true, "data-kb-upload-input": "true",
                    style: { display: "none" },
                    onChange: (event) => {
                        const files = event && event.target && event.target.files;
                        previewFiles(files);
                        if (event && event.target) event.target.value = "";
                    },
                }),
            );

            // ── Windows 11 资源管理器式命令行 ──
            // 只放真实实现：后退用本层面板的历史栈，上一级由路径推出，刷新重新取数；
            // 新建文件夹写 <kbDir>/folders.json（虚拟目录，不动真实磁盘）；重命名 / 复制没有入口，不做假按钮。
            const kbBackStack = Array.isArray(state.kbBack) ? state.kbBack : [];
            const kbCommandBar = h("div", { style: S.explorerBar, "data-kb-commandbar": "true", role: "toolbar", "aria-label": "知识库命令行" },
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, kbBackStack.length === 0 ? S.explorerToolDisabled : {}),
                    "data-kb-nav": "back", disabled: kbBackStack.length === 0,
                    title: kbBackStack.length === 0 ? "没有可后退的记录" : "后退到 " + String(kbBackStack[kbBackStack.length - 1] || "全部文档"),
                    onClick: () => { if (kbBackStack.length === 0) return; kbGoFolder(kbBackStack[kbBackStack.length - 1], "back"); },
                }, "← 后退"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, kbFolderNow === "" ? S.explorerToolDisabled : {}),
                    "data-kb-nav": "up", disabled: kbFolderNow === "",
                    title: kbFolderNow === "" ? "已经在根目录" : "上一级",
                    onClick: () => { if (kbFolderNow !== "") kbGoFolder(kbParentFolder()); },
                }, "↑ 上一级"),
                h("button", {
                    type: "button", className: "om-btn", style: S.explorerTool,
                    "data-kb-nav": "refresh", title: "重新读取当前文件夹",
                    onClick: () => { load({}); scrollPanelTop(); },
                }, "刷新"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, state.saving ? S.explorerToolDisabled : {}),
                    "data-kb-nav": "new-folder", disabled: state.saving === true,
                    title: "在当前文件夹下新建一个空文件夹",
                    onClick: () => patchState({ kbNewFolder: "", kbFolderError: "", actionError: "" }),
                }, "新建文件夹"),
                h("span", { style: { flex: "1 1 auto" } }),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, state.kbView === "details" ? S.explorerToolOn : {}),
                    "data-kb-view": "details", "aria-pressed": state.kbView === "details",
                    title: "详细信息",
                    onClick: () => patchState({ kbView: "details" }),
                }, "详细信息"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, state.kbView === "tiles" ? S.explorerToolOn : {}),
                    "data-kb-view": "tiles", "aria-pressed": state.kbView === "tiles",
                    title: "大图标",
                    onClick: () => patchState({ kbView: "tiles" }),
                }, "大图标"),
                h("label", { style: { display: "inline-flex", alignItems: "center", gap: "5px", fontSize: "11px", color: M.caption } },
                    "排序",
                    h("select", {
                        style: S.explorerSort, "data-kb-sort-select": "true",
                        value: state.kbSort,
                        // 排序是服务端行为（跨页一致）：改了就回第一页重新取。
                        onChange: (event) => load({ kbSort: String(event.target.value || "name"), kbOffset: 0 }),
                    },
                        h("option", { value: "name" }, "名称"),
                        h("option", { value: "chunks" }, "块数"),
                        h("option", { value: "size" }, "大小"),
                        h("option", { value: "time" }, "修改时间"))),
                h("button", {
                    type: "button", className: "om-btn", style: S.explorerTool,
                    "data-kb-sort-dir": state.kbSortDir,
                    title: state.kbSortDir === "asc" ? "升序（点击改为降序）" : "降序（点击改为升序）",
                    onClick: () => load({ kbSortDir: state.kbSortDir === "asc" ? "desc" : "asc", kbOffset: 0 }),
                }, state.kbSortDir === "asc" ? "↑" : "↓"));

            // 行内新建条：不弹窗、不离开当前层。Enter 提交 / Esc 取消，和资源管理器一致。
            const kbNewFolderBar = state.kbNewFolder === null || state.kbNewFolder === undefined
                ? null
                : h("div", {
                    style: Object.assign({}, S.kbSelectBar, { borderStyle: "solid", marginBottom: "10px" }),
                    "data-kb-new-folder": "true",
                },
                    h("span", { style: { color: M.caption, fontSize: "12px" } },
                        "在「" + (kbFolderNow === "" ? "全部文档" : kbFolderName) + "」下新建文件夹"),
                    h("input", {
                        type: "text", className: "om-input", "data-kb-new-folder-input": "true",
                        value: String(state.kbNewFolder || ""), placeholder: "文件夹名（一次一层）",
                        autoFocus: true,
                        onChange: (event) => patchState({ kbNewFolder: String(event.target.value), kbFolderError: "" }),
                        onKeyDown: (event) => {
                            if (event.key === "Enter") {
                                event.preventDefault();
                                kbCreateFolder();
                            } else if (event.key === "Escape") {
                                event.preventDefault();
                                patchState({ kbNewFolder: null, kbFolderError: "", actionError: "" });
                            }
                        },
                    }),
                    h("button", {
                        type: "button", className: "om-btn", style: S.kbIngestBtn,
                        "data-kb-new-folder-ok": "true",
                        onClick: kbCreateFolder,
                    }, "新建"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.explorerTool,
                        "data-kb-new-folder-cancel": "true",
                        onClick: () => patchState({ kbNewFolder: null, kbFolderError: "", actionError: "" }),
                    }, "取消"));

            const kbFolderErrorBar = state.kbFolderError !== ""
                ? h("div", {
                    style: S.editorError, "data-kb-folder-error": "true",
                    // 沿用既有错误条的选择器，失败原因只有一个来源，不会和编辑器里的错误位重复。
                    "data-memory-action-error": "true",
                }, String(state.kbFolderError))
                : null;
            const kbFolderNoticeBar = state.kbFolderNotice !== ""
                ? h("div", { style: S.editorError, "data-kb-folder-notice": "true" },
                    String(state.kbFolderNotice),
                    state.kbFolderRetry !== null && typeof state.kbFolderRetry === "object" && String(state.kbFolderRetry.folder || "") !== ""
                        ? h("button", {
                            type: "button", className: "om-btn", style: S.explorerTool,
                            "data-kb-folder-retry": "true",
                            onClick: () => {
                                const retry = state.kbFolderRetry;
                                if (!retry || !retry.folder) return;
                                const retryPatch = { kbFolder: String(retry.folder) };
                                if (Array.isArray(retry.back)) retryPatch.kbBack = retry.back.slice();
                                load(retryPatch, {
                                    confirmedFolder: retry.confirmed === true ? String(retry.folder) : "",
                                });
                                scrollPanelTop();
                            },
                        }, "重试打开")
                        : null)
                : null;
            const kbFoldersErrorBar = state.tab === "kb" && kbFoldersError !== ""
                ? h("div", { style: S.editorError, "data-kb-folders-error": "true" },
                    "知识库文件夹清单读不出来：" + kbFoldersError + "。现有目录没有被改动，修好 folders.json 后点「刷新」。")
                : null;

            // ── AI 建库开关（默认关）──
            // 勾选会把切好的原文发给已配置的模型（宿主或 OpenAI 兼容端点），可能计费；
            // 模型未就绪时置灰并显示原因，不做「看起来能点、实际静默降级」。
            const kbLlmCfg = data && data.config && data.config.kbLlm ? data.config.kbLlm : null;
            const kbAiReady = kbLlmCfg !== null && kbLlmCfg.ready === true;
            const kbCatalog = state.kbAiCatalog && typeof state.kbAiCatalog === "object" ? state.kbAiCatalog : null;
            const kbModelOptions = [];
            if (kbLlmCfg !== null && kbLlmCfg.mode === "host" && kbCatalog !== null && Array.isArray(kbCatalog.providers)) {
                for (const provider of kbCatalog.providers) {
                    const providerId = String(provider && provider.provider || provider && provider.id || "");
                    for (const model of Array.isArray(provider && provider.models) ? provider.models : []) {
                        const modelId = String(model && model.id || "");
                        if (modelId === "") continue;
                        kbModelOptions.push({ value: providerId + "/" + modelId, label: (providerId || "宿主") + " / " + String(model && model.label || modelId) });
                    }
                }
            }
            // ref 必须在任何读取 .current 的渲染表达式之前声明（暂停卡在下面就要读它）。
            const kbAiPendingRef = React.useRef(null);
            const kbAiBar = h("div", {
                style: Object.assign({}, S.kbSelectBar, { borderStyle: "solid", marginBottom: "10px" }),
                "data-kb-ai": "true",
            },
                h("label", { style: { display: "inline-flex", alignItems: "center", gap: "7px", cursor: kbAiReady ? "pointer" : "default", color: M.textSecondary } },
                    h("input", {
                        type: "checkbox", "data-kb-ai-toggle": "true",
                        checked: state.kbAi === true, disabled: !kbAiReady,
                        onChange: (event) => {
                            const on = Boolean(event.target.checked);
                            if (on && !kbAiReady) return;
                            patchState({ kbAi: on });
                        },
                    }),
                    "上传时用模型切片 / 摘要（发原文给配置的模型，可能计费）"),
                kbLlmCfg === null
                    ? h("span", { style: { color: M.caption } }, "未读到 AI 建库配置；默认关闭。")
                    : kbLlmCfg.mode === "off"
                        ? h("span", { style: { color: M.caption } }, "AI 建库未启用（默认确定性切块）。")
                        : kbLlmCfg.mode === "host" && kbAiReady
                            ? h("select", {
                                style: S.editorSelect, "data-kb-ai-model": "true",
                                value: state.kbAiModel,
                                // 选模型只是选模型：不自动打开左侧的 AI 开关（是否发原文由用户显式勾选）。
                                onChange: (event) => patchState({ kbAiModel: String(event.target.value || "") }),
                            },
                                h("option", { value: "" }, "宿主默认模型"),
                                kbModelOptions.map((option) => h("option", { key: option.value, value: option.value }, option.label)))
                            : kbLlmCfg.mode === "openai" && kbAiReady
                                // OpenAI 兼容端点：模型名由用户手填（端点 / Key 在设置里），
                                // 留空就用设置里那个；不做下拉，因为宿主列不出远端模型。
                                ? h("input", {
                                    type: "text", style: S.editorInput, "data-kb-ai-model": "true",
                                    value: state.kbAiModel,
                                    placeholder: String(kbLlmCfg.model || "模型名，如 gpt-4o-mini"),
                                    "aria-label": "OpenAI 兼容模型名",
                                    onChange: (event) => patchState({ kbAiModel: String(event.target.value || "") }),
                                })
                                : h("span", { style: { color: M.warn } }, String(kbLlmCfg.reason || "模型未就绪，AI 建库不可用。")),
                // host 通道的清单读取状态：RPC 失败只显示错误并给重试，绝不改走 HTTP。
                kbLlmCfg !== null && kbLlmCfg.mode === "host" && state.kbAiLoading === true
                    ? h("span", { style: { color: M.caption }, "data-kb-ai-catalog-loading": "true" }, "正在读取宿主模型清单…")
                    : null,
                kbLlmCfg !== null && kbLlmCfg.mode === "host" && kbCatalog !== null && kbCatalog.ok === false
                    ? h("span", { style: { color: M.warn }, "data-kb-ai-catalog-error": "true" },
                        String(kbCatalog.error || "读取宿主模型清单失败。"),
                        h("button", {
                            type: "button", className: "om-btn",
                            style: Object.assign({}, S.ghostButton, { marginLeft: "6px" }),
                            "data-kb-ai-catalog-retry": "true",
                            onClick: () => { loadModelCatalog(true); },
                        }, "重试"))
                    : null,
                kbLlmCfg !== null && kbLlmCfg.mode === "host" && kbCatalog !== null
                    && Array.isArray(kbCatalog.failures) && kbCatalog.failures.length > 0
                    ? h("span", { style: { color: M.caption }, "data-kb-ai-catalog-failures": "true" },
                        "部分 provider 的模型清单没读到："
                        + kbCatalog.failures.map((item) => {
                            const label = String(item && (item.name || item.id) || "未知 provider");
                            return label + "：" + String(item && item.message || "宿主没有给出原因");
                        }).join("；"))
                    : null,
                kbLlmCfg !== null && kbLlmCfg.mode === "openai" && kbAiReady
                    ? h("span", { style: { color: M.caption } }, "留空 = 设置里的模型")
                    : null,
                kbAiReady && state.kbAi !== true
                    ? h("span", { style: { color: M.caption }, "data-kb-ai-model-hint": "true" },
                        "选模型不会自动开启：要勾选左侧复选框才会调用模型。")
                    : null,
                state.kbAi === true && kbAiReady
                    ? h("span", { style: { color: M.warn } }, "本次上传会调用模型；失败会回退为确定性切块并在回执里说明。")
                    : null);

            // 预估卡：纯计算，不落库、不调模型；数字来自服务端 estimateKbLlmJob，
            // 这里只展示 + 给「开始入库 / 取消」。开始入库重发同一份 base64。
            const kbEstimate = state.kbEstimate && typeof state.kbEstimate === "object" ? state.kbEstimate : null;
            const kbEstimateAi = kbEstimate !== null && kbEstimate.ai && typeof kbEstimate.ai === "object" ? kbEstimate.ai : null;
            // 只有「通道不可用」才拦真实入库；duplicate 是 0 次模型调用但仍可搬 / 换档。
            const kbEstimateUnavailable = kbEstimateAi !== null && kbEstimateAi.available === false;
            const kbEstimateDuplicate = kbEstimate !== null && kbEstimate.duplicate === true;
            const kbEstimateReason = kbEstimateUnavailable
                ? (String(kbEstimateAi.reason || "").trim() || "宿主没有可用的 llm 服务。")
                : "";
            const kbEstimatePanel = kbEstimate === null ? null : h("div", {
                style: Object.assign({}, S.panel, { padding: "12px 14px" }), "data-kb-estimate": "true",
            },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "AI 入库预估：", String(kbEstimate.name || "")),
                    h("span", { style: S.panelNote }, "未写入任何内容")),
                h("div", {
                    style: Object.assign({}, S.note, { color: M.textSecondary }),
                    "data-kb-estimate-folder": String(kbEstimate.folder || ""),
                }, "目标文件夹：「" + (String(kbEstimate.folder || "") === "" ? "全部文档" : String(kbEstimate.folder)) + "」"),
                h("div", {
                    style: Object.assign({}, S.note, { color: M.textSecondary }),
                    "data-kb-estimate-model": String(kbEstimate.aiModel || ""),
                }, "模型：" + (String(kbEstimate.aiModel || "") === "" ? "设置里的默认模型" : String(kbEstimate.aiModel))),
                kbEstimateUnavailable
                    ? h("div", {
                        style: Object.assign({}, S.note, { color: M.warn }),
                        "data-kb-estimate-unavailable": "true",
                    }, "模型通道不可用：" + kbEstimateReason
                        + (/[。！？]$/.test(kbEstimateReason) ? "" : "。")
                        + "这次不会发起任何模型请求。")
                    : null,
                kbEstimateDuplicate
                    ? h("div", {
                        style: Object.assign({}, S.note, { color: M.textSecondary }),
                        "data-kb-estimate-duplicate": "true",
                    }, "同一版内容已在库里：实际入库 0 次模型调用，只更新来源档 / 文件夹。")
                    : null,
                h("div", { style: Object.assign({}, S.note, { color: M.textSecondary }) },
                    "字符 " + String(kbEstimate.chars || 0)
                    + " · 单元 " + String(kbEstimate.estimate && kbEstimate.estimate.units || 0)
                    + " · 批次 " + String(kbEstimate.estimate && kbEstimate.estimate.batches || 0)
                    + " · 调用 " + String(kbEstimate.estimate && kbEstimate.estimate.calls || 0)),
                kbEstimate.estimate === null
                    ? null
                    : h("div", { style: Object.assign({}, S.note, { color: M.caption }) },
                        "预计输入 " + String(kbEstimate.estimate.estimatedInputTokens || 0) + " tokens"
                        + " · 输出 " + String(kbEstimate.estimate.estimatedOutputTokens || 0) + " tokens"
                        + " · 上限 " + String(kbEstimate.estimate.maxCalls || 0) + " 次调用"),
                kbEstimate.estimate !== null && kbEstimate.estimate.overBudget === true
                    ? h("div", { style: Object.assign({}, S.note, { color: M.warn }) }, "超出预算：这份文件按当前预算处理不了，服务端不会调用模型。")
                    : null,
                String(kbEstimate.note || "") === "" ? null : h("div", { style: Object.assign({}, S.note, { color: M.caption }) }, String(kbEstimate.note)),
                h("div", { style: { display: "flex", gap: "8px", marginTop: "8px" } },
                    h("button", {
                        type: "button", className: "om-btn", "data-kb-estimate-start": "true",
                        style: Object.assign({}, S.kbMode, kbEstimateUnavailable ? { opacity: 0.55, cursor: "not-allowed" } : {}),
                        disabled: kbEstimateUnavailable,
                        title: kbEstimateUnavailable ? "模型通道不可用：先修好通道再入库" : undefined,
                        onClick: () => { startKbEstimated(); },
                    }, kbEstimateUnavailable ? "通道不可用" : (kbEstimateDuplicate ? "按此入库（0 次模型调用）" : "开始入库")),
                    h("button", {
                        type: "button", className: "om-btn", "data-kb-estimate-cancel": "true",
                        style: S.kbMode,
                        onClick: () => { kbAiPendingRef.current = null; patchState({ kbEstimate: null }); },
                    }, "取消")));

            // 暂停卡：服务端 paused 表示「本轮到预算上限」，manifest 还没写，知识库里看不到；
            // 续跑必须同 name + 完全相同的 data（任务键含正文哈希）。失去内存内容时按钮禁用。
            const kbPaused = state.kbPaused && typeof state.kbPaused === "object" ? state.kbPaused : null;
            const kbPausedPanel = kbPaused === null ? null : h("div", {
                style: Object.assign({}, S.panel, { padding: "12px 14px" }), "data-kb-paused": "true",
            },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "已暂停：", String(kbPaused.name || "")),
                    h("span", { style: S.panelNote }, "批次 " + String(kbPaused.doneBatches || 0) + " / " + String(kbPaused.batches || 0))),
                h("div", { style: Object.assign({}, S.note, { color: M.textSecondary }) },
                    "已调用 " + String(kbPaused.calls || 0) + " 次 · 状态 " + String(kbPaused.status || "paused")),
                String(kbPaused.note || "") === "" ? null : h("div", { style: Object.assign({}, S.note, { color: M.caption }) }, String(kbPaused.note)),
                kbAiPendingRef.current === null
                    ? h("div", { style: Object.assign({}, S.note, { color: M.warn }) }, "这份文件的内容已不在内存里；重新选择同一个文件（内容一致）后可以续跑。")
                    : null,
                h("div", { style: { display: "flex", gap: "8px" } },
                    h("button", {
                        type: "button", className: "om-btn", "data-kb-resume": "true",
                        style: S.kbMode, disabled: kbAiPendingRef.current === null,
                        onClick: () => { kbResumeUpload(); },
                    }, "继续入库"),
                    h("button", {
                        type: "button", className: "om-btn", "data-kb-paused-cancel": "true",
                        style: S.kbMode, disabled: String(state.kbUploading || "") !== "",
                        title: "不再续跑；丢弃内存里的正文与任务键，之后可重新选择文件开始",
                        onClick: () => { cancelKbPausedUpload(); },
                    }, "取消暂停")));

            const kbUploadNote = String(state.kbUploadNote || "") === "" ? null
                : h("div", { style: { fontSize: "11px", color: M.textSecondary, marginBottom: "8px" }, "data-kb-upload-note": "true" }, String(state.kbUploadNote));

            // 流式进度条（第一百一十轮）：阶段来自服务端 NDJSON，模型批次按 done/total 画。
            // 只有真正在跑的上传才有 progress；预估、预览与普通入库不显示。
            const kbProgressData = state.kbProgress !== null && typeof state.kbProgress === "object" ? state.kbProgress : null;
            let kbProgressLabel = "";
            let kbProgressPercent = 6;
            if (kbProgressData !== null) {
                const phase = String(kbProgressData.phase || "");
                const done = Number(kbProgressData.doneBatches) || 0;
                const total = Number(kbProgressData.totalBatches) || 0;
                const calls = Number(kbProgressData.calls) || 0;
                if (phase === "extract") {
                    kbProgressLabel = "正在抽取正文…";
                    kbProgressPercent = 18;
                } else if (phase === "chunks") {
                    kbProgressLabel = "正在切片并规划模型批次…";
                    kbProgressPercent = 26;
                } else if (phase === "plan") {
                    kbProgressLabel = total > 0 ? "模型批次已规划：" + String(total) + " 批" : "正在规划模型批次…";
                    kbProgressPercent = 30;
                } else if (phase === "model") {
                    kbProgressPercent = total > 0 ? 30 + Math.round((Math.min(done, total) / total) * 69) : 40;
                    const detail = ["模型 " + String(done) + "/" + String(total) + " 批", "已调用 " + String(calls) + " 次"];
                    if (Number(kbProgressData.degraded) > 0) detail.push("降级 " + String(kbProgressData.degraded) + " 批");
                    if (Number(kbProgressData.truncated) > 0) detail.push("截断 " + String(kbProgressData.truncated) + " 次");
                    kbProgressLabel = detail.join(" · ");
                } else {
                    kbProgressLabel = "正在上传并处理…";
                    kbProgressPercent = 10;
                }
            }
            const kbProgressBar = kbProgressData === null ? null : h("div", {
                style: { margin: "2px 0 8px" },
                "data-kb-progress": String(kbProgressData.phase || ""),
                role: "progressbar",
                "aria-valuemin": 0,
                "aria-valuemax": 100,
                "aria-valuenow": kbProgressPercent,
                "aria-label": kbProgressLabel,
            },
                h("div", {
                    style: {
                        height: "5px", borderRadius: "999px", overflow: "hidden",
                        background: "color-mix(in srgb, " + M.stroke + " 70%, transparent)",
                    },
                }, h("div", {
                    style: {
                        width: String(kbProgressPercent) + "%", height: "100%",
                        background: M.accent, transition: "width 160ms ease",
                    },
                })),
                h("div", { style: { marginTop: "4px", fontSize: "11px", color: M.textSecondary } }, kbProgressLabel));

            const kbModeButton = (key, label) => {
                const on = state.kbMode === key;
                return h("button", {
                    key: key, type: "button", className: "om-btn",
                    style: on ? Object.assign({}, S.kbMode, S.kbModeOn) : S.kbMode,
                    "data-kb-mode": key, "aria-pressed": on,
                    onClick: () => {
                        if (on) return;
                        load({ kbMode: key, kbDoc: "", kbChunk: "" });
                        scrollPanelTop();
                    },
                }, label);
            };

            // AI 入库的「已读到的文件内容」只留在内存里：预估 → 确认入库 / 暂停后续跑
            // 都要原样重发同一份 base64（服务端按内容 + 模型 profile 认任务，按文件名认不出）。
            const kbReadFileBase64 = async (file) => {
                return await new Promise((settle, fail) => {
                    const reader = new FileReader();
                    reader.onload = () => {
                        const text = String(reader.result || "");
                        const at = text.indexOf("base64,");
                        settle(at >= 0 ? text.slice(at + 7) : "");
                    };
                    reader.onerror = () => fail(new Error("读不了这个文件。"));
                    reader.readAsDataURL(file);
                });
            };
            /**
             * 一次 kb-upload：只认视图层的 { ok, action:'kb-upload', result } 包装，
             * 业务失败（ok:false）也走 200，所以必须按 body.ok 判，不能只看 HTTP。
             */
            async function kbUploadRequest(payload, options) {
                if (typeof fetch !== "function") throw new Error("当前环境没有 fetch，无法上传（这一格只在浏览器里工作）。");
                const onProgress = options && typeof options.onProgress === "function" ? options.onProgress : null;
                const response = await fetch(KB_UPLOAD_PATH + "?cwd=" + encodeURIComponent(cwdFor || ""), {
                    method: "POST",
                    headers: { "content-type": "application/json" },
                    body: JSON.stringify(payload),
                });
                // 流式进度（第一百一十轮）：服务端边跑边写 NDJSON，客户端按行读，
                // 模型批次不再是一段无法解释的「正在处理…」。假 Response / 旧服务端
                // 没有 body 流时，退回下面原来的整段 JSON 读取。
                // 不能只看 body：真实 fetch 的响应都有 body 流，旧服务端回 application/json
                // 时按行读会把整段 JSON 当成一条没有 type 的事件，最后报「没有收到结果」。
                const responseType = response.headers && typeof response.headers.get === "function"
                    ? String(response.headers.get("content-type") || "").toLowerCase()
                    : "";
                if (payload && payload.stream === true && response.ok
                    && responseType.includes("application/x-ndjson")
                    && response.body && typeof response.body.getReader === "function") {
                    const reader = response.body.getReader();
                    const decoder = new TextDecoder();
                    let buffer = "";
                    let result = null;
                    let failure = "";
                    const consume = (line) => {
                        const text = String(line || "").trim();
                        if (text === "") return;
                        let event = null;
                        try { event = JSON.parse(text); } catch { return; }
                        if (event && event.type === "progress") {
                            if (onProgress !== null) {
                                try { onProgress(event); } catch { /* 画进度不能打断上传 */ }
                            }
                        } else if (event && event.type === "result") {
                            result = event.result && typeof event.result === "object" ? event.result : {};
                        } else if (event && event.type === "error") {
                            failure = String(event.error || "上传失败，端点没有给出原因");
                        }
                    };
                    for (;;) {
                        const step = await reader.read();
                        if (step.done) break;
                        buffer += decoder.decode(step.value, { stream: true });
                        let at = buffer.indexOf("\n");
                        while (at >= 0) {
                            consume(buffer.slice(0, at));
                            buffer = buffer.slice(at + 1);
                            at = buffer.indexOf("\n");
                        }
                    }
                    buffer += decoder.decode();
                    consume(buffer);
                    if (failure !== "") throw new Error(failure);
                    if (result === null) throw new Error("上传流结束了，但没有收到结果；请重试一次。");
                    return result;
                }
                // 与 OCR 同一条 read-text-first 纪律：413 空 body / 网关 HTML / 非 JSON 时
                // response.json() 只会抛浏览器原生 SyntaxError，HTTP 状态全丢；这里保状态。
                const parsed = await readKbUploadResponse(response);
                if (parsed.body === null) throw new Error(parsed.problem);
                const body = parsed.body;
                if (!response.ok || body.ok !== true) {
                    const detail = String(body.error || "").trim() || "上传失败，端点没有给出原因";
                    const code = String(body.code || "").trim();
                    // 失败回执保留服务端文案，并附 HTTP 状态与机器可读 code（如 body_too_large）。
                    throw new Error(detail + "（HTTP " + parsed.status + (code !== "" ? "；code " + code : "") + "）");
                }
                return body.result && typeof body.result === "object" ? body.result : {};
            }
            /**
             * 预估：只做纯计算，不落库、不调模型；用户看过调用数 / 预算后再决定是否入库。
             *
             * `content` 是「送什么正文」的那个对象：二进制通道是 `{ data }`，OCR 正文通道是
             * `{ text, ocr: true }`。两条通道共用同一条预估 → 确认 → 续跑流水线，所以
             * pending 里必须原样保留 content 的字段（`startKbEstimated` 只加 ai 开关）。
             */
            async function estimateKbPayload(name, content, folder) {
                // 暂停期间不允许从预览卡再起一轮预估：那会绕开续跑，变成从 0 开始的付费任务。
                // 要重新开始，必须先点暂停卡上的「取消暂停」（显式放弃这份任务）。
                const pausedNow = latest.current.kbPaused;
                if (pausedNow !== null && typeof pausedNow === "object") {
                    patchState({ actionError: "已有暂停的入库任务：请先「继续入库」或「取消暂停」；暂停期间不会发起新的模型请求。" });
                    return;
                }
                patchState({ kbUploading: name, actionError: "", kbUploadNote: "", kbEstimate: null, kbPaused: null });
                try {
                    const body = content !== null && typeof content === "object" ? content : {};
                    const hasBody = (typeof body.data === "string" && body.data !== "")
                        || (typeof body.text === "string" && body.text !== "");
                    if (!hasBody) throw new Error(name + " 的内容是空的。");
                    // 目标文件夹必须是**预览时捕获的那一份**：预估与随后真正的入库都用
                    // pending.folder。若这里改读 state 的当前目录，用户中途换层就会变成
                    // 「预估说 A、实际写 B」。
                    const targetFolder = typeof folder === "string" ? folder : String(kbFolderNow || "");
                    // 模型也必须在预估时定死：用户在看过 0 次调用的 duplicate 预估后换模型，
                    // 不能让真实入库 / 续跑改用另一个（可能计费的）模型或另一个密钥。
                    const targetModel = String(state.kbAiModel || "");
                    // pending 必须记住内容、目标文件夹与模型：确认入库时原样重发，服务端按内容认任务。
                    // `data` 与 `text`+`ocr` 都是正文载体，谁在 content 里就带谁。
                    kbAiPendingRef.current = Object.assign({ name: name, folder: targetFolder }, body, { aiModel: targetModel });
                    const result = await kbUploadRequest(Object.assign({}, body, {
                        name: name, folder: targetFolder,
                        ai: true, aiModel: targetModel,
                        aiEstimate: true, aiResume: false,
                    }));
                    patchState({
                        kbUploading: "",
                        kbEstimate: {
                            name: name, chars: Number(result.chars) || 0, folder: targetFolder,
                            aiModel: targetModel,
                            estimate: result.estimate && typeof result.estimate === "object" ? result.estimate : null,
                            note: String(result.note || ""),
                            ai: result.ai && typeof result.ai === "object" ? result.ai : null,
                            // duplicate：同内容 + 同切块档已在库里，实际入库 0 次模型调用（仍可
                            // 搬文件夹 / 换来源档），不能因此禁用「开始入库」。
                            duplicate: result.duplicate === true,
                        },
                    });
                } catch (error) {
                    patchState({ kbUploading: "", kbProgress: null, actionError: "预估失败：" + String(error && error.message ? error.message : error) });
                }
            }
            async function startKbEstimated() {
                // 按钮在 available === false 时已禁用；这里再兜一层，程序化触发也不发。
                const shown = latest.current.kbEstimate;
                if (shown && shown.ai && typeof shown.ai === "object" && shown.ai.available === false) return;
                // 暂停后只允许续跑（aiResume:true）或显式取消；旧按钮回调、
                // 连点两次都不能再发一次 aiResume:false 的付费请求。
                if (String(latest.current.kbUploading || "") !== "") return;
                const shownPaused = latest.current.kbPaused;
                if (shownPaused !== null && typeof shownPaused === "object") return;
                const pending = kbAiPendingRef.current;
                if (pending === null || pending === undefined) {
                    patchState({ actionError: "这份文件的内容已不在内存里，请重新选择同一个文件。" });
                    return;
                }
                patchState({ kbUploading: pending.name, actionError: "", kbUploadNote: "", kbProgress: { phase: "upload" } });
                const onProgress = (event) => {
                    patchState({ kbProgress: {
                        phase: String(event.phase || "model"),
                        doneBatches: Number(event.doneBatches) || 0,
                        totalBatches: Number(event.totalBatches) || 0,
                        calls: Number(event.calls) || 0,
                        degraded: Number(event.degraded) || 0,
                        truncated: Number(event.truncated) || 0,
                    } });
                };
                try {
                    // 不再读 state.kbAiModel：模型已随 pending 捕获，换下拉不改变本次入库。
                    const result = await kbUploadRequest(Object.assign({}, pending, {
                        ai: true, aiEstimate: false, aiResume: false, stream: true,
                    }), { onProgress: onProgress });
                    await kbHandleIngestResult(pending, result);
                } catch (error) {
                    patchState({ kbUploading: "", kbProgress: null, actionError: "入库失败：" + String(error && error.message ? error.message : error) });
                }
            }
            async function kbResumeUpload() {
                // 只有仍然停在暂停卡上才续跑；连点或任务已结束后旧回调都不能再发请求。
                const pausedNow = latest.current.kbPaused;
                if (pausedNow === null || typeof pausedNow !== "object") return;
                if (String(latest.current.kbUploading || "") !== "") return;
                const pending = kbAiPendingRef.current;
                if (pending === null || pending === undefined) {
                    patchState({ actionError: "续跑需要同一份文件内容；刷新后请重新选择同一个文件（内容一致才能接上）。" });
                    return;
                }
                patchState({ kbUploading: pending.name, actionError: "", kbUploadNote: "", kbProgress: { phase: "upload" } });
                const onProgress = (event) => {
                    patchState({ kbProgress: {
                        phase: String(event.phase || "model"),
                        doneBatches: Number(event.doneBatches) || 0,
                        totalBatches: Number(event.totalBatches) || 0,
                        calls: Number(event.calls) || 0,
                        degraded: Number(event.degraded) || 0,
                        truncated: Number(event.truncated) || 0,
                    } });
                };
                try {
                    // 续跑必须和预估同一个模型，否则可能接到另一个密钥 / 计费通道。
                    const result = await kbUploadRequest(Object.assign({}, pending, {
                        ai: true, aiEstimate: false, aiResume: true, stream: true,
                    }), { onProgress: onProgress });
                    await kbHandleIngestResult(pending, result);
                } catch (error) {
                    patchState({ kbUploading: "", kbProgress: null, actionError: "续跑失败：" + String(error && error.message ? error.message : error) });
                }
            }
            /** 显式放弃暂停任务：丢掉内存里的正文与任务键，允许重新选择文件开始。 */
            function cancelKbPausedUpload() {
                if (String(latest.current.kbUploading || "") !== "") return;
                kbAiPendingRef.current = null;
                patchState({
                    kbPaused: null, kbEstimate: null, actionError: "",
                    kbUploadNote: "已取消暂停的入库任务：没有再调用模型；重新选择文件可以重新开始。",
                });
            }
            /** 目标文件夹的可读写法：空路径就是根目录「全部文档」。 */
            function kbFolderLabel(folder) {
                const path = String(folder === undefined || folder === null ? "" : folder);
                return path === "" ? "全部文档" : path;
            }
            /**
             * 入库回执：只按服务端返回的**结构化字段**说话。
             *
             * - 普通入库与 AI 入库必须一眼区分；AI 回执写清调用、降级、摘要与错误条数。
             * - 错误只报条数：服务端的 ai.errors 里可能夹带模型端点返回的原文片段，不能
             *   被原样复制到面板上。
             * - unchanged 时不能说成重写了一遍；服务端给了说明（note）就附上。
             */
            function kbIngestReceipt(name, result, fallbackFolder) {
                const row = result && typeof result === "object" ? result : {};
                const folder = typeof row.folder === "string"
                    ? row.folder
                    : String(fallbackFolder === undefined || fallbackFolder === null ? "" : fallbackFolder);
                const ai = row.ai && typeof row.ai === "object" ? row.ai : null;
                const parts = ["已入库：" + String(name || "")];
                const chunks = Number(row.chunks) || 0;
                const chars = Number(row.chars) || 0;
                if (row.unchanged === true) {
                    if (row.moved === true) parts.push("内容没变，只把文件夹改到「" + kbFolderLabel(folder) + "」");
                    else if (row.retiered === true) parts.push("内容没变，只改了来源档（" + String(row.tier || "—") + "）");
                    else parts.push("内容没变：库里的版本原样保留，没有重写");
                } else {
                    parts.push("进入「" + kbFolderLabel(folder) + "」"
                        + (chunks > 0 ? " · " + String(chunks) + " 块" : "")
                        + (chars > 0 ? " · " + String(chars) + " 字符" : ""));
                }
                // OCR 正文与普通正文必须一眼分开：它是识别结果，不是源文件的文字层。
                if (row.ocr === true) parts.push("正文来自 OCR 识别结果（已替换这份文件之前的识别内容）");
                if (typeof row.removedOcrNote === "string" && row.removedOcrNote !== "") {
                    parts.push("早先的笔记 " + row.removedOcrNote + " 已一并收掉");
                }
                if (ai === null || ai.mode === "off") {
                    parts.push("确定性切块：未调用模型、没有模型摘要");
                } else {
                    const calls = Number(ai.calls) || 0;
                    const degraded = Number(ai.degraded) || 0;
                    const summaries = ai.summaries === undefined || ai.summaries === null ? null : Number(ai.summaries) || 0;
                    // 只取条数：errors 的文本可能含模型端点的原始响应。
                    const errors = Array.isArray(ai.errors) ? ai.errors.length : 0;
                    const detail = [
                        "AI 建库",
                        "模型调用 " + String(calls) + " 次",
                        ai.chunking === true ? "模型切片" : "未重切正文",
                        ai.summary === true
                            ? (summaries === null ? "本次未生成摘要" : "摘要 " + String(summaries) + " 条")
                            : "模型摘要未开",
                    ];
                    if (degraded > 0) detail.push(String(degraded) + " 批降级为确定性切块");
                    if (errors > 0) detail.push("错误 " + String(errors) + " 条");
                    parts.push(detail.join(" · "));
                }
                const note = String((ai && ai.note) || row.note || "").trim();
                if (note !== "") parts.push("说明：" + note);
                return parts.join(" · ");
            }
            /** paused / 完成两种回执：paused 留着续跑按钮，完成才清掉临时内容。 */
            async function kbHandleIngestResult(pending, result) {
                if (result && result.paused === true) {
                    // 暂停后必须撤掉预估卡：留着 enabled 的「开始入库」会以 aiResume:false
                    // 从游标 0 再跑一遍，产生一次重复的付费调用。
                    patchState({
                        kbUploading: "",
                        kbProgress: null,
                        kbEstimate: null,
                        kbPaused: {
                            name: pending.name,
                            status: String(result.status || "paused"),
                            calls: Number(result.calls) || 0,
                            doneBatches: Number(result.doneBatches) || 0,
                            batches: Number(result.batches) || 0,
                            note: String(result.note || ""),
                            stateKey: String(result.stateKey || ""),
                        },
                    });
                    return;
                }
                kbAiPendingRef.current = null;
                // AI 入库（预估确认 / 续跑 / OCR 勾模型）走的是这条收尾：预览卡必须
                // 跟着一起撤掉。否则库里已经有了这份文档，界面上还挂着一张「待入库」，
                // 再点一次「入库」就是重复写入 —— 非 AI 路径在 ingestPreview 里按 key
                // 删卡，两条路的口径要一致。
                const pendingName = String(pending && pending.name ? pending.name : "");
                const allPreviews = Array.isArray(latest.current.kbPreviews) ? latest.current.kbPreviews : [];
                const droppedKeys = allPreviews
                    .filter((one) => String(one.name) === pendingName)
                    .map((one) => one.key);
                for (const key of droppedKeys) delete kbPreviewDataRef.current[key];
                patchState({
                    kbUploading: "", kbProgress: null, kbEstimate: null, kbPaused: null,
                    kbPreviews: allPreviews.filter((one) => String(one.name) !== pendingName),
                    kbUploadNote: kbIngestReceipt(pending.name, result, pending.folder),
                });
                load({ kbMode: "docs", kbDoc: "", kbChunk: "" });
            }

            // 函数声明而不是 const：按钮在 kbCrumb 里（更靠前），const 会在渲染到这里之前
            // 就因 TDZ 抛「Cannot access before initialization」。
            //
            // 统一入口的第一步：**只预览**。服务端 kbPreview 只做抽取与切片计数，
            // 不写 manifest、不建块、不调模型；正文留在内存 ref 里，等用户显式入库。
            async function previewFiles(fileList) {
                const files = Array.from(fileList || []);
                if (files.length === 0) return;
                kbPreviewSeq.current += 1;
                const seqNow = kbPreviewSeq.current;
                // 目标文件夹在选文件的这一刻捕获：卡片、预览请求与随后的入库都用它，
                // 不跟左侧目录树后来怎么导航走（否则会悄悄入到别的文件夹）。
                const targetFolder = String(kbFolderNow || "");
                // 新一批选择替换上一批：预览卡是「这次要入库什么」，不是历史记录。
                // 但暂停任务的内存正文必须保留，否则点「继续入库」就再也接不上；暂停期间
                // 新预览卡的 AI 入库也会被 estimateKbPayload 拦住，只能先显式取消。
                const pausedBeforePreview = latest.current.kbPaused;
                if (pausedBeforePreview === null || typeof pausedBeforePreview !== "object") kbAiPendingRef.current = null;
                const kept = [];
                for (const file of files) {
                    const name = String(file && file.name || "文件");
                    const size = Number(file && file.size) || 0;
                    const ext = kbExtOf(name);
                    const ocrType = kbOcrEligibleName(name);
                    const extractable = KB_EXTRACTABLE_EXTS.indexOf(ext) >= 0;
                    const previewPatch = { kbUploading: name, actionError: "", kbUploadNote: "", kbEstimate: null };
                    // 暂停任务不能被一次新预览顺手清掉；只有非暂停时选新文件才重置暂停态。
                    if (pausedBeforePreview === null || typeof pausedBeforePreview !== "object") previewPatch.kbPaused = null;
                    patchState(previewPatch);
                    if (size > KB_OCR_CONTENT_MAX) {
                        // 连 OCR 内容通道也不收：不读、不发，直接说清上限。
                        patchState({ actionError: "“" + name + "”有 " + String(size) + " 字节，超过单文件上限 24 MB（识别通道也是这个上限）：先压小或拆页再试。" });
                        continue;
                    }
                    const key = String(name) + "#" + String(kept.length) + "#" + String(Date.now());
                    const card = (fields) => {
                        // 本地判定出来的卡也要能显式 OCR：正文在用户点按钮时才读，
                        // 预览阶段绝不把内容发给远端。
                        kbPreviewDataRef.current[key] = { file: file };
                        kept.push(Object.assign({
                            key: key, name: name, ext: ext, chars: 0, chunks: 0, bytes: size,
                            images: 0, head: "", note: "", truncated: false, folder: targetFolder,
                        }, fields));
                    };
                    if (!extractable) {
                        card({
                            kind: "unsupported",
                            ocrEligible: ocrType,
                            note: ocrType
                                ? "这个格式不能直接入库：支持文本、PDF、Word / PPT / Excel 与常见图片。要读里面的文字，可以点「识别文字」，结果能存成知识库笔记。"
                                : "这个格式不能直接入库：支持文本、PDF、Word / PPT / Excel 与常见图片。可以先另存为其中的一种。",
                        });
                        continue;
                    }
                    if (KB_BINARY_EXTS.indexOf(ext) < 0 && size > KB_MAX_TEXT_BYTES) {
                        // 与 kbIngest 的文本口径一致：2 MB 以上不抽、不显示「入库」。
                        card({
                            kind: "too_large",
                            ocrEligible: false,
                            note: "文本文件有 " + String(size) + " 字节，超过单文档上限 2 MB：先切小或只放需要的那一节。",
                        });
                        continue;
                    }
                    // 大 PDF 是唯一例外：服务端会按页拆成多份（8–24 MB），拆完每份都
                    // 落在 8 MB 抽取上限内，所以这里不能拿 8 MB 直接挡掉。其余格式仍
                    // 沿用「超过 8 MB 不预览」的诚实边界。
                    const bigPdf = ext === "pdf" && size <= KB_OCR_CONTENT_MAX;
                    if (size > KB_PREVIEW_LOCAL_MAX && !bigPdf) {
                        // 诚实边界：没抽取就不能断言它是扫描件，只能说「没判定文本层」。
                        card(ocrType ? {
                            kind: "unknown",
                            ocrEligible: true,
                            note: "文件超过本地预览上限（8 MB），没有判定文本层；要读里面的文字请点「识别文字」，内容会发给 DeepSeek OCR。",
                        } : {
                            kind: "too_large",
                            ocrEligible: false,
                            note: "文件超过可抽取的单文件上限（8 MB）：先拆分或只放需要的那一节。",
                        });
                        continue;
                    }
                    try {
                        const encoded = await kbReadFileBase64(file);
                        if (encoded === "") throw new Error(name + " 的内容是空的。");
                        const result = await kbUploadRequest({
                            name: name, folder: targetFolder, data: encoded, preview: true,
                        });
                        if (seqNow !== kbPreviewSeq.current) return;
                        const splitParts = result && result.kind === "split" && Array.isArray(result.parts)
                            ? result.parts
                            : null;
                        if (splitParts !== null && splitParts.length > 0) {
                            // 大 PDF 自动拆分：每份一张卡，各自带回可直接入库的 base64。
                            // 用户逐份确认，避免一次点下就把整本 24 MB 的扫描件发出去。
                            splitParts.forEach((part, partIndex) => {
                                const partName = String(part && part.name || (name + "（第" + (partIndex + 1) + "份）"));
                                const partKey = partName + "#" + String(kept.length) + "#" + String(Date.now()) + "#" + String(partIndex);
                                if (part && typeof part.data === "string" && part.data !== "") {
                                    kbPreviewDataRef.current[partKey] = part.data;
                                }
                                kept.push({
                                    key: partKey,
                                    name: partName,
                                    ext: "pdf",
                                    kind: String(part.kind || "unknown"),
                                    ocrEligible: part.ocrEligible === true,
                                    chars: Number(part.chars) || 0,
                                    chunks: Number(part.chunks) || 0,
                                    bytes: Number(part.bytes) || 0,
                                    images: Number(part.images) || 0,
                                    head: String(part.head || ""),
                                    note: String(part.note || ""),
                                    truncated: part.truncated === true,
                                    pageFrom: Number(part.pageFrom) || 0,
                                    pageTo: Number(part.pageTo) || 0,
                                    folder: targetFolder,
                                });
                            });
                        } else {
                            kbPreviewDataRef.current[key] = encoded;
                            kept.push({
                                key: key,
                                name: String(result.name || name),
                                ext: String(result.ext || ""),
                                kind: String(result.kind || "text"),
                                ocrEligible: result.ocrEligible === true,
                                chars: Number(result.chars) || 0,
                                chunks: Number(result.chunks) || 0,
                                bytes: Number(result.bytes) || 0,
                                images: Number(result.images) || 0,
                                head: String(result.head || ""),
                                note: String(result.note || ""),
                                truncated: result.truncated === true,
                                folder: targetFolder,
                            });
                        }
                    } catch (error) {
                        patchState({ actionError: "预览失败：" + String(error && error.message ? error.message : error) });
                    }
                }
                if (seqNow !== kbPreviewSeq.current) return;
                // 之前批次（含本次选择里被丢掉的卡）的正文都不再需要：只按「还在卡上的
                // key」保留，其余全删。只看 latest 会漏掉被并发覆盖的那一批 ref。
                const keepKeys = new Set(kept.map((one) => one.key));
                for (const key of Object.keys(kbPreviewDataRef.current)) {
                    if (!keepKeys.has(key)) delete kbPreviewDataRef.current[key];
                }
                patchState({ kbUploading: "", kbPreviews: kept });
            }
            /** 预览卡上的「入库」：未勾 AI 走确定性切块（零模型调用）；勾了就走预估 → 确认。 */
            async function ingestPreview(item) {
                const stored = kbPreviewDataRef.current[item.key];
                // 卡上的正文有两种载体：预览读到的原始文件（binary，base64 字符串），
                // 或识别完成后的现成正文（{ text, ocr:true }）。两条都从这里进同一个入口。
                const content = typeof stored === "string" && stored !== ""
                    ? { data: stored }
                    : (stored !== null && typeof stored === "object"
                        && typeof stored.text === "string" && stored.text !== ""
                        ? { text: stored.text, ocr: true }
                        : null);
                if (content === null) {
                    patchState({ actionError: "这份文件的内容已不在内存里，请重新选择同一个文件。" });
                    return;
                }
                if (state.kbAi === true && (item.kind === "text" || item.kind === "ocr")) {
                    await estimateKbPayload(item.name, content, item.folder);
                    return;
                }
                patchState({ kbUploading: item.name, actionError: "", kbUploadNote: "" });
                try {
                    const result = await kbUploadRequest(Object.assign({ name: item.name, folder: item.folder, ai: false, aiModel: "" }, content));
                    delete kbPreviewDataRef.current[item.key];
                    patchState({
                        kbUploading: "",
                        kbPreviews: (Array.isArray(latest.current.kbPreviews) ? latest.current.kbPreviews : []).filter((one) => one.key !== item.key),
                        kbUploadNote: kbIngestReceipt(item.name, result, item.folder)
                            + "；如需识别图片或扫描件里的文字，重新添加后再点「识别文字」。",
                    });
                    load({ kbMode: "docs", kbDoc: "", kbChunk: "", kbOffset: 0 });
                } catch (error) {
                    patchState({ kbUploading: "", actionError: "入库失败：" + String(error && error.message ? error.message : error) });
                }
            }
            function pickFiles() {
                const input = document.querySelector("[data-kb-upload-input]");
                if (input && typeof input.click === "function") input.click();
            }
            // UTF-8 文本 → base64：OCR 结果要复用 kb-upload 那条入库通道（JSON + base64）。
            function textToBase64(text) {
                const bytes = new TextEncoder().encode(String(text || ""));
                let binary = "";
                const CHUNK = 0x8000;
                for (let at = 0; at < bytes.length; at += CHUNK) {
                    binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(at, at + CHUNK)));
                }
                return btoa(binary);
            }

            /**
             * OCR（第九十五轮）：只在预览卡上显式点「识别文字」时才发。
             *
             * 发的是**内容**（data + name），不是客户端路径：服务端把内容写成自己
             * 临时目录里的文件，识别完即删，响应也不回路径（信任边界只在服务端）。
             * 预览本身不触发这一条 —— 扫描件在入库前不会自动发给任何远端。
             */
            /**
             * 读一个 JSON 端点的响应：**先读文本再自己解析**，不打无准备的 response.json()。
             *
             * 直接 json() 有两个坑，而且都是「空响应」：宿主没命中路由时是 404 空 body、
             * 兜底是 400 空 body，浏览器就原样抛 `Failed to execute 'json' on 'Response':
             * Unexpected end of JSON input`；反向代理/网关插进来的 HTML 错误页同理。
             * 两种都会把 HTTP 状态一起丢掉，页面上只剩浏览器的原生报错。这里自己判定
             * 「空 / HTML / 非 JSON」并带上状态与 content-type；**不回响应体原文** ——
             * 响应体可能带机器上的路径或内部页面，不该被复制到面板上。
             *
             * endpoint 只影响文案里的端点名：OCR 与知识库上传共用这套判定，但报错时必须
             * 说清是哪个端点（把 KB 的失败叫成 OCR 会把人带到错的方向）。
             */
            async function readJsonEndpointResponse(response, endpoint) {
                const status = Number(response && response.status) || 0;
                const contentType = String(
                    response && response.headers && typeof response.headers.get === "function"
                        ? response.headers.get("content-type") || ""
                        : "",
                );
                let raw = "";
                try {
                    raw = await response.text();
                } catch (error) {
                    return {
                        status, contentType, body: null,
                        problem: endpoint + "响应读到一半就断开了（HTTP " + status + "）："
                            + String(error && error.message ? error.message : error),
                    };
                }
                const trimmed = String(raw || "").trim();
                if (trimmed === "") {
                    // 只陈述观测到的事实：哪个端点、什么状态、body 是空的。404 时补一句
                    // 能照做的核对步骤（路由这一层）；「原因」没有证据，不写推测。
                    const pista = status === 404
                        ? "插件里没有这个路由：确认办公模式插件已启用，并重载页面后再试。"
                        : "";
                    return {
                        status, contentType, body: null,
                        problem: endpoint + "返回了空响应（HTTP " + status
                            + "，content-type: " + (contentType || "无") + "）。" + pista,
                    };
                }
                let body = null;
                try { body = JSON.parse(trimmed); } catch { body = null; }
                if (body === null || typeof body !== "object" || Array.isArray(body)) {
                    const looksHtml = /html|xml/i.test(contentType) || /^</.test(trimmed);
                    const kind = looksHtml ? "HTML/XML 页面" : "非 JSON 文本";
                    // 同样只给状态与分类：这份 body 可能是网关错误页或内网报文，
                    // 既不该复制进面板，也不该替它猜原因（上游/代理/本地路由都没有证据）。
                    return {
                        status, contentType, body: null,
                        problem: endpoint + "返回的是 " + kind + "，不是约定的 JSON（HTTP " + status
                            + "，content-type: " + (contentType || "无") + "）。",
                    };
                }
                return { status, contentType, body, problem: "" };
            }
            /** OCR 端点（/office-memory/ocr）的响应。 */
            async function readOcrResponse(response) {
                return await readJsonEndpointResponse(response, "OCR 端点");
            }
            /** 知识库上传端点（/office-memory/kb-upload）的响应。 */
            async function readKbUploadResponse(response) {
                return await readJsonEndpointResponse(response, "知识库上传端点");
            }

            async function requestOcr(payload, displayName, folder, itemKey) {
                const name = String(displayName || "文件");
                // 识别本身不落库；文件夹只是给「存为笔记」用的：沿用预览那一刻的目标，
                // 不跟用户之后在目录树里的导航走。
                const targetFolder = typeof folder === "string" ? folder : String(kbFolderNow || "");
                patchState({ ocrBrowsing: false, ocr: { name: name, path: "", loading: true, error: "", text: "", pageCount: 0, truncatedPages: [], folder: targetFolder } });
                try {
                    const response = await fetch(OCR_PATH + "?cwd=" + encodeURIComponent(cwdFor || ""), {
                        method: "POST",
                        headers: { "content-type": "application/json" },
                        body: JSON.stringify(payload),
                    });
                    const parsed = await readOcrResponse(response);
                    if (parsed.body === null) throw new Error(parsed.problem);
                    const body = parsed.body;
                    if (!response.ok || body.ok !== true) {
                        const detail = String(body.error || "").trim();
                        const code = String(body.code || "").trim();
                        const hint = String(body.hint || "").trim();
                        // 状态码不吞：非 2xx 时无论端点写没写原因都附上 HTTP 状态；
                        // code 是服务端/上游的机器可读分类（如 gateway_timeout_error）。
                        const notes = [];
                        if (!response.ok) notes.push("HTTP " + parsed.status);
                        if (code !== "") notes.push("code " + code);
                        if (hint !== "") notes.push(hint);
                        throw new Error((detail || "OCR 失败，端点没有给出原因") + (notes.length > 0 ? "（" + notes.join("；") + "）" : ""));
                    }
                    const recognized = String(body.text || "");
                    patchState({
                        ocr: {
                            name: String(body.fileName || name), path: "", loading: false, error: "",
                            text: recognized, pageCount: Number(body.pageCount) || 0,
                            truncatedPages: Array.isArray(body.truncatedPages) ? body.truncatedPages : [],
                            ms: Number(body.ms) || 0, usage: body.usage || null, folder: targetFolder,
                        },
                    });
                    // 识别结果**替换待入库卡上的正文**（第一百零一轮补）：PDF 的文本层可能
                    // 很脏（版式、页眉、乱序），上面那张卡原来显示的就是它。识别完成后把
                    // 卡上的正文换成识别文本，用户点「入库」写进去的才是准确的那一份，
                    // 不需要再单独点「存为知识库文档」—— 两条路径写同一个文档，重复保存
                    // 只会多一次换代。
                    if (typeof itemKey === "string" && itemKey !== "" && recognized.trim() !== "") {
                        const encodedText = textToBase64(recognized);
                        kbPreviewDataRef.current[itemKey] = { text: encodedText, ocr: true };
                        const list = Array.isArray(latest.current.kbPreviews) ? latest.current.kbPreviews : [];
                        patchState({
                            kbPreviews: list.map((one) => (one.key === itemKey ? Object.assign({}, one, {
                                kind: "ocr",
                                chars: recognized.length,
                                chunks: 0,
                                images: 0,
                                head: recognized.slice(0, 1200),
                                truncated: recognized.length > 1200,
                                note: "上面已换成识别结果；入库会写入这份文件自己的文档，并收掉旧的 .ocr.md 笔记。",
                            }) : one)),
                        });
                    }
                } catch (error) {
                    patchState({ ocr: {
                        name: name, path: "", loading: false, text: "", pageCount: 0, truncatedPages: [], folder: targetFolder,
                        error: "识别失败：" + String(error && error.message ? error.message : error),
                    } });
                }
            }
            /**
             * 预览卡上的显式 OCR：只在用户点按钮时读内容，不发客户端路径。
             *
             * 8–24 MB 的文件在预览阶段**没有读过**（只登记了 File），所以这里才
             * FileReader —— 这正是「不点就不发」的一部分：预览不读大文件、不发远端。
             */
            async function runOcrPreview(item) {
                const stored = kbPreviewDataRef.current[item.key];
                let data = typeof stored === "string" ? stored : "";
                if (data === "" && stored && stored.file) {
                    try {
                        data = await kbReadFileBase64(stored.file);
                    } catch (error) {
                        patchState({ actionError: "读这份文件失败：" + String(error && error.message ? error.message : error) });
                        return;
                    }
                }
                if (typeof data !== "string" || data === "") {
                    patchState({ actionError: "这份文件的内容已不在内存里，请重新选择同一个文件再识别。" });
                    return;
                }
                await requestOcr({ name: item.name, data: data }, item.name, item.folder, item.key);
            }

            /**
             * 把 OCR 正文存进知识库（第一百零一轮）。
             *
             * 落点是**源文件自己的那一份文档**：`uploads/<文件夹>/<原文件名>`。旧版本把正文
             * 旁挂成 `<stem>.ocr.md`，结果同一份扫描件在库里有两条互相矛盾的正文；现在
             * 再次识别同一个文件 = 同一个 docId 换代，服务端会顺手收掉那份旧笔记。
             *
             * 勾了「用模型切片 / 摘要」时走与普通文件相同的预估 → 确认 → 续跑流水线：
             * 切块与摘要都发生在 OCR 正文上，且只有用户确认过调用数才会真的调模型。
             */
            function saveOcrAsNote() {
                const draft = state.ocr;
                if (draft === null || draft === undefined) return;
                const text = String(draft.text || "").trim();
                if (text === "") { patchState({ actionError: "识别结果是空的，没有可保存的正文。" }); return; }
                // 名字必须是**原文件名**：显示路径由它决定，改名就等于另存一份新文档。
                const name = String(draft.name || "").trim() || ("ocr-" + String(Date.now()));
                // 预览时捕获的文件夹优先，老状态里没有 folder 才退回当前目录。
                const folder = typeof draft.folder === "string" ? draft.folder : String(kbFolderNow || "");
                const content = { text: textToBase64(text), ocr: true };
                patchState({ ocrSaving: true, actionError: "" });
                if (state.kbAi === true) {
                    // 面板收起：正文已经转交预估卡，留着可编辑的文本框会让「再存一次」变成
                    // 第二次付费入库。预估失败时 actionError 会写明原因，正文不丢在状态机里。
                    patchState({ ocr: null, ocrSaving: false });
                    estimateKbPayload(name, content, folder);
                    return;
                }
                return kbUploadRequest(Object.assign({ name: name, folder: folder, ai: false, aiModel: "" }, content))
                    .then((result) => {
                        patchState({
                            ocrSaving: false, ocr: null,
                            kbUploadNote: kbIngestReceipt(name, result, folder),
                        });
                        load({ kbMode: "docs", kbDoc: "", kbChunk: "" });
                    })
                    .catch((error) => {
                        patchState({ ocrSaving: false, actionError: "存进知识库失败：" + String(error && error.message ? error.message : error) });
                    });
            }

            const ocrPanel = state.ocr === null || state.ocr === undefined ? null : h("div", {
                style: Object.assign({}, S.editor, { marginBottom: "10px" }), "data-kb-ocr-panel": "true",
            },
                h("div", { style: S.editorTitle }, "识别结果 · " + String(state.ocr.name || "")
                    + (state.ocr.loading
                        ? "（正在识别…）"
                        : state.ocr.error
                            // 失败时没有页数与耗时：写「0 页 / 0 ms」等于替失败编了一个结果。
                            ? "（识别失败）"
                            : "（" + String(state.ocr.pageCount || 0) + " 页 / " + String(state.ocr.ms || 0) + " ms）")),
                state.ocr.error ? h("div", { style: S.editorError, "data-kb-ocr-error": "true" }, String(state.ocr.error)) : null,
                !state.ocr.loading && !state.ocr.error && Array.isArray(state.ocr.truncatedPages) && state.ocr.truncatedPages.length > 0
                    ? h("div", {
                        style: Object.assign({}, S.note, { color: M.warn }),
                        "data-kb-ocr-truncated": String(state.ocr.truncatedPages.length),
                    }, "有 " + String(state.ocr.truncatedPages.length) + " 页没有正常结束（"
                        + state.ocr.truncatedPages.map((page) => "#" + String(page && page.index) + " " + String(page && page.finishReason)).join("；")
                        + "），这些页的正文可能不完整；保存前请先核对。")
                    : null,
                state.ocr.loading ? h("div", { style: dim(T.note) }, "正在上传给 DeepSeek OCR 并等待结果（官方总时长上限 180 秒）…") : null,
                state.ocr.loading || state.ocr.error ? null : h("textarea", {
                    style: Object.assign({}, S.editorTextarea, { minHeight: "220px" }),
                    "data-kb-ocr-text": "true",
                    value: String(state.ocr.text || ""),
                    onChange: (event) => patchState({ ocr: Object.assign({}, latest.current.ocr || {}, { text: String(event.target.value || "") }) }),
                }),
                state.ocr.loading || state.ocr.error ? null : h("div", {
                    style: Object.assign({}, S.note, { color: M.caption }),
                    "data-kb-ocr-target": String(state.ocr.folder || ""),
                }, "保存会写入「" + (String(state.ocr.folder || "") === "" ? "全部文档" : String(state.ocr.folder))
                    + "」里这份文件自己的文档（" + String(state.ocr.name || "") + "），替换它之前的识别结果；"
                    + "早先存成 .ocr.md 的旧笔记会一并收掉。"),
                state.ocr.loading || state.ocr.error ? null : h("label", {
                    style: Object.assign({}, S.note, {
                        display: "flex", gap: "6px", alignItems: "center", color: M.caption,
                    }),
                    "data-kb-ocr-ai": "true",
                },
                    h("input", {
                        type: "checkbox",
                        checked: state.kbAi === true,
                        onChange: (event) => patchState({ kbAi: event.target.checked === true }),
                    }),
                    "用模型切片 / 摘要（先预估调用数，确认后才调用）"),
                h("div", { style: S.editorActions },
                    state.ocr.loading || state.ocr.error ? null : h("button", {
                        type: "button", className: "om-btn om-primary", style: S.memPrimary,
                        "data-kb-ocr-save": "true", disabled: state.ocrSaving === true,
                        onClick: saveOcrAsNote,
                    }, state.ocrSaving === true ? "保存中…" : "存为知识库文档"),
                    state.ocr.loading || state.ocr.error ? null : h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-kb-ocr-copy": "true",
                        onClick: () => copyText(String(state.ocr.text || "")),
                    }, "复制全文"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-kb-ocr-close": "true",
                        onClick: () => patchState({ ocr: null }),
                    }, "关闭")),
            );

            // 预览卡（第九十五轮）：先看清楚「这份文件会入库成什么」，再决定要不要。
            // 文本/PDF：给字符数、块数与开头一段；扫描件/图片：提示需要识别文字；
            // 损坏或不支持的格式：只说原因，不提供任何发送按钮。
            const kbPreviews = Array.isArray(state.kbPreviews) ? state.kbPreviews : [];
            const kbPreviewPanel = kbPreviews.length === 0 ? null : h("div", {
                style: Object.assign({}, S.panel, { padding: "12px 14px" }), "data-kb-preview": "true",
            },
                h("div", { style: S.panelTitle },
                    h("span", { style: S.panelTitleText }, "待入库 · " + String(kbPreviews.length) + " 份"),
                    h("span", { style: S.panelNote }, "只是预览，还没有写入任何内容")),
                kbPreviews.map((item) => {
                    const kind = String(item.kind || "text");
                    // ocr：识别结果已经替换了卡上的正文（第一百零一轮补），与文本同一条
                    // 入库路径，只是写入时带 ocr 标记、顺手收掉旧笔记。
                    const isText = kind === "text" || kind === "ocr";
                    // 图片可以显式入库（存缩略图与图片块）：服务端预览确认为 image 才给
                    // 「入库」；本地只登记了 File 的 >8 MB 卡（unknown）不行，它没有正文。
                    const canIngest = isText || kind === "image";
                    // 显式 OCR 的入口只由类型能力决定，不由体积决定：8–24 MB 的扫描件
                    // 在卡上是 unknown（没判定文本层），同样能点「识别文字」。
                    const canOcr = item.ocrEligible === true
                        || kind === "scan" || kind === "image" || kind === "unknown";
                    return h("div", {
                        key: item.key,
                        style: { borderTop: "1px solid " + M.stroke, paddingTop: "8px", marginTop: "8px" },
                        "data-kb-preview-item": String(item.name || ""),
                        "data-kb-preview-kind": kind,
                    },
                        h("div", { style: { display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" } },
                            h(Badges, { items: [
                                kind === "ocr" ? "识别文字"
                                    : isText ? "文本"
                                    : kind === "scan" ? "扫描件"
                                        : kind === "image" ? "图片"
                                            : kind === "broken" ? "读不出来"
                                                : kind === "unsupported" ? "不支持"
                                                    : kind === "unknown" ? "未预览"
                                                        : kind === "too_large" ? "太大"
                                                            : kind,
                                (humanBytes(item.bytes) === null ? "—" : humanBytes(item.bytes)),
                                isText ? String(item.chars || 0) + " 字 · " + String(item.chunks || 0) + " 块" : null,
                            ] }),
                            h("span", { style: { flex: "1 1 160px", minWidth: 0, color: M.textSecondary, overflow: "hidden", textOverflow: "ellipsis" } }, String(item.name || "")),
                            canIngest
                                ? h("button", {
                                    type: "button", className: "om-btn om-primary", style: S.memPrimary,
                                    "data-kb-preview-ingest": String(item.name || ""),
                                    disabled: Boolean(state.kbUploading),
                                    onClick: () => { ingestPreview(item); },
                                }, state.kbUploading === item.name ? "入库中…" : "入库")
                                : null,
                            canOcr
                                ? h("button", {
                                    type: "button", className: "om-btn", style: S.kbMode,
                                    "data-kb-preview-ocr": String(item.name || ""),
                                    title: "把这份文件的内容发给 DeepSeek OCR 识别文字（可能计费）；不点就不会发",
                                    onClick: () => { runOcrPreview(item); },
                                }, "识别文字")
                                : null,
                            h("button", {
                                type: "button", className: "om-btn", style: S.ghostButton,
                                "data-kb-preview-remove": String(item.name || ""),
                                onClick: () => {
                                    delete kbPreviewDataRef.current[item.key];
                                    patchState({ kbPreviews: (Array.isArray(latest.current.kbPreviews) ? latest.current.kbPreviews : []).filter((one) => one.key !== item.key) });
                                },
                            }, "移除")),
                        // 有「入库」或「识别文字」才有目标可讲：损坏 / 不支持 / 太大的卡
                        // 没有任何投递动作，写一个目标只会让人以为还能入库。
                        canIngest || canOcr
                            ? h("div", {
                                style: Object.assign({}, S.note, { color: M.caption }),
                                "data-kb-preview-folder": String(item.folder || ""),
                            }, (canIngest ? "入库目标：" : "识别后存笔记会进入：")
                                + "「" + (String(item.folder || "") === "" ? "全部文档" : String(item.folder)) + "」")
                            : null,
                        String(item.note || "") === ""
                            ? null
                            : h("div", { style: Object.assign({}, S.note, { color: M.caption }), "data-kb-preview-note": "true" }, String(item.note)),
                        isText && String(item.head || "") !== ""
                            ? h("pre", { style: Object.assign({}, S.kbPre, { maxHeight: "140px", overflow: "auto" }), "data-kb-preview-head": "true" },
                                String(item.head) + (item.truncated ? "\n…（预览到这里，入库才是全文）" : ""))
                            : null);
                }));
            const kbToolbar = h("div", { style: S.kbBar, "data-kb-bar": "true" },
                h("div", { style: S.kbModeTrack, role: "group", "aria-label": "知识库看法" },
                    kbModeButton("docs", "文档（" + (finite(counts.kb) === null ? "—" : counts.kb) + "）"),
                    kbModeButton("search", "块检索"),
                ),
                state.kbMode === "search"
                    ? h("div", { style: S.searchWrap },
                        h("span", { style: S.searchIcon }, h(SearchGlyph, null)),
                        h("input", {
                            type: "text", className: "om-input", "data-kb-search-input": "true",
                            // 受控输入：清空检索、切来源、切页签都要把输入框一起复位。
                            // 早先用 defaultValue 时，DOM 只在挂载时读一次，点「清空检索」
                            // 之后输入框还留着旧词，下一次回车又把旧词查了一遍。
                            value: state.kbQuery,
                            placeholder: "在「" + kbFolderName + "」及子文件夹里检索知识块，回车",
                            onChange: (event) => patchState({ kbQuery: String(event.target.value || "") }),
                            onKeyDown: (event) => {
                                if (event.key !== "Enter") return;
                                const next = String(event.target.value || "").trim();
                                load({
                                    kbMode: "search", kbQuery: next, appliedKb: next,
                                    kbDoc: "", kbChunk: "",
                                    // 顶部搜索框显示同一句草稿；块检索不叠加文档过滤词。
                                    query: next, applied: "",
                                });
                                scrollPanelTop();
                            },
                            style: Object.assign({}, S.searchInput, { width: "300px" }),
                        }),
                    )
                    : null,
                state.kbMode === "search" && state.appliedKb
                    ? h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-kb-search-clear": "true",
                        onClick: () => { load({ appliedKb: "", kbQuery: "", kbDoc: "", kbChunk: "", query: "" }); scrollPanelTop(); },
                    }, "清空检索")
                    : null,
            );

            // 文档下钻的抬头：谁、哪个根、多少块；退回时按来路决定文案。
            const kbDocHeader = kbDocData === null ? null : h("div", { style: S.kbDocHead, "data-kb-doc": "true" },
                h("button", {
                    type: "button", className: "om-btn", style: S.kbBack, "data-kb-back": "doc",
                    onClick: () => { load({ kbDoc: "", kbChunk: "" }); scrollPanelTop(); },
                }, state.kbMode === "search" ? "← 检索结果" : "← 文档清单"),
                h("div", { style: { minWidth: 0 } },
                    h("div", { style: S.kbDocTitle }, String(kbDocData.title || kbDocData.path || "这份文档")),
                    h("div", { style: S.kbDocMeta },
                        String(kbDocData.tier || "—")
                        + " · 共 " + String(kbDocData.chunks_total || 0) + " 块"
                        + (finite(kbDocData.chunks_missing) > 0 ? "，缺 " + String(kbDocData.chunks_missing) + " 块" : "")),
                ),
                kbDocData.synthetic === true || String(kbDocData.path || "").startsWith("uploads/") ? null : h("button", {
                    type: "button", className: "om-btn", style: S.kbBack, "data-kb-edit-text": "true",
                    onClick: () => openKbTextEditor(kbDocData),
                }, "改正文"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.kbBack, { color: M.error }), "data-kb-doc-delete": String(kbDocData.docId || state.kbDoc || ""),
                    onClick: () => askDeleteDoc(kbDocData),
                }, "删除"),
            );

            // 摘要的来源块（第一百零二轮）：摘要单独检索，但「它盖的是哪几块」必须在
            // 摘要页上直接看得到，点一下能跳回原文块。来源块可能已被删/改，缺额单独说。
            const kbSummarySources = kbDocData === null || !Array.isArray(kbDocData.sources) || kbDocData.sources.length === 0
                ? (kbDocData !== null && finite(kbDocData.sourcesMissing) > 0
                    ? h("div", { style: S.kbDocMeta, "data-kb-summary-sources": "missing" },
                        "来源块 " + String(kbDocData.sourcesMissing) + " 块已不在索引里（文档被删或内容变化）。")
                    : null)
                : h("div", { style: { margin: "0 0 12px" }, "data-kb-summary-sources": "true" },
                    h("div", { style: S.memSectionHead },
                        h("span", { style: S.memSectionTitle }, "来源块"),
                        h("span", { style: S.memSectionNote },
                            String(kbDocData.sources.length) + " 块"
                            + (finite(kbDocData.sourcesMissing) > 0 ? " · 另有 " + String(kbDocData.sourcesMissing) + " 块已不在索引里" : "")
                            + " · 点一块回原文")),
                    kbDocData.sources.map((source) => h("button", {
                        key: String(source.id),
                        type: "button", className: "om-btn",
                        style: Object.assign({}, S.browserItem, { display: "block", width: "100%", textAlign: "left", cursor: "pointer", marginBottom: "6px" }),
                        "data-kb-summary-source": String(source.id),
                        onClick: () => {
                            load({
                                kbMode: "docs",
                                kbDoc: String(source.docId || ""),
                                kbChunk: String(source.id || ""),
                                kbRelated: "",
                            });
                            scrollPanelTop();
                        },
                    },
                        h("div", { style: S.kbDocMeta },
                            "第 " + String(source.n || 1) + " 块"
                            + (String(source.spanText || "") !== "" ? " · " + String(source.spanText) : "")
                            + (String(source.docPath || "") !== "" ? " · " + String(source.docPath) : "")),
                        h("div", { style: { fontSize: "12px", color: M.text, lineHeight: "1.5", marginTop: "3px" } },
                            String(source.head || source.docTitle || "").slice(0, 120)))));

            const kbRelatedData = data.kbRelated;
            const kbRelatedList = kbRelatedData && Array.isArray(kbRelatedData.items) && kbRelatedData.items.length > 0
                ? h("div", { style: S.kbRelated, "data-kb-related-list": "true" },
                    h("div", { style: S.kbDocMeta }, "相邻块 · " + kbRelatedData.items.length + " 条"),
                    kbRelatedData.items.map((item) => h("div", {
                        key: item.chunkId,
                        style: { display: "flex", alignItems: "flex-start", gap: "6px" },
                    },
                        h("button", {
                            type: "button", style: Object.assign({}, S.kbRelatedItem, { flex: "1 1 auto", textAlign: "left" }),
                            "data-kb-related-item": String(item.chunkId),
                            onClick: () => {
                                load({ kbDoc: String(item.docId || state.kbDoc), kbChunk: String(item.chunkId), kbRelated: "" });
                                scrollPanelTop();
                            },
                        }, "[" + String(item.kind || "related") + " · 第 " + String(item.n || 1) + " 块] "
                            + String(item.title || item.path || "") + " — " + String(item.text || "").slice(0, 60)),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbBack,
                            "data-kb-unlink": String(item.chunkId),
                            title: "解除这条关系（kb-unlink）",
                            onClick: () => removeChunkLink(String(kbChunkData && kbChunkData.id || state.kbChunk), item),
                        }, "解除"))))
                : null;

            const toggleKbSelected = (id) => {
                const next = Object.assign({}, latest.current.kbSelected || {});
                if (next[id] === true) delete next[id];
                else next[id] = true;
                patchState({ kbSelected: next });
            };
            const selectedKbIds = Object.keys(state.kbSelected || {}).filter((id) => state.kbSelected[id] === true);
            const saveKbSummary = () => {
                const draft = state.kbSummary || {};
                const text = String(draft.text || "").trim();
                if (selectedKbIds.length < 2) { patchState({ actionError: "摘要至少选两块。" }); return; }
                if (text === "") { patchState({ actionError: "摘要正文不能为空。" }); return; }
                mutatePanel("kb-summarize", {
                    title: String(draft.title || "").trim() || ("摘要 · " + selectedKbIds.length + " 块"),
                    text: text, chunkIds: selectedKbIds,
                });
            };
            const kbSummaryPanel = state.kbSummary === null ? null : h("div", { style: S.editor, "data-kb-summary-editor": "true" },
                h("div", { style: S.editorTitle }, "把选中的 " + selectedKbIds.length + " 块合成摘要块"),
                h("label", { style: S.editorLabel }, "标题",
                    h("input", {
                        type: "text", style: S.editorInput, "data-kb-summary-title": "true",
                        value: state.kbSummary.title || "",
                        onChange: (event) => patchState({ kbSummary: Object.assign({}, latest.current.kbSummary || {}, { title: String(event.target.value || "") }) }),
                    })),
                h("label", { style: S.editorLabel }, "摘要正文",
                    h("textarea", {
                        style: S.editorTextarea, "data-kb-summary-text": "true",
                        value: state.kbSummary.text || "",
                        onChange: (event) => patchState({ kbSummary: Object.assign({}, latest.current.kbSummary || {}, { text: String(event.target.value || "") }) }),
                    })),
                h("div", { style: S.editorActions },
                    h("button", {
                        type: "button", className: "om-btn om-primary", style: S.memPrimary,
                        "data-kb-summary-save": "true", disabled: state.saving === true,
                        onClick: saveKbSummary,
                    }, state.saving === true ? "保存中…" : "保存摘要"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-kb-summary-cancel": "true",
                        onClick: () => patchState({ kbSummary: null }),
                    }, "取消")),
                state.actionError
                    ? h("div", { style: S.editorError }, String(state.actionError))
                    : null,
            );

            // 块间关系编辑器：source 固定为当前块，目标是索引里另一个块 id。
            // 保存后重新取相邻块，关系列表里立刻能看到（服务端 kb-link 会校验两端存在）。
            const renderChunkLinkEditor = (chunkData) => {
                const draft = state.kbLink;
                if (draft === null || draft === undefined) return null;
                if (String(draft.sourceId || "") !== String(chunkData.id || "")) return null;
                const set = (patch) => patchState({ kbLink: Object.assign({}, latest.current.kbLink || {}, patch) });
                return h("div", { style: S.editor, "data-kb-link-editor": "true" },
                    h("div", { style: S.editorTitle }, "给这一块加一条关系"),
                    h("div", { style: { display: "grid", gridTemplateColumns: "minmax(0,1fr) 160px", gap: "8px" } },
                        h("label", { style: S.editorLabel }, "目标块 id",
                            h("input", {
                                type: "text", style: S.editorInput, "data-kb-link-target": "true",
                                value: String(draft.targetId || ""),
                                placeholder: "kb:<hash>:<n>",
                                onChange: (event) => set({ targetId: String(event.target.value || "") }),
                            })),
                        h("label", { style: S.editorLabel }, "关系",
                            h("select", {
                                style: S.editorInput, "data-kb-link-kind": "true",
                                value: KB_LINK_KINDS.indexOf(String(draft.kind || "")) >= 0 ? String(draft.kind) : "related",
                                onChange: (event) => set({ kind: String(event.target.value || "related") }),
                            }, KB_LINK_KINDS.map((kind) => h("option", { key: kind, value: kind }, kind))))),
                    h("label", { style: S.editorLabel }, "备注（可选）",
                        h("input", {
                            type: "text", style: S.editorInput, "data-kb-link-note": "true",
                            value: String(draft.note || ""),
                            onChange: (event) => set({ note: String(event.target.value || "") }),
                        })),
                    h("div", { style: S.editorActions },
                        h("button", {
                            type: "button", className: "om-btn om-primary", style: S.memPrimary,
                            "data-kb-link-save": "true", disabled: state.saving === true,
                            onClick: saveChunkLink,
                        }, state.saving === true ? "保存中…" : "保存关系"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton,
                            "data-kb-link-cancel": "true",
                            onClick: () => patchState({ kbLink: null }),
                        }, "取消")),
                    state.actionError
                        ? h("div", { style: S.editorError }, String(state.actionError))
                        : null,
                );
            };

            // 块正文：只显示当前选中的那一块，不把整份文档搬进面板。
            // 位置很关键：正文要跟在「选中那一块」的卡片后面，而不是插在清单上方——
            // 否则用户点第 1 块，正文跑到清单外，选中的卡片还留在下方，看起来像多出一条。
            const kbChunkView = (chunkData) => chunkData === null ? null : h("div", {
                style: { marginBottom: "10px" }, "data-kb-chunk-text": String(chunkData.id || ""),
            },
                h("div", { style: S.kbDocHead },
                    h("div", { style: { minWidth: 0, flex: "1 1 260px" } },
                        h("div", { style: S.kbDocTitle },
                            "第 " + String(chunkData.n || 1) + " 块正文"),
                        h("div", { style: S.kbDocMeta },
                            String(chunkData.chars || 0) + " 字"
                            + (chunkData.spanText ? " · " + String(chunkData.spanText) : "")),
                    ),
                    h("button", {
                        type: "button", className: "om-btn", style: S.kbBack, "data-kb-chunk-close": "true",
                        onClick: () => { load({ kbChunk: "", kbRelated: "" }); },
                    }, "收起正文"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.kbBack, "data-kb-related-open": String(chunkData.id || ""),
                        onClick: () => { load({ kbRelated: String(chunkData.id || ""), kbChunk: String(chunkData.id || "") }); },
                    }, "看相邻块"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.kbBack, "data-kb-link-open": String(chunkData.id || ""),
                        onClick: () => patchState({
                            actionError: "",
                            kbLink: { sourceId: String(chunkData.id || ""), targetId: "", kind: "related", note: "" },
                        }),
                    }, "加关系"),
                    h("button", {
                        type: "button", className: "om-btn",
                        style: Object.assign({}, S.kbBack, { color: M.error }), "data-kb-chunk-delete": String(chunkData.id || ""),
                        onClick: () => askDeleteDoc(kbDocData || chunkData),
                    }, "删除文档"),
                ),
                renderChunkLinkEditor(chunkData),
                h("pre", { style: S.kbPre, "data-kb-chunk-body": "true" }, String(chunkData.text || "")),
                // 上传时抽出的图片以 Markdown 引用留在正文里；这里把它们真的画出来，
                // 用户才看得到「只保留切片和图片」中的图片部分。
                (() => {
                    const text = String(chunkData.text || "");
                    const refs = [];
                    const pattern = /!\[[^\]]*\]\((assets\/[^)\s]+)\)/g;
                    let found = pattern.exec(text);
                    while (found !== null) {
                        if (refs.indexOf(found[1]) < 0) refs.push(found[1]);
                        found = pattern.exec(text);
                    }
                    return refs.length === 0 ? null : h("div", { style: { display: "flex", gap: "10px", flexWrap: "wrap", margin: "2px 0 12px" }, "data-kb-assets": "true" },
                        refs.map((ref) => h("a", {
                            key: ref, href: KB_ASSET_PATH + "?cwd=" + encodeURIComponent(cwdFor || "") + "&path=" + encodeURIComponent(ref),
                            target: "_blank", rel: "noreferrer",
                        }, h("img", {
                            src: KB_ASSET_PATH + "?cwd=" + encodeURIComponent(cwdFor || "") + "&path=" + encodeURIComponent(ref),
                            alt: ref, loading: "lazy",
                            style: { maxWidth: "260px", maxHeight: "180px", borderRadius: "9px", border: "1px solid " + M.stroke, background: M.surface },
                        }))));
                })(),
                kbRelatedList,
            );

            const kbChunkList = kbDocData === null ? null : h("div", null,
                // 多块选择（第八十五轮）：摘要块的原料是「选中的这几块」，所以选择态
                // 与「读正文」分开 —— 勾选不会切走当前正文，读正文也不会清掉勾选。
                h("div", { style: S.kbSelectBar, "data-kb-select-bar": "true" },
                    h("span", null, "已选 " + selectedKbIds.length + " 块"),
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-kb-summarize-open": "true",
                        disabled: selectedKbIds.length < 2,
                        onClick: () => patchState({ kbSummary: { title: "", text: "" }, actionError: "" }),
                    }, "合成摘要"),
                    selectedKbIds.length > 0
                        ? h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton,
                            "data-kb-select-clear": "true",
                            onClick: () => patchState({ kbSelected: {} }),
                        }, "清空选择")
                        : null),
                kbSummaryPanel,
                kbDocChunks.length === 0
                    ? h("div", { style: S.emptyState, "data-kb-chunks-empty": "true" },
                        h("div", { style: S.emptyTitle }, "这份文档没有可读的块。"),
                        h("div", { style: S.emptyHint }, "块文件可能缺失或损坏，服务端的 kb-read 会给出具体原因。"))
                    : h("div", { style: S.browserList, "data-kb-chunks": "true" }, kbDocChunks.map((chunk) => {
                        const selected = String(state.kbChunk) === String(chunk.id);
                        const checked = state.kbSelected[String(chunk.id)] === true;
                        const card = h("div", {
                            key: String(chunk.id),
                            style: Object.assign({}, S.browserItem, { cursor: "pointer" }, selected ? S.kbChunkCardOn : {}),
                            "data-kb-chunk": String(chunk.id),
                            onClick: () => {
                                load({ kbDoc: String(kbDocData.docId || state.kbDoc), kbChunk: String(chunk.id), kbRelated: "" });
                            },
                        },
                            h("div", { style: { display: "flex", alignItems: "center", gap: "8px" } },
                                h("input", {
                                    type: "checkbox", checked: checked,
                                    "data-kb-select": String(chunk.id),
                                    "aria-label": "选中第 " + String(chunk.n || 1) + " 块用于摘要",
                                    onClick: (event) => { if (event && typeof event.stopPropagation === "function") event.stopPropagation(); },
                                    onChange: () => toggleKbSelected(String(chunk.id)),
                                }),
                                h(Badges, { items: ["第 " + String(chunk.n || 1) + " 块", String(chunk.chars || 0) + " 字"] })),
                            h("div", { style: S.kbDocTitle }, String(chunk.header || chunk.outline || "（无块头）")),
                            h("div", { style: S.kbDocMeta }, String(chunk.spanText || "")),
                        );
                        // 正文紧跟在选中的卡片下面：点哪一块，正文就在哪一块下面展开。
                        return selected && kbChunkData !== null
                            ? h(React.Fragment, { key: String(chunk.id) + "-row" }, card, kbChunkView(kbChunkData))
                            : card;
                    })),
            );

            const kbSearchView = kbSearchData === null
                ? h("div", { style: S.emptyState, "data-kb-search-empty": "true" },
                    h("div", { style: S.emptyGlyph }, "⌕"),
                    h("div", { style: S.emptyTitle }, state.appliedKb ? "没有拿到检索结果。" : "按内容检索知识块。"),
                    h("div", { style: S.emptyHint }, state.appliedKb
                        ? "服务端这次没有返回命中集，切回文档清单或换一个词再试。"
                        : "输入关键词后回车：检索正文与块头（中文按 bigram），命中给出块号、路径与分数。"))
                : h("div", { "data-kb-search": "true" },
                    h("div", { style: S.kbSearchMeta },
                        (String(state.appliedKb || "") === "" ? "" : "搜索「" + String(state.appliedKb) + "」 · ")
                        + "命中 " + String(finite(kbSearchData.matched) === null ? kbSearchHits.length : kbSearchData.matched)
                        + " 块" + (kbSearchData.truncated === true ? "（已截断，只列前 " + String(kbSearchHits.length) + " 块）" : "")
                        + (kbReasonText(kbSearchData.reason) === "" ? "" : " · " + kbReasonText(kbSearchData.reason))),
                    kbSearchHits.length === 0
                        ? h("div", { style: S.emptyState, "data-kb-search-empty": "true" },
                            h("div", { style: S.emptyTitle }, "没有命中。"),
                            h("div", { style: S.emptyHint }, kbReasonText(kbSearchData.reason) || "换一个更具体的词，或切回文档清单按路径找。"))
                        : h("div", { style: S.browserList, "data-kb-hits": "true" }, kbSearchHits.map((hit) => {
                            const score = finite(hit.score);
                            return h("div", {
                                key: String(hit.chunkId),
                                style: Object.assign({}, S.browserItem, { cursor: "pointer" }),
                                "data-kb-hit": String(hit.chunkId),
                                role: "button", tabIndex: 0,
                                "aria-label": "打开第 " + String(hit.n || 1) + " 块的正文",
                                onClick: () => {
                                    load({ kbDoc: String(hit.docId || ""), kbChunk: String(hit.chunkId || "") });
                                    scrollPanelTop();
                                },
                                onKeyDown: (event) => {
                                    if (!event || (event.key !== "Enter" && event.key !== " ")) return;
                                    event.preventDefault();
                                    load({ kbDoc: String(hit.docId || ""), kbChunk: String(hit.chunkId || "") });
                                    scrollPanelTop();
                                },
                            },
                                h(Badges, { items: [
                                    hit.tier,
                                    "第 " + String(hit.n || 1) + " 块",
                                    score === null ? null : "分 " + score.toFixed(2),
                                ] }),
                                h("div", { style: S.kbDocTitle }, String(hit.title || hit.path || "")),
                                h("div", { style: S.kbDocMeta }, String(hit.path || "") + (hit.spanText ? " · " + String(hit.spanText) : "")),
                                h("div", { style: S.browserBody }, String(hit.preview || "")),
                            );
                        })),
                );

            // 文档清单空的时候不能只画一个空 div：知识库是独立的一棵树，空库
            // （或来源筛选后为空）要给出「下一步怎么入库 / 怎么切回去」。
            const kbEmptyTitle = kbFolderNow !== ""
                ? "这个文件夹是空的。"
                : state.origin !== ""
                    ? "这个来源下没有知识库文档。"
                    : state.applied !== "" ? "没有匹配的知识库文档。" : "知识库还是空的。";
            const kbEmptyHint = kbFolderNow !== ""
                ? "点命令行「新建文件夹」建下一层；这一层没有文档，检索范围切到它也是空的。"
                : state.origin !== ""
                    ? "切回「全部」或换一个来源；来源筛选只影响这一页列出的文档。"
                    : state.applied !== ""
                        ? "换一个词，或清空搜索框看全部文档。"
                        : "点右上角「＋ 添加文件」选文件，或把文件拖进这一页；先预览、确认后再入库，入库后只保留切片与图片，原文件不会复制进知识库。";
            const kbEmptyAction = state.applied !== ""
                ? h("button", {
                    type: "button", className: "om-btn", style: S.ghostButton,
                    "data-kb-empty-action": "clear",
                    onClick: () => {
                        load({ applied: "", query: "", appliedKb: "", kbQuery: "", kbDoc: "", kbChunk: "" });
                        scrollPanelTop();
                    },
                }, "清空搜索")
                : kbFolderNow !== ""
                    ? h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-kb-empty-action": "root", onClick: () => kbGoFolder(""),
                    }, "全部文档")
                    : h("button", {
                        type: "button", className: "om-btn", style: S.kbIngestBtn,
                        "data-kb-empty-action": "upload", disabled: Boolean(state.kbUploading),
                        onClick: pickFiles,
                    }, "＋ 添加文件");
            const kbEmpty = h("div", { style: S.emptyState, "data-kb-empty": "true" },
                h("div", { style: S.emptyGlyph }, "▤"),
                h("div", { style: S.emptyTitle }, kbEmptyTitle),
                h("div", { style: S.emptyHint }, kbEmptyHint),
                h("div", { style: { display: "flex", justifyContent: "center", marginTop: "10px" } }, kbEmptyAction));
            // 摘要块清单（第八十五轮 / 第一百零七轮）：它们是独立的检索单元，但
            // 摘要正文很长 —— 全部铺在文档清单上方，会把文档顶到几屏之外。现在
            // 放在文档之后，并默认折叠：收起时只留标题与来源块数，展开才看预览。
            const kbSummaryMeta = (summary) => {
                const chunks = Array.isArray(summary.chunkIds) ? summary.chunkIds.length : 0;
                const docPath = String(summary.docPath || "");
                const created = String(summary.createdAt || "");
                return chunks + " 块来源"
                    + (docPath !== "" ? " · " + docPath : "")
                    + (created !== "" ? " · " + created.slice(0, 10) : "");
            };
            const kbSummaryPage = data.kbSummariesPage !== null && typeof data.kbSummariesPage === "object"
                ? data.kbSummariesPage : null;
            const kbSummaryPager = kbSummaryPage === null ? null : h("div", {
                style: { display: "flex", alignItems: "center", gap: "6px", marginTop: "8px", flexWrap: "wrap", fontSize: "11px", color: M.caption },
                "data-kb-summaries-pager": "true",
            },
                h("button", {
                    type: "button", className: "om-btn", style: S.ghostButton,
                    "data-kb-summary-page-prev": "true", disabled: kbSummaryPage.hasPrev !== true,
                    onClick: () => load({
                        kbSummaryOffset: Math.max(0, Number(kbSummaryPage.offset) - Number(kbSummaryPage.limit)),
                        kbSummaryLimit: Number(kbSummaryPage.limit) || 10,
                    }),
                }, "上一页"),
                h("span", { "data-kb-summary-page-info": "true" },
                    (Number(kbSummaryPage.total) === 0 ? "共 0 条" :
                        "第 " + (Number(kbSummaryPage.offset) + 1) + "–"
                        + Math.min(Number(kbSummaryPage.total), Number(kbSummaryPage.offset) + Number(kbSummaryPage.limit))
                        + " 条 / 共 " + Number(kbSummaryPage.total) + " 条")),
                h("button", {
                    type: "button", className: "om-btn", style: S.ghostButton,
                    "data-kb-summary-page-next": "true", disabled: kbSummaryPage.hasNext !== true,
                    onClick: () => load({
                        kbSummaryOffset: Number(kbSummaryPage.offset) + Number(kbSummaryPage.limit),
                        kbSummaryLimit: Number(kbSummaryPage.limit) || 10,
                    }),
                }, "下一页"),
                h("select", {
                    className: "om-input", "data-kb-summary-page-size": "true",
                    value: String(kbSummaryPage.limit || 10),
                    onChange: (event) => load({
                        kbSummaryLimit: Number(event.target.value) || 10,
                        kbSummaryOffset: 0,
                    }),
                    style: Object.assign({}, S.input, { width: "auto", padding: "3px 6px", fontSize: "11px" }),
                }, SUMMARY_PAGE_SIZES.map((size) => h("option", { key: size, value: String(size) }, String(size) + " 条/页"))));
            const kbSummariesList = Array.isArray(data.kbSummaries) && data.kbSummaries.length > 0
                ? h("div", { style: { marginTop: "16px", marginBottom: "14px" }, "data-kb-summaries": "true" },
                    h("div", { style: S.memSectionHead },
                        h("span", { style: Object.assign({}, S.memSectionDot, { background: metricColor("kb") }) }),
                        h("span", { style: S.memSectionTitle }, "摘要块"),
                        h("span", { style: S.memSectionNote },
                            (kbSummaryPage === null ? data.kbSummaries.length : Number(kbSummaryPage.total) || 0) + " 条 · 多块合成")),
                    data.kbSummaries.map((summary) => h(Fold, {
                        key: summary.id,
                        name: "kb-summary-" + String(summary.id || ""),
                        label: String(summary.title || summary.id),
                        count: (Array.isArray(summary.chunkIds) ? summary.chunkIds.length : 0) + " 块",
                        defaultOpen: false,
                        style: { marginBottom: "6px" },
                    },
                        h("div", { style: { padding: "2px 2px 4px" } },
                            h("div", { style: S.kbDocMeta }, kbSummaryMeta(summary)),
                            h("div", { style: S.kbDocMeta }, String(summary.text || "").slice(0, 400)),
                            h("button", {
                                type: "button", className: "om-btn",
                                style: Object.assign({}, S.ghostButton, { marginTop: "6px" }),
                                "data-kb-summary": String(summary.id),
                                onClick: () => {
                                    // 摘要已经是独立检索单元：点它看摘要本身（来源块仍可从摘要抬头回到文档）。
                                    load({ kbDoc: "summary:" + String(summary.id || ""), kbChunk: "summary:" + String(summary.id || "") + ":1", kbRelated: "" });
                                    scrollPanelTop();
                                },
                            }, "打开摘要")))),
                    kbSummaryPager)
                : null;

            const kbContent = kbDocData !== null
                ? h("div", { "data-kb-drill": "true" }, kbDocHeader, kbSummarySources, kbChunkList, kbDocData === null ? kbChunkView(kbChunkData) : null)
                : state.kbMode === "search"
                    ? kbSearchView
                    : total === 0 && kbFolderRows.length === 0
                        ? kbEmpty
                        : h("div", null,
                            // 第一百零七轮：文档在前、摘要块在后。摘要块默认折叠，
                            // 不把整页滚动长度交给「每条摘要的正文」。
                            h("div", { style: S.memSectionHead, "data-memory-section": "kb" },
                                h("span", { style: Object.assign({}, S.memSectionDot, { background: metricColor("kb") }) }),
                                h("span", { style: S.memSectionTitle }, "文档 · " + kbFolderName),
                                h("span", { style: S.memSectionNote },
                                    (state.kbRecursive === false ? "仅当前层" : "含子文件夹") + " · " + total + " 篇")),
                            // 知识库用服务端分页控件（在 kbStatusBar 里），
                            // 不再复用记忆页签的本地「再显示 N 条」。
                            kbNote, kbExplorerList, kbStatusBar, detailBar, kbSummariesList);
            // 归档在现行 UI 里是知识库的虚拟文件夹（不再有独立归档页签）：只有进到
            // 「归档」子树才显示隔离提示，空文件夹也显示；其它文件夹的 DOM 不变。
            const archiveFolderView = state.tab === "kb" && (kbFolderNow === "归档" || kbFolderNow.startsWith("归档/"));
            const kbContentWithWarnings = archiveFolderView && archiveWarningBox !== null
                ? h("div", { "data-kb-archive-warning-wrap": "true" }, archiveWarningBox, kbContent)
                : kbContent;
            // 下钻失败不能只剩一个「没有命中」：服务端已经把原因（块文件缺失、
            // 文档不在这个来源里等）写进 kbSearchError / kbDocError / kbChunkError，
            // 这里原样顶在内容上方，和预览面板的错误处理同一口径。
            const kbDrillError = data.kbSearchError || data.kbDocError || data.kbChunkError || data.kbRelatedError || data.kbSummariesError || "";
            const kbErrorNote = kbDrillError === ""
                ? null
                : h("div", { style: S.error, "data-kb-error": "true" }, String(kbDrillError));

            // 知识库的主列：面包屑（我在哪一层）→ 工具条（文档 / 块检索）→ 内容。
            // 左树与主列同属一个资源管理器版式；别的页签仍是单列，不套这层壳。
            const kbMain = h("div", {
                style: Object.assign({}, S.kbMain, state.kbUploading ? { outline: "2px dashed " + M.stroke, outlineOffset: "-6px" } : {}),
                "data-kb-dropzone": "true",
                onDragOver: (event) => { event.preventDefault(); if (event.dataTransfer) event.dataTransfer.dropEffect = "copy"; },
                onDrop: (event) => {
                    event.preventDefault();
                    const files = event.dataTransfer && event.dataTransfer.files;
                    if (files && files.length > 0) previewFiles(files);
                },
            },
                kbCommandBar,
                kbNewFolderBar,
                kbFolderErrorBar,
                kbFolderNoticeBar,
                kbFoldersErrorBar,
                kbCrumb,
                kbAiBar,
                kbUploadNote,
                kbEstimatePanel,
                kbPausedPanel,
                kbToolbar,
                kbPreviewPanel,
                ocrPanel,
                state.kbUploading
                    ? (kbProgressBar === null
                        ? h("div", { style: Object.assign({}, dim(T.note), { padding: "6px 0" }), "data-kb-uploading": "true" }, "正在处理 " + String(state.kbUploading) + "（抽取 → 切片 → 只留块与图片）…")
                        : kbProgressBar)
                    : null,
                state.actionError && state.editor === null
                    ? h("div", { style: S.editorError, "data-memory-action-error": "true" }, String(state.actionError))
                    : null,
                kbErrorNote, editorPanel, kbContentWithWarnings);
            const kbExplorer = h("div", { className: "om-kb-explorer", style: S.kbExplorer, "data-kb-explorer": "true" }, kbTreeSide, kbMain);

            // ── 图库（独立一级页签）────────────────────────────────────────
            // 数据面：GET 列表 / 搜索 / 状态，POST 上传、手工标签、描述、向量、删除。
            // 云端能力一律「先勾选允许、再显式点按钮」：未配置或未探针时按钮禁用并写原因，
            // 不画看起来能用、点下去静默降级的假能力。

            /** 当前列表（data-gallery-item 的排序与选中都基于它）。 */
            function galleryItemsNow() {
                const owner = latest.current;
                const body = owner.galleryData && typeof owner.galleryData === "object" ? owner.galleryData : null;
                return body && Array.isArray(body.items) ? body.items : [];
            }
            function galleryItemById(id) {
                const wanted = String(id || "");
                return galleryItemsNow().find((item) => String(item && item.id || "") === wanted) || null;
            }
            function galleryErrorMessage(error) {
                return String(error && error.message ? error.message : error);
            }
            function gallerySelectItem(id, item) {
                const next = item || galleryItemById(id);
                patchState({
                    gallerySelected: String(id || ""),
                    galleryDraft: next === null ? null : galleryDraftOf(next),
                    galleryDescribeConfirm: false,
                    galleryRemoveConfirm: false,
                    galleryFolderRemoveConfirm: false,
                    galleryPreview: false,
                    galleryError: "",
                });
            }
            /** 筛选条件是一次 GET 的一部分：切分类 / 收藏 / 搜索都要回到第一页并保留选中。 */
            function galleryLoadFilter(patch) {
                // 换浏览层 / 筛选时，新建文件夹条不该跟着留在新位置。
                // 清掉文字检索（patch 把 galleryApplied 置空）时，顶部输入框也要一起复位：
                // 否则框里还留着旧词，下一次回车又按旧词查一遍。
                const merged = Object.assign({ galleryFolderNewOpen: false, galleryFolderName: "" }, patch);
                if (Object.prototype.hasOwnProperty.call(patch || {}, "galleryApplied")
                    && String(patch.galleryApplied || "") === "") merged.query = "";
                if (patch && patch.galleryImageSearch === true) {
                    // 以图搜图的结果只有相似度有意义，别按加入时间把最像的一张排到后面。
                    merged.gallerySort = "score";
                    merged.gallerySortDir = "desc";
                }
                const stillSearching = String(merged.galleryApplied || "") !== "" || merged.galleryImageSearch === true;
                const clearedSearch = (Object.prototype.hasOwnProperty.call(patch || {}, "galleryApplied")
                    && String(patch.galleryApplied || "") === "")
                    || (patch && patch.galleryImageSearch === false);
                const currentSort = Object.prototype.hasOwnProperty.call(merged, "gallerySort")
                    ? merged.gallerySort : latest.current.gallerySort;
                if (clearedSearch && !stillSearching && currentSort === "score") {
                    merged.gallerySort = "added";
                    merged.gallerySortDir = "desc";
                }
                loadGallery(merged);
                scrollPanelTop();
            }
            /**
             * 换文件夹：只改浏览位置，不把文件夹塞进检索条件。
             *
             * 和知识库一样维护一层历史栈（后退恢复上一路径，不缓存清单），
             * 进入 / 后退都清掉当前选中：选中的图已经不在这一层，详情不该继续挂着。
             */
            function galleryGoFolder(next, mode) {
                const target = String(next || "");
                const current = String(latest.current.galleryFolder || "");
                if (mode !== "back" && target === current) return;
                const stack = Array.isArray(latest.current.galleryBack) ? latest.current.galleryBack : [];
                const nextStack = mode === "back"
                    ? stack.slice(0, Math.max(0, stack.length - 1))
                    : stack.concat([current]);
                galleryLoadFilter({
                    galleryFolder: target, galleryBack: nextStack,
                    galleryApplied: "", galleryQuery: "", galleryTags: "", galleryCategory: "",
                    galleryFavorite: false, galleryImageSearch: false,
                    galleryFolderRemoveConfirm: false,
                    gallerySelected: "", galleryDraft: null, galleryPreview: false,
                });
            }
            async function galleryReadFile(file) {
                return await new Promise((settle, fail) => {
                    const reader = new FileReader();
                    reader.onload = () => {
                        const text = String(reader.result || "");
                        const at = text.indexOf("base64,");
                        settle(at >= 0 ? text.slice(at + 7) : "");
                    };
                    reader.onerror = () => fail(new Error("读不了这个文件。"));
                    reader.readAsDataURL(file);
                });
            }
            async function galleryUploadFiles(fileList) {
                const files = Array.from(fileList || []);
                if (files.length === 0) return;
                let added = 0;
                patchState({ galleryError: "", galleryNotice: "" });
                for (const file of files) {
                    patchState({ galleryUploading: String(file.name || "图片") });
                    try {
                        const encoded = await galleryReadFile(file);
                        if (encoded === "") throw new Error(String(file.name || "图片") + " 的内容是空的。");
                        const body = await galleryAction("add", {
                            name: String(file.name || ""), data: encoded,
                            folder: String(latest.current.galleryFolder || ""),
                        }, {});
                        if (body !== null) {
                            if (body.item) gallerySelectItem(body.item.id, body.item);
                            if (body.duplicate === true) {
                                // 重复内容不是新增：不把它算进「已加入 N 张」，提示也留在这一句。
                                patchState({ galleryNotice: "这张图已经在图库里（内容相同），已选中原来那张。" });
                            } else {
                                added += 1;
                            }
                        }
                    } catch (error) {
                        patchState({ galleryError: "上传失败：" + galleryErrorMessage(error) + (added > 0 ? "（已成功 " + added + " 张）" : "") });
                        break;
                    }
                }
                patchState({ galleryUploading: "" });
                if (added > 0 && latest.current.galleryError === "") {
                    // 上传会把图片归到当前文件夹，直接计数也跟着变。
                    loadGalleryStatus();
                    patchState({ galleryNotice: "已加入图库：" + added + " 张。" });
                }
            }
            function galleryPickFiles() {
                const input = document.querySelector("[data-gallery-upload-input]");
                if (input && typeof input.click === "function") input.click();
            }
            function gallerySaveDraft() {
                const draft = latest.current.galleryDraft;
                if (draft === null || draft === undefined || String(draft.id || "") === "") return;
                galleryAction("manual", {
                    id: String(draft.id),
                    patch: {
                        name: String(draft.name || "").trim(),
                        description: String(draft.description || ""),
                        tags: galleryTagList(draft.tags),
                        category: String(draft.category || "").trim(),
                        folder: String(draft.folder || "").trim(),
                        note: String(draft.note || ""),
                        favorite: draft.favorite === true,
                    },
                }, {}).then((body) => {
                    if (body && body.item) {
                        gallerySelectItem(body.item.id, body.item);
                        // 新文件夹可能刚由这次移动创建：目录树与计数要跟着刷新。
                        loadGalleryStatus();
                        patchState({ galleryNotice: "细节已保存。" });
                    }
                });
            }
            function galleryFolderParentOf(path) {
                const text = String(path || "");
                const at = text.lastIndexOf("/");
                return at < 0 ? "" : text.slice(0, at);
            }
            function galleryFolderLeafOf(path) {
                const text = String(path || "");
                const at = text.lastIndexOf("/");
                return at < 0 ? text : text.slice(at + 1);
            }
            /** 新建文件夹：建在当前文件夹下；成功后直接进去，空目录也看得见。 */
            function galleryCreateFolder() {
                const name = String(latest.current.galleryFolderName || "").trim();
                if (name === "") {
                    patchState({ galleryError: "先给文件夹起个名字。" });
                    return;
                }
                const parent = String(latest.current.galleryFolder || "");
                const path = parent === "" ? name : parent + "/" + name;
                const stack = Array.isArray(latest.current.galleryBack) ? latest.current.galleryBack : [];
                galleryAction("folder-create", { path: path }, {
                    galleryFolder: path, galleryBack: stack.concat([parent]),
                    galleryFolderName: "", galleryFolderNewOpen: false,
                    gallerySelected: "", galleryDraft: null,
                }).then((body) => {
                    if (body === null) return;
                    loadGalleryStatus();
                    patchState({ galleryNotice: "已创建文件夹：" + path + "。" });
                });
            }
            /** 改名只改当前文件夹的最后一段，父路径保持不变。 */
            function galleryRenameFolder() {
                const from = String(latest.current.galleryFolder || "");
                if (from === "") return;
                const nextName = String(latest.current.galleryFolderEditValue || "").trim();
                if (nextName === "") {
                    patchState({ galleryError: "文件夹名不能为空。" });
                    return;
                }
                if (nextName === galleryFolderLeafOf(from)) {
                    patchState({ galleryFolderEdit: "", galleryFolderEditValue: "", galleryError: "" });
                    return;
                }
                const parent = galleryFolderParentOf(from);
                const to = parent === "" ? nextName : parent + "/" + nextName;
                galleryAction("folder-rename", { from: from, to: to }, {
                    galleryFolder: to, galleryFolderEdit: "", galleryFolderEditValue: "",
                }).then((body) => {
                    if (body === null) return;
                    loadGalleryStatus();
                    patchState({ galleryNotice: "已改名：" + from + " → " + to + "。" });
                });
            }
            /** 非空目录服务端会拒绝并给出张数；这里只负责两步确认。 */
            function galleryRemoveFolder() {
                const path = String(latest.current.galleryFolder || "");
                if (path === "") return;
                if (latest.current.galleryFolderRemoveConfirm !== true) {
                    patchState({ galleryFolderRemoveConfirm: true, galleryError: "" });
                    return;
                }
                const parent = galleryFolderParentOf(path);
                galleryAction("folder-remove", { path: path }, {
                    galleryFolder: parent, galleryFolderRemoveConfirm: false,
                    galleryFolderEdit: "", galleryFolderEditValue: "",
                }).then((body) => {
                    if (body === null) return;
                    loadGalleryStatus();
                    patchState({ galleryNotice: "已删除文件夹：" + path + "。" });
                });
            }
            function galleryDescribeRequest() {
                const selected = latest.current.gallerySelected;
                const item = galleryItemById(selected);
                if (item === null) { patchState({ galleryError: "先选一张图，再生成描述。" }); return; }
                if (latest.current.galleryAllowUpload !== true) {
                    patchState({ galleryError: "生成视觉描述会把原图发给视觉模型（可能产生云端费用）：先勾选「允许内容发给模型」。" });
                    return;
                }
                const status = latest.current.galleryStatus;
                const vision = status && status.vision ? status.vision : null;
                // 缓存里的能力状态可能停在设置保存之前：已经配置好视觉模型的人点一次
                // 「生成视觉描述」却被告知未配置，就是这里拿旧状态挡掉的（第一百零七轮）。
                // 未配置时先向服务端重取一次状态；仍不可用才报错。已配置则直接确认。
                if (vision !== null && vision.configured === true) {
                    patchState({ galleryDescribeConfirm: true, galleryError: "" });
                    return;
                }
                loadGalleryStatus().then((fresh) => {
                    const current = fresh && fresh.vision ? fresh.vision : null;
                    if (current !== null && current.configured !== true) {
                        patchState({ galleryError: "视觉描述未配置：" + String(current.note || "服务端没有可用的视觉模型。")
                            + " 去「设置 → 办公与记忆 → 图库 → 视觉描述」选一个通道。" });
                        return;
                    }
                    // 状态取不到时不再用旧缓存拦住动作：真未配置由服务端回执说明。
                    patchState({ galleryDescribeConfirm: true, galleryError: "" });
                });
            }
            function galleryDescribeConfirm() {
                const selected = latest.current.gallerySelected;
                if (selected === "") return;
                // 两步确认之间用户可能又把「允许内容发给模型」撤掉了：确认按钮还在
                // 屏幕上，但授权已经收回，这时候不能照着旧状态发 allowUpload。
                if (latest.current.galleryAllowUpload !== true) {
                    patchState({ galleryDescribeConfirm: false, galleryError: "授权已撤回：要生成视觉描述，请重新勾选「允许内容发给模型」。" });
                    return;
                }
                patchState({ galleryDescribeConfirm: false });
                galleryAction("describe", { id: selected, allowUpload: true }, {}).then((body) => {
                    if (body && body.item) gallerySelectItem(body.item.id, body.item);
                    if (body !== null) {
                        patchState({ galleryNotice: String(body.note || "视觉描述已生成；未声明的能力不会被当成已确认。") });
                    }
                    loadGalleryStatus();
                });
            }
            function galleryProbe() {
                if (latest.current.galleryAllowUpload !== true) {
                    patchState({ galleryError: "探针会向多模态端点发一条文本和一张 1×1 图片（可能计费）：先勾选「允许内容发给模型」。" });
                    return;
                }
                galleryAction("probe", { allowUpload: true }, {}).then((body) => {
                    if (body !== null) patchState({ galleryNotice: "探针通过：" + String(body.dim || "?") + " 维 · " + String(body.model || body.provider || "") });
                    loadGalleryStatus();
                });
            }
            function galleryVectors() {
                const selected = latest.current.gallerySelected;
                if (selected === "") { patchState({ galleryError: "先选一张图，再建向量索引。" }); return; }
                if (latest.current.galleryAllowUpload !== true) {
                    patchState({ galleryError: "图片向量化会把原图发给多模态端点（可能计费）：先勾选「允许内容发给模型」。" });
                    return;
                }
                galleryAction("vectors", { ids: [selected], text: true, image: true, allowUpload: true }, {}).then((body) => {
                    if (body !== null) {
                        const text = body.text && body.text.note ? String(body.text.note) : "";
                        const image = body.image && body.image.note ? String(body.image.note) : "";
                        patchState({ galleryNotice: "向量索引已更新。" + (text ? " " + text : "") + (image ? " " + image : "") });
                    }
                    loadGalleryStatus();
                });
            }
            function gallerySearchByImage() {
                const selected = latest.current.gallerySelected;
                if (selected === "") { patchState({ galleryError: "先选一张图，再以它为查询图。" }); return; }
                if (latest.current.galleryAllowUpload !== true) {
                    patchState({ galleryError: "以图搜图会把查询图发给多模态端点（可能计费）：先勾选「允许内容发给模型」。" });
                    return;
                }
                galleryLoadFilter({
                    galleryImageSearch: true, galleryApplied: "", galleryTags: "",
                    galleryCategory: "", galleryFavorite: false, galleryQuery: "", galleryFolder: "",
                });
            }
            function galleryRemove() {
                const selected = latest.current.gallerySelected;
                if (selected === "") return;
                if (latest.current.galleryRemoveConfirm !== true) {
                    patchState({ galleryRemoveConfirm: true, galleryError: "" });
                    return;
                }
                galleryAction("remove", { id: selected }, {}).then((body) => {
                    if (body !== null) {
                        // 移除图片后「已归类 N 张」与各文件夹计数都会变。
                        loadGalleryStatus();
                        patchState({
                            galleryRemoveConfirm: false, gallerySelected: "", galleryDraft: null,
                            galleryPreview: false, galleryNotice: "已从图库移除（文件不再由图库服务）。",
                        });
                    }
                });
            }

            const galleryStatusData = state.galleryStatus && state.galleryStatus.ok === true ? state.galleryStatus : null;
            // 文本向量如果是远端模式，连普通文字搜索都会把搜索词发给那个端点；
            // 这和「图片上传开关」是两件事，状态栏必须分开说，不能一律写「本地」。
            const galleryTextRemote = Boolean(galleryStatusData && galleryStatusData.textEmbedding && galleryStatusData.textEmbedding.remote === true);
            const galleryDataObj = state.galleryData && state.galleryData.ok === true ? state.galleryData : null;
            const galleryAllItems = galleryDataObj && Array.isArray(galleryDataObj.items) ? galleryDataObj.items : [];
            const galleryTotal = galleryDataObj && Number.isFinite(Number(galleryDataObj.total))
                ? Number(galleryDataObj.total) : galleryAllItems.length;
            const galleryRootCount = galleryStatusData && galleryStatusData.counts && Number.isFinite(Number(galleryStatusData.counts.items))
                ? Number(galleryStatusData.counts.items) : galleryTotal;
            const gallerySorted = galleryAllItems.slice().sort((left, right) => galleryCompare(left, right, state.gallerySort, state.gallerySortDir));
            const gallerySelectedItem = gallerySorted.find((item) => String(item && item.id || "") === String(state.gallerySelected || "")) || null;
            const galleryDraft = state.galleryDraft && String(state.galleryDraft.id || "") === String(state.gallerySelected || "")
                ? state.galleryDraft : (gallerySelectedItem === null ? null : galleryDraftOf(gallerySelectedItem));
            const galleryCategories = [];
            for (const item of galleryAllItems) {
                const category = String(item && item.category || "").trim();
                if (category !== "" && galleryCategories.indexOf(category) < 0) galleryCategories.push(category);
            }
            galleryCategories.sort();
            const gallerySearching = state.galleryImageSearch === true || String(state.galleryApplied || "") !== ""
                || String(state.galleryTags || "") !== "" || String(state.galleryCategory || "") !== ""
                || state.galleryFavorite === true;
            // 文件夹来自 status.folders（显式空目录也在里面）；旧响应只有字符串时也能画。
            const galleryFolderRows = (galleryStatusData !== null && Array.isArray(galleryStatusData.folders) ? galleryStatusData.folders : [])
                .map((row) => (typeof row === "string"
                    ? { path: row, name: row.split("/").pop(), parent: "", depth: Math.max(0, row.split("/").length - 1), direct: 0, total: 0 }
                    : row))
                .filter((row) => row && typeof row === "object" && String(row.path || "") !== "");
            const galleryFolderPaths = galleryFolderRows.map((row) => String(row.path));
            // 状态里没有、但图片实际带着的路径（旧数据 / 刚移动）也要能进树；
            // 连同祖先一起补齐，否则子目录会被挂到一个不存在的父节点下面。
            const galleryEnsureAncestors = (path) => {
                const queue = [];
                let at = String(path || "");
                while (at !== "" && galleryFolderPaths.indexOf(at) < 0) {
                    queue.push(at);
                    const cut = at.lastIndexOf("/");
                    at = cut < 0 ? "" : at.slice(0, cut);
                }
                for (const item of queue) galleryFolderPaths.push(item);
            };
            for (const row of galleryFolderRows) galleryEnsureAncestors(String(row.path || ""));
            for (const item of galleryAllItems) galleryEnsureAncestors(String(item && item.folder || "").trim());
            galleryFolderPaths.sort();

            // 左树由路径现算：状态给 path/parent/depth/direct/total，图片补缺失的路径。
            // 折叠状态只活在面板里，不写回 manifest；默认全展开。
            const galleryFolderNodeMap = new Map();
            for (const row of galleryFolderRows) {
                galleryFolderNodeMap.set(String(row.path), Object.assign({}, row, { children: [] }));
            }
            for (const path of galleryFolderPaths) {
                if (!galleryFolderNodeMap.has(path)) {
                    const parts = path.split("/");
                    galleryFolderNodeMap.set(path, {
                        path: path, name: parts[parts.length - 1],
                        parent: parts.length > 1 ? parts.slice(0, -1).join("/") : "",
                        depth: parts.length - 1, direct: 0, total: 0, children: [],
                    });
                }
            }
            for (const node of galleryFolderNodeMap.values()) {
                const parent = galleryFolderNodeMap.get(String(node.parent || ""));
                if (parent !== undefined) parent.children.push(node);
            }
            for (const node of galleryFolderNodeMap.values()) {
                node.children.sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "zh-Hans-CN", { numeric: true, sensitivity: "base" }));
            }
            const galleryFolderNow = String(state.galleryFolder || "");
            const galleryFolderLeaf = galleryFolderNow === "" ? "" : galleryFolderNow.split("/").pop();
            const galleryCurrentFolderNode = galleryFolderNow === "" ? null : (galleryFolderNodeMap.get(galleryFolderNow) || null);
            const galleryRootChildren = Array.from(galleryFolderNodeMap.values())
                .filter((node) => String(node.parent || "") === "")
                .sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "zh-Hans-CN", { numeric: true, sensitivity: "base" }));
            // 当前层的直属子文件夹：详细信息里排在图片前面，大图标里排在图片卡片前面。
            // 根目录没有「当前节点」，它的子文件夹就是顶层文件夹。
            const galleryChildFolders = gallerySearching
                ? []
                : (galleryFolderNow === "" ? galleryRootChildren : (galleryCurrentFolderNode === null ? [] : (galleryCurrentFolderNode.children || []).slice()));
            // 只看某一层（没有搜索条件）时，空态与文案要说「这个文件夹」而不是「图库还是空的」。
            const galleryFolderOnly = !gallerySearching && galleryFolderNow !== "";
            // 信号报告：used 的进「融合」，used:false 的带原因列在下面。缺信号的通道本来
            // 就不进分母，界面也照实说它没参与，而不是把它算成 0 分。
            const galleryUsedSignals = [];
            const gallerySkippedSignals = [];
            const galleryHasSignals = Boolean(galleryDataObj && galleryDataObj.signals && typeof galleryDataObj.signals === "object");
            // 状态按服务端这一次真实返回的 signals 说话：textVector.mode 才是「这次用了
            // local 还是 remote」，不能只看设置里配没配远端。
            const galleryTextSignal = galleryHasSignals && galleryDataObj.signals.textVector && typeof galleryDataObj.signals.textVector === "object"
                ? galleryDataObj.signals.textVector : null;
            const galleryTextSignalMode = galleryTextSignal && typeof galleryTextSignal.mode === "string" ? galleryTextSignal.mode : "";
            const galleryImageTextUsed = Boolean(galleryHasSignals && galleryDataObj.signals.imageText && galleryDataObj.signals.imageText.used === true);
            const galleryTextHintOn = String(state.galleryApplied || "") !== "" && state.galleryImageSearch !== true;
            const galleryTextHint = !galleryTextHintOn
                ? ""
                : galleryTextSignalMode === "remote"
                    ? (galleryImageTextUsed ? "云端检索" : "文本向量：远端")
                    : galleryImageTextUsed
                        ? "图文对齐：云端"
                        : galleryStatusData === null && galleryTextSignalMode === ""
                            ? "检索"
                            : (galleryTextRemote && state.galleryCloudSearch !== true
                                ? "本地检索（远端向量需勾选）"
                                : "本地检索");
            const galleryTextHintTitle = galleryTextSignalMode === "remote"
                ? (galleryImageTextUsed
                    ? "这次搜索词已发给远端文本与图文端点（allowUpload 已允许）。"
                    : "这次搜索词已发给远端文本 embedding 端点（allowUpload 已允许）。")
                : galleryImageTextUsed
                    ? "这次搜索把查询词发给了图文端点（allowUpload 已允许）。"
                    : galleryTextRemote
                        ? "远端文本向量未获允许：这次用本地哈希向量，没有联网。"
                        : "本地词法与哈希向量，没有联网。";
            if (galleryHasSignals) {
                const names = [["lexical", "词法"], ["tag", "标签"], ["textVector", "文本向量"], ["imageText", "图文对齐"], ["imageImage", "以图搜图"]];
                for (const pair of names) {
                    const report = galleryDataObj.signals[pair[0]];
                    if (report === null || report === undefined) continue;
                    if (report.used === true) galleryUsedSignals.push(pair[1]);
                    else gallerySkippedSignals.push(pair[1] + "（" + String(report.reason || "未参与") + "）");
                }
            }
            const galleryTruncated = galleryTotal > galleryAllItems.length;

            const galleryCrumb = h("div", { style: S.kbCrumb, "data-gallery-crumb": String(galleryFolderNow || state.galleryApplied || state.galleryCategory || state.galleryTags || "") },
                h("button", {
                    type: "button", className: "om-btn", style: S.kbCrumbBtn, "data-gallery-crumb-root": "true",
                    onClick: () => galleryLoadFilter({
                        galleryApplied: "", galleryQuery: "", galleryTags: "", galleryCategory: "",
                        galleryFavorite: false, galleryImageSearch: false, galleryFolder: "",
                        galleryFolderRemoveConfirm: false, galleryBack: [],
                        gallerySelected: "", galleryDraft: null, galleryPreview: false,
                    }),
                }, "图库"),
                galleryFolderNow === "" ? null : galleryFolderNow.split("/").map((name, index, all) => {
                    const path = all.slice(0, index + 1).join("/");
                    const isLast = index === all.length - 1;
                    return h(React.Fragment, { key: "folder-" + path },
                        h("span", { style: S.kbCrumbSep }, "/"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbCrumbBtn,
                            "data-gallery-crumb-folder": path, "aria-current": isLast ? "location" : undefined,
                            onClick: () => { if (!isLast) galleryGoFolder(path); },
                        }, name));
                }),
                state.galleryFavorite === true
                    ? h(React.Fragment, { key: "favorite" },
                        h("span", { style: S.kbCrumbSep }, "/"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbCrumbBtn, "data-gallery-crumb-favorite": "true",
                            onClick: () => galleryLoadFilter({ galleryFavorite: false }),
                        }, "收藏"))
                    : null,
                String(state.galleryCategory || "") !== ""
                    ? h(React.Fragment, { key: "category" },
                        h("span", { style: S.kbCrumbSep }, "/"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbCrumbBtn, "data-gallery-crumb-category": String(state.galleryCategory),
                            onClick: () => galleryLoadFilter({ galleryCategory: "" }),
                        }, String(state.galleryCategory)))
                    : null,
                String(state.galleryApplied || "") !== ""
                    ? h(React.Fragment, { key: "query" },
                        h("span", { style: S.kbCrumbSep }, "/"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbCrumbBtn, "data-gallery-crumb-query": String(state.galleryApplied),
                            onClick: () => galleryLoadFilter({ galleryApplied: "", galleryQuery: "" }),
                        }, "搜索：" + String(state.galleryApplied)))
                    : null,
                String(state.galleryTags || "") !== ""
                    ? h(React.Fragment, { key: "tags" },
                        h("span", { style: S.kbCrumbSep }, "/"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbCrumbBtn, "data-gallery-crumb-tags": String(state.galleryTags),
                            onClick: () => galleryLoadFilter({ galleryTags: "" }),
                        }, "标签：" + String(state.galleryTags)))
                    : null,
                state.galleryImageSearch === true
                    ? h(React.Fragment, { key: "image" },
                        h("span", { style: S.kbCrumbSep }, "/"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.kbCrumbBtn, "data-gallery-crumb-image": "true",
                            onClick: () => galleryLoadFilter({ galleryImageSearch: false }),
                        }, "以图搜图"))
                    : null,
            );

            const gallerySortOptions = [{ value: "added", label: "加入时间" }, { value: "name", label: "名称" }, { value: "size", label: "大小" }];
            if (gallerySearching) gallerySortOptions.push({ value: "score", label: "相似度" });
            const galleryBackStack = Array.isArray(state.galleryBack) ? state.galleryBack : [];
            const galleryToolbar = h("div", { style: S.explorerBar, "data-gallery-commandbar": "true", role: "toolbar", "aria-label": "图库命令行" },
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, galleryBackStack.length === 0 ? S.explorerToolDisabled : {}),
                    "data-gallery-nav": "back", disabled: galleryBackStack.length === 0,
                    title: galleryBackStack.length === 0 ? "没有可后退的记录" : "后退到 " + String(galleryBackStack[galleryBackStack.length - 1] || "全部图片"),
                    onClick: () => { if (galleryBackStack.length === 0) return; galleryGoFolder(galleryBackStack[galleryBackStack.length - 1], "back"); },
                }, "← 后退"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, galleryFolderNow === "" ? S.explorerToolDisabled : {}),
                    "data-gallery-nav": "up", disabled: galleryFolderNow === "",
                    title: galleryFolderNow === "" ? "已经在全部图片" : "上一级",
                    onClick: () => { if (galleryFolderNow !== "") galleryGoFolder(galleryFolderParentOf(galleryFolderNow)); },
                }, "↑ 上一级"),
                h("button", {
                    type: "button", className: "om-btn", style: S.explorerTool, "data-gallery-refresh": "true",
                    title: "重新读取图库与能力状态",
                    onClick: () => { loadGalleryStatus(); loadGallery({}); scrollPanelTop(); },
                }, "刷新"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, state.galleryBusy ? S.explorerToolDisabled : {}),
                    "data-gallery-new-folder": "true", "aria-expanded": state.galleryFolderNewOpen === true,
                    title: "在当前文件夹下新建一个空文件夹",
                    onClick: () => patchState({
                        galleryFolderNewOpen: state.galleryFolderNewOpen !== true,
                        galleryFolderName: "",
                        galleryFolderRemoveConfirm: false,
                        galleryError: "",
                    }),
                }, "新建文件夹"),
                h("span", { style: { flex: "1 1 auto" } }),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, state.galleryMode === "list" ? S.explorerToolOn : {}),
                    "data-gallery-view": "list", "aria-pressed": state.galleryMode === "list",
                    title: "详细信息",
                    onClick: () => patchState({ galleryMode: "list" }),
                }, "详细信息"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.explorerTool, state.galleryMode === "grid" ? S.explorerToolOn : {}),
                    "data-gallery-view": "grid", "aria-pressed": state.galleryMode === "grid",
                    title: "大图标",
                    onClick: () => patchState({ galleryMode: "grid" }),
                }, "大图标"),
                h("label", { style: { display: "inline-flex", alignItems: "center", gap: "5px", fontSize: "11px", color: M.caption } },
                    "排序",
                    h("select", {
                        style: S.explorerSort, "data-gallery-sort": "true",
                        value: state.gallerySort,
                        onChange: (event) => patchState({ gallerySort: String(event.target.value || "added") }),
                    }, gallerySortOptions.map((option) => h("option", { key: option.value, value: option.value }, option.label)))),
                h("button", {
                    type: "button", className: "om-btn", style: S.explorerTool,
                    "data-gallery-sort-dir": state.gallerySortDir,
                    title: state.gallerySortDir === "asc" ? "升序（点击改为降序）" : "降序（点击改为升序）",
                    onClick: () => patchState({ gallerySortDir: state.gallerySortDir === "asc" ? "desc" : "asc" }),
                }, state.gallerySortDir === "asc" ? "↑" : "↓"),
                h("label", {
                    style: { display: "inline-flex", alignItems: "center", gap: "6px", marginLeft: "auto", fontSize: "11px", color: M.textSecondary },
                    title: "默认只走本地词法 / 标签 / 本地哈希向量。勾选后，搜索词可能发给文本或图文 embedding 端点，并可能为标签、描述生成远端向量（可能计费）。",
                },
                    h("input", {
                        type: "checkbox", "data-gallery-cloud-search": "true",
                        checked: state.galleryCloudSearch === true,
                        onChange: (event) => {
                            const on = Boolean(event.target.checked);
                            patchState({ galleryCloudSearch: on, galleryError: "" });
                            // 正在按文字检索时立刻换通道重查，用户不必再按一次回车。
                            if (String(latest.current.galleryApplied || "") !== "" && latest.current.galleryImageSearch !== true) {
                                loadGallery({ galleryCloudSearch: on });
                            }
                        },
                    }),
                    "文字搜索走云端（含向量）"),
                h("label", {
                    style: { display: "inline-flex", alignItems: "center", gap: "6px", fontSize: "11px", color: M.textSecondary },
                    title: "勾选后，描述 / 探针 / 向量化（图片与文本）/ 以图搜图才会把内容发给已配置的模型（可能产生云端费用）",
                },
                    h("input", {
                        type: "checkbox", "data-gallery-allow": "true",
                        checked: state.galleryAllowUpload === true,
                        onChange: (event) => patchState({ galleryAllowUpload: Boolean(event.target.checked), galleryError: "" }),
                    }),
                    "允许内容发给模型"),
                h("button", {
                    type: "button", className: "om-btn", style: S.kbIngestBtn, "data-gallery-upload": "true",
                    disabled: Boolean(state.galleryUploading),
                    onClick: galleryPickFiles,
                }, state.galleryUploading ? ("上传中 · " + String(state.galleryUploading).slice(0, 12)) : "＋ 上传图片"),
                h("input", {
                    type: "file", accept: "image/png,image/jpeg,image/gif,image/webp", multiple: true,
                    "data-gallery-upload-input": "true", style: { display: "none" },
                    onChange: (event) => {
                        const files = event && event.target && event.target.files;
                        galleryUploadFiles(files);
                        if (event && event.target) event.target.value = "";
                    },
                }),
            );

            // 左树：和知识库同一套树形读法。根是「全部图片」，下面是真实文件夹；
            // 折叠只影响这一栏，点文件夹换浏览层，当前层才给改名 / 删除。
            const galleryTreeRows = [];
            const galleryWalkTree = (node, depth) => {
                const path = String(node.path || "");
                const children = Array.isArray(node.children) ? node.children : [];
                const active = galleryFolderNow === path;
                const closed = children.length > 0 && state.galleryClosed && state.galleryClosed[path] === true;
                const toggle = children.length === 0 ? null : (event) => {
                    if (event && typeof event.stopPropagation === "function") event.stopPropagation();
                    patchState({ galleryClosed: Object.assign({}, state.galleryClosed || {}, { [path]: !closed }) });
                };
                const direct = Number(node.direct) || 0;
                const total = Number(node.total) || 0;
                const row = h("div", {
                    key: "folder-" + path,
                    style: { display: "flex", alignItems: "center", gap: "2px", paddingLeft: String(Math.max(0, depth - 1) * 12) + "px" },
                    "data-gallery-folder-row": path,
                },
                    h("button", {
                        type: "button", className: "om-btn", title: path + " · 直属 " + String(direct) + " 张 / 含子文件夹 " + String(total) + " 张",
                        style: Object.assign({}, S.kbTreeRow, { flex: "1 1 auto", minWidth: 0 }, active ? S.kbTreeRowOn : {}),
                        "data-gallery-folder": path,
                        "data-gallery-folder-depth": String(Math.max(0, depth - 1)),
                        "aria-pressed": active,
                        onClick: () => galleryGoFolder(path),
                    },
                        h("span", {
                            style: Object.assign({}, S.kbTreeGlyph, { display: "inline-flex", alignItems: "center", justifyContent: "center" }, toggle === null ? {} : { cursor: "pointer" }),
                            "data-gallery-toggle": toggle === null ? undefined : (closed ? "closed" : "open"),
                            "data-gallery-toggle-folder": toggle === null ? undefined : path,
                            onClick: toggle || undefined,
                        }, toggle === null ? h(ExplorerGlyph, { kind: "folder", size: 13 }) : (closed ? "▸" : "▾")),
                        h("span", { style: S.kbTreeName }, String(node.name || galleryFolderLeafOf(path))),
                        h("span", { style: S.kbTreeCount }, total === direct ? String(total) : String(direct) + "/" + String(total))),
                    active ? h("button", {
                        type: "button", className: "om-btn", style: Object.assign({}, S.ghostButton, { padding: "2px 6px", fontSize: "10px" }),
                        "data-gallery-folder-rename": "true",
                        onClick: () => patchState({
                            galleryFolderEdit: path,
                            galleryFolderEditValue: galleryFolderLeafOf(path),
                            galleryFolderRemoveConfirm: false,
                            galleryError: "",
                        }),
                    }, "改名") : null,
                    active ? h("button", {
                        type: "button", className: "om-btn",
                        style: Object.assign({}, S.ghostButton, { padding: "2px 6px", fontSize: "10px", color: state.galleryFolderRemoveConfirm === true ? M.error : undefined }),
                        "data-gallery-folder-remove": "true",
                        title: state.galleryFolderRemoveConfirm === true ? "再点一次确认删除这个空文件夹" : "删除这个文件夹；里面还有图时会被拒绝",
                        onClick: galleryRemoveFolder,
                    }, state.galleryFolderRemoveConfirm === true ? "确认删" : "删除") : null);
                if (!active || state.galleryFolderEdit !== path) {
                    galleryTreeRows.push(row);
                } else {
                    galleryTreeRows.push(h(React.Fragment, { key: "folder-edit-" + path },
                        row,
                        h("div", { style: { display: "flex", gap: "4px", padding: "2px 4px 5px " + String(8 + Math.max(0, depth - 1) * 12) + "px" } },
                            h("input", {
                                type: "text", className: "om-input", "data-gallery-folder-rename-input": "true",
                                value: state.galleryFolderEditValue,
                                onChange: (event) => patchState({ galleryFolderEditValue: String(event.target.value || "") }),
                                onKeyDown: (event) => { if (event.key === "Enter") galleryRenameFolder(); },
                                style: Object.assign({}, S.input, { textAlign: "left", minWidth: 0, flex: "1 1 auto" }),
                            }),
                            h("button", {
                                type: "button", className: "om-btn om-primary", style: S.memPrimary,
                                "data-gallery-folder-rename-submit": "true", onClick: galleryRenameFolder,
                            }, "保存"),
                            h("button", {
                                type: "button", className: "om-btn", style: S.ghostButton,
                                "data-gallery-folder-rename-cancel": "true",
                                onClick: () => patchState({ galleryFolderEdit: "", galleryFolderEditValue: "" }),
                            }, "取消"))));
                }
                if (!closed) for (const child of children) galleryWalkTree(child, depth + 1);
            };
            for (const child of galleryRootChildren) galleryWalkTree(child, 1);
            const galleryAllActive = !gallerySearching && galleryFolderNow === "";
            const gallerySide = h("div", { className: "om-gallery-side", style: S.gallerySide, "data-gallery-sidebar": "true", role: "navigation", "aria-label": "图库文件夹" },
                h("div", { style: S.kbTreeHead }, "图库 · 文件夹"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.kbTreeRow, galleryAllActive ? S.kbTreeRowOn : {}),
                    "data-gallery-nav": "all", "aria-pressed": galleryAllActive,
                    onClick: () => galleryLoadFilter({
                        galleryApplied: "", galleryQuery: "", galleryTags: "", galleryCategory: "",
                        galleryFavorite: false, galleryImageSearch: false, galleryFolder: "",
                        galleryFolderRemoveConfirm: false, galleryBack: [],
                        gallerySelected: "", galleryDraft: null, galleryPreview: false,
                    }),
                },
                    h("span", { style: Object.assign({}, S.kbTreeGlyph, { display: "inline-flex", alignItems: "center", justifyContent: "center" }) },
                        h(ExplorerGlyph, { kind: "folder", size: 13 })),
                    h("span", { style: S.kbTreeName }, "全部图片"),
                    h("span", { style: S.kbTreeCount }, String(galleryRootCount))),
                galleryRootChildren.length === 0 && gallerySearching
                    ? null
                    : galleryRootChildren.length === 0
                        ? h("div", { style: S.kbTreeEmpty }, "还没有文件夹")
                        : galleryTreeRows,
                h("div", { style: Object.assign({}, S.kbTreeHead, { paddingTop: "10px" }) }, "筛选"),
                h("button", {
                    type: "button", className: "om-btn",
                    style: Object.assign({}, S.kbTreeRow, state.galleryFavorite === true ? S.kbTreeRowOn : {}),
                    "data-gallery-nav": "favorite", "aria-pressed": state.galleryFavorite === true,
                    onClick: () => galleryLoadFilter({
                        galleryFavorite: true, galleryApplied: "", galleryQuery: "", galleryTags: "",
                        galleryCategory: "", galleryImageSearch: false, galleryFolder: "",
                        galleryFolderRemoveConfirm: false, galleryBack: [],
                        gallerySelected: "", galleryDraft: null, galleryPreview: false,
                    }),
                },
                    h("span", { style: Object.assign({}, S.kbTreeGlyph, { color: "#c08a2e" }) }, "★"),
                    h("span", { style: S.kbTreeName }, "收藏")),
                h("div", { style: Object.assign({}, S.kbTreeHead, { paddingTop: "10px" }) }, "分类"),
                galleryCategories.length === 0
                    ? h("div", { style: S.kbTreeEmpty }, "还没有分类")
                    : galleryCategories.map((category) => h("button", {
                        key: category, type: "button", className: "om-btn",
                        style: Object.assign({}, S.kbTreeRow, String(state.galleryCategory || "") === category ? S.kbTreeRowOn : {}),
                        "data-gallery-category": category, "aria-pressed": String(state.galleryCategory || "") === category,
                        onClick: () => galleryLoadFilter({
                            galleryCategory: category, galleryApplied: "", galleryQuery: "", galleryTags: "",
                            galleryFavorite: false, galleryImageSearch: false, galleryFolder: "",
                            galleryFolderRemoveConfirm: false, galleryBack: [],
                            gallerySelected: "", galleryDraft: null, galleryPreview: false,
                        }),
                    }, category)),
                h("div", { style: { flex: "1 1 auto" } }),
                h("div", { style: { padding: "6px 8px 2px", fontSize: "10.5px", color: M.caption, lineHeight: "1.5" } },
                    String(galleryTotal) + " 张 · 文件夹 " + String(galleryFolderRows.length)
                    + " · 已描述 " + String(galleryStatusData && galleryStatusData.counts ? galleryStatusData.counts.withAesthetic : "—")),
            );

            const galleryCard = (item) => {
                const id = String(item.id || "");
                const on = id === String(state.gallerySelected || "");
                return h("button", {
                    key: id, type: "button", "data-gallery-item": id, "aria-pressed": on,
                    title: id + " · " + String(item.mediaType || ""),
                    style: Object.assign({}, S.galleryCard, on ? S.galleryCardOn : {}),
                    onClick: () => gallerySelectItem(id, item),
                    onDoubleClick: () => { gallerySelectItem(id, item); patchState({ galleryPreview: true }); },
                    // 回车 = 打开预览（Win11 资源管理器里回车就是打开）。preventDefault
                    // 是为了压掉 <button> 随后合成的那次 click —— 那会把刚打开的预览又重置掉。
                    onKeyDown: (event) => {
                        if (!event || event.key !== "Enter") return;
                        event.preventDefault();
                        gallerySelectItem(id, item);
                        patchState({ galleryPreview: true });
                    },
                },
                    item.thumb
                        ? h("img", {
                            src: galleryImageHref(id, "thumb"), alt: "图库缩略图 " + galleryShortId(id),
                            loading: "lazy", style: S.galleryThumb,
                        })
                        : h("span", { style: Object.assign({}, S.galleryThumb, { display: "flex", alignItems: "center", justifyContent: "center", color: M.caption, fontSize: "11px" }) }, "无缩略图"),
                    h("span", { style: S.galleryCardBody },
                        h("span", { style: { display: "flex", alignItems: "center", gap: "5px", minWidth: 0 } },
                            h("span", { style: { fontFamily: M.mono, fontSize: "11px", color: M.text, overflow: "hidden", textOverflow: "ellipsis" } }, String(item.name || galleryShortId(id))),
                            item.favorite === true ? h("span", { style: { color: "#c08a2e", fontSize: "11px" }, title: "已收藏" }, "★") : null),
                        h("span", { style: { fontFamily: M.mono, fontSize: "9.5px", color: M.caption } },
                            (item.width && item.height ? String(item.width) + "×" + String(item.height) : "尺寸未知") + " · " + String(humanBytes(item.bytes) || "—")),
                        Array.isArray(item.tags) && item.tags.length > 0
                            ? h("span", { style: { fontSize: "10px", color: M.textSecondary, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, item.tags.slice(0, 3).join(" · "))
                            : null,
                        String(item.folder || "") !== ""
                            ? h("span", { style: { fontSize: "10px", color: M.caption, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, "文件夹：" + String(item.folder))
                            : null,
                        item.aesthetic
                            ? h("span", { style: { fontSize: "10px", color: M.caption, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, String(item.aesthetic.summary || item.aesthetic.subject || "有视觉描述"))
                            : null));
            };

            const galleryMediaLabel = (item) => {
                const mime = String(item && item.mediaType || "");
                const at = mime.indexOf("/");
                return at < 0 ? (mime || "—") : mime.slice(at + 1).toUpperCase();
            };
            const galleryItemDescriptionOf = (item) => String(item && (item.description || (item.aesthetic && item.aesthetic.summary)) || "").trim();
            const gallerySortHead = (label, key) => h("button", {
                type: "button", className: "om-btn",
                style: Object.assign({}, S.kbExplorerMeta, { textAlign: "left", background: "transparent", border: "none", padding: 0, cursor: "pointer" }),
                "data-gallery-sort-head": key,
                "aria-pressed": state.gallerySort === key,
                onClick: () => patchState({ gallerySort: key, gallerySortDir: state.gallerySort === key && state.gallerySortDir === "asc" ? "desc" : "asc" }),
            }, label + (state.gallerySort === key ? (state.gallerySortDir === "asc" ? " ↑" : " ↓") : ""));
            const galleryExplorerHead = h("div", {
                style: Object.assign({}, S.kbExplorerHead, { gridTemplateColumns: GALLERY_EXPLORER_COLUMNS }),
                "data-gallery-file-head": "true",
            },
                h("span", null, ""),
                gallerySortHead("名称", "name"),
                h("span", { style: S.kbExplorerMeta }, "类型"),
                h("span", { style: S.kbExplorerMeta }, "数量"),
                gallerySortHead("大小", "size"),
                gallerySortHead("加入时间", "added"));

            const galleryFolderRow = (node) => {
                const path = String(node.path || "");
                const name = String(node.name || galleryFolderLeafOf(path));
                const total = Number(node.total) || 0;
                return h("div", {
                    key: "folder-row:" + path,
                    style: Object.assign({}, S.kbExplorerRow, {
                        gridTemplateColumns: GALLERY_EXPLORER_COLUMNS,
                        cursor: "pointer", background: "transparent",
                    }),
                    "data-gallery-subfolder": path,
                    tabIndex: 0,
                    role: "button",
                    title: "进入文件夹 " + path,
                    onClick: () => galleryGoFolder(path),
                    onKeyDown: (event) => {
                        if (!event || event.key !== "Enter") return;
                        event.preventDefault();
                        galleryGoFolder(path);
                    },
                },
                    h("span", { style: { width: "46px", display: "inline-flex", justifyContent: "center", color: M.warn } }, h(ExplorerGlyph, { kind: "folder", size: 22 })),
                    h("span", { style: Object.assign({}, S.kbExplorerName, { fontWeight: 500 }) }, name),
                    h("span", { style: S.kbExplorerMeta }, "文件夹"),
                    h("span", { style: S.kbExplorerMeta }, total === 0 ? "空" : String(total) + " 张"),
                    h("span", { style: S.kbExplorerMeta }, "—"),
                    h("span", { style: S.kbExplorerMeta }, "—"));
            };
            const galleryFolderTile = (node) => {
                const path = String(node.path || "");
                const total = Number(node.total) || 0;
                return h("button", {
                    key: "folder-tile:" + path, type: "button", className: "om-btn",
                    style: S.explorerTile,
                    "data-gallery-subfolder": path,
                    title: "进入文件夹 " + path,
                    onClick: () => galleryGoFolder(path),
                },
                    h("span", { style: Object.assign({}, S.explorerTileIcon, { color: M.warn }) }, h(ExplorerGlyph, { kind: "folder", size: 40 })),
                    h("span", { style: S.explorerTileName }, String(node.name || galleryFolderLeafOf(path))),
                    h("span", { style: S.explorerTileMeta }, total === 0 ? "空文件夹" : String(total) + " 张"));
            };

            const galleryListItem = (item) => {
                const id = String(item.id || "");
                const on = id === String(state.gallerySelected || "");
                return h("button", {
                    key: id, type: "button", "data-gallery-item": id, "aria-pressed": on,
                    style: Object.assign({}, S.galleryRow, on ? { background: M.sunken, boxShadow: "inset 2px 0 0 " + M.accent } : {}),
                    onClick: () => gallerySelectItem(id, item),
                    onDoubleClick: () => { gallerySelectItem(id, item); patchState({ galleryPreview: true }); },
                    onKeyDown: (event) => {
                        if (!event || event.key !== "Enter") return;
                        event.preventDefault();
                        gallerySelectItem(id, item);
                        patchState({ galleryPreview: true });
                    },
                },
                    item.thumb
                        ? h("img", { src: galleryImageHref(id, "thumb"), alt: "", loading: "lazy", style: S.galleryRowThumb })
                        : h("span", { style: Object.assign({}, S.galleryRowThumb, { display: "flex", alignItems: "center", justifyContent: "center", fontSize: "9px", color: M.caption }) }, "—"),
                    h("span", { style: { display: "flex", flexDirection: "column", gap: "1px", minWidth: 0 } },
                        h("span", { style: { display: "flex", alignItems: "center", gap: "5px", minWidth: 0 } },
                            h("span", { style: { fontFamily: M.mono, fontSize: "11.5px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, String(item.name || galleryShortId(id))),
                            item.favorite === true ? h("span", { style: { color: "#c08a2e", fontSize: "11px" }, title: "已收藏" }, "★") : null),
                        h("span", { style: { fontSize: "10px", color: M.caption, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } },
                            (item.width && item.height ? String(item.width) + "×" + String(item.height) + " · " : "")
                            + (Array.isArray(item.tags) && item.tags.length > 0 ? item.tags.join(" · ") : "无标签")
                            + (String(item.category || "") !== "" ? " · " + String(item.category) : "")
                            + (galleryItemDescriptionOf(item) !== "" ? " · " + galleryItemDescriptionOf(item) : ""))),
                    h("span", { style: S.galleryRowMeta }, galleryMediaLabel(item)),
                    h("span", { style: S.galleryRowMeta }, "—"),
                    h("span", { style: S.galleryRowMeta }, String(humanBytes(item.bytes) || "—")),
                    h("span", { style: S.galleryRowMeta }, galleryWhen(item.addedAt)));
            };

            const gallerySubfolderRows = galleryChildFolders.map(galleryFolderRow);
            const gallerySubfolderTiles = galleryChildFolders.map(galleryFolderTile);
            const galleryEmpty = h("div", { style: S.emptyState, "data-gallery-empty": "true" },
                h("div", { style: S.emptyGlyph }, "▧"),
                h("div", { style: S.emptyTitle }, gallerySearching
                    ? "没有匹配的图片。"
                    : galleryFolderOnly ? "这个文件夹还没有图片。" : "图库还是空的。"),
                h("div", { style: S.emptyHint }, gallerySearching
                    ? "清掉筛选或换一个检索词；以图搜图要选一张图并允许内容发给模型。"
                    : galleryFolderOnly
                        ? "把图片移到这里，或点上方「＋ 上传图片」；左侧「全部图片」回到全部。"
                        : "点工具条「＋ 上传图片」；上传只保存到图库目录。"),
                h("div", { style: { display: "flex", justifyContent: "center", gap: "8px", marginTop: "10px" } },
                    gallerySearching
                        ? h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton,
                            "data-gallery-empty-action": "clear",
                            onClick: () => galleryLoadFilter({
                                galleryApplied: "", galleryQuery: "", galleryTags: "", galleryCategory: "",
                                galleryFavorite: false, galleryImageSearch: false,
                                gallerySelected: "", galleryDraft: null, galleryPreview: false,
                            }),
                        }, "清空搜索")
                        : h("button", {
                            type: "button", className: "om-btn", style: S.kbIngestBtn,
                            "data-gallery-empty-action": "upload", onClick: galleryPickFiles,
                        }, "＋ 上传图片")));
            const galleryContent = state.galleryLoading && galleryDataObj === null
                ? h("div", { style: S.loadingState }, h("span", { className: "om-spin", style: S.spinner }), "正在读取图库…")
                : gallerySorted.length === 0 && galleryChildFolders.length === 0
                    ? galleryEmpty
                    : state.galleryMode === "list"
                        ? h("div", { style: S.kbExplorerList, "data-gallery-list": "true", "data-gallery-view": "details" },
                            galleryExplorerHead,
                            gallerySubfolderRows,
                            gallerySorted.map(galleryListItem),
                            gallerySorted.length === 0
                                ? h("div", { style: { padding: "10px 12px", fontSize: "11.5px", color: M.caption } }, "这个文件夹还没有直属图片；进入子文件夹，或点上方「＋ 上传图片」。")
                                : null)
                        : h("div", { style: S.galleryGrid, "data-gallery-grid": "true" },
                            gallerySubfolderTiles,
                            gallerySorted.map(galleryCard));

            const galleryAesthetic = gallerySelectedItem && gallerySelectedItem.aesthetic ? gallerySelectedItem.aesthetic : null;
            const galleryDetail = gallerySelectedItem === null ? null : h("div", { style: S.galleryDetail, "data-gallery-detail-shown": String(state.gallerySelected || "") },
                h("div", { style: { flex: "0 0 300px", minWidth: 0, display: "flex", flexDirection: "column", gap: "7px" } },
                    h("a", {
                        href: galleryImageHref(state.gallerySelected, "image"), target: "_blank", rel: "noreferrer",
                        style: { display: "block" }, title: "在新标签打开原图",
                    },
                        h("img", {
                            src: galleryImageHref(state.gallerySelected, "image"), alt: "原图预览 " + galleryShortId(state.gallerySelected),
                            "data-gallery-original": "true", style: S.galleryPreview,
                        })),
                    h("div", { style: { display: "flex", gap: "8px", flexWrap: "wrap" } },
                        h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton, "data-gallery-preview": "true",
                            onClick: () => patchState({ galleryPreview: true }),
                        }, "查看原图"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton, "data-gallery-search-image": "true",
                            disabled: galleryStatusData !== null && galleryStatusData.imageEmbedding
                                && galleryStatusData.imageEmbedding.configured !== true,
                            title: galleryStatusData && galleryStatusData.imageEmbedding ? String(galleryStatusData.imageEmbedding.note || "") : "读取能力状态后可用",
                            onClick: gallerySearchByImage,
                        }, "以图搜图"))),
                h("div", { style: { flex: "1 1 260px", minWidth: "240px", display: "flex", flexDirection: "column", gap: "7px" } },
                    h("div", { style: { fontFamily: M.mono, fontSize: "11px", color: M.text, overflowWrap: "anywhere" } }, String(state.gallerySelected)),
                    h("label", { style: { display: "flex", flexDirection: "column", gap: "3px", fontSize: "11px", color: M.caption } },
                        "名称",
                        h("input", {
                            type: "text", className: "om-input", "data-gallery-name": "true",
                            value: galleryDraft === null ? "" : galleryDraft.name,
                            placeholder: "可选；上传时默认用文件名",
                            onChange: (event) => patchState({ galleryDraft: Object.assign({}, galleryDraft || galleryDraftOf(gallerySelectedItem), { name: String(event.target.value || "") }) }),
                            style: Object.assign({}, S.input, { textAlign: "left" }),
                        })),
                    h("div", { style: { fontSize: "11px", color: M.caption } },
                        (gallerySelectedItem.width && gallerySelectedItem.height ? String(gallerySelectedItem.width) + "×" + String(gallerySelectedItem.height) : "尺寸未知")
                        + " · " + String(humanBytes(gallerySelectedItem.bytes) || "—")
                        + " · " + String(gallerySelectedItem.mediaType || "")
                        + " · " + galleryWhen(gallerySelectedItem.addedAt)),
                    h("label", { style: { display: "flex", flexDirection: "column", gap: "3px", fontSize: "11px", color: M.caption } },
                        "标签（逗号分隔）",
                        h("input", {
                            type: "text", className: "om-input", "data-gallery-tags": "true",
                            value: galleryDraft === null ? "" : galleryDraft.tags,
                            placeholder: "例：界面, 深色, 截图",
                            onChange: (event) => patchState({ galleryDraft: Object.assign({}, galleryDraft || galleryDraftOf(gallerySelectedItem), { tags: String(event.target.value || "") }) }),
                            style: Object.assign({}, S.input, { textAlign: "left" }),
                        })),
                    h("label", { style: { display: "flex", flexDirection: "column", gap: "3px", fontSize: "11px", color: M.caption } },
                        "文件夹",
                        h("input", {
                            type: "text", className: "om-input", "data-gallery-folder-input": "true",
                            value: galleryDraft === null ? "" : galleryDraft.folder,
                            placeholder: "例：界面/深色；留空 = 未归类",
                            list: "om-gallery-folders",
                            onChange: (event) => patchState({ galleryDraft: Object.assign({}, galleryDraft || galleryDraftOf(gallerySelectedItem), { folder: String(event.target.value || "") }) }),
                            style: Object.assign({}, S.input, { textAlign: "left" }),
                        }),
                        h("datalist", { id: "om-gallery-folders" }, galleryFolderPaths.map((path) => h("option", { key: path, value: path })))),
                    h("div", { style: { display: "flex", gap: "8px", flexWrap: "wrap" } },
                        h("label", { style: { display: "flex", flexDirection: "column", gap: "3px", fontSize: "11px", color: M.caption, flex: "1 1 120px" } },
                            "分类",
                            h("input", {
                                type: "text", className: "om-input", "data-gallery-category-input": "true",
                                value: galleryDraft === null ? "" : galleryDraft.category,
                                placeholder: "例：界面参考",
                                list: "om-gallery-categories",
                                onChange: (event) => patchState({ galleryDraft: Object.assign({}, galleryDraft || galleryDraftOf(gallerySelectedItem), { category: String(event.target.value || "") }) }),
                                style: Object.assign({}, S.input, { textAlign: "left" }),
                            }),
                            h("datalist", { id: "om-gallery-categories" }, galleryCategories.map((category) => h("option", { key: category, value: category })))),
                        h("label", { style: { display: "inline-flex", alignItems: "center", gap: "6px", fontSize: "11px", color: M.textSecondary, alignSelf: "flex-end", paddingBottom: "7px" } },
                            h("input", {
                                type: "checkbox", "data-gallery-favorite": "true",
                                checked: galleryDraft !== null && galleryDraft.favorite === true,
                                onChange: (event) => patchState({ galleryDraft: Object.assign({}, galleryDraft || galleryDraftOf(gallerySelectedItem), { favorite: Boolean(event.target.checked) }) }),
                            }),
                            "收藏")),
                    h("label", { style: { display: "flex", flexDirection: "column", gap: "3px", fontSize: "11px", color: M.caption } },
                        "描述",
                        h("textarea", {
                            className: "om-input", "data-gallery-description": "true", rows: 4,
                            value: galleryDraft === null ? "" : galleryDraft.description,
                            placeholder: "可编辑；留空则显示视觉模型生成的总结",
                            onChange: (event) => patchState({ galleryDraft: Object.assign({}, galleryDraft || galleryDraftOf(gallerySelectedItem), { description: String(event.target.value || "") }) }),
                            style: Object.assign({}, S.editorTextarea, { minHeight: "76px" }),
                        })),
                    h("label", { style: { display: "flex", flexDirection: "column", gap: "3px", fontSize: "11px", color: M.caption } },
                        "细节",
                        h("textarea", {
                            className: "om-input", "data-gallery-note": "true", rows: 2,
                            value: galleryDraft === null ? "" : galleryDraft.note,
                            onChange: (event) => patchState({ galleryDraft: Object.assign({}, galleryDraft || galleryDraftOf(gallerySelectedItem), { note: String(event.target.value || "") }) }),
                            style: Object.assign({}, S.editorTextarea, { minHeight: "48px" }),
                        })),
                    h("div", { style: { display: "flex", gap: "6px", flexWrap: "wrap" } },
                        h("button", {
                            type: "button", className: "om-btn om-primary", style: S.memPrimary,
                            "data-gallery-save": "true", disabled: state.galleryBusy === true,
                            onClick: gallerySaveDraft,
                        }, state.galleryBusy === true ? "保存中…" : "保存细节"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton,
                            "data-gallery-describe": "true", disabled: state.galleryBusy === true,
                            onClick: galleryDescribeRequest,
                        }, "生成视觉描述"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton,
                            "data-gallery-vectors": "true", disabled: state.galleryBusy === true,
                            onClick: galleryVectors,
                        }, "建向量索引"),
                        h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton,
                            "data-gallery-probe": "true", disabled: state.galleryBusy === true,
                            onClick: galleryProbe,
                        }, "探针图文对齐"),
                        h("button", {
                            type: "button", className: "om-btn",
                            style: Object.assign({}, S.ghostButton, { color: M.error, borderColor: M.error }),
                            "data-gallery-remove": state.galleryRemoveConfirm === true ? "confirm" : "delete",
                            disabled: state.galleryBusy === true,
                            onClick: galleryRemove,
                        }, state.galleryRemoveConfirm === true ? "再点一次确认移除" : "移出图库")),
                    state.galleryDescribeConfirm === true
                        ? h("div", { style: { padding: "7px 9px", borderRadius: "8px", border: "1px solid " + M.warn, background: "color-mix(in srgb, " + M.warn + " 8%, transparent)", fontSize: "11.5px", color: M.textSecondary } },
                            "这张原图会发给已配置的视觉模型（可能产生云端费用）。",
                            h("div", { style: { display: "flex", gap: "6px", marginTop: "6px" } },
                                h("button", {
                                    type: "button", className: "om-btn om-primary", style: S.memPrimary,
                                    "data-gallery-describe-confirm": "true", onClick: galleryDescribeConfirm,
                                }, "确认发送并生成"),
                                h("button", {
                                    type: "button", className: "om-btn", style: S.ghostButton,
                                    "data-gallery-describe-cancel": "true",
                                    onClick: () => patchState({ galleryDescribeConfirm: false }),
                                }, "取消")))
                        : null,
                    state.galleryRemoveConfirm === true
                        ? h("div", { style: { fontSize: "11px", color: M.error } }, "移除会从图库清单与图库目录里删掉这张图，面板里无法撤回。")
                        : null,
                    galleryAesthetic === null
                        ? h("div", { style: { fontSize: "11px", color: M.caption }, "data-gallery-aesthetic": "none" }, "还没有视觉描述。")
                        : h("div", { style: { display: "flex", flexDirection: "column", gap: "3px", padding: "8px 9px", borderRadius: "8px", border: "1px solid " + M.stroke, background: M.sunken }, "data-gallery-aesthetic": "true" },
                            h("div", { style: { fontSize: "10.5px", color: M.caption } }, "视觉描述来源（描述已放进上面的编辑框）"),
                            galleryAesthetic.tags && galleryAesthetic.tags.length > 0
                                ? h("div", { style: { fontSize: "10.5px", color: M.textSecondary } }, "描述标签：" + galleryAesthetic.tags.join(" · "))
                                : null,
                            h("div", { style: { fontSize: "10px", color: M.caption } },
                                [galleryAesthetic.provider, galleryAesthetic.model, galleryAesthetic.channel].filter(Boolean).join(" / ")
                                + (galleryAesthetic.generatedAt ? " · " + galleryWhen(galleryAesthetic.generatedAt) : "")),
                            Array.isArray(galleryAesthetic.droppedFields) && galleryAesthetic.droppedFields.length > 0
                                ? h("div", { style: { fontSize: "10px", color: M.warn } }, "模型漏掉的字段：" + galleryAesthetic.droppedFields.join("、"))
                                : null)));

            const galleryCaps = [];
            if (galleryStatusData !== null) {
                const thumb = galleryStatusData.thumbnail || {};
                const lexical = galleryStatusData.lexical || {};
                const textEmbedding = galleryStatusData.textEmbedding || {};
                const imageEmbedding = galleryStatusData.imageEmbedding || {};
                const vision = galleryStatusData.vision || {};
                galleryCaps.push({ key: "thumbnail", label: "缩略图", ok: thumb.configured === true, note: thumb.configured === true ? "本地生成，不联网。" : String(thumb.lastError || "服务端没有可用的缩略图器。") });
                galleryCaps.push({ key: "lexical", label: "词法", ok: lexical.available === true, note: String(lexical.note || "") });
                // 远端文本向量的联网条件由服务端 note 说清（要显式 allowUpload）；
                // 这里只照抄，不再自己补一句「搜索会发出去」——未勾选时它其实会退回本地。
                galleryCaps.push({ key: "text-vector", label: "文本向量", ok: true, note: String(textEmbedding.note || "") });
                galleryCaps.push({
                    key: "image-embedding", label: "图文对齐",
                    ok: imageEmbedding.configured === true,
                    note: String(imageEmbedding.note || ""),
                });
                galleryCaps.push({
                    key: "vision", label: "视觉描述",
                    ok: vision.configured === true,
                    note: vision.configured === true
                        ? "已配置；仍要勾选允许并确认，图片才会上传。"
                        : String(vision.note || "未配置视觉模型。") + " 到「设置 → 办公与记忆 → 图库 → 视觉描述」配置。",
                });
            }
            const galleryCapPanel = galleryStatusData === null ? null : h("div", {
                style: Object.assign({}, S.editor, { marginTop: "10px", padding: "9px 11px" }), "data-gallery-capabilities": "true",
            },
                h("div", { style: S.editorTitle }, "能力与联网边界"),
                galleryCaps.map((cap) => h("div", { key: cap.key, style: S.galleryCap, "data-gallery-capability": cap.key },
                    h("span", { style: S.galleryCapLabel }, cap.label),
                    h("span", { style: { color: cap.ok ? M.textSecondary : M.caption, lineHeight: "1.5" } },
                        cap.ok ? "可用 · " : "不可用 · ", cap.note))),
                h("div", { style: { marginTop: "6px", fontSize: "10.5px", color: M.caption, lineHeight: "1.5" } },
                    String(galleryStatusData.uploadPolicy || "默认不联网；描述 / 探针 / 图片向量化 / 以图搜图都需显式允许。")),
                h("div", { style: { marginTop: "3px", fontSize: "10.5px", color: M.caption, lineHeight: "1.5" } },
                    "默认本地搜索；启用云端后，搜索词及用于向量化的标签、描述可能发送给供应商，并可能计费。"),
                h("div", { style: { marginTop: "3px", fontSize: "10px", color: M.caption, overflowWrap: "anywhere" } },
                    "图库目录：" + String(galleryStatusData.root || "—")));

            const galleryNewFolderBar = state.galleryFolderNewOpen !== true ? null : h("div", {
                style: Object.assign({}, S.kbSelectBar, { borderStyle: "solid", marginBottom: "10px" }),
                "data-gallery-new-folder-bar": "true",
            },
                h("span", { style: { color: M.caption, fontSize: "12px" } },
                    "在「" + (galleryFolderNow === "" ? "全部图片" : galleryFolderNow) + "」下新建文件夹"),
                h("input", {
                    type: "text", className: "om-input", "data-gallery-new-folder-input": "true",
                    value: String(state.galleryFolderName || ""), placeholder: "文件夹名（一次一层）",
                    autoFocus: true,
                    onChange: (event) => patchState({ galleryFolderName: String(event.target.value || ""), galleryError: "" }),
                    onKeyDown: (event) => {
                        if (event.key === "Enter") {
                            event.preventDefault();
                            galleryCreateFolder();
                        } else if (event.key === "Escape") {
                            event.preventDefault();
                            patchState({ galleryFolderNewOpen: false, galleryFolderName: "", galleryError: "" });
                        }
                    },
                }),
                h("button", {
                    type: "button", className: "om-btn", style: S.kbIngestBtn,
                    "data-gallery-new-folder-submit": "true", onClick: galleryCreateFolder,
                }, "新建"),
                h("button", {
                    type: "button", className: "om-btn", style: S.explorerTool,
                    "data-gallery-new-folder-cancel": "true",
                    onClick: () => patchState({ galleryFolderNewOpen: false, galleryFolderName: "", galleryError: "" }),
                }, "取消"));

            const gallerySelectedLabel = String(state.gallerySelected || "") === ""
                ? "未选中"
                : "已选 · " + String(gallerySelectedItem && (gallerySelectedItem.name || galleryShortId(gallerySelectedItem.id)) || galleryShortId(state.gallerySelected));
            const galleryStatusLine = h("div", { style: S.explorerStatus, "data-gallery-status": "true" },
                h("span", { "data-gallery-status-count": "true" },
                    "文件夹 " + String(galleryChildFolders.length) + " · 图片 " + String(galleryAllItems.length) + " / " + String(galleryTotal)),
                h("span", { style: { color: M.textSecondary } }, gallerySelectedLabel),
                galleryFolderNow !== ""
                    ? h("span", { style: { color: M.textSecondary }, "data-gallery-status-folder": galleryFolderNow }, "位置：" + galleryFolderNow)
                    : null,
                galleryTruncated ? h("span", { style: { color: M.warn } }, "清单已截断，仅显示这一页") : null,
                galleryHasSignals
                    ? h("span", { title: gallerySkippedSignals.join("；") },
                        "融合信号：" + (galleryUsedSignals.length === 0 ? "无" : galleryUsedSignals.join("、")))
                    : h("span", null, gallerySearching ? "筛选（标签 / 分类 / 收藏）" : "浏览全部"),
                gallerySkippedSignals.length > 0
                    ? h("span", { style: { color: M.caption, maxWidth: "100%", overflow: "hidden", textOverflow: "ellipsis" }, title: gallerySkippedSignals.join("；") },
                        "跳过：" + gallerySkippedSignals[0] + (gallerySkippedSignals.length > 1 ? " 等 " + String(gallerySkippedSignals.length) + " 项" : ""))
                    : null,
                galleryTextHintOn
                    ? h("span", { style: { color: M.caption }, title: galleryTextHintTitle, "data-gallery-text-hint": galleryTextSignalMode || "local" }, galleryTextHint)
                    : null,
                h("span", { style: { marginLeft: "auto" } }, state.galleryMode === "list" ? "详细信息" : "大图标"));

            const galleryOverlay = state.galleryPreview !== true || gallerySelectedItem === null ? null : h("div", {
                style: S.galleryOverlay, "data-gallery-overlay": "true",
                role: "dialog", "aria-modal": "true", "aria-label": "原图预览",
                onKeyDown: (event) => { if (event.key === "Escape") patchState({ galleryPreview: false }); },
                onClick: (event) => { if (event.target === event.currentTarget) patchState({ galleryPreview: false }); },
            },
                h("img", {
                    src: galleryImageHref(state.gallerySelected, "image"),
                    alt: "原图 " + galleryShortId(state.gallerySelected),
                    style: S.galleryOverlayImg,
                }),
                h("button", {
                    type: "button", className: "om-btn", style: Object.assign({}, S.ghostButton, { position: "absolute", top: "16px", right: "16px" }),
                    "data-gallery-preview-close": "true",
                    onClick: () => patchState({ galleryPreview: false }),
                }, "关闭"));

            // 原图预览是不抢焦点的浮层：Esc 挂到文档上，键盘用户不必先点一下浮层；
            // 浮层自己的 onKeyDown 保留，聚焦在里面时同样能关。
            React.useEffect(() => {
                if (state.galleryPreview !== true || typeof document === "undefined") return undefined;
                const onKey = (event) => { if (event.key === "Escape") patchState({ galleryPreview: false }); };
                document.addEventListener("keydown", onKey);
                return () => document.removeEventListener("keydown", onKey);
            }, [state.galleryPreview]);

            const galleryView = h("div", { className: "om-gallery-explorer om-kb-explorer", style: S.kbExplorer, "data-gallery-panel": "true" },
                gallerySide,
                h("div", { style: S.kbMain },
                    galleryToolbar,
                    galleryNewFolderBar,
                    galleryCrumb,
                    state.galleryError ? h("div", { style: S.editorError, "data-gallery-error": "true" }, String(state.galleryError)) : null,
                    state.galleryNotice ? h("div", { style: Object.assign({}, S.editorHint, { marginBottom: "8px" }), "data-gallery-notice": "true" }, String(state.galleryNotice)) : null,
                    galleryContent,
                    galleryStatusLine,
                    galleryDetail,
                    galleryCapPanel),
                galleryOverlay);

            // 图库不依赖记忆快照：它有自己的列表 / 状态端点，所以放在是否 loading 之前
            // 判断，打开页签就能用，不必等记忆首屏。
            const body = state.tab === "gallery"
                ? galleryView
                : state.loading
                ? h("div", { style: S.loadingState },
                    h("span", { className: "om-spin", style: S.spinner }),
                    "正在读取记忆…")
                : state.tab === "kb"
                    ? h("div", { "data-kb-panel": "true" }, kbExplorer)
                    : withArchiveWarnings(items.length === 0
                        ? h("div", { style: S.emptyState },
                            h("div", { style: S.emptyGlyph }, "◌"),
                            h("div", { style: S.emptyTitle }, "这一层还没有内容。"),
                            h("div", { style: S.emptyHint }, layerKey === "ledger"
                                ? "台账在 office_run 写出第一份交付物时自动登记，不用手工补。"
                                : layerKey === "archive"
                                    ? "归档由热记忆下沉与台账滚动产生；这里空就是还没下沉过。"
                                    : "点右上角「记一条」写入个人偏好或项目记忆。"))
                        : h("div", null, contentHead, editorPanel, archiveNote, kbNote, list, moreBar, detailBar));

            // ── 可视化块 ──
            //
            // 只在真的拿到过数据时画：错误页上不摆一排 0，那会让人以为记忆是空的。
            // 每一块自己判断有没有内容（返回 null 就什么都不占），所以这里不做
            // 逐块的条件判断，免得条件与组件里的判断各自漂移。
            //
            // 版式（第八十一轮）：**内容优先**。指标卡 / 容量 / 存储域统一放到列表
            // 之后，列表先出来；否则打开面板先看到一屏统计，热记忆被推到首屏之外。
            const dashboards = [];
            const statsVisuals = [];
            if (hasData) {
                statsVisuals.push(h(MetricTiles, { key: "tiles", counts: counts }));
                const rows = capacityRows(counts, limits);
                if (rows.length > 0) dashboards.push(h(CapacityGauges, { key: "capacity", rows: rows }));
                if (stores.length > 0 || data.kbStore) dashboards.push(h(StoreCards, { key: "stores", stores: stores, kbStore: data.kbStore }));
                if (state.tab === "ledger") {
                    statsVisuals.push(h(TimelineBars, {
                        key: "ledger-timeline", name: "ledger-timeline", title: "台账按天分布",
                        color: "#22a879", buckets: bucketize(ledger, ledgerDay),
                    }));
                }
                if (state.tab === "archive") {
                    statsVisuals.push(h(TimelineBars, {
                        key: "archive-months", name: "archive-months", title: "归档按月分布",
                        color: "#c08a2e", buckets: bucketize(archive, archiveMonth),
                    }));
                }
            }

            const layersOn = config && config.layers
                ? Object.keys(config.layers).filter((key) => config.layers[key]).join("、")
                : "";
            const footerChip = (key, label, value) => h("span", { key: key, style: S.footerChip },
                h("span", { style: S.footerKey }, label), value);

            // 待办指令：面板**只生成指令文本**，不落盘。
            // 条目上的「改写 / 忘记 / 删关系」会到这里；新增记忆不走这里 ——
            // 标题栏的「+ 记一条」直接 POST memory-add，保存后重新取数。
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
                    jumpToChat !== null
                        ? h("button", {
                            type: "button", className: "om-btn", style: S.ghostButton, "data-memory-jump": "true",
                            onClick: jumpToChat,
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

            // 直接写一条：不用弹编辑器，也不用生成会话指令。
            // 只在「个人偏好 / 项目记忆」两档出现 —— 台账与知识库没有「写一条」的语义。
            const quickComposer = section === "memory" && (layerKey === "preference" || layerKey === "project")
                ? (state.quickOpen === true
                    ? h("div", { style: S.editor, "data-memory-composer": "true" },
                        h("div", { style: { display: "flex", alignItems: "baseline", gap: "8px", flexWrap: "wrap", marginBottom: "6px" } },
                            h("span", { style: { fontSize: "13px", fontWeight: 600, color: M.text } }, "直接记一条"),
                            h("span", { style: S.editorHint }, "一两句就够；写进哪一档由下面的目标决定。"),
                            h("span", { style: { flex: "1 1 auto" } }),
                            h("button", {
                                type: "button", className: "om-btn", style: S.ghostButton,
                                "data-memory-quick-close": "true",
                                onClick: () => patchState({ quickOpen: false }),
                            }, "收起")),
                        h("textarea", {
                            style: Object.assign({}, S.editorTextarea, { minHeight: "64px" }), "data-memory-quick-text": "true",
                            rows: 2,
                            value: state.quickText,
                            placeholder: "例：我不喜欢用「不是 A 而是 B」这种句式；或者：这个项目只提交一版公开历史。",
                            onChange: (event) => patchState({ quickText: event.target.value }),
                        }),
                        h("div", { style: S.editorActions },
                            h("div", { style: S.chips, "data-memory-quick-target": state.quickTarget },
                                h("button", {
                                    type: "button", className: "om-btn",
                                    style: state.quickTarget === "user" ? Object.assign({}, S.chip, S.chipOn) : Object.assign({}, S.chip, S.chipOff),
                                    "data-memory-quick-target-user": "true",
                                    onClick: () => patchState({ quickTarget: "user" }),
                                }, "个人偏好（全局）"),
                                h("button", {
                                    type: "button", className: "om-btn",
                                    style: state.quickTarget === "project" ? Object.assign({}, S.chip, S.chipOn) : Object.assign({}, S.chip, S.chipOff),
                                    "data-memory-quick-target-project": "true",
                                    onClick: () => patchState({ quickTarget: "project" }),
                                }, "项目记忆（本工作区）")),
                            h("span", { style: { flex: "1 1 auto" } }),
                            h("button", {
                                type: "button", className: "om-btn",
                                style: Object.assign({}, S.ghostButton, {
                                    borderColor: M.accent, color: M.accent, fontWeight: 600,
                                    opacity: state.quickBusy === true || String(state.quickText || "").trim() === "" ? 0.55 : 1,
                                }),
                                "data-memory-quick-add": "true",
                                disabled: state.quickBusy === true || String(state.quickText || "").trim() === "",
                                onClick: quickAdd,
                            }, state.quickBusy === true ? "写入中…" : "添加")),
                    )
                    : h("button", {
                        type: "button", className: "om-btn om-quick-open", style: S.quickOpen,
                        "data-memory-quick-open": "true",
                        onClick: () => patchState({ quickOpen: true }),
                    },
                        h("span", { style: { color: M.accent, fontSize: "15px", lineHeight: 1 } }, "+"),
                        h("span", null, "记一条")))
                : null;

            // 删除确认：kb-drop 删的是**整份文档**（见 src/memory.js:4512），
            // 所以把标题 / 路径 / 块数都摆出来，并说明关系边一起清掉，别让用户以为是删一块。
            const deleteCard = state.deleteTarget === null || state.deleteTarget === undefined ? null : h("div", {
                style: S.pending, "data-memory-delete-card": String(state.deleteTarget.id),
            },
                h("div", { style: S.pendingHead },
                    h("span", { style: S.pendingTitle }, "删除文档：" + state.deleteTarget.title),
                    h("span", { style: S.pendingNote }, "删整篇 · 含 " + String(state.deleteTarget.chunks || 0) + " 块"),
                ),
                state.deleteTarget.path !== ""
                    ? h("div", { style: S.pendingHint }, state.deleteTarget.path)
                    : null,
                h("div", { style: S.pendingActions },
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-memory-delete-cancel": "true",
                        onClick: () => patchState({ deleteTarget: null }),
                    }, "取消"),
                    h("button", {
                        type: "button", className: "om-btn",
                        style: Object.assign({}, S.ghostButton, { borderColor: M.error, color: M.error, fontWeight: 600 }),
                        "data-memory-delete-confirm": "true",
                        disabled: state.saving === true,
                        onClick: confirmDeleteDoc,
                    }, state.saving === true ? "删除中…" : "确认删除整篇"),
                ),
                h("div", { style: S.pendingHint }, "文档、它的块与块间关系会一起清掉；原文件不保留，无法从面板恢复。"),
            );

            // 根元素是面板的**唯一外框**（不是滚动容器，见 S.browser 的注释）；
            // om-scroll 仍然留着 —— 它管 reduced-motion 下的 scroll-behavior，
            // 也给宿主滚动容器的滚动条让出一套样式命名；ref 给「再显示 N 条」用。
            return h("div", { className: "om-mem om-scroll", style: S.browser, ref: listRef, "data-memory-panel": "scroll" },
                // 伪类 / 关键帧 / 媒体查询在这里；内容是常量，没有注入面。
                h("style", { key: "panel-css", dangerouslySetInnerHTML: { __html: PANEL_CSS } }),
                // 标题不吸顶；页签与筛选单独 sticky，正文紧跟着往下铺。
                topBar,
                stickyNav,
                quickComposer,
                deleteCard,
                // 出错时列表照样显示（可能是上一次的数据）：只把错误顶在上面，
                // 而不是用错误页把已有内容盖掉。
                state.error !== null ? h("div", { style: S.error }, state.error) : null,
                pendingCard,
                body,
                // 统计与容量在内容之后：打开面板先读记忆，再决定要不要看用量。
                // 图库是独立视图：记忆的指标卡 / 容量 / 存储域与它无关，不往图库下面塞。
                hasData && state.tab !== "gallery"
                    ? h("div", { style: S.overview, "data-memory-overview": "true", "data-memory-dashboard": "true" },
                        h("div", { style: S.dashKicker }, "MEMORY DASHBOARD"),
                        h("div", { style: S.overviewTitle },
                            h("span", null, "统计与容量"),
                            h("span", { style: S.overviewNote }, "内容优先 · 指标在后"),
                        ),
                        statsVisuals,
                        dashboards.length > 0
                            ? h("div", { className: "om-grid-dash", style: S.panelGrid }, dashboards)
                            : null,
                    )
                    : null,
                config && state.tab !== "gallery"
                    ? h("div", { style: S.footerMeta },
                        footerChip("scope", "范围", String(config.scope || "—")),
                        footerChip("user", "用户偏好", String(config.userScope || "—")),
                        footerChip("layers", "开启的层", layersOn || "（无）"),
                        footerChip("capture", "主动记录", config.autoCapture ? "开" : "关"),
                    )
                    : null,
            );
        }

        // ── 预览面板 ──────────────────────────────────────────────────────

        /**
         * 预览面板在会话页签行里的 id（`conversation.view` 的一格，order 50）。
         *
         * 排在「记忆」（40）右侧：两个面板是同一类东西 —— 「会话里看不到的东西」的
         * 只读窗口；用户要从左到右依次是 对话｜轨迹｜上下文｜编队｜记忆｜预览。
         * `order` 是唯一决定这件事的东西，别再按注册先后去推。
         */
        const PREVIEW_PANEL_ID = "office-preview";

        /**
         * 三条数据端点（宿主侧 src/preview-view.js 注册）。
         *
         * 与记忆端点同一条通道口径：同源相对路径、只读、失败时 body 里带 error。
         * 拆成三条而不是一条：清单要快（打开面板就列），大纲要解析 pptx（不贵），
         * 取图要起渲染引擎（最贵，按页来）—— 合成一条请求会让「看个清单」也等 3 秒。
         */
        const PREVIEW_FILES_PATH = "/office-preview/files";
        const PREVIEW_OUTLINE_PATH = "/office-preview/outline";
        const PREVIEW_PAGE_PATH = "/office-preview/page";
        /**
         * 文本文件（.md/.csv/.tex…）的原文通道（第七十一轮）。
         *
         * 第六十八轮做预览面板时只实现了「PPT 逐页大纲 + 页面图」这一条链，其它格式
         * 只在清单里出现、选中后右列一句 hint —— 用户实测「预览无法看 markdown 文件」
         * 说的就是它。这条通道把文件正文取回来（服务端只读、2 MiB 截断），
         * 客户端再按扩展名分别渲染：.md 走下面的极简渲染器，其余按等宽原文。
         */
        const PREVIEW_TEXT_PATH = "/office-preview/text";

        /** 默认 dpi 与可选档位（与服务端默认值 110 同步；档位取「够看清又不至于几 MB」）。 */
        const PREVIEW_DPI_DEFAULT = 110;
        const PREVIEW_DPI_CHOICES = [72, 96, 110, 144, 200];

        /** 从一条风险文案里读页码（`第 3 页 …` / `P3 …` 两种写法都认），读不出返回 null。 */
        function previewPageOfWarning(text) {
            const matched = /(?:第\s*|\bP)(\d+)\s*页?/.exec(String(text));
            return matched === null ? null : Number(matched[1]);
        }

        /** 文件体积的短写法：面板里的下拉项只放得下这么几个字。 */
        function previewSizeText(bytes) {
            const value = Number(bytes);
            if (!Number.isFinite(value) || value <= 0) return "";
            if (value < 1024) return value + " B";
            if (value < 1024 * 1024) return Math.round(value / 1024) + " KB";
            return (Math.round(value / (1024 * 1024) * 10) / 10) + " MB";
        }

        /**
         * 取一份预览端点数据。
         *
         * 失败一律转成 Error，body 里的 error 原样带出来（服务端已经写好「下一步」，
         * 面板不该把它换成「加载失败」这种没有信息量的话）。
         * fetch 只在浏览器里有：没有它时给一条能照做的错，而不是让整格白屏。
         */
        function fetchPreview(url) {
            if (typeof fetch !== "function") {
                return Promise.reject(new Error("当前环境没有 fetch，预览面板只在浏览器里工作。"));
            }
            let pending;
            try {
                pending = fetch(url, { headers: { accept: "application/json" } });
            } catch (error) {
                return Promise.reject(error);
            }
            return Promise.resolve(pending).then(
                (response) => Promise.resolve(response && typeof response.json === "function" ? response.json() : null)
                    .then((body) => ({ status: response && response.status ? response.status : 0, body }), () => ({ status: response && response.status ? response.status : 0, body: null })),
            ).then((result) => {
                const body = result.body;
                if (body === null || typeof body !== "object") {
                    throw new Error("预览端点没有返回 JSON（HTTP " + result.status + "）。");
                }
                if (body.ok !== true) {
                    const error = new Error(String(body.error || "读取预览数据失败。"));
                    error.status = result.status;
                    throw error;
                }
                return body;
            });
        }

        // ── Markdown 渲染（零依赖） ────────────────────────────────────────
        //
        // 为什么自己写而不是引 marked / markdown-to-jsx：lib/client.js 是**手写的
        // 单文件产物**（没有构建步骤），运行时只从宿主拿 `react` 与
        // `react/jsx-runtime`，第三方包进不了这个 bundle。而 .md 预览真正需要的
        // 只是「读懂一篇文档」的那一小片语法 —— 标题、段落、列表、代码、引用、
        // 表格、链接与强调 —— 不是完整 CommonMark 规范。
        //
        // 安全口径（比「语法全不全」要紧）：
        //   · 不用 dangerouslySetInnerHTML：所有文本都走 React 文本节点，文档里写
        //     <script> 只会被当成那八个字符显示出来；
        //   · 链接只放行 http/https/mailto 与相对路径，`javascript:` / `data:` 降级成
        //     纯文本；
        //   · 图片一律不加载（远端图会暴露「谁打开了这份预览」，而 .md 里的图多是
        //     相对路径，在这里也解析不了），显示成 `🖼 说明` 的占位。
        //
        // 预算：正文 200 KB / 6000 行 / 6000 个节点 / 6 层嵌套。服务端单次最多给
        // 2 MiB，但把 2 MiB 的日志铺成几十万个 DOM 节点会把标签页卡死；超出预算就
        // 渲染到预算为止并挂一行提示（「完整正文请用编辑器打开」）。
        const MD_MAX_CHARS = 200000;
        const MD_MAX_LINES = 6000;
        const MD_MAX_NODES = 6000;
        const MD_MAX_DEPTH = 6;
        const MD_TRUNCATED = "…（内容过长，只渲染了前一部分；完整正文请用编辑器打开这份文件）";

        /** .md 渲染后的版式：跟随宿主类型刻度与主题令牌，不另起一套。 */
        const MDV = {
            wrap: {
                display: "flex", flexDirection: "column", gap: "8px",
                maxWidth: "100%", overflowWrap: "anywhere", wordBreak: "break-word",
                color: M.text
            },
            h1: { fontSize: "20px", lineHeight: "30px", fontWeight: 600, margin: "8px 0 0" },
            h2: { fontSize: "17px", lineHeight: "26px", fontWeight: 600, margin: "8px 0 0" },
            h3: { fontSize: "15px", lineHeight: "24px", fontWeight: 600, margin: "6px 0 0" },
            h4: { fontSize: "14px", lineHeight: "22px", fontWeight: 600, margin: "4px 0 0" },
            h5: { fontSize: "13px", lineHeight: "20px", fontWeight: 600, margin: "4px 0 0" },
            h6: { fontSize: "12px", lineHeight: "18px", fontWeight: 600, margin: "4px 0 0", color: M.textSecondary },
            p: Object.assign({ margin: 0 }, T.body),
            list: Object.assign({ margin: "0 0 0 4px", paddingLeft: "20px", display: "flex", flexDirection: "column", gap: "3px" }, T.body),
            listItem: { margin: 0, paddingLeft: "2px" },
            quote: {
                margin: 0, padding: "2px 0 2px 12px",
                borderLeft: "3px solid " + M.strokeStrong, color: M.textSecondary
            },
            pre: {
                margin: 0, padding: "10px 12px", borderRadius: M.radiusSm,
                border: "1px solid " + M.stroke, background: M.sunken,
                fontFamily: M.mono, fontSize: "12px", lineHeight: "18px",
                overflowX: "auto", whiteSpace: "pre", tabSize: 4
            },
            code: {
                padding: "1px 5px", borderRadius: "5px",
                background: M.sunken, border: "1px solid " + M.stroke,
                fontFamily: M.mono, fontSize: "0.92em"
            },
            link: { color: M.business, textDecoration: "underline", textUnderlineOffset: "2px" },
            del: { opacity: 0.62 },
            hr: { margin: "6px 0", border: 0, borderTop: "1px solid " + M.stroke },
            table: { borderCollapse: "collapse", width: "100%", fontSize: "12px", lineHeight: "18px" },
            th: {
                textAlign: "left", padding: "5px 8px", border: "1px solid " + M.stroke,
                background: M.sunken, color: M.textSecondary, fontWeight: 600, whiteSpace: "nowrap"
            },
            td: { padding: "5px 8px", border: "1px solid " + M.stroke, verticalAlign: "top" },
            imageChip: {
                display: "inline-flex", alignItems: "center", gap: "2px",
                padding: "0 5px", borderRadius: M.pill,
                border: "1px solid " + M.stroke, background: M.sunken, color: M.textTertiary
            },
            truncated: Object.assign({ padding: "6px 10px", borderRadius: M.radiusSm, background: M.sunken, color: M.textTertiary }, T.tiny),
        };

        /** 允许的链接：http/https/mailto 与相对路径；其余（含 javascript:）返回 null。 */
        function mdSafeHref(raw) {
            const value = String(raw === undefined || raw === null ? "" : raw).trim();
            if (value === "") return null;
            if (/^(?:javascript|vbscript|data|file):/i.test(value)) return null;
            if (/^[a-z][a-z0-9+.-]*:/i.test(value)) return /^(?:https?|mailto):/i.test(value) ? value : null;
            return value;
        }

        /** 表格行 → 单元格（认 `\|` 转义；首尾的竖线不算空单元格）。 */
        function mdTableCells(line) {
            let text = String(line).trim();
            if (text.charAt(0) === "|") text = text.slice(1);
            if (text.charAt(text.length - 1) === "|") text = text.slice(0, -1);
            if (text === "") return [];
            const cells = [];
            let current = "";
            for (let index = 0; index < text.length; index += 1) {
                const char = text.charAt(index);
                if (char === "\\" && text.charAt(index + 1) === "|") { current += "|"; index += 1; continue; }
                if (char === "|") { cells.push(current.trim()); current = ""; continue; }
                current += char;
            }
            cells.push(current.trim());
            return cells;
        }

        /**
         * 行内 Markdown：`code`、**粗**、*斜*、~~删~~、[链接](url)、<自动链接>、反斜杠转义。
         *
         * 用**一条全局正则**扫一遍，而不是逐字符 `text.slice(index)`：后者在长文档里
         * 每次复制剩余串，整体会退化成 O(n²)（200 KB 的文档能卡住好几秒）。
         */
        const MD_INLINE = /\\[\\`*_{}\[\]()#+\-.!~|>]|`[^`\n]+`|!\[[^\]\n]*\]\([^)\s]+\)|\[[^\]\n]+\]\([^)\s]+\)|<(?:https?|mailto):[^>\s]+>|(\*\*|__)(?:(?!\1)[^\n])+?\1|~~[^~\n]+?~~|\*[^*\n]+?\*|_[^_\n]+?_/g;

        function mdInline(text, keyBase) {
            const source = String(text === undefined || text === null ? "" : text);
            const out = [];
            let cursor = 0;
            let serial = 0;
            // 每次调用一个**新的**扫描器：mdInline 会递归（粗体 / 斜体 / 删除线里再解行内），
            // 共用同一个带 /g 的正则时，内层把 lastIndex 重置回 0，外层循环就会反复
            // 从同一处重新匹配 —— 表现为「一直 push 节点直到堆爆掉」（实测 4 GB OOM）。
            const scanner = new RegExp(MD_INLINE.source, "g");
            let match = scanner.exec(source);
            while (match !== null) {
                if (match.index > cursor) out.push(source.slice(cursor, match.index));
                cursor = match.index + match[0].length;
                const token = match[0];
                const key = keyBase + "i" + serial;
                serial += 1;
                if (token.charAt(0) === "\\") { out.push(token.slice(1)); }
                else if (token.charAt(0) === "`") { out.push(h("code", { key: key, style: MDV.code }, token.slice(1, -1))); }
                else if (token.slice(0, 2) === "![") {
                    const parts = /^!\[([^\]]*)\]\(([^)\s]+)\)$/.exec(token);
                    out.push(h("span", { key: key, style: MDV.imageChip, title: parts === null ? "" : parts[2] },
                        "🖼 " + (parts === null || parts[1] === "" ? "图片" : parts[1])));
                }
                else if (token.charAt(0) === "[") {
                    const parts = /^\[([^\]]+)\]\(([^)\s]+)\)$/.exec(token);
                    const href = parts === null ? null : mdSafeHref(parts[2]);
                    if (href === null) out.push(parts === null ? token : parts[1]);
                    else out.push(h("a", { key: key, href: href, target: "_blank", rel: "noreferrer noopener", style: MDV.link }, parts[1]));
                }
                else if (token.charAt(0) === "<") {
                    const href = mdSafeHref(token.slice(1, -1));
                    if (href === null) out.push(token);
                    else out.push(h("a", { key: key, href: href, target: "_blank", rel: "noreferrer noopener", style: MDV.link }, token.slice(1, -1)));
                }
                else if (token.slice(0, 2) === "**" || token.slice(0, 2) === "__") {
                    out.push(h("strong", { key: key }, mdInline(token.slice(2, -2), key)));
                }
                else if (token.slice(0, 2) === "~~") {
                    out.push(h("del", { key: key, style: MDV.del }, mdInline(token.slice(2, -2), key)));
                }
                else {
                    out.push(h("em", { key: key }, mdInline(token.slice(1, -1), key)));
                }
                match = scanner.exec(source);
            }
            if (cursor < source.length) out.push(source.slice(cursor));
            return out;
        }

        /**
         * 块级 Markdown → React 元素数组。
         *
         * 语法子集：ATX / Setext 标题、段落、围栏代码、引用、有序与无序列表
         * （含嵌套与任务项）、GFM 表格、分隔线。HTML 块按普通文本渲染。
         */
        function mdBlocks(source, depth, budget) {
            const lines = Array.isArray(source) ? source : String(source).split("\n");
            const out = [];
            const nextKey = budget.nextKey;
            let index = 0;
            const atBudget = () => budget.count >= MD_MAX_NODES;
            const bid = () => { budget.count += 1; return nextKey(); };

            while (index < lines.length) {
                if (atBudget()) { out.push(h("div", { key: nextKey(), style: MDV.truncated }, MD_TRUNCATED)); break; }
                const line = String(lines[index]);
                if (/^\s*$/.test(line)) { index += 1; continue; }

                // ① 围栏代码块（``` / ~~~，闭合要同字符且不短于开标记）
                const fence = /^ {0,3}(`{3,}|~{3,})\s*([^`]*)$/.exec(line);
                if (fence !== null) {
                    const marker = fence[1];
                    const language = String(fence[2] || "").trim();
                    const close = new RegExp("^ {0,3}" + (marker.charAt(0) === "`" ? "`" : "~") + "{" + marker.length + ",}\\s*$");
                    const body = [];
                    index += 1;
                    while (index < lines.length && !close.test(String(lines[index]))) {
                        body.push(String(lines[index]));
                        index += 1;
                    }
                    index += 1;
                    out.push(h("pre", { key: bid(), style: MDV.pre, "data-md-code": language === "" ? "true" : language },
                        h("code", null, body.join("\n"))));
                    continue;
                }

                // ② ATX 标题
                const atx = /^ {0,3}(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line);
                if (atx !== null) {
                    const level = atx[1].length;
                    out.push(h("h" + level, { key: bid(), style: MDV["h" + level] }, mdInline(atx[2], nextKey())));
                    index += 1;
                    continue;
                }

                // ③ 分隔线（必须排在列表之前：`---` 不该被当成空列表项）
                if (/^ {0,3}(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$/.test(line)) {
                    out.push(h("hr", { key: bid(), style: MDV.hr }));
                    index += 1;
                    continue;
                }

                // ④ GFM 表格：本行有 `|`，下一行是「每格至少一个 -」的分隔行
                if (line.indexOf("|") >= 0 && index + 1 < lines.length) {
                    const header = mdTableCells(line);
                    const delimiter = mdTableCells(lines[index + 1]);
                    const aligned = header.length > 0 && delimiter.length === header.length
                        && delimiter.every((cell) => /^:?-+:?$/.test(cell.replace(/\s+/g, "")));
                    if (aligned) {
                        const rowKey = nextKey();
                        const cellStyle = header.map((cell, at) => {
                            const spec = delimiter[at].replace(/\s+/g, "");
                            if (spec.charAt(0) === ":" && spec.charAt(spec.length - 1) === ":") return Object.assign({}, MDV.th, { textAlign: "center" });
                            if (spec.charAt(spec.length - 1) === ":") return Object.assign({}, MDV.th, { textAlign: "right" });
                            return MDV.th;
                        });
                        const rows = [];
                        index += 2;
                        while (index < lines.length && String(lines[index]).indexOf("|") >= 0 && !/^\s*$/.test(String(lines[index]))) {
                            rows.push(mdTableCells(String(lines[index])));
                            index += 1;
                        }
                        budget.count += 1;
                        out.push(h("table", { key: rowKey, style: MDV.table, "data-md-table": "true" },
                            h("thead", null,
                                h("tr", null, header.map((cell, at) => h("th", { key: rowKey + "h" + at, style: cellStyle[at] }, mdInline(cell, rowKey + "h" + at))))),
                            h("tbody", null, rows.map((row, rowIndex) => h("tr", { key: rowKey + "r" + rowIndex },
                                header.map((cell, at) => h("td", { key: rowKey + "r" + rowIndex + "c" + at, style: MDV.td },
                                    mdInline(row[at] === undefined ? "" : row[at], rowKey + "r" + rowIndex + "c" + at))))))
                        ));
                        continue;
                    }
                }

                // ⑤ 引用：连续的 `>` 行，剥掉标记后递归
                if (/^ {0,3}>/.test(line)) {
                    const body = [];
                    while (index < lines.length && /^ {0,3}>/.test(String(lines[index]))) {
                        body.push(String(lines[index]).replace(/^ {0,3}> ?/, ""));
                        index += 1;
                    }
                    out.push(h("blockquote", { key: bid(), style: MDV.quote },
                        depth >= MD_MAX_DEPTH
                            ? h("div", null, body.join("\n"))
                            : mdBlocks(body, depth + 1, budget)));
                    continue;
                }

                // ⑥ 列表：与块首同缩进的 marker 开新项；更深的行留给该项的递归
                const listHead = /^( {0,12})([-*+]|\d{1,9}[.)])\s+/.exec(line);
                if (listHead !== null) {
                    const baseIndent = listHead[1].length;
                    const ordered = /\d/.test(listHead[2].charAt(0));
                    const blockLines = [];
                    while (index < lines.length) {
                        const current = String(lines[index]);
                        if (/^\s*$/.test(current)) {
                            // 空行：下一行还接得上就继续，否则本块结束（空行不要带进项里）
                            let probe = index + 1;
                            while (probe < lines.length && /^\s*$/.test(String(lines[probe]))) probe += 1;
                            if (probe >= lines.length) break;
                            const probeIndent = (/^(\s*)/.exec(String(lines[probe]))[1] || "").length;
                            if (probeIndent <= baseIndent) break;
                            blockLines.push("");
                            index += 1;
                            continue;
                        }
                        const marker = /^( {0,12})([-*+]|\d{1,9}[.)])\s+/.exec(current);
                        if (marker !== null) {
                            if (marker[1].length < baseIndent) break;
                            blockLines.push(current);
                            index += 1;
                            continue;
                        }
                        const indent = (/^(\s*)/.exec(current)[1] || "").length;
                        if (indent <= baseIndent) break;
                        blockLines.push(current);
                        index += 1;
                    }
                    // 切项 + 去缩进：续行按「marker 之后的正文列」去缩进，递归看见的
                    // 才是段落而不是 4 空格代码块（嵌套列表因为缩进更深自己会再进一层）。
                    const items = [];
                    let itemContent = null;
                    let contentColumn = baseIndent + 2;
                    for (const blockLine of blockLines) {
                        const marker = /^( {0,12})([-*+]|\d{1,9}[.)])\s+(.*)$/.exec(blockLine);
                        if (marker !== null && marker[1].length === baseIndent) {
                            if (itemContent !== null) items.push(itemContent);
                            contentColumn = baseIndent + marker[2].length + 1;
                            itemContent = [marker[3]];
                            continue;
                        }
                        if (itemContent === null) continue;
                        itemContent.push(blockLine.replace(new RegExp("^ {0," + contentColumn + "}"), ""));
                    }
                    if (itemContent !== null) items.push(itemContent);
                    const start = /^\d{1,9}/.exec(listHead[2]);
                    // 任务项（GFM）：`- [ ]` / `- [x]` 画成方框字符，比原样显示 `[ ]` 好读。
                    const renderItem = (itemLines, itemIndex) => {
                        const itemKey = nextKey() + "li" + itemIndex;
                        const first = itemLines.length > 0 ? itemLines[0] : "";
                        const task = /^\[([ xX])\]\s*(.*)$/.exec(first);
                        const box = task === null ? "" : (task[1].toLowerCase() === "x" ? "☑ " : "☐ ");
                        const inline = task === null ? first : task[2];
                        if (itemLines.length === 1) {
                            return h("li", { key: itemKey, style: MDV.listItem }, h("span", null, box, mdInline(inline, itemKey + "i")));
                        }
                        // 嵌套层数封顶后不再递归：把续行拍平，宁可少一层缩进也不能无限展开。
                        if (depth >= MD_MAX_DEPTH) {
                            return h("li", { key: itemKey, style: MDV.listItem }, h("span", null, box, mdInline(itemLines.join(" "), itemKey + "i")));
                        }
                        return h("li", { key: itemKey, style: MDV.listItem }, h("div", null, mdBlocks(itemLines, depth + 1, budget)));
                    };
                    out.push(ordered
                        ? h("ol", { key: bid(), style: MDV.list, start: start === null ? 1 : Number(start[0]) },
                            items.map((itemLines, itemIndex) => renderItem(itemLines, itemIndex)))
                        : h("ul", { key: bid(), style: MDV.list },
                            items.map((itemLines, itemIndex) => renderItem(itemLines, itemIndex))));
                    continue;
                }

                // ⑦ Setext 标题：正文行 + 下一行 `===` / `---`
                if (index + 1 < lines.length && /^ {0,3}(?:=+|-+)\s*$/.test(String(lines[index + 1])) && line.indexOf("|") < 0) {
                    const level = String(lines[index + 1]).trim().charAt(0) === "=" ? 1 : 2;
                    out.push(h("h" + level, { key: bid(), style: MDV["h" + level] }, mdInline(line.trim(), nextKey())));
                    index += 2;
                    continue;
                }

                // ⑧ 段落：收集到空行或下一个块起始
                const para = [];
                while (index < lines.length) {
                    const current = String(lines[index]);
                    if (/^\s*$/.test(current)) break;
                    if (para.length > 0 && (
                        /^ {0,3}#{1,6}\s+/.test(current)
                        || /^ {0,3}(?:`{3,}|~{3,})/.test(current)
                        || /^ {0,3}>/.test(current)
                        || /^ {0,12}(?:[-*+]|\d{1,9}[.)])\s+/.test(current)
                        || /^ {0,3}(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$/.test(current)
                    )) break;
                    para.push(current.trim());
                    index += 1;
                }
                if (para.length > 0) {
                    out.push(h("p", { key: bid(), style: MDV.p }, mdInline(para.join(" "), nextKey())));
                }
            }
            return out;
        }

        /**
         * 一份 .md 的完整渲染入口：做预算裁剪后交给块级解析器。
         * 返回 { nodes, truncated } —— truncated 为 true 时调用方在正文上方挂提示。
         */
        function renderMarkdown(source) {
            let text = String(source === undefined || source === null ? "" : source);
            let truncated = false;
            if (text.length > MD_MAX_CHARS) { text = text.slice(0, MD_MAX_CHARS); truncated = true; }
            const lines = text.split("\n");
            if (lines.length > MD_MAX_LINES) { text = lines.slice(0, MD_MAX_LINES).join("\n"); truncated = true; }
            const budget = { count: 0 };
            let serial = 0;
            // 闭包而不是对象方法：mdBlocks 里把它取成局部变量再调，方法写法会丢 this。
            budget.nextKey = () => { serial += 1; return "md-" + serial; };
            const nodes = mdBlocks(text, 0, budget);
            if (budget.count >= MD_MAX_NODES) truncated = true;
            return { nodes: nodes, truncated: truncated };
        }

        /**
         * 预览面板：顶部选文件 + 刷新 + dpi，左列逐页大纲，右列当前页大图，底部风险清单。
         *
         * 取数同样写成显式的 loadFiles / loadOutline / loadPage（不靠 effect 依赖数组），
         * 理由与记忆面板一致：测试里的 React 替身对 effect 只跑一次，抽成函数才测得到
         * 「文件 → 大纲 → 取图」这条链。
         *
         * 一条约定：面板**不写**任何东西。渲染图只落在宿主缓存目录，这个面板连
         * 「导出」按钮都不给 —— 要落盘请走会话里的 office_run。
         */
        function PreviewBrowserPanel(props) {
            /**
             * 当前会话的工作目录：与记忆面板同一条口径（`conversation.view` 的
             * session 作用域标准 props）。拿不到就不取数，显示「没有工作目录」的空态 ——
             * 让宿主随便挑一个根去列文件，比空态更糟：用户会以为那是自己目录里的东西。
             */
            const useSessions = typeof props.useSessions === "function" ? props.useSessions : useNoSessions;
            const sessionCwdRaw = useSessions((snapshot) => (
                typeof props.sessionId === "string" && props.sessionId !== ""
                    ? snapshot?.byId?.[props.sessionId]?.cwd
                    : undefined
            ));
            const bootCwd = typeof sessionCwdRaw === "string" ? sessionCwdRaw : "";

            const [state, setState] = React.useState({
                cwd: bootCwd,
                filesLoading: bootCwd !== "",
                files: [],
                root: "",
                label: "",
                total: 0,
                truncated: false,
                selected: "",
                outline: null,
                outlineLoading: false,
                page: 1,
                dpi: PREVIEW_DPI_DEFAULT,
                image: null,
                imageLoading: false,
                // 文本文件（.md/.csv/.tex…）的正文通道：与「页面图」互斥 —— 一份文件
                // 要么走渲染图，要么走原文，不会同时要两份数据（见 loadOutline 的分支）。
                text: null,
                textLoading: false,
                textError: null,
                // 顶部错误条：清单 / 大纲这种「没它就什么都看不了」的失败。
                error: null,
                // 右侧取图失败：大纲还在，用户换个页码或 dpi 就能再试，不该整页报错。
                notice: null,
            });
            // 异步回来时拿「最新的」state：闭包里那份是发起请求那一刻的旧值。
            const latest = React.useRef(state);
            latest.current = state;

            // patch 立刻并进 latest.current：files 落地后紧接着 loadOutline 时，
            // 两处 patch 都基于同一份 latest —— 只写 setState 的话，第二次会把第一次
            // 的 files 覆盖掉（两次 patch 之间没有渲染）。
            const patchState = (patch) => {
                const merged = Object.assign({}, latest.current, patch);
                latest.current = merged;
                setState(merged);
            };

            const failWith = (error) => {
                patchState({
                    filesLoading: false, outlineLoading: false, imageLoading: false,
                    error: error && error.message ? error.message : String(error),
                });
            };

            const loadPage = (target, overrides) => {
                const opts = overrides || {};
                const base = Object.assign({}, latest.current, opts);
                if (base.selected === "" || base.cwd === "") return;
                const outline = base.outline || {};
                // 不可渲染的格式（.md/.csv）：不请求取图 —— 服务端会 400，请求一次只是白花时间。
                if (outline.renderable === false) return;
                const pageNo = Math.max(1, Math.trunc(Number(target)) || 1);
                patchState(Object.assign({}, opts, { page: pageNo, imageLoading: true, image: null, notice: null }));
                const params = [
                    "cwd=" + encodeURIComponent(base.cwd),
                    "path=" + encodeURIComponent(base.selected),
                    "page=" + encodeURIComponent(String(pageNo)),
                    "dpi=" + encodeURIComponent(String(base.dpi)),
                ];
                fetchPreview(PREVIEW_PAGE_PATH + "?" + params.join("&")).then((body) => {
                    patchState({ imageLoading: false, image: body, notice: null });
                }, (error) => {
                    // 引擎不可用 / 渲染失败都落在这里：右侧显示原文（服务端写了下一步），
                    // 左列大纲与底部风险清单照旧可用。
                    patchState({ imageLoading: false, image: null, notice: error && error.message ? error.message : String(error) });
                });
            };

            /**
             * 取文本文件的正文（.md/.csv/.tex…）。
             *
             * 与 loadPage 互斥：这条链成功时右列画的是正文，不会再有页面图。
             * 失败也不整页报错 —— 大纲与清单还在，用户换一份文件或点「刷新」就能再试。
             */
            const loadText = (filePath, overrides) => {
                const opts = overrides || {};
                const base = Object.assign({}, latest.current, opts);
                if (filePath === "" || base.cwd === "") return;
                patchState(Object.assign({}, opts, { textLoading: true, text: null, textError: null }));
                const url = PREVIEW_TEXT_PATH
                    + "?cwd=" + encodeURIComponent(base.cwd)
                    + "&path=" + encodeURIComponent(filePath);
                fetchPreview(url).then((body) => {
                    patchState({ textLoading: false, text: body, textError: null });
                }, (error) => {
                    patchState({
                        textLoading: false, text: null,
                        textError: error && error.message ? error.message : String(error),
                    });
                });
            };

            const loadOutline = (filePath, overrides) => {
                const opts = overrides || {};
                const base = Object.assign({}, latest.current, opts);
                if (filePath === "" || base.cwd === "") return;
                patchState(Object.assign({}, opts, {
                    selected: filePath, outlineLoading: true, outline: null,
                    image: null, imageLoading: false, page: 1, notice: null, error: null,
                    // 换文件时把上一份的正文一起丢掉：不清的话，选中 .md 之前会先闪一下
                    // 上一份的正文，选中 PPT 之后又会残留一段没人看的文本。
                    text: null, textLoading: false, textError: null,
                }));
                const url = PREVIEW_OUTLINE_PATH
                    + "?cwd=" + encodeURIComponent(base.cwd)
                    + "&path=" + encodeURIComponent(filePath);
                fetchPreview(url).then((body) => {
                    patchState(Object.assign({}, opts, { selected: filePath, outlineLoading: false, outline: body }));
                    // 文本类文件（服务端给 text:true）没有页面图可渲染：走原文通道。
                    // 第七十一轮补上的就是这条分支 —— 第六十八轮只做了清单 + 大纲，
                    // 于是用户在预览里选 .md 时右列只有一句「只做文件清单」。
                    if (body !== null && body.text === true) {
                        loadText(filePath, Object.assign({}, opts, { selected: filePath, outline: body }));
                        return;
                    }
                    // 打开文件就自动渲染第 1 页：用户的下一步动作永远是「看一眼」。
                    loadPage(1, Object.assign({}, opts, { selected: filePath, outline: body }));
                }, failWith);
            };

            const loadFiles = () => {
                const base = latest.current;
                if (base.cwd === "") {
                    patchState({ filesLoading: false });
                    return;
                }
                patchState({ filesLoading: true, error: null });
                fetchPreview(PREVIEW_FILES_PATH + "?cwd=" + encodeURIComponent(base.cwd)).then((body) => {
                    const files = Array.isArray(body.files) ? body.files : [];
                    const cwd = body.root || latest.current.cwd;
                    patchState({
                        filesLoading: false, files: files, cwd: cwd,
                        root: body.root || "", label: body.label || "",
                        total: Number(body.total) || files.length,
                        truncated: body.truncated === true,
                        error: null,
                    });
                    // 刷新后保持当前选择；选的没了（刚删掉）就回退到最新的那一份。
                    const keep = files.find((item) => item.path === latest.current.selected);
                    if (keep !== undefined) loadOutline(keep.path, { cwd: cwd });
                    else if (files.length > 0) loadOutline(files[0].path, { cwd: cwd });
                    else patchState({ selected: "", outline: null, image: null, notice: null });
                }, failWith);
            };

            // 挂载时取一次清单。之后换文件 / 刷新 / 换 dpi 都是用户自己的动作，不在这里覆盖。
            React.useEffect(() => { loadFiles(); }, []);

            const files = Array.isArray(state.files) ? state.files : [];
            const outline = state.outline && typeof state.outline === "object" ? state.outline : null;
            const slides = outline && Array.isArray(outline.slides) ? outline.slides : [];
            const warnings = outline && Array.isArray(outline.warnings) ? outline.warnings : [];
            const renderable = outline === null || outline.renderable !== false;
            const pageCount = outline && Number.isFinite(outline.pages) ? outline.pages : null;
            // ── 文本类文件（.md/.csv/.tex…，服务端在 outline 里给 text:true）──
            //
            // 这一路与「页面图」互斥：文本不渲染成图，直接把正文画出来。
            // .md 走 mdBlocks 那套极简渲染，其余（csv/tex/txt/code）按等宽原文。
            const textKind = outline !== null && outline.text === true;
            const textBody = state.text !== null && typeof state.text === "object" ? state.text : null;
            const textPath = textBody !== null && typeof textBody.path === "string" && textBody.path !== ""
                ? textBody.path
                : state.selected;
            const textIsMarkdown = /\.(?:md|markdown|mdx)$/i.test(textPath);
            const textMeta = textBody === null ? null : h("div", { style: PV.textMeta, "data-preview-text-meta": "true" },
                h("span", null, textBody.name || textPath),
                h("span", null, previewSizeText(textBody.bytes)),
                Number.isFinite(textBody.lines) ? h("span", null, textBody.lines + " 行") : null,
                textBody.truncated === true ? h("span", { style: PV.cachedChip, title: "服务端只读了前 2 MiB" }, "已截断") : null,
            );
            const textNotice = textBody === null || textBody.truncated !== true ? null
                : h("div", { style: S.pendingHint, "data-preview-text-truncated": "true" },
                    typeof textBody.hint === "string" && textBody.hint !== "" ? textBody.hint : "文件较大：只预览了前面一部分。");
            let textRendered = null;
            if (textBody !== null) {
                if (textIsMarkdown) {
                    const rendered = renderMarkdown(textBody.text);
                    textRendered = h("div", { style: MDV.wrap, "data-preview-text": "markdown" },
                        rendered.truncated ? h("div", { style: MDV.truncated }, MD_TRUNCATED) : null,
                        rendered.nodes);
                } else {
                    textRendered = h("pre", { style: PV.textPre, "data-preview-text": "plain" },
                        String(textBody.text === undefined || textBody.text === null ? "" : textBody.text));
                }
            }
            // 风险角标按「第 N 页」聚合：同一条告警里提到几页就标几页（最多标完为止）。
            const riskyPages = new Set();
            const riskyRows = [];
            for (const warning of warnings) {
                const page = previewPageOfWarning(warning);
                riskyRows.push({ text: String(warning), page: page });
                if (page !== null) riskyPages.add(page);
            }
            // 没有逐页大纲的格式（docx/pdf）也要有「第 1 页」这一行，否则左列是空的，
            // 用户会以为「这份文件打不开」而不是「它只有页没有章节」。
            const pageRows = slides.length > 0
                ? slides
                : (outline !== null && outline.pages === null && renderable ? [{ page: 1, layout: "", title: "" }] : []);
            // 同一件事在右列再说一遍（服务端写的 hint）：左列只给一行时，
            // 用户需要知道这是格式本身没有章节，不是解析失败。
            const formatHint = outline !== null && renderable && outline.pages === null
                && typeof outline.hint === "string" && outline.hint !== ""
                ? h("div", { style: S.pendingHint, "data-preview-hint": "true" }, outline.hint)
                : null;

            const fileOptions = files.map((item) => ({
                value: item.path,
                label: item.name + (previewSizeText(item.bytes) === "" ? "" : " · " + previewSizeText(item.bytes)),
            }));
            const dpiOptions = PREVIEW_DPI_CHOICES.map((value) => ({ value: String(value), label: value + " dpi" }));

            const head = h("div", { style: S.browserHead, "data-preview-head": "sticky" },
                h("div", { style: S.browserHeadIcon }, "▦"),
                h("div", { style: S.browserTitleBlock },
                    h("div", { style: S.browserTitle }, "预览" + (state.label ? " · " + state.label : "")),
                    h("div", { style: S.browserSubtitle },
                        state.selected === ""
                            ? "只读面板：渲染图只进缓存目录，不写工作目录"
                            : state.selected
                                + (pageCount === null ? "" : " · 共 " + pageCount + " 页")
                                + (textKind && textBody !== null && Number.isFinite(textBody.lines)
                                    ? " · " + textBody.lines + " 行"
                                    : "")),
                ),
                h("div", { style: S.browserTools },
                    files.length > 0
                        ? h("select", {
                            value: state.selected,
                            "data-preview-file": "true",
                            onChange: (event) => loadOutline(String(event.target.value)),
                            style: S.select,
                        }, fileOptions.map((option) => h("option", { key: option.value, value: option.value }, option.label)))
                        : null,
                    pageRows.length > 0
                        ? h("select", {
                            value: String(state.dpi),
                            "data-preview-dpi": "true",
                            onChange: (event) => loadPage(latest.current.page, { dpi: Number(event.target.value) }),
                            style: S.select,
                        }, dpiOptions.map((option) => h("option", { key: option.value, value: option.value }, option.label)))
                        : null,
                    h("button", {
                        type: "button", className: "om-btn", style: S.ghostButton,
                        "data-preview-refresh": "true",
                        onClick: () => loadFiles(),
                    }, "刷新"),
                ),
            );

            const emptyPanel = (title, hint, dataName) => h("div", {
                style: S.emptyState, "data-preview-empty": dataName,
            },
                h("div", { style: S.emptyGlyph }, "▦"),
                h("div", { style: S.emptyTitle }, title),
                h("div", { style: S.emptyHint }, hint),
            );

            const pageList = pageRows.length === 0 ? null : h("div", { style: PV.pageList, "data-preview-pages": "true" },
                pageRows.map((row) => {
                    const on = row.page === state.page;
                    const risky = riskyPages.has(row.page);
                    return h("div", {
                        key: String(row.page),
                        className: "om-card",
                        style: on ? Object.assign({}, PV.pageRow, PV.pageRowOn) : PV.pageRow,
                        role: "button",
                        tabIndex: 0,
                        "aria-pressed": on,
                        "data-preview-page": String(row.page),
                        onClick: () => loadPage(row.page),
                        onKeyDown: (event) => {
                            if (event.key !== "Enter" && event.key !== " ") return;
                            event.preventDefault();
                            loadPage(row.page);
                        },
                    },
                        h("span", { style: PV.pageNo }, "P" + row.page),
                        h("span", { style: PV.pageLayout }, row.layout === "" ? "页" : row.layout),
                        h("span", { style: PV.pageTitle }, row.title === "" ? "（无标题）" : row.title),
                        risky ? h("span", { style: PV.riskBadge, title: "这一页有排版风险，见底部清单" }, "⚠") : null,
                    );
                }),
            );

            const pageMeta = state.image === null ? null : h("div", { style: PV.imageMeta },
                h("span", null, "第 " + state.image.page + " 页" + (state.image.pages === null ? "" : " / " + state.image.pages)),
                h("span", null, state.image.width + "×" + state.image.height),
                h("span", null, state.image.dpi + " dpi"),
                state.image.cached === true ? h("span", { style: PV.cachedChip }, "缓存") : null,
                h("span", { style: PV.imagePath }, state.selected),
            );

            const viewer = (() => {
                if (state.selected === "") {
                    return h("div", { style: S.emptyInline }, "从上方选一份产物，或点「刷新」。");
                }
                if (state.outlineLoading) {
                    return h("div", { style: S.loadingState },
                        h("span", { className: "om-spin", style: S.spinner }),
                        "正在读取大纲…");
                }
                // 文本类文件（.md/.csv/.tex…）：右列直接画正文，没有页面图可渲染。
                // 这个分支必须排在 !renderable 之前 —— 文本文件也是 renderable=false，
                // 排在后面就永远走不到（那正是第六十八轮「预览看不了 Markdown」的样子）。
                if (textKind) {
                    if (state.textLoading) {
                        return h("div", { style: S.loadingState },
                            h("span", { className: "om-spin", style: S.spinner }),
                            "正在读取正文…");
                    }
                    if (state.textError !== null) {
                        return h("div", { style: S.error, "data-preview-text-error": "true" },
                            state.textError,
                            h("div", { style: S.pendingHint }, "下一步：点上方「刷新」重试；确认这份文件没有被删除或改名。"));
                    }
                    if (textBody === null) {
                        return h("div", { style: S.emptyInline }, "这份文本还没有取回来：点「刷新」再试一次。");
                    }
                    return h("div", { style: PV.textView }, textNotice, textMeta, textRendered);
                }
                if (!renderable) {
                    return h("div", { style: S.loadingState }, outline && outline.hint
                        ? outline.hint
                        : "这份格式没有页面图。");
                }
                if (state.notice !== null) {
                    return h("div", { style: S.error, "data-preview-notice": "true" }, state.notice);
                }
                if (state.imageLoading) {
                    return h("div", { style: S.loadingState },
                        h("span", { className: "om-spin", style: S.spinner }),
                        "正在渲染第 " + state.page + " 页…（首次要 2～5 秒）");
                }
                if (state.image !== null && typeof state.image.png === "string") {
                    return h("div", { style: PV.imageWrap },
                        pageMeta,
                        h("img", {
                            src: state.image.png,
                            alt: "第 " + state.image.page + " 页预览",
                            "data-preview-image": "true",
                            style: PV.image,
                        }));
                }
                return h("div", { style: S.emptyInline }, "选一页取图：点左侧页码，或换一个 dpi 再渲染。");
            })();

            const warningsBox = riskyRows.length === 0 ? null : h("div", { style: PV.warnings, "data-preview-warnings": "true" },
                h("div", { style: PV.warningsTitle }, "排版风险 " + riskyRows.length + " 条 · 点一条跳到对应页"),
                riskyRows.slice(0, 12).map((row, index) => h("div", {
                    key: "warn-" + index,
                    className: "om-act",
                    style: PV.warningRow,
                    role: "button",
                    tabIndex: 0,
                    "data-preview-warning": String(index),
                    title: row.page === null ? "改稿时看这一条" : "跳到第 " + row.page + " 页",
                    onClick: () => {
                        if (row.page === null) return;
                        loadPage(row.page);
                    },
                },
                    h("span", { style: PV.warningPage }, row.page === null ? "—" : "P" + row.page),
                    h("span", { style: PV.warningText }, row.text),
                )),
                riskyRows.length > 12
                    ? h("div", { style: PV.warningMore }, "只列前 12 条，其余 " + (riskyRows.length - 12) + " 条请在会话里用 office.ppt.read 看。")
                    : null,
            );

            const body = state.error !== null && files.length === 0
                ? emptyPanel("读不到这个工作目录的产物", state.error, "error")
                : state.filesLoading && files.length === 0
                    ? h("div", { style: S.loadingState },
                        h("span", { className: "om-spin", style: S.spinner }),
                        "正在扫描工作目录…")
                    : files.length === 0
                        ? emptyPanel(
                            "这个工作目录里还没有办公产物。",
                            "用 office_run 生成 Word / Excel / PPT / PDF 后点「刷新」；面板只列 cwd 下两层目录里的办公文件。",
                            "no-files",
                        )
                        : h("div", null,
                            state.truncated
                                ? h("div", { style: S.pendingHint, "data-preview-truncated": "true" },
                                    "这一页列了 " + files.length + " 份（最多 20 份），共 " + state.total + " 份；先选一份看，其余改完名再刷新。")
                                : null,
                            // 文本文件没有「页」：单列铺满，左列页码整块省掉（不给一份空的
                            // 导航列占 240px —— 正文才是这一格唯一要看的东西）。
                            textKind
                                ? h("div", { style: PV.textGrid, "data-preview-viewer": "true" }, viewer)
                                : h("div", { style: PV.grid },
                                    h("div", { style: PV.left },
                                        pageList,
                                        pageList === null
                                            ? h("div", { style: S.emptyInline }, "这份格式没有页结构。")
                                            : null,
                                    ),
                                    h("div", { style: PV.right, "data-preview-viewer": "true" }, formatHint, viewer),
                                ),
                            warningsBox,
                        );

            return h("div", { className: "om-prev", style: S.browser, "data-preview-panel": "true" },
                h("style", { key: "panel-css", dangerouslySetInnerHTML: { __html: PANEL_CSS } }),
                head,
                // 出错时清单照旧显示（可能是上一次的数据）：只把错误顶在上面，
                // 而不是用错误页把已有内容盖掉 —— 与记忆面板同一条纪律。
                state.error !== null && files.length > 0 ? h("div", { style: S.error }, state.error) : null,
                // 判 state.cwd 而不是 bootCwd：重渲染时宿主可能不再传 owner props
                // （测试 / 异常路径），那时 cwd 已经写进 state 了，空态不该把内容盖掉。
                state.cwd === ""
                    ? emptyPanel(
                        "这个会话还没有工作目录。",
                        "打开一个带工作目录的会话（或在会话里跑一次办公工具）再回到这一格：预览只读当前会话的 cwd。",
                        "no-cwd",
                    )
                    : body,
                h("div", { style: PV.footer },
                    h("span", null, "只读：图片进缓存目录，交付物仍由 office_run 写"),
                    state.root === "" ? null : h("span", { style: PV.imagePath }, state.root),
                ),
            );
        }

        /**
         * 预览面板的版式：左列页码 + 右列大图。
         *
         * 左列固定 240px、右列吃剩余宽度：大图是这一格的主角，页码只是导航。
         * 窄一点的面板（宿主侧栏展开 / 窗口小）用容器查询塌成一列，避免图片被压成一条。
         */
        const PV = {
            grid: {
                display: "grid",
                gridTemplateColumns: "minmax(180px, 240px) minmax(0, 1fr)",
                gap: "14px",
                alignItems: "start"
            },
            left: { minWidth: 0, display: "flex", flexDirection: "column", gap: "6px" },
            right: { minWidth: 0 },
            pageList: { display: "flex", flexDirection: "column", gap: "4px", maxHeight: "560px", overflowY: "auto" },
            pageRow: {
                display: "grid",
                gridTemplateColumns: "34px minmax(0,1fr) auto",
                gridTemplateRows: "auto auto",
                gap: "0 8px",
                padding: "7px 9px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.stroke,
                background: M.surface,
                cursor: "pointer",
                boxShadow: M.soft
            },
            pageRowOn: { borderColor: M.business, background: "color-mix(in srgb, " + M.business + " 9%, " + M.surface + ")" },
            pageNo: Object.assign({ gridRow: "1 / span 2", alignSelf: "center", color: M.business, fontSize: "11px", fontFamily: M.mono }, NUM),
            pageLayout: Object.assign({ color: M.textSecondary, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }, T.tiny),
            pageTitle: Object.assign({ gridColumn: "2", color: M.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }, T.note),
            riskBadge: { gridColumn: "3", gridRow: "1 / span 2", alignSelf: "center", color: M.warn, fontSize: "12px", lineHeight: 1 },
            imageWrap: { display: "flex", flexDirection: "column", gap: "8px" },
            imageMeta: Object.assign({ display: "flex", flexWrap: "wrap", alignItems: "baseline", gap: "10px", color: M.textTertiary }, T.tiny),
            cachedChip: { padding: "1px 6px", borderRadius: M.pill, border: "1px solid " + M.stroke, color: M.textSecondary, background: M.sunken },
            // 文本类文件（第七十一轮）：单列版式。这里不放 maxHeight / overflow —— 面板
            // 本身不滚（滚动归宿主的 scrollBody），正文再套一层定高滚动就是双滚动条，
            // 也是「拉到底闪烁」的一个来源（页脚与 sticky 头会跟着内层滚动抖）。
            textGrid: { display: "flex", flexDirection: "column", gap: "10px", minWidth: 0 },
            textView: { display: "flex", flexDirection: "column", gap: "8px", minWidth: 0 },
            textMeta: Object.assign({ display: "flex", flexWrap: "wrap", alignItems: "baseline", gap: "10px", color: M.textTertiary }, T.tiny),
            textPre: {
                margin: 0,
                padding: "12px 14px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.stroke,
                background: M.sunken,
                color: M.text,
                fontFamily: M.mono,
                fontSize: "12px",
                lineHeight: "18px",
                overflowX: "auto",
                whiteSpace: "pre",
                tabSize: 4
            },
            imagePath: Object.assign({ color: M.caption, overflowWrap: "anywhere" }, T.tiny),
            image: {
                display: "block",
                maxWidth: "100%",
                height: "auto",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.strokeStrong,
                background: "#fff",
                boxShadow: M.lift
            },
            warnings: { marginTop: "14px", display: "flex", flexDirection: "column", gap: "4px" },
            warningsTitle: Object.assign({ color: M.textSecondary, marginBottom: "2px" }, T.note),
            warningRow: {
                display: "grid",
                gridTemplateColumns: "42px minmax(0,1fr)",
                gap: "8px",
                alignItems: "baseline",
                padding: "5px 8px",
                borderRadius: M.radiusSm,
                border: "1px solid " + M.stroke,
                background: M.surface,
                color: M.textTertiary,
                cursor: "pointer"
            },
            warningPage: Object.assign({ color: M.warn, fontFamily: M.mono }, T.tiny),
            warningText: Object.assign({ color: M.textTertiary, lineHeight: 1.5 }, T.tiny),
            warningMore: Object.assign({ color: M.caption, marginTop: "2px" }, T.tiny),
            footer: Object.assign({ display: "flex", flexWrap: "wrap", gap: "10px", justifyContent: "space-between", marginTop: "16px", color: M.caption }, T.tiny),
        };


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

        /**
         * 输入框左侧的「知识库范围」选择器（第十轮）。
         *
         * 只做两件事：显示当前会话生效的范围、把用户选中的虚拟文件夹写回去。
         * **范围判定全在服务端**（src/kb-session-scope.js）：这里选什么、甚至直接改
         * 请求参数，都放不出范围；这个组件挂了也只是少一个入口，检索不会因此变宽。
         *
         * 读写都走宿主的命令通道（/kb-scope），不经任何 cwd-only GET：范围属于
         * **会话**，不属于工作目录。命令代理由 apply 里 ctx.inject(['remote.commands'])
         * 拿到的子上下文提供 —— 宿主客户端上下文对没写进 inject 的服务属性是**抛错**，
         * 直接读 ctx.remote 会让整个 entry 变 failed（浏览器里报
         * `cannot get property "remote.session" without inject`）。
         *
         * 选项来自已有的记忆端点（快照里的 kbTree，人类面板本来就看它），只读目录名，
         * 不带会话范围值；不新增一条「按 cwd 读会话范围」的路由。
         */
        const KB_SCOPE_STATE_LABELS = {
            unset: "未选择 · 禁止检索",
            empty: "范围为空 · 禁止检索",
            corrupt: "范围损坏 · 禁止检索",
            ancestry: "父会话不可用 · 禁止检索",
            origin: "来源不可确认 · 禁止检索",
            cycle: "父链循环 · 禁止检索",
            depth: "父链过深 · 禁止检索",
            "no-session": "无会话 · 禁止检索",
            state: "范围读取失败 · 禁止检索",
            unknown: "范围不可确认 · 禁止检索",
        };

        const KB_SCOPE_STATE_HINTS = {
            unset: "这个会话还没有选择知识库范围，当前禁止检索（不会默认放开全库）。请选一个文件夹；要查全部，请显式选择「全部知识库（不限制）」。",
            empty: "继承到的范围与本会话的记录没有交集，当前能查的目录为空，禁止检索。请在父会话里放宽范围，或在这里重新选择目录。",
            corrupt: "这个会话的知识库范围状态读不出来。为安全起见拒绝检索（不会退回全库）；请重新选择目录，或让管理员检查状态文件。",
            ancestry: "父会话的记录已不可用，无法确认继承范围，当前禁止检索（fail-closed）。",
            origin: "这个会话没有可核对的父会话记录（seeded fork，或没有父链的子代理），无法确认它继承的范围。这个入口不可用；请回到原会话调整范围。",
            cycle: "会话的父链出现循环，无法确认知识库范围，当前禁止检索（fail-closed）。",
            depth: "会话的父链太深，无法确认知识库范围，当前禁止检索（fail-closed）。",
            "no-session": "没有会话上下文，读不到知识库范围。",
            state: "读知识库范围状态失败，当前禁止检索（fail-closed）。",
            unknown: "无法确认知识库范围状态，当前按禁止检索处理（不会退回全库）。",
        };

        function blockedScope(reason) {
            return { kind: "blocked", reason: reason };
        }

        /** 回执只认首行：服务端正文里可能引用“全部知识库”等词，不能被当成状态。 */
        function scopeFirstLine(text) {
            const value = typeof text === "string" ? text : "";
            const cut = value.indexOf("\n");
            return (cut === -1 ? value : value.slice(0, cut)).trim();
        }

        /**
         * 把 /kb-scope 的回执文本解析成稳定的客户端状态。
         *
         * 宿主在 dsh-commands 里会重新打包 CommandResult，只保留 kind / text /
         * sourceEventSeq（dsh-commands/src/index.ts:228-256），所以服务端无法在回执上
         * 多加一个 scope 字段；能跨宿主稳定传递的只有 text。这里**只认首行的已知整句**：
         * 先认拒绝 token，再认“已设为「X」（含子目录）” / “全部知识库”这两类正式回执。
         * 正文里出现「全部知识库（不限制）」「已清空」等字样不算数；不认识的文本一律
         * blocked/unknown，既不放开全库，也不按刚发出的选择猜成功。
         */
        function scopeFromText(text) {
            const line = scopeFirstLine(text);
            if (line === "") return blockedScope("unknown");
            // 1) 拒绝 token 优先：拒绝句里可能引用了“全部知识库（不限制）”作为下一步。
            const refusals = [
                ["ancestry", /^知识库范围[：:]\s*父会话不可用[。，]/],
                ["origin", /^知识库范围[：:]\s*来源不可确认[。，]/],
                ["cycle", /^知识库范围[：:]\s*父链循环[。，]/],
                ["depth", /^知识库范围[：:]\s*父链过深[。，]/],
                ["corrupt", /^知识库范围[：:]\s*状态损坏[。，]/],
                ["empty", /^知识库范围[：:]\s*范围为空[。，]/],
                ["no-session", /^知识库范围[：:]\s*无会话[。，]/],
                ["state", /^知识库范围[：:]\s*状态读取失败[。，]/],
                ["unset", /^知识库范围[：:]\s*未设置[。，]/],
                ["unknown", /^知识库范围[：:]\s*无法确认[。，]/],
                ["unset", /^📚?\s*这个会话还没有设置知识库范围/],
            ];
            for (const entry of refusals) {
                if (entry[1].test(line)) return blockedScope(entry[0]);
            }
            // 2) 正式成功回执：整行锚定；folder 用贪婪捕获，名字里带「」也保留。
            let quoted = /^知识库范围：已设为「(.+)」（含子目录）。$/.exec(line);
            if (quoted === null) quoted = /^知识库范围：「(.+)」（含子目录）。$/.exec(line);
            if (quoted !== null) return { kind: "folder", folder: quoted[1] };
            if (/^知识库范围：全部知识库。$/.test(line)) return { kind: "all" };
            if (/^知识库范围：已清空（全部知识库）。/.test(line)) return { kind: "all" };
            if (/^知识库范围：已清空。?$/.test(line)) return { kind: "all" };
            // 3) 旧后端文案：同样整行锚定，避免正文里的“全部知识库”被误读。
            quoted = /^当前会话的知识库范围已设为「(.+)」及其子目录。$/.exec(line);
            if (quoted === null) quoted = /^当前会话的知识库范围：「(.+)」及其子目录。$/.exec(line);
            if (quoted !== null) return { kind: "folder", folder: quoted[1] };
            if (/^当前会话的知识库范围：全部知识库（不限目录）。$/.test(line)) return { kind: "all" };
            if (/^当前会话的知识库范围已清空：可以检索全部知识库。$/.test(line)) return { kind: "all" };
            return blockedScope("unknown");
        }

        function scopeIsBlocked(state) {
            return state !== null && state !== undefined && state.kind === "blocked";
        }

        function scopeBlockedReason(state) {
            return state && typeof state.reason === "string" ? state.reason : "unknown";
        }

        function scopeLabel(state) {
            if (state === null || state === undefined) return "知识库范围";
            if (state.kind === "all") return "全部知识库（不限制）";
            if (state.kind === "folder") return "文件夹：" + state.folder;
            return KB_SCOPE_STATE_LABELS[scopeBlockedReason(state)] || KB_SCOPE_STATE_LABELS.unknown;
        }

        function scopeDescription(state) {
            if (state === null || state === undefined) return "";
            if (state.kind === "all") return "全部知识库（不限制）";
            if (state.kind === "folder") return "文件夹「" + state.folder + "」（含子目录）";
            return KB_SCOPE_STATE_LABELS[scopeBlockedReason(state)] || KB_SCOPE_STATE_LABELS.unknown;
        }

        function KbScopePicker(props) {
            const sessionId = props && typeof props.sessionId === "string" ? props.sessionId : "";
            const getCommands = props && typeof props.getCommands === "function" ? props.getCommands : () => null;
            // useSessions 是宿主给 session 作用域 slot 的标准 prop（勾子）。它对这个
            // 已挂载的条目是常量，所以「存在才调用」不会让同一个实例的勾子顺序变化。
            const useSessions = props && typeof props.useSessions === "function" ? props.useSessions : null;
            const cwd = useSessions !== null
                ? useSessions((state) => {
                    const row = state && state.byId ? state.byId[sessionId] : null;
                    return row && typeof row.cwd === "string" ? row.cwd : "";
                })
                : "";
            const [ui, setUi] = React.useState({
                open: false, ready: false, state: null, folders: [], error: "", busy: false,
                foldersLoading: false, foldersError: "", foldersLoaded: false,
            });
            // 目录读取的序号：关闭再打开、连续点重试时只认最后一次结果，迟到响应不能
            // 覆盖更新的列表（否则会拿旧列表把刚失效的路径又摆出来）。
            const folderLoadSeq = React.useRef(0);
            // 悬停/聚焦高亮：内联样式写不了 :hover，只能把「现在指着哪一行」记在状态里。
            // 行被禁用时不写它（禁用的按钮本来也收不到鼠标事件），于是禁用态不会被悬停改色。
            const [hovered, setHovered] = React.useState("");
            const hoverProps = (key) => ({
                onMouseEnter: () => setHovered(key),
                onMouseLeave: () => setHovered((prev) => (prev === key ? "" : prev)),
                onFocus: () => setHovered(key),
                onBlur: () => setHovered((prev) => (prev === key ? "" : prev)),
            });

            /** 跑一条 /kb-scope；成功回 { text, state }，失败把原因写进面板并回 null。 */
            const runCommand = async (line) => {
                const commands = getCommands();
                if (commands === null || commands === undefined || typeof commands.execute !== "function") {
                    setUi((prev) => Object.assign({}, prev, { busy: false, error: "命令通道还没就绪，范围暂时改不了。" }));
                    return null;
                }
                let result;
                try {
                    result = await commands.execute(sessionId, line, []);
                } catch (error) {
                    setUi((prev) => Object.assign({}, prev, { busy: false, error: "设置范围失败：" + (error && error.message ? error.message : String(error)) }));
                    return null;
                }
                if (!result || result.ok !== true) {
                    const message = result && result.error ? String(result.error.message || result.error.code) : "未知错误";
                    setUi((prev) => Object.assign({}, prev, { busy: false, error: "设置范围失败：" + message }));
                    return null;
                }
                const value = result.value && result.value.result ? result.value.result : null;
                if (value === null || value.kind !== "success") {
                    setUi((prev) => Object.assign({}, prev, { busy: false, error: value && value.text ? value.text : "设置范围失败。" }));
                    return null;
                }
                const text = typeof value.text === "string" ? value.text : "";
                return { text: text, state: scopeFromText(text) };
            };

            /** 拉一遍文件夹目录（只是目录名，供选择器列出选项）；错误必须回传，不能当空列表。 */
            const fetchFolders = async () => {
                if (typeof fetch !== "function") {
                    return { ok: false, folders: [], error: "当前环境不能读取虚拟文件夹列表。" };
                }
                if (cwd === "") {
                    return { ok: false, folders: [], error: "会话工作目录还没就绪，读不到虚拟文件夹列表。" };
                }
                const url = MEMORY_VIEW_PATH + "?cwd=" + encodeURIComponent(cwd) + "&kblimit=1";
                let response;
                try {
                    response = await fetch(url, { headers: { accept: "application/json" } });
                } catch (error) {
                    return { ok: false, folders: [], error: "读取目录失败：" + (error && error.message ? error.message : String(error)) };
                }
                if (!response || response.ok !== true) {
                    // 403（目录不受信）/ 500 等都要让用户看见；读失败绝不等于“没有目录”。
                    let detail = "";
                    try {
                        const body = response && typeof response.json === "function" ? await response.json() : null;
                        if (body && body.error) detail = String(body.error);
                    } catch (error) {
                        detail = "";
                    }
                    const status = response && response.status ? "HTTP " + response.status : "网络错误";
                    return { ok: false, folders: [], error: "读取目录失败（" + status + "）" + (detail === "" ? "。" : "：" + detail) };
                }
                let body;
                try {
                    body = typeof response.json === "function" ? await response.json() : null;
                } catch (error) {
                    return { ok: false, folders: [], error: "目录接口返回的不是有效 JSON。" };
                }
                if (body === null || typeof body !== "object") {
                    return { ok: false, folders: [], error: "目录接口没有返回内容。" };
                }
                if (body.ok === false) {
                    return { ok: false, folders: [], error: "读取目录失败：" + (body.error ? String(body.error) : "服务端返回错误。") };
                }
                const root = body.kbTree ? body.kbTree : null;
                const folders = [];
                const walk = (node, depth) => {
                    if (node === null || node === undefined) return;
                    if (typeof node.path === "string" && node.path !== "") {
                        folders.push({ path: node.path, name: String(node.name || node.path), depth: depth });
                    }
                    for (const child of node.children || []) walk(child, depth + 1);
                };
                walk(root, 0);
                return { ok: true, folders: folders, error: "" };
            };

            const refreshFolders = async () => {
                const seq = ++folderLoadSeq.current;
                setUi((prev) => Object.assign({}, prev, { foldersLoading: true, foldersError: "" }));
                const result = await fetchFolders();
                if (folderLoadSeq.current !== seq) return result;
                setUi((prev) => Object.assign({}, prev, {
                    foldersLoading: false,
                    // 失败时保留上一份列表只用于展示（行会被禁用）；绝不写成空列表。
                    folders: result.ok ? result.folders : prev.folders,
                    foldersLoaded: result.ok ? true : prev.foldersLoaded,
                    foldersError: result.ok ? "" : result.error,
                }));
                return result;
            };

            const toggle = () => {
                const next = !ui.open;
                // 收起时把高亮清掉：否则下次打开会残留一行“被指着”的底色。
                setHovered("");
                setUi((prev) => Object.assign({}, prev, { open: next, error: "" }));
                if (!next) return;
                if (!ui.ready && !ui.busy) {
                    setUi((prev) => Object.assign({}, prev, { busy: true }));
                    runCommand("/kb-scope").then((reply) => {
                        if (reply === null) return;
                        setUi((prev) => Object.assign({}, prev, { busy: false, ready: true, state: reply.state }));
                    });
                }
                // 每次打开都重新取目录：空目录/旧归档可能在这期间新建或改名，
                // 缓存一份列表会让“刚建好的目录”永远不出现。
                refreshFolders();
            };

            /** 真正把选择写回服务端，并以回读到的 status 为准。 */
            const commitPick = (folder) => {
                const line = folder === "" ? "/kb-scope clear" : "/kb-scope set " + folder;
                runCommand(line).then((reply) => {
                    if (reply === null) return;
                    // 回执只说明“这条命令被接受了”，不证明范围已经是请求的那个值：
                    // 再回读一次服务端 status，以它为准。回读失败时保留原状态并留在面板里。
                    return runCommand("/kb-scope").then((status) => {
                        if (status === null) {
                            setUi((prev) => Object.assign({}, prev, { busy: false, open: true }));
                            return;
                        }
                        const state = status.state;
                        // 服务端给出拒绝态时不要收起面板：用户得看见“为什么没生效”。
                        setUi((prev) => Object.assign({}, prev, {
                            busy: false, ready: true, state: state, open: scopeIsBlocked(state),
                        }));
                    });
                });
            };

            const pick = (folder) => {
                if (ui.busy || ui.foldersLoading) return;
                setUi((prev) => Object.assign({}, prev, { busy: true, error: "" }));
                // 清空不依赖目录列表，直接发。选具体目录前先刷新一次：面板里的行可能
                // 是上一轮读到的，目录刚被删/改名时不能再把旧路径提交给服务端。
                if (folder === "") {
                    commitPick("");
                    return;
                }
                refreshFolders().then((result) => {
                    if (!result.ok) {
                        // 错误已经由 refreshFolders 写进 foldersError；面板保持打开供重试。
                        setUi((prev) => Object.assign({}, prev, { busy: false, open: true }));
                        return;
                    }
                    if (!result.folders.some((item) => item.path === folder)) {
                        setUi((prev) => Object.assign({}, prev, {
                            busy: false, open: true,
                            error: "这个目录刚刚不可用了（可能已被删除或改名）。列表已刷新，请重新选择。",
                        }));
                        return;
                    }
                    commitPick(folder);
                });
            };

            if (sessionId === "") return null;
            const blocked = ui.ready && scopeIsBlocked(ui.state);
            const reason = blocked ? scopeBlockedReason(ui.state) : "unknown";
            const label = "📚 " + scopeLabel(ui.ready ? ui.state : null);
            const stateToken = !ui.ready || ui.state === null || ui.state === undefined
                ? "loading"
                : ui.state.kind === "folder" || ui.state.kind === "all"
                    ? ui.state.kind
                    : scopeBlockedReason(ui.state);
            const rowsDisabled = ui.busy || ui.foldersLoading || ui.foldersError !== "";
            const allSelected = ui.ready && ui.state !== null && ui.state !== undefined && ui.state.kind === "all";
            const allDisabled = ui.busy || ui.foldersLoading;
            const rows = [h("button", Object.assign({
                key: "__all__",
                type: "button",
                "data-kb-scope-all": "",
                "aria-pressed": allSelected ? "true" : "false",
                title: "显式放开全部知识库（不限制）",
                disabled: allDisabled,
                // 选中态的底色/竖条只由 kbScopeRowStyle 给：aria-pressed 之外还要看得见。
                style: kbScopeRowStyle({
                    selected: allSelected,
                    hovered: !allDisabled && hovered === "__all__",
                    disabled: allDisabled,
                }),
                onClick: () => pick(""),
            }, allDisabled ? null : hoverProps("__all__")), "全部知识库（不限制）")];
            for (const item of ui.folders) {
                // 目录恰好叫“全部知识库”时，入口和行都必须和“不限制”区分开，
                // 否则用户会以为自己在放开全库，实际只是锁定了一个同名目录。
                const folderLabel = item.name === "全部知识库" ? "全部知识库（文件夹）" : item.name;
                const selected = ui.ready && ui.state && ui.state.kind === "folder" && ui.state.folder === item.path;
                rows.push(h("button", Object.assign({
                    key: item.path,
                    type: "button",
                    "data-kb-scope-folder": item.path,
                    "aria-pressed": selected ? "true" : "false",
                    // 列表读取失败/正在刷新时不允许拿旧行提交；pick 自己也会再核一次。
                    disabled: rowsDisabled,
                    style: kbScopeRowStyle({
                        selected: selected,
                        hovered: !rowsDisabled && hovered === item.path,
                        disabled: rowsDisabled,
                        depth: item.depth,
                    }),
                    onClick: () => pick(item.path),
                }, rowsDisabled ? null : hoverProps(item.path)), folderLabel));
            }
            const toggleStyle = ui.open
                ? KB_SCOPE_BUTTON_OPEN
                : (hovered === "__toggle__" ? KB_SCOPE_BUTTON_HOVER : KB_SCOPE_BUTTON);
            return h("div", { style: { position: "relative", flex: "0 0 auto" } },
                h("button", Object.assign({
                    type: "button",
                    "data-kb-scope-toggle": "",
                    "data-kb-scope-ready": ui.ready ? "true" : "false",
                    "data-kb-scope-state": stateToken,
                    "data-kb-scope-blocked": blocked ? "true" : "false",
                    "aria-expanded": ui.open,
                    title: blocked
                        ? "当前禁止知识库检索：" + (KB_SCOPE_STATE_LABELS[reason] || KB_SCOPE_STATE_LABELS.unknown)
                        : "当前会话只检索所选虚拟文件夹及其子目录；范围由服务端强制",
                    style: toggleStyle,
                    onClick: toggle,
                }, hoverProps("__toggle__")), label + " ▾"),
                ui.open ? h("div", { "data-kb-scope-panel": "", style: KB_SCOPE_PANEL },
                    ui.ready ? h("div", { "data-kb-scope-current": "", style: KB_SCOPE_NOTE }, "当前：" + scopeDescription(ui.state)) : null,
                    blocked ? h("div", {
                        role: "alert",
                        "data-kb-scope-notice": "",
                        style: kbScopeAlertStyle("warn"),
                    }, KB_SCOPE_STATE_HINTS[reason] || KB_SCOPE_STATE_HINTS.unknown) : null,
                    ui.foldersLoading ? h("div", {
                        role: "status",
                        "data-kb-scope-folders-loading": "",
                        style: KB_SCOPE_NOTE,
                    }, "正在读取虚拟文件夹…") : null,
                    ui.foldersError !== "" ? h("div", {
                        role: "alert",
                        "data-kb-scope-folders-error": "",
                        style: kbScopeAlertStyle("error"),
                    }, ui.foldersError, h("button", {
                        type: "button",
                        "data-kb-scope-folders-retry": "",
                        disabled: ui.busy || ui.foldersLoading,
                        style: Object.assign({}, KB_SCOPE_ROW, { padding: "0 8px", textDecoration: "underline" }),
                        onClick: () => refreshFolders(),
                    }, "重试")) : null,
                    ui.foldersLoaded && !ui.foldersLoading && ui.foldersError === "" && ui.folders.length === 0
                        ? h("div", { "data-kb-scope-folders-empty": "", style: KB_SCOPE_NOTE }, "没有读到可选的虚拟文件夹（目录来自已有文档）。")
                        : null,
                    rows,
                    ui.busy ? h("div", { style: KB_SCOPE_NOTE }, "正在生效…") : null,
                    ui.error !== "" ? h("div", { role: "alert", "data-kb-scope-error": "", style: kbScopeAlertStyle("error") }, ui.error) : null,
                ) : null,
            );
        }

        // 宿主的主题 token：浅色 / 深色两套值由 dsh-client-ui-theme 成对定义在
        // body 与 body[data-ds-dark-theme] 里，这里只取别名，不搬具体色值。
        //
        // 第九十八轮修的是**深色下整块看不见**：面板底色原来写 var(--surface, #fff)，
        // 而宿主从没定义过 --surface（theme bundle 里 0 处定义），于是深色模式下
        // 面板永远是 #fff、文字却继承 --dsw-alias-label-primary（#f9fafb），
        // 实测对比度 1.04:1。同一处还硬编码了 rgba(0,0,0,0.6) 与 #b3261e，
        // 深色下同样不可读；选中 / 悬停 / 禁用三态此前没有任何视觉差别。
        // 现在颜色一律走 --dsw-alias-*（fallback 只兜底“宿主太旧没这个 token”），
        // 深浅两套的实测对比度由 test/kb-scope-contrast.mjs 钉住。
        const KB_TOKENS = {
            text: "var(--dsw-alias-label-primary, #1f2328)",
            note: "var(--dsw-alias-label-secondary, #57606a)",
            muted: "var(--dsw-alias-label-tertiary, #6e7781)",
            surface: "var(--dsw-alias-bg-layer-1, #ffffff)",
            border: "var(--dsw-alias-border-l2, rgba(0,0,0,0.16))",
            hover: "var(--dsw-alias-interactive-bg-hover, rgba(38,49,72,0.06))",
            active: "var(--dsw-alias-interactive-bg-active, rgba(38,49,72,0.10))",
            accent: "var(--dsw-alias-state-business-primary, #4176e6)",
            error: "var(--dsw-alias-state-error-primary, #b3261e)",
            warn: "var(--dsw-alias-state-warn-primary, #b45309)",
            shadow: "var(--dsw-alias-bg-mask-2, rgba(0,0,0,0.12))",
        };

        const KB_SCOPE_BUTTON = {
            font: "inherit", fontSize: "12px", lineHeight: "20px", padding: "2px 8px",
            border: "1px solid " + KB_TOKENS.border, borderRadius: "10px",
            background: "transparent", color: KB_TOKENS.text, cursor: "pointer", whiteSpace: "nowrap",
            maxWidth: "220px", overflow: "hidden", textOverflow: "ellipsis",
        };
        /** 入口按钮的悬停态与展开态：展开时重一档，和「点开了」对得上。 */
        const KB_SCOPE_BUTTON_HOVER = Object.assign({}, KB_SCOPE_BUTTON, { background: KB_TOKENS.hover });
        const KB_SCOPE_BUTTON_OPEN = Object.assign({}, KB_SCOPE_BUTTON, { background: KB_TOKENS.active });

        const KB_SCOPE_PANEL = {
            position: "absolute", left: "0", bottom: "calc(100% + 6px)", zIndex: 30,
            minWidth: "200px", maxHeight: "280px", overflowY: "auto",
            padding: "6px", borderRadius: "10px", border: "1px solid " + KB_TOKENS.border,
            // 不透明主题底色 + 主题文字色：深色下不会再出现“白底浅字”的一整块。
            background: KB_TOKENS.surface, color: KB_TOKENS.text,
            boxShadow: "0 6px 24px " + KB_TOKENS.shadow,
            display: "flex", flexDirection: "column", gap: "2px",
        };

        /**
         * 一行目录的样式（选中 / 悬停 / 禁用三态必须一眼分得开）。
         *
         * 选中 = active 底 + 加粗 + 左侧 2px 业务色竖条；悬停 = hover 底；
         * 禁用 = 三级文字色 + not-allowed，且不再响应悬停（调用方不给 hoverProps）。
         */
        function kbScopeRowStyle(options) {
            const opts = options || {};
            const style = {
                font: "inherit", fontSize: "12px", lineHeight: "24px", textAlign: "left",
                padding: "0 10px", border: "0", borderRadius: "6px",
                background: "transparent", color: KB_TOKENS.text, cursor: "pointer", whiteSpace: "nowrap",
            };
            if (opts.depth) style.paddingLeft = (10 + opts.depth * 12) + "px";
            if (opts.disabled) {
                style.color = KB_TOKENS.muted;
                style.cursor = "not-allowed";
                return style;
            }
            if (opts.selected) {
                style.background = KB_TOKENS.active;
                style.fontWeight = "600";
                style.boxShadow = "inset 2px 0 0 " + KB_TOKENS.accent;
                return style;
            }
            if (opts.hovered === true) style.background = KB_TOKENS.hover;
            return style;
        }

        const KB_SCOPE_ROW = kbScopeRowStyle({});
        const KB_SCOPE_NOTE = { fontSize: "11px", lineHeight: "16px", padding: "4px 10px", color: KB_TOKENS.note, whiteSpace: "normal" };

        /**
         * 提示条（禁止检索 / 读目录失败 / 命令失败）。
         *
         * 文字仍用主文字色 —— 两种主题下都 ≥15:1；红/黄只放在左侧 3px 竖条上做语义提示
         * （装饰件，不承载信息）。不拿它当小字正文：实测竖条自身对比度 error 4.50/4.77、
         * warn 2.15/7.31（浅/深），浅色下黄字当正文是读不出来的。
         */
        function kbScopeAlertStyle(tone) {
            return {
                fontSize: "11px", lineHeight: "16px", padding: "6px 10px", borderRadius: "6px",
                whiteSpace: "normal", color: KB_TOKENS.text, background: KB_TOKENS.hover,
                borderLeft: "3px solid " + (tone === "error" ? KB_TOKENS.error : KB_TOKENS.warn),
            };
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
                // 写入用的直连远程通道（第九十一轮）：configForms 那条只回 true/false，
                // 服务端原因全被丢掉（见 formToScope 的注释）。remote 在宿主客户端里叫
                // "remote"（dsh-api-remotes 导出 inject = ['remote']）。**单独一段 inject**：
                // 和 configForms 写在一起的话，部署里没有 remote 服务时这段回调根本不跑，
                // 设置页连读都读不到 —— 读依赖 configForms，写才额外依赖 remote。
                try {
                    ctx.inject(["remote"], (rctx) => {
                        // 模型清单要用整份 remote（session.modelCatalog）；settings
                        // 只服务写入。两者一起记，迟绑定时选择器会据此刷新。
                        try {
                            setHostRemote(rctx && rctx.remote ? rctx.remote : null);
                        } catch (error) {
                            setHostRemote(null);
                        }
                        // 这一份只是兜底：remote 就绪时 settings 可能还没挂上
                        // （它排在 $mount 列表第 5 位），所以「没有」不是定论 ——
                        // 写入时一律走 currentSettingsRemote() 现取。
                        try {
                            settingsRemote = rctx && rctx.remote && rctx.remote.settings ? rctx.remote.settings : null;
                        } catch (error) {
                            settingsRemote = null;
                        }
                    });
                } catch (error) {
                    settingsRemote = null;
                    setHostRemote(null);
                }

                // 直连通道是**独立的可选服务**：只有把它写进注入列表，cordis 才会等到
                // 它挂上再回调（宿主设置面自己也是这么声明的：
                // dsh-client-ui-settings/src/client/index.ts:31）。不声明的话，回调在
                // remote 就绪时就跑完了，那时 remote.settings 往往还是 undefined，
                // 之后再也不会重跑 —— 直连通道等于从未生效（现场那句固定文案的根因）。
                // 宿主没有这个命名空间时回调不跑，写入照旧退回 configForms。
                try {
                    ctx.inject(["remote", "remote.settings"], (rctx) => {
                        let nested = null;
                        try {
                            nested = rctx && rctx.remote && rctx.remote.settings ? rctx.remote.settings : null;
                        } catch (error) {
                            // 拿不到就保持现状：写入时还会现取一次。
                        }
                        if (nested !== null) settingsRemote = nested;
                        try {
                            setHostRemote(rctx && rctx.remote ? rctx.remote : null);
                        } catch (error) {
                            setHostRemote(null);
                        }
                        return () => {
                            if (settingsRemote === nested) settingsRemote = null;
                        };
                    });
                } catch (error) {
                    // 旧宿主不认识嵌套服务名：保持 configForms 兜底，不清掉已有引用。
                }

                // 会话模型清单同样挂在**独立的 remote.session** 上。旧实现从
                // ['remote'] 的回调里读 hostRemote.session：严格 Cordis 下当场抛
                // 「cannot get property "remote.session" without inject」，模型选择器
                // 显示 cordis 错误、RPC 永远不发出（.office/runtime/diag/web-verify.json:1138）。
                // 显式声明 remote.session：宿主挂上时回调才跑，卸载时用 disposer 清掉，
                // 避免重新挂载后继续拿旧代理。旧宿主没有这个命名空间 → 回调不跑，
                // sessionModelCatalog() 返回 null，照旧退回只读 HTTP 清单。
                try {
                    ctx.inject(["remote", "remote.session"], (rctx) => {
                        let nested = null;
                        try {
                            nested = rctx && rctx.remote && rctx.remote.session ? rctx.remote.session : null;
                        } catch (error) {
                            nested = null;
                        }
                        setSessionRemote(nested);
                        return () => {
                            if (sessionRemote === nested) setSessionRemote(null);
                        };
                    });
                } catch (error) {
                    // 旧宿主不认识嵌套服务名：保持 HTTP 兜底。
                }
            }

            // 命令通道（/kb-scope）：**必须显式注入嵌套服务名 remote.commands**，
            // 回调拿到的是宿主给这个子上下文的 remote；从它上面取 commands 才是
            // 有授权的代理。只注入 remote 再读 remote.commands 在严格 Cordis 下会抛
            // 「cannot get property ... without inject」整页 failed（现场实测）。
            // 它是可选服务：没有命令通道的部署里选择器只显示不可用的提示。
            let commandsProxy = null;
            if (typeof ctx.inject === "function") {
                try {
                    ctx.inject(["remote", "remote.commands"], (rctx) => {
                        try {
                            commandsProxy = rctx && rctx.remote && rctx.remote.commands ? rctx.remote.commands : null;
                        } catch (error) {
                            commandsProxy = null;
                        }
                    });
                } catch (error) {
                    commandsProxy = null;
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

            // 资源管理器接口（第八十八轮）：知识库目录选择必须走宿主自己的目录选择器，
            // 而不是在插件里再实现一个文件浏览器 —— 宿主的 picker 有权限、有最近目录，
            // 也跟用户对「选择目录」的预期一致。它同样是可选服务：拿不到时设置页
            // 保留手工输入路径，不会因为一个部署没有 picker 就白屏。
            HOST_PICK_DIRECTORY = null;
            if (typeof ctx.inject === "function") {
                try {
                    ctx.inject(["uiWorkspace"], (sctx) => {
                        const workspace = sctx && sctx.uiWorkspace ? sctx.uiWorkspace : null;
                        if (workspace && typeof workspace.pickDirectory === "function") {
                            HOST_PICK_DIRECTORY = () => workspace.pickDirectory();
                        }
                    });
                } catch (error) {
                    HOST_PICK_DIRECTORY = null;
                }
            }

            // 第一百零五轮起只留一个一级设置页：办公与记忆合并，内部用页签分成
            // 总览 / 办公 / 检索 / 文档 / 记忆 / 知识库 / 召回。两个旧的 section
            // 不能再注册 —— 它们会各占一格导航，合并就落空了。
            ctx.slots.inject("settings.section", () =>
                ctx.slots.register(
                    {
                        name: "settings.section",
                        id: NAMESPACE,
                        order: 519,
                        label: () => h(OfficeNavLabel, null),
                    },
                    (view) => h(MergedSettingsSection, {
                        scope: scope,
                        initialTab: view && typeof view.initialTab === "string" ? view.initialTab : undefined,
                    }),
                ),
            );

            // ── 记忆浏览面板：会话页签行里的一格（对话｜轨迹｜上下文｜编队｜记忆） ──
            //
            // 第六十二轮（新-22）把它从「sidebar.panellist 一格 + main 一格」迁到
            // conversation.view：用户要的是「从插件和技能中心的中间，挪到上下文和
            // 编队的右侧」，而 `上下文`（order 20）与 `编队`（order 30）本身是
            // conversation.view 的两格，不是 sidebar.panellist 的条目 —— 所以
            // 只有注册进同一个 slot 的 order 40，才是字面上的「在它们右侧」。
            //
            // 作用域提醒：conversation.view 是 **session 作用域**，与上下文 / 编队
            // 一致 —— 没有打开会话时这排页签根本不存在，记忆面板也就没有入口
            // （这是这次迁移的已知代价，写进了文档，不是 bug）。
            //
            // 渲染的是同一份 MemoryBrowserPanel：中间那一列还是那一列（点它的前后，
            // 中心列都是同一个宽度），所以面板本身的观感不变。
            ctx.slots.inject("conversation.view", () =>
                ctx.slots.register(
                    { name: "conversation.view", id: MEMORY_PANEL_ID, order: 40, label: () => "记忆" },
                    // owner props（inspectCall / viewRequest / openView / completeViewRequest）
                    // 由宿主逐次传入；openView 用来把「切到会话」变成「切回对话这一格」。
                    //
                    // 第六十四轮补两个：`sessionId` 与 `useSessions` —— 这一格是 session
                    // 作用域，宿主给的标准 props 里有它们（dsh 自己的 ConversationContent
                    // 就用 `useSessions((s) => s.byId[sessionId]?.cwd)` 取当前工作目录）。
                    // 面板拿它把**默认根**定成「我现在在哪个工作区」，不再由宿主挑第一个。
                    (view) => h(MemoryBrowserPanel, {
                        layout: layout,
                        openView: view && typeof view.openView === "function" ? view.openView : null,
                        sessionId: view && typeof view.sessionId === "string" ? view.sessionId : "",
                        useSessions: view && typeof view.useSessions === "function" ? view.useSessions : null,
                    }),
                ),
            );

            // ── 预览面板：会话页签行里再插一格，order 50（紧挨记忆右侧） ──
            //
            // 与记忆同一套 session 作用域 props：sessionId / useSessions 用来把
            // 根定成「当前会话的 cwd」。面板自己不猜工作目录 —— 猜错就会列出别人
            // 目录里的文件，而预览图看起来又完全正常，这种错没人能一眼发现。
            ctx.slots.inject("conversation.view", () =>
                ctx.slots.register(
                    { name: "conversation.view", id: PREVIEW_PANEL_ID, order: 50, label: () => "预览" },
                    (view) => h(PreviewBrowserPanel, {
                        sessionId: view && typeof view.sessionId === "string" ? view.sessionId : "",
                        useSessions: view && typeof view.useSessions === "function" ? view.useSessions : null,
                    }),
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

            // ── 输入框左侧的知识库范围选择器（session 作用域） ──
            //
            // conversation.input.left 是 list / session 作用域（宿主
            // dsh-client-ui-conversation/src/client/apply.ts:472 注册），没有会话时
            // 整条不渲染。范围读写都走 /kb-scope 命令，命令代理由上面那段
            // ctx.inject(['remote.commands']) 提供（不依赖未注入的 hostRemote）。
            ctx.slots.inject("conversation.input.left", () =>
                ctx.slots.register(
                    { name: "conversation.input.left", id: "office-kb-scope", order: 25 },
                    (bar) => h(KbScopePicker, {
                        key: bar && typeof bar.sessionId === "string" ? bar.sessionId : "",
                        sessionId: bar && typeof bar.sessionId === "string" ? bar.sessionId : "",
                        useSessions: bar && typeof bar.useSessions === "function" ? bar.useSessions : null,
                        getCommands: () => commandsProxy,
                    }),
                ),
            );
        }

        exports.apply = apply;
        exports.inject = inject;
        return module.exports;
    },
});