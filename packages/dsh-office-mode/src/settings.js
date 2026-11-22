/**
 * 设置界面（Settings > 插件 > 办公模式）。
 *
 * 设计取向：插件在别人的部署里跑，用户不该为了关一个工具或改一个超时去
 * 编辑 YAML。这里把「能安全调节的东西」列成一张 schema，交给宿主自带的
 * 设置界面渲染，插件自己不写前端。
 *
 * 三条约束：
 *   1. **工具开关只管声明，不管实现**。关掉某个工具只是不注册它，
 *      已经写进会话历史里的调用不会因此失效。
 *   2. **参数只在安全区间内可调**。超时、并发数、结果上限都给上下界，
 *      设置页调到越界会被 schema 挡下，不会让插件在运行时崩。
 *   3. **空配置永远可用**。schema 的默认值 = 不设任何东西时的行为，
 *      与 config.js 的既有默认保持一致。
 *
 * @module dsh-office-mode/settings
 */
/**
 * schemastery **不是**本插件的依赖：它由宿主提供，且只在 profile 的
 * node_modules 里可解析。本插件是 link: 进 profile 的（源码在仓库里），
 * 直接 `import '@deepseek-ai/schemastery'` 会 ERR_MODULE_NOT_FOUND。
 * 所以 schema 用工厂函数造，z 由调用方（apply 时从宿主拿到）传进来；
 * 拿不到就不注册设置页，插件其余功能照常。
 */
import { DEFAULT_SEARCH_DIR, DEFAULT_MAX_PARALLEL } from './search.js';
import { DEFAULT_KB_LLM } from './kb-llm.js';
import {
    DEFAULT_MEMORY,
    DEFAULT_SCRIPT_TIMEOUT_MS,
    DEFAULT_MAX_SCRIPT_CHARS,
    DEFAULT_CACHE_TTL_MINUTES,
    DEFAULT_PDF_DPI,
    DEFAULT_PDF_MAX_PAGES,
    DEFAULT_PDF_ENGINE,
    DEFAULT_TEX_ENGINE,
    DEFAULT_TEX_TIMEOUT_MS,
    DEFAULT_PYTHON,
    DEFAULT_AV,
    DEFAULT_SUBAGENT_MODEL,
    DEFAULT_GUARD,
    DEFAULT_FETCH_WEB,
    DEFAULT_SEARCH,
    SEARCH_ENGINES,
} from './config.js';
import { AV_LANGUAGES } from './av.js';
import { DEFAULT_OCR, OCR_PURPOSES } from './ocr.js';
import { CACHE_DIR } from './office-paths.js';
import { DEFAULT_LAYERS, DEFAULT_RECALL_QUALITY, DEFAULT_QUOTA, MEMORY_SCOPES, USER_SCOPES } from './memory.js';
import { PROVIDER_IDS } from './web-providers.js';
import { SITE_PRIORITY_LIMITS, SITE_TYPE_IDS, siteCatalog } from './site-catalog.js';
import { DEFAULT_THEME_ID, THEMES } from './engine/theme.js';

/** 设置命名空间。必须是小写字母开头、只含小写字母/数字/连字符。 */
export const OFFICE_SETTINGS_NS = 'dsh-office-mode';

/** 八个工具的名字，供开关与提示词共用。 */
export const TOOL_NAMES = ['office_help', 'office_run', 'office_memory', 'office_web_search', 'office_web_fetch', 'office_search_run', 'office_search_dispatch', 'office_parse_findings'];

/** 可调参数的上下界（设置页与校验共用同一份，避免两处漂移）。 */
export const LIMITS = {
    scriptTimeoutMs: { min: 1_000, max: 600_000, step: 1_000 },
    maxScriptChars: { min: 1_000, max: 4_000_000, step: 1_000 },
    cacheTtlMinutes: { min: 0, max: 43_200, step: 10 },
    pdfDpi: { min: 36, max: 400, step: 1 },
    pdfMaxPages: { min: 1, max: 500, step: 1 },
    texTimeoutMs: { min: 30_000, max: 1_800_000, step: 10_000 },
    pythonTimeoutMs: { min: 5_000, max: 900_000, step: 5_000 },
    avThreads: { min: 1, max: 16, step: 1 },
    avChunkSeconds: { min: 10, max: 600, step: 10 },
    avOverlapSeconds: { min: 0, max: 30, step: 0.5 },
    avMaxSeconds: { min: 10, max: 21_600, step: 10 },
    avTimeoutMs: { min: 5_000, max: 1_800_000, step: 5_000 },
    ocrTimeoutMs: { min: 10_000, max: 300_000, step: 5_000 },
    avMaxFrames: { min: 1, max: 60, step: 1 },
    avFrames: { min: 1, max: 60, step: 1 },
    avFrameWidth: { min: 0, max: 4_096, step: 2 },
    searchMaxParallel: { min: 1, max: 8, step: 1 },
    searchBriefMaxChannels: { min: 1, max: 12, step: 1 },
    searchResultLimit: { min: 1, max: 40, step: 1 },
    searchTimeoutMs: { min: 5_000, max: 300_000, step: 5_000 },
    searchFetchTimeoutMs: { min: 5_000, max: 300_000, step: 5_000 },
    searchFetchMaxResults: { min: 1, max: 20, step: 1 },
    searchFetchPages: { min: 0, max: 5, step: 1 },
    searchFetchMaxBytes: { min: 32_768, max: 33_554_432, step: 32_768 },
    searchFetchMaxChars: { min: 1_000, max: 400_000, step: 1_000 },
    searchFetchMaxRedirects: { min: 0, max: 10, step: 1 },
    // 第二十九轮：取正文撞上 PDF 时的体积上限（放弃线，不是截断线）。
    searchFetchPdfMaxBytes: { min: 1_048_576, max: 134_217_728, step: 1_048_576 },
    memoryUserBytes: { min: 512, max: 65_536, step: 512 },
    memoryProjectBytes: { min: 1_024, max: 262_144, step: 1_024 },
    // 投影预算：0 = 不限（第四十八轮 P0-2）
    memoryProjectionBytes: { min: 0, max: 131_072, step: 1_024 },
    // 单条上限与建议线：0 = 关掉（第四十八轮 新-11）
    memoryEntryBytes: { min: 0, max: 131_072, step: 100 },
    memoryLedgerLimit: { min: 10, max: 10_000, step: 10 },
    memoryArchiveKeep: { min: 1, max: 1_000, step: 1 },
    recallLowScore: { min: 0, max: 1, step: 0.05 },
    recallHighScore: { min: 0, max: 1, step: 0.05 },
    recallCandidateMultiplier: { min: 1, max: 10, step: 1 },
    recallMaxMedium: { min: 0, max: 40, step: 1 },
    recallMaxUnknown: { min: 0, max: 40, step: 1 },
    quotaPerTurn: { min: 0, max: 50, step: 1 },
    // 知识库检索（第八十五轮）：权重 0–1，候选池与 topN 有界，超时防止一次检索拖死一回合。
    kbWeight: { min: 0, max: 1, step: 0.05 },
    kbCandidateLimit: { min: 5, max: 100, step: 5 },
    kbEmbeddingDimensions: { min: 0, max: 4_096, step: 1 },
    kbEmbeddingBatch: { min: 1, max: 64, step: 1 },
    kbEmbeddingTimeoutMs: { min: 1_000, max: 120_000, step: 1_000 },
    // 本地推理（第一百零三轮）：内存是有硬上限的，超了看门狗直接杀进程
    kbLocalIdleMs: { min: 5_000, max: 3_600_000, step: 5_000 },
    kbLocalMaxRss: { min: 200, max: 4_096, step: 10 },
    kbLocalStartupMs: { min: 5_000, max: 300_000, step: 5_000 },
    kbLocalMaxLength: { min: 64, max: 512, step: 32 },
    kbLocalThreads: { min: 1, max: 16, step: 1 },
    // 候选筛选（替代重排）：只发编号 + 截断块头，只回编号
    kbSelectorCandidates: { min: 2, max: 30, step: 1 },
    kbSelectorMaxSelected: { min: 1, max: 20, step: 1 },
    kbSelectorHeadChars: { min: 80, max: 2_000, step: 20 },
    kbSelectorTimeoutMs: { min: 1_000, max: 120_000, step: 1_000 },
    // ANN 索引：M / efConstruction 建图时定死，efSearch 每次检索可调
    kbAnnM: { min: 4, max: 64, step: 4 },
    kbAnnEfConstruction: { min: 16, max: 1_000, step: 8 },
    kbAnnEfSearch: { min: 8, max: 512, step: 8 },
    kbAnnMaxBuild: { min: 100, max: 20_000, step: 100 },
};

/** 主题 id 列表（设置页渲染成下拉）。 */
export const THEME_IDS = THEMES.map((theme) => theme.id);

/** 通道参数里每个键的说明（各条通道共用一套措辞）。 */
const PROVIDER_FIELD_NOTE = {
    apiKey: '这条通道的 Key 字面量（一般留空，用下面的环境变量名；填了以它为准）',
    apiKeyEnv: '取 Key 的环境变量名（先问宿主凭据服务、再看环境变量、最后看 $DSH_HOME/.credentials.yaml）',
    baseURL: '这条通道的端点（末尾斜杠会被去掉）',
    model: '这条通道用的模型（要支持它的联网检索能力）',
};

/**
 * 把每一个可编辑的叶子字段标成 volatile。
 *
 * 为什么必须标：0.1.7 起设置页的写入只接受 volatile 路径 —— dsh-settings 的
 * `write()` 会拿 `isVolatilePath()` 逐条校验，非 volatile 的字段直接报
 * 「Config field "x" is not volatile」。而 volatile 字段的解析结果是
 * Volatile 引用、由宿主就地更新，插件不必重挂就能读到新值 —— 这正是
 * 「改完即时生效」的依据（旧版靠 settings.onChange 回调，那条路已经没有了）。
 *
 * 为什么整片置位而不是逐个写 `.volatile()`：schema 有四十多个字段，逐个写既
 * 啰嗦又容易漏掉一个 —— 而漏掉的那一刻就是设置页上一个点不动的开关。
 * 运行期 volatile 就是 `meta` 上的一个布尔位（schemastery 的 `volatile()` 实现
 * 就是 `this.extra('volatile', true)`），对每个叶子置位与逐个调用等价。
 *
 * 约束：volatile 字段不能嵌在另一个 volatile 字段里
 * （`validateVolatileSchema` 会抛「volatile fields require a fixed object path」），
 * 所以只标叶子，容器对象（tools / search / memory / documents / subagentModel
 * 以及 layers / recallQuality / quota）一律不标。
 *
 * @param {object} schema schemastery 造出来的对象 schema
 */
function markLeavesVolatile(schema) {
    // schema 是可调用的（`schema({})` 就是解析），所以 typeof 是 'function' 而不是
    // 'object' —— 只判 'object' 会让整个函数在第一行静默返回。
    if (schema === null || (typeof schema !== 'object' && typeof schema !== 'function')) return;
    // 容器：只往下走，自己不标。
    if (schema.type === 'object' && schema.dict !== undefined) {
        for (const key of Object.keys(schema.dict)) markLeavesVolatile(schema.dict[key]);
        return;
    }
    if (schema.meta !== undefined && schema.meta.volatile !== true) {
        schema.meta.volatile = true;
    }
}

/**
 * 造设置 schema。字段的 description 会直接显示在设置界面里，所以要写成人话。
 *
 * 分六组：工具开关 / 检索编排 / 记忆 / 后台任务模型 / 文档与缓存 / Python 计算与绘图。
 *
 * @param {object} z schemastery 实例（由宿主提供）
 */
export function createOfficeSettings(z) {
    if (z === undefined || z === null || typeof z.object !== 'function') {
        throw new Error('createOfficeSettings 需要一个 schemastery 实例。');
    }
    const schema = z.object({
    // ── 工具开关 ──────────────────────────────────────────────────────────
    tools: z.object({
        office_help: z.boolean()
            .description('办公文档的按需用法查询（Word / Excel / PPT 写法、主题、检索说明）')
            .default(true),
        office_run: z.boolean()
            .description('批量执行：生成或修改 Word / Excel / PPT，以及批量改文本文件')
            .default(true),
        office_memory: z.boolean()
            .description('办公记忆：读写热记忆（偏好与约定）、检索台账与归档；office_run 的自动台账也归它管')
            .default(true),
        office_web_search: z.boolean()
            .description('查一个事实点（网页抓取通道，免 Key）：给几条查询，抓 DuckDuckGo HTML 解析，来源直接回到对话')
            .default(true),
        office_web_fetch: z.boolean()
            .description('打开一个页面取正文（网页抓取通道）：自带安全抓取，来源是 PDF 时交给 office.pdf 抽文本')
            .default(true),
        office_search_run: z.boolean()
            .description('直接查一轮（内置检索）：给几条查询，插件自己的联网通道去找来源并落盘。不起子代理、不需要提纲')
            .default(true),
        office_search_dispatch: z.boolean()
            .description('渠道覆盖检索：给 topic 自动出提纲（brief.md 落盘）并逐渠道执行；有子代理就派子代理，否则自动用内置检索')
            .default(true),
        office_parse_findings: z.boolean()
            .description('检索第三步：把子代理写下的结果文件读成紧凑摘要')
            .default(true),
    })
        .description('按需启用工具。关掉的工具不会出现在模型面前，也不占每次请求的 schema 开销')
        .default({}),

    // ── 检索编排 ──────────────────────────────────────────────────────────
    search: z.object({
        maxParallel: z.number()
            .description('同时铺开几个检索子代理。渠道多时不要一次全开，容易被限流')
            .min(LIMITS.searchMaxParallel.min).max(LIMITS.searchMaxParallel.max).step(LIMITS.searchMaxParallel.step)
            .default(DEFAULT_MAX_PARALLEL),
        resultLimit: z.number()
            .description('解析结果时最多列出多少条结论（摘要里超出部分只报总数）')
            .min(LIMITS.searchResultLimit.min).max(LIMITS.searchResultLimit.max).step(LIMITS.searchResultLimit.step)
            .default(40),
        maxChannels: z.number()
            .description('一次提纲最多列几个渠道，防止渠道爆炸式增长')
            .min(LIMITS.searchBriefMaxChannels.min).max(LIMITS.searchBriefMaxChannels.max).step(LIMITS.searchBriefMaxChannels.step)
            .default(12),
        outputDir: z.string()
            .description('检索提纲与结果文件的目录（相对会话工作目录）')
            .default(DEFAULT_SEARCH_DIR),
        engine: z.union(SEARCH_ENGINES.map((id) => z.const(id)))
            .description('检索的执行引擎：auto = 先派子代理、跑不了就用插件内置检索；subagent = 只用子代理；builtin = 只用内置检索。'
                + '子代理用插件的抓取工具（office_web_search / office_web_fetch），auto 在派工跑不了时自动走内置检索')
            .default(DEFAULT_SEARCH.engine),
        provider: z.union(['auto', ...PROVIDER_IDS].map((id) => z.const(id)))
            .description('用哪条联网通道。auto = 按下面的顺序依次试（第一条成功的就用它）；'
                + '固定成某一条时只走那一条。默认顺序是 deepseek-api（直连 Search API）→ 宿主 web 服务 → 免 Key 抓取；'
                + '三方检索 API（openai / tavily / bocha / serper …）要显式点名或加回顺序')
            .default(DEFAULT_SEARCH.provider),
        providerOrder: z.array(z.union(PROVIDER_IDS.map((id) => z.const(id))))
            .description('auto 的尝试顺序。默认 deepseek-api → seam → duckduckgo → searxng：'
                + '先直连 Search API（快、不跑辅助模型），宿主接缝第二手，免 Key 抓取兜底；'
                + '三方检索 API 不进默认顺序，加回来才会被自动尝试')
            .default([...DEFAULT_SEARCH.providerOrder]),
        proxy: z.string()
            .description('出口代理（第二十九轮新增），形如 http://127.0.0.1:7897；留空 = 用本进程原有出口直连。'
                + '优先只影响本插件（宿主装了 undici 时建私有分派器，改值当场生效）；拿不到 undici 才退回 Node 的 '
                + 'http.setGlobalProxyFromEnv()，那条是**进程级**且**只能装不能清**（清空后要重启 DSH 才回直连）。'
                + '本地地址（127.0.0.1 / localhost）两条路都不经代理；只认 http / https 代理')
            .default(DEFAULT_SEARCH.proxy),
        sites: z.object({
            enabled: z.boolean()
                .description('站点优先检索（第三十四轮新增）：开启后先把查询限定到清单里的站点（site:<域名>），'
                    + '命中的来源排在前面并按类型汇总；关掉就完全按原来的路子走')
                .default(DEFAULT_SEARCH.sites.enabled),
            fallback: z.boolean()
                .description('限定站点一无所获时退回不限定来源的泛搜，并在反馈里点名哪些站点空手 —— '
                    + '站点被墙或没收录时不至于把「清单查不到」当成「没有资料」')
                .default(DEFAULT_SEARCH.sites.fallback),
            maxPerCall: z.number()
                .description('单次检索最多把几条查询限定到站点上（限定会多花请求，默认只 4 条）')
                .min(SITE_PRIORITY_LIMITS.maxPerCall[0]).max(SITE_PRIORITY_LIMITS.maxPerCall[1])
                .step(1).default(DEFAULT_SEARCH.sites.maxPerCall),
            entries: z.array(z.object({
                type: z.union(SITE_TYPE_IDS.map((id) => z.const(id)))
                    .description('类型：academic 学术 / book 图书 / code 代码 / custom 自定义').default('custom'),
                domain: z.string().description('域名，不带协议与路径（限定查询拼成 site:<域名>）').default(''),
                label: z.string().description('显示名（设置页与反馈里用它）').default(''),
                note: z.string().description('一句话说明').default(''),
                enabled: z.boolean().description('是否参与优先检索').default(true),
            }))
                .description('站点清单本体（第三十四轮新增）。默认值就是内置目录；把行删空 = 这次不限定任何站点。'
                    + '影子图书馆类（sci-hub / z-lib / Anna\'s Archive）只进目录、默认关闭')
                .default(siteCatalog()),
        }).description('站点清单与站点优先检索（第三十四轮新增）')
            .default({}),
        providers: z.object(Object.fromEntries(PROVIDER_IDS.map((id) => [id, z.object(Object.fromEntries(
            Object.entries(DEFAULT_SEARCH.providers[id] ?? {}).map(([key, value]) => [key, key === 'timeoutMs'
                ? z.number().description('这条通道单次检索的超时（毫秒）')
                    .min(LIMITS.searchTimeoutMs.min).max(LIMITS.searchTimeoutMs.max).step(LIMITS.searchTimeoutMs.step)
                    .default(value)
                : z.string().description(PROVIDER_FIELD_NOTE[key] ?? key).default(value)]),
        )).default({})])))
            .description('每条通道各自的 Key / 端点 / 模型 / 超时。deepseek-api 那条的 Key 默认取 DEEPSEEK_API_KEY，端点是 Search API 的基础地址（末尾拼 /search）')
            .default({}),
        preprocess: z.object({
            mode: z.union(['article', 'plain', 'off'].map((id) => z.const(id)))
                .description('取回来的网页先做预处理：article = 去导航与样板、挑主容器；plain = 只做行级清洗；off = 原样返回')
                .default(DEFAULT_SEARCH.preprocess.mode),
            dropBoilerplate: z.boolean()
                .description('丢掉 cookie 条 / 登录注册 / 订阅分享 / 版权声明这类样板行')
                .default(DEFAULT_SEARCH.preprocess.dropBoilerplate),
            dedupeLines: z.boolean()
                .description('同一条重复出现的行只留第一次（导航条收成文本后最常撞这条）')
                .default(DEFAULT_SEARCH.preprocess.dedupeLines),
            minLineChars: z.number()
                .description('短于这个长度、又不像句子的行当导航丢掉。0 = 关掉这条（清单型页面建议关）')
                .min(0).max(80).step(1)
                .default(DEFAULT_SEARCH.preprocess.minLineChars),
            keepTitle: z.boolean()
                .description('抽出标题 / 作者 / 发布时间留给报告（不进正文）')
                .default(DEFAULT_SEARCH.preprocess.keepTitle),
        })
            .description('网页预处理：取正文的产物先清洗再进上下文，报告里给出「洗掉多少」')
            .default({}),
        fetch: z.object({
            maxResults: z.number()
                .description('一次检索最多要几条来源（每个渠道 / 每条直查查询）')
                .min(LIMITS.searchFetchMaxResults.min).max(LIMITS.searchFetchMaxResults.max).step(LIMITS.searchFetchMaxResults.step)
                .default(DEFAULT_FETCH_WEB.maxResults),
            fetchPages: z.number()
                .description('每个渠道 / 每条查询最多打开几页取正文摘录。0 = 只用搜索返回的标题与引用片段')
                .min(LIMITS.searchFetchPages.min).max(LIMITS.searchFetchPages.max).step(LIMITS.searchFetchPages.step)
                .default(DEFAULT_FETCH_WEB.fetchPages),
            fetchTimeoutMs: z.number()
                .description('单页取正文的超时（毫秒）')
                .min(LIMITS.searchFetchTimeoutMs.min).max(LIMITS.searchFetchTimeoutMs.max).step(LIMITS.searchFetchTimeoutMs.step)
                .default(DEFAULT_FETCH_WEB.fetchTimeoutMs),
            maxBytes: z.number()
                .description('单页正文的体积上限（字节），超出即截断')
                .min(LIMITS.searchFetchMaxBytes.min).max(LIMITS.searchFetchMaxBytes.max).step(LIMITS.searchFetchMaxBytes.step)
                .default(DEFAULT_FETCH_WEB.maxBytes),
            maxChars: z.number()
                .description('单页正文的字符上限，超出即截断')
                .min(LIMITS.searchFetchMaxChars.min).max(LIMITS.searchFetchMaxChars.max).step(LIMITS.searchFetchMaxChars.step)
                .default(DEFAULT_FETCH_WEB.maxChars),
            maxRedirects: z.number()
                .description('取正文最多跟随几次同源跳转（跨站跳转一律不跟）')
                .min(LIMITS.searchFetchMaxRedirects.min).max(LIMITS.searchFetchMaxRedirects.max).step(LIMITS.searchFetchMaxRedirects.step)
                .default(DEFAULT_FETCH_WEB.maxRedirects),
            pdfMaxBytes: z.number()
                .description('取正文撞上 PDF 时的体积上限（字节）。截断的 PDF 抽不出文本，所以超上限是**放弃**而不是截断：'
                    + '换来源，或先下载到本地再用 office.pdf 读')
                .min(LIMITS.searchFetchPdfMaxBytes.min).max(LIMITS.searchFetchPdfMaxBytes.max).step(LIMITS.searchFetchPdfMaxBytes.step)
                .default(DEFAULT_FETCH_WEB.pdfMaxBytes),
            pdfMaxPages: z.number()
                .description('PDF 抽文本最多抽前几页，用来圈住抽取耗时')
                .min(LIMITS.pdfMaxPages.min).max(LIMITS.pdfMaxPages.max).step(LIMITS.pdfMaxPages.step)
                .default(DEFAULT_FETCH_WEB.pdfMaxPages),
        })
            .description('取正文与 PDF 的上限（与用哪条检索通道无关）：单页超时 / 体积 / 字符、跳转次数、'
                + 'PDF 体积与页数。检索端点与 Key 在上面的「通道参数」里配')
            .default({}),
    })
        .description('检索三步走的编排参数')
        .default({}),

    // ── 记忆 ──────────────────────────────────────────────────────────────
    memory: z.object({
        enabled: z.boolean()
            .description('开启三层记忆：热记忆（用户偏好与项目约定）、台账（交付物自动登记）、归档（下沉的旧条目）。关掉后 office_memory 工具与自动台账一起停')
            .default(DEFAULT_MEMORY.enabled),
        dir: z.string()
            .description('记忆目录（相对会话工作目录）。一个项目一份记忆：换目录就是换一套记忆')
            .default(DEFAULT_MEMORY.dir),
        autoLedger: z.boolean()
            .description('office_run 每写出一份文档就自动登记进台账（路径、格式、主题、复检统计）。关掉后只能用 office_memory({action:"log"}) 手工登记')
            .default(DEFAULT_MEMORY.autoLedger),
        promptHint: z.boolean()
            .description('往系统提示里加一段静态说明（记忆在哪、怎么用，不含记忆内容）。办公模式的 persona 是完整的，那一段在那里会被丢掉；关掉主要影响其它模式')
            .default(DEFAULT_MEMORY.promptHint),
        userLimitBytes: z.number()
            .description('热记忆里「用户偏好」的容量上限（字节）。超出时把最旧、最不重要的条目下沉到归档')
            .min(LIMITS.memoryUserBytes.min).max(LIMITS.memoryUserBytes.max).step(LIMITS.memoryUserBytes.step)
            .default(DEFAULT_MEMORY.userLimitBytes),
        projectLimitBytes: z.number()
            .description('热记忆里「项目与环境」的容量上限（字节）。同上，超出下沉到归档并在 MEMORY.md 里留一条指路条目')
            .min(LIMITS.memoryProjectBytes.min).max(LIMITS.memoryProjectBytes.max).step(LIMITS.memoryProjectBytes.step)
            .default(DEFAULT_MEMORY.projectLimitBytes),
        projectionBudgetBytes: z.number()
            .description('一次投影里热记忆条目正文的字节上限（0 = 不限）。超出预算的条目不丢：折叠行报条数与读取入口。它管的是「一次反馈贴多少」，不是「库里能放多少」')
            .min(LIMITS.memoryProjectionBytes.min).max(LIMITS.memoryProjectionBytes.max).step(LIMITS.memoryProjectionBytes.step)
            .default(DEFAULT_MEMORY.projectionBudgetBytes),
        entryLimitBytes: z.number()
            .description('单条记忆的字节上限（0 = 不限）。超过就拒绝写入并给出压缩提示：热记忆是每次动笔前照办的清单，不是文档仓库')
            .min(LIMITS.memoryEntryBytes.min).max(LIMITS.memoryEntryBytes.max).step(LIMITS.memoryEntryBytes.step)
            .default(DEFAULT_MEMORY.entryLimitBytes),
        entryHintBytes: z.number()
            .description('单条记忆的建议字数线（字节，0 = 关掉）。超过只在写入回执里提醒，不拒绝 —— 用来给「写成长段落」一个可见的反馈')
            .min(LIMITS.memoryEntryBytes.min).max(LIMITS.memoryEntryBytes.max).step(LIMITS.memoryEntryBytes.step)
            .default(DEFAULT_MEMORY.entryHintBytes),
        ledgerLimit: z.number()
            .description('台账最多保留多少条，超出的最旧记录滚成月度归档摘要')
            .min(LIMITS.memoryLedgerLimit.min).max(LIMITS.memoryLedgerLimit.max).step(LIMITS.memoryLedgerLimit.step)
            .default(DEFAULT_MEMORY.ledgerLimit),
        archiveKeep: z.number()
            .description('归档最多保留多少卷摘要（一个月一卷，超 200 条的月份开下一卷）。归档也是有界的：超出的最旧卷会被删掉（它已经是压缩过的老内容）')
            .min(LIMITS.memoryArchiveKeep.min).max(LIMITS.memoryArchiveKeep.max).step(LIMITS.memoryArchiveKeep.step)
            .default(DEFAULT_MEMORY.archiveKeep),

        // ── 层拓扑（关掉不删数据） ──
        layers: z.object({
            hot: z.boolean()
                .description('热记忆层：用户偏好与项目约定，常驻并随 office_help / office_run 投影出来')
                .default(DEFAULT_LAYERS.hot),
            ledger: z.boolean()
                .description('台账层：每份交付物一条自动登记。关掉后 office_run 不再登记，手工 log 也停')
                .default(DEFAULT_LAYERS.ledger),
            archive: z.boolean()
                .description('归档层：下沉的旧条目与滚动的旧台账。关掉后检索不到归档，但文件仍在盘上')
                .default(DEFAULT_LAYERS.archive),
        })
            .description('每层独立开关。关掉只是不再读 / 不再写，已有数据不会被删除，重新打开就回来')
            .default({}),

        // ── 记忆范围（跨项目层） ──
        scope: z.union(MEMORY_SCOPES.map((id) => z.const(id)))
            .description('记忆放哪儿：workspace = 只跟着当前工作目录（一个项目一份）；global = 所有项目共用一份；both = 两者并存，全局偏好在前、项目约定在后')
            .default(DEFAULT_MEMORY.scope),
        globalDir: z.string()
            .description('全局层目录。留空 = $DSH_HOME/.office/memory；相对路径相对用户主目录（全局层不该跟着工作目录搬家）')
            .default(DEFAULT_MEMORY.globalDir),
        userScope: z.union(USER_SCOPES.map((id) => z.const(id)))
            .description('「用户偏好」这类条目放哪儿：memory = 跟着上面的范围；global = 始终落全局层，换工作目录不必重记一遍')
            .default(DEFAULT_MEMORY.userScope),
        kbDir: z.string()
            .description('知识库目录（单库）：文档、SQLite 索引与下载的模型都在这里。留空 = 全局记忆目录的兄弟 kb/；相对路径相对用户主目录。不再分全局库 / 工作区库')
            .default(DEFAULT_MEMORY.kbDir),
        galleryDir: z.string()
            .description('审美参考图库目录。留空 = 知识库目录旁边的 gallery/；原图、缩略图与索引都在这里。相对路径相对用户主目录')
            .default(''),

        // ── 知识库检索（第八十五轮）──
        embedding: z.object({
            mode: z.union(['auto', 'remote', 'local', 'off'].map((id) => z.const(id)))
                .description('向量化通道：auto = 本地推理优先、其次已配齐的远端，都没有就纯词法；remote = 只走远端；local = 本地推理，没开才用本地哈希；off = 只用词法')
                .default(DEFAULT_MEMORY.embedding.mode),
            baseURL: z.string()
                .description('OpenAI 兼容的 /embeddings 端点（如 https://api.siliconflow.cn/v1）。留空 = 远端不可用，不报错')
                .default(DEFAULT_MEMORY.embedding.baseURL),
            model: z.string()
                .description('embedding 模型 id（如 BAAI/bge-m3）。留空 = 不启用远端；auto 下直接退回词法')
                .default(DEFAULT_MEMORY.embedding.model),
            apiKey: z.string()
                .description('embedding 通道的 Key 字面量。一般留空，用下面的环境变量名')
                .role('secret')
                .default(DEFAULT_MEMORY.embedding.apiKey),
            apiKeyEnv: z.string()
                .description('取 embedding Key 的环境变量名（先查宿主凭据服务，再看环境变量）')
                .default(DEFAULT_MEMORY.embedding.apiKeyEnv),
            dimensions: z.number()
                .description('向量维度（0 = 由模型决定 / 本地默认 256）。远端模型不支持自定义维度时保持 0')
                .min(LIMITS.kbEmbeddingDimensions.min).max(LIMITS.kbEmbeddingDimensions.max).step(LIMITS.kbEmbeddingDimensions.step)
                .default(DEFAULT_MEMORY.embedding.dimensions),
            batchSize: z.number()
                .description('一次请求送多少个块去向量化')
                .min(LIMITS.kbEmbeddingBatch.min).max(LIMITS.kbEmbeddingBatch.max).step(LIMITS.kbEmbeddingBatch.step)
                .default(DEFAULT_MEMORY.embedding.batchSize),
            timeoutMs: z.number()
                .description('单次 embedding 请求的超时（毫秒）；超时或报错退回词法，不阻塞入库')
                .min(LIMITS.kbEmbeddingTimeoutMs.min).max(LIMITS.kbEmbeddingTimeoutMs.max).step(LIMITS.kbEmbeddingTimeoutMs.step)
                .default(DEFAULT_MEMORY.embedding.timeoutMs),
            local: z.object({
                enabled: z.boolean()
                    .description('按需启动本机推理服务（OpenAI 兼容的 /v1/embeddings）。默认关：它要起进程、占几百 MB 内存。打开后第一次需要向量时才启动，空闲超时自动退出')
                    .default(DEFAULT_MEMORY.embedding.local.enabled),
                command: z.string()
                    .description('解释器或可执行文件；留空 = 依次探测 python / py / python3。只在这个字段里填本机路径，不联网')
                    .default(DEFAULT_MEMORY.embedding.local.command),
                modelDir: z.string()
                    .description('模型目录；留空 = <知识库目录>/models/bge-small-zh-v1.5。目录里要有 config.json、分词器与权重')
                    .default(DEFAULT_MEMORY.embedding.local.modelDir),
                idleTimeoutMs: z.number()
                    .description('空闲多久（毫秒）自动退出。退出时杀整棵进程树，不留孤儿进程')
                    .min(LIMITS.kbLocalIdleMs.min).max(LIMITS.kbLocalIdleMs.max).step(LIMITS.kbLocalIdleMs.step)
                    .default(DEFAULT_MEMORY.embedding.local.idleTimeoutMs),
                maxRssMb: z.number()
                    .description('峰值内存硬上限（MB）。实测 bge-small-zh 在 256 token / 2 线程下约 458–558MB；超了看门狗杀进程并退回词法检索')
                    .min(LIMITS.kbLocalMaxRss.min).max(LIMITS.kbLocalMaxRss.max).step(LIMITS.kbLocalMaxRss.step)
                    .default(DEFAULT_MEMORY.embedding.local.maxRssMb),
                startupTimeoutMs: z.number()
                    .description('启动超时（毫秒）：进程起来但没打印就绪行，就判失败并退回词法检索')
                    .min(LIMITS.kbLocalStartupMs.min).max(LIMITS.kbLocalStartupMs.max).step(LIMITS.kbLocalStartupMs.step)
                    .default(DEFAULT_MEMORY.embedding.local.startupTimeoutMs),
                maxLength: z.number()
                    .description('单条分块送进模型的最长 token 数。512 更准但更吃内存；256 是本机实测的安全档')
                    .min(LIMITS.kbLocalMaxLength.min).max(LIMITS.kbLocalMaxLength.max).step(LIMITS.kbLocalMaxLength.step)
                    .default(DEFAULT_MEMORY.embedding.local.maxLength),
                threads: z.number()
                    .description('推理线程数。本机实测 2 线程比 1 线程快约 6 倍，再往上收益递减')
                    .min(LIMITS.kbLocalThreads.min).max(LIMITS.kbLocalThreads.max).step(LIMITS.kbLocalThreads.step)
                    .default(DEFAULT_MEMORY.embedding.local.threads),
            })
                .description('本机推理服务：惰性启动、空闲退出、内存有硬上限。默认关；打开后不联网也能有真正的语义向量')
                .default({}),
        })
            .description('知识库的向量化通道。默认 auto：本地推理优先、其次远端；都没有就纯词法（不会偷偷用本地哈希）')
            .default({}),
        selector: z.object({
            mode: z.union(['off', 'llm'].map((id) => z.const(id)))
                .description('候选筛选：off = 只按融合分返回；llm = 把「编号 + 每块前若干字」发给一个独立模型，只让它回编号。默认 off')
                .default(DEFAULT_MEMORY.selector.mode),
            baseURL: z.string()
                .description('筛选模型的 OpenAI 兼容 /chat/completions 端点。必填（没有沿用检索端点的回退）；留空 = 不筛选')
                .default(DEFAULT_MEMORY.selector.baseURL),
            model: z.string()
                .description('筛选模型 id。留空 = 不启用；mode=llm 且端点/模型齐全时才真的调用')
                .default(DEFAULT_MEMORY.selector.model),
            apiKey: z.string()
                .description('筛选通道的 Key 字面量。一般留空，用下面的环境变量名')
                .role('secret')
                .default(DEFAULT_MEMORY.selector.apiKey),
            apiKeyEnv: z.string()
                .description('取筛选 Key 的环境变量名（先查宿主凭据服务，再看环境变量）')
                .default(DEFAULT_MEMORY.selector.apiKeyEnv),
            timeoutMs: z.number()
                .description('筛选调用的超时（毫秒）；超时原样返回融合结果，不阻断检索')
                .min(LIMITS.kbSelectorTimeoutMs.min).max(LIMITS.kbSelectorTimeoutMs.max).step(LIMITS.kbSelectorTimeoutMs.step)
                .default(DEFAULT_MEMORY.selector.timeoutMs),
            maxCandidates: z.number()
                .description('送进筛选的候选数（只发编号 + 截断块头，不发正文）')
                .min(LIMITS.kbSelectorCandidates.min).max(LIMITS.kbSelectorCandidates.max).step(LIMITS.kbSelectorCandidates.step)
                .default(DEFAULT_MEMORY.selector.maxCandidates),
            maxSelected: z.number()
                .description('筛选后最多保留几条。主上下文只拿被选中的块，这是省上下文的关键')
                .min(LIMITS.kbSelectorMaxSelected.min).max(LIMITS.kbSelectorMaxSelected.max).step(LIMITS.kbSelectorMaxSelected.step)
                .default(DEFAULT_MEMORY.selector.maxSelected),
            headChars: z.number()
                .description('每个候选发给筛选模型的块头字数。越小越省，但也越难判断')
                .min(LIMITS.kbSelectorHeadChars.min).max(LIMITS.kbSelectorHeadChars.max).step(LIMITS.kbSelectorHeadChars.step)
                .default(DEFAULT_MEMORY.selector.headChars),
        })
            .description('候选筛选：独立模型只回编号，主上下文只拿被选中的少量块。默认 off')
            .default({}),
        ann: z.object({
            enabled: z.boolean()
                .description('向量 ANN 索引（纯 JS HNSW）。关掉就退回精确余弦；索引文件坏了、模型变了都会自动重建，不会让检索失败')
                .default(DEFAULT_MEMORY.ann.enabled),
            m: z.number()
                .description('每个节点的邻居数（建图时定死，改了会重建）。越大越准、文件越大')
                .min(LIMITS.kbAnnM.min).max(LIMITS.kbAnnM.max).step(LIMITS.kbAnnM.step)
                .default(DEFAULT_MEMORY.ann.m),
            efConstruction: z.number()
                .description('建图时候选队列的宽度。越大图越好、建得越慢')
                .min(LIMITS.kbAnnEfConstruction.min).max(LIMITS.kbAnnEfConstruction.max).step(LIMITS.kbAnnEfConstruction.step)
                .default(DEFAULT_MEMORY.ann.efConstruction),
            efSearch: z.number()
                .description('每次检索的候选宽度。越大越准越慢；默认 64')
                .min(LIMITS.kbAnnEfSearch.min).max(LIMITS.kbAnnEfSearch.max).step(LIMITS.kbAnnEfSearch.step)
                .default(DEFAULT_MEMORY.ann.efSearch),
            maxBuild: z.number()
                .description('超过这么多块就不建图（同步建图会把一次检索拖成几十秒），退回精确扫描')
                .min(LIMITS.kbAnnMaxBuild.min).max(LIMITS.kbAnnMaxBuild.max).step(LIMITS.kbAnnMaxBuild.step)
                .default(DEFAULT_MEMORY.ann.maxBuild),
        })
            .description('向量 ANN 索引：索引是 embeddings 表的派生缓存，可随时重建')
            .default({}),
        hybrid: z.object({
            lexicalWeight: z.number()
                .description('词法（BM25/FTS）一路在融合分里的权重')
                .min(LIMITS.kbWeight.min).max(LIMITS.kbWeight.max).step(LIMITS.kbWeight.step)
                .default(DEFAULT_MEMORY.hybrid.lexicalWeight),
            vectorWeight: z.number()
                .description('向量一路在融合分里的权重')
                .min(LIMITS.kbWeight.min).max(LIMITS.kbWeight.max).step(LIMITS.kbWeight.step)
                .default(DEFAULT_MEMORY.hybrid.vectorWeight),
            candidateLimit: z.number()
                .description('每路先取多少个候选再融合（候选池越大越稳，也越慢）')
                .min(LIMITS.kbCandidateLimit.min).max(LIMITS.kbCandidateLimit.max).step(LIMITS.kbCandidateLimit.step)
                .default(DEFAULT_MEMORY.hybrid.candidateLimit),
        })
            .description('词法 + 向量两路候选的融合权重。两路都为空时才退回纯词法')
            .default({}),


        // ── 知识库大模型（可选，默认全关） ──
        // 只在用户显式打开时才调用模型；默认 off，不联网、不花钱。
        llm: z.object({
            mode: z.union(['off', 'host', 'openai'].map((id) => z.const(id)))
                .description('切片 / 摘要走哪个模型通道：off = 不用模型（默认）；host = 用宿主已配置的模型；openai = 任意 OpenAI 兼容的 /chat/completions 端点')
                .default(DEFAULT_KB_LLM.mode),
            chunking: z.boolean()
                .description('让模型决定块边界。只在显式打开并选了通道后生效；块正文仍按字符偏移从原文切出，模型不改写正文')
                .default(DEFAULT_KB_LLM.chunking),
            summary: z.boolean()
                .description('让模型给每块 / 每段写摘要。摘要另存，不替代原文块；检索命中摘要会带原文块 id')
                .default(DEFAULT_KB_LLM.summary),
            provider: z.string()
                .description('host 通道的 provider id（从宿主已配置的模型里选）；openai 通道可留空')
                .default(DEFAULT_KB_LLM.provider),
            model: z.string()
                .description('模型 id（host 通道从宿主列表选；openai 通道手填，如 gpt-4o-mini）')
                .default(DEFAULT_KB_LLM.model),
            baseURL: z.string()
                .description('openai 通道的端点，到 /v1 这一层（插件补 /chat/completions）。填了它就会把原文发给这个地址')
                .default(DEFAULT_KB_LLM.baseURL),
            apiKey: z.string()
                .description('openai 通道的 Key 字面量。一般留空，用下面的环境变量名。密钥只用于请求，不会写进回执')
                .role('secret')
                .default(DEFAULT_KB_LLM.apiKey),
            apiKeyEnv: z.string()
                .description('取 Key 的环境变量名；两处都没有时明确报「没有 Key」，不会静默失败')
                .default(DEFAULT_KB_LLM.apiKeyEnv),
            timeoutMs: z.number()
                .description('单次模型调用的超时（毫秒）；超时按确定性规则回退，不阻塞入库')
                .min(1000).max(120000).step(1000)
                .default(DEFAULT_KB_LLM.timeoutMs),
            maxInputChars: z.number()
                .description('一次送给模型的原文上限（字符）。按章分批，绝不把整本一次塞进去')
                .min(1000).max(60000).step(500)
                .default(DEFAULT_KB_LLM.maxInputChars),
            maxOutputTokens: z.number()
                .description('一次模型调用的输出上限（token）')
                .min(200).max(8000).step(100)
                .default(DEFAULT_KB_LLM.maxOutputTokens),
            maxCallsPerRun: z.number()
                .description('单次入库最多调用几次模型（可续跑）；到了就暂停并保留进度')
                .min(1).max(100).step(1)
                .default(DEFAULT_KB_LLM.maxCallsPerRun),
            maxCalls: z.number()
                .description('这一份文档的模型调用总预算。超出预算直接拒绝，不会先烧钱再报错')
                .min(1).max(5000).step(1)
                .default(DEFAULT_KB_LLM.maxCalls),
            maxChunkChars: z.number()
                .description('模型切片时单块字符上限；超过再按小节切')
                .min(500).max(10000).step(100)
                .default(DEFAULT_KB_LLM.maxChunkChars),
            minChunkChars: z.number()
                .description('模型切片时单块建议下限（合并太碎的段）')
                .min(100).max(3000).step(50)
                .default(DEFAULT_KB_LLM.minChunkChars),
            maxUnitsPerBatch: z.number()
                .description('一次调用最多包含多少个小节（防止编号太多模型数错）')
                .min(10).max(200).step(10)
                .default(DEFAULT_KB_LLM.maxUnitsPerBatch),
        })
            .description('知识库的可选大模型增强：切片边界与摘要。默认关闭；打开后按章分批调用，有预算上限、失败自动回退到确定性切块，原文一个字不改')
            .default({}),
        // ── 主动记录 ──
        autoCapture: z.boolean()
            .description('主动记录：在工具反馈里提示「用户这轮说出的偏好与纠正当场记一条」，不必等他说「记住」。这是一句指引，不是自主记录器')
            .default(DEFAULT_MEMORY.autoCapture),

        // ── 召回质量（strict-v1） ──
        recallQuality: z.object({
            policy: z.union(['strict-v1', 'off'].map((id) => z.const(id)))
                .description('strict-v1 = 按相关度分档，低相关的丢掉；off = 只按关键词排序、不丢结果')
                .default(DEFAULT_RECALL_QUALITY.policy),
            lowScoreThreshold: z.number()
                .description('命中词占比低于它算「未知」档。注意：本插件的分数是「命中了查询里几成的词」，不是向量余弦相似度')
                .min(LIMITS.recallLowScore.min).max(LIMITS.recallLowScore.max)
                .default(DEFAULT_RECALL_QUALITY.lowScoreThreshold),
            highScoreThreshold: z.number()
                .description('达到它算「高分」档，直接采纳。必须大于上面那个下限，写反了整组会退回默认')
                .min(LIMITS.recallHighScore.min).max(LIMITS.recallHighScore.max)
                .default(DEFAULT_RECALL_QUALITY.highScoreThreshold),
            candidateMultiplier: z.number()
                .description('先取「条数上限 × 这个倍数」个候选再分档，防止低分结果把高分结果挤出候选池')
                .min(LIMITS.recallCandidateMultiplier.min).max(LIMITS.recallCandidateMultiplier.max).step(LIMITS.recallCandidateMultiplier.step)
                .default(DEFAULT_RECALL_QUALITY.candidateMultiplier),
            maxMediumResults: z.number()
                .description('「中档」最多采纳几条')
                .min(LIMITS.recallMaxMedium.min).max(LIMITS.recallMaxMedium.max).step(LIMITS.recallMaxMedium.step)
                .default(DEFAULT_RECALL_QUALITY.maxMediumResults),
            maxUnknownResults: z.number()
                .description('「未知档」最多采纳几条（0 = 一条都不要）')
                .min(LIMITS.recallMaxUnknown.min).max(LIMITS.recallMaxUnknown.max).step(LIMITS.recallMaxUnknown.step)
                .default(DEFAULT_RECALL_QUALITY.maxUnknownResults),
        })
            .description('召回质量：一次检索返回什么。管的是「结果好不好」，每回合配额管的是「能查几次」')
            .default({}),

        // ── 每回合配额 ──
        quota: z.object({
            recallPerTurn: z.number()
                .description('一个回合里第一次带关键词的记忆检索（0 = 不限制）')
                .min(LIMITS.quotaPerTurn.min).max(LIMITS.quotaPerTurn.max).step(LIMITS.quotaPerTurn.step)
                .default(DEFAULT_QUOTA.recallPerTurn),
            recallRefinePerTurn: z.number()
                .description('同一回合里后续的换词细化次数。用完就拒绝，并说明原因')
                .min(LIMITS.quotaPerTurn.min).max(LIMITS.quotaPerTurn.max).step(LIMITS.quotaPerTurn.step)
                .default(DEFAULT_QUOTA.recallRefinePerTurn),
            kbSearchPerTurn: z.number()
                .description('一个回合里的知识库（kb）检索次数。kb 检索是词法检索（bigram + BM25），换词重查容易一直查下去，所以与记忆检索分开算')
                .min(LIMITS.quotaPerTurn.min).max(LIMITS.quotaPerTurn.max).step(LIMITS.quotaPerTurn.step)
                .default(DEFAULT_QUOTA.kbSearchPerTurn),
            relatedPerTurn: z.number()
                .description('一个回合里沿知识库块关系（kb-related）跳转的次数')
                .min(LIMITS.quotaPerTurn.min).max(LIMITS.quotaPerTurn.max).step(LIMITS.quotaPerTurn.step)
                .default(DEFAULT_QUOTA.relatedPerTurn),
        })
            .description('每回合配额：防止「没查到就换个词再查」把上下文填满。回合号从会话日志的 turn/start 读出来')
            .default({}),
    })
        .description('三层记忆：热记忆常驻、台账自动登记、归档只读。记忆不是缓存，不会被 TTL 清掉')
        .default({}),

    // ── 文件识别（OCR）──────────────────────────────────────────────────
    // 官方契约（api-docs.deepseek.com/zh-cn/api/ocr）：POST https://api.deepseek.com/v1/tools/ocr，
    // multipart 只收 file + purpose，和本组字段一一对应，不多传任何参数。
    ocr: z.object({
        enabled: z.boolean()
            .description('把 PDF / Word / PPT / 图片里的文字读出来；关掉后「识别文件」提交会直接报错，不影响其他功能')
            .default(DEFAULT_OCR.enabled),
        baseURL: z.string()
            .description('默认官方 https://api.deepseek.com；带 /v1 也认，插件补 /tools/ocr')
            .default(DEFAULT_OCR.baseURL),
        apiKey: z.string()
            .description('留空读下面的环境变量；两处都没有时明确报「没有 Key」，不会静默失败')
            .role('secret')
            .default(DEFAULT_OCR.apiKey),
        apiKeyEnv: z.string()
            .description('Key 的环境变量名（默认 DEEPSEEK_API_KEY）')
            .default(DEFAULT_OCR.apiKeyEnv),
        reuseOfficialProvider: z.boolean()
            .description('显式允许在官方端点上复用宿主已保存的 deepseek-official 凭据；默认关。'
                + '只在本地直连官方端点、且专用 Key 与环境变量都没有时生效')
            .default(false),
        purpose: z.union(OCR_PURPOSES.map((id) => z.const(id)))
            .description('text_recognition 只要文字；rich_layout_recognition 连版面标记一起返回，适合还原排版')
            .default(DEFAULT_OCR.purpose),
        timeoutMs: z.number()
            .description('单次超时（毫秒）；官方总时长上限 180 秒，调大也不超过')
            .min(LIMITS.ocrTimeoutMs.min).max(LIMITS.ocrTimeoutMs.max).step(LIMITS.ocrTimeoutMs.step)
            .default(DEFAULT_OCR.timeoutMs),
    })
        .description('文件识别（OCR）：走 DeepSeek 官方 /v1/tools/ocr，按页计费（每 1000 页 1.5 元，官方文档口径）')
        .default({}),
    // ── 后台任务 Agent 的模型路由 ──────────────────────────────────────────
    subagentModel: z.object({
        mode: z.union(['inherit', 'fixed'].map((id) => z.const(id)))
            .description('检索子代理用哪个模型：inherit = 跟随主会话；fixed = 用下面指定的模型')
            .default(DEFAULT_SUBAGENT_MODEL.mode),
        provider: z.string()
            .description('fixed 时的 provider 路由名。留空 = 只覆盖 model，仍留在主会话那条路由上')
            .default(DEFAULT_SUBAGENT_MODEL.provider),
        model: z.string()
            .description('fixed 时的模型 id。留空按 inherit 处理。只影响检索子代理，不影响主对话')
            .default(DEFAULT_SUBAGENT_MODEL.model),
    })
        .description('后台任务 Agent 的模型路由。办公插件里的后台任务只有一处：检索三步走的派工子代理')
        .default({}),

    // ── 审美参考图库（默认不联网）──────────────────────────────────────
    // 两件会花钱的事分开配：视觉描述（看图写主体/构图/配色…）与图文对齐
    // embedding（真正把图片和文字放进同一个向量空间）。默认都是关的，
    // 打开并显式 allowUpload 才会把图片发出去。
    gallery: z.object({
        vision: z.object({
            channel: z.union(['off', 'host', 'openai'].map((id) => z.const(id)))
                .description('视觉描述走哪条通道：off = 不生成描述（默认，只保存图与手工标签）；host = 用宿主已配置的多模态模型；openai = 任意 OpenAI 兼容的 chat 端点')
                .default('off'),
            provider: z.string()
                .description('host 通道的 provider id；openai 通道可留空')
                .default(''),
            model: z.string()
                .description('视觉模型 id（host 通道从宿主列表选；openai 通道手填，如 gpt-4o-mini）')
                .default(''),
            baseURL: z.string()
                .description('openai 通道端点，到 /v1 这一层（插件补 /chat/completions）。填了它就会把缩略图发给这个地址')
                .default(''),
            apiKey: z.string()
                .description('openai 通道 Key 字面量；一般留空用环境变量名。密钥只用于请求，不会写进回执')
                .role('secret')
                .default(''),
            apiKeyEnv: z.string()
                .description('取 Key 的环境变量名；两处都没有时明确报「未配置」，不会假成功')
                .default(''),
            maxTokens: z.number()
                .description('一次描述的输出上限（token）')
                .min(200).max(2000).step(100)
                .default(900),
            timeoutMs: z.number()
                .description('单次视觉调用超时（毫秒）')
                .min(5000).max(180000).step(1000)
                .default(60000),
            language: z.union(['zh', 'en'].map((id) => z.const(id)))
                .description('描述语言')
                .default('zh'),
        })
            .description('可选的视觉描述：生成主体/构图/配色/光线/材质/氛围与审美标签。默认 off；描述只作审美参考，不推断作者与版权')
            .default({}),
        imageEmbedding: z.object({
            provider: z.union(['none', 'voyage', 'cohere', 'vllm'].map((id) => z.const(id)))
                .description('图文对齐 embedding 通道：none = 不做图片向量（默认）；voyage / cohere 用官方多模态接口；vllm = 任意 OpenAI 兼容的多模态 /embeddings')
                .default('none'),
            baseURL: z.string()
                .description('vllm 通道必填；voyage/cohere 留空用官方端点。注意：普通文本 embedding 接口不能当多模态用')
                .default(''),
            model: z.string()
                .description('多模态 embedding 模型 id。留空时 voyage/cohere 用官方默认；vllm 必填')
                .default(''),
            apiKey: z.string()
                .description('Key 字面量；一般留空用环境变量名')
                .role('secret')
                .default(''),
            apiKeyEnv: z.string()
                .description('取 Key 的环境变量名；与模型、端点一起校验，缺一项就明确报未配置')
                .default(''),
            dimensions: z.number()
                .description('期望维度（0 = 用模型默认）。与文本向量不一致时检索会明确报维度不匹配，不会硬拼')
                .min(0).max(4096).step(1)
                .default(0),
            batchSize: z.number()
                .description('一次请求最多几张图（1-16）')
                .min(1).max(16).step(1)
                .default(4),
            timeoutMs: z.number()
                .description('单次 embedding 请求超时（毫秒）')
                .min(5000).max(180000).step(1000)
                .default(30000),
        })
            .description('真正的图文对齐向量：文字搜图、以图搜图与标签/描述/图片向量融合都靠它。默认 none，未配置时明确报未配置，绝不拿文本向量冒充')
            .default({}),
    })
        .description('审美参考图库：原图与缩略图、手工标签/收藏/分类、可选视觉描述与多模态向量。默认完全本地、不上传、不计费')
        .default({}),

    // ── 文档与缓存 ────────────────────────────────────────────────────────
    documents: z.object({
        defaultTheme: z.union(THEME_IDS.map((id) => z.const(id)))
            .description('不传主题时用哪一套配色。plain = 素色网格：Excel 只留细边框、Word 白纸黑字，不加任何背景填充')
            .default(DEFAULT_THEME_ID),
        scriptTimeoutMs: z.number()
            .description('单次 office_run 脚本的超时（毫秒）。批量生成很多文件时可以调大')
            .min(LIMITS.scriptTimeoutMs.min).max(LIMITS.scriptTimeoutMs.max).step(LIMITS.scriptTimeoutMs.step)
            .default(DEFAULT_SCRIPT_TIMEOUT_MS),
        maxScriptChars: z.number()
            .description('单次脚本的字符上限。脚本很大时调大，但过大会拖慢错误定位')
            .min(LIMITS.maxScriptChars.min).max(LIMITS.maxScriptChars.max).step(LIMITS.maxScriptChars.step)
            .default(DEFAULT_MAX_SCRIPT_CHARS),
        cacheDir: z.string()
            .description('中间产物目录（相对会话工作目录）。跨调用保留，按下面的存活时间自动清理')
            .default(CACHE_DIR),
        keepCache: z.boolean()
            .description('调用结束后保留缓存里的中间文件。保留才能跨调用复用（例如渲染好的 PDF 页面图）；关掉则每次调用结束就清空')
            .default(true),
        cacheTtlMinutes: z.number()
            .description('中间产物多少分钟没被碰过就清掉（每次 office_run 开始时清理）。0 = 不按时间清理')
            .min(LIMITS.cacheTtlMinutes.min).max(LIMITS.cacheTtlMinutes.max).step(LIMITS.cacheTtlMinutes.step)
            .default(DEFAULT_CACHE_TTL_MINUTES),
        pdfDpi: z.number()
            .description('PDF 渲染成图片时的默认分辨率。手写笔记建议 120-200；越大越慢、图片越大')
            .min(LIMITS.pdfDpi.min).max(LIMITS.pdfDpi.max).step(LIMITS.pdfDpi.step)
            .default(DEFAULT_PDF_DPI),
        pdfMaxPages: z.number()
            .description('单次 office.pdf.pages() 最多渲染多少页。渲染是重活，防止一次卡住整个脚本')
            .min(LIMITS.pdfMaxPages.min).max(LIMITS.pdfMaxPages.max).step(LIMITS.pdfMaxPages.step)
            .default(DEFAULT_PDF_MAX_PAGES),
        pdfEngine: z.union(['auto', 'fitz', 'pdftoppm', 'pdftocairo', 'mutool', 'gs'])
            .description('PDF 渲染引擎。auto = 自动探测（PyMuPDF / poppler / MuPDF / Ghostscript），一个都没有时 office.pdf.pages() 会明确报错')
            .default(DEFAULT_PDF_ENGINE),
        texEngine: z.union(['auto', 'latexmk', 'xelatex', 'lualatex'])
            .description('LaTeX 编译入口。auto = 有 latexmk 就用它（自动处理 bibtex 与重复编译），否则用 xelatex。thuthesis 需要 fontspec，pdflatex 不可用')
            .default(DEFAULT_TEX_ENGINE),
        texTimeoutMs: z.number()
            .description('单次 LaTeX 编译的超时（毫秒）。学位论文要跑三四遍 xelatex 加 bibtex，大论文要调大')
            .min(LIMITS.texTimeoutMs.min).max(LIMITS.texTimeoutMs.max).step(LIMITS.texTimeoutMs.step)
            .default(DEFAULT_TEX_TIMEOUT_MS),
        texTemplateDir: z.string()
            .description('thuthesis 模板目录（含 thuthesis.cls 与两个校徽）。留空时自动探测：工作目录里的 thuthesis* 目录，其次 TeX Live 自带的版本')
            .default(''),
    })
        .description('文档生成与缓存')
        .default({}),

    // ── preset 工具守卫（新-15①） ─────────────────────────────────────────
    guard: z.object({
        hideDevToolsInOffice: z.boolean()
            .description('打开后，办公会话里真正掩掉 preset 声明要关的那批宿主工具'
                + '（命令行 / 后台任务 / 工作流 / 子代理工具 / 待办 / 技能 / 目标 / 插件管理 / 任务看板）。'
                + '默认关闭：这些入口消失属削减能力，该由用户明确要求。'
                + '检索三步走的派工不受影响（它用的是宿主 subagents 服务，不是 subagent 工具）。'
                + '即时生效：已经在跑的办公会话也会立刻重算（那些入口当场消失，不必重开）')
            .default(DEFAULT_GUARD.hideDevToolsInOffice),
    })
        .description('preset 工具守卫（第五十四轮 新-15①）：把 preset 里那些「声明要关、运行期不生效」的禁用行'
            + '真正落到工具面上。只动办公 preset，默认不隐藏')
        .default({}),

    // ── Python 计算与绘图 ─────────────────────────────────────────────────
    python: z.object({
        enabled: z.boolean()
            .description('开启 Python 通道（office.python.run / file / check）：科学计算与绘图。这是办公模式里唯一的编程出口，命令行仍然是关掉的')
            .default(DEFAULT_PYTHON.enabled),
        bin: z.string()
            .description('Python 解释器：留空自动探测 PATH 里的 python / python3 / py；也可以填绝对路径（如 C:\\Python313\\python.exe）。应用商店的占位程序不算可用')
            .default(DEFAULT_PYTHON.bin),
        timeoutMs: z.number()
            .description('单次 Python 运行的超时（毫秒）。跑长计算或大批量绘图时调大')
            .min(LIMITS.pythonTimeoutMs.min).max(LIMITS.pythonTimeoutMs.max).step(LIMITS.pythonTimeoutMs.step)
            .default(DEFAULT_PYTHON.timeoutMs),
        outDir: z.string()
            .description('Python 产物（图 / 表 / 数据）的目录，相对缓存目录。脚本里拿 OUT_DIR 就能拿到它的绝对路径')
            .default(DEFAULT_PYTHON.outDir),
    })
        .description('Python 计算与绘图：只跑 Python，不开命令行；产物只落缓存目录，按缓存规则自动清理')
        .default({}),

    // ── 音频与视频的内容提取 ──────────────────────────────────────────────
    av: z.object({
        enabled: z.boolean()
            .description('开启音频与视频的内容提取（office.av）：声音转文字、视频按时间点抽帧。关掉只是让 office.av 报错，不删任何东西')
            .default(DEFAULT_AV.enabled),
        ffmpegPath: z.string()
            .description('ffmpeg 可执行文件：留空自动探测（PATH，其次 C:\\app\\ffmpeg\\bin 这类常见目录）。ffprobe 默认取它同目录的那一个')
            .default(DEFAULT_AV.ffmpegPath),
        ffprobePath: z.string()
            .description('ffprobe 可执行文件：留空 = 先按 ffmpeg 同目录找，再走 PATH')
            .default(DEFAULT_AV.ffprobePath),
        modelDir: z.string()
            .description('SenseVoice 模型目录（含 model.int8.onnx 与 tokens.txt）。留空 = 语音输入下载的那份：$DSH_HOME/speech-to-text/sensevoice/models/sensevoice-onnx')
            .default(DEFAULT_AV.modelDir),
        language: z.union(AV_LANGUAGES.map((id) => z.const(id)))
            .description('转写语言：auto = 让模型自己判（支持中 / 粤 / 英 / 日 / 韩）。固定成某一种时，混语种录音可能被强行翻成那一种')
            .default(DEFAULT_AV.language),
        chunkSeconds: z.number()
            .description('单块秒数。模型一次吃的音频越长内存越高，120 秒是官方那条链 131 秒上限留了余量的取值')
            .min(LIMITS.avChunkSeconds.min).max(LIMITS.avChunkSeconds.max).step(LIMITS.avChunkSeconds.step)
            .default(DEFAULT_AV.chunkSeconds),
        overlapSeconds: z.number()
            .description('相邻两块的重叠秒数（默认 0.5）。块不再是完全切开的两段：跨在切点上的短音能在同一块里解完；'
                + '跨得更深的长句由边界回退兜住（不丢字也不重复），所以不必靠调大它来换效果。上限是半块')
            .min(LIMITS.avOverlapSeconds.min).max(LIMITS.avOverlapSeconds.max).step(LIMITS.avOverlapSeconds.step)
            .default(DEFAULT_AV.overlapSeconds),
        maxSeconds: z.number()
            .description('单次处理的音频时长上限（秒）。超过就报错并要求先切段 —— 这是内存保护，切块只解决单块大小，不改变这条上限')
            .min(LIMITS.avMaxSeconds.min).max(LIMITS.avMaxSeconds.max).step(LIMITS.avMaxSeconds.step)
            .default(DEFAULT_AV.maxSeconds),
        timeoutMs: z.number()
            .description('单次解码 / 抽帧 / 转写的超时（毫秒）。长会议录音调大')
            .min(LIMITS.avTimeoutMs.min).max(LIMITS.avTimeoutMs.max).step(LIMITS.avTimeoutMs.step)
            .default(DEFAULT_AV.timeoutMs),
        precision: z.union(['int8', 'fp32'].map((id) => z.const(id)))
            .description('模型精度：int8（默认，约 228 MB）够快够准；fp32（约 894 MB）在本机没有实测收益')
            .default(DEFAULT_AV.precision),
        threads: z.number()
            .description('推理线程数。默认 2：再往上收益递减，而且会和宿主抢 CPU')
            .min(LIMITS.avThreads.min).max(LIMITS.avThreads.max).step(LIMITS.avThreads.step)
            .default(DEFAULT_AV.threads),
        frames: z.number()
            .description('视频默认抽几张画面（不给 at / every / count 时）。抽出来的图交给 read_image 看')
            .min(LIMITS.avFrames.min).max(LIMITS.avFrames.max).step(LIMITS.avFrames.step)
            .default(DEFAULT_AV.frames),
        maxFrames: z.number()
            .description('单次抽帧数上限，防止一次抽几百张把缓存与上下文撑爆')
            .min(LIMITS.avMaxFrames.min).max(LIMITS.avMaxFrames.max).step(LIMITS.avMaxFrames.step)
            .default(DEFAULT_AV.maxFrames),
        frameFormat: z.union(['jpg', 'png'].map((id) => z.const(id)))
            .description('抽帧格式：jpg 体积小（默认，适合照片与实拍）；png 无损，图上文字更清楚')
            .default(DEFAULT_AV.frameFormat),
        frameWidth: z.number()
            .description('抽帧缩放后的宽度（像素）。0 = 不缩放（原分辨率）；4K 视频建议填 1280 之类，图小、读起来也快')
            .min(LIMITS.avFrameWidth.min).max(LIMITS.avFrameWidth.max).step(LIMITS.avFrameWidth.step)
            .default(DEFAULT_AV.frameWidth),
    })
        .description('音频与视频的内容提取：ffmpeg 解码 + 本机 SenseVoice 离线转写 + 视频抽帧。'
            + 'VAD 阈值、单句上限、词级时间戳、缓存子目录等细项只在配置文件里可调（见 office_help topic:av 的 detail 层）')
        .default({}),
    });
    markLeavesVolatile(schema);
    return schema;
}

/**
 * schema 的默认值：空配置下解析出来的那一份。
 * @param {object} z schemastery 实例
 */
export function settingsDefaults(z) {
    return createOfficeSettings(z)({});
}