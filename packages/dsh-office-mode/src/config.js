/**
 * 插件配置解析。空配置永远可用，这一条是硬约束：装错了也不该弄崩宿主。
 *
 * @module dsh-office-mode/config
 */
import { DEFAULT_THEME_ID, THEMES } from './engine/theme.js';
import { SITE_PRIORITY_LIMITS } from './site-catalog.js';
import { DEFAULT_PREPROCESS, PREPROCESS_MODES } from './web-preprocess.js';
import { PROVIDER_IDS } from './web-providers.js';
import { DEFAULT_PYTHON_OUT_DIR, DEFAULT_PYTHON_TIMEOUT_MS } from './python.js';
import {
    AV_CHUNK_OVERLAP_SECONDS,
    AV_CHUNK_SECONDS,
    AV_DEFAULT_FRAMES,
    AV_LANGUAGES,
    AV_MAX_FRAMES,
    AV_MAX_SECONDS,
    AV_TIMEOUT_MS,
} from './av.js';
import {
    DEFAULT_ARCHIVE_KEEP,
    DEFAULT_LEDGER_LIMIT,
    DEFAULT_MEMORY_DIR,
    DEFAULT_PROJECT_LIMIT_BYTES,
    DEFAULT_PROJECTION_BUDGET_BYTES,
    DEFAULT_ENTRY_HINT_BYTES,
    DEFAULT_ENTRY_LIMIT_BYTES,
    DEFAULT_USER_LIMIT_BYTES,
    DEFAULT_LAYERS,
    DEFAULT_RECALL_QUALITY,
    DEFAULT_QUOTA,
    MEMORY_SCOPES,
    USER_SCOPES,
} from './memory.js';

export const DEFAULT_SCRIPT_TIMEOUT_MS = 60_000;
export const DEFAULT_MAX_SCRIPT_CHARS = 400_000;

/**
 * 缓存存活时间（分钟）：超过这个时间没被碰过的中间产物，在下一次 office_run
 * 开始时被清掉。默认 12 小时 —— 够一个工作会话跨很多次调用复用（最典型的是
 * 渲染好的 PDF 页面图），又不至于把工作目录当垃圾场。0 = 不按时间清理。
 */
export const DEFAULT_CACHE_TTL_MINUTES = 720;
/** 缓存总容量上限（字节），超出后从最旧的开始删。默认 512 MB。 */
export const DEFAULT_CACHE_MAX_BYTES = 512 * 1024 * 1024;
/** PDF 渲染的默认 DPI（够看清手写；越高越慢、图越大）。 */
export const DEFAULT_PDF_DPI = 120;
/** 单次 office.pdf.pages() 最多渲染多少页（渲染是重活，防止一次卡死会话）。 */
export const DEFAULT_PDF_MAX_PAGES = 20;
/** PDF 渲染引擎：auto = 自动探测（fitz → pdftoppm → pdftocairo → mutool → gs）。 */
export const DEFAULT_PDF_ENGINE = 'auto';
/**
 * LaTeX 编译入口：auto = 有 latexmk 就用它（它自己会算 bibtex 与重复编译次数），
 * 否则退回 xelatex / lualatex。thuthesis 需要 fontspec，pdflatex 不在候选里。
 */
export const DEFAULT_TEX_ENGINE = 'auto';
/**
 * 单次 LaTeX 编译的超时（毫秒）。学位论文要跑三四遍 xelatex 加 bibtex，
 * 目录越长越慢，所以给的默认值比脚本超时大得多。
 */
export const DEFAULT_TEX_TIMEOUT_MS = 300_000;
export const CACHE_DIR_ENV = 'DSH_OFFICE_CACHE_DIR';
export const THEME_ENV = 'DSH_OFFICE_THEME';
export const PDF_ENGINE_ENV = 'DSH_OFFICE_PDF_ENGINE';
export const TEX_ENGINE_ENV = 'DSH_OFFICE_TEX_ENGINE';
export const TEX_TEMPLATE_ENV = 'DSH_OFFICE_TEX_TEMPLATE';
export const MEMORY_DIR_ENV = 'DSH_OFFICE_MEMORY_DIR';
export const MEMORY_GLOBAL_DIR_ENV = 'DSH_OFFICE_MEMORY_GLOBAL_DIR';
export const PYTHON_BIN_ENV = 'DSH_OFFICE_PYTHON';
export const PYTHON_OUT_DIR_ENV = 'DSH_OFFICE_PYTHON_OUT';
/** 音频 / 视频通道：ffmpeg 与 SenseVoice 模型也可以用环境变量给（默认按设备上已有的那份找）。 */
export const AV_FFMPEG_ENV = 'DSH_OFFICE_FFMPEG';
export const AV_MODEL_DIR_ENV = 'DSH_OFFICE_SENSEVOICE_DIR';
/** 内置检索的端点/模型/Key 名也可以用环境变量给（与宿主 web-search-deepseek 的 DEEPSEEK_SEARCH_BASE_URL 对齐）。 */
export const BUILTIN_SEARCH_BASE_URL_ENV = 'DSH_OFFICE_SEARCH_BASE_URL';
export const BUILTIN_SEARCH_MODEL_ENV = 'DSH_OFFICE_SEARCH_MODEL';
export const BUILTIN_SEARCH_KEY_ENV = 'DSH_OFFICE_SEARCH_API_KEY_ENV';

function text(value) {
    return typeof value === 'string' ? value.trim() : '';
}

function positiveInt(value, fallback, min, max) {
    const n = typeof value === 'number' ? Math.trunc(value) : Number.parseInt(String(value), 10);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

/** 默认开启的工具（= 设置页里的默认勾选状态）。 */
export const DEFAULT_TOOLS = {
    office_help: true,
    office_run: true,
    office_memory: true,
    office_web_search: true,
    office_web_fetch: true,
    office_search_run: true,
    office_search_brief: true,
    office_search_dispatch: true,
    office_parse_findings: true,
};

/**
 * 三层记忆的默认参数。
 *
 * 与缓存不同：记忆**不是**中间产物，不该被 TTL 清掉，也不该在设置页里当
 * 「用完即弃」的开关看待。所以这里给的是容量与范围，清理只能显式做。
 *
 * `promptHint` 只注册一段**静态**说明（记忆在哪、怎么用），不含任何记忆内容：
 * section 的 text 只接受同步函数，而读盘是异步的。它服务的正是「mnemon 关掉
 * 之后谁来提醒模型有记忆」这件事；办公模式的 persona 是 complete 的，这段在
 * 办公模式下会被组装丢掉（那里的投影由 office_help / office_run 的返回值负责）。
 */
export const DEFAULT_MEMORY = {
    enabled: true,
    dir: DEFAULT_MEMORY_DIR,
    autoLedger: true,
    promptHint: true,
    userLimitBytes: DEFAULT_USER_LIMIT_BYTES,
    projectLimitBytes: DEFAULT_PROJECT_LIMIT_BYTES,
    /**
     * 一次投影里条目正文的字节上限（第四十八轮 P0-2）。0 = 不限。
     *
     * 热记忆的容量上限管的是「库里能放多少」，这个管的是「一次反馈里贴多少」——
     * 两件事在旧实现里混着：本机实测全文 15,444 B，一轮 20 次调用累计 76,656 B，
     * 是九个工具全部声明（15,451 B）的 4.96 倍。超出预算的条目不丢：折叠行报条数并给
     * 读取入口（第三十六轮的三条纪律照旧）。
     */
    projectionBudgetBytes: DEFAULT_PROJECTION_BUDGET_BYTES,
    /**
     * 单条记忆的硬上限与建议线（第四十八轮 新-11），都是 0 = 关掉。
     *
     * 硬上限拒绝写入（附压缩模板），建议线只在回执里提醒。两者与容量上限是**三件不同的事**：
     * 容量上限管「这一层能装多少」，硬上限管「一条能多长」，建议线管「多长算写歪了」。
     */
    entryLimitBytes: DEFAULT_ENTRY_LIMIT_BYTES,
    entryHintBytes: DEFAULT_ENTRY_HINT_BYTES,
    ledgerLimit: DEFAULT_LEDGER_LIMIT,
    archiveKeep: DEFAULT_ARCHIVE_KEEP,
    /** 记忆层拓扑（C6）：每层独立开关，关掉不删数据。 */
    layers: { ...DEFAULT_LAYERS },
    /** 记忆范围（C3）：workspace / global / both。 */
    scope: 'workspace',
    /** 全局层目录；留空 = $DSH_HOME/.office/memory（见 memory.js 的 globalMemoryDir）。 */
    globalDir: '',
    /** 用户档案范围（C4）：memory = 跟着当前范围；global = USER 始终落全局层。 */
    userScope: 'memory',
    /** 图关系开关：关掉后 link / related / 实体视图都不可用。 */
    links: true,
    /** 主动记录指引（C7）：往工具反馈里加一段「值得长期保留的事实要主动记」的说明。 */
    autoCapture: true,
    /** 召回质量策略（E3）：strict-v1 的阈值与名额。 */
    recallQuality: { ...DEFAULT_RECALL_QUALITY },
    /** 每回合配额（E2）：一个回合里每类检索最多几次。0 = 不限制。 */
    quota: { ...DEFAULT_QUOTA },
};

/**
 * 后台任务 Agent 的模型路由（C9）。
 *
 * 办公插件里的「后台任务」只有一处：检索三步走的派工子代理。inherit 跟随主
 * 会话的模型；fixed 时给子代理指定一个模型 id（留空按 inherit 处理）。
 */
export const DEFAULT_SUBAGENT_MODEL = { mode: 'inherit', provider: '', model: '' };

/**
 * 内置检索（office 自己的联网通道）的默认参数。
 *
 * 为什么需要它：它让「办公会话没有宿主联网工具（preset 不声明 tool-web）」
 * 时也能查资料 —— 这正是第三十轮起的常态。默认端点是 DeepSeek 的 Anthropic
 * 兼容 Messages 接口 + 原生 web_search 工具，与宿主 web-search-deepseek 的
 * 默认一致；Key 默认取 DEEPSEEK_API_KEY。它只在显式点名 anthropic 通道或把它
 * 加回 providerOrder 时才会用到（默认顺序是免 Key 的抓取通道）。
 */
export const DEFAULT_BUILTIN_WEB = {
    /** 字面量 Key（设置页里填的）；留空则按 apiKeyEnv 去凭据服务/环境/凭据文件里取。 */
    apiKey: '',
    apiKeyEnv: 'DEEPSEEK_API_KEY',
    baseURL: 'https://api.deepseek.com/anthropic/v1',
    model: 'deepseek-v4-flash',
    apiVersion: '2023-06-01',
    /** 一次检索最多让模型调用几次原生 web_search。 */
    maxUses: 5,
    maxTokens: 2048,
    /** 单次检索 / 单页取正文的超时（毫秒）。 */
    searchTimeoutMs: 60_000,
    fetchTimeoutMs: 20_000,
    /** 一次检索最多要几条来源。 */
    maxResults: 8,
    /** 每个渠道（直查时每条查询）最多打开几页取正文摘录。0 = 只取搜索返回的标题与引用片段。 */
    fetchPages: 2,
    /** 单页正文的体积与字符上限。 */
    maxBytes: 2 * 1024 * 1024,
    maxChars: 20_000,
    maxRedirects: 3,
    userAgent: 'dsh-office-mode/0.1 (+built-in web access)',
    /**
     * 取正文撞上 PDF 时的上限（第二十九轮新增）。
     *
     * 检索回来的来源里 PDF 很常见（学生题库、期刊与机构站点），以前这一律报
     * 「不支持的内容类型」；现在下载后交给 office.pdf 抽文本。截断的 PDF 抽不出
     * 东西，所以体积上限**不是截断线而是放弃线**：超了就明确说「太大，换来源」。
     * 页数上限用来圈住抽取耗时（pdftotext / fitz 都是按页跑的）。
     */
    pdfMaxBytes: 24 * 1024 * 1024,
    pdfMaxPages: 30,
};

/** 检索编排的默认参数。 */
export const DEFAULT_SEARCH = {
    maxParallel: 4,
    resultLimit: 40,
    maxChannels: 12,
    outputDir: '.office/search',
    requireCrossSource: true,
    fallbackOnPlatformError: true,
    /**
     * 检索的执行引擎。
     *   auto     —— 先按老路子派子代理；组合里没有联网检索工具（精简部署就是这样）
     *               或没有 subagents 服务时，自动改用插件内置的检索通道。
     *   subagent —— 只用子代理（老行为）。
     *   builtin  —— 只用内置通道，不起子代理。
     */
    engine: 'auto',
    /**
     * 用哪条联网通道（第二十轮新增）。'auto' = 按 providerOrder 依次尝试；
     * 也可以固定成某一条 id（seam / anthropic / openai / tavily / brave /
     * bocha / exa / serper / searxng / duckduckgo）。
     */
    provider: 'auto',
    /**
     * auto 时的尝试顺序（第三十轮起换主次）。
     *
     * 办公模式的资料搜索走**自己的网页抓取通道**：DuckDuckGo HTML 抓取（免 Key）
     * 最前，自建 SearXNG 随后，宿主的 web 服务（seam）只做兜底。三方检索 API
     * （anthropic / openai / tavily / brave / bocha / exa / serper）不再进默认
     * 顺序 —— 要用就在设置页把 providerOrder 加回去，或用 provider 参数按次点名
     * （office_search_run({ provider: 'bocha' })）。
     */
    providerOrder: ['duckduckgo', 'searxng', 'seam'],
    /**
     * 站点清单与「站点优先检索」（第三十四轮新增）。
     *
     * 语义三层，别混：
     *   enabled  总开关。关掉 = 检索完全按原来的路子走，不看清单。
     *   fallback 限定站点一无所获时退回不限定来源的泛搜（用户要求「被墙了就不纠结」）。
     *   entries  清单本体。`undefined` = 没配过 → 用内置目录；`[]` = 用户把行删空了
     *            → 这次真的不限定任何站点（不再兜回内置目录，否则「删空」在界面上等于没生效）。
     *
     * maxPerCall 是一次最多把几条查询限定到站点上：限定会多花请求，所以默认只 4 条。
     */
    sites: {
        enabled: true,
        fallback: true,
        maxPerCall: 4,
        entries: undefined,
    },
    /**
     * 出口代理（第二十九轮新增）：留空 = 不动本进程原有的出口（直连，或宿主启动时
     * 从环境变量装好的那条代理）。
     *
     * 填了优先装成**本插件私有**的分派器（宿主装了 undici 时；只有本插件的请求走它，
     * 改值 / 清空当场生效）；取不到 undici 才退回 Node 24+ 的
     * `http.setGlobalProxyFromEnv()` —— 那条是**进程级**、会顶掉宿主按环境变量装的
     * 全局 dispatcher、代理值也会进 process.env 被子进程继承，而且**只能装不能清**
     * （清空要重启 DSH 才回直连）。两条路都让回环地址直连。细节见 src/web-proxy.js。
     */
    proxy: '',
    /**
     * 每条通道各自的参数。**每条的键集就是它真正用得上的那几个**：
     * 接缝不吃参数；检索 API 只需要 Key；Anthropic / OpenAI 兼容还要端点与模型；
     * 自建 SearXNG 只要实例地址。这样设置页里不会出现「DuckDuckGo 的模型名」
     * 这种没有意义的输入框。
     */
    providers: {
        seam: {},
        anthropic: { apiKey: '', apiKeyEnv: 'DEEPSEEK_API_KEY', baseURL: 'https://api.deepseek.com/anthropic/v1', model: 'deepseek-v4-flash', timeoutMs: 60_000 },
        openai: { apiKey: '', apiKeyEnv: 'OPENAI_API_KEY', baseURL: 'https://api.openai.com/v1', model: 'gpt-4o-mini', timeoutMs: 60_000 },
        tavily: { apiKey: '', apiKeyEnv: 'TAVILY_API_KEY', timeoutMs: 30_000 },
        brave: { apiKey: '', apiKeyEnv: 'BRAVE_API_KEY', timeoutMs: 30_000 },
        bocha: { apiKey: '', apiKeyEnv: 'BOCHA_API_KEY', timeoutMs: 30_000 },
        exa: { apiKey: '', apiKeyEnv: 'EXA_API_KEY', timeoutMs: 30_000 },
        serper: { apiKey: '', apiKeyEnv: 'SERPER_API_KEY', timeoutMs: 30_000 },
        searxng: { baseURL: 'http://127.0.0.1:8080', timeoutMs: 30_000 },
        duckduckgo: { baseURL: 'https://html.duckduckgo.com/html', timeoutMs: 30_000 },
    },
    /**
     * 网页预处理（第二十轮新增）：取回来的页面先清洗再进上下文。
     *   article（默认）去样板 + 挑主容器 + 行级清洗；
     *   plain           只做行级清洗；
     *   off             原样返回（回到第十八轮的行为）。
     */
    preprocess: { ...DEFAULT_PREPROCESS },
    builtin: { ...DEFAULT_BUILTIN_WEB },
};

/** 检索引擎的取值。 */
export const SEARCH_ENGINES = ['auto', 'subagent', 'builtin'];

/**
 * Python 计算与绘图通道（office.python）的默认参数。
 *
 * 默认**开启**：这是办公模式里唯一的科学计算与绘图出口（命令行是关掉的），
 * 关掉它等于把「画一张数据图」重新变成手工活。关掉只是让 office.python 报错，
 * 不删任何东西。
 */
export const DEFAULT_PYTHON = {
    enabled: true,
    /** 解释器：留空自动探测（python / python3 / py），也可以填绝对路径。 */
    bin: '',
    timeoutMs: DEFAULT_PYTHON_TIMEOUT_MS,
    /** 产物目录（缓存目录内）：脚本写出的图与表都落这里。 */
    outDir: DEFAULT_PYTHON_OUT_DIR,
};

/**
 * 音频与视频的内容提取（office.av）的默认参数。
 *
 * 默认**开启**，并按「设备上本来就有 ffmpeg 与 SenseVoice」的口径找它们：
 *   - ffmpeg / ffprobe：留空 = 先 PATH、再常见安装目录；
 *   - SenseVoice：默认指语音输入下载的那份模型目录（$DSH_HOME/speech-to-text/
 *     sensevoice/models/sensevoice-onnx）；
 *   - sherpa-onnx-node：留空 = 从宿主安装位置解析（随语音输入装在 profile 里）。
 *
 * 四样里缺任何一样都**明确报错**（office.av.check() 一次报全，并说清怎么补），
 * 不做静默降级：宁可说「没有 ffmpeg」，也不要假装转写出了一段空文字。
 */
export const DEFAULT_AV = {
    enabled: true,
    ffmpegPath: '',
    ffprobePath: '',
    /** 留空 = $DSH_HOME/speech-to-text/sensevoice/models/sensevoice-onnx。 */
    modelDir: '',
    /** 留空 = 同一棵树下的 models/silero/silero_vad.onnx。 */
    vadModel: '',
    /** 留空 = 从宿主 profile 解析 sherpa-onnx-node（绝对路径也可以写死，便于离线部署）。 */
    runtimePath: '',
    precision: 'int8',
    threads: 2,
    language: 'auto',
    /** 单块秒数：官方那条链一次 131 秒封顶，这里默认 120 秒留余量。 */
    chunkSeconds: AV_CHUNK_SECONDS,
    /**
     * 相邻块的重叠秒数（第三十一轮）：块不再是完全切开的两段，跨在切点上的短音
     * 能在同一块里解完。跨得更深的长句由 worker 的边界回退兜住，不靠调大这个数。
     */
    overlapSeconds: AV_CHUNK_OVERLAP_SECONDS,
    /** 单次处理的音频时长上限（秒）：超过就报错，不把内存吃满。 */
    maxSeconds: AV_MAX_SECONDS,
    timeoutMs: AV_TIMEOUT_MS,
    /** 解码出来的规范 WAV 与抽出来的帧都只落缓存目录。 */
    audioDir: 'av/audio',
    framesDir: 'av/frames',
    /** 单次抽帧数上限与默认张数。 */
    maxFrames: AV_MAX_FRAMES,
    frames: AV_DEFAULT_FRAMES,
    frameFormat: 'jpg',
    /** 抽帧缩放宽度（0 = 原始分辨率）。 */
    frameWidth: 0,
    vadThreshold: 0.5,
    minSpeechSeconds: 0.25,
    minSilenceSeconds: 0.5,
    maxSegmentSeconds: 30,
    /** 是否带词级时间戳（默认不带：体积大，逐字稿用句子时间戳已经够用）。 */
    words: false,
};

function bool(value, fallback) {
    return typeof value === 'boolean' ? value : fallback;
}

/** 枚举收敛：不认识的值退回默认（配置写坏了不该弄崩插件）。 */
function oneOf(value, allowed, fallback) {
    return allowed.includes(value) ? value : fallback;
}

/** 记忆层拓扑：三个开关，缺省跟随默认，未知键忽略。 */
function resolveLayers(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const layers = {};
    for (const [name, fallback] of Object.entries(DEFAULT_LAYERS)) {
        layers[name] = bool(source[name], fallback);
    }
    return layers;
}

/**
 * 召回质量策略。strict-v1 的三个阈值必须保持 低 < 高，否则整个策略没有意义；
 * 配置里写反了就整体退回默认，而不是悄悄用一个自相矛盾的区间。
 */
function resolveRecallQuality(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const low = number(source.lowScoreThreshold, DEFAULT_RECALL_QUALITY.lowScoreThreshold, 0, 1);
    const high = number(source.highScoreThreshold, DEFAULT_RECALL_QUALITY.highScoreThreshold, 0, 1);
    if (!(low < high)) return { ...DEFAULT_RECALL_QUALITY };
    return {
        policy: oneOf(source.policy, ['strict-v1', 'off'], DEFAULT_RECALL_QUALITY.policy),
        lowScoreThreshold: low,
        highScoreThreshold: high,
        candidateMultiplier: positiveInt(source.candidateMultiplier, DEFAULT_RECALL_QUALITY.candidateMultiplier, 1, 10),
        maxMediumResults: positiveInt(source.maxMediumResults, DEFAULT_RECALL_QUALITY.maxMediumResults, 0, 40),
        maxUnknownResults: positiveInt(source.maxUnknownResults, DEFAULT_RECALL_QUALITY.maxUnknownResults, 0, 40),
    };
}

/**
 * 内置检索参数：字符串项给默认值，数值项在安全区间内收敛。
 *
 * 环境变量只做「有就覆盖、没有就用默认」：端点与模型名与宿主的
 * `DEEPSEEK_SEARCH_BASE_URL` 对齐，Key 名默认仍是 DEEPSEEK_API_KEY。
 */
function resolveBuiltinWeb(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const base = DEFAULT_BUILTIN_WEB;
    return {
        apiKey: text(source.apiKey),
        apiKeyEnv: text(source.apiKeyEnv) || text(process.env[BUILTIN_SEARCH_KEY_ENV]) || base.apiKeyEnv,
        baseURL: text(source.baseURL) || text(process.env[BUILTIN_SEARCH_BASE_URL_ENV])
            || text(process.env.DEEPSEEK_SEARCH_BASE_URL) || base.baseURL,
        model: text(source.model) || text(process.env[BUILTIN_SEARCH_MODEL_ENV]) || base.model,
        apiVersion: text(source.apiVersion) || base.apiVersion,
        maxUses: positiveInt(source.maxUses, base.maxUses, 1, 20),
        maxTokens: positiveInt(source.maxTokens, base.maxTokens, 256, 16_384),
        searchTimeoutMs: positiveInt(source.searchTimeoutMs, base.searchTimeoutMs, 5_000, 300_000),
        fetchTimeoutMs: positiveInt(source.fetchTimeoutMs, base.fetchTimeoutMs, 5_000, 300_000),
        maxResults: positiveInt(source.maxResults, base.maxResults, 1, 20),
        fetchPages: positiveInt(source.fetchPages, base.fetchPages, 0, 5),
        maxBytes: positiveInt(source.maxBytes, base.maxBytes, 32_768, 32 * 1024 * 1024),
        maxChars: positiveInt(source.maxChars, base.maxChars, 1_000, 400_000),
        maxRedirects: positiveInt(source.maxRedirects, base.maxRedirects, 0, 10),
        userAgent: text(source.userAgent) || base.userAgent,
        pdfMaxBytes: positiveInt(source.pdfMaxBytes, base.pdfMaxBytes, 1024 * 1024, 128 * 1024 * 1024),
        pdfMaxPages: positiveInt(source.pdfMaxPages, base.pdfMaxPages, 1, 500),
    };
}

/**
 * 每条检索通道的参数：**形状由 DEFAULT_SEARCH.providers[id] 的键集决定** ——
 * 每条通道只留它用得上的那几个键（Key / 端点 / 模型 / 超时），设置页与
 * schema 才不会有「DuckDuckGo 的模型名」这种输入框。
 *
 * 认不出来的通道名直接丢掉（写错一个不存在的通道不该让整份配置失效）；
 * 缺省值来自 DEFAULT_SEARCH.providers —— 那是「这条通道的默认 Key 名与端点」。
 *
 * `anthropic` 这一条要跟**老入口** `search.builtin` 合流：第十六～十九轮的配置
 * 只写了 builtin（apiKey / apiKeyEnv / baseURL / model / searchTimeoutMs），
 * 第二十轮起同一件事的正门是 `providers.anthropic`。两份都给了同名键时以
 * **providers.anthropic 为准**；只给了 builtin 时，builtin 的值必须真的生效 ——
 * 否则「配置了 Key 却仍报没 Key」这种 bug 会极其难查。
 */
function resolveProviders(raw, builtin) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const providers = {};
    for (const id of PROVIDER_IDS) {
        const fallback = DEFAULT_SEARCH.providers[id] ?? {};
        const given = source[id] !== null && typeof source[id] === 'object' ? source[id] : {};
        const legacy = id === 'anthropic' && builtin !== null && typeof builtin === 'object'
            ? {
                apiKey: builtin.apiKey,
                apiKeyEnv: builtin.apiKeyEnv,
                baseURL: builtin.baseURL,
                model: builtin.model,
                timeoutMs: builtin.searchTimeoutMs,
            }
            : {};
        const entry = {};
        for (const key of Object.keys(fallback)) {
            if (key === 'timeoutMs') {
                entry.timeoutMs = positiveInt(given.timeoutMs ?? legacy.timeoutMs, fallback.timeoutMs, 5_000, 300_000);
                continue;
            }
            const base = key === 'baseURL' ? fallback[key] : fallback[key];
            entry[key] = text(given[key]) || text(legacy[key]) || text(base);
            if (key === 'baseURL') entry[key] = entry[key].replace(/\/+$/, '');
        }
        providers[id] = entry;
    }
    return providers;
}

/** auto 时的通道顺序：丢掉认不出来的名字，空表退回默认顺序（抓取通道优先）。 */
function resolveProviderOrder(raw) {
    const list = Array.isArray(raw) ? raw.map((item) => text(item)).filter((item) => PROVIDER_IDS.includes(item)) : [];
    const unique = [...new Set(list)];
    return unique.length > 0 ? unique : [...DEFAULT_SEARCH.providerOrder];
}

/**
 * 站点清单（第三十四轮）：把设置里的四件事收敛成可用值。
 *
 * `entries` 三种形态要分开，别合并：
 *   缺省（undefined）→ 返回 undefined，交给 site-catalog 决定「用内置目录」；
 *   数组             → 逐条规范化（非法域名丢掉、重复域名去重）；
 *   其它类型         → 空数组（用户写坏了就当没站点，不猜）。
 */
function resolveSites(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const entries = source.entries === undefined
        ? undefined
        : (Array.isArray(source.entries) ? source.entries : []);
    return {
        enabled: bool(source.enabled, DEFAULT_SEARCH.sites.enabled),
        fallback: bool(source.fallback, DEFAULT_SEARCH.sites.fallback),
        maxPerCall: positiveInt(source.maxPerCall, DEFAULT_SEARCH.sites.maxPerCall, ...SITE_PRIORITY_LIMITS.maxPerCall),
        entries,
    };
}

/** 预处理模式：认不出来的模式与坏值一律退回默认。 */function resolvePreprocess(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const mode = PREPROCESS_MODES.includes(source.mode) ? source.mode : DEFAULT_PREPROCESS.mode;
    return {
        mode,
        dropBoilerplate: bool(source.dropBoilerplate, DEFAULT_PREPROCESS.dropBoilerplate),
        dedupeLines: bool(source.dedupeLines, DEFAULT_PREPROCESS.dedupeLines),
        minLineChars: positiveInt(source.minLineChars, DEFAULT_PREPROCESS.minLineChars, 0, 80),
        keepTitle: bool(source.keepTitle, DEFAULT_PREPROCESS.keepTitle),
    };
}

/** 每回合配额：0 表示该类不限制。 */
function resolveQuota(raw) {    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const quota = {};
    for (const [name, fallback] of Object.entries(DEFAULT_QUOTA)) {
        quota[name] = positiveInt(source[name], fallback, 0, 50);
    }
    return quota;
}

function number(value, fallback, min, max) {
    const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

/**
 * 音频与视频通道的参数：字符串项给默认值（留空 = 自动探测），
 * 数值项在安全区间内收敛，枚举不认识就退回默认。
 */
function resolveAv(raw) {
    const source = raw !== null && typeof raw === 'object' ? raw : {};
    const chunkSeconds = positiveInt(source.chunkSeconds, DEFAULT_AV.chunkSeconds, 10, 600);
    return {
        enabled: bool(source.enabled, DEFAULT_AV.enabled),
        ffmpegPath: text(source.ffmpegPath) || text(process.env[AV_FFMPEG_ENV]),
        ffprobePath: text(source.ffprobePath),
        modelDir: text(source.modelDir) || text(process.env[AV_MODEL_DIR_ENV]),
        vadModel: text(source.vadModel),
        runtimePath: text(source.runtimePath),
        precision: oneOf(source.precision, ['int8', 'fp32'], DEFAULT_AV.precision),
        threads: positiveInt(source.threads, DEFAULT_AV.threads, 1, 16),
        language: oneOf(source.language, AV_LANGUAGES, DEFAULT_AV.language),
        chunkSeconds,
        // 重叠上限是半块（与 av.js 的 avSettings 同一条判据）：再大就不是重叠，
        // 而是「同一段音频解两遍」，重叠区里的话会被两块各报一次。
        overlapSeconds: Math.min(number(source.overlapSeconds, DEFAULT_AV.overlapSeconds, 0, 30), chunkSeconds / 2),
        maxSeconds: positiveInt(source.maxSeconds, DEFAULT_AV.maxSeconds, 10, 21_600),
        timeoutMs: positiveInt(source.timeoutMs, DEFAULT_AV.timeoutMs, 5_000, 1_800_000),
        audioDir: text(source.audioDir) || DEFAULT_AV.audioDir,
        framesDir: text(source.framesDir) || DEFAULT_AV.framesDir,
        maxFrames: positiveInt(source.maxFrames, DEFAULT_AV.maxFrames, 1, 60),
        frames: positiveInt(source.frames, DEFAULT_AV.frames, 1, 60),
        frameFormat: oneOf(source.frameFormat, ['jpg', 'png'], DEFAULT_AV.frameFormat),
        frameWidth: positiveInt(source.frameWidth, DEFAULT_AV.frameWidth, 0, 4_096),
        vadThreshold: number(source.vadThreshold, DEFAULT_AV.vadThreshold, 0, 1),
        minSpeechSeconds: number(source.minSpeechSeconds, DEFAULT_AV.minSpeechSeconds, 0, 30),
        minSilenceSeconds: number(source.minSilenceSeconds, DEFAULT_AV.minSilenceSeconds, 0.01, 30),
        maxSegmentSeconds: number(source.maxSegmentSeconds, DEFAULT_AV.maxSegmentSeconds, 1, 120),
        words: bool(source.words, DEFAULT_AV.words),
    };
}

/**
 * 校验并补全配置。
 *
 * 配置来源有两层：组合里的 config（composition 层）与设置页写入的
 * settings（用户层）。两层都在这里收敛成同一份 resolved，运行期只认
 * 这一份 —— 避免「设置页改了但某处还在读旧值」。
 *
 * @returns {{cacheDir: string, scriptTimeoutMs: number, defaultTheme: string, maxScriptChars: number, injectGuide: boolean, tools: object, search: object, memory: object, python: object, keepCache: boolean}}
 */
export function resolveConfig(raw) {
    if (raw !== undefined && raw !== null && (typeof raw !== 'object' || Array.isArray(raw))) {
        throw new Error('dsh-office-mode 配置必须是对象。');
    }
    const config = raw ?? {};
    for (const key of ['cacheDir', 'defaultTheme', 'texTemplateDir']) {
        if (config[key] !== undefined && typeof config[key] !== 'string') {
            throw new Error(`${key} 必须是字符串。`);
        }
    }
    if (config.injectGuide !== undefined && typeof config.injectGuide !== 'boolean') {
        throw new Error('injectGuide 必须是布尔值。');
    }

    const cacheDir = text(config.cacheDir) || text(process.env[CACHE_DIR_ENV]) || '.office/cache';
    // 默认主题 = 素色网格：不传 theme 时产出「什么都没加」的文档，
    // 要配色得显式选主题（配置项或 create({theme})）。
    const themeRaw = (text(config.defaultTheme) || text(process.env[THEME_ENV]) || DEFAULT_THEME_ID).toLowerCase();
    const known = THEMES.map((theme) => theme.id);
    if (!known.includes(themeRaw)) {
        throw new Error(`defaultTheme 只支持：${known.join(' / ')}`);
    }
    // 工具开关：缺省 = 全开。未知键忽略（设置页字段增删不会让旧配置失效）。
    const rawTools = config.tools !== null && typeof config.tools === 'object' ? config.tools : {};
    const tools = {};
    for (const [name, fallback] of Object.entries(DEFAULT_TOOLS)) {
        tools[name] = bool(rawTools[name], fallback);
    }

    // 检索编排参数：同样在安全区间内收敛。
    const rawSearch = config.search !== null && typeof config.search === 'object' ? config.search : {};
    const builtin = resolveBuiltinWeb(rawSearch.builtin);
    const search = {
        engine: oneOf(rawSearch.engine, SEARCH_ENGINES, DEFAULT_SEARCH.engine),
        maxParallel: positiveInt(rawSearch.maxParallel, DEFAULT_SEARCH.maxParallel, 1, 8),
        resultLimit: positiveInt(rawSearch.resultLimit, DEFAULT_SEARCH.resultLimit, 1, 40),
        maxChannels: positiveInt(rawSearch.maxChannels, DEFAULT_SEARCH.maxChannels, 1, 12),
        outputDir: text(rawSearch.outputDir) || DEFAULT_SEARCH.outputDir,
        requireCrossSource: bool(rawSearch.requireCrossSource, DEFAULT_SEARCH.requireCrossSource),
        fallbackOnPlatformError: bool(rawSearch.fallbackOnPlatformError, DEFAULT_SEARCH.fallbackOnPlatformError),
        provider: oneOf(rawSearch.provider, ['auto', ...PROVIDER_IDS], DEFAULT_SEARCH.provider),
        providerOrder: resolveProviderOrder(rawSearch.providerOrder),
        proxy: text(rawSearch.proxy),
        sites: resolveSites(rawSearch.sites),
        providers: resolveProviders(rawSearch.providers, builtin),
        preprocess: resolvePreprocess(rawSearch.preprocess),
        builtin,
    };

    // 三层记忆：目录、容量、自动台账、是否往提示里注入。全部有缺省值，
    // 写坏了也只是退回默认（除了 dir 类型不对会明确报错）。
    const rawMemory = config.memory !== null && typeof config.memory === 'object' ? config.memory : {};
    if (rawMemory.dir !== undefined && typeof rawMemory.dir !== 'string') {
        throw new Error('memory.dir 必须是字符串。');
    }
    if (rawMemory.globalDir !== undefined && typeof rawMemory.globalDir !== 'string') {
        throw new Error('memory.globalDir 必须是字符串。');
    }
    const memory = {
        enabled: bool(rawMemory.enabled, DEFAULT_MEMORY.enabled),
        dir: text(rawMemory.dir) || text(process.env[MEMORY_DIR_ENV]) || DEFAULT_MEMORY.dir,
        autoLedger: bool(rawMemory.autoLedger, DEFAULT_MEMORY.autoLedger),
        promptHint: bool(rawMemory.promptHint, DEFAULT_MEMORY.promptHint),
        userLimitBytes: positiveInt(rawMemory.userLimitBytes, DEFAULT_MEMORY.userLimitBytes, 512, 65_536),
        projectLimitBytes: positiveInt(rawMemory.projectLimitBytes, DEFAULT_MEMORY.projectLimitBytes, 1_024, 262_144),
        // 0 = 不限（第四十八轮 P0-2）：投影预算与「库里能放多少」是两件事，所以下限给 0
        projectionBudgetBytes: positiveInt(rawMemory.projectionBudgetBytes, DEFAULT_MEMORY.projectionBudgetBytes, 0, 131_072),
        // 0 = 关掉（新-11）：硬上限与建议线也允许关 —— 有人就是要把长篇结论放热记忆里
        entryLimitBytes: positiveInt(rawMemory.entryLimitBytes, DEFAULT_MEMORY.entryLimitBytes, 0, 131_072),
        entryHintBytes: positiveInt(rawMemory.entryHintBytes, DEFAULT_MEMORY.entryHintBytes, 0, 131_072),
        ledgerLimit: positiveInt(rawMemory.ledgerLimit, DEFAULT_MEMORY.ledgerLimit, 10, 10_000),
        archiveKeep: positiveInt(rawMemory.archiveKeep, DEFAULT_MEMORY.archiveKeep, 1, 1_000),
        layers: resolveLayers(rawMemory.layers),
        scope: oneOf(rawMemory.scope, MEMORY_SCOPES, DEFAULT_MEMORY.scope),
        globalDir: text(rawMemory.globalDir) || text(process.env[MEMORY_GLOBAL_DIR_ENV]) || DEFAULT_MEMORY.globalDir,
        userScope: oneOf(rawMemory.userScope, USER_SCOPES, DEFAULT_MEMORY.userScope),
        links: bool(rawMemory.links, DEFAULT_MEMORY.links),
        autoCapture: bool(rawMemory.autoCapture, DEFAULT_MEMORY.autoCapture),
        recallQuality: resolveRecallQuality(rawMemory.recallQuality),
        quota: resolveQuota(rawMemory.quota),
    };

    // Python 计算与绘图通道：解释器、超时、产物目录。空配置 = 自动探测解释器、
    // 默认超时、产物落缓存目录下的 python/out。
    const rawPython = config.python !== null && typeof config.python === 'object' ? config.python : {};
    if (rawPython.bin !== undefined && typeof rawPython.bin !== 'string') {
        throw new Error('python.bin 必须是字符串。');
    }
    if (rawPython.outDir !== undefined && typeof rawPython.outDir !== 'string') {
        throw new Error('python.outDir 必须是字符串。');
    }
    const python = {
        enabled: bool(rawPython.enabled, DEFAULT_PYTHON.enabled),
        bin: text(rawPython.bin) || text(process.env[PYTHON_BIN_ENV]),
        timeoutMs: positiveInt(rawPython.timeoutMs, DEFAULT_PYTHON.timeoutMs, 5_000, 900_000),
        outDir: text(rawPython.outDir) || text(process.env[PYTHON_OUT_DIR_ENV]) || DEFAULT_PYTHON.outDir,
    };

    const rawSubagentModel = config.subagentModel !== null && typeof config.subagentModel === 'object' ? config.subagentModel : {};
    const subagentModel = {
        mode: oneOf(rawSubagentModel.mode, ['inherit', 'fixed'], DEFAULT_SUBAGENT_MODEL.mode),
        provider: text(rawSubagentModel.provider),
        model: text(rawSubagentModel.model),
    };

    // 音频与视频的内容提取：四样能力（ffmpeg / ffprobe / 模型 / 运行时）都不在
    // 这里校验存在性 —— 解析配置不该碰磁盘（空配置永远可用的那条硬约束）。
    // 存在性由 office.av.check() 与每次调用前的探测负责。
    const av = resolveAv(config.av);

    return {
        cacheDir,
        scriptTimeoutMs: positiveInt(config.scriptTimeoutMs, DEFAULT_SCRIPT_TIMEOUT_MS, 1_000, 600_000),
        defaultTheme: themeRaw,
        maxScriptChars: positiveInt(config.maxScriptChars, DEFAULT_MAX_SCRIPT_CHARS, 1_000, 4_000_000),
        injectGuide: config.injectGuide === true,
        // keepCache：调用结束后**是否保留**缓存目录里的中间文件。
        // 2026-09-23 起默认保留 —— 清掉就没有跨调用复用（PDF 页面图渲染一次、
        // 下一步 read_image 就读不到了）。要恢复「每次调用结束都清空」就设 false。
        keepCache: bool(config.keepCache, true),
        cacheTtlMinutes: positiveInt(config.cacheTtlMinutes, DEFAULT_CACHE_TTL_MINUTES, 0, 43_200),
        cacheMaxBytes: positiveInt(config.cacheMaxBytes, DEFAULT_CACHE_MAX_BYTES, 1_000_000, 8 * 1024 * 1024 * 1024),
        pdfDpi: positiveInt(config.pdfDpi, DEFAULT_PDF_DPI, 36, 400),
        pdfMaxPages: positiveInt(config.pdfMaxPages, DEFAULT_PDF_MAX_PAGES, 1, 500),
        pdfEngine: pdfEngine(config.pdfEngine),
        // LaTeX：编译入口、超时、模板目录。texTemplateDir 留空时自动探测
        // （配置 → 环境变量 → 工作目录里的 thuthesis* → TeX Live 自带）。
        texEngine: texEngine(config.texEngine),
        texTimeoutMs: positiveInt(config.texTimeoutMs, DEFAULT_TEX_TIMEOUT_MS, 30_000, 1_800_000),
        texTemplateDir: text(config.texTemplateDir) || text(process.env[TEX_TEMPLATE_ENV]),
        tools,
        search,
        memory,
        python,
        av,
        subagentModel,
    };
}

/** PDF 引擎名：auto 或探测表里认识的 id，其余按 auto 处理（不弄崩配置）。 */
function pdfEngine(value) {
    const wanted = (text(value) || text(process.env[PDF_ENGINE_ENV]) || DEFAULT_PDF_ENGINE).toLowerCase();
    const known = ['auto', 'fitz', 'pdftoppm', 'pdftocairo', 'mutool', 'gs', 'gswin64c'];
    return known.includes(wanted) ? wanted : DEFAULT_PDF_ENGINE;
}

/** LaTeX 编译入口：auto 或 latexmk / xelatex / lualatex，其余按 auto 处理。 */
function texEngine(value) {
    const wanted = (text(value) || text(process.env[TEX_ENGINE_ENV]) || DEFAULT_TEX_ENGINE).toLowerCase();
    const known = ['auto', 'latexmk', 'xelatex', 'lualatex'];
    return known.includes(wanted) ? wanted : DEFAULT_TEX_ENGINE;
}
