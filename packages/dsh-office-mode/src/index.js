/**
 * dsh-office-mode —— 办公模式插件。
 *
 * 一个插件行带来三件事：
 *   1. 工具面（极简：默认七个，写法和细节按需取）；
 *   2. 按需可开的办公模式提示段落（injectGuide，默认关闭 —— 「模式」这件事
 *      交给 preset 决定，插件本身不往别人的会话里塞提示词）；
 *   3. 设置页（Settings > 插件 > 办公模式）：工具开关与参数调节。
 *
 * 「办公模式」这个模式本身由本地 preset `office` 组合：preset 负责说话方式
 * 与工具目录，插件负责能力。两者分开的好处是：换说话方式不用动引擎，加格式
 * 不用动 preset。
 *
 * @module dsh-office-mode
 */
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';

import { CAPABILITY_VARIABLE, renderCapabilityMap, syncCapabilities } from './capabilities.js';
import { resolveConfig } from './config.js';
import { OFFICE_GUIDE } from './guide.js';
import { MEMORY_PROMPT_HINT } from './memory.js';
import { createOfficeSettings, OFFICE_SETTINGS_NS } from './settings.js';
import { subagentsOf } from './search.js';
import { buildTools, officeRunErrorResult } from './tools.js';
import { registerMemoryView } from './view.js';
import { disposeProxy, installProxy } from './web-proxy.js';

export const name = 'dsh-office-mode';
export const inject = ['tools'];

/** 提示段落的排序位；只在本插件被显式要求注入时才用得上。 */
const GUIDE_SECTION_ORDER = 500;
/** 记忆说明段落的排序位：排在办公约定之后，属于补充说明。 */
const MEMORY_HINT_SECTION_ORDER = 520;

/**
 * 同步从宿主的安装位置解析一个包。
 *
 * 为什么要「同步」：`Config` 必须是模块求值时就已经存在的导出（cordis 只在
 * runtime.Config 为真值时才用它校验 config）。而本插件是 link: 进 profile 的
 * （源码在工作区），ESM 按**模块自身所在目录**解析裸包名，工作区里没有
 * @deepseek-ai/*，所以 `import '@deepseek-ai/schemastery'` 必然
 * ERR_MODULE_NOT_FOUND —— 只能借宿主的解析位置同步 require。
 *
 * 锚点按可靠性排序：profile 根（profile 启动时一定有）→ 宿主安装位置 →
 * 宿主入口所在目录。全部失败返回 undefined，调用方各自降级。
 */
const hostModules = new Map();

function resolveHostModule(specifier) {
    if (hostModules.has(specifier)) return hostModules.get(specifier);
    const anchors = [
        process.env.DSH_PROFILE_DIR,
        process.env.DSH_HOST_ROOT,
        process.env.DSH_CHECKOUT,
    ].filter((value) => typeof value === 'string' && value !== '');
    if (typeof process.argv[1] === 'string' && process.argv[1] !== '') {
        anchors.push(dirname(process.argv[1]));
    }
    let found;
    for (const anchor of anchors) {
        try {
            const req = createRequire(pathToFileURL(join(anchor, 'package.json')).href);
            found = req(req.resolve(specifier));
            break;
        } catch {
            // 换下一个锚点
        }
    }
    const value = found === undefined ? undefined : (found.default ?? found);
    hostModules.set(specifier, value);
    return value;
}

const schemastery = resolveHostModule('@deepseek-ai/schemastery');
const cosmokit = resolveHostModule('@deepseek-ai/cosmokit');

/**
 * 插件 Config：设置页的可编辑字段全靠它。
 *
 * 宿主 dsh-settings 只把「声明了 Config 的插件行」投影成设置命名空间，且写入
 * 只接受 volatile 字段（schema 里已整片标好，见 settings.js 的
 * markLeavesVolatile）。拿不到 schemastery 时退回「无 schema」—— cordis 的
 * resolveConfig 会把 config 原样透传，插件其余功能照常。
 *
 * **整段包在 try/catch 里**：这是模块求值期的代码，抛一次就是「插件加载失败、
 * web boot 少一个 entry」；而设置页少一份 schema 只是降级。两件事的代价差得
 * 太远，所以这里宁可吃掉异常也不让它冒出去。
 */
function buildConfig() {
    if (schemastery === undefined) return undefined;
    try {
        return createOfficeSettings(schemastery);
    } catch (error) {
        return undefined;
    }
}

export const Config = buildConfig();

/**
 * 把 config 里的 Volatile 引用拆成纯值。
 *
 * volatile 字段解析出来的是引用（宿主就地更新、不重挂插件），而本插件的
 * resolveConfig 与下游全部按纯值读，所以每次重建前都要拆一层。拿不到
 * cosmokit 时不可能有 volatile 字段（那意味着连 schemastery 都没有），
 * 原样返回即可。
 *
 * 单点失败只降级不抛错：某一个 `.get()` 出问题就退回 undefined，让下游的
 * resolveConfig 用自己的默认值兜底。
 *
 * @param {unknown} value 配置值（可能含 Volatile 引用）
 */
function unwrapVolatile(value) {
    if (cosmokit !== undefined && cosmokit.isVolatile(value)) {
        try {
            return unwrapVolatile(value.get());
        } catch {
            return undefined;
        }
    }
    if (Array.isArray(value)) return value.map((item) => unwrapVolatile(item));
    if (value !== null && typeof value === 'object') {
        const result = {};
        for (const [key, child] of Object.entries(value)) result[key] = unwrapVolatile(child);
        return result;
    }
    return value;
}

/**
 * 把任意配置解析成运行期配置，解析失败退回空配置。
 *
 * resolveConfig 自带完整默认值，所以「空配置」永远是可用的那一份 —— 这正是
 * 「空配置永远可用」这条约束的兜底：配置坏了不该让插件起不来。
 *
 * @param {unknown} raw 原始配置（可能含 volatile 引用）
 */
function safeResolveConfig(raw) {
    try {
        return resolveConfig(unwrapVolatile(raw));
    } catch (error) {
        return resolveConfig({});
    }
}

export function apply(ctx, config = {}) {
    // resolved 是运行期唯一真源：组合层 config 与设置页用户层在这里收敛。
    // 设置页的写入拼进 profile 里本插件那一行的用户层，宿主按 volatile 就地
    // 更新引用，所以每次重建都重新拆引用 + 重新解析。
    let resolved = safeResolveConfig(config);
    let disposers = [];

    const disposeAll = () => {
        for (const dispose of disposers) {
            try {
                dispose();
            } catch {
                // 释放失败不应影响其它资源回收。
            }
        }
        disposers = [];
    };

    const registerTools = () => {
        for (const definition of buildTools(resolved, ctx)) {
            disposers.push(ctx.tools.register(definition));
        }
    };

    /**
     * 能力快照 + 把它注入 persona（`{{office_capabilities}}`）。
     *
     * 为什么非要变量不可：办公 preset 的 persona 是 `complete: true` 的**唯一**提示段，
     * 宿主 dsh-system-prompt 在装配最后会把 sections 收敛成那一段 —— 插件注册的
     * section（guide / 记忆说明）在办公会话里根本进不了提示词，只有 `{{变量}}` 的插值
     * 发生在收集之后的 renderPrompt 里，能跟着 persona 一起到模型面前。
     *
     * 代价写清楚：preset 引用了一个只有本插件才注册的变量，所以**办公 preset 必须与
     * 本插件同装**（两者本来就是一对，见 README）。缺了插件时宿主会明确报
     * `unknown prompt variable "{{office_capabilities}}"`，不是静默少一段。
     */
    let capabilitySnapshot;
    const refreshCapabilities = (options = {}) => {
        try {
            capabilitySnapshot = syncCapabilities({ config: resolved, refresh: options.refresh === true });
        } catch {
            // 探测失败不该让插件起不来：地图会照实说「没探到」。
            capabilitySnapshot = undefined;
        }
    };

    /**
     * 组合里有没有**能用的**子代理服务（决定 `office_search_dispatch` 走哪条路）。
     *
     * 用 search.js 的 `subagentsOf` 而不是自己写 `ctx.get('subagents') !== undefined`：
     * 那一个还查 `service.start`、并把 `ctx.get` 的抛错吃掉（第四十七轮复核 F1 ——
     * 弱的那个判据会让地图说「走子代理」，而派工回执给的是「没有 subagents 服务」）。
     */
    const subagentsPresent = () => subagentsOf(ctx) !== undefined;

    const registerCapabilityMap = () => {
        const systemPrompt = typeof ctx.get === 'function' ? ctx.get('systemPrompt') : undefined;
        if (systemPrompt === undefined || typeof systemPrompt.variable !== 'function') {
            // 这一段不能像 guide / 记忆说明那样静默跳过：preset 的 persona 引用了这个变量，
            // 没注册的话办公会话**每一次装配都会抛** unknown prompt variable。组合顺序出问题
            // 时，这条警告是唯一能指出原因的线索。
            ctx.logger?.warn?.('[dsh-office-mode] 当前组合里没有 systemPrompt 服务：'
                + '{{' + CAPABILITY_VARIABLE + '}} 未注册 —— 办公 preset 的 persona 会因未知变量装配失败。');
            return;
        }
        // provider 在**每一次装配**时求值（每个请求一次），而它现在成了办公会话能不能装配
        // 的前置：探测再抛错也不能把请求带下去（复核 F3）。快照没探到时顺手带 refresh 重探，
        // 免得拿别的配置留下的旧快照（复核 F2）。
        disposers.push(systemPrompt.variable(CAPABILITY_VARIABLE, () => {
            let snapshot = capabilitySnapshot;
            try {
                if (snapshot === undefined) snapshot = syncCapabilities({ config: resolved, refresh: true });
                return renderCapabilityMap(snapshot, { subagents: subagentsPresent() });
            } catch {
                return renderCapabilityMap(undefined, { subagents: subagentsPresent() });
            }
        }));
    };

    /**
     * 把「office_run 脚本失败」变成一个**宿主认识的失败结果**（第十八轮 P0-1）。
     *
     * 宿主调度器只认 `isError`：工具体正常返回就一律算成功。脚本失败在插件这一侧
     * 是 `ok:false`，宿主看不见 —— 会话日志里的失败率因此被低估（第十八轮实测
     * 5/25 漏记）。这里挂在宿主留的 `tools/execute` 环绕点上（`dsh-timeout` 那种
     * 策略插件用的就是同一个点），只对 office_run 且 `ok === false` 的结果动手。
     *
     * 要点：
     *   - 只认 `exec.name === 'office_run'`，别的工具原样透传；
     *   - 已经是失败（宿主自己抛的错）就不动，不把错误再包一层；
     *   - 不改脚本返回值本身，只替换这一次调用的结果形状（见 tools.js 的
     *     officeRunErrorResult），所以 office_run 的内部契约与测试不受影响。
     */
    const registerRunErrorGuard = () => {
        if (typeof ctx.on !== 'function') return;
        disposers.push(ctx.on('tools/execute', async (exec, next) => {
            const result = await next();
            if (exec?.name !== 'office_run') return result;
            if (result === null || typeof result !== 'object' || result.isError === true) return result;
            const failure = officeRunErrorResult(result.value);
            if (failure === null) return result;
            // 结果里已经渲染好的那段文本最完整（含失败前写出的文件与全部告警），用它；
            // meta 与 additionalContexts 要跟着带过去 —— 宿主的 isError 分支只保留
            // content / error / meta / additionalContexts，漏掉就等于把展示元数据丢了。
            return {
                ...failure,
                ...(Array.isArray(result.content) && result.content.length > 0 ? { content: result.content } : {}),
                ...(result.meta !== undefined ? { meta: result.meta } : {}),
                ...(result.additionalContexts !== undefined ? { additionalContexts: result.additionalContexts } : {}),
            };
        }));
    };

    const registerGuide = () => {
        if (!resolved.injectGuide) return;
        const systemPrompt = typeof ctx.get === 'function' ? ctx.get('systemPrompt') : undefined;
        if (systemPrompt !== undefined && typeof systemPrompt.section === 'function') {
            disposers.push(systemPrompt.section({
                name: 'office-mode-guide',
                order: GUIDE_SECTION_ORDER,
                text: OFFICE_GUIDE,
            }));
        } else {
            ctx.logger?.warn?.('[dsh-office-mode] injectGuide 已开启，但当前组合里没有 systemPrompt 服务，提示段落未注册。');
        }
    };

    /**
     * 注册「工作目录里有记忆」这段静态说明。
     *
     * 这是 mnemon 关掉之后顶替它那层「每轮热记忆提醒」的东西：只讲位置与用法，
     * 不含任何记忆内容（section.text 只接受同步函数，读盘是异步的）。
     * 组合里没有 systemPrompt 服务时静默跳过 —— 这一段是默认开启的，
     * 每次启动都报一次警告会变成噪声。
     */
    const registerMemoryHint = () => {
        if (!resolved.memory.enabled || !resolved.memory.promptHint) return;
        if (resolved.tools.office_memory === false) return;
        const systemPrompt = typeof ctx.get === 'function' ? ctx.get('systemPrompt') : undefined;
        if (systemPrompt === undefined || typeof systemPrompt.section !== 'function') return;
        disposers.push(systemPrompt.section({
            name: 'office-mode-memory',
            order: MEMORY_HINT_SECTION_ORDER,
            text: MEMORY_PROMPT_HINT,
        }));
    };

    /** 按当前 resolved 重建工具面与提示段落。设置页改动后走这里。 */
    const rebuild = () => {
        disposeAll();
        registerTools();
        registerRunErrorGuard();
        registerGuide();
        registerMemoryHint();
        // 设置改动可能换了 python.bin / ffmpeg / 模型目录，能力快照跟着重探。
        refreshCapabilities({ refresh: true });
        registerCapabilityMap();
    };

    /** 重读配置：拆掉 volatile 引用，重新解析成运行期配置。 */
    const refresh = () => {
        resolved = safeResolveConfig(config);
    };

    /**
     * 按当前设置装一次出口代理（第二十九轮）。
     *
     * 只装不改：装不上（运行时不支持 / 地址不合法）记一条警告就够了 —— 它是
     * 「让联网换条出口」的增益项，不该让插件激活失败。装了之后**清不掉**
     * （Node 的 setGlobalProxyFromEnv 只有装没有清），所以清空设置时这里只把
     * 「要重启才回直连」照实记下来，见 src/web-proxy.js。
     */
    const applyProxy = () => {
        // resolveHostModule 给的是「宿主那侧的 undici」：拿得到就建私有分派器
        // （只影响本插件），拿不到才退回进程级环境变量那条路。
        const result = installProxy(resolved?.search?.proxy ?? '', { resolveModule: resolveHostModule });
        if (result.ok !== true) {
            ctx.logger?.warn?.('[dsh-office-mode] 代理设置未生效：' + result.reason);
        }
    };

    registerTools();
    registerRunErrorGuard();
    registerGuide();
    registerMemoryHint();
    refreshCapabilities();
    registerCapabilityMap();
    applyProxy();

    // ── 记忆浏览面板的数据端点 ─────────────────────────────────────────────
    //
    // 只注册一次，且**不**进 disposers：rebuild（设置页改动）不该把路由拆了重建。
    // 处理器每次请求都读当前的 resolved.memory，所以配置改了立刻生效。
    //
    // 用 ctx.inject 而不是直接 ctx.get('webServer')：这样不必假设 webServer 一定
    // 在插件行之前就绪（组合顺序变了也不会静默少一条路由）。组合里根本没有
    // webServer（例如 headless）时回调不触发，插件其余功能照常。
    let disposeView = null;
    ctx.inject(['webServer'], (sctx) => {
        disposeView = registerMemoryView(sctx, { getMemory: () => resolved.memory }) ?? null;
    });

    // ── 设置变更 ──────────────────────────────────────────────────────────
    //
    // 0.1.7 起设置页的写入走宿主设置文档：拼进 profile 里本插件那一行的用户层，
    // 由 dsh-settings 落盘并广播 `settings/document-updated`。旧版那套
    // `settings.installSection(...) + onChange` 的契约已经不存在了。
    //
    // volatile 字段是就地更新的引用（宿主不重挂插件），所以这里必须自己听事件、
    // 重新解析配置再重建工具面。监听器**不进 disposers**：rebuild 会清空它。
    //
    // 整段包 try/catch：这条是「锦上添花」的即时生效链路，拿不到 settings 服务、
    // 事件名对不上、或注册被拒，都只该让设置改动等下次重挂才生效，**不该让插件
    // 激活失败**（激活失败会让整条 web boot 少一个 entry）。
    let disposeSettingsWatch = null;
    try {
        ctx.inject(['settings'], (sctx) => {
            if (typeof sctx.on !== 'function') return;
            disposeSettingsWatch = sctx.on('settings/document-updated', (ns) => {
                if (ns !== OFFICE_SETTINGS_NS) return;
                try {
                    refresh();
                    applyProxy();
                    rebuild();
                } catch {
                    // 重建失败不该把插件打挂：下一次工具调用仍用旧的那份配置。
                }
            });
        });
    } catch {
        disposeSettingsWatch = null;
    }

    if (typeof ctx.on === 'function') {
        ctx.on('dispose', () => {
            disposeAll();
            for (const teardown of [disposeView, disposeSettingsWatch]) {
                if (teardown === null) continue;
                try {
                    teardown();
                } catch {
                    // 释放失败不应影响其它资源回收。
                }
            }
            // 私有代理分派器持有连接池，卸载时要放掉；进程级那条不是本插件的资源，不碰。
            try {
                disposeProxy();
            } catch {
                // 同上：释放失败不该影响插件卸载。
            }
        });
    }
}

export { resolveConfig, buildTools };
export { OFFICE_GUIDE } from './guide.js';
export {
    createMemory,
    memoryPaths,
    globalMemoryDir,
    globalMemoryPaths,
    memoryStats,
    memoryExists,
    resolveMemoryConfig,
    applyRecallQuality,
    relevanceOf,
    renderDigest,
    renderRead,
    renderMutation,
    renderLog,
    renderLink,
    renderUnlink,
    renderRelated,
    renderEntities,
    renderExport,
    renderImport,
    renderStatus,
    renderKb,
    tierTag,
    conflictStateText,
    MEMORY_PROMPT_HINT,
    MEMORY_ACTIONS,
    LINK_KINDS,
    CONFLICT_CLASSES,
    CONFLICT_STATES,
    MEMORY_TIME_WORDS,
    MEMORY_TIME_HINTS,
    MEMORY_TIME_PATTERNS,
    DEFAULT_LAYERS,
    DEFAULT_RECALL_QUALITY,
    DEFAULT_QUOTA,
    MEMORY_SCOPES,
    USER_SCOPES,
    PACK_FORMAT,
    PACK_VERSION,
} from './memory.js';
export {
    splitDocument,
    chunkHeaderOf,
    chunkIdOf,
    chunkCharsFor,
    docIdOf,
    kbPathsFor,
    kbDirOf,
    kbCountsOf,
    kbLiveIdsOf,
    kbTierOf,
    normalizeKbChunk,
    normalizeManifestRow,
    KB_TIERS,
    KB_CHUNK_CHARS,
    KB_CHUNK_CHARS_BY_KIND,
    KB_MAX_DOC_BYTES,
    KB_READ_CHARS,
    KB_VERSION,
} from './kb.js';
export { createTurnQuota, turnKeyOf } from './quota.js';
export { registerMemoryView, buildSnapshot, noteWorkspace, knownWorkspaces, resetWorkspaces, registryWorkspacesOf, sessionWorkspacesOf, MEMORY_VIEW_PATH, VIEW_LIMITS } from './view.js';
export { migrateMnemon, renderMigration, readMnemonInsights, mnemonWorkspaceRoot, mnemonGlobalRoot, DEFAULT_MNEMON_DIR } from './migrate.js';
export { pdfEngines, pdfInfo, pdfPages, pdfText, probePdfEngines } from './pdf.js';
export {
    avCheck,
    avExtract,
    avFrames,
    avInfo,
    avTranscribe,
    avSettings,
    probeAv,
    resetAvProbe,
    inspectWav,
    canonicalizeWav,
    isCanonicalWav,
    planChunks,
    resolveChunkBoundary,
    planFrameTimes,
    joinTranscript,
    renderTranscript,
    formatClock,
    dshHome,
    defaultSenseVoiceModelDir,
    defaultSileroVadPath,
    AV_LANGUAGES,
    AV_ERROR_CODES,
    AV_WORKER_PATH,
    AV_CHUNK_OVERLAP_SECONDS,
    AV_AUDIO_EXTENSIONS,
    AV_VIDEO_EXTENSIONS,
} from './av.js';
export {
    pythonCheck,
    pythonFile,
    pythonRun,
    probePython,
    resetPythonProbe,
    DEFAULT_PYTHON_OUT_DIR,
    DEFAULT_PYTHON_TIMEOUT_MS,
    PYTHON_MAX_CODE_CHARS,
} from './python.js';
export { DEFAULT_PYTHON, PYTHON_BIN_ENV, PYTHON_OUT_DIR_ENV } from './config.js';
export {
    createWebAccess,
    engineDescription,
    mapSearchResponse,
    mergeSources,
    providerLabelOf,
    resolveApiKey,
    htmlToText,
    htmlTitle,
    decodeEntities,
    isPublicAddress,
    validateUrl,
    classifyContentType,
    looksLikePdf,
    defaultPdfReader,
    charsetOf,
    httpFetch,
    webSeamOf,
    WebAccessError,
    WEB_ENGINE_BUILTIN,
    WEB_ENGINE_SEAM,
} from './web.js';
export { PROVIDER_IDS, PROVIDERS, resolveProviderOptions, resolveProviderOrder, searchProvider } from './web-providers.js';
export {
    activeDispatcher,
    disposeProxy,
    installProxy,
    maskProxyUrl,
    normalizeProxyUrl,
    PROXY_ENV_KEYS,
    proxyStatus,
    shouldBypassProxy,
} from './web-proxy.js';
export { preprocessPage, resolvePreprocessOptions, extractMeta, PREPROCESS_MODES, DEFAULT_PREPROCESS } from './web-preprocess.js';
export { compileTex, texEngines, probeTexEngines, locateTemplateDir, escapeLatex } from './tex.js';
export {
    CAPABILITY_VARIABLE,
    CAPABILITY_MAP_BUDGET_BYTES,
    syncCapabilities,
    resetCapabilities,
    renderCapabilityMap,
} from './capabilities.js';
export { createOfficeSettings, OFFICE_SETTINGS_NS, SUBAGENT_TOOL_CATALOG, LIMITS } from './settings.js';