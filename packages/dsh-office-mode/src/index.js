/**
 * dsh-office-mode —— 办公模式插件。
 *
 * 一个插件行带来三件事：
 *   1. 工具面（极简：默认七个，写法和细节按需取）；
 *   2. 按需可开的办公模式提示段落（injectGuide，默认关闭 —— 「模式」这件事
 *      交给 preset 决定，插件本身不往别人的会话里塞提示词）；
 *   3. 设置页（Settings > 插件 > 办公模式）：工具开关与参数调节。
 *
 * 另有一层**按 preset 的守卫**（第五十轮）：本插件的工具注册在全局层，每个 preset
 * 的 agent 都继承得到，所以「极简模式里不该出现 office_*」这件事只能靠
 * `ctx.tools.restrict()` 按 agent 隐藏，不能靠 preset 里的禁用行 —— 理由与实测见
 * src/preset-guard.js。
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
import { kbTreeFolderPresence, registerKbScopeCommand } from './kb-session-scope.js';
import { createMemory, HOT_MEMORY_VARIABLE, MEMORY_PROMPT_HINT, readHotMemory } from './memory.js';
import { registerPresetPlugins } from './preset-plugins.js';
import { enabledOfficeToolNames, installPresetGuard } from './preset-guard.js';
import { reportRebuild } from './rebuild-state.js';
import { createOfficeSettings, OFFICE_SETTINGS_NS } from './settings.js';
import { subagentsOf } from './search.js';
import { buildTools, officeRunErrorResult } from './tools.js';
import { resolveHostOcrCredential } from './ocr-host-credential.js';
import { registerMemoryView } from './view.js';
import { registerPreviewView } from './preview-view.js';
import { disposeProxy, installProxy, maskProxyUrl } from './web-proxy.js';

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
    /** 最近一次 `installAll()` 走到哪一步；只在装失败时用来报「哪个阶段炸的」。 */
    let installStep = '';
    /** 重建事务正在跑（宿主的事件广播是同步且可重入的，见 `runSettingsRebuild`）。 */
    let rebuilding = false;
    /** 事务进行中又收到广播：事务结束后用最新的 config 补跑一次。 */
    let pendingRebuild = false;

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

    /**
     * 注册一批工具定义。
     *
     * 收「定义数组」而不是收配置：`buildTools()` 每份配置只该算**一次**（它内部建了一个
     * 跟着这份工具面走的每回合配额），而工具面现在要拆成两步装（见 `installAll`）。
     */
    const registerToolDefinitions = (definitions) => {
        for (const definition of definitions) disposers.push(ctx.tools.register(definition));
    };

    /**
     * 装上「极简模式里不出现 office_*」这层守卫（第五十轮，见 src/preset-guard.js）。
     *
     * 为什么每次 rebuild 都要重装：掩码里写的是**注册期现算**的工具名（restrict 会校验
     * 名字是不是已知的全局工具，设置页关掉的工具不注册、就不能出现在掩码里），所以
     * 工具面变了要重新算一遍 —— 重建后由 `presetGuard.resync()` 把已在跑的 agent 补齐。
     */
    let presetGuard = null;
    const registerPresetGuard = (active) => {
        presetGuard = installPresetGuard(ctx, {
            toolNames: () => enabledOfficeToolNames(active),
            // 宿主工具那一半（第五十四轮 新-15①）：**默认关着**，设置页打开之后才在
            // 办公 preset 里把预设声明要关的那批能力真正掩掉。清单写在 preset-guard.js，
            // `preset-check` 有一条断言核它与 preset 的 disabled 行不脱节。
            hostHidden: active?.guard?.hideDevToolsInOffice === true ? ['office'] : [],
        });
        disposers.push(presetGuard.dispose);
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
    const refreshCapabilities = (active, options = {}) => {
        try {
            capabilitySnapshot = syncCapabilities({ config: active, refresh: options.refresh === true });
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

    const registerCapabilityMap = (active) => {
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
        //
        // provider 闭包捕的是**这一次安装用的那份配置**（不是可变的 `resolved`）：回滚要把
        // 上一份配置重新装一遍，变量跟着自己那一份配置走，谁装的读谁（第五十九轮 S-04）。
        disposers.push(systemPrompt.variable(CAPABILITY_VARIABLE, () => {
            let snapshot = capabilitySnapshot;
            try {
                if (snapshot === undefined) snapshot = syncCapabilities({ config: active, refresh: true });
                return renderCapabilityMap(snapshot, { subagents: subagentsPresent() });
            } catch {
                return renderCapabilityMap(undefined, { subagents: subagentsPresent() });
            }
        }));
    };

    /**
     * 开场记忆变量（`{{office_hot_memory}}`）。
     *
     * 与能力地图同一个通道：persona 是 `complete: true`，插件 section 进不去，
     * 只有变量插值能把它带进系统提示。区别是内容按**会话工作目录**读：
     * provider 拿到 assembly 的 `agent.cwd`，从 `.office/memory/HOT.md` 同步取
     * 预先落盘的小段（写入记忆时生成），所以新会话第一轮就有，不必等 office_help。
     *
     * 读不到（没写过记忆 / 自定义目录还没建）就返回空串，不抛错：变量缺值会让
     * 整个提示装配失败，而「这次没有开场记忆」是合法状态。
     */
    const registerHotMemory = (active) => {
        const systemPrompt = typeof ctx.get === 'function' ? ctx.get('systemPrompt') : undefined;
        if (systemPrompt === undefined || typeof systemPrompt.variable !== 'function') {
            ctx.logger?.warn?.('[dsh-office-mode] 当前组合里没有 systemPrompt 服务：'
                + '{{' + HOT_MEMORY_VARIABLE + '}} 未注册 —— 办公 preset 的 persona 会因未知变量装配失败。');
            return;
        }
        disposers.push(systemPrompt.variable(HOT_MEMORY_VARIABLE, (context) => {
            try {
                // assembly 的 agent 是 ReactLoopAgent：工作目录挂在 immutable session.header 上
                // （agent 本身没有 cwd 字段；`agent.cwd` 是 CreateAgentOptions.meta 的写法）。
                const cwd = context?.agent?.session?.header?.cwd;
                if (typeof cwd !== 'string' || cwd === '') return '';
                return readHotMemory(cwd, active?.memory ?? {});
            } catch {
                return '';
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

    const registerGuide = (active) => {
        if (!active.injectGuide) return;
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
    const registerMemoryHint = (active) => {
        if (!active.memory.enabled || !active.memory.promptHint) return;
        if (active.tools.office_memory === false) return;
        const systemPrompt = typeof ctx.get === 'function' ? ctx.get('systemPrompt') : undefined;
        if (systemPrompt === undefined || typeof systemPrompt.section !== 'function') return;
        disposers.push(systemPrompt.section({
            name: 'office-mode-memory',
            order: MEMORY_HINT_SECTION_ORDER,
            text: MEMORY_PROMPT_HINT,
        }));
    };

    /**
     * 按**一份**配置把全部资源装上（不动 `resolved`，也不回收任何东西）。
     *
     * 为什么每一步都收一个 `active` 参数、而不是读闭包里的 `resolved`：同一段代码要能装
     * 两次 —— 候选配置装失败时，回滚得把**上一份**重新装回去（见 `rebuildWith`）。写死读
     * `resolved` 就做不到，那正是第五十五轮审查 S-04 的形态。
     *
     * ## 步骤顺序不是随手排的
     *
     * 顺序即依赖顺序，但还有一条更硬的理由：**降级态里能剩下什么，就等于「失败点之前装到的
     * 那几步」**。所以两样「缺了会让整条链路更糟」的资源排在前面：
     *   1. `能力变量`（`{{office_capabilities}}`）—— 办公 preset 的 persona 引用了它，
     *      缺了会让**每一次**提示装配抛 `unknown prompt variable`，比「半份能力」更坏；
     *   2. `诊断工具`（`office_help`，其余工具面之前单独一步）—— 它是「只剩部分能力」这件事
     *      唯一能被用户看见的通道（`office_help` 的运行期实况会报降级状态）。
     * 第五十九轮复核 P1：这两样原来都排在工具面之后，于是「工具面注册失败」这个最可能的
     * 降级形态恰好把它们一起丢掉了 —— 实测残留为空。改成现在这个顺序之后，只要失败点不是
     * 它们自己，降级态里这两样都在（`test/rebuild.mjs` 两条用例分别钉住这两种形态）。
     *
     * 中途失败时**已经装进 `disposers` 的那些不在这里回收**：调用方要先知道是哪一步炸的，
     * 再决定「装回旧的」还是「进降级态」。
     *
     * `refresh: false` 只在激活期用（那时配置刚解析出来，没必要强制重探）；重建一律重探 ——
     * 设置改动可能换了 python.bin / ffmpeg / 模型目录。
     */
    const installAll = (active, options = {}) => {
        const step = (label, run) => {
            installStep = label;
            return run();
        };
        // 定义只算一次（配额跟着这一份工具面走），再拆成「诊断工具」与「其余」两步装。
        const definitions = step('工具面清单', () => buildTools(active, ctx));
        const diagnostic = definitions.filter((definition) => definition.name === 'office_help');
        const rest = definitions.filter((definition) => definition.name !== 'office_help');
        // `office_help` 可以被设置页单独关掉（那时它不在这份工具面里，也就没有诊断通道 ——
        // 那是用户的选择，不是降级），所以这里只防「同一件工具出现两次」这种插件自身的 bug。
        // 工具全关时 `definitions` 是空数组，两步都只注册零件，照常继续（第五十九轮回归：
        // 这里原来无条件要求「必须有一个 office_help」，把 `preset-guard` 的「工具全关」用例打红了）。
        if (diagnostic.length > 1) {
            throw new Error(`工具面里 office_help 出现了 ${diagnostic.length} 次（插件自身的 bug）`);
        }
        step('能力探测', () => refreshCapabilities(active, { refresh: options.refresh !== false }));
        step('能力变量', () => registerCapabilityMap(active));
        step('开场记忆变量', () => registerHotMemory(active));
        step('诊断工具', () => registerToolDefinitions(diagnostic));
        step('工具面', () => registerToolDefinitions(rest));
        step('执行守卫', () => registerRunErrorGuard());
        step('preset 守卫', () => registerPresetGuard(active));
        step('提示段落', () => registerGuide(active));
        step('记忆说明段落', () => registerMemoryHint(active));
        // 掩码里的工具名要跟着新工具面重算（已在跑的极简会话由这里补齐）。
        step('守卫重算', () => presetGuard?.resync());
        installStep = '';
    };

    /**
     * 设置改动的重建事务（第五十九轮 S-04）。
     *
     * 宿主的工具注册表、提示变量表与投影注册表都按**名字**查重，新旧两份不能并存 —— 所以
     * 「先把新的装上、再拆旧的」在这里做不到，只能「先拆旧的、再装新的」，中间有一段空窗。
     * 既然做不到原子替换，就必须有**可回滚**的重建协议：
     *
     *   1. 拆掉当前那一份（先记住它是哪一份配置装的）；
     *   2. 按候选配置整份装；
     *   3. 装失败 → 把候选那半份收干净，再把**改动前的那份配置**重新装回去。
     *
     * 第 3 步成功 = 设置页改坏了东西也仍有一份能用的旧工具面（原注释里「下次工具调用仍用旧
     * 配置」那句要表达的就是这件事，只是以前没人保证它真的成立）。
     *
     * 第 3 步也失败 = **显式降级态**：不抛错（触发者是宿主的设置写入，抛回去等于把设置系统
     * 弄挂），由 `reportRebuild()` 留一条可诊断的状态，日志与 `office_help` 的运行期实况照实
     * 报。这一支**故意不再 `disposeAll()`**：残留是「回滚失败点之前装到的那几步」，而按
     * `installAll` 的顺序，那几步里的头两样正是能力变量与 `office_help` ——
     * 降级态下**大概率**既装配得起来、也看得见（失败点如果恰好落在这两样自己身上，
     * 那就真的什么都不剩；这一点照实写进 `describeRebuildState()`，不粉饰）。
     *
     * 出口代理**不在**这个事务里：它是独立生命周期的资源（进程级那条只能装不能清），装不上
     * 只该记一条警告，不该让工具面跟着回滚。调用方在事务提交成功之后才 `applyProxy()`。
     *
     * @param {object} next 候选运行期配置（已拆 volatile 引用、已解析）
     */
    const rebuildWith = (next) => {
        const previous = resolved;
        disposeAll();
        resolved = next;
        try {
            installAll(next);
            return { ok: true, restored: true, phase: '', error: null, rollbackPhase: '', rollbackError: null };
        } catch (error) {
            const phase = installStep;
            disposeAll();
            resolved = previous;
            try {
                installAll(previous);
                return { ok: false, restored: true, phase, error, rollbackPhase: '', rollbackError: null };
            } catch (rollbackError) {
                return {
                    ok: false,
                    restored: false,
                    phase,
                    error,
                    rollbackPhase: installStep,
                    rollbackError,
                };
            }
        }
    };

    /**
     * 重读配置 + 走一次重建事务，并把结果记进可诊断状态。
     *
     * 「重读配置」这一步拆 volatile 引用、重新解析 —— 设置页的写入拼进 profile 里本插件那一
     * 行的用户层，宿主按 volatile 就地更新引用，所以每次都要重新拆一层。
     *
     * 失败时还要报一件容易被漏掉的事：**这一次设置里的代理改动没有生效**。事务失败后运行期
     * 停在改动前那份配置上，而 `search.proxy` 已经写进 profile（设置页显示的是新值）——
     * 不点名的话，用户会以为代理换了、实际一直走旧出口（第五十九轮复核 P2）。
     *
     * @returns {{ok: boolean, restored: boolean, phase: string, error: unknown,
     *            rollbackPhase: string, rollbackError: unknown}} 事务结果
     */
    const rebuildFromConfig = () => {
        const next = safeResolveConfig(config);
        const outcome = rebuildWith(next);
        // 事务失败时 `resolved` 已经回滚成改动前那份，所以这里比的是「候选 vs 实际生效」。
        const droppedProxy = outcome.ok === true || (next?.search?.proxy ?? '') === (resolved?.search?.proxy ?? '')
            ? ''
            : (next?.search?.proxy ?? '');
        const state = reportRebuild({ ...outcome, droppedProxy });
        if (state.ok !== true) warnRebuildFailure(state);
        return outcome;
    };

    /**
     * 重建失败要**说出来**（S-04 的原始缺陷就是那个空 `catch`）。
     *
     * 两种失败形态的退路不同，所以文案必须分开：装回旧配置的话还能接着用；回滚也失败的话
     * 只剩部分能力，得让用户去重挂插件或重启。
     *
     * 参数是 `reportRebuild()` 记下来的那一份（**已过凭据打码**的字符串），不是原始异常 ——
     * 打码只做一次，日志与 `office_help` 拿到的是同一份文本。
     */
    const warnRebuildFailure = (state) => {
        const phase = state.phase === '' ? '未知阶段' : state.phase;
        const proxyNote = state.droppedProxy === '' ? '' : '另外，本次设置里的代理改动没有生效。';
        if (state.restored === true) {
            ctx.logger?.warn?.('[dsh-office-mode] 设置重建失败（' + phase + '）：' + state.error
                + '；已把改动前的配置整份装回（工具面与提示段落仍是上一份，本次设置改动等下次重挂再试）。'
                + proxyNote);
            return;
        }
        ctx.logger?.warn?.('[dsh-office-mode] 设置重建失败（' + phase + '）：' + state.error
            + '；装回旧配置也失败（' + (state.rollbackPhase === '' ? '未知阶段' : state.rollbackPhase)
            + '）：' + state.rollbackError
            + ' —— 当前只剩部分能力，请重挂本插件或重启 DSH。' + proxyNote);
    };

    /**
     * 跑一次「设置改动 → 重建 → 装代理」。
     *
     * ## 为什么要一个重入守卫
     *
     * cordis 的 `ctx.emit()` 是**同步派发、而且可重入**的：监听器里再 emit，会嵌套执行
     * （第五十九轮复核实测顺序 `L1-enter / L1-enter / L1-exit / L1-exit`；无界重入会栈溢出）。
     * 宿主 `dsh-settings` 不只在**写入**时广播这个事件，在**读取设置投影**那条链路上也会
     * 广播。所以「事务跑到一半又收到一次广播」是一条真实路径，而重入的后果正是 S-04 要根治
     * 的那种形态：内层 `rebuildWith` 的 `previous` 拿到的是**候选**、内层 `disposeAll()` 会把
     * 外层刚装好的那份拆掉，`resolved` / 注册表 / `disposers` 三者互相不一致，而两边各自都
     * 返回成功 —— 谁也不会报错。
     *
     * 守卫的做法不是丢掉这次广播，而是记一个 pending：事务结束后用**当前**的 config 再跑一遍
     * （config 是就地更新的引用，所以补跑拿到的一定是最新的那一份；重入的两次广播自然合并成
     * 最后一次）。循环不会自转：这段代码自己不再广播。
     */
    const runSettingsRebuild = () => {
        if (rebuilding === true) {
            pendingRebuild = true;
            return;
        }
        rebuilding = true;
        try {
            do {
                pendingRebuild = false;
                // 事务失败时不要再去动出口代理：运行期已经回滚到改动前那份配置，按候选配置切
                // 出口会留下「工具面是旧的、出口是新的」这种没人解释得清的组合（失败这件事由
                // 上面的 warn 与状态照实报）。
                if (rebuildFromConfig().ok === true) safeApplyProxy();
            } while (pendingRebuild === true);
        } finally {
            rebuilding = false;
        }
    };

    /**
     * 按当前设置装一次出口代理（第二十九轮）。
     *
     * 只装不改：装不上（运行时不支持 / 地址不合法）记一条警告就够了 —— 它是
     * 「让联网换条出口」的增益项，不该让插件激活失败，也**不进重建事务**（第五十九轮
     * S-04：代理有自己的生命周期，装不成不该把工具面回滚掉）。装了之后**清不掉**
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

    /**
     * 装一次代理，并把**抛错**也接住。
     *
     * `installProxy` 内部把能想到的失败都变成了 `ok:false` + reason，正常不会抛；这里再兜一层
     * 是因为它挂在两条主链路上（激活、设置事件）：代理是「让联网换条出口」的增益项，
     * **任何**形态的失败都不该让插件起不来，更不该被算成「设置重建失败」。
     *
     * 消息过一遍 `maskProxyUrl()`：代理地址可能带 `user:pass@`，异常文案里回显它等于把口令
     * 写进日志与 `office_help` 的运行期实况（第二十九轮审查 P1-1 的纪律）。
     */
    const safeApplyProxy = () => {
        try {
            applyProxy();
        } catch (error) {
            ctx.logger?.warn?.('[dsh-office-mode] 装出口代理时抛错（不该发生，代理这一项没生效）：'
                + maskProxyUrl(String(error?.message ?? error)));
        }
    };

    // 激活：装一次（失败要**响亮** —— 宿主会记一条插件加载失败），但不把半份注册留在
    // 注册表里：宿主重试加载时同名注册会撞车。
    try {
        installAll(resolved, { refresh: false });
    } catch (error) {
        disposeAll();
        throw error;
    }
    safeApplyProxy();

    // ── 记忆浏览面板的数据端点 ─────────────────────────────────────────────
    //
    // 只注册一次，且**不**进 disposers：rebuild（设置页改动）不该把路由拆了重建。
    // 处理器每次请求都读当前的 resolved.memory，所以配置改了立刻生效。
    //
    // 用 ctx.inject 而不是直接 ctx.get('webServer')：这样不必假设 webServer 一定
    // 在插件行之前就绪（组合顺序变了也不会静默少一条路由）。组合里根本没有
    // webServer（例如 headless）时回调不触发，插件其余功能照常。
    let disposeView = null;
    let disposePresetPlugins = null;
    let disposePreviewView = null;
    ctx.inject(['webServer'], (sctx) => {
        disposeView = registerMemoryView(sctx, {
            getMemory: () => resolved.memory,
            getConfig: () => resolved,
            // 借宿主官方供应商的 Key：用父 ctx 读 llm / loader / credentials，
            // 这些服务在插件根上下文里已经就绪；只在 OCR 端点就是官方时才会真正解析。
            resolveHostOcrCredential: (options) => resolveHostOcrCredential(ctx, options),
        }) ?? null;
        // 预览面板的三条只读路由（清单 / 大纲 / 单页图）。与记忆那条共用同一个
        // 信任边界（trustedRootFor），每次请求都现读 resolved 里的 cacheDir ——
        // 缓存目录改了，下一次渲染就落到新位置，不用重挂插件。
        disposePreviewView = registerPreviewView(sctx, {
            getMemory: () => resolved.memory,
            getConfig: () => resolved,
        }) ?? null;
        // 新-21（第六十三轮）：「办公模式插件」那一格的两条路由 —— 读 preset 里的插件行、
        // 按行回写 `disabled`（真源 + profile 侧已有的覆盖行）、还原上一次备份。
        // 与记忆面板那条同理：**每条路由用自己注入的 ctx、每次请求现算目标文件**，
        // 所以设置改了、profile 换了都不用重挂插件（写入面与控制流见 src/preset-plugins.js 的头部注释）。
        disposePresetPlugins = registerPresetPlugins(sctx) ?? null;
    });

    // ── 会话知识库范围命令 ─────────────────────────────────────────────────
    //
    // `/kb-scope` 是唯一能改会话范围的入口：它跑在宿主的命令服务里，能拿到真实
    // 的会话 header（会话 id、父会话、origin），所以「谁在改自己的范围」不靠
    // 客户端自报。工具面（office_memory）只**读**范围，参数里没有可写的口子。
    //
    // 与面板那条同理：只注册一次、不进 disposers，rebuild 不重挂它（范围与设置无关）。
    let disposeScopeCommand = null;
    ctx.inject(['commands'], (sctx) => {
        disposeScopeCommand = registerKbScopeCommand(sctx, {
            logger: ctx.logger,
            // 核对用户写下的虚拟文件夹是否真在库里。读的是**未受限**的树（选项本来也来自它），
            // 所以这一步只回答「有没有这个目录」，不泄露该目录里有什么。
            folderExists: async (session, folder) => {
                const cwd = typeof session?.header?.cwd === 'string' ? session.header.cwd.trim() : '';
                if (cwd === '') return null;
                const memory = createMemory({ root: cwd, memory: resolved.memory });
                const tree = await memory.kbTree({});
                // kbTree 返回 { action, root, ... }，树在 tree.root 下；helper 也接受裸节点。
                // folders.json 坏掉时 helper 返回 null（树读不全），命令层会回「暂时核对不了」。
                return kbTreeFolderPresence(tree, folder);
            },
        }) ?? null;
    });

    // ── 设置变更 ──────────────────────────────────────────────────────────
    //
    // 0.1.7 起设置页的写入走宿主设置文档：拼进 profile 里本插件那一行的用户层，
    // 由 dsh-settings 落盘并广播 `settings/document-updated`。旧版那套
    // `settings.installSection(...) + onChange` 的契约已经不存在了。
    //
    // volatile 字段是就地更新的引用（宿主不重挂插件），所以这里必须自己听事件、
    // 重新解析配置再重建工具面。监听器**不进 disposers**：`disposeAll()` 清的是
    // 资源 disposers，把监听器塞进去就等于**第一次重建时先把设置链路自己拆掉**
    // （第五十九轮复核 P3：原注释把理由写成「rebuild 会清空它」，会误导下一个改这里的人）。
    // 它的生命周期挂在插件卸载上（见下面的 `dispose`）。
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
                // 重建走 `runSettingsRebuild()`：它包着**事务**（第五十九轮 S-04）与**重入
                // 守卫**（宿主读取设置投影时也会广播这个事件，而 cordis 的 emit 是同步且可
                // 重入的）。装不成会把改动前的那份配置整份装回来；连回滚都失败才进降级态，
                // 并且一定会留下日志与可诊断状态。以前的空 `catch` 把失败静默吃掉，于是只剩
                // 半份能力而没人知道。
                //
                // 代理**在提交成功之后**才装：它不进事务（装不成不该回滚工具面），但也绝不能
                // 在事务失败、配置已经回滚之后再按候选配置去切出口 —— 那会留下「工具面是旧的、
                // 出口是新的」这种没人解释得清的组合。
                try {
                    runSettingsRebuild();
                } catch (error) {
                    // 连事务本身都抛（不该发生：rebuildWith 把每一步都包住了）。同样不静默 ——
                    // 抛回宿主会让「改一个设置」把设置写入弄失败。
                    const state = reportRebuild({ ok: false, restored: false, phase: '重建事务', error, rollbackPhase: '', rollbackError: null });
                    ctx.logger?.warn?.('[dsh-office-mode] 设置重建事务自身抛错（不该发生）：'
                        + state.error + ' —— 当前状态未知，请重挂本插件。');
                }
            });
        });
    } catch {
        disposeSettingsWatch = null;
    }

    if (typeof ctx.on === 'function') {
        ctx.on('dispose', () => {
            disposeAll();
            for (const teardown of [disposeView, disposePreviewView, disposePresetPlugins, disposeScopeCommand, disposeSettingsWatch]) {
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
    } else {
        // 宿主没有事件服务时「卸载钩子」挂不上：资源只能靠宿主回收。这是降级，不是正常路径，
        // 照实记一条（第五十九轮复核 P3：原来这一段是静默的）。
        ctx.logger?.warn?.('[dsh-office-mode] 当前组合里没有事件服务（ctx.on）：卸载时不会主动释放'
            + '工具面与提示段落，交给宿主回收。');
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
export {
    GUARDED_TOOL_NAMES,
    TOOL_HIDDEN_PRESETS,
    enabledOfficeToolNames,
    installPresetGuard,
    shouldHideOfficeTools,
} from './preset-guard.js';
export { registerMemoryView, buildSnapshot, noteWorkspace, knownWorkspaces, resetWorkspaces, registryWorkspacesOf, sessionWorkspacesOf, trustedRootFor, publicWorkspaces, MEMORY_VIEW_PATH, VIEW_LIMITS } from './view.js';
export { registerPreviewView, PREVIEW_VIEW_PATHS, PREVIEW_FILE_LIMIT } from './preview-view.js';
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
    mapDeepSeekApiResponse,
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
export { createOfficeSettings, OFFICE_SETTINGS_NS, LIMITS } from './settings.js';
export {
    PRESET_BODY_LIMIT,
    PRESET_BACKUP_KEEP,
    PRESET_PLUGINS_PATH,
    PRESET_RESTORE_PATH,
    backupDirOf,
    listBackups,
    parsePresetPlugins,
    parseProfileOverrides,
    readPresetPlugins,
    registerPresetPlugins,
    resolvePresetTargets,
    restoreLastBackup,
    setDisabledOnLine,
    writePluginToggle,
} from './preset-plugins.js';