/**
 * 极简模式（DSH 内置 preset `minimal`）里不注入 office_* 工具。
 *
 * ## 为什么不能靠 preset 里的禁用行
 *
 * 宿主 `dsh-tools` 的工具注册表是**按作用域分层**的：一个 agent 可见的工具 =
 * 全局层 + 它作用域链上每一层（近的覆盖远的），限制掩码只对**继承来的**那一部分
 * 生效。本插件是 profile 层的一行，`ctx.tools.register()` 落在**全局层** ——
 * 每个 preset 的 agent 都继承得到。
 *
 * 于是 preset 里写 `- id: dsh-office-mode / disabled: true` **关不掉它**：那一行
 * 只让该 preset 自己那一份不注册（本 preset 本来也没注册），全局层那份照旧继承。
 * 本机实测（第五十轮）：ptc preset 里 `tool-workflow` / `workflow-ptc` /
 * `tool-ralph` / `tool-plugin-manager` 都写着 `disabled: true`，而 profile 层把
 * 这四行打开之后，ptc 会话的实际工具面里 `workflow` / `ralph` / `plugin_manager`
 * 一个不少（`cordis_inspect_query` 的 Tool/listTools 41 项可查）。
 *
 * ## 所以走 restrict
 *
 * 唯一能「按 agent 隐藏继承来的工具」的机制是 `ctx.tools.restrict({ deny })`：
 * 它作用在**某个 agent 自己的作用域**上，对继承来的工具应用拒绝掩码，而且
 * 「作用域注册保持可见」（agent 自己那一层注册的东西不受掩码影响）。掩码同时
 * 作用于 schema 展示、名字解析与执行，所以隐藏掉的不只是提示词成本。
 *
 * 触发点在 `agent/created`：那时 preset 已经绑好（`agentPresets.composedPreset()`
 * 能读出 preset id），子代理也走同一条路（它们继承父会话的 preset 修订）。
 *
 * @module dsh-office-mode/preset-guard
 */
import { DEFAULT_TOOLS } from './config.js';

/**
 * 在这些 preset 里不出现本插件的工具。
 *
 * 只列 `minimal`（DSH 内置的「极简模式」）：它的定位就是「一个 shell + 一段人格」，
 * 九个 office_* 工具与它无关。办公模式（`office`）与别的模式照旧 —— 开发时在
 * ptc 会话里直接调 office_run 的那条路子不能被这层守卫误伤。
 */
export const TOOL_HIDDEN_PRESETS = ['minimal'];

/**
 * 办公 preset 里「声明要关、运行期却没关掉」的那批**宿主**工具（第五十四轮 新-15①）。
 *
 * 清单是拿 `presets/office/cordis.patch.yml` 的 `disabled: true` 行**逐行对出来的**
 * （行 id → 工具名），所以预设文件改了这里也要跟着改 —— `test/preset-check.mjs`
 * 有一条断言核这两处不脱节。分组与理由：
 *
 *   bash / pwsh            命令行。办公模式的产物由 office_run 的引擎写，不需要 shell
 *   job_*                  后台任务：办公会话是一次做完一批，不铺长期任务
 *   workflow / ralph       编排与迭代循环，属编程工作流
 *   subagent*              模型自己起子代理 / 管子代理的**工具**入口。**注意**：检索
 *                          三步走的派工不受影响 —— 它用的是宿主 `subagents` **服务**
 *                          （ctx.get('subagents') + start()），不是这几个工具。
 *                          Agent Teams 的 spawn_teammate / team_task_* 也**不在**清单里：
 *                          预设没有声明要关它们（协作那一线在办公模式里是保留的）
 *   interrupt_agent / send_message / list_agents  子代理控制面
 *   todo_write / skill / create_goal / get_goal / update_goal  规划与技能（预设有对应禁用行）
 *   plugin_manager         装插件（编程向）
 *   task_board_*           任务看板：第三十二轮起用户要求「办公模式下关闭」
 *
 * **名字写错不会报错吗**：`tools.restrict()` 对不认识的全局工具名是**当场抛错**。
 * 所以先按公开 API `tools.get(name)`（全局视图）筛一遍「本机真的注册在全局层」的名字，
 * 再逐个装（见 `installMask`）；本机没有的名字**跳过并报出来**——不同机器装的能力不同
 * （没有 workflow / 没有看板都是正常的），但「静默什么都没藏」是不允许出现的失败形态。
 */
export const OFFICE_DEV_TOOLS = [
    'bash', 'pwsh',
    'job_list', 'job_output', 'job_kill',
    'workflow', 'ralph',
    'subagent', 'subagent_fork',
    'interrupt_agent', 'send_message', 'list_agents',
    'todo_write', 'skill',
    'create_goal', 'get_goal', 'update_goal',
    'plugin_manager',
    'task_board_list', 'task_board_get', 'task_board_create', 'task_board_update',
    'task_board_set_parent', 'task_board_run', 'task_board_manage', 'task_board_schedule',
];

/** 本插件可能注册的全部工具名（= 设置页工具开关的键集，顺序与工具面一致）。 */
export const GUARDED_TOOL_NAMES = Object.keys(DEFAULT_TOOLS);

/**
 * 当前**真的注册了**的工具名。
 *
 * `restrict()` 会校验每个名字是不是已知的全局工具，写一个没注册的名字会当场抛错；
 * 设置页关掉的工具不注册，所以掩码必须按当前配置现算，不能拿常量表。
 *
 * @param {{tools?: Record<string, boolean>}} [resolved] 运行期配置
 * @returns {string[]} 要隐藏的工具名（可能为空 = 不必装掩码）
 */
export function enabledOfficeToolNames(resolved) {
    const tools = resolved !== null && typeof resolved === 'object' ? resolved.tools : undefined;
    if (tools === null || typeof tools !== 'object') return [...GUARDED_TOOL_NAMES];
    return GUARDED_TOOL_NAMES.filter((name) => tools[name] !== false);
}

/**
 * 这个 preset 要不要隐藏本插件的工具。
 *
 * 读不到 preset（没有 `agentPresets` 服务的部署、preset-free 的组合、读的时候抛错）
 * 一律**不隐藏**：那是「单会话进程级组合」的老部署，本插件的工具本该照旧可用。
 *
 * @param {string|undefined} presetId preset 标识
 * @param {string[]} [hidden] 隐藏清单
 */
export function shouldHideOfficeTools(presetId, hidden = TOOL_HIDDEN_PRESETS) {
    return typeof presetId === 'string' && presetId !== '' && hidden.includes(presetId);
}

/**
 * 装上「极简模式隐藏 office 工具」这层守卫。
 *
 * 三件事：
 *   1. `agent/created` —— 新 agent 绑好 preset 之后判一次；
 *   2. `agent-preset/selected` —— 空白会话换 preset（宿主只允许在开首轮之前换）时重判，
 *      从极简模式换到别的模式要把掩码**撤掉**，否则办公工具会被永久挡住；
 *   3. `agent/disposed` —— 清掉表里的引用（掩码本身是 agent 作用域的资源，随作用域回收）。
 *
 * 返回 `{ dispose, resync }`：`dispose` 交给调用方的 disposers（设置改动重建时会被调），
 * `resync` 在**工具面重建之后**再调一次 —— 掩码里写的是注册期现算的名字，设置页开关
 * 变了要重新算（否则新开的工具会漏进已在跑的极简会话）。
 *
 * 第五十四轮（`新-15①`）起它还管**宿主工具**：`hostHidden` 给出一批 preset id，凡是在
 * 这些 preset 里的 agent，额外把 `hostToolNames()` 那批宿主工具掩掉。这一路默认关闭
 * （`guard.hideDevToolsInOffice`，用户选择「做成设置项、默认不隐藏」）。
 *
 * @param {object} ctx 插件上下文
 * @param {object} [options]
 * @param {() => string[]} [options.toolNames] 现算要隐藏的本插件工具名
 * @param {string[]} [options.hidden] 隐藏本插件工具的 preset 清单（默认 `TOOL_HIDDEN_PRESETS`）
 * @param {() => string[]} [options.hostToolNames] 现算要隐藏的宿主工具名（默认 `OFFICE_DEV_TOOLS`）
 * @param {string[]} [options.hostHidden] 隐藏宿主工具的 preset 清单（默认空 = 这一路关着）
 * @param {object} [options.logger] 记警告用（默认取 `ctx.logger`）
 */
export function installPresetGuard(ctx, options = {}) {
    const hidden = Array.isArray(options.hidden) ? options.hidden : TOOL_HIDDEN_PRESETS;
    const hostHidden = Array.isArray(options.hostHidden) ? options.hostHidden : [];
    const toolNames = typeof options.toolNames === 'function' ? options.toolNames : () => [...GUARDED_TOOL_NAMES];
    const hostToolNames = typeof options.hostToolNames === 'function' ? options.hostToolNames : () => [...OFFICE_DEV_TOOLS];
    const logger = options.logger ?? ctx?.logger;
    const listeners = [];
    /** agent → 撤销掩码的函数。用 Map（不是 WeakMap）：释放时要能遍历。 */
    const applied = new Map();
    let disposed = false;

    if ((hidden.length === 0 && hostHidden.length === 0) || typeof ctx?.on !== 'function') {
        return { dispose: () => {}, resync: () => {} };
    }

    const service = (name) => {
        if (typeof ctx.get !== 'function') return undefined;
        try {
            return ctx.get(name);
        } catch {
            // 组合里没有这个服务（或代理读未 inject 的属性抛错）：按「读不到」处理。
            return undefined;
        }
    };

    /** agent 的 preset id；读不到一律 undefined（= 不隐藏）。 */
    const presetOf = (agent) => {
        const presets = service('agentPresets');
        if (presets === undefined || typeof presets.composedPreset !== 'function') return undefined;
        try {
            return presets.composedPreset(agent.ctx);
        } catch {
            return undefined;
        }
    };

    /** agent 作用域上的 tools 代理（**必须**经它调 restrict，作用域才落在该 agent 上）。 */
    const toolsOf = (agent) => {
        try {
            return agent?.ctx?.tools;
        } catch {
            return undefined;
        }
    };

    const warn = (message) => {
        try {
            logger?.warn?.(message);
        } catch {
            // 记日志失败不该影响守卫本身。
        }
    };

    const release = (agent) => {
        const entry = applied.get(agent);
        if (entry === undefined) return;
        applied.delete(agent);
        try {
            entry.dispose();
        } catch {
            // 撤销失败（作用域已经在回收）不该往外冒。
        }
    };

    /** 两份名字清单是不是同一份（顺序无关，掩码语义与顺序无关）。 */
    const sameNames = (left, right) => left.length === right.length
        && left.every((name) => right.includes(name));

    const describe = (error) => (error instanceof Error ? error.message : String(error));

    /**
     * 装掩码。**正向判据在前、错误文案在后**（第五十四轮复核 P2-3 的修法）。
     *
     * 正向判据用公开 API `tools.get(name)`（不传 scope = 全局视图）先筛出**真的注册在
     * 全局层**的名字：
     *   - `restrict()` 只对**继承来的**工具生效，agent 自己那一层注册的名字它掩不掉
     *     （宿主的 `view()` 把 own 层的注册无条件放进 visible），全局视图正好排掉这一类；
     *   - 筛完再 restrict，就只剩「真错误」会抛 —— 不再依赖第三方错误文案做判断。
     * 只有拿不到 `get`（假 ctx / 精简实现 / 老宿主）时才退回按错误文案分类。
     *
     * 两路的装法仍不同：
     *   - **本插件的工具**一次装完：名字全是本插件自己注册的，出现「未知工具」只可能是
     *     机制坏了（或设置页关掉了没现算），要整条响亮地报出来；
     *   - **宿主工具**逐个装：它们是别的插件注册进全局层的，不同机器装的东西不同
     *     （没有 workflow、没有看板都正常），所以只跳过「本机没有 / 掩不掉」这一种，
     *     其余失败照旧上报 —— 把真错误也当「这台机器没有」，就成了静默的假成功。
     */
    const installMask = (tools, officeNames, hostNames) => {
        const disposers = [];
        const failed = [];
        const skipped = [];
        // 拿不到 get 时按「都存在」处理：那时错误文案是唯一可用的判据（见上面注释）。
        const inGlobalLayer = typeof tools.get === 'function'
            ? (name) => {
                try {
                    return tools.get(name) !== undefined;
                } catch {
                    return true;
                }
            }
            : () => true;
        // 本插件自己的工具缺席 = 真问题（它本来就该在全局层）；宿主工具缺席 = 这台机器没装。
        const present = (list, absent, note) => list.filter((name) => {
            if (inGlobalLayer(name)) return true;
            absent.push(note === undefined ? name : `${name}${note}`);
            return false;
        });
        const officePresent = present(officeNames, failed, '（本插件的工具不在全局层：restrict 掩不掉它）');
        const hostPresent = present(hostNames, skipped);
        if (officePresent.length > 0) {
            try {
                const dispose = tools.restrict({ deny: officePresent });
                // 宿主契约是返回 disposer；万一给了别的（假 ctx / 精简实现），
                // 仍**记一个空操作**占位 —— 不记就等于下次 sync 再装一层，
                // 而且 dispose 时无从撤销（第五十四轮复核 P3-5）。
                disposers.push(typeof dispose === 'function' ? dispose : () => {});
            } catch (error) {
                failed.push(describe(error));
            }
        }
        for (const name of hostPresent) {
            try {
                const dispose = tools.restrict({ deny: [name] });
                disposers.push(typeof dispose === 'function' ? dispose : () => {});
            } catch (error) {
                const message = describe(error);
                if (/unknown global tool/i.test(message)) skipped.push(name);
                else failed.push(`${name}（${message}）`);
            }
        }
        return {
            failed,
            skipped,
            applied: disposers.length,
            dispose: () => {
                for (const dispose of disposers) {
                    try {
                        dispose();
                    } catch {
                        // 作用域可能已经在回收：撤销失败不该往外冒。
                    }
                }
            },
        };
    };

    /** 「跳过的名字」这类告警每个**不同的清单**只说一次（复核 P3-3：否则每个新 agent 都刷一条）。 */
    const skippedWarned = new Set();

    /**
     * 让一个 agent 的掩码与「它的 preset + 当前工具面」一致。
     *
     * 掩码里带着**名字清单**：工具面变了（设置页开关）而清单没变就不用重装，
     * 变了要撤掉旧的再装新的 —— 只判「装过没有」会让新开的工具漏进已在跑的会话。
     * 上一次**有跳过的名字**时不缓存：本机后来装上那个工具（或宿主换了注册方式）时，
     * 下一次 sync 要能把它补上（复核 P3-2）。
     */
    const sync = (agent) => {
        if (disposed) return;
        if (agent === null || typeof agent !== 'object' || agent.ctx === undefined) return;
        const preset = presetOf(agent);
        const officeNames = shouldHideOfficeTools(preset, hidden) ? toolNames() : [];
        const hostNames = shouldHideOfficeTools(preset, hostHidden) ? hostToolNames() : [];
        // 非数组防御要放在展开**之前**（复核 P3-4）：两路里任何一路返回非数组时，
        // 展开会先抛 TypeError，而抛出点在 agent/created 监听器里 —— 宿主的 announce 是
        // await 的，监听器抛错会让「建 agent」这件事本身失败。
        for (const list of [officeNames, hostNames]) {
            if (!Array.isArray(list)) {
                warn('[dsh-office-mode] 守卫的名字清单不是数组（工具清单由回调现算，这里是插件自身的 bug），本次不装掩码。');
                return;
            }
        }
        const names = [...new Set([...officeNames, ...hostNames])];
        // 工具全关（或这一路的清单为空）时没有可掩的东西：`restrict({})` 还会被判成 no-op 抛错。
        if (names.length === 0) {
            release(agent);
            return;
        }
        const current = applied.get(agent);
        if (current !== undefined && sameNames(current.names, names) && current.skipped.length === 0) return;
        if (current !== undefined) release(agent);
        const tools = toolsOf(agent);
        if (tools === undefined || typeof tools.restrict !== 'function') return;
        const mask = installMask(tools, officeNames, hostNames);
        if (mask.applied === 0) {
            // 一个都没装上：两种成因分开说（复核 P3-1 —— 旧文案把「本机一个都没有」也说成「装不上」）。
            const why = mask.skipped.length > 0 && mask.failed.length === 0
                ? `清单里 ${mask.skipped.length} 个名字本机都不存在：${mask.skipped.join('、')}`
                : `装不上的名字：${mask.failed.join('；')}${mask.skipped.length > 0 ? `；另 ${mask.skipped.length} 个本机不存在` : ''}`;
            warn(`[dsh-office-mode] preset「${preset}」下${officeNames.length > 0 ? '隐藏 office 工具失败' : '隐藏工具失败'}：${why}`);
            return;
        }
        // 缓存里要留**跳过的名字**：清单没变但上次有跳过的那些要重试（复核 P3-2），
        // 所以不能只记「装过的清单」。
        applied.set(agent, { dispose: mask.dispose, names: [...names], skipped: [...mask.skipped] });
        if (mask.skipped.length > 0) {
            // 如实报「这台机器上没有这些工具」：跳过是正常的（profile 装的东西不同），
            // 但必须留下痕迹，否则「以为藏了十个、其实一个都没藏」会没人发现。
            // 每个**不同的清单**只说一次（复核 P3-3）：本部署有 8 个幽灵名字（看板 bundle
            // 没进 profile 的 bundles），每个新 agent 都刷一条会把真信号盖掉。
            const key = mask.skipped.slice().sort().join(',');
            if (!skippedWarned.has(key)) {
                skippedWarned.add(key);
                warn(`[dsh-office-mode] preset「${preset}」下这些工具本机不存在，已跳过（同一清单只报一次）：`
                    + `${mask.skipped.join('、')}`);
            }
        }
        if (mask.failed.length > 0) {
            warn(`[dsh-office-mode] preset「${preset}」下部分工具隐藏失败（其余照旧生效）：${mask.failed.join('；')}`);
        }
    };

    const agents = () => {
        const registry = service('agents');
        return registry !== undefined && typeof registry.list === 'function' ? registry : undefined;
    };

    const on = (event, listener) => {
        try {
            const dispose = ctx.on(event, listener);
            // 假 ctx / 精简实现可能不返回 disposer；没有就只记事件名，释放时跳过。
            if (typeof dispose === 'function') listeners.push(dispose);
        } catch (error) {
            warn(`[dsh-office-mode] 订阅 ${event} 失败：${error instanceof Error ? error.message : String(error)}`);
        }
    };

    on('agent/created', (payload) => sync(payload?.agent));
    on('agent/disposed', (payload) => release(payload?.agent));
    // 空白会话换 preset：宿主只允许在开首轮之前换，换完这里重判一次（可能是撤掉掩码）。
    on('agent-preset/selected', (sessionId) => {
        const registry = agents();
        const agent = registry !== undefined && typeof registry.get === 'function' ? registry.get(sessionId) : undefined;
        if (agent !== undefined) sync(agent);
    });

    return {
        /** 重新判一遍所有活着的 agent（工具面重建之后调）。 */
        resync: () => {
            if (disposed) return;
            const registry = agents();
            if (registry === undefined) {
                // 没有 agents 服务就遍历不到活着的 agent —— 上一次装的掩码**留在了它们身上**
                // 而这一轮改了清单也补不上。这种组合少见（单会话进程级部署），但不能静默。
                warn('[dsh-office-mode] 重新计算守卫时读不到 agents 服务：已在跑的会话不会跟着更新，'
                    + '新开的 agent 照旧按 agent/created 判。');
                return;
            }
            let live = [];
            try {
                live = registry.list() ?? [];
            } catch {
                return;
            }
            for (const agent of live) sync(agent);
        },
        dispose: () => {
            if (disposed) return;
            disposed = true;
            for (const dispose of listeners.splice(0)) {
                try {
                    dispose();
                } catch {
                    // 同上：释放失败不该影响别的资源回收。
                }
            }
            for (const agent of [...applied.keys()]) release(agent);
        },
    };
}
