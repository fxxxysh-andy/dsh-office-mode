/**
 * 最近一次「设置改动 → 重建」的结果（第五十九轮，S-04 设置重建原子性）。
 *
 * ## 为什么需要单开一个模块
 *
 * 重建**不是原子的**：宿主的工具注册表与提示变量表都按**名字**查重，新旧两份不能并存，
 * 所以只能「先拆旧的、再装新的」，中间有一段空窗。空窗不能靠原子替换补上，只能靠
 * 「装回旧配置」补 —— 那就是 `src/index.js` 的 `rebuildWith()` 事务。
 *
 * 而回滚本身也可能失败（注册失败往往是环境性的，重试同一批步骤会再失败一次）。那时：
 *   - 异常**不能往外抛**：这条链路的触发者是宿主的设置写入，抛回去等于「改一个设置把
 *     设置系统弄挂」；
 *   - 更不能静默吃掉 —— 静默吞掉正是 S-04 的原始缺陷（旧资源已经拆了、配置已经更新了、
 *     空 `catch` 什么也没留下，于是只剩半份能力而没人知道）。
 *
 * 所以留一份**可诊断的**状态：日志照实报，`office_help` 的运行期实况也读它。正常态
 * （`ok: true`）时 `describeRebuildState()` 返回空串，插件在提示词与工具反馈里**一个字节
 * 都不多花**；只有真的失败过才多一行。
 *
 * ## 口径
 *
 * 与 `web-proxy.js` 的 `active` 一样是**模块级单例**：一个进程里本插件只有一行，多份实例
 * （测试里反复 `apply()`）共享这一份。测试用 `resetRebuildStateForTest()` 归零。
 *
 * 与 web-proxy 的区别要说清：那边是「当前生效的代理」，这边是「最近一次重建的结果」，
 * 后者不参与任何运行期判断（工具面怎么装由事务自己决定），只给人看。
 *
 * @module dsh-office-mode/rebuild-state
 */
import { maskProxyUrl } from './web-proxy.js';

/**
 * 正常态：插件激活时就是这一份（激活成功即「还没发生过重建」）。
 *
 * 字段口径：
 *   - `ok`            最近一次重建是否整份装成；
 *   - `restored`      失败时旧配置有没有装回来（`ok:false` 才有意义）；
 *   - `degraded`      `ok:false` 且 `restored:false` = 只剩部分能力的降级态；
 *   - `phase`         `installAll()` 里失败的那一步（工具面 / 能力变量 / …）；
 *   - `error`         失败原因（只留 message，不留栈；**已过一遍凭据打码**）；
 *   - `rollbackPhase` / `rollbackError`  回滚那一次失败在哪一步、为什么；
 *   - `droppedProxy`  失败时被丢掉的那次代理改动（打码后的地址，空串 = 没有这一档）。
 */
const HEALTHY = Object.freeze({
    ok: true,
    restored: true,
    degraded: false,
    phase: '',
    error: '',
    rollbackPhase: '',
    rollbackError: '',
    droppedProxy: '',
});

let state = HEALTHY;

/**
 * 只留人能读的那一句，并**过一遍凭据打码**。
 *
 * 为什么打码：这条文本会进 `ctx.logger.warn` 与 `office_help`（会话日志），而注册失败的
 * message 来源会随步骤增加而变（第五十九轮复核 P2）。打码用的是与代理 reason 同一把尺子
 * （`web-proxy.js` 的 `maskProxyUrl`，认 `//user:pass@` 那一段），代价是几个字节，
 * 收益是「同一份纪律不因为换了一条链路就失效」。`undefined` / 对象都不该让诊断本身抛错。
 */
function describeError(value) {
    if (value === undefined || value === null) return '';
    if (value instanceof Error) return maskProxyUrl(value.message);
    if (typeof value === 'string') return maskProxyUrl(value);
    return maskProxyUrl(String(value?.message ?? value));
}

/**
 * 记一次重建结果。
 *
 * @param {{ok?: boolean, restored?: boolean, phase?: string, error?: unknown,
 *          rollbackPhase?: string, rollbackError?: unknown, droppedProxy?: string}} outcome
 * @returns {typeof HEALTHY} 记下来的那一份（副本，调用方改不动内部状态）
 */
export function reportRebuild(outcome) {
    state = Object.freeze({
        ok: outcome?.ok === true,
        restored: outcome?.restored === true,
        degraded: outcome?.ok !== true && outcome?.restored !== true,
        phase: describeError(outcome?.phase),
        error: describeError(outcome?.error),
        rollbackPhase: describeError(outcome?.rollbackPhase),
        rollbackError: describeError(outcome?.rollbackError),
        droppedProxy: maskProxyUrl(outcome?.droppedProxy ?? ''),
    });
    return { ...state };
}

/** 最近一次重建的结果（副本）。 */
export function rebuildState() {
    return { ...state };
}

/**
 * 测试用：归零。
 *
 * 与 `web-proxy.js` 的 `resetProxyStateForTest()` 同一口径 —— 上面那份 `state` 是**进程级
 * 单例**（一个进程里本插件只有一行），同一进程里跑多份实例（或一套测试里多个用例）时，
 * `office_help` 读到的是**最后一个**报告者的结果。用例之间要么显式归零，要么自己造出
 * 想要的失败态，别指望上一个用例留下的状态。
 */
export function resetRebuildStateForTest() {
    state = HEALTHY;
}

/**
 * 把最近一次重建的结果写成**给人看的一句话**。
 *
 * 正常态返回空串（调用方据此零成本跳过）；失败态分两种说清，因为退路完全不同：
 *   - 装回了旧配置：工具面还是上一份能用的，改完设置再试一次就行；
 *   - 回滚也失败（降级态）：只剩部分能力，要重挂插件或重启 DSH。
 *
 * 降级态里还有一句**条件句**照实写：残留是「回滚失败点之前装到的那几步」，而能力变量与
 * `office_help` 排在 `installAll()` 最前面（见那里的注释）。所以失败点若落在这两样自己身上
 * （`phase` 或 `rollbackPhase` 是它们），用户手上可能**既没有诊断工具、也缺能力变量** ——
 * 那时这一行只有日志看得到。不粉饰成「诊断通道一定在」。
 *
 * 不写「我不知道」这种话：`phase` / `error` 缺了就照实说缺了。
 *
 * @param {typeof HEALTHY} [current] 要描述的状态，默认取模块内的那一份
 */
export function describeRebuildState(current = state) {
    if (current.ok === true) return '';
    const phase = current.phase === '' ? '未知阶段' : current.phase;
    const reason = current.error === '' ? '未提供原因' : current.error;
    // 代理改动没生效这件事单独点名：设置页显示的是新值，运行期走的却是旧出口（复核 P2）。
    const proxy = current.droppedProxy === ''
        ? ''
        : `本次设置里的代理改动（${current.droppedProxy}）没有生效：运行期仍走改动前的出口，`
            + '下一次成功的重建或重挂插件后生效。';
    const head = `设置重建：最近一次失败（${phase}：${reason}）`;
    if (current.restored === true) {
        return `${head}；已装回改动前的配置（工具面与提示段落仍是上一份）。改完设置可再试一次，或重挂本插件。${proxy}`;
    }
    const rollbackPhase = current.rollbackPhase === '' ? '未知阶段' : current.rollbackPhase;
    const rollbackReason = current.rollbackError === '' ? '未提供原因' : current.rollbackError;
    const blind = current.phase === '能力变量' || current.phase === '诊断工具'
        || current.rollbackPhase === '能力变量' || current.rollbackPhase === '诊断工具'
        ? '（失败点落在诊断工具或能力变量自己身上：这两样可能也没装上，请以日志为准。）'
        : '';
    return `${head}；**装回旧配置也失败**（${rollbackPhase}：${rollbackReason}）—— 当前只剩部分能力：`
        + `请重挂本插件或重启 DSH，先别在不完整的工具面上继续写产物。${blind}${proxy}`;
}
