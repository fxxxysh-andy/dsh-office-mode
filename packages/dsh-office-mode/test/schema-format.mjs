/**
 * 工具 schema 里的**格式指令**（第十八轮 §5.3 / `18-23` 的另一半）。
 *
 * `IFEval-FC`（2509.18420）的结论是：基准只评参数正确性，**不测嵌在参数描述里的格式
 * 指令**；而把格式写进 JSON schema 之后，连 SOTA 模型都频繁违反。本项目的 `office_*`
 * schema 里有大量格式约束（`queries` 1-5 条、`depth` 1-3、`preprocess` 三档、
 * `dryRun` 布尔、`since` 的时间词……），这一套给每一条配 5 条**纯正则/结构判定**用例，
 * 并把「违反率」变成一个可复算的数：
 *
 *   1. **词表核对**：描述里列的值必须与代码里的常量一致（两向）—— 描述不许编值，
 *      常量里的值不许漏在描述外。这一条抓的是「描述腐烂」（第四十七轮实测到一个实例：
 *      `office_help` 的话题清单少列了 4 个、还把 `tex` 写了两遍）。
 *   2. **5 条用例 + 判定分层**：每条参数的 5 条用例各声明一种处理 ——
 *      `ok`（合法）/ `reject`（被拒绝）/ `tolerated`（被收敛或当作没传，必须有理由）。
 *   3. **声称拒绝的必须真的拒绝**：`reject` 用例都带探针（跑一次真的校验路径或工具调用），
 *      探针没拒就是**违反**。违反率 = 没拒住的 `reject` 用例 / 全部 `reject` 用例，断言为 0。
 *      `tolerated` 不计入分母，但必须写明「为什么可以容忍」——容忍是要有理由的，
 *      不是「没人管」。
 *   4. **类型层自动覆盖**：每个参数的 schema 类型都自动配一对用例（合法类型 + 错类型），
 *      所以「没进格式登记表」的参数也不是没人管。
 *
 * 跑法：node test/schema-format.mjs
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { availableHelpTopics, buildHelp } from '../src/docs.js';
import { KB_TIERS, strictTierOf } from '../src/kb.js';
import {
    CONFLICT_CLASSES,
    CONFLICT_STATES,
    DEFAULT_ENTRY_LIMIT_BYTES,
    DEFAULT_LAYERS,
    LINK_KINDS,
    MEMORY_ACTIONS,
    MEMORY_IMPORTANCE,
    MEMORY_TARGETS,
    MEMORY_TIME_HINTS,
    MEMORY_TIME_PATTERNS,
    MEMORY_TIME_WORDS,
} from '../src/memory.js';
import { CONTENT_TYPES } from '../src/search-routes.js';
import { SITE_TYPE_IDS } from '../src/site-catalog.js';
import { buildTools } from '../src/tools.js';
import { resolveConfig } from '../src/config.js';
import { validateUrl } from '../src/web.js';
import { PROVIDER_IDS } from '../src/web-providers.js';
import { PREPROCESS_MODES } from '../src/web-preprocess.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const TOOLS = buildTools(resolveConfig({}));
const toolOf = (name) => TOOLS.find((tool) => tool.name === name);
const paramOf = (tool, param) => {
    const schema = toolOf(tool)?.parameters?.properties?.[param];
    assert.ok(schema !== undefined, `工具面里找不到 ${tool}.${param}（登记表过期了？）`);
    return schema;
};
const descOf = (tool, param) => String(paramOf(tool, param).description ?? '');

// ── 判定工具（纯正则 / 结构） ──────────────────────────────────────────────

const inList = (list) => (value) => list.includes(value);
const inRange = ([min, max]) => (value) => Number.isInteger(value) && value >= min && value <= max;
const isUrl = (value) => {
    try {
        validateUrl(value);
        return true;
    } catch {
        return false;
    }
};
const emptyOk = (predicate) => (value) => value === '' || predicate(value);

/** 时间窗：空 = 没传；时间词与相对写法（真值从 memory.js 现取）；ISO 日期（含日期是否真的存在）。 */
const TIME_WORDS = MEMORY_TIME_WORDS;
const TIME_PATTERNS = MEMORY_TIME_PATTERNS.map((source) => new RegExp(source));
const isIsoDate = (value) => {
    const match = /^(\d{4})-(\d{2})-(\d{2})([T ](\d{2}):(\d{2}))?$/.exec(String(value));
    if (match === null) return false;
    const [, year, month, day, , hour = '0', minute = '0'] = match;
    const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute)));
    return date.getUTCFullYear() === Number(year)
        && date.getUTCMonth() === Number(month) - 1
        && date.getUTCDate() === Number(day)
        && Number(hour) <= 23 && Number(minute) <= 59;
};
const isTime = emptyOk((value) => {
    const text = String(value).toLowerCase().replace(/\s+/g, ' ');
    return TIME_WORDS.includes(text) || TIME_PATTERNS.some((pattern) => pattern.test(text)) || isIsoDate(value);
});

/** 站点优先参数的四态：不传（undefined）/ 布尔 / 类型 id / 域名数组。 */
const isSites = (value) => typeof value === 'boolean'
    || SITE_TYPE_IDS.includes(value)
    || (Array.isArray(value) && value.every((item) => typeof item === 'string' && item !== ''));

// ── 探针（跑真的校验路径 / 真的工具调用；用临时工作区，不碰真实记忆） ────────

const TMP = mkdtempSync(join(tmpdir(), 'office-schema-format-'));
const EXEC = { agent: { session: { header: { cwd: TMP } } }, signal: new AbortController().signal };

/** 跑一次工具调用；抛错 = 拒绝。 */
const toolRejects = (name) => async (args) => {
    try {
        await toolOf(name).execute(args, EXEC);
        return false;
    } catch {
        return true;
    }
};
const memoryRejects = toolRejects('office_memory');
const webSearchRejects = toolRejects('office_web_search');
const searchRunRejects = toolRejects('office_search_run');

/** 纯常量层的拒绝（tier 的 strictTierOf 就是这么规定的）。 */
const tierRejects = (value) => {
    try {
        strictTierOf(value);
        return false;
    } catch {
        return true;
    }
};

/**
 * link 的 conflict / state：探针要**认报错的原因**，不能只看「抛没抛」。
 *
 * 这两条的值域校验发生在 id 查找之前，而探针用的是假 id —— 合法值会继续往下走、
 * 最后因为「找不到 id」抛错。只判「抛没抛」的话合法用例也会被算成「被拒绝」，
 * 正向控制（合法用例不许被拒）就永远过不去，违反率也变成恒真。
 */
const linkFieldRejects = (field) => async (value) => {
    try {
        await toolOf('office_memory').execute({ action: 'link', sourceId: 'a', targetId: 'b', [field]: value }, EXEC);
        return false;
    } catch (error) {
        return new RegExp(`${field} 只支持`).test(String(error?.message ?? error));
    }
};

// ── 登记表：描述里带格式约束的参数 ──────────────────────────────────────────
//
// 用例写 [值, 处理]：
//   ok        —— 合法值
//   reject    —— 必须被拒绝（带探针）
//   tolerated —— 被收敛到合法值、或按「没传」处理（必须写 why）
//
// 每条参数 5 条用例（纯判定用的是 valid / enabled 这类谓词）。

const TOPICS = availableHelpTopics();
const LAYERS = [...Object.keys(DEFAULT_LAYERS), 'all'];
/** 来源档是**不区分大小写**的（kb.js 的 strictTierOf 会先 lower 再比）。 */
const CORE_TIERS = (value) => KB_TIERS.includes(String(value).toLowerCase());
const CONTENT_TYPE_IDS = CONTENT_TYPES.map((entry) => entry.id);

const REGISTRY = [
    {
        tool: 'office_help', param: 'topic',
        vocabulary: TOPICS,
        valid: emptyOk(inList(TOPICS)),
        cases: [['word', 'ok'], ['', 'ok'], ['wrod', 'tolerated'], ['word/excel', 'tolerated'], ['没有这个话题', 'tolerated']],
        why: '认不出的话题不报错：退回索引，并在回执第一行列出可用话题（docs.js 的 buildHelp）——一次调用就能自我纠正',
    },
    {
        tool: 'office_memory', param: 'action',
        vocabulary: MEMORY_ACTIONS,
        valid: inList(MEMORY_ACTIONS),
        cases: [['read', 'ok'], ['kb-search', 'ok'], ['', 'tolerated'], ['nope', 'reject'], ['READ', 'reject']],
        reject: (value) => memoryRejects({ action: value }),
        why: '空的 action 按 read 处理（tools.js 的 `args.action.trim() !== \'\' ? … : \'read\'`）——查询是只读的，兜底成它最安全',
    },
    {
        tool: 'office_memory', param: 'content',
        // 长度约束（第四十八轮 新-11）：阈值从真源现取，描述里的默认值也是同一个常量
        // （tools.js 的 content 描述由 DEFAULT_ENTRY_LIMIT_BYTES 拼出来）。
        valid: (value) => typeof value === 'string' && value.trim() !== ''
            && Buffer.byteLength(value, 'utf8') <= DEFAULT_ENTRY_LIMIT_BYTES,
        cases: [
            ['一条短约定。', 'ok'],
            ['x'.repeat(2000), 'ok'],
            ['', 'reject'],
            ['x'.repeat(DEFAULT_ENTRY_LIMIT_BYTES + 1), 'reject'],
            ['x'.repeat(50_000), 'reject'],
        ],
        reject: (value) => memoryRejects({ action: 'add', target: 'project', content: value }),
        why: '空内容与超长内容都必须当场拒绝：前者会写进一条空记忆，后者会把整层容量吃光 —— '
            + '两条都不是「收敛到合法值」能救的（对照：认不出的 importance / layer 走 tolerated）',
    },
    {
        tool: 'office_memory', param: 'target',
        vocabulary: MEMORY_TARGETS,
        valid: inList(MEMORY_TARGETS),
        cases: [['user', 'ok'], ['project', 'ok'], ['weird', 'reject'], ['', 'reject'], ['User', 'reject']],
        reject: (value) => memoryRejects({ action: 'add', target: value, content: '目标值探针。' }),
    },
    {
        tool: 'office_memory', param: 'tier',
        vocabulary: KB_TIERS,
        valid: emptyOk(CORE_TIERS),
        cases: [['user', 'ok'], ['VERIFIED', 'ok'], ['', 'ok'], ['trusted', 'reject'], ['user-ish', 'reject']],
        reject: async (value) => tierRejects(value),
    },
    {
        tool: 'office_memory', param: 'importance',
        vocabulary: MEMORY_IMPORTANCE,
        valid: inList(MEMORY_IMPORTANCE),
        cases: [['critical', 'ok'], ['normal', 'ok'], ['weird', 'tolerated'], ['', 'tolerated'], ['HIGH', 'tolerated']],
        why: '认不出的重要度退回 normal（memory.js 的 normalizeEntry）——它是排序用的标签，拒绝一条真实记忆不划算',
    },
    {
        tool: 'office_memory', param: 'layer',
        vocabulary: LAYERS,
        valid: inList(LAYERS),
        cases: [['hot', 'ok'], ['all', 'ok'], ['weird', 'tolerated'], ['', 'tolerated'], ['HOT', 'tolerated']],
        why: '认不出的层退回 all（memory.js 的 read：`wanted` 只在三层里挑，其余一律 all）——读多一层比读不到更安全',
    },
    {
        tool: 'office_memory', param: 'kind',
        vocabulary: LINK_KINDS,
        valid: inList(LINK_KINDS),
        cases: [['related', 'ok'], ['contradicts', 'ok'], ['weird', 'tolerated'], ['', 'tolerated'], ['RELATED', 'tolerated']],
        why: '认不出的关系类型退回 related（memory.js 的 link）——它只是边的标签；但冲突类的两个字段是例外，见下两条',
    },
    {
        tool: 'office_memory', param: 'conflict',
        vocabulary: CONFLICT_CLASSES,
        valid: emptyOk(inList(CONFLICT_CLASSES)),
        cases: [['context-memory', 'ok'], ['', 'ok'], ['weird', 'reject'], ['context', 'reject'], ['CONTEXT-MEMORY', 'reject']],
        reject: linkFieldRejects('conflict'),
    },
    {
        tool: 'office_memory', param: 'state',
        vocabulary: CONFLICT_STATES,
        valid: emptyOk(inList(CONFLICT_STATES)),
        cases: [['unresolved', 'ok'], ['', 'ok'], ['weird', 'reject'], ['prefer-source', 'ok'], ['PREFER-SOURCE', 'reject']],
        reject: linkFieldRejects('state'),
    },
    {
        tool: 'office_memory', param: 'depth',
        valid: inRange([1, 3]),
        cases: [[1, 'ok'], [3, 'ok'], [9, 'tolerated'], [0, 'tolerated'], [1.5, 'tolerated']],
        why: '越界被 clamp 到 1..3（memory.js 的 related：`Math.min(3, Math.max(1, boundedInt(...)))`）——多走几跳只是多花点力气',
    },
    {
        tool: 'office_memory', param: 'since',
        valid: isTime,
        cases: [['今天', 'ok'], ['2026-09-27', 'ok'], ['', 'ok'], ['瞎写', 'reject'], ['2026-13-45', 'reject']],
        reject: (value) => memoryRejects({ action: 'read', since: value }),
    },
    {
        tool: 'office_memory', param: 'until',
        valid: isTime,
        cases: [['上周', 'ok'], ['明天', 'ok'], ['2026-09-27', 'ok'], ['瞎写', 'reject'], ['2026-13-45', 'reject']],
        reject: (value) => memoryRejects({ action: 'read', until: value }),
    },
    {
        tool: 'office_memory', param: 'format',
        // 描述里写的是「如 word / excel / ppt / tex / pdf」＝**举例**，不是枚举；
        // 判定用举例集合，落地时别的值也能写（见 why）。
        vocabulary: ['word', 'excel', 'ppt', 'tex', 'pdf'],
        valid: emptyOk(inList(['word', 'excel', 'ppt', 'tex', 'pdf'])),
        cases: [['word', 'ok'], ['', 'ok'], ['xlsx', 'tolerated'], ['随便写', 'tolerated'], ['WORD', 'tolerated']],
        why: '台账的 format 是**自由文本**（描述里写的是「如 …」＝举例，不是枚举）：写别的值只是台账那一行显示成别的名字，不影响任何判断（tools.js 的 log 分支直接把它写进台账）',
    },
    {
        tool: 'office_memory', param: 'source',
        valid: (value) => Array.isArray(value) && value.every((item) => typeof item === 'string'
            // `kb:` 前缀的项是**块引用**，形状是 kb:<12 位十六进制>:<块号>（kb.js 的内容寻址 id）。
            && (item.startsWith('kb:') ? /^kb:[0-9a-f]{12}:\d+$/.test(item) : item !== '')),
        // 第三位 `'existence'`：这一条超出纯判定（要查库才知道块在不在），由探针负责 —— 见下面
        // 「判定与声明一致」那条对它的豁免。
        cases: [[['notes.md'], 'ok'], [[], 'ok'], [['kb:zzz'], 'reject'], [['kb:abcdef123456:1'], 'reject', 'existence'], ['notes.md', 'tolerated']],
        reject: async (value) => memoryRejects({ action: 'add', target: 'project', content: '来源形状探针。', source: value }),
        why: '不是数组时按「没传」处理（tools.js 的 `Array.isArray(args.source) ? args.source : null`）——三态里的「没传」与「传空数组」是刻意分开的',
    },
    {
        tool: 'office_web_search', param: 'queries',
        valid: (value) => Array.isArray(value) && value.length >= 1 && value.length <= 5,
        cases: [[['a'], 'ok'], [['a', 'b', 'c', 'd', 'e'], 'ok'], [[], 'reject'], [['a', 'b', 'c', 'd', 'e', 'f'], 'tolerated'], ['一条查询', 'tolerated']],
        reject: (value) => webSearchRejects({ queries: value }),
        why: '超过 5 条被 slice 到 5（tools.js 的 WEB_SEARCH_MAX_QUERIES）；非数组被当成单条查询。schema 声明的类型是 array —— 按 schema 校验输入的宿主会更早拒绝',
    },
    {
        tool: 'office_web_search', param: 'maxResults',
        valid: inRange([1, 20]),
        cases: [[8, 'ok'], [1, 'ok'], [0, 'tolerated'], [21, 'tolerated'], ['abc', 'tolerated']],
        why: '越界与不是数的值被收敛到 1-20 或默认 8（tools.js 的 `Math.max(1, Math.min(20, …))`）——少取几条来源不是格式错误',
    },
    {
        tool: 'office_web_search', param: 'sites',
        valid: isSites,
        cases: [['academic', 'ok'], [false, 'ok'], [['arxiv.org'], 'ok'], ['nope', 'tolerated'], ['academic/book', 'tolerated']],
        why: '认不出的类型 id 选不出任何站点（site-catalog.js 的 selectSiteEntries 只认 SITE_TYPE_IDS），这一轮等于不限定并退回泛搜',
    },
    {
        tool: 'office_web_fetch', param: 'url',
        valid: isUrl,
        cases: [['https://example.com', 'ok'], ['http://a.b/c', 'ok'], ['not a url', 'reject'], ['ftp://x', 'reject'], ['', 'reject']],
        // 探针走**真的校验函数**（web.js 的 validateUrl），不是把 valid 取反 ——
        // 取反的话这条的违反率定义上恒为 0（复核 瑕疵 4）。
        reject: async (value) => {
            try {
                validateUrl(value);
                return false;
            } catch {
                return true;
            }
        },
    },
    {
        tool: 'office_web_fetch', param: 'preprocess',
        vocabulary: PREPROCESS_MODES,
        valid: emptyOk(inList(PREPROCESS_MODES)),
        cases: [['article', 'ok'], ['off', 'ok'], ['' , 'ok'], ['raw', 'tolerated'], ['ARTICLE', 'tolerated']],
        why: '认不出的强度退回默认 article（web-preprocess.js 的 resolvePreprocessOptions）——预处理强度是「省 token」的旋钮，不是安全性边界',
    },
    {
        tool: 'office_search_run', param: 'queries',
        valid: (value) => Array.isArray(value) && value.length >= 1 && value.length <= 5,
        cases: [[['a'], 'ok'], [['a', 'b', 'c', 'd', 'e'], 'ok'], [[], 'reject'], [['a', 'b', 'c', 'd', 'e', 'f'], 'tolerated'], ['一条查询', 'tolerated']],
        reject: (value) => searchRunRejects({ queries: value }),
        why: '与 office_web_search 同一条口径：超过 5 条被 slice 到 5、非数组按单条处理（tools.js 的 executeSearchRun 与 executeWebSearch）',
    },
    {
        tool: 'office_search_run', param: 'maxResults',
        valid: inRange([1, 20]),
        cases: [[8, 'ok'], [20, 'ok'], [0, 'tolerated'], [99, 'tolerated'], ['abc', 'tolerated']],
        why: '越界被收敛到 1-20（tools.js 的 executeSearchRun），不是数字时按设置页的默认值',
    },
    {
        tool: 'office_search_run', param: 'fetchPages',
        valid: inRange([0, 5]),
        cases: [[2, 'ok'], [0, 'ok'], [9, 'tolerated'], [-1, 'tolerated'], ['x', 'tolerated']],
        why: '越界被收敛到 0-5（tools.js 的 executeSearchRun），不是数字时按设置页的默认值',
    },
    {
        tool: 'office_search_run', param: 'provider',
        vocabulary: PROVIDER_IDS,
        valid: emptyOk(inList(PROVIDER_IDS)),
        cases: [['duckduckgo', 'ok'], ['bocha', 'ok'], ['', 'ok'], ['nope', 'tolerated'], ['DuckDuckGo', 'tolerated']],
        why: '认不出的通道名退回设置页的顺序自动挑（web-providers.js 的 resolveProviderOrder 会过滤掉未知名）——通道名拼错不该让整次检索失败',
    },
    {
        tool: 'office_search_run', param: 'preprocess',
        vocabulary: PREPROCESS_MODES,
        valid: emptyOk(inList(PREPROCESS_MODES)),
        cases: [['article', 'ok'], ['plain', 'ok'], ['', 'ok'], ['raw', 'tolerated'], ['PLAIN', 'tolerated']],
        why: '与抓取工具同一条口径：认不出的强度退回默认 article（web-preprocess.js）',
    },
    {
        tool: 'office_search_run', param: 'sites',
        valid: isSites,
        cases: [['code', 'ok'], [true, 'ok'], [['github.com'], 'ok'], ['nope', 'tolerated'], ['code/book', 'tolerated']],
        why: '与 office_web_search 同一条口径：认不出的类型 id 等于不限定（site-catalog.js 只认 SITE_TYPE_IDS）',
    },
    {
        tool: 'office_search_brief', param: 'type',
        vocabulary: CONTENT_TYPE_IDS,
        valid: emptyOk(inList(CONTENT_TYPE_IDS)),
        cases: [['hotspot', 'ok'], ['mixed', 'ok'], ['', 'ok'], ['nope', 'tolerated'], ['Hotspot', 'tolerated']],
        why: '认不出的内容类型退回「按主题文字自动判断」（search-routes.js 的 byId 查不到就不当类型用）——自动判断本来就是这个参数的默认行为',
    },
    {
        tool: 'office_parse_findings', param: 'type',
        vocabulary: CONTENT_TYPE_IDS,
        valid: emptyOk(inList(CONTENT_TYPE_IDS)),
        cases: [['knowledge', 'ok'], ['manual', 'ok'], ['', 'ok'], ['nope', 'tolerated'], ['KNOWLEDGE', 'tolerated']],
        why: '与 office_search_brief 同一条：认不出就自动判断（search-routes.js 的 byId；这里的自动判断用于生成核对清单）',
    },
];

/** 通用忽略词：出现在描述里但不属于词表的常见英文/参数名（登记着，便于审查）。 */
const GENERIC_IGNORE = new Set([
    'id', 'kb', 'key', 'web', 'url', 'http', 'https', 'pdf', 'markdown',
    'true', 'false', 'default', 'file', 'files', 'path', 'paths', 'page', 'pages',
    // 参数名与用法提示（描述里说明「这个字段什么时候用」，不是值域）：
    'action', 'target', 'content', 'tier', 'oldtext', 'importance', 'entities', 'tags',
    'layer', 'query', 'limit', 'since', 'until', 'sourceid', 'targetid', 'kind', 'conflict',
    'state', 'depth', 'packpath', 'from', 'dryrun', 'format', 'theme', 'purpose', 'source',
    'topic', 'type', 'audience', 'files', 'briefpath', 'outputpaths', 'maxparallel',
    'queries', 'out', 'maxresults', 'fetchpages', 'provider', 'preprocess', 'sites',
    'script', 'keepcache', 'detail', 'path', 'run', 'mode', 'text', 'log',
]);

/** 全部词表的并集：某个参数描述里出现别的参数的词表值（例如 tier 描述里的 kb-ingest）不算编词。 */
const ALL_VOCABULARY = new Set(REGISTRY.flatMap((entry) => entry.vocabulary ?? []));

/** 从一段描述里提取「像值」的英文词（去掉本参数词表的值之后剩下的部分）。 */
function strayTokens(description, vocabulary) {
    let rest = String(description);
    for (const value of [...vocabulary].sort((left, right) => right.length - left.length)) {
        rest = rest.split(value).join('\u0000');
    }
    const tokens = [...rest.matchAll(/[A-Za-z][A-Za-z0-9_-]{1,}/g)].map((match) => match[0]);
    return [...new Set(tokens)].filter((token) => !GENERIC_IGNORE.has(token.toLowerCase())
        && !ALL_VOCABULARY.has(token)
        && !ALL_VOCABULARY.has(token.toLowerCase()));
}

// ── 1. 登记表与工具面一致 ──────────────────────────────────────────────────

await check('格式登记表里的参数都还在工具面里，且每条正好 5 条用例', () => {
    for (const entry of REGISTRY) {
        const schema = paramOf(entry.tool, entry.param);
        assert.ok(schema !== undefined, `${entry.tool}.${entry.param} 不在工具面里`);
        assert.equal(entry.cases.length, 5, `${entry.tool}.${entry.param} 应有 5 条用例，实际 ${entry.cases.length}`);
        assert.ok(typeof entry.valid === 'function', `${entry.tool}.${entry.param} 缺纯判定 valid()`);
        const kinds = entry.cases.map(([, kind]) => kind);
        assert.ok(kinds.includes('ok'), `${entry.tool}.${entry.param} 至少要有一条合法用例`);
        assert.ok(kinds.some((kind) => kind === 'reject' || kind === 'tolerated'),
            `${entry.tool}.${entry.param} 至少要有一条「会出问题」的用例（否则这份登记没有意义）`);
        for (const kind of kinds) assert.ok(['ok', 'reject', 'tolerated'].includes(kind), `未知处理：${kind}`);
        if (kinds.includes('reject')) assert.ok(typeof entry.reject === 'function', `${entry.tool}.${entry.param} 声称拒绝但没有探针`);
        if (kinds.includes('tolerated')) assert.ok(typeof entry.why === 'string' && entry.why !== '', `${entry.tool}.${entry.param} 有容忍用例，必须写清为什么可以容忍`);
    }
    return `${REGISTRY.length} 条参数 / ${REGISTRY.length * 5} 条用例`;
});

// ── 2. 词表核对（描述 ↔ 代码常量，两向） ────────────────────────────────────

await check('词表两向核对：常量里的值都在描述里，描述里没有编出来的值', () => {
    const report = [];
    for (const entry of REGISTRY) {
        if (!Array.isArray(entry.vocabulary) || entry.vocabulary.length === 0) continue;
        const description = descOf(entry.tool, entry.param);
        const missing = entry.vocabulary.filter((value) => !description.includes(value));
        assert.equal(missing.length, 0,
            `${entry.tool}.${entry.param} 的描述里少了这些值：${missing.join('、')}（描述腐烂的典型：新值加进了常量没进描述）`);
        const stray = strayTokens(description, entry.vocabulary);
        assert.equal(stray.length, 0,
            `${entry.tool}.${entry.param} 的描述里出现了常量里没有的值：${stray.join('、')}`
            + '（描述不许编值；是通用词就加进 test/schema-format.mjs 的 GENERIC_IGNORE 并写清理由）');
        report.push(`${entry.tool}.${entry.param}(${entry.vocabulary.length})`);
    }
    return report.join(' ');
});

await check('office_help 的话题清单自同步：工具描述 = 真源 = 未知话题回执', async () => {
    const schema = paramOf('office_help', 'topic');
    for (const topic of TOPICS) {
        assert.ok(schema.description.includes(topic), `工具描述的话题清单少了 ${topic}`);
        // 正向：列出来的每一个都必须真的能查（列一个查不到的比少列更糟）。
        const help = await buildHelp(topic);
        assert.equal(help.topic, topic, `话题「${topic}」列在清单里，但 buildHelp 回的不是它`);
    }
    const reply = await buildHelp('__不存在的話題__');
    const listed = /可用：([^\n]+)。/.exec(reply.text)?.[1].split(/\s*\/\s*/) ?? [];
    assert.deepEqual(listed, TOPICS, '「没有这个话题」的回执必须与真源逐字一致（原来这里重复列了 tex、又漏了四个话题）');
    assert.equal(new Set(listed).size, listed.length, '话题清单里不该有重复项');
    return `${TOPICS.length} 个话题`;
});

await check('两份索引都列全了话题，且每个话题只出现一次', async () => {
    // 索引文本（默认层 COMPACT_INDEX / 全文层的「其他话题」）是**另外两处**手写清单，
    // 不在 availableHelpTopics 的同源核对里（复核 F12）：新话题加进真源却没进索引，
    // 模型就不知道它存在。第四十七轮顺手修掉了 `pdf` / `tex` 在全文层各出现两次。
    for (const detail of [false, true]) {
        const text = (await buildHelp('', { detail })).text;
        const lines = text.split('\n');
        const missing = [];
        const duplicated = [];
        for (const topic of TOPICS) {
            const hits = lines.filter((line) => line.startsWith(`- ${topic}：`)).length;
            if (hits === 0) missing.push(topic);
            if (hits > 1) duplicated.push(`${topic}×${hits}`);
        }
        assert.equal(missing.length, 0, `${detail ? '全文层' : '默认层'}索引里少了这些话题：${missing.join('、')}`);
        assert.equal(duplicated.length, 0, `${detail ? '全文层' : '默认层'}索引里重复列了：${duplicated.join('、')}`);
    }
    return `${TOPICS.length} 个话题 × 两份索引`;
});

await check('时间窗提示自同步：报错里列的每个字都真的能解析', async () => {
    let message = null;
    try {
        await toolOf('office_memory').execute({ action: 'read', since: '__不存在的時間__' }, EXEC);
    } catch (error) {
        message = String(error.message);
    }
    assert.ok(message !== null, '不可能的时间必须报错 —— 可用词表就在那句话里');
    for (const word of MEMORY_TIME_HINTS) {
        assert.ok(message.includes(word), `报错提示里漏了能解析的时间词「${word}」`);
        // 反向：提示里列的每个字都要真的能解析（跑真工具；空库也会给回执）。
        await toolOf('office_memory').execute({ action: 'read', since: word }, EXEC);
    }
    return `${MEMORY_TIME_HINTS.length} 个提示词逐个验过`;
});

// ── 3. 用例的判定与处理一致 ────────────────────────────────────────────────

await check('5 条用例的判定与声明一致（合法=ok，其余=reject/tolerated）', () => {
    let exempt = 0;
    for (const entry of REGISTRY) {
        for (const [value, kind, exemption] of entry.cases) {
            const valid = entry.valid(value);
            if (kind === 'ok') {
                assert.ok(valid, `${entry.tool}.${entry.param}：${JSON.stringify(value)} 声明 ok，但纯判定说它不合法`);
                continue;
            }
            if (exemption === 'existence') {
                // 纯判定看不到「库里的块在不在」这类事实，那一条由 reject 探针负责。
                exempt += 1;
                continue;
            }
            assert.ok(!valid, `${entry.tool}.${entry.param}：${JSON.stringify(value)} 声明 ${kind}，但纯判定说它合法`);
        }
    }
    return `${exempt} 条由探针负责（超出纯判定）`;
});

// ── 4. 声称拒绝的必须真的拒绝（违反率） ────────────────────────────────────

await check('违反率 = 0：每一条 reject 用例都真的被拒绝', async () => {
    let denominator = 0;
    const report = [];
    for (const entry of REGISTRY) {
        for (const [value, kind] of entry.cases) {
            if (kind !== 'reject') continue;
            denominator += 1;
            const rejected = await entry.reject(value);
            if (!rejected) report.push(`${entry.tool}.${entry.param}=${JSON.stringify(value)}`);
        }
    }
    assert.equal(report.length, 0, `这些值声称会被拒绝、实际放行了（违反率 > 0）：${report.join('、')}`);
    return `违反率 0/${denominator}`;
});

await check('正向控制：合法用例的探针不许「一律拒绝」（否则违反率是恒真的）', async () => {
    // 没有这一条，把某个工具的写入路径整条打断（`add`/`link` 永远抛错）也能全绿：
    // 违反率只看 reject 用例，探针恒抛就永远「0 违反」。所以每条带探针的参数都要拿
    // **一条合法用例**过一遍探针，要求它**不**被拒绝。
    //
    // 两个例外照实登记（它们的合法用例要真的走网络，不拿网络当判据）：
    const NETWORK_BOUND = new Set(['office_web_search.queries', 'office_search_run.queries']);
    let controlled = 0;
    const skipped = [];
    for (const entry of REGISTRY) {
        if (typeof entry.reject !== 'function') continue;
        const key = `${entry.tool}.${entry.param}`;
        if (NETWORK_BOUND.has(key)) {
            skipped.push(key);
            continue;
        }
        const okCase = entry.cases.find(([, kind]) => kind === 'ok');
        assert.ok(okCase !== undefined, `${key} 缺合法用例`);
        const rejected = await entry.reject(okCase[0]);
        assert.equal(rejected, false,
            `${key}：合法值 ${JSON.stringify(okCase[0])} 也被拒了 —— 探针恒拒的话「违反率 0」是恒真的，这条守门就是为了不让它空转`);
        controlled += 1;
    }
    return `${controlled} 条参数有正向控制${skipped.length > 0 ? `；${skipped.join('、')} 例外（要联网）` : ''}`;
});

await check('容忍清单是显式的：tolerated 用例都带理由，且理由指到代码出处', () => {
    const tolerated = REGISTRY.flatMap((entry) => entry.cases
        .filter(([, kind]) => kind === 'tolerated')
        .map(([value]) => ({ entry, value })));
    assert.ok(tolerated.length > 0, '容忍用例不该为零（全拒或全放都说明这套登记没在看现实）');
    for (const item of tolerated) {
        // 容忍必须能指到「谁在收敛它」——写成一句感觉良好但找不到出处的话不算。
        assert.match(item.entry.why, /\.js/, `${item.entry.tool}.${item.entry.param} 的容忍理由要指到代码出处：${item.entry.why}`);
    }
    return `${tolerated.length} 条容忍用例`;
});

// ── 5. 类型层：每个参数自动配一对用例 ──────────────────────────────────────

/** JSON Schema 的子集判定（type / enum / minimum / maximum / minItems / maxItems / items.type / pattern）。 */
function schemaAccepts(schema, value) {
    if (schema === undefined) return false;
    const types = Array.isArray(schema.type) ? schema.type : (schema.type === undefined ? [] : [schema.type]);
    if (types.length > 0) {
        const matched = types.some((type) => {
            if (type === 'string') return typeof value === 'string';
            if (type === 'integer') return Number.isInteger(value);
            if (type === 'number') return typeof value === 'number' && Number.isFinite(value);
            if (type === 'boolean') return typeof value === 'boolean';
            if (type === 'array') return Array.isArray(value);
            if (type === 'object') return value !== null && typeof value === 'object' && !Array.isArray(value);
            return true;
        });
        if (!matched) return false;
    }
    if (Array.isArray(schema.enum) && !schema.enum.includes(value)) return false;
    if (typeof value === 'number') {
        if (Number.isFinite(schema.minimum) && value < schema.minimum) return false;
        if (Number.isFinite(schema.maximum) && value > schema.maximum) return false;
    }
    if (Array.isArray(value)) {
        if (Number.isFinite(schema.minItems) && value.length < schema.minItems) return false;
        if (Number.isFinite(schema.maxItems) && value.length > schema.maxItems) return false;
    }
    if (typeof value === 'string' && typeof schema.pattern === 'string' && !(new RegExp(schema.pattern)).test(value)) return false;
    return true;
}

/** 每个类型的一个合法样本；错类型样本从候选里挑第一个**不属于任何声明类型**的。 */
const TYPE_SAMPLES = {
    string: '示例',
    integer: 2,
    number: 2.5,
    boolean: true,
    array: [],
    object: {},
};

/** 一个值是否属于 schema 声明的类型之一（联合类型要逐个看）。 */
function matchesDeclaredType(types, value) {
    return types.some((type) => {
        if (type === 'string') return typeof value === 'string';
        if (type === 'integer') return Number.isInteger(value);
        if (type === 'number') return typeof value === 'number' && Number.isFinite(value);
        if (type === 'boolean') return typeof value === 'boolean';
        if (type === 'array') return Array.isArray(value);
        if (type === 'object') return value !== null && typeof value === 'object' && !Array.isArray(value);
        return true;
    });
}

/** 找一个真的不属于声明类型的值（联合类型如 `['boolean','string','array']` 不能拿字符串当错样本）。 */
function wrongTypeSample(types) {
    const candidates = [42, '例', true, [], {}];
    const found = candidates.find((value) => !matchesDeclaredType(types, value));
    assert.ok(found !== undefined, `找不到与类型 ${JSON.stringify(types)} 不匹配的样本`);
    return found;
}

await check('类型层自动覆盖：每个参数都配「合法类型 + 错类型」一对用例', () => {
    let covered = 0;
    for (const tool of TOOLS) {
        for (const [name, schema] of Object.entries(tool.parameters?.properties ?? {})) {
            const types = Array.isArray(schema.type) ? schema.type : [schema.type];
            const usable = types.filter((type) => TYPE_SAMPLES[type] !== undefined);
            assert.ok(usable.length > 0, `${tool.name}.${name} 的 schema 没有声明可判定的类型`);
            const ok = TYPE_SAMPLES[usable[0]];
            assert.ok(schemaAccepts(schema, ok),
                `${tool.name}.${name}：schema 连自己的合法类型样本都不接受（${JSON.stringify(ok)}）`);
            const bad = wrongTypeSample(types);
            assert.ok(!schemaAccepts(schema, bad),
                `${tool.name}.${name}：schema 接受了错类型的值（${JSON.stringify(bad)}）`);
            covered += 1;
        }
    }
    return `${covered} 个参数 × 2 条`;
});

// ── 输出 ──────────────────────────────────────────────────────────────────

rmSync(TMP, { recursive: true, force: true });

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok && item.note !== undefined ? `  → ${item.note}` : ''}`);
    if (!item.ok) console.log(`       | ${item.error?.message ?? item.error}`);
}
console.log(`schema-format: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);