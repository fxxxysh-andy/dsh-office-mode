/**
 * 工具定义：七个（可用设置页逐个关掉）。
 *
 * office_help        —— 按需取文档（开始时只声明工具，写法和细节在需要时才取）。
 * office_run         —— 一次调用完成一批办公操作，结果自动复检。
 * office_memory      —— 三层记忆的读写与检索。
 * office_search_run  —— 内置检索直查（不需要子代理、也不需要提纲）。
 * office_search_brief / office_search_dispatch / office_parse_findings
 *                    —— 按渠道的检索三步（出提纲 → 执行 → 解析回摘要）。
 *
 * 把「少调用、批量做」做成工具形状而不是提示词要求：模型没有"一条一条来"
 * 的入口，想改十处就只能写一次脚本。
 *
 * @module dsh-office-mode/tools
 */
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, join } from 'node:path';
import { buildHelp } from './docs.js';
import { parseFindings, renderFindings } from './findings.js';
import {
    createMemory,
    renderEntities,
    renderExport,
    renderImport,
    renderLink,
    renderLog,
    renderMutation,
    renderRead,
    renderRelated,
    renderStatus,
    renderUnlink,
} from './memory.js';
import { migrateMnemon, renderMigration } from './migrate.js';
import { hintOnce, projectDigest, sessionIdOf } from './projection.js';
import { createTurnQuota } from './quota.js';
import { executeRun } from './run.js';
import { dispatchSearch, excerptOf, parseResultFiles, renderDispatch, writeBrief } from './search.js';
import { buildBrief, contentTypeIds, guessContentType } from './search-routes.js';
import { noteWorkspace } from './view.js';
import {
    createWebAccess,
    describeWebFailure,
    engineDescription,
    PROVIDER_IDS,
    summarizeWebFailures,
    WEB_FAILURE_KINDS,
    webFailureKind,
} from './web.js';

function textBlock(text) {
    return [{ type: 'text', text }];
}

function formatBytes(bytes) {
    const n = Number(bytes) || 0;
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

function compactJson(value) {
    if (value === null || value === undefined) return '（无）';
    try {
        const text = JSON.stringify(value);
        // 上限从 600 提到 4000：office.pdf.pages() 这类返回值要带一页一条的路径清单，
        // 600 会把后半截页面直接吃掉（模型拿到的是截断 JSON，只能靠猜）。
        return text.length > RETURN_ECHO_CHARS ? `${text.slice(0, RETURN_ECHO_CHARS)}…` : text;
    } catch {
        return String(value);
    }
}

/** 脚本返回值回显的字符上限（4000 提到这里是因为中文下 4000 字符 ≈ 12 KB）。 */
const RETURN_ECHO_CHARS = 1200;

/** 单条 outline 最多回显几项（剩下的只报条数）。 */
const OUTLINE_ECHO_MAX = 8;

/**
 * 把「脚本返回」里**已经能从别处读到**的内容去掉再回显。
 *
 * 依据（第十七轮实测）：一次 office_run 里 `returned.outline` 有 41 条 / 6.3 KB，
 * 其中 37 条能在**同一步的脚本参数**里逐字找到 —— 那是模型上一秒自己写进去的正文
 * 被原样回灌；同一屏里「结构：」行又把目录讲了一遍。所以：
 *
 *   - `outline` 从回显里摘掉（结构行已经给出可读目录，先读文件再返回 outline 的
 *     用法本来就该把 outline 留空）；
 *   - 其余内容按字符上限截断（原来是 4000 字符，中文下约 12 KB）。
 *
 * 只动**回显**，不动工具返回值本身 —— 落盘、台账与其它字段一个字节都不受影响。
 */
function shrinkReturned(value) {
    if (value === null || typeof value !== 'object' || Array.isArray(value)) return value;
    if (!Array.isArray(value.outline)) return value;
    const { outline, ...rest } = value;
    return { ...rest, outline: `…共 ${outline.length} 项，见「结构：」行` };
}

function renderStats(stats) {
    if (stats === null || typeof stats !== 'object') return '';
    const parts = [];
    for (const [key, value] of Object.entries(stats)) {
        if (value === null || value === undefined) continue;
        if (Array.isArray(value)) {
            if (value.length === 0) continue;
            parts.push(`${key} ${value.length} 项`);
            continue;
        }
        if (typeof value === 'object') {
            const inner = Object.entries(value)
                .filter(([, v]) => typeof v !== 'object' || v === null)
                .map(([k, v]) => `${k} ${String(v)}`)
                .join(' ');
            if (inner !== '') parts.push(`${key}(${inner})`);
            continue;
        }
        parts.push(`${key} ${String(value)}`);
    }
    return parts.join('｜');
}

/** 把执行结果渲染成模型直接可读的中文反馈。 */
export function renderRun(value) {
    const lines = [];
    if (value.ok === true) {
        lines.push(`✅ office_run 完成（${value.elapsedMs} ms）${value.purpose ? `：${value.purpose}` : ''}`);
    } else {
        const error = value.error ?? {};
        lines.push(`❌ office_run 执行失败：${error.name ?? 'Error'}：${error.message ?? '未知错误'}`);
        // line / column 可能是 null（找不到位置时），别渲染成「脚本第 null 行」。
        if (typeof error.line === 'number') {
            lines.push(`   脚本第 ${error.line} 行${typeof error.column === 'number' ? ` 第 ${error.column} 列` : ''}`);
        }
        if (Array.isArray(error.stack) && error.stack.length > 0) {
            lines.push(...error.stack.slice(0, 3).map((frame) => `   ${frame}`));
        }
        if (value.files.length > 0 || value.otherFiles.length > 0) {
            lines.push('   （失败前已经写出的文件仍然保留，见下）');
        }
    }

    if (value.files.length > 0) {
        lines.push('', `文件（${value.files.length}）`);
        value.files.forEach((file, index) => {
            const badge = [file.format, file.theme].filter(Boolean).join(' · ');
            lines.push(`${index + 1}) ${file.path}${badge ? `  [${badge}]` : ''}  ${formatBytes(file.bytes)}`);
            const stats = renderStats(file.stats);
            if (stats !== '') lines.push(`   ${stats}`);
            if (Array.isArray(file.outline) && file.outline.length > 0) {
                const head = file.outline.slice(0, OUTLINE_ECHO_MAX);
                lines.push(`   结构：${head.join(' / ')}${file.outline.length > head.length ? ` …共 ${file.outline.length} 项` : ''}`);
            }
            if (file.note !== null && file.note !== undefined) lines.push(`   复检：${file.note}`);
            for (const warning of file.warnings ?? []) lines.push(`   ⚠ ${warning}`);
        });
    }

    if (value.otherFiles.length > 0) {
        lines.push(
            '',
            `其他写入（${value.otherFiles.length}）：${value.otherFiles
                .slice(0, 12)
                .map((file) => `${file.path} ${formatBytes(file.bytes)}`)
                .join('、')}${value.otherFiles.length > 12 ? ' …' : ''}`,
        );
    }

    if (value.returned !== null && value.returned !== undefined) {
        lines.push('', `脚本返回：${compactJson(shrinkReturned(value.returned))}`);
    }

    if (Array.isArray(value.logs) && value.logs.length > 0) {
        lines.push('', `日志：${value.logs.slice(0, 12).join('；')}${value.logs.length > 12 ? ' …' : ''}`);
    }
    if (Array.isArray(value.notes) && value.notes.length > 0) {
        lines.push(`备注：${value.notes.slice(0, 12).join('；')}`);
    }

    // 记忆：台账登记结果 + 动笔前就该看的投影（热记忆与最近台账）。
    // 放在文件清单之后、注意事项之前：先说完这次产出了什么，再说「记住的东西」。
    const memory = value.memory ?? {};
    if (memory.enabled === true) {
        if (typeof memory.logged === 'number' && memory.logged > 0) {
            lines.push('', `记忆：台账登记 ${memory.logged} 条（现共 ${memory.kept} 条）${memory.rolled > 0 ? `，${memory.rolled} 条最旧的已滚进归档` : ''}。`);
        } else if (memory.autoLedger === false && (value.files ?? []).length > 0) {
            lines.push('', '记忆：自动登记台账已关闭（要登记用 office_memory({ action: "log", path, format })，或在设置页打开）。');
        }
        if (typeof memory.digest === 'string' && memory.digest !== '') lines.push('', memory.digest);
        if (typeof memory.error === 'string' && memory.error !== '') lines.push(`记忆：台账写入失败（不影响文件）：${memory.error}`);
    }

    const warnings = (value.warnings ?? []).filter((line) => !value.files.some((file) => (file.warnings ?? []).some((w) => line === `${file.path}：${w}`)));
    if (warnings.length > 0) {
        lines.push('', '需要注意：');
        for (const warning of warnings) lines.push(`- ${warning}`);
    }

    const cache = value.cache ?? {};
    if (cache.clearedAfter === true) {
        lines.push('', `缓存：${cache.dir} 已清空（keepCache:false）`);
    } else {
        // 缓存块压成一行（第十七轮）：原来那行 TTL 说明每次调用都一模一样，
        // 而 kept / pruned / hits 这些数字本来就进缓存前缀、每变一次都会让
        // 后面所有内容重算 —— 值不动的一行比三行更省，也更好读。
        const size = formatBytes(cache.keptBytes ?? 0);
        const reused = cache.hits > 0 ? `，命中复用 ${cache.hits} 项（${formatBytes(cache.hitBytes ?? 0)}）` : '';
        const pruned = Array.isArray(cache.pruned) && cache.pruned.length > 0
            ? `，清掉 ${cache.pruned.length} 个过期文件（${formatBytes(cache.prunedBytes ?? 0)}）`
            : '';
        const ttl = cache.ttlMinutes > 0 ? `，${cache.ttlMinutes} 分钟没碰过下次开始时清掉` : '';
        lines.push('', `缓存：${cache.dir} 保留 ${cache.kept ?? 0} 个中间文件（${size}）${reused}${pruned}${ttl}；立即清空用 office.cache.clear()`);
    }
    return lines.join('\n');
}

const briefParameters = {
    type: 'object',
    properties: {
        topic: {
            type: 'string',
            description: '要检索的主题，一句话写清楚（如「某地化工厂爆炸」）。',
        },
        type: {
            type: 'string',
            description: '内容类型：hotspot（热点事件）/ knowledge（知识类）/ manual（手册类）/ mixed（判不准）。省略时按主题文字自动判断。',
        },
        audience: {
            type: 'string',
            description: '选填：这批资料用在哪（如「写一份 12 页的汇报 PPT」），决定详略取舍。',
        },
        files: {
            type: 'array',
            description: '选填：接收结果的相对路径，按渠道分组，第 i 个对应提纲里第 i 个渠道。',
            items: { type: 'string' },
        },
    },
    required: ['topic'],
};

const dispatchParameters = {
    type: 'object',
    properties: {
        briefPath: {
            type: 'string',
            description: 'office_search_brief 写下的提纲文件路径。',
        },
        outputPaths: {
            type: 'array',
            description: '每个渠道一个结果文件路径；子代理会把找到的材料写进这些文件。顺序与提纲里的渠道一致。',
            items: { type: 'string' },
        },
        maxParallel: {
            type: 'integer',
            description: '同时铺开的子代理数量上限，默认 4。渠道多时不要一次全开。',
        },
    },
    required: ['briefPath', 'outputPaths'],
};

const parseParameters = {
    type: 'object',
    properties: {
        paths: {
            type: 'array',
            description: '子代理写下的结果文件路径（Markdown）。可以一次传多个，按渠道合并解析。',
            items: { type: 'string' },
        },
        type: {
            type: 'string',
            description: '内容类型：hotspot / knowledge / manual / mixed；省略时自动判断。',
        },
        topic: {
            type: 'string',
            description: '选填：主题，用于按类型生成核对清单。',
        },
    },
    required: ['paths'],
};

const briefOutputSchema = {
    type: 'object',
    properties: {
        ok: { type: 'boolean' },
        topic: { type: 'string' },
        type: { type: 'string' },
        typeName: { type: 'string' },
        briefPath: { type: 'string' },
        text: { type: 'string' },
    },
    additionalProperties: true,
};

const dispatchOutputSchema = {
    type: 'object',
    properties: {
        ok: { type: 'boolean' },
        jobs: { type: 'array', items: { type: 'object', additionalProperties: true } },
        text: { type: 'string' },
    },
    additionalProperties: true,
};

const parseOutputSchema = {
    type: 'object',
    properties: {
        ok: { type: 'boolean' },
        files: { type: 'array', items: { type: 'string' } },
        text: { type: 'string' },
    },
    additionalProperties: true,
};

const searchRunParameters = {
    type: 'object',
    properties: {
        queries: {
            type: 'array',
            description: '要查的 1-5 条查询，每条一句话（如「2026 年地方政府债务化解进展」）。',
            items: { type: 'string' },
        },
        out: {
            type: 'string',
            description: '结果文件路径（相对会话工作目录）。省略时按第一条查询落在 .office/search/<slug>/run.md。',
        },
        maxResults: {
            type: 'integer',
            description: '每条查询最多取几条来源（1-20，默认 8）。',
        },
        fetchPages: {
            type: 'integer',
            description: '最多打开几页取正文摘录（0-5，默认 2）。取正文更实但更慢；只想要来源清单就填 0。',
        },
        provider: {
            type: 'string',
            description: '选一条检索通道（不传 = 按设置页的顺序自动挑）。可选：'
                + PROVIDER_IDS.map((id) => (id === 'bocha' ? 'bocha（中文好）' : (id === 'duckduckgo' ? 'duckduckgo（免 Key）' : id))).join(' / ')
                + '。seam = 宿主 web 服务，anthropic / openai = 兼容端点的原生检索。',
        },
        preprocess: {
            type: 'string',
            description: '取正文的预处理强度：article（默认，去导航与样板再进上下文）/ plain（只做行级清洗）/ off（原样返回）。',
        },
    },
    required: ['queries'],
};

const searchRunOutputSchema = {
    type: 'object',
    properties: {
        ok: { type: 'boolean' },
        engine: { type: 'string' },
        file: { type: 'string' },
        sources: { type: 'integer' },
        queries: { type: 'array', items: { type: 'object', additionalProperties: true } },
        failures: { type: 'array', items: { type: 'object', additionalProperties: true } },
        text: { type: 'string' },
    },
    additionalProperties: true,
};

const memoryParameters = {
    type: 'object',
    properties: {
        action: {
            type: 'string',
            description: "read（读记忆）/ add（记一条）/ replace（改一条）/ remove（删一条）/ log（手工把一份交付物登记进台账）"
                + " / link（给两条记忆建关系）/ unlink（删关系）/ related（沿关系找相邻条目）/ entities（看实体）/ status（记忆状态汇总）"
                + " / export（整包导出）/ import（整包导入）/ migrate（把 mnemon 的记忆迁进这套记忆）。",
        },
        target: {
            type: 'string',
            description: 'add / replace / remove 用：user = 用户是谁、偏好与要求；project = 项目与环境的事实、约定、坑。默认 project。',
        },
        content: {
            type: 'string',
            description: 'add / replace 用：要记的内容。一两句、自包含、下次不看上下文也读得懂；不要写检索原文或一次性进度。',
        },
        oldText: {
            type: 'string',
            description: 'replace / remove 用：能唯一命中那一条的原文片段。命中多条会被拒绝，给长一点。',
        },
        importance: {
            type: 'string',
            description: 'critical / normal / low（默认 normal）。critical 用于明确的「必须 / 永远不要」；low 用于只对当前一段时间有用的事实。',
        },
        entities: {
            type: 'array',
            description: 'add / replace 用（选填）：这条记忆涉及的实体名（系统、人名、项目名）。填了才会出现在 action:"entities" 的实体视图里——实体是声明的，不猜。',
            items: { type: 'string' },
        },
        tags: {
            type: 'array',
            description: 'add / replace 用（选填）：自由标签，参与归档检索。',
            items: { type: 'string' },
        },
        layer: {
            type: 'string',
            description: 'read 用：hot（偏好与约定）/ ledger（交付台账）/ archive（归档）/ all（默认，三层共用一份预算）。',
        },
        query: {
            type: 'string',
            description: 'read 用：关键词，用来检索台账与归档（多个词用空格分开）。带 query 的读取会走召回质量分档，并计入每回合配额。',
        },
        limit: {
            type: 'integer',
            description: 'read 用：最多返回几条。有硬上限，调大也不会把整本台账灌进上下文。',
        },
        id: {
            type: 'string',
            description: 'related / unlink 用：related 给起点条目的 id；unlink 给要删掉的关系 id。id 从 read 或 related 的结果里拿。',
        },
        sourceId: { type: 'string', description: 'link 用：起点条目的 id。' },
        targetId: { type: 'string', description: 'link 用：终点条目的 id。两个 id 都必须真实存在，否则报错。' },
        kind: {
            type: 'string',
            description: 'link 用：关系类型 related（默认）/ refines / supersedes / contradicts / supports / derives。related 用：只看某一类关系。',
        },
        depth: {
            type: 'integer',
            description: 'related 用：走几跳（1-3，默认 1）。图一旦放开很容易把整个记忆库拉进上下文，所以默认只走一跳。',
        },
        packPath: {
            type: 'string',
            description: 'export / import 用：Pack 文件路径。export 不传时写到记忆目录下的 pack-<时间戳>.json；import 必传。',
        },
        path: {
            type: 'string',
            description: 'log 用：交付物路径（相对工作目录）。office_run 写出的文件已自动登记，这里只补它没覆盖的产物。',
        },
        from: {
            type: 'string',
            description: 'migrate 用：mnemon 的数据根，默认 .mnemon（相对会话工作目录）。',
        },
        dryRun: {
            type: 'boolean',
            description: 'migrate 用：true 时只统计不落盘（先看一眼会迁什么）。',
        },
        format: { type: 'string', description: 'log 用：格式，如 word / excel / ppt / tex / pdf。' },
        theme: { type: 'string', description: 'log 用：主题 id。' },
        purpose: { type: 'string', description: 'log 用：这份产物是做什么的。' },
        source: {
            type: 'array',
            description: 'log 用：材料来源（文件路径或 URL），便于以后追溯这份产物的依据。',
            items: { type: 'string' },
        },
    },
    required: ['action'],
};

const memoryOutputSchema = {
    type: 'object',
    properties: {
        ok: { type: 'boolean' },
        action: { type: 'string' },
        text: { type: 'string' },
    },
    additionalProperties: true,
};

const helpParameters = {
    type: 'object',
    properties: {
        topic: {
            type: 'string',
            description: '要查的话题：word / excel / ppt / tex / pdf / python / theme / files / cache / memory / search / settings / run / guide。省略则返回全部话题索引。',
        },
        detail: {
            type: 'boolean',
            description: '默认只给最小事实集（接口签名 + 必需边界），够写出对的调用。要看例子、字段全表、为什么这样做与配置明细再传 true —— 整篇文档会留在后续每个请求里，能不问就不问。',
        },
    },
};

const runParameters = {
    type: 'object',
    properties: {
        script: {
            type: 'string',
            description: '要执行的 JavaScript 脚本。可以用 office.word / office.excel / office.ppt 造文档，用 office.tex 生成 LaTeX 论文源码并编译 PDF，用 office.av 把录音 / 视频转成文字并给视频抽帧，用 office.files 批量改文件，用 office.python 做科学计算与绘图，最后 return 一个值。语法细节先 office_help。',
        },
        purpose: {
            type: 'string',
            description: '一句话说明这次要做什么（选填，只用于显示）。',
        },
        keepCache: {
            type: 'boolean',
            description: 'false 时这次调用结束就清空缓存目录（默认保留，保留才能跨调用复用：例如渲染好的 PDF 页面图下一步还要 read_image）。',
        },
    },
    required: ['script'],
};

const helpOutputSchema = {
    type: 'object',
    properties: {
        ok: { type: 'boolean' },
        topic: { type: 'string' },
        text: { type: 'string' },
    },
    additionalProperties: true,
};

// 宿主按 output.schema 逐项校验，**声明过的字段缺键即判非法**（2026-09-21 实测：
// 成功路径不写 error 会报 "value.error" must be an object；不传 purpose 会报
// "value.purpose" must be a string，整次调用被判成「返回了非法输出」，脚本跑成功也拿不到反馈）。
// purpose 与 error 本来就是可选的，因此不在这里声明类型：additionalProperties: true
// 已经允许它们出现，宿主也就不会因为缺键而拒绝整次调用。
const runOutputSchema = {
    type: 'object',
    properties: {
        ok: { type: 'boolean' },
        files: { type: 'array', items: { type: 'object', additionalProperties: true } },
        otherFiles: { type: 'array', items: { type: 'object', additionalProperties: true } },
        returned: {},
        logs: { type: 'array', items: { type: 'string' } },
        notes: { type: 'array', items: { type: 'string' } },
        warnings: { type: 'array', items: { type: 'string' } },
        cache: { type: 'object', additionalProperties: true },
        formats: { type: 'array', items: { type: 'string' } },
        elapsedMs: { type: 'integer' },
    },
    additionalProperties: true,
};

/**
 * office_search_brief：出提纲并落盘。
 *
 * 提纲写成文件而不是只在返回值里给一段文本，是因为紧接着派工与若干轮 * 补检索都要用它当规格；落在盘上，子代理和后续调用都能拿到同一份。
 */
async function executeBrief(rawArgs, exec, config) {
    const root = exec?.agent?.session?.header?.cwd ?? process.cwd();
    const args = rawArgs !== null && typeof rawArgs === 'object' ? rawArgs : {};
    const typeId = typeof args.type === 'string' && args.type.trim() !== '' ? args.type : undefined;
    const { topic, type, text } = buildBrief(args.topic, typeId, { audience: args.audience });
    const outputDir = config?.search?.outputDir;
    const briefPath = await writeBrief(root, topic, type.id, args.audience, text, outputDir);
    const shown = (await import('node:path')).relative(root, briefPath).split('\\').join('/');

    // 结果文件路径：调用方没给就按渠道数生成默认路径，省掉一次来回。
    const { defaultOutputPaths } = await import('./search.js');
    const suggested = defaultOutputPaths(topic, type.channels.length, outputDir);

    const feedback = [
        text,
        '',
        '提纲已写入：' + shown,
        '',
        '下一步（两种都行）',
        '- 交给检索执行：office_search_dispatch({ briefPath: "' + shown + '", outputPaths: [...] })',
        '  建议的结果文件（按渠道顺序）：' + suggested.join('、'),
        '  组合里没有 web_search / 子代理时，这一步会自动改用插件的内置检索通道。',
        '- 只想要提纲：直接把上面的渠道清单当人工检索清单用。',
        '',
        '渠道数：' + type.channels.length + '（★ 为必须覆盖）。派工时 outputPaths 要给同样多个。',
    ].join('\n');

    return { ok: true, topic, type: type.id, typeName: type.name, briefPath: shown, text: feedback };
}

/** office_search_dispatch：把提纲铺成子代理并发检索。 */
async function executeDispatch(rawArgs, exec, config, ctx) {
    const value = await dispatchSearch(rawArgs, exec, config, ctx);
    return { ok: value.ok, jobs: value.jobs, text: renderDispatch(value) };
}

/** office_parse_findings：把子代理的结果文件读成摘要。 */
async function executeParse(rawArgs, exec, config) {
    const value = await parseResultFiles(rawArgs, exec, config?.search?.resultLimit);
    return { ok: true, files: value.files, text: value.text };
}

/** 一次直查最多几条查询 / 每条最多几条来源（与检索提纲的渠道数上限同一量级）。 */
const SEARCH_RUN_MAX_QUERIES = 5;
const SEARCH_RUN_MAX_SOURCES_PER_QUERY = 20;
const SEARCH_RUN_MAX_REPLY_LINES = 5;

/**
 * 直查结果的落盘路径：调用方给了 out 就用它（相对工作目录解析），
 * 否则按第一条查询落在 <outputDir>/<slug>/run.md。
 */
function searchRunPath(root, config, queries, requested) {
    const asked = typeof requested === 'string' ? requested.trim() : '';
    if (asked !== '') return isAbsolute(asked) ? asked : join(root, asked);
    const dir = config?.search?.outputDir || '.office/search';
    const slug = String(queries[0] ?? 'search').trim().replace(/[\\/:*?"<>|\s]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'search';
    return join(root, dir, slug, 'run.md');
}

/**
 * 正文预处理的一句话报告。
 *
 * 它值一句话的位置，因为「洗掉了多少」直接决定这条摘录可不可信：压缩比高到
 * 离谱（比如 99%）通常意味着挑错了主容器，那时把 `search.preprocess.mode`
 * 退回 `off` / `plain` 就回到老行为。页数为 0 时返回空串（没取正文就别报）。
 */
export function preprocessReport(pages) {
    const list = Array.isArray(pages) ? pages.filter((item) => item?.stats !== undefined) : [];
    if (list.length === 0) return '';
    const input = list.reduce((sum, item) => sum + (item.stats.inputChars ?? 0), 0);
    const output = list.reduce((sum, item) => sum + (item.stats.outputChars ?? 0), 0);
    const dropped = list.reduce((sum, item) => sum + (item.stats.droppedLines ?? 0), 0);
    const deduped = list.reduce((sum, item) => sum + (item.stats.dedupedLines ?? 0), 0);
    const mode = list[0].mode ?? 'article';
    const percent = input === 0 ? 0 : Math.round((1 - output / input) * 1000) / 10;
    return `${list.length} 页（${mode} 模式）：${input.toLocaleString('en-US')} → ${output.toLocaleString('en-US')} 字符`
        + `（去掉 ${percent}%；丢样板行 ${dropped}、去重 ${deduped}）`;
}

/**
 * 用插件内置的检索通道直接查一轮：不起子代理，也不需要提纲。
 *
 * 这是「组合里没有 web_search 时也能查资料」的入口：办公 preset 是一份完整的
 * 组合，里面没有 @deepseek-ai/dsh-tool-web，所以办公会话的工具面里没有
 * web_search / web_fetch。内置通道优先用宿主的 web 服务（ctx.web），拿不到就
 * 自己发 HTTP（见 src/web.js）。
 *
 * 落盘与摘要的分工与三步走一致：来源与摘录写进文件，聊天里只回一份紧凑清单。
 */
async function executeSearchRun(rawArgs, exec, config, ctx) {
    const root = cwdOf(exec);
    const args = rawArgs !== null && typeof rawArgs === 'object' ? rawArgs : {};
    const rawQueries = Array.isArray(args.queries) ? args.queries : [args.queries];
    const queries = rawQueries.map((query) => (typeof query === 'string' ? query.trim() : '')).filter((query) => query !== '').slice(0, SEARCH_RUN_MAX_QUERIES);
    if (queries.length === 0) throw new Error('queries 不能为空：给一条或几条要查的话。');

    const settings = config?.search ?? {};
    const maxResults = Number.isFinite(args.maxResults)
        ? Math.max(1, Math.min(SEARCH_RUN_MAX_SOURCES_PER_QUERY, Math.trunc(args.maxResults)))
        : (Number.isFinite(settings?.builtin?.maxResults) ? settings.builtin.maxResults : 8);
    const fetchPages = Number.isFinite(args.fetchPages)
        ? Math.max(0, Math.min(5, Math.trunc(args.fetchPages)))
        : (Number.isFinite(settings?.builtin?.fetchPages) ? settings.builtin.fetchPages : 2);
    // 通道可以按次点名（第二十轮）：不传就用配置里的 provider / providerOrder。
    const wantedProvider = typeof args.provider === 'string' ? args.provider.trim() : '';
    const preprocessMode = typeof args.preprocess === 'string' ? args.preprocess.trim() : '';

    const access = createWebAccess(ctx, settings);
    const status = await access.probe();
    if (!status.ok) {
        throw new Error('内置检索用不了：' + status.reason
            + '。可以在设置页「办公模式 → 检索编排」里配通道与 Key，或用 office_search_dispatch 派子代理。');
    }

    const signal = exec?.signal ?? new AbortController().signal;
    const picked = [];
    const failures = [];
    // 取正文失败原先在这里被**静默吞掉**（`catch {}`）—— 于是「被墙 / 被挡 / 页面没了」
    // 在模型眼里等价于「这条来源没有摘录」，四类失败一类都传不出去（第十八轮 P0-9）。
    const fetchFailures = [];
    let engine = null;
    let provider = null;
    let fetched = 0;
    // 预处理报告：几页、原文多少字符、洗完剩多少、丢了几行 —— 进文件也进反馈。
    // 它是「这条摘录值不值得信」的直接依据：压缩比异常高时多半是挑错了主容器，
    // 那时把 search.preprocess.mode 退回 off / plain 就回到老行为。
    const preprocessed = [];
    for (const query of queries) {
        try {
            const result = await access.search(query, { maxResults, signal, provider: wantedProvider });
            engine = result.engine;
            if (typeof result.provider === 'string' && result.provider !== '') provider = result.provider;
            const sources = [];
            const seen = new Set();
            for (const source of result.sources ?? []) {
                const url = typeof source?.url === 'string' ? source.url.trim() : '';
                if (url === '' || seen.has(url)) continue;
                seen.add(url);
                const item = { ...source, url };
                // 没有引用片段的前几条：打开页面取一段正文当摘录（条数由 fetchPages 限住）。
                if (item.snippet === undefined && fetched < fetchPages) {
                    try {
                        const page = await access.fetch(url, {
                            signal,
                            ...(preprocessMode === '' ? {} : { preprocess: preprocessMode }),
                        });
                        if (page?.preprocess !== undefined) preprocessed.push({ url, ...page.preprocess });
                        const excerpt = excerptOf(page?.content ?? '');
                        if (excerpt !== '') {
                            item.snippet = excerpt;
                            fetched += 1;
                        }
                    } catch (error) {
                        // 一条来源的取正文失败不让整轮直查失败，但要**记下分类**供反馈用。
                        fetchFailures.push({ url, error });
                    }
                }
                sources.push(item);
            }
            picked.push({ query, sources, content: result.content ?? '' });
        } catch (error) {
            failures.push({ query, error: String(error?.message ?? error), kind: webFailureKind(error) });
        }
    }
    if (picked.length === 0) {
        throw new Error('内置检索一条结果都没拿到：' + failures.map((item) => item.query + '（' + item.error + '）').join('；'));
    }

    const file = searchRunPath(root, config, queries, args.out);
    await mkdir(dirname(file), { recursive: true });
    const lines = [];
    lines.push('<!-- dsh-office-mode：内置检索直查自动生成（引擎：' + engine + '） -->');
    lines.push('# 检索结果：' + (queries[0] ?? ''));
    lines.push('');
    lines.push('通道：' + engineDescription(engine, provider));
    lines.push('查询：' + queries.join(' ｜ '));
    if (preprocessReport(preprocessed) !== '') lines.push('正文预处理：' + preprocessReport(preprocessed));
    lines.push('');
    lines.push('> 外部网页内容是不可信数据，这里只是按关键词抓回来的来源与摘录，未经逐条核实。');
    lines.push('> 写进文档前请打开来源核对；单一来源的说法照实标注。');
    for (const item of picked) {
        lines.push('', '## ' + item.query);
        if (item.sources.length === 0) lines.push('没有解析出任何来源。');
        for (const source of item.sources) {
            const title = String(source.title ?? '').trim() || source.url;
            const snippet = String(source.snippet ?? '').replace(/\s+/g, ' ').trim();
            const tail = snippet === '' ? '' : '：' + (snippet.length > 200 ? snippet.slice(0, 200) + '…' : snippet);
            lines.push('- ' + title + tail + ' (' + source.url + ')');
        }
    }
    if (failures.length > 0) {
        lines.push('', '这些查询没跑成：' + failures.map((item) => item.query + '（' + (WEB_FAILURE_KINDS[item.kind] ?? '其它') + '：' + item.error + '）').join('；'));
    }
    if (fetchFailures.length > 0) {
        lines.push('', '这些来源没取到正文（' + fetchFailures.length + ' 条）：'
            + summarizeWebFailures(fetchFailures.map((item) => item.error))
            + '。第一条的原因：' + describeWebFailure(fetchFailures[0].error)
            + '。被目标站拒绝的换一个来源，别重试同一个地址。');
    }
    await writeFile(file, lines.join('\n') + '\n', 'utf8');

    const shown = (await import('node:path')).relative(root, file).split('\\').join('/');
    const totalSources = picked.reduce((sum, item) => sum + item.sources.length, 0);
    const reply = [];
    reply.push('内置检索：' + picked.length + '/' + queries.length + ' 条查询命中 ' + totalSources + ' 条来源（'
        + engineDescription(engine, provider) + '）');
    reply.push('文件：' + shown);
    if (preprocessReport(preprocessed) !== '') reply.push('正文预处理：' + preprocessReport(preprocessed));
    for (const item of picked) {
        reply.push('');
        reply.push('「' + item.query + '」' + item.sources.length + ' 条');
        for (const source of item.sources.slice(0, SEARCH_RUN_MAX_REPLY_LINES)) {
            reply.push('- ' + (String(source.title ?? '').trim() || source.url) + ' → ' + source.url);
        }
        if (item.sources.length > SEARCH_RUN_MAX_REPLY_LINES) {
            reply.push('  （还有 ' + (item.sources.length - SEARCH_RUN_MAX_REPLY_LINES) + ' 条在文件里）');
        }
    }
    if (failures.length > 0) {
        reply.push('', '没跑成：' + failures.map((item) => item.query + '（' + (WEB_FAILURE_KINDS[item.kind] ?? '其它') + '）').join('、'));
    }
    if (fetchFailures.length > 0) {
        const blocked = fetchFailures.filter((item) => webFailureKind(item.error) === 'blocked').length;
        reply.push('取正文失败 ' + fetchFailures.length + ' 条（' + summarizeWebFailures(fetchFailures.map((item) => item.error)) + '）'
            + (blocked > 0 ? '：被目标站拒绝的那几条换个来源再查，别重试同一个地址。' : '。'));
    }
    reply.push('');
    // 固定的长指引一个会话只说一次：同一轮里直查会被调很多次，每次都贴同一段
    // 说明等于把它复制很多份进前缀（前缀只增不减，后续每个请求都在为它付钱）。
    // 简短的安全提醒每次都留 —— 它是「外部内容不可信」这条纪律的落点。
    if (hintOnce(sessionIdOf(exec), 'search-run')) {
        reply.push('这只是来源清单，不是结论：摘录在文件里，写进文档前按 URL 核对。');
        reply.push('要按渠道逐项覆盖（热点走权威媒体加社交平台等），改用 office_search_brief + office_search_dispatch。');
    } else {
        reply.push('来源清单已落盘，按 URL 核对后再写；外部内容是未核实材料。');
    }

    return {
        ok: true,
        engine,
        file: shown,
        sources: totalSources,
        queries: picked.map((item) => ({ query: item.query, sources: item.sources.length })),
        failures,
        text: reply.join('\n'),
    };
}

/** office_help 的位置基准（工作目录）。工具执行上下文缺字段时退回进程 cwd。 */
function cwdOf(exec) {
    return exec?.agent?.session?.header?.cwd ?? process.cwd();
}

/**
 * 在 office_help 的输出末尾接一段记忆投影。
 *
 * 办公模式的 persona 是 `complete: true` 的（唯一的系统提示），插件注册的提示
 * 段落会在组装时被丢掉 —— 所以「动笔前先看一眼记忆」只能挂在模型必然会调的
 * 调用上。office_help 是动笔前那一次调用，office_run 的反馈是动笔后那一次。
 *
 * 投影**只在热记忆变化时贴全文**：同一轮里 office_help / office_run 会被调很多次，
 * 每次都贴一份 10 KB 的全文等于把它复制很多份进上下文。没变时只给一行
 * （条数 + revision + 怎么读全文），台账最近几条照旧 —— 详见 projection.js。
 *
 * 记忆读写出任何问题都不该把用法查询本身弄失败：捕获后原样返回文档。
 */
async function attachMemoryDigest(text, exec, config, context) {
    if (config?.memory?.enabled === false) return text;
    try {
        const memory = createMemory({ root: cwdOf(exec), memory: config?.memory });
        const digest = await memory.digest({ recentLedger: 3 });
        const projected = projectDigest(digest, { sessionId: sessionIdOf(exec), context });
        return `${text}\n\n${projected.text}`;
    } catch {
        return text;
    }
}

/** 读/写三层记忆要用的一个实例（root = 会话工作目录）。 */
function memoryOf(exec, config) {
    return createMemory({ root: cwdOf(exec), memory: config?.memory });
}

/**
 * 带 query 的记忆检索要过每回合配额。
 *
 * 口径（与 mnemon 的「一次初查 + 一次细化」一致）：
 *   本回合第一次带 query 的读取  → recallPerTurn
 *   同一回合里后续的换词细化      → recallRefinePerTurn
 * 两者都用完就拒绝，并在回执里说清楚「本回合已经查过几次」—— 拒绝必须可读，
 * 否则模型只会以为是检索坏了、然后换个工具接着试。
 */
function takeRecallQuota(quota, exec) {
    const first = quota.take('recallPerTurn', exec);
    if (first.allowed) return { allowed: true, decision: first, stage: 'recall' };
    const refine = quota.take('recallRefinePerTurn', exec);
    if (refine.allowed) return { allowed: true, decision: refine, stage: 'refine' };
    return { allowed: false, decision: refine, stage: 'denied' };
}

function quotaRefusal(decision, stage) {
    return [
        `📒 本回合的检索配额已用完（${decision.reason}）。`,
        '- 这是刻意的界：一个回合里查太多次，检索结果会把排版与措辞的余地挤出上下文。',
        '- 继续下去的办法：换 `layer` 只读单层、直接用 action:\'read\' 不带 query 浏览（不占配额）、或者基于已有结果动手写。',
        stage === 'denied' ? '- 要放宽就去设置页调「每回合配额」（0 = 不限制）。' : '',
    ].filter((line) => line !== '').join('\n');
}

/** Pack 文件的默认位置：记忆目录下按时间戳命名，不覆盖上一份。 */
function defaultPackPath(exec, config) {
    const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\..+$/, '').replace('T', '-');
    return join(memoryOf(exec, config).paths.dir, `pack-${stamp}.json`);
}

/**
 * 读/写三层记忆。
 *
 * 记忆目录始终在**会话工作目录**下：一个项目一份记忆。要跨项目共享偏好时，
 * 把设置里的「记忆范围」改成 both / global（全局层默认在 $DSH_HOME/.office/memory）——
 * 这是本轮并进来的 mnemon 跨项目层，不再是「一份记忆走天下」。
 */
async function executeMemory(rawArgs, exec, config, quota) {
    const args = rawArgs !== null && typeof rawArgs === 'object' ? rawArgs : {};
    const action = typeof args.action === 'string' && args.action.trim() !== '' ? args.action.trim() : 'read';
    if (config?.memory?.enabled === false) {
        return { ok: false, action, text: '记忆已关闭（设置页里关掉的）。要看内容请先把「记忆」打开。' };
    }
    const memory = memoryOf(exec, config);

    if (action === 'read') {
        const query = typeof args.query === 'string' ? args.query.trim() : '';
        // 只有「带关键词的检索」才占配额：不带 query 的浏览是看现状，不是查资料。
        if (query !== '') {
            const gate = takeRecallQuota(quota, exec);
            if (!gate.allowed) return { ok: false, action, text: quotaRefusal(gate.decision, gate.stage) };
        }
        const value = await memory.read({
            layer: typeof args.layer === 'string' ? args.layer : 'all',
            query,
            limit: Number.isFinite(args.limit) ? args.limit : 0,
        });
        return { ok: true, action, text: renderRead(value) };
    }

    if (action === 'log') {
        const value = await memory.log([{
            path: typeof args.path === 'string' ? args.path : '',
            format: typeof args.format === 'string' ? args.format : '',
            theme: typeof args.theme === 'string' ? args.theme : '',
            purpose: typeof args.purpose === 'string' ? args.purpose : '',
            source: Array.isArray(args.source) ? args.source : [],
        }]);
        return { ok: true, action, text: renderLog(value) };
    }

    if (action === 'status') {
        return { ok: true, action, text: renderStatus(await memory.status()) };
    }

    if (action === 'link') {
        const value = await memory.link({
            sourceId: typeof args.sourceId === 'string' ? args.sourceId : '',
            targetId: typeof args.targetId === 'string' ? args.targetId : '',
            kind: typeof args.kind === 'string' ? args.kind : '',
            note: typeof args.note === 'string' ? args.note : '',
        });
        return { ok: true, action, text: renderLink(value) };
    }

    if (action === 'unlink') {
        const value = await memory.unlink({ id: typeof args.id === 'string' ? args.id : '' });
        return { ok: true, action, text: renderUnlink(value) };
    }

    if (action === 'related') {
        const gate = quota.take('relatedPerTurn', exec);
        if (!gate.allowed) return { ok: false, action, text: quotaRefusal(gate, 'denied') };
        const value = await memory.related({
            id: typeof args.id === 'string' ? args.id : '',
            depth: Number.isFinite(args.depth) ? args.depth : 1,
            kind: typeof args.kind === 'string' ? args.kind : '',
            limit: Number.isFinite(args.limit) ? args.limit : 0,
        });
        return { ok: true, action, text: renderRelated(value) };
    }

    if (action === 'entities') {
        const value = await memory.entities({
            query: typeof args.query === 'string' ? args.query : '',
            limit: Number.isFinite(args.limit) ? args.limit : 0,
        });
        return { ok: true, action, text: renderEntities(value) };
    }

    if (action === 'export') {
        const pack = await memory.exportPack();
        const target = typeof args.packPath === 'string' && args.packPath.trim() !== ''
            ? (isAbsolute(args.packPath) ? args.packPath : join(cwdOf(exec), args.packPath))
            : defaultPackPath(exec, config);
        await mkdir(dirname(target), { recursive: true });
        const text = `${JSON.stringify(pack, null, 2)}\n`;
        await writeFile(target, text, 'utf8');
        const shown = (await import('node:path')).relative(cwdOf(exec), target).split('\\').join('/');
        return { ok: true, action, text: renderExport({ pack, path: shown, bytes: Buffer.byteLength(text, 'utf8') }) };
    }

    if (action === 'import') {
        const raw = typeof args.packPath === 'string' ? args.packPath.trim() : '';
        if (raw === '') throw new Error('import 需要 packPath（先 export，或用别的会话导出的那份）。');
        const target = isAbsolute(raw) ? raw : join(cwdOf(exec), raw);
        let parsed = null;
        try {
            parsed = JSON.parse(await readFile(target, 'utf8'));
        } catch (error) {
            throw new Error(`读不了这份 Pack（${raw}）：${error?.message ?? error}`);
        }
        const value = await memory.importPack(parsed);
        return { ok: true, action, text: renderImport(value) };
    }

    if (action === 'migrate') {
        const value = await migrateMnemon({
            root: cwdOf(exec),
            memory: config?.memory,
            source: typeof args.from === 'string' && args.from.trim() !== '' ? args.from : undefined,
            dryRun: args.dryRun === true,
        });
        return { ok: true, action, text: renderMigration(value) };
    }

    const value = await memory.mutate({
        action,
        target: typeof args.target === 'string' ? args.target : 'project',
        content: typeof args.content === 'string' ? args.content : '',
        oldText: typeof args.oldText === 'string' ? args.oldText : '',
        importance: typeof args.importance === 'string' ? args.importance : 'normal',
        entities: Array.isArray(args.entities) ? args.entities : [],
        tags: Array.isArray(args.tags) ? args.tags : [],
    });
    return { ok: true, action, text: renderMutation(value) };
}

/**
 * 构造工具定义数组。
 *
 * @param {ReturnType<import('./config.js').resolveConfig>} config
 * @param {object} [ctx] 插件上下文；检索工具用它经 `ctx.get('subagents')` 取子代理
 *     服务（**不能**直接读 `ctx.subagents`：没写进 inject 的服务属性在 cordis
 *     的 ctx 代理上会当场抛错，见 search.js 的 subagentsOf）。
 *     省略时（例如单测）检索派工不可用，其余工具照常。
 */
export function buildTools(config, ctx) {
    // 每回合配额计数器：跟着这一份工具面走（设置改动重建工具面时一并重置）。
    const quota = createTurnQuota({ limits: config?.memory?.quota ?? {} });
    const all = [
        {
            name: 'office_help',
            description:
                '办公模式用法查询：Word / Excel / PPT 的写法、LaTeX 学位论文（thuthesis）的生成与编译、PDF 的读法、Python 科学计算与绘图（python）、可用主题、批量改文件、缓存规则、三层记忆（memory）。'
                + '第一次用某种格式前先查一次（如 office_help({ topic: "word" })）。'
                + '默认只回最小事实集（接口签名与必需边界），够写出对的调用；要看例子与全部解释再传 detail:true —— 文档会留在后续每个请求的前缀里，能省则省。'
                + '返回值末尾会附一段记忆投影（用户偏好与项目约定 + 最近台账），动笔前照着它办。'
                + '不查直接凭记忆写 API 很容易错。',
            parameters: helpParameters,
            output: {
                schema: helpOutputSchema,
                render: (_args, value) => textBlock(String(value?.text ?? '（没有文档）')),
            },
            // office_help 的返回值末尾接一段记忆投影（热记忆 + 最近台账）：
            // 这是办公模式里最接近 mnemon「每轮热记忆投影」的位置。
            async execute(rawArgs, exec) {
                const args = rawArgs !== null && typeof rawArgs === 'object' ? rawArgs : {};
                const { topic, text } = await buildHelp(args.topic, { detail: args.detail === true });
                return { ok: true, topic, text: await attachMemoryDigest(text, exec, config, 'help') };
            },
        },
        {
            name: 'office_run',
            description:
                '一次调用完成一批办公操作：生成或修改 Word(.docx) / Excel(.xlsx) / PPT(.pptx)，生成 LaTeX 学位论文源码并编译 PDF（office.tex），读 PDF（office.pdf），用 Python 做科学计算与绘图（office.python），以及批量修改文本文件。'
                + 'script 里用 office 对象直接写多步；每次调用都从零开始，不要把状态留到下次。'
                + '写出的文件会自动重新打开检查，返回真实统计与排版风险。'
                + '没查过用法请先调用 office_help。',
            parameters: runParameters,
            output: {
                schema: runOutputSchema,
                render: (_args, value) => textBlock(renderRun(value)),
            },
            // 会写文件，不允许与其它调用并发。
            isConcurrencySafe: () => false,
            execute: (rawArgs, exec) => executeRun(rawArgs, exec, config),
        },
        {
            name: 'office_memory',
            description:
                '办公记忆：三层分开的长期记忆，跨会话保留在 .office/memory 里（可配成工作区 / 全局 / 两者并存）。'
                + '热记忆（action read 的 layer:hot）装用户偏好与项目约定，动笔前应当遵守；'
                + '台账（layer:ledger）是每份交付物的自动登记：路径、格式、主题、来源与复检统计，用来找「上次那份东西」；'
                + '归档（layer:archive）装热记忆装不下时下沉的旧条目与滚动的旧台账，只读。'
                + "action:'add' / 'replace' / 'remove' 改热记忆（target: user 记用户偏好，project 记项目与环境），"
                + "可带 entities（实体名，才会进实体视图）与 tags。"
                + "action:'log' 手工把一份产物登记进台账（office_run 写出的文件已自动登记）；"
                + "action:'link' / 'unlink' / 'related' 管条目之间的双向关系（related 沿关系找相邻条目，默认一跳）；"
                + "action:'entities' 看实体视图；action:'status' 看各层计数；"
                + "action:'export' / 'import' 整包备份与恢复（单文件 JSON，只增不改、按 id 幂等）；"
                + "action:'migrate' 把 mnemon 的记忆（runtime 热记忆 + 长期记忆）迁进这套三层记忆，先用 dryRun:true 看会迁什么。"
                + '带 query 的读取走召回质量分档（低相关会被丢掉），并计入每回合配额（默认一回合一次初查 + 一次换词细化）——'
                + '被配额挡住时会说明原因，不要换个工具重复同样的检索。'
                + '只记用户说过的偏好、纠正与稳定的事实；不要记检索原文、一次性进度或自己的猜测。'
                + '用户在这一轮里说出偏好或纠正时当场记一条，不必等他说「记住」。'
                + 'office_help 的返回值末尾本来就会带一份热记忆投影，所以平常不必特意 read。',
            parameters: memoryParameters,
            output: {
                schema: memoryOutputSchema,
                render: (_args, value) => textBlock(String(value?.text ?? '（没有记忆）')),
            },
            isConcurrencySafe: () => false,
            execute: (rawArgs, exec) => executeMemory(rawArgs, exec, config, quota),
        },
        {
            name: 'office_search_run',
            description:
                '直接查一轮（内置检索，不需要子代理、也不需要提纲）：给 1-5 条查询，插件的联网通道去找来源，'
                + '把标题与摘录写进结果文件，聊天里只回一份紧凑的来源清单。'
                + '办公 preset 里没有 web_search / web_fetch 这类工具，查资料就走这里。'
                + '通道有多条（第二十轮起）：默认按设置页的顺序自动挑（宿主 web 服务 → Anthropic 兼容 → 三方检索 API → 免 Key 兜底），'
                + '也可以用 provider 参数按次点名（bocha 中文好、serper 就是 Google、duckduckgo 免 Key …）。'
                + '取回来的网页会先过预处理（去导航与样板）再进上下文，preprocess 参数可以按次改强度。'
                + '要按渠道逐项覆盖（热点走权威媒体加社交平台、知识类走百科等）时，改用 office_search_brief + office_search_dispatch。'
                + '清单只是线索不是结论：写进文档前按 URL 核对。',
            parameters: searchRunParameters,
            output: {
                schema: searchRunOutputSchema,
                render: (_args, value) => textBlock(String(value?.text ?? '（没有结果）')),
            },
            isConcurrencySafe: () => false,
            execute: (rawArgs, exec) => executeSearchRun(rawArgs, exec, config, ctx),
        },
        {
            name: 'office_search_brief',
            description:
                '为一次检索生成提纲：按内容类型（热点事件 / 知识类 / 手册类）给出该走哪些渠道、每个渠道为什么必要、'
                + '查到什么程度算够，并把提纲写成文件。'
                + '规则是固定的——热点走权威媒体加各社交平台，知识走百科、需要更深再下沉文献，手册只认官方参考文档；'
                + '每个类型都先做一轮不限定来源的泛搜，避免只看一个来源。'
                + '提纲写好后交给 office_search_dispatch 执行：那里的材料只落在文件里，不会把原始搜索结果拉进上下文。'
                + '只查一两个事实点、不需要渠道覆盖时，直接用 office_search_run 更快。',
            parameters: briefParameters,
            output: {
                schema: briefOutputSchema,
                render: (_args, value) => textBlock(String(value?.text ?? '（没有提纲）')),
            },
            isConcurrencySafe: () => false,
            execute: (rawArgs, exec) => executeBrief(rawArgs, exec, config),
        },
        {
            name: 'office_search_dispatch',
            description:
                '把一份检索提纲铺开执行：每个渠道一个子代理，各自去调 web_search / advanced_search / platform_search，'
                + '把结果写进指定的结果文件。'
                + '组合里没有这些检索工具、或没有子代理服务时，它会自动改用插件的内置检索通道在进程内跑同样的渠道'
                + '（同一批结果文件、同一种格式，接着照样用 office_parse_findings 读）。'
                + '这是检索的执行环节——不论哪条通道，材料都只落盘，主上下文不会被搜索结果灌满。'
                + '要改渠道或查更深时，先 office_search_brief 重新出提纲。',
            parameters: dispatchParameters,
            output: {
                schema: dispatchOutputSchema,
                render: (_args, value) => textBlock(String(value?.text ?? '（没有派工）')),
            },
            isConcurrencySafe: () => false,
            execute: (rawArgs, exec) => executeDispatch(rawArgs, exec, config, ctx),
        },
        {
            name: 'office_parse_findings',
            description:
                '解析子代理写下的检索结果文件：按渠道归类结论、算出来源覆盖度、标出只有单一来源背书、'
                + '以及完全没带 URL 的条目，返回一份紧凑摘要。'
                + '写文档前用它读结果，这样进上下文的是几十行摘要而不是整页材料；'
                + '摘要里的「仍缺渠道」「待核」提示要照着补检索或照实标注。',
            parameters: parseParameters,
            output: {
                schema: parseOutputSchema,
                render: (_args, value) => textBlock(String(value?.text ?? '（没有结果）')),
            },
            execute: (rawArgs, exec) => executeParse(rawArgs, exec, config),
        },
    ];

    // 工具开关（设置页）：只注册开着的。关掉的工具不出现在模型面前，
    // 也不占每次请求的 schema 开销。缺省（无 config.tools）时全开，
    // 保证空配置永远可用。
    const switches = config?.tools ?? {};
    const enabled = all.filter((tool) => switches[tool.name] !== false);

    // 每个工具执行前记一下工作目录：记忆浏览面板的 HTTP 端点只服务「本进程真的
    // 在这里跑过工具」的根 —— 这是那条信任边界的唯一来源（见 view.js 的说明）。
    return enabled.map((tool) => ({
        ...tool,
        execute: (rawArgs, exec) => {
            noteWorkspace(cwdOf(exec));
            return tool.execute(rawArgs, exec);
        },
    }));
}