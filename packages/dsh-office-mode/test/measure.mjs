/**
 * 测量口径的可执行部分（第十八轮 §5.2 的三条纪律）—— `18-23` 的一半。
 *
 * 第十八轮那份研究给出五条纪律，其中三条能在**没有模型**的情况下落成断言，
 * 这一份就是它们；另外两条（抄评测设计不抄分数、报区间不报单点）体现在
 * `passK` 的返回形状里：**只要有一次不过、或者每次结果不一样，就不算通过**。
 *
 *   1. **`pass^k`（repeat-k 外壳）**：同一批检查跑 k 次，全部通过**且结果逐字相同**
 *      才算通过。它抓的是非确定性（时间、缓存、文件系统顺序）——那种「单次绿、
 *      下次红」的套件靠单跑永远看不出来。
 *   2. **等字节基线**：任何「加了一段提示之后指标变好」的结论，都要有一个**长度精确
 *      匹配**的、与办公无关的对照文本。这一份把它做成函数：字节数精确相等，
 *      且不含任何本项目的词。
 *   3. **多格式扰动回归**：同一份 persona 换标点、换分隔符、打乱条目与分节顺序，
 *      断言守它的判据抽出来的**语义指纹**不变（最差变体也要通过）。测不了模型对
 *      排版的敏感度，但能测「我们的判据依赖的是语义还是排版」——后者才是会腐烂的东西。
 *
 * 跑法：node test/measure.mjs
 */
import assert from 'node:assert/strict';

import { buildTools } from '../src/tools.js';
import { resolveConfig } from '../src/config.js';
import { renderCapabilityMap, syncCapabilities } from '../src/capabilities.js';
import {
    bytes,
    fingerprintEquals,
    personaFingerprint,
    readPersonaPrefix,
    renderPersona,
} from './persona-shape.mjs';

const results = [];
function check(name, fn) {
    try {
        const note = fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

// ── 外壳：三个纯函数（这一份自己的工具） ────────────────────────────────────

/**
 * `pass^k` 外壳：跑 k 次，返回「几次过、几次并一致」。
 *
 * `ok` 的口径是本项目要的那种严格：**全部过 且 每次结果逐字相同**。
 * 只报单点（第一次的结果）在统计上是噪声，`reports` 里给出每次的结果供下结论用。
 *
 * @param {number} k 重复次数
 * @param {() => any} run 一次检查（抛错 = 那一次没过）
 */
export function passK(k, run) {
    const runs = [];
    for (let index = 1; index <= k; index += 1) {
        try {
            runs.push({ index, ok: true, value: run(index) });
        } catch (error) {
            runs.push({ index, ok: false, error: error?.message ?? String(error) });
        }
    }
    const failed = runs.filter((item) => !item.ok);
    const signatures = new Set(runs.map((item) => JSON.stringify(item.ok ? item.value : `✗${item.error}`)));
    return {
        k,
        runs,
        passed: runs.length - failed.length,
        failed: failed.length,
        identical: signatures.size === 1,
        firstFailure: failed[0]?.index,
        ok: failed.length === 0 && signatures.size === 1,
    };
}

/**
 * 等字节基线：生成一段与给定文本**字节数精确相等**的、与办公无关的通用说明。
 *
 * 用途是第十八轮 §5.2 第 1 条：加了提示之后指标变好时，必须能拿一个等长的
 * 无关文本当作对照，排除「只是格式/接口对齐」的解释。中文补一个字符会超字节，
 * 所以最后用单字节字符补齐到精确相等。
 */
export function equalByteBaseline(text, options = {}) {
    const target = bytes(text);
    const filler = options.filler ?? '这是一段与办公任务无关的通用说明，用来做等字节对照。';
    let out = '';
    while (bytes(out + filler) <= target) out += filler;
    while (bytes(out) < target) out += '-';
    return out;
}

/** 确定性洗牌（LCG，不碰 Math.random：同样的种子必然同样结果）。 */
function seededShuffle(items, seed) {
    const list = [...items];
    let state = (seed >>> 0) || 1;
    for (let index = list.length - 1; index > 0; index -= 1) {
        state = (state * 1664525 + 1013904223) >>> 0;
        const pick = state % (index + 1);
        [list[index], list[pick]] = [list[pick], list[index]];
    }
    return list;
}

/** 分节标题所在的行号（判据与 persona-shape 的 personaSections 对齐）。 */
function titleIndices(text) {
    const lines = String(text).split('\n');
    const titles = new Set();
    for (const [index, line] of lines.entries()) {
        if (index === 0) continue;
        const trimmed = line.trim();
        if (trimmed === '' || trimmed.startsWith('-') || line.startsWith(' ')) continue;
        if (/^\{\{[^{}]*\}\}$/.test(trimmed)) continue;
        titles.add(index);
    }
    return [...titles].sort((left, right) => left - right);
}

/** 按分节切块：首块 = 首句到第一个分节标题之前；其余每块 = 一个分节（逐行切，可无损拼回）。 */
function splitSections(text) {
    const lines = String(text).split('\n');
    const cuts = titleIndices(text);
    if (cuts.length === 0) return [String(text)];
    const blocks = [lines.slice(0, cuts[0]).join('\n')];
    for (const [order, start] of cuts.entries()) {
        const end = order + 1 < cuts.length ? cuts[order + 1] : lines.length;
        blocks.push(lines.slice(start, end).join('\n'));
    }
    return blocks;
}

/** 扰动变体：换引号 / 换分隔符 / 换全角冒号 / 行尾加空格 / 每条目内打乱 / 分节打乱。 */
export function perturbPersona(text, variant, seed = 7) {
    const source = String(text);
    if (variant === 'quotes') return source.split('「').join('“').split('」').join('”');
    if (variant === 'separators') return source.split('、').join('，').split('；').join('。');
    if (variant === 'colons') return source.split('：').join(':');
    if (variant === 'slashes') return source.split(' / ').join('／');
    if (variant === 'trailing-space') return source.split('\n').map((line) => `${line} `).join('\n');
    if (variant === 'bullets') {
        return splitSections(source).map((block) => shuffleBullets(block, seed)).join('\n');
    }
    if (variant === 'sections') {
        const blocks = splitSections(source);
        return [blocks[0], ...seededShuffle(blocks.slice(1), seed)].join('\n');
    }
    throw new Error(`未知扰动变体：${variant}`);
}

/** 块内把 `- ` 条目（含缩进续行）打乱，其余行不动。 */
function shuffleBullets(block, seed) {
    const lines = String(block).split('\n');
    const spans = [];
    for (let index = 0; index < lines.length; index += 1) {
        if (!lines[index].trim().startsWith('- ')) continue;
        let end = index + 1;
        while (end < lines.length && lines[end].startsWith(' ') && lines[end].trim() !== '') end += 1;
        spans.push({ start: index, end });
        index = end - 1;
    }
    if (spans.length < 2) return block;
    const shuffled = seededShuffle(spans.map((span) => lines.slice(span.start, span.end)), seed);
    const out = [];
    let cursor = 0;
    for (let index = 0; index < lines.length; index += 1) {
        const span = spans[cursor];
        if (span !== undefined && index === span.start) {
            out.push(...shuffled[cursor]);
            index = span.end - 1;
            cursor += 1;
            continue;
        }
        out.push(lines[index]);
    }
    return out.join('\n');
}

/**
 * 一次真实测量批次：工具面 + 能力地图 + persona 指纹（`pass^k` 拿它当检查对象）。
 *
 * 探测带 `refresh: true`：不带的话第二到第五次会命中进程级快照缓存，`identical` 就变成
 * 「缓存保证的恒真」，而不是「非确定性被排除」（复核 瑕疵 2）。每次重探才是真重复。
 */
function measurementBatch() {
    const general = resolveConfig({});
    const tools = buildTools(general);
    const engines = syncCapabilities({ config: general, refresh: true });
    return {
        tools: tools.map((tool) => `${tool.name}:${bytes(JSON.stringify({
            name: tool.name,
            description: tool.description,
            parameters: tool.parameters,
        }))}`),
        engines,
        map: renderCapabilityMap(engines, { subagents: true }),
        persona: personaFingerprint(readPersonaPrefix().prefix),
    };
}

// ── 1. pass^k ──────────────────────────────────────────────────────────────

check('pass^k：同一批检查跑 5 次，全过且每次结果逐字相同才算通过', () => {
    const outcome = passK(5, measurementBatch);
    assert.equal(outcome.failed, 0, `有 ${outcome.failed} 次没过：${JSON.stringify(outcome.runs.find((item) => !item.ok))}`);
    assert.ok(outcome.identical, '5 次结果不一致 —— 说明这批检查里有非确定性（时间 / 缓存 / 目录顺序）');
    assert.ok(outcome.ok, 'pass^k 的判据是「全过且一致」');
    return `pass^5 = ${outcome.passed}/${outcome.k}，结果一致`;
});

check('pass^k 能抓住「第 3 次才红」的检查（外壳本身要验红）', () => {
    const outcome = passK(5, (index) => {
        if (index === 3) throw new Error('第三次才出错');
        return 'ok';
    });
    assert.equal(outcome.failed, 1);
    assert.equal(outcome.firstFailure, 3);
    assert.equal(outcome.ok, false, '有一次没过就不算通过');
    return '第 3 次红被抓住';
});

check('pass^k 能抓住「每次结果不一样」的非确定性', () => {
    const outcome = passK(3, (index) => `tick-${index}`);
    assert.equal(outcome.failed, 0, '不抛错不等于稳定');
    assert.equal(outcome.identical, false, '三次结果不同必须被判定为不稳定');
    assert.equal(outcome.ok, false);
    return '结果不一致被判定为不稳定';
});

// ── 2. 等字节基线 ──────────────────────────────────────────────────────────

check('等字节基线：字节数精确相等、可复算、不含本项目词', () => {
    const { prefix } = readPersonaPrefix();
    const rendered = renderPersona(prefix, renderCapabilityMap({
        pdf: { render: ['pdftoppm', 'pdftocairo', 'mutool', 'gs'], text: ['pdftotext'] },
        tex: ['latexmk', 'xelatex', 'lualatex'],
        preview: true,
        python: 'C:/Python313/python.exe',
        av: { ffmpeg: true, model: true },
    }, { subagents: false }));
    const baseline = equalByteBaseline(rendered);
    assert.equal(bytes(baseline), bytes(rendered), '等字节基线必须**精确**等字节，±容差不算');
    assert.equal(equalByteBaseline(rendered), baseline, '同样的输入必须给同样的基线');
    assert.ok(!/office_|\.office|DSH|word|excel|ppt/i.test(baseline), '基线里不能出现本项目的词，否则不是「无关说明」');
    return `对照 ${bytes(baseline)} B / 处理 ${bytes(rendered)} B`;
});

check('等字节基线跟着文本长度走（不是常量）', () => {
    const short = equalByteBaseline('短');
    const long = equalByteBaseline('长'.repeat(500));
    assert.equal(bytes(short), bytes('短'));
    assert.equal(bytes(long), bytes('长'.repeat(500)));
    assert.ok(bytes(long) > bytes(short));
});

// ── 3. 多格式扰动回归 ──────────────────────────────────────────────────────

const VARIANTS = ['quotes', 'separators', 'colons', 'slashes', 'trailing-space', 'bullets', 'sections'];

check(`多格式扰动：${VARIANTS.length} 个保语义变体的语义指纹都不变（最差变体也要过）`, () => {
    const { prefix } = readPersonaPrefix();
    const original = personaFingerprint(prefix);
    const lines = [];
    for (const variant of VARIANTS) {
        const perturbed = perturbPersona(prefix, variant);
        assert.notEqual(perturbed, prefix, `变体 ${variant} 没有真的改动文本 —— 这样的「扰动」是空转`);
        const fingerprint = personaFingerprint(perturbed);
        assert.ok(fingerprintEquals(fingerprint, original),
            `变体 ${variant} 之后语义指纹变了：${JSON.stringify(fingerprint)} ≠ ${JSON.stringify(original)}`);
        lines.push(`${variant}@${bytes(perturbed)}B`);
    }
    return lines.join(' ');
});

check('扰动回归的判据不是恒真：真坏掉的 persona 必须被指纹抓住', () => {
    const { prefix } = readPersonaPrefix();
    const original = personaFingerprint(prefix);
    // 负面控制一：删掉能力地图那一节的路由句。
    const withoutRouting = prefix.split('文档走 office_run').join('文档用对应工具');
    assert.ok(!fingerprintEquals(personaFingerprint(withoutRouting), original), '路由句没了必须被抓');
    // 负面控制二：删掉一整节（失败分支）。
    const withoutBranch = prefix.split('失败分支')[0];
    assert.ok(!fingerprintEquals(personaFingerprint(withoutBranch), original), '分节没了必须被抓');
    // 负面控制三：变量占位行没了（persona 就不再引用插件注入的能力）。
    const withoutVariable = prefix.split('{{office_capabilities}}').join('');
    assert.ok(!fingerprintEquals(personaFingerprint(withoutVariable), original), '变量引用没了必须被抓');
});

check('顺序扰动：分节顺序与条目顺序不影响工具面与能力地图的字节', () => {
    const general = resolveConfig({});
    const keys = Object.keys(general.tools ?? {});
    const shuffled = {};
    for (const key of seededShuffle(keys, 11)) shuffled[key] = general.tools[key];
    const left = buildTools(general).map((tool) => `${tool.name}:${bytes(JSON.stringify(tool.parameters))}`);
    const right = buildTools(resolveConfig({ tools: shuffled })).map((tool) => `${tool.name}:${bytes(JSON.stringify(tool.parameters))}`);
    assert.deepEqual(right, left, '工具面不该依赖配置键的顺序');
    const mapA = renderCapabilityMap(syncCapabilities({ config: general, refresh: true }), { subagents: false });
    const mapB = renderCapabilityMap(syncCapabilities({ config: general, refresh: true }), { subagents: false });
    assert.equal(mapB, mapA, '能力地图是纯函数：连读两次必须逐字相同');
});

// ── 输出 ──────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok && item.note !== undefined ? `  → ${item.note}` : ''}`);
    if (!item.ok) console.log(`       | ${item.error?.message ?? item.error}`);
}
console.log(`measure: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);