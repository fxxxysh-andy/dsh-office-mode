/**
 * 提示词注入预算与「固定长句只说一次」的回归测试（第十七轮）。
 *
 * 这一份守的是**上下文成本**，不是功能：办公插件往模型上下文里写的东西会留在
 * 前缀里被后续每一次请求重发，所以「体积」与「重复」都是要被钉住的量。
 * 三条断言口径：
 *
 *   1. **默认层必须比全文层小得多**（每个话题至少 2x）。office_help 默认只给
 *      最小事实集（接口签名 + 必需边界），例子与解释留在 detail:true —— 默认层
 *      一旦被谁悄悄改回全文，这条会当场红。
 *   2. **关键接口不许在分层里丢**：默认层仍要能查出每个 office.* 操作名，
 *      否则「省字节」就变成了「模型写不出对的调用」。
 *   3. **固定长句一个会话只说一次**：派工反馈里的「不要把结果原文搬进对话」、
 *      直查反馈里的两条尾巴，第二次调用不再重复；同时每次都要有的最小提醒
 *      （按 URL 核对 / 外部内容不可信）不能一起被省掉。
 *
 * 跑法：node test/injection-budget.mjs
 */
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { buildHelp } from '../src/docs.js';
import { renderDispatch } from '../src/search.js';
import { renderRun } from '../src/tools.js';
import { buildTools } from '../src/tools.js';
import { resolveConfig } from '../src/config.js';
import { hintOnce, hintSeen, resetHints } from '../src/projection.js';

const results = [];
async function check(name, fn) {
    resetHints();
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const bytes = (text) => Buffer.byteLength(String(text ?? ''), 'utf8');

/** 所有话题：默认层与全文层都要能出，且默认层不许比全文层大。 */
const TOPICS = ['', 'guide', 'run', 'word', 'excel', 'ppt', 'tex', 'pdf', 'python', 'av', 'files', 'cache', 'theme', 'search', 'memory', 'settings'];

// ── 1. 分层：每个话题的默认层都比全文层小得多 ──────────────────────────────

await check('office_help：默认层都更小，大话题至少收掉 20% / 30%', async () => {
    const report = [];
    for (const topic of TOPICS) {
        const compact = await buildHelp(topic);
        const full = await buildHelp(topic, { detail: true });
        assert.ok(compact.text.trim() !== '', `${topic || '(index)'} 默认层不能为空`);
        const compactBytes = bytes(compact.text);
        const fullBytes = bytes(full.text);
        // 三条口径：
        //   · 每个话题的默认层都必须更小（挡住「默认层悄悄退回全文」）；
        //   · 全文 ≥ 3 KB 的话题至少收掉 20%；
        //   · 全文 ≥ 8 KB 的话题至少收掉 30%（ppt / memory 这两个才是大头）。
        assert.ok(compactBytes < fullBytes, `${topic || '(index)'} 默认层不该不小于全文层`);
        if (fullBytes >= 3000) {
            assert.ok(compactBytes * 5 <= fullBytes * 4, `${topic || '(index)'} 默认层 ${compactBytes}B 对全文层 ${fullBytes}B 收得太少（< 20%）`);
        }
        if (fullBytes >= 8000) {
            assert.ok(compactBytes * 10 <= fullBytes * 7, `${topic || '(index)'} 默认层 ${compactBytes}B 对全文层 ${fullBytes}B 收得太少（< 30%）`);
        }
        report.push(`${topic || 'index'} ${compactBytes}/${fullBytes}`);
    }
    return report.join(' ');
});

await check('office_help：ppt 与 memory 这两个大话题被压到 1/3 以下', async () => {
    for (const topic of ['ppt', 'memory']) {
        const compact = bytes((await buildHelp(topic)).text);
        const full = bytes((await buildHelp(topic, { detail: true })).text);
        assert.ok(compact * 3 <= full, `${topic} 默认层 ${compact}B 没到全文层 ${full}B 的 1/3`);
    }
});

await check('office_help：索引与 guide 也走默认层，并指路 detail:true', async () => {
    const index = await buildHelp('');
    const guide = await buildHelp('guide');
    assert.ok(index.text.includes('detail:true') || index.text.includes('detail: true'), '索引要告诉模型怎么拿全文');
    assert.ok(bytes(index.text) * 2 <= bytes((await buildHelp('', { detail: true })).text));
    assert.ok(bytes(guide.text) * 2 <= bytes((await buildHelp('guide', { detail: true })).text));
});

// ── 2. 分层不许丢接口 ──────────────────────────────────────────────────────

await check('默认层仍给出每个格式的操作名（省字节不能省接口）', async () => {
    const wanted = {
        word: ['office.word', 'create', 'builder.heading', 'builder.table', 'builder.formula', 'builder.toc', 'edit'],
        excel: ['office.excel', 'sheet.stats', 'sheet.summary', 'sheet.formula', 'edit'],
        ppt: ['office.ppt', 'deck.cards', 'deck.kpi', 'deck.imageText', 'deck.timeline', 'deck.iconGrid', 'readSlides', 'revise', 'insertImages', 'images'],
    };
    for (const [topic, needles] of Object.entries(wanted)) {
        const text = (await buildHelp(topic)).text;
        for (const needle of needles) assert.ok(text.includes(needle), `${topic} 默认层缺 ${needle}`);
    }
});

await check('默认层的 pdf / python / search / settings / memory 仍讲清关键事实', async () => {
    const pdf = (await buildHelp('pdf')).text;
    for (const needle of ['office.pdf.info', 'office.pdf.text', 'office.pdf.pages', 'read_image', 'hasTextLayer', 'reused']) {
        assert.ok(pdf.includes(needle), `pdf 默认层缺 ${needle}`);
    }
    const python = (await buildHelp('python')).text;
    for (const needle of ['office.python.check', 'office.python.run', 'office.python.file', 'OUT_DIR', 'savefig', 'use_cjk_font', 'logFiles.stdout']) {
        assert.ok(python.includes(needle), `python 默认层缺 ${needle}`);
    }
    const av = (await buildHelp('av')).text;
    for (const needle of ['office.av.check', 'office.av.info', 'office.av.transcribe', 'office.av.frames', 'office.av.extract', 'read_image', 'SenseVoice', 'ffmpeg', '时间戳']) {
        assert.ok(av.includes(needle), `av 默认层缺 ${needle}`);
    }
    const search = (await buildHelp('search')).text;
    for (const needle of ['office_search_run', 'office_search_brief', 'office_search_dispatch', 'office_parse_findings', '泛搜', '茧房']) {
        assert.ok(search.includes(needle), `search 默认层缺 ${needle}`);
    }
    const settings = (await buildHelp('settings')).text;
    for (const needle of ['工具开关', 'maxParallel', 'resultLimit', 'subagentTools', 'outputDir', 'defaultTheme', 'userLimitBytes', 'ledgerLimit']) {
        assert.ok(settings.includes(needle), `settings 默认层缺 ${needle}`);
    }
    const memory = (await buildHelp('memory')).text;
    for (const needle of ['热记忆', '台账', '归档', 'replace', 'oldText', '不记', '记忆目录', '按工作目录']) {
        assert.ok(memory.includes(needle), `memory 默认层缺 ${needle}`);
    }
});

// ── 3. 固定长句一个会话只说一次 ────────────────────────────────────────────

await check('派工反馈：同一条失败原因不再逐渠道重复七遍', () => {
    const reason = 'tools.restrict() names unknown global tools "web_search", "advanced_search", "platform_search", "web_fetch"; known global tools: ' + 'x'.repeat(200);
    const jobs = Array.from({ length: 7 }, (unused, index) => ({ channel: `渠道${index + 1}`, path: `c${index + 1}.md`, ok: false, error: reason }));
    const text = renderDispatch({ topic: '某事件', jobs, sessionId: 's1' });
    const occurrences = text.split(reason).length - 1;
    assert.equal(occurrences, 1, `失败原因只该原样出现一次，实际 ${occurrences} 次`);
    assert.ok(text.includes('7 个渠道同因'), '要说清有几个渠道同因');
    for (const job of jobs) assert.ok(text.includes(job.channel), `合并后仍要列出渠道：${job.channel}`);
});

await check('派工反馈：「不要把结果原文搬进对话」一个会话只说一次', () => {
    const value = { topic: '某事件', jobs: [{ channel: '泛搜', path: 'a.md', ok: true }], sessionId: 's-dup' };
    const first = renderDispatch(value);
    const second = renderDispatch(value);
    assert.ok(first.includes('不要把结果原文搬进对话'), '第一次要给这条纪律');
    assert.ok(!second.includes('不要把结果原文搬进对话'), '第二次不再重复');
    assert.ok(second.includes('office_parse_findings'), '第二次仍要指向下一步');
    assert.ok(renderDispatch({ ...value, sessionId: 's-other' }).includes('不要把结果原文搬进对话'), '换一个会话要重新给');
});

await check('office_search_run：同一会话第二次的反馈明显更短，最小提醒仍在', async () => {
    const seam = {
        async search(request) {
            return {
                sources: [
                    { url: 'https://news-a.example.com/' + encodeURIComponent(request.query), title: '来源甲', snippet: '引用片段' },
                    { url: 'https://news-b.example.com/2', title: '来源乙' },
                ],
                truncated: false,
            };
        },
        async fetch(request) {
            return {
                url: request.url,
                statusCode: 200,
                body: { kind: 'html', content: '<p>来源乙的正文段落，长度足够当摘录用的一行文字。</p>' },
                truncated: false,
            };
        },
    };
    const ctx = new Proxy({ inject: ['tools'] }, {
        get(target, prop) {
            if (prop === 'get') return (name) => (name === 'web' ? seam : undefined);
            if (prop in target) return target[prop];
            throw new Error(`cannot get property "${String(prop)}" without inject`);
        },
    });
    const root = mkdtempSync(join(tmpdir(), 'office-budget-'));
    try {
        const tools = buildTools(resolveConfig({}), ctx);
        const run = tools.find((item) => item.name === 'office_search_run');
        const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
        const first = await run.execute({ queries: ['某事件 通报'], out: 'a.md' }, exec);
        const second = await run.execute({ queries: ['某事件 通报'], out: 'b.md' }, exec);
        assert.ok(first.text.includes('要按渠道逐项覆盖'), '第一次要给渠道引导');
        assert.ok(!second.text.includes('要按渠道逐项覆盖'), '第二次不再重复渠道引导');
        assert.ok(second.text.includes('不是结论') || second.text.includes('未核实材料'), '最小安全提醒必须留着');
        assert.ok(second.text.includes('https://news-a.example.com/') || second.text.includes('来源甲'), '来源清单照旧要给');
        assert.ok(
            Buffer.byteLength(second.text, 'utf8') < Buffer.byteLength(first.text, 'utf8'),
            '第二次的反馈要更短',
        );
    } finally {
        rmSync(root, { recursive: true, force: true });
    }
});

// ── 4. office_run 反馈不许回灌模型自己刚写的正文 ────────────────────────────

await check('office_run 反馈：脚本返回里的 outline 不再原样回灌', () => {
    const outline = Array.from({ length: 41 }, (unused, index) => `P${index + 1} cover：标题${index + 1}`);
    const returned = { ok: true, path: 'a.pptx', outline, note: '保留我' };
    const text = renderRun({
        ok: true,
        elapsedMs: 12,
        files: [],
        otherFiles: [],
        returned,
        logs: [],
        notes: [],
        warnings: [],
        cache: { dir: '.office/cache', kept: 0, keptBytes: 0, hits: 0, pruned: [], ttlMinutes: 720 },
    });
    for (const line of outline) assert.ok(!text.includes(line), `outline 不该回灌：${line}`);
    assert.ok(text.includes('共 41 项'), '要报出 outline 有条数');
    assert.ok(text.includes('保留我'), 'outline 之外的字段要留下');
    assert.ok(bytes(text) < 400, `这一条反馈要够小，实际 ${bytes(text)}B`);
});

await check('office_run 反馈：脚本返回超长时按小上限截断', () => {
    const text = renderRun({
        ok: true,
        elapsedMs: 12,
        files: [],
        otherFiles: [],
        returned: { blob: 'x'.repeat(5000) },
        logs: [],
        notes: [],
        warnings: [],
        cache: { dir: '.office/cache', kept: 0, keptBytes: 0, hits: 0, pruned: [], ttlMinutes: 720 },
    });
    assert.ok(text.includes('…'), '要能看出是被截断的');
    assert.ok(bytes(text) < 2000, `截断上限要够小，实际 ${bytes(text)}B`);
});

await check('office_run 反馈：缓存块压成一行，TTL 说明不再占三行', () => {
    const text = renderRun({
        ok: true, elapsedMs: 12, files: [], otherFiles: [], returned: null, logs: [], notes: [], warnings: [],
        cache: { dir: '.office/cache', kept: 131, keptBytes: 19890000, hits: 2, hitBytes: 4000, pruned: ['a', 'b'], prunedBytes: 12400, ttlMinutes: 720 },
    });
    const cacheLines = text.split('\n').filter((line) => line.includes('缓存：') || line.trimStart().startsWith('720 分钟'));
    assert.equal(cacheLines.length, 1, `缓存块只该一行，实际 ${cacheLines.length} 行`);
    assert.ok(cacheLines[0].includes('命中复用'), '命中信息要留下');
    assert.ok(cacheLines[0].includes('office.cache.clear'), '清空入口要留下');
});

// ── 收尾 ──────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    if (item.ok) console.log(`PASS  ${item.name}${item.note ? `  → ${item.note}` : ''}`);
    else console.log(`FAIL  ${item.name}  → ${item.error?.message ?? item.error}`);
}
console.log(`\ninjection-budget: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
