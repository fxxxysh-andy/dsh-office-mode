/**
 * 引擎与管线的冒烟测试（不依赖格式模块是否完成）。
 *
 * 覆盖：ZIP 往返、XML 解析、文本度量、主题解析、批量文件编辑只写一次盘、
 * 脚本错误行号换算、缓存清理。
 *
 * 跑法：node test/smoke.mjs
 */
import { mkdtempSync, rmSync, existsSync, readFileSync, writeFileSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { zip, unzip, unzipText, crc32, editPackage } from '../src/engine/zip.js';
import { parseXml, textOf, descendants, escapeXml } from '../src/engine/xml.js';
import { estimateEm, wrapText, slugify, ensureExtension, normalizeColumns, normalizeRows, createEnv, toNumber } from '../src/engine/kit.js';
import { DEFAULT_THEME_ID, resolveTheme, shadeOf, themeCatalog, toHex } from '../src/engine/theme.js';
import { createCache } from '../src/engine/cache.js';
import { resolveConfig } from '../src/config.js';
import { buildHelp } from '../src/docs.js';
import { executeRun } from '../src/run.js';
import { buildTools, officeRunErrorResult, renderRun } from '../src/tools.js';

const results = [];
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const root = mkdtempSync(join(tmpdir(), 'office-smoke-'));

await check('crc32 已知向量', () => {
    assert.equal(crc32(new TextEncoder().encode('123456789')), 0xcbf43926);
});

await check('zip 往返（含中文名与压缩）', () => {
    const big = '中文内容'.repeat(200);
    const bytes = zip([
        { name: '[Content_Types].xml', data: '<Types/>' },
        { name: 'word/document.xml', data: big },
        { name: 'bin.dat', data: new Uint8Array([1, 2, 3, 4, 5]) },
    ]);
    const text = unzipText(bytes);
    assert.equal(text.get('[Content_Types].xml'), '<Types/>');
    assert.equal(text.get('word/document.xml'), big);
    const raw = unzip(bytes);
    assert.deepEqual([...raw.get('bin.dat')], [1, 2, 3, 4, 5]);
    assert.equal(raw.size, 3);
});

await check('editPackage 只改指定 part', () => {
    const bytes = zip([{ name: 'a.xml', data: '<a>旧</a>' }, { name: 'b.bin', data: new Uint8Array([9]) }]);
    const next = editPackage(bytes, (name, content) => (name === 'a.xml' ? content.replace('旧', '新') : undefined));
    assert.equal(unzipText(next).get('a.xml'), '<a>新</a>');
    assert.deepEqual([...unzip(next).get('b.bin')], [9]);
});

await check('XML 解析（属性 / 自闭合 / CDATA / 实体）', () => {
    const doc = parseXml('<?xml version="1.0"?><w:body a="1" b=\'x\'>text<w:p/><w:t xml:space="preserve">A &amp; B &#65;</w:t><![CDATA[<raw>]]></w:body>');
    const body = doc.children[0];
    assert.equal(body.name, 'w:body');
    assert.equal(body.attrs.a, '1');
    assert.equal(body.attrs.b, 'x');
    assert.equal(descendants(body, 'w:t').length, 1);
    assert.equal(textOf(descendants(body, 'w:t')[0]), 'A & B A');
    assert.ok(textOf(body).includes('<raw>'));
    assert.equal(escapeXml('<a & "b">'), '&lt;a &amp; &quot;b&quot;&gt;');
});

await check('文本度量：中文比西文宽，折行不溢出', () => {
    assert.ok(estimateEm('中文中文') > estimateEm('abcd'));
    const lines = wrapText('这是一段需要折行的中文说明文字'.repeat(3), 60, 12);
    assert.ok(lines.length >= 3);
    for (const line of lines) assert.ok(estimateEm(line) * 12 <= 60 + 12, `行超宽：${line}`);
});

await check('数据规范化', () => {
    assert.equal(slugify('季度 汇报/2026:Q1'), '季度-汇报-2026-Q1');
    assert.equal(ensureExtension('报告', '.docx'), '报告.docx');
    assert.equal(ensureExtension('报告.docx', '.docx'), '报告.docx');
    const cols = normalizeColumns(['名称', { title: '金额', type: 'number' }]);
    assert.equal(cols[0].title, '名称');
    assert.equal(cols[1].type, 'number');
    assert.deepEqual(normalizeRows([['a']], 3), [['a', '', '']]);
    assert.equal(toNumber('¥1,234.5'), 1234.5);
    assert.equal(toNumber('12%'), 0.12);
});

await check('主题：默认是素色网格，未知 id 回落并报告', () => {
    // 默认主题必须是 plain：不传主题的产出不带任何背景填充
    assert.equal(DEFAULT_THEME_ID, 'plain');
    assert.equal(resolveTheme().theme.id, 'plain');
    assert.equal(resolveTheme('business').fellBack, false);
    const miss = resolveTheme('nope');
    assert.equal(miss.fellBack, true);
    assert.equal(miss.theme.id, DEFAULT_THEME_ID);
    assert.ok(themeCatalog().length >= 5);
    assert.equal(toHex('#abc'), 'AABBCC');
    for (const theme of themeCatalog()) assert.match(theme.primary, /^[0-9A-F]{6}$/);
});

await check('主题：素色主题无填充，shadeOf 把空串与白色归一成不填充', () => {
    const plain = resolveTheme('plain').theme;
    for (const key of ['titleFill', 'headerFill', 'zebraFill', 'accentFill']) {
        assert.equal(plain.table[key], '', `plain.table.${key} 应为空串（不填充）`);
    }
    // 彩色主题仍然有底色，别把主题能力一起改掉
    const business = resolveTheme('business').theme;
    assert.notEqual(business.table.headerFill, '');
    assert.notEqual(business.table.zebraFill, '');
    assert.equal(shadeOf(''), '');
    assert.equal(shadeOf(undefined), '');
    assert.equal(shadeOf('FFFFFF'), '');
    assert.equal(shadeOf('#f2f6fb'), 'F2F6FB');
    assert.equal(shadeOf('1F4E79'), '1F4E79');
});

await check('配置：默认主题跟随 DEFAULT_THEME_ID', () => {
    const saved = process.env.DSH_OFFICE_THEME;
    delete process.env.DSH_OFFICE_THEME;
    try {
        assert.equal(resolveConfig({}).defaultTheme, DEFAULT_THEME_ID);
        assert.equal(resolveConfig({ defaultTheme: 'business' }).defaultTheme, 'business');
    } finally {
        if (saved === undefined) delete process.env.DSH_OFFICE_THEME;
        else process.env.DSH_OFFICE_THEME = saved;
    }
});

await check('缓存：写入 / 列出 / 清空', () => {
    const cache = createCache({ root, dir: '.office/cache' });
    cache.write('preview.html', '<html/>');
    assert.equal(cache.list().length, 1);
    assert.equal(cache.rel, '.office/cache');
    assert.ok(existsSync(cache.path('preview.html')));
    // 2026-09-23：列表要带修改时间（TTL 清理与「这是不是刚渲染的」都靠它）。
    assert.equal(typeof cache.list()[0].modifiedMs, 'number');
    assert.equal(cache.list()[0].modifiedMs > 0, true);
    // 子目录（PDF 渲染结果按内容键分目录放）
    cache.write('pdf/abc/page-1.png', Buffer.from([1, 2, 3]));
    assert.equal(cache.list().some((entry) => entry.name === 'pdf/abc/page-1.png'), true);
    assert.throws(() => cache.path('../escape.txt'), /跳出缓存目录/);
    assert.equal(cache.clear(), true);
    assert.equal(cache.list().length, 0);
    assert.equal(cache.clear(), false);
});

await check('配置：非法值报错，环境变量可兜底', () => {
    const conf = resolveConfig({});
    assert.equal(conf.cacheDir, '.office/cache');
    assert.equal(conf.injectGuide, false);
    assert.throws(() => resolveConfig({ defaultTheme: '不存在' }), /defaultTheme/);
    assert.throws(() => resolveConfig({ cacheDir: 5 }), /cacheDir/);
});

await check('office_help：索引与话题', async () => {
    const index = await buildHelp('');
    assert.match(index.text, /office_run|格式/);
    const guide = await buildHelp('guide');
    assert.match(guide.text, /办公模式/);
    const unknown = await buildHelp('不存在的格式');
    assert.match(unknown.text, /没有「不存在的格式」这个话题/);
    const cache = await buildHelp('cache');
    assert.match(cache.text, /office\.cache\.clear/);
    // tex 的长文档由格式模块的 meta.guide 提供（docs.js 只负责转交）：
    // 这条路径断了，office_help({topic:'tex'}) 会退回自动生成的简短话题。
    const tex = await buildHelp('tex');
    assert.match(tex.text, /office\.tex\.create/);
    assert.match(tex.text, /office\.tex\.compile/);
    assert.match(index.text, /tex/, '索引里应当列出 tex 格式');
});

await check('office_run：批量改 Markdown 只写一次盘', async () => {
    writeFileSync(join(root, 'notes.md'), '# 标题\n\n旧内容A\n旧内容B\n', 'utf8');
    const result = await executeRun({
        script: `
            const r = office.files.edit('notes.md', [['旧内容A', '新内容A'], ['旧内容B', '新内容B'], ['不存在的串', 'x']]);
            office.assert(r.missing.length === 1, '未命中的替换要被报告出来');
            office.files.write('out/小结.md', '# 小结\\n\\n做完了。');
            office.log('改了 ' + r.applied.length + ' 处');
            return { applied: r.applied.length, missing: r.missing.length };
        `,
        purpose: '批量改 Markdown',
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));

    assert.equal(result.ok, true, result.error ? `${result.error.message} @${result.error.line}` : '');
    assert.equal(readFileSync(join(root, 'notes.md'), 'utf8').includes('新内容A'), true);
    assert.equal(result.otherFiles.length, 2);
    assert.equal(result.returned.applied, 2);
    assert.equal(result.returned.missing, 1);
    assert.equal(result.cache.clearedAfter, false, '默认保留缓存，不再每次调用清空');
    assert.equal(result.cache.kept, 0);

    const text = renderRun(result);
    assert.match(text, /✅ office_run 完成/);
    assert.match(text, /notes\.md/);
});

await check('office_run：脚本报错时给出原始行号', async () => {
    const result = await executeRun({
        script: 'const a = 1;\n\noffice.assert(a === 2, "这里应该失败");\n',
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, false);
    assert.equal(result.error.line, 3, `期望第 3 行，实际 ${String(result.error.line)}`);
    assert.match(renderRun(result), /❌ office_run 执行失败/);
});

await check('office_run：脚本内抛出的异常也能定位（vm realm 的错误实例）', async () => {
    // 脚本跑在 vm 里，抛出的 TypeError 是另一个 realm 的实例，宿主侧
    // `instanceof Error` 为假。早先按 String(error) 重新包装，行号与调用栈全丢，
    // 模型只看到「脚本第 null 行」——2026-09-22 的实测里，这直接让模型花了
    // 四五次调用去猜崩溃位置。这里钉住：位置要拿到，且不能渲染成 null。
    const result = await executeRun({
        script: 'const a = 1;\nconst b = a.missing.length;\nreturn b;\n',
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, false);
    assert.equal(result.error.name, 'TypeError');
    assert.match(result.error.message, /length/);
    assert.equal(typeof result.error.line, 'number', `期望拿到行号，实际 ${String(result.error.line)}`);
    assert.equal(result.error.line, 2, `期望第 2 行，实际 ${String(result.error.line)}`);
    assert.equal(typeof result.error.column, 'number');
    const frames = result.error.stack;
    assert.equal(Array.isArray(frames), true);
    assert.equal(frames.some((frame) => frame.includes('office-script.js')), true,
        `调用栈里应当能看到脚本自身的帧：${JSON.stringify(frames)}`);
    assert.doesNotMatch(renderRun(result), /第 null 行/);
});

await check('office_run：无 process / require / fs', async () => {
    const result = await executeRun({
        script: 'return { p: typeof process, r: typeof require, f: typeof fetch, e: typeof eval };',
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.p, 'undefined');
    assert.equal(result.returned.r, 'undefined');
    assert.equal(result.returned.f, 'undefined');
});

await check('office_run：空脚本与超长脚本被拒绝', async () => {
    await assert.rejects(() => executeRun({ script: '  ' }, {}, resolveConfig({})), /script 不能为空/);
    await assert.rejects(
        () => executeRun({ script: 'x'.repeat(2000) }, {}, resolveConfig({ maxScriptChars: 1000 })),
        /script 太长/,
    );
});

await check('office_run：结果必须是可无损 JSON 的值（不能有 undefined）', async () => {
    const result = await executeRun({
        script: 'return { a: undefined, b: [1, undefined], c: { d: undefined } };',
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    const walk = (value, path) => {
        assert.notEqual(value, undefined, `${path} 是 undefined`);
        if (value === null || typeof value !== 'object') return;
        for (const [key, item] of Object.entries(value)) walk(item, `${path}.${key}`);
    };
    walk(result, 'result');
    // 宿主校验用的正是这一步：丢字段会在这里露出来。
    assert.deepEqual(JSON.parse(JSON.stringify(result)), result);
});

await check('office_run：返回值满足宿主对声明字段的校验（声明即必填）', async () => {
    // 宿主按 output.schema 逐项校验返回值，且**声明过的字段缺键即判非法**：
    // 2026-09-21 实测，成功路径不写 error 会被报成
    // `tool "office_run" returned invalid output: "value.error" must be an object`，
    // 不传 purpose 会被报成 `"value.purpose" must be a string` —— 脚本跑成功也拿不到反馈。
    // 所以这里按「声明即必填」断言，而不是像旧版那样 `if (!(key in value)) continue`
    // 把缺键放过（正是那条 continue 让线上故障躲过了这套测试）。
    const run = buildTools(resolveConfig({})).find((item) => item.name === 'office_run');
    assert.ok(run?.output?.schema, 'office_run 必须声明输出 schema');
    const assertShape = (value, label) => {
        for (const [key, spec] of Object.entries(run.output.schema.properties ?? {})) {
            assert.equal(key in value, true, `${label} 缺少声明过的字段 ${key}`);
            const actual = value[key];
            assert.notEqual(actual, undefined, `${label}.${key} 是 undefined`);
            if (spec.type === 'object') {
                assert.equal(typeof actual, 'object', `${label}.${key} 必须是对象`);
                assert.notEqual(actual, null, `${label}.${key} 不能是 null`);
            }
            if (spec.type === 'array') assert.equal(Array.isArray(actual), true, `${label}.${key} 必须是数组`);
            if (spec.type === 'string') assert.equal(typeof actual, 'string', `${label}.${key} 必须是字符串`);
            if (spec.type === 'integer') assert.equal(Number.isInteger(actual), true, `${label}.${key} 必须是整数`);
            if (spec.type === 'boolean') assert.equal(typeof actual, 'boolean', `${label}.${key} 必须是布尔值`);
        }
    };
    // 成功路径 + 传了 purpose：老 schema 会在这里被判非法（缺 error）。
    const ok = await executeRun(
        { script: 'return 1;', purpose: '形状自测' },
        { agent: { session: { header: { cwd: root } } } },
        resolveConfig({}),
    );
    assert.equal(ok.ok, true);
    assert.equal('error' in ok, false, '成功时不该带 error 字段');
    assertShape(ok, 'ok');

    // 失败路径 + 不传 purpose：老 schema 会在「error 是对象」之外再因缺 purpose 被判非法。
    const noPurpose = await executeRun(
        { script: 'throw new Error("boom");' },
        { agent: { session: { header: { cwd: root } } } },
        resolveConfig({}),
    );
    assert.equal(noPurpose.ok, false);
    assertShape(noPurpose, 'noPurpose');

    const bad = await executeRun(
        { script: 'throw new Error("boom");', purpose: '形状自测' },
        { agent: { session: { header: { cwd: root } } } },
        resolveConfig({}),
    );
    assert.equal(bad.ok, false);
    assert.equal(typeof bad.error, 'object');
    assert.notEqual(bad.error, null);
    assertShape(bad, 'bad');
});

await check('office.ppt 暴露 readSlides / revise（SDK 挂载面不能静默丢失）', async () => {
    // 2026-09-22：这两条是「生成后审阅 + 微调」的入口。它们的挂法比较绕 ——
    // 实现在 formats/ppt-revise.js，经 ppt.js 的 `api` 惰性工厂 + sdk.js 的 wrap()
    // 挂到 office.ppt 上。任何一环改名都会让 office_run 里的脚本拿到
    // `undefined is not a function`，而插件本身照常加载、其它格式毫发无损 ——
    // 这种「静默少一个方法」正是最该被断言钉住的一类。这里直接问 SDK 要能力。
    const loaded = await import('../src/registry.js');
    const pptEntry = loaded.formatEntry('ppt');
    assert.ok(pptEntry, '注册表里必须有 ppt');
    const mod = await (await loaded.loadFormat('ppt')).module;
    assert.equal(typeof mod.api?.readSlides, 'function', 'ppt.js 必须导出 api.readSlides 工厂');
    assert.equal(typeof mod.api?.revise, 'function', 'ppt.js 必须导出 api.revise 工厂');

    // 真正跑一遍：造 → 读结构 → 改文字/样式/位置 → 读回复检
    const result = await executeRun({
        script: `
const deck = office.ppt.create({ title: 'SDK 冒烟', theme: 'business', path: 'smoke-deck.pptx' });
deck.cover({ title: '原题' }).bullets({ title: '第二页', items: ['甲', '乙'] }).closing();
deck.save();

const before = await office.ppt.readSlides('smoke-deck.pptx');
const title = before.pages[0].shapes.find((s) => s.name === 'Title');
if (!title) throw new Error('readSlides 没找到封面 Title');
if (typeof title.x !== 'number') throw new Error('readSlides 没给出几何');

const rev = await office.ppt.revise('smoke-deck.pptx', [
  { slide: 1, shape: 'title', setText: '改过的题' },
  { slide: 1, shape: 'title', style: { sizePt: 30, color: '0B5394' } },
  { slide: 1, shape: 'title', move: { x: 3, y: 4 } },
]);
if (!rev.ok || rev.applied.length !== 3) throw new Error('revise 没有全部生效：' + JSON.stringify(rev.skipped));

const after = await office.ppt.readSlides('smoke-deck.pptx');
const moved = after.pages[0].shapes.find((s) => s.name === 'Title');
return { text: moved.text, x: moved.x, y: moved.y, pages: before.pages.length };
`,
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.text, '改过的题');
    assert.equal(result.returned.x, 3);
    assert.equal(result.returned.y, 4);
    assert.equal(result.returned.pages, 3);
});

await check('office.tex 暴露 create / read / edit / compile（SDK 挂载面不能静默丢失）', async () => {
    // 与上面 ppt 那条同一个道理：compile / engines / escape 走的是 tex.js 的 `api`
    // 惰性工厂 + sdk.js 的 wrap()，任何一环改名都会让脚本拿到
    // `undefined is not a function`，而插件照常加载 —— 只能靠断言钉住。
    const loaded = await import('../src/registry.js');
    assert.ok(loaded.formatEntry('tex'), '注册表里必须有 tex');
    assert.equal(loaded.formatByExtension('thesis/thesis.tex')?.id, 'tex', '.tex 必须映射到 tex 格式');
    const mod = await (await loaded.loadFormat('tex')).module;
    for (const name of ['compile', 'engines', 'template', 'escape', 'clearAux']) {
        assert.equal(typeof mod.api?.[name], 'function', `tex.js 必须导出 api.${name} 工厂`);
    }

    const result = await executeRun({
        script: `
const t = office.tex.create({
  path: 'thesis/thesis.tex', title: 'SDK 冒烟论文', degree: 'master',
  abstract: { zh: ['摘要。'], keywordsZh: ['冒烟'] },
  chapters: [{ title: '绪论', sections: [{ title: '背景', blocks: ['正文。'] }] }],
  acknowledgements: ['致谢。'],
});
const report = office.tex.read(t.main);
const engines = await office.tex.engines();
const esc = office.tex.escape('100% & a_1');
return { main: t.main, chapters: report.stats.chapters, kind: report.kind, esc, engines: engines.available.length, files: t.files.length };
`,
        purpose: 'office.tex 冒烟',
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.main, 'thesis/thesis.tex');
    assert.equal(result.returned.chapters, 1);
    assert.equal(result.returned.kind, 'main');
    assert.equal(result.returned.esc, '100\\% \\& a\\_1');
    assert.equal(typeof result.returned.engines, 'number');
    // 写出的 .tex 会被自动复检：判断来自注册表，不是写死三件套。
    assert.ok(result.files.some((file) => file.format === 'tex' && file.path.endsWith('thesis.tex')),
        '写出的 .tex 应当进 files 并被复检');
    assert.ok(result.files.every((file) => file.stats !== null || file.format === null), '复检报告要么有 stats 要么不是注册格式');
    // 产物在复检之后才删：复检发生在脚本结束之后，脚本里自己删掉会被记成「复检失败」。
    rmSync(join(root, 'thesis'), { recursive: true, force: true });
});

await check('缓存：默认保留中间文件（跨调用复用），keepCache:false 才清空', async () => {
    const result = await executeRun({
        script: `office.cache.write('tmp.json', {a:1}); return office.cache.list().filter((f) => !f.dir).length;`,
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned, 1);
    assert.equal(result.cache.clearedAfter, false);
    assert.equal(result.cache.kept, 1);
    assert.equal(existsSync(join(root, '.office/cache', 'tmp.json')), true);
    assert.match(renderRun(result), /保留 1 个中间文件/);

    const cleared = await executeRun({
        script: `return office.cache.list().filter((f) => !f.dir).length;`,
        keepCache: false,
    }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(cleared.ok, true, cleared.error?.message);
    assert.equal(cleared.returned, 1, '清空前脚本仍然看得见上一次的缓存');
    assert.equal(cleared.cache.clearedAfter, true);
    assert.equal(existsSync(join(root, '.office/cache')), false);
    assert.match(renderRun(cleared), /已清空/);
});

await check('工作目录里没有残留垃圾', () => {
    const entries = readdirSync(root).map((name) => name);
    // .office 是插件自己的隐藏目录（缓存 / 检索 / 三层记忆 / 测试中间产物都在里面），
    // 属于「插件自己该有的东西」，不是垃圾；工作目录根上不该再出现第二个插件目录。
    assert.deepEqual(
        entries.sort(),
        ['.office', 'notes.md', 'out', 'smoke-deck.pptx'].filter((n) => existsSync(join(root, n))).sort(),
    );
});

// ── office_run 的失败必须让宿主看见（第十八轮 P0-1，第二十七轮落地） ────────────

await check('office_run 脚本失败：守卫把它换成宿主认识的 isError 结果', async () => {
    const failedValue = {
        ok: false,
        files: [],
        otherFiles: [],
        returned: null,
        logs: [],
        notes: [],
        warnings: ['这次调用没有写出任何文件。'],
        error: { name: 'TypeError', message: 'w.titel is not a function', line: 3, column: 9, stack: [] },
        cache: { dir: '.office/cache', kept: 0, keptBytes: 0, hits: 0, pruned: [], ttlMinutes: 720 },
        elapsedMs: 5,
    };

    // 1) 纯函数：只把 ok:false 认成失败，别的值一律返回 null（别动它）。
    const failure = officeRunErrorResult(failedValue);
    assert.equal(failure.isError, true);
    assert.equal(failure.error.message, 'w.titel is not a function');
    assert.equal(failure.error.info.code, 'OFFICE_RUN_FAILED', '给宿主一个稳定的 code 便于路由与复现');
    assert.equal(failure.error.info.name, 'TypeError');
    assert.match(failure.error.info.reason, /第 3 行/);
    assert.match(failure.content[0].text, /office_run 执行失败/);
    assert.match(failure.content[0].text, /脚本第 3 行/, '失败反馈仍要带脚本行号');
    assert.match(failure.content[0].text, /没有写出任何文件/, '告警不能被丢掉');
    assert.equal(officeRunErrorResult({ ...failedValue, ok: true }), null, '成功的结果不许动');
    assert.equal(officeRunErrorResult(undefined), null);
    assert.equal(officeRunErrorResult('ok'), null);

    // 2) 接线：插件启动时把守卫挂在宿主的 tools/execute 环绕点上。
    const { apply } = await import('../src/index.js');
    const handlers = new Map();
    const ctx = {
        tools: { register: () => () => {} },
        get: () => undefined,
        inject: () => {},
        on: (event, listener) => { handlers.set(event, listener); return () => {}; },
        logger: {},
    };
    apply(ctx, {});
    const guard = handlers.get('tools/execute');
    assert.equal(typeof guard, 'function', '要把守卫挂在 tools/execute 上（与宿主自带的 timeout 策略同一个点）');

    const success = { isError: false, value: failedValue, content: [{ type: 'text', text: '❌ 工具体已经渲染好的失败反馈' }] };
    const guarded = await guard({ name: 'office_run' }, async () => success);
    assert.equal(guarded.isError, true);
    assert.equal(guarded.content[0].text, '❌ 工具体已经渲染好的失败反馈', '要保留已经渲染好的那段文本（里面有文件清单与告警）');
    assert.equal(guarded.error.info.code, 'OFFICE_RUN_FAILED');

    // 宿主拿到这个结果之后会做什么，是本条断言真正依赖的契约 —— 照抄
    // `dsh-tools` 的 `normalizeDispatchResult` / `materializeFinalResult` 的 isError 分支：
    // 只保留 content / error / meta / additionalContexts，**丢掉 value**。
    // 这一步是仿真（没有起真调度器），但它把「哪些字段能活下来」写成了可读的断言：
    // 宿主那两段一旦改成重渲染 content，这里就该跟着改，而不是继续假绿。
    const hostNormalize = (result) => (result.isError === true
        ? {
            isError: true,
            error: result.error,
            content: result.content,
            ...(result.meta !== undefined ? { meta: result.meta } : {}),
            ...(result.additionalContexts !== undefined ? { additionalContexts: result.additionalContexts } : {}),
        }
        : { isError: false, value: result.value, content: result.content });
    const recorded = hostNormalize(guarded);
    assert.equal(recorded.isError, true, '会话日志里要记成失败');
    assert.equal(recorded.error.info.code, 'OFFICE_RUN_FAILED');
    assert.equal(recorded.error.info.name, 'TypeError');
    assert.match(recorded.error.info.reason, /第 3 行/);
    assert.equal(recorded.content[0].text, '❌ 工具体已经渲染好的失败反馈');
    assert.equal(Object.hasOwn(recorded, 'value'), false, 'isError 结果没有 value —— 宿主的契约如此');

    // meta / additionalContexts 要跟着带过去（宿主只保留这两样展示元数据）
    const withMeta = await guard({ name: 'office_run' }, async () => ({ ...success, meta: { surface: 'x' }, additionalContexts: [{ role: 'user' }] }));
    assert.deepEqual(withMeta.meta, { surface: 'x' }, 'meta 不能丢');
    assert.deepEqual(withMeta.additionalContexts, [{ role: 'user' }], 'additionalContexts 不能丢');

    const okResult = { isError: false, value: { ...failedValue, ok: true }, content: [{ type: 'text', text: '✅ 完成' }] };
    assert.equal((await guard({ name: 'office_run' }, async () => okResult)).isError, false, '成功的结果照旧成功');
    const other = { isError: false, value: { ok: false }, content: [] };
    assert.equal(await guard({ name: 'office_search_run' }, async () => other), other, '别的工具原样透传');
    const hostError = { isError: true, error: { message: 'boom' }, content: [{ type: 'text', text: 'Error: boom' }] };
    assert.equal(await guard({ name: 'office_run' }, async () => hostError), hostError, '宿主已经标成失败的结果不许再包一层');
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) {
    for (const item of failed) console.error(item.error);
}
console.log(`smoke: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
