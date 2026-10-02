/**
 * 渲染预览（office.preview，历史遗留 `2-1`）的测试。
 *
 * 这条账说的是「办公会话里没有渲染预览入口」：能生成三件套、也能读回结构，
 * 但看不到画面。本轮的落点是拿宿主随部署安装的 LibreOffice 引擎把自家文档
 * 渲染成页面图（交给 read_image）或 PDF。
 *
 * 分三层：
 *   1. 能力探针与诚实降级：引擎在/不在两种情形都要给出可执行的结论；
 *   2. 渲染与缓存：docx / pptx / xlsx 渲染出真 PNG（魔数、尺寸、页码），
 *      同一份文件同样参数第二次是命中（不重渲染），dpi / pages 换参数就换键；
 *   3. 经 office_run 的全链路：脚本里 office.preview.* 可用、PDF 交付物出现在
 *      反馈的 otherFiles 里、参数错误给的是可执行的错。
 *
 * 引擎解析不到时（非 DSH 环境）整段真渲染 SKIP 并如实说明，不当成通过。
 *
 * 跑法：node test/preview.mjs
 */
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import { loadPreviewKit, PREVIEW_DEFAULT_DPI, previewCheck, resetPreviewKit } from '../src/preview.js';
import { executeRun } from '../src/run.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        if (error?.skip === true) results.push({ name, ok: true, note: `SKIP ${error.message}` });
        else results.push({ name, ok: false, error });
    }
}
function skip(message) {
    const error = new Error(message);
    error.skip = true;
    throw error;
}

const root = mkdtempSync(join(tmpdir(), 'office-preview-'));
const runScript = (script, extra = {}) => executeRun({ script, purpose: '预览测试' }, { agent: { session: { header: { cwd: root } } } }, resolveConfig(extra));

const probe = await previewCheck();
const ready = probe.available === true;

/** PNG 魔数：渲染出来的得是真图，不是空文件。 */
function isPng(path) {
    const bytes = readFileSync(path);
    return bytes.length > 100 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;
}
const isPdf = (path) => readFileSync(path).subarray(0, 5).toString('latin1') === '%PDF-';

// ── 1. 能力探针与诚实降级 ──────────────────────────────────────────────────

await check('preview.check：报出引擎、后端、版本与支持的输入格式', async () => {
    assert.equal(typeof probe.available, 'boolean');
    assert.equal(probe.engine, '@deepseek-ai/libreoffice-kit');
    assert.ok(Array.isArray(probe.formats) && probe.formats.includes('.docx') && probe.formats.includes('.pptx'));
    if (!ready) return `不可用：${probe.hint}`;
    assert.ok(['native', 'wasm'].includes(probe.backend), `backend 异常：${probe.backend}`);
    assert.ok(typeof probe.version === 'string' && probe.version !== '');
    return `${probe.backend} ${probe.version}`;
});

await check('诚实降级：引擎解析不到时 available:false，并说清缺什么（不当成可用）', async () => {
    const saved = {
        DSH_PROFILE_DIR: process.env.DSH_PROFILE_DIR,
        DSH_HOST_ROOT: process.env.DSH_HOST_ROOT,
        DSH_CHECKOUT: process.env.DSH_CHECKOUT,
    };
    const cwd = process.cwd();
    try {
        delete process.env.DSH_PROFILE_DIR;
        delete process.env.DSH_HOST_ROOT;
        delete process.env.DSH_CHECKOUT;
        process.chdir(tmpdir());
        resetPreviewKit();
        const missing = await loadPreviewKit();
        if (missing.module !== undefined) {
            // 本机某些锚点仍然解析得到（例如插件被 link 到带 node_modules 的 profile 里）——
            // 这不是失败，如实说明这条负例这次测不了。
            return 'SKIP 本机仍有锚点能解析到引擎，负例未覆盖';
        }
        assert.match(missing.error, /找不到渲染引擎|加载失败/);
        const value = await previewCheck();
        assert.equal(value.available, false);
        assert.match(value.hint, /找不到渲染引擎/);
        assert.ok(value.formats.includes('.docx'), '不可用时也要给出支持的格式清单');
    } finally {
        for (const [key, value] of Object.entries(saved)) {
            if (value === undefined) delete process.env[key];
            else process.env[key] = value;
        }
        process.chdir(cwd);
        resetPreviewKit();
    }
    return '缺引擎时的结论可执行';
});

// ── 2. 渲染与缓存 ──────────────────────────────────────────────────────────

if (!ready) {
    results.push({ name: '真渲染（docx / pptx / xlsx → 页面图）', ok: true, note: `SKIP ${probe.hint}` });
} else {
    await check('造产物：一份 docx + 一份两页 pptx + 一份 xlsx', async () => {
        const result = await runScript(`
            const wb = office.word.create({ path: '报告.docx', title: '预览测试' });
            wb.heading('一、结论', 1);
            wb.para('这一段用来验证渲染预览：写出来 → 渲染成页面图 → read_image。');
            wb.bullets(['第一点', '第二点', '第三点']);
            wb.table({ columns: ['项', '值'], rows: [['A', 1], ['B', 2]] });
            wb.save();
            const deck = office.ppt.create({ path: '汇报.pptx', title: '预览测试' });
            deck.cover({ title: '预览测试', subtitle: '渲染成页面图' });
            deck.bullets({ title: '要点', items: ['第一点', '第二点'] });
            deck.save();
            const sheet = office.excel.create({ path: '明细.xlsx', title: '预览测试' });
            const s = sheet.sheet('明细');
            s.table({ columns: ['项', '值'], rows: [['A', 1], ['B', 2]], totalRow: true });
            sheet.save();
            return office.files.list('.');
        `);
        assert.equal(result.ok, true, result.error?.message);
        for (const name of ['报告.docx', '汇报.pptx', '明细.xlsx']) {
            assert.ok(existsSync(join(root, name)), `${name} 没写出来`);
        }
        return '3 份产物';
    });

    await check('office.preview.render：docx 渲染出真 PNG，路径可直接交给 read_image', async () => {
        const result = await runScript(`return await office.preview.render('报告.docx');`);
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.equal(value.ok, true);
        assert.equal(value.kind, 'images');
        assert.equal(value.dpi, PREVIEW_DEFAULT_DPI);
        assert.ok(value.count >= 1, '至少渲染出一页');
        assert.equal(value.reused, false);
        for (const image of value.images) {
            assert.ok(image.width > 100 && image.height > 100, `尺寸异常：${JSON.stringify(image)}`);
            assert.ok(image.bytes > 1000, `图太小：${image.bytes}`);
            assert.ok(!image.path.startsWith('.office/cache') === false, '预览图应当落在缓存目录');
            const absolute = join(root, image.path);
            assert.ok(existsSync(absolute), `图不在盘上：${image.path}`);
            assert.equal(statSync(absolute).size, image.bytes);
            assert.ok(isPng(absolute), `${image.path} 不是 PNG`);
        }
        assert.deepEqual(value.missingFonts, []);
        return `${value.count} 张 / ${value.images[0].width}x${value.images[0].height} / ${value.images[0].bytes}B`;
    });

    await check('缓存：同一份文件同样参数第二次是命中（不重渲染，路径一致）', async () => {
        const first = await runScript(`return await office.preview.render('报告.docx');`);
        const second = await runScript(`return await office.preview.render('报告.docx');`);
        assert.equal(second.ok, true, second.error?.message);
        const a = first.returned;
        const b = second.returned;
        assert.equal(b.reused, true, '第二次应当命中');
        assert.equal(b.dir, a.dir);
        assert.deepEqual(b.images.map((item) => item.path), a.images.map((item) => item.path));
        // 命中的返回形状要和新渲染一致（宽高从 IHDR 现读、清单还原元数据）—— 否则按
        // width / pageCount / missingFonts 读的调用方会在第二次调用时拿到 undefined。
        assert.deepEqual(b.images.map((item) => [item.width, item.height]), a.images.map((item) => [item.width, item.height]));
        assert.deepEqual(b.images.map((item) => item.bytes), a.images.map((item) => item.bytes));
        for (const key of ['pageCount', 'backend', 'rasterEngine']) {
            assert.deepEqual(b[key], a[key], `命中的 ${key} 要与新渲染一致`);
        }
        assert.deepEqual(b.missingFonts, a.missingFonts);
        // 命中要记进缓存统计（不然「缓存省了多少」这件事就没有证据）。
        assert.ok(second.cache.hits >= 1, `命中数没记上：${second.cache.hits}`);
        return `命中 ${second.cache.hits} 次`;
    });

    await check('长文档：页数超过 maxPages 时只渲前 maxPages 页并如实报截断（不是直接失败）', async () => {
        const made = await runScript(`
            const deck = office.ppt.create({ path: '长稿.pptx', title: '长稿' });
            deck.cover({ title: '长稿' });
            for (let i = 1; i <= 14; i += 1) deck.bullets({ title: '第 ' + i + ' 页', items: ['a', 'b'] });
            deck.save();
            return office.ppt.read('长稿.pptx').stats.slides;
        `);
        assert.equal(made.ok, true, made.error?.message);
        assert.ok(made.returned > 12, `夹具要超过 maxPages=12，实际 ${made.returned} 页`);
        // 默认调用（pages:'all' + maxPages 12）：引擎会抛 output-too-large，插件要降级成前 12 页
        const result = await runScript(`return await office.preview.render('长稿.pptx');`);
        assert.equal(result.ok, true, `默认调用不该失败：${result.error?.message}`);
        const value = result.returned;
        assert.equal(value.count, 12, `应当渲前 12 页，实际 ${value.count}`);
        assert.equal(value.pageCount, made.returned);
        assert.match(value.hint, /这次只渲染了 12 张（maxPages=12）/);
        assert.ok(value.images.every((item) => isPng(join(root, item.path))));
        // 命中路径也要把「被 maxPages 截断」说清楚
        const again = await runScript(`return await office.preview.render('长稿.pptx');`);
        assert.equal(again.returned.reused, true);
        assert.equal(again.returned.pageCount, made.returned);
        assert.match(again.returned.hint, /maxPages=12/);
        // 放大 maxPages 就能多渲几页
        const bigger = await runScript(`return await office.preview.render('长稿.pptx', { maxPages: 15 });`);
        assert.equal(bigger.returned.count, made.returned, 'maxPages 放大到 15 应当渲全');
        assert.ok(!/只渲染了/.test(bigger.returned.hint));
        return `${made.returned} 页 → 渲 12 张`;
    });

    await check('参数：pages 只渲染指定页，dpi 改分辨率就换缓存键', async () => {
        const page2 = await runScript(`return await office.preview.render('汇报.pptx', { pages: [2] });`);
        assert.equal(page2.ok, true, page2.error?.message);
        assert.equal(page2.returned.count, 1, '只要第 2 页就只渲染一张');
        assert.equal(page2.returned.images[0].page, 2);
        const low = await runScript(`return await office.preview.render('汇报.pptx', { dpi: 72 });`);
        const high = await runScript(`return await office.preview.render('汇报.pptx', { dpi: 160 });`);
        assert.equal(low.ok, true, low.error?.message);
        assert.equal(high.ok, true, high.error?.message);
        assert.notEqual(low.returned.dir, high.returned.dir, 'dpi 不同应当是不同的缓存键');
        assert.ok(high.returned.images[0].width > low.returned.images[0].width * 1.5,
            `dpi 没生效：${low.returned.images[0].width} → ${high.returned.images[0].width}`);
        // 宽高比不随 dpi 变（像素取整会带来千分位误差，给 1% 容差）。
        const ratio = (image) => image.height / image.width;
        assert.ok(Math.abs(ratio(high.returned.images[0]) - ratio(low.returned.images[0])) < 0.01,
            `同一页的宽高比不该随 dpi 变：${ratio(low.returned.images[0])} → ${ratio(high.returned.images[0])}`);
        return `${low.returned.images[0].width}px → ${high.returned.images[0].width}px`;
    });

    await check('office.preview.render：xlsx 按工作表渲染数据区', async () => {
        const result = await runScript(`return await office.preview.render('明细.xlsx', { sheet: '明细' });`);
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.ok(value.count >= 1, '工作表应当渲染出至少一张');
        assert.ok(value.images.every((item) => isPng(join(root, item.path))));
        return `${value.count} 张`;
    });

    await check('office.preview.pdf：to 写到工作目录（交付物）并出现在 otherFiles 里', async () => {
        const result = await runScript(`return await office.preview.pdf('报告.docx', { to: '报告.pdf' });`);
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.equal(value.kind, 'pdf');
        assert.equal(value.path, '报告.pdf');
        assert.ok(isPdf(join(root, '报告.pdf')), '写出来的不是 PDF');
        assert.equal(statSync(join(root, '报告.pdf')).size, value.bytes);
        assert.ok(result.otherFiles.some((file) => file.path === '报告.pdf'),
            `PDF 没出现在 otherFiles 里：${JSON.stringify(result.otherFiles)}`);
        // 再转一次：内容一致，不重写、也不重复登记。
        const again = await runScript(`return await office.preview.pdf('报告.docx', { to: '报告.pdf' });`);
        assert.equal(again.ok, true, again.error?.message);
        assert.equal(again.returned.reused, true);
        assert.equal(again.files.length, 0, '命中时不该再有写出的文件');
        // 不传 to 只放缓存。
        const cachedOnly = await runScript(`return await office.preview.pdf('报告.docx');`);
        assert.equal(cachedOnly.ok, true, cachedOnly.error?.message);
        assert.ok(cachedOnly.returned.path.startsWith('.office/cache/preview/'), cachedOnly.returned.path);
        return `${value.bytes}B`;
    });

    await check('office.preview.pdf：不许覆盖源文件、不许静默覆盖别人的文件、不许写出工作目录', async () => {
        const sourceBefore = readFileSync(join(root, '报告.docx'));
        const self = await runScript(`return await office.preview.pdf('报告.docx', { to: '报告.docx' });`);
        assert.equal(self.ok, false, '把 PDF 写到源文件上必须被拦住');
        assert.match(self.error.message, /不能和输入文件是同一个/);
        assert.ok(sourceBefore.equals(readFileSync(join(root, '报告.docx'))), '源文件必须一个字节都没变');

        // 既有的、内容不同的文件：默认不覆盖
        await runScript(`office.files.write('已有.pdf', 'not a pdf'); return true;`);
        const clash = await runScript(`return await office.preview.pdf('报告.docx', { to: '已有.pdf' });`);
        assert.equal(clash.ok, false, '内容不同的既有文件默认不该被覆盖');
        assert.match(clash.error.message, /已存在且内容不同/);
        assert.equal(readFileSync(join(root, '已有.pdf'), 'utf8'), 'not a pdf');
        const forced = await runScript(`return await office.preview.pdf('报告.docx', { to: '已有.pdf', overwrite: true });`);
        assert.equal(forced.ok, true, forced.error?.message);
        assert.equal(forced.returned.overwrote, true);
        assert.ok(isPdf(join(root, '已有.pdf')));

        // 相对路径不许写出工作目录
        const escape = await runScript(`return await office.preview.pdf('报告.docx', { to: '../逃出去.pdf' });`);
        assert.equal(escape.ok, false);
        assert.match(escape.error.message, /工作目录之外/);
        return '3 类拦截';
    });

    await check('参数错误给的是可执行的错（不是引擎的原始报错）', async () => {
        const badExt = await runScript(`return await office.preview.render('报告.docx', { pages: [0] });`);
        assert.equal(badExt.ok, false);
        assert.match(badExt.error.message, /pages 只能给/);
        const both = await runScript(`return await office.preview.render('明细.xlsx', { pages: [1], sheet: '明细' });`);
        assert.equal(both.ok, false);
        assert.match(both.error.message, /不能同时给/);
        const missing = await runScript(`return await office.preview.render('没有这份.docx');`);
        assert.equal(missing.ok, false);
        assert.match(missing.error.message, /找不到文件/);
        // 不支持的扩展名要把支持清单列出来。
        await runScript(`office.files.write('笔记.md', '# 不是文档'); return true;`);
        const unsupported = await runScript(`return await office.preview.render('笔记.md');`);
        assert.equal(unsupported.ok, false);
        assert.match(unsupported.error.message, /不在支持的清单里/);
        assert.match(unsupported.error.message, /\.docx/);
        return '4 类参数错误';
    });
}

await check('office_help：preview 话题在（默认层给签名，全文层给细节）', async () => {
    const { buildHelp } = await import('../src/docs.js');
    const brief = await buildHelp('preview');
    for (const needle of ['office.preview.check', 'office.preview.render', 'office.preview.pdf', 'read_image', 'dpi', 'pages', 'sheet']) {
        assert.ok(brief.text.includes(needle), `preview 默认层缺 ${needle}`);
    }
    const full = await buildHelp('preview', { detail: true });
    assert.match(full.text, /LibreOffice/);
    assert.match(full.text, /missingFonts/);
    const index = await buildHelp('');
    assert.match(index.text, /preview：把自家 docx/);
    const run = await buildHelp('run');
    assert.match(run.text, /office\.preview/, 'run 默认层要提到预览入口');
    return '默认层 + 全文层 + 索引';
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`preview: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
