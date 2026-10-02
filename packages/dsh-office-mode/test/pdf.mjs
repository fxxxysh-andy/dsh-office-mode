/**
 * PDF 与缓存生命周期的测试。
 *
 * 两件事在这里钉住：
 *   1. PDF 的两条读法 —— 文本型抽文字、图像型（没有文本层）渲染成图片；
 *      渲染结果按内容键缓存，第二次渲染必须**命中**而不是重算。
 *   2. 缓存跨调用保留：这是「每次 office_run 都把缓存清掉」的回归测试。
 *
 * 渲染引擎是外部依赖（PyMuPDF / poppler / MuPDF / Ghostscript），本机没有时
 * 相关断言标记 SKIP 并如实说明 —— 不当成通过。
 *
 * 跑法：node test/pdf.mjs
 */
import { mkdtempSync, mkdirSync, rmSync, existsSync, readFileSync, statSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { createCache } from '../src/engine/cache.js';
import { createEnv } from '../src/engine/kit.js';
import { pdfEngines, pdfInfo, pdfPages, pdfText, probePdfEngines } from '../src/pdf.js';
import { defaultPdfReader } from '../src/web.js';
import { resolveConfig } from '../src/config.js';
import { executeRun } from '../src/run.js';
import { renderRun } from '../src/tools.js';

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

/** 造一份最小但结构完整的 PDF（含 xref），用来测解析与渲染。 */
function makePdf(options = {}) {
    const pages = options.pages ?? [['Page One Hello'], ['Page Two World']];
    const withFonts = options.withFonts !== false;
    const objects = [];
    const pageCount = pages.length;
    // 1 catalog / 2 pages / 3..(2+n) page / 之后是资源与内容流
    const fontId = 3 + pageCount;
    const imageId = withFonts ? fontId : fontId;
    const contentStart = withFonts ? fontId + 1 : fontId + 1;
    const kids = pages.map((_all, index) => `${3 + index} 0 R`).join(' ');
    objects.push('<< /Type /Catalog /Pages 2 0 R >>');
    objects.push(`<< /Type /Pages /Kids [${kids}] /Count ${pageCount} >>`);
    for (let index = 0; index < pageCount; index += 1) {
        const contentId = contentStart + index;
        const resource = withFonts
            ? `/Font << /F1 ${fontId} 0 R >>`
            : `/XObject << /Im1 ${imageId} 0 R >>`;
        objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << ${resource} >> /Contents ${contentId} 0 R >>`);
    }
    if (withFonts) objects.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    else objects.push('<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 12 >>\nstream\n\xff\x00\x00\x00\xff\x00\x00\x00\xff\xff\xff\xff\nendstream');
    for (let index = 0; index < pageCount; index += 1) {
        const body = withFonts
            ? `BT /F1 24 Tf 72 700 Td (${pages[index][0]}) Tj ET`
            : 'q 200 0 0 200 100 500 cm /Im1 Do Q';
        objects.push(`<< /Length ${body.length} >>\nstream\n${body}\nendstream`);
    }
    const infoId = contentStart + pageCount;
    objects.push(`<< /Title <FEFF6D4B8BD5> /Producer (office-mode-test) >>`);

    let pdf = '%PDF-1.4\n';
    const offsets = [];
    objects.forEach((body, index) => {
        offsets.push(pdf.length);
        pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
    });
    const xrefStart = pdf.length;
    pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (const offset of offsets) pdf += `${String(offset).padStart(10, '0')} 00000 n \n`;
    pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R /Info ${infoId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`;
    return Buffer.from(pdf, 'latin1');
}

const root = mkdtempSync(join(tmpdir(), 'office-pdf-'));
const env = createEnv({ root });
const cache = createCache({ root, dir: '.office/cache' });
const config = resolveConfig({});

const textPdf = join(root, 'text.pdf');
const imagePdf = join(root, 'scan.pdf');
writeFileSync(textPdf, makePdf({ pages: [['Page One Hello'], ['Page Two World']] }));
writeFileSync(imagePdf, makePdf({ pages: [['ignored'], ['ignored']], withFonts: false }));

await check('pdf：轻量解析能读出页数与元信息（不依赖外部工具）', async () => {
    const info = await pdfInfo(textPdf, env, { cache });
    assert.equal(info.pages, 2);
    assert.equal(info.bytes, statSync(textPdf).size);
    assert.equal(info.title, '测试');
    assert.equal(info.producer, 'office-mode-test');
    assert.equal(info.hasTextLayer, true);
    assert.equal(info.pageSizeCm.width, 20.99);
    assert.equal(info.pageSizeCm.height, 29.7);
});

await check('pdf：图像型（没有字体）被判成没有文本层', async () => {
    const info = await pdfInfo(imagePdf, env, { cache });
    assert.equal(info.pages, 2);
    assert.equal(info.hasTextLayer, false);
    assert.equal(info.counts.images > 0, true);
    assert.match(info.hint, /渲染成图片/);
});

await check('pdf：不是 PDF 的文件会被明确拒绝', async () => {
    writeFileSync(join(root, 'fake.pdf'), 'hello', 'utf8');
    await assert.rejects(() => pdfInfo(join(root, 'fake.pdf'), env, { cache }), /不是 PDF/);
});

await check('pdf：抽文本（文本型）', async () => {
    const engines = await pdfEngines();
    if (!engines.text.includes('pdftotext') && !engines.text.includes('fitz')) skip('本机没有抽文本引擎');
    const value = await pdfText(textPdf, { out: 'raw.md' }, env, cache);
    assert.equal(value.chars > 0, true, `chars=${value.chars}`);
    assert.match(readFileSync(join(root, 'raw.md'), 'utf8'), /Page One Hello/);
    assert.equal(value.path, 'raw.md');
});

await check('pdf：取正文用的那个读取器在真实引擎上跑得通（第二十九轮）', async () => {
    // 这条守的是 web.js 里 defaultPdfReader 的 env 垫片：pdfText 除了 env.resolve 还会用
    // env.root（displayPath 要它），少一个就在「真的抽一份 PDF」时炸成 TypeError ——
    // 注入假读取器的单测看不出这一点，所以这里用真文件、真引擎跑一次。
    const engines = await pdfEngines();
    if (!engines.text.includes('pdftotext') && !engines.text.includes('fitz')) skip('本机没有抽文本引擎');
    const value = await defaultPdfReader.text(textPdf, { maxPages: 1 });
    assert.equal(value.extractor === 'pdftotext' || value.extractor === 'fitz', true, `engine=${value.extractor}`);
    assert.match(value.text, /Page One Hello/);
    assert.equal(value.pages, 1, 'maxPages 传 1 时只抽一页');
    assert.equal(typeof value.hint, 'string');
});

await check('pdf：渲染页面并复用（缓存命中）', async () => {
    const engines = await pdfEngines();
    if (engines.render.length === 0) skip('本机没有 PDF 渲染引擎');
    const first = await pdfPages(textPdf, { dpi: 72 }, env, cache, config);
    assert.equal(first.files.length, 2);
    assert.equal(first.rendered, 2);
    assert.equal(first.reused, 0);
    for (const file of first.files) {
        assert.equal(existsSync(env.resolve(file.path)), true, file.path);
        assert.equal(file.bytes > 0, true, file.path);
    }
    const before = cache.stats.hits;
    const second = await pdfPages(textPdf, { dpi: 72 }, env, cache, config);
    assert.equal(second.rendered, 0, '第二次不该重算');
    assert.equal(second.reused, 2, '第二次应当全部命中');
    assert.equal(cache.stats.hits, before + 2);
    return `引擎 ${first.engine}，${first.bytes} 字节`;
});

await check('pdf：图像型 PDF 也能渲染（手写笔记那条路）', async () => {
    const engines = await pdfEngines();
    if (engines.render.length === 0) skip('本机没有 PDF 渲染引擎');
    const value = await pdfPages(imagePdf, { dpi: 72, format: 'jpeg' }, env, cache, config);
    assert.equal(value.files.length, 2);
    assert.equal(value.format, 'jpeg');
    for (const file of value.files) assert.equal(/\.jpg$/.test(file.path), true, file.path);
});

await check('pdf：页数超过单次上限时报错并给出分批建议', async () => {
    const engines = await pdfEngines();
    if (engines.render.length === 0) skip('本机没有 PDF 渲染引擎');
    await assert.rejects(
        () => pdfPages(textPdf, { from: 1, to: 2 }, env, cache, { ...config, pdfMaxPages: 1 }),
        /一次最多渲染 1 页/,
    );
});

await check('pdf：指定不存在的引擎时明确报错', async () => {
    const engines = await pdfEngines();
    if (engines.render.includes('mutool')) skip('本机装了 mutool，换个不存在的引擎名测');
    await assert.rejects(
        () => pdfPages(textPdf, {}, env, cache, { ...config, pdfEngine: 'mutool' }),
        /不可用/,
    );
});

await check('pdf：引擎探测结果可读（给报错与设置页用）', async () => {
    const engines = await pdfEngines();
    assert.equal(Array.isArray(engines.render), true);
    assert.equal(engines.info.includes('light'), true);
    const probe = probePdfEngines();
    assert.equal(typeof probe.bin, 'object');
    return `render=[${engines.render.join(',')}] text=[${engines.text.join(',')}]`;
});

await check('缓存：子目录路径可用，向上跳会被拒', () => {
    const written = cache.write('pdf/a/b.png', Buffer.from([1, 2, 3]));
    assert.equal(written.bytes, 3);
    assert.equal(existsSync(cache.locate('pdf/a/b.png')), true);
    assert.equal(cache.list().some((entry) => entry.name === 'pdf/a/b.png'), true);
    assert.throws(() => cache.path('../escape.txt'), /跳出缓存目录/);
});

await check('缓存：TTL 只清过期条目，新文件不动', () => {
    const fresh = cache.write('fresh.txt', 'x');
    const stale = cache.write('stale.txt', 'y');
    const old = new Date(Date.now() - 48 * 3600 * 1000);
    utimesSync(cache.locate('stale.txt'), old, old);
    const pruned = cache.prune({ maxAgeMs: 3600 * 1000 });
    assert.equal(pruned.removed.includes('stale.txt'), true);
    assert.equal(pruned.removed.includes('fresh.txt'), false);
    assert.equal(existsSync(cache.locate('stale.txt')), false);
    assert.equal(existsSync(fresh.absolute), true);
    assert.equal(stale.bytes, 1);
});

await check('缓存：容量上限从最旧的开始删', () => {
    const dir = mkdtempSync(join(tmpdir(), 'office-tmp-cap-'));
    const small = createCache({ root: dir, dir: '.office/cache' });
    small.write('a.bin', Buffer.alloc(40));
    small.write('b.bin', Buffer.alloc(40));
    const old = new Date(Date.now() - 3600 * 1000);
    utimesSync(small.locate('a.bin'), old, old);
    const pruned = small.prune({ maxBytes: 60 });
    assert.equal(pruned.removed.includes('a.bin'), true, JSON.stringify(pruned));
    assert.equal(small.list().filter((entry) => !entry.dir).length, 1);
    rmSync(dir, { recursive: true, force: true });
});

await check('office_run：缓存跨调用保留（这是缓存命中的前提）', async () => {
    const runRoot = mkdtempSync(join(tmpdir(), 'office-tmp-run-'));
    const first = await executeRun({
        script: `office.cache.write('keep.json', {a:1}); return office.cache.list().length;`,
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    assert.equal(first.ok, true, first.error?.message);
    assert.equal(first.cache.clearedAfter, false, '默认不该清空');
    assert.equal(existsSync(join(runRoot, '.office', 'cache', 'keep.json')), true, '第一次写下的中间文件必须还在');

    const second = await executeRun({
        script: `const files = office.cache.list().filter((f) => !f.dir).map((f) => f.name);
            office.assert(files.includes('keep.json'), '第二次调用应当看得见上一次的缓存');
            return { files, stats: office.cache.stats() };`,
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    assert.equal(second.ok, true, second.error?.message);
    assert.equal(second.cache.kept > 0, true);

    const cleared = await executeRun({
        script: `office.cache.write('gone.json', {a:1}); return office.cache.list().length;`,
        keepCache: false,
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    assert.equal(cleared.cache.clearedAfter, true);
    assert.equal(existsSync(join(runRoot, '.office', 'cache')), false, 'keepCache:false 才清空');

    const text = renderRun(second);
    assert.match(text, /缓存：\.office\/cache 保留/);
    rmSync(runRoot, { recursive: true, force: true });
});

await check('office_run：过期缓存会在下次调用开始时清掉', async () => {
    const runRoot = mkdtempSync(join(tmpdir(), 'office-tmp-ttl-'));
    const first = await executeRun({
        script: `office.cache.write('old.json', {a:1}); return 1;`,
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    assert.equal(first.ok, true, first.error?.message);
    const stale = join(runRoot, '.office', 'cache', 'old.json');
    const old = new Date(Date.now() - 48 * 3600 * 1000);
    utimesSync(stale, old, old);

    const second = await executeRun({
        script: `return office.cache.list().filter((f) => !f.dir).map((f) => f.name);`,
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({ cacheTtlMinutes: 60 }));
    assert.equal(second.ok, true, second.error?.message);
    assert.deepEqual(second.returned, []);
    assert.equal(second.cache.pruned.includes('old.json'), true, JSON.stringify(second.cache));
    rmSync(runRoot, { recursive: true, force: true });
});

await check('office_run：office.pdf 挂在 SDK 上（面不能静默丢）', async () => {
    const runRoot = mkdtempSync(join(tmpdir(), 'office-pdf-sdk-'));
    writeFileSync(join(runRoot, 'doc.pdf'), makePdf({ pages: [['SDK Probe']] }));
    const result = await executeRun({
        script: `const info = await office.pdf.info('doc.pdf');
            const engines = await office.pdf.engines();
            return { pages: info.pages, hasTextLayer: info.hasTextLayer, render: engines.render.length };`,
    }, { agent: { session: { header: { cwd: runRoot } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.pages, 1);
    assert.equal(result.returned.hasTextLayer, true);
    assert.equal(typeof result.returned.render, 'number');
    rmSync(runRoot, { recursive: true, force: true });
});

await check('office_help：pdf 话题存在且提到 read_image', async () => {
    const { buildHelp } = await import('../src/docs.js');
    const pdf = await buildHelp('pdf');
    assert.match(pdf.text, /office\.pdf\.pages/);
    assert.match(pdf.text, /read_image/);
    const index = await buildHelp('');
    assert.match(index.text, /pdf：PDF/);
    const cache = await buildHelp('cache');
    assert.match(cache.text, /跨调用保留/);
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`pdf: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
