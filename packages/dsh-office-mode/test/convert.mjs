/**
 * 文档互转与统一出 PDF（office.convert，第五十三轮）的测试。
 *
 * 这条账说的是「手上的文档换个格式只能重新生成一遍」：tex / md / word 之间没有互转入口，
 * md 也一直出不了 PDF。落点是 office.convert —— 三件引擎（pandoc / 本机 TeX / 宿主
 * LibreOffice）各管一段，`check()` 报实况，`file()` 一次转一份。
 *
 * 分五层：
 *   1. 能力探针与矩阵：三个引擎的实况与「哪些路真能走」；
 *   2. 文本互转：md ↔ docx ↔ tex 真转（魔数 / 结构 / 表格与公式不丢）、内嵌图片抽取；
 *   3. 出 PDF：.tex 走 TeX、Office 三件套走 LibreOffice、纯文本源两条路都验；
 *   4. 边界与负例：覆盖闸、复用、写不出去、读不了的方向都给可执行的错；
 *   5. 经 office_run 的全链路：SDK 面上有它，产物进反馈与台账。
 *
 * 缺哪个引擎就 SKIP 它那几条路并如实说明，不当成通过。
 *
 * 跑法：node test/convert.mjs（默认清掉临时目录；加 --keep 留着人工看）
 */
import { existsSync, linkSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import { createCache } from '../src/engine/cache.js';
import { createEnv } from '../src/engine/kit.js';
import { convertCheck, convertFile, probeConvertEngines, resetConvertProbe } from '../src/convert.js';
import { availableHelpTopics, buildHelp } from '../src/docs.js';
import { runPdfProcess } from '../src/pdf.js';
import { executeRun } from '../src/run.js';
import { probeTexEngines } from '../src/tex.js';

const results = [];
let skipped = 0;
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        if (error?.skip === true) {
            // SKIP 单独计数：这一套几乎每条都依赖本机引擎，把跳过的算进「通过」会让
            // 一台什么都没装的机器也打印「20/20 通过」（复核 M5）。
            skipped += 1;
            results.push({ name, ok: true, note: `SKIP ${error.message}`, skipped: true });
        } else {
            results.push({ name, ok: false, error });
        }
    }
}
function skip(message) {
    const error = new Error(message);
    error.skip = true;
    throw error;
}

const KEEP = process.argv.includes('--keep');
const root = mkdtempSync(join(tmpdir(), 'office-convert-'));
const config = resolveConfig({ cacheDir: '.office/runtime/cache', defaultTheme: 'plain' });
const env = createEnv({ root, themeResolver: () => ({ theme: {} }), config });
const cache = createCache({ root, dir: '.office/runtime/cache' });
const runScript = (script) => executeRun(
    { script, purpose: '转换测试' },
    { agent: { session: { header: { cwd: root } } } },
    config,
);

const probe = await convertCheck();
const HAS_PANDOC = probe.pandoc !== null;
const HAS_TEX = probe.latex !== null;
const HAS_KIT = probe.libreoffice === true;
const probeTex = probeTexEngines();

/** 1×1 的红点 PNG（真的 PNG，用来验「内嵌图片抽出来了」）。 */
const PNG = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
);

const isPdf = (path) => readFileSync(path).subarray(0, 5).toString('latin1') === '%PDF-';
const isPng = (path) => {
    const bytes = readFileSync(path);
    return bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;
};
const isZip = (path) => {
    const bytes = readFileSync(path);
    return bytes[0] === 0x50 && bytes[1] === 0x4b;
};

const MD = `# 互转冒烟

中文段落，含**加粗**、\`行内代码\` 与公式 $E=mc^2$。

- 列表一
- 列表二

| 列A | 列B |
| --- | --- |
| 甲 | 乙 |
`;
writeFileSync(join(root, 'note.md'), MD, 'utf8');
writeFileSync(join(root, 'img.md'), '# 带图\n\n中文说明。\n\n![示意图](pic.png)\n', 'utf8');
writeFileSync(join(root, 'pic.png'), PNG);
writeFileSync(join(root, 'stub.pdf'), '%PDF-1.4\n%%EOF\n');
mkdirSync(join(root, 'adir'), { recursive: true });

// ── 1. 能力探针与矩阵 ──────────────────────────────────────────────────────

await check('convert.check：报出三件引擎的实况，且与各引擎自己的探针一致', async () => {
    assert.equal(typeof probe.available, 'boolean');
    assert.deepEqual(probe.targets, ['md', 'docx', 'tex', 'pdf']);
    assert.ok(probe.sources.includes('md') && probe.sources.includes('docx') && probe.sources.includes('tex'));
    if (HAS_PANDOC) {
        assert.ok(existsSync(probe.pandoc.path), `pandoc 路径不可达：${probe.pandoc.path}`);
        assert.ok(typeof probe.pandoc.version === 'string' && /pandoc/i.test(probe.pandoc.version), probe.pandoc.version);
    } else {
        assert.equal(probe.pandoc, null, '探不到时 pandoc 必须是 null（不能是空对象）');
    }
    assert.equal(probe.latex !== null, probeTex.xelatex !== undefined || probeTex.lualatex !== undefined,
        'latex 那半要与 probeTexEngines 同一口径');
    assert.ok(typeof probe.hint === 'string' && probe.hint !== '');
    const missing = [HAS_PANDOC, HAS_TEX, HAS_KIT].filter((flag) => !flag).length;
    return missing === 0 ? '三件引擎齐' : `缺 ${missing} 件（hint 已写明）`;
});

await check('矩阵只列真能走的路：自己转自己不列，缺引擎时 available:false 且带 reason', async () => {
    const md = probe.matrix.find((row) => row.source === 'md');
    assert.deepEqual(md.targets.map((item) => item.target), ['docx', 'tex', 'pdf']);
    const pdfRow = probe.matrix.find((row) => row.source === 'pdf');
    assert.deepEqual(pdfRow.targets.map((item) => item.target), ['pdf'], 'PDF 源只有「原样留一份」这一条');
    for (const row of probe.matrix) {
        for (const item of row.targets) {
            assert.equal(typeof item.available, 'boolean', `${row.source} → ${item.target} 的 available 要是布尔`);
            if (item.available !== true) assert.ok(typeof item.reason === 'string' && item.reason !== '', `${row.source} → ${item.target} 不可用时要写清缺什么`);
        }
    }
    return `${probe.matrix.length} 类源`;
});

await check('DSH_OFFICE_PANDOC 指到不存在的路径时是「提示」不是「禁用」：仍退回 PATH，真探不到才报缺', async () => {
    const saved = process.env.DSH_OFFICE_PANDOC;
    process.env.DSH_OFFICE_PANDOC = join(root, '没有这个可执行文件.exe');
    resetConvertProbe();
    try {
        const offline = await convertCheck();
        if (offline.pandoc !== null) {
            // 环境变量只是「先看这里」：指错了要退回 PATH 上真装的那一份，而不是成一个坏配置。
            // （这条在装了 pandoc 的机器上才跑得到；没装的机器走下面那半，两条互补。）
            assert.ok(existsSync(offline.pandoc.path), offline.pandoc.path);
            return `退回到 PATH 上的 ${offline.pandoc.path}`;
        }
        const row = offline.matrix.find((item) => item.source === 'md').targets.find((item) => item.target === 'docx');
        assert.equal(row.available, false, '真探不到 pandoc 时 md→docx 必须报不可用');
        assert.match(row.reason, /pandoc/);
        await assert.rejects(() => convertFile('note.md', { to: 'docx' }, env, cache), /要 pandoc/);
        return '缺 pandoc 时矩阵与报错都可执行';
    } finally {
        if (saved === undefined) delete process.env.DSH_OFFICE_PANDOC;
        else process.env.DSH_OFFICE_PANDOC = saved;
        resetConvertProbe();
    }
});

// ── 2. 文本互转 ────────────────────────────────────────────────────────────

await check('md → docx：真写出 docx（ZIP 魔数），再转回 md 时正文、表格与公式都还在', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    const docx = await convertFile('note.md', { to: 'docx' }, env, cache);
    assert.equal(docx.ok, true);
    assert.equal(docx.path, 'note.docx');
    assert.equal(docx.engine, 'pandoc');
    assert.ok(isZip(join(root, 'note.docx')), 'docx 是 ZIP 容器（PK 开头）');
    assert.ok(docx.bytes > 1000, `docx 太小：${docx.bytes}B`);
    assert.ok(docx.steps.join(' ').includes('pandoc'));

    const back = await convertFile('note.docx', { to: 'md', out: 'note.back.md' }, env, cache);
    assert.equal(back.path, 'note.back.md');
    const text = readFileSync(join(root, 'note.back.md'), 'utf8');
    for (const needle of ['中文段落', '**加粗**', '`行内代码`', '列表一', '| 列A | 列B |', '$E = mc^{2}$']) {
        assert.ok(text.includes(needle), `回环后的 md 缺 ${needle}：\n${text}`);
    }
    return `${docx.bytes}B → ${back.bytes}B`;
});

await check('md → tex：默认出能直接编译的整份（导言区 + 中文字体）', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    const tex = await convertFile('note.md', { to: 'tex' }, env, cache);
    assert.equal(tex.path, 'note.tex');
    const text = readFileSync(join(root, 'note.tex'), 'utf8');
    assert.match(text, /\\documentclass/, '没有导言区的片段编不了，默认要给整份');
    assert.match(text, /\\section\{互转冒烟\}/);
    if (probe.cjkFont === null) {
        assert.ok(env.warnings.some((line) => line.includes('中文字体')), '探不到中文字体时要警告（pandoc 会静默丢字）');
    } else {
        assert.match(text, /CJKmainfont/, `没挂中文字体：${probe.cjkFont}`);
    }
    // fragment:true 只给正文片段（往自己的论文里并时用）。
    const fragment = await convertFile('note.md', { to: 'tex', out: 'note.frag.tex', fragment: true }, env, cache);
    const fragText = readFileSync(join(root, fragment.path), 'utf8');
    assert.ok(!/\\documentclass/.test(fragText), 'fragment:true 不该带导言区');
    return `${tex.bytes}B / 片段 ${fragment.bytes}B`;
});

await check('tex → md / docx：反向也能转（pandoc 读 LaTeX）', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    if (!existsSync(join(root, 'note.tex'))) skip('上一条没产出 note.tex');
    const md = await convertFile('note.tex', { to: 'md', out: 'from-tex.md' }, env, cache);
    assert.ok(readFileSync(join(root, 'from-tex.md'), 'utf8').includes('中文段落'), readFileSync(join(root, 'from-tex.md'), 'utf8'));
    const docx = await convertFile('note.tex', { to: 'docx', out: 'from-tex.docx' }, env, cache);
    assert.ok(isZip(join(root, 'from-tex.docx')));
    return `${md.bytes}B / ${docx.bytes}B`;
});

await check('docx 里的内嵌图片会被抽成 <目标名>.media/，md 里的引用同时改写', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    await convertFile('img.md', { to: 'docx' }, env, cache);
    const back = await convertFile('img.docx', { to: 'md', out: 'img.back.md' }, env, cache);
    const dir = join(root, 'img.back.media');
    assert.ok(existsSync(dir), `没抽出媒体目录：${readdirSync(root).join('、')}`);
    const files = readdirSync(dir);
    assert.equal(files.length, 1, `媒体文件数不对：${files.join('、')}`);
    assert.ok(isPng(join(dir, files[0])), '抽出来的得是真 PNG');
    assert.ok(Array.isArray(back.media) && back.media.length === 1, 'media 要登记进返回值（才进台账）');
    const text = readFileSync(join(root, 'img.back.md'), 'utf8');
    assert.ok(text.includes('img.back.media/'), `md 里的图片路径没改写：\n${text}`);
    assert.ok(!/\.office[\\/]cache/.test(text), 'md 里不该留缓存目录的路径');

    // media:false 只出文本，并且要警告「引用会断」。
    const skipped = await convertFile('img.docx', { to: 'md', out: 'img.plain.md', media: false }, env, cache);
    assert.equal(skipped.media, undefined);
    assert.ok(!existsSync(join(root, 'img.plain.media')));
    assert.ok(env.warnings.some((line) => line.includes('内嵌图片没有抽出来')), '断了链要说一声');
    // 主文件被覆盖闸拒掉时，媒体不许先落地（否则留一堆没人引用的孤儿图）—— 复核瑕疵。
    writeFileSync(join(root, 'clash-media.md'), '这不是转换出来的内容', 'utf8');
    await assert.rejects(() => convertFile('img.docx', { to: 'md', out: 'clash-media.md' }, env, cache), /已存在且内容不同/);
    assert.ok(!existsSync(join(root, 'clash-media.media')), '主文件被拒时不该留下媒体目录');
    return `1 张图 / ${back.bytes}B`;
});

await check('幂等：同一份源重转命中复用（reused:true，不重写、不重复登记）', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    const first = await convertFile('img.md', { to: 'docx', out: 'idem.docx' }, env, cache);
    const before = readFileSync(join(root, 'idem.docx'));
    const writesBefore = env.writes.length;
    const again = await convertFile('img.md', { to: 'docx', out: 'idem.docx' }, env, cache);
    assert.equal(again.reused, true, '同一份源 + 同样的目标必须命中复用（pandoc 的时间戳要钉住）');
    assert.equal(again.wrote, false);
    assert.ok(before.equals(readFileSync(join(root, 'idem.docx'))), '复用时文件一个字节都不该动');
    assert.equal(env.writes.length, writesBefore, '复用时不该再登记一次（不重复登记）');
    return `${first.bytes}B 命中`;
});

// ── 3. 出 PDF ──────────────────────────────────────────────────────────────

await check('tex → pdf：本机 TeX 真编译（PDF 魔数 + 页数 + 产物登记）', async () => {
    if (!HAS_TEX) skip('本机没有 TeX 发行版');
    if (!existsSync(join(root, 'note.tex'))) skip('没有可编译的 note.tex（需要 pandoc 先生成）');
    const pdf = await convertFile('note.tex', { to: 'pdf' }, env, cache);
    assert.equal(pdf.path, 'note.pdf');
    assert.ok(isPdf(join(root, 'note.pdf')));
    assert.ok(['latexmk', 'xelatex', 'lualatex'].includes(pdf.engine), pdf.engine);
    assert.ok(pdf.pages === null || pdf.pages >= 1, `页数异常：${pdf.pages}`);
    return `${pdf.engine} / ${pdf.bytes}B${pdf.pages === null ? '' : ` / ${pdf.pages} 页`}`;
});

await check('md → pdf：两条路线各自出品（LibreOffice 与 xelatex），engine 写清走了哪条', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    const routes = [];
    if (HAS_KIT) {
        const lo = await convertFile('note.md', { to: 'pdf', out: 'md-lo.pdf', engine: 'libreoffice' }, env, cache);
        assert.ok(isPdf(join(root, 'md-lo.pdf')));
        assert.match(lo.engine, /libreoffice/i);
        routes.push(`libreoffice ${lo.bytes}B`);
    }
    if (HAS_TEX) {
        const tex = await convertFile('note.md', { to: 'pdf', out: 'md-tex.pdf', engine: 'tex' }, env, cache);
        assert.ok(isPdf(join(root, 'md-tex.pdf')));
        assert.match(tex.engine, /xelatex|lualatex/);
        routes.push(`${tex.engine} ${tex.bytes}B`);
    }
    if (routes.length === 0) skip('既没有 LibreOffice 引擎也没有 TeX 发行版');
    return routes.join(' / ');
});

await check('docx → pdf：走宿主 LibreOffice 引擎（与 office.preview 同一份）', async () => {
    if (!HAS_KIT) skip('宿主 LibreOffice 引擎解析不到');
    if (!existsSync(join(root, 'note.docx'))) skip('没有 note.docx');
    const pdf = await convertFile('note.docx', { to: 'pdf', out: 'docx.pdf' }, env, cache);
    assert.ok(isPdf(join(root, 'docx.pdf')));
    assert.equal(pdf.engine, 'libreoffice');
    return `${pdf.bytes}B`;
});

await check('纯文本源出 PDF 的 auto 路线：LibreOffice 优先，产物进 files/otherFiles', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    if (!HAS_KIT && !HAS_TEX) skip('两条 PDF 路线都不可用');
    const pdf = await convertFile('note.md', { to: 'pdf', out: 'auto.pdf' }, env, cache);
    assert.ok(isPdf(join(root, 'auto.pdf')));
    assert.match(pdf.engine, HAS_KIT ? /libreoffice/i : /xelatex|lualatex/);
    return pdf.engine;
});

await check('老格式与 OpenDocument 也能出 PDF：.odt 真转一份，.doc / .ods / .odp 走同一条 direct 路', async () => {
    // 「各种文档都要能出 PDF」这句要落到实处：宿主 LibreOffice 引擎读得了 .doc / .ods / .odp
    // （kit 的 IMAGE_FORMATS 列了它们），所以矩阵里那三类源必须只剩「出 PDF」这一条路。
    for (const id of ['doc', 'odf', 'xlsx', 'pptx']) {
        const row = probe.matrix.find((item) => item.source === id);
        assert.ok(row !== undefined, `矩阵里缺 ${id}`);
        assert.deepEqual(row.targets.map((item) => item.target), ['pdf'], `${id} 只该有出 PDF 这一条`);
        assert.equal(row.targets[0].engine, 'direct');
        assert.equal(row.targets[0].available, HAS_KIT, `${id} 的可用性要跟 LibreOffice 引擎一致`);
    }
    // .doc 源转文本要直说「pandoc 读不了」，而不是含糊地失败。
    writeFileSync(join(root, 'legacy.doc'), 'placeholder', 'utf8');
    await assert.rejects(() => convertFile('legacy.doc', { to: 'md' }, env, cache), /pandoc 读不了/);
    if (!HAS_KIT) skip('宿主 LibreOffice 引擎解析不到（.doc / .odf 的路线已核对）');
    if (!HAS_PANDOC) return '.doc / .odf 的路线已核对（造 .odt 夹具要 pandoc）';
    // 用 pandoc 造一份 .odt 当「外部来的文档」（转换目标里没有 odt，正合场景），再让它出 PDF。
    const fixtureDir = join(root, '.office/runtime/cache/fixture');
    mkdirSync(fixtureDir, { recursive: true });
    const made = await runPdfProcess(probe.pandoc.path, ['note.md', '-o', 'note.odt'], {
        outPath: join(fixtureDir, 'out.txt'),
        errPath: join(fixtureDir, 'err.txt'),
        cwd: root,
        timeoutMs: 60_000,
    });
    assert.equal(made.code, 0, `造 .odt 夹具失败：${made.err}`);
    const pdf = await convertFile('note.odt', { to: 'pdf', out: 'odt.pdf' }, env, cache);
    assert.ok(isPdf(join(root, 'odt.pdf')));
    assert.equal(pdf.engine, 'libreoffice');
    const md = await convertFile('note.odt', { to: 'md', out: 'odt.md' }, env, cache);
    assert.ok(readFileSync(join(root, 'odt.md'), 'utf8').includes('中文段落'), 'pandoc 读得了 .odt 文本');
    return `${pdf.bytes}B`;
});

await check('tex → pdf：clean 收敛不删交付物，默认落点的既有 PDF 先拦后写（复核 D1 / M1）', async () => {
    if (!HAS_TEX) skip('本机没有 TeX 发行版');
    if (!existsSync(join(root, 'note.tex'))) skip('没有可编译的 note.tex（需要 pandoc 先生成）');
    // clean:'all' 在 latexmk 里是 `-C`（连 PDF 一起删）—— 转换必须收敛成 'aux' 并说一声。
    const notesBefore = env.notes.length;
    const cleaned = await convertFile('note.tex', { to: 'pdf', out: 'cleaned.pdf', clean: 'all' }, env, cache);
    assert.ok(isPdf(join(root, 'cleaned.pdf')), 'clean:all 也不许把这次的交付物删掉');
    assert.ok(env.notes.slice(notesBefore).some((line) => line.includes("clean:'all'")), '收敛成 aux 要说一声');
    // clean:'none' 要真的留下中间产物。
    await convertFile('note.tex', { to: 'pdf', out: 'kept.pdf', clean: 'none' }, env, cache);
    assert.ok(existsSync(join(root, 'note.aux')) || existsSync(join(root, 'note.log')), 'clean:none 应当保留 .aux / .log');
    // 默认落点上已经有一份 PDF（前面几条留下的）：编译会直接覆盖它，所以必须在编译**之前**拦住。
    const beforeBytes = readFileSync(join(root, 'note.pdf'));
    await assert.rejects(() => convertFile('note.tex', { to: 'pdf' }, env, cache), /默认落点就是 TeX 写 PDF 的位置/);
    assert.ok(beforeBytes.equals(readFileSync(join(root, 'note.pdf'))), '被拒时那份 PDF 一个字节都不该动');
    // 传了 overwrite:true 才写：登记产物、reused 恒为 false、hint 说明这条例外。
    const warningsBefore = env.warnings.length;
    const landing = await convertFile('note.tex', { to: 'pdf', overwrite: true }, env, cache);
    assert.equal(landing.path, 'note.pdf');
    assert.equal(landing.reused, false, '.tex 路线的默认落点由 TeX 自己写，reused 不适用');
    assert.equal(landing.overwrote, true, '覆盖了既有的 PDF 要照实报 overwrote');
    assert.match(landing.hint, /默认落点就是 TeX/, landing.hint);
    assert.ok(env.warnings.slice(warningsBefore).some((line) => line.includes('原本有一份 PDF')), '覆盖既有的 PDF 要说一声');
    // 显式 out 到别处时走通用覆盖闸：换一份内容不同的源写同一个目标，必须被拒。
    await convertFile('note.tex', { to: 'pdf', out: 'tex-shared.pdf' }, env, cache);
    writeFileSync(join(root, 'note2.tex'), readFileSync(join(root, 'note.tex'), 'utf8').replace('互转冒烟', '另一个标题'), 'utf8');
    await assert.rejects(() => convertFile('note2.tex', { to: 'pdf', out: 'tex-shared.pdf' }, env, cache), /已存在且内容不同/);
    return 'clean 收敛 + 落点先拦后写 + 显式 out 守闸';
});

await check('别名绕闸：硬链接指的同一个文件也不许当目标（复核 D1）', async () => {
    // 只比路径字符串时，`另存一份再转`（硬链接 / 符号链接 / Windows 大小写）会把源文件本身
    // 换成目标格式 —— 第五十一轮给 office.files.move 建过同一套判据，这里补上。
    const source = join(root, 'alias-src.md');
    const link = join(root, 'alias-link.md');
    writeFileSync(source, '# 别名', 'utf8');
    try {
        linkSync(source, link);
    } catch (error) {
        skip(`本机文件系统不支持硬链接：${error.message}`);
    }
    const markdown = readFileSync(source);
    await assert.rejects(
        () => convertFile('alias-src.md', { to: 'docx', out: 'alias-link.md', overwrite: true }, env, cache),
        /目标就是源文件自己/,
    );
    assert.ok(markdown.equals(readFileSync(source)), '被拒时源文件一个字节都不该动');
    return '硬链接也拦得住';
});

await check('中文字体：显式 font 一定挂上；大文件末尾的中文不会被漏判（复核 T4 / M3）', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    // 显式点名 → 即使源里没有中文也要挂（用户的参数不该被自动判据丢掉）。
    const forced = await convertFile('note.md', { to: 'tex', out: 'forced.tex', font: 'Noto Sans CJK SC' }, env, cache);
    assert.match(readFileSync(join(root, 'forced.tex'), 'utf8'), /CJKmainfont.*Noto Sans CJK SC/);
    void forced;
    // 大文件：中文只在 20 万字符之后 —— 判据要整份扫（分块），否则这份 .tex 会静默缺字。
    if (probe.cjkFont === null) skip('本机没探到中文字体（大文件那条判据无处可挂）');
    const big = join(root, 'big.md');
    writeFileSync(big, `# 大文件\n\n${'a'.repeat(250_000)}\n\n末尾的中文段落。\n`, 'utf8');
    const bigTex = await convertFile('big.md', { to: 'tex', out: 'big.tex' }, env, cache);
    assert.match(readFileSync(join(root, 'big.tex'), 'utf8'), /CJKmainfont/, '大文件末尾的中文也必须被判出来');
    // 分块边界：把中文正好放在 1 MiB 块的接缝附近，重叠 3 字节才不会漏。
    const seam = join(root, 'seam.md');
    writeFileSync(seam, `# 接缝\n\n${'b'.repeat((1 << 20) - 4)}中\n`, 'utf8');
    const seamTex = await convertFile('seam.md', { to: 'tex', out: 'seam.tex' }, env, cache);
    assert.match(readFileSync(join(root, 'seam.tex'), 'utf8'), /CJKmainfont/, '块接缝上的中文不能漏');
    void bigTex;
    void seamTex;
    return 'font 生效 + 整份扫描';
});

// ── 4. 边界与负例 ──────────────────────────────────────────────────────────

await check('读不了的方向给出可执行的错（PDF 正文走 office.pdf，不是 pandoc）', async () => {
    await assert.rejects(() => convertFile('stub.pdf', { to: 'docx' }, env, cache), (error) => {
        assert.match(error.message, /pandoc 读不了/);
        assert.match(error.message, /office\.pdf\.text/);
        return true;
    });
    const copy = await convertFile('stub.pdf', { to: 'pdf', out: 'stub-copy.pdf' }, env, cache);
    assert.equal(copy.engine, 'copy');
    assert.ok(readFileSync(join(root, 'stub-copy.pdf')).equals(readFileSync(join(root, 'stub.pdf'))));
    return 'pdf → pdf 原样复用';
});

await check('参数与路径的负例：to 写错 / 源不存在 / 源是目录 / 不认识的源 / 目标就是源', async () => {
    await assert.rejects(() => convertFile('note.md', { to: 'xlsx' }, env, cache), /to 要写成/);
    await assert.rejects(() => convertFile('没有这份.md', { to: 'docx' }, env, cache), /找不到/);
    // 源是目录：先按「是目录不是文件」拦住，别一路走到引擎那儿去。
    await assert.rejects(() => convertFile('adir', { to: 'docx' }, env, cache), /是目录/);
    await assert.rejects(() => convertFile('adir/', { to: 'docx' }, env, cache), /是目录/);
    writeFileSync(join(root, 'weird.zzz'), 'x', 'utf8');
    await assert.rejects(() => convertFile('weird.zzz', { to: 'pdf' }, env, cache), /还不认识/);
    await assert.rejects(() => convertFile('note.md', { to: 'md' }, env, cache), /目标就是源文件自己/);
    await assert.rejects(() => convertFile('note.md', { to: 'docx', out: 'adir' }, env, cache), /是一个目录/);
    await assert.rejects(() => convertFile('note.md', { to: 'docx', out: '../逃出去.docx' }, env, cache), /工作目录之外/);
    return '8 类负例';
});

await check('覆盖闸：内容不同要 overwrite:true；.tex 出 PDF 不接受 LibreOffice', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    writeFileSync(join(root, 'other.md'), `${MD}\n再补一段。\n`, 'utf8');
    await convertFile('other.md', { to: 'docx', out: 'clash.docx' }, env, cache);
    await assert.rejects(
        () => convertFile('note.md', { to: 'docx', out: 'clash.docx' }, env, cache),
        /已存在且内容不同/,
    );
    const forced = await convertFile('note.md', { to: 'docx', out: 'clash.docx', overwrite: true }, env, cache);
    assert.equal(forced.overwrote, true);
    assert.equal(forced.reused, false);
    if (existsSync(join(root, 'note.tex'))) {
        await assert.rejects(
            () => convertFile('note.tex', { to: 'pdf', engine: 'libreoffice' }, env, cache),
            /走不了 LibreOffice/,
        );
    }
    return '覆盖 + 路线冲突';
});

// ── 5. 经 office_run 的全链路 ──────────────────────────────────────────────

await check('office_run：office.convert 在 SDK 面上，产物进反馈（docx 进 files / pdf 进 otherFiles）', async () => {
    if (!HAS_PANDOC) skip('本机没有 pandoc');
    const result = await runScript(`
        const c = await office.convert.check();
        office.assert(c.targets.length === 4, 'targets 不对');
        const docx = await office.convert.file('note.md', { to: 'docx', out: 'run.docx' });
        return { targets: c.targets, path: docx.path, engine: docx.engine, from: docx.from, to: docx.to };
    `);
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.path, 'run.docx');
    assert.equal(result.returned.engine, 'pandoc');
    assert.ok(result.files.some((file) => file.path === 'run.docx'), `docx 没进 files：${JSON.stringify(result.files)}`);
    const docx = result.files.find((file) => file.path === 'run.docx');
    assert.equal(docx.format, 'word', 'docx 要按 word 格式复检');
    assert.ok(docx.stats !== null || (Array.isArray(docx.outline) && docx.outline.length > 0),
        `写出的 docx 会被自动复检：${JSON.stringify(docx)}`);

    if (HAS_KIT || HAS_TEX) {
        const pdfRun = await runScript(`const r = await office.convert.file('note.md', { to: 'pdf', out: 'run.pdf' }); return r.path;`);
        assert.equal(pdfRun.ok, true, pdfRun.error?.message);
        assert.ok(pdfRun.otherFiles.some((file) => file.path === 'run.pdf'), `PDF 没进 otherFiles：${JSON.stringify(pdfRun.otherFiles)}`);
    }
    return 'SDK 面 + 反馈登记';
});

await check('office_help：convert 话题在（默认层签名、全文层矩阵、两份索引与 run 话题都提到）', async () => {
    assert.ok(availableHelpTopics().includes('convert'), 'convert 必须进话题真源');
    const brief = await buildHelp('convert');
    for (const needle of ['office.convert.check', 'office.convert.file', 'pandoc', 'LibreOffice', 'md / docx / tex / pdf', 'overwrite:true']) {
        assert.ok(brief.text.includes(needle), `convert 默认层缺 ${needle}`);
    }
    const full = await buildHelp('convert', { detail: true });
    assert.ok(full.text.includes('SOURCE_DATE_EPOCH'), '全文层要讲复用与时间戳');
    assert.ok(full.text.includes('thuthesis'), '全文层要讲论文源的边界');
    for (const detail of [false, true]) {
        const index = await buildHelp('', { detail });
        assert.ok(index.text.split('\n').some((line) => line.startsWith('- convert：')), `${detail ? '全文' : '默认'}层索引里没列 convert`);
    }
    const run = await buildHelp('run');
    assert.ok(run.text.includes('office.convert'), 'run 默认层要提到转换入口');
    // 两层不许自相矛盾：全文层那句「外部程序几处」也要算上转换（复核 D3）。
    const runFull = await buildHelp('run', { detail: true });
    assert.ok(runFull.text.includes('office.convert'), 'run 全文层也要提到转换入口');
    assert.ok(runFull.text.includes('六处'), '全文层的外部程序处数要与默认层一致');
    const unknown = await buildHelp('__不存在的話題__');
    assert.ok(unknown.text.includes('convert'), '认不出的话题要把它列进可用清单');
    return '默认层 + 全文层 + 索引 + run';
});

await check('探测缓存可复位：resetConvertProbe 之后重新探到同一批引擎', async () => {
    const before = probeConvertEngines();
    resetConvertProbe();
    const after = probeConvertEngines();
    assert.equal(after.pandoc, before.pandoc);
    return after.pandoc === undefined ? 'pandoc 未装' : after.pandoc;
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
// 汇总把「跳过」单独报出来：跳过不是通过（复核 M5）。
console.log(skipped === 0
    ? `convert: ${results.length - failed.length}/${results.length} 通过`
    : `convert: ${results.length - skipped - failed.length} 通过 / ${skipped} 跳过 / ${failed.length} 失败（共 ${results.length} 条；跳过 = 本机缺对应引擎）`);
if (!KEEP) rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);

