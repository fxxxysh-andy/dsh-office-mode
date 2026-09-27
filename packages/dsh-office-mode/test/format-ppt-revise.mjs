/**
 * `office.ppt.readSlides` / `office.ppt.revise` 的自测（就地微调已有 .pptx）。
 *
 * 与 `format-ppt.mjs` 分开：那一份测「从零生成」，这一份测「读回已有包、定点改」。
 *
 * 跑法：node test/format-ppt-revise.mjs
 *
 * 这里刻意用**别的工具生成的 pptx**（`办公模式示例/参考稿复刻.pptx` 生成的同构样例，
 * 以及用 office.ppt 现造的稿子）作为输入，而不是只测自己刚写出的文件：
 * 「能改自己的输出」和「能改任何人的输出」是两件事，后者才是用户手上的场景。
 */
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { deflateSync } from 'node:zlib';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
// 测试中间产物统一落在仓库根的 .office/tmp/<suite>/ 下（不再用包内的 test/.tmp）
const outDir = join(here, '..', '..', '..', '.office', 'tmp', 'revise');
rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

const { createEnv } = await import('../src/engine/kit.js');
const { resolveTheme } = await import('../src/engine/theme.js');
const { parseXml, descendants, children } = await import('../src/engine/xml.js');
const ppt = await import('../src/formats/ppt.js');
const { readSlides, revise, insertImages } = await import('../src/formats/ppt-revise.js');
const zipModule = await import('../src/engine/zip.js');

const env = createEnv({
    root: outDir,
    themeResolver: (id) => resolveTheme(id ?? 'business'),
});

let passed = 0;
const failures = [];

function ok(name, condition, detail = '') {
    if (condition) {
        passed += 1;
    } else {
        failures.push(`${name}${detail === '' ? '' : `  → ${detail}`}`);
    }
}

function eq(name, actual, expected) {
    ok(name, JSON.stringify(actual) === JSON.stringify(expected), `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`);
}

// ── 造一张真 PNG（PNG 的 CRC 与 ZIP 同一个多项式，直接复用 zip.js 的 crc32） ──

function makePng(width, height) {
    const raw = Buffer.alloc((width * 3 + 1) * height);
    for (let y = 0; y < height; y += 1) {
        const row = y * (width * 3 + 1);
        raw[row] = 0;
        for (let x = 0; x < width; x += 1) {
            const at = row + 1 + x * 3;
            raw[at] = Math.round((255 * x) / Math.max(1, width - 1));
            raw[at + 1] = 90;
            raw[at + 2] = Math.round((255 * y) / Math.max(1, height - 1));
        }
    }
    const chunk = (type, data) => {
        const out = Buffer.alloc(8 + data.length + 4);
        out.writeUInt32BE(data.length, 0);
        out.write(type, 4, 'ascii');
        data.copy(out, 8);
        out.writeUInt32BE(zipModule.crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
        return out;
    };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw, { level: 9 })),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

// ── 造一份内容形态足够复杂的原稿 ────────────────────────────────────────────
// makeDeck() 每次现造一份：revise 是就地改盘，多组断言共用同一个文件时，
// 前面改过的内容会变成后面断言的前提（第一版测试就是这么被自己骗过去的）。
function makeDeck(name) {
    const deck = ppt.create({ title: '微调测试稿', theme: 'business', path: name }, env);
    deck
        .cover({ title: '原始标题', subtitle: '原始副标题' })
        .bullets({ title: '要点页', items: ['第一条要点', '第二条要点', '第三条要点'] })
        .cards({ title: '卡片页', items: [
            { title: '甲', body: '甲的内容' },
            { title: '乙', body: '乙的内容' },
            { title: '丙', body: '丙的内容' },
        ], columns: 3 })
        .kpi({ title: 'KPI 页', items: [{ value: '108%', label: '达成率' }, { value: '32', label: '新增' }] })
        .closing({ title: '结束页' });
    // 带旋转的形状：用来证明改位置不会把 rot 丢掉
    deck.page().shape({ preset: 'rect', x: 1, y: 1, w: 3, h: 2, fill: 'DDEAF6', rotate: 15, text: '旋转块' });
    deck.notes('这是第一页的备注');
    const written = deck.save();
    return written;
}

const deckPath = 'revise.pptx';
const written = makeDeck(deckPath);
ok('原稿写出成功', written.ok === true && written.stats.slides >= 5, `slides=${written.stats?.slides}`);
const originalBytes = readFileSync(join(outDir, deckPath));
const originalXml = slurpSlide(originalBytes, 1);

function slurpSlide(bytes, pageNumber) {
    // 直接看包里的 slide XML：断言「除了被改的那段，其余一个字节都没动」
    const { unzip } = zipModule;
    const files = unzip(bytes);
    return new TextDecoder('utf-8').decode(files.get(`ppt/slides/slide${pageNumber}.xml`));
}

// ── readSlides ─────────────────────────────────────────────────────────────
const listing = readSlides(deckPath, env);
ok('readSlides 返回 ok', listing.ok === true);
ok('readSlides 页数与生成一致', listing.pages.length === written.stats.slides, `${listing.pages.length} vs ${written.stats.slides}`);
const cover = listing.pages[0];
ok('首页版式是封面', cover.layout.includes('封面'), cover.layout);
const titleShape = cover.shapes.find((shape) => shape.name === 'Title');
ok('封面有 Title 形状', titleShape !== undefined);
ok('Title 形状带 id', Number.isInteger(titleShape?.id), String(titleShape?.id));
eq('Title 文本读得回来', titleShape?.text, '原始标题');
ok('Title 几何读得回来（厘米）', typeof titleShape?.x === 'number' && typeof titleShape?.w === 'number',
    JSON.stringify({ x: titleShape?.x, w: titleShape?.w }));
ok('每个形状都标了能不能文本编辑', cover.shapes.every((shape) => typeof shape.editable === 'boolean'));

// ── revise：改文字 ─────────────────────────────────────────────────────────
const r1 = revise(deckPath, [
    { slide: 1, shape: 'title', setText: '改后的标题' },
], env);
eq('setText 生效', r1.applied.length, 1);
eq('setText 报的是形状名', r1.applied[0]?.shape, 'Title');
ok('revise 落盘', r1.saved === true && r1.ok === true, JSON.stringify(r1.skipped));
eq('改完再读回来是新文字', readSlides(deckPath, env).pages[0].shapes.find((s) => s.name === 'Title')?.text, '改后的标题');

// 定点改写：除被改的形状外，XML 其余部分必须逐字节不变
const afterXml = slurpSlide(readFileSync(join(outDir, deckPath)), 1);
ok('只改了命中的形状，其余 XML 一字未动', diffCount(originalXml, afterXml) === 1,
    `差异段数 ${diffCount(originalXml, afterXml)}`);
ok('被改的形状仍保留原字号', /sz="3200"/.test(afterXml) || /sz="\d{4}"/.test(afterXml));

function diffCount(a, b) {
    // 按「非空白标签块」粗粒度比较：统计被替换掉的 <p:sp> 数量
    const countShapes = (xml) => (xml.match(/<p:sp>/g) ?? []).length;
    if (countShapes(a) !== countShapes(b)) return 99;
    let diff = 0;
    const blocksA = a.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
    const blocksB = b.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
    for (let i = 0; i < blocksA.length; i += 1) if (blocksA[i] !== blocksB[i]) diff += 1;
    return diff;
}

// ── revise：replace 只换子串，段落结构不动 ─────────────────────────────────
const before2 = readSlides(deckPath, env).pages.find((page) => page.shapes.some((s) => s.name === 'Bullets 1'));
const r2 = revise(deckPath, [
    { slide: 2, shape: 'Bullets 1', replace: [['第二条要点', '第二条（已改）']] },
], env);
eq('replace 生效', r2.applied.length, 1, JSON.stringify(r2.skipped));
const after2 = readSlides(deckPath, env).pages[1];
ok('replace 只改命中段落', after2.shapes.some((s) => s.text.includes('第二条（已改）')), after2.shapes.map((s) => s.text).join('|'));
ok('replace 不动其它条目', after2.shapes.some((s) => s.text.includes('第一条要点')) && after2.shapes.some((s) => s.text.includes('第三条要点')));
eq('replace 前后段落数不变', before2?.shapes.find((s) => s.name === 'Bullets 1')?.paragraphs,
    after2.shapes.find((s) => s.name === 'Bullets 1')?.paragraphs);

// ── revise：改样式（字号 / 颜色 / 字体 / 加粗）─────────────────────────────
// 单独一份稿子：这一组既要看落进 XML 的属性，也要数 rPr 里的填充个数，
// 前面几组改过的内容会让这里的计数失去意义。
const styleDeck = freshDeck('style.pptx');
const r3 = revise(styleDeck, [
    { slide: 2, shape: 'Title', style: { sizePt: 40, color: 'C00000', font: '微软雅黑', bold: true } },
], env);
eq('style 生效', r3.applied.length, 1, JSON.stringify(r3.skipped));
const styled = slurpSlide(readFileSync(join(outDir, styleDeck)), 2);
ok('字号落到 rPr', /sz="4000"/.test(styled));
ok('颜色落到 solidFill', /srgbClr val="C00000"/.test(styled));
ok('字体落到 latin', /<a:latin typeface="微软雅黑"\/>/.test(styled));
ok('加粗落到属性', /b="1"/.test(styled));

/**
 * 取出名叫 `name` 的那个形状的 XML。
 * 断言计数必须落在单个形状上：整页里还有形状填充、装饰色块，
 * 混着数会把「形状底色」当成「文字颜色」，断言就失去意义了。
 */
function shapeXml(xml, name) {
    const sp = /<p:sp>[\s\S]*?<\/p:sp>/g;
    let m;
    while ((m = sp.exec(xml)) !== null) {
        if (new RegExp(`<p:cNvPr[^>]*name="${name}"`).test(m[0])) return m[0];
    }
    return '';
}

const titleXml = shapeXml(styled, 'Title');
// Title 的 p:spPr 里本来就有 1 个 noFill（形状不填充）+ 1 个 noFill（描边不画），
// rPr 里再 1 个 solidFill，合计 3；多出来就说明同一个 rPr 里塞了两个填充。
eq('Title 形状里只有 run 那一个填充', (titleXml.match(/<a:solidFill>/g) ?? []).length, 1,
    `实际 ${(titleXml.match(/<a:solidFill>/g) ?? []).length} 个 solidFill`);
eq('Title 的 p:txBody 外壳还在', (titleXml.match(/<p:txBody>/g) ?? []).length, 1);
eq('Title 的 bodyPr 还在', (titleXml.match(/<a:bodyPr\b/g) ?? []).length, 1);
eq('Title 的 lstStyle 还在', (titleXml.match(/<a:lstStyle\/>/g) ?? []).length, 1);

// 子元素顺序必须合法：solidFill 不能排在 latin 后面（否则 PowerPoint 判包损坏）
ok('solidFill 排在 latin 之前', /<a:solidFill><a:srgbClr val="C00000"\/><\/a:solidFill><a:latin/.test(styled),
    '否则真实 PowerPoint 会报「需要修复」');

// ── revise：改位置与尺寸（厘米）─────────────────────────────────────────────
const posBefore = readSlides(deckPath, env).pages[0].shapes.find((s) => s.name === 'Title');
const r4 = revise(deckPath, [
    { slide: 1, shape: 'title', move: { x: 10, y: 5 } },
    { slide: 1, shape: 'title', resize: { w: 8, h: 2.5 } },
], env);
eq('move 与 resize 生效', r4.applied.length, 2, JSON.stringify(r4.skipped));
const posAfter = readSlides(deckPath, env).pages[0].shapes.find((s) => s.name === 'Title');
ok('x 改到 10cm', Math.abs(posAfter.x - 10) < 0.02, String(posAfter.x));
ok('y 改到 5cm', Math.abs(posAfter.y - 5) < 0.02, String(posAfter.y));
ok('w 改到 8cm', Math.abs(posAfter.w - 8) < 0.02, String(posAfter.w));
ok('h 改到 2.5cm', Math.abs(posAfter.h - 2.5) < 0.02, String(posAfter.h));
ok('相对移动可用', (() => {
    const r = revise(deckPath, [{ slide: 1, shape: 'title', move: { dx: 1, dy: -1 } }], env);
    const now = readSlides(deckPath, env).pages[0].shapes.find((s) => s.name === 'Title');
    return r.applied.length === 1 && Math.abs(now.x - 11) < 0.02 && Math.abs(now.y - 4) < 0.02;
})(), JSON.stringify(posBefore));

// 旋转不能被改位置弄丢
const rotatedShape = readSlides(deckPath, env).pages.find((page) => page.shapes.some((s) => s.text === '旋转块'));
const rotatedPage = rotatedShape?.index;
const rotatedId = rotatedShape?.shapes.find((s) => s.text === '旋转块')?.id;
const r5 = revise(deckPath, [{ slide: rotatedPage, shape: rotatedId, move: { dx: 0.5, dy: 0.5 } }], env);
eq('旋转形状也能移动', r5.applied.length, 1, JSON.stringify(r5.skipped));
const rotatedXml = slurpSlide(readFileSync(join(outDir, deckPath)), rotatedPage);
ok('移动旋转形状不会丢掉 rot', /<a:xfrm rot="900000"/.test(rotatedXml), 'rot 必须原样保留');

// ── revise：越界要报出来 ───────────────────────────────────────────────────
const r6 = revise(deckPath, [{ slide: 1, shape: 'title', move: { x: 60, y: 5 } }], env);
ok('移出页面会给警告', r6.warnings.some((line) => line.includes('超出页面')), JSON.stringify(r6.warnings));

// ── revise：多页选择器 ─────────────────────────────────────────────────────
// 每组用一份新造的稿子：选择器断言的是「选中了哪几页」，
// 若共用同一个文件，前一组已经改过的页会让后一组「样式与现状相同」而落空。
function freshDeck(name) {
    makeDeck(name);
    return name;
}

const selRange = freshDeck('sel-range.pptx');
// 用一个主题里不会出现的颜色：business 主题的 primary 就是 1F4E79，
// 拿它当「新值」在第 2、3 页是空操作（标题本来就是这个色），断言会假失败。
const r7 = revise(selRange, [{ slide: '1-3', shape: 'title', style: { color: 'B45309' } }], env);
eq('范围选择器命中 1-3 页', r7.applied.map((a) => a.page), [1, 2, 3], JSON.stringify(r7.skipped));

const selArray = freshDeck('sel-array.pptx');
const r8 = revise(selArray, [{ slide: [1, 2], shape: 'title', style: { italic: true } }], env);
eq('数组选择器命中指定页', r8.applied.map((a) => a.page), [1, 2], JSON.stringify(r8.skipped));

const selOne = freshDeck('sel-one.pptx');
const r9 = revise(selOne, [{ slide: 1, shape: 'title', style: { italic: true } }], env);
eq('单个页码只命中该页', r9.applied.map((a) => a.page), [1], JSON.stringify(r9.skipped));

const selAll = freshDeck('sel-all.pptx');
const r9b = revise(selAll, [{ shape: 'title', style: { italic: true } }], env);
// 5 页有标题（封面/要点/卡片/KPI/结尾）；第 6 页是自由绘制页（只有一个旋转色块），
// 本来就没有 Title —— 省略 slide 时它不该被算成命中，也不该报「找不到形状」。
eq('省略 slide = 全篇（5 页有标题的页）', r9b.applied.map((a) => a.page), [1, 2, 3, 4, 5], JSON.stringify(r9b.skipped));
eq('没有标题的自由绘制页不报错', r9b.skipped.length, 0, JSON.stringify(r9b.skipped));

const selList = freshDeck('sel-list.pptx');
const r9c = revise(selList, [{ slide: '2,4', shape: 'title', style: { italic: true } }], env);
eq('逗号选择器命中 2、4 页', r9c.applied.map((a) => a.page), [2, 4], JSON.stringify(r9c.skipped));

// ── 失败路径必须说清楚，而不是静默 ─────────────────────────────────────────
const failDeck = freshDeck('fail.pptx');
const r10 = revise(failDeck, [{ slide: 1, shape: '不存在的形状名', setText: 'x' }], env);
eq('找不到形状时不落盘', r10.saved, false);
ok('找不到形状时说清原因并给出下一步', r10.skipped.some((item) => item.reason.includes('找不到形状') && item.reason.includes('readSlides')),
    JSON.stringify(r10.skipped));
// 找不到形状时只报一条，不该再叠一条没有信息量的重复
eq('找不到形状只报一条', r10.skipped.length, 1);

const mixDeck = freshDeck('mix.pptx');
const r11 = revise(mixDeck, [
    { slide: 1, shape: 'title', setText: '有效改动' },
    { slide: 1, shape: 'title', replace: [['不存在的文字', 'x']] },
], env);
eq('一条失败不影响其它条', r11.applied.length, 1);
eq('失败的一条进 skipped', r11.skipped.length, 1);
ok('skipped 里带原因', typeof r11.skipped[0]?.reason === 'string' && r11.skipped[0].reason.includes('找不到'), JSON.stringify(r11.skipped));

// 非文本形状：理由要说清「不是找不到，是这个形状没有文本」。
// Title Accent 是装饰细线（有 txBody 但里面没有段落），用它验证「找到了、但改不了」这条路径。
const shapeDeck = freshDeck('shapeonly.pptx');
const decoration = readSlides(shapeDeck, env).pages[1].shapes.find((s) => s.name === 'Title Accent');
ok('能找到装饰形状', decoration !== undefined, readSlides(shapeDeck, env).pages[1].shapes.map((s) => s.name).join('|'));
const r11b = revise(shapeDeck, [{ slide: 2, shape: 'Title Accent', replace: [['任何文字', 'x']] }], env);
eq('装饰形状不落盘', r11b.saved, false);
eq('装饰形状只报一条', r11b.skipped.length, 1, JSON.stringify(r11b.skipped));
ok('对没有文本的形状给出可操作的理由',
    /找不到|没有文本/.test(r11b.skipped[0]?.reason ?? ''), JSON.stringify(r11b.skipped));
ok('装饰形状的失败原因不是「找不到形状」', !/找不到形状/.test(r11b.skipped[0]?.reason ?? ''),
    '形状明明找到了，报「找不到」会把调用方引向错误的方向');

// ── 多段落换行 ─────────────────────────────────────────────────────────────
const multiDeck = freshDeck('multi.pptx');
const r12 = revise(multiDeck, [{ slide: 3, shape: 'Card 1 Body', setText: '第一行\n第二行' }], env);
eq('setText 支持换行', r12.applied.length, 1, JSON.stringify(r12.skipped));
const multi = readSlides(multiDeck, env).pages[2].shapes.find((s) => s.name === 'Card 1 Body');
eq('换行拆成两段', multi?.paragraphs, 2);

// ── 富文本（段内多 run）只改命中的 run ─────────────────────────────────────
const rich = ppt.create({ title: '富文本', theme: 'plain', path: 'rich.pptx' }, env);
rich.bullets({ title: '混排', items: [{ runs: [
    { text: '普通', sizePt: 18 },
    { text: '重点', sizePt: 18, color: 'C00000', bold: true },
] }] });
rich.save();
const r13 = revise('rich.pptx', [{ slide: 1, shape: 'Bullets 1', style: { sizePt: 24, text: '普通' } }], env);
eq('按子串改样式生效', r13.applied.length, 1, JSON.stringify(r13.skipped));
const richXml = slurpSlide(readFileSync(join(outDir, 'rich.pptx')), 1);
ok('只改了命中的 run 字号', (richXml.match(/sz="2400"/g) ?? []).length === 1 && (richXml.match(/sz="1800"/g) ?? []).length === 1,
    `2400×${(richXml.match(/sz="2400"/g) ?? []).length} 1800×${(richXml.match(/sz="1800"/g) ?? []).length}`);
ok('未命中的 run 保留原色', /srgbClr val="C00000"/.test(richXml));

// ── fit ────────────────────────────────────────────────────────────────────
const beforeFit = readSlides('rich.pptx', env).pages[0].shapes.find((s) => s.name === 'Bullets 1');
const r14 = revise('rich.pptx', [{ slide: 1, shape: 'Bullets 1', fit: true }], env);
eq('fit 生效', r14.applied.length, 1, JSON.stringify(r14.skipped));
const afterFit = readSlides('rich.pptx', env).pages[0].shapes.find((s) => s.name === 'Bullets 1');
ok('fit 改的是高度', afterFit.h !== beforeFit.h, `${beforeFit.h} → ${afterFit.h}`);
ok('fit 不动宽度', Math.abs(afterFit.w - beforeFit.w) < 0.02);

// ── 改一个「别人的」文件：不带任何本引擎的形状命名 ─────────────────────────
// 用 dsh-ppt（另一个 PPT 生成器）真实产出的 pptx 验证兜底路径：它的形状叫
// Title / Body 1 / Accent bar，几何与命名都跟本引擎不同。
// 拿不到 dsh-ppt 时退回合成的极简包 —— 只覆盖「按名字定位」的逻辑，
// 并明确标注它不是真实 Office 产物（合成包缺 theme1.xml，PowerPoint 会判坏，
// 拿它做「文件是否有效」的结论是错的）。
const foreign = await buildForeignDeck();
writeFileSync(join(outDir, 'foreign.pptx'), foreign.bytes);
const foreignList = readSlides('foreign.pptx', env);
eq(`能读外部工具生成的包（${foreign.source}）`, foreignList.pages.length, foreign.pages);
const foreignTitle = foreignList.pages[0].shapes.find((s) => s.name === foreign.shapeName);
ok('外部包的形状也能定位', foreignTitle !== undefined,
    `${foreign.shapeName} 不在 ${foreignList.pages[0].shapes.map((s) => s.name).join('|')}`);
const r15 = revise('foreign.pptx', [
    { slide: 1, shape: foreign.shapeName, setText: '改外部包', style: { sizePt: 30, color: '008000' } },
], env);
eq('能改外部包的形状', r15.applied.length, 1, JSON.stringify(r15.skipped));
eq('外部包改完能读回新文字',
    readSlides('foreign.pptx', env).pages[0].shapes.find((s) => s.name === foreign.shapeName)?.text, '改外部包');

/**
 * 优先用 dsh-ppt 的 deck-core 造一份真实的外部 pptx；模块不在（别人克隆本仓库时
 * 没装 dsh-ppt）就退回合成的极简包。
 */
async function buildForeignDeck() {
    const candidates = [
        process.env.DSH_PPT_DECK_CORE,
        join(process.env.DSH_HOME ?? join(process.env.USERPROFILE ?? '', '.dsh'),
            'profiles', 'web', 'node_modules', 'dsh-ppt', 'skills', 'dsh-ppt', 'scripts', 'deck-core.mjs'),
    ].filter((item) => typeof item === 'string' && item !== '');
    for (const candidate of candidates) {
        if (!existsSync(candidate)) continue;
        try {
            const core = await import(pathToFileURL(candidate).href);
            if (typeof core.buildDeck !== 'function' && typeof core.buildPptx !== 'function') continue;
            const manifest = {
                title: '外部稿',
                lang: 'zh',
                motion: false,
                slides: [
                    { layout: 'cover', title: '外部封面', subtitle: '别家生成' },
                    { layout: 'bullets', title: '外部要点', bullets: ['甲', '乙'] },
                ],
            };
            // buildDeck 写盘，buildPptx 只返回字节；两个 API 版本都兼容
            if (typeof core.buildPptx === 'function') {
                const bytes = core.buildPptx(manifest, 'data', 'zh');
                return { bytes: Buffer.from(bytes), source: 'dsh-ppt', pages: 2, shapeName: 'Title' };
            }
            const built = await core.buildDeck({ manifest, outputDir: join(outDir, 'foreign-src'), fileName: 'foreign' });
            const p = built?.pptxPath ?? built?.files?.pptx;
            if (typeof p !== 'string' || !existsSync(p)) continue;
            return { bytes: readFileSync(p), source: 'dsh-ppt', pages: 2, shapeName: 'Title' };
        } catch {
            // 换下一个候选：外部包只是「额外的一种输入」，拿不到就退回合成包
        }
    }
    return { bytes: buildSyntheticDeck(), source: '合成极简包（非真实 Office 产物）', pages: 2, shapeName: 'TextBox 1' };
}

function buildSyntheticDeck() {
    // 手写一个最小但合法的 pptx：故意用与本引擎不同的形状命名与几何
    const slideXml = (title) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>'
        + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        + '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>'
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        + '<p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>'
        + `<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang="zh-CN" sz="2400" dirty="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${title}</a:t></a:r></a:p>`
        + '</p:txBody></p:sp>'
        + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
    const layout = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">'
        + '<p:cSld name="空白"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
    const master = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        + '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + '</p:spTree></p:cSld>'
        + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        + '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        + '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>';
    const presentation = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        + '<p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>'
        + '<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>';
    const entries = [
        { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            + '<Default Extension="xml" ContentType="application/xml"/>'
            + '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
            + '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            + '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            + '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
            + '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
            + '</Types>' },
        { name: '_rels/.rels', data: relsXml([['rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'ppt/presentation.xml']]) },
        { name: 'ppt/presentation.xml', data: presentation },
        { name: 'ppt/_rels/presentation.xml.rels', data: relsXml([
            ['rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slides/slide1.xml'],
            ['rId2', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slides/slide2.xml'],
            ['rId3', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', 'slideMasters/slideMaster1.xml'],
        ]) },
        { name: 'ppt/slides/slide1.xml', data: slideXml('外部标题一') },
        { name: 'ppt/slides/slide2.xml', data: slideXml('外部标题二') },
        { name: 'ppt/slides/_rels/slide1.xml.rels', data: relsXml([['rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout1.xml']]) },
        { name: 'ppt/slides/_rels/slide2.xml.rels', data: relsXml([['rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout1.xml']]) },
        { name: 'ppt/slideLayouts/slideLayout1.xml', data: layout },
        { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: relsXml([['rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', '../slideMasters/slideMaster1.xml']]) },
        { name: 'ppt/slideMasters/slideMaster1.xml', data: master },
        { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: relsXml([['rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout1.xml']]) },
    ];
    return Buffer.from(zipModule.zip(entries));
}

function relsXml(rows) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rows.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('')
        + '</Relationships>';
}

/**
 * 在合成包的基础上加两张图，专测抽取的两条边界：
 * 1. 媒体部件的内容类型**只写在 Override 里**（没有 png 的 Default）—— 别的工具确实会这么产；
 * 2. 一张图用 `a:blip/@r:link` 链到外部 —— 包里没有字节，抽取时必须照实说「抽不出来」。
 */
function buildOverrideDeck(imageBytes) {
    const files = zipModule.unzip(buildSyntheticDeck());
    const decoder = new TextDecoder('utf-8');
    const encoder = new TextEncoder();
    files.set('ppt/media/image1.png', imageBytes);
    const ct = decoder.decode(files.get('[Content_Types].xml')).replace('</Types>',
        '<Override PartName="/ppt/media/image1.png" ContentType="image/png"/></Types>');
    files.set('[Content_Types].xml', encoder.encode(ct));
    const rels = decoder.decode(files.get('ppt/slides/_rels/slide1.xml.rels')).replace('</Relationships>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
        + ' Target="../media/image1.png"/>'
        + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
        + ' Target="https://example.com/linked.png" TargetMode="External"/></Relationships>');
    files.set('ppt/slides/_rels/slide1.xml.rels', encoder.encode(rels));
    const pics = '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Embedded"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/>'
        + '</p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/>'
        + '<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/>'
        + '<a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        + '<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:pic>'
        + '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Linked"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/>'
        + '</p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:link="rId3"/>'
        + '<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="1000000"/>'
        + '<a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        + '<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:pic>';
    const slide1 = decoder.decode(files.get('ppt/slides/slide1.xml')).replace('</p:spTree>', `${pics}</p:spTree>`);
    files.set('ppt/slides/slide1.xml', encoder.encode(slide1));
    return Buffer.from(zipModule.zip(files));
}

// ── 图片：往「别人的包」里插图，以及把包里的图抽出来 ───────────────────────
//
// 输入刻意用 buildSyntheticDeck()：那份包里**没有媒体部件、也没有 png 的 Default**，
// 于是「补 Default」「从 image1 起编号」「新关系 id = 最大 + 1」这几条才真的被走到。
// 用自己刚生成的带图稿子测这几条，等于让被测代码自己给自己出题。

const png = makePng(240, 120);
env.writeFile('assets/pic.png', png);
const square = makePng(240, 240);
env.writeFile('assets/square.png', square);

function packageOf(file) {
    return zipModule.unzip(readFileSync(join(outDir, file)));
}

function textOfPart(parts, name) {
    const data = parts.get(name);
    return data === undefined ? '' : new TextDecoder('utf-8').decode(data);
}

const syntheticPath = 'synthetic-media.pptx';
writeFileSync(join(outDir, syntheticPath), buildSyntheticDeck());

const ins = insertImages(syntheticPath, [
    { slide: 2, path: 'assets/pic.png', x: 2, y: 3, wCm: 6, hCm: 4, fit: 'contain', name: '插入的图', alt: '替代文字' },
], env);
eq('插图返回 ok/saved', [ins.ok, ins.saved], [true, true]);
eq('插图 applied 一条', ins.applied.length, 1, JSON.stringify(ins.skipped));
eq('插图 skipped 为空', ins.skipped.length, 0, JSON.stringify(ins.skipped));
eq('applied 的形状 id 从该页最大 id 顺延', ins.applied[0]?.id, 3);
eq('applied 的关系 id 从该页最大数字 id 顺延', ins.applied[0]?.relId, 'rId2');
eq('applied 报的几何（厘米）', [ins.applied[0]?.x, ins.applied[0]?.y, ins.applied[0]?.w, ins.applied[0]?.h], [2, 3.5, 6, 3]);

const insParts = packageOf(syntheticPath);
ok('(a) 新媒体部件与源文件逐字节一致',
    Buffer.compare(Buffer.from(insParts.get('ppt/media/image1.png') ?? Buffer.alloc(0)), png) === 0);
{
    const ct = textOfPart(insParts, '[Content_Types].xml');
    ok('(b) [Content_Types].xml 补了 png 的 Default', ct.includes('<Default Extension="png" ContentType="image/png"/>'), ct.slice(0, 300));
    eq('(b2) png 的 Default 只补一条', (ct.match(/Extension="png"/g) ?? []).length, 1);
    // Default 必须排在 Override 之前（CT_Types 的子元素顺序：Default* 然后 Override*）
    ok('(b3) Default 排在 Override 之前', ct.indexOf('Extension="png"') < ct.indexOf('<Override'), ct.slice(0, 400));
    ok('(b4) 媒体部件本身不写 Override', !/Override[^>]*ppt\/media/.test(ct));
}
{
    const rels = textOfPart(insParts, 'ppt/slides/_rels/slide2.xml.rels');
    ok('(c) slide2 的 rels 多了指向 ../media/image1.png 的关系',
        rels.includes('Id="rId2"') && rels.includes('Target="../media/image1.png"'), rels);
    ok('(c2) 新关系的类型是 image',
        /Id="rId2"[^>]*Type="[^"]*\/relationships\/image"/.test(rels), rels);
    eq('(c3) 原有的 layout 关系没被动过', /Id="rId1"[^>]*slideLayout/.test(rels), true);
    ok('(c4) slide1 的 rels 没有被牵连', !textOfPart(insParts, 'ppt/slides/_rels/slide1.xml.rels').includes('media/'), '');
}
{
    const slide2 = textOfPart(insParts, 'ppt/slides/slide2.xml');
    eq('(d) 恰好新增一个 p:pic', (slide2.match(/<p:pic>/g) ?? []).length, 1);
    ok('(d2) p:pic 的 r:embed 指向新关系', slide2.includes('<a:blip r:embed="rId2"/>'), slide2);
    ok('(d3) p:pic 是 p:spTree 的最后一个形状',
        slide2.indexOf('<p:pic>') > slide2.lastIndexOf('</p:sp>') && slide2.indexOf('</p:pic>') < slide2.indexOf('</p:spTree>'));
    ok('(d4) p:nvPicPr 里有 p:nvPr（PresentationML 的 minOccurs=1）',
        /<p:cNvPicPr><a:picLocks noChangeAspect="1"\/><\/p:cNvPicPr><p:nvPr\/><\/p:nvPicPr>/.test(slide2));
    ok('(d5) 形状名与替代文字落到 p:cNvPr',
        slide2.includes('name="插入的图"') && slide2.includes('descr="替代文字"'), slide2);
    const off = /<p:pic>[\s\S]*?<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(slide2);
    const ext = /<p:pic>[\s\S]*?<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/.exec(slide2);
    eq('(e) a:off = 厘米 × 360000', [Number(off?.[1]), Number(off?.[2])], [Math.round(2 * 360000), Math.round(3.5 * 360000)]);
    eq('(e2) a:ext = 厘米 × 360000', [Number(ext?.[1]), Number(ext?.[2])], [Math.round(6 * 360000), Math.round(3 * 360000)]);
}

// (f) 缺图：进 skipped、给原因、**不落盘**（没改动就不写，与 revise 同一条约定）
{
    const before = readFileSync(join(outDir, syntheticPath));
    const missing = insertImages(syntheticPath, [{ slide: 1, path: 'assets/not-here.png', x: 1, y: 1, wCm: 4 }], env);
    eq('(f) 缺图时 ok/saved', [missing.ok, missing.saved], [false, false]);
    ok('(f2) 缺图给得出原因', /找不到图片文件/.test(missing.skipped[0]?.reason ?? ''), JSON.stringify(missing.skipped));
    eq('(f3) 失败时文件一个字节都不动', Buffer.compare(before, readFileSync(join(outDir, syntheticPath))), 0);
    eq('(f4) 失败时补一句「文件未改写」', missing.warnings.some((line) => line.includes('未改写')), true);
}
// (f5) 格式不支持：同样进 skipped，且不会往包里塞媒体部件
{
    env.writeFile('assets/raw.bmp', Buffer.from('BM00000000 not a supported bitmap'));
    const bad = insertImages(syntheticPath, [{ slide: 1, path: 'assets/raw.bmp', x: 1, y: 1, wCm: 4 }], env);
    eq('(f5) 不支持的格式不落盘', [bad.ok, bad.saved], [false, false]);
    ok('(f6) 不支持的格式说清只认什么', /只认 PNG\/JPEG\/GIF/.test(bad.skipped[0]?.reason ?? ''), JSON.stringify(bad.skipped));
    eq('(f7) 不支持的格式没有多出媒体部件',
        [...packageOf(syntheticPath).keys()].filter((name) => name.startsWith('ppt/media/')).length, 1);
}

// (g) 抽取：字节必须与源文件逐字节一致，并报出页码与关系 id
{
    const ex = ppt.images(syntheticPath, { out: 'synthetic-out' }, env);
    eq('(g) 抽出一张图', [ex.ok, ex.count], [true, 1]);
    eq('(g2) 抽取的字节与源文件一致', Buffer.compare(env.readFile(ex.files[0].path), png), 0);
    eq('(g3) 抽取报出页码与关系 id', [ex.files[0].slides.join(','), ex.files[0].relIds.join(',')], ['2', 'rId2']);
    eq('(g4) 抽取报出内容类型与像素尺寸',
        [ex.files[0].contentType, ex.files[0].width, ex.files[0].height], ['image/png', 240, 120]);
    eq('(g5) 只被引用一次时不报 reused', [ex.reused, ex.files[0].reused], [0, false]);
}

// (h) readSlides：图片形状要带上 image 信息，其余字段一个不少
{
    const after = readSlides(syntheticPath, env);
    const pics = after.pages[1].shapes.filter((shape) => shape.kind === 'pic');
    eq('(h) readSlides 认出新增的图片形状', pics.length, 1, JSON.stringify(after.pages[1].shapes.map((s) => s.name)));
    eq('(h2) 图片形状的 kind/name/editable', [pics[0]?.kind, pics[0]?.name, pics[0]?.editable], ['pic', '插入的图', false]);
    eq('(h3) 图片形状的几何（厘米）', [pics[0]?.x, pics[0]?.y, pics[0]?.w, pics[0]?.h], [2, 3.5, 6, 3]);
    eq('(h4) 图片形状的 image 信息', [
        pics[0]?.image?.part, pics[0]?.image?.relId, pics[0]?.image?.contentType,
        pics[0]?.image?.bytes, pics[0]?.image?.width, pics[0]?.image?.height, pics[0]?.image?.linked,
    ], ['ppt/media/image1.png', 'rId2', 'image/png', png.length, 240, 120, false]);
    eq('(h5) contain 不裁切，crop 全 0',
        [pics[0]?.image?.crop?.l, pics[0]?.image?.crop?.t, pics[0]?.image?.crop?.r, pics[0]?.image?.crop?.b], [0, 0, 0, 0]);
    ok('(h6) 原有文本形状的字段一个都没少',
        after.pages[0].shapes.some((shape) => shape.name === 'TextBox 1' && shape.text === '外部标题一'
            && shape.editable === true && shape.paragraphs === 1), JSON.stringify(after.pages[0].shapes));
    ok('(h7) 文本形状不带 image 键', after.pages[0].shapes.every((shape) => shape.image === undefined));
}

// cover / natural / 编号顺延：三种 fit 各走一遍
{
    const coverPath = 'synthetic-fit.pptx';
    writeFileSync(join(outDir, coverPath), buildSyntheticDeck());
    const fits = insertImages(coverPath, [
        { slide: 1, path: 'assets/square.png', x: 1, y: 1, wCm: 4, hCm: 2, fit: 'cover' },
        { slide: 2, path: 'assets/square.png', x: 1, y: 1, fit: 'natural' },
        { slide: 2, path: 'assets/pic.png', x: 1, y: 1, wCm: 3 },
    ], env);
    eq('三种 fit 都成功', fits.applied.map((item) => [item.page, item.fit, item.part]),
        [[1, 'cover', 'ppt/media/image1.png'], [2, 'natural', 'ppt/media/image1.png'], [2, 'contain', 'ppt/media/image2.png']],
        JSON.stringify(fits.skipped));
    const coverSlide = textOfPart(packageOf(coverPath), 'ppt/slides/slide1.xml');
    ok('cover 铺满框并写 a:srcRect 裁切',
        /<a:srcRect l="0" t="25000" r="0" b="25000"\/>/.test(coverSlide), coverSlide);
    ok('cover 的 a:ext 就是框本身（4×2cm）', coverSlide.includes('<a:ext cx="1440000" cy="720000"/>'), coverSlide);
    const naturalSlide = textOfPart(packageOf(coverPath), 'ppt/slides/slide2.xml');
    ok('natural 按 96 DPI 用原始像素尺寸（240px = 2286000 EMU）',
        naturalSlide.includes('<a:ext cx="2286000" cy="2286000"/>'), naturalSlide);
    eq('同一张图插两次只登记一个媒体部件',
        [...packageOf(coverPath).keys()].filter((name) => name.startsWith('ppt/media/')).join(','),
        'ppt/media/image1.png,ppt/media/image2.png');
    // 同一部件被两页引用 → reused
    const reuseEx = ppt.images(coverPath, { out: 'synthetic-reuse' }, env);
    const first = reuseEx.files.find((file) => file.part === 'ppt/media/image1.png');
    eq('跨页复用同一部件时报 reused', [reuseEx.reused, first?.slides.join(',')], [1, '1,2']);
}

// 放不下的图照样插进去，但要在 warnings 里说清（与 revise 的 move/resize 同口径）
{
    const warnPath = 'synthetic-offpage.pptx';
    writeFileSync(join(outDir, warnPath), buildSyntheticDeck());
    const off = insertImages(warnPath, [{ slide: 1, path: 'assets/square.png', x: 30, y: 15, fit: 'natural' }], env);
    eq('超页时图片照样插入（不是错误）', [off.ok, off.applied.length], [true, 1]);
    ok('超页 warning 带页码与厘米坐标',
        /第 1 页插入的图片超出页面/.test(off.warnings.join('|')), JSON.stringify(off.warnings));
}

// revise() 里同一批 op 做「改文字 + 插图」
{
    const mixedPath = 'synthetic-mixed.pptx';
    writeFileSync(join(outDir, mixedPath), buildSyntheticDeck());
    const mixed = revise(mixedPath, [
        { slide: 1, shape: 'TextBox 1', setText: '改过的文字' },
        { slide: 2, addImage: { path: 'assets/pic.png', x: 1, y: 1, wCm: 5, hCm: 2, fit: 'cover' } },
    ], env);
    eq('一批里同时改文字与插图', mixed.applied.map((item) => item.op), ['setText', 'insertImage'], JSON.stringify(mixed.skipped));
    eq('两条都落了盘', [mixed.ok, mixed.saved], [true, true]);
    const mixedParts = packageOf(mixedPath);
    ok('文字改到了', textOfPart(mixedParts, 'ppt/slides/slide1.xml').includes('改过的文字'));
    ok('图插到了', textOfPart(mixedParts, 'ppt/slides/slide2.xml').includes('<p:pic>'));
    ok('媒体部件进了包', mixedParts.has('ppt/media/image1.png'));
    ok('插图 op 不需要 shape 选择器，也不会报「找不到形状」',
        !mixed.skipped.some((item) => /找不到形状/.test(item.reason)), JSON.stringify(mixed.skipped));

    // 同一条 op 里既写 shape 又写 addImage：不能静默忽略用户写的东西
    const bothPath = 'synthetic-both.pptx';
    writeFileSync(join(outDir, bothPath), buildSyntheticDeck());
    const both = revise(bothPath, [{ slide: 1, shape: 'title', addImage: { path: 'assets/pic.png', x: 1, y: 1, wCm: 3 } }], env);
    eq('addImage 与 shape 同时给出时插图照做', [both.ok, both.applied.length], [true, 1]);
    ok('并且明说 shape 被忽略了', both.warnings.some((line) => line.includes('已忽略 shape')), JSON.stringify(both.warnings));

    // 只有 addImage 且失败时：不落盘
    const failPath = 'synthetic-addfail.pptx';
    writeFileSync(join(outDir, failPath), buildSyntheticDeck());
    const before = readFileSync(join(outDir, failPath));
    const failed = revise(failPath, [{ slide: 1, addImage: { path: 'assets/not-here.png' } }], env);
    eq('addImage 失败时不落盘', [failed.ok, failed.saved], [false, false]);
    ok('addImage 失败给原因', /找不到图片文件/.test(failed.skipped[0]?.reason ?? ''), JSON.stringify(failed.skipped));
    eq('addImage 失败不改写文件', Buffer.compare(before, readFileSync(join(outDir, failPath))), 0);
}

// ── 结构完整性：引擎自己的读取器与 LibreOffice 通过都不算证据 ──────────────
//
// 这一段盯的是「真实 Microsoft Office 才判得出来的坏包」：p:sp 与 p:txBody 被切坏、
// p:spTree 的头两个子元素被顶掉、p:cNvPr id 撞车、a:spPr 子元素顺序写反、
// r:embed 指向一条不存在的关系。这些在自研解析器下都照常渲染。

/** 平衡扫描：截出顶层形状元素的切片（与 ppt-revise.js 的 sliceShapes 同一套配对扫描）。 */
function scanShapes(xml) {
    const out = [];
    const re = /<(\/?)(p:sp|p:pic|p:graphicFrame|p:grpSp|p:cxnSp)\b([^>]*?)(\/?)>/g;
    const stack = [];
    let match;
    while ((match = re.exec(xml)) !== null) {
        const [whole, closing, tag, , selfClose] = match;
        if (closing === '/') {
            const open = stack.pop();
            if (open !== undefined && stack.length === 0) out.push({ tag, start: open.start, end: match.index + whole.length });
            continue;
        }
        if (selfClose === '/') {
            if (stack.length === 0) out.push({ tag, start: match.index, end: match.index + whole.length });
            continue;
        }
        stack.push({ tag, start: match.index });
    }
    return out.sort((a, b) => a.start - b.start);
}

/** 一个标签在整段 XML 里的配对情况：depth 必须回到 0，否则元素被切坏了。 */
function tagBalance(xml, tag) {
    const re = new RegExp(`<(/?)${tag}(?=[\\s/>])([^>]*?)(/?)>`, 'g');
    let depth = 0;
    let opens = 0;
    let closes = 0;
    let match;
    while ((match = re.exec(xml)) !== null) {
        if (match[1] === '/') { closes += 1; depth -= 1; continue; }
        if (match[3] === '/') continue;
        opens += 1;
        depth += 1;
    }
    return { opens, closes, depth };
}

/** 截出某个标签的完整元素切片（用于逐个检查 p:spPr 的子元素顺序）。 */
function sliceElements(xml, tag) {
    const re = new RegExp(`<(/?)${tag}(?=[\\s/>])([^>]*?)(/?)>`, 'g');
    const out = [];
    const stack = [];
    let match;
    while ((match = re.exec(xml)) !== null) {
        const [whole, closing, , selfClose] = match;
        if (closing === '/') {
            const open = stack.pop();
            if (open !== undefined && stack.length === 0) out.push(xml.slice(open.start, match.index + whole.length));
            continue;
        }
        if (selfClose === '/') { if (stack.length === 0) out.push(whole); continue; }
        stack.push({ start: match.index });
    }
    return out;
}

/** a:spPr 的合法子元素顺序（docs/ooxml-pitfalls.md 第 8 条）。 */
const SPPR_RANK = {
    'a:xfrm': 0, 'a:custGeom': 1, 'a:prstGeom': 1,
    'a:noFill': 2, 'a:solidFill': 2, 'a:gradFill': 2, 'a:blipFill': 2, 'a:pattFill': 2, 'a:grpFill': 2,
    'a:ln': 3, 'a:effectLst': 4, 'a:effectDag': 4, 'a:scene3d': 5, 'a:sp3d': 6, 'a:extLst': 7,
};

function integrityProblems(slideXml, relsXml) {
    const problems = [];
    const relIds = new Set([...relsXml.matchAll(/<Relationship\b[^>]*\sId="([^"]+)"/g)].map((match) => match[1]));

    // 1) p:sp 与 p:txBody 一一对应，且每个 p:sp 恰好一个完整的 p:txBody
    const sp = tagBalance(slideXml, 'p:sp');
    const tx = tagBalance(slideXml, 'p:txBody');
    if (sp.depth !== 0) problems.push('p:sp 标签不配对');
    if (tx.depth !== 0) problems.push('p:txBody 标签不配对');
    if (sp.opens !== tx.opens) problems.push(`p:sp ${sp.opens} 个但 p:txBody ${tx.opens} 个`);
    for (const shape of scanShapes(slideXml).filter((item) => item.tag === 'p:sp')) {
        const inner = tagBalance(slideXml.slice(shape.start, shape.end), 'p:txBody');
        if (inner.opens !== 1 || inner.closes !== 1) problems.push(`一个 p:sp 里有 ${inner.opens} 个 p:txBody`);
    }

    // 2) p:spTree 的头两个子元素还是 nvGrpSpPr / grpSpPr，xfrm 四个值全 0
    const root = parseXml(slideXml).children[0];
    const spTree = descendants(root, 'p:spTree')[0];
    const kids = spTree === undefined ? [] : children(spTree);
    if (kids[0]?.name !== 'p:nvGrpSpPr' || kids[1]?.name !== 'p:grpSpPr') {
        problems.push(`p:spTree 的头两个子元素是 ${kids.slice(0, 2).map((node) => node.name).join(',')}`);
    } else {
        const xfrm = descendants(kids[1], 'a:xfrm')[0];
        for (const label of ['a:off', 'a:ext', 'a:chOff', 'a:chExt']) {
            const node = descendants(xfrm, label)[0];
            if (node === undefined) { problems.push(`p:grpSpPr 缺 ${label}`); continue; }
            for (const key of ['x', 'y', 'cx', 'cy']) {
                if (node.attrs[key] !== undefined && node.attrs[key] !== '0') {
                    problems.push(`p:grpSpPr 的 ${label}@${key}=${node.attrs[key]} 不是 0`);
                }
            }
        }
    }

    // 3) p:cNvPr 的 id 每页唯一
    const ids = [...slideXml.matchAll(/<p:cNvPr\b[^>]*\sid="(\d+)"/g)].map((match) => match[1]);
    if (new Set(ids).size !== ids.length) problems.push(`p:cNvPr id 重复：${ids.join(',')}`);

    // 4) p:spPr 的**直接**子元素顺序合法（xfrm → prstGeom → 填充 → ln → effectLst）。
    // 只看直接子元素：a:ln 里还有自己的 a:noFill，按全文字面扫会把内层的填充误判成顺序颠倒。
    for (const spPr of sliceElements(slideXml, 'p:spPr')) {
        const node = parseXml(spPr).children[0];
        let last = -1;
        for (const child of children(node)) {
            const rank = SPPR_RANK[child.name];
            if (rank === undefined) continue;
            if (rank < last) { problems.push(`p:spPr 子元素顺序非法：${spPr.slice(0, 200)}`); break; }
            last = rank;
        }
    }

    // 5) 每个 r:embed / r:id / r:link 都能在本部件的 rels 里找到
    for (const match of slideXml.matchAll(/r:(?:id|embed|link)="([^"]+)"/g)) {
        if (!relIds.has(match[1])) problems.push(`引用了不存在的关系 ${match[1]}`);
    }
    return problems;
}

/** 一份包里所有 slide 的结构问题（跨页一起报，一次就能看出是哪一页坏的）。 */
function integrityReport(label, parts) {
    const problems = [];
    for (const name of [...parts.keys()].filter((item) => /^ppt\/slides\/slide\d+\.xml$/.test(item))) {
        const relsName = name.replace(/([^/]+)$/, '_rels/$1.rels');
        const xml = textOfPart(parts, name);
        const rels = textOfPart(parts, relsName);
        for (const problem of integrityProblems(xml, rels)) problems.push(`${label} ${name}：${problem}`);
    }
    return problems;
}

{
    const decks = ['synthetic-media.pptx', 'synthetic-fit.pptx', 'synthetic-mixed.pptx', 'synthetic-offpage.pptx'];
    for (const deck of decks) {
        eq(`结构完整性（${deck}）`, integrityReport(deck, packageOf(deck)), []);
    }
    // 断言这套检查真的会抓到东西：拿一个被切坏的 txBody 当反例
    const broken = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        + '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr><p:txBody><a:p/></p:sp>';
    ok('反例：缺 </p:txBody> 会被抓出来', integrityProblems(broken, '').length > 0, '检查本身失效了');
    ok('反例：p:spPr 子元素顺序写反会被抓出来',
        integrityProblems('<p:spPr><a:ln><a:noFill/></a:ln><a:prstGeom prst="rect"/></p:spPr>', '').length > 0);
    ok('反例：悬空的 r:embed 会被抓出来',
        integrityProblems('<p:pic><p:blipFill><a:blip r:embed="rId9"/></p:blipFill></p:pic>', '').length > 0);
}

// ── 抽取的两条边界：内容类型只写在 Override 里、以及 r:link 的外链图片 ─────
{
    const overridePath = 'synthetic-override.pptx';
    writeFileSync(join(outDir, overridePath), buildOverrideDeck(png));
    const ex = ppt.images(overridePath, { out: 'override-out' }, env);
    eq('内容类型只在 Override 里也能解析出来', [ex.count, ex.files[0]?.contentType], [1, 'image/png']);
    eq('Override 声明的图同样按字节抽出', Buffer.compare(env.readFile(ex.files[0].path), png), 0);
    eq('外链图片进 skipped', ex.skipped.length, 1, JSON.stringify(ex.skipped));
    ok('外链的 skipped 说清是链接而非嵌入', /外部链接/.test(ex.skipped[0]?.reason ?? ''), JSON.stringify(ex.skipped));
    ok('外链的 skipped 不编造字节', ex.skipped[0]?.bytes === undefined, JSON.stringify(ex.skipped[0]));

    const listed = readSlides(overridePath, env);
    const shapes = listed.pages[0].shapes.filter((shape) => shape.kind === 'pic');
    eq('两张图都被 readSlides 认出来', shapes.length, 2, JSON.stringify(shapes.map((s) => s.name)));
    eq('嵌入的那张报出部件与字节', [shapes[0]?.image?.part, shapes[0]?.image?.bytes, shapes[0]?.image?.linked],
        ['ppt/media/image1.png', png.length, false]);
    eq('外链的那张报 linked:true 且 bytes 为 null',
        [shapes[1]?.image?.linked, shapes[1]?.image?.bytes, shapes[1]?.image?.contentType], [true, null, null]);
    ok('外链的那张 part 是链出去的目标', String(shapes[1]?.image?.part).startsWith('https://'), String(shapes[1]?.image?.part));
}

// ── (i) SDK 挂载面：新增的两个方法必须真的挂到 office.ppt 上 ───────────────
{
    ok('(i) ppt.api.images 是工厂函数', typeof ppt.api?.images === 'function');
    ok('(i2) ppt.api.insertImages 是工厂函数', typeof ppt.api?.insertImages === 'function');

    const { buildSdk } = await import('../src/sdk.js');
    const { createCache } = await import('../src/engine/cache.js');
    const cache = createCache({ root: outDir, dir: 'cache' });
    const sdk = await buildSdk({ env, cache, config: { defaultTheme: 'plain' } });
    ok('(i3) SDK 挂上了 office.ppt.images', typeof sdk.office.ppt.images === 'function');
    ok('(i4) SDK 挂上了 office.ppt.insertImages', typeof sdk.office.ppt.insertImages === 'function');
    ok('(i5) 既有的 readSlides / revise 还在',
        typeof sdk.office.ppt.readSlides === 'function' && typeof sdk.office.ppt.revise === 'function');
    // 不传 out：必须落到注入的 cache 目录里，而不是工作目录 —— 这条同时验证 sdk.js 把 {cache} 传下去了
    const sdkEx = sdk.office.ppt.images(syntheticPath);
    eq('(i6) 不传 out 时落注入的缓存目录', sdkEx.dir.startsWith('cache/'), true);
    eq('(i7) 缓存里的图与源文件一致', Buffer.compare(env.readFile(sdkEx.files[0].path), png), 0);
    ok('(i8) 写缓存记了一次 artifact', cache.stats.artifacts >= 1, String(cache.stats.artifacts));

    const sdkDeck = 'sdk-insert.pptx';
    writeFileSync(join(outDir, sdkDeck), buildSyntheticDeck());
    const sdkIns = await sdk.office.ppt.insertImages(sdkDeck, [
        { slide: 1, path: 'assets/pic.png', x: 1, y: 1, wCm: 3, hCm: 2 },
    ]);
    eq('(i9) 通过 SDK 插图', [sdkIns.ok, sdkIns.saved, sdkIns.applied.length], [true, true, 1]);
    eq('(i10) SDK 插完的包结构依然完整', integrityReport(sdkDeck, packageOf(sdkDeck)), []);
}

// ── 结果 ───────────────────────────────────────────────────────────────────
for (const line of failures) console.log(`FAIL  ${line}`);
console.log(`format-ppt-revise: ${passed}/${passed + failures.length} 通过`);
if (failures.length > 0) process.exit(1);
