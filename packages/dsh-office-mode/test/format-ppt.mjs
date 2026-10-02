#!/usr/bin/env node
/**
 * PPT 引擎自测（SPEC §5）。
 *
 * 覆盖：全部 builder 方法 → render() 出字节 → unzip 回读 → 逐个 XML/rels 解析 →
 * 关系目标存在性 → read() 复检 stats → 排版警告的真实性。
 *
 * 产物写在仓库根的 .office/tmp/ppt 下，默认结束时删掉；加 --keep 可以留着做人工/COM 检查。
 *
 *   node test/format-ppt.mjs
 *   node test/format-ppt.mjs --keep
 */
import { deflateSync } from 'node:zlib';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { create, read, images, meta } from '../src/formats/ppt.js';
import { createEnv, cmToEmu } from '../src/engine/kit.js';
import { DEFAULT_THEME_ID, resolveTheme } from '../src/engine/theme.js';
import { unzip, crc32 } from '../src/engine/zip.js';
import { parseXml, children as xmlChildren, descendants } from '../src/engine/xml.js';

const HERE = dirname(fileURLToPath(import.meta.url));
// 测试中间产物统一落在**仓库根**的 .office/tmp/<suite>/ 下（不再用包内的 test/.tmp）。
// 只用 ppt/ 这个子目录：三个格式的测试会同时跑，
// 谁都不许对共享的 .office/tmp 做递归删除。
const TMP = resolve(HERE, '..', '..', '..', '.office', 'tmp', 'ppt');
const KEEP = process.argv.includes('--keep');

const failures = [];
let checks = 0;

function assert(condition, message) {
    checks += 1;
    if (!condition) failures.push(message);
    return condition === true;
}

function equal(actual, expected, message) {
    return assert(actual === expected, `${message}（期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}）`);
}

function fail(message) {
    checks += 1;
    failures.push(message);
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
        out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
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

// ── 造一张真 JPEG（最简基线 JPEG：1 个灰度分量、全 0 系数） ───────────────
//
// 为什么不用「随便一段带 FFD8 开头的字节」：那只能骗过自家探测器，真实 PowerPoint 与
// System.Drawing 都会判它不是图片。这一份是能真正解码的 8×8 灰度 JPEG：
// 两个 1 位码的 Huffman 表（DC 只含 0 类、AC 只含 EOB）配上全 1 的量化表，
// 一个 8×8 块正好一个 MCU，熵数据就是「DC 0 + EOB」两位加填充。
// 已经用 System.Drawing 验过：能加载，尺寸 8×8。
function makeJpeg(width, height) {
    const counts = Buffer.alloc(16);
    counts[0] = 1; // 长度为 1 的码只有一条
    const dqt = Buffer.alloc(65);
    dqt[0] = 0x00; // Pq=0（8 位）、Tq=0
    dqt.fill(0x01, 1); // 全 1 量化表：合法，且不会放大任何误差
    const sof = Buffer.alloc(8);
    sof[0] = 0x08; // 精度 8 位
    sof.writeUInt16BE(height, 1);
    sof.writeUInt16BE(width, 3);
    sof[5] = 0x01; // 分量数 1
    sof[6] = 0x01; // 分量 id
    sof[7] = 0x11; // 1×1 采样、量化表 0
    return Buffer.concat([
        Buffer.from([0xff, 0xd8]),
        Buffer.from([0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00]),
        Buffer.from([0xff, 0xdb, 0x00, 0x43]), dqt,
        Buffer.from([0xff, 0xc0, 0x00, 0x0b]), sof, Buffer.from([0x00]),
        Buffer.from([0xff, 0xc4, 0x00, 0x14, 0x00]), counts, Buffer.from([0x00]), // DC 表：码 0 → 类别 0
        Buffer.from([0xff, 0xc4, 0x00, 0x14, 0x10]), counts, Buffer.from([0x00]), // AC 表：码 0 → EOB
        Buffer.from([0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f, 0x00, 0x3f]),
        Buffer.from([0xff, 0xd9]),
    ]);
}

// ── 包结构工具 ────────────────────────────────────────────────────────────

function ownerOf(relsPath) {
    const at = relsPath.lastIndexOf('/_rels/');
    if (at === -1) return relsPath === '_rels/.rels' ? '' : undefined;
    return `${relsPath.slice(0, at)}/${relsPath.slice(at + 7).replace(/\.rels$/, '')}`;
}

function relsPathOf(partName) {
    const at = partName.lastIndexOf('/');
    const dir = at === -1 ? '' : partName.slice(0, at + 1);
    const base = at === -1 ? partName : partName.slice(at + 1);
    return `${dir}_rels/${base}.rels`;
}

function resolveTarget(owner, target) {
    if (target.startsWith('/')) return target.slice(1);
    const segments = owner === '' ? [] : owner.split('/').slice(0, -1);
    for (const piece of target.split('/')) {
        if (piece === '' || piece === '.') continue;
        if (piece === '..') segments.pop();
        else segments.push(piece);
    }
    return segments.join('/');
}

/**
 * 只实现 images() 会用到的那几个方法的缓存替身。
 *
 * 为什么不直接用 engine/cache.js：这里要验的是「images() 走的是**注入的** cache 对象、
 * 而不是某条写死的路径」。替身一换，硬编码立刻暴露成 dir 不对。
 */
const fakeCache = (() => {
    const root = resolve(TMP, 'fake-cache');
    const api = {
        artifacts: 0,
        dir: root,
        rel: 'fake-cache',
        ensureDir(name) {
            const target = resolve(root, name);
            mkdirSync(target, { recursive: true });
            return target;
        },
        noteArtifact() {
            api.artifacts += 1;
        },
        write(name, data) {
            const target = api.ensureDir('.');
            const file = resolve(target, name);
            writeFileSync(file, data);
            api.artifacts += 1;
            return { path: file, absolute: file, bytes: Buffer.byteLength(data) };
        },
    };
    return api;
})();

// ── 主样例：覆盖全部 builder 方法 ─────────────────────────────────────────

const env = createEnv({ root: TMP, themeResolver: resolveTheme });
const png = makePng(1200, 675);
env.writeFile('assets/chart.png', png);

const deck = create({
    title: '2026 年产品规划',
    theme: 'tech',
    path: 'demo/2026-产品规划.pptx',
    size: '16:9',
    author: '产品委员会',
    date: '2026-01-15',
}, env);

deck.cover({
    title: '2026 年产品规划',
    subtitle: '把「能用」做成「好用」：三条产品线与一个平台底座',
    kicker: '年度战略汇报',
    presenter: '产品委员会',
    date: '2026-01-15',
});
deck.section({ title: '第一部分：现状与判断', subtitle: '过去一年我们做对了什么、错过了什么' });
deck.bullets({
    title: '三条产品线的成绩与短板',
    items: [
        { text: '办公模式：从 0 到 1 完成 Word / Excel / PPT 三件套的零依赖引擎', level: 1 },
        { text: '文档引擎已能通过真实 Office 打开校验', level: 2 },
        { text: '表格引擎覆盖公式、冻结窗格与自动列宽', level: 2 },
        { text: '协作模式：多人批注仍依赖第三方云盘', level: 1 },
        { text: '搜索模式：垂直检索的召回率只有 62%', level: 1 },
    ],
});
deck.bullets({
    title: '明年的四个关键动作',
    columns: 2,
    items: [
        { text: '统一文档模型', level: 1 },
        { text: '把三套 OOXML 渲染收敛成一套版式语言', level: 2 },
        { text: '把溢出风险前置到生成端', level: 2 },
        { text: '开放插件协议', level: 1 },
        { text: '第三方可以注册新的格式引擎', level: 2 },
        { text: '降低集成成本', level: 1 },
        { text: '一次调用完成批量产出', level: 2 },
    ],
});
deck.table({
    title: '三条产品线的季度目标',
    columns: [
        { title: '产品线', width: 14 },
        { title: 'Q1 目标', type: 'number' },
        { title: 'Q2 目标', type: 'number' },
        { title: '负责人', width: 10 },
        { title: '风险' },
    ],
    rows: [
        ['办公模式', '3 个格式引擎', '6 个格式引擎', '李工', '低'],
        ['协作模式', '100 人内测', '1000 人内测', '王工', '中'],
        ['搜索模式', '召回率 62%', '召回率 80%', '赵工', '高'],
        ['平台底座', '插件协议 v1', '插件市场', '陈工', '中'],
        ['合计', '3 项', '4 项', '—', '—'],
    ],
});
deck.notes('这页先讲结论，再展开每一条的实现路径。第二列的数字来自 2025Q4 的实测值。');
deck.quote({
    text: '好的工具应该让人忘记工具本身，只记得自己想做的事。',
    by: '产品委员会',
});
deck.statement({ text: '把复杂度留给我们，把确定性留给用户', sub: '2026 年唯一不变的目标' });
deck.image({
    path: 'assets/chart.png',
    title: '增长曲线',
    caption: '图 1：三个季度的活跃用户增长（数据来源：内部埋点）',
});
deck.notes('图表数据在附录 B，注意口径变化。');
deck.section({ title: '第二部分：资源与节奏', subtitle: '人力、预算与里程碑' });
deck.closing({ title: '谢谢', subtitle: '欢迎提问，也欢迎直接来找我们要 demo' });

const built = deck.inspect();
const bytes = deck.render();
const again = deck.render();
const parts = unzip(bytes);
const textOf = new Map([...parts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));

console.log(`# ${meta.name}（${meta.ext}）`);
console.log(`渲染体积 ${bytes.length} 字节，部件 ${parts.size} 个`);
console.log(`stats  ${JSON.stringify(built.stats)}`);
console.log('outline');
for (const line of built.outline) console.log(`  ${line}`);
console.log(`warnings(${built.warnings.length})  ${built.warnings.join(' / ') || '（无）'}`);

// ── 1. 幂等：同一份内容渲染两次必须得到同一批部件 ─────────────────────────

{
    const second = unzip(again);
    equal(second.size, parts.size, 'render() 两次的部件数应一致');
    for (const [name, data] of parts) {
        if (name === 'docProps/core.xml') continue; // 时间戳必然不同，其余必须逐字节一致
        const other = second.get(name);
        if (!other || Buffer.compare(Buffer.from(data), Buffer.from(other)) !== 0) {
            fail(`render() 两次内容不一致：${name}`);
        }
    }
}

// ── 2. 必需部件齐全 ──────────────────────────────────────────────────────

const required = [
    '[Content_Types].xml', '_rels/.rels', 'docProps/core.xml', 'docProps/app.xml',
    'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels',
    'ppt/theme/theme1.xml',
    'ppt/slideMasters/slideMaster1.xml', 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    'ppt/notesMasters/notesMaster1.xml', 'ppt/notesMasters/_rels/notesMaster1.xml.rels',
    'ppt/media/image1.png',
];
for (const name of required) assert(parts.has(name), `缺少必需部件：${name}`);
for (let i = 1; i <= 8; i += 1) {
    assert(parts.has(`ppt/slideLayouts/slideLayout${i}.xml`), `缺少版式：slideLayout${i}.xml`);
    assert(parts.has(`ppt/slideLayouts/_rels/slideLayout${i}.xml.rels`), `缺少版式关系：slideLayout${i}.xml.rels`);
}
for (let i = 1; i <= 10; i += 1) {
    assert(parts.has(`ppt/slides/slide${i}.xml`), `缺少页面：slide${i}.xml`);
    assert(parts.has(`ppt/slides/_rels/slide${i}.xml.rels`), `缺少页面关系：slide${i}.xml.rels`);
}
for (let i = 1; i <= 2; i += 1) {
    assert(parts.has(`ppt/notesSlides/notesSlide${i}.xml`), `缺少备注页：notesSlide${i}.xml`);
}
equal(parts.size, 52, '主样例的部件总数');

// 图片字节必须原样进包
{
    const stored = parts.get('ppt/media/image1.png');
    equal(stored.length, png.length, '媒体部件体积');
    assert(Buffer.compare(Buffer.from(stored), png) === 0, '媒体部件字节应与源文件一致');
    // 部件名从 image1 起顺延，扩展名按真实格式（不是按源文件后缀）—— 抽图时靠它找回原格式
    assert([...parts.keys()].filter((name) => name.startsWith('ppt/media/')).join(',') === 'ppt/media/image1.png',
        `媒体部件名单：${[...parts.keys()].filter((name) => name.startsWith('ppt/media/')).join(',')}`);
}

// ── 3. 逐个 XML / rels 解析，确认根元素与关键节点 ─────────────────────────

const roots = new Map();
for (const [name, text] of textOf) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    try {
        const doc = parseXml(text);
        const root = doc.children[0];
        if (root === undefined) {
            fail(`${name} 没有根元素`);
            continue;
        }
        roots.set(name, root);
        assert(doc.children.length === 1, `${name} 只能有一个根元素`);
        if (name.endsWith('.rels')) equal(root.name, 'Relationships', `${name} 的根元素`);
        else if (name === '[Content_Types].xml') equal(root.name, 'Types', `${name} 的根元素`);
    } catch (error) {
        fail(`${name} 解析失败：${error.message}`);
    }
}
equal(roots.size, [...parts.keys()].filter((name) => /\.(xml|rels)$/.test(name)).length, '可解析的 XML 部件数');

// presentation：子元素顺序与 id 起点
{
    const root = roots.get('ppt/presentation.xml');
    const names = xmlChildren(root).map((node) => node.name);
    equal(names.join(','), 'p:sldMasterIdLst,p:sldIdLst,p:notesMasterIdLst,p:sldSz,p:notesSz,p:defaultTextStyle',
        'p:presentation 的子元素顺序');
    const master = descendants(root, 'p:sldMasterId')[0];
    equal(master.attrs.id, '2147483648', 'p:sldMasterId 的 id');
    equal(master.attrs['r:id'], 'rId1', 'p:sldMasterId 指向 master');
    // 备注要真的显示在备注页上，必须登记 notesMaster 并有一条 presentation→notesMaster 关系；
    // 只放 notesSlide 部件的话 PowerPoint 不报错，但备注会被静默丢掉
    const notesMasterId = descendants(root, 'p:notesMasterId')[0];
    assert(notesMasterId !== undefined, '有备注时必须写 p:notesMasterIdLst');
    equal(notesMasterId.attrs.id, '2147483649', 'p:notesMasterId 的 id');
    assert(/^rId\d+$/.test(notesMasterId.attrs['r:id'] ?? ''), 'p:notesMasterId 必须指向一个关系');
    const notesRel = descendants(roots.get('ppt/_rels/presentation.xml.rels'), 'Relationship')
        .find((node) => node.attrs.Id === notesMasterId.attrs['r:id']);
    assert(notesRel !== undefined, 'presentation 里缺少 notesMaster 关系');
    equal(notesRel.attrs.Type, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
        'notesMaster 关系的类型');
    equal(notesRel.attrs.Target, 'notesMasters/notesMaster1.xml', 'notesMaster 关系的目标');
    const slideIds = descendants(root, 'p:sldId');
    equal(slideIds.length, 10, 'p:sldId 数量');
    equal(slideIds[0].attrs.id, '256', '第一个 p:sldId 的 id');
    equal(slideIds[9].attrs.id, '265', '最后一个 p:sldId 的 id');
    equal(new Set(slideIds.map((node) => node.attrs['r:id'])).size, 10, 'p:sldId 的 r:id 不能重复');
    const size = descendants(root, 'p:sldSz')[0];
    equal(size.attrs.cx, '12192000', '16:9 宽度');
    equal(size.attrs.cy, '6858000', '16:9 高度');
    const notesSz = descendants(root, 'p:notesSz')[0];
    equal(notesSz.attrs.cx, '6858000', 'notesSz 宽度');
}

// slideMaster：clrMap 12 槽 + txStyles 三块 + sldLayoutIdLst
{
    const root = roots.get('ppt/slideMasters/slideMaster1.xml');
    const clrMap = descendants(root, 'p:clrMap')[0];
    assert(clrMap !== undefined, 'master 缺少 p:clrMap');
    equal(Object.keys(clrMap.attrs).length, 12, 'clrMap 的属性数');
    for (const key of ['bg1', 'tx1', 'bg2', 'tx2', 'hlink', 'folHlink']) {
        assert(clrMap.attrs[key] !== undefined, `clrMap 缺少 ${key}`);
    }
    const txStyles = descendants(root, 'p:txStyles')[0];
    assert(txStyles !== undefined, 'master 缺少 p:txStyles');
    equal(xmlChildren(txStyles).map((node) => node.name).join(','), 'p:titleStyle,p:bodyStyle,p:otherStyle',
        'p:txStyles 的三个块');
    equal(descendants(txStyles, 'a:lvl1pPr').length >= 3, true, 'txStyles 至少要有三块里的 lvl1pPr');
    const layoutIds = descendants(root, 'p:sldLayoutId');
    equal(layoutIds.length, 8, 'p:sldLayoutId 数量');
    equal(layoutIds[0].attrs.id, '2147483649', '第一个 p:sldLayoutId 的 id');
    equal(layoutIds[7].attrs.id, '2147483656', '最后一个 p:sldLayoutId 的 id');
    equal(new Set(layoutIds.map((node) => node.attrs['r:id'])).size, 8, 'p:sldLayoutId 的 r:id 不能重复');
}

// theme1.xml：12 色 + major/minor 字体 + fmtScheme 四个列表各 3 项
{
    const root = roots.get('ppt/theme/theme1.xml');
    equal(root.name, 'a:theme', 'theme 根元素');
    const clrScheme = descendants(root, 'a:clrScheme')[0];
    const slots = xmlChildren(clrScheme).map((node) => node.name);
    equal(slots.length, 12, 'clrScheme 的颜色数');
    equal(slots.join(','), 'a:dk1,a:lt1,a:dk2,a:lt2,a:accent1,a:accent2,a:accent3,a:accent4,a:accent5,a:accent6,a:hlink,a:folHlink',
        'clrScheme 的槽位顺序');
    for (const name of ['a:majorFont', 'a:minorFont']) {
        const font = descendants(root, name)[0];
        assert(font !== undefined, `fontScheme 缺少 ${name}`);
        const kids = xmlChildren(font).map((node) => node.name);
        assert(kids.includes('a:latin') && kids.includes('a:ea') && kids.includes('a:cs'),
            `${name} 必须有 latin/ea/cs`);
        assert(kids.filter((kid) => kid === 'a:font').length >= 1, `${name} 应带中日韩脚本映射`);
    }
    const fmt = descendants(root, 'a:fmtScheme')[0];
    for (const list of ['a:fillStyleLst', 'a:lnStyleLst', 'a:effectStyleLst', 'a:bgFillStyleLst']) {
        const node = descendants(fmt, list)[0];
        assert(node !== undefined, `fmtScheme 缺少 ${list}`);
        equal(xmlChildren(node).length, 3, `${list} 的项数（PowerPoint 按 1-based 索引取用）`);
    }
    assert(descendants(root, 'a:objectDefaults').length === 1, 'theme 缺少 objectDefaults');
    assert(descendants(root, 'a:extraClrSchemeLst').length === 1, 'theme 缺少 extraClrSchemeLst');
}

// 每个 layout / slide / notesSlide：spTree 前两个子元素 + clrMapOvr
{
    const targets = [...parts.keys()].filter((name) => /^ppt\/(slideLayouts\/slideLayout\d+|slides\/slide\d+|notesSlides\/notesSlide\d+)\.xml$/.test(name));
    equal(targets.length, 20, 'layout + slide + notesSlide 的总数');
    for (const name of targets) {
        const root = roots.get(name);
        const spTree = descendants(root, 'p:spTree')[0];
        if (!assert(spTree !== undefined, `${name} 缺少 p:spTree`)) continue;
        const kids = xmlChildren(spTree);
        equal(kids[0]?.name, 'p:nvGrpSpPr', `${name} 的 spTree 第一个子元素`);
        equal(kids[1]?.name, 'p:grpSpPr', `${name} 的 spTree 第二个子元素`);
        const xfrm = descendants(kids[1], 'a:xfrm')[0];
        const zeros = ['a:off', 'a:ext', 'a:chOff', 'a:chExt'].every((tag) => {
            const node = descendants(xfrm, tag)[0];
            return node !== undefined
                && ['x', 'y', 'cx', 'cy'].every((key) => node.attrs[key] === undefined || node.attrs[key] === '0');
        });
        assert(zeros, `${name} 的 grpSpPr xfrm 必须四个值全 0`);
        const clrMapOvr = descendants(root, 'p:clrMapOvr')[0];
        assert(clrMapOvr !== undefined, `${name} 缺少 p:clrMapOvr`);
        assert(descendants(clrMapOvr, 'a:masterClrMapping').length === 1, `${name} 的 clrMapOvr 必须指向 masterClrMapping`);
        // 形状 id 必须唯一且 > 0
        const ids = descendants(spTree, 'p:cNvPr').map((node) => Number(node.attrs.id));
        assert(ids.every((id) => Number.isInteger(id) && id > 0), `${name} 存在非法形状 id`);
        equal(new Set(ids).size, ids.length, `${name} 的形状 id 必须唯一`);
    }
}

// 版式语义：cSld name 必须是中文版式名
{
    const names = [];
    for (let i = 1; i <= 8; i += 1) {
        const root = roots.get(`ppt/slideLayouts/slideLayout${i}.xml`);
        const cSld = descendants(root, 'p:cSld')[0];
        names.push(cSld.attrs.name);
        assert(root.attrs.preserve === '1', `slideLayout${i} 应有 preserve="1"`);
        assert(typeof root.attrs.type === 'string' && root.attrs.type !== '', `slideLayout${i} 应有 type`);
    }
    equal(names.join(','), '封面,章节,要点,表格,引用,陈述,图片,结尾', '8 个版式名');
}

// 内容页的标题带 / 页脚 / 页码
{
    const slide3 = textOf.get('ppt/slides/slide3.xml');
    assert(slide3.includes('name="Title"'), '内容页应有 Title 文本框');
    assert(slide3.includes('name="Page Number"'), '内容页应有页码');
    assert(slide3.includes('>3 / 10<'), '页码文本应为「3 / 10」');
    assert(slide3.includes('<a:buChar char="•"/>'), '一级要点应使用 buChar');
    assert(slide3.includes('<a:buChar char="–"/>'), '二级要点应使用不同的符号');
    assert(slide3.includes('<a:ea typeface="微软雅黑"/>'), '中文必须写 a:ea 字体');
    assert(slide3.includes('lang="zh-CN"'), '文字必须声明 lang="zh-CN"');
    const cover = textOf.get('ppt/slides/slide1.xml');
    assert(cover.includes('name="Kicker"'), '封面应有 kicker');
    assert(cover.includes('name="Presenter"'), '封面应有 presenter');
    assert(!cover.includes('name="Page Number"'), '封面不应有页码');
}

// 表格：表头填充 + 网格列 + 行高
{
    const slide5 = textOf.get('ppt/slides/slide5.xml');
    assert(slide5.includes('<a:tbl>'), '第 5 页应有表格');
    equal((slide5.match(/<a:gridCol /g) ?? []).length, 5, '表格列数');
    equal((slide5.match(/<a:tr /g) ?? []).length, 6, '表格行数（含表头）');
    assert(slide5.includes('val="0E7490"'), '表头应使用主题色填充');
    assert(slide5.includes('<a:lnL w="0"><a:noFill/></a:lnL>'), '竖线应显式关闭');
    assert(!slide5.includes('tableStyleId'), '不应引用不存在的 tableStyleId');
}

// 备注页
{
    const notes = textOf.get('ppt/notesSlides/notesSlide1.xml');
    assert(notes.includes('type="body" idx="1"'), '备注页应有 body 占位符');
    assert(notes.includes('这页先讲结论'), '备注文本应写入 notesSlide');
}

// ── 4. 关系目标存在性 + 内容类型覆盖 + r:id 引用有效性 ────────────────────

for (const [name, root] of roots) {
    if (!name.endsWith('.rels')) continue;
    const owner = ownerOf(name);
    assert(owner !== undefined, `${name} 不是合法的关系部件路径`);
    const ids = [];
    for (const rel of descendants(root, 'Relationship')) {
        ids.push(rel.attrs.Id);
        if (rel.attrs.TargetMode === 'External') continue;
        const target = resolveTarget(owner, rel.attrs.Target);
        assert(parts.has(target), `${name} 的目标不存在：${rel.attrs.Target} → ${target}`);
    }
    equal(new Set(ids).size, ids.length, `${name} 的关系 id 必须唯一`);
}

{
    const ctRoot = roots.get('[Content_Types].xml');
    const overrides = new Set();
    const defaults = new Set();
    for (const node of xmlChildren(ctRoot)) {
        if (node.name === 'Override') overrides.add(node.attrs.PartName.replace(/^\//, ''));
        if (node.name === 'Default') defaults.add(String(node.attrs.Extension).toLowerCase());
    }
    for (const name of parts.keys()) {
        if (overrides.has(name)) continue;
        const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
        assert(defaults.has(ext), `内容类型未声明：${name}`);
    }
    for (const name of overrides) assert(parts.has(name), `[Content_Types].xml 声明了不存在的部件：${name}`);
    assert(defaults.has('rels') && defaults.has('xml'), 'Default 必须覆盖 rels 与 xml');
    assert(overrides.has('ppt/notesSlides/notesSlide1.xml'), 'notesSlide 必须有 Override');
    assert(overrides.has('ppt/media/image1.png') === false, '图片应走 Default 而不是 Override');
}

{
    // 每个部件里出现的 r:id / r:embed 都必须在该部件的 rels 里
    for (const [name, text] of textOf) {
        if (name.endsWith('.rels')) continue;
        const relsName = relsPathOf(name);
        const relsRoot = roots.get(relsName);
        const available = new Set(relsRoot === undefined ? [] : descendants(relsRoot, 'Relationship').map((node) => node.attrs.Id));
        for (const match of text.matchAll(/r:(?:id|embed|link)="([^"]+)"/g)) {
            assert(available.has(match[1]), `${name} 引用了不存在的 ${match[1]}（${relsName}）`);
        }
    }
}

// ── 5. save() 报告 + read() 复检 ──────────────────────────────────────────

const report = deck.save();
equal(report.ok, true, 'save() 的 ok');
equal(report.format, 'pptx', 'save() 的 format');
equal(report.path, 'demo/2026-产品规划.pptx', 'save() 的相对路径');
equal(report.theme, 'tech', 'save() 的主题');
equal(report.bytes, bytes.length, 'save() 的体积应与 render() 一致');
equal(report.stats.slides, 10, 'stats.slides');
equal(report.stats.notes, 2, 'stats.notes');
equal(report.stats.tables, 1, 'stats.tables');
equal(report.stats.images, 1, 'stats.images');
equal(report.stats.layouts.cover, 1, 'stats.layouts.cover');
equal(report.stats.layouts.section, 2, 'stats.layouts.section');
equal(report.stats.layouts.bullets, 2, 'stats.layouts.bullets');
equal(report.stats.layouts.table, 1, 'stats.layouts.table');
equal(report.stats.layouts.quote, 1, 'stats.layouts.quote');
equal(report.stats.layouts.statement, 1, 'stats.layouts.statement');
equal(report.stats.layouts.image, 1, 'stats.layouts.image');
equal(report.stats.layouts.closing, 1, 'stats.layouts.closing');
assert(report.stats.shapes > 40, `stats.shapes 应大于 40，实际 ${report.stats.shapes}`);
assert(report.stats.textBoxes > 30, `stats.textBoxes 应大于 30，实际 ${report.stats.textBoxes}`);
assert(report.stats.words > 150, `stats.words 应大于 150，实际 ${report.stats.words}`);
assert(report.stats.mediaBytes === png.length, 'stats.mediaBytes 应等于图片体积');
equal(report.outline.length, 10, 'outline 行数');
assert(report.outline[0].startsWith('P1 cover：'), `outline[0] 格式：${report.outline[0]}`);
assert(report.outline[4].startsWith('P5 table：'), `outline[4] 格式：${report.outline[4]}`);
assert(report.outline[4].includes('5 行 × 5 列'), `outline[4] 应含表格尺寸：${report.outline[4]}`);
assert(report.outline[7].startsWith('P8 image：'), `outline[7] 格式：${report.outline[7]}`);
assert(report.outline[2].includes('5 条'), `outline[2] 应含要点条数：${report.outline[2]}`);
assert(report.outline[3].includes('双栏'), `outline[3] 应标出双栏：${report.outline[3]}`);
equal(report.warnings.length, 0, `主样例不应有排版警告：${report.warnings.join(' / ')}`);
assert(env.writes.some((write) => write.path === 'demo/2026-产品规划.pptx'), '落盘应走 env.writeFile');

const back = read(report.path, env);
equal(back.ok, true, 'read() 的 ok');
equal(back.format, 'pptx', 'read() 的 format');
equal(back.bytes, bytes.length, 'read() 的体积');
equal(back.theme, 'tech', 'read() 应从 core.xml 取回主题');
equal(back.size, '16:9', 'read() 的尺寸');
equal(back.stats.slides, 10, 'read().stats.slides');
equal(back.stats.notes, 2, 'read().stats.notes');
equal(back.stats.tables, 1, 'read().stats.tables');
equal(back.stats.images, 1, 'read().stats.images');
equal(back.stats.layouts.table, 1, 'read().stats.layouts.table');
equal(back.pages[0].layout, 'cover', '第 1 页版式');
equal(back.pages[4].layout, 'table', '第 5 页版式');
equal(back.pages[4].tableRows, 6, '第 5 页表格行数');
equal(back.pages[4].tableColumns, 5, '第 5 页表格列数');
equal(back.pages[7].layout, 'image', '第 8 页版式');
equal(back.pages[7].images, 1, '第 8 页图片数');
equal(back.pages[2].layout, 'bullets', '第 3 页版式');
equal(back.pages[2].title, '三条产品线的成绩与短板', '第 3 页标题');
equal(back.pages[2].lines, 5, '第 3 页正文段落数');
assert(back.pages[4].notes.startsWith('这页先讲结论'), `第 5 页备注摘要：${back.pages[4].notes}`);
equal(back.pages[9].layout, 'closing', '第 10 页版式');
equal(back.outline.length, 10, 'read().outline 行数');
equal(back.stats.textBoxes, report.stats.textBoxes, 'read() 与 save() 的文本框数应一致');
equal(back.stats.words, report.stats.words, 'read() 与 save() 的字数应一致');
equal(back.warnings.length, 0, `主样例 read() 不应有警告：${back.warnings.join(' / ')}`);

// stats.media：逐张报清楚「包里是哪几张图、什么类型、多大、在第几页」。
// 只给一个 imageParts 计数看不出改哪一张，这一组断言盯着明细与计数的一致性。
{
    equal(back.stats.imageParts, 1, 'read().stats.imageParts');
    equal(back.stats.media.length, back.stats.imageParts, 'stats.media 与 imageParts 必须是同一批条目');
    const item = back.stats.media[0];
    equal(item.part, 'ppt/media/image1.png', 'stats.media[0].part');
    equal(item.contentType, 'image/png', 'stats.media[0].contentType');
    equal(item.bytes, png.length, 'stats.media[0].bytes');
    equal(item.width, 1200, 'stats.media[0].width');
    equal(item.height, 675, 'stats.media[0].height');
    assert(item.slides.includes(8), `stats.media[0].slides 应含第 8 页：${JSON.stringify(item.slides)}`);
    assert(back.pages[7].images === 1 && back.pages[7].layout === 'image', '第 8 页是图片页');
}

// 抽取：把包里的图落成文件，字节必须与源文件一致（这一步是「读图能力」的入口）
{
    const extracted = images(report.path, { out: 'extract' }, env);
    equal(extracted.ok, true, 'images() 的 ok');
    equal(extracted.format, 'pptx', 'images() 的 format');
    equal(extracted.count, 1, 'images() 抽出的图片数');
    equal(extracted.reused, 0, '同一张图只被引用一次时不应报 reused');
    equal(extracted.skipped.length, 0, `主样例不应有跳过的图：${JSON.stringify(extracted.skipped)}`);
    const file = extracted.files[0];
    equal(file.part, 'ppt/media/image1.png', 'images().files[0].part');
    equal(file.ext, 'png', 'images().files[0].ext');
    equal(file.contentType, 'image/png', 'images().files[0].contentType');
    equal(file.bytes, png.length, 'images().files[0].bytes');
    equal(file.width, 1200, 'images().files[0].width');
    equal(file.height, 675, 'images().files[0].height');
    assert(file.slides.includes(8), `images().files[0].slides：${JSON.stringify(file.slides)}`);
    assert(file.relIds.length > 0 && /^rId\d+$/.test(file.relIds[0]), `files[0].relIds：${JSON.stringify(file.relIds)}`);
    assert(file.path.startsWith('extract/'), `files[0].path 应落在 out 目录里：${file.path}`);
    assert(Buffer.compare(env.readFile(file.path), png) === 0, '抽出来的字节必须与源文件逐字节一致');
    equal(extracted.dir, 'extract', 'images() 的 dir');
    assert(typeof extracted.hint === 'string' && extracted.hint.includes('read_image'), 'images() 要给下一步提示');
}

console.log(`read() 复检 stats  ${JSON.stringify(back.stats)}`);
console.log(`read() 第 1/5/8 页  ${back.pages[0].layout} / ${back.pages[4].layout} / ${back.pages[7].layout}`
    + `，第 3 页「${back.pages[2].title}」${back.pages[2].lines} 段`);

// ── 6. 边界样例：4:3 + 未知主题 + 每一条 warning 都要真的触发 ──────────────

const edge = create({ title: '边界样例', theme: '不存在的主题', path: 'edge.pptx', size: '4:3', author: '测试' }, env);
edge.bullets({
    title: '要点过多的页面',
    items: [
        { text: '第一条', level: 1 }, { text: '第二条', level: 1 }, { text: '第三条', level: 1 },
        { text: '第四条', level: 1 }, { text: '第五条', level: 1 }, { text: '第六条', level: 1 },
        { text: '第七条', level: 1 }, { text: '第八条', level: 1 }, { text: '第九条', level: 1 },
        { text: '这是一条三级要点，层级已经过深', level: 3 },
        { text: '这是一条特别长的要点，需要折很多行才放得下'.repeat(6), level: 1 },
        { text: '第二条超长要点，同样会把版心撑爆'.repeat(6), level: 1 },
        { text: '第三条超长要点，用来验证放不下的分支'.repeat(6), level: 1 },
    ],
});
edge.table({
    title: '列很多的表格',
    columns: ['项目名称', '负责人', '开始时间', '结束时间', '预算', '进度', '风险等级'],
    rows: Array.from({ length: 22 }, (_, index) => [
        `任务 ${index + 1}`, `负责人 ${index + 1}`, '2026-01-01', '2026-03-31', '120,000', `${index * 4}%`, '中',
    ]),
});
edge.image({ path: 'assets/not-here.png', title: '缺失的图片' });
edge.statement({ text: '这一句宣言会很长，长到任何字号都放不下'.repeat(12), sub: '压到最小字号' });
const edgeReport = edge.save();

const edgeParts = unzip(edge.render());
const edgeSize = (() => {
    const text = new TextDecoder('utf-8').decode(edgeParts.get('ppt/presentation.xml'));
    return /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(text);
})();
equal(edgeSize?.[1], '9144000', '4:3 宽度');
// 没有备注时不应出现 notesMasterIdLst，否则会多挂一个用不到的主版
assert(!new TextDecoder('utf-8').decode(edgeParts.get('ppt/presentation.xml')).includes('notesMasterIdLst'),
    '无备注时不应写 p:notesMasterIdLst');
assert(!edgeParts.has('ppt/notesMasters/notesMaster1.xml'), '无备注时不应生成 notesMaster 部件');
equal(edgeReport.theme, DEFAULT_THEME_ID, '未知主题应回落到默认主题');
assert(edgeReport.warnings.some((line) => line.includes('不存在，已回落')), '应报告主题回落');
assert(edgeReport.warnings.some((line) => line.includes('超过 7 条')), '应报告要点条数超限');
assert(edgeReport.warnings.some((line) => line.includes('3 级')), '应报告层级过深');
assert(edgeReport.warnings.some((line) => line.includes('列，列宽')), '应报告列数过多');
assert(edgeReport.warnings.some((line) => line.includes('超过可用')), '应报告表格超版心');
assert(edgeReport.warnings.some((line) => line.includes('图片缺失')), '应报告图片缺失');
assert(edgeReport.warnings.some((line) => line.includes('放不下')), '应报告文字放不下');
assert(edgeReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
    `警告里不应出现 undefined/NaN：${edgeReport.warnings.join(' / ')}`);
equal(edgeReport.stats.slides, 4, '边界样例页数');
console.log(`边界样例 warnings(${edgeReport.warnings.length})`);
for (const line of edgeReport.warnings) console.log(`  - ${line}`);
const edgeBack = read(edgeReport.path, env);
equal(edgeBack.size, '4:3', 'read() 应识别 4:3');
equal(edgeBack.stats.slides, 4, '边界样例 read().stats.slides');

// ── 7. 图片能力：contain / cover / natural、显式尺寸、对齐、边框、满页、多图网格 ──
//
// 这一节单独用一个 deck：主样例已经断言了「10 页 / 52 个部件 / 0 条警告」，
// 往里面塞图会连带改掉那些数字，反而看不清是图片能力坏了还是主样例变了。

const img = create({ title: '图片能力样例', theme: 'plain', path: 'images.pptx', size: '16:9', author: '测试' }, env);
// 1400×787 与 1400×1050 两种比例：前者略宽于 16:9，后者是 4:3，按 cover 一定会被裁。
// 宽度都取 1400 是为了「铺满整页时仍有 96 DPI 以上」——低于这个宽度会把满页背景也拖进 DPI 告警，
// 那样就分不清「测试图太小」和「引擎算错了」
env.writeFile('assets/wide.png', makePng(1400, 787));
env.writeFile('assets/tall.png', makePng(1400, 1050));
// 200×100 的小图放到 24cm 宽：DPI 必然低于 96，用来验证「发虚」告警
env.writeFile('assets/tiny.png', makePng(200, 100));

img.image({ path: 'assets/wide.png', title: '默认 contain' });                       // P1
img.image({ path: 'assets/wide.png', title: 'cover 铺满', fit: 'cover', caption: '铺满后左右会各切一刀' }); // P2
img.image({ path: 'assets/wide.png', title: 'natural 原始尺寸', fit: 'natural' });     // P3
img.image({ path: 'assets/wide.png', title: '只给宽度', widthCm: 6 });                 // P4
img.image({ path: 'assets/wide.png', title: '只给高度', heightCm: 4 });                // P5
img.image({ path: 'assets/wide.png', title: '左对齐', align: 'left' });                // P6
img.image({ path: 'assets/wide.png', title: '右对齐 + 边框', align: 'right', frame: true }); // P7
img.image({ path: 'assets/wide.png', title: '满页背景', fullBleed: true });            // P8
img.image({ path: 'assets/tall.png', title: '满页背景（比例不符）', fullBleed: true }); // P9
img.image({ path: 'assets/tiny.png', title: '低分辨率', widthCm: 24 });                // P10
img.image({ path: 'assets/wide.png', title: '非法 fit', fit: 'stretch', align: 'middle' }); // P11
img.image({ path: 'assets/missing.png', title: '缺图' });                              // P12
img.images([
    { path: 'assets/wide.png', caption: '图一' },
    { path: 'assets/tall.png', caption: '图二' },
    { path: 'assets/missing.png', caption: '图三（缺图）' },
    { path: 'assets/wide.png' },
], { title: '2 列网格' });                                                             // P13
img.images([
    { path: 'assets/wide.png' },
    { path: 'assets/tall.png' },
    { path: 'assets/wide.png' },
], { title: '3 列网格', columns: 3, gapCm: 0.5, fit: 'cover' });                       // P14

const imgReport = img.save();
const imgParts = unzip(img.render());
const imgText = new Map([...imgParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
const slideXml = (page) => imgText.get(`ppt/slides/slide${page}.xml`) ?? '';
const extOf = (shape) => {
    const match = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(shape ?? '');
    return match === null ? undefined : { w: Number(match[1]), h: Number(match[2]) };
};
// 只看 p:pic（真图片）与 p:sp 里带 a:xfrm 的（占位框）：spTree 的 grpSpPr 也是 a:ext 全 0，
// 不按元素切出来就会把那个 0 当成图片尺寸
const pictureBoxes = (page) => [...slideXml(page).matchAll(/<p:pic>([\s\S]*?)<\/p:pic>/g)]
    .map((match) => {
        const ext = extOf(match[1]);
        const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(match[1]);
        return {
            x: Number(off[1]), y: Number(off[2]), w: ext.w, h: ext.h, xml: match[1],
        };
    });
const placeholderBoxes = (page) => [...slideXml(page).matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)]
    .filter((match) => match[1].includes('name="Missing Image"'))
    .map((match) => {
        const ext = extOf(match[1]);
        const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(match[1]);
        return {
            x: Number(off[1]), y: Number(off[2]), w: ext.w, h: ext.h, xml: match[1],
        };
    });
// 只取「真图片 + 缺图占位框」：页面上的标题带/页脚/题注都是 p:sp，按 y 排序会被它们顶到最前面。
// 有真图片时优先返回真图片（满页背景页的标题浮层 y 更小，但我们要断言的是图片本身）
const firstBox = (page) => {
    const pics = pictureBoxes(page);
    const source = pics.length > 0 ? pics : placeholderBoxes(page);
    return source.sort((a, b) => a.y - b.y || a.x - b.x)[0];
};
// 版心几何（与 ppt.js 的 makeGeometry 同一套数字，测试里写死是为了「算错了能发现」）
const MARGIN = cmToEmu(0.9);
const CONTENT_W = 12192000 - MARGIN * 2;
const BODY_Y = cmToEmu(3.25);
const FOOTER_Y = 6858000 - cmToEmu(1.10);
const BODY_H = Math.max(cmToEmu(3), FOOTER_Y - cmToEmu(0.25) - BODY_Y);
// 可用高度：contain 会为「有题注」的页面扣掉题注条，cover 的样例带了题注
const AREA_PLAIN = Math.max(cmToEmu(3), BODY_H - cmToEmu(0.25));
const AREA_CAPTION = Math.max(cmToEmu(3), BODY_H - cmToEmu(1.0) - cmToEmu(0.25));
// 期望的等比尺寸：整条链路用同一个缩放系数，w/h 才严格保持纵横比
const fitted = (px, py, areaW, areaH) => {
    const s = Math.min(areaW / (px * 9525), areaH / (py * 9525));
    return { w: Math.max(9525, Math.round(px * 9525 * s)), h: Math.max(9525, Math.round(py * 9525 * s)) };
};

equal(imgReport.stats.slides, 14, '图片样例页数');
equal(imgReport.stats.layouts.image, 14, '图片样例全部走 image 版式（多图网格也归到 image）');
{
    // 逐页核对「真实放置的图片数」：满页背景不计入 stats.images（它是背景不是配图），
    // 但它确实画在页面上，所以按页数出来是 19、按 stats 口径是 17
    const perPage = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]
        .map((page) => [...pictureBoxes(page), ...placeholderBoxes(page)].length);
    equal(perPage.join(','), '1,1,1,1,1,1,1,1,1,1,1,1,4,3',
        '逐页图片形状数（P13 四格含缺图占位框，P14 三格）');
    equal(perPage[7], 1, 'P8 满页背景页面上仍应有 1 张图');
    equal(perPage[8], 1, 'P9 满页背景页面上仍应有 1 张图');
}
equal(imgReport.stats.images, 17, 'stats.images 应为真实放置的图片数（12 单图 + 4 + 3 网格，两张满页背景不计）');
equal(imgReport.stats.imageParts, 3, 'stats.imageParts 应为去重后的媒体部件数');
equal(imgReport.stats.cropped, 5, 'stats.cropped：cover 3 张 + 两张满页背景（16:9 与 4:3 各一张）');
equal(imgReport.stats.lowDpi, 1, 'stats.lowDpi：只有那张 200×100 的小图发虚');
assert(imgParts.has('ppt/media/image1.png') && imgParts.has('ppt/media/image2.png') && imgParts.has('ppt/media/image3.png'),
    '三张不同的图应生成三个媒体部件');
equal(imgParts.has('ppt/media/image4.png'), false, '重复路径不应重复打包媒体');

// contain：等比放进「版心宽 × (版心高 - 题注)」，纵横比必须原样保持
{
    const box = firstBox(1);
    assert(box !== undefined, 'P1 应有一个图片形状');
    const want = fitted(1400, 787, CONTENT_W, AREA_PLAIN);
    equal(box.w, want.w, 'P1 contain 宽度应按可用区域等比缩放');
    equal(box.h, want.h, 'P1 contain 高度应严格等比');
    equal(box.x, Math.round(MARGIN + (CONTENT_W - box.w) / 2), 'P1 默认水平居中');
    equal(box.y, Math.round(BODY_Y + (AREA_PLAIN - box.h) / 2), 'P1 默认纵向居中');
    assert(!slideXml(1).includes('<a:srcRect'), 'contain 不应写 a:srcRect');
}

// cover：铺满可用区域（宽 = 版心宽），并写 a:srcRect 裁掉超出部分
{
    const xml = slideXml(2);
    const box = firstBox(2);
    assert(box !== undefined, 'P2 应有一个图片形状');
    equal(box.w, CONTENT_W, 'P2 cover 宽度应铺满版心');
    equal(box.h, AREA_CAPTION, 'P2 cover 高度应铺满可用高度');
    const srcRect = /<a:srcRect l="(\d+)" t="(\d+)" r="(\d+)" b="(\d+)"\/>/.exec(xml);
    assert(srcRect !== null, 'P2 cover 应写 a:srcRect');
    // 图片比可用区域更宽，所以裁的是上下；裁切量必须对称且不为 0
    assert(Number(srcRect[2]) > 0 && Number(srcRect[4]) > 0, 'P2 偏宽的图按 cover 应裁上下');
    equal(Number(srcRect[2]), Number(srcRect[4]), 'P2 上下裁切量应对称');
    assert(Number(srcRect[1]) < 100000 && Number(srcRect[2]) < 100000,
        'a:srcRect 不能写满 100000（那等于把图裁没了）');
    assert(xml.indexOf('<a:srcRect') < xml.indexOf('<a:stretch>'),
        'a:srcRect 必须排在 a:stretch 之前（schema 顺序）');
}

// natural：基准是 96 DPI 折算的原始像素尺寸，放不下时与 contain 缩到一样大（只缩不放）
{
    const box = firstBox(3);
    assert(box !== undefined, 'P3 应有一个图片形状');
    equal(box.w, firstBox(1).w, 'P3 natural 超版心时应与 contain 缩到同一尺寸');
    equal(box.h, firstBox(1).h, 'P3 natural 高度同样与 contain 一致');
    // 一张比版心小的图：natural 不会把它放大，contain 也不放大，两者都等于原始像素尺寸
    assert(1400 * 9525 > CONTENT_W, '样例图应比版心宽，这样 natural 才走「缩」的分支');
}

// 显式尺寸：只给宽度按比例推高度，只给高度按比例推宽度
{
    const onlyW = firstBox(4);
    equal(onlyW.w, cmToEmu(6), 'P4 显式宽度');
    equal(onlyW.h, Math.round(cmToEmu(6) * (787 / 1400)), 'P4 只给宽度时应按纵横比推高度');
    const onlyH = firstBox(5);
    equal(onlyH.h, cmToEmu(4), 'P5 显式高度');
    equal(onlyH.w, Math.round(cmToEmu(4) * (1400 / 787)), 'P5 只给高度时应按纵横比推宽度');
}

// 对齐：宽度小于版心时 left / right 贴边
{
    const left = firstBox(6);
    equal(left.x, MARGIN, 'P6 align=left 应贴左边距');
    const right = firstBox(7);
    equal(right.x + right.w, 12192000 - MARGIN, 'P7 align=right 应贴右边距');
    assert(slideXml(7).includes('<a:ln w="9525">'), 'P7 frame=true 应画细边框');
    assert(!slideXml(1).includes('<a:ln w="9525">'), 'P1 未要求 frame 时不应加边框');
}

// fullBleed：铺满整页（忽略页边距与标题区），且不要标题带与页脚
{
    const box = firstBox(8);
    equal(box.x, 0, 'P8 满页背景 x 应为 0');
    equal(box.y, 0, 'P8 满页背景 y 应为 0');
    equal(box.w, 12192000, 'P8 满页背景宽度应为整页宽');
    equal(box.h, 6858000, 'P8 满页背景高度应为整页高');
    assert(slideXml(8).includes('Picture Full Bleed'), '满页背景应标记为 Picture Full Bleed');
    assert(!slideXml(8).includes('name="Page Number"'), '满页背景不应再叠页码');
    assert(!slideXml(8).includes('name="Title Band"'), '满页背景不应再叠标题色带');
    assert(!slideXml(8).includes('name="Footer Text"'), '满页背景不应再叠页脚');
}

// 满页背景比例不符：裁切 + 告警
{
    const box = firstBox(9);
    equal(box.w, 12192000, 'P9 比例不符的满页背景仍应铺满整页宽');
    equal(box.h, 6858000, 'P9 比例不符的满页背景仍应铺满整页高');
    assert(slideXml(9).includes('<a:srcRect'), 'P9 比例不符时应用 a:srcRect 裁切');
    assert(imgReport.warnings.some((line) => line.includes('满页背景') && line.includes('裁')),
        '应报告满页背景被裁');
}

// 低 DPI 与裁切告警
assert(imgReport.warnings.some((line) => line.includes('低于 96 DPI')), '应报告 DPI 过低');
assert(imgReport.warnings.some((line) => line.includes('cover 铺满') && line.includes('已被裁切')),
    '应报告 cover 裁切');

// 非法输入只变 warning，不抛异常
assert(imgReport.warnings.some((line) => line.includes('未知的图片适配方式')), '应报告非法 fit');
assert(imgReport.warnings.some((line) => line.includes('未知的图片对齐方式')), '应报告非法 align');
equal(firstBox(11).w, firstBox(1).w, '非法 fit 应回落到 contain');
equal(firstBox(11).x, Math.round(MARGIN + (CONTENT_W - firstBox(11).w) / 2), '非法 align 应回落到 center');
assert(imgReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
    `图片告警里不应出现 undefined/NaN：${imgReport.warnings.join(' / ')}`);

// 缺图占位框
assert(slideXml(12).includes('name="Missing Image"'), 'P12 缺图应画占位框');
assert(slideXml(12).includes('图片缺失：assets/missing.png'), 'P12 占位框应写明缺失路径');

// 多图网格：格宽 = (版心宽 - gap × (列数 - 1)) / 列数，逐格等比放入并居中
{
    const gap = cmToEmu(0.4);
    const cellW2 = Math.round((CONTENT_W - gap) / 2);
    const rowH = Math.round((BODY_H - cmToEmu(0.25)) / 2);
    const cellH = Math.max(cmToEmu(1.2), rowH - cmToEmu(1.05));
    const boxes = [...pictureBoxes(13), ...placeholderBoxes(13)]
        .sort((a, b) => a.y - b.y || a.x - b.x);
    equal(boxes.length, 4, 'P13 应有 4 个图片形状（含缺图占位框）');
    // 第一行两张图：同一起点 y，且都落在自己的格子里
    const row1 = boxes.filter((box) => box.y === BODY_Y);
    equal(row1.length, 2, 'P13 第一行应有 2 格');
    assert(row1.every((box) => box.w <= cellW2 && box.h <= cellH),
        'P13 第一行每格都不应超出版心格宽/格高');
    // 每张图在自己的格子里居中（align 默认 center），两张图分别属于第 1 / 第 2 列
    const centers = row1.map((box) => box.x + box.w / 2).sort((a, b) => a - b);
    equal(Math.round(centers[0]), Math.round(MARGIN + cellW2 / 2), 'P13 第一行第 1 格图片应在格内水平居中');
    equal(Math.round(centers[1]), Math.round(MARGIN + cellW2 + gap + cellW2 / 2), 'P13 第一行第 2 格图片应在格内水平居中');
    // 第二行起点 = 正文起点 + 行高
    const row2 = boxes.filter((box) => box.y > BODY_Y);
    equal(row2.length, 2, 'P13 第二行应有 2 格');
    assert(row2.every((box) => box.y === BODY_Y + rowH), 'P13 第二行起点 = 正文起点 + 行高');
    // 缺图占位框铺满整格（比图更大），且写明路径
    const placeholder = row2.find((box) => box.xml.includes('Missing Image'));
    assert(placeholder !== undefined, 'P13 缺图格子应画占位框且不影响其它格');
    equal(placeholder.w, cellW2, 'P13 缺图占位框应铺满格宽');
    equal(placeholder.h, cellH, 'P13 缺图占位框应铺满格高');
    assert(placeholder.x === MARGIN || placeholder.x === Math.round(MARGIN + cellW2 + gap),
        'P13 缺图占位框应落在自己的格子里');
    assert(slideXml(13).includes('Cell Caption'), 'P13 每格题注应各自成框');
    assert(slideXml(13).includes('图片缺失：assets/missing.png'), 'P13 缺图占位框应写明缺失路径');
    assert(imgReport.outline[12].includes('2 列 × 4 图'), `outline 应标出网格：${imgReport.outline[12]}`);
}
{
    const gap = cmToEmu(0.5);
    const cellW3 = Math.round((CONTENT_W - gap * 2) / 3);
    const boxes = pictureBoxes(14);
    equal(boxes.length, 3, 'P14 应有 3 个图片形状');
    assert(boxes.every((box) => box.w === cellW3), `P14 三列格宽应一致（期望 ${cellW3}）`);
    assert(boxes.every((box) => box.xml.includes('<a:srcRect')), 'P14 fit=cover 的网格每格都应带 a:srcRect');
    assert(imgReport.outline[13].includes('3 列 × 3 图'), `outline 应标出三列网格：${imgReport.outline[13]}`);
}
// 满页背景没写 fit 时按 cover 报（背景图默认铺满）
assert(imgReport.outline[7].includes('cover｜') && imgReport.outline[7].includes('cm'),
    `outline 应含图片尺寸：${imgReport.outline[7]}`);
assert(imgReport.outline[1].includes('（有裁切）'), `outline 应标出裁切：${imgReport.outline[1]}`);

// read() 回读：尺寸 / 裁切 / DPI 都要来自真实解析
{
    const imgBack = read(imgReport.path, env);
    equal(imgBack.stats.slides, 14, '图片样例 read().stats.slides');
    equal(imgBack.stats.images, 17, '图片样例 read().stats.images');
    equal(imgBack.stats.imageParts, 3, '图片样例 read().stats.imageParts');
    equal(imgBack.stats.cropped, 5, '图片样例 read().stats.cropped');
    equal(imgBack.stats.lowDpi, 1, '图片样例 read().stats.lowDpi');
    equal(imgBack.pages[7].fullBleed, true, 'P8 应被读成满页背景');
    equal(imgBack.pages[7].images, 1, 'P8 图片数');
    equal(imgBack.pages[1].cropped, 1, 'P2 应被读成有裁切');
    equal(imgBack.pages[0].cropped, 0, 'P1 无裁切');
    const p1 = firstBox(1);
    equal(imgBack.pages[0].imageSizes[0],
        `${Math.round((p1.w / 360000) * 10) / 10}×${Math.round((p1.h / 360000) * 10) / 10}cm`,
        'P1 展示尺寸应按解析出的 EMU 换算');
    assert(imgBack.pages[9].dpi > 0 && imgBack.pages[9].dpi < 96, 'P10 应被判为低 DPI');
    assert(imgBack.warnings.some((line) => line.includes('低于 96 DPI')), 'read() 应复现 DPI 告警');
    assert(imgBack.warnings.some((line) => line.includes('a:srcRect 裁切')), 'read() 应复现裁切告警');
    // read() 数的是包里的 p:pic；缺图占位框是 p:sp（画出来是虚线框），所以这里是 3 而不是 4
    assert(imgBack.outline[12].includes('3 张图'),
        `read() 的网格行应报出图片张数：${imgBack.outline[12]}`);
    assert(imgBack.outline[0].includes('cm'), `read() outline 应含尺寸：${imgBack.outline[0]}`);
    assert(imgBack.outline[1].includes('（有裁切）'), `read() outline 应标裁切：${imgBack.outline[1]}`);
    assert(imgBack.outline[7].includes('（满页背景）'), `read() outline 应标满页背景：${imgBack.outline[7]}`);
    console.log(`图片样例 warnings(${imgReport.warnings.length})`);
    for (const line of imgReport.warnings) console.log(`  - ${line}`);
}

// ── 8. 自由绘制：形状 / 直线 / 图标 / 富文本 run / 页面母版 / 表格合并 ──────
//
// 这一整节用独立的 deck：主样例已经断言了「10 页 / 52 个部件 / 0 条警告」，
// 往里面塞形状会连带改掉那些数字，反而看不清是自由绘制坏了还是主样例变了。

const xOf = (xml, name) => {
    // 取某个 cNvPr name 的 p:sp / p:pic 的几何：形状 XML 里 a:off/a:ext 只有一组
    const at = xml.indexOf(`name="${name}"`);
    if (at === -1) return undefined;
    const chunk = xml.slice(at, at + 3000);
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(chunk);
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(chunk);
    return off === null || ext === null ? undefined : {
        x: Number(off[1]), y: Number(off[2]), w: Number(ext[1]), h: Number(ext[2]),
    };
};
const countOf = (text, pattern) => (text.match(pattern) ?? []).length;
// 只从 a:rPr 里取 run 的属性：同样形状的 a:solidFill 也出现在形状填充和描边里，
// 不按 run 切出来就会把「形状的底色」当成「run 的颜色」
const runProps = (xml) => [...xml.matchAll(/<a:rPr ([^>]*)>([\s\S]*?)<\/a:rPr>/g)].map((match) => ({
    attrs: match[1],
    color: /<a:srgbClr val="([0-9A-F]{6})"/.exec(match[2])?.[1],
    xml: match[0],
}));

// deck.shape 支持的 preset 清单（与 SPEC / meta 里写的一致）
const PRESETS = [
    'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'hexagon',
    'chevron', 'pentagon', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'bentArrow',
    'curvedRightArrow', 'blockArc', 'donut', 'pie', 'teardrop', 'cloud', 'star5', 'heart',
    'plaque', 'frame', 'halfFrame', 'corner', 'flowChartProcess', 'flowChartDecision',
    'flowChartTerminator', 'flowChartData', 'roundedRectCallout', 'ovalCallout',
];

const canvas = create({ title: '自由绘制样例', theme: 'plain', path: 'canvas.pptx', size: '16:9' }, env);
PRESETS.forEach((preset, index) => {
    canvas.shape({
        preset,
        x: 0.4 + (index % 8) * 4.0,
        y: 1.0 + Math.floor(index / 8) * 1.7,
        w: 2.4,
        h: 1.2,
        fill: '1D4ED8',
        line: '0F172A',
    });
});
// 未知 preset、旋转、圆角、半透明、虚线、双端箭头、阴影、出血坐标、形状内文字
canvas.shape({
    preset: 'star7', x: 30, y: 12, w: 2, h: 2,
});
canvas.shape({
    preset: 'roundRect', x: -1.2, y: -0.8, w: 6.0, h: 3.0,
    rotate: 30, radius: 0.3,
    fill: { color: 'B91C1C', alpha: 0.5 },
    line: {
        color: '111111', widthPt: 2, dash: 'sysDot', alpha: 0.5,
        arrow: { begin: 'oval', end: 'arrow' },
    },
    shadow: { blurPt: 4, distPt: 2, dirDeg: 45, color: '000000', alpha: 0.35 },
    text: '出血卡片',
    textOpts: { size: 14, bold: true, color: 'FFFFFF', align: 'ctr', anchor: 'ctr', marginCm: 0.1, font: '黑体' },
});
// 渐变（含 stop 透明度与角度）、stops 不足 2 个、未知 dash
canvas.shape({
    preset: 'rect', x: 1, y: 12, w: 8, h: 2, fill: 'FFFFFF',
    line: { color: '334155', dash: 'wavy' },
});
canvas.shape({
    preset: 'rect', x: 1, y: 14.2, w: 8, h: 1.2,
    fill: {
        type: 'gradient', angle: 90,
        stops: [{ pos: 0, color: 'B91C1C' }, { pos: 1, color: 'F97316', alpha: 0.8 }],
    },
});
canvas.shape({
    preset: 'rect', x: 1, y: 15.6, w: 8, h: 1.2,
    fill: { type: 'gradient', stops: [{ pos: 0, color: '123456' }] },
});
canvas.shape({ preset: 'ellipse', x: 10, y: 14.2, w: 2, h: 2, fill: 'none' });
// 图标：码点 / U+ 写法 / 单字符 + 自定义字体
canvas.icon({ glyph: 0xE72C, x: 20.0, y: 2.0, sizeCm: 1.2, color: 'B91C1C' });
canvas.icon({ glyph: 'U+E8B7', x: 21.6, y: 2.0, sizeCm: 1.0 });
canvas.icon({ glyph: '★', x: 23.2, y: 2.0, sizeCm: 1.0, font: 'Arial' });
// 直线：正方向、反方向（flipH/flipV）、双端箭头、重合点、缺 to
canvas.line({ from: [1, 12], to: [10, 15], color: '0F172A', widthPt: 2, dash: 'lgDash', arrow: 'end' });
canvas.line({ from: [12, 15], to: [5, 12], color: 'B91C1C', widthPt: 1, alpha: 0.6, arrow: 'both' });
canvas.line({ from: [2, 17], to: [2, 17] });
canvas.line({ from: [3, 17] });

const canvasReport = canvas.save();
const canvasParts = unzip(canvas.render());
const canvasText = new Map([...canvasParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
const c1 = canvasText.get('ppt/slides/slide1.xml') ?? '';

equal(canvasReport.stats.slides, 1, '连续画形状应聚在同一页');
equal(canvasReport.stats.layouts.shape, 1, '自由绘制页在 stats.layouts 里自成一类');
equal(canvasReport.stats.tables, 0, '自由绘制样例没有表格');
console.log(`自由绘制样例 warnings(${canvasReport.warnings.length})`);
for (const line of canvasReport.warnings) console.log(`  - ${line}`);
console.log(`自由绘制样例 outline  ${canvasReport.outline[0]}`);
console.log(`自由绘制样例 shapeKinds  ${JSON.stringify(canvasReport.stats.shapeKinds)}`);

// 口语名 → DrawingML 真名（写原名会让真实 PowerPoint 判文件损坏，见交付报告）
const PRST_ALIASES = {
    flowChartData: 'flowChartInputOutput',
    roundedRectCallout: 'wedgeRoundRectCallout',
    ovalCallout: 'wedgeEllipseCallout',
};
// 每个 preset 都必须原样落到 a:prstGeom prst="…"（三个口语名落到对应的真名）
for (const preset of PRESETS) {
    const expected = PRST_ALIASES[preset] ?? preset;
    assert(c1.includes(`<a:prstGeom prst="${expected}">`), `缺少 preset 的几何：${preset} → ${expected}`);
    assert((canvasReport.stats.shapeKinds[preset] ?? 0) >= 1, `stats.shapeKinds.${preset}`);
}
assert(!/<a:prstGeom prst="(flowChartData|roundedRectCallout|ovalCallout)">/.test(c1),
    '不应把 DrawingML 里不存在的口语名写进 prstGeom');
equal(PRESETS.length, 32, 'preset 清单长度');
assert(c1.includes('<a:prstGeom prst="line">'), '直线形状应写 prstGeom prst="line"');
equal(canvasReport.stats.shapeKinds.line, 4, 'stats.shapeKinds.line');
equal(canvasReport.stats.shapeKinds.icon, 3, 'stats.shapeKinds.icon');
// 未知 preset → 回落 rect + warning（含 4 个显式 rect 与 icon 的 rect 不计入这里）
equal(canvasReport.stats.shapeKinds.rect, 5, '未知 preset 回落 rect 后 rect 共 5 个（4 个显式 + 1 个回落）');
assert(canvasReport.warnings.some((line) => line.includes('preset「star7」不认识，已回落 rect')), '应报告未知 preset');
// 旋转 30° → rot=1800000（60000/度）
assert(c1.includes('rot="1800000"'), 'rotate 应写 a:xfrm rot（60000/度）');
// 出血坐标：允许负值
assert(c1.includes('<a:off x="-432000" y="-288000"/>'), '形状应允许负坐标（出血到页边）');
// 圆角：0.3cm / 短边 3cm → adj=10000（千分之一）
assert(c1.includes('<a:gd name="adj" fmla="val 10000"/>'), 'roundRect 的 radius 应写成 a:avLst adj');
// 半透明填充与描边：alpha 0.5 → 50000（千分之一百分比）
equal(countOf(c1, /<a:alpha val="50000"\/>/g), 2, '填充与描边各写一个 50% 透明度');
// 虚线：sysDot 原样写，wavy 回落 solid
assert(c1.includes('<a:prstDash val="sysDot"/>'), 'dash=sysDot 应原样写入');
assert(canvasReport.warnings.some((line) => line.includes('未知的虚线样式「wavy」')), '应报告未知 dash');
// 箭头端点：begin → headEnd，end → tailEnd，都在 prstDash 之后
{
    const arrow = c1.slice(c1.indexOf('<a:ln w="25400">'), c1.indexOf('<a:ln w="25400">') + 400);
    assert(arrow.includes('<a:headEnd type="oval" w="med" len="med"/>'), 'line.arrow.begin 应写 a:headEnd');
    assert(arrow.includes('<a:tailEnd type="arrow" w="med" len="med"/>'), 'line.arrow.end 应写 a:tailEnd');
    assert(arrow.indexOf('<a:prstDash') < arrow.indexOf('<a:headEnd'), 'a:headEnd 必须排在 a:prstDash 之后');
}
// 阴影：blurPt/distPt → EMU，dirDeg → 60000/度
assert(c1.includes('<a:outerShdw blurRad="50800" dist="25400" dir="2700000" rotWithShape="0">'), '阴影参数换算');
assert(c1.includes('<a:alpha val="35000"/>'), '阴影默认 35% 黑');
// 形状内文字：字体/字号/锚点/内边距
assert(c1.includes('<a:t>出血卡片</a:t>'), '形状内文字应写进 a:t');
assert(c1.includes('<a:latin typeface="黑体"/>'), 'textOpts.font 应覆盖 a:latin');
assert(c1.includes('lIns="36000"'), 'textOpts.marginCm 应写成 bodyPr 内边距');
assert(c1.includes('anchor="ctr"'), 'textOpts.anchor 应写进 bodyPr');
// 渐变：stops 的 pos 与 alpha、以及角度 → a:lin ang
{
    const grad = c1.slice(c1.indexOf('<a:gradFill'));
    assert(grad.includes('<a:gs pos="0"><a:srgbClr val="B91C1C"></a:srgbClr></a:gs>'), '渐变第一个色标');
    assert(grad.includes('<a:gs pos="100000"><a:srgbClr val="F97316"><a:alpha val="80000"/></a:srgbClr></a:gs>'),
        '渐变末色标带 alpha');
    assert(grad.includes('<a:lin ang="5400000" scaled="0"/>'), '渐变角度 90° → ang=5400000');
}
assert(canvasReport.warnings.some((line) => line.includes('stops 只有 1 个，少于 2 个')), '应报告渐变 stops 不足');
// 图标：字形、默认字体、自定义字体、居中
assert(c1.includes(`<a:t>${String.fromCodePoint(0xE72C)}</a:t>`), 'icon 的码点应转成字形');
assert(c1.includes(`<a:t>${String.fromCodePoint(0xE8B7)}</a:t>`), 'icon 应接受 U+ 写法');
assert(c1.includes('<a:latin typeface="Segoe MDL2 Assets"/>'), 'icon 默认字体应是 Segoe MDL2 Assets');
assert(c1.includes('<a:latin typeface="Arial"/>'), 'icon 的自定义字体应生效');
// 直线：正方向与 flipH/flipV、宽度磅 → EMU
{
    const lines = [...c1.matchAll(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?prst="line"(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g)]
        .map((match) => match[0]);
    equal(lines.length, 4, '应有 4 条直线形状');
    const forward = lines.find((xml) => xml.includes('<a:off x="360000" y="4320000"/>'));
    assert(forward !== undefined, '第一条直线应从 (1cm,12cm) 起');
    assert(!forward.includes('flipH'), '向右的直线不应写 flipH');
    assert(forward.includes('cx="3240000" cy="1080000"'), '直线长度应等于两点差值');
    assert(forward.includes('<a:ln w="25400">'), 'widthPt=2 → w=25400');
    const backward = lines.find((xml) => xml.includes('<a:off x="1800000" y="4320000"/>'));
    assert(backward !== undefined, '反向直线应以左上角定位');
    assert(backward.includes('flipH="1"') && backward.includes('flipV="1"'), '向左上的直线应写 flipH/flipV');
    assert(countOf(backward, /<a:headEnd type="triangle"/g) === 1, 'arrow=both 应有 headEnd');
    assert(countOf(backward, /<a:tailEnd type="triangle"/g) === 1, 'arrow=both 应有 tailEnd');
}
assert(canvasReport.warnings.some((line) => line.includes('起点与终点重合')), '应报告直线起止点重合');
assert(canvasReport.warnings.some((line) => line.includes('需要 from:[x,y] 与 to:[x,y]')), '应报告直线缺端点');
assert(canvasReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
    `自由绘制告警里不应出现 undefined/NaN：${canvasReport.warnings.join(' / ')}`);

// read() 复检：自由绘制页必须能被认回来
{
    const canvasBack = read(canvasReport.path, env);
    equal(canvasBack.pages[0].layout, 'shape', 'read() 应把自由绘制页认成 shape');
    equal(canvasBack.stats.layouts.shape, 1, 'read().stats.layouts.shape');
    assert(canvasBack.pages[0].autoShapes >= PRESETS.length, `read() 的形状数应不小于 preset 数：${canvasBack.pages[0].autoShapes}`);
    equal(canvasBack.pages[0].mergedCells, 0, '自由绘制页没有合并单元格');
    assert(canvasBack.outline[0].startsWith('P1 shape：'), `read() outline：${canvasBack.outline[0]}`);
    assert(canvasBack.outline[0].includes('个形状'), `read() outline 应含形状数：${canvasBack.outline[0]}`);
}

// ── 8.1 页面母版：参考稿式整页 ─────────────────────────────────────────────

env.writeFile('assets/bg.png', makePng(1400, 787));
env.writeFile('assets/logo.png', makePng(200, 200));

const ref = create({ title: '参考稿', theme: 'plain', path: 'reference.pptx', size: '16:9', author: '测试' }, env);
ref.master({
    background: { image: { path: 'assets/bg.png', fit: 'cover', dim: 0.45, dimColor: '7F1D1D' } },
    logo: { path: 'assets/logo.png', corner: 'top-right', widthCm: 1.6, marginCm: 0.5 },
    header: { text: '××大学 计算机学院 · 2026 年度学术汇报', size: 14, color: '334155' },
    footer: '内部资料，请勿外传',
    pageNumber: { show: true, corner: 'bottom-right', from: 1, format: '{n} / {total}', size: 10, color: '64748B' },
    accent: [
        { bar: 'left-top', color: 'B91C1C', sizeCm: 0.35, lengthCm: 4.5 },
        { bar: 'bottom', color: '0F766E', sizeCm: 0.2 },
    ],
    skipLayouts: [],
});
// 4 张圆角卡片 + 1 条渐变横幅 + 1 段三色混排文字：一页里同时出现
// 卡片底色刻意用 F8FAFC 而不是 FFFFFF：FFFFFF 按约定等于「不填充」，卡片会变透明
for (let i = 0; i < 4; i += 1) {
    ref.shape({
        preset: 'roundRect', x: 1.2 + i * 7.9, y: 4.4, w: 7.0, h: 5.6,
        fill: 'F8FAFC', line: { color: 'E5E7EB', widthPt: 1 }, radius: 0.25, shadow: true,
        text: `卡片 ${i + 1}`,
        textOpts: { size: 14, bold: true, color: '0F172A', align: 'l', anchor: 't', marginCm: 0.3 },
    });
}
ref.shape({
    preset: 'rect', x: 1.2, y: 2.6, w: 30.6, h: 1.5,
    fill: {
        type: 'gradient', angle: 0,
        stops: [{ pos: 0, color: 'B91C1C' }, { pos: 1, color: 'F97316', alpha: 0.9 }],
    },
    text: '2026 年度研究进展',
    textOpts: { size: 20, bold: true, color: 'FFFFFF', align: 'ctr', anchor: 'ctr' },
});
ref.shape({
    preset: 'rect', x: 1.2, y: 10.6, w: 30, h: 1.2,
    textOpts: {
        size: 16, lineSpacing: 1.2, align: 'l', anchor: 'ctr',
        runs: [
            { text: '关键词：', color: 'B91C1C', bold: true, spacing: 120 },
            { text: '多模态检索', color: '1D4ED8', italic: true, underline: true },
            { text: ' 与 ', color: '64748B' },
            { text: '可控生成', color: '0F766E', size: 18, bold: true },
        ],
    },
});

const refReport = ref.save();
const refParts = unzip(ref.render());
const refText = new Map([...refParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
const r1 = refText.get('ppt/slides/slide1.xml') ?? '';

equal(refReport.stats.slides, 1, '参考稿只有一页');
equal(refReport.stats.layouts.shape, 1, '参考稿是自由绘制页');
equal(refReport.stats.imageParts, 2, '参考稿应打包背景图与 logo 两个媒体部件');
console.log(`参考稿 warnings(${refReport.warnings.length})`);
for (const line of refReport.warnings) console.log(`  - ${line}`);
console.log(`参考稿 outline  ${refReport.outline[0]}`);
assert(refReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
    `参考稿告警里不应出现 undefined/NaN：${refReport.warnings.join(' / ')}`);

// 背景层：满页 p:pic + 半透明压暗矩形
{
    const bgOff = r1.indexOf('<a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/>');
    assert(bgOff !== -1, '满页背景图应铺满整页');
    const pic = r1.slice(0, r1.indexOf('</p:pic>') + 7);
    assert(pic.includes('<p:pic>'), '背景图应写成 p:pic');
    assert(pic.includes('Picture Full Bleed'), '满页背景图应有 Full Bleed 命名');
    assert(pic.includes('r:embed="rId2"'), '背景图应有一条图片关系');
    assert(r1.includes('name="Background Dim"'), 'dim>0 时应画压暗矩形');
    assert(r1.includes('<a:srgbClr val="7F1D1D"><a:alpha val="45000"/></a:srgbClr>'), 'dim 0.45 → alpha 45000，颜色用 dimColor');
}
// logo：右上角，宽 1.6cm，原图等比
{
    const logo = xOf(r1, 'Master Logo');
    assert(logo !== undefined, 'logo 应是一张 p:pic');
    if (logo !== undefined) {
        equal(logo.w, cmToEmu(1.6), 'logo 宽度按 widthCm');
        equal(logo.h, cmToEmu(1.6), '方形 logo 高度应与宽度一致');
        equal(logo.x, 12192000 - cmToEmu(0.5) - cmToEmu(1.6), 'logo 应贴右上角');
        equal(logo.y, cmToEmu(0.5), 'logo 上边距');
    }
}
// 页眉 / 页脚 / 页码 / 装饰条
assert(r1.includes('××大学 计算机学院 · 2026 年度学术汇报'), '页眉文字应写入');
assert(r1.includes('name="Master Header"'), '页眉应有独立文本框');
assert(r1.includes('name="Master Footer"'), '页脚应有独立文本框');
assert(r1.includes('<a:t>1 / 1</a:t>'), '页码应替换 {n} 与 {total}');
{
    const badge = xOf(r1, 'Master Accent 1');
    equal(badge?.w, cmToEmu(0.35), '左上装饰条的宽度 = sizeCm');
    equal(badge?.h, cmToEmu(4.5), '左上装饰条的长度 = lengthCm');
    const bottom = xOf(r1, 'Master Accent 2');
    equal(bottom?.y, 6858000 - cmToEmu(0.2), '底部装饰条应贴下边');
    equal(bottom?.w, 12192000, '底部装饰条默认铺满整页宽');
}
// 渐变横幅 + 四张卡片 + 段内三色文字
{
    const grad = r1.slice(r1.indexOf('<a:gradFill'));
    assert(grad.includes('<a:gs pos="0"><a:srgbClr val="B91C1C"></a:srgbClr></a:gs>'), '横幅渐变的起点色');
    assert(grad.includes('<a:alpha val="90000"/>'), '横幅渐变末色标的 90% 透明度');
    assert(grad.includes('<a:lin ang="0" scaled="0"/>'), '横幅渐变角度 0°');
    equal(countOf(r1, /prst="roundRect"/g), 4, '一页里应有 4 张圆角卡片');
    equal(countOf(r1, /<a:outerShdw /g), 4, '4 张卡片各有一个阴影');
    equal(countOf(r1, /<a:t>卡片 /g), 4, '4 张卡片各有一段文字');
    // 段内三色：同一段里四条 run，颜色/字号/加粗各异
    {
        const colors = runProps(r1).map((run) => run.color);
        for (const color of ['B91C1C', '1D4ED8', '64748B', '0F766E']) {
            assert(colors.includes(color), `段内混排缺少顏色 ${color}（实际 ${colors.join(',')}）`);
        }
        const last = runProps(r1).find((run) => run.color === '0F766E');
        assert(last !== undefined && last.attrs.includes('sz="1800"'), 'run 自己的字号（18pt）应生效');
        assert(last !== undefined && last.attrs.includes('b="1"'), 'run 自己的加粗应生效');
        assert(runProps(r1).some((run) => run.attrs.includes('sz="1600"')), '未写字号的 run 应继承段落基值');
        assert(runProps(r1).some((run) => run.attrs.includes('spc="120"')), 'run 的 spacing 应写 a:rPr spc');
        const blue = runProps(r1).find((run) => run.color === '1D4ED8');
        assert(blue !== undefined && blue.attrs.includes('i="1"') && blue.attrs.includes('u="sng"'),
            'run 的 italic / underline 应写成 i="1" u="sng"');
        // 关键词那段四条 run 必须同属一个 <a:p>
        const pStart = r1.lastIndexOf('<a:p>', r1.indexOf('<a:t>关键词：</a:t>'));
        const pEnd = r1.indexOf('</a:p>', r1.indexOf('<a:t>可控生成</a:t>'));
        const para = r1.slice(pStart, pEnd);
        equal(countOf(para, /<a:r>/g), 4, '四条 run 必须在同一段里');
    }
}
// 图层顺序：背景图在最底、内容居中、母版叠加层在最上
{
    const order = ['Picture Full Bleed', 'Background Dim', 'roundRect', 'Master Header', 'Master Page Number', 'Master Accent 2']
        .map((token) => r1.indexOf(token));
    assert(order.every((at) => at !== -1), '图层元素都应存在');
    for (let i = 1; i < order.length; i += 1) {
        assert(order[i - 1] < order[i], `图层顺序错误：${i}（${order.join(',')}）`);
    }
}
// read() 复检：背景图 + 形状都要读得回来
{
    const refBack = read(refReport.path, env);
    equal(refBack.pages[0].layout, 'shape', '参考稿 read() 应认成 shape');
    equal(refBack.pages[0].fullBleed, true, '参考稿应被读成有满页背景图');
    equal(refBack.stats.imageParts, 2, '参考稿 read() 媒体部件数');
    assert(refBack.outline[0].includes('背景图'), `read() outline 应标出背景图：${refBack.outline[0]}`);
}

// 母版：默认 skipLayouts（封面/结尾）不加页眉/页脚/页码，但背景/logo/装饰条照加
{
    const md = create({ title: '母版缺省', theme: 'plain', path: 'master-default.pptx', size: '16:9', author: '测试' }, env);
    md.master({
        logo: { path: 'assets/logo.png', corner: 'bottom-left', widthCm: 1.0 },
        header: '某某大学 · 汇报',
        footer: '内部资料',
        pageNumber: { show: true, from: 1, format: '第 {n} 页 / 共 {total} 页' },
        accent: { bar: 'top', color: 'B91C1C', sizeCm: 0.2 },
    });
    md.cover({ title: '封面标题' });
    md.bullets({ title: '要点页', items: ['第一条', '第二条'] });
    md.closing({ title: '谢谢' });
    const mdReport = md.save();
    const mdParts = unzip(md.render());
    const mdText = new Map([...mdParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
    const cover = mdText.get('ppt/slides/slide1.xml') ?? '';
    const body = mdText.get('ppt/slides/slide2.xml') ?? '';
    const close = mdText.get('ppt/slides/slide3.xml') ?? '';
    assert(cover.includes('Master Logo') || cover.includes('Picture 2'), '封面应有 logo');
    assert(!cover.includes('name="Master Header"'), '封面默认不加页眉');
    assert(!cover.includes('name="Master Footer"'), '封面默认不加页脚');
    assert(!cover.includes('name="Master Page Number"'), '封面默认不加页码');
    assert(cover.includes('name="Master Accent 1"'), '封面仍应加装饰条');
    assert(body.includes('name="Master Header"'), '内容页应有页眉');
    assert(body.includes('name="Master Footer"'), '内容页应有页脚');
    assert(body.includes('<a:t>第 2 页 / 共 3 页</a:t>'), '页码应使用真实页序号（1 基）');
    assert(!body.includes('name="Footer Rule"'), '配了母版后不应再画版式自带的页脚线');
    assert(!body.includes('name="Page Number"'), '配了母版后不应再画版式自带的页码');
    assert(close.includes('Master Accent 1'), '结尾页仍应加装饰条');
    assert(!close.includes('name="Master Page Number"'), '结尾页默认不加页码');
    assert(!body.includes('第 1 页'), '页码不应把封面算成 1');
    equal(mdReport.stats.layouts.shape, 0, '这个样例没有自由绘制页');
    assert(mdReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
        `母版告警里不应出现 undefined/NaN：${mdReport.warnings.join(' / ')}`);
}
// 母版图片缺失 / 角标写错 / 背景不填白
{
    const bad = create({ title: '母版缺图', theme: 'plain', path: 'master-bad.pptx', size: '16:9' }, env);
    bad.master({
        background: { image: { path: 'assets/nope.png' } },
        logo: { path: 'assets/nope.png', corner: 'middle' },
        accent: { bar: 'diagonal' },
    });
    bad.shape({ preset: 'rect', x: 1, y: 1, w: 2, h: 2, fill: 'FFFFFF' });
    const badReport = bad.save();
    assert(badReport.warnings.some((line) => line.includes('母版背景图 缺失')), '应报告母版背景图缺失');
    assert(badReport.warnings.some((line) => line.includes('母版 logo 缺失')), '应报告母版 logo 缺失');
    assert(badReport.warnings.some((line) => line.includes('corner「middle」')), '应报告非法角标');
    assert(badReport.warnings.some((line) => line.includes('bar「diagonal」')), '应报告非法装饰条');
    const badParts = unzip(bad.render());
    const badXml = new TextDecoder('utf-8').decode(badParts.get('ppt/slides/slide1.xml'));
    assert(badXml.includes('name="Missing Image"'), '缺图应画占位框');
    const shapeAt = badXml.indexOf('name="Shape');
    const shapeSeg = badXml.slice(shapeAt, shapeAt + 600);
    assert(shapeSeg.includes('<a:noFill/>'), 'fill=FFFFFF 视为不填充');
    assert(!shapeSeg.includes('FFFFFF'), 'FFFFFF 不应写进形状填充');
    // deck.background 只影响当前页
    const per = create({ title: '单页背景', theme: 'plain', path: 'per-page.pptx' }, env);
    per.shape({ preset: 'rect', x: 1, y: 1, w: 2, h: 2, fill: '1D4ED8' });
    per.background({ color: 'F1F5F9' });
    per.page();
    per.shape({ preset: 'rect', x: 1, y: 4, w: 2, h: 2, fill: '1D4ED8' });
    per.background({ gradient: { stops: [{ pos: 0, color: 'B91C1C' }, { pos: 1, color: '1D4ED8' }], angle: 90 } });
    const perParts = unzip(per.render());
    const perText = new Map([...perParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
    assert((perText.get('ppt/slides/slide1.xml') ?? '').includes('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="F1F5F9"/>'),
        'deck.background 应把当前页的 p:bg 写成指定颜色');
    equal(per.inspect().stats.slides, 2, 'deck.page() 之后继续画形状应另起一页');
    assert(!(perText.get('ppt/slides/slide2.xml') ?? '').includes('val="F1F5F9"'), '第 2 页不应继承第 1 页的背景');
    assert((perText.get('ppt/slides/slide2.xml') ?? '').includes('<p:bg><p:bgPr><a:gradFill'),
        '渐变背景应写成 p:bgPr 里的 a:gradFill');
}

// ── 8.2 表格合并 / 单元格样式 / 链接 / 富文本 ─────────────────────────────

const mt = create({ title: '合并表格', theme: 'plain', path: 'merge.pptx', size: '16:9' }, env);
mt.table({
    title: '合并与样式',
    columns: [
        { title: '项目', width: 8 },
        { title: 'Q1', align: 'center' },
        { title: 'Q2', align: 'center' },
        { title: '备注' },
    ],
    rows: [
        [{ text: '平台底座', colSpan: 3, fill: 'FEF3C7', bold: true }, '跨越三列'],
        [{ text: '办公', rowSpan: 2 }, 'A1', 'A2', { text: '官网', link: 'https://example.com/a' }],
        ['B1', { runs: [{ text: '红', color: 'DC2626' }, { text: '蓝', color: '2563EB', bold: true }] }, { text: '越界', colSpan: 9 }],
    ],
    headerFill: '0E7490',
    zebra: false,
    border: { color: '94A3B8', widthPt: 0.75 },
    cellPadCm: 0.12,
    rowHeightCm: 1.0,
});
mt.table({
    title: '斑马纹与表头样式',
    columns: ['名称', '值'],
    rows: [['甲', '1'], ['乙', '2'], ['丙', '3']],
    zebra: 'F1F5F9',
    firstRowBold: false,
});
const mtReport = mt.save();
const mtParts = unzip(mt.render());
const mtText = new Map([...mtParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
const t1 = mtText.get('ppt/slides/slide1.xml') ?? '';
const t2 = mtText.get('ppt/slides/slide2.xml') ?? '';
console.log(`合并表格 warnings(${mtReport.warnings.length})`);
for (const line of mtReport.warnings) console.log(`  - ${line}`);

equal(mtReport.stats.tables, 2, '两张表');
equal(mtReport.stats.mergedCells, 2, '合并区域数：colSpan=3 一处 + rowSpan=2 一处');
equal(countOf(t1, /<a:gridCol /g), 4, '网格列数');
equal(countOf(t1, /<a:tr /g), 4, '行数（含表头）');
{
    const rows = [...t1.matchAll(/<a:tr [^>]*>([\s\S]*?)<\/a:tr>/g)].map((match) => match[1]);
    equal(rows.length, 4, '解析出的行数');
    for (const row of rows) {
        equal(countOf(row, /<a:tc[ >]/g), 4, '每行的 a:tc 个数必须等于网格列数（含被覆盖的格子）');
    }
    // 第 1 行：colSpan=3 的起点格 + 2 个 hMerge
    assert(rows[1].includes('<a:tc gridSpan="3">'), '合并起点格应写 a:tc gridSpan');
    equal(countOf(rows[1], /hMerge="1"/g), 2, '被横向合并覆盖的格子应写 hMerge="1"');
    // 第 2 行：rowSpan=2 的起点格
    assert(rows[2].includes('<a:tc rowSpan="2">'), '纵向合并起点格应写 a:tc rowSpan');
    // 第 3 行：第 1 列被上一行的 rowSpan 覆盖 → vMerge
    equal(countOf(rows[3], /vMerge="1"/g), 1, '被纵向合并覆盖的格子应写 vMerge="1"');
}
assert(mtReport.warnings.some((line) => line.includes('colSpan=9 超出 4 列，已裁剪为 1')), '应报告 colSpan 越界裁剪');
// 样式
assert(t1.includes('<a:srgbClr val="0E7490"/>'), 'headerFill 应生效');
assert(t1.includes('<a:srgbClr val="FEF3C7"/>'), '单元格自定义 fill 应生效');
assert(!t1.includes('<a:srgbClr val="F1F5F9"/>'), 'zebra:false 时不应有斑马纹底色');
assert(t1.includes('<a:lnL w="9525"><a:solidFill><a:srgbClr val="94A3B8"/>'), 'border 应给四边画线');
assert(t1.includes('marL="43200"'), 'cellPadCm=0.12 → 单元格内边距');
for (const match of t1.matchAll(/<a:tr h="(\d+)"/g)) {
    assert(Number(match[1]) >= cmToEmu(1.0), `rowHeightCm 应抬高行高，实际 ${match[1]}`);
}
// 链接：真关系 + 主题链接色 + 下划线
{
    assert(t1.includes('<a:hlinkClick r:id="rId2"/>'), '带 link 的单元格应写 a:hlinkClick');
    assert(t1.includes('u="sng"'), '链接单元格应加下划线');
    const rels = mtText.get('ppt/slides/_rels/slide1.xml.rels') ?? '';
    assert(rels.includes('TargetMode="External"'), '超链接关系必须是 External');
    assert(rels.includes('Target="https://example.com/a"'), '超链接关系的目标');
    assert(rels.includes('relationships/hyperlink'), '超链接关系的类型');
}
// 富文本 run：同一格里两种颜色，其中一个加粗
{
    assert(t1.includes('<a:t>红</a:t>') && t1.includes('<a:t>蓝</a:t>'), '单元格 runs 应逐条写入');
    const red = t1.slice(t1.indexOf('<a:t>红</a:t>') - 300, t1.indexOf('<a:t>红</a:t>'));
    const blue = t1.slice(t1.indexOf('<a:t>蓝</a:t>') - 300, t1.indexOf('<a:t>蓝</a:t>'));
    assert(red.includes('val="DC2626"'), 'run 自己的颜色');
    assert(blue.includes('val="2563EB"') && blue.includes('b="1"'), 'run 自己的颜色与加粗');
}
{
    const firstRow = /<a:tr [^>]*>([\s\S]*?)<\/a:tr>/.exec(t2)?.[1] ?? '';
    assert(!firstRow.includes('b="1"'), 'firstRowBold:false 时表头不应加粗');
    assert(t2.includes('<a:srgbClr val="F1F5F9"/>'), 'zebra 传色值时应用该颜色');
}
{
    const mtBack = read(mtReport.path, env);
    equal(mtBack.pages[0].mergedCells, 2, 'read() 应数出 2 处合并');
    assert(mtBack.outline[0].includes('（含 2 处合并）'), `read() outline 应标出合并：${mtBack.outline[0]}`);
    equal(mtBack.stats.mergedCells, 2, 'read().stats.mergedCells');
    assert(mtBack.outline[0].includes('4 行 × 4 列'), `read() outline 表格尺寸：${mtBack.outline[0]}`);
}
// 合并越界的两种方向都要报
{
    const over = create({ title: '越界合并', theme: 'plain', path: 'merge-over.pptx' }, env);
    over.table({
        title: '越界',
        columns: ['A', 'B', 'C'],
        rows: [
            [{ text: '横', colSpan: 5 }],
            [{ text: '纵', rowSpan: 4 }, 'x', 'y'],
        ],
    });
    const overReport = over.save();
    assert(overReport.warnings.some((line) => line.includes('colSpan=5 超出 3 列')), '应报告列方向越界');
    assert(overReport.warnings.some((line) => line.includes('rowSpan=4 超出 2 行')), '应报告行方向越界');
}

// ── 8.3 runs 在所有既有版式里都能用 ────────────────────────────────────────

{
    const rich = create({ title: '富文本', theme: 'plain', path: 'runs.pptx', size: '16:9' }, env);
    rich.cover({
        title: '封面主标题',
        runs: [{ text: '封面', color: 'B91C1C' }, { text: '主标题', color: '1D4ED8', size: 30, bold: true }],
    });
    rich.bullets({
        title: '要点混排',
        items: [
            { text: '第一条', level: 1, runs: [{ text: '第一条', color: '0F766E', bold: true }, { text: '（补充）', color: '64748B' }] },
            '第二条',
        ],
    });
    rich.quote({ text: '引用', runs: [{ text: '引用', color: 'B91C1C' }, { text: '出处', color: '0F766E', size: 14 }], by: '某人' });
    rich.statement({ text: '宣言', runs: [{ text: '宣言', color: 'FFFFFF' }, { text: '副线', color: 'F97316' }], sub: '副标题' });
    rich.closing({ runs: [{ text: '谢', color: 'FFFFFF' }, { text: '谢', color: 'F97316', bold: true }] });
    const richReport = rich.save();
    const richParts = unzip(rich.render());
    const richText = new Map([...richParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
    const pages = [1, 2, 3, 4, 5].map((page) => richText.get(`ppt/slides/slide${page}.xml`) ?? '');
    // 每一页都必须出现两段颜色不同的 run
    for (const [index, xml] of pages.entries()) {
        const colors = new Set(runProps(xml).map((run) => run.color).filter((color) => color !== undefined));
        assert(colors.size >= 2, `第 ${index + 1} 页应有段内混排 run（实际 ${[...colors].join(',')}）`);
    }
    assert(pages[0].includes('<a:t>封面</a:t>') && pages[0].includes('<a:t>主标题</a:t>'), '封面 runs');
    assert(countOf(pages[0], /<a:r>/g) >= 2, '封面同一段应有两条 run');
    assert(pages[1].includes('<a:t>第一条</a:t>') && pages[1].includes('<a:t>（补充）</a:t>'), '要点 runs');
    assert(pages[1].includes('<a:buChar char="•"/>'), '要点混排不应破坏项目符号');
    assert(pages[2].includes('<a:t>引用</a:t>') && pages[2].includes('<a:t>出处</a:t>'), '引用 runs');
    assert(pages[3].includes('<a:t>宣言</a:t>') && pages[3].includes('<a:t>副线</a:t>'), '宣言 runs');
    assert(pages[4].includes('<a:t>谢</a:t>'), '结尾 runs');
    equal(richReport.stats.slides, 5, '富文本样例页数');
    assert(richReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
        `富文本告警里不应出现 undefined/NaN：${richReport.warnings.join(' / ')}`);
}
// 每一条新 warning 的口径：非法输入只变 warning，不抛
{
    const warnDeck = create({ title: '告警覆盖', theme: 'plain', path: 'warn.pptx' }, env);
    warnDeck.shape({ preset: 'nope', x: 0, y: 0, w: 1, h: 1, radius: 0.2, fill: { type: 'gradient', stops: [{ pos: 0, color: '000000' }] } });
    warnDeck.shape({ preset: 'roundRect', x: 0, y: 0, w: 0, h: 0, radius: 1, line: 'none', shadow: false });
    warnDeck.line({ from: [0, 0], to: [1, 1], arrow: 'sideways', dash: 'dots' });
    warnDeck.icon({});
    const warnReport = warnDeck.save();
    for (const text of [
        'preset「nope」不认识', 'stops 只有 1 个', '缺少宽高', '只有 roundRect 支持 radius',
        'arrow「sideways」不认识', '未知的虚线样式「dots」', 'icon 缺少 glyph',
    ]) {
        assert(warnReport.warnings.some((line) => line.includes(text)), `应覆盖告警：${text}`);
    }
    assert(warnReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
        `告警里不应出现 undefined/NaN：${warnReport.warnings.join(' / ')}`);
}
// meta 必须把新 API 都写进去（Lead 据此写帮助文档）
{
    const createMethods = meta.methods.create.join('\n');
    for (const token of ['deck.shape(', 'deck.line(', 'deck.icon(', 'deck.master(', 'deck.background(', 'runs']) {
        assert(createMethods.includes(token), `meta.methods.create 缺少 ${token}`);
    }
    assert(meta.summary.includes('自由绘制'), 'meta.summary 应写明自由绘制能力');
    // 本轮新增的版式助手都必须出现在 meta 里（签名即文档）
    for (const token of [
        'deck.cards(', 'deck.steps(', 'deck.compare(', 'deck.kpi(', 'deck.imageText(',
        'deck.timeline(', 'deck.iconGrid(', 'deck.panel(', 'deck.banner(',
    ]) {
        assert(createMethods.includes(token), `meta.methods.create 缺少 ${token}`);
    }
    assert(meta.summary.includes('版式助手'), 'meta.summary 应写明版式助手能力');
    // 三条必须写进 meta 的说明（上一轮踩坑的结论）
    for (const token of [
        'flowChartData→flowChartInputOutput',
        'roundedRectCallout→wedgeRoundRectCallout',
        'ovalCallout→wedgeEllipseCallout',
        'deck.page()',
        'FFFFFF 与空串 = 不填充',
    ]) {
        assert(createMethods.includes(token), `meta.methods.create 缺少说明：${token}`);
    }
    assert(createMethods.includes('纯白卡片会变透明'), 'meta 应写明纯白卡片会透明');
}

// ── 8.4 版式助手：cards / steps / compare / kpi / imageText / timeline / iconGrid + panel / banner ──

// 助手页的几何全部按厘米算再统一换算，所以断言里也用同一批厘米常量。
// 版心：x=0.9cm、y=3.25cm、宽 32.07cm、高 14.45cm（16:9）。
const HB = { x: 0.9, y: 3.25, w: 32.07, h: 14.45 };

const hp = create({ title: '版式助手', theme: 'plain', path: 'helpers.pptx', size: '16:9' }, env);
hp.cards({
    title: '卡片网格',
    numbered: true,
    items: [
        { title: '卡片一', body: '正文一', accent: '1D4ED8', icon: 0xE72C },
        { title: '卡片二', body: '正文二' },
        { title: '卡片三', body: '正文三' },
        { title: '卡片四', body: '正文四' },
    ],
});
hp.steps({
    title: '编号流程',
    items: [{ title: '第一步', body: 'a' }, { title: '第二步', body: 'b' }, { title: '第三步', body: 'c' }],
});
hp.compare({
    title: '对比双栏',
    left: { title: '方案 A', items: ['优点一', '优点二'], accent: '0F766E' },
    right: { title: '方案 B', items: ['缺点一', '缺点二'] },
});
hp.kpi({
    title: 'KPI 指标',
    items: [
        { value: '128', unit: '万', label: '营收', color: '0F766E' },
        { value: '92', unit: '%', label: '完成率' },
        { value: '3.4', label: 'NPS' },
    ],
});
hp.imageText({
    title: '图文',
    image: { path: 'assets/chart.png', caption: '示意图' },
    items: ['要点一', '要点二'],
    side: 'right',
});
hp.imageText({ title: '图文缺图', image: { path: 'assets/nope.png' }, items: ['甲'], side: 'left' });
hp.timeline({
    title: '时间线',
    items: [
        { time: 'Q1', title: '一', body: 'aa' }, { time: 'Q2', title: '二', body: 'bb' },
        { time: 'Q3', title: '三', body: 'cc' }, { time: 'Q4', title: '四', body: 'dd' },
    ],
});
hp.iconGrid({
    title: '图标网格',
    items: [
        { icon: 0xE72C, label: '搜索', body: '说明' },
        { icon: 'U+E8B7', label: '设置' },
        { icon: '★', label: '收藏' },
        { icon: 0xE713, label: '标记' },
        { icon: 0xE703, label: '编辑' },
        { icon: 0xE74D, label: '删除' },
    ],
});
const hpReport = hp.save();
const hpParts = unzip(hp.render());
const hpText = new Map([...hpParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
const h1 = hpText.get('ppt/slides/slide1.xml') ?? '';
const h2 = hpText.get('ppt/slides/slide2.xml') ?? '';
const h3 = hpText.get('ppt/slides/slide3.xml') ?? '';
const h4 = hpText.get('ppt/slides/slide4.xml') ?? '';
const h5 = hpText.get('ppt/slides/slide5.xml') ?? '';
const h6 = hpText.get('ppt/slides/slide6.xml') ?? '';
const h7 = hpText.get('ppt/slides/slide7.xml') ?? '';
const h8 = hpText.get('ppt/slides/slide8.xml') ?? '';

console.log(`助手样例 warnings(${hpReport.warnings.length})`);
for (const line of hpReport.warnings) console.log(`  - ${line}`);
console.log('助手样例 outline');
for (const line of hpReport.outline) console.log(`  ${line}`);

// 每个助手在 stats.layouts 里自成一类，且都不算自由绘制页
equal(hpReport.stats.slides, 8, '助手样例页数');
for (const kind of ['cards', 'steps', 'compare', 'kpi', 'timeline', 'iconGrid']) {
    equal(hpReport.stats.layouts[kind], 1, `stats.layouts.${kind}`);
}
equal(hpReport.stats.layouts.imageText, 2, 'stats.layouts.imageText');
equal(hpReport.stats.layouts.shape, 0, '助手页不应被算成自由绘制页');
// outline 一眼能看出「这页用了哪种助手」
assert(hpReport.outline[0].startsWith('P1 cards：'), `outline[0]：${hpReport.outline[0]}`);
assert(hpReport.outline[0].includes('4 张卡片'), `outline[0] 应含卡片数：${hpReport.outline[0]}`);
assert(hpReport.outline[1].startsWith('P2 steps：') && hpReport.outline[1].includes('3 步'), `outline[1]：${hpReport.outline[1]}`);
assert(hpReport.outline[2].includes('双栏对比'), `outline[2]：${hpReport.outline[2]}`);
assert(hpReport.outline[3].includes('3 个指标'), `outline[3]：${hpReport.outline[3]}`);
assert(hpReport.outline[4].includes('图 + 文'), `outline[4]：${hpReport.outline[4]}`);
assert(hpReport.outline[6].includes('4 个节点'), `outline[6]：${hpReport.outline[6]}`);
assert(hpReport.outline[7].includes('6 个图标'), `outline[7]：${hpReport.outline[7]}`);

// cards：2 列网格、圆角卡、序号圆、图标、强调色、正文
{
    equal(countOf(h1, /prst="roundRect"/g), 4, 'cards 应有 4 张圆角卡片');
    equal(countOf(h1, /name="Card \d+"/g), 4, 'cards 应给每张卡起名 Card N');
    equal(countOf(h1, /name="Card \d+ Badge"/g), 4, 'numbered:true 时每张卡一个序号圆');
    equal(countOf(h1, /prst="ellipse"/g), 4, '序号圆用 ellipse 画');
    // 第 1 张卡：x 按列宽落点，宽 (32.07-0.5)/2；高度跟着内容走，整块垂直居中。
    // 2026-09-22 起卡片不再一律撑满版心（cardNeedCm / cardGrid 的 preferH）：
    // 之前 42 页的稿子里 34 页正文区一半以上没字，卡片页平均七成空白。
    const card1 = xOf(h1, 'Card 1');
    equal(card1?.x, cmToEmu(0.9), 'cards 第 1 张的 x');
    equal(card1?.w, cmToEmu(15.785), 'cards 卡片宽（2 列，gap 0.5）');
    const fillH = (HB.h - 0.5) / 2;
    assert(card1.h < cmToEmu(fillH - 0.01),
        `cards 卡片应按内容收高：实际 ${card1.h / 360000}cm，均分高度 ${fillH}cm`);
    assert(card1.h >= cmToEmu(1.8), `cards 卡片高度不应低于下限：${card1.h / 360000}cm`);
    // 上下留白相等 = 整块在版心内垂直居中
    const card3 = xOf(h1, 'Card 3');
    const topGap = card1.y - cmToEmu(HB.y);
    const bottomGap = cmToEmu(HB.y + HB.h) - (card3.y + card3.h);
    assert(Math.abs(topGap - bottomGap) <= 2000,
        `cards 整块应垂直居中：上 ${topGap / 360000}cm / 下 ${bottomGap / 360000}cm`);
    // 第 2 张在同一行右移一列；第 3 张换行
    equal(xOf(h1, 'Card 2')?.x, cmToEmu(17.185), 'cards 第 2 张应右移一列宽 + 间距');
    equal(xOf(h1, 'Card 2')?.y, card1.y, 'cards 第 2 张与第 1 张同一行');
    equal(card3?.y, card1.y + card1.h + cmToEmu(0.5), 'cards 第 3 张应换行');
    equal(card3?.x, cmToEmu(0.9), 'cards 第 3 张回到第 1 列');
    // 末行不满时整行居中：5 张卡排 3 列 → 第 2 行 2 张，左右留白相等
    {
        const lr = create({ title: '末行居中', theme: 'plain', path: 'lastrow.pptx', size: '16:9' }, env);
        lr.cards({
            title: '五张卡',
            columns: 3,
            items: [{ title: 'a' }, { title: 'b' }, { title: 'c' }, { title: 'd' }, { title: 'e' }],
        });
        lr.save();
        const lrParts = unzip(lr.render());
        const lrXml = new TextDecoder('utf-8').decode(lrParts.get('ppt/slides/slide1.xml'));
        const c4 = xOf(lrXml, 'Card 4');
        const c5 = xOf(lrXml, 'Card 5');
        equal(c5?.y, c4?.y, '末行两张卡应在同一行');
        const leftGap = c4.x - cmToEmu(HB.x);
        const rightGap = cmToEmu(HB.x + HB.w) - (c5.x + c5.w);
        assert(Math.abs(leftGap - rightGap) <= 2000,
            `末行应居中：左 ${leftGap / 360000}cm / 右 ${rightGap / 360000}cm`);
    }
    // 图标：0.9cm 见方，落在卡片左上（0.9 + 0.3 内边距）
    const icon1 = xOf(h1, 'Card 1 Icon');
    equal(icon1?.x, cmToEmu(1.2), 'cards 图标的 x');
    equal(icon1?.w, cmToEmu(0.9), 'cards 图标边长 = iconSizeCm');
    assert(h1.includes(`<a:t>${String.fromCodePoint(0xE72C)}</a:t>`), 'cards 的 icon 应走 deck.icon 的字形');
    assert(h1.includes('<a:latin typeface="Segoe MDL2 Assets"/>'), 'cards 图标默认字体');
    // 序号圆跟着图标往右让位
    const badge1 = xOf(h1, 'Card 1 Badge');
    assert((badge1?.x ?? 0) > cmToEmu(1.2) + cmToEmu(0.9), '有图标时段序号圆应排在图标右侧');
    assert(h1.includes('<a:t>1</a:t>') && h1.includes('<a:t>4</a:t>'), '序号圆应写 1~4');
    // 卡片底色：素色主题的 surface 是 FFFFFF（= 不填充），助手必须回落到极浅灰
    // 只取卡片圆角矩形自己的那一段：序号圆上的白字（onColorText）也在同一段附近
    const cardSeg = h1.slice(h1.indexOf('name="Card 1"'), h1.indexOf('name="Card 1 Icon"'));
    assert(cardSeg.includes('val="F8FAFC"'), '素色主题下卡片默认底色应回落到 F8FAFC');
    assert(!cardSeg.includes('val="FFFFFF"'), '卡片底色不应是 FFFFFF（那等于不填充）');
    assert(h1.includes('<a:srgbClr val="1D4ED8"/>'), 'cards 的 accent 应生效');
    assert(h1.includes('<a:outerShdw') === false, '未指定 shadow 时卡片不应有阴影');
    assert(h1.includes('<a:t>卡片一</a:t>') && h1.includes('<a:t>正文一</a:t>'), '卡片标题与正文都应写入');
    assert(h1.includes('<a:buNone/>'), '卡片正文不应带项目符号');
    // 强调细线：贴在头部下方
    const rule1 = xOf(h1, 'Card 1 Rule');
    assert(rule1 !== undefined, 'cards 每张卡应有一条强调细线');
    assert((rule1?.y ?? 0) > cmToEmu(3.25 + 0.3 + 0.95), '强调细线应在头部之下');
}
// cards：显式 cardFill=FFFFFF 仍然等于「不填充」（硬约定）
{
    const white = create({ title: '白卡', theme: 'plain', path: 'cards-white.pptx' }, env);
    white.cards({ title: '白卡', items: [{ title: 'a' }], cardFill: 'FFFFFF' });
    const whiteXml = new TextDecoder('utf-8').decode(unzip(white.render()).get('ppt/slides/slide1.xml'));
    const seg = whiteXml.slice(whiteXml.indexOf('name="Card 1"'), whiteXml.indexOf('name="Card 1 Rule"'));
    assert(seg.includes('<a:noFill/>'), 'cardFill=FFFFFF 应写成不填充');
    assert(!seg.includes('val="FFFFFF"'), 'cardFill=FFFFFF 不应写进形状填充');
}
// cards：主题 surface 直接可用
{
    const biz = create({ title: '商务卡', theme: 'business', path: 'cards-biz.pptx' }, env);
    biz.cards({ title: '商务卡', items: [{ title: 'a' }] });
    const bizXml = new TextDecoder('utf-8').decode(unzip(biz.render()).get('ppt/slides/slide1.xml'));
    assert(bizXml.includes('<a:srgbClr val="F2F6FB"/>'), '有配色主题时卡片底色应取主题 surface');
}
// steps：横向流程、卡片之间的三角箭头、序号圆
{
    equal(countOf(h2, /name="Step Card \d+"/g), 3, 'steps 应有 3 张步骤卡');
    equal(countOf(h2, /name="Step Arrow \d+"/g), 2, '3 步之间应有 2 个箭头');
    equal(countOf(h2, /name="Step Card \d+ Badge"/g), 3, 'numbered 默认 true，应画序号圆');
    const step1 = xOf(h2, 'Step Card 1');
    equal(step1?.x, cmToEmu(0.9), 'steps 第 1 步的 x');
    // 横向：3 步、箭头各占 0.9cm → 每步 (32.07 - 1.8) / 3 = 10.09cm
    equal(step1?.w, cmToEmu(10.09), 'steps 横向每步宽度');
    // 高度跟着内容走（2026-09-22）：三张卡各一行正文，按内容收高；下限 2.6cm、上限版心高
    assert(step1 !== undefined, 'steps 第 1 步应存在');
    assert(step1.h <= cmToEmu(HB.h), `steps 卡片不应超过版心高：${step1.h / 360000}cm`);
    assert(step1.h >= cmToEmu(2.6), `steps 卡片不应低于下限 2.6cm：${step1.h / 360000}cm`);
    assert(step1.h < cmToEmu(5.2), `steps 卡片应按内容收高（旧上限 5.2cm）：${step1.h / 360000}cm`);
    // 整块在版心内垂直居中：上下留白相等
    const stepGapTop = step1.y - cmToEmu(HB.y);
    const stepGapBottom = cmToEmu(HB.y + HB.h) - (step1.y + step1.h);
    assert(Math.abs(stepGapTop - stepGapBottom) <= 2000, 'steps 整块应垂直居中');
    equal(xOf(h2, 'Step Card 2')?.x, cmToEmu(11.89), 'steps 第 2 步应与第 1 步间隔一个箭头位');
    assert(xOf(h2, 'Step Arrow 1') !== undefined, 'steps 应画箭头');
    // 箭头是「线 + a:tailEnd 三角」：与 deck.line 的 arrow:'end' 同一套实现
    {
        const arrow = h2.slice(h2.indexOf('name="Step Arrow 1"'), h2.indexOf('name="Step Arrow 1"') + 900);
        assert(arrow.includes('prst="line"'), '步骤箭头应写成直线形状');
        assert(arrow.includes('<a:tailEnd type="triangle"'), '步骤箭头应写 a:tailEnd 三角');
        assert(!arrow.includes('<a:headEnd'), '步骤箭头只在一端，不应写 a:headEnd');
    }
    const arrow1 = xOf(h2, 'Step Arrow 1');
    equal(arrow1?.x, cmToEmu(11.09), '箭头左端应贴着第 1 步右边（0.1cm 留白）');
    // 箭头纵向落在卡片中间
    assert((arrow1?.y ?? 0) > (step1?.y ?? 0) && (arrow1?.y ?? 0) < (step1?.y ?? 0) + (step1?.h ?? 0),
        '箭头应落在卡片高度之内');
    assert(h2.includes('<a:t>a</a:t>') && h2.includes('<a:t>c</a:t>'), '步骤正文应写入');
}
// steps：纵向 + 关闭序号与箭头
{
    const vt = create({ title: '纵向', theme: 'plain', path: 'steps-v.pptx' }, env);
    vt.steps({ title: '纵向流程', direction: 'vertical', items: [{ title: 'a' }, { title: 'b' }, { title: 'c' }] });
    vt.steps({ title: '裸流程', numbered: false, arrows: false, items: [{ title: 'x' }, { title: 'y' }] });
    const vtXml = new TextDecoder('utf-8').decode(unzip(vt.render()).get('ppt/slides/slide1.xml'));
    const bare = new TextDecoder('utf-8').decode(unzip(vt.render()).get('ppt/slides/slide2.xml'));
    const c1 = xOf(vtXml, 'Step Card 1');
    // 纵向：3 步、箭头各 0.7cm → 每步 (14.45 - 1.4) / 3 = 4.35cm，整宽
    equal(c1?.x, cmToEmu(0.9), 'steps 纵向卡片满版心宽');
    equal(c1?.w, cmToEmu(32.07), 'steps 纵向卡片宽度 = 版心宽');
    equal(c1?.h, cmToEmu(4.35), 'steps 纵向每步高度');
    equal(xOf(vtXml, 'Step Card 2')?.y, cmToEmu(3.25 + 4.35 + 0.7), 'steps 纵向第 2 步的 y');
    assert(vtXml.includes('<a:prstGeom prst="line">'), 'steps 纵向箭头也是直线形状');
    assert(!bare.includes('Badge'), 'numbered:false 时不应画序号圆');
    assert(!bare.includes('Step Arrow'), 'arrows:false 时不应画箭头');
    assert(vt.inspect().warnings.length === 0, `纵向 steps 不应有告警：${vt.inspect().warnings.join(' / ')}`);
}
// compare：双栏卡片 + 中间竖分隔线 + 各自 accent
{
    equal(countOf(h3, /prst="roundRect"/g), 2, 'compare 应有两张栏卡');
    const left = xOf(h3, 'Compare Column Left');
    const right = xOf(h3, 'Compare Column Right');
    equal(left?.w, cmToEmu(15.685), 'compare 每栏宽度（gap 0.7）');
    // 栏卡高度跟着条目走（2026-09-22）：两栏各 2 条要点 → 按内容收高并垂直居中
    assert(left !== undefined && right !== undefined, 'compare 两栏卡应存在');
    assert(left.h < cmToEmu(HB.h), `compare 栏卡应按内容收高：${left.h / 360000}cm`);
    assert(left.h >= cmToEmu(2.6), `compare 栏卡不应低于下限 2.6cm：${left.h / 360000}cm`);
    equal(right.h, left.h, 'compare 两栏应等高');
    equal(right.y, left.y, 'compare 两栏应同一顶线');
    const cmpGapTop = left.y - cmToEmu(HB.y);
    const cmpGapBottom = cmToEmu(HB.y + HB.h) - (left.y + left.h);
    assert(Math.abs(cmpGapTop - cmpGapBottom) <= 2000, 'compare 整块应垂直居中');
    equal(right?.x, cmToEmu(17.285), 'compare 右栏的 x');
    assert((left?.x ?? 1) + (left?.w ?? 0) < (right?.x ?? 0), 'compare 两栏不应重叠');
    // 分隔线：竖线画在两栏之间的正中，纵向只贯穿栏卡本身
    const divider = xOf(h3, 'Compare Divider');
    assert(divider !== undefined, 'compare 默认应画竖分隔线');
    equal(divider?.x, cmToEmu(16.935), '分隔线 x = 两栏间隙正中');
    assert(Math.abs((divider?.h ?? 0) - (left.h - cmToEmu(0.2))) <= 1000,
        `分隔线应贯穿栏卡（上下各留 0.1cm）：线 ${(divider?.h ?? 0) / 360000}cm / 卡 ${left.h / 360000}cm`);
    assert(h3.includes('<a:prstGeom prst="line">'), '分隔线用直线形状');
    assert(h3.includes('<a:srgbClr val="0F766E"/>'), 'compare 的 accent 应用在栏标题与项目符号上');
    {
        // 栏标题是 accent 色而不是白字色带；项目符号用 • 且带 buClr
        const seg = h3.slice(h3.indexOf('name="Compare Column Left Title"'), h3.indexOf('name="Compare Column Left Title"') + 900);
        assert(seg.includes('<a:srgbClr val="0F766E"/>'), 'compare 栏标题应用 accent 色');
        assert(h3.includes('<a:buChar char="•"/>'), 'compare 的要点应写项目符号');
        assert(h3.includes('<a:buClr><a:srgbClr val="0F766E"/>'), 'compare 项目符号颜色应跟随 accent');
    }
    // divider:false 关闭
    const bareCmp = create({ title: '无分隔', theme: 'plain', path: 'compare-bare.pptx' }, env);
    bareCmp.compare({ title: '无分隔', divider: false, left: { title: 'a' }, right: { title: 'b' } });
    const bareXml = new TextDecoder('utf-8').decode(unzip(bareCmp.render()).get('ppt/slides/slide1.xml'));
    assert(!bareXml.includes('Compare Divider'), 'divider:false 时不应画分隔线');
    assert(bareCmp.inspect().warnings.length === 0, '开关 divider 不应产生告警');
}
// kpi：大号数字 + 单位小号 run + 标签
{
    equal(countOf(h4, /name="KPI Card \d+"/g), 3, 'kpi 应给每个指标一张卡');
    equal(countOf(h4, /name="KPI Card \d+ Value"/g), 3, 'kpi 每个指标一个大数字文本框');
    equal(countOf(h4, /name="KPI Card \d+ Label"/g), 3, 'kpi 每个指标一个标签');
    const kpiRuns = [...h4.matchAll(/<a:r><a:rPr ([^>]*)>([\s\S]*?)<\/a:rPr><a:t>([^<]*)<\/a:t>/g)]
        .map((match) => ({ attrs: match[1], props: match[2], text: match[3] }));
    const num = kpiRuns.find((run) => run.text === '128');
    const unit = kpiRuns.find((run) => run.text === '万');
    assert(num !== undefined, 'kpi 数字应写成 run');
    assert(unit !== undefined, 'kpi 单位应写成独立的 run');
    assert(Number(/sz="(\d+)"/.exec(num?.attrs ?? '')?.[1] ?? 0) >= 4000, 'kpi 大数字字号应 >= 40pt');
    assert(Number(/sz="(\d+)"/.exec(unit?.attrs ?? '')?.[1] ?? 0) < Number(/sz="(\d+)"/.exec(num?.attrs ?? '')?.[1] ?? 0),
        'kpi 单位字号必须小于数字字号');
    assert((num?.attrs ?? '').includes('b="1"'), 'kpi 大数字应加粗');
    assert((num?.props ?? '').includes('val="0F766E"'), 'kpi 的 color 应作用在数字上');
    assert(h4.includes('<a:t>营收</a:t>'), 'kpi 标签应写入');
    // columns 缺省 = min(条目数, 4)：3 个指标一行铺开
    // 卡片收高到 4.6cm 上限并垂直居中（2026-09-22：原来撑满版心，四张卡变成四个大空框）
    const kpiCard = xOf(h4, 'KPI Card 1');
    assert(kpiCard !== undefined, 'kpi 第 1 张卡应存在');
    equal(kpiCard.h, cmToEmu(4.6), 'kpi 卡片高应收到上限');
    const kpiGapTop = kpiCard.y - cmToEmu(HB.y);
    const kpiGapBottom = cmToEmu(HB.y + HB.h) - (kpiCard.y + kpiCard.h);
    assert(Math.abs(kpiGapTop - kpiGapBottom) <= 2000,
        `kpi 整块应垂直居中：上 ${kpiGapTop / 360000}cm / 下 ${kpiGapBottom / 360000}cm`);
    equal(xOf(h4, 'KPI Card 2')?.y, kpiCard.y, 'kpi 默认单行（三张卡同一行）');
    equal(xOf(h4, 'KPI Card 2')?.x, cmToEmu(0.9 + (32.07 - 1) / 3 + 0.5), 'kpi 每张卡 (32.07-2×0.5)/3 宽');
    assert(hpReport.stats.shapeKinds.ellipse >= 4, 'stats.shapeKinds 应统计到助手的椭圆（序号圆）');
}
// imageText：左右分布、ratio、题注、缺图占位
{
    const pic = xOf(h5, 'ImageText Picture');
    const panel = xOf(h5, 'ImageText Panel');
    assert(pic !== undefined && panel !== undefined, 'imageText 应有图与图侧面板');
    assert((pic?.x ?? 0) > (panel?.x ?? 0), "side:'right' 时图片应在文字面板右侧");
    // 长边 1200x675 按 contain 放进 14.75×14.45cm 的区域，等比后高度受区域限制
    equal(panel?.w, cmToEmu(32.07 - 32.07 * 0.46 - 0.7), 'imageText 文字面板宽度 = 版心 - 图宽 - 间距');
    equal(h5.includes('<a:t>示意图</a:t>'), true, 'imageText 的题注应写入');
    assert(h5.includes('<a:t>要点一</a:t>') && h5.includes('<a:t>要点二</a:t>'), 'imageText 的要点应写入');
    // side:'left' 且缺图 → 虚线占位框在左侧，并带上缺的是哪张
    const missing = xOf(h6, 'Missing Image');
    const panel2 = xOf(h6, 'ImageText Panel');
    assert(missing !== undefined, '缺图应画占位框');
    assert((missing?.x ?? 99999999) < (panel2?.x ?? 0), "side:'left' 时占位框应落在左侧");
    assert(h6.includes('图片缺失：assets/nope.png'), '缺图占位框应写明缺的是哪张');
    assert(hpReport.warnings.some((line) => line.includes('已画占位框：assets/nope.png')), '缺图应进 warnings');
    // ratio 非法 → 钳制 + warning
    const ratioDeck = create({ title: '比例', theme: 'plain', path: 'image-text-ratio.pptx' }, env);
    ratioDeck.imageText({ title: '比例', ratio: 0.95, image: { path: 'assets/nope.png' }, text: '整段文字' });
    const ratioReport = ratioDeck.save();
    assert(ratioReport.warnings.some((line) => line.includes('ratio 0.95 超出 0.2~0.8')), '应报告 ratio 越界');
    // 整段文字（没有 items）也能排
    const txtDeck = create({ title: '整段', theme: 'plain', path: 'image-text-txt.pptx' }, env);
    txtDeck.imageText({ title: '整段文字', text: '这是一整段说明文字。', image: { path: 'assets/nope.png' } });
    const txtXml = new TextDecoder('utf-8').decode(unzip(txtDeck.render()).get('ppt/slides/slide1.xml'));
    assert(txtXml.includes('<a:t>这是一整段说明文字。</a:t>'), 'imageText 的 text 应作为整段写入');
}
// timeline：轴线 + 等距节点 + 上下交错
{
    equal(countOf(h7, /name="Timeline Node \d+"/g), 4, 'timeline 应有 4 个节点');
    equal(countOf(h7, /prst="ellipse"/g), 4, 'timeline 节点用圆表示');
    const axis = xOf(h7, 'Timeline Axis');
    assert(axis !== undefined && axis.w > cmToEmu(31), 'timeline 轴线应横贯版心');
    const midY = cmToEmu(3.25 + 14.45 / 2);
    assert((axis?.y ?? 0) === midY || Math.abs((axis?.y ?? 0) - midY) <= 1, '轴线应在版心纵向正中');
    // 节点等距：第 1/4 个节点中心分别在 1/8 与 7/8 处
    const node1 = xOf(h7, 'Timeline Node 1');
    const node4 = xOf(h7, 'Timeline Node 4');
    equal((node1?.x ?? 0) + (node1?.w ?? 0) / 2, cmToEmu(0.9 + 32.07 / 8), '第 1 个节点的中心');
    equal((node4?.x ?? 0) + (node4?.w ?? 0) / 2, cmToEmu(0.9 + 32.07 * 7 / 8), '第 4 个节点的中心');
    // 上下交错：奇数节点在轴上方（文字贴轴、anchor=b），偶数在下方
    const text1 = xOf(h7, 'Timeline Text 1');
    const text2 = xOf(h7, 'Timeline Text 2');
    assert((text1?.y ?? 0) + (text1?.h ?? 0) <= midY, '第 1 个节点的文字应在轴线之上');
    assert((text2?.y ?? 0) >= midY, '第 2 个节点的文字应在轴线之下');
    assert(h7.includes('anchor="b"') && h7.includes('anchor="t"'), '交错的两侧锚点应一个是 b、一个是 t');
    assert(h7.includes('<a:t>Q1</a:t>') && h7.includes('<a:t>Q4</a:t>'), '时间点文字应写入');
    // 纵向时间线
    const vt = create({ title: '纵向时间线', theme: 'plain', path: 'timeline-v.pptx' }, env);
    vt.timeline({ title: '纵向', direction: 'vertical', items: [{ time: 'T1', title: '一' }, { time: 'T2', title: '二' }] });
    const vtXml = new TextDecoder('utf-8').decode(unzip(vt.render()).get('ppt/slides/slide1.xml'));
    const vAxis = xOf(vtXml, 'Timeline Axis');
    assert((vAxis?.h ?? 0) > cmToEmu(13), '纵向时间线的轴应纵向贯穿版心');
    assert((vAxis?.w ?? 999) <= 1, '纵向轴线的宽度应只有线宽');
    assert(vt.inspect().warnings.length === 0, `纵向 timeline 不应有告警：${vt.inspect().warnings.join(' / ')}`);
}
// iconGrid：字形文本、标签、网格几何
{
    equal(countOf(h8, /name="IconGrid \d+ Icon"/g), 6, 'iconGrid 应给 6 个图标各画一个字形');
    equal(countOf(h8, /name="IconGrid \d+ Label"/g), 6, 'iconGrid 应给 6 个格子各写一个标签');
    assert(h8.includes(`<a:t>${String.fromCodePoint(0xE72C)}</a:t>`), 'iconGrid 的 0xE72C 应转成字形');
    assert(h8.includes(`<a:t>${String.fromCodePoint(0xE8B7)}</a:t>`), 'iconGrid 应接受 U+ 写法');
    assert(h8.includes('<a:t>★</a:t>'), 'iconGrid 应接受单字符字形');
    assert(countOf(h8, /typeface="Segoe MDL2 Assets"/g) >= 2, 'iconGrid 默认字体应为 Segoe MDL2 Assets');
    // 3 列：每格 (32.07 - 1.2) / 3 = 10.29cm，图标在格内水平居中
    const icon1 = xOf(h8, 'IconGrid 1 Icon');
    equal(icon1?.w, cmToEmu(1.1), 'iconGrid 图标边长 = iconSizeCm 默认值');
    equal((icon1?.x ?? 0) + (icon1?.w ?? 0) / 2, cmToEmu(0.9 + 10.29 / 2), '图标应在格内水平居中');
    equal(xOf(h8, 'IconGrid 1 Label')?.y, cmToEmu(3.25 + 0.3 + 1.1 + 0.18), '标签应排在图标下方');
    assert(h8.includes('<a:t>搜索</a:t>') && h8.includes('<a:t>删除</a:t>'), '标签文字应写入');
    assert(h8.includes('<a:t>说明</a:t>'), 'iconGrid 的 body 应写入');
}
// read() 回读：版式名与生成端一字不差（靠形状名标记认页）
{
    const hpBack = read(hpReport.path, env);
    equal(hpBack.stats.slides, 8, 'read() 助手样例页数');
    for (const kind of ['cards', 'steps', 'compare', 'kpi', 'timeline', 'iconGrid']) {
        equal(hpBack.stats.layouts[kind], 1, `read().stats.layouts.${kind}`);
    }
    equal(hpBack.stats.layouts.imageText, 2, 'read().stats.layouts.imageText');
    equal(hpBack.stats.layouts.shape, 0, 'read() 不应把助手页认成自由绘制页');
    const want = ['cards', 'steps', 'compare', 'kpi', 'imageText', 'imageText', 'timeline', 'iconGrid'];
    want.forEach((kind, index) => {
        equal(hpBack.pages[index].layout, kind, `read() 第 ${index + 1} 页版式`);
    });
    // outline 与生成端逐行一致：版式名、标题、条目计数都不能各说各话
    equal(hpBack.outline.length, hpReport.outline.length, 'read() 与生成端 outline 行数');
    for (const [index, line] of hpBack.outline.entries()) {
        equal(line, hpReport.outline[index], `read() 与生成端 outline 第 ${index + 1} 行`);
    }
    equal(hpBack.pages[0].title, '卡片网格', 'read() 应读回助手页标题');
    assert(hpBack.pages[0].autoShapes >= 4, 'read() 应数出助手页的自绘形状');
    assert(hpBack.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
        `read() 告警里不应出现 undefined/NaN：${hpBack.warnings.join(' / ')}`);
}
// 助手页的溢出与参数风险都要进 warnings（越界、条数超出容量、方向非法、列数非法…）
{
    const overDeck = create({ title: '助手边界', theme: 'business', path: 'helpers-edge.pptx' }, env);
    overDeck.cards({
        title: '卡片过多', columns: 4,
        items: Array.from({ length: 40 }, (_, index) => ({ title: `卡 ${index + 1}`, body: '正文' })),
    });
    overDeck.cards({ title: '列数非法', columns: 7, items: [{ title: 'a' }] });
    overDeck.cards({ title: '没有卡片', items: [] });
    overDeck.cards({ title: '列数下限', columns: false, items: [{ title: 'a' }, { title: 'b' }] });
    overDeck.steps({ title: '步骤过多', items: Array.from({ length: 9 }, (_, index) => ({ title: `步骤 ${index + 1}` })) });
    overDeck.steps({ title: '方向非法', direction: 'diagonal', items: [{ title: 'a' }, { title: 'b' }] });
    overDeck.compare({
        title: '要点过多',
        left: { title: 'L', items: Array.from({ length: 14 }, (_, index) => `要点 ${index + 1}`.repeat(16)) },
        right: { title: 'R', items: ['x'] },
    });
    // 4 列 × 3 行 → 每张卡只有 4.48cm 高，大数字必须往下压
    overDeck.kpi({
        title: '数字过长',
        items: Array.from({ length: 12 }, (_, index) => (
            index === 0 ? { value: '1 2 3 4 5 6 7 8 9 0 '.repeat(4), label: '超长数字' } : { value: '1', label: `b${index}` }
        )),
    });
    overDeck.kpi({ title: '列数非法', columns: 9, items: [{ value: '1', label: 'a' }] });
    overDeck.timeline({ title: '节点过多', direction: 'vertical', items: Array.from({ length: 12 }, (_, index) => ({ time: `T${index}`, title: `标题 ${index}`, body: '正文' })) });
    overDeck.iconGrid({ title: '图标过多', items: Array.from({ length: 24 }, (_, index) => ({ icon: 0xE72C, label: `图 ${index + 1}` })) });
    overDeck.iconGrid({ title: '列数非法', columns: 9, items: [{ icon: 0xE72C, label: 'a' }] });
    overDeck.panel({ title: '面板放不下', text: '很长的一段提示词'.repeat(20), x: 1, y: 2, w: 20, h: 1.2 });
    overDeck.banner({ text: '越界横幅', x: 30, y: 18, w: 10, h: 2 });
    const overReport = overDeck.save();
    console.log(`助手边界 warnings(${overReport.warnings.length})`);
    for (const line of overReport.warnings) console.log(`  - ${line}`);
    for (const text of [
        // 卡片数量超出 列数×行数
        '40 张卡片按 4 列排成 10 行',
        '超出可用高度（最多 4×6 = 24 张）',
        '只支持 1 / 2 / 3 / 4 列，已按 2 列排版（收到 7）',
        'cards 没有任何卡片',
        // steps 太窄 / 方向非法
        '9 步横向排布每步只有 2.8cm 宽',
        'direction「diagonal」不认识',
        // compare 要点放不下
        'compare 左栏 14 条要点放不下',
        // kpi 数字过长 / 列数非法
        'kpi 第 1 个数字',
        '只支持 1 / 2 / 3 / 4 列，已按 1 列排版（收到 9）',
        // timeline 节点过多
        'timeline：12 个节点纵向排布每格只有 1.2cm 高',
        // iconGrid 过密 / 列数非法
        'iconGrid：24 个图标排成 3×8',
        '只支持 1 / 2 / 3 / 4 / 5 / 6 列，已按 3 列排版（收到 9）',
        // 元素级越界与放不下
        'panel 的文字放不下',
        'banner 超出页面',
    ]) {
        assert(overReport.warnings.some((line) => line.includes(text)), `助手边界应覆盖告警：${text}`);
    }
    // 卡片溢出时不要再逐张重复报「正文放不下」，否则真正可操作的那条会被埋掉
    equal(overReport.warnings.filter((line) => line.includes('张卡片正文放不下')).length, 0,
        '网格已溢出时不应逐张重复报正文放不下');
    assert(overReport.warnings.some((line) => line.includes('kpi 第 1 个数字') && line.includes('过长')),
        'kpi 数字压到阈值以下时应报「过长」');
    assert(overReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
        `助手边界告警里不应出现 undefined/NaN：${overReport.warnings.join(' / ')}`);
    equal(overReport.stats.slides, 12, '助手边界页数（panel/banner 续在最后一页上）');
    equal(overReport.stats.layouts.cards, 4, '助手边界 stats.layouts.cards');
    // 40 张卡的页面仍然是 cards（read() 能靠 Card N 认出来）
    const overBack = read(overReport.path, env);
    equal(overBack.pages[0].layout, 'cards', 'read() 应把 40 张卡的页面认成 cards');
    equal(overBack.pages[0].helperCount, 40, 'read() 应数出 40 张卡片');
    equal(overBack.pages[1].helperCount, 1, 'read() 应数出 1 张卡片（列数非法回落后）');
    equal(overBack.stats.layouts.shape, 1, 'read() 只应把「没有卡片」那一页退化成形状页');
}
// 一页参考稿式整页：cards + banner + panel 同页（deck.page() 之外的叠加路径）
{
    const ref2 = create({ title: '参考稿助手', theme: 'business', path: 'ref-helper.pptx', size: '16:9' }, env);
    ref2.cards({
        title: '年度进展', columns: 3, cardFill: 'F8FAFC',
        items: [{ title: '研究一' }, { title: '研究二' }, { title: '研究三' }],
    });
    ref2.banner({
        text: '2026 年度研究进展', sub: '学术汇报', x: 0.9, y: 5.6, w: 32.07, h: 1.5,
        gradient: { angle: 0, stops: [{ pos: 0, color: 'B91C1C' }, { pos: 1, color: 'F97316', alpha: 0.9 }] },
    });
    ref2.panel({ title: '提示词示例', text: '请把下面这段内容改写得更简洁。', x: 1.2, y: 8.4, w: 15 });
    const rfReport = ref2.save();
    const rfXml = new TextDecoder('utf-8').decode(unzip(ref2.render()).get('ppt/slides/slide1.xml'));
    equal(rfReport.stats.slides, 1, 'cards + banner + panel 必须落在同一页');
    equal(rfReport.stats.layouts.cards, 1, '组合页仍是 cards 页');
    equal(rfReport.stats.layouts.shape, 0, '组合页不应该多出一张自由绘制页');
    equal(rfReport.stats.shapeKinds.banner, 1, 'stats.shapeKinds.banner');
    equal(rfReport.stats.shapeKinds.panel, 1, 'stats.shapeKinds.panel');
    assert(rfReport.outline[0].startsWith('P1 cards：年度进展｜3 张卡片'), `组合页 outline：${rfReport.outline[0]}`);
    // 图层顺序：卡片在下 → 横幅 → 面板
    {
        const order = ['Card 1', 'name="Banner"', 'Banner Text', 'name="Panel"', 'Panel Text']
            .map((token) => rfXml.indexOf(token));
        assert(order.every((at) => at !== -1), `组合页元素都应存在：${order.join(',')}`);
        for (let i = 1; i < order.length; i += 1) {
            assert(order[i - 1] < order[i], `组合页图层顺序错误：${order.join(',')}`);
        }
    }
    // 横幅：渐变 + 自动对比色（红底写白字）
    {
        const grad = rfXml.slice(rfXml.indexOf('<a:gradFill'));
        assert(grad.includes('<a:gs pos="0"><a:srgbClr val="B91C1C"></a:srgbClr></a:gs>'), '横幅渐变起点色');
        assert(grad.includes('<a:gs pos="100000"><a:srgbClr val="F97316"><a:alpha val="90000"/></a:srgbClr></a:gs>'),
            '横幅渐变末色标带 alpha');
        assert(grad.includes('<a:lin ang="0" scaled="0"/>'), '横幅渐变角度 0°');
        const bannerText = rfXml.slice(rfXml.indexOf('name="Banner Text"'), rfXml.indexOf('name="Banner Text"') + 1500);
        assert(bannerText.includes('<a:srgbClr val="FFFFFF"/>'), '深色横幅上的文字应自动用白色');
        assert(rfXml.includes('<a:t>2026 年度研究进展</a:t>') && rfXml.includes('<a:t>学术汇报</a:t>'), '横幅主副标题');
    }
    // 面板：圆角 + 浅底 + 标题 + 正文；h 缺省时按文字量算出来
    {
        const panel = xOf(rfXml, 'name="Panel"') ?? xOf(rfXml, 'Panel');
        assert(panel !== undefined, 'panel 应是一块矩形');
        const seg = rfXml.slice(rfXml.indexOf('name="Panel"'), rfXml.indexOf('name="Panel"') + 1200);
        assert(seg.includes('prst="roundRect"'), 'panel 应用圆角矩形');
        assert(seg.includes('val="F2F6FB"'), 'panel 的 fill / 主题 surface 应生效');
        const adj = Number(/<a:gd name="adj" fmla="val (\d+)"\/>/.exec(seg)?.[1] ?? 0);
        equal(adj, Math.round((cmToEmu(0.18) / Math.min(panel?.w ?? 1, panel?.h ?? 1)) * 100000), 'panel 的 radius 应写成 avLst adj');
        assert(adj > 0 && adj < 50000, `panel 的 adj 应在合法区间：${adj}`);
        assert(rfXml.includes('<a:t>提示词示例</a:t>'), 'panel 标题应写入');
        assert(rfXml.includes('<a:t>请把下面这段内容改写得更简洁。</a:t>'), 'panel 文本应写入');
    }
    // h 缺省随文字量变化
    {
        const ph = create({ title: '面板高度', theme: 'plain', path: 'panel-h.pptx' }, env);
        ph.panel({ text: '一句话。' });
        ph.page();
        ph.panel({ text: '很长的一段话。'.repeat(30) });
        const parts = unzip(ph.render());
        const short = xOf(new TextDecoder('utf-8').decode(parts.get('ppt/slides/slide1.xml')), 'Panel');
        const long = xOf(new TextDecoder('utf-8').decode(parts.get('ppt/slides/slide2.xml')), 'Panel');
        assert((long?.h ?? 0) > (short?.h ?? 0) + cmToEmu(0.4),
            `panel 缺省高度应随文字量增长：${short?.h} vs ${long?.h}`);
        assert((short?.h ?? 0) > cmToEmu(0.5), 'panel 缺省高度至少装得下一行');
        equal(ph.inspect().stats.layouts.shape, 2, 'panel 不新起页，两张自由页');
    }
    // banner：纯色 + 显式文字色
    {
        const bd = create({ title: '横幅', theme: 'plain', path: 'banner.pptx' }, env);
        bd.banner({ text: '纯色横幅', x: 1, y: 1, w: 20, h: 1.5, fill: '0F766E', color: 'FFFFFF', sizePt: 22 });
        const bx = new TextDecoder('utf-8').decode(unzip(bd.render()).get('ppt/slides/slide1.xml'));
        const bseg = bx.slice(bx.indexOf('name="Banner"'), bx.indexOf('name="Banner"') + 1200);
        assert(bseg.includes('<a:srgbClr val="0F766E"/>'), 'banner 的 fill 应写成纯色填充');
        assert(bx.includes('<a:t>纯色横幅</a:t>'), 'banner 文字应写入');
        assert(bx.includes('sz="2200"'), 'banner 的 sizePt 应生效');
        assert(bx.includes('<a:srgbClr val="FFFFFF"/>'), 'banner 的 color 应显式作用在文字上');
        equal(xOf(bx, 'Banner')?.w, cmToEmu(20), 'banner 宽度按 w');
        equal(xOf(bx, 'Banner')?.h, cmToEmu(1.5), 'banner 高度按 h');
        equal(bd.inspect().stats.layouts.shape, 1, 'banner 画在当前自由页');
    }
    // 元素级助手与 deck.shape 同级：中间插 shape 不新起页
    {
        const mix = create({ title: '混排', theme: 'plain', path: 'mix.pptx' }, env);
        mix.panel({ text: '面板' });
        mix.banner({ text: '横幅' });
        mix.shape({ preset: 'rect', x: 1, y: 12, w: 4, h: 2, fill: '1D4ED8' });
        equal(mix.inspect().stats.slides, 1, 'panel/banner/shape 应画在同一页');
        equal(mix.inspect().stats.shapeKinds.panel, 1, 'stats.shapeKinds.panel');
        equal(mix.inspect().stats.shapeKinds.banner, 1, 'stats.shapeKinds.banner');
        equal(mix.inspect().stats.shapeKinds.rect, 1, 'stats.shapeKinds.rect');
        // deck.page() 之后另起一页（与既有语义一致）
        mix.page();
        mix.banner({ text: '第二页横幅' });
        equal(mix.inspect().stats.slides, 2, 'deck.page() 之后 banner 另起一页');
    }
    const rfBack = read(rfReport.path, env);
    equal(rfBack.pages[0].layout, 'cards', '组合页 read() 仍是 cards');
    assert(rfBack.outline[0].includes('3 张卡片'), `组合页 read() outline：${rfBack.outline[0]}`);
    assert(rfReport.warnings.every((line) => !line.includes('undefined') && !line.includes('NaN')),
        `组合页告警里不应出现 undefined/NaN：${rfReport.warnings.join(' / ')}`);
}

// ── 8.5 公式（原生 OMML）与三级字体 ─────────────────────────────────────────

{
    const math = create({
        title: '公式与字体', theme: 'business', path: 'math-font.pptx', author: '测试',
        // 全篇字体覆盖：正文槽 + 标题槽分开指定
        font: { body: '楷体', bodyEn: 'Georgia', title: '黑体', titleEn: 'Georgia' },
    }, env);
    // 行内公式混排在 bullets 里
    math.bullets({
        title: '质能方程',
        items: [
            { runs: [{ text: '质能方程 ' }, { math: 'E=mc^2', size: 20 }, { text: ' 是狭义相对论的推论' }] },
            { runs: [{ text: '欧拉恒等式 ' }, { math: 'e^{i\\pi}+1=0' }] },
            { text: '没有公式的普通要点' },
        ],
    });
    // 展示式公式：statement 整段只有一条 math run
    math.statement({ runs: [{ math: '\\int_{-\\infty}^{\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}' }] });
    // 表格单元格里放公式
    math.table({
        title: '公式表',
        columns: ['名称', '表达式'],
        rows: [['勾股定理', { runs: [{ math: 'a^2+b^2=c^2' }] }]],
    });
    // 解析失败的公式：回退成文字 + warning
    math.bullets({ title: '坏公式', items: [{ runs: [{ math: '\\unknowncmd{x}' }] }] });
    const mfReport = math.save();
    const mfParts = unzip(math.render());
    const mfText = new Map([...mfParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
    const mfSlide = (page) => mfText.get(`ppt/slides/slide${page}.xml`) ?? '';

    // 结构：AlternateContent + a14:m + oMath + fallback
    const s1 = mfSlide(1);
    assert(s1.includes('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'),
        '公式用 mc:AlternateContent 包装');
    assert(s1.includes('xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"'),
        'mc:Choice 声明 a14 并 Requires');
    assert((s1.match(/<a14:m>/g) ?? []).length === 2, '第 1 页两条行内公式');
    assert(s1.includes('<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'),
        '行内公式是 oMath（无 oMathPara）');
    assert(s1.includes('E = mc²'), 'fallback 里是纯文本近似');
    assert(s1.includes('sz="2000"'), '公式的 size 写进 m:r 内嵌的 a:rPr');

    // 展示式：statement 页
    const s2 = mfSlide(2);
    assert(s2.includes('<m:oMathPara'), '整段一条 math run 排成展示式（oMathPara）');
    assert(s2.includes('<m:nary>'), '积分是 nary');
    assert(s2.includes('<m:rad>'), '根号是 rad');
    // 浅色主题的陈述页公式必须用深色（onPrimary 白字在浅色主题上隐形 —— 第三十七轮修）
    assert(!s2.includes('<a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><m:t>'),
        '浅色主题 statement 的公式不能是白字');

    // 表格与坏公式
    assert(mfSlide(3).includes('<a14:m>'), '表格单元格里的公式也能写入');
    const s4 = mfSlide(4);
    assert(!s4.includes('<a14:m>'), '解析失败的公式不写 a14:m');
    assert(s4.includes('\\unknowncmd{x}'), '坏公式按原样排成文字');
    assert(mfReport.warnings.some((line) => line.includes('公式无法解析') && line.includes('第 4 页')),
        `坏公式要报带页码的 warning：${mfReport.warnings.join(' / ')}`);
    // 第四十八轮 P0-1：告警要能定位到「哪个形状、第几条 math run」，并且必须双写 env.warn ——
    // office_run 顶层的 warnings 只由 env.warnings + read() 复检警告拼成，只放进 report
    // 的返回值等于没报（批量脚本一不 return 那份 report，用户就什么都看不到）
    assert(mfReport.warnings.some((line) => line.includes('个 math run')),
        `坏公式的警告要指到形状与 run 序号：${mfReport.warnings.join(' / ')}`);
    equal(mfReport.stats.formulaErrors, 1, 'create 侧 stats.formulaErrors 数出那一条坏公式');
    equal(mfReport.stats.formulas, 4, '坏公式不算进 formulas（a14:m 只有 4 条）');
    assert(env.warnings.some((line) => line.startsWith('ppt：') && line.includes('公式无法解析')),
        `构建期警告要双写 env.warn：${env.warnings.join(' / ')}`);
    const warningsBeforeRecompile = env.warnings.length;
    math.render();
    equal(env.warnings.length, warningsBeforeRecompile, '重编译不重复报同一条构建期警告');
    // read() 侧的兜底扫描：只数 a14:m 的话，这一页会被报成「没有公式、也没有问题」
    const badRead = read('math-font.pptx', env);
    equal(badRead.stats.formulaErrors, 1, 'read() 侧扫出残留的 LaTeX 文本');
    assert(badRead.warnings.some((line) => line.includes('未解析的 LaTeX')),
        `read() 要报残留 LaTeX：${badRead.warnings.join(' / ')}`);

    // 字体：全篇覆盖要落到 theme 的 fontScheme 与页面文字
    const theme1 = mfText.get('ppt/theme/theme1.xml') ?? '';
    assert(theme1.includes('<a:majorFont><a:latin typeface="Georgia"/><a:ea typeface="黑体"/>'),
        'fontScheme majorFont 用标题槽覆盖');
    assert(theme1.includes('<a:minorFont><a:latin typeface="Georgia"/><a:ea typeface="楷体"/>'),
        'fontScheme minorFont 用正文槽覆盖');
    assert((mfSlide(1).match(/<a:latin typeface="Georgia"\/>/g) ?? []).length >= 4, '正文 run 用覆盖后的西文字体');
    assert((mfSlide(1).match(/<a:ea typeface="楷体"\/>/g) ?? []).length >= 4, '正文 run 用覆盖后的中文字体');
    assert(mfSlide(1).includes('<a:ea typeface="黑体"/>'), '页面标题用标题槽字体');

    // 统计：create 侧与 read() 侧的 formulas 同口径
    equal(mfReport.stats.formulas, 4, 'create 侧 stats.formulas（2 行内 + 1 展示式 + 1 表格）');
    const mfBack = read(mfReport.path, env);
    equal(mfBack.stats.formulas, 4, 'read() 侧 stats.formulas');
    equal(mfBack.pages[0].formulas, 2, '第 1 页 pages[0].formulas');
    equal(mfBack.pages[1].formulas, 1, '第 2 页 pages[1].formulas');
    // 字数统计把公式近似算进去（不然「带公式的要点」按空段落估高）
    assert(mfBack.pages[0].words > 0, '带公式的页 words 不为 0');

    // 段落/run 级字体：字符串与 {en,cn} 两种写法
    const ff = create({ title: 'run 字体', theme: 'plain', path: 'run-font.pptx' }, env);
    ff.shape({
        preset: 'rect', x: 1, y: 1, w: 10, h: 4, fill: 'F1F5F9',
        textOpts: {
            font: '隶书',
            paras: [{ runs: [{ text: '段落字体' }, { text: 'run 覆盖', font: { en: 'Consolas', cn: '宋体' } }] }],
        },
    });
    const ffText = new Map([...unzip(ff.render())].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
    const ff1 = ffText.get('ppt/slides/slide1.xml') ?? '';
    assert(ff1.includes('<a:ea typeface="隶书"/>'), '段落 font 字符串作用于未覆盖的 run');
    assert(ff1.includes('<a:latin typeface="Consolas"/><a:ea typeface="宋体"/>'),
        'run font 对象 {en,cn} 分别落 latin/ea');
    // 全篇 font 也可以是字符串
    const fs2 = create({ title: '字符串字体', theme: 'plain', path: 'str-font.pptx', font: '微软雅黑' }, env);
    fs2.bullets({ title: '一页', items: ['甲', '乙'] });
    const fsText = new Map([...unzip(fs2.render())].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));
    assert((fsText.get('ppt/slides/slide1.xml') ?? '').includes('<a:latin typeface="微软雅黑"/>'),
        'font 字符串连西文槽一起换');
    // 超长字体名要 warning 且被忽略
    const longFont = create({ title: '超长', theme: 'plain', path: 'long.pptx', font: 'x'.repeat(70) }, env);
    longFont.bullets({ title: '一页', items: ['甲'] });
    const longReport = longFont.save();
    assert(longReport.warnings.some((line) => line.includes('超过 64')), '超长字体名报 warning');
    assert(!(unzip(longFont.render()).get('ppt/slides/slide1.xml')?.toString('utf8') ?? '').includes('x'.repeat(70)),
        '超长字体名没有写进文件');
}

// ── 9. 枚举属性收敛（真实 PowerPoint 对非法取值零容忍）────────────────────

{
    // 踩过的坑：母版页眉写成 {align:'left'} 时，旧实现直接落 `algn="left"`，
    // 自研解析器与 LibreOffice 都照常渲染，真实 PowerPoint 却拒绝打开整个文件。
    // 这里把口语值全部过一遍收敛表，并断言不会写出任何非法枚举。
    const en = create({ title: '枚举收敛', theme: 'plain', path: 'enum.pptx' }, env);
    en.master({
        header: { text: '页眉', align: 'left', size: 15 },
        footer: { text: '页脚', align: 'center' },
    });
    en.shape({
        preset: 'roundRect', x: 2, y: 6, w: 8, h: 3,
        text: '形状文字',
        textOpts: { align: 'center', anchor: 'middle', size: 14 },
    });
    const ex = new TextDecoder('utf-8').decode(unzip(en.render()).get('ppt/slides/slide1.xml'));
    assert(ex.includes('algn="l"'), 'align:"left" 应收敛成 algn="l"');
    assert(ex.includes('algn="ctr"'), 'align:"center" 应收敛成 algn="ctr"');
    assert(ex.includes('anchor="ctr"'), 'anchor:"middle" 应收敛成 anchor="ctr"');
    assert(!/algn="(left|center|right|middle|justify)"/.test(ex),
        `不应写出非法 algn：${(ex.match(/algn="[^"]*"/g) ?? []).join(',')}`);
    assert(!/anchor="(top|middle|bottom|center)"/.test(ex),
        `不应写出非法 anchor：${(ex.match(/anchor="[^"]*"/g) ?? []).join(',')}`);
}

// ── 10. 图片：JPEG 的 Default 注册、不支持的格式、抽取 ────────────────────
//
// 主样例只覆盖了 PNG。这一组补两个「另一条分支」：jpeg 的 Default 与抽取时的扩展名映射，
// 以及 takeMedia 里「文件在、但不是 PNG/JPEG/GIF」的占位框分支（此前没有任何测试盯着它）。
{
    const jpeg = makeJpeg(8, 8);
    env.writeFile('assets/photo.jpg', jpeg);
    env.writeFile('assets/chart.bmp', Buffer.from('BM00000000 not a supported bitmap'));

    const pic = create({ title: '图片边界', theme: 'plain', path: 'media-edge.pptx' }, env);
    pic.image({ path: 'assets/photo.jpg', title: 'JPEG 照片' });
    pic.image({ path: 'assets/chart.bmp', title: '不支持的格式' });
    const picReport = pic.save();
    const picParts = unzip(pic.render());
    const picText = new Map([...picParts].map(([name, data]) => [name, new TextDecoder('utf-8').decode(data)]));

    // 生成侧口径：images 数「页面上放了几张图」（含缺图/坏格式的占位框），
    // imageParts 只数真的进了包的媒体部件 —— 两个数字分开报才看得出「哪张没进去」
    equal(picReport.stats.images, 2, 'stats.images 含占位框');
    equal(picReport.stats.imageParts, 1, 'stats.imageParts 只算真的进了包的图');
    equal(picReport.stats.mediaBytes, jpeg.length, 'stats.mediaBytes 只算真的进了包的图');
    assert(picReport.warnings.some((line) => line.includes('图片格式不支持')), '不支持的格式要给 warning');
    assert(picParts.has('ppt/media/image1.jpeg'), 'JPEG 的部件扩展名应是 .jpeg（按真实格式，不按 .jpg 后缀）');
    assert(Buffer.compare(Buffer.from(picParts.get('ppt/media/image1.jpeg')), jpeg) === 0, 'JPEG 字节应原样进包');

    const picCt = picText.get('[Content_Types].xml');
    assert(picCt.includes('<Default Extension="jpeg" ContentType="image/jpeg"/>'), 'jpeg 必须走 Default 注册');
    assert(!/Override[^>]*ppt\/media/.test(picCt), '媒体部件不该写 Override（体积大且容易漏）');

    // 不支持的格式只画占位框（p:sp），不能凭空多一个 p:pic
    const slide2 = picText.get('ppt/slides/slide2.xml');
    equal((slide2.match(/<p:pic>/g) ?? []).length, 0, '不支持的格式不该产出 p:pic');
    assert(slide2.includes('图片格式不支持'), '占位框里要写清是哪种问题');
    equal((picText.get('ppt/slides/slide1.xml').match(/<p:pic>/g) ?? []).length, 1, 'JPEG 页应有一张 p:pic');

    const picBack = read(picReport.path, env);
    equal(picBack.stats.images, 1, 'read().stats.images 只数真的 p:pic');
    equal(picBack.stats.media.length, 1, 'read().stats.media 只报真的进了包的图');
    equal(picBack.stats.media[0].part, 'ppt/media/image1.jpeg', 'read().stats.media[0].part');
    equal(picBack.stats.media[0].contentType, 'image/jpeg', 'read().stats.media[0].contentType');
    equal(picBack.stats.media[0].width, 8, 'read().stats.media[0].width');
    equal(picBack.stats.media[0].height, 8, 'read().stats.media[0].height');
    equal(picBack.stats.media[0].slides.join(','), '1', 'read().stats.media[0].slides');

    const picEx = images(picReport.path, { out: 'extract-jpeg' }, env);
    equal(picEx.count, 1, 'JPEG 抽取数量');
    equal(picEx.files[0].ext, 'jpeg', '抽取的扩展名应映射回 jpeg');
    equal(picEx.files[0].contentType, 'image/jpeg', '抽取的内容类型');
    equal(picEx.files[0].bytes, jpeg.length, '抽取的字节数');
    equal(picEx.files[0].slides.join(','), '1', '抽取报出的页码');
    assert(picEx.files[0].path.endsWith('.jpeg'), `抽出来的文件名：${picEx.files[0].path}`);
    assert(Buffer.compare(env.readFile(picEx.files[0].path), jpeg) === 0, '抽出来的 JPEG 必须逐字节一致');

    // 没有 out 时落缓存目录：这条走的是注入的 cache，不是硬编码路径
    const cached = images(picReport.path, {}, env, { cache: fakeCache });
    assert(cached.dir.startsWith('fake-cache/'), `默认应落缓存目录：${cached.dir}`);
    assert(cached.files[0].path.startsWith('fake-cache/'), `缓存里的路径：${cached.files[0].path}`);
    equal(cached.count, 1, '缓存路径同样能抽出来');
    assert(fakeCache.artifacts === 1, `写缓存要记一次 artifact：${fakeCache.artifacts}`);
}

// ── 11. 结果 ─────────────────────────────────────────────────────────────

if (failures.length > 0) {
    console.error(`\nFAIL ppt：${failures.length} 项不通过（共 ${checks} 项断言）`);
    for (const line of failures) console.error(`  ✗ ${line}`);
    if (!KEEP) rmSync(TMP, { recursive: true, force: true });
    process.exit(1);
}

console.log(`\nPASS ppt bytes=${bytes.length} parts=${parts.size} warnings=${report.warnings.length}`);
console.log(`断言 ${checks} 项全部通过；边界样例覆盖了 ${edgeReport.warnings.length} 条排版风险`);
if (KEEP) {
    console.log(`产物保留在 ${TMP}`);
} else {
    rmSync(TMP, { recursive: true, force: true });
    console.log(`临时产物已清理：${TMP}`);
}
