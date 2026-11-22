/**
 * kb-extract 的自测：只跑这一份，不碰全量套件。
 *
 * 所有样例包都在 os.tmpdir() 下现造现删；OOXML 用现有的 engine/zip 打包，
 * 不引第三方依赖，也不改仓库里任何既有文件。
 */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';

import { zip } from '../src/engine/zip.js';
import {
    KB_EXTRACT_MAX_BYTES,
    isExtractableKbPath,
    extractKbSource,
    extractDocxText,
    extractPptxText,
    extractXlsxText,
} from '../src/kb-extract.js';

const root = mkdtempSync(join(tmpdir(), 'dsh-kb-extract-'));

after(() => {
    rmSync(root, { recursive: true, force: true, maxRetries: 3, retryDelay: 20 });
});

function writeAt(relPath, data) {
    const absolute = join(root, relPath);
    mkdirSync(dirname(absolute), { recursive: true });
    writeFileSync(absolute, data);
    return absolute;
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PNG_BYTES = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x01, 0x02, 0x03]);
const JPG_BYTES = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46]);
const WEBP_BYTES = Buffer.from('RIFF....WEBPVP8 ', 'binary');

function packageOf(entries) {
    return zip(new Map(entries));
}

function docxPackage(bodyXml, media = {}) {
    const entries = [
        ['[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>'],
        ['word/document.xml', XML_HEAD
            + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + '<w:body>' + bodyXml + '</w:body></w:document>'],
    ];
    for (const [name, bytes] of Object.entries(media)) entries.push([name, bytes]);
    return packageOf(entries);
}

function pptxSlide(texts) {
    const paragraphs = texts
        .map((text) => '<a:p><a:r><a:t>' + text + '</a:t></a:r></a:p>')
        .join('');
    return XML_HEAD
        + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        + '<p:cSld><p:spTree><p:sp><p:txBody>' + paragraphs + '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
}

function pptxPackage(slides, media = {}) {
    const entries = [
        ['[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
    ];
    for (const [name, texts] of Object.entries(slides)) entries.push([name, pptxSlide(texts)]);
    for (const [name, bytes] of Object.entries(media)) entries.push([name, bytes]);
    return packageOf(entries);
}

const XLSX_SHARED = XML_HEAD
    + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">'
    + '<si><t>姓名</t></si>'
    + '<si><r><t>张</t></r><r><t>三</t></r></si>'
    + '<si><t xml:space="preserve"> note </t></si>'
    + '</sst>';

function sheetXml(rowsXml) {
    return XML_HEAD
        + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + '<sheetData>' + rowsXml + '</sheetData></worksheet>';
}

const SHEET_ONE = sheetXml(
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
    + '<row r="2"><c r="A2"><v>42</v></c><c r="C2" t="inlineStr"><is><t>direct</t></is></c><c r="D2"><f>A1</f><v>7</v></c></row>'
    + '<row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3" t="b"><v>1</v></c></row>',
);
const SHEET_TWO = sheetXml('<row r="1"><c r="A1" t="inlineStr"><is><t>页二</t></is></c></row>');
const SHEET_THREE = sheetXml('<row r="1"><c r="A1" t="s"><v>99</v></c></row>');

function xlsxPackage(media = {}) {
    const entries = [
        ['[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
        ['xl/sharedStrings.xml', XLSX_SHARED],
        ['xl/worksheets/sheet1.xml', SHEET_ONE],
        ['xl/worksheets/sheet2.xml', SHEET_TWO],
        ['xl/worksheets/sheet3.xml', SHEET_THREE],
    ];
    for (const [name, bytes] of Object.entries(media)) entries.push([name, bytes]);
    return packageOf(entries);
}

/* ── 路径判断与文本 ─────────────────────────────────────────────────────── */

test('isExtractableKbPath 只看扩展名', () => {
    for (const file of ['a.txt', 'b.md', 'c.Markdown', 'd.jsonl', 'e.docx', 'f.PPTX', 'g.xlsx', 'h.pdf', 'i.png', 'j.JPEG', 'k.webp', 'l.gif', 'm.js']) {
        assert.equal(isExtractableKbPath(file), true, file + ' 应可抽取');
    }
    for (const file of ['a.exe', 'b.doc', 'c', '', '   ', 'd.tar.gz', 'e.']) {
        assert.equal(isExtractableKbPath(file), false, file + ' 不应可抽取');
    }
    assert.equal(isExtractableKbPath(null), false);
    assert.equal(KB_EXTRACT_MAX_BYTES, 8 * 1024 * 1024);
});

test('txt：原样读出、title 用文件名', async () => {
    const text = '第一行\n第二行\n';
    writeAt('note.txt', text);
    const result = await extractKbSource({ root, path: 'note.txt' });
    assert.equal(result.ok, true);
    assert.equal(result.kind, 'text');
    assert.equal(result.title, 'note.txt');
    assert.equal(result.text, text);
    assert.equal(result.bytes, Buffer.byteLength(text, 'utf8'));
    assert.deepEqual(result.images, []);
    assert.deepEqual(result.warnings, []);
});

test('含 NUL 的伪文本被拒绝', async () => {
    writeAt('binary.txt', 'abc\u0000def');
    await assert.rejects(extractKbSource({ root, path: 'binary.txt' }), /NUL/);
});

test('越界路径被拒绝', async () => {
    await assert.rejects(extractKbSource({ root, path: '../escape.txt' }), /越界/);
    await assert.rejects(extractKbSource({ root, path: 'docs/../../escape.txt' }), /越界/);
});

test('空文本与超过 maxBytes 都有明确错误', async () => {
    writeAt('empty.md', '   \n\t\n');
    await assert.rejects(extractKbSource({ root, path: 'empty.md' }), /没有可抽取文本/);

    writeAt('big.txt', 'x'.repeat(64));
    await assert.rejects(extractKbSource({ root, path: 'big.txt', maxBytes: 16 }), /上限 16/);
    await assert.rejects(extractKbSource({ root, path: 'big.txt', maxBytes: 0 }), /maxBytes/);
});

test('不支持的扩展名被拒绝', async () => {
    writeAt('tool.exe', 'MZ');
    await assert.rejects(extractKbSource({ root, path: 'tool.exe' }), /不是可抽取/);
});

/* ── DOCX ──────────────────────────────────────────────────────────────── */

test('docx：段落、表格、实体解码与 word/media 图片', async () => {
    const bodyXml = '<w:p><w:r><w:t>Hello DOCX &amp; friends</w:t></w:r></w:p>'
        + '<w:p><w:r><w:t>第二段</w:t></w:r></w:p>'
        + '<w:tbl>'
        + '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>'
        + '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>'
        + '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>'
        + '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>'
        + '</w:tbl>'
        + '<w:p><w:r><w:t>&#20320;好</w:t></w:r></w:p>';
    const bytes = docxPackage(bodyXml, { 'word/media/image1.png': PNG_BYTES });
    writeAt('docs/report.docx', bytes);

    const expected = 'Hello DOCX & friends\n第二段\nA1\tB1\nA2\tB2\n你好';
    const direct = extractDocxText(bytes);
    assert.equal(direct.text, expected);
    assert.equal(direct.images.length, 1);
    assert.equal(direct.images[0].name, 'image1.png');
    assert.equal(direct.images[0].ext, 'png');
    assert.equal(direct.images[0].mime, 'image/png');
    assert.ok(Buffer.isBuffer(direct.images[0].data));
    assert.ok(direct.images[0].data.equals(PNG_BYTES));

    const result = await extractKbSource({ root, path: 'docs/report.docx' });
    assert.equal(result.kind, 'docx');
    assert.equal(result.title, 'report.docx');
    assert.equal(result.text, expected);
    assert.equal(result.images.length, 1);
    assert.equal(result.bytes, bytes.length);
    assert.deepEqual(result.warnings, []);
});

test('docx：没有任何文字时报没有可抽取文本', async () => {
    const bytes = docxPackage('<w:p><w:r><w:t></w:t></w:r></w:p>', { 'word/media/image2.png': PNG_BYTES });
    writeAt('docs/empty.docx', bytes);
    await assert.rejects(extractKbSource({ root, path: 'docs/empty.docx' }), /没有可抽取文本/);
});

test('docx：缺少 word/document.xml 说清是什么包坏了', async () => {
    const bytes = packageOf([['[Content_Types].xml', XML_HEAD + '<Types/>']]);
    writeAt('docs/broken.docx', bytes);
    await assert.rejects(extractKbSource({ root, path: 'docs/broken.docx' }), /document\.xml/);
});

/* ── PPTX ──────────────────────────────────────────────────────────────── */

test('pptx：按 slide 数字排序、页间空行与 ppt/media 图片', async () => {
    const bytes = pptxPackage({
        'ppt/slides/slide10.xml': ['ten'],
        'ppt/slides/slide2.xml': ['two-a', 'two-b'],
        'ppt/slides/slide1.xml': ['one'],
    }, {
        'ppt/media/pic.png': PNG_BYTES,
        'ppt/media/photo.jpg': JPG_BYTES,
        'ppt/media/vector.webp': WEBP_BYTES,
    });
    writeAt('docs/deck.pptx', bytes);

    const expected = 'one\n\ntwo-a\ntwo-b\n\nten';
    const direct = extractPptxText(bytes);
    assert.equal(direct.text, expected);
    assert.deepEqual(direct.images.map((image) => [image.name, image.ext, image.mime]), [
        ['pic.png', 'png', 'image/png'],
        ['photo.jpg', 'jpg', 'image/jpeg'],
        ['vector.webp', 'webp', 'image/webp'],
    ]);

    const result = await extractKbSource({ root, path: 'docs/deck.pptx' });
    assert.equal(result.kind, 'pptx');
    assert.equal(result.text, expected);
    assert.equal(result.images.length, 3);
    assert.equal(result.images[1].mime, 'image/jpeg');
    assert.deepEqual(result.warnings, []);
});

/* ── XLSX ──────────────────────────────────────────────────────────────── */

test('pptx：没有 a:p 的最小幻灯片也不丢 a:t', () => {
    const bytes = packageOf([
        ['[Content_Types].xml', XML_HEAD + '<Types/>'],
        ['ppt/slides/slide1.xml', XML_HEAD
            + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            + '<a:t>bare run</a:t></p:sld>'],
    ]);
    assert.equal(extractPptxText(bytes).text, 'bare run');
});

test('xlsx：共享字符串、inlineStr、数字、列位与 xl/media 图片', async () => {
    const bytes = xlsxPackage({ 'xl/media/logo.png': PNG_BYTES });
    writeAt('docs/book.xlsx', bytes);

    const expected = '姓名\t张三\n42\t\tdirect\t7\n note \t1\n\n页二';
    const direct = extractXlsxText(bytes);
    assert.equal(direct.text, expected);
    assert.equal(direct.images.length, 1);
    assert.equal(direct.images[0].name, 'logo.png');
    assert.equal(direct.images[0].mime, 'image/png');

    const result = await extractKbSource({ root, path: 'docs/book.xlsx' });
    assert.equal(result.kind, 'xlsx');
    assert.equal(result.text, expected);
    assert.equal(result.images.length, 1);
    assert.equal(result.warnings.length, 1);
    assert.match(result.warnings[0], /sheet3\.xml/);
    assert.match(result.warnings[0], /#99/);
});

/* ── 图片与 PDF ────────────────────────────────────────────────────────── */

test('图片文件：kind=image、占位文本与图片资产', async () => {
    writeAt('pics/photo.PNG', PNG_BYTES);
    const result = await extractKbSource({ root, path: 'pics/photo.PNG' });
    assert.equal(result.kind, 'image');
    assert.equal(result.title, 'photo.PNG');
    assert.equal(result.text, '[图片：photo.PNG]');
    assert.equal(result.images.length, 1);
    assert.equal(result.images[0].name, 'photo.PNG');
    assert.equal(result.images[0].ext, 'png');
    assert.equal(result.images[0].mime, 'image/png');
    assert.ok(result.images[0].data.equals(PNG_BYTES));
});

test('pdf：没有 env 时明确要 env', async () => {
    writeAt('docs/paper.pdf', '%PDF-1.4\n%%EOF\n');
    await assert.rejects(extractKbSource({ root, path: 'docs/paper.pdf' }), /需要 env/);
});
