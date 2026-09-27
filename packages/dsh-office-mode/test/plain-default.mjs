/**
 * 默认素色网格自测：不传主题时，Excel 保持普通表格的网格形态、Word 保持白纸黑字，
 * 两者都不带任何背景填充。
 *
 * 这条要求单独立一个测试盯着，因为它是「默认观感」而不是某个函数的行为：
 *   1) 默认产出里没有 solid 填充、没有 `w:shd` 底纹；
 *   2) 显式选主题时原有配色必须还在（防止一刀切把主题能力改掉）；
 *   3) read() 反查默认主题得到 plain，PPT 共享同一套默认主题也不会写出空色值。
 *
 * 跑法：node test/plain-default.mjs
 */
import { rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { createEnv } from '../src/engine/kit.js';
import { unzipText } from '../src/engine/zip.js';
import { attr, children as xmlChildren, descendants, parseXml } from '../src/engine/xml.js';
import { DEFAULT_THEME_ID } from '../src/engine/theme.js';
import { create as createWord, edit as editWord, read as readWord } from '../src/formats/word.js';
import { create as createExcel, read as readExcel } from '../src/formats/excel.js';
import { create as createPpt } from '../src/formats/ppt.js';

const ROOT = fileURLToPath(new URL('../../../.office/tmp/plain/', import.meta.url));
const failures = [];
let checks = 0;

function check(condition, label, detail = '') {
    checks += 1;
    if (!condition) failures.push(detail ? `${label} —— ${detail}` : label);
}

function checkEqual(actual, expected, label) {
    check(actual === expected, label, `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`);
}

function checkMatch(text, pattern, label) {
    check(pattern.test(text), label, `没有匹配 ${String(pattern)}`);
}

function checkNotMatch(text, pattern, label) {
    check(!pattern.test(text), label, `不该出现 ${String(pattern)}`);
}

rmSync(ROOT, { recursive: true, force: true });

// ---------------------------------------------------------------- Excel

function excelRun(theme, file) {
    const env = createEnv({ root: ROOT });
    const spec = { path: file, title: '默认网格' };
    if (theme !== undefined) spec.theme = theme;
    const wb = createExcel(spec, env);
    const sheet = wb.sheet('明细');
    sheet.title('2026 年 Q1 明细', { span: 4 });
    sheet.note('说明行同样不带背景');
    sheet.table({
        columns: [
            { title: '季度', type: 'text' },
            { title: '营收', type: 'number' },
            { title: '同比', type: 'percent' },
            { title: '负责人', type: 'text' },
        ],
        rows: [['Q1', 1280, 0.12, '张三'], ['Q2', 1530, 0.18, '李四'], ['Q3', 1410, 0.07, '王五']],
        totalRow: true,
        freeze: true,
        autofilter: true,
    });
    const report = wb.save();
    const texts = unzipText(env.readFile(file));
    return { env, report, styles: texts.get('xl/styles.xml'), sheet: texts.get('xl/worksheets/sheet1.xml') };
}

const plainXlsx = excelRun(undefined, 'plain.xlsx');
checkEqual(plainXlsx.report.theme, DEFAULT_THEME_ID, 'Excel 不传主题时用的是默认主题');
{
    const styles = parseXml(plainXlsx.styles).children[0];
    const fills = descendants(styles, 'fills')[0];
    checkEqual(xmlChildren(fills).length, 2, '默认只有 schema 要求的 none + gray125 两个 fill');
    checkNotMatch(plainXlsx.styles, /patternType="solid"/, '默认没有任何纯色填充');
    checkNotMatch(plainXlsx.styles, /fgColor/, '默认不写 fgColor');
    checkNotMatch(plainXlsx.styles, /applyFill/, '默认没有单元格应用填充');
    const xfs = xmlChildren(descendants(styles, 'cellXfs')[0]);
    check(xfs.length > 0, 'cellXfs 非空');
    for (const xf of xfs) {
        checkEqual(String(attr(xf, 'fillId', '0')), '0', '每个 cellXfs 的 fillId 都是 0');
    }
    // 网格形态还在：表格仍然写细边框（样式表里），只是没有底色
    checkMatch(plainXlsx.styles, /<left style="thin">/, '默认表格仍然写细边框');
    checkMatch(plainXlsx.sheet, /<c r="[A-Z]+\d+" s="\d+"/, '单元格挂了带边框的样式');
}
checkEqual(readExcel('plain.xlsx', plainXlsx.env).theme, DEFAULT_THEME_ID, 'Excel read() 反查默认主题');

const themedXlsx = excelRun('business', 'business.xlsx');
checkEqual(themedXlsx.report.theme, 'business', '显式 business 主题仍然生效');
checkMatch(themedXlsx.styles, /patternType="solid"/, '显式主题仍然写纯色填充');
checkMatch(themedXlsx.styles, /FF1F4E79/, '表头底色来自主题');
checkMatch(themedXlsx.styles, /FFF2F6FB/, '斑马纹底色来自主题');

// ---------------------------------------------------------------- Word

function wordRun(theme, file) {
    const env = createEnv({ root: ROOT });
    const spec = { title: '默认素色文档', path: file, author: '测试' };
    if (theme !== undefined) spec.theme = theme;
    const builder = createWord(spec, env);
    builder.title('默认素色文档', { subtitle: '不带任何背景填充' })
        .heading('一、概览', 1)
        .para('正文段落。')
        .quote('引用块只留左侧竖线。')
        .code('const x = 1;')
        .table({ columns: ['季度', '营收'], rows: [['Q1', '1,280'], ['Q2', '1,530']] });
    const report = builder.save();
    const texts = unzipText(env.readFile(file));
    return { env, report, styles: texts.get('word/styles.xml'), document: texts.get('word/document.xml') };
}

const plainDocx = wordRun(undefined, 'plain.docx');
checkEqual(plainDocx.report.theme, DEFAULT_THEME_ID, 'Word 不传主题时用的是默认主题');
checkNotMatch(plainDocx.document, /<w:shd/, '默认正文与表格都不写 w:shd 底纹');
checkNotMatch(plainDocx.styles, /w:fill="/, '默认样式里没有任何填充');
checkNotMatch(plainDocx.styles, /<w:shd/, '默认样式里没有底纹');
checkMatch(plainDocx.document, /<w:tblBorders>/, '默认表格仍然有边框（网格形态保留）');
checkMatch(plainDocx.styles, /<w:pBdr>/, '标题/引用仍靠边框区分，而不是底色');
checkEqual(readWord('plain.docx', plainDocx.env).theme, DEFAULT_THEME_ID, 'Word read() 反查默认主题');

// edit() 追加内容同样不能引入底纹：素色文档改完还是素色
editWord('plain.docx', [
    { append: { type: 'table', columns: ['项目', '值'], rows: [['甲', '1'], ['乙', '2']] } },
    { append: { type: 'para', text: '追加的正文。' } },
], plainDocx.env);
const editedDocx = unzipText(plainDocx.env.readFile('plain.docx')).get('word/document.xml');
checkNotMatch(editedDocx, /<w:shd/, '追加表格/段落后仍然没有底纹');
checkMatch(editedDocx, /追加的正文/, '追加内容真的写进去了');

const themedDocx = wordRun('business', 'business.docx');
checkEqual(themedDocx.report.theme, 'business', '显式 business 主题仍然生效');
checkMatch(themedDocx.document, /<w:shd[^>]*w:fill="1F4E79"/, '显式主题的表头底色仍在');
checkMatch(themedDocx.document, /<w:shd[^>]*w:fill="F2F6FB"/, '显式主题的斑马纹仍在');
checkMatch(themedDocx.styles, /w:fill="F2F6FB"/, '显式主题的引用/代码底色仍在');

// ---------------------------------------------------------------- PPT（共享默认主题）

const pptEnv = createEnv({ root: ROOT });
const deck = createPpt({ title: '默认素色演示', path: 'plain.pptx' }, pptEnv);
deck.cover({ title: '默认素色演示', subtitle: 'plain' })
    .bullets({ title: '要点', items: ['第一点', '第二点'] })
    .table({ title: '数据', columns: ['项', '值'], rows: [['甲', '1'], ['乙', '2']] })
    .closing({ title: '谢谢' });
const pptReport = deck.save();
checkEqual(pptReport.theme, DEFAULT_THEME_ID, 'PPT 不传主题时用的是默认主题');
{
    const texts = unzipText(pptEnv.readFile('plain.pptx'));
    let emptyColor = '';
    for (const [name, text] of texts) {
        if (!/\.(xml|rels)$/.test(name)) continue;
        if (/<a:srgbClr val=""/.test(text)) emptyColor = name;
    }
    checkEqual(emptyColor, '', 'PPT 默认产出里没有空色值（空串底色必须回退）');
    check(pptReport.warnings.every((line) => !/不存在|已回落/.test(line)), 'PPT 默认主题不应触发主题回落');
}

// ---------------------------------------------------------------- 结果

rmSync(ROOT, { recursive: true, force: true });
if (failures.length > 0) {
    for (const line of failures) console.log(`FAIL  ${line}`);
    console.log(`\nplain-default: ${checks - failures.length}/${checks} 通过`);
    process.exitCode = 1;
} else {
    console.log(`PASS plain-default: ${checks}/${checks} 项断言全部通过`);
}
