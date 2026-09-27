/**
 * 背景填充巡检：把 Office 产物拆开，只统计「底色 / 底纹」相关的东西。
 *
 * 用途：验收「默认不要背景」这条要求。它不判断对错，只把真实写进字节里的
 * 填充列出来，人（或模型）一眼就能看出默认产出到底有没有上底色。
 *
 * 用法：
 *   node test/inspect-fills.mjs <文件…>
 *   node test/inspect-fills.mjs "D:\...\办公模式示例\分月明细.xlsx"
 */
import { readFileSync } from 'node:fs';
import { unzipText } from '../src/engine/zip.js';

const files = process.argv.slice(2).filter((arg) => !arg.startsWith('--'));
if (files.length === 0) {
    console.log('用法：node test/inspect-fills.mjs <文件…>（.docx / .xlsx / .pptx）');
    process.exit(1);
}

/**
 * 每个格式关心的模式与「干净」时的说明。
 * PPT 的实色填充是版式的一部分（封面底色等），这里只盯空色值这种坏值。
 */
const PATTERNS = {
    docx: {
        clean: '没有任何底纹（w:shd）',
        list: [['w:shd 底纹', /<w:shd[^>]*>/g]],
    },
    xlsx: {
        clean: '没有任何纯色填充，也没有单元格应用填充',
        list: [
            ['纯色填充', /<patternFill patternType="solid">[\s\S]*?<\/patternFill>/g],
            ['单元格应用填充', /<xf [^>]*fillId="(?!0")[0-9]+"[^>]*>/g],
        ],
    },
    pptx: {
        clean: '没有空色值（实色填充不在此统计）',
        list: [['空色值', /<a:srgbClr val=""\/>/g]],
    },
};

function scan(label, text, patterns) {
    const hits = [];
    for (const [name, re] of patterns) {
        const found = text.match(re) ?? [];
        if (found.length > 0) hits.push(`    ⚠ ${label}：${name} ×${String(found.length)}`);
    }
    return hits;
}

for (const file of files) {
    const ext = file.slice(file.lastIndexOf('.')).toLowerCase();
    const kind = ext.slice(1);
    const spec = PATTERNS[kind];
    console.log(file);
    if (spec === undefined) {
        console.log('  （不是 .docx/.xlsx/.pptx，跳过）');
        continue;
    }
    const texts = unzipText(readFileSync(file));
    const interesting = [...texts.keys()].filter((name) => {
        if (kind === 'xlsx') return name === 'xl/styles.xml';
        if (kind === 'docx') return name === 'word/document.xml' || name === 'word/styles.xml';
        return /^ppt\/slides\/slide[0-9]+\.xml$/.test(name);
    });
    const hits = interesting.flatMap((name) => scan(name, texts.get(name) ?? '', spec.list));
    console.log(hits.length === 0 ? `  ✓ ${spec.clean}` : hits.join('\n'));
}
