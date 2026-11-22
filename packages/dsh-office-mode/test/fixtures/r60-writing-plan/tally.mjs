/**
 * 第六十轮对照实验的测量口径（用户点名的 `新-24`：生成前先给每一部分定篇幅）。
 *
 * 这一份是**固化证据 + 可复算脚本**，不是套件：`test/measure.mjs` 会读同目录的
 * `manifest.json`，把下面的数字与两个 `.md` 现算一遍对上（谁改了产物，套件当场红）。
 * 手工复算：node test/fixtures/r60-writing-plan/tally.mjs
 *
 * 为什么放在 test/fixtures 里、而不是 `.office/tmp/`：第六十轮第一次把这批产物放在
 * `.office/tmp/r60/`，被另一个并发复核子代理当成自己的草稿目录**递归删了两次** ——
 * 交付档里那句「口径脚本留在 .office/tmp/…」当场变成死路径。所以这一轮的实验材料
 * 进版本库、随套件守门；以后复核方统一用 `.office/tmp/review-<sessionid>/`。
 *
 * 口径（引用数字时必须带全，否则会串）：
 *   - **三节正文之和**：`## ` 小节正文去掉空白后的字符数 / 汉字数（不含 H1 标题行）；
 *   - **全文**：整份文件去掉空白后的字符数 / 汉字数（含 H1 标题行）；
 *   - **占比**：每节 `三节正文之和` 的字符占比；
 *   - **极差比**：最长节 ÷ 最短节（字符口径）；
 *   - **CV**：各节占比的**标准差 ÷ 均值**（越小越「平均用力」）。
 *
 * 三条从复核里学到的纪律（都写进交付档 §四了）：
 *   1. 子代理自报的字数**不是错**，是**口径不同**：B 报的「172 / 487 / 255 = 914」是
 *      正文汉字（逐字相同）；A 报的「199 / 340 / 362 = 901」是正文汉字**加上本节标题**的
 *      汉字（+2 / +6 / +7）。引用时写清哪一种，别写成「自报与文件对不上」。
 *   2. 读数只在**子代理 settled 之后**取。第六十轮在飞读过两次（722 / 845），
 *      都是被整份替换掉的完整旧稿，不是半截文件。
 *   3. 一次读盘，别分两次读（否则并发写时会给出「跨版本的字节 + 字符」组合）。
 */
import { readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

const here = import.meta.dirname;
const strip = (text) => text.replace(/\s+/gu, '');
const hanzi = (text) => (text.match(/[\u4e00-\u9fff]/gu) ?? []).length;

/** 按 `## ` 切小节；H1（文档标题）与它上面的前言不计入各节。 */
export function sectionsOf(text) {
    const sections = [];
    let current = null;
    for (const line of text.split(/\r?\n/)) {
        const heading = /^(#{1,3})\s+(.*)$/u.exec(line);
        if (heading) {
            if (heading[1].length === 1 && /^##\s/mu.test(text)) { current = null; continue; }
            current = { title: heading[2].trim(), lines: [] };
            sections.push(current);
            continue;
        }
        if (current !== null) current.lines.push(line);
    }
    return sections.map((section) => {
        const body = section.lines.join('\n');
        return { title: section.title, chars: strip(body).length, hanzi: hanzi(body) };
    });
}

/** 一份产物的全部指标（与 manifest 里的字段一一对应）。 */
export function measure(file) {
    const buffer = readFileSync(file);              // 一次读盘：字节与字符同源
    const text = buffer.toString('utf8');
    const sections = sectionsOf(text);
    const bodyChars = sections.reduce((sum, item) => sum + item.chars, 0);
    const bodyHanzi = sections.reduce((sum, item) => sum + item.hanzi, 0);
    const shares = sections.map((item) => item.chars / bodyChars);
    const mean = shares.reduce((a, b) => a + b, 0) / shares.length;
    const sd = Math.sqrt(shares.reduce((a, b) => a + (b - mean) ** 2, 0) / shares.length);
    const counts = sections.map((item) => item.chars);
    const longest = Math.max(...counts);
    const shortest = Math.min(...counts);
    return {
        bytes: statSync(file).size,
        fileChars: strip(text).length,
        fileHanzi: hanzi(text),
        bodyChars,
        bodyHanzi,
        sections,
        ratio: shortest === 0 ? Number.NaN : longest / shortest,
        cv: sd / mean,
        longestTitle: sections[counts.indexOf(longest)]?.title,
    };
}

const same = (left, right) => JSON.stringify(left) === JSON.stringify(right);

if (import.meta.filename === process.argv[1]) {
    const manifest = JSON.parse(readFileSync(join(here, 'manifest.json'), 'utf8'));
    const rows = [];
    for (const [arm, expected] of Object.entries(manifest.arms)) {
        const file = join(here, expected.file);
        const actual = measure(file);
        console.log(`\n== 臂 ${arm}（${expected.treatment}）==`);
        console.log(`   字节 ${actual.bytes}｜三节正文之和 ${actual.bodyChars} 字符 / ${actual.bodyHanzi} 汉字｜全文 ${actual.fileChars} 字符 / ${actual.fileHanzi} 汉字（含标题行）`);
        for (const item of actual.sections) {
            console.log(`   ${String(item.chars).padStart(4)} 字符 / ${String(item.hanzi).padStart(4)} 汉字  ${((item.chars / actual.bodyChars) * 100).toFixed(1).padStart(5)}%  ${item.title}`);
        }
        console.log(`   极差比 ${actual.ratio.toFixed(2)}｜占比 CV ${actual.cv.toFixed(4)}｜最长一节：${actual.longestTitle}`);
        const expectedNumbers = {
            bytes: expected.bytes,
            fileChars: expected.fileChars,
            fileHanzi: expected.fileHanzi,
            bodyChars: expected.bodyChars,
            bodyHanzi: expected.bodyHanzi,
            sections: expected.sections,
        };
        const actualNumbers = {
            bytes: actual.bytes,
            fileChars: actual.fileChars,
            fileHanzi: actual.fileHanzi,
            bodyChars: actual.bodyChars,
            bodyHanzi: actual.bodyHanzi,
            sections: actual.sections,
        };
        console.log(same(expectedNumbers, actualNumbers)
            ? '   ✓ 与 manifest.json 逐项一致'
            : `   ✗ 与 manifest.json 不一致：\n     期望 ${JSON.stringify(expectedNumbers)}\n     实际 ${JSON.stringify(actualNumbers)}`);
        rows.push({ arm, expected, actual, ok: same(expectedNumbers, actualNumbers) });
    }
    const bad = rows.filter((row) => !row.ok);
    console.log(`\n汇总：${rows.length - bad.length}/${rows.length} 臂与 manifest 一致`);
    process.exit(bad.length === 0 ? 0 : 1);
}
