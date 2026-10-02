/**
 * ppt-math（LaTeX 子集 → OMML）单测。
 *
 * 断言三层：
 * 1. 结构：omml 片段用到的元素/属性/顺序必须是 OMML schema 认的 —— PowerPoint 对
 *    顺序错误零容忍，这里逐个元素核对关键结构；
 * 2. plain：mc:Fallback 与字数估算用的纯文本近似要「人能读」且稳定；
 * 3. 边界：不认识的命令、不配对的括号要带着字符位置报错，而不是产出半截公式。
 *
 * 跑法：node test/format-ppt-math.mjs
 */
import { isKnownCommand, knownCommands, latexToMath, looksLikeLatex, mathPlain } from '../src/formats/ppt-math.js';

const failures = [];
let checks = 0;

function assert(condition, message) {
    checks += 1;
    if (!condition) failures.push(message);
}

function equal(actual, expected, message) {
    checks += 1;
    if (actual !== expected) failures.push(`${message}：期望 ${JSON.stringify(expected)}，得到 ${JSON.stringify(actual)}`);
}

const om = (latex) => latexToMath(latex).blocks.join(' ');

// ── 1. 基础元素 ─────────────────────────────────────────────────────────────

{
    const r = latexToMath('E=mc^2');
    assert(r.ok, 'E=mc^2 应解析成功');
    equal(r.plain, 'E = mc²', 'E=mc² 的 plain 近似');
    assert(om('E=mc^2').includes('<m:sSup>'), '上标用 m:sSup');
    assert(om('E=mc^2').includes('<m:sup><m:r>'), 'm:sSup 有 sup 内容');
    assert(!om('E=mc^2').includes('m:oMathPara'), '行内公式不写 oMathPara');
}

{
    const r = latexToMath('a_i');
    equal(r.plain, 'aᵢ', '单字符下标用 Unicode 下标');
    assert(om('a_i').includes('<m:sSub>'), '下标用 m:sSub');
}

{
    // 上下标同时有：无论先写谁，OMML 都是 sSubSup(e, sub, sup)
    const a = om('x_{i}^{n}');
    const b = om('x^{n}_{i}');
    assert(a.includes('<m:sSubSup>'), 'x_{i}^{n} 用 m:sSubSup');
    equal(a, b, '上下标顺序不影响产出');
    assert(a.indexOf('<m:e>') < a.indexOf('<m:sub>') && a.indexOf('<m:sub>') < a.indexOf('<m:sup>'),
        'sSubSup 子元素顺序 e → sub → sup');
}

{
    assert(om("f' + g''").includes('′'), '撇号转 Unicode prime');
    equal(latexToMath("f'").plain, 'f′', '撇号的 plain');
}

// ── 2. 分式 / 根式 / 二项式 ─────────────────────────────────────────────────

{
    const r = latexToMath('\\frac{a+b}{c-d}');
    assert(r.ok && r.blocks[0].includes('<m:f><m:num>'), '分式用 m:f + m:num');
    assert(r.blocks[0].indexOf('<m:num>') < r.blocks[0].indexOf('<m:den>'), 'num 在 den 前');
    equal(r.plain, '(a + b)/(c - d)', '分式 plain 带括号');
    equal(latexToMath('\\frac{a}{b}').plain, 'a/b', '短分子分母不加括号');
}

{
    const r = latexToMath('\\sqrt{x^2+1}');
    assert(r.blocks[0].includes('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'), '平方根写 degHide');
    assert(r.blocks[0].indexOf('<m:deg>') < r.blocks[0].indexOf('<m:e>'), 'rad 子元素顺序 deg → e');
    equal(r.plain, '√(x² + 1)', '根式 plain');
    const n = latexToMath('\\sqrt[3]{x}');
    assert(n.blocks[0].includes('<m:deg><m:r>'), 'n 次根把次数写进 m:deg');
    assert(!n.blocks[0].includes('degHide'), '有次数时不写 degHide');
    equal(n.plain, '3√x', 'n 次根 plain');
}

{
    const r = latexToMath('\\binom{n}{k}');
    assert(r.blocks[0].includes('<m:type m:val="noBar"/>'), '二项式用 noBar 分式');
    assert(r.blocks[0].indexOf('<m:type m:val="noBar"/>') < r.blocks[0].indexOf('<m:num>'),
        'fPr 在 num 之前（CT_F 顺序）');
}

// ── 3. 大运算符与极限 ───────────────────────────────────────────────────────

{
    const r = latexToMath('\\sum_{i=1}^{n} a_i');
    const o = r.blocks[0];
    assert(o.includes('<m:chr m:val="∑"/>'), '求和字符 ∑');
    assert(o.includes('<m:limLoc m:val="undOvr"/>'), '求和默认上下限在上下（undOvr）');
    assert(o.indexOf('<m:chr') < o.indexOf('<m:limLoc'), 'naryPr 里 chr 在 limLoc 前');
    assert(o.includes('<m:sub>') && o.includes('<m:sup>') && o.includes('<m:e/>'),
        'nary 的 sub/sup/e 三件都在（隐藏时也要有空元素）');
    equal(r.plain, '∑_(i = 1)ⁿ aᵢ', '求和的 plain');
}

{
    const o = latexToMath('\\int_0^{2\\pi}').blocks[0];
    assert(o.includes('<m:chr m:val="∫"/>'), '积分字符 ∫');
    assert(o.includes('<m:limLoc m:val="subSup"/>'), '积分默认上下限在侧边（subSup）');
    assert(o.includes('<m:supHide m:val="1"/>') === false, '有上限时不写 supHide');
    // 关键回归：单 token 下标不能把后面的 ^ 吞成自己的上标
    assert(o.indexOf('<m:sub><m:r>') >= 0 && o.indexOf('<m:sub><m:sSup>') === -1,
        '\\int_0^{2\\pi} 的下标是裸 0（不是 0^2π）');
}

{
    const bare = latexToMath('\\sum');
    assert(bare.blocks[0].includes('<m:subHide m:val="1"/>') && bare.blocks[0].includes('<m:supHide m:val="1"/>'),
        '裸 \\sum 隐藏上下限');
    assert(bare.blocks[0].includes('<m:sub></m:sub>'), '隐藏的下标仍要有空元素');
    const forced = latexToMath('\\int\\limits_0^1');
    assert(forced.blocks[0].includes('<m:limLoc m:val="undOvr"/>'), '\\limits 把积分改到上下');
}

{
    const r = latexToMath('\\lim_{x \\to 0} \\frac{f(x)}{x}');
    const o = r.blocks[0];
    assert(o.includes('<m:limLow>'), '极限用 m:limLow');
    assert(o.includes('<m:func><m:fName>'), '极限名称包 m:func/fName');
    assert(o.indexOf('<m:e>') < o.indexOf('<m:lim>'), 'limLow 子元素顺序 e → lim');
    equal(r.plain, 'lim_(x → 0) (f(x))/x', '极限 plain');
}

{
    const r = latexToMath('\\log_2 8');
    assert(r.blocks[0].includes('<m:sSub>') || r.blocks[0].includes('<m:sSubSup>'), '\\log_2 的底数挂函数名');
    equal(r.plain, 'log₂8', '\\log_2 的 plain');
    assert(latexToMath('\\sin x').blocks[0].includes('<m:func>'), '\\sin x 用 m:func');
    assert(latexToMath('\\sin x').blocks[0].includes('<m:sty m:val="p"/>'), '函数名是正体');
}

// ── 4. 定界符 / 重音 / 样式 ─────────────────────────────────────────────────

{
    const o = latexToMath('\\left(\\frac{a}{b}\\right)').blocks[0];
    assert(o.includes('<m:begChr m:val="("/>') && o.includes('<m:endChr m:val=")"/>'), '\\left( \\right) 的定界符');
    assert(o.indexOf('<m:begChr') < o.indexOf('<m:endChr'), 'dPr 里 begChr 在 endChr 前');
    assert(latexToMath('\\left[x\\right]').blocks[0].includes('<m:endChr m:val="]"/>'), '方括号定界符');
    const dot = latexToMath('\\left. \\frac{dy}{dx} \\right|_{x=0}');
    assert(dot.blocks[0].includes('<m:begChr m:val=""/>'), '空定界符 . 写空 begChr');
    assert(dot.blocks[0].includes('<m:endChr m:val="|"/>'), '\\right| 写竖线');
    assert(latexToMath('\\left\\langle x \\right\\rangle').ok, '\\langle/\\rangle 定界符');
}

{
    assert(latexToMath('\\hat{y}').blocks[0].includes('<m:acc>'), '重音用 m:acc');
    assert(latexToMath('\\vec{v}').blocks[0].includes('<m:chr m:val="⃗"/>'), '向量重音字符');
    const bar = latexToMath('\\overline{X}');
    assert(bar.blocks[0].includes('<m:pos m:val="top"/>'), 'overline 用 m:bar top');
    const under = latexToMath('\\underline{X}');
    assert(under.blocks[0].includes('<m:pos m:val="bot"/>'), 'underline 用 m:bar bot');
    assert(latexToMath('\\overrightarrow{AB}').blocks[0].includes('<m:groupChr>'), 'overrightarrow 用 m:groupChr');
}

{
    const t = latexToMath('\\text{当且仅当 } x > 0');
    assert(t.blocks[0].includes('<m:sty m:val="p"/>'), '\\text 是正体');
    assert(t.blocks[0].includes('<m:t>当且仅当 </m:t>'), '\\text 保留原文与空格');
    equal(t.plain, '当且仅当 x>0', '\\text 的 plain');
    assert(latexToMath('\\mathbf{v}').blocks[0].includes('<m:sty m:val="b"/>'), '\\mathbf 粗体样式');
    assert(latexToMath('\\mathbb{R}^n').plain === 'ℝⁿ', '\\mathbb{R} 映射双线体');
    const esc = latexToMath('\\text{a \\& b \\% c}');
    assert(esc.blocks[0].includes('<m:t>a &amp; b % c</m:t>'), '\\text 里还原转义（XML 转义不丢）');
}

// ── 5. 矩阵 / cases / aligned ──────────────────────────────────────────────

{
    const r = latexToMath('\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}');
    const o = r.blocks[0];
    assert(o.includes('<m:m><m:mPr><m:mcs>'), '矩阵用 m:m + m:mPr/m:mcs');
    assert((o.match(/<m:mr>/g) ?? []).length === 2, '两行矩阵有两个 m:mr');
    assert((o.match(/<m:mc>/g) ?? []).length === 2, '两列矩阵的 mcs 有两个 m:mc');
    // 带定界符的矩阵：m:m 整体包在 m:d 里（Word 的 pmatrix 结构）
    assert(o.startsWith('<m:d><m:dPr><m:begChr m:val="("/>') && o.includes('<m:e><m:m>'),
        'pmatrix 的 m:m 包在 m:d（定界符在 dPr）');
    equal(r.plain, '(a，b；c，d)', 'pmatrix 的 plain');
    const b = latexToMath('\\begin{bmatrix} 1 \\end{bmatrix}');
    assert(b.blocks[0].includes('<m:begChr m:val="["/>') && b.blocks[0].includes('<m:endChr m:val="]"/>'),
        'bmatrix 用方括号');
    const bare = latexToMath('\\begin{matrix} 1 & 2 \\\\ 3 & 4 \\end{matrix}');
    assert(bare.blocks[0].startsWith('<m:m>'), '裸 matrix 不包 m:d');
}

{
    const r = latexToMath('\\begin{cases} x=1 & a \\\\ y=2 & b \\end{cases}');
    const o = r.blocks[0];
    assert(o.includes('<m:begChr m:val="{"/>') && o.includes('<m:endChr m:val=""/>'), 'cases 左花括号无右括号');
    assert(o.includes('<m:eqArr>'), 'cases 的行排进 eqArr');
    equal(r.plain, '{x = 1 a；y = 2 b', 'cases 的 plain');
    const aligned = latexToMath('\\begin{aligned} a&=1 \\\\ b&=2 \\end{aligned}');
    equal(aligned.plain, 'a = 1；b = 2', 'aligned 的 plain 不带对齐逗号');
}

// ── 6. 符号 / 空白 / 注释 / 多行 ───────────────────────────────────────────

{
    equal(latexToMath('\\alpha+\\beta\\leq\\Gamma').plain, 'α + β ≤ Γ', '希腊字母与关系符');
    equal(latexToMath('a \\times b \\neq c \\to d').plain, 'a × b ≠ c → d', '运算符 plain');
    equal(latexToMath('x \\, y \\quad z').plain, 'x y z', '空白命令不出乱码');
    const comment = latexToMath('a + % 这是注释\n b');
    equal(comment.plain, 'a + b', '注释行被跳过');
    const multi = latexToMath('a=b \\\\ c=d');
    assert(multi.blocks.length === 2, '顶层 \\\\ 分成两条 oMath');
    equal(multi.plain, 'a = b；c = d', '多行 plain 用分号连接');
}

// ── 7. 注入与转义 ──────────────────────────────────────────────────────────

{
    const withRpr = latexToMath('xy', { rPr: '<a:rPr sz="2400"/>' });
    assert(withRpr.blocks[0].includes('<a:rPr sz="2400"/><m:t>x'), 'rPr 注入到 m:r 里、在 m:t 之前');
    const xml = latexToMath('\\text{a<b & c>d}').blocks[0];
    assert(xml.includes('a&lt;b &amp; c&gt;d'), 'm:t 内容做 XML 转义');
    // 同一公式带不同 rPr 是不同缓存键
    const a = latexToMath('x', { rPr: '<a:rPr sz="1"/>' });
    const b = latexToMath('x', { rPr: '<a:rPr sz="2"/>' });
    assert(a.blocks[0] !== b.blocks[0], 'rPr 不同产出不同（缓存键正确）');
    const again = latexToMath('x', { rPr: '<a:rPr sz="1"/>' });
    equal(again.blocks[0], a.blocks[0], '缓存命中结果一致');
}

// ── 8. 错误路径 ─────────────────────────────────────────────────────────────

const badCases = [
    ['\\foo{x}', '不认识命令'],
    // 第四十八轮 P0-1：定界符名只在 \left…\right 里可用 —— 单独写 \langle 必须失败，
    // 而它偏偏出现在帮助文本的定界符清单里，最容易让人以为能独立用
    ['\\langle x \\rangle', '不认识命令'],
    ['\\lfloor x \\rfloor', '不认识命令'],
    ['\\frac{a', '之前结束'],
    ['{', '之前结束'],
    ['x^', '缺少参数'],
    ['&', '只能出现在'],
    ['\\left( x', '之前结束'],
    ['\\begin{matrix} a', '没有配对的 \\end'],
    ['\\end{matrix}', '没有对应的 \\begin'],
    ['\\begin{unknown} a \\end{unknown}', '不支持环境'],
    ['\\begin{matrix} a \\end{pmatrix}', '不匹配'],
];
for (const [latex, needle] of badCases) {
    const r = latexToMath(latex);
    checks += 2;
    if (r.ok !== false) failures.push(`「${latex}」应当解析失败，却成功了：${r.plain}`);
    else if (!r.error.includes(needle)) failures.push(`「${latex}」的报错应包含「${needle}」，得到：${r.error}`);
    else if (!/第 \d+ 字符附近/.test(r.error) && !r.error.includes('之前结束')) {
        failures.push(`「${latex}」的报错应带位置：${r.error}`);
    }
}

{
    equal(mathPlain('\\unknown{x}'), '\\unknown{x}', 'mathPlain 解析失败时退回原样（量算不归零）');
    equal(mathPlain(''), '', '空公式给空串');
    equal(latexToMath('   ').plain, '', '纯空白的 plain 为空');
}

// ── 8.5 支持域目录与「残留 LaTeX」判据（第四十八轮 P0-1） ──────────────────

{
    // 目录从表现取：文档里那份手写清单（ppt.js 的 meta）要与它同源，
    // 数字也会跟着表走 —— 老代码写「约 150 个」，实际 178 个，就是这么漂的。
    const catalogue = knownCommands();
    const names = catalogue.map((item) => item.name);
    assert(names.length >= 250, `支持域目录应覆盖全部表项，实际 ${names.length}`);
    assert(new Set(names).size === names.length, '目录里不该有重名');
    for (const needed of ['epsilon', 'frac', 'sqrt', 'sum', 'hat', 'pmatrix', 'text']) {
        assert(names.includes(needed), `目录应含 \\${needed}`);
    }
    // where 标出语境：\langle 只能出现在 \left…\right 里，这是「表里有名字 ≠ 能独立用」的证据
    const langle = catalogue.find((item) => item.name === 'langle');
    assert(langle?.where === 'only-in-left-right', `\\langle 的语境应为 only-in-left-right，得到 ${langle?.where}`);
    assert(catalogue.find((item) => item.name === 'limits')?.where === 'nary-modifier', '\\limits 只紧接大运算符');
    assert(catalogue.find((item) => item.name === 'epsilon')?.where === 'standalone', '\\epsilon 可独立使用');
    assert(isKnownCommand('epsilon') && !isKnownCommand('unknowncmd'), 'isKnownCommand 认得出支持域');
    assert(latexToMath('\\left\\langle x \\right\\rangle').ok, '\\left\\langle…\\right\\rangle 必须成功');
}

{
    // 残留判据：宁可漏报也不误报 —— Windows 路径与普通文字都不能命中
    const yes = ['\\langle x \\rangle + \\epsilon', '\\frac{a}{b}', '\\sqrt{2}', 'E=mc^2 \\cdot x',
        '\\unknowncmd{x}'];
    const no = ['D:\\work\\demo project\\demo1', 'C:\\Users\\alpha\\file', 'x_i^2', 'a_b', 'E = mc²',
        '\\\\server\\share', '路径 C:\\temp 里有反斜杠', 'D:\\my_files\\notes.txt'];
    for (const text of yes) assert(looksLikeLatex(text), `「${text}」应被判成残留公式`);
    for (const text of no) assert(!looksLikeLatex(text), `「${text}」不该被判成残留公式`);
}

// ── 9. 真实世界的公式 ──────────────────────────────────────────────────────

const realWorld = [
    '\\sum_{k=0}^{n} \\binom{n}{k} x^k y^{n-k}',
    '\\int_{-\\infty}^{\\infty} e^{-x^2} \\, dx = \\sqrt{\\pi}',
    '\\nabla \\times \\vec{B} = \\mu_0 \\vec{J} + \\mu_0 \\varepsilon_0 \\frac{\\partial \\vec{E}}{\\partial t}',
    'f(x) = \\begin{cases} 1 & x \\in \\mathbb{Q} \\\\ 0 & x \\notin \\mathbb{Q} \\end{cases}',
    '\\text{SNR}_{dB} = 10 \\log_{10} \\frac{P_s}{P_n}',
    'e^{i\\pi} + 1 = 0',
];
for (const latex of realWorld) {
    const r = latexToMath(latex);
    checks += 1;
    if (!r.ok) failures.push(`真实公式应可解析：${latex} → ${r.error}`);
    else if (r.plain.trim() === '') failures.push(`真实公式的 plain 不应为空：${latex}`);
}

// 结果
if (failures.length > 0) {
    console.error(`\nFAIL format-ppt-math：${failures.length} 项不通过（共 ${checks} 项断言）`);
    for (const line of failures) console.error(`  ✗ ${line}`);
    process.exit(1);
}
console.log(`\nPASS format-ppt-math：${checks} 项断言全部通过`);
