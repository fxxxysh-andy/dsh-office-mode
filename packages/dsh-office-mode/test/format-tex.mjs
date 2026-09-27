/**
 * LaTeX 学位论文引擎自测（office.tex）。
 *
 * 跑法：`node test/format-tex.mjs`（默认连真实编译一起测，约 1 分钟）；
 * `node test/format-tex.mjs --no-compile` 只测生成/读取/编辑，不碰 TeX。
 * 加 `--keep` 保留 .office/tmp/tex/ 里的产物（用于人工打开 PDF 核对版式）。
 *
 * 这里刻意覆盖三件容易出错的事：
 *   1. 生成的项目要能真的被 xelatex 编出 PDF（含 bibtex 与两遍以上编译）；
 *   2. 编译**失败**时要给出真实错误，而不是假装成功；
 *   3. 复检要报出「图片不存在 / 引用对不上 / label 不存在」这类会卡住编译的问题。
 */
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { createEnv } from '../src/engine/kit.js';
import { createCache } from '../src/engine/cache.js';
import { resolveConfig } from '../src/config.js';
import { executeRun } from '../src/run.js';
import { create, edit, read, findMainFile, meta, api } from '../src/formats/tex.js';
import { texEngines, locateTemplateDir } from '../src/tex.js';
import { pdfInfo } from '../src/pdf.js';

const KEEP = process.argv.includes('--keep');
const NO_COMPILE = process.argv.includes('--no-compile');
const ROOT = fileURLToPath(new URL('../../../.office/tmp/tex/', import.meta.url));
const TEMPLATE = fileURLToPath(new URL('../../../thuthesis-v7.7.1/', import.meta.url));
/** 工作区里有没有 thuthesis 模板：没有时模板块的断言自动降级（改用 TeX Live 自带的版本）。 */
const HAS_TEMPLATE = existsSync(`${TEMPLATE}thuthesis.cls`);

const failures = [];
let checks = 0;

function check(condition, label, detail = '') {
    checks += 1;
    if (!condition) failures.push(detail ? `${label} —— ${detail}` : label);
}

function checkEqual(actual, expected, label) {
    check(actual === expected, label, `期望 ${JSON.stringify(expected)}，实际 ${JSON.stringify(actual)}`);
}

const config = resolveConfig({
    cacheDir: '.office/cache',
    texTemplateDir: TEMPLATE,
    defaultTheme: 'plain',
});
// 测试用的 env：与 office_run 里那一份同形（root + 主题 + config）。
const env = createEnv({ root: ROOT, themeResolver: () => ({ theme: {} }), config });
const cache = createCache({ root: ROOT, dir: '.office/cache' });

/** 取第一段 `\thusetup{...}` 的正文（用于断言「里面没有空行」）。 */
function firstSetupBlock(text) {
    const start = text.indexOf('\\thusetup{');
    if (start === -1) return '';
    let depth = 0;
    for (let index = start + '\\thusetup'.length; index < text.length; index += 1) {
        if (text[index] === '{') depth += 1;
        else if (text[index] === '}') {
            depth -= 1;
            if (depth === 0) return text.slice(start, index + 1);
        }
    }
    return text.slice(start);
}

console.log(`工作目录：${ROOT}`);
console.log(`模板目录：${TEMPLATE}${HAS_TEMPLATE ? '' : '（不存在，模板块断言降级）'}`);

/* ── 0. 环境 ────────────────────────────────────────────────────────────── */

{
    const engines = await texEngines();
    console.log(`TeX 引擎：${engines.available.join(' / ') || '（无）'}`);
    check(engines.available.length > 0, '本机应当探测到 TeX 发行版（测试环境是 TeX Live）', engines.hint);
    const template = locateTemplateDir(env, TEMPLATE);
    if (HAS_TEMPLATE) checkEqual(template?.complete, true, '模板目录应当被判定为「成套」');
}

/* ── 1. create：写出一整个项目 ───────────────────────────────────────────── */

const SPEC = {
    path: 'thesis/thesis.tex',
    title: '面向大规模场景的测试论文题目研究',
    titleEn: 'Research on a Test Thesis Topic for Large-scale Scenarios',
    degree: 'master',
    degreeType: 'academic',
    output: 'electronic',
    author: '张三',
    authorEn: 'Zhang San',
    supervisor: '李四, 教授',
    supervisorEn: 'Professor Li Si',
    department: '计算机科学与技术系',
    discipline: '计算机科学与技术',
    disciplineEn: 'Computer Science and Technology',
    degreeCategory: '工学硕士',
    degreeCategoryEn: 'Master of Science',
    date: '2026-06-01',
    abstract: {
        zh: ['本文研究了一个用于自测的题目。', '第二段摘要。'],
        en: ['This thesis studies a self-test topic.'],
        keywordsZh: ['测试', '排版', '学位论文'],
        keywordsEn: ['testing', 'typesetting', 'thesis'],
    },
    denotation: [['PI', '聚酰亚胺'], { term: 'MPI', note: '模型化合物' }],
    chapters: [
        {
            title: '绪论',
            sections: [
                {
                    title: '研究背景',
                    blocks: [
                        '这是第一段正文，引用一篇文献\\cite{zhang2024}。',
                        { items: ['第一点', '第二点'], ordered: false },
                        { equation: 'E = mc^2', label: 'eq:emc' },
                    ],
                },
                { title: '本文工作', blocks: ['本文的主要工作如下。'] },
            ],
        },
        {
            title: '方法',
            sections: [{
                title: '总体框架',
                blocks: [
                    { table: { caption: '参数表', header: ['参数', '取值'], rows: [['学习率', '0.01'], ['批大小', '32']] } },
                    { figure: { file: 'figures/example-image-a.pdf', caption: '框架图', width: 0.6 } },
                    { align: ['a &= b + c', 'd &= e - f'] },
                ],
            }],
        },
    ],
    references: [
        { key: 'zhang2024', type: 'article', author: '张三 and 李四', title: '一篇测试论文', journal: '测试学报', year: '2024', volume: '34', number: '2', pages: '1--7' },
        '@book{wang2020, author = {王五}, title = {测试图书}, publisher = {某出版社}, year = {2020}}',
    ],
    acknowledgements: ['感谢导师的指导。'],
    resume: { bio: ['1998 年生于某地。'], achievements: ['张三. 一篇测试论文[J]. 测试学报, 2024.'] },
    committee: { supervision: [['李四', '教授', '清华大学']], reviewers: [['赵六', '教授', '北京大学']], defense: [['钱七', '教授', '清华大学']] },
    appendix: { title: '补充材料', blocks: ['附录正文。'] },
};

{
    // 图片素材：把模板自带的示例图拷进 figures/，让编译真的有一张图可插。
    const figureSource = fileURLToPath(new URL('../../../thuthesis-v7.7.1/figures/example-image-a.pdf', import.meta.url));
    env.writeFile('thesis/figures/example-image-a.pdf', readFileSync(figureSource));

    const report = create(SPEC, env);
    checkEqual(report.ok, true, 'create 应当成功');
    checkEqual(report.format, 'tex', 'create 报告 format 应为 tex');
    checkEqual(report.main, 'thesis/thesis.tex', '主文件路径');
    if (HAS_TEMPLATE) {
        checkEqual(report.support.mode, 'copied', '默认应当把模板文件整套拷进项目');
        check(report.files.some((file) => file.path.endsWith('thuthesis.cls')), '应当拷进 thuthesis.cls');
        check(report.files.some((file) => file.path.endsWith('thu-fig-logo.pdf')), '应当拷进封面校徽 thu-fig-logo.pdf');
        check(report.files.some((file) => file.path.endsWith('thuthesis-numeric.bst')), '应当拷进参考文献样式 thuthesis-numeric.bst');
    }
    checkEqual(report.stats.chapters, 2, '章节数');
    for (const name of ['thesis/thesis.tex', 'thesis/thusetup.tex', 'thesis/data/abstract.tex', 'thesis/data/chap01.tex',
        'thesis/data/chap02.tex', 'thesis/data/denotation.tex', 'thesis/data/acknowledgements.tex',
        'thesis/data/resume.tex', 'thesis/data/committee.tex', 'thesis/data/appendix.tex', 'thesis/ref/refs.bib']) {
        check(env.exists(name), `应当写出 ${name}`);
    }

    const main = env.readText('thesis/thesis.tex');
    check(main.includes('\\documentclass[degree=master, degree-type=academic, language=chinese, fontset=windows]{thuthesis}'),
        '主文件的文档类选项要写对');
    check(main.includes('\\input{data/chap01}') && main.includes('\\input{data/chap02}'), '主文件要 \\input 两章');
    check(main.includes('\\bibliography{ref/refs}'), '主文件要有 \\bibliography');
    check(!/\n{4,}/.test(main), '主文件不该出现连续三个以上空行');
    const setup = env.readText('thesis/thusetup.tex');
    check(setup.includes('output = electronic,'), 'thusetup 应当写入 output');
    check(!/\n[ \t]*\n/.test(firstSetupBlock(setup)), '\\thusetup 块里不应有空行', firstSetupBlock(setup).slice(0, 200));
    check(setup.includes('\\bibliographystyle{thuthesis-numeric}'), '默认参考文献样式应为顺序编码制');
    const chapter = env.readText('thesis/data/chap01.tex');
    check(chapter.includes('\\chapter{绪论}'), '第 1 章标题');
    check(chapter.includes('\\begin{equation}') && chapter.includes('\\label{eq:emc}'), '公式与给定 label');
    const chapter2 = env.readText('thesis/data/chap02.tex');
    check(chapter2.includes('\\toprule') && chapter2.includes('\\bottomrule'), '表格应当用三线表');
    check(/\\includegraphics\[width=0\.6\\linewidth\]\{figures\/example-image-a\.pdf\}/.test(chapter2), '图片宽度与路径');
    check(/\\label\{fig:2-1\}/.test(chapter2), '没给 label 的图应当自动补 label（fig:2-1）');
    const bib = env.readText('thesis/ref/refs.bib');
    check(bib.includes('@article{zhang2024,') && bib.includes('@book{wang2020,'), '两条参考文献都要写进 refs.bib');
}

/* ── 2. read：项目级复检 ────────────────────────────────────────────────── */

{
    const report = read('thesis/thesis.tex', env);
    checkEqual(report.kind, 'main', '主文件应被识别为 main');
    checkEqual(report.stats.chapters, 2, 'read 应数出 2 章正文');
    checkEqual(report.stats.appendixChapters, 1, '附录单独计数，不混进正文的章数');
    check(report.stats.sections >= 3, 'read 应数出各节', String(report.stats.sections));
    checkEqual(report.stats.figures, 1, '图数');
    checkEqual(report.stats.tables, 1, '表数');
    checkEqual(report.stats.equations, 2, '公式数（1 个 equation + 1 个 align 环境）');
    checkEqual(report.stats.references, 2, 'refs.bib 条目数');
    check(report.stats.words > 50, '字数应当被数出来', String(report.stats.words));
    checkEqual(report.stats.degree, 'master', 'read 应解出学位');
    checkEqual(report.stats.fontset, 'windows', 'read 应解出字体库');
    check(report.outline.some((line) => line.includes('第 1 章 绪论')), 'outline 里应有第 1 章');
    check(report.outline.some((line) => line.includes('附录 补充材料')), 'outline 里附录应标成附录而不是第 3 章');
    check(report.outline.some((line) => line.includes('研究背景')), 'outline 里应有节标题');
    checkEqual(report.warnings.length, 0, '这个项目不应有 warnings', JSON.stringify(report.warnings));
}

/* ── 3. read：片段文件 ──────────────────────────────────────────────────── */

{
    const report = read('thesis/data/chap01.tex', env);
    checkEqual(report.kind, 'fragment', 'data/chap01.tex 是片段');
    check(report.stats.sections === 2 && report.stats.chapters === 1, '片段也要数出章与节', JSON.stringify(report.stats));
    checkEqual(report.warnings.length, 0, '片段不应报项目级 warnings');
    const report2 = read('thesis/thusetup.tex', env);
    checkEqual(report2.kind, 'fragment', 'thusetup.tex 也是片段');
}

/* ── 4. 复检真的会报问题 ────────────────────────────────────────────────── */

{
    // 复制一份项目，故意制造四类问题：图片不存在、引用不存在的文献、\ref 不存在的 label、段落过长。
    const clone = (rel) => env.writeFile(`broken/${rel}`, env.readText(`thesis/${rel}`));
    for (const rel of ['thesis.tex', 'thusetup.tex', 'data/abstract.tex', 'data/chap01.tex', 'data/chap02.tex',
        'data/denotation.tex', 'data/acknowledgements.tex', 'data/resume.tex', 'data/committee.tex',
        'data/appendix.tex', 'ref/refs.bib']) clone(rel);
    env.writeFile('broken/data/chap01.tex', env.readText('broken/data/chap01.tex')
        .replace('\\cite{zhang2024}', '\\cite{zhang2024,missing2025}')
        .replace('\\begin{equation}', '本文引用了一个不存在的标签\\ref{eq:nowhere}。\n\n\\begin{equation}')
        .replace('这是第一段正文', '很长的一段'.repeat(300)));
    env.writeFile('broken/data/chap02.tex', env.readText('broken/data/chap02.tex')
        .replace('figures/example-image-a.pdf', 'figures/not-there.pdf'));

    const report = read('broken/thesis.tex', env);
    const text = report.warnings.join('\n');
    check(/图片文件不存在：figures\/not-there\.pdf/.test(text), '应当报出图片不存在');
    check(/missing2025/.test(text), '应当报出 bib 里没有的引用键');
    check(/\\ref 指向了不存在的 label：eq:nowhere/.test(text), '应当报出 \ref 指向不存在的 label');
    check(/约 \d+ 字/.test(text), '应当报出段落过长');
    check(!report.warnings.some((line) => line.includes('zhang2024') && line.includes('没有的文献')), 'zhang2024 在 bib 里，不该被报缺');
}

/* ── 5. edit：跨文件替换 / 改字段 / 追章 / 加文献 ────────────────────────── */

{
    const report = edit('thesis/thesis.tex', [
        { find: '本文的主要工作如下。', replace: '本文的主要工作如下（已改）。' },
        { set: { title: '改过的论文标题', date: '2026-07-01' } },
        { set: { keywords: '新关键词一, 新关键词二' }, file: 'data/abstract.tex' },
        { appendTo: { file: 'data/chap02.tex', text: '这是追加的一段。' } },
        { appendChapter: { title: '实验', sections: [{ title: '设置', blocks: ['实验设置。'] }] } },
        { addReference: { key: 'li2025', type: 'inproceedings', author: 'Li M', title: 'A test paper', booktitle: 'Proc. of Test', year: '2025' } },
        { addReference: { key: 'li2025', type: 'inproceedings', title: '重复的 key 应当被跳过' } },
        { removeAll: true },
    ], env);
    checkEqual(report.changed, true, 'edit 应当有改动');
    check(report.applied.some((line) => line.includes('已改')), '跨文件替换应当命中');
    check(report.applied.some((line) => line.includes('title = {改过的论文标题}')), 'set.title 应当改掉 thusetup');
    check(report.applied.some((line) => line.includes('date = {2026-07-01}')), 'set.date 应当改掉 thusetup');
    check(report.applied.some((line) => line.includes('keywords = {新关键词一, 新关键词二}')), 'set 可以用 file 指到 abstract.tex');
    check(report.applied.some((line) => line.includes('data/chap03.tex')), 'appendChapter 应当新建 chap03');
    check(report.applied.some((line) => line.includes('新增文献 li2025')), 'addReference 应当加进 li2025');
    check(report.skipped.some((entry) => /li2025 已经存在/.test(entry.reason)), '同 key 的第二条应当被跳过');
    check(report.skipped.some((entry) => /未知操作/.test(entry.reason)), '未知操作应当进 skipped');

    const main = env.readText('thesis/thesis.tex');
    check(/\\input\{data\/chap03\}\n\\bibliography/.test(main), '新章应当插在 \\bibliography 之前，且 \\input 不带后缀', main.slice(-300));
    check(env.exists('thesis/data/chap03.tex'), 'chap03.tex 应当存在');
    const chap02 = env.readText('thesis/data/chap02.tex');
    check(chap02.includes('\\chapter{方法}') && chap02.includes('这是追加的一段。'), 'appendTo 必须真的追加，不能把原文件覆盖成只剩追加内容', chap02);
    check(env.readText('thesis/thusetup.tex').includes('改过的论文标题'), 'thusetup 应当被改写');
    check(env.readText('thesis/ref/refs.bib').includes('@inproceedings{li2025,'), 'refs.bib 应当加上 li2025');
    const again = read('thesis/thesis.tex', env);
    checkEqual(again.stats.chapters, 3, 'edit 之后应当有 3 章');
    check(again.warnings.length === 0, 'edit 之后仍不应有 warnings', JSON.stringify(again.warnings));
    checkEqual(again.stats.references, 3, 'edit 之后 refs.bib 应有 3 条');
}

/* ── 6. 本科生：跳过 committee/comments/resolution ──────────────────────── */

{
    const report = create({
        path: 'bachelor/thesis.tex',
        title: '本科综合论文训练',
        degree: 'bachelor',
        author: '李四',
        supervisor: '王五, 教授',
        abstract: { zh: ['摘要。'], keywordsZh: ['本科'] },
        chapters: [{ title: '绪论', sections: [{ title: '背景', blocks: ['正文。'] }] }],
        committee: { supervision: [['王五', '教授', '清华大学']] },
        acknowledgements: ['致谢。'],
    }, env);
    check(report.warnings.some((line) => line.includes('本科生不需要 committee')), '应当提示本科生不需要 committee');
    check(!env.exists('bachelor/data/committee.tex'), '本科生项目不该写 committee.tex');
    const main = env.readText('bachelor/thesis.tex');
    check(main.includes('\\documentclass[degree=bachelor, language=chinese, fontset=windows]{thuthesis}'), '本科生不要 degree-type', main.split('\n')[5]);
    check(!main.includes('degree-type'), '本科生选项里不应出现 degree-type');
}

/* ── 7. 编译：真的编出 PDF ──────────────────────────────────────────────── */

if (NO_COMPILE) {
    console.log('（--no-compile：跳过真实编译）');
} else {
    const started = Date.now();
    const engines = await texEngines();
    if (engines.available.length === 0) {
        check(false, '编译测试需要本机有 TeX 发行版');
    } else {
        const result = await api.compile(env)('thesis/thesis.tex', { clean: 'none' });
        console.log(`编译耗时 ${Math.round((Date.now() - started) / 1000)} 秒：ok=${result.ok} engine=${result.engine} exit=${result.exitCode}`);
        checkEqual(result.ok, true, '编译应当成功', JSON.stringify(result.log.errors.slice(0, 3)));
        check(result.pdf !== null, '应当产出 PDF');
        if (result.pdf !== null) {
            check(result.pdf.bytes > 50_000, 'PDF 体积应当像一份真论文', String(result.pdf.bytes));
            check(result.pdf.pages > 5, 'PDF 页数应当大于 5', String(result.pdf.pages));
            check(env.exists('thesis/thesis.pdf'), 'PDF 应当落在项目目录里');
            const info = await pdfInfo('thesis/thesis.pdf', env, { cache });
            checkEqual(info.hasTextLayer, true, '生成的 PDF 应当有文本层');
            check(info.pages >= 5, '用 pdf 解析器复核页数', String(info.pages));
        }
        checkEqual(result.log.errors.length, 0, '编译日志里不应有错误', JSON.stringify(result.log.errors.slice(0, 3)));

        // 中间产物默认清掉（clean: 'aux'），但这次显式 clean:'none'，.log 应当在。
        check(env.exists('thesis/thesis.log'), 'clean:none 时应保留 .log 供排查');

        // 编译失败也要给得出真话：把主文件改坏再编一次。
        env.writeFile('thesis/broken.tex', env.readText('thesis/thesis.tex').replace('\\begin{document}', '\\begin{document}\n\\undefinedcommandhere'));
        const broken = await api.compile(env)('thesis/broken.tex', { clean: 'none' });
        checkEqual(broken.ok, false, '坏文件不应当被报成编译成功');
        check(broken.log.errors.some((item) => /undefined command|Undefined control sequence/i.test(item.message)), '错误信息里应当有 Undefined control sequence', JSON.stringify(broken.log.errors.slice(0, 3)));
        check(typeof broken.hint === 'string' && broken.hint.length > 0, '失败报告要给 hint');
        env.remove('thesis/broken.tex');
    }
}

/* ── 8. meta 与 office_help 的一致性 ────────────────────────────────────── */

{
    checkEqual(meta.id, 'tex', 'meta.id');
    checkEqual(meta.ext, '.tex', 'meta.ext');
    check(typeof meta.guide === 'string' && meta.guide.includes('office.tex.create'), 'meta.guide 要能当 office_help 的 tex 话题');
    check(meta.guide.includes('office.tex.compile'), 'meta.guide 要讲到编译');
    check(!meta.guide.includes('undefined'), 'meta.guide 里不该出现 undefined');
    for (const name of ['compile', 'engines', 'template', 'escape', 'clearAux']) {
        check(typeof api[name] === 'function', `api.${name} 应当存在`);
    }
    checkEqual(typeof api.compile(env), 'function', 'api.compile 是「拿 env 返回函数」的工厂');
    const located = findMainFile('thesis/data/chap01.tex', env);
    checkEqual(located?.rel, 'thesis/thesis.tex', '片段文件应当能靠 % !TEX root 找回主文件');
}

/* ── 9. 端到端：走 office_run 的真实路径（vm 沙箱 → SDK → 复检） ────────── */

if (!NO_COMPILE && HAS_TEMPLATE) {
    const E2E = fileURLToPath(new URL('../../../.office/tmp/tex/e2e/', import.meta.url));
    const exec = { agent: { session: { header: { cwd: E2E } } } };
    const result = await executeRun({
        purpose: 'LaTeX 论文端到端',
        script: `
const t = office.tex.create({
  path: 'thesis/thesis.tex', degree: 'master',
  title: '端到端测试论文', author: '张三', supervisor: '李四, 教授',
  department: '计算机科学与技术系', discipline: '计算机科学与技术',
  degreeCategory: '工学硕士',
  abstract: { zh: ['这是一段用于端到端自测的摘要。'], keywordsZh: ['自测', '排版'] },
  chapters: [
    { title: '绪论', sections: [{ title: '背景', blocks: ['正文第一段\\\\cite{zhang2024}。', { equation: 'E = mc^2' }] }] },
    { title: '方法', sections: [{ title: '框架', blocks: [{ table: { caption: '参数', header: ['参数', '值'], rows: [['学习率', '0.01']] } }] }] },
  ],
  references: [{ key: 'zhang2024', type: 'article', author: '张三', title: '题名', journal: '学报', year: '2024' }],
  acknowledgements: ['感谢导师。'],
});
const r = await office.tex.compile(t.main, { clean: 'none' });
const engines = await office.tex.engines();
return {
  main: t.main, files: t.files.length, support: t.support.mode,
  ok: r.ok, pages: r.pdf ? r.pdf.pages : 0, errors: r.log.errors.length,
  engine: r.engine, engines: engines.available,
  escaped: office.tex.escape('100% & a_1'),
};
`,
    }, exec, resolveConfig({ cacheDir: '.office/cache', texTemplateDir: TEMPLATE, defaultTheme: 'plain' }));

    checkEqual(result.ok, true, 'office_run 应当成功', JSON.stringify(result.error));
    checkEqual(result.returned?.ok, true, '脚本里的 compile 应当成功', JSON.stringify(result.returned?.errors));
    check((result.returned?.pages ?? 0) > 5, '端到端产出的 PDF 应当有页数', String(result.returned?.pages));
    checkEqual(result.returned?.escaped, '100\\% \\& a\\_1', 'office.tex.escape 应当挂在 SDK 上');
    checkEqual(result.returned?.support, 'copied', '端到端应当走模板拷贝');
    check((result.returned?.engines ?? []).includes('latexmk'), 'SDK 上的 office.tex.engines 应当探测到 latexmk');

    // 写出的 .tex 必须与三件套一样被自动复检（注册表驱动，而不是写死三件套）。
    const texRecords = result.files.filter((file) => file.format === 'tex');
    check(texRecords.length >= 3, '写出的 .tex 应当逐个被复检', JSON.stringify(result.files.map((file) => `${file.path}:${file.format}`)));
    check(texRecords.some((file) => file.path.endsWith('thesis.tex') && file.stats?.chapters === 2), '主文件复检应报出 2 章', JSON.stringify(texRecords.map((file) => [file.path, file.stats?.chapters])));
    check(result.otherFiles.some((file) => file.path.endsWith('thuthesis.cls')), '模板文件应当出现在 otherFiles 里');
    check(!result.warnings.some((line) => line.includes('部分格式不可用')), '不应该有格式加载失败', JSON.stringify(result.warnings));
    check(existsSync(join(E2E, 'thesis', 'thesis.pdf')), 'PDF 应当落在 e2e 目录里');
    if (!KEEP) env.remove('e2e');
}

/* ── 结果 ───────────────────────────────────────────────────────────────── */

if (!KEEP) {
    // .office/tmp 与其它引擎共用，只能删自己产出的目录。
    env.remove('thesis');
    env.remove('broken');
    env.remove('bachelor');
    env.remove('.office/cache');
    console.log('已清理本次产物（保留 .office/tmp 下其它文件）');
} else {
    console.log(`保留了产物：${ROOT}`);
}

if (failures.length) {
    console.error(`\nFAIL tex —— ${failures.length}/${checks} 项断言未通过：`);
    for (const item of failures) console.error(`  ✗ ${item}`);
    process.exitCode = 1;
} else {
    console.log(`\nPASS tex checks=${checks}`);
}
