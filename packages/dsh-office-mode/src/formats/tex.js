/**
 * LaTeX 学位论文（thuthesis）格式引擎。
 *
 * 与 word / excel / ppt 的区别：那三个格式的产物是「一个文件」，而一篇学位论文
 * 是「一个项目」—— 主文件、thusetup.tex、data/ 下按章分文件、ref/refs.bib，
 * 外加模板的运行期文件（thuthesis.cls、两个校徽、参考文献样式）。所以这里：
 *
 *   - create() 一次写出一整个能直接编译的项目，而不是一个 .tex；
 *   - read()  按项目解析（顺着 \input 递归收集）：章节结构、图表公式计数、
 *             引用与参考文献对不对得上、有没有指向不存在的文件；
 *   - edit()  按项目编辑：跨文件字面替换、改 \thusetup 字段、追加章节、加文献条目；
 *   - api.compile() 调本机 TeX 发行版编出 PDF（见 src/tex.js），api.engines()
 *             报告本机有什么，api.escape() 给不写 LaTeX 的普通文本转义。
 *
 * 排版这件事不归插件管：thuthesis.cls 已经把清华学位论文的版式写死了（字体、
 * 行距、页边距、题注、参考文献格式）。插件保证「结构、字段、文件、引用」都对，
 * 让模板自己去排版 —— 这也是「符合排版」的正确落点：手工拼版式过不了格式审查。
 *
 * @module dsh-office-mode/formats/tex
 */
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { asArray, asString, displayPath, ensureExtension } from '../engine/kit.js';
import { cleanBuild, compileTex, copySupport, escapeLatex, locateTemplateDir, texEngines } from '../tex.js';

/* ── 常量 ───────────────────────────────────────────────────────────────── */

/** 模板支持的学位与选项（照 thuthesis 文档；写错会让编译在第一行就停住）。 */
const DEGREES = ['doctor', 'master', 'bachelor', 'postdoc'];
const DEGREE_TYPES = ['academic', 'professional'];
const LANGUAGES = ['chinese', 'english'];
const FONTSETS = ['windows', 'mac', 'fandol', 'ubuntu'];
const BIB_STYLES = {
    numeric: { package: '\\usepackage[sort]{natbib}', style: 'thuthesis-numeric' },
    'author-year': { package: '\\usepackage{natbib}', style: 'thuthesis-author-year' },
    bachelor: { package: '\\usepackage[sort]{natbib}', style: 'thuthesis-bachelor' },
};

/** 不需要「指导小组名单 / 评语 / 决议书」三种表页的学位。 */
const SKIP_COMMITTEE = new Set(['bachelor']);

/* ── 小工具 ─────────────────────────────────────────────────────────────── */

function asObject(value, fallback = {}) {
    return value !== null && typeof value === 'object' && !Array.isArray(value) ? value : fallback;
}

function escapeRegExp(text) {
    return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** 去掉行内注释（`\%` 不是注释起点）。 */
function stripComments(source) {
    return String(source ?? '')
        .split(/\r?\n/)
        .map((line) => {
            for (let index = 0; index < line.length; index += 1) {
                if (line[index] !== '%') continue;
                let backslashes = 0;
                for (let back = index - 1; back >= 0 && line[back] === '\\'; back -= 1) backslashes += 1;
                if (backslashes % 2 === 0) return line.slice(0, index);
            }
            return line;
        })
        .join('\n');
}

/** 中文按字、西文按词估字数；命令、注释、环境名都不算。 */
function countWords(source) {
    const clean = stripComments(String(source ?? ''))
        .replace(/\\(?:begin|end)\s*\{[^}]*\}/g, ' ')
        .replace(/\\(?:cite|citep|citet|upcite|parencite|textcite|ref|eqref|autoref|pageref|label|includegraphics|input|include|bibliography)\s*(?:\[[^\]]*\])*\s*\{[^}]*\}/g, ' ')
        .replace(/\\[a-zA-Z@]+\*?(?:\[[^\]]*\])?/g, ' ')
        .replace(/[{}$&#_^~\\]/g, ' ');
    const cjk = (clean.match(/[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/g) ?? []).length;
    const western = (clean.match(/[A-Za-z][A-Za-z'’-]*/g) ?? []).length;
    return { cjk, western, words: cjk + western };
}

/** 标题里的 LaTeX 命令去掉，只留人看的字（outline 用）。 */
function plainTitle(text) {
    return stripComments(String(text ?? ''))
        .replace(/\\(?:textbf|textit|emph|LaTeX|TeX|text|texttt)\s*\{([^{}]*)\}/g, '$1')
        .replace(/\\[a-zA-Z@]+\*?/g, '')
        .replace(/[{}]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * 找一个命令的花括号参数区间（跳过 `\{`，允许一层嵌套）。
 * setSetupFields 靠它在 \thusetup{...} 的末尾补字段。
 */
function findBracedArg(text, command, from = 0) {
    const start = text.indexOf(`${command}{`, from);
    if (start === -1) return undefined;
    const open = start + command.length;
    let depth = 0;
    for (let index = open; index < text.length; index += 1) {
        const ch = text[index];
        if (ch === '\\') {
            index += 1;
            continue;
        }
        if (ch === '{') depth += 1;
        else if (ch === '}') {
            depth -= 1;
            if (depth === 0) return { start: open, end: index, body: text.slice(open + 1, index) };
        }
    }
    return undefined;
}

/* ── 内容块渲染 ─────────────────────────────────────────────────────────── */

/**
 * 把内容块渲染成 LaTeX。块可以是字符串（一段正文）或对象：
 *   {p} {items, ordered} {equation, label} {align} {figure} {table} {code} {raw}
 * 文本一律按 **LaTeX 源码** 处理：字面的 % & _ # $ 要自己转义，
 * 手里是普通文本时先调 office.tex.escape()。
 */
function renderBlocks(blocks, ctx) {
    const lines = [];
    for (const raw of asArray(blocks)) {
        if (typeof raw === 'string') {
            if (raw.trim() !== '') lines.push(raw.trim(), '');
            continue;
        }
        const block = asObject(raw);
        if (block.p !== undefined) {
            const text = asString(block.p).trim();
            if (text !== '') lines.push(text, '');
            continue;
        }
        if (block.items !== undefined) {
            const env = block.ordered === true ? 'enumerate' : 'itemize';
            lines.push(`\\begin{${env}}`);
            for (const item of asArray(block.items)) {
                const text = typeof item === 'string' ? item : asString(asObject(item).text ?? asObject(item).p);
                lines.push(`  \\item ${text}`);
            }
            lines.push(`\\end{${env}}`, '');
            continue;
        }
        if (block.equation !== undefined) {
            const label = asString(block.label) || ctx.autoLabel('eq');
            lines.push('\\begin{equation}');
            lines.push(`  ${asString(block.equation).trim()}`);
            lines.push(`  \\label{${label}}`);
            lines.push('\\end{equation}', '');
            continue;
        }
        if (block.align !== undefined) {
            lines.push('\\begin{align}');
            for (const row of asArray(block.align)) {
                const text = asString(row).trim();
                lines.push(`  ${/\\\\$/.test(text) ? text : `${text} \\\\`}`);
            }
            lines.push('\\end{align}', '');
            continue;
        }
        if (block.figure !== undefined) {
            const figure = typeof block.figure === 'string' ? { file: block.figure } : asObject(block.figure);
            const file = asString(figure.file ?? figure.path);
            const caption = asString(figure.caption);
            const label = asString(figure.label) || ctx.autoLabel('fig');
            const width = figure.widthCm !== undefined
                ? `width=${Number(figure.widthCm)}cm`
                : `width=${figure.width !== undefined ? Number(figure.width) : 0.7}\\linewidth`;
            lines.push('\\begin{figure}[htbp]');
            lines.push('  \\centering');
            lines.push(`  \\includegraphics[${width}]{${file}}`);
            if (caption !== '') lines.push(`  \\caption{${caption}}`);
            lines.push(`  \\label{${label}}`);
            lines.push('\\end{figure}', '');
            ctx.noteFigure(file);
            continue;
        }
        if (block.table !== undefined) {
            lines.push(...renderTable(asObject(block.table), ctx));
            continue;
        }
        if (block.code !== undefined) {
            lines.push('\\begin{verbatim}');
            lines.push(String(block.code).replace(/\s+$/, ''));
            lines.push('\\end{verbatim}', '');
            continue;
        }
        if (block.raw !== undefined) {
            lines.push(String(block.raw).replace(/\s+$/, ''), '');
            continue;
        }
        ctx.warnings.push(`未知的内容块，已跳过：${Object.keys(block).join(',') || '(空)'}`);
    }
    return lines;
}

/** 三线表（thuthesis 自带 booktabs，\toprule/\midrule/\bottomrule 就是它的版式）。 */
function renderTable(table, ctx) {
    const header = asArray(table.header ?? table.columns)
        .map((cell) => (typeof cell === 'string' ? cell : asString(asObject(cell).title ?? asObject(cell).text)));
    const rows = asArray(table.rows)
        .map((row) => asArray(row).map((cell) => (typeof cell === 'string' || typeof cell === 'number' ? String(cell) : asString(asObject(cell).text))));
    const width = Math.max(header.length, ...rows.map((row) => row.length), 1);
    const align = Array.isArray(table.align)
        ? table.align.join('')
        : (typeof table.align === 'string' && table.align.trim() !== ''
            ? table.align.trim()
            : ['l', ...Array.from({ length: width - 1 }, () => 'c')].join(''));
    const caption = asString(table.caption);
    const label = asString(table.label) || ctx.autoLabel('tab');
    const pad = (cells) => {
        const out = cells.slice(0, width);
        while (out.length < width) out.push('');
        return out;
    };
    const lines = ['\\begin{table}[htbp]', '  \\centering'];
    if (caption !== '') lines.push(`  \\caption{${caption}}`);
    lines.push(`  \\label{${label}}`, `  \\begin{tabular}{${align}}`, '    \\toprule');
    if (header.length > 0) {
        lines.push(`    ${pad(header).join(' & ')} \\\\`, '    \\midrule');
    }
    for (const row of rows) lines.push(`    ${pad(row).join(' & ')} \\\\`);
    lines.push('    \\bottomrule', '  \\end{tabular}', '\\end{table}', '');
    return lines;
}

/** 造一个内容块的渲染上下文（自动 label 的编号与「图片文件在不在」的记账都靠它）。 */
function createRenderContext(warnings, chapterIndex = 1) {
    const counters = { fig: 0, tab: 0, eq: 0 };
    const figures = [];
    const ctx = {
        chapterIndex,
        warnings,
        figures,
        autoLabel(kind) {
            counters[kind] += 1;
            return `${kind}:${ctx.chapterIndex}-${counters[kind]}`;
        },
        noteFigure(file) {
            figures.push({ file, chapter: ctx.chapterIndex });
        },
    };
    return ctx;
}

/* ── 生成：各文件模板 ───────────────────────────────────────────────────── */

function mainFile(path, layout) {
    const include = layout.include;
    const lines = [
        '% !TEX encoding = UTF-8',
        '% !TEX program = xelatex',
        '%',
        '% 清华大学学位论文主文件（thuthesis），由 dsh-office-mode 的 office.tex.create() 生成。',
        `% 编译：office.tex.compile('${path}')；正文按章分文件放在 data/ 下。`,
        '',
        `\\documentclass[${layout.classOptions}]{thuthesis}`,
        '',
        '\\input{thusetup}',
        '',
        '\\begin{document}',
        '',
        '% 封面',
        '\\maketitle',
    ];
    if (include.committee) lines.push('', '% 学位论文指导小组、公开评阅人和答辩委员会名单', '\\input{data/committee}');
    lines.push('', '% 使用授权的说明', '\\copyrightpage', '', '\\frontmatter');
    if (include.abstract) lines.push('% 中英文摘要', '\\input{data/abstract}');
    lines.push('', '% 目录', '\\tableofcontents', '', '% 插图和附表清单', '\\listoffigures', '\\listoftables');
    if (include.denotation) lines.push('', '% 符号对照表', '\\input{data/denotation}');
    lines.push('', '', '% 正文', '\\mainmatter');
    for (const chapter of layout.chapters) lines.push(`\\input{data/${chapter.file.replace(/\.tex$/, '')}}`);
    lines.push('', '% 参考文献（BibTeX + natbib）', '\\bibliography{ref/refs}');
    if (include.appendix) lines.push('', '% 附录', '\\appendix', '\\input{data/appendix}');
    lines.push('', '% 其它部分', '\\backmatter');
    if (include.acknowledgements) lines.push('', '% 致谢', '\\input{data/acknowledgements}');
    lines.push('', '% 声明', '\\statement');
    if (include.resume) lines.push('', '% 个人简历、在学期间完成的相关学术成果', '\\input{data/resume}');
    if (include.comments) lines.push('', '% 指导教师评语', '\\input{data/comments}');
    if (include.resolution) lines.push('', '% 答辩委员会决议书', '\\input{data/resolution}');
    lines.push('', '\\end{document}', '');
    return lines.join('\n');
}

function setupFile(spec) {
    const fields = [];
    const put = (key, value) => {
        if (value === undefined || value === null || value === '') return;
        if (typeof value === 'boolean' || typeof value === 'number') fields.push(`  ${key} = ${String(value)},`);
        else fields.push(`  ${key} = {${String(value)}},`);
    };
    fields.push('  % 输出格式：print（打印版，插空白页便于双面打印）| electronic（电子版）');
    fields.push(`  output = ${spec.output},`);
    put('title', spec.title);
    put('title*', spec.titleEn);
    put('degree-category', spec.degreeCategory);
    put('degree-category*', spec.degreeCategoryEn);
    put('department', spec.department);
    put('discipline', spec.discipline);
    put('discipline*', spec.disciplineEn);
    if (spec.degree === 'postdoc') {
        put('clc', spec.clc);
        put('udc', spec.udc);
        put('id', spec.id);
        put('discipline-level-1', spec.discipline);
        put('discipline-level-2', spec.disciplineLevel2);
        put('start-date', spec.startDate);
    } else {
        put('student-id', spec.studentId);
    }
    put('author', spec.author);
    put('author*', spec.authorEn);
    put('supervisor', spec.supervisor);
    put('supervisor*', spec.supervisorEn);
    put('associate-supervisor', spec.associateSupervisor);
    put('associate-supervisor*', spec.associateSupervisorEn);
    put('co-supervisor', spec.coSupervisor);
    put('professional-field', spec.professionalField);
    put('professional-field*', spec.professionalFieldEn);
    put('secret-level', spec.secretLevel);
    put('secret-year', spec.secretYear);
    put('date', spec.date);
    put('include-spine', spec.includeSpine === true);
    for (const [key, value] of Object.entries(asObject(spec.setup))) put(key, value);

    const bib = BIB_STYLES[spec.bibStyle];
    return [
        `% !TEX root = ./${spec.fileName}`,
        '',
        '% 论文基本配置：字段含义与可用选项见 thuthesis 文档（thuthesis.pdf）。',
        '% 注意：\\thusetup{} 里面不要出现空行；不需要的字段整行删掉即可。',
        '',
        '\\thusetup{',
        ...fields,
        '}',
        '',
        '% 载入所需的宏包',
        '\\usepackage{amsthm}',
        '\\thusetup{',
        '  math-font = xits,   % stix | xits | libertinus',
        '}',
        '',
        '\\usepackage{threeparttable}   % 表格加脚注',
        '\\usepackage{multirow}         % 表格跨行',
        '\\usepackage{longtable}        % 跨页表格',
        '\\usepackage{algorithm}',
        '\\usepackage{algorithmic}',
        '\\usepackage{siunitx}          % 量和单位',
        '',
        `% 参考文献：${spec.bibStyle === 'author-year' ? '著者-出版年制' : '顺序编码制'}`,
        bib.package,
        `\\bibliographystyle{${bib.style}}`,
        '',
        '% 图片默认在 figures/ 下',
        '\\graphicspath{{figures/}}',
        '',
        '\\makeatletter',
        '\\newcommand\\dif{%  % 微分符号',
        '  \\mathop{}\\mathord{}%',
        '  \\ifthu@math@style@TeX d\\else\\mathrm{d}\\fi',
        '}',
        '\\makeatother',
        '',
        '% hyperref 宏包在最后调用',
        '\\usepackage{hyperref}',
        '',
    ].join('\n');
}

function abstractFile(spec) {
    const keywordsZh = asArray(spec.abstract.keywordsZh).map((item) => asString(item)).filter((item) => item !== '');
    const keywordsEn = asArray(spec.abstract.keywordsEn).map((item) => asString(item)).filter((item) => item !== '');
    const lines = [`% !TEX root = ../${spec.fileName}`, ''];
    // 只有真的给了内容才写对应的摘要环境：空的关键词表在模板里会输出一个空行，
    // 而且「中文论文」只给中文摘要也是很常见的写法。
    if (spec.abstract.zh.length > 0 || keywordsZh.length > 0) {
        lines.push('\\begin{abstract}');
        for (const paragraph of spec.abstract.zh) lines.push(`  ${paragraph}`, '');
        lines.push('  % 关键词用“英文逗号”分隔，输出时自动处理为正确分隔符（不超过 5 个）');
        lines.push('  \\thusetup{');
        lines.push(`    keywords = {${keywordsZh.join(', ')}},`);
        lines.push('  }', '\\end{abstract}', '');
    }
    if (spec.abstract.en.length > 0 || keywordsEn.length > 0) {
        lines.push('\\begin{abstract*}');
        for (const paragraph of spec.abstract.en) lines.push(`  ${paragraph}`, '');
        lines.push('  % Use comma as separator when inputting (at most 5)');
        lines.push('  \\thusetup{');
        lines.push(`    keywords* = {${keywordsEn.join(', ')}},`);
        lines.push('  }', '\\end{abstract*}', '');
    }
    if (lines.length === 2) lines.push('% 没有提供摘要内容：中英文摘要是学位论文的必需部分。', '');
    return lines.join('\n');
}

function denotationFile(entries, fileName) {
    const lines = [`% !TEX root = ../${fileName}`, '', '\\begin{denotation}[3cm]'];
    for (const entry of entries) {
        if (typeof entry === 'string') {
            const at = entry.indexOf(':');
            const term = at === -1 ? entry : entry.slice(0, at);
            const note = at === -1 ? '' : entry.slice(at + 1);
            lines.push(`  \\item[${term.trim()}] ${note.trim()}`);
            continue;
        }
        const item = asObject(entry);
        lines.push(`  \\item[${asString(item.term ?? item.name)}] ${asString(item.note ?? item.definition ?? item.text)}`);
    }
    lines.push('\\end{denotation}', '');
    return lines.join('\n');
}

function chapterFile(chapter, fileName, ctx) {
    const lines = [`% !TEX root = ../${fileName}`, ''];
    lines.push(`\\chapter{${asString(chapter.title, `第 ${String(ctx.chapterIndex)} 章`)}}`);
    if (asString(chapter.label) !== '') lines.push(`\\label{${asString(chapter.label)}}`);
    lines.push('');
    lines.push(...renderBlocks(chapter.blocks ?? chapter.paragraphs, ctx));
    for (const rawSection of asArray(chapter.sections)) {
        const section = asObject(rawSection);
        lines.push(`\\section{${asString(section.title)}}`);
        if (asString(section.label) !== '') lines.push(`\\label{${asString(section.label)}}`);
        lines.push('');
        lines.push(...renderBlocks(section.blocks ?? section.paragraphs, ctx));
        for (const rawSub of asArray(section.subsections)) {
            const sub = asObject(rawSub);
            lines.push(`\\subsection{${asString(sub.title)}}`);
            if (asString(sub.label) !== '') lines.push(`\\label{${asString(sub.label)}}`);
            lines.push('');
            lines.push(...renderBlocks(sub.blocks ?? sub.paragraphs, ctx));
        }
    }
    return lines.join('\n');
}

function acknowledgementFile(text, fileName) {
    const paragraphs = Array.isArray(text) ? text : (text === undefined ? [] : [text]);
    const lines = [`% !TEX root = ../${fileName}`, '', '\\begin{acknowledgements}'];
    for (const paragraph of paragraphs.length > 0 ? paragraphs : ['（致谢）']) lines.push(`  ${asString(paragraph)}`, '');
    lines.push('\\end{acknowledgements}', '');
    return lines.join('\n');
}

function resumeFile(resume, fileName) {
    const data = asObject(resume);
    const bio = asArray(data.bio ?? data.paragraphs);
    const achievements = asArray(data.achievements);
    const lines = [`% !TEX root = ../${fileName}`, '', '\\begin{resume}', ''];
    if (bio.length > 0) {
        lines.push('  \\section*{个人简历}', '');
        for (const paragraph of bio) lines.push(`  ${asString(paragraph)}`, '');
    }
    if (achievements.length > 0) {
        lines.push('  \\section*{在学期间完成的相关学术成果}', '');
        lines.push('  \\subsection*{学术论文}', '');
        lines.push('  \\begin{achievements}');
        for (const item of achievements) lines.push(`    \\item ${asString(item)}`);
        lines.push('  \\end{achievements}', '');
    }
    lines.push('\\end{resume}', '');
    return lines.join('\n');
}

function committeeFile(committee, fileName) {
    const data = asObject(committee);
    const lines = [`% !TEX root = ../${fileName}`, '', '\\begin{committee}[name={学位论文指导小组、公开评阅人和答辩委员会名单}]', ''];
    lines.push('  \\newcolumntype{C}[1]{@{}>{\\centering\\arraybackslash}p{#1}}', '');
    const groups = asArray(data.groups).length > 0
        ? asArray(data.groups).map((item) => asObject(item))
        : [
            { title: '指导小组名单', rows: asArray(data.supervision) },
            { title: '公开评阅人名单', rows: asArray(data.reviewers) },
            { title: '答辩委员会名单', rows: asArray(data.defense) },
        ];
    for (const group of groups) {
        lines.push(`  \\section*{${asString(group.title, '名单')}}`, '');
        lines.push('  \\begin{center}', '    \\begin{tabular}{C{3cm}C{3cm}C{9cm}@{}}');
        for (const row of asArray(group.rows)) {
            lines.push(`      ${asArray(row).map((cell) => asString(cell)).join(' & ')} \\\\`);
        }
        lines.push('    \\end{tabular}', '  \\end{center}', '');
    }
    lines.push('\\end{committee}', '');
    return lines.join('\n');
}

function simpleEnvFile(text, env, fileName) {
    const paragraphs = Array.isArray(text) ? text : (text === undefined ? [] : [text]);
    const lines = [`% !TEX root = ../${fileName}`, '', `\\begin{${env}}`];
    for (const paragraph of paragraphs.length > 0 ? paragraphs : ['（待补充）']) lines.push(`  ${asString(paragraph)}`, '');
    lines.push(`\\end{${env}}`, '');
    return lines.join('\n');
}

/** 参考文献条目 → BibTeX。对象按 type 组装，字符串（以 @ 开头）原样保留。 */
function renderBibEntry(entry, fallbackKey) {
    if (typeof entry === 'string') return entry.trim().startsWith('@') ? `${entry.trim()}\n` : null;
    const item = asObject(entry);
    const type = asString(item.type, 'article');
    const key = asString(item.key, fallbackKey);
    const fieldOrder = ['author', 'title', 'journal', 'booktitle', 'publisher', 'address', 'school', 'institution', 'editor', 'edition', 'volume', 'number', 'pages', 'year', 'month', 'doi', 'url', 'note', 'language'];
    const lines = [`@${type}{${key},`];
    const used = new Set(['type', 'key']);
    for (const field of fieldOrder) {
        if (item[field] === undefined || item[field] === '') continue;
        used.add(field);
        lines.push(`  ${field.padEnd(10)}= {${String(item[field])}},`);
    }
    for (const [field, value] of Object.entries(item)) {
        if (used.has(field) || value === undefined || value === '') continue;
        lines.push(`  ${field.padEnd(10)}= {${String(value)}},`);
    }
    lines.push('}');
    return `${lines.join('\n')}\n`;
}

/** 从条目里推一个 bibtex key（作者姓氏 + 年份）。 */
function deriveKey(item, index) {
    const object = asObject(item);
    const author = asString(object.author, 'ref');
    const year = asString(object.year, String(2000 + (index % 25)));
    const first = author.split(/\s+and\s+/)[0].trim().split(/[,\s]+/)[0];
    const latin = first.replace(/[^A-Za-z]/g, '') || `ref${index + 1}`;
    return `${latin.toLowerCase()}${year}`;
}

/* ── create ─────────────────────────────────────────────────────────────── */

/** 归一化 spec：默认值、选项合法性、文件布局一次定下来。 */
function normalizeSpec(raw) {
    const spec = asObject(raw);
    const degree = DEGREES.includes(asString(spec.degree)) ? asString(spec.degree) : 'master';
    const degreeType = asString(spec.degreeType) === 'professional' ? 'professional' : 'academic';
    const abstract = asObject(spec.abstract);
    const toParagraphs = (...values) => {
        for (const value of values) {
            if (Array.isArray(value)) return value.map((item) => asString(item)).filter((item) => item !== '');
            if (typeof value === 'string' && value.trim() !== '') return [value];
        }
        return [];
    };

    // 主文件名一律 ASCII：中文文件名在 TeX 的日志、aux 与 bibtex 里都可能变乱码，
    // 而论文标题是中文 —— 标题写进 thusetup，文件名不跟着标题走。
    const wanted = asString(spec.path, '');
    const path = ensureExtension(wanted === '' ? 'thesis/thesis' : wanted, '.tex').replace(/\\/g, '/');
    const dir = dirname(path) === '.' ? '' : dirname(path);
    const fileName = path.slice(path.lastIndexOf('/') + 1);
    const chapters = asArray(spec.chapters).map((item, index) => ({
        data: asObject(item),
        file: `chap${String(index + 1).padStart(2, '0')}.tex`,
    }));

    return {
        raw: spec,
        path,
        dir,
        fileName,
        degree,
        degreeType,
        language: LANGUAGES.includes(asString(spec.language)) ? asString(spec.language) : 'chinese',
        fontset: FONTSETS.includes(asString(spec.fontset)) ? asString(spec.fontset) : 'windows',
        output: asString(spec.output) === 'electronic' ? 'electronic' : 'print',
        bibStyle: BIB_STYLES[asString(spec.bibStyle)] !== undefined ? asString(spec.bibStyle) : 'numeric',
        classOptions: [
            `degree=${degree}`,
            ...(degree === 'bachelor' ? [] : [`degree-type=${degreeType}`]),
            `language=${LANGUAGES.includes(asString(spec.language)) ? asString(spec.language) : 'chinese'}`,
            `fontset=${FONTSETS.includes(asString(spec.fontset)) ? asString(spec.fontset) : 'windows'}`,
        ].join(', '),
        title: asString(spec.title),
        titleEn: asString(spec.titleEn),
        degreeCategory: asString(spec.degreeCategory),
        degreeCategoryEn: asString(spec.degreeCategoryEn),
        department: asString(spec.department),
        discipline: asString(spec.discipline),
        disciplineEn: asString(spec.disciplineEn),
        disciplineLevel2: asString(spec.disciplineLevel2),
        studentId: asString(spec.studentId),
        author: asString(spec.author),
        authorEn: asString(spec.authorEn),
        supervisor: asString(spec.supervisor),
        supervisorEn: asString(spec.supervisorEn),
        associateSupervisor: asString(spec.associateSupervisor),
        associateSupervisorEn: asString(spec.associateSupervisorEn),
        coSupervisor: asString(spec.coSupervisor),
        professionalField: asString(spec.professionalField),
        professionalFieldEn: asString(spec.professionalFieldEn),
        secretLevel: asString(spec.secretLevel),
        secretYear: asString(spec.secretYear),
        clc: asString(spec.clc),
        udc: asString(spec.udc),
        id: asString(spec.id),
        startDate: asString(spec.startDate),
        date: asString(spec.date),
        includeSpine: spec.includeSpine === true,
        setup: asObject(spec.setup),
        chapters,
        abstract: {
            zh: toParagraphs(abstract.zh, abstract.text),
            en: toParagraphs(abstract.en, abstract.textEn),
            keywordsZh: asArray(abstract.keywordsZh ?? abstract.keywords),
            keywordsEn: asArray(abstract.keywordsEn),
        },
        include: {
            abstract: toParagraphs(abstract.zh, abstract.text).length > 0 || toParagraphs(abstract.en, abstract.textEn).length > 0,
            denotation: asArray(spec.denotation).length > 0,
            appendix: spec.appendix !== undefined,
            acknowledgements: spec.acknowledgements !== undefined,
            resume: spec.resume !== undefined,
            committee: spec.committee !== undefined && !SKIP_COMMITTEE.has(degree),
            comments: spec.comments !== undefined && !SKIP_COMMITTEE.has(degree),
            resolution: spec.resolution !== undefined && !SKIP_COMMITTEE.has(degree),
        },
    };
}

export function create(spec, env) {
    const normalized = normalizeSpec(spec);
    const raw = normalized.raw;
    const warnings = [];
    const contentFiles = [];
    const supportFiles = [];
    const projectDir = normalized.dir === '' ? '.' : normalized.dir;

    if (normalized.title === '') warnings.push('没有给 title：封面标题会是空的。');
    else if ([...normalized.title].length > 30) warnings.push(`标题 ${[...normalized.title].length} 字偏长，封面与书脊可能折行；可用 \\\\ 手动控制换行。`);
    if (normalized.chapters.length === 0) warnings.push('没有给 chapters：只生成了骨架，正文是空的。');
    if (normalized.include.abstract === false) warnings.push('没有给 abstract：中英文摘要是学位论文的必需部分。');
    if (shapeCount(normalized.abstract.keywordsZh) > 5 || shapeCount(normalized.abstract.keywordsEn) > 5) warnings.push('关键词超过 5 个；模板要求不超过 5 个。');
    if (raw.committee !== undefined && SKIP_COMMITTEE.has(normalized.degree)) warnings.push('本科生不需要 committee（指导小组与答辩委员会名单），已跳过。');

    // 模板文件：项目自包含才换机器也能编；拷不到就退化成用 TeX Live 自带的版本。
    const supportMode = asString(raw.support, 'auto');
    const template = supportMode === 'none' || supportMode === 'system'
        ? undefined
        : locateTemplateDir(env, raw.templateDir ?? env.config?.texTemplateDir);
    let support;
    if (supportMode === 'none' || supportMode === 'system') {
        support = { mode: supportMode, copied: [], note: supportMode === 'system' ? '按设置不拷贝模板：编译时用 TeX Live 自带的 thuthesis。' : '按设置不写模板文件。' };
    } else if (template === undefined) {
        support = { mode: 'missing', copied: [], note: '没找到 thuthesis 模板目录：源文件照样生成，编译时依赖 TeX Live 自带的版本。' };
        warnings.push(support.note);
    } else if (template.complete !== true) {
        support = { mode: 'texlive', copied: [], source: displayPath(env.root, template.dir), note: `只有 TeX Live 自带的 thuthesis（${displayPath(env.root, template.dir)}）；要与模板版本一致就给 create 传 templateDir。` };
    } else {
        const copied = copySupport(template.dir, projectDir, env);
        support = {
            mode: 'copied',
            source: displayPath(env.root, template.dir),
            copied: copied.copied.map((item) => item.name),
            missing: copied.missing,
            note: `已从 ${displayPath(env.root, template.dir)} 拷贝模板文件，项目自包含。`,
        };
        for (const item of copied.copied) supportFiles.push({ path: item.path, role: 'support', bytes: item.bytes });
        if (copied.missing.length > 0) warnings.push(`模板目录里缺 ${copied.missing.join('、')}：编译可能在封面或参考文献处失败。`);
    }

    const write = (relPath, content, role) => {
        const entry = env.writeFile(join(projectDir, relPath).replace(/\\/g, '/'), content);
        contentFiles.push({ path: entry.path, role, bytes: entry.bytes });
        return entry;
    };

    const existing = [];
    for (const name of [normalized.fileName, 'thusetup.tex']) {
        if (env.exists(join(projectDir, name).replace(/\\/g, '/'))) existing.push(name);
    }

    const ctx = createRenderContext(warnings, 1);

    const mainEntry = write(normalized.fileName, mainFile(normalized.path, normalized), 'main');
    write('thusetup.tex', setupFile(normalized), 'setup');
    if (normalized.include.abstract) write('data/abstract.tex', abstractFile(normalized), 'abstract');
    if (normalized.include.denotation) write('data/denotation.tex', denotationFile(asArray(raw.denotation), normalized.fileName), 'denotation');

    normalized.chapters.forEach((chapter, index) => {
        ctx.chapterIndex = index + 1;
        write(`data/${chapter.file}`, chapterFile(chapter.data, normalized.fileName, ctx), 'chapter');
    });

    if (normalized.include.appendix) {
        const appendix = asObject(raw.appendix);
        const blocks = appendix.blocks ?? (typeof raw.appendix === 'string' ? undefined : raw.appendix);
        ctx.chapterIndex = 'A';
        write('data/appendix.tex', chapterFile({ title: appendix.title ?? '补充内容', blocks }, normalized.fileName, ctx), 'appendix');
    }
    if (normalized.include.acknowledgements) write('data/acknowledgements.tex', acknowledgementFile(raw.acknowledgements, normalized.fileName), 'acknowledgements');
    if (normalized.include.resume) write('data/resume.tex', resumeFile(raw.resume, normalized.fileName), 'resume');
    if (normalized.include.committee) write('data/committee.tex', committeeFile(raw.committee, normalized.fileName), 'committee');
    if (normalized.include.comments) write('data/comments.tex', simpleEnvFile(raw.comments, 'comments', normalized.fileName), 'comments');
    if (normalized.include.resolution) write('data/resolution.tex', simpleEnvFile(raw.resolution, 'resolution', normalized.fileName), 'resolution');

    // 参考文献：没给条目也照样建 refs.bib —— 文件齐全比缺一个文件重要，
    // 而且 read() 会提醒「一条文献都没有」。
    const bibLines = ['% 参考文献数据库（BibTeX）。key 与正文 \\cite{...} 对应。', ''];
    asArray(raw.references).forEach((entry, index) => {
        const rendered = renderBibEntry(entry, deriveKey(entry, index));
        if (rendered !== null) bibLines.push(rendered);
    });
    write('ref/refs.bib', `${bibLines.join('\n').replace(/\n{3,}/g, '\n\n').replace(/\s+$/, '')}\n`, 'bibliography');

    // 图表文件是否真的存在：不存在的话编译会停在半路，写盘时就说清楚。
    for (const figure of ctx.figures) {
        if (figure.file === '') {
            warnings.push('有一个 figure 没给 file，已生成空的 \\includegraphics。');
            continue;
        }
        const direct = join(projectDir, figure.file).replace(/\\/g, '/');
        const inFigures = join(projectDir, 'figures', figure.file.split('/').pop()).replace(/\\/g, '/');
        if (!env.exists(direct) && !env.exists(inFigures)) {
            warnings.push(`图片文件不存在：${figure.file}（放到 ${projectDir === '.' ? '' : `${projectDir}/`}figures/ 下，或把 file 写成相对主文件的路径）。`);
        }
    }
    if (existing.length > 0) warnings.push(`覆盖了已存在的文件：${existing.join('、')}。create 是全量重建，只想改内容请用 edit。`);

    const files = [...contentFiles, ...supportFiles];
    const totalBytes = files.reduce((sum, file) => sum + (file.bytes ?? 0), 0);
    return {
        ok: true,
        format: 'tex',
        kind: 'project',
        path: mainEntry.path,
        main: mainEntry.path,
        projectDir: displayPath(env.root, env.resolve(projectDir)),
        bytes: totalBytes,
        theme: null,
        files,
        support,
        stats: {
            files: files.length,
            contentFiles: contentFiles.length,
            chapters: normalized.chapters.length,
            degree: normalized.degree,
            degreeType: normalized.degreeType,
            language: normalized.language,
            fontset: normalized.fontset,
            bibStyle: normalized.bibStyle,
            references: asArray(raw.references).length,
            bytes: totalBytes,
        },
        outline: [
            `学位：${normalized.degree}（${normalized.degreeType}）｜语言：${normalized.language}｜字体库：${normalized.fontset}｜参考文献：${normalized.bibStyle}`,
            `主文件：${mainEntry.path}（项目目录 ${displayPath(env.root, env.resolve(projectDir))}）`,
            ...normalized.chapters.map((chapter, index) => `第 ${index + 1} 章 ${plainTitle(asString(chapter.data.title)) || '（未命名）'}`),
            support.mode === 'copied' ? `模板文件：已拷贝 ${support.copied.length} 个（${support.copied.join('、')}）` : `模板文件：${support.note}`,
        ],
        warnings,
        next: `改完内容后用 office.tex.compile('${mainEntry.path}') 编出 PDF（先跑 office.tex.engines() 确认本机有 TeX 发行版）。`,
    };
}

function shapeCount(value) {
    return asArray(value).length;
}

/* ── 项目解析（read 共用） ──────────────────────────────────────────────── */

/** 主文件引用的所有 .tex（含主文件自己），深度优先、去重。 */
function collectProject(mainPath, env) {
    const mainAbsolute = env.resolve(mainPath);
    const out = [];
    const seen = new Set();
    const visit = (absolute, depth, appendix) => {
        if (depth > 3 || seen.has(absolute) || !existsSync(absolute)) return;
        seen.add(absolute);
        let text = '';
        try {
            text = readFileSync(absolute, 'utf8');
        } catch {
            return;
        }
        out.push({ absolute, rel: displayPath(env.root, absolute), text, isMain: out.length === 0, appendix });
        // \appendix 之后的 \input 是附录：它的 \chapter 不占正文的章号，
        // 报告里要分开数（否则「共几章」会把附录也算进去）。
        const appendixAt = text.indexOf('\\appendix');
        for (const match of text.matchAll(/\\(?:input|include)\s*\{([^}]*)\}/g)) {
            const name = match[1].trim();
            if (name === '') continue;
            const target = name.startsWith('/') ? name : join(dirname(absolute), ensureExtension(name, '.tex')).replace(/\\/g, '/');
            visit(target, depth + 1, appendix || (appendixAt !== -1 && match.index > appendixAt));
        }
    };
    visit(mainAbsolute, 0, false);
    out.projectDir = dirname(mainAbsolute);
    return out;
}

/** 找项目里真正的主文件（含 \documentclass 的那个）。 */
export function findMainFile(startPath, env) {
    const absolute = env.resolve(startPath);
    let text = '';
    try {
        text = readFileSync(absolute, 'utf8');
    } catch {
        return undefined;
    }
    if (/\\documentclass/.test(text)) return { absolute, rel: displayPath(env.root, absolute) };
    // 片段文件可以靠 `% !TEX root = ../xxx.tex` 指回主文件。
    const rootHint = /%\s*!TEX\s+root\s*=\s*(\S+)/.exec(text);
    if (rootHint !== null) {
        const candidate = join(dirname(absolute), rootHint[1]).replace(/\\/g, '/');
        if (existsSync(candidate)) return { absolute: candidate, rel: displayPath(env.root, candidate) };
    }
    // 再退一步：同目录或上一级目录里找唯一一个带 \documentclass 的 .tex。
    for (const dir of [dirname(absolute), dirname(dirname(absolute))]) {
        let names = [];
        try {
            names = readdirSync(dir);
        } catch {
            continue;
        }
        const mains = names.filter((name) => name.endsWith('.tex') && name !== 'thusetup.tex').filter((name) => {
            try {
                return /\\documentclass/.test(readFileSync(join(dir, name), 'utf8'));
            } catch {
                return false;
            }
        });
        if (mains.length === 1) {
            const candidate = join(dir, mains[0]).replace(/\\/g, '/');
            return { absolute: candidate, rel: displayPath(env.root, candidate) };
        }
    }
    return undefined;
}

/** 解析 .bib 里的条目 key → 类型。 */
function parseBibKeys(text) {
    const keys = new Map();
    for (const match of String(text ?? '').matchAll(/@(\w+)\s*\{\s*([^,\s}]+)\s*,/g)) keys.set(match[2], match[1].toLowerCase());
    return keys;
}

/** 段落级提醒：「一段太长了，该分段」。 */
function analyzeParagraphs(text, source, warnings) {
    const body = stripComments(text)
        .split(/\r?\n/)
        .filter((line) => !/^\s*\\/.test(line))
        .join('\n');
    for (const paragraph of body.split(/\n\s*\n/)) {
        const clean = paragraph.replace(/\\[a-zA-Z@]+\*?(\[[^\]]*\])?(\{[^{}]*\})?/g, ' ').trim();
        if (clean === '') continue;
        const { words } = countWords(clean);
        if (words > 800) warnings.push(`${source}：有一处段落约 ${words} 字，学位论文建议拆成多段。`);
    }
}

/* ── read ───────────────────────────────────────────────────────────────── */

export function read(path, env) {
    const target = env.resolve(path);
    if (!existsSync(target)) throw new Error(`找不到文件：${path}`);
    const text = readFileSync(target, 'utf8');
    const warnings = [];
    if (!/\\documentclass/.test(text)) return readFragment(path, text, env, warnings);

    const project = collectProject(path, env);
    const main = project[0];
    const joined = project.map((item) => item.text).join('\n');
    const stats = {
        files: project.length,
        chapters: 0,
        appendixChapters: 0,
        sections: 0,
        subsections: 0,
        figures: 0,
        tables: 0,
        equations: 0,
        citations: 0,
        citationKeys: 0,
        references: 0,
        words: 0,
        inputFiles: project.slice(1).map((item) => item.rel),
    };
    const outline = [];
    const figures = [];
    const citeKeys = new Set();
    const citeCounts = new Map();
    const labels = new Set();
    const refs = [];
    let citationCount = 0;

    for (const item of project) {
        const body = stripComments(item.text);
        const count = (pattern) => (body.match(pattern) ?? []).length;
        const chapterMatches = [...body.matchAll(/\\chapter\*?\s*\{([^}]*)\}/g)];
        const sectionMatches = [...body.matchAll(/\\section\*?\s*\{([^}]*)\}/g)];
        stats.sections += sectionMatches.length;
        stats.subsections += count(/\\subsection\*?\s*\{/g);
        stats.figures += count(/\\begin\{figure\*?\}/g);
        stats.tables += count(/\\begin\{(?:table|longtable)\*?\}/g);
        stats.equations += count(/\\begin\{(?:equation|align|gather|multline|eqnarray)\*?\}/g) + count(/\\\[/g);
        for (const match of body.matchAll(/\\(?:cite|citep|citet|upcite|parencite|textcite)\s*(?:\[[^\]]*\])*\s*\{([^}]*)\}/g)) {
            for (const key of match[1].split(',').map((entry) => entry.trim()).filter((entry) => entry !== '')) {
                citeKeys.add(key);
                citeCounts.set(key, (citeCounts.get(key) ?? 0) + 1);
                citationCount += 1;
            }
        }
        for (const match of body.matchAll(/\\label\s*\{([^}]*)\}/g)) labels.add(match[1].trim());
        for (const match of body.matchAll(/\\(?:ref|eqref|autoref|pageref)\s*\{([^}]*)\}/g)) {
            for (const key of match[1].split(',').map((entry) => entry.trim()).filter((entry) => entry !== '')) refs.push({ key, source: item.rel });
        }
        for (const match of body.matchAll(/\\includegraphics\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}/g)) {
            figures.push({ file: match[1].trim(), source: item.rel });
        }
        stats.words += countWords(item.text).words;
        analyzeParagraphs(item.text, item.rel, warnings);

        if (chapterMatches.length > 0) {
            if (item.appendix) {
                stats.appendixChapters += chapterMatches.length;
                outline.push(`附录 ${plainTitle(chapterMatches[0][1])}`);
            } else {
                stats.chapters += chapterMatches.length;
                outline.push(`第 ${stats.chapters - chapterMatches.length + 1} 章 ${plainTitle(chapterMatches[0][1])}`);
            }
        }
        for (const match of sectionMatches) outline.push(`  ${plainTitle(match[1])}`);
    }
    stats.citations = citationCount;
    stats.citationKeys = citeKeys.size;

    // 参考文献：\bibliography{ref/refs} → ref/refs.bib（相对主文件所在目录）
    const bibKeys = new Map();
    for (const match of main.text.matchAll(/\\(?:bibliography|addbibresource)\s*\{([^}]*)\}/g)) {
        for (const raw of match[1].split(',').map((entry) => entry.trim()).filter((entry) => entry !== '')) {
            const file = join(dirname(main.absolute), ensureExtension(raw, '.bib')).replace(/\\/g, '/');
            if (!existsSync(file)) {
                warnings.push(`参考文献库不存在：${displayPath(env.root, file)}（主文件的 \\bibliography 指向它）。`);
                continue;
            }
            for (const [key, type] of parseBibKeys(readFileSync(file, 'utf8'))) if (!bibKeys.has(key)) bibKeys.set(key, type);
        }
    }
    stats.references = bibKeys.size;
    const missingKeys = [...citeKeys].filter((key) => !bibKeys.has(key));
    if (missingKeys.length > 0) {
        warnings.push(`正文引用了 ${missingKeys.length} 个 refs.bib 里没有的文献：${missingKeys.slice(0, 6).join('、')}${missingKeys.length > 6 ? ' …' : ''}；先补条目再编译，否则参考文献处会出现 [?]。`);
    }
    if (bibKeys.size > 0 && citeKeys.size === 0) warnings.push(`refs.bib 里有 ${bibKeys.size} 条文献却一处都没引用：正文用 \\cite{key} 引用，或删掉多余条目。`);
    if (stats.references === 0) warnings.push('参考文献库里没有任何条目：学位论文的参考文献不能为空。');
    const missingRefs = [...new Set(refs.filter((item) => !labels.has(item.key)).map((item) => item.key))];
    if (missingRefs.length > 0) warnings.push(`\\ref 指向了不存在的 label：${missingRefs.slice(0, 6).join('、')}（图表/公式要用 \\label 打标签）。`);

    // 图片与 \input 指向的文件在不在：编译会在这里直接停下，写盘时就该说清楚。
    for (const figure of figures) {
        const direct = join(dirname(main.absolute), figure.file).replace(/\\/g, '/');
        const inFigures = join(dirname(main.absolute), 'figures', figure.file.split('/').pop()).replace(/\\/g, '/');
        if (!existsSync(direct) && !existsSync(inFigures)) warnings.push(`图片文件不存在：${figure.file}（${figure.source}）`);
    }
    for (const item of project) {
        for (const match of item.text.matchAll(/\\(?:input|include)\s*\{([^}]*)\}/g)) {
            const name = match[1].trim();
            if (name === '') continue;
            const candidate = join(dirname(item.absolute), ensureExtension(name, '.tex')).replace(/\\/g, '/');
            if (!existsSync(candidate)) warnings.push(`\\input 指向的文件不存在：${item.rel} → ${name}`);
        }
    }

    // 结构与必需部分
    if (stats.chapters === 0) warnings.push('正文里一个 \\chapter 都没有：学位论文必须有章节。');
    if (stats.sections === 0 && stats.chapters > 0) warnings.push('有章节但没有 \\section：建议按节组织内容。');
    if (!/\\begin\{abstract\}/.test(joined)) warnings.push('没有中英文摘要（\\begin{abstract}）：学位论文必需。');
    if (!/\\begin\{acknowledgements\}/.test(joined)) warnings.push('没有致谢（\\begin{acknowledgements}）：学位论文必需。');
    if (!/\\tableofcontents/.test(main.text)) warnings.push('主文件里没有 \\tableofcontents：学位论文必须有目录。');
    for (const item of project) {
        for (const match of item.text.matchAll(/\\begin\{(figure|table)\*?\}([\s\S]*?)\\end\{\1\*?\}/g)) {
            if (!/\\caption\s*\{/.test(match[2])) warnings.push(`${item.rel}：有 ${match[1]} 环境没有 \\caption（图/表必须有题注）。`);
            if (!/\\label\s*\{/.test(match[2])) warnings.push(`${item.rel}：有 ${match[1]} 环境没有 \\label，正文无法引用它。`);
        }
    }
    const titleMatch = /title\s*=\s*\{([^}]*)\}/.exec(joined);
    if (titleMatch !== null) {
        const length = [...plainTitle(titleMatch[1])].length;
        if (length > 30) warnings.push(`论文标题 ${length} 字偏长，封面可能折行；建议精简或用 \\\\ 手动换行。`);
    }

    const options = {};
    const classOptions = /\\documentclass\s*\[([^\]]*)\]\s*\{([^}]*)\}/.exec(main.text);
    if (classOptions !== null) {
        for (const part of classOptions[1].split(',')) {
            const [key, value] = part.split('=').map((entry) => entry.trim());
            if (key !== '') options[key] = value ?? true;
        }
    }
    stats.degree = asString(options.degree, 'master');
    stats.degreeType = asString(options['degree-type'], 'academic');
    stats.language = asString(options.language, 'chinese');
    stats.fontset = asString(options.fontset, 'windows');

    return {
        ok: true,
        format: 'tex',
        kind: 'main',
        path: displayPath(env.root, target),
        main: displayPath(env.root, target),
        projectDir: displayPath(env.root, dirname(target)),
        bytes: Buffer.byteLength(text, 'utf8'),
        theme: null,
        stats,
        outline: [
            `学位：${stats.degree}（${stats.degreeType}）｜语言：${stats.language}｜字体库：${stats.fontset}`,
            `共 ${stats.files} 个 .tex（含 ${stats.inputFiles.length} 个 \\input 文件）｜约 ${stats.words} 字｜${stats.chapters} 章 / ${stats.sections} 节${stats.appendixChapters > 0 ? ` / 附录 ${stats.appendixChapters} 章` : ''}`,
            `图表公式：图 ${stats.figures}、表 ${stats.tables}、公式 ${stats.equations}｜引用 ${stats.citations} 处（${stats.citationKeys} 条文献）/ refs.bib ${stats.references} 条`,
            ...outline,
        ],
        warnings,
    };
}

/** 片段文件（data/chap01.tex 这种）：只报它自己有什么，不拿整篇论文的标准挑毛病。 */
function readFragment(path, text, env, warnings) {
    const body = stripComments(text);
    const count = (pattern) => (body.match(pattern) ?? []).length;
    const sections = [...body.matchAll(/\\(?:chapter|section|subsection)\*?\s*\{([^}]*)\}/g)].map((match) => plainTitle(match[1]));
    const words = countWords(text).words;
    analyzeParagraphs(text, displayPath(env.root, env.resolve(path)), warnings);
    for (const match of text.matchAll(/\\begin\{(figure|table)\*?\}([\s\S]*?)\\end\{\1\*?\}/g)) {
        if (!/\\caption\s*\{/.test(match[2])) warnings.push(`有 ${match[1]} 环境没有 \\caption（图/表必须有题注）。`);
    }
    return {
        ok: true,
        format: 'tex',
        kind: 'fragment',
        path: displayPath(env.root, env.resolve(path)),
        bytes: Buffer.byteLength(text, 'utf8'),
        theme: null,
        stats: {
            words,
            chapters: count(/\\chapter\*?\s*\{/g),
            sections: count(/\\section\*?\s*\{/g),
            subsections: count(/\\subsection\*?\s*\{/g),
            figures: count(/\\begin\{figure\*?\}/g),
            tables: count(/\\begin\{(?:table|longtable)\*?\}/g),
            equations: count(/\\begin\{(?:equation|align|gather|multline|eqnarray)\*?\}/g) + count(/\\\[/g),
            citations: count(/\\(?:cite|citep|citet|upcite)\s*(?:\[[^\]]*\])*\s*\{/g),
            lines: text.split(/\r?\n/).length,
        },
        outline: [`片段文件（不是主文件）：${sections.length} 个标题｜约 ${words} 字`, ...sections.map((item) => `  ${item}`)],
        warnings,
    };
}

/* ── edit ───────────────────────────────────────────────────────────────── */

/** 项目里的所有 .tex（主文件 + \input）；传进来的可以是主文件，也可以是片段。 */
function projectTexts(path, env) {
    const main = findMainFile(path, env);
    if (main === undefined) {
        const absolute = env.resolve(path);
        if (!existsSync(absolute)) throw new Error(`找不到文件：${path}`);
        return { main: undefined, items: [{ absolute, rel: displayPath(env.root, absolute), text: readFileSync(absolute, 'utf8') }] };
    }
    return { main, items: collectProject(main.rel, env).map((item) => ({ absolute: item.absolute, rel: item.rel, text: item.text })) };
}

/** 在 \thusetup{...} 里写字段：命中就替换值，没命中就补在最后一个 \thusetup 块的末尾。 */
function setSetupFields(text, pairs) {
    let body = text;
    const applied = [];
    const skipped = [];
    for (const [key, value] of pairs) {
        const rendered = typeof value === 'boolean' || typeof value === 'number' ? String(value) : `{${String(value)}}`;
        const pattern = new RegExp(`(^|[\\s{,])(${escapeRegExp(key)}\\s*=\\s*)(\\{(?:[^{}]|\\{[^{}]*\\})*\\}|[^,\\n}]+)`, 'm');
        const match = pattern.exec(body);
        if (match !== null) {
            body = `${body.slice(0, match.index)}${match[1]}${match[2]}${rendered}${body.slice(match.index + match[0].length)}`;
            applied.push(`\\thusetup.${key} = ${rendered}`);
            continue;
        }
        let searchFrom = 0;
        let last;
        for (;;) {
            const found = findBracedArg(body, '\\thusetup', searchFrom);
            if (found === undefined) break;
            last = found;
            searchFrom = found.end;
        }
        if (last === undefined) {
            skipped.push({ key, reason: '这个文件里没有 \\thusetup{} 块' });
            continue;
        }
        body = `${body.slice(0, last.end)}\n  ${key} = ${rendered},${body.slice(last.end)}`;
        applied.push(`\\thusetup.${key} = ${rendered}（新增）`);
    }
    return { text: body, applied, skipped };
}

export function edit(path, ops, env) {
    const { main, items } = projectTexts(path, env);
    const byRel = new Map(items.map((item) => [item.rel.replace(/\\/g, '/'), { ...item, changed: false }]));
    const applied = [];
    const skipped = [];
    const changes = [];
    const projectDir = main === undefined ? dirname(env.resolve(path)) : dirname(main.absolute);

    const pick = (file) => {
        if (file === undefined) return [...byRel.values()];
        const wanted = String(file).replace(/\\/g, '/').replace(/^\.\//, '');
        const hit = byRel.get(wanted) ?? [...byRel.values()].find((item) => item.rel.endsWith(wanted));
        return hit === undefined ? [] : [hit];
    };
    /**
     * 找项目里的文件（按后缀匹配，与 pick 同一套规则），找不到就是「要新建一个」。
     * 这条必须走后缀匹配：调用方给的是项目内路径（data/chap02.tex），
     * 而 byRel 的键是相对工作目录的路径（thesis/data/chap02.tex）——
     * 直接用键查会查不到，然后新建一份只含追加内容的同名文件，把原文件覆盖掉。
     */
    const resolveTarget = (file) => {
        const hit = pick(file)[0];
        if (hit !== undefined) return hit;
        const rel = String(file).replace(/\\/g, '/');
        return { absolute: env.resolve(join(projectDir, rel)), rel: projectPath(rel), text: '', changed: false };
    };
    /** 新文件在项目里叫 rel（写进 \input），但落盘路径要相对工作目录。 */
    const projectPath = (rel) => displayPath(env.root, env.resolve(join(projectDir, rel).replace(/\\/g, '/')));

    for (const raw of asArray(ops)) {
        const op = asObject(raw);
        const where = Object.keys(op).join(',') || '(空)';

        if (op.find !== undefined) {
            const find = asString(op.find);
            if (find === '') {
                skipped.push({ op: where, reason: 'find 不能为空' });
                continue;
            }
            const targets = pick(op.file);
            if (targets.length === 0) {
                skipped.push({ op: where, reason: `项目里找不到文件：${String(op.file)}` });
                continue;
            }
            let hits = 0;
            for (const item of targets) {
                const count = item.text.split(find).length - 1;
                if (count === 0) continue;
                if (op.all === false) {
                    item.text = item.text.replace(find, asString(op.replace));
                    hits += 1;
                    applied.push(`${item.rel}：${find} → ${asString(op.replace)}`);
                } else {
                    item.text = item.text.split(find).join(asString(op.replace));
                    hits += count;
                    applied.push(`${item.rel}：${find} → ${asString(op.replace)} ×${count}`);
                }
                item.changed = true;
                changes.push({ file: item.rel, replace: find, count: op.all === false ? 1 : count });
            }
            if (hits === 0) skipped.push({ op: where, reason: '未命中' });
            continue;
        }

        if (op.set !== undefined) {
            const pairs = Object.entries(asObject(op.set));
            if (pairs.length === 0) {
                skipped.push({ op: where, reason: 'set 里没有任何字段' });
                continue;
            }
            // 默认改 thusetup.tex；也可以用 file 指到别处（摘要的关键词就在 abstract.tex）。
            const targets = op.file !== undefined
                ? pick(op.file)
                : (pick('thusetup.tex').length > 0 ? pick('thusetup.tex') : [...byRel.values()].slice(0, 1));
            if (targets.length === 0) {
                skipped.push({ op: where, reason: '没找到 thusetup.tex，请用 file 指定目标文件' });
                continue;
            }
            const item = targets[0];
            const result = setSetupFields(item.text, pairs);
            if (result.applied.length > 0) {
                item.text = result.text;
                item.changed = true;
                applied.push(...result.applied.map((line) => `${item.rel}：${line}`));
                changes.push({ file: item.rel, set: pairs.map(([key, value]) => `${key}=${String(value)}`) });
            }
            skipped.push(...result.skipped.map((entry) => ({ op: where, reason: `${entry.key}：${entry.reason}` })));
            continue;
        }

        if (op.appendTo !== undefined) {
            const spec = typeof op.appendTo === 'string' ? { file: op.appendTo } : asObject(op.appendTo);
            const text = asString(spec.text ?? op.text);
            const rel = asString(spec.file);
            if (rel === '' || text === '') {
                skipped.push({ op: where, reason: 'appendTo 需要 {file, text}' });
                continue;
            }
            const target = resolveTarget(rel);
            target.text = target.text === '' ? `${text}\n` : `${target.text.replace(/\s+$/, '')}\n\n${text}\n`;
            target.changed = true;
            byRel.set(target.rel, target);
            applied.push(`${target.rel}：追加 ${[...text].length} 字`);
            changes.push({ file: target.rel, append: text.slice(0, 40) });
            continue;
        }

        if (op.appendChapter !== undefined) {
            if (main === undefined) {
                skipped.push({ op: where, reason: '找不到主文件，无法插入 \\input' });
                continue;
            }
            const chapter = asObject(op.appendChapter);
            const mainItem = byRel.get(main.rel.replace(/\\/g, '/'));
            const used = new Set([...byRel.values()].map((item) => (/(chap\d+)\.tex$/.exec(item.rel) ?? [])[1]).filter(Boolean));
            let index = 1;
            while (used.has(`chap${String(index).padStart(2, '0')}`)) index += 1;
            const file = `data/chap${String(index).padStart(2, '0')}.tex`;
            const ctx = createRenderContext([], index);
            const content = chapterFile(chapter, main.rel.slice(main.rel.lastIndexOf('/') + 1), ctx);
            const rel = projectPath(file);
            byRel.set(rel, { absolute: env.resolve(join(projectDir, file)), rel, text: content, changed: true });
            // 插在 \bibliography 之前；没有就插在 \appendix / \backmatter 之前。
            // 写法与 create 生成的主文件保持一致：\input 不带 .tex 后缀。
            const anchor = /\\bibliography\s*\{/.exec(mainItem.text) ?? /\\appendix\b/.exec(mainItem.text) ?? /\\backmatter\b/.exec(mainItem.text);
            const input = `\\input{${file.replace(/\.tex$/, '')}}`;
            if (anchor === null) mainItem.text = mainItem.text.replace(/\\end\{document\}/, `${input}\n\n\\end{document}`);
            else mainItem.text = `${mainItem.text.slice(0, anchor.index)}${input}\n${mainItem.text.slice(anchor.index)}`;
            mainItem.changed = true;
            applied.push(`${rel}：新增章节「${asString(chapter.title)}」，并插入 \\input{${file}}`);
            changes.push({ file: rel, chapter: asString(chapter.title) });
            continue;
        }

        if (op.addReference !== undefined) {
            const entry = op.addReference;
            const key = typeof entry === 'string'
                ? (/(?:@\w+\s*\{\s*)([^,\s}]+)/.exec(entry) ?? [])[1]
                : asString(asObject(entry).key);
            const rendered = renderBibEntry(entry, key ?? `ref${Date.now().toString(36)}`);
            if (rendered === null) {
                skipped.push({ op: where, reason: 'addReference 只接受 {key, type, ...} 对象或以 @ 开头的 BibTeX 字符串' });
                continue;
            }
            let bibItem = [...byRel.values()].find((item) => item.rel.endsWith('.bib'));
            if (bibItem === undefined) {
                // 文献库必须先读出来再追加：collectProject 只收集 .tex，
                // 若从空字符串开始，整份 refs.bib 会被这一条覆盖掉。
                const mainText = main === undefined ? '' : (byRel.get(main.rel.replace(/\\/g, '/'))?.text ?? '');
                const declared = /\\bibliography\s*\{([^}]*)\}/.exec(mainText) ?? /\\addbibresource\s*\{([^}]*)\}/.exec(mainText);
                const first = declared === null ? '' : declared[1].split(',')[0].trim();
                const rel = projectPath(first === '' ? 'ref/refs.bib' : ensureExtension(first, '.bib'));
                bibItem = { absolute: env.resolve(rel), rel, text: env.exists(rel) ? env.readText(rel) : '', changed: false };
            }
            if (key !== undefined && key !== '' && parseBibKeys(bibItem.text).has(key)) {
                skipped.push({ op: where, reason: `文献 ${key} 已经存在（同 key 不重复加）` });
                continue;
            }
            bibItem.text = `${bibItem.text.replace(/\s+$/, '')}\n\n${rendered}`;
            bibItem.changed = true;
            byRel.set(bibItem.rel, bibItem);
            applied.push(`${bibItem.rel}：新增文献 ${key ?? '(无 key)'}`);
            changes.push({ file: bibItem.rel, reference: key });
            continue;
        }

        skipped.push({ op: where, reason: `未知操作，支持的键：find/replace、set、appendTo、appendChapter、addReference（收到 ${where}）` });
    }

    const writtenFiles = [];
    for (const item of byRel.values()) {
        if (!item.changed) continue;
        const entry = env.writeFile(item.rel, item.text);
        writtenFiles.push({ path: entry.path, bytes: entry.bytes });
    }

    const report = read(main === undefined ? path : main.rel, env);
    return {
        ...report,
        ok: true,
        path: displayPath(env.root, env.resolve(path)),
        applied,
        skipped,
        changes,
        changed: writtenFiles.length > 0,
        files: writtenFiles,
        bytes: writtenFiles.reduce((sum, file) => sum + file.bytes, 0),
    };
}

/* ── api（挂在 office.tex 上，与 create / read / edit 并列） ──────────────── */

async function compileProject(env, path, options) {
    const wanted = asString(path);
    if (wanted === '') throw new Error('office.tex.compile 需要一个主文件路径（含 \\documentclass 的那个 .tex）。');
    const main = findMainFile(wanted, env);
    const report = await compileTex(main === undefined ? wanted : main.rel, env, { ...asObject(options), config: env.config });
    if (report.ok !== true) {
        // 编译失败时把常见成因直接翻译出来，别让模型去读几百行 .log。
        const diagnosis = [];
        const tail = report.log.tail;
        if (report.timedOut !== true && report.log.errors.length === 0) {
            if (/fontspec|Font .* not found|The font/.test(tail)) diagnosis.push('字体库问题：Windows 上建议 fontset=windows；换平台时模板需要该平台的字体。');
            if (/File .* not found|not found on input/.test(tail)) diagnosis.push('有文件找不到：先照着 read() 的 warnings 检查图片与 \\input 路径，再确认模板文件（thuthesis.cls 与两个校徽）齐不齐。');
            if (/I found no \\citation|empty bibliography/.test(tail)) diagnosis.push('参考文献是空的：先 addReference 或直接编辑 ref/refs.bib 加条目，再在正文里 \\cite。');
        }
        report.diagnosis = diagnosis;
    }
    return report;
}

/** 清掉中间产物（.aux/.log/.synctex.gz 等），保留 PDF 与源文件。 */
async function clearAux(env, path, mode) {
    const main = findMainFile(asString(path), env);
    if (main === undefined) throw new Error('找不到 LaTeX 主文件（含 \\documentclass 的 .tex）。');
    const fileName = main.rel.slice(main.rel.lastIndexOf('/') + 1);
    return cleanBuild(env, dirname(main.absolute), fileName, mode === 'all' ? 'all' : 'aux', 'manual');
}

export const api = {
    compile: (env) => (path, options) => compileProject(env, path, options),
    engines: () => () => texEngines(),
    template: (env) => (dir) => locateTemplateDir(env, dir),
    escape: () => (text) => escapeLatex(text),
    clearAux: (env) => (path, mode) => clearAux(env, path, mode),
};

/* ── meta ───────────────────────────────────────────────────────────────── */

/** 长文档：进 office_help({topic:'tex'})。写法与细节放这里，工具声明只有两行。 */
const GUIDE = `LaTeX 学位论文（office.tex）—— thuthesis 模板

一句话：office.tex.create() 一次写出**能直接编译的整篇论文项目**（主文件 +
thusetup.tex + data/ 按章分文件 + ref/refs.bib + 模板运行期文件），
office.tex.compile() 调本机 TeX 发行版编出 PDF。版式由 thuthesis.cls 决定，
插件只保证「结构、字段、文件、引用」都对 —— 手工拼版式反而过不了格式审查。

项目结构（create 一次写齐）
  thesis/thesis.tex        主文件：\\documentclass + 封面/目录/正文/参考文献骨架
  thesis/thusetup.tex      论文基本信息（标题/作者/导师/学科…）与宏包
  thesis/data/abstract.tex 中英文摘要与关键词
  thesis/data/chap01.tex   第 1 章（chap02、chap03… 依次）
  thesis/data/denotation.tex        符号对照表（给了才写）
  thesis/data/acknowledgements.tex  致谢；resume.tex 简历与成果；appendix.tex 附录；
  thesis/data/committee.tex         指导小组/评阅人/答辩委员会名单
  thesis/ref/refs.bib      参考文献库
  thesis/thuthesis.cls …   模板运行期文件（从工作区的 thuthesis 模板目录整套拷贝）

基本用法（一次调用里生成 + 编译）
  const t = office.tex.create({
    path: 'thesis/thesis.tex',          // 主文件；文件名用 ASCII（中文标题写 title）
    title: '基于××的××研究', titleEn: 'Research on ...',
    degree: 'master',                   // doctor | master | bachelor | postdoc
    degreeType: 'academic',             // academic | professional
    author: '张三', authorEn: 'Zhang San',
    supervisor: '李四, 教授', supervisorEn: 'Professor Li Si',
    department: '计算机科学与技术系',
    discipline: '计算机科学与技术', disciplineEn: 'Computer Science and Technology',
    degreeCategory: '工学硕士', degreeCategoryEn: 'Master of Science',
    date: '2026-06-01',
    abstract: { zh: ['摘要正文…'], en: ['Abstract…'],
                keywordsZh: ['关键词1', '关键词2'], keywordsEn: ['keyword 1', 'keyword 2'] },
    chapters: [
      { title: '绪论', sections: [
        { title: '研究背景', blocks: ['第一段…', { items: ['第一点', '第二点'] }] },
      ] },
    ],
    references: [{ key: 'zhang2024', type: 'article', author: '张三 and 李四',
                   title: '论文题名', journal: '某学报', year: '2024', volume: '34', pages: '1--7' }],
    acknowledgements: ['感谢…'],
    resume: { bio: ['1998 年生于…'], achievements: ['Zhang S. Paper title[J]. Journal, 2025.'] },
  });
  const r = await office.tex.compile('thesis/thesis.tex');
  // r.ok / r.pdf.path / r.pdf.pages / r.log.errors / r.hint

内容块（blocks）：字符串＝一段正文；对象形式有
  {p: '正文'}                             段落
  {items: ['…'], ordered: true}           列表
  {equation: 'E = mc^2', label}           公式（自动编号）
  {align: ['a &= b \\\\', 'c &= d']}      多行对齐公式
  {figure: {file: 'figures/x.pdf', caption: '题注', label, width: 0.7, widthCm}}
  {table: {caption: '题注', header: ['列1','列2'], rows: [['a','b']], align: 'lcc'}}  三线表
  {code: '…'}                             verbatim 代码块
  {raw: '\\vspace{1cm}'}                   原样插入 LaTeX
section 下面还可以有 subsections: [{title, blocks}]。

文本按 LaTeX 源码处理：标题、段落、表格单元格里想写字面的 % & _ # $ 要自己转义，
例如 '达成率 95\\%'、'变量 a\\_1'、'A \\& B'。手里是普通文本（从别处复制来的）
就先过一遍 office.tex.escape(text)。中文标点直接写，ctex 会处理。

改内容（office.tex.edit，一次调用做完所有修改）
  [{find: '旧结论', replace: '新结论'}]           跨项目所有 .tex 的替换（file 可限定单文件）
  [{set: {title: '新标题', date: '2026-07-01'}}]  改 \\thusetup 字段（默认改 thusetup.tex）
  [{set: {keywords: '新关键词1, 新关键词2'}, file: 'data/abstract.tex'}]
  [{appendTo: {file: 'data/chap01.tex', text: '补充一段…'}}]
  [{appendChapter: {title: '实验', sections: [{title: '设置', blocks: ['…']}]}}]  新章 + 自动插 \\input
  [{addReference: {key: 'li2025', type: 'inproceedings', author: 'Li M', title: '…',
                   booktitle: '…', year: '2025'}}]
edit 之后要重新 compile 才会更新 PDF。

编译（office.tex.compile）
  const r = await office.tex.compile('thesis/thesis.tex', { timeoutMs: 600000, clean: 'aux' });
  - r.ok 为 true 才算编出 PDF：r.pdf = {path, bytes, pages}；
  - r.log.errors 是解析过的真实错误（文件:行号: 说明），r.log.tail 是日志尾部，
    完整日志在 r.log.path（可以 office.files.read 读）；
  - r.log.overfull / underfull 是排版提示，不一定是错误；
  - r.support 说明模板文件从哪来（项目自带 / TeX Live 自带）；
  - 失败先看 r.log.errors 第一条与 r.diagnosis，不要盲改源文件；
  - office.tex.engines() 看本机有什么（latexmk / xelatex / lualatex / bibtex）；
    一个都没有时只能生成源文件，编不出 PDF；
  - office.tex.clearAux('thesis/thesis.tex') 清中间产物（.aux/.log/.synctex.gz）。

本机没有 TeX 发行版时
  create / read / edit 照常可用（源文件本身就是交付物），只有 compile 会明确报错。
  装 TeX Live 或 MiKTeX 之后不用改配置，插件自动探测。

常见编译错误对照
  ! Undefined control sequence.             命令拼错，或用了没载入的宏包
  ! LaTeX Error: File \`x.pdf' not found.   图片路径不对（默认在 figures/ 下找）
  ! Package fontspec Error                  fontset 与当前平台不匹配（Windows 用 windows）
  ! Missing $ inserted                      数学符号没放进 $...$ 或公式环境
  \\input 指向的文件不存在                    先看 read() 的 warnings

注意
  - 本科生（degree=bachelor）不需要 committee / comments / resolution，给了也会跳过；
  - 关键词不超过 5 个；\\thusetup{} 里不要出现空行；
  - 交付前建议用 office.pdf.pages() 渲染封面与目录几页 + read_image 亲眼看一下；
  - 字体、行距、页边距、题注格式都由模板负责，不要自己写 \\vspace、\\fontsize 调版式。`;

/**
 * office_help 的默认层（缓存提示词预算：全文那份 6.4 KB 只留给 detail:true）。
 * 只写「照着它就能写出对的调用」—— 操作名、必需字段、块语法、转义规则与
 * 编译结果的读法。完整示例、错误对照表与版式约定在 GUIDE 里。
 */
const GUIDE_BRIEF = `LaTeX 学位论文（office.tex）—— thuthesis 模板

一句话：office.tex.create() 一次写出**能直接编译的整篇论文项目**（主文件 +
thusetup.tex + data/ 按章分文件 + ref/refs.bib + 模板运行期文件），
office.tex.compile() 调本机 TeX 发行版编出 PDF。版式由 thuthesis.cls 决定，
插件只保证「结构、字段、文件、引用」都对 —— 手工拼版式反而过不了格式审查。

基本用法（一次调用里生成 + 编译）
  const t = office.tex.create({
    path: 'thesis/thesis.tex',          // 主文件；文件名用 ASCII（中文标题写 title）
    title: '基于××的××研究', titleEn: 'Research on ...',
    degree: 'master',                   // doctor | master | bachelor | postdoc
    degreeType: 'academic',             // academic | professional
    author: '张三', authorEn: 'Zhang San',
    supervisor: '李四, 教授', department: '计算机科学与技术系',
    discipline: '计算机科学与技术', degreeCategory: '工学硕士', date: '2026-06-01',
    abstract: { zh: ['摘要正文…'], en: ['Abstract…'],
                keywordsZh: ['关键词1', '关键词2'], keywordsEn: ['keyword 1', 'keyword 2'] },
    chapters: [{ title: '绪论', sections: [
      { title: '研究背景', blocks: ['第一段…', { items: ['第一点', '第二点'] }] }] }],
    references: [{ key: 'zhang2024', type: 'article', author: '张三 and 李四',
                   title: '论文题名', journal: '某学报', year: '2024', volume: '34', pages: '1--7' }],
    acknowledgements: ['感谢…'],
    resume: { bio: ['1998 年生于…'], achievements: ['Zhang S. Paper title[J]. Journal, 2025.'] },
  });
  const r = await office.tex.compile('thesis/thesis.tex');
  // r.ok / r.pdf.path / r.pdf.pages / r.log.errors / r.hint

项目结构（create 一次写齐）：thesis.tex 主文件 + thusetup.tex 基本信息 +
data/{abstract,chap01,chap02,…,denotation,acknowledgements,resume,appendix,committee}.tex +
ref/refs.bib + thuthesis.cls 等模板运行期文件（整套从模板目录拷贝）。

内容块（blocks）：字符串＝段正文；对象形式有
  {p} 段落 / {items: [...], ordered} 列表 / {equation, label} 公式（自动编号）/
  {align: ['a &= b \\\\', 'c &= d']} 多行对齐 / {figure: {file, caption, label, width}} /
  {table: {caption, header, rows, align: 'lcc'}} 三线表 / {code} 代码块 / {raw} 原样插入。
section 下面还可以有 subsections: [{title, blocks}]。

文本按 LaTeX 源码处理：要写字面的 % & _ # $ 得自己转义（'达成率 95\\%'、'a\\_1'）；
普通文本先过 office.tex.escape(text)。中文标点直接写，ctex 会处理。

改内容（office.tex.edit，一次调用改完所有位置）
  [{find, replace, all, file?}] 跨项目字面替换 / [{set: {title, date}, file?}] 改 \\thusetup 字段 /
  [{appendTo: {file, text}}] / [{appendChapter: {title, sections}}] 新建一章并自动插 \\input /
  [{addReference: {key, type, author, title, year, …}}]。edit 之后要重新 compile 才更新 PDF。

编译（office.tex.compile(path, {timeoutMs, clean, engine, support})）
  r.ok 为 true 才算编出 PDF：r.pdf = {path, bytes, pages}；r.log.errors 是解析过的真实错误
  （文件:行号: 说明），完整日志在 r.log.path；r.log.overfull 是排版提示不一定是错；
  r.support 说明模板文件从哪来；失败先看 r.log.errors 第一条与 r.diagnosis。
  office.tex.engines() 看本机有什么（latexmk / xelatex / lualatex / bibtex）；
  office.tex.clearAux(path, 'aux'|'all') 清中间产物。本机没有 TeX 发行版时
  create / read / edit 照常可用，只有 compile 会明确报错。

注意：本科生（bachelor）不需要 committee；关键词不超过 5 个；\\thusetup{} 里不要出现空行；
字体行距页边距题注格式都由模板负责，不要自己写 \\vspace / \\fontsize 调版式。
完整示例、项目结构逐文件说明、常见编译错误对照：office_help({ topic:'tex', detail:true })。`;

export const meta = {
    id: 'tex',
    name: 'LaTeX 学位论文（thuthesis）',
    ext: '.tex',
    summary: '生成/读取/编辑清华大学学位论文的 LaTeX 源文件（thuthesis 模板：主文件 + thusetup + data/ 分章 + refs.bib + 模板文件），并可调用本机 TeX Live 编出 PDF',
    guide: GUIDE,
    guideBrief: GUIDE_BRIEF,
    methods: {
        create: [
            'create({path, title, titleEn, degree, degreeType, language, fontset, author, supervisor, department, discipline, degreeCategory, date, abstract, chapters, references, acknowledgements, resume, committee, denotation, appendix, bibStyle, templateDir, support}) → report',
            '  degree: doctor | master | bachelor | postdoc；degreeType: academic | professional',
            '  chapters: [{title, label?, blocks?, sections: [{title, label?, blocks?, subsections}]}]',
            '  abstract: {zh: [段…], en: [段…], keywordsZh: [≤5], keywordsEn: [≤5]}',
            "  references: [{key, type, author, title, journal, year, volume, pages, …}] 或 '@article{key, …}' 字符串",
            '  一次写出整套项目（含模板运行期文件）；报告里有 files / support / stats / warnings',
            '  support: auto（默认，找到模板就整套拷贝）| system（用 TeX Live 自带的）| none',
        ],
        read: [
            "read(path, env) → report：顺 \\input 递归解析整个项目（片段文件只报自己）",
            '  stats: files / chapters / sections / figures / tables / equations / citations / references / words / degree / language / fontset',
            '  outline: 学位与规模摘要 + 每章每节标题',
            '  warnings: 图片与 \\input 指向的文件是否存在、\\cite 的 key 有没有对应 bib 条目、图表有没有题注与 label、\\ref 的 label 是否存在、段落是否过长、摘要与致谢是否缺失',
        ],
        edit: [
            'edit(path, ops, env) → report + {applied, skipped, changes}，一次调用改完所有位置',
            "  [{find, replace, all=true, file?}]        跨项目字面替换（file 限定单个文件）",
            "  [{set: {title: '…', date: '…'}, file?}]   改写 \\thusetup 字段（默认改 thusetup.tex）",
            '  [{appendTo: {file, text}}]                往某个文件末尾追加 LaTeX',
            '  [{appendChapter: {title, sections}}]      新建 data/chapNN.tex 并插入 \\input',
            '  [{addReference: {key, type, author, title, year, …}}]  追加一条 BibTeX（同 key 不重复）',
        ],
        api: [
            "await office.tex.compile(path, {timeoutMs, clean:'aux'|'all'|'none', engine, support}) → {ok, pdf, log:{errors, tail, path, overfull}, diagnosis}",
            'await office.tex.engines() → {available:[latexmk|xelatex|lualatex], paths, bibliography, hint}',
            "office.tex.template(dir?) → {dir, source, complete} 模板目录从哪来",
            'office.tex.escape(text) → 把普通文本里的 % & _ # $ 转义成 LaTeX 能编译的形式',
            "await office.tex.clearAux(path, 'aux'|'all') → 清中间产物",
        ],
    },
};

export { escapeLatex, GUIDE as TEX_GUIDE };
