/**
 * LaTeX 子集 → Office Math（OMML）转换器：PPT 公式功能的纯函数内核。
 *
 * 为什么不用图片方案：图片公式在 PowerPoint 里不能再编辑、不能跟着主题换色，
 * 而原生 OMML（DrawingML 2010 的 a14:m 包装）在 PowerPoint 2010+ / WPS / LibreOffice
 * 里都是可编辑公式。本模块只产出「m:oMath 的内部内容」与「纯文本近似」，
 * AlternateContent / mc:Fallback 的包装由 ppt.js 做 —— 那一层需要 run 的字号与颜色。
 *
 * 支持的子集（写幻灯片公式足够；超出子集会带着字符位置报错，由调用方回退成文字）：
 * - 字母/数字/运算符；希腊字母与常用符号（\times \leq \in \to \forall …共 178 个）
 * - \frac \dfrac \tfrac \binom、\sqrt（含 [n] 次根）、上下标与撇号（x'、x''）
 * - \sum \prod \int 等大运算符带上下限（\limits / \nolimits 可改位置）、
 *   \lim \max \min 等带下限的极限函数
 * - \sin \cos \log \ln 等函数名（自动正体）、\text \mathrm \mathbf \mathit \mathcal \mathfrak \mathbb
 * - \left…\right 定界符（含 \lfloor \langle 与空定界符 .）、\hat \bar \vec \dot \tilde 等重音、
 *   \overline \underline、\overrightarrow、\overbrace \underbrace
 *   —— 注意 \langle 这类定界符名**只在 \left…\right 里可用**，单独写 \langle 会报「不认识命令」；
 *   目录由本模块的 `knownCommands()` 现取（每一处都标了它能在什么语境下用）
 * - 矩阵族 matrix/pmatrix/bmatrix/vmatrix/Bmatrix、cases、aligned/align/gather
 *   （& 分列、\\\\ 换行；aligned 的对齐标记退化为按内容排列）
 * - 间距 \\, \\: \\; \\quad \\qquad \\! \\  与 ~、注释 %、换行 \\\\（展示式里换行 = 多条 oMath）
 *
 * 纯文本近似（plain）服务于两处：mc:Fallback 里给不支持 a14 的阅读器看的文字，
 * 以及 ppt.js 的字数统计与折行估算。公式没有「字数」概念，但量算必须算进去，
 * 否则「带公式的要点」会按空段落估高。
 *
 * @module dsh-office-mode/formats/ppt-math
 */

/** XML 文本转义（m:t 的内容只需要这三个）。 */
function esc(value) {
    return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 上标/下标 Unicode 映射：能完整映射才用 Unicode 角标，否则退回 ^(..) / _(..)。 */
const SUP_MAP = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸',
    '9': '⁹', '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾', 'n': 'ⁿ', 'i': 'ⁱ',
};
const SUB_MAP = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈',
    '9': '₉', '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎', 'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ',
    'i': 'ᵢ', 'j': 'ⱼ', 'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ', 'p': 'ₚ', 'r': 'ᵣ',
    's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ', 'v': 'ᵥ', 'x': 'ₓ',
};

/** 希腊字母（小写 + 常用变体 + 大写）。 */
const GREEK = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε', zeta: 'ζ',
    eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν',
    xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ', varrho: 'ϱ', sigma: 'σ', varsigma: 'ς', tau: 'τ',
    upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ',
    Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

/** 符号命令（二元/关系/箭头/杂项），间距由数学引擎自动摆。 */
const SYMBOLS = {
    times: '×', div: '÷', cdot: '⋅', bullet: '•', ast: '∗', star: '⋆', circ: '∘',
    oplus: '⊕', ominus: '⊖', otimes: '⊗', oslash: '⊘', odot: '⊙', pm: '±', mp: '∓',
    cup: '∪', cap: '∩', uplus: '⊎', sqcap: '⊓', sqcup: '⊔', vee: '∨', wedge: '∧',
    setminus: '∖', wr: '≀', diamond: '⋄', bigtriangleup: '△', bigtriangledown: '▽',
    triangleleft: '⊲', triangleright: '⊳',
    leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', equiv: '≡', approx: '≈',
    cong: '≅', simeq: '≃', sim: '∼', asymp: '≍', doteq: '≐', propto: '∝',
    prec: '≺', succ: '≻', preceq: '⪯', succeq: '⪰', ll: '≪', gg: '≫',
    subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇', nsubseteq: '⊈', nsupseteq: '⊉',
    in: '∈', notin: '∉', ni: '∋', models: '⊨', vdash: '⊢', dashv: '⊣',
    perp: '⊥', parallel: '∥', angle: '∠', measuredangle: '∡', triangle: '△',
    to: '→', rightarrow: '→', leftarrow: '←', leftrightarrow: '↔',
    Rightarrow: '⇒', Leftarrow: '⇐', Leftrightarrow: '⇔',
    longrightarrow: '⟶', longleftarrow: '⟵', longleftrightarrow: '⟷',
    mapsto: '↦', longmapsto: '⟼', hookrightarrow: '↪', hookleftarrow: '↩',
    uparrow: '↑', downarrow: '↓', updownarrow: '↕', Uparrow: '⇑', Downarrow: '⇓', Updownarrow: '⇕',
    nearrow: '↗', searrow: '↘', swarrow: '↙', nwarrow: '↖',
    rightharpoonup: '⇀', leftharpoonup: '↼', rightleftharpoons: '⇌',
    implies: '⟹', impliedby: '⟸', iff: '⟺',
    neg: '¬', lnot: '¬', land: '∧', lor: '∨',
    forall: '∀', exists: '∃', nexists: '∄', top: '⊤', bot: '⊥',
    infty: '∞', partial: '∂', nabla: '∇', emptyset: '∅', varnothing: '∅',
    hbar: 'ℏ', ell: 'ℓ', imath: 'ı', jmath: 'ȷ', aleph: 'ℵ', wp: '℘', Re: 'ℜ', Im: 'ℑ',
    prime: '′', degree: '°', because: '∵', therefore: '∴',
    square: '□', blacksquare: '■', checkmark: '✓', dagger: '†', ddagger: '‡',
    vdots: '⋮', cdots: '⋯', ddots: '⋱', ldots: '…', dots: '…',
    '%': '%', '&': '&', '#': '#', '_': '_', '{': '{', '}': '}', '$': '$',
};

/** 大运算符：side=true 时上下限在侧边（积分族），否则在上下（求和/乘积族）。 */
const BIG_OPS = {
    sum: { chr: '∑', side: false }, prod: { chr: '∏', side: false }, coprod: { chr: '∐', side: false },
    int: { chr: '∫', side: true }, iint: { chr: '∬', side: true }, iiint: { chr: '∭', side: true },
    oint: { chr: '∮', side: true }, oiint: { chr: '∯', side: true }, oiiint: { chr: '∰', side: true },
    bigcup: { chr: '⋃', side: false }, bigcap: { chr: '⋂', side: false }, bigsqcup: { chr: '⨆', side: false },
    biguplus: { chr: '⨄', side: false }, bigoplus: { chr: '⨁', side: false }, bigotimes: { chr: '⨂', side: false },
    bigvee: { chr: '⋁', side: false }, bigwedge: { chr: '⋀', side: false },
};

/** 极限类：\lim_{x \to 0} 排成「下方极限」（m:limLow）；无下限时就是正体函数名。 */
const LIMIT_FUNCS = new Set(['lim', 'limsup', 'liminf', 'max', 'min', 'sup', 'inf', 'gcd', 'Pr', 'argmax', 'argmin']);

/** 函数名：\sin x 排成 m:func（名称正体 + 参数）；无参数时只出名称。 */
const FUNCS = new Set([
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'arccot',
    'sinh', 'cosh', 'tanh', 'coth', 'log', 'ln', 'lg', 'exp', 'deg', 'dim', 'hom', 'ker', 'arg', 'mod', 'det',
]);

/** 重音命令 → 组合字符（m:acc 的 m:chr）。 */
const ACCENTS = {
    hat: '̂', widehat: '̂', tilde: '̃', widetilde: '̃', bar: '̄', vec: '⃗',
    dot: '̇', ddot: '̈', acute: '́', grave: '̀', check: '̌', breve: '̆', mathring: '̊',
};

/** \overrightarrow 族 → m:groupChr（字符 + 位置）。 */
const GROUP_CHARS = {
    overrightarrow: { chr: '→', pos: 'top' },
    overleftarrow: { chr: '←', pos: 'top' },
    overleftrightarrow: { chr: '↔', pos: 'top' },
    overbrace: { chr: '⏞', pos: 'top' },
    underbrace: { chr: '⏟', pos: 'bot' },
};

/** 空白命令 → 空格类字符（\! 负空格没有对应物，直接吞掉）。 */
const SPACES = {
    ',': ' ', ':': ' ', ';': ' ', ' ': ' ', '~': ' ', quad: ' ', qquad: '  ', enspace: ' ',
};

/** \mathbb 的双线体映射（没映射到的字母退回原字母）。 */
const BB_MAP = {
    A: '𝔸', B: '𝔹', C: 'ℂ', D: '𝔻', E: '𝔼', F: '𝔽', G: '𝔾', H: 'ℍ', I: '𝕀', J: '𝕁',
    K: '𝕂', L: '𝕃', M: '𝕄', N: 'ℕ', O: '𝕆', P: 'ℙ', Q: 'ℚ', R: 'ℝ', S: '𝕊', T: '𝕋',
    U: '𝕌', V: '𝕍', W: '𝕎', X: '𝕏', Y: '𝕐', Z: 'ℤ',
};

/** 文字样式命令 → m:rPr 的 m:sty 值；这类命令的参数按「原样文字」读（保留空格）。 */
const TEXT_STYLES = { text: 'p', mathrm: 'p', mathbf: 'b', mathit: 'i', mathcal: 'scr', mathfrak: 'frk', mathbb: 'p' };

/** \left / \right 接受的定界符 → 单字符。 */
const DELIMS = {
    '(': '(', ')': ')', '[': '[', ']': ']', '{': '{', '}': '}', '|': '|', '/': '/', '\\': '\\',
    lfloor: '⌊', rfloor: '⌋', lceil: '⌈', rceil: '⌉', langle: '⟨', rangle: '⟩',
    lvert: '|', rvert: '|', lVert: '‖', rVert: '‖', uparrow: '↑', downarrow: '↓',
    updownarrow: '↕', Uparrow: '⇑', Downarrow: '⇓', Updownarrow: '⇕', lgroup: '⟮', rgroup: '⟯',
};

/** 矩阵族的定界符（空串 = 无定界符）。 */
const MATRIX_DELIMS = {
    matrix: ['', ''], pmatrix: ['(', ')'], bmatrix: ['[', ']'], Bmatrix: ['{', '}'],
    vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'],
};

/** plain 近似里需要两侧留空白的字符。 */
const SPACED_OPS = new Set([
    '=', '+', '−', '×', '÷', '±', '∓', '≤', '≥', '≠', '≈', '≡', '→', '←', '⇒', '⇔', '↦',
    '∈', '∉', '⊂', '⊆', '∪', '∩', '⋅', '∝', '∼', '≅', '⊕', '⊗', '⟶', '⟵',
]);

/** \text 类命令里 `\x` 的字面量还原。 */
const TEXT_ESCAPES = { '\\': '\\', '%': '%', '&': '&', '#': '#', '_': '_', '{': '{', '}': '}', '$': '$' };

class MathError extends Error {}

function err(pos, message) {
    throw new MathError(`第 ${pos} 字符附近：${message}`);
}

/**
 * 命令目录：由上面那些表 + 只由 switch 分支识别的命令现算出来。
 *
 * 为什么要它：文档里那份「支持的 LaTeX 子集」是手写的，与表没有任何生成或校验关系，
 * 已经漂移过（写「约 150 个常用符号」，实际 178 个）。有了这个目录，
 * `office_help` 的清单与测试断言都能从真源现取，不会各说各话。
 *
 * `where` 标出每个命令**能在什么语境下用** —— 表里有名字不等于能独立使用：
 * `delims` 里的 27 个（\langle \lfloor \rVert …）只在 `\left…\right` 内可达，
 * `nary` 里的 \limits / \nolimits 只紧接大运算符时可达。
 */
const SWITCH_ONLY_COMMANDS = ['frac', 'dfrac', 'tfrac', 'binom', 'sqrt', 'left', 'right',
    'begin', 'end', 'overline', 'underline', '!', 'limits', 'nolimits'];

function commandCatalogue() {
    const table = new Map();
    const add = (names, where) => {
        for (const name of names) {
            if (name === '') continue;
            const current = table.get(name);
            // 同一个名字若既能独立用、又能在 \left…\right 里用，按更强的那个记
            if (current === undefined || where === 'standalone') table.set(name, where);
        }
    };
    add(Object.keys(GREEK), 'standalone');
    add(Object.keys(SYMBOLS), 'standalone');
    add(Object.keys(BIG_OPS), 'standalone');
    add([...LIMIT_FUNCS], 'standalone');
    add([...FUNCS], 'standalone');
    add(Object.keys(ACCENTS), 'standalone');
    add(Object.keys(GROUP_CHARS), 'standalone');
    add(Object.keys(SPACES).filter((name) => name.length > 1), 'standalone');
    add(Object.keys(TEXT_STYLES), 'standalone');
    add(Object.keys(MATRIX_DELIMS), 'environment');
    add(Object.keys(DELIMS).filter((name) => name.length > 1), 'only-in-left-right');
    for (const name of SWITCH_ONLY_COMMANDS) {
        if (name === 'limits' || name === 'nolimits') table.set(name, 'nary-modifier');
        else if (name === 'left' || name === 'right' || name === 'begin' || name === 'end') table.set(name, 'standalone');
        else add([name], 'standalone');
    }
    return table;
}

const COMMAND_CATALOGUE = commandCatalogue();

/** 命令是否在支持的子集里（不带反斜杠的名字）。 */
export function isKnownCommand(name) {
    return COMMAND_CATALOGUE.has(String(name ?? ''));
}

/**
 * 支持域清单（给文档 / 测试用）：按名字排序，每条带 `where`。
 * `where`：standalone（可独立使用）/ only-in-left-right（只在 \left…\right 里）/
 * nary-modifier（只紧接大运算符）/ environment（\begin{} 的环境名）。
 */
export function knownCommands() {
    return [...COMMAND_CATALOGUE.entries()]
        .map(([name, where]) => ({ name, where }))
        .sort((left, right) => (left.name < right.name ? -1 : (left.name > right.name ? 1 : 0)));
}

/**
 * 一段**普通文本**是不是「本该是公式、却按字面量留下来了」的 LaTeX。
 *
 * 用途：`read()` 侧的兜底扫描（公式在写入时解析失败会回落成普通 run，原文进 a:t / w:t，
 * 之后再读就只是普通文本）。判据刻意保守 —— 误报比漏报更烦人：
 *   1. 文本里至少要有一个 `\命令`（否则连像公式都谈不上）；
 *   2. 已知命令 ≥ 2 个；或出现 `{}`（路径里几乎不会有的数学写法）；
 *   3. 只有一个已知命令时，还要有 `[]` `_` `^` 里的一种。
 * 于是 Windows 路径（`D:\work\temp`、`C:\Users\alpha`、`D:\my_files\x`）不会命中，
 * 而 `\langle x \rangle + \epsilon`、`\frac{a}{b}`、`\unknowncmd{x}`、`E=mc^2 \cdot x` 会命中
 * —— 最后这条要紧：解析失败的公式**必然**含未知命令，只按支持域判就永远抓不到它。
 */
export function looksLikeLatex(text) {
    const source = String(text ?? '');
    if (!source.includes('\\')) return false;
    const names = source.match(/\\[a-zA-Z]+/g) ?? [];
    if (names.length === 0) return false;
    const known = names.filter((item) => COMMAND_CATALOGUE.has(item.slice(1)));
    if (known.length >= 2) return true;
    if (/[{}]/.test(source)) return true;
    return known.length >= 1 && /[[\]_^]/.test(source);
}

// ─── 词法 ───────────────────────────────────────────────────────────────────

/**
 * 词法：把 latex 源码切成 token。
 * kind: char（普通字符）| cmd（\name，含单字符命令）|lbrace|rbrace|sup|sub|amp|nl|space
 * 被词法吞掉的原始空格以 space token 保留 —— \text 的原样文字要靠它还原。
 */
function tokenize(input) {
    const tokens = [];
    let i = 0;
    const n = input.length;
    while (i < n) {
        const c = input[i];
        if (c === '%') {
            while (i < n && input[i] !== '\n') i += 1;
            continue;
        }
        if (c === '\n' || c === '\r' || c === '\t' || c === ' ') {
            let j = i;
            while (j < n && /[\n\r\t ]/.test(input[j])) j += 1;
            tokens.push({ kind: 'space', pos: i + 1 });
            i = j;
            continue;
        }
        if (c === '\\') {
            const next = input[i + 1];
            if (next === '\\') {
                tokens.push({ kind: 'nl', pos: i + 1 });
                i += 2;
                continue;
            }
            if (next !== undefined && /[a-zA-Z]/.test(next)) {
                let j = i + 1;
                while (j < n && /[a-zA-Z]/.test(input[j])) j += 1;
                tokens.push({ kind: 'cmd', name: input.slice(i + 1, j), pos: i + 1 });
                i = j;
                continue;
            }
            tokens.push({ kind: 'cmd', name: next ?? '', pos: i + 1 });
            i += 2;
            continue;
        }
        if (c === '{') { tokens.push({ kind: 'lbrace', pos: i + 1 }); i += 1; continue; }
        if (c === '}') { tokens.push({ kind: 'rbrace', pos: i + 1 }); i += 1; continue; }
        if (c === '^') { tokens.push({ kind: 'sup', pos: i + 1 }); i += 1; continue; }
        if (c === '_') { tokens.push({ kind: 'sub', pos: i + 1 }); i += 1; continue; }
        if (c === '&') { tokens.push({ kind: 'amp', pos: i + 1 }); i += 1; continue; }
        if (c === '~') { tokens.push({ kind: 'cmd', name: '~', pos: i + 1 }); i += 1; continue; }
        tokens.push({ kind: 'char', ch: c, pos: i + 1 });
        i += 1;
    }
    return tokens;
}

// ─── 语法 ───────────────────────────────────────────────────────────────────

/**
 * 递归下降解析器，一次产出 omml（m:oMath 内部内容）与 plain（纯文本近似）。
 * 节点为 { o, p }，只拼字符串、不建树。
 */
function createParser(tokens, rPr) {
    let at = 0;

    const peek = (offset = 0) => tokens[at + offset];
    const next = () => tokens[at++];

    /** 跳过空白 token；数学模式下空白不改变语义（间距由渲染引擎摆）。 */
    function skipSpace() {
        while (peek() !== undefined && peek().kind === 'space') at += 1;
    }

    /** 一条 m:r；sty 缺省 = 默认样式（变量斜体/数字正体由渲染器决定）。 */
    function run(text, sty) {
        const styPr = sty ? `<m:rPr><m:sty m:val="${sty}"/></m:rPr>` : '';
        return { o: `<m:r>${styPr}${rPr}<m:t>${esc(text)}</m:t></m:r>`, p: text };
    }

    function join(nodes) {
        return { o: nodes.map((node) => node.o).join(''), p: nodes.map((node) => node.p).join('') };
    }

    /** 读一个参数：{…} 组或单个 token；返回节点。 */
    function parseArg(what) {
        skipSpace();
        const token = peek();
        if (token === undefined) err(0, `${what} 缺少参数`);
        if (token.kind === 'lbrace') {
            next();
            const nodes = parseList((t) => t.kind === 'rbrace');
            const close = next();
            if (close === undefined || close.kind !== 'rbrace') err(token.pos, `${what} 的 { 没有配对的 }`);
            return join(nodes);
        }
        if (['rbrace', 'sup', 'sub', 'amp', 'nl'].includes(token.kind)) {
            err(token.pos, `${what} 的参数不合法（${token.kind}）`);
        }
        return parseElement();
    }

    /**
     * 读「原样文字组」（\text{...} 一族）：从 token 流还原字面文本，保留空格，
     * 只还原 \% \& \_ \{ \} \\ \# \$ 这类转义；组内允许嵌套 {}（按字面处理）。
     */
    function readRawGroup(what) {
        skipSpace();
        const open = peek();
        if (open === undefined || open.kind !== 'lbrace') {
            // 无组：按单 token 处理（\mathrm x 这类写法）
            const arg = parseArg(what);
            return arg.p;
        }
        next();
        let depth = 1;
        let text = '';
        for (;;) {
            const token = next();
            if (token === undefined) err(open.pos, `${what} 的 { 没有配对的 }`);
            if (token.kind === 'lbrace') { depth += 1; text += '{'; continue; }
            if (token.kind === 'rbrace') {
                depth -= 1;
                if (depth === 0) break;
                text += '}';
                continue;
            }
            if (token.kind === 'space') { text += ' '; continue; }
            if (token.kind === 'cmd') {
                if (TEXT_ESCAPES[token.name] !== undefined) { text += TEXT_ESCAPES[token.name]; continue; }
                if (token.name === '') err(token.pos, `孤立的反斜杠`);
                text += `\\${token.name}`;
                continue;
            }
            if (token.kind === 'nl') { text += ' '; continue; }
            if (token.kind === 'char') { text += token.ch; continue; }
            // sup/sub/amp 在 \text 里按字面算（少见，但比报错好）
            text += token.kind === 'sup' ? '^' : token.kind === 'sub' ? '_' : token.kind === 'amp' ? '&' : '';
        }
        return text;
    }

    /**
     * 读一个「不偷角标」的参数：给 \sum_…^… / \log_… 这类场景用 ——
     * 单 token 参数不会再把紧随其后的 ^ / _ 吞成自己的角标
     * （否则 \int_0^{2\pi} 会变成「下标是 0^2π」）。
     */
    function parseArgNoScript(what) {
        skipSpace();
        const token = peek();
        if (token === undefined) err(0, `${what} 缺少参数`);
        if (token.kind === 'lbrace') return parseArg(what);
        if (['rbrace', 'sup', 'sub', 'amp', 'nl'].includes(token.kind)) {
            err(token.pos, `${what} 的参数不合法（${token.kind}）`);
        }
        next();
        return token.kind === 'char' ? charNode(token) : cmdNode(token);
    }

    /** 角标 → Unicode（能完整映射才用），否则 ^(..) / _(..)。 */
    function mapScript(text, map, mark) {
        const chars = [...text];
        if (chars.length > 0 && chars.every((ch) => map[ch] !== undefined)) {
            return chars.map((ch) => map[ch]).join('');
        }
        return `${mark}(${text})`;
    }

    const toSup = (text) => mapScript(text, SUP_MAP, '^');
    const toSub = (text) => mapScript(text, SUB_MAP, '_');

    function wrapScript(base, sub, sup) {
        if (sup !== null && sub !== null) {
            const o = `<m:sSubSup><m:e>${base.o}</m:e><m:sub>${sub.o}</m:sub><m:sup>${sup.o}</m:sup></m:sSubSup>`;
            return { o, p: `${base.p}${toSub(sub.p)}${toSup(sup.p)}` };
        }
        if (sup !== null) {
            const o = `<m:sSup><m:e>${base.o}</m:e><m:sup>${sup.o}</m:sup></m:sSup>`;
            return { o, p: `${base.p}${toSup(sup.p)}` };
        }
        const o = `<m:sSub><m:e>${base.o}</m:e><m:sub>${sub.o}</m:sub></m:sSub>`;
        return { o, p: `${base.p}${toSub(sub.p)}` };
    }

    /** 把紧跟 base 的上/下标/撇号挂上去。 */
    function attachScripts(base) {
        let node = base;
        for (;;) {
            skipSpace();
            const token = peek();
            if (token === undefined) return node;
            if (token.kind === 'char' && token.ch === "'") {
                let count = 0;
                while (peek() !== undefined && peek().kind === 'char' && peek().ch === "'") { next(); count += 1; }
                const sup = run('′'.repeat(count), 'p');
                node = {
                    o: `<m:sSup><m:e>${node.o}</m:e><m:sup>${sup.o}</m:sup></m:sSup>`,
                    p: `${node.p}′`,
                };
                continue;
            }
            let sup = null;
            let sub = null;
            let saw = false;
            while (peek() !== undefined && (peek().kind === 'sup' || peek().kind === 'sub')) {
                const marker = next();
                saw = true;
                const arg = parseArg(marker.kind === 'sup' ? '^' : '_');
                if (marker.kind === 'sup') sup = sup === null ? arg : join([sup, arg]);
                else sub = sub === null ? arg : join([sub, arg]);
            }
            if (!saw) return node;
            node = wrapScript(node, sub, sup);
        }
    }

    /** 元素级解析：一个原子 + 它自带的上下标。 */
    function parseElement() {
        const token = next();
        if (token === undefined) err(0, '公式意外结束');
        if (token.kind === 'lbrace') {
            const nodes = parseList((t) => t.kind === 'rbrace');
            const close = next();
            if (close === undefined || close.kind !== 'rbrace') err(token.pos, '{ 没有配对的 }');
            return attachScripts(join(nodes));
        }
        if (token.kind === 'rbrace') err(token.pos, '出现多余的 }');
        if (token.kind === 'sup' || token.kind === 'sub') err(token.pos, `${token.kind === 'sup' ? '^' : '_'} 前面没有内容`);
        if (token.kind === 'amp') err(token.pos, '& 只能出现在矩阵 / aligned 环境');
        if (token.kind === 'nl') err(token.pos, '\\\\ 只能用在环境或多行公式');
        if (token.kind === 'space') return parseElement();
        if (token.kind === 'char') return attachScripts(charNode(token));
        return attachScripts(cmdNode(token));
    }

    /** 普通字符：同类合并成一条 run（字母一条、数字一条、运算符单字符一条）。 */
    function charNode(token) {
        const ch = token.ch;
        if (/[a-zA-Z]/.test(ch)) {
            let word = ch;
            while (peek() !== undefined && peek().kind === 'char' && /[a-zA-Z]/.test(peek().ch)) {
                word += next().ch;
            }
            return run(word);
        }
        if (/[0-9.]/.test(ch)) {
            let num = ch;
            while (peek() !== undefined && peek().kind === 'char' && /[0-9.]/.test(peek().ch)) {
                num += next().ch;
            }
            return run(num, 'p');
        }
        if (ch === "'") err(token.pos, '撇号前面没有内容');
        const plain = SPACED_OPS.has(ch) || ch === '-' ? ` ${ch} ` : ch;
        return { o: run(ch, 'p').o, p: plain };
    }

    /** 命令解析（子集主体）。 */
    function cmdNode(token) {
        const name = token.name;
        const pos = token.pos;
        if (name === 'frac' || name === 'dfrac' || name === 'tfrac') {
            const num = parseArg(`\\${name}`);
            const den = parseArg(`\\${name}`);
            const o = `<m:f><m:num>${num.o}</m:num><m:den>${den.o}</m:den></m:f>`;
            return { o, p: `${wrapParen(num.p)}/${wrapParen(den.p)}` };
        }
        if (name === 'binom') {
            const a = parseArg('\\binom');
            const b = parseArg('\\binom');
            const o = `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${a.o}</m:num><m:den>${b.o}</m:den></m:f>`;
            return { o, p: `(${a.p};${b.p})` };
        }
        if (name === 'sqrt') {
            let deg = { o: '', p: '' };
            skipSpace();
            if (peek() !== undefined && peek().kind === 'char' && peek().ch === '[') {
                next();
                const nodes = parseList((t) => t.kind === 'char' && t.ch === ']');
                const close = next();
                if (close === undefined || close.kind !== 'char' || close.ch !== ']') {
                    err(pos, '\\sqrt 的 [ 没有配对的 ]');
                }
                deg = join(nodes);
            }
            const body = parseArg('\\sqrt');
            const radPr = deg.o === '' ? '<m:radPr><m:degHide m:val="1"/></m:radPr>' : '';
            const o = `<m:rad>${radPr}<m:deg>${deg.o}</m:deg><m:e>${body.o}</m:e></m:rad>`;
            return { o, p: `${deg.p}√${wrapParen(body.p)}` };
        }
        if (BIG_OPS[name] !== undefined) {
            const spec = BIG_OPS[name];
            let side = spec.side;
            let sub = null;
            let sup = null;
            for (;;) {
                skipSpace();
                const marker = peek();
                if (marker !== undefined && marker.kind === 'cmd' && (marker.name === 'limits' || marker.name === 'nolimits')) {
                    next();
                    side = marker.name === 'nolimits';
                    continue;
                }
                if (marker !== undefined && (marker.kind === 'sub' || marker.kind === 'sup')) {
                    next();
                    const arg = parseArgNoScript(marker.kind === 'sub' ? '_' : '^');
                    if (marker.kind === 'sub') sub = arg;
                    else sup = arg;
                    continue;
                }
                break;
            }
            const hideSub = sub === null ? '<m:subHide m:val="1"/>' : '';
            const hideSup = sup === null ? '<m:supHide m:val="1"/>' : '';
            // 空的 m:e 是合法的（Word 对裸 \sum 也写空 e）：运算对象作为兄弟节点排在后面
            const o = `<m:nary><m:naryPr><m:chr m:val="${spec.chr}"/>`
                + `<m:limLoc m:val="${side ? 'subSup' : 'undOvr'}"/>${hideSub}${hideSup}</m:naryPr>`
                + `<m:sub>${sub === null ? '' : sub.o}</m:sub><m:sup>${sup === null ? '' : sup.o}</m:sup>`
                + `<m:e/></m:nary>`;
            return { o, p: `${spec.chr}${sub === null ? '' : toSub(sub.p)}${sup === null ? '' : toSup(sup.p)} ` };
        }
        if (LIMIT_FUNCS.has(name)) {
            skipSpace();
            let sub = null;
            if (peek() !== undefined && peek().kind === 'sub') {
                next();
                sub = parseArgNoScript('_');
            }
            const nameNode = run(name, 'p');
            if (sub === null) return nameNode;
            const o = `<m:limLow><m:e><m:func><m:fName>${nameNode.o}</m:fName><m:e/></m:func></m:e>`
                + `<m:lim>${sub.o}</m:lim></m:limLow>`;
            return { o, p: `${name}${toSub(sub.p)} ` };
        }
        if (FUNCS.has(name)) {
            const nameNode = run(name, 'p');
            skipSpace();
            const marker = peek();
            if (marker === undefined || marker.kind === 'rbrace' || marker.kind === 'nl' || marker.kind === 'amp') {
                return nameNode;
            }
            // \log_2 8：函数名自带角标（底数）—— 排成 sSub/sSup/sSubSup 包住 m:func
            if (marker.kind === 'sub' || marker.kind === 'sup') {
                const funcNode = { o: `<m:func><m:fName>${nameNode.o}</m:fName><m:e/></m:func>`, p: name };
                return attachScripts(funcNode);
            }
            const arg = parseArg(`\\${name}`);
            const o = `<m:func><m:fName>${nameNode.o}</m:fName><m:e>${arg.o}</m:e></m:func>`;
            return { o, p: `${name} ${arg.p}` };
        }
        if (TEXT_STYLES[name] !== undefined) {
            const text = readRawGroup(`\\${name}`);
            if (name === 'mathbb') {
                const mapped = [...text].map((ch) => BB_MAP[ch] ?? ch).join('');
                return run(mapped, 'p');
            }
            return run(text, TEXT_STYLES[name]);
        }
        if (name === 'left') {
            skipSpace();
            const openToken = next();
            if (openToken === undefined) err(pos, '\\left 后面缺少定界符');
            let begChr = '';
            if (openToken.kind === 'char' && openToken.ch === '.') begChr = '';
            else if (openToken.kind === 'char' && DELIMS[openToken.ch] !== undefined) begChr = DELIMS[openToken.ch];
            else if (openToken.kind === 'cmd' && DELIMS[openToken.name] !== undefined) begChr = DELIMS[openToken.name];
            else err(openToken.pos, `\\left 不认识定界符`);
            const inner = parseList((t) => t.kind === 'cmd' && t.name === 'right');
            const rightToken = next();
            if (rightToken === undefined || rightToken.kind !== 'cmd' || rightToken.name !== 'right') {
                err(pos, '\\left 没有配对的 \\right');
            }
            skipSpace();
            const closeToken = next();
            if (closeToken === undefined) err(rightToken.pos, '\\right 后面缺少定界符');
            let endChr = '';
            if (closeToken.kind === 'char' && closeToken.ch === '.') endChr = '';
            else if (closeToken.kind === 'char' && DELIMS[closeToken.ch] !== undefined) endChr = DELIMS[closeToken.ch];
            else if (closeToken.kind === 'cmd' && DELIMS[closeToken.name] !== undefined) endChr = DELIMS[closeToken.name];
            else err(closeToken.pos, '\\right 不认识定界符');
            const o = `<m:d><m:dPr><m:begChr m:val="${begChr}"/>${endChr === '' ? '<m:endChr m:val=""/>' : `<m:endChr m:val="${endChr}"/>`}</m:dPr>`
                + `<m:e>${inner.map((node) => node.o).join('')}</m:e></m:d>`;
            const p = `${begChr}${inner.map((node) => node.p).join('')}${endChr}`;
            return { o, p };
        }
        if (name === 'right') err(pos, '\\right 出现时没有对应的 \\left');
        if (name === 'begin') return environment(token);
        if (name === 'end') err(pos, '\\end 出现时没有对应的 \\begin');
        if (ACCENTS[name] !== undefined) {
            const arg = parseArg(`\\${name}`);
            const o = `<m:acc><m:accPr><m:chr m:val="${ACCENTS[name]}"/></m:accPr><m:e>${arg.o}</m:e></m:acc>`;
            return { o, p: `${arg.p}${ACCENTS[name]}` };
        }
        if (name === 'overline' || name === 'underline') {
            const arg = parseArg(`\\${name}`);
            const posAttr = name === 'overline' ? 'top' : 'bot';
            const o = `<m:bar><m:barPr><m:pos m:val="${posAttr}"/></m:barPr><m:e>${arg.o}</m:e></m:bar>`;
            const mark = name === 'overline' ? '̅' : '̲';
            return { o, p: `${arg.p}${mark}` };
        }
        if (GROUP_CHARS[name] !== undefined) {
            const spec = GROUP_CHARS[name];
            const arg = parseArg(`\\${name}`);
            const o = `<m:groupChr><m:groupChrPr><m:chr m:val="${spec.chr}"/><m:pos m:val="${spec.pos}"/></m:groupChrPr>`
                + `<m:e>${arg.o}</m:e></m:groupChr>`;
            return { o, p: `${arg.p}${spec.chr}` };
        }
        if (SPACES[name] !== undefined) return run(SPACES[name], 'p');
        if (name === '!') return { o: '', p: '' };
        if (GREEK[name] !== undefined) {
            const glyph = GREEK[name];
            const plain = SPACED_OPS.has(glyph) ? ` ${glyph} ` : glyph;
            return { o: run(glyph).o, p: plain };
        }
        if (SYMBOLS[name] !== undefined) {
            const glyph = SYMBOLS[name];
            const plain = SPACED_OPS.has(glyph) ? ` ${glyph} ` : glyph;
            return { o: run(glyph, 'p').o, p: plain };
        }
        err(pos, `不认识命令 \\${name || '(空)'}（不在支持的 LaTeX 子集里）`);
        return undefined;
    }

    /** \begin{env}：读环境名 → 行（\\\\）/ 列（&）→ 矩阵族 / cases / aligned 族。 */
    function environment(beginToken) {
        skipSpace();
        const openBrace = next();
        if (openBrace === undefined || openBrace.kind !== 'lbrace') err(beginToken.pos, '\\begin 后面要有 {环境名}');
        let env = '';
        for (;;) {
            const token = peek();
            if (token === undefined) err(beginToken.pos, `\\begin 的 { 没有闭合`);
            if (token.kind === 'rbrace') { next(); break; }
            if (token.kind === 'char' && /[a-zA-Z*]/.test(token.ch)) { env += next().ch; continue; }
            err(token.pos, `\\begin 的环境名里有意外字符`);
        }
        if (env === '') err(beginToken.pos, '\\begin 的环境名为空');

        const rows = [];
        let row = [];
        let cell = [];
        const pushCell = () => { row.push(join(cell)); cell = []; };
        const pushRow = () => { pushCell(); rows.push(row); row = []; };
        for (;;) {
            skipSpace();
            const token = peek();
            if (token === undefined) err(beginToken.pos, `\\begin{${env}} 没有配对的 \\end{${env}}`);
            if (token.kind === 'cmd' && token.name === 'end') break;
            if (token.kind === 'nl') { next(); pushRow(); continue; }
            if (token.kind === 'amp') { next(); pushCell(); continue; }
            cell.push(parseElement());
        }
        pushRow();

        const endToken = next();
        skipSpace();
        const endOpen = next();
        if (endOpen === undefined || endOpen.kind !== 'lbrace') err(endToken.pos, '\\end 后面要有 {环境名}');
        let endName = '';
        for (;;) {
            const token = peek();
            if (token === undefined) err(endToken.pos, `\\end 的 { 没有闭合`);
            if (token.kind === 'rbrace') { next(); break; }
            if (token.kind === 'char' && /[a-zA-Z*]/.test(token.ch)) { endName += next().ch; continue; }
            err(token.pos, `\\end 的环境名里有意外字符`);
        }
        if (endName !== env) err(endToken.pos, `\\end{${endName}} 与 \\begin{${env}} 不匹配`);

        const columns = Math.max(1, ...rows.map((line) => line.length));
        const filled = rows.map((line) => {
            const cells = [...line];
            while (cells.length < columns) cells.push({ o: '', p: '' });
            return cells;
        });
        const plainRows = filled.map((line) => line.map((node) => node.p).join('，'));
        // aligned 族的 & 是对齐标记不是内容：plain 里空格连接即可，别把「a， = 1」这种逗号带出来
        const plainAlignRows = filled.map((line) => line.map((node) => node.p).filter((text) => text !== '').join(' '));

        if (MATRIX_DELIMS[env] !== undefined) {
            const [beg, end] = MATRIX_DELIMS[env];
            const mcPr = Array.from({ length: columns }, () => '<m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="center"/></m:mcPr></m:mc>').join('');
            const mrs = filled.map((line) => `<m:mr>${line.map((node) => `<m:e>${node.o}</m:e>`).join('')}</m:mr>`).join('');
            const matrix = `<m:m><m:mPr><m:mcs>${mcPr}</m:mcs></m:mPr>${mrs}</m:m>`;
            if (beg === '' && end === '') return { o: matrix, p: plainRows.join('；') };
            // 带定界符的矩阵：m:m 整体放进 m:d（Word 的 pmatrix 就是这个结构）
            const o = `<m:d><m:dPr><m:begChr m:val="${beg}"/><m:endChr m:val="${end}"/></m:dPr>`
                + `<m:e>${matrix}</m:e></m:d>`;
            return { o, p: `${beg}${plainRows.join('；')}${end}` };
        }
        if (env === 'cases') {
            const es = filled.map((line) => `<m:e>${line.map((node) => node.o).join('')}</m:e>`).join('');
            const o = `<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr>`
                + `<m:e><m:eqArr>${es}</m:eqArr></m:e></m:d>`;
            return { o, p: `{${plainAlignRows.join('；')}` };
        }
        if (['aligned', 'align', 'align*', 'gather', 'gathered', 'eqnarray'].includes(env)) {
            const es = filled.map((line) => `<m:e>${line.map((node) => node.o).join('')}</m:e>`).join('');
            const o = `<m:eqArr>${es}</m:eqArr>`;
            return { o, p: plainAlignRows.join('；') };
        }
        err(beginToken.pos, `不支持环境 ${env}（可用 matrix/pmatrix/bmatrix/vmatrix/Bmatrix/cases/aligned）`);
        return undefined;
    }

    /**
     * 序列解析：直到 stop(token) 为真（停止 token 由调用方消费）。
     * 组内的 \\\\ 按空格容忍（与 LaTeX 一致）。
     */
    function parseList(stop) {
        const nodes = [];
        for (;;) {
            skipSpace();
            const token = peek();
            if (token === undefined) {
                if (stop === undefined) return nodes;
                err(0, '公式在 } / \\right / \\end 之前结束了');
            }
            if (stop(token)) return nodes;
            if (token.kind === 'nl') { next(); continue; }
            nodes.push(parseElement());
        }
    }

    /** 顶层：\\\\ 分块（展示式里每块一条 oMath）。 */
    function parseTop() {
        const blocks = [];
        let nodes = [];
        for (;;) {
            skipSpace();
            const token = peek();
            if (token === undefined) break;
            if (token.kind === 'nl') {
                next();
                blocks.push(join(nodes));
                nodes = [];
                continue;
            }
            nodes.push(parseElement());
        }
        blocks.push(join(nodes));
        return blocks;
    }

    return { parseTop };
}

/** plain 里给分式分子/分母按需加括号（短的干净内容直接用）。 */
function wrapParen(text) {
    const trimmed = text.trim();
    if (trimmed === '') return '';
    if ([...trimmed].length <= 3 && !/[\s+\-/=<>]/.test(trimmed)) return trimmed;
    return `(${trimmed})`;
}

/** plain 清理：压缩连续空白。 */
function cleanPlain(text) {
    return text.replace(/\s+/g, ' ').trim();
}

const cache = new Map();

/**
 * LaTeX → OMML。
 *
 * @param {string} latex 公式源码
 * @param {{rPr?: string}} [options] rPr 插到每条 m:r 里（a:rPr：字号/颜色由调用方拼好）
 * @returns {{ok:true, blocks:string[], plain:string} | {ok:false, error:string}}
 *   blocks：按 \\\\ 分块，每块是 m:oMath 的内部内容；plain：纯文本近似
 */
export function latexToMath(latex, options = {}) {
    const rPr = typeof options.rPr === 'string' ? options.rPr : '';
    // 缓存键必须带 rPr 全文（带不同字号的同一公式是不同产出，只按长度会撞键）
    const key = `${rPr}｜${latex}`;
    if (cache.has(key)) return cache.get(key);
    let result;
    try {
        const parser = createParser(tokenize(String(latex ?? '')), rPr);
        const blocks = parser.parseTop();
        const plain = cleanPlain(blocks.map((block) => block.p).join('；'));
        result = { ok: true, blocks: blocks.map((block) => block.o), plain };
    } catch (error) {
        result = { ok: false, error: error instanceof MathError ? error.message : String(error?.message ?? error) };
    }
    if (cache.size > 400) cache.clear();
    cache.set(key, result);
    return result;
}

/**
 * LaTeX → 纯文本近似（字数统计 / 折行估算用）。
 * 解析失败时返回原样字符串 —— 量算宁可偏大也不能归零。
 */
export function mathPlain(latex) {
    const parsed = latexToMath(String(latex ?? ''));
    if (parsed.ok) return parsed.plain;
    return String(latex ?? '');
}
