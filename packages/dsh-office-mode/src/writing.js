/**
 * 写作规划与文本自检（第七十三轮）。
 *
 * 为什么要有这个模块：提示词能说「重点给足、难点写透」，但模型不知道哪一节
 * 是重点、该给多少字、那节要完成什么动作。这里把「按难度与读者收益分配注意力」
 * 变成可计算的表：每一节拿到目标字数、深度档、必须写到的内容、语言与比喻的边界、
 * 以及一个可回查的自检入口。
 *
 * 两条边界：
 *   1. 规划不是结构模板。节名来自用户；没有给节时按文体给一套可改的骨架，
 *      但真正的产物是每节的注意力与写法，不是六个小标题。
 *   2. 自检是启发式，不是 AI 检测器。它只报可解释的信号（句长过匀、套话密度、
 *      重点节没写透、通篇没有让步），不声称「像人」可以被一个分数证明。
 *
 * @module dsh-office-mode/writing
 */

const KIND_PROFILES = {
    context: {
        label: '背景', base: 0.55, defaultDifficulty: 0.8, defaultValue: 1,
        purpose: '只交代读者不知道的那部分来处，让后文的问题成立。',
        readerQuestion: '为什么现在要谈这件事？',
        mustCover: ['读者已知的用一句带过', '真正缺的那块信息', '这件事与读者的关系'],
        moves: {
            open: '从一个具体的时间、现象或冲突切入，不要从「随着……的发展」开始。',
            develop: ['只写读者不知道的部分，常识一句带过', '把「为什么现在」落到一个可核对的变化上'],
            turn: '如果背景里有反直觉之处，在这里点一句，留给后文。',
            land: '最后一句交给问题，不要总结背景。',
        },
        avoid: ['从盘古开天辟地写起', '用「重要性/复杂性」代替事实'],
        register: 'plain', figurative: 'none',
    },
    problem: {
        label: '问题', base: 0.95, defaultDifficulty: 1.4, defaultValue: 2,
        purpose: '把「大家都觉得没问题」的地方撬开，给出一个可回答的问题。',
        readerQuestion: '原来的解释哪里不够用？',
        mustCover: ['旧解释的裂缝', '一个具体的反例或异常', '本文要回答的那个问句'],
        moves: {
            open: '先给一个旧解释能解释的现象，再给一个它解释不了的。',
            develop: ['把裂缝写成一个可检验的差异，不要只说「存在不足」', '说明这个不足为什么值得现在解决'],
            turn: '承认旧解释在多数场景仍然有效，限定你的问题边界。',
            land: '把问题收成一句读者能记住的话。',
        },
        avoid: ['用「仍有待深入」代替具体缺口', '把文献综述写成流水账'],
        register: 'mixed', figurative: 'one-precise-analogy',
    },
    method: {
        label: '方法', base: 1.05, defaultDifficulty: 2, defaultValue: 2,
        purpose: '让读者能复现关键步骤，或至少判断结论是怎么来的。',
        readerQuestion: '你是怎么得到这个结果的？',
        mustCover: ['关键步骤与选择理由', '数据/材料从哪来', '一个能复现的最小细节'],
        moves: {
            open: '先给方法的整体动作（一句话），再拆步骤。',
            develop: ['每一步写清输入、动作、输出', '解释为什么选它而不是更常见的替代', '给出一个最小例子走通全流程'],
            turn: '指出方法在哪个环节最脆弱。',
            land: '落在这套方法能回答与不能回答的边界上。',
        },
        avoid: ['只列工具名不写动作', '把方法写成广告'],
        register: 'precise', figurative: 'none',
    },
    theory: {
        label: '理论/机制', base: 1.3, defaultDifficulty: 2.6, defaultValue: 3,
        purpose: '把「为什么会这样」拆到读者能自己往下推一步。',
        readerQuestion: '机制到底怎么运转？',
        mustCover: ['关键概念的一次性定义', 'A 如何导致 B 的链条', '一个最小例子或反例检验链条'],
        moves: {
            open: '从一个具体的困惑或反直觉的事实切入，不要从定义开始。',
            develop: [
                '把关键概念拆到可操作：它在什么条件下出现、什么条件下不出现',
                '写出机制链条：先发生什么、再发生什么、哪一步是必要的',
                '用一个最小例子走一遍链条，再给一个链条失效的边界',
            ],
            turn: '主动给出一个竞争解释，并说明你为什么没有采用它。',
            land: '用一句短句说清它改变了什么判断。',
        },
        avoid: ['堆术语不解释', '用「机制复杂」代替机制', '连续三段纯抽象推演'],
        register: 'precise', figurative: 'one-precise-analogy',
    },
    evidence: {
        label: '证据', base: 1.05, defaultDifficulty: 1.5, defaultValue: 2.5,
        purpose: '让结论有可核对的依据，而不是形容词。',
        readerQuestion: '凭什么这么说？',
        mustCover: ['关键数字或对照', '来源与口径', '不确定性与替代解释'],
        moves: {
            open: '先给结论数字或对照结果，再解释它意味着什么。',
            develop: ['写清样本、口径、比较对象', '把「显著」换成具体差值', '给一个不支持结论的观察'],
            turn: '指出证据最弱的一环。',
            land: '落到「这条证据支持到什么程度」。',
        },
        avoid: ['只写「显著提升」不给数', '把相关性写成因果'],
        register: 'precise', figurative: 'none',
    },
    analysis: {
        label: '分析', base: 1.2, defaultDifficulty: 2.2, defaultValue: 2.5,
        purpose: '把证据与机制接起来，回答「所以呢」。',
        readerQuestion: '这些材料合起来说明什么？',
        mustCover: ['材料与结论之间的那一步推理', '一个替代解释', '结论的适用范围'],
        moves: {
            open: '先给一个需要解释的模式或矛盾。',
            develop: ['把证据按机制链条排，而不是按时间排', '写出替代解释并排除', '给一个边界条件'],
            turn: '承认一处解释不了的材料。',
            land: '落到一个可被检验的判断。',
        },
        avoid: ['复述结果不推进', '用「这说明」跳过推理'],
        register: 'precise', figurative: 'one-precise-analogy',
    },
    leap: {
        label: '跃迁', base: 1.45, defaultDifficulty: 2.8, defaultValue: 3,
        purpose: '把一个洞察推到读者原来的问题框架之外，但每一步都留脚印。',
        readerQuestion: '如果这是真的，什么会变？',
        mustCover: ['一句可检验的跃迁命题', '它从哪一步推出来', '最强反对意见与一个可证伪的观察'],
        moves: {
            open: '先把跃迁写成一句可检验的话：如果它成立，读者会怎么看待原来的问题。',
            develop: [
                '交代它从哪一步推出来，不要把跳跃藏在形容词里',
                '把假设摆到桌面上：它依赖什么、不依赖什么',
                '给出一个能证伪它的观察，或一个最小的思想实验',
            ],
            turn: '写出最强的反对意见，不要设稻草人。',
            land: '不要喊口号，落在一个具体后果或下一步动作上。',
        },
        avoid: ['用「颠覆/范式转移」代替论证', '把推测写成事实', '连续排比抒情'],
        register: 'precise', figurative: 'one-precise-analogy',
    },
    implication: {
        label: '含义', base: 1.05, defaultDifficulty: 1.8, defaultValue: 2.2,
        purpose: '说清这个结论对谁、在什么时间尺度上、意味着什么。',
        readerQuestion: '这跟我有什么关系？',
        mustCover: ['受影响的对象', '影响发生的条件', '一个可观察的后果'],
        moves: {
            open: '先点名一个具体对象或场景。',
            develop: ['把「影响」拆成可观察的变化', '给出时间尺度与前提'],
            turn: '指出影响不成立的场景。',
            land: '落在一个人能做的判断或动作上。',
        },
        avoid: ['泛泛说「具有重要意义」', '把愿望写成预测'],
        register: 'mixed', figurative: 'one-precise-analogy',
    },
    example: {
        label: '例子', base: 0.95, defaultDifficulty: 1.2, defaultValue: 2,
        purpose: '用一个具体场景把抽象机制落地，同时暴露它的边界。',
        readerQuestion: '能举个具体例子吗？',
        mustCover: ['一个可想象的场景', '场景里机制如何起作用', '这个例子在哪里不适用'],
        moves: {
            open: '直接进入场景：谁、在什么时候、遇到什么。',
            develop: ['给两三个具体细节，不要只给标签', '把场景与机制逐步对上', '指出例子为了清楚省略了什么'],
            turn: '用一个反例或边界把例子收住。',
            land: '回到读者能带走的那一句。',
        },
        avoid: ['例子只用来装饰', '编造具体人名与数据'],
        register: 'plain', figurative: 'everyday-metaphor',
    },
    howto: {
        label: '做法', base: 1.0, defaultDifficulty: 1.5, defaultValue: 2,
        purpose: '让读者能照着做，或判断自己该不该做。',
        readerQuestion: '具体怎么做？',
        mustCover: ['前置条件', '按顺序的关键动作', '失败时怎么判断与回退'],
        moves: {
            open: '先说做完是什么样，再给第一步。',
            develop: ['按动作顺序写，每步一个判断点', '给一个常见失败与检查方法', '标出可以跳过的步骤'],
            turn: '说明这套做法在什么情况下不适用。',
            land: '给一个最小的下一步。',
        },
        avoid: ['步骤里夹大段原理', '每步都写「注意」'],
        register: 'plain', figurative: 'everyday-metaphor',
    },
    limitation: {
        label: '局限', base: 0.8, defaultDifficulty: 1.8, defaultValue: 2,
        purpose: '把结论的边界说清，让读者知道什么时候不该信。',
        readerQuestion: '什么情况下会失效？',
        mustCover: ['最可能失效的条件', '一个具体反例', '对结论的影响程度'],
        moves: {
            open: '直接给最可能失效的那个条件。',
            develop: ['用一个具体场景说明怎么失效', '区分「结论错」与「适用范围窄」', '给出补救或下一步检验'],
            turn: '不要用「样本量不足」这类通用句代替具体局限。',
            land: '落到读者该如何使用这个结论。',
        },
        avoid: ['把局限写成自我批评清单', '只写「未来研究可以改进」'],
        register: 'precise', figurative: 'none',
    },
    summary: {
        label: '收束', base: 0.6, defaultDifficulty: 0.8, defaultValue: 1.2,
        purpose: '把读者带走的那一句留下，不重复全文。',
        readerQuestion: '最后该记住什么？',
        mustCover: ['一个判断或后果', '一个未解决的张力', '一个可做的下一步'],
        moves: {
            open: '直接给判断，不要「综上所述」。',
            develop: ['只回收前文没有明说的那层意思', '留一个具体的未决问题'],
            turn: '可以承认一个仍不确定的地方。',
            land: '用一句短句收住，不写祝福语。',
        },
        avoid: ['逐节复述', '用「希望本文有所帮助」结尾'],
        register: 'plain', figurative: 'none',
    },
    faq: {
        label: '答疑', base: 0.75, defaultDifficulty: 1.0, defaultValue: 1.5,
        purpose: '回答读者最可能卡住的那两三个问题。',
        readerQuestion: '我遇到的情况算不算例外？',
        mustCover: ['一个真实会被问的问题', '直接回答', '适用条件'],
        moves: {
            open: '用读者会问的原话提问。',
            develop: ['先给短答案，再给条件', '给一个具体场景'],
            turn: '指出问题本身可能问错了的地方。',
            land: '给一个可操作的判断标准。',
        },
        avoid: ['自问自答凑节数', '答案绕回定义'],
        register: 'plain', figurative: 'none',
    },
};

const GENRE_PROFILES = {
    academic: {
        label: '学术论文', register: 'precise', figurative: 'one-precise-analogy',
        rhythm: '以 20–35 字的中长句为主；每 4–6 句插一个 10 字以内的短句落地；避免连续三句同长。',
        sections: [
            { heading: '问题的来处', kind: 'context', difficulty: 1, value: 1.2 },
            { heading: '旧解释的裂缝', kind: 'problem', difficulty: 1.8, value: 2.4 },
            { heading: '方法与证据', kind: 'method', difficulty: 2.2, value: 2.2 },
            { heading: '结果与机制', kind: 'analysis', difficulty: 2.6, value: 3 },
            { heading: '跃迁：这意味着什么', kind: 'leap', difficulty: 3, value: 3 },
            { heading: '局限与反例', kind: 'limitation', difficulty: 2, value: 2 },
        ],
    },
    explainer: {
        label: '说明文', register: 'plain', figurative: 'everyday-metaphor',
        rhythm: '16–28 字为主；允许口语短句与一个不完整的停顿；关键判断用短句，解释用长句。',
        sections: [
            { heading: '从一个具体场景开始', kind: 'example', difficulty: 1, value: 1.5 },
            { heading: '反直觉的地方', kind: 'problem', difficulty: 1.8, value: 2.4 },
            { heading: '机制：它是怎么发生的', kind: 'theory', difficulty: 2.4, value: 3 },
            { heading: '一个日常比喻与它的边界', kind: 'example', difficulty: 1.4, value: 2.2 },
            { heading: '你可能遇到的情况', kind: 'howto', difficulty: 1.6, value: 2 },
            { heading: '边界与常见误解', kind: 'limitation', difficulty: 1.8, value: 1.8 },
        ],
    },
    report: {
        label: '工作汇报', register: 'precise', figurative: 'none',
        rhythm: '18–30 字为主；结论句短，依据句可以长；每段第一句承担信息，不要用「另外」开头。',
        sections: [
            { heading: '结论先行', kind: 'summary', difficulty: 1, value: 3 },
            { heading: '依据与数据', kind: 'evidence', difficulty: 1.6, value: 3 },
            { heading: '机制与原因', kind: 'analysis', difficulty: 2.2, value: 2.4 },
            { heading: '风险与反例', kind: 'limitation', difficulty: 2, value: 2.2 },
            { heading: '下一步建议', kind: 'howto', difficulty: 1.5, value: 2.4 },
        ],
    },
    memo: {
        label: '备忘/提案', register: 'precise', figurative: 'none',
        rhythm: '短句为主；一段只说一件事；判断句与依据句交替。',
        sections: [
            { heading: '要决定的事', kind: 'problem', difficulty: 1.2, value: 3 },
            { heading: '选项与代价', kind: 'analysis', difficulty: 2, value: 2.6 },
            { heading: '建议与理由', kind: 'analysis', difficulty: 2, value: 3 },
            { heading: '风险与回退', kind: 'limitation', difficulty: 1.6, value: 2 },
        ],
    },
    essay: {
        label: '随笔/评论', register: 'mixed', figurative: 'everyday-metaphor',
        rhythm: '长短交替；叙事段可以短促，议论段允许一个长句；不要每段都同样长。',
        sections: [
            { heading: '一个具体的入口', kind: 'example', difficulty: 1, value: 1.5 },
            { heading: '我们通常怎么想', kind: 'context', difficulty: 1, value: 1.4 },
            { heading: '哪里不对', kind: 'problem', difficulty: 2, value: 2.6 },
            { heading: '另一种看法', kind: 'leap', difficulty: 2.6, value: 3 },
            { heading: '它改变什么', kind: 'implication', difficulty: 2.2, value: 2.4 },
            { heading: '收束', kind: 'summary', difficulty: 1, value: 1.6 },
        ],
    },
    howto: {
        label: '操作指南', register: 'plain', figurative: 'everyday-metaphor',
        rhythm: '14–24 字为主；步骤句短，解释句可长；每步只做一件事。',
        sections: [
            { heading: '做完是什么样', kind: 'example', difficulty: 1, value: 2 },
            { heading: '开始之前', kind: 'context', difficulty: 1, value: 2 },
            { heading: '具体步骤', kind: 'howto', difficulty: 2, value: 3 },
            { heading: '常见失败与回退', kind: 'limitation', difficulty: 1.8, value: 2.4 },
        ],
    },
    general: {
        label: '通用文稿', register: 'mixed', figurative: 'one-precise-analogy',
        rhythm: '长短交替；关键判断用短句，解释用长句；段落长度不要持平。',
        sections: [
            { heading: '是什么', kind: 'context', difficulty: 1, value: 1.5 },
            { heading: '为什么难', kind: 'problem', difficulty: 1.8, value: 2.4 },
            { heading: '怎么做/怎么运作', kind: 'method', difficulty: 2, value: 2.6 },
            { heading: '边界与代价', kind: 'limitation', difficulty: 1.8, value: 2 },
            { heading: '收束', kind: 'summary', difficulty: 1, value: 1.5 },
        ],
    },
};

const GENRE_ALIASES = {
    paper: 'academic', thesis: 'academic', 学术: 'academic', 论文: 'academic',
    article: 'explainer', 说明: 'explainer', 科普: 'explainer', 教程: 'howto',
    business: 'report', 汇报: 'report', 报告: 'report',
    proposal: 'memo', 提案: 'memo', 备忘: 'memo',
    comment: 'essay', 评论: 'essay', 随笔: 'essay',
    guide: 'howto', 指南: 'howto', 操作: 'howto',
    general: 'general', 通用: 'general',
};

const REGISTER_LABELS = {
    precise: '准确克制', plain: '平实', conversational: '放松', mixed: '张弛交替',
};

const FIGURATIVE_LABELS = {
    none: '不用比喻',
    'one-precise-analogy': '最多一个精确类比，并当场兑现成机制',
    'everyday-metaphor': '可用日常比喻（每个概念最多一个，随后兑现成机制）',
    scene: '用一个具体场景代替比喻',
};

const DEPTH_LABELS = { compress: '收住', touch: '点到', develop: '展开', dwell: '写透' };

const GLOBAL_ANTI_PATTERNS = [
    '不要把「首先/其次/最后」当骨架；段落靠内容推进，不靠序号。',
    '不要每节结尾都总结；总结只在全文收束处出现一次。',
    '不要用「不仅……而且」「随着……的发展」「在当今……时代」凑句子。',
    '不要把「重要性/复杂性/挑战性」当内容；换成事实、数字或场景。',
    '不要把每段写到差不多长；该短的地方短，该展开的地方展开。',
    '不要在学术段里堆形容词；比喻最多一个，且要说明它在哪里不成立。',
    '不要写没有依据的「显著提升/深刻改变」；给出差值与口径。',
    '不要为了像人而编造个人经历、对话或数据；不确定就标注推测。',
    '不要把「可能/或许」当安全带；该下判断时下判断，并把依据放在旁边。',
    '不要在结尾重复开头；留一个未决问题或下一步。',
];

const CLICHE_PATTERNS = [
    { term: '首先/其次/最后', re: /(?:^|[。！？\n])\s*(?:首先|其次|最后|再者|此外)[，,]/gu },
    { term: '综上所述', re: /综上所述|总而言之|总的来说/gu },
    { term: '由此可见', re: /由此可见|不难看出|不难发现/gu },
    { term: '值得注意的是', re: /值得注意的是|需要指出的是|值得一提的是/gu },
    { term: '众所周知', re: /众所周知/gu },
    { term: '随着……发展', re: /随着[^。！？\n]{0,16}(?:的)?(?:发展|进步|普及|兴起)/gu },
    { term: '在当今……', re: /在当今[^。！？\n]{0,12}(?:时代|社会|背景|环境)/gu },
    { term: '不仅……而且', re: /不仅[^。！？\n]{0,24}而且/gu },
    { term: '一方面……另一方面', re: /一方面[^。！？\n]{0,48}另一方面/gu },
    { term: '赋能/抓手/闭环', re: /赋能|抓手|闭环|底层逻辑|顶层设计|降本增效|深度融合/gu },
    { term: '全方位/多维度', re: /全方位|多维度|全链条|系统性/gu },
    { term: '助力/打造/彰显', re: /助力|打造|彰显|凸显/gu },
    { term: '极大地/有效地', re: /极大地|显著地|有效地|更好地|持续地/gu },
    { term: 'delve/tapestry', re: /delve|tapestry|a testament to|navigate the landscape/giu },
    { term: 'moreover/in conclusion', re: /moreover|furthermore|in conclusion|it is important to note/giu },
];

const HEDGE_RE = /可能|或许|也许|大概|似乎|一定程度上|在某种意义上|或许可以说/gu;
const METAPHOR_RE = /好像|就像|好比|如同|仿佛|简直是|相当于|打个比方|如同.{0,12}一般|像[^。！？\n]{0,16}一样/gu;
const TURN_RE = /但|不过|然而|反过来|反例|问题在于|恰恰|局限|失效|不成立|例外|代价|反对/gu;
const ABSTRACT_RE = /[\u4e00-\u9fa5]{1,3}(?:性|化|度|主义|机制|逻辑|范式|维度|层面)/gu;
const EXAMPLE_RE = /例如|比如|举个例子|具体来说|有一次|设想|假设你/gu;

function clamp(value, min, max) {
    const number = Number(value);
    if (!Number.isFinite(number)) return min;
    return Math.min(max, Math.max(min, number));
}

function textOf(value, fallback = '') {
    return typeof value === 'string' && value.trim() !== '' ? value.trim() : fallback;
}

function countChars(value) {
    return Array.from(String(value ?? '').replace(/\s+/gu, '')).length;
}

function roundTo(value, step = 10) {
    return Math.round(Number(value) / step) * step;
}

function normalizeHeading(value) {
    return String(value ?? '')
        .replace(/^#{1,6}\s*/u, '')
        .replace(/^[一二三四五六七八九十\d]+[、.．)）]\s*/u, '')
        .replace(/[\s*_~:：。，,、；;！!？?（）()【】\[\]「」『』]/gu, '')
        .toLowerCase();
}

function genreOf(value) {
    const key = textOf(value, 'general').toLowerCase();
    if (GENRE_PROFILES[key] !== undefined) return key;
    const alias = GENRE_ALIASES[key];
    return alias !== undefined && GENRE_PROFILES[alias] !== undefined ? alias : 'general';
}

function kindOf(value) {
    const key = textOf(value, '').toLowerCase();
    if (KIND_PROFILES[key] !== undefined) return key;
    const aliases = {
        background: 'context', intro: 'context', 背景: 'context',
        gap: 'problem', 问题: 'problem',
        methodology: 'method', 方法: 'method',
        mechanism: 'theory', 机制: 'theory', 理论: 'theory',
        result: 'evidence', 结果: 'evidence', 数据: 'evidence',
        discussion: 'analysis', 分析: 'analysis',
        insight: 'leap', 洞察: 'leap', 跃迁: 'leap',
        impact: 'implication', 影响: 'implication', 意义: 'implication',
        case: 'example', 例子: 'example', 案例: 'example',
        step: 'howto', 步骤: 'howto', 操作: 'howto',
        risk: 'limitation', 风险: 'limitation', 局限: 'limitation',
        conclusion: 'summary', 总结: 'summary', 结论: 'summary',
        question: 'faq', 答疑: 'faq',
    };
    return aliases[key] ?? null;
}

function inferKind(heading) {
    const value = String(heading ?? '');
    const rules = [
        [/背景|现状|来处|已知|context|background|intro/i, 'context'],
        [/问题|裂缝|张力|反直觉|缺口|problem|gap/i, 'problem'],
        [/方法|做法|流程|步骤|实验设计|method|methodology/i, 'method'],
        [/理论|机制|原理|推导|模型|theory|mechanism/i, 'theory'],
        [/证据|数据|结果|实验|evidence|result/i, 'evidence'],
        [/分析|讨论|原因|为何|解释|analysis|discussion/i, 'analysis'],
        [/跃迁|洞察|启示|第三种|leap|insight/i, 'leap'],
        [/影响|意义|含义|展望|implication|impact/i, 'implication'],
        [/例子|案例|场景|比喻|故事|生平|人物|example|case/i, 'example'],
        [/步骤|做法|操作|指南|怎么|howto|tutorial/i, 'howto'],
        [/局限|风险|边界|反例|不足|限制|limitation|risk/i, 'limitation'],
        [/总结|结论|收束|summary|conclusion/i, 'summary'],
        [/答疑|问答|常见问题|faq/i, 'faq'],
    ];
    for (const pair of rules) {
        if (pair[0].test(value)) return pair[1];
    }
    return null;
}

function registerFor(kind, genre, override) {
    if (typeof override === 'string' && REGISTER_LABELS[override] !== undefined) return override;
    const profile = KIND_PROFILES[kind] ?? KIND_PROFILES.analysis;
    if (kind === 'example' || kind === 'howto') {
        return genre === 'academic' || genre === 'report' || genre === 'memo' ? 'plain' : profile.register;
    }
    if (genre === 'explainer' && (kind === 'theory' || kind === 'analysis')) return 'plain';
    return profile.register;
}

function figurativeFor(kind, genre, override) {
    if (typeof override === 'string' && FIGURATIVE_LABELS[override] !== undefined) return override;
    if (genre === 'academic' || genre === 'report' || genre === 'memo') {
        if (kind === 'example') return 'scene';
        if (kind === 'theory' || kind === 'analysis' || kind === 'leap' || kind === 'problem' || kind === 'implication') return 'one-precise-analogy';
        return 'none';
    }
    if (genre === 'explainer' || genre === 'howto' || genre === 'essay') {
        if (kind === 'example' || kind === 'theory' || kind === 'howto' || kind === 'analysis') return 'everyday-metaphor';
        return 'none';
    }
    if (kind === 'example') return 'everyday-metaphor';
    if (kind === 'theory' || kind === 'analysis' || kind === 'leap') return 'one-precise-analogy';
    return 'none';
}

function rhythmFor(register, genre) {
    if (register === 'precise') {
        return genre === 'report'
            ? '18–30 字为主；结论句短，依据句可以长；每段第一句承担信息。'
            : '20–35 字为主；每 4–6 句插一个 10 字以内的短句落地；避免连续三句同长。';
    }
    if (register === 'plain') return '16–28 字为主；可以用口语短句；关键判断用短句，解释用长句。';
    if (register === 'conversational') return '12–24 字为主；可以用「你」「其实」，但每段最多一次。';
    return '长短交替；关键判断用短句，解释用长句；不要让每段一样长。';
}

function normalizeMoves(profile, override) {
    const moves = {
        open: textOf(override?.open, profile.moves.open),
        develop: Array.isArray(override?.develop) && override.develop.length > 0
            ? override.develop.map((line) => textOf(line, '')).filter(Boolean)
            : profile.moves.develop.slice(),
        turn: textOf(override?.turn, profile.moves.turn),
        land: textOf(override?.land, profile.moves.land),
    };
    if (moves.develop.length === 0) moves.develop = profile.moves.develop.slice();
    return moves;
}

function normalizeSection(item, index, genre) {
    const source = typeof item === 'string' ? { heading: item } : (item ?? {});
    const heading = textOf(source.heading, '第 ' + (index + 1) + ' 节');
    const kind = kindOf(source.kind) ?? inferKind(heading) ?? 'analysis';
    const profile = KIND_PROFILES[kind] ?? KIND_PROFILES.analysis;
    const difficulty = clamp(source.difficulty ?? profile.defaultDifficulty, 0, 3);
    const value = clamp(source.value ?? profile.defaultValue, 0, 3);
    const register = registerFor(kind, genre, source.register);
    const figurative = figurativeFor(kind, genre, source.figurative);
    return {
        heading,
        kind,
        difficulty: Math.round(difficulty * 10) / 10,
        value: Math.round(value * 10) / 10,
        purpose: textOf(source.purpose, profile.purpose),
        readerQuestion: textOf(source.readerQuestion, profile.readerQuestion),
        mustCover: Array.isArray(source.mustCover) && source.mustCover.length > 0
            ? source.mustCover.map((line) => textOf(line, '')).filter(Boolean)
            : profile.mustCover.slice(),
        moves: normalizeMoves(profile, source.moves),
        register,
        figurative,
        rhythm: textOf(source.rhythm, rhythmFor(register, genre)),
        avoid: Array.isArray(source.avoid) && source.avoid.length > 0
            ? source.avoid.map((line) => textOf(line, '')).filter(Boolean)
            : profile.avoid.slice(),
    };
}

function templateSections(genre) {
    return GENRE_PROFILES[genre].sections.map((item) => ({ ...item }));
}

function allocate(sections, totalChars) {
    const weights = sections.map((section) => {
        const profile = KIND_PROFILES[section.kind] ?? KIND_PROFILES.analysis;
        return profile.base * (1 + section.difficulty * 0.16) * (1 + section.value * 0.10);
    });
    const count = sections.length;
    const cap = count <= 1 ? 1 : count === 2 ? 0.65 : count === 3 ? 0.5 : 0.4;
    const shares = new Array(count).fill(0);
    let remaining = 1;
    let active = weights.map((unused, index) => index);
    let guard = 0;
    while (active.length > 0 && guard < 100) {
        guard += 1;
        const sum = active.reduce((acc, index) => acc + weights[index], 0);
        let capped = false;
        for (const index of active) {
            const share = sum === 0 ? 1 / active.length : (remaining * weights[index]) / sum;
            if (share > cap) {
                shares[index] = cap;
                remaining -= cap;
                capped = true;
            }
        }
        if (!capped) {
            for (const index of active) {
                shares[index] = sum === 0 ? remaining / active.length : (remaining * weights[index]) / sum;
            }
            break;
        }
        active = active.filter((index) => shares[index] === 0);
    }
    if (shares.some((share) => !Number.isFinite(share) || share < 0)) {
        for (let index = 0; index < count; index += 1) shares[index] = 1 / count;
    }

    const minChars = Math.max(20, Math.min(80, Math.floor(totalChars / count / 10) * 10));
    const maxChars = Math.max(minChars, roundTo(totalChars * cap, 10));
    const targets = shares.map((share) => Math.max(minChars, roundTo(totalChars * share, 10)));
    let difference = totalChars - targets.reduce((sum, value) => sum + value, 0);
    let adjustGuard = 0;
    while (difference !== 0 && adjustGuard < 20000) {
        adjustGuard += 1;
        const direction = difference > 0 ? 1 : -1;
        const order = sections
            .map((section, index) => ({ index, weight: weights[index], target: targets[index] }))
            .filter((entry) => entry.target + direction * 10 >= minChars && entry.target + direction * 10 <= maxChars)
            .sort((left, right) => (direction > 0 ? right.weight - left.weight : left.weight - right.weight));
        if (order.length === 0) break;
        targets[order[0].index] += direction * 10;
        difference -= direction * 10;
    }
    if (difference !== 0) targets[targets.length - 1] = Math.max(minChars, targets[targets.length - 1] + difference);
    const sum = targets.reduce((acc, value) => acc + value, 0);
    if (sum !== totalChars) targets[targets.length - 1] = Math.max(minChars, targets[targets.length - 1] + (totalChars - sum));

    return targets.map((target, index) => ({
        weight: Math.round(weights[index] * 100) / 100,
        targetChars: target,
        share: totalChars === 0 ? 0 : target / totalChars,
    }));
}

function depthOf(share, count) {
    if (count === 1 || share >= 0.25) return 'dwell';
    if (share >= 0.15) return 'develop';
    if (share >= 0.08) return 'touch';
    return 'compress';
}

function planBrief(input, options) {
    const source = typeof input === 'string' ? { brief: input } : (input ?? {});
    const title = textOf(source.title, typeof input === 'string' ? textOf(input.split(/\r?\n/u)[0], '未命名文稿').slice(0, 40) : '未命名文稿');
    const genre = genreOf(source.genre ?? options.genre);
    const audience = textOf(source.audience ?? options.audience, '不熟悉这个问题的读者');
    const goal = textOf(source.goal ?? options.goal, '');
    const targetChars = clamp(source.targetChars ?? options.targetChars ?? 1800, 300, 20000);
    const throughLine = textOf(source.throughLine ?? options.throughLine, '');
    return { source, title, genre, audience, goal, targetChars: Math.round(targetChars), throughLine };
}

function buildChecks(brief, sections) {
    const checks = [
        '三问①：读者要同时拿住几个新概念？超过 2 个就拆步、给例子。',
        '三问②：这一节 / 这段话改变读者哪个判断？不改变就删，或压成一句。',
        '三问③：全文最长的一节是不是最难的那点？若最长的是背景 / 常识，规划失败，重分篇幅。',
    ];
    if (brief.goal === '') checks.push('还没有目标：先补一句「读者读完后会改变什么判断」，否则注意力没有依据。');
    if (brief.throughLine === '') checks.push('主线空着：写正文前用一句话写下「如果读者只记住一句，是哪句」。');
    if (sections.length >= 4) {
        const weights = sections.map((section) => section.weight);
        const max = Math.max(...weights);
        const min = Math.min(...weights);
        if (max / min < 1.3) checks.push('各节权重接近（最大/最小 < 1.3）：确认是否真的同等重要；如果不是，调高难点/跃迁节的 difficulty 或 value。');
    }
    if (sections.length >= 3) {
        const byTarget = [...sections].sort((left, right) => right.targetChars - left.targetChars)[0];
        const byNeed = [...sections].sort((left, right) => (right.difficulty * 1.2 + right.value) - (left.difficulty * 1.2 + left.value))[0];
        if (byTarget.id !== byNeed.id) {
            checks.push('最长的一节是「' + byTarget.heading + '」（' + byTarget.targetChars + ' 字），但最难 / 最值钱的是「' + byNeed.heading + '」：确认是不是把笔墨给了铺垫。');
        }
    }
    if (brief.targetChars < 600) checks.push('总字数较短：规划只保留最关键的三节，别为了凑表把每节写成摘要。');
    if (sections.length > 7) checks.push('节数偏多：读者记不住七节以上；能合并的合并，把笔墨留给难点。');
    return checks;
}

/**
 * 做一份写作规划。
 *
 * input 可以是：
 *   · 一句话 brief（string）；
 *   · { title, genre, audience, goal, targetChars, sections, throughLine }；
 *   · sections 里每项是标题字符串，或 { heading, kind, difficulty, value, purpose, mustCover, moves, register, figurative, rhythm, avoid }。
 *
 * genre：academic / explainer / report / memo / essay / howto / general（也认中英文别名）。
 * 返回值里的 markdown 是给模型直接读的版本；调用方通常用 office.log(plan.markdown)
 * 把它带回上下文，再按它写。
 */
export function writingPlan(input, options = {}) {
    const brief = planBrief(input, options ?? {});
    const sourceSections = Array.isArray(brief.source.sections)
        ? brief.source.sections
        : Array.isArray(brief.source.outline)
            ? brief.source.outline
            : null;
    const rawSections = sourceSections !== null && sourceSections.length > 0
        ? sourceSections
        : templateSections(brief.genre);
    const sections = rawSections.map((item, index) => normalizeSection(item, index, brief.genre));
    const allocation = allocate(sections, brief.targetChars);
    const withAllocation = sections.map((section, index) => {
        const share = allocation[index].share;
        const depth = depthOf(share, sections.length);
        return {
            id: 's' + (index + 1),
            ...section,
            weight: allocation[index].weight,
            share: Math.round(share * 1000) / 1000,
            targetChars: allocation[index].targetChars,
            depth,
            depthLabel: DEPTH_LABELS[depth],
            registerLabel: REGISTER_LABELS[section.register],
            figurativeLabel: FIGURATIVE_LABELS[section.figurative],
        };
    });
    const checks = buildChecks(brief, withAllocation);
    const strategy = {
        throughLine: brief.throughLine,
        reader: { audience: brief.audience, goal: brief.goal, question: '读者读完要能自己回答的那个问题' },
        rule: '篇幅按内容难度与读者收益分配，不按节数平均；难点/跃迁/理论多给笔墨，背景与常识收住。',
        texture: {
            anchor: '每 150 字至少一个可核对的细节：数字（带口径）、日期、专名、原话或具体场景；没有就写清「这是推测」，不要编。',
            turn: '每个核心主张配一条最强反对或竞争解释，并写清它在什么条件下成立；位置放在论证最顺的时候，不要放在末尾当免责声明。',
            landing: '每节最后一句写判断或后果，不写「综上所述」；全文只在收束处总结一次，总结句必须带新信息（条件 / 代价 / 下一步）。',
            paragraph: '一段只推进一个信息差：段首给主张，段内先旧信息后新信息，段末放最重的那句；不要一段里换两个功能。',
            transition: '连接词是逻辑的冗余编码：靠语序、关键词重复、代词回指能接上就删；「首先/其次/最后」全文只许出现一组。',
            uncertainty: '语言强度跟证据强度对齐：相关性写「与……相关（n=…）」，机制未验证写「一种解释是」，外推写「在……条件下」。',
            detail: '背景与常识压到 1–2 句；省下的笔墨给新定义、反直觉结论、机制与跃迁点。',
        },
        rhythm: GENRE_PROFILES[brief.genre].rhythm,
        figurative: GENRE_PROFILES[brief.genre].figurative,
        antiPatterns: GLOBAL_ANTI_PATTERNS.slice(),
    };
    const plan = {
        version: 1,
        brief,
        strategy,
        sections: withAllocation,
        checks,
        totalChars: brief.targetChars,
        markdown: '',
    };
    plan.markdown = renderPlan(plan);
    return plan;
}

function joinMove(move) {
    const develop = Array.isArray(move.develop) ? move.develop.filter(Boolean).join('；') : '';
    return '先：' + move.open + '\n  - ' + develop + '\n  转折：' + move.turn + '\n  落点：' + move.land;
}

function renderPlan(plan) {
    const lines = [];
    const brief = plan.brief;
    lines.push('# 写作规划：' + brief.title);
    lines.push(plan.totalChars + ' 字左右｜' + GENRE_PROFILES[brief.genre].label + '｜读者：' + brief.audience + (brief.goal === '' ? '' : '｜目标：' + brief.goal));
    lines.push('主线：' + (brief.throughLine === '' ? '待补——写正文前用一句话写下「读者只记住的一句」' : brief.throughLine));
    lines.push('');
    lines.push('篇幅（按难度与读者收益分，不平均）：');
    lines.push(plan.sections.map((section) => section.heading + ' ' + Math.round(section.share * 100) + '%·' + section.targetChars + '字·' + section.depthLabel).join(' ｜ '));
    lines.push('');
    lines.push('全局质感：');
    lines.push('- 细节：' + plan.strategy.texture.anchor);
    lines.push('- 让步：' + plan.strategy.texture.turn);
    lines.push('- 落点：' + plan.strategy.texture.landing);
    lines.push('- 段落：' + plan.strategy.texture.paragraph);
    lines.push('- 连接：' + plan.strategy.texture.transition);
    lines.push('- 语气：' + plan.strategy.texture.uncertainty);
    lines.push('- 详略：' + plan.strategy.texture.detail);
    lines.push('- 节奏：' + plan.strategy.rhythm);
    lines.push('- 比喻：' + FIGURATIVE_LABELS[plan.strategy.figurative]);
    lines.push('- 不要：' + plan.strategy.antiPatterns.slice(0, 4).join(' '));
    for (const section of plan.sections) {
        lines.push('');
        lines.push('## ' + section.id.replace('s', '') + '. ' + section.heading + '（' + section.targetChars + ' 字，' + section.depthLabel + '）');
        lines.push('- 目的：' + section.purpose);
        lines.push('- 读者问题：' + section.readerQuestion);
        lines.push('- 必须写到：' + section.mustCover.join('；'));
        lines.push('- 写法：');
        lines.push('  - ' + joinMove(section.moves));
        lines.push('- 语言：' + section.registerLabel + '；比喻：' + section.figurativeLabel + '；节奏：' + section.rhythm);
        lines.push('- 不要：' + section.avoid.join('；'));
    }
    if (plan.checks.length > 0) {
        lines.push('');
        lines.push('写之前先确认：');
        for (const check of plan.checks) lines.push('- ' + check);
    }
    return lines.join('\n');
}

/** 把规划对象渲染成可读的 Markdown。 */
export function renderWritingPlan(plan) {
    if (plan !== null && typeof plan === 'object' && typeof plan.markdown === 'string' && plan.markdown !== '') return plan.markdown;
    if (plan !== null && typeof plan === 'object' && Array.isArray(plan.sections)) return renderPlan(plan);
    throw new Error('office.writing.render：需要一份 writingPlan 返回的规划对象。');
}

function stripMarkdown(text) {
    return String(text)
        .replace(/\x60\x60\x60[\s\S]*?\x60\x60\x60/gu, '')
        .replace(/!\[[^\]]*\]\([^)]*\)/gu, '')
        .replace(/\[([^\]]+)\]\([^)]*\)/gu, '$1')
        .replace(/^#{1,6}\s+.*$/gmu, '')
        .replace(/[*_~]{1,3}/gu, '')
        .replace(/^>\s?/gmu, '')
        .replace(/\u00a0/gu, ' ');
}

function splitParagraphs(text) {
    return stripMarkdown(text)
        .split(/\n\s*\n/gu)
        .map((paragraph) => paragraph.replace(/\n/gu, ' ').trim())
        .filter((paragraph) => countChars(paragraph) >= 2);
}

function splitSentences(text) {
    return stripMarkdown(text)
        .replace(/\s+/gu, ' ')
        .split(/(?<=[。！？!?…])/gu)
        .map((sentence) => sentence.trim())
        .filter((sentence) => countChars(sentence) >= 2 && !/^[-*+>#]/u.test(sentence));
}

function mean(values) {
    if (values.length === 0) return 0;
    return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function stddev(values) {
    if (values.length < 2) return 0;
    const average = mean(values);
    const variance = values.reduce((sum, value) => sum + (value - average) ** 2, 0) / values.length;
    return Math.sqrt(variance);
}

function cv(values) {
    const average = mean(values);
    return average === 0 ? 0 : stddev(values) / average;
}

function matchCount(text, pattern) {
    const re = new RegExp(pattern.source, pattern.flags.includes('g') ? pattern.flags : pattern.flags + 'g');
    const matches = text.match(re);
    return matches === null ? 0 : matches.length;
}

function topEntries(map, limit) {
    return [...map.entries()]
        .sort((left, right) => right[1] - left[1])
        .slice(0, limit)
        .map((entry) => ({ term: entry[0], count: entry[1] }));
}

/** 不用模型的可复核信号：从句长/段长/套话/细节/比喻/让步这些表面量算出来。 */
export function writingSignals(text) {
    const body = stripMarkdown(String(text ?? ''));
    const paragraphs = splitParagraphs(text);
    const paragraphLengths = paragraphs.map(countChars);
    const sentences = splitSentences(text);
    const sentenceLengths = sentences.map(countChars);
    const chars = countChars(body);
    const shortRatio = sentenceLengths.length === 0 ? 0 : sentenceLengths.filter((length) => length <= 12).length / sentenceLengths.length;
    const longRatio = sentenceLengths.length === 0 ? 0 : sentenceLengths.filter((length) => length >= 45).length / sentenceLengths.length;
    const clicheMatches = [];
    for (const item of CLICHE_PATTERNS) {
        const count = matchCount(body, item.re);
        if (count > 0) clicheMatches.push({ term: item.term, count });
    }
    const clicheTotal = clicheMatches.reduce((sum, item) => sum + item.count, 0);
    const openingMap = new Map();
    for (const sentence of sentences) {
        const opening = Array.from(sentence.replace(/[\s，,。！？!?、；;：:""''（）()]/gu, '')).slice(0, 4).join('');
        if (opening.length >= 2) openingMap.set(opening, (openingMap.get(opening) ?? 0) + 1);
    }
    const repeatedOpenings = topEntries(openingMap, 3).filter((item) => item.count >= 3);
    return {
        chars,
        paragraphs: paragraphs.length,
        sentences: sentences.length,
        sentenceLengths,
        paragraphLengths,
        meanSentence: Math.round(mean(sentenceLengths) * 10) / 10,
        sdSentence: Math.round(stddev(sentenceLengths) * 10) / 10,
        cvSentence: Math.round(cv(sentenceLengths) * 1000) / 1000,
        meanParagraph: Math.round(mean(paragraphLengths) * 10) / 10,
        cvParagraph: Math.round(cv(paragraphLengths) * 1000) / 1000,
        shortRatio: Math.round(shortRatio * 1000) / 1000,
        longRatio: Math.round(longRatio * 1000) / 1000,
        cliches: clicheMatches,
        clicheTotal,
        clicheDensity: chars === 0 ? 0 : Math.round((clicheTotal / chars) * 1000 * 10) / 10,
        metaphorHits: matchCount(body, METAPHOR_RE),
        hedgeHits: matchCount(body, HEDGE_RE),
        abstractHits: matchCount(body, ABSTRACT_RE),
        concreteHits: matchCount(body, /\d+(?:\.\d+)?/gu) + matchCount(body, EXAMPLE_RE),
        turnHits: matchCount(body, TURN_RE),
        summaryHits: matchCount(body, /综上所述|总而言之|总的来说|由此可见|总之/gu),
        questionHits: matchCount(body, /[？?]/gu),
        repeatedOpenings,
    };
}

function splitSections(text) {
    const sections = [];
    let current = { heading: '(全文)', body: [] };
    for (const line of String(text).split(/\r?\n/u)) {
        const match = /^#{1,6}\s+(.+?)\s*$/u.exec(line);
        if (match !== null) {
            if (current.body.join('').trim() !== '') sections.push({ heading: current.heading, text: current.body.join('\n') });
            current = { heading: match[1].trim(), body: [] };
        } else {
            current.body.push(line);
        }
    }
    if (current.body.join('').trim() !== '') sections.push({ heading: current.heading, text: current.body.join('\n') });
    return sections;
}

function matchPlannedSection(actual, planned) {
    const left = normalizeHeading(actual.heading);
    for (const section of planned) {
        const right = normalizeHeading(section.heading);
        if (left === right || left.includes(right) || right.includes(left)) return section;
    }
    return null;
}

function planOf(value) {
    if (value !== null && typeof value === 'object' && Array.isArray(value.sections) && value.brief !== undefined) return value;
    if (value !== null && typeof value === 'object' && Array.isArray(value.sections)) return value;
    return null;
}

function warning(code, message, section) {
    return { code, message, ...(section === undefined ? {} : { section }) };
}

/**
 * 自检一份正文。
 *
 * 第二个参数可以传 writingPlan 的返回值（会按节名对表），也可以省略。
 * warnings 是给模型改稿用的：每条都能落到一句具体动作，不是一个笼统分数。
 * score 只是这些信号的加权，不是「像人程度」的证明。
 */
export function writingAudit(text, planOrOptions) {
    if (typeof text !== 'string' || text.trim() === '') {
        throw new Error('office.writing.audit：没有正文可查；传完整的 Markdown 或纯文本。');
    }
    const plan = planOf(planOrOptions);
    const signals = writingSignals(text);
    const actualSections = splitSections(text).map((section) => {
        const chars = countChars(section.text);
        return { heading: section.heading, chars, text: section.text };
    });
    const warnings = [];
    const suggestions = [];
    let score = 100;

    if (signals.chars < 300) {
        warnings.push(warning('short-text', '正文不足 300 字，句长/段长信号不稳定；这一条只做提示，不扣分。'));
    }

    if (signals.sentences >= 8 && signals.cvSentence < 0.28) {
        score -= 14;
        warnings.push(warning('uniform-sentences', '句长几乎一样（平均 ' + signals.meanSentence + ' 字，变异系数 ' + signals.cvSentence + '）：读起来像同一条流水线。'));
        suggestions.push('挑连续三句同长的，拆成「长—短—中」；至少让一句落到 12 字以内。');
    } else if (signals.sentences >= 8 && signals.cvSentence < 0.38) {
        score -= 7;
        warnings.push(warning('flat-rhythm', '句长偏匀（变异系数 ' + signals.cvSentence + '）：还可以再拉开一点。'));
    }
    if (signals.longRatio > 0.45) {
        score -= 5;
        warnings.push(warning('long-sentences', '45 字以上长句占 ' + Math.round(signals.longRatio * 100) + '%：读者会断气。'));
    }
    if (signals.shortRatio < 0.08 && signals.meanSentence > 30) {
        score -= 5;
        warnings.push(warning('no-short-sentences', '几乎没有 12 字以内的短句：关键判断没有落地的地方。'));
    }
    if (signals.paragraphs >= 4 && signals.cvParagraph < 0.25) {
        score -= 8;
        warnings.push(warning('uniform-paragraphs', '段落长度几乎一样（变异系数 ' + signals.cvParagraph + '）：该短的地方没短，该展开的地方没展开。'));
    }
    if (signals.clicheTotal >= 3 && signals.clicheDensity > 1.0) {
        const penalty = Math.min(20, 6 + signals.clicheTotal * 2);
        score -= penalty;
        warnings.push(warning('cliches', '套话/模板连接词 ' + signals.clicheTotal + ' 处（' + signals.cliches.slice(0, 3).map((item) => item.term).join('、') + '）：换成一个具体动作或事实。'));
        suggestions.push('先把「首先/其次/最后」「综上所述」「值得注意的是」这类骨架删掉，看段落还能不能靠内容接上。');
    }
    if (signals.repeatedOpenings.length > 0 && signals.repeatedOpenings[0].count >= 4) {
        score -= 6;
        warnings.push(warning('repeated-openings', '有 ' + signals.repeatedOpenings[0].count + ' 句用同样的开头「' + signals.repeatedOpenings[0].term + '」：换主语或从句中段起句。'));
    }
    if (signals.abstractHits >= 8 && signals.concreteHits <= 2) {
        score -= 6;
        warnings.push(warning('abstract-heavy', '抽象名词（性/化/机制/逻辑）' + signals.abstractHits + ' 处，具体数字或例子只有 ' + signals.concreteHits + ' 处：读者抓不住。'));
        suggestions.push('每个抽象判断后面补一个数字、场景或反例；补不出来就说明它只是推测。');
    }
    if (signals.chars >= 600) {
        const anchorNeed = Math.max(1, Math.ceil(signals.chars / 150));
        if (signals.concreteHits < anchorNeed) {
            score -= 5;
            warnings.push(warning('anchor-sparse', '具体锚点（数字 / 例子）' + signals.concreteHits + ' 处，按每 150 字一个算应有 ' + anchorNeed + ' 处：抽象判断没有落点。'));
            suggestions.push('每 150 字补一个可核对的锚点：数字带口径、专名、日期、原话或具体场景；补不出来就写「这是推测」。');
        }
    }
    if (signals.metaphorHits > Math.max(3, Math.ceil(signals.chars / 400))) {
        score -= 5;
        warnings.push(warning('metaphor-overuse', '比喻 ' + signals.metaphorHits + ' 处，密度偏高：比喻要兑现成机制，不要连续换喻体。'));
    }
    if (signals.hedgeHits > 0 && (signals.hedgeHits / Math.max(1, signals.chars)) * 1000 > 8) {
        score -= 4;
        warnings.push(warning('over-hedging', '「可能/或许/一定程度上」密度偏高：该下判断的地方下判断，把依据放在旁边。'));
    }
    if (signals.chars >= 600 && signals.turnHits === 0) {
        score -= 8;
        warnings.push(warning('no-turn', '通篇顺推，没有一处让步、反例或竞争解释：读起来像在说服而不是在思考。'));
        suggestions.push('在论证最顺的地方加一句「不过，有一种情况会推翻它」，再把它接住。');
    }
    if (signals.chars >= 600 && signals.summaryHits >= 3) {
        score -= 5;
        warnings.push(warning('summary-repeat', '总结性连接词出现 ' + signals.summaryHits + ' 次：每节都收一次，读者会觉得在原地走。'));
    }
    if (actualSections.length >= 2) {
        const summaryTails = actualSections.filter((section) => /(?:综上所述|总而言之|总的来说|总之)[^。！？!?]*[。！？!?]?\s*$/u.test(section.text.trim()));
        if (summaryTails.length >= 2) {
            score -= 4;
            warnings.push(warning('section-summary', '有 ' + summaryTails.length + ' 个小节在结尾做总结（' + summaryTails.slice(0, 2).map((section) => section.heading).join('、') + '）：复述不增量，只留带条件 / 代价 / 下一步的那一句。'));
        }
    }

    const planSections = [];
    if (plan !== null) {
        let matched = 0;
        const matchedPairs = [];
        for (const section of actualSections) {
            const planned = matchPlannedSection(section, plan.sections);
            if (planned === null) {
                planSections.push({ heading: section.heading, chars: section.chars, planned: null });
                continue;
            }
            matched += 1;
            const target = Number(planned.targetChars) || 0;
            const deviation = target > 0 ? (section.chars - target) / target : 0;
            const entry = {
                heading: section.heading,
                plannedHeading: planned.heading,
                chars: section.chars,
                targetChars: target,
                share: planned.share,
                depth: planned.depth,
                deviation: Math.round(deviation * 1000) / 1000,
            };
            planSections.push(entry);
            matchedPairs.push({ entry, planned });
            if (planned.depth === 'dwell' || planned.depth === 'develop') {
                if (section.chars < target * 0.68) {
                    score -= 8;
                    warnings.push(warning('underdeveloped', '「' + section.heading + '」按规划是' + planned.depthLabel + '（约 ' + target + ' 字），实际 ' + section.chars + ' 字：重点被一笔带过。', section.heading));
                    suggestions.push('「' + section.heading + '」至少扩到 ' + Math.round(target * 0.9) + ' 字：补机制链条、一个具体例子，或一个反例。');
                }
            }
            if (planned.depth === 'compress' && section.chars > target * 1.7) {
                score -= 4;
                warnings.push(warning('overexpanded', '「' + section.heading + '」按规划该收住（约 ' + target + ' 字），实际 ' + section.chars + ' 字：铺垫抢了重点的篇幅。', section.heading));
                suggestions.push('「' + section.heading + '」压到 ' + Math.round(target * 1.2) + ' 字以内：只留读者不知道的那部分。');
            }
        }
        if (matched === 0) {
            warnings.push(warning('unmatched-plan', '正文的小标题与规划里的节名对不上：要么按规划改标题，要么把规划当成另一份稿子的。'));
        } else if (matchedPairs.length >= 3) {
            const plannedShares = matchedPairs.map((pair) => Number(pair.planned.share) || 0);
            const actualChars = matchedPairs.map((pair) => pair.entry.chars);
            const plannedSpread = Math.max(...plannedShares) / Math.max(0.001, Math.min(...plannedShares));
            const actualSpread = Math.max(...actualChars) / Math.max(1, Math.min(...actualChars));
            if (plannedSpread >= 1.5 && actualSpread < 1.2) {
                score -= 10;
                warnings.push(warning('uniform-allocation', '规划里各节篇幅差 ' + plannedSpread.toFixed(1) + ' 倍，正文实际只差 ' + actualSpread.toFixed(2) + ' 倍：又写成了平均用力。'));
                suggestions.push('把规划里占比最低的节压掉三成，把最高的那一节扩到目标字数；不要靠加小标题解决。');
            }
            const longest = matchedPairs.reduce((left, right) => (right.entry.chars > left.entry.chars ? right : left));
            const hardest = matchedPairs.reduce((left, right) => (
                (right.planned.difficulty * 1.2 + right.planned.value) > (left.planned.difficulty * 1.2 + left.planned.value) ? right : left
            ));
            const softKinds = new Set(['context', 'summary', 'example', 'faq']);
            if (longest.entry.heading !== hardest.entry.heading
                && (softKinds.has(longest.planned.kind) || longest.entry.chars > hardest.entry.chars * 1.15)) {
                score -= 8;
                warnings.push(warning('longest-not-hardest', '最长的一节是「' + longest.entry.heading + '」（' + longest.entry.chars + ' 字），最难 / 最值钱的是「' + hardest.entry.heading + '」：笔墨给错了地方。', longest.entry.heading));
                suggestions.push('把「' + longest.entry.heading + '」压到能说清就停，省下的字数扩「' + hardest.entry.heading + '」的机制、反例或边界。');
            }
        }
        const unmatched = planSections.filter((section) => section.planned === null).length;
        if (unmatched > 0 && matched > 0) suggestions.push('有 ' + unmatched + ' 个小标题没对上规划：检查是不是把规划节名改了却没改目标。');
    }

    score = Math.max(0, Math.min(100, Math.round(score)));
    const band = score >= 82 ? '自然' : score >= 68 ? '尚可，仍有机器味' : '模板感重';
    const uniqueSuggestions = [...new Set(suggestions)].slice(0, 6);
    return {
        ok: true,
        score,
        band,
        chars: signals.chars,
        signals,
        sections: planSections.length > 0 ? planSections : actualSections.map((section) => ({ heading: section.heading, chars: section.chars })),
        warnings,
        suggestions: uniqueSuggestions,
        note: 'score 是启发式信号加权，不是「像人程度」的证明；优先看 warnings 与 suggestions。',
    };
}

// ── 去 AI 味报告（第七十五轮）──────────────────────────────────────────────
// 只报可解释的句式与文体偏离：没有分数，不推断作者，也不改事实。
// A1 对举、A2 提示语冒号、M1 元话语起手式是默认层；F1/A4/A5/B1/C2 只进 stats，
// 因为上一轮标定里它们在人类对照上方向反或过火（见 docs/办公模式第七十四轮-标定结果与规则修正）。
const DEAI_A1_RE = /(?:不是|并非)[^。！？\n]{1,24}?(?:而是|是)/gu;
const DEAI_A2_RE = /(?:原因|结论|核心|关键|问题|答案|本质|真相|重点|逻辑|前提|事实)(?:不复杂|很简单|只有一个|是|在于)?[：:]/gu;
const DEAI_M1_RE = /(?:先说|直接说|简单说|一句话|说白了|说穿了)[^。！？\n：:]{0,8}[：:]/gu;
const DEAI_F1_RE = /——/gu;
const DEAI_A4_RE = /[\u4e00-\u9fa5]{12,}的[\u4e00-\u9fa5]{2,6}/gu;
const DEAI_A5_RE = /、/gu;
const DEAI_C2_RE = /(?:^|[。！？\n])\s*(?:因此|所以|然而|但是|此外|同时|总之|综上|换句话说|值得注意的是)/gu;

const DEAI_REGISTERS = {
    general: {
        label: '通用',
        note: '没有指定文体：按通用口径报 A1/A2/M1；要更准请传 register（platform/official/academic/personal）。',
        levels: { A1: 'high', A2: 'high', M1: 'high' },
    },
    platform: {
        label: '平台体/自媒体',
        note: '平台体常见先结论、数字、个人经历；A1/A2/M1 都按高置信报，但仍不是 AI 检测。',
        levels: { A1: 'high', A2: 'high', M1: 'high' },
    },
    official: {
        label: '公文/汇报',
        note: '公文常用「原因在于：」等提示语：A2 只作提示，不当作 AI 证据；元话语起手式仍可疑。',
        levels: { A1: 'medium', A2: 'low', M1: 'high' },
    },
    academic: {
        label: '学术',
        note: '学术里的「不是 A 而是 B」可能是真论证：A1 中置信，A2 只作提示。',
        levels: { A1: 'medium', A2: 'low', M1: 'medium' },
    },
    personal: {
        label: '个人随笔',
        note: '随笔里的对举和口语不是错误：A1 只作提示，重点看提示语与文体偏离。',
        levels: { A1: 'low', A2: 'medium', M1: 'high' },
    },
};

const DEAI_REGISTER_ALIASES = {
    platform: 'platform', 平台: 'platform', 平台体: 'platform', 自媒体: 'platform',
    official: 'official', 公文: 'official', 汇报: 'official',
    academic: 'academic', 学术: 'academic',
    personal: 'personal', 随笔: 'personal',
    general: 'general', 通用: 'general',
};

const DEAI_METRIC_SPECS = [
    { key: 'meanSentence', label: '平均句长', unit: '字', abs: 4, rel: 0.35 },
    { key: 'cvSentence', label: '句长变异', unit: '', abs: 0.12, rel: 0.35 },
    { key: 'meanParagraph', label: '平均段长', unit: '字', abs: 30, rel: 0.35 },
    { key: 'cvParagraph', label: '段长变异', unit: '', abs: 0.15, rel: 0.35 },
    { key: 'shortRatio', label: '短句占比', unit: '', abs: 0.08, rel: 0.4 },
    { key: 'a1Density', label: '对举密度', unit: '/千字', abs: 2, rel: 0.5 },
    { key: 'promptColonDensity', label: '提示语冒号密度', unit: '/千字', abs: 2, rel: 0.5 },
    { key: 'metaDensity', label: '元话语密度', unit: '/千字', abs: 1, rel: 0.5 },
    { key: 'clicheDensity', label: '套话密度', unit: '/千字', abs: 2, rel: 0.5 },
    { key: 'specificityDensity', label: '具体锚点密度', unit: '/千字', abs: 2, rel: 0.5 },
    { key: 'questionDensity', label: '问句密度', unit: '/千字', abs: 2, rel: 0.5 },
];

const DEAI_SAFE_RULES = [
    {
        code: 'M1',
        re: /(^|[。！？\n]\s*)(先说结论|直接说|简单说|一句话|说白了|说穿了)\s*[：:]\s*/gu,
        replace: (match) => match[1] ?? '',
    },
    {
        code: 'A2',
        re: /(^|[。！？\n，,]\s*)(原因不复杂|原因很简单)\s*[：:]\s*/gu,
        replace: (match) => (match[1] ?? '') + '原因是',
    },
    {
        code: 'A2',
        re: /(^|[。！？\n，,]\s*)(这个原因)\s*[：:]\s*/gu,
        replace: (match) => (match[1] ?? '') + '这个原因，',
    },
];

function maskForDeai(text) {
    const blank = (match) => match.replace(/[^\n]/gu, ' ');
    return String(text)
        .replace(/\x60\x60\x60[\s\S]*?\x60\x60\x60/gu, blank)
        .replace(/\x60[^\x60\n]*\x60/gu, blank)
        .replace(/https?:\/\/[^\s)）]+/gu, blank);
}

function deaiSpans(text, pattern) {
    const re = new RegExp(pattern.source, pattern.flags.includes('g') ? pattern.flags : pattern.flags + 'g');
    return [...text.matchAll(re)].map((match) => ({
        text: match[0],
        start: match.index,
        end: match.index + match[0].length,
    }));
}

function perK(count, chars) {
    const total = Math.max(0, Number(chars) || 0);
    return total === 0 ? 0 : Math.round((count / total) * 1000 * 10) / 10;
}

function deaiStats(text) {
    const signals = writingSignals(text);
    const body = maskForDeai(stripMarkdown(text));
    const metaSpans = deaiSpans(body, DEAI_M1_RE);
    const promptSpans = deaiSpans(body, DEAI_A2_RE).filter((span) => (
        !metaSpans.some((meta) => span.start < meta.end && span.end > meta.start)
    ));
    const counts = {
        a1: matchCount(body, DEAI_A1_RE),
        a2: promptSpans.length,
        m1: metaSpans.length,
        f1: matchCount(body, DEAI_F1_RE),
        a4: matchCount(body, DEAI_A4_RE),
        a5: matchCount(body, DEAI_A5_RE),
        b1: matchCount(body, ABSTRACT_RE),
        c2: matchCount(body, DEAI_C2_RE),
    };
    const chars = signals.chars;
    return {
        ...signals,
        ...counts,
        a1Density: perK(counts.a1, chars),
        a2Density: perK(counts.a2, chars),
        m1Density: perK(counts.m1, chars),
        f1Density: perK(counts.f1, chars),
        a4Density: perK(counts.a4, chars),
        a5Density: perK(counts.a5, chars),
        b1Density: perK(counts.b1, chars),
        c2Density: perK(counts.c2, chars),
        promptColonDensity: perK(counts.a2 + counts.m1, chars),
        metaDensity: perK(counts.m1, chars),
        specificityDensity: perK(signals.concreteHits, chars),
        questionDensity: perK(signals.questionHits, chars),
    };
}

function compactDeaiStats(stats) {
    return {
        chars: stats.chars,
        sentences: stats.sentences,
        paragraphs: stats.paragraphs,
        meanSentence: stats.meanSentence,
        cvSentence: stats.cvSentence,
        meanParagraph: stats.meanParagraph,
        cvParagraph: stats.cvParagraph,
        shortRatio: stats.shortRatio,
        longRatio: stats.longRatio,
        a1: stats.a1,
        a2: stats.a2,
        m1: stats.m1,
        f1: stats.f1,
        a4: stats.a4,
        a5: stats.a5,
        b1: stats.b1,
        c2: stats.c2,
        a1Density: stats.a1Density,
        a2Density: stats.a2Density,
        m1Density: stats.m1Density,
        f1Density: stats.f1Density,
        promptColonDensity: stats.promptColonDensity,
        metaDensity: stats.metaDensity,
        clicheDensity: stats.clicheDensity,
        specificityDensity: stats.specificityDensity,
        questionDensity: stats.questionDensity,
    };
}

function deaiStyleRef(body, refText) {
    const ref = deaiStats(refText);
    const weak = ref.chars < 200;
    const specs = weak
        ? DEAI_METRIC_SPECS.filter((spec) => ['meanSentence', 'cvSentence', 'meanParagraph', 'cvParagraph', 'shortRatio'].includes(spec.key))
        : DEAI_METRIC_SPECS;
    const metrics = [];
    for (const spec of specs) {
        const left = Number(body[spec.key]) || 0;
        const right = Number(ref[spec.key]) || 0;
        const delta = left - right;
        const relative = Math.abs(delta) / Math.max(Math.abs(right), spec.abs / 2, 0.001);
        if (Math.abs(delta) < spec.abs || relative < spec.rel) continue;
        const percent = right === 0 ? '' : '（' + (delta >= 0 ? '高' : '低') + ' ' + Math.min(999, Math.round(relative * 100)) + '%）';
        metrics.push({
            key: spec.key,
            label: spec.label,
            unit: spec.unit,
            refValue: right,
            bodyValue: left,
            delta: Math.round(delta * 1000) / 1000,
            relative: Math.round(relative * 1000) / 1000,
            direction: delta >= 0 ? 'higher' : 'lower',
            note: spec.label + '：正文 ' + left + (spec.unit || '') + ' vs 参考 ' + right + (spec.unit || '') + percent,
        });
    }
    metrics.sort((left, right) => right.relative - left.relative);
    return {
        provided: true,
        chars: ref.chars,
        weak,
        metrics: metrics.slice(0, 6),
        note: weak
            ? '参考稿不足 200 字：只比句长/段长方向，不要当阈值。'
            : '只报与参考稿的偏离，不报 AI 概率；参考稿应是同一文体的人类稿。',
    };
}

function deaiSafeEdits(text) {
    const masked = maskForDeai(text);
    const found = [];
    for (const rule of DEAI_SAFE_RULES) {
        const re = new RegExp(rule.re.source, rule.re.flags);
        for (const match of masked.matchAll(re)) {
            const from = text.slice(match.index, match.index + match[0].length);
            const to = rule.replace(match);
            if (from !== to) {
                found.push({ code: rule.code, from, to, start: match.index, end: match.index + match[0].length });
            }
        }
    }
    const sorted = found.sort((left, right) => left.start - right.start || right.end - left.end);
    const edits = [];
    for (const edit of sorted) {
        const previous = edits[edits.length - 1];
        if (previous === undefined || edit.start >= previous.end) edits.push(edit);
    }
    let result = text;
    for (const edit of [...edits].sort((left, right) => right.start - left.start)) {
        result = result.slice(0, edit.start) + edit.to + result.slice(edit.end);
    }
    return { text: result, edits };
}

/**
 * 去 AI 味报告：只列可解释的句式与文体偏离。
 *
 * 没有分数、没有「AI 含量%」、不推断作者。A1 对举 / A2 提示语冒号 / M1 元话语
 * 起手式是默认可疑项；F1 破折号、A4 长定语、A5 顿号、B1 抽象、C2 连接词只进
 * stats，因为标定里它们在人类稿上也会命中、方向不稳。
 *
 * options：
 *   mode: 'report'（默认，只报）| 'safe'（只删/改提示语前缀，不动论证）
 *   register: 'platform' | 'official' | 'academic' | 'personal' | 'general'
 *   styleRef: 一段人类参考稿；给了就报「与参考稿的偏离」，不报 AI 概率
 */
export function writingDeai(text, options = {}) {
    if (typeof text !== 'string' || text.trim() === '') {
        throw new Error('office.writing.deai：没有正文可查；传完整的 Markdown 或纯文本。');
    }
    const opts = typeof options === 'string' ? { register: options } : (options ?? {});
    const register = normalizeDeaiRegister(opts.register);
    const mode = opts.mode === 'safe' ? 'safe' : 'report';
    const stats = deaiStats(text);
    const masked = maskForDeai(text);
    const metaSpans = deaiSpans(masked, DEAI_M1_RE);
    const promptSpans = deaiSpans(masked, DEAI_A2_RE).filter((span) => (
        !metaSpans.some((meta) => span.start < meta.end && span.end > meta.start)
    ));
    const contrastSpans = deaiSpans(masked, DEAI_A1_RE);
    const flags = [];
    const notes = [register.note];
    const consider = (code, label, spans, reason, falsePositiveHint) => {
        if (spans.length === 0) return;
        const confidence = register.levels[code] ?? 'medium';
        const density = code === 'A1' ? stats.a1Density : code === 'A2' ? stats.a2Density : stats.m1Density;
        if (confidence === 'high' || confidence === 'medium') {
            flags.push({
                code,
                label,
                count: spans.length,
                density,
                confidence,
                spans: spans.slice(0, 6),
                reason,
                falsePositiveHint,
            });
            return;
        }
        notes.push(label + ' ' + spans.length + ' 处（密度 ' + density + '/千字）：' + falsePositiveHint);
    };
    consider(
        'A1',
        '对举/翻案腔',
        contrastSpans,
        '「不是 A 而是 B」把对比写成了固定骨架；连续出现时像模板。',
        '学术/评论里的真对比也可能用这个句式；看是否连续出现、是否只是空对比。',
    );
    consider(
        'A2',
        '提示语冒号',
        promptSpans,
        '先宣布「这是什么」再给内容；平台体常见，读完容易觉得在听讲解。',
        '公文/教程里的「原因在于：」是正常结构；official/academic 已自动降为提示。',
    );
    consider(
        'M1',
        '元话语起手式',
        metaSpans,
        '「先说结论/说白了」先讲写法再讲内容，是典型的口播/平台体起手。',
        '演讲或口语稿也可能这样起；看是否每段都用。',
    );
    if (stats.f1 > 0) notes.push('破折号 ' + stats.f1 + ' 处：模型指纹，方向随模型与文体变，只进 stats。');
    if (stats.a4 > 0) notes.push('长前置定语 ' + stats.a4 + ' 处：人类公文也常见，只进 stats。');
    if (stats.a5 > 0) notes.push('顿号 ' + stats.a5 + ' 处：人类并列往往更多，只进 stats。');
    if (stats.b1 > 0) notes.push('抽象名词 ' + stats.b1 + ' 处：人类说明文也会用，只进 stats。');
    if (stats.c2 > 0) notes.push('段首连接词 ' + stats.c2 + ' 处：弱信号，只作风格提示。');

    let styleRef = { provided: false };
    if (typeof opts.styleRef === 'string' && opts.styleRef.trim() !== '') {
        styleRef = deaiStyleRef(stats, opts.styleRef);
    }
    const safe = mode === 'safe' ? deaiSafeEdits(text) : null;
    return {
        ok: true,
        mode,
        register: { id: register.id, label: register.label },
        chars: stats.chars,
        flags,
        stats: compactDeaiStats(stats),
        styleRef,
        ...(safe === null ? {} : { text: safe.text, edits: safe.edits }),
        notes,
        note: '只报可解释的句式与文体偏离，不是 AI 检测器，也没有分数；人类也会写 A1/A2。改稿先看 flags/spans。',
    };
}

function normalizeDeaiRegister(value) {
    const key = String(value ?? 'general').trim().toLowerCase();
    const id = DEAI_REGISTER_ALIASES[key] ?? (DEAI_REGISTERS[key] !== undefined ? key : 'general');
    return { id, ...DEAI_REGISTERS[id] };
}
