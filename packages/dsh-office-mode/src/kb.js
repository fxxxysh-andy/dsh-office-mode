/**
 * 知识库（kb）—— 记忆之外的第二棵树：**来源文档**的结构化切块与可回溯指针。
 *
 * 为什么不是第四层记忆：第三层（热记忆 / 台账 / 归档）管的是「生命周期」，
 * 装的都是**本插件自己产出的**结论与登记；而这里的每一块都是**外部原文**，
 * 它不该按生命周期下沉，也不该被摘要化 —— 引用必须落回原文的字符区间。
 * 第十九轮定的三条线（台账 → 块、记忆 → 块、块 ↔ 台账/记忆）都靠这个 id 系上。
 *
 * 存储（`.office/kb/`，与 `.office/memory` 同级的兄弟目录）：
 *
 *   manifest.jsonl              一文档一行：相对路径、内容哈希、字节、块数、来源档、入库时间
 *   chunks/<docHash>/<n>.json   一块一文件：**只放与这一块有关的东西**（正文、字符区间
 *                               span、标题路径、自校验哈希）；来源路径 / 入库时刻 / 来源档
 *                               是文档级的，从 manifest 现取
 *   index.json                  计数与体积（给面板容量条；**真源是前两份**，它只是投影）
 *
 * 一条容易踩的纪律：**`docHash` 里包含切块档**（`kbHash(text, profile)`）。切块是
 * 「内容 + 档位」的函数，块 id 又是 `kb:<哈希>:<块号>` —— 哈希里不带档位，两份边界不同的
 * 块就会共用一个 id 空间（第四十一轮独立验证 F3 实测：同一份内容从 `.office/search/`（900 档）
 * 与普通路径（1200 档）各入库一次，读前者拿到后者的边界，manifest 说 4 块而盘上 5 个文件）。
 * 同一份内容被两份**路径不同**的文档入库时仍然共用一个块目录 —— 所以块文件里不许写
 * 文档级元数据（否则互相覆盖，读 a.md 的块头会写着「来源 b.md」）。
 *
 * 四条纪律（都有第二十四轮的出处，别再重新论证）：
 *   1. **原文只存指针，但 span 必须存**。大文件不复制进 kb；引用落到字符区间 ——
 *      把来源压成摘要后引用精度对摘要 0.86、对源 span 只有 0.12（2609.14245），
 *      而 CAMS 靠「claim → 逐字引用 → token span」把多源归因从 38% 提到 64%。
 *   2. **块头拼「标题路径 + 来源 + 日期」**。这是 Anthropic Contextual Retrieval 的
 *      零依赖版本（官方数字 −35% / −49% / −67% 全部来自「给块前置 50—100 token
 *      专属上下文」）；官方同时否定「给块加通用文档摘要」与「摘要式索引」。
 *   3. **切块是结构切块**（按 Markdown 标题分层 + 段落聚合），不做语义切块：
 *      2410.13070 在真实文档上定长常优于语义切块，而语义切块还要引入嵌入。
 *      块大小**按来源分档**（2505.21700：最优块大小随数据集从 64 变到 1024）。
 *   4. **切块必须确定性**：同一份正文任何时候切出同一批块（同 id、同 span、
 *      同正文），否则「重复入库幂等」与「按 id 引用」都不成立。块文件另带一个
 *      自校验哈希（正文与它不等时按「缺块」处理 / 导入时整份拒收）。
 *
 * 检索（第二十四轮 24-11 接线）：词法的多路 BM25（正文 + 块头），中文按 bigram 展开。
 * 一期的「显式挑读」仍然成立（入库、清单、有界整份读都不变），检索只是多了一条入口 ——
 * 语料规模的那条线（10k token 全给更好、100k 才反转）由「命中只给预览 + 有界扫描」兜住。
 * 打分器是纯函数（`kbRankChunks`），所以探针与 action 面量的是**同一份**实现。
 *
 * @module dsh-office-mode/kb
 */
import { createHash } from 'node:crypto';
import { dirname, join, posix } from 'node:path';

/** 存储格式版本；将来改结构时用它做迁移判断。 */
export const KB_VERSION = 1;
/** kb 目录名：与记忆目录同级的兄弟目录（`.office/kb`）。 */
export const KB_DIR_NAME = 'kb';
/**
 * 来源三档（第二十四轮 §5.3 的硬规则，**不是**可信度权重）。
 *
 *   user        用户放进工作目录的文档（第十三轮信任边界）：可进 kb，可被引用
 *   verified    ≥2 个独立来源确认过的簇：可进 kb，可被引用
 *   unverified  单源检索结果：**只进 kb 的待核区**，永不进热记忆
 *
 * 为什么必须是硬规则：Utility Under Attack 的实测是「加性 provenance 权重与完全
 * 无防御不可区分（p = 0.80）」，而按来源类别**排除**不可信内容才有效
 * （0.3167 → 0.7000）。同一篇还给了方向：写入期设防没有可测的良性代价，
 * 读时过滤要付 4.4 个准确率点并误隔离 33.6% 的合法记忆 —— 所以安全设在写侧。
 */
export const KB_TIERS = ['user', 'verified', 'unverified'];
export const DEFAULT_KB_TIER = 'user';
/** 单块字符数上限（按来源分档：检索结果文件比正经文档碎，块要小一些）。 */
export const KB_CHUNK_CHARS = 1200;
export const KB_CHUNK_CHARS_BY_KIND = { doc: 1200, search: 900 };
/** 单文档字节上限：超过就拒绝入库（不复制大文件进 kb，也不静默截断）。 */
export const KB_MAX_DOC_BYTES = 2 * 1024 * 1024;
/** 「有界整份读」的字符额度：与 READ_LIMITS.totalChars 同一量级。 */
export const KB_READ_CHARS = 4000;
/** 一次 kb-list / kb-read 最多给几条 / 几块。 */
export const KB_LIST_LIMIT = 20;
export const KB_READ_CHUNKS = 12;

// ── 检索（第二十四轮 24-11 的接线面） ────────────────────────────────────────
//
// 一期（第四十一轮）只做「显式入库 + 清单 + 有界整份读」；检索器等自测集先变绿再启用。
// 第四十二轮交出了那套自测集（`test/memory-probes.mjs` 的 C 组：15 条标注查询上候选
// 打分器 15/15、朴素整串只有 5/15），本轮把它搬进这里并接进 `office_memory`。
//
// 三条口径（都有出处，别再重新论证）：
//   1. **词法检索，不做嵌入**。中文按字符二元组展开（LeCaRDv2：完全不切 nDCG@10 0.359
//      ＜ 字符级 0.567；SQLite FTS5 对中文基本失效），拉丁词与数字整词保留。
//   2. **BM25 排序**，k1 = 1.2 / b = 0.75 —— 先按词频饱和，再按块长度归一。参数不暴露
//      成旋钮：它们是在自测集上量过的，不该被随手拧。
//   3. **多路按归一化加权和融合**：正文一路、块头一路（标题路径 + 来源 + 日期，就是
//      Contextual Retrieval 的零依赖版本）。每路先各自归一化到 0..1 再加权，
//      否则「块头短、词频高」的那一路会把正文那一整片盖过去。
export const KB_SEARCH_LIMIT = 8;
/** 一次检索最多读多少块正文（多路打分的语料有界：语料再大也不会变成全库扫描）。 */
export const KB_SEARCH_CHUNKS = 2000;
/** BM25 的两个参数（自测集上量出来的那两个数，不出旋钮）。 */
export const KB_BM25 = { k1: 1.2, b: 0.75 };
/** 多路融合的权重：正文为主、块头为辅。 */
export const KB_ROUTE_WEIGHTS = { text: 0.75, head: 0.25 };
/** 一条命中给多少字符的预览（块正文由 kb-read 按需整块读）。 */
export const KB_HIT_PREVIEW = 280;

function asText(value) {
    return typeof value === 'string' ? value.trim() : '';
}

function finiteNumber(value) {
    const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
    return Number.isFinite(n) ? n : 0;
}

function boundedInt(value, fallback, min, max) {
    const n = typeof value === 'number' ? Math.trunc(value) : Number.parseInt(String(value), 10);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

/**
 * 内容哈希（sha1 前 12 位十六进制）：块 id 与文档版本都用它。
 *
 * `profile` 一定要传（块大小档）：**切块是内容 + 档位的函数**。同内容在文档档（1200）与
 * 检索结果档（900）下切出来的块数与边界不同，而块 id 是 `kb:<哈希>:<块号>` 形式 ——
 * 哈希里不带档位，两份不同边界的块就会共用一个 id 空间（第四十一轮独立验证 F3 实测：
 * 同一份内容从 `.office/search/` 与普通路径各入库一次，读前者拿到后者的边界，
 * manifest 说 4 块而盘上 5 个文件）。带上档位之后，「同一个哈希」就一定「同一批块」。
 */
export function kbHash(text, profile = '') {
    return createHash('sha1').update(`${String(profile ?? '')}\u0000${String(text ?? '')}`, 'utf8').digest('hex').slice(0, 12);
}

/**
 * kb 目录：记忆目录的**兄弟目录**。
 *
 * 为什么跟着记忆目录走而不是写死 `.office/kb`：用户把 `memory.dir` 改成别处时，
 * 两份数据不该一个跟走、一个留在原地（那会变成「记忆搬了、来源还指着旧目录」）。
 */
export function kbDirOf(memoryDir) {
    return join(dirname(memoryDir), KB_DIR_NAME);
}

/** kb 一棵树的文件位置（manifest 是真源，index.json 只是投影）。 */
export function kbPathsFor(memoryDir) {
    const dir = kbDirOf(memoryDir);
    return {
        dir,
        manifest: join(dir, 'manifest.jsonl'),
        chunks: join(dir, 'chunks'),
        index: join(dir, 'index.json'),
    };
}

/** 文档 id：按**路径**寻址（同一份文件重复入库要能认出来是同一份）。 */
export function docIdOf(relPath) {
    return `kb:${kbHash(asText(relPath))}`;
}

/** 块 id：按**内容**寻址（同一个 docHash 切出的第 n 块永远是同一个 id）。 */
export function chunkIdOf(docHash, n) {
    return `kb:${asText(docHash)}:${boundedInt(n, 0, 0, 1_000_000)}`;
}

/** 块文件名（相对 chunks/ 目录，posix 写法，便于进 Pack 的键）。 */
export function chunkRelOf(docHash, n) {
    return posix.join(asText(docHash), `${boundedInt(n, 0, 0, 1_000_000)}.json`);
}

/** 来源类别收敛：不认识的值按 user 处理（写坏配置不该产生一个不存在的档）。 */
export function kbTierOf(value) {
    const wanted = asText(value).toLowerCase();
    return KB_TIERS.includes(wanted) ? wanted : DEFAULT_KB_TIER;
}

/**
 * 显式传入的来源档：**认不出来就报错**，不静默退回。
 *
 * 与 `kbTierOf` 的分工是刻意的：读落盘数据要宽松（老文件、外部改过的行不该读不动），
 * 而调用方写进来的 `tier`（`add` / `kb-ingest`）必须看得懂 —— 拼错一个字母就静默降档，
 * 等于把「已验证」悄悄变成「用户提供」，那是看不见的档位漂移（独立验证 F13 实测）。
 */
export function strictTierOf(value, where = 'tier') {
    const wanted = asText(value).toLowerCase();
    if (wanted === '') return '';
    if (!KB_TIERS.includes(wanted)) {
        throw new Error(`${where} 只支持 ${KB_TIERS.join(' / ')}，收到「${asText(value)}」。`);
    }
    return wanted;
}

/** 两处对同一份来源的档位说法取**更保守**的那个（顺序：verified > user > unverified）。 */
export function strictestTier(left, right) {
    const rank = { verified: 2, user: 1, unverified: 0 };
    const a = kbTierOf(left);
    const b = kbTierOf(right);
    return (rank[a] ?? 1) <= (rank[b] ?? 1) ? a : b;
}

/** 按来源分档的块大小：`.office/search/` 里的检索结果文件用小块。 */
export function chunkCharsFor(sourcePath) {
    const kind = /(^|[\\/])\.office[\\/]search[\\/]/.test(String(sourcePath ?? '')) ? 'search' : 'doc';
    return KB_CHUNK_CHARS_BY_KIND[kind] ?? KB_CHUNK_CHARS;
}

/** Markdown 标题行：切块的骨架。 */
const HEADING = /^(#{1,6})[ \t]+(\S.*?)[ \t]*$/;
/** 硬切时的可断点（往左找最近一个）。 */
const BREAK_AFTER = /[\n。！？；：.!?;:、，,）)】」』]$/;

/**
 * 把一段正文切成块（纯函数，确定性）。
 *
 * 两条规则：
 *   - **结构**：按 Markdown 标题分层得到「标题路径」（`A / B / C`），每个段落是
 *     一个原子单位；
 *   - **聚合**：同一标题路径下的段落依次累积到 `chunkChars` 为止；单段本身就超限
 *     时按可断点（换行 / 句末标点）硬切，最多回退 1/6 个块。
 *
 * 返回的 `span` 是**源正文的字符区间**，且满足 `text === source.slice(start, end)`
 * —— 这条等式就是「span 可回溯」的全部含义，测试里逐块断言。
 */
export function splitDocument(text, { chunkChars = KB_CHUNK_CHARS } = {}) {
    const source = typeof text === 'string' ? text : '';
    const limit = Math.max(200, boundedInt(chunkChars, KB_CHUNK_CHARS, 200, 8000));
    /** 每行连同它在源正文里的区间（行尾的 \r 不进正文）。 */
    const lines = [];
    {
        let offset = 0;
        for (const raw of source.split('\n')) {
            const clean = raw.endsWith('\r') ? raw.slice(0, -1) : raw;
            lines.push({ text: clean, start: offset, end: offset + clean.length });
            offset += raw.length + 1;
        }
    }

    const pieces = [];
    let pending = null;   // { outline: string, ranges: [{start,end}], chars }
    const flush = () => {
        if (pending !== null && pending.ranges.length > 0) pieces.push(pending);
        pending = null;
    };
    const open = (outline) => {
        pending = { outline, ranges: [], chars: 0 };
    };
    const pushRange = (range) => {
        if (pending === null) open('');
        pending.ranges.push(range);
        pending.chars += range.end - range.start;
    };

    /** 单个段落自己超限：按可断点硬切，切出来的每一段各成一块。 */
    const pushLong = (range, outline) => {
        let cursor = range.start;
        while (range.end - cursor > limit) {
            const hard = cursor + limit;
            const floor = Math.max(cursor + 60, hard - Math.floor(limit / 6));
            let cut = hard;
            for (let index = hard; index > floor; index -= 1) {
                if (BREAK_AFTER.test(source.slice(index - 1, index))) {
                    cut = index;
                    break;
                }
            }
            flush();
            open(outline);
            pushRange({ start: cursor, end: cut });
            flush();
            cursor = cut;
        }
        if (range.end > cursor) {
            if (pending === null) open(outline);
            pushRange({ start: cursor, end: range.end });
        }
    };

    let stack = [];
    let paragraph = null;   // 连续非空行合成一个段落
    /** 标题路径：跳级标题（`###` 出现在 `#` 之前）不该拼出空层级（` /  / 标题`）。 */
    const outlineOf = () => stack.filter((item) => typeof item === 'string' && item !== '').join(' / ');
    const closeParagraph = () => {
        if (paragraph === null) return;
        const range = { start: paragraph.start, end: paragraph.end };
        paragraph = null;
        const outline = outlineOf();
        if (pending !== null && pending.outline !== outline) flush();
        if (pending === null) open(outline);
        // 段落跨过额度就先收口（单段自己超限走硬切）。收口之后必须**按当前标题路径
        // 重新开一块** —— 让 `pushRange` 去兜底开块的话，它不知道标题路径，新块会变成
        // 「没有标题路径」的孤儿（同一个标题下超长的正文被切碎时尤其明显）。
        if (pending.chars > 0 && pending.chars + (range.end - range.start) > limit) {
            flush();
            open(outline);
        }
        if (range.end - range.start > limit) pushLong(range, outline);
        else pushRange(range);
    };

    for (const line of lines) {
        const heading = HEADING.exec(line.text);
        if (heading !== null) {
            closeParagraph();
            flush();
            const depth = heading[1].length;
            stack = stack.slice(0, depth - 1);
            stack[depth - 1] = heading[2].trim();
            continue;
        }
        if (line.text.trim() === '') {
            closeParagraph();
            continue;
        }
        if (paragraph === null) paragraph = { start: line.start, end: line.end };
        else paragraph.end = line.end;
    }
    closeParagraph();
    flush();

    // 收敛成对外形态：trim 掉两端空白，span 随之收窄，保证 slice 等式成立。
    const out = [];
    for (const piece of pieces) {
        const start = piece.ranges[0].start;
        const end = piece.ranges[piece.ranges.length - 1].end;
        let from = start;
        let to = end;
        while (from < to && /\s/.test(source[from])) from += 1;
        while (to > from && /\s/.test(source[to - 1])) to -= 1;
        if (to <= from) continue;
        out.push({
            n: out.length + 1,
            text: source.slice(from, to),
            span: { start: from, end: to },
            outline: piece.outline,
        });
    }
    return out;
}

/**
 * 块头：标题路径 + 来源 + 日期（Contextual Retrieval 的零依赖版本）。
 * 渲染给模型时前置在块正文之前，**不**混进 `text`（span 要与源正文逐字相等）。
 */
export function chunkHeaderOf({ outline = '', source = '', at = '' } = {}) {
    const parts = [];
    if (asText(outline) !== '') parts.push(asText(outline));
    if (asText(source) !== '') parts.push(`来源 ${asText(source)}`);
    const date = asText(at).slice(0, 10);
    if (date !== '') parts.push(date);
    return parts.join(' · ');
}

// ── 词法检索的零件（纯函数） ────────────────────────────────────────────────

/** 检索串的分词边界：与记忆侧 `queryTerms` 同一套分隔符。 */
const KB_TERM_SPLIT = /[\s,，、;；|/]+/;
/** CJK 与非 CJK 的分段（中英混排要分段处理）。 */
const KB_SEGMENT = /[\u3400-\u9fff]+|[^\u3400-\u9fff]+/g;
const KB_CJK_ONLY = /^[\u3400-\u9fff]+$/;

/**
 * 切词（**保留重复**，BM25 要词频）。
 *
 * 纯 CJK 段展开成字符二元组（长度 1 的段落没有二元组，原样保留），其余段整词保留。
 * 二元组而不是单字或词：单字太松（任意两个常见字就命中），词需要词典（本插件零依赖）；
 * 二元组是 Elasticsearch 的 CJK 默认做法，也是本机实测里唯一能把「词序颠倒」与
 * 「合称 vs 全称」救回来的粒度（第二十四轮 16 条中文查询：整串 hit@3 8/16 → 16/16）。
 */
export function kbTokens(text) {
    const tokens = [];
    for (const raw of String(text ?? '').toLowerCase().split(KB_TERM_SPLIT)) {
        if (raw === '') continue;
        for (const segment of raw.match(KB_SEGMENT) ?? []) {
            if (segment === '') continue;
            if (!KB_CJK_ONLY.test(segment)) {
                tokens.push(segment);
                continue;
            }
            const chars = [...segment];
            if (chars.length < 2) {
                tokens.push(segment);
                continue;
            }
            for (let i = 0; i + 1 < chars.length; i += 1) tokens.push(chars[i] + chars[i + 1]);
        }
    }
    return tokens;
}

/** 查询侧的词表（去重）：同一批词，BM25 只按词表迭代一次。 */
export function kbTerms(text) {
    return [...new Set(kbTokens(text))];
}

/**
 * 一个字段上的 BM25 原始分。
 *
 * 多路融合**必须**先各自算原始分再归一化：两路的量纲不同（正文长、块头短），
 * 直接相加等于让短的那一路凭词频饱和把长的那一路盖掉。
 */
function bm25Scores(texts, terms, { k1 = KB_BM25.k1, b = KB_BM25.b } = {}) {
    const wanted = new Set(terms);
    const tfs = texts.map((text) => {
        const map = new Map();
        for (const token of kbTokens(text)) map.set(token, (map.get(token) ?? 0) + 1);
        return map;
    });
    const lengths = tfs.map((map) => {
        let total = 0;
        for (const value of map.values()) total += value;
        return total;
    });
    let sum = 0;
    for (const length of lengths) sum += length;
    const avg = sum / Math.max(1, lengths.length);
    const df = new Map();
    for (const map of tfs) {
        for (const token of map.keys()) {
            if (wanted.has(token)) df.set(token, (df.get(token) ?? 0) + 1);
        }
    }
    const total_ = texts.length;
    return tfs.map((tf, index) => {
        let score = 0;
        for (const token of wanted) {
            const frequency = tf.get(token) ?? 0;
            if (frequency === 0) continue;
            const docs = df.get(token) ?? 0;
            const idf = Math.log(1 + (total_ - docs + 0.5) / (docs + 0.5));
            score += (idf * (frequency * (k1 + 1)))
                / (frequency + k1 * (1 - b + b * (lengths[index] / avg)));
        }
        return score;
    });
}

/** 每路归一化到 0..1（除以本路最高分；全 0 时原样返回）。 */
function normalizeScores(scores) {
    let max = 0;
    for (const score of scores) {
        if (score > max) max = score;
    }
    if (max <= 0) return scores.map(() => 0);
    return scores.map((score) => score / max);
}

/**
 * 在候选块上排序（纯函数、确定性）。
 *
 * 输入 `chunks` 每项至少要有 `{ id, text, head }`（`head` 是块头）。返回按融合分降序的
 * `[{ chunk, score, scoreText, scoreHead }]`，**只保留至少一路有非零原始分的块** ——
 * 「一句都不沾」就是弃答，不许把词表零重叠的块端出来（LongMemEval 的 ABS 口径）。
 * 同分按块 id 定序：块 id 是内容寻址的，所以同一批语料任何时候排出来都一样。
 */
export function kbRankChunks(chunks, query, { k1 = KB_BM25.k1, b = KB_BM25.b, weights = KB_ROUTE_WEIGHTS } = {}) {
    const list = Array.isArray(chunks) ? chunks : [];
    const terms = kbTerms(query);
    if (terms.length === 0 || list.length === 0) return [];
    const headWeight = weights?.head ?? KB_ROUTE_WEIGHTS.head;
    const textWeight = weights?.text ?? KB_ROUTE_WEIGHTS.text;
    const textRaw = bm25Scores(list.map((chunk) => chunk.text ?? ''), terms, { k1, b });
    const headRaw = bm25Scores(list.map((chunk) => chunk.head ?? ''), terms, { k1, b });
    const textNorm = normalizeScores(textRaw);
    const headNorm = normalizeScores(headRaw);
    return list
        .map((chunk, index) => ({
            chunk,
            score: textWeight * textNorm[index] + headWeight * headNorm[index],
            scoreText: textNorm[index],
            scoreHead: headNorm[index],
            rawText: textRaw[index],
            rawHead: headRaw[index],
        }))
        .filter((row) => row.rawText > 0 || row.rawHead > 0)
        .sort((left, right) => (right.score - left.score)
            || String(left.chunk.id ?? '').localeCompare(String(right.chunk.id ?? '')));
}

/**
 * 多路分的两处分位写法（回执里用，便于核对「这条为什么被排上来」）。
 * 只写到三位小数：这是给人看的依据，不是能拿来做二次计算的量。
 */
export function scoreText(value) {
    return Number.isFinite(value) ? value.toFixed(3) : '0.000';
}

/** manifest 一行的规范化（坏行直接丢掉：manifest 是追加写的，半截行不该让整份读不出来）。 */
export function normalizeManifestRow(raw) {
    if (raw === null || typeof raw !== 'object') return null;
    const id = asText(raw.id);
    const hash = asText(raw.hash);
    const path = asText(raw.path);
    if (id === '' || hash === '') return null;
    return {
        id,
        path,
        hash,
        title: asText(raw.title) || path,
        bytes: finiteNumber(raw.bytes),
        chars: finiteNumber(raw.chars),
        chunks: boundedInt(raw.chunks, 0, 0, 100_000),
        tier: kbTierOf(raw.tier),
        at: asText(raw.at),
    };
}

/** 块文件的规范化。
 *
 * 块文件里**只放与「这一块本身」有关的东西**：id、块号、正文、span、标题路径、自校验哈希。
 * 来源路径 / 入库时刻 / 来源档这些是**文档级**的，一律从 manifest 行现取 —— 两份内容相同
 * 的文档共用同一批块文件，把文档级元数据写进块文件就会互相覆盖（独立验证 F3 实测：
 * 读 a.md 的块头写着「来源 b.md」）。
 */
export function normalizeKbChunk(raw) {
    if (raw === null || typeof raw !== 'object') return null;
    const id = asText(raw.id);
    const text = typeof raw.text === 'string' ? raw.text : '';
    if (id === '' || text === '') return null;
    const span = raw.span !== null && typeof raw.span === 'object' ? raw.span : {};
    return {
        id,
        n: boundedInt(raw.n, 0, 0, 1_000_000),
        text,
        span: { start: finiteNumber(span.start), end: finiteNumber(span.end) },
        outline: asText(raw.outline),
        // 自校验：正文被动过 / 被截断时能查出来（读侧按「缺块」处理，导入侧直接拒收）。
        hash: asText(raw.hash) || kbHash(text),
    };
}

/** 一块块文件是不是自洽（正文与自校验哈希一致）。 */
export function chunkSelfConsistent(chunk) {
    return chunk !== null && asText(chunk.hash) === kbHash(chunk.text);
}

/** 计数与体积（status 与 index.json 用它；只吃 manifest 行，不读正文）。 */
export function kbCountsOf(rows) {
    const tiers = {};
    for (const tier of KB_TIERS) tiers[tier] = 0;
    let chunks = 0;
    let bytes = 0;
    for (const row of rows) {
        chunks += row.chunks;
        bytes += row.bytes;
        tiers[row.tier] = (tiers[row.tier] ?? 0) + 1;
    }
    return { docs: rows.length, chunks, bytes, tiers };
}

/**
 * 全部活着的 kb id（文档 id + 块 id）——**不读正文**。
 *
 * 块 id 是内容寻址的（`kb:<docHash>:<n>`），所以只要知道文档的哈希与块数就能
 * 把全部块 id 算出来：悬空引用这条不变量要的是事实，不该为了它把每个块文件读一遍。
 */
export function kbLiveIdsOf(rows) {
    const ids = new Set();
    for (const row of rows) {
        ids.add(row.id);
        for (let n = 1; n <= row.chunks; n += 1) ids.add(chunkIdOf(row.hash, n));
    }
    return ids;
}

/** 块在源文档里的区间写法（渲染用：`字符 120–480`）。 */
export function spanText(span) {
    const start = boundedInt(span?.start, 0, 0, Number.MAX_SAFE_INTEGER);
    const end = boundedInt(span?.end, 0, 0, Number.MAX_SAFE_INTEGER);
    return `字符 ${start}–${end}`;
}
