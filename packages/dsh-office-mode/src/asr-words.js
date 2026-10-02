/**
 * 词级时间：从 sherpa-onnx 的 tokens + timestamps 现算。
 *
 * 为什么单独一个模块（第二十三轮 `22-4` 的实测结论）：SenseVoice 通道**不返回**
 * 词级时间 —— sherpa-onnx-node 1.13.8 里 `result.words` 恒为空数组，绑定里也没有
 * 开关（`enable_word_timestamps` / `word_timestamps` / `enable_words` 一个都没有）。
 * 所以「词级时间戳」不能靠打开一个参数得到，只能拿逐 token 的起始秒自己算。
 * 以前 `words:true` 只是往每句里塞一个 `words: []`：体积 +16 B/句、信息为零、
 * 看起来却像「做了」—— 这条账记的正是它。
 *
 * 尺度与定义（两者都写在这里，因为下游的时间轴对齐靠它）：
 *   - 中日韩按**字**、西文按**词**（SentencePiece 的 `▁` 是词边界，子词并回前一个词），
 *     标点并进前一个词；
 *   - 词的起点 = 它第一个 token 的起点，终点 = **下一个 token 的起点**
 *     （CTC 对齐给出的定义，最后一个词收在本句终点）。
 *
 * 拿不到 tokens/timestamps（空、长度对不上）时返回 `undefined`：调用方据此
 * **不写 words 字段**，而不是写空数组假装有。
 *
 * @module dsh-office-mode/asr-words
 */

const CJK_CHAR = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uac00-\ud7a3]/;
const PUNCT_ONLY = /^[.,;:!?'"()[\]{}<>《》“”‘’、，。！？；：（）【】…—－·~]+$/;

const round3 = (value) => Math.round((Number(value) || 0) * 1000) / 1000;

/** 词级时间戳：上游可能给数组，也可能给 JSON 字符串。 */
export function normalizeWords(raw) {
    if (Array.isArray(raw)) return raw;
    if (typeof raw === 'string' && raw.trim() !== '') {
        try {
            const parsed = JSON.parse(raw);
            return Array.isArray(parsed) ? parsed : undefined;
        } catch {
            return undefined;
        }
    }
    return undefined;
}

/**
 * 从一条解码结果算出词级时间。
 *
 * `baseSeconds` 是这一段音频在整份媒体里的起点：sherpa 给的 `timestamps` 是
 * **相对这一段**的秒数，所以要加上基址才是绝对时间。少了这一步，词的时间会
 * 全部落在句子的开头（实测：某句 0.67–5.70 秒，词级时间却是 0.06–5.10 —— 看起来
 * 「有数据」，对不回听点）。同一条边界也用来夹住越界值。
 *
 * @param {object} result sherpa 的解码结果（tokens / timestamps / words）
 * @param {number} lengthSeconds 这一段的实际时长（秒），用作最后一个词的终点
 * @param {number} [baseSeconds] 这一段在整份媒体里的起点（秒）
 * @returns {Array<{text: string, start: number, end: number}>|undefined}
 */
export function wordsOf(result, lengthSeconds, baseSeconds = 0) {
    const base = Math.max(0, Number(baseSeconds) || 0);
    const upstream = normalizeWords(result?.words);
    if (Array.isArray(upstream) && upstream.length > 0) {
        const mapped = upstream
            .map((item) => ({
                text: String(item?.text ?? item?.word ?? ''),
                start: round3(base + (Number(item?.start ?? item?.startTime ?? item?.start_time) || 0)),
                end: round3(base + (Number(item?.end ?? item?.endTime ?? item?.end_time) || 0)),
            }))
            .filter((item) => item.text !== '');
        if (mapped.length > 0) return mapped;
    }
    const tokens = Array.isArray(result?.tokens) ? result.tokens : [];
    const stamps = Array.isArray(result?.timestamps) ? result.timestamps : [];
    if (tokens.length === 0 || tokens.length !== stamps.length) return undefined;
    const duration = Math.max(0, Number(lengthSeconds) || 0);
    const upper = base + duration;
    const clamp = (value) => Math.min(upper, Math.max(base, round3(base + value)));
    const entries = [];
    for (let i = 0; i < tokens.length; i += 1) {
        const raw = String(tokens[i] ?? '');
        const bare = raw.replace(/^[▁\s]+/, '');
        if (bare === '') continue;
        const start = clamp(Number(stamps[i]));
        const end = clamp(i + 1 < tokens.length ? Number(stamps[i + 1]) : duration);
        const previous = entries[entries.length - 1];
        if (PUNCT_ONLY.test(bare) && previous !== undefined) {
            previous.text += bare;
            previous.end = Math.max(previous.end, end);
            continue;
        }
        const boundary = raw.startsWith('▁') || CJK_CHAR.test(bare) || previous === undefined;
        const continuesLatin = previous !== undefined
            && !boundary
            && /[A-Za-z0-9]$/.test(previous.text)
            && /^[A-Za-z0-9]/.test(bare);
        if (continuesLatin) {
            previous.text += bare;
            previous.end = Math.max(previous.end, end);
            continue;
        }
        entries.push({ text: bare, start, end: Math.max(start, end) });
    }
    return entries.length === 0 ? undefined : entries;
}
