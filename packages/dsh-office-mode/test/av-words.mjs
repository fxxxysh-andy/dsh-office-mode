/**
 * 词级时间（office.av.transcribe 的 `words`）的回归与**量化**测试（历史遗留 `22-4`）。
 *
 * 这条账原来的说法是「`words` 实现了但没量化」。实测（sherpa-onnx-node 1.13.8 +
 * SenseVoice）`words` 是**没实现**：模型侧不返回词级时间，绑定里也没有开关，
 * `result.words` 恒为 `[]` —— 老代码只是往每句塞了一个空数组，体积 +16 B/句、
 * 信息为零。本轮的修法是拿逐 token 的起始秒（tokens + timestamps）现算，并把
 * 代价与收益量出来（这个文件就是那把尺子）。
 *
 * 三段：
 *   1. 纯函数 `wordsOf`：中日韩按字、西文按词、标点并入前词、缺 tokens 时返回
 *      `undefined`（不写空数组假装有）。
 *   2. 转写稿渲染：词级时间那一节与 600 词封顶。
 *   3. 真跑（缺 ffmpeg / 模型 / 运行时时整段 SKIP 并如实说明）：用 Windows SAPI
 *      合成一段中文语音 → 规范 WAV → 同一个 worker 跑 words on/off，
 *      量 **体积 / 推理耗时 / 时间定位精度**，并断言 words 的覆盖与单调性。
 *
 * 跑法：node test/av-words.mjs
 */
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { normalizeWords, wordsOf } from '../src/asr-words.js';
import { AV_WORKER_PATH, canonicalizeWav, inspectWav, probeAv, renderTranscript } from '../src/av.js';
import { resolveConfig } from '../src/config.js';
import { runPdfProcess } from '../src/pdf.js';
import { executeRun } from '../src/run.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        if (error?.skip === true) results.push({ name, ok: true, note: `SKIP ${error.message}` });
        else results.push({ name, ok: false, error });
    }
}
function skip(message) {
    const error = new Error(message);
    error.skip = true;
    throw error;
}

const root = mkdtempSync(join(tmpdir(), 'office-av-words-'));
const measured = {};

// ── 1. 纯函数：wordsOf 的尺度与边界 ─────────────────────────────────────────

await check('wordsOf：中日韩按字，标点并进前一个字', async () => {
    const result = {
        tokens: ['今', '天', '下', '午', '3', '点', '开', '会', '，', '好', '。'],
        timestamps: [0, 0.2, 0.4, 0.6, 0.8, 1.0, 1.2, 1.4, 1.6, 1.8, 2.0],
    };
    const words = wordsOf(result, 2.2);
    assert.equal(words.length, 9, '11 个 token 里两个标点并入前字，应为 9 条');
    assert.equal(words.map((item) => item.text).join(''), '今天下午3点开会，好。');
    assert.equal(words[0].text, '今');
    assert.equal(words[0].start, 0);
    assert.equal(words[0].end, 0.2, '词的终点是下一个 token 的起点');
    assert.equal(words[7].text, '会，', '标点并入前一个字');
    assert.equal(words[8].text, '好。');
    assert.equal(words[8].end, 2.2, '最后一个词收在本句终点');
    return `${words.length} 条`;
});

await check('wordsOf：西文按词（▁ 是边界，子词并回前一个词）', async () => {
    const result = {
        tokens: ['▁Quar', 'ter', 'ly', '▁re', 'view', '▁meeting', '.'],
        timestamps: [0, 0.3, 0.5, 0.7, 0.9, 1.2, 1.6],
    };
    const words = wordsOf(result, 1.7);
    assert.deepEqual(words.map((item) => item.text), ['Quarterly', 'review', 'meeting.']);
    assert.equal(words[0].start, 0);
    assert.equal(words[2].end, 1.7);
    return words.map((item) => item.text).join(' ');
});

await check('wordsOf：拿不到 tokens / timestamps 时返回 undefined（不写空数组）', async () => {
    assert.equal(wordsOf({ words: [] }, 1), undefined, '上游给空 words 且没有 tokens → undefined');
    assert.equal(wordsOf({ tokens: ['a'], timestamps: [] }, 1), undefined, '长度对不上 → undefined');
    assert.equal(wordsOf({}, 1), undefined);
    assert.equal(wordsOf({ tokens: ['▁', ' '], timestamps: [0, 0.1] }, 1), undefined, '全是空白 token → undefined');
    // 上游哪天真的给了 words，就照上游的用（形状归一化）。
    const upstream = wordsOf({ words: [{ text: '你好', start: 0.5, end: 1.1 }] }, 2);
    assert.deepEqual(upstream, [{ text: '你好', start: 0.5, end: 1.1 }]);
    assert.deepEqual(normalizeWords('[]'), [], 'JSON 字符串也认');
    assert.equal(normalizeWords('不是 JSON'), undefined);
    return '全部返回 undefined / 归一化';
});

await check('wordsOf：相对时间要加基址、越界值夹在本句区间内', async () => {
    // timestamps 是相对本段开头的秒数：基址 0.66 的句子里，第一个 token 的 0.06
    // 必须落成 0.72（绝对时间），而不是 0.06 —— 少了这一步，词的时间会全挤在句首。
    const offsets = wordsOf({ tokens: ['甲', '乙'], timestamps: [0, 0.5] }, 1, 0.66);
    assert.deepEqual(offsets, [{ text: '甲', start: 0.66, end: 1.16 }, { text: '乙', start: 1.16, end: 1.66 }]);
    const clamped = wordsOf({ tokens: ['甲', '乙'], timestamps: [5, 5.5] }, 1, 0.5);
    assert.deepEqual(clamped, [{ text: '甲', start: 1.5, end: 1.5 }, { text: '乙', start: 1.5, end: 1.5 }], '越界的时间戳被夹到本句时长');
    return '基址 + clamp 生效';
});

// ── 2. 转写稿渲染 ────────────────────────────────────────────────────────────

await check('renderTranscript：有词级时间时多一节，且按 TRANSCRIPT_WORD_LIMIT 封顶', async () => {
    const segment = (index) => ({
        index,
        start: index * 2,
        end: index * 2 + 2,
        clock: `00:0${index}-00:0${index + 1}`,
        text: '你好',
        words: [{ text: '你', start: index * 2, end: index * 2 + 1 }, { text: '好', start: index * 2 + 1, end: index * 2 + 2 }],
    });
    const withWords = renderTranscript({ path: 'a.wav', segments: [segment(0)], segmentCount: 1, audioSeconds: 2 });
    assert.match(withWords, /## 词级时间/);
    assert.match(withWords, /0\.00 你 \/ 1\.00 好/);
    const many = renderTranscript({ path: 'a.wav', segments: Array.from({ length: 400 }, (_, i) => segment(i)), segmentCount: 400 });
    assert.match(many, /被省略：单份转写稿最多列 600 个词/);
    // 硬上限：正文里印出来的词不超过 600 个（跨线的那一句只印装得下的那几个）
    const printedWords = (many.match(/[0-9]+\.[0-9]{2} /g) ?? []).length;
    assert.ok(printedWords <= 600, `词级时间印了 ${printedWords} 个词，超过 600 的上限`);
    assert.ok(printedWords >= 598, `印得太少也不对：${printedWords}`);
    const plain = renderTranscript({ path: 'a.wav', segments: [{ index: 0, start: 0, end: 1, clock: '00:00-00:01', text: 'hi' }], segmentCount: 1 });
    assert.ok(!plain.includes('## 词级时间'), '没开 words 时不该出现这一节');
    return '一节 + 封顶';
});

// ── 3. 真跑：合成语音 → worker words on/off → 量代价与精度 ────────────────────

const probe = await probeAv(resolveConfig({}));
const ready = probe.available === true;

if (!ready) {
    results.push({ name: '真跑量化（词级时间 vs 句级时间戳）', ok: true, note: `SKIP ${probe.hint}` });
} else {
    /** 用 Windows SAPI 合成一段语音（开发期夹具，不进仓库）。 */
    async function synthesize(text, outPath) {
        if (process.platform !== 'win32') throw Object.assign(new Error('只有 Windows 有 SAPI，本机不合成语音夹具'), { skip: true });
        const script = [
            'Add-Type -AssemblyName System.Speech',
            '$sp = New-Object System.Speech.Synthesis.SpeechSynthesizer',
            '$zh = $sp.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Culture.Name -like "zh*" } | Select-Object -First 1',
            'if ($zh) { $sp.SelectVoice($zh.VoiceInfo.Name) }',
            '$fmt = New-Object System.Speech.AudioFormat.SpeechAudioFormatInfo(16000, [System.Speech.AudioFormat.AudioBitsPerSample]::Sixteen, [System.Speech.AudioFormat.AudioChannel]::Mono)',
            `$sp.SetOutputToWaveFile('${outPath.replace(/'/g, "''")}', $fmt)`,
            `$sp.Speak('${text.replace(/'/g, "''")}')`,
            '$sp.Dispose()',
        ].join('\n');
        const scriptPath = join(root, 'make-speech.ps1');
        writeFileSync(scriptPath, `\ufeff${script}`, 'utf8');
        const run = await runPdfProcess('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath], {
            outPath: join(root, 'ps.out.txt'),
            errPath: join(root, 'ps.err.txt'),
            timeoutMs: 60_000,
        });
        if (run.code !== 0 || !existsSync(outPath)) {
            throw Object.assign(new Error(`SAPI 合成失败：${(run.err ?? '').split('\n').slice(-3).join(' ').slice(0, 200)}`), { skip: true });
        }
        return outPath;
    }

    const spoke = join(root, 'speech.raw.wav');
    const wavPath = join(root, 'speech.wav');
    const speechText = '今天下午三点开季度会议，请提前准备经营报表。';

    await check('造夹具：SAPI 合成中文语音并转成规范 WAV', async () => {
        await synthesize(speechText, spoke);
        const canonical = canonicalizeWav(readFileSync(spoke));
        writeFileSync(wavPath, canonical);
        const info = inspectWav(readFileSync(wavPath), readFileSync(wavPath).length);
        assert.equal(info.ok, true, info.reason);
        assert.ok(info.seconds > 2 && info.seconds < 30, `时长异常：${info.seconds}`);
        measured.audioSeconds = info.seconds;
        return `${info.seconds.toFixed(2)} 秒`;
    });

    /** 直接驱动 worker（与 av.js 同一条 argv 契约），结果写进临时目录。 */
    async function runWorker(words) {
        const outPath = join(root, `result-${words ? 'words' : 'plain'}.json`);
        const config = {
            wav: wavPath,
            outPath,
            runtimePath: probe.sensevoice.runtime,
            model: probe.sensevoice.model,
            tokens: probe.sensevoice.tokens,
            vad: probe.sensevoice.vad,
            threads: 2,
            language: 'auto',
            chunkSeconds: 120,
            overlapSeconds: 0.5,
            maxSegmentSeconds: 30,
            vadThreshold: 0.5,
            minSpeechSeconds: 0.25,
            minSilenceSeconds: 0.5,
            words,
        };
        const run = await runPdfProcess(process.execPath, [AV_WORKER_PATH, JSON.stringify(config)], {
            outPath: join(root, `worker-${words}.out.txt`),
            errPath: join(root, `worker-${words}.err.txt`),
            timeoutMs: 300_000,
            cwd: root,
        });
        assert.ok(existsSync(outPath), `worker 没写结果文件（code=${run.code}）：${(run.err ?? '').slice(-300)}`);
        const payload = JSON.parse(readFileSync(outPath, 'utf8'));
        assert.equal(payload.ok, true, payload.error);
        payload.bytes = readFileSync(outPath).length;
        return payload;
    }

    let plain = null;
    let withWords = null;

    await check('真跑：words:false 时**没有** words 字段（老代码的无条件 words:[] 是回归点）', async () => {
        plain = await runWorker(false);
        assert.ok(plain.segments.length > 0, '这段语音至少该识别出一句');
        for (const segment of plain.segments) {
            assert.ok(!Object.prototype.hasOwnProperty.call(segment, 'words'), '没开 words 却写了 words 字段');
        }
        measured.plainBytes = plain.bytes;
        measured.plainSeconds = plain.inferenceSeconds;
        return `${plain.segments.length} 句 / ${plain.bytes} B / ${plain.inferenceSeconds}s`;
    });

    await check('真跑：words:true 真的给出词级时间，且覆盖整句、时间单调不越界', async () => {
        if (plain === null) throw Object.assign(new Error('上一步没跑成'), { skip: true });
        withWords = await runWorker(true);
        const withWordSegments = withWords.segments.filter((item) => Array.isArray(item.words) && item.words.length > 0);
        assert.ok(withWordSegments.length > 0, '开了 words 却没有任何词级时间 —— 就是这个 bug 要防的');
        let previousEnd = -1;
        for (const segment of withWordSegments) {
            assert.equal(segment.words.map((word) => word.text).join(''), segment.text, '词级时间的字面要拼回整句（不丢字）');
            for (const word of segment.words) {
                assert.ok(Number.isFinite(word.start) && Number.isFinite(word.end), 'start/end 要是数字');
                assert.ok(word.start <= word.end, `词的 start 不能大于 end：${JSON.stringify(word)}`);
                assert.ok(word.start >= segment.start - 0.01 && word.end <= segment.end + 0.01,
                    `词的时间越出本句区间：${JSON.stringify(word)} vs ${segment.start}-${segment.end}`);
            }
            const first = segment.words[0].start;
            assert.ok(first >= previousEnd - 0.01, '相邻句的词级时间不能回退');
            previousEnd = segment.words[segment.words.length - 1].end;
        }
        measured.wordsBytes = withWords.bytes;
        measured.wordsSeconds = withWords.inferenceSeconds;
        measured.wordSegments = withWordSegments.length;
        measured.wordCount = withWordSegments.reduce((sum, item) => sum + item.words.length, 0);
        return `${measured.wordSegments} 句 / ${measured.wordCount} 词 / ${withWords.bytes} B`;
    });

    await check('量化：词级时间定位精度 vs 只用句级时间戳（中点估计）', async () => {
        if (withWords === null) throw Object.assign(new Error('真跑没完成'), { skip: true });
        const words = withWords.segments.flatMap((segment) => (segment.words ?? []).map((word) => ({ ...word, segment })));
        assert.ok(words.length > 0, '没有词级时间无法量化');
        const segmentDurations = withWords.segments.map((item) => item.end - item.start);
        const wordDurations = words.map((word) => word.end - word.start);
        const median = (list) => {
            const sorted = [...list].sort((a, b) => a - b);
            return sorted[Math.floor(sorted.length / 2)];
        };
        // 只用句级时间戳时，一个词的时间只能取句子的中点 —— 偏差 = |词的起点 - 句中点|。
        const errors = words.map((word) => Math.abs(word.start - (word.segment.start + word.segment.end) / 2));
        measured.segmentMedianSeconds = median(segmentDurations);
        measured.wordMedianSeconds = median(wordDurations);
        measured.segmentMidpointMeanError = errors.reduce((sum, value) => sum + value, 0) / errors.length;
        measured.segmentMidpointMaxError = Math.max(...errors);
        measured.sizeGrowth = measured.wordsBytes / measured.plainBytes;
        measured.timeDelta = measured.wordsSeconds - measured.plainSeconds;
        assert.ok(measured.wordMedianSeconds < measured.segmentMedianSeconds,
            '词的粒度应当细于句子（否则词级时间没有意义）');
        assert.ok(measured.sizeGrowth > 1, '开了词级时间体积必须变大 —— 这条是代价，不是收益');
        return `句级中位 ${measured.segmentMedianSeconds.toFixed(2)}s → 词级中位 ${measured.wordMedianSeconds.toFixed(2)}s；`
            + `只用句级时间戳时词的定位偏差均值 ${measured.segmentMidpointMeanError.toFixed(2)}s（最大 ${measured.segmentMidpointMaxError.toFixed(2)}s）`;
    });

    await check('量化：体积与推理耗时的代价（打印成表，供档案引用）', async () => {
        if (measured.wordsBytes === undefined) throw Object.assign(new Error('真跑没完成'), { skip: true });
        const fmt = (bytes) => `${(bytes / 1024).toFixed(1)} KB`;
        measured.report = [
            `音频 ${measured.audioSeconds.toFixed(2)}s；句级 ${fmt(measured.plainBytes)} / ${measured.plainSeconds}s 推理`,
            `词级 ${fmt(measured.wordsBytes)} / ${measured.wordsSeconds}s 推理；体积 ×${measured.sizeGrowth.toFixed(2)}，推理 ${measured.timeDelta >= 0 ? '+' : ''}${measured.timeDelta.toFixed(3)}s`,
            `词粒度中位 ${measured.wordMedianSeconds.toFixed(3)}s，句粒度中位 ${measured.segmentMedianSeconds.toFixed(3)}s；只用句级时间戳时定位偏差均值 ${measured.segmentMidpointMeanError.toFixed(2)}s`,
        ].join('\n');
        return `×${measured.sizeGrowth.toFixed(2)} 体积`;
    });

    await check('office_run 全链路：words:true 时 wordTimings 计数在，out 的转写稿带词级一节', async () => {
        if (withWords === null) throw Object.assign(new Error('真跑没完成'), { skip: true });
        const script = `return await office.av.transcribe(${JSON.stringify(wavPath)}, { language: 'zh', words: true, out: 'words-transcript.md' });`;
        const result = await executeRun({ script }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.ok(value.wordTimings !== undefined, 'words:true 时顶层要有 wordTimings 计数');
        assert.ok(value.wordTimings.count > 0, `wordTimings.count 应为正：${JSON.stringify(value.wordTimings)}`);
        assert.match(value.wordTimings.unit, /中日韩按字/);
        const transcript = readFileSync(join(root, 'words-transcript.md'), 'utf8');
        assert.match(transcript, /## 词级时间/);
        // 不开 words 时不该有 wordTimings（同一个入口，只有开关不同）。
        const off = await executeRun({
            script: `return await office.av.transcribe(${JSON.stringify(wavPath)}, { language: 'zh' });`,
        }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
        assert.equal(off.ok, true, off.error?.message);
        assert.equal(off.returned.wordTimings, undefined);
        assert.ok(!off.returned.segments.some((item) => item.words !== undefined), '没开 words 时段里不该有 words');
        return `wordTimings ${value.wordTimings.segments} 句 / ${value.wordTimings.count} 词`;
    });
}

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
for (const line of (measured.report ?? '').split('\n').filter((item) => item !== '')) console.log(`量化  ${line}`);
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`av-words: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
