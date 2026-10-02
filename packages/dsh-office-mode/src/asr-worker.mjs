/**
 * SenseVoice 推理子进程（office.av 的一个内部入口，脚本里看不到它）。
 *
 * 为什么单开一个进程：SenseVoice 的解码是**同步**的原生调用，一段十分钟的
 * 音频在 2 线程下要几秒到几十秒，跑在宿主进程里会把整个会话卡住。这里沿用
 * 上游语音输入那条链的同一种形态（托管子进程），但换成一个一次性进程：
 * 父进程把「要转写的规范 WAV + 模型路径 + 分块参数」用 argv 传进来，本进程
 * 把结果 JSON 写到指定文件后退出。
 *
 * 三条纪律：
 *   - **不打印正文**：stdout / stderr 只用来留日志，结果只走结果文件。
 *   - **失败也要写结果文件**：父进程据此报出可执行的错，而不是「进程挂了」。
 *   - **VAD 切句**：整段直接喂给模型在长音频上会退化，所以按块跑 Silero VAD，
 *     逐句解码并给出绝对时间戳 —— 逐字稿要能被引用回听点。
 *   - **块之间有重叠、切点会回退**（第三十一轮）：块不再是完全切开的两段，见
 *     av.js 的 planChunks / resolveChunkBoundary —— 窗口比自有区间多出
 *     overlapSeconds，被窗口末尾切断的句子交给下一块从话头重新解，保证一句话
 *     不会在切点处被切成两半、也不会重复出稿。
 *
 * 跑法（父进程自动完成，手跑用于排查）：
 *   node src/asr-worker.mjs '{"wav":"…","outPath":"…","runtimePath":"…","model":"…","tokens":"…","vad":"…"}'
 */
import { createRequire } from 'node:module';
import { writeFileSync } from 'node:fs';
import { wordsOf } from './asr-words.js';
import { joinTranscript, planChunks, resolveChunkBoundary } from './av.js';

/** 把结果（含失败）写到父进程指定的文件；写完就退出。 */
function finish(outPath, payload) {
    writeFileSync(outPath, `${JSON.stringify(payload, null, 1)}\n`, 'utf8');
    process.exit(payload.ok === true ? 0 : 1);
}

let config;
try {
    config = JSON.parse(process.argv[2] ?? '{}');
} catch (error) {
    // 连参数都读不出来时没有结果文件可写，只能打日志。
    process.stderr.write(`asr-worker: 参数不是 JSON：${error.message}\n`);
    process.exit(2);
}

/** 把一段样本解码成一条结果（同步原生调用）。 */
function makeDecoder(recognizer, language) {
    const native = {
        featConfig: { sampleRate: 16_000, featureDim: 80 },
        modelConfig: {
            senseVoice: {
                model: config.model,
                language,
                useInverseTextNormalization: 1,
            },
            tokens: config.tokens,
            numThreads: config.threads,
            provider: 'cpu',
            debug: 0,
        },
    };
    recognizer.setConfig(native);
    return (samples) => {
        const stream = recognizer.createStream();
        stream.acceptWaveform({ sampleRate: 16_000, samples });
        recognizer.decode(stream);
        return recognizer.getResult(stream);
    };
}

try {
    const require = createRequire(import.meta.url);
    const sherpa = require(config.runtimePath);
    const wave = sherpa.readWave(config.wav);
    if (wave?.sampleRate !== 16_000) {
        finish(config.outPath, { ok: false, error: `WAV 采样率是 ${wave?.sampleRate}，必须是 16000（office.av 会先规范化，走到这里说明文件被改过）` });
    }
    const samples = wave.samples;
    const language = typeof config.language === 'string' && config.language !== '' ? config.language : 'auto';
    const chunkSeconds = Math.max(1, Number(config.chunkSeconds) || 120);
    const overlapSeconds = Math.max(0, Math.min(Number(config.overlapSeconds) || 0, chunkSeconds / 2));
    // 单句上限夹到不超过单块秒数：一句话必须能被某个窗口完整装下，边界回退才有
    // 地方可退。不夹的话（例如单块 10 秒、单句上限 30 秒），一句连续的 30 秒话
    // 会横跨整块窗口，回退点落在块首 —— 循环原地打转。
    const maxSegmentSeconds = Math.max(1, Math.min(Number(config.maxSegmentSeconds) || 30, chunkSeconds));
    const recognizer = new sherpa.OfflineRecognizer({
        featConfig: { sampleRate: 16_000, featureDim: 80 },
        modelConfig: {
            senseVoice: { model: config.model, language, useInverseTextNormalization: 1 },
            tokens: config.tokens,
            numThreads: config.threads,
            provider: 'cpu',
            debug: 0,
        },
    });
    const detector = new sherpa.Vad({
        sileroVad: {
            model: config.vad,
            threshold: config.vadThreshold,
            minSilenceDuration: config.minSilenceSeconds,
            minSpeechDuration: config.minSpeechSeconds,
            maxSpeechDuration: maxSegmentSeconds,
            windowSize: 512,
        },
        sampleRate: 16_000,
        numThreads: config.threads,
        provider: 'cpu',
        debug: 0,
    }, Math.ceil(maxSegmentSeconds + config.minSilenceSeconds + 1));
    const decode = makeDecoder(recognizer, language);

    const samplesPerChunk = Math.round(chunkSeconds * 16_000);
    const overlapSamples = Math.round(overlapSeconds * 16_000);
    const segments = [];
    const detected = new Set();
    let inferenceSeconds = 0;
    let speechSamples = 0;
    let chunkCount = 0;
    let rolls = 0;

    const totalSamples = samples.length;
    const totalSeconds = totalSamples / 16_000;
    const toSeconds = (sampleCount) => Math.round((sampleCount / 16_000) * 1000) / 1000;

    /**
     * 解一段样本，返回一条结果（没识别出文字就返回 null）。
     *
     * 只解码不登记：登记（出稿 / 计数）由调用方决定 —— 被边界回退交给下一块的
     * 句子同样要解码，但不能进逐字稿，否则就重复了。
     */
    const decodeSlice = (slice, startSample, endSample) => {
        const started = performance.now();
        const result = decode(slice);
        inferenceSeconds += (performance.now() - started) / 1000;
        const text = String(result?.text ?? '').trim();
        if (text === '') return null;
        const lang = String(result?.lang ?? '');
        if (lang !== '') detected.add(lang.replace(/[<>|]/g, ''));
        // 词级时间（`words`）要自己算：sherpa-onnx 的 SenseVoice 通道**不返回**
        // words（1.13.8 实测恒为 `[]`，绑定里也没有开词级时间的参数），
        // 真正可用的是逐 token 的起始秒（tokens[] + timestamps[]）。算法见 asr-words.js。
        // 第三个参数是**基址**：timestamps 相对这一段的开头，不加基址就对不回听点。
        // 算不出来时**不写这个字段**，而不是写一个空数组假装有。
        const words = config.words === true
            ? wordsOf(result, (endSample - startSample) / 16_000, toSeconds(startSample))
            : undefined;
        return {
            start: toSeconds(startSample),
            end: toSeconds(endSample),
            length: endSample - startSample,
            text,
            lang: lang || null,
            emotion: result?.emotion === undefined ? null : String(result.emotion),
            event: result?.event === undefined ? null : String(result.event),
            ...(words === undefined ? {} : { words }),
        };
    };

    /** 把一条解码结果收进逐字稿。 */
    const emit = (item) => {
        speechSamples += item.length;
        segments.push({
            index: segments.length,
            start: item.start,
            end: item.end,
            text: item.text,
            lang: item.lang,
            emotion: item.emotion,
            event: item.event,
            ...(item.words === undefined ? {} : { words: item.words }),
        });
    };

    // 逐块跑，但块起点不再固定按 chunkSeconds 前进：窗口末尾切在句子中间时，
    // 下一块的起点会回退到那句话的话头（见 resolveChunkBoundary）。
    // `emittedUntil` 是**已出稿到的时刻**：重叠让相邻两块都看得见同一段音频，
    // 没有这条水位线，上一块刚解出来的完整句子会在下一块里再报一次。
    let baseSample = 0;
    let emittedUntil = 0;
    while (baseSample < totalSamples) {
        const base = baseSample / 16_000;
        const ownedEndSample = Math.min(totalSamples, baseSample + samplesPerChunk);
        const winEndSample = Math.min(totalSamples, ownedEndSample + overlapSamples);
        const ownedEnd = ownedEndSample / 16_000;
        const winEnd = winEndSample / 16_000;
        chunkCount += 1;

        detector.reset();
        const raw = [];
        // 抽干 VAD 里已完成的句子：`fed` 是当前已喂进去的样本数（用于算绝对时间）。
        const drain = (fed) => {
            while (!detector.isEmpty()) {
                const segment = detector.front(false);
                const length = segment?.samples?.length ?? 0;
                const endSample = baseSample + fed;
                const startSample = Math.max(baseSample, endSample - length);
                if (length > 0) {
                    const item = decodeSlice(segment.samples, startSample, endSample);
                    if (item !== null) raw.push(item);
                }
                detector.pop();
            }
        };

        const window = samples.subarray(baseSample, winEndSample);
        let fed = 0;
        for (let offset = 0; offset < window.length; offset += 512) {
            const part = window.subarray(offset, Math.min(window.length, offset + 512));
            detector.acceptWaveform(part);
            fed = offset + part.length;
            drain(fed);
        }
        detector.flush();
        fed = window.length;
        drain(fed);

        const decision = resolveChunkBoundary({ base, ownedEnd, winEnd, total: totalSeconds, raw, emittedUntil });
        for (const index of decision.keep) {
            const item = raw[index];
            emit(item);
            emittedUntil = Math.max(emittedUntil, item.end);
        }
        if (decision.rolledBack) rolls += 1;

        // VAD 判断「整块都不是语音」（纯音乐、纯静音、或信号太弱）时一句话都不会出，
        // 这时退回整段解码 —— 否则一段人是能听清的录音会静默变成空稿。
        // **只解自有区间里还没出过稿的那一段**：重叠区属于下一块，已出稿的前半截
        // 也不能再解一遍，否则同一段文字会出现两次。
        const fallbackFrom = Math.max(baseSample, Math.round(emittedUntil * 16_000));
        if (raw.length === 0 && fallbackFrom < ownedEndSample) {
            const item = decodeSlice(samples.subarray(fallbackFrom, ownedEndSample), fallbackFrom, ownedEndSample);
            if (item !== null) {
                emit(item);
                emittedUntil = Math.max(emittedUntil, item.end);
            }
        }

        // 回退点换算回样本；resolveChunkBoundary 已保证前进，这里再钉一次 ——
        // 浮点四舍五入到同一采样点时最坏的情况是原地打转，那是死循环，不能留。
        const nextSample = Math.round(decision.nextBase * 16_000);
        baseSample = nextSample > baseSample ? nextSample : baseSample + 1;
    }

    // 时间戳补齐成 `mm:ss-mm:ss`，逐字稿与报告都能直接用。
    for (const segment of segments) {
        segment.clock = `${clock(segment.start)}-${clock(segment.end)}`;
    }

    finish(config.outPath, {
        ok: true,
        model: `SenseVoiceSmall (${String(config.model).toLowerCase().includes('int8') ? 'INT8' : 'FP32'})`,
        language,
        detectedLanguages: [...detected],
        audioSeconds: Math.round((samples.length / 16_000) * 1000) / 1000,
        speechSeconds: Math.round((speechSamples / 16_000) * 1000) / 1000,
        inferenceSeconds: Math.round(inferenceSeconds * 1000) / 1000,
        chunkSeconds,
        overlapSeconds,
        chunks: chunkCount,
        rolls,
        segments,
        text: joinTranscript(segments.map((item) => item.text)),
    });
} catch (error) {
    finish(config.outPath, { ok: false, error: `${error?.name ?? 'Error'}: ${error?.message ?? String(error)}` });
}

/** 秒 → `mm:ss` / `h:mm:ss`（与 av.js 的 formatClock 同一口径）。 */
function clock(seconds) {
    const total = Math.max(0, Math.round(Number(seconds) || 0));
    const hours = Math.floor(total / 3600);
    const minutes = Math.floor((total % 3600) / 60);
    const secs = total % 60;
    const pad = (value) => String(value).padStart(2, '0');
    return hours > 0 ? `${hours}:${pad(minutes)}:${pad(secs)}` : `${pad(minutes)}:${pad(secs)}`;
}
