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
 *
 * 跑法（父进程自动完成，手跑用于排查）：
 *   node src/asr-worker.mjs '{"wav":"…","outPath":"…","runtimePath":"…","model":"…","tokens":"…","vad":"…"}'
 */
import { createRequire } from 'node:module';
import { writeFileSync } from 'node:fs';
import { joinTranscript } from './av.js';

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
            maxSpeechDuration: config.maxSegmentSeconds,
            windowSize: 512,
        },
        sampleRate: 16_000,
        numThreads: config.threads,
        provider: 'cpu',
        debug: 0,
    }, Math.ceil(config.maxSegmentSeconds + config.minSilenceSeconds + 1));
    const decode = makeDecoder(recognizer, language);

    const chunkSeconds = Math.max(1, Number(config.chunkSeconds) || 120);
    const samplesPerChunk = Math.round(chunkSeconds * 16_000);
    const segments = [];
    const detected = new Set();
    let inferenceSeconds = 0;
    let speechSamples = 0;
    let chunkCount = 0;

    /** 抽干 VAD 里已完成的句子：`fed` 是当前已喂进去的样本数（用于算绝对时间）。 */
    const drain = (fed, base) => {
        while (!detector.isEmpty()) {
            const segment = detector.front(false);
            const length = segment?.samples?.length ?? 0;
            const endSample = base + fed;
            const startSample = Math.max(base, endSample - length);
            if (length > 0) {
                const started = performance.now();
                const result = decode(segment.samples);
                inferenceSeconds += (performance.now() - started) / 1000;
                speechSamples += length;
                const text = String(result?.text ?? '').trim();
                if (text !== '') {
                    const lang = String(result?.lang ?? '');
                    if (lang !== '') detected.add(lang.replace(/[<>|]/g, ''));
                    segments.push({
                        index: segments.length,
                        start: Math.round((startSample / 16_000) * 1000) / 1000,
                        end: Math.round((endSample / 16_000) * 1000) / 1000,
                        text,
                        lang: lang || null,
                        emotion: result?.emotion === undefined ? null : String(result.emotion),
                        event: result?.event === undefined ? null : String(result.event),
                        ...(config.words === true ? { words: normalizeWords(result?.words) } : {}),
                    });
                }
            }
            detector.pop();
        }
    };

    for (let base = 0; base < samples.length; base += samplesPerChunk) {
        const slice = samples.subarray(base, Math.min(samples.length, base + samplesPerChunk));
        chunkCount += 1;
        detector.reset();
        let fed = 0;
        for (let offset = 0; offset < slice.length; offset += 512) {
            const part = slice.subarray(offset, Math.min(slice.length, offset + 512));
            detector.acceptWaveform(part);
            fed = offset + part.length;
            drain(fed, base);
        }
        detector.flush();
        fed = slice.length;
        drain(fed, base);
        // VAD 判断「整块都不是语音」（纯音乐、纯静音、或信号太弱）时一句话都不会出，
        // 这时退回整块解码 —— 否则一段人是能听清的录音会静默变成空稿。
        const produced = segments.filter((item) => item.end > base / 16_000 && item.start < (base + slice.length) / 16_000);
        if (produced.length === 0 && slice.length > 0) {
            const started = performance.now();
            const result = decode(slice);
            inferenceSeconds += (performance.now() - started) / 1000;
            speechSamples += slice.length;
            const text = String(result?.text ?? '').trim();
            if (text !== '') {
                const lang = String(result?.lang ?? '');
                if (lang !== '') detected.add(lang.replace(/[<>|]/g, ''));
                segments.push({
                    index: segments.length,
                    start: Math.round((base / 16_000) * 1000) / 1000,
                    end: Math.round((Math.min(samples.length, base + slice.length) / 16_000) * 1000) / 1000,
                    text,
                    lang: lang || null,
                    emotion: result?.emotion === undefined ? null : String(result.emotion),
                    event: result?.event === undefined ? null : String(result.event),
                    ...(config.words === true ? { words: normalizeWords(result?.words) } : {}),
                });
            }
        }
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
        chunks: chunkCount,
        segments,
        text: joinTranscript(segments.map((item) => item.text)),
    });
} catch (error) {
    finish(config.outPath, { ok: false, error: `${error?.name ?? 'Error'}: ${error?.message ?? String(error)}` });
}

/** 词级时间戳：上游可能给数组，也可能给 JSON 字符串。 */
function normalizeWords(raw) {
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

/** 秒 → `mm:ss` / `h:mm:ss`（与 av.js 的 formatClock 同一口径）。 */
function clock(seconds) {
    const total = Math.max(0, Math.round(Number(seconds) || 0));
    const hours = Math.floor(total / 3600);
    const minutes = Math.floor((total % 3600) / 60);
    const secs = total % 60;
    const pad = (value) => String(value).padStart(2, '0');
    return hours > 0 ? `${hours}:${pad(minutes)}:${pad(secs)}` : `${pad(minutes)}:${pad(secs)}`;
}
