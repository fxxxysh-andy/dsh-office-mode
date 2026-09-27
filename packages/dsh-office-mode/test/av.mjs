/**
 * 音频与视频内容提取（office.av）的测试（第二十二轮）。
 *
 * 钉住四件事：
 *   1. **面在**：office.av.check / info / transcribe / frames / extract 挂在 SDK 上，
 *      设置能关掉它，缺件时的报错文案是可执行的（不是「功能不存在」）。
 *   2. **规范 WAV 的判据唯一**：只认 fmt 块 16 字节、data 落在偏移 36 的
 *      16 kHz 单声道 PCM16 —— SAPI 的 fmt18 头与 ffmpeg 默认封装（LIST 块）
 *      都被判不合格，且能就地规范化。这是整条链最容易被误判成「模型不能用」的地方。
 *   3. **纯函数**：分块、抽帧时间点、逐句拼接、时间戳格式。
 *   4. **真跑**：ffmpeg 造一段测试视频，走 info → transcribe → frames → extract，
 *      断言帧是真 JPEG、逐字稿文件真落盘、同一份文件第二次转写命中缓存。
 *
 * ffmpeg / SenseVoice 模型 / sherpa-onnx 运行时都是外部依赖：缺任何一样时相关断言
 * 标记 SKIP 并如实说明（与 test/python.mjs 对解释器、test/pdf.mjs 对渲染引擎同一口径）。
 *
 * 跑法：node test/av.mjs
 */
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { resolveConfig } from '../src/config.js';
import { executeRun } from '../src/run.js';
import { runPdfProcess } from '../src/pdf.js';
import {
    AV_ERROR_CODES,
    avSettings,
    canonicalizeWav,
    formatClock,
    inspectWav,
    joinTranscript,
    planChunks,
    planFrameTimes,
    probeAv,
    resetAvProbe,
} from '../src/av.js';

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
function cleanup(dir) {
    try {
        rmSync(dir, { recursive: true, force: true });
    } catch {
        // Windows 上被占住的目录留给系统临时目录清理
    }
}

const root = mkdtempSync(join(tmpdir(), 'office-av-'));
const probe = await probeAv(resolveConfig({}));
const ready = probe.available === true;

// ── 1. 规范 WAV 的判据与修复 ────────────────────────────────────────────────

/** 造一个 16 kHz 单声道 PCM16 WAV；`extraChunk` 会在 fmt 之后插一个块（模拟 LIST/fact）。 */
function makeWav({ seconds = 0.5, sampleRate = 16_000, channels = 1, bits = 16, extraChunk = false, brokenRiffSize = false } = {}) {
    const frames = Math.round(seconds * sampleRate);
    const data = Buffer.alloc(frames * channels * (bits / 8));
    for (let i = 0; i < data.length; i += 2) data.writeInt16LE(Math.round(Math.sin(i / 8) * 8000), i);
    const fmt = Buffer.alloc(16);
    fmt.writeUInt16LE(1, 0);
    fmt.writeUInt16LE(channels, 2);
    fmt.writeUInt32LE(sampleRate, 4);
    fmt.writeUInt32LE(sampleRate * channels * (bits / 8), 8);
    fmt.writeUInt16LE(channels * (bits / 8), 12);
    fmt.writeUInt16LE(bits, 14);
    const parts = [];
    parts.push(Buffer.from('fmt ', 'ascii'), (() => { const size = Buffer.alloc(4); size.writeUInt32LE(16, 0); return size; })(), fmt);
    if (extraChunk) parts.push(Buffer.from('LIST', 'ascii'), (() => { const size = Buffer.alloc(4); size.writeUInt32LE(6, 0); return size; })(), Buffer.from('INFOxx', 'ascii'));
    parts.push(Buffer.from('data', 'ascii'), (() => { const size = Buffer.alloc(4); size.writeUInt32LE(data.length, 0); return size; })(), data);
    const body = Buffer.concat(parts);
    const header = Buffer.alloc(12);
    header.write('RIFF', 0, 'ascii');
    header.writeUInt32LE(brokenRiffSize ? body.length + 99 : body.length + 4, 4);
    header.write('WAVE', 8, 'ascii');
    return Buffer.concat([header, body]);
}

await check('inspectWav：规范 WAV 通过，并算出时长', () => {
    const buffer = makeWav({ seconds: 1 });
    const info = inspectWav(buffer);
    assert.equal(info.ok, true, info.reason);
    assert.equal(info.sampleRate, 16_000);
    assert.equal(info.channels, 1);
    assert.equal(info.seconds, 1);
    return `${buffer.length} B / ${info.seconds} 秒`;
});

await check('inspectWav：fmt 块不是 16 字节（SAPI 的形状）被判不合格', () => {
    const buffer = makeWav({ seconds: 0.2 });
    // 把 fmt 块大小改成 18（SAPI 会带 cbSize），data 随之后移 —— 模拟真实失败形态。
    const patched = Buffer.concat([buffer.subarray(0, 16), (() => { const size = Buffer.alloc(4); size.writeUInt32LE(18, 0); return size; })(), Buffer.from([0, 0]), buffer.subarray(20)]);
    const info = inspectWav(patched);
    assert.equal(info.ok, false);
    assert.match(info.reason, /fmt 块大小/);
    return info.reason;
});

await check('inspectWav：多一个块把 data 推到偏移 36 之后（ffmpeg 默认封装）被判不合格', () => {
    const buffer = makeWav({ seconds: 0.2, extraChunk: true });
    const info = inspectWav(buffer);
    assert.equal(info.ok, false);
    assert.match(info.reason, /data 块不在偏移 36/);
    return info.reason;
});

await check('inspectWav：立体声 / 采样率不对 / RIFF 长度不符都被挡下', () => {
    assert.match(inspectWav(makeWav({ channels: 2 })).reason, /声道数/);
    assert.match(inspectWav(makeWav({ sampleRate: 44_100 })).reason, /采样率/);
    assert.match(inspectWav(makeWav({ brokenRiffSize: true })).reason, /RIFF 声明长度/);
    assert.equal(inspectWav(Buffer.alloc(10)).ok, false);
    return '四类坏形状各有各的原因';
});

await check('canonicalizeWav：带额外块的 WAV 能重写成规范头，且 PCM 一字节不丢', () => {
    const source = makeWav({ seconds: 0.3, extraChunk: true });
    const fixed = canonicalizeWav(source);
    const info = inspectWav(fixed);
    assert.equal(info.ok, true, info.reason);
    assert.equal(info.seconds, 0.3);
    assert.equal(fixed.length, source.length - 14); // LIST 块（8 字节头 + 6 字节体）被去掉
    return `${source.length} → ${fixed.length} B`;
});

await check('canonicalizeWav：不是 16 kHz 单声道 PCM16 时明确报错（不做隐式重采样）', () => {
    assert.throws(() => canonicalizeWav(makeWav({ sampleRate: 44_100 })), /不是 16 kHz 单声道 PCM16/);
    assert.throws(() => canonicalizeWav(Buffer.from('not a wav at all')), /不是 RIFF\/WAVE/);
    return '两类输入各自报清原因';
});

// ── 2. 纯函数：分块 / 抽帧时间点 / 拼接 / 时间戳 ──────────────────────────────

await check('planChunks：按块切分并覆盖到尾（最后一块可以更短）', () => {
    const chunks = planChunks(250, 120);
    assert.deepEqual(chunks.map((item) => [item.start, item.end]), [[0, 120], [120, 240], [240, 250]]);
    assert.equal(chunks.length, 3);
    assert.deepEqual(planChunks(0, 120), [{ index: 0, start: 0, end: 0, seconds: 0 }]);
    return '250 秒 → 3 块';
});

await check('planFrameTimes：at > every > count，且 from/to 能限制区间', () => {
    assert.deepEqual(planFrameTimes(60, { at: [30, 5, 30] }), [5, 30]);
    const every = planFrameTimes(10, { every: 4 });
    assert.deepEqual(every, [0, 4, 8]);
    const count = planFrameTimes(100, { count: 4 });
    assert.equal(count.length, 4);
    assert.ok(count[0] > 0 && count[count.length - 1] < 100, '均匀铺开不该取到首尾端点');
    const ranged = planFrameTimes(100, { count: 2, from: 20, to: 40 });
    assert.ok(ranged.every((t) => t >= 20 && t <= 40), `from/to 没生效：${ranged.join(',')}`);
    assert.throws(() => planFrameTimes(10, { at: ['x'] }), /没有有效的时间点/);
    return `count=4 → ${count.join(',')}`;
});

await check('joinTranscript：中日韩之间不加空格，拉丁语之间加一个空格', () => {
    assert.equal(joinTranscript(['今天下午三点', '开会。']), '今天下午三点开会。');
    assert.equal(joinTranscript(['Hello', 'world']), 'Hello world');
    assert.equal(joinTranscript(['第一句', 'second part']), '第一句second part');
    assert.equal(joinTranscript(['', '  ', '有内容']), '有内容');
    return '中英混排的拼接规则';
});

await check('formatClock：分秒与小时两种形态', () => {
    assert.equal(formatClock(0), '00:00');
    assert.equal(formatClock(75), '01:15');
    assert.equal(formatClock(3725), '1:02:05');
    return `${formatClock(3725)}`;
});

// ── 3. 配置与探测 ───────────────────────────────────────────────────────────

await check('avSettings：默认值来自「设备上本来就有」，坏值在区间内收敛', () => {
    const settings = avSettings({});
    assert.equal(settings.enabled, true);
    assert.equal(settings.language, 'auto');
    assert.equal(settings.chunkSeconds, 120);
    assert.equal(settings.maxSeconds, 3600);
    assert.equal(settings.precision, 'int8');
    assert.ok(settings.modelDir.includes('speech-to-text'), '默认模型目录要指向语音输入下载的那份');
    const clamped = avSettings({ av: { chunkSeconds: 9999, language: 'xx', precision: 'q4', threads: 99 } });
    assert.equal(clamped.chunkSeconds, 600);
    assert.equal(clamped.language, 'auto');
    assert.equal(clamped.precision, 'int8');
    assert.equal(clamped.threads, 16);
    return `模型目录 ${settings.modelDir}`;
});

await check('probeAv：路径全写错时报出缺什么（不静默降级）', async () => {
    resetAvProbe();
    const broken = await probeAv(resolveConfig({ av: { ffmpegPath: 'C:\\definitely\\missing\\ffmpeg.exe', ffprobePath: 'C:\\definitely\\missing\\ffprobe.exe', modelDir: 'C:\\definitely\\missing\\models', vadModel: 'C:\\definitely\\missing\\vad.onnx' } }));
    assert.equal(broken.available, false);
    for (const item of ['ffmpeg', 'ffprobe', 'SenseVoice 模型']) {
        assert.ok(broken.missing.includes(item), `missing 里该有「${item}」，实际 ${broken.missing.join(' / ')}`);
    }
    assert.match(broken.hint, /装 ffmpeg/);
    assert.match(broken.hint, /语音输入/);
    resetAvProbe();
    return `缺：${broken.missing.join(' / ')}`;
});

await check('错误码漂移守卫：src/av.js 里抛出的每个码都登记在 AV_ERROR_CODES', () => {
    const source = readFileSync(new URL('../src/av.js', import.meta.url), 'utf8');
    const found = [...source.matchAll(/avError\(\s*'([A-Z_]+)'/g)].map((match) => match[1]);
    assert.ok(found.length >= 8, `只扫到 ${found.length} 个错误码，正则可能坏了`);
    const unique = [...new Set(found)];
    for (const code of unique) assert.ok(AV_ERROR_CODES.includes(code), `${code} 没有登记进 AV_ERROR_CODES`);
    assert.ok(AV_ERROR_CODES.length >= unique.length, '登记表不能比实际用到的少');
    return `${unique.length} 个码全部已登记`;
});

await check('SDK 面在：office.av.check() 经 office_run 拿得到能力清单', async () => {
    const script = 'return await office.av.check();';
    const result = await executeRun({ script }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({}));
    assert.equal(result.ok, true, result.error?.message);
    const value = result.returned;
    assert.ok(value !== null && typeof value === 'object', 'office.av.check 要有返回值');
    for (const key of ['available', 'missing', 'ffmpeg', 'ffprobe', 'sensevoice', 'chunkSeconds', 'maxSeconds', 'maxFrames', 'hint']) {
        assert.ok(Object.prototype.hasOwnProperty.call(value, key), `check() 缺少 ${key}`);
    }
    assert.ok(Array.isArray(value.missing), 'missing 要是数组');
    return `available=${value.available}`;
});

await check('设置能关掉它：av.enabled=false 时 office.av 明确报错并指到设置页', async () => {
    const script = 'return await office.av.info("whatever.mp4");';
    const result = await executeRun({ script }, { agent: { session: { header: { cwd: root } } } }, resolveConfig({ av: { enabled: false } }));
    assert.equal(result.ok, false);
    assert.match(result.error.message, /设置里关闭/);
    return result.error.message.slice(0, 60);
});

// ── 4. 真跑（外部依赖缺失时跳过） ───────────────────────────────────────────

if (!ready) {
    results.push({ name: '真跑（ffmpeg / SenseVoice 模型 / sherpa-onnx 运行时）', ok: true, note: `SKIP ${probe.hint}` });
} else {
    const mediaDir = join(root, 'av-media');
    const video = join(mediaDir, 'meeting.mp4');
    const audio = join(mediaDir, 'tone.mp3');
    const silentVideo = join(mediaDir, 'silent.mp4');
    const bin = probe.ffmpeg.path;
    const runFfmpeg = (args) => runPdfProcess(bin, args, {
        outPath: join(mediaDir, 'ffmpeg.out.txt'),
        errPath: join(mediaDir, 'ffmpeg.err.txt'),
        timeoutMs: 120_000,
    });

    await check('造测试媒体：带声音的视频 / 只有声音的音频 / 没有声音的视频', async () => {
        const { mkdirSync } = await import('node:fs');
        mkdirSync(mediaDir, { recursive: true });
        const one = await runFfmpeg(['-hide_banner', '-loglevel', 'error', '-y', '-f', 'lavfi', '-i', 'color=c=navy:s=320x240:d=4', '-f', 'lavfi', '-i', 'sine=frequency=440:duration=4', '-shortest', '-c:v', 'libx264', '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-map_metadata', '-1', video]);
        assert.equal(one.code, 0, `造视频失败：${one.err}`);
        const two = await runFfmpeg(['-hide_banner', '-loglevel', 'error', '-y', '-f', 'lavfi', '-i', 'sine=frequency=440:duration=3', '-ac', '1', '-ar', '16000', '-c:a', 'libmp3lame', audio]);
        assert.equal(two.code, 0, `造音频失败：${two.err}`);
        const three = await runFfmpeg(['-hide_banner', '-loglevel', 'error', '-y', '-f', 'lavfi', '-i', 'color=c=black:s=320x240:d=2', '-c:v', 'libx264', '-pix_fmt', 'yuv420p', silentVideo]);
        assert.equal(three.code, 0, `造无声视频失败：${three.err}`);
        return '3 份测试媒体';
    });

    /** 在临时工作目录里跑一段 office_run 脚本。 */
    const runScript = (script, extra = {}) => executeRun({ script }, { agent: { session: { header: { cwd: root } } } }, resolveConfig(extra));

    await check('office.av.info：认出视频的两条轨与时长', async () => {
        const result = await runScript(`return await office.av.info(${JSON.stringify(video)});`);
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.equal(value.kind, 'video');
        assert.equal(value.hasAudio, true);
        assert.equal(value.hasVideo, true);
        assert.ok(value.durationSeconds > 3 && value.durationSeconds < 5, `时长异常：${value.durationSeconds}`);
        assert.equal(value.video[0].width, 320);
        return `${value.kind} ${value.durationClock} ${value.video[0].codec}/${value.audio[0].codec}`;
    });

    await check('office.av.transcribe：无声的音频也能走完整条链（识别出 0 句不报错）', async () => {
        const result = await runScript(`return await office.av.transcribe(${JSON.stringify(audio)}, { language: 'zh' });`);
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.equal(value.ok, true);
        assert.ok(Array.isArray(value.segments), 'segments 要是数组');
        assert.ok(value.audioSeconds > 2.5, `解码时长异常：${value.audioSeconds}`);
        assert.equal(typeof value.text, 'string');
        assert.ok(value.model.includes('SenseVoice'), `模型名异常：${value.model}`);
        return `语音 ${value.audioSeconds}s / ${value.segmentCount} 句 / 推理 ${value.inferenceSeconds}s`;
    });

    await check('office.av.transcribe：out 落成带时间戳的 Markdown 逐字稿', async () => {
        const out = 'av-out/transcript.md';
        const result = await runScript(`return await office.av.transcribe(${JSON.stringify(video)}, { language: 'zh', out: ${JSON.stringify(out)} });`);
        assert.equal(result.ok, true, result.error?.message);
        const written = result.returned.out;
        assert.ok(written !== undefined, 'out 没写出来');
        const text = readFileSync(join(root, out), 'utf8');
        assert.match(text, /转写稿/);
        assert.match(text, /\| # \| 时间 \| 文本 \|/);
        assert.ok(written.chars > 50, `逐字稿太短：${written.chars}`);
        return `${written.chars} 字 → ${out}`;
    });

    await check('office.av.transcribe：同一份文件第二次命中解码缓存', async () => {
        // 用一份新拷贝：同一份文件在前面的用例里已经解码过，缓存命中的对比要干净。
        const { copyFileSync } = await import('node:fs');
        const fresh = join(mediaDir, 'meeting-reuse.mp4');
        copyFileSync(video, fresh);
        const script = `return await office.av.transcribe(${JSON.stringify(fresh)}, { language: 'zh' });`;
        const first = await runScript(script);
        const second = await runScript(script);
        assert.equal(first.ok, true, first.error?.message);
        assert.equal(second.ok, true, second.error?.message);
        assert.equal(first.returned.audio.reused, false, '第一次不该是缓存命中');
        assert.equal(second.returned.audio.reused, true, '第二次该命中缓存');
        return `第二次 reused=${second.returned.audio.reused}`;
    });

    await check('office.av.frames：抽出来的确实是 JPEG，且时间点在区间内', async () => {
        const result = await runScript(`return await office.av.frames(${JSON.stringify(video)}, { count: 2, width: 160 });`);
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.equal(value.count, 2);
        assert.equal(value.format, 'jpg');
        for (const file of value.files) {
            const bytes = readFileSync(join(root, file.path));
            assert.equal(bytes[0], 0xff, `${file.path} 不是 JPEG`);
            assert.equal(bytes[1], 0xd8, `${file.path} 不是 JPEG`);
            assert.ok(file.at > 0 && file.at < 4, `时间点越界：${file.at}`);
        }
        return `${value.count} 张 → ${value.dir}`;
    });

    await check('office.av.extract：一条调用同时拿到逐字稿与画面', async () => {
        const out = 'av-out/extract.md';
        const result = await runScript(`return await office.av.extract(${JSON.stringify(video)}, { language: 'zh', out: ${JSON.stringify(out)}, count: 2 });`);
        assert.equal(result.ok, true, result.error?.message);
        const value = result.returned;
        assert.ok(value.transcript !== null, 'extract 该带回 transcript');
        assert.ok(value.frames !== null, 'extract 该带回 frames');
        assert.equal(value.frames.count, 2);
        assert.match(readFileSync(join(root, out), 'utf8'), /转写稿/);
        return `transcript + ${value.frames.count} 帧`;
    });

    await check('office.av.frames：音频文件上抽帧明确报「没有视频轨」', async () => {
        const result = await runScript(`return await office.av.frames(${JSON.stringify(audio)}, {});`);
        assert.equal(result.ok, false);
        assert.match(result.error.message, /没有视频轨/);
        return result.error.message.slice(0, 40);
    });

    await check('office.av.transcribe：没有音频轨的视频明确报「没有音频轨」', async () => {
        const result = await runScript(`return await office.av.transcribe(${JSON.stringify(silentVideo)}, {});`);
        assert.equal(result.ok, false);
        assert.match(result.error.message, /没有音频轨/);
        return result.error.message.slice(0, 40);
    });

    await check('找不到文件时给的是可执行的错（而不是 ffmpeg 的原始报错）', async () => {
        const result = await runScript('return await office.av.info("不存在的录音.m4a");');
        assert.equal(result.ok, false);
        assert.match(result.error.message, /找不到文件/);
        return result.error.message.slice(0, 40);
    });
}

cleanup(root);

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    if (item.ok) console.log(`PASS  ${item.name}${item.note ? `  → ${item.note}` : ''}`);
    else console.log(`FAIL  ${item.name}  → ${item.error?.message ?? item.error}`);
}
console.log(`\nav: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
