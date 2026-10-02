/**
 * 音频与视频的内容提取（office.av）。
 *
 * 办公模式此前只能读「已经是文字或图片」的输入：PDF 走 office.pdf、图片走
 * read_image。会议录音、访谈视频、手机录的语音条这类**只有声音**的输入没有
 * 任何出口。这一层补的就是它，两件事：
 *
 *   - **声音 → 文字**：SenseVoice（本机语音输入下载下来的那份 ONNX 模型）
 *     离线转写，VAD 切句并给每句时间戳；
 *   - **画面 → 图片**：视频按时间点抽帧成 JPEG / PNG，交给 read_image 看 ——
 *     和 PDF 图像型那条路一样，是读图能力，不是 OCR。
 *
 * ## 为什么自己去探测 ffmpeg 与 SenseVoice，而不是借语音输入那条链
 *
 * 语音输入插件（`@deepseek-ai/dsh-experimental-speech-to-text-sensevoice`）已经
 * 把模型与原生运行时装在了本机，但它的服务形态是**为麦克风录音设计的**：
 * 一次请求一段规范 16 kHz 单声道 PCM16 WAV、默认上限 4 MiB（≈131 秒）、
 * 回给调用方的只有 `text`（情感与事件标签被它自己的 strict 契约丢掉）。
 * 办公场景要的是「任意格式的长文件 + 分块 + 时间戳 + 尽量不改动宿主」，
 * 所以这里直接驱动同一份模型：
 *
 *   - 解码与规范化交给 **ffmpeg**（`-map_metadata -1 -fflags +bitexact` 才能
 *     得到 data 偏移 36 的规范 WAV，否则后面那层会拒收），随后自己再验一遍；
 *   - 推理交给 `sherpa-onnx-node`（宿主 profile 里已随语音输入装好，**用宿主
 *     解析位置 require**，本插件因此仍然是零依赖）；
 *   - 推理跑在**独立子进程**里（`asr-worker.mjs`），不占宿主的事件循环 ——
 *     一段十分钟的音频在 2 线程下要几秒到几十秒，放主进程里会卡住整个会话。
 *
 * ## 分块：重叠 + 边界回退（第三十一轮）
 *
 * 长音频要切成块才能控住内存与单次超时，但块是硬切的，VAD 又在块内独立跑 ——
 * 切点落在句子中间时两边各拿到半句。第三十一轮起两条措施一起上：
 *
 *   1. 相邻块留 `overlapSeconds`（默认 0.5 秒）的重叠：跨在切点上的短音在同一块
 *      里被完整解出来（`planChunks` 的 `end` 比 `ownedEnd` 多出这一段）；
 *   2. 跨得更深、在窗口末尾被切断的句子，本块不出稿，把下一块的起点回退到它的
 *      话头，让它整句在下一块里解一遍（`resolveChunkBoundary`）—— 不丢字、不重复。
 *
 * 只做重叠是不够的：重叠区之外的长句会在两块里各断一次、逐字稿里出现重复片段。
 *
 * ## 默认值来自「设备上本来就有」
 *
 * 默认认为设备上有 ffmpeg 与 SenseVoice（用户口径）。四样东西缺任何一个都
 * **明确报错并说清怎么补**，不静默降级、不假装支持：ffmpeg、ffprobe、
 * 模型目录、sherpa-onnx 运行时。`office.av.check()` 一次报全。
 *
 * @module dsh-office-mode/av
 */
import { createHash } from 'node:crypto';
import { closeSync, existsSync, openSync, readFileSync, readSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { displayPath } from './engine/kit.js';
// 与 python.js / tex.js 同一个进程执行器：受限沙箱下 stdio:'pipe' 会 spawn EPERM，
// 统一把 stdout/stderr 重定向到文件句柄。
import { findPdfExecutable as findExecutable, runPdfProcess as runProcess } from './pdf.js';

/** 认得的音频扩展名（真正的判据是 ffprobe 说有音频流，这只是给报错用的初判）。 */
export const AV_AUDIO_EXTENSIONS = new Set([
    '.wav', '.mp3', '.m4a', '.aac', '.flac', '.ogg', '.oga', '.opus', '.wma', '.amr', '.aif', '.aiff', '.ape', '.mka',
]);
/** 认得的视频扩展名。 */
export const AV_VIDEO_EXTENSIONS = new Set([
    '.mp4', '.m4v', '.mov', '.mkv', '.avi', '.webm', '.flv', '.wmv', '.mpg', '.mpeg', '.ts', '.m2ts', '.3gp', '.rmvb',
]);
/** 转写语言提示（与 SenseVoice 支持的语言一致）。 */
export const AV_LANGUAGES = ['auto', 'zh', 'en', 'yue', 'ja', 'ko'];
/** 单块默认秒数：官方那条链默认上限 131 秒，这里取 120 秒留余量。 */
export const AV_CHUNK_SECONDS = 120;
/**
 * 相邻两块之间的默认重叠秒数（第三十一轮）。
 *
 * 为什么要有它：块是硬切开的，VAD 又在块内独立跑 —— 一句话正好跨在切点上时，
 * 两边各拿到半句，转写结果是两段都不像话的碎片。给块尾巴留 0.5 秒的重叠，
 * 跨在切点上的短音（一个词、一个音节）就能在**同一块**里被完整解出来。
 * 为什么是 0.5 秒：它只买「切点附近那一点音频的余量」，不是拿重叠换上下文；
 * 重叠越长，重复解码的量越大（每块多算 overlap/chunk 的比例），而真正
 * 「一句话不许被切开」的保证由 `resolveChunkBoundary` 的边界回退给出（见下），
 * 不靠把这个数字调大。
 */
export const AV_CHUNK_OVERLAP_SECONDS = 0.5;
/** 单次处理的音频时长上限（秒）：超过就报错并要求先切，而不是把内存吃满。 */
export const AV_MAX_SECONDS = 3600;
/** 单次转写 / 抽帧的默认超时（毫秒）。 */
export const AV_TIMEOUT_MS = 300_000;
/** 单次抽帧数上限。 */
export const AV_MAX_FRAMES = 12;
/** 抽帧默认张数（不给 at / every / count 时）。 */
export const AV_DEFAULT_FRAMES = 6;
/** 规范的 16 kHz 单声道 PCM16 WAV 头长度。 */
export const AV_WAV_HEADER_BYTES = 44;

/** PATH 里没有 ffmpeg 时再找这些常见安装目录。 */
const AV_EXTRA_BIN_DIRS = ['C:\\app\\ffmpeg\\bin', 'C:\\ffmpeg\\bin', 'C:\\Program Files\\ffmpeg\\bin', '/usr/local/bin', '/opt/homebrew/bin'];

/** 子进程里跑的推理入口（与本源文件同目录）。 */
export const AV_WORKER_PATH = fileURLToPath(new URL('./asr-worker.mjs', import.meta.url));

/** 带错误码的错误：测试与反馈文案都按码认，不按字符串匹配。 */
function avError(code, message) {
    const error = new Error(message);
    error.code = code;
    return error;
}

/** 本插件抛出的全部错误码（test/av.mjs 有一条漂移守卫逐个断言已登记）。 */
export const AV_ERROR_CODES = [
    'OFFICE_AV_DISABLED',
    'OFFICE_AV_INPUT_MISSING',
    'OFFICE_AV_NO_FFMPEG',
    'OFFICE_AV_NO_FFPROBE',
    'OFFICE_AV_NO_MODEL',
    'OFFICE_AV_NO_RUNTIME',
    'OFFICE_AV_PROBE_FAILED',
    'OFFICE_AV_NO_AUDIO_STREAM',
    'OFFICE_AV_NO_VIDEO_STREAM',
    'OFFICE_AV_TOO_LONG',
    'OFFICE_AV_DECODE_FAILED',
    'OFFICE_AV_TRANSCRIBE_FAILED',
    'OFFICE_AV_FRAMES_FAILED',
    'OFFICE_AV_BAD_OPTION',
];

/** DSH 主目录：优先环境变量，其次 ~/.dsh。 */
export function dshHome() {
    const raw = typeof process.env.DSH_HOME === 'string' ? process.env.DSH_HOME.trim() : '';
    return raw !== '' ? raw : join(homedir(), '.dsh');
}

/**
 * 语音输入那份 SenseVoice 模型的默认位置。
 *
 * 这是上游语音输入 bundle 的 `dataRoot`（patch 里写的是
 * `dshHomePath('speech-to-text','sensevoice')`）+ 它自己的 `models/sensevoice-onnx`。
 * 换机器、或者把模型复制到别处时，用设置页的「模型目录」指过去即可。
 */
export function defaultSenseVoiceModelDir() {
    const override = typeof process.env.DSH_OFFICE_SENSEVOICE_DIR === 'string' ? process.env.DSH_OFFICE_SENSEVOICE_DIR.trim() : '';
    return override !== '' ? override : join(dshHome(), 'speech-to-text', 'sensevoice', 'models', 'sensevoice-onnx');
}

/** Silero VAD 的默认位置（与模型目录同一棵 trees）。 */
export function defaultSileroVadPath() {
    const override = typeof process.env.DSH_OFFICE_SILERO_VAD === 'string' ? process.env.DSH_OFFICE_SILERO_VAD.trim() : '';
    return override !== '' ? override : join(dshHome(), 'speech-to-text', 'sensevoice', 'models', 'silero', 'silero_vad.onnx');
}

/** 从配置里取 av 通道的参数（配置缺失时用默认值，绝不因为没配置就不可用）。 */
export function avSettings(config) {
    const raw = config?.av !== null && typeof config?.av === 'object' ? config.av : {};
    const num = (value, fallback, min, max) => {
        const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
        if (!Number.isFinite(n)) return fallback;
        return Math.min(max, Math.max(min, n));
    };
    const str = (value) => (typeof value === 'string' ? value.trim() : '');
    const language = AV_LANGUAGES.includes(str(raw.language)) ? str(raw.language) : 'auto';
    const chunkSeconds = num(raw.chunkSeconds, AV_CHUNK_SECONDS, 10, 600);
    // 重叠必须小于半块：等于或超过半块时「重叠」已经变成「同一段音频解两遍」，
    // 重叠区里的话会被两块各报一次。上限取块的一半，是能保住去重不变量的边界。
    const overlapSeconds = Math.min(num(raw.overlapSeconds, AV_CHUNK_OVERLAP_SECONDS, 0, 30), chunkSeconds / 2);
    return {
        enabled: raw.enabled !== false,
        ffmpegPath: str(raw.ffmpegPath),
        ffprobePath: str(raw.ffprobePath),
        modelDir: str(raw.modelDir) || defaultSenseVoiceModelDir(),
        vadModel: str(raw.vadModel) || defaultSileroVadPath(),
        runtimePath: str(raw.runtimePath),
        precision: raw.precision === 'fp32' ? 'fp32' : 'int8',
        threads: Math.round(num(raw.threads, 2, 1, 16)),
        language,
        chunkSeconds,
        overlapSeconds,
        maxSeconds: num(raw.maxSeconds, AV_MAX_SECONDS, 10, 21_600),
        timeoutMs: num(raw.timeoutMs, AV_TIMEOUT_MS, 5_000, 1_800_000),
        framesDir: str(raw.framesDir) || 'av/frames',
        audioDir: str(raw.audioDir) || 'av/audio',
        maxFrames: Math.round(num(raw.maxFrames, AV_MAX_FRAMES, 1, 60)),
        frames: Math.round(num(raw.frames, AV_DEFAULT_FRAMES, 1, 60)),
        frameFormat: raw.frameFormat === 'png' ? 'png' : 'jpg',
        frameWidth: Math.round(num(raw.frameWidth, 0, 0, 4096)),
        vadThreshold: num(raw.vadThreshold, 0.5, 0, 1),
        minSpeechSeconds: num(raw.minSpeechSeconds, 0.25, 0, 30),
        minSilenceSeconds: num(raw.minSilenceSeconds, 0.5, 0.01, 30),
        maxSegmentSeconds: num(raw.maxSegmentSeconds, 30, 1, 120),
        words: raw.words === true,
    };
}

// ── 可执行文件与宿主模块的解析 ────────────────────────────────────────────────

/**
 * 在配置路径 → PATH → 常见安装目录里找一个媒体工具。
 *
 * **配了路径就用它**（不存在就是缺件，不再退回 PATH）：显式配置写错还静默用另一个
 * ffmpeg，会让「我明明指了路径」和「实际跑的是哪一个」对不上，排查时最费时间。
 * 留空才走自动探测（PATH → 常见安装目录）。
 */
function findAvExecutable(names, configured = '') {
    const wanted = typeof configured === 'string' ? configured.trim() : '';
    if (wanted !== '') return existsSync(wanted) ? wanted : undefined;
    const onPath = findExecutable(names);
    if (onPath !== undefined) return onPath;
    for (const dir of AV_EXTRA_BIN_DIRS) {
        for (const name of names) {
            for (const ext of process.platform === 'win32' ? ['.exe', '.cmd', ''] : ['']) {
                const candidate = join(dir, `${name}${ext}`);
                try {
                    if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
                } catch {
                    // 目录不可读就跳过
                }
            }
        }
    }
    return undefined;
}

/**
 * 从宿主的安装位置解析一个包（与 index.js 的 resolveHostModule 同一套锚点 + 一层兜底）。
 *
 * 为什么需要：本插件是 link: 进 profile 的（源码在工作区），而 sherpa-onnx-node
 * 装在宿主那边。ESM 按模块自身位置解析裸包名必然失败，只能借宿主的解析位置。
 *
 * 为什么还要那层兜底：profile 的 node_modules 里只有**宿主直接依赖的几个包**
 * （实测本机是 @deepseek-ai/cosmokit 与 @deepseek-ai/schemastery），
 * sherpa-onnx-node 不在那儿 —— 它在 `@deepseek-ai/dsh` 自己的 node_modules 里。
 * 所以第一次解析失败时，先解析 `@deepseek-ai/dsh/package.json`（那个一定解析得到），
 * 再从它那里解析目标包。
 */
const hostAssets = new Map();

/** 可用的解析锚点，按可靠性排序。 */
function hostAnchors() {
    const anchors = [
        process.env.DSH_PROFILE_DIR,
        process.env.DSH_HOST_ROOT,
        process.env.DSH_CHECKOUT,
        process.env.DSH_HOME,
    ].filter((value) => typeof value === 'string' && value !== '');
    if (typeof process.argv[1] === 'string' && process.argv[1] !== '') anchors.push(dirname(process.argv[1]));
    anchors.push(process.cwd());
    return [...new Set(anchors)];
}

/** 从一个锚点建一个 require（锚点不可用时返回 undefined）。 */
function hostRequire(anchor) {
    try {
        return createRequire(pathToFileURL(join(anchor, 'package.json')).href);
    } catch {
        return undefined;
    }
}

/** 在锚点链上解析一个包，都解析不到返回 undefined。 */
function resolveFromAnchors(specifier) {
    for (const anchor of hostAnchors()) {
        const req = hostRequire(anchor);
        if (req === undefined) continue;
        try {
            return req.resolve(specifier);
        } catch {
            // 换下一个锚点
        }
    }
    return undefined;
}

function resolveHostAsset(specifier) {
    if (hostAssets.has(specifier)) return hostAssets.get(specifier);
    let found = resolveFromAnchors(specifier);
    if (found === undefined) {
        const dshPackage = resolveFromAnchors('@deepseek-ai/dsh/package.json');
        if (dshPackage !== undefined) {
            try {
                found = createRequire(dshPackage).resolve(specifier);
            } catch {
                found = undefined;
            }
        }
    }
    hostAssets.set(specifier, found);
    return found;
}

/** 真跑一次 `-version`：PATH 里存在不等于能用（装了一半的 ffmpeg 很常见）。 */
async function probeTool(command, timeoutMs) {
    const stamp = `${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    const outPath = join(tmpdir(), `office-av-probe-${stamp}.out.txt`);
    const errPath = join(tmpdir(), `office-av-probe-${stamp}.err.txt`);
    const result = await runProcess(command, ['-version'], { outPath, errPath, timeoutMs });
    const text = `${result.out ?? ''}${result.err ?? ''}`.trim();
    const version = text.split('\n').map((line) => line.trim()).find((line) => line !== '') ?? '';
    return {
        ok: result.code === 0,
        version: version.slice(0, 200),
        error: typeof result.error === 'string' ? result.error : (result.code === 0 ? null : `退出码 ${String(result.code)}`),
    };
}

/** 按精度挑模型文件；两个都没有时报错用。 */
function modelFileFor(settings) {
    const wanted = settings.precision === 'fp32' ? 'model.onnx' : 'model.int8.onnx';
    const primary = join(settings.modelDir, wanted);
    if (existsSync(primary)) return primary;
    const alternative = join(settings.modelDir, settings.precision === 'fp32' ? 'model.int8.onnx' : 'model.onnx');
    return existsSync(alternative) ? alternative : primary;
}

/** 进程内探测缓存：同样的路径组合只探一次（设置里改了路径会因 key 变化重探）。 */
const probeCache = new Map();

/** 探测结果按需刷新（设置里改了 ffmpeg 或模型路径之后）。 */
export function resetAvProbe() {
    probeCache.clear();
}

/**
 * **同步**版的可用性判据（能力地图用，见 src/capabilities.js）。
 *
 * 为什么要有它：地图必须与 `office.av.check` 说同一件事。第四十七轮复核实测过两处
 * 不一致 —— 只按 `ffmpeg` 判时，`ffprobe` 配错也报「可用」（`probeAv` 的 missing 里有
 * `ffprobe`），而按 PATH 找 ffmpeg 时又漏掉 `AV_EXTRA_BIN_DIRS` 里的那份（报「缺件」
 * 而通道其实可用）。所以判据留在这里、与 `probeAv` 共用同一批解析函数：
 *
 *   - 五样（ffmpeg / ffprobe / SenseVoice 权重 + 词表 / Silero VAD / sherpa-onnx 运行时）
 *     齐全才算 `available`，与 `probeAv().missing.length === 0` 同一口径；
 *   - **不真跑** ffmpeg / ffprobe（那是 `probeAv` 的异步部分），所以「找得到」不等于
 *     「跑得起来」——地图末尾那句「能不能真跑看各自的 check」就是这条边界；
 *   - `enabled:false` 照实报关闭，不假装有能力。
 *
 * @param {object} [config] 插件配置（取 av 组）
 * @returns {{enabled: boolean, ffmpeg: boolean, ffprobe: boolean, model: boolean, vad: boolean, runtime: boolean, available: boolean, missing: string[]}}
 */
export function probeAvSync(config = {}) {
    const settings = avSettings(config);
    if (settings.enabled === false) {
        return { enabled: false, ffmpeg: false, ffprobe: false, model: false, vad: false, runtime: false, available: false, missing: ['已关闭'] };
    }
    const ffmpeg = findAvExecutable(['ffmpeg'], settings.ffmpegPath) !== undefined;
    const ffprobe = findAvExecutable(['ffprobe'], settings.ffprobePath) !== undefined;
    const model = existsSync(modelFileFor(settings)) && existsSync(join(settings.modelDir, 'tokens.txt'));
    const vad = existsSync(settings.vadModel);
    // 与 probeAv 同一口径：配置路径优先，否则从宿主锚点解析，且**必须真的存在**。
    const resolved = settings.runtimePath !== '' && existsSync(settings.runtimePath)
        ? settings.runtimePath
        : resolveHostAsset('sherpa-onnx-node');
    const runtime = typeof resolved === 'string' && resolved !== '' && existsSync(resolved);
    const missing = [];
    if (!ffmpeg) missing.push('ffmpeg');
    if (!ffprobe) missing.push('ffprobe');
    if (!model) missing.push('SenseVoice 模型');
    if (!vad) missing.push('Silero VAD');
    if (!runtime) missing.push('sherpa-onnx 运行时');
    return {
        enabled: true,
        ffmpeg,
        ffprobe,
        model,
        vad,
        runtime,
        available: missing.length === 0,
        missing,
    };
}

/**
 * 探测本机的「能读到什么程度」。
 *
 * @param {object} [config] 插件配置（取 av 组）
 * @returns {Promise<object>} 四项能力的可用性与版本，缺什么列在 missing 里
 */
export async function probeAv(config = {}) {
    const settings = avSettings(config);
    const key = JSON.stringify([
        settings.ffmpegPath, settings.ffprobePath, settings.modelDir, settings.vadModel,
        settings.runtimePath, settings.precision, settings.enabled,
    ]);
    if (probeCache.has(key)) return probeCache.get(key);

    const ffmpeg = findAvExecutable(['ffmpeg'], settings.ffmpegPath);
    const ffprobe = findAvExecutable(['ffprobe'], settings.ffprobePath);
    const ffmpegProbe = ffmpeg === undefined ? undefined : await probeTool(ffmpeg, 20_000);
    const ffprobeProbe = ffprobe === undefined ? undefined : await probeTool(ffprobe, 20_000);
    const runtime = settings.runtimePath !== '' && existsSync(settings.runtimePath)
        ? settings.runtimePath
        : resolveHostAsset('sherpa-onnx-node');
    const model = modelFileFor(settings);
    const tokens = join(settings.modelDir, 'tokens.txt');
    const modelOk = existsSync(model) && existsSync(tokens);
    const vadOk = existsSync(settings.vadModel);
    const runtimeOk = typeof runtime === 'string' && runtime !== '' && existsSync(runtime);

    const missing = [];
    if (ffmpegProbe?.ok !== true) missing.push('ffmpeg');
    if (ffprobeProbe?.ok !== true) missing.push('ffprobe');
    if (!modelOk) missing.push('SenseVoice 模型');
    if (!vadOk) missing.push('Silero VAD');
    if (!runtimeOk) missing.push('sherpa-onnx 运行时');

    const value = {
        enabled: settings.enabled,
        available: missing.length === 0,
        missing,
        ffmpeg: { path: ffmpeg ?? null, ok: ffmpegProbe?.ok === true, version: ffmpegProbe?.version ?? '', error: ffmpegProbe?.error ?? null },
        ffprobe: { path: ffprobe ?? null, ok: ffprobeProbe?.ok === true, version: ffprobeProbe?.version ?? '', error: ffprobeProbe?.error ?? null },
        sensevoice: {
            modelDirectory: settings.modelDir,
            model: modelOk ? model : null,
            modelPath: model,
            tokens,
            vad: settings.vadModel,
            runtime: runtimeOk ? runtime : null,
            precision: settings.precision,
            threads: settings.threads,
            languages: [...AV_LANGUAGES],
        },
        language: settings.language,
        chunkSeconds: settings.chunkSeconds,
        overlapSeconds: settings.overlapSeconds,
        maxSeconds: settings.maxSeconds,
        hint: missing.length === 0
            ? 'ffmpeg / ffprobe / SenseVoice 模型 / sherpa-onnx 运行时都在，可以转写与抽帧。'
            : missingFfmpegHint(missing, settings),
    };
    probeCache.set(key, value);
    return value;
}

/** 缺件时的说明：一次把「装什么、在哪儿改」讲清楚。 */
function missingFfmpegHint(missing, settings) {
    const lines = [`视频与音频内容提取当前不可用，缺：${missing.join(' / ')}。`];
    if (missing.includes('ffmpeg') || missing.includes('ffprobe')) {
        if (settings.ffmpegPath !== '' || settings.ffprobePath !== '') {
            lines.push(`设置里填的路径不存在（ffmpeg：${settings.ffmpegPath || '留空'}；ffprobe：${settings.ffprobePath || '留空'}）——`
                + '路径填了就以它为准，不再退回 PATH；改对或者清空让它自动探测。');
        }
        lines.push('装 ffmpeg（它的 bin 目录里同时有 ffprobe）并确保在 PATH 里，或者在本插件设置页的「音频与视频」里填绝对路径。');
    }
    if (missing.includes('SenseVoice 模型') || missing.includes('Silero VAD')) {
        lines.push(`SenseVoice 模型默认取语音输入下载的那份：${settings.modelDir}（缺 model.int8.onnx / tokens.txt），VAD 在 ${settings.vadModel}。`
            + '装一次语音输入（@deepseek-ai/dsh-experimental-voice-input-bundle）它会自己下载；也可以把模型复制到别处后在设置页改「模型目录」。');
    }
    if (missing.includes('sherpa-onnx 运行时')) {
        lines.push('sherpa-onnx-node 随语音输入一起装在宿主的 node_modules 里；本插件按宿主安装位置解析它，解析不到就说明它没装。');
    }
    return lines.join('\n');
}

/** 缺件时统一从这里抛，报错文案与 check() 的 hint 是同一份。 */
async function requireCapability(config) {
    const probe = await probeAv(config);
    if (probe.enabled === false) {
        throw avError('OFFICE_AV_DISABLED', '音频与视频的内容提取已在本插件的设置里关闭（「音频与视频」→ 启用）。');
    }
    if (probe.ffmpeg.ok !== true) {
        throw avError('OFFICE_AV_NO_FFMPEG', missingFfmpegHint(['ffmpeg'], avSettings(config)));
    }
    if (probe.ffprobe.ok !== true) {
        throw avError('OFFICE_AV_NO_FFPROBE', missingFfmpegHint(['ffprobe'], avSettings(config)));
    }
    return probe;
}

/** 转写前再确认模型与运行时；缺件时分别给码，便于反馈分层。 */
function requireSpeech(probe, settings) {
    if (probe.sensevoice.model === null) {
        throw avError('OFFICE_AV_NO_MODEL', missingFfmpegHint(['SenseVoice 模型'], settings));
    }
    if (probe.sensevoice.runtime === null) {
        throw avError('OFFICE_AV_NO_RUNTIME', missingFfmpegHint(['sherpa-onnx 运行时'], settings));
    }
}

// ── WAV 的规范形状 ───────────────────────────────────────────────────────────

/**
 * 按固定偏移解析 WAV 头（与宿主语音输入那条链同一套判据）。
 *
 * 判据必须是**唯一**的：fmt 块大小恰为 16、PCM、单声道、16000 Hz、16 bit、
 * `data` 落在偏移 36。多一个 fact / LIST 块都会让 `data` 挪位，后面那层直接
 * 判「不是规范 WAV」—— 实测 SAPI 的默认输出（fmt 18）与 ffmpeg 的默认封装
 * （LIST 在 36，data 在 70）都属于这类，所以这里既验也修。
 *
 * @param {Buffer} buffer 文件（或文件头）字节
 * @param {number} [totalBytes] 文件总长度（只给了头部时要传）
 */
export function inspectWav(buffer, totalBytes) {
    const total = Number.isFinite(totalBytes) ? Number(totalBytes) : buffer.length;
    const fail = (reason) => ({
        ok: false, reason, channels: null, sampleRate: null, bitsPerSample: null, dataBytes: null, seconds: null, chunks: [],
    });
    if (buffer.length < 44) return fail('字节数不足 44，连头都不完整');
    if (buffer.toString('ascii', 0, 4) !== 'RIFF' || buffer.toString('ascii', 8, 12) !== 'WAVE') return fail('不是 RIFF/WAVE');
    if (buffer.toString('ascii', 12, 16) !== 'fmt ') return fail('第一个块不是 fmt');
    if (buffer.readUInt32LE(16) !== 16) return fail(`fmt 块大小是 ${buffer.readUInt32LE(16)}，必须是 16`);
    if (buffer.readUInt16LE(20) !== 1) return fail('不是 PCM（audioFormat ≠ 1）');
    const channels = buffer.readUInt16LE(22);
    const sampleRate = buffer.readUInt32LE(24);
    const byteRate = buffer.readUInt32LE(28);
    const blockAlign = buffer.readUInt16LE(32);
    const bitsPerSample = buffer.readUInt16LE(34);
    if (channels !== 1) return fail(`声道数是 ${channels}，必须是 1（单声道）`);
    if (sampleRate !== 16_000) return fail(`采样率是 ${sampleRate}，必须是 16000`);
    if (byteRate !== 32_000) return fail(`byteRate 是 ${byteRate}，必须是 32000`);
    if (blockAlign !== 2) return fail(`blockAlign 是 ${blockAlign}，必须是 2`);
    if (bitsPerSample !== 16) return fail(`位深是 ${bitsPerSample}，必须是 16`);
    if (buffer.toString('ascii', 36, 40) !== 'data') return fail('data 块不在偏移 36（前面多了别的块）');
    const dataBytes = buffer.readUInt32LE(40);
    const riffSize = buffer.readUInt32LE(4);
    if (riffSize !== total - 8) return fail(`RIFF 声明长度 ${riffSize} 与文件实际 ${total - 8} 不符`);
    if (dataBytes !== total - 44) return fail(`data 声明长度 ${dataBytes} 与文件实际 ${total - 44} 不符`);
    if ((total - 44) % 2 !== 0) return fail('data 长度不是偶数');
    return {
        ok: true,
        reason: '',
        channels,
        sampleRate,
        bitsPerSample,
        byteRate,
        blockAlign,
        dataBytes,
        seconds: (total - 44) / 32_000,
        chunks: [{ id: 'fmt ', offset: 12, size: 16 }, { id: 'data', offset: 36, size: dataBytes }],
    };
}

/**
 * 把任意可解析的 RIFF/WAVE 重写成规范 44 字节头（只保留 fmt 与 data）。
 * 不是 PCM16 单声道 16 kHz 时抛错并说明原因 —— 这一步不做重采样，
 * 采样率与声道的统一由 ffmpeg 负责。
 */
export function canonicalizeWav(buffer) {
    if (buffer.toString('ascii', 0, 4) !== 'RIFF' || buffer.toString('ascii', 8, 12) !== 'WAVE') {
        throw avError('OFFICE_AV_DECODE_FAILED', '不是 RIFF/WAVE，无法规范化。');
    }
    let offset = 12;
    let fmt;
    let data;
    const chunks = [];
    while (offset + 8 <= buffer.length) {
        const id = buffer.toString('ascii', offset, offset + 4);
        const size = buffer.readUInt32LE(offset + 4);
        const start = offset + 8;
        const end = Math.min(start + size, buffer.length);
        chunks.push({ id, offset, size });
        if (id === 'fmt ' && fmt === undefined) fmt = buffer.subarray(start, start + 16);
        if (id === 'data' && data === undefined) data = buffer.subarray(start, end);
        offset = start + size + (size % 2);
    }
    if (fmt === undefined || fmt.length < 16) throw avError('OFFICE_AV_DECODE_FAILED', 'WAV 里没有 fmt 块。');
    if (data === undefined) throw avError('OFFICE_AV_DECODE_FAILED', 'WAV 里没有 data 块。');
    const channels = fmt.readUInt16LE(2);
    const sampleRate = fmt.readUInt32LE(4);
    const bitsPerSample = fmt.readUInt16LE(14);
    if (fmt.readUInt16LE(0) !== 1 || channels !== 1 || sampleRate !== 16_000 || bitsPerSample !== 16) {
        throw avError('OFFICE_AV_DECODE_FAILED', `WAV 不是 16 kHz 单声道 PCM16（${sampleRate} Hz / ${channels} 声道 / ${bitsPerSample} bit），先转码再规范化。`);
    }
    const header = Buffer.alloc(AV_WAV_HEADER_BYTES);
    header.write('RIFF', 0, 'ascii');
    header.writeUInt32LE(36 + data.length, 4);
    header.write('WAVE', 8, 'ascii');
    header.write('fmt ', 12, 'ascii');
    header.writeUInt32LE(16, 16);
    header.writeUInt16LE(1, 20);
    header.writeUInt16LE(1, 22);
    header.writeUInt32LE(16_000, 24);
    header.writeUInt32LE(32_000, 28);
    header.writeUInt16LE(2, 32);
    header.writeUInt16LE(16, 34);
    header.write('data', 36, 'ascii');
    header.writeUInt32LE(data.length, 40);
    return Buffer.concat([header, data]);
}

/**
 * 盘上的 WAV 是不是已经是规范形状。
 *
 * 只读文件头（64 字节）再配合文件长度判断 —— 这里被反复调用（缓存命中检查、
 * 解码后复验），一次读进整个 PCM（一小时音频 115 MB）会让「命中缓存」这件事
 * 比重新解码还贵。
 */
export function isCanonicalWav(path) {
    let fd;
    try {
        const info = statSync(path);
        if (!info.isFile() || info.size < AV_WAV_HEADER_BYTES) return false;
        fd = openSync(path, 'r');
        const head = Buffer.alloc(Math.min(64, info.size));
        const read = readSync(fd, head, 0, head.length, 0);
        return inspectWav(head.subarray(0, read), info.size).ok;
    } catch {
        return false;
    } finally {
        if (fd !== undefined) {
            try {
                closeSync(fd);
            } catch {
                // 已经关掉了
            }
        }
    }
}

// ── 输入与命名 ───────────────────────────────────────────────────────────────

/** 把脚本给的路径收敛成绝对路径，并检查存在。 */
function resolveInput(filePath, env) {
    const wanted = typeof filePath === 'string' ? filePath.trim() : '';
    if (wanted === '') throw avError('OFFICE_AV_INPUT_MISSING', '要给出音频或视频文件的路径。');
    const root = resolve(env?.root ?? process.cwd());
    const absolute = isAbsolute(wanted) ? wanted : resolve(root, wanted);
    if (!existsSync(absolute)) {
        throw avError('OFFICE_AV_INPUT_MISSING', `找不到文件：${displayPath(root, absolute)}`);
    }
    let info;
    try {
        info = statSync(absolute);
    } catch (error) {
        throw avError('OFFICE_AV_INPUT_MISSING', `读不到文件：${displayPath(root, absolute)}（${error.message}）`);
    }
    if (!info.isFile()) throw avError('OFFICE_AV_INPUT_MISSING', `不是普通文件：${displayPath(root, absolute)}`);
    return { absolute, relative: displayPath(root, absolute), bytes: info.size, modifiedMs: info.mtimeMs, root };
}

/** 文件名 → 可放进缓存路径的一段（保留可读性，去掉危险字符）。 */
function slugOf(filePath) {
    const base = basename(String(filePath)).replace(/\.[^.]+$/, '');
    const cleaned = base.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-').replace(/^\.+/, '').trim();
    return (cleaned === '' ? 'media' : cleaned).slice(0, 60);
}

/** 内容键：源文件的大小与修改时间 + 影响解码结果的参数。 */
function contentKey(bytes, modifiedMs, extra) {
    return createHash('sha1').update(`${bytes}:${Math.round(modifiedMs)}:${extra}`).digest('hex').slice(0, 10);
}

/** 时间戳文本：`mm:ss` / `h:mm:ss`。 */
export function formatClock(seconds) {
    const total = Math.max(0, Math.round(Number(seconds) || 0));
    const hours = Math.floor(total / 3600);
    const minutes = Math.floor((total % 3600) / 60);
    const secs = total % 60;
    const pad = (n) => String(n).padStart(2, '0');
    return hours > 0 ? `${hours}:${pad(minutes)}:${pad(secs)}` : `${pad(minutes)}:${pad(secs)}`;
}

/** 临时文件路径（探测、ffprobe JSON 这类用完即弃的东西）。 */
function tempPath(tag) {
    return join(tmpdir(), `office-av-${tag}-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`);
}

/**
 * 成功时删掉这一次的进程日志。
 *
 * 缓存目录只该留有用的东西：ffmpeg 每次调用都会产出 stdout / stderr 两个文件，
 * 成功路径上它们是空的，留在缓存里只会让 `.office/cache/av/` 长出一堆 0 字节
 * 的噪声文件（还会被 office.cache.list 当成产物统计）。失败时**保留**，
 * 那是排查的唯一线索，并且报错文案里给出了路径。
 */
function cleanupLogs(paths) {
    for (const path of paths) {
        try {
            rmSync(path, { force: true });
        } catch {
            // 删不掉不影响结论
        }
    }
}

// ── ffprobe：媒体信息 ────────────────────────────────────────────────────────

/** 跑一次 ffprobe，拿回 JSON（失败时报错带 stderr 尾部）。 */
async function ffprobeJson(absolute, tool, timeoutMs, args) {
    const outPath = tempPath('probe.json');
    const errPath = tempPath('probe.err');
    const result = await runProcess(tool, args, { outPath, errPath, timeoutMs });
    if (result.code !== 0 || result.timedOut === true) {
        const detail = `${result.err ?? ''}`.trim().split('\n').slice(-3).join(' ').slice(0, 400);
        throw avError('OFFICE_AV_PROBE_FAILED', `ffprobe 读不出这份文件的媒体信息${result.timedOut ? '（超时）' : ''}：${detail || result.error || '没有输出'}`);
    }
    try {
        return JSON.parse(result.out ?? '{}');
    } catch (error) {
        throw avError('OFFICE_AV_PROBE_FAILED', `ffprobe 的输出不是 JSON：${error.message}`);
    }
}

/** 把 ffprobe 的原始结构收敛成反馈里要用的那几个量。 */
function summarizeMedia(probe, absolute, relative, bytes) {
    const streams = Array.isArray(probe?.streams) ? probe.streams : [];
    const format = probe?.format ?? {};
    const video = streams.filter((s) => s.codec_type === 'video' && s.disposition?.attached_pic !== 1).map((s) => ({
        index: s.index,
        codec: s.codec_name,
        width: s.width,
        height: s.height,
        fps: parseFps(s.avg_frame_rate ?? s.r_frame_rate),
        frames: Number.isFinite(Number(s.nb_frames)) ? Number(s.nb_frames) : null,
        durationSeconds: Number.isFinite(Number(s.duration)) ? Number(s.duration) : null,
    }));
    const audio = streams.filter((s) => s.codec_type === 'audio').map((s) => ({
        index: s.index,
        codec: s.codec_name,
        sampleRate: Number.isFinite(Number(s.sample_rate)) ? Number(s.sample_rate) : null,
        channels: Number.isFinite(Number(s.channels)) ? Number(s.channels) : null,
        durationSeconds: Number.isFinite(Number(s.duration)) ? Number(s.duration) : null,
    }));
    const durationSeconds = Number.isFinite(Number(format.duration)) && Number(format.duration) > 0
        ? Number(format.duration)
        : (video[0]?.durationSeconds ?? audio[0]?.durationSeconds ?? null);
    const kind = video.length > 0 ? 'video' : (audio.length > 0 ? 'audio' : 'other');
    const extension = extname(absolute).toLowerCase();
    return {
        path: relative,
        bytes,
        kind,
        knownExtension: AV_VIDEO_EXTENSIONS.has(extension) ? 'video' : (AV_AUDIO_EXTENSIONS.has(extension) ? 'audio' : 'other'),
        container: format.format_name ?? '',
        durationSeconds,
        durationClock: durationSeconds === null ? null : formatClock(durationSeconds),
        bitRate: Number.isFinite(Number(format.bit_rate)) ? Number(format.bit_rate) : null,
        hasAudio: audio.length > 0,
        hasVideo: video.length > 0,
        audio,
        video,
    };
}

/** `30000/1001` 这类分数 → 小数；认不出来给 null。 */
function parseFps(value) {
    if (typeof value !== 'string' || value === '') return null;
    const [top, bottom] = value.split('/').map((part) => Number.parseFloat(part));
    if (!Number.isFinite(top)) return null;
    if (!Number.isFinite(bottom) || bottom === 0) return Number.isFinite(top) ? top : null;
    const fps = top / bottom;
    return fps > 0 ? Math.round(fps * 1000) / 1000 : null;
}

/**
 * office.av.info(path)：这份媒体里有什么。
 *
 * @returns {Promise<object>} kind / durationSeconds / hasAudio / hasVideo / 轨道明细
 */
export async function avInfo(filePath, options, env, cache, config) {
    const settings = avSettings(config);
    const probe = await requireCapability(config);
    const input = resolveInput(filePath, env);
    const raw = await ffprobeJson(input.absolute, probe.ffprobe.path, settings.timeoutMs, [
        '-v', 'error', '-print_format', 'json', '-show_format', '-show_streams', input.absolute,
    ]);
    const summary = summarizeMedia(raw, input.absolute, input.relative, input.bytes);
    return {
        ...summary,
        hint: summary.hasAudio
            ? (summary.kind === 'video'
                ? '要文字就 office.av.transcribe(path)，要画面就 office.av.frames(path)，两样一起 office.av.extract(path)。'
                : '要文字就 office.av.transcribe(path)（可传 language / chunkSeconds / overlapSeconds）。')
            : '这份文件里没有音频轨，转写不了；有画面就 office.av.frames(path) 抽帧。',
    };
}

// ── 解码：任意格式 → 规范 16 kHz 单声道 PCM16 WAV ───────────────────────────

/**
 * 用 ffmpeg 解码成规范 WAV，并按内容键缓存在 `.office/cache/av/audio/` 下。
 *
 * 两个关键点：
 *   1. `-map_metadata -1 -fflags +bitexact` 才会省掉 LIST/INFO 块，`data` 才落在
 *      偏移 36。少了这两个开关，ffmpeg 的默认封装就会被后面的规范检查拒收。
 *   2. 解码完**自己再验一遍**（只信字节不信命令行），不合格就地规范化；
 *      规范化失败才报错 —— 报错里带 ffmpeg 的 stderr 尾部。
 */
async function decodeToCanonicalWav(input, probe, settings, cache, logBase) {
    const key = contentKey(input.bytes, input.modifiedMs, `pcm16-16k-mono:${settings.timeoutMs}`);
    const cached = cache !== undefined && cache !== null && typeof cache.path === 'function'
        ? cache.path(`${settings.audioDir}/${slugOf(input.relative)}-${key}.wav`)
        : tempPath(`${slugOf(input.relative)}-${key}.wav`);
    if (existsSync(cached) && isCanonicalWav(cached)) {
        const info = statSync(cached);
        if (cache !== undefined && cache !== null && typeof cache.noteHit === 'function') cache.noteHit(info.size);
        return { path: cached, bytes: info.size, seconds: (info.size - AV_WAV_HEADER_BYTES) / 32_000, reused: true };
    }
    const outPath = `${logBase}.ffmpeg.out.txt`;
    const errPath = `${logBase}.ffmpeg.err.txt`;
    const result = await runProcess(probe.ffmpeg.path, [
        '-hide_banner', '-nostdin', '-v', 'error', '-y',
        '-i', input.absolute,
        '-map', '0:a:0', '-vn',
        '-ac', '1', '-ar', '16000', '-c:a', 'pcm_s16le',
        '-map_metadata', '-1', '-fflags', '+bitexact',
        '-f', 'wav', cached,
    ], { outPath, errPath, timeoutMs: settings.timeoutMs });
    if (result.code !== 0 || result.timedOut === true) {
        const detail = `${result.err ?? ''}`.trim().split('\n').slice(-4).join(' ').slice(0, 500);
        throw avError('OFFICE_AV_DECODE_FAILED', `ffmpeg 解码失败${result.timedOut ? '（超时）' : ''}：${detail || result.error || '没有输出'}`);
    }
    cleanupLogs([outPath, errPath]);
    let bytes = statSync(cached).size;
    let seconds = (bytes - AV_WAV_HEADER_BYTES) / 32_000;
    let normalized = false;
    if (!isCanonicalWav(cached)) {
        const buffer = canonicalizeWav(readFileSync(cached));
        writeFileSync(cached, buffer);
        bytes = buffer.length;
        seconds = (bytes - AV_WAV_HEADER_BYTES) / 32_000;
        normalized = true;
    }
    if (cache !== undefined && cache !== null && typeof cache.noteArtifact === 'function') cache.noteArtifact();
    return { path: cached, bytes, seconds, reused: false, normalized, logFiles: { stdout: outPath, stderr: errPath } };
}

// ── 分块 ─────────────────────────────────────────────────────────────────────

/**
 * 把总时长切成若干块（最后一块可能更短）。
 * 块是给推理端用的：一次请求的音频越长，内存与单次超时越难控。
 *
 * 第三十一轮起每块多两个量，**别把 start/end 当成「这块的全部音频」**：
 *
 *   - `ownedSeconds` 这块**自己负责出稿**的区间，是 `[start, ownedEnd)`；
 *   - `end` 是这块真正喂给 VAD 的窗口末尾 —— 比 ownedEnd 多出 `overlapSeconds`
 *     的重叠（最后一块没有下一块，重叠自然截到总长）。
 *
 * 重叠的作用是让跨在 `ownedEnd` 上的短音在同一块里被完整解出来；跨得更深的
 * 长句由 `resolveChunkBoundary` 的边界回退兜住。两者合起来才是不丢字、不重复的
 * 完整契约，只用重叠是不够的（见那个函数的注释）。
 *
 * @param {number} seconds 音频总秒数
 * @param {number} chunkSeconds 单块秒数（own 区间长度）
 * @param {number} overlapSeconds 相邻块的重叠秒数
 */
export function planChunks(seconds, chunkSeconds, overlapSeconds = AV_CHUNK_OVERLAP_SECONDS) {
    const total = Math.max(0, Number(seconds) || 0);
    const size = Math.max(1, Number(chunkSeconds) || AV_CHUNK_SECONDS);
    const overlap = Math.max(0, Number(overlapSeconds) || 0);
    const round = (value) => Math.round(value * 1000) / 1000;
    if (total === 0) return [{ index: 0, start: 0, end: 0, ownedEnd: 0, seconds: 0, ownedSeconds: 0 }];
    const chunks = [];
    for (let start = 0; start < total; start += size) {
        const ownedEnd = Math.min(total, start + size);
        const end = Math.min(total, ownedEnd + overlap);
        chunks.push({
            index: chunks.length,
            start: round(start),
            end: round(end),
            ownedEnd: round(ownedEnd),
            seconds: round(end - start),
            ownedSeconds: round(ownedEnd - start),
        });
    }
    return chunks;
}

/**
 * 一块窗口解出来的句子，哪些留在本块、下一块从哪儿接着跑。
 *
 * ## 为什么光有重叠不够
 *
 * 重叠只保证「结束点落在重叠区里」的句子能在同一块里解完。一句话如果开始得早、
 * 结束得比窗口末尾还晚，它在两块里**都是断的**：本块截到窗口末尾，下一块从窗口
 * 起点开始（话头已经被上一块吃掉了），逐字稿里就会出现一小段重复、且这句被切成
 * 两段。做了重叠之后这个问题反而更容易暴露 —— 旧代码（完全没有重叠）只是把它
 * 切成两句，不会重复。
 *
 * ## 契约
 *
 * 输入是「本块窗口解出来的全部句子」（绝对时间），输出是两条：
 *
 *   - `keep`：留在本块的句子下标 —— 起点落在自有区间 `[base, ownedEnd)` 里、
 *     没有被窗口末尾切断的句子。起点落在重叠区（`>= ownedEnd`）的句子属于
 *     下一块，本块不留：下一块的窗口从 `nextBase` 起一定覆盖它。
 *   - `nextBase`：下一块的起点。默认是 `ownedEnd`；只要有一条句子要被下一块
 *     重新解（起点在重叠区里，或者被窗口末尾切断），就回退到它的起点，
 *     让它连同话头一起在下一块里被完整解一遍。
 *
 * 唯一的例外是「回退无路可走」：被切断的那句话起点就是窗口起点（回退等于原地
 * 打转）时照常出稿、截断在窗口末尾 —— 宁可少几个字，也不能为了回退把整句丢掉，
 * 更不能让循环挂住。
 *
 * 回退为什么是安全的（这是整个分块方案的正确性依据）：
 *   1. VAD 的句子在时间上不重叠且有先后顺序，被回退的那些句子一定排在所有
 *      `keep` 的句子之后 —— 所以回退不会把已经出稿的音频再解一遍（不重复）。
 *   2. 回退点是那条句子自己的起点 —— 所以它不会丢话头（不丢字）。
 *   3. 回退点在 `(base, ownedEnd]` 里，窗口长度又是 `chunk + overlap`，所以
 *      「一句话能被某个窗口完整装下」的条件只要求这句话不长于单块秒数；
 *      而 VAD 的 `maxSpeechDuration` 在 worker 里被夹到不超过单块秒数，
 *      这条前提是成立的（同时也是这个 while 循环一定前进的依据）。
 *
 * ## 水位线：`emittedUntil`
 *
 * 光有重叠会**重复出稿**：上一块的窗口尾巴把一句完整的话解了出来（这是重叠
 * 想要的效果），下一块的窗口从更早的地方开始，又看见了这句话的尾巴 ——
 * 只按「起点落在自有区间里」判断的话，这条尾巴会在下一块再报一次。
 * 所以还要记住**已经出稿到哪一秒**（worker 每收一条就推进一次），并据此处理
 * 本窗口解出来的句子：
 *
 *   - 起点在水位线之前、整条都在水位线之前：上一块已经报过，丢掉；
 *   - 起点在水位线之前、跨在水位线上（前半截报过）：本块不出稿，把下一块的
 *     起点放到水位线上，让后半截在下一块里从水位线接着解 —— 不重复，也不丢
 *     后半截；
 *   - 起点在水位线之后：照常按上面的契约判断。
 *
 * 判「起点在水位线之前」用的是**窗口起点被夹住**（`start === base`）而不是
 * 「离水位线够远」：VAD 按 512 样本（≈32 ms）切帧，上一块留下的尾巴可能只有
 * 几毫秒，按距离判会漏掉它、让这半句在下一块里再报一次。
 *
 * @param {object} input
 *   base          本块窗口起点（秒）
 *   ownedEnd      本块自有区间末尾（秒）
 *   winEnd        本块窗口末尾（秒）
 *   total         音频总时长（秒）—— 窗口末尾正好等于总长时，末尾那句是自然结束，不能回退
 *   raw           本块窗口解出来的句子 `[{start, end}]`（绝对秒，按时间升序）
 *   emittedUntil  已经出稿到的时刻（秒），默认 0
 * @returns {{keep: number[], nextBase: number, rolledBack: boolean}}
 */
export function resolveChunkBoundary({ base, ownedEnd, winEnd, total, raw = [], emittedUntil = 0 }) {
    const epsilon = 0.05;
    const waterlineEpsilon = 0.005;
    const keep = [];
    let nextBase = ownedEnd;
    let rolledBack = false;
    for (let index = 0; index < raw.length; index += 1) {
        const segment = raw[index];
        // 起点被窗口起点夹住（`start === base`）说明这句话**可能在窗口之前就开始了**；
        // 这时只要窗口起点落在水位线之前，它就是上一块报过的那句的尾巴 —— 逐字稿里
        // 不能再来一份。VAD 的句子是按 512 样本（≈32 ms）切帧的，尾巴可能只有几毫秒，
        // 所以判据是「夹在窗口起点」而不是「离水位线够远」。
        const clampedAtBase = segment.start <= base + 1e-9;
        const behindWaterline = (clampedAtBase && base < emittedUntil - waterlineEpsilon)
            || segment.start < emittedUntil - waterlineEpsilon;
        if (behindWaterline) {
            if (segment.end > emittedUntil + waterlineEpsilon) {
                // 前半截报过、后半截还没：把下一块挪到水位线上接着解后半截。
                // 必须就此打住：下一块会从水位线（比后面几句的起点更早）重跑，
                // 后面几句跟着一起交给它，水位线才能一直是「这之前的都出过稿」。
                nextBase = Math.max(base, emittedUntil);
                break;
            }
            // 整条都在水位线之前：上一块已经报过，丢掉；后面的句子照常处理。
            continue;
        }
        // 窗口末尾被切断（且后面还有音频）→ 这句不能在本块出稿，交给下一块从它的
        // 起点重新解。0.05 秒的余量是给 VAD 的帧长（512 样本 ≈ 32 ms）留的：
        // 一句在窗口末尾自然结束、又恰好落在余量里时最多多解一遍，不会丢字。
        const cutAtEnd = winEnd < total - 1e-9 && segment.end >= winEnd - epsilon;
        if (cutAtEnd && segment.start > base + 1e-9) {
            nextBase = Math.max(base, segment.start);
            rolledBack = true;
            break;
        }
        // 起点落在重叠区：这一句属于下一块，本块不留（下一块的窗口从 nextBase 起
        // 一定覆盖它）。
        if (segment.start >= ownedEnd - 1e-9) {
            nextBase = Math.max(base, segment.start);
            break;
        }
        // 剩下的两种都留在本块：
        //   - 完整的句子（起点在自有区间里、没被窗口末尾切断）；
        //   - 被窗口末尾切断、但**回退无路可走**的句子（它的起点就是窗口起点，
        //     回退等于原地打转）。这种情况只在一句话长到横跨整个窗口时出现
        //     （VAD 的单句上限已被夹到不超过单块秒数，所以理论上不该发生），
        //     宁可照出稿、让它截断在窗口末尾，也不能为了回退把整句丢掉。
        keep.push(index);
    }
    // 兜底：回退点至少要往前走一个样本级别，否则同一块会被反复解（理论上前提
    // 不成立时才会走到这里，但那意味着死循环，宁可少一次回退也不能挂住）。
    if (!(nextBase > base)) nextBase = ownedEnd > base ? ownedEnd : base + 1;
    return { keep, nextBase, rolledBack };
}

/** 转写结果里每段的文本拼接规则：中日韩之间不加空格，拉丁语之间加一个空格。 */
export function joinTranscript(parts) {
    const items = parts.filter((part) => typeof part === 'string' && part.trim() !== '');
    let out = '';
    for (const item of items) {
        const text = item.trim();
        if (out === '') {
            out = text;
            continue;
        }
        const previous = out.slice(-1);
        const next = text.slice(0, 1);
        const cjk = /[\u3000-\u9fff\uf900-\ufaff\uff00-\uffef]/;
        out += cjk.test(previous) || cjk.test(next) ? text : ` ${text}`;
    }
    return out;
}

// ── 转写 ─────────────────────────────────────────────────────────────────────

/** 选一个进程内协作者：把结果 JSON 读回来。 */
function readJsonFile(path) {
    try {
        return JSON.parse(readFileSync(path, 'utf8'));
    } catch (error) {
        throw avError('OFFICE_AV_TRANSCRIBE_FAILED', `读不到转写结果（${basename(path)}）：${error.message}`);
    }
}

/**
 * office.av.transcribe(path, options)：把声音转成文字。
 *
 * options: { language?, chunkSeconds?, overlapSeconds?, maxSeconds?, out?, words?, keepAudio? }
 *   - language      语言提示（auto / zh / en / yue / ja / ko），默认取设置
 *   - chunkSeconds  分块秒数（默认 120）
 *   - overlapSeconds 相邻两块的重叠秒数（默认 0.5，上限是半块）—— 块不再是完全切开，
 *                    跨在切点上的短音能在同一块里解完；再深的长句由边界回退兜住
 *   - out           把带时间戳的转写稿写成工作目录里的文件（长稿必须给，否则会被反馈截断）
 *   - words         是否带上词级时间戳（默认不带，体积大）
 *
 * @returns {Promise<object>} 文本、分块与逐句时间戳、耗时、产物路径
 */
export async function avTranscribe(filePath, options, env, cache, config) {
    const settings = avSettings(config);
    const probe = await requireCapability(config);
    requireSpeech(probe, settings);
    const input = resolveInput(filePath, env);
    const info = await avInfo(filePath, options, env, cache, config);
    if (!info.hasAudio) {
        throw avError('OFFICE_AV_NO_AUDIO_STREAM', `${info.path} 里没有音频轨，转写不了。要是只想看画面，用 office.av.frames(path)。`);
    }
    const maxSeconds = clampOption(options?.maxSeconds, settings.maxSeconds, 1, 21_600);
    const duration = Number(info.durationSeconds) || 0;
    if (duration > maxSeconds) {
        throw avError('OFFICE_AV_TOO_LONG', `${info.path} 时长 ${formatClock(duration)} 超过单次上限 ${formatClock(maxSeconds)}。`
            + `先在设置页调大「单次时长上限」，或者用外部工具切成几段再逐段转写（分块只解决内存，不改变这条上限）。`);
    }

    const started = Date.now();
    const logBase = cache !== undefined && cache !== null && typeof cache.path === 'function'
        ? cache.path(`${settings.audioDir}/log-${slugOf(input.relative)}-${Date.now().toString(36)}`)
        : tempPath('transcribe');
    const audio = await decodeToCanonicalWav(input, probe, settings, cache, logBase);
    if (audio.seconds > maxSeconds) {
        throw avError('OFFICE_AV_TOO_LONG', `解码出来的音频时长 ${formatClock(audio.seconds)} 超过单次上限 ${formatClock(maxSeconds)}。`);
    }
    const chunkSeconds = clampOption(options?.chunkSeconds, settings.chunkSeconds, 10, 600);
    const overlapSeconds = Math.min(
        clampOption(options?.overlapSeconds, settings.overlapSeconds, 0, 30),
        chunkSeconds / 2,
    );
    const chunks = planChunks(audio.seconds, chunkSeconds, overlapSeconds);
    const resultPath = cache !== undefined && cache !== null && typeof cache.path === 'function'
        ? cache.path(`${settings.audioDir}/result-${slugOf(input.relative)}-${Date.now().toString(36)}.json`)
        : tempPath('result.json');
    const workerConfig = {
        wav: audio.path,
        outPath: resultPath,
        runtimePath: probe.sensevoice.runtime,
        model: probe.sensevoice.model,
        tokens: probe.sensevoice.tokens,
        vad: probe.sensevoice.vad,
        threads: settings.threads,
        language: AV_LANGUAGES.includes(String(options?.language ?? '')) ? String(options.language) : settings.language,
        chunkSeconds,
        overlapSeconds,
        maxSegmentSeconds: settings.maxSegmentSeconds,
        vadThreshold: settings.vadThreshold,
        minSpeechSeconds: settings.minSpeechSeconds,
        minSilenceSeconds: settings.minSilenceSeconds,
        words: options?.words === true || settings.words,
    };
    const stdoutPath = `${logBase}.worker.out.txt`;
    const stderrPath = `${logBase}.worker.err.txt`;
    const worker = await runProcess(process.execPath, [AV_WORKER_PATH, JSON.stringify(workerConfig)], {
        outPath: stdoutPath,
        errPath: stderrPath,
        timeoutMs: Math.min(1_800_000, Math.max(30_000, settings.timeoutMs + chunks.length * 5_000)),
        cwd: cache !== undefined && cache !== null && typeof cache.dir === 'string' ? cache.dir : tmpdir(),
    });
    const payload = existsSync(resultPath) ? readJsonFile(resultPath) : undefined;
    if (worker.code !== 0 || worker.timedOut === true || payload?.ok !== true) {
        const detail = `${payload?.error ?? ''}`.trim()
            || `${worker.err ?? ''}`.trim().split('\n').slice(-5).join(' ').slice(0, 500)
            || worker.error
            || '子进程没有输出';
        throw avError('OFFICE_AV_TRANSCRIBE_FAILED', `SenseVoice 转写失败${worker.timedOut ? '（超时）' : ''}：${detail}`
            + `（日志：${displayPath(input.root, stderrPath)}）`);
    }

    const segments = Array.isArray(payload.segments) ? payload.segments : [];
    const text = typeof payload.text === 'string' ? payload.text : joinTranscript(segments.map((item) => item.text));
    // 词级时间是否真的产出了：请求了 words 但模型侧没给出 tokens/timestamps 时，
    // 段里不会有 words 字段 —— 这一行让「开了但没算出」不用把整段 words 读一遍就能看出来。
    const wordSegments = settings.words === true || options?.words === true
        ? segments.filter((item) => Array.isArray(item.words) && item.words.length > 0)
        : [];
    cleanupLogs([stdoutPath, stderrPath]);
    const value = {
        ok: true,
        path: input.relative,
        kind: info.kind,
        model: payload.model ?? `SenseVoiceSmall (${settings.precision.toUpperCase()})`,
        language: payload.language ?? workerConfig.language,
        detectedLanguages: payload.detectedLanguages ?? [],
        durationSeconds: info.durationSeconds,
        audioSeconds: Math.round(audio.seconds * 1000) / 1000,
        speechSeconds: payload.speechSeconds ?? null,
        inferenceSeconds: payload.inferenceSeconds ?? null,
        realtimeFactor: payload.inferenceSeconds !== null && audio.seconds > 0
            ? Math.round((payload.inferenceSeconds / audio.seconds) * 1000) / 1000
            : null,
        chunkSeconds,
        overlapSeconds,
        chunks: payload.chunks ?? chunks.length,
        // 边界回退的次数：0 表示每块都正好在句子之间切开（最省），大于 0 表示有
        // 句子跨在切点上、被下一块从话头重新解过。它是「重叠够不够」的实测信号。
        rolls: payload.rolls ?? 0,
        segmentCount: segments.length,
        ...(settings.words === true || options?.words === true
            ? {
                wordTimings: {
                    segments: wordSegments.length,
                    count: wordSegments.reduce((sum, item) => sum + item.words.length, 0),
                    unit: '中日韩按字 / 西文按词',
                },
            }
            : {}),
        segments: segments.map((item) => ({
            index: item.index,
            start: item.start,
            end: item.end,
            clock: item.clock ?? `${formatClock(item.start)}-${formatClock(item.end)}`,
            text: item.text,
            lang: item.lang ?? null,
            emotion: item.emotion ?? null,
            event: item.event ?? null,
            ...(item.words === undefined ? {} : { words: item.words }),
        })),
        text,
        audio: {
            path: displayPath(input.root, audio.path),
            seconds: Math.round(audio.seconds * 1000) / 1000,
            reused: audio.reused === true,
            normalized: audio.normalized === true,
            bytes: audio.bytes,
        },
        logFiles: undefined,
        elapsedMs: Date.now() - started,
    };
    // 成功路径不留进程日志（空文件只会在缓存里变成噪声）；失败时它们才是有用的线索，
    // 那条路径在抛错前会把 stderr 的路径写进报错文案。
    delete value.logFiles;

    if (typeof options?.out === 'string' && options.out.trim() !== '') {
        const markdown = renderTranscript(value);
        const written = env.writeFile(options.out, markdown);
        value.out = { path: written.path, chars: [...markdown].length, bytes: written.bytes };
    }
    return value;
}

/** 转写稿里词级时间的封顶（个词）：默认关的字段，写进文件时也要有界。 */
const TRANSCRIPT_WORD_LIMIT = 600;

/** 转写稿的 Markdown 形态：头部元信息 + 逐句时间戳。 */
export function renderTranscript(result) {
    const lines = [
        `# ${basename(String(result.path ?? '媒体'))} 转写稿`,
        '',
        `- 来源：${result.path}`,
        `- 时长：${result.durationSeconds === null || result.durationSeconds === undefined ? '未知' : formatClock(result.durationSeconds)}`,
        `- 语音：${formatClock(result.audioSeconds ?? 0)}（识别出 ${result.segmentCount ?? 0} 句）`,
        `- 语言：${result.language ?? 'auto'}${Array.isArray(result.detectedLanguages) && result.detectedLanguages.length > 0 ? `（判定 ${result.detectedLanguages.join(' / ')}）` : ''}`,
        `- 模型：${result.model ?? 'SenseVoiceSmall'}`,
        `- 耗时：${result.inferenceSeconds ?? '?'} 秒${result.realtimeFactor ? `（实时率 ${result.realtimeFactor}）` : ''}`,
        '',
        '| # | 时间 | 文本 |',
        '| --- | --- | --- |',
    ];
    for (const segment of result.segments ?? []) {
        const text = String(segment.text ?? '').replace(/\|/g, '\\|').trim();
        lines.push(`| ${Number(segment.index ?? 0) + 1} | ${segment.clock ?? ''} | ${text} |`);
    }
    if ((result.segments ?? []).length === 0) lines.push('| — | — | （没有识别出语音） |');
    const extras = (result.segments ?? []).filter((item) => item.emotion && item.emotion !== '<|NEUTRAL|>');
    if (extras.length > 0) {
        lines.push('', '## 非中性情绪 / 非语音事件', '');
        for (const item of extras) lines.push(`- ${item.clock}：${item.emotion} ${item.event ?? ''}`);
    }
    // 词级时间只在真的算出来时才有这一节（words:true 且 tokens/timestamps 可用）。
    // 有界：词级时间默认关就是因为它会长，写进稿子时也要封顶，不然一份长会议能到几百 KB。
    const withWords = (result.segments ?? []).filter((item) => Array.isArray(item.words) && item.words.length > 0);
    if (withWords.length > 0) {
        lines.push('', `## 词级时间（秒；中日韩按字、西文按词；共 ${withWords.length} 句）`, '');
        let printed = 0;
        let omittedWords = 0;
        let omittedSegments = 0;
        for (const item of withWords) {
            const room = TRANSCRIPT_WORD_LIMIT - printed;
            if (room <= 0) {
                // 硬上限：跨过上限的那一句只印**装得下**的那几个词，
                // 而不是整句印完再停（那会让「最多 600 个词」这句话不成立）。
                omittedSegments += 1;
                omittedWords += item.words.length;
                continue;
            }
            const shown = item.words.slice(0, room);
            printed += shown.length;
            if (shown.length < item.words.length) {
                omittedWords += item.words.length - shown.length;
                omittedSegments += 1;
            }
            lines.push(`- ${item.clock ?? ''}：${shown.map((word) => `${Number(word.start ?? 0).toFixed(2)} ${String(word.text ?? '')}`).join(' / ')}`);
        }
        if (omittedWords > 0) {
            lines.push(`- （另有 ${omittedWords} 个词（${omittedSegments} 句）被省略：单份转写稿最多列 ${TRANSCRIPT_WORD_LIMIT} 个词）`);
        }
    }
    return `${lines.join('\n')}\n`;
}

// ── 抽帧 ─────────────────────────────────────────────────────────────────────

/**
 * 决定抽哪几个时间点。
 *
 * 优先级：at（明确给点）> every（定步长）> count（默认 6 张均匀铺开）。
 * from / to 可以限制区间。结果按时间排序去重。
 */
export function planFrameTimes(durationSeconds, options = {}, defaults = {}) {
    const duration = Math.max(0, Number(durationSeconds) || 0);
    const from = Math.max(0, Number(options.from) || 0);
    const to = Math.min(duration > 0 ? duration : Number.POSITIVE_INFINITY, Number.isFinite(Number(options.to)) ? Number(options.to) : (duration > 0 ? duration : 0));
    const upper = duration > 0 ? Math.max(from, to) : Math.max(from, to || from);
    if (Array.isArray(options.at) || Number.isFinite(Number(options.at))) {
        const raw = Array.isArray(options.at) ? options.at : [options.at];
        const times = raw.map((item) => Number(item)).filter((item) => Number.isFinite(item) && item >= 0);
        if (times.length === 0) throw avError('OFFICE_AV_BAD_OPTION', 'frames 的 at 里没有有效的时间点（秒）。');
        return [...new Set(times.map((t) => Math.round(t * 1000) / 1000))].sort((a, b) => a - b);
    }
    if (Number.isFinite(Number(options.every)) && Number(options.every) > 0) {
        const step = Number(options.every);
        const times = [];
        for (let t = from; t <= upper && times.length < 1000; t += step) times.push(Math.round(t * 1000) / 1000);
        if (times.length === 0) times.push(from);
        return times;
    }
    const count = Math.max(1, Math.min(Number(options.count) || defaults.count || AV_DEFAULT_FRAMES, defaults.max ?? 60));
    if (duration <= 0) return [0];
    const span = Math.max(0, (Number.isFinite(upper) ? upper : duration) - from);
    if (count === 1) return [Math.round((from + span / 2) * 1000) / 1000];
    const times = [];
    for (let i = 0; i < count; i += 1) times.push(Math.round((from + (span * (i + 0.5)) / count) * 1000) / 1000);
    return times;
}

/**
 * office.av.frames(path, options)：按时间点抽帧成图片。
 *
 * options: { at?, every?, count?, from?, to?, format?, width?, outDir? }
 * @returns {Promise<object>} 每张图的路径与时间点；路径直接交给 read_image
 */
export async function avFrames(filePath, options, env, cache, config) {
    const settings = avSettings(config);
    const probe = await requireCapability(config);
    const input = resolveInput(filePath, env);
    const info = await avInfo(filePath, options, env, cache, config);
    if (!info.hasVideo) {
        throw avError('OFFICE_AV_NO_VIDEO_STREAM', `${info.path} 里没有视频轨，抽不了帧。要文字用 office.av.transcribe(path)。`);
    }
    const times = planFrameTimes(info.durationSeconds, options ?? {}, { count: settings.frames, max: settings.maxFrames });
    if (times.length > settings.maxFrames) {
        throw avError('OFFICE_AV_BAD_OPTION', `一次最多抽 ${settings.maxFrames} 张（这次算了 ${times.length} 张）。调小 count / 放大 every，或分批抽。`);
    }
    const format = options?.format === 'png' ? 'png' : (options?.format === 'jpg' || options?.format === undefined ? settings.frameFormat : String(options.format));
    if (format !== 'png' && format !== 'jpg') throw avError('OFFICE_AV_BAD_OPTION', 'frames 的 format 只支持 jpg / png。');
    const width = clampOption(options?.width, settings.frameWidth, 0, 4096);
    const dirName = typeof options?.outDir === 'string' && options.outDir.trim() !== ''
        ? options.outDir.trim()
        : `${settings.framesDir}/${slugOf(input.relative)}`;
    const dir = cache !== undefined && cache !== null && typeof cache.ensureDir === 'function'
        ? cache.ensureDir(dirName)
        : join(tmpdir(), `office-av-frames-${process.pid}`);
    const started = Date.now();
    const files = [];
    for (const [index, at] of times.entries()) {
        const name = `frame-${String(index + 1).padStart(3, '0')}-${formatClock(at).replace(/:/g, 'm')}s.${format}`;
        const target = join(dir, name);
        const args = ['-hide_banner', '-nostdin', '-v', 'error', '-y', '-ss', String(at), '-i', input.absolute, '-frames:v', '1', '-q:v', '3'];
        if (width > 0) args.push('-vf', `scale=${width}:-2`);
        args.push(target);
        const outPath = `${target}.ffmpeg.out.txt`;
        const errPath = `${target}.ffmpeg.err.txt`;
        const result = await runProcess(probe.ffmpeg.path, args, { outPath, errPath, timeoutMs: settings.timeoutMs });
        if (result.code !== 0 || result.timedOut === true || !existsSync(target)) {
            const detail = `${result.err ?? ''}`.trim().split('\n').slice(-3).join(' ').slice(0, 300);
            throw avError('OFFICE_AV_FRAMES_FAILED', `抽第 ${index + 1} 帧（${formatClock(at)}）失败：${detail || result.error || '没有输出'}`
                + `（ffmpeg 日志：${displayPath(input.root, errPath)}）`);
        }
        cleanupLogs([outPath, errPath]);
        const bytes = statSync(target).size;
        if (cache !== undefined && cache !== null && typeof cache.noteArtifact === 'function') cache.noteArtifact();
        files.push({
            index: index + 1,
            at: Math.round(at * 1000) / 1000,
            clock: formatClock(at),
            path: displayPath(input.root, target),
            bytes,
        });
    }
    return {
        ok: true,
        path: input.relative,
        kind: info.kind,
        durationSeconds: info.durationSeconds,
        dir: displayPath(input.root, dir),
        format,
        width: width > 0 ? width : null,
        count: files.length,
        files,
        hint: '逐张 read_image({ file_path }) 就能看到画面；图上若有文字，那是读图能力认出来的，不是 OCR。',
        elapsedMs: Date.now() - started,
    };
}

/** 把选项收敛到区间里；没给就用默认。 */
function clampOption(value, fallback, min, max) {
    const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

// ── 一次做完 ─────────────────────────────────────────────────────────────────

/**
 * office.av.extract(path, options)：信息 + 转写（有音频时）+ 抽帧（有视频时）。
 *
 * 这是「一次调用干完一批」的那个入口：会议视频一条调用就能拿到逐字稿与
 * 若干画面。options 同时接受 transcribe / frames 两边的参数，另有：
 *   - out            把逐字稿写成工作目录里的 Markdown
 *   - frames         false 时不抽帧；数字时当作抽几张
 *   - transcript     false 时不转写（只抽帧）
 */
export async function avExtract(filePath, options, env, cache, config) {
    const settings = avSettings(config);
    const opts = options !== null && typeof options === 'object' ? options : {};
    const info = await avInfo(filePath, opts, env, cache, config);
    const wantTranscript = opts.transcript !== false && info.hasAudio;
    const framesWanted = opts.frames !== false && info.hasVideo;
    const frameOptions = (() => {
        const { frames, transcript, out, ...rest } = opts;
        if (typeof frames === 'number') return { ...rest, count: frames };
        return rest;
    })();
    const result = {
        ok: true,
        path: info.path,
        kind: info.kind,
        info,
        frames: null,
        transcript: null,
        notes: [],
    };
    if (wantTranscript) {
        result.transcript = await avTranscribe(filePath, { ...opts, out: opts.out }, env, cache, config);
    } else if (!info.hasAudio) {
        result.notes.push('没有音频轨，跳过转写。');
    } else {
        result.notes.push('按参数要求跳过转写。');
    }
    if (framesWanted) {
        result.frames = await avFrames(filePath, frameOptions, env, cache, config);
    } else if (!info.hasVideo) {
        result.notes.push('没有视频轨，跳过抽帧。');
    } else {
        result.notes.push('按参数要求跳过抽帧。');
    }
    if (result.transcript !== null && result.frames !== null) {
        result.next = '逐字稿在 transcript（或 out 指的文件）里，画面用 frames.files 逐张 read_image。';
    } else if (result.transcript !== null) {
        result.next = '逐字稿在 transcript（长稿建议给 out 落成文件再读）。';
    } else if (result.frames !== null) {
        result.next = '画面用 frames.files 逐张 read_image。';
    }
    return result;
}

/** office.av.check()：本机能不能做音频 / 视频的内容提取。 */
export async function avCheck(options = {}) {
    const settings = avSettings(options.config);
    const probe = await probeAv(options.config);
    return {
        available: probe.available,
        enabled: probe.enabled,
        missing: probe.missing,
        ffmpeg: probe.ffmpeg,
        ffprobe: probe.ffprobe,
        sensevoice: {
            modelDirectory: probe.sensevoice.modelDirectory,
            model: probe.sensevoice.model === null ? null : displayPath(process.cwd(), probe.sensevoice.model),
            tokens: probe.sensevoice.tokens,
            vad: probe.sensevoice.vad,
            runtime: probe.sensevoice.runtime,
            precision: probe.sensevoice.precision,
            threads: probe.sensevoice.threads,
            languages: probe.sensevoice.languages,
        },
        language: settings.language,
        chunkSeconds: settings.chunkSeconds,
        overlapSeconds: settings.overlapSeconds,
        maxSeconds: settings.maxSeconds,
        maxFrames: settings.maxFrames,
        audioDir: settings.audioDir,
        framesDir: settings.framesDir,
        hint: probe.hint,
    };
}
