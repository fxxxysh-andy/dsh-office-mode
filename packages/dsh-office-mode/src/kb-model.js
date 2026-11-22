/**
 * 本地嵌入模型的下载器（第八十八轮；第一百零三轮换成小模型）。
 *
 * 只做一件事：把本地 embedding 模型从**国内可直连的来源**（魔搭 ModelScope）
 * 拉到知识库目录下的 `models/<模型名>/`，并给面板一份可轮询的状态。魔搭失败时
 * 再退回 Hugging Face 与 hf-mirror（老路径不删，境外用户仍然能用）。
 *
 * 第一百零三轮把默认模型从 Qwen3-Embedding-0.6B 换成 BAAI/bge-small-zh-v1.5：
 * 0.6B 即使 q4 量化也要 350–400MB 权重，加运行时必然超过 600MB 预算；bge-small-zh
 * 的 int8/fp32 在 2 线程、maxlen 256 下实测峰值 RSS 约 480–515MB（见
 * kb-local-embed.js 与第 103 轮文档）。
 *
 * **不在这里加载模型**——推理走本机 Python 服务（kb-local-embed-server.py）；
 * 下载器只负责把文件放对位置。
 */
import { createWriteStream } from 'node:fs';
import { mkdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';

/** 面板与端点共用的模型清单：kind 是稳定 id，repo 是模型仓库（两个来源同 id）。 */
export const MODEL_SPECS = [
    // kind 仍是 'embedding'：它是面板与下载端点的稳定键，换的只是模型本身。
    // label 直接决定下载目录 models/<label>/，与 kb-local-embed 的默认目录保持一致。
    { kind: 'embedding', label: 'bge-small-zh-v1.5', repo: 'BAAI/bge-small-zh-v1.5' },
];

const MANIFEST = '.dsh-office-model.json';
const JOBS = new Map();

const MODELSCOPE_ENDPOINT = 'https://modelscope.cn';

/**
 * Hugging Face 端点：保留原行为（HF_ENDPOINT 只指定一个；否则直连 + 镜像兜底）。
 * 它现在排在魔搭之后，是**第二选择**，不再是默认第一跳。
 */
const DEFAULT_MIRRORS = ['https://hf-mirror.com'];
const HF_ENDPOINT = String(process.env.HF_ENDPOINT || '').replace(/\/$/, '');
const HF_ENDPOINTS = (HF_ENDPOINT !== ''
    ? [HF_ENDPOINT]
    : ['https://huggingface.co', ...String(process.env.HF_MIRRORS || DEFAULT_MIRRORS.join(','))
        .split(',').map((item) => item.trim().replace(/\/$/, '')).filter(Boolean)])
    .filter((item, index, all) => all.indexOf(item) === index);

/**
 * 来源顺序：魔搭在前（用户实测 huggingface.co 在本机 20 秒无响应，hf-mirror 也不稳），
 * 魔搭失败再退回 HF。每个来源自带清单接口与文件 URL 形状，下载时按同一顺序回退。
 */
const SOURCES = [
    { id: 'modelscope', label: '魔搭 ModelScope', endpoint: MODELSCOPE_ENDPOINT },
    ...HF_ENDPOINTS.map((endpoint) => ({
        id: 'huggingface',
        label: endpoint.includes('hf-mirror') ? 'HF 镜像' : 'Hugging Face',
        endpoint,
    })),
];

function sourceKey(source) {
    if (source === null || source === undefined) return '';
    return source.id === 'modelscope' ? 'modelscope' : 'hf:' + source.endpoint;
}

/** 单次网络操作的超时：metadata 15s，文件按块空闲 60s（不是整文件总时长）。 */
const META_TIMEOUT_MS = 15_000;
const BODY_IDLE_TIMEOUT_MS = 60_000;

/** 给 fetch 一个可读的失败原因：超时、DNS、连接重置分别说清楚。 */
function networkReason(error) {
    const name = String(error?.name || '');
    if (name === 'TimeoutError' || name === 'AbortError') return '连接超时';
    const code = String(error?.cause?.code || error?.code || '');
    if (code === 'ENOTFOUND') return '域名解析失败';
    if (code === 'ECONNREFUSED') return '连接被拒绝';
    if (code === 'ECONNRESET') return '连接被重置';
    return String(error?.message || error || '网络错误');
}

/** 模型根：知识库目录下的 models/（不另建第二处真源）。 */
export function modelRootOf(kbDir) {
    return join(String(kbDir || ''), 'models');
}

/** 某个模型的目标目录。kind 不合法时抛错，别让路径拼出一个任意目录。 */
export function modelDirOf(kbDir, kind) {
    const spec = MODEL_SPECS.find((item) => item.kind === kind);
    if (spec === undefined) throw new Error(`不认识的模型：${kind}（只支持 embedding）。`);
    return join(modelRootOf(kbDir), spec.label);
}

async function sizeOf(path) {
    try {
        const info = await stat(path);
        return info.isFile() ? info.size : 0;
    } catch {
        return 0;
    }
}

/** 下载完成才算「有」：manifest 是下载器自己写的提交点。 */
async function readMarker(dir) {
    try {
        const parsed = JSON.parse(await readFile(join(dir, MANIFEST), 'utf8'));
        return parsed !== null && typeof parsed === 'object' ? parsed : null;
    } catch {
        return null;
    }
}

async function inspectOne(kbDir, spec) {
    const dir = modelDirOf(kbDir, spec.kind);
    const marker = await readMarker(dir);
    const active = JOBS.get(spec.kind);
    const running = active !== undefined && active.state === 'running';
    const downloaded = marker !== null && marker.complete === true;
    return {
        kind: spec.kind,
        label: spec.label,
        repo: spec.repo,
        dir,
        downloaded,
        files: Number.isFinite(marker?.files) ? marker.files : 0,
        bytes: Number.isFinite(marker?.bytes) ? marker.bytes : 0,
        completedAt: typeof marker?.completedAt === 'string' ? marker.completedAt : '',
        source: running ? String(active.source || '') : String(marker?.source || ''),
        sourceLabel: running ? String(active.sourceLabel || '') : String(marker?.sourceLabel || ''),
        endpoint: running ? String(active.endpoint || '') : String(marker?.endpoint || ''),
        state: running ? 'running' : (active?.state === 'error' ? 'error' : (downloaded ? 'ready' : 'missing')),
        progress: running ? progressOf(active) : (downloaded ? 1 : 0),
        downloadedBytes: running ? active.downloadedBytes : 0,
        totalBytes: running ? active.totalBytes : 0,
        file: running ? active.file : '',
        error: active?.state === 'error' ? active.error : '',
    };
}

function progressOf(job) {
    if (!Number.isFinite(job.totalBytes) || job.totalBytes <= 0) return 0;
    return Math.max(0, Math.min(1, job.downloadedBytes / job.totalBytes));
}

/** 面板轮询的状态：两个模型各一条，外加它们共同的根目录。 */
export async function modelStatus({ kbDir }) {
    const root = modelRootOf(kbDir);
    const models = [];
    for (const spec of MODEL_SPECS) models.push(await inspectOne(kbDir, spec));
    return {
        ok: true,
        root,
        kbDir: String(kbDir || ''),
        sources: SOURCES.map((source) => ({ id: sourceKey(source), label: source.label, endpoint: source.endpoint })),
        models,
    };
}

function safeRel(name) {
    const value = String(name || '').replace(/\\/g, '/');
    if (value === '' || value.startsWith('/') || value.includes('..')) return '';
    return value;
}

function encodePath(rel) {
    return rel.split('/').map((part) => encodeURIComponent(part)).join('/');
}

/**
 * 魔搭的仓库清单：`/api/v1/models/<repo>/repo/files?Revision=master&Recursive=true`。
 * 返回的 Data.Files 同时含 tree（目录）与 blob（文件），目录不下载。
 * 字段名是 Path / Type / Size（与 HF 的 rfilename / size 不同）。
 */
async function listModelScope(repo) {
    const response = await fetch(`${MODELSCOPE_ENDPOINT}/api/v1/models/${repo}/repo/files?Revision=master&Recursive=true`, {
        headers: { 'user-agent': 'dsh-office-mode/model-downloader' },
        signal: AbortSignal.timeout(META_TIMEOUT_MS),
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    const rows = Array.isArray(payload?.Data?.Files) ? payload.Data.Files : [];
    const files = rows
        .filter((row) => String(row?.Type || 'blob').toLowerCase() !== 'tree')
        .map((row) => ({ name: safeRel(row?.Path), size: Number(row?.Size) || 0 }))
        .filter((item) => item.name !== '');
    if (files.length === 0) throw new Error('清单里没有可下载的文件');
    return files;
}

/** Hugging Face 的仓库清单：`/api/models/<repo>?blobs=true`（原路径）。 */
async function listHuggingFace(repo, endpoint) {
    const response = await fetch(`${endpoint}/api/models/${repo}?blobs=true`, {
        headers: { 'user-agent': 'dsh-office-mode/model-downloader' },
        signal: AbortSignal.timeout(META_TIMEOUT_MS),
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    const siblings = Array.isArray(payload?.siblings) ? payload.siblings : [];
    return siblings.map((item) => ({ name: safeRel(item?.rfilename), size: Number(item?.size) || 0 }));
}

/**
 * 按来源顺序读清单：返回第一个成功的来源与它的文件表。
 * 魔搭优先；失败后自动换 Hugging Face / hf-mirror，错误里保留每个来源的原因。
 */
async function fetchModelInfo(repo) {
    const failures = [];
    for (const source of SOURCES) {
        try {
            const files = source.id === 'modelscope'
                ? await listModelScope(repo)
                : await listHuggingFace(repo, source.endpoint);
            return { source, endpoint: source.endpoint, files };
        } catch (error) {
            failures.push(`${source.label}（${source.endpoint}）：${networkReason(error)}`);
        }
    }
    throw new Error(`读取仓库清单失败（已尝试 ${SOURCES.length} 个来源）：${failures.join('；')}。`
        + '请检查网络/代理，或设置 HF_ENDPOINT 指向可用镜像后重试。');
}

/** 文件在某个来源上的下载地址。魔搭用 resolve/master，HF 用 resolve/main。 */
function fileUrlOf(source, repo, name) {
    return source.id === 'modelscope'
        ? `${source.endpoint}/models/${repo}/resolve/master/${encodePath(name)}`
        : `${source.endpoint}/${repo}/resolve/main/${encodePath(name)}`;
}

/**
 * 下载一个文件：每个来源一份全新的 .part；断掉/超时就清掉再换下一个来源。
 *
 * 超时用「空闲」而不是「总时长」：0.6B 的权重有一千多兆，固定总超时会让慢网
 * 永远下不完；只要还能收到字节就继续，连续 60s 没有新字节才算断。
 */
async function downloadFile({ target, job, startBytes, preferred }) {
    await mkdir(dirname(target), { recursive: true });
    const order = [...SOURCES];
    const wanted = sourceKey(preferred);
    const first = order.findIndex((source) => sourceKey(source) === wanted);
    if (first > 0) order.unshift(order.splice(first, 1)[0]);
    const failures = [];
    for (const source of order) {
        if (job.cancel === true) throw new Error('已取消。');
        const url = fileUrlOf(source, job.repo, job.name);
        const temp = target + '.part';
        await rm(temp, { force: true });
        job.downloadedBytes = startBytes;
        let written = 0;
        const controller = new AbortController();
        let idle = null;
        const arm = () => {
            if (idle !== null) clearTimeout(idle);
            idle = setTimeout(() => controller.abort(), BODY_IDLE_TIMEOUT_MS);
        };
        try {
            arm();
            const response = await fetch(url, {
                headers: { 'user-agent': 'dsh-office-mode/model-downloader' },
                signal: controller.signal,
            });
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            if (response.body === null) throw new Error('响应没有正文');
            const stream = Readable.fromWeb(response.body);
            stream.on('data', (chunk) => {
                written += chunk.length;
                job.downloadedBytes = startBytes + written;
                arm();
            });
            await pipeline(stream, createWriteStream(temp));
            if (idle !== null) clearTimeout(idle);
            await rename(temp, target);
            job.source = sourceKey(source);
            job.sourceLabel = source.label;
            job.endpoint = source.endpoint;
            return source;
        } catch (error) {
            if (idle !== null) clearTimeout(idle);
            await rm(temp, { force: true });
            failures.push(`${source.label}（${source.endpoint}）：${networkReason(error)}`);
        }
    }
    throw new Error(`下载 ${job.name} 失败（已尝试 ${order.length} 个来源）：${failures.join('；')}`);
}

async function runDownload(kbDir, spec, job) {
    job.state = 'running';
    job.startedAt = new Date().toISOString();
    job.file = '读取仓库清单…';
    const found = await fetchModelInfo(spec.repo);
    let preferred = found.source;
    job.source = sourceKey(preferred);
    job.sourceLabel = preferred.label;
    job.endpoint = found.endpoint;
    const files = found.files
        .filter((item) => item.name !== '.gitattributes' && !item.name.endsWith('.md'));
    if (files.length === 0) throw new Error('仓库清单是空的：来源没有返回可下载的文件。');
    job.totalBytes = files.reduce((sum, item) => sum + item.size, 0);
    job.files = 0;
    job.fileCount = files.length;
    const dir = modelDirOf(kbDir, spec.kind);
    await mkdir(dir, { recursive: true });
    for (const item of files) {
        if (job.cancel === true) throw new Error('已取消。');
        job.file = item.name;
        job.repo = spec.repo;
        job.name = item.name;
        const target = join(dir, item.name);
        const existing = await sizeOf(target);
        if (item.size > 0 && existing === item.size) {
            job.downloadedBytes += item.size;
            job.files += 1;
            continue;
        }
        const startBytes = job.downloadedBytes;
        const used = await downloadFile({ target, job, startBytes, preferred });
        preferred = used;
        job.files += 1;
        if (item.size === 0) job.downloadedBytes = job.totalBytes > 0 ? job.totalBytes : job.downloadedBytes;
    }
    await writeFile(join(dir, MANIFEST), JSON.stringify({
        version: 1,
        repo: spec.repo,
        label: spec.label,
        source: job.source,
        sourceLabel: job.sourceLabel,
        endpoint: job.endpoint,
        complete: true,
        files: files.length,
        bytes: job.downloadedBytes,
        completedAt: new Date().toISOString(),
    }, null, 2) + '\n', 'utf8');
    job.state = 'ready';
    job.progress = 1;
    job.file = '';
    job.finishedAt = new Date().toISOString();
}

/**
 * 启动一次下载。同一个模型重复点击返回同一份状态，不会并发写同一个目录。
 *
 * 下载是后台任务：端点立刻返回 running，面板轮询 status。进程退出会留下 `.part`，
 * 下次点下载会从零重新拉（各来源都没有稳定的断点续传契约，宁可从零写）。
 */
export function startModelDownload({ kbDir, kind }) {
    const spec = MODEL_SPECS.find((item) => item.kind === kind);
    if (spec === undefined) throw new Error(`不认识的模型：${kind}（只支持 embedding）。`);
    if (String(kbDir || '') === '') throw new Error('没有知识库目录：先在设置里选一个目录。');
    const jobs = JOBS.get(kind);
    if (jobs !== undefined && jobs.state === 'running') return inspectOne(kbDir, spec);
    const job = {
        kind,
        state: 'queued',
        downloadedBytes: 0,
        totalBytes: 0,
        files: 0,
        fileCount: 0,
        file: '',
        source: '',
        sourceLabel: '',
        endpoint: '',
        error: '',
    };
    JOBS.set(kind, job);
    runDownload(String(kbDir), spec, job).catch((error) => {
        job.state = 'error';
        job.error = String(error?.message ?? error);
        job.finishedAt = new Date().toISOString();
    });
    return inspectOne(kbDir, spec);
}

/** 清掉下载器缓存（测试用；真实删除模型请直接删 models/ 目录）。 */
export function resetModelJobs() {
    JOBS.clear();
}

export async function removeModel({ kbDir, kind }) {
    const dir = modelDirOf(kbDir, kind);
    await rm(dir, { recursive: true, force: true });
    JOBS.delete(kind);
    return { ok: true, kind, dir };
}
