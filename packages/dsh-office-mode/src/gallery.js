/**
 * 审美参考图库（不是 OCR）：原图保存、缩略图、审美描述与标签、手工整理、可解释检索。
 *
 * ## 这一层负责什么
 *
 *   originals/<sha256>.<ext>   原图，按内容哈希寻址；同内容天然去重
 *   thumbs/<sha256>.jpg        缩略图（Pillow 生成；不可用时明确回执，不假装有）
 *   manifest.json              真源：事实、审美描述、手工标签/分类/收藏分开存
 *   vectors/text/<key>/<id>.json   文本空间向量（描述 + 标签），key 含模型 + 维度 + 版本
 *   vectors/image/<key>/<id>.json  图文对齐空间向量，key 含通道 + 端点 + 模型 + 维度 + 版本
 *
 * ## 三条不可妥协的纪律
 *
 * 1. **事实与审美分开**：facts 只放哈希、字节、宽高、MIME 与调用方给的来源；
 *    审美描述单独放 aesthetic；用户手工标签/分类/收藏放 manual。不推断作者/版权。
 * 2. **两个向量空间绝不混比**：文本向量与图片向量分目录、分 key、分读取路径；
 *    没有配置多模态 embedding、或探针（文本与图片同维）没过时，图文信号明确跳过，
 *    绝不拿文本向量去比图片向量，也不写占位向量。
 * 3. **默认不联网**：描述图片、探针、图片向量化都要显式 allowUpload:true；
 *    文本向量化只上传描述/标签文本（走已有的 embedding 配置），同样在回执里说明。
 *
 * 检索是可解释的融合：词法、标签精确匹配、文本向量、图文对齐向量四路，
 * 每路单独给分，并报告哪一路被用了/被跳过及原因。没有语义能力时就说没有，
 * 不把本地哈希向量包装成「语义召回」。
 *
 * @module dsh-office-mode/gallery
 */
import { createHash, randomUUID } from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { realpathSync } from 'node:fs';
import { mkdir, open, readFile, readdir, realpath, rename, stat, unlink, writeFile } from 'node:fs/promises';
import { homedir, hostname } from 'node:os';
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { sniffImageBytes } from './image-bytes.js';
import {
    LOCAL_EMBEDDING_MODEL,
    asVector,
    canUseRemoteEmbedding,
    cosineSimilarity,
    embedTexts,
} from './kb-embed.js';
import { kbTokens } from './kb.js';
import {
    GALLERY_PROMPT_VERSION,
    IMAGE_EMBED_VERSION,
    createImageEmbedder,
    describeImage,
    imageEmbeddingReady,
    normalizeGalleryConfig,
    parseAestheticJson,
    probeImageEmbedding,
    visionReady,
} from './gallery-model.js';

/** 图库数据版本：结构变了就升它（manifest 里的 version）。 */
export const GALLERY_VERSION = 2;
export const GALLERY_DIRNAME = 'gallery';
/** 单张图上限：原图保存、缩略图、向量化共用这一条。 */
export const MAX_IMAGE_BYTES = 20 * 1024 * 1024;
/** 解码前的像素上限：挡「小文件大画布」的解压炸弹。 */
export const MAX_PIXELS = 60 * 1000 * 1000;

/** 多模态端点自己的单图/单请求上限（官方文档口径，按十进制 MB）。 */
const IMAGE_EMBED_LIMITS = Object.freeze({
    voyage: { bytes: 20_000_000, pixels: 16_000_000 },
    cohere: { bytes: 5_000_000, requestBytes: 20_000_000 },
});
export const MAX_TAGS = 40;
export const MAX_TAG_CHARS = 48;
export const MAX_NOTE_CHARS = 2000;
export const MAX_NAME_CHARS = 120;
export const MAX_DESCRIPTION_CHARS = 600;
export const MAX_CATEGORY_CHARS = 40;
export const MAX_FOLDER_CHARS = 200;
export const MAX_FOLDER_NAME_CHARS = 40;
export const MAX_FOLDER_DEPTH = 8;
export const AESTHETIC_FACETS = ['subject', 'composition', 'light', 'material', 'mood'];
export const SUMMARY_MAX = 600;
export const PALETTE_MAX = 8;
export const TEXT_VECTOR_VERSION = 'text-v1';
export const DEFAULT_WEIGHTS = { lexical: 0.34, tag: 0.1, textVector: 0.3, imageText: 0.26, imageImage: 0.26 };

/** 跨进程 manifest 锁：锁只罩住本地读改写，不跨云端请求。 */
const MANIFEST_LOCK_FILE = 'manifest.lock';
const MANIFEST_LOCK_TIMEOUT_MS = 15000;
const MANIFEST_LOCK_STALE_MS = 120000;

const execFileAsync = promisify(execFile);

/** 图库在磁盘上的固定布局。 */
export function galleryPaths(root) {
    const base = resolve(root);
    return {
        root: base,
        manifest: join(base, 'manifest.json'),
        originals: join(base, 'originals'),
        thumbs: join(base, 'thumbs'),
        vectors: join(base, 'vectors'),
    };
}

function asText(value) {
    if (typeof value === 'string') return value;
    if (value === null || value === undefined) return '';
    return String(value);
}

function sha256Hex(value) {
    return createHash('sha256').update(value).digest('hex');
}

export function hashBytes(bytes) {
    return sha256Hex(bytes);
}

/** 去掉控制字符、折叠空白、限长；审美描述与标签都走这里，避免把不可信文本原样塞进 UI。 */
export function cleanText(value, max = 200) {
    const text = asText(value)
        .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    return text.length > max ? text.slice(0, max) : text;
}

/**
 * 文件夹路径规范化：只做逻辑分层，不落成真实目录。
 * 接受 / 或 \\\\ 分隔、首尾斜杠与多余空格；拒绝 . / ..（避免把 UI 路径变成
 * 越界语义）、超过 8 层或单段超过 40 字。空串表示「根目录 / 未归类」。
 */
function folderPathOf(raw) {
    const text = asText(raw).trim();
    if (text === '') return { ok: true, path: '' };
    const out = [];
    for (const part of text.replace(/\\/g, '/').split('/')) {
        const name = part
            .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
        if (name === '') continue;
        if (name === '.' || name === '..') {
            return { ok: false, error: '文件夹路径不能包含 . 或 ..：' + text };
        }
        if (/^\.+$/.test(name)) {
            return { ok: false, error: '文件夹名不能只有点：' + name };
        }
        if (name.length > MAX_FOLDER_NAME_CHARS) {
            return { ok: false, error: '文件夹名太长（最多 ' + MAX_FOLDER_NAME_CHARS + ' 字）：' + name };
        }
        out.push(name);
    }
    if (out.length > MAX_FOLDER_DEPTH) {
        return { ok: false, error: '文件夹层级太深（最多 ' + MAX_FOLDER_DEPTH + ' 层）：' + text };
    }
    const path = out.join('/');
    if (path.length > MAX_FOLDER_CHARS) {
        return { ok: false, error: '文件夹路径太长（最多 ' + MAX_FOLDER_CHARS + ' 字）：' + path };
    }
    return { ok: true, path };
}

/** 规范化文件夹路径；非法路径抛错，调用方要转成明确的 bad-folder 回执。 */
export function normalizeFolderPath(raw) {
    const result = folderPathOf(raw);
    if (!result.ok) throw new Error(result.error);
    return result.path;
}

/** 路径的祖先链（含自身）：A/B/C → A、A/B、A/B/C。 */
function folderChain(path) {
    const out = [];
    if (path === '') return out;
    const parts = path.split('/');
    for (let index = 1; index <= parts.length; index += 1) out.push(parts.slice(0, index).join('/'));
    return out;
}

function folderOf(item) {
    return asText(item?.manual?.folder);
}

function folderFilterOf(filter) {
    const parsed = folderPathOf(filter?.folder);
    if (!parsed.ok) return null;
    return parsed.path;
}

/** 标签规范化：支持数组或逗号串；去空白、去重（不分大小写）、限长限量。 */
export function normalizeTags(raw) {
    const list = Array.isArray(raw)
        ? raw
        : (typeof raw === 'string' && raw.trim() !== '' ? raw.split(/[,，、;；\n]+/) : []);
    const seen = new Set();
    const tags = [];
    for (const item of list) {
        const tag = cleanText(item, MAX_TAG_CHARS);
        if (tag === '') continue;
        const key = tag.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        tags.push(tag);
        if (tags.length >= MAX_TAGS) break;
    }
    return tags;
}

/** 审美描述规范化：只保留六个画面维度 + palette + summary + tags；作者/版权类字段直接丢弃。 */
const DROPPED_AESTHETIC_FIELDS = ['author', 'artist', 'copyright', 'credit', 'license', 'owner', 'source'];

export function normalizeAesthetic(raw, meta = {}) {
    const source = raw !== null && typeof raw === 'object' && !Array.isArray(raw) ? raw : {};
    const out = {};
    for (const facet of AESTHETIC_FACETS) {
        out[facet] = cleanText(source[facet], facet === 'subject' || facet === 'composition' ? 200 : 140);
    }
    out.palette = normalizeTags(source.palette).slice(0, PALETTE_MAX).map((color) => cleanText(color, 40));
    out.summary = cleanText(source.summary, SUMMARY_MAX);
    out.tags = normalizeTags(source.tags).slice(0, 24);
    out.source = 'model';
    out.channel = cleanText(meta.channel, 20);
    out.provider = cleanText(meta.provider, 80);
    out.model = cleanText(meta.model, 120);
    out.promptVersion = cleanText(meta.promptVersion ?? '', 40);
    out.generatedAt = cleanText(meta.generatedAt ?? new Date().toISOString(), 40);
    const dropped = Object.keys(source).filter((key) => DROPPED_AESTHETIC_FIELDS.includes(key.toLowerCase()));
    if (dropped.length > 0) out.droppedFields = dropped;
    return out;
}

/**
 * 参与检索的文本：只拼名称、审美描述、用户改写的描述、标签、分类与细节；
 * **不拼来源路径/哈希/尺寸**这些事实字段（否则「搜 1920」会把所有大图都召回，
 * 语义也全被污染）。改写描述存在时不再重复模型总结，避免两段近义文本互相稀释。
 */
export function documentTextOf(item) {
    const aesthetic = item?.aesthetic ?? {};
    const manual = item?.manual ?? {};
    const parts = [];
    if (cleanText(manual.name, MAX_NAME_CHARS) !== '') parts.push(cleanText(manual.name, MAX_NAME_CHARS));
    for (const facet of AESTHETIC_FACETS) {
        if (cleanText(aesthetic[facet], 400) !== '') parts.push(cleanText(aesthetic[facet], 400));
    }
    for (const color of aesthetic.palette ?? []) {
        const text = cleanText(color, 64);
        if (text !== '') parts.push(text);
    }
    const description = cleanText(manual.description, MAX_DESCRIPTION_CHARS) || cleanText(aesthetic.summary, SUMMARY_MAX);
    if (description !== '') parts.push(description);
    for (const tag of aesthetic.tags ?? []) {
        const text = cleanText(tag, MAX_TAG_CHARS);
        if (text !== '') parts.push(text);
    }
    for (const tag of manual.tags ?? []) {
        const text = cleanText(tag, MAX_TAG_CHARS);
        if (text !== '') parts.push(text);
    }
    if (cleanText(manual.category, MAX_CATEGORY_CHARS) !== '') parts.push(cleanText(manual.category, MAX_CATEGORY_CHARS));
    if (cleanText(manual.note, MAX_NOTE_CHARS) !== '') parts.push(cleanText(manual.note, MAX_NOTE_CHARS));
    return parts.join(' ').trim();
}

export function documentHashOf(item) {
    return sha256Hex(documentTextOf(item));
}

function digest12(text) {
    return sha256Hex(text).slice(0, 12);
}

/** 文本向量目录名：模型 + 维度 + 版本都进 key，换模型不会读到旧空间。 */
export function textVectorKey(model, dim) {
    return 'txt_' + digest12(asText(model) + '|' + dim + '|' + TEXT_VECTOR_VERSION) + '__' + dim;
}

/** 图文对齐空间的稳定标识：通道 + 端点 + 模型 + 版本（不含维度，维度在目录名里）。 */
export function imageSpaceKey(imageConfig) {
    const cfg = imageConfig ?? {};
    return digest12([cfg.provider, cfg.baseURL, cfg.model, IMAGE_EMBED_VERSION].join('|'));
}

export function imageVectorKey(imageConfig, dim) {
    return 'img_' + imageSpaceKey(imageConfig) + '__' + dim;
}

function isInside(parent, target) {
    const rel = relative(parent, target);
    return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel));
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Windows 上 rename 可能被杀毒 / 索引器 / 仍开着的句柄瞬时挡住（EPERM/EBUSY）。
 * 只重试这几类瞬时错误，重试耗尽就把原始错误抛出去，不吞掉。
 */
async function renameWithRetry(from, to, attempts = 8) {
    let last;
    for (let index = 0; index < attempts; index += 1) {
        try {
            await rename(from, to);
            return;
        } catch (error) {
            last = error;
            const transient = error?.code === 'EPERM' || error?.code === 'EACCES' || error?.code === 'EBUSY';
            if (!transient || index === attempts - 1) throw error;
            await sleep(10 + index * 15);
        }
    }
    throw last;
}

async function writeAtomic(target, data) {
    await mkdir(dirname(target), { recursive: true });
    const tmp = target + '.tmp-' + process.pid + '-' + Math.random().toString(16).slice(2);
    await writeFile(tmp, data);
    try {
        await renameWithRetry(tmp, target);
    } catch (error) {
        try { await unlink(tmp); } catch { /* 临时文件可能已经被 rename 或不在 */ }
        throw new Error('原子写 ' + target + ' 失败（' + (error?.code ?? 'unknown') + '）：' + (error?.message ?? String(error)));
    }
}

/**
 * Pillow 缩略图生成器（可选依赖）。Pillow 不在时返回 ok:false 与原因，
 * 图库照常保存原图，只是没有缩略图 —— 不拿原图冒充缩略图。
 */
export function createPillowThumbnailer({ python = 'python', timeoutMs = 30000 } = {}) {
    const script = [
        'import sys',
        'from PIL import Image',
        'src, dst, side = sys.argv[1], sys.argv[2], int(sys.argv[3])',
        'im = Image.open(src)',
        'im.draft("RGB", (side, side))',
        'im.thumbnail((side, side))',
        'if im.mode in ("RGBA", "LA", "P"):',
        '    im = im.convert("RGBA")',
        '    bg = Image.new("RGB", im.size, (255, 255, 255))',
        '    bg.paste(im, mask=im.split()[-1])',
        '    im = bg',
        'elif im.mode != "RGB":',
        '    im = im.convert("RGB")',
        'im.save(dst, "JPEG", quality=82, optimize=True)',
    ].join('\n');
    return async function pillowThumbnail({ sourcePath, targetPath, maxSide = 480 }) {
        try {
            await execFileAsync(python, ['-c', script, sourcePath, targetPath, String(maxSide)], {
                timeout: timeoutMs,
                windowsHide: true,
                maxBuffer: 1024 * 1024,
            });
        } catch (error) {
            const stderr = typeof error?.stderr === 'string' ? error.stderr.trim() : '';
            return { ok: false, reason: 'Pillow 缩略图失败：' + (stderr || error?.message || String(error)) };
        }
        try {
            const info = await stat(targetPath);
            if (info.size === 0) return { ok: false, reason: 'Pillow 生成的缩略图是空文件。' };
            return { ok: true, mediaType: 'image/jpeg' };
        } catch {
            return { ok: false, reason: 'Pillow 没有写出缩略图文件。' };
        }
    };
}

/** 词法打分（纯函数，测试与检索用同一份）：查询 token 的 IDF 覆盖度，0..1。 */
export function lexicalScores(items, query) {
    const queryTokens = [...new Set(kbTokens(asText(query)))];
    const scores = new Map();
    if (queryTokens.length === 0) return scores;
    const docs = items.map((item) => ({ id: item.id, tokens: new Set(kbTokens(documentTextOf(item))) }));
    const documentFrequency = new Map();
    for (const doc of docs) {
        for (const token of doc.tokens) {
            documentFrequency.set(token, (documentFrequency.get(token) ?? 0) + 1);
        }
    }
    const total = docs.length;
    const idf = (token) => {
        const df = documentFrequency.get(token) ?? 0;
        return Math.log(1 + (total - df + 0.5) / (df + 0.5));
    };
    const denominator = queryTokens.reduce((sum, token) => sum + idf(token), 0) || 1;
    for (const doc of docs) {
        let sum = 0;
        for (const token of queryTokens) {
            if (doc.tokens.has(token)) sum += idf(token);
        }
        scores.set(doc.id, Math.min(1, sum / denominator));
    }
    return scores;
}

/** 标签精确/包含匹配分（0 / 0.6 / 1）：用于「搜一个标签」这种最常用的动作。 */
export function tagScoreOf(item, query) {
    const wanted = asText(query).trim().toLowerCase();
    if (wanted === '') return 0;
    const tags = [
        ...(item?.manual?.tags ?? []),
        ...(item?.aesthetic?.tags ?? []),
    ].map((tag) => asText(tag).toLowerCase());
    if (tags.includes(wanted)) return 1;
    if (tags.some((tag) => tag.includes(wanted) || wanted.includes(tag))) return 0.6;
    return 0;
}

function summarize(item, extra = {}) {
    const manual = item.manual ?? { tags: [], category: '', favorite: false, note: '' };
    const aesthetic = item.aesthetic ?? null;
    const tags = [...new Set([...(manual.tags ?? []), ...((aesthetic?.tags) ?? [])])];
    return {
        id: item.id,
        name: manual.name ?? '',
        description: cleanText(manual.description, MAX_DESCRIPTION_CHARS) || cleanText(aesthetic?.summary, SUMMARY_MAX),
        bytes: item.facts?.bytes ?? 0,
        mediaType: item.facts?.mediaType ?? '',
        width: item.facts?.width ?? null,
        height: item.facts?.height ?? null,
        addedAt: item.facts?.addedAt ?? item.createdAt ?? '',
        thumb: item.thumb ?? null,
        category: manual.category ?? '',
        folder: manual.folder ?? '',
        favorite: manual.favorite === true,
        note: manual.note ?? '',
        tags,
        manualTags: [...(manual.tags ?? [])],
        aesthetic: aesthetic === null ? null : {
            subject: aesthetic.subject ?? '',
            composition: aesthetic.composition ?? '',
            palette: [...(aesthetic.palette ?? [])],
            light: aesthetic.light ?? '',
            material: aesthetic.material ?? '',
            mood: aesthetic.mood ?? '',
            summary: aesthetic.summary ?? '',
            tags: [...(aesthetic.tags ?? [])],
            model: aesthetic.model ?? '',
            provider: aesthetic.provider ?? '',
            channel: aesthetic.channel ?? '',
            generatedAt: aesthetic.generatedAt ?? '',
            ...(Array.isArray(aesthetic.droppedFields) ? { droppedFields: [...aesthetic.droppedFields] } : {}),
        },
        ...extra,
    };
}

/**
 * 图库目录：显式 memory.galleryDir 优先，否则与知识库同级（kbDir 的兄弟目录 gallery）。
 * 与 kbRootDir 同样的收口方式，避免图库把数据写到 KB 里面。
 */
export function galleryRootDir(memoryConfig = {}, kbRoot = '') {
    const explicit = typeof memoryConfig.galleryDir === 'string' ? memoryConfig.galleryDir.trim() : '';
    if (explicit !== '') {
        // 设置页约定：相对路径相对用户主目录（与 kbDir / globalDir 同一口径），
        // 不跟着某个工作目录搬家。
        return isAbsolute(explicit) ? resolve(explicit) : resolve(homedir(), explicit);
    }
    const kb = typeof kbRoot === 'string' ? kbRoot.trim() : '';
    if (kb === '') return '';
    return join(dirname(resolve(kb)), GALLERY_DIRNAME);
}

/**
 * 同一物理图库的进程内共享状态。
 *
 * 面板长活实例、工具调用实例、不同 cwd 的实例都可能指向同一个 gallery/；
 * 如果各自缓存一份 manifest，后写者会把先写者的条目覆盖掉。这里按规范化后的
 * 物理目录共享同一份 state 与串行写链，并在每次读写前用 manifest 的
 * mtime/ctime/size/inode 判断磁盘是否被别的进程改过，改过就重读。
 */
const GALLERY_STORES = new Map();

function physicalRootKey(root) {
    const resolved = resolve(root);
    let physical = resolved;
    try {
        physical = realpathSync(resolved);
    } catch {
        try {
            physical = join(realpathSync(dirname(resolved)), basename(resolved));
        } catch {
            physical = resolved;
        }
    }
    return process.platform === 'win32' ? physical.toLowerCase() : physical;
}

function galleryStoreFor(root) {
    const key = physicalRootKey(root);
    let store = GALLERY_STORES.get(key);
    if (store === undefined) {
        store = { state: null, stamp: null, loading: null, chain: Promise.resolve(), version: 0 };
        GALLERY_STORES.set(key, store);
    }
    return store;
}

/**
 * 创建图库实例。
 *
 * @param {object} options
 * @param {string} options.root 图库目录（绝对路径；缺省为 cwd/.office/gallery）
 * @param {string} [options.cwd] 相对路径的基准与会话工作目录
 * @param {string[]} [options.roots] 允许导入图片的信任根（缺省 cwd 与图库目录）
 * @param {Function|null} [options.thumbnailer] 缩略图生成器；传 null 表示明确不生成
 * @param {object} [options.config] 图库配置（文本向量 / 图片向量 / 视觉通道）
 */
export function createGallery(options = {}) {
    const root = resolve(options.root ?? join(process.cwd(), '.office', GALLERY_DIRNAME));
    const paths = galleryPaths(root);
    const cwd = resolve(options.cwd ?? process.cwd());
    const trustRoots = (Array.isArray(options.roots) ? options.roots : [cwd, root])
        .filter((value) => typeof value === 'string' && value !== '')
        .map((value) => resolve(value));
    const maxBytes = Number.isFinite(options.maxBytes) ? options.maxBytes : MAX_IMAGE_BYTES;
    const now = typeof options.now === 'function' ? options.now : () => new Date().toISOString();
    const fetchImpl = options.fetchImpl ?? globalThis.fetch;
    const env = options.env ?? process.env;
    const thumbnailer = options.thumbnailer === undefined
        ? createPillowThumbnailer(options.pillow ?? {})
        : options.thumbnailer;
    let config = normalizeGalleryConfig(options.config ?? {});
    let lastThumbError = '';
    let state = null;
    const store = galleryStoreFor(root);
    const lockTimeoutMs = Number.isFinite(options.lockTimeoutMs)
        ? Math.max(50, Math.trunc(options.lockTimeoutMs))
        : MANIFEST_LOCK_TIMEOUT_MS;
    const lockStaleMs = Number.isFinite(options.lockStaleMs)
        ? Math.max(1000, Math.trunc(options.lockStaleMs))
        : MANIFEST_LOCK_STALE_MS;
    const lockPath = join(root, MANIFEST_LOCK_FILE);

    function adoptState(next) {
        state = next;
        return next;
    }

    async function readLockInfo() {
        try {
            const raw = await readFile(lockPath, 'utf8');
            try {
                const parsed = JSON.parse(raw);
                return parsed !== null && typeof parsed === 'object' ? parsed : { raw: raw.slice(0, 200) };
            } catch {
                return { raw: raw.slice(0, 200) };
            }
        } catch (error) {
            return error?.code === 'ENOENT' ? null : { raw: error?.message ?? String(error) };
        }
    }

    function lockTimeoutError(info) {
        const ageMs = info && typeof info.at === 'string' ? Date.now() - Date.parse(info.at) : NaN;
        const stale = Number.isFinite(ageMs) && ageMs > lockStaleMs;
        const holder = info === null
            ? '锁文件已不存在'
            : info.raw !== undefined
                ? (String(info.raw).trim() === ''
                    ? '锁文件为空（0 字节），可能是进程创建后没来得及写入就退出了'
                    : '锁文件内容不可解析：' + String(info.raw))
                : 'pid=' + String(info.pid ?? '?') + ' host=' + String(info.host ?? '?') + ' at=' + String(info.at ?? '?');
        return {
            ok: false,
            code: 'lock-timeout',
            error: '图库清单锁 ' + lockPath + ' 在 ' + lockTimeoutMs + 'ms 内没拿到（' + holder
                + (stale ? '，已超过 ' + lockStaleMs + 'ms，可能是崩溃残留' : '')
                + '）。确认没有进程在写后手动删除这个锁文件再重试；不会自动删除，避免误删仍在使用的锁。',
        };
    }

    /**
     * 跨进程 manifest 锁：open(..., 'wx') 是原子的独占创建，Windows / POSIX 都可用。
     * 只罩住本地「读最新 → 改 → 原子写」，不把云端请求包进来。
     * 锁文件残留时不自动删（避免 TOCTOU 误删活锁），超时错误里给出手动恢复路径。
     */
    async function acquireManifestLock() {
        await mkdir(root, { recursive: true });
        const deadline = Date.now() + lockTimeoutMs;
        for (;;) {
            const token = randomUUID();
            let handle = null;
            try {
                handle = await open(lockPath, 'wx');
                await handle.writeFile(JSON.stringify({
                    token,
                    pid: process.pid,
                    host: hostname(),
                    at: new Date().toISOString(),
                    kind: 'gallery-manifest',
                }));
                await handle.close();
                return { ok: true, path: lockPath, token };
            } catch (error) {
                if (handle !== null) {
                    try { await handle.close(); } catch { /* 已经关掉 */ }
                }
                // EEXIST 是正常的锁竞争；Windows 上杀毒/索引器还可能在文件刚创建或
                // 刚删除时给 EPERM/EACCES/EBUSY，这些也按瞬时竞争重试，不能直接崩掉上传。
                const retryable = error?.code === 'EEXIST'
                    || error?.code === 'EPERM'
                    || error?.code === 'EACCES'
                    || error?.code === 'EBUSY';
                if (!retryable) throw error;
                const info = await readLockInfo();
                if (Date.now() >= deadline) return lockTimeoutError(info);
                await sleep(15 + Math.floor(Math.random() * 20));
            }
        }
    }

    async function releaseManifestLock(lock) {
        let info;
        try {
            info = JSON.parse(await readFile(lock.path, 'utf8'));
        } catch {
            return;
        }
        if (info === null || typeof info !== 'object' || info.token !== lock.token) return;
        // Windows 上杀毒/索引器可能短暂占着锁文件，unlink 会 EPERM/EBUSY；重试几次，
        // 仍失败就抛给 withWriteLock 的 finally，让下一个写者按超时文案人工恢复。
        for (let attempt = 0; ; attempt += 1) {
            try {
                await unlink(lock.path);
                return;
            } catch (error) {
                if (error?.code === 'ENOENT') return;
                const retryable = error?.code === 'EPERM' || error?.code === 'EACCES' || error?.code === 'EBUSY';
                if (!retryable || attempt >= 5) throw error;
                await sleep(10 + attempt * 15);
            }
        }
    }

    function withWriteLock(fn) {
        const run = async () => {
            const lock = await acquireManifestLock();
            if (lock.ok !== true) return lock;
            try {
                return await fn();
            } finally {
                try {
                    await releaseManifestLock(lock);
                } catch {
                    // 写已经完成；释放失败会留下锁，下一个写者会超时并看到手动恢复说明。
                }
            }
        };
        const next = store.chain.then(run, run);
        store.chain = next.then(() => undefined, () => undefined);
        return next;
    }

    /** manifest 的磁盘指纹：任何一个字段变了都当外部修改，重读。 */
    async function manifestStamp() {
        try {
            const info = await stat(paths.manifest);
            return [info.mtimeMs, info.ctimeMs, info.size, info.ino].join(':');
        } catch (error) {
            if (error?.code === 'ENOENT') return 'missing';
            throw error;
        }
    }

    async function readManifestFromDisk() {
        try {
            const raw = await readFile(paths.manifest, 'utf8');
            const parsed = JSON.parse(raw);
            if (parsed === null || typeof parsed !== 'object' || parsed.items === null || typeof parsed.items !== 'object') {
                throw new Error('manifest 结构损坏（没有 items 对象）');
            }
            // 旧 manifest 没有 folders：按图片里的 folder 反推祖先，读的时候补齐，
            // 下一次写入就会把显式目录落盘。非法路径不能带进 state，否则过滤会失真。
            const folders = new Set();
            for (const value of Array.isArray(parsed.folders) ? parsed.folders : []) {
                const path = folderPathOf(value);
                if (path.ok) for (const ancestor of folderChain(path.path)) folders.add(ancestor);
            }
            for (const item of Object.values(parsed.items)) {
                if (item === null || typeof item !== 'object') continue;
                if (item.manual === null || typeof item.manual !== 'object') item.manual = {};
                const path = folderPathOf(item.manual.folder);
                item.manual.folder = path.ok ? path.path : '';
                for (const ancestor of folderChain(item.manual.folder)) folders.add(ancestor);
            }
            return {
                version: GALLERY_VERSION,
                updatedAt: asText(parsed.updatedAt),
                folders: [...folders].sort((left, right) => left.localeCompare(right, 'zh')),
                items: parsed.items,
            };
        } catch (error) {
            if (error?.code === 'ENOENT') {
                return { version: GALLERY_VERSION, updatedAt: '', folders: [], items: {} };
            }
            throw new Error('图库清单读不了：' + (error?.message ?? String(error)));
        }
    }

    async function loadState(force = false) {
        if (force) {
            // 写事务已经在跨进程锁里：必须绕过 stamp 缓存，直接从磁盘读最新一份。
            if (store.loading !== null) await store.loading.catch(() => {});
            const next = await readManifestFromDisk();
            store.state = next;
            store.stamp = await manifestStamp();
            store.version += 1;
            return adoptState(next);
        }
        for (;;) {
            const stamp = await manifestStamp();
            if (store.state !== null && store.stamp === stamp) return adoptState(store.state);
            if (store.loading !== null) {
                // 别的实例正在读；等它读完再按新指纹判断一次，避免拿到旧对象。
                await store.loading.catch(() => {});
                continue;
            }
            const expectedVersion = store.version;
            store.loading = (async () => {
                const next = await readManifestFromDisk();
                const after = await manifestStamp();
                // 读的过程中有别的实例写入（或磁盘被外部改了）：这份数据已经过期，
                // 丢掉，让循环按新指纹重读，绝不能把旧 state 盖在新 stamp 上。
                if (store.version !== expectedVersion || after !== stamp) return null;
                store.state = next;
                store.stamp = after;
                return next;
            })();
            try {
                await store.loading;
            } finally {
                store.loading = null;
            }
            // 回到循环顶部再核一次指纹：读的过程中磁盘又变了就重读。
        }
    }

    async function saveState() {
        adoptState(store.state);
        state.updatedAt = now();
        try {
            await writeAtomic(paths.manifest, JSON.stringify(state, null, 2));
        } catch (error) {
            // 写失败时内存里的 state 已经改过；不能让下一次读继续拿这份
            // 「磁盘上不存在」的快照，清掉缓存强制从磁盘重读。
            store.state = null;
            store.stamp = null;
            throw error;
        }
        store.stamp = await manifestStamp();
        store.version += 1;
    }

    async function ensureDirs() {
        await mkdir(paths.originals, { recursive: true });
        await mkdir(paths.thumbs, { recursive: true });
        await mkdir(paths.vectors, { recursive: true });
    }

    /** 把路径及其祖先登记进 manifest.folders（显式空目录靠它保留）。 */
    function addFolderState(path) {
        if (!Array.isArray(state.folders)) state.folders = [];
        const known = new Set(state.folders);
        for (const ancestor of folderChain(path)) known.add(ancestor);
        state.folders = [...known].sort((left, right) => left.localeCompare(right, 'zh'));
    }

    /**
     * 文件夹视图：显式目录与图片里的目录合并，按路径排序。
     * direct 是直属图片数，total 含所有后代；这样父目录空着也能显示，且不会
     * 因为只看图片而丢掉用户手工建的空目录。
     */
    function folderRows() {
        const known = new Set(Array.isArray(state.folders) ? state.folders : []);
        const direct = new Map();
        const total = new Map();
        for (const item of Object.values(state.items)) {
            const path = folderOf(item);
            if (path === '') continue;
            for (const ancestor of folderChain(path)) {
                known.add(ancestor);
                total.set(ancestor, (total.get(ancestor) ?? 0) + 1);
            }
            direct.set(path, (direct.get(path) ?? 0) + 1);
        }
        return [...known].sort((left, right) => left.localeCompare(right, 'zh')).map((path) => {
            const parts = path.split('/');
            return {
                path,
                name: parts[parts.length - 1],
                parent: parts.length > 1 ? parts.slice(0, -1).join('/') : '',
                depth: parts.length,
                direct: direct.get(path) ?? 0,
                total: total.get(path) ?? 0,
            };
        });
    }

    /**
     * 新建文件夹（幂等）。父目录自动补齐；路径只存 manifest，不碰磁盘目录，
     * 这样 Windows / Linux 上同一份图库的行为一致。
     */
    async function createFolder(rawPath) {
        const parsed = folderPathOf(rawPath);
        if (!parsed.ok) return { ok: false, code: 'bad-folder', error: parsed.error };
        if (parsed.path === '') {
            return { ok: false, code: 'bad-folder', error: '根目录不是可新建的文件夹；给一个非空路径。' };
        }
        const path = parsed.path;
        return withWriteLock(async () => {
            await loadState(true);
            // 图片已经归到这个路径时，目录在视图里存在但不在 manifest.folders：
            // 也算「已存在」，否则回执会把老目录说成刚建好。
            const existed = folderRows().some((entry) => entry.path === path);
            addFolderState(path);
            if (!existed) await saveState();
            return {
                ok: true,
                created: !existed,
                path,
                folder: folderRows().find((entry) => entry.path === path) ?? null,
            };
        });
    }

    /** 重命名文件夹（含所有后代与图片），拒绝移动到自己的子目录或已存在的目标。 */
    async function renameFolder(rawFrom, rawTo) {
        const from = folderPathOf(rawFrom);
        const to = folderPathOf(rawTo);
        if (!from.ok) return { ok: false, code: 'bad-folder', error: from.error };
        if (!to.ok) return { ok: false, code: 'bad-folder', error: to.error };
        if (from.path === '') return { ok: false, code: 'bad-folder', error: '要给出原来的文件夹路径。' };
        if (to.path === '') return { ok: false, code: 'bad-folder', error: '目标文件夹不能是根目录。' };
        if (from.path === to.path) return { ok: true, from: from.path, to: to.path, moved: 0, renamedFolders: 0 };
        if (to.path.startsWith(from.path + '/')) {
            return { ok: false, code: 'bad-folder', error: '不能把文件夹移动到它自己的子目录：' + from.path + ' → ' + to.path };
        }
        return withWriteLock(async () => {
            await loadState(true);
            const before = folderRows();
            const known = new Set(before.map((entry) => entry.path));
            if (!known.has(from.path) && !before.some((entry) => entry.path.startsWith(from.path + '/'))) {
                return { ok: false, code: 'folder-not-found', error: '没有这个文件夹：' + from.path };
            }
            const rewritten = (path) => (path === from.path || path.startsWith(from.path + '/'))
                ? to.path + path.slice(from.path.length)
                : path;
            if (known.has(to.path)) {
                return { ok: false, code: 'folder-exists', error: '目标文件夹已存在：' + to.path };
            }
            for (const entry of before) {
                if (entry.path !== from.path && !entry.path.startsWith(from.path + '/')) continue;
                const target = rewritten(entry.path);
                if (known.has(target)) {
                    return { ok: false, code: 'folder-exists', error: '目标位置已有同名文件夹：' + target };
                }
            }
            const folders = (Array.isArray(state.folders) ? state.folders : []).map(rewritten);
            state.folders = [...new Set(folders)].sort((left, right) => left.localeCompare(right, 'zh'));
            let moved = 0;
            let renamedFolders = 0;
            for (const item of Object.values(state.items)) {
                const path = folderOf(item);
                if (path !== from.path && !path.startsWith(from.path + '/')) continue;
                item.manual.folder = rewritten(path);
                item.updatedAt = now();
                moved += 1;
            }
            for (const entry of before) {
                if (entry.path === from.path || entry.path.startsWith(from.path + '/')) renamedFolders += 1;
            }
            addFolderState(to.path);
            await saveState();
            return { ok: true, from: from.path, to: to.path, moved, renamedFolders };
        });
    }

    /** 删除空文件夹（含其后代空目录）；还有图片时明确拒绝，绝不悄悄把图片搬走。 */
    async function removeFolder(rawPath) {
        const parsed = folderPathOf(rawPath);
        if (!parsed.ok) return { ok: false, code: 'bad-folder', error: parsed.error };
        if (parsed.path === '') return { ok: false, code: 'bad-folder', error: '要给出要删除的文件夹路径。' };
        const path = parsed.path;
        return withWriteLock(async () => {
            await loadState(true);
            const before = folderRows();
            const inSubtree = (entryPath) => entryPath === path || entryPath.startsWith(path + '/');
            if (!before.some((entry) => inSubtree(entry.path))) {
                return { ok: false, code: 'folder-not-found', error: '没有这个文件夹：' + path };
            }
            const total = before.find((entry) => entry.path === path)?.total ?? 0;
            const direct = before.find((entry) => entry.path === path)?.direct ?? 0;
            if (total > 0) {
                return {
                    ok: false,
                    code: 'folder-not-empty',
                    error: '文件夹里还有 ' + total + ' 张图（直属 ' + direct + ' 张），先把图片移到别处再删。',
                };
            }
            const removed = before.filter((entry) => inSubtree(entry.path)).length;
            state.folders = (Array.isArray(state.folders) ? state.folders : [])
                .filter((entry) => !inSubtree(entry));
            await saveState();
            return { ok: true, path, removedFolders: removed };
        });
    }

    async function listFolders() {
        await loadState();
        return { ok: true, folders: folderRows() };
    }

    async function realRoots() {
        const out = [];
        for (const value of [...trustRoots, root]) {
            try {
                out.push(await realpath(value));
            } catch {
                out.push(value);
            }
        }
        return [...new Set(out.map((value) => resolve(value)))];
    }

    async function resolveTrusted(inputPath) {
        const shown = asText(inputPath).trim();
        if (shown === '') return { ok: false, code: 'bad-path', error: '要给出图片路径。' };
        const absolute = isAbsolute(shown) ? resolve(shown) : resolve(cwd, shown);
        let real;
        try {
            real = await realpath(absolute);
        } catch (error) {
            return {
                ok: false,
                code: error?.code === 'ENOENT' ? 'not-found' : 'unreadable',
                error: '读不到这个文件：' + (error?.message ?? String(error)),
            };
        }
        const trusted = await realRoots();
        if (!trusted.some((value) => isInside(value, real))) {
            return {
                ok: false,
                code: 'outside-root',
                error: '只接受会话工作目录、知识库目录或图库目录内的图片；这个路径在信任边界外：' + real,
                roots: trusted,
            };
        }
        let info;
        try {
            info = await stat(real);
        } catch (error) {
            return { ok: false, code: 'unreadable', error: '读不到文件状态：' + (error?.message ?? String(error)) };
        }
        if (!info.isFile()) return { ok: false, code: 'not-file', error: '不是普通文件。' };
        if (info.size === 0) return { ok: false, code: 'empty', error: '文件是空的（0 字节）。' };
        if (info.size > maxBytes) {
            return { ok: false, code: 'too-large', error: '图片 ' + info.size + ' 字节，超过上限 ' + maxBytes + ' 字节。' };
        }
        return { ok: true, path: real, bytes: info.size, roots: trusted };
    }

    async function makeThumb(sourcePath, targetPath) {
        if (typeof thumbnailer !== 'function') {
            return { ok: false, reason: '没有配置缩略图生成器（Pillow 不可用）。' };
        }
        try {
            const out = await thumbnailer({ sourcePath, targetPath, maxSide: 480 });
            if (out?.ok === true) {
                return { ok: true, mediaType: out.mediaType ?? 'image/jpeg' };
            }
            return { ok: false, reason: out?.reason ?? '缩略图生成器没有返回成功。' };
        } catch (error) {
            return { ok: false, reason: '缩略图生成失败：' + (error?.message ?? String(error)) };
        }
    }

    function sanitizeOrigin(raw) {
        const source = raw !== null && typeof raw === 'object' && !Array.isArray(raw) ? raw : {};
        const origin = { kind: cleanText(source.kind, 16) || 'unknown' };
        if (cleanText(source.path, 1024) !== '') origin.path = cleanText(source.path, 1024);
        if (cleanText(source.url, 2048) !== '') origin.url = cleanText(source.url, 2048);
        if (cleanText(source.label, 160) !== '') origin.label = cleanText(source.label, 160);
        return origin;
    }

    async function addBytes(inputBytes, meta = {}) {
        const bytes = Buffer.isBuffer(inputBytes) ? inputBytes : Buffer.from(inputBytes ?? []);
        if (bytes.length === 0) return { ok: false, code: 'empty', error: '图片是空的（0 字节）。' };
        if (bytes.length > maxBytes) {
            return { ok: false, code: 'too-large', error: '图片 ' + bytes.length + ' 字节，超过上限 ' + maxBytes + ' 字节。' };
        }
        const sniffed = sniffImageBytes(bytes);
        if (sniffed === undefined) {
            return {
                ok: false,
                code: 'unsupported',
                error: '不是认得的位图（只按魔数认 PNG / JPEG / GIF / WebP；SVG、TIFF、BMP 不入库）。',
            };
        }
        const width = Number.isFinite(sniffed.width) ? sniffed.width : null;
        const height = Number.isFinite(sniffed.height) ? sniffed.height : null;
        if (width !== null && height !== null && width * height > MAX_PIXELS) {
            return {
                ok: false,
                code: 'too-many-pixels',
                error: '图片解码后 ' + width + '×' + height + ' 像素，超过上限 ' + MAX_PIXELS + ' 像素。',
            };
        }
        const folder = folderPathOf(meta.folder);
        if (!folder.ok) return { ok: false, code: 'bad-folder', error: folder.error };
        const sha = hashBytes(bytes);
        const originalRel = 'originals/' + sha + '.' + sniffed.ext;
        const thumbRel = 'thumbs/' + sha + '.jpg';
        const tempSuffix = process.pid + '-' + randomUUID();
        const sourceTmp = join(root, '.thumb-source-' + sha + '-' + tempSuffix + '.' + sniffed.ext);
        const thumbTmp = join(paths.thumbs, sha + '.jpg.tmp-' + tempSuffix);
        let thumb = { ok: false, reason: '没有配置缩略图生成器（Pillow 不可用）。' };
        try {
            await ensureDirs();
            if (typeof thumbnailer === 'function') {
                await writeFile(sourceTmp, bytes);
                // 缩略图生成可能很慢（Pillow 默认上限 30s），必须在 manifest 锁外做，
                // 否则会占着跨进程锁让别的上传等到 lock-timeout。
                thumb = await makeThumb(sourceTmp, thumbTmp);
                if (!thumb.ok) lastThumbError = thumb.reason;
            }
        } catch (error) {
            thumb = { ok: false, reason: '缩略图准备失败：' + (error?.message ?? String(error)) };
            lastThumbError = thumb.reason;
        }
        const cleanupTemps = async () => {
            for (const tempPath of [sourceTmp, thumbTmp]) {
                try {
                    await unlink(tempPath);
                } catch {
                    // 临时文件可能已经不存在，或已经 rename 成正式缩略图。
                }
            }
        };
        try {
            return await withWriteLock(async () => {
                await loadState(true);
                await ensureDirs();
                const existing = state.items[sha];
                if (existing !== undefined) {
                    await cleanupTemps();
                    return {
                        ok: true,
                        added: false,
                        duplicate: true,
                        item: summarize(existing),
                        note: '同内容图片已存在（内容哈希相同），没有重复保存。',
                    };
                }
                try {
                    await writeAtomic(join(root, originalRel), bytes);
                } catch (error) {
                    await cleanupTemps();
                    return {
                        ok: false,
                        code: 'write-failed',
                        error: '写原图失败：' + (error?.message ?? String(error)),
                    };
                }
                let thumbStored = false;
                if (thumb.ok) {
                    try {
                        await renameWithRetry(thumbTmp, join(root, thumbRel));
                        thumbStored = true;
                    } catch (error) {
                        lastThumbError = '缩略图落盘失败：' + (error?.message ?? String(error));
                    }
                }
                const item = {
                    id: sha,
                    version: GALLERY_VERSION,
                    facts: {
                        sha256: sha,
                        bytes: bytes.length,
                        mediaType: sniffed.mime,
                        ext: sniffed.ext,
                        width,
                        height,
                        addedAt: now(),
                        origin: sanitizeOrigin(meta.origin),
                    },
                    original: originalRel,
                    thumb: thumbStored ? thumbRel : null,
                    aesthetic: null,
                    manual: {
                        name: cleanText(meta.name, MAX_NAME_CHARS),
                        description: cleanText(meta.description, MAX_DESCRIPTION_CHARS),
                        tags: normalizeTags(meta.tags),
                        category: cleanText(meta.category, MAX_CATEGORY_CHARS),
                        folder: folder.path,
                        favorite: meta.favorite === true,
                        note: cleanText(meta.note, MAX_NOTE_CHARS),
                    },
                    createdAt: now(),
                    updatedAt: now(),
                };
                state.items[sha] = item;
                addFolderState(folder.path);
                await saveState();
                return {
                    ok: true,
                    added: true,
                    duplicate: false,
                    item: summarize(item),
                    ...(thumbStored ? {} : { thumb: { ok: false, reason: lastThumbError || thumb.reason } }),
                };
            });
        } finally {
            await cleanupTemps();
        }
    }

    async function addFromPath(inputPath, meta = {}) {
        const resolved = await resolveTrusted(inputPath);
        if (!resolved.ok) return resolved;
        let bytes;
        try {
            bytes = await readFile(resolved.path);
        } catch (error) {
            return { ok: false, code: 'unreadable', error: '读不到这个文件：' + (error?.message ?? String(error)) };
        }
        if (bytes.length > maxBytes) {
            return { ok: false, code: 'too-large', error: '图片 ' + bytes.length + ' 字节，超过上限 ' + maxBytes + ' 字节。' };
        }
        return addBytes(bytes, {
            ...meta,
            name: cleanText(meta.name, MAX_NAME_CHARS) || basename(resolved.path),
            origin: { ...(meta.origin ?? {}), kind: 'file', path: resolved.path },
        });
    }

    async function get(id) {
        const key = asText(id).trim().toLowerCase();
        if (!/^[0-9a-f]{64}$/.test(key)) return null;
        await loadState();
        return state.items[key] ?? null;
    }

    function matchesFilters(item, filter = {}) {
        const folder = folderFilterOf(filter);
        // 非法路径（比如 ..）不能退化成「不过滤」，否则越界语义会变成全库。
        if (folder === null) return false;
        if (folder !== '' && folderOf(item) !== folder) return false;
        const wanted = normalizeTags(filter.tags ?? []);
        if (wanted.length > 0) {
            const owned = new Set([
                ...(item.manual?.tags ?? []),
                ...(item.aesthetic?.tags ?? []),
            ].map((tag) => asText(tag).toLowerCase()));
            if (!wanted.every((tag) => owned.has(tag.toLowerCase()))) return false;
        }
        if (cleanText(filter.category, MAX_CATEGORY_CHARS) !== '') {
            if ((item.manual?.category ?? '') !== cleanText(filter.category, MAX_CATEGORY_CHARS)) return false;
        }
        if (filter.favorite === true && item.manual?.favorite !== true) return false;
        if (filter.hasAesthetic === true && item.aesthetic === null) return false;
        return true;
    }

    async function list(filter = {}) {
        await loadState();
        const items = Object.values(state.items).filter((item) => matchesFilters(item, filter));
        items.sort((left, right) => asText(right.facts?.addedAt).localeCompare(asText(left.facts?.addedAt)));
        const offset = Number.isFinite(filter.offset) ? Math.max(0, Math.trunc(filter.offset)) : 0;
        const limit = Number.isFinite(filter.limit) ? Math.min(500, Math.max(1, Math.trunc(filter.limit))) : 200;
        return {
            ok: true,
            total: items.length,
            offset,
            limit,
            items: items.slice(offset, offset + limit).map((item) => summarize(item)),
        };
    }

    async function updateManual(id, patch = {}) {
        let nextFolder;
        if (patch.folder !== undefined) {
            const parsed = folderPathOf(patch.folder);
            if (!parsed.ok) return { ok: false, code: 'bad-folder', error: parsed.error };
            nextFolder = parsed.path;
        }
        const item = await get(id);
        if (item === null) return { ok: false, code: 'not-found', error: '图库里没有这张图。' };
        return withWriteLock(async () => {
            await loadState(true);
            const current = state.items[item.id];
            if (current === undefined) return { ok: false, code: 'not-found', error: '图库里没有这张图。' };
            const manual = current.manual ?? { tags: [], category: '', favorite: false, note: '' };
            let tags = [...(manual.tags ?? [])];
            if (patch.tags !== undefined) tags = normalizeTags(patch.tags);
            if (Array.isArray(patch.addTags)) tags = normalizeTags([...tags, ...patch.addTags]);
            if (Array.isArray(patch.removeTags)) {
                const remove = new Set(normalizeTags(patch.removeTags).map((tag) => tag.toLowerCase()));
                tags = tags.filter((tag) => !remove.has(tag.toLowerCase()));
            }
            current.manual = {
                name: patch.name !== undefined ? cleanText(patch.name, MAX_NAME_CHARS) : (manual.name ?? ''),
                description: patch.description !== undefined ? cleanText(patch.description, MAX_DESCRIPTION_CHARS) : (manual.description ?? ''),
                tags,
                category: patch.category !== undefined ? cleanText(patch.category, MAX_CATEGORY_CHARS) : (manual.category ?? ''),
                folder: nextFolder !== undefined ? nextFolder : folderOf(current),
                favorite: patch.favorite !== undefined ? patch.favorite === true : manual.favorite === true,
                note: patch.note !== undefined ? cleanText(patch.note, MAX_NOTE_CHARS) : (manual.note ?? ''),
            };
            addFolderState(current.manual.folder);
            current.updatedAt = now();
            await saveState();
            return { ok: true, item: summarize(current) };
        });
    }

    async function setAesthetic(id, raw, meta = {}) {
        const item = await get(id);
        if (item === null) return { ok: false, code: 'not-found', error: '图库里没有这张图。' };
        return withWriteLock(async () => {
            await loadState(true);
            const current = state.items[item.id];
            if (current === undefined) return { ok: false, code: 'not-found', error: '图库里没有这张图。' };
            current.aesthetic = raw === null ? null : normalizeAesthetic(raw, meta);
            // 重新生成视觉描述时，清掉用户上一次的改写，否则界面上永远显示旧文本。
            if (meta.resetManualDescription === true) current.manual = { ...(current.manual ?? {}), description: '' };
            current.updatedAt = now();
            await saveState();
            return { ok: true, item: summarize(current) };
        });
    }

    async function remove(id) {
        const item = await get(id);
        if (item === null) return { ok: false, code: 'not-found', error: '图库里没有这张图。' };
        return withWriteLock(async () => {
            await loadState(true);
            delete state.items[item.id];
            await saveState();
            const targets = [join(root, item.original)];
            if (item.thumb !== null) targets.push(join(root, item.thumb));
            for (const target of targets) {
                try { await unlink(target); } catch { /* 文件可能已不在 */ }
            }
            try {
                for (const kind of ['text', 'image']) {
                    const base = join(paths.vectors, kind);
                    const keys = await readdir(base).catch(() => []);
                    for (const key of keys) {
                        try { await unlink(join(base, key, item.id + '.json')); } catch { /* 没有这个向量 */ }
                    }
                }
            } catch { /* 向量目录可能还没建 */ }
            return { ok: true, removed: item.id };
        });
    }

    async function vectorPath(kind, key, id) {
        return join(paths.vectors, kind, key, id + '.json');
    }

    async function readVector(kind, key, id) {
        try {
            return JSON.parse(await readFile(await vectorPath(kind, key, id), 'utf8'));
        } catch {
            return null;
        }
    }

    async function writeVector(kind, key, id, payload) {
        await writeAtomic(await vectorPath(kind, key, id), JSON.stringify(payload));
    }

    async function readProbe() {
        try {
            return JSON.parse(await readFile(join(paths.vectors, 'probe-' + imageSpaceKey(config.imageEmbedding) + '.json'), 'utf8'));
        } catch {
            return null;
        }
    }

    function probeFile() {
        return join(paths.vectors, 'probe-' + imageSpaceKey(config.imageEmbedding) + '.json');
    }

    async function probeImage(options = {}) {
        if (options.allowUpload !== true) {
            return {
                ok: false,
                code: 'needs-confirmation',
                error: '探针会向多模态端点发送一条文本和一个 1×1 图片（可能计费）。确认后传 allowUpload:true。',
            };
        }
        const ready = imageEmbeddingReady(config.imageEmbedding);
        if (!ready.ok) return { ok: false, code: ready.code, error: ready.error };
        const embedder = createImageEmbedder(config.imageEmbedding, { fetchImpl, env });
        const result = await probeImageEmbedding(embedder, { signal: options.signal });
        const prior = await readProbe();
        if (result.ok !== true && prior !== null && prior.ok === true) {
            // 一次网络抖动不能把上次成功的探针抹掉：保留成功结果，只把这次错误
            // 作为 stale/lastError 带回去，检索继续按已验证的同维空间工作。
            return { ...prior, stale: true, lastError: result.error ?? '这次探针失败。' };
        }
        const payload = {
            ...result,
            provider: embedder.provider,
            model: embedder.model,
            baseURL: embedder.baseURL,
            space: imageSpaceKey(config.imageEmbedding),
            at: now(),
        };
        await writeAtomic(probeFile(), JSON.stringify(payload, null, 2));
        return payload;
    }

    async function ensureTextVectors(items, { force = false, signal, allowUpload = false } = {}) {
        const textConfig = config.textEmbedding;
        if (textConfig.mode === 'off') {
            return { ok: false, code: 'disabled', error: '文本向量在设置里关掉了。', written: 0, skipped: 0 };
        }
        const remoteConfigured = canUseRemoteEmbedding(textConfig);
        // 远端文本向量也发文本：描述、标签，索引时还有正文摘要。配置了端点不等于
        // 同意每次把内容发出去；未显式 allowUpload 时退回本地哈希向量，零请求。
        const localOnly = !remoteConfigured || allowUpload !== true;
        const effectiveConfig = remoteConfigured && localOnly ? { ...textConfig, mode: 'local' } : textConfig;
        const modelHint = localOnly ? LOCAL_EMBEDDING_MODEL : textConfig.model;
        const dimHint = textConfig.dimensions || (localOnly ? 256 : 0);
        const knownKey = modelHint !== '' && dimHint > 0 ? textVectorKey(modelHint, dimHint) : null;
        const work = [];
        let skipped = 0;
        for (const item of items) {
            const text = documentTextOf(item);
            if (text === '') {
                skipped += 1;
                continue;
            }
            const textHash = sha256Hex(text);
            if (!force && knownKey !== null) {
                const cached = await readVector('text', knownKey, item.id);
                if (cached !== null && cached.textHash === textHash && cached.vector?.length === dimHint) continue;
            }
            work.push({ item, text, textHash });
        }
        if (work.length === 0) {
            return {
                ok: true,
                written: 0,
                skipped,
                model: modelHint || (localOnly ? LOCAL_EMBEDDING_MODEL : ''),
                dim: dimHint,
                mode: localOnly ? 'local' : 'remote',
                degraded: false,
                error: '',
                note: remoteConfigured && localOnly ? '远端文本向量需要显式 allowUpload:true；这次用本地哈希向量，没有联网。' : '',
            };
        }
        const batchSize = textConfig.batchSize;
        let written = 0;
        let lastModel = '';
        let lastDim = 0;
        let mode = localOnly ? 'local' : 'remote';
        let degraded = false;
        let firstError = '';
        for (let start = 0; start < work.length; start += batchSize) {
            const batch = work.slice(start, start + batchSize);
            const result = await embedTexts({
                texts: batch.map((entry) => entry.text),
                config: effectiveConfig,
                fetchImpl,
                signal,
            });
            if (!result.ok) {
                return {
                    ok: false,
                    code: 'embedding-failed',
                    error: result.error,
                    written,
                    skipped,
                    model: result.model,
                    dim: 0,
                    mode: result.mode,
                    degraded: true,
                };
            }
            lastModel = result.model;
            lastDim = result.vectors[0]?.length ?? 0;
            mode = result.mode;
            degraded = degraded || result.degraded === true;
            if (firstError === '' && result.error !== '') firstError = result.error;
            if (lastDim === 0) {
                return { ok: false, code: 'empty-vector', error: 'embedding 返回了空向量，没有写缓存。', written, skipped };
            }
            const key = textVectorKey(lastModel, lastDim);
            for (let index = 0; index < batch.length; index += 1) {
                const vector = result.vectors[index];
                if (!vector || vector.length !== lastDim) continue;
                await writeVector('text', key, batch[index].item.id, {
                    kind: 'text',
                    model: lastModel,
                    dim: lastDim,
                    key,
                    textHash: batch[index].textHash,
                    vector: [...vector],
                    createdAt: now(),
                });
                written += 1;
            }
        }
        return {
            ok: true,
            written,
            skipped,
            model: lastModel,
            dim: lastDim,
            mode,
            degraded,
            error: firstError,
            note: remoteConfigured && localOnly
                ? '远端文本向量需要显式 allowUpload:true；这次用本地哈希向量，没有联网。'
                : degraded
                    ? '远端文本 embedding 不可用，已用本地哈希向量（只补字面召回，不是语义模型）。'
                    : '',
        };
    }

    /**
     * 选一张能发给当前多模态通道的图：原图超限时改用缩略图（同一通道统一预处理），
     * 没有缩略图就明确失败，绝不静默截断或把超大图硬发出去。
     */
    async function embeddingImageInput(item, cfg) {
        const limits = cfg.provider === 'voyage'
            ? IMAGE_EMBED_LIMITS.voyage
            : cfg.provider === 'cohere'
                ? IMAGE_EMBED_LIMITS.cohere
                : { bytes: maxBytes };
        let bytes;
        try {
            bytes = await readFile(join(root, item.original));
        } catch {
            return { ok: false, error: '原图读不到，没有生成图片向量。' };
        }
        const width = item.facts?.width;
        const height = item.facts?.height;
        const tooManyPixels = Number.isFinite(limits.pixels) && Number.isFinite(width) && Number.isFinite(height)
            && width * height > limits.pixels;
        if (bytes.length <= limits.bytes && !tooManyPixels) {
            return { ok: true, image: { bytes, mediaType: item.facts.mediaType } };
        }
        if (typeof item.thumb === 'string' && item.thumb !== '') {
            try {
                const thumb = await readFile(join(root, item.thumb));
                if (thumb.length <= limits.bytes) {
                    return { ok: true, image: { bytes: thumb, mediaType: 'image/jpeg' }, downscaled: true };
                }
            } catch {
                // 缩略图读不到就按下面的明确失败处理。
            }
        }
        const why = tooManyPixels
            ? '像素 ' + width + '×' + height + ' 超过当前通道上限 ' + limits.pixels
            : bytes.length + ' 字节超过当前通道单图上限 ' + limits.bytes;
        return {
            ok: false,
            error: '这张图' + why + '，且没有可用缩略图可降采样；没有生成图片向量。',
        };
    }

    async function ensureImageVectors(items, { force = false, allowUpload = false, signal } = {}) {
        const ready = imageEmbeddingReady(config.imageEmbedding);
        if (!ready.ok) {
            return {
                ok: false,
                code: 'unconfigured',
                error: ready.error + ' 图片向量没有生成；图文对齐检索会明确跳过，不会拿文本向量代替。',
                written: 0,
            };
        }
        if (allowUpload !== true) {
            return {
                ok: false,
                code: 'needs-confirmation',
                error: '把图片发给多模态 embedding 端点会联网（可能计费）。确认后传 allowUpload:true。',
                written: 0,
            };
        }
        const probe = await readProbe();
        if (probe === null || probe.ok !== true) {
            return {
                ok: false,
                code: 'not-probed',
                error: '多模态端点还没通过探针（文本与图片必须同维）。先跑 probe，再向量化。',
                written: 0,
            };
        }
        const cfg = config.imageEmbedding;
        const embedder = createImageEmbedder(cfg, { fetchImpl, env });
        const key = imageVectorKey(cfg, probe.dim);
        const work = [];
        for (const item of items) {
            if (!force) {
                const cached = await readVector('image', key, item.id);
                if (cached !== null && cached.dim === probe.dim && cached.vector?.length === probe.dim) continue;
            }
            work.push(item);
        }
        if (work.length === 0) {
            return { ok: true, written: 0, dim: probe.dim, provider: cfg.provider, model: cfg.model, error: '' };
        }
        const pending = [];
        let skipped = 0;
        let firstSkipReason = '';
        for (const item of work) {
            const input = await embeddingImageInput(item, cfg);
            if (!input.ok) {
                skipped += 1;
                if (firstSkipReason === '') firstSkipReason = input.error;
                continue;
            }
            pending.push({ item, image: input.image });
        }
        const skipNote = skipped > 0
            ? '有 ' + skipped + ' 张图未向量化：' + (firstSkipReason || '超过当前通道限制')
            : '';
        if (pending.length === 0) {
            return {
                ok: true,
                written: 0,
                skipped,
                dim: probe.dim,
                provider: cfg.provider,
                model: cfg.model,
                error: skipNote,
            };
        }
        let written = 0;
        const requestMaxBytes = Number.isFinite(IMAGE_EMBED_LIMITS[cfg.provider]?.requestBytes)
            ? IMAGE_EMBED_LIMITS[cfg.provider].requestBytes
            : Number.POSITIVE_INFINITY;
        let cursor = 0;
        while (cursor < pending.length) {
            const batch = [];
            let batchBytes = 0;
            while (cursor < pending.length && batch.length < cfg.batchSize) {
                const entry = pending[cursor];
                if (batch.length > 0 && batchBytes + entry.image.bytes.length > requestMaxBytes) break;
                batch.push(entry);
                batchBytes += entry.image.bytes.length;
                cursor += 1;
            }
            if (batch.length === 0) {
                cursor += 1;
                continue;
            }
            const result = await embedder.embedImages(batch.map((entry) => entry.image), { inputType: 'document', signal });
            if (!result.ok) {
                return { ok: false, code: 'embedding-failed', error: result.error, written, skipped, dim: probe.dim, provider: cfg.provider, model: cfg.model };
            }
            if (result.dim !== probe.dim) {
                return {
                    ok: false,
                    code: 'dim-mismatch',
                    error: '端点这次返回 ' + result.dim + ' 维，探针是 ' + probe.dim + ' 维；没有写缓存。',
                    written,
                    skipped,
                };
            }
            for (let index = 0; index < batch.length; index += 1) {
                const item = batch[index].item;
                const vector = result.vectors[index];
                if (!vector) continue;
                await writeVector('image', key, item.id, {
                    kind: 'image',
                    provider: cfg.provider,
                    model: cfg.model,
                    dim: probe.dim,
                    key,
                    space: imageSpaceKey(cfg),
                    vector: [...vector],
                    createdAt: now(),
                });
                written += 1;
            }
        }
        return {
            ok: true,
            written,
            skipped,
            dim: probe.dim,
            provider: cfg.provider,
            model: cfg.model,
            error: skipNote,
        };
    }

    async function ensureVectors(options = {}) {
        await loadState();
        const ids = Array.isArray(options.ids) ? options.ids.map((id) => asText(id).toLowerCase()) : [];
        const items = ids.length > 0
            ? ids.map((id) => state.items[id]).filter(Boolean)
            : Object.values(state.items);
        const report = { ok: true, items: items.length, text: null, image: null };
        if (options.text !== false) report.text = await ensureTextVectors(items, options);
        if (options.image === true) report.image = await ensureImageVectors(items, options);
        report.ok = (report.text?.ok ?? true) && (report.image?.ok ?? true);
        return report;
    }

    async function describeItem(id, options = {}) {
        const item = await get(id);
        if (item === null) return { ok: false, code: 'not-found', error: '图库里没有这张图。' };
        if (options.allowUpload !== true) {
            return {
                ok: false,
                code: 'needs-confirmation',
                error: '会把这张图发给视觉模型（可能产生云端费用）。确认后才上传：传 allowUpload:true。',
                vision: { channel: config.vision.channel, model: config.vision.model, provider: config.vision.provider },
            };
        }
        const ready = visionReady(config.vision);
        if (!ready.ok) return { ok: false, code: ready.code, error: ready.error };
        let bytes;
        try {
            bytes = await readFile(join(root, item.original));
        } catch (error) {
            return { ok: false, code: 'unreadable', error: '原图读不到：' + (error?.message ?? String(error)) };
        }
        const result = await describeImage({
            config: config.vision,
            bytes,
            mediaType: item.facts.mediaType,
            host: options.host ?? null,
            fetchImpl,
            env,
            signal: options.signal,
        });
        if (!result.ok) {
            return { ok: false, code: result.code ?? 'vision-failed', error: result.error };
        }
        const parsed = parseAestheticJson(result.text);
        if (!parsed.ok) {
            return {
                ok: false,
                code: 'bad-output',
                error: parsed.error,
                preview: parsed.preview,
                note: '模型原始输出没有写进图库，避免把不可解析的内容当描述。',
            };
        }
        const stored = await setAesthetic(item.id, parsed.value, {
            resetManualDescription: true,
            channel: result.channel,
            provider: result.provider,
            model: result.model,
            promptVersion: GALLERY_PROMPT_VERSION,
            generatedAt: now(),
        });
        return {
            ok: true,
            item: stored.item,
            aesthetic: stored.item.aesthetic,
            note: result.modalitiesKnown === false
                ? '模型没有声明是否支持图片输入（inputModalities 缺失），这次调用成功了，但能力未经声明确认。'
                : '',
        };
    }

    async function status() {
        await loadState();
        const items = Object.values(state.items);
        const probe = await readProbe();
        const ready = imageEmbeddingReady(config.imageEmbedding);
        const vision = visionReady(config.vision);
        const withThumb = items.filter((item) => item.thumb !== null).length;
        const withAesthetic = items.filter((item) => item.aesthetic !== null).length;
        const withFolder = items.filter((item) => folderOf(item) !== '').length;
        const folders = folderRows();
        return {
            ok: true,
            root,
            version: GALLERY_VERSION,
            counts: { items: items.length, withThumb, withAesthetic, withFolder, folders: folders.length },
            folders,
            thumbnail: { configured: typeof thumbnailer === 'function', lastError: lastThumbError },
            lexical: { available: true, note: '中文按二元组切词 + IDF；不依赖外部服务。' },
            textEmbedding: {
                kind: 'text',
                mode: config.textEmbedding.mode,
                remote: canUseRemoteEmbedding(config.textEmbedding),
                model: canUseRemoteEmbedding(config.textEmbedding) ? config.textEmbedding.model : LOCAL_EMBEDDING_MODEL,
                dimensions: config.textEmbedding.dimensions || 0,
                note: canUseRemoteEmbedding(config.textEmbedding)
                    ? '描述、标签与查询词会发给远端文本 embedding（可能计费）；搜索/索引仍要显式 allowUpload:true，未确认时用本地哈希向量。文本 embedding 与图文多模态是两个能力。'
                    : '当前是本地哈希向量：只补字面召回，不是语义模型。',
            },
            imageEmbedding: {
                kind: 'multimodal',
                provider: config.imageEmbedding.provider,
                model: config.imageEmbedding.model,
                configured: ready.ok,
                probed: probe?.ok === true,
                dim: probe?.ok === true ? probe.dim : 0,
                note: !ready.ok
                    ? ready.error
                    : (probe?.ok === true
                        ? '图文对齐已通过探针；文字搜图 / 以图搜图可用，但发查询文本或图片前仍需显式 allowUpload:true（可能计费）。'
                        : '端点已配置但还没通过探针：图文信号会跳过，先跑 probe。'),
            },
            vision: {
                kind: 'vision',
                channel: config.vision.channel,
                provider: config.vision.provider,
                model: config.vision.model,
                configured: vision.ok,
                note: vision.ok ? '描述仍要显式 allowUpload 才会把图片发出去（可能计费）。' : vision.error,
            },
            uploadPolicy: '默认不联网：describe / probe / 图片向量化 / 文字搜图 / 以图搜图 / 远端文本向量化都会把对应内容（图片，或查询词/描述/标签）发给远端，可能计费，均需显式 allowUpload:true；未确认时文本走本地哈希。远端文本 embedding 与图文多模态是两个独立能力。',
        };
    }

    async function search(params = {}) {
        await loadState();
        const cfg = params.config !== undefined && params.config !== null
            ? normalizeGalleryConfig(params.config)
            : config;
        const query = asText(params.query).trim();
        const imageId = asText(params.imageId ?? params.image).trim().toLowerCase();
        const allowUpload = params.allowUpload === true;
        const weights = { ...DEFAULT_WEIGHTS, ...(params.weights ?? {}) };
        const all = Object.values(state.items);
        let candidates = all.filter((item) => matchesFilters(item, params));
        let queryImage = null;
        if (imageId !== '') {
            queryImage = state.items[imageId] ?? null;
            if (queryImage === null) {
                return { ok: false, code: 'not-found', error: '作为查询的那张图不在图库里。' };
            }
            candidates = candidates.filter((item) => item.id !== queryImage.id);
        }
        const signals = [];
        const reports = {};

        if (query !== '') {
            const scores = lexicalScores(candidates, query);
            signals.push({ name: 'lexical', weight: Number(weights.lexical) || 0, scores });
            reports.lexical = { used: true, note: '词法：中文二元组 + IDF 覆盖度。' };
        } else {
            reports.lexical = { used: false, reason: '没有查询词。' };
        }

        if (query !== '') {
            const scores = new Map(candidates.map((item) => [item.id, tagScoreOf(item, query)]));
            const nonZero = [...scores.values()].some((value) => value > 0);
            signals.push({ name: 'tag', weight: Number(weights.tag) || 0, scores });
            reports.tag = { used: true, matched: nonZero, note: nonZero ? '命中标签（精确 1.0 / 包含 0.6）。' : '没有标签命中。' };
        } else {
            reports.tag = { used: false, reason: '没有查询词。' };
        }

        if (query !== '' && cfg.textEmbedding.mode !== 'off' && candidates.length > 0) {
            const remoteText = canUseRemoteEmbedding(cfg.textEmbedding);
            const textAllowed = allowUpload === true;
            const textCfg = remoteText && !textAllowed ? { ...cfg.textEmbedding, mode: 'local' } : cfg.textEmbedding;
            const embedded = await embedTexts({ texts: [query], config: textCfg, fetchImpl });
            if (!embedded.ok) {
                reports.textVector = { used: false, reason: embedded.error };
            } else {
                const vector = embedded.vectors[0];
                const dim = vector?.length ?? 0;
                const model = embedded.model;
                const key = dim > 0 ? textVectorKey(model, dim) : '';
                const scores = new Map();
                let covered = 0;
                for (const item of candidates) {
                    const cached = key === '' ? null : await readVector('text', key, item.id);
                    const textHash = documentHashOf(item);
                    if (cached !== null && cached.textHash === textHash && cached.vector?.length === dim) {
                        scores.set(item.id, Math.max(0, cosineSimilarity(vector, asVector(cached.vector))));
                        covered += 1;
                    } else {
                        scores.set(item.id, 0);
                    }
                }
                const usable = covered > 0;
                if (usable) signals.push({ name: 'textVector', weight: Number(weights.textVector) || 0, scores });
                reports.textVector = {
                    used: usable,
                    model,
                    dim,
                    mode: embedded.mode,
                    degraded: embedded.degraded === true,
                    covered,
                    total: candidates.length,
                    note: !usable
                        ? (remoteText && !textAllowed
                            ? '远端文本向量需要显式 allowUpload:true；这次用本地哈希向量，本地还没有缓存（先索引）。'
                            : '没有与当前文本向量模型匹配的缓存向量：先运行向量化；这次没有把缺失项当成 0 分参与融合。')
                        : (embedded.degraded || embedded.mode === 'local'
                            ? (remoteText && !textAllowed
                                ? '远端文本向量需要显式 allowUpload:true；这次用本地哈希向量（未联网），只补字面召回。'
                                : '描述与标签的本地哈希向量：只补字面召回，不是语义模型。')
                            : '描述与标签的文本向量（语义能力取决于所配端点）。'),
                    ...(embedded.error ? { error: embedded.error } : {}),
                    ...(remoteText && !textAllowed ? { remoteBlocked: true } : {}),
                };
            }
        } else {
            reports.textVector = { used: false, reason: cfg.textEmbedding.mode === 'off' ? '文本向量在设置里关掉了。' : '没有查询词或没有候选。' };
        }

        let probe = null;
        if (query !== '' && candidates.length > 0) {
            const ready = imageEmbeddingReady(cfg.imageEmbedding);
            probe = ready.ok ? await readProbe() : null;
            if (!ready.ok) {
                reports.imageText = { used: false, reason: ready.error + ' 没有拿文本向量去比图片向量。' };
            } else if (allowUpload !== true) {
                // 默认零请求：配置了多模态端点不等于同意把查询词发出去。
                // 只有调用方显式 allowUpload:true 才走 imageText。
                reports.imageText = {
                    used: false,
                    reason: '文字查询的多模态向量会把查询词发给图文端点（可能计费）。确认后传 allowUpload:true；这次没有发请求。',
                };
            } else if (probe === null || probe.ok !== true) {
                reports.imageText = { used: false, reason: '多模态端点还没通过探针（文本与图片同维），图文对齐信号跳过。' };
            } else {
                const embedder = createImageEmbedder(cfg.imageEmbedding, { fetchImpl, env });
                const result = await embedder.embedTexts([query], { inputType: 'query' });
                if (!result.ok) {
                    reports.imageText = { used: false, reason: '查询文本的多模态向量失败：' + result.error };
                } else if (result.dim !== probe.dim) {
                    reports.imageText = { used: false, reason: '查询文本向量 ' + result.dim + ' 维与图片空间 ' + probe.dim + ' 维不一致，已跳过。' };
                } else {
                    const key = imageVectorKey(cfg.imageEmbedding, probe.dim);
                    const scores = new Map();
                    let covered = 0;
                    for (const item of candidates) {
                        const cached = await readVector('image', key, item.id);
                        if (cached !== null && cached.dim === probe.dim && cached.vector?.length === probe.dim) {
                            scores.set(item.id, Math.max(0, cosineSimilarity(result.vectors[0], asVector(cached.vector))));
                            covered += 1;
                        } else {
                            scores.set(item.id, 0);
                        }
                    }
                    if (covered > 0) {
                        signals.push({ name: 'imageText', weight: Number(weights.imageText) || 0, scores });
                    }
                    reports.imageText = {
                        used: covered > 0,
                        provider: cfg.imageEmbedding.provider,
                        model: cfg.imageEmbedding.model,
                        dim: probe.dim,
                        covered,
                        total: candidates.length,
                        note: covered > 0
                            ? '文字查询与图片向量在同一多模态空间里比较。'
                            : '还没有图片向量：先向量化（需要显式 allowUpload）。',
                    };
                }
            }
        } else {
            reports.imageText = { used: false, reason: '没有查询词或没有候选。' };
        }

        if (queryImage !== null) {
            const ready = imageEmbeddingReady(cfg.imageEmbedding);
            probe = probe ?? (ready.ok ? await readProbe() : null);
            if (!ready.ok) {
                reports.imageImage = { used: false, reason: ready.error + ' 以图搜图不可用；没有拿文本向量代替。' };
            } else if (probe === null || probe.ok !== true) {
                reports.imageImage = { used: false, reason: '多模态端点还没通过探针，以图搜图信号跳过。' };
            } else if (allowUpload !== true) {
                reports.imageImage = { used: false, reason: '以图搜图会把查询图发给多模态端点（可能计费）。确认后传 allowUpload:true。' };
            } else if (candidates.length === 0) {
                reports.imageImage = { used: false, reason: '图库里还没有其它候选图（查询图本身不计入结果）。' };
            } else {
                const embedder = createImageEmbedder(cfg.imageEmbedding, { fetchImpl, env });
                const queryInput = await embeddingImageInput(queryImage, cfg.imageEmbedding);
                if (!queryInput.ok) {
                    reports.imageImage = { used: false, reason: queryInput.error };
                } else {
                    const result = await embedder.embedImages([queryInput.image], { inputType: 'query' });
                    if (!result.ok) {
                        reports.imageImage = { used: false, reason: '查询图的向量失败：' + result.error };
                    } else if (result.dim !== probe.dim) {
                        reports.imageImage = { used: false, reason: '查询图向量 ' + result.dim + ' 维与图片空间 ' + probe.dim + ' 维不一致，已跳过。' };
                    } else {
                        const key = imageVectorKey(cfg.imageEmbedding, probe.dim);
                        const scores = new Map();
                        let covered = 0;
                        for (const item of candidates) {
                            const cached = await readVector('image', key, item.id);
                            if (cached !== null && cached.dim === probe.dim && cached.vector?.length === probe.dim) {
                                scores.set(item.id, Math.max(0, cosineSimilarity(result.vectors[0], asVector(cached.vector))));
                                covered += 1;
                            } else {
                                scores.set(item.id, 0);
                            }
                        }
                        // imageImage 是独立信号；老配置只写 imageText 权重时沿用旧值，
                        // 避免升级后以图搜图的排序突然换一套权重。
                        const imageImageWeight = params.weights?.imageImage !== undefined
                            ? Number(params.weights.imageImage) || 0
                            : params.weights?.imageText !== undefined
                                ? Number(params.weights.imageText) || 0
                                : DEFAULT_WEIGHTS.imageImage;
                        if (covered > 0) signals.push({ name: 'imageImage', weight: imageImageWeight, scores });
                        reports.imageImage = {
                            used: covered > 0,
                            dim: probe.dim,
                            covered,
                            total: candidates.length,
                            note: covered > 0 ? '以图搜图：查询图与候选图在同一多模态空间里比较。' : '还没有候选图片向量。',
                        };
                    }
                }
            }
        } else {
            reports.imageImage = { used: false, reason: '没有提供图库内的查询图。' };
        }

        const active = signals.filter((signal) => signal.weight > 0);
        const denominator = active.reduce((sum, signal) => sum + signal.weight, 0) || 1;
        const scored = candidates.map((item) => {
            const parts = {};
            let sum = 0;
            for (const signal of active) {
                const value = signal.scores.get(item.id) ?? 0;
                parts[signal.name] = Number(value.toFixed(4));
                sum += value * signal.weight;
            }
            return { item, score: sum / denominator, parts };
        });
        scored.sort((left, right) => right.score - left.score
            || asText(right.item.facts?.addedAt).localeCompare(asText(left.item.facts?.addedAt)));
        const limit = Number.isFinite(params.limit) ? Math.min(200, Math.max(1, Math.trunc(params.limit))) : 30;
        const queryActive = query !== '' || queryImage !== null;
        // 有查询词（或查询图）时只保留真正有信号的候选。零分条目是「没命中」，
        // 不是「相似度为 0 的命中」：把它们一起返回会让客户端永远进不了空态，
        // 看起来像「搜什么都能搜到整库」。只有硬过滤（标签/分类/收藏）没有查询
        // 条件时才完整返回过滤结果。
        // 零分与负分都不是「命中」：负的图文余弦表示不相似，留着它会让一次
        // 无结果的检索看起来仍然搜到了东西。只保留正分候选；只有硬过滤
        // （标签 / 分类 / 收藏）没有查询条件时才完整返回过滤结果。
        const matched = queryActive ? scored.filter((entry) => entry.score > 0) : scored;
        return {
            ok: true,
            query,
            imageId: queryImage?.id ?? '',
            total: matched.length,
            limit,
            reason: queryActive && matched.length === 0 ? 'no-match' : '',
            signals: reports,
            activeSignals: active.map((signal) => signal.name),
            items: matched.slice(0, limit).map((entry) => summarize(entry.item, {
                score: Number(entry.score.toFixed(4)),
                parts: entry.parts,
                ...(queryActive ? {} : { score: null, parts: {} }),
            })),
            notes: buildSearchNotes({ reports, signals, queryActive, matched: matched.length }),
        };
    }

    function buildSearchNotes({ reports, signals, queryActive, matched = 0 }) {
        const notes = [];
        if (!queryActive) return ['没有查询词也没有查询图：按入库时间列最近图片。'];
        if (matched === 0) notes.push('没有匹配的图片：换个词，或先确认已建文本向量 / 图文向量。');
        if (signals.length === 0) notes.push('没有可用的检索信号。');
        if (reports.textVector?.used === false && reports.textVector?.reason !== undefined) notes.push('文本向量：' + reports.textVector.reason);
        if (reports.textVector?.remoteBlocked === true && reports.textVector?.note) notes.push('文本向量：' + reports.textVector.note);
        if (reports.imageText?.used === false && reports.imageText?.reason !== undefined) notes.push('图文对齐：' + reports.imageText.reason);
        if (reports.imageImage?.used === false && reports.imageImage?.reason !== undefined) notes.push('以图搜图：' + reports.imageImage.reason);
        notes.push('分数是各信号加权后的可解释融合；不是概率，也不代表语义能力已经可用。');
        return notes;
    }

    return {
        root,
        paths,
        getConfig: () => config,
        configure(raw) {
            config = normalizeGalleryConfig(raw ?? {});
            return config;
        },
        addFromPath,
        addBytes,
        get,
        list,
        listFolders,
        createFolder,
        renameFolder,
        removeFolder,
        updateManual,
        setAesthetic,
        remove,
        describeImage: describeItem,
        ensureVectors,
        probeImage,
        status,
        search,
        documentTextOf,
        summarize,
    };
}
