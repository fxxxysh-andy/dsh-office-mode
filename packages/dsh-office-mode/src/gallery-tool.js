/**
 * 图库的统一工具入口（挂载到 office_memory 的一个 action，不暴露内部分层）。
 *
 * 对外只有 op：
 *   list    看最近入库
 *   search  文字 / 标签 / 以图搜图
 *   add     从会话工作目录导入一张图
 *   tag     手工标签、分类、文件夹、收藏、备注
 *   folder-list / folder-create / folder-rename / folder-remove  文件夹
 *   describe 让视觉模型生成审美描述（必须 allowUpload:true）
 *   index   准备检索向量（文本默认；图片需要 allowUpload:true）
 *   status  能力与未配置项
 *   remove  删除一张图
 *
 * 这一层不碰公共文件：挂载时只要在 tools.js 的 action 表里加一个分支，
 * 把 deps（galleryFor / hostOf）接进去即可。
 *
 * @module dsh-office-mode/gallery-tool
 */
import { join } from 'node:path';

function asText(value) {
    return typeof value === 'string' ? value : value === null || value === undefined ? '' : String(value);
}

function idOf(raw) {
    const id = asText(raw).trim().toLowerCase();
    return /^[0-9a-f]{64}$/.test(id) ? id : '';
}

function boolOf(value) {
    return value === true || value === '1' || value === 'true';
}

function shortId(id) {
    return asText(id).slice(0, 8);
}

function tagsLine(item) {
    const tags = Array.isArray(item.tags) ? item.tags : [];
    return tags.length === 0 ? '' : ' 标签：' + tags.join('、');
}

function itemLine(item, gallery) {
    const title = item.name || item.description || item.aesthetic?.summary || item.aesthetic?.subject || item.note || shortId(item.id);
    const parts = [];
    parts.push('- ' + title + '（' + shortId(item.id) + '，' + item.width + '×' + item.height + '，' + item.mediaType + '）');
    parts.push('  路径：' + join(gallery.root, 'originals', item.id + '.' + (item.mediaType === 'image/jpeg' ? 'jpeg' : item.mediaType.replace('image/', ''))));
    if (item.thumb) parts.push('  缩略图：' + join(gallery.root, item.thumb));
    if (asText(item.folder) !== '') parts.push('  文件夹：' + item.folder);
    const tags = tagsLine(item);
    if (tags !== '') parts.push(tags);
    if (item.category !== '') parts.push('  分类：' + item.category + (item.favorite ? '（收藏）' : ''));
    else if (item.favorite) parts.push('  收藏');
    return parts.join('\n');
}

function scoreLine(item) {
    if (item.score === null || item.score === undefined) return '';
    const parts = item.parts ?? {};
    const labels = { lexical: '词法', tag: '标签', textVector: '文本向量', imageText: '图文对齐', imageImage: '以图搜图' };
    const detail = Object.entries(parts)
        .filter(([, value]) => Number(value) > 0)
        .map(([key, value]) => (labels[key] ?? key) + ' ' + Number(value).toFixed(2))
        .join(' / ');
    return '  分数：' + Number(item.score).toFixed(2) + (detail === '' ? '' : '（' + detail + '）');
}

/** 把一次图库调用的结果渲染成给人看的短文本。 */
export function renderGalleryValue(value) {
    if (value === null || typeof value !== 'object') return '（图库没有返回结果）';
    if (value.ok === false && value.op !== 'index') {
        const extra = value.code === 'needs-confirmation'
            ? '\n这条动作会把图片或文本发给外部模型/端点。确认后重新调用，并显式传 allowUpload: true。'
            : '';
        const reason = asText(value.error) || asText(value.text) || '原因未知';
        const suffix = asText(value.error) !== '' && asText(value.code) !== '' ? '（' + value.code + '）' : '';
        return '图库没有完成：' + reason + suffix + extra;
    }
    if (value.op === 'status') {
        const status = value.status ?? {};
        const lines = [
            '图库：' + (status.root ?? '') + '，共 ' + (status.counts?.items ?? 0) + ' 张（有描述 ' + (status.counts?.withAesthetic ?? 0) + '，有缩略图 ' + (status.counts?.withThumb ?? 0) + '，已归类 ' + (status.counts?.withFolder ?? 0) + '；文件夹 ' + (status.counts?.folders ?? 0) + ' 个）。',
            '缩略图：' + (status.thumbnail?.configured ? '可用' : '不可用（' + (status.thumbnail?.lastError || '没有生成器') + '）'),
            '文本向量：' + (status.textEmbedding?.model ?? '') + '（' + (status.textEmbedding?.note ?? '') + '）',
            '图文对齐：' + (status.imageEmbedding?.configured
                ? (status.imageEmbedding.probed ? '可用，' + status.imageEmbedding.dim + ' 维' : '已配置但未通过探针')
                : '未配置（' + (status.imageEmbedding?.note ?? '') + '）'),
            '视觉描述：' + (status.vision?.configured ? status.vision.channel + ' / ' + status.vision.model : '未配置（' + (status.vision?.note ?? '') + '）'),
            '上传策略：' + (status.uploadPolicy ?? ''),
        ];
        return lines.join('\n');
    }
    if (value.op === 'describe') {
        const aesthetic = value.item?.aesthetic ?? {};
        const lines = ['已生成审美描述（' + (aesthetic.channel ?? '') + ' / ' + (aesthetic.model ?? '') + '）：'];
        for (const key of ['subject', 'composition', 'light', 'material', 'mood', 'summary']) {
            if (asText(aesthetic[key]) !== '') lines.push('- ' + key + '：' + aesthetic[key]);
        }
        if (Array.isArray(aesthetic.palette) && aesthetic.palette.length > 0) lines.push('- palette：' + aesthetic.palette.join('、'));
        if (Array.isArray(aesthetic.tags) && aesthetic.tags.length > 0) lines.push('- tags：' + aesthetic.tags.join('、'));
        if (Array.isArray(aesthetic.droppedFields) && aesthetic.droppedFields.length > 0) {
            lines.push('（模型给出的 ' + aesthetic.droppedFields.join('、') + ' 属于作者/版权信息，已丢弃，不写进图库。）');
        }
        if (asText(value.note) !== '') lines.push(asText(value.note));
        return lines.join('\n');
    }
    if (value.op === 'index') {
        const lines = ['图库向量：'];
        if (value.text) {
            lines.push('- 文本：' + (value.text.ok
                ? '写入 ' + value.text.written + ' 条（' + (value.text.model || '') + '，' + value.text.dim + ' 维）' + (value.text.degraded ? '，已降级' : '')
                : '没有完成：' + value.text.error));
            if (value.text.note) lines.push('  说明：' + value.text.note);
        }
        if (value.image) {
            lines.push('- 图片：' + (value.image.ok
                ? '写入 ' + value.image.written + ' 条（' + value.image.provider + ' / ' + value.image.model + '，' + value.image.dim + ' 维）'
                : '没有完成：' + value.image.error));
        }
        return lines.join('\n');
    }
    if (value.op === 'add') {
        const item = value.item;
        if (item === undefined) return asText(value.text);
        const lines = [value.duplicate ? '这张图已经在图库里（同内容哈希），没有重复保存：' : '已入库：'];
        lines.push(itemLine(item, value.gallery));
        if (value.thumb?.ok === false) lines.push('缩略图没有生成：' + value.thumb.reason);
        return lines.join('\n');
    }
    if (value.op === 'tag') {
        return '已更新：\n' + itemLine(value.item, value.gallery);
    }
    if (value.op === 'folder-create') {
        return (value.created ? '已创建文件夹：' : '文件夹已存在：') + value.path;
    }
    if (value.op === 'folder-rename') {
        return '已重命名文件夹：' + value.from + ' → ' + value.to
            + '（图片 ' + value.moved + ' 张，目录 ' + value.renamedFolders + ' 个）。';
    }
    if (value.op === 'folder-remove') {
        return '已删除文件夹：' + value.path + '（连同空的子目录共 ' + value.removedFolders + ' 个）。';
    }
    if (value.op === 'folder-list') {
        const folders = Array.isArray(value.folders) ? value.folders : [];
        if (folders.length === 0) return '图库还没有文件夹：入库或 tag 时传 folder（如 参考/风景）即可归类。';
        const lines = ['图库文件夹（直属 / 含子目录）：'];
        for (const folder of folders) {
            lines.push('- ' + folder.path + '（' + folder.direct + ' / ' + folder.total + '）');
        }
        return lines.join('\n');
    }
    if (value.op === 'remove') {
        return '已删除 ' + shortId(value.removed) + '（原图、缩略图与向量缓存一并清理）。';
    }
    const items = Array.isArray(value.items) ? value.items : [];
    if (items.length === 0) {
        return value.op === 'search'
            ? '图库里没有匹配的图片。' + (Array.isArray(value.notes) ? '\n' + value.notes.join('\n') : '')
            : '图库还是空的：用 add 从工作目录导入图片。';
    }
    const lines = [value.op === 'search'
        ? '找到 ' + items.length + ' 张（按融合分排序）：'
        : '最近 ' + items.length + ' 张：'];
    for (const item of items) {
        lines.push(itemLine(item, value.gallery) + (scoreLine(item) === '' ? '' : '\n' + scoreLine(item)));
    }
    if (Array.isArray(value.notes) && value.notes.length > 0) {
        lines.push('');
        for (const note of value.notes) lines.push('> ' + note);
    }
    return lines.join('\n');
}

/**
 * 执行一次图库统一入口调用。
 *
 * @param {object} rawArgs { op, ... }；op 缺省 list
 * @param {{galleryFor: () => object|null, hostOf?: () => object|null}} deps
 */
export async function executeGalleryAction(rawArgs, deps = {}) {
    const args = rawArgs !== null && typeof rawArgs === 'object' ? rawArgs : {};
    const op = asText(args.op) || 'list';
    const gallery = typeof deps.galleryFor === 'function' ? deps.galleryFor() : null;
    if (gallery === null) {
        return { ok: false, op, text: '图库没有初始化：请先在设置里启用办公记忆/图库，或确认数据目录可写。', code: 'not-ready' };
    }
    if (op === 'status') {
        const status = await gallery.status();
        return { ok: true, op, status, gallery: { root: gallery.root } };
    }
    if (op === 'list') {
        const listed = await gallery.list({
            tags: args.tags,
            category: args.category,
            folder: args.folder,
            favorite: args.favorite,
            limit: Number.isFinite(args.limit) ? args.limit : 50,
            offset: Number.isFinite(args.offset) ? args.offset : 0,
        });
        return { ok: true, op, ...listed, gallery: { root: gallery.root } };
    }
    if (op === 'search') {
        const found = await gallery.search({
            query: args.query,
            imageId: args.image,
            tags: args.tags,
            category: args.category,
            folder: args.folder,
            favorite: args.favorite,
            allowUpload: boolOf(args.allowUpload),
            limit: Number.isFinite(args.limit) ? args.limit : 12,
        });
        return { ...found, op, gallery: { root: gallery.root } };
    }
    if (op === 'add') {
        const added = await gallery.addFromPath(args.path, {
            name: args.name,
            description: args.description,
            tags: args.tags,
            category: args.category,
            folder: args.folder,
            favorite: args.favorite,
            note: args.note,
        });
        return { ...added, op, gallery: { root: gallery.root } };
    }
    if (op === 'tag') {
        const id = idOf(args.id);
        if (id === '') return { ok: false, op, error: '要给出图库里那张图的 id（64 位十六进制）。', code: 'bad-id' };
        const updated = await gallery.updateManual(id, {
            name: args.name,
            description: args.description,
            tags: args.tags,
            addTags: args.addTags,
            removeTags: args.removeTags,
            category: args.category,
            folder: args.folder,
            favorite: args.favorite,
            note: args.note,
        });
        return { ...updated, op, gallery: { root: gallery.root } };
    }
    if (op === 'folder-list') {
        const listed = await gallery.listFolders();
        return { ...listed, op, gallery: { root: gallery.root } };
    }
    if (op === 'folder-create') {
        const created = await gallery.createFolder(args.path ?? args.folder);
        return { ...created, op, gallery: { root: gallery.root } };
    }
    if (op === 'folder-rename') {
        const renamed = await gallery.renameFolder(args.from, args.to);
        return { ...renamed, op, gallery: { root: gallery.root } };
    }
    if (op === 'folder-remove') {
        const removed = await gallery.removeFolder(args.path ?? args.folder);
        return { ...removed, op, gallery: { root: gallery.root } };
    }
    if (op === 'describe') {
        const id = idOf(args.id);
        if (id === '') return { ok: false, op, error: '要给出图库里那张图的 id。', code: 'bad-id' };
        const described = await gallery.describeImage(id, {
            allowUpload: boolOf(args.allowUpload),
            host: typeof deps.hostOf === 'function' ? deps.hostOf() : null,
        });
        return { ...described, op, gallery: { root: gallery.root } };
    }
    if (op === 'index') {
        const indexed = await gallery.ensureVectors({
            ids: Array.isArray(args.ids) ? args.ids : null,
            text: args.text !== false,
            image: boolOf(args.image),
            force: boolOf(args.force),
            allowUpload: boolOf(args.allowUpload),
        });
        return { ...indexed, op, gallery: { root: gallery.root } };
    }
    if (op === 'remove') {
        const id = idOf(args.id);
        if (id === '') return { ok: false, op, error: '要给出图库里那张图的 id。', code: 'bad-id' };
        const removed = await gallery.remove(id);
        return { ...removed, op, gallery: { root: gallery.root } };
    }
    return {
        ok: false,
        op,
        error: '不认识的 op：' + op,
        code: 'bad-op',
        text: '可用 op：list / search / add / tag / folder-list / folder-create / folder-rename / folder-remove / describe / index / status / remove。',
    };
}
