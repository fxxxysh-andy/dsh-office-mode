/**
 * 图库的挂载胶水（不碰公共文件）。
 *
 * memory.js / tools.js / view.js 只需要各加一两行：
 *   const gallery = createGalleryFor({ root, kbRoot, config: config.gallery, galleryDir: memory.galleryDir });
 * 根目录判定、信任根、缩略图器都在这里统一，避免三处各写一份。
 *
 * @module dsh-office-mode/gallery-mount
 */
import { dirname, join } from 'node:path';
import { createGallery, createPillowThumbnailer, galleryRootDir } from './gallery.js';

/**
 * 图库根：显式 galleryDir 优先；否则与知识库同级。
 * @param {{kbRoot?:string, galleryDir?:string}} options
 */
export function galleryRootFor({ kbRoot = '', galleryDir = '' } = {}) {
    const explicit = galleryRootDir({ galleryDir }, kbRoot);
    if (explicit !== '') return explicit;
    return kbRoot === '' ? '' : join(dirname(kbRoot), 'gallery');
}

/**
 * 建一个图库实例。createGallery 本身是同步构造，只有它的方法才是异步的，
 * 所以 memory.gallery() / tools 可以直接同步拿到实例。
 *
 * @param {{root:string, kbRoot?:string, cwd?:string, config?:object, galleryDir?:string,
 *   thumbnailer?:object|null}} options
 */
export function createGalleryFor({ root, kbRoot = '', cwd = '', config = {}, galleryDir = '' } = {}) {
    const base = typeof root === 'string' ? root : '';
    const dir = galleryRootFor({ kbRoot, galleryDir });
    const roots = [base, kbRoot, dir].filter((value) => typeof value === 'string' && value !== '');
    return createGallery({
        root: dir,
        cwd: cwd === '' ? base : cwd,
        roots,
        thumbnailer: createPillowThumbnailer(),
        config: config ?? {},
    });
}
