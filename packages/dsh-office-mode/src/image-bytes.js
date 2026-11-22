/**
 * 图片魔数的嗅探（PNG / JPEG / GIF / WebP），只认字节、不碰文件、不联网。
 *
 * 单独成一个模块是为了**切断循环依赖**：图库与图片资产都要用嗅探，
 * 但 image-source 还要用 web.js（取图通道），而 web.js 又依赖 config.js；
 * config.js 反过来要读 memory.js 的 DEFAULT_MEMORY_DIR。图库→image-source→web→config→memory
 * 这条链在 config 先执行时会把 memory 尚未初始化的绑定读成 TDZ 错误（第九十二轮修）。
 * 嗅探本身是纯函数，放这里谁都能直接 import，不拖任何人下水。
 *
 * @module dsh-office-mode/image-bytes
 */

/** 图片魔数嗅探：认得 PNG / JPEG / GIF / WebP，并读出像素尺寸。 */
export function sniffImageBytes(bytes) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (bytes.length >= 24 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
        return { mime: 'image/png', ext: 'png', width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
        // JPEG：扫段找 SOF0/SOF2。
        let at = 2;
        while (at + 9 < bytes.length) {
            if (bytes[at] !== 0xff) { at += 1; continue; }
            const marker = bytes[at + 1];
            const length = view.getUint16(at + 2);
            if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
                return { mime: 'image/jpeg', ext: 'jpeg', height: view.getUint16(at + 5), width: view.getUint16(at + 7) };
            }
            if (length <= 0) break;
            at += 2 + length;
        }
        return { mime: 'image/jpeg', ext: 'jpeg', width: null, height: null };
    }
    if (bytes.length >= 10 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
        return { mime: 'image/gif', ext: 'gif', width: view.getUint16(6, true), height: view.getUint16(8, true) };
    }
    if (bytes.length >= 30 && bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46
        && bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50) {
        const fourcc = String.fromCharCode(bytes[12], bytes[13], bytes[14], bytes[15]);
        // VP8X（扩展格式）：24 位宽高（存的是 宽-1 / 高-1）。
        if (fourcc === 'VP8X' && bytes.length >= 30) {
            const width = 1 + (bytes[24] | (bytes[25] << 8) | (bytes[26] << 16));
            const height = 1 + (bytes[27] | (bytes[28] << 8) | (bytes[29] << 16));
            return { mime: 'image/webp', ext: 'webp', width, height };
        }
        // 有损 WebP（VP8 ）：块头后是帧头，宽高各 14 位小端（帧起点的 26 / 28）。
        if (fourcc === 'VP8 ' && bytes.length >= 30) {
            const width = (bytes[26] | (bytes[27] << 8)) & 0x3fff;
            const height = (bytes[28] | (bytes[29] << 8)) & 0x3fff;
            return { mime: 'image/webp', ext: 'webp', width: width || null, height: height || null };
        }
        // 无损 WebP（VP8L）：0x2f 签名后 4 字节，低 14 位宽、高 14 位高（都存 值-1）。
        if (fourcc === 'VP8L' && bytes.length >= 25 && bytes[20] === 0x2f) {
            const bits = bytes[21] | (bytes[22] << 8) | (bytes[23] << 16) | (bytes[24] << 24);
            return { mime: 'image/webp', ext: 'webp', width: (bits & 0x3fff) + 1, height: ((bits >>> 14) & 0x3fff) + 1 };
        }
        return { mime: 'image/webp', ext: 'webp', width: null, height: null };
    }
    return undefined;
}
