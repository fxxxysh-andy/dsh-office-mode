/**
 * 最小 ZIP 读写（OOXML 容器）。零第三方依赖，只用 node:zlib。
 *
 * 为什么自己写：办公模式插件要在任意 profile 里零依赖安装，而 fflate/jszip
 * 在 preset 作用域下不一定可解析。ZIP 的存储/打包格式足够小，这里实现
 * 「读中央目录 + 局部头切片」和「写局部头 + 中央目录 + EOCD」两条路径。
 *
 * @module dsh-office-mode/engine/zip
 */
import { deflateRawSync, inflateRawSync } from 'node:zlib';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;

const CRC_TABLE = (() => {
    const table = new Int32Array(256);
    for (let i = 0; i < 256; i += 1) {
        let c = i;
        for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[i] = c;
    }
    return table;
})();

/** 计算 ZIP 用的 CRC-32。 */
export function crc32(bytes) {
    let crc = -1;
    for (let i = 0; i < bytes.length; i += 1) {
        crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xff];
    }
    return (crc ^ -1) >>> 0;
}

function toBytes(data) {
    // 用 ArrayBuffer.isView 而不是 instanceof：脚本在 vm 里跑，跨 realm 的
    // TypedArray 不满足宿主侧的原型判断，但同样是合法视图。
    if (typeof data === 'string') return new TextEncoder().encode(data);
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    if (data instanceof Uint8Array) return data;
    throw new TypeError('zip: 只接受字符串、Uint8Array 或 ArrayBuffer');
}

/**
 * 2007-01-01 00:00:00 的固定 DOS 时间戳。
 * 固定值让同一份内容产出同一份字节，测试可以直接比对。
 */
const DOS_TIME = 0;
const DOS_DATE = ((2007 - 1980) << 9) | (1 << 5) | 1;

/**
 * 打包成 ZIP。
 * @param {Array<{name: string, data: string|Uint8Array}>|Map<string, string|Uint8Array>} entries
 *   条目按给定顺序写入；OOXML 约定 `[Content_Types].xml` 放在最前。
 * @returns {Uint8Array}
 */
export function zip(entries) {
    const list = entries instanceof Map
        ? [...entries].map(([name, data]) => ({ name, data }))
        : [...entries];
    const encoder = new TextEncoder();
    const chunks = [];
    const central = [];
    let offset = 0;
    for (const entry of list) {
        const nameBytes = encoder.encode(entry.name);
        const raw = toBytes(entry.data);
        const crc = crc32(raw);
        const deflated = deflateRawSync(raw, { level: 9 });
        const useDeflate = deflated.length < raw.length;
        const body = useDeflate ? deflated : raw;
        const method = useDeflate ? 8 : 0;

        const local = new Uint8Array(30 + nameBytes.length);
        const lv = new DataView(local.buffer);
        lv.setUint32(0, SIG_LOCAL, true);
        lv.setUint16(4, 20, true);
        lv.setUint16(6, 0x0800, true); // UTF-8 文件名
        lv.setUint16(8, method, true);
        lv.setUint16(10, DOS_TIME, true);
        lv.setUint16(12, DOS_DATE, true);
        lv.setUint32(14, crc, true);
        lv.setUint32(18, body.length, true);
        lv.setUint32(22, raw.length, true);
        lv.setUint16(26, nameBytes.length, true);
        lv.setUint16(28, 0, true);
        local.set(nameBytes, 30);
        chunks.push(local, body);

        const cd = new Uint8Array(46 + nameBytes.length);
        const cv = new DataView(cd.buffer);
        cv.setUint32(0, SIG_CENTRAL, true);
        cv.setUint16(4, 20, true);
        cv.setUint16(6, 20, true);
        cv.setUint16(8, 0x0800, true);
        cv.setUint16(10, method, true);
        cv.setUint16(12, DOS_TIME, true);
        cv.setUint16(14, DOS_DATE, true);
        cv.setUint32(16, crc, true);
        cv.setUint32(20, body.length, true);
        cv.setUint32(24, raw.length, true);
        cv.setUint16(28, nameBytes.length, true);
        cv.setUint32(42, offset, true);
        cd.set(nameBytes, 46);
        central.push(cd);

        offset += local.length + body.length;
    }
    const centralSize = central.reduce((sum, part) => sum + part.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, SIG_EOCD, true);
    ev.setUint16(8, list.length, true);
    ev.setUint16(10, list.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    return concat([...chunks, ...central, eocd]);
}

function concat(parts) {
    const total = parts.reduce((sum, part) => sum + part.length, 0);
    const out = new Uint8Array(total);
    let at = 0;
    for (const part of parts) {
        out.set(part, at);
        at += part.length;
    }
    return out;
}

/**
 * 解开 ZIP。返回 `Map<条目名, Uint8Array>`，顺序与包内一致。
 * 只按中央目录寻址，因此包尾注释、数据描述符都不影响解析。
 */
export function unzip(bytes) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    const floor = Math.max(0, bytes.length - 0x10000 - 22);
    for (let i = bytes.length - 22; i >= floor; i -= 1) {
        if (view.getUint32(i, true) === SIG_EOCD) {
            eocd = i;
            break;
        }
    }
    if (eocd < 0) throw new Error('zip: 找不到中央目录结尾（不是有效的 ZIP/OOXML 包）');
    const count = view.getUint16(eocd + 10, true);
    let cursor = view.getUint32(eocd + 16, true);
    const decoder = new TextDecoder('utf-8');
    const files = new Map();
    for (let i = 0; i < count; i += 1) {
        if (view.getUint32(cursor, true) !== SIG_CENTRAL) throw new Error('zip: 中央目录项签名错误');
        const method = view.getUint16(cursor + 10, true);
        const compSize = view.getUint32(cursor + 20, true);
        const rawSize = view.getUint32(cursor + 24, true);
        const nameLen = view.getUint16(cursor + 28, true);
        const extraLen = view.getUint16(cursor + 30, true);
        const commentLen = view.getUint16(cursor + 32, true);
        const localAt = view.getUint32(cursor + 42, true);
        const name = decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + nameLen));
        if (view.getUint32(localAt, true) !== SIG_LOCAL) throw new Error(`zip: 局部头签名错误（${name}）`);
        const localNameLen = view.getUint16(localAt + 26, true);
        const localExtraLen = view.getUint16(localAt + 28, true);
        const start = localAt + 30 + localNameLen + localExtraLen;
        const body = bytes.subarray(start, start + compSize);
        let data;
        if (method === 0) data = body;
        else if (method === 8) data = new Uint8Array(inflateRawSync(body));
        else throw new Error(`zip: 不支持的压缩方式 ${String(method)}（${name}）`);
        if (rawSize !== data.length) data = data.subarray(0, rawSize);
        files.set(name, data);
        cursor += 46 + nameLen + extraLen + commentLen;
    }
    return files;
}

/** 把 ZIP 条目按 UTF-8 解成文本。 */
export function unzipText(bytes) {
    const decoder = new TextDecoder('utf-8');
    const out = new Map();
    for (const [name, data] of unzip(bytes)) {
        out.set(name, decoder.decode(data));
    }
    return out;
}

/**
 * 用「读 → 改 → 写」的方式编辑已有的 OOXML 包。
 * `mutate(name, text)` 返回新的文本内容；返回 `undefined` 表示不修改。
 */
export function editPackage(bytes, mutate) {
    const files = unzip(bytes);
    const decoder = new TextDecoder('utf-8');
    const encoder = new TextEncoder();
    const out = new Map();
    for (const [name, data] of files) {
        if (/\.(xml|rels)$/i.test(name)) {
            const next = mutate(name, decoder.decode(data));
            out.set(name, typeof next === 'string' ? encoder.encode(next) : data);
        } else {
            out.set(name, data);
        }
    }
    return zip(out);
}
