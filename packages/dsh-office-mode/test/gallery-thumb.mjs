/**
 * 缩略图真实路径测试（本机 Pillow）。
 *
 * 这个测试会真的调 python + PIL 生成缩略图；没有 Pillow 时跳过并明确说明，
 * 而不是拿原图冒充缩略图。
 *
 * 跑法：node test/gallery-thumb.mjs
 */
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { deflateSync } from 'node:zlib';
import { createGallery, createPillowThumbnailer } from '../src/gallery.js';
import { sniffImageBytes } from '../src/image-source.js';

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
        let c = n;
        for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

function crc32(buffer) {
    let crc = 0xffffffff;
    for (const byte of buffer) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
    const length = Buffer.alloc(4);
    length.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([length, body, crc]);
}

function pngOf(width, height) {
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;
    const stride = width * 3 + 1;
    const raw = Buffer.alloc(stride * height);
    for (let y = 0; y < height; y += 1) {
        raw[y * stride] = 0;
        for (let x = 0; x < width; x += 1) {
            const at = y * stride + 1 + x * 3;
            raw[at] = 40 + (x * 7) % 200;
            raw[at + 1] = 90 + (y * 5) % 150;
            raw[at + 2] = 160;
        }
    }
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw)),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

const root = await mkdtemp(join(tmpdir(), 'dsh-gallery-thumb-'));
const source = join(root, 'big.png');
await writeFile(source, pngOf(1200, 800));

const thumbnailer = createPillowThumbnailer();
const probe = await thumbnailer({ sourcePath: source, targetPath: join(root, 'probe.jpg'), maxSide: 48 });
if (probe.ok !== true) {
    console.log('skip：本机 Pillow 不可用，没有跑真实缩略图路径');
    console.log('     ' + probe.reason);
    await rm(root, { recursive: true, force: true });
    process.exitCode = 0;
} else {
    const gallery = createGallery({ root: join(root, 'gallery'), cwd: root, roots: [root], thumbnailer });
    const added = await gallery.addFromPath(source);
    assert.equal(added.ok, true, JSON.stringify(added));
    assert.equal(added.thumb, undefined, '有真缩略图时不应返回 thumb 失败原因');
    assert.match(added.item.thumb, /^thumbs\/[0-9a-f]{64}\.jpg$/);

    const thumbBytes = await readFile(join(gallery.root, added.item.thumb));
    assert.ok(thumbBytes.length > 0, '缩略图不能是空文件');
    const sniffed = sniffImageBytes(thumbBytes);
    assert.equal(sniffed?.mime, 'image/jpeg', '缩略图必须是真 JPEG');
    assert.ok((sniffed.width ?? 0) <= 480 && (sniffed.height ?? 0) <= 480, '缩略图长边不超过 480');

    const status = await gallery.status();
    assert.equal(status.counts.withThumb, 1);
    console.log('ok   真实 Pillow 缩略图：' + sniffed.width + '×' + sniffed.height + '，' + thumbBytes.length + ' 字节');
    console.log('1/1 passed');
    await rm(root, { recursive: true, force: true });
}
