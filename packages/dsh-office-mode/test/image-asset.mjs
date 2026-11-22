/**
 * 知识库图片资产安全读取的测试。
 *
 * 覆盖：assets/ 前缀与 .. 拦截、realpath 根内校验（符号链接/目录联接逃逸）、
 * 只按魔数认 PNG/JPEG/GIF/WebP、扩展名不参与 MIME、大小上限、空文件与错误码。
 *
 * 跑法：node test/image-asset.mjs
 */
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readSafeAsset, normalizeAssetRelative, assetSecurityHeaders } from '../src/image-asset.js';

const results = [];
class SkipError extends Error {}
async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
    } catch (error) {
        if (error instanceof SkipError) results.push({ name, ok: false, skipped: true, reason: error.message });
        else results.push({ name, ok: false, error });
    }
}

// 1×1 透明 PNG（合法魔数 + IHDR 宽高）。
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lzq1WQAAAABJRU5ErkJggg==', 'base64');
// 最小 GIF：GIF89a + 宽高 1×1。
const GIF = Buffer.concat([Buffer.from('GIF89a', 'ascii'), Buffer.from([1, 0, 1, 0, 0, 0, 0])]);

await check('normalizeAssetRelative：只接受 assets/ 下的规范相对路径', async () => {
    assert.equal(normalizeAssetRelative('assets/a.png').ok, true);
    for (const bad of ['', 'a.png', 'assets/../x.png', 'assets/./x.png', 'assets//x.png', 'assets/a/../../b.png', '/assets/a.png', 'assets/a\0b.png']) {
        const got = normalizeAssetRelative(bad);
        assert.equal(got.ok, false, bad);
        assert.equal(got.status, 400, bad);
    }
});

await check('合法 PNG：按魔数返回 image/png，扩展名撒谎也不影响', async () => {
    const root = await mkdtemp(join(tmpdir(), 'asset-'));
    try {
        await mkdir(join(root, 'assets'));
        await writeFile(join(root, 'assets', 'looks-like.jpg'), PNG);
        const got = await readSafeAsset({ root, relativePath: 'assets/looks-like.jpg' });
        assert.equal(got.ok, true);
        assert.equal(got.mime, 'image/png');
        assert.equal(got.bytes.length, PNG.length);
        assert.equal(got.relativePath, 'assets/looks-like.jpg');
        const headers = assetSecurityHeaders(got.mime, got.size);
        assert.equal(headers['x-content-type-options'], 'nosniff');
        assert.match(headers['content-security-policy'], /sandbox/);
        assert.equal(headers['content-type'], 'image/png');
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

await check('GIF 魔数被接受', async () => {
    const root = await mkdtemp(join(tmpdir(), 'asset-'));
    try {
        await mkdir(join(root, 'assets'));
        await writeFile(join(root, 'assets', 'a.gif'), GIF);
        const got = await readSafeAsset({ root, relativePath: 'assets/a.gif' });
        assert.equal(got.ok, true);
        assert.equal(got.mime, 'image/gif');
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

await check('非图片 / SVG：415，不按扩展名当图片回', async () => {
    const root = await mkdtemp(join(tmpdir(), 'asset-'));
    try {
        await mkdir(join(root, 'assets'));
        await writeFile(join(root, 'assets', 'evil.svg'), '<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>');
        await writeFile(join(root, 'assets', 'note.png'), 'just text');
        const svg = await readSafeAsset({ root, relativePath: 'assets/evil.svg' });
        assert.equal(svg.ok, false);
        assert.equal(svg.status, 415);
        const txt = await readSafeAsset({ root, relativePath: 'assets/note.png' });
        assert.equal(txt.ok, false);
        assert.equal(txt.status, 415);
        assert.equal(txt.code, 'not-image');
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

await check('空文件 415、超限 413、缺失 404', async () => {
    const root = await mkdtemp(join(tmpdir(), 'asset-'));
    try {
        await mkdir(join(root, 'assets'));
        await writeFile(join(root, 'assets', 'empty.png'), Buffer.alloc(0));
        await writeFile(join(root, 'assets', 'big.png'), PNG);
        const empty = await readSafeAsset({ root, relativePath: 'assets/empty.png' });
        assert.equal(empty.status, 415);
        const big = await readSafeAsset({ root, relativePath: 'assets/big.png', maxBytes: 10 });
        assert.equal(big.status, 413);
        assert.match(big.error, /上限/);
        const missing = await readSafeAsset({ root, relativePath: 'assets/none.png' });
        assert.equal(missing.status, 404);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

await check('目录不是普通文件：404', async () => {
    const root = await mkdtemp(join(tmpdir(), 'asset-'));
    try {
        await mkdir(join(root, 'assets', 'folder.png'), { recursive: true });
        const got = await readSafeAsset({ root, relativePath: 'assets/folder.png' });
        assert.equal(got.status, 404);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

await check('符号链接/目录联接逃逸到根外：realpath 拒绝 403', async () => {
    const root = await mkdtemp(join(tmpdir(), 'asset-'));
    const outside = await mkdtemp(join(tmpdir(), 'asset-out-'));
    try {
        await mkdir(join(root, 'assets'));
        await writeFile(join(outside, 'secret.png'), PNG);
        const linkPath = join(root, 'assets', 'leak');
        try {
            await symlink(outside, linkPath, process.platform === 'win32' ? 'junction' : 'dir');
        } catch (error) {
            // Windows 没开开发者模式时可能无权建联接；记为 SKIP，不计入 passed。
            throw new SkipError('无法创建符号链接：' + error.code);
        }
        const got = await readSafeAsset({ root, relativePath: 'assets/leak/secret.png' });
        assert.equal(got.ok, false);
        assert.equal(got.status, 403);
        assert.equal(got.code, 'outside');
    } finally {
        await rm(root, { recursive: true, force: true });
        await rm(outside, { recursive: true, force: true });
    }
});

for (const result of results) {
    if (result.ok) console.log('ok   ' + result.name);
    else if (result.skipped) console.log('SKIP ' + result.name + '（' + result.reason + '）');
    else {
        console.log('FAIL ' + result.name);
        console.log(result.error?.stack ?? result.error);
    }
}
const failed = results.filter((result) => !result.ok && !result.skipped).length;
const passed = results.filter((result) => result.ok).length;
const skipped = results.filter((result) => result.skipped).length;
console.log('');
console.log(passed + '/' + (results.length - skipped) + ' passed' + (skipped > 0 ? '，' + skipped + ' skipped' : ''));
if (failed > 0) process.exitCode = 1;