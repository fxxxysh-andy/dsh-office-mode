/** 图库挂载胶水的测试：根目录判定、信任根、同步实例。跑法：node test/gallery-mount.mjs */
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { homedir, tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createGalleryFor, galleryRootFor } from '../src/gallery-mount.js';

const results = [];
async function check(name, fn) {
    try { await fn(); results.push({ name, ok: true }); }
    catch (error) { results.push({ name, ok: false, error }); }
}

await check('根目录：显式 galleryDir 优先，否则与 kb 同级，空 kb 返回空', async () => {
    assert.equal(galleryRootFor({ kbRoot: join('tmp', 'data', 'kb'), galleryDir: join('tmp', 'pics') }), resolve(homedir(), 'tmp/pics'));
    assert.equal(galleryRootFor({ kbRoot: join('tmp', 'data', 'kb') }), resolve(join('tmp', 'data', 'gallery')));
    assert.equal(galleryRootFor({}), '');
});

await check('createGalleryFor：同步返回实例，信任根含 cwd/kb/图库，且能入库检索', async () => {
    const base = await mkdtemp(join(tmpdir(), 'gm-'));
    const kb = join(base, 'kb');
    const { mkdir } = await import('node:fs/promises');
    await mkdir(kb, { recursive: true });
    try {
        const gallery = createGalleryFor({ root: base, kbRoot: kb, config: {} });
        assert.equal(gallery.root, resolve(join(base, 'gallery')));
        const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lzq1WQAAAABJRU5ErkJggg==', 'base64');
        const file = join(base, 'one.png');
        await writeFile(file, png);
        const added = await gallery.addFromPath(file, { tags: ['测试'] });
        assert.equal(added.ok, true);
        assert.equal(added.item.tags.includes('测试'), true);
        // 信任根外的文件必须拒绝（根 = base/kb/gallery，不包含系统临时目录的其它位置）。
        const outside = await mkdtemp(join(tmpdir(), 'gm-out-'));
        try {
            const foreign = join(outside, 'x.png');
            await writeFile(foreign, png);
            const bad = await gallery.addFromPath(foreign);
            assert.equal(bad.ok, false);
            assert.equal(bad.code, 'outside-root');
        } finally {
            await rm(outside, { recursive: true, force: true });
        }
    } finally {
        await rm(base, { recursive: true, force: true });
    }
});

for (const r of results) {
    if (r.ok) console.log('ok   ' + r.name);
    else { console.log('FAIL ' + r.name); console.log(r.error?.stack ?? r.error); }
}
const failed = results.filter((r) => !r.ok).length;
console.log('');
console.log((results.length - failed) + '/' + results.length + ' passed');
if (failed > 0) process.exitCode = 1;