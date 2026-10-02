/**
 * 配图来源与获取（office.image，历史遗留 `1-2` + `11-1`）的测试。
 *
 * 两条账合流：`11-1`（网络取图与图库检索）与 `1-2`（配图无来源）。
 * 分三层：
 *   1. **图库解析**：两条免 Key 通道的响应映射成统一候选（标题 / 尺寸 / 许可 / 作者 /
 *      一行 credit / 完整 attribution），—— 用注入的 fetchBytes 打桩，不联网。
 *   2. **取图与署名**：魔数校验（HTML 说明页要明确报「不是图片」）、落盘路径、
 *      引用进 env.citations、以及「把 credit 传给文档的 source 之后产物里真的有来源」。
 *   3. **诚实边界**：空 query / 未知通道 / 没结果 / 已存在目标 / 不可嵌入格式（WebP）。
 *
 * 真联网那一段按项目惯例：直连出不去就 SKIP 并如实说明（要过代理设 DSH_OFFICE_TEST_PROXY）。
 *
 * 跑法：node test/image.mjs
 */
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import { deflateSync } from 'node:zlib';

import { createEnv } from '../src/engine/kit.js';
import { unzipText } from '../src/engine/zip.js';
import { imageChannels, imageFetch, imageSearch, sniffImageBytes, stripHtml } from '../src/image-source.js';
import { installProxy } from '../src/web-proxy.js';
import { resolveConfig } from '../src/config.js';
import { executeRun } from '../src/run.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        if (error?.skip === true) results.push({ name, ok: true, note: `SKIP ${error.message}` });
        else results.push({ name, ok: false, error });
    }
}
function skip(message) {
    const error = new Error(message);
    error.skip = true;
    throw error;
}

const root = mkdtempSync(join(tmpdir(), 'office-image-'));
const env = createEnv({ root });

/** 造一张真 PNG（1x1，手写 IHDR/IDAT/IEND）。 */
function pngBytes(width = 1, height = 1) {
    const chunk = (type, data) => {
        const length = Buffer.alloc(4);
        length.writeUInt32BE(data.length, 0);
        const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
        const crc = Buffer.alloc(4);
        crc.writeUInt32BE(0, 0); // CRC 不在我们的嗅探范围内，写 0 足够
        return Buffer.concat([length, body, crc]);
    };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8; ihdr[9] = 2; // 8 bit / truecolor
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(Buffer.alloc(width * height * 3))),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}
const PNG = pngBytes(4, 3);
/** 一张真 PNG 落在工作目录里：脚本里没法构造字节，所以由测试先写好。 */
writeFileSync(join(root, 'pic.png'), PNG);

/** 打桩的字节通道：按 URL 里的关键字回不同的响应，并记下每次调用的选项。 */
function stubFetch(routes) {
    const calls = [];
    const fetchBytes = async (url, options) => {
        calls.push({ url: String(url), options: options ?? {} });
        for (const [needle, reply] of Object.entries(routes)) {
            if (String(url).includes(needle)) {
                if (typeof reply === 'function') return reply(String(url), options);
                return { url: String(url), statusCode: 200, contentType: 'application/json', bytes: Buffer.from(JSON.stringify(reply), 'utf8') };
            }
        }
        throw new Error(`stub 没有为 ${url} 配响应`);
    };
    return { calls, fetchBytes };
}

const COMMONS_PAYLOAD = {
    query: {
        pages: [
            {
                title: 'File:Wind turbine.jpg',
                imageinfo: [{
                    url: 'https://upload.wikimedia.org/wikipedia/commons/a/ab/Wind_turbine.jpg',
                    thumburl: 'https://upload.wikimedia.org/wikipedia/commons/thumb/a/ab/Wind_turbine.jpg/1024px-Wind_turbine.jpg',
                    descriptionurl: 'https://commons.wikimedia.org/wiki/File:Wind_turbine.jpg',
                    width: 4000,
                    height: 3000,
                    mime: 'image/jpeg',
                    extmetadata: {
                        LicenseShortName: { value: 'CC BY-SA 4.0' },
                        LicenseUrl: { value: 'https://creativecommons.org/licenses/by-sa/4.0/' },
                        Artist: { value: '<a href="//commons.wikimedia.org/wiki/User:Someone">Someone</a>' },
                    },
                }],
            },
            {
                title: 'File:Logo.svg',
                imageinfo: [{
                    url: 'https://upload.wikimedia.org/wikipedia/commons/b/bc/Logo.svg',
                    thumburl: 'https://upload.wikimedia.org/wikipedia/commons/thumb/b/bc/Logo.svg/1024px-Logo.svg.png',
                    descriptionurl: 'https://commons.wikimedia.org/wiki/File:Logo.svg',
                    width: 512, height: 512, mime: 'image/svg+xml',
                    extmetadata: { LicenseShortName: { value: 'Public domain' } },
                }],
            },
        ],
    },
};

const OPENVERSE_PAYLOAD = {
    results: [
        {
            title: 'Solar panels',
            url: 'https://live.staticflickr.com/1/2_solar.jpg',
            thumbnail: 'https://api.openverse.org/v1/images/2/thumb/',
            foreign_landing_url: 'https://www.flickr.com/photos/x/2',
            width: 1600, height: 1200,
            license: 'by', license_version: '4.0',
            creator: 'Jane Doe',
        },
    ],
};

// ── 1. 通道与解析 ──────────────────────────────────────────────────────────

await check('通道清单：两条免 Key 图库，含说明与 id', async () => {
    const channels = imageChannels();
    assert.deepEqual(channels.map((item) => item.id), ['commons', 'openverse']);
    assert.ok(channels.every((item) => item.needsKey === false), '默认通道都不该要 Key');
    assert.ok(channels.every((item) => typeof item.note === 'string' && item.note !== ''));
    return channels.map((item) => item.id).join(' / ');
});

await check('image.search(commons)：映射出标题 / 尺寸 / 许可 / 作者，并给出 credit 与 attribution', async () => {
    const stub = stubFetch({ 'commons.wikimedia.org/w/api.php': COMMONS_PAYLOAD });
    const report = await imageSearch('风力发电机', { limit: 5 }, { fetchBytes: stub.fetchBytes });
    assert.equal(report.channel, 'commons');
    assert.equal(report.count, 2);
    const first = report.results[0];
    assert.equal(first.index, 1);
    assert.equal(first.title, 'Wind turbine.jpg');
    assert.match(first.thumbUrl, /1024px/);
    assert.equal(first.pageUrl, 'https://commons.wikimedia.org/wiki/File:Wind_turbine.jpg');
    assert.equal(first.license, 'CC BY-SA 4.0');
    assert.equal(first.licenseUrl, 'https://creativecommons.org/licenses/by-sa/4.0/');
    assert.equal(first.author, 'Someone', 'Artist 里的 HTML 要去掉');
    assert.equal(first.mime, 'image/jpeg');
    assert.equal(first.fetchUrl, first.url, '原图就是位图时取原图');
    assert.equal(first.insertable, true);
    assert.match(first.credit, /维基共享资源/);
    assert.match(first.credit, /Someone/);
    assert.match(first.credit, /CC BY-SA 4.0/);
    assert.match(first.attribution, /来源页：https:\/\/commons\.wikimedia\.org/);
    // SVG 的缩略图是 PNG → fetchUrl 指向缩略图（原图取回来嵌不进去）
    const svg = report.results[1];
    assert.equal(svg.insertable, true, 'SVG 的 PNG 缩略图可以嵌');
    assert.match(svg.thumbUrl, /\.png$/);
    assert.equal(svg.fetchUrl, svg.thumbUrl, 'SVG 原图不是位图，fetchUrl 要指向 PNG 缩略图');
    assert.equal(svg.mime, 'image/svg+xml', 'mime 报的是原图格式（如实）');
    // 请求本身要带检索词与 limit（别把参数丢了）
    assert.match(stub.calls[0].url, /gsrsearch=/);
    assert.match(stub.calls[0].url, /gsrlimit=5/);
    // 选项要真的传到字节通道（maxBytes / accept / 超时）—— 只传 URL 的实现会在这里露馅
    assert.equal(stub.calls[0].options.accept, 'application/json');
    assert.ok(stub.calls[0].options.maxBytes > 0);
    assert.ok(stub.calls[0].options.fetchTimeoutMs > 0);
    return `${report.count} 条 / ${first.credit}`;
});

await check('image.search(openverse)：映射 creator 与 license', async () => {
    const stub = stubFetch({ 'api.openverse.org': OPENVERSE_PAYLOAD });
    const report = await imageSearch('solar', { channel: 'openverse', limit: 3 }, { fetchBytes: stub.fetchBytes });
    assert.equal(report.channel, 'openverse');
    const first = report.results[0];
    assert.equal(first.author, 'Jane Doe');
    assert.equal(first.license, 'BY 4.0');
    assert.equal(first.pageUrl, 'https://www.flickr.com/photos/x/2');
    assert.match(first.credit, /Openverse/);
    return first.credit;
});

await check('image.search：auto 在第一条通道失败时换下一条', async () => {
    const stub = stubFetch({
        'commons.wikimedia.org': () => { throw new Error('HTTP 503'); },
        'api.openverse.org': OPENVERSE_PAYLOAD,
    });
    const report = await imageSearch('solar', {}, { fetchBytes: stub.fetchBytes });
    assert.equal(report.channel, 'openverse', 'commons 挂了应当退到 openverse');
    assert.equal(stub.calls.length, 2);
    return 'commons → openverse';
});

await check('image.search：接口在 200 里回错误信封时转述原因，不说成「没有结果」', async () => {
    // 维基共享资源：iiurlwidth 超过原图宽度时回 {"error":{"code":"badvalue",…}}
    const commonsError = stubFetch({
        'commons.wikimedia.org': { error: { code: 'badvalue', info: 'Invalid value for iiurlwidth' } },
    });
    await assert.rejects(
        () => imageSearch('x', { channel: 'commons' }, { fetchBytes: commonsError.fetchBytes }),
        /接口报错（badvalue）：Invalid value for iiurlwidth/,
    );
    // Openverse：{"detail":"…"}
    const openverseError = stubFetch({ 'api.openverse.org': { detail: 'Request was throttled.' } });
    await assert.rejects(
        () => imageSearch('x', { channel: 'openverse' }, { fetchBytes: openverseError.fetchBytes }),
        /接口报错：Request was throttled/,
    );
    return 'badvalue / throttled';
});

await check('image.search：空 query / 未知通道 / 没有结果都给可执行的错', async () => {
    await assert.rejects(() => imageSearch('   ', {}, { fetchBytes: async () => { throw new Error('不该被调用'); } }), /query 不能为空/);
    await assert.rejects(
        () => imageSearch('x', { channel: 'getty' }, { fetchBytes: async () => ({ bytes: Buffer.from('{}'), contentType: 'application/json', url: 'u' }) }),
        /没有这条图库通道「getty」.*commons \/ openverse/,
    );
    const empty = stubFetch({ 'commons.wikimedia.org': { query: { pages: [] } }, 'api.openverse.org': { results: [] } });
    await assert.rejects(() => imageSearch('zzz', {}, { fetchBytes: empty.fetchBytes }), /没有取到图/);
    // 接口返回 HTML（拦截页）时要报「不是 JSON」，不是崩在 JSON.parse 上
    const html = async () => ({ url: 'u', statusCode: 200, contentType: 'text/html', bytes: Buffer.from('<html>blocked</html>', 'utf8') });
    await assert.rejects(() => imageSearch('x', { channel: 'commons' }, { fetchBytes: html }), /不是 JSON/);
    return '3 类错误';
});

// ── 2. 取图、署名与产物 ────────────────────────────────────────────────────

await check('image.fetch：下载 → 魔数校验 → 落 assets/ → 返回 credit 与 source', async () => {
    const stub = stubFetch({
        'upload.wikimedia.org': () => ({ url: 'https://upload.wikimedia.org/x.jpg', statusCode: 200, contentType: 'image/png', bytes: PNG }),
    });
    const candidate = {
        title: '风力发电机',
        url: 'https://upload.wikimedia.org/x.jpg',
        pageUrl: 'https://commons.wikimedia.org/wiki/File:X',
        license: 'CC BY-SA 4.0',
        author: 'Someone',
        channel: 'commons',
    };
    const fetched = await imageFetch(candidate, {}, env, { fetchBytes: stub.fetchBytes });
    assert.equal(fetched.ok, true);
    assert.equal(fetched.mime, 'image/png', '魔数说了算，不看 URL 后缀');
    assert.equal(fetched.width, 4);
    assert.equal(fetched.height, 3);
    assert.match(fetched.path, /^assets\/风力发电机-[0-9a-f]{6}\.png$/);
    assert.ok(existsSync(join(root, fetched.path)));
    assert.equal(readFileSync(join(root, fetched.path)).length, fetched.bytes);
    assert.equal(fetched.insertable, true);
    assert.match(fetched.credit, /CC BY-SA 4.0/);
    assert.equal(fetched.source.pageUrl, 'https://commons.wikimedia.org/wiki/File:X');
    // 引用要进 env.citations（office_run 会把它们写进台账的 source）
    assert.equal(env.citations.length, 1);
    assert.equal(env.citations[0].url, 'https://commons.wikimedia.org/wiki/File:X');
    assert.equal(env.citations[0].license, 'CC BY-SA 4.0');
    return fetched.path;
});

await check('image.fetch：返回 HTML 说明页时明确报「不是图片」（不当成下载成功）', async () => {
    const stub = stubFetch({
        'example.com': () => ({ url: 'https://example.com/x.jpg', statusCode: 200, contentType: 'text/html', bytes: Buffer.from('<!doctype html><html>404 not found</html>', 'utf8') }),
    });
    await assert.rejects(() => imageFetch('https://example.com/x.jpg', {}, env, { fetchBytes: stub.fetchBytes }), /不是图片.*HTML/s);
    const binary = stubFetch({
        'example.com': () => ({ url: 'https://example.com/x.bin', statusCode: 200, contentType: 'application/octet-stream', bytes: Buffer.from([0, 1, 2, 3, 4, 5]) }),
    });
    await assert.rejects(() => imageFetch('https://example.com/x.bin', {}, env, { fetchBytes: binary.fetchBytes }), /不是图片/);
    return 'HTML / 二进制都挡住';
});

await check('image.fetch：WebP 能取回来但标记为不可嵌入（Word 只吃 PNG/JPEG）', async () => {
    const webp = Buffer.concat([
        Buffer.from('RIFF', 'latin1'), Buffer.alloc(4), Buffer.from('WEBP', 'latin1'),
        Buffer.from('VP8X', 'latin1'), Buffer.alloc(4),
        // flags(1) + reserved(3) + width-1(3) + height-1(3) = 10 字节的 VP8X 数据
        Buffer.from([0, 0, 0, 0, 3, 0, 0, 1, 0, 0]),
    ]);
    const stub = stubFetch({ 'example.com': () => ({ url: 'https://example.com/a.webp', statusCode: 200, contentType: 'image/webp', bytes: webp }) });
    const fetched = await imageFetch('https://example.com/a.webp', { to: 'assets/a.webp' }, env, { fetchBytes: stub.fetchBytes });
    assert.equal(fetched.mime, 'image/webp');
    assert.equal(fetched.width, 4);
    assert.equal(fetched.height, 2);
    assert.equal(fetched.insertable, false);
    assert.equal(fetched.insertableInWord, false);
    assert.equal(fetched.insertableInPpt, false);
    assert.match(fetched.hint, /Word 只吃 PNG\/JPEG/);
    return `${fetched.width}x${fetched.height} webp`;
});

await check('image.fetch：GIF 分宿主报可嵌入性（PPT 行、Word 不行）', async () => {
    const gif = Buffer.concat([
        Buffer.from('GIF89a', 'latin1'),
        Buffer.from([0x04, 0x00, 0x03, 0x00]),  // 宽 4、高 3（小端）
        Buffer.alloc(20),
    ]);
    const stub = stubFetch({ 'example.com': () => ({ url: 'https://example.com/a.gif', statusCode: 200, contentType: 'image/gif', bytes: gif }) });
    const fetched = await imageFetch('https://example.com/a.gif', { to: 'assets/a.gif' }, env, { fetchBytes: stub.fetchBytes });
    assert.equal(fetched.mime, 'image/gif');
    // Word 的 builder.image 会跳过 GIF（word.js 只认 PNG/JPEG），PPT 能放 —— 一个布尔量说不清
    assert.equal(fetched.insertableInPpt, true);
    assert.equal(fetched.insertableInWord, false);
    assert.equal(fetched.insertable, false);
    assert.match(fetched.hint, /能嵌进 PPT/);
    assert.match(fetched.hint, /Word 只吃 PNG\/JPEG/);
    return 'ppt 行 / word 不行';
});

await check('image.fetch：相对路径不许写出工作目录', async () => {
    const stub = stubFetch({ 'example.com': () => ({ url: 'https://example.com/a.png', statusCode: 200, contentType: 'image/png', bytes: PNG }) });
    await assert.rejects(
        () => imageFetch('https://example.com/a.png', { to: '../逃出去.png' }, env, { fetchBytes: stub.fetchBytes }),
        /工作目录之外/,
    );
    return '拦下 ../';
});

await check('image.fetch：目标已存在时默认不覆盖（要覆盖得显式说）', async () => {
    const stub = stubFetch({ 'example.com': () => ({ url: 'https://example.com/a.png', statusCode: 200, contentType: 'image/png', bytes: PNG }) });
    await imageFetch('https://example.com/a.png', { to: 'assets/fixed.png' }, env, { fetchBytes: stub.fetchBytes });
    await assert.rejects(() => imageFetch('https://example.com/a.png', { to: 'assets/fixed.png' }, env, { fetchBytes: stub.fetchBytes }), /已存在/);
    const again = await imageFetch('https://example.com/a.png', { to: 'assets/fixed.png', overwrite: true }, env, { fetchBytes: stub.fetchBytes });
    assert.equal(again.bytes, PNG.length);
    return '不覆盖 / 显式覆盖';
});

await check('sniffImageBytes / stripHtml：单元级边界', async () => {
    assert.equal(sniffImageBytes(PNG).mime, 'image/png');
    assert.equal(sniffImageBytes(Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10])).mime, 'image/jpeg');
    assert.equal(sniffImageBytes(Buffer.from([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 2, 0, 3, 0])).width, 2);
    assert.equal(sniffImageBytes(Buffer.from('hello world')), undefined);
    assert.equal(stripHtml('<a href="x">A &amp; B</a>'), 'A & B');
    return '4 种输入';
});

// ── 3. 经 office_run：署名真的写进产物 ──────────────────────────────────────

const runScript = (script, extra = {}) => executeRun({ script, purpose: '配图测试' }, { agent: { session: { header: { cwd: root } } } }, resolveConfig(extra));

await check('SDK 面：office.image.channels 在脚本里可用（两条免 Key 通道）', async () => {
    const result = await runScript('return office.image.channels();');
    assert.equal(result.ok, true, result.error?.message);
    assert.deepEqual(result.returned.map((item) => item.id), ['commons', 'openverse']);
    assert.ok(result.returned.every((item) => item.needsKey === false));
    return '2 条通道';
});

await check('字节通道 httpFetchBytes：限长、同源跳转、SSRF 与超时分类都走同一套', async () => {
    // 这是 office.image 唯一的网络出口，所以它自己的护栏要有测试（以前一条都没有）。
    const { httpFetchBytes } = await import('../src/web.js');
    const { WebAccessError } = await import('../src/web-errors.js');
    const png = PNG;
    const network = (reply) => ({
        lookup: async () => [{ address: '93.184.216.34', family: 4 }],
        fetch: async (url, init) => (typeof reply === 'function' ? reply(url, init) : reply),
    });
    const imageReply = () => new Response(png, { status: 200, headers: { 'content-type': 'image/png' } });

    // 正常取回
    const ok = await httpFetchBytes('https://example.com/a.png', {}, undefined, network(imageReply));
    assert.equal(ok.bytes.length, png.length);
    assert.equal(ok.contentType, 'image/png');

    // 限长：超过 maxBytes 直接报错（截断的图片不能当成功）
    await assert.rejects(
        () => httpFetchBytes('https://example.com/a.png', { maxBytes: 8 }, undefined, network(imageReply)),
        (error) => error instanceof WebAccessError && /超过 1 KB 上限|超过/.test(error.message),
    );

    // 同源跳转跟随、跨站跳转拒绝
    let hops = 0;
    const redirects = network(async (url) => {
        hops += 1;
        if (hops === 1) return new Response(null, { status: 302, headers: { location: '/b.png' } });
        return imageReply();
    });
    const followed = await httpFetchBytes('https://example.com/a.png', {}, undefined, redirects);
    assert.equal(followed.url, 'https://example.com/b.png');
    const cross = network(() => new Response(null, { status: 302, headers: { location: 'https://evil.test/x.png' } }));
    await assert.rejects(() => httpFetchBytes('https://example.com/a.png', {}, undefined, cross), /跨站跳转/);

    // SSRF：解析到私网地址就拒绝（不发给 fetch）
    const privateHost = { lookup: async () => [{ address: '127.0.0.1', family: 4 }], fetch: async () => { throw new Error('不该走到 fetch'); } };
    await assert.rejects(() => httpFetchBytes('https://internal.test/a.png', {}, undefined, privateHost), /非公网地址/);

    // 非 http(s) 与空响应体
    await assert.rejects(() => httpFetchBytes('ftp://example.com/a.png', {}, undefined, network(imageReply)), /只支持 http \/ https/);
    await assert.rejects(
        () => httpFetchBytes('https://example.com/empty.png', {}, undefined, network(() => new Response('', { status: 200, headers: { 'content-type': 'image/png' } }))),
        /响应体是空的/,
    );
    return '6 类护栏';
});

await check('产物里真的有来源：Word 图注下多一行「来源：…」，PPT 并进题注', async () => {
    const result = await runScript(`
        const wb = office.word.create({ path: '配图.docx', title: '配图' });
        wb.para('下面这张图来自图库，署名必须跟着它。');
        wb.image('pic.png', { widthCm: 8, caption: '陆上风电场', source: '维基共享资源 · Someone · CC BY-SA 4.0' });
        wb.save();
        const deck = office.ppt.create({ path: '配图.pptx', title: '配图' });
        deck.cover({ title: '配图' });
        deck.image({ path: 'pic.png', caption: '陆上风电场', source: '维基共享资源 · Someone · CC BY-SA 4.0' });
        deck.save();
        return true;
    `);
    assert.equal(result.ok, true, result.error?.message);
    const docx = unzipText(readFileSync(join(root, '配图.docx')));
    const documentXml = docx.get('word/document.xml');
    assert.match(documentXml, /来源：维基共享资源 · Someone · CC BY-SA 4\.0/);
    assert.match(documentXml, /descr="[^"]*来源：维基共享资源/, '署名也要进替代文本');
    assert.match(documentXml, /陆上风电场/);
    assert.ok(docx.has('word/media/image1.png'), '图片部件要真的在包里');
    const pptx = unzipText(readFileSync(join(root, '配图.pptx')));
    const slideXml = [...pptx.entries()].filter(([name]) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).map(([, text]) => text).join('');
    assert.match(slideXml, /陆上风电场 ｜ 来源：维基共享资源 · Someone · CC BY-SA 4\.0/);
    return 'docx + pptx 都带来源';
});

await check('台账来源：取图的引用（来源页 + 许可）会进 office_run 台账的 source', async () => {
    // 走真实链路：把 globalThis.fetch 打桩成「返回一张 PNG 的 200」，
    // office.image.fetch → env.cite → office_run 的 sourcePaths → 台账 source。
    const original = globalThis.fetch;
    const fetchedUrls = [];
    globalThis.fetch = async (url) => {
        fetchedUrls.push(String(url));
        return new Response(PNG, { status: 200, headers: { 'content-type': 'image/png' } });
    };
    let result;
    try {
        result = await runScript(`
            const img = await office.image.fetch('https://example.com/turbine.png', { to: 'assets/turbine.png' });
            const wb = office.word.create({ path: '带引用.docx', title: '带引用' });
            wb.image(img.path, { widthCm: 6, caption: '风机', source: img.credit });
            wb.save();
            return { path: img.path, credit: img.credit };
        `);
    } finally {
        globalThis.fetch = original;
    }
    if (!result.ok && /解析不了主机|没有解析到任何地址/.test(String(result.error?.message ?? ''))) {
        skip('本机 DNS 解析不了 example.com（离线环境），这条负例测不了');
    }
    assert.equal(result.ok, true, result.error?.message);
    assert.equal(result.returned.path, 'assets/turbine.png');
    assert.ok(existsSync(join(root, 'assets/turbine.png')));
    assert.ok(fetchedUrls.some((url) => url.includes('example.com/turbine.png')), '没真的走到 fetch');

    const { createMemory } = await import('../src/memory.js');
    const memory = createMemory({ root, memory: resolveConfig({}).memory });
    // 台账行的 source 不进 read 的投影，但它是**检索键的一部分**：用图片地址当关键词
    // 能命中这一行，就说明「这份产物依据了哪张图」真的记下来了。
    const hit = await memory.read({ layer: 'ledger', query: 'turbine.png' });
    const hitItems = hit.ledger?.items ?? [];
    assert.ok(hitItems.some((item) => String(item.path).includes('带引用.docx')),
        `用图片地址检索不到这条台账：${JSON.stringify(hitItems).slice(0, 200)}`);
    // 再看一眼真源（ledger.jsonl）里的 source 字段本身。
    const ledgerFile = join(root, '.office', 'memory', 'ledger.jsonl');
    const raw = readFileSync(ledgerFile, 'utf8');
    const row = raw.split('\n').map((line) => { try { return JSON.parse(line); } catch { return undefined; } })
        .filter(Boolean).find((item) => String(item.path ?? '').includes('带引用.docx'));
    assert.ok(row !== undefined, 'ledger.jsonl 里没有这一行');
    assert.deepEqual(row.source, ['https://example.com/turbine.png'], `source 字段不对：${JSON.stringify(row.source)}`);
    return row.source.join(' ');
});

// ── 4. 真联网（按项目惯例：直连出不去就 SKIP） ──────────────────────────────

await check('真联网：经代理检索并取回一张真图（没配代理或出不去就 SKIP）', async () => {
    const proxy = process.env.DSH_OFFICE_TEST_PROXY ?? '';
    if (proxy === '') skip('没有设 DSH_OFFICE_TEST_PROXY（本机直连出不去）');
    // 单进程跑测试时拿不到宿主装配期的 undici 解析位置，所以这里补一个：
    // 从 profile / 宿主安装位置解析 undici，代理才会走**私有分派器**那条路
    // （进程级那条要 NODE_USE_ENV_PROXY=1，本进程起不到作用）。
    const { hostAnchors } = await import('./host-modules.mjs');
    const resolveModule = (specifier) => {
        for (const anchor of hostAnchors()) {
            try {
                return createRequire(pathToFileURL(join(anchor, 'package.json')).href)(specifier);
            } catch {
                // 换下一个锚点
            }
        }
        return undefined;
    };
    const installed = installProxy(proxy, { resolveModule });
    if (!installed.ok || !installed.applied) skip(`代理装不上：${installed.reason ?? ''}`);
    let report;
    try {
        report = await imageSearch('wind turbine', { channel: 'commons', limit: 3, width: 640 });
    } catch (error) {
        if (/NETWORK|TIMEOUT|取资源失败/.test(String(error?.message ?? ''))) skip(`出不去：${error.message.slice(0, 80)}`);
        throw error;
    }
    assert.ok(report.count > 0, '真联网至少要拿到一条候选');
    const first = report.results.find((item) => item.insertable) ?? report.results[0];
    const fetched = await imageFetch(first, { to: 'assets/real.png', maxBytes: 4 * 1024 * 1024 }, env);
    assert.equal(fetched.ok, true);
    assert.ok(fetched.bytes > 1000, `取回来的图太小：${fetched.bytes}`);
    assert.ok(['image/png', 'image/jpeg', 'image/gif'].includes(fetched.mime), fetched.mime);
    assert.match(fetched.credit, /维基共享资源/);
    return `${fetched.mime} ${fetched.width}x${fetched.height} ${fetched.bytes}B · ${fetched.credit}`;
});

await check('office_help：image 话题在（默认层给签名，全文层给示例）', async () => {
    const { buildHelp } = await import('../src/docs.js');
    const brief = await buildHelp('image');
    for (const needle of ['office.image.channels', 'office.image.search', 'office.image.fetch', 'credit', 'insertable']) {
        assert.ok(brief.text.includes(needle), `image 默认层缺 ${needle}`);
    }
    const full = await buildHelp('image', { detail: true });
    assert.match(full.text, /source: img\.credit/);
    assert.match(full.text, /维基共享资源/);
    const index = await buildHelp('');
    assert.match(index.text, /image：配图的来源与获取/);
    return '默认层 + 全文层 + 索引';
});

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    const note = item.note === undefined ? '' : `  （${item.note}）`;
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${note}${item.ok ? '' : `  → ${item.error?.message}`}`);
}
if (failed.length > 0) for (const item of failed) console.error(item.error);
console.log(`image: ${results.length - failed.length}/${results.length} 通过`);
rmSync(root, { recursive: true, force: true });
process.exit(failed.length === 0 ? 0 : 1);
