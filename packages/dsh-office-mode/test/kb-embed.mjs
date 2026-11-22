/**
 * `src/kb-embed.js` 的通道路由（第一百零四轮）。
 *
 * 设置页顶部显示的「实际通道」与服务端 hybridKbSearch 必须读同一份口径：
 * off 不因为本地开关而偷偷起进程；remote 不被本地抢占；auto 也没有远端时
 * 不能假装有向量。优先级抽成纯函数后，这里把全组合钉住。
 *
 * 跑法：`node test/kb-embed.mjs`
 */
import assert from 'node:assert/strict';
import test from 'node:test';

import { embedTexts, embeddingRouteOf } from '../src/kb-embed.js';

test('embeddingRouteOf：mode 优先，不被别的开关抢占', () => {
    assert.equal(embeddingRouteOf({ mode: 'off', localEnabled: true, remoteReady: true }), 'off',
        'off 时必须纯词法，本地/远端都不该被拉起');
    assert.equal(embeddingRouteOf({ mode: 'local', localEnabled: true, remoteReady: true }), 'local',
        'local 且本地推理可用时优先本地');
    assert.equal(embeddingRouteOf({ mode: 'local', localEnabled: false, remoteReady: true }), 'local-hash',
        '显式 local 但本地推理没开：允许零依赖哈希，不能改用远端');
    assert.equal(embeddingRouteOf({ mode: 'remote', localEnabled: true, remoteReady: true }), 'remote',
        'remote 不该被本机推理抢占');
    assert.equal(embeddingRouteOf({ mode: 'remote', localEnabled: true, remoteReady: false }), 'none',
        'remote 端点/模型不齐时退回词法，不落到哈希');
});

test('embeddingRouteOf：auto 的优先级是本地推理 > 远端 > 纯词法', () => {
    assert.equal(embeddingRouteOf({ mode: 'auto', localEnabled: true, remoteReady: true }), 'local');
    assert.equal(embeddingRouteOf({ mode: 'auto', localEnabled: false, remoteReady: true }), 'remote');
    assert.equal(embeddingRouteOf({ mode: 'auto', localEnabled: true, remoteReady: false }), 'local');
    assert.equal(embeddingRouteOf({ mode: 'auto', localEnabled: false, remoteReady: false }), 'none',
        'auto 没有本地也没有远端时是纯词法，不是本地哈希');
});

test('embeddingRouteOf：缺省与坏值按 auto 收口，不抛错', () => {
    assert.equal(embeddingRouteOf(), 'none');
    assert.equal(embeddingRouteOf({ mode: 'bogus', localEnabled: true }), 'local', '坏 mode 落回 auto');
    assert.equal(embeddingRouteOf(null), 'none', 'null 参数不抛错');
    assert.equal(embeddingRouteOf({ mode: 'AUTO', remoteReady: true }), 'remote', '大小写不匹配按坏值收口成 auto');
});

test('embedTexts：remote 失败时 strict 决定「退回词法」还是「哈希兜底」', async () => {
    const config = { mode: 'remote', baseURL: 'https://api.example.com/v1', model: 'bge-m3', dimensions: 32 };
    const fetchImpl = async () => { throw new Error('connect ECONNREFUSED'); };

    const loose = await embedTexts({ texts: ['知识库检索'], config, fetchImpl });
    assert.equal(loose.ok, true);
    assert.equal(loose.mode, 'local');
    assert.equal(loose.degraded, true);
    assert.equal(loose.vectors.length, 1, '非 strict 才允许本地哈希兜底');

    const strict = await embedTexts({ texts: ['知识库检索'], config, fetchImpl, strict: true });
    assert.equal(strict.ok, false, 'strict 时必须把失败报给调用方，交给它退回词法');
    assert.equal(strict.mode, 'remote');
    assert.equal(strict.degraded, true);
    assert.equal(strict.vectors.length, 0);
    assert.match(strict.error, /ECONNREFUSED/);
});

