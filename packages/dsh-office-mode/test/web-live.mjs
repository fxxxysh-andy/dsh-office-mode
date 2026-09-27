/**
 * 内置检索的联网自测（**要联网**，所以不进默认测试套件）。
 *
 * 默认测试套件一律离线：通道选择、落盘格式、HTML 收文本、SSRF 分类这些都能用
 * 假接缝钉死。但「真的查得到东西」这件事假接缝证明不了，所以单独留一个脚本：
 *
 *   node test/web-live.mjs
 *
 * 它跑的是**自带实现**（不注入 web 接缝），走 DeepSeek 的 Anthropic 兼容接口 +
 * 原生 web_search。没有可用 Key 时明确跳过（退出码 0），不是失败。
 * 需要联网与额度，所以不要把它挂进 npm test。
 */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';

import { createWebAccess, WEB_ENGINE_BUILTIN } from '../src/web.js';
import { excerptOf } from '../src/search.js';
import { resolveConfig } from '../src/config.js';
import { buildTools } from '../src/tools.js';

const root = mkdtempSync(join(tmpdir(), 'office-web-live-'));
const config = resolveConfig({ search: { builtin: { maxResults: 5, fetchPages: 1 } } });
const access = createWebAccess(undefined, config.search.builtin);

const status = await access.probe();
if (!status.ok) {
    console.log('SKIP  没有可用的联网通道（缺 API Key 或不出网）：' + status.reason);
    rmSync(root, { recursive: true, force: true });
    process.exit(0);
}
console.log('通道：' + status.reason);

let failed = 0;
async function check(name, fn) {
    try {
        await fn();
        console.log('PASS  ' + name);
    } catch (error) {
        failed += 1;
        console.log('FAIL  ' + name + '\n      ' + (error?.message ?? error));
    }
}

await check('自带检索能拿到带 URL 的来源', async () => {
    const result = await access.search('2026年地方政府债务化解 最新进展', { maxResults: 5 });
    assert.equal(result.engine, WEB_ENGINE_BUILTIN);
    assert.ok(result.sources.length > 0, '一条来源都没有');
    for (const source of result.sources) {
        assert.match(source.url, /^https?:\/\//, `来源不是 URL：${source.url}`);
    }
    console.log('      ' + result.sources.slice(0, 3).map((s) => s.title || s.url).join(' ｜ '));
});

await check('自带取正文能把真实页面收成可读文本，并挑出正文摘录', async () => {
    const page = await access.fetch('https://example.com/');
    assert.equal(page.statusCode, 200);
    assert.ok(page.content.length > 20, '正文太短，多半没取到');
    assert.ok(!page.content.includes('<html'), 'HTML 标签要收掉');
    const excerpt = excerptOf(page.content);
    assert.ok(excerpt.length > 10, '摘录为空');
    console.log('      摘录：' + excerpt.slice(0, 60));
});

await check('office_search_run 端到端：落盘 + 回清单', async () => {
    const tools = buildTools(config, undefined);
    const tool = tools.find((t) => t.name === 'office_search_run');
    const exec = { agent: { session: { header: { cwd: root } } }, signal: new AbortController().signal };
    const value = await tool.execute({ queries: ['清华大学 形势与政策'], fetchPages: 0 }, exec);
    assert.equal(value.ok, true);
    assert.ok(value.sources > 0);
    const text = readFileSync(join(root, value.file), 'utf8');
    assert.ok(text.includes('https://'), '结果文件里要有来源 URL');
    assert.ok(text.includes('## 清华大学 形势与政策'), '按查询分组');
});

rmSync(root, { recursive: true, force: true });
console.log(`web-live: ${failed === 0 ? '全部通过' : failed + ' 项失败'}`);
process.exit(failed > 0 ? 1 : 0);
