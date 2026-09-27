/**
 * 端到端测试：完全走 office_run 的真实路径（vm 沙箱 → SDK → 三个格式模块 → 复检），
 * 一次调用产出 Word / Excel / PPT 三件套 + 一份 Markdown，然后校验：
 *   - 三个文件都存在且是合法 ZIP
 *   - 复检报告里有真实 stats/outline
 *   - 缓存默认保留（跨调用复用），显式 keepCache:false 才清空
 *
 * 产物留在 <仓库根>/.office/tmp/e2e/ 供 validate-com.vbs 用真实 Office 再验一遍。
 *
 * 跑法：node test/e2e.mjs [--keep]
 */
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

import { resolveConfig } from '../src/config.js';
import { executeRun } from '../src/run.js';
import { renderRun } from '../src/tools.js';
import { unzip } from '../src/engine/zip.js';

// 测试中间产物统一落在仓库根的 .office/tmp/<suite>/ 下：与插件运行期的中间产物
// 同处一棵 .office/ 树，而不是在包里另开一个 .tmp。
const outDir = fileURLToPath(new URL('../../../.office/tmp/e2e/', import.meta.url));
rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

const SCRIPT = `
const business = 'business';

// ── Word：一份有结构的季度汇报 ──────────────────────────────────────────────
const w = office.word.create({ title: '2026 年第一季度业绩汇报', theme: business, author: '运营中心' });
w.title('2026 年第一季度业绩汇报')
 .para('编制：运营中心　　数据截止：2026 年 3 月 31 日')
 .heading('一、整体情况', 1)
 .para('本季度整体达成率 108%，超额完成目标。三条主要业务线均实现正增长，其中企业服务线增长最快。')
 .bullets(['营收 3,860 万元，同比 +18.4%', '新增客户 32 家，其中付费客户 21 家', '平均交付周期由 21 天缩短到 17 天'])
 .heading('二、分月数据', 1)
 .table({
     columns: [{ title: '月份', width: 0 }, { title: '目标（万元）', type: 'number' }, { title: '完成（万元）', type: 'number' }, { title: '达成率', type: 'percent' }],
     rows: [['1 月', 1200, 1302, '108.5%'], ['2 月', 1100, 1144, '104.0%'], ['3 月', 1300, 1414, '108.8%']],
 })
 .heading('三、下季度计划', 1)
 .steps(['把企业服务线的交付模板标准化', '六月前完成客户成功团队扩编', '试点按季度结算的付费方式'])
 .quote('把交付做成标准件，增长才不依赖人手。')
 .save('季度汇报.docx');

// ── Excel：一份带公式和合计的数据表 ────────────────────────────────────────
const x = office.excel.create({ theme: business });
const s = x.sheet('分月明细');
s.title('2026 年第一季度分月明细', { span: 5 })
 .table({
     columns: ['月份', '目标（万元）', '完成（万元）', '差额（万元）', '达成率'],
     rows: [
         ['1 月', 1200, 1302, null, '108.5%'],
         ['2 月', 1100, 1144, null, '104.0%'],
         ['3 月', 1300, 1414, null, '108.8%'],
     ],
     totalRow: true,
     freeze: true,
 });
s.note('差额 = 完成 - 目标；达成率 = 完成 / 目标。');
x.save('季度数据.xlsx');

// ── PPT：一套对外汇报的演示 ────────────────────────────────────────────────
const p = office.ppt.create({ title: '2026 Q1 业绩汇报', theme: business });
p.cover({ title: '2026 年第一季度业绩汇报', subtitle: '运营中心 · 2026 年 4 月', presenter: '运营中心' })
 .bullets({ title: '本季度三件事', items: ['整体达成率 108%', '新增客户 32 家', '交付周期缩短 4 天'] })
 .table({ title: '分月达成情况', columns: ['月份', '目标', '完成', '达成率'], rows: [['1 月', '1,200', '1,302', '108.5%'], ['2 月', '1,100', '1,144', '104.0%'], ['3 月', '1,300', '1,414', '108.8%']] })
 .quote({ text: '把交付做成标准件，增长才不依赖人手。' })
 .section({ title: '下季度计划' })
 .bullets({ title: '三件要事', items: ['交付模板标准化', '客户成功团队扩编', '试点季度结算'] })
 .closing({ title: '谢谢', subtitle: '运营中心' })
 .save('季度汇报.pptx');

// ── Markdown：一次改多处 ──────────────────────────────────────────────────
office.files.write('会议纪要.md', '# 季度复盘会\\n\\n## 结论\\n- 待定\\n\\n## 待办\\n- 待认领\\n');
const edited = office.files.edit('会议纪要.md', [['待定', '达成率 108%，超额完成'], ['待认领', '张伟：交付模板标准化'], ['不存在的段落', 'x']]);
office.assert(edited.missing.length === 1, '未命中的替换必须被报告');

office.log('三件套 + 纪要已生成');
office.cache.write('intermediate.json', { step: 'done' });
return { files: 4, unmatched: edited.missing.length };
`;

const config = resolveConfig({});
const exec = { agent: { session: { header: { cwd: outDir } } } };
const started = Date.now();
const result = await executeRun({ script: SCRIPT, purpose: '生成季度汇报三件套' }, exec, config);
const elapsed = Date.now() - started;

console.log(renderRun(result));
console.log('');

const problems = [];
if (result.ok !== true) problems.push(`脚本失败：${result.error?.message}（第 ${result.error?.line} 行）`);

const expected = ['季度汇报.docx', '季度数据.xlsx', '季度汇报.pptx'];
for (const name of expected) {
    const full = join(outDir, name);
    if (!existsSync(full)) {
        problems.push(`缺少产物：${name}`);
        continue;
    }
    const bytes = readFileSync(full);
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) problems.push(`${name} 不是 ZIP（PK 头）`);
    let parts = 0;
    try {
        parts = unzip(new Uint8Array(bytes)).size;
    } catch (error) {
        problems.push(`${name} 解压失败：${error.message}`);
    }
    if (parts < 5) problems.push(`${name} 包内部件过少（${parts}）`);
}

if (result.files.length !== 3) problems.push(`复检报告应有 3 个 Office 文件，实际 ${result.files.length}`);
for (const file of result.files) {
    if (file.stats === null || typeof file.stats !== 'object') problems.push(`${file.path} 缺少 stats`);
    if (!Array.isArray(file.outline) || file.outline.length === 0) problems.push(`${file.path} 缺少 outline`);
}
// 会议纪要.md 被 write 后又 edit，按路径去重后只算一个。
if (result.otherFiles.length !== 1) problems.push(`其他写入应为 1 个（纪要），实际 ${result.otherFiles.length}`);
// 2026-09-23 起缓存默认保留：中间产物跨调用复用（PDF 页面图那条路靠它）。
if (result.cache.clearedAfter !== false) problems.push('默认不该清空缓存');
if (result.cache.kept !== 1) problems.push(`缓存里应保留 1 个中间文件，实际 ${result.cache.kept}`);
if (!existsSync(join(outDir, '.office', 'cache', 'intermediate.json'))) problems.push('中间产物没有留在缓存里');
if (result.returned?.unmatched !== 1) problems.push('未命中替换没有正确返回');

// 三层记忆的台账：三个 Office 产物应各登记一条，非 Office 的会议纪要.md 不进台账。
if (result.memory?.logged !== 3) problems.push(`台账应自动登记 3 条，实际 ${result.memory?.logged}`);
const ledgerFile = join(outDir, '.office', 'memory', 'ledger.jsonl');
if (!existsSync(ledgerFile)) {
    problems.push('记忆目录里没有 ledger.jsonl（自动台账没写成）');
} else {
    const lines = readFileSync(ledgerFile, 'utf8').trim().split('\n').filter(Boolean);
    const officeName = [...expected].sort();
    const logged = lines.map((line) => JSON.parse(line).path).sort();
    if (logged.join(',') !== officeName.join(',')) problems.push(`台账登记的路径不对：${logged.join('、')}`);
}

// 显式 keepCache:false 时才清空（老行为仍然可用）。
const cleared = await executeRun({
    script: `return office.cache.list().filter((f) => !f.dir).length;`,
    purpose: '清空缓存',
    keepCache: false,
}, exec, config);
if (cleared.cache.clearedAfter !== true) problems.push('keepCache:false 没有清空缓存');
if (existsSync(join(outDir, '.office', 'cache'))) problems.push('.office/cache 目录在 keepCache:false 之后仍然存在');
if (cleared.returned !== 1) problems.push('清空前脚本应当看得见缓存里的中间文件');

// .office/ 是插件自己该有的东西（缓存 + 三层记忆），不是「预期外的文件」：
// 它必须**只有这一个**目录名，工作目录根上不该再冒出 .office/cache / .office/memory。
const leftovers = readdirSync(outDir).filter((name) => !expected.includes(name) && name !== '会议纪要.md' && name !== '.office');
if (leftovers.length > 0) problems.push(`工作目录有预期外的文件：${leftovers.join('、')}`);

console.log(`e2e: ${problems.length === 0 ? 'PASS' : 'FAIL'}  ${elapsed} ms`);
for (const problem of problems) console.log(`  - ${problem}`);
if (problems.length === 0) {
    const summary = result.files.map((file) => `${file.path} ${statSync(join(outDir, file.path)).size}B`).join('，');
    console.log(`  产物：${summary}`);
}
process.exit(problems.length === 0 ? 0 : 1);
