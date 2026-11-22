/**
 * CI / 发布守门的名单核对（第五十六轮 S-06）。
 *
 * 这一份守的不是某个功能，而是**「发布时到底测了什么」这句话是不是真的**：
 *   - 两份工作流（`.github/workflows/ci.yml` / `release.yml`）都必须只调
 *     `node scripts/run-suites.mjs --ci`，不许再手写一串套件名 —— 名单只有
 *     `scripts/ci-suites.mjs` 一处真源。从前两边各写一串，于是 `format-excel` /
 *     `format-ppt` / `format-ppt-revise` / `chart` 这四个核心格式回归**漏在守门外十一轮**，
 *     而 CI 一直是绿的（绿的是「跑过的那 14 套」，不是「代码没坏」）。
 *   - CI 档必须**包含**核心格式回归，且**不含**任何要本机引擎 / 要出网 / 要宿主包的套件
 *     （两向核对：只查「该在的在」会放过误加的，只查「不该在的不在」会放过漏加的）。
 *   - CI 档里的每个名字都要在 `run-suites.mjs` 的 `ORDER` 里（那里是「文件真的存在」的
 *     校验点；`--ci` 启动时也会再核一遍，这里做文本级的早期发现）。
 *
 * 文本级检查的自白：`ORDER` 那段是按源码文本解析的（`const ORDER = [` … `];`），不是
 * 运行时对象 —— 运行时的那份由 `run-suites.mjs --ci` 自己在启动期核对（缺名字直接非零
 * 退出，见那里的纪律三）。两处互为补充：这里拦住「改了名字没改名单」，那里拦住「名单里
 * 写了不存在的套件」。
 *
 * 跑法：node test/ci-guard.mjs
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { CI_FORBIDDEN, CI_SUITES } from '../scripts/ci-suites.mjs';

const results = [];
function check(name, fn) {
    try {
        fn();
        results.push({ name, ok: true });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

const WORKFLOWS = ['ci.yml', 'release.yml'];
const workflowPath = (name) => fileURLToPath(new URL(`../../../.github/workflows/${name}`, import.meta.url));
const runnerSource = readFileSync(fileURLToPath(new URL('../scripts/run-suites.mjs', import.meta.url)), 'utf8');

/** 从一份工作流里取出所有「调 run-suites.mjs」的那一行（去掉行首缩进，排除注释）。 */
function runnerLines(text) {
    return text.split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => !line.startsWith('#') && line.includes('run-suites.mjs'));
}

/** 一行里出现几个「套件名」词（用来识别写死的名单）。 */
function suiteHits(line) {
    const names = new Set([...CI_SUITES, ...CI_FORBIDDEN]);
    let hits = 0;
    for (const name of names) {
        if (new RegExp(`(^|[^\\w-])${name.replace(/[-]/g, '\\-')}([^\\w-]|$)`).test(line)) hits += 1;
    }
    return hits;
}

/** `ORDER` 数组的源码文本（只取 `const ORDER = [` 到下一个 `];`）。 */
function orderSource() {
    const start = runnerSource.indexOf('const ORDER = [');
    assert.ok(start >= 0, 'run-suites.mjs 里找不到 ORDER 数组');
    const end = runnerSource.indexOf('];', start);
    assert.ok(end > start, 'ORDER 数组没有正常结束');
    return runnerSource.slice(start, end);
}

// ── 1. 两份工作流都只点名 `--ci` ─────────────────────────────────────────────

check('工作流只调 run-suites.mjs --ci，不再手写名单（自动建的 Initial commit 之后最容易腐烂的一处）', () => {
    for (const name of WORKFLOWS) {
        const text = readFileSync(workflowPath(name), 'utf8');
        const lines = runnerLines(text);
        assert.equal(lines.length, 1, `${name} 里应当只有一处调 run-suites.mjs，实际 ${lines.length} 处：${lines.join(' | ')}`);
        assert.equal(lines[0], 'run: node scripts/run-suites.mjs --ci',
            `${name} 的回归步骤必须逐字是「run: node scripts/run-suites.mjs --ci」，实际「${lines[0]}」`);
    }
});

check('工作流里不许再出现套件名清单（写死的名字就是第二份真源）', () => {
    for (const name of WORKFLOWS) {
        const text = readFileSync(workflowPath(name), 'utf8');
        // 只查**没被注释掉**的行：注释里点名「谁要什么环境」是解释，不是名单。
        // 判据是「一行里出现**两个以上**套件名」—— 单个词可能是巧合（release.yml 里
        // `gh release view` 的 `view` 就与套件名同形），一串名字不会是。
        const live = text.split(/\r?\n/).filter((line) => !line.trim().startsWith('#'));
        const suspects = live.filter((line) => suiteHits(line) >= 2);
        assert.deepEqual(suspects, [], `${name} 里还有套件名清单（应全部搬进 scripts/ci-suites.mjs）：${suspects.join(' | ')}`);
    }
});

// ── 2. CI 档的两向核对 ──────────────────────────────────────────────────────

check('CI 档包含核心格式回归（S-06 点名的四个 + 原本就在的两个）', () => {
    for (const name of ['chart', 'format-excel', 'format-ppt', 'format-ppt-revise', 'format-ppt-math', 'format-word']) {
        assert.ok(CI_SUITES.includes(name), `CI 档必须包含 ${name}（核心格式回归不能漏在发布守门之外）`);
    }
});

check('CI 档不含任何要本机引擎 / 要出网 / 要宿主包的套件', () => {
    const both = CI_SUITES.filter((name) => CI_FORBIDDEN.includes(name));
    assert.deepEqual(both, [], `CI 档里混进了环境相关的套件：${both.join('、')}`);
});

check('两份名单都没有重复项（重复会让「N/M 套件」的 N 虚高）', () => {
    assert.equal(new Set(CI_SUITES).size, CI_SUITES.length, 'CI 档里有重复名字');
    assert.equal(new Set(CI_FORBIDDEN).size, CI_FORBIDDEN.length, 'CI_FORBIDDEN 里有重复名字');
});

check('CI 档里的每个名字都在 ORDER 里（否则 --ci 会跑出一套「不在全量回归里」的东西）', () => {
    const order = orderSource();
    const missing = CI_SUITES.filter((name) => !order.includes(`'${name}'`));
    assert.deepEqual(missing, [], `ci-suites.mjs 里有 ORDER 里没有的套件：${missing.join('、')}`);
});

check('守门名单自己也在 CI 档里（守门不被守门就等于没守）', () => {
    assert.ok(CI_SUITES.includes('ci-guard'), 'ci-guard 必须在 CI 档里');
});

// ── 收尾 ────────────────────────────────────────────────────────────────────

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    if (item.ok) console.log(`PASS  ${item.name}`);
    else console.log(`FAIL  ${item.name}  → ${item.error?.message ?? item.error}`);
}
console.log(`\nci-guard: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
