/**
 * 逐套件跑测试（给本轮回归用）：按固定顺序跑，逐套件计时并记 exit code。
 *
 * 用法：node scripts/run-suites.mjs [套件名...]   （或 npm run test:suites）
 *   - 不传名字 = 跑默认顺序里的全部套件（含要联网的 `web-live` 与要无头 Chrome 的
 *     `web-verify`，它们排在最后）；
 *   - 传名字 = 只跑这些（例如 `node scripts/run-suites.mjs client memory`）。
 *
 * 为什么要它：此前每轮「全量回归」都是手敲二十几条 `node test/*.mjs`，跑漏一个也看不出来。
 * 这个入口把「全绿」变成可复跑的一句话，输出一行一套件，末尾给「N/M 套件 exit=0」。
 *
 * 两条纪律（第四十二轮补的，都是被真实缺陷逼出来的）：
 *   1. **名字必须真的对得上文件**。`ORDER` 里写 `seam`、文件却叫 `subagent-seam.mjs` 时，
 *      过滤是 `Array.includes` 精确匹配，这一套会被**静默丢掉** —— 于是「26/26 全绿」
 *      其实少跑了一套（第四十二轮实测：`subagent-seam` 4/4、`host-modules` exit=0、
 *      `validate-libreoffice` 3 通过 / 0 失败，三套都不在默认顺序里）。现在映射集中写在
 *      `fileOf` 里，缺映射就是脚本自己的 bug，不再靠猜。
 *   2. **打错的名字必须报错**。以前按名字点一个不存在的套件会走成「0/1 套件 exit=0」——
 *      退出码 0 的假绿，比跑红更糟。现在找不到脚本直接非零退出。
 *
 * 环境相关的套件分两档：
 *   - 默认顺序末尾（能跑就跑，跑不动也会如实红/绿）：`web-live`（要联网，本机直连到不了
 *     DuckDuckGo，要过须设 `DSH_OFFICE_TEST_PROXY`，见该脚本头部）、`web-verify`（要无头
 *     Chrome）；
 *   - `OPTIONAL`（**不进默认顺序**，按名字点才会跑）：`validate-libreoffice`（要本机装了
 *     LibreOffice、且 .office/tmp 下已有样例产物；没产物时它自己 exit 1）。
 */
import { spawn } from 'node:child_process';
import { readdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');

/** 跑得慢、且依赖本机环境的套件排在最后（超时也拖不住前面的结论）。 */
const ORDER = [
    'smoke', 'memory', 'kb', 'memory-recall', 'memory-probes', 'injection-budget', 'preset-check', 'capabilities',
    'schema-format', 'measure',
    'settings', 'client', 'view', 'pdf', 'python', 'archive', 'preview', 'image', 'search', 'web', 'web-proxy', 'web-live',
    'sites', 'seam', 'plain-default', 'e2e', 'av', 'av-words', 'chart',
    'format-ppt-math', 'format-tex', 'format-word', 'format-excel', 'format-ppt',
    'format-ppt-revise', 'web-verify',
];

/** 名字 → 文件名：**每个名字都必须在这里有解**（默认之外的名字也一样，按名字点要能跑）。 */
const fileOf = {
    sites: 'site-catalog.mjs',
    seam: 'subagent-seam.mjs',
    'memory-probes': 'memory-probes.mjs',
    'host-modules': 'host-modules.mjs',
    'validate-libreoffice': 'validate-libreoffice.mjs',
    'validate-chart-com': 'validate-chart-com.mjs',
};

/** 不进默认顺序（要本机环境 / 要已有样例），按名字点才跑。 */
const OPTIONAL = ['validate-libreoffice', 'validate-chart-com'];

/**
 * `test/` 下**不是套件**的文件：被别的套件 import 的 helper，与只产出样例的生成器。
 * 它们没有 check、没有断言、跑起来零输出 —— 放进 ORDER 只会虚增「N/N 全绿」的行数
 * （第四十二轮复核抓到的 P1-2：`host-modules.mjs` 就是这种）。列出来是为了让下面那条
 * 反向检查能分清「故意不跑」与「忘了加」。
 */
const HELPERS = ['host-modules.mjs', 'persona-shape.mjs', 'inspect-fills.mjs', 'make-samples.mjs', 'make-reference-deck.mjs'];

function scriptFor(name) {
    const file = fileOf[name] || `${name}.mjs`;
    return { file, path: resolve(root, 'test', file) };
}

const present = new Set(readdirSync(resolve(root, 'test')));
const missing = ORDER.filter((name) => !present.has(scriptFor(name).file));
if (missing.length > 0) {
    // 静默少跑是这套入口最大的坑：宁可在这里红，也不要在末尾给一个假的「全绿」。
    console.error(`RED  顺序里的套件找不到脚本：${missing.map((name) => `${name}（${scriptFor(name).file}）`).join('、')}`);
    process.exit(2);
}

// 反向检查（第四十二轮复核 P1-3）：只查 ORDER→文件是**单向**的，下一轮新加的
// `test/xxx.mjs` 会被静默丢掉而这里照样「全绿」。所以盘上每个 .mjs 都必须被三档之一收编。
const accounted = new Set([
    ...ORDER.map((name) => scriptFor(name).file),
    ...OPTIONAL.map((name) => scriptFor(name).file),
    ...HELPERS,
]);
const strays = [...present].filter((file) => file.endsWith('.mjs') && !accounted.has(file));
if (strays.length > 0) {
    console.error(`RED  test/ 下有既不在 ORDER、也不在 OPTIONAL / HELPERS 里的脚本：${strays.join('、')}`
        + ' —— 它们不会被任何一次「全量回归」跑到。要么进 ORDER（是套件），要么进 OPTIONAL / HELPERS 并写清理由。');
    process.exit(2);
}

const wanted = process.argv.slice(2);
for (const name of wanted) {
    if (!present.has(scriptFor(name).file)) {
        const hint = OPTIONAL.includes(name) ? '' : `；可选套件是 ${OPTIONAL.join('、')}`;
        console.error(`RED  找不到套件「${name}」（期望文件 test/${scriptFor(name).file}）${hint}`);
        process.exit(2);
    }
}
const names = wanted.length > 0 ? wanted : ORDER;

function run(name) {
    return new Promise((done) => {
        const started = Date.now();
        const child = spawn(process.execPath, [scriptFor(name).path], {
            cwd: root,
            stdio: ['ignore', 'pipe', 'pipe'],
        });
        let out = '';
        child.stdout.on('data', (chunk) => { out += chunk; });
        child.stderr.on('data', (chunk) => { out += chunk; });
        child.on('close', (code) => {
            const seconds = ((Date.now() - started) / 1000).toFixed(1);
            // 摘要行：套件的输出样式不统一（`xx: N/N 通过`、`PASS plain-default: …`、
            // `SKIP …`、`断言 N 项`、`exit=0` 无输出），所以几种都要认。
            // 认不出来会留下空摘要 —— 别把空摘要读成「没跑」（复核 P2-8）。
            const summary = out.trim().split(/\r?\n/)
                .filter((line) => /通过|通过数|fail|FAIL|失败|PASS|SKIP|断言|项/.test(line))
                .slice(-1)[0] || '';
            console.log(`${code === 0 ? 'OK  ' : 'RED '} ${name.padEnd(20)} exit=${code} ${seconds}s  ${summary.slice(0, 110)}`);
            if (code !== 0) {
                const tail = out.trim().split(/\r?\n/).slice(-18).join('\n');
                console.log(tail.split('\n').map((line) => `       | ${line}`).join('\n'));
            }
            done(code === 0);
        });
    });
}

let failed = 0;
for (const name of names) {
    // eslint-disable-next-line no-await-in-loop
    const ok = await run(name);
    if (!ok) failed += 1;
}
console.log(`\n汇总：${names.length - failed}/${names.length} 套件 exit=0`
    + (wanted.length === 0 && OPTIONAL.length > 0 ? `（另有可选套件未跑：${OPTIONAL.join('、')}）` : ''));
process.exit(failed === 0 ? 0 : 1);
