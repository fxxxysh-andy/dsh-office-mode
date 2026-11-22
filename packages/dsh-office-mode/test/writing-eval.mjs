#!/usr/bin/env node
/**
 * 大模型写作 A/B 评测（第七十三轮）自测。
 *
 * 钉住五件事：
 *   1. 纯函数：§5 反模板清单 10 条 0/1、±10% 限长判定、评审 JSON 解析；
 *   2. --mock 完全离线跑通，产物结构（samples/key/report）与匿名化正确，且可复现；
 *   3. 无 Key 时 SKIP + exit 0，不写任何假结果；
 *   4. 真 HTTP 路径：本地 127.0.0.1 mock server 模拟 OpenAI 兼容端点，验证两臂提示词
 *      不同（planned 带写作规划）、换序 position、限长标记、评审打分与 judgeError 降级；
 *   5. 参数错误 exit 2、生成失败 exit 1。
 *
 * 跑法：node test/writing-eval.mjs（加 --keep 保留 .office/runtime/tmp/writing-eval/）。
 * 本套件不联网：mock server 只监听 127.0.0.1 的随机端口。
 */
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { createServer } from 'node:http';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
    WRITING_TASKS,
    antiTemplateChecklist,
    lengthVerdict,
    parseJudgeOutput,
    runWritingEval,
} from '../scripts/writing-eval.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..');
const SCRIPT = resolve(ROOT, 'scripts', 'writing-eval.mjs');
const TMP = resolve(ROOT, '..', '..', '.office', 'runtime', 'tmp', 'writing-eval');
const KEEP = process.argv.includes('--keep');

const failures = [];
let checks = 0;

function check(condition, label, detail = '') {
    checks += 1;
    if (!condition) failures.push(detail ? label + ' —— ' + detail : label);
}

function checkEqual(actual, expected, label) {
    check(actual === expected, label, '期望 ' + JSON.stringify(expected) + '，实际 ' + JSON.stringify(actual));
}

function checkIncludes(haystack, needle, label) {
    const hit = Array.isArray(haystack) ? haystack.includes(needle) : String(haystack).includes(needle);
    check(hit, label, '没有找到 ' + JSON.stringify(needle));
}

function runCli(args, env = process.env) {
    return new Promise((done) => {
        const child = spawn(process.execPath, [SCRIPT, ...args], {
            cwd: ROOT,
            env,
            stdio: ['ignore', 'pipe', 'pipe'],
            windowsHide: true,
        });
        let stdout = '';
        let stderr = '';
        const timer = setTimeout(() => {
            try { child.kill('SIGKILL'); } catch { /* ignore */ }
        }, 120000);
        child.stdout.on('data', (chunk) => { stdout += chunk; });
        child.stderr.on('data', (chunk) => { stderr += chunk; });
        child.on('close', (code) => {
            clearTimeout(timer);
            done({ code, stdout, stderr });
        });
    });
}

function readKey(outDir) {
    return JSON.parse(readFileSync(join(outDir, 'key.json'), 'utf8'));
}

function sampleFiles(outDir) {
    const dir = join(outDir, 'samples');
    if (!existsSync(dir)) return [];
    return readdirSync(dir).filter((name) => name.endsWith('.md')).sort();
}

/** 本地 OpenAI 兼容假端点：只监听 127.0.0.1 随机端口，不联网。 */
function startMockServer(reply) {
    return new Promise((done) => {
        const requests = [];
        const server = createServer((req, res) => {
            let raw = '';
            req.on('data', (chunk) => { raw += chunk; });
            req.on('end', () => {
                let body = null;
                try { body = JSON.parse(raw); } catch { body = null; }
                requests.push({ method: req.method, url: req.url, headers: req.headers, body });
                const content = reply(body);
                res.writeHead(200, { 'content-type': 'application/json' });
                res.end(JSON.stringify({
                    choices: [{ message: { role: 'assistant', content } }],
                    usage: { prompt_tokens: 10, completion_tokens: 10 },
                }));
            });
        });
        server.listen(0, '127.0.0.1', () => {
            const { port } = server.address();
            done({
                port,
                requests,
                close: () => new Promise((closed) => server.close(closed)),
            });
        });
    });
}

const PLANNED_MARK = '规划路径标记Q7';
const BASELINE_MARK = '普通路径标记M4';
const LONG_TAIL = '先把结论放在前面，再把证据按强度排开，最后说明这条结论在什么范围内成立。'
    + '数据来自一次线上对照，口径写在附录里，不要把它当成因果。';
const JUDGE_REPLY = JSON.stringify({ naturalness: 5, planning: 4, credibility: 3, comment: '本地 mock 评审' });

function generationReply(body) {
    const planned = (body?.messages ?? []).some((message) => String(message?.content ?? '').includes('## 写作规划'));
    if (planned) return PLANNED_MARK + '。这一版按规划写，重点放在机制与边界，长度刻意压短。';
    return BASELINE_MARK + '。' + LONG_TAIL.repeat(12);
}

async function main() {
    rmSync(TMP, { recursive: true, force: true });
    mkdirSync(TMP, { recursive: true });

    // ---------------------------------------------------------------- 1. 纯函数
    {
        checkEqual(WRITING_TASKS.length, 3, '研究材料 §4 有三道题');
        checkEqual(WRITING_TASKS.map((task) => task.id).join(','), 'a,b,c', '任务编号 a/b/c');
        checkEqual(WRITING_TASKS.map((task) => task.targetChars).join(','), '550,475,350', '三题目标字数与 §4 一致');
        check(WRITING_TASKS.every((task) => typeof task.brief === 'string' && task.brief.length > 100), '每道题都有完整任务要求');
        check(WRITING_TASKS.every((task) => Array.isArray(task.sections) && task.sections.length >= 3), '每道题都有分节规划');

        const cleanText = '问题不在模型，而在我们怎么问。先把候选压到五条，再看正确率是否回升。'
            + '如果回升，噪声假设就站得住；如果不动，问题更可能出在排序。'
            + '但这只是假设，还没有证据。这个判断有前提：单轮问答、同语言、答案落在单一片段里。'
            + '多跳场景会把这些因素搅在一起，实验分不开两者。';
        const clean = antiTemplateChecklist(cleanText);
        const template = antiTemplateChecklist('首先，我们要提升效率。其次，我们要优化流程。最后，我们要加强协同。'
            + '综上所述，这项工作意义重大。总之，我们要继续努力。'
            + '这不仅是技术问题，而且是管理问题。这不仅影响当下，而且决定未来。');
        checkEqual(clean.items.length, 10, '反模板清单是 10 条');
        check(clean.items.every((item) => item.hit === 0 || item.hit === 1), '每条命中只有 0/1');
        check(clean.items.every((item) => typeof item.label === 'string' && item.label.length > 0), '每条都有可读标签');
        check(clean.hits <= 2, '正常文本不应大面积命中', '实际 ' + clean.hits);
        check(template.hits > clean.hits, '模板文本命中数应明显更多', '模板 ' + template.hits + ' / 正常 ' + clean.hits);
        checkEqual(template.items.find((item) => item.id === 1).hit, 1, '「首先/其次/最后」为唯一结构要命中');
        checkEqual(template.items.find((item) => item.id === 2).hit, 1, '每节结尾总结要命中');
        checkEqual(template.items.find((item) => item.id === 10).hit, 0, '模板文本没有数字，第 10 条不命中');

        const exact = lengthVerdict(550, 550);
        check(exact.truncated === false && exact.short === false, '正好达标不算超限/偏短');
        checkEqual(exact.minChars, 495, '±10% 下限');
        checkEqual(exact.maxChars, 605, '±10% 上限');
        checkEqual(lengthVerdict(700, 550).truncated, true, '超出上限标 truncated');
        checkEqual(lengthVerdict(300, 550).short, true, '低于下限标 short');

        const judged = parseJudgeOutput('评分如下：{"naturalness":5,"planning":4,"credibility":3,"comment":"边界清楚"}');
        checkEqual(judged.naturalness, 5, '评审 JSON 能从解释文字里解析');
        checkEqual(judged.comment, '边界清楚', '评审 comment 保留');
        let missing = '';
        try { parseJudgeOutput('{"naturalness":5}'); } catch (error) { missing = String(error.message); }
        checkIncludes(missing, 'planning', '缺字段要报出字段名');
        let range = '';
        try { parseJudgeOutput('{"naturalness":9,"planning":3,"credibility":3}'); } catch (error) { range = String(error.message); }
        checkIncludes(range, '1–5', '越界分数要报出范围');
    }

    // ---------------------------------------------------------------- 2. --mock 离线全流程
    const mockDir = join(TMP, 'mock');
    const mockDir2 = join(TMP, 'mock-repeat');
    {
        const run = await runCli(['--tasks', 'a', '--runs', '2', '--mock', '--out', mockDir]);
        checkEqual(run.code, 0, '--mock 应 exit 0');
        checkIncludes(run.stdout, 'mock', 'stdout 标明 mock 模式');
        checkEqual(sampleFiles(mockDir).length, 4, '1 任务 × 2 臂 × 2 轮 = 4 份样本');
        const files = sampleFiles(mockDir);
        const bodies = files.map((name) => readFileSync(join(mockDir, 'samples', name), 'utf8'));
        check(bodies.every((body) => body.startsWith('<!-- sample ')), '每份样本只以编号注释开头');
        check(bodies.every((body) => !/baseline|planned|mock|deepseek/iu.test(body)), '样本正文不泄露臂名/模型名');
        check(bodies.every((body) => body.replace(/^<!-- sample [0-9]+ -->\s*/u, '').trim().length > 200), '样本正文非空');

        const key = readKey(mockDir);
        checkEqual(key.version, 1, 'key.json 版本');
        checkEqual(key.mock, true, 'key.json 标明 mock');
        checkEqual(key.model, 'mock', 'mock 模式模型名固定');
        checkEqual(key.runs, 2, 'key.json 记录轮次');
        checkEqual(key.tasks.join(','), 'a', 'key.json 记录任务');
        checkEqual(key.arms.join(','), 'baseline,planned', 'key.json 记录两个臂');
        checkEqual(key.samples.length, 4, 'key.json 记录 4 份样本');
        check(key.samples.every((sample) => typeof sample.id === 'string' && /^[0-9]{4}$/u.test(sample.id)), '样本编号是四位数字');
        checkEqual(new Set(key.samples.map((sample) => sample.id)).size, 4, '样本编号不重复');
        check(key.samples.every((sample) => Number.isInteger(sample.score) && typeof sample.band === 'string'
            && Array.isArray(sample.warnings) && typeof sample.antiTemplateHits === 'number'), 'key 记录 score/band/warnings/反模板');
        check(key.samples.every((sample) => Object.keys(sample.antiTemplate).length === 10), '每份样本记满 10 条反模板');
        check(key.samples.every((sample) => sample.judge === null && sample.judgeError === null), '不开评审时不写 judge');
        check(key.samples.every((sample) => sample.model === 'mock'), 'key 里每份样本模型一致');
        check(key.samples.every((sample) => files.includes(sample.id + '.md')), 'key 里每个编号都有对应样本文件');
        checkEqual(key.samples.filter((sample) => sample.position === 1).length, 2, '两轮各有一份样本在左侧');

        const report = readFileSync(join(mockDir, 'report.md'), 'utf8');
        checkIncludes(report, '反模板', 'report 含反模板清单');
        checkIncludes(report, '长度', 'report 含长度统计');
        checkIncludes(report, 'planned', 'report 标明两个臂');
        checkIncludes(report, 'mock', 'report 标明 mock 模式');
        checkIncludes(report, '盲评', 'report 含盲评指引');
        checkIncludes(report, '位置控制', 'report 说明换序控制');

        const again = await runCli(['--tasks', 'a', '--runs', '2', '--mock', '--out', mockDir2]);
        checkEqual(again.code, 0, 'mock 重跑仍 exit 0');
        const keyAgain = readKey(mockDir2);
        checkEqual(
            keyAgain.samples.map((sample) => sample.id + ':' + sample.score).join(','),
            key.samples.map((sample) => sample.id + ':' + sample.score).join(','),
            '同一输入两次 mock 的编号与分数可复现',
        );
    }

    // ---------------------------------------------------------------- 3. 无 Key → SKIP，exit 0
    {
        const cleanEnv = { ...process.env };
        delete cleanEnv.WRITING_EVAL_API_KEY;
        delete cleanEnv.DEEPSEEK_API_KEY;
        const skipDir = join(TMP, 'skip');
        const run = await runCli(['--tasks', 'a', '--out', skipDir], cleanEnv);
        checkEqual(run.code, 0, '无 Key 应 exit 0（SKIP 不是失败）');
        checkIncludes(run.stdout, 'SKIP', '无 Key 打印 SKIP');
        checkIncludes(run.stdout, 'WRITING_EVAL_API_KEY', '提示 WRITING_EVAL_API_KEY');
        checkIncludes(run.stdout, 'DEEPSEEK_API_KEY', '提示 DEEPSEEK_API_KEY');
        check(!existsSync(join(skipDir, 'samples')), '无 Key 不写 samples');
        check(!existsSync(join(skipDir, 'key.json')), '无 Key 不写 key.json');
        check(!existsSync(join(skipDir, 'report.md')), '无 Key 不写 report.md');
    }

    // ---------------------------------------------------------------- 4. 参数错误 exit 2
    {
        const badRuns = await runCli(['--runs', '0', '--mock', '--out', join(TMP, 'bad-runs')]);
        checkEqual(badRuns.code, 2, '--runs 0 是参数错误');
        checkIncludes(badRuns.stderr, '参数错误', '参数错误要有明确提示');
        const badTask = await runCli(['--tasks', 'z', '--mock', '--out', join(TMP, 'bad-task')]);
        checkEqual(badTask.code, 2, '未知任务是参数错误');
        const badUrl = await runCli(['--base-url', 'ftp://example.com', '--mock', '--out', join(TMP, 'bad-url')]);
        checkEqual(badUrl.code, 2, '非 http(s) base-url 是参数错误');
        const badTokens = await runCli(['--max-tokens', '10', '--mock', '--out', join(TMP, 'bad-tokens')]);
        checkEqual(badTokens.code, 2, '越界 max-tokens 是参数错误');
        const help = await runCli(['--help']);
        checkEqual(help.code, 0, '--help exit 0');
        checkIncludes(help.stdout, '--self-test', 'help 列出 --self-test');
    }

    // ---------------------------------------------------------------- 5. 真 HTTP 路径（本地假端点）
    const server = await startMockServer((body) => (body?.model === 'writing-judge-model' ? JUDGE_REPLY : generationReply(body)));
    const httpDir = join(TMP, 'http');
    let httpKey;
    try {
        const env = { ...process.env, WRITING_EVAL_TEST_KEY: 'writing-test-key' };
        const before = server.requests.length;
        const run = await runCli([
            '--tasks', 'a', '--runs', '2',
            '--out', httpDir,
            '--model', 'writing-test-model',
            '--base-url', 'http://127.0.0.1:' + server.port,
            '--api-key-env', 'WRITING_EVAL_TEST_KEY',
        ], env);
        checkEqual(run.code, 0, '本地 mock 端点路径 exit 0');
        const requests = server.requests.slice(before);
        checkEqual(requests.length, 4, '1 任务 × 2 臂 × 2 轮 = 4 次生成请求');
        check(requests.every((request) => request.method === 'POST' && request.url === '/chat/completions'), '都打到 POST /chat/completions');
        check(requests.every((request) => request.headers.authorization === 'Bearer writing-test-key'), 'Authorization 带 --api-key-env 的 Key');
        check(requests.every((request) => request.body.model === 'writing-test-model'), '请求体带 --model');
        check(requests.every((request) => request.body.stream === false), '请求体 stream=false');
        const plannedRequests = requests.filter((request) => JSON.stringify(request.body.messages).includes('## 写作规划'));
        checkEqual(plannedRequests.length, 2, 'planned 臂的提示词带写作规划');
        check(plannedRequests.every((request) => JSON.stringify(request.body.messages).includes('召回率') || JSON.stringify(request.body.messages).includes('任务')), 'planned 提示词带任务要求');
        checkEqual(requests.length - plannedRequests.length, 2, 'baseline 臂不带写作规划');
        check(requests.filter((request) => !JSON.stringify(request.body.messages).includes('## 写作规划'))
            .every((request) => JSON.stringify(request.body.messages).includes('## 写作任务')), 'baseline 只给任务与篇幅');

        httpKey = readKey(httpDir);
        checkEqual(httpKey.mock, false, '真路径 key 不标 mock');
        checkEqual(httpKey.model, 'writing-test-model', 'key 记录生成模型');
        checkEqual(httpKey.baseUrl, 'http://127.0.0.1:' + server.port, 'key 记录 baseUrl');
        checkEqual(httpKey.keyEnv, 'WRITING_EVAL_TEST_KEY', 'key 记录 Key 来源变量名');
        check(!readFileSync(join(httpDir, 'key.json'), 'utf8').includes('writing-test-key'), 'key.json 不落 Key 明文');
        check(!readFileSync(join(httpDir, 'report.md'), 'utf8').includes('writing-test-key'), 'report 不落 Key 明文');
        checkEqual(httpKey.samples.length, 4, '真路径也写 4 份样本');
        checkEqual(sampleFiles(httpDir).length, 4, '真路径 samples 目录 4 份');

        for (const sample of httpKey.samples) {
            const body = readFileSync(join(httpDir, 'samples', sample.id + '.md'), 'utf8');
            const expected = sample.arm === 'planned' ? PLANNED_MARK : BASELINE_MARK;
            checkIncludes(body, expected, '样本 ' + sample.id + ' 的正文与 key 里的臂一致');
            check(!/baseline|planned/iu.test(body), '样本 ' + sample.id + ' 不出现臂名');
        }
        const baselineSamples = httpKey.samples.filter((sample) => sample.arm === 'baseline');
        const plannedSamples = httpKey.samples.filter((sample) => sample.arm === 'planned');
        check(baselineSamples.every((sample) => sample.truncated === true), '超长 baseline 样本被标 truncated');
        check(plannedSamples.every((sample) => sample.short === true), '过短 planned 样本被标 short');
        checkEqual(new Set(baselineSamples.map((sample) => sample.position)).size, 2, 'baseline 两轮左右位置各一次');
        checkEqual(new Set(plannedSamples.map((sample) => sample.position)).size, 2, 'planned 两轮左右位置各一次');
        check(baselineSamples.every((sample) => sample.truncated === true && sample.chars > sample.maxChars), 'truncated 与 chars 一致');

        const report = readFileSync(join(httpDir, 'report.md'), 'utf8');
        checkIncludes(report, 'writing-test-model', 'report 记录生成模型');
        checkIncludes(report, '位置控制', 'report 说明位置控制');
        checkIncludes(report, '长度与限长', 'report 含长度与限长表');
        checkIncludes(report, 'truncated', 'report 解释 truncated 标记');

        // ------------------------------------------------------------ 6. 评审模型路径
        const judgeDir = join(TMP, 'judge');
        const beforeJudge = server.requests.length;
        const judged = await runCli([
            '--tasks', 'a', '--runs', '2',
            '--out', judgeDir,
            '--model', 'writing-test-model',
            '--judge', 'writing-judge-model',
            '--base-url', 'http://127.0.0.1:' + server.port,
            '--api-key-env', 'WRITING_EVAL_TEST_KEY',
        ], env);
        checkEqual(judged.code, 0, '带 --judge 应 exit 0');
        const judgeRequests = server.requests.slice(beforeJudge).filter((request) => request.body?.model === 'writing-judge-model');
        checkEqual(judgeRequests.length, 4, '每份样本一次评审请求');
        check(judgeRequests.every((request) => JSON.stringify(request.body.messages).includes('待评正文')), '评审提示词带待评正文');
        check(judgeRequests.every((request) => request.body.temperature === 0), '评审 temperature=0');
        const judgeKey = readKey(judgeDir);
        checkEqual(judgeKey.judgeModel, 'writing-judge-model', 'key 记录评审模型');
        check(judgeKey.samples.every((sample) => sample.judge !== null && [1, 2, 3, 4, 5].includes(sample.judge.naturalness)), '评审分数写进 key');
        check(judgeKey.samples.every((sample) => sample.judgeError === null), '本地合法 JSON 不应有 judgeError');
        const judgeReport = readFileSync(join(judgeDir, 'report.md'), 'utf8');
        checkIncludes(judgeReport, '评审模型打分', 'report 含评审打分表');
        checkIncludes(judgeReport, '自然度', 'report 含三维之一');
    } finally {
        await server.close();
    }

    // ---------------------------------------------------------------- 6b. judgeError 降级（in-process）
    {
        const dir = join(TMP, 'inproc-judge-error');
        const fakeFetch = async (url, init) => {
            const body = JSON.parse(init.body);
            const content = body.model === 'judge-x' ? '抱歉，我无法评分。' : '正文一。正文二。正文三。';
            return { ok: true, status: 200, text: async () => JSON.stringify({ choices: [{ message: { content } }] }) };
        };
        const result = await runWritingEval({
            tasks: ['a'],
            runs: 1,
            outDir: dir,
            apiKey: 'fake-key',
            model: 'gen-x',
            judgeModel: 'judge-x',
            fetchImpl: fakeFetch,
            log: () => {},
        });
        checkEqual(result.ok, true, '评审解析失败不应拖垮整轮');
        checkEqual(result.samples.length, 2, '生成照常完成');
        check(result.samples.every((sample) => typeof sample.judgeError === 'string' && sample.judgeError.length > 0), '解析失败记 judgeError');
        check(result.samples.every((sample) => sample.judge === null), '解析失败不写分数');
        const key = readKey(dir);
        check(key.samples.every((sample) => typeof sample.judgeError === 'string'), 'key.json 也记录 judgeError');
        checkIncludes(readFileSync(join(dir, 'report.md'), 'utf8'), '解析失败', 'report 有解析失败列');
    }

    // ---------------------------------------------------------------- 6c. 生成失败 → ok=false / exit 1
    {
        const dir = join(TMP, 'inproc-fail');
        const result = await runWritingEval({
            tasks: ['a'],
            runs: 1,
            outDir: dir,
            apiKey: 'fake-key',
            fetchImpl: async () => { throw new Error('boom'); },
            log: () => {},
        });
        checkEqual(result.ok, false, '生成失败时 ok=false');
        checkEqual(result.failures.length, 2, '两个臂各记一次失败');
        checkIncludes(readFileSync(join(dir, 'report.md'), 'utf8'), '生成失败', 'report 含生成失败段');

        const closedDir = join(TMP, 'closed-port');
        const run = await runCli([
            '--tasks', 'a', '--runs', '1',
            '--out', closedDir,
            '--base-url', 'http://127.0.0.1:' + server.port,
            '--api-key-env', 'WRITING_EVAL_TEST_KEY',
        ], { ...process.env, WRITING_EVAL_TEST_KEY: 'writing-test-key' });
        checkEqual(run.code, 1, '连不上端点时 exit 1');
    }

    // ---------------------------------------------------------------- 7. --self-test
    {
        const run = await runCli(['--self-test']);
        checkEqual(run.code, 0, '--self-test exit 0');
        checkIncludes(run.stdout, 'SELF-TEST PASS', '--self-test 打印 PASS');
        checkIncludes(run.stdout, 'checks=', '--self-test 报告检查数');
    }

    // ---------------------------------------------------------------- 结果
    if (failures.length > 0) {
        console.error('\nFAIL writing-eval：' + failures.length + ' 项不通过（共 ' + checks + ' 项断言）');
        for (const line of failures) console.error('  ✗ ' + line);
        if (!KEEP) rmSync(TMP, { recursive: true, force: true });
        process.exit(1);
    }
    console.log('\nPASS writing-eval：' + checks + ' 通过 / 0 失败');
    if (KEEP) {
        console.log('产物保留在 ' + TMP);
    } else {
        rmSync(TMP, { recursive: true, force: true });
        console.log('临时产物已清理：' + TMP);
    }
}

await main();
