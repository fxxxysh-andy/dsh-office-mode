/**
 * 「办公模式插件」开关的回归测试（第六十三轮，用户点名的 `新-21`）。
 *
 * 这一份守的是三件在别处看不见的事：
 *   1. **逐行手术**：开关只动目标行的 `disabled`，其余字节逐字节不变（这个文件的注释
 *      本身就是设计记载，整份重排等于把 diff 变成不可读）；
 *   2. **两处同写**：preset 真源 + profile 侧**已存在的**覆盖行；profile 里没有覆盖行的
 *      行只改真源（凭空新建覆盖行会把「办公模式不要」变成「所有模式都不要」）；
 *   3. **失败形态**：备份写不成 → 一个文件都不动；第二处写失败 → 第一处回滚回去；
 *      未知 id / 非法值 / 错 content-type 当场拒绝。这三条都用**可换的写入器接缝**注入
 *      失败（`writeImpl` / `snapshotImpl`），所以探针验的是真路径，不是读代码相信。
 *
 * 真实文件的两向核对（设置页列出的集合 = preset 里的插件行集合）在 `preset-check.mjs`
 * 里（那套要读本机 profile，不进 CI 档）；这里只用一份**假的 DSH_HOME**（临时目录），
 * 所以本套件是纯 Node、零依赖、不读 `~/.dsh`，可以进 CI 档。
 *
 * 跑法：node test/preset-plugins.mjs
 */
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { join } from 'node:path';

import {
    PRESET_BODY_LIMIT,
    PRESET_PLUGINS_PATH,
    PRESET_RESTORE_PATH,
    atomicWriteFile,
    backupDirOf,
    dshHomeDir,
    listBackups,
    packageRootOf,
    parsePresetPlugins,
    parseProfileOverrides,
    readPresetPlugins,
    registerPresetPlugins,
    resolvePresetTargets,
    restoreLastBackup,
    setDisabledOnLine,
    writePluginToggle,
} from '../src/preset-plugins.js';

const results = [];
async function check(name, fn) {
    try {
        const note = await fn();
        results.push({ name, ok: true, note: typeof note === 'string' ? note : undefined });
    } catch (error) {
        results.push({ name, ok: false, error });
    }
}

/** 本轮探针的工作区（`.office/runtime/tmp` 不进版本库；这里用系统临时目录，CI 上也可写）。 */
const sandbox = mkdtempSync(join(tmpdir(), 'office-preset-plugins-'));
process.on('exit', () => {
    try {
        rmSync(sandbox, { recursive: true, force: true });
    } catch {
        // 清理失败不该影响结论。
    }
});

/**
 * 一份**假 preset**：结构与真文件同形（8 空格 `plugins:`、10 空格的行、12 空格的行内键），
 * 并故意在 `prefix: |-` 的正文里埋两条陷阱（`- id:` 与 `disabled:` 形状的文字）。
 */
const PRESET_SAMPLE = [
    '# 假的办公 preset（测探针用，不是真文件）',
    '- insert:',
    '    - id: preset-office',
    "      name: '@deepseek-ai/dsh-agent-preset'",
    '      config:',
    '        id: office',
    '        plugins:',
    '          # ── 身份与说话方式 ──',
    '          - id: persona',
    "            name: '@deepseek-ai/dsh-persona'",
    '            config:',
    '              complete: true',
    '              prefix: |-',
    '                陷阱一：这一行长得像插件行',
    '                - id: 不是插件行',
    '                陷阱二：这一行长得很像禁用行',
    '                disabled: true',
    '',
    '          # ── 文件 ──',
    '          - id: tool-fs',
    "            name: '@deepseek-ai/dsh-tool-fs'",
    '',
    '          - id: tool-bash',
    "            name: '@deepseek-ai/dsh-tool-bash'",
    '            disabled: true',
    '',
    '          - id: ui-task-board',
    "            name: '@linxin666/dsh-client-ui-task-board'",
    '            disabled: true',
    '        # 列表结束',
    '',
].join('\n');

/** 一份**假 profile 补丁**：覆盖行只有 persona 与 tool-fs（ui-task-board 故意没有）。 */
const PROFILE_SAMPLE = [
    '# Your patch layer for this dsh profile',
    '- id: some-other-row',
    '  name: "@deepseek-ai/dsh-something"',
    '  config:',
    '    value: 1',
    '- id: dsh-office-mode',
    '  config:',
    '    search:',
    '      engine: auto',
    '- id: persona',
    '  disabled: false',
    '- id: tool-fs',
    '  disabled: false',
    '',
].join('\n');

/** 搭一份假的 DSH_HOME：一个 profile（link 指向假包 + link 指向假 preset）。 */
function makeHome() {
    const home = mkdtempSync(join(sandbox, 'home-'));
    const profileDir = join(home, 'profiles', 'web');
    const presetDir = join(home, 'presets', 'office');
    const packageDir = join(home, 'packages', 'dsh-office-mode');
    mkdirSync(join(profileDir, 'node_modules'), { recursive: true });
    mkdirSync(presetDir, { recursive: true });
    mkdirSync(packageDir, { recursive: true });
    writeFileSync(join(presetDir, 'cordis.patch.yml'), PRESET_SAMPLE, 'utf8');
    writeFileSync(join(presetDir, 'package.json'), JSON.stringify({
        name: '@local/dsh-office-preset',
        dsh: { bundle: { patch: './cordis.patch.yml' } },
    }), 'utf8');
    writeFileSync(join(packageDir, 'package.json'), JSON.stringify({ name: 'dsh-office-mode' }), 'utf8');
    writeFileSync(join(profileDir, 'cordis.patch.yml'), PROFILE_SAMPLE, 'utf8');
    writeFileSync(join(profileDir, 'package.json'), JSON.stringify({
        name: 'dsh-profile-web',
        dependencies: {
            'dsh-office-mode': `link:${packageDir}`,
            '@local/dsh-office-preset': `link:${presetDir}`,
        },
        dsh: { profile: { bundles: ['dsh-office-mode', '@local/dsh-office-preset'] } },
    }), 'utf8');
    mkdirSync(join(profileDir, 'node_modules', 'dsh-office-mode'), { recursive: true });
    return { home, profileDir, presetDir, packageDir, presetPatch: join(presetDir, 'cordis.patch.yml'), profilePatch: join(profileDir, 'cordis.patch.yml') };
}

const fixture = makeHome();
const targets = resolvePresetTargets({ dshHome: fixture.home, packageDir: fixture.packageDir });

/** 每次探针前把两份文件复位（探针之间不许互相影响）。 */
function resetFiles() {
    writeFileSync(fixture.presetPatch, PRESET_SAMPLE, 'utf8');
    writeFileSync(fixture.profilePatch, PROFILE_SAMPLE, 'utf8');
    rmSync(backupDirOf(targets), { recursive: true, force: true });
}

const presetText = () => readFileSync(fixture.presetPatch, 'utf8');
const profileText = () => readFileSync(fixture.profilePatch, 'utf8');

/** 两次文本的逐行差（用来证明「只动了目标行」）：对齐首尾相同段，中间那段就是改动。 */
function lineDelta(before, after) {
    const a = before.split('\n');
    const b = after.split('\n');
    let head = 0;
    while (head < a.length && head < b.length && a[head] === b[head]) head += 1;
    let tailA = a.length;
    let tailB = b.length;
    while (tailA > head && tailB > head && a[tailA - 1] === b[tailB - 1]) {
        tailA -= 1;
        tailB -= 1;
    }
    return { before: a.slice(head, tailA), after: b.slice(head, tailB) };
}

// ── 1 发现 ────────────────────────────────────────────────────────────────
await check('发现：从假 $DSH_HOME 里找到装本插件的 profile 与 preset bundle', () => {
    assert.equal(targets.ok, true, targets.error);
    assert.equal(targets.profileName, 'web');
    assert.equal(targets.profileDir, fixture.profileDir);
    assert.equal(targets.presetDir, fixture.presetDir);
    assert.equal(targets.presetPatch, fixture.presetPatch);
    assert.equal(targets.profilePatch, fixture.profilePatch);
    assert.equal(dshHomeDir({ DSH_HOME: fixture.home }), fixture.home);
    assert.equal(dshHomeDir({}), join(homedir(), '.dsh'));
    // packageRootOf：本模块在 src/ 下，上一级就是包根（真包，不是假的）
    assert.equal(packageRootOf(), packageRootOf(new URL('../src/preset-plugins.js', import.meta.url).href));
    return `${targets.profileName} / ${targets.presetDir}`;
});

// ── 2 解析（含陷阱行） ────────────────────────────────────────────────────
await check('解析：只认 plugins: 下深一级的行，prefix 正文里的 `- id:` / `disabled:` 不算', () => {
    const parsed = parsePresetPlugins(PRESET_SAMPLE);
    assert.equal(parsed.ok, true);
    assert.deepEqual(parsed.rows.map((row) => row.id), ['persona', 'tool-fs', 'tool-bash', 'ui-task-board']);
    assert.deepEqual(parsed.rows.map((row) => row.name), [
        '@deepseek-ai/dsh-persona', '@deepseek-ai/dsh-tool-fs',
        '@deepseek-ai/dsh-tool-bash', '@linxin666/dsh-client-ui-task-board',
    ]);
    const persona = parsed.rows[0];
    assert.equal(persona.disabled, undefined, 'persona 的 disabled 是陷阱行里的，不该被认成行内键');
    assert.equal(persona.disabledLine, -1);
    assert.equal(parsed.rows[2].disabled, true);
    assert.equal(parsed.rows[3].disabled, true);
    return `${parsed.rows.length} 行`;
});

await check('解析：真 preset 文件里解析出的行与独立扫出的 `- id:` 行集合一致', () => {
    const real = readFileSync(new URL('../../../presets/office/cordis.patch.yml', import.meta.url), 'utf8');
    const parsed = parsePresetPlugins(real);
    assert.equal(parsed.ok, true);
    // 独立扫法（不借解析器的递归状态）：插件列表区间里缩进正好 10 空格、形如 `- id: x` 的行。
    const lines = real.split('\n');
    const seen = [];
    for (const line of lines) {
        const match = /^ {10}-\s+id:\s*(\S+)\s*$/.exec(line);
        if (match !== null) seen.push(match[1]);
    }
    assert.deepEqual(parsed.rows.map((row) => row.id), seen);
    assert.equal(new Set(seen).size, seen.length, '行 id 必须唯一');
    assert.ok(seen.length >= 20, `真 preset 的插件行看起来太少了：${seen.length}`);
    return `${seen.length} 行`;
});

await check('解析：profile 覆盖行只认顶层直接键，config 里的同名键不算', () => {
    const items = parseProfileOverrides(PROFILE_SAMPLE);
    assert.deepEqual(items.map((item) => item.id), ['some-other-row', 'dsh-office-mode', 'persona', 'tool-fs']);
    assert.equal(items[0].hasName, true);
    assert.equal(items[0].disabledLine, -1, 'some-other-row 没有 disabled 行');
    assert.equal(items[2].disabled, false);
    assert.equal(items[2].disabledLine, 10);
    return `${items.length} 条`;
});

// ── 3 逐行手术 ────────────────────────────────────────────────────────────
await check('手术：改已有 disabled 行只换那一个词（其余字节全等）', () => {
    const parsed = parsePresetPlugins(PRESET_SAMPLE);
    const row = parsed.rows.find((item) => item.id === 'tool-bash');
    const result = setDisabledOnLine(PRESET_SAMPLE, {
        lineIndex: row.disabledLine, insertAfter: row.nameLine, indent: '            ', value: 'false',
    });
    assert.equal(result.ok, true);
    assert.equal(result.changed, 'value');
    const delta = lineDelta(PRESET_SAMPLE, result.text);
    assert.deepEqual(delta.before, ['            disabled: true']);
    assert.deepEqual(delta.after, ['            disabled: false']);
    assert.equal(result.text.split('\n').length, PRESET_SAMPLE.split('\n').length);
    return '1 行 / 1 词';
});

await check('手术：没有 disabled 行的行插一行，插在 name 下面、缩进与 name 同级', () => {
    const parsed = parsePresetPlugins(PRESET_SAMPLE);
    const row = parsed.rows.find((item) => item.id === 'tool-fs');
    const result = setDisabledOnLine(PRESET_SAMPLE, {
        lineIndex: row.disabledLine, insertAfter: row.nameLine, indent: '            ', value: 'true',
    });
    assert.equal(result.changed, 'insert');
    const lines = result.text.split('\n');
    assert.equal(lines[row.nameLine + 1], '            disabled: true');
    assert.equal(lines[row.nameLine + 2], '');
    assert.equal(lines.length, PRESET_SAMPLE.split('\n').length + 1);
    const delta = lineDelta(PRESET_SAMPLE, result.text);
    assert.deepEqual(delta.before, []);
    assert.deepEqual(delta.after, ['            disabled: true']);
    return '+1 行';
});

await check('手术：行尾 CRLF 不会被规范化掉', () => {
    const crlf = PRESET_SAMPLE.replace(/\n/g, '\r\n');
    const parsed = parsePresetPlugins(crlf);
    const row = parsed.rows.find((item) => item.id === 'tool-fs');
    const result = setDisabledOnLine(crlf, {
        lineIndex: row.disabledLine, insertAfter: row.nameLine, indent: '            ', value: 'true',
        eol: '\r\n',
    });
    assert.equal(result.ok, true);
    assert.equal(result.text.includes('\r\n'), true);
    assert.equal(result.text.replace(/\r\n/g, '').includes('\n'), false, '不该混进裸 \\n');
    return 'CRLF 保持';
});

// ── 4 两处同写 ────────────────────────────────────────────────────────────
await check('开关：关掉 persona → 真源那行变 true，profile 覆盖行同步变 true，且只动这两行', () => {
    resetFiles();
    const before = { preset: presetText(), profile: profileText() };
    const result = writePluginToggle(targets, { id: 'persona', enabled: false });
    assert.equal(result.ok, true, result.error);
    assert.deepEqual(result.changed, { preset: true, profile: true });
    assert.equal(result.profileOverridden, true);
    assert.ok(result.snapshot !== null && existsSync(result.snapshot), '必须落一份快照');
    const presetDelta = lineDelta(before.preset, presetText());
    assert.deepEqual(presetDelta.before, []);
    assert.deepEqual(presetDelta.after, ['            disabled: true'], '真源只该「多出一行」，不该改别的行');
    const profileDelta = lineDelta(before.profile, profileText());
    assert.deepEqual(profileDelta.before, ['  disabled: false']);
    assert.deepEqual(profileDelta.after, ['  disabled: true']);
    // 列表读回来也必须是关着的
    const listed = readPresetPlugins(targets);
    assert.equal(listed.rows.find((row) => row.id === 'persona').enabled, false);
    assert.equal(listed.rows.find((row) => row.id === 'persona').profileDisabled, true);
    return 'preset + profile';
});

await check('开关：profile 里没有覆盖行的行（ui-task-board）只改真源', () => {
    resetFiles();
    const beforeProfile = profileText();
    const result = writePluginToggle(targets, { id: 'ui-task-board', enabled: true });
    assert.equal(result.ok, true, result.error);
    assert.deepEqual(result.changed, { preset: true, profile: false });
    assert.equal(result.profileOverridden, false);
    assert.equal(profileText(), beforeProfile, 'profile 文件一个字节都不该动');
    const listed = readPresetPlugins(targets);
    assert.equal(listed.rows.find((row) => row.id === 'ui-task-board').enabled, true);
    return '只写真源';
});

await check('幂等：重复开关不改文件、不落新快照，回执照实说「已经是这个状态」', () => {
    resetFiles();
    writePluginToggle(targets, { id: 'tool-fs', enabled: false });
    const once = { preset: presetText(), profile: profileText(), backups: listBackups(targets).length };
    const again = writePluginToggle(targets, { id: 'tool-fs', enabled: false });
    assert.equal(again.ok, true, again.error);
    assert.deepEqual(again.changed, { preset: false, profile: false });
    assert.equal(again.snapshot, null);
    assert.match(again.diff, /幂等/);
    assert.equal(presetText(), once.preset);
    assert.equal(profileText(), once.profile);
    assert.equal(listBackups(targets).length, once.backups, '幂等时不该多一份快照');
    return '无改动';
});

await check('往返：关掉再打开，语义回到原状态（只多一行显式 `disabled: false`）', () => {
    resetFiles();
    writePluginToggle(targets, { id: 'tool-fs', enabled: false });
    assert.notEqual(presetText(), PRESET_SAMPLE);
    const back = writePluginToggle(targets, { id: 'tool-fs', enabled: true });
    assert.equal(back.ok, true, back.error);
    const delta = lineDelta(PRESET_SAMPLE, presetText());
    // 不删行：那一行是「谁写的」记不住，删掉用户手写的 `disabled: false` 比留一行更糟。
    // 语义上等价（不写这一行 = 启用），所以这里断言的是「只多这一行、行状态是启用的」。
    assert.deepEqual(delta.before, []);
    assert.deepEqual(delta.after, ['            disabled: false']);
    assert.equal(readPresetPlugins(targets).rows.find((row) => row.id === 'tool-fs').enabled, true);
    assert.equal(profileText(), PROFILE_SAMPLE, 'profile 覆盖行要回到原样');
    return '语义还原 / +1 行';
});

// ── 5 失败形态 ────────────────────────────────────────────────────────────
await check('拒绝：未知 id 不写任何文件', () => {
    resetFiles();
    const result = writePluginToggle(targets, { id: 'not-a-row', enabled: false });
    assert.equal(result.ok, false);
    assert.match(result.error, /没有这一行/);
    assert.equal(presetText(), PRESET_SAMPLE);
    assert.equal(profileText(), PROFILE_SAMPLE);
    assert.equal(listBackups(targets).length, 0, '拒绝时不该落快照');
    return '未动';
});

await check('退路先行：备份写不成 → 一个文件都不动', () => {
    resetFiles();
    const result = writePluginToggle(targets, {
        id: 'persona', enabled: false,
        snapshotImpl: () => { throw new Error('磁盘满了（探针注入）'); },
    });
    assert.equal(result.ok, false);
    assert.match(result.error, /写备份失败/);
    assert.equal(presetText(), PRESET_SAMPLE);
    assert.equal(profileText(), PROFILE_SAMPLE);
    return '未动';
});

await check('回滚：第二处写失败 → 第一处被写回原样，且回执标明 rolledBack', () => {
    resetFiles();
    let calls = 0;
    const result = writePluginToggle(targets, {
        id: 'persona', enabled: false,
        writeImpl: (file, text) => {
            calls += 1;
            if (calls === 2) throw new Error('目标被占用（探针注入）');
            atomicWriteFile(file, text);
        },
    });
    assert.equal(result.ok, false);
    assert.match(result.error, /写入失败/);
    assert.equal(result.rolledBack, true);
    assert.equal(calls, 3, '第三次调用是回滚第一处');
    assert.equal(presetText(), PRESET_SAMPLE, '真源必须回到改动前');
    assert.equal(profileText(), PROFILE_SAMPLE);
    return 'rolledBack=true';
});

// ── 6 备份与还原 ──────────────────────────────────────────────────────────
await check('还原：关掉一行再还原 → 两份文件都回到快照那一版', () => {
    resetFiles();
    writePluginToggle(targets, { id: 'persona', enabled: false });
    writePluginToggle(targets, { id: 'tool-fs', enabled: false });
    assert.notEqual(presetText(), PRESET_SAMPLE);
    const backups = listBackups(targets);
    assert.equal(backups.length, 2);
    const restored = restoreLastBackup(targets);
    assert.equal(restored.ok, true, restored.error);
    assert.equal(restored.restored.length, 2);
    // 最新一份快照是「tool-fs 那次改动前」：persona 已关、tool-fs 还没关。
    assert.notEqual(presetText(), PRESET_SAMPLE);
    const listed = readPresetPlugins(targets);
    assert.equal(listed.rows.find((row) => row.id === 'persona').enabled, false);
    assert.equal(listed.rows.find((row) => row.id === 'tool-fs').enabled, true);
    assert.match(profileText(), /- id: persona\n  disabled: true/);
    return `${restored.restored.length} 个文件`;
});

await check('还原：没有备份时明确拒绝', () => {
    resetFiles();
    const result = restoreLastBackup(targets);
    assert.equal(result.ok, false);
    assert.match(result.error, /还没有备份/);
    return '拒绝';
});

// ── 7 路由 ────────────────────────────────────────────────────────────────
function fakeWebServer() {
    const routes = new Map();
    return {
        routes,
        register({ kind, path, handler }) {
            routes.set(`${kind}:${path}`, handler);
            return () => routes.delete(`${kind}:${path}`);
        },
    };
}

function fakeRes() {
    const res = { status: 0, body: '', headers: null };
    res.writeHead = (status, headers) => { res.status = status; res.headers = headers; };
    res.end = (chunk) => { res.body = typeof chunk === 'string' ? chunk : String(chunk ?? ''); };
    return res;
}

function fakeReq({ method = 'GET', headers = {}, body = '' } = {}) {
    const req = new EventEmitter();
    req.method = method;
    req.headers = headers;
    req.destroy = () => {};
    if (body !== '') {
        setImmediate(() => {
            req.emit('data', Buffer.from(body, 'utf8'));
            req.emit('end');
        });
    } else {
        setImmediate(() => req.emit('end'));
    }
    return req;
}

async function callRoute(web, path, { method = 'GET', headers = {}, body = '' } = {}) {
    const handler = web.routes.get(`exact:${path}`);
    assert.ok(handler !== undefined, `没有注册路由 ${path}`);
    const res = fakeRes();
    await handler(fakeReq({ method, headers, body }), res);
    return { res, json: res.body === '' ? null : JSON.parse(res.body) };
}

const web = fakeWebServer();
const disposeRoutes = registerPresetPlugins({ get: (name) => (name === 'webServer' ? web : undefined) }, { targets });

await check('路由：注册了两条（列表 / 还原），没有 webServer 时返回 undefined', () => {
    assert.equal(typeof disposeRoutes, 'function');
    assert.equal(web.routes.has(`exact:${PRESET_PLUGINS_PATH}`), true);
    assert.equal(web.routes.has(`exact:${PRESET_RESTORE_PATH}`), true);
    assert.equal(registerPresetPlugins({ get: () => undefined }), undefined);
    assert.equal(registerPresetPlugins(undefined), undefined);
    return '2 条';
});

await check('路由：GET 给出行清单与两处路径（列表实时来自磁盘）', async () => {
    resetFiles();
    const { res, json } = await callRoute(web, PRESET_PLUGINS_PATH);
    assert.equal(res.status, 200);
    assert.equal(json.ok, true);
    assert.equal(json.rows.length, 4);
    assert.equal(json.presetPath, fixture.presetPatch);
    assert.equal(json.profilePath, fixture.profilePatch);
    assert.equal(json.effect, '新建的办公会话');
    // 「实时」：绕过路由直接改文件，下一次 GET 必须读到新状态
    const flipped = PRESET_SAMPLE.replace(
        "            name: '@deepseek-ai/dsh-tool-bash'\n            disabled: true",
        "            name: '@deepseek-ai/dsh-tool-bash'\n            disabled: false",
    );
    assert.notEqual(flipped, PRESET_SAMPLE, '探针的替换没命中（样本变了？）');
    writeFileSync(fixture.presetPatch, flipped, 'utf8');
    const again = await callRoute(web, PRESET_PLUGINS_PATH);
    assert.equal(again.json.rows.find((row) => row.id === 'tool-bash').enabled, true);
    return `${json.rows.length} 行`;
});

await check('路由：POST 非 JSON content-type → 415；非法 body → 400；PUT → 405', async () => {
    resetFiles();
    const wrongType = await callRoute(web, PRESET_PLUGINS_PATH, { method: 'POST', headers: { 'content-type': 'text/plain' }, body: '{}' });
    assert.equal(wrongType.res.status, 415);
    const badBody = await callRoute(web, PRESET_PLUGINS_PATH, { method: 'POST', headers: { 'content-type': 'application/json' }, body: '{"id":"tool-fs"}' });
    assert.equal(badBody.res.status, 400);
    const badJson = await callRoute(web, PRESET_PLUGINS_PATH, { method: 'POST', headers: { 'content-type': 'application/json' }, body: '{' });
    assert.equal(badJson.res.status, 400);
    const put = await callRoute(web, PRESET_PLUGINS_PATH, { method: 'PUT' });
    assert.equal(put.res.status, 405);
    const getRestore = await callRoute(web, PRESET_RESTORE_PATH);
    assert.equal(getRestore.res.status, 405);
    assert.equal(presetText(), PRESET_SAMPLE);
    return '415 / 400 / 405';
});

await check('路由：POST 合法请求真的改盘，回执带回执与 applied；未知 id 走 ok:false', async () => {
    resetFiles();
    const ok = await callRoute(web, PRESET_PLUGINS_PATH, {
        method: 'POST', headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ id: 'tool-fs', enabled: false }),
    });
    assert.equal(ok.res.status, 200);
    assert.equal(ok.json.ok, true);
    assert.deepEqual(ok.json.applied.changed, { preset: true, profile: true });
    assert.equal(readPresetPlugins(targets).rows.find((row) => row.id === 'tool-fs').enabled, false);
    const unknown = await callRoute(web, PRESET_PLUGINS_PATH, {
        method: 'POST', headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ id: 'nope', enabled: false }),
    });
    assert.equal(unknown.res.status, 200);
    assert.equal(unknown.json.ok, false);
    assert.match(unknown.json.error, /没有这一行/);
    return 'applied 到位';
});

await check('路由：POST /restore 还原上一次；没有备份时 ok:false', async () => {
    resetFiles();
    await callRoute(web, PRESET_PLUGINS_PATH, {
        method: 'POST', headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ id: 'persona', enabled: false }),
    });
    assert.notEqual(presetText(), PRESET_SAMPLE);
    const restored = await callRoute(web, PRESET_RESTORE_PATH, {
        method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}',
    });
    assert.equal(restored.res.status, 200);
    assert.equal(restored.json.ok, true);
    assert.equal(presetText(), PRESET_SAMPLE);
    assert.equal(profileText(), PROFILE_SAMPLE);
    resetFiles();
    const empty = await callRoute(web, PRESET_RESTORE_PATH, {
        method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}',
    });
    assert.equal(empty.json.ok, false);
    assert.match(empty.json.error, /还没有备份/);
    return '还原到位';
});

await check('边界：请求体上限是个有限的数（不给无限上传留口子）', () => {
    assert.equal(Number.isFinite(PRESET_BODY_LIMIT), true);
    assert.ok(PRESET_BODY_LIMIT > 0 && PRESET_BODY_LIMIT <= 64 * 1024, `上限可疑：${PRESET_BODY_LIMIT}`);
    return `${PRESET_BODY_LIMIT} B`;
});

// ── 8 找不到时的形态 ──────────────────────────────────────────────────────
await check('发现失败：空目录给 ok:false 与一句能照办的说明，列表为空', () => {
    const lonely = mkdtempSync(join(sandbox, 'lonely-'));
    const mine = resolvePresetTargets({ dshHome: lonely, packageDir: join(lonely, 'packages', 'dsh-office-mode') });
    assert.equal(mine.ok, false);
    assert.match(mine.error, /找不到办公 preset/);
    const listed = readPresetPlugins(mine);
    assert.equal(listed.ok, false);
    assert.equal(listed.rows.length, 0);
    const wrote = writePluginToggle(mine, { id: 'persona', enabled: false });
    assert.equal(wrote.ok, false);
    return 'ok:false';
});

// ── 收尾 ──────────────────────────────────────────────────────────────────
disposeRoutes();
const failed = results.filter((item) => item.ok !== true);
for (const item of results) {
    if (item.ok === true) console.log(`PASS  ${item.name}${item.note === undefined ? '' : `  —— ${item.note}`}`);
    else console.log(`RED   ${item.name}\n      ${item.error?.message ?? item.error}`);
}
console.log(`\npreset-plugins: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length === 0 ? 0 : 1);
