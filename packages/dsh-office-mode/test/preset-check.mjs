/**
 * 本地 preset 的形状检查。
 *
 * 复刻 dsh-agent-presets 的判定规则（entryListProblem）：顶层必须是数组，
 * 每一行必须是带非空 `name` 的映射，group 行递归检查。preset 只要有一条不满足，
 * 在选择器里就是「坏 preset」，不可选也删不掉 —— 所以这个检查值得单独跑。
 *
 * 用 deployment 里的 js-yaml（本地 workspace 没有 YAML 解析器），
 * 找不到时退化为「按缩进找顶层条目」的弱检查，并明确说明是弱检查。
 *
 * 跑法：node test/preset-check.mjs
 */
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { dshPackageDir, resolveShippedBundle } from './host-modules.mjs';

const here = dirname(fileURLToPath(import.meta.url));
// 默认检查仓库里的 preset bundle；用 --dir 可以检查别处的副本。
//
// 2026-09-24（第九轮）：preset 从「遗留目录式」迁成了**声明式 bundle**。
// 0.1.7 起 DSH 不再读 $DSH_HOME/.agent-presets/<id>/（dsh-agent-preset 的编辑
// 技能原话：Nothing reads that directory any more），预设必须是一个 bundle，
// patch 里插入一行 @deepseek-ai/dsh-agent-preset 声明；组合清单从
// config.plugins 取，不再有独立的 agent.cordis.yml / preset.yml。
const dirArg = process.argv.indexOf('--dir');
const presetDir = dirArg === -1 || process.argv[dirArg + 1] === undefined
    ? resolve(here, '../../../presets/office')
    : resolve(process.argv[dirArg + 1]);
const compositionPath = join(presetDir, 'cordis.patch.yml');
const manifestPath = join(presetDir, 'package.json');

// YAML 解析器与随部署安装的 bundle 都从宿主的解析位置找（见 host-modules.mjs）：
// 本插件是 link: 进 profile 的，工作区里没有这些包，把某台机器的安装路径写死，
// 到别的机器上必然落空。
const YAML_CANDIDATES = [
    resolveShippedBundle('js-yaml'),
].filter((candidate) => typeof candidate === 'string' && candidate !== '');

// 随部署安装的 bundle（@deepseek-ai/*）不在 profile 的 node_modules 里，
// 用 dsh 安装根下的 node_modules 作为部署根，避免再解析一次。
const DSH_PACKAGE_DIR = dshPackageDir();
const SHIPPED_ROOTS = [
    process.env.DSH_SHIPPED_ROOT,
    DSH_PACKAGE_DIR === undefined ? undefined : join(DSH_PACKAGE_DIR, 'node_modules'),
].filter((root) => typeof root === 'string' && root !== '');

async function loadYaml() {
    for (const candidate of YAML_CANDIDATES) {
        if (!existsSync(candidate)) continue;
        try {
            const module = await import(pathToFileURL(candidate).href);
            return module.default ?? module;
        } catch {
            // 换下一个候选；全部失败时走弱检查。
        }
    }
    return undefined;
}

/** 与 dsh-agent-presets 的 entryListProblem 等价。 */
function entryListProblem(rows, at = '') {
    if (!Array.isArray(rows)) {
        return at === '' ? '顶层必须是插件行数组' : `group ${at} 必须是数组`;
    }
    for (const [index, row] of rows.entries()) {
        const label = at === '' ? `第 ${index + 1} 行` : `${at} 第 ${index + 1} 行`;
        if (typeof row !== 'object' || row === null || Array.isArray(row)) return `${label} 不是插件行`;
        const { name, group, config } = row;
        if (typeof name !== 'string' || name === '') return `${label} 缺少 name`;
        if (group === true) {
            const nested = entryListProblem(config, label);
            if (nested !== undefined) return nested;
        }
    }
    return undefined;
}

/**
 * profile 组合里真实存在的行 id。
 *
 * 用途：preset 里 `disabled: true` 的行是按 id 关掉宿主组合里的行的，
 * **id 写错不会报错，只是静默失效**（2026-09-22 的实测：预设以为关掉了
 * 子代理与命令行，后来装进来的 Agent Teams、dsh-ppt 却全漏进了办公会话，
 * 工具面从 9 个涨到 41 个）。这里把每个 disabled id 拿去核对一遍。
 *
 * 第三方 bundle 从 profile 自己的 node_modules 读；@deepseek-ai/* 这类随
 * 部署安装的 bundle 不在 profile 里，用 DSH_SHIPPED_ROOT 或已知候选路径兜底。
 * 两个来源都读不到时跳过这项检查，并明确打印说明。
 */
function collectProfileRowIds() {
    return new Map(collectProfileRows().map((row) => [row.id, row.bundle]));
}

/** 同上，但保留行本身（id + bundle + disabled），供 Agent Teams 这类「按实际写法」核对用。 */
function collectProfileRows() {
    const rows = [];
    const seen = new Set();
    const home = process.env.DSH_HOME;
    if (typeof home !== 'string' || home === '') return rows;
    const profilesDir = join(home, 'profiles');
    if (!existsSync(profilesDir)) return rows;

    for (const profileName of readdirSync(profilesDir)) {
        const manifest = join(profilesDir, profileName, 'package.json');
        if (!existsSync(manifest)) continue;
        let bundles;
        let dependencies;
        try {
            const pkg = JSON.parse(readFileSync(manifest, 'utf8'));
            bundles = pkg?.dsh?.profile?.bundles;
            dependencies = pkg?.dependencies;
        } catch {
            continue;
        }
        const enabled = new Set(bundles ?? []);
        // 同时扫「已安装但未启用」的 bundle（dependencies 里挂着、bundles 列表里没有）：
        // 它们贡献的行不在生效组合里，但 preset 里对这些 id 的 disabled 行仍然「有指向」
        // —— 2026-09-23 关闭 dsh-mnemon 之后，办公预设里那 9 行就属于这种情况。
        // 不扫它们，那些行会被误判成「id 拼错、静默失效」。
        const candidatesBundles = [...enabled, ...Object.keys(dependencies ?? {})];
        for (const bundle of new Set(candidatesBundles)) {
            const parts = String(bundle).split('/');
            const candidates = [
                join(profilesDir, profileName, 'node_modules', ...parts, 'cordis.patch.yml'),
                ...SHIPPED_ROOTS.map((root) => join(root, ...parts, 'cordis.patch.yml')),
            ];
            const patch = candidates.find((file) => existsSync(file));
            if (patch === undefined) continue;
            const text = readFileSync(patch, 'utf8');
            // 按「一条顶层 patch 条目」切块，这样同一块里的 disabled 能跟 id 对上。
            const blocks = text.split(/^(?=\s*-\s*(?:id|insert):)/m);
            for (const block of blocks) {
                const idMatch = /^\s*-\s*id:\s*([^\s#]+)/m.exec(block);
                if (idMatch === null) continue;
                const id = idMatch[1].replace(/^['"]|['"]$/g, '');
                if (seen.has(id)) continue;
                seen.add(id);
                rows.push({
                    id,
                    bundle: `${profileName}/${bundle}`,
                    disabled: /^\s*disabled:\s*true\s*$/m.test(block),
                    bundleEnabled: enabled.has(bundle),
                });
            }
        }
    }
    return rows;
}

/** 没有 YAML 解析器时的弱检查：只看顶层 `- ` 条目里有没有 `name:`。 */
function weakCheck(source) {
    const problems = [];
    let current = undefined;
    let sawName = false;
    for (const [index, line] of source.split(/\r?\n/).entries()) {
        if (/^- /.test(line) || /^-\s*$/.test(line)) {
            if (current !== undefined && !sawName) problems.push(`${current} 缺少 name`);
            current = `第 ${index + 1} 行`;
            sawName = /(^|\s)name:/.test(line);
        } else if (current !== undefined && /^\s+name:/.test(line)) {
            sawName = true;
        }
    }
    if (current !== undefined && !sawName) problems.push(`${current} 缺少 name`);
    return problems;
}

const results = [];
function record(name, ok, detail) {
    results.push({ name, ok, detail });
}

// 1. bundle 形状
record('preset bundle 目录存在', existsSync(presetDir), presetDir);
record('cordis.patch.yml 存在', existsSync(compositionPath), compositionPath);
record('package.json 存在', existsSync(manifestPath), '');
if (existsSync(manifestPath)) {
    let manifest;
    try {
        manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
        record('package.json 可解析', true, '');
    } catch (error) {
        record('package.json 可解析', false, error.message);
    }
    if (manifest !== undefined) {
        record('声明了 dsh.bundle.patch', typeof manifest?.dsh?.bundle?.patch === 'string', '');
        record(
            'patch 路径指向 cordis.patch.yml',
            manifest?.dsh?.bundle?.patch === './cordis.patch.yml',
            String(manifest?.dsh?.bundle?.patch),
        );
    }
}

// 2. 组合形状
if (existsSync(compositionPath)) {
    const source = readFileSync(compositionPath, 'utf8');
    const yaml = await loadYaml();
    if (yaml === undefined) {
        const problems = weakCheck(source);
        record('（弱检查）每行都有 name', problems.length === 0, problems.join('；'));
        console.log('注意：没找到 js-yaml，只做了弱检查。');
    } else {
        let rows;
        try {
            const patch = yaml.load(source);
            const declared = Array.isArray(patch) ? patch[0]?.insert?.[0] : undefined;
            record(
                '声明行形状正确（insert → preset-office / @deepseek-ai/dsh-agent-preset）',
                declared?.id === 'preset-office' && declared?.name === '@deepseek-ai/dsh-agent-preset',
                '',
            );
            record('声明的 id 是 office', declared?.config?.id === 'office', String(declared?.config?.id));
            record('声明有 name', typeof declared?.config?.name === 'string' && declared.config.name !== '', '');
            record(
                '声明有 description',
                typeof declared?.config?.description === 'string' && declared.config.description !== '',
                '',
            );
            record('声明不含 order（沿用遗留 preset.yml 的取舍）', declared?.config?.order === undefined, '');
            rows = declared?.config?.plugins;
            record('YAML 可解析', true, '');
        } catch (error) {
            record('YAML 可解析', false, error.message);
        }
        if (rows !== undefined) {
            const problem = entryListProblem(rows);
            record('形状合法（每行都是带 name 的映射）', problem === undefined, problem ?? '');
            if (Array.isArray(rows)) {
                const ids = rows.map((row) => row?.id).filter((id) => typeof id === 'string');
                const duplicates = ids.filter((id, index) => ids.indexOf(id) !== index);
                record('id 无重复', duplicates.length === 0, [...new Set(duplicates)].join('、'));
                const disabledWithoutName = rows.filter((row) => row?.disabled === true && typeof row?.name !== 'string');
                record('disabled 行也带 name', disabledWithoutName.length === 0, `${disabledWithoutName.length} 行`);

                const enabled = rows.filter((row) => row?.disabled !== true);
                record('至少启用一行', enabled.length > 0, `${enabled.length} 行启用`);
                record(
                    '启用的行都有显式 name',
                    enabled.every((row) => typeof row.name === 'string' && row.name !== ''),
                    '',
                );
                const officeRow = rows.find((row) => row?.id === 'persona');
                record('persona 行存在且 complete', officeRow?.config?.complete === true, '');
                // 若将来出现「非办公型」的 preset（例如纯检索助手），它不生成文档、也不派活，
                // 所以「办公约定」「协作约定」这两项对它不适用。按 persona 里认领的
                // 职责判断，而不是按目录名 —— 目录可以复制改名，职责文本不会。
                const personaPrefix = typeof officeRow?.config?.prefix === 'string' ? officeRow.config.prefix : '';
                const isSearchAgent = personaPrefix.includes('检索子代理');
                if (isSearchAgent) {
                    record(
                        'persona 里写了检索职责（检索子代理）',
                        personaPrefix.includes('写进') && personaPrefix.includes('来源 URL'),
                        '',
                    );
                } else {
                    record(
                        'persona 里写了办公约定',
                        personaPrefix.includes('office_help'),
                        '',
                    );
                    // 记忆能力在插件里（office_memory + office_run 的自动台账），
                    // 但 persona 不写「什么时候记、记到哪一层」就等于没加：
                    // 模型不会自己发现「用户纠正过的事下次要照办」。
                    record(
                        'persona 里写了记忆约定（office_memory）',
                        personaPrefix.includes('office_memory') && personaPrefix.includes('记忆'),
                        '三层记忆的能力在插件里，用法要写进 persona',
                    );
                    // Python 通道是办公模式里唯一的编程出口（命令行是关掉的）：
                    // persona 不写它，模型就会继续拿手工办法画图，或者干脆说做不到。
                    record(
                        'persona 里写了 Python 约定（office.python）',
                        personaPrefix.includes('office.python') && personaPrefix.includes('OUT_DIR'),
                        'office.python 的用法（check / OUT_DIR / savefig / use_cjk_font）要写进 persona',
                    );
                    // 音频与视频的内容提取（第二十二轮）：能力在插件的 office.av 里，
                    // persona 不写它，模型遇到录音与视频只会说「做不到」或者去手工抄写。
                    record(
                        'persona 里写了音频与视频约定（office.av）',
                        personaPrefix.includes('office.av') && personaPrefix.includes('SenseVoice'),
                        'office.av 的用法（check / transcribe / frames / extract）与缺件报错要写进 persona',
                    );
                    // 办公 preset 是一份完整的组合，里面没有 web_search 这类工具：
                    // persona 必须写清「查资料走插件自带的联网通道」，否则模型会以为
                    // 环境不支持检索而放弃（2026-09-25 的真实会话就是这样停住的）。
                    record(
                        'persona 里写了内置检索（没有 web_search 也能查）',
                        personaPrefix.includes('office_search_run') && personaPrefix.includes('内置检索'),
                        '办公 preset 里没有 web_search：要写清 office_search_run 与派工的内置回退',
                    );
                    // 通道真不可用时必须照实说，不能假装查过 —— 这是检索这件事的诚实底线。
                    record(
                        'persona 里写了通道不可用时照实说明',
                        personaPrefix.includes('内置检索') && personaPrefix.includes('照实告诉用户'),
                        '两条通道都不可用时要说明缺哪个 Key、在哪配，不要编造来源',
                    );
                    // 跨渠道矛盾的处理纪律原先只写在 office_search_brief 输出的提纲里
                    // （search-routes.js），直查（office_search_run）那条路上没有 —— 而直查
                    // 的调用量更大（第十八轮 P1-12：纪律触达率与调用量正好相反）。
                    // 放进 persona 才与走哪条路无关。
                    record(
                        'persona 里写了跨来源矛盾的处理纪律',
                        personaPrefix.includes('不一致时两种都写出来') && personaPrefix.includes('不要替用户选一种'),
                        '两个来源说法冲突时要并列写出，不要替用户选一种：这条要在唯一系统提示里，不能只在提纲里',
                    );
                }
                record('没有误关 present', !rows.some((row) => row?.id === 'present' && row?.disabled === true), '');

                const rowIds = collectProfileRowIds();
                if (rowIds.size === 0) {
                    console.log('注意：读不到 profile 组合的行清单，跳过 disabled id 核对。');
                } else {
                    const disabledIds = rows
                        .filter((row) => row?.disabled === true && typeof row?.id === 'string')
                        .map((row) => row.id);
                    const dead = disabledIds.filter((id) => !rowIds.has(id));
                    record(
                        `disabled 的 id 都能在 profile 组合里找到（核对 ${disabledIds.length} 个）`,
                        dead.length === 0,
                        dead.length === 0 ? '' : `这些 id 不存在，禁用静默失效：${dead.join('、')}`,
                    );
                }

                // mnemon（记忆）在办公模式下必须关干净。关不干净的坑有两层：
                //   1) 整套 mnemon 组件挂在一个 group 行 mnemon-bundle 下，
                //      只关组行时组内每个组件行仍会被 loader 逐条解析；
                //   2) 组内组件又各自是独立的行 id，漏一个就漏一片。
                // 所以这里核对「profile 里真实存在的 mnemon 行」是否都被关掉了。
                const mnemonRows = collectProfileRows().filter((row) => row.id.startsWith('mnemon'));
                if (mnemonRows.length === 0) {
                    console.log('注意：profile 里找不到 mnemon 行（bundle 已卸载？），跳过 mnemon 核对。');
                } else {
                    // 2026-09-23 起整个 dsh-mnemon bundle 已在 profile 里停用（不再出现在
                    // dsh.profile.bundles 里），所以「生效的 mnemon 行」= 既没被 preset 关掉、
                    // bundle 也还启用着的那些。停用的 bundle 贡献的行本来就不在组合里。
                    const bundleOff = mnemonRows.every((row) => row.bundleEnabled === false);
                    const open = mnemonRows
                        .filter((row) => !row.disabled && row.bundleEnabled !== false)
                        .map((row) => row.id);
                    const closed = new Set(
                        rows.filter((row) => row?.disabled === true).map((row) => row.id),
                    );
                    const leaking = open.filter((id) => !closed.has(id));
                    record(
                        `mnemon 在办公模式下已全部关闭（profile 里有 ${mnemonRows.length} 个 mnemon 行${bundleOff ? '，整个 bundle 已停用' : ''}）`,
                        leaking.length === 0,
                        leaking.length === 0 ? '' : `这些 mnemon 行没被关掉，会出现在办公会话里：${leaking.join('、')}`,
                    );
                }

                // Agent Teams 是「按 profile 实际写法」核对的一项，不是「关掉即可」的一项。
                // 2026-09-22 先把它关掉、随后按用户要求补回：真正的坑是两头都不报错——
                //   1) 写 disabled 行的 id 拼错 → 静默失效，工具照旧在；
                //   2) 补回时只删掉 disabled 行，却忘了 profile 里的行分布在
                //      @deepseek-ai/dsh-experimental-agent-team-profile（服务行 + 工具行）
                //      与 ...-web-profile（界面行）两个 bundle 里 → 也是静默的。
                // 所以这里核对 preset 的意图与组合的事实是否一致。
                record(
                    'Agent Teams 已启用（没有把 agent-team / tool-agent-team 关掉）',
                    !rows.some((row) => row?.disabled === true
                        && (row?.id === 'agent-team' || row?.id === 'tool-agent-team')),
                    '被 disabled 关掉后 spawn_teammate 等工具不会出现',
                );
                const teamRows = collectProfileRows();
                if (teamRows.length === 0) {
                    console.log('注意：读不到 profile 组合的行清单，跳过 Agent Teams 行核对。');
                } else {
                    const missing = ['agent-team', 'tool-agent-team']
                        .filter((id) => !teamRows.some((row) => row.id === id));
                    record(
                        'profile 组合里有 agent-team 与 tool-agent-team 两行',
                        missing.length === 0,
                        missing.length === 0 ? '' : `缺少：${missing.join('、')}（检查 profile 是否装了 agent-team bundle）`,
                    );
                    const personaRow = rows.find((row) => row?.id === 'persona');
                    const prefix = typeof personaRow?.config?.prefix === 'string' ? personaRow.config.prefix : '';
                    // 检索子代理不派活给更下一层，协作约定对它不适用。
                    if (!prefix.includes('检索子代理')) {
                        record(
                            'persona 里写了协作约定（spawn_teammate）',
                            prefix.includes('spawn_teammate'),
                            '启用了但没告诉模型怎么用，等于白开',
                        );
                        record(
                            'persona 的工具清单包含 Agent Teams 的收口工具',
                            ['send_message', 'team_task', 'wait_agent'].every((name) => prefix.includes(name)),
                            '',
                        );
                    }
                }
            }
        }
    }
}

const failed = results.filter((item) => !item.ok);
for (const item of results) {
    console.log(`${item.ok ? 'PASS' : 'FAIL'}  ${item.name}${item.ok || item.detail === '' ? '' : `  → ${item.detail}`}`);
}
console.log(`preset-check: ${results.length - failed.length}/${results.length} 通过`);
process.exit(failed.length > 0 ? 1 : 0);