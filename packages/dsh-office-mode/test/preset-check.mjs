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
import {
    bytes,
    personaBranchQuotes,
    personaSections,
    personaToolBullets,
    personaToolSection,
    personaVariables,
    renderPersona,
} from './persona-shape.mjs';
import { CAPABILITY_MAP_BUDGET_BYTES, CAPABILITY_VARIABLE, renderCapabilityMap } from '../src/capabilities.js';
import { NO_SUBAGENTS_HINT } from '../src/search.js';
import { WEB_FAILURE_KINDS } from '../src/web.js';

/**
 * persona 的字节预算（第十八轮 P2-5 / `18-21` 的减法审计守门）。
 *
 * 第四十三轮审计前是 **17,262 B**（200+ 行），审到 **15,934 B**。预算取实测 + 余量，
 * 与工具面预算同一条纪律：**新增能力不靠加宽预算落地** —— 想加一句先删一句，
 * 或者显式调高并写清理由。
 *
 * 第四十七轮补了一条同口径的断言：**渲染后**的 persona（把 `{{office_capabilities}}`
 * 换成最坏情况的能力地图）也在这个预算内 —— 只看 preset 文件里的字节会漏掉注入文本，
 * 而注入文本同样进每一次请求。
 */
const PERSONA_BUDGET_BYTES = 16300;

/**
 * persona 的分节清单（每一节都要能回答「删掉它会重现哪个失败模式」）。
 *
 * 这张表是**有意识决定的登记**：新增一节必须同时改这里，等于逼一次「它守的是
 * 哪个实测失败模式」的自问 —— 第三十九轮那句「只并进、没删过一句」不会再无声发生。
 */
const PERSONA_SECTIONS = [
    '说话',
    '做事',
    '探能力只探一次',
    '失败与收尾',
    '失败分支',
    '读 PDF',
    '做文档',
    '算与画',
    '音频与视频',
    '内容',
    '检索',
    '演示文稿要多走一步',
    '记忆',
    '协作',
    '工具',
];

/** 「工具」一节的字节上界：一句路由 + 一个变量引用（实测 453 B，留一点余量）。 */
const TOOL_MAP_SECTION_BUDGET_BYTES = 620;

/**
 * 最坏情况的能力地图：每类引擎都满、且派工走长的那一支（没有 subagents）。
 *
 * 渲染后的 persona 是**每一次请求**都要重发的那一份，所以预算要按最坏情况钉，
 * 而不是按「本机这台恰好探到什么」钉。三处都是实测出来的最长分支：
 *   - 音视频故意**只给一部分**（缺 VAD）：「部分缺件」那句比「三样齐」那句长；
 *   - 解释器给一个**超长**文件名（地图里会被截到 16 字符，否则最坏情况没有上界）；
 *   - 派工走「没有 subagents」那一支（那句话更长）。
 */
const WORST_CAPABILITY_MAP = renderCapabilityMap({
    pdf: { render: ['pdftoppm', 'pdftocairo', 'mutool', 'gs'], text: ['pdftotext'] },
    tex: ['latexmk', 'xelatex', 'lualatex'],
    preview: true,
    python: 'C:/somewhere/aaaaaaaaaaaaaaaaaaaaaaaaaaaa.exe',
    av: { enabled: true, available: false, ffmpeg: true, ffprobe: true, model: true, vad: false, runtime: true },
}, { subagents: false });

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
                    // 办公 preset 的检索口子（第三十轮）：资料搜索只走插件自己的抓取通道，
                    // persona 必须写清两个抓取工具的适用场合（查一个事实点 / 打开一个页面）
                    // 与「渠道覆盖与落盘」的三步走，否则模型会在两套路子之间反复试探。
                    record(
                        'persona 里写了内置检索（渠道覆盖与落盘）',
                        personaPrefix.includes('office_search_run') && personaPrefix.includes('内置检索'),
                        '要写清 office_search_run 与派工的内置回退',
                    );
                    record(
                        'persona 里写了两个抓取工具（office_web_search / office_web_fetch）',
                        personaPrefix.includes('office_web_search') && personaPrefix.includes('office_web_fetch'),
                        'preset 不声明 tool-web：persona 要说清这两个插件工具的适用场合',
                    );
                    // 宿主的 tool-web 行必须**不在** preset 的 plugins 列表里（第三十轮收掉）：
                    // 声明了它，宿主的 web_search / web_fetch 就会回到办公会话 —— 那正是
                    // 用户点名不要的「走宿主的路」。
                    const toolWeb = rows.find((row) => row?.id === 'tool-web');
                    record(
                        'preset 不声明宿主的 tool-web（检索只走插件的抓取通道）',
                        toolWeb === undefined,
                        'plugins 列表里不该再有 - id: tool-web / name: @deepseek-ai/dsh-tool-web',
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
                    // ── 第十八轮 P0-2～P0-6（第三十八轮补进 persona）──────────────────
                    //
                    // 这五条各有一次实测反证，逐条钉住。放在 persona 里的理由都一样：
                    // 唯一的系统提示是唯一「每次请求都在」的位置，写进 office_help 只有查过的人看得到。
                    //
                    // P0-2：读 PDF 那一节的例子原先只有 dpi，本身就会撞上 20 页上限
                    // （V5；session4 因此造出 pages:[…] 这种不存在的写法）。
                    record(
                        'persona 的读 PDF 一节写了 from/to 与单次页数上限（P0-2）',
                        personaPrefix.includes('一次最多 20 页')
                        && personaPrefix.includes('from / to')
                        && personaPrefix.includes('from: 1, to: 20'),
                        'PDF 读取要写明「先 info 拿页数，再按 20 页一批传 from/to」，否则长文档一次渲染就被上限拒绝',
                    );
                    // P0-3：session1 里同一步连做 5 次探测调用（thinking 里自己引了「不要试探」又违例）。
                    record(
                        'persona 里写了「探能力只探一次、优先 office_help」（P0-3）',
                        personaPrefix.includes('只允许一次')
                        && personaPrefix.includes('office_help({ topic })'),
                        '把「自由权衡」改成有上限的规则：一次探测不够就交付，不要连着换工具试',
                    );
                    // P0-4：V9 —— 一次工具失败被写成 [critical] 的「环境事实」，污染了后续会话。
                    record(
                        'persona 里写了「一次失败 ≠ 能力不存在，环境级结论要两轮复现」（P0-4）',
                        personaPrefix.includes('两轮复现')
                        && personaPrefix.includes('环境级结论'),
                        '记忆污染比一次失败贵得多：环境级结论必须两轮复现才成立',
                    );
                    // P0-5：session3 三十次调用 0 交付、两次被用户打断。
                    record(
                        'persona 里写了「做不到就当场交付结论与缺口」（P0-5）',
                        personaPrefix.includes('做不到就当场用一段话交付')
                        && personaPrefix.includes('为什么缺'),
                        '失败时先说结论与缺口，不要为了把工具跑通继续试',
                    );
                    // P0-6：V10 —— 11 条折行警告改到剩 1 条（差一个字）还再跑一轮，且收尾没提它。
                    record(
                        'persona 里写了「收尾必须列未消解警告」（P0-6）',
                        personaPrefix.includes('未消解的警告')
                        && personaPrefix.includes('已达标'),
                        '收尾不列未消解警告 = 静默带过；达标时给一句正面回执，省掉为 1 个字再跑的一轮',
                    );
                    // 3-1：第三轮定下的「版式与内容量匹配」在 persona 里长期只有近似表述
                    // （「要点超过 7 条」「卡片正文只有一行却占了半页」），没有可照办的
                    // 行数与条数阈值。这一条把三处判据钉住 —— 判据写不清，模型就只能凭感觉
                    // 排页，「一句话拆四页」与「半页一句话」两种极端都会回来。
                    record(
                        'persona 里写了版式与内容量匹配的判据（3-1）',
                        personaPrefix.includes('卡片正文不足两行')
                        && personaPrefix.includes('要点超过 6 条')
                        && personaPrefix.includes('少于 3 条'),
                        '判据要写成行数与条数：卡片正文不足两行改要点页、要点 >6 拆页、<3 并页或补足',
                    );
                    // 18-17：P1-8 —— session4 里同一步对同一份 main.tex 连发了可合并的
                    // 多次 edit。规则要写成「合并成一次调用」并点名 ops 数组（word / excel
                    // 的 edit 与 office.files.edit 都吃它），否则模型只会看到「一次做完」这种
                    // 泛泛表述，仍旧一步十次小改。
                    record(
                        'persona 里写了「同一步多处修改合并成一次调用」（18-17）',
                        personaPrefix.includes('合并成一次调用') && personaPrefix.includes('ops 数组'),
                        '要点名 ops 数组：不写它，模型不知道 edit 能一次吃多条改动（`office.ppt.revise` 里的 ops 不算）',
                    );
                    // 18-15 / 18-22：persona 的**分支条件**必须与插件真实文案对齐。
                    //
                    // V8 的根因就是「persona 的分支条件写了一串从未出现过的错误文本」：
                    // 旧版写「如果它报『当前组合里没有 subagents 服务』…」，而当时真实
                    // 报错是 `tools.restrict() names unknown global tools …`，模型照旧规则
                    // 去派了子代理。这一条把「失败分支」一节里每行第一个「」的词逐个与
                    // 插件常量做**机器核对**（两向：persona 不许编词；四类失败一类不许漏）。
                    const branchQuotes = personaBranchQuotes(personaPrefix);
                    const branchSources = [
                        ...Object.values(WEB_FAILURE_KINDS),
                        NO_SUBAGENTS_HINT,
                    ];
                    const invented = branchQuotes.filter(
                        (quoted) => !branchSources.some((source) => source.includes(quoted)),
                    );
                    record(
                        `persona 的失败分支词都能在插件文案里找到（核对 ${branchQuotes.length} 个）`,
                        branchQuotes.length > 0 && invented.length === 0,
                        invented.length === 0
                            ? ''
                            : `这些词插件从来不会报出来（V8 同类）：${invented.join('、')}`,
                    );
                    const missingKinds = ['配置缺失', '网络出口不可达', '目标站拒绝', '没拿到结果']
                        .filter((kind) => !branchQuotes.includes(kind));
                    record(
                        'persona 的失败分支覆盖了四类失败（18-15）',
                        missingKinds.length === 0,
                        missingKinds.length === 0 ? '' : `漏了：${missingKinds.join('、')}`,
                    );
                    // 18-21：IFScale 思路的减法审计要有**守门**，否则下一轮又会「只并进、
                    // 没删过一句」（第三十九轮自述 persona 增重 2.5 KB 就是这么来的）。
                    // 两条：① 总字节预算；② 分节清单 —— 每一节都是「删掉它会重现哪个失败
                    // 模式」的答案，新增一节必须同时改这张清单（=一次有意识的决定）。
                    record(
                        `persona 总字节在预算内（${bytes(personaPrefix)} / ${PERSONA_BUDGET_BYTES}）`,
                        bytes(personaPrefix) <= PERSONA_BUDGET_BYTES,
                        '减法审计的守门：要加一句就得先删一句（或显式调高预算并写清理由）',
                    );
                    const sections = personaSections(personaPrefix);
                    const missingSections = PERSONA_SECTIONS.filter((title) => !sections.includes(title));
                    const extraSections = sections.filter((title) => !PERSONA_SECTIONS.includes(title));
                    record(
                        `persona 的分节与清单一致（${sections.length} 节）`,
                        missingSections.length === 0 && extraSections.length === 0,
                        missingSections.length === 0 && extraSections.length === 0
                            ? ''
                            : `缺：${missingSections.join('、') || '无'}；多：${extraSections.join('、') || '无'}`,
                    );

                    // ── 18-20 / 23-2：能力地图一句话 + 实际能力由插件注入 ──────────────
                    //
                    // 第十八轮 P2-4：persona 里放一句「哪件事找谁」，不列工具的完整参数；
                    // 第二十三轮 23-2：写死的能力清单会与运行期错配，改成插件注入实际能力。
                    // 两件事落在同一节里，所以守门也放在一起：
                    //   ① 那一节是一句路由（只许一条 `- `），字节有上界；
                    //   ② 引用的变量必须由插件注册（两边改名会同时被这条抓住）；
                    //   ③ 渲染后的 persona（最坏能力地图）仍在 persona 预算内 —— 这是
                    //      「先减后加」真正要守的量，光看 preset 文件里的字节会漏掉注入文本。
                    const toolSection = personaToolSection(personaPrefix);
                    record(
                        `persona 的「工具」一节只留一句能力地图（${bytes(toolSection)} / ${TOOL_MAP_SECTION_BUDGET_BYTES} B，${personaToolBullets(toolSection)} 条）`,
                        toolSection !== undefined
                        && bytes(toolSection) <= TOOL_MAP_SECTION_BUDGET_BYTES
                        && personaToolBullets(toolSection) === 1,
                        'P2-4：把「哪件事找谁」压成一句；工具的参数与清单由工具面和插件注入承担',
                    );
                    const personaVars = personaVariables(personaPrefix);
                    const knownVars = ['cwd', CAPABILITY_VARIABLE];
                    const unknownVars = personaVars.filter((name) => !knownVars.includes(name));
                    record(
                        `persona 引用的变量都是宿主或插件注册过的（${personaVars.join(' / ')}）`,
                        unknownVars.length === 0,
                        unknownVars.length === 0
                            ? ''
                            : `这些变量没人注册，装配时会抛 unknown prompt variable：${unknownVars.join('、')}`,
                    );
                    record(
                        `persona 引用了插件的能力地图变量 {{${CAPABILITY_VARIABLE}}}（23-2）`,
                        personaVars.includes(CAPABILITY_VARIABLE),
                        'persona 是 complete 段：插件注册的 section 进不来，实际能力只能走 {{变量}} 注入',
                    );
                    record(
                        `能力地图在最坏情况下不超预算（${bytes(WORST_CAPABILITY_MAP)} / ${CAPABILITY_MAP_BUDGET_BYTES}）`,
                        bytes(WORST_CAPABILITY_MAP) <= CAPABILITY_MAP_BUDGET_BYTES,
                        '地图跟着 persona 进每一次请求：引擎清单封顶，超出就要先减再改预算',
                    );
                    const renderedPersona = renderPersona(personaPrefix, WORST_CAPABILITY_MAP);
                    record(
                        `渲染后的 persona 在预算内（${bytes(renderedPersona)} / ${PERSONA_BUDGET_BYTES}）`,
                        bytes(renderedPersona) <= PERSONA_BUDGET_BYTES,
                        '注入文本也算常驻成本：按最坏能力地图钉，别只看 preset 文件里的字节',
                    );
                }
                record('没有误关 present', !rows.some((row) => row?.id === 'present' && row?.disabled === true), '');

                // 办公插件那一行必须是**启用的**（第四十七轮复核 F13）：preset 的 persona 引用
                // 了插件注册的 `{{office_capabilities}}`，插件停用而 preset 还装着的话，
                // 办公会话**每一轮**都会在 preStep 抛 unknown prompt variable ——
                // 停用插件不是「少几个工具」，是整个模式起不来。
                const pluginRows = collectProfileRows().filter((row) => row.id === 'dsh-office-mode');
                if (pluginRows.length === 0) {
                    console.log('注意：profile 里找不到 dsh-office-mode 行（插件未装？），跳过启用态核对。');
                } else {
                    const enabledPlugin = pluginRows.some((row) => !row.disabled && row.bundleEnabled !== false);
                    record(
                        `办公插件在 profile 里是启用的（找到 ${pluginRows.length} 行）`,
                        enabledPlugin,
                        'persona 引用了插件注册的 {{office_capabilities}}：插件停用而 preset 还在，办公会话每轮装配都会失败',
                    );
                }

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
                // 第三十轮起 preset **不再写** mnemon 的禁用行（bundle 已在 profile 停用，
                // 平时拦一个没挂载的东西是死配置）；这套核对因此变成「守门」：只要
                // bundle 保持停用，办公会话里就不会出现 mnemon；谁把 bundle 重新启用
                // 而没有在 preset 里补禁用行，这里会当场红。
                const mnemonRows = collectProfileRows().filter((row) => row.id.startsWith('mnemon'));
                if (mnemonRows.length === 0) {
                    console.log('注意：profile 里找不到 mnemon 行（bundle 已卸载？），跳过 mnemon 核对。');
                } else {
                    // 「生效的 mnemon 行」= 既没被 preset 关掉、bundle 也还启用着的那些。
                    // 停用的 bundle 贡献的行本来就不在组合里。
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
                        leaking.length === 0 ? '' : `这些 mnemon 行没被关掉，会出现在办公会话里：${leaking.join('、')}`
                            + '（preset 已不再写 mnemon 禁用行 —— 要么保持 bundle 停用，要么在 preset 里补禁用行）',
                    );
                }

                // 任务看板（第三十二轮，用户要求「办公模式下关闭 taskboard 功能」）。                // 与 mnemon 那条同一套判据：bundle 在 profile 里启用着，所以只能靠 preset
                // 的禁用行把它挡在办公会话之外；挡不住的话，宿主半侧的 10 个 task_board_*
                // 工具与浏览器半侧的看板界面会一起漏进来（工具面与提示都白涨）。
                // 两条断言：① preset 里确实写了这一行（改 id 会被抓）；② profile 里那一行
                // 真的被关掉了（bundle 换名 / 换 id 时会被抓）。
                record(
                    'preset 里写了任务看板的禁用行（id: ui-task-board）',
                    rows.some((row) => row?.id === 'ui-task-board' && row?.disabled === true),
                    '办公模式下不要任务看板：10 个 task_board_* 工具 + 看板界面都该关掉',
                );
                const boardRows = collectProfileRows().filter((row) => row.id === 'ui-task-board');
                if (boardRows.length === 0) {
                    console.log('注意：profile 里找不到任务看板行（bundle 未装？），跳过任务看板核对。');
                } else {
                    const boardClosed = new Set(rows.filter((row) => row?.disabled === true).map((row) => row.id));
                    const boardLeaking = boardRows
                        .filter((row) => !row.disabled && row.bundleEnabled !== false)
                        .map((row) => row.id)
                        .filter((id) => !boardClosed.has(id));
                    record(
                        `任务看板在办公模式下已关闭（profile 里有 ${boardRows.length} 个任务看板行）`,
                        boardLeaking.length === 0,
                        boardLeaking.length === 0 ? '' : `这些行没被关掉，task_board_* 与看板界面会出现在办公会话里：${boardLeaking.join('、')}`,
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