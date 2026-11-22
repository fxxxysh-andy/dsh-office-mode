/**
 * 「办公模式插件」列表的服务端：列 preset 里的插件行、按行回写 `disabled`、备份与还原。
 *
 * 为什么单独一个模块：设置页那一格改的不是**插件自己的设置文档**（`tools.*` /
 * `search.*`，走 schemastery + dsh-settings），而是**办公 preset 的声明文件**
 * `presets/office/cordis.patch.yml`。两者的写入协议完全不同：设置文档是结构化覆盖，
 * 这里要的是「只动目标行的一个字段，其余逐字节不变」—— YAML 不能整份重排（重排会把
 * 注释与顺序一起改掉，diff 就没人看得懂，而这个文件的注释本身就是设计记载）。
 *
 * 三件必须一起成立的事实（本项目此前的实测，见 preset 文件里的长注释）：
 *
 *   1. **真源是 preset 文件**：`presets/office/cordis.patch.yml`。装进 profile 时是
 *      `link:`（`@local/dsh-office-preset` → 工作区 `presets/office`），所以「工作区那份」
 *      与「安装目录那份」是同一个文件，不存在「写哪一边」的分歧。
 *   2. **profile 侧还有一层覆盖行**：`~/.dsh/profiles/<profile>/cordis.patch.yml` 里
 *      留着 install_bundle 当时展平出来的 `- id: <行 id> / disabled: false`。profile 层
 *      在 bundle 层之后应用，于是这些覆盖行会**盖掉 preset 自己的 disabled**（第五十轮的
 *      实测结论：preset 里的禁用行关不掉 profile 层那一份）。所以只写 preset 真源时，
 *      那些行在运行期看不出变化 —— 开关必须**同时**写这两处才真的生效。
 *      **只更新已存在的覆盖行，绝不新建**：凭空给一行加 profile 覆盖会把「办公模式不要
 *      这个能力」变成「所有模式都不要」，那不是用户按开关的本意。
 *   3. **改了只影响新建的会话**：preset（含插件组合）在新建会话时读取，已经在跑的会话
 *      不受影响 —— 这正是用户要的口径，写在界面回执里。
 *
 * 写入纪律：
 *   - 逐行手术（只替换 / 插入目标行的那一行），其余字节照原样拼回；
 *   - 写前先落一份快照（两个文件的原文 + 时间戳）到 `$DSH_HOME/.office/preset-backups/`，
 *     设置页可以「还原上一次」；**快照写不成就不动文件**（宁可拒绝，也不留下没退路的改动）；
 *   - 临时文件 + 原子改名；第二处写失败时把第一处回滚回去，不留半份改动。
 *
 * @module preset-plugins
 */
import {
    existsSync,
    mkdirSync,
    readFileSync,
    readdirSync,
    realpathSync,
    renameSync,
    statSync,
    unlinkSync,
    writeFileSync,
} from 'node:fs';
import { homedir } from 'node:os';
import { dirname, isAbsolute, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

/** 只读列表与写入共用一条路径（GET 读、POST 写，同一份数据）。 */
export const PRESET_PLUGINS_PATH = '/office-preset/plugins';
/** 还原上一次备份。 */
export const PRESET_RESTORE_PATH = '/office-preset/restore';

/** 请求体上限：这一格只传 `{ id, enabled }`，8 KB 已经宽出一个数量级。 */
export const PRESET_BODY_LIMIT = 8 * 1024;
/** 快照保留份数（多了只会占地方，还原只认最新一份）。 */
export const PRESET_BACKUP_KEEP = 20;
/** preset 文件里插件行的固定缩进层级：`plugins:` 下面一级。 */
const PRESET_ROW_OFFSET = 2;

function asText(value) {
    return typeof value === 'string' ? value : (value === undefined || value === null ? '' : String(value));
}

function realPathOf(target) {
    try {
        return realpathSync(target);
    } catch {
        return resolve(target);
    }
}

/** 按行切开，**每段自带行尾**（CRLF 不会在拼回时被悄悄规范化）。 */
function splitLines(text) {
    const lines = asText(text).match(/[^\n]*\n|[^\n]+$/g);
    return lines === null ? [] : lines;
}

function eolOf(line) {
    const match = /(\r?\n)$/.exec(asText(line));
    return match === null ? '\n' : match[1];
}

/** `$DSH_HOME`（取不到就退 `~/.dsh`）——与 memory.js 的全局目录同一条口径。 */
export function dshHomeDir(env = process.env) {
    const configured = asText(env?.DSH_HOME).trim();
    if (configured !== '') return isAbsolute(configured) ? configured : resolve(homedir(), configured);
    return join(homedir(), '.dsh');
}

/** 本模块所在包的根目录（`src/` 的上一级）。 */
export function packageRootOf(hereUrl) {
    const here = hereUrl === undefined ? import.meta.url : hereUrl;
    const file = here.startsWith('file:') ? fileURLToPath(here) : here;
    return resolve(dirname(file), '..');
}

function readJson(file) {
    try {
        return JSON.parse(readFileSync(file, 'utf8'));
    } catch {
        return null;
    }
}

function readText(file) {
    try {
        return readFileSync(file, 'utf8');
    } catch {
        return null;
    }
}

/** 这个目录里的 cordis.patch.yml 是不是一份 agent preset 的 bundle patch。 */
function isPresetDir(dir) {
    const text = readText(join(dir, 'cordis.patch.yml'));
    return text !== null && text.includes('dsh-agent-preset');
}

/** `link:` / `file:` 依赖的值 → 绝对目录（其它形状返回 null）。 */
function linkTargetOf(profileDir, value) {
    const raw = asText(value).trim();
    const spec = raw.startsWith('link:') ? raw.slice(5) : (raw.startsWith('file:') ? raw.slice(5) : '');
    if (spec === '') return null;
    return isAbsolute(spec) ? resolve(spec) : resolve(profileDir, spec);
}

/**
 * 找出「哪个 profile 装着本插件」以及两处要写的文件。
 *
 * 判据按可靠性排序：profile 的 `node_modules/dsh-office-mode` 指回本包目录 →
 * profile 的依赖里声明了 `dsh-office-mode` → 任一 profile 里能找到 preset bundle。
 * 找不到时给 `error`（照实说清缺什么），调用方据此把界面标成不可用，而不是猜一个路径。
 *
 * @param {{dshHome?: string, packageDir?: string, env?: object}} options
 */
export function resolvePresetTargets(options = {}) {
    const home = options.dshHome ?? dshHomeDir(options.env ?? process.env);
    const packageDir = options.packageDir ?? packageRootOf();
    const profilesDir = join(home, 'profiles');

    const profiles = [];
    if (existsSync(profilesDir)) {
        for (const name of readdirSync(profilesDir)) {
            const dir = join(profilesDir, name);
            let stat = null;
            try {
                stat = statSync(dir);
            } catch {
                continue;
            }
            if (!stat.isDirectory()) continue;
            const manifest = readJson(join(dir, 'package.json'));
            if (manifest === null) continue;
            const deps = Object.assign({}, manifest.dependencies, manifest.devDependencies);
            const presetDir = Object.keys(deps)
                .map((key) => linkTargetOf(dir, deps[key]))
                .find((target) => target !== null && isPresetDir(target)) ?? null;
            const installed = realPathOf(join(dir, 'node_modules', 'dsh-office-mode'));
            profiles.push({
                name,
                dir,
                manifest,
                hasOfficeDep: deps['dsh-office-mode'] !== undefined,
                pointsHere: installed === realPathOf(packageDir),
                presetDir,
                patch: join(dir, 'cordis.patch.yml'),
            });
        }
    }

    const chosen = profiles.find((item) => item.pointsHere && item.presetDir !== null)
        ?? profiles.find((item) => item.hasOfficeDep && item.presetDir !== null)
        ?? profiles.find((item) => item.pointsHere)
        ?? null;

    // 兜底：没有任何 profile 认领时，按工作区布局找 preset（本仓就是这种布局）。
    const fallbackPreset = resolve(packageDir, '..', '..', 'presets', 'office');
    const presetDir = chosen?.presetDir ?? (isPresetDir(fallbackPreset) ? fallbackPreset : null);
    if (presetDir === null) {
        return {
            ok: false,
            dshHome: home,
            packageDir,
            profileName: chosen?.name ?? null,
            error: '找不到办公 preset 的 bundle 目录（profile 里没有指向它的 link:，工作区里也没有 presets/office）。',
        };
    }

    const presetPatch = join(presetDir, 'cordis.patch.yml');
    return {
        ok: true,
        dshHome: home,
        packageDir,
        profileName: chosen?.name ?? null,
        profileDir: chosen?.dir ?? null,
        profilePatch: chosen === null || !existsSync(chosen.patch) ? null : chosen.patch,
        presetDir,
        presetPatch,
        presetPatchExists: existsSync(presetPatch),
    };
}

/** 快照目录：与全局记忆同一层（`$DSH_HOME/.office/...`）。 */
export function backupDirOf(targets) {
    return join(targets.dshHome, '.office', 'preset-backups');
}

/**
 * 解析 preset 文件里的插件行。
 *
 * 只认 `plugins:` 下面**正好深一级**的 `- id:` 行（`plugins:` 八空格、行十空格那种），
 * 所以 persona 的 `prefix: |-` 正文里形如 `- xxx` 的行不会被误当成插件行。注释行、空行
 * 都留在原位，行号照原文记下 —— 回写时按行号手术。
 *
 * @returns {{ok: boolean, error?: string, pluginsLine: number, rowIndent: number, rows: Array<object>}}
 */
export function parsePresetPlugins(text) {
    const lines = splitLines(text);
    let pluginsLine = -1;
    let pluginsIndent = -1;
    for (let i = 0; i < lines.length; i += 1) {
        const match = /^(\s*)plugins:\s*$/.exec(lines[i]);
        if (match === null) continue;
        pluginsLine = i;
        pluginsIndent = match[1].length;
        break;
    }
    if (pluginsLine < 0) {
        return { ok: false, error: '这份 preset 里没有 `plugins:` 列表（它不是 agent preset 的 bundle patch）。', pluginsLine: -1, rowIndent: -1, rows: [] };
    }
    const rowIndent = pluginsIndent + PRESET_ROW_OFFSET;
    const keyIndent = rowIndent + PRESET_ROW_OFFSET;
    const rows = [];
    let current = null;
    const push = () => {
        if (current !== null) rows.push(current);
        current = null;
    };
    for (let i = pluginsLine + 1; i < lines.length; i += 1) {
        const line = lines[i];
        if (/^\s*$/.test(line)) continue;
        const indent = /^\s*/.exec(line)[0].length;
        if (indent < rowIndent) break; // 回到上层：插件列表结束
        if (indent === rowIndent) {
            // 同级只可能是「新的行」或分隔注释；注释不打断当前行（它在行之前）。
            const start = /^(\s*)-\s+id:\s*(\S+)\s*$/.exec(line);
            if (start !== null) {
                push();
                current = { id: start[2], idLine: i, name: '', nameLine: -1, disabled: undefined, disabledLine: -1 };
                continue;
            }
            if (!/^\s*#/.test(line)) push();
            continue;
        }
        if (current === null) continue;
        // 行的**直接键**（比行深一级）才认：`config:` 下面的正文、`prefix: |-` 的文本
        // 里也会有 `name:` / `disabled:` 形状的文字，认了就会读错（persona 那一行最长）。
        if (indent !== keyIndent) continue;
        const nameMatch = /^\s*name:\s*(.*?)\s*$/.exec(line);
        if (nameMatch !== null && current.nameLine < 0) {
            current.name = nameMatch[1].replace(/^['"]|['"]$/g, '');
            current.nameLine = i;
            continue;
        }
        const disabledMatch = /^\s*disabled:\s*(\S+)\s*$/.exec(line);
        if (disabledMatch !== null && current.disabledLine < 0) {
            current.disabled = disabledMatch[1] === 'true';
            current.disabledLine = i;
        }
    }
    push();
    return { ok: true, pluginsLine, rowIndent, rows };
}

/**
 * 解析 profile 侧那些 id 定向的覆盖行（顶层 `- id: x` / `  disabled: false`）。
 *
 * 这一份与 preset 那份的区别：这里**只认已存在的行**（调用方据此决定要不要同步写），
 * 并且顶层行也可能带 `name` / `config`（profile 自己声明的行），一并记下来 —— 界面要能
 * 说清「这一行的开关同时落在 profile 层」。
 */
export function parseProfileOverrides(text) {
    const lines = splitLines(text);
    const items = [];
    let current = null;
    const push = () => {
        if (current !== null) items.push(current);
        current = null;
    };
    for (let i = 0; i < lines.length; i += 1) {
        const line = lines[i];
        const start = /^-\s+id:\s*(\S+)\s*$/.exec(line);
        if (start !== null) {
            push();
            current = { id: start[1], idLine: i, name: '', hasName: false, disabled: undefined, disabledLine: -1, insert: false };
            continue;
        }
        if (current === null) continue;
        if (/^\s*insert:\s*$/.test(line)) current.insert = true;
        // 顶层行的直接键（缩进 2）才认：`config:` 下面还有更深的键。
        if (!/^ {2}\S/.test(line)) continue;
        const nameMatch = /^\s+name:\s*(.*?)\s*$/.exec(line);
        if (nameMatch !== null) {
            current.hasName = true;
            current.name = nameMatch[1].replace(/^['"]|['"]$/g, '');
        }
        const disabledMatch = /^\s+disabled:\s*(\S+)\s*$/.exec(line);
        if (disabledMatch !== null && current.disabledLine < 0) {
            current.disabled = disabledMatch[1] === 'true';
            current.disabledLine = i;
        }
    }
    push();
    return items;
}

/**
 * 把某一行（按行号）的 `disabled` 设成指定值；没有这一行就在参考行下面插一行。
 *
 * @returns {{ok: boolean, error?: string, text?: string, changed: 'value'|'insert'|'same'}}
 */
export function setDisabledOnLine(text, { lineIndex, insertAfter, indent = '', value, eol = '\n' }) {
    const lines = splitLines(text);
    const target = lineIndex >= 0 ? lineIndex : insertAfter;
    if (target < 0 || target >= lines.length) return { ok: false, error: '行号越界（文件在磁盘上变了？）。' };
    if (lineIndex >= 0) {
        const line = lines[lineIndex];
        const match = /^(\s*)disabled:(\s*)(\S*)(\s*)(\r?\n)?$/.exec(line);
        if (match === null) return { ok: false, error: '目标行的形状不像 `disabled: <值>`，拒绝改写。' };
        const next = `${match[1]}disabled:${match[2]}${value}${match[4]}${match[5] ?? ''}`;
        if (next === line) return { ok: true, text, changed: 'same' };
        lines[lineIndex] = next;
        return { ok: true, text: lines.join(''), changed: 'value' };
    }
    lines.splice(target + 1, 0, `${indent}disabled: ${value}${eol}`);
    return { ok: true, text: lines.join(''), changed: 'insert' };
}

/** 原子写：临时文件 + 改名；改名撞车（Windows 上目标被占用）时退一步删了再改。 */
export function atomicWriteFile(file, text) {
    const tmp = `${file}.tmp-${process.pid}-${Date.now()}`;
    writeFileSync(tmp, text, 'utf8');
    try {
        renameSync(tmp, file);
    } catch (error) {
        try {
            unlinkSync(file);
        } catch {
            // 目标本来就不在：继续试着改名。
        }
        try {
            renameSync(tmp, file);
        } catch (second) {
            try {
                unlinkSync(tmp);
            } catch {
                // 清理失败不该盖掉真正的原因。
            }
            throw second === undefined ? error : second;
        }
    }
    return file;
}

/** 落一份快照（两个文件的原文）；返回快照路径。 */
export function writeSnapshot(targets, files, at = new Date()) {
    const dir = backupDirOf(targets);
    mkdirSync(dir, { recursive: true });
    const stamp = `${at.getFullYear()}${String(at.getMonth() + 1).padStart(2, '0')}${String(at.getDate()).padStart(2, '0')}`
        + `-${String(at.getHours()).padStart(2, '0')}${String(at.getMinutes()).padStart(2, '0')}${String(at.getSeconds()).padStart(2, '0')}`
        + `-${String(at.getMilliseconds()).padStart(3, '0')}`;
    // 同一毫秒内连续两次改动会撞同一个文件名：第二份把第一份覆盖掉，还原就少了那一档。
    // 撞名时加 `~1`、`~2` 后缀（`~` 不是时间戳里的字，排序仍能按 base + 序号读）。
    let file = join(dir, `preset-${stamp}.json`);
    for (let seq = 1; existsSync(file) && seq <= 999; seq += 1) {
        file = join(dir, `preset-${stamp}~${seq}.json`);
    }
    const payload = {
        at: at.toISOString(),
        preset: targets.presetPatch,
        profile: targets.profilePatch,
        files: files.map((item) => ({ path: item.path, text: item.text })),
    };
    atomicWriteFile(file, `${JSON.stringify(payload, null, 2)}\n`);
    // 只留最近 PRESET_BACKUP_KEEP 份：还原只认最新一份，多的只是占地方。
    const all = listBackups(targets);
    for (const stale of all.slice(PRESET_BACKUP_KEEP)) {
        try {
            unlinkSync(stale.path);
        } catch {
            // 删不掉就留着，不该让一次写入失败。
        }
    }
    return file;
}

/** 快照文件名拆成「时间戳 + 撞名序号」，供排序用（序号缺省是 0）。 */
function snapshotRank(name) {
    const body = String(name).replace(/^preset-/, '').replace(/\.json$/, '');
    const mark = body.lastIndexOf('~');
    if (mark < 0) return { base: body, seq: 0 };
    const seq = Number(body.slice(mark + 1));
    return Number.isFinite(seq) ? { base: body.slice(0, mark), seq } : { base: body, seq: 0 };
}

/** 快照清单：新的在前（同一毫秒的按撞名序号排，序号大的更新）。 */
export function listBackups(targets) {
    const dir = backupDirOf(targets);
    if (!existsSync(dir)) return [];
    return readdirSync(dir)
        .filter((name) => /^preset-.*\.json$/.test(name))
        .map((name) => {
            const path = join(dir, name);
            let at = null;
            let size = 0;
            try {
                size = statSync(path).size;
            } catch {
                size = 0;
            }
            try {
                at = JSON.parse(readFileSync(path, 'utf8')).at ?? null;
            } catch {
                at = null;
            }
            return { path, name, at, size };
        })
        .sort((a, b) => {
            const left = snapshotRank(a.name);
            const right = snapshotRank(b.name);
            if (left.base !== right.base) return right.base.localeCompare(left.base);
            if (left.seq !== right.seq) return right.seq - left.seq;
            return String(b.name).localeCompare(String(a.name));
        });
}

/** 读出 preset 的插件行 + 每一行在 profile 侧有没有覆盖行。 */
export function readPresetPlugins(targets) {
    if (targets === null || targets.ok !== true) {
        return { ok: false, error: targets?.error ?? '预设位置没找到。', rows: [] };
    }
    const text = readText(targets.presetPatch);
    if (text === null) {
        return { ok: false, error: `读不到 preset 文件：${targets.presetPatch}`, rows: [] };
    }
    const parsed = parsePresetPlugins(text);
    if (parsed.ok !== true) return { ok: false, error: parsed.error, rows: [] };

    const profileText = targets.profilePatch === null ? null : readText(targets.profilePatch);
    const overrides = profileText === null ? [] : parseProfileOverrides(profileText);
    const byId = new Map();
    for (const item of overrides) {
        // 同 id 多条覆盖时以**最后一条**为准（profile 补丁按顺序生效）。
        if (!item.insert) byId.set(item.id, item);
    }
    const rows = parsed.rows.map((row) => {
        const override = byId.get(row.id) ?? null;
        return {
            id: row.id,
            name: row.name,
            enabled: row.disabled !== true,
            declared: row.disabled === undefined ? 'default' : (row.disabled === true ? 'disabled' : 'enabled'),
            profileDisabled: override === null ? null : (override.disabled === true),
            profileOverridden: override !== null,
        };
    });
    const backups = listBackups(targets);
    return {
        ok: true,
        rows,
        presetPath: targets.presetPatch,
        profilePath: targets.profilePatch,
        profileName: targets.profileName,
        backup: backups.length === 0 ? null : backups[0],
        backups: backups.length,
        effect: '新建的办公会话',
    };
}

/**
 * 按 id 开关一行：同时写 preset 真源与 profile 侧**已存在的**覆盖行。
 *
 * 两个写入器是**可换的接缝**：生产路径用默认值，探针用它们注入「第二处写失败」与
 * 「备份写不成」这两种失败形态 —— 不留真接缝的话，「失败要回滚、备份写不成不许动文件」
 * 这两条纪律就只能靠读代码相信。
 *
 * @returns {{ok: boolean, error?: string, ...}} 失败时**不留改动**（备份写不成就不写）。
 */
export function writePluginToggle(targets, {
    id,
    enabled,
    at = new Date(),
    writeImpl = atomicWriteFile,
    snapshotImpl = writeSnapshot,
} = {}) {
    if (targets === null || targets.ok !== true) return { ok: false, error: targets?.error ?? '预设位置没找到。' };
    const want = enabled === true;
    const presetText = readText(targets.presetPatch);
    if (presetText === null) return { ok: false, error: `读不到 preset 文件：${targets.presetPatch}` };
    const parsed = parsePresetPlugins(presetText);
    if (parsed.ok !== true) return { ok: false, error: parsed.error };
    const row = parsed.rows.find((item) => item.id === id);
    if (row === undefined) {
        return { ok: false, error: `preset 里没有这一行：${asText(id)}。列表可能过期了，刷新一次再试。` };
    }

    const presetNext = setDisabledOnLine(presetText, {
        lineIndex: row.disabledLine,
        insertAfter: row.nameLine >= 0 ? row.nameLine : row.idLine,
        indent: ' '.repeat(parsed.rowIndent + PRESET_ROW_OFFSET),
        value: want ? 'false' : 'true',
        eol: eolOf(splitLines(presetText)[row.nameLine >= 0 ? row.nameLine : row.idLine] ?? '\n'),
    });
    if (presetNext.ok !== true) return { ok: false, error: `改 preset 行失败：${presetNext.error}` };

    const presetChanged = presetNext.changed !== 'same';
    let profileNext = null;
    let profileBefore = null;
    let profileChanged = false;
    if (targets.profilePatch !== null) {
        const profileText = readText(targets.profilePatch);
        if (profileText !== null) {
            const overrideRow = parseProfileOverrides(profileText).filter((item) => item.id === id && !item.insert).pop() ?? null;
            if (overrideRow !== null) {
                const result = setDisabledOnLine(profileText, {
                    lineIndex: overrideRow.disabledLine,
                    insertAfter: overrideRow.idLine,
                    indent: '  ',
                    value: want ? 'false' : 'true',
                    eol: eolOf(splitLines(profileText)[overrideRow.idLine] ?? '\n'),
                });
                if (result.ok !== true) return { ok: false, error: `改 profile 覆盖行失败：${result.error}` };
                profileBefore = profileText;
                profileNext = result.text;
                profileChanged = result.changed !== 'same';
            }
        }
    }

    // 幂等：两处都已经是目标值时不写盘、不落新快照 —— 重复开关不该改文件时间戳。
    if (!presetChanged && !profileChanged) {
        return {
            ok: true,
            id,
            enabled: want,
            changed: { preset: false, profile: false },
            profileOverridden: profileNext !== null,
            snapshot: null,
            diff: '已经是这个状态（幂等，没有改动）。',
        };
    }

    const files = [];
    if (presetChanged) files.push({ path: targets.presetPatch, text: presetText });
    if (profileChanged) files.push({ path: targets.profilePatch, text: profileBefore });

    // 退路先行：快照写不成就不动任何文件。
    let snapshot = null;
    try {
        snapshot = snapshotImpl(targets, files, at);
    } catch (error) {
        return { ok: false, error: `写备份失败，没有改动任何文件：${error?.message ?? error}` };
    }

    const changed = { preset: presetChanged, profile: false };
    try {
        if (presetChanged) writeImpl(targets.presetPatch, presetNext.text);
        if (profileChanged) {
            writeImpl(targets.profilePatch, profileNext);
            changed.profile = true;
        }
    } catch (error) {
        // 第二处写失败 → 把第一处回滚回去，不留半份改动。
        let rolledBack = true;
        try {
            if (changed.preset) writeImpl(targets.presetPatch, presetText);
            if (changed.profile && profileBefore !== null) writeImpl(targets.profilePatch, profileBefore);
        } catch {
            rolledBack = false;
        }
        return {
            ok: false,
            error: `写入失败：${error?.message ?? error}`,
            rolledBack,
            snapshot,
        };
    }
    return {
        ok: true,
        id,
        enabled: want,
        changed,
        profileOverridden: profileNext !== null,
        snapshot,
        diff: '只改了目标行的 disabled。',
    };
}

/** 还原最新一份快照（两个文件一起写回）。 */
export function restoreLastBackup(targets) {
    if (targets === null || targets.ok !== true) return { ok: false, error: targets?.error ?? '预设位置没找到。' };
    const backups = listBackups(targets);
    if (backups.length === 0) return { ok: false, error: '还没有备份可还原。' };
    const newest = backups[0];
    let payload = null;
    try {
        payload = JSON.parse(readFileSync(newest.path, 'utf8'));
    } catch (error) {
        return { ok: false, error: `备份读不出来：${error?.message ?? error}` };
    }
    const wrote = [];
    for (const item of Array.isArray(payload.files) ? payload.files : []) {
        if (typeof item?.path !== 'string' || typeof item?.text !== 'string') continue;
        try {
            atomicWriteFile(item.path, item.text);
            wrote.push(item.path);
        } catch (error) {
            return { ok: false, error: `还原失败（${item.path}）：${error?.message ?? error}`, wrote };
        }
    }
    if (wrote.length === 0) return { ok: false, error: '这份备份里没有可还原的文件。' };
    return { ok: true, restored: wrote, at: payload.at ?? newest.at ?? null, snapshot: newest.path };
}

/**
 * 注册两条路由：`GET/POST /office-preset/plugins` 与 `POST /office-preset/restore`。
 *
 * 写入面的边界（写给下一个改这里的人）：
 *   - 方法只认 GET / POST，其余 405；
 *   - POST **必须**是 `application/json`（浏览器里跨站的简单请求带不了这个 content-type，
 *     于是 CSRF 那条路被挡住）；请求体上限 8 KB；
 *   - 只接受 `{ id, enabled }` 两个字段，id 必须**正好是 preset 里的一行**，enabled 必须是
 *     布尔 —— 非法值当场 400，不猜；
 *   - 能改的文件只有发现到的那两个（preset 真源 + profile 覆盖文件），没有路径参数。
 *
 * @param {object} ctx 插件上下文（需要 ctx.webServer）
 * @param {{targets?: object}} options targets 固定时传进来（测试用；不传则每次现算）
 * @returns {() => void|undefined}
 */
export function registerPresetPlugins(ctx, { targets } = {}) {
    const webServer = typeof ctx?.get === 'function' ? ctx.get('webServer') : undefined;
    if (webServer === undefined || webServer === null || typeof webServer.register !== 'function') {
        return undefined;
    }
    const resolveTargets = () => targets ?? resolvePresetTargets({});

    const readBody = (req) => new Promise((settle, fail) => {
        let size = 0;
        const chunks = [];
        req.on('data', (chunk) => {
            size += chunk.length;
            if (size > PRESET_BODY_LIMIT) {
                fail(new Error('请求体太大。'));
                try {
                    req.destroy();
                } catch {
                    // 已经断开就没什么可做的。
                }
                return;
            }
            chunks.push(chunk);
        });
        req.on('end', () => settle(Buffer.concat(chunks).toString('utf8')));
        req.on('error', (error) => fail(error));
    });

    const disposers = [];
    disposers.push(webServer.register({
        kind: 'exact',
        path: PRESET_PLUGINS_PATH,
        handler: async (req, res) => {
            const method = String(req.method ?? 'GET').toUpperCase();
            if (method !== 'GET' && method !== 'HEAD' && method !== 'POST') {
                sendJson(res, 405, { ok: false, error: '这个端点只支持 GET 与 POST。' });
                return;
            }
            if (method === 'GET' || method === 'HEAD') {
                const current = resolveTargets();
                sendJson(res, 200, readPresetPlugins(current));
                return;
            }
            const contentType = String(req.headers?.['content-type'] ?? '');
            if (!contentType.includes('application/json')) {
                sendJson(res, 415, { ok: false, error: 'POST 要用 application/json（跨站表单发不了这个类型）。' });
                return;
            }
            let payload = null;
            try {
                payload = JSON.parse(await readBody(req));
            } catch (error) {
                sendJson(res, 400, { ok: false, error: `请求体不是 JSON：${error?.message ?? error}` });
                return;
            }
            const id = asText(payload?.id).trim();
            if (id === '' || typeof payload?.enabled !== 'boolean') {
                sendJson(res, 400, { ok: false, error: '请求体要形如 { "id": "<插件行 id>", "enabled": true|false }。' });
                return;
            }
            const current = resolveTargets();
            const result = writePluginToggle(current, { id, enabled: payload.enabled });
            if (result.ok !== true) {
                sendJson(res, 200, Object.assign({}, readPresetPlugins(current), { ok: false, error: result.error, rolledBack: result.rolledBack }));
                return;
            }
            sendJson(res, 200, Object.assign({}, readPresetPlugins(current), {
                applied: { id: result.id, enabled: result.enabled, changed: result.changed, profileOverridden: result.profileOverridden, diff: result.diff, snapshot: result.snapshot },
            }));
        },
    }));
    disposers.push(webServer.register({
        kind: 'exact',
        path: PRESET_RESTORE_PATH,
        handler: async (req, res) => {
            const method = String(req.method ?? 'GET').toUpperCase();
            if (method !== 'POST') {
                sendJson(res, 405, { ok: false, error: '还原只认 POST。' });
                return;
            }
            const result = restoreLastBackup(resolveTargets());
            const current = resolveTargets();
            sendJson(res, 200, Object.assign({}, readPresetPlugins(current), { ok: result.ok === true, error: result.ok === true ? undefined : result.error, restored: result.restored, restoredAt: result.at }));
        },
    }));

    return () => {
        for (const dispose of disposers) {
            try {
                dispose?.();
            } catch {
                // 释放失败不该影响其它路由。
            }
        }
    };
}

function sendJson(res, status, payload) {
    const body = `${JSON.stringify(payload)}\n`;
    res.writeHead(status, {
        'content-type': 'application/json; charset=utf-8',
        'content-length': Buffer.byteLength(body, 'utf8'),
        'cache-control': 'no-store',
    });
    res.end(body);
}
