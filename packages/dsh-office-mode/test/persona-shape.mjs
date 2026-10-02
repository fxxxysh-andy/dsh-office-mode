/**
 * persona 的形状抽取（`preset-check` 与 `measure` 共用）。
 *
 * 为什么单独一个模块：第十八轮 §5.2 第 2、3 条要求「同一批输入换标点 / 换分隔符 /
 * 调条目顺序，断言最差变体仍通过」。要做到这一点，守 persona 的那些断言必须抽**语义**
 * 而不是**排版**：分节标题、失败分支词、变量引用、能力地图那一节 —— 换一种引号、
 * 换一个分隔符、把条目打乱，抽出来的东西都该一样。抽取器住在这里，`preset-check`
 * 与 `test/measure.mjs` 用的是同一份实现，扰动回归测的就是它们本身。
 *
 * 只做纯文本处理，不读盘、不依赖 YAML。
 *
 * @module dsh-office-mode/test/persona-shape
 */
import { existsSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { resolveShippedBundle } from './host-modules.mjs';
import { CAPABILITY_VARIABLE } from '../src/capabilities.js';

/** UTF-8 字节（persona 预算按字节算，与工具面预算同一口径）。 */
export function bytes(text) {
    return Buffer.byteLength(String(text ?? ''), 'utf8');
}

/**
 * persona 里所有成对的引号形态。
 *
 * 失败分支那一节用「」包住插件真实报出的词；换一种引号（“”/『』）在语义上完全等价，
 * 抽取器要认得出，否则扰动回归会误报「分支词丢了」。半角双引号不收：它在正文里
 * 承担别的角色（引用设置项、举例），收进来会把噪声当成分支词。
 */
const QUOTE_PAIRS = [['「', '」'], ['“', '”'], ['『', '』']];

/** 一个段落（到下一个空行为止）里按出现顺序取出的成对引用内容。 */
function quotedWordsIn(lines, start) {
    const words = [];
    for (let index = start; index < lines.length; index += 1) {
        const line = lines[index];
        if (line.trim() === '') break;
        const pattern = new RegExp(QUOTE_PAIRS.map(([open, close]) => `${open}([^${close}]+)${close}`).join('|'));
        const match = pattern.exec(line);
        if (match !== null) words.push(match.slice(1).find((value) => value !== undefined));
    }
    return words;
}

/**
 * persona 的分节标题：非空、不以「-」开头、不缩进、不是第一行（首句）、不是变量占位行。
 */
export function personaSections(prefix) {
    const titles = [];
    for (const [index, line] of String(prefix ?? '').split(/\r?\n/).entries()) {
        if (index === 0) continue;
        const text = line.trim();
        if (text === '' || text.startsWith('-') || line.startsWith(' ')) continue;
        if (/^\{\{[^{}]*\}\}$/.test(text)) continue;
        titles.push(text.split(/[（(]/)[0].trim());
    }
    return titles;
}

/**
 * 「失败分支」一节里每行**第一个**引用里的词（`18-15` / `18-22` 的机器核对对象）。
 *
 * 只取第一个：一行里后面出现的引用往往是设置页格子名、状态词，不是分支条件本身。
 */
export function personaBranchQuotes(prefix) {
    const lines = String(prefix ?? '').split(/\r?\n/);
    const start = lines.findIndex((line) => line.trim().startsWith('失败分支'));
    if (start === -1) return [];
    return quotedWordsIn(lines, start + 1);
}

/**
 * persona 里引用过的 `{{变量}}`。
 *
 * 办公 preset 的 persona 是 `complete: true` 的**唯一**提示段，插件注册的 section 在
 * 办公会话里进不了提示词（宿主 dsh-system-prompt 最后会把 sections 收敛成那一段），
 * 能进这份提示词的只有 `{{变量}}` 的插值。所以「引用了哪个变量」是硬约束：
 * 缺一个就是每一次装配都抛 `unknown prompt variable`，会话直接起不来。
 */
export function personaVariables(prefix) {
    const names = [];
    for (const match of String(prefix ?? '').matchAll(/\{\{([^{}]*)\}\}/g)) names.push(match[1]);
    return names;
}

/**
 * persona 的「工具」一节（标题行到下一个分节标题之前）。
 *
 * 第十八轮 P2-4 要的就是这一节：一句「哪件事找谁」的能力地图，**不列工具的参数**。
 * 它现在还有第二个角色（23-2）：实际能力由插件注入，这里只留路由。
 */
export function personaToolSection(prefix) {
    const lines = String(prefix ?? '').split(/\r?\n/);
    const start = lines.findIndex((line) => line.trim().startsWith('工具（能力地图'));
    if (start === -1) return undefined;
    const body = [];
    for (let index = start; index < lines.length; index += 1) {
        const line = lines[index];
        // 一节的内容 = 标题 + 空行 + 缩进的续行 + `- ` 条目 + 变量占位行；
        // 其余非空行就是下一节的标题（personaSections 的判据与之对齐）。
        const inside = index === start
            || line.trim() === ''
            || line.startsWith(' ')
            || line.trim().startsWith('- ')
            || /^\{\{[^{}]*\}\}$/.test(line.trim());
        if (!inside) break;
        body.push(line);
    }
    return body.join('\n').trimEnd();
}

/** 「工具」一节里 `- ` 起头的条目数（P2-4：只许一句路由）。 */
export function personaToolBullets(section) {
    return String(section ?? '').split(/\r?\n/).filter((line) => line.trim().startsWith('- ')).length;
}

/** 把 persona 里的能力地图变量换成给定文本（渲染后的真身）。 */
export function renderPersona(prefix, map) {
    return String(prefix ?? '').split(`{{${CAPABILITY_VARIABLE}}}`).join(map);
}

/** persona 的语义指纹：扰动回归比的就是它（排版变了、这组东西不该变）。 */
export function personaFingerprint(prefix) {
    const sections = personaSections(prefix);
    const branches = personaBranchQuotes(prefix);
    const variables = personaVariables(prefix);
    return {
        // 分节与分支按**集合**比：守门判据不该依赖位置（第十八轮 §5.2 第 3 条的
        // 位置依赖测的是模型，这里能测的是「我们的抽取器与断言位置无关」）。
        sections: [...sections].sort(),
        branches: [...branches].sort(),
        variables: [...variables].sort(),
        toolBullets: personaToolBullets(personaToolSection(prefix)),
        hasRouting: String(prefix ?? '').includes('文档走 office_run'),
        hasTeams: ['spawn_teammate', 'send_message', 'wait_agent', 'team_task'].every((name) => String(prefix ?? '').includes(name)),
    };
}

/** 语义指纹是否一致（同样的键、同样的值）。 */
export function fingerprintEquals(left, right) {
    return JSON.stringify(left) === JSON.stringify(right);
}

/**
 * 读 presets/office 里的 persona（YAML 解析器用随部署安装的那份，同步 require）。
 *
 * 找不到解析器或文件时抛错：这一份是测量口径的输入，静默跳过等于把「没跑」当「通过」。
 *
 * @param {string} [presetDir] 默认仓库里的 presets/office
 * @returns {{ composition: string, prefix: string }}
 */
export function readPersonaPrefix(presetDir) {
    const here = dirname(fileURLToPath(import.meta.url));
    const dir = presetDir ?? resolve(here, '../../../presets/office');
    const yamlPath = resolveShippedBundle('js-yaml');
    if (typeof yamlPath !== 'string' || !existsSync(yamlPath)) {
        throw new Error('找不到随部署安装的 js-yaml，无法解析 preset —— 别把「找不到解析器」当成通过');
    }
    const composition = join(dir, 'cordis.patch.yml');
    if (!existsSync(composition)) throw new Error(`找不到 preset 组合文件：${composition}`);
    const require = createRequire(pathToFileURL(yamlPath).href);
    const yaml = require(yamlPath);
    const rows = yaml.load(readFileSync(composition, 'utf8'));
    const declared = Array.isArray(rows) ? rows[0]?.insert?.[0] : undefined;
    const personaRow = (declared?.config?.plugins ?? []).find((row) => row?.id === 'persona');
    if (personaRow === undefined) throw new Error('preset 里找不到 persona 行');
    return { composition, prefix: String(personaRow.config?.prefix ?? '') };
}