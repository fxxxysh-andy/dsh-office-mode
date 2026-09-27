/**
 * 宿主（DSH）模块与随部署安装的 bundle 的解析（测试专用）。
 *
 * 为什么不直接 `import '@deepseek-ai/schemastery'`：本插件是 link: 进 profile 的
 * （源码在工作区），ESM 按模块自身位置解析裸包名，工作区里没有 @deepseek-ai/*，
 * 所以只能借宿主自己的解析位置。锚点与 src/index.js 的 resolveHostModule 同一套：
 * DSH_PROFILE_DIR → DSH_HOST_ROOT → DSH_CHECKOUT → 惯例位置 → 宿主入口目录 → cwd。
 *
 * 「惯例位置」补两类：$DSH_HOME/profiles/<name>（profile 目录里一定有 package.json），
 * 以及 node 安装目录下的 node_modules（nvm / fnm / volta 都把全局包放在那里）。
 * 随部署安装的 bundle（@deepseek-ai/*）不在 profile 的 node_modules 里，得先把
 * @deepseek-ai/dsh/package.json 解析出来，再从它那里解析目标包 —— 见
 * dshPackageDir() 与 resolveShippedBundle()。
 *
 * 解析不到时一律返回 undefined，调用方各自降级（跳过并明确说明），
 * 绝不把「解析不到」当成「检查通过」。
 */
import { createRequire } from 'node:module';
import { existsSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { homedir } from 'node:os';
import { pathToFileURL } from 'node:url';

/** DSH 主目录：优先环境变量，其次 ~/.dsh。 */
export function dshHome() {
    const raw = typeof process.env.DSH_HOME === 'string' ? process.env.DSH_HOME.trim() : '';
    return raw !== '' ? raw : join(homedir(), '.dsh');
}

/** profile 目录的候选：环境变量点名的那个 + $DSH_HOME/profiles 下真实存在的每一个。 */
function profileDirs() {
    const dirs = [];
    const names = [process.env.DSH_PROFILE, process.env.DSH_PROFILE_NAME, 'web', 'default'];
    for (const name of names) {
        if (typeof name === 'string' && name !== '') dirs.push(join(dshHome(), 'profiles', name));
    }
    try {
        for (const name of readdirSync(join(dshHome(), 'profiles'))) dirs.push(join(dshHome(), 'profiles', name));
    } catch {
        // 没有 profiles 目录就跳过
    }
    return dirs;
}

/** node 安装目录下的模块根：版本管理器把全局包装在这些位置。 */
function nodeInstallDirs() {
    const execDir = dirname(process.execPath);
    return [
        execDir,
        join(execDir, 'node_modules'),
        join(execDir, '..', 'lib', 'node_modules'),
        join(execDir, '..', 'node_modules'),
    ];
}

/** 可用的解析锚点，按可靠性排序（与 src/index.js 同一套，另加惯例位置）。 */
export function hostAnchors(extra = []) {
    const anchors = [
        process.env.DSH_PROFILE_DIR,
        process.env.DSH_HOST_ROOT,
        process.env.DSH_CHECKOUT,
        ...profileDirs(),
        ...nodeInstallDirs(),
        ...extra,
    ].filter((value) => typeof value === 'string' && value !== '');
    if (typeof process.argv[1] === 'string' && process.argv[1] !== '') anchors.push(dirname(process.argv[1]));
    anchors.push(process.cwd());
    return [...new Set(anchors)];
}

/** 从一个锚点建一个 require（锚点不可用时返回 undefined）。 */
export function hostRequire(anchor) {
    try {
        return createRequire(pathToFileURL(join(anchor, 'package.json')).href);
    } catch {
        return undefined;
    }
}

/** 在指定目录的解析位置解析一个包；解析不到返回 undefined。 */
export function resolveFromDir(dir, specifier) {
    const req = hostRequire(dir);
    if (req === undefined) return undefined;
    try {
        return req.resolve(specifier);
    } catch {
        return undefined;
    }
}

/** 在锚点链上解析一个包的绝对路径；都解析不到返回 undefined。 */
export function resolveHostPath(specifier, extra = []) {
    for (const anchor of hostAnchors(extra)) {
        const found = resolveFromDir(anchor, specifier);
        if (found !== undefined) return found;
    }
    return undefined;
}

/** @deepseek-ai/dsh 的安装目录；解析不到返回 undefined。 */
export function dshPackageDir(extra = []) {
    const manifest = resolveHostPath('@deepseek-ai/dsh/package.json', extra);
    if (manifest !== undefined) return dirname(manifest);
    // 万一某个版本的 exports 不暴露 ./package.json：从入口往回找最近的 package.json。
    const entry = resolveHostPath('@deepseek-ai/dsh', extra);
    if (entry === undefined) return undefined;
    let dir = dirname(entry);
    for (let i = 0; i < 6; i += 1) {
        if (existsSync(join(dir, 'package.json'))) return dir;
        const up = dirname(dir);
        if (up === dir) break;
        dir = up;
    }
    return undefined;
}

/** 从部署安装位置解析一个随宿主安装的包（@deepseek-ai/* 与 dsh 自己的依赖）。 */
export function resolveShippedBundle(specifier, extra = []) {
    const dir = dshPackageDir(extra);
    if (dir === undefined) return undefined;
    const direct = resolveFromDir(dir, specifier);
    if (direct !== undefined) return direct;
    return resolveFromDir(join(dir, 'node_modules'), specifier);
}

/** 宿主 Web 入口 lib/bin.js 的路径；DSH_BIN 显式指过就用它。 */
export function dshBinPath() {
    const explicit = typeof process.env.DSH_BIN === 'string' ? process.env.DSH_BIN.trim() : '';
    if (explicit !== '' && existsSync(explicit)) return explicit;
    const dir = dshPackageDir();
    if (dir === undefined) return undefined;
    const candidate = join(dir, 'lib', 'bin.js');
    return existsSync(candidate) ? candidate : undefined;
}
