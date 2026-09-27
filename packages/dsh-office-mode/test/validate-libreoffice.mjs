/**
 * 独立校验：用本机部署自带的 LibreOfficeKit 把生成的 Office 文件转成 PDF。
 *
 * 为什么值得单独做一次：单元测试只能证明「文件符合我们自己理解的 OOXML」。
 * LibreOffice 是另一套实现，它肯把文件渲染出来，才是「文件真的没坏」的旁证。
 * 这一步只用于开发期校验，不属于插件本身 —— 插件保持零依赖。
 *
 * 用法：
 *   node test/validate-libreoffice.mjs [--dir .office/tmp/e2e] [--kit <libreoffice-kit/lib/index.js>]
 */
import { existsSync, mkdirSync, readdirSync, rmSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { resolveShippedBundle } from './host-modules.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);

function argValue(name, fallback) {
    const at = args.indexOf(name);
    return at === -1 || args[at + 1] === undefined ? fallback : args[at + 1];
}

// --dir 相对「当前工作目录」解析，不是相对脚本目录：脚本目录下再套一层
// test/ 会得到一个不存在的路径（这个坑已经踩过一次）。
const targetDir = argValue('--dir', '') === ''
    ? fileURLToPath(new URL('../../../.office/tmp/e2e/', import.meta.url))
    : resolve(process.cwd(), argValue('--dir', ''));
// libreoffice-kit 是随部署安装的 bundle，从宿主的解析位置找（见 host-modules.mjs）；
// --kit 显式指过就用它。
const kitCandidates = [
    argValue('--kit', ''),
    resolveShippedBundle('@deepseek-ai/libreoffice-kit'),
].filter((candidate) => typeof candidate === 'string' && candidate !== '');

async function loadKit() {
    for (const candidate of kitCandidates) {
        if (!existsSync(candidate)) continue;
        try {
            const module = await import(pathToFileURL(candidate).href);
            if (typeof module.createConverter === 'function') return module;
        } catch (error) {
            console.log(`  跳过 ${candidate}：${error.message.split('\n')[0]}`);
        }
    }
    return undefined;
}

const kit = await loadKit();
if (kit === undefined) {
    console.log('libreoffice-kit 不可用，跳过独立校验（如实报告，不当成通过）。');
    process.exit(2);
}

const files = existsSync(targetDir)
    ? readdirSync(targetDir)
        .filter((name) => /\.(docx|xlsx|pptx)$/i.test(name))
        .map((name) => join(targetDir, name))
    : [];

if (files.length === 0) {
    console.log(`没有可校验的文件：${targetDir}`);
    process.exit(1);
}

const outDir = fileURLToPath(new URL('../../../.office/tmp/libreoffice-pdf/', import.meta.url));
rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

const converter = await kit.createConverter({ timeoutMs: 180_000 });
console.log(`引擎后端：${converter.backend}`);
let ok = 0;
let bad = 0;
try {
    for (const file of files) {
        const name = file.slice(file.lastIndexOf('\\') + 1);
        // 输出名必须带上原扩展名：同一份内容的 .docx 与 .pptx 会撞到同一个 PDF 名，
        // 而 render() 拒绝写已存在的输出路径。
        const outputPath = join(outDir, `${name}.pdf`);
        try {
            const result = await converter.render({ inputPath: file, outputPath });
            const bytes = statSync(outputPath).size;
            const missing = result.missingFonts.length > 0 ? `（缺字体：${result.missingFonts.join('、')}）` : '';
            console.log(`OK    ${name}  → PDF ${(bytes / 1024).toFixed(1)} KB${missing}`);
            ok += 1;
        } catch (error) {
            console.log(`FAIL  ${name}  → ${error.name}: ${String(error.message).split('\n')[0]}`);
            bad += 1;
        }
    }
} finally {
    await converter.dispose();
}

console.log(`libreoffice-validate: ${ok} 通过 / ${bad} 失败（PDF 在 ${outDir}）`);
process.exit(bad > 0 ? 1 : 0);
