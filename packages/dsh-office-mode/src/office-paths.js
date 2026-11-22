/**
 * .office 目录的标准布局（第八十四轮）。
 *
 * 只做一件事：把「长期数据」与「随时可删的运行期文件」分开。以前缓存、测试临时目录、
 * 活体验证截图都直接铺在 .office/ 下，时间一长就分不清哪些能删；现在运行期文件统一
 * 收进 .office/runtime/，根目录只留用户看得懂的持久目录：
 *
 *   .office/memory   长期记忆（memory.json / Markdown 投影 / 台账 / 归档）
 *   .office/kb       知识库（manifest + 块文件）
 *   .office/backups  备份
 *   .office/search   检索派工留下的提纲与摘录（可重跑，但常被引用）
 *   .office/runtime  运行期文件，随时可删
 *
 * 迁移旧目录：node scripts/organize-office.mjs --dry-run 先看，再去掉 --dry-run。
 * 清理运行期文件：node scripts/clean-office.mjs --all（或 --older-than 7）。
 */

/** 运行期文件的根。 */
export const RUNTIME_DIR = '.office/runtime';

/** 中间产物目录（PDF 页图、PPT 抽图、Python 产物、音频解码 WAV 等）。 */
export const CACHE_DIR = '.office/runtime/cache';

/** 测试与临时脚本的落点。 */
export const TMP_DIR = '.office/runtime/tmp';

/** 活体验证的截图与 JSON 报告。 */
export const DIAG_DIR = '.office/runtime/diag';

/** 检索派工的结果目录（保留在根，便于文档引用与人工翻看）。 */
export const SEARCH_DIR = '.office/search';

/** 长期记忆目录。 */
export const MEMORY_DIR = '.office/memory';
