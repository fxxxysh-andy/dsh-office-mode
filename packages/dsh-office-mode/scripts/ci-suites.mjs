/**
 * CI 与发布守门跑的那一批套件 —— **唯一真源**（第五十六轮 S-06 的落地）。
 *
 * 为什么单独抽一份：`.github/workflows/ci.yml` 与 `release.yml` 从前各自手写同一串名字，
 * 于是「发布守门到底测了什么」有两处真源，而且都容易忘 —— 第三～五十四轮加的
 * `format-excel` / `format-ppt` / `format-ppt-revise` / `chart`（Excel 与 PPT 的 XML、
 * 关系与统计回归）一直**不在**守门里：它们在默认 `ORDER` 里，但工作流是**显式点名**跑的，
 * 名字没写进去就等于没跑。反向文件检查只能证明测试已登记，不能证明发布执行过它。
 * 现在工作流只写 `node scripts/run-suites.mjs --ci`，名单只有这一份，改一处两边都变。
 *
 * 收录判据（**只收不依赖本机环境的**）：纯 Node、零第三方依赖、不联网、不需要
 * LibreOffice / Office / 无头 Chrome / ffmpeg / Python / pandoc / 本机 TeX，也不读
 * `~/.dsh` 的 profile。够不上这条线的别加进来 —— CI runner 上它们必然红，红了也不说明
 * 代码有问题，只会把守门变成噪音（环境相关的套件按名字点着跑：`validate-libreoffice` /
 * `validate-chart-com` 是可选档，`preview` / `convert` / `av` / `python` / `web-verify` /
 * `web-live` 在本机全量回归里跑）。
 *
 * 守着这份名单的三处：
 *   1. `scripts/run-suites.mjs --ci`（工作流实际调用的入口，会核对每个名字都在 `ORDER`
 *      里、脚本真的存在，缺一个当场非零退出）；
 *   2. `test/ci-guard.mjs`（断言两份工作流都只写 `--ci`、名单包含四个核心格式回归、
 *      且不含任何环境相关套件）；
 *   3. `test/ci-guard.mjs` 自己也在名单里 —— 守门自己也要被守门。
 */

/** 跑在 CI / 发布守门里的套件名（按 `run-suites.mjs` 的 `ORDER` 口径书写）。 */
export const CI_SUITES = [
    // 记忆与知识库：契约最密、最容易一次改动踩坏别处的一档
    'smoke', 'memory', 'memory-archive-scope', 'memory-archive-integration', 'kb', 'kb-folders', 'kb-folders-scope', 'kb-folders-view', 'kb-scope', 'kb-db', 'kb-ann', 'kb-selector', 'kb-local-embed', 'kb-embed', 'kb-extract', 'memory-recall', 'memory-probes', 'citation', 'injection-budget',
    // 写作规划（第七十三轮）：writing 测 writingPlan/Audit/Signals 契约，writing-eval 测
    // A/B 评测脚本本身（--mock 全离线、真 HTTP 只打 127.0.0.1 假端点）。
    'writing', 'writing-eval',
    // 检索与派工：不联网的那几层（真联网的 web-live 在环境档里）
    'archive', 'sites', 'seam',
    // 格式与提示词约束：纯结构断言。`preview-view`（第六十八轮）测预览面板的三条只读
    // 路由：假 webServer + 假渲染引擎，不碰本机 LibreOffice，所以可以进 CI 档。
    'view', 'view-ocr', 'kb-scope-contrast', 'ocr', 'kb-llm', 'kb-ai-integration', 'gallery', 'gallery-model', 'gallery-routes', 'gallery-tool', 'gallery-mount', 'gallery-integration', 'image-asset', 'preview-view', 'plain-default', 'schema-format',
    // preset 声明文件的逐行回写（第六十三轮 新-21）：探针在系统临时目录里搭一份假
    // `$DSH_HOME`（假 profile + 假 preset），所以不读本机 profile、不联网 ——
    // 与 preset-check 的分工是：那一套读**真实的** profile 与 preset（只在本机跑）。
    'preset-plugins',
    // 生命周期契约：设置重建事务的失败形态（第五十九轮 S-04）。纯 Node、不读本机环境 ——
    // 断言的是「失败时回滚到哪一份」，与探到的引擎无关
    'rebuild',
    // 核心格式回归（S-06 点名补上的四个：Excel / PPT 的部件、关系与统计）
    'chart', 'format-excel', 'format-ppt', 'format-ppt-revise', 'format-ppt-math', 'format-word', 'image-export',
    'e2e',
    // 守门名单自己的回归
    'ci-guard',
];

/**
 * 明确**不许**进 CI 档的名字（两向核对用：既查「该在的在」，也查「不该在的不在」）。
 *
 * 第五十六轮之前正是没人做这一步核对，才让四个核心格式套件漏了十一轮。按**原因**分三组，
 * 免得下一轮把「因为它要 LibreOffice」和「因为它要出网」记成一回事：
 *   - 要本机引擎或本机安装：preview / convert / validate-libreoffice / validate-chart-com /
 *     pdf / tex / format-tex / python / av / av-words / image；
 *   - 要出网或要无头浏览器：web / web-proxy / web-live / web-verify / search；
 *   - 要宿主的包或本机环境：preset-check（读 `$DSH_HOME` 的 profile 与 preset bundle）、
 *     settings / client（`resolveHostPath` 解析 `@deepseek-ai/schemastery`、主题包）、
 *     capabilities（现探本机 PDF / TeX / Python / pandoc / 音视频引擎，结论随环境变）。
 */
export const CI_FORBIDDEN = [
    'preview', 'convert', 'pdf', 'python', 'av', 'av-words', 'image', 'tex', 'format-tex',
    'validate-libreoffice', 'validate-chart-com',
    'web', 'web-proxy', 'web-live', 'web-verify', 'search',
    'view-ocr-socket', 'kb-upload-socket', // 真实 loopback 监听仅在默认档运行
    'preset-check', 'settings', 'settings-save', 'client', 'capabilities',
];
