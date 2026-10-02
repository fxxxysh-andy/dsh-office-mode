# dsh-office-mode（办公模式）

给 DeepSeek Harness 的办公插件与配套 preset：把 Word / Excel / PPT 的生成、读取、编辑、
排版、主题与原生图表，LaTeX 学位论文（thuthesis 模板）的源文件生成与 PDF 编译，PDF 解析
与渲染预览，Python 计算与绘图，音频与视频的内容提取（含词级时间戳），按内容类型分流的
联网检索与三步走检索编排，以及按工作目录存的三层记忆与知识库，收进「一次调用干完一批活」
的极简工具面里。插件保持零第三方依赖。

## 仓库结构

| 目录 | 是什么 |
| --- | --- |
| `packages/dsh-office-mode` | 插件本体（能力）：只注册 `office_help` / `office_run` / `office_memory` 等少数常用工具，其余能力按需取用 |
| `presets/office` | 「办公模式」preset bundle（说话方式 + 精简工具目录） |

## 安装

两条命令都在仓库根目录下执行。

```powershell
# 1) 插件（能力）：装进当前 profile，所有模式都能用
dsh plugin --profile web add ./packages/dsh-office-mode

# 2) 模式（说话方式 + 精简工具目录）：声明式 preset bundle
#    plugin_manager → install_bundle → ./presets/office
```

装完插件后，任意会话里 `office_help` 就能用；选择「办公模式」启动的会话才会吃到
persona 与精简工具目录。

## 文档

| 文档 | 内容 |
| --- | --- |
| [插件说明](packages/dsh-office-mode/README.md) | 安装、工具清单、三层记忆、检索链路、音视频、设置页与全部话题 |
| [格式规范](packages/dsh-office-mode/SPEC-formats.md) | Word / Excel / PPT 的字段全表与排版规则 |
| [OOXML 踩坑记录](packages/dsh-office-mode/docs/ooxml-pitfalls.md) | 真实 Office 打不开文件的那一类坑 |

## 测试

各套件都是独立脚本，在插件目录下直接跑：

```powershell
cd packages/dsh-office-mode
node test/smoke.mjs
```

其余套件见 `packages/dsh-office-mode/package.json` 的 `scripts`（memory / av / settings /
client / search / web / pdf / python / preset-check / injection-budget 等）。

## 许可

MIT，见 [LICENSE](LICENSE)。
