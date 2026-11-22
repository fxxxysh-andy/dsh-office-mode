# 审美参考图库

面向「找感觉」的图片库：保存原图、生成缩略图、用文字/标签/相似图找图。它不是 OCR，也不把图片里的文字当检索目标。

## 用户看到的能力

- **入库**：从会话工作目录导入 PNG / JPEG / GIF / WebP，原图按内容哈希保存，同一张图不会重复占空间。
- **看**：缩略图网格、大图预览（原图字节原样返回）。
- **描述**：可选让视觉模型生成主体 / 构图 / 配色 / 光线 / 材质 / 氛围与标签；只描述看得见的审美特征，**不推断作者、版权、出处**。模型如果返回 author/copyright 一类字段，直接丢弃，并告诉用户丢了哪些。
- **手工整理**：标签、分类、收藏、备注随时改。
- **搜**：文字搜图（描述 + 标签）、按标签/分类/收藏过滤、以图搜图；有图文对齐向量时做融合排序。
- **如实报告能力**：没配置视觉模型或多模态 embedding 时，回执写「未配置」，不假装成功，不拿文本向量冒充图片语义向量。

## 统一入口

工具层只有一个 gallery 动作（挂在 office_memory 的 action 表里），用 op 区分：

| op | 用途 |
| --- | --- |
| list | 最近入库，可按标签/分类/收藏过滤 |
| search | 文字 / 标签 / 以图搜图（query、图库图 ID `image`、tags） |
| add | 从工作目录导入一张图（path） |
| tag | 改标签、分类、收藏、备注 |
| describe | 让视觉模型写审美描述；必须显式 allowUpload: true |
| index | 准备检索向量；使用远端模型必须显式 allowUpload: true |
| status | 看当前能力、已配置项、未配置原因 |
| remove | 删除图片及缩略图、向量缓存 |

HTTP 端点统一在 /office-memory/gallery，用 query/body 分区（?view=image&id=…、POST {action:"add"} 等）。图片只按 64 位十六进制 id 取，路径由服务端拼，不接受任意路径。

## 配置（默认全关，不联网）

设置页只有一组「图库」：

- **视觉描述**：通道 off / host（宿主模型）/ openai（OpenAI 兼容端点），以及 provider、model、baseURL、apiKey。
- **图文对齐**：通道 off / voyage / cohere / vllm（OpenAI 兼容自建端点），model、baseURL、apiKey、维度、批量、超时。
- **文本向量**：沿用知识库的文本 embedding 配置；远端或本地进程不可用时退回词法并标 `degraded`，只有显式 `local` 且推理没开时才用本地哈希向量，并在回执里写明「只补字面召回，不是语义模型」。

默认只在本地检索，不上传图片或文本。视觉描述、探针、图片向量化、云端文字搜图和以图搜图须显式确认。远端文本向量可能发送搜索词、标签和描述；未确认时回退本地哈希。外部模型可能计费。

## 数据与缓存

- 图库根目录（默认在知识库目录同级 gallery/）：
  - originals/<sha256>.<ext> 原图
  - thumbs/<sha256>.jpg 缩略图
  - manifest.json 真源，facts / aesthetic / manual 三层分开
  - vectors/text/<key>/<id>.json 文本向量
  - vectors/image/<key>/<id>.json 图文对齐向量
  - vectors/probe-<space>.json 同维探针结果
- 向量 key 带模型、维度、版本：文本 txt_<model|dim|text-v1>，图片 img_<provider|baseURL|model|mm-v1>，互不混用。
- 带内容哈希与文本哈希：同图、同描述可复用缓存。
- 同一物理图库的清单写入经过进程内队列和跨进程独占锁。拿不到锁会超时，不绕过锁写入。若进程异常退出遗留 `manifest.lock`，先确认没有进程持有它，再手动移除；插件不会自动清理可能仍有效的锁。

## 检索怎么算分

硬过滤（标签/分类/收藏）之后，按可用信号加权：

- lexical：中文 bigram + IDF 覆盖度（描述、标签、备注）
- tag：标签精确匹配 1.0 / 包含 0.6
- textVector：查询文本向量 vs 描述向量（文本空间）
- imageText：查询文本的多模态向量 vs 图片向量（图文对齐空间）
- imageImage：查询图向量 vs 图片向量（同一对齐空间）

融合分 = 激活信号的加权平均。**不可用信号不会出现分项**；文本向量和图片向量永远不会互相比。每个结果都带分项分与信号说明，notes 会写明「分数是融合分，不代表语义能力已可用」。

有查询词或查询图时，**只有融合分 > 0 才算命中**：0 分表示所有激活信号都没命中，
负的图文余弦表示不相似，两者都不进入结果。无命中返回 `items: []`、`total: 0` 与
`reason: "no-match"`，面板显示空态和「清空搜索」。只有标签 / 分类 / 收藏这类硬过滤、
没有查询条件时，仍完整返回过滤后的图片（`score: null`）。

## 安全边界

- 导入路径先 realpath，必须落在信任根（会话工作目录、知识库目录、图库目录）内；不是普通文件、空文件、超过 20 MiB 都拒绝。
- 只按魔数认 PNG / JPEG / GIF / WebP；解码前检查像素总数（60M 上限）。
- 图片是非可信数据：只当字节保存与展示，不执行、不当 HTML。响应带 nosniff、CSP sandbox。
- 原图/缩略图只按 id 查找，不接受调用方给的文件路径；manifest 里的路径也必须等于由 id 推导的规范相对路径，篡改后直接 404。
- 不推断作者与版权；来源只记录调用方显式提供的字段。

## 真实能力边界

- 多模态 embedding 与视觉模型的网络行为**尚未端到端验证**：测试用假 fetch / 假宿主覆盖请求体、维度校验、未配置与未确认路径；本机没有 Key，也没有图文对齐模型。
- 默认状态：视觉通道 off、图片向量 off；此时 status 与 search 明确报「未配置」。
- 缩略图依赖本机 Python + Pillow；不可用时保留原图、不生成缩略图，不用原图冒充。
- 供应商自己的上限也按契约守住：Voyage 单图 20MB / 16M 像素，Cohere 单图 5MB / 单请求 20MB；超过时自动改用缩略图（JPEG）再向量化，没有缩略图就明确跳过并说明原因。
- 本地文本向量是 local-hash-v1（256 维），只是字面召回，不是语义模型。

## 测试

在 packages/dsh-office-mode 下：

```
node test/gallery.mjs
node test/gallery-model.mjs
node test/gallery-routes.mjs
node test/gallery-tool.mjs
node test/gallery-thumb.mjs   # 需要本机 Pillow；没有则 skip
```

## 代码位置

- src/gallery.js：存储、路径信任、标签/收藏、文本与图片向量缓存、检索融合
- src/gallery-model.js：视觉描述（宿主 / OpenAI 兼容）与多模态 embedding（Voyage / Cohere / vLLM）
- src/gallery-routes.js：HTTP 端点（注册在 /office-memory/gallery）
- src/gallery-tool.js：统一入口与文本回执
