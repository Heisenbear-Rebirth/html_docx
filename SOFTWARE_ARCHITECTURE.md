# html_docx 软件架构设计

## 1. 架构目标

本项目的目标是实现一个面向 AI Agent 的 DOCX 可逆编辑系统：

```text
DOCX <-> H-DOCX bundle
```

系统必须满足：

- 未修改往返严格保真。
- 局部修改只影响目标 OOXML fragment。
- HTML 作为 Agent 可读可编辑投影。
- H-CSS 作为自定义 DOCX 样式编辑 DSL。
- manifest、原始 package、OOXML fragment 和 hash 作为事实来源。
- 不安全、不明确、不可证明的编辑必须拒绝。

本架构设计不是普通转换器架构，而是一个“DOCX package 保真 + 可编辑投影 + patch compiler + validation”的系统。

## 2. 顶层数据流

```text
input.docx
  -> package reader
  -> OPC inventory
  -> OOXML scanner
  -> manifest builder
  -> HTML projector
  -> H-CSS scaffold generator
  -> H-DOCX bundle

H-DOCX bundle
  -> HTML/H-CSS parser
  -> manifest validator
  -> edit intent detector
  -> H-CSS compiler
  -> OOXML patch planner
  -> safety validator
  -> fragment patcher
  -> DOCX repacker
  -> output validator
  -> output.docx
```

核心原则：

```text
export 阶段只建立可逆投影。
apply 阶段只执行可证明安全的 patch。
validate 阶段证明未修改内容未变化。
```

## 3. H-DOCX Bundle 格式

H-DOCX 不应是单个普通 HTML 文件，而应是一个目录或 zip 包。

### 3.1 目录格式

推荐开发期格式：

```text
document.hdocx/
  manifest.json
  document.html
  styles.generated.css
  agent.edits.hcss
  audit.log.jsonl
  original/
    original.docx
    entries.json
  parts/
    word/
      document.xml
      styles.xml
      numbering.xml
      settings.xml
      comments.xml
      footnotes.xml
      endnotes.xml
      header1.xml
      footer1.xml
    _rels/
  media/
    image1.png
  previews/
    image1.preview.png
```

说明：

- `manifest.json` 是事实索引，不是可选文件。
- `document.html` 是 Agent 主编辑面。
- `styles.generated.css` 是预览 CSS，可再生成，默认只读。
- `agent.edits.hcss` 是 Agent 修改排版的 H-CSS 脚本。
- `audit.log.jsonl` 记录每次导出、修改、回写、验证摘要。
- `original/original.docx` 保存原始 DOCX bytes，用于未修改直接返回和严格校验。
- `original/entries.json` 保存 ZIP entry 顺序、压缩方式、时间戳、CRC、大小等元信息。
- `parts/` 保存解析用的 part 副本，不能替代原始 bytes。
- `media/` 保存媒体文件副本和浏览器可用路径。
- `previews/` 保存不可直接预览媒体的预览图。

### 3.2 打包格式

发布期可以支持：

```text
document.hdocx.zip
```

但内部结构与目录格式一致。

### 3.3 单文件 HTML 模式

单文件模式理论可行，但 V1 不建议作为主格式。

如果以后支持，需要把 manifest 和 original DOCX 以 base64 嵌入 HTML：

```html
<script type="application/hdocx-manifest+json">...</script>
<script type="application/hdocx-package+base64">...</script>
```

缺点：

- 文件巨大。
- Agent 编辑风险高。
- diff 噪音大。
- 不利于分层验证。

V1 不实现单文件模式。

## 4. 核心模块

### 4.1 PackageReader

职责：

- 读取 `.docx` zip package。
- 枚举所有 ZIP entries。
- 保存原始 entry bytes 和 metadata。
- 解析 `[Content_Types].xml`。
- 解析 package-level 和 part-level relationships。

输入：

```text
input.docx
```

输出：

```text
PackageInventory
RawEntryStore
RelationshipGraph
```

关键要求：

- 不修改原始 bytes。
- 不规范化路径。
- 保留外部关系 `TargetMode="External"`。
- 识别加密或损坏 package 并拒绝。

### 4.2 OOXMLScanner

职责：

- 扫描 WordprocessingML parts。
- 识别 document、styles、numbering、settings、headers、footers、footnotes、endnotes、comments、commentsExtended、theme、customXml 等。
- 建立 paragraph/run/table/range/object/style/field/revision/comment 索引。
- 标记 editable/protected/opaque 节点。

输出：

```text
PartModel[]
NodeIndex
RangeMap
StyleMap
NumberingMap
ObjectMap
ProtectedMap
```

关键要求：

- 解析用于定位和投影，不承担全量重写。
- 未知 namespace 和扩展结构进入 opaque/protected map。
- 复杂 range 使用 start/end anchor 模型。

### 4.3 ManifestBuilder

职责：

- 构建 `manifest.json`。
- 给节点分配稳定 `data-hdocx-id`。
- 记录 part、locator、hash、lock、source fragment、parent/child、range relation。
- 记录 original package entry metadata。

输出：

```text
manifest.json
```

关键要求：

- manifest 必须足够支撑回写校验。
- ID 必须稳定且唯一。
- hash 必须覆盖未修改检测需要的 fragment。

### 4.4 HTMLProjector

职责：

- 将 PartModel 投影成 `document.html`。
- 输出 Agent 可读结构。
- 对 main document、header、footer、footnote、endnote、comment 等 part 生成独立 article。
- 为 protected 结构生成可见或可定位占位。

输出：

```text
document.html
```

关键要求：

- `<p>` 对应 `w:p`。
- `<span data-hdocx-type="run">` 对应 `w:r`。
- 表格使用 `<table>/<tr>/<td>` 投影。
- 标题仍用 `<p>`，通过 `data-hdocx-style-id`、`data-hdocx-outline-level`、ARIA 表达。
- 列表仍用 `<p>`，通过 `data-hdocx-num-id`、`data-hdocx-ilvl` 表达。
- 重叠范围使用 boundary anchors，不强行 wrapper。

### 4.5 PreviewStyleGenerator

职责：

- 根据 manifest、styles.xml、numbering.xml、theme 生成 `styles.generated.css`。
- 提供浏览器预览效果。

输出：

```text
styles.generated.css
```

关键要求：

- 只用于预览。
- 可再生成。
- 不作为 OOXML 回写事实来源。

### 4.6 HCSSParser

职责：

- 解析 `agent.edits.hcss`。
- 支持 `@hdocx-token`、`@hdocx-set`、`@hdocx-format`、`@hdocx-include`、`@hdocx-range`、`@hdocx-edit mode(...)`。
- 进行语法、类型、单位、冲突检查。

输出：

```text
HCSSProgram
EditRequests
SelectorMatchReport
```

关键要求：

- H-CSS 是自定义 DSL，不是浏览器 CSS。
- 0 匹配默认报错。
- protected 命中默认报错。
- 属性层级错配默认报错。
- 冲突默认报错。

### 4.7 HTMLChangeDetector

职责：

- 比较当前 `document.html` 与 manifest 记录的导出状态。
- 检测 Agent 对文本、属性、节点结构的直接修改。
- 将允许的 HTML 修改转成 EditIntent。

输出：

```text
EditIntent[]
HTMLChangeReport
```

关键要求：

- ID 被修改则失败。
- protected node 被修改则失败。
- 未声明新增节点需要进入 insert flow。
- 删除节点需要确认是否包含 protected/range/object。

### 4.8 EditIntentResolver

职责：

- 合并 HTML 直接修改和 H-CSS 修改。
- 将用户编辑意图归一化为明确 patch 目标。

输出：

```text
ResolvedEdit[]
ConflictReport
```

关键要求：

- 同一目标同一属性冲突必须失败。
- style-definition 与 direct-formatting 不得混淆。
- 段落选择器设置 run 属性必须有明确 mode。
- run 选择器设置段落属性默认失败。

### 4.9 PatchPlanner

职责：

- 把 ResolvedEdit 编译成 OOXML patch plan。
- 决定修改 part、fragment、节点、属性和插入位置。
- 规划 run split、range split、paragraph patch、style patch。

输出：

```text
PatchPlan
```

关键要求：

- 只修改必要 fragment。
- 新 run 继承原 run 未修改属性。
- run split 不能切坏复杂边界。
- 修改 style definition 必须生成影响摘要。

### 4.10 SafetyValidator

职责：

- 在写入前验证 PatchPlan 安全性。

检查：

- manifest hash 是否匹配。
- protected 节点是否被修改。
- range boundary 是否完整。
- run split 是否落在 grapheme cluster 边界。
- field/comment/bookmark/revision/content control 是否被破坏。
- relationship 是否冲突。
- 修改是否违反文档保护。

输出：

```text
SafetyReport
```

关键要求：

- 失败优先。
- 输出明确错误原因和定位。

### 4.11 FragmentPatcher

职责：

- 对被修改 XML part 进行 fragment-level patch。
- 未修改 fragment 保留原始 bytes。
- 尽量避免整份 DOM 重写。

输出：

```text
ModifiedPart[]
PatchAudit
```

关键要求：

- 保留 namespace prefix、属性顺序、whitespace、comments、processing instructions。
- 不折叠 `mc:AlternateContent`。
- 不删除 unknown namespace。

实现上可以分阶段：

V1 允许对被修改的简单 XML part 使用保守序列化，但必须在验证中证明未修改子树 hash 不变。后续再升级到 offset-based fragment patch。

### 4.12 DocxRepacker

职责：

- 重新打包 output.docx。
- 未修改 entry 使用原始 bytes 和 metadata。
- 修改 entry 使用 FragmentPatcher 输出。
- 必要时更新 content types 和 relationships。

输出：

```text
output.docx
```

关键要求：

- 未修改往返时直接返回 original.docx bytes。
- 局部修改时尽量保留 entry 顺序和 metadata。
- 不无故改变 relationship id。

### 4.13 OutputValidator

职责：

- 验证输出 DOCX。
- 执行 package diff、part hash、relationship graph、protected structure、WordprocessingML 基本结构检查。

输出：

```text
ValidationReport
```

关键要求：

- 未修改往返必须完全等价。
- 局部修改必须只出现预期 diff。
- 输出必须可被 DOCX 解析器重新读取。

## 5. 核心数据模型

以下为概念模型，具体实现语言确定后再落成类型定义。

### 5.1 PackageInventory

```json
{
  "packageId": "pkg-001",
  "entries": [
    {
      "path": "word/document.xml",
      "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      "relationshipPart": "word/_rels/document.xml.rels",
      "sha256": "...",
      "compressedSize": 1234,
      "uncompressedSize": 5678,
      "compressionMethod": "deflate",
      "zipOrder": 12
    }
  ]
}
```

### 5.2 PartModel

```json
{
  "partPath": "/word/document.xml",
  "partKind": "main-document",
  "hash": "...",
  "nodes": ["p-0001", "p-0002"],
  "relationships": ["rId1", "rId2"]
}
```

### 5.3 NodeRecord

```json
{
  "id": "r-00042",
  "kind": "run",
  "partPath": "/word/document.xml",
  "locator": {
    "strategy": "structural-path",
    "path": ["body", 0, "p", 12, "r", 3]
  },
  "hash": "...",
  "lock": "editable",
  "parent": "p-00012",
  "children": [],
  "styleRefs": [],
  "rangeRefs": [],
  "sourceFragmentRef": "frag-abc123"
}
```

### 5.4 FragmentRecord

```json
{
  "id": "frag-abc123",
  "partPath": "/word/document.xml",
  "byteRange": [1024, 1180],
  "xmlKind": "w:r",
  "sha256": "...",
  "protected": false
}
```

V1 如果暂不实现 byte offset patch，可以先使用 structural locator + canonical subtree hash，但文档中仍保留 byteRange 目标。

### 5.5 RangeRecord

```json
{
  "id": "cmt-3",
  "kind": "comment",
  "startAnchor": "anchor-cmt-3-start",
  "endAnchor": "anchor-cmt-3-end",
  "partPath": "/word/document.xml",
  "crossingRanges": ["bm-1"],
  "lock": "protected-boundary"
}
```

### 5.6 StyleRecord

```json
{
  "styleId": "Normal",
  "type": "paragraph",
  "name": "Normal",
  "basedOn": null,
  "next": "Normal",
  "linked": null,
  "sourcePart": "/word/styles.xml",
  "hash": "..."
}
```

### 5.7 EditIntent

```json
{
  "id": "edit-0001",
  "source": "hcss",
  "target": {
    "kind": "set",
    "name": "body-paragraphs"
  },
  "mode": "style-definition",
  "properties": {
    "hdocx-font-size": "12pt",
    "hdocx-line-spacing": 1.5
  }
}
```

### 5.8 PatchOperation

```json
{
  "id": "patch-0001",
  "partPath": "/word/styles.xml",
  "targetNode": "style-Normal",
  "operation": "set-run-property",
  "property": "w:sz",
  "value": "24",
  "expectedOldHash": "..."
}
```

## 6. Manifest 结构

`manifest.json` 顶层建议：

```json
{
  "hdocxVersion": "0.1",
  "createdBy": "html_docx",
  "sourceDocx": {
    "fileName": "input.docx",
    "sha256": "...",
    "size": 123456
  },
  "package": {},
  "parts": {},
  "relationships": {},
  "nodes": {},
  "fragments": {},
  "ranges": {},
  "styles": {},
  "numbering": {},
  "media": {},
  "protected": {},
  "generatedClasses": {},
  "validation": {
    "exportHash": "...",
    "roundtripPolicy": "original-bytes-if-unmodified"
  }
}
```

要求：

- manifest 是回写的硬依赖。
- manifest 不应被 Agent 随意编辑。
- manifest 被破坏时不能输出 DOCX。

## 7. CLI 设计

V1 优先做 CLI，方便 Codex/Claude Code 调用。

### 7.1 export

```text
html-docx export input.docx --out document.hdocx
```

作用：

- DOCX -> H-DOCX bundle。
- 生成 HTML、manifest、preview CSS、空 H-CSS。

输出：

```text
document.hdocx/
```

### 7.2 validate

```text
html-docx validate document.hdocx
```

作用：

- 校验 bundle 完整性。
- 校验 manifest 与 HTML 一致。
- 校验 H-CSS 语法和选择器匹配。
- 不输出 DOCX。

### 7.3 plan

```text
html-docx plan document.hdocx --report plan.json
```

作用：

- 解析 HTML/H-CSS 修改。
- 输出将要修改哪些 part、节点、属性。
- 输出选择器命中数量和 style-definition 影响摘要。
- 不写 DOCX。

### 7.4 apply

```text
html-docx apply document.hdocx --out output.docx
```

作用：

- 执行 validate。
- 生成 PatchPlan。
- 写 output.docx。
- 执行输出校验。

### 7.5 diff

```text
html-docx diff input.docx output.docx --report diff.json
```

作用：

- 比较两个 DOCX package。
- 输出 part-level、relationship-level、XML fragment-level diff。

### 7.6 inspect

```text
html-docx inspect document.hdocx --id p-00042
```

作用：

- 查看某个 H-DOCX 节点对应的 part、XML fragment、style、range、lock 状态。

### 7.7 roundtrip

```text
html-docx roundtrip input.docx --work work.hdocx --out roundtrip.docx
```

作用：

- 一键执行 export + apply。
- 用于未修改往返测试。

## 8. API 设计

后续可以提供库 API。

概念接口：

```text
export_docx(input_docx_path, output_hdocx_dir) -> ExportReport
validate_hdocx(hdocx_dir) -> ValidationReport
plan_hdocx(hdocx_dir) -> PatchPlan
apply_hdocx(hdocx_dir, output_docx_path) -> ApplyReport
diff_docx(left_docx_path, right_docx_path) -> DiffReport
inspect_node(hdocx_dir, node_id) -> NodeInspection
```

API 返回对象必须机器可读，方便 Agent 决策。

## 9. 回写策略

### 9.1 未修改路径

如果：

- `document.html` 未修改；
- `agent.edits.hcss` 无有效 edit；
- manifest 与 original hash 匹配；

则：

```text
直接复制 original/original.docx 到 output.docx
```

这是未修改严格双射的最强策略。

### 9.2 简单文本修改

条件：

- 修改目标是 editable run text。
- run 不包含 protected 子节点。
- 修改不移动 range boundary。

策略：

- patch 该 run 的 `w:t`。
- 必要时更新 `xml:space="preserve"`。
- 未修改 siblings 不动。

### 9.3 局部字符格式修改

条件：

- 修改范围在普通文本内。
- 可安全 run split。

策略：

- 拆分 run。
- 新 run 继承原 run properties。
- 对目标 run 添加或修改 `w:rPr`。

### 9.4 段落格式修改

条件：

- 目标是 paragraph。
- 属性属于 paragraph-level。

策略：

- patch `w:pPr`。
- 不触碰 run。

### 9.5 样式定义修改

条件：

- H-CSS 使用 `mode(style-definition)`。
- styleId 唯一。
- 用户确认或接受影响摘要。

策略：

- patch `word/styles.xml` 中对应 `w:style`。
- 输出影响摘要。

### 9.6 表格修改

V1 先支持：

- 单元格文本修改。
- 单元格内 run/paragraph 格式修改。
- 简单单元格底纹、边框、对齐。

暂缓：

- 复杂合并结构编辑。
- tblGrid 重建。
- 自动调整列宽。

### 9.7 protected 修改

默认拒绝。

包括：

- field。
- equation。
- revision。
- SmartArt。
- chart。
- OLE。
- unknown opaque fragment。
- signature-related part。

## 10. 验证链

### 10.1 Export 验证

检查：

- package 可读。
- manifest 完整。
- 所有 part hash 已记录。
- 所有节点 ID 唯一。
- HTML 中所有 `data-hdocx-id` 存在于 manifest。
- protected 节点已标记。

### 10.2 Pre-apply 验证

检查：

- manifest 未缺失。
- original DOCX 存在且 hash 匹配。
- HTML ID 未被篡改。
- protected 节点未被修改。
- H-CSS 合法。
- 选择器命中数量可接受。
- 没有属性冲突。
- PatchPlan 安全。

### 10.3 Output 验证

检查：

- output.docx 是合法 ZIP/OPC package。
- 所有应存在 entries 存在。
- relationships 图一致或只发生预期变化。
- 未修改 parts hash 不变。
- 修改 parts 只有预期 fragment 变化。
- media hash 不变，除非用户替换。
- styles/numbering/settings/comments/revisions/footnotes/endnotes 未被意外改变。

### 10.4 可选 Word 打开验证

在有可用环境时，可以增加：

- 用 LibreOffice 或 Word 自动打开并导出 PDF。
- 用 Open XML SDK validator 验证。
- 用渲染截图做视觉对比。

V1 不应把这些作为唯一验证，因为严格双射主要靠 package/OOXML diff 证明。

## 11. 错误模型

错误必须机器可读。

示例：

```json
{
  "code": "HCSS_SELECTOR_MATCHED_PROTECTED_NODE",
  "severity": "error",
  "message": "Selector body-paragraphs matched protected field node fld-0004.",
  "location": {
    "file": "agent.edits.hcss",
    "line": 12,
    "nodeId": "fld-0004",
    "partPath": "/word/document.xml"
  },
  "suggestion": "Exclude [data-hdocx-lock=\"protected\"] or use a narrower selector."
}
```

错误分类：

- `PACKAGE_*`
- `MANIFEST_*`
- `HTML_*`
- `HCSS_*`
- `SELECTOR_*`
- `PATCH_*`
- `PROTECTED_*`
- `RANGE_*`
- `VALIDATION_*`
- `SECURITY_*`

## 12. 技术实现建议

### 12.1 语言选择

推荐先用 Python 实现 V1。

理由：

- 标准库支持 zip/xml。
- 便于快速构建 CLI。
- Agent 生态友好。
- 文档处理和测试工具丰富。

约束：

- 所有依赖必须安装在项目内虚拟环境。
- 不做全局 pip install。
- 不写当前目录外缓存或配置。

### 12.2 XML 处理

V1 可组合使用：

- `zipfile` 读取和打包。
- `lxml` 或标准库 XML 解析。
- 自定义 namespace-preserving fragment patch 策略。

注意：

- 如果使用 `lxml`，必须在项目内 `.venv` 安装。
- 不能依赖全局 Python 包。

### 12.3 H-CSS 解析

V1 可以先实现一个小型解析器：

- 只支持必要 at-rules。
- 只支持属性声明。
- 不支持任意 CSS。
- 不支持复杂 cascade。

后续可引入 parser generator，但仍应保持 H-CSS 自定义语义。

### 12.4 Hash 策略

推荐：

- package entry bytes hash。
- XML part bytes hash。
- fragment canonical hash。
- visible text hash。
- protected node hash。

不同 hash 用途不同，不能混用。

## 13. 测试架构

### 13.1 Fixture 分类

```text
tests/fixtures/
  basic/
    simple_paragraph.docx
    mixed_runs.docx
    table_basic.docx
  academic/
    headings_toc_refs.docx
    footnotes_endnotes.docx
    captions_crossrefs.docx
    equations.docx
  review/
    comments.docx
    revisions.docx
  objects/
    images.docx
    charts.docx
    smartart.docx
    ole.docx
  edge/
    alternate_content.docx
    custom_xml.docx
    bidi_east_asian.docx
    nested_ranges.docx
```

### 13.2 测试类型

未修改往返：

```text
docx -> hdocx -> docx
```

局部文本修改：

```text
modify one run text -> only target run fragment changes
```

局部格式修改：

```text
format one word -> run split + target rPr change only
```

H-CSS 修改：

```text
style-definition/direct-formatting/range-formatting
```

失败测试：

```text
protected edit
ID tamper
selector zero match
selector protected match
range boundary broken
conflicting H-CSS rules
```

### 13.3 Golden Report

每个测试输出：

```text
export_report.json
plan_report.json
apply_report.json
diff_report.json
validation_report.json
```

这样 Agent 可以直接读报告判断问题。

## 14. MVP 开发顺序

### Phase 0: 项目骨架

目标：

- 建立 Python 包和 CLI。
- 建立项目内 `.venv` 使用说明。
- 建立基础测试框架。

产物：

- `src/html_docx/`
- `tests/`
- `pyproject.toml`
- CLI 可运行。

### Phase 1: Package roundtrip

目标：

- 读取 DOCX。
- 生成 H-DOCX bundle。
- 未修改 apply 直接返回 original DOCX。
- package hash 验证通过。

成功标准：

- `roundtrip` 对任意未加密 DOCX 返回原始 bytes。

### Phase 2: Basic document projection

目标：

- 投影 main document 中 paragraph/run/text。
- 生成 `document.html` 和 manifest。
- 支持普通文本修改回写。

成功标准：

- 修改一个普通 run 文本，只改变对应 `w:t`。

### Phase 3: Run properties and paragraph properties

目标：

- 支持常见字符格式。
- 支持常见段落格式。
- 支持 run split。

成功标准：

- 一个段落内多个字号。
- 一个词加粗。
- 半句改颜色。
- 非目标 fragment 不变。

### Phase 4: H-CSS V1

目标：

- 支持 token/set/format/include/edit mode。
- 支持选择器匹配和冲突检查。
- 支持 style-definition 和 direct-formatting。

成功标准：

- Agent 能用 H-CSS 批量修改正文字号、行距、缩进。

### Phase 5: Tables and secondary parts

目标：

- 支持表格基本投影和单元格内编辑。
- 支持 header/footer/footnote/endnote article。

成功标准：

- 页眉、脚注、表格单元格内文本和局部格式可安全修改。

### Phase 6: Protected academic structures

目标：

- field、equation、comment、revision、drawing、chart、SmartArt、OLE 等 protected 占位。
- 防止 Agent 误改。

成功标准：

- 含目录、批注、修订、公式、图片的文档未修改严格保真。
- 尝试修改 protected 节点时报错。

### Phase 7: Deep validation and diff

目标：

- 完整 diff 报告。
- 更精确 fragment patch。
- 更多 fixture。

成功标准：

- 可以证明局部修改的影响范围。

## 15. 风险与对应策略

| 风险 | 策略 |
|---|---|
| XML 重新序列化造成无关 diff | fragment patch、未修改 bytes 原样回填 |
| H-CSS 选择器误伤 | dry-run、命中统计、protected 报错 |
| run split 破坏复杂边界 | safety validator、grapheme 边界检查 |
| style-definition 影响范围过大 | plan 阶段输出影响摘要 |
| 未知 OOXML 扩展丢失 | opaque fragment 保真 |
| Agent 篡改 manifest/ID | hash 和 ID 校验失败 |
| Word 修订/批注结构复杂 | V1 只读保真，编辑后置 |
| 数字签名/加密 | 拒绝或明确确认 |

## 16. V1 不做的事

V1 不承诺：

- 单文件 HTML 双射。
- 自动重算目录、交叉引用、文献域。
- 编辑修订链。
- 语义编辑公式、图表、SmartArt、OLE。
- 完整 Word 渲染一致性。
- 任意 CSS 到 DOCX。
- 任意 HTML 到 DOCX。
- 支持加密 DOCX。

V1 要承诺：

- 未修改严格保真。
- 普通正文局部文本/格式精确修改。
- 批量基础排版修改。
- 学术关键结构只读保真。
- 不安全编辑拒绝。

## 17. 文档关系

当前架构依赖以下设计文档：

- `FUNCTIONAL_SPEC.md`：功能边界和严格双射原则。
- `HDOCX_HTML_DESIGN.md`：HTML/H-CSS 表示设计。
- `EDGE_CASE_TEST_MATRIX.md`：DOCX 结构边缘性测试。
- `EDITING_EDGE_CASES.md`：微观编辑语义边界。
- `SELECTOR_AND_REUSE_DESIGN.md`：选择器、集合、格式复用设计。
- `SELECTOR_EDGE_CASES_AND_GUARDS.md`：选择器复用系统护栏。

本文件负责把这些原则落成软件模块和开发顺序。

## 18. 最终架构原则

```text
H-DOCX 是 bundle，不是普通 HTML。
manifest 是事实索引，不是注释。
HTML 是编辑投影，不是完整 DOCX。
H-CSS 是 OOXML patch DSL，不是网页 CSS。
Patch 必须最小化。
未修改必须原样保真。
复杂结构默认 protected。
不能证明安全就拒绝。
所有行为必须可审计。
```
