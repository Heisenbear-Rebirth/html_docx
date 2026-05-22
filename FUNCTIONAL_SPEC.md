# html_docx 功能性描述

## 1. 项目定位

本项目开发一套面向 AI Agent 的 DOCX 可逆编辑系统，使 Codex、Claude Code 等代码型 Agent 能够通过类似 HTML 的可读结构安全修改 DOCX 文档，同时保证 DOCX 与项目定义的增强 HTML 表示之间严格双射。

项目目标不是普通的 `docx -> html` 导出，也不是普通的 `html -> docx` 生成，而是建立：

```text
DOCX <-> H-DOCX
```

当用户需要从零创建新文档时，系统采用内置 canonical DOCX 模板生成新的 DOCX，
然后仍然进入同一条 `DOCX <-> H-DOCX` 可逆链路。Agent 不应手写任意 OOXML 包。

其中 `H-DOCX` 是本项目定义的增强 HTML 文档表示。它包含：

- Agent 可读、可编辑的 HTML 表层。
- 记录 DOCX 原始结构、关系、资源、样式、编号、XML 子树和校验信息的 manifest。
- 原始 OPC package 中所有需要保真的 parts、relationships、media、custom XML 和扩展数据。

因此，本项目中的“HTML”不是普通浏览器 HTML，而是面向 DOCX 的可逆编辑层。

## 2. 严格双射原则

本项目必须满足严格双射，而不是“尽量保持格式”。

严格双射定义如下：

1. 对任意受支持 DOCX，执行 `docx -> h-docx -> docx`，在用户未修改 H-DOCX 的情况下，输出文档必须与输入文档在规范化后完全等价。
2. 当用户或 Agent 只修改 H-DOCX 中某个声明为可编辑的节点时，回写 DOCX 时只允许对应 OOXML 子树发生变化。
3. 所有未修改的 OOXML parts、relationships、styles、numbering、media、headers、footers、footnotes、endnotes、comments、revisions、customXml、metadata 和未知扩展必须原样保留。
4. 如果某个功能无法被语义编辑，也必须通过受保护节点和原始 XML/资源保真机制完整保留。
5. 如果某个输入内容既无法编辑也无法保真，转换必须失败并给出明确错误，不能静默丢弃或降级。
6. 如果 H-DOCX 中的受保护节点、manifest、hash、relationship id 或结构锚点被破坏，回写 DOCX 必须失败。

## 3. 功能分级

所有 Word/DOCX 功能在本项目中分为四类。

| 级别 | 名称 | 含义 |
|---|---|---|
| A | 可编辑且严格保真 | 在 H-DOCX 中有语义表示，Agent 可以修改，回写 DOCX 时只改变目标结构 |
| B | 只读且严格保真 | 在 H-DOCX 中可见或以占位符表示，但默认禁止编辑，原始 XML/资源完整保留 |
| C | 完全透明保真 | Agent 通常不需要看到，系统在 manifest/package 层完整保留 |
| D | 拒绝处理 | 无法保证安全双射时直接报错，不产生损坏文档 |

“暂不支持编辑”不能理解为“可以丢弃”。只要文档能够进入本系统，除 D 类外所有内容都必须严格保真。

## 4. 学术写作核心功能

学术写作场景是本项目的核心场景。以下三组功能必须作为一等功能对待，不能作为可选兼容项。

### 4.1 基础排版

基础排版必须严格保真，并在 V1 中提供主要编辑能力。

包括：

- 字体、中文字体、西文字体、字号。
- 加粗、斜体、下划线、删除线、上下标、小型大写、颜色、高亮。
- 段落样式、字符样式、标题样式。
- 左对齐、居中、右对齐、两端对齐。
- 首行缩进、悬挂缩进、左右缩进。
- 段前、段后、行距、分页控制。
- 项目符号、多级编号、标题编号。
- 表格结构、表格样式、边框、底纹、合并单元格。
- 图片、题注、图片大小和位置。
- 页眉、页脚、页码。
- 页面大小、方向、页边距、分节。

### 4.2 学术结构

学术结构必须严格保真。V1 可以先支持受控编辑，但不能丢失或破坏。

包括：

- 目录 TOC。
- 题注。
- 交叉引用。
- 书签。
- 超链接。
- 脚注、尾注。
- 公式。
- 图表编号。
- 页码域。
- 文档属性域。
- 引文和参考文献相关域。
- 索引、图表目录、表格目录。

域类内容在 V1 中默认不主动重算。系统应保留原始 field code、field result 和相关 XML。若用户修改了会影响域结果的内容，应标记“需要 Word 更新域”，而不是自行猜测重算。

### 4.3 协作与审阅信息

批注和修订对论文指导、期刊返修、毕业论文审阅非常关键，必须严格保真。

包括：

- 批注内容。
- 批注范围。
- 批注作者、时间、id。
- 现代 threaded comments。
- 插入修订。
- 删除修订。
- 移动修订。
- 格式修订。
- 表格修订。
- 修订作者、时间、id。

V1 中批注和修订默认作为只读保真功能处理。后续可以增加受控编辑能力，但必须保证不会破坏修订链和批注范围。

## 5. Word/DOCX 功能边界

### 5.1 OPC 包结构

决策：C 类，完全透明保真。

包括：

- `[Content_Types].xml`
- package-level relationships
- part-level relationships
- part name
- content type
- media 资源
- embedded package
- custom XML part

要求：所有未修改 part 必须按原始内容回写，不能因解析和序列化导致无关变化。

### 5.2 主文档结构

决策：A 类。

包括：

- document body
- section
- paragraph
- run
- text
- tab
- line break
- page break
- column break
- bookmark
- hyperlink

要求：H-DOCX 中每个主要结构节点必须有稳定 id、源 XML 定位信息和 hash。

### 5.3 字符格式

决策：A 类核心编辑，未知属性保真。

包括：

- run style
- font family
- East Asian font
- ASCII font
- complex script font
- font size
- bold
- italic
- underline
- color
- highlight
- strike
- double strike
- superscript
- subscript
- character spacing
- hidden text
- language
- RTL

要求：Agent 修改字符格式时只改变对应 run properties。

### 5.4 段落格式

决策：A 类核心编辑，未知属性保真。

包括：

- paragraph style
- alignment
- indentation
- first line indent
- hanging indent
- spacing before/after
- line spacing
- keep with next
- keep lines
- page break before
- widow control
- outline level
- tabs
- paragraph border
- paragraph shading
- text direction

要求：段落属性必须与具体 paragraph 绑定，不能只转成 CSS 后丢失 OOXML 属性。

V1 H-CSS 必须明确声明当前可编辑的格式子集，并在 `plan` 阶段解释：

- 字符格式：`font-family`、`ascii-font`、`hansi-font`、`eastAsia-font`/`east-asia-font`、`cs-font`、`font-size`、`bold`、`italic`、`color`。
- 段落格式：`text-align`/`align`、`first-line-indent`、`line-spacing`、`line-spacing-exact`、`space-before`、`space-after`。
- `plan` 输出必须包含选择器命中节点、声明是否支持、规范化值、OOXML 映射、将生成的 patch id，以及不支持声明的行号和原因。

### 5.5 样式系统

决策：A/B 类。应用样式属于 A 类；复杂样式定义修改初期属于 B 类。

包括：

- paragraph style
- character style
- table style
- numbering style
- basedOn
- next style
- linked style
- latent styles
- theme binding

要求：V1 支持应用已有样式和修改常见样式属性。复杂继承关系和主题关系必须保真。

### 5.6 编号与项目符号

决策：A 类常规编辑，复杂编号保真。

包括：

- abstract numbering
- concrete numbering
- level
- start value
- level text
- bullet font
- numbering indentation
- picture bullet
- restart numbering

要求：多级标题编号和论文常用列表编号必须稳定往返。

### 5.7 表格

决策：A 类核心编辑。

包括：

- table
- row
- cell
- grid
- width
- border
- shading
- cell margin
- vertical alignment
- horizontal alignment
- merged cell
- repeated header row
- nested table
- table style

要求：单元格内容和常用表格排版可编辑；复杂表格属性必须保真。

### 5.8 页面布局

决策：A/B 类。常见页面设置可编辑，复杂设置保真。

包括：

- section properties
- page size
- orientation
- margins
- columns
- page border
- page background
- watermark
- different first page
- odd/even headers and footers

要求：分节是页面布局的关键边界，不能在 HTML 表示中被压平。

### 5.9 页眉页脚

决策：A 类文本编辑，复杂对象保真。

包括：

- default header/footer
- first page header/footer
- odd/even header/footer
- page number field
- image in header/footer
- table in header/footer

要求：页眉页脚必须作为独立 part 映射到 H-DOCX，不可只作为正文附属信息。

### 5.10 脚注尾注

决策：A 类受控编辑。

包括：

- footnote
- endnote
- note reference
- note body
- numbering format
- separator

要求：脚注尾注在学术写作中必须严格保真。V1 至少支持正文内容编辑和引用关系保真。

### 5.11 域与引用

决策：B 类为主，严格保真。

包括：

- TOC
- page number
- cross-reference
- caption
- bibliography field
- date field
- document property field
- mail merge field
- index

要求：V1 不主动重算域。若相关内容变化，应保留 field code 并标记需要更新域。

### 5.12 批注

决策：B 类，严格保真。

包括：

- comment range start/end
- comment reference
- comment body
- author
- date
- initials
- threaded comment

要求：批注范围必须稳定，不能因正文编辑错位。

### 5.13 修订

决策：B 类，严格保真。

包括：

- insertion
- deletion
- move from/to
- formatting change
- table revision
- revision author
- revision date
- revision id

要求：V1 默认不编辑修订结构。对含修订文档进行正文编辑时，必须避免破坏原修订链。

### 5.14 图片与媒体

决策：A/B 类。

包括：

- image relationship
- inline image
- floating image
- crop
- size
- position
- wrapping
- alt text
- external image link

要求：V1 支持图片保真、alt text 修改、尺寸修改和受控替换。复杂定位和环绕必须保真。

### 5.15 绘图对象

决策：B 类，严格保真。

包括：

- DrawingML shape
- VML shape
- text box
- WordArt
- icon
- SmartArt
- 3D model

要求：默认以受保护对象呈现，原始 XML 和资源关系完整保留。

### 5.16 图表

决策：B 类，严格保真。

包括：

- chart part
- embedded workbook
- chart style
- chart data relationship

要求：V1 不编辑图表数据，必须保留图表及其嵌入数据。

### 5.17 公式

决策：B 类起步，后续可扩展为 A 类。

包括：

- OMML
- equation layout
- equation numbering relationship

要求：公式是学术写作核心功能。V1 必须严格保真，可先以受保护节点呈现。后续可以增加 OMML 与可编辑数学表示之间的受控映射。

### 5.18 内容控件与表单

决策：B 类，简单文本控件后续可进入 A 类。

包括：

- SDT
- rich text control
- plain text control
- checkbox
- dropdown
- date picker
- XML mapping
- legacy form field

要求：数据绑定和控件属性必须保真。

### 5.19 自定义 XML

决策：C 类，完全透明保真。

包括：

- customXml parts
- XML schema
- data binding

要求：未被明确编辑时必须完全原样回写。

### 5.20 东亚排版

决策：A/B 类，严格保真。

包括：

- 拼音指南
- 着重号
- 纵排
- 双行合一
- 字符网格
- 禁则
- 中文字体
- 中文字号

要求：中文学术文档必须保留东亚排版属性。常用字体、字号、段落属性进入 A 类。

### 5.21 保护与权限

决策：B/D 类。

包括：

- read-only recommendation
- restrict editing
- form protection
- password hash

要求：保护设置必须保真。如果用户请求的编辑会违反保护语义，系统应拒绝或要求用户明确解除保护。

### 5.22 签名与加密

决策：D 类为主。

包括：

- digital signature
- encrypted OOXML

要求：加密 DOCX 初期拒绝处理。带数字签名的文档可以读取和保真，但任何编辑都会使签名失效，必须明确拒绝或要求用户确认签名失效。

### 5.23 宏与 ActiveX

决策：D 类。

包括：

- DOCM
- VBA project
- ActiveX control

要求：本项目面向 DOCX，不执行、不编辑宏。遇到宏或 ActiveX 应拒绝或只在明确安全策略下做透明保真。

### 5.24 嵌入对象

决策：B/C 类，严格保真。

包括：

- OLE object
- embedded Excel
- embedded PDF
- embedded Visio
- package object

要求：对象二进制和关系必须保真，V1 不编辑。

### 5.25 兼容性与扩展

决策：C 类，完全透明保真。

包括：

- `mc:Ignorable`
- `mc:AlternateContent`
- unknown namespace
- Transitional OOXML
- legacy VML
- application-specific extension

要求：未知内容默认原样保留，不得解析后重写导致信息损失。

### 5.26 应用级功能

决策：不属于文件双射目标。

包括：

- Word 自动保存
- 云端版本历史
- 实时协作状态
- 拼写语法建议 UI
- 翻译
- 听写
- 阅读模式
- Word 比较文档命令

要求：这些不是 DOCX 文件内稳定内容，项目不主动处理。

## 6. H-DOCX 表示要求

H-DOCX 必须同时服务两个目标：Agent 可读可改，以及 DOCX 严格还原。

### 6.0 编辑粒度要求

H-DOCX 的编辑粒度必须符合 WordprocessingML 的真实结构：

- 段落级格式作用于 `w:pPr`，例如对齐、缩进、段前段后、行距、分页控制。
- 字符级格式作用于 `w:rPr`，例如字体、字号、加粗、颜色、上下标。
- 一个段落可以包含多个 run，因此同一段落中可以有多个字号、字体、颜色和加粗状态。
- 对一个词、半句或几个字符做局部格式修改时，应通过拆分 run 实现。
- 如果局部修改会破坏域、批注、书签、修订、内容控件或其他受保护边界，系统必须拒绝或要求更明确的编辑策略。

### 6.1 HTML 表层要求

HTML 表层应满足：

- 结构清晰，便于 Agent 定位段落、标题、表格、脚注、页眉页脚。
- 可编辑节点显式标注 `data-hdocx-id`。
- 节点标注来源 part、XPath 或等价定位信息。
- 节点标注功能类型，例如 paragraph、run、table、field、comment、revision。
- 受保护节点必须显式标注，不允许 Agent 误改。
- 不能只依赖 CSS 表达 DOCX 语义。

### 6.2 Manifest 要求

manifest 至少记录：

- package file list
- content types
- relationships
- part hash
- editable node map
- protected node map
- style map
- numbering map
- media map
- field map
- comments map
- revisions map
- footnotes/endnotes map
- original XML fragments
- unknown extension map

### 6.3 回写要求

回写 DOCX 时必须：

- 校验 H-DOCX 与 manifest 一致。
- 校验未修改节点 hash。
- 只 patch 已声明可编辑节点。
- 未修改 part 优先使用原始 bytes 回写。
- 关系 id 不应无故变化。
- 未知 XML 不应被重新格式化或规范化。
- 输出后运行结构校验和双射校验。

## 7. Agent 使用场景

本项目至少支持以下典型场景：

1. Agent 修改论文正文文字，但不改变排版。
2. Agent 按用户要求修改字号、字体、行距、缩进、段前段后。
3. Agent 调整标题层级和样式。
4. Agent 编辑表格内容和常见表格格式。
5. Agent 修改页眉页脚中的文本。
6. Agent 修改脚注尾注文本。
7. Agent 在不破坏批注和修订的前提下编辑正文。
8. Agent 识别目录、题注、交叉引用等域，并避免错误改写。
9. Agent 替换图片或修改图片说明。
10. Agent 输出可审计的修改报告，说明 DOCX 中哪些结构被改变。

## 8. 验收标准

### 8.1 未修改往返测试

每个测试 DOCX 必须通过：

```text
input.docx -> output.hdocx -> roundtrip.docx
```

验收：

- OPC 规范化后完全等价。
- 所有 part 存在。
- 所有 relationships 存在。
- 所有 media hash 一致。
- 所有 styles、numbering、settings、comments、revisions、footnotes、endnotes 一致。

### 8.2 局部修改测试

对 H-DOCX 做单点修改后回写 DOCX。

验收：

- 目标 XML 子树发生预期变化。
- 非目标 XML 子树 hash 不变。
- 非目标资源 hash 不变。
- Word 能正常打开文档。
- 学术结构、批注、修订、域不被破坏。

### 8.3 失败优先测试

当出现以下情况时必须失败：

- manifest 丢失。
- protected node 被修改。
- 未知结构无法保真。
- relationship id 冲突。
- part hash 不匹配。
- 加密文档无法读取。
- 签名文档被请求编辑但用户未确认签名失效。

## 9. V1 建议范围

V1 应优先完成：

- DOCX package 解包和原样回包。
- main document 的 paragraph/run/text 可逆映射。
- 字符格式和段落格式的常用编辑。
- styles、numbering、settings 的严格保真。
- 表格常用结构和内容编辑。
- 页眉页脚文本编辑。
- 脚注尾注文本编辑。
- 图片保真和基础替换。
- 域、公式、批注、修订、图表、SmartArt、嵌入对象的只读保真。
- 未修改往返测试。
- 局部修改最小变更测试。

V1 不应承诺：

- 重算目录、交叉引用、文献域。
- 语义编辑复杂修订链。
- 编辑 SmartArt、图表数据、嵌入对象。
- 支持宏、ActiveX、加密文档。
- 用普通 HTML 表示全部 DOCX 语义。

## 10. 核心结论

本项目的底线是：

```text
可编辑的，必须精确回写。
不可编辑的，必须严格保真。
不能保真的，必须拒绝处理。
```

只要文档成功进入系统，系统就不能以“HTML 不支持”为理由丢失 Word/DOCX 信息。HTML 只是 Agent 的编辑界面，严格双射由 H-DOCX 的完整结构、manifest、原始 OOXML 和校验机制共同保证。
