# H-DOCX 边缘性测试矩阵

## 1. 测试目标

本文档用于压力测试当前 H-DOCX 设计是否能够覆盖 DOCX 的极端情况，并维持严格双射。

这里的“测试”不是运行代码，而是设计级边缘性测试：

```text
给定一个困难 DOCX 特性
-> H-DOCX 是否能表示
-> Agent 是否能安全编辑
-> 回写 DOCX 时如何保证未修改内容不变
-> 如果不能保证，是否会拒绝而不是损坏文档
```

## 2. 总体结论

当前设计在补充以下硬约束后，可以支撑严格双射：

1. H-DOCX 必须是一个 bundle，而不是单独一个 HTML 文件。
2. H-DOCX bundle 必须保存原始 DOCX package bytes 或等价的原始 ZIP entry 信息。
3. HTML 只是 Agent 可编辑投影，不是完整事实来源。
4. H-CSS 是自定义 DOCX 样式语言，不是普通网页 CSS。
5. 未修改内容必须原样回填，不能整份 XML 重新序列化。
6. 可能重叠的 Word 范围必须用 start/end 锚点和 manifest range map 表示，不能强行用 HTML wrapper。
7. 任何无法证明安全的编辑都必须拒绝。

因此，严格双射成立的前提是：

```text
DOCX <-> H-DOCX bundle
```

而不是：

```text
DOCX <-> plain HTML
```

单独的普通 HTML 无法覆盖 DOCX 全部结构。

## 3. 双射域定义

### 3.1 支持域

V1 设计支持以下输入域：

- 未加密 `.docx`。
- 符合 OPC/OOXML 基本结构的 package。
- 包含未知 namespace、扩展 part、`mc:AlternateContent`、VML、DrawingML、custom XML、批注、修订、域、公式等复杂结构的文档。

只要这些内容可以作为原始 package/part/fragment 保留，就属于可进入系统的文档。

### 3.2 拒绝域

V1 应拒绝：

- 加密 DOCX。
- 损坏到无法读取 package 目录的 DOCX。
- 需要执行宏或 ActiveX 才能理解内容的文档。
- 用户要求编辑会导致数字签名失效但未明确确认的文档。
- manifest 缺失或被篡改的 H-DOCX。
- protected node 被修改且无法安全恢复的 H-DOCX。

拒绝是严格双射设计的一部分。不能安全双射时，正确行为是失败。

## 4. 关键不变量

### 4.1 Package 不变量

未修改往返：

```text
decode(encode(input.docx)) == input.docx
```

最强策略：

- 如果 H-DOCX 没有修改，直接返回原始 DOCX bytes。

局部修改：

- 未修改 ZIP entries 必须保留原始 bytes 或规范化后等价。
- 未修改 relationships 不变。
- 未修改 content types 不变。
- 未修改 media hash 不变。
- 未修改 XML fragments 不变。

### 4.2 XML 不变量

对未修改 XML 内容：

- namespace prefix 不应变化。
- attribute order 不应无故变化。
- whitespace 不应无故变化。
- comments 和 processing instructions 不应丢失。
- `mc:Ignorable`、`mc:AlternateContent` 不应被折叠。
- unknown namespace elements/attributes 不应被删除。
- `w:rsid*` 等 Word 编辑痕迹不应被清理。

因此，回写应使用 fragment patch，而不是整份 DOM parse/serialize。

### 4.3 HTML 不变量

HTML 中：

- 每个可编辑节点必须有稳定 `data-hdocx-id`。
- 每个受保护节点必须有 `data-hdocx-lock="protected"` 或等价标记。
- 可能交叉的范围必须用 start/end 锚点。
- Agent 不能通过删除 wrapper 的方式隐式删除原始 OOXML 结构。

### 4.4 H-CSS 不变量

H-CSS 中：

- 选择器必须只匹配 H-DOCX 节点。
- 属性必须来自白名单。
- 单位必须能精确转换为 OOXML 单位。
- 冲突规则必须有确定优先级。
- 修改 style definition 还是 direct formatting 必须显式声明。
- 无法映射的声明必须报错，不能静默忽略。

## 5. 边缘性测试矩阵

### 5.1 ZIP/OPC 层

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| ZIP entry 顺序特殊 | 重新打包后字节差异 | 保存原始 entry 顺序和 metadata | 未修改时直接返回原始 bytes；修改时尽量保持顺序 |
| ZIP entry 时间戳/压缩参数特殊 | 字节级 diff 变大 | 保存原始 ZIP metadata | 未修改 entry 原样保留 |
| `[Content_Types].xml` 有未知 Override | 未知 part 丢失 | manifest 记录所有 content type | 原样保留 |
| package-level `.rels` 有外部链接 | 外链被内联或删除 | relationships map 保留 TargetMode | 原样保留 |
| part name 含空格、中文、特殊字符 | 路径编码错误 | manifest 使用 OPC part name 原值 | 不规范化路径 |
| custom XML part | 数据绑定断裂 | C 类透明保真 | 不展示或只读，原样回写 |
| thumbnail、docProps | 元数据变化 | C 类透明保真 | 未修改时 hash 不变 |
| 数字签名 part | 编辑后签名失效 | B/D 类 | 未编辑可保真；编辑需拒绝或确认 |
| 加密 DOCX | 无法读取内部 parts | D 类 | 拒绝 |

结论：只要 H-DOCX 保存原始 package/entry 信息，OPC 层可严格保真。

### 5.2 XML 序列化层

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| namespace prefix 非常规 | parse/serialize 改 prefix | 原始 XML fragment hash | 未修改 fragment 不重写 |
| 属性顺序特殊 | 重新序列化导致 diff | 原始 fragment 保留 | 未修改不动 |
| `mc:AlternateContent` | fallback 被错误选择 | protected fragment | 原样保留 |
| `mc:Ignorable` 扩展 | 扩展属性丢失 | unknown namespace map | 原样保留 |
| XML comments | 注释丢失 | fragment 保留 | 原样保留 |
| processing instruction | PI 丢失 | fragment 保留 | 原样保留 |
| `w:rsid*` | Word 修订痕迹变化 | 原始属性保留 | 不清理 |

结论：必须采用 patch 模型，不能把 XML 当普通 DOM 全量重写。

### 5.3 段落与 run

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| 一个段落有多个 run | 合并 run 后格式丢失 | `<p>` + 多个 run span | 保持 run 边界 |
| 一个 run 有多个 `w:t` | 文本节点合并造成空格丢失 | manifest 保存 run 内子结构 | 简单编辑可 patch；复杂编辑可拆 run |
| `xml:space="preserve"` | 首尾空格丢失 | 显式 whitespace 记录 | 回写保留 |
| tab、line break、page break | 被当普通空白 | 专用节点 | 回写为原 OOXML |
| soft hyphen/no-break hyphen | 字符被普通化 | 专用 inline token 或原文保留 | 不丢失 |
| 符号字体字符 | Unicode 与字体语义混淆 | run properties 保留 | 不改字体属性 |
| RTL/bidi 文本 | 顺序错乱 | run/paragraph bidi 属性 | 保真，编辑需校验 |
| 拼音指南 ruby | HTML 无等价 | protected inline object | 原样保真 |
| 东亚着重号 | CSS 近似不足 | OOXML 属性保留 | H-CSS 可映射时才编辑 |

结论：paragraph/run 粒度正确，但 run 内部仍需 manifest 保存子结构。

### 5.4 重叠范围

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| 批注范围与书签交叉 | HTML wrapper 无法表示 | start/end 锚点 + range map | 可表示；破坏边界则拒绝 |
| 域 begin/separate/end 跨多个 run | 结果和代码错位 | field boundary map | 默认 protected |
| move range 与 paragraph 交叉 | 树结构不匹配 | boundary anchors | 严格保真 |
| permission range | 权限边界丢失 | range anchors | 严格保真 |
| proofErr 范围 | 拼写错误标记丢失 | range anchors | 严格保真 |

结论：这是设计中最容易出错的地方。必须禁止“所有复杂结构都用 wrapper”的假设。

### 5.5 样式与 H-CSS

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| 修改所有正文字号 | 不清楚改样式还是直接格式 | `@hdocx-edit mode(...)` | 必须显式选择 |
| 样式继承链复杂 | 视觉值与定义值混淆 | styles.xml 保真 + computed preview | 回写只改目标层 |
| linked style | 段落/字符样式混淆 | style map | 保真，编辑需指定类型 |
| latent style | 不显示但影响 Word | style map | 原样保留 |
| 主题字体 | CSS 字体名不等于主题绑定 | theme map | 不展开替换，除非明确 |
| 冲突 H-CSS 规则 | 输出不确定 | 固定优先级和冲突诊断 | 报错或按规则确定 |
| `2char` 缩进 | 依赖中文字符宽度 | H-CSS 单位转换规则 | 编译成 OOXML twips/chars |
| 行距 `1.5` | Word line rule 语义不同 | `hdocx-line-spacing` 明确定义 | 精确映射为 OOXML |

结论：H-CSS 可行，但必须是自定义 DSL，并带类型系统、单位系统和冲突规则。

### 5.6 编号与列表

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| 多级标题编号 | `<ol>` 无法表达 | paragraph + numId/ilvl | 保真 |
| 列表跨非列表段落 | HTML list 结构断裂 | paragraph 属性表达 | 保真 |
| lvlOverride | 编号起点丢失 | numbering map | 保真 |
| restart numbering | 重启信息丢失 | numPr/numbering map | 保真 |
| 图片项目符号 | HTML bullet 不够 | protected numbering resource | 保真 |
| 符号字体项目符号 | 字符/字体混淆 | numbering level font | 保真 |
| legal numbering | 编号格式特殊 | numbering map | 保真 |

结论：列表源结构必须是 paragraph，不能依赖 `<ol>/<ul>`。

### 5.7 表格

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| 横向合并 `gridSpan` | colspan 不等价 | `data-hdocx-grid-span` + manifest | 保真 |
| 纵向合并 `vMerge` | rowspan 不等价 | vMerge map | 保真 |
| omitted cells | HTML 表格无法自然表示 | row gridBefore/gridAfter map | 保真 |
| tblGrid 与实际单元格不一致 | 自动修复导致变化 | 原始 tblGrid 保留 | 不修复 |
| 嵌套表格 | 结构复杂 | td 内嵌 table | 可表示 |
| 重复标题行 | HTML 无语义 | row property | 保真 |
| 表格跨页控制 | HTML 无语义 | trPr/tblPr 保留 | 保真 |
| 单元格文本方向 | CSS 近似不足 | tcPr 保留 | 保真 |

结论：HTML 表格标签可用于投影，但真实表格语义在 manifest/OOXML。

### 5.8 分节与页面布局

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| `sectPr` 在段落属性中 | section wrapper 锚点错 | section map 绑定 paragraph/body | 保真 |
| 最后一节 `sectPr` 在 body 末尾 | 边界丢失 | body-level section map | 保真 |
| continuous section | 错当分页 | section type 保留 | 保真 |
| 每节不同页眉页脚 | part 关系错配 | section -> header/footer refs | 保真 |
| 奇偶页/首页不同 | header/footer 类型丢失 | sectPr map | 保真 |
| 分栏 | HTML 难显示 | section properties | 保真 |
| 页码重启 | 页码域显示错误 | section pgNumType | 保真 |

结论：section 是核心结构边界，不能只靠视觉切块。

### 5.9 页眉页脚、脚注尾注

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| header/footer 多 part | 只保留一个 | 每个 part 一个 article | 保真 |
| 页眉里有表格/图片 | 简化成文本 | article 内完整结构 | 保真 |
| 脚注 separator | 被当普通脚注 | special note map | 保真 |
| continuation separator | 丢失 | special note map | 保真 |
| 脚注中有批注/修订 | 嵌套复杂 | 同正文规则 | 保真 |
| 尾注编号格式 | 格式丢失 | notes properties | 保真 |

结论：这些必须作为独立 part/article，而不是主文档附属文本。

### 5.10 域、目录、交叉引用

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| complex field | begin/separate/end 丢失 | field map + anchors | 保真 |
| simple field | 结构被展开 | field node | 保真 |
| 嵌套 field | 边界错乱 | stack/range map | 保真 |
| TOC 跨多个段落 | 单节点无法表示 | protected field range | 保真 |
| REF/PAGEREF | 书签引用断裂 | field code 保留 | 保真 |
| Citation/Bibliography | 引文数据丢失 | field + custom XML 保留 | 保真 |
| 用户改标题文本 | TOC 结果变旧 | 标记 needs-update-fields | 不自动猜测 |

结论：V1 不重算域是正确选择。严格保真优先。

### 5.11 批注与现代评论

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| comment range start/end | 范围错位 | boundary anchors | 保真 |
| comments.xml 文本 | 批注正文丢失 | 独立 article | 保真 |
| commentsExtended.xml | 现代属性丢失 | related part 保留 | 保真 |
| threaded comments | 线程关系丢失 | people/thread map | 保真 |
| 批注引用点被删除 | dangling comment | 校验失败或连同范围安全删除 | 默认拒绝 |

结论：批注可展示，但 V1 默认不编辑范围结构。

### 5.12 修订

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| `w:ins` | 插入记录丢失 | protected revision node | 保真 |
| `w:del` + `w:delText` | 删除文本变普通文本 | protected revision node | 保真 |
| moveFrom/moveTo | 移动关系断裂 | revision range map | 保真 |
| paragraph property revision | 格式修订丢失 | pPrChange map | 保真 |
| table revision | 行列修订丢失 | table revision map | 保真 |
| 编辑修订内部文本 | 修订语义改变 | 默认拒绝 | 后续专门设计 |

结论：修订是 V1 高风险区。默认只读保真是必要的。

### 5.13 图片、绘图、对象

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| inline drawing | 尺寸/关系丢失 | img preview + drawing map | 保真 |
| floating anchor | 位置/环绕丢失 | protected anchor map | 保真 |
| crop | 图片显示变化 | blipFill/srcRect 保留 | 保真 |
| EMF/WMF/SVG | 浏览器预览不支持 | preview fallback + 原资源 | 原资源保真 |
| VML fallback | fallback 丢失 | AlternateContent/VML fragment | 保真 |
| text box 含正文 | 被忽略 | protected 或独立子 article | 保真 |
| chart | 数据关系丢失 | chart part map | 保真 |
| SmartArt | data model 丢失 | protected object | 保真 |
| OLE embedded object | 二进制损坏 | opaque part/resource | 保真 |

结论：图形对象可以预览，但事实来源必须是原始 DrawingML/VML/object parts。

### 5.14 公式

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| inline OMML | 转 MathML 后丢失 | protected OMML fragment | 保真 |
| display equation | 段落布局丢失 | equation object + paragraph map | 保真 |
| 公式编号 | 与题注/域关系断裂 | field/bookmark map | 保真 |
| 公式内特殊结构 | 转换不完整 | 原始 OMML | 保真 |

结论：公式初期必须以 OMML 保真为主，编辑功能后置。

### 5.15 内容控件、表单、自定义 XML

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| SDT rich text | 控件属性丢失 | sdt wrapper/protected map | 保真 |
| checkbox/dropdown/date | 控件值和属性混淆 | control map | 保真 |
| XML mapping | 数据绑定断裂 | custom XML + mapping map | 保真 |
| legacy form field | 字段损坏 | protected field/control | 保真 |

结论：可以显示内容，但控件结构默认保护。

### 5.16 保护、安全与格式异常

| 测试项 | 风险 | H-DOCX 表示 | 期望行为 |
|---|---|---|---|
| restrict editing | 用户越权修改 | protection map | 拒绝违反保护的编辑 |
| read-only recommended | 元数据丢失 | settings 保留 | 保真 |
| password hash | hash 丢失 | settings 保留 | 保真 |
| macro/ActiveX | 安全风险 | D 类 | 拒绝或透明保真但不执行 |
| 非法但 Word 可打开的 XML | 修复后 diff | raw fragment | 不主动修复 |

结论：安全相关内容必须保守处理。

## 6. 发现的设计修正点

边缘性测试暴露出当前设计必须明确以下修正。

### 6.1 H-CSS 必须被定义为自定义 DSL

不能描述为“普通 CSS 子集”。它只是借用 CSS 外形。

必须新增：

- 属性白名单。
- 选择器白名单。
- 单位系统。
- 冲突优先级。
- style-definition/direct-formatting 模式。
- 无法映射时报错。

### 6.2 HTML wrapper 不能表示所有范围

批注、书签、域、权限范围、proofing range、move range 等可能交叉。必须统一支持：

- `range-start`
- `range-end`
- manifest range map
- 边界 hash
- 破坏边界时拒绝

### 6.3 未修改 XML 必须 fragment patch

为了严格双射，不能把整份 `document.xml` 读进 DOM 后重新输出。应尽量：

- 保存原始 XML bytes。
- 建立节点 offset 或等价定位。
- 只替换被修改的 fragment。
- 未修改 fragment 原样拼回。

### 6.4 run 内部结构不能只看文本

一个 run 内可能有：

- `w:t`
- `w:tab`
- `w:br`
- `w:fldChar`
- `w:instrText`
- `w:drawing`
- `w:object`
- `w:sym`

所以 run 的 HTML 文本只是投影，manifest 必须保存 run child sequence。

### 6.5 双射应以 bundle 为单位

如果用户要求“单个 HTML 文件本身必须包含全部 DOCX 信息”，理论上也可以通过嵌入 base64 package 和 manifest 实现，但这会非常臃肿，不利于 Agent 编辑。

推荐实际格式：

```text
document.hdocx/
  document.html
  manifest.json
  styles.generated.css
  agent.edits.hcss
  package/
    original.docx
    parts/
    media/
```

或者打包为：

```text
document.hdocx.zip
```

## 7. 通过/失败判定

### 7.1 设计通过的情况

以下情况下设计可维持严格双射：

- 文档未修改，直接输出原始 DOCX bytes。
- 只修改普通段落文本，且不跨 protected/range 边界。
- 只修改普通 run 属性，且目标 run 不含字段、修订、批注边界或复杂 inline 对象。
- 只通过 H-CSS 修改白名单内格式，且选择器、单位、模式明确。
- 未知对象、复杂对象、扩展 XML 均未被编辑，作为 opaque fragment 原样回写。

### 7.2 必须拒绝的情况

以下情况必须失败：

- Agent 删除或改写 protected 节点。
- Agent 改动导致 range-start/range-end 不成对。
- Agent 改动导致批注、书签、域、修订边界无法定位。
- H-CSS 使用未声明属性或无法转换单位。
- H-CSS 规则冲突且无确定优先级。
- manifest hash 与 HTML 不一致。
- 原始 package 缺失。
- 编辑会破坏数字签名但用户未确认。

## 8. 最终结论

当前设计经边缘性测试后，结论是：

```text
可以严格双射，但必须把 H-DOCX 定义为带原始 DOCX 保真层的 bundle。
HTML 不能独自承担双射。
H-CSS 可以作为自定义排版编辑语言，但必须编译为受控 OOXML patch。
所有复杂、未知、重叠、不可安全编辑的结构必须原样保真或拒绝。
```

这一设计能够囊括 DOCX 的复杂情况，因为它不要求所有 DOCX 语义都被 HTML 原生表达，而是把 HTML 当作可编辑投影，把完整 OOXML/OPC 信息保存在 manifest 和原始 package 层中。
