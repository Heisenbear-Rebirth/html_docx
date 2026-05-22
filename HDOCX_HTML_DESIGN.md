# H-DOCX HTML 设计草案

## 1. 总体结论

H-DOCX 的 HTML 设计应遵循一个核心原则：

```text
HTML 是 Agent 的编辑投影，manifest 和原始 OOXML 才是严格双射的事实来源。
```

因此，本项目不应把普通 HTML/CSS 当成 DOCX 的完整等价物。正确设计是：

- HTML 负责让 Agent 看懂文档结构并进行局部修改。
- CSS 负责浏览器预览和受控排版修改。
- manifest 负责记录完整 DOCX 结构、原始 XML、关系、资源、hash 和回写规则。
- 原始 DOCX package parts 负责未修改内容的字节级或规范化级保真。

## 2. 是否沿用 CSS 的做法

结论：应该沿用 CSS 的“选择器 + 样式声明”思想，但这里的 CSS 指自定义的 H-CSS，而不是普通网页 CSS。H-CSS 的语法可以接近 CSS，语义必须是 DOCX/OOXML patch 语义。

### 2.1 为什么不能直接使用普通 CSS

普通 CSS 与 DOCX/OOXML 存在根本差异：

- CSS 的 cascade/inheritance 与 Word 样式继承机制不同。
- CSS 的单位、行距、缩进、分页控制与 Word 的 twip、half-point、section properties 并不完全等价。
- Word 有 paragraph properties、run properties、table properties、section properties 等多层属性，CSS 容易把它们混成视觉样式。
- Word 的域、编号、修订、批注、题注、交叉引用等不是 CSS 能表达的。
- CSS 允许大量浏览器属性，若 Agent 随意编辑，会产生无法稳定回写到 OOXML 的状态。

所以，普通 CSS 可以用于预览，但不能成为严格双射的唯一格式描述。

### 2.2 推荐的三层样式模型

H-DOCX 应采用三层样式模型。

#### 第一层：OOXML 原始样式层

来源：DOCX 原始 parts 和 manifest。

内容：

- `word/styles.xml`
- `word/numbering.xml`
- `word/theme/theme1.xml`
- paragraph/run/table/section 的直接属性
- 样式继承关系
- 未知属性和扩展属性

职责：

- 严格保真。
- 作为回写 DOCX 的事实依据。
- 不因 HTML/CSS 预览而丢失或重排。

#### 第二层：生成的预览 CSS

文件建议命名：

```text
styles.generated.css
```

职责：

- 让浏览器中的 H-DOCX 尽量接近 Word 显示效果。
- 帮助 Agent 和用户阅读。
- 可以从 manifest 重新生成。

限制：

- 默认只读。
- 不作为回写 OOXML 的唯一依据。
- Agent 不应直接修改该文件。

示例：

```css
.w-pstyle-Normal {
  font-family: "Times New Roman", "SimSun", serif;
  font-size: 12pt;
  line-height: 1.5;
}

.w-pstyle-Heading1 {
  font-size: 16pt;
  font-weight: 700;
  text-align: center;
}
```

#### 第三层：自定义 H-CSS 编辑样式层

文件建议命名：

```text
agent.edits.hcss
```

这里定义一种自定义样式语言，暂称 H-CSS。它使用 CSS 风格的选择器和声明块，但不是浏览器 CSS，也不是普通 CSS 子集。H-CSS 只允许项目声明过的 DOCX 可回写属性，每条声明都必须能编译成确定的 OOXML patch。

职责：

- 允许 Agent 批量表达排版意图。
- 由 H-DOCX 编译器解析并转换为 OOXML patch。
- 所有声明必须可验证、可追踪、可回滚。

示例：

```css
/* 将所有正文段落设置为小四、1.5 倍行距、首行缩进 2 字符 */
[data-hdocx-style-id="Normal"] {
  hdocx-font-size: 12pt;
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}

/* 将指定标题居中加粗 */
[data-hdocx-id="p-00042"] {
  hdocx-align: center;
  hdocx-bold: true;
}
```

H-CSS 允许的属性必须是白名单，例如：

- `hdocx-font-size`
- `hdocx-font-family`
- `hdocx-ascii-font`
- `hdocx-hansi-font`
- `hdocx-eastAsia-font` / `hdocx-east-asia-font`
- `hdocx-bold`
- `hdocx-italic`
- `hdocx-underline`
- `hdocx-color`
- `hdocx-highlight`
- `hdocx-align`
- `hdocx-left-indent`
- `hdocx-right-indent`
- `hdocx-first-line-indent`
- `hdocx-hanging-indent`
- `hdocx-line-spacing`
- `hdocx-space-before`
- `hdocx-space-after`
- `hdocx-page-break-before`
- `hdocx-keep-with-next`
- `hdocx-style-id`

普通 CSS 属性可以作为兼容别名进入编译器，例如 `font-size`、`font-weight`、`text-align`，但最终必须规范化为 `hdocx-*` 属性。无法映射的 CSS 属性必须报错或警告，不能静默忽略。

## 3. 标签粒度设计

结论：H-DOCX 的主要标签粒度应跟随 OOXML 结构，而不是自然语言结构。

也就是说：

- 一个 `<p>` 对应一个 Word 段落 `w:p`。
- 一个 `<span>` 对应一个 Word run `w:r` 或 run-like inline range。
- 一个 `<table>` 对应一个 Word 表格 `w:tbl`。
- 一个 `<tr>` 对应一个 Word 表格行 `w:tr`。
- 一个 `<td>` 对应一个 Word 单元格 `w:tc`。
- 不应把“一句话”作为基础标签单位。

### 3.1 为什么不按句子切分

DOCX 本身不存储“句子”这个结构。句子只是自然语言语义，不是 WordprocessingML 的稳定编辑边界。

按句子切分会带来问题：

- 一个句子可能跨多个 run，因为中间有加粗、公式、超链接、批注、修订或域。
- 一个 run 可能包含多个句子。
- 批注和修订范围可能从句子中间开始或结束。
- 中文、英文、编号、参考文献、缩写和公式会让句子切分不稳定。
- 回写时很难证明只改变了原始 DOCX 的对应结构。

因此，句子可以作为可选的辅助分析层，但不能成为 H-DOCX 的源结构。

### 3.2 推荐结构层级

H-DOCX 的 HTML 层建议采用以下结构：

```text
html
  head
    metadata
    generated preview css
    editable H-CSS
  body
    article / part
      section
        paragraph
          run / inline object / field / comment anchor / revision range
        table
          row
            cell
              paragraph
```

其中：

- `article` 表示一个 DOCX part，例如 main document、header、footer、footnote、endnote、comment。
- `section` 表示 Word section 边界。
- `p` 表示 Word paragraph。
- `span` 表示 Word run。
- `table/tr/td` 表示 Word 表格结构。
- 复杂对象用受保护 wrapper 表示。

## 4. 推荐 HTML 骨架

示例：

```html
<!doctype html>
<html data-hdocx-version="0.1" data-hdocx-package-id="pkg-001">
  <head>
    <meta charset="utf-8">
    <title>document</title>
    <link rel="stylesheet" href="styles.generated.css" data-hdocx-role="preview">
    <link rel="stylesheet" href="agent.edits.hcss" data-hdocx-role="editable-style-requests">
  </head>
  <body>
    <article
      data-hdocx-type="part"
      data-hdocx-part="/word/document.xml"
      data-hdocx-id="part-main">

      <section
        data-hdocx-type="section"
        data-hdocx-id="sec-00001"
        data-hdocx-source="/word/document.xml">

        <p
          data-hdocx-type="paragraph"
          data-hdocx-id="p-00001"
          data-hdocx-style-id="Title"
          data-hdocx-lock="editable"
          class="w-p w-pstyle-Title">
          <span
            data-hdocx-type="run"
            data-hdocx-id="r-00001"
            data-hdocx-lock="editable"
            class="w-r">
            论文标题
          </span>
        </p>

        <p
          data-hdocx-type="paragraph"
          data-hdocx-id="p-00002"
          data-hdocx-style-id="Normal"
          data-hdocx-lock="editable"
          class="w-p w-pstyle-Normal">
          <span
            data-hdocx-type="run"
            data-hdocx-id="r-00002"
            data-hdocx-lock="editable"
            class="w-r">
            正文内容。
          </span>
        </p>

      </section>
    </article>
  </body>
</html>
```

## 5. 段落、标题与列表

### 5.1 段落

所有 Word 段落都使用 `<p>` 表示，包括普通正文、标题、题注、列表项、目录项。

理由：

- 在 DOCX 中它们本质上都是 `w:p`。
- 标题是带样式和大纲级别的 paragraph，不是独立结构。
- 统一使用 `<p>` 可以减少回写歧义。

标题不建议直接使用 `<h1>`、`<h2>` 作为源结构。可以通过属性表达标题语义：

```html
<p
  data-hdocx-type="paragraph"
  data-hdocx-id="p-01000"
  data-hdocx-style-id="Heading1"
  data-hdocx-outline-level="1"
  role="heading"
  aria-level="1">
  <span data-hdocx-type="run" data-hdocx-id="r-01000">第一章 绪论</span>
</p>
```

这样浏览器和 Agent 能识别标题，但回写时仍然准确对应 Word paragraph。

### 5.2 列表

Word 列表不应直接等同于 HTML 的 `<ul>` / `<ol>`。

原因：

- Word 列表由 `numbering.xml` 中的 abstract numbering 和 concrete numbering 决定。
- 同一列表可能跨多个不连续段落。
- 标题编号和普通列表共用编号机制。

推荐 V1 使用 `<p>` 表示列表项，并保留编号信息：

```html
<p
  data-hdocx-type="paragraph"
  data-hdocx-id="p-02001"
  data-hdocx-style-id="ListParagraph"
  data-hdocx-num-id="7"
  data-hdocx-ilvl="0"
  data-hdocx-list-label="1."
  class="w-p w-list">
  <span data-hdocx-type="run" data-hdocx-id="r-02001">列表内容</span>
</p>
```

浏览器预览层可以把连续列表项渲染成类似 `<ol>` 的视觉效果，但源结构仍然是 paragraph。

## 6. Run 与文本

一个 `<span data-hdocx-type="run">` 对应一个 `w:r`。

示例：

```html
<p data-hdocx-type="paragraph" data-hdocx-id="p-03001">
  <span data-hdocx-type="run" data-hdocx-id="r-03001">这是普通文本，</span>
  <span data-hdocx-type="run" data-hdocx-id="r-03002" data-hdocx-bold="true">这是加粗文本</span>
  <span data-hdocx-type="run" data-hdocx-id="r-03003">。</span>
</p>
```

设计要求：

- Agent 修改 run 中的文本时，只 patch 对应 `w:t`。
- 如果 Agent 在一个 run 内新增局部格式，编译器可以拆分 run。
- 如果 Agent 删除 run，manifest 必须验证该 run 不包含受保护对象、域边界、批注边界或修订边界。
- 空格、制表符、换行符必须显式保真，不能依赖 HTML whitespace 折叠。

对于空格，必要时使用受控节点：

```html
<span data-hdocx-type="tab" data-hdocx-id="tab-00001"></span>
<br data-hdocx-type="line-break" data-hdocx-id="br-00001">
```

### 6.1 局部格式与 run split

一个 Word 段落可以包含任意多个字号、字体、颜色和字符级格式。H-DOCX 的表示方式不是让同一个 `<p>` 承担多个字符格式，而是使用多个 run：

```html
<p data-hdocx-type="paragraph" data-hdocx-id="p-10001">
  <span data-hdocx-type="run" data-hdocx-id="r-10001" data-hdocx-font-size="16pt">研究目的</span>
  <span data-hdocx-type="run" data-hdocx-id="r-10002" data-hdocx-font-size="12pt">是分析……</span>
</p>
```

当 Agent 只修改一个 run 中的部分文字格式时，编译器应执行 run split：

```html
<span data-hdocx-type="run" data-hdocx-id="r-10003a">本文提出一种</span>
<span data-hdocx-type="run" data-hdocx-id="r-10003b" data-hdocx-bold="true">新方法</span>
<span data-hdocx-type="run" data-hdocx-id="r-10003c">。</span>
```

run split 必须遵守安全边界：

- 不能切开 field、comment、bookmark、revision、permission range 等受保护边界。
- 不能切在 Unicode grapheme cluster 中间。
- 新 run 必须继承原 run 的未修改属性。
- 只有目标 run fragment 可以变化。
- 如果原 run 含有 drawing、object、field char、instrText、tab、break 等复杂子节点，必须先验证可拆分，否则拒绝。

## 7. 表格设计

Word 表格使用标准 HTML 表格标签，但必须附加 DOCX 元信息。

示例：

```html
<table
  data-hdocx-type="table"
  data-hdocx-id="tbl-00001"
  data-hdocx-style-id="TableGrid">
  <tr data-hdocx-type="table-row" data-hdocx-id="tr-00001">
    <td
      data-hdocx-type="table-cell"
      data-hdocx-id="tc-00001"
      data-hdocx-grid-span="2"
      colspan="2">
      <p data-hdocx-type="paragraph" data-hdocx-id="p-04001">
        <span data-hdocx-type="run" data-hdocx-id="r-04001">单元格内容</span>
      </p>
    </td>
  </tr>
</table>
```

要求：

- `colspan` / `rowspan` 只是浏览器预览辅助。
- 真正的合并单元格信息必须在 `data-hdocx-*` 和 manifest 中保留。
- 垂直合并、网格定义、单元格边距、表格缩进、重复标题行必须保真。

## 8. 域、公式、批注、修订的表示

这些功能在学术写作中必须严格保真，但 V1 默认不允许随意编辑。

### 8.0 重叠范围规则

HTML 是树结构，但 WordprocessingML 中存在大量“范围型结构”，它们可能与 run、段落、批注、书签、域、修订、权限范围发生交叉或重叠。典型例子包括：

- comment range
- bookmark range
- move range
- permission range
- field begin/separate/end
- proofing error range
- custom XML range

因此，H-DOCX 不能假设所有 Word 范围都能用一个 HTML wrapper 正确表示。凡是可能发生交叉的范围，必须使用 start/end 边界锚点加 manifest range map 表示，而不是强行嵌套。

示例：

```html
<span data-hdocx-type="range-start" data-hdocx-range-kind="bookmark" data-hdocx-range-id="bm-1"></span>
<span data-hdocx-type="run" data-hdocx-id="r-10001">文本 A</span>
<span data-hdocx-type="range-start" data-hdocx-range-kind="comment" data-hdocx-range-id="cmt-3"></span>
<span data-hdocx-type="run" data-hdocx-id="r-10002">文本 B</span>
<span data-hdocx-type="range-end" data-hdocx-range-kind="bookmark" data-hdocx-range-id="bm-1"></span>
<span data-hdocx-type="run" data-hdocx-id="r-10003">文本 C</span>
<span data-hdocx-type="range-end" data-hdocx-range-kind="comment" data-hdocx-range-id="cmt-3"></span>
```

manifest 中必须记录这些范围的原始 XML 节点、位置、id、hash 和交叉关系。若 Agent 修改导致范围边界无法回写，编译器必须拒绝。

### 8.1 域

字段代码和结果需要同时保留。

示例：

```html
<span
  data-hdocx-type="field"
  data-hdocx-id="fld-00001"
  data-hdocx-field-kind="TOC"
  data-hdocx-lock="protected">
  目录
</span>
```

manifest 中保存完整 field begin/separate/end 结构、field code 和 field result。

### 8.2 公式

V1 推荐受保护占位：

```html
<span
  data-hdocx-type="equation"
  data-hdocx-id="eq-00001"
  data-hdocx-lock="protected">
  [equation: eq-00001]
</span>
```

manifest 中保存完整 OMML XML。后续如果实现公式编辑，可以引入 MathML/LaTeX 的受控映射，但不能替代原始 OMML 保真层。

### 8.3 批注

批注范围需要以锚点表示，而不是只显示批注文本。

```html
<span data-hdocx-type="comment-start" data-hdocx-comment-id="3"></span>
<span data-hdocx-type="run" data-hdocx-id="r-05001">被批注的文本</span>
<span data-hdocx-type="comment-end" data-hdocx-comment-id="3"></span>
```

批注正文作为独立 part/article 表示：

```html
<article
  data-hdocx-type="part"
  data-hdocx-part="/word/comments.xml"
  data-hdocx-id="part-comments">
  <section data-hdocx-type="comment" data-hdocx-comment-id="3">
    <p data-hdocx-type="paragraph" data-hdocx-id="comment-p-00003">
      <span data-hdocx-type="run" data-hdocx-id="comment-r-00003">批注意见</span>
    </p>
  </section>
</article>
```

### 8.4 修订

修订结构默认受保护。

```html
<span
  data-hdocx-type="revision"
  data-hdocx-revision-kind="insertion"
  data-hdocx-revision-id="12"
  data-hdocx-author="Author"
  data-hdocx-lock="protected">
  <span data-hdocx-type="run" data-hdocx-id="r-06001">插入内容</span>
</span>
```

V1 中 Agent 不应直接改修订 wrapper。若需要编辑含修订文本，编译器必须明确判断是否安全。

## 9. 图片与对象

图片使用 `<img>`，但关系和 DrawingML/VML 结构必须由 manifest 保真。

```html
<figure
  data-hdocx-type="drawing"
  data-hdocx-id="draw-00001"
  data-hdocx-lock="editable-metadata">
  <img
    src="media/image1.png"
    data-hdocx-type="image"
    data-hdocx-id="img-00001"
    data-hdocx-rid="rId8"
    alt="图片说明">
</figure>
```

要求：

- `src` 是预览路径，不是唯一事实来源。
- `rId`、原始 DrawingML、尺寸、裁剪、环绕、锚定信息必须保真。
- V1 可支持替换图片、修改 alt text、修改简单尺寸。
- 浮动位置、复杂环绕、SmartArt、图表、OLE 默认受保护。

## 10. Style 修改方式

H-DOCX 需要同时支持两类格式修改。

### 10.1 局部直接格式

适合用户说：

```text
把这一段居中。
把这个词加粗。
```

HTML 表示：

```html
<p
  data-hdocx-type="paragraph"
  data-hdocx-id="p-07001"
  data-hdocx-align="center">
  <span data-hdocx-type="run" data-hdocx-id="r-07001">标题</span>
</p>
```

回写目标：修改该段落的 `w:pPr`。

### 10.2 样式级批量修改

适合用户说：

```text
把所有正文改成小四，1.5 倍行距，首行缩进 2 字符。
```

H-CSS 表示：

```css
[data-hdocx-style-id="Normal"] {
  hdocx-font-size: 12pt;
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}
```

编译器需要决定这是：

- 修改 `styles.xml` 中的 `Normal` 样式；或
- 给所有匹配段落添加直接格式。

这个选择不能靠 Agent 猜，必须由 H-DOCX 命令或配置明确指定。

推荐增加作用域声明：

```css
@hdocx-edit mode(style-definition);

[data-hdocx-style-id="Normal"] {
  hdocx-font-size: 12pt;
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}
```

或者：

```css
@hdocx-edit mode(direct-formatting);
```

## 11. ID 与定位

每个可编辑或受保护节点必须有稳定 ID。

推荐 ID 类型：

- `part-main`
- `sec-00001`
- `p-00001`
- `r-00001`
- `tbl-00001`
- `tr-00001`
- `tc-00001`
- `img-00001`
- `fld-00001`
- `cmt-00001`
- `rev-00001`

HTML 中的 ID 用于 Agent 定位；manifest 中必须记录：

- part path
- XML source locator
- original hash
- node kind
- editable/protected status
- parent/children relation
- OOXML namespace context

不要只依赖 XPath，因为 XML 被 patch 后 XPath 可能变化。应使用稳定 ID + 结构路径 + hash 的组合定位。

### 11.1 class、命名集合与格式复用

H-DOCX 有必要提供类似 `class/id` 的复用能力，但职责必须分开：

- `data-hdocx-id` 是权威节点身份，用于精确定位。
- 生成的 HTML `class` 用于预览、阅读和批量选择，例如 `hstyle-Normal`、`hdocx-p`、`hlock-editable`。
- Agent 不应通过随意修改普通 `class` 来表达 DOCX 修改。
- Agent 需要复用选择器时，应使用 H-CSS 的 `@hdocx-set`。
- Agent 需要复用格式声明时，应使用 H-CSS 的 `@hdocx-format`。
- Agent 需要复用数值时，应使用 H-CSS 的 `@hdocx-token`。

示例：

```css
@hdocx-set body-paragraphs {
  select: [data-hdocx-type="paragraph"][data-hdocx-style-id="Normal"];
  exclude: [data-hdocx-lock="protected"];
}

@hdocx-format thesis-body-text {
  hdocx-font-size: 12pt;
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}

@hdocx-edit mode(style-definition);

body-paragraphs {
  @hdocx-include thesis-body-text;
}
```

详细规则见 `SELECTOR_AND_REUSE_DESIGN.md`。

## 12. 修改与校验流程

推荐流程：

```text
1. DOCX -> H-DOCX
2. Agent 编辑 document.html 和/或 agent.edits.hcss
3. 编译器读取 HTML、H-CSS、manifest
4. 校验受保护节点未被修改
5. 校验未修改节点 hash
6. 将允许的修改转换为 OOXML patch
7. 未修改 parts 使用原始 bytes 回写
8. 输出 DOCX
9. 再执行 docx -> h-docx 或 package diff 做双射校验
```

严格双射还要求 H-DOCX bundle 保存原始 DOCX package bytes 或等价的原始 ZIP entry 信息。未修改往返时，最安全策略是直接返回原始 DOCX bytes。发生局部修改时，未修改 ZIP entries 应尽量复用原始 entry bytes、metadata、顺序和 relationships；被修改的 XML part 必须使用最小 patch 策略，避免整份 XML parse 后重新序列化造成 namespace、属性顺序、空白、`mc:AlternateContent` 或扩展属性变化。

## 13. 设计取舍

### 13.1 标签以段落和 run 为主

最终决策：

- 段落用 `<p>`。
- run 用 `<span>`。
- 表格用 `<table>/<tr>/<td>`。
- 标题仍用 `<p>`，通过 style/outline/ARIA 表达标题语义。
- 列表仍用 `<p>`，通过 numbering 属性表达列表语义。
- 句子不作为源结构。

### 13.2 CSS 可以用，但必须受控

最终决策：

- `styles.generated.css` 用于预览，只读，可再生成。
- `agent.edits.hcss` 用于 Agent 批量排版修改。
- H-CSS 只允许白名单属性和可验证选择器。
- 普通 CSS 属性如果被支持，只能作为 H-CSS 属性别名，不能覆盖 DOCX 事实层。

### 13.3 受保护结构必须显式存在

最终决策：

- 域、公式、批注、修订、图表、SmartArt、OLE 等都必须在 HTML 中有可见或可定位占位。
- Agent 不应误删它们。
- manifest 负责保存完整原始结构。

## 14. V1 推荐实现范围

V1 的 HTML 设计应先支持：

- main document part。
- paragraph/run/text 映射。
- 常用 run properties。
- 常用 paragraph properties。
- styles.generated.css 预览生成。
- agent.edits.hcss 的受限解析。
- table/tr/td 基础映射。
- header/footer/footnote/endnote 作为独立 article。
- field/equation/comment/revision 的受保护占位。
- manifest 中的 ID、part、hash、lock 状态。

V1 不应先做：

- 任意 CSS 到 DOCX 的自由转换。
- 按句子切分作为源结构。
- 自动重算目录和交叉引用。
- 复杂修订链编辑。
- SmartArt、图表、OLE 的语义编辑。

## 15. 核心规则

```text
结构跟随 OOXML。
显示可以像 HTML。
批量排版可以像 CSS。
回写必须像补丁。
未知内容必须保真。
不能证明安全就拒绝。
```
