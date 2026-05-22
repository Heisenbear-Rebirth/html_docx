# H-DOCX 编辑语义边界测试

## 1. 目标

本文档测试 H-DOCX 在实际编辑场景中的表达能力，重点回答类似问题：

- 一个段落中能不能有两个字号？
- 一个词能不能加粗？
- 半句话能不能改颜色？
- 一个 run 内部局部改格式怎么办？
- 段落样式和局部格式冲突时谁优先？
- H-CSS 应该作用在段落、run、范围还是样式定义上？

结论先行：

```text
一个段落当然可以有多个字号、字体、颜色、加粗状态。
实现方式不是给同一个段落设置多个字号，而是把段落拆成多个 run。
段落级属性写入 w:pPr。
字符级属性写入 w:rPr。
局部文本格式修改通过 run split 实现。
```

## 2. 基础编辑单位

H-DOCX 必须区分三类编辑单位。

| 单位 | HTML 表示 | OOXML 对应 | 适合修改 |
|---|---|---|---|
| 段落 | `<p>` | `w:p` | 对齐、缩进、段前段后、行距、分页、段落样式 |
| Run | `<span data-hdocx-type="run">` | `w:r` | 字体、字号、加粗、颜色、上下标、字符样式 |
| 文本范围 | `range-start/range-end` 或编译器临时选择区 | run split 后映射到多个 `w:r` | 半句、一个词、几个字符的局部格式 |

不能把所有格式都挂在段落上，也不能把所有段落排版都挂在 run 上。

## 3. 一个段落内两个字号

### 3.1 场景

用户要求：

```text
同一个段落中，“研究目的”四个字用三号，其余文字用小四。
```

### 3.2 H-DOCX 表示

```html
<p
  data-hdocx-type="paragraph"
  data-hdocx-id="p-0001"
  data-hdocx-style-id="Normal">
  <span
    data-hdocx-type="run"
    data-hdocx-id="r-0001"
    data-hdocx-font-size="16pt">研究目的</span>
  <span
    data-hdocx-type="run"
    data-hdocx-id="r-0002"
    data-hdocx-font-size="12pt">是分析……</span>
</p>
```

### 3.3 回写规则

- `p-0001` 回写为一个 `w:p`。
- `r-0001` 回写为一个 `w:r`，其 `w:rPr/w:sz` 为三号对应值。
- `r-0002` 回写为另一个 `w:r`，其 `w:rPr/w:sz` 为小四对应值。
- 段落属性不应被改成两个字号，因为字号是 run-level 属性。

### 3.4 判定

支持，属于 A 类可编辑。

## 4. 一个词加粗

### 4.1 原始结构

```html
<p data-hdocx-type="paragraph" data-hdocx-id="p-0002">
  <span data-hdocx-type="run" data-hdocx-id="r-0003">本文提出一种新方法。</span>
</p>
```

用户要求：

```text
把“新方法”加粗。
```

### 4.2 编译器处理

原始 run 需要拆分为三个 run：

```html
<p data-hdocx-type="paragraph" data-hdocx-id="p-0002">
  <span data-hdocx-type="run" data-hdocx-id="r-0003a">本文提出一种</span>
  <span data-hdocx-type="run" data-hdocx-id="r-0003b" data-hdocx-bold="true">新方法</span>
  <span data-hdocx-type="run" data-hdocx-id="r-0003c">。</span>
</p>
```

### 4.3 严格双射要求

- 只有 `r-0003` 对应 XML fragment 被替换。
- 新增 run 必须继承原 run 的未修改 `w:rPr`。
- 加粗 run 只增加或修改 `w:b`。
- 如果原 run 内包含 field、comment boundary、revision boundary、drawing、tab、break 等复杂子节点，必须先判断是否可安全拆分。

### 4.4 判定

支持，但要求 run split 安全检查。

## 5. 同一段落中混合字体

### 5.1 场景

```text
中文用宋体，英文用 Times New Roman。
```

### 5.2 推荐表示

如果 Word 原本使用同一 run 的东亚字体和西文字体属性：

```html
<span
  data-hdocx-type="run"
  data-hdocx-id="r-0010"
  data-hdocx-east-asia-font="SimSun"
  data-hdocx-font-family="Times New Roman">中文 English</span>
```

如果用户要求某几个字符使用完全不同字体，则拆分 run：

```html
<span data-hdocx-type="run" data-hdocx-id="r-0011" data-hdocx-east-asia-font="SimSun">中文</span>
<span data-hdocx-type="run" data-hdocx-id="r-0012" data-hdocx-font-family="Times New Roman">English</span>
```

### 5.3 判定

支持。中英文字体应使用独立 H-CSS 属性，不能只用一个 `font-family` 混过去。

## 6. 段落居中但局部文字加粗

### 6.1 场景

```text
整段居中，其中“关键结论”加粗。
```

### 6.2 表示

```html
<p
  data-hdocx-type="paragraph"
  data-hdocx-id="p-0010"
  data-hdocx-align="center">
  <span data-hdocx-type="run" data-hdocx-id="r-0020">本文的</span>
  <span data-hdocx-type="run" data-hdocx-id="r-0021" data-hdocx-bold="true">关键结论</span>
  <span data-hdocx-type="run" data-hdocx-id="r-0022">如下。</span>
</p>
```

### 6.3 判定

支持。段落级和 run 级属性互不冲突：

- `data-hdocx-align` -> `w:pPr/w:jc`
- `data-hdocx-bold` -> `w:rPr/w:b`

## 7. 一个 run 内局部改字号

### 7.1 场景

原始：

```html
<span data-hdocx-type="run" data-hdocx-id="r-0030">abcde</span>
```

用户要求：

```text
把 bcd 改成 14pt。
```

### 7.2 处理

拆成：

```html
<span data-hdocx-type="run" data-hdocx-id="r-0030a">a</span>
<span data-hdocx-type="run" data-hdocx-id="r-0030b" data-hdocx-font-size="14pt">bcd</span>
<span data-hdocx-type="run" data-hdocx-id="r-0030c">e</span>
```

### 7.3 判定

支持，但必须满足：

- 原 run 只包含可拆分文本。
- 拆分点不落在 field、revision、comment、bookmark、proofing range、surrogate pair、组合字符、emoji 序列中间。
- 拆分后继承原 run 属性。

## 8. 修改跨越多个 run 的文字

### 8.1 场景

原始：

```html
<span data-hdocx-id="r-0040">这是</span>
<span data-hdocx-id="r-0041" data-hdocx-bold="true">重要</span>
<span data-hdocx-id="r-0042">结论</span>
```

用户要求：

```text
把“是重要结”改成红色。
```

### 8.2 处理

编译器需要：

1. 在 `r-0040` 内拆分出“是”。
2. 保留 `r-0041` 原有加粗，同时叠加红色。
3. 在 `r-0042` 内拆分出“结”。
4. 只修改目标文本范围对应的 run。

结果类似：

```html
<span data-hdocx-id="r-0040a">这</span>
<span data-hdocx-id="r-0040b" data-hdocx-color="#ff0000">是</span>
<span data-hdocx-id="r-0041" data-hdocx-bold="true" data-hdocx-color="#ff0000">重要</span>
<span data-hdocx-id="r-0042a" data-hdocx-color="#ff0000">结</span>
<span data-hdocx-id="r-0042b">论</span>
```

### 8.3 判定

支持，但这是比单 run 修改更高风险的 A 类操作。必须执行范围安全检查。

## 9. 修改跨越段落的范围

### 9.1 场景

用户选中两个段落之间的一段文字，要求改字体。

### 9.2 处理

允许，但应拆成多个 paragraph-local patch：

- 第一个段落修改选区起点到段末。
- 中间完整段落修改全部 run。
- 最后一个段落修改段首到选区终点。

### 9.3 限制

如果选区跨越：

- 表格边界
- 域边界
- 修订边界
- 批注边界
- 内容控件边界
- 受保护对象

则必须降级为逐段安全编辑或拒绝。

## 10. 段落样式与局部格式冲突

### 10.1 场景

`Normal` 样式定义为 12pt，但某个 run 直接设置为 16pt。

### 10.2 规则

Word 的实际优先级大致为：

```text
直接 run 属性 > 字符样式 > 段落样式中的 run 属性 > 文档默认值
```

H-DOCX 必须同时记录：

- declared value：XML 中显式声明的值。
- computed value：继承后用于预览的值。
- edit target：本次编辑到底修改哪一层。

### 10.3 判定

支持，但 H-CSS 必须显式声明编辑层：

```css
@hdocx-edit mode(style-definition);
```

或：

```css
@hdocx-edit mode(direct-formatting);
```

不能让 Agent 自己猜。

## 11. 对段落选择器设置字符属性

### 11.1 场景

```css
[data-hdocx-id="p-0100"] {
  hdocx-font-size: 12pt;
}
```

问题：字号是 run 属性，选择器命中的是段落。

### 11.2 规则

H-CSS 必须有明确策略。推荐支持三种模式：

```css
@hdocx-edit mode(paragraph-default-run-properties);
```

含义：修改该段落 `w:pPr/w:rPr`，作为段落默认字符属性。

```css
@hdocx-edit mode(all-runs);
```

含义：修改该段落内所有可编辑 run 的 `w:rPr`。

```css
@hdocx-edit mode(style-definition);
```

含义：如果段落使用某个样式，则修改样式定义中的 run properties。

### 11.3 判定

默认不允许模糊执行。没有模式声明时应报错。

## 12. 对 run 设置段落属性

### 12.1 场景

```css
[data-hdocx-id="r-0100"] {
  hdocx-align: center;
}
```

`hdocx-align` 是段落属性，选择器命中 run。

### 12.2 判定

默认报错。可选支持：

```css
@hdocx-edit lift-to-parent-paragraph(true);
```

但 V1 建议先不支持隐式提升，避免 Agent 误改整段。

## 13. 半个加粗词再改斜体

### 13.1 场景

原文：

```html
<span data-hdocx-id="r-0200" data-hdocx-bold="true">重要结论</span>
```

用户要求：

```text
只把“结论”改成斜体。
```

### 13.2 结果

```html
<span data-hdocx-id="r-0200a" data-hdocx-bold="true">重要</span>
<span data-hdocx-id="r-0200b" data-hdocx-bold="true" data-hdocx-italic="true">结论</span>
```

### 13.3 判定

支持。新 run 继承原加粗属性，再叠加斜体。

## 14. 删除带格式的部分文本

### 14.1 场景

```html
<span data-hdocx-id="r-0300">A</span>
<span data-hdocx-id="r-0301" data-hdocx-bold="true">B</span>
<span data-hdocx-id="r-0302">C</span>
```

用户删除 B。

### 14.2 判定

支持，前提：

- `r-0301` 不包含 protected 子结构。
- 删除不会造成批注、书签、域、修订范围悬空。

否则拒绝，或要求用户明确选择“连同相关结构一起删除”。

## 15. 插入带格式文本

### 15.1 场景

用户在 run 中间插入红色文字。

### 15.2 处理

拆分原 run，并插入新 run：

```html
<span data-hdocx-id="r-0400a">原文前半</span>
<span data-hdocx-id="r-new-0001" data-hdocx-color="#ff0000">插入文字</span>
<span data-hdocx-id="r-0400b">原文后半</span>
```

### 15.3 判定

支持。新 run 应继承插入点上下文的合理默认属性，但 manifest 需要标记为新增节点。

## 16. 合并相邻 run

### 16.1 场景

Agent 为了简化 HTML，把两个相邻 run 合并。

### 16.2 判定

默认不允许无意义合并。

理由：

- 即使视觉格式相同，原始 run 可能有不同 `rsid`、语言、proofing、书签边界或隐藏属性。
- 合并会破坏最小变更原则。

编译器可以在输出阶段做安全合并优化，但 V1 不建议做。

## 17. 空白字符

### 17.1 场景

用户修改含首尾空格、多个连续空格、tab、换行的文本。

### 17.2 规则

- 普通空格必须考虑 `xml:space="preserve"`。
- tab 必须是专用 token，不应变成空格。
- line break/page break/column break 必须是专用节点。
- HTML 渲染中的空白折叠不能影响 DOCX 回写。

### 17.3 判定

支持，但文本编辑器层必须禁用“自动整理空白”的行为。

## 18. Unicode 复杂字符

### 18.1 场景

用户修改 emoji、组合音标、生僻字、代理对、变体选择符。

### 18.2 规则

run split 不能切在 grapheme cluster 中间。编译器需要按 Unicode grapheme cluster 判断编辑边界。

### 18.3 判定

支持，但需要 Unicode 边界算法。V1 如果没有实现，应在复杂字符附近拒绝局部拆分。

## 19. 批注范围内改格式

### 19.1 场景

某个词处于批注范围内，用户要求改字号。

### 19.2 判定

支持，前提是：

- 不移动 comment start/end。
- 不删除 comment reference。
- 不拆坏批注边界。

如果修改范围跨过批注边界，则需要拆成边界内外多个 patch；无法安全拆分时拒绝。

## 20. 修订文本内改格式

### 20.1 场景

被插入修订包裹的文本，用户要求改颜色。

### 20.2 判定

V1 默认拒绝。

理由：

- 这会改变修订记录内部内容。
- 可能生成新的格式修订。
- 不同 Word 版本对修订内修改的行为复杂。

后续可以设计“在保留修订语义的前提下追加新修订”的专门模式。

## 21. 域结果中改文字

### 21.1 场景

用户直接修改目录或交叉引用显示文本。

### 21.2 判定

V1 默认拒绝或要求明确确认。

理由：

- 域结果可能被 Word 更新覆盖。
- 正确修改对象通常是源标题、题注或书签，而不是 field result。

如果用户明确要求改 field result，系统必须标记：

```text
field-result-manually-modified
```

并保留 field code。

## 22. 表格单元格内局部格式

### 22.1 场景

表格某单元格内同一段文字有两个字号。

### 22.2 判定

支持。表格单元格内仍然是 paragraph/run 结构，规则与正文一致。

## 23. 标题中局部格式

### 23.1 场景

一级标题整体使用 Heading1，但标题中的英文缩写用斜体。

### 23.2 判定

支持。

表示：

```html
<p data-hdocx-style-id="Heading1" data-hdocx-outline-level="1">
  <span data-hdocx-id="r-0500">基于 </span>
  <span data-hdocx-id="r-0501" data-hdocx-italic="true">AI</span>
  <span data-hdocx-id="r-0502"> 的研究</span>
</p>
```

标题样式仍属于 paragraph style，局部斜体属于 run direct formatting。

## 24. 脚注中局部格式

### 24.1 场景

脚注正文里一个词加粗。

### 24.2 判定

支持。脚注 part 也用 article/paragraph/run 表示，编辑规则与主文档一致。

## 25. 页眉页脚中局部格式

### 25.1 场景

页眉中学校名称加粗，页码域保持不变。

### 25.2 判定

支持，前提：

- 修改普通 run。
- 页码域 protected。
- 不移动 field boundary。

## 26. 题注中局部格式

### 26.1 场景

图题 “图 1-1 系统结构” 中，“系统结构”改为加粗。

### 26.2 判定

支持，但题注编号域必须 protected。只能改普通文本 run。

## 27. H-CSS 选择器粒度

H-CSS 应支持以下选择器类型。

### 27.1 ID 选择器

```css
[data-hdocx-id="r-0010"] {
  hdocx-font-size: 14pt;
}
```

适合精确修改。

### 27.2 类型选择器

```css
[data-hdocx-type="paragraph"] {
  hdocx-line-spacing: 1.5;
}
```

适合全局修改，但必须谨慎，建议要求确认影响范围。

### 27.3 样式选择器

```css
[data-hdocx-style-id="Normal"] {
  hdocx-first-line-indent: 2char;
}
```

必须配合 `@hdocx-edit mode(...)`。

### 27.4 范围选择器

推荐额外支持命名选择范围：

```css
@hdocx-range target-intro {
  start: r-0100:grapheme(3);
  end: r-0104:grapheme(8);
}

@hdocx-edit mode(range-formatting);

target-intro {
  hdocx-color: #ff0000;
}
```

这用于表达“半句”“几个词”这样的精确局部修改。

## 28. 必须拒绝的编辑

以下编辑默认拒绝：

- 对 protected field/equation/revision/object 直接改文本。
- 对 run 设置段落属性且未声明提升策略。
- 对 paragraph 设置 run 属性且未声明作用模式。
- 删除 range-start 或 range-end。
- 合并包含不同隐藏属性或边界锚点的 run。
- 在 grapheme cluster 中间拆分。
- 跨越多个结构域的模糊选择修改。
- H-CSS 属性无法映射到 OOXML。
- H-CSS 选择器匹配到 protected 节点。

## 29. 最终规则

```text
一个段落可以有任意多个字号。
一个句子可以有任意多种格式。
一个词甚至一个字符也可以单独设置格式。
实现方式是 run split，而不是让 paragraph 承担字符格式。

段落属性只写 paragraph。
字符属性只写 run。
范围修改先拆 run，再 patch。
复杂边界不安全就拒绝。
```

这套规则与 WordprocessingML 的结构一致，因此既能表达真实 Word 文档中的混合格式，又能维持严格双射和最小修改。
