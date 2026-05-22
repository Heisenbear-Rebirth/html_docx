# H-DOCX 选择器与复用系统设计

## 1. 结论

有必要引入类似 `id`、`class` 的系统，而且这是让 Agent 高效书写 H-DOCX 修改指令的关键。

但是该系统不能照搬普通 HTML 的 `id/class` 语义。H-DOCX 应区分四类东西：

| 层级 | 名称 | 作用 | 是否属于 DOCX 事实来源 |
|---|---|---|---|
| 1 | 稳定节点 ID | 精确定位 DOCX 节点 | 是 |
| 2 | 生成语义类 | 基于 DOCX 结构生成，便于选择和预览 | 是，因可由 DOCX 确定生成 |
| 3 | Agent 命名集合 | 避免重复写复杂选择器 | 否，属于编辑脚本 |
| 4 | Agent 格式预设 | 避免重复写格式声明 | 否，属于编辑脚本 |

核心原则：

```text
data-hdocx-id 用于身份。
生成 class 用于可读和选择。
Agent 自定义集合用于复用选择器。
Agent 格式预设用于复用声明。
只有可回写的 H-CSS 声明会变成 OOXML patch。
```

## 2. 为什么需要这套系统

如果没有类似 `class/id` 的机制，Agent 会被迫反复书写大量选择器和格式声明，例如：

```css
[data-hdocx-style-id="Normal"][data-hdocx-type="paragraph"]:not([data-hdocx-lock="protected"]) {
  hdocx-font-family-east-asia: "SimSun";
  hdocx-font-family-ascii: "Times New Roman";
  hdocx-font-size: 12pt;
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}
```

这会带来几个问题：

- Agent 更容易写错选择器。
- 同一套格式要复制多次。
- 修改范围不容易审计。
- 用户说“把正文统一改成……”时，需要一个稳定的命名目标。

因此，H-DOCX 需要提供类似类名、变量、格式预设和命名范围的能力。

## 3. 稳定 ID

### 3.1 设计

每个可编辑或受保护的关键节点都必须有 `data-hdocx-id`。

示例：

```html
<p
  data-hdocx-type="paragraph"
  data-hdocx-id="p-00042"
  data-hdocx-style-id="Normal">
  <span data-hdocx-type="run" data-hdocx-id="r-00120">正文内容</span>
</p>
```

`data-hdocx-id` 是 H-DOCX 的权威节点身份，不建议依赖普通 HTML `id`。

可以为了浏览器锚点同步生成普通 `id`：

```html
<p id="hdocx-p-00042" data-hdocx-id="p-00042">
```

但普通 `id` 只是辅助；回写 DOCX 时以 `data-hdocx-id` 和 manifest 为准。

### 3.2 规则

- `data-hdocx-id` 由转换器生成。
- Agent 不应修改已有 `data-hdocx-id`。
- 新增节点由编译器分配新 ID。
- ID 必须在 H-DOCX bundle 内唯一，或通过 part path + local id 唯一。
- ID 对应 manifest 中的 part、XML locator、hash、node kind、lock 状态。

## 4. 生成语义类

### 4.1 设计

H-DOCX 可以生成普通 HTML `class`，用于预览、阅读和选择。但这些 class 必须是从 DOCX 结构确定生成的，不应由 Agent 随意发明。

示例：

```html
<p
  class="hdocx-p hpart-main hstyle-Normal hlock-editable"
  data-hdocx-type="paragraph"
  data-hdocx-id="p-00042"
  data-hdocx-style-id="Normal">
```

推荐生成类：

- `hdocx-p`
- `hdocx-r`
- `hdocx-table`
- `hdocx-cell`
- `hpart-main`
- `hpart-header`
- `hpart-footer`
- `hstyle-Normal`
- `hstyle-Heading1`
- `hlist-level-0`
- `hlock-editable`
- `hlock-protected`
- `hfield`
- `hcomment-anchor`
- `hrevision`

### 4.2 严格双射影响

生成类不破坏严格双射，因为它们是 DOCX 结构的确定投影。

也就是说：

```text
DOCX -> H-DOCX
```

时可以重新生成这些类。它们不是额外文档内容。

### 4.3 限制

- Agent 不应通过修改普通 `class` 改变 DOCX。
- 回写 DOCX 时，普通 `class` 不作为事实依据。
- 若 `class` 与 `data-hdocx-*` 冲突，以 `data-hdocx-*` 和 manifest 为准。

## 5. Agent 命名集合

### 5.1 设计

为了避免重复写复杂选择器，H-CSS 应支持命名集合。

建议语法：

```css
@hdocx-set body-paragraphs {
  select: [data-hdocx-type="paragraph"][data-hdocx-style-id="Normal"];
  exclude: [data-hdocx-lock="protected"];
}

@hdocx-set chapter-headings {
  select: [data-hdocx-type="paragraph"][data-hdocx-style-id="Heading1"];
}
```

之后可以直接使用集合名：

```css
@hdocx-edit mode(style-definition);

body-paragraphs {
  hdocx-font-size: 12pt;
  hdocx-line-spacing: 1.5;
}
```

### 5.2 集合不是 HTML class

`@hdocx-set` 不一定要把类名写回 HTML 节点。

这是重要边界：

- 如果 Agent 只是为了复用选择器，不需要污染 `document.html`。
- 集合属于 `agent.edits.hcss` 的编辑脚本。
- 集合本身不进入 DOCX，只有集合命中的节点上的 H-CSS 声明会编译为 OOXML patch。

### 5.3 持久化选择

如果确实需要长期保存 Agent 命名集合，有两个选择：

1. 保存在 H-DOCX bundle 的 sidecar 文件，例如 `agent.sets.hjson`。
2. 如果要求 `H-DOCX -> DOCX -> H-DOCX` 后仍保留这些集合，则必须写入 DOCX 的 custom XML part，并明确标记为 H-DOCX 元数据。

默认 V1 建议使用第 1 种。不要默认污染用户 DOCX。

## 6. Agent 格式预设

### 6.1 设计

为了避免重复写格式声明，H-CSS 应支持命名格式预设。

示例：

```css
@hdocx-format thesis-body-text {
  hdocx-font-family-east-asia: "SimSun";
  hdocx-font-family-ascii: "Times New Roman";
  hdocx-font-size: 12pt;
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
  hdocx-space-before: 0pt;
  hdocx-space-after: 0pt;
}

@hdocx-edit mode(style-definition);

body-paragraphs {
  @hdocx-include thesis-body-text;
}
```

### 6.2 格式预设与 Word 样式的区别

`@hdocx-format` 不是 Word style。

它只是 H-CSS 中的一组可复用声明。它可以被编译到：

- Word 样式定义；
- 段落直接格式；
- run 直接格式；
- 表格样式或表格直接格式；

具体写到哪里，必须由 `@hdocx-edit mode(...)` 决定。

### 6.3 映射为 Word 样式

如果用户明确希望创建或修改 Word 样式，应使用显式声明：

```css
@hdocx-word-style BodyText {
  type: paragraph;
  name: "Body Text";
  based-on: Normal;
}

@hdocx-map-format thesis-body-text to-style BodyText;
```

V1 可以先不支持创建新 Word 样式，但应支持修改已有 style definition。

## 7. 变量与设计 token

为了减少重复数值，H-CSS 可以支持 token。

示例：

```css
@hdocx-token font.cn.body "SimSun";
@hdocx-token font.en.body "Times New Roman";
@hdocx-token size.body 12pt;
@hdocx-token line.body 1.5;
@hdocx-token indent.first-line 2char;

@hdocx-format thesis-body-text {
  hdocx-font-family-east-asia: token(font.cn.body);
  hdocx-font-family-ascii: token(font.en.body);
  hdocx-font-size: token(size.body);
  hdocx-line-spacing: token(line.body);
  hdocx-first-line-indent: token(indent.first-line);
}
```

规则：

- token 必须在编译期解析成具体 H-CSS 值。
- token 值必须符合 H-CSS 类型系统。
- 未使用 token 不影响 DOCX。
- token 不是 Word theme，除非显式映射。

## 8. 命名范围

有些 Agent 操作不是按现有样式选择，而是按用户选区选择，例如“把摘要第二句改成斜体”。

H-CSS 应支持命名范围：

```css
@hdocx-range abstract-second-sentence {
  start: r-01000:grapheme(0);
  end: r-01004:grapheme(12);
}

@hdocx-edit mode(range-formatting);

abstract-second-sentence {
  hdocx-italic: true;
}
```

规则：

- range 必须以 grapheme cluster 为边界。
- range 可以跨 run。
- range 跨段落时必须拆成 paragraph-local patch。
- range 不能破坏 protected boundary。

## 9. 推荐 H-CSS 示例

用户要求：

```text
把正文统一改为宋体/Times New Roman，小四，1.5 倍行距，首行缩进 2 字符；
一级标题居中加粗；
摘要中的关键词加粗。
```

推荐写法：

```css
@hdocx-token font.cn.body "SimSun";
@hdocx-token font.en.body "Times New Roman";
@hdocx-token size.body 12pt;

@hdocx-set body-paragraphs {
  select: [data-hdocx-type="paragraph"][data-hdocx-style-id="Normal"];
  exclude: [data-hdocx-lock="protected"];
}

@hdocx-set heading-1 {
  select: [data-hdocx-type="paragraph"][data-hdocx-style-id="Heading1"];
}

@hdocx-range abstract-keywords {
  start: r-02010:grapheme(0);
  end: r-02012:grapheme(6);
}

@hdocx-format body-text {
  hdocx-font-family-east-asia: token(font.cn.body);
  hdocx-font-family-ascii: token(font.en.body);
  hdocx-font-size: token(size.body);
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}

@hdocx-edit mode(style-definition);

body-paragraphs {
  @hdocx-include body-text;
}

heading-1 {
  hdocx-align: center;
  hdocx-bold: true;
}

@hdocx-edit mode(range-formatting);

abstract-keywords {
  hdocx-bold: true;
}
```

## 10. 冲突规则

H-CSS 必须定义确定的冲突规则。

推荐优先级：

```text
explicit data-hdocx-id selector
> named range
> named set with explicit priority
> generated class/style selector
> type selector
```

如果两个规则同级且设置同一属性为不同值，应报错，除非显式声明：

```css
@hdocx-priority heading-1 100;
@hdocx-priority body-paragraphs 10;
```

不建议完全照搬网页 CSS cascade，因为 Word 样式系统和 DOCX patch 语义不同。

## 10.1 边缘护栏

选择器与复用系统必须配套强校验。至少包括：

- ID 唯一性检查。
- manifest 与 HTML ID/hash 一致性检查。
- 选择器匹配数量统计。
- 0 匹配默认报错。
- protected 节点命中默认报错。
- H-CSS 属性层级类型检查。
- H-CSS 单位类型检查。
- 属性冲突默认报错。
- run split 安全检查。
- range boundary 安全检查。
- `mode(style-definition)` 的影响摘要。

详细失败模式和处理规则见 `SELECTOR_EDGE_CASES_AND_GUARDS.md`。

## 11. 与严格双射的关系

### 11.1 不破坏双射的部分

以下内容不破坏 DOCX -> H-DOCX -> DOCX：

- `data-hdocx-id`，因为它映射到 manifest 中的原始 DOCX 节点。
- 生成 class，因为它可由 DOCX 确定生成。
- H-CSS 命名集合和格式预设，因为它们是编辑脚本，不是 DOCX 原文。
- token，因为它们只是编译期值。

### 11.2 需要谨慎的部分

如果希望 Agent 自定义 class 成为文档状态，并且要求 `H-DOCX -> DOCX -> H-DOCX` 后仍保留，那么必须把这些 class 作为 H-DOCX 元数据写入 DOCX custom XML part。

否则它们只能作为编辑脚本存在，不属于最终 DOCX 文档语义。

### 11.3 默认策略

V1 默认策略：

```text
不允许 Agent 通过随意修改 HTML class 来表达 DOCX 修改。
允许 Agent 在 H-CSS 中定义 @hdocx-set、@hdocx-format、@hdocx-token。
这些定义只参与编译，不默认写入 DOCX。
```

## 12. V1 建议实现

V1 应支持：

- `data-hdocx-id` 精确定位。
- 生成 `class`，例如 `hstyle-*`、`hdocx-*`、`hlock-*`。
- H-CSS `@hdocx-set`。
- H-CSS `@hdocx-format`。
- H-CSS `@hdocx-token`。
- H-CSS `@hdocx-include`。
- H-CSS `@hdocx-edit mode(...)`。
- 冲突检测。

V1 可以暂缓：

- 创建新的 Word style。
- 将 Agent 自定义集合写入 DOCX custom XML。
- 复杂优先级系统。
- 类似 Sass 的复杂嵌套。

## 13. 最终规则

```text
需要类似 class/id 的能力。
id 用于定位，不用于复用。
生成 class 用于阅读和批量选择，不作为独立事实来源。
Agent 复用选择器用 @hdocx-set。
Agent 复用格式声明用 @hdocx-format。
Agent 复用数值用 @hdocx-token。
不要让普通 HTML class 承担 DOCX 语义。
```

这样既能让 Agent 写得简洁，又不会因为引入额外 HTML 语义而破坏严格双射。
