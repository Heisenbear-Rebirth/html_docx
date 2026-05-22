# H-DOCX 选择器复用系统边缘风险与护栏

## 1. 结论

类似 `id`、`class`、命名集合和格式预设的系统是必要的，但它会引入新的错误来源。主要风险不是“表达不了”，而是：

- Agent 选择了错误范围。
- 一个规则意外命中过多节点。
- 多个规则冲突后结果不确定。
- 自定义 class 或集合变成隐藏文档状态。
- run split 后旧选择器失效。
- 样式级修改和直接格式修改混淆。
- 选择器命中了 protected 节点。

因此，选择器复用系统必须是强校验、强类型、失败优先的设计。

```text
宁可拒绝一个模糊选择，也不能默默生成错误 DOCX。
```

## 2. 风险总表

| 边缘情况 | 可能错误 | 覆盖策略 |
|---|---|---|
| `data-hdocx-id` 重复 | patch 写到错误节点 | 转换阶段报错 |
| Agent 修改已有 ID | manifest 对不上 | 回写阶段报错 |
| 选择器匹配 0 个节点 | 用户以为已修改 | 默认报错，可显式允许 |
| 选择器匹配过多节点 | 意外批量修改 | 需要 dry-run 摘要和影响数量 |
| 选择器命中 protected 节点 | 域/修订/公式被破坏 | 默认报错 |
| 多个集合重叠 | 同一属性冲突 | 冲突检测，要求 priority |
| 生成 class 与 data 属性冲突 | 事实来源混乱 | 以 manifest/data 为准，并报 warning |
| Agent 自定义 class | 状态不可追踪 | V1 不允许作为修改语义 |
| run split 后旧 run ID 消失 | 后续规则失效 | 建立 split lineage map |
| 新增节点没有 ID | 无法回写或审计 | 编译器分配临时 ID 并登记 |
| 样式名重复或本地化 | style selector 误判 | 使用 styleId，不用显示名称 |
| Word 样式继承复杂 | 预览值和写入值混淆 | 区分 declared/computed/edit target |
| H-CSS 属性作用层错误 | run/paragraph 写错层 | 类型检查报错 |
| 规则顺序影响结果 | 不可复现 | 禁止网页 CSS 式隐式 cascade |
| 规则跨越边界 | 批注/域/修订错位 | range 安全检查，不安全则拒绝 |

## 3. ID 相关边缘情况

### 3.1 重复 ID

错误示例：

```html
<p data-hdocx-id="p-001">...</p>
<p data-hdocx-id="p-001">...</p>
```

风险：回写时无法判断 patch 目标。

规则：

- H-DOCX bundle 内 `data-hdocx-id` 必须唯一。
- 如果采用 part-local ID，则必须用 `part path + local id` 组合唯一。
- 发现重复 ID 必须立即失败。

### 3.2 Agent 修改 ID

风险：HTML 看起来正常，但 manifest 中原节点 hash、locator、lock 状态全部失效。

规则：

- 已存在节点的 `data-hdocx-id` 是只读属性。
- 如果 HTML 中 ID 与 manifest 不一致，回写失败。
- 新增节点不能由 Agent 手写最终 ID，应使用临时标记，编译器分配正式 ID。

推荐新增节点写法：

```html
<span data-hdocx-type="run" data-hdocx-new="true">新增文字</span>
```

编译器输出时分配：

```html
<span data-hdocx-id="r-new-0001" data-hdocx-origin="inserted">新增文字</span>
```

### 3.3 run split 后 ID 变化

场景：

```html
<span data-hdocx-id="r-100">重要结论</span>
```

被拆成：

```html
<span data-hdocx-id="r-100a">重要</span>
<span data-hdocx-id="r-100b">结论</span>
```

风险：后续 H-CSS 仍引用 `r-100`。

规则：

- manifest 必须记录 split lineage：

```json
{
  "splitFrom": "r-100",
  "children": ["r-100a", "r-100b"]
}
```

- 如果后续规则命中已拆分的旧 ID，应报错并提示可用新 ID。
- 不建议自动把旧 ID 展开到所有子 run，除非规则显式声明：

```css
@hdocx-resolve-split include-children;
```

## 4. class 相关边缘情况

### 4.1 生成 class 与 data 属性冲突

错误示例：

```html
<p
  class="hstyle-Heading1"
  data-hdocx-style-id="Normal">
```

风险：Agent 或预览器误以为这是标题。

规则：

- 回写时以 `data-hdocx-style-id` 和 manifest 为准。
- 生成 class 可以重新生成，不作为事实来源。
- 如果 class 与 data 属性冲突，应发出 warning；如果 H-CSS 依赖冲突 class，应报错。

### 4.2 Agent 自定义 class 表达修改

错误示例：

```html
<p class="make-red">正文</p>
```

风险：`make-red` 的语义不在 manifest 中，无法严格双射。

规则：

- V1 不允许普通 HTML class 表示 DOCX 修改。
- 如果要复用选择器，使用 `@hdocx-set`。
- 如果要复用格式，使用 `@hdocx-format`。

### 4.3 class 名称非法或碰撞

风险：Word styleId 可能包含空格、中文、特殊字符，直接变成 class 会碰撞。

规则：

- 生成 class 必须经过稳定编码，例如：

```text
hstyle-Normal
hstyle-encoded-<stable-hash>
```

- manifest 记录 class 到 styleId 的映射。
- H-CSS 推荐直接用 `[data-hdocx-style-id="..."]`，而不是依赖 class。

## 5. `@hdocx-set` 边缘情况

### 5.1 集合匹配 0 个节点

示例：

```css
@hdocx-set body {
  select: [data-hdocx-style-id="BodyText"];
}
```

如果文档没有 `BodyText`。

规则：

- 默认报错。
- 如果用户确实允许空匹配，必须显式声明：

```css
@hdocx-set body {
  select: [data-hdocx-style-id="BodyText"];
  allow-empty: true;
}
```

### 5.2 集合匹配过多节点

示例：

```css
@hdocx-set all {
  select: [data-hdocx-type="paragraph"];
}
```

风险：正文、标题、目录、脚注、批注都被改。

规则：

- 编译前必须输出 dry-run 影响摘要：

```text
set all matched:
- main document paragraphs: 120
- headings: 8
- footnotes: 14
- comments: 3
- protected fields: 2
```

- 如果集合命中 protected 节点，默认失败。
- 大范围选择器建议要求显式确认或 `max-match`：

```css
@hdocx-set body {
  select: [data-hdocx-style-id="Normal"];
  max-match: 300;
}
```

### 5.3 集合递归引用

错误示例：

```css
@hdocx-set a { include: b; }
@hdocx-set b { include: a; }
```

规则：

- 检测引用环。
- 发现环必须报错。

### 5.4 集合差集误伤

示例：

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
  exclude: [data-hdocx-style-id="Heading1"];
}
```

风险：仍然命中脚注、目录、批注、页眉页脚。

规则：

- 集合应支持 part 维度过滤：

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"][data-hdocx-style-id="Normal"];
  within: /word/document.xml;
  exclude: [data-hdocx-lock="protected"];
}
```

- V1 推荐所有批量集合都必须显式声明 `within` 或 part scope。

## 6. `@hdocx-format` 边缘情况

### 6.1 格式预设包含不同层级属性

错误示例：

```css
@hdocx-format mixed {
  hdocx-font-size: 12pt;
  hdocx-align: center;
}
```

这同时包含 run 属性和 paragraph 属性。

规则：

- 格式预设必须声明 target kind：

```css
@hdocx-format body-paragraph target(paragraph) {
  hdocx-align: justify;
  hdocx-line-spacing: 1.5;
}

@hdocx-format body-run target(run) {
  hdocx-font-size: 12pt;
}
```

- 如果不声明 target，编译器根据属性推断；推断出多层级属性时必须报错。

### 6.2 include 后属性冲突

示例：

```css
@hdocx-format a { hdocx-font-size: 12pt; }
@hdocx-format b { hdocx-font-size: 14pt; }

body {
  @hdocx-include a;
  @hdocx-include b;
}
```

规则：

- 同一规则内 include 冲突必须报错。
- 允许显式覆盖：

```css
@hdocx-include a;
@hdocx-include b override;
```

V1 可以先不支持 override，全部冲突报错。

### 6.3 token 类型不匹配

示例：

```css
@hdocx-token size.body "large";

body {
  hdocx-font-size: token(size.body);
}
```

规则：

- token 必须有类型校验。
- `hdocx-font-size` 只接受 point、half-point 或预定义中文字号。
- 类型不匹配必须报错。

## 7. 选择器与属性层级错配

### 7.1 段落选择器设置 run 属性

示例：

```css
[data-hdocx-id="p-1"] {
  hdocx-font-size: 12pt;
}
```

风险：不知道是改 `w:pPr/w:rPr`，还是改段内所有 runs。

规则：必须显式 mode。

```css
@hdocx-edit mode(paragraph-default-run-properties);
```

或：

```css
@hdocx-edit mode(all-runs);
```

没有 mode 时失败。

### 7.2 run 选择器设置段落属性

示例：

```css
[data-hdocx-id="r-1"] {
  hdocx-align: center;
}
```

规则：

- V1 默认报错。
- 不自动提升到父段落。

## 8. protected 节点命中

### 8.1 格式命中 protected field result

风险：目录、交叉引用、页码域被改坏。

规则：

- H-CSS 命中 protected 节点即失败。
- 如果用户只想改 field result 外观，必须使用专门模式：

```css
@hdocx-edit mode(field-result-formatting) acknowledge-field-risk;
```

V1 可以不支持该模式。

### 8.2 集合包含修订文本

风险：修改修订内部文本会改变审阅语义。

规则：

- V1 默认 protected revision 内部所有节点。
- 批量格式不得进入 revision。

## 9. 范围选择边缘情况

### 9.1 range 切在复杂字符中间

风险：Unicode 字符被破坏。

规则：

- range start/end 必须以 grapheme cluster 为单位。
- 不能用 UTF-16 code unit 作为用户级边界。

### 9.2 range 跨 protected boundary

风险：批注、域、书签、修订错位。

规则：

- 编译器应尝试拆成多个安全子范围。
- 如果无法拆分，拒绝。

### 9.3 range 锚点文本已改变

风险：之前定义的 range 不再指向原文本。

规则：

- range 定义应包含 anchor hash。
- anchor hash 不匹配时失败，要求重新定位。

## 10. 样式系统边缘情况

### 10.1 styleId 与显示名称不同

Word 样式显示名可能本地化，例如“正文”“Normal”。

规则：

- 选择器必须优先使用 styleId。
- 显示名称只能作为辅助属性。

### 10.2 styleId 缺失或重复

风险：不合规文档或 Word 兼容文档可能异常。

规则：

- 正常情况下 styleId 应唯一。
- 如果发现重复 styleId，样式级编辑必须拒绝。
- 仍可对具体节点做直接格式编辑。

### 10.3 修改 style definition 导致大量节点变化

风险：用户以为只改一段，实际全局受影响。

规则：

- `mode(style-definition)` 必须输出影响摘要。
- 摘要应列出受该 style 影响的段落数量和 part。

## 11. 顺序与优先级边缘情况

### 11.1 规则顺序导致隐式覆盖

网页 CSS 允许后写覆盖先写，但 H-DOCX 不应默认这样做。

规则：

- 同一属性多处赋不同值时默认报错。
- 只有显式 priority 或 override 才允许覆盖。

### 11.2 命名集合之间优先级

规则：

```css
@hdocx-priority heading-1 100;
@hdocx-priority body 10;
```

V1 可以暂不支持 priority，改为冲突即失败。

## 12. 严格双射边缘情况

### 12.1 H-CSS 编辑脚本是否进入 DOCX

默认不进入 DOCX。

风险：`H-DOCX -> DOCX -> H-DOCX` 后 Agent 自定义集合丢失。

规则：

- 这不破坏 DOCX 内容双射，因为集合不是 DOCX 原始内容。
- 如果用户要求保留 Agent 元数据，必须显式写入 custom XML part。
- V1 默认不写 custom XML，避免污染用户文档。

### 12.2 generated class 是否进入 DOCX

不进入。

规则：

- generated class 可由 DOCX 重新生成。
- 不作为持久事实。

### 12.3 修改后能否回到完全相同的 H-DOCX

严格双射的主目标是 DOCX 与 H-DOCX bundle 的双射。修改后如果重新导出，生成类和 ID 应稳定，但 Agent 临时脚本不一定保留，除非它被纳入 bundle。

V1 建议：

- H-DOCX bundle 内保留 `agent.edits.hcss` 作为审计记录。
- 输出 DOCX 默认不携带 `agent.edits.hcss`。

## 13. 必须实现的护栏

V1 最少需要以下护栏：

1. ID 唯一性检查。
2. manifest 与 HTML ID/hash 一致性检查。
3. H-CSS 选择器匹配数量统计。
4. 0 匹配默认报错。
5. protected 命中默认报错。
6. 属性层级类型检查。
7. H-CSS 属性白名单。
8. H-CSS 单位类型检查。
9. 冲突属性默认报错。
10. run split 安全检查。
11. range boundary 安全检查。
12. style-definition 修改影响摘要。
13. 未修改 XML/package hash 校验。

## 14. 无法完全覆盖但可安全处理的情况

以下情况 V1 可能无法语义编辑，但可以安全保真或拒绝：

- 修订链内部批量格式修改。
- 域结果局部编辑。
- SmartArt 内部文本精确修改。
- 图表数据格式联动修改。
- OLE 对象内部内容修改。
- 受保护文档中的越权编辑。
- 数字签名文档编辑。
- 异常或非规范 OOXML 中的样式定义批量修改。

处理策略：

```text
能透明保真的，保真。
不能安全编辑的，拒绝。
不能保真的，拒绝输入。
```

## 15. 最终判断

这套选择器与复用系统本身可以覆盖 Agent 常见书写需求，但如果没有护栏，会产生严重边缘错误。

因此最终设计应定为：

```text
允许类似 class/id 的复用能力。
禁止普通 class 成为 DOCX 修改语义。
所有选择器都必须可解释、可统计、可审计。
所有属性都必须有明确 OOXML 写入层。
所有冲突和模糊情况默认失败。
```

这样才能既减少 Agent 重复书写，又不牺牲严格双射。
