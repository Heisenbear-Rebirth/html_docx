# html_docx 全局交付计划

## 1. 目标

本计划的终点是交付一款完整可用的软件，使 Codex、Claude Code 等 Agent 能够安全编辑 DOCX，同时满足项目核心原则：

```text
DOCX <-> H-DOCX bundle

可编辑的，必须精确 patch。
不可编辑的，必须严格保真。
不能证明安全的，必须拒绝。
```

这里的“完整可用软件”不是普通格式转换器，而是一个面向 Agent 的 DOCX 可逆编辑系统。它应包含：

- 命令行工具。
- Python 库 API。
- H-DOCX bundle 格式。
- HTML 编辑投影。
- H-CSS 自定义样式 DSL。
- 严格 manifest 校验。
- OOXML patch 回写。
- DOCX diff/validate/report 工具。
- 测试夹具和真实文档验证链。
- 面向 Agent 的使用文档和错误报告。

## 2. 当前状态

当前已经完成从 Phase 1 到若干核心可编辑能力的第一轮实现：

```text
DOCX -> H-DOCX bundle -> DOCX
```

已具备：

- Python 包骨架。
- CLI 入口。
- `export` / `validate` / `plan` / `apply` / `roundtrip` / `diff` / `inspect` 基础命令。
- 原始 DOCX bytes 保存。
- 未修改 `apply` 直接复制原始 DOCX。
- 最小 `document.html` 投影。
- 最小 `manifest.json`。
- 字节级 roundtrip 测试。
- 普通 run 文本修改。
- run 级 bold/italic/font-size/color 修改。
- `run-segment` 局部 run split。
- 段落 align/line-spacing/first-line-indent 修改。
- 表格基础投影和单元格文本修改。
- 页眉、脚注等 secondary parts 投影和文本修改。
- H-CSS V1 基础能力，包括 set、token、format、include、paragraph-formatting、all-runs、style-definition。
- CLI `--report` JSON 文件输出。

当前限制：

- XML 修改仍主要通过 ElementTree 重写被修改 part，尚未实现 byte-range fragment patch。
- 批注、修订、公式、图表、SmartArt、OLE 等高级结构仍以保护/保真为主，专门投影和编辑尚未完成。
- numbering/list、图片受控替换、完整 diff/report、真实大型 DOCX 压力测试尚未完成。

## 3. 最终交付形态

### 3.1 CLI

最终 CLI 至少包含：

```powershell
python -m html_docx export input.docx --out work.hdocx
python -m html_docx validate work.hdocx
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx
python -m html_docx diff input.docx output.docx --report diff.json
python -m html_docx inspect work.hdocx --id p-000001
python -m html_docx roundtrip input.docx --work work.hdocx --out roundtrip.docx
```

### 3.2 Python API

最终库 API 至少包含：

```python
export_docx(input_docx, output_hdocx)
validate_hdocx(hdocx_dir)
plan_hdocx(hdocx_dir)
apply_hdocx(hdocx_dir, output_docx)
diff_docx(left_docx, right_docx)
inspect_node(hdocx_dir, node_id)
```

### 3.3 Agent 工作流

Agent 的标准工作流：

```text
1. export DOCX
2. 阅读 document.html / manifest / plan
3. 修改 document.html 或 agent.edits.hcss
4. plan 查看影响范围
5. apply 输出 DOCX
6. diff/validate 证明修改范围
```

### 3.4 最终软件边界

最终版本必须支持：

- 未修改 DOCX 严格往返。
- 正文文本修改。
- run 级字符格式修改。
- 段落级格式修改。
- 一个段落内多字号、多字体、多颜色。
- run split。
- 表格单元格文本和基础格式修改。
- 页眉页脚文本修改。
- 脚注尾注文本修改。
- 基础图片元数据和受控替换。
- H-CSS 批量修改字体、字号、行距、缩进、对齐、段前段后。
- styles.xml 中已有样式的受控修改。
- 目录、域、公式、批注、修订、图表、SmartArt、OLE、未知扩展的严格保真。
- 受保护或不安全修改拒绝。
- 清晰可机读错误报告。

## 4. 总体开发策略

开发顺序遵循：

```text
先证明严格保真。
再支持最小文本编辑。
再支持局部格式。
再支持批量 H-CSS。
再扩展复杂结构。
最后做真实文档压力测试和交付包装。
```

每个阶段必须有：

- 可运行代码。
- 自动化测试。
- CLI 烟测。
- 明确失败行为。
- 文档同步。

不允许为了“看起来能转换”牺牲严格双射。

## 5. 阶段计划

## Phase 1: Package Roundtrip 基础闭环

状态：基本完成。

目标：

- 读取 DOCX package。
- 保存原始 bytes。
- 生成 H-DOCX bundle。
- 未修改时输出字节级相同 DOCX。

已完成：

- package 读取。
- entries manifest。
- original.docx 保存。
- 未修改 roundtrip。
- 单元测试。

剩余增强：

- 更完整的 ZIP entry metadata。
- 加密文档识别。
- 损坏 package 错误细化。
- `diff` 输出 JSON report 文件。

验收：

- `roundtrip` 对最小 DOCX 字节一致。
- `diff` 显示无差异。
- manifest 缺失时拒绝。

## Phase 2: 普通文本回写

目标：

实现第一个真实编辑闭环：

```text
修改 document.html 中普通 run 文本
-> 回写 word/document.xml 中对应 w:t
-> 输出 DOCX
```

任务：

- 在 manifest 中记录 paragraph/run 的结构定位。
- 记录 run 是否 simple-editable。
- 检测 `document.html` 中 run 文本变化。
- 只允许 editable/simple run 修改。
- 回写对应 `w:t`。
- 正确处理 `xml:space="preserve"`。
- 输出 patch report。

必须拒绝：

- 修改 protected run。
- 删除或改写 `data-hdocx-id`。
- 一个 run 内有多个复杂子节点时盲目修改。
- 修改 field/drawing/object/instrText。

验收：

- 修改一个普通 run 文本成功。
- 非目标 entries 不变。
- `word/document.xml` 只有目标文本变化。
- protected run 修改失败。
- ID 篡改失败。

## Phase 3: Run Split 与局部字符格式

目标：

支持一个词、半句、几个字符的格式修改。

任务：

- 实现 Unicode grapheme cluster 边界检测。
- 实现 run split。
- 新 run 继承原 run properties。
- 支持 run 属性：
  - 字号。
  - 字体。
  - 中文字体。
  - 西文字体。
  - 加粗。
  - 斜体。
  - 下划线。
  - 颜色。
  - 高亮。
  - 上下标。
- 支持字符级 direct formatting。
- 支持 range-formatting 的内部数据模型。

必须拒绝：

- 拆分 field/revision/comment boundary。
- 在 grapheme cluster 中间拆分。
- 修改修订内部文本。
- 修改域结果，除非有专门确认模式。

验收：

- 一个段落内两个字号。
- 一个词加粗。
- 半句改颜色。
- 跨多个 run 改同一格式。
- 格式修改不影响非目标 run。

## Phase 4: 段落格式回写

目标：

支持常用段落排版修改。

任务：

- 支持 `w:pPr` patch。
- 支持：
  - 对齐。
  - 左右缩进。
  - 首行缩进。
  - 悬挂缩进。
  - 段前段后。
  - 行距。
  - 分页前。
  - 与下段同页。
  - 段落样式应用。
- 区分 declared value 与 computed value。
- 区分 direct formatting 与 style-definition。

验收：

- 单段居中。
- 正文首行缩进 2 字符。
- 正文 1.5 倍行距。
- 标题段前段后修改。
- run 属性不被误写到 paragraph。
- paragraph 属性不被误写到 run。

## Phase 5: H-CSS V1

目标：

实现 Agent 友好的批量排版 DSL。

任务：

- 实现 H-CSS parser。
- 支持：
  - `@hdocx-token`
  - `@hdocx-set`
  - `@hdocx-format`
  - `@hdocx-include`
  - `@hdocx-range`
  - `@hdocx-edit mode(...)`
- 支持选择器匹配。
- 支持属性白名单。
- 支持单位类型系统。
- 支持冲突检测。
- 支持 0 匹配默认失败。
- 支持 protected 命中默认失败。
- 支持 plan 阶段影响摘要。

验收：

- 用 H-CSS 批量设置正文小四、1.5 倍行距、首行缩进。
- 用 H-CSS 设置标题居中加粗。
- 用 H-CSS 修改命名 range。
- 冲突规则失败。
- protected 命中失败。
- 0 匹配失败。

## Phase 6: 样式系统

目标：

支持已有 Word style 的安全读取、应用和修改。

任务：

- 解析 `word/styles.xml`。
- 建立 style map。
- 支持 paragraph style、character style、table style。
- 支持 basedOn、next、linked 的保真。
- 支持修改已有 style definition 的常用属性。
- 支持 style-definition 影响摘要。

暂缓：

- 新建复杂 Word style。
- 自动重构样式继承链。
- 主题字体深度编辑。

验收：

- 修改 Normal 样式字号。
- 修改 Heading1 对齐和加粗。
- 应用已有样式到段落。
- styleId 重复时拒绝样式级编辑。
- latent styles 保真。

## Phase 7: 表格支持

目标：

支持论文常见表格编辑。

任务：

- 投影 table/tr/tc。
- manifest 记录 tblGrid、gridSpan、vMerge。
- 支持单元格文本编辑。
- 支持单元格内 run/paragraph 格式。
- 支持基础表格属性：
  - 边框。
  - 底纹。
  - 宽度。
  - 单元格对齐。
  - 表格样式应用。

暂缓：

- 自动重算复杂 tblGrid。
- 任意插入删除复杂合并行列。

验收：

- 表格单元格内一个词加粗。
- 单元格文本修改只影响目标 cell。
- 横向/纵向合并表格未修改保真。
- 嵌套表格未修改保真。

## Phase 8: Secondary Parts

目标：

支持主文档以外的学术关键 parts。

任务：

- header part 投影和编辑。
- footer part 投影和编辑。
- footnotes part 投影和编辑。
- endnotes part 投影和编辑。
- comments part 投影，V1 默认只读。
- part 之间 relationship 保真。

验收：

- 页眉文本修改成功。
- 页脚页码域 protected。
- 脚注正文文本修改成功。
- 尾注正文文本修改成功。
- 批注正文默认只读保真。

## Phase 9: 学术结构只读保真

目标：

对学术写作核心高级结构建立可靠保护层。

任务：

- field map：
  - TOC。
  - REF。
  - PAGEREF。
  - SEQ。
  - PAGE。
  - Citation/Bibliography。
- equation map。
- caption/bookmark/cross-reference map。
- comment range map。
- revision range map。
- protected node 占位。
- boundary anchor。

验收：

- 含目录文档未修改 roundtrip 保真。
- 含交叉引用文档未修改保真。
- 含公式文档未修改保真。
- 含批注文档未修改保真。
- 含修订文档未修改保真。
- 试图修改 protected field/equation/revision 失败。

## Phase 10: 图片、绘图与嵌入对象

目标：

支持图片基础编辑，并对复杂对象严格保真。

任务：

- 解析 image relationship。
- 投影 `<img>` preview。
- 支持 alt text 修改。
- 支持简单尺寸修改。
- 支持受控图片替换。
- DrawingML/VML 原始结构保真。
- chart、SmartArt、OLE、text box 默认 protected。

验收：

- 普通图片未修改保真。
- 修改图片 alt text 成功。
- 替换图片时关系正确。
- chart/SmartArt/OLE 未修改保真。
- 修改复杂对象内容默认失败。

## Phase 11: Fragment Patch 强化

目标：

降低 XML 重新序列化导致的非目标 diff。

任务：

- fragment-level patcher。
- source fragment byte range 或稳定结构 locator。
- 未修改 XML fragment 原样拼回。
- namespace prefix 保留。
- 属性顺序、空白、XML comment、PI 保留。
- `mc:AlternateContent` 不折叠。

验收：

- 局部文本修改时非目标 fragment hash 不变。
- unknown namespace 不丢。
- `mc:AlternateContent` 未修改保真。
- Word `rsid` 不被清理。

## Phase 12: Diff、Plan 与审计报告

目标：

让 Agent 和用户能明确知道软件做了什么。

任务：

- `plan` 输出 patch plan。
- `diff` 输出 package/part/node/fragment 四层 diff。
- `validate` 输出结构化错误。
- `audit.log.jsonl` 完整记录。
- 错误码体系稳定。

验收：

- 用户能看到将修改哪些节点。
- 用户能看到 style-definition 影响多少段。
- 用户能看到失败原因和修正建议。
- 所有 CLI 输出机器可读 JSON。

## Phase 13: 真实文档压力测试

目标：

用真实学术 DOCX 验证可用性。

Fixture 分类：

- 简单正文。
- 混合 run。
- 多级标题编号。
- 目录与交叉引用。
- 脚注尾注。
- 题注。
- 表格。
- 图片。
- 公式。
- 批注。
- 修订。
- 页眉页脚。
- customXml。
- `mc:AlternateContent`。
- 东亚排版。

验收：

- 每类 fixture 未修改 roundtrip 通过。
- 每类可编辑区域局部修改通过。
- 每类 protected 修改失败。
- diff 报告清晰。

## Phase 14: Agent 使用体验

目标：

让 Codex/Claude Code 能稳定使用软件，不需要猜。

任务：

- 写 Agent 操作指南。
- 写 H-CSS reference。
- 写 HTML 编辑规则。
- 写错误码说明。
- 提供示例：
  - 修改正文。
  - 修改字号。
  - 修改行距。
  - 修改标题样式。
  - 修改表格。
  - 修改脚注。
  - 处理失败报告。
- 提供 `examples/`。

验收：

- Agent 能根据文档完成常见编辑。
- plan/apply/validate 工作流顺畅。
- 错误报告能指导 Agent 自动修复 H-CSS。

## Phase 15: 打包与发布

目标：

交付用户可以安装和使用的软件。

任务：

- 完整 `pyproject.toml`。
- 本地 wheel 构建。
- CLI entrypoint。
- 版本号策略。
- changelog。
- release checklist。
- 可选 Windows 可执行打包。

注意：

- 当前项目默认不改变系统环境。
- 依赖必须项目本地安装或由用户明确安装。
- 打包流程也应尽量在项目目录内完成。

验收：

- 新环境可按文档运行。
- 测试全通过。
- CLI 可用。
- 示例可跑通。

## 6. 质量门

每个阶段合并前必须通过：

```powershell
python -m unittest discover -s tests
```

核心命令必须至少烟测：

```powershell
python -m html_docx export ...
python -m html_docx validate ...
python -m html_docx plan ...
python -m html_docx apply ...
python -m html_docx diff ...
```

涉及 DOCX 输出的阶段必须证明：

- output.docx 可作为 ZIP 打开。
- 必需 parts 存在。
- 未修改 entries hash 不变。
- 修改 entries diff 符合预期。
- protected structures 未变化。

## 7. 完整可用软件的验收清单

软件达到“完整可用”时，必须满足：

### 7.1 功能验收

- `export` 能处理真实未加密 DOCX。
- `validate` 能发现 bundle 损坏。
- `plan` 能预告修改范围。
- `apply` 能输出 DOCX。
- `diff` 能说明变化。
- `inspect` 能定位节点。
- 普通正文文本可编辑。
- 常见字符格式可编辑。
- 常见段落格式可编辑。
- 表格内容可编辑。
- 页眉页脚可编辑。
- 脚注尾注可编辑。
- H-CSS 批量排版可用。
- 学术高级结构保真。
- 不安全修改拒绝。

### 7.2 严格双射验收

- 未修改 roundtrip 字节一致，或在明确记录的规范化策略下等价。
- 修改后非目标 package entries 不变。
- 修改后非目标 XML fragment 不变。
- 原始 media hash 不变，除非用户替换。
- relationships 不无故变化。
- unknown XML 不丢失。
- `mc:AlternateContent` 不被折叠。

### 7.3 Agent 可用性验收

- 文档结构清晰。
- 错误报告机器可读。
- H-CSS 可批量修改。
- 选择器命中数量可审计。
- protected 命中会提示具体节点。
- 示例足够 Agent 模仿。

### 7.4 工程验收

- 自动化测试覆盖核心路径。
- 真实 fixture 覆盖学术写作常见结构。
- 无全局环境依赖。
- 所有生成物在项目目录或用户指定输出目录。
- README 能指导用户完成第一轮使用。

## 8. 开发优先级

最高优先级：

1. 普通文本 patch。
2. run split。
3. 段落格式 patch。
4. H-CSS V1。
5. 表格和 secondary parts。
6. protected academic structures。
7. diff/report 强化。

暂不优先：

- GUI。
- 单文件 HTML。
- 自动重算 TOC。
- 修订链语义编辑。
- SmartArt/Chart/OLE 语义编辑。
- 任意 HTML 到 DOCX。

## 9. 风险管理

| 风险 | 处理策略 |
|---|---|
| XML 序列化引入非目标 diff | 先用 hash 检测，后续实现 fragment patch |
| Agent 误选大范围 | plan 阶段输出命中数量，默认 protected 失败 |
| Word 高级结构复杂 | 默认 protected/opaque 保真 |
| H-CSS 语义膨胀 | 属性白名单和类型系统 |
| 样式继承导致误写 | declared/computed/edit-target 三者分离 |
| 真实 DOCX 不规范 | 尽量保真，不主动修复；不能保真则拒绝 |
| 验证不足 | 每个阶段配套 fixture 和 diff report |

## 10. 下一步立即执行项

下一步应进入 Phase 2：

```text
实现普通 editable run 文本修改回写。
```

具体任务：

1. 给 run manifest 增加 `runIndex`、`paragraphIndex`、`textNodeIndex`、`simpleEditable`。
2. 增强 HTML change detector，提取 run 当前文本。
3. 比较 manifest 中的 textHash。
4. 定位 `word/document.xml` 中对应 run。
5. 修改单一 `w:t`。
6. 重新打包 DOCX。
7. 生成 patch report。
8. 增加测试：
   - 普通文本修改成功。
   - 非目标 entry 不变。
   - protected 修改失败。
   - ID 篡改失败。

完成 Phase 2 后，软件才真正进入“可编辑 DOCX”的阶段。

## 11. 计划维护规则

本计划是开发总线：

- 每完成一个 Phase，要更新当前状态。
- 每发现新的 DOCX 边界，要补进对应设计文档和测试。
- 如果实现与设计冲突，以严格双射和失败优先原则为准。
- 不为了快速支持某个视觉效果而破坏保真。

最终交付不是“能导出一个看起来像的 Word”，而是：

```text
Agent 能放心编辑，用户能验证变化，DOCX 未修改部分能严格保真。
```
