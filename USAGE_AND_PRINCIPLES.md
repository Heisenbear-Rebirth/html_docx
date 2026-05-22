# H-DOCX 使用与原理说明

本文是项目的总入口说明，面向开发者、Codex、Claude Code 以及其他需要安全编辑 DOCX 的 agent。

## 项目目标

H-DOCX 的目标不是把 DOCX 尽量转换成普通 HTML，而是在 DOCX 与一种 HTML-like 工作目录之间建立严格可验证的双向关系：

```text
DOCX <-> H-DOCX bundle <-> DOCX
```

核心要求是：

- 未编辑的 DOCX 往返后必须字节级一致。
- 被编辑的 DOCX 必须只改变用户明确要求改变的部分。
- 所有未触碰的 OOXML、媒体、关系、样式、编号、页眉页脚、脚注尾注等必须原样保留。
- 不能证明安全的编辑必须失败并生成报告，而不是猜测式改写。

这使 agent 可以直接修改 `document.html` 或 `agent.edits.hcss`，同时避免破坏 Word 原有排版。

## 基本概念

### DOCX 包

`.docx` 本质上是一个 ZIP 包，内部包含：

- `word/document.xml`：正文主结构。
- `word/styles.xml`：样式定义。
- `word/numbering.xml`：列表与编号。
- `word/header*.xml`、`word/footer*.xml`：页眉页脚。
- `word/footnotes.xml`、`word/endnotes.xml`：脚注尾注。
- `word/media/*`：图片等媒体。
- `_rels` 与 `.rels`：关系文件。
- 其他高级对象、字段、公式、批注、修订、自定义 XML 等。

普通 HTML 转换通常会丢失这些信息。H-DOCX 不丢弃它们，而是把可编辑内容投影出来，把不可安全编辑的内容作为受保护结构保存。

### H-DOCX bundle

`export` 命令会把一个 DOCX 导出为 `.hdocx` 工作目录。典型结构包括：

```text
work.hdocx/
  manifest.json
  document.html
  agent.edits.hcss
  original/original.docx
  parts/
```

其中：

- `document.html` 是 agent 主要阅读和修改的 HTML-like 投影。
- `agent.edits.hcss` 是项目自定义的 H-CSS，用来表达字号、缩进、行距、对齐、列表等受控格式修改。
- `manifest.json` 是内部索引与校验信息，不应手动编辑。
- `original/original.docx` 是原始文件副本，不应手动编辑。
- `parts/` 保存原始 OOXML 与媒体，用于精确回写。

### 受控编辑

H-DOCX 只允许两类安全编辑：

1. 直接修改已投影的可编辑文本。
2. 通过 H-CSS 使用受支持的格式或对象替换操作。

H-CSS 不是浏览器 CSS。所有格式声明都必须使用 `hdocx-` 前缀；
`plan` 会报告选择器命中的节点、每条声明是否支持、规范化值、OOXML 映射、patch id，
以及不支持声明的行号和原因。

例如：

```css
@hdocx-edit mode(paragraph-formatting);

#p-000001 {
  hdocx-text-align: justify;
  hdocx-first-line-indent: 2char;
  hdocx-line-spacing-exact: 18pt;
  hdocx-space-before: 0;
  hdocx-space-after: 0;
}

@hdocx-edit mode(all-runs);

#p-000001 {
  hdocx-font-family: "Times New Roman";
  hdocx-eastAsia-font: "SimSun";
  hdocx-font-size: 10.5pt;
}
```

当前支持的段落声明包括：`hdocx-text-align`/`hdocx-align`、
`hdocx-first-line-indent`、`hdocx-line-spacing`、`hdocx-line-spacing-exact`、
`hdocx-space-before`、`hdocx-space-after`。当前支持的 run 声明包括：
`hdocx-font-family`、`hdocx-eastAsia-font`/`hdocx-east-asia-font`、
`hdocx-ascii-font`、`hdocx-hansi-font`、`hdocx-cs-font`、`hdocx-font-size`、
`hdocx-bold`、`hdocx-italic`、`hdocx-color`。

高级对象如公式、字段、批注、修订、SmartArt、OLE、AlternateContent 等默认受保护。只有存在专门支持的编辑模式时才可以改。

## 安装与运行

开发时建议在项目根目录运行：

```powershell
$env:PYTHONPATH = "src"
python -m html_docx doctor
```

如果使用已构建并安装到项目本地 `.venv` 的命令行工具：

```powershell
.\.venv\Scripts\html-docx.exe doctor
```

本项目不需要网络运行，不需要全局安装依赖。若将来增加依赖，也应安装到项目本地 `.venv`。

## 常用工作流

### 1. 新建 DOCX

如果没有已有 DOCX，先从 canonical blank 模板创建：

```powershell
$env:PYTHONPATH = "src"
python -m html_docx create --out new.docx --title "Draft Title" --paragraph "First paragraph." --export-to new.hdocx --force
python -m html_docx check new.docx --work new-check.hdocx --out new-checked.docx --force --report new-check.json
```

这不是让 Agent 手写 OOXML，而是由工具生成可验证的 DOCX 包。传入
`--export-to` 后可以继续编辑 `new.hdocx/document.html` 和
`new.hdocx/agent.edits.hcss`。

### 2. 未编辑双射检查

用于证明一个 DOCX 可以无损往返：

```powershell
$env:PYTHONPATH = "src"
python -m html_docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
```

成功标准：

- `ok: true`
- `acceptance.byteIdentical: true`
- `acceptance.semanticIdentical: true`
- 输入与输出 SHA256 一致

如果 SHA256 一致，说明输出 DOCX 每一个字节都与输入相同。这比视觉渲染一致更强。

### 3. 受控编辑

推荐流程：

```powershell
$env:PYTHONPATH = "src"
python -m html_docx audit input.docx --report audit.json
python -m html_docx export input.docx --out work.hdocx --force
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
```

Agent 修改前应先 inspect 目标节点、样式、列表、表格或图片，避免使用过宽选择器。

### 4. H-CSS 批量定位

H-CSS 支持命名集合，方便 agent 避免重复书写：

```css
@hdocx-set body-style {
  select: style(BodyText);
}

@hdocx-set first-level-list {
  select: list(1, 0);
}

@hdocx-set header-paragraphs {
  select: part(/word/header1.xml, paragraph);
}
```

如果某个选择器可以合法匹配不到任何节点，必须显式声明：

```css
@hdocx-set optional-notes {
  select: .maybe-note;
  allow-empty: true;
}
```

### 5. 压力测试

修改转换逻辑后运行：

```powershell
$env:PYTHONPATH = "src"
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check pressure-fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

内置压力样例覆盖：

- 最小 DOCX。
- 样式与编号。
- 表格。
- 页眉、脚注、图片。
- 批注、修订、公式。
- 高级受保护对象。

### 6. 可选渲染 QA

如果系统 PATH 中存在 LibreOffice/soffice，可以运行：

```powershell
$env:PYTHONPATH = "src"
python -m html_docx render-check output.docx --out render-out --force --allow-missing --report render.json
```

如果报告为 `renderer-missing`，表示外部渲染器不可用，不代表核心双射失败。

## Agent 使用规则

Agent 应先阅读：

- `AGENTS.md`：仓库内硬规则。
- `AGENT_GUIDE.md`：详细命令与编辑示例。

跨工作区使用时，推荐把 `html-docx-mcp` 加入 `PATH`，再通过 MCP JSON 暴露 tools，
而不是依赖 agent 自动发现 skill。MCP tools 会复用同一套
audit/export/plan/apply/diff/check 逻辑，并要求所有路径位于声明的工作区 root 内。

Agent 可以编辑：

- `.hdocx/document.html` 中的可编辑文本。
- `.hdocx/agent.edits.hcss` 中的受支持 H-CSS 规则。
- H-CSS 支持操作引用的 bundle 内资源。
- `parts/word/media/` 下被明确允许替换的媒体文件。

Agent 不得编辑：

- `manifest.json`
- `original/original.docx`
- 受保护占位结构
- `data-hdocx-id`、`data-hdocx-part`、样式 id、编号 id 等只读元数据
- 没有专门操作支持的 `parts/` 内 OOXML 文件

## 验收标准

完成任何任务前，应根据任务类型运行验收：

- 修改源码：运行 `python -m unittest discover -s tests`。
- 检查某个 DOCX 双射：运行 `check`。
- 做了编辑：运行 `apply` 和 `diff`，确认只有预期变化。
- 修改转换核心：运行压力 fixture 的 `batch-check`。
- 修改排版且有渲染器：运行 `render-check`。

不得只凭“看起来能打开”宣称成功。

## 当前设计边界

严格双射并不意味着把 Word 所有高级功能都变成普通 HTML。项目的策略是：

- 学术写作核心排版要可表达、可定位、可回写。
- 能安全编辑的对象提供受控操作。
- 不能安全编辑的对象作为受保护结构完整保存。
- 任何不确定情况都要报告，而不是静默降级。

因此，H-DOCX 的 HTML 是项目自定义的可逆投影，不是浏览器网页，也不是通用 HTML/CSS 规范。

## 重要文件

- `README.md`：项目快速介绍。
- `AGENTS.md`：agent 硬规则。
- `AGENT_GUIDE.md`：agent 操作手册。
- `src/html_docx/mcp_server.py`：H-DOCX MCP stdio server。
- `FUNCTIONAL_SPEC.md`：功能边界。
- `HDOCX_HTML_DESIGN.md`：H-DOCX/HTML 设计。
- `SELECTOR_AND_REUSE_DESIGN.md`：选择器与复用设计。
- `EDGE_CASE_TEST_MATRIX.md`：边缘/压力测试矩阵。
- `PRESSURE_FIXTURES.md`：压力样例说明。
- `RELEASE_CHECKLIST.md`：发布验收清单。
