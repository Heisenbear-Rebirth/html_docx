# html_docx

[English README](README.md)

面向 Agent 的可逆 DOCX 编辑工具，通过 H-DOCX bundle 让 Codex、Claude Code
等代码 Agent 可以检查、修改并验证 Word 文档，同时尽量避免破坏原有排版和
OOXML 结构。

`html_docx` 并不是普通的 DOCX 转 HTML 工具。它不会把 Word 文档降级成浏览器
HTML，而是把 Agent 需要编辑的内容投影成 HTML-like 文件，并把原始 OOXML 包
作为唯一可信源保存下来。

```text
DOCX <-> H-DOCX bundle <-> DOCX
```

核心契约是严格的：

- 未编辑的往返转换必须字节级一致。
- 编辑后的往返转换必须精确保留所有未触碰的 OOXML 部分。
- 不支持或无法证明安全的编辑必须失败并输出报告，而不是猜测式改写。
- 成功与否以验证报告为准，不能只靠“看起来能打开”判断。

## 为什么需要 H-DOCX

普通 DOCX-to-HTML 转换适合展示，但不适合 Agent 修改学术或专业 Word 文档。
真实 DOCX 中包含样式、编号、脚注、尾注、页眉页脚、字段、公式、批注、修订、
关系文件、媒体、包元数据等结构，这些信息很难无损映射到普通 HTML。

H-DOCX 采用不同策略：

- 把可编辑内容投影到 `document.html`。
- 用 `agent.edits.hcss` 表达受控的格式和结构修改。
- 保存原始 DOCX、原始 OOXML parts、关系文件和元数据。
- 只把被允许的修改片段 patch 回原始包。
- 在声明成功前执行 audit、plan、apply、diff 和验收检查。

这样 Agent 可以读写熟悉的 HTML-like 文件，同时由工具保证 DOCX 安全边界。

## 当前状态

这是一个早期但可用的实现，重点是严格可逆、受控编辑和 Agent 工作流。

当前版本的本地验收结果：

- `101` 个单元测试通过。
- 内置压力 fixture 往返：`6/6` 字节级一致。
- 真实样例往返：字节级一致，语义节点一致。
- MCP stdio smoke test 通过。

## 环境要求

- Windows PowerShell，用于执行随仓库提供的安装脚本。
- Python `>=3.11`。
- 包本身没有 Python 运行时依赖。

项目设计上避免全局安装 Python 包。建议使用工作区 `.venv`，或使用下面的用户级
隔离安装脚本。

## 快速开始

### 克隆仓库

```powershell
git clone https://github.com/Heisenbear-Rebirth/html_docx.git
cd html_docx
```

### 从源码验证

```powershell
$env:PYTHONPATH = "src"
python -m html_docx doctor
python -m unittest discover -s tests
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check pressure-fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

这些命令不需要网络，也不会修改全局 Python 环境。

### 将命令加入 PATH

H-DOCX 不会自动修改你的 MCP 客户端配置。请把包安装到你自己控制的 Python
环境里，然后把该环境的命令目录加入 `PATH`，让任意工作区都能直接运行：

```text
html-docx
html-docx-mcp
```

例如，如果安装在某个 venv 中，Windows 下可以把它的 `Scripts` 目录加入当前终端：

```powershell
$env:PATH = "C:\Tools\hdocx\.venv\Scripts;$env:PATH"
```

安装位置由你决定。若要长期可用，请把同一个命令目录写入你的 shell profile 或
Windows 用户环境变量 `PATH`。

打开新终端后验证：

```powershell
html-docx doctor
'{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' | html-docx-mcp
```

支持通用 `mcpServers` 结构的客户端可以使用这段 MCP JSON：

```json
{
  "mcpServers": {
    "hdocx": {
      "command": "html-docx-mcp",
      "args": []
    }
  }
}
```

如果你的 MCP 客户端外层字段不同，只保留 `hdocx` server object，按客户端要求调整
外层包装即可。Agent 应在调用具体 tool 时传入当前工作区的 `root` 参数。
`HDOCX_MCP_ROOT` 只作为客户端无法传 `root` 时的可选兜底；如果省略，server 会继续
尝试 `CLAUDE_PROJECT_DIR`，最后使用自身当前目录。

修改 MCP 配置后，请重启 MCP 客户端。

## 编辑工作流

### 新建 DOCX

如果任务不是修改已有文档，而是从零开始创建文档，请使用内置 canonical blank
模板：

```powershell
html-docx create --out new.docx --title "Draft Title" --paragraph "First paragraph." --export-to new.hdocx --force
html-docx check new.docx --work new-check.hdocx --out new-checked.docx --force --report new-check.json
```

`create` 会先写出一个有效 DOCX 包。如果传入 `--export-to`，它还会同时导出
H-DOCX bundle，便于 agent 继续通过 `document.html` 和 `agent.edits.hcss` 编辑。

### 未编辑双射检查

遇到新的文档类型时，先证明它可以无损往返：

```powershell
html-docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
```

成功标准：

- `ok: true`
- `acceptance.byteIdentical: true`
- `acceptance.semanticIdentical: true`
- 输入和输出 SHA256 一致

如果 SHA256 一致，说明 DOCX 字节流完全一致。对于未编辑往返，这比渲染一致更强。

### 受控编辑

```powershell
html-docx audit input.docx --report audit.json
html-docx export input.docx --out work.hdocx --force
html-docx inspect work.hdocx --kind node --id p-000001

# 修改 work.hdocx/document.html 或 work.hdocx/agent.edits.hcss。

html-docx plan work.hdocx --report plan.json
html-docx apply work.hdocx --out output.docx --report apply.json
html-docx diff input.docx output.docx --report diff.json
```

如果直接从源码运行，而不是使用已安装的 CLI，请改用：

```powershell
$env:PYTHONPATH = "src"
python -m html_docx ...
```

## H-DOCX Bundle 结构

导出的 bundle 是普通目录：

```text
work.hdocx/
  manifest.json
  document.html
  agent.edits.hcss
  styles.generated.css
  audit.log.jsonl
  original/
    original.docx
    entries.json
  parts/
    ...
```

重要规则：

- 只在投影为可编辑的位置修改 `document.html`。
- 用 `agent.edits.hcss` 表达受支持的格式和结构操作。
- 不要手动编辑 `manifest.json`。
- 不要手动编辑 `original/original.docx`。
- 不要修改受保护占位符。
- 除非有专门操作明确支持，否则不要直接改写 `parts/` 下的 OOXML。

## CLI 命令

| 命令 | 用途 |
| --- | --- |
| `doctor` | 报告运行时能力。 |
| `create` | 新建 canonical DOCX，可选同步导出 H-DOCX。 |
| `audit` | 检测高风险 DOCX 结构和保留策略。 |
| `export` | 把 DOCX 导出为 H-DOCX bundle。 |
| `validate` | 在 apply 前验证 H-DOCX bundle。 |
| `inspect` | 按 id 检查节点、样式、列表、表格或图片。 |
| `query` / `find` | 按文本、样式、字体、段落格式、是否含图片或疑似一级标题返回结构化匹配。 |
| `assert` | 对 bundle 执行断言式验收检查。 |
| `plan` | 只规划编辑，不写 DOCX。 |
| `apply` | 把 bundle 应用回 DOCX。 |
| `diff` | 比较两个 DOCX 的包、语义节点和片段差异。 |
| `roundtrip` | 不编辑，执行 export + apply。 |
| `check` | 对单个 DOCX 执行 export/apply/diff 验收。 |
| `batch-check` | 对文件或目录批量执行 `check`。 |
| `generate-fixtures` | 生成内置压力 DOCX fixture。 |
| `mcp` | 启动 stdio MCP server。 |

包也会安装专用的 `html-docx-mcp` 命令，供 MCP 客户端使用。

## 结构化查询与断言

定位目标时优先使用 `query` / `find`，不要让 agent 直接读
`document.html` 原文。查询结果来自投影 manifest，会返回节点 id、宿主段落、
图片 run、有效格式属性，以及疑似一级标题的启发式理由。

```powershell
html-docx query work.hdocx --text "关键词"
html-docx query work.hdocx --align center --font-size 14pt --font-family SimHei --suspected-heading-level1
html-docx find work.hdocx --kind image
```

编辑后可以用 `assert` 做验收，减少 agent 自己写临时 Python 脚本：

```powershell
html-docx assert work.hdocx --assertion text-payload-unchanged
html-docx assert work.hdocx --assertion images-host-paragraph-not-exact-line-spacing
html-docx assert work.hdocx --assertions-json "[{\"type\":\"paragraphs-have-empty-before\",\"paragraphIds\":[\"p-000006\"]}]"
html-docx assert work.hdocx --assertions-json "[{\"type\":\"paragraphs-have-empty-before\",\"paragraphIds\":[\"p-000006\"],\"afterApply\":true}]"
```

内置断言包括：

- `text-payload-unchanged`
- `paragraphs-have-empty-before`
- `paragraphs-have-empty-after`
- `images-host-paragraph-not-exact-line-spacing`
- `level1-headings-have-empty-paragraph-before`

默认情况下，结构和格式断言检查当前导出的 bundle 投影。若某条断言对象设置
`afterApply: true` 或 `plannedOutput: true`，工具会在 bundle 内部生成临时
planned 输出 DOCX，重新 export 后检查应用后的投影状态。`text-payload-unchanged`
默认检查计划 patch 是否改正文；设置 `afterApply: true` 后会比较原始文档与
planned 输出的文本 payload。

`level1-headings-have-empty-paragraph-before` 使用一级标题启发式。它支持
`includeRegex` 和 `excludeRegex`，并默认排除“摘要 / Abstract / 目录 /
Contents / 关键词”等前置标题。若这些前置标题本来就是目标，可设置
`useDefaultExcludes: false`。

## H-CSS 示例

H-CSS 是本项目自定义的编辑语言，不是浏览器 CSS。所有格式声明都必须使用
`hdocx-` 前缀；普通 CSS 声明会在 `plan` 阶段被报告为不支持。

`hdocx_plan` 是 Agent 写格式声明时的契约边界。它会对每条 H-CSS 规则报告：
规则行号、命中的 H-DOCX 节点 id、每条声明是否支持、规范化后的值、OOXML 映射、
以及该规则会生成的 patch id。

### 段落格式

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-text-align: justify;
  hdocx-line-spacing-exact: 18pt;
  hdocx-first-line-indent: 2char;
  hdocx-space-before: 0;
  hdocx-space-after: 0;
}
```

支持的段落声明：

| 声明 | 取值 | OOXML 映射 |
| --- | --- | --- |
| `hdocx-text-align` / `hdocx-align` | `left`、`center`、`right`、`justify`/`both` | `w:pPr/w:jc @w:val` |
| `hdocx-first-line-indent` | 非负 `char` 或 `pt` | `w:pPr/w:ind @w:firstLineChars` 或 `@w:firstLine` |
| `hdocx-line-spacing` | 正数倍数或精确 `pt` | `w:pPr/w:spacing @w:line` 和 `@w:lineRule` |
| `hdocx-line-spacing-exact` | 正数 `pt` | `w:pPr/w:spacing @w:lineRule="exact"` |
| `hdocx-space-before` | `0`、非负 `pt` 或 `line` | `w:pPr/w:spacing @w:before` 或 `@w:beforeLines` |
| `hdocx-space-after` | `0`、非负 `pt` 或 `line` | `w:pPr/w:spacing @w:after` 或 `@w:afterLines` |
| `hdocx-manual-page-break-before` | `true` 或 `false` | 在目标段落前插入幂等的手动分页段落 |

### 段落结构

当目标输出需要真实的 Word 结构段落，而不是只调已有段落的视觉间距时，
使用 `@hdocx-edit mode(paragraph-structure);`。空行属于这一类：它会被写成
真正的空 `<w:p>`。`plan` 会报告 `insert-empty-paragraph-after` 等结构操作；
`apply` 在相邻位置已经有空段落时会跳过重复插入；`diff` 会把空行变化报告到
`emptyParagraphDiff`，并对齐后续语义节点，避免把未改正文误报成文本变化。

```css
@hdocx-edit mode(paragraph-structure);

#p-000010 {
  hdocx-insert-empty-paragraph-after: true;
  hdocx-empty-paragraph-line-spacing-exact: 12pt;
  hdocx-empty-paragraph-space-before: 0;
  hdocx-empty-paragraph-space-after: 0;
}
```

支持的段落结构声明：

| 声明 | 取值 | OOXML 映射 |
| --- | --- | --- |
| `hdocx-insert-empty-paragraph-before` | `true` 或 `false` | 在目标段落前插入幂等的空 `<w:p>` |
| `hdocx-insert-empty-paragraph-after` | `true` 或 `false` | 在目标段落后插入幂等的空 `<w:p>` |
| `hdocx-empty-paragraph-style-id` | 已存在的简单样式 id | 插入空段落的 `w:pPr/w:pStyle @w:val` |
| `hdocx-empty-paragraph-line-spacing` | 正数倍数或精确 `pt` | 插入空段落的 `w:pPr/w:spacing` |
| `hdocx-empty-paragraph-line-spacing-exact` | 正数 `pt` | 插入空段落的 `w:pPr/w:spacing @w:lineRule="exact"` |
| `hdocx-empty-paragraph-space-before` | `0`、非负 `pt` 或 `line` | 插入空段落的 `w:pPr/w:spacing @w:before` 或 `@w:beforeLines` |
| `hdocx-empty-paragraph-space-after` | `0`、非负 `pt` 或 `line` | 插入空段落的 `w:pPr/w:spacing @w:after` 或 `@w:afterLines` |

如果只是让两段看起来分开，优先用 `hdocx-space-before` /
`hdocx-space-after`。如果 Word 里必须存在一个可拥有独立行距和段前段后距的
真实空行，就使用段落结构模式。

### 函数选择器

```css
@hdocx-set target-paragraph {
  select: id(p-000001);
}

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

选择器支持范围刻意较小：id、class、精确属性选择器、class+属性复合选择器
（例如 `.hdocx-r[data-hdocx-id="r-000001"]`）、逗号分隔的 selector list，
以及上面的 H-DOCX 函数。

安全分组应写在 `agent.edits.hcss` 里；不要为了分组去修改
`document.html` 添加自定义 class 或其他投影元数据。普通规则和
`@hdocx-set` 的 `select` 都支持 selector list：

```css
@hdocx-set body {
  select: id(p-000007), id(p-000008), id(p-000009);
}

.role-body,
.role-reference {
  hdocx-font-size: 10.5pt;
}
```

### Run 格式

```css
@hdocx-edit mode(all-runs);

#r-000001 {
  hdocx-font-family: "Times New Roman";
  hdocx-eastAsia-font: "SimSun";
  hdocx-font-size: 12pt;
  hdocx-bold: true;
}
```

支持的 run 声明：

| 声明 | 取值 | OOXML 映射 |
| --- | --- | --- |
| `hdocx-font-family` | 带引号或不带引号的字体名 | `w:rFonts @w:ascii` 和 `@w:hAnsi` |
| `hdocx-eastAsia-font` / `hdocx-east-asia-font` | 带引号或不带引号的字体名 | `w:rFonts @w:eastAsia` |
| `hdocx-ascii-font` | 带引号或不带引号的字体名 | `w:rFonts @w:ascii` |
| `hdocx-hansi-font` | 带引号或不带引号的字体名 | `w:rFonts @w:hAnsi` |
| `hdocx-cs-font` | 带引号或不带引号的字体名 | `w:rFonts @w:cs` |
| `hdocx-font-size` | 正数 `pt`，例如 `10.5pt` | `w:sz` 半磅值 |
| `hdocx-bold` | `true` 或 `false` | `w:b` |
| `hdocx-italic` | `true` 或 `false` | `w:i` |
| `hdocx-color` | `#RRGGBB` | `w:color @w:val` |

### 图片格式

已有投影图片使用 `@hdocx-edit mode(image-formatting);`。这个模式可以修改
DrawingML 元数据和尺寸，也可以修改图片所在段落的行距/段距，避免 inline 图片
因为继承了正文固定行距而被裁掉或看起来像被隐藏。

```css
@hdocx-edit mode(image-formatting);

#r-000001 {
  hdocx-width-emu: 1828800;
  hdocx-height-emu: 914400;
  hdocx-alt: "Scaled figure";
  hdocx-paragraph-line-spacing: 1;
  hdocx-paragraph-space-before: 0;
  hdocx-paragraph-space-after: 0;
}
```

支持的图片声明：

| 声明 | 取值 | OOXML 映射 |
| --- | --- | --- |
| `hdocx-alt` | 带引号或不带引号的文本 | `wp:docPr @descr` |
| `hdocx-width-emu` | 正整数 EMU | `wp:extent @cx` |
| `hdocx-height-emu` | 正整数 EMU | `wp:extent @cy` |
| `hdocx-paragraph-line-spacing` | 正数倍数或精确 `pt` | 图片所在段落的 `w:pPr/w:spacing` |
| `hdocx-paragraph-line-spacing-exact` | 正数 `pt` | 图片所在段落的精确行距 |
| `hdocx-paragraph-space-before` | `0`、非负 `pt` 或 `line` | 图片所在段落的段前距 |
| `hdocx-paragraph-space-after` | `0`、非负 `pt` 或 `line` | 图片所在段落的段后距 |
| `hdocx-paragraph-text-align` / `hdocx-paragraph-align` | `left`、`center`、`right`、`justify`/`both` | 图片所在段落对齐 |

如果图片在应用论文正文固定行距后被裁切，优先定位图片 run，并设置
`hdocx-paragraph-line-spacing: 1`，让图片所在段落回到不会裁切 inline 图片的自动单倍行距。

### 论文正文格式示例

```css
@hdocx-set body {
  select: style(BodyText);
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-text-align: justify;
  hdocx-first-line-indent: 2char;
  hdocx-line-spacing-exact: 18pt;
  hdocx-space-before: 0;
  hdocx-space-after: 0;
}

@hdocx-edit mode(all-runs);

body {
  hdocx-font-family: "Times New Roman";
  hdocx-eastAsia-font: "SimSun";
  hdocx-font-size: 10.5pt;
}
```

### 手动分页符

需要显式手动分页时，使用 `hdocx-manual-page-break-before: true`，不要用
Word 的自动分页属性 `pageBreakBefore`。`plan` 会报告
`insert-manual-page-break-before`，`apply` 会在目标段落前插入独立的分页段落；
如果同一位置已经存在该分页符，会跳过重复插入。`diff` 会把这类变化报告到
`manualPageBreakDiff`，并对齐后续语义节点，避免仅因插入分页段落而误报正文改变。

```css
@hdocx-edit mode(paragraph-formatting);

.hdocx-p[data-hdocx-id="p-000006"] {
  hdocx-manual-page-break-before: true;
}
```

### 插入图片

```css
@hdocx-insert-image after(#p-000001) {
  source: assets/figure.png;
  alt: "Figure 1";
  width-emu: 914400;
  height-emu: 457200;
}
```

### 插入表格行

```css
@hdocx-insert-table-row after(#tr-000001) {
  cells: "New A|New B";
}
```

### 创建并应用样式

```css
@hdocx-style AgentBody {
  type: paragraph;
  name: "Agent Body";
  based-on: Normal;
  hdocx-font-size: 13pt;
}

@hdocx-edit mode(paragraph-style);

#p-000001 {
  hdocx-style-id: AgentBody;
}
```

### 修改编号定义

```css
@hdocx-set list-items {
  select: [data-hdocx-num-id="1"][data-hdocx-ilvl="0"];
}

@hdocx-edit mode(numbering-definition);

list-items {
  hdocx-num-format: upperLetter;
  hdocx-level-text: "Appendix %1)";
  hdocx-start: 3;
}
```

### 替换公式

```css
@hdocx-edit mode(equation-omml);

[data-hdocx-protected-kind="equation"] {
  hdocx-omml-source: equations/replacement.omml;
}
```

## MCP Agent 集成

仓库把 DOCX 工作流暴露为本地 stdio MCP server：

```text
html-docx-mcp
```

先把 `html-docx-mcp` 加入 `PATH`，然后在 MCP 客户端中加入 JSON 配置：

```json
{
  "mcpServers": {
    "hdocx": {
      "command": "html-docx-mcp",
      "args": []
    }
  }
}
```

MCP server 提供 `hdocx_create`、`hdocx_audit`、`hdocx_export`、
`hdocx_query`、`hdocx_find`、`hdocx_inspect`、`hdocx_plan`、
`hdocx_apply`、`hdocx_diff`、`hdocx_assert`、`hdocx_check`、
`hdocx_batch_check`、`hdocx_guidance` 等 tools。

它还暴露 agent 可直接读取的 resources 和 prompts：

```text
hdocx://guide/workflow
hdocx://guide/writing-format
hdocx://guide/query
hdocx://guide/hcss
hdocx://guide/acceptance
hdocx://guide/edge-cases

hdocx_create_docx
hdocx_safe_edit
hdocx_format_change
hdocx_roundtrip_check
```

Agent 编辑前应读取相关 resource。若某个 MCP 客户端不展示 resources 或 prompts，
也可以调用 `hdocx_guidance`，通过普通 tool response 获取同一套书写规则。

每个面向文件的 tool 都支持可选 `root` 参数。所有文件路径必须解析到该 root 内；
如果不传 `root`，server 会依次使用 `HDOCX_MCP_ROOT`、`CLAUDE_PROJECT_DIR`，
最后才使用 MCP server 当前目录。

Tool 调用应串行执行。如果同一个 server 实例内两个 tool handler 真的重叠，
server 会返回结构化的 `MCP_SERVER_BUSY`，而不是冒险打断 stdio transport。
有些客户端即使让 agent 并行调用，也会把短调用排队后顺序送到 server；这种
情况下两个调用都成功是正常的，因为它们实际没有在 server 内重叠。

`hdocx_doctor` 会报告已加载模块路径、包版本、git commit 或模块时间戳、
支持的 H-CSS edit mode、支持的 H-CSS 属性、stdio 编码和 guidance 来源哈希。
如果 README 与 MCP 实际行为不一致，先运行 `hdocx_doctor`，确认 MCP 进程到底
加载了哪份代码。
项目升级后需要重启 MCP 客户端或 server 进程；长时间运行的 MCP 进程会继续使用
启动时已经加载的旧代码。

Agent 策略：

- 大范围编辑前必须 inspect。
- 新建文档时调用 `hdocx_create`，不要手写 OOXML 包。
- 优先使用 id、命名集合和窄选择器。
- 除非有专门模式支持，否则高级结构一律视为受保护。
- 声称成功前必须运行 `plan`、`apply` 和 `diff`。

## 验证与 QA

按任务风险选择验证：

- 修改源码：`python -m unittest discover -s tests`
- 从零新建 DOCX：`html-docx create` 后运行 `html-docx check`
- 新 DOCX 或未知 DOCX：`html-docx audit` 和 `html-docx check`
- 编辑 DOCX：`html-docx apply` 后运行 `html-docx diff`
- 修改转换核心：`generate-fixtures` 后运行 `batch-check`

## 支持的编辑面

当前支持：

- 使用内置 blank 模板新建 canonical DOCX。
- 可编辑 run 文本。
- Run 格式和 run split。
- 段落格式。
- 段落样式应用、样式创建、安全删除未使用样式。
- 编号/列表元数据投影和受控编号级别编辑。
- 新建单级或多级列表。
- 表格单元格文本、简单非合并表格的行列操作。
- 页眉、脚注、尾注等二级 part 投影。
- 图片 alt text、尺寸元数据、受控媒体替换和插入图片。
- 受控批注文本编辑。
- 修订接受/拒绝操作。
- 从 bundle 内文件执行整段 OMML 公式替换。
- 包级、语义级、片段级 diff。

## 保留策略

以下高风险结构默认保留并保护，除非存在专门操作支持精确修改：

- 自定义 XML
- 图表
- SmartArt
- OLE
- AlternateContent
- 字段
- 公式
- 批注
- 修订
- 文本框
- VML

这是有意设计。严格可逆比假装所有 Word 功能都能变成普通 HTML 更重要。

## 仓库结构

```text
src/html_docx/                  # CLI 和库实现
tests/                          # 单元测试和往返测试
AGENTS.md                       # 仓库级 Agent 规则
CLAUDE.md                       # Claude Code 入口
USAGE_AND_PRINCIPLES.md         # 完整使用与原理说明
```

## 开发

运行测试：

```powershell
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
```

在本地环境构建 wheel：

```powershell
.\.venv\Scripts\python.exe -m pip wheel . --no-deps --no-build-isolation -w dist
```

生成的 fixture、渲染输出、临时 bundle、本地 DOCX 样例和本地虚拟环境均已被
`.gitignore` 排除。

## 文档地图

- `USAGE_AND_PRINCIPLES.md`：使用与实现原理。
- `FUNCTIONAL_SPEC.md`：功能边界。
- `HDOCX_HTML_DESIGN.md`：H-DOCX/HTML 表示设计。
- `SELECTOR_AND_REUSE_DESIGN.md`：选择器与复用模型。
- `SELECTOR_EDGE_CASES_AND_GUARDS.md`：选择器边界与保护。
- `EDITING_EDGE_CASES.md`：编辑边界情况。
- `EDGE_CASE_TEST_MATRIX.md`：边缘和压力测试矩阵。
- `SOFTWARE_ARCHITECTURE.md`：软件架构。
- `GLOBAL_DELIVERY_PLAN.md`：全局交付计划。
- `COMPLETION_PLAN.md`：完成门槛。
- `IMPLEMENTATION_STATUS.md`：实现状态。
- `AGENT_GUIDE.md`：Agent 详细工作流。
- `PRESSURE_FIXTURES.md`：压力 fixture 覆盖说明。
- `RELEASE_CHECKLIST.md`：发布验收清单。
