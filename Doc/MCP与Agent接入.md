# MCP 与 Agent 接入

> 面向 AI/Agent 的 docx/xlsx/pptx/pdf 读写接入说明。核心思路：**先看结构（read）→ 再按锚点改（apply）**，
> 避免让模型凭记忆猜文档结构、也避免"改完静默不生效"。

---

## 1. 三种接入方式

| 方式 | 适用场景 | 入口 |
|------|----------|------|
| **C# 库** | AI 生成 C# 代码直接调用 | `NewLife.Office.OfficeDocumentTool` 统一入口，或各格式 API |
| **CLI** | 脚本、批处理、任何能执行命令的 Agent | `NewLife.Office.Cli` 的 `read` / `apply` / `text` / `range` / `search` |
| **MCP** | 支持 MCP 的 Agent（VS Code、Claude 等） | `NewLife.Office.Cli mcp`（stdio 传输，零外部依赖） |

---

## 2. CLI

```bash
# 结构快照（docx/xlsx/pptx/pdf 自动识别）
NewLife.Office.Cli read report.xlsx

# 按补丁修改（任一条失败则不写出文件）
NewLife.Office.Cli apply report.xlsx patch.json out.xlsx

# 分段读文本（避免一次灌满上下文）
NewLife.Office.Cli text big.docx --max-chars 4000
NewLife.Office.Cli text big.docx --max-chars 4000 --cursor 4000

# 工作簿按行分页读取
NewLife.Office.Cli range book.xlsx "Sheet1!A1:F2000" --max-rows 200

# 搜索（工作簿单元格 / PDF 文本）
NewLife.Office.Cli search book.xlsx "华东"
NewLife.Office.Cli search report.pdf "合同金额"

# 文档转 Markdown
NewLife.Office.Cli convert report.docx

# 图片落地成文件 + 输出页眉页脚（Agent 可直接读图片文件）
NewLife.Office.Cli convert report.docx report.md --images ./report-images --header-footer
```

- `read` 输出到 stdout；**被管道/重定向时按 UTF-8 写字节**，避免控制台代码页把中文转成乱码。
- `apply` 退出码：`0` 全部成功、`1` 参数/文件错误、`2` 格式/加密错误、`3` 有操作失败（此时**不产生输出文件**）。
- `text` / `range` 在 stderr 输出续读游标提示（`nextCursor=N`）。
- `convert` 的 `--images <dir>` 把图片写成 `{文档名}-image1.png`…（带源文档名前缀，多个文档落同一目录不会互相覆盖）并在 Markdown 中以文件名引用；`--header-footer` 把页眉页脚以引用块标注输出（默认均关闭，避免污染正文）。

---

## 3. MCP 服务端

```bash
NewLife.Office.Cli mcp
```

stdio 传输，每行一条 JSON-RPC 2.0 消息；实现了 `initialize` / `tools/list` / `tools/call` / `ping`。

### 工具清单

| 工具 | 入参 | 说明 |
|------|------|------|
| `office_read_outline` | `path` | 任意格式结构快照（按扩展名/内容路由） |
| `office_read_text` | `path`、`maxChars`、`cursor` | 分段读纯文本，返回 `nextCursor` |
| `office_search` | `path`、`text`、`regex`、`maxResults` | 工作簿或 PDF 内搜索 |
| `office_apply_patch` | `path`、`output`、`patch` | 任意格式补丁（失败不写文件） |
| `docx_read_outline` | `path` | docx 结构快照（块类型、标题层级、表格矩阵、`hasUnmodeledContent`） |
| `docx_apply_patch` | `path`、`output`、`patch` | docx 补丁 |
| `xlsx_read_outline` | `path` | 工作表/表头/预览/列摘要/表格对象/图表统计 |
| `xlsx_read_range` | `path`、`range`、`maxRows`、`cursor` | A1 区域分页读取 |
| `xlsx_apply_patch` | `path`、`output`、`patch` | 工作簿补丁 |
| `pptx_read_outline` | `path` | 幻灯片/文本框/形状/图片/表格/图表/备注 |
| `pptx_apply_patch` | `path`、`output`、`patch` | 演示文稿补丁 |
| `pdf_read_outline` | `path` | 页/书签/表单/注释/附件/字体/加密签名状态 |
| `pdf_search` | `path`、`text`、`regex`、`maxResults` | 返回页码与矩形坐标 |
| `pdf_apply_patch` | `path`、`output`、`patch` | PDF 补丁 |

### Agent 侧配置示例（VS Code `mcp.json`）

```json
{
  "servers": {
    "newlife-office": {
      "type": "stdio",
      "command": "dotnet",
      "args": ["run", "--project", "C:/X/NewLife.Office/NewLife.Office.Cli", "--", "mcp"]
    }
  }
}
```

---

## 4. 补丁指令集

顶层结构：`{ "ops": [ ... ] }`，按数组顺序执行。

### 4.1 docx（锚点 `p#3` / `tbl#5` / `img#7`）

| op | 必填字段 | 行为 |
|----|----------|------|
| `replace_text` | `find`、`replace`，可选 `regex` | 全文替换（跨 Run）；`regex: true` 时按正则，支持 `$1` 分组 |
| `set_text` | `anchor`、`text` | 整段替换，保留首 Run 格式 |
| `insert_paragraph` | `anchor`、`text`，可选 `position`(after/before)、`style` | 在锚点段落前后插入；`anchor` 可省略表示追加到末尾 |
| `delete` | `anchor` | 删除该顶层元素（段落/表格/图片） |
| `set_style` | `anchor`、`style` | 设置段落样式（`Heading1`~`Heading6` 或自定义样式 ID） |
| `set_cell_text` | `anchor`（表格）、`row`、`column`、`text` | 写单元格 |
| `insert_table` | `anchor`，`cells` 或 `rows`/`columns` | 插入表格（`cells` 为二维数组时首行作表头） |
| `add_comment` | `anchor`、`text`，可选 `author` | 给整段加批注 |

示例：

```json
{
  "ops": [
    { "op": "replace_text", "find": "A4", "replace": "A5" },
    { "op": "set_text", "anchor": "p#2", "text": "结论：同意上线。" },
    { "op": "add_comment", "anchor": "p#2", "text": "此处需补充数据来源", "author": "AI 审阅" },
    { "op": "set_cell_text", "anchor": "tbl#5", "row": 1, "column": 2, "text": "市场部" }
  ]
}
```

### 4.2 xlsx（锚点 `Sheet1!A1:D20`）

| op | 必填字段 | 行为 |
|----|----------|------|
| `set_cells` | `range`、`values`（二维或标量） | 批量写值；字符串以 `=` 开头写公式 |
| `set_cell` | `range`、`value` | 单格写值 |
| `set_formula` | `range`、`formula` | 写公式（多格同公式会告警：相对引用不自动调整） |
| `append_rows` | `sheet`、`values` | 追加到最后一个有内容的行之后 |
| `insert_rows` / `delete_rows` | `range`（行范围，如 `Sheet1!3:4`） | 增删行并平移样式/公式/批注/合并区 |
| `merge` / `unmerge` | `range` | 合并/取消合并 |
| `set_style` | `range` + `bold`/`italic`/`fontSize`/`fontName`/`fontColor`/`backgroundColor`/`hAlign`/`vAlign`/`wrapText`/`border`/`numberFormat` | 区域样式（未传的属性保持原值） |
| `set_number_format` | `range`、`format` | 数字格式 |
| `set_column_width` | `sheet`、`columns`、`width` | 列宽 |
| `set_row_height` | `sheet`、`rows`、`height` | 行高 |
| `freeze` / `autofilter` | `sheet` + `rows`/`cols`；`range` | 冻结/筛选 |
| `add_sheet` / `rename_sheet` / `delete_sheet` | `name`；`from`/`to` | 工作表管理 |
| `add_table` | `range`、`name`，可选 `style`/`columns` | 结构化表格 |
| `add_chart` | `range`（首行系列名、首列分类）、`type`、`title` | 嵌入图表 |

示例：

```json
{
  "ops": [
    { "op": "set_cells", "range": "Sheet1!A1:B2", "values": [["姓名", "部门"], ["张三", "研发"]] },
    { "op": "set_formula", "range": "Sheet1!C2", "formula": "=SUM(B2:B2)" },
    { "op": "set_style", "range": "Sheet1!A1:B1", "bold": true, "backgroundColor": "EEEEEE" }
  ]
}
```

### 4.3 pptx（锚点 `s#2`、`s#2/tb#0`、`s#2/shape#0`、`s#2/pic#0`、`s#2/tbl#0`、`s#2/chart#0`）

| op | 必填字段 | 行为 |
|----|----------|------|
| `set_shape_text` | `anchor`、`text` | 改写文本框/形状文本（保留文本框级格式） |
| `set_table_cell` | `anchor`、`row`、`column`、`text` | 写表格单元格 |
| `update_chart` | `anchor`，可选 `categories`/`series`/`title` | 更新图表数据 |
| `replace_image` | `anchor`、`base64` 或 `path` | 替换图片 |
| `set_notes` | `anchor`（幻灯片）、`text` | 演讲者备注 |
| `set_title` | `anchor`、`text` | 按标题占位符定位改写 |
| `add_slide` | 可选 `at`/`layoutIndex`/`title`/`text`/`notes` | 新增幻灯片 |
| `delete_slide` / `move_slide` / `clone_slide` | `anchor`；`from`/`to` | 页编排（克隆为浅拷贝，附告警） |

### 4.4 pdf（页码 1 起始，坐标原点左下角）

| op | 必填字段 | 行为 |
|----|----------|------|
| `delete_pages` / `extract_pages` | `pages`（`"1-3,5"` 或 `[1,2]`） | 删除/只保留指定页 |
| `move_page` / `rotate_page` | `from`/`to`；`pages`、`rotation` | 页序与旋转 |
| `set_metadata` | `title`/`author`/`subject`/`keywords`/`creator`/`producer` | 文档信息 |
| `add_text` | `page`、`text`、`x`、`y`、`fontSize`、`color`、`rotation` | 叠加文字（Helvetica，Latin-1） |
| `add_watermark` | `text`、`fontSize`、`color` | 全页 45° 水印 |
| `add_highlight` / `add_note` | `page`、`x`、`y`、`width`/`height`、`text` | 高亮/便签（配合 `pdf_search` 的坐标） |
| `fill_form` | `name`、`value` 或 `fields` | 表单填充（复选框传 `true`/`false`） |
| `merge` | `path` | 追加另一份 PDF 的页面 |

示例：

```json
{
  "ops": [
    { "op": "set_metadata", "title": "2026 年度报告" },
    { "op": "add_highlight", "page": 1, "x": 72, "y": 700, "width": 120, "height": 14, "text": "重点" },
    { "op": "rotate_page", "pages": "2", "rotation": 90 }
  ]
}
```

### 4.5 两条通用契约

1. **锚点是 read 时刻的快照位置**：结构性操作（insert/delete/move/add_slide 等）会让后续位置偏移，请把结构性操作排在补丁最后，或按锚点从大到小排列；改动后重新 `read`。
2. **失败不落盘**：某条失败时已成功的操作仍留在内存模型上，CLI/MCP 因此**不写出文件**；如需重试，必须重新读取源文件，不要复用已部分修改的对象。

---

## 5. Agent Skill 模板

把下面内容放进 Agent 的 Skill/指令里，可显著减少试错：

```markdown
# 办公文档读写技能

处理 Word/Excel/PPT/PDF 任务时：
1. 先用 `office_read_outline`（或 CLI `read`）拿结构快照，**不要凭猜测定位内容**。
2. 只使用快照里出现过的锚点（`p#3`、`Sheet1!A1`、`s#2/tb#0`、页码+矩形），不要臆造。
3. 长内容用 `office_read_text` / `xlsx_read_range` 分段读，按 `nextCursor` 续读，避免一次灌满上下文。
4. 文本定位优先用 `replace_text` / `office_search`；结构性改动（insert/delete/add_slide）排在补丁最后。
5. PDF 标注先用 `pdf_search` 拿矩形，再用 `add_highlight`/`add_note` 的 `x`/`y`/`width`/`height`。
6. 需要"建议修改"时用修订（C# `SetTextAsRevision` / `StartRevisions`）；需要"留批注"时用 `add_comment`，不要往正文插说明文字。
7. 一次任务尽量用一个补丁完成；失败时读取错误码（Unsupported/Malformed/Encrypted/AnchorMissing/ResourceLimit）后重新 `read` 再试。
8. 加密文档会被拒绝（Encrypted），不要反复重试。
```

---

## 6. 已知限制

| 限制 | 说明 |
|------|------|
| 锚点稳定性 | 锚点是"快照内序号"，改动后需重新 `read` 取新锚点 |
| docx 修订/批注 | 读取修订记录、批注内容可用；批注锚点区间读取未建模（程序化新增批注不受影响） |
| docx 目录 | `AddTableOfContents` 生成 TOC 域，需 Word/WPS 打开时更新域；离线可见用 `AddStaticTableOfContents`（纯文本条目，无跳转） |
| xlsx 图表/透视表 | 未建模部件原样透传；`add_chart` 仅支持按数据区域生成基础图表 |
| pptx 保真 | 按模型重建（L2），SmartArt/复杂动画等未建模元素不透传；另存时以源文件为模板复用母版/主题 |
| pdf 加密 | 加密 PDF 拒绝对象级修改，需先解密 |
| pdf 叠加文字 | 仅 Helvetica + Latin-1；中文请先生成图片或改用模板/生成式 PDF |

---

## 7. 实现状态

| 能力 | 状态 |
|------|------|
| docx 快照 + 补丁 + CLI/MCP | ✅ W29 |
| xlsx 快照 + A1 读写 + 补丁 + CLI/MCP | ✅ M34 |
| pptx 快照 + 补丁 + CLI/MCP | ✅ S21 |
| pdf 快照 + 搜索定位 + 对象级补丁 + CLI/MCP | ✅ P10 |
| 跨格式统一入口与分页/错误契约 | ✅ GEN-4 |
