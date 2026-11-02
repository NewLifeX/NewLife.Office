# MCP 与 Agent 接入

> 面向 AI/Agent 的 docx 读写接入说明。核心思路：**先看结构（read）→ 再按锚点改（apply）**，
> 避免让模型凭记忆猜文档结构、也避免"改完静默不生效"。

---

## 1. 三种接入方式

| 方式 | 适用场景 | 入口 |
|------|----------|------|
| **C# 库** | AI 生成 C# 代码直接调用 | `NewLife.Office.Word` 命名空间（`WordReader` / `WordDocumentEditExtensions` / `DocxPatch`） |
| **CLI** | 脚本、批处理、任何能执行命令的 Agent | `NewLife.Office.Cli` 的 `read` / `apply` 命令 |
| **MCP** | 支持 MCP 的 Agent（VS Code、Claude 等） | `NewLife.Office.Cli mcp`（stdio 传输，零外部依赖） |

---

## 2. CLI

```bash
# 结构快照（块级 JSON，含 p#3 / tbl#5 锚点）
NewLife.Office.Cli read report.docx

# 按补丁修改（任一条失败则不写出文件）
NewLife.Office.Cli apply report.docx patch.json out.docx

# 文档转 Markdown（已有能力）
NewLife.Office.Cli convert report.docx

# 图片落地成文件 + 输出页眉页脚（Agent 可直接读图片文件）
NewLife.Office.Cli convert report.docx report.md --images ./report-images --header-footer
```

- `read` 输出到 stdout；**被管道/重定向时按 UTF-8 写字节**，避免控制台代码页把中文转成乱码。
- `apply` 退出码：`0` 全部成功、`1` 参数/文件错误、`3` 有操作失败（此时**不产生输出文件**）。
- `convert` 的 `--images <dir>` 把图片写成 `image1.png`… 并在 Markdown 中以文件名引用；`--header-footer` 把页眉页脚以引用块标注输出（默认均关闭，避免污染正文）。

---

## 3. MCP 服务端

```bash
NewLife.Office.Cli mcp
```

stdio 传输，每行一条 JSON-RPC 2.0 消息；实现了 `initialize` / `tools/list` / `tools/call` / `ping`。

### 工具清单

| 工具 | 入参 | 说明 |
|------|------|------|
| `docx_read_outline` | `path` | 返回结构快照 JSON 文本（块类型、标题层级、表格矩阵、`hasUnmodeledContent` 标记） |
| `docx_apply_patch` | `path`、`output`、`patch` | 执行补丁并写出文件；失败时返回 `isError: true` 且不写文件 |

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

---

## 5. Agent Skill 模板

把下面内容放进 Agent 的 Skill/指令里，可显著减少试错：

```markdown
# docx 编辑技能

处理 Word 文档任务时：
1. 先用 `docx_read_outline`（或 CLI `read`）拿到结构快照，**不要凭猜测定位内容**。
2. 只使用快照里出现过的 anchor（如 `p#3`、`tbl#5`），不要臆造。
3. 段落文本定位优先用 `replace_text` + `find`；结构性改动用 `insert_paragraph` / `delete`。
4. 需要"建议修改"时用修订（C# `SetTextAsRevision` / `StartRevisions`），不要直接改原文。
5. 需要"留批注"时用 `add_comment`，不要往正文里插入说明文字。
6. `hasUnmodeledContent: true` 的块（含域/OLE/批注锚点）尽量避免重建，改整段文字优先用 `replace_text`。
7. 一次任务尽量用一个补丁完成；`apply` 任一条失败会整体不写文件，此时读取失败原因后重试。
```

---

## 6. 已知限制

| 限制 | 说明 |
|------|------|
| 锚点稳定性 | 锚点是"快照内序号"，改动后需重新 `read` 取新锚点 |
| 修订/批注 | 读取修订记录、读取批注内容可用；批注锚点区间读取未建模（程序化新增批注不受影响） |
| 目录 | `AddTableOfContents` 生成 TOC 域，需 Word/WPS 打开时更新域；离线可见用 `AddStaticTableOfContents`（纯文本条目，无跳转） |
| 页脚链接 | 未提供"链接到上一节"显式开关；不写 `footerReference` 即 Word 默认链接行为 |
| Markdown 导出 | 多级无序列表已保留层级；页眉页脚、图片落地尚未输出 |
