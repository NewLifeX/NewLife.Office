# Markdown 竞品分析

> 版本：v2.0 | 日期：2026-08-06
> 返回：[竞品分析报告](竞品分析报告.md)

---

## 1. Markdown 竞品概览

| 库名 | 许可证 | 特点 | Stars | NuGet 下载量 |
|------|--------|------|-------|--------------|
| **本项目（NewLife.Office）** | MIT | C# 办公文档级 Markdown 引擎，MD↔docx/PDF/HTML/Word/PPT/Excel 双向全格式转换中枢 | — | — |
| **Markdig** | BSD-2-Clause | 最全面的 CommonMark + GFM 扩展，可扩展管道，.NET 事实标准 | 5.3k | 1 亿+ |
| **CommonMark.NET** | BSD-3-Clause | 严格 CommonMark 规范实现，轻量（已停更，2017） | 1k | — |
| **ReverseMarkdown** | MIT | HTML → Markdown 反向转换器，活跃维护（v5.4） | 395 | — |
| **MarkdownSharp** | MIT | StackOverflow 早期使用的轻量 Markdown 实现（已停更） | 263 | — |
| **Firecrawl anydoc** | MIT | Rust 文档解析引擎，8 大格式家族（Word/PPT/Excel/ODT/RTF/EPUB/CSV/PDF）→ GFM Markdown，中位 4.4ms | 5.8k | —（非 NuGet） |

> **定位差异**：Markdig / CommonMark.NET / MarkdownSharp 是纯 Markdown 解析器；ReverseMarkdown 是 HTML→MD 单向转换器；anydoc 是文档→MD 单向转换器（Rust）；NewLife.Office 是唯一同时覆盖 **MD 解析（CommonMark + GFM + 高级扩展）** 与 **办公文档双向转换** 的 .NET 库。

---

## 2. 功能对比矩阵

> 标记说明：✅ 完整支持 | ⚠️ 部分可用 | ❌ 不支持 | — 不适用（该库不处理此方向）。NewLife.Office 列基于 v1.3.x 当前能力实测标注；anydoc 不解析 Markdown，ReverseMarkdown 不解析 Markdown。

### 2.1 解析能力

| 功能 | NewLife.Office | Markdig | CommonMark.NET | MarkdownSharp | ReverseMarkdown | anydoc |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| CommonMark 核心语法 | ✅ | ✅ | ✅ | ⚠️ | — | — |
| GFM 表格 | ✅ | ✅ | ❌ | ❌ | — | — |
| GFM 任务列表 | ✅ | ✅ | ❌ | ❌ | — | — |
| GFM 删除线 | ✅ | ✅ | ❌ | ❌ | — | — |
| Setext 标题 | ✅ | ✅ | ✅ | ✅ | — | — |
| 引用块嵌套 | ⚠️ | ✅ | ✅ | ⚠️ | — | — |
| 围栏代码块 | ✅ | ✅ | ✅ | ❌ | — | — |
| HTML 块保留 | ✅ | ✅ | ✅ | ✅ | — | — |
| 图片引用提取 | ✅ | ⚠️ | ❌ | ❌ | — | — |
| 自动链接 (AutoLinks) | ✅ | ✅ | ❌ | ❌ | — | — |
| YAML Front Matter | ✅ | ✅ | ❌ | ❌ | — | — |
| 脚注 | ✅ | ✅ | ❌ | ❌ | — | — |
| 数学公式 (LaTeX) | ✅ | ✅ | ❌ | ❌ | — | — |
| Emoji 短码 | ✅ | ✅ | ❌ | ❌ | — | — |
| 自定义属性 `{.class #id}` | ✅ | ✅ | ❌ | ❌ | — | — |
| 定义列表 | ✅ | ✅ | ❌ | ❌ | — | — |
| 缩写（`*[ABBR]:`） | ✅ | ✅ | ❌ | ❌ | — | — |
| 源码位置追踪（行列号） | ✅ | ✅ | ❌ | ❌ | — | — |
| 往返渲染 (Roundtrip) | ✅ | ✅ | ❌ | ❌ | — | — |
| Mermaid 图表 | ❌ | ✅ | ❌ | ❌ | — | — |
| 媒体嵌入 | ❌ | ✅ | ❌ | ❌ | — | — |

### 2.2 输出/转换能力（核心差异化）

| 功能 | NewLife.Office | Markdig | CommonMark.NET | MarkdownSharp | ReverseMarkdown | anydoc |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| MD → HTML | ✅ | ✅ | ✅ | ✅ | — | — |
| MD → 完整 HTML 页面（ToHtmlPage） | ✅ | ❌ | ❌ | ❌ | — | — |
| MD → Markdown（序列化） | ✅ | ✅ | ❌ | ❌ | — | — |
| MD → Word (.docx) | ✅ | ❌ | ❌ | ❌ | — | — |
| MD → PDF | ✅ | ❌ | ❌ | ❌ | — | — |
| **HTML → MD** | ✅ (HtmlToMarkdownConverter) | ❌ | ❌ | ❌ | ✅ | — |
| **Word → MD** | ✅ (IMarkdownExtractable) | ❌ | ❌ | ❌ | ❌ | ✅ |
| **PDF → MD** | ✅ (FormatToMarkdown) | ❌ | ❌ | ❌ | ❌ | ✅ |
| **PPT → MD** | ✅ (FormatToMarkdown) | ❌ | ❌ | ❌ | ❌ | ✅ |
| **Excel → MD** | ✅ (FormatToMarkdown) | ❌ | ❌ | ❌ | ❌ | ✅ |
| ODS/RTF/EPUB → MD | ❌ | ❌ | ❌ | ❌ | ❌ | ✅ |
| 元数据 FrontMatter（标题/源文件/格式） | ✅ | ❌ | ❌ | ❌ | ❌ | ⚠️ |
| HTML 安全选项 (SafeLinks) | ✅ | ❌ | ❌ | ⚠️ | ⚠️ | — |

> anydoc 覆盖面更广（8 大格式家族），但方向单一（仅文档→MD，无反向）；NewLife.Office 独占 **MD→办公文档正向通道**，且全部反向转换均支持可配置选项。

### 2.3 API 与架构

| 功能 | NewLife.Office | Markdig | CommonMark.NET | MarkdownSharp | ReverseMarkdown | anydoc |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 管线/管道模式 | ✅ (MarkdownPipeline) | ✅ | ❌ | ❌ | ❌ | — |
| 强类型 AST 继承体系 | ✅ (12 种块类型) | ✅ | ❌ | ❌ | ❌ | ✅ (统一 Document) |
| 扩展注册机制 | ✅ (IMarkdownExtension) | ✅ (20+ 扩展) | ❌ | ❌ | ⚠️ (标签别名) | — |
| 统一 AST 转换层（多 Reader→块模型，单一序列化器） | ✅ | — | — | — | — | ✅ |
| 源码位置追踪 | ✅ | ✅ | ❌ | ❌ | ❌ | — |
| 往返渲染 (Roundtrip) | ✅ | ✅ | ❌ | ❌ | ❌ | — |
| 文本规范化（控制字符/NBSP/软连字符） | ✅ (MarkdownTextCleaner) | ❌ | ❌ | ❌ | ❌ | ✅ |
| 类型化转换错误 | ✅ (ConvertErrorException) | ❌ | ❌ | ❌ | ❌ | ✅ (ConvertError) |
| 流式解析（Stream） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 文件解析 | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| 目标框架覆盖 | net45/netstandard2.0/2.1 | netstandard2.0+（无 net45） | net45+ | net45+ | net8.0+ | Rust 绑定 |
| 外部依赖数 | 1 (NewLife.Core) | 0 | 0 | 0 | 1 (HtmlAgilityPack) | — |
| NuGet 包体积 | <50KB | ~500KB | ~100KB | ~100KB | ~200KB | —（非 NuGet） |

**API 易用性**：解析并输出 HTML 三者均一行调用，NewLife.Office 额外返回 `MarkdownDocument` 对象可继续转 Word/PDF，灵活性更高：

```csharp
// NewLife.Office
var doc = MarkdownDocument.Parse("# Hello\nWorld");
var html = doc.ToHtml();

// Markdig
var html = Markdown.ToHtml("# Hello\nWorld");

// CommonMark.NET
var html = CommonMarkConverter.Convert("# Hello\nWorld");
```

---

## 3. 差距分析

### 3.1 NewLife.Office Markdown 核心优势

| 优势 | 说明 |
|------|------|
| **办公文档级双向转换** | 独有 MD→Word(.docx)/MD→PDF 正向转换 + Word/HTML/PDF/PPT/Excel→MD 反向全路径，任何纯 Markdown 库或 anydoc（单向）均无法做到 |
| **统一 AST 转换层** | `IMarkdownDocumentExtractable.ToMarkdownDocument()`，各 Reader 直接产出块模型，单一序列化器保证输出一致性（对标 anydoc 统一 Document 模型） |
| **全格式转换中枢** | `FormatToMarkdown.Convert/ToDocument` 统一入口，PDF/PPT/Excel/Word→MD 一行转换，自动附加元数据 FrontMatter |
| **零依赖轻量** | 仅依赖 NewLife.Core，包体积极小（<50KB） |
| **全框架覆盖** | net45 ~ netstandard2.1 全覆盖，Markdig 已放弃 net45，ReverseMarkdown 仅 net8.0+ |
| **强类型 AST** | 12 种块类型继承体系（HeadingBlock/CodeBlock/ListItemBlock/TableBlock 等），对标 Markdig |
| **完整 HTML 页面** | `ToHtmlPage()` 一键输出含内联样式的完整页面 |
| **安全选项** | SafeLinks 危险链接过滤，适合服务端渲染场景 |
| **统一办公文档生态** | 与 Excel/Word/PDF/PPT 模块同一库，API 风格一致，零对接成本 |

### 3.2 差距项（已全部补齐 ✅）

| 优先级 | 功能 | 竞品参考 | 状态 |
|--------|------|---------|:---:|
| P0 | HTML→MD 反向转换 | ReverseMarkdown | ✅ MD04-01 |
| P0 | 自动链接 (AutoLinks) | Markdig | ✅ MD05-01 |
| P1 | YAML Front Matter | Markdig | ✅ MD05-02 |
| P1 | 管线/扩展机制 | Markdig | ✅ MD06-01/02 |
| P1 | PDF→MD 文本提取 | — | ✅ MD07-01 |
| P1 | 统一 AST 转换层 | anydoc | ✅ MD07-04 |
| P2 | 脚注 | Markdig | ✅ MD05-03 |
| P2 | 数学公式 (LaTeX) | Markdig | ✅ MD05-04 |
| P2 | Emoji 短码 | Markdig | ✅ MD05-05 |
| P2 | 自定义属性/定义列表/缩写 | Markdig | ✅ MD05-06/07/08 |
| P3 | 往返渲染 (Roundtrip) | Markdig | ✅ MD06-03 |
| P3 | 源码位置追踪 | Markdig | ✅ MD06-04 |
| P3 | 元数据 FrontMatter/文本规范化/类型化错误 | anydoc | ✅ MD07-05/06/07 |

> 对标 Markdig 与 anydoc 的全部差距项已补齐完成。**有意不实现的项**：Mermaid 图表渲染（需引入 JS 引擎，偏离零依赖原则）、媒体嵌入（非 Markdown 核心需求）。

### 3.3 双向转换矩阵（NewLife.Office 特有）

| 源格式 ╲ 目标格式 | MD | HTML | Word | PDF |
|------|:---:|:---:|:---:|:---:|
| **MD** | ✅ 序列化 | ✅ ToHtml/ToHtmlPage | ✅ ToWord/SaveWord | ✅ ToPdf/SavePdf |
| **HTML** | ✅ HtmlToMarkdownConverter | — | — | — |
| **Word** | ✅ IMarkdownExtractable | ✅ WordHtmlConverter | — | ✅ WordPdfConverter |
| **PDF** | ✅ PdfReader→MD | — | — | — |
| **PPT** | ✅ PptxReader→MD | — | — | — |
| **Excel** | ✅ ExcelReader→MD | — | — | — |

> 该矩阵是 NewLife.Office 独有的双向全路径：纯 Markdown 解析库（Markdig 等）只有左上角一行；anydoc 只有第一列（文档→MD）；NewLife.Office 已闭环 6 格式 × 双向。

---

## 4. 架构演进路线

对标 Markdig 的成熟架构 + anydoc 的统一 Document 模型 + NewLife.Office 独有的办公文档转换能力，分四步演进，现已全部完成：

### 4.1 v2.0 强类型 AST

- **目标**：消除属性混杂，建立类型安全的块模型
- **交付**：12 种块类型继承体系（HeadingBlock/CodeBlock/ListItemBlock/TableBlock 等），每种块持有独立强类型属性
- **意义**：为管线、扩展、源码位置追踪、往返渲染提供统一 AST 基础，对标 Markdig 的 MarkdownObject 体系

### 4.2 v3.0 MarkdownPipeline 管线 + 扩展注册

- **目标**：可配置、可插拔的解析→转换→渲染管线
- **交付**：`MarkdownPipeline`（扩展粒度 Enable 开关）+ `IMarkdownExtension` 扩展注册机制（`pipeline.Use(ext)`），补齐 AutoLinks/YAML Front Matter/脚注/数学公式/Emoji/自定义属性/定义列表/缩写
- **意义**：功能可裁剪（CreateStrict/CreateGfm 预设管线），第三方可扩展，对标 Markdig 的扩展管道

### 4.3 v4.0 全格式转换中枢（对标 anydoc）

- **目标**：办公文档→Markdown 统一入口，输出一致性
- **交付**：
  - `FormatToMarkdown.Convert/ToDocument`：PDF/PPT/Excel/Word→MD 一行转换
  - `IMarkdownDocumentExtractable.ToMarkdownDocument()`：各 Reader 直接产出 AST 块模型，单一序列化器保证输出一致性（对标 anydoc 统一 Document 模型 + 单一 GFM 序列化器）
  - 元数据 FrontMatter（标题/源文件/格式）、文本规范化（MarkdownTextCleaner）、类型化错误（ConvertErrorException）
- **意义**：完成 .NET 生态唯一的办公文档 ↔ Markdown 双向全格式转换闭环

| 阶段 | 目标 | 关键交付 | 状态 |
|------|------|---------|:---:|
| **v1.0** | 基础解析 + 正向转换 | CommonMark+GFM 解析，MD→HTML/Word/PDF | ✅ 完成 |
| **v2.0** | AST 继承体系 | 12 种块类型子类，消除属性混杂 | ✅ 完成 |
| **v2.1** | 反向转换闭环 | HTML→MD（对标 ReverseMarkdown） | ✅ 完成 |
| **v3.0** | 管线 + 扩展 + 补齐解析 | MarkdownPipeline + AutoLinks/YAML/脚注/数学/Emoji 等扩展 | ✅ 完成 |
| **v4.0** | 全格式转换中枢 | PDF/PPT/Excel/Word→MD + 统一 AST 转换层（对标 anydoc） | ✅ 完成 |

---

## 5. 结论与差异化定位

**结论**：NewLife.Office 在 **办公文档级 Markdown 转换**（MD→Word/PDF 正向、HTML/Word/PDF/PPT/Excel→MD 反向全路径）方面具有独特且不可替代的优势——这是任何纯 Markdown 解析库（Markdig/CommonMark.NET/MarkdownSharp）、单向转换器（ReverseMarkdown）或 Rust 文档解析引擎（anydoc）都无法做到的。v2.0 强类型 AST、v2.1 反向转换闭环、v3.0 管线/扩展、v4.0 全格式转换中枢**四大里程碑全部完成**，已实现 **.NET 生态中最完整的办公文档 ↔ Markdown 双向全格式转换矩阵**。

**战略定位**：NewLife.Office 不只是一个 Markdown 解析器，而是 **.NET 生态中唯一的办公文档 ↔ Markdown 双向转换中枢**，能力覆盖 Markdig（解析）+ ReverseMarkdown（HTML→MD）+ anydoc（文档→MD）三者的并集，并补上了三者都没有的 **MD→办公文档正向通道**。

> **2026-08 读写完善（MD08）**：完成 CommonMark 核心语法补全——引用链接三形态、缩进代码块、HTML 实体解码、行内 HTML 与 6 种 HTML 块类型、列表续行与嵌套；写入侧上下文相关转义与往返修改检测（快照比较）；`MarkdownPipeline` 开关真实生效；文件编码自动检测（BOM/严格 UTF-8/GBK）。日常读写解析能力已对标 Markdig 核心语法。

**真实未实现项（不虚标）**：
- **内容识别格式（magic bytes）**：anydoc 通过文件内容而非扩展名识别格式；NewLife.Office 目前按扩展名/显式入口分发
- **Agent Skill / CLI**：anydoc 提供 Node.js/Python 绑定与 CLI；NewLife.Office 目前为纯 .NET 库 API
- **Mermaid 图表 / 媒体嵌入**：Markdig 独有，本项目有意不实现（需引入 JS 引擎，偏离零依赖原则）
- **ODS/RTF/EPUB → MD**：anydoc 覆盖的格式家族，本项目未纳入转换中枢
