# Markdown 竞品分析

> 版本：v2.0 | 日期：2026-07-03

← 返回 [竞品分析报告.md](竞品分析报告.md)

---

## 1. Markdown 竞品概览

| 项目 | 类型 | 语言 | 许可证 | 最新版本 | Stars | 定位 |
|------|------|------|--------|---------|-------|------|
| **NewLife.Office** | 开源 | C# | MIT | — | — | **唯一的办公文档级 Markdown 引擎**，MD↔docx/PDF/HTML 双向全格式转换中枢 |
| **Markdig** | 开源 | C# | BSD-2-Clause | v1.3.2 (2026-06) | 5.3k | 最快、功能最全的 .NET Markdown 处理器 |
| **CommonMark.NET** | 开源(停更) | C# | BSD-3-Clause | v0.15.1 (2017) | 1k | 严格 CommonMark 规范实现 |
| **ReverseMarkdown** | 开源 | C# | MIT | v5.4.0 (2026-06) | 389 | HTML → Markdown 反向转换器 |

---

## 2. 功能对比矩阵

### 2.1 解析能力

| 功能 | NewLife.Office | Markdig | CommonMark.NET | ReverseMarkdown |
|------|:---:|:---:|:---:|:---:|
| CommonMark 核心语法 | ✅ | ✅ | ✅ | — |
| GFM 表格 | ✅ | ✅ | ❌ | ✅ (HTML→MD) |
| GFM 任务列表 | ✅ | ✅ | ❌ | ✅ |
| GFM 删除线 | ✅ | ✅ | ❌ | ✅ |
| Setext 标题 | ✅ | ✅ | ✅ | — |
| 引用块嵌套 | ⚠️ | ✅ | ✅ | ✅ |
| 围栏代码块 | ✅ | ✅ | ✅ | ✅ |
| HTML 块保留 | ✅ | ✅ | ✅ | ✅ |
| YAML Front Matter | ✅ | ✅ | ❌ | ❌ |
| 脚注 | ✅ | ✅ | ❌ | ❌ |
| 数学公式 (LaTeX) | ✅ | ✅ | ❌ | ❌ |
| 图表 (Mermaid) | ❌ | ✅ | ❌ | ❌ |
| Emoji 短码 | ✅ | ✅ | ❌ | ❌ |
| 自定义属性 `{.class}` | ✅ | ✅ | ❌ | ❌ |
| 定义列表 | ✅ | ✅ | ❌ | ❌ |
| 缩写 | ✅ | ✅ | ❌ | ❌ |
| 自动链接 | ✅ | ✅ | ❌ | ❌ |
| 媒体嵌入 | ❌ | ✅ | ❌ | ❌ |
| 源码位置追踪 | ✅ | ✅ | ❌ | ❌ |
| 往返渲染 (Roundtrip) | ✅ | ✅ | ❌ | ❌ |

### 2.2 输出/转换能力（核心差异化）

| 功能 | NewLife.Office | Markdig | CommonMark.NET | ReverseMarkdown |
|------|:---:|:---:|:---:|:---:|
| MD → HTML | ✅ | ✅ | ✅ | — |
| MD → 完整 HTML 页面 | ✅ | ❌ | ❌ | ❌ |
| MD → Word (.docx) | ✅ | ❌ | ❌ | ❌ |
| MD → PDF | ✅ | ❌ | ❌ | ❌ |
| MD → Markdown (序列化) | ✅ | ✅ | ❌ | — |
| **HTML → MD** | ✅ (HtmlToMarkdownConverter) | ❌ | ❌ | ✅ |
| **Word → MD** | ✅ (IMarkdownExtractable) | ❌ | ❌ | ❌ |
| **PDF → MD** | ✅ (MD07-01, PdfReader→MD) | ❌ | ❌ | ❌ |
| HTML 安全选项 (SafeLinks) | ✅ | ❌ | ❌ | ⚠️ |

### 2.3 API 与架构

| 功能 | NewLife.Office | Markdig | CommonMark.NET | ReverseMarkdown |
|------|:---:|:---:|:---:|:---:|
| 管线/管道模式 | ✅ (MarkdownPipeline) | ✅ | ❌ | ❌ |
| **强类型 AST 继承体系** | ✅ (v2.0 已实现) | ✅ | ❌ | ❌ |
| 扩展机制 | ✅ (IMarkdownExtension) | ✅ (20+扩展) | ❌ | ⚠️ (标签别名) |
| 流式解析 | ✅ | ✅ | ✅ | ✅ |
| 文件解析 | ✅ | ✅ | ❌ | ❌ |
| NuGet 包体积 | <50KB | ~500KB | ~100KB | ~200KB |
| 外部依赖数 | 1 (NewLife.Core) | 0 | 0 | 1 (HtmlAgilityPack) |

---

## 3. 差距分析

### 3.1 NewLife.Office Markdown 核心优势

| 优势 | 说明 |
|------|------|
| **办公文档级双向转换** | 独有 MD→Word(.docx) 和 MD→PDF 转换，且已有 Word→MD 反向通道（IMarkdownExtractable），竞品无人能及 |
| **零依赖轻量** | 仅依赖 NewLife.Core，包体积极小 (<50KB) |
| **全框架覆盖** | net45 ~ net10 全覆盖，Markdig 已放弃 net45，ReverseMarkdown 仅 net8.0+ |
| **强类型 AST** | v2.0 已实现 12 种块类型继承体系（HeadingBlock/CodeBlock/ListItemBlock 等），对标 Markdig |
| **完整 HTML 页面** | `ToHtmlPage()` 一键输出含内联样式的完整页面 |
| **安全选项** | SafeLinks 危险链接过滤，适合服务端渲染场景 |
| **统一办公文档生态** | 与 Excel/Word/PDF/PPT 模块同一库，API 风格一致，零对接成本 |

### 3.2 差距项（已全部补齐 ✅）

| 优先级 | 功能 | 竞品参考 | 状态 |
|--------|------|---------|:---:|
| ~~P0~~ | ~~HTML→MD 反向转换~~ | ReverseMarkdown | ✅ MD04-01 |
| ~~P0~~ | ~~自动链接 (AutoLinks)~~ | Markdig | ✅ MD05-01 |
| ~~P1~~ | ~~YAML Front Matter~~ | Markdig | ✅ MD05-02 |
| ~~P1~~ | ~~管线/扩展机制~~ | Markdig | ✅ MD06-01/02 |
| ~~P2~~ | ~~脚注~~ | Markdig | ✅ MD05-03 |
| ~~P2~~ | ~~数学公式 (LaTeX)~~ | Markdig | ✅ MD05-04 |
| ~~P2~~ | ~~Emoji 短码~~ | Markdig | ✅ MD05-05 |
| ~~P2~~ | ~~PDF→MD 文本提取~~ | — | ✅ MD07-01 |
| ~~P3~~ | ~~往返渲染 (Roundtrip)~~ | Markdig | ✅ MD06-03 |
| ~~P3~~ | ~~源码位置追踪~~ | Markdig | ✅ MD06-04 |

> **v3.0 全部差距项已于 2026-07-03 前补齐完成。** 剩余 Markdig 独有但 NewLife.Office 不计划的：Mermaid 图表渲染（需引入 JS 引擎）、媒体嵌入（非 Markdown 核心需求）。

### 3.3 双向转换矩阵（NewLife.Office 特有）

| 源格式 ╲ 目标格式 | MD | HTML | Word | PDF |
|------|:---:|:---:|:---:|:---:|
| **MD** | ✅ 序列化 | ✅ ToHtml/ToHtmlPage | ✅ ToWord/SaveWord | ✅ ToPdf/SavePdf |
| **HTML** | 🔄 P0 | — | — | — |
| **Word** | ✅ IMarkdownExtractable | ✅ WordHtmlConverter | — | ✅ WordPdfConverter |
| **PDF** | 🔄 P2 | — | — | — |

---

## 4. API 易用性对比

实现"解析 Markdown 并输出 HTML"的代码量对比：

**NewLife.Office**:
```csharp
var doc = MarkdownDocument.Parse("# Hello\nWorld");
var html = doc.ToHtml();
```

**Markdig**:
```csharp
var html = Markdown.ToHtml("# Hello\nWorld");
```

**CommonMark.NET**:
```csharp
var html = CommonMarkConverter.Convert("# Hello\nWorld");
```

> 三者均为一行调用，NewLife.Office 额外返回 `MarkdownDocument` 对象可进一步操作（转 Word/PDF），灵活性更高。

---

## 5. 架构演进路线

对标 Markdig 的成熟架构 + NewLife.Office 独有的办公文档转换能力，分四步演进：

| 阶段 | 目标 | 关键交付 | 状态 |
|------|------|---------|:---:|
| **v1.0** | 基础解析 + 正向转换 | CommonMark+GFM 解析，MD→HTML/Word/PDF | ✅ 完成 |
| **v2.0** | AST 继承体系 | 12 种块类型子类（HeadingBlock/CodeBlock/ListItemBlock 等），消除属性混杂 | ✅ 完成 |
| **v2.1** | 反向转换闭环 | HTML→MD（对标 ReverseMarkdown），实现 MD↔HTML/Word 双向全路径 | ✅ 完成 |
| **v3.0** | 管线 + 扩展 + 补齐解析 | MarkdownPipeline 可配置管线，AutoLinks/YAML/脚注/数学/Emoji 扩展 | ✅ 完成 |
| **v4.0** | 全格式转换中枢 | PDF→MD、PPT→MD，文档格式任意方向互转 | ✅ 完成 |

---

## 6. 结论

NewLife.Office 在 **办公文档级 Markdown 转换**（MD→Word/PDF）方面具有独特且不可替代的优势——这是任何纯 Markdown 解析库无法做到的。v2.0 AST 继承体系、v2.1 反向转换闭环、v3.0 管线/扩展/解析增强、v4.0 全格式转换中枢**四大里程碑全部完成**。NewLife.Office 现已实现 **.NET 生态中最完整的办公文档 ↔ Markdown 双向全格式转换矩阵**。

**战略定位**：NewLife.Office 不只是一个 Markdown 解析器，而是 **.NET 生态中唯一的办公文档 ↔ Markdown 双向转换中枢**（已完成）。
