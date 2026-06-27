# PDF 竞品分析

← 返回 [竞品分析报告.md](竞品分析报告.md)

---

## 1. PDF 竞品概览

| 库名 | 许可证 | 依赖大小 | 主要能力 | GitHub Stars | NuGet下载量 |
|------|--------|---------|---------|-------------|------------|
| **iText 7** | AGPL / 商业 | ~5MB | 创建/编辑/表单/签名 | ~1.5k | 5000万+ |
| **QuestPDF** | MIT（年收入<$1M）/ 商业 | ~3MB | 创建（Fluent API） | ~12k+ | 500万+ |
| **PdfSharp** | MIT | ~2MB | 创建/合并/基础编辑 | ~4k+ | 2000万+ |
| **MigraDoc** | MIT | ~2MB（含PdfSharp） | 文档模型→PDF | 同PdfSharp | 同PdfSharp |
| **PdfPig** | Apache 2.0 | ~3MB | 读取/文本提取 | ~1.5k | 200万+ |
| **Docnet** | MIT | ~5MB（含PDFium） | 读取/渲染 | ~300+ | 50万+ |
| **Aspose.PDF** | 商业 | ~50MB | 全功能 | N/A | 500万+ |
| **Spire.PDF Free** | 免费受限/商业 | ~20MB | 创建/编辑（10页限制） | N/A | 200万+ |
| **IronPDF** | 商业 | ~100MB+ | HTML→PDF/编辑 | N/A | 300万+ |

---

## 2. PDF 功能对比矩阵

### 2.1 创建与生成

| 功能 | NewLife.Office | iText 7 | QuestPDF | PdfSharp | MigraDoc | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 从零创建 PDF | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 文本排版 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 表格 | ✅ | ✅ | ✅ | ❌（需手绘） | ✅ | ✅ |
| 图片插入 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 多页文档 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 页眉/页脚 | ✅ | ✅ | ✅ | 手动 | ✅ | ✅ |
| 目录生成 | 部分 | 部分 | ✅ | ❌ | ✅ | ✅ |
| 水印 | 部分 | ✅ | ✅ | ✅ | ❌ | ✅ |
| 条码/二维码 | ❌ | ✅ | 需第三方 | ❌ | ❌ | ✅ |
| 中文字体支持 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Fluent API | ✅ | ❌ | ✅ | ❌ | 否 | ❌ |

### 2.2 读取与提取

| 功能 | NewLife.Office | iText 7 | PdfPig | PdfSharp | Docnet | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 文本提取 | ✅ 含 FlateDecode | ✅ | ✅ | 部分 | ✅ | ✅ |
| 带位置信息的文本提取 | ✅ xref+解压 | ✅ | ✅ | ❌ | 部分 | ✅ |
| 图片提取 | ✅ xref+解压 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 元数据读取 | ✅ xref优先 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 页面渲染为图片 | 🔄规划 | ❌ | ❌ | ❌ | ✅ | ✅ |
| 结构化数据提取（表格） | ❌ | 部分 | 部分 | ❌ | ❌ | ✅ |

### 2.3 编辑与操作

| 功能 | NewLife.Office | iText 7 | PdfSharp | PdfPig | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|
| 合并 PDF | ✅ | ✅ | ✅ | ❌ | ✅ |
| 拆分 PDF | ✅ | ✅ | ✅ | ❌ | ✅ |
| 页面旋转/删除/重排 | ✅ | ✅ | ✅ | ❌ | ✅ |
| 添加文字/图片覆盖 | ✅ | ✅ | ✅ | ❌ | ✅ |
| 表单填充（AcroForm） | ✅ 创建/填充 | ✅ | ❌ | ❌ | ✅ |
| 数字签名 | ❌ | ✅ | ❌ | ❌ | ✅ |
| 加密/权限控制 | ✅ RC4 | ✅ RC4/AES | ❌ | ❌ | ✅ RC4/AES |
| PDF/A 合规 | ❌ | ✅ | ❌ | ❌ | ✅ |
| 书签/大纲 | ✅ 读写 | ✅ 读写 | 读取 | ❌ | ✅ |
| 注释/批注 | ✅ Link | ✅ 全部 | 读取 | ❌ | ✅ |

---

## 3. PDF 竞品优劣势分析

### 3.1 iText 7

**优势**：.NET 生态中功能最全面的 PDF 库，支持创建、编辑、表单、签名、PDF/A 等企业级特性，文档丰富。  
**劣势**：**AGPL 许可**要求使用者也开源，否则需购买商业许可（价格不低）；API 较复杂。  
**亮点**：表单填充、数字签名、PDF/A 合规等企业级功能是其核心竞争力。

### 3.2 QuestPDF

**优势**：现代化 Fluent API 设计，开发体验极佳，MIT 许可（年收入<$1M），支持热重载预览，社区活跃度极高（12k+ Stars）。  
**劣势**：**仅支持创建**，不能读取或编辑已有 PDF；年收入>$1M 的公司需商业许可。  
**亮点**：C# 声明式布局、自动分页、组件复用的设计理念，在 PDF 生成领域代表了最先进的开发体验。

### 3.3 PdfSharp / MigraDoc

**优势**：MIT 许可，轻量，PdfSharp 提供底层绘图 API，MigraDoc 提供文档模型（段落/表格/图片），可搭配使用。  
**劣势**：PdfSharp 无高层表格 API（需手动绘制线条），文本提取能力有限，无表单/签名支持。  
**亮点**：PDF 合并功能简洁高效；MigraDoc 可同时输出 PDF 和 RTF。

### 3.4 PdfPig

**优势**：Apache 2.0 许可，专注 PDF 文本提取，支持逐字/逐行提取并保留位置信息，适合数据挖掘场景。  
**劣势**：**只读**，不能创建或编辑 PDF。  
**亮点**：文本提取的精度和位置信息获取能力在免费库中最优。

### 3.5 Aspose.PDF

**优势**：功能最全，支持创建/编辑/转换/表单/签名/OCR 等全部 PDF 操作，零原生依赖。  
**劣势**：商业许可价格高昂，包体积极大（~50MB），闭源。  
**亮点**：HTML→PDF 转换、PDF→Word/Excel 反向转换的能力是其独特卖点。

---

## 4. 差异化定位

- **vs PdfSharp**：功能更全面（文本提取、合并拆分、水印等），提供高层 Fluent API
- **vs iText 7**：完全免费，无 AGPL 限制，闭源项目可放心使用；提供 xref+FlateDecode 读取、AcroForm 表单创建等高级能力
- **vs QuestPDF**：支持读取和编辑（QuestPDF 仅生成），支持 net45，无年收入门槛
- **vs PdfPig**：同时支持读取和创建（PdfPig 仅读取），且读取端提供 xref+解压+书签等完整能力

PDF 库的许可证问题是行业痛点。NewLife.Office 以 MIT 许可提供全面的 PDF 操作能力，无论创建、读取、编辑、合并拆分均可免费商用。

---

## 5. 技术架构深度对比

### 5.1 PDF 解析架构

PDF 文件结构的核心是**交叉引用表（xref）**——它是一个索引，记录了每个对象的字节偏移量。正确解析 xref 表是可靠读取 PDF 的前提。

| 能力 | NewLife.Office | iText 7 | PdfPig | PdfSharp |
|------|:---:|:---:|:---:|:---:|
| 传统 xref 表解析 | ✅ | ✅ | ✅ | ❌ |
| xref 流（PDF 1.5+） | ✅ | ✅ | ✅ | ❌ |
| 增量更新链（/Prev） | ✅ | ✅ | ✅ | ❌ |
| 对象流（ObjStm） | ✅ | ✅ | ✅ | ❌ |
| 纯字符串扫描（无 xref） | 回退方案 | — | — | ✅（唯一方式） |

**关键差异**：PdfSharp 使用字符串扫描方式定位 `stream`/`endstream` 关键字，在二进制流中包含 "endstream" 字面量时会误匹配。NewLife.Office 实现了完整的 xref 表解析器（`PdfXRefTable`），按对象号精确定位，是免费库中读取可靠性最高的方案。

### 5.2 内容流解压缩

绝大多数 PDF 文件使用 FlateDecode（zlib/Deflate）压缩内容流以减小体积。不解压缩则无法正确提取文本。

| 能力 | NewLife.Office | iText 7 | PdfPig | PdfSharp |
|------|:---:|:---:|:---:|:---:|
| FlateDecode（zlib） | ✅ DeflateStream | ✅ | ✅ | ❌ |
| ASCII85Decode | ✅ | ✅ | ✅ | ❌ |
| ASCIIHexDecode | ✅ | ✅ | ✅ | ❌ |
| 多重过滤器链 | ✅ | ✅ | ✅ | ❌ |
| LZWDecode | ❌ | ✅ | ✅ | ❌ |

### 5.3 字体处理对比

中文字体处理是 PDF 库的核心难点之一。

| 能力 | NewLife.Office | iText 7 | QuestPDF | PdfSharp |
|------|:---:|:---:|:---:|:---:|
| 系统 TrueType 嵌入 | ✅ 完整映射链 | ✅ | ✅ | ❌ |
| CIDFontType2 + Identity-H | ✅ | ✅ | ✅ | ❌ |
| CIDToGIDMap 流 | ✅ 自动生成 | ✅ | ✅ | ❌ |
| ToUnicode CMap | ✅ Identity-UCS2 | ✅ | ✅ | ❌ |
| 字体子集化 | ❌ | ✅ | ✅ | ❌ |
| Adobe CJK 回退 | ✅ STSong-Light | — | — | ❌ |

## 6. API 易用性对比（代码示例）

### 6.1 创建含表格的 PDF

**NewLife.Office（Fluent API）**：
```csharp
using var doc = new PdfFluentDocument();
doc.Title = "报表";
doc.AddText("销售数据汇总", fontSize: 20)
   .AddEmptyLine()
   .AddTable(new[] {
       new[]{"姓名","部门","销售额"},
       new[]{"张三","技术","¥120,000"},
       new[]{"李四","销售","¥250,000"},
   }, firstRowHeader: true);
doc.Save("report.pdf");
```

**iText 7**（~25 行，需手动管理 Document/PdfWriter/Table/Cell 对象）：
```csharp
using var pdf = new PdfDocument(new PdfWriter("report.pdf"));
using var doc = new Document(pdf);
doc.Add(new Paragraph("销售数据汇总").SetFontSize(20));
var table = new Table(3);
table.AddHeaderCell("姓名"); table.AddHeaderCell("部门"); table.AddHeaderCell("销售额");
table.AddCell("张三"); table.AddCell("技术"); table.AddCell("¥120,000");
doc.Add(table); doc.Close();
```

**QuestPDF**（声明式，最简洁但仅创建）：
```csharp
Document.Create(c => c.Page(p => {
    p.Content().Column(c => {
        c.Item().Text("销售数据汇总").FontSize(20);
        c.Item().Table(t => {
            t.ColumnsDefinition(c => { c.RelativeColumn(); c.RelativeColumn(); c.RelativeColumn(); });
            t.Header(h => { h.Cell().Text("姓名"); h.Cell().Text("部门"); h.Cell().Text("销售额"); });
            t.Cell().Text("张三"); t.Cell().Text("技术"); t.Cell().Text("¥120,000");
        });
    });
})).GeneratePdf("report.pdf");
```

### 6.2 文本提取

**NewLife.Office**（3 行）：
```csharp
using var reader = new PdfReader("input.pdf");
var text = reader.ExtractText();
var meta = reader.ReadMetadata();
```

**PdfPig**（5 行）：
```csharp
using var doc = PdfDocument.Open("input.pdf");
var text = string.Join(" ", doc.GetPages().Select(p => p.Text));
```

**iText 7**（8 行，需手动遍历策略）：
```csharp
using var pdf = new PdfDocument(new PdfReader("input.pdf"));
var strategy = new SimpleTextExtractionStrategy();
for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
    PdfTextExtractor.GetTextFromPage(pdf.GetPage(i), strategy);
```

## 7. 许可证陷阱深度分析

| 库 | 许可证 | 陷阱说明 | 风险评估 |
|---|--------|---------|---------|
| **iText 7** | AGPLv3 | **病毒式传染**：只要你的软件"通过网络提供服务"（如 Web API 生成 PDF），就必须开源全部代码。商业许可 $2,500+/年 | 🔴 高风险 |
| **QuestPDF** | MIT → 商业 | 年收入 < $1M 免费；超过后必须购买商业许可（$599/年起）。动态检测许可，超限后抛异常 | 🟡 中等风险 |
| **PdfSharp** | MIT | 真正免费，但功能有限（无 xref 解析、无解压） | 🟢 低风险 |
| **PdfPig** | Apache 2.0 | 真正免费，但仅读取（不能创建 PDF） | 🟢 低风险 |
| **Aspose.PDF** | 商业 | $999/年起，按开发者席位收费，无上限 | 🔴 高成本 |
| **Spire.PDF Free** | 免费受限 | 免费版限 10 页/文档，超出自动截断 | 🟡 受限 |
| **NewLife.Office** | MIT | 完全免费，无任何限制，闭源商用均无法律风险 | 🟢 零风险 |

> **关键建议**：如果你的项目涉及 Web API 生成 PDF 且不愿开源，iText 7 的 AGPL 是法律陷阱。QuestPDF 的年收入门槛对快速增长的公司也是隐患。NewLife.Office PDF 以 MIT 许可消除了这些风险，是闭源项目的最安全选择。

## 8. 选型决策树

```
需要什么能力？
├── 仅创建 PDF
│   ├── 需要 MIT 且零门槛 → NewLife.Office / PdfSharp
│   ├── 需要最佳声明式 API → QuestPDF（注意收入门槛）
│   └── 需要全功能企业级 → Aspose.PDF（商业）
├── 仅读取 PDF
│   ├── 需要 MIT 且含文本提取 → NewLife.Office / PdfPig
│   ├── 需要页面渲染为图片 → Docnet（需 PDFium 原生依赖）
│   └── 需要结构化表格提取 → iText 7 / Aspose.PDF
├── 读 + 写 + 编辑
│   ├── MIT 许可首选 → NewLife.Office（功能最全的免费选择）
│   ├── 商业项目且预算充足 → Aspose.PDF
│   └── 开源项目可接受 AGPL → iText 7
├── 表单填充（AcroForm）
│   ├── MIT 许可 → NewLife.Office（支持创建/填充）
│   ├── 开源项目 → iText 7
│   └── 商业 → Aspose.PDF
└── 数字签名 / PDF/A
    ├── 商业 → Aspose.PDF
    └── 开源 → iText 7 (AGPL)
```

---

← 返回 [竞品分析报告.md](竞品分析报告.md)
