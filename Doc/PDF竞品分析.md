# PDF 竞品分析

> 版本：v2.0 | 日期：2026-08-06
> 返回：[竞品分析报告](竞品分析报告.md)

---

## 1. PDF 竞品概览

| 竞品 | 类型 | 许可 | 定位 |
|------|------|------|------|
| **本项目（NewLife.Office）** | 开源库（v1.3.x） | MIT | 零外部依赖全能型：PDF 创建/读取/编辑/加密/签名/PDF-A 全覆盖，Fluent API，net45 起步 |
| **iText 7** | 开源库（AGPL）/商业 | AGPL / 商业 | 功能最全的 .NET PDF 库，企业级特性（表单/签名/PDF-A），闭源项目必须付费 |
| **QuestPDF** | 开源库 | MIT（年收入>$1M 需付费）/商业 | 现代声明式 Fluent API 生成 PDF，仅创建不支持读取 |
| **PdfSharp / MigraDoc** | 开源库 | MIT | 轻量创建/合并为主：PdfSharp 底层绘图 + MigraDoc 文档模型，读取能力弱 |
| **PdfPig** | 开源库 | Apache 2.0 | 专注 PDF 读取/文本提取，保留位置信息，仅读取 |
| **Aspose.PDF** | 商业库 | 商业 | 功能最全，含转换/OCR/渲染，包体积约 50 MB |

### 1.1 市场与体积概览

| 库名 | 依赖大小 | 外部依赖 | GitHub Stars | NuGet 下载量 |
|------|---------|---------|-------------|------------|
| **NewLife.Office（PDF 模块）** | <500KB（全库） | 1（NewLife.Core） | — | — |
| **iText 7** | ~5MB | 3+ | ~1.5k | 5000万+ |
| **QuestPDF** | ~3MB | 2-3 | ~12k+ | 500万+ |
| **PdfSharp** | ~2MB | 0 | ~4k+ | 2000万+ |
| **PdfPig** | ~3MB | 1-2 | ~1.5k | 200万+ |
| **Aspose.PDF** | ~50MB | 0 | N/A | 500万+ |

---

## 2. PDF 功能对比矩阵

> 标记说明：✅ 完整支持 | ⚠️ 部分可用 | ❌ 不支持。NewLife.Office 列基于 v1.3.x 当前能力实测标注。

### 2.1 创建与生成

| 功能 | NewLife.Office | iText 7 | QuestPDF | PdfSharp | MigraDoc | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 从零创建 PDF | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 文本排版（字体/大小/颜色/粗斜体） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 中文字体（TrueType 嵌入完整映射链） | ✅ | ✅ | ✅ | ⚠️ | ⚠️ | ✅ |
| 表格（斑马纹/表头/边框/合并列） | ✅ | ✅ | ✅ | ❌（需手绘） | ✅ | ✅ |
| 图片插入 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 多页文档 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 页眉/页脚（含页码） | ✅ | ✅ | ✅ | 手动 | ✅ | ✅ |
| 目录生成 | ⚠️ | ⚠️ | ✅ | ❌ | ✅ | ✅ |
| 水印（文字/图片叠加） | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 条码/二维码 | ✅ QR码 | ✅ | 需第三方 | ❌ | ❌ | ✅ |

### 2.2 读取与提取

| 功能 | NewLife.Office | iText 7 | PdfPig | PdfSharp | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|
| 文本提取（PDF 1.0-1.7） | ✅ | ✅ | ✅ | ⚠️ | ✅ |
| 逐页提取 | ✅ | ✅ | ✅ | ⚠️ | ✅ |
| 带坐标文本提取 | ✅ | ✅ | ✅ | ❌ | ✅ |
| 图片提取 | ✅ | ✅ | ✅ | ❌ | ✅ |
| 元数据读取 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 页面信息（页数/MediaBox/旋转） | ✅ | ✅ | ✅ | ✅ | ✅ |
| 书签/大纲读取（递归 Outline 树） | ✅ | ✅ | ✅ | ❌ | ✅ |
| 内容流重建（操作符序列） | ✅ | ⚠️ | ⚠️ | ❌ | ✅ |
| 页面渲染为图片 | ⚠️（Rendering 扩展包） | ❌ | ❌ | ❌ | ✅ |
| 结构化数据提取（表格） | ❌ | ⚠️ | ⚠️ | ❌ | ✅ |

### 2.3 编辑与操作

| 功能 | NewLife.Office | iText 7 | PdfSharp | PdfPig | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|
| 合并 PDF | ✅ | ✅ | ✅ | ❌ | ✅ |
| 拆分 PDF | ✅ | ✅ | ✅ | ❌ | ✅ |
| 页面旋转/删除/重排 | ✅ | ✅ | ✅ | ❌ | ✅ |
| 叠加文字/图片（盖章/水印/页码） | ✅ | ✅ | ✅ | ❌ | ✅ |
| 对象级编辑 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 书签/大纲读写 | ✅ | ✅ | 读取 | ❌ | ✅ |

### 2.4 高级功能（加密/签名/表单/注释/PDF-A）

| 功能 | NewLife.Office | iText 7 | QuestPDF | PdfSharp | PdfPig | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 加密（RC4 40/128-bit） | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| 加密（AES-128/256） | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| 权限控制 | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| 数字签名（PKCS#7 分离签名+可见签名域） | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| AcroForm 表单创建（文本框/复选框/下拉框/签名） | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| 表单填充 | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| 注释/批注（全类型写入） | ✅ | ✅ | 读取 | 读取 | ❌ | ✅ |
| 超链接 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| PDF/A 合规（A-1B/2B/3B） | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |

### 2.5 Fluent API 与开发体验

| 功能 | NewLife.Office | QuestPDF | iText 7 | PdfSharp | MigraDoc |
|------|:---:|:---:|:---:|:---:|:---:|
| 声明式布局（Row/Column/Text/Image/Table） | ✅ | ✅ | ❌ | ❌ | ❌ |
| 自动分页 | ✅ | ✅ | ⚠️（手动布局） | ❌ | ✅ |
| 组件复用 | ✅ | ✅ | ❌ | ❌ | ⚠️ |
| 低层绘图 API（线/矩形/圆角矩形/椭圆/多边形/圆弧/渐变/贝塞尔） | ✅ | ✅（Canvas） | ✅（PdfCanvas） | ✅（XGraphics） | ❌ |
| 虚线/点线/透明度/图片旋转/字符间距 | ✅ | ✅ | ✅ | ✅ | ❌ |
| 自定义页码 {page}/{total} | ✅ | ✅ | ⚠️ | ⚠️ | ✅ |

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
**劣势**：PdfSharp 无高层表格 API（需手动绘制线条），文本提取能力有限，无 xref 解析，无表单/签名/加密支持。  
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

### 4.1 核心优势

- **vs PdfSharp**：功能更全面（文本提取、合并拆分、水印、加密签名等），提供高层 Fluent API，读取端具备完整 xref 解析
- **vs iText 7**：完全免费，无 AGPL 限制，闭源项目可放心使用；提供 xref+FlateDecode 读取、AcroForm 表单创建、数字签名、PDF/A 等高级能力
- **vs QuestPDF**：支持读取和编辑（QuestPDF 仅生成），支持 net45，无年收入门槛
- **vs PdfPig**：同时支持读取和创建（PdfPig 仅读取），且读取端提供 xref+解压+书签等完整能力

### 4.2 许可证优势

| 库 | 许可证 | 陷阱说明 | 风险评估 |
|---|--------|---------|---------|
| **iText 7** | AGPLv3 | **病毒式传染**：只要你的软件"通过网络提供服务"（如 Web API 生成 PDF），就必须开源全部代码。商业许可 $2,500+/年 | 🔴 高风险 |
| **QuestPDF** | MIT → 商业 | 年收入 < $1M 免费；超过后必须购买商业许可（$599/年起）。动态检测许可，超限后抛异常 | 🟡 中等风险 |
| **PdfSharp** | MIT | 真正免费，但功能有限（无 xref 解析、无解压） | 🟢 低风险 |
| **PdfPig** | Apache 2.0 | 真正免费，但仅读取（不能创建 PDF） | 🟢 低风险 |
| **Aspose.PDF** | 商业 | $999/年起，按开发者席位收费，无上限 | 🔴 高成本 |
| **NewLife.Office** | MIT | 完全免费，无任何限制，闭源商用均无法律风险 | 🟢 零风险 |

### 4.3 目标赛道

PDF 库的许可证问题是行业痛点。NewLife.Office 以 MIT 许可提供全面的 PDF 操作能力，无论创建、读取、编辑、合并拆分、加密签名均可免费商用，定位为**闭源商业项目零许可风险的 PDF 全功能方案**，与 NewLife.Core 生态（XCode/Cube 等）深度集成，服务报表生成、文档处理、数据导出等企业场景。

---

## 5. 技术架构深度对比

### 5.1 PDF 解析架构（xref）

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
| LZWDecode | ✅ | ✅ | ✅ | ❌ |
| RunLengthDecode | ✅ | ✅ | ✅ | ❌ |

### 5.3 字体处理对比

中文字体处理是 PDF 库的核心难点之一。NewLife.Office 实现了系统 TrueType 嵌入 + CIDFontType2 + Identity-H + CIDToGIDMap + ToUnicode 的完整映射链，保证中文 PDF 生成与提取双向无乱码。

| 能力 | NewLife.Office | iText 7 | QuestPDF | PdfSharp |
|------|:---:|:---:|:---:|:---:|
| 系统 TrueType 嵌入 | ✅ 完整映射链 | ✅ | ✅ | ❌ |
| CIDFontType2 + Identity-H | ✅ | ✅ | ✅ | ❌ |
| CIDToGIDMap 流 | ✅ 自动生成 | ✅ | ✅ | ❌ |
| ToUnicode CMap | ✅ Identity-UCS2 | ✅ | ✅ | ❌ |
| 字体子集化 | ❌ | ✅ | ✅ | ❌ |
| Adobe CJK 回退 | ✅ STSong-Light | — | — | ❌ |

### 5.4 加密与签名架构

| 能力 | NewLife.Office | iText 7 | PdfSharp | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|
| RC4 40/128-bit 加密 | ✅ PdfEncryptor 策略模式 | ✅ | ❌ | ✅ |
| AES-128/256 加密 | ✅ | ✅ | ❌ | ✅ |
| 权限控制（打印/复制/修改） | ✅ | ✅ | ❌ | ✅ |
| 数字签名 | ✅ PdfSigner，PKCS#7 分离签名 + 可见签名域（手动 ASN.1 DER） | ✅ | ❌ | ✅ |

---

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

### 6.3 加密与数字签名（NewLife.Office）

**AES-256 加密 + 权限控制**：
```csharp
using var writer = new PdfWriter();
writer.UserPassword = "user123";
writer.OwnerPassword = "owner456";
writer.Permissions = -1;                        // 全权限；按位与可控制打印/复制/修改
writer.CipherRevision = CipherRevision.Aes_256;
writer.DrawText("加密文档", 40, 40, 20);
writer.Save("encrypted.pdf");
```

**PKCS#7 分离签名 + 可见签名域**：
```csharp
using var cert = new X509Certificate2("cert.pfx", "password");
PdfSigner.Sign("input.pdf", "signed.pdf", cert);
```

## 7. 结论

### 7.1 综合评估

| 维度 | NewLife.Office | iText 7 | QuestPDF | PdfSharp | PdfPig | Aspose.PDF |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 功能覆盖 | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐ | ⭐⭐ | ⭐⭐⭐⭐⭐ |
| 许可风险 | 🟢 零 | 🔴 AGPL | 🟡 收入门槛 | 🟢 | 🟢 | 🔴 高成本 |
| 依赖体积 | <500KB | ~5MB | ~3MB | ~2MB | ~3MB | ~50MB |
| 框架兼容 | net45+ | netstandard2.0+ | net6.0+ | netstandard2.0+ | netstandard2.0+ | netstandard2.0+ |
| 读写编辑一体化 | ✅ | ✅ | ❌（仅创建） | ⚠️（读取弱） | ❌（仅读取） | ✅ |

### 7.2 选型决策树

```
需要什么能力？
├── 仅创建 PDF
│   ├── 需要 MIT 且零门槛 → NewLife.Office / PdfSharp
│   ├── 需要最佳声明式 API → QuestPDF（注意收入门槛）
│   └── 需要全功能企业级 → Aspose.PDF（商业）
├── 仅读取 PDF
│   ├── 需要 MIT 且含文本提取 → NewLife.Office / PdfPig
│   ├── 需要页面渲染为图片 → NewLife.Office.Rendering / Docnet
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
    ├── MIT 许可 → NewLife.Office（A-1B/2B/3B + PKCS#7 签名）
    ├── 开源项目 → iText 7 (AGPL)
    └── 商业 → Aspose.PDF
```

### 7.3 真实未实现项（已知边界）

以下能力当前不在 NewLife.Office 主库范围内，需要时请评估第三方方案：

- **页面渲染为图片**：主库不包含渲染引擎，由 `NewLife.Office.Rendering` 扩展包提供文档预览能力（`DocumentPreview`）
- **HTML → PDF 高保真**：不支持 HTML 排版引擎级别的转换
- **OCR**：不支持扫描件文字识别
- **字体子集化**：中文字体以完整 TrueType 嵌入，暂不支持子集化瘦身

### 7.4 总结

NewLife.Office 是免费库中**读写编辑一体化能力最完整**的 PDF 方案：创建端提供 QuestPDF 风格的 Fluent API 与完整中文字体映射链，读取端具备 xref/内容流/对象级的高保真解析，编辑端覆盖合并拆分/水印/表单/加密/签名/PDF-A，且 MIT 许可 + 仅依赖 NewLife.Core + net45 起步，是闭源商业项目零许可风险的最安全选择。若需要页面级渲染、HTML 转换或 OCR 等延伸能力，可组合 `NewLife.Office.Rendering` 与第三方库使用。

---

← 返回 [竞品分析报告.md](竞品分析报告.md)
