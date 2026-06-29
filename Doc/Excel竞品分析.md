# Excel 竞品分析

← 返回 [竞品分析报告.md](竞品分析报告.md)

---

## 1. Excel 竞品概览

| 库名 | 许可证 | 最新版本 | 依赖大小 | 支持格式 | GitHub Stars | NuGet下载量 |
|------|--------|---------|---------|---------|-------------|------------|
| **EPPlus** | Polyform Noncommercial / 商业 | 8.x | ~5MB | xlsx | ~2k | 1亿+ |
| **NPOI** | Apache 2.0 + OSMFEULA（2.8+） | 2.8.x | ~10MB | xls/xlsx/docx | ~6.2k | 5000万+ |
| **ClosedXML** | MIT | 0.105.x | ~3MB（含OpenXML SDK） | xlsx | ~5.6k | 4000万+ |
| **MiniExcel** | Apache 2.0 | 1.x | <1MB | xlsx/csv | ~2.5k | 500万+ |
| **Open XML SDK** | MIT | 3.x | ~2MB | xlsx/docx/pptx | ~4k | 3000万+ |
| **Aspose.Cells** | 商业 | 25.x | ~30MB | xls/xlsx/csv/pdf等 | N/A | 1000万+ |
| **ExcelDataReader** | MIT | 3.x | <1MB | xls/xlsx（只读） | ~3.5k | 5000万+ |

> ⚠️ **NPOI 注意**：v2.8.0 引入 OSMFEULA，要求营利性组织支付月度维护费，不再完全免费商用。  
> ⚠️ **EPPlus 注意**：v8 改为 Polyform Noncommercial 许可，个人/非商业免费，商业必须购买许可。

---

## 2. Excel 功能对比矩阵

### 2.1 基础读写

| 功能 | NewLife.Office | EPPlus 8 | NPOI 2.8 | ClosedXML 0.105 | MiniExcel | Aspose |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 读取 xlsx | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 写入 xlsx | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 读取 xls | ✅ | ❌ | ✅ | ❌ | ✅ | ✅ |
| 写入 xls | ⚠️部分可用 | ❌ | ✅ | ❌ | ❌ | ✅ |
| CSV 支持 | ✅（NewLife.Core CsvFile） | ❌ | ❌ | ❌ | ✅ | ✅ |
| 多工作表 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 流式读取 | ✅ | 部分 | 部分 | ❌ | ✅ | ✅ |
| 流式写入 | ✅ | ❌ | ✅ | ❌ | ✅ | ✅ |

### 2.2 单元格与样式

| 功能 | NewLife.Office | EPPlus | NPOI | ClosedXML | MiniExcel | Aspose |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 字体样式 | ✅ | ✅ | ✅ | ✅ | 部分 | ✅ |
| 背景色 | ✅ | ✅ | ✅ | ✅ | 部分 | ✅ |
| 边框 | ✅ | ✅ | ✅ | ✅ | 部分 | ✅ |
| 对齐方式 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 自定义数字格式 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 条件格式 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |

### 2.3 布局与结构

| 功能 | NewLife.Office | EPPlus | NPOI | ClosedXML | MiniExcel | Aspose |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 合并单元格 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 冻结窗格 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 自动筛选 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 列宽设置 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 行高设置 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 自动列宽 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |

### 2.4 高级数据

| 功能 | NewLife.Office | EPPlus | NPOI | ClosedXML | MiniExcel | Aspose |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 超链接 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 数据验证 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 公式 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 批注 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 数据透视表 | 部分 | ✅ | 部分 | ❌ | ❌ | ✅ |

### 2.5 图片与图表

| 功能 | NewLife.Office | EPPlus | NPOI | ClosedXML | MiniExcel | Aspose |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 插入图片 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 图表（Writer 集成） | ✅（模型+M21完成） | ✅ | ✅ | 部分 | ❌ | ✅ |

### 2.6 打印与页面

| 功能 | NewLife.Office | EPPlus | NPOI | ClosedXML | MiniExcel | Aspose |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 页面方向/纸张 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 页边距 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 页眉页脚 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 打印标题行 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| 工作表保护 | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |

### 2.7 便捷 API

| 功能 | NewLife.Office | EPPlus | NPOI | ClosedXML | MiniExcel | Aspose |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| 对象映射导出 | ✅ | 部分 | ❌ | ❌ | ✅ | 部分 |
| 对象映射导入 | ✅ | 部分 | ❌ | ❌ | ✅ | 部分 |
| DataTable 支持 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 模板填充 | ✅ | ✅ | ❌ | ❌ | ✅ | ✅ |
| Attribute 映射 | ✅ | 部分 | ❌ | ❌ | ✅ | 部分 |
| ExcelHelper 静态API（MiniExcel风格） | ✅ | ❌ | ❌ | ❌ | ✅ | 部分 |

> ✅ = 已支持 | ❌ = 不支持 | 部分 = 有限支持 | ⚠️ = 部分可用

---

### 2.8 高级数据结构（高保真差距分析）

> 以下为 NewLife.Office 与竞品在高保真 xlsx 读写方面的关键差距，对应 M18-M19 待实现功能。

| 功能 | NewLife.Office | EPPlus 8 | NPOI 2.8 | ClosedXML 0.105 | Aspose.Cells |
|------|:---:|:---:|:---:|:---:|:---:|
| 命名范围 (Defined Names) | ✅ | ✅ | ✅ | ✅ | ✅ |
| 结构化表格 (`<table>` 元素) | ✅ | ✅ | ✅ | ✅ | ✅ |
| 表格样式/带状行 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 单元格富文本 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 渐变填充 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 图案填充 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 对角线边框 | ✅ | ✅ | ✅ | ✅ | ✅ |

### 2.9 图表与可视化（高保真差距分析）

> 对应 M21-M22 待实现功能。

| 功能 | NewLife.Office | EPPlus 8 | NPOI 2.8 | ClosedXML 0.105 | Aspose.Cells |
|------|:---:|:---:|:---:|:---:|:---:|
| 图表集成到 Writer (AddChart) | ✅ | ✅ | ✅ | ✅ | ✅ |
| 图表数据读取 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 迷你图 (Sparklines) | ✅ | ✅ | ❌ | ❌ | ✅ |
| 条件格式图标集 (IconSet) | ✅ | ✅ | ✅ | ✅ | ✅ |
| 条件格式自定义公式 | ✅ | ✅ | ✅ | ✅ | ✅ |

### 2.10 布局与排版（高保真差距分析）

> 对应 M20、M23 待实现功能。

| 功能 | NewLife.Office | EPPlus 8 | NPOI 2.8 | ClosedXML 0.105 | Aspose.Cells |
|------|:---:|:---:|:---:|:---:|:---:|
| 四边独立边框 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 文本旋转 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 缩进 (Indent) | ✅ | ✅ | ✅ | ✅ | ✅ |
| 缩小以填充 (ShrinkToFit) | ✅ | ✅ | ✅ | ✅ | ✅ |
| 删除线/上标/下标 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 列分组/大纲级别 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 行分组/大纲级别 | ✅ | ✅ | ✅ | ✅ | ✅ |

### 2.11 企业特性（高保真差距分析）

> 对应 M24 待实现功能。

| 功能 | NewLife.Office | EPPlus 8 | NPOI 2.8 | ClosedXML 0.105 | Aspose.Cells |
|------|:---:|:---:|:---:|:---:|:---:|
| 工作表标签颜色 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 工作簿保护（结构/窗口） | ✅ | ✅ | ✅ | ✅ | ✅ |
| 计算选项 (`<calcPr>`) | ✅ | ✅ | 部分 | ✅ | ✅ |
| 切片器 (Slicers) | ❌ | ✅ | ❌ | ❌ | ✅ |
| 线程化批注 | ❌ | ✅ | ❌ | ❌ | ✅ |

---

## 3. Excel 非功能对比

### 3.1 依赖与体积

| 库 | 外部依赖数 | 包体积 | 运行时内存 |
|---|-----------|--------|-----------|
| **NewLife.Office** | 1（NewLife.Core） | <100KB | 极低 |
| EPPlus | 2-3 | ~5MB | 中 |
| NPOI | 5+ | ~10MB | 高 |
| ClosedXML | 3+（含OpenXML SDK） | ~3MB | 中高 |
| MiniExcel | 0 | <200KB | 极低 |
| Open XML SDK | 1 | ~2MB | 中 |
| Aspose.Cells | 0 | ~30MB | 高 |

### 3.2 框架兼容性

| 库 | net45 | netstandard2.0 | net6.0+ | net8.0+ |
|---|:---:|:---:|:---:|:---:|
| **NewLife.Office** | ✅ | ✅ | ✅ | ✅ |
| EPPlus | ❌(v5+) | ✅ | ✅ | ✅ |
| NPOI | ✅ | ✅ | ✅ | ✅ |
| ClosedXML | ❌ | ✅ | ✅ | ✅ |
| MiniExcel | ❌ | ✅ | ✅ | ✅ |
| Aspose.Cells | ✅ | ✅ | ✅ | ✅ |

### 3.3 许可证风险

| 库 | 许可 | 商业使用风险 |
|---|------|------------|
| **NewLife.Office** | MIT | 无，完全免费 |
| EPPlus | Polyform Noncommercial / 商业 | **v5+商业使用需付费** |
| NPOI | Apache 2.0 + OSMFEULA（v2.8+） | **营利性组织需支付维护费** |
| ClosedXML | MIT | 无 |
| MiniExcel | Apache 2.0 | 低 |
| Aspose.Cells | 商业 | **必须购买许可** |

---

## 4. Excel 竞品优劣势分析

### 4.1 EPPlus

**优势**：功能最全面的开源方案，API 设计友好，文档丰富，Excel 特性覆盖率极高。  
**劣势**：v5 起商业使用需购买许可证（v8 进一步收紧）；不支持 xls 格式；内存占用较高。

### 4.2 NPOI

**优势**：支持 xls/xlsx，功能覆盖广泛，社区活跃。  
**劣势**：包体积大（~10MB），依赖多，API 较底层（Java POI 风格），内存占用高；**v2.8.0 起营利性组织需付费**。

### 4.3 ClosedXML

**优势**：API 友好、代码可读性好，MIT 许可，0.105 版本持续活跃维护。  
**劣势**：依赖 Open XML SDK（较重），不支持 net45，不支持 xls。

### 4.4 MiniExcel

**优势**：极致轻量，流式读写，内存占用极低，支持模板填充和对象映射。  
**劣势**：样式支持有限，不支持合并单元格、冻结窗格等高级特性，不支持图表。

### 4.5 Aspose.Cells

**优势**：功能最全，企业级品质，支持几乎所有 Excel 特性（含 Sparklines/Slicers 等高端特性）。  
**劣势**：**价格高昂**（单开发者License数千美元），闭源，包体积大（~30MB）。

---

## 5. API 代码易用性对比

以下对比各库完成最典型 Excel 操作所需代码量与风格，直观体现 NewLife.Office 的开发效率优势。

### 5.1 对象集合导出

```csharp
// ✅ NewLife.Office（ExcelHelper 最简洁模式）— 1 行完成
ExcelHelper.SaveAs("report.xlsx", users);
```

```csharp
// ✅ NewLife.Office（ExcelWriter 完整控制）— 样式随行设置
using var writer = new ExcelWriter("report.xlsx");
writer.WriteObjects("Sheet1", users, new CellStyle { Bold = true, Background = "4472C4", ForeColor = "FFFFFF" });
writer.Save();
```

```csharp
// EPPlus — LoadFromCollection 完成映射，样式需逐步设置
using var package = new ExcelPackage("report.xlsx");
var sheet = package.Workbook.Worksheets.Add("Sheet1");
sheet.Cells["A1"].LoadFromCollection(users, PrintHeaders: true);
package.Save();
```

```csharp
// NPOI — 无内置对象映射，需完整手写反射循环（约 20 行）
var workbook = new XSSFWorkbook();
var sheet = workbook.CreateSheet("Sheet1");
var props = typeof(User).GetProperties();
var header = sheet.CreateRow(0);
for (var i = 0; i < props.Length; i++) header.CreateCell(i).SetCellValue(props[i].Name);
var rowIdx = 1;
foreach (var u in users)
{
    var row = sheet.CreateRow(rowIdx++);
    for (var i = 0; i < props.Length; i++)
        row.CreateCell(i).SetCellValue(props[i].GetValue(u)?.ToString());
}
using var fs = File.Create("report.xlsx");
workbook.Write(fs);
```

```csharp
// MiniExcel — 最简洁，但不支持任何样式设置
await MiniExcel.SaveAsAsync("report.xlsx", users);
```

### 5.2 对象集合导入

```csharp
// ✅ NewLife.Office（ExcelHelper 最简洁模式）— 1 行完成
var list = ExcelHelper.Query<User>("report.xlsx").ToList();
```

```csharp
// ✅ NewLife.Office（ExcelReader 显式控制）— 自动按列名/DisplayName/Description 映射
using var reader = new ExcelReader("report.xlsx");
var list2 = reader.ReadObjects<User>().ToList();
```

```csharp
// EPPlus — 需指定列映射关系，代码量中等
using var package = new ExcelPackage("report.xlsx");
var sheet = package.Workbook.Worksheets[0];
var list = sheet.Cells["A1:Z1000"].ToCollectionWithMappings(
    row => new User { Name = row.GetValue<String>(1), Age = row.GetValue<Int32>(2) },
    options => options.HeaderRow = 0);
```

```csharp
// MiniExcel — 同样简洁
var list = await MiniExcel.QueryAsync<User>("report.xlsx");

// NPOI — 无内置支持，需手写列名→属性映射（约 20-30 行）
```

### 5.3 模板填充

```csharp
// ✅ NewLife.Office（ExcelHelper 最简洁模式）— 1 行完成
ExcelHelper.SaveByTemplate("output.xlsx", "template.xlsx", new { Name = "张三", Date = DateTime.Today, Total = 9800m });
```

```csharp
// ✅ NewLife.Office（ExcelTemplate 显式控制）
var tpl = new ExcelTemplate("template.xlsx");
tpl.Fill("output.xlsx", new Dictionary<String, Object>
{
    ["Name"] = "张三", ["Date"] = DateTime.Today, ["Total"] = 9800m
});
```

```csharp
// MiniExcel（模式接近，语法略有差异）
var data = new { Name = "张三", Date = DateTime.Today, Total = 9800m };
await MiniExcel.SaveAsByTemplateAsync("output.xlsx", "template.xlsx", data);

// EPPlus / NPOI 无原生模板填充能力，需自己实现占位符替换逻辑
```

### 5.4 流式读取大文件

```csharp
// ✅ NewLife.Office — IEnumerable 逐行 yield，内存极低
using var reader = new ExcelReader("bigfile.xlsx");
foreach (var row in reader.ReadRows())   // row: Object?[]
    Process(row);
```

```csharp
// MiniExcel — 同样支持流式，API 风格相似
await foreach (var row in MiniExcel.QueryAsync("bigfile.xlsx"))
    Process(row);

// EPPlus / ClosedXML — 全量加载到内存，不适合超大文件
```

### 5.5 单元格样式设置

```csharp
// ✅ NewLife.Office — 值对象风格，一次构建跨行复用
var style = new CellStyle
{
    Bold = true, FontColor = "FF0000", Background = "FFFF00",
    HorizontalAlignment = HorizontalAlignment.Center,
    Border = CellBorderStyle.Thin
};
writer.WriteRow(null, new Object[] { "总计", 9800m }, style);
```

```csharp
// EPPlus — 每个属性单独设置，代码量多但可精细控制
var cell = sheet.Cells["A1"];
cell.Style.Font.Bold = true;
cell.Style.Font.Color.SetColor(Color.Red);
cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
cell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
```

```csharp
// NPOI — 需在 workbook 级别预先创建样式对象，API 源自 Java 风格
var font = workbook.CreateFont();
font.IsBold = true;
font.Color = IndexedColors.Red.Index;
var cellStyle = workbook.CreateCellStyle();
cellStyle.SetFont(font);
cellStyle.FillForegroundColor = IndexedColors.Yellow.Index;
cellStyle.FillPattern = FillPattern.SolidForeground;
cellStyle.Alignment = HorizontalAlignment.Center;
cell.CellStyle = cellStyle;
```

> **小结**：NewLife.Office 提供 ExcelHelper 静态入口（单行即可完成导入/导出/模板，媲美 MiniExcel 最简洁用法）；同时内置完整样式、图表、高级特性支持，远超 MiniExcel；以值对象风格的 API 比 EPPlus/NPOI 减少 60–80% 代码量。

---

## 6. 差异化定位

- **vs EPPlus**：完全免费，无商业许可限制，框架兼容性更好（支持 net45）
- **vs NPOI**：API 更简洁现代，不引入 Java 风格，2.8+ 版本商业使用存在许可风险
- **vs ClosedXML**：无 Open XML SDK 依赖，支持 net45，包体积更小
- **vs MiniExcel**：功能远超（样式/图表/公式/图片/页面设置等全面支持）

### 高保真 xlsx 完成计划（M18-M25）

| 模块 | 功能 | 优先级 | 状态 |
|------|------|--------|------|
| M18 | 命名范围 (Defined Names) | 🔴高 | ✅ 完成 |
| M19 | 结构化 Table 元素 | 🔴高 | ✅ 完成 |
| M20 | 富文本/渐变填充/四边独立边框 | 🟡中 | ✅ 完成 |
| M21 | 图表集成到 Writer/Reader | 🟡中 | ✅ 完成 |
| M22 | 条件格式图标集/自定义公式 | 🟢低 | ✅ 完成 |
| M23 | 文本旋转/缩进/分组大纲 | 🟢低 | ✅ 完成 |
| M24 | 标签颜色/工作簿保护/calcPr | 🟢低 | ✅ 完成 |
| M25 | xls 基础样式 (BiffWriter) | 🟢低 | ✅ 完成 |

---

← 返回 [竞品分析报告.md](竞品分析报告.md)
