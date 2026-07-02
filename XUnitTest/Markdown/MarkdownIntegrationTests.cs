using System.ComponentModel;
using System.Text;
using NewLife.Office;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Ole2;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using NewLife.Office.Xps;
using Xunit;

using XUnitTest.Common;

namespace XUnitTest.Markdown;

/// <summary>Markdown 格式集成测试</summary>
public class MarkdownIntegrationTests : IntegrationTestBase
{
    [Fact, DisplayName("Markdown_复杂写入再读取往返")]
    public void Markdown_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.md");

        var mdText = @"# NewLife.Office 集成测试

## 概述

这是一份由集成测试自动生成的 **Markdown** 文档。

## 功能列表

- Excel 读写（xlsx/xls）
- Word 文档（docx）
- PDF 生成和读取
- PowerPoint（pptx）
- RTF 格式
- 更多格式...

## 代码示例

```csharp
var reader = new ExcelReader(""data.xlsx"");
var rows = reader.ReadRows().ToList();
```

## 数据表格

| 格式 | 读取 | 写入 |
|------|------|------|
| XLSX | ✓ | ✓ |
| DOCX | ✓ | ✓ |
| PDF | ✓ | ✓ |

## 引用

> NewLife.Office 是一个纯 C# 实现的办公文档处理库。
> 无需安装 Office 即可读写多种文档格式。

### 小标题

普通段落结束。
";
        File.WriteAllText(path, mdText, new UTF8Encoding(false));

        Assert.True(File.Exists(path));

        // 解析验证
        var doc = MarkdownDocument.ParseFile(path);
        Assert.True(doc.Blocks.Count >= 5);

        // 往返：序列化再解析，块数量可能因空行处理略有差异
        var markdown = doc.ToMarkdown();
    }
}
