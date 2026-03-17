using System.ComponentModel;
using System.Text;
using NewLife.Office;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>Markdown 格式集成测试</summary>
public class MarkdownTests : IntegrationTestBase
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
        Assert.Contains("# NewLife.Office 集成测试", markdown);
        Assert.Contains("Excel 读写", markdown);

        var doc2 = MarkdownDocument.Parse(markdown);
        Assert.True(doc2.Blocks.Count >= 5);

        // 转 HTML
        var html = doc.ToHtml();
        Assert.Contains("<h1", html);
        Assert.Contains("<table", html);
        Assert.Contains("csharp", html);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<MarkdownDocument>(factoryReader);
    }

    [Fact, DisplayName("Markdown转Word")]
    public void Markdown_To_Word()
    {
        var mdPath = Path.Combine(OutputDir, "convert_md.md");
        var docxPath = Path.Combine(OutputDir, "converted_from_md.docx");

        File.WriteAllText(mdPath, "# 标题\n\n这是正文段落。\n\n## 二级标题\n\n另一个段落。\n", new UTF8Encoding(false));

        var doc = MarkdownDocument.ParseFile(mdPath);
        var converter = new MarkdownWordConverter();
        var bytes = converter.ToBytes(doc);
        File.WriteAllBytes(docxPath, bytes);

        Assert.True(File.Exists(docxPath));

        using var reader = new WordReader(docxPath);
        var paras = reader.ReadParagraphs().ToList();
        Assert.Contains("标题", paras);
        Assert.Contains("这是正文段落。", paras);
    }

    [Fact, DisplayName("Markdown转PDF")]
    public void Markdown_To_Pdf()
    {
        var mdPath = Path.Combine(OutputDir, "convert_md_pdf.md");
        var pdfPath = Path.Combine(OutputDir, "converted_from_md.pdf");

        File.WriteAllText(mdPath, "# PDF标题\n\n正文内容。\n\n| A | B |\n|---|---|\n| 1 | 2 |\n", new UTF8Encoding(false));

        var doc = MarkdownDocument.ParseFile(mdPath);
        var converter = new MarkdownPdfConverter();
        var bytes = converter.ToBytes(doc);
        File.WriteAllBytes(pdfPath, bytes);

        Assert.True(File.Exists(pdfPath));
        Assert.True(bytes.Length > 0);

        // 编码提供程序已在基类中注册
        using var pdfReader = new PdfReader(pdfPath);
        Assert.True(pdfReader.GetPageCount() >= 1);
    }
}
