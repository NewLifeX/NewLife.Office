using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>Word docx 格式集成测试</summary>
public class WordDocxTests : IntegrationTestBase
{
    [Fact, DisplayName("Word_docx_复杂写入再读取往返")]
    public void Word_Docx_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.docx");

        using (var w = new WordWriter())
        {
            w.DocumentProperties.Title = "集成测试文档";
            w.DocumentProperties.Author = "NewLife Office";

            w.AppendHeading("第一章 概述", 1);
            w.AppendParagraph("这是一份由 NewLife.Office 自动生成的测试文档，用于验证 Word 文件的读写功能。");
            w.AppendParagraph("本文档包含标题、段落、表格、列表等多种元素。");

            w.AppendHeading("第二章 数据表格", 2);
            w.AppendParagraph("以下是一个示例数据表格：");

            var tableData = new[]
            {
                new[] { "产品", "价格", "库存" },
                new[] { "笔记本", "5999", "100" },
                new[] { "手机", "3999", "500" },
                new[] { "平板", "2999", "200" },
            };
            w.AppendTable(tableData);

            w.AppendHeading("第三章 格式化文本", 2);
            w.AppendParagraph("普通段落文本。", WordParagraphStyle.Normal,
                new WordRunProperties { FontSize = 14f, Bold = true });
            w.AppendParagraph("这是另一个段落。");

            w.AppendHeading("附录", 3);
            w.AppendParagraph("文档结束。");

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        // 读取验证
        using var reader = new WordReader(path);
        var paragraphs = reader.ReadParagraphs().ToList();
        Assert.True(paragraphs.Count >= 5);
        Assert.Contains("第一章 概述", paragraphs);
        Assert.Contains("文档结束。", paragraphs);

        // ReadFullText 返回正文文本，Title 属于文档属性不在正文中
        var fullText = reader.ReadFullText();
        Assert.Contains("数据表格", fullText);
        Assert.Contains("格式化文本", fullText);

        // 读取表格
        var tables = reader.ReadTables().ToList();
        Assert.True(tables.Count >= 1);
        Assert.Equal("产品", tables[0][0][0]);
        Assert.Equal("笔记本", tables[0][1][0]);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<WordReader>(factoryReader);
        (factoryReader as IDisposable)?.Dispose();
    }

    [Fact, DisplayName("Word_docx转PDF")]
    public void Word_Docx_To_Pdf()
    {
        var docxPath = Path.Combine(OutputDir, "convert_word.docx");
        var pdfPath = Path.Combine(OutputDir, "converted_from_word.pdf");

        using (var w = new WordWriter())
        {
            w.AppendHeading("Word转PDF测试", 1);
            w.AppendParagraph("这段文字应该出现在PDF中。");
            w.AppendParagraph("支持多段落转换。");
            w.Save(docxPath);
        }

        // 转换
        var converter = new WordPdfConverter();
        converter.ConvertToFile(docxPath, pdfPath);

        Assert.True(File.Exists(pdfPath));

        // 验证 PDF
        using var pdfReader = new PdfReader(pdfPath);
        var text = pdfReader.ExtractText();
        Assert.Contains("Word", text);
        Assert.True(pdfReader.GetPageCount() >= 1);
    }
}
