using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>PDF 格式集成测试</summary>
public class PdfTests : IntegrationTestBase
{
    [Fact, DisplayName("PDF_复杂写入再读取")]
    public void Pdf_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.pdf");

        using (var w = new PdfFluentDocument())
        {
            w.Title = "PDF集成测试";
            w.Author = "NewLife Office";

            w.AddText("PDF集成测试文档", 24f);
            w.AddEmptyLine(12f);
            w.AddText("这是一份由 NewLife.Office 自动生成的 PDF 测试文档。", 12f);
            w.AddText("本文档包含多种元素：文本、表格等。", 12f);
            w.AddEmptyLine(12f);
            w.AddText("第一章 数据表格", 18f);
            w.AddEmptyLine(8f);

            var tableData = new List<String[]>
            {
                new[] { "姓名", "年龄", "城市" },
                new[] { "张三", "28", "北京" },
                new[] { "李四", "35", "上海" },
                new[] { "王五", "42", "广州" },
            };
            w.AddTable(tableData, firstRowHeader: true);

            w.AddEmptyLine(12f);
            w.AddText("第二章 其它内容", 18f);
            w.AddText("这是第二章的内容。", 12f);
            w.AddText("分隔线以上是正文内容。", 10f);

            // 第二页
            w.PageBreak();
            w.AddText("第二页内容", 16f);
            w.AddText("多页 PDF 测试。", 12f);

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        // 读取验证（需要 CodePagesEncodingProvider，已在基类静态构造中注册）
        using var reader = new PdfReader(path);
        Assert.Equal(2, reader.GetPageCount());

        var text = reader.ExtractText();
        Assert.Contains("PDF", text);

        var meta = reader.ReadMetadata();
        Assert.Equal(2, meta.PageCount);
        Assert.NotNull(meta.PdfVersion);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<PdfReader>(factoryReader);
        (factoryReader as IDisposable)?.Dispose();
    }
}
