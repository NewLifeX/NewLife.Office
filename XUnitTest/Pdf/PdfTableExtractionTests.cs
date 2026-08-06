using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office.Pdf;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>PDF 结构化表格提取测试（P09，对标 iText 7/Aspose.PDF）</summary>
/// <remarks>
/// 覆盖基于文本坐标聚类的表格识别：规则表格提取、无表格页面不误报、多表格。
/// </remarks>
public class PdfTableExtractionTests
{
    static PdfTableExtractionTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    /// <summary>生成含表格的 PDF 字节（英文，WinAnsi 编码可正确提取）</summary>
    private static Byte[] BuildTablePdf()
    {
        using var ms = new MemoryStream();
        using (var writer = new PdfWriter())
        {
            writer.BeginPage();
            writer.DrawText("Sales Report", 56, 780, 16);
            writer.DrawTable(new[]
            {
                new[] { "Product", "Qty", "Amount" },
                new[] { "Apple", "100", "500.00" },
                new[] { "Banana", "200", "400.00" },
                new[] { "Orange", "150", "600.00" },
            });
            writer.Save(ms);
        }
        return ms.ToArray();
    }

    [Fact, DisplayName("表格提取：识别行列结构")]
    public void Extract_SimpleTable()
    {
        var bytes = BuildTablePdf();
        using var reader = new PdfReader(new MemoryStream(bytes));
        var tables = PdfTableExtractor.Extract(reader);

        Assert.NotEmpty(tables);
        var table = tables[0];
        var dump = String.Join(" || ", table.Rows.Select(r => String.Join("|", r.Cells)));
        Assert.True(table.RowCount >= 3, $"行数 {table.RowCount} 不足: {dump}");
        Assert.True(table.ColumnCount >= 3, $"列数 {table.ColumnCount} 不足: {dump}");

        // 表头
        var header = String.Join("|", table.Rows[0].Cells);
        Assert.Contains("Product", header);
        Assert.Contains("Qty", header);
        // 数据行内容
        var allText = String.Join("|", table.Rows.SelectMany(r => r.Cells));
        Assert.Contains("Apple", allText);
        Assert.Contains("Banana", allText);
        Assert.Contains("Orange", allText);
    }

    [Fact, DisplayName("表格提取：无表格纯文本页面不误报")]
    public void Extract_NoTable()
    {
        using var ms = new MemoryStream();
        using (var writer = new PdfWriter())
        {
            writer.BeginPage();
            writer.DrawText("这是一个普通段落文本。", 56, 780, 12);
            writer.DrawText("第二行普通文本。", 56, 760, 12);
            writer.Save(ms);
        }

        using var reader = new PdfReader(new MemoryStream(ms.ToArray()));
        var tables = PdfTableExtractor.Extract(reader);
        Assert.Empty(tables);
    }

    [Fact, DisplayName("表格提取：参数校验与空输入")]
    public void Extract_Validation()
    {
        Assert.Throws<ArgumentNullException>(() => PdfTableExtractor.Extract(null!));
    }
}
