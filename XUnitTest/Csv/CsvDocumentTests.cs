using System.ComponentModel;
using System.Text;
using NewLife.Office.Csv;
using Xunit;

namespace XUnitTest.Csv;

/// <summary>CSV 文档模型测试：RFC4180 解析、编码/分隔符识别与下游模型转换</summary>
public class CsvDocumentTests
{
    static CsvDocumentTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    #region 工具
    private static String TempFile(String name, Byte[] bytes)
    {
        var dir = Path.Combine(Path.GetTempPath(), "nlo_csv_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, name);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    private static String TempFile(String name, String text) => TempFile(name, new UTF8Encoding(false).GetBytes(text));
    #endregion

    [Fact, DisplayName("CSV 读取：引号字段含逗号、换行与转义双引号")]
    public void Read_HandlesQuotedFields()
    {
        var path = TempFile("quote.csv", "name,note\r\n\"张,三\",\"第一行\r\n第二行\"\r\n李四,\"他说\"\"你好\"\"\"\r\n");

        var csv = CsvDocument.ReadFile(path);

        Assert.Equal(3, csv.RowCount);
        Assert.Equal("张,三", csv.Rows[1][0]);
        Assert.Equal("第一行\r\n第二行", csv.Rows[1][1]);
        Assert.Equal("他说\"你好\"", csv.Rows[2][1]);
    }

    [Fact, DisplayName("CSV 读取：UTF-8 BOM 不污染首列表头")]
    public void Read_StripsUtf8Bom()
    {
        var bytes = new Byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes("地区,金额\r\n华东,100\r\n")).ToArray();
        var path = TempFile("bom.csv", bytes);

        var csv = CsvDocument.ReadFile(path);

        Assert.Equal("地区", csv.Rows[0][0]);
        Assert.Equal(',', csv.Separator);
    }

    [Fact, DisplayName("CSV 读取：中文 GBK 文件自动识别编码")]
    public void Read_DetectsGbk()
    {
        var gbk = Encoding.GetEncoding(936);
        var path = TempFile("gbk.csv", gbk.GetBytes("地区,金额\r\n华东,100\r\n华南,200\r\n"));

        var csv = CsvDocument.ReadFile(path);

        Assert.Equal("地区", csv.Rows[0][0]);
        Assert.Equal("华东", csv.Rows[1][0]);
        Assert.Equal(3, csv.RowCount);
    }

    [Fact, DisplayName("CSV 读取：分号分隔自动识别")]
    public void Read_DetectsSemicolon()
    {
        var path = TempFile("semi.csv", "地区;金额\r\n华东;100\r\n华南;200\r\n");

        var csv = CsvDocument.ReadFile(path);

        Assert.Equal(';', csv.Separator);
        Assert.Equal(2, csv.Rows[0].Length);
        Assert.Equal("华东", csv.Rows[1][0]);
    }

    [Fact, DisplayName("CSV 提取：Markdown 表格首行为表头")]
    public void ExtractMarkdown_UsesFirstRowAsHeader()
    {
        var path = TempFile("md.csv", "地区,金额\r\n华东,100\r\n");

        var markdown = CsvDocument.ReadFile(path).ExtractMarkdown();

        Assert.Contains("| 地区 | 金额 |", markdown);
        Assert.Contains("| 华东 | 100 |", markdown);
    }

    [Fact, DisplayName("CSV 转换：工作簿模型保留原始行")]
    public void ToExcelDocument_KeepsRows()
    {
        var path = TempFile("book.csv", "地区,金额\r\n华东,100\r\n");

        var document = CsvDocument.ReadFile(path).ToExcelDocument("销售");

        var sheet = Assert.Single(document.Worksheets);
        Assert.Equal("销售", sheet.Name);
        Assert.Equal(2, sheet.Rows.Count);
        Assert.Equal("华东", sheet.Rows[1][0]);
    }

    [Fact, DisplayName("CSV 读取：空文件返回零行")]
    public void Read_EmptyFile_ReturnsNoRows()
    {
        var path = TempFile("empty.csv", "");

        var csv = CsvDocument.ReadFile(path);

        Assert.Equal(0, csv.RowCount);
        Assert.Null(csv.ExtractText());
        Assert.Null(csv.ExtractMarkdown());
    }
}
