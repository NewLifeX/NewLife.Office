using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>XLSX 输出 OOXML 规范校验（OpenXmlValidator 必须 0 错误）</summary>
/// <remarks>
/// Excel 对包结构比 Word/PPT 更严格：Content-Type 不匹配、子元素错位会直接导致"文件已损坏"而无法打开。
/// 本类把历史缺陷固化为用例（图表 drawing 内容类型、图表关系挂载点、条件格式/分页符顺序、迷你图扩展位置）。
/// </remarks>
public class XlsxSchemaTests
{
    #region 辅助
    /// <summary>校验 xlsx 字节流，返回错误描述列表</summary>
    /// <param name="data">xlsx 包字节</param>
    /// <returns>错误描述列表，空表示完全合法</returns>
    private static List<String> Validate(Byte[] data)
    {
        var list = new List<String>();
        try
        {
            using var ms = new MemoryStream(data);
            using var doc = SpreadsheetDocument.Open(ms, false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            foreach (var e in validator.Validate(doc))
                list.Add($"{e.Part?.Uri} | {e.Path?.XPath} | {e.Description}");
        }
        catch (Exception ex) { list.Add("EX " + ex.GetType().Name + ": " + ex.Message); }
        return list;
    }

    /// <summary>把校验结果压成文本，便于断言失败时定位</summary>
    /// <param name="errors">错误列表</param>
    /// <returns>多行描述</returns>
    private static String Describe(List<String> errors) => String.Join("\n", errors);

    /// <summary>读取 zip 中的文本部件</summary>
    /// <param name="data">xlsx 包字节</param>
    /// <param name="path">部件路径</param>
    /// <returns>部件文本，不存在返回 null</returns>
    private static String? ReadPart(Byte[] data, String path)
    {
        using var ms = new MemoryStream(data);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var entry = za.GetEntry(path);
        if (entry == null) return null;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    private static Byte[] Build(Action<ExcelWriter> build)
    {
        using var ms = new MemoryStream();
        using (var w = new ExcelWriter(ms))
        {
            build(w);
            w.Save();
        }
        return ms.ToArray();
    }

    private static ExcelChart BarChart() => new()
    {
        Type = "bar",
        Title = "金额",
        Categories = ["甲", "乙"],
        Series = [new ExcelChartSeries { Name = "金额", Data = [1, 2] }],
        AnchorRow = 6,
        AnchorCol = 0,
    };
    #endregion

    #region 用例
    [Fact(DisplayName = "规范-基础表格0校验错误")]
    public void BasicSheet_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.WriteHeader("S1", ["Id", "Name"]);
            w.WriteRows("S1", [new Object?[] { 1, "甲" }]);
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-仅含图表的工作表可被正常打开")]
    public void ChartOnly_DrawingContentType_Openable()
    {
        var data = Build(w =>
        {
            w.WriteHeader("S1", ["Id", "Amount"]);
            w.WriteRows("S1", [new Object?[] { 1, 10 }, new Object?[] { 2, 20 }]);
            w.AddChart("S1", BarChart());
        });

        // 修正前：drawing1.xml 缺 Override，落回 application/xml，Excel 直接判定文件损坏
        var ct = ReadPart(data, "[Content_Types].xml")!;
        Assert.Contains("/xl/drawings/drawing1.xml", ct);
        Assert.Contains("drawing+xml", ct);

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-图表关系挂在drawing而非worksheet")]
    public void ChartRelationship_OnDrawingPart_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.WriteHeader("S1", ["Id", "Amount"]);
            w.WriteRows("S1", [new Object?[] { 1, 10 }]);
            w.AddChart("S1", BarChart());
        });

        var sheetRels = ReadPart(data, "xl/worksheets/_rels/sheet1.xml.rels")!;
        Assert.DoesNotContain("relationships/chart", sheetRels);
        Assert.Contains("relationships/drawing", sheetRels);

        var drawRels = ReadPart(data, "xl/drawings/_rels/drawing1.xml.rels")!;
        Assert.Contains("relationships/chart", drawRels);

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-条件格式与分页符顺序合法")]
    public void ConditionalFormatAndBreaks_ElementOrder_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }, new Object?[] { 3, 4 }]);
            w.AddConditionalFormat("S1", "B2:B3", ConditionalFormatValues.GreaterThan, "1", "FFC7CE");
            w.SetPageBreak("S1", 2);
            w.SetColumnPageBreak("S1", 2);
            w.SetPageSetup("S1", PageOrientation.Landscape);
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var xml = ReadPart(data, "xl/worksheets/sheet1.xml")!;
        var cf = xml.IndexOf("<conditionalFormatting", StringComparison.Ordinal);
        var brk = xml.IndexOf("<rowBreaks", StringComparison.Ordinal);
        var pageSetup = xml.IndexOf("<pageSetup", StringComparison.Ordinal);
        Assert.True(cf >= 0 && pageSetup > cf, "条件格式必须早于 pageSetup");
        Assert.True(brk > pageSetup, "分页符必须晚于 pageSetup");
    }

    [Fact(DisplayName = "规范-迷你图写在extLst扩展内且可回读")]
    public void Sparklines_InExtLst_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B", "C", "D"]);
            w.WriteRows("S1", [new Object?[] { 1, 2, 3, 4 }]);
            w.AddSparklineGroup("S1", "S1!B2:D2", "S1!E2", "line", "FF0000", "00FF00");
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var xml = ReadPart(data, "xl/worksheets/sheet1.xml")!;
        var extLst = xml.IndexOf("<extLst", StringComparison.Ordinal);
        var groups = xml.IndexOf("sparklineGroups", StringComparison.Ordinal);
        Assert.True(extLst >= 0 && groups > extLst, "sparklineGroups 必须位于 extLst 内");

        using var ms = new MemoryStream(data);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var groupsRead = reader.ReadSparklines("S1");
        Assert.Single(groupsRead);
        Assert.Equal("S1!B2:D2", groupsRead[0].DataRange);
        Assert.Contains("S1!E2", groupsRead[0].Sparklines);
    }

    [Fact(DisplayName = "规范-综合特性工作簿0校验错误")]
    public void RichSheet_NoSchemaErrors()
    {        var data = Build(w =>
        {
            w.WriteHeader("S1", ["Id", "Name", "Amount", "Flag"]);
            w.WriteRows("S1", [
                new Object?[] { 1, "甲", 1234.5, true },
                new Object?[] { 2, "乙", 0.25, false },
            ]);
            w.SetColumnWidth("S1", 0, 12);
            w.FreezePane("S1", 1, 1);
            w.SetAutoFilter("S1", "A1:D3");
            w.MergeCell("S1", "A5:C5");
            w.AddHyperlink("S1", 5, 0, "https://example.com", "示例");
            w.AddDropdownValidation("S1", "D2:D3", ["是", "否"]);
            w.AddRangeValidation("S1", "C2:C3", "decimal", "between", "0", "10000");
            w.SetHeaderFooter("S1", "&C表头 &D", "&R第 &P 页");
            w.SetPrintTitleRows("S1", 1, 1);
            w.SetPrintArea("S1", "A1:D10");
            w.AddTable("A1:D3", "T1", null, ["Id", "Name", "Amount", "Flag"]);
            w.AddComment("S1", 1, 1, "备注内容", "作者");
            w.AddIconSetConditionalFormat("S1", "C2:C3", "3Arrows");
            w.AddExpressionConditionalFormat("S1", "C2:C3", "$C2>0", "C6EFCE");
            w.ProtectSheet("S1", "123");
            w.AddChart("S1", BarChart());
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-图片与多工作表0校验错误")]
    public void ImageAndMultiSheet_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.AddImage("S1", 2, 0, [0x89, 0x50, 0x4E, 0x47], "png", 80, 60);
            w.WriteHeader("S2", ["X", "Y"]);
            w.WriteRows("S2", [new Object?[] { 3, 4 }]);
            w.AddChart("S2", BarChart());
            w.AddTable("S2", "A1:B2", "T2", null, ["X", "Y"]);
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-线程化批注与切片器0校验错误")]
    public void ThreadedCommentAndSlicer_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }, new Object?[] { 3, 4 }]);
            w.AddTable("S1", "A1:B3", "T1", null, ["A", "B"]);
            w.AddSlicer("S1", "T1", "A");
            w.AddThreadedComment("S1", 1, 1, "线程批注", "作者");
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }
    #endregion
}
