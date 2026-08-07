using System.ComponentModel;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>ExcelReader 真实世界健壮性测试 — 1904日期/错误单元格/ISO日期/富文本共享字符串/隐藏行列/缺失部件容错</summary>
public class ExcelReaderRobustnessTests
{
    #region 辅助
    /// <summary>构造最小 xlsx（可注入 workbookPr/共享字符串/样式）</summary>
    private static MemoryStream BuildExcel(String sheetXml, String? sharedStrings = null, String? styles = null, String workbookPr = "", Boolean includeShared = true, Boolean includeStyles = true)
    {
        var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, true, Encoding.UTF8))
        {
            // [Content_Types].xml
            using (var sw = new StreamWriter(za.CreateEntry("[Content_Types].xml").Open(), Encoding.UTF8))
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
            // workbook.xml（可带 workbookPr）
            using (var sw = new StreamWriter(za.CreateEntry("xl/workbook.xml").Open(), Encoding.UTF8))
                sw.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{workbookPr}<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/></sheets></workbook>");
            // styles.xml
            if (includeStyles)
            {
                using var sw = new StreamWriter(za.CreateEntry("xl/styles.xml").Open(), Encoding.UTF8);
                sw.Write(styles ?? "<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs></styleSheet>");
            }
            // sharedStrings.xml
            if (includeShared)
            {
                using var sw = new StreamWriter(za.CreateEntry("xl/sharedStrings.xml").Open(), Encoding.UTF8);
                sw.Write(sharedStrings ?? "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"><si><t>Str</t></si></sst>");
            }
            // sheet
            using (var sw = new StreamWriter(za.CreateEntry("xl/worksheets/sheet1.xml").Open(), Encoding.UTF8))
                sw.Write(sheetXml);
        }
        ms.Position = 0;
        return ms;
    }

    private static readonly String DateStyle = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs></styleSheet>";
    #endregion

    #region 1904 日期系统
    [Fact]
    [DisplayName("Excel读—1904日期系统：序列值偏移1462天")]
    public void Date1904_System()
    {
        // 1904 系统下序列 0 = 1904-01-01，43831 = 1904-01-01 + 43831 天
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" s=\"1\"><v>0</v></c><c r=\"B1\" s=\"1\"><v>43831</v></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, styles: DateStyle, workbookPr: "<workbookPr date1904=\"1\"/>");
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal(new DateTime(1904, 1, 1), (DateTime)row[0]!);
        Assert.Equal(new DateTime(1904, 1, 1).AddDays(43831), (DateTime)row[1]!);
    }

    [Fact]
    [DisplayName("Excel读—1900系统序列边界：1/59/60/61")]
    public void Date1900_SerialBoundaries()
    {
        // 1900 系统：1=1900-01-01, 59=1900-02-28, 60=虚构闰日(钳制为1900-02-28), 61=1900-03-01
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\">" +
            "<c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"B1\" s=\"1\"><v>59</v></c><c r=\"C1\" s=\"1\"><v>60</v></c><c r=\"D1\" s=\"1\"><v>61</v></c>" +
            "</row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, styles: DateStyle);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal(new DateTime(1900, 1, 1), (DateTime)row[0]!);
        Assert.Equal(new DateTime(1900, 2, 28), (DateTime)row[1]!);
        Assert.Equal(new DateTime(1900, 2, 28), (DateTime)row[2]!); // 虚构闰日钳制
        Assert.Equal(new DateTime(1900, 3, 1), (DateTime)row[3]!);
    }

    [Fact]
    [DisplayName("Excel读—日期含时间精度")]
    public void Date_WithTime()
    {
        // 43831.5 = 2020-01-01 12:00（1900 系统）
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" s=\"1\"><v>43831.5</v></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, styles: DateStyle);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal(new DateTime(2020, 1, 1, 12, 0, 0), (DateTime)row[0]!);
    }
    #endregion

    #region t=e 错误单元格 / t=d ISO 日期
    [Fact]
    [DisplayName("Excel读—错误单元格t=e保持错误字符串")]
    public void ErrorCell()
    {
        // 带日期样式 s=1 的错误单元格也不应被转换
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\">" +
            "<c r=\"A1\" t=\"e\"><v>#DIV/0!</v></c>" +
            "<c r=\"B1\" t=\"e\" s=\"1\"><v>#N/A</v></c>" +
            "<c r=\"C1\" t=\"e\"><v>#REF!</v></c>" +
            "</row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, styles: DateStyle);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal("#DIV/0!", row[0]);
        Assert.Equal("#N/A", row[1]);
        Assert.Equal("#REF!", row[2]);
    }

    [Fact]
    [DisplayName("Excel读—ISO日期单元格t=d转为DateTime")]
    public void IsoDateCell()
    {
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\">" +
            "<c r=\"A1\" t=\"d\"><v>2020-01-15T00:00:00</v></c>" +
            "<c r=\"B1\" t=\"d\" s=\"1\"><v>2021-06-01T08:30:00</v></c>" +
            "</row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, styles: DateStyle);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal(new DateTime(2020, 1, 15), (DateTime)row[0]!);
        Assert.Equal(new DateTime(2021, 6, 1, 8, 30, 0), (DateTime)row[1]!);
    }
    #endregion

    #region 富文本共享字符串
    [Fact]
    [DisplayName("Excel读—富文本共享字符串拼接run文本并跳过音标")]
    public void RichTextSharedString()
    {
        // si 含多个 <r> run + 音标 rPh，应拼接 run 文本且跳过音标内容
        var shared = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">" +
            "<si><r><rPr><b/></rPr><t>你好</t></r><r><t>世界</t></r><rPh sb=\"0\" eb=\"1\"><t>ㄋㄧˇ</t></rPh></si></sst>";
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, sharedStrings: shared);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal("你好世界", row[0]);
    }

    [Fact]
    [DisplayName("Excel读—共享字符串xml:space=preserve保留空格")]
    public void SharedString_PreserveSpace()
    {
        var shared = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">" +
            "<si><t xml:space=\"preserve\">  前导空格  </t></si><si><t>普通</t></si></sst>";
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, sharedStrings: shared);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal("  前导空格  ", row[0]);
        Assert.Equal("普通", row[1]);
    }
    #endregion

    #region 隐藏行列
    [Fact]
    [DisplayName("Excel读—隐藏行/隐藏列读取")]
    public void HiddenRowsAndColumns()
    {
        // 第2行 hidden，B列 hidden
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<cols><col min=\"2\" max=\"2\" hidden=\"1\"/></cols>" +
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row><row r=\"2\" hidden=\"1\"><c r=\"A2\"><v>3</v></c></row><row r=\"3\"><c r=\"A3\"><v>4</v></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var hiddenRows = reader.ReadHiddenRows("Sheet1");
        var hiddenCols = reader.ReadHiddenColumns("Sheet1");
        Assert.Equal(new[] { 1 }, hiddenRows); // 0基行1 = 第2行
        Assert.Equal(new[] { 1 }, hiddenCols); // 0基列1 = B列
    }

    [Fact]
    [DisplayName("Excel读—隐藏行列快照ReadWorksheet")]
    public void HiddenRows_InReadWorksheet()
    {
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<cols><col min=\"3\" max=\"3\" hidden=\"true\"/></cols>" +
            "<sheetData><row r=\"1\" hidden=\"true\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var sd = reader.ReadWorksheet("Sheet1");
        Assert.Contains(0, sd.HiddenRows);
        Assert.Contains(2, sd.HiddenColumns);
    }
    #endregion

    #region 主题色解析
    [Fact]
    [DisplayName("Excel读—主题色theme:N解析为实际RGB")]
    public void ResolveThemeColor_ResolvesRgb()
    {
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet);
        // 向 zip 追加 theme1.xml（clrScheme：dk1=FFFFFF, lt1=000000, accent1=4472C4）
        using (var za = new ZipArchive(ms, ZipArchiveMode.Update, true, Encoding.UTF8))
        {
            var entry = za.CreateEntry("xl/theme/theme1.xml");
            using var sw = new StreamWriter(entry.Open(), Encoding.UTF8);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:clrScheme name=\"Office\">" +
                "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
                "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
                "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>" +
                "<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>" +
                "<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1>" +
                "<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>" +
                "<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>" +
                "<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>" +
                "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>" +
                "<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>" +
                "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>" +
                "<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>" +
                "</a:clrScheme></a:themeElements></a:theme>");
        }
        ms.Position = 0;

        using var reader = new ExcelReader(ms, Encoding.UTF8);
        Assert.Equal("000000", reader.ResolveThemeColor("theme:0"));  // dk1（sysClr lastClr）
        Assert.Equal("FFFFFF", reader.ResolveThemeColor("theme:1"));  // lt1
        Assert.Equal("44546A", reader.ResolveThemeColor("theme:2"));  // dk2
        Assert.Equal("4472C4", reader.ResolveThemeColor("theme:4"));  // accent1
        Assert.Equal("954F72", reader.ResolveThemeColor("theme:11")); // folHlink
        Assert.Equal("FF0000", reader.ResolveThemeColor("FF0000"));   // 非主题色原样返回
    }
    #endregion

    #region 隐藏行列写入与往返
    [Fact]
    [DisplayName("Excel写—SetRowHidden/SetColumnHidden生成hidden属性")]
    public void Write_HiddenRowsAndCols()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { "A", "B", "C" });
            w.WriteRow("Sheet1", new Object?[] { "D", "E", "F" });
            w.SetRowHidden("Sheet1", 2);
            w.SetColumnHidden("Sheet1", 1);
            w.Save();
        }

        using var reader = new ExcelReader(path);
        var hiddenRows = reader.ReadHiddenRows("Sheet1");
        var hiddenCols = reader.ReadHiddenColumns("Sheet1");
        Assert.Equal(new[] { 1 }, hiddenRows); // 0基行1 = 第2行
        Assert.Equal(new[] { 1 }, hiddenCols); // 0基列1 = B列
    }

    [Fact]
    [DisplayName("Excel往返—隐藏行列WriteExcel回写")]
    public void RoundTrip_HiddenRowsAndCols()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { "A", "B", "C" });
            w.WriteRow("Sheet1", new Object?[] { "D", "E", "F" });
            w.SetRowHidden("Sheet1", 2);
            w.SetColumnHidden("Sheet1", 1);
            w.Save();
        }

        // 读快照 → 写回 → 再读
        ExcelDocument data;
        using (var reader = new ExcelReader(path))
        {
            data = reader.ReadExcel();
        }
        Assert.Contains(1, data.Worksheets[0].HiddenRows);
        Assert.Contains(1, data.Worksheets[0].HiddenColumns);

        var path2 = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w2 = new ExcelWriter(path2))
        {
            w2.WriteExcel(data);
            w2.Save();
        }

        using var reader2 = new ExcelReader(path2);
        var sd = reader2.ReadWorksheet("Sheet1");
        Assert.Equal(new[] { 1 }, sd.HiddenRows);
        Assert.Equal(new[] { 1 }, sd.HiddenColumns);
    }
    #endregion

    #region 批注增强
    [Fact]
    [DisplayName("Excel批注—尺寸与可见性写入读取")]
    public void Comment_SizeVisibility()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { "A", "B" });
            w.AddComment("Sheet1", 1, 0, "隐藏批注", "作者A", 200, 80, false);
            w.AddComment("Sheet1", 1, 1, "可见批注", "作者B", 150, 100, true);
            w.Save();
        }

        using var reader = new ExcelReader(path);
        var comments = reader.ReadComments("Sheet1");
        Assert.Equal(2, comments.Count);

        var c1 = comments[(0, 0)];
        Assert.Equal("隐藏批注", c1.Text);
        Assert.Equal("作者A", c1.Author);
        Assert.False(c1.Visible);
        Assert.True(Math.Abs(c1.Width - 200) < 0.5, $"宽度={c1.Width}");
        Assert.True(Math.Abs(c1.Height - 80) < 0.5, $"高度={c1.Height}");

        var c2 = comments[(0, 1)];
        Assert.True(c2.Visible);
        Assert.True(Math.Abs(c2.Width - 150) < 0.5, $"宽度={c2.Width}");
        Assert.True(Math.Abs(c2.Height - 100) < 0.5, $"高度={c2.Height}");
    }

    [Fact]
    [DisplayName("Excel往返—批注尺寸可见性WriteExcel回写")]
    public void Comment_RoundtripWriteExcel()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { "A" });
            w.AddComment("Sheet1", 1, 0, "测试批注", "作者", 180, 90, true);
            w.Save();
        }

        ExcelDocument data;
        using (var reader = new ExcelReader(path))
        {
            data = reader.ReadExcel();
        }
        Assert.True(data.Worksheets[0].Comments[(0, 0)].Visible);

        var path2 = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w2 = new ExcelWriter(path2))
        {
            w2.WriteExcel(data);
            w2.Save();
        }

        using var reader2 = new ExcelReader(path2);
        var cm = reader2.ReadComments("Sheet1")[(0, 0)];
        Assert.Equal("测试批注", cm.Text);
        Assert.True(cm.Visible);
        Assert.True(Math.Abs(cm.Width - 180) < 0.5);
        Assert.True(Math.Abs(cm.Height - 90) < 0.5);
    }
    #endregion

    #region 大文件流式性能
    [Fact]
    [DisplayName("Excel性能—10万行流式读取时间")]
    public void Performance_100kRows()
    {
        // 构造 10万行 × 8 列的 xlsx（共享字符串池化，避免写放大）
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        var pool = new String[8];
        for (var i = 0; i < pool.Length; i++) pool[i] = "类别" + i;

        var swWrite = Stopwatch.StartNew();
        using (var w = new ExcelWriter(path))
        {
            w.AutoFitColumnWidth = false;
            w.WriteHeader("Sheet1", new[] { "编号", "比率", "类别A", "类别B", "类别C", "类别D", "类别E", "类别F" });
            for (var r = 0; r < 100_000; r++)
            {
                var row = new Object?[8];
                row[0] = r + 1;
                row[1] = r * 0.5;
                for (var c = 2; c < 8; c++) row[c] = pool[(r + c) % pool.Length];
                w.WriteRow("Sheet1", row);
            }
            w.Save();
        }
        swWrite.Stop();

        // 流式读取
        var swRead = Stopwatch.StartNew();
        var count = 0;
        using (var reader = new ExcelReader(path))
        {
            foreach (var row in reader.ReadRows("Sheet1"))
            {
                Assert.Equal(8, row.Length);
                count++;
            }
        }
        swRead.Stop();

        Assert.Equal(100_001, count); // 表头 + 10万数据行
        Console.WriteLine($"10万行 xlsx：写入 {swWrite.ElapsedMilliseconds}ms，流式读取 {swRead.ElapsedMilliseconds}ms，共 {count} 行");

        // 读取时间宽松阈值（避免 CI 抖动）：10万行 × 8 列应 < 5 秒
        Assert.True(swRead.ElapsedMilliseconds < 5000, $"10万行流式读取耗时 {swRead.ElapsedMilliseconds}ms 超过 5 秒阈值");
    }
    #endregion

    #region 打印区域与分页符
    [Fact]
    [DisplayName("Excel往返—打印区域与分页符WriteExcel回写")]
    public void PrintArea_PageBreaks_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { "A", "B", "C" });
            w.WriteRow("Sheet1", new Object?[] { "D", "E", "F" });
            w.WriteRow("Sheet1", new Object?[] { "G", "H", "I" });
            w.SetPrintArea("Sheet1", "A1:C3");
            w.SetPageBreak("Sheet1", 2);
            w.SetColumnPageBreak("Sheet1", 2);
            w.Save();
        }

        ExcelDocument data;
        using (var reader = new ExcelReader(path))
        {
            data = reader.ReadExcel();
        }
        var sd = data.Worksheets[0];
        Assert.NotNull(sd.PrintArea);
        Assert.Contains(2, sd.RowPageBreaks);
        Assert.Contains(2, sd.ColumnPageBreaks);

        var path2 = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w2 = new ExcelWriter(path2))
        {
            w2.WriteExcel(data);
            w2.Save();
        }

        using var reader2 = new ExcelReader(path2);
        var sd2 = reader2.ReadWorksheet("Sheet1");
        Assert.Equal(sd.PrintArea, sd2.PrintArea);
        Assert.Equal(sd.RowPageBreaks, sd2.RowPageBreaks);
        Assert.Equal(sd.ColumnPageBreaks, sd2.ColumnPageBreaks);
    }
    #endregion

    #region 显示设置往返
    [Fact]
    [DisplayName("Excel往返—网格线/缩放/默认行高WriteExcel回写")]
    public void DisplaySettings_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { "A", "B" });
            w.SetGridlines("Sheet1", false);
            w.SetZoomScale("Sheet1", 150);
            w.SetDefaultRowHeight("Sheet1", 18);
            w.Save();
        }

        ExcelDocument data;
        using (var reader = new ExcelReader(path))
        {
            data = reader.ReadExcel();
        }
        var sd = data.Worksheets[0];
        Assert.False(sd.Gridlines);
        Assert.Equal(150, sd.ZoomScale);
        Assert.Equal(18, sd.DefaultRowHeight);

        var path2 = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w2 = new ExcelWriter(path2))
        {
            w2.WriteExcel(data);
            w2.Save();
        }

        using var reader2 = new ExcelReader(path2);
        var sd2 = reader2.ReadWorksheet("Sheet1");
        Assert.False(sd2.Gridlines);
        Assert.Equal(150, sd2.ZoomScale);
        Assert.Equal(18, sd2.DefaultRowHeight);
    }
    #endregion

    #region 条件格式 dxf 颜色往返
    [Fact]
    [DisplayName("Excel往返—条件格式填充色dxf写入读取（多规则dxfId分配）")]
    public void ConditionalFormat_Color_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { 100 });
            w.WriteRow("Sheet1", new Object?[] { 0 });
            w.AddConditionalFormat("Sheet1", "A1:A2", ConditionalFormatValues.GreaterThan, "0", "92D050");
            w.AddConditionalFormat("Sheet1", "A1:A2", ConditionalFormatValues.LessThan, "50", "FF0000");
            w.AddExpressionConditionalFormat("Sheet1", "A1:A2", "A1>200", "FFFF00");
            w.Save();
        }

        // 直接读取
        using (var reader = new ExcelReader(path))
        {
            var cfs = reader.ReadConditionalFormats("Sheet1").ToList();
            Assert.Equal(3, cfs.Count);
            var gt = cfs.First(c => c.Type == ConditionalFormatValues.GreaterThan);
            Assert.Equal("92D050", gt.Color);
            var lt = cfs.First(c => c.Type == ConditionalFormatValues.LessThan);
            Assert.Equal("FF0000", lt.Color);
            var ex = cfs.First(c => c.Type == ConditionalFormatValues.Expression);
            Assert.Equal("FFFF00", ex.Color);
        }

        // 快照往返后仍保留
        ExcelDocument data;
        using (var reader = new ExcelReader(path))
        {
            data = reader.ReadExcel();
        }
        var path2 = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w2 = new ExcelWriter(path2))
        {
            w2.WriteExcel(data);
            w2.Save();
        }
        using var reader2 = new ExcelReader(path2);
        var cfs2 = reader2.ReadConditionalFormats("Sheet1").ToList();
        var gt2 = cfs2.First(c => c.Type == ConditionalFormatValues.GreaterThan);
        Assert.Equal("92D050", gt2.Color);
        var lt2 = cfs2.First(c => c.Type == ConditionalFormatValues.LessThan);
        Assert.Equal("FF0000", lt2.Color);
        var ex2 = cfs2.First(c => c.Type == ConditionalFormatValues.Expression);
        Assert.Equal("FFFF00", ex2.Color);
    }
    #endregion

    #region 条件格式 dxf 字体样式往返
    [Fact]
    [DisplayName("Excel往返—条件格式dxf字体颜色/加粗/边框写入读取")]
    public void ConditionalFormat_Font_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteRow("Sheet1", new Object?[] { 100 });
            w.WriteRow("Sheet1", new Object?[] { 0 });
            w.AddConditionalFormat("Sheet1", "A1:A2", ConditionalFormatValues.GreaterThan, "0", "92D050", null, "FF0000", true, "000000");
            w.AddConditionalFormat("Sheet1", "A1:A2", ConditionalFormatValues.LessThan, "50", "FF0000", null, "0000FF", false);
            w.AddExpressionConditionalFormat("Sheet1", "A1:A2", "A1>200", "FFFF00");
            w.Save();
        }

        // 直接读取
        using (var reader = new ExcelReader(path))
        {
            var cfs = reader.ReadConditionalFormats("Sheet1").ToList();
            Assert.Equal(3, cfs.Count);
            var gt = cfs.First(c => c.Type == ConditionalFormatValues.GreaterThan);
            Assert.Equal("92D050", gt.Color);
            Assert.Equal("FF0000", gt.FontColor);
            Assert.Equal("000000", gt.BorderColor);
            Assert.True(gt.IsBold);
            var lt = cfs.First(c => c.Type == ConditionalFormatValues.LessThan);
            Assert.Equal("FF0000", lt.Color);
            Assert.Equal("0000FF", lt.FontColor);
            Assert.Null(lt.BorderColor);
            Assert.False(lt.IsBold);
            var ex = cfs.First(c => c.Type == ConditionalFormatValues.Expression);
            Assert.Equal("FFFF00", ex.Color);
            Assert.Null(ex.FontColor);
        }

        // 快照往返后仍保留
        ExcelDocument data;
        using (var reader = new ExcelReader(path))
        {
            data = reader.ReadExcel();
        }
        var path2 = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        using (var w2 = new ExcelWriter(path2))
        {
            w2.WriteExcel(data);
            w2.Save();
        }
        using var reader2 = new ExcelReader(path2);
        var cfs2 = reader2.ReadConditionalFormats("Sheet1").ToList();
        var gt2 = cfs2.First(c => c.Type == ConditionalFormatValues.GreaterThan);
        Assert.Equal("92D050", gt2.Color);
        Assert.Equal("FF0000", gt2.FontColor);
        Assert.Equal("000000", gt2.BorderColor);
        Assert.True(gt2.IsBold);
        var lt2 = cfs2.First(c => c.Type == ConditionalFormatValues.LessThan);
        Assert.Equal("FF0000", lt2.Color);
        Assert.Equal("0000FF", lt2.FontColor);
        Assert.Null(lt2.BorderColor);
        Assert.False(lt2.IsBold);
        var ex2 = cfs2.First(c => c.Type == ConditionalFormatValues.Expression);
        Assert.Equal("FFFF00", ex2.Color);
        Assert.Null(ex2.FontColor);
    }
    #endregion

    #region 缺失部件容错
    [Fact]
    [DisplayName("Excel读—无sharedStrings/styles部件仍可读取")]
    public void MissingParts_StillReads()
    {
        // 无 sharedStrings、无 styles 部件，纯内联数字单元格
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\"><v>123</v></c><c r=\"B1\" t=\"inlineStr\"><is><t>文本</t></is></c></row></sheetData></worksheet>";
        using var ms = BuildExcel(sheet, includeShared: false, includeStyles: false);
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var row = reader.ReadRows().First();
        Assert.Equal(123, row[0]);
        Assert.Equal("文本", row[1]);
    }

    [Fact]
    [DisplayName("Excel读—无workbook.xml时Sheets为空")]
    public void MissingWorkbook_SheetsEmpty()
    {
        // 只有 sheet 部件，没有 workbook.xml —— ReadRows 应返回空而非抛异常
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, true, Encoding.UTF8))
        {
            using (var sw = new StreamWriter(za.CreateEntry("xl/worksheets/sheet1.xml").Open(), Encoding.UTF8))
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>");
        }
        ms.Position = 0;

        using var reader = new ExcelReader(ms, Encoding.UTF8);
        Assert.Empty(reader.ReadRows());
        Assert.Empty(reader.ReadExcel().Worksheets);
    }
    #endregion
}
