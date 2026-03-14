using System.ComponentModel;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

public class ExcelWriterTests
{
    [Fact, DisplayName("单Sheet全类型往返")]
    public void SingleSheet_AllTypes_Roundtrip()
    {
        using var ms = new MemoryStream();
        var writer = new ExcelWriter(ms);

        writer.WriteHeader(null!, new[] { "Name", "Percent", "Date", "DateTime", "Time", "Int", "Long", "Long2", "Long3", "DecFrac", "DecInt", "DoubleFrac", "BoolT", "BoolF", "BigNum", "LeadingZero", "IdCard", "PercentTextFail", "GapTest" });

        var dateOnly = new DateTime(2024, 7, 1);
        var dateTime = new DateTime(2024, 7, 1, 12, 34, 56);
        var time = TimeSpan.FromHours(5) + TimeSpan.FromMinutes(6) + TimeSpan.FromSeconds(7); // 05:06:07
        // 各列说明：Name(string), Percent("12.5%" 成功解析), Date(DateOnly), DateTime(Date+Time), Time(TimeSpan), Int, Long, DecFrac(有小数), DecInt(无小数), DoubleFrac, BoolT, BoolF, BigNum(>12位), LeadingZero(前导0), IdCard(含X), PercentTextFail("abc%"失败分支), GapTest(中间前面留一个null)
        var row = new Object?[]
        {
            "Alice", "12.5%", dateOnly, dateTime, time, 123, 2147483648L, 214748364899L, 2147483648999999L, 123.45m, 456m, 0.125d, true, false,
            "1234567890123", "00123", "12345619900101888X", "abc%", null
        };
        writer.WriteRows(null, new[] { row });

        writer.Save();

        File.WriteAllBytes("ew.xlsx", ms.ToArray());

        // 用 ExcelReader 读取验证类型与数值
        ms.Position = 0;
        var reader = new ExcelReader(ms, Encoding.UTF8);
        var rows = reader.ReadRows().ToList();
        Assert.Equal(2, rows.Count); // header + 1 数据行
        var header = rows[0].Select(e => e + "").ToArray();
        Assert.Equal("Name", header[0]);

        var data = rows[1];
        // Percent => Double 0.125
        Assert.Equal("Alice", data[0]);
        Assert.True(data[1] is Double && Math.Abs((Double)data[1]! - 0.125d) < 1e-9);
        Assert.True(data[2] is DateTime && ((DateTime)data[2]!).Date == dateOnly.Date && ((DateTime)data[2]!).TimeOfDay == TimeSpan.Zero);
        Assert.True(data[3] is DateTime && (DateTime)data[3]! == dateTime);
        Assert.True(data[4] is TimeSpan && (TimeSpan)data[4]! == time);
        Assert.Equal("123", data[5] + "");
        Assert.Equal(2147483648L, (Int64)data[6]!); // long
        Assert.Equal(214748364899L, (Int64)data[7]!); // long
        Assert.Equal("2147483648999999", data[8]!); // long
        Assert.True(data[9] is Decimal or Double); // 小数
        Assert.True(data[10] is Int32 or Int64 or Decimal); // 整数小数样式不变
        Assert.True(data[11] is Decimal or Double);
        Assert.True(data[12] is Boolean && (Boolean)data[12]!);
        Assert.True(data[13] is Boolean && !(Boolean)data[13]!);
        Assert.Equal("1234567890123", data[14]); // 大数字保留文本
        Assert.Equal("00123", data[15]); // 前导0保留
        Assert.Equal("12345619900101888X", data[16]); // 身份证含X
        Assert.Equal("abc%", data[17]); // 百分比解析失败 -> 文本
        // GapTest (最后列前提供 null) => ExcelWriter 跳过，读取时为缺失列应自动补 null
        Assert.Null(data[18]);
    }

    [Fact, DisplayName("多Sheet导出与读取")]
    public void MultiSheet_Export_Read()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader("Users", new[] { "Id", "Name" });
        w.WriteRows("Users", new[] { new Object?[] { 1, "Tom" }, new Object?[] { 2, "Jerry" } });

        w.WriteHeader("Stats", new[] { "Metric", "Value" });
        w.WriteRows("Stats", new[] { new Object?[] { "Count", 2 }, new Object?[] { "Rate", "50%" } });
        w.Save();

        ms.Position = 0;
        var r = new ExcelReader(ms, Encoding.UTF8);
        var users = r.ReadRows("Users").ToList();
        Assert.Equal(3, users.Count); // header + 2
        var stats = r.ReadRows("Stats").ToList();
        Assert.Equal(3, stats.Count);

        // 百分比在第二个sheet中解析为 Double 0.5
        Assert.True(stats[2][1] is Double d && Math.Abs(d - 0.5) < 1e-9);
    }

    [Fact, DisplayName("无字符串时不生成sharedStrings")]
    public void NoSharedStrings_FileStructure()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        // 全数字/日期/时间，不含字符串
        w.WriteRows(null, new[] { new Object?[] { 1, 2.5m, DateTime.Today, TimeSpan.FromMinutes(30) } });
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        Assert.Null(za.GetEntry("xl/sharedStrings.xml")); // 不存在
        Assert.NotNull(za.GetEntry("xl/styles.xml"));
        Assert.NotNull(za.GetEntry("xl/worksheets/sheet1.xml"));
    }

    [Fact, DisplayName("Dispose自动保存文件路径")]
    public void Dispose_AutoSave_File()
    {
        var path = Path.Combine(Path.GetTempPath(), "excelwriter_test_" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var w = new ExcelWriter(path))
            {
                w.WriteHeader(null!, new[] { "A" });
                w.WriteRows(null, new[] { new Object?[] { 123 } });
            } // Dispose 触发保存

            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var r = new ExcelReader(fs, Encoding.UTF8);
            var rows = r.ReadRows().ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("123", rows[1][0] + "");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("空Writer保存生成空Sheet")]
    public void EmptyWriter_Save()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.Save();
        ms.Position = 0;
        var r = new ExcelReader(ms, Encoding.UTF8);
        var list = r.ReadRows().ToList();
        Assert.Empty(list); // 无数据行
    }

    [Fact, DisplayName("Int64使用整数样式避免科学计数")]
    public void Int64_Uses_Integer_Style()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "LongVal" });
        var longVal = 1234567890123456789L; // 19位，超过15位后将改为共享字符串，避免精度与科学计数
        w.WriteRows(null, new[] { new Object?[] { longVal } });
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        var sheet = za.GetEntry("xl/worksheets/sheet1.xml");
        Assert.NotNull(sheet);
        using var sr = new StreamReader(sheet!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        // 现在超过15位的 Int64 以共享字符串方式写入，A1=表头(LongVal)->索引0，A2=长数字->索引1
        Assert.Contains("<c r=\"A2\" t=\"s\"><v>1</v></c>", xml);

        // 同时校验 sharedStrings.xml 中包含该长数字文本
        var sharedEntry = za.GetEntry("xl/sharedStrings.xml");
        Assert.NotNull(sharedEntry);
        using (var ssr = new StreamReader(sharedEntry!.Open(), Encoding.UTF8))
        {
            var sharedXml = ssr.ReadToEnd();
            Assert.Contains("LongVal", sharedXml);
            Assert.Contains(longVal.ToString(), sharedXml);
        }

        // 读取验证：返回为字符串（避免精度丢失），调用方可自行再解析
        ms.Position = 0;
        var r2 = new ExcelReader(ms, Encoding.UTF8);
        var rows = r2.ReadRows().ToList();
        Assert.Equal(longVal.ToString(), rows[1][0]);
        Assert.True(rows[1][0] is String);
    }

    [Fact, DisplayName("三个不同Sheet表头与数据互不干扰")]
    public void MultiSheet_ThreeSheets_DifferentHeaders_And_Data()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);

        // Sheet 1: Users
        w.WriteHeader("Users", new[] { "UserId", "UserName", "Active" });
        w.WriteRows("Users", new[]
        {
            new Object?[] { 1, "Tom", true },
            new Object?[] { 2, "Jerry", false }
        });

        // Sheet 2: Orders
        w.WriteHeader("Orders", new[] { "OrderId", "Amount", "Date" });
        var orderDate = new DateTime(2024, 1, 2);
        w.WriteRows("Orders", new[]
        {
            new Object?[] { 1001, 123.45m, orderDate },
            new Object?[] { 1002, 200m, orderDate.AddDays(1) },
            new Object?[] { 1003, 0.5m, orderDate.AddDays(2) }
        });

        // Sheet 3: Logs （包含时间与文本混合，不同列数）
        w.WriteHeader("Logs", new[] { "Seq", "Level", "Message", "Time" });
        var t0 = DateTime.Now.Date.AddHours(8).AddMinutes(15).AddSeconds(30);
        w.WriteRows("Logs", new[]
        {
            new Object?[] { 1, "INFO", "Start", t0 },
            new Object?[] { 2, "WARN", "Latency", t0.AddMinutes(5) },
            new Object?[] { 3, "ERROR", "Failed", t0.AddMinutes(10) },
            new Object?[] { 4, "INFO", "Done", t0.AddMinutes(15) }
        });

        w.Save();

        File.WriteAllBytes("ew2.xlsx", ms.ToArray());

        ms.Position = 0;
        var r = new ExcelReader(ms, Encoding.UTF8);
        // 验证 sheet 名称集合包含三个
        var sheets = r.Sheets?.ToList();
        Assert.NotNull(sheets);
        Assert.Contains("Users", sheets!);
        Assert.Contains("Orders", sheets!);
        Assert.Contains("Logs", sheets!);

        // Users
        var users = r.ReadRows("Users").ToList();
        Assert.Equal(3, users.Count); // header + 2
        Assert.Equal("UserId", users[0][0]);
        Assert.Equal(1, users[1][0]);
        Assert.True(users[2][2] is Boolean && !(Boolean)users[2][2]!); // Active 列第二行 false

        // Orders
        var orders = r.ReadRows("Orders").ToList();
        Assert.Equal(4, orders.Count); // header + 3
        Assert.Equal("Amount", orders[0][1]);
        Assert.True(orders[2][2] is DateTime dt2 && dt2.Date == orderDate.AddDays(1).Date);
        Assert.True(orders[3][1] is Decimal or Double); // 金额小数

        // Logs
        var logs = r.ReadRows("Logs").ToList();
        Assert.Equal(5, logs.Count); // header + 4
        Assert.Equal("Level", logs[0][1]);
        Assert.Equal("ERROR", logs[3][1]); // 第3条日志（数据行 Seq=3）
        Assert.True(logs[4][3] is DateTime); // 时间列

        // 互不串表：确认 Users 的列数 != Logs 的列数
        Assert.NotEqual(users[0].Length, logs[0].Length);
    }

    #region 样式测试
    [Fact, DisplayName("WriteHeader带样式生成粗体字体")]
    public void WriteHeader_WithStyle_GeneratesBoldFont()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        var style = new CellStyle { Bold = true, FontSize = 12, FontName = "Arial" };
        w.WriteHeader(null!, new[] { "A", "B" }, style);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        var stylesEntry = za.GetEntry("xl/styles.xml");
        Assert.NotNull(stylesEntry);
        using var sr = new StreamReader(stylesEntry!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<b/>", xml);
        Assert.Contains("val=\"12\"", xml);
        Assert.Contains("val=\"Arial\"", xml);
    }

    [Fact, DisplayName("WriteRow带背景色样式")]
    public void WriteRow_WithBackgroundColor()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        var style = new CellStyle { BackgroundColor = "FF0000" };
        w.WriteRow(null, new Object?[] { "Red" }, style);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/styles.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("FFFF0000", xml); // fgColor with FF prefix
        Assert.Contains("solid", xml);
    }

    [Fact, DisplayName("WriteRow带边框样式")]
    public void WriteRow_WithBorder()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        var style = new CellStyle { Border = CellBorderStyle.Thin, BorderColor = "000000" };
        w.WriteRow(null, new Object?[] { "Bordered" }, style);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/styles.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("style=\"thin\"", xml);
    }

    [Fact, DisplayName("WriteRow带自定义数字格式")]
    public void WriteRow_WithCustomNumberFormat()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        var style = new CellStyle { NumberFormat = "#,##0.00" };
        w.WriteRow(null, new Object?[] { 12345.678 }, style);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/styles.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("numFmtId=\"164\"", xml); // 第一个自定义格式
        Assert.Contains("#,##0.00", xml);
    }

    [Fact, DisplayName("WriteRow带对齐和换行样式")]
    public void WriteRow_WithAlignmentAndWrapText()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        var style = new CellStyle { HAlign = HorizontalAlignment.Center, VAlign = VerticalAlignment.Center, WrapText = true };
        w.WriteRow(null, new Object?[] { "Centered" }, style);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/styles.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("horizontal=\"center\"", xml);
        Assert.Contains("vertical=\"center\"", xml);
        Assert.Contains("wrapText=\"1\"", xml);
    }
    #endregion

    #region 合并单元格测试
    [Fact, DisplayName("MergeCell生成mergeCells节点")]
    public void MergeCell_GeneratesMergeCellsNode()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A", "B", "C" });
        w.MergeCell(null, "A1:C1");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<mergeCells", xml);
        Assert.Contains("ref=\"A1:C1\"", xml);
    }

    [Fact, DisplayName("MergeCell按行列索引")]
    public void MergeCell_ByRowColIndex()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A", "B", "C", "D" });
        w.WriteRow(null, new Object?[] { "data", null, null, null });
        w.MergeCell(null, 1, 0, 1, 3); // 第2行A-D合并（0基）
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("ref=\"A2:D2\"", xml);
    }
    #endregion

    #region 冻结窗格测试
    [Fact, DisplayName("FreezePane冻结首行")]
    public void FreezePane_FirstRow()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Name" });
        w.FreezePane(null, 1);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("ySplit=\"1\"", xml);
        Assert.Contains("state=\"frozen\"", xml);
    }

    [Fact, DisplayName("FreezePane冻结行列")]
    public void FreezePane_RowAndCol()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A", "B" });
        w.FreezePane(null, 1, 1);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("xSplit=\"1\"", xml);
        Assert.Contains("ySplit=\"1\"", xml);
        Assert.Contains("activePane=\"bottomRight\"", xml);
    }
    #endregion

    #region 自动筛选测试
    [Fact, DisplayName("SetAutoFilter生成autoFilter节点")]
    public void SetAutoFilter_GeneratesAutoFilterNode()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A", "B", "C" });
        w.SetAutoFilter(null, "A1:C1");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<autoFilter ref=\"A1:C1\"", xml);
    }
    #endregion

    #region 行高测试
    [Fact, DisplayName("SetRowHeight生成行高属性")]
    public void SetRowHeight_GeneratesHeightAttribute()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A" });
        w.SetRowHeight(null, 1, 30);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("ht=\"30\"", xml);
        Assert.Contains("customHeight=\"1\"", xml);
    }
    #endregion

    #region 超链接测试
    [Fact, DisplayName("AddHyperlink生成超链接节点和关系")]
    public void AddHyperlink_GeneratesHyperlinkAndRels()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Link" });
        w.WriteRow(null, new Object?[] { "Click" });
        w.AddHyperlink(null, 2, 0, "https://example.com", "Example");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<hyperlinks>", xml);
        Assert.Contains("ref=\"A2\"", xml);
        Assert.Contains("display=\"Example\"", xml);

        // 验证关系文件
        var rels = za.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
        Assert.NotNull(rels);
        using var sr2 = new StreamReader(rels!.Open(), Encoding.UTF8);
        var relsXml = sr2.ReadToEnd();
        Assert.Contains("https://example.com", relsXml);
        Assert.Contains("hyperlink", relsXml);
    }
    #endregion

    #region 数据验证测试
    [Fact, DisplayName("AddDropdownValidation生成数据验证节点")]
    public void AddDropdownValidation_GeneratesDataValidation()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Status" });
        w.AddDropdownValidation(null, "A2:A100", new[] { "Active", "Inactive", "Pending" });
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<dataValidations", xml);
        Assert.Contains("type=\"list\"", xml);
        Assert.Contains("sqref=\"A2:A100\"", xml);
        Assert.Contains("Active", xml);
    }
    #endregion

    #region 图片测试
    [Fact, DisplayName("AddImage生成drawing和media文件")]
    public void AddImage_GeneratesDrawingAndMedia()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Image" });
        // 最小有效 PNG（1x1 transparent pixel）
        var pngData = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52 };
        w.AddImage(null, 2, 0, pngData, "png", 50, 50);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);

        // 验证 drawing 文件
        var drawing = za.GetEntry("xl/drawings/drawing1.xml");
        Assert.NotNull(drawing);
        using var sr = new StreamReader(drawing!.Open(), Encoding.UTF8);
        var drawXml = sr.ReadToEnd();
        Assert.Contains("twoCellAnchor", drawXml);
        Assert.Contains("blipFill", drawXml);

        // 验证 media 文件
        var media = za.GetEntry("xl/media/image1.png");
        Assert.NotNull(media);

        // 验证 sheet 引用 drawing
        using var sr2 = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var sheetXml = sr2.ReadToEnd();
        Assert.Contains("<drawing", sheetXml);
    }
    #endregion

    #region 页面设置测试
    [Fact, DisplayName("SetPageSetup生成页面设置节点")]
    public void SetPageSetup_GeneratesPageSetupNode()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A" });
        w.SetPageSetup(null, PageOrientation.Landscape, PaperSize.A4);
        w.SetPageMargins(null, 1.0, 1.0, 0.5, 0.5);
        w.SetHeaderFooter(null, "Header Text", "Footer Text");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("orientation=\"landscape\"", xml);
        Assert.Contains("paperSize=\"9\"", xml);
        Assert.Contains("top=\"1\"", xml);
        Assert.Contains("left=\"0.5\"", xml);
        Assert.Contains("Header Text", xml);
        Assert.Contains("Footer Text", xml);
    }

    [Fact, DisplayName("SetPrintTitleRows生成打印标题行")]
    public void SetPrintTitleRows_GeneratesDefinedNames()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A" });
        w.SetPrintTitleRows(null, 1, 1);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/workbook.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("_xlnm.Print_Titles", xml);
    }
    #endregion

    #region 工作表保护测试
    [Fact, DisplayName("ProtectSheet生成保护节点")]
    public void ProtectSheet_GeneratesProtectionNode()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A" });
        w.ProtectSheet(null, "password123");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<sheetProtection", xml);
        Assert.Contains("sheet=\"1\"", xml);
        Assert.Contains("password=", xml);
    }

    [Fact, DisplayName("ProtectSheet无密码")]
    public void ProtectSheet_WithoutPassword()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "A" });
        w.ProtectSheet(null);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<sheetProtection", xml);
        Assert.DoesNotContain("password=", xml);
    }
    #endregion

    #region 条件格式测试
    [Fact, DisplayName("AddConditionalFormat生成条件格式节点")]
    public void AddConditionalFormat_GeneratesConditionalFormatting()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Score" });
        w.WriteRow(null, new Object?[] { 90 });
        w.WriteRow(null, new Object?[] { 50 });
        w.AddConditionalFormat(null, "A2:A3", ConditionalFormatType.GreaterThan, "80", "00FF00");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("<conditionalFormatting", xml);
        Assert.Contains("operator=\"greaterThan\"", xml);
        Assert.Contains("<formula>80</formula>", xml);
    }

    [Fact, DisplayName("AddConditionalFormat数据条")]
    public void AddConditionalFormat_DataBar()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Value" });
        w.WriteRow(null, new Object?[] { 10 });
        w.AddConditionalFormat(null, "A2:A2", ConditionalFormatType.DataBar, null, "4472C4");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("type=\"dataBar\"", xml);
    }

    [Fact, DisplayName("AddConditionalFormat色阶")]
    public void AddConditionalFormat_ColorScale()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Value" });
        w.WriteRow(null, new Object?[] { 10 });
        w.AddConditionalFormat(null, "A2:A2", ConditionalFormatType.ColorScale, null, "FF6347");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("type=\"colorScale\"", xml);
    }

    [Fact, DisplayName("AddConditionalFormat介于")]
    public void AddConditionalFormat_Between()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Score" });
        w.WriteRow(null, new Object?[] { 75 });
        w.AddConditionalFormat(null, "A2:A2", ConditionalFormatType.Between, "60", "FFFF00", "90");
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("operator=\"between\"", xml);
    }
    #endregion

    #region 对象映射测试
    [Fact, DisplayName("WriteObjects导出对象集合")]
    public void WriteObjects_ExportObjectCollection()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        var users = new[]
        {
            new TestUser { Id = 1, Name = "Alice", Age = 30 },
            new TestUser { Id = 2, Name = "Bob", Age = 25 },
        };
        w.WriteObjects<TestUser>(null, users, CellStyle.Header);
        w.Save();

        ms.Position = 0;
        var r = new ExcelReader(ms, Encoding.UTF8);
        var rows = r.ReadRows().ToList();
        Assert.Equal(3, rows.Count); // header + 2 data

        // 验证表头使用 DisplayName
        Assert.Equal("编号", rows[0][0]);
        Assert.Equal("姓名", rows[0][1]);
        Assert.Equal("Age", rows[0][2]); // 无 DisplayName，使用属性名

        // 验证数据
        Assert.Equal("1", rows[1][0] + "");
        Assert.Equal("Alice", rows[1][1]);
    }

    [Fact, DisplayName("WriteDataTable导出DataTable")]
    public void WriteDataTable_ExportDataTable()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);

        var dt = new DataTable();
        dt.Columns.Add("Product", typeof(String));
        dt.Columns.Add("Price", typeof(Decimal));
        dt.Columns.Add("Qty", typeof(Int32));
        dt.Rows.Add("Apple", 3.5m, 100);
        dt.Rows.Add("Banana", 2.0m, 200);

        w.WriteDataTable(null, dt, CellStyle.Header);
        w.Save();

        ms.Position = 0;
        var r = new ExcelReader(ms, Encoding.UTF8);
        var rows = r.ReadRows().ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal("Product", rows[0][0]);
        Assert.Equal("Apple", rows[1][0]);
    }
    #endregion

    #region 列宽测试
    [Fact, DisplayName("SetColumnWidth手工设置列宽")]
    public void SetColumnWidth_ManuallySetWidth()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.WriteHeader(null!, new[] { "Name" });
        w.SetColumnWidth(null, 0, 20);
        w.Save();

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);
        using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("width=\"20\"", xml);
        Assert.Contains("customWidth=\"1\"", xml);
    }
    #endregion

    #region 辅助类
    private class TestUser
    {
        [DisplayName("编号")]
        public Int32 Id { get; set; }

        [DisplayName("姓名")]
        public String Name { get; set; } = "";

        public Int32 Age { get; set; }
    }
    #endregion
}
