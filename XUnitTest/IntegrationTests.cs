using System.ComponentModel;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>集成测试：完整写入所有功能后读取验证</summary>
public class IntegrationTests
{
    [Fact, DisplayName("生成复杂Excel文档供人工验收")]
    public void GenerateComplexExcel_ForManualInspection()
    {
        //var outputDir = Path.Combine(Path.GetTempPath(), "NewLife.Office", "artifacts");
        var outputDir = ".".GetFullPath();
        Directory.CreateDirectory(outputDir);

        var path = Path.Combine(outputDir, $"complex_feature_preview_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        using (var w = new ExcelWriter(path))
        {
            var headerStyle = new CellStyle
            {
                Bold = true,
                FontSize = 11,
                BackgroundColor = "4472C4",
                FontColor = "FFFFFF",
                HAlign = HorizontalAlignment.Center,
                Border = CellBorderStyle.Thin
            };

            w.WriteHeader("Data", new[] { "编号", "姓名", "年龄", "入职日期", "薪资", "在职" }, headerStyle);

            var dataStyle = new CellStyle { Border = CellBorderStyle.Thin };
            w.WriteRow("Data", new Object?[] { 1, "张三", 28, new DateTime(2020, 1, 15), 8500.50m, true }, dataStyle);
            w.WriteRow("Data", new Object?[] { 2, "李四", 35, new DateTime(2018, 6, 1), 12000m, true }, dataStyle);
            w.WriteRow("Data", new Object?[] { 3, "王五", 42, new DateTime(2015, 3, 20), 15000.75m, false }, dataStyle);

            w.WriteRow("Data", new Object?[] { });
            w.WriteRow("Data", new Object?[] { "部门汇总", null, null, null, null, null });
            w.MergeCell("Data", "A5:F5");
            w.FreezePane("Data", 1);
            w.SetAutoFilter("Data", "A1:F1");
            w.SetRowHeight("Data", 1, 25);
            w.SetColumnWidth("Data", 0, 8);
            w.SetColumnWidth("Data", 1, 12);
            w.AddHyperlink("Data", 2, 1, "https://example.com/zhangsan", "张三主页");
            w.AddDropdownValidation("Data", "F2:F100", new[] { "是", "否" });
            w.SetPageSetup("Data", PageOrientation.Landscape, PaperSize.A4);
            w.SetPageMargins("Data", 1.0, 1.0, 0.75, 0.75);
            w.SetHeaderFooter("Data", "员工信息表", "第&P页/共&N页");
            w.SetPrintTitleRows("Data", 1, 1);
            w.ProtectSheet("Data", "pass123");
            w.AddConditionalFormat("Data", "E2:E4", ConditionalFormatType.GreaterThan, "10000", "92D050");
            w.AddConditionalFormat("Data", "C2:C4", ConditionalFormatType.Between, "25", "FFFF00", "40");

            var products = new[]
            {
                new ProductInfo { Name = "Laptop", Price = 5999.99m, Stock = 100 },
                new ProductInfo { Name = "Phone", Price = 3999m, Stock = 500 },
                new ProductInfo { Name = "Tablet", Price = 2999.50m, Stock = 200 },
            };
            w.WriteObjects("Products", products, CellStyle.Header);

            var dt = new DataTable();
            dt.Columns.Add("Region", typeof(String));
            dt.Columns.Add("Sales", typeof(Decimal));
            dt.Columns.Add("Quarter", typeof(String));
            dt.Rows.Add("East", 150000m, "Q1");
            dt.Rows.Add("West", 120000m, "Q2");
            dt.Rows.Add("North", 90000m, "Q3");
            w.WriteDataTable("Sales", dt, CellStyle.Header);
            w.AddConditionalFormat("Sales", "B2:B4", ConditionalFormatType.DataBar, null, "4472C4");

            w.WriteHeader("Images", new[] { "Description", "Photo" });
            w.WriteRow("Images", new Object?[] { "Test Image", "" });
            var pngData = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            w.AddImage("Images", 2, 1, pngData, "png", 80, 60);
        }

        Assert.True(File.Exists(path));

        using var reader = new ExcelReader(path);
        var sheets = reader.Sheets?.ToList();
        Assert.NotNull(sheets);
        Assert.Equal(4, sheets!.Count);

        Console.WriteLine($"复杂Excel人工验收文件已生成：{path}");
    }

    [Fact, DisplayName("完整功能写入再读取往返测试")]
    public void FullFeature_WriteAndRead_Roundtrip()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);

        // === Sheet1: 数据表 ===
        var headerStyle = new CellStyle
        {
            Bold = true,
            FontSize = 11,
            BackgroundColor = "4472C4",
            FontColor = "FFFFFF",
            HAlign = HorizontalAlignment.Center,
            Border = CellBorderStyle.Thin
        };

        w.WriteHeader("Data", new[] { "编号", "姓名", "年龄", "入职日期", "薪资", "在职" }, headerStyle);

        var dataStyle = new CellStyle { Border = CellBorderStyle.Thin };
        w.WriteRow("Data", new Object?[] { 1, "张三", 28, new DateTime(2020, 1, 15), 8500.50m, true }, dataStyle);
        w.WriteRow("Data", new Object?[] { 2, "李四", 35, new DateTime(2018, 6, 1), 12000m, true }, dataStyle);
        w.WriteRow("Data", new Object?[] { 3, "王五", 42, new DateTime(2015, 3, 20), 15000.75m, false }, dataStyle);

        // 合并标题行
        w.WriteRow("Data", new Object?[] { }, null); // 空行
        w.WriteRow("Data", new Object?[] { "部门汇总", null, null, null, null, null }, null);
        w.MergeCell("Data", "A5:F5"); // 合并标题
        w.MergeCell("Data", 4, 0, 4, 5); // 另一种合并方式（覆盖上一行空行）

        // 冻结首行
        w.FreezePane("Data", 1);

        // 自动筛选
        w.SetAutoFilter("Data", "A1:F1");

        // 行高
        w.SetRowHeight("Data", 1, 25);

        // 列宽
        w.SetColumnWidth("Data", 0, 8);
        w.SetColumnWidth("Data", 1, 12);

        // 超链接
        w.AddHyperlink("Data", 2, 1, "https://example.com/zhangsan", "张三主页");

        // 数据验证
        w.AddDropdownValidation("Data", "F2:F100", new[] { "是", "否" });

        // 页面设置
        w.SetPageSetup("Data", PageOrientation.Landscape, PaperSize.A4);
        w.SetPageMargins("Data", 1.0, 1.0, 0.75, 0.75);
        w.SetHeaderFooter("Data", "员工信息表", "第&P页/共&N页");
        w.SetPrintTitleRows("Data", 1, 1);

        // 工作表保护
        w.ProtectSheet("Data", "pass123");

        // 条件格式：薪资>10000 绿色
        w.AddConditionalFormat("Data", "E2:E4", ConditionalFormatType.GreaterThan, "10000", "92D050");

        // 条件格式：年龄在 25-40 之间 黄色
        w.AddConditionalFormat("Data", "C2:C4", ConditionalFormatType.Between, "25", "FFFF00", "40");

        // === Sheet2: 对象映射导出 ===
        var products = new[]
        {
            new ProductInfo { Name = "Laptop", Price = 5999.99m, Stock = 100 },
            new ProductInfo { Name = "Phone", Price = 3999m, Stock = 500 },
            new ProductInfo { Name = "Tablet", Price = 2999.50m, Stock = 200 },
        };
        w.WriteObjects("Products", products, CellStyle.Header);

        // === Sheet3: DataTable 导出 ===
        var dt = new DataTable();
        dt.Columns.Add("Region", typeof(String));
        dt.Columns.Add("Sales", typeof(Decimal));
        dt.Columns.Add("Quarter", typeof(String));
        dt.Rows.Add("East", 150000m, "Q1");
        dt.Rows.Add("West", 120000m, "Q2");
        dt.Rows.Add("North", 90000m, "Q3");
        w.WriteDataTable("Sales", dt, CellStyle.Header);

        // DataBar 条件格式
        w.AddConditionalFormat("Sales", "B2:B4", ConditionalFormatType.DataBar, null, "4472C4");

        // === Sheet4: 图片 ===
        w.WriteHeader("Images", new[] { "Description", "Photo" });
        w.WriteRow("Images", new Object?[] { "Test Image", "" });
        var pngData = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        w.AddImage("Images", 2, 1, pngData, "png", 80, 60);

        w.Save();

        // ============ 读取验证 ============
        ms.Position = 0;
        var r = new ExcelReader(ms, Encoding.UTF8);

        // 验证工作表数量
        var sheets = r.Sheets?.ToList();
        Assert.NotNull(sheets);
        Assert.Equal(4, sheets!.Count);
        Assert.Contains("Data", sheets);
        Assert.Contains("Products", sheets);
        Assert.Contains("Sales", sheets);
        Assert.Contains("Images", sheets);

        // === Data sheet ===
        var dataRows = r.ReadRows("Data").ToList();
        Assert.True(dataRows.Count >= 4); // header + 3 data + extras
        // 表头
        Assert.Equal("编号", dataRows[0][0]);
        Assert.Equal("姓名", dataRows[0][1]);
        Assert.Equal("薪资", dataRows[0][4]);
        // 数据行
        Assert.Equal("张三", dataRows[1][1]);
        Assert.True(dataRows[1][3] is DateTime); // 入职日期
        Assert.True(dataRows[1][5] is Boolean); // 在职

        // 合并区域
        var merges = r.GetMergeRanges("Data");
        Assert.NotNull(merges);
        Assert.True(merges!.Count >= 1);

        // === Products sheet (对象映射) ===
        var prodRows = r.ReadRows("Products").ToList();
        Assert.Equal(4, prodRows.Count); // header + 3
        // 验证表头用了 DisplayName
        Assert.Equal("商品名称", prodRows[0][0]);
        Assert.Equal("单价", prodRows[0][1]);
        Assert.Equal("库存", prodRows[0][2]);
        // 数据
        Assert.Equal("Laptop", prodRows[1][0]);

        // === Sales sheet (DataTable) ===
        var salesRows = r.ReadRows("Sales").ToList();
        Assert.Equal(4, salesRows.Count);
        Assert.Equal("Region", salesRows[0][0]);
        Assert.Equal("Sales", salesRows[0][1]);

        // === 读取为对象 ===
        ms.Position = 0;
        var r2 = new ExcelReader(ms, Encoding.UTF8);
        var prodObjects = r2.ReadObjects<ProductInfo>("Products").ToList();
        Assert.Equal(3, prodObjects.Count);
        Assert.Equal("Laptop", prodObjects[0].Name);
        Assert.Equal(5999.99m, prodObjects[0].Price);
        Assert.Equal(100, prodObjects[0].Stock);

        // === 读取为 DataTable ===
        ms.Position = 0;
        var r3 = new ExcelReader(ms, Encoding.UTF8);
        var salesDt = r3.ReadDataTable("Sales");
        Assert.Equal(3, salesDt.Columns.Count);
        Assert.Equal(3, salesDt.Rows.Count);
        Assert.Equal("East", salesDt.Rows[0][0]);

        // ============ 验证 ZIP 结构 ============
        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true, Encoding.UTF8);

        // styles.xml 包含自定义样式
        var stylesXml = ReadEntry(za, "xl/styles.xml");
        Assert.Contains("<b/>", stylesXml); // 粗体
        Assert.Contains("FF4472C4", stylesXml); // 背景色
        Assert.Contains("style=\"thin\"", stylesXml); // 边框

        // sheet1 包含各功能节点
        var sheet1Xml = ReadEntry(za, "xl/worksheets/sheet1.xml");
        Assert.Contains("<sheetProtection", sheet1Xml);
        Assert.Contains("<autoFilter", sheet1Xml);
        Assert.Contains("<mergeCells", sheet1Xml);
        Assert.Contains("<conditionalFormatting", sheet1Xml);
        Assert.Contains("<dataValidations", sheet1Xml);
        Assert.Contains("<hyperlinks>", sheet1Xml);
        Assert.Contains("ySplit=\"1\"", sheet1Xml); // 冻结
        Assert.Contains("ht=\"25\"", sheet1Xml); // 行高
        Assert.Contains("orientation=\"landscape\"", sheet1Xml);

        // workbook 包含打印标题
        var wbXml = ReadEntry(za, "xl/workbook.xml");
        Assert.Contains("_xlnm.Print_Titles", wbXml);

        // drawing 和 media
        Assert.NotNull(za.GetEntry("xl/drawings/drawing4.xml"));
        Assert.NotNull(za.GetEntry("xl/media/image1.png"));

        // rels
        Assert.NotNull(za.GetEntry("xl/worksheets/_rels/sheet1.xml.rels"));
    }

    [Fact, DisplayName("写入文件再读取文件往返测试")]
    public void FileBasedWriteAndRead_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), "integration_test_" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            // 写入
            using (var w = new ExcelWriter(path))
            {
                w.WriteHeader("Sheet1", new[] { "Id", "Name", "Score" }, CellStyle.Header);
                w.WriteRow("Sheet1", new Object?[] { 1, "Alice", 95.5 });
                w.WriteRow("Sheet1", new Object?[] { 2, "Bob", 87.0 });
                w.FreezePane("Sheet1", 1);
                w.SetAutoFilter("Sheet1", "A1:C1");
            } // Dispose 自动保存

            Assert.True(File.Exists(path));

            // 读取
            using var reader = new ExcelReader(path);
            var rows = reader.ReadRows("Sheet1").ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal("Id", rows[0][0]);
            Assert.Equal("Alice", rows[1][1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("多Sheet对象映射和DataTable往返")]
    public void MultiSheet_ObjectAndDataTable_Roundtrip()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);

        // 对象写入
        var items = new[]
        {
            new ProductInfo { Name = "A", Price = 10m, Stock = 1 },
            new ProductInfo { Name = "B", Price = 20m, Stock = 2 },
        };
        w.WriteObjects("Items", items);

        // DataTable 写入
        var dt = new DataTable();
        dt.Columns.Add("X", typeof(Int32));
        dt.Columns.Add("Y", typeof(String));
        dt.Rows.Add(100, "Hello");
        dt.Rows.Add(200, "World");
        w.WriteDataTable("Table", dt);

        w.Save();

        // 读取 Items 为对象
        ms.Position = 0;
        var r1 = new ExcelReader(ms, Encoding.UTF8);
        var readItems = r1.ReadObjects<ProductInfo>("Items").ToList();
        Assert.Equal(2, readItems.Count);
        Assert.Equal("A", readItems[0].Name);
        Assert.Equal(10m, readItems[0].Price);

        // 读取 Table 为 DataTable
        ms.Position = 0;
        var r2 = new ExcelReader(ms, Encoding.UTF8);
        var readDt = r2.ReadDataTable("Table");
        Assert.Equal(2, readDt.Columns.Count);
        Assert.Equal(2, readDt.Rows.Count);
        Assert.Equal("100", readDt.Rows[0][0] + "");
    }

    [Fact, DisplayName("空Excel写入读取")]
    public void Empty_WriteAndRead()
    {
        using var ms = new MemoryStream();
        var w = new ExcelWriter(ms);
        w.Save();

        ms.Position = 0;
        var r = new ExcelReader(ms, Encoding.UTF8);
        var rows = r.ReadRows().ToList();
        Assert.Empty(rows);
    }

    #region 辅助
    private static String ReadEntry(ZipArchive za, String entryPath)
    {
        var entry = za.GetEntry(entryPath);
        if (entry == null) return "";
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }
    #endregion

    #region 辅助类
    private class ProductInfo
    {
        [DisplayName("商品名称")]
        public String Name { get; set; } = "";

        [DisplayName("单价")]
        public Decimal Price { get; set; }

        [DisplayName("库存")]
        public Int32 Stock { get; set; }
    }
    #endregion
}
