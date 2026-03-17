using System.ComponentModel;
using System.Data;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>BiffWriter xls BIFF8 写入器单元测试</summary>
public class BiffWriterTests
{
    #region 基础写入

    [Fact]
    [DisplayName("写入单行数据并读回验证")]
    public void WriteAndRead_SingleRow()
    {
        using var writer = new BiffWriter();
        writer.WriteRow(new Object?[] { "Hello", 42.0, true });

        var bytes = writer.ToBytes();
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 512, "xls 文件应大于 512 字节（OLE2 最小扇区）");

        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        Assert.Equal(3, rows[0].Length);
        Assert.Equal("Hello", rows[0][0]);
        Assert.Equal(42.0, rows[0][1]);
        Assert.Equal(true, rows[0][2]);
    }

    [Fact]
    [DisplayName("写入标题行和多数据行后读回")]
    public void WriteHeader_And_DataRows()
    {
        using var writer = new BiffWriter();
        writer.WriteHeader(["姓名", "年龄", "工资"]);
        writer.WriteRow(["Alice", 30, 5000.0]);
        writer.WriteRow(["Bob", 25, 4500.5]);

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();

        Assert.Equal(3, rows.Count);
        Assert.Equal("姓名", rows[0][0]);
        Assert.Equal("Alice", rows[1][0]);
        Assert.Equal(30.0, rows[1][1]);
        Assert.Equal(5000.0, rows[1][2]);
        Assert.Equal("Bob", rows[2][0]);
    }

    [Fact]
    [DisplayName("写入 null 单元格")]
    public void WriteRow_WithNullValues()
    {
        using var writer = new BiffWriter();
        writer.WriteRow(new Object?[] { "A", null, 3.14 });

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        Assert.Equal("A", rows[0][0]);
        Assert.Null(rows[0][1]);
        Assert.Equal(3.14, rows[0][2]);
    }

    [Fact]
    [DisplayName("写入多工作表")]
    public void WriteMultipleSheets()
    {
        using var writer = new BiffWriter();
        writer.WriteRow(["Sheet1 Data"]);

        writer.SheetName = "Sheet2";
        writer.WriteRow(["Sheet2 Data"]);

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        Assert.Equal(2, reader.SheetNames.Count);

        var s1 = reader.ReadSheet("Sheet1").ToList();
        Assert.Single(s1);
        Assert.Equal("Sheet1 Data", s1[0][0]);

        var s2 = reader.ReadSheet("Sheet2").ToList();
        Assert.Single(s2);
        Assert.Equal("Sheet2 Data", s2[0][0]);
    }

    [Fact]
    [DisplayName("保存到文件并重新读取")]
    public void SaveToFile_And_Read()
    {
        var path = Path.Combine(Path.GetTempPath(), $"BiffWriterTest_{Guid.NewGuid():N}.xls");
        try
        {
            using var writer = new BiffWriter();
            writer.WriteHeader(["ID", "Value"]);
            writer.WriteRow([1, "TestValue"]);

            writer.Save(path);

            Assert.True(File.Exists(path), "文件应被创建");
            Assert.True(new FileInfo(path).Length > 512, "文件大小应大于 512 字节");

            using var reader = new BiffReader(path);
            var rows = reader.ReadSheet().ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("ID", rows[0][0]);
            Assert.Equal("Value", rows[0][1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    #endregion

    #region 数据类型

    [Fact]
    [DisplayName("写入不同数值类型")]
    public void WriteRow_NumericTypes()
    {
        using var writer = new BiffWriter();
        writer.WriteRow(new Object?[]
        {
            100,
            999999999L,
            3.14f,
            9999.99m,
            Double.MaxValue
        });

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        Assert.Equal(5, rows[0].Length);
        Assert.Equal(100.0, rows[0][0]);
        Assert.Equal(999999999.0, rows[0][1]);
    }

    [Fact]
    [DisplayName("写入 DateTime 并读回（精度到天）")]
    public void WriteRow_DateTime()
    {
        var dt = new DateTime(2025, 6, 15, 0, 0, 0, DateTimeKind.Unspecified);
        using var writer = new BiffWriter();
        writer.WriteRow(new Object?[] { dt });

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        // Excel 日期序列号：2025-06-15 = 45823 + 偏移
        var serial = (Double?)rows[0][0];
        Assert.NotNull(serial);
        Assert.True(serial > 40000, "日期序列号应在合理范围内");
    }

    [Fact]
    [DisplayName("写入布尔值")]
    public void WriteRow_Boolean()
    {
        using var writer = new BiffWriter();
        writer.WriteRow(new Object?[] { true, false });

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        Assert.Equal(true, rows[0][0]);
        Assert.Equal(false, rows[0][1]);
    }

    [Fact]
    [DisplayName("写入空字符串")]
    public void WriteRow_EmptyString()
    {
        using var writer = new BiffWriter();
        writer.WriteRow(new Object?[] { "", "有内容" });

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        Assert.Equal("", rows[0][0]);
        Assert.Equal("有内容", rows[0][1]);
    }

    [Fact]
    [DisplayName("写入 Unicode/中文字符")]
    public void WriteRow_CjkStrings()
    {
        using var writer = new BiffWriter();
        writer.WriteRow(["中文", "日本語", "한국어", "العربية"]);

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        Assert.Equal("中文", rows[0][0]);
        Assert.Equal("日本語", rows[0][1]);
        Assert.Equal("한국어", rows[0][2]);
        Assert.Equal("العربية", rows[0][3]);
    }

    #endregion

    #region 对象映射

    [Fact]
    [DisplayName("WriteObjects 映射 POCO 集合到工作表")]
    public void WriteObjects_PocoColl()
    {
        var data = new List<SampleModel>
        {
            new() { Id = 1, Name = "Alice", Score = 95.5 },
            new() { Id = 2, Name = "Bob", Score = 88.0 },
        };

        using var writer = new BiffWriter();
        writer.WriteObjects(data);

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Equal(3, rows.Count); // 1 标题 + 2 数据

        Assert.Equal("Id", rows[0][0]);
        Assert.Equal("Name", rows[0][1]);
        Assert.Equal("Score", rows[0][2]);

        Assert.Equal(1.0, rows[1][0]);
        Assert.Equal("Alice", rows[1][1]);
        Assert.Equal(95.5, rows[1][2]);
    }

    [Fact]
    [DisplayName("WriteObjects 使用 DisplayName 作为列标题")]
    public void WriteObjects_DisplayName()
    {
        var data = new List<SampleWithDisplayName>
        {
            new() { Id = 1, FullName = "TestUser" }
        };

        using var writer = new BiffWriter();
        writer.WriteObjects(data);

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal("编号", rows[0][0]);
        Assert.Equal("全名", rows[0][1]);
    }

    [Fact]
    [DisplayName("WriteDataTable 映射 DataTable 到工作表")]
    public void WriteDataTable_Basic()
    {
        var table = new DataTable();
        table.Columns.Add("产品", typeof(String));
        table.Columns.Add("数量", typeof(Int32));
        table.Columns.Add("单价", typeof(Double));
        table.Rows.Add("苹果", 100, 2.5);
        table.Rows.Add("香蕉", 200, 1.8);

        using var writer = new BiffWriter();
        writer.WriteDataTable(table);

        var bytes = writer.ToBytes();
        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal("产品", rows[0][0]);
        Assert.Equal("苹果", rows[1][0]);
        Assert.Equal(100.0, rows[1][1]);
        Assert.Equal(2.5, rows[1][2]);
    }

    #endregion

    #region 大数据量

    [Fact]
    [DisplayName("写入 1000 行数据性能验证")]
    public void Write_1000Rows()
    {
        using var writer = new BiffWriter();
        writer.WriteHeader(["Index", "Text", "Value"]);

        for (var i = 1; i <= 1000; i++)
        {
            writer.WriteRow([i, $"Row_{i}", i * 1.5]);
        }

        var bytes = writer.ToBytes();
        Assert.True(bytes.Length > 10_000, "1000行数据文件应大于 10KB");

        using var reader = new BiffReader(new MemoryStream(bytes));
        var rows = reader.ReadSheet().ToList();
        Assert.Equal(1001, rows.Count); // 1 标题 + 1000 数据
        Assert.Equal("Index", rows[0][0]);
        Assert.Equal(1000.0, rows[1000][0]);
    }

    #endregion

    #region 辅助类型

    private class SampleModel
    {
        public Int32 Id { get; set; }
        public String Name { get; set; } = "";
        public Double Score { get; set; }
    }

    private class SampleWithDisplayName
    {
        [DisplayName("编号")]
        public Int32 Id { get; set; }

        [DisplayName("全名")]
        public String FullName { get; set; } = "";
    }

    #endregion
}
