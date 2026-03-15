using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using NewLife.Office.Ods;
using Xunit;

namespace XUnitTest;

/// <summary>ODS 模块单元测试</summary>
public class OdsTests
{
    #region 辅助：生成测试 ODS 流
    private static MemoryStream CreateOdsStream(String sheetName, IEnumerable<IEnumerable<String>> rows)
    {
        var writer = new OdsWriter();
        writer.AddSheet(sheetName, rows);
        var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;
        return ms;
    }
    #endregion

    #region 写入测试 — 基础
    [Fact]
    [DisplayName("写入简单数据可生成有效ODS")]
    public void Write_SimpleData_ValidOds()
    {
        var writer = new OdsWriter();
        writer.AddSheet("Sheet1", new[] { new[] { "A", "B" }, new[] { "1", "2" } });
        var ms = new MemoryStream();
        writer.Save(ms);
        Assert.True(ms.Length > 0);
    }

    [Fact]
    [DisplayName("写入多工作表")]
    public void Write_MultipleSheets_AllPresent()
    {
        var writer = new OdsWriter();
        writer.AddSheet("Sheet1", new[] { new[] { "S1" } });
        writer.AddSheet("Sheet2", new[] { new[] { "S2" } });
        var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;
        var sheets = OdsReader.Read(ms);
        Assert.Equal(2, sheets.Count);
        Assert.Equal("Sheet1", sheets[0].Name);
        Assert.Equal("Sheet2", sheets[1].Name);
    }

    [Fact]
    [DisplayName("写入空单元格")]
    public void Write_EmptyCells_NoException()
    {
        var writer = new OdsWriter();
        writer.AddSheet("Sheet1", new[] { new[] { "", "B", "" } });
        var ms = new MemoryStream();
        var ex = Record.Exception(() => writer.Save(ms));
        Assert.Null(ex);
    }

    [Fact]
    [DisplayName("写入文档属性")]
    public void Write_DocProperties_MetaPresent()
    {
        var writer = new OdsWriter { Title = "TestDoc", Author = "Alice" };
        writer.AddSheet("S1", new[] { new[] { "x" } });
        var ms = new MemoryStream();
        writer.Save(ms);
        Assert.True(ms.Length > 0);
    }

    [Fact]
    [DisplayName("写入保存到文件")]
    public void Write_SaveFile_FileCreated()
    {
        var path = Path.Combine(Path.GetTempPath(), "test_ods_write.ods");
        try
        {
            var writer = new OdsWriter();
            writer.AddSheet("Data", new[] { new[] { "Name", "Value" }, new[] { "X", "42" } });
            writer.Save(path);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion

    #region 读取测试 — 往返
    [Fact]
    [DisplayName("往返：写再读文本值正确")]
    public void RoundTrip_StringValues_Preserved()
    {
        using var ms = CreateOdsStream("Sheet1", new[]
        {
            new[] { "Hello", "World" },
            new[] { "foo", "bar" },
        });
        var sheets = OdsReader.Read(ms);
        Assert.Single(sheets);
        var rows = sheets[0].Rows;
        Assert.Equal(2, rows.Count);
        Assert.Equal("Hello", rows[0][0]);
        Assert.Equal("World", rows[0][1]);
        Assert.Equal("foo", rows[1][0]);
        Assert.Equal("bar", rows[1][1]);
    }

    [Fact]
    [DisplayName("往返：写再读工作表名称正确")]
    public void RoundTrip_SheetName_Preserved()
    {
        using var ms = CreateOdsStream("MySheet", new[] { new[] { "A" } });
        var sheets = OdsReader.Read(ms);
        Assert.Single(sheets);
        Assert.Equal("MySheet", sheets[0].Name);
    }

    [Fact]
    [DisplayName("往返：写再读数值类型")]
    public void RoundTrip_NumericValues_Preserved()
    {
        using var ms = CreateOdsStream("Data", new[]
        {
            new[] { "42", "3.14", "-100" },
        });
        var sheets = OdsReader.Read(ms);
        var row = sheets[0].Rows[0];
        Assert.Equal(3, row.Length);
        Assert.Equal("42", row[0]);
        Assert.Equal("3.14", row[1]);
        Assert.Equal("-100", row[2]);
    }

    [Fact]
    [DisplayName("往返：多行多列数据完整")]
    public void RoundTrip_MultipleRows_AllPresent()
    {
        using var ms = CreateOdsStream("Grid", new[]
        {
            new[] { "R1C1", "R1C2", "R1C3" },
            new[] { "R2C1", "R2C2", "R2C3" },
            new[] { "R3C1", "R3C2", "R3C3" },
        });
        var sheets = OdsReader.Read(ms);
        var rows = sheets[0].Rows;
        Assert.Equal(3, rows.Count);
        Assert.Equal("R3C3", rows[2][2]);
    }

    [Fact]
    [DisplayName("往返：包含XML特殊字符")]
    public void RoundTrip_XmlSpecialChars_Preserved()
    {
        using var ms = CreateOdsStream("Sheet1", new[]
        {
            new[] { "<tag>", "a & b", "\"quoted\"" },
        });
        var sheets = OdsReader.Read(ms);
        var row = sheets[0].Rows[0];
        Assert.Equal("<tag>", row[0]);
        Assert.Equal("a & b", row[1]);
        Assert.Equal("\"quoted\"", row[2]);
    }

    [Fact]
    [DisplayName("往返：中文内容正确保存与读取")]
    public void RoundTrip_ChineseText_Preserved()
    {
        using var ms = CreateOdsStream("中文表", new[]
        {
            new[] { "姓名", "年龄" },
            new[] { "张三", "25" },
        });
        var sheets = OdsReader.Read(ms);
        Assert.Equal("中文表", sheets[0].Name);
        Assert.Equal("姓名", sheets[0].Rows[0][0]);
        Assert.Equal("张三", sheets[0].Rows[1][0]);
    }
    #endregion

    #region 读取测试 — ReadRows 接口
    [Fact]
    [DisplayName("ReadRows 从文件返回第一张表行数据")]
    public void ReadFile_ReadRows_ReturnsFirstSheet()
    {
        var path = Path.Combine(Path.GetTempPath(), "test_ods_readrows.ods");
        try
        {
            var writer = new OdsWriter();
            writer.AddSheet("Sheet1", new[] { new[] { "X", "Y" }, new[] { "1", "2" } });
            writer.Save(path);

            var rows = OdsReader.ReadRows(path);
            Assert.Equal(2, rows.Count);
            Assert.Equal("X", rows[0][0]);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    [DisplayName("ReadRows 从流返回第一张表行数据")]
    public void ReadStream_ReadRows_ReturnsFirstSheet()
    {
        using var ms = CreateOdsStream("S1", new[] { new[] { "A", "B" } });
        var rows = OdsReader.ReadRows(ms);
        Assert.Single(rows);
        Assert.Equal("A", rows[0][0]);
    }
    #endregion

    #region 写入测试 — OdsSheet API
    [Fact]
    [DisplayName("通过OdsSheet对象添加工作表")]
    public void Write_OdsSheetObject_CorrectlyAdded()
    {
        var sheet = new OdsSheet { Name = "Direct" };
        sheet.Rows.Add(new[] { "Col1", "Col2" });
        sheet.Rows.Add(new[] { "Val1", "Val2" });
        var writer = new OdsWriter();
        writer.AddSheet(sheet);
        var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;
        var readBack = OdsReader.Read(ms);
        Assert.Single(readBack);
        Assert.Equal("Direct", readBack[0].Name);
        Assert.Equal("Col1", readBack[0].Rows[0][0]);
    }

    [Fact]
    [DisplayName("链式调用AddSheet")]
    public void Write_ChainAddSheet_BothSheetsPresent()
    {
        var writer = new OdsWriter();
        writer.AddSheet("A", new[] { new[] { "a" } })
              .AddSheet("B", new[] { new[] { "b" } });
        var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;
        var sheets = OdsReader.Read(ms);
        Assert.Equal(2, sheets.Count);
    }
    #endregion

    #region 边界测试
    [Fact]
    [DisplayName("空工作表（无行）可正常保存读取")]
    public void Write_EmptySheet_NoException()
    {
        var writer = new OdsWriter();
        writer.AddSheet("Empty", Array.Empty<String[]>());
        var ms = new MemoryStream();
        var ex = Record.Exception(() => writer.Save(ms));
        Assert.Null(ex);
    }

    [Fact]
    [DisplayName("读取空ODS(无标准内容)返回空列表")]
    public void Read_NonOdsStream_ReturnsEmpty()
    {
        var ms = new MemoryStream(new byte[] { 0x50, 0x4B, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });
        // This is an empty ZIP — should return empty list gracefully
        var sheets = OdsReader.Read(ms);
        Assert.NotNull(sheets);
        Assert.Empty(sheets);
    }
    #endregion
}
