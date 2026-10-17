using System.ComponentModel;
using System.Data;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>ExcelHelper 静态便捷入口与 CellFormula 公式封装测试</summary>
public class ExcelHelperTests
{
    #region 辅助
    private static String TempFile() => Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");

    public class Person
    {
        public Int32 Id { get; set; }
        public String? Name { get; set; }
        public Double Score { get; set; }
    }
    #endregion

    #region SaveAs / Query
    [Fact, DisplayName("ExcelHelper_对象集合导出并强类型读回")]
    public void SaveAs_Query_Objects_RoundTrip()
    {
        var path = TempFile();
        try
        {
            var data = new[]
            {
                new Person { Id = 1, Name = "张三", Score = 95.5 },
                new Person { Id = 2, Name = "李四", Score = 88 },
            };

            ExcelHelper.SaveAs(path, data, "人员");

            // 原始行读取：表头 + 2 行数据
            var rows = ExcelHelper.Query(path, "人员").ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal("Id", rows[0][0]);
            Assert.Equal("Name", rows[0][1]);
            Assert.Equal("Score", rows[0][2]);
            Assert.Equal(1L, Convert.ToInt64(rows[1][0]));
            Assert.Equal("张三", rows[1][1]);

            // 强类型读取
            var back = ExcelHelper.Query<Person>(path, "人员").ToList();
            Assert.Equal(2, back.Count);
            Assert.Equal(1, back[0].Id);
            Assert.Equal("张三", back[0].Name);
            Assert.Equal(95.5, back[0].Score);
            Assert.Equal(2, back[1].Id);
            Assert.Equal(88d, back[1].Score);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("ExcelHelper_DataTable导出并读回DataTable")]
    public void SaveAs_Query_DataTable_RoundTrip()
    {
        var path = TempFile();
        try
        {
            var table = new DataTable("销售");
            table.Columns.Add("产品", typeof(String));
            table.Columns.Add("数量", typeof(Int32));
            table.Rows.Add("键盘", 10);
            table.Rows.Add("鼠标", 20);

            ExcelHelper.SaveAs(path, table, "销售");

            var back = ExcelHelper.QueryAsDataTable(path, "销售");
            Assert.Equal(2, back.Columns.Count);
            Assert.Equal(2, back.Rows.Count);
            Assert.Equal("产品", back.Columns[0].ColumnName);
            Assert.Equal("键盘", back.Rows[0]["产品"]);
            Assert.Equal(20, Convert.ToInt32(back.Rows[1]["数量"]));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("ExcelHelper_二维行数据导出")]
    public void SaveAs_Rows_WritesAllRows()
    {
        var path = TempFile();
        try
        {
            var rows = new List<Object?[]>
            {
                new Object?[] { "编号", "名称" },
                new Object?[] { 1, "甲" },
                new Object?[] { 2, "乙" },
                new Object?[] { 3, null },
            };

            ExcelHelper.SaveAs(path, (IEnumerable<Object?[]>)rows, "数据");

            var back = ExcelHelper.Query(path, "数据").ToList();
            Assert.Equal(4, back.Count);
            Assert.Equal("编号", back[0][0]);
            Assert.Equal(3L, Convert.ToInt64(back[3][0]));
            Assert.True(back[3].Length < 2 || back[3][1] == null);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("ExcelHelper_写入流并读回")]
    public void SaveAs_Stream_WritesReadableWorkbook()
    {
        var path = TempFile();
        try
        {
            using var ms = new MemoryStream();
            ExcelHelper.SaveAs(ms, new[] { new Person { Id = 7, Name = "流", Score = 1.5 } }, "流表");

            Assert.True(ms.Length > 0);
            File.WriteAllBytes(path, ms.ToArray());

            using var reader = new ExcelReader(path);
            var document = reader.ReadExcel();
            Assert.Single(document.Worksheets);
            Assert.Equal(2, document.Worksheets[0].Rows.Count);
            Assert.Equal("流", document.Worksheets[0].Rows[1][1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("ExcelHelper_参数校验抛出ArgumentNullException")]
    public void Arguments_NullOrEmpty_Throw()
    {
        using var ms = new MemoryStream();
        var data = new[] { new Person() };

        Assert.Throws<ArgumentNullException>(() => ExcelHelper.SaveAs((String)null!, data));
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.SaveAs("", data));
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.SaveAs((Stream)null!, data));
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.SaveAs<Person>(ms, null!));
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.SaveAs((String)null!, new DataTable()));
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.SaveAs("x.xlsx", (DataTable)null!));
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.QueryAsDataTable(null!));
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.Query<Person>(null!).ToList());
        Assert.Throws<ArgumentNullException>(() => ExcelHelper.Query((String)null!).ToList());
    }
    #endregion

    #region CellFormula
    [Fact, DisplayName("CellFormula_属性与文本")]
    public void CellFormula_Properties()
    {
        var formula = new CellFormula("SUM(A1:A2)", 3);

        Assert.Equal("SUM(A1:A2)", formula.Formula);
        Assert.Equal(3, formula.CachedValue);
        Assert.Equal("=SUM(A1:A2)", formula.ToString());
        Assert.Null(new CellFormula("NOW()").CachedValue);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [DisplayName("CellFormula_空公式抛出异常")]
    public void CellFormula_Empty_Throws(String? formula) =>
        Assert.Throws<ArgumentNullException>(() => new CellFormula(formula!));

    [Fact, DisplayName("CellFormula_写入公式与缓存值并可读回")]
    public void CellFormula_WriteAndRead_RoundTrip()
    {
        var path = TempFile();
        try
        {
            using (var writer = new ExcelWriter(path))
            {
                writer.WriteRow("Sheet1", [1, 2]);
                writer.WriteRow("Sheet1", [new CellFormula("SUM(A1:A2)", 3), new CellFormula("A1*2", 2)]);
                writer.Save();
            }

            using var reader = new ExcelReader(path);
            var document = reader.ReadExcel();
            var sheet = document.Worksheets[0];

            Assert.Equal("SUM(A1:A2)", sheet.Formulas[(1, 0)]);
            Assert.Equal("A1*2", sheet.Formulas[(1, 1)]);
            Assert.Equal(3d, Convert.ToDouble(sheet.Rows[1][0]));
            Assert.Equal(2d, Convert.ToDouble(sheet.Rows[1][1]));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
    #endregion
}
