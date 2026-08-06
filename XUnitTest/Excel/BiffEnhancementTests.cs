using System.ComponentModel;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>BiffReader/BiffWriter xls 增强测试 — 合并/页面设置/保护/数据验证</summary>
public class BiffEnhancementTests
{
    private static Byte[] Build(Action<BiffWriter> build)
    {
        using var writer = new BiffWriter();
        build(writer);
        return writer.ToBytes();
    }

    [Fact]
    [DisplayName("xls—合并单元格写入与读取往返")]
    public void Merges_Roundtrip()
    {
        var bytes = Build(w =>
        {
            w.WriteRow(["A", "B", "C"]);
            w.MergeCells(0, 0, 0, 2);   // A1:C1
            w.MergeCells(1, 1, 2, 1);   // B2:B3
        });

        using var reader = new BiffReader(new MemoryStream(bytes));
        var merges = reader.GetMerges();
        Assert.Equal(2, merges.Count);
        Assert.Equal((0, 0, 0, 2), merges[0]);
        Assert.Equal((1, 1, 2, 1), merges[1]);
    }

    [Fact]
    [DisplayName("xls—页面设置写入与读取往返")]
    public void PageSetup_Roundtrip()
    {
        var bytes = Build(w =>
        {
            w.SetPageSetup(landscape: true, paperSize: 9, headerMargin: 0.5, footerMargin: 0.4);
            w.SetHeaderFooter("表头", "第&P页");
            w.WriteRow(["数据"]);
        });

        using var reader = new BiffReader(new MemoryStream(bytes));
        var (landscape, paperSize, headerMargin, footerMargin, header, footer) = reader.GetPageSetup();
        Assert.True(landscape);
        Assert.Equal(9, paperSize);
        Assert.True(Math.Abs(headerMargin - 0.5) < 0.01, $"headerMargin={headerMargin}");
        Assert.True(Math.Abs(footerMargin - 0.4) < 0.01, $"footerMargin={footerMargin}");
        Assert.Equal("表头", header);
        Assert.Equal("第&P页", footer);
    }

    [Fact]
    [DisplayName("xls—工作表保护写入与读取")]
    public void Protection_Roundtrip()
    {
        // 无密码保护
        var bytes1 = Build(w =>
        {
            w.WriteRow(["A"]);
            w.ProtectSheet();
        });
        using (var reader = new BiffReader(new MemoryStream(bytes1)))
        {
            Assert.True(reader.GetProtection());
        }

        // 带密码保护
        var bytes2 = Build(w =>
        {
            w.WriteRow(["A"]);
            w.ProtectSheet("secret");
        });
        using (var reader = new BiffReader(new MemoryStream(bytes2)))
        {
            Assert.True(reader.GetProtection());
        }

        // 未保护
        var bytes3 = Build(w => w.WriteRow(["A"]));
        using (var reader = new BiffReader(new MemoryStream(bytes3)))
        {
            Assert.False(reader.GetProtection());
        }
    }

    [Fact]
    [DisplayName("xls—下拉列表数据验证写入与读取往返")]
    public void DropdownValidation_Roundtrip()
    {
        var bytes = Build(w =>
        {
            w.WriteRow(["状态"]);
            w.WriteRow(["启用"]);
            w.AddDropdownValidation("A2:A10", new[] { "启用", "禁用", "待定" });
        });

        using var reader = new BiffReader(new MemoryStream(bytes));
        var validations = reader.GetValidations();
        Assert.Single(validations);

        var dv = validations[0];
        Assert.Equal("A2:A10", dv.CellRange);
        Assert.Equal("list", dv.ValidationType);
        Assert.NotNull(dv.Items);
        Assert.Equal(new[] { "启用", "禁用", "待定" }, dv.Items);
    }

    [Fact]
    [DisplayName("xls—未设置时增强读取返回默认值")]
    public void Empty_Defaults()
    {
        var bytes = Build(w => w.WriteRow(["A"]));

        using var reader = new BiffReader(new MemoryStream(bytes));
        Assert.Empty(reader.GetMerges());
        Assert.False(reader.GetProtection());
        Assert.Empty(reader.GetValidations());

        var (landscape, paperSize, _, _, header, footer) = reader.GetPageSetup();
        Assert.False(landscape);
        Assert.Equal(0, paperSize);
        Assert.Equal(String.Empty, header);
        Assert.Equal(String.Empty, footer);
    }

    [Fact]
    [DisplayName("xls—单元格样式写入与读取往返")]
    public void CellFormats_Roundtrip()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader(new[] { "名称", "金额" });
            var style = new CellFormat { Bold = true, FontSize = 12, FontColor = "FF0000" };
            w.WriteRow(new Object?[] { "苹果", 3.5 }, style);
        });

        using var reader = new BiffReader(new MemoryStream(bytes));
        var formats = reader.ReadCellFormats();

        // 数据行（第2行）应有字体样式
        Assert.True(formats.TryGetValue((1, 0), out var cf0), "A2 应有样式");
        Assert.True(cf0.Bold, "应识别粗体");
        Assert.True(Math.Abs(cf0.FontSize - 12) < 0.5, $"FontSize={cf0.FontSize}");
        Assert.Equal("FF0000", cf0.FontColor);

        Assert.True(formats.TryGetValue((1, 1), out var cf1), "B2 应有样式");
        Assert.True(cf1.Bold);

        // 表头行（默认样式）不应有样式条目
        Assert.False(formats.ContainsKey((0, 0)));
        Assert.False(formats.ContainsKey((0, 1)));
    }

    [Fact]
    [DisplayName("xls—未设置样式时ReadCellFormats为空")]
    public void CellFormats_None()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader(new[] { "A", "B" });
            w.WriteRow(new Object?[] { 1, 2 });
        });

        using var reader = new BiffReader(new MemoryStream(bytes));
        Assert.Empty(reader.ReadCellFormats());
    }
}
