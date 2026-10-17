using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Common;

/// <summary>跨格式共享模型（Color / OfficeFont / Theme）单元测试</summary>
public class OfficeModelTests
{
    #region Color
    [Fact, DisplayName("Color_默认构造为不透明黑色")]
    public void Color_Default_IsOpaqueBlack()
    {
        var color = new Color();

        Assert.Equal(0, color.R);
        Assert.Equal(0, color.G);
        Assert.Equal(0, color.B);
        Assert.Equal((Byte)255, color.A);
        Assert.Equal("000000", color.Hex);
        Assert.Equal("FF000000", color.HexArgb);
        Assert.False(color.IsTransparent);
    }

    [Theory]
    [InlineData("FF0000", "FF0000", "FFFF0000")]
    [InlineData("#00FF00", "00FF00", "FF00FF00")]
    [InlineData("#0000FF", "0000FF", "FF0000FF")]
    [DisplayName("Color_FromHex解析6位RGB")]
    public void Color_FromHex_Rgb(String hex, String expectedHex, String expectedArgb)
    {
        var color = Color.FromHex(hex);

        Assert.Equal(expectedHex, color.Hex);
        Assert.Equal(expectedArgb, color.HexArgb);
        Assert.Equal((Byte)255, color.A);
    }

    [Fact, DisplayName("Color_FromHex解析8位ARGB")]
    public void Color_FromHex_Argb()
    {
        var color = Color.FromHex("80FF8800");

        Assert.Equal((Byte)128, color.A);
        Assert.Equal("FF8800", color.Hex);
        Assert.Equal("80FF8800", color.HexArgb);
        Assert.False(color.IsTransparent);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("XYZ")]
    [InlineData("12345")]
    [DisplayName("Color_FromHex非法输入回退黑色")]
    public void Color_FromHex_Invalid_ReturnsBlack(String? hex)
    {
        var color = Color.FromHex(hex);

        Assert.Equal(0, color.R);
        Assert.Equal(0, color.G);
        Assert.Equal(0, color.B);
        Assert.Equal((Byte)255, color.A);
    }

    [Fact, DisplayName("Color_WithAlpha生成新实例且不修改原值")]
    public void Color_WithAlpha_ReturnsCopy()
    {
        var red = Color.Red;

        var transparent = red.WithAlpha(0);

        Assert.True(transparent.IsTransparent);
        Assert.Equal(0, transparent.A);
        Assert.Equal("FF0000", transparent.Hex);
        Assert.Equal((Byte)255, red.A);
        Assert.False(red.IsTransparent);
    }

    [Fact, DisplayName("Color_预定义颜色取值正确")]
    public void Color_Predefined_Values()
    {
        Assert.Equal("000000", Color.Black.Hex);
        Assert.Equal("FFFFFF", Color.White.Hex);
        Assert.Equal("FF0000", Color.Red.Hex);
        Assert.Equal("008000", Color.Green.Hex);
        Assert.Equal("0000FF", Color.Blue.Hex);
        Assert.True(Color.Transparent.IsTransparent);
    }
    #endregion

    #region Theme
    [Fact, DisplayName("Theme_内置色板名称不区分大小写")]
    public void Theme_Builtin_IsCaseInsensitive()
    {
        var lower = Theme.Get("blue");
        var upper = Theme.Get("BLUE");

        Assert.Equal(6, lower.Length);
        Assert.Equal(lower, upper);
        Assert.Equal("2563EB", lower[0]);
        Assert.Equal("DBEAFE", lower[5]);
    }

    [Fact, DisplayName("Theme_空值与未知名称回退blue色板")]
    public void Theme_Unknown_FallsBackToBlue()
    {
        var expected = Theme.Get("blue");

        Assert.Equal(expected, Theme.Get(null));
        Assert.Equal(expected, Theme.Get(""));
        Assert.Equal(expected, Theme.Get("not-exists"));
    }

    [Fact, DisplayName("Theme_自定义解析器优先且可回退内置色板")]
    public void Theme_CustomResolver_PrecedenceAndFallback()
    {
        String[] custom = ["111111", "222222", "333333", "444444", "555555", "666666"];

        Assert.Equal(custom, Theme.Get("brand", _ => custom));
        Assert.Equal("111111", Theme.GetPrimary("brand", _ => custom));
        Assert.Equal("666666", Theme.GetLight("brand", _ => custom));

        // 解析器返回 null 或空数组时回退内置色板
        Assert.Equal(Theme.Get("blue"), Theme.Get("brand", _ => null));
        Assert.Equal(Theme.Get("blue"), Theme.Get("brand", _ => []));
        // 主题为空时不调用解析器
        Assert.Equal(Theme.Get("blue"), Theme.Get(null, _ => custom));
    }

    [Theory]
    [InlineData("blue", "2563EB")]
    [InlineData("green", "16A34A")]
    [InlineData("ocean", "0EA5E9")]
    [InlineData("amber", "F59E0B")]
    [DisplayName("Theme_GetPrimary返回Accent1")]
    public void Theme_GetPrimary(String theme, String expected) => Assert.Equal(expected, Theme.GetPrimary(theme));

    [Theory]
    [InlineData("dark", "0F172A")]
    [InlineData("ocean", "0F172A")]
    [InlineData("blue", "FFFFFF")]
    [InlineData(null, "FFFFFF")]
    [DisplayName("Theme_GetBackground按主题返回背景色")]
    public void Theme_GetBackground(String? theme, String expected) => Assert.Equal(expected, Theme.GetBackground(theme));
    #endregion

    #region OfficeFont
    [Fact, DisplayName("OfficeFont_默认值")]
    public void OfficeFont_Defaults()
    {
        var font = new OfficeFont();

        Assert.Null(font.Name);
        Assert.Null(font.EastAsianName);
        Assert.Null(font.EffectiveName);
        Assert.False(font.Bold);
        Assert.False(font.Italic);
        Assert.False(font.Underline);
        Assert.False(font.Strikethrough);
        Assert.Null(font.Color);
    }

    [Theory]
    [InlineData("Arial", null, "Arial")]
    [InlineData("Arial", "微软雅黑", "微软雅黑")]
    [InlineData(null, "宋体", "宋体")]
    [DisplayName("OfficeFont_EffectiveName优先东亚字体")]
    public void OfficeFont_EffectiveName(String? name, String? eastAsianName, String? expected)
    {
        var font = new OfficeFont { Name = name, EastAsianName = eastAsianName };

        Assert.Equal(expected, font.EffectiveName);
    }
    #endregion
}
