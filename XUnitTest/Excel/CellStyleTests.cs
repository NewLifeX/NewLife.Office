using System.ComponentModel;
using NewLife.Office;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Ole2;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using NewLife.Office.Xps;
using Xunit;

namespace XUnitTest.Excel;

public class CellStyleTests
{
    [Fact, DisplayName("默认CellStyle属性值")]
    public void Default_Properties()
    {
        var style = new NewLife.Office.Excel.CellFormat();
        Assert.Null(style.FontName);
        Assert.Equal(0, style.FontSize);
        Assert.False(style.Bold);
        Assert.False(style.Italic);
        Assert.False(style.Underline);
        Assert.Null(style.FontColor);
        Assert.Null(style.BackgroundColor);
        Assert.Equal(HorizontalAlignment.General, style.HAlign);
        Assert.Equal(VerticalAlignment.Top, style.VAlign);
        Assert.False(style.WrapText);
        Assert.Equal(NewLife.Office.Excel.BorderStyle.None, style.Border);
        Assert.Null(style.BorderColor);
        Assert.Null(style.NumberFormat);
    }

    [Fact, DisplayName("Header静态样式")]
    public void Header_StaticStyle()
    {
        var style = NewLife.Office.Excel.CellFormat.Header;
        Assert.True(style.Bold);
        Assert.Equal(0, style.FontSize);
        Assert.Equal(HorizontalAlignment.General, style.HAlign);
    }

    [Fact, DisplayName("Title静态样式")]
    public void Title_StaticStyle()
    {
        var style = NewLife.Office.Excel.CellFormat.Title;
        Assert.True(style.Bold);
        Assert.Equal(14, style.FontSize);
        Assert.Equal(HorizontalAlignment.Center, style.HAlign);
    }

    [Fact, DisplayName("设置所有属性")]
    public void Set_All_Properties()
    {
        var style = new NewLife.Office.Excel.CellFormat
        {
            FontName = "Consolas",
            FontSize = 16,
            Bold = true,
            Italic = true,
            Underline = true,
            FontColor = "FF0000",
            BackgroundColor = "00FF00",
            HAlign = HorizontalAlignment.Right,
            VAlign = VerticalAlignment.Bottom,
            WrapText = true,
            Border = NewLife.Office.Excel.BorderStyle.Medium,
            BorderColor = "0000FF",
            NumberFormat = "yyyy-MM-dd"
        };

        Assert.Equal("Consolas", style.FontName);
        Assert.Equal(16, style.FontSize);
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.True(style.Underline);
        Assert.Equal("FF0000", style.FontColor);
        Assert.Equal("00FF00", style.BackgroundColor);
        Assert.Equal(HorizontalAlignment.Right, style.HAlign);
        Assert.Equal(VerticalAlignment.Bottom, style.VAlign);
        Assert.True(style.WrapText);
        Assert.Equal(NewLife.Office.Excel.BorderStyle.Medium, style.Border);
        Assert.Equal("0000FF", style.BorderColor);
        Assert.Equal("yyyy-MM-dd", style.NumberFormat);
    }

    [Fact, DisplayName("枚举值完整性")]
    public void Enum_Values()
    {
        // HorizontalAlignment
        Assert.Equal(6, Enum.GetValues(typeof(HorizontalAlignment)).Length);

        // VerticalAlignment
        Assert.Equal(3, Enum.GetValues(typeof(VerticalAlignment)).Length);

        // CellBorderStyle
        Assert.Equal(7, Enum.GetValues(typeof(NewLife.Office.Excel.BorderStyle)).Length);

        // PageOrientation
        Assert.Equal(2, Enum.GetValues(typeof(PageOrientation)).Length);

        // PaperSize
        Assert.Equal(4, Enum.GetValues(typeof(PaperSize)).Length);

        // ConditionalFormatType
        Assert.Equal(10, Enum.GetValues(typeof(ConditionalFormatValues)).Length);
    }

    [Fact, DisplayName("Header和Title每次返回新实例")]
    public void Header_Title_NewInstance()
    {
        var h1 = NewLife.Office.Excel.CellFormat.Header;
        var h2 = NewLife.Office.Excel.CellFormat.Header;
        Assert.NotSame(h1, h2);

        var t1 = NewLife.Office.Excel.CellFormat.Title;
        var t2 = NewLife.Office.Excel.CellFormat.Title;
        Assert.NotSame(t1, t2);
    }
}
