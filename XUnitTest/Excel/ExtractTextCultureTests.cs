using System.ComponentModel;
using System.Globalization;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>文本提取的区域无关性：数值不得随当前区域改变小数点</summary>
/// <remarks>
/// 回归：ExtractText 曾用 ToString()，德语等区域下 1.5 输出成 1,5，逗号被当成分列符而破坏结构。
/// </remarks>
public class ExtractTextCultureTests
{
    private static void RunInGerman(Action action)
    {
        var origin = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            action();
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = origin;
        }
    }

    [Fact, DisplayName("文本提取：xlsx 小数使用固定区域小数点")]
    public void Xlsx_UsesInvariantDecimalPoint()
    {
        var path = Path.Combine(Path.GetTempPath(), $"nlo_extract_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var writer = new ExcelWriter(path))
            {
                writer.WriteHeader("Sheet1", ["名称", "金额"]);
                writer.WriteRow("Sheet1", ["苹果", 1.5]);
                writer.Save();
            }

            RunInGerman(() =>
            {
                using var reader = new ExcelReader(path);
                var text = reader.ExtractText();

                Assert.Contains("苹果,1.5", text);
                Assert.DoesNotContain("1,5", text);
            });
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("文本提取：xls 小数使用固定区域小数点")]
    public void Xls_UsesInvariantDecimalPoint()
    {
        using var writer = new BiffWriter();
        writer.WriteHeader(["名称", "金额"]);
        writer.WriteRow(["苹果", 1.5]);
        var bytes = writer.ToBytes();

        RunInGerman(() =>
        {
            using var reader = new BiffReader(new MemoryStream(bytes));
            var text = reader.ExtractText();

            Assert.Contains("苹果,1.5", text);
            Assert.DoesNotContain("1,5", text);
        });
    }
}
