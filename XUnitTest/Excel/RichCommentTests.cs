using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>Excel 批注富文本测试（M26，对标 EPPlus）</summary>
/// <remarks>
/// 覆盖富文本批注的写入（多段 + 加粗/颜色）与读取还原，以及普通批注兼容。
/// </remarks>
public class RichCommentTests
{
    [Fact, DisplayName("批注富文本：写入与读取还原")]
    public void RichComment_WriteRead()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.SheetName = "Sheet1";
            writer.AddComment("Sheet1", 1, 0, "普通批注", "作者A");
            writer.AddComment("Sheet1", 2, 0,
                [
                    CommentSegment.Create("第一段"),
                    CommentSegment.BoldText("加粗段"),
                    CommentSegment.Colored("红字", "FFFF0000"),
                ],
                "作者B");
            writer.Save();
        }

        ms.Position = 0;
        using var reader = new ExcelReader(ms, Encoding.UTF8);

        // 富文本读取
        var rich = reader.ReadCommentSegments("Sheet1");
        Assert.Equal(2, rich.Count);

        var (author, segments) = rich[(1, 0)];
        Assert.Equal("作者B", author);
        Assert.Equal(3, segments.Count);
        Assert.Equal("第一段", segments[0].Text);
        Assert.False(segments[0].Bold);

        Assert.Equal("加粗段", segments[1].Text);
        Assert.True(segments[1].Bold);

        Assert.Equal("红字", segments[2].Text);
        Assert.Equal("FFFF0000", segments[2].Color);
    }

    [Fact, DisplayName("批注富文本：普通批注兼容")]
    public void RichComment_PlainCompatible()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.SheetName = "Sheet1";
            writer.AddComment("Sheet1", 1, 0, "普通批注", "作者A");
            writer.Save();
        }

        ms.Position = 0;
        using var reader = new ExcelReader(ms, Encoding.UTF8);

        var plain = reader.ReadComments("Sheet1");
        Assert.True(plain.TryGetValue((0, 0), out var item));
        Assert.Equal("普通批注", item.Text);
        Assert.Equal("作者A", item.Author);

        // 富文本读取也应兼容单段
        var rich = reader.ReadCommentSegments("Sheet1");
        Assert.True(rich.TryGetValue((0, 0), out var r));
        Assert.Single(r.Segments);
        Assert.Equal("普通批注", r.Segments[0].Text);
    }

    [Fact, DisplayName("批注富文本：字号与字体还原")]
    public void RichComment_FontSizeAndName()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.SheetName = "Sheet1";
            writer.AddComment("Sheet1", 1, 0,
                [new CommentSegment { Text = "大字", FontSize = 14, FontName = "宋体" }],
                "作者");
            writer.Save();
        }

        ms.Position = 0;
        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var rich = reader.ReadCommentSegments("Sheet1");
        var seg = rich[(0, 0)].Segments[0];
        Assert.Equal(14, seg.FontSize);
        Assert.Equal("宋体", seg.FontName);
    }
}
