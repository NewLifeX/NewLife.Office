using System.ComponentModel;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using NewLife.Office.Ppt;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace XUnitTest.Ppt;

/// <summary>PPTX 交叉兼容：用 OpenXml SDK 按语义读取本项目生成的 pptx</summary>
/// <remarks>
/// OpenXmlValidator 只能证明"包结构合法"，证明不了"标准解析器能按语义读出内容"。
/// 本类用 SDK 的对象模型读取文本、表格、图片、备注与形状填充，
/// 能发现"结构合法但关系断链/内容放错位置"这类缺陷（例如备注部件没挂到幻灯片上）。
/// </remarks>
public class PptCrossCompatibilityTests
{
    private static Byte[] Build(Action<PptxWriter> build)
    {
        using var w = new PptxWriter();
        build(w);
        using var ms = new MemoryStream();
        w.Save(ms);
        return ms.ToArray();
    }

    /// <summary>最小合法 PNG（1x1 透明像素）</summary>
    private static Byte[] MinimalPng { get; } = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8AAAwAB/AF+p0n3AAAAAElFTkSuQmCC");

    /// <summary>收集幻灯片内全部 a:t 文本</summary>
    /// <param name="slidePart">幻灯片部件</param>
    /// <returns>拼接文本</returns>
    private static String CollectText(SlidePart slidePart)
    {
        var sb = new StringBuilder();
        foreach (var t in slidePart.Slide.Descendants<A.Text>()) sb.Append(t.Text);
        return sb.ToString();
    }

    [Fact, DisplayName("P1 SDK 读取幻灯片文本")]
    public void P1_SDK_ReadsSlideText()
    {
        var bytes = Build(w =>
        {
            w.AddSlide(0);
            w.AddTextBox(0, "第一页标题", 2, 2, 20, 3, 28, true);
            w.AddSlide(1);
            w.AddTextBox(1, "第二页正文", 2, 2, 20, 3, 20, false);
        });

        using var ms = new MemoryStream(bytes);
        using var ppt = PresentationDocument.Open(ms, false);
        var parts = ppt.PresentationPart!.SlideParts.ToList();
        Assert.Equal(2, parts.Count);
        Assert.Contains("第一页标题", CollectText(parts[0]));
        Assert.Contains("第二页正文", CollectText(parts[1]));
    }

    [Fact, DisplayName("P2 SDK 读取表格单元格文本")]
    public void P2_SDK_ReadsTable()
    {
        var bytes = Build(w =>
        {
            w.AddSlide(0);
            w.AddTable(0, [["甲", "乙"], ["1", "2"]], 2, 2, 16, true);
        });

        using var ms = new MemoryStream(bytes);
        using var ppt = PresentationDocument.Open(ms, false);
        var table = ppt.PresentationPart!.SlideParts.First().Slide.Descendants<A.Table>().Single();
        var rows = table.Elements<A.TableRow>().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(2, rows[0].Elements<A.TableCell>().Count());
        var cells = rows.SelectMany(e => e.Elements<A.TableCell>())
            .Select(e => String.Concat(e.Descendants<A.Text>().Select(t => t.Text)))
            .ToList();
        Assert.Equal(["甲", "乙", "1", "2"], cells);
    }

    [Fact, DisplayName("P3 SDK 读取图片关系")]
    public void P3_SDK_ReadsPicture()
    {
        var bytes = Build(w =>
        {
            w.AddSlide(0);
            w.AddImage(0, MinimalPng, "png", 2, 2, 8, 6);
        });

        using var ms = new MemoryStream(bytes);
        using var ppt = PresentationDocument.Open(ms, false);
        var slidePart = ppt.PresentationPart!.SlideParts.First();

        var blip = slidePart.Slide.Descendants<A.Blip>().FirstOrDefault();
        Assert.NotNull(blip);
        var embedId = blip!.Embed!.Value!;
        var imagePart = (ImagePart)slidePart.GetPartById(embedId);
        using var data = imagePart.GetStream();
        Assert.True(data.Length > 0);
        Assert.Equal("image/png", imagePart.ContentType);
    }

    [Fact, DisplayName("P4 SDK 读取演讲者备注")]
    public void P4_SDK_ReadsNotes()
    {
        var bytes = Build(w =>
        {
            w.AddSlide(0);
            w.SetNotes(0, "演讲者备注内容");
        });

        using var ms = new MemoryStream(bytes);
        using var ppt = PresentationDocument.Open(ms, false);
        var slidePart = ppt.PresentationPart!.SlideParts.First();

        // 备注必须通过幻灯片的关系可达，否则 PowerPoint 侧看不到
        var notesPart = slidePart.NotesSlidePart;
        Assert.NotNull(notesPart);
        Assert.Contains("演讲者备注内容", String.Concat(notesPart!.NotesSlide.Descendants<A.Text>().Select(t => t.Text)));

        // notesSlide 必须挂着 notesMaster，缺母版会被判为损坏
        Assert.NotNull(notesPart.NotesMasterPart);
    }

    [Fact, DisplayName("P5 SDK 读取形状填充色")]
    public void P5_SDK_ReadsShapeFill()
    {
        var bytes = Build(w =>
        {
            w.AddSlide(0);
            w.AddShape(0, "rect", 2, 2, 8, 4, "00AAFF");
        });

        using var ms = new MemoryStream(bytes);
        using var ppt = PresentationDocument.Open(ms, false);
        var shape = ppt.PresentationPart!.SlideParts.First().Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        var color = shape.Descendants<A.RgbColorModelHex>().FirstOrDefault();
        Assert.NotNull(color);
        Assert.Equal("00AAFF", color!.Val!.Value);
    }
}
