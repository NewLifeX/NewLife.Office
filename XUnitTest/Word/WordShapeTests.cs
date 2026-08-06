using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>Word 自选图形创建测试（W12，对标 Open XML SDK/GemBox）</summary>
/// <remarks>
/// 覆盖 DrawingML 形状（矩形/椭圆/直线）的写入与 document.xml 结构验证。
/// </remarks>
public class WordShapeTests
{
    /// <summary>生成含形状的 docx 并返回 document.xml 文本</summary>
    private static (Byte[] Docx, String DocumentXml) BuildDocx(Action<WordWriter> build)
    {
        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
        {
            build(writer);
            writer.Save(ms);
        }
        var bytes = ms.ToArray();
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entry = za.GetEntry("word/document.xml");
        Assert.NotNull(entry);
        using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
        return (bytes, sr.ReadToEnd());
    }

    [Fact, DisplayName("自选图形：矩形/椭圆/直线写入")]
    public void Shape_WriteBasics()
    {
        var (_, xml) = BuildDocx(w =>
        {
            w.AppendParagraph("形状测试");
            w.AppendShape(WordShape.Rect(4, 3, "FF0000"));
            w.AppendShape(WordShape.Ellipse(3, 3, "00FF00"));
            w.AppendShape(WordShape.Line(6, "0000FF", 2));
        });

        // wps:wsp 形状存在
        Assert.Contains("wordprocessingShape", xml);
        Assert.Contains("wps:wsp", xml);
        // 三种几何类型
        Assert.Contains("prst=\"rect\"", xml);
        Assert.Contains("prst=\"ellipse\"", xml);
        Assert.Contains("prst=\"line\"", xml);
        // 填充色
        Assert.Contains("val=\"FF0000\"", xml);
        Assert.Contains("val=\"00FF00\"", xml);
        // 线宽（2 磅 = 25400 EMU）
        Assert.Contains("w=\"25400\"", xml);
    }

    [Fact, DisplayName("自选图形：形状文本与无填充")]
    public void Shape_TextAndNoFill()
    {
        var (_, xml) = BuildDocx(w =>
        {
            var shape = WordShape.Rect(5, 2, "336699");
            shape.Text = "形状内文本";
            w.AppendShape(shape);
            w.AppendShape(new WordShape { Type = "triangle", WidthCm = 3, HeightCm = 3 });
        });

        Assert.Contains("<wps:txbx>", xml);
        Assert.Contains("形状内文本", xml);
        Assert.Contains("prst=\"triangle\"", xml);
        Assert.Contains("<a:noFill/>", xml); // 未设置填充色 → noFill
    }

    [Fact, DisplayName("自选图形：参数校验")]
    public void Shape_Validation()
    {
        using var writer = new WordWriter();
        Assert.Throws<ArgumentNullException>(() => writer.AppendShape(null!));
    }

    [Fact, DisplayName("自选图形：可被 WordReader 读回不抛异常")]
    public void Shape_ReadBack()
    {
        var (bytes, _) = BuildDocx(w =>
        {
            w.AppendShape(WordShape.Rect(4, 3, "FF0000"));
        });
        using var ms = new MemoryStream(bytes);
        var reader = new WordReader(ms);
        var doc = reader.ReadDocument();
        Assert.NotNull(doc);
        // WordReader 对 drawing 段落保留原始 XML 透传，读取不抛异常即视为通过
        Assert.NotNull(doc.DocumentXml);
    }
}
