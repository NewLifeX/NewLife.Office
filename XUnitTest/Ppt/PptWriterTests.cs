using System.ComponentModel;
using System.Text;
using NewLife.Office.Ole2;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PptWriter .ppt 二进制格式写入器单元测试</summary>
/// <remarks>
/// 写入后使用 <see cref="PptReader"/> 读回验证（自闭环），
/// 并校验 OLE2 容器包含 "PowerPoint Document" 与 "Current User" 两个流。
/// </remarks>
public class PptWriterTests
{
    /// <summary>写入器 → 读回器（内存流）</summary>
    private static PptReader RoundTrip(Action<PptWriter> build)
    {
        using var writer = new PptWriter();
        build(writer);
        var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;
        return new PptReader(ms);
    }

    [Fact, DisplayName("PPT—单张幻灯片写读往返")]
    public void WriteAndRead_SingleSlide()
    {
        using var reader = RoundTrip(w => w.AddSlide("Hello PPT"));
        Assert.Equal(1, reader.SlideCount);
        Assert.Contains("Hello PPT", reader.GetSlideText(0));
    }

    [Fact, DisplayName("PPT—多张幻灯片写读往返")]
    public void WriteAndRead_MultipleSlides()
    {
        using var reader = RoundTrip(w =>
        {
            w.AddSlide("第一张", "内容 A");
            w.AddSlide("第二张");
            w.AddSlide("第三张", "内容 B", "内容 C");
        });
        Assert.Equal(3, reader.SlideCount);
        Assert.Contains("第一张", reader.GetSlideText(0));
        Assert.Contains("内容 A", reader.GetSlideText(0));
        Assert.Contains("第二张", reader.GetSlideText(1));
        Assert.Contains("内容 B", reader.GetSlideText(2));
        Assert.Contains("内容 C", reader.GetSlideText(2));
    }

    [Fact, DisplayName("PPT—多行文本保留段落结构")]
    public void WriteAndRead_MultiLine()
    {
        using var reader = RoundTrip(w => w.AddSlide("标题", "正文第一行", "正文第二行"));
        var text = reader.GetSlideText(0);
        Assert.Contains("标题", text);
        Assert.Contains("正文第一行", text);
        Assert.Contains("正文第二行", text);
        // 段落以换行分隔
        Assert.Contains("\n", text);
    }

    [Fact, DisplayName("PPT—中文文本写读往返")]
    public void WriteAndRead_ChineseText()
    {
        using var reader = RoundTrip(w => w.AddSlide("智能制造系统商业计划书", "核心优势", "零依赖 · MIT 许可"));
        Assert.Contains("智能制造系统商业计划书", reader.GetSlideText(0));
        Assert.Contains("核心优势", reader.GetSlideText(0));
        Assert.Contains("零依赖 · MIT 许可", reader.GetSlideText(0));
    }

    [Fact, DisplayName("PPT—超长文本自动拆分续记录并完整读回")]
    public void WriteAndRead_LongText()
    {
        var longText = new String('中', 4200) + "尾部标记";
        using var reader = RoundTrip(w => w.AddSlide(longText));
        Assert.Equal(longText, reader.GetSlideText(0));
    }

    [Fact, DisplayName("PPT—空演示文稿可写入且 SlideCount=0")]
    public void Write_Empty()
    {
        using var reader = RoundTrip(_ => { });
        Assert.Equal(0, reader.SlideCount);
    }

    [Fact, DisplayName("PPT—OLE2 容器含 PowerPoint Document 与 Current User 流")]
    public void Save_ContainsBothStreams()
    {
        using var writer = new PptWriter();
        writer.AddSlide("标题");
        var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var doc = CfbDocument.Open(ms, leaveOpen: true);
        var ppt = doc.GetStreamData("PowerPoint Document");
        var user = doc.GetStreamData("Current User");
        Assert.NotNull(ppt);
        Assert.True(ppt!.Length > 0, "PowerPoint Document 流不能为空");
        Assert.NotNull(user);
        Assert.True(user!.Length >= 28, "Current User 流至少 28 字节");
    }

    [Fact, DisplayName("PPT—保存到文件路径可读回")]
    public void Save_ToPath()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".ppt");
        try
        {
            using (var writer = new PptWriter())
            {
                writer.Author = "单元测试";
                writer.AddSlide("文件保存测试");
                writer.Save(path);
            }
            Assert.True(File.Exists(path));
            using var reader = new PptReader(path);
            Assert.Equal(1, reader.SlideCount);
            Assert.Contains("文件保存测试", reader.GetSlideText(0));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("PPT—ToBytes 与 Save 结果一致")]
    public void ToBytes_MatchesSave()
    {
        using var writer = new PptWriter();
        writer.AddSlide("A", "B");

        var bytes = writer.ToBytes();
        using var ms = new MemoryStream();
        writer.Save(ms);

        Assert.True(bytes.Length > 0);
        Assert.Equal(bytes.Length, ms.Length);
        Assert.True(bytes.AsSpan().SequenceEqual(ms.ToArray()), "ToBytes 与 Save 输出不一致");
    }
}
