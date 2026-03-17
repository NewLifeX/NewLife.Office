using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>PptReader .ppt 二进制格式读取器单元测试</summary>
public class PptReaderTests
{
    // ─── PPT 记录构建辅助 ────────────────────────────────────────────────

    private static Byte[] LE2(UInt16 v) => new Byte[] { (Byte)(v & 0xFF), (Byte)(v >> 8) };
    private static Byte[] LE4(UInt32 v) =>
        new Byte[] { (Byte)(v & 0xFF), (Byte)((v >> 8) & 0xFF), (Byte)((v >> 16) & 0xFF), (Byte)(v >> 24) };

    private static Byte[] Concat(params Byte[][] parts)
    {
        var total = 0;
        foreach (var p in parts) total += p.Length;
        var buf = new Byte[total];
        var pos = 0;
        foreach (var p in parts) { Array.Copy(p, 0, buf, pos, p.Length); pos += p.Length; }
        return buf;
    }

    // PPT 记录头：recVerType(2) + recType(2) + recLen(4)
    private static Byte[] RecHeader(UInt16 recVer, UInt16 recInstance, UInt16 recType, UInt32 recLen)
    {
        var verType = (UInt16)((recInstance << 4) | (recVer & 0x0F));
        return Concat(LE2(verType), LE2(recType), LE4(recLen));
    }

    // TextCharsAtom (0x03F2)：UTF-16LE 文本原子
    private static Byte[] TextCharsAtom(String text)
    {
        var textBytes = Encoding.Unicode.GetBytes(text);
        return Concat(RecHeader(0x00, 0x0000, 0x03F2, (UInt32)textBytes.Length), textBytes);
    }

    // TextBytesAtom (0x03F0)：ANSI 文本原子
    private static Byte[] TextBytesAtom(String text)
    {
        var textBytes = Encoding.ASCII.GetBytes(text);
        return Concat(RecHeader(0x00, 0x0000, 0x03F0, (UInt32)textBytes.Length), textBytes);
    }

    // SlideContainer (0x03EE)：幻灯片容器
    private static Byte[] SlideContainer(Byte[] body)
    {
        return Concat(RecHeader(0x0F, 0x0000, 0x03EE, (UInt32)body.Length), body);
    }

    // 将 PPT 流打包进 OLE2 容器
    private static Stream BuildPpt(Byte[] pptStream)
    {
        var doc = new CfbDocument();
        doc.Root.AddStream("PowerPoint Document", pptStream);
        return new MemoryStream(doc.ToBytes());
    }

    // ─── 测试 ─────────────────────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("SlideCount 返回 1（单幻灯片）")]
    public void SlideCount_SingleSlide_ReturnsOne()
    {
        var slide = SlideContainer(TextCharsAtom("Title"));
        var ppt = BuildPpt(slide);
        using var reader = new PptReader(ppt);
        Assert.Equal(1, reader.SlideCount);
    }

    [Fact, System.ComponentModel.DisplayName("SlideCount 返回正确幻灯片数量")]
    public void SlideCount_MultipleSlides()
    {
        var slide1 = SlideContainer(TextCharsAtom("Slide 1"));
        var slide2 = SlideContainer(TextCharsAtom("Slide 2"));
        var slide3 = SlideContainer(TextCharsAtom("Slide 3"));
        using var ppt = BuildPpt(Concat(slide1, slide2, slide3));
        using var reader = new PptReader(ppt);
        Assert.Equal(3, reader.SlideCount);
    }

    [Fact, System.ComponentModel.DisplayName("GetSlideText 返回正确文本（TextCharsAtom）")]
    public void GetSlideText_TextCharsAtom()
    {
        var slide = SlideContainer(TextCharsAtom("Hello PPT"));
        using var ppt = BuildPpt(slide);
        using var reader = new PptReader(ppt);
        Assert.Contains("Hello PPT", reader.GetSlideText(0));
    }

    [Fact, System.ComponentModel.DisplayName("GetSlideText 返回正确文本（TextBytesAtom）")]
    public void GetSlideText_TextBytesAtom()
    {
        var slide = SlideContainer(TextBytesAtom("Bytes text"));
        using var ppt = BuildPpt(slide);
        using var reader = new PptReader(ppt);
        Assert.Contains("Bytes text", reader.GetSlideText(0));
    }

    [Fact, System.ComponentModel.DisplayName("GetSlideText 合并一个幻灯片内多个文本块")]
    public void GetSlideText_MultipleTextAtoms_MergedWithNewlines()
    {
        var body = Concat(TextCharsAtom("Title"), TextCharsAtom("Body text"));
        var slide = SlideContainer(body);
        using var ppt = BuildPpt(slide);
        using var reader = new PptReader(ppt);
        var text = reader.GetSlideText(0);
        Assert.Contains("Title", text);
        Assert.Contains("Body text", text);
    }

    [Fact, System.ComponentModel.DisplayName("ReadAllSlides 返回所有幻灯片文本序列")]
    public void ReadAllSlides_ReturnsAllSlides()
    {
        var slide1 = SlideContainer(TextCharsAtom("Alpha"));
        var slide2 = SlideContainer(TextCharsAtom("Beta"));
        using var ppt = BuildPpt(Concat(slide1, slide2));
        using var reader = new PptReader(ppt);
        var texts = reader.ReadAllSlides().ToList();
        Assert.Equal(2, texts.Count);
        Assert.Contains("Alpha", texts[0]);
        Assert.Contains("Beta", texts[1]);
    }

    [Fact, System.ComponentModel.DisplayName("PPT 流不含 SlideContainer 时 SlideCount = 0")]
    public void NoSlides_SlideCount_Zero()
    {
        // 构造一个含有合法记录但不含 SlideContainer 的 PPT 流
        // 用一个普通原子记录（DocumentContainer 以外的类型，不含文本）
        var dummyAtom = Concat(RecHeader(0x00, 0, 0x03E8, 0));  // recLen=0
        using var ppt = BuildPpt(dummyAtom);
        using var reader = new PptReader(ppt);
        Assert.Equal(0, reader.SlideCount);
    }

    [Fact, System.ComponentModel.DisplayName("GetSlideText 越界抛 ArgumentOutOfRangeException")]
    public void GetSlideText_OutOfRange_Throws()
    {
        var slide = SlideContainer(TextCharsAtom("Data"));
        using var ppt = BuildPpt(slide);
        using var reader = new PptReader(ppt);
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetSlideText(1));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetSlideText(-1));
    }

    [Fact, System.ComponentModel.DisplayName("非 OLE2 格式抛出 InvalidDataException")]
    public void InvalidOle2_Throws()
    {
        var fakeData = new Byte[512];
        using var ms = new MemoryStream(fakeData);
        Assert.Throws<InvalidDataException>(() => new PptReader(ms));
    }

    [Fact, System.ComponentModel.DisplayName("嵌套容器内的文本也被提取")]
    public void NestedContainer_TextExtracted()
    {
        // 在 SlideContainer 内套一个普通容器，文本依然属于该幻灯片
        var innerBody = TextCharsAtom("Nested text");
        var innerContainer = Concat(RecHeader(0x0F, 0, 0x03F9, (UInt32)innerBody.Length), innerBody);
        var slideBody = Concat(innerContainer);
        var slide = SlideContainer(slideBody);
        using var ppt = BuildPpt(slide);
        using var reader = new PptReader(ppt);
        Assert.Equal(1, reader.SlideCount);
        Assert.Contains("Nested text", reader.GetSlideText(0));
    }
}
