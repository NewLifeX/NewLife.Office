using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>图片插入 API 测试：AddPicture 系列、自然尺寸探测、透传模式下 media/rels/content-types 补齐</summary>
public class PictureApiTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>1×1 透明 PNG</summary>
    private const String Png1X1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";

    #region 工具
    private static Byte[] BuildDocx(String bodyInnerXml)
    {
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>");
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static Document Read(String bodyInnerXml)
    {
        using var ms = new MemoryStream(BuildDocx(bodyInnerXml));
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static Byte[] Save(Document doc)
    {
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.Save(ms, doc);
        return ms.ToArray();
    }

    private static Document ReadBack(Byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static Boolean HasEntry(Byte[] data, String path)
    {
        using var ms = new MemoryStream(data);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        return za.GetEntry(path) != null;
    }

    private static String ReadEntry(Byte[] data, String path)
    {
        using var ms = new MemoryStream(data);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(za.GetEntry(path)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    /// <summary>构造带页眉的 docx（读入后 <see cref="Document.Headers"/> 非空 → 保存走重建分支）</summary>
    private static Byte[] BuildDocxWithHeader()
    {
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\"><w:body>"
            + "<w:p><w:r><w:t>正文</w:t></w:r></w:p>"
            + "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rHdr1\"/>"
            + "<w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\"/></w:sectPr>"
            + "</w:body></w:document>";
        var headerXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:hdr xmlns:w=\"{W}\"><w:p><w:r><w:t>公司页眉</w:t></w:r></w:p></w:hdr>";
        var relsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rHdr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>"
            + "</Relationships>";

        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
                + "</Types>");
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/_rels/document.xml.rels", relsXml);
            WriteEntry(za, "word/header1.xml", headerXml);
        }
        return ms.ToArray();
    }
    #endregion

    #region 插入
    [Fact(DisplayName = "图片-透传文档插入后补齐 media/rels/content-types")]
    public void AddPicture_ReadDoc_PersistsParts()
    {
        var doc = Read("<w:p><w:r><w:t>正文</w:t></w:r></w:p>");
        doc.BeginEdit();

        var img = doc.AddPicture(Convert.FromBase64String(Png1X1), "png", altText: "示例图");
        Assert.Equal("rPic1", img.RelId);
        Assert.Equal("示例图", img.AltText);
        Assert.True(doc.Images.ContainsKey("rPic1"));

        var bytes = Save(doc);
        Assert.True(HasEntry(bytes, "word/media/rPic1.png"));
        Assert.Contains("Id=\"rPic1\"", ReadEntry(bytes, "word/_rels/document.xml.rels"));
        Assert.Contains("Extension=\"png\"", ReadEntry(bytes, "[Content_Types].xml"));

        var back = ReadBack(bytes);
        Assert.Contains(back.Elements, e => e.Type == ElementType.Image);
        Assert.True(back.Images.ContainsKey("rPic1"));
    }

    [Fact(DisplayName = "图片-页眉重建模式下插入图片仍补齐 rels 与 content-types")]
    public void AddPicture_WithHeaderRebuild_PersistsParts()
    {
        Document doc;
        using (var ms = new MemoryStream(BuildDocxWithHeader()))
        using (var reader = new WordReader(ms))
            doc = reader.ReadDocument();

        // 前提：有页眉 → 保存时走重建分支（[Content_Types].xml 与 document.xml.rels 会从透传集中移除）
        Assert.NotEmpty(doc.Headers);
        doc.BeginEdit();

        var img = doc.AddPicture(Convert.FromBase64String(Png1X1), "png");
        var bytes = Save(doc);

        Assert.True(HasEntry(bytes, $"word/media/{img.RelId}.png"));
        var rels = ReadEntry(bytes, "word/_rels/document.xml.rels");
        Assert.Contains($"Id=\"{img.RelId}\"", rels);
        Assert.Contains("Extension=\"png\"", ReadEntry(bytes, "[Content_Types].xml"));
        // 页眉关系仍需保留（重建不得把页眉一起丢掉）
        Assert.Contains("header", rels);

        var back = ReadBack(bytes);
        Assert.True(back.Images.ContainsKey(img.RelId));
        Assert.Contains(back.Elements, e => e.Type == ElementType.Image);
    }

    [Fact(DisplayName = "图片-只给宽度时按自然宽高比换算高度")]
    public void AddPicture_OnlyWidth_KeepsAspect()
    {
        var doc = new Document();
        var img = doc.AddPicture(Convert.FromBase64String(Png1X1), "png", width: Length.FromCm(4));

        // 1×1 像素图片：宽高比 1:1
        Assert.Equal(4, Math.Round(img.WidthEmu / Length.EmuPerCm, 3));
        Assert.Equal(4, Math.Round(img.HeightEmu / Length.EmuPerCm, 3));
    }

    [Fact(DisplayName = "图片-只给高度时按自然宽高比换算宽度")]
    public void AddPicture_OnlyHeight_KeepsAspect()
    {
        var doc = new Document();
        var img = doc.AddPicture(Convert.FromBase64String(Png1X1), "png", height: Length.FromCm(3));

        Assert.Equal(3, Math.Round(img.HeightEmu / Length.EmuPerCm, 3));
        Assert.Equal(3, Math.Round(img.WidthEmu / Length.EmuPerCm, 3));
    }

    [Fact(DisplayName = "图片-同时给宽高时不做换算")]
    public void AddPicture_BothSize_UsedAsIs()
    {
        var doc = new Document();
        var img = doc.AddPicture(Convert.FromBase64String(Png1X1), "png", Length.FromCm(5), Length.FromCm(2));

        Assert.Equal(5, Math.Round(img.WidthEmu / Length.EmuPerCm, 3));
        Assert.Equal(2, Math.Round(img.HeightEmu / Length.EmuPerCm, 3));
    }

    [Fact(DisplayName = "图片-Base64 data URI 可识别格式")]
    public void AddPictureBase64_DataUri()
    {
        var doc = new Document();
        var img = doc.AddPictureBase64($"data:image/png;base64,{Png1X1}");

        Assert.Equal("png", img.Extension);
        Assert.True(doc.Images.ContainsKey(img.RelId));
    }

    [Fact(DisplayName = "图片-多次插入分配到不同关系 ID")]
    public void AddPicture_Multiple_UniqueRelIds()
    {
        var doc = new Document();
        var a = doc.AddPicture(Convert.FromBase64String(Png1X1), "png");
        var b = doc.AddPicture(Convert.FromBase64String(Png1X1), "png");

        Assert.NotEqual(a.RelId, b.RelId);
        Assert.Equal(2, doc.Images.Count);
    }

    [Fact(DisplayName = "图片-空数据抛异常")]
    public void AddPicture_EmptyData_Throws()
    {
        var doc = new Document();
        Assert.Throws<ArgumentNullException>(() => doc.AddPicture((Byte[])null!));
        Assert.Throws<ArgumentException>(() => doc.AddPicture([]));
    }

    [Fact(DisplayName = "图片-非法 data URI 抛 FormatException")]
    public void AddPictureBase64_InvalidDataUri_Throws()
    {
        var doc = new Document();
        Assert.Throws<FormatException>(() => doc.AddPictureBase64("data:image/png;AAAA"));
    }
    #endregion

    #region 元信息探测
    [Fact(DisplayName = "图片探测-PNG 1×1 识别正确")]
    public void ImageInfo_Png()
    {
        var data = Convert.FromBase64String(Png1X1);
        Assert.Equal("png", ImageInfo.DetectExtension(data));
        Assert.True(ImageInfo.TryGetPixelSize(data, out var w, out var h, out _));
        Assert.Equal(1, w);
        Assert.Equal(1, h);
    }

    [Fact(DisplayName = "图片探测-GIF 尺寸小端解析")]
    public void ImageInfo_Gif()
    {
        var data = new Byte[32];
        data[0] = (Byte)'G'; data[1] = (Byte)'I'; data[2] = (Byte)'F'; data[3] = (Byte)'8'; data[4] = (Byte)'9'; data[5] = (Byte)'a';
        data[6] = 100; data[7] = 0;    // width = 100
        data[8] = 50; data[9] = 0;     // height = 50

        Assert.Equal("gif", ImageInfo.DetectExtension(data));
        Assert.True(ImageInfo.TryGetPixelSize(data, out var w, out var h, out _));
        Assert.Equal(100, w);
        Assert.Equal(50, h);
    }

    [Fact(DisplayName = "图片探测-BMP 尺寸解析（负高度取绝对值）")]
    public void ImageInfo_Bmp()
    {
        var data = new Byte[40];
        data[0] = (Byte)'B'; data[1] = (Byte)'M';
        data[18] = 0x80; data[19] = 0x02; data[20] = 0; data[21] = 0;         // width = 640
        data[22] = 0xCE; data[23] = 0xFF; data[24] = 0xFF; data[25] = 0xFF; // height = -50

        Assert.Equal("bmp", ImageInfo.DetectExtension(data));
        Assert.True(ImageInfo.TryGetPixelSize(data, out var w, out var h, out _));
        Assert.Equal(640, w);
        Assert.Equal(50, h);
    }

    [Fact(DisplayName = "图片探测-未知格式返回空")]
    public void ImageInfo_Unknown()
    {
        Assert.Null(ImageInfo.DetectExtension([1, 2, 3, 4, 5]));
        Assert.False(ImageInfo.TryGetPixelSize([1, 2, 3, 4, 5], out _, out _, out _));
    }
    #endregion
}
