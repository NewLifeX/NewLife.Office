using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>
/// 读取保真测试：三态格式（w:val="0" 显式关闭）、Run 多子元素文本、
/// 列表 numId→numFmt 映射判定、新增格式属性（高亮/小型大写/隐藏/东亚字体/语言）。
/// </summary>
public class ReaderFidelityTests
{
    #region 构造 docx 辅助
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>构建最小 docx 字节（可附加 styles.xml / numbering.xml）</summary>
    private static Byte[] BuildDocx(String bodyInnerXml, String? numberingXml = null, String? stylesXml = null)
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
            if (stylesXml != null) WriteEntry(za, "word/styles.xml", stylesXml);
            if (numberingXml != null) WriteEntry(za, "word/numbering.xml", numberingXml);
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    /// <summary>读取首段首 Run</summary>
    private static (Paragraph Para, Run Run) ReadFirstRun(String bodyInnerXml, String? numberingXml = null, String? stylesXml = null)
    {
        var bytes = BuildDocx(bodyInnerXml, numberingXml, stylesXml);
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();
        var para = Assert.Single(doc.Elements, e => e.Type == ElementType.Paragraph).Paragraph!;
        return (para, Assert.Single(para.Runs));
    }
    #endregion

    #region 三态格式解析
    [Fact(DisplayName = "三态—w:b w:val=0 显式关闭粗体")]
    public void ThreeState_BoldExplicitOff()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:rPr><w:b w:val=\"0\"/></w:rPr><w:t>abc</w:t></w:r></w:p>");
        Assert.False(run.Properties!.Bold);
        Assert.Equal("abc", run.Text);
    }

    [Fact(DisplayName = "三态—w:b 无 val 默认开启")]
    public void ThreeState_BoldOn()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>abc</w:t></w:r></w:p>");
        Assert.True(run.Properties!.Bold);
    }

    [Fact(DisplayName = "三态—w:b w:val=1 开启")]
    public void ThreeState_BoldValOne()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:rPr><w:b w:val=\"1\"/></w:rPr><w:t>abc</w:t></w:r></w:p>");
        Assert.True(run.Properties!.Bold);
    }

    [Fact(DisplayName = "三态—w:i w:val=false 显式关闭斜体")]
    public void ThreeState_ItalicExplicitOff()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:rPr><w:i w:val=\"false\"/></w:rPr><w:t>abc</w:t></w:r></w:p>");
        Assert.False(run.Properties!.Italic);
    }

    [Fact(DisplayName = "三态—w:strike w:val=0 显式关闭删除线")]
    public void ThreeState_StrikeExplicitOff()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:rPr><w:strike w:val=\"0\"/></w:rPr><w:t>abc</w:t></w:r></w:p>");
        Assert.False(run.Properties!.Strikethrough);
    }

    [Fact(DisplayName = "三态—w:u w:val=none 显式取消下划线")]
    public void ThreeState_UnderlineNone()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:rPr><w:u w:val=\"none\"/></w:rPr><w:t>abc</w:t></w:r></w:p>");
        Assert.False(run.Properties!.Underline);
    }

    [Fact(DisplayName = "三态—w:vertAlign baseline 取消上下标")]
    public void ThreeState_VertAlignBaseline()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:rPr><w:vertAlign w:val=\"baseline\"/></w:rPr><w:t>abc</w:t></w:r></w:p>");
        Assert.False(run.Properties!.Superscript);
        Assert.False(run.Properties!.Subscript);
    }
    #endregion

    #region Run 多子元素文本
    [Fact(DisplayName = "Run文本—多个 w:t 拼接")]
    public void RunText_MultipleWt()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:t>ab</w:t><w:t>cd</w:t></w:r></w:p>");
        Assert.Equal("abcd", run.Text);
    }

    [Fact(DisplayName = "Run文本—w:tab 转制表符")]
    public void RunText_Tab()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>");
        Assert.Equal("a\tb", run.Text);
    }

    [Fact(DisplayName = "Run文本—w:br 换行 / w:cr 回车")]
    public void RunText_BrAndCr()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t><w:cr/><w:t>c</w:t></w:r></w:p>");
        Assert.Equal("a\nb\nc", run.Text);
    }

    [Fact(DisplayName = "Run文本—noBreakHyphen 转连字符")]
    public void RunText_NoBreakHyphen()
    {
        var (_, run) = ReadFirstRun("<w:p><w:r><w:t>a</w:t><w:noBreakHyphen/><w:t>b</w:t></w:r></w:p>");
        Assert.Equal("a-b", run.Text);
    }
    #endregion

    #region 列表 numId 映射判定
    private const String NumberingXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + $"<w:numbering xmlns:w=\"{W}\">"
        + "<w:abstractNum w:abstractNumId=\"3\">"
        + "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"•\"/></w:lvl>"
        + "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"–\"/></w:lvl>"
        + "</w:abstractNum>"
        + "<w:abstractNum w:abstractNumId=\"4\">"
        + "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl>"
        + "</w:abstractNum>"
        + "<w:num w:numId=\"7\"><w:abstractNumId w:val=\"3\"/></w:num>"
        + "<w:num w:numId=\"8\"><w:abstractNumId w:val=\"4\"/></w:num>"
        + "</w:numbering>";

    [Fact(DisplayName = "列表—numId 映射 bullet 判定为无序列表")]
    public void List_NumIdBullet()
    {
        var body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"7\"/></w:numPr></w:pPr>"
            + "<w:r><w:t>项目一</w:t></w:r></w:p>";
        var (para, _) = ReadFirstRun(body, NumberingXml);
        Assert.True(para.IsBullet);
        Assert.False(para.IsOrderedList);
        Assert.Equal(7, para.NumId);
        Assert.Equal("bullet", para.ListFormat);
        Assert.Equal(0, para.ListLevel);
    }

    [Fact(DisplayName = "列表—numId 映射 decimal 判定为有序列表")]
    public void List_NumIdOrdered()
    {
        var body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"8\"/></w:numPr></w:pPr>"
            + "<w:r><w:t>第一项</w:t></w:r></w:p>";
        var (para, _) = ReadFirstRun(body, NumberingXml);
        Assert.False(para.IsBullet);
        Assert.True(para.IsOrderedList);
        Assert.Equal(8, para.NumId);
        Assert.Equal("decimal", para.ListFormat);
    }

    [Fact(DisplayName = "列表—numId=0 无编号")]
    public void List_NumIdZero()
    {
        var body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"0\"/></w:numPr></w:pPr>"
            + "<w:r><w:t>正文</w:t></w:r></w:p>";
        var (para, _) = ReadFirstRun(body, NumberingXml);
        Assert.False(para.IsBullet);
        Assert.False(para.IsOrderedList);
        Assert.Equal(0, para.NumId);
    }

    [Fact(DisplayName = "列表—numId 二级级别判定 bullet")]
    public void List_NumIdLevel1()
    {
        var body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"7\"/></w:numPr></w:pPr>"
            + "<w:r><w:t>子项</w:t></w:r></w:p>";
        var (para, _) = ReadFirstRun(body, NumberingXml);
        Assert.True(para.IsBullet);
        Assert.Equal(1, para.ListLevel);
    }
    #endregion

    #region 新增格式属性
    [Fact(DisplayName = "格式—高亮/小型大写/全大写/隐藏/语言/东亚字体")]
    public void Format_NewProperties()
    {
        var body = "<w:p><w:r><w:rPr>"
            + "<w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:eastAsia=\"宋体\"/>"
            + "<w:highlight w:val=\"yellow\"/>"
            + "<w:smallCaps/>"
            + "<w:caps w:val=\"0\"/>"
            + "<w:vanish/>"
            + "<w:lang w:val=\"zh-CN\"/>"
            + "</w:rPr><w:t>abc</w:t></w:r></w:p>";
        var (_, run) = ReadFirstRun(body);
        var p = run.Properties!;
        Assert.Equal("Arial", p.FontName);
        Assert.Equal("宋体", p.EastAsiaFontName);
        Assert.Equal("yellow", p.HighlightColor);
        Assert.True(p.SmallCaps);
        Assert.False(p.AllCaps);
        Assert.True(p.Hidden);
        Assert.Equal("zh-CN", p.Language);
    }
    #endregion

    #region 空段落保留
    [Fact(DisplayName = "保真—空段落保留在模型中（读入写出结构完整）")]
    public void Fidelity_EmptyParagraphKept()
    {
        var bytes = BuildDocx("<w:p><w:r><w:t>一</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>二</w:t></w:r></w:p>");
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();
        Assert.Equal(3, doc.Elements.Count);
        Assert.Equal(ElementType.Paragraph, doc.Elements[1].Type);
        Assert.Empty(doc.Elements[1].Paragraph!.Runs);
    }

    [Fact(DisplayName = "保真—含书签的空段落保留")]
    public void Fidelity_BookmarkParagraphKept()
    {
        var bytes = BuildDocx("<w:p><w:bookmarkStart w:id=\"0\" w:name=\"锚点\"/><w:bookmarkEnd w:id=\"0\"/></w:p>");
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();
        var para = Assert.Single(doc.Elements).Paragraph!;
        Assert.Equal("锚点", para.BookmarkName);
    }
    #endregion

    #region 三态写回
    [Fact(DisplayName = "写回—Bold=false 输出 w:val=0")]
    public void Write_ExplicitOff()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter())
            {
                var para = w.AppendParagraph("abc", ParagraphStyle.Normal, new RunProperties { Bold = false });
                w.Save(path);
            }
            // 读取生成的 document.xml 验证 w:val=0 写回
            using var za = ZipFile.OpenRead(path);
            var entry = za.GetEntry("word/document.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("<w:b w:val=\"0\"/>", xml);

            // 读回模型验证三态保持
            using var reader = new WordReader(path);
            var doc = reader.ReadDocument();
            var run = doc.Elements[0].Paragraph!.Runs[0];
            Assert.False(run.Properties!.Bold);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion
}
