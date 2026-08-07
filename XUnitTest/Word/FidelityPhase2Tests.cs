using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>
/// 阶段2 保真测试：表格 vMerge/宽度/行高、SVG/anchor 图片、富文本页眉页脚、多节支持。
/// </summary>
public class FidelityPhase2Tests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const String WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    #region 构造 docx 辅助
    /// <summary>构建 docx 字节，可选页眉/页脚部件与 document.xml.rels 内容</summary>
    private static Byte[] BuildDocx(String bodyInnerXml, String? headerXml = null, String? footerXml = null,
        String? docRelsXml = null, String? numberingXml = null, String? stylesXml = null, String? media = null)
    {
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:wp=\"{WP}\" xmlns:a=\"{A}\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\"><w:body>{bodyInnerXml}</w:body></w:document>";

        if (docRelsXml == null)
            docRelsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>";

        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var ct = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
                + "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>"
                + "<Default Extension=\"svg\" ContentType=\"image/svg+xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>";
            if (headerXml != null)
                ct += "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>";
            if (footerXml != null)
                ct += "<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>";
            ct += "</Types>";
            WriteEntry(za, "[Content_Types].xml", ct);

            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/_rels/document.xml.rels", docRelsXml);
            if (headerXml != null) WriteEntry(za, "word/header1.xml", headerXml);
            if (footerXml != null) WriteEntry(za, "word/footer1.xml", footerXml);
            if (stylesXml != null) WriteEntry(za, "word/styles.xml", stylesXml);
            if (numberingXml != null) WriteEntry(za, "word/numbering.xml", numberingXml);
            if (media != null) WriteEntry(za, "word/media/image1.svg", media);
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static Document Read(String bodyInnerXml, String? headerXml = null, String? footerXml = null,
        String? docRelsXml = null, String? numberingXml = null, String? stylesXml = null, String? media = null)
    {
        var bytes = BuildDocx(bodyInnerXml, headerXml, footerXml, docRelsXml, numberingXml, stylesXml, media);
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }
    #endregion

    #region 表格读取
    [Fact(DisplayName = "表格—vMerge restart/continue 区分 + 单元格宽度")]
    public void Table_VMergeRestartContinue()
    {
        var body = "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>"
            + "<w:tr>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>起点</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>合并起点</w:t></w:r></w:p></w:tc>"
            + "</w:tr>"
            + "<w:tr>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>下</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
            + "</w:tr></w:tbl>";
        var doc = Read(body);
        var el = Assert.Single(doc.Elements, e => e.Type == ElementType.Table);
        var rows = el.TableRows!;
        Assert.Equal(2, rows.Count);
        Assert.Equal(2500, rows[0][0].Width);
        Assert.Equal(-1, rows[0][1].RowSpan);  // restart
        Assert.Equal(0, rows[1][1].RowSpan);   // continue
    }

    [Fact(DisplayName = "表格—Table 富模型写回（行高/表头/单元格宽度）")]
    public void Table_ModelWriteBack()
    {
        var table = new Table
        {
            FirstRowHeader = true,
            Width = 6000,
            ColumnWidths = [3000, 3000],
        };
        table.Rows.Add(new TableRow
        {
            IsHeader = true,
            Height = 500,
            Cells =
            [
                new Cell { Width = 3000, Paragraphs = { new Paragraph { Runs = { new Run { Text = "列A" } } } } },
                new Cell { Width = 3000, Paragraphs = { new Paragraph { Runs = { new Run { Text = "列B" } } } } },
            ],
        });
        table.Rows.Add(new TableRow
        {
            Cells =
            [
                new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "a1" } } } } },
                new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "b1" } } } } },
            ],
        });

        var doc = new Document { DocumentXml = null };
        doc.Elements.Add(new Element { Type = ElementType.Table, Table = table });

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter()) w.Save(path, doc);

            // 验证生成的 document.xml 含 trHeight/tblHeader/tcW
            using var za = ZipFile.OpenRead(path);
            using var sr = new StreamReader(za.GetEntry("word/document.xml")!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("<w:trHeight w:val=\"500\"", xml);
            Assert.Contains("<w:tblHeader/>", xml);
            Assert.Contains("<w:tcW w:w=\"3000\"", xml);

            // 读回模型验证
            using var reader = new WordReader(path);
            var read = reader.ReadDocument();
            var el = Assert.Single(read.Elements, e => e.Type == ElementType.Table);
            Assert.NotNull(el.Table);
            Assert.Equal(500, el.Table!.Rows[0].Height);
            Assert.True(el.Table.Rows[0].IsHeader);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion

    #region 图片解析
    [Fact(DisplayName = "图片—SVG asvg:svgBlip 解析")]
    public void Image_SvgParse()
    {
        var body = "<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">"
            + "<wp:extent cx=\"3600000\" cy=\"2700000\"/>"
            + "<wp:docPr id=\"1\" name=\"svg1\"/>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
            + "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"\"/><pic:cNvPicPr/></pic:nvPicPr>"
            + "<pic:blipFill><a:blip r:embed=\"rIdImg\"><asvg:svgBlip r:embed=\"rIdImg\"/></a:blip>"
            + "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
            + "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"3600000\" cy=\"2700000\"/></a:xfrm>"
            + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
            + "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rIdImg\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.svg\"/>"
            + "</Relationships>";
        var doc = Read(body, docRelsXml: rels, media: "<svg xmlns=\"http://www.w3.org/2000/svg\"/>");

        var img = Assert.Single(doc.Images.Values);
        Assert.Equal("svg", img.Extension);
        var el = Assert.Single(doc.Elements, e => e.Type == ElementType.Image);
        Assert.True(el.Image!.IsSvg);
        Assert.Equal("svg", el.Image.Extension);
    }

    [Fact(DisplayName = "图片—wp:anchor 浮动位置与环绕解析")]
    public void Image_AnchorParse()
    {
        var body = "<w:p><w:r><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251658240\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
            + "<wp:simplePos x=\"0\" y=\"0\"/>"
            + "<wp:positionH relativeFrom=\"column\"><wp:align>center</wp:align></wp:positionH>"
            + "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>457200</wp:posOffset></wp:positionV>"
            + "<wp:wrapSquare wrapText=\"bothSides\"/>"
            + "<wp:extent cx=\"3600000\" cy=\"2700000\"/>"
            + "<wp:docPr id=\"2\" name=\"img2\" descr=\"logo\"/>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
            + "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"\"/><pic:cNvPicPr/></pic:nvPicPr>"
            + "<pic:blipFill><a:blip r:embed=\"rIdImg2\"/>"
            + "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
            + "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"3600000\" cy=\"2700000\"/></a:xfrm>"
            + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
            + "</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>";
        var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rIdImg2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image2.png\"/>"
            + "</Relationships>";
        // 需在 BuildDocx 中支持 png 媒体 —— 用 svg 代替无法匹配，改用内联构造
        var bytes = BuildDocx(body, docRelsXml: rels, media: "<png/>");
        using (var ms = new MemoryStream(bytes))
        using (var reader = new WordReader(ms))
        {
            var doc = reader.ReadDocument();
            var el = Assert.Single(doc.Elements, e => e.Type == ElementType.Image);
            var img = el.Image!;
            Assert.Equal("anchor", img.AnchorType);
            Assert.Equal("center", img.AnchorPosH);
            Assert.Equal(457200, img.AnchorOffsetY);
            Assert.Equal("square", img.Wrap);
            Assert.Equal("logo", img.AltText);
        }
    }
    #endregion

    #region 页眉页脚
    [Fact(DisplayName = "页眉—富文本页眉读取（default 类型）")]
    public void Header_RichRead()
    {
        var body = "<w:p><w:r><w:t>正文</w:t></w:r></w:p>"
            + "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rHdr1\"/>"
            + "<w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>";
        var headerXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:hdr xmlns:w=\"{W}\"><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>机密文件</w:t></w:r></w:p></w:hdr>";
        var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rHdr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>"
            + "</Relationships>";
        var doc = Read(body, headerXml: headerXml, docRelsXml: rels);

        Assert.Single(doc.Headers);
        var hdr = doc.Headers[0];
        Assert.Equal("default", hdr.Type);
        var para = Assert.Single(hdr.Elements).Paragraph!;
        Assert.Equal("机密文件", Assert.Single(para.Runs).Text);
        Assert.Equal("机密文件", doc.HeaderText);
    }

    [Fact(DisplayName = "页眉—程序化富文本页眉写入后读回")]
    public void Header_RichWrite()
    {
        var doc = new Document { DocumentXml = null };
        doc.Headers.Add(new Header
        {
            Type = "default",
            Elements =
            [
                new Element
                {
                    Type = ElementType.Paragraph,
                    Paragraph = new Paragraph
                    {
                        Alignment = "center",
                        Runs = { new Run { Text = "公司内部资料", Properties = new RunProperties { Bold = true } } },
                    },
                },
            ],
        });
        doc.Footers.Add(new Footer
        {
            Type = "default",
            Elements =
            [
                new Element
                {
                    Type = ElementType.Paragraph,
                    Paragraph = new Paragraph
                    {
                        Runs = { new Run { Text = "第 " } },
                    },
                },
            ],
        });
        doc.Elements.Add(new Element
        {
            Type = ElementType.Paragraph,
            Paragraph = new Paragraph { Runs = { new Run { Text = "正文内容" } } },
        });

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter()) w.Save(path, doc);

            using var za = ZipFile.OpenRead(path);
            using var sr = new StreamReader(za.GetEntry("word/header1.xml")!.Open(), Encoding.UTF8);
            var headerXml = sr.ReadToEnd();
            Assert.Contains("公司内部资料", headerXml);
            Assert.NotNull(za.GetEntry("word/footer1.xml"));

            using var reader = new WordReader(path);
            var read = reader.ReadDocument();
            Assert.Single(read.Headers);
            Assert.Equal("default", read.Headers[0].Type);
            var para = Assert.Single(read.Headers[0].Elements).Paragraph!;
            Assert.Equal("公司内部资料", Assert.Single(para.Runs).Text);
            Assert.True(Assert.Single(para.Runs).Properties!.Bold);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion

    #region 多节
    [Fact(DisplayName = "多节—读取 2 节文档（段内嵌 sectPr 切分）")]
    public void Section_ReadTwo()
    {
        var body = "<w:p><w:r><w:t>第一节内容</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr></w:pPr>"
            + "<w:r><w:t>第一节末</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>第二节内容</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/></w:sectPr></w:pPr>"
            + "<w:r><w:t>第二节末</w:t></w:r></w:p>";
        var doc = Read(body);

        Assert.Equal(2, doc.Sections.Count);
        Assert.Equal(2, doc.Sections[0].Elements.Count);
        Assert.Equal(2, doc.Sections[1].Elements.Count);
        Assert.False(doc.Sections[0].PageSettings.Landscape);
        Assert.True(doc.Sections[1].PageSettings.Landscape);
        Assert.NotNull(doc.Sections[0].SectPrXml);
        Assert.NotNull(doc.Sections[1].SectPrXml);
        // 全文档平铺视图保持完整
        Assert.Equal(4, doc.Elements.Count);
    }

    [Fact(DisplayName = "多节—程序化 2 节写入（首节纵向/次节横向）")]
    public void Section_WriteTwo()
    {
        var doc = new Document { DocumentXml = null };
        var sec1 = new Section();
        sec1.Elements.Add(new Element
        {
            Type = ElementType.Paragraph,
            Paragraph = new Paragraph { Runs = { new Run { Text = "第一节" } } },
        });
        sec1.PageSettings.Landscape = false;
        var sec2 = new Section();
        sec2.Elements.Add(new Element
        {
            Type = ElementType.Paragraph,
            Paragraph = new Paragraph { Runs = { new Run { Text = "第二节" } } },
        });
        sec2.PageSettings.Landscape = true;
        doc.Sections.Add(sec1);
        doc.Sections.Add(sec2);

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter()) w.Save(path, doc);

            using var reader = new WordReader(path);
            var read = reader.ReadDocument();
            Assert.Equal(2, read.Sections.Count);
            Assert.False(read.Sections[0].PageSettings.Landscape);
            Assert.True(read.Sections[1].PageSettings.Landscape);
            Assert.Contains("第一节", GetSectionText(read.Sections[0]));
            Assert.Contains("第二节", GetSectionText(read.Sections[1]));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    private static String GetSectionText(Section sec)
    {
        var sb = new StringBuilder();
        foreach (var el in sec.Elements)
        {
            if (el.Paragraph != null)
                foreach (var r in el.Paragraph.Runs) sb.Append(r.Text);
        }
        return sb.ToString();
    }
    #endregion
}
