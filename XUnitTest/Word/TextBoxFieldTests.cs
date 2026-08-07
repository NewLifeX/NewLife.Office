using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>
/// 文本框与域代码读取正确性测试（W22）
/// </summary>
/// <remarks>
/// 覆盖：
/// <list type="bullet">
/// <item>ReadParagraphs 文本框文本去重（此前重复出现）</item>
/// <item>文本框模型（Paragraph.TextBoxes）+ FindText/ReplaceText 覆盖</item>
/// <item>文本框替换持久化（同步外层段落 RawXml，保存后生效且形状保留）</item>
/// <item>域代码读取（fldChar 结果可读、指令文本不泄漏）</item>
/// </list>
/// </remarks>
public class TextBoxFieldTests
{
    #region 辅助

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
        using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
        return (bytes, sr.ReadToEnd());
    }

    /// <summary>含文本框的文档</summary>
    private static (Byte[] Docx, String DocumentXml) BuildTextboxDocx(String textboxText = "文本框内文本")
    {
        return BuildDocx(w =>
        {
            w.AppendParagraph("正文段落");
            var shape = WordShape.Rect(5, 2, "336699");
            shape.Text = textboxText;
            w.AppendShape(shape);
        });
    }

    #endregion

    #region 读取去重与模型

    [Fact, DisplayName("W22_文本框_ReadParagraphs文本去重")]
    public void TextBox_ReadParagraphs_Dedup()
    {
        var (bytes, _) = BuildTextboxDocx();
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var paras = reader.ReadParagraphs().ToList();

        // 文本框文本只出现一次（此前外层段落 .//w:t 与内层段落 //w:p 双重匹配导致重复）
        Assert.Equal(2, paras.Count);
        Assert.Contains("正文段落", paras);
        Assert.Equal(1, paras.Count(p => p == "文本框内文本"));
    }

    [Fact, DisplayName("W22_文本框_模型TextBoxes解析")]
    public void TextBox_Model_TextBoxesParsed()
    {
        var (bytes, _) = BuildTextboxDocx();
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();

        var shapePara = doc.Elements.Select(e => e.Paragraph)
            .FirstOrDefault(p => p != null && p.TextBoxes.Count > 0);
        Assert.NotNull(shapePara);
        Assert.Single(shapePara!.TextBoxes);
        Assert.Equal("文本框内文本", String.Concat(shapePara.TextBoxes[0].Runs.Select(r => r.Text)));
        Assert.NotNull(shapePara.TextBoxes[0].RawXml);
    }

    [Fact, DisplayName("W22_文本框_FindText命中文本框")]
    public void TextBox_FindText_Hits()
    {
        var (bytes, _) = BuildTextboxDocx();
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();

        var matches = doc.FindText("文本框内文本");
        Assert.Single(matches);
    }

    [Fact, DisplayName("W22_文本框_ReplaceText命中并更新模型")]
    public void TextBox_ReplaceText_UpdatesModel()
    {
        var (bytes, _) = BuildTextboxDocx("旧文本内容");
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();

        var count = doc.ReplaceText("旧文本内容", "新文本内容");
        Assert.Equal(1, count);

        var shapePara = doc.Elements.Select(e => e.Paragraph)
            .First(p => p!.TextBoxes.Count > 0)!;
        Assert.Equal("新文本内容", String.Concat(shapePara.TextBoxes[0].Runs.Select(r => r.Text)));
    }

    [Fact, DisplayName("W22_文本框_替换后保存持久化且形状保留")]
    public void TextBox_ReplaceText_PersistsAfterSave()
    {
        var (bytes, _) = BuildTextboxDocx("旧文本内容");
        var outPath = Path.Combine(Path.GetTempPath(), $"tb_{Guid.NewGuid():N}.docx");
        try
        {
            Document doc;
            using (var ms = new MemoryStream(bytes))
            using (var reader = new WordReader(ms))
                doc = reader.ReadDocument();

            var count = doc.ReplaceText("旧文本内容", "新文本内容");
            Assert.Equal(1, count);

            // 持久化工作流：关闭整文透传，保留元素级 RawXml（形状绘制靠它保留）
            doc.DocumentXml = null;
            using (var writer = new WordWriter())
                writer.Save(outPath, doc);

            // 回读：文本框文本已更新
            using (var reader = new WordReader(outPath))
            {
                var re = reader.ReadDocument();
                var shapePara = re.Elements.Select(e => e.Paragraph)
                    .First(p => p!.TextBoxes.Count > 0)!;
                Assert.Equal("新文本内容", String.Concat(shapePara.TextBoxes[0].Runs.Select(r => r.Text)));

                // 形状绘制保留（txbxContent + 形状 XML 仍存在）
                Assert.Contains("<wps:txbx>", re.DocumentXml!);
                Assert.Contains("wordprocessingShape", re.DocumentXml!);
            }

            // 全文档文本提取一致（去重且为新文本）
            using (var reader = new WordReader(outPath))
            {
                var paras = reader.ReadParagraphs().ToList();
                Assert.Contains("新文本内容", paras);
                Assert.Equal(1, paras.Count(p => p.Contains("文本内容")));
            }
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    #endregion

    #region 域代码读取

    [Fact, DisplayName("W22_域代码_MERGEFIELD结果可读且指令不泄漏")]
    public void Field_MergeField_ResultReadable()
    {
        var (bytes, _) = BuildDocx(w =>
        {
            w.AppendParagraph("正文开始");
            w.AppendMergeField("CustomerName");
            w.AppendParagraph("正文结束");
        });
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);

        var paras = reader.ReadParagraphs().ToList();
        Assert.Contains("正文开始", paras);
        Assert.Contains("正文结束", paras);
        // 域缓存结果可读
        Assert.Contains(paras, p => p.Contains("CustomerName"));
        // 指令文本（MERGEFIELD 指令）不泄漏到纯文本
        Assert.DoesNotContain(paras, p => p.Contains("MERGEFIELD"));

        var doc = reader.ReadDocument();
        var fieldPara = doc.Elements.Select(e => e.Paragraph)
            .First(p => p != null && p.Runs.Any(r => r.Text.Contains("CustomerName")));
        Assert.NotNull(fieldPara);
    }

    [Fact, DisplayName("W22_域代码_无缓存结果的空域不产生空段落文本")]
    public void Field_EmptyResult_NoLeak()
    {
        // 构造完整 docx：PAGE 域（begin + 指令 + end，无 separate 结果）
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var bodyInner = "<w:p><w:r><w:t>页脚上方</w:t></w:r></w:p>"
            + "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + "<w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>"
            + "<w:p><w:r><w:t>正文</w:t></w:r></w:p>";
        var bytes = BuildMinimalDocx(bodyInner);

        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var paras = reader.ReadParagraphs().ToList();
        Assert.Contains("页脚上方", paras);
        Assert.Contains("正文", paras);
        // PAGE 指令不泄漏；无结果时不产生空段落文本
        Assert.DoesNotContain(paras, p => p.Contains("PAGE"));
    }

    [Fact, DisplayName("W22_ReadParagraphs_tab与换行语义")]
    public void ReadParagraphs_TabAndBreak()
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var bodyInner = "<w:p><w:r><w:t>姓名</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>部门</w:t></w:r>"
            + "<w:r><w:br/></w:r><w:r><w:t>第二行</w:t></w:r></w:p>";
        var bytes = BuildMinimalDocx(bodyInner);

        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var para = Assert.Single(reader.ReadParagraphs());
        Assert.Equal("姓名\t部门\n第二行", para);
    }

    [Fact, DisplayName("W22_内联SDT_内容并入段落Runs且Find可达")]
    public void InlineSdt_RunsParsed()
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var bodyInner = "<w:p><w:r><w:t>开始</w:t></w:r>"
            + "<w:sdt><w:sdtPr><w:alias w:val=\"字段\"/><w:tag w:val=\"FIELD\"/><w:id w:val=\"123\"/><w:text/></w:sdtPr>"
            + "<w:sdtContent><w:r><w:t>控件内容</w:t></w:r></w:sdtContent></w:sdt>"
            + "<w:r><w:t>结束</w:t></w:r></w:p>";
        var bytes = BuildMinimalDocx(bodyInner);

        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);

        // 纯文本提取包含控件内容
        var para = Assert.Single(reader.ReadParagraphs());
        Assert.Equal("开始控件内容结束", para);

        // 模型：SDT 内容并入 Runs，FindText 可达
        var doc = reader.ReadDocument();
        var p = Assert.Single(doc.Elements).Paragraph!;
        Assert.Contains(p.Runs, r => r.Text == "控件内容");
        Assert.Single(doc.FindText("控件内容"));
    }

    [Fact, DisplayName("W22_文本框_GetText与Markdown转换包含文本框")]
    public void TextBox_GetTextAndMarkdown()
    {
        var (bytes, _) = BuildTextboxDocx("框内说明文字");
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);

        // 模型 GetText 包含文本框
        var doc = reader.ReadDocument();
        Assert.Contains("框内说明文字", doc.GetText());

        // Markdown 转换包含文本框（Word 原生导出行为）
        var md = reader.ExtractMarkdown();
        Assert.NotNull(md);
        Assert.Contains("框内说明文字", md);
    }

    [Fact, DisplayName("W22_文本框_表格单元格内替换持久化")]
    public void TextBox_InTableCell_ReplacePersists()
    {
        // 表格单元格内嵌含文本框的形状段落
        var shapePara = "<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">"
            + "<wp:extent cx=\"1800000\" cy=\"720000\"/><wp:docPr id=\"1\" name=\"Shape1\"/>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
            + "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1800000\" cy=\"720000\"/></a:xfrm>"
            + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=\"336699\"/></a:solidFill></wps:spPr>"
            + "<wps:txbx><w:txbxContent><w:p><w:r><w:t>CELLTB</w:t></w:r></w:p></w:txbxContent></wps:txbx>"
            + "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        var bodyInner = "<w:p><w:r><w:t>表格前</w:t></w:r></w:p>"
            + "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>"
            + shapePara + "</w:tc></w:tr></w:tbl>"
            + "<w:p><w:r><w:t>表格后</w:t></w:r></w:p>";
        var bytes = BuildMinimalDocx(bodyInner, fullNs: true);
        var outPath = Path.Combine(Path.GetTempPath(), $"tbcell_{Guid.NewGuid():N}.docx");
        try
        {
            Document doc;
            using (var ms = new MemoryStream(bytes))
            using (var reader = new WordReader(ms))
                doc = reader.ReadDocument();

            // FindText 可达单元格内文本框
            Assert.Single(doc.FindText("CELLTB"));

            var count = doc.ReplaceText("CELLTB", "NEWCELL");
            Assert.Equal(1, count);

            // 持久化：关闭整文透传，表格元素 RawXml 的 txbxContent 已同步
            doc.DocumentXml = null;
            using (var writer = new WordWriter())
                writer.Save(outPath, doc);

            using (var reader = new WordReader(outPath))
            {
                var re = reader.ReadDocument();
                Assert.Single(re.FindText("NEWCELL"));
                Assert.Empty(re.FindText("CELLTB"));
            }
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    #endregion

    #region 辅助

    /// <summary>构建最小 docx（自定义 body 内容，可选完整 DrawingML 命名空间）</summary>
    private static Byte[] BuildMinimalDocx(String bodyInnerXml, Boolean fullNs = false)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var nsDecl = $"<w:document xmlns:w=\"{W}\"";
        if (fullNs)
        {
            nsDecl += " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";
        }
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + nsDecl + $"><w:body>{bodyInnerXml}</w:body></w:document>";
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void WriteEntry(String path, String content)
            {
                using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
                sw.Write(content);
            }

            WriteEntry("[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>");
            WriteEntry("_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry("word/document.xml", documentXml);
            WriteEntry("word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        }
        return ms.ToArray();
    }

    #endregion
}
