using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>OOXML 子元素顺序测试：rPr / pPr / tcPr / sectPr 必须按 schema 顺序输出</summary>
/// <remarks>
/// 顺序错误不会让 Word 必然打不开，但 OpenXmlValidator 判为 "unexpected child element"，
/// 严格校验器/第三方工具会报文件损坏。这里取「单个属性的首段」逐对断言先后关系。
/// </remarks>
public class SchemaOrderTests
{
    private static Byte[] Save(Document doc)
    {
        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
            writer.Save(ms, doc);
        return ms.ToArray();
    }

    private static String ReadDocumentXml(Byte[] bytes)
    {
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var sr = new StreamReader(za.GetEntry("word/document.xml")!.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    /// <summary>取首个 startTag..endTag 片段（含首尾）</summary>
    private static String Slice(String xml, String startTag, String endTag)
    {
        var from = xml.IndexOf(startTag, StringComparison.Ordinal);
        Assert.True(from >= 0, $"未找到 {startTag}");
        var to = xml.IndexOf(endTag, from, StringComparison.Ordinal);
        Assert.True(to >= 0, $"未找到 {endTag}");
        return xml[from..(to + endTag.Length)];
    }

    /// <summary>断言片段中各标记依次出现（不存在的标记跳过）</summary>
    private static void AssertOrder(String xml, params String[] order)
    {
        var last = -1;
        var lastName = "";
        foreach (var tag in order)
        {
            var idx = xml.IndexOf(tag, StringComparison.Ordinal);
            if (idx < 0) continue;
            Assert.True(idx > last, $"子元素顺序错误：{tag} 出现在 {lastName} 之前（{xml}）");
            last = idx;
            lastName = tag;
        }
    }

    [Fact(DisplayName = "顺序-rPr 中 rFonts 在最前、color/sz/u 依次靠后")]
    public void RunProperties_SchemaOrder()
    {
        var doc = new Document();
        var el = doc.AddParagraph("带格式文本");
        el.Paragraph!.Runs[0].Properties = new RunProperties
        {
            FontName = "宋体",
            EastAsiaFontName = "宋体",
            Bold = true,
            ForeColor = "FF0000",
            FontSize = 12f,
            UnderlineStyle = "single",
            CharacterSpacing = 20,
        };

        var rPr = Slice(ReadDocumentXml(Save(doc)), "<w:rPr>", "</w:rPr>");

        AssertOrder(rPr, "<w:rPr>", "<w:rFonts", "<w:b", "<w:color", "<w:spacing", "<w:sz", "<w:u ");
    }

    [Fact(DisplayName = "顺序-pPr 中 pStyle/keepNext 在前，spacing/ind 在 shd 之后，jc 最后")]
    public void ParagraphProperties_SchemaOrder()
    {
        var doc = new Document();
        var el = doc.AddHeading("标题", 1);
        el.Paragraph!.KeepNext = true;
        el.Paragraph!.BackgroundColor = "FFFF00";
        el.Paragraph!.SpaceAfter = 120;
        el.Paragraph!.IndentLeft = 360;
        el.Paragraph!.Alignment = "center";

        var pPr = Slice(ReadDocumentXml(Save(doc)), "<w:pPr>", "</w:pPr>");

        AssertOrder(pPr, "<w:pPr>", "<w:pStyle", "<w:keepNext", "<w:shd", "<w:spacing", "<w:ind", "<w:jc");
    }

    [Fact(DisplayName = "顺序-tcPr 中 tcW/gridSpan 在 tcBorders、shd、vAlign 之前")]
    public void CellProperties_SchemaOrder()
    {
        var doc = new Document();
        doc.AddTable([["表头", "值"], ["合并", "单元"]]);
        var table = doc.Elements[0].Table!;
        table.Rows[1].Cells[0].ColSpan = 2;
        table.Rows[1].Cells[0].BackgroundColor = "EEEEEE";
        table.Rows[1].Cells[0].VerticalAlignment = "center";

        var tcPr = Slice(ReadDocumentXml(Save(doc)), "<w:tcPr>", "</w:tcPr>");

        AssertOrder(tcPr, "<w:tcPr>", "<w:tcW", "<w:gridSpan", "<w:tcBorders", "<w:shd", "<w:vAlign");
    }

    [Fact(DisplayName = "顺序-sectPr 中 pgSz/pgMar 在 pgBorders 之前，titlePg 在 cols 之后")]
    public void SectionProperties_SchemaOrder()
    {
        var doc = new Document();
        doc.AddParagraph("正文");
        doc.PageSettings.PageBorder = new PageBorder { Top = "single" };
        doc.PageSettings.ColumnCount = 2;
        doc.PageSettings.LineNumber = new LineNumberSettings();
        doc.PageSettings.TitlePage = true;

        var sectPr = Slice(ReadDocumentXml(Save(doc)), "<w:sectPr", "</w:sectPr>");

        AssertOrder(sectPr, "<w:sectPr", "<w:pgSz", "<w:pgMar", "<w:pgBorders", "<w:lnNumType", "<w:cols", "<w:titlePg");
    }

    [Fact(DisplayName = "顺序-奇偶页不同写入 settings.xml 而非 sectPr")]
    public void EvenAndOddHeaders_GoesToSettings()
    {
        var doc = new Document();
        doc.AddParagraph("正文");
        doc.PageSettings.EvenAndOddHeaders = true;

        var bytes = Save(doc);
        var documentXml = ReadDocumentXml(bytes);
        Assert.DoesNotContain("evenAndOddHeaders", documentXml);

        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var sr = new StreamReader(za.GetEntry("word/settings.xml")!.Open(), Encoding.UTF8);
        Assert.Contains("evenAndOddHeaders", sr.ReadToEnd());
    }

    [Fact(DisplayName = "顺序-超链接 run 的 color/u 只出现一次且位置合法")]
    public void Hyperlink_RunProperties_NoDuplicate()
    {
        var doc = new Document();
        var p = doc.AddParagraph("");
        p.Paragraph!.Runs.Add(new Run { Text = "链接文字", HyperlinkRelId = "rLink1", Properties = new RunProperties() });

        var xml = ReadDocumentXml(Save(doc));

        // 取出带 Hyperlink 样式的那段 rPr
        var idx = xml.IndexOf("w:rStyle w:val=\"Hyperlink\"", StringComparison.Ordinal);
        Assert.True(idx > 0);
        var from = xml.LastIndexOf("<w:rPr>", idx, StringComparison.Ordinal);
        var to = xml.IndexOf("</w:rPr>", idx, StringComparison.Ordinal);
        var rPr = xml[from..(to + "</w:rPr>".Length)];

        Assert.Single(System.Text.RegularExpressions.Regex.Matches(rPr, "<w:color "));
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(rPr, "<w:u "));
        AssertOrder(rPr, "<w:rStyle", "<w:color", "<w:u");
    }

    [Fact(DisplayName = "顺序-单元格 tcBorders 不得重复输出")]
    public void CellProperties_SingleBordersElement()
    {
        var doc = new Document();
        doc.AddTable([["A", "B"]]);
        // 同时给表格样式与单元格自定义边框：旧实现会输出两个 w:tcBorders（元素重复非法）
        doc.Elements[0].TableStyle = new TableStyle { BorderColor = "000000" };
        doc.Elements[0].Table!.Rows[0].Cells[0].Borders = new TableBorders { Top = new Border { Style = BorderStyle.Single, Width = 4 } };

        var tcPr = Slice(ReadDocumentXml(Save(doc)), "<w:tcPr>", "</w:tcPr>");

        Assert.Contains("<w:tcBorders>", tcPr);
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(tcPr, "<w:tcBorders>"));
    }
}
