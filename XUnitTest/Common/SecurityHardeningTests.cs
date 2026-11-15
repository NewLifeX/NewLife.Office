using System.ComponentModel;
using System.IO.Compression;
using System.Text;
using NewLife.Office;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using Xunit;

namespace XUnitTest.Common;

/// <summary>不可信输入加固与输出兼容回归</summary>
public class SecurityHardeningTests
{
    [Fact, DisplayName("超宽单元格引用被拒绝，不会展开成千万级列")]
    public void Excel_OversizedColumnReference_Throws()
    {
        using var ms = BuildMinimalXlsx("""
            <?xml version="1.0" encoding="utf-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData><row r="1"><c r="ZZZZZ1" t="inlineStr"><is><t>x</t></is></c></row></sheetData>
            </worksheet>
            """);

        using var reader = new ExcelReader(ms, Encoding.UTF8);
        Assert.Throws<InvalidDataException>(() => reader.ReadRows("S1").ToList());
    }

    [Fact, DisplayName("XML 内部实体不会展开（防亿枚笑弹）")]
    public void Xml_InternalEntityExpansion_IsBlocked()
    {
        var payload = new StringBuilder();
        payload.Append("<?xml version=\"1.0\"?><!DOCTYPE lolz [<!ENTITY lol \"lol\">");
        for (var i = 1; i <= 8; i++)
        {
            payload.Append($"<!ENTITY lol{i} \"");
            for (var j = 0; j < 10; j++) payload.Append(i == 1 ? "&lol;" : $"&lol{i - 1};");
            payload.Append("\">");
        }
        payload.Append("]><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        payload.Append("<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&lol8;</t></is></c></row></sheetData></worksheet>");

        using var ms = BuildMinimalXlsx(payload.ToString());
        using var reader = new ExcelReader(ms, Encoding.UTF8);

        String? value = null;
        try
        {
            value = reader.ReadRows("S1").FirstOrDefault()?[0] as String;
        }
        catch (Exception ex) when (ex is InvalidDataException or System.Xml.XmlException)
        {
            // 实体未声明或超过文档字符上限都会直接拒绝，同样说明没有按实体展开
        }

        Assert.True(value == null || value.Length < 1024, $"内部实体被展开了，长度 {value?.Length}");
    }

    [Fact, DisplayName("ODS 文本 00123/true 读回仍是字符串")]
    public void Ods_LeadingZeroText_StaysString()
    {
        var writer = new OdsWriter();
        writer.AddSheet("S", [["00123", "true", "12.5"]]);
        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        var sheets = OdsReader.Read(ms);
        var rows = sheets.First(s => s.Name == "S").Rows;
        Assert.Equal("00123", rows[0][0]);
        Assert.Equal("true", rows[0][1]);
        Assert.Equal("12.5", rows[0][2]);
    }

    [Fact, DisplayName("EPUB 章节名无法跳出 OEBPS 目录")]
    public void Epub_ChapterName_CannotEscape()
    {
        var doc = new EpubDocument { Title = "T", Author = "A" };
        doc.Chapters.Add(new EpubChapter { Title = "一", Content = "<p>1</p>", FileName = "../evil.xhtml" });

        using var ms = new MemoryStream();
        new EpubWriter().Write(doc, ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, true);
        Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains("..") || e.FullName.StartsWith("/"));
        Assert.Contains(zip.Entries, e => e.FullName == "OEBPS/evil.xhtml");
    }

    [Fact, DisplayName("iCal 带 TZID 的时间换算为 UTC")]
    public void ICal_Tzid_ConvertsToUtc()
    {
        var ics = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:1\r\nSUMMARY:会议\r\n"
                + "DTSTART;TZID=America/New_York:20240102T120000\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n";

        var doc = new ICalReader().Parse(ics);
        var evt = Assert.Single(doc.Events);
        Assert.Equal(new DateTime(2024, 1, 2, 17, 0, 0, DateTimeKind.Utc), evt.Start!.Value.UtcDateTime);
    }

    [Fact, DisplayName("Markdown 危险链接被识别，控制字符不能绕过")]
    public void Markdown_DangerousLinks_AreDetected()
    {
        var html = MarkdownDocument.Parse("[a](java\tscript:alert(1)) [b](javascript:alert(2))").ToHtml();
        Assert.DoesNotContain("href=\"java", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=\"javascript", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, DisplayName("EML 无 BOM 且 QP 换行不重复")]
    public void Eml_NoBom_AndNoDoubleCr()
    {
        var msg = new Message
        {
            From = "a@example.com",
            Subject = "hi",
            TextBody = "第一行\r\n第二行",
        };

        using var ms = new MemoryStream();
        new EmlWriter().Write(msg, ms);
        var bytes = ms.ToArray();

        Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "EML 不应带 UTF-8 BOM");

        var text = Encoding.UTF8.GetString(bytes);
        Assert.DoesNotContain("\r\r\n", text);

        var parsed = new EmlReader().Parse(bytes);
        Assert.Equal("hi", parsed.Subject);
        Assert.Contains("第二行", parsed.TextBody);
    }

    [Fact, DisplayName("RTF 高位字符使用有符号 uN 输出")]
    public void Rtf_HighUnicode_UsesSignedValue()
    {
        var writer = new RtfWriter();
        writer.AddParagraph("😀 汉");

        var rtf = writer.ToString();

        // 代理对必须拆成两个有符号 \uN：0xD83D-0x10000=-10179，0xDE00-0x10000=-8704
        Assert.Contains("\\u-10179?", rtf);
        Assert.Contains("\\u-8704?", rtf);
        Assert.Contains("\\u27721?", rtf); // 汉 = 0x6C49 = 27721
    }

    [Fact, DisplayName("vCard/iCal 属性值中的 CRLF 被剔除")]
    public void TextFormats_StripCrLfInjection()
    {
        var contact = new Contact
        {
            FullName = "正常",
            Uid = "id\r\nX-EVIL:1",
            Note = "备注",
        };

        var card = new VCardWriter().Build(contact);
        Assert.DoesNotContain("\r\nX-EVIL", card);

        var evt = new ICalEvent { Uid = "u\r\nX-EVIL:1", Summary = "s" };
        var cal = new ICalDocument();
        cal.Events.Add(evt);
        var ics = new ICalWriter().Build(cal);
        Assert.DoesNotContain("\r\nX-EVIL", ics);
    }

    /// <summary>构造一个只含单个工作表的 xlsx，工作表 XML 由调用方提供</summary>
    private static MemoryStream BuildMinimalXlsx(String sheetXml)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            WriteEntry(zip, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "</Types>");
            WriteEntry(zip, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                + "</Relationships>");
            WriteEntry(zip, "xl/workbook.xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            WriteEntry(zip, "xl/_rels/workbook.xml.rels",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "</Relationships>");
            WriteEntry(zip, "xl/worksheets/sheet1.xml", sheetXml);
        }

        ms.Position = 0;
        return ms;
    }

    private static void WriteEntry(ZipArchive zip, String name, String content)
    {
        var entry = zip.CreateEntry(name);
        using var s = entry.Open();
        var bytes = new UTF8Encoding(false).GetBytes(content);
        s.Write(bytes, 0, bytes.Length);
    }
}
