using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>Word 写出器样式测试。覆盖默认样式美化、StylesXml 注入与表头字体颜色</summary>
public class WriterStyleTests
{
    private static String ReadEntry(String path, String entryName)
    {
        using var za = new ZipArchive(File.OpenRead(path), ZipArchiveMode.Read);
        var entry = za.GetEntry(entryName);
        using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    [Fact(DisplayName = "WordWriter—默认样式为现代排版（字体/行距/标题配色/表格边框）")]
    public void DefaultStyles_AreModernized()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using var writer = new WordWriter();
            writer.AppendHeading("章节标题", 1);
            writer.AppendParagraph("正文内容");
            writer.Save(tempFile);

            var xml = ReadEntry(tempFile, "word/styles.xml");
            // 默认字体：西文 Segoe UI + 中文微软雅黑，五号字（10.5pt）
            Assert.Contains("Segoe UI", xml);
            Assert.Contains("微软雅黑", xml);
            Assert.Contains("<w:sz w:val=\"21\"/>", xml);
            // 默认段落：1.3 倍行距 + 段后 6pt
            Assert.Contains("<w:spacing w:line=\"312\" w:lineRule=\"auto\" w:after=\"120\"/>", xml);
            // 标题：品牌深蓝配色 + 段前后间距 + keepNext
            Assert.Contains("<w:color w:val=\"1F4E79\"/>", xml);
            Assert.Contains("<w:keepNext/>", xml);
            // 表格边框：浅灰替代纯黑
            Assert.Contains("w:color=\"BFBFBF\"", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "WordWriter—StylesXml注入后原样写出（替换默认样式表）")]
    public void StylesXml_Injection_OverridesDefaultStyles()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            const String custom = "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia=\"仿宋\"/><w:sz w:val=\"28\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
                "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
                "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/></w:style>" +
                "</w:styles>";

            using var writer = new WordWriter { StylesXml = custom };
            writer.AppendHeading("章节标题", 1);
            writer.Save(tempFile);

            var xml = ReadEntry(tempFile, "word/styles.xml");
            Assert.Equal(custom, xml);
            Assert.DoesNotContain("Segoe UI", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "WordWriter—表格表头字色与背景色输出")]
    public void Table_HeaderStyling_Applied()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using var writer = new WordWriter();
            var rows = new List<IEnumerable<String>>
            {
                new String[] { "姓名", "分数" },
                new String[] { "张三", "85" },
            };
            writer.AppendTable(rows, firstRowHeader: true, new TableStyle
            {
                HeaderBgColor = "1F4E79",
                HeaderFontColor = "FFFFFF",
                StripeColor = "F1F5F9",
            });
            writer.Save(tempFile);

            var xml = ReadEntry(tempFile, "word/document.xml");
            Assert.Contains("w:fill=\"1F4E79\"", xml);            // 表头底色
            Assert.Contains("<w:color w:val=\"FFFFFF\"/>", xml);   // 表头白字
            Assert.Contains("w:fill=\"F1F5F9\"", xml);            // 斑马纹
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    [Fact(DisplayName = "WordWriter—表格默认内边距（避免文字贴边框）")]
    public void Table_DefaultCellMargins_Applied()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using var writer = new WordWriter();
            var rows = new List<IEnumerable<String>>
            {
                new String[] { "姓名", "分数" },
                new String[] { "张三", "85" },
            };
            // 自定义样式与内置 TableGrid 两种路径都应有内边距
            writer.AppendTable(rows, firstRowHeader: true, new TableStyle { HeaderBgColor = "2563EB" });
            writer.AppendTable(rows, firstRowHeader: true);
            writer.Save(tempFile);

            var xml = ReadEntry(tempFile, "word/document.xml");
            Assert.Contains("<w:tblCellMar>", xml);
            Assert.Contains("<w:left w:w=\"108\" w:type=\"dxa\"/>", xml);
            Assert.Contains("<w:right w:w=\"108\" w:type=\"dxa\"/>", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }}
