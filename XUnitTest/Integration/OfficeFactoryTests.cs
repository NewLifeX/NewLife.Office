using System.ComponentModel;
using System.Text;
using NewLife.Office;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>OfficeFactory 工厂类自身功能测试</summary>
public class OfficeFactoryTests : IntegrationTestBase
{
    [Fact, DisplayName("IsSupported_支持的后缀返回true")]
    public void IsSupported_Returns_True_For_All_Supported()
    {
        var extensions = new[] { ".xlsx", ".xls", ".docx", ".doc", ".pptx", ".ppt", ".pdf", ".rtf", ".ods", ".epub", ".vcf", ".eml", ".ics", ".md", ".xps" };
        foreach (var ext in extensions)
        {
            Assert.True(OfficeFactory.IsSupported(ext), $"应支持 {ext}");
        }
    }

    [Fact, DisplayName("IsSupported_不带点号也返回true")]
    public void IsSupported_WithoutDot_Returns_True()
    {
        Assert.True(OfficeFactory.IsSupported("xlsx"));
        Assert.True(OfficeFactory.IsSupported("pdf"));
        Assert.True(OfficeFactory.IsSupported("md"));
    }

    [Fact, DisplayName("IsSupported_不支持的后缀返回false")]
    public void IsSupported_Returns_False_For_Unsupported()
    {
        Assert.False(OfficeFactory.IsSupported(".txt"));
        Assert.False(OfficeFactory.IsSupported(".csv"));
        Assert.False(OfficeFactory.IsSupported(".zip"));
        Assert.False(OfficeFactory.IsSupported(""));
        Assert.False(OfficeFactory.IsSupported(null!));
    }

    [Fact, DisplayName("IsSupported_大小写不敏感")]
    public void IsSupported_CaseInsensitive()
    {
        Assert.True(OfficeFactory.IsSupported(".XLSX"));
        Assert.True(OfficeFactory.IsSupported(".Pdf"));
        Assert.True(OfficeFactory.IsSupported("DOCX"));
    }

    [Fact, DisplayName("CreateReader_文件不存在抛FileNotFoundException")]
    public void CreateReader_FileNotFound_Throws()
    {
        Assert.Throws<FileNotFoundException>(() => OfficeFactory.CreateReader("not_exist.xlsx"));
    }

    [Fact, DisplayName("CreateReader_空路径抛ArgumentNullException")]
    public void CreateReader_NullPath_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => OfficeFactory.CreateReader(null!));
        Assert.Throws<ArgumentNullException>(() => OfficeFactory.CreateReader(""));
        Assert.Throws<ArgumentNullException>(() => OfficeFactory.CreateReader("   "));
    }

    [Fact, DisplayName("CreateReader_不支持格式抛NotSupportedException")]
    public void CreateReader_UnsupportedFormat_Throws()
    {
        var path = Path.Combine(OutputDir, "test.txt");
        File.WriteAllText(path, "hello");
        Assert.Null(OfficeFactory.CreateReader(path));
    }

    [Fact, DisplayName("SupportedExtensions_包含15种格式")]
    public void SupportedExtensions_Has15()
    {
        Assert.Equal(15, OfficeFactory.SupportedExtensions.Count);
    }

    #region ReadText

    [Fact, DisplayName("ReadText_null路径返回null")]
    public void ReadText_NullPath_Returns_Null()
    {
        Assert.Null(OfficeFactory.ReadText(null!));
        Assert.Null(OfficeFactory.ReadText(""));
        Assert.Null(OfficeFactory.ReadText("   "));
    }

    [Fact, DisplayName("ReadText_从xlsx文件提取CSV格式文本")]
    public void ReadText_Xlsx_Returns_CsvText()
    {
        var path = Path.Combine(OutputDir, "factory_readtext.xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteHeader("Sheet1", new[] { "姓名", "年龄", "薪资" });
            w.WriteRow("Sheet1", new Object?[] { "张三", 28, 8500m });
            w.WriteRow("Sheet1", new Object?[] { "李四", 35, 12000m });
            w.Save();
        }

        var text = OfficeFactory.ReadText(path);

        Assert.NotNull(text);
        Assert.Contains("姓名", text);
        Assert.Contains("张三", text);
        Assert.Contains("李四", text);
    }

    [Fact, DisplayName("ReadText_从docx文件提取纯文本")]
    public void ReadText_Docx_Returns_PlainText()
    {
        var path = Path.Combine(OutputDir, "factory_readtext.docx");
        using (var w = new WordWriter())
        {
            w.AppendHeading("测试标题", 1);
            w.AppendParagraph("这是一段正文内容。");
            w.AppendParagraph("第二段文字。");
            w.Save(path);
        }

        var text = OfficeFactory.ReadText(path);

        Assert.NotNull(text);
        Assert.Contains("测试标题", text);
        Assert.Contains("这是一段正文内容", text);
    }

    [Fact, DisplayName("ReadText_从md文件提取纯文本")]
    public void ReadText_Md_Returns_PlainText()
    {
        var path = Path.Combine(OutputDir, "factory_readtext.md");
        File.WriteAllText(path, "# 标题\n\n这是正文段落。\n\n## 二级标题\n\n另一段文字。\n", new UTF8Encoding(false));

        var text = OfficeFactory.ReadText(path);

        Assert.NotNull(text);
        Assert.Contains("标题", text);
        Assert.Contains("正文段落", text);
    }

    [Fact, DisplayName("ReadText_从流提取xlsx文本")]
    public void ReadText_Stream_Xlsx_Returns_Text()
    {
        using var ms = new MemoryStream();
        using (var w = new ExcelWriter(ms))
        {
            w.WriteHeader("Data", new[] { "品名", "数量" });
            w.WriteRow("Data", new Object?[] { "苹果", 100 });
            w.Save();
        }

        ms.Position = 0;
        var text = OfficeFactory.ReadText(ms, ".xlsx");

        Assert.NotNull(text);
        Assert.Contains("品名", text);
        Assert.Contains("苹果", text);
    }

    [Fact, DisplayName("ReadText_流为null返回null")]
    public void ReadText_NullStream_Returns_Null()
    {
        Assert.Null(OfficeFactory.ReadText(null!, ".xlsx"));
    }

    [Fact, DisplayName("ReadText_不支持格式的流返回null")]
    public void ReadText_UnsupportedStream_Returns_Null()
    {
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
        Assert.Null(OfficeFactory.ReadText(ms, ".unknown"));
    }

    #endregion

    #region ReadMarkdown

    [Fact, DisplayName("ReadMarkdown_null路径返回null")]
    public void ReadMarkdown_NullPath_Returns_Null()
    {
        Assert.Null(OfficeFactory.ReadMarkdown(null!));
        Assert.Null(OfficeFactory.ReadMarkdown(""));
        Assert.Null(OfficeFactory.ReadMarkdown("   "));
    }

    [Fact, DisplayName("ReadMarkdown_从xlsx文件提取Markdown表格")]
    public void ReadMarkdown_Xlsx_Returns_MarkdownTable()
    {
        var path = Path.Combine(OutputDir, "factory_readmd.xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteHeader("Sheet1", new[] { "城市", "人口", "面积" });
            w.WriteRow("Sheet1", new Object?[] { "北京", 2154, 16410 });
            w.WriteRow("Sheet1", new Object?[] { "上海", 2428, 6340 });
            w.Save();
        }

        var md = OfficeFactory.ReadMarkdown(path);

        Assert.NotNull(md);
        Assert.Contains("城市", md);
        Assert.Contains("北京", md);
        Assert.Contains("|", md);
    }

    [Fact, DisplayName("ReadMarkdown_从docx文件提取Markdown文本")]
    public void ReadMarkdown_Docx_Returns_Markdown()
    {
        var path = Path.Combine(OutputDir, "factory_readmd.docx");
        using (var w = new WordWriter())
        {
            w.AppendHeading("文档标题", 1);
            w.AppendParagraph("段落内容示例。");
            w.Save(path);
        }

        var md = OfficeFactory.ReadMarkdown(path);

        Assert.NotNull(md);
        Assert.Contains("文档标题", md);
        Assert.Contains("段落内容示例", md);
    }

    [Fact, DisplayName("ReadMarkdown_从md文件提取Markdown原文")]
    public void ReadMarkdown_Md_Returns_OriginalMarkdown()
    {
        var path = Path.Combine(OutputDir, "factory_readmd.md");
        var mdContent = "# 测试标题\n\n正文内容。\n\n## 子标题\n\n- 列表项1\n- 列表项2\n";
        File.WriteAllText(path, mdContent, new UTF8Encoding(false));

        var md = OfficeFactory.ReadMarkdown(path);

        Assert.NotNull(md);
        Assert.Contains("测试标题", md);
        Assert.Contains("正文内容", md);
        Assert.Contains("列表项", md);
    }

    [Fact, DisplayName("ReadMarkdown_从流提取xlsx的Markdown表格")]
    public void ReadMarkdown_Stream_Xlsx_Returns_MarkdownTable()
    {
        using var ms = new MemoryStream();
        using (var w = new ExcelWriter(ms))
        {
            w.WriteHeader("Sheet1", new[] { "产品", "价格" });
            w.WriteRow("Sheet1", new Object?[] { "笔记本", 5999m });
            w.Save();
        }

        ms.Position = 0;
        var md = OfficeFactory.ReadMarkdown(ms, ".xlsx");

        Assert.NotNull(md);
        Assert.Contains("产品", md);
        Assert.Contains("笔记本", md);
        Assert.Contains("|", md);
    }

    [Fact, DisplayName("ReadMarkdown_流为null返回null")]
    public void ReadMarkdown_NullStream_Returns_Null()
    {
        Assert.Null(OfficeFactory.ReadMarkdown(null!, ".xlsx"));
    }

    [Fact, DisplayName("ReadMarkdown_不支持格式的流返回null")]
    public void ReadMarkdown_UnsupportedStream_Returns_Null()
    {
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
        Assert.Null(OfficeFactory.ReadMarkdown(ms, ".unknown"));
    }

    #endregion
}
