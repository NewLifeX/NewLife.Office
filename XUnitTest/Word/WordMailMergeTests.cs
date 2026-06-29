using System.IO.Compression;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Word;

/// <summary>WordTemplate.MailMerge 单元测试 — MERGEFIELD 域执行引擎</summary>
public class WordMailMergeTests
{
    #region 基础合并
    [Fact(DisplayName = "MailMerge—单字段替换")]
    public void MailMerge_SingleField()
    {
        // 创建含 MERGEFIELD 的模板
        var templatePath = CreateMergeTemplate("FirstName");

        var data = new Dictionary<String, Object?>
        {
            ["FirstName"] = "张三"
        };

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_single.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, data);

            Assert.True(File.Exists(outputPath));

            // 验证输出文档中包含替换后的值
            var text = ReadText(outputPath);
            Assert.Contains("张三", text);
            Assert.DoesNotContain("«FirstName»", text);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }

    [Fact(DisplayName = "MailMerge—多字段替换")]
    public void MailMerge_MultipleFields()
    {
        var templatePath = CreateMergeTemplate("FirstName", "LastName", "Company");

        var data = new Dictionary<String, Object?>
        {
            ["FirstName"] = "张三",
            ["LastName"] = "李四",
            ["Company"] = "新生命团队"
        };

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_multi.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, data);

            var text = ReadText(outputPath);
            Assert.Contains("张三", text);
            Assert.Contains("李四", text);
            Assert.Contains("新生命团队", text);
            Assert.DoesNotContain("«FirstName»", text);
            Assert.DoesNotContain("«LastName»", text);
            Assert.DoesNotContain("«Company»", text);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }

    [Fact(DisplayName = "MailMerge—缺失字段保留原占位")]
    public void MailMerge_MissingField()
    {
        var templatePath = CreateMergeTemplate("FirstName");

        var data = new Dictionary<String, Object?>
        {
            ["SomeOther"] = "hello"
        };

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_missing.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, data);

            var text = ReadText(outputPath);
            // 未匹配字段保留原显示文本
            Assert.Contains("FirstName", text);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }
    #endregion

    #region 多记录合并
    [Fact(DisplayName = "MailMerge—多记录合并（2条）")]
    public void MailMerge_MultipleRecords()
    {
        var templatePath = CreateMergeTemplate("Name", "Amount");

        var records = new[]
        {
            new Dictionary<String, Object?> { ["Name"] = "张三", ["Amount"] = "100.00" },
            new Dictionary<String, Object?> { ["Name"] = "李四", ["Amount"] = "200.50" },
        };

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_multi_record.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, records);

            Assert.True(File.Exists(outputPath));
            // 验证 XML 格式有效
            using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
            var entry = zip.GetEntry("word/document.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("张三", xml);
            Assert.Contains("李四", xml);
            // 确保 XML 格式良好
            Assert.True(xml.StartsWith("<?xml") || xml.StartsWith("<w:document"));
            Assert.True(xml.EndsWith("</w:document>") || xml.Contains("</w:document>"));
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }

    [Fact(DisplayName = "MailMerge—多记录合并（1条不回退）")]
    public void MailMerge_MultipleRecords_Single()
    {
        var templatePath = CreateMergeTemplate("Name");

        var records = new[]
        {
            new Dictionary<String, Object?> { ["Name"] = "张三" },
        };

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_multi_single.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, records);

            Assert.True(File.Exists(outputPath));
            var text = ReadText(outputPath);
            Assert.Contains("张三", text);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }

    [Fact(DisplayName = "MailMerge—空记录列表不抛异常")]
    public void MailMerge_EmptyRecords()
    {
        var templatePath = CreateMergeTemplate("Name");

        var records = Array.Empty<Dictionary<String, Object?>>();

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_empty.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, records);

            // 空记录应创建最小文档
            Assert.True(File.Exists(outputPath));
            var info = new FileInfo(outputPath);
            Assert.True(info.Length > 0);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }
    #endregion

    #region 边界场景
    [Fact(DisplayName = "MailMerge—空数据字典不修改文档")]
    public void MailMerge_EmptyData()
    {
        var templatePath = CreateMergeTemplate("FirstName");

        var data = new Dictionary<String, Object?>();

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_empty_data.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, data);

            var text = ReadText(outputPath);
            // 空数据时应保留原占位符
            Assert.Contains("FirstName", text);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }

    [Fact(DisplayName = "MailMerge—Null值替换为空字符串")]
    public void MailMerge_NullValue()
    {
        var templatePath = CreateMergeTemplate("Optional");

        var data = new Dictionary<String, Object?>
        {
            ["Optional"] = null
        };

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_null.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, data);

            var text = ReadText(outputPath);
            Assert.DoesNotContain("«Optional»", text);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }

    [Fact(DisplayName = "MailMerge—XML特殊字符正确转义")]
    public void MailMerge_XmlEscape()
    {
        var templatePath = CreateMergeTemplate("Content");

        var data = new Dictionary<String, Object?>
        {
            ["Content"] = "A < B & C > D"
        };

        var outputPath = Path.Combine(Path.GetTempPath(), "merge_escape.docx");
        try
        {
            var template = new WordTemplate(templatePath);
            template.MailMerge(outputPath, data);

            var xml = ReadDocumentXml(outputPath);
            Assert.Contains("A &lt; B &amp; C &gt; D", xml);
        }
        finally
        {
            DeleteIfExists(templatePath);
            DeleteIfExists(outputPath);
        }
    }
    #endregion

    #region 辅助方法
    /// <summary>创建含 MERGEFIELD 的 docx 模板</summary>
    private static String CreateMergeTemplate(params String[] fieldNames)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        using var writer = new WordWriter();
        writer.AppendParagraph("邮件合并测试文档");
        foreach (var name in fieldNames)
        {
            writer.AppendMergeField(name);
        }
        writer.Save(tempPath);
        return tempPath;
    }

    /// <summary>读取 docx 中的纯文本内容</summary>
    private static String ReadText(String docxPath)
    {
        using var reader = new WordReader(docxPath);
        return reader.ExtractText() ?? String.Empty;
    }

    /// <summary>读取 docx 中 word/document.xml 原始内容</summary>
    private static String ReadDocumentXml(String docxPath)
    {
        using var fs = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        var entry = zip.GetEntry("word/document.xml");
        if (entry == null) return String.Empty;
        using var stream = entry.Open();
        using var sr = new StreamReader(stream, Encoding.UTF8);
        return sr.ReadToEnd();
    }

    private static void DeleteIfExists(String path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }
    #endregion
}
