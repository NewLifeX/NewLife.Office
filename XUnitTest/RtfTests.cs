using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using NewLife.Office.Rtf;
using Xunit;

namespace XUnitTest;

/// <summary>RTF 模块单元测试</summary>
public class RtfTests
{
    #region 解析测试 — 基础
    [Fact]
    [DisplayName("解析空字符串返回空文档")]
    public void Parse_Empty_ReturnsEmpty()
    {
        var doc = RtfDocument.Parse("");
        Assert.NotNull(doc);
        Assert.Empty(doc.Blocks);
    }

    [Fact]
    [DisplayName("解析非RTF字符串返回空文档")]
    public void Parse_NonRtf_ReturnsEmpty()
    {
        var doc = RtfDocument.Parse("Hello world, not RTF");
        Assert.NotNull(doc);
        Assert.Empty(doc.Blocks);
    }

    [Fact]
    [DisplayName("解析简单文本段落")]
    public void Parse_SimpleParagraph_TextExtracted()
    {
        var rtf = @"{\rtf1\ansi Hello World\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("Hello World", text);
    }

    [Fact]
    [DisplayName("解析多段落")]
    public void Parse_MultipleParagraphs_AllExtracted()
    {
        var rtf = @"{\rtf1\ansi First\par Second\par Third\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("First", text);
        Assert.Contains("Second", text);
        Assert.Contains("Third", text);
    }
    #endregion

    #region 解析测试 — 格式
    [Fact]
    [DisplayName("解析粗体格式")]
    public void Parse_BoldText_RunHasBold()
    {
        var rtf = @"{\rtf1\ansi {\b Bold Text}\par}";
        var doc = RtfDocument.Parse(rtf);
        var paras = doc.Paragraphs.ToList();
        Assert.NotEmpty(paras);
        var boldRun = paras.SelectMany(p => p.Runs).FirstOrDefault(r => r.Bold);
        Assert.NotNull(boldRun);
        Assert.Contains("Bold", boldRun.Text);
    }

    [Fact]
    [DisplayName("解析斜体格式")]
    public void Parse_ItalicText_RunHasItalic()
    {
        var rtf = @"{\rtf1\ansi {\i Italic Text}\par}";
        var doc = RtfDocument.Parse(rtf);
        var italicRun = doc.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.Italic);
        Assert.NotNull(italicRun);
    }

    [Fact]
    [DisplayName("解析下划线格式")]
    public void Parse_UnderlineText_RunHasUnderline()
    {
        var rtf = @"{\rtf1\ansi {\ul Underline}\par}";
        var doc = RtfDocument.Parse(rtf);
        var ulRun = doc.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.Underline);
        Assert.NotNull(ulRun);
    }

    [Fact]
    [DisplayName("解析字号")]
    public void Parse_FontSize_RunHasFontSize()
    {
        var rtf = @"{\rtf1\ansi {\fs40 Large Text}\par}";
        var doc = RtfDocument.Parse(rtf);
        var run = doc.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.FontSize == 40);
        Assert.NotNull(run);
    }

    [Fact]
    [DisplayName("解析段落对齐")]
    public void Parse_CenterAlign_ParagraphCentered()
    {
        var rtf = @"{\rtf1\ansi \pard\qc Center Text\par}";
        var doc = RtfDocument.Parse(rtf);
        var paras = doc.Paragraphs.ToList();
        Assert.NotEmpty(paras);
        var centerPara = paras.FirstOrDefault(p => p.Alignment == RtfAlignment.Center);
        Assert.NotNull(centerPara);
    }
    #endregion

    #region 解析测试 — 颜色与字体
    [Fact]
    [DisplayName("解析颜色表并应用颜色")]
    public void Parse_ColorTable_RunHasForeColor()
    {
        // cf1 应引用颜色表第1个颜色（红色）
        var rtf = @"{\rtf1\ansi{\colortbl;\red255\green0\blue0;}\pard{\cf1 Red Text}\par}";
        var doc = RtfDocument.Parse(rtf);
        var run = doc.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.ForeColor >= 0);
        Assert.NotNull(run);
        // 颜色应该是红色 0xFF0000
        Assert.Equal(0xFF0000, run.ForeColor);
    }

    [Fact]
    [DisplayName("解析字体表并应用字体名")]
    public void Parse_FontTable_RunHasFontName()
    {
        var rtf = @"{\rtf1\ansi{\fonttbl{\f0\froman Arial;}}{\f0 Arial Text}\par}";
        var doc = RtfDocument.Parse(rtf);
        var run = doc.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => !String.IsNullOrEmpty(r.FontName));
        Assert.NotNull(run);
        Assert.Contains("Arial", run.FontName);
    }
    #endregion

    #region 解析测试 — 特殊字符
    [Fact]
    [DisplayName("解析十六进制ANSI字符")]
    public void Parse_HexAnsi_CharDecoded()
    {
        // \'e9 = é in Windows-1252
        var rtf = @"{\rtf1\ansi caf\'e9\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("café", text);
    }

    [Fact]
    [DisplayName("解析Unicode字符")]
    public void Parse_Unicode_CharDecoded()
    {
        // \u8364? = € (U+20AC)
        var rtf = @"{\rtf1\ansi \u8364?\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("€", text);
    }

    [Fact]
    [DisplayName("解析换行符\\line")]
    public void Parse_LineBreak_RunMarkedAsBreak()
    {
        var rtf = @"{\rtf1\ansi Line1\line Line2\par}";
        var doc = RtfDocument.Parse(rtf);
        var hasBreak = doc.Paragraphs.SelectMany(p => p.Runs).Any(r => r.IsLineBreak);
        Assert.True(hasBreak);
    }

    [Fact]
    [DisplayName("解析制表符\\tab")]
    public void Parse_Tab_AppendedToText()
    {
        var rtf = @"{\rtf1\ansi Col1\tab Col2\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains('\t', text);
    }
    #endregion

    #region 解析测试 — 文档属性
    [Fact]
    [DisplayName("解析文档标题")]
    public void Parse_InfoTitle_TitleSet()
    {
        var rtf = @"{\rtf1\ansi{\info{\title My Doc}}Hello\par}";
        var doc = RtfDocument.Parse(rtf);
        Assert.Equal("My Doc", doc.Title);
    }

    [Fact]
    [DisplayName("解析文档作者")]
    public void Parse_InfoAuthor_AuthorSet()
    {
        var rtf = @"{\rtf1\ansi{\info{\author John}}Hello\par}";
        var doc = RtfDocument.Parse(rtf);
        Assert.Equal("John", doc.Author);
    }
    #endregion

    #region 解析测试 — 跳过目标组
    [Fact]
    [DisplayName("跳过可选目标组\\*")]
    public void Parse_SkipOptionalDestination_NoGarbage()
    {
        var rtf = @"{\rtf1\ansi {\*\generator Word;} Real Text\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("Real Text", text);
        Assert.DoesNotContain("generator", text);
    }

    [Fact]
    [DisplayName("跳过样式表")]
    public void Parse_SkipStylesheet_NoGarbage()
    {
        var rtf = @"{\rtf1\ansi{\stylesheet{\s0 Normal;}}Content\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("Content", text);
        Assert.DoesNotContain("Normal", text);
    }
    #endregion

    #region 写入测试 — 基础
    [Fact]
    [DisplayName("写入简单段落生成合法RTF")]
    public void Write_SimpleParagraph_ValidRtf()
    {
        var writer = new RtfWriter();
        writer.AddParagraph("Hello RTF");
        var rtf = writer.ToString();
        Assert.StartsWith("{\\rtf1", rtf);
        Assert.Contains("Hello RTF", rtf);
        Assert.Contains("\\par", rtf);
    }

    [Fact]
    [DisplayName("写入多段落")]
    public void Write_MultiParagraph_AllPresent()
    {
        var writer = new RtfWriter();
        writer.AddParagraph("First").AddParagraph("Second").AddParagraph("Third");
        var rtf = writer.ToString();
        Assert.Contains("First", rtf);
        Assert.Contains("Second", rtf);
        Assert.Contains("Third", rtf);
    }

    [Fact]
    [DisplayName("写入带粗体格式段落")]
    public void Write_BoldParagraph_BoldTagPresent()
    {
        var para = new RtfParagraph();
        para.Runs.Add(new RtfRun { Text = "Bold", Bold = true });
        var writer = new RtfWriter();
        writer.AddParagraph(para);
        var rtf = writer.ToString();
        Assert.Contains("\\b", rtf);
        Assert.Contains("Bold", rtf);
    }

    [Fact]
    [DisplayName("写入带颜色格式")]
    public void Write_ColoredText_ColorTablePresent()
    {
        var para = new RtfParagraph();
        para.Runs.Add(new RtfRun { Text = "Red", ForeColor = 0xFF0000 });
        var writer = new RtfWriter();
        writer.AddParagraph(para);
        var rtf = writer.ToString();
        Assert.Contains("\\colortbl", rtf);
        Assert.Contains("\\red255", rtf);
        Assert.Contains("\\cf", rtf);
    }

    [Fact]
    [DisplayName("写入文档属性")]
    public void Write_DocProperties_InfoBlockPresent()
    {
        var writer = new RtfWriter { Title = "Test", Author = "Alice" };
        writer.AddParagraph("Content");
        var rtf = writer.ToString();
        Assert.Contains("\\info", rtf);
        Assert.Contains("Test", rtf);
        Assert.Contains("Alice", rtf);
    }
    #endregion

    #region 写入测试 — 表格
    [Fact]
    [DisplayName("写入简单表格")]
    public void Write_Table_TableTagsPresent()
    {
        var writer = new RtfWriter();
        writer.AddTable(new[]
        {
            new[] { "R1C1", "R1C2" },
            new[] { "R2C1", "R2C2" },
        });
        var rtf = writer.ToString();
        Assert.Contains("\\trowd", rtf);
        Assert.Contains("\\cell", rtf);
        Assert.Contains("\\row", rtf);
        Assert.Contains("R1C1", rtf);
        Assert.Contains("R2C2", rtf);
    }
    #endregion

    #region 写入测试 — 保存/流
    [Fact]
    [DisplayName("保存到文件后可读取")]
    public void Write_SaveFile_CanReadBack()
    {
        var path = Path.Combine(Path.GetTempPath(), "test_rtf.rtf");
        try
        {
            var writer = new RtfWriter();
            writer.AddParagraph("Save Test");
            writer.Save(path);
            Assert.True(File.Exists(path));
            var content = File.ReadAllText(path);
            Assert.Contains("Save Test", content);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    [DisplayName("保存到流后可读取")]
    public void Write_SaveStream_CanReadBack()
    {
        var writer = new RtfWriter();
        writer.AddParagraph("Stream Test");
        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;
        var content = new System.IO.StreamReader(ms).ReadToEnd();
        Assert.Contains("Stream Test", content);
    }
    #endregion

    #region 往返测试
    [Fact]
    [DisplayName("写再读纯文本往返")]
    public void RoundTrip_PlainText_Preserved()
    {
        var writer = new RtfWriter();
        writer.AddParagraph("Round Trip Text");
        var rtf = writer.ToString();
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("Round Trip Text", text);
    }

    [Fact]
    [DisplayName("写再读表格内容往返")]
    public void RoundTrip_Table_CellsPreserved()
    {
        var writer = new RtfWriter();
        writer.AddTable(new[] { new[] { "Cell A", "Cell B" } });
        var rtf = writer.ToString();
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("Cell A", text);
        Assert.Contains("Cell B", text);
    }
    #endregion

    #region 模板填充测试
    [Fact]
    [DisplayName("模板填充替换占位符")]
    public void FillTemplate_ReplacesPlaceholders()
    {
        var rtf = @"{\rtf1\ansi Hello {{Name}}, you are {{Age}} years old.\par}";
        var values = new Dictionary<String, String> { ["Name"] = "Alice", ["Age"] = "30" };
        var result = RtfWriter.FillTemplate(rtf, values);
        Assert.Contains("Alice", result);
        Assert.Contains("30", result);
        Assert.DoesNotContain("{{Name}}", result);
    }

    [Fact]
    [DisplayName("模板填充保留未匹配占位符")]
    public void FillTemplate_UnknownKey_Preserved()
    {
        var rtf = @"{\rtf1\ansi {{Known}} and {{Unknown}}\par}";
        var values = new Dictionary<String, String> { ["Known"] = "X" };
        var result = RtfWriter.FillTemplate(rtf, values);
        Assert.Contains("X", result);
        Assert.Contains("{{Unknown}}", result);
    }

    [Fact]
    [DisplayName("模板填充空值不崩溃")]
    public void FillTemplate_NullOrEmpty_ReturnsOriginal()
    {
        Assert.Equal("", RtfWriter.FillTemplate("", null));
        Assert.Null(RtfWriter.FillTemplate(null, new Dictionary<String, String>()));
    }

    [Fact]
    [DisplayName("GetPlainText 提取所有文本")]
    public void GetPlainText_AllBlocks_AllText()
    {
        var rtf = @"{\rtf1\ansi Para1\par Para2\par}";
        var doc = RtfDocument.Parse(rtf);
        var text = doc.GetPlainText();
        Assert.Contains("Para1", text);
        Assert.Contains("Para2", text);
    }
    #endregion
}
