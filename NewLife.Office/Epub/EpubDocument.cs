using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office;

/// <summary>EPUB 电子书文档模型</summary>
public class EpubDocument : ITextExtractable, IMarkdownExtractable
{
    #region 属性

    /// <summary>书名</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>作者</summary>
    public String Author { get; set; } = String.Empty;

    /// <summary>语言（BCP 47，如 zh-CN/en）</summary>
    public String Language { get; set; } = "zh-CN";

    /// <summary>出版商</summary>
    public String Publisher { get; set; } = String.Empty;

    /// <summary>描述/简介</summary>
    public String Description { get; set; } = String.Empty;

    /// <summary>发布日期</summary>
    public String PublishDate { get; set; } = String.Empty;

    /// <summary>ISBN 或其他唯一标识</summary>
    public String Identifier { get; set; } = String.Empty;

    /// <summary>封面图片数据（PNG/JPEG）</summary>
    public Byte[]? Cover { get; set; }

    /// <summary>封面图片 MIME 类型</summary>
    public String CoverMediaType { get; set; } = "image/jpeg";

    /// <summary>章节列表（按顺序）</summary>
    public List<EpubChapter> Chapters { get; set; } = [];

    /// <summary>自定义 CSS 样式表内容</summary>
    public String StyleSheet { get; set; } = String.Empty;

    #endregion

    #region 文本提取
    /// <summary>提取纯文本（去除 HTML 标签）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText()
    {
        if (Chapters == null || Chapters.Count == 0) return null;

        var sb = new StringBuilder();
        ExtractChaptersText(Chapters, sb);
        return sb.ToString();
    }

    /// <summary>提取 Markdown 格式（保留章节标题结构）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown()
    {
        if (Chapters == null || Chapters.Count == 0) return null;

        var sb = new StringBuilder();
        ExtractChaptersMarkdown(Chapters, sb, 1);
        return sb.ToString();
    }

    private static void ExtractChaptersText(List<EpubChapter> chapters, StringBuilder sb)
    {
        foreach (var ch in chapters)
        {
            if (!String.IsNullOrEmpty(ch.Title))
                sb.AppendLine(ch.Title);
            if (!String.IsNullOrEmpty(ch.Content))
                sb.AppendLine(StripHtml(ch.Content));
            sb.AppendLine();

            if (ch.Children != null && ch.Children.Count > 0)
                ExtractChaptersText(ch.Children, sb);
        }
    }

    private static void ExtractChaptersMarkdown(List<EpubChapter> chapters, StringBuilder sb, Int32 headingLevel)
    {
        var prefix = new String('#', Math.Min(headingLevel, 6));
        foreach (var ch in chapters)
        {
            if (!String.IsNullOrEmpty(ch.Title))
            {
                sb.AppendLine($"{prefix} {ch.Title}");
                sb.AppendLine();
            }
            if (!String.IsNullOrEmpty(ch.Content))
            {
                sb.AppendLine(StripHtml(ch.Content));
                sb.AppendLine();
            }

            if (ch.Children != null && ch.Children.Count > 0)
                ExtractChaptersMarkdown(ch.Children, sb, headingLevel + 1);
        }
    }

    private static String StripHtml(String html)
    {
        if (String.IsNullOrEmpty(html)) return "";
        // 移除 HTML 标签
        var text = Regex.Replace(html, "<[^>]+>", " ");
        // 解码常见 HTML 实体
        text = text.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
                   .Replace("&quot;", "\"").Replace("&#39;", "'").Replace("&nbsp;", " ");
        // 合并多余空白
        text = Regex.Replace(text, @"\s+", " ").Trim();
        return text;
    }
    #endregion
}

/// <summary>EPUB 章节</summary>
public class EpubChapter
{
    #region 属性

    /// <summary>章节标题</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>章节 HTML 内容（完整 XHTML 片段或正文部分）</summary>
    public String Content { get; set; } = String.Empty;

    /// <summary>文件名（不含路径，如 chapter01.xhtml）</summary>
    public String FileName { get; set; } = String.Empty;

    /// <summary>子章节（嵌套 TOC 支持）</summary>
    public List<EpubChapter> Children { get; set; } = [];

    #endregion
}
