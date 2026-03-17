using System.IO;
using System.Text;

namespace NewLife.Office.Markdown;

/// <summary>Markdown 文档对象模型</summary>
/// <remarks>
/// 表示解析后的完整 Markdown 文档，提供从字符串/流解析，或序列化回 Markdown 的功能。
/// <para>读取示例：</para>
/// <code>
/// var doc = MarkdownDocument.Parse("# Hello\nWorld");
/// var html = doc.ToHtml();
/// </code>
/// <para>创建示例：</para>
/// <code>
/// var doc = new MarkdownDocument();
/// doc.Blocks.Add(MarkdownBlock.CreateHeading(1, ...));
/// var md = doc.ToMarkdown();
/// </code>
/// </remarks>
public sealed class MarkdownDocument
{
    #region 属性
    /// <summary>文档块列表（顶层）</summary>
    public List<MarkdownBlock> Blocks { get; } = [];
    #endregion

    #region 解析
    /// <summary>从 Markdown 字符串解析文档</summary>
    /// <param name="text">Markdown 文本</param>
    /// <returns>已解析的文档对象</returns>
    public static MarkdownDocument Parse(String text)
    {
        if (String.IsNullOrEmpty(text)) return new MarkdownDocument();
        return new MarkdownParser().Parse(text);
    }

    /// <summary>从流解析文档</summary>
    /// <param name="stream">输入流（UTF-8 或带 BOM）</param>
    /// <returns>已解析的文档对象</returns>
    public static MarkdownDocument Parse(Stream stream)
    {
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return Parse(reader.ReadToEnd());
    }

    /// <summary>从文件解析文档</summary>
    /// <param name="path">文件路径</param>
    /// <returns>已解析的文档对象</returns>
    public static MarkdownDocument ParseFile(String path)
    {
        var text = File.ReadAllText(path, Encoding.UTF8);
        return Parse(text);
    }
    #endregion

    #region 输出
    /// <summary>序列化为 Markdown 文本</summary>
    /// <returns>Markdown 字符串</returns>
    public String ToMarkdown()
    {
        var writer = new MarkdownWriter();
        return writer.ToMarkdown(this);
    }

    /// <summary>转换为 HTML 字符串</summary>
    /// <param name="options">HTML 转换选项（null 使用默认）</param>
    /// <returns>HTML 片段（不含 &lt;html&gt;/&lt;body&gt; 包裹）</returns>
    public String ToHtml(MarkdownHtmlOptions? options = null)
    {
        var converter = new MarkdownHtmlConverter(options ?? new MarkdownHtmlOptions());
        return converter.Convert(this);
    }

    /// <summary>转换为完整 HTML 页面</summary>
    /// <param name="title">页面标题</param>
    /// <param name="options">HTML 转换选项（null 使用默认）</param>
    /// <returns>完整 HTML 文档字符串</returns>
    public String ToHtmlPage(String title = "Document", MarkdownHtmlOptions? options = null)
    {
        var body = ToHtml(options);
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>")
          .AppendLine("<html lang=\"zh\">")
          .AppendLine("<head>")
          .AppendLine("<meta charset=\"UTF-8\">")
          .AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">")
          .Append("<title>").Append(HtmlEncode(title)).AppendLine("</title>")
          .AppendLine("<style>")
          .AppendLine("body{font-family:-apple-system,BlinkMacSystemFont,\"Segoe UI\",Helvetica,Arial,sans-serif;font-size:16px;line-height:1.6;max-width:900px;margin:0 auto;padding:2rem;color:#1a1a2e}")
          .AppendLine("h1,h2,h3,h4,h5,h6{margin-top:1.5rem;margin-bottom:.5rem;font-weight:600}")
          .AppendLine("code{background:#f6f8fa;padding:.2em .4em;border-radius:3px;font-size:.9em;font-family:monospace}")
          .AppendLine("pre{background:#f6f8fa;padding:1rem;border-radius:6px;overflow:auto}")
          .AppendLine("pre code{background:none;padding:0}")
          .AppendLine("blockquote{margin:0;padding-left:1rem;border-left:4px solid #dfe2e5;color:#6a737d}")
          .AppendLine("table{border-collapse:collapse;width:100%}")
          .AppendLine("th,td{border:1px solid #dfe2e5;padding:.5rem .75rem;text-align:left}")
          .AppendLine("th{background:#f6f8fa;font-weight:600}")
          .AppendLine("tr:nth-child(even){background:#fafbfc}")
          .AppendLine("a{color:#0366d6;text-decoration:none}")
          .AppendLine("a:hover{text-decoration:underline}")
          .AppendLine("img{max-width:100%}")
          .AppendLine("hr{border:none;border-top:1px solid #e1e4e8;margin:1.5rem 0}")
          .AppendLine(".task-list-item{list-style:none;padding-left:.2rem}")
          .AppendLine(".task-list-item input{margin-right:.5rem}")
          .AppendLine("del{opacity:.7}")
          .AppendLine("</style>")
          .AppendLine("</head>")
          .AppendLine("<body>")
          .AppendLine(body)
          .AppendLine("</body>")
          .AppendLine("</html>");
        return sb.ToString();
    }

    private static String HtmlEncode(String s) => s
        .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
        .Replace("\"", "&quot;");
    #endregion

    #region Word/PDF 转换
    /// <summary>将文档转换为 .docx 字节数组（MD03-02）</summary>
    /// <returns>docx 字节数组</returns>
    public Byte[] ToWord()
    {
        return new MarkdownWordConverter().ToBytes(this);
    }

    /// <summary>将文档保存为 .docx 文件（MD03-02）</summary>
    /// <param name="path">目标文件路径</param>
    public void SaveWord(String path)
    {
        var bytes = ToWord();
        File.WriteAllBytes(path, bytes);
    }

    /// <summary>将文档写入流（Word 格式，MD03-02）</summary>
    /// <param name="stream">目标可写流</param>
    public void SaveWord(Stream stream)
    {
        var bytes = ToWord();
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>将文档转换为 PDF 字节数组（MD03-03）</summary>
    /// <returns>PDF 字节数组</returns>
    public Byte[] ToPdf()
    {
        return new MarkdownPdfConverter().ToBytes(this);
    }

    /// <summary>将文档保存为 PDF 文件（MD03-03）</summary>
    /// <param name="path">目标文件路径</param>
    public void SavePdf(String path)
    {
        var bytes = ToPdf();
        File.WriteAllBytes(path, bytes);
    }

    /// <summary>将文档写入流（PDF 格式，MD03-03）</summary>
    /// <param name="stream">目标可写流</param>
    public void SavePdf(Stream stream)
    {
        var bytes = ToPdf();
        stream.Write(bytes, 0, bytes.Length);
    }
    #endregion
}

/// <summary>Markdown → HTML 转换选项</summary>
public sealed class MarkdownHtmlOptions
{
    /// <summary>是否为代码块添加 language-xxx CSS 类（默认 true）</summary>
    public Boolean AddLanguageClass { get; set; } = true;

    /// <summary>是否在链接上添加 target="_blank"（默认 false）</summary>
    public Boolean ExternalLinkTarget { get; set; }

    /// <summary>是否对链接添加 rel="noopener noreferrer"（默认 false）</summary>
    public Boolean SafeLinks { get; set; }
}
