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
public sealed class MarkdownDocument : ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>文档块列表（顶层）</summary>
    public List<MarkdownBlock> Blocks { get; } = [];

    /// <summary>YAML Front Matter 元数据（解析自文档开头的 --- 块），键值对集合</summary>
    public Dictionary<String, String> FrontMatter { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>缩写映射表（MD05-08），键为缩写文本，值为全称</summary>
    public Dictionary<String, String> Abbreviations { get; } = new(StringComparer.Ordinal);

    /// <summary>引用链接定义（[id]: url），键为引用标识（CommonMark 引用链接）</summary>
    public Dictionary<String, String> References { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>引用链接定义原始行（解析时捕获，含标题与格式），供往返渲染原位输出</summary>
    internal List<String> ReferenceSourceLines { get; } = [];

    /// <summary>引用链接定义原始行号（对应 ReferenceSourceLines），供往返渲染定位</summary>
    internal List<Int32> ReferenceLineIndexes { get; } = [];

    /// <summary>是否启用往返渲染模式（MD06-03），解析→修改→序列化时保留未修改块的原始格式</summary>
    public Boolean Roundtrip { get; set; }
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

    /// <summary>从 Markdown 字符串解析文档（使用指定管线）</summary>
    /// <param name="text">Markdown 文本</param>
    /// <param name="pipeline">处理管线（控制启用/禁用的扩展）</param>
    /// <returns>已解析的文档对象</returns>
    public static MarkdownDocument Parse(String text, MarkdownPipeline pipeline)
    {
        if (String.IsNullOrEmpty(text)) return new MarkdownDocument();
        var parser = new MarkdownParser();
        parser.Pipeline = pipeline;
        return parser.Parse(text);
    }

    /// <summary>从 Markdown 字符串解析文档并启用往返渲染模式（MD06-03）</summary>
    /// <param name="text">Markdown 文本</param>
    /// <returns>已解析的文档对象（<see cref="Roundtrip"/>=true）</returns>
    /// <remarks>往返模式下，解析→修改→序列化时未修改块保留原始格式与块间空行</remarks>
    public static MarkdownDocument ParseRoundtrip(String text)
        => ParseRoundtrip(text, null);

    /// <summary>从 Markdown 字符串解析文档并启用往返渲染模式（MD06-03）</summary>
    /// <param name="text">Markdown 文本</param>
    /// <param name="pipeline">处理管线（可为 null 使用默认全功能）</param>
    /// <returns>已解析的文档对象（<see cref="Roundtrip"/>=true）</returns>
    public static MarkdownDocument ParseRoundtrip(String text, MarkdownPipeline? pipeline)
    {
        if (String.IsNullOrEmpty(text)) return new MarkdownDocument { Roundtrip = true };

        var parser = new MarkdownParser { Roundtrip = true };
        if (pipeline != null) parser.Pipeline = pipeline;
        var doc = parser.Parse(text);
        doc.Roundtrip = true;
        return doc;
    }

    /// <summary>从流解析文档</summary>
    /// <param name="stream">输入流（自动检测编码：BOM/UTF-8/GBK）</param>
    /// <returns>已解析的文档对象</returns>
    public static MarkdownDocument Parse(Stream stream)
    {
        var text = ReadAllText(stream);
        return Parse(text);
    }

    /// <summary>从文件解析文档</summary>
    /// <param name="path">文件路径（自动检测编码：BOM/UTF-8/GBK）</param>
    /// <returns>已解析的文档对象</returns>
    public static MarkdownDocument ParseFile(String path)
    {
        var bytes = File.ReadAllBytes(path);
        return Parse(DecodeText(bytes));
    }

    /// <summary>读取文件文本（编码自动检测：BOM/严格 UTF-8/GBK 兜底），原始文本未经解析</summary>
    /// <param name="path">文件路径</param>
    /// <returns>解码后文本</returns>
    public static String ReadFileText(String path)
        => DecodeText(File.ReadAllBytes(path));

    /// <summary>读取流文本（编码自动检测：BOM/严格 UTF-8/GBK 兜底），原始文本未经解析</summary>
    /// <param name="stream">输入流</param>
    /// <returns>解码后文本</returns>
    public static String ReadStreamText(Stream stream)
        => ReadAllText(stream);

    /// <summary>读取流文本，自动检测编码（BOM/严格 UTF-8/GBK 兜底）</summary>
    /// <param name="stream">输入流</param>
    /// <returns>解码后文本</returns>
    private static String ReadAllText(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return DecodeText(ms.ToArray());
    }

    /// <summary>字节数组解码：BOM → 严格 UTF-8 → GBK 兜底（中文环境）</summary>
    /// <param name="bytes">字节数组</param>
    /// <returns>解码后文本</returns>
    private static String DecodeText(Byte[] bytes)
    {
        if (bytes == null || bytes.Length == 0) return "";

        // BOM 检测（UTF-8 / UTF-16 LE / UTF-16 BE）
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);

        // 严格 UTF-8 校验，避免乱码
        try
        {
            return new UTF8Encoding(false, true).GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            // 非 UTF-8：尝试 GBK（net45 原生支持；netstandard 无提供程序时回退 UTF-8）
            try
            {
                return Encoding.GetEncoding(936).GetString(bytes);
            }
            catch
            {
                return Encoding.UTF8.GetString(bytes);
            }
        }
    }

    /// <summary>从 HTML 字符串解析文档（反向转换，MD04）</summary>
    /// <param name="html">HTML 文本</param>
    /// <param name="options">转换选项（null 使用默认）</param>
    /// <returns>已解析的文档对象</returns>
    /// <remarks>
    /// 支持 h1-h6/p/em/strong/b/i/a/img/ul/ol/li/table/tr/td/th/blockquote/code/pre/hr/br/del 等常用标签。
    /// </remarks>
    public static MarkdownDocument FromHtml(String html, HtmlToMarkdownOptions? options = null)
    {
        if (String.IsNullOrEmpty(html)) return new MarkdownDocument();
        return new HtmlToMarkdownConverter(options).Convert(html);
    }

    /// <summary>从 HTML 流解析文档</summary>
    /// <param name="stream">输入流（UTF-8 或带 BOM）</param>
    /// <param name="options">转换选项（null 使用默认）</param>
    /// <returns>已解析的文档对象</returns>
    public static MarkdownDocument FromHtml(Stream stream, HtmlToMarkdownOptions? options = null)
    {
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return FromHtml(reader.ReadToEnd(), options);
    }
    #endregion

    #region 输出
    /// <summary>序列化为 Markdown 文本</summary>
    /// <param name="bulletChar">无序列表项目符号（默认 "-"），可设为 "*" 或 "+"</param>
    /// <returns>Markdown 字符串</returns>
    public String ToMarkdown(String bulletChar = "-")
    {
        var writer = new MarkdownWriter(bulletChar);
        var body = writer.ToMarkdown(this);

        // 若有 FrontMatter，前置 YAML 块（值必要时加引号，保证往返一致）
        if (FrontMatter.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine("---");
            foreach (var kv in FrontMatter)
                sb.AppendLine($"{kv.Key}: {EscapeYamlValue(kv.Value)}");
            sb.AppendLine("---");
            sb.AppendLine();
            sb.Append(body);
            return sb.ToString();
        }

        return body;
    }

    /// <summary>序列化 YAML 值（含冒号/引号/特殊起始字符/首尾空白时加双引号）</summary>
    /// <param name="value">原始值</param>
    /// <returns>可安全往返的 YAML 值</returns>
    private static String EscapeYamlValue(String value)
    {
        if (String.IsNullOrEmpty(value)) return "\"\"";
        if (value.Trim() != value) return QuoteYaml(value);
        if (value[0] is '#' or '"' or '\'' or '-' or '[' or ']' or '{' or '}' or '&' or '*' or '!' or '|' or '>' or '%' or '@' or '`')
            return QuoteYaml(value);
        if (value.Contains(": ") || value.Contains(" #") || value.Contains('\t'))
            return QuoteYaml(value);
        return value;
    }

    /// <summary>双引号包裹并转义 YAML 值（反斜杠与双引号）</summary>
    /// <param name="value">原始值</param>
    /// <returns>引号包裹后的值</returns>
    private static String QuoteYaml(String value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";

    /// <summary>转换为 HTML 字符串</summary>
    /// <param name="options">HTML 转换选项（null 使用默认）</param>
    /// <returns>HTML 片段（不含 &lt;html&gt;/&lt;body&gt; 包裹）</returns>
    public String ToHtml(MarkdownHtmlOptions? options = null)
    {
        options ??= new MarkdownHtmlOptions();
        // 自动传入文档中的缩写映射
        if (options.Abbreviations == null && Abbreviations.Count > 0)
            options.Abbreviations = Abbreviations;
        var converter = new MarkdownHtmlConverter(options);
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
          .Append("<title>").Append(MarkdownHtmlConverter.HtmlEncode(title)).AppendLine("</title>")
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

    #region 文本提取
    /// <summary>提取纯文本（去除 Markdown 标记）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText()
    {
        if (Blocks.Count == 0) return null;

        var sb = new StringBuilder();
        foreach (var block in Blocks)
        {
            ExtractBlockText(block, sb);
        }
        return sb.ToString();
    }

    /// <summary>提取 Markdown 格式（原始 Markdown）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown() => ToMarkdown();

    private static void ExtractBlockText(MarkdownBlock block, StringBuilder sb)
    {
        switch (block.Type)
        {
            case MarkdownBlockType.Heading:
            case MarkdownBlockType.Paragraph:
                ExtractInlinesText(block.Inlines, sb);
                sb.AppendLine();
                break;
            case MarkdownBlockType.CodeBlock:
                sb.AppendLine(((CodeBlock)block).RawText);
                break;
            case MarkdownBlockType.HtmlBlock:
                sb.AppendLine(((HtmlBlock)block).RawText);
                break;
            case MarkdownBlockType.ThematicBreak:
                sb.AppendLine();
                break;
            case MarkdownBlockType.Table:
                foreach (var row in block.Children)
                {
                    foreach (var cell in row.Children)
                    {
                        ExtractInlinesText(cell.Inlines, sb);
                        sb.Append('\t');
                    }
                    sb.AppendLine();
                }
                break;
            default:
                // 容器块（列表、引用块等）递归
                if (block.Inlines.Count > 0)
                {
                    ExtractInlinesText(block.Inlines, sb);
                    sb.AppendLine();
                }
                foreach (var child in block.Children)
                {
                    ExtractBlockText(child, sb);
                }
                break;
        }
    }

    private static void ExtractInlinesText(List<MarkdownInline> inlines, StringBuilder sb)
    {
        foreach (var inline in inlines)
        {
            if (!String.IsNullOrEmpty(inline.Text))
                sb.Append(inline.Text);
            if (!String.IsNullOrEmpty(inline.Alt))
                sb.Append(inline.Alt);
            if (inline.Children.Count > 0)
                ExtractInlinesText(inline.Children, sb);
        }
    }
    #endregion
}
