using System.Text;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Rtf;

namespace NewLife.Office;

/// <summary>办公文档工厂，提供文件格式校验和读取器创建</summary>
/// <remarks>
/// 支持的格式：xlsx、xls、docx、doc、pptx、ppt、pdf、rtf、ods、epub、vcf、eml、ics、md、xps。
/// <para>用法示例：</para>
/// <code>
/// if (OfficeFactory.IsSupported(".xlsx"))
/// {
///     using var reader = OfficeFactory.CreateReader("data.xlsx");
///     // reader 实际类型为 ExcelReader，可强制转换使用
/// }
/// </code>
/// </remarks>
public static class OfficeFactory
{
    #region 属性
    /// <summary>所有支持的文件后缀（含点号，小写）</summary>
    public static IReadOnlyList<String> SupportedExtensions { get; } =
    [
        ".xlsx", ".xls",
        ".docx", ".doc",
        ".pptx", ".ppt",
        ".pdf",
        ".rtf",
        ".ods",
        ".epub",
        ".vcf",
        ".eml",
        ".ics",
        ".md",
        ".xps",
    ];
    #endregion

    #region 方法
    /// <summary>校验是否支持指定文件后缀</summary>
    /// <param name="extension">文件后缀，可带点号（如 ".xlsx"）或不带（如 "xlsx"）</param>
    /// <returns>是否支持</returns>
    public static Boolean IsSupported(String extension)
    {
        if (String.IsNullOrWhiteSpace(extension)) return false;

        if (!extension.StartsWith("."))
            extension = "." + extension;

        return SupportedExtensions.Contains(extension.ToLowerInvariant());
    }

    /// <summary>根据文件路径创建对应的读取器</summary>
    /// <remarks>
    /// 返回的对象实际类型取决于文件后缀：
    /// <list type="bullet">
    /// <item><description>.xlsx → <see cref="ExcelReader"/></description></item>
    /// <item><description>.xls → <see cref="BiffReader"/></description></item>
    /// <item><description>.docx → <see cref="WordReader"/></description></item>
    /// <item><description>.doc → <see cref="DocReader"/></description></item>
    /// <item><description>.pptx → <see cref="PptxReader"/></description></item>
    /// <item><description>.ppt → <see cref="PptReader"/></description></item>
    /// <item><description>.pdf → <see cref="PdfReader"/></description></item>
    /// <item><description>.rtf → <see cref="RtfDocument"/>（通过 RtfDocument.ParseFile 返回）</description></item>
    /// <item><description>.ods → <see cref="OdsDocument"/>（OdsReader 包装）</description></item>
    /// <item><description>.epub → <see cref="EpubDocument"/></description></item>
    /// <item><description>.vcf → <see cref="VCardDocument"/>（VCardReader 包装）</description></item>
    /// <item><description>.eml → <see cref="EmlMessage"/></description></item>
    /// <item><description>.ics → <see cref="ICalDocument"/></description></item>
    /// <item><description>.md → <see cref="MarkdownDocument"/>（通过 ParseFile 返回）</description></item>
    /// <item><description>.xps → <see cref="XpsDocument"/>（XpsReader 包装）</description></item>
    /// </list>
    /// 调用方应在使用完毕后释放返回对象（若其实现 IDisposable）。
    /// </remarks>
    /// <param name="filePath">文件路径</param>
    /// <returns>读取器对象，实际类型取决于文件后缀</returns>
    /// <exception cref="ArgumentNullException">filePath 为空</exception>
    /// <exception cref="NotSupportedException">不支持的文件后缀</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    public static Object? CreateReader(String filePath)
    {
        if (String.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

        var fullPath = filePath.GetFullPath();
        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"文件不存在: {fullPath}", fullPath);

        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" => new ExcelReader(fullPath),
            ".xls" => new BiffReader(fullPath),
            ".docx" => new WordReader(fullPath),
            ".doc" => new DocReader(fullPath),
            ".pptx" => new PptxReader(fullPath),
            ".ppt" => new PptReader(fullPath),
            ".pdf" => new PdfReader(fullPath),
            ".rtf" => RtfDocument.ParseFile(fullPath),
            ".ods" => new OdsDocument(OdsReader.ReadFile(fullPath)),
            ".epub" => new EpubReader().Read(fullPath),
            ".vcf" => new VCardDocument(new VCardReader().ReadAll(fullPath)),
            ".eml" => new EmlReader().Read(fullPath),
            ".ics" => new ICalReader().Read(fullPath),
            ".md" => MarkdownDocument.ParseFile(fullPath),
            ".xps" => new XpsDocument(new XpsReader().Read(fullPath)),
            _ => null,
        };
    }

    /// <summary>根据数据流和扩展名创建对应的读取器</summary>
    /// <param name="stream">数据流</param>
    /// <param name="extension">文件后缀，可带点号（如 ".xlsx"）或不带（如 "xlsx"）</param>
    /// <returns>读取器对象，不支持的格式返回 null</returns>
    public static Object? CreateReader(Stream stream, String extension)
    {
        if (stream == null) return null;
        if (String.IsNullOrWhiteSpace(extension)) return null;

        if (!extension.StartsWith("."))
            extension = "." + extension;

        return extension.ToLowerInvariant() switch
        {
            ".xlsx" => new ExcelReader(stream, Encoding.UTF8),
            ".xls" => new BiffReader(stream),
            ".docx" => new WordReader(stream),
            ".doc" => new DocReader(stream),
            ".pptx" => new PptxReader(stream),
            ".ppt" => new PptReader(stream),
            ".pdf" => new PdfReader(stream),
            ".rtf" => RtfDocument.Parse(stream),
            ".ods" => new OdsDocument(OdsReader.Read(stream)),
            ".epub" => new EpubReader().Read(stream),
            ".vcf" => new VCardDocument(new VCardReader().ReadAll(stream)),
            ".eml" => new EmlReader().Read(stream),
            ".ics" => new ICalReader().Read(stream),
            ".md" => MarkdownDocument.Parse(stream),
            ".xps" => new XpsDocument(new XpsReader().Read(stream)),
            _ => null,
        };
    }
    #endregion

    #region 文本提取
    /// <summary>从文件提取纯文本</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>纯文本，不支持或无内容返回 null</returns>
    public static String? ReadText(String filePath)
    {
        if (String.IsNullOrWhiteSpace(filePath)) return null;

        var reader = CreateReader(filePath);
        try
        {
            return (reader as ITextExtractable)?.ExtractText();
        }
        finally
        {
            (reader as IDisposable)?.Dispose();
        }
    }

    /// <summary>从数据流提取纯文本</summary>
    /// <param name="stream">数据流</param>
    /// <param name="extension">文件后缀，可带点号（如 ".xlsx"）或不带（如 "xlsx"）</param>
    /// <returns>纯文本，不支持或无内容返回 null</returns>
    public static String? ReadText(Stream stream, String extension)
    {
        var reader = CreateReader(stream, extension);
        if (reader == null) return null;

        try
        {
            return (reader as ITextExtractable)?.ExtractText();
        }
        finally
        {
            (reader as IDisposable)?.Dispose();
        }
    }

    /// <summary>从文件提取 Markdown 格式文本</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>Markdown 文本，不支持或无内容返回 null</returns>
    public static String? ReadMarkdown(String filePath)
    {
        if (String.IsNullOrWhiteSpace(filePath)) return null;

        var reader = CreateReader(filePath);
        try
        {
            return (reader as IMarkdownExtractable)?.ExtractMarkdown();
        }
        finally
        {
            (reader as IDisposable)?.Dispose();
        }
    }

    /// <summary>从数据流提取 Markdown 格式文本</summary>
    /// <param name="stream">数据流</param>
    /// <param name="extension">文件后缀，可带点号（如 ".xlsx"）或不带（如 "xlsx"）</param>
    /// <returns>Markdown 文本，不支持或无内容返回 null</returns>
    public static String? ReadMarkdown(Stream stream, String extension)
    {
        var reader = CreateReader(stream, extension);
        if (reader == null) return null;

        try
        {
            return (reader as IMarkdownExtractable)?.ExtractMarkdown();
        }
        finally
        {
            (reader as IDisposable)?.Dispose();
        }
    }
    #endregion
}
