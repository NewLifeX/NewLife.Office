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
    /// <item><description>.ods → <see cref="OdsReader"/>（静态方法包装）</description></item>
    /// <item><description>.epub → <see cref="EpubReader"/></description></item>
    /// <item><description>.vcf → <see cref="VCardReader"/></description></item>
    /// <item><description>.eml → <see cref="EmlReader"/></description></item>
    /// <item><description>.ics → <see cref="ICalReader"/></description></item>
    /// <item><description>.md → MarkdownDocument（通过 ParseFile 返回）</description></item>
    /// <item><description>.xps → <see cref="XpsReader"/></description></item>
    /// </list>
    /// 调用方应在使用完毕后释放返回对象（若其实现 IDisposable）。
    /// </remarks>
    /// <param name="filePath">文件路径</param>
    /// <returns>读取器对象，实际类型取决于文件后缀</returns>
    /// <exception cref="ArgumentNullException">filePath 为空</exception>
    /// <exception cref="NotSupportedException">不支持的文件后缀</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    public static Object CreateReader(String filePath)
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
            ".ods" => OdsReader.ReadFile(fullPath),
            ".epub" => new EpubReader().Read(fullPath),
            ".vcf" => new VCardReader(),
            ".eml" => new EmlReader(),
            ".ics" => new ICalReader(),
            ".md" => MarkdownDocument.ParseFile(fullPath),
            ".xps" => new XpsReader(),
            _ => throw new NotSupportedException($"不支持的文件格式: {ext}")
        };
    }
    #endregion
}
