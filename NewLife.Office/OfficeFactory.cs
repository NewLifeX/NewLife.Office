using System.IO.Compression;
using System.Text;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Ole2;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using NewLife.Office.Xps;

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
    /// <item><description>.eml → <see cref="Message"/></description></item>
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

    #region 内容识别（GEN-1，对标 anydoc）
    /// <summary>按文件头 magic bytes 识别文档格式（不依赖扩展名）</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>格式名（如 "xlsx"/"pdf"/"md"），无法识别返回 null</returns>
    /// <exception cref="ArgumentNullException">filePath 为空</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    public static String? Detect(String filePath)
    {
        if (String.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

        var fullPath = filePath.GetFullPath();
        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"文件不存在: {fullPath}", fullPath);

        using var fs = File.OpenRead(fullPath);
        return Detect(fs);
    }

    /// <summary>按流内容识别文档格式（不依赖扩展名）</summary>
    /// <param name="stream">输入流</param>
    /// <returns>格式名（如 "xlsx"/"pdf"/"md"），无法识别返回 null</returns>
    /// <exception cref="ArgumentNullException">stream 为空</exception>
    public static String? Detect(Stream stream)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) return null;

        // 复制到可寻址内存流，保证 OLE2/ZIP 容器解析
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        ms.Position = 0;
        return DetectCore(ms);
    }

    /// <summary>按内容识别核心实现</summary>
    private static String? DetectCore(Stream stream)
    {
        var head = new Byte[8];
        var n = stream.Read(head, 0, 8);
        if (n <= 0) return null;

        // PDF: %PDF-
        if (n >= 5 && head[0] == '%' && head[1] == 'P' && head[2] == 'D' && head[3] == 'F' && head[4] == '-')
            return "pdf";

        // RTF: {\rtf
        if (n >= 6 && head[0] == '{' && head[1] == '\\' && head[2] == 'r' && head[3] == 't' && head[4] == 'f')
            return "rtf";

        // OLE2 复合文档（xls/doc/ppt 共用容器）
        if (n >= 8 && head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0 &&
            head[4] == 0xA1 && head[5] == 0xB1 && head[6] == 0x1A && head[7] == 0xE1)
            return DetectOle2(stream);

        // ZIP 容器（xlsx/docx/pptx/epub/ods/xps）
        if (n >= 4 && head[0] == 0x50 && head[1] == 0x4B && (head[2] is 0x03 or 0x05 or 0x07))
            return DetectZip(stream);

        // 文本类（md/vcf/ics/eml）
        return DetectText(stream);
    }

    /// <summary>识别 OLE2 复合文档内部格式（按根存储流名）</summary>
    private static String? DetectOle2(Stream stream)
    {
        try
        {
            stream.Position = 0;
            var reader = new CfbReader(stream);
            var root = reader.Parse();
            if (root.GetStream("Workbook") != null) return "xls";
            if (root.GetStream("WordDocument") != null) return "doc";
            if (root.GetStream("PowerPoint Document") != null) return "ppt";
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>识别 ZIP 容器内格式（按目录结构与内容类型）</summary>
    private static String? DetectZip(Stream stream)
    {
        try
        {
            stream.Position = 0;
            using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            var hasWord = false;
            var hasXl = false;
            var hasPpt = false;
            var hasMimetype = false;
            var hasManifest = false;
            var isXps = false;
            foreach (var entry in zip.Entries)
            {
                var name = entry.FullName;
                if (name == "mimetype") { hasMimetype = true; continue; }
                if (name == "META-INF/manifest.xml") { hasManifest = true; continue; }
                if (name == "[Content_Types].xml")
                {
                    // XPS 固定版式：内容类型含 FixedDocumentSequence
                    using var s = entry.Open();
                    using var sr = new StreamReader(s, Encoding.UTF8);
                    var contentTypes = sr.ReadToEnd();
                    isXps = contentTypes.Contains("fixeddocument-sequence", StringComparison.OrdinalIgnoreCase);
                    continue;
                }
                if (name.StartsWith("word/", StringComparison.Ordinal)) hasWord = true;
                else if (name.StartsWith("xl/", StringComparison.Ordinal)) hasXl = true;
                else if (name.StartsWith("ppt/", StringComparison.Ordinal)) hasPpt = true;

                if (hasPpt || hasWord || hasXl) break;
            }

            if (hasPpt) return "pptx";
            if (hasWord) return "docx";
            if (hasXl) return "xlsx";
            if (hasMimetype) return "epub";
            if (hasManifest) return "ods";
            if (isXps) return "xps";
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>识别文本类格式（md/vcf/ics/eml）</summary>
    private static String? DetectText(Stream stream)
    {
        try
        {
            stream.Position = 0;
            var buf = new Byte[4096];
            var n = stream.Read(buf, 0, buf.Length);
            if (n <= 0) return null;

            // 含 NUL 的二进制但未知格式
            for (var i = 0; i < n; i++)
            {
                if (buf[i] == 0) return null;
            }

            var text = Encoding.UTF8.GetString(buf, 0, n);
            var trimmed = text.TrimStart('\uFEFF', ' ', '\t', '\r', '\n');
            if (trimmed.StartsWith("BEGIN:VCARD", StringComparison.OrdinalIgnoreCase)) return "vcf";
            if (trimmed.StartsWith("BEGIN:VCALENDAR", StringComparison.OrdinalIgnoreCase)) return "ics";
            if (text.StartsWith("MIME-Version:", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("Return-Path:", StringComparison.OrdinalIgnoreCase))
                return "eml";
            return "md";
        }
        catch
        {
            return null;
        }
    }

    /// <summary>按内容识别格式并创建读取器（扩展名错误/缺失时兜底）</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>读取器对象，无法识别返回 null</returns>
    /// <exception cref="ArgumentNullException">filePath 为空</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    public static Object? CreateReaderByContent(String filePath)
    {
        if (String.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

        var fullPath = filePath.GetFullPath();
        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"文件不存在: {fullPath}", fullPath);

        var format = Detect(fullPath);
        if (format == null) return null;
        return CreateReaderByFormat(format, fullPath);
    }

    /// <summary>按内容识别格式并从流创建读取器</summary>
    /// <param name="stream">数据流</param>
    /// <returns>读取器对象，无法识别返回 null</returns>
    public static Object? CreateReaderByContent(Stream stream)
    {
        if (stream == null) return null;

        var format = Detect(stream);
        if (format == null) return null;
        return format switch
        {
            "xlsx" => new ExcelReader(stream, Encoding.UTF8),
            "xls" => new BiffReader(stream),
            "docx" => new WordReader(stream),
            "doc" => new DocReader(stream),
            "pptx" => new PptxReader(stream),
            "ppt" => new PptReader(stream),
            "pdf" => new PdfReader(stream),
            "rtf" => RtfDocument.Parse(stream),
            "ods" => new OdsDocument(OdsReader.Read(stream)),
            "epub" => new EpubReader().Read(stream),
            "vcf" => new VCardDocument(new VCardReader().ReadAll(stream)),
            "eml" => new EmlReader().Read(stream),
            "ics" => new ICalReader().Read(stream),
            "md" => MarkdownDocument.Parse(stream),
            "xps" => new XpsDocument(new XpsReader().Read(stream)),
            _ => null,
        };
    }

    /// <summary>按识别出的格式名从文件创建读取器（内部共用）</summary>
    private static Object? CreateReaderByFormat(String format, String fullPath)
    {
        return format switch
        {
            "xlsx" => new ExcelReader(fullPath),
            "xls" => new BiffReader(fullPath),
            "docx" => new WordReader(fullPath),
            "doc" => new DocReader(fullPath),
            "pptx" => new PptxReader(fullPath),
            "ppt" => new PptReader(fullPath),
            "pdf" => new PdfReader(fullPath),
            "rtf" => RtfDocument.ParseFile(fullPath),
            "ods" => new OdsDocument(OdsReader.ReadFile(fullPath)),
            "epub" => new EpubReader().Read(fullPath),
            "vcf" => new VCardDocument(new VCardReader().ReadAll(fullPath)),
            "eml" => new EmlReader().Read(fullPath),
            "ics" => new ICalReader().Read(fullPath),
            "md" => MarkdownDocument.ParseFile(fullPath),
            "xps" => new XpsDocument(new XpsReader().Read(fullPath)),
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

    #region 格式转换
    /// <summary>将 doc（97-2003）文件转换为 docx 文件（W20）</summary>
    /// <param name="inputPath">doc 输入路径</param>
    /// <param name="outputPath">docx 输出路径</param>
    public static void ConvertDocToDocx(String inputPath, String outputPath)
    {
        using var reader = new DocReader(inputPath);
        reader.SaveAsDocx(outputPath);
    }

    /// <summary>将 doc（97-2003）数据流转换为 docx 数据流（W20）</summary>
    /// <param name="inputStream">doc 输入流</param>
    /// <param name="outputStream">docx 输出流</param>
    public static void ConvertDocToDocx(Stream inputStream, Stream outputStream)
    {
        using var reader = new DocReader(inputStream);
        reader.SaveAsDocx(outputStream);
    }
    #endregion
}
