using System;
using System.IO;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Ods;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.Word;

namespace NewLife.Office.Markdown;

/// <summary>全格式→Markdown 转换中枢（MD07）</summary>
/// <remarks>
/// 统一入口，将 Excel/Word/PPT/PDF 等格式转换为 Markdown。
/// 优先走各 Reader 的 <see cref="IMarkdownDocumentExtractable"/> 结构化 AST（对标 anydoc 统一 Document 模型），
/// 回退到 <see cref="IMarkdownExtractable"/> 字符串实现。
/// </remarks>
public static class FormatToMarkdown
{
    #region Excel → Markdown (MD07-03)
    /// <summary>将 Excel 文件转换为 Markdown GFM 表格</summary>
    /// <param name="filePath">xlsx/xls 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromExcel(String filePath) => FromExcel(filePath, null);

    /// <summary>将 Excel 文件转换为 Markdown GFM 表格</summary>
    /// <param name="filePath">xlsx/xls 文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromExcel(String filePath, MarkdownConverterOptions? options)
    {
        using var reader = new ExcelReader(filePath);
        return ToMarkdownString(reader, filePath, options);
    }

    /// <summary>将 Excel 流转换为 Markdown</summary>
    /// <param name="stream">xlsx/xls 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromExcel(Stream stream) => FromExcel(stream, null);

    /// <summary>将 Excel 流转换为 Markdown</summary>
    /// <param name="stream">xlsx/xls 流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromExcel(Stream stream, MarkdownConverterOptions? options)
    {
        using var reader = new ExcelReader(stream, System.Text.Encoding.UTF8);
        return ToMarkdownString(reader, null, options);
    }
    #endregion

    #region Word → Markdown
    /// <summary>将 Word 文件转换为 Markdown</summary>
    /// <param name="filePath">docx/doc 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromWord(String filePath) => FromWord(filePath, null);

    /// <summary>将 Word 文件转换为 Markdown</summary>
    /// <param name="filePath">docx/doc 文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromWord(String filePath, MarkdownConverterOptions? options)
    {
        using var reader = new WordReader(filePath);
        return ToMarkdownString(reader, filePath, options);
    }

    /// <summary>将 Word 流转换为 Markdown</summary>
    /// <param name="stream">docx/doc 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromWord(Stream stream) => FromWord(stream, null);

    /// <summary>将 Word 流转换为 Markdown</summary>
    /// <param name="stream">docx/doc 流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromWord(Stream stream, MarkdownConverterOptions? options)
    {
        using var reader = new WordReader(stream);
        return ToMarkdownString(reader, null, options);
    }
    #endregion

    #region PPT → Markdown (MD07-02)
    /// <summary>将 PPT 文件转换为 Markdown</summary>
    /// <param name="filePath">pptx/ppt 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPpt(String filePath) => FromPpt(filePath, null);

    /// <summary>将 PPT 文件转换为 Markdown</summary>
    /// <param name="filePath">pptx/ppt 文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPpt(String filePath, MarkdownConverterOptions? options)
    {
        using var reader = new PptxReader(filePath);
        return ToMarkdownString(reader, filePath, options);
    }

    /// <summary>将 PPT 流转换为 Markdown</summary>
    /// <param name="stream">pptx/ppt 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPpt(Stream stream) => FromPpt(stream, null);

    /// <summary>将 PPT 流转换为 Markdown</summary>
    /// <param name="stream">pptx/ppt 流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPpt(Stream stream, MarkdownConverterOptions? options)
    {
        using var reader = new PptxReader(stream);
        return ToMarkdownString(reader, null, options);
    }
    #endregion

    #region PDF → Markdown (MD07-01)
    /// <summary>将 PDF 文件转换为 Markdown</summary>
    /// <param name="filePath">pdf 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPdf(String filePath) => FromPdf(filePath, null);

    /// <summary>将 PDF 文件转换为 Markdown</summary>
    /// <param name="filePath">pdf 文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPdf(String filePath, MarkdownConverterOptions? options)
    {
        using var reader = new PdfReader(filePath);
        return ToMarkdownString(reader, filePath, options);
    }

    /// <summary>将 PDF 流转换为 Markdown</summary>
    /// <param name="stream">pdf 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPdf(Stream stream) => FromPdf(stream, null);

    /// <summary>将 PDF 流转换为 Markdown</summary>
    /// <param name="stream">pdf 流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromPdf(Stream stream, MarkdownConverterOptions? options)
    {
        using var reader = new PdfReader(stream);
        return ToMarkdownString(reader, null, options);
    }
    #endregion

    #region ODS → Markdown (MD09-01)
    /// <summary>将 ODS 文件转换为 Markdown GFM 表格</summary>
    /// <param name="filePath">ods 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromOds(String filePath) => FromOds(filePath, null);

    /// <summary>将 ODS 文件转换为 Markdown GFM 表格</summary>
    /// <param name="filePath">ods 文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromOds(String filePath, MarkdownConverterOptions? options)
    {
        var doc = new OdsDocument(OdsReader.ReadFile(filePath));
        return ToMarkdownString(doc, filePath, options);
    }

    /// <summary>将 ODS 流转换为 Markdown</summary>
    /// <param name="stream">ods 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromOds(Stream stream) => FromOds(stream, null);

    /// <summary>将 ODS 流转换为 Markdown</summary>
    /// <param name="stream">ods 流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromOds(Stream stream, MarkdownConverterOptions? options)
    {
        var doc = new OdsDocument(OdsReader.Read(stream));
        return ToMarkdownString(doc, null, options);
    }
    #endregion

    #region RTF → Markdown (MD09-02)
    /// <summary>将 RTF 文件转换为 Markdown</summary>
    /// <param name="filePath">rtf 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromRtf(String filePath) => FromRtf(filePath, null);

    /// <summary>将 RTF 文件转换为 Markdown</summary>
    /// <param name="filePath">rtf 文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromRtf(String filePath, MarkdownConverterOptions? options)
    {
        var doc = RtfDocument.ParseFile(filePath);
        return ToMarkdownString(doc, filePath, options);
    }

    /// <summary>将 RTF 流转换为 Markdown</summary>
    /// <param name="stream">rtf 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromRtf(Stream stream) => FromRtf(stream, null);

    /// <summary>将 RTF 流转换为 Markdown</summary>
    /// <param name="stream">rtf 流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromRtf(Stream stream, MarkdownConverterOptions? options)
    {
        var doc = RtfDocument.Parse(stream);
        return ToMarkdownString(doc, null, options);
    }
    #endregion

    #region EPUB → Markdown (MD09-03)
    /// <summary>将 EPUB 文件转换为 Markdown（含章节标题结构）</summary>
    /// <param name="filePath">epub 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromEpub(String filePath) => FromEpub(filePath, null);

    /// <summary>将 EPUB 文件转换为 Markdown（含章节标题结构）</summary>
    /// <param name="filePath">epub 文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromEpub(String filePath, MarkdownConverterOptions? options)
    {
        var doc = new EpubReader().Read(filePath);
        return ToMarkdownString(doc, filePath, options);
    }

    /// <summary>将 EPUB 流转换为 Markdown</summary>
    /// <param name="stream">epub 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromEpub(Stream stream) => FromEpub(stream, null);

    /// <summary>将 EPUB 流转换为 Markdown</summary>
    /// <param name="stream">epub 流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromEpub(Stream stream, MarkdownConverterOptions? options)
    {
        var doc = new EpubReader().Read(stream);
        return ToMarkdownString(doc, null, options);
    }
    #endregion

    #region 通用入口
    /// <summary>根据文件扩展名自动选择转换器</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? Convert(String filePath) => Convert(filePath, null);

    /// <summary>根据文件扩展名自动选择转换器</summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? Convert(String filePath, MarkdownConverterOptions? options)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" or ".xls" => FromExcel(filePath, options),
            ".docx" or ".doc" => FromWord(filePath, options),
            ".pptx" or ".ppt" => FromPpt(filePath, options),
            ".pdf" => FromPdf(filePath, options),
            ".ods" => FromOds(filePath, options),
            ".rtf" => FromRtf(filePath, options),
            ".epub" => FromEpub(filePath, options),
            ".md" or ".markdown" => ConvertMarkdownFile(filePath, options),
            // 扩展名未知/不支持时按内容识别兜底（magic bytes，对标 anydoc）
            _ => ConvertByContent(filePath, options),
        };
    }

    /// <summary>根据扩展名从流转换</summary>
    /// <param name="stream">数据流</param>
    /// <param name="extension">文件后缀（可带点号），用于选择转换器</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    public static String? Convert(Stream stream, String extension, MarkdownConverterOptions? options)
    {
        if (stream == null) return null;

        if (!extension.StartsWith(".")) extension = "." + extension;
        return extension.ToLowerInvariant() switch
        {
            ".xlsx" or ".xls" => FromExcel(stream, options),
            ".docx" or ".doc" => FromWord(stream, options),
            ".pptx" or ".ppt" => FromPpt(stream, options),
            ".pdf" => FromPdf(stream, options),
            ".ods" => FromOds(stream, options),
            ".rtf" => FromRtf(stream, options),
            ".epub" => FromEpub(stream, options),
            ".md" or ".markdown" => ConvertMarkdownStream(stream, options),
            // 扩展名未知/不支持时按内容识别兜底（magic bytes，对标 anydoc）
            _ => ConvertByContent(stream, options),
        };
    }

    /// <summary>Markdown 文件直通读取（已是 Markdown，原样返回不重复转换；编码自动检测 BOM/UTF-8/GBK）</summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="options">转换选项（直通模式忽略）</param>
    /// <returns>Markdown 文本</returns>
    private static String? ConvertMarkdownFile(String filePath, MarkdownConverterOptions? options)
        => MarkdownDocument.ReadFileText(filePath);

    /// <summary>Markdown 流直通读取（编码自动检测 BOM/UTF-8/GBK，原样返回不重复转换）</summary>
    /// <param name="stream">数据流</param>
    /// <param name="options">转换选项（直通模式忽略）</param>
    /// <returns>Markdown 文本</returns>
    private static String? ConvertMarkdownStream(Stream stream, MarkdownConverterOptions? options)
        => MarkdownDocument.ReadStreamText(stream);

    /// <summary>扩展名未知/不支持时按内容识别兜底（复用 <see cref="OfficeFactory.CreateReaderByContent(String)"/>）</summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    private static String? ConvertByContent(String filePath, MarkdownConverterOptions? options)
    {
        Object? reader;
        try
        {
            reader = OfficeFactory.CreateReaderByContent(filePath);
        }
        catch (Exception ex)
        {
            // 内容识别失败（文件不存在/无法解析）统一归为 Unsupported
            throw new ConvertErrorException(ConvertErrorType.Unsupported, $"无法识别文件内容或格式不支持: {Path.GetExtension(filePath)}", ex);
        }
        if (reader is not IMarkdownExtractable extractable)
        {
            (reader as IDisposable)?.Dispose();
            throw ConvertErrorException.Unsupported($"无法识别文件内容或格式不支持: {Path.GetExtension(filePath)}");
        }
        try
        {
            return ToMarkdownString(extractable, filePath, options);
        }
        finally
        {
            (reader as IDisposable)?.Dispose();
        }
    }

    /// <summary>扩展名未知/不支持时按内容识别兜底（复用 <see cref="OfficeFactory.CreateReaderByContent(Stream)"/>）</summary>
    /// <param name="stream">数据流</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    private static String? ConvertByContent(Stream stream, MarkdownConverterOptions? options)
    {
        Object? reader;
        try
        {
            reader = OfficeFactory.CreateReaderByContent(stream);
        }
        catch (Exception ex)
        {
            // 内容识别失败统一归为 Unsupported
            throw new ConvertErrorException(ConvertErrorType.Unsupported, "无法识别流内容或格式不支持", ex);
        }
        if (reader is not IMarkdownExtractable extractable)
        {
            (reader as IDisposable)?.Dispose();
            throw ConvertErrorException.Unsupported("无法识别流内容或格式不支持");
        }
        try
        {
            return ToMarkdownString(extractable, null, options);
        }
        finally
        {
            (reader as IDisposable)?.Dispose();
        }
    }
    #endregion

    #region 文档级转换
    /// <summary>将文件转换为 MarkdownDocument（结构化 AST）</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>Markdown 文档对象，失败返回 null</returns>
    public static MarkdownDocument? ToDocument(String filePath) => ToDocument(filePath, null);

    /// <summary>将文件转换为 MarkdownDocument（结构化 AST）</summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 文档对象，失败返回 null</returns>
    public static MarkdownDocument? ToDocument(String filePath, MarkdownConverterOptions? options)
    {
        var reader = OfficeFactory.CreateReader(filePath);
        if (reader == null) return null;
        try
        {
            var doc = (reader as IMarkdownDocumentExtractable)?.ToMarkdownDocument();
            if (doc != null) return ApplyOptions(doc, filePath, options);

            // 回退：旧字符串实现 → 解析
            var md = (reader as IMarkdownExtractable)?.ExtractMarkdown();
            if (md == null) return null;
            doc = MarkdownDocument.Parse(md);
            return ApplyOptions(doc, filePath, options);
        }
        finally
        {
            (reader as IDisposable)?.Dispose();
        }
    }
    #endregion

    #region 辅助
    /// <summary>从读取器提取 Markdown 字符串（优先 AST，回退字符串实现）</summary>
    /// <param name="reader">支持 Markdown 提取的读取器</param>
    /// <param name="filePath">源文件路径（可选，用于元数据）</param>
    /// <param name="options">转换选项，null 使用默认</param>
    /// <returns>Markdown 字符串</returns>
    private static String? ToMarkdownString(IMarkdownExtractable reader, String? filePath, MarkdownConverterOptions? options)
    {
        var doc = (reader as IMarkdownDocumentExtractable)?.ToMarkdownDocument();
        if (doc != null)
        {
            doc = ApplyOptions(doc, filePath, options);
            return doc.ToMarkdown();
        }
        return reader.ExtractMarkdown();
    }

    /// <summary>应用转换选项：元数据 FrontMatter + 噪音过滤</summary>
    /// <param name="doc">Markdown 文档</param>
    /// <param name="filePath">源文件路径（可选）</param>
    /// <param name="options">转换选项</param>
    /// <returns>处理后的文档</returns>
    private static MarkdownDocument ApplyOptions(MarkdownDocument doc, String? filePath, MarkdownConverterOptions? options)
    {
        options ??= new MarkdownConverterOptions();

        // 元数据 FrontMatter
        if (options.Metadata && !String.IsNullOrEmpty(filePath))
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            if (!String.IsNullOrEmpty(fileName) && !doc.FrontMatter.ContainsKey("title"))
                doc.FrontMatter["title"] = fileName!;
            if (!doc.FrontMatter.ContainsKey("source"))
                doc.FrontMatter["source"] = Path.GetFileName(filePath) ?? filePath!;
            var ext = Path.GetExtension(filePath)?.TrimStart('.');
            if (!String.IsNullOrEmpty(ext) && !doc.FrontMatter.ContainsKey("format"))
                doc.FrontMatter["format"] = ext!;
        }

        // 噪音过滤：空段落
        if (options.CleanNoise)
        {
            for (var i = doc.Blocks.Count - 1; i >= 0; i--)
            {
                var block = doc.Blocks[i];
                if (block.Type == MarkdownBlockType.Paragraph && String.IsNullOrWhiteSpace(block.GetPlainText()))
                    doc.Blocks.RemoveAt(i);
            }
        }

        return doc;
    }
    #endregion
}
