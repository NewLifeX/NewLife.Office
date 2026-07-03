using System;
using System.IO;
using NewLife.Office.Excel;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Word;

namespace NewLife.Office.Markdown;

/// <summary>全格式→Markdown 转换中枢（MD07）</summary>
/// <remarks>
/// 统一入口，将 Excel/Word/PPT/PDF 等格式转换为 Markdown。
/// 底层复用各 Reader 已有的 <see cref="IMarkdownExtractable"/> 实现。
/// </remarks>
public static class FormatToMarkdown
{
    #region Excel → Markdown (MD07-03)
    /// <summary>将 Excel 文件转换为 Markdown GFM 表格</summary>
    /// <param name="filePath">xlsx/xls 文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromExcel(String filePath)
    {
        using var reader = new ExcelReader(filePath);
        return reader.ExtractMarkdown();
    }

    /// <summary>将 Excel 流转换为 Markdown</summary>
    /// <param name="stream">xlsx/xls 流</param>
    /// <returns>Markdown 字符串</returns>
    public static String? FromExcel(Stream stream)
    {
        using var reader = new ExcelReader(stream, System.Text.Encoding.UTF8);
        return reader.ExtractMarkdown();
    }
    #endregion

    #region Word → Markdown
    /// <summary>将 Word 文件转换为 Markdown</summary>
    public static String? FromWord(String filePath)
    {
        using var reader = new WordReader(filePath);
        return reader.ExtractMarkdown();
    }

    /// <summary>将 Word 流转换为 Markdown</summary>
    public static String? FromWord(Stream stream)
    {
        using var reader = new WordReader(stream);
        return reader.ExtractMarkdown();
    }
    #endregion

    #region PPT → Markdown (MD07-02)
    /// <summary>将 PPT 文件转换为 Markdown</summary>
    public static String? FromPpt(String filePath)
    {
        using var reader = new PptxReader(filePath);
        return reader.ExtractMarkdown();
    }

    /// <summary>将 PPT 流转换为 Markdown</summary>
    public static String? FromPpt(Stream stream)
    {
        using var reader = new PptxReader(stream);
        return reader.ExtractMarkdown();
    }
    #endregion

    #region PDF → Markdown (MD07-01)
    /// <summary>将 PDF 文件转换为 Markdown</summary>
    public static String? FromPdf(String filePath)
    {
        using var reader = new PdfReader(filePath);
        return ((IMarkdownExtractable)reader).ExtractMarkdown();
    }

    /// <summary>将 PDF 流转换为 Markdown</summary>
    public static String? FromPdf(Stream stream)
    {
        using var reader = new PdfReader(stream);
        return ((IMarkdownExtractable)reader).ExtractMarkdown();
    }
    #endregion

    #region 通用入口
    /// <summary>根据文件扩展名自动选择转换器</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>Markdown 字符串</returns>
    public static String? Convert(String filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" or ".xls" => FromExcel(filePath),
            ".docx" => FromWord(filePath),
            ".pptx" => FromPpt(filePath),
            ".pdf" => FromPdf(filePath),
            _ => throw new NotSupportedException($"不支持的文件格式: {ext}"),
        };
    }
    #endregion

    #region 文档级转换
    /// <summary>将文件转换为 MarkdownDocument</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>Markdown 文档对象，失败返回 null</returns>
    public static MarkdownDocument? ToDocument(String filePath)
    {
        var md = Convert(filePath);
        return md != null ? MarkdownDocument.Parse(md) : null;
    }
    #endregion
}
