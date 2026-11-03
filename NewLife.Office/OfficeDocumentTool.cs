using System.Text;
using System.Text.RegularExpressions;
using NewLife.Office.Excel;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Word;

namespace NewLife.Office;

/// <summary>AI 工具错误码</summary>
public static class OfficeToolErrorCode
{
    /// <summary>不支持的格式</summary>
    public const String Unsupported = "Unsupported";

    /// <summary>文件结构非法或解析失败</summary>
    public const String Malformed = "Malformed";

    /// <summary>文档已加密</summary>
    public const String Encrypted = "Encrypted";

    /// <summary>锚点/页码不存在</summary>
    public const String AnchorMissing = "AnchorMissing";

    /// <summary>超出资源上限</summary>
    public const String ResourceLimit = "ResourceLimit";
}

/// <summary>统一工具结果：成功内容 + 错误码 + 分页信息</summary>
public class OfficeToolResult
{
    /// <summary>是否成功</summary>
    public Boolean Success { get; set; }

    /// <summary>格式（docx/xlsx/pptx/pdf）</summary>
    public String? Format { get; set; }

    /// <summary>内容（JSON 或纯文本）</summary>
    public String? Content { get; set; }

    /// <summary>错误码（见 <see cref="OfficeToolErrorCode"/>）</summary>
    public String? Code { get; set; }

    /// <summary>错误信息</summary>
    public String? Message { get; set; }

    /// <summary>内容是否被截断</summary>
    public Boolean Truncated { get; set; }

    /// <summary>续读游标（字符偏移或行偏移）</summary>
    public Int32 NextCursor { get; set; }

    /// <summary>告警</summary>
    public List<String> Warnings { get; } = [];

    /// <summary>输出 JSON 结果</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"success\": ").Append(Success ? "true" : "false");
        if (!Format.IsNullOrEmpty()) sb.Append(", \"format\": ").Append(OfficeJson.Escape(Format));
        if (!Code.IsNullOrEmpty()) sb.Append(", \"code\": ").Append(OfficeJson.Escape(Code));
        if (!Message.IsNullOrEmpty()) sb.Append(", \"message\": ").Append(OfficeJson.Escape(Message));
        if (Truncated) sb.Append(", \"truncated\": true, \"nextCursor\": ").Append(NextCursor);
        if (Warnings.Count > 0) sb.Append(", \"warnings\": ").Append(OfficeJson.Array(Warnings));
        if (Content != null) sb.Append(",").Append(nl).Append(pad).Append("\"content\": ").Append(OfficeJson.Escape(Content));
        sb.Append(nl).Append('}');
        return sb.ToString();
    }
}

/// <summary>跨格式 AI 工具层：按内容/扩展名路由到 docx/xlsx/pptx/pdf 的快照、读取、搜索与补丁</summary>
/// <remarks>
/// 所有写入都"整批成功才落盘"，失败返回 <see cref="OfficeToolErrorCode"/> 错误码。
/// <example>
/// <code>
/// var outline = OfficeDocumentTool.ReadOutline("report.xlsx");
/// Console.WriteLine(outline.Content);
///
/// var patch = OfficeDocumentTool.ApplyPatch("report.xlsx", json, "out.xlsx");
/// if (!patch.Success) Console.WriteLine($"{patch.Code}: {patch.Message}");
/// </code>
/// </example>
/// </remarks>
public static class OfficeDocumentTool
{
    #region 格式识别
    /// <summary>识别文档格式</summary>
    /// <param name="path">文件路径</param>
    /// <returns>docx/xlsx/pptx/pdf，无法识别返回未知后缀</returns>
    public static String Detect(String path)
    {
        if (path == null) throw new ArgumentNullException(nameof(path));

        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        if (ext is "docx" or "xlsx" or "pptx" or "pdf") return ext;

        return OfficeFactory.Detect(path) ?? ext;
    }

    private static Boolean IsEncryptedOoxml(String path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            var header = new Byte[8];
            var read = fs.Read(header, 0, header.Length);
            return read >= 8 && header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0;
        }
        catch
        {
            return false;
        }
    }

    private static Boolean IsEncryptedPdf(String path)
    {
        try
        {
            using var reader = new PdfReader(path);
            return reader.XRefTable?.Trailer.ContainsKey("Encrypt") == true;
        }
        catch
        {
            return false;
        }
    }
    #endregion

    #region 读取
    /// <summary>读取结构快照</summary>
    /// <param name="path">文件路径</param>
    /// <returns>工具结果（Content 为 JSON）</returns>
    public static OfficeToolResult ReadOutline(String path)
    {
        try
        {
            var format = Detect(path);
            if (format is not ("docx" or "xlsx" or "pptx" or "pdf"))
                return Fail(format, OfficeToolErrorCode.Unsupported, L.T("Unsupported structure snapshot format: {0}", format));
            if (format != "pdf" && IsEncryptedOoxml(path))
                return Fail(format, OfficeToolErrorCode.Encrypted, L.T("The document is encrypted; decrypt it first."));

            String json;
            switch (format)
            {
                case "docx":
                    using (var reader = new WordReader(path)) json = reader.ReadOutline().ToJson();
                    break;
                case "xlsx":
                case "xls":
                    using (var reader = new ExcelReader(path)) json = reader.ReadExcel().ToOutline().ToJson();
                    break;
                case "pptx":
                    using (var reader = new PptxReader(path)) json = reader.ReadDocument().ToOutline().ToJson();
                    break;
                default:
                    using (var reader = new PdfReader(path))
                    {
                        if (reader.XRefTable?.Trailer.ContainsKey("Encrypt") == true)
                            return Fail(format, OfficeToolErrorCode.Encrypted, L.T("The PDF is encrypted; decrypt it first."));
                        json = PdfSnapshot.From(reader).ToJson();
                    }
                    break;
            }

            return new OfficeToolResult { Success = true, Format = format, Content = json };
        }
        catch (Exception ex)
        {
            return Fail(null, MapErrorCode(ex), ex.Message);
        }
    }

    /// <summary>分段读取纯文本（任一种受支持格式）</summary>
    /// <param name="path">文件路径</param>
    /// <param name="maxChars">本次最多返回字符数（0 表示不限）</param>
    /// <param name="cursor">续读游标（上次返回的 NextCursor）</param>
    /// <returns>工具结果（Content 为纯文本片段）</returns>
    public static OfficeToolResult ReadText(String path, Int32 maxChars = 0, Int32 cursor = 0)
    {
        try
        {
            var format = Detect(path);
            if (format != "pdf" && IsEncryptedOoxml(path))
                return Fail(format, OfficeToolErrorCode.Encrypted, L.T("The document is encrypted; decrypt it first."));
            if (format == "pdf" && IsEncryptedPdf(path))
                return Fail(format, OfficeToolErrorCode.Encrypted, L.T("The PDF is encrypted; decrypt it first."));

            var text = OfficeFactory.ReadText(path);
            if (text == null) return Fail(format, OfficeToolErrorCode.Unsupported, L.T("Cannot read text from format: {0}", format));

            if (cursor < 0) cursor = 0;
            if (cursor > text.Length) cursor = text.Length;

            var length = text.Length - cursor;
            if (maxChars > 0 && length > maxChars) length = maxChars;

            var content = text.Substring(cursor, length);
            var truncated = cursor + length < text.Length;
            return new OfficeToolResult
            {
                Success = true,
                Format = format,
                Content = content,
                Truncated = truncated,
                NextCursor = cursor + length,
            };
        }
        catch (Exception ex)
        {
            return Fail(null, MapErrorCode(ex), ex.Message);
        }
    }

    /// <summary>分页读取工作簿区域</summary>
    /// <param name="path">工作簿路径</param>
    /// <param name="range">A1 区域（如 <c>Sheet1!A1:D100</c>）</param>
    /// <param name="maxRows">本次最多返回行数（0 表示不限）</param>
    /// <param name="cursor">续读游标（已跳过的行数）</param>
    /// <returns>工具结果（Content 为 JSON）</returns>
    public static OfficeToolResult ReadRange(String path, String range, Int32 maxRows = 0, Int32 cursor = 0)
    {
        try
        {
            var format = Detect(path);
            if (format is not ("xlsx" or "xls"))
                return Fail(format, OfficeToolErrorCode.Unsupported, L.T("Range reading is only supported for workbooks: {0}", format));

            using var reader = new ExcelReader(path);
            var document = reader.ReadExcel();
            var cellRange = ExcelCellRange.Parse(range, document.Worksheets.FirstOrDefault()?.Name)
                ?? throw new FormatException(L.T("Cannot parse range: {0}", range));
            var sheet = document.GetSheet(cellRange.Sheet);

            cellRange = cellRange.Clamp(ExcelSheetEditor.GetRowCount(sheet), ExcelSheetEditor.GetColumnCount(sheet));
            if (cursor < 0) cursor = 0;
            if (cursor >= cellRange.RowCount) cursor = Math.Max(0, cellRange.RowCount);

            var available = cellRange.RowCount - cursor;
            var take = maxRows > 0 ? Math.Min(maxRows, available) : available;
            var sub = new ExcelCellRange(sheet.Name, cellRange.StartRow + cursor, cellRange.StartCol,
                cellRange.StartRow + cursor + Math.Max(0, take - 1), cellRange.EndCol);
            var rows = take > 0 ? ExcelSheetEditor.ReadRange(sheet, sub) : new List<Object?[]>();

            var json = new StringBuilder();
            json.Append('{');
            json.Append("\"range\": ").Append(OfficeJson.Escape(sub.FullAddress));
            json.Append(", \"rows\": ").Append(OfficeJson.Matrix(rows));
            var truncated = cursor + take < cellRange.RowCount;
            if (truncated) json.Append(", \"truncated\": true, \"nextCursor\": ").Append(cursor + take);
            json.Append('}');

            return new OfficeToolResult
            {
                Success = true,
                Format = format,
                Content = json.ToString(),
                Truncated = truncated,
                NextCursor = cursor + take,
            };
        }
        catch (Exception ex)
        {
            return Fail(null, MapErrorCode(ex), ex.Message);
        }
    }

    /// <summary>在工作簿或 PDF 中搜索文本</summary>
    /// <param name="path">文件路径</param>
    /// <param name="text">搜索内容</param>
    /// <param name="regex">是否正则</param>
    /// <param name="maxResults">最大命中数</param>
    /// <returns>工具结果（Content 为 JSON）</returns>
    public static OfficeToolResult Search(String path, String text, Boolean regex = false, Int32 maxResults = 100)
    {
        try
        {
            var format = Detect(path);
            switch (format)
            {
                case "xlsx":
                case "xls":
                {
                    using var reader = new ExcelReader(path);
                    var result = ExcelSearch.Search(reader.ReadExcel(), text, regex, false, maxResults);
                    return new OfficeToolResult { Success = true, Format = format, Content = result.ToJson() };
                }
                case "pdf":
                {
                    using var reader = new PdfReader(path);
                    if (reader.XRefTable?.Trailer.ContainsKey("Encrypt") == true)
                        return Fail(format, OfficeToolErrorCode.Encrypted, L.T("The PDF is encrypted; decrypt it first."));
                    var hits = reader.SearchText(text, regex, false, maxResults);
                    var result = new PdfTextSearchResult { Truncated = hits.Count >= maxResults };
                    result.Hits.AddRange(hits);
                    return new OfficeToolResult { Success = true, Format = format, Content = result.ToJson() };
                }
                default:
                    return Fail(format, OfficeToolErrorCode.Unsupported, L.T("Search is not supported for format: {0}", format));
            }
        }
        catch (RegexMatchTimeoutException)
        {
            return Fail(null, OfficeToolErrorCode.ResourceLimit, L.T("Regular expression evaluation timed out."));
        }
        catch (Exception ex)
        {
            return Fail(null, MapErrorCode(ex), ex.Message);
        }
    }
    #endregion

    #region 补丁
    /// <summary>应用 JSON 补丁；仅在全部成功时写出文件</summary>
    /// <param name="path">源文件路径</param>
    /// <param name="patchJson">补丁 JSON</param>
    /// <param name="outputPath">输出路径</param>
    /// <returns>工具结果（Content 为补丁执行结果 JSON）</returns>
    public static OfficeToolResult ApplyPatch(String path, String patchJson, String outputPath)
    {
        try
        {
            if (patchJson == null) throw new ArgumentNullException(nameof(patchJson));
            if (outputPath == null) throw new ArgumentNullException(nameof(outputPath));

            var format = Detect(path);
            if (format is not ("docx" or "xlsx" or "xls" or "pptx" or "pdf"))
                return Fail(format, OfficeToolErrorCode.Unsupported, L.T("Unsupported patch format: {0}", format));
            if (format == "xls" && outputPath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                return Fail(format, OfficeToolErrorCode.Unsupported, L.T("xls does not support model-level write-back; change the output path to .xlsx."));
            if (format != "pdf" && IsEncryptedOoxml(path))
                return Fail(format, OfficeToolErrorCode.Encrypted, L.T("The document is encrypted; decrypt it first."));

            String resultJson;
            var warnings = new List<String>();
            switch (format)
            {
                case "docx":
                {
                    using var reader = new WordReader(path);
                    var doc = reader.ReadDocument();
                    doc.BeginEdit();
                    var patch = DocxPatch.ParseJson(patchJson);
                    var result = DocxPatch.Apply(doc, patch);
                    if (!result.Success)
                        return Fail(format, ClassifyPatchFailure(result.FailedCodes.ToArray()), String.Join("; ", result.Failed));

                    warnings.AddRange(result.Warnings);
                    resultJson = new OfficeResultBridge(result.Applied.Count, result.Failed, result.Warnings).ToJson();
                    WriteAtomic(outputPath, temp =>
                    {
                        using var writer = new WordWriter();
                        writer.Save(temp, doc);
                    });
                    break;
                }
                case "xlsx":
                case "xls":
                {
                    using var reader = new ExcelReader(path);
                    var doc = reader.ReadExcel();
                    var patch = ExcelPatch.ParseJson(patchJson);
                    var result = patch.Apply(doc);
                    if (!result.Success)
                        return Fail(format, ClassifyPatchFailure(result.FailedCodes.ToArray()), String.Join("; ", result.Failed));

                    warnings.AddRange(result.Warnings);
                    resultJson = result.ToJson();
                    WriteAtomic(outputPath, temp => doc.Save(temp));
                    break;
                }
                case "pptx":
                {
                    using var reader = new PptxReader(path);
                    var doc = reader.ReadDocument();
                    var patch = PptxPatch.ParseJson(patchJson);
                    var result = patch.Apply(doc);
                    if (!result.Success)
                        return Fail(format, ClassifyPatchFailure(result.FailedCodes.ToArray()), String.Join("; ", result.Failed));

                    warnings.AddRange(result.Warnings);
                    resultJson = result.ToJson();
                    WriteAtomic(outputPath, temp => doc.Save(temp, path));
                    break;
                }
                default:
                {
                    var doc = PdfObjectDocument.Load(path);
                    if (doc.IsEncrypted) return Fail(format, OfficeToolErrorCode.Encrypted, L.T("The PDF is encrypted; decrypt it first."));
                    var patch = PdfPatch.ParseJson(patchJson);
                    var result = patch.Apply(doc);
                    if (!result.Success)
                        return Fail(format, ClassifyPatchFailure(result.FailedCodes.ToArray()), String.Join("; ", result.Failed));

                    warnings.AddRange(result.Warnings);
                    resultJson = result.ToJson();
                    WriteAtomic(outputPath, temp => doc.Save(temp));
                    break;
                }
            }

            var toolResult = new OfficeToolResult { Success = true, Format = format, Content = resultJson };
            toolResult.Warnings.AddRange(warnings);
            return toolResult;
        }
        catch (Exception ex)
        {
            return Fail(null, MapErrorCode(ex), ex.Message);
        }
    }
    #endregion

    #region 辅助
    private static OfficeToolResult Fail(String? format, String code, String message) => new()
    {
        Success = false,
        Format = format,
        Code = code,
        Message = message,
    };

    private static String MapErrorCode(Exception ex) => ex switch
    {
        NotSupportedException => OfficeToolErrorCode.Unsupported,
        TimeoutException => OfficeToolErrorCode.ResourceLimit,
        ArgumentOutOfRangeException => OfficeToolErrorCode.AnchorMissing,
        FileNotFoundException => OfficeToolErrorCode.Malformed,
        InvalidDataException => OfficeToolErrorCode.Malformed,
        FormatException => OfficeToolErrorCode.Malformed,
        _ => OfficeToolErrorCode.Malformed,
    };

    private static String ClassifyPatchFailure(String[] codes)
    {
        foreach (var code in codes)
        {
            if (code == nameof(NotSupportedException)) return OfficeToolErrorCode.Unsupported;
        }
        foreach (var code in codes)
        {
            if (code == nameof(ArgumentException) || code == nameof(ArgumentOutOfRangeException) || code == nameof(KeyNotFoundException))
                return OfficeToolErrorCode.AnchorMissing;
        }
        return OfficeToolErrorCode.Malformed;
    }

    private static void WriteAtomic(String outputPath, Action<String> write)
    {
        var file = outputPath.GetFullPath();
        file.EnsureDirectory(true);
        var temp = file + ".tmp";
        try
        {
            write(temp);
            if (File.Exists(file)) File.Delete(file);
            File.Move(temp, file);
        }
        catch
        {
            if (File.Exists(temp)) File.Delete(temp);
            throw;
        }
    }
    #endregion
}

/// <summary>补丁结果适配（统一 applied/failed/warnings 结构）</summary>
internal class OfficeResultBridge
{
    private readonly Int32 _applied;
    private readonly List<String> _failed;
    private readonly List<String> _warnings;

    public OfficeResultBridge(Int32 applied, List<String> failed, List<String> warnings)
    {
        _applied = applied;
        _failed = failed;
        _warnings = warnings;
    }

    public String ToJson()
    {
        var sb = new StringBuilder();
        sb.Append("{\"success\": ").Append(_failed.Count == 0 ? "true" : "false");
        sb.Append(", \"applied\": ").Append(_applied);
        sb.Append(", \"failed\": ").Append(OfficeJson.Array(_failed));
        sb.Append(", \"warnings\": ").Append(OfficeJson.Array(_warnings));
        sb.Append('}');
        return sb.ToString();
    }
}
