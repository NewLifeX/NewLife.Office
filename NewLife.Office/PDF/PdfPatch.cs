using System.Globalization;
using System.Text;

namespace NewLife.Office.Pdf;

/// <summary>PDF JSON 补丁：页面操作、元数据、标注与表单填充，任一条失败不写文件</summary>
/// <remarks>
/// 页码使用 1 起始（与阅读器一致），坐标为 PDF 用户空间（原点左下角，单位点）。
/// 执行只改内存对象，调用方应在 <see cref="PdfPatchResult.Success"/> 为 true 时再保存。
/// <example>
/// <code>
/// var patch = PdfPatch.ParseJson("""
/// { "ops": [
///   { "op": "set_metadata", "title": "2026 年度报告" },
///   { "op": "add_highlight", "page": 1, "x": 72, "y": 700, "width": 120, "height": 14, "text": "重点" },
///   { "op": "rotate_page", "pages": "2", "rotation": 90 }
/// ] }
/// """);
/// var doc = PdfObjectDocument.Load("in.pdf");
/// if (patch.Apply(doc).Success) doc.Save("out.pdf");
/// </code>
/// </example>
/// </remarks>
public class PdfPatch
{
    #region 属性
    /// <summary>指令集合</summary>
    public List<PdfPatchOp> Ops { get; } = [];
    #endregion

    #region 解析
    /// <summary>解析 JSON 补丁</summary>
    /// <param name="json">JSON 文本</param>
    /// <returns>补丁对象</returns>
    public static PdfPatch ParseJson(String json)
    {
        if (json == null) throw new ArgumentNullException(nameof(json));

        var root = JsonLite.Parse(json) as Dictionary<String, Object?>
            ?? throw new FormatException(L.T("Patch root must be a JSON object."));
        if (!root.TryGetValue("ops", out var opsValue) || opsValue is not List<Object?> list)
            throw new FormatException(L.T("Patch must contain an 'ops' array."));

        var patch = new PdfPatch();
        foreach (var item in list)
        {
            if (item is not Dictionary<String, Object?> obj)
                throw new FormatException(L.T("Each op must be a JSON object."));
            patch.Ops.Add(new PdfPatchOp
            {
                Op = JsonLite.GetString(obj, "op") ?? throw new FormatException(L.T("Each op must have an 'op' name.")),
                Args = obj,
            });
        }
        return patch;
    }
    #endregion

    #region 执行
    /// <summary>在对象级文档上执行补丁</summary>
    /// <param name="document">对象级文档</param>
    /// <returns>执行结果</returns>
    public PdfPatchResult Apply(PdfObjectDocument document)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.IsEncrypted) throw new NotSupportedException(L.T("Patches are not supported for encrypted PDFs; decrypt the file first."));

        var result = new PdfPatchResult();
        for (var i = 0; i < Ops.Count; i++)
        {
            var op = Ops[i];
            try
            {
                ApplyOp(document, op, result);
                result.Applied.Add($"{i}:{op.Op}");
            }
            catch (Exception ex)
            {
                result.Failed.Add($"{i}:{op.Op}: {ex.Message}");
            }
        }
        return result;
    }

    private static void ApplyOp(PdfObjectDocument document, PdfPatchOp op, PdfPatchResult result)
    {
        switch (op.Op)
        {
            case "delete_pages":
            {
                var remove = OpPages(op);
                var pages = document.GetPageObjectNumbers();
                foreach (var page in remove)
                    if (page > pages.Count) throw new ArgumentOutOfRangeException(nameof(op), L.T("Page number out of range: {0}", page));

                var keep = new List<Int32>();
                for (var i = 0; i < pages.Count; i++)
                    if (!remove.Contains(i + 1)) keep.Add(i);
                if (keep.Count == 0) throw new ArgumentException(L.T("Cannot delete all pages."));
                document.KeepPages(keep);
                return;
            }

            case "extract_pages":
            case "keep_pages":
            {
                var keep = OpPages(op).Select(e => e - 1).OrderBy(e => e).ToList();
                var count = document.GetPageObjectNumbers().Count;
                foreach (var index in keep)
                    if (index < 0 || index >= count) throw new ArgumentOutOfRangeException(nameof(op), L.T("Page number out of range: {0}", index + 1));
                document.KeepPages(keep);
                return;
            }

            case "move_page":
            {
                var from = op.GetInt32("from") ?? throw new ArgumentException(L.T("move_page requires 'from'."));
                var to = op.GetInt32("to") ?? throw new ArgumentException(L.T("move_page requires 'to'."));
                var pages = document.GetPageObjectNumbers();
                if (from < 1 || from > pages.Count) throw new ArgumentOutOfRangeException(nameof(op), L.T("Page number out of range: {0}", from));
                if (to < 1 || to > pages.Count) throw new ArgumentOutOfRangeException(nameof(op), L.T("Page number out of range: {0}", to));

                var page = pages[from - 1];
                pages.RemoveAt(from - 1);
                pages.Insert(to - 1, page);
                document.SetPageOrder(pages);
                return;
            }

            case "rotate_page":
            {
                var rotation = op.GetInt32("rotation") ?? throw new ArgumentException(L.T("rotate_page requires 'rotation'."));
                foreach (var page in OpPages(op))
                    document.RotatePage(page - 1, rotation);
                return;
            }

            case "set_metadata":
            {
                var values = new Dictionary<String, String?>();
                Add(values, "Title", op.GetString("title"));
                Add(values, "Author", op.GetString("author"));
                Add(values, "Subject", op.GetString("subject"));
                Add(values, "Keywords", op.GetString("keywords"));
                Add(values, "Creator", op.GetString("creator"));
                Add(values, "Producer", op.GetString("producer"));
                if (values.Count == 0) throw new ArgumentException(L.T("set_metadata requires at least one field."));
                document.SetMetadata(values);
                return;
            }

            case "add_text":
            {
                var page = RequirePage(op);
                document.AddText(page - 1,
                    op.GetString("text") ?? throw new ArgumentException(L.T("add_text requires 'text'.")),
                    op.GetSingle("x") ?? 72f,
                    op.GetSingle("y") ?? 72f,
                    op.GetSingle("fontSize") ?? 12f,
                    op.GetString("color"),
                    op.GetSingle("rotation") ?? 0f);
                return;
            }

            case "add_watermark":
            {
                document.AddWatermark(
                    op.GetString("text") ?? throw new ArgumentException(L.T("add_watermark requires 'text'.")),
                    op.GetSingle("fontSize") ?? 48f,
                    op.GetString("color"));
                return;
            }

            case "add_highlight":
            {
                var page = RequirePage(op);
                var x = op.GetSingle("x") ?? throw new ArgumentException(L.T("add_highlight requires '{0}'.", "x"));
                var y = op.GetSingle("y") ?? throw new ArgumentException(L.T("add_highlight requires '{0}'.", "y"));
                var width = op.GetSingle("width") ?? throw new ArgumentException(L.T("add_highlight requires '{0}'.", "width"));
                var height = op.GetSingle("height") ?? throw new ArgumentException(L.T("add_highlight requires '{0}'.", "height"));

                var annotation = BuildHighlight(x, y, width, height, op.GetString("color") ?? "FFEB3B",
                    op.GetString("text"), op.GetString("author"));
                document.AddAnnotation(page - 1, annotation);
                return;
            }

            case "add_note":
            {
                var page = RequirePage(op);
                var x = op.GetSingle("x") ?? 72f;
                var y = op.GetSingle("y") ?? 72f;
                var annotation = BuildNote(x, y, op.GetString("text") ?? String.Empty, op.GetString("author"));
                document.AddAnnotation(page - 1, annotation);
                return;
            }

            case "fill_form":
            {
                if (op.Args.TryGetValue("fields", out var fieldsValue) && fieldsValue is Dictionary<String, Object?> fields)
                {
                    var applied = false;
                    foreach (var kv in fields)
                        if (document.FillForm(kv.Key, Convert.ToString(kv.Value, CultureInfo.InvariantCulture) ?? String.Empty)) applied = true;
                    if (!applied) throw new ArgumentException(L.T("No matching form field."));
                    return;
                }

                var name = op.GetString("name") ?? throw new ArgumentException(L.T("fill_form requires 'name' or 'fields'."));
                var value = op.GetString("value") ?? String.Empty;
                if (!document.FillForm(name, value)) throw new ArgumentException(L.T("Form field not found: {0}", name));
                return;
            }

            case "merge":
            {
                var path = op.GetString("path") ?? throw new ArgumentException(L.T("merge requires 'path'."));
                if (!File.Exists(path)) throw new FileNotFoundException(L.T("File not found: {0}", path));
                var other = PdfObjectDocument.Load(path);
                document.Merge(other);
                return;
            }

            default:
                throw new ArgumentException(L.T("Unknown operation: {0}", op.Op));
        }
    }

    private static void Add(Dictionary<String, String?> values, String key, String? value)
    {
        if (value != null) values[key] = value;
    }

    private static Int32 RequirePage(PdfPatchOp op)
    {
        var page = op.GetInt32("page") ?? throw new ArgumentException(L.T("{0} requires a 1-based 'page'.", op.Op));
        if (page < 1) throw new ArgumentOutOfRangeException(nameof(op), L.T("Page numbers start at 1: {0}", page));
        return page;
    }

    /// <summary>解析 pages 参数：支持 "1-3,5" 或 [1,2,3]；缺省为全部页</summary>
    private static List<Int32> OpPages(PdfPatchOp op)
    {
        var pages = new List<Int32>();
        if (op.Args.TryGetValue("pages", out var value))
        {
            if (value is List<Object?> list)
            {
                foreach (var item in list)
                    pages.Add(Convert.ToInt32(item, CultureInfo.InvariantCulture));
            }
            else if (value is String text && text.Length > 0)
            {
                foreach (var part in text.Split(','))
                {
                    var token = part.Trim();
                    if (token.Length == 0) continue;
                    var dash = token.IndexOf('-');
                    if (dash > 0)
                    {
                        var start = Convert.ToInt32(token[..dash], CultureInfo.InvariantCulture);
                        var end = Convert.ToInt32(token[(dash + 1)..], CultureInfo.InvariantCulture);
                        for (var i = start; i <= end; i++) pages.Add(i);
                    }
                    else
                        pages.Add(Convert.ToInt32(token, CultureInfo.InvariantCulture));
                }
            }
        }

        if (pages.Count == 0) throw new ArgumentException(L.T("{0} requires 'pages' (for example \"1-3,5\" or [1,2,3]).", op.Op));
        if (pages.Any(e => e < 1)) throw new ArgumentOutOfRangeException(nameof(op), L.T("Page numbers start at 1."));
        return pages;
    }

    private static PdfObject BuildHighlight(Single x, Single y, Single width, Single height, String colorHex, String? text, String? author)
    {
        var dict = new PdfDict
        {
            ["Type"] = new PdfName("Annot"),
            ["Subtype"] = new PdfName("Highlight"),
            ["Rect"] = new PdfArray
            {
                Items =
                {
                    new PdfNumber(x), new PdfNumber(y), new PdfNumber(x + width), new PdfNumber(y + height),
                },
            },
            ["QuadPoints"] = new PdfArray
            {
                Items =
                {
                    new PdfNumber(x), new PdfNumber(y + height),
                    new PdfNumber(x + width), new PdfNumber(y + height),
                    new PdfNumber(x), new PdfNumber(y),
                    new PdfNumber(x + width), new PdfNumber(y),
                },
            },
            ["C"] = ColorArray(colorHex),
            ["F"] = new PdfNumber(4),
        };
        if (!text.IsNullOrEmpty()) dict["Contents"] = PdfObjectDocument.NewTextString(text!);
        if (!author.IsNullOrEmpty()) dict["T"] = PdfObjectDocument.NewTextString(author!);
        return new DictObj(dict);
    }

    private static PdfObject BuildNote(Single x, Single y, String text, String? author)
    {
        var dict = new PdfDict
        {
            ["Type"] = new PdfName("Annot"),
            ["Subtype"] = new PdfName("Text"),
            ["Rect"] = new PdfArray
            {
                Items = { new PdfNumber(x), new PdfNumber(y), new PdfNumber(x + 20), new PdfNumber(y + 20) },
            },
            ["Contents"] = PdfObjectDocument.NewTextString(text),
            ["Name"] = new PdfName("Comment"),
            ["F"] = new PdfNumber(4),
        };
        if (!author.IsNullOrEmpty()) dict["T"] = PdfObjectDocument.NewTextString(author!);
        return new DictObj(dict);
    }

    private static PdfArray ColorArray(String colorHex)
    {
        var array = new PdfArray();
        if (colorHex.Length >= 6)
        {
            array.Items.Add(new PdfNumber(Convert.ToInt32(colorHex.Substring(0, 2), 16) / 255f));
            array.Items.Add(new PdfNumber(Convert.ToInt32(colorHex.Substring(2, 2), 16) / 255f));
            array.Items.Add(new PdfNumber(Convert.ToInt32(colorHex.Substring(4, 2), 16) / 255f));
        }
        else
        {
            array.Items.Add(new PdfNumber(1));
            array.Items.Add(new PdfNumber(1));
            array.Items.Add(new PdfNumber(0));
        }
        return array;
    }
    #endregion
}

/// <summary>单条 PDF 补丁指令</summary>
public class PdfPatchOp
{
    /// <summary>指令名</summary>
    public String Op { get; set; } = String.Empty;

    /// <summary>参数字典</summary>
    public Dictionary<String, Object?> Args { get; set; } = [];

    /// <summary>读取字符串参数</summary>
    /// <param name="key">键</param>
    /// <returns>字符串值</returns>
    public String? GetString(String key) => JsonLite.GetString(Args, key);

    /// <summary>读取整数参数</summary>
    /// <param name="key">键</param>
    /// <returns>整数值</returns>
    public Int32? GetInt32(String key) => JsonLite.GetInt32(Args, key);

    /// <summary>读取浮点参数</summary>
    /// <param name="key">键</param>
    /// <returns>浮点值</returns>
    public Single? GetSingle(String key)
    {
        if (!Args.TryGetValue(key, out var value) || value == null) return null;
        if (value is Double d) return (Single)d;
        if (value is String s && Single.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)) return parsed;
        return null;
    }
}

/// <summary>PDF 补丁执行结果</summary>
public class PdfPatchResult
{
    /// <summary>成功的指令</summary>
    public List<String> Applied { get; } = [];

    /// <summary>失败的指令与原因</summary>
    public List<String> Failed { get; } = [];

    /// <summary>执行告警</summary>
    public List<String> Warnings { get; } = [];

    /// <summary>是否全部成功（为 false 时不应保存）</summary>
    public Boolean Success => Failed.Count == 0;

    /// <summary>输出 JSON 结果</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"success\": ").Append(Success ? "true" : "false").Append(',').Append(nl);
        sb.Append(pad).Append("\"applied\": ").Append(OfficeJson.Array(Applied)).Append(',').Append(nl);
        sb.Append(pad).Append("\"failed\": ").Append(OfficeJson.Array(Failed)).Append(',').Append(nl);
        sb.Append(pad).Append("\"warnings\": ").Append(OfficeJson.Array(Warnings)).Append(nl);
        sb.Append('}');
        return sb.ToString();
    }
}
