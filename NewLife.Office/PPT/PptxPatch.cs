using System.Globalization;
using System.Text;

namespace NewLife.Office.Ppt;

/// <summary>演示文稿 JSON 补丁：按幻灯片/元素锚点批量修改，任一条失败不写文件</summary>
/// <remarks>
/// 锚点：<c>s#2</c>（幻灯片）、<c>s#2/tb#0</c>（文本框）、<c>s#2/shape#0</c>（形状）、
/// <c>s#2/pic#0</c>（图片）、<c>s#2/tbl#0</c>（表格）、<c>s#2/chart#0</c>（图表）。
/// 执行只改内存模型，调用方应在 <see cref="PptxPatchResult.Success"/> 为 true 时再保存。
/// <example>
/// <code>
/// var patch = PptxPatch.ParseJson("""
/// { "ops": [
///   { "op": "set_shape_text", "anchor": "s#0/tb#1", "text": "2026 年第三季度" },
///   { "op": "set_notes", "anchor": "s#0", "text": "汇报要点：营收增长" }
/// ] }
/// """);
/// using var reader = new PptxReader("in.pptx");
/// var doc = reader.ReadDocument();
/// if (patch.Apply(doc).Success) doc.Save("out.pptx", "in.pptx");
/// </code>
/// </example>
/// </remarks>
public class PptxPatch
{
    /// <summary>每厘米对应的 EMU</summary>
    private const Int64 EmuPerCm = 360000;

    #region 属性
    /// <summary>指令集合</summary>
    public List<PptxPatchOp> Ops { get; } = [];
    #endregion

    #region 解析
    /// <summary>解析 JSON 补丁</summary>
    /// <param name="json">JSON 文本</param>
    /// <returns>补丁对象</returns>
    public static PptxPatch ParseJson(String json)
    {
        if (json == null) throw new ArgumentNullException(nameof(json));

        var root = JsonLite.Parse(json) as Dictionary<String, Object?>
            ?? throw new FormatException(L.T("Patch root must be a JSON object."));
        if (!root.TryGetValue("ops", out var opsValue) || opsValue is not List<Object?> list)
            throw new FormatException(L.T("Patch must contain an 'ops' array."));

        var patch = new PptxPatch();
        foreach (var item in list)
        {
            if (item is not Dictionary<String, Object?> obj)
                throw new FormatException(L.T("Each op must be a JSON object."));
            patch.Ops.Add(new PptxPatchOp
            {
                Op = JsonLite.GetString(obj, "op") ?? throw new FormatException(L.T("Each op must have an 'op' name.")),
                Args = obj,
            });
        }
        return patch;
    }
    #endregion

    #region 执行
    /// <summary>在演示文稿模型上执行补丁</summary>
    /// <param name="document">演示文稿模型</param>
    /// <returns>执行结果</returns>
    public PptxPatchResult Apply(Presentation document)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var result = new PptxPatchResult();
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
                result.FailedCodes.Add(ex.GetType().Name);
            }
        }
        return result;
    }

    private static void ApplyOp(Presentation document, PptxPatchOp op, PptxPatchResult result)
    {
        switch (op.Op)
        {
            case "set_shape_text":
            {
                var anchor = PptxAnchor.Parse(op.GetString("anchor") ?? throw new ArgumentException(L.T("set_shape_text requires 'anchor'.")));
                var slide = GetSlide(document, anchor.SlideIndex);
                var text = op.GetString("text") ?? String.Empty;

                switch (anchor.Kind)
                {
                    case "tb":
                        SetTextBoxText(GetItem(slide.TextBoxes, anchor.Index), text);
                        break;
                    case "shape":
                        GetItem(slide.Shapes, anchor.Index).Text = text;
                        break;
                    case "tbl":
                        throw new ArgumentException(L.T("Use set_table_cell for tables."));
                    case "chart":
                        throw new ArgumentException(L.T("Use update_chart for charts."));
                    case "pic":
                        throw new ArgumentException(L.T("Use replace_image for pictures."));
                    case null:
                        throw new ArgumentException(L.T("set_shape_text requires a concrete element anchor, such as s#0/tb#0."));
                    default:
                        throw new ArgumentException(L.T("Unsupported element type: {0}", anchor.Kind));
                }
                return;
            }

            case "set_table_cell":
            {
                var anchor = PptxAnchor.Parse(op.GetString("anchor") ?? throw new ArgumentException(L.T("set_table_cell requires 'anchor'.")));
                if (anchor.Kind != "tbl") throw new ArgumentException(L.T("set_table_cell requires a table anchor, such as s#0/tbl#0."));
                var slide = GetSlide(document, anchor.SlideIndex);
                var table = GetItem(slide.Tables, anchor.Index);
                var row = op.GetInt32("row") ?? throw new ArgumentException(L.T("set_table_cell requires 'row'."));
                var column = op.GetInt32("column") ?? throw new ArgumentException(L.T("set_table_cell requires 'column'."));
                if (row < 0 || row >= table.Rows.Count) throw new ArgumentOutOfRangeException(nameof(row), L.T("Row index out of range: {0}", row));
                var cells = table.Rows[row];
                if (column < 0 || column >= cells.Length) throw new ArgumentOutOfRangeException(nameof(column), L.T("Column index out of range: {0}", column));
                cells[column] = op.GetString("text") ?? String.Empty;
                return;
            }

            case "update_chart":
            {
                var anchor = PptxAnchor.Parse(op.GetString("anchor") ?? throw new ArgumentException(L.T("update_chart requires 'anchor'.")));
                if (anchor.Kind != "chart") throw new ArgumentException(L.T("update_chart requires a chart anchor, such as s#0/chart#0."));
                var slide = GetSlide(document, anchor.SlideIndex);
                var chart = GetItem(slide.Charts, anchor.Index);

                var categories = op.GetStrings("categories");
                if (categories != null) chart.Categories = categories;

                if (op.Args.TryGetValue("series", out var seriesValue) && seriesValue is List<Object?> seriesList)
                {
                    chart.Series.Clear();
                    foreach (var item in seriesList)
                    {
                        if (item is not Dictionary<String, Object?> obj) continue;
                        var series = new ChartSeries { Name = JsonLite.GetString(obj, "name") ?? String.Empty };
                        if (obj.TryGetValue("values", out var valuesValue) && valuesValue is List<Object?> values)
                        {
                            var data = new Double[values.Count];
                            for (var i = 0; i < values.Count; i++)
                                data[i] = values[i] == null ? 0d : Convert.ToDouble(values[i], CultureInfo.InvariantCulture);
                            series.Values = data;
                        }
                        chart.Series.Add(series);
                    }
                }

                var title = op.GetString("title");
                if (title != null) chart.Title = title;

                if (chart.Categories.Length > 0)
                {
                    foreach (var series in chart.Series)
                        if (series.Values.Length != chart.Categories.Length)
                            throw new ArgumentException(L.T("Series '{0}' has {1} data points but there are {2} categories.", series.Name, series.Values.Length, chart.Categories.Length));
                }
                return;
            }

            case "replace_image":
            {
                var anchor = PptxAnchor.Parse(op.GetString("anchor") ?? throw new ArgumentException(L.T("replace_image requires 'anchor'.")));
                if (anchor.Kind != "pic") throw new ArgumentException(L.T("replace_image requires a picture anchor, such as s#0/pic#0."));
                var slide = GetSlide(document, anchor.SlideIndex);
                var picture = GetItem(slide.Images, anchor.Index);

                var data = op.GetBytes();
                if (data == null || data.Length == 0) throw new ArgumentException(L.T("replace_image requires 'base64' or 'path'."));
                picture.Data = data;
                var extension = op.GetString("extension");
                var sourcePath = op.GetString("path");
                if (!extension.IsNullOrEmpty()) picture.Extension = extension!;
                else if (!sourcePath.IsNullOrEmpty()) picture.Extension = Path.GetExtension(sourcePath!).TrimStart('.').ToLowerInvariant();
                else if (SniffImageExtension(data) is { } sniffed) picture.Extension = sniffed;
                return;
            }

            case "set_notes":
            {
                var anchor = PptxAnchor.Parse(op.GetString("anchor") ?? throw new ArgumentException(L.T("set_notes requires 'anchor'.")));
                var slide = GetSlide(document, anchor.SlideIndex);
                slide.Notes = op.GetString("text");
                return;
            }

            case "set_title":
            {
                var anchor = PptxAnchor.Parse(op.GetString("anchor") ?? throw new ArgumentException(L.T("set_title requires 'anchor'.")));
                var slide = GetSlide(document, anchor.SlideIndex);
                var text = op.GetString("text") ?? String.Empty;
                var title = slide.TextBoxes.FirstOrDefault(e => String.Equals(e.Role, "title", StringComparison.OrdinalIgnoreCase)
                        || String.Equals(e.Role, "ctrTitle", StringComparison.OrdinalIgnoreCase))
                    ?? slide.TextBoxes.FirstOrDefault();
                if (title == null) throw new ArgumentException(L.T("Slide s#{0} has no text box usable as a title.", anchor.SlideIndex));
                SetTextBoxText(title, text);
                return;
            }

            case "add_slide":
            {
                var layout = op.GetInt32("layoutIndex") ?? 0;
                var slide = new Slide { LayoutIndex = layout };
                var title = op.GetString("title");
                var body = op.GetString("text");
                if (!title.IsNullOrEmpty())
                {
                    slide.TextBoxes.Add(new TextBox
                    {
                        Text = title!,
                        Role = "title",
                        Left = EmuPerCm,
                        Top = EmuPerCm,
                        Width = Math.Max(EmuPerCm, document.SlideWidth - 2 * EmuPerCm),
                        Height = 2 * EmuPerCm,
                        FontSize = 32,
                        Bold = true,
                    });
                }
                if (!body.IsNullOrEmpty())
                {
                    slide.TextBoxes.Add(new TextBox
                    {
                        Text = body!,
                        Role = "body",
                        Left = EmuPerCm,
                        Top = 3 * EmuPerCm,
                        Width = Math.Max(EmuPerCm, document.SlideWidth - 2 * EmuPerCm),
                        Height = Math.Max(EmuPerCm, document.SlideHeight - 4 * EmuPerCm),
                        FontSize = 18,
                    });
                }
                var notes = op.GetString("notes");
                if (!notes.IsNullOrEmpty()) slide.Notes = notes;

                var at = op.GetInt32("at") ?? document.Slides.Count;
                if (at < 0) at = 0;
                if (at > document.Slides.Count) at = document.Slides.Count;
                document.Slides.Insert(at, slide);
                result.Warnings.Add(L.T("add_slide: inserted a new slide at position {0}.", at + 1));
                return;
            }

            case "delete_slide":
            {
                var index = ResolveSlideIndex(document, op);
                if (document.Slides.Count <= 1) throw new ArgumentException(L.T("Cannot delete the last slide."));
                document.Slides.RemoveAt(index);
                return;
            }

            case "move_slide":
            {
                var from = op.GetInt32("from") ?? throw new ArgumentException(L.T("move_slide requires 'from'."));
                var to = op.GetInt32("to") ?? throw new ArgumentException(L.T("move_slide requires 'to'."));
                if (from < 0 || from >= document.Slides.Count) throw new ArgumentOutOfRangeException(nameof(from), L.T("Slide index out of range: {0}", from));
                if (to < 0 || to >= document.Slides.Count) throw new ArgumentOutOfRangeException(nameof(to), L.T("Slide index out of range: {0}", to));
                var slide = document.Slides[from];
                document.Slides.RemoveAt(from);
                document.Slides.Insert(to, slide);
                return;
            }

            case "clone_slide":
            {
                var from = op.GetInt32("from") ?? throw new ArgumentException(L.T("clone_slide requires 'from'."));
                var slide = GetSlide(document, from);
                var clone = CloneSlide(slide);
                var at = op.GetInt32("to") ?? document.Slides.Count;
                if (at < 0) at = 0;
                if (at > document.Slides.Count) at = document.Slides.Count;
                document.Slides.Insert(at, clone);
                return;
            }

            default:
                throw new ArgumentException(L.T("Unknown operation: {0}", op.Op));
        }
    }
    #endregion

    #region 辅助
    private static Slide GetSlide(Presentation document, Int32 index)
    {
        if (index < 0 || index >= document.Slides.Count) throw new ArgumentOutOfRangeException(nameof(index), L.T("Slide index out of range: {0}", index));
        return document.Slides[index];
    }

    private static T GetItem<T>(List<T> list, Int32 index)
    {
        if (index < 0 || index >= list.Count) throw new ArgumentOutOfRangeException(nameof(index), L.T("Element index out of range: {0}", index));
        return list[index];
    }

    private static Int32 ResolveSlideIndex(Presentation document, PptxPatchOp op)
    {
        var anchor = op.GetString("anchor");
        if (!anchor.IsNullOrEmpty()) return PptxAnchor.Parse(anchor!).SlideIndex;
        var slide = op.GetInt32("slide") ?? throw new ArgumentException(L.T("{0} requires 'anchor' or 'slide'.", op.Op));
        if (slide < 0 || slide >= document.Slides.Count) throw new ArgumentOutOfRangeException(nameof(slide), L.T("Slide index out of range: {0}", slide));
        return slide;
    }

    private static void SetTextBoxText(TextBox box, String text)
    {
        // 段落级富文本优先于 Text，必须一并清理，否则改写不生效
        box.Text = text;
        box.Paragraphs.Clear();
        box.Runs.Clear();
    }

    /// <summary>按魔数推断图片扩展名（base64 分支没有文件名可用）</summary>
    /// <param name="data">图片字节</param>
    /// <returns>扩展名，无法识别返回 null</returns>
    private static String? SniffImageExtension(Byte[] data)
    {
        if (data.Length >= 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) return "png";
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF) return "jpg";
        if (data.Length >= 6 && data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46) return "gif";
        if (data.Length >= 2 && data[0] == 0x42 && data[1] == 0x4D) return "bmp";
        if (data.Length >= 12 && data[0] == 0x52 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x46
            && data[8] == 0x57 && data[9] == 0x45 && data[10] == 0x42 && data[11] == 0x50) return "webp";
        if (data.Length >= 3 && ((data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A) || (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00))) return "tiff";
        return null;
    }

    private static Slide CloneSlide(Slide source) => PptxCloner.CloneSlide(source);

    #endregion
}

/// <summary>补丁锚点（<c>s#2</c> 或 <c>s#2/tb#0</c>）</summary>
public class PptxAnchor
{
    /// <summary>幻灯片索引</summary>
    public Int32 SlideIndex { get; set; }

    /// <summary>元素类别（tb/shape/pic/tbl/chart/grp/conn），null 表示整页</summary>
    public String? Kind { get; set; }

    /// <summary>元素索引</summary>
    public Int32 Index { get; set; }

    /// <summary>解析锚点</summary>
    /// <param name="text">锚点文本</param>
    /// <returns>锚点对象</returns>
    public static PptxAnchor Parse(String text)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));

        var s = text.Trim();
        if (!s.StartsWith("s#", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException(L.T("Invalid anchor: {0}", text));

        var anchor = new PptxAnchor();
        var slash = s.IndexOf('/');
        var slidePart = slash >= 0 ? s[..slash] : s;
        if (!Int32.TryParse(slidePart[2..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var slideIndex))
            throw new ArgumentException(L.T("Invalid anchor: {0}", text));
        anchor.SlideIndex = slideIndex;

        if (slash >= 0)
        {
            var rest = s[(slash + 1)..];
            var hash = rest.IndexOf('#');
            if (hash <= 0) throw new ArgumentException(L.T("Invalid anchor: {0}", text));
            anchor.Kind = rest[..hash].ToLowerInvariant();
            if (!Int32.TryParse(rest[(hash + 1)..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
                throw new ArgumentException(L.T("Invalid anchor: {0}", text));
            anchor.Index = index;
        }

        return anchor;
    }

    /// <summary>输出锚点文本</summary>
    /// <returns>锚点</returns>
    public override String ToString() => Kind == null ? $"s#{SlideIndex}" : $"s#{SlideIndex}/{Kind}#{Index}";
}

/// <summary>单条 PPT 补丁指令</summary>
public class PptxPatchOp
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

    /// <summary>读取字符串数组参数</summary>
    /// <param name="key">键</param>
    /// <returns>字符串数组</returns>
    public String[]? GetStrings(String key)
    {
        if (!Args.TryGetValue(key, out var value) || value is not List<Object?> list) return null;
        return list.Select(e => Convert.ToString(e, CultureInfo.InvariantCulture) ?? String.Empty).ToArray();
    }

    /// <summary>内嵌媒体上限（32MB），防止超大文件吃满内存</summary>
    internal const Int64 MaxMediaBytes = 32L * 1024 * 1024;

    /// <summary>读取图片字节（base64 或 path）</summary>
    /// <returns>字节数据，缺失返回 null</returns>
    public Byte[]? GetBytes()
    {
        var base64 = GetString("base64");
        if (!base64.IsNullOrEmpty())
        {
            var data = Convert.FromBase64String(base64!);
            if (data.Length > MaxMediaBytes) throw new ArgumentException(L.T("The embedded media exceeds the {0} MB limit.", MaxMediaBytes / 1024 / 1024));
            return data;
        }

        var path = GetString("path");
        if (!path.IsNullOrEmpty() && File.Exists(path))
        {
            if (new FileInfo(path).Length > MaxMediaBytes) throw new ArgumentException(L.T("The embedded media exceeds the {0} MB limit.", MaxMediaBytes / 1024 / 1024));
            return File.ReadAllBytes(path!);
        }
        return null;
    }
}

/// <summary>PPT 补丁执行结果</summary>
public class PptxPatchResult
{
    /// <summary>成功的指令</summary>
    public List<String> Applied { get; } = [];

    /// <summary>失败的指令与原因</summary>
    public List<String> Failed { get; } = [];

    internal List<String> FailedCodes { get; } = [];

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
