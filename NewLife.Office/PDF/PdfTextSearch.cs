using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office.Pdf;

/// <summary>PDF 文本搜索与按阅读顺序取文本</summary>
/// <remarks>
/// 命中项返回页码与矩形坐标（原点在左下角），可直接作为高亮/便签/文字叠加的锚点。
/// 文本块坐标由内容流文本矩阵推算，宽度按字符数估算（CJK 按 1 字宽、西文按 0.5 字宽），用于标注框大致定位。
/// </remarks>
public static class PdfTextSearch
{
    /// <summary>在整份 PDF 中搜索文本</summary>
    /// <param name="reader">PDF 读取器</param>
    /// <param name="text">搜索内容（或正则）</param>
    /// <param name="regex">是否按正则表达式匹配</param>
    /// <param name="caseSensitive">是否区分大小写</param>
    /// <param name="maxResults">最大命中数</param>
    /// <returns>命中列表</returns>
    public static List<PdfTextHit> SearchText(this PdfReader reader, String text, Boolean regex = false, Boolean caseSensitive = false, Int32 maxResults = 100)
    {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (maxResults <= 0) maxResults = 100;

        var hits = new List<PdfTextHit>();
        Regex? rx = null;
        if (regex)
        {
            var options = RegexOptions.None;
            if (!caseSensitive) options |= RegexOptions.IgnoreCase;
            rx = new Regex(text, options);
        }

        var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var pageCount = reader.GetPageCount();
        for (var page = 0; page < pageCount; page++)
        {
            foreach (var block in OrderedBlocks(reader, page))
            {
                if (block.Text.IsNullOrEmpty()) continue;

                if (rx != null)
                {
                    foreach (Match match in rx.Matches(block.Text))
                    {
                        if (!match.Success || match.Length == 0) continue;
                        hits.Add(BuildHit(page, block, match.Index, match.Value));
                        if (hits.Count >= maxResults) return hits;
                    }
                }
                else
                {
                    var index = 0;
                    while (index < block.Text.Length)
                    {
                        var found = block.Text.IndexOf(text, index, comparison);
                        if (found < 0) break;
                        hits.Add(BuildHit(page, block, found, block.Text.Substring(found, text.Length)));
                        index = found + Math.Max(1, text.Length);
                        if (hits.Count >= maxResults) return hits;
                    }
                }
            }
        }
        return hits;
    }

    /// <summary>按阅读顺序读取整页文本</summary>
    /// <param name="reader">PDF 读取器</param>
    /// <param name="pageIndex">页索引（0 基）</param>
    /// <param name="lineTolerance">同一行的 Y 容差（点）</param>
    /// <returns>逐行文本</returns>
    public static List<PdfTextLine> ReadPageLines(this PdfReader reader, Int32 pageIndex, Single lineTolerance = 3f)
    {
        if (reader == null) throw new ArgumentNullException(nameof(reader));

        var lines = new List<PdfTextLine>();
        var blocks = reader.GetPageTextBlocks(pageIndex);
        if (blocks.Count == 0) return lines;

        var ordered = blocks
            .Where(e => !e.Text.IsNullOrEmpty())
            .OrderByDescending(e => e.Y)
            .ThenBy(e => e.X)
            .ToList();

        PdfTextLine? current = null;
        foreach (var block in ordered)
        {
            if (current == null || Math.Abs(current.Y - block.Y) > lineTolerance)
            {
                current = new PdfTextLine { Y = block.Y };
                lines.Add(current);
            }
            current.Blocks.Add(block);
        }

        foreach (var line in lines)
        {
            line.Blocks.Sort((a, b) => a.X.CompareTo(b.X));
            var sb = new StringBuilder();
            var lastRight = (Single?)null;
            foreach (var block in line.Blocks)
            {
                if (lastRight.HasValue && block.X - lastRight.Value > 1f) sb.Append(' ');
                sb.Append(block.Text);
                lastRight = block.X + Measure(block.Text, block.FontSize);
            }
            line.Text = sb.ToString();
        }

        return lines;
    }

    private static List<PdfText> OrderedBlocks(PdfReader reader, Int32 pageIndex)
    {
        var blocks = reader.GetPageTextBlocks(pageIndex);
        return blocks
            .Where(e => !e.Text.IsNullOrEmpty())
            .OrderByDescending(e => e.Y)
            .ThenBy(e => e.X)
            .ToList();
    }

    private static PdfTextHit BuildHit(Int32 pageIndex, PdfText block, Int32 offset, String match)
    {
        var fontSize = block.FontSize > 0 ? block.FontSize : 10f;
        var prefix = offset > 0 ? block.Text[..offset] : String.Empty;
        var x = block.X + Measure(prefix, fontSize);

        return new PdfTextHit
        {
            PageIndex = pageIndex,
            Text = match,
            BlockText = block.Text,
            X = x,
            Y = block.Y,
            Width = Measure(match, fontSize),
            Height = fontSize * 1.2f,
        };
    }

    /// <summary>估算文本宽度（CJK 按 1 字宽、西文按 0.5 字宽）</summary>
    /// <param name="text">文本</param>
    /// <param name="fontSize">字号</param>
    /// <returns>宽度（点）</returns>
    public static Single Measure(String? text, Single fontSize = 10f)
    {
        if (text == null) return 0;
        var size = fontSize > 0 ? fontSize : 10f;
        var width = 0f;
        foreach (var ch in text)
            width += ch >= 0x2E80 ? size : size * 0.5f;
        return width;
    }
}

/// <summary>文本命中项</summary>
public class PdfTextHit
{
    /// <summary>页索引（0 基）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>命中的文本片段</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>命中所在的完整文本块</summary>
    public String BlockText { get; set; } = String.Empty;

    /// <summary>命中框左下角 X（点）</summary>
    public Single X { get; set; }

    /// <summary>命中框左下角 Y（点）</summary>
    public Single Y { get; set; }

    /// <summary>命中框宽度（点，估算）</summary>
    public Single Width { get; set; }

    /// <summary>命中框高度（点，估算）</summary>
    public Single Height { get; set; }

    /// <summary>页码（1 起始）</summary>
    public Int32 Page => PageIndex + 1;
}

/// <summary>按阅读顺序聚合的一行文本</summary>
public class PdfTextLine
{
    /// <summary>行基线 Y 坐标</summary>
    public Single Y { get; set; }

    /// <summary>行文本</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>构成该行的文本块</summary>
    public List<PdfText> Blocks { get; } = [];
}

/// <summary>文本搜索结果</summary>
public class PdfTextSearchResult
{
    /// <summary>命中列表</summary>
    public List<PdfTextHit> Hits { get; } = [];

    /// <summary>是否截断</summary>
    public Boolean Truncated { get; set; }

    /// <summary>输出 JSON</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"count\": ").Append(Hits.Count).Append(',').Append(nl);
        sb.Append(pad).Append("\"truncated\": ").Append(Truncated ? "true" : "false").Append(',').Append(nl);
        sb.Append(pad).Append("\"hits\": [");
        for (var i = 0; i < Hits.Count; i++)
        {
            var hit = Hits[i];
            if (i > 0) sb.Append(',');
            sb.Append(nl).Append(pad).Append(pad).Append('{');
            sb.Append("\"page\": ").Append(hit.Page);
            sb.Append(", \"text\": ").Append(OfficeJson.Escape(hit.Text));
            sb.Append(", \"x\": ").Append(OfficeJson.Value(hit.X));
            sb.Append(", \"y\": ").Append(OfficeJson.Value(hit.Y));
            sb.Append(", \"width\": ").Append(OfficeJson.Value(hit.Width));
            sb.Append(", \"height\": ").Append(OfficeJson.Value(hit.Height));
            sb.Append('}');
        }
        sb.Append(nl).Append(pad).Append(']').Append(nl).Append('}');
        return sb.ToString();
    }
}
