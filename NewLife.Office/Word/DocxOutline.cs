using System.Text;

namespace NewLife.Office.Word;

/// <summary>文档结构快照：面向 AI/Agent 的块级视图，每块带稳定锚点便于后续定位修改</summary>
/// <remarks>
/// 与 <see cref="WordReader.ReadFullText"/> 的区别：本类保留块边界、类型、标题层级与表格矩阵，
/// 并给每块分配锚点（<c>p#3</c> / <c>tbl#5</c> / <c>img#7</c>，数字为 <see cref="Element"/> 在正文中的序号），
/// 让模型能"先看结构、再按锚点改"。
/// <para>锚点在"同一份快照"内稳定；执行修改后需重新取快照。</para>
/// <example>
/// <code>
/// using var reader = new WordReader("report.docx");
/// var outline = reader.ReadDocument().ToOutline();
/// Console.WriteLine(outline.ToJson());     // 结构化 JSON，可直接交给 LLM
/// </code>
/// </example>
/// </remarks>
public class DocxOutline
{
    #region 属性
    /// <summary>块列表（按文档顺序）</summary>
    public List<OutlineBlock> Blocks { get; } = [];

    /// <summary>标题块数量</summary>
    public Int32 HeadingCount => Blocks.Count(b => b.Level > 0);

    /// <summary>表格块数量</summary>
    public Int32 TableCount => Blocks.Count(b => b.Type == "table");
    #endregion

    #region 构造
    /// <summary>从文档模型构建快照</summary>
    /// <param name="doc">文档</param>
    /// <returns>结构快照</returns>
    /// <exception cref="ArgumentNullException">doc 为空</exception>
    public static DocxOutline From(Document doc)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        var outline = new DocxOutline();
        for (var i = 0; i < doc.Elements.Count; i++)
        {
            var el = doc.Elements[i];
            var block = new OutlineBlock { Index = i };

            if (el.Paragraph != null)
            {
                var para = el.Paragraph;
                block.Type = para.Style == ParagraphStyle.Normal ? "paragraph" : "heading";
                block.Level = para.Style == ParagraphStyle.Normal ? 0 : (Int32)para.Style;
                block.Style = para.StyleId ?? (para.Style == ParagraphStyle.Normal ? "Normal" : $"Heading{(Int32)para.Style}");
                block.Text = ParagraphText(para);
                block.HasUnmodeledContent = para.HasUnmodeledContent;
                block.Anchor = $"p#{i}";
            }
            else if (el.Table != null)
            {
                block.Type = "table";
                block.Anchor = $"tbl#{i}";
                block.Rows = el.Table.Rows.Count;
                block.Columns = el.Table.Rows.Count > 0 ? el.Table.Rows.Max(r => r.Cells.Count) : 0;
                block.Cells = [];
                foreach (var row in el.Table.Rows)
                {
                    var cells = new List<String>();
                    foreach (var cell in row.Cells)
                    {
                        var sb = new StringBuilder();
                        foreach (var p in cell.Paragraphs) sb.Append(ParagraphText(p));
                        cells.Add(sb.ToString());
                    }
                    block.Cells.Add(cells);
                }
                block.Text = String.Empty;
            }
            else if (el.Image != null)
            {
                block.Type = "image";
                block.Anchor = $"img#{i}";
                block.Text = el.Image.AltText ?? String.Empty;
                block.Style = el.Image.Extension;
                block.Rows = (Int32)Math.Round(el.Image.WidthEmu / Length.EmuPerCm, 2);
                block.Columns = (Int32)Math.Round(el.Image.HeightEmu / Length.EmuPerCm, 2);
            }
            else if (el.Type == ElementType.Unknown)
            {
                block.Type = "unknown";
                block.Anchor = $"x#{i}";
                block.HasUnmodeledContent = true;
            }
            else
            {
                block.Type = el.Type.ToString();
                block.Anchor = $"e#{i}";
            }

            outline.Blocks.Add(block);
        }
        return outline;
    }

    /// <summary>拼接段落文本（域 Run 取缓存结果）</summary>
    private static String ParagraphText(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var run in para.Runs)
        {
            if (!String.IsNullOrEmpty(run.FieldCode))
            {
                if (!String.IsNullOrEmpty(run.FieldResult)) sb.Append(run.FieldResult);
                continue;
            }
            sb.Append(run.Text);
        }
        return sb.ToString();
    }
    #endregion

    #region JSON
    /// <summary>输出 JSON（结构化快照，可直接交给 LLM 或写入文件）</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"headingCount\": ").Append(HeadingCount).Append(',').Append(nl);
        sb.Append(pad).Append("\"tableCount\": ").Append(TableCount).Append(',').Append(nl);
        sb.Append(pad).Append("\"blocks\": [");
        for (var i = 0; i < Blocks.Count; i++)
        {
            var b = Blocks[i];
            sb.Append(nl).Append(pad).Append(pad).Append('{');
            sb.Append("\"anchor\": ").Append(Json(b.Anchor)).Append(", ");
            sb.Append("\"index\": ").Append(b.Index).Append(", ");
            sb.Append("\"type\": ").Append(Json(b.Type));
            if (b.Level > 0) sb.Append(", \"level\": ").Append(b.Level);
            if (!String.IsNullOrEmpty(b.Style)) sb.Append(", \"style\": ").Append(Json(b.Style));
            sb.Append(", \"text\": ").Append(Json(b.Text));
            if (b.HasUnmodeledContent) sb.Append(", \"hasUnmodeledContent\": true");
            if (b.Cells != null)
            {
                sb.Append(", \"rows\": ").Append(b.Rows).Append(", \"columns\": ").Append(b.Columns);
                sb.Append(", \"cells\": [");
                for (var r = 0; r < b.Cells.Count; r++)
                {
                    if (r > 0) sb.Append(", ");
                    sb.Append('[');
                    var row = b.Cells[r];
                    for (var c = 0; c < row.Count; c++)
                    {
                        if (c > 0) sb.Append(", ");
                        sb.Append(Json(row[c]));
                    }
                    sb.Append(']');
                }
                sb.Append(']');
            }
            sb.Append('}');
            if (i < Blocks.Count - 1) sb.Append(',');
        }
        sb.Append(nl).Append(pad).Append(']').Append(nl).Append('}');
        return sb.ToString();
    }

    /// <summary>JSON 字符串转义</summary>
    private static String Json(String? value)
    {
        if (value == null) return "null";

        var sb = new StringBuilder(value.Length + 2);
        sb.Append('"');
        foreach (var ch in value)
        {
            switch (ch)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (ch < ' ') sb.Append("\\u").Append(((Int32)ch).ToString("x4"));
                    else sb.Append(ch);
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }
    #endregion
}

/// <summary>结构快照中的单个块</summary>
public class OutlineBlock
{
    #region 属性
    /// <summary>稳定锚点：<c>p#3</c>（段落）/ <c>tbl#5</c>（表格）/ <c>img#7</c>（图片）/ <c>x#9</c>（未建模）</summary>
    public String Anchor { get; set; } = String.Empty;

    /// <summary>该块在 <see cref="Document.Elements"/> 中的序号</summary>
    public Int32 Index { get; set; }

    /// <summary>块类型：paragraph / heading / table / image / unknown</summary>
    public String Type { get; set; } = String.Empty;

    /// <summary>标题级别（1~6，0 表示非标题）</summary>
    public Int32 Level { get; set; }

    /// <summary>样式 ID（如 Normal、Heading2、自定义样式名）</summary>
    public String? Style { get; set; }

    /// <summary>纯文本（表格块为空，见 <see cref="Cells"/>）</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>是否含模型无法表达的内容（域/OLE/批注锚点等），修改该块可能丢失内容</summary>
    public Boolean HasUnmodeledContent { get; set; }

    /// <summary>表格行数（表格块与图片块记录尺寸）</summary>
    public Int32 Rows { get; set; }

    /// <summary>表格列数</summary>
    public Int32 Columns { get; set; }

    /// <summary>表格单元格文本矩阵（仅表格块）</summary>
    public List<List<String>>? Cells { get; set; }
    #endregion
}
