using System.Text;

namespace NewLife.Office.Ppt;

/// <summary>演示文稿结构快照：面向 AI/Agent 的幻灯片与元素视图</summary>
/// <remarks>
/// 锚点形态：幻灯片 <c>s#2</c>，元素 <c>s#2/tb#0</c>（文本框）、<c>s#2/shape#0</c>（形状）、
/// <c>s#2/pic#0</c>（图片）、<c>s#2/tbl#0</c>（表格）、<c>s#2/chart#0</c>（图表）。
/// 序号是快照内的集合下标，修改后需重新取快照。
/// <example>
/// <code>
/// using var reader = new PptxReader("report.pptx");
/// var outline = reader.ReadDocument().ToOutline();
/// Console.WriteLine(outline.ToJson());
/// </code>
/// </example>
/// </remarks>
public class PptxOutline
{
    #region 属性
    /// <summary>幻灯片快照集合</summary>
    public List<PptxSlideOutline> Slides { get; } = [];
    #endregion

    #region 构造
    /// <summary>从演示文稿模型构建快照</summary>
    /// <param name="document">演示文稿模型</param>
    /// <returns>结构快照</returns>
    public static PptxOutline From(Presentation document)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var outline = new PptxOutline();
        for (var i = 0; i < document.Slides.Count; i++)
        {
            var slide = document.Slides[i];
            var item = new PptxSlideOutline
            {
                Index = i,
                Anchor = $"s#{i}",
                LayoutIndex = slide.LayoutIndex,
                Layout = slide.Layout,
                Hidden = slide.Hidden,
                Notes = slide.Notes,
            };

            for (var n = 0; n < slide.TextBoxes.Count; n++)
            {
                var tb = slide.TextBoxes[n];
                item.TextBoxes.Add(new PptxElementOutline
                {
                    Anchor = $"s#{i}/tb#{n}",
                    Kind = "textbox",
                    Text = tb.Text,
                    Role = tb.Role,
                    Left = tb.Left,
                    Top = tb.Top,
                    Width = tb.Width,
                    Height = tb.Height,
                    FontSize = tb.FontSize,
                    Bold = tb.Bold,
                });
            }

            for (var n = 0; n < slide.Shapes.Count; n++)
            {
                var sp = slide.Shapes[n];
                item.Shapes.Add(new PptxElementOutline
                {
                    Anchor = $"s#{i}/shape#{n}",
                    Kind = "shape",
                    Id = sp.Id,
                    Type = sp.ShapeType,
                    Text = sp.Text,
                    Left = sp.Left,
                    Top = sp.Top,
                    Width = sp.Width,
                    Height = sp.Height,
                    FontSize = sp.FontSize,
                    Bold = sp.Bold,
                });
            }

            for (var n = 0; n < slide.Images.Count; n++)
            {
                var pic = slide.Images[n];
                item.Pictures.Add(new PptxElementOutline
                {
                    Anchor = $"s#{i}/pic#{n}",
                    Kind = "picture",
                    Type = pic.Extension,
                    Left = pic.Left,
                    Top = pic.Top,
                    Width = pic.Width,
                    Height = pic.Height,
                });
            }

            for (var n = 0; n < slide.Tables.Count; n++)
            {
                var table = slide.Tables[n];
                var info = new PptxTableOutline
                {
                    Anchor = $"s#{i}/tbl#{n}",
                    Rows = table.Rows.Count,
                    Columns = table.Rows.Count > 0 ? table.Rows.Max(r => r.Length) : 0,
                };
                foreach (var row in table.Rows)
                    info.Cells.Add(row.Select(e => (String?)e).ToList());
                item.Tables.Add(info);
            }

            for (var n = 0; n < slide.Charts.Count; n++)
            {
                var chart = slide.Charts[n];
                var info = new PptxChartOutline
                {
                    Anchor = $"s#{i}/chart#{n}",
                    Type = chart.ChartType,
                    Title = chart.Title,
                };
                if (chart.Categories != null) info.Categories.AddRange(chart.Categories);
                foreach (var series in chart.Series)
                {
                    info.Series.Add(new PptxSeriesOutline
                    {
                        Name = series.Name,
                        Values = series.Values,
                    });
                }
                item.Charts.Add(info);
            }

            item.GroupCount = slide.Groups.Count;
            item.ConnectorCount = slide.Connectors.Count;
            outline.Slides.Add(item);
        }
        return outline;
    }
    #endregion

    #region JSON
    /// <summary>输出 JSON 快照</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"slideCount\": ").Append(Slides.Count).Append(',').Append(nl);
        sb.Append(pad).Append("\"slides\": [");
        for (var i = 0; i < Slides.Count; i++)
        {
            var slide = Slides[i];
            sb.Append(nl).Append(pad).Append(pad).Append('{');
            sb.Append("\"anchor\": ").Append(OfficeJson.Escape(slide.Anchor));
            sb.Append(", \"index\": ").Append(slide.Index);
            sb.Append(", \"layoutIndex\": ").Append(slide.LayoutIndex);
            if (!slide.Layout.IsNullOrEmpty()) sb.Append(", \"layout\": ").Append(OfficeJson.Escape(slide.Layout));
            if (slide.Hidden) sb.Append(", \"hidden\": true");
            if (!slide.Notes.IsNullOrEmpty()) sb.Append(", \"notes\": ").Append(OfficeJson.Escape(slide.Notes));
            sb.Append(", \"textBoxes\": ").Append(Elements(slide.TextBoxes, pretty));
            sb.Append(", \"shapes\": ").Append(Elements(slide.Shapes, pretty));
            if (slide.Pictures.Count > 0) sb.Append(", \"pictures\": ").Append(Elements(slide.Pictures, pretty));
            if (slide.Tables.Count > 0) sb.Append(", \"tables\": ").Append(Tables(slide.Tables, pretty));
            if (slide.Charts.Count > 0) sb.Append(", \"charts\": ").Append(Charts(slide.Charts, pretty));
            if (slide.GroupCount > 0) sb.Append(", \"groups\": ").Append(slide.GroupCount);
            if (slide.ConnectorCount > 0) sb.Append(", \"connectors\": ").Append(slide.ConnectorCount);
            sb.Append('}');
            if (i < Slides.Count - 1) sb.Append(',');
        }
        sb.Append(nl).Append(pad).Append(']').Append(nl).Append('}');
        return sb.ToString();
    }

    private static String Elements(List<PptxElementOutline> items, Boolean pretty)
    {
        var sb = new StringBuilder();
        sb.Append('[');
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            if (i > 0) sb.Append(",");
            if (pretty) sb.Append("\n      ");
            sb.Append('{');
            sb.Append("\"anchor\": ").Append(OfficeJson.Escape(item.Anchor));
            sb.Append(", \"kind\": ").Append(OfficeJson.Escape(item.Kind));
            if (item.Id != 0) sb.Append(", \"id\": ").Append(item.Id);
            if (!item.Type.IsNullOrEmpty()) sb.Append(", \"type\": ").Append(OfficeJson.Escape(item.Type));
            if (!item.Role.IsNullOrEmpty()) sb.Append(", \"role\": ").Append(OfficeJson.Escape(item.Role));
            if (!item.Text.IsNullOrEmpty()) sb.Append(", \"text\": ").Append(OfficeJson.Escape(item.Text));
            sb.Append(", \"left\": ").Append(item.Left);
            sb.Append(", \"top\": ").Append(item.Top);
            sb.Append(", \"width\": ").Append(item.Width);
            sb.Append(", \"height\": ").Append(item.Height);
            if (item.FontSize > 0) sb.Append(", \"fontSize\": ").Append(item.FontSize);
            if (item.Bold) sb.Append(", \"bold\": true");
            sb.Append('}');
        }
        if (pretty && items.Count > 0) sb.Append("\n    ");
        sb.Append(']');
        return sb.ToString();
    }

    private static String Tables(List<PptxTableOutline> items, Boolean pretty)
    {
        var sb = new StringBuilder();
        sb.Append('[');
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            if (i > 0) sb.Append(",");
            if (pretty) sb.Append("\n      ");
            sb.Append('{');
            sb.Append("\"anchor\": ").Append(OfficeJson.Escape(item.Anchor));
            sb.Append(", \"rows\": ").Append(item.Rows);
            sb.Append(", \"columns\": ").Append(item.Columns);
            sb.Append(", \"cells\": [");
            for (var r = 0; r < item.Cells.Count; r++)
            {
                if (r > 0) sb.Append(", ");
                sb.Append(OfficeJson.Array(item.Cells[r]));
            }
            sb.Append("]}");
        }
        if (pretty && items.Count > 0) sb.Append("\n    ");
        sb.Append(']');
        return sb.ToString();
    }

    private static String Charts(List<PptxChartOutline> items, Boolean pretty)
    {
        var sb = new StringBuilder();
        sb.Append('[');
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            if (i > 0) sb.Append(",");
            if (pretty) sb.Append("\n      ");
            sb.Append('{');
            sb.Append("\"anchor\": ").Append(OfficeJson.Escape(item.Anchor));
            sb.Append(", \"type\": ").Append(OfficeJson.Escape(item.Type));
            if (!item.Title.IsNullOrEmpty()) sb.Append(", \"title\": ").Append(OfficeJson.Escape(item.Title));
            sb.Append(", \"categories\": ").Append(OfficeJson.Array(item.Categories));
            sb.Append(", \"series\": [");
            for (var s = 0; s < item.Series.Count; s++)
            {
                if (s > 0) sb.Append(", ");
                var series = item.Series[s];
                sb.Append("{\"name\": ").Append(OfficeJson.Escape(series.Name)).Append(", \"values\": [");
                for (var v = 0; v < series.Values.Length; v++)
                {
                    if (v > 0) sb.Append(", ");
                    sb.Append(OfficeJson.Value(series.Values[v]));
                }
                sb.Append("]}");
            }
            sb.Append("]}");
        }
        if (pretty && items.Count > 0) sb.Append("\n    ");
        sb.Append(']');
        return sb.ToString();
    }
    #endregion
}

/// <summary>幻灯片结构快照</summary>
public class PptxSlideOutline
{
    /// <summary>锚点（如 <c>s#2</c>）</summary>
    public String Anchor { get; set; } = String.Empty;

    /// <summary>幻灯片索引（0 基）</summary>
    public Int32 Index { get; set; }

    /// <summary>版式索引</summary>
    public Int32 LayoutIndex { get; set; }

    /// <summary>自动排版策略（title_content 等），null 表示未设置</summary>
    public String? Layout { get; set; }

    /// <summary>是否隐藏</summary>
    public Boolean Hidden { get; set; }

    /// <summary>演讲者备注</summary>
    public String? Notes { get; set; }

    /// <summary>文本框</summary>
    public List<PptxElementOutline> TextBoxes { get; } = [];

    /// <summary>形状</summary>
    public List<PptxElementOutline> Shapes { get; } = [];

    /// <summary>图片</summary>
    public List<PptxElementOutline> Pictures { get; } = [];

    /// <summary>表格</summary>
    public List<PptxTableOutline> Tables { get; } = [];

    /// <summary>图表</summary>
    public List<PptxChartOutline> Charts { get; } = [];

    /// <summary>组形状数量</summary>
    public Int32 GroupCount { get; set; }

    /// <summary>连接线数量</summary>
    public Int32 ConnectorCount { get; set; }
}

/// <summary>幻灯片元素快照</summary>
public class PptxElementOutline
{
    /// <summary>锚点（如 <c>s#2/tb#0</c>）</summary>
    public String Anchor { get; set; } = String.Empty;

    /// <summary>元素类别：textbox/shape/picture</summary>
    public String Kind { get; set; } = String.Empty;

    /// <summary>形状 ID（仅 shape）</summary>
    public Int32 Id { get; set; }

    /// <summary>类型（形状几何类型或图片扩展名）</summary>
    public String? Type { get; set; }

    /// <summary>占位符角色（title/subtitle/body 等）</summary>
    public String? Role { get; set; }

    /// <summary>文本</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>字号（磅）</summary>
    public Int32 FontSize { get; set; }

    /// <summary>是否粗体</summary>
    public Boolean Bold { get; set; }
}

/// <summary>表格快照</summary>
public class PptxTableOutline
{
    /// <summary>锚点（如 <c>s#2/tbl#0</c>）</summary>
    public String Anchor { get; set; } = String.Empty;

    /// <summary>行数</summary>
    public Int32 Rows { get; set; }

    /// <summary>列数</summary>
    public Int32 Columns { get; set; }

    /// <summary>单元格文本矩阵</summary>
    public List<List<String?>> Cells { get; } = [];
}

/// <summary>图表快照</summary>
public class PptxChartOutline
{
    /// <summary>锚点（如 <c>s#2/chart#0</c>）</summary>
    public String Anchor { get; set; } = String.Empty;

    /// <summary>图表类型</summary>
    public String Type { get; set; } = String.Empty;

    /// <summary>图表标题</summary>
    public String? Title { get; set; }

    /// <summary>分类轴标签</summary>
    public List<String> Categories { get; } = [];

    /// <summary>数据系列</summary>
    public List<PptxSeriesOutline> Series { get; } = [];
}

/// <summary>图表系列快照</summary>
public class PptxSeriesOutline
{
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数据点</summary>
    public Double[] Values { get; set; } = [];
}
