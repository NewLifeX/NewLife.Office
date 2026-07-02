namespace NewLife.Office.Word;

/// <summary>Word 四边边框集合，用于表格、单元格或段落边框设置</summary>
/// <remarks>
/// null 表示对应边不设置（使用默认或父级样式）。
/// <example>
/// <code>
/// var borders = new TableBorders
/// {
///     Top    = new Border { Style = BorderStyle.Single, Width = 8 },
///     Bottom = new Border { Style = BorderStyle.Single, Width = 8 },
///     Left   = new Border { Style = BorderStyle.None },
///     Right  = new Border { Style = BorderStyle.None },
///     InsideH = new Border { Style = BorderStyle.Dotted },
///     InsideV = new Border { Style = BorderStyle.None },
/// };
/// </code>
/// </example>
/// </remarks>
public class TableBorders
{
    #region 属性
    /// <summary>上边框</summary>
    public Border? Top { get; set; }

    /// <summary>下边框</summary>
    public Border? Bottom { get; set; }

    /// <summary>左边框</summary>
    public Border? Left { get; set; }

    /// <summary>右边框</summary>
    public Border? Right { get; set; }

    /// <summary>表格内部水平分隔线</summary>
    public Border? InsideH { get; set; }

    /// <summary>表格内部垂直分隔线</summary>
    public Border? InsideV { get; set; }
    #endregion

    #region 工厂方法
    /// <summary>创建四边统一边框</summary>
    /// <param name="style">线型</param>
    /// <param name="color">颜色（hex）</param>
    /// <param name="width">粗细（八分之一磅）</param>
    public static TableBorders All(BorderStyle style, String? color = null, Int32 width = 4)
    {
        var b = new Border { Style = style, Color = color, Width = width };
        return new TableBorders { Top = b, Bottom = b, Left = b, Right = b, InsideH = b, InsideV = b };
    }

    /// <summary>创建仅外框线（无内部分隔线）</summary>
    public static TableBorders OutlineOnly(BorderStyle style = BorderStyle.Single, String? color = null, Int32 width = 4)
    {
        var b = new Border { Style = style, Color = color, Width = width };
        var none = new Border { Style = BorderStyle.None };
        return new TableBorders { Top = b, Bottom = b, Left = b, Right = b, InsideH = none, InsideV = none };
    }
    #endregion
}
