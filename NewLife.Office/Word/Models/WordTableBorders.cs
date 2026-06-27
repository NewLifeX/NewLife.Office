namespace NewLife.Office;

/// <summary>Word 四边边框集合，用于表格、单元格或段落边框设置</summary>
/// <remarks>
/// null 表示对应边不设置（使用默认或父级样式）。
/// <example>
/// <code>
/// var borders = new WordTableBorders
/// {
///     Top    = new WordBorder { Style = WordBorderStyle.Single, Width = 8 },
///     Bottom = new WordBorder { Style = WordBorderStyle.Single, Width = 8 },
///     Left   = new WordBorder { Style = WordBorderStyle.None },
///     Right  = new WordBorder { Style = WordBorderStyle.None },
///     InsideH = new WordBorder { Style = WordBorderStyle.Dotted },
///     InsideV = new WordBorder { Style = WordBorderStyle.None },
/// };
/// </code>
/// </example>
/// </remarks>
public class WordTableBorders
{
    #region 属性
    /// <summary>上边框</summary>
    public WordBorder? Top { get; set; }

    /// <summary>下边框</summary>
    public WordBorder? Bottom { get; set; }

    /// <summary>左边框</summary>
    public WordBorder? Left { get; set; }

    /// <summary>右边框</summary>
    public WordBorder? Right { get; set; }

    /// <summary>表格内部水平分隔线</summary>
    public WordBorder? InsideH { get; set; }

    /// <summary>表格内部垂直分隔线</summary>
    public WordBorder? InsideV { get; set; }
    #endregion

    #region 工厂方法
    /// <summary>创建四边统一边框</summary>
    /// <param name="style">线型</param>
    /// <param name="color">颜色（hex）</param>
    /// <param name="width">粗细（八分之一磅）</param>
    public static WordTableBorders All(WordBorderStyle style, String? color = null, Int32 width = 4)
    {
        var b = new WordBorder { Style = style, Color = color, Width = width };
        return new WordTableBorders { Top = b, Bottom = b, Left = b, Right = b, InsideH = b, InsideV = b };
    }

    /// <summary>创建仅外框线（无内部分隔线）</summary>
    public static WordTableBorders OutlineOnly(WordBorderStyle style = WordBorderStyle.Single, String? color = null, Int32 width = 4)
    {
        var b = new WordBorder { Style = style, Color = color, Width = width };
        var none = new WordBorder { Style = WordBorderStyle.None };
        return new WordTableBorders { Top = b, Bottom = b, Left = b, Right = b, InsideH = none, InsideV = none };
    }
    #endregion
}
