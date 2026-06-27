namespace NewLife.Office;

/// <summary>页面设置</summary>
public class WordPageSettings
{
    #region 属性
    /// <summary>页面宽度（twips，1440 twips = 1英寸）</summary>
    public Int32 PageWidth { get; set; } = 11906; // A4: 210mm

    /// <summary>页面高度（twips）</summary>
    public Int32 PageHeight { get; set; } = 16838; // A4: 297mm

    /// <summary>上边距（twips）</summary>
    public Int32 MarginTop { get; set; } = 1440;

    /// <summary>下边距（twips）</summary>
    public Int32 MarginBottom { get; set; } = 1440;

    /// <summary>左边距（twips）</summary>
    public Int32 MarginLeft { get; set; } = 1800;

    /// <summary>右边距（twips）</summary>
    public Int32 MarginRight { get; set; } = 1800;

    /// <summary>横向</summary>
    public Boolean Landscape { get; set; }

    /// <summary>页眉文本</summary>
    public String? HeaderText { get; set; }

    /// <summary>页脚文本</summary>
    public String? FooterText { get; set; }

    /// <summary>水印文字（null 表示无水印）</summary>
    public String? WatermarkText { get; set; }

    /// <summary>分栏数量（1 表示不分栏，默认 1）</summary>
    public Int32 ColumnCount { get; set; } = 1;

    /// <summary>分栏间距（twips，默认 720 = 0.5 英寸）</summary>
    public Int32 ColumnSpacing { get; set; } = 720;

    /// <summary>页面边框设置（null 表示无边框）</summary>
    public WordPageBorder? PageBorder { get; set; }
    #endregion
}

/// <summary>页面边框设置</summary>
public class WordPageBorder
{
    /// <summary>上边框样式（single/double/dotted/dash/thick/wave 等）</summary>
    public String? Top { get; set; }

    /// <summary>下边框样式</summary>
    public String? Bottom { get; set; }

    /// <summary>左边框样式</summary>
    public String? Left { get; set; }

    /// <summary>右边框样式</summary>
    public String? Right { get; set; }

    /// <summary>边框颜色（16进制RGB，如 "FF0000"），null 表示自动</summary>
    public String? Color { get; set; }

    /// <summary>边框宽度（缇，1/8点），默认 4 = 0.5pt</summary>
    public Int32 Size { get; set; } = 4;

    /// <summary>边框距页面边缘的距离（缇），默认 24</summary>
    public Int32 Space { get; set; } = 24;

    /// <summary>边框距页面文字的距离（缇），默认 24</summary>
    public Int32 OffsetFrom { get; set; } = 24; // 0=text, 1=page edge
}
