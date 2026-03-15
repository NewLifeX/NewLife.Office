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
    #endregion
}
