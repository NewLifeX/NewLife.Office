namespace NewLife.Office.Excel;

/// <summary>单元格批注（注释）</summary>
/// <remarks>
/// 包含批注文本、作者以及批注框的尺寸/位置/可见性（源自 xl/drawings/vmlDrawingN.vml）。
/// </remarks>
public class Comment
{
    /// <summary>批注文本</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>批注作者</summary>
    public String Author { get; set; } = String.Empty;

    /// <summary>批注框宽度（磅，默认 108pt）</summary>
    public Double Width { get; set; } = 108;

    /// <summary>批注框高度（磅，默认 59.25pt）</summary>
    public Double Height { get; set; } = 59.25;

    /// <summary>是否可见（默认隐藏，鼠标悬停时显示）</summary>
    public Boolean Visible { get; set; }

    /// <summary>批注框左边距（磅，相对单元格左上角）</summary>
    public Double Left { get; set; } = 59.25;

    /// <summary>批注框上边距（磅，相对单元格左上角）</summary>
    public Double Top { get; set; } = 1.5;
}
