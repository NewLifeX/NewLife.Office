namespace NewLife.Office;

/// <summary>PDF 描边笔（线条/边框样式）</summary>
/// <remarks>
/// 描述 PDF 绘图操作中的线条样式，包括颜色、粗细、端点和虚线图案。
/// <example>
/// <code>
/// var pen = new PdfPen { Color = OfficeColor.Red, Width = 2f };
/// var dashedPen = new PdfPen { Width = 1f, DashPattern = [4f, 2f] };  // 4pt实线+2pt空白
/// </code>
/// </example>
/// </remarks>
public class PdfPen
{
    #region 属性
    /// <summary>线条颜色，null 表示使用黑色</summary>
    public OfficeColor? Color { get; set; }

    /// <summary>线条粗细（磅），默认 1pt</summary>
    public Single Width { get; set; } = 1f;

    /// <summary>线端帽样式</summary>
    public PdfLineCap LineCap { get; set; } = PdfLineCap.Butt;

    /// <summary>线段连接样式</summary>
    public PdfLineJoin LineJoin { get; set; } = PdfLineJoin.Miter;

    /// <summary>虚线图案（单位磅，交替表示实线/空白长度），null = 实线</summary>
    public Single[]? DashPattern { get; set; }

    /// <summary>虚线起始偏移（磅）</summary>
    public Single DashOffset { get; set; }
    #endregion

    #region 预定义
    /// <summary>细实线（0.5pt 黑色）</summary>
    public static readonly PdfPen Hairline = new() { Width = 0.5f };

    /// <summary>标准实线（1pt 黑色）</summary>
    public static readonly PdfPen Solid = new() { Width = 1f };

    /// <summary>粗实线（2pt 黑色）</summary>
    public static readonly PdfPen Thick = new() { Width = 2f };

    /// <summary>标准虚线（1pt 黑色，4pt实/2pt空）</summary>
    public static readonly PdfPen Dashed = new() { Width = 1f, DashPattern = [4f, 2f] };
    #endregion
}

/// <summary>PDF 线端帽样式</summary>
public enum PdfLineCap
{
    /// <summary>平头端点（截止于线段终点）</summary>
    Butt = 0,

    /// <summary>圆头端点</summary>
    Round = 1,

    /// <summary>方头端点（超出终点半个线宽）</summary>
    Square = 2,
}

/// <summary>PDF 线段连接样式</summary>
public enum PdfLineJoin
{
    /// <summary>尖角连接</summary>
    Miter = 0,

    /// <summary>圆角连接</summary>
    Round = 1,

    /// <summary>斜切连接</summary>
    Bevel = 2,
}
