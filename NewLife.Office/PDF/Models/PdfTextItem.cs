using System.Text;

namespace NewLife.Office;

/// <summary>PDF 文本项，包含文本内容和近似坐标</summary>
/// <remarks>
/// 坐标系以页面左下角为原点，单位为 PDF 用户空间单位（通常约等于磅/pt）。
/// </remarks>
public class PdfTextItem
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>近似 X 坐标（PDF 用户空间单位）</summary>
    public Single X { get; set; }

    /// <summary>近似 Y 坐标（PDF 用户空间单位）</summary>
    public Single Y { get; set; }

    /// <summary>字体大小（通过 Tf 操作符获取，0 表示未知）</summary>
    public Single FontSize { get; set; }
    #endregion
}