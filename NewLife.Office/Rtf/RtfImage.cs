using System;

namespace NewLife.Office.Rtf;

/// <summary>RTF 图片</summary>
/// <remarks>
/// 表示 RTF 文档中 \pict 组嵌入的图片，支持 PNG、JPEG、EMF 和 WMF 格式。
/// </remarks>
public sealed class RtfImage
{
    #region 属性
    /// <summary>图片格式（png/jpg/emf/wmf）</summary>
    public String Format { get; set; } = "png";

    /// <summary>图片原始字节数据</summary>
    public Byte[] Data { get; set; } = new Byte[0];

    /// <summary>图片宽度（twips，1 英寸=1440 twips），读取时来自 \picw；写入时为显示宽度</summary>
    public Int32 Width { get; set; }

    /// <summary>图片高度（twips），读取时来自 \pich；写入时为显示高度</summary>
    public Int32 Height { get; set; }
    #endregion

    /// <inheritdoc/>
    public override String ToString() => $"[{Format}] {Data.Length} bytes {Width}x{Height} twips";
}
