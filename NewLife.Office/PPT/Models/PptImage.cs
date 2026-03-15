namespace NewLife.Office;

/// <summary>PPT 幻灯片图片元素</summary>
public class PptImage
{
    #region 属性
    /// <summary>图片字节数据</summary>
    public Byte[] Data { get; set; } = [];

    /// <summary>扩展名（png/jpg）</summary>
    public String Extension { get; set; } = "png";

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; } = 3000000;

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; } = 2000000;

    /// <summary>关系ID（内部用）</summary>
    public String RelId { get; set; } = String.Empty;
    #endregion
}
