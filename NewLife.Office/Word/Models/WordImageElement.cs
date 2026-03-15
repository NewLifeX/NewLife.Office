namespace NewLife.Office;

/// <summary>图片元素</summary>
public class WordImageElement
{
    #region 属性
    /// <summary>图片数据</summary>
    public Byte[] ImageData { get; set; } = [];

    /// <summary>扩展名（png/jpg）</summary>
    public String Extension { get; set; } = "png";

    /// <summary>宽度（EMU，914400 = 1英寸）</summary>
    public Int64 WidthEmu { get; set; } = 3600000;

    /// <summary>高度（EMU）</summary>
    public Int64 HeightEmu { get; set; } = 2700000;

    /// <summary>关系ID</summary>
    public String RelId { get; set; } = String.Empty;
    #endregion
}
