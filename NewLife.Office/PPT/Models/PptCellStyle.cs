namespace NewLife.Office;

/// <summary>PPT 表格单元格样式（S10-02）</summary>
public class PptCellStyle
{
    #region 属性
    /// <summary>单元格背景色（16进制 RGB），null 表示表格默认色</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>字体颜色（16进制 RGB），null 表示继承</summary>
    public String? FontColor { get; set; }

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>字号（磅），0 表示继承</summary>
    public Int32 FontSize { get; set; }
    #endregion
}
