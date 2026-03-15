namespace NewLife.Office;

/// <summary>表格样式配置</summary>
public class WordTableStyle
{
    #region 属性
    /// <summary>边框颜色（16进制 RGB，默认黑色）</summary>
    public String BorderColor { get; set; } = "000000";

    /// <summary>边框线宽（pt8，默认4=0.5pt）</summary>
    public Int32 BorderSize { get; set; } = 4;

    /// <summary>表头行背景色（16进制 RGB，null=不设置）</summary>
    public String? HeaderBgColor { get; set; }

    /// <summary>表头行字体加粗</summary>
    public Boolean HeaderBold { get; set; } = true;

    /// <summary>斑马纹颜色（奇数行背景色，null=不设置）</summary>
    public String? StripeColor { get; set; }

    /// <summary>列宽列表（twips，null=自动均分）</summary>
    public Int32[]? ColumnWidths { get; set; }
    #endregion
}
