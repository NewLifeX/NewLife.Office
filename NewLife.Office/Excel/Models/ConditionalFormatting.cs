namespace NewLife.Office.Excel;

/// <summary>条件格式信息</summary>
public class ConditionalFormatting
{
    #region 属性
    /// <summary>应用范围（如 "A1:A100"）</summary>
    public String Range { get; set; } = String.Empty;

    /// <summary>条件类型</summary>
    public ConditionalFormatValues Type { get; set; }

    /// <summary>条件值（如 "10000"）</summary>
    public String? Value { get; set; }

    /// <summary>第二条件值（仅 Between 类型使用）</summary>
    public String? Value2 { get; set; }

    /// <summary>颜色（RGB十六进制）</summary>
    public String? Color { get; set; }

    /// <summary>字体颜色（RGB十六进制，dxf 字体样式）</summary>
    public String? FontColor { get; set; }

    /// <summary>边框颜色（RGB十六进制，dxf 边框样式）</summary>
    public String? BorderColor { get; set; }

    /// <summary>是否加粗（dxf 字体样式）</summary>
    public Boolean IsBold { get; set; }

    /// <summary>图标集类型（仅 IconSet 类型，如 "3Arrows"/"3TrafficLights1"）</summary>
    public String? IconSetType { get; set; }

    /// <summary>自定义公式（仅 Expression 类型）</summary>
    public String? Formula { get; set; }
    #endregion
}
