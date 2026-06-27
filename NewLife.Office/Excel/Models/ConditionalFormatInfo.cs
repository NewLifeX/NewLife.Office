namespace NewLife.Office;

/// <summary>条件格式信息</summary>
public class ExcelConditionalFormat
{
    #region 属性
    /// <summary>应用范围（如 "A1:A100"）</summary>
    public String Range { get; set; } = String.Empty;

    /// <summary>条件类型</summary>
    public ExcelConditionalFormatType Type { get; set; }

    /// <summary>条件值（如 "10000"）</summary>
    public String? Value { get; set; }

    /// <summary>第二条件值（仅 Between 类型使用）</summary>
    public String? Value2 { get; set; }

    /// <summary>颜色（RGB十六进制）</summary>
    public String? Color { get; set; }
    #endregion
}
