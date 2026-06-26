namespace NewLife.Office;

/// <summary>数据验证信息</summary>
public class ValidationInfo
{
    #region 属性
    /// <summary>验证范围（如 "A2:A100"）</summary>
    public String CellRange { get; set; } = String.Empty;

    /// <summary>下拉选项列表（列表验证时使用）</summary>
    public String[]? Items { get; set; }

    /// <summary>验证类型：whole/decimal/date/time/textLength</summary>
    public String? ValidationType { get; set; }

    /// <summary>运算符：between/notBetween/equal/notEqual/greaterThan/lessThan 等</summary>
    public String? Operator { get; set; }

    /// <summary>公式1（最小值或比较值）</summary>
    public String? Formula1 { get; set; }

    /// <summary>公式2（最大值，仅 between/notBetween 使用）</summary>
    public String? Formula2 { get; set; }
    #endregion
}
