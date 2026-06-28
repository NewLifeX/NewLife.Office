namespace NewLife.Office;

/// <summary>条件格式类型</summary>
public enum ExcelConditionalFormatType
{
    /// <summary>大于</summary>
    GreaterThan = 0,

    /// <summary>小于</summary>
    LessThan,

    /// <summary>等于</summary>
    Equal,

    /// <summary>介于</summary>
    Between,

    /// <summary>不等于</summary>
    NotEqual,

    /// <summary>不介于</summary>
    NotBetween,

    /// <summary>数据条</summary>
    DataBar,

    /// <summary>色阶</summary>
    ColorScale,

    /// <summary>图标集（3/4/5 级箭头、旗标、星级等）</summary>
    IconSet,

    /// <summary>自定义公式（任意 Excel 公式表达式）</summary>
    Expression,
}
