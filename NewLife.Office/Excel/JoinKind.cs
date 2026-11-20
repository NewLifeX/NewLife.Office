namespace NewLife.Office.Excel;

/// <summary>关联类型</summary>
/// <remarks>供 <see cref="ExcelDataFrame.Merge(ExcelDataFrame, String, String, JoinKind, String)"/> 选择关联方式。</remarks>
public enum JoinKind
{
    /// <summary>仅保留两侧匹配的行</summary>
    Inner,

    /// <summary>保留全部左侧行</summary>
    Left,

    /// <summary>保留全部右侧行</summary>
    Right,

    /// <summary>保留两侧所有行</summary>
    Full,
}
