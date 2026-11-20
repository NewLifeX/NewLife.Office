namespace NewLife.Office.Excel;

/// <summary>分组汇总函数</summary>
/// <remarks>用于 <see cref="AggregateSpec"/>、<see cref="ExcelDataFrame.Pivot(String, String, String, AggregateFunction)"/> 等聚合入口。</remarks>
public enum AggregateFunction
{
    /// <summary>非空值数量；<see cref="AggregateSpec.CountRows(String)"/> 则统计组内行数</summary>
    Count,

    /// <summary>数值求和</summary>
    Sum,

    /// <summary>数值平均值</summary>
    Average,

    /// <summary>最小值</summary>
    Min,

    /// <summary>最大值</summary>
    Max,
}
