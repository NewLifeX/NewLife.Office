namespace NewLife.Office.Ppt;

/// <summary>PPT 嵌入图表</summary>
public class Chart
{
    #region 属性
    /// <summary>图表类型（bar/line/pie/area/scatter）</summary>
    public String ChartType { get; set; } = "bar";

    /// <summary>图表标题，null 表示不显示</summary>
    public String? Title { get; set; }

    /// <summary>分类轴标签</summary>
    public String[] Categories { get; set; } = [];

    /// <summary>系列集合</summary>
    public List<ChartSeries> Series { get; } = [];

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; } = 6000000;

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; } = 4000000;

    /// <summary>图表关系ID（由写入器内部设置）</summary>
    public String RelId { get; set; } = String.Empty;

    /// <summary>图表文件编号（由写入器内部设置）</summary>
    internal Int32 ChartNumber { get; set; }

    /// <summary>数值轴最小值（仅数值轴，分类轴无效），null 表示自动</summary>
    public Double? AxisMinValue { get; set; }

    /// <summary>数值轴最大值（仅数值轴，分类轴无效），null 表示自动</summary>
    public Double? AxisMaxValue { get; set; }

    /// <summary>图表样式索引（1-48，null 不写 c:style，S19）</summary>
    public Int32? StyleIndex { get; set; }

    /// <summary>图例位置（b/r/l/t，null 默认底部）</summary>
    public String? LegendPosition { get; set; }
    #endregion
}
