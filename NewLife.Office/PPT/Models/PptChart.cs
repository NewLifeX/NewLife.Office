namespace NewLife.Office;

/// <summary>PPT 嵌入图表</summary>
public class PptChart
{
    #region 属性
    /// <summary>图表类型（bar/line/pie/area/scatter）</summary>
    public String ChartType { get; set; } = "bar";

    /// <summary>图表标题，null 表示不显示</summary>
    public String? Title { get; set; }

    /// <summary>分类轴标签</summary>
    public String[] Categories { get; set; } = [];

    /// <summary>系列集合</summary>
    public List<PptChartSeries> Series { get; } = [];

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
    #endregion
}
