namespace NewLife.Office.Word;

/// <summary>Word 图表系列（数据列）</summary>
public class WordChartSeries
{
    /// <summary>系列名称（图表图例）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>系列数值（与 Categories 长度一致）</summary>
    public Double[] Values { get; set; } = [];
}

/// <summary>Word 图表（W13，对标 Open XML SDK/Aspose.Words）</summary>
/// <remarks>
/// 表示插入到 Word 文档中的 DrawingML 图表（柱状/折线/饼图）。
/// 生成 chartSpace XML + 嵌入式 Excel 数据包，通过 drawing 部件挂载到段落。
/// </remarks>
public class WordChart
{
    /// <summary>图表类型：bar/column/line/pie</summary>
    public String Type { get; set; } = "column";

    /// <summary>图表标题</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>分类轴类别（如月份/产品名）</summary>
    public String[] Categories { get; set; } = [];

    /// <summary>数据系列列表</summary>
    public List<WordChartSeries> Series { get; set; } = [];

    /// <summary>宽度（厘米）</summary>
    public Double WidthCm { get; set; } = 12;

    /// <summary>高度（厘米）</summary>
    public Double HeightCm { get; set; } = 7.5;

    /// <summary>快速创建单系列柱状图</summary>
    /// <param name="title">图表标题</param>
    /// <param name="categories">分类</param>
    /// <param name="seriesName">系列名称</param>
    /// <param name="values">数值</param>
    /// <returns>图表对象</returns>
    public static WordChart Column(String title, String[] categories, String seriesName, Double[] values)
    {
        return new WordChart
        {
            Type = "column",
            Title = title,
            Categories = categories,
            Series = [new WordChartSeries { Name = seriesName, Values = values }],
        };
    }

    /// <summary>快速创建折线图</summary>
    /// <param name="title">图表标题</param>
    /// <param name="categories">分类</param>
    /// <param name="seriesName">系列名称</param>
    /// <param name="values">数值</param>
    /// <returns>图表对象</returns>
    public static WordChart Line(String title, String[] categories, String seriesName, Double[] values)
    {
        return new WordChart
        {
            Type = "line",
            Title = title,
            Categories = categories,
            Series = [new WordChartSeries { Name = seriesName, Values = values }],
        };
    }

    /// <summary>快速创建饼图（单系列）</summary>
    /// <param name="title">图表标题</param>
    /// <param name="categories">分类</param>
    /// <param name="values">数值</param>
    /// <returns>图表对象</returns>
    public static WordChart Pie(String title, String[] categories, Double[] values)
    {
        return new WordChart
        {
            Type = "pie",
            Title = title,
            Categories = categories,
            Series = [new WordChartSeries { Name = title, Values = values }],
        };
    }
}
