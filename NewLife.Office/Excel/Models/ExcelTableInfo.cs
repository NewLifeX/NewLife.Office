namespace NewLife.Office;

/// <summary>Excel 结构化表格（OOXML table 元素）</summary>
/// <remarks>
/// 对应 Open XML 规范中的 xl/tables/tableN.xml 部件及 &lt;tableParts&gt; 引用。
/// 通过 ExcelWriter.AddTable 写入，Reader.ReadTables() 读取。
/// <example>
/// <code>
/// writer.AddTable("A1:E10", "销售明细", "TableStyleMedium9");
/// // 或带列名
/// writer.AddTable("A1:E10", "销售明细", "TableStyleMedium9",
///     new[] { "日期", "产品", "数量", "单价", "金额" });
/// </code>
/// </example>
/// </remarks>
public class ExcelTableInfo
{
    #region 属性
    /// <summary>表格范围（Excel 记法，如 "A1:E10"）</summary>
    public String Range { get; set; } = String.Empty;

    /// <summary>表格名称（XML 中的 name/displayName，同时用作命名范围）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>表格样式名称（如 "TableStyleMedium9"，null 表示使用默认）</summary>
    public String? StyleName { get; set; } = "TableStyleMedium9";

    /// <summary>列名集合（与范围内的列一一对应）；null 时按 Column1/Column2... 自动生成</summary>
    public String[]? ColumnNames { get; set; }

    /// <summary>是否显示表头行（默认 true）</summary>
    public Boolean ShowHeader { get; set; } = true;

    /// <summary>是否显示行带状（斑马条纹），默认 true</summary>
    public Boolean ShowRowStripes { get; set; } = true;

    /// <summary>是否显示列带状</summary>
    public Boolean ShowColumnStripes { get; set; }

    /// <summary>是否高亮第一列</summary>
    public Boolean ShowFirstColumn { get; set; }

    /// <summary>是否高亮最后列</summary>
    public Boolean ShowLastColumn { get; set; }

    /// <summary>是否显示筛选按钮（默认 true）</summary>
    public Boolean ShowFilterButton { get; set; } = true;
    #endregion
}
