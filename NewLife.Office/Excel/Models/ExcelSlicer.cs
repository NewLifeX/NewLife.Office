namespace NewLife.Office.Excel;

/// <summary>Excel 切片器（Slicer，M28，对标 EPPlus/Aspose.Cells）</summary>
/// <remarks>
/// 切片器是用于快速筛选结构化表格数据的可视化组件（Excel 2010+）。
/// OOXML 实现包含 slicerCache 部件（数据缓存）与 slicer 部件（外观），
/// 通过工作表的 x14:slicerList 扩展引用挂载。
/// </remarks>
public class ExcelSlicer
{
    /// <summary>切片器名称（XML name 属性，默认自动生成）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>显示标题（caption，默认取列名）</summary>
    public String Caption { get; set; } = String.Empty;

    /// <summary>关联的结构化表格名称</summary>
    public String TableName { get; set; } = String.Empty;

    /// <summary>筛选字段列名（表格中的列名）</summary>
    public String ColumnName { get; set; } = String.Empty;

    /// <summary>所属工作表名称</summary>
    public String Sheet { get; set; } = String.Empty;

    /// <summary>切片器样式（默认 SlicerStyleLight1）</summary>
    public String Style { get; set; } = "SlicerStyleLight1";

    /// <summary>关联的切片器缓存名称（Save 时生成，读取时还原）</summary>
    public String? CacheName { get; set; }
}
