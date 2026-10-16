namespace NewLife.Office.Excel;

/// <summary>Excel 数据透视表定义（Reader 回读模型，E17）</summary>
/// <remarks>
/// 由 <see cref="ExcelReader.ReadPivotTables"/> 返回，<see cref="Fields"/> 保留
/// 行列/数据/筛选标记与汇总方式，字段名从透视缓存（cacheFields）还原。
/// </remarks>
public class PivotTableInfo
{
    #region 属性
    /// <summary>透视表名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>透视表位置（单元格区域，如 A3:C10）</summary>
    public String Location { get; set; } = String.Empty;

    /// <summary>透视缓存 ID</summary>
    public Int32 CacheId { get; set; }

    /// <summary>字段集合（含行列/数据/筛选标记；数据字段带汇总方式与显示名称）</summary>
    public List<PivotField> Fields { get; } = [];
    #endregion
}
