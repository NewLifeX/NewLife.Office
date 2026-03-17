using System.Data;

namespace NewLife.Office;

/// <summary>Excel 静态便捷入口，提供与 MiniExcel 风格对齐的最简洁 API</summary>
/// <remarks>
/// 无需实例化，所有方法均为静态调用。
/// 适合一行代码完成导入/导出/模板填充的快捷场景；
/// 需要精细控制样式、多 Sheet、页面设置等能力时，请直接使用 <see cref="ExcelWriter"/>/<see cref="ExcelReader"/>/<see cref="ExcelTemplate"/>。
/// </remarks>
public static class ExcelHelper
{
    #region 写入（SaveAs）

    /// <summary>将对象集合导出为 xlsx 文件</summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="path">目标文件路径</param>
    /// <param name="data">对象集合</param>
    /// <param name="sheetName">工作表名称，默认 "Sheet1"</param>
    /// <param name="headerStyle">可选表头样式</param>
    public static void SaveAs<T>(String path, IEnumerable<T> data, String sheetName = "Sheet1", CellStyle? headerStyle = null) where T : class
    {
        if (path.IsNullOrEmpty()) throw new ArgumentNullException(nameof(path));
        if (data == null) throw new ArgumentNullException(nameof(data));

        using var writer = new ExcelWriter(path);
        writer.SheetName = sheetName;
        writer.WriteObjects(sheetName, data, headerStyle);
        writer.Save();
    }

    /// <summary>将对象集合导出为 xlsx 后写入流</summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="stream">目标可写流</param>
    /// <param name="data">对象集合</param>
    /// <param name="sheetName">工作表名称，默认 "Sheet1"</param>
    /// <param name="headerStyle">可选表头样式</param>
    public static void SaveAs<T>(Stream stream, IEnumerable<T> data, String sheetName = "Sheet1", CellStyle? headerStyle = null) where T : class
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (data == null) throw new ArgumentNullException(nameof(data));

        using var writer = new ExcelWriter(stream);
        writer.SheetName = sheetName;
        writer.WriteObjects(sheetName, data, headerStyle);
        writer.Save();
    }

    /// <summary>将 DataTable 导出为 xlsx 文件</summary>
    /// <param name="path">目标文件路径</param>
    /// <param name="table">DataTable</param>
    /// <param name="sheetName">工作表名称，默认 "Sheet1"</param>
    /// <param name="headerStyle">可选表头样式</param>
    public static void SaveAs(String path, DataTable table, String sheetName = "Sheet1", CellStyle? headerStyle = null)
    {
        if (path.IsNullOrEmpty()) throw new ArgumentNullException(nameof(path));
        if (table == null) throw new ArgumentNullException(nameof(table));

        using var writer = new ExcelWriter(path);
        writer.SheetName = sheetName;
        writer.WriteDataTable(sheetName, table, headerStyle);
        writer.Save();
    }

    /// <summary>将二维数据（行集合）导出为 xlsx 文件</summary>
    /// <param name="path">目标文件路径</param>
    /// <param name="rows">行集合，每行为 Object?[]</param>
    /// <param name="sheetName">工作表名称，默认 "Sheet1"</param>
    public static void SaveAs(String path, IEnumerable<Object?[]> rows, String sheetName = "Sheet1")
    {
        if (path.IsNullOrEmpty()) throw new ArgumentNullException(nameof(path));
        if (rows == null) throw new ArgumentNullException(nameof(rows));

        using var writer = new ExcelWriter(path);
        writer.WriteRows(sheetName, rows);
        writer.Save();
    }

    #endregion

    #region 读取（Query）

    /// <summary>读取 xlsx 文件并映射为强类型集合</summary>
    /// <typeparam name="T">目标类型（需有无参构造函数）</typeparam>
    /// <param name="path">xlsx 文件路径</param>
    /// <param name="sheetName">工作表名称（可空，空时读取第一个）</param>
    /// <returns>对象集合（第一行作为列名映射）</returns>
    public static IEnumerable<T> Query<T>(String path, String? sheetName = null) where T : new()
    {
        if (path.IsNullOrEmpty()) throw new ArgumentNullException(nameof(path));

        using var reader = new ExcelReader(path);
        foreach (var item in reader.ReadObjects<T>(sheetName))
        {
            yield return item;
        }
    }

    /// <summary>读取 xlsx 文件的原始行数据</summary>
    /// <param name="path">xlsx 文件路径</param>
    /// <param name="sheetName">工作表名称（可空，空时读取第一个）</param>
    /// <returns>行集合，每行为 Object?[]</returns>
    public static IEnumerable<Object?[]> Query(String path, String? sheetName = null)
    {
        if (path.IsNullOrEmpty()) throw new ArgumentNullException(nameof(path));

        using var reader = new ExcelReader(path);
        foreach (var row in reader.ReadRows(sheetName))
        {
            yield return row;
        }
    }

    /// <summary>读取 xlsx 文件到 DataTable</summary>
    /// <param name="path">xlsx 文件路径</param>
    /// <param name="sheetName">工作表名称（可空，空时读取第一个）</param>
    /// <returns>DataTable（第一行作为列名）</returns>
    public static DataTable QueryAsDataTable(String path, String? sheetName = null)
    {
        if (path.IsNullOrEmpty()) throw new ArgumentNullException(nameof(path));

        using var reader = new ExcelReader(path);
        return reader.ReadDataTable(sheetName);
    }

    #endregion

    #region 模板填充（SaveByTemplate）

    /// <summary>基于 xlsx 模板填充占位符并输出到新文件</summary>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="templatePath">模板 xlsx 文件路径</param>
    /// <param name="data">占位符数据（键为占位符名称，值为替换内容）</param>
    public static void SaveByTemplate(String outputPath, String templatePath, IDictionary<String, Object> data)
    {
        if (outputPath.IsNullOrEmpty()) throw new ArgumentNullException(nameof(outputPath));
        if (templatePath.IsNullOrEmpty()) throw new ArgumentNullException(nameof(templatePath));
        if (data == null) throw new ArgumentNullException(nameof(data));

        var tpl = new ExcelTemplate(templatePath);
        tpl.Fill(outputPath, data);
    }

    /// <summary>基于 xlsx 模板填充对象属性并输出到新文件（自动从对象属性构建字典）</summary>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="templatePath">模板 xlsx 文件路径</param>
    /// <param name="dataObject">数据对象，属性名即占位符键名</param>
    public static void SaveByTemplate(String outputPath, String templatePath, Object dataObject)
    {
        if (outputPath.IsNullOrEmpty()) throw new ArgumentNullException(nameof(outputPath));
        if (templatePath.IsNullOrEmpty()) throw new ArgumentNullException(nameof(templatePath));
        if (dataObject == null) throw new ArgumentNullException(nameof(dataObject));

        var dict = new Dictionary<String, Object>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in dataObject.GetType().GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
        {
            if (prop.CanRead)
            {
                var val = prop.GetValue(dataObject);
                if (val != null) dict[prop.Name] = val;
            }
        }

        var tpl = new ExcelTemplate(templatePath);
        tpl.Fill(outputPath, dict);
    }

    #endregion
}
