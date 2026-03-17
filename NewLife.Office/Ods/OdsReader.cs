using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml;

namespace NewLife.Office.Ods;

/// <summary>ODS 电子表格读取器</summary>
/// <remarks>
/// 读取 OpenDocument Spreadsheet（.ods）格式文件，提取工作表名称、单元格数据及合并区域。
/// ODS 是基于 ZIP 的 XML 格式，核心内容在 content.xml 中。
/// </remarks>
public sealed class OdsReader
{
    #region 常量
    private const String NsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private const String NsTable = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private const String NsText = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private const String NsStyle = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    private const String NsFo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    #endregion

    #region 方法 — 读取
    /// <summary>从文件路径读取所有工作表数据</summary>
    /// <param name="path">ODS 文件路径</param>
    /// <returns>工作表列表</returns>
    public static List<OdsSheet> ReadFile(String path)
    {
        using var fs = File.OpenRead(path);
        return Read(fs);
    }

    /// <summary>从流读取所有工作表数据</summary>
    /// <param name="stream">ODS 输入流</param>
    /// <returns>工作表列表</returns>
    public static List<OdsSheet> Read(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var entry = zip.GetEntry("content.xml");
        if (entry == null) return [];

        using var contentStream = entry.Open();
        return ParseContentXml(contentStream);
    }

    /// <summary>从文件路径读取第一张工作表的数据行</summary>
    /// <param name="path">ODS 文件路径</param>
    /// <returns>行列表，每行为字符串数组</returns>
    public static List<String[]> ReadRows(String path)
    {
        var sheets = ReadFile(path);
        return sheets.Count > 0 ? sheets[0].Rows : [];
    }

    /// <summary>从流读取第一张工作表的数据行</summary>
    /// <param name="stream">ODS 输入流</param>
    /// <returns>行列表，每行为字符串数组</returns>
    public static List<String[]> ReadRows(Stream stream)
    {
        var sheets = Read(stream);
        return sheets.Count > 0 ? sheets[0].Rows : [];
    }

    /// <summary>从文件读取第一张工作表并映射为对象集合</summary>
    /// <typeparam name="T">目标类型，需要有公共无参构造函数</typeparam>
    /// <param name="path">ODS 文件路径</param>
    /// <param name="sheetName">工作表名称（空则取第一张）</param>
    /// <returns>对象集合</returns>
    public static IEnumerable<T> ReadObjects<T>(String path, String? sheetName = null) where T : new()
    {
        var sheets = ReadFile(path);
        var sheet = FindSheet(sheets, sheetName);
        return sheet == null ? [] : MapObjects<T>(sheet.Rows);
    }

    /// <summary>从流读取工作表并映射为对象集合</summary>
    /// <typeparam name="T">目标类型</typeparam>
    /// <param name="stream">ODS 流</param>
    /// <param name="sheetName">工作表名称（空则取第一张）</param>
    /// <returns>对象集合</returns>
    public static IEnumerable<T> ReadObjects<T>(Stream stream, String? sheetName = null) where T : new()
    {
        var sheets = Read(stream);
        var sheet = FindSheet(sheets, sheetName);
        return sheet == null ? [] : MapObjects<T>(sheet.Rows);
    }

    /// <summary>从文件读取工作表为 DataTable</summary>
    /// <param name="path">ODS 文件路径</param>
    /// <param name="sheetName">工作表名称（空则取第一张）</param>
    /// <returns>DataTable（第一行作为列名）</returns>
    public static DataTable ReadDataTable(String path, String? sheetName = null)
    {
        var sheets = ReadFile(path);
        var sheet = FindSheet(sheets, sheetName);
        return sheet == null ? new DataTable() : BuildDataTable(sheet.Rows);
    }

    /// <summary>从流读取工作表为 DataTable</summary>
    /// <param name="stream">ODS 流</param>
    /// <param name="sheetName">工作表名称（空则取第一张）</param>
    /// <returns>DataTable</returns>
    public static DataTable ReadDataTable(Stream stream, String? sheetName = null)
    {
        var sheets = Read(stream);
        var sheet = FindSheet(sheets, sheetName);
        return sheet == null ? new DataTable() : BuildDataTable(sheet.Rows);
    }
    #endregion

    #region XML 解析
    private static List<OdsSheet> ParseContentXml(Stream xmlStream)
    {
        var result = new List<OdsSheet>();
        var settings = new XmlReaderSettings { IgnoreWhitespace = false, IgnoreComments = true };
        using var reader = XmlReader.Create(xmlStream, settings);

        OdsSheet? currentSheet = null;
        List<String>? currentRow = null;
        StringBuilder? cellText = null;
        var inTextP = false;
        var cellRepeat = 1;
        var rowRepeat = 1;
        var colSpan = 1;
        var rowSpan = 1;
        var currentRowIndex = 0;
        var currentColIndex = 0;

        // 样式解析状态
        var styleMap = new Dictionary<String, OdsCellStyle>();
        var inAutoStyles = false;
        String? currentStyleName = null;
        OdsCellStyle? currentCellStyle = null;
        String? currentCellStyleName = null;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                var ns = reader.NamespaceURI;
                var name = reader.LocalName;

                // 进入 automatic-styles 区段
                if (ns == NsOffice && name == "automatic-styles")
                {
                    inAutoStyles = true;
                    continue;
                }

                // 样式定义：仅处理 family="table-cell"
                if (inAutoStyles && ns == NsStyle && name == "style")
                {
                    var family = reader.GetAttribute("family", NsStyle);
                    if (family == "table-cell")
                    {
                        currentStyleName = reader.GetAttribute("name", NsStyle);
                        currentCellStyle = new OdsCellStyle();
                    }
                    continue;
                }

                // 文本属性：字体粗体/斜体/大小/颜色
                if (inAutoStyles && currentCellStyle != null && ns == NsStyle && name == "text-properties")
                {
                    var weight = reader.GetAttribute("font-weight", NsFo);
                    if (weight == "bold") currentCellStyle.FontBold = true;
                    var fontStyle = reader.GetAttribute("font-style", NsFo);
                    if (fontStyle == "italic") currentCellStyle.FontItalic = true;
                    var size = reader.GetAttribute("font-size", NsFo);
                    if (size != null) currentCellStyle.FontSize = ParsePtValue(size);
                    var color = reader.GetAttribute("color", NsFo);
                    if (!String.IsNullOrEmpty(color)) currentCellStyle.FontColor = color;
                    continue;
                }

                // 单元格属性：背景色
                if (inAutoStyles && currentCellStyle != null && ns == NsStyle && name == "table-cell-properties")
                {
                    var bg = reader.GetAttribute("background-color", NsFo);
                    if (!String.IsNullOrEmpty(bg) && bg != "transparent")
                        currentCellStyle.BackgroundColor = bg;
                    continue;
                }

                // 段落属性：水平对齐
                if (inAutoStyles && currentCellStyle != null && ns == NsStyle && name == "paragraph-properties")
                {
                    var align = reader.GetAttribute("text-align", NsFo);
                    if (!String.IsNullOrEmpty(align)) currentCellStyle.HAlign = align;
                    continue;
                }

                if (ns == NsTable && name == "table")
                {
                    var sheetName = reader.GetAttribute("name", NsTable) ?? "";
                    currentSheet = new OdsSheet { Name = sheetName };
                    result.Add(currentSheet);
                    currentRowIndex = 0;
                    currentColIndex = 0;
                    continue;
                }

                if (ns == NsTable && name == "table-row")
                {
                    rowRepeat = GetRepeatAttr(reader, NsTable, "number-rows-repeated");
                    currentRow = [];
                    currentColIndex = 0;
                    continue;
                }

                if (ns == NsTable && name == "table-cell")
                {
                    cellRepeat = GetRepeatAttr(reader, NsTable, "number-columns-repeated");
                    colSpan = GetRepeatAttr(reader, NsTable, "number-columns-spanned");
                    rowSpan = GetRepeatAttr(reader, NsTable, "number-rows-spanned");
                    currentCellStyleName = reader.GetAttribute("style-name", NsTable);
                    cellText = new StringBuilder();
                    inTextP = false;
                    continue;
                }

                if (ns == NsText && name == "p")
                {
                    inTextP = true;
                    continue;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                var ns = reader.NamespaceURI;
                var name = reader.LocalName;

                // 离开 automatic-styles
                if (ns == NsOffice && name == "automatic-styles")
                {
                    inAutoStyles = false;
                    currentStyleName = null;
                    currentCellStyle = null;
                    continue;
                }

                // 完成一条样式定义
                if (inAutoStyles && ns == NsStyle && name == "style")
                {
                    if (currentStyleName != null && currentCellStyle != null)
                        styleMap[currentStyleName] = currentCellStyle;
                    currentStyleName = null;
                    currentCellStyle = null;
                    continue;
                }

                if (ns == NsText && name == "p")
                {
                    inTextP = false;
                    continue;
                }

                if (ns == NsTable && name == "table-cell")
                {
                    if (currentRow != null && cellText != null)
                    {
                        var val = cellText.ToString();
                        // 记录合并区域（仅记录一次，不重复 cellRepeat）
                        if (currentSheet != null && (colSpan > 1 || rowSpan > 1))
                        {
                            currentSheet.MergedCells.Add(new OdsMergeRegion
                            {
                                Row = currentRowIndex,
                                Col = currentColIndex,
                                ColSpan = colSpan,
                                RowSpan = rowSpan,
                            });
                        }
                        // 记录单元格样式（仅首列，不展开 repeat）
                        if (currentSheet != null && currentCellStyleName != null &&
                            styleMap.TryGetValue(currentCellStyleName, out var cellStyle))
                        {
                            currentSheet.CellStyles[(currentRowIndex, currentColIndex)] = cellStyle;
                        }
                        for (var i = 0; i < cellRepeat; i++)
                        {
                            currentRow.Add(val);
                        }
                        currentColIndex += cellRepeat;
                    }
                    cellText = null;
                    inTextP = false;
                    colSpan = 1;
                    rowSpan = 1;
                    currentCellStyleName = null;
                    continue;
                }

                if (ns == NsTable && name == "table-row")
                {
                    if (currentSheet != null && currentRow != null)
                    {
                        var trimmedRow = TrimTrailingEmpty(currentRow);
                        if (trimmedRow != null || rowRepeat == 1)
                        {
                            var arr = trimmedRow ?? [];
                            for (var i = 0; i < rowRepeat; i++)
                            {
                                currentSheet.Rows.Add(arr);
                            }
                        }
                    }
                    currentRow = null;
                    currentRowIndex += rowRepeat;
                    rowRepeat = 1;
                    continue;
                }
            }
            else if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.SignificantWhitespace)
            {
                if (inTextP && cellText != null)
                    cellText.Append(reader.Value);
            }
        }

        return result;
    }

    /// <summary>解析 "12pt"/"12.5pt" 格式的字体大小，返回 Single</summary>
    private static Single ParsePtValue(String val)
    {
        if (val.EndsWith("pt", StringComparison.OrdinalIgnoreCase) &&
            Single.TryParse(val[..^2],
                NumberStyles.Float,
                CultureInfo.InvariantCulture, out var f))
            return f;
        return 0f;
    }

    private static Int32 GetRepeatAttr(XmlReader reader, String ns, String localName)
    {
        var val = reader.GetAttribute(localName, ns);
        return val != null && Int32.TryParse(val, out var n) ? n : 1;
    }

    private static String[]? TrimTrailingEmpty(List<String> row)
    {
        var end = row.Count - 1;
        while (end >= 0 && String.IsNullOrEmpty(row[end])) end--;
        if (end < 0) return null; // all empty
        var arr = new String[end + 1];
        for (var i = 0; i <= end; i++) arr[i] = row[i];
        return arr;
    }
    #endregion

    #region 对象映射辅助
    private static OdsSheet? FindSheet(List<OdsSheet> sheets, String? sheetName)
    {
        if (sheets.Count == 0) return null;
        if (String.IsNullOrEmpty(sheetName)) return sheets[0];
        foreach (var s in sheets)
        {
            if (String.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                return s;
        }
        return sheets[0];
    }

    private static IEnumerable<T> MapObjects<T>(List<String[]> rows) where T : new()
    {
        if (rows.Count == 0) yield break;

        var headers = rows[0];
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite).ToArray();

        // 建立列 → 属性映射
        var mapping = new PropertyInfo?[headers.Length];
        for (var i = 0; i < headers.Length; i++)
        {
            var h = headers[i];
            if (String.IsNullOrEmpty(h)) continue;
            foreach (var p in props)
            {
                if (String.Equals(p.Name, h, StringComparison.OrdinalIgnoreCase)) { mapping[i] = p; break; }
                var dn = p.GetCustomAttribute<DisplayNameAttribute>();
                if (dn != null && dn.DisplayName == h) { mapping[i] = p; break; }
                var desc = p.GetCustomAttribute<DescriptionAttribute>();
                if (desc != null && desc.Description == h) { mapping[i] = p; break; }
            }
        }

        for (var r = 1; r < rows.Count; r++)
        {
            var row = rows[r];
            var item = new T();
            for (var c = 0; c < Math.Min(row.Length, mapping.Length); c++)
            {
                var prop = mapping[c];
                if (prop == null || c >= row.Length) continue;
                var s = row[c];
                if (String.IsNullOrEmpty(s)) continue;
                try
                {
                    var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                    Object? val = targetType == typeof(String) ? s
                        : targetType == typeof(Int32) ? s.ToInt()
                        : targetType == typeof(Int64) ? (Int64.TryParse(s, out var v64) ? v64 : 0L)
                        : targetType == typeof(Double) ? s.ToDouble()
                        : targetType == typeof(Boolean) ? s.ToBoolean()
                        : targetType == typeof(DateTime) ? s.ToDateTime()
                        : Convert.ChangeType(s, targetType, CultureInfo.InvariantCulture);
                    prop.SetValue(item, val);
                }
                catch { /* 转换失败跳过 */ }
            }
            yield return item;
        }
    }

    private static DataTable BuildDataTable(List<String[]> rows)
    {
        var dt = new DataTable();
        if (rows.Count == 0) return dt;

        foreach (var col in rows[0])
        {
            dt.Columns.Add(String.IsNullOrEmpty(col) ? $"Column{dt.Columns.Count + 1}" : col);
        }

        for (var r = 1; r < rows.Count; r++)
        {
            var row = dt.NewRow();
            var data = rows[r];
            for (var c = 0; c < Math.Min(data.Length, dt.Columns.Count); c++)
            {
                row[c] = data[c] ?? "";
            }
            dt.Rows.Add(row);
        }
        return dt;
    }
    #endregion
}

/// <summary>ODS 工作表数据</summary>
public sealed class OdsSheet
{
    /// <summary>工作表名称</summary>
    public String Name { get; set; } = "";

    /// <summary>数据行列表（每行为字符串数组）</summary>
    public List<String[]> Rows { get; } = [];

    /// <summary>合并单元格区域列表</summary>
    public List<OdsMergeRegion> MergedCells { get; } = [];

    /// <summary>单元格样式字典，键为（行索引 0 基, 列索引 0 基）</summary>
    public Dictionary<(Int32 Row, Int32 Col), OdsCellStyle> CellStyles { get; } = [];
}

/// <summary>ODS 合并单元格区域</summary>
public sealed class OdsMergeRegion
{
    /// <summary>起始行（0 基）</summary>
    public Int32 Row { get; set; }

    /// <summary>起始列（0 基）</summary>
    public Int32 Col { get; set; }

    /// <summary>跨行数，≥ 1</summary>
    public Int32 RowSpan { get; set; } = 1;

    /// <summary>跨列数，≥ 1</summary>
    public Int32 ColSpan { get; set; } = 1;
}

/// <summary>ODS 单元格样式（字体、颜色、对齐）</summary>
public sealed class OdsCellStyle
{
    /// <summary>字体是否加粗</summary>
    public Boolean FontBold { get; set; }

    /// <summary>字体是否斜体</summary>
    public Boolean FontItalic { get; set; }

    /// <summary>字体大小（磅），0 表示未设置</summary>
    public Single FontSize { get; set; }

    /// <summary>字体颜色，格式 #RRGGBB；null 表示未设置</summary>
    public String? FontColor { get; set; }

    /// <summary>背景颜色，格式 #RRGGBB；null 表示未设置</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>水平对齐：start / center / end / left / right；null 表示未设置</summary>
    public String? HAlign { get; set; }
}
