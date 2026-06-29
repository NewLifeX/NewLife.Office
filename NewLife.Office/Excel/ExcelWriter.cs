using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Text;

namespace NewLife.Office;

/// <summary>轻量级Excel写入器，支持多个工作表</summary>
/// <remarks>
/// 目标：快速导出简单数据，支持多工作表的列头与多行数据；识别常见数据类型并使用合适样式，避免长数字（如身份证、长整型）被 Excel / WPS 显示为科学计数。
/// 支持单元格样式（字体/填充/边框/对齐）、合并单元格、冻结窗格、自动筛选、超链接、数据验证、图片、页面设置、条件格式等功能。
/// </remarks>
public partial class ExcelWriter : DisposeBase
{
    #region 内部类型
    /// <summary>单元格样式（值为 Excel 内置 numFmtId）。</summary>
    private enum ExcelCellStyle : Int32
    {
        General = 0,  // General
        Integer = 1,  // 0 （整数，避免长整型使用科学计数）
        Decimal = 2,  // 0.00
        Percent = 10, // 0.00%
        Date = 14,    // mm-dd-yy
        Time = 21,    // h:mm:ss
        DateTime = 22 // m/d/yy h:mm
    }

    private static readonly ExcelCellStyle[] _cellStyles = (ExcelCellStyle[])Enum.GetValues(typeof(ExcelCellStyle));

    private record FontEntry(String? Name, Double Size, Boolean Bold, Boolean Italic, Boolean Underline, String? Color, Boolean Strike, String? VerticalAlign);
    private record FillEntry(String? BgColor, String PatternType, String? GradientType = null, String? GradientColor1 = null, String? GradientColor2 = null, String? PatternFgColor = null, String? PatternTypeName = null);
    private record BorderEntry(
        ExcelCellBorderStyle Left, String? LeftColor,
        ExcelCellBorderStyle Right, String? RightColor,
        ExcelCellBorderStyle Top, String? TopColor,
        ExcelCellBorderStyle Bottom, String? BottomColor,
        ExcelCellBorderStyle Diagonal = ExcelCellBorderStyle.None, String? DiagonalColor = null);
    private record XfEntry(Int32 NumFmtId, Int32 FontId, Int32 FillId, Int32 BorderId, ExcelHorizontalAlignment HAlign, ExcelVerticalAlignment VAlign, Boolean WrapText, Int32 TextRotation, Int32 Indent, Boolean ShrinkToFit);

    private class SheetHyperlink
    {
        public Int32 Row { get; set; }
        public Int32 Col { get; set; }
        public String Url { get; set; } = null!;
        public String? Display { get; set; }
    }

    private class SheetValidation
    {
        public String CellRange { get; set; } = null!;
        public String[]? Items { get; set; }          // 下拉列表选项
        public String? ValidationType { get; set; }   // decimal, whole, date, time, textLength
        public String? Operator { get; set; }         // between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual
        public String? Formula1 { get; set; }
        public String? Formula2 { get; set; }
    }

    private class SheetImage
    {
        public Int32 Row { get; set; }
        public Int32 Col { get; set; }
        public Byte[] Data { get; set; } = null!;
        public String Extension { get; set; } = "png";
        public Double Width { get; set; }
        public Double Height { get; set; }
        public Int64 FromColOff { get; set; }
        public Int64 FromRowOff { get; set; }
        public Int32 ToRow { get; set; } = -1;
        public Int32 ToCol { get; set; } = -1;
        public Int64 ToColOff { get; set; }
        public Int64 ToRowOff { get; set; }
        public String EditAs { get; set; } = "oneCell";
    }

    private class SheetPageSetup
    {
        public ExcelPageOrientation Orientation { get; set; }
        public ExcelPaperSize PaperSize { get; set; }
        public Double MarginTop { get; set; } = 0.75;
        public Double MarginBottom { get; set; } = 0.75;
        public Double MarginLeft { get; set; } = 0.7;
        public Double MarginRight { get; set; } = 0.7;
        public String? HeaderText { get; set; }
        public String? FooterText { get; set; }
        public Int32 PrintTitleStartRow { get; set; } = -1;
        public Int32 PrintTitleEndRow { get; set; } = -1;
    }

    private class ConditionalFormatEntry
    {
        public String Range { get; set; } = null!;
        public ExcelConditionalFormatType Type { get; set; }
        public String? Value { get; set; }
        public String? Value2 { get; set; }
        public String? Color { get; set; }
        /// <summary>图标集类型名（如 "3Arrows"、"3Flags"、"5Rating"），仅 IconSet 使用</summary>
        public String? IconSetType { get; set; }
        /// <summary>自定义公式字符串（不含 = 号），仅 Expression 使用</summary>
        public String? Formula { get; set; }
    }

    private class SheetComment
    {
        public Int32 Row { get; set; }   // 1-based
        public Int32 Col { get; set; }   // 0-based
        public String Text { get; set; } = null!;
        public String Author { get; set; } = String.Empty;
    }

    /// <summary>Excel 文档属性（写入 docProps/core.xml）</summary>
    public class ExcelDocumentProperties
    {
        /// <summary>标题</summary>
        public String? Title { get; set; }

        /// <summary>作者/创建者</summary>
        public String? Creator { get; set; }

        /// <summary>主题</summary>
        public String? Subject { get; set; }

        /// <summary>描述</summary>
        public String? Description { get; set; }
    }
    #endregion

    #region 属性
    /// <summary>文件路径（Save 时写入）</summary>
    public String? FileName { get; }

    /// <summary>目标流（若提供则写入该流，调用方负责生命周期）</summary>
    public Stream? Stream { get; }

    /// <summary>默认工作表名称（当调用 API 未指定 sheet 时使用）</summary>
    public String SheetName { get; set; } = "Sheet1";

    /// <summary>文档属性</summary>
    public ExcelDocumentProperties? DocumentProperties { get; set; }

    /// <summary>文本编码</summary>
    public Encoding Encoding { get; set; } = Encoding.UTF8;

    /// <summary>超过该数字有效位数阈值（或极小值有大量前导0小数）则写为文本以避免科学计数法。默认 11。</summary>
    private const Int32 LongNumberAsTextThreshold = 11;

    /// <summary>是否自动根据数据内容估算列宽，并写入 <c>&lt;cols&gt;</c> 来避免 WPS/Excel 出现########。默认 true。</summary>
    public Boolean AutoFitColumnWidth { get; set; } = true;

    // 多 sheet：保持插入顺序，写 workbook.xml 时用于 sheetId 顺序
    private readonly List<String> _sheetNames = [];
    private readonly Dictionary<String, List<String>> _sheetRows = new(StringComparer.OrdinalIgnoreCase); // sheet -> 行XML集合
    private readonly Dictionary<String, Int32> _sheetRowIndex = new(StringComparer.OrdinalIgnoreCase);     // sheet -> 当前行号（1基）

    // 每个 sheet 的列最大显示宽度（字符数估算），下标 0 基，对应 Excel 列 1 基
    private readonly Dictionary<String, List<Double>> _sheetColWidths = new(StringComparer.OrdinalIgnoreCase);

    private readonly Dictionary<String, Int32> _shared = new(StringComparer.Ordinal); // 共享字符串去重
    private Int32 _sharedCount; // 总引用次数（含重复）

    // 样式管理（字体/填充/边框/XF 去重）
    private readonly List<FontEntry> _fonts = [new(null, 0, false, false, false, null, false, null)]; // index 0 = 默认字体
    private readonly List<FillEntry> _fills = [new(null, "none"), new(null, "gray125")]; // 0=none, 1=gray125 (Excel 要求)
    private readonly List<BorderEntry> _borders = [new(ExcelCellBorderStyle.None, null, ExcelCellBorderStyle.None, null, ExcelCellBorderStyle.None, null, ExcelCellBorderStyle.None, null)]; // index 0 = 无边框
    private readonly Dictionary<String, Int32> _numFmtMap = new(StringComparer.Ordinal); // formatCode → numFmtId
    private Int32 _nextNumFmtId = 164; // 自定义 numFmt 从 164 开始
    private readonly List<XfEntry> _xfEntries;
    private readonly Dictionary<String, Int32> _xfCache = new(StringComparer.Ordinal); // 复合键 → XF 索引

    // 合并单元格：sheet -> [(startRow, startCol, endRow, endCol)]
    private readonly Dictionary<String, List<(Int32, Int32, Int32, Int32)>> _sheetMerges = new(StringComparer.OrdinalIgnoreCase);
    // 冻结窗格：sheet -> (rows, cols)
    private readonly Dictionary<String, (Int32 Rows, Int32 Cols)> _sheetFreezes = new(StringComparer.OrdinalIgnoreCase);
    // 自动筛选：sheet -> ref ("A1:F1")
    private readonly Dictionary<String, String> _sheetAutoFilters = new(StringComparer.OrdinalIgnoreCase);
    // 行高：sheet -> { rowIndex(1基) -> height }
    private readonly Dictionary<String, Dictionary<Int32, Double>> _sheetRowHeights = new(StringComparer.OrdinalIgnoreCase);
    // 列大纲：sheet -> { colIndex(0基) -> (level, collapsed) }
    private readonly Dictionary<String, Dictionary<Int32, (Int32 Level, Boolean Collapsed)>> _sheetColOutlines = new(StringComparer.OrdinalIgnoreCase);
    // 行大纲：sheet -> { rowIndex(1基) -> (level, collapsed) }
    private readonly Dictionary<String, Dictionary<Int32, (Int32 Level, Boolean Collapsed)>> _sheetRowOutlines = new(StringComparer.OrdinalIgnoreCase);
    // 工作表标签颜色：sheet -> RGB六位十六进制
    private readonly Dictionary<String, String> _sheetTabColors = new(StringComparer.OrdinalIgnoreCase);
    // 工作簿保护密码哈希（null 表示不保护）
    private String? _workbookProtectionHash;
    private Boolean _workbookLockStructure;
    private Boolean _workbookLockWindows;

    // 图表：sheet → 图表列表
    private readonly Dictionary<String, List<ExcelChart>> _sheetCharts = new(StringComparer.OrdinalIgnoreCase);
    // 超链接
    private readonly Dictionary<String, List<SheetHyperlink>> _sheetHyperlinks = new(StringComparer.OrdinalIgnoreCase);
    // 数据验证
    private readonly Dictionary<String, List<SheetValidation>> _sheetValidations = new(StringComparer.OrdinalIgnoreCase);
    // 图片
    private readonly Dictionary<String, List<SheetImage>> _sheetImages = new(StringComparer.OrdinalIgnoreCase);
    // 页面设置
    private readonly Dictionary<String, SheetPageSetup> _sheetPageSetups = new(StringComparer.OrdinalIgnoreCase);
    // 工作表保护
    private readonly Dictionary<String, String?> _sheetProtection = new(StringComparer.OrdinalIgnoreCase);
    // 工作表可见性：sheet → state ("visible"/"hidden"/"veryHidden")
    private readonly Dictionary<String, String> _sheetStates = new(StringComparer.OrdinalIgnoreCase);
    // 条件格式
    private readonly Dictionary<String, List<ConditionalFormatEntry>> _sheetCondFormats = new(StringComparer.OrdinalIgnoreCase);
    // 批注
    private readonly Dictionary<String, List<SheetComment>> _sheetComments = new(StringComparer.OrdinalIgnoreCase);

    // 分页符：sheet → 行号列表（1基，此行开始新页）
    private readonly Dictionary<String, List<Int32>> _sheetPageBreaks = new(StringComparer.OrdinalIgnoreCase);
    // 垂直分页符：sheet → 列号列表（0基，此列开始新页）
    private readonly Dictionary<String, List<Int32>> _sheetColPageBreaks = new(StringComparer.OrdinalIgnoreCase);

    // 迷你图组：sheet → 迷你图组列表
    private readonly Dictionary<String, List<SparklineGroup>> _sheetSparklineGroups = new(StringComparer.OrdinalIgnoreCase);

    // 每单元格样式覆盖（0基行, 0基列）→ CellStyle
    private readonly Dictionary<String, Dictionary<(Int32 Row, Int32 Col), Office.ExcelCellStyle>> _cellStyleOverrides = new(StringComparer.OrdinalIgnoreCase);

    // OtherParts 透传：Reader 收集的原始 ZIP 部件，Save 时原样写回
    private Dictionary<String, Byte[]> _otherParts = [];

    // 用户自定义命名范围（排除 _xlnm.* 系统名）
    private readonly List<(String Name, String Formula)> _definedNames = [];

    // 结构化表格：sheet → 表格列表
    private readonly Dictionary<String, List<ExcelTableInfo>> _sheetTables = new(StringComparer.OrdinalIgnoreCase);
    #endregion

    #region 构造
    /// <summary>使用文件路径实例化写入器</summary>
    /// <param name="fileName">目标 xlsx 文件</param>
    public ExcelWriter(String fileName)
    {
        FileName = fileName.GetFullPath();
        _xfEntries = InitBuiltinXfEntries();
    }

    /// <summary>使用外部流实例化写入器</summary>
    /// <param name="stream">目标可写流</param>
    public ExcelWriter(Stream stream)
    {
        Stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _xfEntries = InitBuiltinXfEntries();
    }

    /// <summary>销毁释放。仅显式 Dispose 时自动 Save，析构函数不写文件（GC 回收时托管资源可能已释放）</summary>
    /// <param name="disposing"></param>
    protected override void Dispose(Boolean disposing)
    {
        base.Dispose(disposing);
        if (disposing && Stream == null) Save();
    }

    private static List<XfEntry> InitBuiltinXfEntries()
    {
        // 按 _cellStyles 枚举值升序，生成内置 XF 条目（全使用默认字体/填充/边框）
        var list = new List<XfEntry>();
        foreach (var st in _cellStyles)
        {
            list.Add(new XfEntry((Int32)st, 0, 0, 0, ExcelHorizontalAlignment.General, ExcelVerticalAlignment.Top, false, 0, 0, false));
        }
        return list;
    }
    #endregion

    #region 写入接口
    /// <summary>写入列头到指定工作表</summary>
    /// <param name="sheet">工作表名称（可空，空时使用 <see cref="SheetName"/>）</param>
    /// <param name="headers">列头文本集合</param>
    public void WriteHeader(String sheet, IEnumerable<String> headers)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        if (headers == null) throw new ArgumentNullException(nameof(headers));

        EnsureSheet(sheet);

        var arr = headers as String[] ?? headers.ToArray();
        AddRow(sheet, arr.Select(e => (Object?)e).ToArray());
    }

    /// <summary>写入列头到指定工作表（带样式）</summary>
    /// <param name="sheet">工作表名称（可空，空时使用 <see cref="SheetName"/>）</param>
    /// <param name="headers">列头文本集合</param>
    /// <param name="style">表头单元格样式</param>
    public void WriteHeader(String sheet, IEnumerable<String> headers, Office.ExcelCellStyle? style)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        if (headers == null) throw new ArgumentNullException(nameof(headers));

        EnsureSheet(sheet);

        var arr = headers as String[] ?? headers.ToArray();
        AddRow(sheet, arr.Select(e => (Object?)e).ToArray(), style);
    }

    /// <summary>写入多行数据到指定工作表</summary>
    /// <param name="sheet">工作表名称（可空，空时使用 <see cref="SheetName"/>）</param>
    /// <param name="data">数据集合，每行一个对象数组</param>
    public void WriteRows(String? sheet, IEnumerable<Object?[]> data)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));

        if (sheet.IsNullOrEmpty())
            sheet = SheetName;
        else
            SheetName = sheet; // 同步默认值为最近使用

        EnsureSheet(sheet);

        foreach (var row in data)
        {
            AddRow(sheet, row);
        }
    }

    /// <summary>写入多行数据到指定工作表（带统一样式）</summary>
    /// <param name="sheet">工作表名称（可空，空时使用 <see cref="SheetName"/>）</param>
    /// <param name="data">数据集合，每行一个对象数组</param>
    /// <param name="style">统一单元格样式</param>
    public void WriteRows(String? sheet, IEnumerable<Object?[]> data, Office.ExcelCellStyle? style)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));

        if (sheet.IsNullOrEmpty())
            sheet = SheetName;
        else
            SheetName = sheet;

        EnsureSheet(sheet);

        foreach (var row in data)
        {
            AddRow(sheet, row, style);
        }
    }

    /// <summary>写入单行数据</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="values">单行数据</param>
    /// <param name="style">单元格样式</param>
    public void WriteRow(String? sheet, Object?[] values, Office.ExcelCellStyle? style = null)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        AddRow(sheet, values, style);
    }

    /// <summary>手工设置列宽（字符宽度，近似），0 基列序号。需在 Save 之前调用。</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="columnIndex">列序号（0基）</param>
    /// <param name="width">字符宽度</param>
    public void SetColumnWidth(String? sheet, Int32 columnIndex, Double width)
    {
        if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet!);

        var list = _sheetColWidths[sheet!];
        while (list.Count <= columnIndex) list.Add(0);
        if (width > list[columnIndex]) list[columnIndex] = width;
    }
    #endregion

    #region 布局设置
    /// <summary>合并单元格（Excel 记法，如 "A1:F1"）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="range">合并范围，如 "A1:F1"</param>
    public void MergeCell(String? sheet, String range)
    {
        if (range.IsNullOrEmpty()) throw new ArgumentNullException(nameof(range));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        var parts = range.Split(':');
        if (parts.Length != 2) throw new ArgumentException("范围格式应为 A1:F1", nameof(range));

        var (r1, c1) = ParseCellRef(parts[0]);
        var (r2, c2) = ParseCellRef(parts[1]);
        MergeCell(sheet, r1, c1, r2, c2);
    }

    /// <summary>合并单元格（行列索引，0基）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="startRow">起始行（0基）</param>
    /// <param name="startCol">起始列（0基）</param>
    /// <param name="endRow">结束行（0基）</param>
    /// <param name="endCol">结束列（0基）</param>
    public void MergeCell(String? sheet, Int32 startRow, Int32 startCol, Int32 endRow, Int32 endCol)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetMerges.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetMerges[sheet] = list;
        }
        list.Add((startRow, startCol, endRow, endCol));
    }

    /// <summary>冻结窗格</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="rows">冻结的行数（如 1 = 冻结首行）</param>
    /// <param name="cols">冻结的列数（如 1 = 冻结首列）</param>
    public void FreezePane(String? sheet, Int32 rows, Int32 cols = 0)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        _sheetFreezes[sheet] = (rows, cols);
    }

    /// <summary>设置自动筛选（Excel 记法，如 "A1:F1"）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="range">筛选范围，如 "A1:F1"</param>
    public void SetAutoFilter(String? sheet, String range)
    {
        if (range.IsNullOrEmpty()) throw new ArgumentNullException(nameof(range));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        _sheetAutoFilters[sheet] = range;
    }

    /// <summary>设置行高</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="row">行号（1基）</param>
    /// <param name="height">行高（磅值）</param>
    public void SetRowHeight(String? sheet, Int32 row, Double height)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetRowHeights.TryGetValue(sheet, out var dict))
        {
            dict = [];
            _sheetRowHeights[sheet] = dict;
        }
        dict[row] = height;
    }

    /// <summary>设置列大纲/分组级别（用于折叠展开列）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="columnIndex">列号（0基）</param>
    /// <param name="level">大纲级别（1-8，0 表示取消分组）</param>
    /// <param name="collapsed">是否默认折叠</param>
    public void SetColumnOutlineLevel(String? sheet, Int32 columnIndex, Int32 level, Boolean collapsed = false)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!_sheetColOutlines.TryGetValue(sheet, out var dict))
        {
            dict = [];
            _sheetColOutlines[sheet] = dict;
        }
        dict[columnIndex] = (level, collapsed);
    }

    /// <summary>设置行大纲/分组级别（用于折叠展开行）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="row">行号（1基）</param>
    /// <param name="level">大纲级别（1-8，0 表示取消分组）</param>
    /// <param name="collapsed">是否默认折叠</param>
    public void SetRowOutlineLevel(String? sheet, Int32 row, Int32 level, Boolean collapsed = false)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!_sheetRowOutlines.TryGetValue(sheet, out var dict))
        {
            dict = [];
            _sheetRowOutlines[sheet] = dict;
        }
        dict[row] = (level, collapsed);
    }

    /// <summary>设置工作表标签颜色</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="color">RGB 六位十六进制（如 "FF0000"），null 表示清除颜色</param>
    public void SetSheetTabColor(String? sheet, String? color)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (color.IsNullOrEmpty())
            _sheetTabColors.Remove(sheet!);
        else
            _sheetTabColors[sheet!] = color!;
    }

    /// <summary>设置工作簿保护（防止结构/窗口被修改）</summary>
    /// <param name="password">保护密码（null 表示无密码保护）</param>
    /// <param name="lockStructure">是否锁定工作表结构（添加/移动/删除/重命名）</param>
    /// <param name="lockWindows">是否锁定窗口位置和大小</param>
    public void ProtectWorkbook(String? password, Boolean lockStructure = true, Boolean lockWindows = false)
    {
        _workbookLockStructure = lockStructure;
        _workbookLockWindows = lockWindows;
        if (password.IsNullOrEmpty())
        {
            _workbookProtectionHash = String.Empty; // 无密码但启用保护
        }
        else
        {
            // 使用 xor + count 算法（与 Excel 97-2003 兼容的简单哈希）
            // 注：xlsx 实际支持更安全的哈希算法，这里使用最简单的兼容实现
            _workbookProtectionHash = ComputeXorHash(password!);
        }
    }

    private static String ComputeXorHash(String password)
    {
        // Excel 97-2003 式密码保护哈希（xor 算法），与 ProtectSheet 复用同一实现
        var hash = 0;
        if (password.Length == 0) return "0000";
        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= password[i];
        }
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
        hash ^= (password.Length + 0x8000);
        return hash.ToString("X4");
    }
    #endregion

    #region 超链接
    /// <summary>添加超链接</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="row">行号（1基）</param>
    /// <param name="col">列号（0基）</param>
    /// <param name="url">链接地址</param>
    /// <param name="displayText">显示文本（可空，空时显示 URL）</param>
    public void AddHyperlink(String? sheet, Int32 row, Int32 col, String url, String? displayText = null)
    {
        if (url.IsNullOrEmpty()) throw new ArgumentNullException(nameof(url));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetHyperlinks.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetHyperlinks[sheet] = list;
        }
        list.Add(new SheetHyperlink { Row = row, Col = col, Url = url, Display = displayText });
    }
    #endregion

    #region 数据验证
    /// <summary>添加下拉列表数据验证</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="cellRange">验证范围（如 "A2:A100"）</param>
    /// <param name="items">下拉选项列表</param>
    public void AddDropdownValidation(String? sheet, String cellRange, String[] items)
    {
        if (cellRange.IsNullOrEmpty()) throw new ArgumentNullException(nameof(cellRange));
        if (items == null || items.Length == 0) throw new ArgumentNullException(nameof(items));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetValidations.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetValidations[sheet] = list;
        }
        list.Add(new SheetValidation { CellRange = cellRange, Items = items });
    }

    /// <summary>添加数值/日期范围数据验证</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="cellRange">验证范围（如 "B2:B100"）</param>
    /// <param name="validationType">验证类型：whole（整数）、decimal（小数）、date（日期）、time（时间）、textLength（文本长度）</param>
    /// <param name="operator">运算符：between、notBetween、equal、notEqual、greaterThan、lessThan、greaterThanOrEqual、lessThanOrEqual</param>
    /// <param name="formula1">最小值（或比较值）</param>
    /// <param name="formula2">最大值（仅 between 和 notBetween 有效）</param>
    public void AddRangeValidation(String? sheet, String cellRange,
        String validationType = "whole",
        String @operator = "between",
        String formula1 = "0",
        String? formula2 = null)
    {
        if (cellRange.IsNullOrEmpty()) throw new ArgumentNullException(nameof(cellRange));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetValidations.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetValidations[sheet] = list;
        }
        list.Add(new SheetValidation
        {
            CellRange = cellRange,
            ValidationType = validationType,
            Operator = @operator,
            Formula1 = formula1,
            Formula2 = formula2,
        });
    }
    #endregion

    #region 图片
    /// <summary>插入图片</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="row">行号（1基）</param>
    /// <param name="col">列号（0基）</param>
    /// <param name="imageData">图片数据</param>
    /// <param name="extension">图片格式（如 "png"、"jpeg"）</param>
    /// <param name="widthPx">图片宽度（像素）</param>
    /// <param name="heightPx">图片高度（像素）</param>
    public void AddImage(String? sheet, Int32 row, Int32 col, Byte[] imageData, String extension = "png", Double widthPx = 100, Double heightPx = 100)
    {
        AddImage(sheet, row, col, imageData, extension, widthPx, heightPx, 0, 0, -1, -1, 0, 0, "oneCell");
    }

    /// <summary>插入图片（完整锚点信息）</summary>
    private void AddImage(String? sheet, Int32 fromRow, Int32 fromCol, Byte[] imageData, String extension,
        Double widthPx, Double heightPx, Int64 fromColOff, Int64 fromRowOff,
        Int32 toRow, Int32 toCol, Int64 toColOff, Int64 toRowOff, String editAs)
    {
        if (imageData == null || imageData.Length == 0) throw new ArgumentNullException(nameof(imageData));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetImages.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetImages[sheet] = list;
        }
        list.Add(new SheetImage
        {
            Row = fromRow, Col = fromCol,
            Data = imageData, Extension = extension.ToLower().TrimStart('.'),
            Width = widthPx, Height = heightPx,
            FromColOff = fromColOff, FromRowOff = fromRowOff,
            ToRow = toRow, ToCol = toCol,
            ToColOff = toColOff, ToRowOff = toRowOff,
            EditAs = editAs,
        });
    }
    #endregion

    #region 页面设置
    /// <summary>设置页面方向和纸张大小</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="orientation">页面方向</param>
    /// <param name="paperSize">纸张大小</param>
    public void SetPageSetup(String? sheet, ExcelPageOrientation orientation, ExcelPaperSize paperSize = ExcelPaperSize.A4)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        var ps = GetOrCreatePageSetup(sheet);
        ps.Orientation = orientation;
        ps.PaperSize = paperSize;
    }

    /// <summary>设置页边距（英寸）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="top">上边距</param>
    /// <param name="bottom">下边距</param>
    /// <param name="left">左边距</param>
    /// <param name="right">右边距</param>
    public void SetPageMargins(String? sheet, Double top, Double bottom, Double left, Double right)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        var ps = GetOrCreatePageSetup(sheet);
        ps.MarginTop = top;
        ps.MarginBottom = bottom;
        ps.MarginLeft = left;
        ps.MarginRight = right;
    }

    /// <summary>设置页眉页脚文本</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="header">页眉文本</param>
    /// <param name="footer">页脚文本</param>
    public void SetHeaderFooter(String? sheet, String? header, String? footer)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        var ps = GetOrCreatePageSetup(sheet);
        ps.HeaderText = header;
        ps.FooterText = footer;
    }

    /// <summary>设置打印标题行（每页重复打印）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="startRow">起始行（1基）</param>
    /// <param name="endRow">结束行（1基）</param>
    public void SetPrintTitleRows(String? sheet, Int32 startRow, Int32 endRow)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        var ps = GetOrCreatePageSetup(sheet);
        ps.PrintTitleStartRow = startRow;
        ps.PrintTitleEndRow = endRow;
    }

    /// <summary>添加用户自定义命名范围</summary>
    /// <param name="name">名称（须符合 Excel 命名规则，不可以 _xlnm. 开头）</param>
    /// <param name="formula">公式或范围引用（如 "Sheet1!$A$1:$B$10" 或 "'数据'!$C:$C"）</param>
    public void AddDefinedName(String name, String formula)
    {
        if (name.IsNullOrEmpty()) throw new ArgumentNullException(nameof(name));
        if (formula.IsNullOrEmpty()) throw new ArgumentNullException(nameof(formula));
        _definedNames.Add((name, formula));
    }

    /// <summary>按名称获取已添加的命名范围公式</summary>
    /// <param name="name">命名范围名称（大小写不敏感）</param>
    /// <returns>范围公式字符串，未找到返回 null</returns>
    public String? GetRangeByName(String name)
    {
        if (name.IsNullOrEmpty()) return null;
        foreach (var (n, f) in _definedNames)
        {
            if (n.EqualIgnoreCase(name)) return f;
        }
        return null;
    }

    /// <summary>设置打印区域</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="range">区域引用（如 "A1:F50"）</param>
    public void SetPrintArea(String? sheet, String range)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!range.Contains('!'))
            range = $"'{sheet}'!{range}";
        _definedNames.RemoveAll(dn => dn.Name.EqualIgnoreCase("_xlnm.Print_Area") && dn.Formula.Contains($"'{sheet}'!"));
        AddDefinedName("_xlnm.Print_Area", range);
    }

    /// <summary>添加迷你图组（行内微型图表）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="dataRange">数据区域（如 "Sheet1!B2:F2"）</param>
    /// <param name="cellRange">放置迷你图的单元格区域（如 "Sheet1!G2:G2"）</param>
    /// <param name="type">类型：line/column/stacked</param>
    /// <param name="lineColor">线条/柱颜色（16进制RGB）</param>
    /// <param name="markerColor">标记点颜色（可空）</param>
    /// <returns>迷你图组对象</returns>
    public SparklineGroup AddSparklineGroup(String? sheet, String dataRange, String cellRange, String type = "line", String? lineColor = "FF0000", String? markerColor = null)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        var sg = new SparklineGroup
        {
            Type = type,
            DataRange = dataRange,
            CellRange = cellRange,
            LineColor = lineColor ?? "FF0000",
            MarkerColor = markerColor
        };
        if (!_sheetSparklineGroups.TryGetValue(sheet, out var list))
            _sheetSparklineGroups[sheet] = list = [];
        list.Add(sg);
        return sg;
    }

    /// <summary>迷你图组定义</summary>
    public class SparklineGroup
    {
        /// <summary>类型：line/column/stacked</summary>
        public String Type { get; set; } = "line";
        /// <summary>数据区域</summary>
        public String DataRange { get; set; } = String.Empty;
        /// <summary>放置单元格区域</summary>
        public String CellRange { get; set; } = String.Empty;
        /// <summary>线条/柱颜色（16进制RGB）</summary>
        public String LineColor { get; set; } = "FF0000";
        /// <summary>标记点颜色</summary>
        public String? MarkerColor { get; set; }
    }

    /// <summary>设置水平分页符</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="row">分页符所在行（1基，此行开始新页）</param>
    public void SetPageBreak(String? sheet, Int32 row)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!_sheetPageBreaks.TryGetValue(sheet, out var breaks))
            _sheetPageBreaks[sheet] = breaks = [];
        if (!breaks.Contains(row))
            breaks.Add(row);
    }

    /// <summary>设置垂直分页符</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="col">分页符所在列（1基，此列开始新页）</param>
    public void SetColumnPageBreak(String? sheet, Int32 col)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!_sheetColPageBreaks.TryGetValue(sheet, out var breaks))
            _sheetColPageBreaks[sheet] = breaks = [];
        if (!breaks.Contains(col))
            breaks.Add(col);
    }

    /// <summary>在当前工作表中添加结构化表格（OOXML table 元素）</summary>
    /// <param name="range">表格范围（Excel 记法，如 "A1:E10"，含表头行）</param>
    /// <param name="name">表格名称（同时作为表格引用标识）</param>
    /// <param name="style">表格样式名称（如 "TableStyleMedium9"，默认不传时使用 Medium9）</param>
    /// <param name="columnNames">列名集合；null 时从范围列位置自动生成 Column1/Column2...</param>
    public void AddTable(String range, String name, String? style = null, String[]? columnNames = null)
    {
        if (range.IsNullOrEmpty()) throw new ArgumentNullException(nameof(range));
        if (name.IsNullOrEmpty()) throw new ArgumentNullException(nameof(name));
        EnsureSheet(SheetName);
        if (!_sheetTables.TryGetValue(SheetName, out var tables))
        {
            tables = [];
            _sheetTables[SheetName] = tables;
        }
        tables.Add(new ExcelTableInfo
        {
            Range = range,
            Name = name,
            StyleName = style ?? "TableStyleMedium9",
            ColumnNames = columnNames,
        });
    }

    /// <summary>在指定工作表中添加结构化表格（OOXML table 元素）</summary>
    /// <param name="sheet">工作表名称</param>
    /// <param name="range">表格范围</param>
    /// <param name="name">表格名称</param>
    /// <param name="style">表格样式名称</param>
    /// <param name="columnNames">列名集合</param>
    public void AddTable(String sheet, String range, String name, String? style = null, String[]? columnNames = null)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        if (range.IsNullOrEmpty()) throw new ArgumentNullException(nameof(range));
        if (name.IsNullOrEmpty()) throw new ArgumentNullException(nameof(name));
        EnsureSheet(sheet);
        if (!_sheetTables.TryGetValue(sheet, out var tables))
        {
            tables = [];
            _sheetTables[sheet] = tables;
        }
        tables.Add(new ExcelTableInfo
        {
            Range = range,
            Name = name,
            StyleName = style ?? "TableStyleMedium9",
            ColumnNames = columnNames,
        });
    }

    /// <summary>添加图表到工作表</summary>
    /// <param name="sheet">工作表名称（可空，空时用当前工作表）</param>
    /// <param name="chart">图表定义对象</param>
    public void AddChart(String? sheet, ExcelChart chart)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!_sheetCharts.TryGetValue(sheet, out var charts))
        {
            charts = [];
            _sheetCharts[sheet] = charts;
        }
        charts.Add(chart);
    }

    private SheetPageSetup GetOrCreatePageSetup(String sheet)
    {
        if (!_sheetPageSetups.TryGetValue(sheet, out var ps))
        {
            ps = new SheetPageSetup();
            _sheetPageSetups[sheet] = ps;
        }
        return ps;
    }
    #endregion

    #region 工作表保护
    /// <summary>保护工作表</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="password">保护密码（可空，空时仅启用保护无密码）</param>
    public void ProtectSheet(String? sheet, String? password = null)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        _sheetProtection[sheet] = password;
    }

    /// <summary>设置工作表可见性</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="veryHidden">true=深度隐藏（仅 VBA 可取消隐藏），false=普通隐藏（用户可从 UI 取消隐藏）</param>
    public void HideSheet(String? sheet, Boolean veryHidden = false)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        _sheetStates[sheet] = veryHidden ? "veryHidden" : "hidden";
    }

    /// <summary>恢复工作表可见</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    public void UnhideSheet(String? sheet)
    {
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        _sheetStates.Remove(sheet);
    }
    #endregion

    #region 公式
    /// <summary>在指定行写入公式单元格（与 WriteRow 配合使用）</summary>
    /// <remarks>更简单的方式是在 WriteRow 的 values 数组中直接传入 <see cref="ExcelFormula"/> 实例。</remarks>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="formula">公式文本（不含等号，如 "SUM(A1:A10)"）</param>
    /// <param name="cachedValue">缓存值（可空）</param>
    public void AppendFormula(String? sheet, String formula, Object? cachedValue = null)
    {
        if (formula.IsNullOrEmpty()) throw new ArgumentNullException(nameof(formula));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        // 包装为 ExcelFormula 放入当前行
        AddRow(sheet, [new ExcelFormula(formula, cachedValue)]);
    }
    #endregion

    #region 批注
    /// <summary>为指定单元格添加批注</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="row">行号（1基）</param>
    /// <param name="col">列号（0基）</param>
    /// <param name="text">批注文本</param>
    /// <param name="author">批注作者（可空）</param>
    public void AddComment(String? sheet, Int32 row, Int32 col, String text, String? author = null)
    {
        if (text.IsNullOrEmpty()) throw new ArgumentNullException(nameof(text));
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetComments.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetComments[sheet] = list;
        }
        list.Add(new SheetComment { Row = row, Col = col, Text = text, Author = author ?? String.Empty });
    }

    /// <summary>设置指定单元格的样式（覆盖行级样式）</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="row">行号（0基）</param>
    /// <param name="col">列号（0基）</param>
    /// <param name="style">单元格样式</param>
    public void SetCellStyle(String? sheet, Int32 row, Int32 col, Office.ExcelCellStyle style)
    {
        if (style == null) throw new ArgumentNullException(nameof(style));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_cellStyleOverrides.TryGetValue(sheet, out var dict))
        {
            dict = [];
            _cellStyleOverrides[sheet] = dict;
        }
        dict[(row, col)] = style;
    }
    #endregion

    #region 条件格式
    /// <summary>添加条件格式</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="range">应用范围（如 "A1:A100"）</param>
    /// <param name="type">条件类型</param>
    /// <param name="value">条件值</param>
    /// <param name="color">满足条件时的背景色（RGB十六进制）</param>
    /// <param name="value2">第二个条件值（仅 Between 类型使用）</param>
    public void AddConditionalFormat(String? sheet, String range, ExcelConditionalFormatType type, String? value, String? color, String? value2 = null)
    {
        if (range.IsNullOrEmpty()) throw new ArgumentNullException(nameof(range));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetCondFormats.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetCondFormats[sheet] = list;
        }
        list.Add(new ConditionalFormatEntry { Range = range, Type = type, Value = value, Value2 = value2, Color = color });
    }

    /// <summary>添加图标集条件格式</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="range">应用范围（如 "A1:A100"）</param>
    /// <param name="iconSetType">图标集类型（如 "3Arrows"、"3Flags"、"3TrafficLights1"、"4Rating"、"5Rating"）</param>
    public void AddIconSetConditionalFormat(String? sheet, String range, String iconSetType = "3Arrows")
    {
        if (range.IsNullOrEmpty()) throw new ArgumentNullException(nameof(range));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!_sheetCondFormats.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetCondFormats[sheet] = list;
        }
        list.Add(new ConditionalFormatEntry { Range = range, Type = ExcelConditionalFormatType.IconSet, IconSetType = iconSetType });
    }

    /// <summary>添加自定义公式条件格式</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="range">应用范围（如 "A1:A100"）</param>
    /// <param name="formula">Excel 公式（不含 = 号，如 "A1&gt;100"、"AND(A1&gt;0,B1&lt;10)"）</param>
    /// <param name="color">满足条件时的背景色（RGB十六进制）</param>
    public void AddExpressionConditionalFormat(String? sheet, String range, String formula, String? color)
    {
        if (range.IsNullOrEmpty()) throw new ArgumentNullException(nameof(range));
        if (formula.IsNullOrEmpty()) throw new ArgumentNullException(nameof(formula));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);
        if (!_sheetCondFormats.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetCondFormats[sheet] = list;
        }
        list.Add(new ConditionalFormatEntry { Range = range, Type = ExcelConditionalFormatType.Expression, Formula = formula, Color = color });
    }
    #endregion

    #region 对象映射
    /// <summary>将对象集合导出到工作表</summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="data">对象集合</param>
    /// <param name="headerStyle">表头样式</param>
    public void WriteObjects<T>(String? sheet, IEnumerable<T> data, Office.ExcelCellStyle? headerStyle = null) where T : class
    {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(e => e.CanRead)
            .ToArray();

        // 表头：优先使用 DisplayName → Description → 属性名
        var headers = new String[props.Length];
        for (var i = 0; i < props.Length; i++)
        {
            var dn = props[i].GetCustomAttribute<DisplayNameAttribute>();
            if (dn != null && !dn.DisplayName.IsNullOrEmpty()) { headers[i] = dn.DisplayName; continue; }
            var desc = props[i].GetCustomAttribute<DescriptionAttribute>();
            if (desc != null && !desc.Description.IsNullOrEmpty()) { headers[i] = desc.Description; continue; }
            headers[i] = props[i].Name;
        }
        WriteHeader(sheet, headers, headerStyle);

        // 数据行
        foreach (var item in data)
        {
            var values = new Object?[props.Length];
            for (var i = 0; i < props.Length; i++)
            {
                values[i] = props[i].GetValue(item);
            }
            AddRow(sheet, values);
        }
    }

    /// <summary>将 DataTable 导出到工作表</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="table">DataTable</param>
    /// <param name="headerStyle">表头样式</param>
    public void WriteDataTable(String? sheet, DataTable table, Office.ExcelCellStyle? headerStyle = null)
    {
        if (table == null) throw new ArgumentNullException(nameof(table));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        var headers = new String[table.Columns.Count];
        for (var i = 0; i < table.Columns.Count; i++)
        {
            headers[i] = table.Columns[i].ColumnName;
        }
        WriteHeader(sheet, headers, headerStyle);

        foreach (DataRow dr in table.Rows)
        {
            AddRow(sheet, dr.ItemArray);
        }
    }
    #endregion

    #region ExcelData写入
    /// <summary>从完整快照写入工作簿</summary>
    /// <param name="data">ExcelData 快照数据</param>
    public void WriteExcel(ExcelDocument data)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));

        var autoFit = AutoFitColumnWidth;
        AutoFitColumnWidth = false; // 写 ExcelData 时用预设列宽
        _otherParts = data.OtherParts.Count > 0 ? new Dictionary<String, Byte[]>(data.OtherParts) : [];

        // 用源文件的默认字体覆盖 font[0]，确保行/列标题字体与原文件一致
        if (data.DefaultFont != null)
        {
            var df = data.DefaultFont;
            _fonts[0] = new FontEntry(df.Name, df.Size, df.Bold, false, false, df.Color, false, null);
        }

        try
        {
            foreach (var sd in data.Sheets)
            {
                var sheet = sd.Name;
                EnsureSheet(sheet);

                // 写入数据行（带每单元格样式）
                var prevActualRow = -1;
                for (var r = 0; r < sd.Rows.Count; r++)
                {
                    // 计算此行的实际 Excel 行号（0基）
                    var actualRow = sd.ActualRowNumbers != null ? sd.ActualRowNumbers[r] : r;

                // 若源文件存在跳行（如行 13 为空），推进行号计数器跳过缺失的行，
                // 确保后续行的 r 属性与原始 Excel 行号一致（不插入空行元素）
                if (actualRow > prevActualRow + 1)
                    AdvanceToRow(sheet, actualRow + 1); // actualRow 是 0 基，目标行号 = actualRow+1（1基）
                prevActualRow = actualRow;

                    // 先设置该行的每单元格样式覆盖（需在 AddRow 之前）
                    foreach (var kv in sd.CellStyles)
                    {
                        var (cr, cc) = kv.Key;
                        if (cr == actualRow)
                            SetCellStyle(sheet, actualRow, cc, kv.Value);
                    }

                    // 检查公式——将公式单元格的值包装为 ExcelFormula（不修改原始数据）
                    var row = (Object?[])sd.Rows[r].Clone();
                    for (var c = 0; c < row.Length; c++)
                    {
                        if (sd.Formulas.TryGetValue((actualRow, c), out var formula) && !formula.IsNullOrEmpty())
                        {
                            row[c] = new ExcelFormula(formula, row[c]);
                        }
                    }
                    AddRow(sheet, row);
                }

                // 合并区域
                foreach (var (sr, sc, er, ec) in sd.Merges)
                {
                    MergeCell(sheet, sr, sc, er, ec);
                }

                // 冻结窗格
                if (sd.FreezePane.HasValue)
                    FreezePane(sheet, sd.FreezePane.Value.Rows, sd.FreezePane.Value.Cols);

                // 自动筛选
                if (!sd.AutoFilter.IsNullOrEmpty())
                    SetAutoFilter(sheet, sd.AutoFilter!);

                // 行高（0基→1基）
                foreach (var kv in sd.RowHeights)
                {
                    SetRowHeight(sheet, kv.Key + 1, kv.Value);
                }

                // 列宽（已经是0基）
                foreach (var kv in sd.ColumnWidths)
                {
                    SetColumnWidth(sheet, kv.Key, kv.Value);
                }

                // 超链接（0基行列→1基行）
                foreach (var kv in sd.Hyperlinks)
                {
                    var (r, c) = kv.Key;
                    AddHyperlink(sheet, r + 1, c, kv.Value.Url, kv.Value.Display);
                }

                // 图片
                foreach (var img in sd.Images)
                {
                    AddImage(sheet, img.Row, img.Col, img.Data, img.Extension, img.Width, img.Height,
                        img.FromColOff, img.FromRowOff, img.ToRow, img.ToCol, img.ToColOff, img.ToRowOff, img.EditAs);
                }

                // 页面设置
                if (sd.Orientation != ExcelPageOrientation.Portrait || sd.PaperSize != ExcelPaperSize.Default)
                    SetPageSetup(sheet, sd.Orientation, sd.PaperSize);
                SetPageMargins(sheet, sd.MarginTop, sd.MarginBottom, sd.MarginLeft, sd.MarginRight);
                if (!sd.HeaderText.IsNullOrEmpty() || !sd.FooterText.IsNullOrEmpty())
                    SetHeaderFooter(sheet, sd.HeaderText, sd.FooterText);
                if (sd.PrintTitleStartRow > 0)
                    SetPrintTitleRows(sheet, sd.PrintTitleStartRow, sd.PrintTitleEndRow);

                // 工作表保护
                if (sd.ProtectionPassword != null)
                    ProtectSheet(sheet, sd.ProtectionPassword);

                // 条件格式
                foreach (var cf in sd.ConditionalFormats)
                {
                    AddConditionalFormat(sheet, cf.Range, cf.Type, cf.Value, cf.Color, cf.Value2);
                }

                // 批注（0基→1基行）
                foreach (var kv in sd.Comments)
                {
                    var (r, c) = kv.Key;
                    AddComment(sheet, r + 1, c, kv.Value.Text, kv.Value.Author);
                }

                // 数据验证
                foreach (var v in sd.Validations)
                {
                    if (v.Items != null && v.Items.Length > 0)
                        AddDropdownValidation(sheet, v.CellRange, v.Items);
                    else if (!v.ValidationType.IsNullOrEmpty())
                        AddRangeValidation(sheet, v.CellRange, v.ValidationType!, v.Operator ?? "between", v.Formula1 ?? "0", v.Formula2);
                }

                // 结构化表格
                foreach (var tbl in sd.Tables)
                {
                    AddTable(sheet, tbl.Range, tbl.Name, tbl.StyleName, tbl.ColumnNames);
                }

                // 图表
                foreach (var ch in sd.Charts)
                {
                    AddChart(sheet, ch);
                }
            }

            // 用户自定义命名范围
            foreach (var kv in data.DefinedNames)
            {
                AddDefinedName(kv.Key, kv.Value);
            }
        }
        finally
        {
            AutoFitColumnWidth = autoFit;
        }
    }
    #endregion
}