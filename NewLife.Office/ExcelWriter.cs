using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Security;
using System.Text;
using NewLife.Collections;

namespace NewLife.Office;

/// <summary>轻量级Excel写入器，支持多个工作表</summary>
/// <remarks>
/// 目标：快速导出简单数据，支持多工作表的列头与多行数据；识别常见数据类型并使用合适样式，避免长数字（如身份证、长整型）被 Excel / WPS 显示为科学计数。
/// 支持单元格样式（字体/填充/边框/对齐）、合并单元格、冻结窗格、自动筛选、超链接、数据验证、图片、页面设置、条件格式等功能。
/// </remarks>
public class ExcelWriter : DisposeBase
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

    private record FontEntry(String? Name, Double Size, Boolean Bold, Boolean Italic, Boolean Underline, String? Color);
    private record FillEntry(String? BgColor, String PatternType);
    private record BorderEntry(CellBorderStyle Style, String? Color);
    private record XfEntry(Int32 NumFmtId, Int32 FontId, Int32 FillId, Int32 BorderId, HorizontalAlignment HAlign, VerticalAlignment VAlign, Boolean WrapText);

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
    }

    private class SheetPageSetup
    {
        public PageOrientation Orientation { get; set; }
        public PaperSize PaperSize { get; set; }
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
        public ConditionalFormatType Type { get; set; }
        public String? Value { get; set; }
        public String? Value2 { get; set; }
        public String? Color { get; set; }
    }

    private class SheetComment
    {
        public Int32 Row { get; set; }   // 1-based
        public Int32 Col { get; set; }   // 0-based
        public String Text { get; set; } = null!;
        public String Author { get; set; } = String.Empty;
    }
    #endregion

    #region 属性
    /// <summary>文件路径（Save 时写入）</summary>
    public String? FileName { get; }

    /// <summary>目标流（若提供则写入该流，调用方负责生命周期）</summary>
    public Stream? Stream { get; }

    /// <summary>默认工作表名称（当调用 API 未指定 sheet 时使用）</summary>
    public String SheetName { get; set; } = "Sheet1";

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
    private readonly List<FontEntry> _fonts = [new(null, 0, false, false, false, null)]; // index 0 = 默认字体
    private readonly List<FillEntry> _fills = [new(null, "none"), new(null, "gray125")]; // 0=none, 1=gray125 (Excel 要求)
    private readonly List<BorderEntry> _borders = [new(CellBorderStyle.None, null)]; // index 0 = 无边框
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
    // 条件格式
    private readonly Dictionary<String, List<ConditionalFormatEntry>> _sheetCondFormats = new(StringComparer.OrdinalIgnoreCase);
    // 批注
    private readonly Dictionary<String, List<SheetComment>> _sheetComments = new(StringComparer.OrdinalIgnoreCase);
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

    /// <summary>销毁释放</summary>
    /// <param name="disposing"></param>
    protected override void Dispose(Boolean disposing)
    {
        base.Dispose(disposing);
        if (Stream == null) Save();
    }

    private static List<XfEntry> InitBuiltinXfEntries()
    {
        // 按 _cellStyles 枚举值升序，生成内置 XF 条目（全使用默认字体/填充/边框）
        var list = new List<XfEntry>();
        foreach (var st in _cellStyles)
        {
            list.Add(new XfEntry((Int32)st, 0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Top, false));
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
    public void WriteHeader(String sheet, IEnumerable<String> headers, CellStyle? style)
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
    public void WriteRows(String? sheet, IEnumerable<Object?[]> data, CellStyle? style)
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
    public void WriteRow(String? sheet, Object?[] values, CellStyle? style = null)
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
        if (imageData == null || imageData.Length == 0) throw new ArgumentNullException(nameof(imageData));
        if (sheet.IsNullOrEmpty()) sheet = SheetName;
        EnsureSheet(sheet);

        if (!_sheetImages.TryGetValue(sheet, out var list))
        {
            list = [];
            _sheetImages[sheet] = list;
        }
        list.Add(new SheetImage { Row = row, Col = col, Data = imageData, Extension = extension.ToLower().TrimStart('.'), Width = widthPx, Height = heightPx });
    }
    #endregion

    #region 页面设置
    /// <summary>设置页面方向和纸张大小</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="orientation">页面方向</param>
    /// <param name="paperSize">纸张大小</param>
    public void SetPageSetup(String? sheet, PageOrientation orientation, PaperSize paperSize = PaperSize.A4)
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
    #endregion

    #region 条件格式
    /// <summary>添加条件格式</summary>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="range">应用范围（如 "A1:A100"）</param>
    /// <param name="type">条件类型</param>
    /// <param name="value">条件值</param>
    /// <param name="color">满足条件时的背景色（RGB十六进制）</param>
    /// <param name="value2">第二个条件值（仅 Between 类型使用）</param>
    public void AddConditionalFormat(String? sheet, String range, ConditionalFormatType type, String? value, String? color, String? value2 = null)
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
    #endregion

    #region 对象映射
    /// <summary>将对象集合导出到工作表</summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="sheet">工作表名称（可空）</param>
    /// <param name="data">对象集合</param>
    /// <param name="headerStyle">表头样式</param>
    public void WriteObjects<T>(String? sheet, IEnumerable<T> data, CellStyle? headerStyle = null) where T : class
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
    public void WriteDataTable(String? sheet, DataTable table, CellStyle? headerStyle = null)
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

    #region 内部写入
    private void EnsureSheet(String sheet)
    {
        if (!_sheetRows.ContainsKey(sheet))
        {
            _sheetRows[sheet] = [];
            _sheetRowIndex[sheet] = 0;
            _sheetNames.Add(sheet);
            _sheetColWidths[sheet] = [];
        }
    }

    private void AddRow(String sheet, Object?[]? values, CellStyle? rowStyle = null)
    {
        EnsureSheet(sheet);

        var rowIndex = ++_sheetRowIndex[sheet];
        values ??= [];

        var sb = Pool.StringBuilder.Get();
        sb.Append("<row r=\"").Append(rowIndex).Append("\">");

        for (var i = 0; i < values.Length; i++)
        {
            var val = values[i];
            if (val == null) continue; // 缺失列：解析时自动补 null

            var cellRef = GetColumnName(i) + rowIndex; // A1 / B2 ...

            // 公式快捷路径
            if (val is ExcelFormula fval)
            {
                var fxml = SecurityElement.Escape(fval.Formula) ?? fval.Formula;
                sb.Append("<c r=\"").Append(cellRef).Append('"');
                String? fType = null;
                String fInner;
                switch (fval.CachedValue)
                {
                    case Boolean b:
                        fType = "b";
                        fInner = b ? "1" : "0";
                        break;
                    case String str:
                        fType = "str";
                        fInner = SecurityElement.Escape(str) ?? str;
                        break;
                    case null:
                        fInner = String.Empty;
                        break;
                    default:
                        fInner = Convert.ToString(fval.CachedValue, CultureInfo.InvariantCulture) ?? String.Empty;
                        break;
                }
                if (fType != null) sb.Append(" t=\"").Append(fType).Append('"');
                sb.Append("><f>").Append(fxml).Append("</f><v>").Append(fInner).Append("</v></c>");
                continue;
            }

            // 识别类型
            var autoStyle = ExcelCellStyle.General;
            String? tAttr = null; // t="s" / "b"
            String? inner = null; // <v>值</v>
            var displayLen = 0;   // 估算显示长度用于列宽

            switch (val)
            {
                case String str:
                    {
                        // 百分比：形如 "12.3%" / "45%"
                        if (str.Length > 0 && str.EndsWith("%") && TryParsePercent(str, out var pct))
                        {
                            autoStyle = ExcelCellStyle.Percent;
                            inner = (pct / 100).ToString("0.##########", CultureInfo.InvariantCulture);
                            //displayLen = inner.Length + 1;
                            break;
                        }
                        else
                        {
                            // 普通字符串走共享字符串，减少体积 & 避免被推断
                            tAttr = "s";
                            inner = GetSharedStringIndex(str).ToString();
                        }
                        break;
                    }
                case Boolean b:
                    {
                        tAttr = "b";
                        inner = b ? "1" : "0";
                        //displayLen = 5;
                        break;
                    }
                case DateTime dt:
                    {
                        var baseDate = new DateTime(1900, 1, 1);
                        if (dt < baseDate)
                        {
                            // Excel 无法表示 1900-01-01 之前（或无效）日期，这里写入空字符串
                            tAttr = "s";
                            inner = GetSharedStringIndex(String.Empty).ToString();
                            break;
                        }
                        // Excel 序列值：1=1900/1/1（含闰年Bug），读取时减2，这里写入需补2
                        var serial = (dt - baseDate).TotalDays + 2; // 包含时间小数
                        var hasTime = dt.TimeOfDay.Ticks != 0;
                        autoStyle = hasTime ? ExcelCellStyle.DateTime : ExcelCellStyle.Date;
                        inner = serial.ToString("0.###############", CultureInfo.InvariantCulture);
                        // 为避免 WPS 显示 ########，这里按常见完整格式长度估算：yyyy-MM-dd 或 yyyy-MM-dd HH:mm:ss
                        //displayLen = hasTime ? 16 - 1 : 10 - 1;
                        displayLen = hasTime ? 14 : 0;
                        break;
                    }
                case TimeSpan ts:
                    autoStyle = ExcelCellStyle.Time;
                    inner = ts.TotalDays.ToString("0.###############", CultureInfo.InvariantCulture);
                    //displayLen = inner.Length;
                    break;
                case Int16 or Int32 or Int64 or Byte or SByte or UInt16 or UInt32 or UInt64:
                    {
                        // 如果太长，为了避免出现科学计数法，改用字符串表示
                        var numStr = Convert.ToString(val, CultureInfo.InvariantCulture)!;
                        if (ShouldWriteAsText(numStr, 15))
                        {
                            tAttr = "s";
                            inner = GetSharedStringIndex(numStr).ToString();
                        }
                        else
                        {
                            autoStyle = ExcelCellStyle.Integer;
                            inner = numStr; // 使用 General，避免两位截断
                        }
                        displayLen = numStr.Length < 8 ? 0 : numStr.Length;
                        break;
                    }
                case Decimal dec:
                    {
                        var numStr = dec.ToString(CultureInfo.InvariantCulture);
                        if (ShouldWriteAsText(numStr, LongNumberAsTextThreshold))
                        {
                            tAttr = "s";
                            inner = GetSharedStringIndex(numStr).ToString();
                        }
                        else
                        {
                            inner = numStr; // 使用 General，避免两位截断
                        }
                        displayLen = numStr.Length < 8 ? 0 : numStr.Length;
                        break;
                    }
                case Double d:
                    {
                        var numStr = d.ToString("0.###############", CultureInfo.InvariantCulture);
                        if (ShouldWriteAsText(numStr, LongNumberAsTextThreshold))
                        {
                            tAttr = "s";
                            inner = GetSharedStringIndex(numStr).ToString();
                        }
                        else
                        {
                            inner = numStr; // General
                        }
                        displayLen = numStr.Length < 8 ? 0 : numStr.Length;
                        break;
                    }
                case Single f:
                    {
                        var numStr = f.ToString("0.###############", CultureInfo.InvariantCulture);
                        if (ShouldWriteAsText(numStr, LongNumberAsTextThreshold))
                        {
                            tAttr = "s";
                            inner = GetSharedStringIndex(numStr).ToString();
                        }
                        else
                        {
                            inner = numStr; // General
                        }
                        displayLen = numStr.Length < 8 ? 0 : numStr.Length;
                        break;
                    }
                default:
                    {
                        // 其它类型调用 ToString() 后按字符串处理
                        var str = val + "";
                        tAttr = "s";
                        inner = GetSharedStringIndex(str).ToString();
                        break;
                    }
            }

            // 计算最终 XF 索引
            var sIndex = -1;
            if (rowStyle != null)
            {
                // 用户指定了样式：合并自动检测的 numFmtId 与用户样式的字体/填充/边框/对齐
                var numFmtId = (Int32)autoStyle;
                // 如果用户样式指定了自定义数字格式，则覆盖自动检测
                if (!rowStyle.NumberFormat.IsNullOrEmpty())
                    numFmtId = GetOrCreateNumFmt(rowStyle.NumberFormat!);
                sIndex = GetOrCreateXf(rowStyle, numFmtId);
            }
            else if (tAttr == null)
            {
                // 无用户样式、非字符串/布尔：使用内置样式
                sIndex = Array.IndexOf(_cellStyles, autoStyle);
            }

            sb.Append("<c r=\"").Append(cellRef).Append('"');
            if (tAttr != null) sb.Append(' ').Append("t=\"").Append(tAttr).Append('"');
            if (sIndex >= 0) sb.Append(' ').Append("s=\"").Append(sIndex).Append('"');
            sb.Append("><v>").Append(inner).Append("</v></c>");

            // 自动列宽
            if (AutoFitColumnWidth && displayLen > 0)
            {
                var list = _sheetColWidths[sheet];
                while (list.Count <= i) list.Add(0);
                // Excel 列宽：字符数 + 2 边距（粗略），限制最大值适度（如 80）
                var w = displayLen + 2; // 经验值
                if (w > 80) w = 80;
                if (w > list[i]) list[i] = w;
            }
        }

        sb.Append("</row>");
        _sheetRows[sheet].Add(sb.Return(true));
    }

    /// <summary>判断一个数值字符串是否应转为文本以避免被 Excel 自动显示为科学计数法。</summary>
    private static Boolean ShouldWriteAsText(String numStr, Int32 maxLength)
    {
        if (numStr.IsNullOrEmpty()) return false;

        var digits = 0;
        for (var i = 0; i < numStr.Length; i++)
        {
            var ch = numStr[i];
            if (ch >= '0' && ch <= '9') digits++;
        }
        if (digits > maxLength) return true;         // 有效数字过长（>11）
        if (numStr.StartsWith("0.0000000")) return true;            // 很小的数值（大量前导0）
        return false;
    }

    private static Boolean TryParsePercent(String str, out Decimal value)
    {
        value = 0m;
        var txt = str.Trim().TrimEnd('%');
        if (Decimal.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) { value = d; return true; }
        return false;
    }

    private Int32 GetSharedStringIndex(String str)
    {
        _sharedCount++;
        if (_shared.TryGetValue(str, out var idx)) return idx;
        idx = _shared.Count;
        _shared[str] = idx;
        return idx;
    }

    private static String GetColumnName(Int32 index)
    {
        // 0 -> A
        index++; // 转为 1 基
        var sb = Pool.StringBuilder.Get();
        while (index > 0)
        {
            var mod = (index - 1) % 26;
            sb.Insert(0, (Char)('A' + mod));
            index = (index - 1) / 26;
        }
        return sb.Return(true);
    }
    #endregion

    #region 样式管理
    /// <summary>根据用户样式和数字格式，查找或创建 XF 条目并返回索引</summary>
    private Int32 GetOrCreateXf(CellStyle cs, Int32 numFmtId)
    {
        // 找或创建字体
        var font = new FontEntry(cs.FontName, cs.FontSize, cs.Bold, cs.Italic, cs.Underline, cs.FontColor);
        var fontId = FindOrAdd(_fonts, font);

        // 找或创建填充
        var fillId = 0;
        if (!cs.BackgroundColor.IsNullOrEmpty())
        {
            var fill = new FillEntry(cs.BackgroundColor, "solid");
            fillId = FindOrAdd(_fills, fill);
        }

        // 找或创建边框
        var borderId = 0;
        if (cs.Border != CellBorderStyle.None)
        {
            var border = new BorderEntry(cs.Border, cs.BorderColor);
            borderId = FindOrAdd(_borders, border);
        }

        // 复合键去重
        var key = $"{numFmtId}-{fontId}-{fillId}-{borderId}-{(Int32)cs.HAlign}-{(Int32)cs.VAlign}-{(cs.WrapText ? 1 : 0)}";
        if (_xfCache.TryGetValue(key, out var idx)) return idx;

        var xf = new XfEntry(numFmtId, fontId, fillId, borderId, cs.HAlign, cs.VAlign, cs.WrapText);
        idx = _xfEntries.Count;
        _xfEntries.Add(xf);
        _xfCache[key] = idx;
        return idx;
    }

    /// <summary>获取或创建自定义数字格式</summary>
    private Int32 GetOrCreateNumFmt(String formatCode)
    {
        if (_numFmtMap.TryGetValue(formatCode, out var id)) return id;
        id = _nextNumFmtId++;
        _numFmtMap[formatCode] = id;
        return id;
    }

    private static Int32 FindOrAdd<T>(List<T> list, T item) where T : notnull
    {
        for (var i = 0; i < list.Count; i++)
        {
            if (list[i].Equals(item)) return i;
        }
        list.Add(item);
        return list.Count - 1;
    }

    /// <summary>解析单元格引用（如 "A1"）返回 (行0基, 列0基)</summary>
    private static (Int32 Row, Int32 Col) ParseCellRef(String cellRef)
    {
        var colLen = 0;
        for (var i = 0; i < cellRef.Length; i++)
        {
            var ch = cellRef[i];
            if (ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z') colLen++;
            else break;
        }

        var colIndex = 0;
        for (var i = 0; i < colLen; i++)
        {
            var ch = cellRef[i];
            if (ch is >= 'a' and <= 'z') ch = (Char)(ch - 'a' + 'A');
            colIndex = colIndex * 26 + (ch - 'A' + 1);
        }
        colIndex--; // 转 0 基

        var rowStr = cellRef.Substring(colLen);
        var rowIndex = Int32.Parse(rowStr) - 1; // 转 0 基

        return (rowIndex, colIndex);
    }

    /// <summary>生成单元格引用（如 "A1"），行列均为 0 基</summary>
    private static String MakeCellRef(Int32 row, Int32 col) => GetColumnName(col) + (row + 1);

    /// <summary>获取边框 OOXML 样式名</summary>
    private static String GetBorderStyleName(CellBorderStyle style) => style switch
    {
        CellBorderStyle.Thin => "thin",
        CellBorderStyle.Medium => "medium",
        CellBorderStyle.Thick => "thick",
        CellBorderStyle.Dashed => "dashed",
        CellBorderStyle.Dotted => "dotted",
        CellBorderStyle.DoubleLine => "double",
        _ => "thin",
    };
    #endregion

    #region 保存
    /// <summary>保存到文件或目标流</summary>
    public void Save()
    {
        // 若未写任何 sheet，创建一个空的默认工作表，避免生成非法 workbook
        if (_sheetNames.Count == 0) EnsureSheet(SheetName);

        var target = Stream;
        if (target == null)
        {
            if (FileName.IsNullOrEmpty()) throw new InvalidOperationException("未指定输出位置");

            var file = FileName.EnsureDirectory(true).GetFullPath();
            target = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
        }

        // 判断哪些 sheet 有图片
        var sheetsWithImages = new HashSet<Int32>();
        var globalImageIndex = 0;
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetImages.TryGetValue(_sheetNames[i], out var imgs) && imgs.Count > 0)
                sheetsWithImages.Add(i);
        }

        // 判断哪些 sheet 有超链接
        var sheetsWithHyperlinks = new HashSet<Int32>();
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetHyperlinks.TryGetValue(_sheetNames[i], out var links) && links.Count > 0)
                sheetsWithHyperlinks.Add(i);
        }

        // 判断哪些 sheet 需要打印标题行
        var sheetsWithPrintTitles = new HashSet<Int32>();
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetPageSetups.TryGetValue(_sheetNames[i], out var ps) && ps.PrintTitleStartRow > 0)
                sheetsWithPrintTitles.Add(i);
        }

        // 判断哪些 sheet 有批注
        var sheetsWithComments = new HashSet<Int32>();
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetComments.TryGetValue(_sheetNames[i], out var cmts) && cmts.Count > 0)
                sheetsWithComments.Add(i);
        }

        using var za = new ZipArchive(target, ZipArchiveMode.Create, leaveOpen: Stream != null, entryNameEncoding: Encoding);

        // _rels/.rels
        using (var sw = new StreamWriter(za.CreateEntry("_rels/.rels").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        }

        // [Content_Types].xml
        using (var sw = new StreamWriter(za.CreateEntry("[Content_Types].xml").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sw.Write("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                sw.Write("<Override PartName=\"/xl/worksheets/sheet");
                sw.Write(i + 1);
                sw.Write(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }
            if (_shared.Count > 0)
            {
                sw.Write("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            }
            sw.Write("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            // 图片类型
            var imageExts = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in _sheetImages)
            {
                foreach (var img in kv.Value)
                {
                    imageExts.Add(img.Extension);
                }
            }
            foreach (var ext in imageExts)
            {
                var mime = ext == "png" ? "image/png" : ext == "jpeg" || ext == "jpg" ? "image/jpeg" : ext == "gif" ? "image/gif" : "image/png";
                sw.Write($"<Default Extension=\"{ext}\" ContentType=\"{mime}\"/>");
            }
            // Drawing
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                if (sheetsWithImages.Contains(i))
                    sw.Write($"<Override PartName=\"/xl/drawings/drawing{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
            }
            // 批注
            if (sheetsWithComments.Count > 0)
                sw.Write("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>");
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                if (sheetsWithComments.Contains(i))
                    sw.Write($"<Override PartName=\"/xl/comments{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>");
            }
            sw.Write("</Types>");
        }

        // workbook.xml
        using (var sw = new StreamWriter(za.CreateEntry("xl/workbook.xml").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                var name = SecurityElement.Escape(_sheetNames[i]) ?? _sheetNames[i];
                sw.Write($"<sheet name=\"{name}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
            }
            sw.Write("</sheets>");
            // 打印标题行的 definedNames
            if (sheetsWithPrintTitles.Count > 0)
            {
                sw.Write("<definedNames>");
                foreach (var si in sheetsWithPrintTitles)
                {
                    var ps = _sheetPageSetups[_sheetNames[si]];
                    var sn = SecurityElement.Escape(_sheetNames[si]) ?? _sheetNames[si];
                    sw.Write($"<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"{si}\">'{sn}'!${ ps.PrintTitleStartRow}:${ps.PrintTitleEndRow}</definedName>");
                }
                sw.Write("</definedNames>");
            }
            sw.Write("</workbook>");
        }

        // workbook 关系
        using (var sw = new StreamWriter(za.CreateEntry("xl/_rels/workbook.xml.rels").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (var i = 0; i < _sheetNames.Count; i++) sw.Write($"<Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
            var nextId = _sheetNames.Count + 1;
            sw.Write($"<Relationship Id=\"rId{nextId++}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            if (_shared.Count > 0) sw.Write($"<Relationship Id=\"rId{nextId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            sw.Write("</Relationships>");
        }

        // styles.xml（完整版：numFmts + fonts + fills + borders + cellXfs）
        WriteStylesXml(za);

        // sharedStrings.xml
        if (_shared.Count > 0)
        {
            using var sw = new StreamWriter(za.CreateEntry("xl/sharedStrings.xml").Open(), Encoding);
            sw.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{_sharedCount}\" uniqueCount=\"{_shared.Count}\">");
            foreach (var kv in _shared.OrderBy(e => e.Value))
            {
                var txt = SecurityElement.Escape(kv.Key) ?? String.Empty;
                sw.Write("<si><t>");
                sw.Write(txt);
                sw.Write("</t></si>");
            }
            sw.Write("</sst>");
        }

        // worksheets
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            var sheet = _sheetNames[i];
            var entry = za.CreateEntry($"xl/worksheets/sheet{i + 1}.xml");
            using var sw = new StreamWriter(entry.Open(), Encoding);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:etc=\"http://www.wps.cn/officeDocument/2017/etCustomData\">");

            // sheetViews（冻结窗格）
            if (_sheetFreezes.TryGetValue(sheet, out var freeze) && (freeze.Rows > 0 || freeze.Cols > 0))
            {
                sw.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">");
                var topLeft = MakeCellRef(freeze.Rows, freeze.Cols);
                if (freeze.Rows > 0 && freeze.Cols > 0)
                {
                    sw.Write($"<pane xSplit=\"{freeze.Cols}\" ySplit=\"{freeze.Rows}\" topLeftCell=\"{topLeft}\" activePane=\"bottomRight\" state=\"frozen\"/>");
                }
                else if (freeze.Rows > 0)
                {
                    sw.Write($"<pane ySplit=\"{freeze.Rows}\" topLeftCell=\"{topLeft}\" activePane=\"bottomLeft\" state=\"frozen\"/>");
                }
                else
                {
                    sw.Write($"<pane xSplit=\"{freeze.Cols}\" topLeftCell=\"{topLeft}\" activePane=\"topRight\" state=\"frozen\"/>");
                }
                sw.Write("</sheetView></sheetViews>");
            }

            // cols（列宽）
            if (AutoFitColumnWidth && _sheetColWidths.TryGetValue(sheet, out var widths) && widths.Count > 0)
            {
                if (widths.Any(e => e > 0))
                {
                    sw.Write("<cols>");
                    for (var c = 0; c < widths.Count; c++)
                    {
                        var w = widths[c];
                        if (w <= 0) continue;
                        sw.Write($"<col min=\"{c + 1}\" max=\"{c + 1}\" width=\"{w:0.##}\" customWidth=\"1\"/>");
                    }
                    sw.Write("</cols>");
                }
            }

            // sheetData（带行高注入）
            sw.Write("<sheetData>");
            if (_sheetRows.TryGetValue(sheet, out var list))
            {
                var hasHeights = _sheetRowHeights.TryGetValue(sheet, out var heights) && heights.Count > 0;
                var rowNum = 1;
                foreach (var r in list)
                {
                    if (hasHeights && heights!.TryGetValue(rowNum, out var ht))
                    {
                        sw.Write(r.Replace($"<row r=\"{rowNum}\"", $"<row r=\"{rowNum}\" ht=\"{ht:0.##}\" customHeight=\"1\""));
                    }
                    else
                    {
                        sw.Write(r);
                    }
                    rowNum++;
                }
            }
            sw.Write("</sheetData>");

            // sheetProtection
            if (_sheetProtection.TryGetValue(sheet, out var pwd))
            {
                sw.Write("<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"");
                if (!pwd.IsNullOrEmpty())
                {
                    var hash = ComputeSheetProtectionHash(pwd);
                    sw.Write($" password=\"{hash}\"");
                }
                sw.Write("/>");
            }

            // autoFilter
            if (_sheetAutoFilters.TryGetValue(sheet, out var filter))
            {
                sw.Write($"<autoFilter ref=\"{filter}\"/>");
            }

            // mergeCells
            if (_sheetMerges.TryGetValue(sheet, out var merges) && merges.Count > 0)
            {
                sw.Write($"<mergeCells count=\"{merges.Count}\">");
                foreach (var (sr, sc, er, ec) in merges)
                {
                    sw.Write($"<mergeCell ref=\"{MakeCellRef(sr, sc)}:{MakeCellRef(er, ec)}\"/>");
                }
                sw.Write("</mergeCells>");
            }

            // conditionalFormatting
            if (_sheetCondFormats.TryGetValue(sheet, out var conds) && conds.Count > 0)
            {
                var priority = 1;
                foreach (var cf in conds)
                {
                    sw.Write($"<conditionalFormatting sqref=\"{cf.Range}\">");
                    switch (cf.Type)
                    {
                        case ConditionalFormatType.GreaterThan:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"greaterThan\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.LessThan:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"lessThan\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.Equal:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"equal\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.Between:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"between\"><formula>{SecurityElement.Escape(cf.Value)}</formula><formula>{SecurityElement.Escape(cf.Value2)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.DataBar:
                            sw.Write($"<cfRule type=\"dataBar\" priority=\"{priority++}\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF{cf.Color ?? "4472C4"}\"/></dataBar></cfRule>");
                            break;
                        case ConditionalFormatType.ColorScale:
                            sw.Write($"<cfRule type=\"colorScale\" priority=\"{priority++}\"><colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFFFFFFF\"/><color rgb=\"FF{cf.Color ?? "4472C4"}\"/></colorScale></cfRule>");
                            break;
                    }
                    sw.Write("</conditionalFormatting>");
                }
            }

            // dataValidations
            if (_sheetValidations.TryGetValue(sheet, out var validations) && validations.Count > 0)
            {
                sw.Write($"<dataValidations count=\"{validations.Count}\">");
                foreach (var v in validations)
                {
                    if (v.Items != null)
                    {
                        var formula = "\"" + String.Join(",", v.Items.Select(e => SecurityElement.Escape(e))) + "\"";
                        sw.Write($"<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"{v.CellRange}\"><formula1>{formula}</formula1></dataValidation>");
                    }
                    else if (!v.ValidationType.IsNullOrEmpty())
                    {
                        var op = v.Operator ?? "between";
                        sw.Write($"<dataValidation type=\"{v.ValidationType}\" operator=\"{op}\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"{v.CellRange}\">");
                        sw.Write($"<formula1>{SecurityElement.Escape(v.Formula1 ?? "0")}</formula1>");
                        if (!v.Formula2.IsNullOrEmpty()) sw.Write($"<formula2>{SecurityElement.Escape(v.Formula2!)}</formula2>");
                        sw.Write("</dataValidation>");
                    }
                }
                sw.Write("</dataValidations>");
            }

            // hyperlinks
            if (sheetsWithHyperlinks.Contains(i) && _sheetHyperlinks.TryGetValue(sheet, out var hyperlinks))
            {
                sw.Write("<hyperlinks>");
                for (var h = 0; h < hyperlinks.Count; h++)
                {
                    var hl = hyperlinks[h];
                    var cellRef = MakeCellRef(hl.Row - 1, hl.Col);
                    sw.Write($"<hyperlink ref=\"{cellRef}\" r:id=\"rHl{h + 1}\"");
                    if (!hl.Display.IsNullOrEmpty()) sw.Write($" display=\"{SecurityElement.Escape(hl.Display)}\"");
                    sw.Write("/>");
                }
                sw.Write("</hyperlinks>");
            }

            // pageMargins + pageSetup + headerFooter
            if (_sheetPageSetups.TryGetValue(sheet, out var pageSetup))
            {
                sw.Write($"<pageMargins left=\"{pageSetup.MarginLeft:0.##}\" right=\"{pageSetup.MarginRight:0.##}\" top=\"{pageSetup.MarginTop:0.##}\" bottom=\"{pageSetup.MarginBottom:0.##}\" header=\"0.3\" footer=\"0.3\"/>");
                var orient = pageSetup.Orientation == PageOrientation.Landscape ? "landscape" : "portrait";
                sw.Write($"<pageSetup orientation=\"{orient}\"");
                if (pageSetup.PaperSize != PaperSize.Default) sw.Write($" paperSize=\"{(Int32)pageSetup.PaperSize}\"");
                sw.Write("/>");
                if (!pageSetup.HeaderText.IsNullOrEmpty() || !pageSetup.FooterText.IsNullOrEmpty())
                {
                    sw.Write("<headerFooter>");
                    if (!pageSetup.HeaderText.IsNullOrEmpty()) sw.Write($"<oddHeader>{SecurityElement.Escape(pageSetup.HeaderText)}</oddHeader>");
                    if (!pageSetup.FooterText.IsNullOrEmpty()) sw.Write($"<oddFooter>{SecurityElement.Escape(pageSetup.FooterText)}</oddFooter>");
                    sw.Write("</headerFooter>");
                }
            }

            // drawing（图片引用）
            if (sheetsWithImages.Contains(i))
            {
                sw.Write($"<drawing r:id=\"rDr1\"/>");
            }

            // legacyDrawing（批注 VML 引用）
            if (sheetsWithComments.Contains(i))
            {
                sw.Write($"<legacyDrawing r:id=\"rVml1\"/>");
            }

            sw.Write("</worksheet>");
            sw.Dispose();

            // sheet rels（超链接 + 图片 drawing + 批注关系）
            if (sheetsWithHyperlinks.Contains(i) || sheetsWithImages.Contains(i) || sheetsWithComments.Contains(i))
            {
                var relEntry = za.CreateEntry($"xl/worksheets/_rels/sheet{i + 1}.xml.rels");
                using var rsw = new StreamWriter(relEntry.Open(), Encoding);
                rsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                if (sheetsWithHyperlinks.Contains(i) && _sheetHyperlinks.TryGetValue(sheet, out var rels))
                {
                    for (var h = 0; h < rels.Count; h++)
                    {
                        rsw.Write($"<Relationship Id=\"rHl{h + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{SecurityElement.Escape(rels[h].Url)}\" TargetMode=\"External\"/>");
                    }
                }
                if (sheetsWithImages.Contains(i))
                {
                    rsw.Write($"<Relationship Id=\"rDr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing{i + 1}.xml\"/>");
                }
                if (sheetsWithComments.Contains(i))
                {
                    rsw.Write($"<Relationship Id=\"rVml1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing{i + 1}.vml\"/>");
                    rsw.Write($"<Relationship Id=\"rCmt1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments{i + 1}.xml\"/>");
                }
                rsw.Write("</Relationships>");
            }
        }

        // Drawings 和媒体文件
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (!sheetsWithImages.Contains(i)) continue;
            var sheet = _sheetNames[i];
            var images = _sheetImages[sheet];

            // drawing{i+1}.xml
            var drawEntry = za.CreateEntry($"xl/drawings/drawing{i + 1}.xml");
            using (var dsw = new StreamWriter(drawEntry.Open(), Encoding))
            {
                dsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                for (var j = 0; j < images.Count; j++)
                {
                    var img = images[j];
                    var emuW = (Int64)(img.Width * 9525); // px → EMU
                    var emuH = (Int64)(img.Height * 9525);
                    dsw.Write($"<xdr:twoCellAnchor><xdr:from><xdr:col>{img.Col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{img.Row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>");
                    dsw.Write($"<xdr:to><xdr:col>{img.Col + 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{img.Row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>");
                    dsw.Write($"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{j + 2}\" name=\"Image{globalImageIndex + 1}\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>");
                    dsw.Write($"<xdr:blipFill><a:blip r:embed=\"rImg{j + 1}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
                    dsw.Write($"<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{emuW}\" cy=\"{emuH}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>");
                    globalImageIndex++;
                }
                dsw.Write("</xdr:wsDr>");
            }

            // drawing rels
            var drawRelEntry = za.CreateEntry($"xl/drawings/_rels/drawing{i + 1}.xml.rels");
            using (var drsw = new StreamWriter(drawRelEntry.Open(), Encoding))
            {
                drsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (var j = 0; j < images.Count; j++)
                {
                    drsw.Write($"<Relationship Id=\"rImg{j + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{globalImageIndex - images.Count + j + 1}.{images[j].Extension}\"/>");
                }
                drsw.Write("</Relationships>");
            }

            // 媒体文件
            for (var j = 0; j < images.Count; j++)
            {
                var img = images[j];
                var mediaEntry = za.CreateEntry($"xl/media/image{globalImageIndex - images.Count + j + 1}.{img.Extension}");
                using var ms2 = mediaEntry.Open();
                ms2.Write(img.Data, 0, img.Data.Length);
            }
        }

        // 批注文件：xl/commentsN.xml + xl/drawings/vmlDrawingN.vml
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (!sheetsWithComments.Contains(i)) continue;
            var sheet = _sheetNames[i];
            var comments = _sheetComments[sheet];

            // 收集所有不同作者（保持插入顺序，用 List 去重）
            var authors = new List<String>();
            foreach (var c in comments)
            {
                if (!authors.Contains(c.Author)) authors.Add(c.Author);
            }

            // xl/commentsN.xml
            using (var csw = new StreamWriter(za.CreateEntry($"xl/comments{i + 1}.xml").Open(), Encoding))
            {
                csw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                csw.Write("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                csw.Write("<authors>");
                foreach (var a in authors) csw.Write($"<author>{SecurityElement.Escape(a)}</author>");
                csw.Write("</authors><commentList>");
                foreach (var c in comments)
                {
                    var cellRef = MakeCellRef(c.Row - 1, c.Col);
                    var authorId = authors.IndexOf(c.Author);
                    csw.Write($"<comment ref=\"{cellRef}\" authorId=\"{authorId}\">");
                    csw.Write($"<text><r><t xml:space=\"preserve\">{SecurityElement.Escape(c.Text)}</t></r></text>");
                    csw.Write("</comment>");
                }
                csw.Write("</commentList></comments>");
            }

            // xl/drawings/vmlDrawingN.vml
            using (var vsw = new StreamWriter(za.CreateEntry($"xl/drawings/vmlDrawing{i + 1}.vml").Open(), Encoding))
            {
                vsw.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                vsw.Write("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
                vsw.Write("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m0,0l0,21600,21600,21600,21600,0xe\">");
                vsw.Write("<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");
                for (var j = 0; j < comments.Count; j++)
                {
                    var c = comments[j];
                    vsw.Write($"<v:shape id=\"_x0000_s{1025 + j}\" type=\"#_x0000_t202\" " +
                              "style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" " +
                              "fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
                    vsw.Write("<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
                    vsw.Write("<v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>");
                    vsw.Write("<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/>");
                    vsw.Write($"<x:Row>{c.Row - 1}</x:Row><x:Column>{c.Col}</x:Column>");
                    vsw.Write("</x:ClientData></v:shape>");
                }
                vsw.Write("</xml>");
            }
        }

        target.Flush();
    }

    /// <summary>生成完整的 styles.xml</summary>
    private void WriteStylesXml(ZipArchive za)
    {
        using var sw = new StreamWriter(za.CreateEntry("xl/styles.xml").Open(), Encoding);
        sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        // numFmts（自定义）
        if (_numFmtMap.Count > 0)
        {
            sw.Write($"<numFmts count=\"{_numFmtMap.Count}\">");
            foreach (var kv in _numFmtMap)
            {
                sw.Write($"<numFmt numFmtId=\"{kv.Value}\" formatCode=\"{SecurityElement.Escape(kv.Key)}\"/>");
            }
            sw.Write("</numFmts>");
        }

        // fonts
        sw.Write($"<fonts count=\"{_fonts.Count}\">");
        foreach (var f in _fonts)
        {
            sw.Write("<font>");
            if (f.Bold) sw.Write("<b/>");
            if (f.Italic) sw.Write("<i/>");
            if (f.Underline) sw.Write("<u/>");
            if (f.Size > 0) sw.Write($"<sz val=\"{f.Size}\"/>");
            if (!f.Color.IsNullOrEmpty()) sw.Write($"<color rgb=\"FF{f.Color}\"/>");
            if (!f.Name.IsNullOrEmpty()) sw.Write($"<name val=\"{SecurityElement.Escape(f.Name)}\"/>");
            sw.Write("</font>");
        }
        sw.Write("</fonts>");

        // fills
        sw.Write($"<fills count=\"{_fills.Count}\">");
        foreach (var f in _fills)
        {
            sw.Write("<fill>");
            if (f.PatternType == "none")
                sw.Write("<patternFill patternType=\"none\"/>");
            else if (f.PatternType == "gray125")
                sw.Write("<patternFill patternType=\"gray125\"/>");
            else
                sw.Write($"<patternFill patternType=\"solid\"><fgColor rgb=\"FF{f.BgColor}\"/></patternFill>");
            sw.Write("</fill>");
        }
        sw.Write("</fills>");

        // borders
        sw.Write($"<borders count=\"{_borders.Count}\">");
        foreach (var b in _borders)
        {
            if (b.Style == CellBorderStyle.None)
            {
                sw.Write("<border><left/><right/><top/><bottom/><diagonal/></border>");
            }
            else
            {
                var sn = GetBorderStyleName(b.Style);
                var ca = b.Color.IsNullOrEmpty() ? "" : $"<color rgb=\"FF{b.Color}\"/>";
                sw.Write($"<border><left style=\"{sn}\">{ca}</left><right style=\"{sn}\">{ca}</right><top style=\"{sn}\">{ca}</top><bottom style=\"{sn}\">{ca}</bottom><diagonal/></border>");
            }
        }
        sw.Write("</borders>");

        // cellXfs
        sw.Write($"<cellXfs count=\"{_xfEntries.Count}\">");
        foreach (var xf in _xfEntries)
        {
            sw.Write($"<xf numFmtId=\"{xf.NumFmtId}\" fontId=\"{xf.FontId}\" fillId=\"{xf.FillId}\" borderId=\"{xf.BorderId}\"");
            if (xf.FontId > 0) sw.Write(" applyFont=\"1\"");
            if (xf.FillId > 0) sw.Write(" applyFill=\"1\"");
            if (xf.BorderId > 0) sw.Write(" applyBorder=\"1\"");
            if (xf.NumFmtId > 0) sw.Write(" applyNumberFormat=\"1\"");
            if (xf.HAlign != HorizontalAlignment.General || xf.VAlign != VerticalAlignment.Top || xf.WrapText)
            {
                sw.Write(" applyAlignment=\"1\"><alignment");
                if (xf.HAlign != HorizontalAlignment.General) sw.Write($" horizontal=\"{xf.HAlign.ToString().ToLower()}\"");
                if (xf.VAlign != VerticalAlignment.Top) sw.Write($" vertical=\"{xf.VAlign.ToString().ToLower()}\"");
                if (xf.WrapText) sw.Write(" wrapText=\"1\"");
                sw.Write("/></xf>");
            }
            else
            {
                sw.Write("/>");
            }
        }
        sw.Write("</cellXfs>");

        // 条件格式需要的 dxf（差异格式）
        var totalDxf = 0;
        foreach (var kv in _sheetCondFormats)
        {
            foreach (var cf in kv.Value)
            {
                if (cf.Type < ConditionalFormatType.DataBar) totalDxf++;
            }
        }
        if (totalDxf > 0)
        {
            sw.Write($"<dxfs count=\"{totalDxf}\">");
            foreach (var kv in _sheetCondFormats)
            {
                foreach (var cf in kv.Value)
                {
                    if (cf.Type >= ConditionalFormatType.DataBar) continue;
                    sw.Write("<dxf>");
                    if (!cf.Color.IsNullOrEmpty())
                        sw.Write($"<fill><patternFill><bgColor rgb=\"FF{cf.Color}\"/></patternFill></fill>");
                    sw.Write("</dxf>");
                }
            }
            sw.Write("</dxfs>");
        }

        sw.Write("</styleSheet>");
    }

    /// <summary>计算工作表保护密码哈希（Excel 传统算法）</summary>
    private static String ComputeSheetProtectionHash(String password)
    {
        var hash = 0;
        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash ^= password[i];
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        }
        hash ^= password.Length;
        hash ^= 0xCE4B;
        return hash.ToString("X4");
    }
    #endregion
}
