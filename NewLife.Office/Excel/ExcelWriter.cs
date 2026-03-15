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
}