using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Text;
using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>xls（BIFF8）格式写入器</summary>
/// <remarks>
/// 生成 Microsoft Excel 97-2003 二进制格式（BIFF8）的 .xls 文件，
/// 打包在 OLE2/CFB 容器中，无需外部依赖。
/// <para>支持多工作表、字符串/数值/日期/布尔/公式单元格写入，
/// 以及对象集合和 DataTable 的批量映射。</para>
/// <para>写入示例：</para>
/// <code>
/// using var writer = new BiffWriter();
/// writer.WriteHeader(new[] { "姓名", "年龄", "成绩" });
/// writer.WriteRow(new Object?[] { "Alice", 28, 95.5 });
/// writer.Save("data.xls");
/// </code>
/// </remarks>
public sealed class BiffWriter : IDisposable
{
    #region 常量

    private const UInt16 RecBof = 0x0809;
    private const UInt16 RecEof = 0x000A;
    private const UInt16 RecBoundSheet = 0x0085;
    private const UInt16 RecSst = 0x00FC;
    private const UInt16 RecDimensions = 0x0200;
    private const UInt16 RecRow = 0x0208;
    private const UInt16 RecLabelSst = 0x00FD;
    private const UInt16 RecNumber = 0x0203;
    private const UInt16 RecBoolErr = 0x0205;
    private const UInt16 RecBlank = 0x0201;
    private const UInt16 RecXf = 0x00E0;
    private const UInt16 RecFont = 0x0031;
    private const UInt16 RecFormat = 0x041E;
    private const UInt16 RecContinue = 0x003C;
    private const UInt16 RecColInfo = 0x007D;
    private const UInt16 RecFormula = 0x0006;
    private const UInt16 RecString = 0x0207;
    private const UInt16 RecWindow2 = 0x023E;
    private const Int32 MaxRecordDataSize = 8224;

    // BIFF8 日期纪元：1900-01-01（含 1900 闰年兼容性偏移 +1）
    private static readonly DateTime DateEpoch = new(1900, 1, 1);
    private const Int32 DateEpochOffset = 2; // Excel 的 1900 闰年兼容 bug

    #endregion

    #region 属性

    /// <summary>当前活动工作表名称</summary>
    public String SheetName
    {
        get => _currentSheet;
        set
        {
            if (!_sheetData.ContainsKey(value))
            {
                _sheetNames.Add(value);
                _sheetData[value] = [];
            }
            _currentSheet = value;
        }
    }

    #endregion

    #region 私有字段

    private readonly List<String> _sheetNames = [];
    private readonly Dictionary<String, List<(List<Object?> Values, ExcelCellStyle? Style)>> _sheetData = new(StringComparer.Ordinal);

    // 共享字符串表
    private readonly List<String> _sst = [];
    private readonly Dictionary<String, Int32> _sstIndex = new(StringComparer.Ordinal);

    // 列宽：Key = sheetName, Value = (colIndex → width)
    private readonly Dictionary<String, Dictionary<Int32, Int32>> _sheetColWidths = new(StringComparer.Ordinal);

    // 列数字格式：Key = sheetName, Value = (colIndex → formatString)
    private readonly Dictionary<String, Dictionary<Int32, String>> _columnFormats = new(StringComparer.Ordinal);

    // 冻结窗格：Key = sheetName, Value = (freezeRow, freezeCol)
    private readonly Dictionary<String, (Int32 Row, Int32 Col)> _sheetFreezePanes = new(StringComparer.Ordinal);

    private String _currentSheet = "Sheet1";
    private Boolean _disposed;

    #endregion

    #region 构造

    /// <summary>创建新的 xls 写入器</summary>
    public BiffWriter()
    {
        _sheetNames.Add(_currentSheet);
        _sheetData[_currentSheet] = [];
    }

    /// <summary>释放资源</summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }

    #endregion

    #region 写入方法

    /// <summary>写入标题行（字符串数组）</summary>
    /// <param name="headers">列标题</param>
    public void WriteHeader(IEnumerable<String> headers)
    {
        WriteRow(headers.Cast<Object?>());
    }

    /// <summary>写入一行数据</summary>
    /// <param name="values">单元格值序列（支持 String/Int32/Double/DateTime/Boolean/null）</param>
    public void WriteRow(IEnumerable<Object?> values)
    {
        WriteRow(values, null);
    }

    /// <summary>写入一行数据（带样式）</summary>
    /// <param name="values">单元格值序列</param>
    /// <param name="style">行级单元格样式（字体/填充/边框），null 使用默认样式</param>
    public void WriteRow(IEnumerable<Object?> values, ExcelCellStyle? style)
    {
        var sheet = GetCurrentSheet();
        sheet.Add((values.ToList(), style));
    }

    /// <summary>将对象集合写入当前工作表（第一行为属性名标题）</summary>
    /// <typeparam name="T">对象类型</typeparam>
    /// <param name="data">对象集合</param>
    public void WriteObjects<T>(IEnumerable<T> data) where T : class
    {
        var props = GetMappableProperties<T>();
        var headers = props.Select(GetPropertyDisplayName).ToArray();
        WriteHeader(headers);

        foreach (var obj in data)
        {
            var row = props.Select(p =>
            {
                var val = p.GetValue(obj);
                return val;
            }).Cast<Object?>();
            WriteRow(row);
        }
    }

    /// <summary>将 DataTable 写入当前工作表（第一行为列名标题）</summary>
    /// <param name="table">数据表</param>
    public void WriteDataTable(DataTable table)
    {
        WriteHeader(table.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        foreach (DataRow row in table.Rows)
        {
            WriteRow(row.ItemArray.Cast<Object?>());
        }
    }

    /// <summary>设置当前工作表中指定列的宽度</summary>
    /// <param name="columnIndex">列索引（0基）</param>
    /// <param name="width">列宽（单位：字符宽度，约等于默认字体字符宽度）</param>
    /// <remarks>
    /// 列宽以最大字符宽度（256分之一的字符宽度）为单位存储。
    /// 例如 width=10 表示约 10 个字符宽度，内部存储为 10*256=2560。
    /// </remarks>
    public void SetColumnWidth(Int32 columnIndex, Double width)
    {
        if (!_sheetColWidths.TryGetValue(_currentSheet, out var colMap))
        {
            colMap = [];
            _sheetColWidths[_currentSheet] = colMap;
        }
        // BIFF8 COLINFO 使用 1/256 字符宽度为单位
        colMap[columnIndex] = (Int32)(width * 256);
    }

    /// <summary>设置当前工作表中指定列的数字格式</summary>
    /// <param name="columnIndex">列索引（0基）</param>
    /// <param name="format">Excel 数字格式字符串（如 yyyy-mm-dd、#,##0.00、0%）</param>
    public void SetColumnNumberFormat(Int32 columnIndex, String format)
    {
        if (!_columnFormats.TryGetValue(_currentSheet, out var colMap))
        {
            colMap = [];
            _columnFormats[_currentSheet] = colMap;
        }
        colMap[columnIndex] = format;
    }

    /// <summary>设置当前工作表的冻结窗格</summary>
    /// <param name="freezeRow">冻结行数（0 不冻结行），标题行下方的行数</param>
    /// <param name="freezeCol">冻结列数（0 不冻结列），左侧的列数</param>
    /// <remarks>例如 SetFreezePane(1, 0) 冻结首行；SetFreezePane(1, 2) 冻结首行和前两列</remarks>
    public void SetFreezePane(Int32 freezeRow, Int32 freezeCol)
    {
        _sheetFreezePanes[_currentSheet] = (freezeRow, freezeCol);
    }

    #endregion

    #region 保存

    /// <summary>将 xls 数据保存到指定文件</summary>
    /// <param name="path">目标文件路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs);
    }

    /// <summary>将 xls 数据写入流</summary>
    /// <param name="stream">可写输出流</param>
    public void Save(Stream stream)
    {
        BuildSstIndex();
        var workbookBytes = BuildWorkbookStream();

        var doc = new CfbDocument();
        doc.PutStream("Workbook", workbookBytes);
        doc.Save(stream);
    }

    /// <summary>将 xls 数据序列化为字节数组</summary>
    /// <returns>xls 格式的字节数组</returns>
    public Byte[] ToBytes()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }

    #endregion

    #region BIFF8 流构建

    private void BuildSstIndex()
    {
        _sst.Clear();
        _sstIndex.Clear();

        foreach (var sheetName in _sheetNames)
        {
            if (!_sheetData.TryGetValue(sheetName, out var rows)) continue;
            foreach (var (values, _) in rows)
            {
                foreach (var cell in values)
                {
                    if (cell is String s && !_sstIndex.ContainsKey(s))
                    {
                        // 公式不加入 SST（以 = 开头）
                        if (s.Length > 0 && s[0] == '=') continue;
                        _sstIndex[s] = _sst.Count;
                        _sst.Add(s);
                    }
                }
            }
        }
    }

    private Byte[] BuildWorkbookStream()
    {
        // 收集所有不重复的行样式，分配字体/XF 索引
        var styleMap = new Dictionary<ExcelCellStyle, Int32>(); // style → xfIndex (1基，0=默认)
        var styleFonts = new List<ExcelCellStyle>();
        var nextXfIndex = 21; // 0-20 为内置默认
        var nextFontIndex = 6;  // 0-5 为默认字体

        foreach (var sheetName in _sheetNames)
        {
            if (!_sheetData.TryGetValue(sheetName, out var rows)) continue;
            foreach (var (_, style) in rows)
            {
                if (style != null && !styleMap.ContainsKey(style))
                {
                    styleMap[style] = nextXfIndex++;
                    styleFonts.Add(style);
                }
            }
        }

        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms, Encoding.Unicode, leaveOpen: true);

        // 提前收集格式映射（需要先于字体/XF 记录使用）
        var formatMap = new Dictionary<String, Int32>(); // formatString → formatIndex
        var nextFormatIdx = 165;
        foreach (var sheetName in _sheetNames)
        {
            if (!_columnFormats.TryGetValue(sheetName, out var colFmts)) continue;
            foreach (var kv in colFmts)
            {
                var fmt = kv.Value;
                if (!formatMap.ContainsKey(fmt))
                    formatMap[fmt] = nextFormatIdx++;
            }
        }

        // 为每个自定义格式创建 XF 索引
        var formatXfIndex = new Dictionary<Int32, Int32>(); // formatIndex → xfIndex
        foreach (var kv in formatMap)
        {
            formatXfIndex[kv.Value] = nextXfIndex;
            nextFontIndex++; // 每个格式需要自己的字体槽位
            nextXfIndex++;
        }

        // 1. Globals BOF
        WriteRecord(bw, RecBof, BuildBofData(0x0005));

        // 2. 字体记录（6默认 + 样式字体 + 每自定义格式一个字体槽位）
        for (var fi = 0; fi < 6; fi++)
            WriteRecord(bw, RecFont, BuildFontRecord(null));
        foreach (var s in styleFonts)
            WriteRecord(bw, RecFont, BuildFontRecord(s));
        // 为每个自定义格式写一个默认字体
        foreach (var _ in formatMap)
            WriteRecord(bw, RecFont, BuildFontRecord(null));

        // 3. 格式记录（默认日期格式 + 自定义数字格式）
        // 写入内置日期格式
        WriteRecord(bw, RecFormat, BuildFormatRecord(164, "yyyy/mm/dd"));
        // 写入自定义列格式
        foreach (var kv in formatMap)
            WriteRecord(bw, RecFormat, BuildFormatRecord(kv.Value, kv.Key));

        // 4. XF 记录：21 条内置 + N 条自定义样式
        WriteXfRecords(bw, styleMap, nextFontIndex, formatXfIndex);

        // 5. BoundSheet
        var boundSheetPositions = new List<Int64>();
        foreach (var sheetName in _sheetNames)
        {
            boundSheetPositions.Add(ms.Position + 4);
            WriteRecord(bw, RecBoundSheet, BuildBoundSheetData(sheetName, 0));
        }

        // 6. SST
        WriteRecord(bw, RecSst, BuildSstRecord());

        // 7. Globals EOF
        WriteRecord(bw, RecEof, []);

        // 8. 写入各工作表
        for (var si = 0; si < _sheetNames.Count; si++)
        {
            var sheetName = _sheetNames[si];
            var sheetBofOffset = (Int32)ms.Position;
            var savedPos = ms.Position;
            ms.Position = boundSheetPositions[si];
            bw.Write(sheetBofOffset);
            ms.Position = savedPos;
            WriteSheetStream(bw, sheetName, styleMap, formatMap, formatXfIndex);
        }

        bw.Flush();
        return ms.ToArray();
    }

    private void WriteSheetStream(BinaryWriter bw, String sheetName, Dictionary<ExcelCellStyle, Int32> styleMap,
        Dictionary<String, Int32> formatMap, Dictionary<Int32, Int32> formatXfIndex)
    {
        var rows = _sheetData.TryGetValue(sheetName, out var r) ? r : [];

        // Sheet BOF
        WriteRecord(bw, RecBof, BuildBofData(0x0010));

        // DIMENSIONS
        var rowCount = rows.Count;
        var colCount = rows.Count > 0 ? rows.Max(r2 => r2.Values.Count) : 0;
        WriteRecord(bw, RecDimensions, BuildDimensionsData(rowCount, colCount));

        // WINDOW2 — 冻结窗格
        if (_sheetFreezePanes.TryGetValue(sheetName, out var freeze))
            WriteRecord(bw, RecWindow2, BuildWindow2Data(freeze.Row, freeze.Col));
        else
            WriteRecord(bw, RecWindow2, BuildWindow2Data(0, 0));

        // COLINFO — 列宽
        if (_sheetColWidths.TryGetValue(sheetName, out var colWidths))
        {
            foreach (var kv in colWidths.OrderBy(kv => kv.Key))
            {
                WriteRecord(bw, RecColInfo, BuildColInfoData(kv.Key, kv.Key, kv.Value));
            }
        }

        // 获取当前工作表的列格式映射
        _columnFormats.TryGetValue(sheetName, out var sheetColFmts);

        // ROW + 单元格记录
        for (var ri = 0; ri < rows.Count; ri++)
        {
            var (values, style) = rows[ri];
            var colMax = values.Count;

            // 确定此行的 XF 索引基础值
            var baseXf = 15; // 默认
            if (style != null && styleMap.TryGetValue(style, out var sx))
                baseXf = sx;

            WriteRecord(bw, RecRow, BuildRowData(ri, 0, colMax));

            for (var ci = 0; ci < values.Count; ci++)
            {
                // 获取此列的 XF 索引（格式优先于行样式）
                var cellXf = baseXf;
                if (sheetColFmts != null && sheetColFmts.TryGetValue(ci, out var colFmt) &&
                    formatMap.TryGetValue(colFmt, out var fmtIdx) &&
                    formatXfIndex.TryGetValue(fmtIdx, out var fmtXf))
                    cellXf = fmtXf;

                var cell = values[ci];
                if (cell == null)
                {
                    WriteRecord(bw, RecBlank, BuildBlankData(ri, ci, cellXf));
                }
                else if (cell is String strVal)
                {
                    // 检测公式：以 = 开头的字符串视为公式
                    if (strVal.Length > 0 && strVal[0] == '=')
                    {
                        WriteRecord(bw, RecFormula, BuildFormulaData(ri, ci, cellXf, strVal));
                    }
                    else
                    {
                        var sstIdx = _sstIndex.TryGetValue(strVal, out var idx) ? idx : 0;
                        WriteRecord(bw, RecLabelSst, BuildLabelSstData(ri, ci, sstIdx, cellXf));
                    }
                }
                else if (cell is Boolean boolVal)
                {
                    WriteRecord(bw, RecBoolErr, BuildBoolErrData(ri, ci, boolVal ? (Byte)1 : (Byte)0, false, cellXf));
                }
                else if (cell is DateTime dtVal)
                {
                    var serial = DateToSerial(dtVal);
                    // 日期使用列格式 XF 或内置日期 XF(1)
                    var dateXf = cellXf != 15 ? cellXf : 1;
                    WriteRecord(bw, RecNumber, BuildNumberData(ri, ci, serial, xfIndex: dateXf));
                }
                else
                {
                    var dbl = ConvertToDouble(cell);
                    WriteRecord(bw, RecNumber, BuildNumberData(ri, ci, dbl, cellXf));
                }
            }
        }

        // Sheet EOF
        WriteRecord(bw, RecEof, []);
    }

    #endregion

    #region 记录构建辅助

    private static Byte[] BuildBofData(UInt16 bofType)
    {
        var buf = new Byte[16];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)0x0600);   // BIFF8 version
        writer.Write(bofType);           // type
        writer.Write((UInt16)0x0DBB);   // build identifier
        writer.Write((UInt16)0x07CC);   // build year (1996)
        writer.Write(0x00000041u);       // file history flags
        writer.Write(0x00000006u);       // runtime version
        return buf;
    }

    private static Byte[] BuildBoundSheetData(String name, Int32 bofOffset)
    {
        var nameBytes = Encoding.Unicode.GetBytes(name);
        var buf = new Byte[8 + nameBytes.Length];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt32)bofOffset);
        writer.Write((Byte)0x00); // grbit (visible + worksheet)
        writer.Write((Byte)0x00);
        writer.Write((Byte)name.Length); // cch
        writer.Write((Byte)0x01); // fHighByte = UTF-16LE
        Array.Copy(nameBytes, 0, buf, 8, nameBytes.Length);
        return buf;
    }

    private Byte[] BuildSstRecord()
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms, Encoding.Unicode, leaveOpen: true);

        // 总字符串引用数 + 唯一字符串数
        var totalRefs = _sst.Count; // 简化：引用数 = 唯一数
        bw.Write(totalRefs);
        bw.Write(_sst.Count);

        foreach (var s in _sst)
        {
            // XLUnicodeString：cch(2) + flags(1) + UTF-16LE 数据
            bw.Write((UInt16)s.Length);
            bw.Write((Byte)0x01); // fHighByte = 1（UTF-16LE）
            foreach (var ch in s)
            {
                bw.Write((UInt16)ch);
            }
        }

        bw.Flush();
        return ms.ToArray();
    }

    private static Byte[] BuildDimensionsData(Int32 rowCount, Int32 colCount)
    {
        var buf = new Byte[14];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write(0u); // first row
        writer.Write((UInt32)Math.Max(rowCount, 1)); // last row + 1
        writer.Write((UInt16)0); // first col
        writer.Write((UInt16)Math.Max(colCount, 1)); // last col + 1
        writer.Write((UInt16)0); // reserved
        return buf;
    }

    private static Byte[] BuildColInfoData(Int32 firstCol, Int32 lastCol, Int32 width)
    {
        // COLINFO 记录：colFirst(2) + colLast(2) + coldx(2) + ixfe(2) + grbit(2)
        var buf = new Byte[12];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)firstCol);
        writer.Write((UInt16)lastCol);
        writer.Write((UInt16)width);     // 1/256 字符宽度
        writer.Write((UInt16)0x000F);    // XF index (default=15)
        writer.Write((UInt16)0x0000);    // grbit (not hidden, default)
        return buf;
    }

    private static Byte[] BuildFormulaData(Int32 row, Int32 col, Int32 xfIndex, String formula)
    {
        // 将公式字符串编码为字节
        var formulaBytes = Encoding.UTF8.GetBytes(formula);
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);

        bw.Write((UInt16)row);
        bw.Write((UInt16)col);
        bw.Write((UInt16)xfIndex);
        // Result: 8 bytes (0 = string/empty result)
        bw.Write(0L);
        // Options: 0x0001 = recalc always
        bw.Write((UInt16)0x0001);
        // Not used (4 bytes)
        bw.Write(0u);
        // Formula expression length (2 bytes) + raw formula bytes
        bw.Write((UInt16)formulaBytes.Length);
        bw.Write(formulaBytes);

        bw.Flush();
        return ms.ToArray();
    }

    private static Byte[] BuildRowData(Int32 row, Int32 firstCol, Int32 lastCol)
    {
        var buf = new Byte[16];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)firstCol);
        writer.Write((UInt16)lastCol);
        writer.Write((UInt16)0x00FF); // row height = 255 twips (default)
        writer.Write((UInt16)0);      // unused
        writer.Write((UInt16)0);      // unused
        writer.Write((UInt16)0x0100); // default row attributes
        writer.Write((UInt16)0x0F);   // XF index 15 (default)
        return buf;
    }

    private static Byte[] BuildLabelSstData(Int32 row, Int32 col, Int32 sstIndex, Int32 xfIndex = 15)
    {
        var buf = new Byte[10];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)col);
        writer.Write((UInt16)xfIndex);
        writer.Write((UInt32)sstIndex);
        return buf;
    }

    private static Byte[] BuildNumberData(Int32 row, Int32 col, Double value, Int32 xfIndex = 15)
    {
        var buf = new Byte[14];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)col);
        writer.Write((UInt16)xfIndex);
        writer.Write(value);
        return buf;
    }

    private static Byte[] BuildBoolErrData(Int32 row, Int32 col, Byte value, Boolean isError, Int32 xfIndex = 15)
    {
        var buf = new Byte[8];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)col);
        writer.Write((UInt16)xfIndex);
        writer.Write(value);
        writer.Write(isError ? (Byte)1 : (Byte)0);
        return buf;
    }

    private static Byte[] BuildBlankData(Int32 row, Int32 col, Int32 xfIndex = 15)
    {
        var buf = new Byte[6];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)col);
        writer.Write((UInt16)xfIndex);
        return buf;
    }

    /// <summary>构建 WINDOW2 记录（含冻结窗格信息）</summary>
    /// <param name="freezeRow">冻结行数（0=不冻结）</param>
    /// <param name="freezeCol">冻结列数（0=不冻结）</param>
    private static Byte[] BuildWindow2Data(Int32 freezeRow, Int32 freezeCol)
    {
        var buf = new Byte[18];
        var writer = new SpanWriter(buf, 0, buf.Length);
        // grbit: bit3(0x08)=frozen, bit4(0x10)=no split panes
        var flags = freezeRow > 0 || freezeCol > 0 ? (UInt16)0x0018 : (UInt16)0x0000;
        writer.Write(flags);
        writer.Write((UInt16)freezeRow);  // top row visible in frozen pane
        writer.Write((UInt16)freezeCol);  // left column visible in frozen pane
        writer.Write((UInt16)0x0040);     // color index (default: system foreground)
        writer.Write((UInt16)0x0000);     // reserved
        writer.Write((UInt16)0x0000);     // frozen scroll row (0 = not split)
        writer.Write((UInt16)0x0000);     // frozen scroll column
        writer.Write((UInt16)0x0040);     // Use gridline color
        return buf;
    }

    private static Byte[] BuildFontRecord(ExcelCellStyle? style = null)
    {
        var name = style?.FontName.IsNullOrEmpty() != false ? "Arial" : style.FontName!;
        var nameBytes = Encoding.Unicode.GetBytes(name);
        var size = style?.FontSize > 0 ? (Int32)(style.FontSize * 20) : 200; // 200 = 10pt
        var bold = style?.Bold == true ? 0x02BC : 0x0190; // 700 vs 400
        var italic = style?.Italic == true ? (UInt16)0x0002 : (UInt16)0;
        var colorIdx = MapColorToIndex(style?.FontColor);
        var buf = new Byte[16 + nameBytes.Length];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)size);
        writer.Write((UInt16)italic);
        writer.Write((UInt16)colorIdx);
        writer.Write((UInt16)bold);
        writer.Write((UInt16)0);
        writer.Write((Byte)0);
        writer.Write((Byte)0);
        writer.Write((Byte)0);
        writer.Write((Byte)0);
        writer.Write((Byte)name.Length);
        writer.Write((Byte)0x01);
        Array.Copy(nameBytes, 0, buf, 16, nameBytes.Length);
        return buf;
    }

    // 简单映射：常用颜色 → BIFF8 颜色索引
    private static UInt16 MapColorToIndex(String? rgb)
    {
        if (rgb.IsNullOrEmpty()) return 0x7FFF;
        return rgb?.ToUpper() switch
        {
            "000000" => 0x7FFF, // black → auto
            "FF0000" => 10,      // red
            "00FF00" => 11,      // bright green
            "0000FF" => 12,      // blue
            "FFFF00" => 13,      // yellow
            "FF00FF" => 14,      // magenta
            "00FFFF" => 15,      // cyan
            "800000" => 16,      // dark red
            "008000" => 17,      // dark green
            "000080" => 18,      // dark blue
            "808000" => 19,      // dark yellow
            "800080" => 20,      // dark magenta
            "008080" => 21,      // dark cyan
            "C0C0C0" => 22,      // silver
            "808080" => 23,      // gray
            _ => 0x7FFF,
        };
    }

    private static Byte[] BuildFormatRecord(Int32 formatIndex, String formatString)
    {
        // BIFF8 FORMAT 记录: ixfe(2) + cch(2) + fHighByte(1) + rgch(n*2)
        var fmtBytes = Encoding.Unicode.GetBytes(formatString);
        var buf = new Byte[5 + fmtBytes.Length];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)formatIndex);       // format index
        writer.Write((UInt16)formatString.Length); // character count
        writer.Write((Byte)0x01);                 // fHighByte = UTF-16LE
        Array.Copy(fmtBytes, 0, buf, 5, fmtBytes.Length);
        return buf;
    }

    private static void WriteXfRecords(BinaryWriter bw, Dictionary<ExcelCellStyle, Int32> styleMap, Int32 startFontIndex, Dictionary<Int32, Int32> formatXfIndex)
    {
        // 21 内置 XF：索引 0-14 = 普通, 15 = 默认, 16-20 = 标题
        for (var i = 0; i < 21; i++)
        {
            var xfData = BuildBuiltinXfRecord(i);
            WriteRecord(bw, RecXf, xfData);
        }
        // 自定义样式 XF
        foreach (var kv in styleMap)
        {
            var style = kv.Key;
            var fontIdx = startFontIndex + styleMap.Keys.ToList().IndexOf(style);
            var bgColor = style.BackgroundColor.IsNullOrEmpty() ? 0x40u : 0u;
            var xfData = BuildStyledXfRecord((UInt16)fontIdx, bgColor, style.Border != ExcelCellBorderStyle.None);
            WriteRecord(bw, RecXf, xfData);
        }
        // 自定义格式 XF（每个格式一个简约 XF 记录，仅设置格式索引）
        var fmtFontIdx = (UInt16)(startFontIndex + styleMap.Count);
        foreach (var kv in formatXfIndex)
        {
            var xfData = BuildFormatXfRecord(fmtFontIdx, (UInt16)kv.Key);
            WriteRecord(bw, RecXf, xfData);
            fmtFontIdx++;
        }
    }

    private static Byte[] BuildBuiltinXfRecord(Int32 index)
    {
        var buf = new Byte[22];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)0);
        writer.Write(index == 1 ? (UInt16)164 : (UInt16)0);
        writer.Write(index < 16 ? (UInt16)0xFFF5 : (UInt16)0x0001);
        writer.Write((UInt16)0x20C0);
        writer.Write((UInt16)0);
        writer.Write((UInt16)0);
        writer.Write((UInt16)0);
        writer.Write(0u);
        writer.Write(0u);
        return buf;
    }

    private static Byte[] BuildStyledXfRecord(UInt16 fontIndex, UInt32 bgColor, Boolean hasBorder)
    {
        var buf = new Byte[22];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write(fontIndex);
        writer.Write((UInt16)0); // format index
        writer.Write((UInt16)0xFFF5); // cell style
        writer.Write((UInt16)0x20C0); // default alignment (center-vert, no-wrap)
        writer.Write((UInt16)0);
        writer.Write((UInt16)0);
        writer.Write((UInt16)0);
        writer.Write(hasBorder ? 0x0000000Fu : 0u); // thin borders (4 nibbles)
        writer.Write(bgColor); // fill pattern + bg color
        return buf;
    }

    private static Byte[] BuildFormatXfRecord(UInt16 fontIndex, UInt16 formatIndex)
    {
        var buf = new Byte[22];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write(fontIndex);
        writer.Write(formatIndex);     // 指向 FORMAT 记录的索引
        writer.Write((UInt16)0xFFF5);  // cell style
        writer.Write((UInt16)0x20C0);  // default alignment
        writer.Write((UInt16)0);
        writer.Write((UInt16)0);
        writer.Write((UInt16)0);
        writer.Write(0u);              // no borders
        writer.Write(0x40u);           // no fill
        return buf;
    }

    private static void WriteRecord(BinaryWriter bw, UInt16 recType, Byte[] data)
    {
        if (data.Length <= MaxRecordDataSize)
        {
            bw.Write(recType);
            bw.Write((UInt16)data.Length);
            bw.Write(data);
            return;
        }

        // 超长数据需拆分 CONTINUE 记录
        var offset = 0;
        var first = true;
        while (offset < data.Length)
        {
            var chunk = Math.Min(MaxRecordDataSize, data.Length - offset);
            bw.Write(first ? recType : RecContinue);
            bw.Write((UInt16)chunk);
            bw.Write(data, offset, chunk);
            offset += chunk;
            first = false;
        }
    }

    #endregion

    #region 辅助

    private List<(List<Object?> Values, ExcelCellStyle? Style)> GetCurrentSheet()
    {
        if (!_sheetData.TryGetValue(_currentSheet, out var rows))
        {
            rows = [];
            _sheetData[_currentSheet] = rows;
            _sheetNames.Add(_currentSheet);
        }
        return rows;
    }

    private static Double DateToSerial(DateTime dt)
    {
        // Excel 日期序列号：从 1900-01-00 开始（含 1900 年闰年 bug：+1）
        var days = (dt.Date - DateEpoch).TotalDays + DateEpochOffset;
        var time = dt.TimeOfDay.TotalDays;
        return days + time;
    }

    private static Double ConvertToDouble(Object? value)
    {
        return value switch
        {
            Double d => d,
            Single f => (Double)f,
            Decimal dec => (Double)dec,
            Int32 i => i,
            Int64 l => l,
            Int16 sh => sh,
            Byte b => b,
            SByte sb2 => sb2,
            UInt16 us => us,
            UInt32 ui => ui,
            UInt64 ul => ul,
            _ => Convert.ToDouble(value)
        };
    }

    private static PropertyInfo[] GetMappableProperties<T>()
    {
        return typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
            .ToArray();
    }

    private static String GetPropertyDisplayName(PropertyInfo p)
    {
        var dn = p.GetCustomAttributes<DisplayNameAttribute>(false).FirstOrDefault();
        return dn?.DisplayName ?? p.Name;
    }

    #endregion
}
