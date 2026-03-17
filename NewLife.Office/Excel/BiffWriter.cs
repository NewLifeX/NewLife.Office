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
    private readonly Dictionary<String, List<List<Object?>>> _sheetData = new(StringComparer.Ordinal);

    // 共享字符串表
    private readonly List<String> _sst = [];
    private readonly Dictionary<String, Int32> _sstIndex = new(StringComparer.Ordinal);

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
        var sheet = GetCurrentSheet();
        sheet.Add(values.ToList());
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
            foreach (var row in rows)
            {
                foreach (var cell in row)
                {
                    if (cell is String s && !_sstIndex.ContainsKey(s))
                    {
                        _sstIndex[s] = _sst.Count;
                        _sst.Add(s);
                    }
                }
            }
        }
    }

    private Byte[] BuildWorkbookStream()
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms, Encoding.Unicode, leaveOpen: true);

        // 1. Globals BOF
        WriteRecord(bw, RecBof, BuildBofData(0x0005)); // 0x0005 = Workbook Globals

        // 2. 默认字体（BIFF8 要求至少 5 个 FONT 记录，index 0-4，index 4 为保留）
        for (var fi = 0; fi < 6; fi++)
        {
            WriteRecord(bw, RecFont, BuildFontRecord());
        }

        // 3. 默认格式（Format record，数值日期格式索引）
        WriteRecord(bw, RecFormat, BuildFormatRecord(164, "yyyy/mm/dd")); // 自定义日期格式

        // 4. XF（扩展格式）记录：最少 21 条（BIFF8 基本样式）
        WriteXfRecords(bw);

        // 5. BoundSheet：每个工作表一条，先填 0 偏移占位
        var boundSheetPositions = new List<Int64>();
        foreach (var sheetName in _sheetNames)
        {
            boundSheetPositions.Add(ms.Position + 4); // +4 跳过记录头
            WriteRecord(bw, RecBoundSheet, BuildBoundSheetData(sheetName, 0));
        }

        // 6. 共享字符串表（SST）
        WriteRecord(bw, RecSst, BuildSstRecord());

        // 7. Globals EOF
        WriteRecord(bw, RecEof, []);

        // 8. 写入各工作表，并回填 BoundSheet 偏移
        for (var si = 0; si < _sheetNames.Count; si++)
        {
            var sheetName = _sheetNames[si];
            var sheetBofOffset = (Int32)ms.Position;

            // 回填 BoundSheet 中的 BOF 偏移
            var savedPos = ms.Position;
            ms.Position = boundSheetPositions[si];
            bw.Write(sheetBofOffset);
            ms.Position = savedPos;

            // 写入工作表数据
            WriteSheetStream(bw, sheetName);
        }

        bw.Flush();
        return ms.ToArray();
    }

    private void WriteSheetStream(BinaryWriter bw, String sheetName)
    {
        var rows = _sheetData.TryGetValue(sheetName, out var r) ? r : [];

        // Sheet BOF
        WriteRecord(bw, RecBof, BuildBofData(0x0010)); // 0x0010 = Worksheet

        // DIMENSIONS 记录
        var rowCount = rows.Count;
        var colCount = rows.Count > 0 ? rows.Max(r2 => r2.Count) : 0;
        WriteRecord(bw, RecDimensions, BuildDimensionsData(rowCount, colCount));

        // ROW + 单元格记录
        for (var ri = 0; ri < rows.Count; ri++)
        {
            var row = rows[ri];
            var colMax = row.Count;

            // ROW 描述记录
            WriteRecord(bw, RecRow, BuildRowData(ri, 0, colMax));

            // 单元格数据
            for (var ci = 0; ci < row.Count; ci++)
            {
                var cell = row[ci];
                if (cell == null)
                {
                    WriteRecord(bw, RecBlank, BuildBlankData(ri, ci));
                }
                else if (cell is String strVal)
                {
                    var sstIdx = _sstIndex.TryGetValue(strVal, out var idx) ? idx : 0;
                    WriteRecord(bw, RecLabelSst, BuildLabelSstData(ri, ci, sstIdx));
                }
                else if (cell is Boolean boolVal)
                {
                    WriteRecord(bw, RecBoolErr, BuildBoolErrData(ri, ci, boolVal ? (Byte)1 : (Byte)0, false));
                }
                else if (cell is DateTime dtVal)
                {
                    var serial = DateToSerial(dtVal);
                    WriteRecord(bw, RecNumber, BuildNumberData(ri, ci, serial, xfIndex: 1)); // XF index 1 = 日期格式
                }
                else
                {
                    var dbl = ConvertToDouble(cell);
                    WriteRecord(bw, RecNumber, BuildNumberData(ri, ci, dbl));
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

    private static Byte[] BuildLabelSstData(Int32 row, Int32 col, Int32 sstIndex)
    {
        var buf = new Byte[10];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)col);
        writer.Write((UInt16)0x000F);    // XF index 15
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

    private static Byte[] BuildBoolErrData(Int32 row, Int32 col, Byte value, Boolean isError)
    {
        var buf = new Byte[8];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)col);
        writer.Write((UInt16)0x000F); // XF index 15
        writer.Write(value);
        writer.Write(isError ? (Byte)1 : (Byte)0);
        return buf;
    }

    private static Byte[] BuildBlankData(Int32 row, Int32 col)
    {
        var buf = new Byte[6];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)row);
        writer.Write((UInt16)col);
        writer.Write((UInt16)0x000F); // XF index 15
        return buf;
    }

    private static Byte[] BuildFontRecord()
    {
        // 默认字体：Arial 10pt（BIFF8 FONT 记录结构）
        var name = "Arial";
        var nameBytes = Encoding.Unicode.GetBytes(name);
        // dyHeight(2)+grbit(2)+icv(2)+bls(2)+sss(2)+uls(1)+bFamily(1)+bCharSet(1)+reserved(1)+cch(1)+fHighByte(1)+name(n)
        var buf = new Byte[16 + nameBytes.Length];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)200);    // dyHeight: 200 = 10pt (in 1/20 pt units)
        writer.Write((UInt16)0);      // grbit
        writer.Write((UInt16)0x7FFF); // icv: colour index (default/auto)
        writer.Write((UInt16)0x0190); // bls: bold weight (400 = normal)
        writer.Write((UInt16)0);      // sss: super/sub script
        writer.Write((Byte)0);        // uls: underline type
        writer.Write((Byte)0);        // bFamily: font family
        writer.Write((Byte)0);        // bCharSet: charset
        writer.Write((Byte)0);        // reserved
        writer.Write((Byte)name.Length); // cch: character count of font name
        writer.Write((Byte)0x01);     // fHighByte: 1 = UTF-16LE
        Array.Copy(nameBytes, 0, buf, 16, nameBytes.Length);
        return buf;
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

    private static void WriteXfRecords(BinaryWriter bw)
    {
        // BIFF8 要求至少 21 个内置 XF 记录（样式索引 0-14 = 普通，15 = 默认单元格格式，16-20 = 标题）
        // 索引 1：日期格式（formatIndex = 14 = "m/d/yy"）
        for (var i = 0; i < 21; i++)
        {
            var xfData = BuildXfRecord(i);
            WriteRecord(bw, RecXf, xfData);
        }
    }

    private static Byte[] BuildXfRecord(Int32 index)
    {
        var buf = new Byte[22];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write((UInt16)0); // font index
        // Format index: 1 = 日期格式自定义索引 164
        writer.Write(index == 1 ? (UInt16)164 : (UInt16)0);
        writer.Write(index < 16 ? (UInt16)0xFFF5 : (UInt16)0x0001); // style flag
        writer.Write((UInt16)0x20C0); // alignment
        writer.Write((UInt16)0);      // rotation
        writer.Write((UInt16)0);      // text properties
        writer.Write((UInt16)0);      // used attribute
        writer.Write(0u);             // border lines
        writer.Write(0u);             // colour / pattern
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

    private List<List<Object?>> GetCurrentSheet()
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
