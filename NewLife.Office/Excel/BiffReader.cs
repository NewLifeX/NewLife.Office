using System.Text;

namespace NewLife.Office;

/// <summary>xls（BIFF8）格式读取器</summary>
/// <remarks>
/// 通过 OLE2/CFB 容器解析 Microsoft Excel 97-2003 二进制格式（BIFF8），
/// 提取工作表数据、共享字符串表（SST）、单元格数值/布尔等内容。
/// <para>读取示例：</para>
/// <code>
/// using var reader = new BiffReader("data.xls");
/// foreach (var name in reader.SheetNames)
/// {
///     var rows = reader.ReadSheet(name).ToList();
/// }
/// </code>
/// </remarks>
public sealed class BiffReader : IDisposable
{

    #region 属性
    /// <summary>工作表名称列表（按顺序）</summary>
    public IReadOnlyList<String> SheetNames => _sheetNames;
    #endregion

    #region 私有字段
    private readonly Byte[] _workbook;
    private String[] _sst = [];
    private List<String> _sheetNames = [];
    private List<Int32> _sheetBofOffsets = [];
    private Boolean _disposed;
    #endregion

    #region 构造与打开
    /// <summary>从 xls 文件路径打开</summary>
    /// <param name="path">xls 文件路径</param>
    public BiffReader(String path)
    {
        using var doc = CfbDocument.Open(path);
        _workbook = GetWorkbookStream(doc);
        Parse();
    }

    /// <summary>从流打开（需包含 xls 的完整 OLE2 容器内容）</summary>
    /// <param name="stream">可读流</param>
    public BiffReader(Stream stream)
    {
        using var doc = CfbDocument.Open(stream, leaveOpen: true);
        _workbook = GetWorkbookStream(doc);
        Parse();
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

    private static Byte[] GetWorkbookStream(CfbDocument doc)
    {
        // Excel 97+ 使用 "Workbook"；更早版本使用 "Book"
        var data = doc.GetStreamData("Workbook") ?? doc.GetStreamData("Book");
        if (data == null || data.Length == 0)
            throw new InvalidDataException("找不到 Workbook 流，文件可能不是有效的 xls 格式。");
        return data;
    }
    #endregion

    #region 解析
    private void Parse()
    {
        // 第一遍：收集所有记录；合并 CONTINUE 扩展数据
        var records = ReadAllRecords(_workbook);

        // 找到 Globals BOF
        var globalsStart = FindGlobalsBof(records);
        if (globalsStart < 0)
            throw new InvalidDataException("未找到 BIFF8 Globals BOF 记录，文件可能损坏。");

        // 解析 Globals 段（从 globalsStart 到第一个 EOF）
        var pos = globalsStart;
        while (pos < records.Count)
        {
            var (type, data) = records[pos];
            pos++;
            if (type == RecEof) break;
            if (type == RecSst)
                _sst = ParseSst(data);
            else if (type == RecBoundSheet)
                ParseBoundSheet(data);
        }
    }

    /// <summary>读取所有 BIFF 记录（自动合并 CONTINUE 内容）</summary>
    /// <param name="buf">工作簿字节流</param>
    /// <returns>记录列表（type, data）</returns>
    private static List<(UInt16 type, Byte[] data)> ReadAllRecords(Byte[] buf)
    {
        var result = new List<(UInt16, Byte[])>();
        var pos = 0;
        while (pos + 4 <= buf.Length)
        {
            var type = ReadUInt16(buf, pos);
            var len = (Int32)ReadUInt16(buf, pos + 2);
            pos += 4;
            if (pos + len > buf.Length) break;

            var data = Slice(buf, pos, len);
            pos += len;

            // 将紧随的 CONTINUE 记录合并（SST 等超长记录会跨多个 CONTINUE）
            while (pos + 4 <= buf.Length)
            {
                var nextType = ReadUInt16(buf, pos);
                if (nextType != RecContinue) break;
                var nextLen = (Int32)ReadUInt16(buf, pos + 2);
                pos += 4;
                if (pos + nextLen > buf.Length) break;
                var continuation = Slice(buf, pos, nextLen);
                pos += nextLen;
                var merged = new Byte[data.Length + continuation.Length];
                Array.Copy(data, 0, merged, 0, data.Length);
                Array.Copy(continuation, 0, merged, data.Length, continuation.Length);
                data = merged;
            }

            result.Add(((UInt16)type, data));
        }
        return result;
    }

    /// <summary>定位第一个 Globals 类型 BOF（recType=BOF, recVerType=0x0005）</summary>
    /// <param name="records">记录列表</param>
    /// <returns>索引，-1 表示未找到</returns>
    private static Int32 FindGlobalsBof(List<(UInt16, Byte[])> records)
    {
        for (var i = 0; i < records.Count; i++)
        {
            var (t, d) = records[i];
            if (t != RecBof) continue;
            if (d.Length < 4) continue;
            var version = ReadUInt16(d, 0);
            var bofType = ReadUInt16(d, 2);
            if (version == 0x0600 && bofType == 0x0005) return i + 1; // 指向 BOF 后的第一条
        }
        return -1;
    }

    /// <summary>解析共享字符串表（SST）</summary>
    /// <param name="data">SST 记录数据（含 CONTINUE 已合并）</param>
    /// <returns>字符串数组</returns>
    private static String[] ParseSst(Byte[] data)
    {
        if (data.Length < 8) return [];

        // 4字节总引用数，4字节唯一字符串数
        var uniqueCount = (Int32)ReadUInt32(data, 4);
        var strings = new String[uniqueCount];
        var pos = 8;

        for (var i = 0; i < uniqueCount; i++)
        {
            if (pos + 2 > data.Length) break;
            strings[i] = ReadXluString(data, ref pos);
        }

        return strings;
    }

    /// <summary>解析 XLUnicodeRichExtendedString</summary>
    /// <param name="data">字节数组</param>
    /// <param name="pos">当前偏移（解析后自动前进）</param>
    /// <returns>字符串内容</returns>
    private static String ReadXluString(Byte[] data, ref Int32 pos)
    {
        if (pos + 3 > data.Length) return String.Empty;

        var cch = (Int32)ReadUInt16(data, pos);     // 字符数
        var flags = data[pos + 2];                   // 标志字节
        pos += 3;

        var fHighByte = (flags & 0x01) != 0;         // 1=UTF-16LE，0=Latin-1
        var fRichString = (flags & 0x08) != 0;       // 含富文本
        var fExtString = (flags & 0x04) != 0;        // 含扩展数据

        var cRun = 0;
        if (fRichString)
        {
            if (pos + 2 > data.Length) return String.Empty;
            cRun = (Int32)ReadUInt16(data, pos);
            pos += 2;
        }

        var cbExtRst = 0;
        if (fExtString)
        {
            if (pos + 4 > data.Length) return String.Empty;
            cbExtRst = (Int32)ReadUInt32(data, pos);
            pos += 4;
        }

        // 读取字符数据
        String result;
        if (fHighByte)
        {
            // UTF-16LE，每字符 2 字节
            var byteCount = cch * 2;
            if (pos + byteCount > data.Length) byteCount = data.Length - pos;
            result = Encoding.Unicode.GetString(data, pos, byteCount);
            pos += cch * 2;
        }
        else
        {
            // Latin-1/扩展 ASCII，每字符 1 字节
            if (pos + cch > data.Length) cch = data.Length - pos;
            result = DecodeLatin1(data, pos, cch);
            pos += cch;
        }

        // 跳过富文本索引（每条 4 字节）
        pos += cRun * 4;
        // 跳过扩展数据
        if (cbExtRst > 0)
            pos += cbExtRst;

        return result;
    }

    /// <summary>解析 BoundSheet 记录，提取工作表名称和 BOF 偏移</summary>
    /// <param name="data">BoundSheet 记录数据</param>
    private void ParseBoundSheet(Byte[] data)
    {
        if (data.Length < 8) return;

        var bofOffset = (Int32)ReadUInt32(data, 0);
        // data[4..5] = grbit（可见性+类型）
        var nameLen = data[6];   // cch: 字符数
        var flags = data[7];     // fHighByte: 0=Latin-1, 1=Unicode
        var isUnicode = (flags & 0x01) != 0;

        String name;
        if (isUnicode)
            name = Encoding.Unicode.GetString(data, 8, nameLen * 2);
        else
            name = DecodeLatin1(data, 8, nameLen);

        _sheetNames.Add(name);
        _sheetBofOffsets.Add(bofOffset);
    }
    #endregion

    #region 读取方法
    /// <summary>逐行读取工作表数据</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>行数组序列，每行为对象数组（String/Double/Boolean/null）</returns>
    public IEnumerable<Object?[]> ReadSheet(String? sheet = null)
    {
        var idx = ResolveSheetIndex(sheet);
        if (idx < 0) yield break;

        // 定位到工作表 BOF（从文件流偏移）
        var bofFileOffset = _sheetBofOffsets[idx];
        var sheetRecords = ReadSheetRecords(_workbook, bofFileOffset);

        // 收集所有单元格
        var cells = new SortedDictionary<Int32, SortedDictionary<Int32, Object?>>();
        var maxCol = 0;

        foreach (var (type, data) in sheetRecords)
        {
            if (type == RecLabelSst)
                ParseLabelSst(data, cells, ref maxCol);
            else if (type == RecNumber)
                ParseNumber(data, cells, ref maxCol);
            else if (type == RecRk)
                ParseRk(data, cells, ref maxCol);
            else if (type == RecMulRk)
                ParseMulRk(data, cells, ref maxCol);
            else if (type == RecBoolErr)
                ParseBoolErr(data, cells, ref maxCol);
            else if (type == RecLabel)
                ParseLabel(data, cells, ref maxCol);
            // RecBlank / RecMulBlank：空单元格，跳过
        }

        // 按行返回数据
        if (cells.Count == 0) yield break;
        var lastRow = cells.Keys.Max();
        for (var r = 0; r <= lastRow; r++)
        {
            var row = new Object?[maxCol + 1];
            if (cells.TryGetValue(r, out var cols))
            {
                foreach (var kv in cols)
                {
                    if (kv.Key <= maxCol)
                        row[kv.Key] = kv.Value;
                }
            }
            yield return row;
        }
    }

    /// <summary>读取工作表数据并映射到对象集合</summary>
    /// <typeparam name="T">目标类型</typeparam>
    /// <param name="sheet">工作表名，null 取第一个</param>
    /// <returns>对象序列</returns>
    public IEnumerable<T> ReadObjects<T>(String? sheet = null) where T : class, new()
    {
        var props = typeof(T).GetProperties();
        var rows = ReadSheet(sheet).ToList();
        if (rows.Count < 2) yield break;

        var headers = rows[0];
        for (var ri = 1; ri < rows.Count; ri++)
        {
            var row = rows[ri];
            var obj = new T();
            for (var ci = 0; ci < Math.Min(headers.Length, row.Length); ci++)
            {
                var hdr = (headers[ci]?.ToString() ?? String.Empty).Trim();
                var prop = props.FirstOrDefault(p =>
                    p.Name.Equals(hdr, StringComparison.OrdinalIgnoreCase) ||
                    p.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), false)
                     .OfType<System.ComponentModel.DisplayNameAttribute>().Any(a => a.DisplayName == hdr));
                if (prop == null) continue;
                try
                {
                    var value = row[ci];
                    if (value == null) continue;
                    if (prop.PropertyType == typeof(String))
                        prop.SetValue(obj, Convert.ToString(value));
                    else
                        prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType));
                }
                catch { /* 跳过转换失败 */ }
            }
            yield return obj;
        }
    }
    #endregion

    #region 单元格解析辅助
    private void ParseLabelSst(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+sstIndex(4) = 共10字节
        if (data.Length < 10) return;
        var row = (Int32)ReadUInt16(data, 0);
        var col = (Int32)ReadUInt16(data, 2);
        var sstIdx = (Int32)ReadUInt32(data, 6);
        var value = sstIdx >= 0 && sstIdx < _sst.Length ? (Object?)_sst[sstIdx] : null;
        SetCell(cells, row, col, value, ref maxCol);
    }

    private static void ParseNumber(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+double(8) = 14字节
        if (data.Length < 14) return;
        var row = (Int32)ReadUInt16(data, 0);
        var col = (Int32)ReadUInt16(data, 2);
        var value = BitConverterToDouble(data, 6);
        SetCell(cells, row, col, (Object?)value, ref maxCol);
    }

    private static void ParseRk(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+rk(4) = 10字节
        if (data.Length < 10) return;
        var row = (Int32)ReadUInt16(data, 0);
        var col = (Int32)ReadUInt16(data, 2);
        var rk = (Int32)ReadUInt32(data, 6);
        var value = DecodeRk(rk);
        SetCell(cells, row, col, (Object?)value, ref maxCol);
    }

    private static void ParseMulRk(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+firstCol(2)+[xf(2)+rk(4)]*n+lastCol(2)
        if (data.Length < 6) return;
        var row = (Int32)ReadUInt16(data, 0);
        var firstCol = (Int32)ReadUInt16(data, 2);
        var lastCol = (Int32)ReadUInt16(data, data.Length - 2);
        var count = lastCol - firstCol + 1;
        for (var i = 0; i < count; i++)
        {
            var offset = 4 + i * 6;
            if (offset + 6 > data.Length - 2) break;
            var rk = (Int32)ReadUInt32(data, offset + 2);
            var value = DecodeRk(rk);
            SetCell(cells, row, firstCol + i, (Object?)value, ref maxCol);
        }
    }

    private static void ParseBoolErr(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+boolOrErr(1)+isError(1) = 8字节
        if (data.Length < 8) return;
        var row = (Int32)ReadUInt16(data, 0);
        var col = (Int32)ReadUInt16(data, 2);
        var isError = data[7] != 0;
        if (!isError)
        {
            var boolVal = data[6] != 0;
            SetCell(cells, row, col, (Object?)boolVal, ref maxCol);
        }
        // 错误值暂时跳过
    }

    private static void ParseLabel(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+cch(2)+fHighByte(1)+chars
        if (data.Length < 9) return;
        var row = (Int32)ReadUInt16(data, 0);
        var col = (Int32)ReadUInt16(data, 2);
        var cch = (Int32)ReadUInt16(data, 6);
        var fHighByte = data[8];
        String value;
        if (fHighByte != 0)
            value = Encoding.Unicode.GetString(data, 9, Math.Min(cch * 2, data.Length - 9));
        else
            value = DecodeLatin1(data, 9, Math.Min(cch, data.Length - 9));
        SetCell(cells, row, col, (Object?)value, ref maxCol);
    }

    private static void SetCell(SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells,
        Int32 row, Int32 col, Object? value, ref Int32 maxCol)
    {
        if (!cells.TryGetValue(row, out var rowDict))
        {
            rowDict = new SortedDictionary<Int32, Object?>();
            cells[row] = rowDict;
        }
        rowDict[col] = value;
        if (col > maxCol) maxCol = col;
    }

    /// <summary>解码 RK 压缩数值</summary>
    /// <param name="rk">4字节 RK 值</param>
    /// <returns>解码后的浮点数</returns>
    private static Double DecodeRk(Int32 rk)
    {
        Double value;
        if ((rk & 0x02) != 0)
        {
            // 整数：RK 值右移 2 位取整数
            value = rk >> 2;
        }
        else
        {
            // 浮点：将 RK 高 30 位放入 double 的高 30 位
            var d64 = ((Int64)(rk & unchecked((Int32)0xFFFFFFFC))) << 32;
            var bytes = BitConverter.GetBytes(d64);
            value = BitConverter.ToDouble(bytes, 0);
        }

        // bit 1 = 0 表示乘以 100，否则不处理（注意：BIFF8 中 bit1=1 表示 ÷100）
        if ((rk & 0x01) != 0)
            value /= 100.0;

        return value;
    }
    #endregion

    #region 工作表记录读取
    /// <summary>从指定文件流偏移读取工作表段的所有记录（到 EOF）</summary>
    /// <param name="buf">工作簿字节流</param>
    /// <param name="bofOffset">工作表 BOF 在字节流中的偏移</param>
    /// <returns>记录列表</returns>
    private static List<(UInt16, Byte[])> ReadSheetRecords(Byte[] buf, Int32 bofOffset)
    {
        var result = new List<(UInt16, Byte[])>();
        var pos = bofOffset;
        var depth = 0;

        while (pos + 4 <= buf.Length)
        {
            var type = ReadUInt16(buf, pos);
            var len = (Int32)ReadUInt16(buf, pos + 2);
            pos += 4;
            if (pos + len > buf.Length) break;

            var data = Slice(buf, pos, len);
            pos += len;

            if (type == RecBof)
            {
                depth++;
            }
            else if (type == RecEof)
            {
                // depth==1 说明这是工作表级别的 EOF，读取结束
                if (depth == 1) break;
                depth--;
            }

            result.Add((type, data));
        }
        return result;
    }

    private Int32 ResolveSheetIndex(String? name)
    {
        if (_sheetNames.Count == 0) return -1;
        if (String.IsNullOrEmpty(name)) return 0;
        for (var i = 0; i < _sheetNames.Count; i++)
            if (_sheetNames[i].Equals(name, StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }
    #endregion

    #region 字节工具
    private static UInt16 ReadUInt16(Byte[] buf, Int32 pos) =>
        (UInt16)(buf[pos] | (buf[pos + 1] << 8));

    private static UInt32 ReadUInt32(Byte[] buf, Int32 pos) =>
        (UInt32)(buf[pos] | (buf[pos + 1] << 8) | (buf[pos + 2] << 16) | (buf[pos + 3] << 24));

    private static Double BitConverterToDouble(Byte[] buf, Int32 pos)
    {
        var tmp = new Byte[8];
        Array.Copy(buf, pos, tmp, 0, 8);
        return BitConverter.ToDouble(tmp, 0);
    }

    private static Byte[] Slice(Byte[] buf, Int32 offset, Int32 length)
    {
        var result = new Byte[length];
        Array.Copy(buf, offset, result, 0, length);
        return result;
    }

    /// <summary>ISO-8859-1 直接映射：每个字节直接转为 Unicode 同码点的字符</summary>
    /// <param name="data">字节数组</param>
    /// <param name="pos">起始偏移</param>
    /// <param name="count">字节数</param>
    /// <returns>解码后的字符串</returns>
    private static String DecodeLatin1(Byte[] data, Int32 pos, Int32 count)
    {
        var chars = new Char[count];
        for (var i = 0; i < count; i++)
            chars[i] = (Char)data[pos + i];
        return new String(chars);
    }
    #endregion

    #region 记录类型常量
    // BIFF8 记录类型常量
    private const UInt16 RecBof = 0x0809;
    private const UInt16 RecEof = 0x000A;
    private const UInt16 RecContinue = 0x003C;
    private const UInt16 RecSst = 0x00FC;
    private const UInt16 RecBoundSheet = 0x0085;
    private const UInt16 RecLabelSst = 0x00FD;
    private const UInt16 RecNumber = 0x0203;
    private const UInt16 RecRk = 0x027E;
    private const UInt16 RecMulRk = 0x00BE;
    private const UInt16 RecBoolErr = 0x0205;
    private const UInt16 RecLabel = 0x0204;
    private const UInt16 RecBlank = 0x0201;
    private const UInt16 RecMulBlank = 0x00BF;
    #endregion
}
