using System.Globalization;
using System.Text;
using NewLife.Buffers;
using NewLife.Office.Ole2;

namespace NewLife.Office.Excel;

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
public sealed class BiffReader : IDisposable, ITextExtractable, IMarkdownExtractable
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
    private readonly List<(Int32 Row, Int32 Col, String Url)> _hyperlinks = [];
    private readonly Dictionary<Int32, Double> _columnWidths = [];
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
        // 创建单一 SpanReader，整个解析过程共用，不在循环内重复分配
        var reader = new SpanReader(_workbook, 0, _workbook.Length);

        // 第一遍：收集所有记录；合并 CONTINUE 扩展数据
        var records = ReadAllRecords(ref reader);

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
    /// <param name="reader">已定位到工作簿头部的读取器</param>
    /// <returns>记录列表（type, data）</returns>
    private static List<(UInt16 type, Byte[] data)> ReadAllRecords(ref SpanReader reader)
    {
        var result = new List<(UInt16, Byte[])>();
        while (reader.Available >= 4)
        {
            var type = reader.ReadUInt16();
            var len = (Int32)reader.ReadUInt16();
            if (reader.Available < len) break;

            var data = reader.ReadBytes(len).ToArray();

            // 将紧随的 CONTINUE 记录合并（SST 等超长记录会跨多个 CONTINUE）
            while (reader.Available >= 4)
            {
                var savedPos = reader.Position;
                var nextType = reader.ReadUInt16();
                if (nextType != RecContinue) { reader.Position = savedPos; break; }
                var nextLen = (Int32)reader.ReadUInt16();
                if (reader.Available < nextLen) { reader.Position = savedPos; break; }
                var continuation = reader.ReadBytes(nextLen).ToArray();
                var merged = new Byte[data.Length + continuation.Length];
                Array.Copy(data, 0, merged, 0, data.Length);
                Array.Copy(continuation, 0, merged, data.Length, continuation.Length);
                data = merged;
            }

            result.Add((type, data));
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
            var reader = new SpanReader(d, 0, 4);
            var version = reader.ReadUInt16();
            var bofType = reader.ReadUInt16();
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
        var reader = new SpanReader(data, 4, 4);
        var uniqueCount = reader.ReadInt32();
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
            cRun = ReadUInt16(data, pos);
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

        var reader = new SpanReader(data, 0, data.Length);
        var bofOffset = reader.ReadInt32();
        reader.Advance(2); // grbit（可见性+类型）
        var nameLen = reader.ReadByte();   // cch: 字符数
        var flags = reader.ReadByte();     // fHighByte: 0=Latin-1, 1=Unicode
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
        var sheetReader = new SpanReader(_workbook, 0, _workbook.Length);
        var sheetRecords = ReadSheetRecords(ref sheetReader, bofFileOffset);

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
            else if (type == RecFormula)
                ParseFormula(data, cells, ref maxCol);
            else if (type == RecHyperlink)
                ParseHyperlink(data);
            else if (type == RecColInfo)
                ParseColInfo(data);
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

    #region 增强读取（合并/页面设置/保护/验证）
    /// <summary>读取指定工作表的合并单元格区域</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>合并区域列表（R1, C1, R2, C2，均 0 基）</returns>
    public List<(Int32 R1, Int32 C1, Int32 R2, Int32 C2)> GetMerges(String? sheet = null)
    {
        var result = new List<(Int32, Int32, Int32, Int32)>();
        foreach (var (type, data) in ReadSheetRecords(sheet))
        {
            if (type != RecMergedCells || data.Length < 2) continue;

            var count = ReadUInt16(data, 0);
            for (var i = 0; i < count; i++)
            {
                var pos = 2 + i * 8;
                if (pos + 8 > data.Length) break;
                // 字节序：r1(2) r2(2) c1(2) c2(2)
                var r1 = ReadUInt16(data, pos);
                var r2 = ReadUInt16(data, pos + 2);
                var c1 = ReadUInt16(data, pos + 4);
                var c2 = ReadUInt16(data, pos + 6);
                result.Add((r1, c1, r2, c2));
            }
        }
        return result;
    }

    /// <summary>读取指定工作表的页面设置（SETUP + HEADER + FOOTER）</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>(横向, 纸张大小, 页眉边距英寸, 页脚边距英寸, 页眉, 页脚)</returns>
    public (Boolean Landscape, Int32 PaperSize, Double HeaderMargin, Double FooterMargin, String Header, String Footer) GetPageSetup(String? sheet = null)
    {
        var landscape = false;
        var paperSize = 0;
        var headerMargin = 0.0;
        var footerMargin = 0.0;
        var header = String.Empty;
        var footer = String.Empty;

        foreach (var (type, data) in ReadSheetRecords(sheet))
        {
            if (type == RecSetup && data.Length >= 22)
            {
                paperSize = ReadUInt16(data, 0);
                landscape = (ReadUInt16(data, 10) & 0x01) != 0;
                headerMargin = ReadUInt16(data, 16) / 100.0;
                footerMargin = ReadUInt16(data, 20) / 100.0;
            }
            else if (type == RecHeader && data.Length >= 2)
            {
                header = ReadHeaderFooterString(data);
            }
            else if (type == RecFooter && data.Length >= 2)
            {
                footer = ReadHeaderFooterString(data);
            }
        }
        return (landscape, paperSize, headerMargin, footerMargin, header, footer);
    }

    /// <summary>读取指定工作表是否受保护（PROTECT 记录）</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>是否受保护</returns>
    public Boolean GetProtection(String? sheet = null)
    {
        foreach (var (type, data) in ReadSheetRecords(sheet))
        {
            if (type == RecProtect && data.Length >= 2)
                return ReadUInt16(data, 0) != 0;
        }
        return false;
    }

    /// <summary>读取指定工作表的数据验证（DV 记录，支持下拉列表）</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>数据验证列表</returns>
    public List<DataValidation> GetValidations(String? sheet = null)
    {
        var result = new List<DataValidation>();
        foreach (var (type, data) in ReadSheetRecords(sheet))
        {
            if (type == RecDv) ParseDv(data, result);
        }
        return result;
    }

    /// <summary>读取指定工作表的所有记录（不解析单元格）</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    private List<(UInt16, Byte[])> ReadSheetRecords(String? sheet)
    {
        var idx = ResolveSheetIndex(sheet);
        if (idx < 0) return [];
        var sheetReader = new SpanReader(_workbook, 0, _workbook.Length);
        return ReadSheetRecords(ref sheetReader, _sheetBofOffsets[idx]);
    }

    private static void ParseDv(Byte[] data, List<DataValidation> result)
    {
        // DV 记录：fOption1(2)+fOption2(2)+ixfe(2)+r1(2)+c1(2)+r2(2)+c2(2)+titleCch(2)+promptCch(2)+title+prompt+cch1(2)+cch2(2)+formula1+formula2
        if (data.Length < 18) return;

        var option1 = ReadUInt16(data, 0);
        var type = option1 & 0x0F;
        var op = (option1 >> 4) & 0x0F;
        var r1 = ReadUInt16(data, 6);
        var c1 = ReadUInt16(data, 8);
        var r2 = ReadUInt16(data, 10);
        var c2 = ReadUInt16(data, 12);
        var titleCch = ReadUInt16(data, 14);
        var promptCch = ReadUInt16(data, 16);
        var pos = 18;

        // title / prompt（UnicodeStringNoCch：1 字节 fHighByte + 字符）
        pos = SkipDvString(data, pos, titleCch);
        pos = SkipDvString(data, pos, promptCch);
        if (pos + 4 > data.Length) return;

        var cch1 = ReadUInt16(data, pos);
        var cch2 = ReadUInt16(data, pos + 2);
        pos += 4;

        var dv = new DataValidation
        {
            CellRange = $"{ColumnName(c1)}{r1 + 1}:{ColumnName(c2)}{r2 + 1}",
            ValidationType = type switch { 1 => "whole", 2 => "decimal", 3 => "range", 4 => "date", 5 => "time", 6 => "textLength", _ => "list" },
            Operator = op switch { 1 => "notBetween", 2 => "equal", 3 => "notEqual", 4 => "greaterThan", 5 => "lessThan", 6 => "greaterThanOrEqual", 7 => "lessThanOrEqual", _ => "between" },
        };

        if (cch1 > 0 && pos + cch1 <= data.Length)
        {
            // 公式1：list 类型为带引号的逗号分隔 UTF-16LE 字符串
            var formula1 = Encoding.Unicode.GetString(data, pos, cch1);
            dv.Formula1 = formula1;
            if (type == 0)
            {
                var items = formula1.Trim('"').Split(',').Select(e => e.Trim()).Where(e => e.Length > 0).ToArray();
                if (items.Length > 0) dv.Items = items;
            }
        }
        pos += cch1;
        if (cch2 > 0 && pos + cch2 <= data.Length)
        {
            dv.Formula2 = Encoding.Unicode.GetString(data, pos, cch2);
        }

        result.Add(dv);
    }

    /// <summary>跳过 DV 记录中的 UnicodeStringNoCch 字符串，返回新的偏移</summary>
    private static Int32 SkipDvString(Byte[] data, Int32 pos, Int32 cch)
    {
        if (cch <= 0) return pos;
        if (pos >= data.Length) return pos;
        var fHighByte = (data[pos] & 0x01) != 0;
        return pos + 1 + cch * (fHighByte ? 2 : 1);
    }

    /// <summary>解析 HEADER/FOOTER 记录字符串（兼容本库写入格式与标准格式）</summary>
    private static String ReadHeaderFooterString(Byte[] data)
    {
        if (data.Length < 2) return String.Empty;

        // 本库写入格式：cch(2) + UTF-16LE（无标志字节）
        var cch2 = ReadUInt16(data, 0);
        if (data.Length == 2 + cch2 * 2)
            return Encoding.Unicode.GetString(data, 2, cch2 * 2);

        // 标准格式：cch(1) + grbit(1) + 字符
        var cch = data[0];
        var grbit = data.Length > 1 ? data[1] : (Byte)0;
        var pos = 2;
        if (cch == 0 || data.Length <= pos) return String.Empty;
        if ((grbit & 0x01) != 0)
            return Encoding.Unicode.GetString(data, pos, Math.Min(cch * 2, data.Length - pos));
        return DecodeLatin1(data, pos, Math.Min(cch, data.Length - pos));
    }

    /// <summary>列索引（0基）转列名（A/AA/AB）</summary>
    private static String ColumnName(Int32 col)
    {
        var name = String.Empty;
        var c = col + 1;
        while (c > 0)
        {
            c--;
            name = (Char)('A' + c % 26) + name;
            c /= 26;
        }
        return name;
    }
    #endregion

    #region 样式与命名范围读取
    /// <summary>读取指定工作表的单元格样式（FONT/FORMAT/XF 记录解析）</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>(行, 列) → CellFormat（0基），仅包含应用了非默认样式的单元格</returns>
    public Dictionary<(Int32 Row, Int32 Col), CellFormat> ReadCellFormats(String? sheet = null)
    {
        var result = new Dictionary<(Int32, Int32), CellFormat>();
        var idx = ResolveSheetIndex(sheet);
        if (idx < 0) return result;

        // 1. 解析 Globals 段样式表（FONT/FORMAT/XF）
        var fonts = new List<CellFormat>();
        var formats = new Dictionary<Int32, String>(); // formatIndex → 格式码
        var xfs = new List<XfEntry>();                 // xfIndex → 解析信息
        ParseGlobalStyles(fonts, formats, xfs);
        if (xfs.Count == 0) return result;

        // 2. 解析工作表单元格的 xfIndex
        var sheetReader = new SpanReader(_workbook, 0, _workbook.Length);
        var sheetRecords = ReadSheetRecords(ref sheetReader, _sheetBofOffsets[idx]);
        var cells = new Dictionary<(Int32, Int32), Int32>();
        foreach (var (type, data) in sheetRecords)
        {
            ParseCellXf(type, data, cells);
        }

        // 3. 组装 CellFormat
        foreach (var kv in cells)
        {
            var xf = kv.Value;
            if (xf <= 0 || xf >= xfs.Count) continue; // 0 = 默认样式，跳过
            var cf = BuildCellFormat(xfs[xf], fonts, formats);
            if (cf != null) result[kv.Key] = cf;
        }
        return result;
    }

    /// <summary>读取工作簿中所有命名范围（NAME 记录，不含系统 _xlnm.*）</summary>
    /// <returns>名称 → 范围引用（如 "Sheet1!$A$1:$B$10"）</returns>
    public Dictionary<String, String> GetDefinedNames()
    {
        var result = new Dictionary<String, String>();
        var reader = new SpanReader(_workbook, 0, _workbook.Length);
        var records = ReadAllRecords(ref reader);

        var globalsStart = FindGlobalsBof(records);
        if (globalsStart < 0) return result;

        var sheetNames = _sheetNames.Count > 0 ? _sheetNames[0] : "Sheet1";
        var pos = globalsStart;
        while (pos < records.Count)
        {
            var (type, data) = records[pos];
            pos++;
            if (type == RecEof) break;
            if (type == RecName && data.Length >= 14)
            {
                var (name, range) = ParseNameRecord(data, sheetNames);
                if (!name.IsNullOrEmpty() && !name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase) && !range.IsNullOrEmpty())
                    result[name!] = range!;
            }
        }
        return result;
    }

    /// <summary>读取第一个工作表的打印区域（_xlnm.Print_Area）</summary>
    /// <returns>范围（如 "A1:F20"），未设置返回 null</returns>
    public String? GetPrintArea()
    {
        var reader = new SpanReader(_workbook, 0, _workbook.Length);
        var records = ReadAllRecords(ref reader);
        var globalsStart = FindGlobalsBof(records);
        if (globalsStart < 0) return null;

        var sheetNames = _sheetNames.Count > 0 ? _sheetNames[0] : "Sheet1";
        var pos = globalsStart;
        while (pos < records.Count)
        {
            var (type, data) = records[pos];
            pos++;
            if (type == RecEof) break;
            if (type == RecName && data.Length >= 14)
            {
                var (name, range) = ParseNameRecord(data, sheetNames);
                if (name == "_xlnm.Print_Area") return range;
            }
        }
        return null;
    }

    /// <summary>解析 NAME 记录，返回（名称, 范围引用）</summary>
    private static (String? Name, String? Range) ParseNameRecord(Byte[] data, String defaultSheet)
    {
        // NAME 记录：option(2)+keyboard(1)+cch(1)+cce(2)+ixals(2)+reserved(2)+cce2(2)+cceFormula(1)+name+formula
        if (data.Length < 14) return (null, null);

        var option = ReadUInt16(data, 0);
        var cch = data[3];          // 名称字符数
        var cceFormula = data[12];  // 公式长度
        var namePos = 14;
        if (namePos + cch > data.Length) return (null, null);

        // 名称（Latin-1 或 UTF-16，取决于 option bit0）
        String name;
        if ((option & 0x01) != 0)
            name = Encoding.Unicode.GetString(data, namePos, cch * 2);
        else
            name = DecodeLatin1(data, namePos, cch);
        namePos += (option & 0x01) != 0 ? cch * 2 : cch;

        if (cceFormula <= 0 || namePos + cceFormula > data.Length) return (name, null);

        // 解析公式 Rgce：grbit(2) + ptgArea3d(1) + ixti(2) + r1(2)+c1(2)+r2(2)+c2(2)
        var fpos = namePos;
        var grbit = ReadUInt16(data, fpos);
        fpos += 2;
        if (fpos + 1 > data.Length) return (name, null);
        var ptg = data[fpos];
        fpos += 1;
        if (ptg is 0x3A or 0x3B) // ptgRef3d / ptgArea3d
        {
            fpos += 2; // skip ixti
            if (ptg == 0x3A)
            {
                if (fpos + 4 > data.Length) return (name, null);
                var r = ReadUInt16(data, fpos);
                var c = ReadUInt16(data, fpos + 2);
                return (name, $"{defaultSheet}!{ColumnName(c)}{r + 1}");
            }
            else
            {
                if (fpos + 8 > data.Length) return (name, null);
                var r1 = ReadUInt16(data, fpos);
                var c1 = ReadUInt16(data, fpos + 2);
                var r2 = ReadUInt16(data, fpos + 4);
                var c2 = ReadUInt16(data, fpos + 6);
                return (name, $"{defaultSheet}!{ColumnName(c1)}{r1 + 1}:{ColumnName(c2)}{r2 + 1}");
            }
        }
        return (name, null);
    }

    /// <summary>解析 Globals 段的 FONT/FORMAT/XF 记录</summary>
    private void ParseGlobalStyles(List<CellFormat> fonts, Dictionary<Int32, String> formats, List<XfEntry> xfs)
    {
        var reader = new SpanReader(_workbook, 0, _workbook.Length);
        var records = ReadAllRecords(ref reader);
        var globalsStart = FindGlobalsBof(records);
        if (globalsStart < 0) return;

        var pos = globalsStart;
        while (pos < records.Count)
        {
            var (type, data) = records[pos];
            pos++;
            if (type == RecEof) break;

            if (type == RecFont)
                fonts.Add(ParseFont(data));
            else if (type == RecFormat && data.Length >= 5)
            {
                var ifmt = ReadUInt16(data, 0);
                var cch = ReadUInt16(data, 2);
                var fHighByte = data[4];
                var code = (fHighByte & 0x01) != 0
                    ? Encoding.Unicode.GetString(data, 5, Math.Min(cch * 2, data.Length - 5))
                    : DecodeLatin1(data, 5, Math.Min(cch, data.Length - 5));
                formats[ifmt] = code;
            }
            else if (type == RecXf && data.Length >= 20)
                xfs.Add(ParseXf(data));
        }
    }

    /// <summary>解析 FONT 记录为字体近似样式</summary>
    private static CellFormat ParseFont(Byte[] data)
    {
        var cf = new CellFormat();
        if (data.Length < 16) return cf;

        var height = ReadUInt16(data, 0); // twips
        cf.FontSize = height / 20.0;
        var grbit = ReadUInt16(data, 2);
        cf.Italic = (grbit & 0x02) != 0;
        cf.Underline = (grbit & 0x0C) != 0;
        var weight = ReadUInt16(data, 6);
        cf.Bold = weight >= 600;
        var colorIdx = ReadUInt16(data, 4);
        cf.FontColor = MapColorIndexToRgb(colorIdx);

        var nameLen = data[14];
        var fHighByte = data[15];
        if (nameLen > 0 && 16 + nameLen * (fHighByte == 1 ? 2 : 1) <= data.Length)
        {
            cf.FontName = fHighByte == 1
                ? Encoding.Unicode.GetString(data, 16, nameLen * 2)
                : DecodeLatin1(data, 16, nameLen);
        }
        return cf;
    }

    /// <summary>解析 XF 记录（兼容标准 20 字节与本库 22 字节布局）</summary>
    private static XfEntry ParseXf(Byte[] data)
    {
        var xf = new XfEntry();
        xf.FontId = ReadUInt16(data, 0);
        xf.FormatId = ReadUInt16(data, 2);

        // fAlign（标准 20 字节布局与 22 字节布局均在偏移 6）
        var fAlign = ReadUInt16(data, 6);
        xf.Wrap = (fAlign & 0x01) != 0;
        xf.HAlign = (HorizontalAlignment)((fAlign >> 3) & 0x07);
        xf.VAlign = (VerticalAlignment)((fAlign >> 9) & 0x03);

        // 边框样式 nibble：标准布局在偏移 10，本库 22 字节布局在偏移 14
        var borderPos = data.Length == 22 ? 14 : 10;
        if (borderPos + 4 <= data.Length)
        {
            xf.LeftBorder = (BorderStyle)(data[borderPos] & 0x0F);
            xf.RightBorder = (BorderStyle)((data[borderPos] >> 4) & 0x0F);
            xf.TopBorder = (BorderStyle)(data[borderPos + 1] & 0x0F);
            xf.BottomBorder = (BorderStyle)((data[borderPos + 1] >> 4) & 0x0F);
        }

        // 填充图案：标准布局在偏移 18；本库 22 字节布局末字节为 0x40/0
        if (data.Length == 22)
        {
            xf.FillPattern = (UInt32)data[21];
        }
        else if (data.Length >= 20)
        {
            xf.FillPattern = data[18];
            // 标准布局填充色：16-17 前景，18 图案（颜色索引在 palette，best-effort）
            xf.FgColor = ReadUInt16(data, 16);
        }
        return xf;
    }

    /// <summary>从单元格记录提取 (行, 列) → xfIndex</summary>
    private static void ParseCellXf(UInt16 type, Byte[] data, Dictionary<(Int32, Int32), Int32> cells)
    {
        if (data.Length < 6) return;
        var row = (Int32)ReadUInt16(data, 0);
        var col = (Int32)ReadUInt16(data, 2);
        var xf = (Int32)ReadUInt16(data, 4);

        // MULRK 的 xf 在每条记录内；LABELSST/NUMBER/RK/BOOLERR/LABEL/FORMULA/BLANK 的 xf 在偏移 4
        if (type == RecMulRk)
        {
            var firstCol = col;
            var lastCol = (Int32)ReadUInt16(data, data.Length - 2);
            var p = 6;
            for (var c = firstCol; c <= lastCol; c++)
            {
                if (p + 6 > data.Length - 2) break;
                var cx = (Int32)ReadUInt16(data, p);
                cells[(row, c)] = cx;
                p += 6;
            }
        }
        else
        {
            cells[(row, col)] = xf;
        }
    }

    /// <summary>组装 CellFormat（跳过与默认样式无差异的项）</summary>
    private static CellFormat? BuildCellFormat(XfEntry xf, List<CellFormat> fonts, Dictionary<Int32, String> formats)
    {
        var cf = new CellFormat();
        var changed = false;

        if (xf.FontId > 0 && xf.FontId < fonts.Count) // 字体 0 为默认字体，不构成自定义样式
        {
            var f = fonts[xf.FontId];
            if (!f.FontName.IsNullOrEmpty()) { cf.FontName = f.FontName; changed = true; }
            if (f.FontSize > 0) { cf.FontSize = f.FontSize; changed = true; }
            if (f.Bold) { cf.Bold = true; changed = true; }
            if (f.Italic) { cf.Italic = true; changed = true; }
            if (!f.FontColor.IsNullOrEmpty()) { cf.FontColor = f.FontColor; changed = true; }
        }

        if (formats.TryGetValue(xf.FormatId, out var fmt) && !fmt.IsNullOrEmpty())
        {
            cf.NumberFormat = fmt;
            changed = true;
        }

        if (xf.Wrap) { cf.WrapText = true; changed = true; }
        if (xf.HAlign != HorizontalAlignment.General) { cf.HAlign = xf.HAlign; changed = true; }
        if (xf.VAlign != VerticalAlignment.Top) { cf.VAlign = xf.VAlign; changed = true; }
        if (xf.LeftBorder != BorderStyle.None) { cf.LeftBorder = xf.LeftBorder; changed = true; }
        if (xf.RightBorder != BorderStyle.None) { cf.RightBorder = xf.RightBorder; changed = true; }
        if (xf.TopBorder != BorderStyle.None) { cf.TopBorder = xf.TopBorder; changed = true; }
        if (xf.BottomBorder != BorderStyle.None) { cf.BottomBorder = xf.BottomBorder; changed = true; }
        if (xf.FillPattern is > 0 and not 0x40) { cf.BackgroundColor = MapColorIndexToRgb(xf.FgColor); changed = true; }

        return changed ? cf : null;
    }

    /// <summary>BIFF8 常用颜色索引 → RGB（调色板近似）</summary>
    private static String? MapColorIndexToRgb(UInt16 idx)
    {
        if (idx == 0x7FFF || idx == 64) return null; // 自动/系统
        var palette = new[] { "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
            "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF" };
        if (idx < palette.Length) return palette[idx];
        return null;
    }

    /// <summary>XF 记录解析结果</summary>
    private sealed class XfEntry
    {
        public Int32 FontId;
        public Int32 FormatId;
        public Boolean Wrap;
        public HorizontalAlignment HAlign;
        public VerticalAlignment VAlign;
        public BorderStyle LeftBorder;
        public BorderStyle RightBorder;
        public BorderStyle TopBorder;
        public BorderStyle BottomBorder;
        public UInt32 FillPattern;
        public UInt16 FgColor;
    }
    #endregion

    #region 条件格式读取
    /// <summary>读取指定工作表的条件格式（CONDFMT + CF 记录）</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>条件格式列表（范围/类型/数值条件）</returns>
    /// <remarks>当前支持 cellIs（数值比较）与 formula（表达式）类型；colorScale/dataBar/iconSet 复杂格式暂跳过。</remarks>
    public List<ConditionalFormatting> GetConditionalFormats(String? sheet = null)
    {
        var result = new List<ConditionalFormatting>();
        var idx = ResolveSheetIndex(sheet);
        if (idx < 0) return result;

        var sheetReader = new SpanReader(_workbook, 0, _workbook.Length);
        var sheetRecords = ReadSheetRecords(ref sheetReader, _sheetBofOffsets[idx]);

        var pendingCount = 0;
        var pendingRange = String.Empty;

        foreach (var (type, data) in sheetRecords)
        {
            if (type == RecCondFmt && data.Length >= 12)
            {
                // CONDFMT：count(2)+r1(2)+r2(2)+c1(2)+c2(2)+flags(2)
                var count = ReadUInt16(data, 0);
                var r1 = ReadUInt16(data, 2);
                var r2 = ReadUInt16(data, 4);
                var c1 = ReadUInt16(data, 6);
                var c2 = ReadUInt16(data, 8);
                pendingCount = count;
                pendingRange = $"{ColumnName(c1)}{r1 + 1}:{ColumnName(c2)}{r2 + 1}";
            }
            else if (type == RecCf && pendingCount > 0)
            {
                pendingCount--;
                var cf = ParseCf(data, pendingRange);
                if (cf != null) result.Add(cf);
            }
        }
        return result;
    }

    private static ConditionalFormatting? ParseCf(Byte[] data, String range)
    {
        // CF：flags(2)+f1(8)+f2(8)+字体(4)+边框(4)+图案(4)+temp(2)[+公式]
        if (data.Length < 18) return null;

        var flags = ReadUInt16(data, 0);
        var type = flags & 0x0F;
        var op = (flags >> 4) & 0x0F;
        var f1 = BitConverter.Int64BitsToDouble(BitConverter.ToInt64(data, 2));
        var f2 = BitConverter.Int64BitsToDouble(BitConverter.ToInt64(data, 10));

        var info = new ConditionalFormatting { Range = range };
        if (type == 1) // cellIs
        {
            info.Type = op switch
            {
                1 => ConditionalFormatValues.NotBetween,
                2 => ConditionalFormatValues.Equal,
                3 => ConditionalFormatValues.NotEqual,
                4 => ConditionalFormatValues.GreaterThan,
                5 => ConditionalFormatValues.LessThan,
                _ => ConditionalFormatValues.Between,
            };
            info.Value = FormatCfValue(f1);
            info.Value2 = FormatCfValue(f2);
        }
        else if (type == 2) // formula
        {
            info.Type = ConditionalFormatValues.Expression;
            // 公式为 Rgce 编码（token），固定部分之后；best-effort 读取
            if (data.Length > 32)
                info.Formula = Encoding.UTF8.GetString(data, 32, data.Length - 32).TrimEnd('\0');
        }
        else return null; // colorScale/dataBar/iconSet 复杂，跳过

        // 图案样式（offset 26）：pattern(1)+fgColor(2)+bgColor(1)，前景色索引映射调色板还原 Color
        if (data.Length >= 30)
        {
            var fgIdx = ReadUInt16(data, 27);
            var fgColor = MapColorIndexToRgb(fgIdx);
            if (!fgColor.IsNullOrEmpty()) info.Color = fgColor;
        }

        return info;
    }

    /// <summary>格式化 CF 条件值（整数值不带小数）</summary>
    private static String FormatCfValue(Double value)
    {
        if (Math.Abs(value - Math.Truncate(value)) < 0.0001)
            return ((Int64)value).ToString();
        return value.ToString(CultureInfo.InvariantCulture);
    }
    #endregion

    #region 批注读取
    /// <summary>读取指定工作表的批注（NOTE + TXO 记录）</summary>
    /// <param name="sheet">工作表名称，null 取第一个</param>
    /// <returns>(0基行, 0基列) → 批注（文本/作者）</returns>
    /// <remarks>
    /// NOTE 记录含位置与作者；批注文本在 TXO 记录中。同一批注的 NOTE/TXO 按流顺序配对（简单文件 1:1）。
    /// </remarks>
    public Dictionary<(Int32 Row, Int32 Col), Comment> GetComments(String? sheet = null)
    {
        var result = new Dictionary<(Int32, Int32), Comment>();
        var idx = ResolveSheetIndex(sheet);
        if (idx < 0) return result;

        var sheetReader = new SpanReader(_workbook, 0, _workbook.Length);
        var sheetRecords = ReadSheetRecords(ref sheetReader, _sheetBofOffsets[idx]);

        var notes = new List<(Int32 Row, Int32 Col, String Author)>();
        var texts = new List<String>();

        foreach (var (type, data) in sheetRecords)
        {
            if (type == RecNote && data.Length >= 11)
            {
                // NOTE：row(2)+col(2)+objectId(2)+flags(2)+cch(2)+fHighByte(1)+作者
                var row = (Int32)ReadUInt16(data, 0);
                var col = (Int32)ReadUInt16(data, 2);
                var cch = ReadUInt16(data, 8);
                var fHighByte = data[10];
                var author = fHighByte != 0
                    ? Encoding.Unicode.GetString(data, 11, Math.Min(cch * 2, data.Length - 11))
                    : DecodeLatin1(data, 11, Math.Min(cch, data.Length - 11));
                notes.Add((row, col, author));
            }
            else if (type == RecTxo && data.Length >= 12)
            {
                // TXO：grbit(2)+rot(2)+cchText(2)+cbRuns(2)+ifntEmpty(2)+reserved(2)+文本(UTF-16LE)
                var cchText = ReadUInt16(data, 4);
                var text = cchText > 0
                    ? Encoding.Unicode.GetString(data, 12, Math.Min(cchText * 2, data.Length - 12))
                    : String.Empty;
                texts.Add(text);
            }
        }

        // NOTE 与 TXO 按流顺序配对（每个批注恰好一个 TXO）
        for (var i = 0; i < notes.Count; i++)
        {
            var (row, col, author) = notes[i];
            var text = i < texts.Count ? texts[i] : String.Empty;
            result[(row, col)] = new Comment { Text = text, Author = author };
        }
        return result;
    }
    #endregion

    #region 单元格解析辅助
    private void ParseLabelSst(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+sstIndex(4) = 共10字节
        if (data.Length < 10) return;
        var reader = new SpanReader(data, 0, data.Length);
        var row = (Int32)reader.ReadUInt16();
        var col = (Int32)reader.ReadUInt16();
        reader.Advance(2); // skip xf
        var sstIdx = (Int32)reader.ReadUInt32();
        var value = sstIdx >= 0 && sstIdx < _sst.Length ? (Object?)_sst[sstIdx] : null;
        SetCell(cells, row, col, value, ref maxCol);
    }

    private static void ParseNumber(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+double(8) = 14字节
        if (data.Length < 14) return;
        var reader = new SpanReader(data, 0, data.Length);
        var row = (Int32)reader.ReadUInt16();
        var col = (Int32)reader.ReadUInt16();
        reader.Advance(2); // skip xf
        var value = reader.ReadDouble();
        SetCell(cells, row, col, (Object?)value, ref maxCol);
    }

    private static void ParseRk(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+rk(4) = 10字节
        if (data.Length < 10) return;
        var reader = new SpanReader(data, 0, data.Length);
        var row = (Int32)reader.ReadUInt16();
        var col = (Int32)reader.ReadUInt16();
        reader.Advance(2); // skip xf
        var rk = reader.ReadInt32();
        var value = DecodeRk(rk);
        SetCell(cells, row, col, (Object?)value, ref maxCol);
    }

    private static void ParseMulRk(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+firstCol(2)+[xf(2)+rk(4)]*n+lastCol(2)
        if (data.Length < 6) return;
        var reader = new SpanReader(data, 0, data.Length);
        var row = (Int32)reader.ReadUInt16();
        var firstCol = (Int32)reader.ReadUInt16();
        var lastCol = (Int32)new SpanReader(data, data.Length - 2, 2).ReadUInt16();
        var count = lastCol - firstCol + 1;
        for (var i = 0; i < count; i++)
        {
            if (reader.Position + 6 > data.Length - 2) break;
            reader.Advance(2); // skip xf
            var rk = reader.ReadInt32();
            var value = DecodeRk(rk);
            SetCell(cells, row, firstCol + i, (Object?)value, ref maxCol);
        }
    }

    private static void ParseBoolErr(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+boolOrErr(1)+isError(1) = 8字节
        if (data.Length < 8) return;
        var reader = new SpanReader(data, 0, data.Length);
        var row = (Int32)reader.ReadUInt16();
        var col = (Int32)reader.ReadUInt16();
        reader.Advance(2); // skip xf
        var boolOrErr = reader.ReadByte();
        var isError = reader.ReadByte() != 0;
        if (!isError)
        {
            SetCell(cells, row, col, (Object?)(boolOrErr != 0), ref maxCol);
        }
        // 错误值暂时跳过
    }

    private static void ParseLabel(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+cch(2)+fHighByte(1)+chars
        if (data.Length < 9) return;
        var reader = new SpanReader(data, 0, data.Length);
        var row = (Int32)reader.ReadUInt16();
        var col = (Int32)reader.ReadUInt16();
        reader.Advance(2); // skip xf
        var cch = (Int32)reader.ReadUInt16();
        var fHighByte = reader.ReadByte();
        String value;
        if (fHighByte != 0)
            value = Encoding.Unicode.GetString(data, 9, Math.Min(cch * 2, data.Length - 9));
        else
            value = DecodeLatin1(data, 9, Math.Min(cch, data.Length - 9));
        SetCell(cells, row, col, (Object?)value, ref maxCol);
    }

    private static void ParseFormula(Byte[] data, SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells, ref Int32 maxCol)
    {
        // row(2)+col(2)+xf(2)+num(8)+grbit(2)+chn(4)+formulaLen(2)+formulaBytes
        if (data.Length < 22) return;
        var reader = new SpanReader(data, 0, data.Length);
        var row = (Int32)reader.ReadUInt16();
        var col = (Int32)reader.ReadUInt16();
        reader.Advance(12); // skip xf(2)+num(8)+grbit(2)
        reader.Advance(4);  // skip chn(4)
        var formulaLen = (Int32)reader.ReadUInt16();
        if (formulaLen > 0 && 22 + formulaLen <= data.Length)
        {
            var formulaText = Encoding.UTF8.GetString(data, 22, formulaLen);
            SetCell(cells, row, col, (Object?)formulaText, ref maxCol);
        }
        else
        {
            SetCell(cells, row, col, (Object?)"=?", ref maxCol);
        }
    }

    private static void SetCell(SortedDictionary<Int32, SortedDictionary<Int32, Object?>> cells,
        Int32 row, Int32 col, Object? value, ref Int32 maxCol)
    {
        if (!cells.TryGetValue(row, out var rowDict))
        {
            rowDict = [];
            cells[row] = rowDict;
        }
        rowDict[col] = value;
        if (col > maxCol) maxCol = col;
    }

    /// <summary>解析 HYPERLINK 记录 (0x01B8)</summary>
    private void ParseHyperlink(Byte[] data)
    {
        if (data.Length < 30) return;
        var reader = new SpanReader(data, 0, data.Length);
        var firstRow = (Int32)reader.ReadUInt16();
        var lastRow = (Int32)reader.ReadUInt16();
        var firstCol = (Int32)reader.ReadUInt16();
        var lastCol = (Int32)reader.ReadUInt16();
        reader.Advance(20); // skip 16-byte GUID + 4-byte options
        var urlLen = (Int32)reader.ReadUInt16();
        if (urlLen > 0 && 30 + urlLen <= data.Length)
        {
            var url = Encoding.UTF8.GetString(data, 30, urlLen);
            _hyperlinks.Add((firstRow, firstCol, url));
        }
    }

    /// <summary>获取已解析的超链接列表</summary>
    /// <returns>超链接列表（行、列、URL）</returns>
    public IReadOnlyList<(Int32 Row, Int32 Col, String Url)> GetHyperlinks() => _hyperlinks.AsReadOnly();

    /// <summary>获取已解析的列宽（列索引 → 宽度，以字符宽度为单位）</summary>
    /// <returns>列宽字典的只读包装</returns>
    public IReadOnlyDictionary<Int32, Double> GetColumnWidths() => new Dictionary<Int32, Double>(_columnWidths);

    private void ParseColInfo(Byte[] data)
    {
        if (data.Length < 10) return;
        var reader = new SpanReader(data, 0, data.Length);
        var firstCol = (Int32)reader.ReadUInt16();
        var lastCol = (Int32)reader.ReadUInt16();
        var widthRaw = reader.ReadUInt16(); // 1/256 of character width
        reader.Advance(4); // skip xfIndex(2) + flags(2)
        var width = widthRaw / 256.0;
        for (var c = firstCol; c <= lastCol; c++)
            _columnWidths[c] = width;
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
            value = BitConverter.Int64BitsToDouble(d64);
        }

        // bit 1 = 0 表示乘以 100，否则不处理（注意：BIFF8 中 bit1=1 表示 ÷100）
        if ((rk & 0x01) != 0)
            value /= 100.0;

        return value;
    }
    #endregion

    #region 工作表记录读取
    /// <summary>从指定文件流偏移读取工作表段的所有记录（到 EOF）</summary>
    /// <param name="reader">工作簿读取器</param>
    /// <param name="bofOffset">工作表 BOF 在字节流中的偏移</param>
    /// <returns>记录列表</returns>
    private static List<(UInt16, Byte[])> ReadSheetRecords(ref SpanReader reader, Int32 bofOffset)
    {
        reader.Position = bofOffset;
        var result = new List<(UInt16, Byte[])>();
        var depth = 0;

        while (reader.Available >= 4)
        {
            var type = reader.ReadUInt16();
            var len = (Int32)reader.ReadUInt16();
            if (reader.Available < len) break;

            var data = reader.ReadBytes(len).ToArray();

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
        {
            if (_sheetNames[i].Equals(name, StringComparison.OrdinalIgnoreCase))
                return i;
        }
        return -1;
    }
    #endregion

    #region 字节工具
    private static UInt16 ReadUInt16(Byte[] buf, Int32 pos)
    {
        var reader = new SpanReader(buf.AsSpan(pos));
        return reader.ReadUInt16();
    }

    private static UInt32 ReadUInt32(Byte[] buf, Int32 pos)
    {
        var reader = new SpanReader(buf.AsSpan(pos));
        return reader.ReadUInt32();
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
        {
            chars[i] = (Char)data[pos + i];
        }
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
    private const UInt16 RecHyperlink = 0x01B8;
    private const UInt16 RecRk = 0x027E;
    private const UInt16 RecMulRk = 0x00BE;
    private const UInt16 RecBoolErr = 0x0205;
    private const UInt16 RecLabel = 0x0204;
    private const UInt16 RecBlank = 0x0201;
    private const UInt16 RecFormula = 0x0006;
    private const UInt16 RecMulBlank = 0x00BF;
    private const UInt16 RecColInfo = 0x007D;
    private const UInt16 RecMergedCells = 0x00E5;
    private const UInt16 RecSetup = 0x00A1;
    private const UInt16 RecHeader = 0x0014;
    private const UInt16 RecFooter = 0x0015;
    private const UInt16 RecProtect = 0x0012;
    private const UInt16 RecDv = 0x01BE;
    private const UInt16 RecFont = 0x0031;
    private const UInt16 RecFormat = 0x041E;
    private const UInt16 RecXf = 0x00E0;
    private const UInt16 RecName = 0x0018;
    private const UInt16 RecNote = 0x001C;
    private const UInt16 RecTxo = 0x01B6;
    private const UInt16 RecCondFmt = 0x01B0;
    private const UInt16 RecCf = 0x01B1;
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（CSV 格式，逗号分隔）</summary>
    /// <returns>CSV 格式文本</returns>
    public String? ExtractText()
    {
        if (_sheetNames.Count == 0) return null;

        var sb = new StringBuilder();
        for (var si = 0; si < _sheetNames.Count; si++)
        {
            var sheetName = _sheetNames[si];
            if (_sheetNames.Count > 1)
            {
                if (si > 0) sb.AppendLine();
                sb.AppendLine($"## {sheetName}");
            }

            foreach (var row in ReadSheet(sheetName))
            {
                for (var i = 0; i < row.Length; i++)
                {
                    if (i > 0) sb.Append(',');
                    sb.Append(CsvEscape(row[i]?.ToString()));
                }
                sb.AppendLine();
            }
        }
        return sb.ToString();
    }

    /// <summary>提取 Markdown 格式（表格）</summary>
    /// <returns>Markdown 表格字符串</returns>
    public String? ExtractMarkdown()
    {
        if (_sheetNames.Count == 0) return null;

        var sb = new StringBuilder();
        for (var si = 0; si < _sheetNames.Count; si++)
        {
            var sheetName = _sheetNames[si];
            if (_sheetNames.Count > 1)
            {
                if (si > 0) sb.AppendLine();
                sb.AppendLine($"## {sheetName}");
                sb.AppendLine();
            }

            var rows = ReadSheet(sheetName).ToList();
            if (rows.Count == 0) continue;

            // 第一行作为表头
            var header = rows[0];
            sb.Append('|');
            foreach (var cell in header)
            {
                sb.Append(' ').Append(MdEscape(cell?.ToString())).Append(" |");
            }
            sb.AppendLine();

            // 分隔线
            sb.Append('|');
            for (var i = 0; i < header.Length; i++)
            {
                sb.Append(" --- |");
            }
            sb.AppendLine();

            // 数据行
            for (var ri = 1; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                sb.Append('|');
                for (var i = 0; i < header.Length; i++)
                {
                    var val = i < row.Length ? row[i]?.ToString() : "";
                    sb.Append(' ').Append(MdEscape(val)).Append(" |");
                }
                sb.AppendLine();
            }
        }
        return sb.ToString();
    }

    private static String CsvEscape(String? value)
    {
        if (String.IsNullOrEmpty(value)) return "";
        if (value!.IndexOfAny([',', '"', '\n', '\r']) >= 0)
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }

    private static String MdEscape(String? value)
    {
        if (String.IsNullOrEmpty(value)) return "";
        return value!.Replace("|", "\\|").Replace("\n", " ").Replace("\r", "");
    }
    #endregion
}
