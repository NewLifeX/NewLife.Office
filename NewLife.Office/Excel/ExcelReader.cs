using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace NewLife.Office;

/// <summary>轻量级Excel读取器，支持读写往返</summary>
/// <remarks>
/// 文档 https://newlifex.com/core/excel_reader
/// 仅支持xlsx格式，本质上是压缩包，内部xml。
/// 支持完整读取：共享字符串、样式（字体/填充/边框/对齐/数字格式）、单元格数据、
/// 合并区域、超链接、图片、页面设置、条件格式、批注、数据验证、公式等。
/// 通过 <see cref="ReadExcel"/> 一键获取工作簿完整快照。
/// </remarks>
public class ExcelReader : DisposeBase, ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>文件名</summary>
    public String? FileName { get; }

    /// <summary>工作表集合（键为工作表名称）</summary>
    public ICollection<String>? Sheets => _orderedSheets;

    private ZipArchive _zip;
    private String[]? _sharedStrings;
    private ExcelNumberFormat?[]? _styles;
    private IDictionary<String, ZipArchiveEntry>? _entries;
    private List<String>? _orderedSheets;

    // 完整样式解析
    private List<FontInfo>? _fontInfos;
    private List<FillInfo>? _fillInfos;
    private List<BorderInfo>? _borderInfos;
    private List<XfInfo>? _xfInfos;
    private Dictionary<Int32, String>? _numFmtCodes;
    #endregion

    #region 内部类型
    private class FontInfo
    {
        public String? Name;
        public Double Size;
        public Boolean Bold;
        public Boolean Italic;
        public Boolean Underline;
        public String? Color;
    }

    private class FillInfo
    {
        public String? BgColor;
        public String? PatternType = "none";
    }

    private class BorderInfo
    {
        public CellBorderStyle Style;
        public String? Color;
    }

    private class XfInfo
    {
        public Int32 NumFmtId;
        public Int32 FontId;
        public Int32 FillId;
        public Int32 BorderId;
        public HorizontalAlignment HAlign;
        public VerticalAlignment VAlign;
        public Boolean WrapText;
    }
    #endregion

    #region 构造
    /// <summary>实例化读取器</summary>
    /// <param name="fileName">Excel文件路径（xlsx）</param>
    public ExcelReader(String fileName)
    {
        if (fileName.IsNullOrEmpty()) throw new ArgumentNullException(nameof(fileName));

        FileName = fileName;

        //_zip = ZipFile.OpenRead(fileName.GetFullPath());
        // 共享访问，避免文件被其它进程打开时再次访问抛出异常
        var fs = new FileStream(fileName.GetFullPath(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        _zip = new ZipArchive(fs, ZipArchiveMode.Read, false);

        Parse();
    }

    /// <summary>实例化读取器</summary>
    /// <param name="stream">Excel数据流（需保持可读，调用方负责其生命周期）</param>
    /// <param name="encoding">压缩文件内各xml条目的编码（一般为UTF-8）</param>
    public ExcelReader(Stream stream, Encoding encoding)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        if (stream is FileStream fs) FileName = fs.Name;

        _zip = new ZipArchive(stream, ZipArchiveMode.Read, true, encoding);

        Parse();
    }

    /// <summary>销毁</summary>
    /// <param name="disposing"></param>
    protected override void Dispose(Boolean disposing)
    {
        base.Dispose(disposing);

        _entries?.Clear();
        _zip?.Dispose();
    }
    #endregion

    #region 方法
    private void Parse()
    {
        // 读取共享字符串（可缺失）
        {
            var entry = _zip.GetEntry("xl/sharedStrings.xml");
            if (entry != null)
            {
                using var es = entry.Open(); // 确保及时释放，避免后续再打开时报本地文件头损坏
                _sharedStrings = ReadStrings(es);
            }
        }

        // 读取样式（含数字格式 + 完整字体/填充/边框/XF）
        {
            var entry = _zip.GetEntry("xl/styles.xml");
            if (entry != null)
            {
                using var es = entry.Open();
                _styles = ReadStyles(es);
            }
        }

        // 读取sheet条目索引
        {
            _entries = ReadSheets(_zip);
        }

        // 二次解析完整样式（需要重新读取 styles.xml，因为流已关闭）
        {
            var entry = _zip.GetEntry("xl/styles.xml");
            if (entry != null)
            {
                using var es = entry.Open();
                ParseFullStyles(es);
            }
        }
    }

    private static DateTime _1900 = new(1900, 1, 1);

    /// <summary>逐行读取数据，第一行通常是表头。支持超过26列（AA/AB等）以及缺失列自动补 null。</summary>
    /// <param name="sheet">工作表名。默认 null 取第一个数据表</param>
    /// <returns>按行返回对象数组。根据样式尝试转换为 DateTime / TimeSpan / 数值 / 布尔，否则为字符串</returns>
    public IEnumerable<Object?[]> ReadRows(String? sheet = null)
    {
        ThrowIfDisposed();

        if (Sheets == null || _entries == null) yield break;

        if (sheet.IsNullOrEmpty()) sheet = Sheets.FirstOrDefault();
        if (sheet.IsNullOrEmpty()) throw new ArgumentNullException(nameof(sheet));

        if (!_entries.TryGetValue(sheet, out var entry)) throw new ArgumentOutOfRangeException(nameof(sheet), "Unable to find worksheet");

        using var esheet = entry.Open(); // 及时释放单个 sheet 流
        var doc = XDocument.Load(esheet);
        if (doc.Root == null) yield break;

        var data = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("sheetData"));
        if (data == null) yield break;

        // 加快样式判断速度
        var styles = _styles;
        if (styles != null && styles.Length == 0) styles = null;

        var headerColumnCount = -1; // 记录首行列数，用于补齐后续行尾部缺失列

        foreach (var row in data.Elements())
        {
            var vs = new List<Object?>();
            var curIndex = 0; // 当前列（0基）
            foreach (var col in row.Elements())
            {
                // 单元格引用。例如 A1 / AB23
                var r = col.Attribute("r")?.Value;
                if (!r.IsNullOrEmpty())
                {
                    var targetIndex = GetColumnIndex(r!); // 0基
                    // 补齐缺失列
                    while (curIndex < targetIndex)
                    {
                        vs.Add(null);
                        curIndex++;
                    }
                }

                // 默认原始值。优先取 <v> 子节点（统一行为），否则使用节点聚合值
                Object? val = null;
                var vNode = col.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
                if (vNode != null)
                    val = vNode.Value;
                else
                    val = col.Value; // inlineStr 等情况会走这里

                // t=DataType: s=SharedString, b=Boolean, n=Number(默认), d=Date(较少出现), str=公式结果文本, inlineStr=内联字符串
                var t = col.Attribute("t")?.Value;
                if (t == "s")
                {
                    // 共享字符串
                    if (val is String s2 && Int32.TryParse(s2, out var sharedIndex)) val = _sharedStrings != null && sharedIndex >= 0 && sharedIndex < _sharedStrings.Length ? _sharedStrings[sharedIndex] : null;
                }
                else if (t == "b")
                {
                    // 布尔：0 / 1 以及 true / false
                    if (val is String sb) val = sb == "1" || sb.EqualIgnoreCase("true");
                }
                else if (t == "inlineStr")
                {
                    // 已经在 col.Value 中
                }
                else if (t == "str")
                {
                    // 公式结果文本，不再特别处理
                }

                // 样式转换（日期 / 时间 / 数字）。仅当未被布尔/共享字符串提前转换
                if (val is String && styles != null)
                {
                    var sAttr = col.Attribute("s"); // StyleIndex
                    if (sAttr != null)
                    {
                        var si = sAttr.Value.ToInt();
                        if (si >= 0 && si < styles.Length)
                        {
                            // 按引用格式转换数值，没有引用格式时不转换
                            var st = styles[si];
                            if (st != null) val = ChangeType(val, st);
                        }
                    }
                    else if (t.IsNullOrEmpty())
                    {
                        // OOXML 规定：无 t 属性且无 s 属性时，<v> 值默认为数字（General 格式）
                        if (val is String rawStr)
                        {
                            if (Int32.TryParse(rawStr, out var ni)) val = ni;
                            else if (Int64.TryParse(rawStr, out var nl)) val = nl;
                            else if (Double.TryParse(rawStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var nd)) val = nd;
                        }
                    }
                }

                vs.Add(val);
                curIndex++; // 移动到下一列
            }

            // 记录首行列数
            if (headerColumnCount == -1)
            {
                headerColumnCount = vs.Count;
            }
            else if (headerColumnCount > 0 && vs.Count < headerColumnCount)
            {
                // 补齐尾部缺失列（例如数据行末尾空值未写入单元格）
                while (vs.Count < headerColumnCount) vs.Add(null);
            }

            yield return vs.ToArray();
        }
    }

    /// <summary>逐行读取数据，同时返回每行的实际 Excel 行号（0基）</summary>
    private IEnumerable<(Int32 RowIdx, Object?[] Row)> ReadRowsWithIndex(String? sheet = null)
    {
        ThrowIfDisposed();

        if (Sheets == null || _entries == null) yield break;

        if (sheet.IsNullOrEmpty()) sheet = Sheets.FirstOrDefault();
        if (sheet.IsNullOrEmpty()) throw new ArgumentNullException(nameof(sheet));

        if (!_entries.TryGetValue(sheet, out var entry)) yield break;

        using var esheet = entry.Open();
        var doc = XDocument.Load(esheet);
        if (doc.Root == null) yield break;

        var dataEl = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("sheetData"));
        if (dataEl == null) yield break;

        var styles = _styles;
        if (styles != null && styles.Length == 0) styles = null;
        var headerColumnCount = -1;

        foreach (var row in dataEl.Elements())
        {
            var rAttr = row.Attribute("r");
            if (rAttr == null) continue;
            var rowIdx = rAttr.Value.ToInt(-1) - 1; // 转为 0 基
            if (rowIdx < 0) continue;

            var vs = new List<Object?>();
            var curIndex = 0;
            foreach (var col in row.Elements())
            {
                var r = col.Attribute("r")?.Value;
                if (!r.IsNullOrEmpty())
                {
                    var targetIndex = GetColumnIndex(r!);
                    while (curIndex < targetIndex) { vs.Add(null); curIndex++; }
                }

                Object? val = null;
                var vNode = col.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
                val = vNode != null ? vNode.Value : col.Value;

                var t = col.Attribute("t")?.Value;
                if (t == "s")
                {
                    if (val is String s2 && Int32.TryParse(s2, out var sharedIndex))
                        val = _sharedStrings != null && sharedIndex >= 0 && sharedIndex < _sharedStrings.Length ? _sharedStrings[sharedIndex] : null;
                }
                else if (t == "b")
                {
                    if (val is String sb) val = sb == "1" || sb.EqualIgnoreCase("true");
                }

                if (val is String && styles != null)
                {
                    var sAttr = col.Attribute("s");
                    if (sAttr != null)
                    {
                        var si = sAttr.Value.ToInt();
                        if (si >= 0 && si < styles.Length)
                        {
                            var st = styles[si];
                            if (st != null) val = ChangeType(val, st);
                        }
                    }
                    else if (t.IsNullOrEmpty())
                    {
                        if (val is String rawStr)
                        {
                            if (Int32.TryParse(rawStr, out var ni)) val = ni;
                            else if (Int64.TryParse(rawStr, out var nl)) val = nl;
                            else if (Double.TryParse(rawStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var nd)) val = nd;
                        }
                    }
                }

                vs.Add(val);
                curIndex++;
            }

            if (headerColumnCount == -1)
                headerColumnCount = vs.Count;
            else if (headerColumnCount > 0 && vs.Count < headerColumnCount)
                while (vs.Count < headerColumnCount) vs.Add(null);

            yield return (rowIdx, vs.ToArray());
        }
    }

    /// <summary>按 Excel 数字格式尝试转换值</summary>
    private Object? ChangeType(Object? val, ExcelNumberFormat st)
    {
        // 日期格式。Excel 以 1900-1-1 为基准(含虚构闰年Bug)，序列值 1 = 1900-01-01。这里减2 与历史实现保持兼容。
        if (st.Format.Contains("yy") || st.Format.Contains("mmm") || st.NumFmtId >= 14 && st.NumFmtId <= 17 || st.NumFmtId == 22)
        {
            if (val is String str && Double.TryParse(str, out var d))
            {
                // 暂时不明白为何要减2，实际上这么做就对了
                //val = _1900.AddDays(str.ToDouble() - 2);
                // 取整秒，剔除毫秒部分，避免浮点误差
                val = _1900.AddSeconds(Math.Round((d - 2) * 24 * 3600));
                //var ss = str.Split('.');
                //var dt = _1900.AddDays(ss[0].ToInt() - 2);
                //dt = dt.AddSeconds(ss[1].ToLong() / 115740);
                //val = dt.ToFullString();
            }
        }
        else if (st.NumFmtId is >= 18 and <= 21 or >= 45 and <= 47)
        {
            if (val is String str && Double.TryParse(str, out var d2))
            {
                val = TimeSpan.FromSeconds(Math.Round(d2 * 24 * 3600));
            }
        }
        // 自动处理0/General
        else if (st.NumFmtId == 0)
        {
            if (val is String str)
            {
                if (Int32.TryParse(str, out var n)) return n;
                if (Int64.TryParse(str, out var m)) return m;
                if (Decimal.TryParse(str, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return d;
                if (Double.TryParse(str, out var d2)) return d2;
            }
        }
        else if (st.NumFmtId is 1 or 3 or 37 or 38)
        {
            if (val is String str)
            {
                if (Int32.TryParse(str, out var n)) return n;
                if (Int64.TryParse(str, out var m)) return m;
            }
        }
        else if (st.NumFmtId is 2 or 4 or 11 or 39 or 40)
        {
            if (val is String str)
            {
                if (Decimal.TryParse(str, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return d;
                if (Double.TryParse(str, out var d2)) return d2;
            }
        }
        else if (st.NumFmtId is 9 or 10)
        {
            if (val is String str)
            {
                if (Double.TryParse(str, out var d2)) return d2;
            }
        }
        // 文本Text
        else if (st.NumFmtId == 49)
        {
            if (val is String str)
            {
                if (Decimal.TryParse(str, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return d.ToString();
                if (Double.TryParse(str, out var d2)) return d2.ToString();
            }
        }

        return val;
    }

    private String[]? ReadStrings(Stream ms)
    {
        var doc = XDocument.Load(ms);
        if (doc?.Root == null) return null;

        var list = new List<String>();
        foreach (var item in doc.Root.Elements())
        {
            list.Add(item.Value);
        }

        return list.ToArray();
    }

    private ExcelNumberFormat?[]? ReadStyles(Stream ms)
    {
        var doc = XDocument.Load(ms);
        if (doc?.Root == null) return null;

        // 内置默认样式
        var fmts = new Dictionary<Int32, String>
        {
            [0] = "General",
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [14] = "mm-dd-yy",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "m/d/yy h:mm",
            [37] = "#,##0 ;(#,##0)",
            [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)",
            [40] = "#,##0.00;[Red](#,##0.00)",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [48] = "##0.0E+0",
            [49] = "@"
        };

        // 自定义样式
        var numFmts = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "numFmts");
        if (numFmts != null)
        {
            foreach (var item in numFmts.Elements())
            {
                var id = item.Attribute("numFmtId");
                var code = item.Attribute("formatCode");
                if (id != null && code != null) fmts[id.Value.ToInt()] = code.Value;
            }
        }

        var list = new List<ExcelNumberFormat?>();
        var xfs = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "cellXfs");
        if (xfs != null)
        {
            foreach (var item in xfs.Elements())
            {
                var fid = item.Attribute("numFmtId");
                if (fid == null) continue;

                var id = fid.Value.ToInt();
                if (fmts.TryGetValue(id, out var code))
                    list.Add(new ExcelNumberFormat(id, code));
                else
                    list.Add(null);
            }
        }

        return list.ToArray();
    }

    /// <summary>解析完整样式（字体/填充/边框/XF）</summary>
    private void ParseFullStyles(Stream ms)
    {
        var doc = XDocument.Load(ms);
        if (doc?.Root == null) return;

        var ns = doc.Root.Name.Namespace;

        // numFmts（自定义数字格式）
        _numFmtCodes = new Dictionary<Int32, String>
        {
            [0] = "General", [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00",
            [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
            [14] = "mm-dd-yy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy",
            [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss",
            [22] = "m/d/yy h:mm",
            [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)",
            [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@"
        };
        var numFmts = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "numFmts");
        if (numFmts != null)
        {
            foreach (var item in numFmts.Elements())
            {
                var id = item.Attribute("numFmtId");
                var code = item.Attribute("formatCode");
                if (id != null && code != null) _numFmtCodes[id.Value.ToInt()] = code.Value;
            }
        }

        // fonts
        _fontInfos = [];
        var fontsEl = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "fonts");
        if (fontsEl != null)
        {
            foreach (var f in fontsEl.Elements())
            {
                var fi = new FontInfo();
                fi.Bold = f.Element(ns + "b") != null;
                fi.Italic = f.Element(ns + "i") != null;
                fi.Underline = f.Element(ns + "u") != null;
                var sz = f.Element(ns + "sz");
                if (sz != null) fi.Size = sz.Attribute("val")?.Value.ToDouble() ?? 0;
                var color = f.Element(ns + "color");
                if (color != null) fi.Color = NormalizeColorAttr(color);
                var name = f.Element(ns + "name");
                if (name != null) fi.Name = name.Attribute("val")?.Value;
                _fontInfos.Add(fi);
            }
        }

        // fills
        _fillInfos = [];
        var fillsEl = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "fills");
        if (fillsEl != null)
        {
            foreach (var f in fillsEl.Elements())
            {
                var fli = new FillInfo();
                var pf = f.Element(ns + "patternFill");
                if (pf != null)
                {
                    fli.PatternType = pf.Attribute("patternType")?.Value ?? "none";
                    var fg = pf.Element(ns + "fgColor");
                    if (fg != null) fli.BgColor = NormalizeColorAttr(fg);
                    if (fli.BgColor.IsNullOrEmpty())
                    {
                        var bg = pf.Element(ns + "bgColor");
                        if (bg != null) fli.BgColor = NormalizeColorAttr(bg);
                    }
                }
                _fillInfos.Add(fli);
            }
        }

        // borders
        _borderInfos = [];
        var bordersEl = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "borders");
        if (bordersEl != null)
        {
            foreach (var b in bordersEl.Elements())
            {
                var bi = new BorderInfo();
                var left = b.Element(ns + "left");
                if (left != null)
                {
                    bi.Style = ParseBorderStyle(left.Attribute("style")?.Value);
                    var lc = left.Element(ns + "color");
                    if (lc != null) bi.Color = NormalizeRgb(lc.Attribute("rgb")?.Value);
                }
                if (bi.Style == CellBorderStyle.None)
                {
                    var bottom = b.Element(ns + "bottom");
                    if (bottom != null)
                    {
                        bi.Style = ParseBorderStyle(bottom.Attribute("style")?.Value);
                        var bc = bottom.Element(ns + "color");
                        if (bc != null) bi.Color = NormalizeRgb(bc.Attribute("rgb")?.Value);
                    }
                }
                _borderInfos.Add(bi);
            }
        }

        // cellXfs
        _xfInfos = [];
        var xfsEl = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "cellXfs");
        if (xfsEl != null)
        {
            foreach (var xf in xfsEl.Elements())
            {
                var xi = new XfInfo
                {
                    NumFmtId = xf.Attribute("numFmtId")?.Value.ToInt() ?? 0,
                    FontId = xf.Attribute("fontId")?.Value.ToInt() ?? 0,
                    FillId = xf.Attribute("fillId")?.Value.ToInt() ?? 0,
                    BorderId = xf.Attribute("borderId")?.Value.ToInt() ?? 0,
                };
                var al = xf.Element(ns + "alignment");
                if (al != null)
                {
                    xi.HAlign = ParseHAlign(al.Attribute("horizontal")?.Value);
                    xi.VAlign = ParseVAlign(al.Attribute("vertical")?.Value);
                    var wt = al.Attribute("wrapText")?.Value;
                    xi.WrapText = wt == "1" || wt == "true";
                }
                _xfInfos.Add(xi);
            }
        }
    }

    private static CellBorderStyle ParseBorderStyle(String? style) => style switch
    {
        "thin" => CellBorderStyle.Thin,
        "medium" => CellBorderStyle.Medium,
        "thick" => CellBorderStyle.Thick,
        "dashed" => CellBorderStyle.Dashed,
        "dotted" => CellBorderStyle.Dotted,
        "double" => CellBorderStyle.DoubleLine,
        _ => CellBorderStyle.None,
    };

    /// <summary>规范化 OOXML RGB 颜色值：仅去除前导 FF alpha 前缀，保留实际 RGB</summary>
    private static String? NormalizeRgb(String? rgb)
    {
        if (rgb.IsNullOrEmpty() || rgb!.Length < 6) return rgb;
        if (rgb.Length >= 8 && rgb.StartsWith("FF"))
            return rgb[2..];
        return rgb;
    }

    /// <summary>规范化颜色 XML 元素：优先读 rgb 属性，无则读 theme 属性存为 "theme:N" 格式</summary>
    private static String? NormalizeColorAttr(XElement? colorEl)
    {
        if (colorEl == null) return null;
        var rgb = NormalizeRgb(colorEl.Attribute("rgb")?.Value);
        if (!rgb.IsNullOrEmpty()) return rgb;
        var theme = colorEl.Attribute("theme")?.Value;
        if (!theme.IsNullOrEmpty()) return $"theme:{theme}";
        return null;
    }

    private static HorizontalAlignment ParseHAlign(String? h) => h switch
    {
        "left" => HorizontalAlignment.Left,
        "center" => HorizontalAlignment.Center,
        "right" => HorizontalAlignment.Right,
        "fill" => HorizontalAlignment.Fill,
        "justify" => HorizontalAlignment.Justify,
        _ => HorizontalAlignment.General,
    };

    private static VerticalAlignment ParseVAlign(String? v) => v switch
    {
        "center" => VerticalAlignment.Center,
        "bottom" => VerticalAlignment.Bottom,
        _ => VerticalAlignment.Top,
    };

    private IDictionary<String, ZipArchiveEntry> ReadSheets(ZipArchive zip)
    {
        var dic = new Dictionary<String, String?>();

        var entry = zip.GetEntry("xl/workbook.xml");
        if (entry != null)
        {
            using var es = entry.Open(); // 释放 workbook.xml 流
            var doc = XDocument.Load(es);
            if (doc?.Root != null)
            {
                //var list = new List<String>();
                var sheets = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "sheets");
                if (sheets != null)
                {
                    foreach (var item in sheets.Elements())
                    {
                        var id = item.Attribute("sheetId");
                        var name = item.Attribute("name");
                        if (id != null) dic[id.Value] = name?.Value;
                    }

                    // 按 workbook.xml 中 <sheets> 的顺序保存工作表名称列表
                    _orderedSheets = [];
                    foreach (var item in sheets.Elements())
                    {
                        var name = item.Attribute("name")?.Value;
                        if (name != null) _orderedSheets.Add(name);
                    }
                }
            }
        }

        //_entries = _zip.Entries.Where(e =>
        //    e.FullName.StartsWithIgnoreCase("xl/worksheets/") &&
        //    e.Name.EndsWithIgnoreCase(".xml"))
        //    .ToDictionary(e => e.Name.TrimEnd(".xml"), e => e);

        var dic2 = new Dictionary<String, ZipArchiveEntry>();
        foreach (var item in zip.Entries)
        {
            if (item.FullName.StartsWithIgnoreCase("xl/worksheets/") && item.Name.EndsWithIgnoreCase(".xml"))
            {
                var name = item.Name.TrimEnd(".xml");
                if (dic.TryGetValue(name.TrimStart("sheet"), out var str)) name = str;
                name ??= String.Empty;

                dic2[name] = item;
            }
        }

        return dic2;
    }
    #endregion

    #region 辅助
    /// <summary>解析单元格引用（如 A1 / AB23）得到列索引（0基）。失败返回 0。</summary>
    private static Int32 GetColumnIndex(String cellRef)
    {
        // 提取前导字母部分
        var len = 0;
        for (var i = 0; i < cellRef.Length; i++)
        {
            var ch = cellRef[i];
            if (ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z') len++;
            else break;
        }
        if (len == 0) return 0;

        var index = 0;
        for (var i = 0; i < len; i++)
        {
            var ch = cellRef[i];
            if (ch is >= 'a' and <= 'z') ch = (Char)(ch - 'a' + 'A');
            index = index * 26 + (ch - 'A' + 1);
        }
        return index - 1; // 转为0基
    }
    #endregion

    #region 对象映射
    /// <summary>将工作表数据映射到强类型对象集合</summary>
    /// <typeparam name="T">目标类型（需有无参构造函数）</typeparam>
    /// <param name="sheet">工作表名称（可空，空时取第一个）</param>
    /// <returns>对象枚举（第一行作为表头映射列名）</returns>
    public IEnumerable<T> ReadObjects<T>(String? sheet = null) where T : new()
    {
        ThrowIfDisposed();

        using var enumerator = ReadRows(sheet).GetEnumerator();
        if (!enumerator.MoveNext()) yield break;

        // 第一行作为表头
        var headers = enumerator.Current.Select(e => e?.ToString() ?? "").ToArray();

        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(e => e.CanWrite)
            .ToArray();

        // 建立列索引 → 属性映射
        var mapping = new PropertyInfo?[headers.Length];
        for (var i = 0; i < headers.Length; i++)
        {
            var h = headers[i];
            if (h.IsNullOrEmpty()) continue;

            foreach (var p in props)
            {
                // 按属性名匹配
                if (p.Name.EqualIgnoreCase(h)) { mapping[i] = p; break; }
                // 按 DisplayName 匹配
                var dn = p.GetCustomAttribute<DisplayNameAttribute>();
                if (dn != null && dn.DisplayName == h) { mapping[i] = p; break; }
                // 按 Description 匹配
                var desc = p.GetCustomAttribute<DescriptionAttribute>();
                if (desc != null && desc.Description == h) { mapping[i] = p; break; }
            }
        }

        // 数据行
        while (enumerator.MoveNext())
        {
            var row = enumerator.Current;
            var item = new T();
            for (var c = 0; c < Math.Min(row.Length, mapping.Length); c++)
            {
                var prop = mapping[c];
                if (prop == null || row[c] == null) continue;

                try
                {
                    var val = row[c];
                    var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

                    if (val is String s)
                    {
                        // 字符串到目标类型转换
                        if (targetType == typeof(String))
                            prop.SetValue(item, s);
                        else if (targetType == typeof(Int32))
                            prop.SetValue(item, s.ToInt());
                        else if (targetType == typeof(Int64))
                            prop.SetValue(item, Int64.TryParse(s, out var v64) ? v64 : 0L);
                        else if (targetType == typeof(Double))
                            prop.SetValue(item, s.ToDouble());
                        else if (targetType == typeof(Decimal))
                            prop.SetValue(item, Decimal.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var vd) ? vd : 0m);
                        else if (targetType == typeof(Boolean))
                            prop.SetValue(item, s.ToBoolean());
                        else if (targetType == typeof(DateTime))
                            prop.SetValue(item, s.ToDateTime());
                        else
                            prop.SetValue(item, Convert.ChangeType(val, targetType, CultureInfo.InvariantCulture));
                    }
                    else if (val != null)
                    {
                        // 其他类型直接或转换后赋值
                        var valType = val.GetType();
                        if (targetType.IsAssignableFrom(valType))
                            prop.SetValue(item, val);
                        else
                            prop.SetValue(item, Convert.ChangeType(val, targetType, CultureInfo.InvariantCulture));
                    }
                }
                catch
                {
                    // 转换失败跳过该字段
                }
            }
            yield return item;
        }
    }

    /// <summary>将工作表数据读取为 DataTable</summary>
    /// <param name="sheet">工作表名称（可空，空时取第一个）</param>
    /// <returns>DataTable（第一行作为列名）</returns>
    public DataTable ReadDataTable(String? sheet = null)
    {
        ThrowIfDisposed();

        var dt = new DataTable();
        var isFirst = true;

        foreach (var row in ReadRows(sheet))
        {
            if (isFirst)
            {
                // 第一行作为列名
                for (var i = 0; i < row.Length; i++)
                {
                    var colName = row[i]?.ToString() ?? $"Column{i + 1}";
                    dt.Columns.Add(colName);
                }
                isFirst = false;
                continue;
            }

            var dr = dt.NewRow();
            for (var i = 0; i < Math.Min(row.Length, dt.Columns.Count); i++)
            {
                dr[i] = row[i] ?? DBNull.Value;
            }
            dt.Rows.Add(dr);
        }

        return dt;
    }

    /// <summary>获取合并单元格区域列表</summary>
    /// <param name="sheet">工作表名称（可空，空时取第一个）</param>
    /// <returns>合并区域列表，每项为 (起始行0基, 起始列0基, 结束行0基, 结束列0基)</returns>
    public IList<(Int32 StartRow, Int32 StartCol, Int32 EndRow, Int32 EndCol)>? GetMergeRanges(String? sheet = null)
    {
        ThrowIfDisposed();

        if (Sheets == null || _entries == null) return null;

        if (sheet.IsNullOrEmpty()) sheet = Sheets.FirstOrDefault();
        if (sheet.IsNullOrEmpty()) return null;

        if (!_entries.TryGetValue(sheet, out var entry)) return null;

        using var esheet = entry.Open();
        var doc = XDocument.Load(esheet);
        if (doc.Root == null) return null;

        var mergeNode = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("mergeCells"));
        if (mergeNode == null) return null;

        var result = new List<(Int32, Int32, Int32, Int32)>();
        foreach (var mc in mergeNode.Elements())
        {
            var refAttr = mc.Attribute("ref")?.Value;
            if (refAttr.IsNullOrEmpty()) continue;

            var parts = refAttr!.Split(':');
            if (parts.Length != 2) continue;

            var (r1, c1) = ParseCellRef(parts[0]);
            var (r2, c2) = ParseCellRef(parts[1]);
            result.Add((r1, c1, r2, c2));
        }

        return result;
    }

    /// <summary>读取工作表超链接</summary>
    /// <param name="sheet">工作表名称（可空，空时取第一个）</param>
    /// <returns>单元格引用到 URL 的字典（如 "A1" → "https://..."）</returns>
    public IDictionary<String, String> ReadHyperlinks(String? sheet = null)
    {
        ThrowIfDisposed();
        var result = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);

        if (Sheets == null || _entries == null) return result;
        if (sheet.IsNullOrEmpty()) sheet = Sheets.FirstOrDefault();
        if (sheet.IsNullOrEmpty()) return result;
        if (!_entries.TryGetValue(sheet, out var entry)) return result;

        // 读取 .rels 文件（r:id → URL）
        var relsPath = "xl/worksheets/_rels/" + entry.Name + ".rels";
        var relsEntry = _zip.GetEntry(relsPath);
        var urlMap = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        if (relsEntry != null)
        {
            using var rs = relsEntry.Open();
            var relsDoc = XDocument.Load(rs);
            if (relsDoc.Root != null)
            {
                foreach (var rel in relsDoc.Root.Elements())
                {
                    var type = rel.Attribute("Type")?.Value ?? String.Empty;
                    if (!type.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase)) continue;
                    var id = rel.Attribute("Id")?.Value;
                    var target = rel.Attribute("Target")?.Value;
                    if (id != null && target != null) urlMap[id] = target;
                }
            }
        }

        // 读取 sheet.xml 中的 <hyperlinks> 节点
        using var esheet = entry.Open();
        var doc = XDocument.Load(esheet);
        if (doc.Root == null) return result;

        var hyperlinks = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "hyperlinks");
        if (hyperlinks == null) return result;

        foreach (var hl in hyperlinks.Elements())
        {
            var cellRef = hl.Attribute("ref")?.Value;
            if (cellRef.IsNullOrEmpty()) continue;

            // 外部链接：通过 r:id 查找 URL
            var rId = hl.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
            if (rId != null && urlMap.TryGetValue(rId, out var url))
            {
                result[cellRef!] = url;
            }
            else
            {
                // 内部位置超链接（#SheetName!A1 格式）
                var loc = hl.Attribute("location")?.Value;
                if (!loc.IsNullOrEmpty()) result[cellRef!] = "#" + loc;
            }
        }

        return result;
    }

    /// <summary>解析单元格引用返回 (行0基, 列0基)</summary>
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
        colIndex--;

        var rowStr = cellRef[colLen..];
        var rowIndex = Int32.Parse(rowStr) - 1;

        return (rowIndex, colIndex);
    }

    /// <summary>打开指定工作表的 XML 文档</summary>
    private XDocument OpenSheetXml(String sheet)
    {
        if (_entries == null || !_entries.TryGetValue(sheet, out var entry))
            throw new ArgumentOutOfRangeException(nameof(sheet), "Unable to find worksheet");
        using var es = entry.Open();
        return XDocument.Load(es);
    }
    #endregion

    #region 完整读取
    /// <summary>读取工作簿完整快照（数据+样式+元数据）</summary>
    /// <returns>ExcelData 完整快照</returns>
    public ExcelData ReadExcel()
    {
        ThrowIfDisposed();

        var data = new ExcelData();
        if (Sheets == null) return data;

        // 提取默认字体（font[0]）供 Writer 重建 styles.xml
        if (_fontInfos != null && _fontInfos.Count > 0)
        {
            var f0 = _fontInfos[0];
            data.DefaultFont = new DefaultFontInfo
            {
                Name = f0.Name,
                Size = f0.Size,
                Bold = f0.Bold,
                Color = f0.Color,
            };
        }

        foreach (var sheet in Sheets)
        {
            data.Sheets.Add(ReadSheet(sheet));
        }

        // 收集所有未解析的 ZIP 部件，确保往返不丢内容
        CollectOtherParts(data);

        return data;
    }

    /// <summary>读取单个工作表的完整快照</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>SheetData 快照</returns>
    public SheetData ReadSheet(String sheet)
    {
        var sd = new SheetData { Name = sheet };

        // 行数据（带实际行号，用于还原跳行结构）
        var rowNumbers = new List<Int32>();
        var rows = new List<Object?[]>();
        foreach (var (rowIdx, rowData) in ReadRowsWithIndex(sheet))
        {
            rowNumbers.Add(rowIdx);
            rows.Add(rowData);
        }
        sd.Rows = rows;
        // 仅当存在跳行时才记录行号映射（优化：连续行不需要）
        var hasGaps = false;
        for (var i = 0; i < rowNumbers.Count; i++)
        {
            if (rowNumbers[i] != i) { hasGaps = true; break; }
        }
        if (hasGaps) sd.ActualRowNumbers = rowNumbers;

        // 单元格样式
        sd.CellStyles = ReadCellStyles(sheet);

        // 合并区域
        var merges = GetMergeRanges(sheet);
        if (merges != null) sd.Merges = merges.ToList();

        // 冻结窗格
        sd.FreezePane = ReadFreezePanes(sheet);

        // 自动筛选
        sd.AutoFilter = ReadAutoFilter(sheet);

        // 行高
        sd.RowHeights = ReadRowHeights(sheet);

        // 列宽
        sd.ColumnWidths = ReadColumnWidths(sheet);

        // 超链接
        var links = ReadHyperlinks(sheet);
        foreach (var kv in links)
        {
            var (r, c) = ParseCellRef(kv.Key);
            sd.Hyperlinks[(r, c)] = (kv.Value, null);
        }

        // 图片
        sd.Images = ReadImages(sheet).ToList();

        // 页面设置
        ReadPageSetup(sheet, sd);

        // 工作表保护
        sd.ProtectionPassword = ReadSheetProtection(sheet);

        // 条件格式
        sd.ConditionalFormats = ReadConditionalFormats(sheet).ToList();

        // 批注
        sd.Comments = ReadComments(sheet);

        // 数据验证
        sd.Validations = ReadDataValidations(sheet).ToList();

        // 公式
        sd.Formulas = ReadFormulas(sheet);

        return sd;
    }

    /// <summary>读取指定工作表每单元格的完整样式</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>(行, 列) → CellStyle 字典，行列均为0基</returns>
    public Dictionary<(Int32 Row, Int32 Col), CellStyle> ReadCellStyles(String sheet)
    {
        var result = new Dictionary<(Int32, Int32), CellStyle>();
        if (_xfInfos == null) return result;

        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return result;

        var data = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("sheetData"));
        if (data == null) return result;

        foreach (var row in data.Elements())
        {
            var rAttr = row.Attribute("r");
            if (rAttr == null) continue;
            var rowIndex = rAttr.Value.ToInt(-1) - 1;
            foreach (var col in row.Elements())
            {
                var r = col.Attribute("r")?.Value;
                if (r.IsNullOrEmpty()) continue;
                var (_, colIndex) = ParseCellRef(r!);

                var sAttr = col.Attribute("s");
                if (sAttr == null) continue;

                var sIdx = sAttr.Value.ToInt(-1);
                if (sIdx < 0 || sIdx >= _xfInfos.Count) continue;

                var xf = _xfInfos[sIdx];
                var cs = new CellStyle();

                // 字体
                if (_fontInfos != null && xf.FontId >= 0 && xf.FontId < _fontInfos.Count)
                {
                    var fi = _fontInfos[xf.FontId];
                    cs.FontName = fi.Name;
                    cs.FontSize = fi.Size;
                    cs.Bold = fi.Bold;
                    cs.Italic = fi.Italic;
                    cs.Underline = fi.Underline;
                    cs.FontColor = fi.Color;
                }

                // 填充
                if (_fillInfos != null && xf.FillId >= 0 && xf.FillId < _fillInfos.Count)
                {
                    var fli = _fillInfos[xf.FillId];
                    if (fli.PatternType == "solid" && !fli.BgColor.IsNullOrEmpty())
                        cs.BackgroundColor = fli.BgColor;
                }

                // 边框
                if (_borderInfos != null && xf.BorderId >= 0 && xf.BorderId < _borderInfos.Count)
                {
                    var bi = _borderInfos[xf.BorderId];
                    cs.Border = bi.Style;
                    cs.BorderColor = bi.Color;
                }

                // 对齐
                cs.HAlign = xf.HAlign;
                cs.VAlign = xf.VAlign;
                cs.WrapText = xf.WrapText;

                // 数字格式
                if (_numFmtCodes != null && _numFmtCodes.TryGetValue(xf.NumFmtId, out var fmt) && fmt != "General")
                    cs.NumberFormat = fmt;

                result[(rowIndex, colIndex)] = cs;
            }
        }

        return result;
    }

    /// <summary>读取指定工作表的列宽</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>0基列号 → 字符宽度</returns>
    public Dictionary<Int32, Double> ReadColumnWidths(String sheet)
    {
        var result = new Dictionary<Int32, Double>();
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return result;

        var cols = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("cols"));
        if (cols == null) return result;

        foreach (var col in cols.Elements())
        {
            var min = col.Attribute("min")?.Value.ToInt(1) ?? 1;
            var max = col.Attribute("max")?.Value.ToInt(1) ?? 1;
            var width = col.Attribute("width")?.Value.ToDouble() ?? 0;
            for (var c = min - 1; c < max; c++)
                result[c] = width;
        }
        return result;
    }

    /// <summary>读取指定工作表的行高</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>0基行号 → 磅值</returns>
    public Dictionary<Int32, Double> ReadRowHeights(String sheet)
    {
        var result = new Dictionary<Int32, Double>();
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return result;

        var data = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("sheetData"));
        if (data == null) return result;

        foreach (var row in data.Elements())
        {
            var ht = row.Attribute("ht");
            if (ht == null) continue;
            var rAttr = row.Attribute("r");
            var r = rAttr != null ? rAttr.Value.ToInt(-1) : -1;
            if (r < 1) continue;
            result[r - 1] = ht.Value.ToDouble();
        }
        return result;
    }

    /// <summary>读取指定工作表的冻结窗格</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>冻结(行数, 列数)，null 表示未冻结</returns>
    public (Int32 Rows, Int32 Cols)? ReadFreezePanes(String sheet)
    {
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return null;

        var sheetViews = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("sheetViews"));
        if (sheetViews == null) return null;

        var pane = sheetViews.Elements().Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("pane"));
        if (pane == null) return null;

        var state = pane.Attribute("state")?.Value;
        if (state != "frozen") return null;

        var ySplit = pane.Attribute("ySplit")?.Value.ToInt() ?? 0;
        var xSplit = pane.Attribute("xSplit")?.Value.ToInt() ?? 0;
        return (ySplit, xSplit);
    }

    /// <summary>读取指定工作表的自动筛选范围</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>筛选范围（如 "A1:F1"），null 表示未设置</returns>
    public String? ReadAutoFilter(String sheet)
    {
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return null;

        var af = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("autoFilter"));
        return af?.Attribute("ref")?.Value;
    }

    /// <summary>读取指定工作表的图片</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>图片列表</returns>
    public IEnumerable<ExcelImage> ReadImages(String sheet)
    {
        if (_entries == null || !_entries.TryGetValue(sheet, out var sheetEntry)) yield break;

        // 通过 sheet rels 找到 drawing
        var sheetIdx = sheetEntry.Name.TrimEnd(".xml").TrimStart("sheet").ToInt(-1);
        if (sheetIdx < 1) yield break;
        var relsPath = $"xl/worksheets/_rels/sheet{sheetIdx}.xml.rels";
        var relsEntry = _zip.GetEntry(relsPath);
        if (relsEntry == null) yield break;

        String? drawingRId = null;
        using (var rs = relsEntry.Open())
        {
            var relsDoc = XDocument.Load(rs);
            if (relsDoc.Root != null)
            {
                foreach (var rel in relsDoc.Root.Elements())
                {
                    var type = rel.Attribute("Type")?.Value ?? String.Empty;
                    if (type.Contains("drawing", StringComparison.OrdinalIgnoreCase))
                    {
                        drawingRId = rel.Attribute("Id")?.Value;
                        break;
                    }
                }
            }
        }
        if (drawingRId == null) yield break;

        // 读 drawing XML 获取图片位置和 rId
        var drawingPath = $"xl/drawings/drawing{sheetIdx}.xml";
        var drawingEntry = _zip.GetEntry(drawingPath);
        if (drawingEntry == null) yield break;

        var images = new List<(Int32 Row, Int32 Col, Int64 FromColOff, Int64 FromRowOff, Int32 ToRow, Int32 ToCol, Int64 ToColOff, Int64 ToRowOff, String EditAs, String ImgRId, Int64 EmuW, Int64 EmuH)>();
        using (var ds = drawingEntry.Open())
        {
            var drawDoc = XDocument.Load(ds);
            if (drawDoc.Root == null) yield break;

            var xdrNs = drawDoc.Root.Name.Namespace;
            var aNs = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            foreach (var anchor in drawDoc.Root.Elements())
            {
                var from = anchor.Element(xdrNs + "from");
                if (from == null) continue;
                var colEl = from.Element(xdrNs + "col");
                var rowEl = from.Element(xdrNs + "row");
                if (colEl == null || rowEl == null) continue;

                var col = colEl.Value.ToInt();
                var row = rowEl.Value.ToInt();
                var fromColOff = Int64.TryParse(from.Element(xdrNs + "colOff")?.Value, out var fco) ? fco : 0L;
                var fromRowOff = Int64.TryParse(from.Element(xdrNs + "rowOff")?.Value, out var fro) ? fro : 0L;

                var toRow = -1; var toCol = -1; var toColOff = 0L; var toRowOff = 0L;
                var toEl = anchor.Element(xdrNs + "to");
                if (toEl != null)
                {
                    toCol = toEl.Element(xdrNs + "col")?.Value.ToInt() ?? -1;
                    toRow = toEl.Element(xdrNs + "row")?.Value.ToInt() ?? -1;
                    Int64.TryParse(toEl.Element(xdrNs + "colOff")?.Value, out toColOff);
                    Int64.TryParse(toEl.Element(xdrNs + "rowOff")?.Value, out toRowOff);
                }

                var editAs = anchor.Attribute("editAs")?.Value ?? "oneCell";

                var pic = anchor.Element(xdrNs + "pic");
                if (pic == null) continue;
                var blipFill = pic.Element(xdrNs + "blipFill");
                if (blipFill == null) continue;
                var blip = blipFill.Element(aNs + "blip");
                if (blip == null) continue;
                var embed = blip.Attributes().FirstOrDefault(a => a.Name.LocalName == "embed");
                if (embed == null) continue;

                var spPr = pic.Element(xdrNs + "spPr");
                Int64 emuW = 952500, emuH = 952500;
                if (spPr != null)
                {
                    var xfrm = spPr.Element(aNs + "xfrm");
                    if (xfrm != null)
                    {
                        var ext = xfrm.Element(aNs + "ext");
                        if (ext != null)
                        {
                            emuW = Int64.TryParse(ext.Attribute("cx")?.Value, out var cxVal) ? cxVal : 952500;
                            emuH = Int64.TryParse(ext.Attribute("cy")?.Value, out var cyVal) ? cyVal : 952500;
                        }
                    }
                }

                images.Add((row, col, fromColOff, fromRowOff, toRow, toCol, toColOff, toRowOff, editAs, embed.Value, emuW, emuH));
            }
        }

        // 读 drawing rels 获取图片路径
        var drawRelsPath = $"xl/drawings/_rels/drawing{sheetIdx}.xml.rels";
        var drawRelsEntry = _zip.GetEntry(drawRelsPath);
        var imgPathMap = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        if (drawRelsEntry != null)
        {
            using var drs = drawRelsEntry.Open();
            var drDoc = XDocument.Load(drs);
            if (drDoc.Root != null)
            {
                foreach (var rel in drDoc.Root.Elements())
                {
                    var id = rel.Attribute("Id")?.Value;
                    var target = rel.Attribute("Target")?.Value;
                    if (id != null && target != null)
                    {
                        // 规范化路径：消除 target 中的 ".."（如 "../media/image1.png" → "media/image1.png"）
                        var baseDir = "xl/drawings/";
                        var combined = baseDir + target;
                        var parts = combined.Split('/');
                        var normalized = new List<String>();
                        foreach (var p in parts)
                        {
                            if (p == "..")
                            {
                                if (normalized.Count > 0) normalized.RemoveAt(normalized.Count - 1);
                            }
                            else if (p != "." && !p.IsNullOrEmpty())
                            {
                                normalized.Add(p);
                            }
                        }
                        imgPathMap[id] = String.Join("/", normalized);
                    }
                }
            }
        }

        // 输出图片
        foreach (var (row, col, fromColOff, fromRowOff, toRow, toCol, toColOff, toRowOff, editAs, imgRId, emuW, emuH) in images)
        {
            if (!imgPathMap.TryGetValue(imgRId, out var mediaPath)) continue;
            var mediaEntry = _zip.GetEntry(mediaPath);
            if (mediaEntry == null) continue;

            Byte[] data;
            using (var ms2 = new MemoryStream())
            {
                using var mediaStream = mediaEntry.Open();
                mediaStream.CopyTo(ms2);
                data = ms2.ToArray();
            }

            var ext = Path.GetExtension(mediaPath).TrimStart('.').ToLower();
            if (ext.IsNullOrEmpty()) ext = "png";

            yield return new ExcelImage
            {
                Data = data,
                Extension = ext,
                Row = row,
                Col = col,
                Width = Math.Round(emuW / 9525.0, 1),
                Height = Math.Round(emuH / 9525.0, 1),
                FromColOff = fromColOff,
                FromRowOff = fromRowOff,
                ToRow = toRow,
                ToCol = toCol,
                ToColOff = toColOff,
                ToRowOff = toRowOff,
                EditAs = editAs,
            };
        }
    }

    /// <summary>读取页面设置信息并填入 SheetData</summary>
    /// <param name="sheet">工作表名称</param>
    /// <param name="sd">目标 SheetData</param>
    public void ReadPageSetup(String sheet, SheetData sd)
    {
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return;

        // pageMargins
        var margins = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("pageMargins"));
        if (margins != null)
        {
            sd.MarginTop = margins.Attribute("top")?.Value.ToDouble() ?? 0.75;
            sd.MarginBottom = margins.Attribute("bottom")?.Value.ToDouble() ?? 0.75;
            sd.MarginLeft = margins.Attribute("left")?.Value.ToDouble() ?? 0.7;
            sd.MarginRight = margins.Attribute("right")?.Value.ToDouble() ?? 0.7;
        }

        // pageSetup
        var ps = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("pageSetup"));
        if (ps != null)
        {
            var orient = ps.Attribute("orientation")?.Value;
            sd.Orientation = orient == "landscape" ? PageOrientation.Landscape : PageOrientation.Portrait;
            var psVal = ps.Attribute("paperSize")?.Value.ToInt() ?? 0;
            sd.PaperSize = psVal > 0 ? (PaperSize)psVal : PaperSize.Default;
        }

        // headerFooter
        var hf = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("headerFooter"));
        if (hf != null)
        {
            sd.HeaderText = hf.Elements().FirstOrDefault(e => e.Name.LocalName == "oddHeader")?.Value;
            sd.FooterText = hf.Elements().FirstOrDefault(e => e.Name.LocalName == "oddFooter")?.Value;
        }

        // print titles（需从 workbook.xml 读取 definedNames）
        try
        {
            var wbEntry = _zip.GetEntry("xl/workbook.xml");
            if (wbEntry != null)
            {
                using var ws = wbEntry.Open();
                var wbDoc = XDocument.Load(ws);
                if (wbDoc.Root != null)
                {
                    var definedNames = wbDoc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("definedNames"));
                    if (definedNames != null)
                    {
                        foreach (var dn in definedNames.Elements())
                        {
                            if (dn.Attribute("name")?.Value != "_xlnm.Print_Titles") continue;
                            var text = dn.Value;
                            // 格式: 'SheetName'!$1:$1
                            var parts = text.Split('!');
                            if (parts.Length == 2)
                            {
                                var sn = parts[0].Trim('\'');
                                if (!sn.EqualIgnoreCase(sheet)) continue;
                                var range = parts[1].Trim('$').Split(':');
                                if (range.Length == 2)
                                {
                                    sd.PrintTitleStartRow = range[0].ToInt();
                                    sd.PrintTitleEndRow = range[1].ToInt();
                                }
                            }
                        }
                    }
                }
            }
        }
        catch { /* 非关键 */ }
    }

    /// <summary>读取工作表保护密码哈希</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>密码哈希，null 表示未保护</returns>
    public String? ReadSheetProtection(String sheet)
    {
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return null;

        var sp = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("sheetProtection"));
        return sp?.Attribute("password")?.Value;
    }

    /// <summary>读取条件格式</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>条件格式列表</returns>
    public IEnumerable<ConditionalFormatInfo> ReadConditionalFormats(String sheet)
    {
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) yield break;

        foreach (var cf in doc.Root.Elements().Where(e => e.Name.LocalName.EqualIgnoreCase("conditionalFormatting")))
        {
            var range = cf.Attribute("sqref")?.Value;
            if (range.IsNullOrEmpty()) continue;

            foreach (var rule in cf.Elements())
            {
                var type = rule.Attribute("type")?.Value;
                if (type.IsNullOrEmpty()) continue;

                var info = new ConditionalFormatInfo { Range = range! };

                if (type == "cellIs")
                {
                    var op = rule.Attribute("operator")?.Value;
                    info.Type = op switch
                    {
                        "greaterThan" => ConditionalFormatType.GreaterThan,
                        "lessThan" => ConditionalFormatType.LessThan,
                        "equal" => ConditionalFormatType.Equal,
                        "between" => ConditionalFormatType.Between,
                        _ => ConditionalFormatType.GreaterThan,
                    };
                    var formulas = rule.Elements().Where(e => e.Name.LocalName == "formula").ToList();
                    if (formulas.Count > 0) info.Value = formulas[0].Value;
                    if (formulas.Count > 1) info.Value2 = formulas[1].Value;

                    // 尝试从 dxf 获取颜色（简化：从 styles.xml 的 dxfs 读取）
                    var dxfId = rule.Attribute("dxfId")?.Value.ToInt(-1) ?? -1;
                    if (dxfId >= 0 && _fillInfos != null)
                    {
                        // dxf 颜色解析较复杂，暂不处理
                    }
                }
                else if (type == "dataBar")
                {
                    info.Type = ConditionalFormatType.DataBar;
                    var dataBar = rule.Element(rule.Name.Namespace + "dataBar");
                    var color = dataBar?.Element(rule.Name.Namespace + "color");
                    info.Color = NormalizeRgb(color?.Attribute("rgb")?.Value);
                }
                else if (type == "colorScale")
                {
                    info.Type = ConditionalFormatType.ColorScale;
                }

                yield return info;
            }
        }
    }

    /// <summary>读取批注</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>(行, 列) → (文本, 作者)，行列均为0基</returns>
    public Dictionary<(Int32 Row, Int32 Col), (String Text, String Author)> ReadComments(String sheet)
    {
        var result = new Dictionary<(Int32, Int32), (String, String)>();
        if (_entries == null || !_entries.TryGetValue(sheet, out var sheetEntry)) return result;

        var sheetIdx = sheetEntry.Name.TrimEnd(".xml").TrimStart("sheet").ToInt(-1);
        if (sheetIdx < 1) return result;

        var commentsPath = $"xl/comments{sheetIdx}.xml";
        var commentsEntry = _zip.GetEntry(commentsPath);
        if (commentsEntry == null) return result;

        // 读作者列表
        var authors = new List<String>();
        using (var cs = commentsEntry.Open())
        {
            var doc = XDocument.Load(cs);
            if (doc.Root == null) return result;

            var authorsEl = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "authors");
            if (authorsEl != null)
            {
                foreach (var a in authorsEl.Elements())
                    authors.Add(a.Value);
            }

            var commentList = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "commentList");
            if (commentList == null) return result;

            foreach (var cmt in commentList.Elements())
            {
                var refAttr = cmt.Attribute("ref")?.Value;
                if (refAttr.IsNullOrEmpty()) continue;

                var (row, col) = ParseCellRef(refAttr!);
                var authorId = cmt.Attribute("authorId")?.Value.ToInt() ?? 0;
                var author = authorId >= 0 && authorId < authors.Count ? authors[authorId] : String.Empty;

                var text = String.Empty;
                var textEl = cmt.Elements().FirstOrDefault(e => e.Name.LocalName == "text");
                if (textEl != null)
                {
                    var rEl = textEl.Elements().FirstOrDefault(e => e.Name.LocalName == "r");
                    if (rEl != null)
                    {
                        var tEl = rEl.Elements().FirstOrDefault(e => e.Name.LocalName == "t");
                        if (tEl != null) text = tEl.Value;
                    }
                }

                result[(row, col)] = (text, author);
            }
        }

        return result;
    }

    /// <summary>读取数据验证</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>数据验证列表</returns>
    public IEnumerable<ValidationInfo> ReadDataValidations(String sheet)
    {
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) yield break;

        var dvs = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("dataValidations"));
        if (dvs == null) yield break;

        foreach (var dv in dvs.Elements())
        {
            var info = new ValidationInfo
            {
                CellRange = dv.Attribute("sqref")?.Value ?? String.Empty,
                ValidationType = dv.Attribute("type")?.Value,
                Operator = dv.Attribute("operator")?.Value,
            };

            var formulas = dv.Elements().Where(e => e.Name.LocalName == "formula1" || e.Name.LocalName == "formula2").ToList();
            foreach (var f in formulas)
            {
                if (f.Name.LocalName == "formula1") info.Formula1 = f.Value.Trim('"');
                else info.Formula2 = f.Value.Trim('"');
            }

            // 下拉列表
            if (info.ValidationType == "list" && !info.Formula1.IsNullOrEmpty())
                info.Items = info.Formula1!.Split(',').Select(e => e.Trim('"')).ToArray();

            yield return info;
        }
    }

    /// <summary>读取公式</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>(行, 列) → 公式文本（不含等号），行列均为0基</returns>
    public Dictionary<(Int32 Row, Int32 Col), String> ReadFormulas(String sheet)
    {
        var result = new Dictionary<(Int32, Int32), String>();
        var doc = OpenSheetXml(sheet);
        if (doc.Root == null) return result;

        var data = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName.EqualIgnoreCase("sheetData"));
        if (data == null) return result;

        // 共享公式字典：si → (基础公式文本, 基础行, 基础列)
        var sharedFormulas = new Dictionary<Int32, (String Formula, Int32 BaseRow, Int32 BaseCol)>();

        foreach (var row in data.Elements())
        {
            var rAttr2 = row.Attribute("r");
            if (rAttr2 == null) continue;
            var rowIndex = rAttr2.Value.ToInt(-1) - 1;
            if (rowIndex < 0) continue;

            foreach (var col in row.Elements())
            {
                var r = col.Attribute("r")?.Value;
                if (r.IsNullOrEmpty()) continue;
                var (_, colIndex) = ParseCellRef(r!);

                var fEl = col.Elements().FirstOrDefault(e => e.Name.LocalName == "f");
                if (fEl == null) continue;

                var fType = fEl.Attribute("t")?.Value;
                var val = fEl.Value;

                if (fType == "shared")
                {
                    var siStr = fEl.Attribute("si")?.Value;
                    if (siStr.IsNullOrEmpty()) continue;
                    var si = siStr.ToInt(-1);
                    if (si < 0) continue;

                    if (!val.IsNullOrEmpty())
                    {
                        // 定义行：包含公式文本
                        sharedFormulas[si] = (val, rowIndex, colIndex);
                        result[(rowIndex, colIndex)] = val;
                    }
                    else if (sharedFormulas.TryGetValue(si, out var baseEntry))
                    {
                        // 引用行：根据行偏移调整公式
                        var adjusted = AdjustFormula(baseEntry.Formula, rowIndex - baseEntry.BaseRow, colIndex - baseEntry.BaseCol);
                        if (!adjusted.IsNullOrEmpty())
                            result[(rowIndex, colIndex)] = adjusted!;
                    }
                }
                else
                {
                    if (val.IsNullOrEmpty()) continue;
                    result[(rowIndex, colIndex)] = val;
                }
            }
        }

        return result;
    }

    /// <summary>按行/列偏移调整公式中的 A1 风格单元格引用</summary>
    /// <param name="formula">原始公式文本（不含等号）</param>
    /// <param name="rowDelta">行偏移（正数表示向下）</param>
    /// <param name="colDelta">列偏移（正数表示向右）</param>
    private static String? AdjustFormula(String formula, Int32 rowDelta, Int32 colDelta)
    {
        if (formula.IsNullOrEmpty()) return formula;
        if (rowDelta == 0 && colDelta == 0) return formula;

        // 匹配 A1 风格的单元格引用：可选 $ + 列字母 + 可选 $ + 行数字
        // 使用简单的字符扫描替换，避免引入 Regex
        var sb = new System.Text.StringBuilder(formula.Length + 8);
        var i = 0;
        while (i < formula.Length)
        {
            // 寻找单元格引用起始：$字母 或 字母（不在单词中间）
            var ch = formula[i];
            var absCol = false;
            var absRow = false;

            // 检查是否是引用的起始
            var isStart = ch == '$' || (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
            // 排除函数名（后跟'('）和紧跟字母的字母
            if (isStart && i > 0)
            {
                var prev = formula[i - 1];
                if ((prev >= 'A' && prev <= 'Z') || (prev >= 'a' && prev <= 'z') || (prev >= '0' && prev <= '9') || prev == '_')
                    isStart = false;
            }

            if (!isStart)
            {
                sb.Append(ch);
                i++;
                continue;
            }

            var start = i;
            // 解析 $?列字母
            if (i < formula.Length && formula[i] == '$') { absCol = true; i++; }
            var colStart = i;
            while (i < formula.Length && ((formula[i] >= 'A' && formula[i] <= 'Z') || (formula[i] >= 'a' && formula[i] <= 'z'))) i++;
            var colStr = formula[colStart..i];

            if (colStr.IsNullOrEmpty() || i >= formula.Length)
            {
                // 不是引用，还原
                sb.Append(formula[start..i]);
                continue;
            }

            // 解析 $?行数字
            if (i < formula.Length && formula[i] == '$') { absRow = true; i++; }
            var rowStart = i;
            while (i < formula.Length && formula[i] >= '0' && formula[i] <= '9') i++;
            var rowStr = formula[rowStart..i];

            if (rowStr.IsNullOrEmpty())
            {
                // 只有列字母无行数字（可能是列范围引用），不是单元格引用
                sb.Append(formula[start..i]);
                continue;
            }

            // 确认下一个字符不是字母（排除命名范围）
            if (i < formula.Length && ((formula[i] >= 'A' && formula[i] <= 'Z') || (formula[i] >= 'a' && formula[i] <= 'z') || formula[i] == '_'))
            {
                sb.Append(formula[start..i]);
                continue;
            }

            // 计算新的行列
            var colNum = 0;
            foreach (var c in colStr.ToUpperInvariant()) colNum = colNum * 26 + (c - 'A' + 1);
            var rowNum = rowStr.ToInt();

            if (!absCol) colNum += colDelta;
            if (!absRow) rowNum += rowDelta;

            // 确保引用有效（行列 >= 1）
            if (colNum < 1 || rowNum < 1)
            {
                sb.Append(formula[start..i]);
                continue;
            }

            // 转回列字母
            var newColStr = new System.Text.StringBuilder();
            var cn = colNum;
            while (cn > 0) { newColStr.Insert(0, (Char)('A' + (cn - 1) % 26)); cn = (cn - 1) / 26; }

            if (absCol) sb.Append('$');
            sb.Append(newColStr);
            if (absRow) sb.Append('$');
            sb.Append(rowNum);
        }

        return sb.ToString();
    }

    /// <summary>收集所有未显式解析的 ZIP 部件，确保往返不丢内容</summary>
    private void CollectOtherParts(ExcelData data)
    {
        // 由 Writer 重新生成的部件，不从原始 ZIP 收集
        var handled = new HashSet<String>(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        };
        // 工作表 XML 由 Writer 根据 SheetData 重建
        foreach (var sheet in data.Sheets)
        {
            var idx = data.Sheets.IndexOf(sheet) + 1;
            handled.Add($"xl/worksheets/sheet{idx}.xml");
        }

        foreach (var entry in _zip.Entries)
        {
            var name = entry.FullName;
            if (name.EndsWith("/")) continue;
            if (handled.Contains(name)) continue;

            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            data.OtherParts[name] = ms.ToArray();
        }
    }
    #endregion

    #region 内嵌类
    class ExcelNumberFormat(Int32 numFmtId, String format)
    {
        public Int32 NumFmtId { get; set; } = numFmtId;
        public String Format { get; set; } = format;
    }
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（CSV 格式，逗号分隔）</summary>
    /// <returns>CSV 格式文本</returns>
    public String? ExtractText()
    {
        var sheets = Sheets;
        if (sheets == null || sheets.Count == 0) return null;

        var sb = new StringBuilder();
        var sheetList = sheets.ToList();
        for (var si = 0; si < sheetList.Count; si++)
        {
            var sheetName = sheetList[si];
            if (sheetList.Count > 1)
            {
                if (si > 0) sb.AppendLine();
                sb.AppendLine($"## {sheetName}");
            }

            foreach (var row in ReadRows(sheetName))
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
        var sheets = Sheets;
        if (sheets == null || sheets.Count == 0) return null;

        var sb = new StringBuilder();
        var sheetList = sheets.ToList();
        for (var si = 0; si < sheetList.Count; si++)
        {
            var sheetName = sheetList[si];
            if (sheetList.Count > 1)
            {
                if (si > 0) sb.AppendLine();
                sb.AppendLine($"## {sheetName}");
                sb.AppendLine();
            }

            var rows = ReadRows(sheetName).ToList();
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
        if (value.IndexOfAny([',', '"', '\n', '\r']) >= 0)
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }

    private static String MdEscape(String? value)
    {
        if (String.IsNullOrEmpty(value)) return "";
        return value.Replace("|", "\\|").Replace("\n", " ").Replace("\r", "");
    }
    #endregion
}
