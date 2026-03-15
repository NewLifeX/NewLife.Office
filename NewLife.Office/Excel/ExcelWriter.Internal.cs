using System.Globalization;
using System.Security;
using NewLife.Collections;

namespace NewLife.Office;

partial class ExcelWriter
{
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
}
