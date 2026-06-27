using System.Globalization;
using System.Text;

namespace NewLife.Office;

/// <summary>PDF 交叉引用表（xref），按对象号精确定位每个间接对象</summary>
/// <remarks>
/// 解析 PDF 文件尾部的交叉引用表，支持传统 xref 表和 PDF 1.5+ 的 xref 流对象。
/// 支持增量更新链（多个 xref section 通过 /Prev 链接）。
/// 存储每个对象号的字节偏移量，为 PdfObjectParser 提供精确定位能力。
/// </remarks>
public class PdfXRefTable
{
    #region 属性
    /// <summary>对象号 → 条目（偏移量、代数、是否在用、所在流号）</summary>
    public Dictionary<Int32, XRefEntry> Entries { get; } = [];

    /// <summary>trailer 字典（合并所有 section 的 trailer）</summary>
    public PdfDict Trailer { get; } = new();

    /// <summary>对象流条目（ObjStm 压缩对象数流 → 对象号列表）</summary>
    public Dictionary<Int32, List<XRefEntry>> ObjectStreams { get; } = [];
    #endregion

    #region 构造
    /// <summary>从 PDF 字节数组解析交叉引用表</summary>
    /// <param name="data">PDF 完整字节</param>
    public PdfXRefTable(Byte[] data)
    {
        var latin1 = Encoding.GetEncoding(28591);
        var text = latin1.GetString(data);

        // 步骤 1：定位文件尾部的 startxref 指针
        var eofIdx = text.LastIndexOf("%%EOF", StringComparison.Ordinal);
        if (eofIdx < 0) return;

        // 从 %%EOF 向前搜索 startxref
        var startxrefIdx = text.LastIndexOf("startxref", eofIdx, StringComparison.Ordinal);
        if (startxrefIdx < 0) return;

        // 解析 startxref 后的偏移量
        var offsetStr = String.Empty;
        var pos = startxrefIdx + 9;
        while (pos < eofIdx && (text[pos] == ' ' || text[pos] == '\r' || text[pos] == '\n')) pos++;
        while (pos < eofIdx && text[pos] >= '0' && text[pos] <= '9')
            offsetStr += text[pos++];

        if (!Int64.TryParse(offsetStr, out var firstXrefOffset)) return;

        // 步骤 2：从第一个 xref 开始，沿 /Prev 链逐层解析
        var xrefOffsets = new HashSet<Int64>();
        ParseXRefChain(data, latin1, text, firstXrefOffset, xrefOffsets);
    }
    #endregion

    #region 方法
    /// <summary>获取指定对象号的字节偏移量，-1 表示不存在</summary>
    /// <param name="objNum">对象号</param>
    /// <returns>字节偏移量</returns>
    public Int64 GetOffset(Int32 objNum) => Entries.TryGetValue(objNum, out var e) ? e.Offset : -1;

    /// <summary>判断指定对象号是否存在且在用</summary>
    /// <param name="objNum">对象号</param>
    /// <returns>true 表示存在且在用的间接对象</returns>
    public Boolean HasObject(Int32 objNum) => Entries.TryGetValue(objNum, out var e) && e.InUse;

    /// <summary>获取对象所在的流号（压缩对象流场景），0 表示直接在文件偏移处</summary>
    /// <param name="objNum">对象号</param>
    /// <returns>流对象号</returns>
    public Int32 GetStreamObjNum(Int32 objNum) => Entries.TryGetValue(objNum, out var e) ? e.StreamObjNum : 0;

    /// <summary>获取对象在对象流中的索引</summary>
    /// <param name="objNum">对象号</param>
    /// <returns>对象流内索引</returns>
    public Int32 GetStreamIndex(Int32 objNum) => Entries.TryGetValue(objNum, out var e) ? e.StreamIndex : -1;
    #endregion

    #region 私有方法
    /// <summary>递归解析 xref 链（传统表+流格式+增量更新）</summary>
    private void ParseXRefChain(Byte[] data, Encoding latin1, String text, Int64 xrefOffset, HashSet<Int64> visited)
    {
        if (!visited.Add(xrefOffset) || xrefOffset >= data.Length) return;

        // 检查是否为 xref 流对象（PDF 1.5+ 格式）
        // xref 流对象以 "NNN 0 obj" 开头，字典中含 /Type /XRef
        if (IsXRefStream(data, xrefOffset, latin1, text))
        {
            ParseXRefStream(data, xrefOffset, latin1, text);
        }
        else
        {
            ParseXRefTable(data, xrefOffset, latin1, text);
        }

        // 检查 /Prev 链接到前一个 xref section（增量更新）
        if (Trailer.TryGetValue("Prev", out var prevVal) && prevVal is PdfNumber prevNum)
        {
            ParseXRefChain(data, latin1, text, (Int64)prevNum.Value, visited);
        }
    }

    /// <summary>解析传统 xref 表：xref\n0 6\n0000000000 65535 f \n...</summary>
    private void ParseXRefTable(Byte[] data, Int64 offset, Encoding latin1, String text)
    {
        var pos = (Int32)offset;
        // 跳过 "xref" 和换行
        while (pos < text.Length && text[pos] != '\n') pos++;
        if (pos < text.Length) pos++; // 跳过 \n

        var sectionStart = pos;
        while (sectionStart < text.Length)
        {
            // 读取 "startObj count" 行
            var lineEnd = text.IndexOf('\n', sectionStart);
            if (lineEnd < 0) break;
            var headerLine = text[sectionStart..lineEnd].Trim();
            sectionStart = lineEnd + 1;

            var parts = headerLine.Split(' ');
            if (!Int32.TryParse(parts[0], out var startObj) ||
                !Int32.TryParse(parts[1], out var count))
                break;

            // 避免无效的巨型 section
            if (count > 5000000) break;

            // 读取 count 条 xref 条目（每行 20 字节："NNNNNNNNNN GGGGG eol\n"）
            for (var i = 0; i < count && sectionStart + 19 < data.Length; i++)
            {
                var entryBytes = new Byte[20];
                var bytePos = sectionStart;
                // xref 条目使用 Latin-1 编码，每个字符 1 字节
                if (bytePos + 20 > data.Length) break;

                var entryLine = latin1.GetString(data, bytePos, 20);
                if (entryLine.Length < 18) break;

                var entryOffset = entryLine[..10];
                var entryGen = entryLine[11..16];
                var entryStatus = entryLine[17];

                if (Int64.TryParse(entryOffset, out var off) &&
                    Int32.TryParse(entryGen, out var gen))
                {
                    var objNum = startObj + i;
                    var inUse = entryStatus == 'n';
                    if (!Entries.ContainsKey(objNum))
                    {
                        Entries[objNum] = new XRefEntry
                        {
                            Offset = off,
                            GenNum = gen,
                            InUse = inUse,
                        };
                    }
                }
                sectionStart += 20;
            }

            // 偷看下一行，如果是数字则继续下一段
            if (sectionStart + 3 >= text.Length) break;
            var peekText = text.Substring(sectionStart, Math.Min(20, text.Length - sectionStart)).Trim();
            // 跳过空格看是否数字开头
            var peekParts = peekText.Split(' ');
            if (peekParts.Length < 2 || !Int32.TryParse(peekParts[0], out _) || !Int32.TryParse(peekParts[1], out var nextCount) || nextCount <= 0 || nextCount > 5000000)
                break;
        }

        // 读取 trailer 字典（在 "trailer" 关键字之后）
        var trailerIdx = text.IndexOf("trailer", sectionStart, StringComparison.Ordinal);
        if (trailerIdx < 0) return;

        var dictStart = text.IndexOf("<<", trailerIdx, Math.Min(200, text.Length - trailerIdx), StringComparison.Ordinal);
        if (dictStart < 0) return;

        var dictStr = ReadDictString(text, dictStart);
        if (dictStr != null)
            MergeTrailerDict(PdfObjectParser.ParseDictString(dictStr));
    }

    /// <summary>检查指定偏移处是否为 xref 流对象</summary>
    private static Boolean IsXRefStream(Byte[] data, Int64 offset, Encoding latin1, String text)
    {
        var o = (Int32)offset;
        // 读取 "NNN 0 obj" 
        var lineEnd = text.IndexOf('\n', o);
        if (lineEnd < 0 || lineEnd - o > 30) return false;
        var headerLine = text[o..lineEnd].Trim();
        if (!headerLine.EndsWith("obj")) return false;

        // 检查后续字典中是否有 /Type /XRef
        var dictStart = text.IndexOf("<<", lineEnd, Math.Min(200, text.Length - lineEnd), StringComparison.Ordinal);
        if (dictStart < 0 || dictStart - lineEnd > 100) return false;

        var dictText = ReadDictString(text, dictStart);
        return dictText != null && dictText.IndexOf("/Type /XRef", StringComparison.Ordinal) >= 0;
    }

    /// <summary>解析 xref 流对象（PDF 1.5+ 格式）</summary>
    private void ParseXRefStream(Byte[] data, Int64 offset, Encoding latin1, String text)
    {
        var o = (Int32)offset;
        var lineEnd = text.IndexOf('\n', o);
        if (lineEnd < 0) return;

        // 解析对象字典
        var dictStart = text.IndexOf("<<", lineEnd, StringComparison.Ordinal);
        if (dictStart < 0) return;

        var dictStr = ReadDictString(text, dictStart);
        if (dictStr == null) return;

        var dict = PdfObjectParser.ParseDictString(dictStr);
        MergeTrailerDict(dict); // xref 流字典同时是 trailer

        // 解析流数据
        var streamStart = text.IndexOf("stream", dictStart + dictStr.Length, StringComparison.Ordinal);
        if (streamStart < 0) return;

        var contentStart = streamStart + 6;
        if (contentStart < text.Length && text[contentStart] == '\r') contentStart++;
        if (contentStart < text.Length && text[contentStart] == '\n') contentStart++;

        var streamEnd = text.IndexOf("endstream", contentStart, StringComparison.Ordinal);
        if (streamEnd < 0) return;

        var streamData = new Byte[streamEnd - contentStart];
        Array.Copy(data, contentStart, streamData, 0, streamData.Length);

        // 解压流数据（xref 流通常使用 FlateDecode）
        if (dict.TryGetValue("Filter", out var filterVal) && filterVal is PdfName fn && fn.Value == "FlateDecode")
        {
            try
            {
                streamData = PdfReader.DecompressFlate(streamData);
            }
            catch
            {
                return; // 解压失败则跳过
            }
        }

        ParseXRefStreamData(streamData, dict);
    }

    /// <summary>解析解压后的 xref 流数据</summary>
    private void ParseXRefStreamData(Byte[] streamData, PdfDict dict)
    {
        // 读取 /W 数组（每条目宽度：[type_size, field2_size, field3_size]）
        if (!dict.TryGetValue("W", out var wVal) || wVal is not PdfArray wArr || wArr.Items.Count < 3) return;

        var w0 = (Int32)((PdfNumber)wArr.Items[0]).Value;
        var w1 = (Int32)((PdfNumber)wArr.Items[1]).Value;
        var w2 = (Int32)((PdfNumber)wArr.Items[2]).Value;
        var entrySize = w0 + w1 + w2;
        if (entrySize <= 0 || entrySize > 20) return;

        // 读取 /Index 数组（[startObj, count, ...]），默认 [0, Size]
        var indexPairs = new List<(Int32 Start, Int32 Count)>();
        if (dict.TryGetValue("Index", out var idxVal) && idxVal is PdfArray idxArr)
        {
            for (var i = 0; i + 1 < idxArr.Items.Count; i += 2)
                indexPairs.Add(((Int32)((PdfNumber)idxArr.Items[i]).Value, (Int32)((PdfNumber)idxArr.Items[i + 1]).Value));
        }
        else if (dict.TryGetValue("Size", out var szVal) && szVal is PdfNumber szNum)
        {
            indexPairs.Add((0, (Int32)szNum.Value));
        }
        else return;

        var pos = 0;
        foreach (var (startObj, count) in indexPairs)
        {
            for (var i = 0; i < count && pos + entrySize <= streamData.Length; i++)
            {
                var objNum = startObj + i;
                var type = ReadField(streamData, ref pos, w0);
                var field2 = ReadField(streamData, ref pos, w1);
                var field3 = ReadField(streamData, ref pos, w2);

                switch (type)
                {
                    case 0: // 空闲对象
                        if (!Entries.ContainsKey(objNum))
                            Entries[objNum] = new XRefEntry { Offset = field2, GenNum = (Int32)field3, InUse = false };
                        break;
                    case 1: // 在用对象（直接偏移）
                        if (!Entries.ContainsKey(objNum))
                            Entries[objNum] = new XRefEntry { Offset = field2, GenNum = (Int32)field3, InUse = true };
                        break;
                    case 2: // 压缩对象（在对象流中）
                        if (!Entries.ContainsKey(objNum))
                            Entries[objNum] = new XRefEntry
                            {
                                StreamObjNum = (Int32)field2,
                                StreamIndex = (Int32)field3,
                                InUse = true,
                            };
                        // 记录对象流
                        if (!ObjectStreams.ContainsKey((Int32)field2))
                            ObjectStreams[(Int32)field2] = [];
                        ObjectStreams[(Int32)field2].Add(new XRefEntry
                        {
                            Offset = objNum,
                            GenNum = (Int32)field3,
                            InUse = true,
                        });
                        break;
                }
            }
        }
    }

    /// <summary>从指定位置读取宽度为 w 字节的大端整数字段</summary>
    private static Int64 ReadField(Byte[] data, ref Int32 pos, Int32 w)
    {
        Int64 val = 0;
        for (var i = 0; i < w && pos < data.Length; i++)
            val = (val << 8) | data[pos++];
        return val;
    }

    /// <summary>读取 PDF 字典字符串（<< ... >>），处理嵌套</summary>
    private static String? ReadDictString(String text, Int32 start)
    {
        var pos = start;
        if (pos + 2 > text.Length || text[pos] != '<' || text[pos + 1] != '<') return null;

        var depth = 0;
        var sb = new StringBuilder();
        while (pos < text.Length)
        {
            var c = text[pos];
            sb.Append(c);

            if (c == '<' && pos + 1 < text.Length && text[pos + 1] == '<')
            {
                depth++;
                sb.Append('<');
                pos += 2;
                continue;
            }
            if (c == '>' && pos + 1 < text.Length && text[pos + 1] == '>')
            {
                sb.Append('>');
                depth--;
                pos += 2;
                if (depth == 0) return sb.ToString();
                continue;
            }
            pos++;
        }
        return null;
    }

    /// <summary>合并 trailer 字典到主字典（后续 section 的值不覆盖已有值）</summary>
    private void MergeTrailerDict(PdfDict newDict)
    {
        foreach (var kv in newDict)
        {
            if (!Trailer.ContainsKey(kv.Key))
                Trailer[kv.Key] = kv.Value;
        }
    }
    #endregion
}

/// <summary>交叉引用表条目</summary>
public class XRefEntry
{
    #region 属性
    /// <summary>字节偏移量（传统 xref）或对象号（对象流索引）</summary>
    public Int64 Offset { get; set; }

    /// <summary>代数号</summary>
    public Int32 GenNum { get; set; }

    /// <summary>对象是否在用</summary>
    public Boolean InUse { get; set; }

    /// <summary>对象流号（压缩对象，type=2），0 表示直接偏移</summary>
    public Int32 StreamObjNum { get; set; }

    /// <summary>在对象流中的对象索引</summary>
    public Int32 StreamIndex { get; set; }
    #endregion
}
