using System.Text;
using NewLife.Buffers;
using NewLife.Collections;
using NewLife.Office;
using NewLife.Office.Ole2;

namespace NewLife.Office.Word;

/// <summary>Word 97-2003 二进制（.doc）文档读取器</summary>
/// <remarks>
/// 基于 OLE2/CFB 容器解析 MS-DOC 格式，通过 CLX 段信息提取纯文本与段落结构。
/// 仅支持 Word 97 及以后生成的 .doc 文件（二进制格式标识 0xA5EC）。
/// <para>用法示例：</para>
/// <code>
/// using var reader = new DocReader("document.doc");
/// foreach (var para in reader.ReadParagraphs())
///     Console.WriteLine(para);
/// </code>
/// </remarks>
public sealed class DocReader : IDisposable, ITextExtractable, IMarkdownExtractable
{
    #region 属性

    /// <summary>文档全文（已缓存）</summary>
    private String? _fullText;

    private Boolean _disposed;

    #endregion

    #region 私有字段

    private readonly Byte[] _wordDoc;

    /// <summary>表格流（1Table/0Table），fast-saved 文档的 CHPX/PAPX/CLX 可能存于此</summary>
    private readonly Byte[]? _tableDoc;

    #endregion

    #region 构造

    /// <summary>从 doc 文件路径打开</summary>
    /// <param name="path">doc 文件路径</param>
    public DocReader(String path)
    {
        using var doc = CfbDocument.Open(path);
        _wordDoc = GetWordDocStream(doc);
        _tableDoc = doc.GetStreamData("1Table") ?? doc.GetStreamData("0Table");
        ValidateFib(_wordDoc);
    }

    /// <summary>从流打开（需包含 doc 的完整 OLE2 容器内容）</summary>
    /// <param name="stream">可读流</param>
    public DocReader(Stream stream)
    {
        using var doc = CfbDocument.Open(stream, leaveOpen: true);
        _wordDoc = GetWordDocStream(doc);
        _tableDoc = doc.GetStreamData("1Table") ?? doc.GetStreamData("0Table");
        ValidateFib(_wordDoc);
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

    private static Byte[] GetWordDocStream(CfbDocument doc)
    {
        var data = doc.GetStreamData("WordDocument");
        if (data == null || data.Length < 32)
            throw new InvalidDataException("找不到 WordDocument 流，文件可能不是有效的 .doc 格式。");
        return data;
    }

    private static void ValidateFib(Byte[] buf)
    {
        var reader = new SpanReader(buf, 0, 2);
        var wIdent = reader.ReadUInt16();
        // 0xA5EC = Word 二进制文档；0xA5DC = Word 模板
        if (wIdent != 0xA5EC && wIdent != 0xA5DC)
            throw new InvalidDataException($"不支持的文档格式：wIdent = 0x{wIdent:X4}，仅支持 Word 97-2003 二进制格式。");
    }

    #endregion

    #region 读取方法

    /// <summary>读取文档全文</summary>
    /// <returns>文档全文，段落以换行符分隔</returns>
    public String ReadFullText()
    {
        if (_fullText == null) _fullText = BuildFullText();
        return _fullText;
    }

    /// <summary>逐段落读取文档文本</summary>
    /// <returns>非空段落序列</returns>
    public IEnumerable<String> ReadParagraphs()
    {
        var text = ReadFullText();
        var start = 0;
        for (var i = 0; i <= text.Length; i++)
        {
            if (i == text.Length || text[i] == '\n')
            {
                var len = i - start;
                if (len > 0)
                    yield return text.Substring(start, len);
                start = i + 1;
            }
        }
    }

    /// <summary>读取文档中的所有表格</summary>
    /// <remarks>
    /// 通过检测原始文本流中的 0x07（表格单元格结束符）识别表格行，
    /// 将连续的表格行聚合为表格，每格内容已 Trim 处理。
    /// 空表格行（所有单元格均为空）会被自动跳过。
    /// </remarks>
    /// <returns>表格序列，每张表格为 String[][] （行 × 列）</returns>
    public IEnumerable<String[][]> ReadTables()
    {
        var rawText = BuildRawText(keepTableMarkers: true);
        var table = new List<String[]>();
        var start = 0;

        for (var i = 0; i <= rawText.Length; i++)
        {
            if (i < rawText.Length && rawText[i] != '\n') continue;

            var line = rawText[start..i];
            start = i + 1;

            if (line.Contains('\x07'))
            {
                // 表格行：以 \x07 为分隔符分割单元格
                var parts = line.Split('\x07');
                // 最后一个 \x07 之后通常是空字符串，过滤
                var cells = parts
                    .Select(p => p.Trim())
                    .ToArray();
                // 去掉尾部多余的空单元格（行末的 \x07 产生）
                var lastNonEmpty = cells.Length - 1;
                while (lastNonEmpty >= 0 && cells[lastNonEmpty].Length == 0)
                {
                    lastNonEmpty--;
                }

                if (lastNonEmpty >= 0)
                    table.Add(cells.Take(lastNonEmpty + 1).ToArray());
            }
            else
            {
                // 非表格行：如果之前有积累的表格行，则结束当前表格
                if (table.Count > 0)
                {
                    yield return table.ToArray();
                    table.Clear();
                }
            }
        }

        // 文档末尾若有未输出的表格
        if (table.Count > 0)
            yield return table.ToArray();
    }

    /// <summary>读取带格式段落（基于 CHPX/PAPX 解析粗体/斜体/字号/字体/颜色/对齐/缩进/列表）</summary>
    /// <remarks>
    /// 通过 FIB 定位 PlcBte 表（CHPX/PAPX），按 CP 范围解析段落格式。
    /// 支持 WordDocument 与 1Table/0Table 两种存放位置；未知 sprm 处停止解析（尽力而为）。
    /// </remarks>
    /// <returns>带格式段落序列</returns>
    public IEnumerable<DocParagraph> ReadParagraphsWithFormat()
    {
        var buf = _wordDoc;
        if (buf.Length < 300) yield break;

        // FIB FibRgFcLcb97：12=PlcBteChpx，13=CLX，14=PlcBtePapx
        List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)>? chpxList = null;
        List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)>? papxList = null;
        if (TryGetFcLcb(buf, 12, out var fcChpx, out var lcbChpx) &&
            TryGetFcLcb(buf, 14, out var fcPapx, out var lcbPapx))
        {
            chpxList = ParsePlcBte(fcChpx, lcbChpx);
            papxList = ParsePlcBte(fcPapx, lcbPapx);
        }
        // 无 CHPX/PAPX 表时回退纯文本段落（无格式）
        var hasFormat = chpxList is { Count: > 0 };
        if (!hasFormat)
        {
            foreach (var para in ReadParagraphs())
                yield return new DocParagraph { Text = para };
            yield break;
        }

        // 构建全文（记录每个字符的真实 CP；域控制符占位不产出，文本索引与 CP 偏移不一致）
        // 段落起始 CP 用 cps 映射，避免含域/表格的文档段落格式错配
        var (text, cps) = BuildTextWithCps();
        if (text.Length == 0) yield break;

        // 按段落符（0x0D 段落 / 0x0C 分页 / 0x0B 分栏）切分
        var paraStart = 0;
        for (var i = 0; i <= text.Length; i++)
        {
            var isSep = i == text.Length || text[i] == '\r' || text[i] == '\n' || text[i] == '\f' || text[i] == '\v';
            if (!isSep) continue;

            var len = i - paraStart;
            // 去掉表格单元格结束符 0x07（保留 CP 定位，但不进入文本输出）
            var paraText = text.Substring(paraStart, len).TrimEnd('\r', '\n', '\f', '\v').Replace("\x07", "");
            if (paraText.Length > 0)
            {
                var cp = paraStart < cps.Length ? cps[paraStart] : paraStart;
                yield return BuildDocParagraph(paraText, cp, chpxList!, papxList!);
            }
            paraStart = i + 1;
        }
    }

    #endregion

    #region 格式提取辅助

    /// <summary>从 FIB FibRgFcLcb97 读取指定索引的 fc/lcb 对</summary>
    private static Boolean TryGetFcLcb(Byte[] buf, Int32 index, out Int32 fc, out Int32 lcb)
    {
        fc = 0;
        lcb = 0;
        if (buf.Length < 300) return false;

        var reader = new SpanReader(buf, 32, buf.Length - 32);
        var csw = (Int32)reader.ReadUInt16();
        if (csw < 1) csw = 14;

        var cslwOffset = 34 + csw * 2;
        if (cslwOffset + 2 > buf.Length) return false;

        reader = new SpanReader(buf, cslwOffset, buf.Length - cslwOffset);
        var cslw = (Int32)reader.ReadUInt16();
        if (cslw < 1) cslw = 22;

        var fcLcbStart = cslwOffset + 2 + cslw * 4 + 2;
        var pos = fcLcbStart + index * 8;
        if (pos + 8 > buf.Length) return false;

        reader = new SpanReader(buf, pos, 8);
        fc = (Int32)reader.ReadUInt32();
        lcb = (Int32)reader.ReadUInt32();
        return fc >= 0 && lcb > 0;
    }

    /// <summary>解析 PlcBte（CP 数组 + BTE 数组），返回 CP 范围 → grpprl 字节</summary>
    private List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)>? ParsePlcBte(Int32 fc, Int32 lcb)
    {
        var src = _wordDoc;
        if (fc < 0 || lcb <= 4 || fc + lcb > src.Length) return null;
        var n = (lcb - 4) / 8;
        if (n <= 0) return null;

        var result = new List<(Int32, Int32, Byte[])>(n);
        var reader = new SpanReader(src, fc, lcb);
        var cps = new Int32[n + 1];
        for (var i = 0; i <= n; i++) cps[i] = (Int32)reader.ReadUInt32();

        for (var i = 0; i < n; i++)
        {
            var bte = (Int32)reader.ReadUInt32();
            var fTable = (bte & 0x40000000) != 0; // bit30：数据在表格流
            var dataFc = bte & 0x3FFFFFFF;
            Byte[]? data = null;
            if (fTable)
            {
                if (_tableDoc != null && dataFc >= 0 && dataFc < _tableDoc.Length)
                    data = _tableDoc;
            }
            else if (dataFc >= 0 && dataFc < src.Length)
            {
                data = src;
            }

            var grpprl = new Byte[0];
            if (data != null)
            {
                // CHPX/PAPX：1 字节 cbGrpprl + grpprl（PAPX 其后还有 12 字节 PHE，不影响读取）
                var cbGrpprl = data[dataFc];
                var grpprlLen = Math.Min(cbGrpprl, data.Length - dataFc - 1);
                if (grpprlLen > 0)
                {
                    grpprl = new Byte[grpprlLen];
                    Array.Copy(data, dataFc + 1, grpprl, 0, grpprlLen);
                }
            }
            result.Add((cps[i], cps[i + 1], grpprl));
        }
        return result;
    }

    /// <summary>按 CP 查找所在范围的 grpprl</summary>
    private static Byte[] GetGrpprlAt(List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)> list, Int32 cp)
    {
        foreach (var (s, e, g) in list)
        {
            if (cp >= s && cp < e) return g;
        }
        if (list.Count > 0 && cp >= list[^1].Item1) return list[^1].Item3;
        return [];
    }

    /// <summary>根据 CP 位置构建带格式段落</summary>
    private static DocParagraph BuildDocParagraph(String text, Int32 cp,
        List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)> chpxList,
        List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)> papxList)
    {
        var p = new DocParagraph { Text = text };
        // 段落格式取段落首字符（或段落标记位置）的 PAPX
        var papx = GetGrpprlAt(papxList, cp);
        if (papx.Length > 0) ParsePapx(papx, p);
        // 字符格式取段落首字符的 CHPX（合并整段同格式的常见场景）
        var chpx = GetGrpprlAt(chpxList, cp);
        if (chpx.Length > 0) ParseChpx(chpx, p);
        return p;
    }

    /// <summary>解析 CHPX grpprl 中的字符格式 sprms</summary>
    private static void ParseChpx(Byte[] grpprl, DocParagraph p)
    {
        var i = 0;
        while (i < grpprl.Length)
        {
            var first = grpprl[i];
            var twoByte = (first & 0x80) != 0;
            Int32 sprm;
            if (twoByte)
            {
                if (i + 1 >= grpprl.Length) break;
                sprm = ((first & 0x7F) << 8) | grpprl[i + 1];
                i += 2;
            }
            else
            {
                sprm = first;
                i += 1;
            }

            switch (sprm)
            {
                case 0x0835: // sprmCFBold
                    if (i < grpprl.Length) p.Bold = grpprl[i] != 0;
                    i += 1;
                    break;
                case 0x0836: // sprmCFItalic
                    if (i < grpprl.Length) p.Italic = grpprl[i] != 0;
                    i += 1;
                    break;
                case 0x0837: // sprmCFStrike
                    if (i < grpprl.Length) p.Strikethrough = grpprl[i] != 0;
                    i += 1;
                    break;
                case 0x083A: // sprmCIS：bit0=斜体，bit1=粗体
                    if (i < grpprl.Length)
                    {
                        var v = grpprl[i];
                        p.Italic = (v & 0x01) != 0;
                        p.Bold = (v & 0x02) != 0;
                    }
                    i += 1;
                    break;
                case 0x083C: // sprmCFS：1 字节半磅字号
                    if (i < grpprl.Length) p.FontSize = grpprl[i] / 2f;
                    i += 1;
                    break;
                case 0x0425: // sprmCHps：2 字节半磅字号
                    if (i + 1 < grpprl.Length) p.FontSize = (grpprl[i] | (grpprl[i + 1] << 8)) / 2f;
                    i += 2;
                    break;
                case 0x4A2F: // sprmCFontName 短字符串：2 字节长度 + ANSI(CP1252)
                    i = ParseFontName(grpprl, i, p, false);
                    break;
                case 0x6A2F: // sprmCFontName 长字符串：2 字节长度 + UTF-16
                    i = ParseFontName(grpprl, i, p, true);
                    break;
                case 0x4C2E: // sprmCCv：4 字节 COLORREF（BGR）
                    if (i + 3 < grpprl.Length)
                    {
                        var b = grpprl[i]; var g = grpprl[i + 1]; var r = grpprl[i + 2];
                        p.ForeColor = $"{r:X2}{g:X2}{b:X2}";
                    }
                    i += 4;
                    break;
                default:
                    // 未知 sprm：操作数长度不定，停止解析（尽力而为）
                    return;
            }
        }
    }

    /// <summary>解析 sprmCFontName 字体名字符串，返回操作数结束后的偏移</summary>
    private static Int32 ParseFontName(Byte[] grpprl, Int32 i, DocParagraph p, Boolean unicode)
    {
        if (i + 1 >= grpprl.Length) return grpprl.Length;
        var len = grpprl[i] | (grpprl[i + 1] << 8);
        var start = i + 2;
        var byteLen = unicode ? len * 2 : len;
        if (byteLen >= 0 && start + byteLen <= grpprl.Length)
        {
            if (len > 0)
                p.FontName = (unicode ? Encoding.Unicode : Encoding.GetEncoding(1252)).GetString(grpprl, start, byteLen);
            return start + byteLen;
        }
        return grpprl.Length;
    }

    /// <summary>解析 PAPX grpprl 中的段落格式 sprms</summary>
    private static void ParsePapx(Byte[] grpprl, DocParagraph p)
    {
        var i = 0;
        while (i < grpprl.Length)
        {
            var first = grpprl[i];
            var twoByte = (first & 0x80) != 0;
            Int32 sprm;
            if (twoByte)
            {
                if (i + 1 >= grpprl.Length) break;
                sprm = ((first & 0x7F) << 8) | grpprl[i + 1];
                i += 2;
            }
            else
            {
                sprm = first;
                i += 1;
            }

            switch (sprm)
            {
                case 0x2403: // sprmPJc：0=left 1=center 2=right 3=justify
                    if (i < grpprl.Length)
                    {
                        p.Alignment = grpprl[i] switch
                        {
                            1 => "center",
                            2 => "right",
                            3 => "justify",
                            _ => "left",
                        };
                    }
                    i += 1;
                    break;
                case 0x2405: // sprmPDxaLeft（4 字节有符号）
                    if (i + 3 < grpprl.Length) p.IndentLeft = ReadInt32(grpprl, i);
                    i += 4;
                    break;
                case 0x2406: // sprmPDxaRight（4 字节）
                    i += 4;
                    break;
                case 0x2407: // sprmPDxaLeft1（4 字节首行缩进）
                    if (i + 3 < grpprl.Length)
                    {
                        var v = ReadInt32(grpprl, i);
                        // 正值=首行缩进，负值=悬挂缩进
                        p.FirstLineIndent = v;
                    }
                    i += 4;
                    break;
                case 0x2416: // sprmPDyaBefore（4 字节）
                    if (i + 3 < grpprl.Length) p.SpaceBefore = ReadInt32(grpprl, i);
                    i += 4;
                    break;
                case 0x2417: // sprmPDyaAfter（4 字节）
                    if (i + 3 < grpprl.Length) p.SpaceAfter = ReadInt32(grpprl, i);
                    i += 4;
                    break;
                case 0x241A: // sprmPIlvl（1 字节列表级别）
                    if (i < grpprl.Length) p.ListLevel = grpprl[i];
                    i += 1;
                    break;
                case 0x2A0D: // sprmPNumId（2 字节编号定义 ID）
                    if (i + 1 < grpprl.Length) p.NumberingId = grpprl[i] | (grpprl[i + 1] << 8);
                    i += 2;
                    break;
                case 0x241B: // sprmPDyaLine（4 字节行距）
                    i += 4;
                    break;
                default:
                    return;
            }
        }
    }

    private static Int32 ReadInt32(Byte[] buf, Int32 i) =>
        buf[i] | (buf[i + 1] << 8) | (buf[i + 2] << 16) | (buf[i + 3] << 24);





    #endregion

    #region FIB 解析与文本提取

    /// <summary>解析 FIB，定位 CLX，提取所有文本</summary>
    private String BuildFullText() => BuildRawText(keepTableMarkers: false);

    /// <summary>提取文档文本，可选保留表格单元格标记符（0x07）</summary>
    /// <param name="keepTableMarkers">true = 保留 0x07 用于表格检测</param>
    private String BuildRawText(Boolean keepTableMarkers)
    {
        var buf = _wordDoc;
        if (buf.Length < 300) return String.Empty;

        // ─── 1. 定位 FibRgFcLcb97 中的 fcClx / lcbClx ──────────────────────
        var reader = new SpanReader(buf, 32, buf.Length - 32);
        var csw = (Int32)reader.ReadUInt16();         // FIB base = 32 bytes
        if (csw < 1) csw = 14;                        // Word 97 默认 14

        var cslwOffset = 34 + csw * 2;
        if (cslwOffset + 2 > buf.Length) return String.Empty;

        reader = new SpanReader(buf, cslwOffset, buf.Length - cslwOffset);
        var cslw = (Int32)reader.ReadUInt16();
        if (cslw < 1) cslw = 22;                      // Word 97 默认 22

        // FibRgFcLcb97 起始偏移
        var fcLcbStart = cslwOffset + 2 + cslw * 4 + 2;
        // CLX 条目在 FibRgFcLcb97 中的索引为 13
        var fcClxPos = fcLcbStart + 13 * 8;

        if (fcClxPos + 8 > buf.Length) return String.Empty;

        reader = new SpanReader(buf, fcClxPos, 8);
        var fcClx = (Int32)reader.ReadUInt32();
        var lcbClx = (Int32)reader.ReadUInt32();

        if (fcClx < 0 || lcbClx <= 0 || (Int64)fcClx + lcbClx > buf.Length)
            return String.Empty;

        // ─── 2. 解析 CLX ──────────────────────────────────────────────────────
        var clxReader = new SpanReader(buf, fcClx, lcbClx);
        var lcbPlcPcd = 0;
        var foundPcdt = false;

        while (clxReader.Position < clxReader.Capacity)
        {
            var clxt = clxReader.ReadByte();
            if (clxt == 0x01)
            {
                // PRC 记录：跳过
                if (clxReader.Position + 2 > clxReader.Capacity) break;
                var cbGrpprl = (Int16)clxReader.ReadUInt16();
                clxReader.Advance(cbGrpprl);
            }
            else if (clxt == 0x02)
            {
                // PCDT 记录
                if (clxReader.Position + 4 > clxReader.Capacity) return String.Empty;
                lcbPlcPcd = (Int32)clxReader.ReadUInt32();
                foundPcdt = true;
                break;
            }
            else
            {
                // 未知 clxt 类型，无法继续
                return String.Empty;
            }
        }

        if (!foundPcdt) return String.Empty;

        var plcPcdStart = fcClx + (Int32)clxReader.Position;
        var plcPcdEnd = plcPcdStart + lcbPlcPcd;
        if (plcPcdEnd > fcClx + lcbClx) plcPcdEnd = fcClx + lcbClx;

        // ─── 3. 解析 PlcPcd：(n+1) 个 CP 值 + n 个 PCD（各8字节）───────────
        // n = (lcbPlcPcd - 4) / 12
        var pieceCount = (lcbPlcPcd - 4) / 12;
        if (pieceCount <= 0) return String.Empty;

        // CP 数组：pieceCount+1 个 UInt32 值
        var cpReader = new SpanReader(buf, plcPcdStart, (pieceCount + 1) * 4);
        // PCD 数组：pieceCount 个 8 字节结构
        var pcdReader = new SpanReader(buf, plcPcdStart + (pieceCount + 1) * 4, pieceCount * 8);

        // ─── 4. 提取每个 piece 的文本 ─────────────────────────────────────
        var sb = Pool.StringBuilder.Get();
        var cpStart = (Int32)cpReader.ReadUInt32();

        for (var i = 0; i < pieceCount; i++)
        {
            var cpEnd = (Int32)cpReader.ReadUInt32();
            var charCount = cpEnd - cpStart;

            // PCD 结构 8 字节：clsPcd(2) + FcCompressed(4) + Prm(2)
            pcdReader.Advance(2);
            var fcCompressed = (Int32)pcdReader.ReadUInt32();
            pcdReader.Advance(2);

            cpStart = cpEnd;

            if (charCount <= 0) continue;

            var fCompressed = ((fcCompressed >> 30) & 1) == 1;
            var fc = fcCompressed & 0x3FFFFFFF;

            if (fCompressed)
            {
                // ANSI（Latin-1）存储：fc 是压缩偏移，实际字节偏移 = fc / 2
                var byteOffset = fc / 2;
                var byteCount = charCount;
                if (byteOffset < 0 || byteOffset + byteCount > buf.Length)
                    continue;

                for (var c = 0; c < byteCount; c++)
                {
                    var ch = (Char)buf[byteOffset + c];
                    AppendDocChar(sb, ch, keepTableMarkers);
                }
            }
            else
            {
                // Unicode 存储：每字符 2 字节
                var byteOffset = fc;
                var byteCount = charCount * 2;
                if (byteOffset < 0 || byteOffset + byteCount > buf.Length)
                    continue;

                var textReader = new SpanReader(buf, byteOffset, byteCount);
                for (var c = 0; c < charCount; c++)
                {
                    var ch = (Char)textReader.ReadUInt16();
                    AppendDocChar(sb, ch, keepTableMarkers);
                }
            }
        }

        return sb.Return(true);
    }

    /// <summary>将文档字符追加到 StringBuilder，过滤控制字符并转换段落符</summary>
    /// <param name="sb">目标 StringBuilder</param>
    /// <param name="ch">文档字符</param>
    /// <param name="keepTableMarkers">是否保留 0x07 表格单元格标记符</param>
    private static void AppendDocChar(StringBuilder sb, Char ch, Boolean keepTableMarkers = false)
    {
        switch (ch)
        {
            case '\r':   // 段落结束符（0x0D）
            case '\f':   // 分页符（0x0C）
            case '\v':   // 分栏符（0x0B）
                sb.Append('\n');
                break;
            case '\x07': // 表格单元格结束符
                if (keepTableMarkers) sb.Append('\x07');
                break;
            case '\x13': // 域开始，跳过
            case '\x14': // 域分隔，跳过
            case '\x15': // 域结束，跳过
                break;
            default:
                if (ch >= ' ' || ch == '\t')
                    sb.Append(ch);
                break;
        }
    }

    #endregion

    #region 文本提取
    /// <summary>提取纯文本（段落间换行分隔）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText() => ReadFullText();

    /// <summary>提取 Markdown 格式（段落间空行分隔）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown()
    {
        var sb = new StringBuilder();
        foreach (var para in ReadParagraphs())
        {
            sb.AppendLine(para);
            sb.AppendLine();
        }
        return sb.ToString();
    }
    #endregion

    #region doc → docx 转换（W20）

    /// <summary>
    /// 将 doc 文档转换为 docx 文档模型，配合 <see cref="WordWriter.Save(String, Document)"/> 输出 docx。
    /// </summary>
    /// <remarks>
    /// 保留段落/表格的文档顺序与常用格式（粗体/斜体/删除线/字号/字体/颜色/对齐/缩进/列表级别）。
    /// 表格单元格以 0x07 标记识别，段落格式基于 CHPX/PAPX 按真实 CP 对齐。
    /// </remarks>
    /// <returns>docx 文档模型</returns>
    public Document ToDocument()
    {
        var doc = new Document();

        // 1. 解析 CHPX/PAPX 格式表（与 ReadParagraphsWithFormat 相同定位）
        List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)>? chpxList = null;
        List<(Int32 CpStart, Int32 CpEnd, Byte[] Grpprl)>? papxList = null;
        if (TryGetFcLcb(_wordDoc, 12, out var fcChpx, out var lcbChpx) &&
            TryGetFcLcb(_wordDoc, 14, out var fcPapx, out var lcbPapx))
        {
            chpxList = ParsePlcBte(fcChpx, lcbChpx);
            papxList = ParsePlcBte(fcPapx, lcbPapx);
        }

        // 2. 构建保留表格标记（\x07）与段落符（\r）的文本及真实 CP 数组
        var (text, cps) = BuildTextWithCps();
        if (text.Length == 0) return doc;

        var currentTable = new List<List<Cell>>();
        var start = 0;

        // 结束当前表格（遇到非表格行或文档末尾）
        void FlushTable()
        {
            if (currentTable.Count > 0)
            {
                // 复制列表再清空，避免引用别名导致已加入元素的 TableRows 被清空
                doc.Elements.Add(new Element { Type = ElementType.Table, TableRows = new List<List<Cell>>(currentTable) });
                currentTable.Clear();
            }
        }

        for (var i = 0; i <= text.Length; i++)
        {
            var isSep = i == text.Length || text[i] == '\r' || text[i] == '\f' || text[i] == '\v';
            if (!isSep) continue;

            var line = text[start..i];
            start = i + 1;

            if (line.Contains('\x07'))
            {
                // 表格行：以 \x07 为分隔符分割单元格
                var cells = line.Split('\x07').Select(p => p.Trim()).ToList();
                while (cells.Count > 0 && cells[^1].Length == 0) cells.RemoveAt(cells.Count - 1);
                if (cells.Count > 0)
                {
                    var row = new List<Cell>();
                    foreach (var cellText in cells)
                        row.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = cellText } } } } });
                    currentTable.Add(row);
                }
            }
            else if (line.Length > 0)
            {
                // 普通段落：结束当前表格，按 CP 应用段落/字符格式
                FlushTable();
                var lineCp = cps[start - line.Length - 1]; // 行首字符的真实 CP
                var dp = BuildDocParagraph(line, lineCp, chpxList ?? [], papxList ?? []);
                doc.Elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = ToWordParagraph(dp) });
            }
        }
        FlushTable();

        return doc;
    }

    /// <summary>
    /// 将 doc 文档转换为 docx 文件（W20）。
    /// </summary>
    /// <param name="outputPath">输出 docx 文件路径</param>
    public void SaveAsDocx(String outputPath)
    {
        using var writer = new WordWriter();
        writer.Save(outputPath, ToDocument());
    }

    /// <summary>
    /// 将 doc 文档转换为 docx 并写入流（W20）。
    /// </summary>
    /// <param name="outputStream">输出流</param>
    public void SaveAsDocx(Stream outputStream)
    {
        using var writer = new WordWriter();
        writer.Save(outputStream, ToDocument());
    }

    /// <summary>将 doc 带格式段落映射为 docx 段落模型</summary>
    private static Paragraph ToWordParagraph(DocParagraph dp)
    {
        var para = new Paragraph
        {
            Alignment = dp.Alignment,
            IndentLeft = dp.IndentLeft,
            FirstLineIndent = dp.FirstLineIndent,
            SpaceBefore = dp.SpaceBefore,
            SpaceAfter = dp.SpaceAfter,
        };

        // 列表：PAPX 列表级别 + 编号定义 ID
        if (dp.ListLevel >= 0 && dp.NumberingId > 0)
        {
            para.IsBullet = true; // doc 编号简化为项目符号（无 numbering.xml 定义）
            para.ListLevel = dp.ListLevel;
        }

        var props = new RunProperties
        {
            Bold = dp.Bold ? true : null,
            Italic = dp.Italic ? true : null,
            Strikethrough = dp.Strikethrough ? true : null,
            FontSize = dp.FontSize,
            FontName = dp.FontName,
            ForeColor = dp.ForeColor,
        };
        para.Runs.Add(new Run { Text = dp.Text, Properties = props });
        return para;
    }

    /// <summary>
    /// 构建保留表格单元格标记（\x07）与段落分隔符（\r）的文本，并返回每个字符的真实 CP。
    /// CP 计数包含 \x07 等控制符，与 CHPX/PAPX 的 PlcBte 表对齐，用于段落格式定位。
    /// </summary>
    private (String Text, Int32[] Cps) BuildTextWithCps()
    {
        var buf = _wordDoc;
        if (buf.Length < 300) return (String.Empty, []);

        // ─── 定位 CLX / PlcPcd（与 BuildRawText 相同）────────────────────
        var reader = new SpanReader(buf, 32, buf.Length - 32);
        var csw = (Int32)reader.ReadUInt16();
        if (csw < 1) csw = 14;
        var cslwOffset = 34 + csw * 2;
        if (cslwOffset + 2 > buf.Length) return (String.Empty, []);
        reader = new SpanReader(buf, cslwOffset, buf.Length - cslwOffset);
        var cslw = (Int32)reader.ReadUInt16();
        if (cslw < 1) cslw = 22;
        var fcLcbStart = cslwOffset + 2 + cslw * 4 + 2;
        var fcClxPos = fcLcbStart + 13 * 8;
        if (fcClxPos + 8 > buf.Length) return (String.Empty, []);
        reader = new SpanReader(buf, fcClxPos, 8);
        var fcClx = (Int32)reader.ReadUInt32();
        var lcbClx = (Int32)reader.ReadUInt32();
        if (fcClx < 0 || lcbClx <= 0 || (Int64)fcClx + lcbClx > buf.Length) return (String.Empty, []);

        var clxReader = new SpanReader(buf, fcClx, lcbClx);
        var lcbPlcPcd = 0;
        var foundPcdt = false;
        while (clxReader.Position < clxReader.Capacity)
        {
            var clxt = clxReader.ReadByte();
            if (clxt == 0x01)
            {
                if (clxReader.Position + 2 > clxReader.Capacity) break;
                var cbGrpprl = (Int16)clxReader.ReadUInt16();
                clxReader.Advance(cbGrpprl);
            }
            else if (clxt == 0x02)
            {
                if (clxReader.Position + 4 > clxReader.Capacity) return (String.Empty, []);
                lcbPlcPcd = (Int32)clxReader.ReadUInt32();
                foundPcdt = true;
                break;
            }
            else return (String.Empty, []);
        }
        if (!foundPcdt) return (String.Empty, []);

        var plcPcdStart = fcClx + (Int32)clxReader.Position;
        var plcPcdEnd = plcPcdStart + lcbPlcPcd;
        if (plcPcdEnd > fcClx + lcbClx) plcPcdEnd = fcClx + lcbClx;
        var pieceCount = (lcbPlcPcd - 4) / 12;
        if (pieceCount <= 0) return (String.Empty, []);

        var cpReader = new SpanReader(buf, plcPcdStart, (pieceCount + 1) * 4);
        var pcdReader = new SpanReader(buf, plcPcdStart + (pieceCount + 1) * 4, pieceCount * 8);

        var sb = Pool.StringBuilder.Get();
        var cps = new List<Int32>();
        var cpStart = (Int32)cpReader.ReadUInt32();

        for (var i = 0; i < pieceCount; i++)
        {
            var cpEnd = (Int32)cpReader.ReadUInt32();
            var charCount = cpEnd - cpStart;
            pcdReader.Advance(2);
            var fcCompressed = (Int32)pcdReader.ReadUInt32();
            pcdReader.Advance(2);
            cpStart = cpEnd;
            if (charCount <= 0) continue;

            var fCompressed = ((fcCompressed >> 30) & 1) == 1;
            var fc = fcCompressed & 0x3FFFFFFF;

            if (fCompressed)
            {
                var byteOffset = fc / 2;
                if (byteOffset < 0 || byteOffset + charCount > buf.Length) continue;
                for (var c = 0; c < charCount; c++)
                {
                    var ch = (Char)buf[byteOffset + c];
                    if (AppendDocCharKeepCp(sb, ch)) cps.Add(cpStart - charCount + c);
                }
            }
            else
            {
                var byteOffset = fc;
                if (byteOffset < 0 || byteOffset + charCount * 2 > buf.Length) continue;
                for (var c = 0; c < charCount; c++)
                {
                    var ch = (Char)(buf[byteOffset + c * 2] | (buf[byteOffset + c * 2 + 1] << 8));
                    if (AppendDocCharKeepCp(sb, ch)) cps.Add(cpStart - charCount + c);
                }
            }
        }

        var text = sb.Return(true);
        return (text, cps.ToArray());
    }

    /// <summary>追加字符（保留 \x07 表格标记与 \r 段落符），返回是否产出（域控制符不产出）</summary>
    private static Boolean AppendDocCharKeepCp(StringBuilder sb, Char ch)
    {
        switch (ch)
        {
            case '\x07': // 表格单元格标记：保留
            case '\r':   // 段落结束符：保留
            case '\f':   // 分页符：保留（与 ReadParagraphsWithFormat 一致以 \f 切分）
            case '\v':   // 分栏符：保留
            case '\t':
                sb.Append(ch);
                return true;
            case '\x13': // 域开始/分隔/结束：跳过（不产出，但占 CP）
            case '\x14':
            case '\x15':
                return false;
            default:
                if (ch >= ' ')
                {
                    sb.Append(ch);
                    return true;
                }
                return false;
        }
    }

    #endregion
}
