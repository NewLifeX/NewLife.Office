using System.Text;
using NewLife.Buffers;
using NewLife.Collections;
using NewLife.Office;

namespace NewLife.Office;

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

    #endregion

    #region 构造

    /// <summary>从 doc 文件路径打开</summary>
    /// <param name="path">doc 文件路径</param>
    public DocReader(String path)
    {
        using var doc = CfbDocument.Open(path);
        _wordDoc = GetWordDocStream(doc);
        ValidateFib(_wordDoc);
    }

    /// <summary>从流打开（需包含 doc 的完整 OLE2 容器内容）</summary>
    /// <param name="stream">可读流</param>
    public DocReader(Stream stream)
    {
        using var doc = CfbDocument.Open(stream, leaveOpen: true);
        _wordDoc = GetWordDocStream(doc);
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
}
