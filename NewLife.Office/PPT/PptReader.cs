using System.Text;
using NewLife.Office;

namespace NewLife.Office;

/// <summary>PowerPoint 97-2003 二进制（.ppt）演示文稿读取器</summary>
/// <remarks>
/// 基于 OLE2/CFB 容器解析 MS-PPT 格式，从 "PowerPoint Document" 流中提取各幻灯片文本。
/// 通过递归扫描 PPT 记录树，定位 SlideContainer，并收集其中的文本原子记录。
/// <para>用法示例：</para>
/// <code>
/// using var reader = new PptReader("slides.ppt");
/// for (var i = 0; i &lt; reader.SlideCount; i++)
///     Console.WriteLine(reader.GetSlideText(i));
/// </code>
/// </remarks>
public sealed class PptReader : IDisposable
{
    #region 属性

    /// <summary>幻灯片数量</summary>
    public Int32 SlideCount => _slideTexts.Count;

    private Boolean _disposed;

    #endregion

    #region 私有字段

    private readonly List<String> _slideTexts = [];

    #endregion

    #region 构造

    /// <summary>从 ppt 文件路径打开</summary>
    /// <param name="path">ppt 文件路径</param>
    public PptReader(String path)
    {
        using var doc = CfbDocument.Open(path);
        var data = GetPptStream(doc);
        ParseStream(data);
    }

    /// <summary>从流打开（需包含 ppt 的完整 OLE2 容器内容）</summary>
    /// <param name="stream">可读流</param>
    public PptReader(Stream stream)
    {
        using var doc = CfbDocument.Open(stream, leaveOpen: true);
        var data = GetPptStream(doc);
        ParseStream(data);
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

    private static Byte[] GetPptStream(CfbDocument doc)
    {
        var data = doc.GetStreamData("PowerPoint Document");
        if (data == null || data.Length == 0)
            throw new InvalidDataException("找不到 PowerPoint Document 流，文件可能不是有效的 .ppt 格式。");
        return data;
    }

    #endregion

    #region 读取方法

    /// <summary>获取指定幻灯片的全部文本</summary>
    /// <param name="index">幻灯片索引（从 0 开始）</param>
    /// <returns>幻灯片文本</returns>
    public String GetSlideText(Int32 index)
    {
        if (index < 0 || index >= _slideTexts.Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"幻灯片索引 {index} 超出范围（共 {_slideTexts.Count} 张）。");
        return _slideTexts[index];
    }

    /// <summary>依次读取所有幻灯片文本</summary>
    /// <returns>幻灯片文本序列</returns>
    public IEnumerable<String> ReadAllSlides() => _slideTexts;

    #endregion

    #region PPT 流解析

    /// <summary>解析 PowerPoint Document 流，提取所有幻灯片文本</summary>
    /// <param name="buf">PPT 流字节</param>
    private void ParseStream(Byte[] buf)
    {
        var slides = new List<List<String>>();
        ScanRecords(buf, 0, buf.Length, slides, null);

        foreach (var ts in slides)
        {
            _slideTexts.Add(String.Join("\n", ts));
        }
    }

    /// <summary>递归扫描 PPT 记录树，收集幻灯片文本</summary>
    /// <param name="buf">字节流</param>
    /// <param name="start">扫描起点</param>
    /// <param name="end">扫描终点</param>
    /// <param name="slides">累积的幻灯片文本列表</param>
    /// <param name="currentSlide">当前幻灯片文本收集器（null 代表还未进入任何幻灯片）</param>
    private static void ScanRecords(Byte[] buf, Int32 start, Int32 end,
        List<List<String>> slides, List<String> currentSlide)
    {
        var pos = start;
        while (pos + 8 <= end)
        {
            var verType = ReadUInt16(buf, pos);
            var recType = ReadUInt16(buf, pos + 2);
            var recLen = (Int32)ReadUInt32(buf, pos + 4);
            var recVer = verType & 0x0F;
            pos += 8;

            if (recLen < 0 || pos + recLen > end) break;

            if (recType == RecTextCharsAtom && recLen >= 2)
            {
                // UTF-16LE 文本：每字符 2 字节，长度需按 2 对齐
                var charBytes = recLen & ~1;
                var text = Encoding.Unicode.GetString(buf, pos, charBytes).TrimEnd('\r', '\n');
                if (text.Length > 0)
                    (currentSlide ?? GetOrAddSlide(slides))?.Add(text);
            }
            else if (recType == RecTextBytesAtom && recLen >= 1)
            {
                // ANSI 文本：直接字节→字符映射（ISO-8859-1）
                var text = DecodeLatin1(buf, pos, recLen).TrimEnd('\r', '\n');
                if (text.Length > 0)
                    (currentSlide ?? GetOrAddSlide(slides))?.Add(text);
            }
            else if (recVer == 0x0F)
            {
                // 容器记录
                if (recType == RecSlideContainer)
                {
                    // 进入一个新幻灯片
                    var slideTexts = new List<String>();
                    slides.Add(slideTexts);
                    ScanRecords(buf, pos, pos + recLen, slides, slideTexts);
                }
                else
                {
                    // 其他容器—继续在当前幻灯片上下文中递归
                    ScanRecords(buf, pos, pos + recLen, slides, currentSlide);
                }
            }

            pos += recLen;
        }
    }

    /// <summary>确保 slides 中至少有一个 slide，并返回最后一个</summary>
    /// <param name="slides">幻灯片文本列表</param>
    /// <returns>最后一个幻灯片的文本收集器</returns>
    private static List<String> GetOrAddSlide(List<List<String>> slides)
    {
        if (slides.Count == 0) slides.Add([]);
        return slides[slides.Count - 1];
    }

    /// <summary>ISO-8859-1 字节→字符映射</summary>
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

    #region 字节工具

    private static UInt16 ReadUInt16(Byte[] buf, Int32 pos) =>
        (UInt16)(buf[pos] | (buf[pos + 1] << 8));

    private static UInt32 ReadUInt32(Byte[] buf, Int32 pos) =>
        (UInt32)(buf[pos] | (buf[pos + 1] << 8) | (buf[pos + 2] << 16) | (buf[pos + 3] << 24));

    #endregion

    #region PPT 记录类型常量

    // 幻灯片容器
    private const UInt16 RecSlideContainer = 0x03EE;
    // 文本原子（UTF-16LE）
    private const UInt16 RecTextCharsAtom = 0x03F2;
    // 文本原子（ANSI/Latin-1）
    private const UInt16 RecTextBytesAtom = 0x03F0;

    #endregion
}
