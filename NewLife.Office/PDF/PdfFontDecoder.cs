using System.Text;

namespace NewLife.Office.Pdf;

/// <summary>PDF 字体解码器。解析 Type0（CID）/简单字体的编码映射，把内容流中的字形编码还原为 Unicode</summary>
/// <remarks>
/// <para>PDF 内容流里的文本是<b>字形编码</b>而不是 Unicode：Type0 字体常见 Identity-H 编码（2 字节 CID），
/// 必须经 /ToUnicode CMap 才能还原；简单字体则为单字节编码 + WinAnsi/MacRoman/Differences。</para>
/// <para>本类只做文本还原所需的最小解析：codespacerange 决定编码字节宽度，bfchar/bfrange 建立映射。
/// 缺少 /ToUnicode 时按 Identity + UTF-16BE 兜底（部分生成器直接写入 Unicode 码点）。</para>
/// </remarks>
internal sealed class PdfFontDecoder
{
    /// <summary>字体基础名</summary>
    public String? BaseFont { get; set; }

    /// <summary>编码字节宽度是否为 2 字节（Type0/CID 字体）</summary>
    public Boolean TwoByte { get; set; }

    /// <summary>编码到 Unicode 文本的映射</summary>
    public Dictionary<Int32, String> Map { get; } = [];

    /// <summary>解码字形编码为文本</summary>
    /// <param name="bytes">内容流中的字符串字节</param>
    /// <returns>还原后的文本</returns>
    public String Decode(Byte[] bytes)
    {
        if (bytes == null || bytes.Length == 0) return String.Empty;

        var sb = new StringBuilder(bytes.Length);
        if (TwoByte)
        {
            for (var i = 0; i + 1 < bytes.Length; i += 2)
            {
                var code = (bytes[i] << 8) | bytes[i + 1];

                // 有映射优先取映射；无映射（无 ToUnicode）时按 Identity 约定当作 Unicode 码点
                if (Map.TryGetValue(code, out var text))
                    sb.Append(text);
                else if (Map.Count == 0)
                    sb.Append((Char)code);
            }

            return sb.ToString();
        }

        foreach (var b in bytes)
        {
            if (Map.TryGetValue(b, out var text))
                sb.Append(text);
            else
                sb.Append((Char)b);
        }

        return sb.ToString();
    }
}

/// <summary>PDF 字体表加载。从页面 /Resources /Font 构建字体解码器</summary>
internal static class PdfFonts
{
    /// <summary>单个字体映射的最大条目数。畸形 CMap 声明超大区间时按此截断</summary>
    internal const Int32 MaxMapEntries = 65536;

    /// <summary>加载页面的字体解码表</summary>
    /// <param name="data">PDF 原始字节</param>
    /// <param name="xref">交叉引用表</param>
    /// <param name="resources">页面资源字典（可空）</param>
    /// <returns>字体名（不含 /）到解码器的映射</returns>
    public static Dictionary<String, PdfFontDecoder> Load(Byte[] data, PdfXRefTable xref, PdfDict? resources)
    {
        var fonts = new Dictionary<String, PdfFontDecoder>(StringComparer.Ordinal);

        if (resources == null || !resources.TryGetValue("Font", out var fontVal)) return fonts;
        if (Resolve(data, xref, fontVal) is not DictObj fontDict) return fonts;

        foreach (var kv in fontDict.Value)
        {
            if (Resolve(data, xref, kv.Value) is not DictObj obj) continue;

            var font = Parse(data, xref, obj.Value);
            if (font != null) fonts[kv.Key] = font;
        }

        return fonts;
    }

    /// <summary>解析单个字体字典</summary>
    private static PdfFontDecoder? Parse(Byte[] data, PdfXRefTable xref, PdfDict dict)
    {
        var font = new PdfFontDecoder();
        if (dict.TryGetValue("BaseFont", out var baseFont) && baseFont is PdfName bn) font.BaseFont = bn.Value;

        var subtype = dict.TryGetValue("Subtype", out var st) && st is PdfName stn ? stn.Value : null;
        if (subtype == "Type0") font.TwoByte = true;

        // 编码：Identity-H/V 为 2 字节；CMap 流按 codespacerange 判定
        if (dict.TryGetValue("Encoding", out var encVal))
        {
            var enc = Resolve(data, xref, encVal);
            if (enc is PdfName en)
            {
                if (en.Value.StartsWith("Identity", StringComparison.Ordinal)) font.TwoByte = true;
            }
            else if (enc is PdfStream encStream)
            {
                var cmap = ReadStreamText(encStream);
                if (cmap != null && GetCodeByteWidth(cmap) >= 2) font.TwoByte = true;
            }
        }

        // ToUnicode 映射
        if (dict.TryGetValue("ToUnicode", out var tuVal))
        {
            var tu = Resolve(data, xref, tuVal);
            if (tu is PdfStream tuStream)
            {
                var cmap = ReadStreamText(tuStream);
                if (!cmap.IsNullOrEmpty()) ParseCMapToUnicode(cmap!, font.Map, out var width);
            }
        }

        return font;
    }

    /// <summary>读取流内容为文本（自动解压）</summary>
    private static String? ReadStreamText(PdfStream stream)
    {
        try
        {
            var bytes = PdfReader.DecompressStreamData(stream.Data, stream.Dict);
            return bytes.Length == 0 ? null : Encoding.GetEncoding(28591).GetString(bytes);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>解析 /ToUnicode CMap：codespacerange 决定编码宽度，bfchar/bfrange 建立映射</summary>
    /// <param name="cmap">CMap 文本</param>
    /// <param name="map">输出映射</param>
    /// <param name="codeWidth">输出编码字节宽度</param>
    internal static void ParseCMapToUnicode(String cmap, Dictionary<Int32, String> map, out Int32 codeWidth)
    {
        codeWidth = 1;
        var tokens = Tokenize(cmap);

        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (token == "begincodespacerange")
            {
                // <lo> <hi> 对，取低值长度作为编码宽度
                if (i + 1 < tokens.Count && tokens[i + 1].Length > 0 && tokens[i + 1][0] == '<')
                    codeWidth = Math.Max(1, (tokens[i + 1].Length - 2) / 2);
            }
            else if (token == "beginbfchar")
            {
                i++;
                for (; i + 1 < tokens.Count && !tokens[i].StartsWith("end"); i += 2)
                {
                    if (!IsHexToken(tokens[i]) || !IsHexToken(tokens[i + 1])) break;
                    if (map.Count >= MaxMapEntries) break;

                    var code = ParseHexInt(tokens[i]);
                    var text = DecodeUtf16(tokens[i + 1]);
                    if (code >= 0) map[code] = text;
                }
                i--;
            }
            else if (token == "beginbfrange")
            {
                i++;
                for (; i < tokens.Count && !tokens[i].StartsWith("end");)
                {
                    if (i + 2 >= tokens.Count || !IsHexToken(tokens[i]) || !IsHexToken(tokens[i + 1])) break;

                    var lo = ParseHexInt(tokens[i]);
                    var hi = ParseHexInt(tokens[i + 1]);
                    i += 2;

                    // 上限保护：畸形 CMap 可能声明超大区间，限制单次展开规模
                    if (hi < lo) continue;
                    if (hi - lo >= MaxMapEntries) hi = lo + MaxMapEntries - 1;

                    if (i < tokens.Count && tokens[i] == "[")
                    {
                        // <lo> <hi> [ <d1> <d2> ... ]
                        i++;
                        var code = lo;
                        for (; i < tokens.Count && tokens[i] != "]"; i++)
                        {
                            if (tokens[i].Length > 0 && tokens[i][0] == '<' && code <= hi)
                                map[code++] = DecodeUtf16(tokens[i]);
                        }
                        if (i < tokens.Count && tokens[i] == "]") i++;
                    }
                    else if (i < tokens.Count && tokens[i].Length > 0 && tokens[i][0] == '<')
                    {
                        // <lo> <hi> <dst>：目标码点按顺序递增
                        var dst = ParseHexBytes(tokens[i]);
                        i++;
                        for (var code = lo; code <= hi; code++)
                        {
                            map[code] = DecodeUtf16(IncrementHex(dst, code - lo));
                        }
                    }
                    else
                        break;
                }
                i--;
            }
        }
    }

    /// <summary>获取 CMap 的编码字节宽度；未声明时返回 0</summary>
    private static Int32 GetCodeByteWidth(String cmap)
    {
        ParseCMapToUnicode(cmap, [], out var width);
        return width;
    }

    /// <summary>把 CMap 文本切分为十六进制串、数组括号与关键字</summary>
    private static List<String> Tokenize(String text)
    {
        var tokens = new List<String>();
        var i = 0;
        while (i < text.Length)
        {
            var c = text[i];
            if (Char.IsWhiteSpace(c))
            {
                i++;
                continue;
            }

            if (c == '<')
            {
                var end = text.IndexOf('>', i + 1);
                if (end < 0) break;
                tokens.Add(text[i..(end + 1)]);
                i = end + 1;
                continue;
            }

            if (c == '[' || c == ']')
            {
                tokens.Add(c.ToString());
                i++;
                continue;
            }

            if (c == '%')
            {
                // 注释到行尾
                var nl = text.IndexOf('\n', i);
                if (nl < 0) break;
                i = nl + 1;
                continue;
            }

            var start = i;
            while (i < text.Length && !Char.IsWhiteSpace(text[i]) && text[i] != '<' && text[i] != '[' && text[i] != ']') i++;
            tokens.Add(text[start..i]);
        }

        return tokens;
    }

    /// <summary>判断词元是否为 &lt;hex&gt; 形式</summary>
    private static Boolean IsHexToken(String token) => token.Length > 0 && token[0] == '<';

    /// <summary>解析 &lt;hex&gt; 为整数</summary>
    private static Int32 ParseHexInt(String token)
    {
        var bytes = ParseHexBytes(token);
        var value = 0;
        foreach (var b in bytes) value = (value << 8) | b;

        return value;
    }

    /// <summary>把 &lt;hex&gt; 解析为字节数组</summary>
    private static Byte[] ParseHexBytes(String token)
    {
        var hex = token.Trim('<', '>');
        if (hex.Length % 2 == 1) hex += "0";

        var bytes = new Byte[hex.Length / 2];
        for (var i = 0; i < bytes.Length; i++)
        {
            if (!Byte.TryParse(hex.Substring(i * 2, 2), System.Globalization.NumberStyles.HexNumber, null, out var b)) return [];
            bytes[i] = b;
        }

        return bytes;
    }

    /// <summary>把 &lt;hex&gt; 目的串解码为 UTF-16BE 文本</summary>
    private static String DecodeUtf16(String token) => DecodeUtf16(ParseHexBytes(token));

    private static String DecodeUtf16(Byte[] bytes)
    {
        if (bytes.Length == 0) return String.Empty;
        if (bytes.Length == 1) return ((Char)bytes[0]).ToString();

        var sb = new StringBuilder(bytes.Length / 2);
        for (var i = 0; i + 1 < bytes.Length; i += 2)
            sb.Append((Char)((bytes[i] << 8) | bytes[i + 1]));

        return sb.ToString();
    }

    /// <summary>对十六进制目的串做增量（bfrange 目标码点递增）</summary>
    private static String IncrementHex(Byte[] bytes, Int32 offset)
    {
        if (offset == 0 || bytes.Length < 2) return ByteToHex(bytes);

        var copy = (Byte[])bytes.Clone();
        var value = (copy[^2] << 8) | copy[^1];
        value += offset;
        copy[^2] = (Byte)(value >> 8);
        copy[^1] = (Byte)value;

        return ByteToHex(copy);
    }

    private static String ByteToHex(Byte[] bytes)
    {
        var sb = new StringBuilder(bytes.Length * 2 + 2);
        sb.Append('<');
        foreach (var b in bytes) sb.Append(b.ToString("X2"));
        sb.Append('>');

        return sb.ToString();
    }

    /// <summary>解析间接引用</summary>
    private static PdfObject? Resolve(Byte[] data, PdfXRefTable xref, PdfObject? obj)
        => obj is Ref r ? PdfObjectParser.ReadObject(data, xref, r.ObjNum) : obj;
}
