using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office;

/// <summary>vCard 联系人文件读取器（RFC 6350）</summary>
/// <remarks>
/// 支持 vCard 2.1、3.0、4.0 格式。
/// 一个 .vcf 文件可包含多个联系人（BEGIN:VCARD ... END:VCARD）。
/// </remarks>
public class VCardReader
{
    #region 读取方法

    /// <summary>从文件读取所有联系人</summary>
    /// <param name="path">vCard 文件路径（.vcf）</param>
    /// <returns>联系人列表</returns>
    public List<VCardContact> ReadAll(String path)
    {
        var text = File.ReadAllText(path, Encoding.UTF8);
        return ParseAll(text);
    }

    /// <summary>从文件读取第一个联系人</summary>
    /// <param name="path">vCard 文件路径</param>
    /// <returns>联系人，若文件为空则返回 null</returns>
    public VCardContact? Read(String path)
    {
        var all = ReadAll(path);
        return all.Count > 0 ? all[0] : null;
    }

    /// <summary>从流读取所有联系人</summary>
    /// <param name="stream">可读流</param>
    /// <returns>联系人列表</returns>
    public List<VCardContact> ReadAll(Stream stream)
    {
        using var sr = new StreamReader(stream, Encoding.UTF8);
        return ParseAll(sr.ReadToEnd());
    }

    /// <summary>从字符串解析所有联系人</summary>
    /// <param name="text">vCard 文本</param>
    /// <returns>联系人列表</returns>
    public List<VCardContact> ParseAll(String text)
    {
        var contacts = new List<VCardContact>();
        var lines = UnfoldLines(text);
        VCardContact? current = null;

        foreach (var line in lines)
        {
            if (String.IsNullOrWhiteSpace(line)) continue;
            var colonIdx = line.IndexOf(':');
            if (colonIdx <= 0) continue;
            var left = line.Substring(0, colonIdx).Trim();
            var value = line.Substring(colonIdx + 1).Trim();

            // 拆分属性名和参数（; 分隔）
            var parts = left.Split(';');
            var propName = parts[0].ToUpperInvariant();
            var paramStr = parts.Length > 1 ? String.Join(";", parts, 1, parts.Length - 1) : String.Empty;

            switch (propName)
            {
                case "BEGIN":
                    if (value.Equals("VCARD", StringComparison.OrdinalIgnoreCase))
                        current = new VCardContact();
                    break;
                case "END":
                    if (value.Equals("VCARD", StringComparison.OrdinalIgnoreCase) && current != null)
                    {
                        contacts.Add(current);
                        current = null;
                    }
                    break;
                default:
                    if (current != null)
                        ApplyProp(current, propName, paramStr, value);
                    break;
            }
        }

        return contacts;
    }

    #endregion

    #region 私有方法

    private static List<String> UnfoldLines(String text)
    {
        var lines = new List<String>();
        var sb = new StringBuilder();
        foreach (var line in text.Split('\n'))
        {
            var l = line.TrimEnd('\r');
            if (l.Length > 0 && (l[0] == ' ' || l[0] == '\t'))
                sb.Append(l.Substring(1));
            else
            {
                if (sb.Length > 0) lines.Add(sb.ToString());
                sb.Clear();
                sb.Append(l);
            }
        }
        if (sb.Length > 0) lines.Add(sb.ToString());
        return lines;
    }

    private static void ApplyProp(VCardContact c, String name, String param, String value)
    {
        var typeVal = ExtractParam(param, "TYPE");
        switch (name)
        {
            case "FN":
                c.FullName = UnescapeText(DecodeValue(value, param));
                break;
            case "N":
                c.Name = ParseName(UnescapeText(value));
                break;
            case "ORG":
                c.Organization = UnescapeText(value).Split(';')[0];
                break;
            case "TITLE":
                c.Title = UnescapeText(value);
                break;
            case "NOTE":
                c.Note = UnescapeText(DecodeValue(value, param));
                break;
            case "URL":
                c.Url = value;
                break;
            case "UID":
                c.Uid = value;
                break;
            case "BDAY":
                c.Birthday = ParseDate(value);
                break;
            case "REV":
                if (DateTimeOffset.TryParse(value, out var rev)) c.Revision = rev;
                break;
            case "PHOTO":
                c.Photo = value;
                break;
            case "TEL":
                c.Phones.Add(new VCardPhone { Number = value.Trim(), Type = typeVal });
                break;
            case "EMAIL":
                c.Emails.Add(new VCardEmail { Address = value.Trim(), Type = typeVal });
                break;
            case "ADR":
                c.Addresses.Add(ParseAddress(UnescapeText(value), typeVal));
                break;
            default:
                c.ExtraProps[name.ToLowerInvariant()] = value;
                break;
        }
    }

    private static String? ExtractParam(String param, String key)
    {
        // 查找 TYPE=xxx 或 TYPE="xxx"
        var m = Regex.Match(param, key + @"=([^;""]+|""[^""]*"")", RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value.Trim('"') : null;
    }

    private static VCardName ParseName(String value)
    {
        var parts = value.Split(';');
        return new VCardName
        {
            Family     = parts.Length > 0 ? parts[0] : null,
            Given      = parts.Length > 1 ? parts[1] : null,
            Additional = parts.Length > 2 ? parts[2] : null,
            Prefix     = parts.Length > 3 ? parts[3] : null,
            Suffix     = parts.Length > 4 ? parts[4] : null,
        };
    }

    private static VCardAddress ParseAddress(String value, String? type)
    {
        var parts = value.Split(';');
        return new VCardAddress
        {
            PoBox      = parts.Length > 0 ? parts[0] : null,
            Extended   = parts.Length > 1 ? parts[1] : null,
            Street     = parts.Length > 2 ? parts[2] : null,
            City       = parts.Length > 3 ? parts[3] : null,
            Region     = parts.Length > 4 ? parts[4] : null,
            PostalCode = parts.Length > 5 ? parts[5] : null,
            Country    = parts.Length > 6 ? parts[6] : null,
            Type       = type,
        };
    }

    private static DateTime? ParseDate(String value)
    {
        // YYYYMMDD or YYYY-MM-DD
        var clean = value.Replace("-", "").Split('T')[0];
        if (clean.Length == 8 && DateTime.TryParseExact(clean, "yyyyMMdd",
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var d))
            return d;
        return null;
    }

    private static String DecodeValue(String value, String param)
    {
        if (param.IndexOf("ENCODING=QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase) >= 0 ||
            param.IndexOf("ENCODING=QP", StringComparison.OrdinalIgnoreCase) >= 0)
            return DecodeQuotedPrintable(value);
        if (param.IndexOf("ENCODING=BASE64", StringComparison.OrdinalIgnoreCase) >= 0 ||
            param.IndexOf("ENCODING=B", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            try { return Encoding.UTF8.GetString(Convert.FromBase64String(Regex.Replace(value, @"\s+", ""))); }
            catch { return value; }
        }
        return value;
    }

    private static String DecodeQuotedPrintable(String input)
    {
        var ms = new MemoryStream();
        var i = 0;
        while (i < input.Length)
        {
            if (input[i] == '=' && i + 2 < input.Length)
            {
                if (Byte.TryParse(input.Substring(i + 1, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
                {
                    ms.WriteByte(b);
                    i += 3;
                }
                else ms.WriteByte((Byte)input[i++]);
            }
            else ms.WriteByte((Byte)input[i++]);
        }
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    private static String UnescapeText(String value)
    {
        return value.Replace("\\n", "\n").Replace("\\N", "\n")
                    .Replace("\\,", ",").Replace("\\;", ";")
                    .Replace("\\\\", "\\");
    }

    #endregion
}
