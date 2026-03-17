using System.Text;

namespace NewLife.Office;

/// <summary>vCard 联系人文件写入器（RFC 6350）</summary>
/// <remarks>
/// 默认生成 vCard 4.0 格式。
/// 行长度超过 75 字节时自动折行（CRLF + 空格）。
/// </remarks>
public class VCardWriter
{
    #region 属性

    /// <summary>vCard 版本，默认 4.0</summary>
    public String Version { get; set; } = "4.0";

    #endregion

    #region 写入方法

    /// <summary>将联系人写入文件</summary>
    /// <param name="contact">联系人</param>
    /// <param name="path">输出文件路径（.vcf）</param>
    public void Write(VCardContact contact, String path)
    {
        var content = Build(contact);
        File.WriteAllText(path, content, new UTF8Encoding(false));
    }

    /// <summary>将多个联系人写入文件</summary>
    /// <param name="contacts">联系人列表</param>
    /// <param name="path">输出文件路径（.vcf）</param>
    public void WriteAll(IEnumerable<VCardContact> contacts, String path)
    {
        var sb = new StringBuilder();
        foreach (var c in contacts)
        {
            sb.Append(Build(c));
        }
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
    }

    /// <summary>将联系人写入流</summary>
    /// <param name="contact">联系人</param>
    /// <param name="stream">可写输出流</param>
    public void Write(VCardContact contact, Stream stream)
    {
        var content = Build(contact);
        var bytes = new UTF8Encoding(false).GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>将联系人序列化为 vCard 字符串</summary>
    /// <param name="contact">联系人</param>
    /// <returns>vCard 格式字符串</returns>
    public String Build(VCardContact contact)
    {
        var sb = new StringBuilder();
        AppendLine(sb, "BEGIN:VCARD");
        AppendLine(sb, "VERSION:" + Version);

        if (contact.Uid != null)
            AppendLine(sb, "UID:" + contact.Uid);

        if (contact.FullName != null)
            AppendLine(sb, "FN:" + EscapeText(contact.FullName));

        if (contact.Name != null)
        {
            var n = contact.Name;
            var nameStr = $"{n.Family ?? ""};{n.Given ?? ""};{n.Additional ?? ""};{n.Prefix ?? ""};{n.Suffix ?? ""}";
            AppendLine(sb, "N:" + EscapeText(nameStr));
        }

        if (contact.Organization != null)
            AppendLine(sb, "ORG:" + EscapeText(contact.Organization));

        if (contact.Title != null)
            AppendLine(sb, "TITLE:" + EscapeText(contact.Title));

        if (contact.Birthday.HasValue)
            AppendLine(sb, "BDAY:" + contact.Birthday.Value.ToString("yyyyMMdd"));

        if (contact.Note != null)
            AppendLine(sb, "NOTE:" + EscapeText(contact.Note));

        if (contact.Url != null)
            AppendLine(sb, "URL:" + contact.Url);

        if (contact.Photo != null)
            AppendLine(sb, "PHOTO:" + contact.Photo);

        foreach (var phone in contact.Phones)
        {
            var prop = phone.Type != null ? $"TEL;TYPE={phone.Type}" : "TEL";
            AppendLine(sb, prop + ":" + (phone.Number ?? ""));
        }

        foreach (var email in contact.Emails)
        {
            var prop = email.Type != null ? $"EMAIL;TYPE={email.Type}" : "EMAIL";
            AppendLine(sb, prop + ":" + (email.Address ?? ""));
        }

        foreach (var adr in contact.Addresses)
        {
            var prop = adr.Type != null ? $"ADR;TYPE={adr.Type}" : "ADR";
            var adrStr = $"{adr.PoBox ?? ""};{adr.Extended ?? ""};{adr.Street ?? ""}" +
                         $";{adr.City ?? ""};{adr.Region ?? ""};{adr.PostalCode ?? ""};{adr.Country ?? ""}";
            AppendLine(sb, prop + ":" + EscapeText(adrStr));
        }

        if (contact.Revision.HasValue)
            AppendLine(sb, "REV:" + contact.Revision.Value.ToUniversalTime().ToString("yyyyMMddTHHmmssZ"));

        foreach (var kv in contact.ExtraProps)
        {
            AppendLine(sb, kv.Key.ToUpperInvariant() + ":" + kv.Value);
        }

        AppendLine(sb, "END:VCARD");
        return sb.ToString();
    }

    #endregion

    #region 私有方法

    /// <summary>添加折行支持（RFC 6350 要求内容行最长 75 字节）</summary>
    private static void AppendLine(StringBuilder sb, String line)
    {
        var bytes = Encoding.UTF8.GetBytes(line);
        if (bytes.Length <= 75)
        {
            sb.Append(line).Append("\r\n");
            return;
        }

        var pos = 0;
        var first = true;
        while (pos < line.Length)
        {
            if (!first) sb.Append(' ');
            var take = TakeChars(line, pos, first ? 75 : 74);
            sb.Append(line, pos, take);
            sb.Append("\r\n");
            pos += take;
            first = false;
        }
    }

    private static Int32 TakeChars(String text, Int32 pos, Int32 maxBytes)
    {
        var bytes = 0;
        var count = 0;
        while (pos + count < text.Length)
        {
            var c = text[pos + count];
            var cb = Encoding.UTF8.GetByteCount(new Char[] { c });
            if (bytes + cb > maxBytes) break;
            bytes += cb;
            count++;
        }
        return count > 0 ? count : 1;
    }

    private static String EscapeText(String value)
    {
        return value.Replace("\\", "\\\\").Replace(";", "\\;")
                    .Replace(",", "\\,").Replace("\n", "\\n").Replace("\r", "");
    }

    #endregion
}
