using System.Text;

namespace NewLife.Office;

/// <summary>Outlook MSG 邮件文件读取器（OLE2/CFB 容器，MAPI 属性）</summary>
/// <remarks>
/// 支持读取 .msg 文件，提取主题、正文（纯文本/HTML）、收发件人地址、附件等。
/// 底层依赖 <see cref="CfbDocument"/> 解析 OLE2 复合文档。
/// <para>用法示例：</para>
/// <code>
/// var reader = new MsgReader();
/// var msg = reader.Read("email.msg");
/// Console.WriteLine(msg.Subject);
/// </code>
/// </remarks>
public class MsgReader
{
    #region MAPI 属性常量

    // 属性 ID（大写十六进制 4 字符）
    private const String PidTagSubject            = "0037";
    private const String PidTagSenderName         = "0C1A";
    private const String PidTagSenderEmailAddr    = "0C1F";
    private const String PidTagSentRepresentingName     = "0042";
    private const String PidTagSentRepresentingEmailAddr= "0065";
    private const String PidTagDisplayTo          = "0E04";
    private const String PidTagDisplayCc          = "0E03";
    private const String PidTagBody              = "1000";
    private const String PidTagBodyHtml          = "1013";

    // 附件属性 ID
    private const String PidTagAttachFilename     = "3704";  // 短文件名（ANSI）
    private const String PidTagAttachLongFilename = "3707";  // 长文件名（Unicode）
    private const String PidTagAttachMimeTag      = "370E";  // MIME 类型
    private const String PidTagAttachDataBin      = "3701";  // 二进制数据

    // 收件人属性 ID
    private const String PidTagRecipientType      = "0C15";  // 1=TO 2=CC 3=BCC
    private const String PidTagRecipientEmailAddr = "3003";
    private const String PidTagRecipientDisplayName = "3001";

    // MAPI 属性类型
    private const String TypeUnicode = "001F";  // PT_UNICODE（UTF-16LE）
    private const String TypeAnsi    = "001E";  // PT_STRING8（Windows-1252）
    private const String TypeBinary  = "0102";  // PT_BINARY
    private const String TypeLong    = "0003";  // PT_LONG（Int32）

    #endregion

    #region 读取方法

    /// <summary>从文件路径读取 MSG</summary>
    /// <param name="path">MSG 文件路径</param>
    /// <returns>解析后的邮件消息</returns>
    public EmlMessage Read(String path)
    {
        using var doc = CfbDocument.Open(path);
        return ParseMsg(doc.Root);
    }

    /// <summary>从流读取 MSG</summary>
    /// <param name="stream">包含 MSG OLE2 内容的可寻址流</param>
    /// <returns>解析后的邮件消息</returns>
    public EmlMessage Read(Stream stream)
    {
        using var doc = CfbDocument.Open(stream, leaveOpen: true);
        return ParseMsg(doc.Root);
    }

    #endregion

    #region 解析核心

    private EmlMessage ParseMsg(CfbStorage root)
    {
        var msg = new EmlMessage();

        msg.Subject  = ReadString(root, PidTagSubject);
        msg.TextBody = ReadString(root, PidTagBody);
        msg.HtmlBody = ReadString(root, PidTagBodyHtml);

        // 发件人：先尝试 SenderEmailAddr，不存在则用 SentRepresentingEmailAddr
        var senderAddr = ReadString(root, PidTagSenderEmailAddr)
                      ?? ReadString(root, PidTagSentRepresentingEmailAddr);
        var senderName = ReadString(root, PidTagSenderName)
                      ?? ReadString(root, PidTagSentRepresentingName);
        msg.From = FormatAddress(senderName, senderAddr);

        // DisplayTo / DisplayCc（分号分隔）→ 解析到 To/Cc
        var displayTo = ReadString(root, PidTagDisplayTo);
        if (!String.IsNullOrEmpty(displayTo))
            foreach (var addr in SplitAddresses(displayTo))
            {
                msg.To.Add(addr);
            }

        var displayCc = ReadString(root, PidTagDisplayCc);
        if (!String.IsNullOrEmpty(displayCc))
            foreach (var addr in SplitAddresses(displayCc))
            {
                msg.Cc.Add(addr);
            }

        // 精确收件人地址（来自 __recip_version1.0_#xxxxx 子存储）
        ParseRecipients(root, msg);

        // 附件（来自 __attach_version1.0_#xxxxx 子存储）
        ParseAttachments(root, msg);

        return msg;
    }

    /// <summary>从根存储读取字符串 MAPI 属性（先尝试 Unicode，再尝试 ANSI）</summary>
    private static String? ReadString(CfbStorage store, String propId)
    {
        // 先尝试 Unicode
        var unicodeStream = store.GetStream($"__substg1.0_{propId}{TypeUnicode}");
        if (unicodeStream != null && unicodeStream.Data.Length > 0)
        {
            // UTF-16LE，移除末尾 null 字符
            var text = Encoding.Unicode.GetString(unicodeStream.Data);
            return text.TrimEnd('\0');
        }

        // 再尝试 ANSI（Windows-1252 降级）
        var ansiStream = store.GetStream($"__substg1.0_{propId}{TypeAnsi}");
        if (ansiStream != null && ansiStream.Data.Length > 0)
        {
            Encoding enc;
            try { enc = Encoding.GetEncoding(1252); }
            catch { enc = Encoding.GetEncoding("iso-8859-1"); }
            var text = enc.GetString(ansiStream.Data);
            return text.TrimEnd('\0');
        }

        return null;
    }

    /// <summary>读取 PT_LONG 属性（直接从属性流 __properties_version1.0 中读取）</summary>
    private static Int32 ReadLongFromPropsStream(CfbStorage store, String propId)
    {
        var propsStream = store.GetStream("__properties_version1.0");
        if (propsStream == null) return 0;

        var data = propsStream.Data;
        // MAPI 属性头为 16 字节（根存储）或 8 字节（子存储）
        var offset = store.Parent == null ? 16 : 8;
        var tag = $"{propId}0003";  // PT_LONG

        while (offset + 16 <= data.Length)
        {
            var propTag = BitConverter.ToString(data, offset, 4).Replace("-", "").ToUpperInvariant();
            var typePart = propTag.Substring(4, 4);
            var idPart   = propTag[..4];

            if (idPart == propId.ToUpperInvariant() && typePart == "0300")
            {
                // 值在 offset+8 的 4 字节
                return BitConverter.IsLittleEndian
                    ? (data[offset + 8] | (data[offset + 9] << 8) | (data[offset + 10] << 16) | (data[offset + 11] << 24))
                    : throw new InvalidOperationException("Big-endian not supported");
            }
            offset += 16;
        }
        return 0;
    }

    private static void ParseRecipients(CfbStorage root, EmlMessage msg)
    {
        // 如果 To 已经有内容（来自 DisplayTo），则不再重复添加
        var alreadyHasTo = msg.To.Count > 0;
        msg.To.Clear();
        msg.Cc.Clear();

        foreach (var storage in root.Storages)
        {
            if (!storage.Name.StartsWith("__recip_version1.0_", StringComparison.OrdinalIgnoreCase))
                continue;

            var email = ReadString(storage, PidTagRecipientEmailAddr);
            var name  = ReadString(storage, PidTagRecipientDisplayName);

            // 读取收件人类型（PT_LONG 直接在 __properties 流中）
            var recipType = ReadLongFromPropsStream(storage, "0C15");

            var addr = FormatAddress(name, email);
            if (String.IsNullOrEmpty(addr)) continue;

            switch (recipType)
            {
                case 2: msg.Cc.Add(addr); break;
                case 3: msg.Bcc.Add(addr); break;
                default: msg.To.Add(addr); break;  // 1 = TO，0 = 未知时也归 TO
            }
        }

        // 如果解析失败（无子存储），恢复 DisplayTo
        if (msg.To.Count == 0 && alreadyHasTo)
        {
            var displayTo = ReadString(root, PidTagDisplayTo);
            if (!String.IsNullOrEmpty(displayTo))
                foreach (var addr in SplitAddresses(displayTo))
                {
                    msg.To.Add(addr);
                }
        }
    }

    private static void ParseAttachments(CfbStorage root, EmlMessage msg)
    {
        foreach (var storage in root.Storages)
        {
            if (!storage.Name.StartsWith("__attach_version1.0_", StringComparison.OrdinalIgnoreCase))
                continue;

            // 附件文件名：优先长文件名，其次短文件名
            var longName  = ReadString(storage, PidTagAttachLongFilename);
            var shortName = ReadString(storage, PidTagAttachFilename);
            var filename  = !String.IsNullOrEmpty(longName) ? longName : shortName;

            var mimeTag  = ReadString(storage, PidTagAttachMimeTag);
            var dataBinStream = storage.GetStream($"__substg1.0_{PidTagAttachDataBin}{TypeBinary}");

            if (dataBinStream == null) continue;

            var attach = new EmlAttachment
            {
                FileName    = filename ?? "attachment",
                ContentType = mimeTag ?? "application/octet-stream",
                Data        = dataBinStream.Data,
            };
            msg.Attachments.Add(attach);
        }
    }

    #endregion

    #region 辅助方法

    private static String? FormatAddress(String? name, String? email)
    {
        if (!String.IsNullOrEmpty(name) && !String.IsNullOrEmpty(email))
            return $"{name} <{email}>";
        if (!String.IsNullOrEmpty(email)) return email;
        if (!String.IsNullOrEmpty(name)) return name;
        return null;
    }

    private static String[] SplitAddresses(String addresses) =>
        addresses.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);

    #endregion
}
