using System.Text;

namespace NewLife.Office;

/// <summary>VCard 文档包装，封装联系人列表并提供文本/Markdown 提取能力</summary>
public class VCardDocument : ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>联系人列表</summary>
    public List<VCardContact> Contacts { get; }
    #endregion

    #region 构造
    /// <summary>实例化 VCard 文档包装</summary>
    /// <param name="contacts">联系人列表</param>
    public VCardDocument(List<VCardContact> contacts) => Contacts = contacts ?? [];
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（每个联系人按字段输出）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText()
    {
        if (Contacts == null || Contacts.Count == 0) return null;

        var sb = new StringBuilder();
        for (var i = 0; i < Contacts.Count; i++)
        {
            if (i > 0) sb.AppendLine();
            var c = Contacts[i];
            if (!String.IsNullOrEmpty(c.FullName)) sb.AppendLine(c.FullName);
            if (!String.IsNullOrEmpty(c.Organization)) sb.AppendLine($"组织: {c.Organization}");
            if (!String.IsNullOrEmpty(c.Title)) sb.AppendLine($"职位: {c.Title}");
            if (c.Birthday != null) sb.AppendLine($"生日: {c.Birthday:yyyy-MM-dd}");
            foreach (var phone in c.Phones)
            {
                if (!String.IsNullOrEmpty(phone.Number))
                    sb.AppendLine($"电话: {phone.Number}" + (String.IsNullOrEmpty(phone.Type) ? "" : $" ({phone.Type})"));
            }
            foreach (var email in c.Emails)
            {
                if (!String.IsNullOrEmpty(email.Address))
                    sb.AppendLine($"邮箱: {email.Address}" + (String.IsNullOrEmpty(email.Type) ? "" : $" ({email.Type})"));
            }
            foreach (var addr in c.Addresses)
            {
                var parts = new List<String>();
                if (!String.IsNullOrEmpty(addr.Country)) parts.Add(addr.Country!);
                if (!String.IsNullOrEmpty(addr.Region)) parts.Add(addr.Region!);
                if (!String.IsNullOrEmpty(addr.City)) parts.Add(addr.City!);
                if (!String.IsNullOrEmpty(addr.Street)) parts.Add(addr.Street!);
                if (parts.Count > 0)
                    sb.AppendLine($"地址: {String.Join(" ", parts)}" + (String.IsNullOrEmpty(addr.Type) ? "" : $" ({addr.Type})"));
            }
            if (!String.IsNullOrEmpty(c.Url)) sb.AppendLine($"网址: {c.Url}");
            if (!String.IsNullOrEmpty(c.Note)) sb.AppendLine($"备注: {c.Note}");
        }
        return sb.ToString();
    }

    /// <summary>提取 Markdown 格式（每个联系人用标题分隔）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown()
    {
        if (Contacts == null || Contacts.Count == 0) return null;

        var sb = new StringBuilder();
        for (var i = 0; i < Contacts.Count; i++)
        {
            if (i > 0) sb.AppendLine();
            var c = Contacts[i];
            sb.AppendLine($"## {c.FullName ?? "未知联系人"}");
            sb.AppendLine();
            if (!String.IsNullOrEmpty(c.Organization)) sb.AppendLine($"- **组织**: {c.Organization}");
            if (!String.IsNullOrEmpty(c.Title)) sb.AppendLine($"- **职位**: {c.Title}");
            if (c.Birthday != null) sb.AppendLine($"- **生日**: {c.Birthday:yyyy-MM-dd}");
            foreach (var phone in c.Phones)
            {
                if (!String.IsNullOrEmpty(phone.Number))
                    sb.AppendLine($"- **电话**: {phone.Number}" + (String.IsNullOrEmpty(phone.Type) ? "" : $" ({phone.Type})"));
            }
            foreach (var email in c.Emails)
            {
                if (!String.IsNullOrEmpty(email.Address))
                    sb.AppendLine($"- **邮箱**: {email.Address}" + (String.IsNullOrEmpty(email.Type) ? "" : $" ({email.Type})"));
            }
            foreach (var addr in c.Addresses)
            {
                var parts = new List<String>();
                if (!String.IsNullOrEmpty(addr.Country)) parts.Add(addr.Country!);
                if (!String.IsNullOrEmpty(addr.Region)) parts.Add(addr.Region!);
                if (!String.IsNullOrEmpty(addr.City)) parts.Add(addr.City!);
                if (!String.IsNullOrEmpty(addr.Street)) parts.Add(addr.Street!);
                if (parts.Count > 0)
                    sb.AppendLine($"- **地址**: {String.Join(" ", parts)}" + (String.IsNullOrEmpty(addr.Type) ? "" : $" ({addr.Type})"));
            }
            if (!String.IsNullOrEmpty(c.Url)) sb.AppendLine($"- **网址**: {c.Url}");
            if (!String.IsNullOrEmpty(c.Note)) sb.AppendLine($"- **备注**: {c.Note}");
        }
        return sb.ToString();
    }
    #endregion
}
