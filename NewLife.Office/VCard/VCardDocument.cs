using System.Text;

namespace NewLife.Office.VCard;

/// <summary>VCard 文档包装，封装联系人列表并提供文本/Markdown 提取能力</summary>
public class VCardDocument : ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>联系人列表</summary>
    public List<Contact> Contacts { get; }
    #endregion

    #region 构造
    /// <summary>实例化 VCard 文档包装</summary>
    /// <param name="contacts">联系人列表</param>
    public VCardDocument(List<Contact> contacts) => Contacts = contacts ?? [];
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
            if (!String.IsNullOrEmpty(c.Organization)) sb.AppendLine(L.T("Organization: {0}", c.Organization));
            if (!String.IsNullOrEmpty(c.Title)) sb.AppendLine(L.T("Job title: {0}", c.Title));
            if (c.Birthday != null) sb.AppendLine(L.T("Birthday: {0:yyyy-MM-dd}", c.Birthday));
            foreach (var phone in c.Phones)
            {
                if (!String.IsNullOrEmpty(phone.Number))
                {
                    var type = String.IsNullOrEmpty(phone.Type) ? "" : " (" + phone.Type + ")";
                    sb.AppendLine(L.T("Phone: {0}", phone.Number) + type);
                }
            }
            foreach (var email in c.Emails)
            {
                if (!String.IsNullOrEmpty(email.Address))
                {
                    var type = String.IsNullOrEmpty(email.Type) ? "" : " (" + email.Type + ")";
                    sb.AppendLine(L.T("Email: {0}", email.Address) + type);
                }
            }
            foreach (var addr in c.Addresses)
            {
                var parts = new List<String>();
                if (!String.IsNullOrEmpty(addr.Country)) parts.Add(addr.Country!);
                if (!String.IsNullOrEmpty(addr.Region)) parts.Add(addr.Region!);
                if (!String.IsNullOrEmpty(addr.City)) parts.Add(addr.City!);
                if (!String.IsNullOrEmpty(addr.Street)) parts.Add(addr.Street!);
                if (parts.Count > 0)
                {
                    var type = String.IsNullOrEmpty(addr.Type) ? "" : " (" + addr.Type + ")";
                    sb.AppendLine(L.T("Address: {0}", String.Join(" ", parts)) + type);
                }
            }
            if (!String.IsNullOrEmpty(c.Url)) sb.AppendLine(L.T("Website: {0}", c.Url));
            if (!String.IsNullOrEmpty(c.Note)) sb.AppendLine(L.T("Note: {0}", c.Note));
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
            sb.AppendLine("## " + (c.FullName ?? L.T("Unknown contact")));
            sb.AppendLine();
            if (!String.IsNullOrEmpty(c.Organization)) sb.AppendLine(L.T("- **Organization**: {0}", c.Organization));
            if (!String.IsNullOrEmpty(c.Title)) sb.AppendLine(L.T("- **Job title**: {0}", c.Title));
            if (c.Birthday != null) sb.AppendLine(L.T("- **Birthday**: {0:yyyy-MM-dd}", c.Birthday));
            foreach (var phone in c.Phones)
            {
                if (!String.IsNullOrEmpty(phone.Number))
                {
                    var type = String.IsNullOrEmpty(phone.Type) ? "" : " (" + phone.Type + ")";
                    sb.AppendLine(L.T("- **Phone**: {0}", phone.Number) + type);
                }
            }
            foreach (var email in c.Emails)
            {
                if (!String.IsNullOrEmpty(email.Address))
                {
                    var type = String.IsNullOrEmpty(email.Type) ? "" : " (" + email.Type + ")";
                    sb.AppendLine(L.T("- **Email**: {0}", email.Address) + type);
                }
            }
            foreach (var addr in c.Addresses)
            {
                var parts = new List<String>();
                if (!String.IsNullOrEmpty(addr.Country)) parts.Add(addr.Country!);
                if (!String.IsNullOrEmpty(addr.Region)) parts.Add(addr.Region!);
                if (!String.IsNullOrEmpty(addr.City)) parts.Add(addr.City!);
                if (!String.IsNullOrEmpty(addr.Street)) parts.Add(addr.Street!);
                if (parts.Count > 0)
                {
                    var type = String.IsNullOrEmpty(addr.Type) ? "" : " (" + addr.Type + ")";
                    sb.AppendLine(L.T("- **Address**: {0}", String.Join(" ", parts)) + type);
                }
            }
            if (!String.IsNullOrEmpty(c.Url)) sb.AppendLine(L.T("- **Website**: {0}", c.Url));
            if (!String.IsNullOrEmpty(c.Note)) sb.AppendLine(L.T("- **Note**: {0}", c.Note));
        }
        return sb.ToString();
    }
    #endregion
}
