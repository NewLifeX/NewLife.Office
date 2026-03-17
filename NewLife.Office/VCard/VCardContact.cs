namespace NewLife.Office;

/// <summary>vCard 联系人（RFC 6350，vCard 4.0）</summary>
/// <remarks>
/// 支持 vCard 2.1、3.0、4.0 格式的读写。
/// 核心属性：FN/N/TEL/EMAIL/ADR/PHOTO/URL/NOTE/BDAY/ORG/TITLE。
/// </remarks>
public class VCardContact
{
    #region 属性

    /// <summary>格式化姓名（FN）</summary>
    public String? FullName { get; set; }

    /// <summary>姓名分量（N）：Family;Given;Additional;Prefix;Suffix</summary>
    public VCardName? Name { get; set; }

    /// <summary>组织（ORG）</summary>
    public String? Organization { get; set; }

    /// <summary>职位（TITLE）</summary>
    public String? Title { get; set; }

    /// <summary>生日（BDAY）</summary>
    public DateTime? Birthday { get; set; }

    /// <summary>备注（NOTE）</summary>
    public String? Note { get; set; }

    /// <summary>个人主页（URL）</summary>
    public String? Url { get; set; }

    /// <summary>照片 URL 或 Base64（PHOTO）</summary>
    public String? Photo { get; set; }

    /// <summary>UID</summary>
    public String? Uid { get; set; }

    /// <summary>修订时间（REV）</summary>
    public DateTimeOffset? Revision { get; set; }

    /// <summary>电话号码列表（TEL）</summary>
    public List<VCardPhone> Phones { get; } = new List<VCardPhone>();

    /// <summary>电子邮件列表（EMAIL）</summary>
    public List<VCardEmail> Emails { get; } = new List<VCardEmail>();

    /// <summary>地址列表（ADR）</summary>
    public List<VCardAddress> Addresses { get; } = new List<VCardAddress>();

    /// <summary>扩展属性（X-* 和未知属性）</summary>
    public Dictionary<String, String> ExtraProps { get; } = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);

    #endregion
}

/// <summary>vCard 姓名分量（N 属性）</summary>
public class VCardName
{
    /// <summary>姓（Family Name）</summary>
    public String? Family { get; set; }

    /// <summary>名（Given Name）</summary>
    public String? Given { get; set; }

    /// <summary>中间名（Additional Name）</summary>
    public String? Additional { get; set; }

    /// <summary>前缀（Prefix）</summary>
    public String? Prefix { get; set; }

    /// <summary>后缀（Suffix）</summary>
    public String? Suffix { get; set; }
}

/// <summary>vCard 电话（TEL）</summary>
public class VCardPhone
{
    /// <summary>电话号码</summary>
    public String? Number { get; set; }

    /// <summary>类型（WORK/HOME/CELL/FAX 等，逗号分隔）</summary>
    public String? Type { get; set; }
}

/// <summary>vCard 邮件（EMAIL）</summary>
public class VCardEmail
{
    /// <summary>邮箱地址</summary>
    public String? Address { get; set; }

    /// <summary>类型（WORK/HOME 等）</summary>
    public String? Type { get; set; }
}

/// <summary>vCard 地址（ADR）：PO Box;;Street;City;Region;PostalCode;Country</summary>
public class VCardAddress
{
    /// <summary>邮政信箱</summary>
    public String? PoBox { get; set; }

    /// <summary>扩展地址</summary>
    public String? Extended { get; set; }

    /// <summary>街道</summary>
    public String? Street { get; set; }

    /// <summary>城市</summary>
    public String? City { get; set; }

    /// <summary>省/州</summary>
    public String? Region { get; set; }

    /// <summary>邮编</summary>
    public String? PostalCode { get; set; }

    /// <summary>国家</summary>
    public String? Country { get; set; }

    /// <summary>类型（WORK/HOME 等）</summary>
    public String? Type { get; set; }
}
