namespace NewLife.Office;

/// <summary>PDF 字体信息</summary>
public class PdfFontInfo
{
    #region 属性
    /// <summary>字体名称（BaseFont 值）</summary>
    public String? Name { get; set; }

    /// <summary>字体类型（如 Type1、TrueType、Type0、CIDFontType0、CIDFontType2）</summary>
    public String? Type { get; set; }

    /// <summary>编码方式（如 StandardEncoding、Identity-H）</summary>
    public String? Encoding { get; set; }

    /// <summary>是否嵌入</summary>
    public Boolean IsEmbedded => Name?.Contains('+') == true;
    #endregion
}
