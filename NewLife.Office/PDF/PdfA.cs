namespace NewLife.Office;

/// <summary>PDF/A 合规级别</summary>
public enum PdfACompliance
{
    /// <summary>PDF/A-1b（ISO 19005-1，Level B 基础合规）</summary>
    PDF_A_1B,

    /// <summary>PDF/A-2b（ISO 19005-2，Level B 基础合规）</summary>
    PDF_A_2B,

    /// <summary>PDF/A-3b（ISO 19005-3，Level B 基础合规）</summary>
    PDF_A_3B,
}

/// <summary>PDF/A 输出意图（sRGB IEC61966-2.1 ICC 颜色配置文件）</summary>
/// <remarks>
/// PDF/A 规范要求文档包含输出意图（/OutputIntent），指向目标颜色空间。
/// sRGB IEC61966-2.1 是最常用的默认输出意图。
/// 完整实现需要嵌入 ICC 配置文件字节，此处使用标准引用。
/// </remarks>
internal static class PdfAConstants
{
    /// <summary>sRGB IEC61966-2.1 ICC 配置文件标识符</summary>
    public const String SrgbIecProfileIdentifier = "sRGB IEC61966-2.1";

    /// <summary>PDF/A-1 部分标识符</summary>
    public const String PdfA1Part = "1";

    /// <summary>PDF/A-2 部分标识符</summary>
    public const String PdfA2Part = "2";

    /// <summary>PDF/A-3 部分标识符</summary>
    public const String PdfA3Part = "3";

    /// <summary>生成 XMP 元数据流</summary>
    /// <param name="part">PDF/A 部分（1/2/3）</param>
    /// <param name="conformance">合规级别（B）</param>
    /// <returns>XMP 字节</returns>
    public static Byte[] GenerateXmpMetadata(Int32 part, Char conformance = 'B')
    {
        var xmp = $@"<?xpacket begin=""﻿"" id=""W5M0MpCehiHzreSzNTczkc9d""?>
<x:xmpmeta xmlns:x=""adobe:ns:meta/"">
  <rdf:RDF xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#"">
    <rdf:Description rdf:about="""" xmlns:pdfaid=""http://www.aiim.org/pdfa/ns/id/"">
      <pdfaid:part>{part}</pdfaid:part>
      <pdfaid:conformance>{conformance}</pdfaid:conformance>
    </rdf:Description>
  </rdf:RDF>
</x:xmpmeta>
<?xpacket end=""w""?>";
        return System.Text.Encoding.UTF8.GetBytes(xmp);
    }
}
