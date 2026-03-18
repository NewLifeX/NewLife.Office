using System.Text;

namespace NewLife.Office;

/// <summary>PDF 字体定义</summary>
public class PdfFont
{
    #region 属性
    /// <summary>字体资源名（如 F1）</summary>
    public String Name { get; }

    /// <summary>基础字体名（Type1 标准字体或嵌入 TrueType 名）</summary>
    public String BaseFont { get; }

    /// <summary>是否中文字体（使用 Identity-H 编码）</summary>
    public Boolean IsCjk { get; }

    /// <summary>字体文件路径（TrueType/TTC 字体文件的完整路径），null 表示使用 Adobe 内建 CJK 字体</summary>
    public String? FontFilePath { get; set; }

    /// <summary>TTC 字体集合中的字体索引（仅对 .ttc/.otc 文件有效，默认 0）</summary>
    public Int32 TtcFontIndex { get; set; }

    /// <summary>是否将字体文件嵌入 PDF（仅对找到字体文件的 TrueType 字体有效；Type1 内置字体无需嵌入）</summary>
    /// <remarks>设为 false 时写入字体引用但不嵌入字体数据，可显著减小文件体积，但 PDF 阅读器需自行安装对应字体。</remarks>
    public Boolean EmbedFont { get; set; } = true;
    #endregion

    #region 构造
    /// <summary>实例化字体</summary>
    /// <param name="name">资源名</param>
    /// <param name="baseFont">基础字体名</param>
    /// <param name="isCjk">是否中文字体</param>
    public PdfFont(String name, String baseFont, Boolean isCjk = false)
    {
        Name = name;
        BaseFont = baseFont;
        IsCjk = isCjk;
    }
    #endregion
}