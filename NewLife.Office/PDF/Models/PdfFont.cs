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