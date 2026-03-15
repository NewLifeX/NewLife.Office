namespace NewLife.Office;

/// <summary>文字格式属性</summary>
public class WordRunProperties
{
    #region 属性
    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>下划线</summary>
    public Boolean Underline { get; set; }

    /// <summary>前景色（16进制 RGB，如 "FF0000"）</summary>
    public String? ForeColor { get; set; }

    /// <summary>字号（磅）</summary>
    public Single? FontSize { get; set; }

    /// <summary>字体名称</summary>
    public String? FontName { get; set; }
    #endregion
}
