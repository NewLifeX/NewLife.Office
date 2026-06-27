namespace NewLife.Office;

/// <summary>字体值对象，跨 Word/PPT/Excel/PDF 四格式通用</summary>
/// <remarks>
/// 轻量值对象，描述文字字体、大小、样式及颜色，可在各格式 Writer API 中统一传递。
/// <example>
/// <code>
/// var heading = new OfficeFont { Name = "微软雅黑", Size = 18, Bold = true, Color = "1F497D" };
/// var body    = new OfficeFont { Name = "宋体",    Size = 11 };
/// </code>
/// </example>
/// </remarks>
public class OfficeFont
{
    #region 属性
    /// <summary>西文/拉丁字体名称（如 "Arial"），null 表示继承默认</summary>
    public String? Name { get; set; }

    /// <summary>东亚/中文字体名称（如 "微软雅黑"），null 表示与 Name 相同</summary>
    public String? EastAsianName { get; set; }

    /// <summary>字号（磅），null 表示继承默认</summary>
    public Single? Size { get; set; }

    /// <summary>是否粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>是否斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>是否下划线</summary>
    public Boolean Underline { get; set; }

    /// <summary>是否删除线</summary>
    public Boolean Strikethrough { get; set; }

    /// <summary>文字颜色（16进制 RGB，无 # 前缀，如 "FF0000"），null 表示继承默认</summary>
    public String? Color { get; set; }

    /// <summary>有效的拉丁字体名（EastAsianName 优先，其次 Name）</summary>
    public String? EffectiveName => EastAsianName ?? Name;
    #endregion

    #region 静态预设
    /// <summary>正文默认字体（宋体 11pt）</summary>
    public static readonly OfficeFont Body = new() { Name = "Calibri", EastAsianName = "宋体", Size = 11 };

    /// <summary>标题默认字体（微软雅黑 18pt 粗体）</summary>
    public static readonly OfficeFont Heading = new() { Name = "Calibri", EastAsianName = "微软雅黑", Size = 18, Bold = true };

    /// <summary>代码字体（Courier New 10pt）</summary>
    public static readonly OfficeFont Code = new() { Name = "Courier New", Size = 10 };
    #endregion
}
