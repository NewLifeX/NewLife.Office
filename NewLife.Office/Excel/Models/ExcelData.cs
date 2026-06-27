namespace NewLife.Office;

/// <summary>Excel 工作簿完整数据快照，用于读写往返</summary>
/// <remarks>
/// 包含所有工作表及其完整数据、样式、布局和元数据。
/// 通过 <see cref="ExcelReader.ReadExcel"/> 读取，通过 <see cref="ExcelWriter.WriteExcel"/> 写入。
/// </remarks>
public class ExcelDocument
{
    #region 属性
    /// <summary>工作表集合</summary>
    public List<ExcelSheet> Sheets { get; set; } = [];

    /// <summary>
    /// 原样透传的其他 ZIP 部件（主题/文档属性/外部链接/VBA 等）。
    /// Key 为 ZIP 入口路径（如 "xl/theme/theme1.xml"），Value 为原始字节。
    /// Reader 填充，Writer 在 Save 时原样写回，确保未解析内容不丢失。
    /// </summary>
    public Dictionary<String, Byte[]> OtherParts { get; set; } = [];

    /// <summary>
    /// 用户自定义命名范围（排除 _xlnm.* 系统名如打印标题）。
    /// Key 为名称，Value 为公式/范围引用（如 "Sheet1!$A$1:$B$10"）。
    /// </summary>
    public Dictionary<String, String> DefinedNames { get; set; } = [];

    /// <summary>
    /// 默认字体（对应 styles.xml 中 fonts[0]，用于行/列标题渲染）。
    /// 由 Reader 从源文件中读取，Writer 用于生成 font[0]，确保标题字体与原文件一致。
    /// </summary>
    public ExcelDefaultFont? DefaultFont { get; set; }
    #endregion
}

/// <summary>默认字体信息（对应 styles.xml fonts[0]）</summary>
public class ExcelDefaultFont
{
    /// <summary>字体名称（如 "宋体"、"Calibri"）</summary>
    public String? Name { get; set; }

    /// <summary>字体大小（磅）</summary>
    public Double Size { get; set; }

    /// <summary>是否粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>颜色（RGB 六位十六进制或 "theme:N" 格式）</summary>
    public String? Color { get; set; }
}
