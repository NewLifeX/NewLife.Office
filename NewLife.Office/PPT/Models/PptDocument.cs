namespace NewLife.Office;

/// <summary>PPT 演示文稿文档模型，对标 WordDocument / ExcelDocument / PdfDocument</summary>
/// <remarks>
/// 纯数据容器（POCO），不含读写逻辑。
/// 通过 <see cref="PptxReader.ReadDocument"/> 读取，通过 <see cref="PptxWriter.Save(String,PptDocument)"/> 写入。
/// <example>
/// <code>
/// // 构建演示文稿并保存
/// var doc = new PptDocument { Properties = { Title = "季度报告" } };
/// var slide = new PptSlide { Layout = "title_only" };
/// slide.TextBoxes.Add(new PptTextBox { Text = "2026年Q2", Role = "title" });
/// doc.Slides.Add(slide);
///
/// using var writer = new PptxWriter();
/// writer.Save("report.pptx", doc);
///
/// // 读取已有演示文稿
/// using var reader = new PptxReader("report.pptx");
/// var doc2 = reader.ReadDocument();
/// </code>
/// </example>
/// </remarks>
public class PptDocument
{
    #region 属性
    /// <summary>幻灯片集合</summary>
    public List<PptSlide> Slides { get; set; } = [];

    /// <summary>可编程母版（null 表示使用内置默认母版）</summary>
    public PptMaster? Master { get; set; }

    /// <summary>幻灯片宽度（EMU），默认 16:9 = 12192000</summary>
    public Int64 SlideWidth { get; set; } = 12192000;

    /// <summary>幻灯片高度（EMU），默认 16:9 = 6858000</summary>
    public Int64 SlideHeight { get; set; } = 6858000;

    /// <summary>主题强调色数组（Accent1~6，16进制 RGB 无 # 前缀），默认 Office 蓝色系</summary>
    public String[] AccentColors { get; set; } = ["4472C4", "ED7D31", "A9D18E", "FF0000", "FFC000", "70AD47"];

    /// <summary>文档属性（标题、作者、主题等）</summary>
    public PptDocumentProperties Properties { get; set; } = new();

    /// <summary>页眉页脚设置（幻灯片编号/日期/页脚文本），null 表示不显示</summary>
    public PptHeaderFooter? HeaderFooter { get; set; }
    #endregion
}
