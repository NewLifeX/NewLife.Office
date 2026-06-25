namespace NewLife.Office;

/// <summary>PPT 可编程幻灯片版式（S04-Master 增强）</summary>
/// <remarks>支持在编程式母版中自定义版式的占位符和显示名称，无需依赖外部模板文件。</remarks>
public class PptLayout
{
    #region 属性

    /// <summary>版式显示名称（如 "标题幻灯片"）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>版式类型（如 blank、title、twoContent、ctrTitle 等）</summary>
    public String LayoutType { get; set; } = "blank";

    /// <summary>版式上的形状集合（如占位符文本框、装饰图形）</summary>
    public List<PptShape> Shapes { get; } = [];

    /// <summary>版式上的文本框集合</summary>
    public List<PptTextBox> TextBoxes { get; } = [];

    #endregion
}
