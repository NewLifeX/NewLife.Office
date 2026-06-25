namespace NewLife.Office;

/// <summary>PPT 可编程幻灯片母版（S04-Master 增强）</summary>
/// <remarks>
/// 支持无需 PPT 模板文件，通过代码创建自定义母版和版式。
/// 与 <see cref="PptxWriter.LoadMaster(String)"/> 模板加载模式互补。
/// </remarks>
public class PptMaster
{
    #region 属性

    /// <summary>母版背景色（16进制 RGB，如 "1F497D"），null 表示使用主题背景</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>母版上的形状集合（如公司 Logo、页脚装饰等）</summary>
    public List<PptShape> Shapes { get; } = [];

    /// <summary>关联的版式列表</summary>
    public List<PptLayout> Layouts { get; } = [];

    #endregion

    #region 方法

    /// <summary>添加新版式</summary>
    /// <param name="name">版式显示名称（如 "标题幻灯片"、"内容页"）</param>
    /// <param name="layoutType">版式类型（如 blank、title、twoContent 等）</param>
    /// <returns>新创建的版式对象</returns>
    public PptLayout AddLayout(String name, String layoutType = "blank")
    {
        var layout = new PptLayout { Name = name, LayoutType = layoutType };
        Layouts.Add(layout);
        return layout;
    }

    #endregion
}
