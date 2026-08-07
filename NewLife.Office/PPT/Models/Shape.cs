namespace NewLife.Office.Ppt;

/// <summary>PPT 幻灯片文本形状</summary>
public class Shape
{
    #region 属性
    /// <summary>形状ID</summary>
    public Int32 Id { get; set; }

    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>形状类型（如 textBox, rect, ellipse, roundRect, triangle, diamond 等）</summary>
    public String ShapeType { get; set; } = String.Empty;

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>填充色（16进制 RGB），null 表示无填充（写入时使用）</summary>
    public String? FillColor { get; set; }

    /// <summary>图片填充数据（写入时使用），设置后覆盖 FillColor，使用 blipFill 替代 solidFill</summary>
    public Byte[]? FillImage { get; set; }

    /// <summary>图片填充扩展名（默认 "png"），配合 FillImage 使用</summary>
    public String FillImageExt { get; set; } = "png";

    /// <summary>渐变填充类型（"linear" 或 "radial"），null 表示不使用渐变</summary>
    public String? GradientType { get; set; }

    /// <summary>渐变起始色（16进制 RGB），配合 GradientType 使用</summary>
    public String? GradientColor1 { get; set; }

    /// <summary>渐变结束色（16进制 RGB），配合 GradientType 使用</summary>
    public String? GradientColor2 { get; set; }

    /// <summary>渐变角度（度，仅线性渐变有效，0=左到右，90=下到上），默认 90</summary>
    public Int32 GradientAngle { get; set; } = 90;

    /// <summary>形状图片填充的关系 ID（内部用）</summary>
    public String? ShapeImageRelId { get; set; }

    /// <summary>线条颜色（16进制 RGB），null 表示无线条（写入时使用）</summary>
    public String? LineColor { get; set; }

    /// <summary>线宽（EMU，12700=1pt，写入时使用）</summary>
    public Int32 LineWidth { get; set; } = 12700;

    /// <summary>文字字号（磅，写入时使用）</summary>
    public Int32 FontSize { get; set; } = 14;

    /// <summary>文字颜色（16进制 RGB，写入时使用）</summary>
    public String? FontColor { get; set; }

    /// <summary>文字粗体（写入时使用）</summary>
    public Boolean Bold { get; set; }

    /// <summary>拉丁/西文字体名称（如"Arial"），null 表示使用默认字体</summary>
    public String? LatinFontName { get; set; }

    /// <summary>东亚/中文字体名称（如"微软雅黑"），null 表示使用默认字体</summary>
    public String? EastAsianFontName { get; set; }

    /// <summary>复杂脚本字体名称（如阿拉伯/泰文），null 表示使用默认字体</summary>
    public String? ComplexScriptFontName { get; set; }

    /// <summary>符号字体名称，null 表示使用默认字体</summary>
    public String? SymbolFontName { get; set; }

    /// <summary>旋转角度（S15-02），以 60000 分之一度为单位（如 5400000=90°）</summary>
    public Int32 Rotation { get; set; }

    /// <summary>替换文本/无障碍描述（对应 OOXML descr 属性）</summary>
    public String? AltText { get; set; }

    /// <summary>圆角半径（仅 roundRect 形状有效，EMU）</summary>
    public Int64 CornerRadius { get; set; }

    /// <summary>水平翻转（对应 OOXML a:xfrm flipH）</summary>
    public Boolean FlipHorizontal { get; set; }

    /// <summary>垂直翻转（对应 OOXML a:xfrm flipV）</summary>
    public Boolean FlipVertical { get; set; }

    /// <summary>文本方向（"horz"=横排/"vert"=竖排/"vert270"=竖排270），null=默认横排</summary>
    public String? TextDirection { get; set; }

    /// <summary>文本自动适应（"norm"=自动缩小/"noAuto"=不自动适应），null=默认</summary>
    public String? TextAutoFit { get; set; }

    /// <summary>文本左边距（EMU），null=默认（25400=0.1英寸）</summary>
    public Int32? TextMarginLeft { get; set; }

    /// <summary>文本右边距（EMU），null=默认</summary>
    public Int32? TextMarginRight { get; set; }

    /// <summary>文本上边距（EMU），null=默认</summary>
    public Int32? TextMarginTop { get; set; }

    /// <summary>文本下边距（EMU），null=默认</summary>
    public Int32? TextMarginBottom { get; set; }

    /// <summary>文本是否自动换行（wrap），null=默认（true）</summary>
    public Boolean? WordWrap { get; set; }

    /// <summary>线条虚线样式（如 "dash"/"dot"/"dashDot"），null 表示实线</summary>
    public String? DashStyle { get; set; }

    /// <summary>文本垂直锚定方式（bodyPr anchor 属性：t/ctr/b），null 表示默认</summary>
    public String? Anchor { get; set; }

    /// <summary>文本左内边距（EMU，bodyPr lIns），null 表示默认</summary>
    public Int32? LeftInset { get; set; }

    /// <summary>文本右内边距（EMU，bodyPr rIns），null 表示默认</summary>
    public Int32? RightInset { get; set; }

    /// <summary>文本上内边距（EMU，bodyPr tIns），null 表示默认</summary>
    public Int32? TopInset { get; set; }

    /// <summary>文本下内边距（EMU，bodyPr bIns），null 表示默认</summary>
    public Int32? BottomInset { get; set; }

    /// <summary>形状级超链接（URL 或文件路径，写入 cNvPr hlinkClick，S20）</summary>
    public String? HyperlinkUrl { get; set; }
    #endregion

    #region 方法
    /// <summary>克隆形状，可指定新 ID 和偏移</summary>
    /// <param name="newId">新形状 ID</param>
    /// <param name="offsetX">X 偏移（EMU）</param>
    /// <param name="offsetY">Y 偏移（EMU）</param>
    /// <returns>克隆的形状</returns>
    public Shape Clone(Int32 newId, Int64 offsetX = 0, Int64 offsetY = 0)
    {
        return new Shape
        {
            Id = newId,
            Text = Text,
            ShapeType = ShapeType,
            Left = Left + offsetX,
            Top = Top + offsetY,
            Width = Width,
            Height = Height,
            FillColor = FillColor,
            FillImage = FillImage,
            FillImageExt = FillImageExt,
            GradientType = GradientType,
            GradientColor1 = GradientColor1,
            GradientColor2 = GradientColor2,
            GradientAngle = GradientAngle,
            LineColor = LineColor,
            LineWidth = LineWidth,
            FontSize = FontSize,
            FontColor = FontColor,
            Bold = Bold,
            LatinFontName = LatinFontName,
            EastAsianFontName = EastAsianFontName,
            ComplexScriptFontName = ComplexScriptFontName,
            SymbolFontName = SymbolFontName,
            Rotation = Rotation,
            AltText = AltText,
            CornerRadius = CornerRadius,
            FlipHorizontal = FlipHorizontal,
            FlipVertical = FlipVertical,
            TextDirection = TextDirection,
            TextAutoFit = TextAutoFit,
            TextMarginLeft = TextMarginLeft,
            TextMarginRight = TextMarginRight,
            TextMarginTop = TextMarginTop,
            TextMarginBottom = TextMarginBottom,
            WordWrap = WordWrap,
            DashStyle = DashStyle,
            Anchor = Anchor,
            LeftInset = LeftInset,
            RightInset = RightInset,
            TopInset = TopInset,
            BottomInset = BottomInset,
        };
    }
    #endregion
}
