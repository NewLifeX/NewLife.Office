namespace NewLife.Office;

/// <summary>PPT 幻灯片文本框</summary>
public class PptTextBox
{
    #region 属性
    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字号（磅）</summary>
    public Int32 FontSize { get; set; } = 18;

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>文字颜色（16进制 RGB，如 "000000"）</summary>
    public String? FontColor { get; set; }

    /// <summary>对齐（l/ctr/r）</summary>
    public String Alignment { get; set; } = "l";

    /// <summary>背景色（16进制 RGB），null 表示透明</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>拉丁/西文字体名称（如"Arial"），null 表示使用默认字体</summary>
    public String? LatinFontName { get; set; }

    /// <summary>东亚/中文字体名称（如"微软雅黑"），null 表示使用默认字体</summary>
    public String? EastAsianFontName { get; set; }

    /// <summary>复杂脚本字体名称（如阿拉伯/泰文），null 表示使用默认字体</summary>
    public String? ComplexScriptFontName { get; set; }

    /// <summary>符号字体名称，null 表示使用默认字体</summary>
    public String? SymbolFontName { get; set; }

    /// <summary>字体名称（如"微软雅黑"），null 表示使用默认字体。兼容属性：getter 返回 EastAsianFontName ?? LatinFontName</summary>
    public String? FontName
    {
        get => EastAsianFontName ?? LatinFontName;
        set => LatinFontName = EastAsianFontName = value;
    }

    /// <summary>超链接 URL</summary>
    public String? HyperlinkUrl { get; set; }

    /// <summary>bodyPr 自动适应模式：0=normAutofit 1=spAutoFit(auto-height) 2=noAutofit</summary>
    public Int32 AutoFit { get; set; }  // 0=norm 1=spAutoFit 2=no

    /// <summary>文本垂直锁定方式（bodyPr anchor 属性：t=顶部/ctr=居中/b=底部），空表示不设置</summary>
    public String Anchor { get; set; } = String.Empty;

    /// <summary>行间距（10万分为单位的百分比，如 100000 = 100%），0 表示不设置</summary>
    public Int32 LineSpacingPct { get; set; }

    /// <summary>段前间距（pt，如 0）</summary>
    public Int32 SpaceBeforePt { get; set; }

    /// <summary>富文本片段集合（向后兼容：优先使用 Paragraphs）</summary>
    public List<PptTextRun> Runs { get; } = [];

    /// <summary>段落集合（多段文本时使用，每个段落含独立 Run 列表和段落级格式）。非空时 Writer 优先使用此属性逐段写入</summary>
    public List<PptParagraph> Paragraphs { get; } = [];

    /// <summary>文本框左内边距（EMU，bodyPr lIns），0 表示使用默认值</summary>
    public Int32 LeftInset { get; set; }

    /// <summary>文本框右内边距（EMU，bodyPr rIns），0 表示使用默认值</summary>
    public Int32 RightInset { get; set; }

    /// <summary>文本框上内边距（EMU，bodyPr tIns），0 表示使用默认值</summary>
    public Int32 TopInset { get; set; }

    /// <summary>文本框下内边距（EMU，bodyPr bIns），0 表示使用默认值</summary>
    public Int32 BottomInset { get; set; }

    /// <summary>语义角色，供 LayoutEngine 自动排版使用，null 表示不参与自动排版</summary>
    /// <remarks>支持：title/subtitle/body/kpi/caption</remarks>
    public String? Role { get; set; }

    /// <summary>旋转角度（S15-02），以 60000 分之一度为单位</summary>
    public Int32 Rotation { get; set; }

    /// <summary>替换文本/无障碍描述（对应 OOXML descr 属性）</summary>
    public String? AltText { get; set; }

    /// <summary>文本方向（vert 属性值：horz=水平/vert=垂直/vert270=旋转270°/eaVert=东亚竖排），null 表示默认水平</summary>
    public String? TextDirection { get; set; }
    #endregion
}
