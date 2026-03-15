namespace NewLife.Office;

/// <summary>Word 文档属性（由 docProps/core.xml 读取）</summary>
public class WordProperties
{
    #region 属性
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>描述</summary>
    public String? Description { get; set; }

    /// <summary>创建时间</summary>
    public DateTime? Created { get; set; }
    #endregion
}

/// <summary>Word 段落样式</summary>
public enum WordParagraphStyle
{
    /// <summary>普通</summary>
    Normal,
    /// <summary>一级标题</summary>
    Heading1,
    /// <summary>二级标题</summary>
    Heading2,
    /// <summary>三级标题</summary>
    Heading3,
    /// <summary>四级标题</summary>
    Heading4,
    /// <summary>五级标题</summary>
    Heading5,
    /// <summary>六级标题</summary>
    Heading6,
}

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

/// <summary>文字段（Run）</summary>
public class WordRun
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>格式属性</summary>
    public WordRunProperties? Properties { get; set; }

    /// <summary>超链接关系ID（内部用）</summary>
    public String? HyperlinkRelId { get; set; }
    #endregion
}

/// <summary>段落</summary>
public class WordParagraph
{
    #region 属性
    /// <summary>段落样式</summary>
    public WordParagraphStyle Style { get; set; } = WordParagraphStyle.Normal;

    /// <summary>文字段集合</summary>
    public List<WordRun> Runs { get; } = [];

    /// <summary>对齐方式（left/center/right/both）</summary>
    public String? Alignment { get; set; }

    /// <summary>左缩进（twips）</summary>
    public Int32? IndentLeft { get; set; }

    /// <summary>右缩进（twips）</summary>
    public Int32? IndentRight { get; set; }

    /// <summary>首行缩进（twips，正值=缩进，负值=悬挂缩进）</summary>
    public Int32? FirstLineIndent { get; set; }

    /// <summary>段前间距（twips）</summary>
    public Int32? SpaceBefore { get; set; }

    /// <summary>段后间距（twips）</summary>
    public Int32? SpaceAfter { get; set; }

    /// <summary>行距（percent × 100，如 100=单倍, 150=1.5倍, 200=双倍）</summary>
    public Int32? LineSpacingPct { get; set; }

    /// <summary>是否分页符</summary>
    public Boolean IsPageBreak { get; set; }

    /// <summary>是否项目符号列表</summary>
    public Boolean IsBullet { get; set; }

    /// <summary>书签名称（非空时在段落首尾添加书签）</summary>
    public String? BookmarkName { get; set; }
    #endregion
}

/// <summary>表格单元格</summary>
public class WordCell
{
    #region 属性
    /// <summary>段落集合</summary>
    public List<WordParagraph> Paragraphs { get; } = [];

    /// <summary>背景色（16进制 RGB）</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>合并列数</summary>
    public Int32 ColSpan { get; set; } = 1;

    /// <summary>合并行数（垂直合并）</summary>
    public Int32 RowSpan { get; set; } = 1;
    #endregion
}

/// <summary>图片元素</summary>
public class WordImageElement
{
    #region 属性
    /// <summary>图片数据</summary>
    public Byte[] ImageData { get; set; } = [];

    /// <summary>扩展名（png/jpg）</summary>
    public String Extension { get; set; } = "png";

    /// <summary>宽度（EMU，914400 = 1英寸）</summary>
    public Int64 WidthEmu { get; set; } = 3600000;

    /// <summary>高度（EMU）</summary>
    public Int64 HeightEmu { get; set; } = 2700000;

    /// <summary>关系ID</summary>
    public String RelId { get; set; } = String.Empty;
    #endregion
}

/// <summary>文档元素联合类型</summary>
public class WordElement
{
    /// <summary>类型</summary>
    public WordElementType Type { get; set; }

    /// <summary>段落（Type=Paragraph 时有效）</summary>
    public WordParagraph? Paragraph { get; set; }

    /// <summary>表格行集合（Type=Table 时有效）</summary>
    public List<List<WordCell>>? TableRows { get; set; }

    /// <summary>首行是否表头（Type=Table 时有效）</summary>
    public Boolean TableFirstRowHeader { get; set; }

    /// <summary>表格样式（Type=Table 时有效）</summary>
    public WordTableStyle? TableStyle { get; set; }

    /// <summary>图片（Type=Image 时有效）</summary>
    public WordImageElement? Image { get; set; }
}

/// <summary>表格样式配置</summary>
public class WordTableStyle
{
    #region 属性
    /// <summary>边框颜色（16进制 RGB，默认黑色）</summary>
    public String BorderColor { get; set; } = "000000";

    /// <summary>边框线宽（pt×8，默认4=0.5pt）</summary>
    public Int32 BorderSize { get; set; } = 4;

    /// <summary>表头行背景色（16进制 RGB，null=不设置）</summary>
    public String? HeaderBgColor { get; set; }

    /// <summary>表头行字体加粗</summary>
    public Boolean HeaderBold { get; set; } = true;

    /// <summary>斑马纹颜色（奇数行背景色，null=不设置）</summary>
    public String? StripeColor { get; set; }

    /// <summary>列宽列表（twips，null=自动均分）</summary>
    public Int32[]? ColumnWidths { get; set; }
    #endregion
}

/// <summary>文档元素类型</summary>
public enum WordElementType
{
    /// <summary>段落</summary>
    Paragraph,
    /// <summary>表格</summary>
    Table,
    /// <summary>图片</summary>
    Image,
}

/// <summary>页面设置</summary>
public class WordPageSettings
{
    #region 属性
    /// <summary>页面宽度（twips，1440 twips = 1英寸）</summary>
    public Int32 PageWidth { get; set; } = 11906; // A4: 210mm

    /// <summary>页面高度（twips）</summary>
    public Int32 PageHeight { get; set; } = 16838; // A4: 297mm

    /// <summary>上边距（twips）</summary>
    public Int32 MarginTop { get; set; } = 1440;

    /// <summary>下边距（twips）</summary>
    public Int32 MarginBottom { get; set; } = 1440;

    /// <summary>左边距（twips）</summary>
    public Int32 MarginLeft { get; set; } = 1800;

    /// <summary>右边距（twips）</summary>
    public Int32 MarginRight { get; set; } = 1800;

    /// <summary>横向</summary>
    public Boolean Landscape { get; set; }

    /// <summary>页眉文本</summary>
    public String? HeaderText { get; set; }

    /// <summary>页脚文本</summary>
    public String? FooterText { get; set; }

    /// <summary>水印文字（null 表示无水印）</summary>
    public String? WatermarkText { get; set; }
    #endregion
}

/// <summary>文档属性</summary>
public class WordDocumentProperties
{
    #region 属性
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>描述</summary>
    public String? Description { get; set; }
    #endregion
}
