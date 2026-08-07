namespace NewLife.Office.Word;

/// <summary>图片元素</summary>
public class Image
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

    /// <summary>是否 SVG 矢量图片（对应 asvg:svgBlip，而非 a:blip）</summary>
    public Boolean IsSvg { get; set; }

    /// <summary>环绕/定位类型：inline=嵌入文本行（默认），anchor=浮动锚定</summary>
    public String AnchorType { get; set; } = "inline";

    /// <summary>浮动图片水平位置（wp:positionH 的 align/offset 原始值，如 "center" / "offset"）</summary>
    public String? AnchorPosH { get; set; }

    /// <summary>浮动图片垂直位置（wp:positionV 的 align/offset 原始值，如 "top" / "offset"）</summary>
    public String? AnchorPosV { get; set; }

    /// <summary>浮动图片水平偏移（wp:positionH w:posOffset，EMU），非 offset 时可能为 null</summary>
    public Int64? AnchorOffsetX { get; set; }

    /// <summary>浮动图片垂直偏移（wp:positionV w:posOffset，EMU）</summary>
    public Int64? AnchorOffsetY { get; set; }

    /// <summary>环绕方式（anchor 时有效）：square/tight/through/topAndBottom/none（对应 wp:wrap* 子元素）</summary>
    public String? Wrap { get; set; }

    /// <summary>替代文本（wp:docPr descr / a:blip cstate），用于图片占位符识别</summary>
    public String? AltText { get; set; }
    #endregion
}
