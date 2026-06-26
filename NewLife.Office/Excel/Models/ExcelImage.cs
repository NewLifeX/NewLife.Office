namespace NewLife.Office;

/// <summary>Excel 工作表图片元素</summary>
public class ExcelImage
{
    #region 属性
    /// <summary>图片字节数据</summary>
    public Byte[] Data { get; set; } = [];

    /// <summary>扩展名（png/jpg/gif）</summary>
    public String Extension { get; set; } = "png";

    /// <summary>起始行（0基，来自 drawing from.row）</summary>
    public Int32 Row { get; set; }

    /// <summary>起始列（0基，来自 drawing from.col）</summary>
    public Int32 Col { get; set; }

    /// <summary>宽度（像素）</summary>
    public Double Width { get; set; } = 100;

    /// <summary>高度（像素）</summary>
    public Double Height { get; set; } = 100;

    /// <summary>起始列偏移（EMU，来自 from.colOff）</summary>
    public Int64 FromColOff { get; set; }

    /// <summary>起始行偏移（EMU，来自 from.rowOff）</summary>
    public Int64 FromRowOff { get; set; }

    /// <summary>结束列（0基，来自 to.col）；-1 表示未设置</summary>
    public Int32 ToCol { get; set; } = -1;

    /// <summary>结束行（0基，来自 to.row）；-1 表示未设置</summary>
    public Int32 ToRow { get; set; } = -1;

    /// <summary>结束列偏移（EMU，来自 to.colOff）</summary>
    public Int64 ToColOff { get; set; }

    /// <summary>结束行偏移（EMU，来自 to.rowOff）</summary>
    public Int64 ToRowOff { get; set; }

    /// <summary>锚点编辑属性（"oneCell"/"twoCell"/"absolute"；默认 "oneCell"）</summary>
    public String EditAs { get; set; } = "oneCell";
    #endregion
}
