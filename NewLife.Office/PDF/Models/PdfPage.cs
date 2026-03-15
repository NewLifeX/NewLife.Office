using System.Text;

namespace NewLife.Office;

/// <summary>PDF 页面对象（记录每页内容）</summary>
public class PdfPage
{
    #region 属性
    /// <summary>页面宽度（点，1 pt = 1/72 英寸）</summary>
    public Single Width { get; set; } = 595f; // A4

    /// <summary>页面高度（点）</summary>
    public Single Height { get; set; } = 842f; // A4

    /// <summary>内容流字节</summary>
    public Byte[] ContentBytes { get; set; } = [];

    /// <summary>此页引用的图片 XObject 名称→数据</summary>
    public Dictionary<String, (Byte[] Data, Int32 Width, Int32 Height, Boolean IsJpeg)> Images { get; } = [];

    /// <summary>页面旋转角度（0/90/180/270）</summary>
    public Int32 Rotation { get; set; } = 0;

    /// <summary>页面超链接注释列表（PDF 坐标：原点在左下角）</summary>
    public List<(Single X, Single Y, Single W, Single H, String Url)> LinkAnnotations { get; } = [];

    /// <summary>PDF 对象号（catalog=1, pages=2, page=3...）</summary>
    internal Int32 PageObjId { get; set; }

    /// <summary>内容流对象号</summary>
    internal Int32 ContentObjId { get; set; }
    #endregion
}