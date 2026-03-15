using System.Text;

namespace NewLife.Office;

/// <summary>PDF 嵌入图片流</summary>
public class PdfImageStream
{
    #region 属性
    /// <summary>图片在文档中的顺序索引（从 0 开始）</summary>
    public Int32 Index { get; set; }

    /// <summary>图片宽度（像素）</summary>
    public Int32 Width { get; set; }

    /// <summary>图片高度（像素）</summary>
    public Int32 Height { get; set; }

    /// <summary>编码过滤器名称，如 DCTDecode、FlateDecode 等</summary>
    public String Filter { get; set; } = String.Empty;

    /// <summary>原始流字节（未解压缩），对 JPEG 可直接使用</summary>
    public Byte[] RawData { get; set; } = [];

    /// <summary>是否为 JPEG（DCTDecode）图片，可直接将 RawData 保存为 .jpg</summary>
    public Boolean IsJpeg => Filter.IndexOf("DCTDecode", StringComparison.OrdinalIgnoreCase) >= 0;
    #endregion
}