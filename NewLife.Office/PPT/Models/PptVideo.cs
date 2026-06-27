namespace NewLife.Office;

/// <summary>PPT 幻灯片嵌入视频/音频媒体</summary>
public class PptVideo
{
    #region 属性
    /// <summary>媒体字节数据</summary>
    public Byte[] Data { get; set; } = [];

    /// <summary>扩展名（mp4/mov/avi/wav/mp3 等）</summary>
    public String Extension { get; set; } = "mp4";

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; } = 6000000;

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; } = 4000000;

    /// <summary>媒体关系ID（内部用）</summary>
    public String RelId { get; set; } = String.Empty;

    /// <summary>视频海报帧/缩略图字节数据，null 表示无缩略图（写入时自动生成占位）</summary>
    public Byte[]? ThumbnailData { get; set; }

    /// <summary>缩略图扩展名（如"png"），默认 png</summary>
    public String ThumbnailExtension { get; set; } = "png";

    /// <summary>缩略图关系ID（内部用，写入时自动分配）</summary>
    public String ThumbnailRelId { get; set; } = String.Empty;
    #endregion
}
