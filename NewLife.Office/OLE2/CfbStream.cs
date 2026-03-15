namespace NewLife.Office;

/// <summary>CFB 流节点（相当于文件）</summary>
/// <remarks>
/// 流节点包含实际数据，通过 <see cref="Data"/> 属性读取全部内容，
/// 或通过 <see cref="OpenRead"/> 获取只读流。
/// </remarks>
public sealed class CfbStream
{
    #region 属性
    /// <summary>流名称</summary>
    public String Name { get; internal set; } = String.Empty;

    /// <summary>流数据（从 CFB 文件读出的完整字节数组）</summary>
    public Byte[] Data { get; internal set; } = [];

    /// <summary>父存储</summary>
    public CfbStorage? Parent { get; internal set; }

    /// <summary>流大小（字节）</summary>
    public Int32 Length => Data.Length;
    #endregion

    #region 方法
    /// <summary>以只读方式打开流</summary>
    /// <returns>内存流</returns>
    public Stream OpenRead() => new MemoryStream(Data, false);

    /// <summary>以指定偏移和长度获取子数据</summary>
    /// <param name="offset">起始偏移</param>
    /// <param name="length">长度</param>
    /// <returns>子数组</returns>
    public Byte[] GetBytes(Int32 offset, Int32 length)
    {
        var result = new Byte[length];
        Array.Copy(Data, offset, result, 0, length);
        return result;
    }
    #endregion
}
