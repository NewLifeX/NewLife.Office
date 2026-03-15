using NewLife.Log;

namespace NewLife.Office;

/// <summary>CFB（Compound File Binary / OLE2）文档</summary>
/// <remarks>
/// 实现 MS-CFB 规范，提供对 OLE2 复合文档的读写能力。
/// 旧版 Office 格式（.xls/.doc/.ppt 等）均以 CFB 作为容器格式。
/// <para>读取示例：</para>
/// <code>
/// using var doc = CfbDocument.Open("file.xls");
/// var stream = doc.Root.GetStorage("Workbook")?.Data;
/// </code>
/// <para>写入示例：</para>
/// <code>
/// var doc = new CfbDocument();
/// doc.Root.AddStream("MyStream", data);
/// doc.Save("output.ole");
/// </code>
/// </remarks>
public sealed class CfbDocument : IDisposable
{
    #region 属性
    /// <summary>根存储节点</summary>
    public CfbStorage Root { get; private set; } = new CfbStorage { Name = "Root Entry" };

    /// <summary>日志</summary>
    public ILog Log { get; set; } = Logger.Null;
    #endregion

    #region 构造与打开
    /// <summary>创建空的 CFB 文档（用于写入）</summary>
    public CfbDocument() { }

    /// <summary>从文件路径打开 CFB 文档</summary>
    /// <param name="path">CFB 文件路径</param>
    /// <returns>已解析的文档对象</returns>
    public static CfbDocument Open(String path)
    {
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return OpenCore(fs, true);
    }

    /// <summary>从流打开 CFB 文档</summary>
    /// <param name="stream">可寻址的输入流</param>
    /// <param name="leaveOpen">解析后是否保持流开启（false 则由文档管理流生命周期）</param>
    /// <returns>已解析的文档对象</returns>
    public static CfbDocument Open(Stream stream, Boolean leaveOpen = true)
    {
        if (!stream.CanSeek)
            throw new ArgumentException("Stream must be seekable.", nameof(stream));
        return OpenCore(stream, !leaveOpen);
    }

    private static CfbDocument Open(Byte[] data)
    {
        return OpenCore(new MemoryStream(data, false), true);
    }

    private static CfbDocument OpenCore(Stream stream, Boolean ownStream)
    {
        var doc = new CfbDocument();
        var reader = new CfbReader(stream);
        doc.Root = reader.Parse();
        if (ownStream) stream.Dispose();
        return doc;
    }
    #endregion

    #region 保存
    /// <summary>保存到文件</summary>
    /// <param name="path">输出文件路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs);
    }

    /// <summary>序列化为字节数组</summary>
    /// <returns>CFB 格式的字节数组</returns>
    public Byte[] ToBytes()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }

    /// <summary>将文档写入输出流</summary>
    /// <param name="output">可写输出流</param>
    public void Save(Stream output)
    {
        var writer = new CfbWriter();
        writer.Write(Root, output);
    }
    #endregion

    #region 路径访问
    /// <summary>按路径访问流（路径分隔符为 /，如 "Workbook" 或 "Storage1/Stream2"）</summary>
    /// <param name="path">流路径</param>
    /// <returns>找到则返回字节数组，否则 null</returns>
    public Byte[]? GetStreamData(String path)
    {
        var parts = path.Split('/');
        var storage = Root;
        for (var i = 0; i < parts.Length - 1; i++)
        {
            storage = storage.GetStorage(parts[i]);
            if (storage == null) return null;
        }
        return storage.GetStream(parts[parts.Length - 1])?.Data;
    }

    /// <summary>按路径添加流</summary>
    /// <param name="path">流路径（父存储若不存在则自动创建）</param>
    /// <param name="data">流数据</param>
    public void PutStream(String path, Byte[] data)
    {
        var parts = path.Split('/');
        var storage = Root;
        for (var i = 0; i < parts.Length - 1; i++)
        {
            storage = storage.GetStorage(parts[i]) ?? storage.AddStorage(parts[i]);
        }

        var streamName = parts[parts.Length - 1];
        // 若已存在则替换
        var existing = storage.GetStream(streamName);
        if (existing != null)
        {
            existing.Data = data;
            return;
        }
        storage.AddStream(streamName, data);
    }
    #endregion

    #region IDisposable
    /// <summary>释放资源</summary>
    public void Dispose() { }
    #endregion
}
