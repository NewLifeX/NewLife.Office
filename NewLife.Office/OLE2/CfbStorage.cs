namespace NewLife.Office;

/// <summary>CFB 存储节点（相当于目录/文件夹）</summary>
/// <remarks>
/// 存储节点可以包含子存储和子流，通过 <see cref="Children"/> 访问，
/// 通过 <see cref="GetStorage"/> / <see cref="GetStream"/> 按名称检索。
/// </remarks>
public sealed class CfbStorage
{
    #region 属性
    /// <summary>存储名称</summary>
    public String Name { get; internal set; } = String.Empty;

    /// <summary>子节点列表（存储和流）</summary>
    public List<Object> Children { get; } = [];

    /// <summary>父存储，根节点为 null</summary>
    public CfbStorage? Parent { get; internal set; }
    #endregion

    #region 方法
    /// <summary>按名称获取子存储</summary>
    /// <param name="name">存储名称（区分大小写）</param>
    /// <returns>找到则返回 <see cref="CfbStorage"/>，否则返回 null</returns>
    public CfbStorage? GetStorage(String name)
    {
        foreach (var child in Children)
        {
            if (child is CfbStorage st && st.Name == name) return st;
        }
        return null;
    }

    /// <summary>按名称获取子流</summary>
    /// <param name="name">流名称（区分大小写）</param>
    /// <returns>找到则返回 <see cref="CfbStream"/>，否则返回 null</returns>
    public CfbStream? GetStream(String name)
    {
        foreach (var child in Children)
        {
            if (child is CfbStream st && st.Name == name) return st;
        }
        return null;
    }

    /// <summary>枚举所有子存储</summary>
    public IEnumerable<CfbStorage> Storages
    {
        get
        {
            foreach (var child in Children)
            {
                if (child is CfbStorage st) yield return st;
            }
        }
    }

    /// <summary>枚举所有子流</summary>
    public IEnumerable<CfbStream> Streams
    {
        get
        {
            foreach (var child in Children)
            {
                if (child is CfbStream st) yield return st;
            }
        }
    }

    /// <summary>添加子流（写入时使用）</summary>
    /// <param name="name">流名称</param>
    /// <param name="data">流数据</param>
    /// <returns>创建的流节点</returns>
    public CfbStream AddStream(String name, Byte[] data)
    {
        var stream = new CfbStream { Name = name, Data = data, Parent = this };
        Children.Add(stream);
        return stream;
    }

    /// <summary>添加子存储（写入时使用）</summary>
    /// <param name="name">存储名称</param>
    /// <returns>创建的存储节点</returns>
    public CfbStorage AddStorage(String name)
    {
        var storage = new CfbStorage { Name = name, Parent = this };
        Children.Add(storage);
        return storage;
    }
    #endregion
}
