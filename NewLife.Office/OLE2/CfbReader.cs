using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>CFB（Compound File Binary）格式读取器</summary>
/// <remarks>
/// 实现 MS-CFB 规范的解析逻辑，支持：
/// 版本 3（512 字节扇区）和版本 4（4096 字节扇区）；
/// 普通流（FAT 链）和迷你流（Mini FAT 链）；
/// DIFAT 扩展 FAT（支持超过 109 个 FAT 扇区的大文件）。
/// </remarks>
internal sealed class CfbReader
{
    #region 字段
    private readonly Stream _fs;
    private CfbHeader _header = null!;
    private Int32[] _fat = [];        // FAT 链表
    private Int32[] _miniFat = [];    // Mini FAT 链表
    private Byte[] _miniStream = [];  // 根存储的迷你流数据
    private CfbDirectoryEntry[] _dirs = [];  // 所有目录条目
    #endregion

    #region 构造
    /// <summary>从流构造读取器</summary>
    /// <param name="stream">可寻址的输入流</param>
    public CfbReader(Stream stream) => _fs = stream;
    #endregion

    #region 解析
    /// <summary>解析 CFB 文件，返回根存储树</summary>
    /// <returns>根存储节点</returns>
    public CfbStorage Parse()
    {
        // 1. 读取并解析文件头
        var headerBuf = new Byte[512];
        _fs.Seek(0, SeekOrigin.Begin);
        ReadFull(_fs, headerBuf);
        _header = CfbHeader.ReadFrom(headerBuf);

        // 2. 构建完整的 FAT 扇区 ID 列表（含 DIFAT 扩展）
        var fatSectorIds = BuildFatSectorList();

        // 3. 读入全部 FAT 数据
        _fat = ReadFatArray(fatSectorIds);

        // 4. 读取所有目录条目
        _dirs = ReadAllDirectoryEntries();

        // 5. 读取迷你 FAT
        if (_header.FirstMiniFatSectorId != CfbSectorMarker.EndOfChain &&
            _header.FirstMiniFatSectorId != CfbSectorMarker.FreeSect)
        {
            _miniFat = ReadFatLikeChain(_header.FirstMiniFatSectorId);
        }

        // 6. 读取迷你流（根存储的流数据）
        var rootEntry = _dirs[0];
        if (rootEntry.StartingSectorId != CfbSectorMarker.EndOfChain &&
            rootEntry.StreamSize > 0)
        {
            _miniStream = ReadSectorChain(rootEntry.StartingSectorId, (Int32)rootEntry.StreamSize);
        }

        // 7. 构建树结构
        var root = new CfbStorage { Name = "Root Entry" };
        if (rootEntry.ChildSid != CfbSectorMarker.NoEntry)
            BuildTree(root, rootEntry.ChildSid);

        return root;
    }

    /// <summary>构建完整的 FAT 扇区 ID 列表（含 DIFAT 扩展链）</summary>
    private List<Int32> BuildFatSectorList()
    {
        var list = new List<Int32>(_header.FatSectorCount);

        // 从文件头 DIFAT 数组读取前 109 个 FAT 扇区 ID
        foreach (var id in _header.DifatArray)
        {
            if (id == CfbSectorMarker.FreeSect || id == CfbSectorMarker.EndOfChain) break;
            list.Add(id);
        }

        // 若存在 DIFAT 扇区链，继续追加
        var difatSid = _header.FirstDifatSectorId;
        while (difatSid != CfbSectorMarker.EndOfChain && difatSid != CfbSectorMarker.FreeSect)
        {
            var buf = ReadSector(difatSid);
            var reader = new SpanReader(buf, 0, buf.Length);
            var entriesPerSector = (_header.SectorSize / 4) - 1; // 最后 4 字节是下 DIFAT 扇区 ID
            for (var i = 0; i < entriesPerSector; i++)
            {
                var id = reader.ReadInt32();
                if (id != CfbSectorMarker.FreeSect && id != CfbSectorMarker.EndOfChain)
                    list.Add(id);
            }
            difatSid = reader.ReadInt32(); // 最后4字节：下一个 DIFAT 扇区
        }

        return list;
    }

    /// <summary>从 FAT 扇区 ID 列表读入全部 FAT 数据</summary>
    private Int32[] ReadFatArray(List<Int32> fatSectorIds)
    {
        var entriesPerSector = _header.SectorSize / 4;
        var fat = new Int32[fatSectorIds.Count * entriesPerSector];
        var idx = 0;
        foreach (var sid in fatSectorIds)
        {
            var buf = ReadSector(sid);
            var reader = new SpanReader(buf, 0, buf.Length);
            for (var i = 0; i < entriesPerSector; i++)
            {
                fat[idx++] = reader.ReadInt32();
            }
        }
        return fat;
    }

    /// <summary>读取类 FAT 链（Mini FAT 使用相同格式）</summary>
    private Int32[] ReadFatLikeChain(Int32 startSid)
    {
        var data = ReadSectorChain(startSid, _header.MiniFatSectorCount * _header.SectorSize);
        var count = data.Length / 4;
        var arr = new Int32[count];
        var reader = new SpanReader(data, 0, data.Length);
        for (var i = 0; i < count; i++)
        {
            arr[i] = reader.ReadInt32();
        }
        return arr;
    }

    /// <summary>读取所有目录条目</summary>
    private CfbDirectoryEntry[] ReadAllDirectoryEntries()
    {
        var dirData = ReadSectorChain(_header.FirstDirSectorId, -1);
        var entrySize = 128;
        var count = dirData.Length / entrySize;
        var entries = new CfbDirectoryEntry[count];
        for (var i = 0; i < count; i++)
        {
            var buf = Slice(dirData, i * entrySize, entrySize);
            entries[i] = CfbDirectoryEntry.ReadFrom(buf, i);
        }
        return entries;
    }

    /// <summary>递归构建存储树</summary>
    private void BuildTree(CfbStorage parent, Int32 sid)
    {
        if (sid == CfbSectorMarker.NoEntry || sid < 0 || sid >= _dirs.Length) return;

        var entry = _dirs[sid];
        if (entry.ObjectType == CfbObjectType.Empty) return;

        // 先处理左兄弟（红黑树中序遍历）
        if (entry.LeftSibSid != CfbSectorMarker.NoEntry)
            BuildTree(parent, entry.LeftSibSid);

        // 处理当前节点
        if (entry.ObjectType == CfbObjectType.Stream)
        {
            var data = ReadStreamData(entry);
            var cfbStream = new CfbStream { Name = entry.Name, Data = data, Parent = parent };
            parent.Children.Add(cfbStream);
        }
        else if (entry.ObjectType == CfbObjectType.Storage)
        {
            var storage = new CfbStorage { Name = entry.Name, Parent = parent };
            parent.Children.Add(storage);
            if (entry.ChildSid != CfbSectorMarker.NoEntry)
                BuildTree(storage, entry.ChildSid);
        }

        // 处理右兄弟
        if (entry.RightSibSid != CfbSectorMarker.NoEntry)
            BuildTree(parent, entry.RightSibSid);
    }

    /// <summary>读取一个流条目的数据（自动识别普通流或迷你流）</summary>
    private Byte[] ReadStreamData(CfbDirectoryEntry entry)
    {
        var size = (Int32)entry.StreamSize;
        if (size == 0) return [];

        if (size < _header.MiniStreamCutoff && _miniStream.Length > 0)
            return ReadMiniStreamChain(entry.StartingSectorId, size);

        return ReadSectorChain(entry.StartingSectorId, size);
    }

    /// <summary>通过 FAT 链读取普通扇区数据</summary>
    /// <param name="startSid">起始扇区 ID</param>
    /// <param name="expectedSize">期望大小（-1 表示读取整个链）</param>
    private Byte[] ReadSectorChain(Int32 startSid, Int32 expectedSize)
    {
        var chunks = new List<Byte[]>();
        var total = 0;
        var sid = startSid;
        while (sid != CfbSectorMarker.EndOfChain && sid != CfbSectorMarker.FreeSect && sid >= 0)
        {
            var sector = ReadSector(sid);
            chunks.Add(sector);
            total += sector.Length;
            if (sid >= _fat.Length) break;
            sid = _fat[sid];
        }

        // 如果没有数据
        if (total == 0) return [];

        // 拼接并截断到期望大小
        var result = new Byte[total];
        var pos = 0;
        foreach (var chunk in chunks)
        {
            Array.Copy(chunk, 0, result, pos, chunk.Length);
            pos += chunk.Length;
        }

        if (expectedSize > 0 && expectedSize < total)
            return Slice(result, 0, expectedSize);

        return result;
    }

    /// <summary>通过 Mini FAT 链读取迷你流数据</summary>
    private Byte[] ReadMiniStreamChain(Int32 startMiniSid, Int32 expectedSize)
    {
        var miniSectorSize = _header.MiniSectorSize;
        var chunks = new List<Byte[]>();
        var total = 0;
        var msid = startMiniSid;
        while (msid != CfbSectorMarker.EndOfChain && msid != CfbSectorMarker.FreeSect && msid >= 0)
        {
            var offset = msid * miniSectorSize;
            if (offset + miniSectorSize <= _miniStream.Length)
            {
                chunks.Add(Slice(_miniStream, offset, miniSectorSize));
                total += miniSectorSize;
            }
            if (msid >= _miniFat.Length) break;
            msid = _miniFat[msid];
        }

        var result = new Byte[total];
        var pos = 0;
        foreach (var chunk in chunks)
        {
            Array.Copy(chunk, 0, result, pos, chunk.Length);
            pos += chunk.Length;
        }

        if (expectedSize > 0 && expectedSize < total)
            return Slice(result, 0, expectedSize);

        return result;
    }

    /// <summary>读取指定编号的扇区（512 或 4096 字节）</summary>
    private Byte[] ReadSector(Int32 sectorId)
    {
        var buf = new Byte[_header.SectorSize];
        var offset = (Int64)(sectorId + 1) * _header.SectorSize;
        _fs.Seek(offset, SeekOrigin.Begin);
        ReadFull(_fs, buf);
        return buf;
    }

    /// <summary>兼容所有目标框架的全量读取（替代 net7+ 的 ReadExactly）</summary>
    private static void ReadFull(Stream stream, Byte[] buf)
    {
        var offset = 0;
        while (offset < buf.Length)
        {
            var read = stream.Read(buf, offset, buf.Length - offset);
            if (read == 0) throw new EndOfStreamException("Unexpected end of CFB stream.");
            offset += read;
        }
    }

    /// <summary>兼容所有目标框架的数组切片（替代 C# 8 range 语法）</summary>
    private static Byte[] Slice(Byte[] src, Int32 start, Int32 length)
    {
        var result = new Byte[length];
        Array.Copy(src, start, result, 0, length);
        return result;
    }
    #endregion
}
