using NewLife.Buffers;

namespace NewLife.Office.Ole2;

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
        // 1. 顺序读取 512 字节文件头（SpanReader 适合此处的顺序读取）
        _fs.Seek(0, SeekOrigin.Begin);
        var headerReader = new SpanReader(_fs, bufferSize: 512);
        _header = CfbHeader.ReadFrom(ref headerReader);

        // 文件长度决定合理的扇区总数，避免头部伪造超大计数直接分配
        var maxSectors = _fs.Length / _header.SectorSize + 2;
        if (maxSectors > Int32.MaxValue) maxSectors = Int32.MaxValue;
        var sectorLimit = (Int32)maxSectors;

        // 2. 构建完整的 FAT 扇区 ID 列表（含 DIFAT 扩展）
        var fatSectorIds = BuildFatSectorList(sectorLimit);

        // 3. 读入全部 FAT 数据
        _fat = ReadFatArray(fatSectorIds, sectorLimit);

        // 4. 读取所有目录条目
        _dirs = ReadAllDirectoryEntries();
        if (_dirs.Length == 0)
            throw new InvalidDataException(L.T("The CFB directory is empty."));

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
            if (rootEntry.StreamSize > Int32.MaxValue)
                throw new InvalidDataException(L.T("The CFB mini stream is too large."));
            _miniStream = ReadSectorChain(rootEntry.StartingSectorId, (Int32)rootEntry.StreamSize);
        }

        // 7. 构建树结构
        var root = new CfbStorage { Name = "Root Entry" };
        if (rootEntry.ChildSid != CfbSectorMarker.NoEntry)
            BuildTree(root, rootEntry.ChildSid, new HashSet<Int32>(), 0);

        return root;
    }

    /// <summary>构建完整的 FAT 扇区 ID 列表（含 DIFAT 扩展链）</summary>
    /// <param name="sectorLimit">文件长度允许的最大扇区数</param>
    private List<Int32> BuildFatSectorList(Int32 sectorLimit)
    {
        var capacity = Math.Min(_header.FatSectorCount, Math.Max(0, sectorLimit));
        var list = new List<Int32>(capacity);
        var visited = new HashSet<Int32>();

        // 从文件头 DIFAT 数组读取前 109 个 FAT 扇区 ID
        foreach (var id in _header.DifatArray)
        {
            if (id == CfbSectorMarker.FreeSect || id == CfbSectorMarker.EndOfChain) break;
            if (id < 0 || id >= sectorLimit)
                throw new InvalidDataException(L.T("Invalid CFB FAT sector id."));
            list.Add(id);
        }

        // 若存在 DIFAT 扇区链，继续追加
        var difatSid = _header.FirstDifatSectorId;
        var sectorBuf = new Byte[_header.SectorSize];
        while (difatSid != CfbSectorMarker.EndOfChain && difatSid != CfbSectorMarker.FreeSect)
        {
            if (difatSid < 0 || difatSid >= sectorLimit || !visited.Add(difatSid))
                throw new InvalidDataException(L.T("Invalid or cyclic CFB DIFAT chain."));
            if (list.Count > sectorLimit)
                throw new InvalidDataException(L.T("The CFB FAT sector list exceeds the resource limit."));

            ReadAt((Int64)(difatSid + 1) * _header.SectorSize, sectorBuf, 0, _header.SectorSize);
            var sr = new SpanReader(sectorBuf, 0, _header.SectorSize);
            var entriesPerSector = (_header.SectorSize / 4) - 1; // 最后 4 字节是下 DIFAT 扇区 ID
            for (var i = 0; i < entriesPerSector; i++)
            {
                var id = sr.ReadInt32();
                if (id != CfbSectorMarker.FreeSect && id != CfbSectorMarker.EndOfChain)
                {
                    if (id < 0 || id >= sectorLimit)
                        throw new InvalidDataException(L.T("Invalid CFB FAT sector id."));
                    list.Add(id);
                }
            }
            difatSid = sr.ReadInt32(); // 最后4字节：下一个 DIFAT 扇区
        }

        return list;
    }

    /// <summary>从 FAT 扇区 ID 列表读入全部 FAT 数据</summary>
    /// <param name="fatSectorIds">FAT 扇区 ID 列表</param>
    /// <param name="sectorLimit">文件长度允许的最大扇区数</param>
    private Int32[] ReadFatArray(List<Int32> fatSectorIds, Int32 sectorLimit)
    {
        var entriesPerSector = _header.SectorSize / 4;
        var total = (Int64)fatSectorIds.Count * entriesPerSector;
        if (total > OfficeLimits.MaxCfbFatEntries || total > (Int64)sectorLimit + entriesPerSector)
            throw new InvalidDataException(L.T("The CFB FAT table exceeds the resource limit."));

        var fat = new Int32[(Int32)total];
        var idx = 0;
        var sectorBuf = new Byte[_header.SectorSize];
        foreach (var sid in fatSectorIds)
        {
            if (sid < 0 || sid >= sectorLimit)
                throw new InvalidDataException(L.T("Invalid CFB FAT sector id."));

            ReadAt((Int64)(sid + 1) * _header.SectorSize, sectorBuf, 0, _header.SectorSize);
            var sr = new SpanReader(sectorBuf, 0, _header.SectorSize);
            for (var i = 0; i < entriesPerSector; i++)
                fat[idx++] = sr.ReadInt32();
        }
        return fat;
    }

    /// <summary>读取类 FAT 链（Mini FAT 使用相同格式）</summary>
    private Int32[] ReadFatLikeChain(Int32 startSid)
    {
        var expected = (Int64)_header.MiniFatSectorCount * _header.SectorSize;
        if (expected > Int32.MaxValue || expected > OfficeLimits.MaxCfbStreamBytes)
            throw new InvalidDataException(L.T("The CFB mini FAT exceeds the resource limit."));

        var data = ReadSectorChain(startSid, (Int32)expected);
        var count = data.Length / 4;
        var arr = new Int32[count];
        var sr = new SpanReader(data, 0, data.Length);
        for (var i = 0; i < count; i++)
            arr[i] = sr.ReadInt32();
        return arr;
    }

    /// <summary>读取所有目录条目</summary>
    private CfbDirectoryEntry[] ReadAllDirectoryEntries()
    {
        var dirData = ReadSectorChain(_header.FirstDirSectorId, -1);
        var entrySize = 128;
        var count = dirData.Length / entrySize;
        if (count > OfficeLimits.MaxCfbDirectoryEntries)
            throw new InvalidDataException(L.T("The CFB directory exceeds the resource limit."));

        var entries = new CfbDirectoryEntry[count];
        for (var i = 0; i < count; i++)
            entries[i] = CfbDirectoryEntry.ReadFrom(dirData.AsSpan(i * entrySize, entrySize), i);
        return entries;
    }

    /// <summary>递归构建存储树（带 visited 与深度上限，防止恶意环/深链）</summary>
    private void BuildTree(CfbStorage parent, Int32 sid, HashSet<Int32> visited, Int32 depth)
    {
        if (sid == CfbSectorMarker.NoEntry || sid < 0 || sid >= _dirs.Length) return;
        if (depth > OfficeLimits.MaxCfbTreeDepth)
            throw new InvalidDataException(L.T("The CFB directory tree is nested too deeply."));
        if (!visited.Add(sid))
            throw new InvalidDataException(L.T("The CFB directory tree contains a cycle."));

        var entry = _dirs[sid];
        if (entry.ObjectType == ObjectType.Empty) return;

        // 先处理左兄弟（红黑树中序遍历）
        if (entry.LeftSibSid != CfbSectorMarker.NoEntry)
            BuildTree(parent, entry.LeftSibSid, visited, depth + 1);

        // 处理当前节点
        if (entry.ObjectType == ObjectType.Stream)
        {
            var data = ReadStreamData(entry);
            var cfbStream = new CfbStream { Name = entry.Name, Data = data, Parent = parent };
            parent.Children.Add(cfbStream);
        }
        else if (entry.ObjectType == ObjectType.Storage)
        {
            var storage = new CfbStorage { Name = entry.Name, Parent = parent };
            parent.Children.Add(storage);
            if (entry.ChildSid != CfbSectorMarker.NoEntry)
                BuildTree(storage, entry.ChildSid, visited, depth + 1);
        }

        // 处理右兄弟
        if (entry.RightSibSid != CfbSectorMarker.NoEntry)
            BuildTree(parent, entry.RightSibSid, visited, depth + 1);
    }

    /// <summary>读取一个流条目的数据（自动识别普通流或迷你流）</summary>
    private Byte[] ReadStreamData(CfbDirectoryEntry entry)
    {
        if (entry.StreamSize < 0 || entry.StreamSize > Int32.MaxValue || entry.StreamSize > OfficeLimits.MaxCfbStreamBytes)
            throw new InvalidDataException(L.T("The CFB stream exceeds the resource limit."));

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
        // 先遍历 FAT 链计算扇区数量
        var sectorSize = _header.SectorSize;
        var sectorCount = 0;
        var sid = startSid;
        while (sid != CfbSectorMarker.EndOfChain && sid != CfbSectorMarker.FreeSect && sid >= 0)
        {
            sectorCount++;
            if (sectorCount > _fat.Length + 1)
                throw new InvalidDataException(L.T("The CFB FAT chain contains a cycle."));
            if (sid >= _fat.Length) break;
            sid = _fat[sid];
        }
        if (sectorCount == 0) return [];

        // 分配结果缓冲区并通过 ReadAt 按扇区读取
        var total = (Int64)sectorCount * sectorSize;
        var resultSize = expectedSize > 0 && expectedSize < total ? expectedSize : total;
        if (resultSize > Int32.MaxValue || resultSize > OfficeLimits.MaxCfbStreamBytes)
            throw new InvalidDataException(L.T("The CFB stream exceeds the resource limit."));

        var limit = (Int32)resultSize;
        var result = new Byte[limit];

        var pos = 0;
        sid = startSid;
        while (sid != CfbSectorMarker.EndOfChain && sid != CfbSectorMarker.FreeSect && sid >= 0 && pos < limit)
        {
            var copyLen = Math.Min(sectorSize, limit - pos);
            ReadAt((Int64)(sid + 1) * sectorSize, result, pos, copyLen);
            pos += copyLen;
            if (sid >= _fat.Length) break;
            sid = _fat[sid];
        }

        return result;
    }

    /// <summary>从流的指定偏移处读取确定数量的字节</summary>
    /// <param name="offset">文件偏移</param>
    /// <param name="dest">目标缓冲区</param>
    /// <param name="destOffset">写入起始偏移</param>
    /// <param name="count">读取字节数</param>
    private void ReadAt(Int64 offset, Byte[] dest, Int32 destOffset, Int32 count)
    {
        _fs.Seek(offset, SeekOrigin.Begin);
        var read = 0;
        while (read < count)
        {
            var r = _fs.Read(dest, destOffset + read, count - read);
            if (r == 0) throw new EndOfStreamException(L.T("Unexpected end of CFB stream."));
            read += r;
        }
    }

    /// <summary>通过 Mini FAT 链读取迷你流数据</summary>
    private Byte[] ReadMiniStreamChain(Int32 startMiniSid, Int32 expectedSize)
    {
        var miniSectorSize = _header.MiniSectorSize;

        // 先遍历 Mini FAT 链：检测环，并统计落在迷你流范围内的有效扇区
        var sectorCount = 0;
        var steps = 0;
        var msid = startMiniSid;
        while (msid != CfbSectorMarker.EndOfChain && msid != CfbSectorMarker.FreeSect && msid >= 0)
        {
            if (++steps > _miniFat.Length + 1)
                throw new InvalidDataException(L.T("The CFB mini FAT chain contains a cycle."));
            if ((Int64)msid * miniSectorSize + miniSectorSize <= _miniStream.Length)
                sectorCount++;
            if (msid >= _miniFat.Length) break;
            msid = _miniFat[msid];
        }
        if (sectorCount == 0) return [];

        // 分配结果缓冲区并直接从迷你流复制
        var total = (Int64)sectorCount * miniSectorSize;
        var resultSize = expectedSize > 0 && expectedSize < total ? expectedSize : total;
        if (resultSize > Int32.MaxValue || resultSize > OfficeLimits.MaxCfbStreamBytes)
            throw new InvalidDataException(L.T("The CFB mini stream exceeds the resource limit."));

        var limit = (Int32)resultSize;
        var result = new Byte[limit];

        var pos = 0;
        msid = startMiniSid;
        while (msid != CfbSectorMarker.EndOfChain && msid != CfbSectorMarker.FreeSect && msid >= 0 && pos < limit)
        {
            var offset = msid * miniSectorSize;
            if (offset + miniSectorSize <= _miniStream.Length)
            {
                var copyLen = Math.Min(miniSectorSize, limit - pos);
                Array.Copy(_miniStream, offset, result, pos, copyLen);
                pos += copyLen;
            }
            if (msid >= _miniFat.Length) break;
            msid = _miniFat[msid];
        }

        return result;
    }

    #endregion
}
