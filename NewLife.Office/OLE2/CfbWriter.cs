using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>CFB（Compound File Binary）格式写入器</summary>
/// <remarks>
/// 将存储树序列化为 CFB 字节流（版本 3，512 字节扇区）。
/// 小于 4096 字节的流使用迷你扇区以节省空间，较大流使用普通 FAT 扇区。
/// </remarks>
internal sealed class CfbWriter
{
    #region 常量
    private const Int32 SectorSize = 512;
    private const Int32 MiniSectorSize = 64;
    private const Int32 MiniStreamCutoff = 4096;
    private const Int32 FatEntriesPerSector = SectorSize / 4;    // 128
    private const Int32 DirEntriesPerSector = SectorSize / 128;  // 4
    #endregion

    #region 写入
    /// <summary>将存储树序列化为 CFB 字节流</summary>
    /// <param name="root">根存储节点</param>
    /// <param name="output">输出流</param>
    public void Write(CfbStorage root, Stream output)
    {
        // 1. 收集所有流和存储，分配目录条目 SID
        var allEntries = new List<EntryInfo>();
        var rootInfo = new EntryInfo("Root Entry", CfbObjectType.RootStorage, null);
        allEntries.Add(rootInfo);
        CollectEntries(root, rootInfo, allEntries);

        // 2. 将小流放入迷你流，大流直接写入普通扇区
        var miniData = new List<Byte[]>();  // mini 流中的每个流字节
        var bigStreams = new List<EntryInfo>();

        foreach (var e in allEntries)
        {
            if (e.Type != CfbObjectType.Stream || e.Data == null) continue;
            if (e.Data.Length < MiniStreamCutoff)
                miniData.Add(e.Data);
            else
                bigStreams.Add(e);
        }

        // 3. 计算布局
        // mini 扇区分配
        var miniStreamBytes = BuildMiniStream(allEntries);
        var miniStreamSectorCount = CeilDiv(miniStreamBytes.Length, SectorSize);

        // FAT 分配：big streams + mini stream sectors + dir sectors
        // 先粗估 FAT 扇区需求，再精确分配
        var layout = AllocateSectors(allEntries, miniStreamBytes, miniStreamSectorCount);

        // 4. 构建目录条目字节数组
        var dirBytes = BuildDirectoryBytes(allEntries, layout, miniStreamBytes.Length);

        // 5. 构建 FAT 字节数组
        var fatBytes = BuildFat(layout);

        // 6. 写入文件
        WriteOutput(output, layout, fatBytes, dirBytes, miniStreamBytes, allEntries);
    }
    #endregion

    #region 内部数据
    private sealed class EntryInfo
    {
        public String Name { get; }
        public CfbObjectType Type { get; }
        public Byte[]? Data { get; }
        public EntryInfo? Parent { get; }
        public List<EntryInfo> Children { get; } = [];
        public Int32 Sid { get; set; }
        public Int32 StartSector { get; set; } = CfbSectorMarker.EndOfChain;
        public Int32 MiniStartSector { get; set; } = CfbSectorMarker.EndOfChain;

        public Int32 LeftSibSid { get; set; } = CfbSectorMarker.NoEntry;
        public Int32 RightSibSid { get; set; } = CfbSectorMarker.NoEntry;
        public Int32 ChildRootSid { get; set; } = CfbSectorMarker.NoEntry;

        public CfbColorFlag ColorFlag { get; set; } = CfbColorFlag.Black;

        public EntryInfo(String name, CfbObjectType type, EntryInfo? parent, Byte[]? data = null)
        {
            Name = name; Type = type; Parent = parent; Data = data;
        }
    }

    private sealed class SectorLayout
    {
        public List<Int32> FatSectorIds { get; } = [];
        public List<Int32> DirSectorIds { get; } = [];
        public Int32 MiniStreamStartSector { get; set; } = CfbSectorMarker.EndOfChain;
        public Int32 MiniFatStartSector { get; set; } = CfbSectorMarker.EndOfChain;
        public Int32 TotalSectors { get; set; }

        // FAT 条目数组；索引 = sectorId，值 = 下一个 sectorId 或标记
        public Int32[] FatArray { get; set; } = [];

        // Mini FAT 条目
        public Int32[] MiniFatArray { get; set; } = [];
    }
    #endregion

    #region 辅助方法
    private static void CollectEntries(CfbStorage storage, EntryInfo parent, List<EntryInfo> allEntries)
    {
        foreach (var child in storage.Children)
        {
            EntryInfo info;
            if (child is CfbStorage childStorage)
            {
                info = new EntryInfo(childStorage.Name, CfbObjectType.Storage, parent);
                allEntries.Add(info);
                parent.Children.Add(info);
                CollectEntries(childStorage, info, allEntries);
            }
            else if (child is CfbStream childStream)
            {
                info = new EntryInfo(childStream.Name, CfbObjectType.Stream, parent, childStream.Data);
                allEntries.Add(info);
                parent.Children.Add(info);
            }
            // else: ignore unknown types
        }

        // 分配 SID（遍历完所有条目后统一分配）
        for (var i = 0; i < allEntries.Count; i++)
        {
            allEntries[i].Sid = i;
        }
    }

    private static Byte[] BuildMiniStream(List<EntryInfo> allEntries)
    {
        var miniData = new List<Byte[]>();
        var miniSid = 0;

        foreach (var e in allEntries)
        {
            if (e.Type != CfbObjectType.Stream || e.Data == null || e.Data.Length >= MiniStreamCutoff) continue;

            e.MiniStartSector = miniSid;
            var paddedSize = CeilDiv(e.Data.Length, MiniSectorSize) * MiniSectorSize;
            var padded = new Byte[paddedSize];
            Array.Copy(e.Data, padded, e.Data.Length);
            miniData.Add(padded);
            miniSid += paddedSize / MiniSectorSize;
        }

        if (miniData.Count == 0) return [];

        var totalMiniBytes = miniData.Sum(d => d.Length);
        var result = new Byte[totalMiniBytes];
        var pos = 0;
        foreach (var d in miniData) { Array.Copy(d, 0, result, pos, d.Length); pos += d.Length; }
        return result;
    }

    private SectorLayout AllocateSectors(List<EntryInfo> allEntries, Byte[] miniStreamBytes, Int32 miniStreamSectorCount)
    {
        var layout = new SectorLayout();
        var nextSector = 0;

        // Assign sectors to big streams
        foreach (var e in allEntries)
        {
            if (e.Type != CfbObjectType.Stream || e.Data == null || e.Data.Length < MiniStreamCutoff) continue;
            e.StartSector = nextSector;
            var count = CeilDiv(e.Data.Length, SectorSize);
            nextSector += count;
        }

        // Assign sectors to mini stream (from root)
        if (miniStreamBytes.Length > 0)
        {
            layout.MiniStreamStartSector = nextSector;
            nextSector += miniStreamSectorCount;

            // Mini FAT sectors
            var miniSectorCount = miniStreamBytes.Length / MiniSectorSize;
            var miniFatSectorCount = CeilDiv(miniSectorCount, FatEntriesPerSector);
            layout.MiniFatStartSector = nextSector;
            nextSector += miniFatSectorCount;
        }

        // Directory sectors
        var dirSectorCount = CeilDiv(allEntries.Count, DirEntriesPerSector);
        for (var i = 0; i < dirSectorCount; i++)
        {
            layout.DirSectorIds.Add(nextSector + i);
        }
        nextSector += dirSectorCount;

        // Now compute FAT size iteratively (FAT sectors itself consume entries)
        var fatSectorCount = 1;
        while (true)
        {
            var totalEntries = nextSector + fatSectorCount;
            var needed = CeilDiv(totalEntries, FatEntriesPerSector);
            if (needed <= fatSectorCount) break;
            fatSectorCount = needed;
        }

        // FAT sectors go after directory sectors
        for (var i = 0; i < fatSectorCount; i++)
        {
            layout.FatSectorIds.Add(nextSector + i);
        }
        nextSector += fatSectorCount;

        layout.TotalSectors = nextSector;

        // Build FAT and mini FAT arrays
        BuildFatArrays(allEntries, miniStreamBytes, layout);

        return layout;
    }

    private static void BuildFatArrays(List<EntryInfo> allEntries, Byte[] miniStreamBytes, SectorLayout layout)
    {
        var fat = new Int32[layout.TotalSectors];
        for (var i = 0; i < fat.Length; i++) fat[i] = CfbSectorMarker.FreeSect;

        // Big streams
        foreach (var e in allEntries)
        {
            if (e.Type != CfbObjectType.Stream || e.Data == null || e.Data.Length < MiniStreamCutoff) continue;
            var count = CeilDiv(e.Data.Length, SectorSize);
            for (var i = 0; i < count - 1; i++)
            {
                fat[e.StartSector + i] = e.StartSector + i + 1;
            }
            fat[e.StartSector + count - 1] = CfbSectorMarker.EndOfChain;
        }

        // Mini stream sectors
        if (layout.MiniStreamStartSector != CfbSectorMarker.EndOfChain)
        {
            var count = CeilDiv(miniStreamBytes.Length, SectorSize);
            for (var i = 0; i < count - 1; i++)
            {
                fat[layout.MiniStreamStartSector + i] = layout.MiniStreamStartSector + i + 1;
            }
            fat[layout.MiniStreamStartSector + count - 1] = CfbSectorMarker.EndOfChain;
        }

        // Mini FAT sectors
        if (layout.MiniFatStartSector != CfbSectorMarker.EndOfChain)
        {
            var miniSectorCount = miniStreamBytes.Length / MiniSectorSize;
            var miniFatSectorCount = CeilDiv(miniSectorCount, FatEntriesPerSector);
            for (var i = 0; i < miniFatSectorCount - 1; i++)
            {
                fat[layout.MiniFatStartSector + i] = layout.MiniFatStartSector + i + 1;
            }
            fat[layout.MiniFatStartSector + miniFatSectorCount - 1] = CfbSectorMarker.EndOfChain;
        }

        // Directory sectors
        for (var i = 0; i < layout.DirSectorIds.Count - 1; i++)
        {
            fat[layout.DirSectorIds[i]] = layout.DirSectorIds[i + 1];
        }
        if (layout.DirSectorIds.Count > 0)
            fat[layout.DirSectorIds[layout.DirSectorIds.Count - 1]] = CfbSectorMarker.EndOfChain;

        // FAT sectors themselves
        foreach (var fs in layout.FatSectorIds)
        {
            fat[fs] = CfbSectorMarker.FatSect;
        }

        layout.FatArray = fat;

        // Build mini FAT
        if (miniStreamBytes.Length > 0)
        {
            var miniSectorCount = miniStreamBytes.Length / MiniSectorSize;
            var miniFat = new Int32[miniSectorCount];
            for (var i = 0; i < miniFat.Length; i++) miniFat[i] = CfbSectorMarker.FreeSect;

            foreach (var e in allEntries)
            {
                if (e.Type != CfbObjectType.Stream || e.Data == null || e.Data.Length >= MiniStreamCutoff || e.Data.Length == 0) continue;
                var count = CeilDiv(e.Data.Length, MiniSectorSize);
                for (var i = 0; i < count - 1; i++)
                {
                    miniFat[e.MiniStartSector + i] = e.MiniStartSector + i + 1;
                }
                miniFat[e.MiniStartSector + count - 1] = CfbSectorMarker.EndOfChain;
            }

            layout.MiniFatArray = miniFat;
        }
    }

    private static Byte[] BuildDirectoryBytes(List<EntryInfo> allEntries, SectorLayout layout, Int32 miniStreamSize)
    {
        var dirCount = CeilDiv(allEntries.Count, DirEntriesPerSector) * DirEntriesPerSector;
        var dirBytes = new Byte[dirCount * 128];

        // First pass: build sibling trees for all storages
        foreach (var e in allEntries)
        {
            if (e.Type is CfbObjectType.Storage or CfbObjectType.RootStorage && e.Children.Count > 0)
                e.ChildRootSid = BuildSiblingTree(e.Children, 0, e.Children.Count - 1);
        }

        // Second pass: build directory entry bytes
        for (var i = 0; i < allEntries.Count; i++)
        {
            var e = allEntries[i];
            var dirEntry = new CfbDirectoryEntry
            {
                Sid = e.Sid,
                Name = e.Name,
                ObjectType = e.Type,
                ColorFlag = CfbColorFlag.Black,
                LeftSibSid = e.LeftSibSid,
                RightSibSid = e.RightSibSid,
                ChildSid = e.ChildRootSid,
            };

            // Set stream data location
            if (e.Type == CfbObjectType.Stream && e.Data != null)
            {
                dirEntry.StreamSize = e.Data.Length;
                if (e.Data.Length >= MiniStreamCutoff)
                    dirEntry.StartingSectorId = e.StartSector;
                else
                    dirEntry.StartingSectorId = e.MiniStartSector;
            }

            // Root gets mini stream location and size
            if (e.Type == CfbObjectType.RootStorage)
            {
                dirEntry.StartingSectorId = layout.MiniStreamStartSector;
                dirEntry.StreamSize = miniStreamSize;
            }

            dirEntry.WriteTo(dirBytes, i * 128);
        }

        // Fill remaining slots with empty entries
        for (var i = allEntries.Count; i < dirCount; i++)
        {
            var empty = new CfbDirectoryEntry
            {
                Sid = i, Name = String.Empty, ObjectType = CfbObjectType.Empty,
                ColorFlag = CfbColorFlag.Red,
                LeftSibSid = CfbSectorMarker.NoEntry,
                RightSibSid = CfbSectorMarker.NoEntry,
                ChildSid = CfbSectorMarker.NoEntry,
                StartingSectorId = CfbSectorMarker.EndOfChain,
            };
            empty.WriteTo(dirBytes, i * 128);
        }

        return dirBytes;
    }

    /// <summary>将子节点列表构建为平衡二叉 sibling 树，同时设置各子节点的 LeftSibSid/RightSibSid，返回子根 SID</summary>
    private static Int32 BuildSiblingTree(List<EntryInfo> children, Int32 lo, Int32 hi)
    {
        if (lo > hi) return CfbSectorMarker.NoEntry;
        var mid = (lo + hi) / 2;
        var entry = children[mid];
        entry.ColorFlag = CfbColorFlag.Black;
        entry.LeftSibSid = BuildSiblingTree(children, lo, mid - 1);
        entry.RightSibSid = BuildSiblingTree(children, mid + 1, hi);
        return entry.Sid;
    }

    private static Byte[] BuildFat(SectorLayout layout)
    {
        var fatEntryCount = layout.FatSectorIds.Count * FatEntriesPerSector;
        var fatBytes = new Byte[fatEntryCount * 4];
        var writer = new SpanWriter(fatBytes, 0, fatBytes.Length);
        foreach (var entry in layout.FatArray)
        {
            writer.Write(entry);
        }
        // Pad remaining with FreeSect
        while (writer.Position < fatBytes.Length)
        {
            writer.Write(CfbSectorMarker.FreeSect);
        }
        return fatBytes;
    }

    private void WriteOutput(Stream output, SectorLayout layout, Byte[] fatBytes, Byte[] dirBytes,
        Byte[] miniStreamBytes, List<EntryInfo> allEntries)
    {
        // 计算 mini stream 实际大小（需填入根目录条目）
        var miniSectorCount = miniStreamBytes.Length > 0 ? miniStreamBytes.Length / MiniSectorSize : 0;
        var miniFatSectorCount = layout.MiniFatStartSector != CfbSectorMarker.EndOfChain
            ? CeilDiv(miniSectorCount, FatEntriesPerSector) : 0;

        // 写入文件头
        var headerBuf = new Byte[SectorSize];
        var difatEntries = layout.FatSectorIds.ToArray();
        CfbHeader.WriteTo(headerBuf,
            fatSectorCount: layout.FatSectorIds.Count,
            firstDirSector: layout.DirSectorIds.Count > 0 ? layout.DirSectorIds[0] : CfbSectorMarker.EndOfChain,
            firstMiniFat: layout.MiniFatStartSector,
            miniFatCount: miniFatSectorCount,
            difatEntries: difatEntries);
        output.Write(headerBuf, 0, SectorSize);

        // 按 sectorId 顺序写入所有扇区
        // 构建一个 sectorId → bytes 映射
        var sectors = new Dictionary<Int32, Byte[]>();

        // Big stream sectors
        foreach (var e in allEntries)
        {
            if (e.Type != CfbObjectType.Stream || e.Data == null || e.Data.Length < MiniStreamCutoff) continue;
            var count = CeilDiv(e.Data.Length, SectorSize);
            for (var i = 0; i < count; i++)
            {
                var start = i * SectorSize;
                var len = Math.Min(SectorSize, e.Data.Length - start);
                var buf = new Byte[SectorSize];
                Array.Copy(e.Data, start, buf, 0, len);
                sectors[e.StartSector + i] = buf;
            }
        }

        // Mini stream sectors
        if (miniStreamBytes.Length > 0)
        {
            var count = CeilDiv(miniStreamBytes.Length, SectorSize);
            for (var i = 0; i < count; i++)
            {
                var start = i * SectorSize;
                var len = Math.Min(SectorSize, miniStreamBytes.Length - start);
                var buf = new Byte[SectorSize];
                Array.Copy(miniStreamBytes, start, buf, 0, len);
                sectors[layout.MiniStreamStartSector + i] = buf;
            }
        }

        // Mini FAT sectors
        if (layout.MiniFatStartSector != CfbSectorMarker.EndOfChain && layout.MiniFatArray.Length > 0)
        {
            var miniFatBytes = new Byte[miniFatSectorCount * SectorSize];
            var mfw = new SpanWriter(miniFatBytes, 0, miniFatBytes.Length);
            foreach (var entry in layout.MiniFatArray) mfw.Write(entry);
            while (mfw.Position < miniFatBytes.Length) mfw.Write(CfbSectorMarker.FreeSect);

            for (var i = 0; i < miniFatSectorCount; i++)
            {
                var buf = new Byte[SectorSize];
                Array.Copy(miniFatBytes, i * SectorSize, buf, 0, SectorSize);
                sectors[layout.MiniFatStartSector + i] = buf;
            }
        }

        // Directory sectors
        for (var i = 0; i < layout.DirSectorIds.Count; i++)
        {
            var start = i * SectorSize;
            var buf = new Byte[SectorSize];
            var len = Math.Min(SectorSize, dirBytes.Length - start);
            if (len > 0) Array.Copy(dirBytes, start, buf, 0, len);
            sectors[layout.DirSectorIds[i]] = buf;
        }

        // FAT sectors
        for (var i = 0; i < layout.FatSectorIds.Count; i++)
        {
            var start = i * SectorSize;
            var buf = new Byte[SectorSize];
            var len = Math.Min(SectorSize, fatBytes.Length - start);
            if (len > 0) Array.Copy(fatBytes, start, buf, 0, len);
            sectors[layout.FatSectorIds[i]] = buf;
        }

        // 顺序写出
        for (var sid = 0; sid < layout.TotalSectors; sid++)
        {
            if (sectors.TryGetValue(sid, out var s))
                output.Write(s, 0, SectorSize);
            else
            {
                var empty = new Byte[SectorSize];
                output.Write(empty, 0, SectorSize);
            }
        }
    }

    private static Int32 CeilDiv(Int32 value, Int32 divisor) => (value + divisor - 1) / divisor;
    #endregion
}
