using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>CFB 文件头（512 字节）</summary>
/// <remarks>
/// 解析 Compound File Binary 的固定 512 字节文件头，包含 FAT/DIFAT/目录扇区位置信息。
/// 版本 3 扇区大小 512 字节，版本 4 扇区大小 4096 字节。
/// </remarks>
internal sealed class CfbHeader
{
    #region 属性
    /// <summary>主版本号（3 或 4）</summary>
    public UInt16 MajorVersion { get; private set; }

    /// <summary>扇区大小（字节），等于 2^SectorSizePow</summary>
    public Int32 SectorSize { get; private set; }

    /// <summary>迷你扇区大小（字节），等于 2^MiniSectorSizePow</summary>
    public Int32 MiniSectorSize { get; private set; }

    /// <summary>FAT 扇区数量</summary>
    public Int32 FatSectorCount { get; private set; }

    /// <summary>目录起始扇区 ID</summary>
    public Int32 FirstDirSectorId { get; private set; }

    /// <summary>迷你流截断大小（小于此值的流使用迷你流，默认 4096）</summary>
    public Int32 MiniStreamCutoff { get; private set; }

    /// <summary>第一个迷你 FAT 扇区 ID</summary>
    public Int32 FirstMiniFatSectorId { get; private set; }

    /// <summary>迷你 FAT 扇区数量</summary>
    public Int32 MiniFatSectorCount { get; private set; }

    /// <summary>第一个 DIFAT 扇区 ID（无则为 EndOfChain）</summary>
    public Int32 FirstDifatSectorId { get; private set; }

    /// <summary>DIFAT 扇区数量</summary>
    public Int32 DifatSectorCount { get; private set; }

    /// <summary>头部内嵌的前 109 个 FAT 扇区 ID 数组（DIFAT 数组）</summary>
    public Int32[] DifatArray { get; private set; } = [];
    #endregion

    #region 静态
    /// <summary>CFB 魔数签名</summary>
    private static readonly Byte[] MagicBytes = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <summary>从二进制缓冲区解析文件头</summary>
    /// <param name="buf">至少 512 字节的缓冲区</param>
    /// <returns>解析后的文件头对象</returns>
    public static CfbHeader ReadFrom(Byte[] buf)
    {
        var reader = new SpanReader(buf, 0, 512);
        var header = new CfbHeader();

        // 校验魔数
        var magic = reader.ReadBytes(8).ToArray();
        for (var i = 0; i < 8; i++)
        {
            if (magic[i] != MagicBytes[i])
                throw new InvalidDataException("Not a valid CFB file: magic number mismatch.");
        }

        reader.Advance(16); // Reserved CLSID
        reader.Advance(2);  // Minor version
        header.MajorVersion = reader.ReadUInt16();
        reader.Advance(2);  // Byte order mark (0xFFFE)

        var sectorSizePow = reader.ReadUInt16();
        header.SectorSize = 1 << sectorSizePow;

        var miniSectorSizePow = reader.ReadUInt16();
        header.MiniSectorSize = 1 << miniSectorSizePow;

        reader.Advance(6);  // Reserved

        reader.ReadInt32(); // Dir sector count (0 for v3)
        header.FatSectorCount = reader.ReadInt32();
        header.FirstDirSectorId = reader.ReadInt32();
        reader.Advance(4);  // Transaction signature

        header.MiniStreamCutoff = reader.ReadInt32();
        header.FirstMiniFatSectorId = reader.ReadInt32();
        header.MiniFatSectorCount = reader.ReadInt32();
        header.FirstDifatSectorId = reader.ReadInt32();
        header.DifatSectorCount = reader.ReadInt32();

        // 读取头部内嵌的 109 个 DIFAT 条目
        var difat = new Int32[109];
        for (var i = 0; i < 109; i++)
            difat[i] = reader.ReadInt32();
        header.DifatArray = difat;

        return header;
    }

    /// <summary>将文件头写入 512 字节缓冲区</summary>
    /// <param name="buf">输出缓冲区，长度至少 512</param>
    /// <param name="fatSectorCount">FAT 扇区数量</param>
    /// <param name="firstDirSector">第一个目录扇区 ID</param>
    /// <param name="firstMiniFat">第一个迷你 FAT 扇区 ID</param>
    /// <param name="miniFatCount">迷你 FAT 扇区数量</param>
    /// <param name="difatEntries">FAT 扇区 ID 列表（填入 DIFAT 数组）</param>
    public static void WriteTo(Byte[] buf, Int32 fatSectorCount, Int32 firstDirSector,
        Int32 firstMiniFat, Int32 miniFatCount, Int32[] difatEntries)
    {
        var writer = new SpanWriter(buf, 0, 512);

        // 魔数
        writer.Write((Byte)0xD0); writer.Write((Byte)0xCF); writer.Write((Byte)0x11); writer.Write((Byte)0xE0);
        writer.Write((Byte)0xA1); writer.Write((Byte)0xB1); writer.Write((Byte)0x1A); writer.Write((Byte)0xE1);

        // Reserved CLSID (16 bytes zero)
        for (var i = 0; i < 16; i++) writer.Write((Byte)0);

        writer.Write((UInt16)0x003E);  // Minor version
        writer.Write((UInt16)0x0003);  // Major version (3)
        writer.Write((UInt16)0xFFFE);  // Byte order (little-endian)
        writer.Write((UInt16)9);       // Sector size = 2^9 = 512
        writer.Write((UInt16)6);       // Mini sector size = 2^6 = 64

        // Reserved 6 bytes
        for (var i = 0; i < 6; i++) writer.Write((Byte)0);

        writer.Write(0);               // Dir sector count (0 for v3)
        writer.Write(fatSectorCount);  // FAT sector count
        writer.Write(firstDirSector);  // First directory sector
        writer.Write(0);               // Transaction signature

        writer.Write(4096);            // Mini stream cutoff
        writer.Write(firstMiniFat);    // First mini FAT sector
        writer.Write(miniFatCount);    // Mini FAT sector count
        writer.Write(CfbSectorMarker.EndOfChain);  // No DIFAT chain
        writer.Write(0);               // DIFAT sector count

        // DIFAT 数组（109 个条目）
        for (var i = 0; i < 109; i++)
        {
            var id = i < difatEntries.Length ? difatEntries[i] : CfbSectorMarker.FreeSect;
            writer.Write(id);
        }
    }
    #endregion
}
