using System.Text;
using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>CFB 目录项（128 字节）</summary>
/// <remarks>
/// 每个目录项描述一个存储或流节点，通过左兄弟/右兄弟/子节点 SID 组成红黑树。
/// 根存储的 SID 为 0，占据目录第一个槽位。
/// </remarks>
internal sealed class CfbDirectoryEntry
{
    #region 属性
    /// <summary>条目 SID（在目录数组中的索引）</summary>
    public Int32 Sid { get; set; }

    /// <summary>条目名称（UTF-16LE，最多 31 字符）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>对象类型</summary>
    public CfbObjectType ObjectType { get; set; }

    /// <summary>红黑树颜色</summary>
    public CfbColorFlag ColorFlag { get; set; }

    /// <summary>左兄弟 SID</summary>
    public Int32 LeftSibSid { get; set; }

    /// <summary>右兄弟 SID</summary>
    public Int32 RightSibSid { get; set; }

    /// <summary>第一个子节点 SID</summary>
    public Int32 ChildSid { get; set; }

    /// <summary>CLSID（根/存储节点有意义）</summary>
    public Byte[] Clsid { get; set; } = new Byte[16];

    /// <summary>流/存储数据起始扇区 ID</summary>
    public Int32 StartingSectorId { get; set; }

    /// <summary>流大小（字节）</summary>
    public Int64 StreamSize { get; set; }
    #endregion

    #region 方法
    /// <summary>从 128 字节缓冲区解析一个目录项</summary>
    /// <param name="buf">128 字节只读缓冲区</param>
    /// <param name="sid">本条目 SID</param>
    /// <returns>解析后的目录项</returns>
    public static CfbDirectoryEntry ReadFrom(ReadOnlySpan<Byte> buf, Int32 sid)
    {
        var reader = new SpanReader(buf);
        var entry = new CfbDirectoryEntry { Sid = sid };

        // 名称（64 字节，UTF-16LE，包含 null 终止符）
        var nameLen = 0;
        var nameRaw = reader.ReadBytes(64).ToArray();
        // 先读名称长度（offset 64-65）
        var nameLenBytes = reader.ReadBytes(2).ToArray();
        nameLen = nameLenBytes[0] | (nameLenBytes[1] << 8);
        if (nameLen >= 2)
        {
            // nameLen 包含 null 终止符的字节数
            var charCount = (nameLen - 2) / 2; // 去掉 null 的字符数
            if (charCount > 0 && charCount <= 31)
                entry.Name = Encoding.Unicode.GetString(nameRaw, 0, nameLen - 2);
        }

        entry.ObjectType = (CfbObjectType)reader.ReadByte();
        entry.ColorFlag = (CfbColorFlag)reader.ReadByte();
        entry.LeftSibSid = reader.ReadInt32();
        entry.RightSibSid = reader.ReadInt32();
        entry.ChildSid = reader.ReadInt32();
        entry.Clsid = reader.ReadBytes(16).ToArray();

        reader.Advance(4); // State bits
        reader.Advance(8); // Created time
        reader.Advance(8); // Modified time

        entry.StartingSectorId = reader.ReadInt32();
        var sizeLow = (UInt32)reader.ReadInt32();
        var sizeHigh = (UInt32)reader.ReadInt32();
        entry.StreamSize = (Int64)((UInt64)sizeHigh << 32 | sizeLow);

        return entry;
    }

    /// <summary>将目录项写入缓冲区</summary>
    /// <param name="buf">输出缓冲区</param>
    /// <param name="offset">写入起始偏移（默认 0）</param>
    public void WriteTo(Byte[] buf, Int32 offset = 0)
    {
        var writer = new SpanWriter(buf, offset, 128);

        // 名称（64 字节 UTF-16LE，不足补零）
        var nameBytes = Encoding.Unicode.GetBytes(Name);
        var written = Math.Min(nameBytes.Length, 62); // 最多 31 字符 = 62 字节
        // SpanWriter.Write(byte[]) 需要完整数组，此处手动复制
        for (var i = 0; i < written; i++) writer.Write(nameBytes[i]);
        // 补齐到 64 字节
        for (var i = written; i < 64; i++) writer.Write((Byte)0);

        // 名称长度（含 null 终止符）
        var nameLen = (UInt16)(Name.Length > 0 ? (written + 2) : 0);
        writer.Write(nameLen);

        writer.Write((Byte)ObjectType);
        writer.Write((Byte)ColorFlag);
        writer.Write(LeftSibSid);
        writer.Write(RightSibSid);
        writer.Write(ChildSid);

        // CLSID (16 bytes)
        for (var i = 0; i < 16; i++)
        {
            writer.Write(i < Clsid.Length ? Clsid[i] : (Byte)0);
        }

        writer.Write(0);   // State bits
        writer.Write(0L);  // Created time
        writer.Write(0L);  // Modified time

        writer.Write(StartingSectorId);
        writer.Write((UInt32)(StreamSize & 0xFFFFFFFF));  // Size low
        writer.Write((UInt32)(StreamSize >> 32));          // Size high
    }
    #endregion
}
