namespace NewLife.Office;

/// <summary>CFB 扇区 ID 特殊标记值</summary>
public static class CfbSectorMarker
{
    /// <summary>不可用/空扇区（0xFFFFFFFF）</summary>
    public const Int32 FreeSect = unchecked((Int32)0xFFFFFFFF);

    /// <summary>FAT 链尾标记（0xFFFFFFFE）</summary>
    public const Int32 EndOfChain = unchecked((Int32)0xFFFFFFFE);

    /// <summary>FAT 扇区本身的标记（0xFFFFFFFD）</summary>
    public const Int32 FatSect = unchecked((Int32)0xFFFFFFFD);

    /// <summary>DIFAT 扇区标记（0xFFFFFFFC）</summary>
    public const Int32 DifSect = unchecked((Int32)0xFFFFFFFC);

    /// <summary>不存在的 Directory Entry SID（0xFFFFFFFF）</summary>
    public const Int32 NoEntry = unchecked((Int32)0xFFFFFFFF);
}
