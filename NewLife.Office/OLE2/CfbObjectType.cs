namespace NewLife.Office;

/// <summary>CFB 目录项对象类型</summary>
public enum CfbObjectType : Byte
{
    /// <summary>空闲槽</summary>
    Empty = 0,

    /// <summary>存储对象（相当于目录）</summary>
    Storage = 1,

    /// <summary>流对象（相当于文件）</summary>
    Stream = 2,

    /// <summary>根存储</summary>
    RootStorage = 5,
}
