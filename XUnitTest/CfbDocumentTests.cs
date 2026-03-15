using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>CFB/OLE2 复合文档单元测试</summary>
public class CfbDocumentTests
{
    // ─── 辅助 ──────────────────────────────────────────────────────────────
    private static Byte[] MakeBytes(Int32 size, Byte fill = 0xAB)
    {
        var buf = new Byte[size];
        for (var i = 0; i < size; i++) buf[i] = (Byte)(fill ^ (i & 0xFF));
        return buf;
    }

    private static void AssertBytesEqual(Byte[] expected, Byte[]? actual, String msg = "")
    {
        Assert.NotNull(actual);
        Assert.Equal(expected.Length, actual!.Length);
        for (var i = 0; i < expected.Length; i++)
            Assert.True(expected[i] == actual[i], $"{msg} byte[{i}] expected {expected[i]:X2} got {actual[i]:X2}");
    }

    // ─── 基础往返测试 ───────────────────────────────────────────────────────

    [Fact, DisplayName("空文档（仅根存储）写入后可正确读回")]
    public void EmptyDocument_RoundTrip()
    {
        var doc = new CfbDocument();
        var bytes = doc.ToBytes();

        Assert.True(bytes.Length >= 512, "至少包含文件头扇区");
        Assert.Equal(0xD0, bytes[0]);
        Assert.Equal(0xCF, bytes[1]);
        Assert.Equal(0x11, bytes[2]);
        Assert.Equal(0xE0, bytes[3]);

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        Assert.NotNull(doc2.Root);
        Assert.Empty(doc2.Root.Children);
    }

    [Fact, DisplayName("小流（< 4096 字节，使用迷你流）写入读回数据完整")]
    public void SmallStream_MiniStream_RoundTrip()
    {
        var data = MakeBytes(256);

        var doc = new CfbDocument();
        doc.Root.AddStream("TestStream", data);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        var stream = doc2.Root.GetStream("TestStream");
        Assert.NotNull(stream);
        AssertBytesEqual(data, stream!.Data, "SmallStream");
    }

    [Fact, DisplayName("大流（>= 4096 字节，使用正常 FAT）写入读回数据完整")]
    public void LargeStream_FatChain_RoundTrip()
    {
        var data = MakeBytes(8192);

        var doc = new CfbDocument();
        doc.Root.AddStream("BigStream", data);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        var stream = doc2.Root.GetStream("BigStream");
        Assert.NotNull(stream);
        AssertBytesEqual(data, stream!.Data, "LargeStream");
    }

    [Fact, DisplayName("恰好 4096 字节的流边界值测试")]
    public void BoundaryStream_Exactly4096_RoundTrip()
    {
        var data = MakeBytes(4096);

        var doc = new CfbDocument();
        doc.Root.AddStream("BoundaryStream", data);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(data, doc2.Root.GetStream("BoundaryStream")?.Data, "Boundary4096");
    }

    [Fact, DisplayName("根存储中多个流同时写入读回正确")]
    public void MultipleStreams_RootStorage_RoundTrip()
    {
        var d1 = MakeBytes(100);
        var d2 = MakeBytes(5000, 0x55);
        var d3 = Encoding.UTF8.GetBytes("Hello CFB World");

        var doc = new CfbDocument();
        doc.Root.AddStream("Stream1", d1);
        doc.Root.AddStream("Stream2", d2);
        doc.Root.AddStream("Stream3", d3);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(d1, doc2.Root.GetStream("Stream1")?.Data, "Stream1");
        AssertBytesEqual(d2, doc2.Root.GetStream("Stream2")?.Data, "Stream2");
        AssertBytesEqual(d3, doc2.Root.GetStream("Stream3")?.Data, "Stream3");
    }

    [Fact, DisplayName("嵌套存储（Storage/Stream 层级）往返正确")]
    public void NestedStorage_RoundTrip()
    {
        var inner = MakeBytes(300);

        var doc = new CfbDocument();
        var sub = doc.Root.AddStorage("SubStorage");
        sub.AddStream("InnerStream", inner);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        var sub2 = doc2.Root.GetStorage("SubStorage");
        Assert.NotNull(sub2);
        var stream = sub2!.GetStream("InnerStream");
        Assert.NotNull(stream);
        AssertBytesEqual(inner, stream!.Data, "NestedStream");
    }

    [Fact, DisplayName("多层嵌套存储（三层）往返正确")]
    public void DeepNestedStorage_3Level_RoundTrip()
    {
        var data = Encoding.UTF8.GetBytes("deep content");

        var doc = new CfbDocument();
        var l1 = doc.Root.AddStorage("Level1");
        var l2 = l1.AddStorage("Level2");
        l2.AddStream("DeepStream", data);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        var stream = doc2.Root.GetStorage("Level1")?.GetStorage("Level2")?.GetStream("DeepStream");
        Assert.NotNull(stream);
        AssertBytesEqual(data, stream!.Data, "DeepStream");
    }

    [Fact, DisplayName("空字节数组流（长度 0）往返正确")]
    public void EmptyStream_ZeroLength_RoundTrip()
    {
        var doc = new CfbDocument();
        doc.Root.AddStream("EmptyStream", []);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        var stream = doc2.Root.GetStream("EmptyStream");
        Assert.NotNull(stream);
        Assert.Empty(stream!.Data ?? []);
    }

    // ─── 路径 API 测试 ──────────────────────────────────────────────────────

    [Fact, DisplayName("GetStreamData 单层路径返回正确数据")]
    public void GetStreamData_TopLevelPath_ReturnsData()
    {
        var data = MakeBytes(512);
        var doc = new CfbDocument();
        doc.Root.AddStream("Workbook", data);

        var bytes = doc.ToBytes();
        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(data, doc2.GetStreamData("Workbook"), "TopLevel");
    }

    [Fact, DisplayName("GetStreamData 嵌套路径（/ 分隔）返回正确数据")]
    public void GetStreamData_NestedPath_ReturnsData()
    {
        var data = MakeBytes(128);
        var doc = new CfbDocument();
        doc.Root.AddStorage("Storage1").AddStream("Sub", data);

        var bytes = doc.ToBytes();
        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(data, doc2.GetStreamData("Storage1/Sub"), "NestedPath");
    }

    [Fact, DisplayName("GetStreamData 路径不存在返回 null")]
    public void GetStreamData_NotFound_ReturnsNull()
    {
        var doc = new CfbDocument();
        doc.Root.AddStream("Exists", [0x01]);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        Assert.Null(doc2.GetStreamData("DoesNotExist"));
        Assert.Null(doc2.GetStreamData("Missing/Stream"));
    }

    [Fact, DisplayName("PutStream 写入后使用 GetStreamData 可读回")]
    public void PutStream_ThenGet_RoundTrip()
    {
        var data = MakeBytes(200);
        var doc = new CfbDocument();
        doc.PutStream("MyStream", data);

        var bytes = doc.ToBytes();
        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(data, doc2.GetStreamData("MyStream"), "PutGet");
    }

    [Fact, DisplayName("PutStream 写入同一路径两次，第二次覆盖数据")]
    public void PutStream_Overwrite_ReplacesData()
    {
        var data1 = MakeBytes(100);
        var data2 = MakeBytes(200, 0xCC);

        var doc = new CfbDocument();
        doc.PutStream("Stream", data1);
        doc.PutStream("Stream", data2); // overwrite

        var bytes = doc.ToBytes();
        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(data2, doc2.GetStreamData("Stream"), "Overwrite");
    }

    [Fact, DisplayName("PutStream 嵌套路径自动创建父存储")]
    public void PutStream_NestedPath_AutoCreatesStorage()
    {
        var data = MakeBytes(64);
        var doc = new CfbDocument();
        doc.PutStream("Storage1/Stream1", data);

        var bytes = doc.ToBytes();
        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(data, doc2.GetStreamData("Storage1/Stream1"), "AutoCreateStorage");
    }

    // ─── 大文件测试 ─────────────────────────────────────────────────────────

    [Fact, DisplayName("超大流（512 KB）FAT 链完整性测试")]
    public void VeryLargeStream_HalfMegabyte_RoundTrip()
    {
        var data = MakeBytes(512 * 1024, 0x77);

        var doc = new CfbDocument();
        doc.Root.AddStream("HugeStream", data);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(data, doc2.Root.GetStream("HugeStream")?.Data, "512KB");
    }

    [Fact, DisplayName("混合大小流（小+大）共存往返正确")]
    public void MixedStreams_SmallAndLarge_RoundTrip()
    {
        var small = MakeBytes(500);
        var large = MakeBytes(10000, 0x33);

        var doc = new CfbDocument();
        doc.Root.AddStream("Small", small);
        doc.Root.AddStream("Large", large);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        AssertBytesEqual(small, doc2.Root.GetStream("Small")?.Data, "Small");
        AssertBytesEqual(large, doc2.Root.GetStream("Large")?.Data, "Large");
    }

    // ─── ToBytes / Stream / File API 测试 ──────────────────────────────────

    [Fact, DisplayName("ToBytes 与 Save(Stream) 产生相同字节内容")]
    public void ToBytes_And_SaveStream_Identical()
    {
        var data = MakeBytes(1024);
        var doc = new CfbDocument();
        doc.Root.AddStream("S", data);

        var bytes1 = doc.ToBytes();

        using var ms = new MemoryStream();
        doc.Save(ms);
        var bytes2 = ms.ToArray();

        Assert.Equal(bytes1.Length, bytes2.Length);
        AssertBytesEqual(bytes1, bytes2, "ToBytesVsSaveStream");
    }

    [Fact, DisplayName("Open(Stream, leaveOpen:true) 不关闭流")]
    public void Open_LeaveOpen_True_StreamRemainsOpen()
    {
        var data = MakeBytes(100);
        var doc = new CfbDocument();
        doc.Root.AddStream("S", data);
        var bytes = doc.ToBytes();

        var ms = new MemoryStream(bytes);
        var _ = CfbDocument.Open(ms, leaveOpen: true);
        Assert.True(ms.CanRead, "Stream should remain open");
    }

    // ─── 文件系统往返测试 ───────────────────────────────────────────────────

    [Fact, DisplayName("保存到文件再打开验证往返完整")]
    public void SaveToFile_ThenOpen_RoundTrip()
    {
        var data = MakeBytes(2048, 0xF1);
        var path = Path.Combine(Path.GetTempPath(), $"cfb_test_{Guid.NewGuid():N}.ole");
        try
        {
            var doc = new CfbDocument();
            doc.Root.AddStream("FileStream", data);
            doc.Save(path);

            Assert.True(File.Exists(path));
            var doc2 = CfbDocument.Open(path);
            AssertBytesEqual(data, doc2.Root.GetStream("FileStream")?.Data, "FileRoundtrip");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ─── 错误处理测试 ───────────────────────────────────────────────────────

    [Fact, DisplayName("解析无效魔数抛出 InvalidDataException")]
    public void Open_InvalidMagic_ThrowsInvalidDataException()
    {
        var buf = new Byte[512];
        // not a CFB file
        Encoding.ASCII.GetBytes("PK\x03\x04").CopyTo(buf, 0);
        Assert.Throws<InvalidDataException>(() => CfbDocument.Open(new MemoryStream(buf)));
    }

    [Fact, DisplayName("Open 不可寻址流抛出 ArgumentException")]
    public void Open_NonSeekableStream_ThrowsArgumentException()
    {
        // Use a non-seekable stream wrapper
        Assert.Throws<ArgumentException>(() => CfbDocument.Open(new NonSeekableStream()));
    }

    // ─── 存储/流查找 API 测试 ──────────────────────────────────────────────

    [Fact, DisplayName("GetStorage 不存在返回 null")]
    public void GetStorage_NotFound_ReturnsNull()
    {
        var doc = new CfbDocument();
        Assert.Null(doc.Root.GetStorage("NoSuchStorage"));
    }

    [Fact, DisplayName("GetStream 不存在返回 null")]
    public void GetStream_NotFound_ReturnsNull()
    {
        var doc = new CfbDocument();
        Assert.Null(doc.Root.GetStream("NoSuchStream"));
    }

    [Fact, DisplayName("Storages 属性列出所有子存储")]
    public void Storages_Property_ListsSubStorages()
    {
        var doc = new CfbDocument();
        doc.Root.AddStorage("A");
        doc.Root.AddStorage("B");
        doc.Root.AddStream("C", [0x01]);

        var storages = doc.Root.Storages.ToList();
        Assert.Equal(2, storages.Count);
        Assert.True(storages.Any(s => s.Name == "A"), "Should contain A");
        Assert.True(storages.Any(s => s.Name == "B"), "Should contain B");
    }

    [Fact, DisplayName("Streams 属性列出所有子流")]
    public void Streams_Property_ListsChildStreams()
    {
        var doc = new CfbDocument();
        doc.Root.AddStorage("Sub");
        doc.Root.AddStream("S1", [0x01]);
        doc.Root.AddStream("S2", [0x02]);

        var streams = doc.Root.Streams.ToList();
        Assert.Equal(2, streams.Count);
        Assert.True(streams.Any(s => s.Name == "S1"), "Should contain S1");
        Assert.True(streams.Any(s => s.Name == "S2"), "Should contain S2");
    }

    [Fact, DisplayName("CfbStream.OpenRead 返回可读流，内容正确")]
    public void CfbStream_OpenRead_ReturnsReadableStream()
    {
        var data = MakeBytes(200);
        var doc = new CfbDocument();
        doc.Root.AddStream("RS", data);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        var stream = doc2.Root.GetStream("RS")!;
        using var rs = stream.OpenRead();
        var read = new Byte[data.Length];
        var total = 0;
        int n;
        while ((n = rs.Read(read, total, read.Length - total)) > 0) total += n;
        Assert.Equal(data.Length, total);
        AssertBytesEqual(data, read, "OpenRead");
    }

    [Fact, DisplayName("CfbStream.GetBytes 返回数据副本")]
    public void CfbStream_GetBytes_ReturnsCopy()
    {
        var data = MakeBytes(64);
        var doc = new CfbDocument();
        doc.Root.AddStream("B", data);
        var bytes = doc.ToBytes();

        var doc2 = CfbDocument.Open(new MemoryStream(bytes));
        var stream = doc2.Root.GetStream("B")!;
        var copy = stream.GetBytes(0, data.Length);
        AssertBytesEqual(data, copy, "GetBytes");
        Assert.False(ReferenceEquals(stream.Data, copy), "GetBytes 应返回副本而非原始数组");
    }

    // ─── 辅助类型 ──────────────────────────────────────────────────────────

    private sealed class NonSeekableStream : Stream
    {
        public override Boolean CanRead => true;
        public override Boolean CanSeek => false;
        public override Boolean CanWrite => false;
        public override Int64 Length => throw new NotSupportedException();
        public override Int64 Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override Int32 Read(Byte[] buffer, Int32 offset, Int32 count) => 0;
        public override Int64 Seek(Int64 offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(Int64 value) => throw new NotSupportedException();
        public override void Write(Byte[] buffer, Int32 offset, Int32 count) => throw new NotSupportedException();
    }
}
