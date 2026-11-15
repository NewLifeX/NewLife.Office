using System.IO.Compression;
using System.Text;

namespace NewLife.Office;

/// <summary>带资源上限的流读取辅助</summary>
internal static class OfficeStreams
{
    /// <summary>打开 ZIP 条目并限制解压后的字节数</summary>
    /// <param name="entry">ZIP 条目</param>
    /// <param name="limit">上限，默认 <see cref="OfficeLimits.MaxZipEntryBytes"/></param>
    /// <returns>受限读取流，调用方负责释放</returns>
    public static Stream OpenEntry(ZipArchiveEntry entry, Int64? limit = null)
    {
        if (entry == null) throw new ArgumentNullException(nameof(entry));

        var max = limit ?? OfficeLimits.MaxZipEntryBytes;
        if (entry.Length > max)
            throw new InvalidDataException(L.T("ZIP entry '{0}' exceeds the resource limit ({1} bytes).", entry.FullName, max));

        return new LimitedReadStream(entry.Open(), max);
    }

    /// <summary>读取 ZIP 条目全部字节并限制解压后的字节数</summary>
    /// <param name="entry">ZIP 条目</param>
    /// <param name="limit">上限，默认 <see cref="OfficeLimits.MaxZipEntryBytes"/></param>
    /// <returns>条目数据</returns>
    public static Byte[] ReadAllBytes(ZipArchiveEntry entry, Int64? limit = null)
    {
        using var stream = OpenEntry(entry, limit);
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>读取 ZIP 条目全部文本并限制解压后的字节数</summary>
    /// <param name="entry">ZIP 条目</param>
    /// <param name="encoding">文本编码，默认 UTF-8</param>
    /// <param name="limit">上限，默认 <see cref="OfficeLimits.MaxZipEntryBytes"/></param>
    /// <returns>条目文本</returns>
    public static String ReadAllText(ZipArchiveEntry entry, Encoding? encoding = null, Int64? limit = null)
    {
        var bytes = ReadAllBytes(entry, limit);
        return (encoding ?? Encoding.UTF8).GetString(bytes).TrimStart('\uFEFF');
    }

    /// <summary>校验 ZIP 条目数量，超限抛异常</summary>
    /// <param name="archive">ZIP 归档</param>
    public static void CheckEntryCount(ZipArchive archive)
    {
        if (archive == null) return;
        if (archive.Entries.Count > OfficeLimits.MaxZipEntries)
            throw new InvalidDataException(L.T("The archive contains too many entries ({0}).", archive.Entries.Count));
    }

    /// <summary>校验 ZIP 归档的条目数量、单条目大小与解压总大小</summary>
    /// <param name="archive">ZIP 归档</param>
    public static void ValidateArchive(ZipArchive archive)
    {
        if (archive == null) return;

        var entries = archive.Entries;
        CheckEntryCount(archive);

        var total = 0L;
        foreach (var entry in entries)
        {
            // 目录条目长度为 0，但同样占用条目配额，已在上一步统一检查
            var length = entry.Length;
            if (length > OfficeLimits.MaxZipEntryBytes)
                throw new InvalidDataException(L.T("ZIP entry '{0}' exceeds the resource limit ({1} bytes).", entry.FullName, OfficeLimits.MaxZipEntryBytes));

            total += length;
            if (total > OfficeLimits.MaxZipTotalBytes)
                throw new InvalidDataException(L.T("The decompressed archive exceeds the resource limit ({0} bytes).", OfficeLimits.MaxZipTotalBytes));
        }
    }

    /// <summary>以只读方式打开 ZIP 归档并校验资源上限</summary>
    /// <param name="stream">归档数据流</param>
    /// <param name="leaveOpen">释放归档时是否保留数据流</param>
    /// <param name="entryNameEncoding">条目名编码</param>
    /// <returns>已校验的归档</returns>
    public static ZipArchive OpenReadArchive(Stream stream, Boolean leaveOpen = false, Encoding? entryNameEncoding = null)
    {
        var zip = entryNameEncoding == null
            ? new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen)
            : new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen, entryNameEncoding);
        try
        {
            ValidateArchive(zip);
        }
        catch
        {
            zip.Dispose();
            throw;
        }

        return zip;
    }

    /// <summary>把输入流复制到内存流，并限制最大字节数</summary>
    /// <param name="stream">输入流</param>
    /// <param name="limit">上限</param>
    /// <returns>内存流，位置已复位到 0</returns>
    public static MemoryStream CopyToMemory(Stream stream, Int64 limit)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        using var limited = new LimitedReadStream(stream, limit, leaveOpen: true);
        var ms = new MemoryStream();
        limited.CopyTo(ms);
        ms.Position = 0;
        return ms;
    }

    /// <summary>限制读取字节数的包装流</summary>
    private sealed class LimitedReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly Int64 _limit;
        private readonly Boolean _leaveOpen;
        private Int64 _read;

        public LimitedReadStream(Stream inner, Int64 limit, Boolean leaveOpen = false)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _limit = limit > 0 ? limit : throw new ArgumentOutOfRangeException(nameof(limit));
            _leaveOpen = leaveOpen;
        }

        public override Boolean CanRead => _inner.CanRead;
        public override Boolean CanSeek => false;
        public override Boolean CanWrite => false;
        public override Int64 Length => throw new NotSupportedException();
        public override Int64 Position
        {
            get => _read;
            set => throw new NotSupportedException();
        }

        public override Int32 Read(Byte[] buffer, Int32 offset, Int32 count)
        {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset + count > buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));

            var remain = _limit - _read;
            if (remain <= 0)
            {
                // 到达上限后必须区分“恰好结束”和“还有数据”，避免误伤合法文件
                var probe = new Byte[1];
                var extra = _inner.Read(probe, 0, 1);
                if (extra == 0) return 0;

                throw new InvalidDataException(L.T("The decompressed data exceeds the resource limit ({0} bytes).", _limit));
            }

            if (count > remain) count = (Int32)remain;
            var n = _inner.Read(buffer, offset, count);
            _read += n;
            return n;
        }

        public override void Flush() { }
        public override Int64 Seek(Int64 offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(Int64 value) => throw new NotSupportedException();
        public override void Write(Byte[] buffer, Int32 offset, Int32 count) => throw new NotSupportedException();

        protected override void Dispose(Boolean disposing)
        {
            if (disposing && !_leaveOpen) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
