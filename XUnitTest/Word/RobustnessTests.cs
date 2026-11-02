using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>异常输入健壮性测试：损坏/伪装/截断文件不得挂起，且要给出可诊断的错误</summary>
public class RobustnessTests
{
    private static String TempFile(String ext, Byte[]? content = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"Robust_{Guid.NewGuid():N}.{ext}");
        File.WriteAllBytes(path, content ?? []);
        return path;
    }

    /// <summary>构造一个最小合法 docx（可选截断到前 n 字节）</summary>
    private static Byte[] BuildMinimalDocx(Int32 truncateTo = -1)
    {
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            using (var sw = new StreamWriter(za.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false)))
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p></w:body></w:document>");
        }
        var bytes = ms.ToArray();
        return truncateTo > 0 && truncateTo < bytes.Length ? bytes[..truncateTo] : bytes;
    }

    [Fact(DisplayName = "健壮性-非 zip 内容伪装成 docx 时抛异常而非挂起")]
    public void NotAZip_Throws()
    {
        var path = TempFile("docx", Encoding.UTF8.GetBytes("这不是一个 zip 文件"));
        try
        {
            Assert.ThrowsAny<Exception>(() =>
            {
                using var reader = new WordReader(path);
                reader.ReadDocument();
            });
        }
        finally { File.Delete(path); }
    }

    [Fact(DisplayName = "健壮性-截断的 docx 抛异常而非返回半成品")]
    public void Truncated_Throws()
    {
        var path = TempFile("docx", BuildMinimalDocx(truncateTo: 40));
        try
        {
            Assert.ThrowsAny<Exception>(() =>
            {
                using var reader = new WordReader(path);
                reader.ReadDocument();
            });
        }
        finally { File.Delete(path); }
    }

    [Fact(DisplayName = "健壮性-空文件抛异常")]
    public void EmptyFile_Throws()
    {
        var path = TempFile("docx", []);
        try
        {
            Assert.ThrowsAny<Exception>(() =>
            {
                using var reader = new WordReader(path);
                reader.ReadDocument();
            });
        }
        finally { File.Delete(path); }
    }

    [Fact(DisplayName = "健壮性-损坏文件转 Markdown 返回 null 并记录 LastError")]
    public void Broken_Markdown_ReturnsNull_WithError()
    {
        // 能打开成 zip，但缺少 word/document.xml
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        using (var sw = new StreamWriter(za.CreateEntry("dummy.txt").Open(), new UTF8Encoding(false)))
            sw.Write("no document part");

        ms.Position = 0;
        using var reader = new WordReader(ms);
        var md = reader.ToMarkdownDocument();

        // 契约：不抛异常，返回 null，并可通过 LastError 区分“空文档”和“解析失败”
        Assert.Null(md);
        Assert.False(String.IsNullOrEmpty(reader.LastError), "解析失败应记录 LastError");
    }

    [Fact(DisplayName = "健壮性-不存在的文件路径立即报 FileNotFound")]
    public void MissingFile_ThrowsFileNotFound()
    {
        var path = Path.Combine(Path.GetTempPath(), $"NoSuch_{Guid.NewGuid():N}.docx");
        Assert.ThrowsAny<IOException>(() => new WordReader(path));
    }

    [Fact(DisplayName = "健壮性-损坏的补丁 JSON 抛 FormatException 而非静默空补丁")]
    public void BrokenPatchJson_Throws()
    {
        Assert.Throws<FormatException>(() => DocxPatch.ParseJson("{not json"));
        Assert.Throws<FormatException>(() => DocxPatch.ParseJson("[{\"op\":\"set_text\"}]"));   // 顶层必须是对象
        Assert.Throws<FormatException>(() => DocxPatch.ParseJson("{\"ops\":\"should-be-array\"}"));
    }
}
