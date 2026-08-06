using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office;
using NewLife.Office.Pdf;
using Xunit;

namespace XUnitTest.Common;

/// <summary>OfficeFactory 内容识别测试（GEN-1，对标 anydoc）</summary>
/// <remarks>
/// 覆盖按文件头 magic bytes 识别格式（PDF/RTF/OLE2/ZIP 容器/文本类），
/// 以及扩展名错误时按内容创建读取器。
/// </remarks>
public class OfficeFactoryDetectionTests
{
    private static String BinDir => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, ".."));

    private static String OutputDir
    {
        get
        {
            var dir = Path.GetFullPath("./files");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    [Fact, DisplayName("Detect_文件头识别PDF")]
    public void Detect_Pdf()
    {
        var path = Path.Combine(BinDir, "星语产品宣传.pdf");
        if (!File.Exists(path)) return;
        Assert.Equal("pdf", OfficeFactory.Detect(path));
    }

    [Fact, DisplayName("Detect_文件头识别xlsx")]
    public void Detect_Xlsx()
    {
        var path = Path.Combine(BinDir, "卫生间铝型材.xlsx");
        if (!File.Exists(path)) return;
        Assert.Equal("xlsx", OfficeFactory.Detect(path));
    }

    [Fact, DisplayName("Detect_文件头识别docx")]
    public void Detect_Docx()
    {
        var path = Path.Combine(BinDir, "A4工业计算机_v2.0.docx");
        if (!File.Exists(path)) return;
        Assert.Equal("docx", OfficeFactory.Detect(path));
    }

    [Fact, DisplayName("Detect_文件头识别pptx")]
    public void Detect_Pptx()
    {
        var path = Path.Combine(BinDir, "NET最爱Redis消息队列.pptx");
        if (!File.Exists(path)) return;
        Assert.Equal("pptx", OfficeFactory.Detect(path));
    }

    [Fact, DisplayName("Detect_RTF文本识别")]
    public void Detect_Rtf()
    {
        var text = "{\\rtf1\\ansi Hello";
        Assert.Equal("rtf", OfficeFactory.Detect(new MemoryStream(Encoding.UTF8.GetBytes(text))));
    }

    [Fact, DisplayName("Detect_Markdown文本识别")]
    public void Detect_Markdown()
    {
        var text = "# 标题\n\n这是 Markdown 文档。\n";
        Assert.Equal("md", OfficeFactory.Detect(new MemoryStream(Encoding.UTF8.GetBytes(text))));
    }

    [Fact, DisplayName("Detect_vCard文本识别")]
    public void Detect_VCard()
    {
        var text = "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:张三\r\nEND:VCARD\r\n";
        Assert.Equal("vcf", OfficeFactory.Detect(new MemoryStream(Encoding.UTF8.GetBytes(text))));
    }

    [Fact, DisplayName("Detect_iCalendar文本识别")]
    public void Detect_ICal()
    {
        var text = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR\r\n";
        Assert.Equal("ics", OfficeFactory.Detect(new MemoryStream(Encoding.UTF8.GetBytes(text))));
    }

    [Fact, DisplayName("Detect_未知二进制返回null")]
    public void Detect_UnknownBinary()
    {
        var bytes = new Byte[] { 0x01, 0x02, 0x00, 0x05, 0x07 };
        Assert.Null(OfficeFactory.Detect(new MemoryStream(bytes)));
    }

    [Fact, DisplayName("Detect_空流返回null")]
    public void Detect_EmptyStream()
    {
        Assert.Null(OfficeFactory.Detect(new MemoryStream()));
    }

    [Fact, DisplayName("CreateReaderByContent_扩展名错误按内容识别")]
    public void CreateReaderByContent_WrongExtension()
    {
        var path = Path.Combine(BinDir, "星语产品宣传.pdf");
        if (!File.Exists(path)) return;

        // 复制为错误的 .docx 扩展名
        var wrongPath = Path.Combine(OutputDir, "wrong_ext.docx");
        File.Copy(path, wrongPath, true);

        var reader = OfficeFactory.CreateReaderByContent(wrongPath);
        Assert.NotNull(reader);
        Assert.IsType<PdfReader>(reader);
        (reader as PdfReader)?.Dispose();
    }

    [Fact, DisplayName("CreateReaderByContent_正常文件按内容读取")]
    public void CreateReaderByContent_Normal()
    {
        var path = Path.Combine(BinDir, "卫生间铝型材.xlsx");
        if (!File.Exists(path)) return;

        var reader = OfficeFactory.CreateReaderByContent(path);
        Assert.NotNull(reader);
        (reader as IDisposable)?.Dispose();
    }

    [Fact, DisplayName("Detect_文件不存在抛FileNotFoundException")]
    public void Detect_FileNotFound()
    {
        Assert.Throws<FileNotFoundException>(() => OfficeFactory.Detect("不存在.xyz"));
    }
}
