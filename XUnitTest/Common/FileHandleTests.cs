using System.ComponentModel;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.Word;
using NewLife.Office.Xps;
using Xunit;
using WDocument = NewLife.Office.Word.Document;

namespace XUnitTest.Common;

/// <summary>文件句柄释放验证：写入/读取结束后不得继续占用文件，且可安全覆盖或删除</summary>
/// <remarks>
/// Windows 下文件被进程独占时 <c>File.Delete</c> 或再次覆盖写会抛 IOException。
/// 用户"生成完文件却删不掉/覆盖不了"是典型痛点，根因多为 writer/reader 内部 Stream 未释放。
/// 本类对主要格式逐一验证：写完可再覆盖写、读完可删除。
/// </remarks>
public class FileHandleTests
{
    private static String NewDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "fh_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    /// <summary>写入一次、再同名覆盖写一次，可选读一次，最后删除（每步都必须成功）</summary>
    /// <param name="path">目标文件路径</param>
    /// <param name="write">写入动作</param>
    /// <param name="read">读取动作（可为 null）</param>
    private static void VerifyReusable(String path, Action write, Action<String>? read)
    {
        write();
        Assert.True(File.Exists(path), "首次写入未产生文件：" + path);

        // 覆盖写：若上一次的句柄未释放，这里会抛 IOException
        write();
        Assert.True(File.Exists(path));

        read?.Invoke(path);

        // 读取结束后必须可以删除
        File.Delete(path);
        Assert.False(File.Exists(path));
    }

    [Fact, DisplayName("句柄-DOCX 写完可覆盖/删除，读完可删除")]
    public void Docx_ReleasesHandle()
    {
        var dir = NewDir();
        try
        {
            var path = Path.Combine(dir, "a.docx");
            VerifyReusable(path, () =>
            {
                var doc = new WDocument();
                doc.AddParagraph("内容");
                new WordWriter().Save(path, doc);
            }, p =>
            {
                using var reader = new WordReader(p);
                Assert.NotEmpty(reader.ReadDocument().Elements);
            });
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("句柄-XLSX 写完可覆盖/删除，读完可删除")]
    public void Xlsx_ReleasesHandle()
    {
        var dir = NewDir();
        try
        {
            var path = Path.Combine(dir, "a.xlsx");
            VerifyReusable(path, () =>
            {
                using var w = new ExcelWriter(path);
                w.WriteHeader("S1", ["A", "B"]);
                w.WriteRows("S1", [new Object?[] { 1, 2 }]);
                w.Save();
            }, p =>
            {
                using var r = new ExcelReader(p);
                Assert.NotEmpty(r.Sheets!);
            });
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("句柄-PPTX 写完可覆盖/删除，读完可删除")]
    public void Pptx_ReleasesHandle()
    {
        var dir = NewDir();
        try
        {
            var path = Path.Combine(dir, "a.pptx");
            VerifyReusable(path, () =>
            {
                using var w = new PptxWriter();
                w.AddSlide(0);
                w.AddTextBox(0, "标题", 2, 2, 10, 2, 20, false);
                w.Save(path);
            }, p =>
            {
                using var r = new PptxReader(p);
                Assert.True(r.GetSlideCount() >= 1);
            });
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("句柄-PDF 写完可覆盖/删除，读完可删除")]
    public void Pdf_ReleasesHandle()
    {
        var dir = NewDir();
        try
        {
            var path = Path.Combine(dir, "a.pdf");
            VerifyReusable(path, () =>
            {
                using var w = new PdfWriter();
                w.BeginPage();
                w.DrawText("PDF content", 50, 700);
                w.EndPage();
                w.Save(path);
            }, p =>
            {
                using var r = new PdfReader(p);
                Assert.False(String.IsNullOrEmpty(r.ExtractText()));
            });
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("句柄-RTF 写完可覆盖/删除")]
    public void Rtf_ReleasesHandle()
    {
        var dir = NewDir();
        try
        {
            var path = Path.Combine(dir, "a.rtf");
            VerifyReusable(path, () =>
            {
                var w = new RtfWriter();
                w.AddParagraph("RTF 内容");
                w.Save(path);
            }, p => Assert.True(new FileInfo(p).Length > 0));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("句柄-EPUB 写完可覆盖/删除")]
    public void Epub_ReleasesHandle()
    {
        var dir = NewDir();
        try
        {
            var path = Path.Combine(dir, "a.epub");
            VerifyReusable(path, () =>
            {
                var doc = new EpubDocument { Title = "书", Author = "作者" };
                doc.Chapters.Add(new EpubChapter { Title = "第一章", Content = "<p>正文</p>" });
                new EpubWriter().Write(doc, path);
            }, p =>
            {
                var back = new EpubReader().Read(p);
                Assert.Equal("书", back.Title);
            });
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("句柄-XPS 写完可覆盖/删除")]
    public void Xps_ReleasesHandle()
    {
        var dir = NewDir();
        try
        {
            var path = Path.Combine(dir, "a.xps");
            VerifyReusable(path, () =>
            {
                var w = new XpsWriter();
                w.AddPage(816, 1056, [("XPS content", 50, 100, 12)]);
                w.Save(path);
            }, p => Assert.True(new FileInfo(p).Length > 0));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}
