using System.ComponentModel;
using System.Text;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Ods;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Interop;

/// <summary>跨库互操作夹具：本库产出供第三方读取，第三方产出供本库读取</summary>
/// <remarks>
/// 由 CI 分三步驱动：
/// 1. 运行 <see cref="GenerateOurDocuments"/>，把本库写出的文档放进 {OFFICE_INTEROP_DIR}/ours；
/// 2. Python 侧用 openpyxl/python-docx/python-pptx/pypdf/odfpy/icalendar/vobject 等读取并写出 {dir}/theirs；
/// 3. 运行 <see cref="ReadThirdPartyDocuments"/> 让本库读取第三方文件。
/// 目录由环境变量 OFFICE_INTEROP_DIR 指定，本地默认落到临时目录。
/// </remarks>
public class InteropFixtureTests
{
    /// <summary>互操作根目录</summary>
    public static String RootDir
    {
        get
        {
            var dir = Environment.GetEnvironmentVariable("OFFICE_INTEROP_DIR");
            if (dir == null || dir.Length == 0) dir = Path.Combine(Path.GetTempPath(), "nloffice-interop");

            return dir.GetFullPath();
        }
    }

    private static String EnsureDir(String sub)
    {
        var dir = Path.Combine(RootDir, sub);
        Directory.CreateDirectory(dir);
        return dir;
    }

    [Fact, DisplayName("互操作—生成本库文档供第三方读取")]
    public void GenerateOurDocuments()
    {
        var dir = EnsureDir("ours");

        // xlsx
        using (var w = new ExcelWriter(Path.Combine(dir, "ours.xlsx")))
        {
            w.WriteHeader("S1", ["Id", "Name", "Amount", "Flag", "Date"]);
            w.WriteRows("S1",
            [
                new Object?[] { 1, "甲", 1234.5, true, new DateTime(2024, 1, 2) },
                new Object?[] { 2, "乙", -0.25, false, new DateTime(2024, 3, 4) },
            ]);
            w.Save();
        }

        // xls
        using (var w = new BiffWriter())
        {
            w.SheetName = "S1";
            w.WriteHeader(["Id", "Name"]);
            w.WriteRow([1, "甲"]);
            w.Save(Path.Combine(dir, "ours.xls"));
        }

        // docx
        using (var ms = new MemoryStream())
        {
            using (var w = new WordWriter())
            {
                w.AppendParagraph("第一段中文内容 paragraph one.");
                w.Save(ms);
            }
            File.WriteAllBytes(Path.Combine(dir, "ours.docx"), ms.ToArray());
        }

        // pptx
        using (var ms = new MemoryStream())
        {
            using (var w = new PptxWriter())
            {
                w.AddSlide();
                w.AddTextBox(0, "幻灯片标题", 2, 2, 20, 3);
                w.Save(ms);
            }
            File.WriteAllBytes(Path.Combine(dir, "ours.pptx"), ms.ToArray());
        }

        // pdf
        using (var ms = new MemoryStream())
        {
            using (var w = new PdfWriter())
            {
                w.DrawText("Hello PDF 第一页", 50, 700, 14);
                w.Save(ms);
            }
            File.WriteAllBytes(Path.Combine(dir, "ours.pdf"), ms.ToArray());
        }

        // ods
        var ods = new OdsWriter();
        ods.AddSheet("S1", [["Id", "Name"], ["1", "甲"]]);
        using (var ms = new MemoryStream())
        {
            ods.Save(ms);
            File.WriteAllBytes(Path.Combine(dir, "ours.ods"), ms.ToArray());
        }

        // epub
        var epub = new EpubDocument { Title = "互操作", Author = "NewLife" };
        epub.Chapters.Add(new EpubChapter { Title = "第一章", Content = "<p>第一章内容</p>", FileName = "chapter01.xhtml" });
        using (var ms = new MemoryStream())
        {
            new EpubWriter().Write(epub, ms);
            File.WriteAllBytes(Path.Combine(dir, "ours.epub"), ms.ToArray());
        }

        // ics
        var cal = new ICalDocument { CalendarName = "互操作日历" };
        cal.Events.Add(new ICalEvent
        {
            Uid = "evt-1@newlife",
            Summary = "会议;含逗号,和分号",
            Start = new DateTimeOffset(2024, 1, 2, 12, 0, 0, TimeSpan.Zero),
            End = new DateTimeOffset(2024, 1, 2, 13, 0, 0, TimeSpan.Zero),
        });
        File.WriteAllText(Path.Combine(dir, "ours.ics"), new ICalWriter().Build(cal), new UTF8Encoding(false));

        // vcf
        var contact = new Contact
        {
            FullName = "张三",
            Uid = "contact-1@newlife",
            Organization = "NewLife",
        };
        contact.Emails.Add(new NewLife.Office.VCard.Email { Address = "zhangsan@example.com", Type = "WORK" });
        File.WriteAllText(Path.Combine(dir, "ours.vcf"), new VCardWriter().Build(contact), new UTF8Encoding(false));

        // eml
        var msg = new Message
        {
            From = "sender@example.com",
            Subject = "互操作测试",
            TextBody = "第一行\r\n第二行",
        };
        msg.To.Add("receiver@example.com");
        using (var ms = new MemoryStream())
        {
            new EmlWriter().Write(msg, ms);
            File.WriteAllBytes(Path.Combine(dir, "ours.eml"), ms.ToArray());
        }

        Assert.True(File.Exists(Path.Combine(dir, "ours.xlsx")));
    }

    [Fact, DisplayName("互操作—读取第三方生成的文档")]
    public void ReadThirdPartyDocuments()
    {
        var dir = Path.Combine(RootDir, "theirs");
        if (!Directory.Exists(dir))
        {
            Console.WriteLine($"跳过：未找到第三方文档目录 {dir}");
            return;
        }

        var files = Directory.GetFiles(dir);
        if (files.Length == 0)
        {
            Console.WriteLine($"跳过：第三方文档目录为空 {dir}");
            return;
        }

        var failures = new List<String>();

        foreach (var file in files)
        {
            try
            {
                switch (Path.GetExtension(file).ToLowerInvariant())
                {
                    case ".xlsx":
                        using (var reader = new ExcelReader(file)) Assert.NotEmpty(reader.ReadRows().ToList());
                        break;
                    case ".xls":
                        using (var reader = new BiffReader(file)) Assert.NotEmpty(reader.ReadSheet().ToList());
                        break;
                    case ".docx":
                        using (var reader = new WordReader(file)) Assert.NotEmpty(reader.ReadParagraphs().ToList());
                        break;
                    case ".pptx":
                        using (var reader = new PptxReader(file)) Assert.NotEmpty(reader.ReadAllText());
                        break;
                    case ".pdf":
                        using (var reader = new PdfReader(file)) Assert.True(reader.GetPageCount() > 0);
                        break;
                    case ".ods":
                        Assert.NotEmpty(OdsReader.ReadFile(file));
                        break;
                    case ".epub":
                        Assert.NotNull(new EpubReader().Read(file));
                        break;
                    case ".ics":
                        Assert.NotEmpty(new ICalReader().Read(file).Events);
                        break;
                    case ".vcf":
                        Assert.NotEmpty(new VCardReader().ReadAll(file));
                        break;
                    case ".eml":
                        Assert.NotNull(new EmlReader().Read(file));
                        break;
                    default:
                        continue;
                }
            }
            catch (Exception ex)
            {
                failures.Add($"{Path.GetFileName(file)}: {ex.GetType().Name} {ex.Message}");
            }
        }

        Assert.True(failures.Count == 0, "第三方文档读取失败：" + String.Join("; ", failures));
    }
}
