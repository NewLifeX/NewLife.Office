// NewLife.Office NativeAOT 冒烟：23 场景逐项 PASS/FAIL，任一失败以非 0 退出码结束。
// 同一工程双跑对照：JIT（dotnet run -c Release）与 NativeAOT（dotnet publish -c Release -r <rid> 后运行产物）。
// 由 .github/workflows/aot.yml 自动执行（win-x64/win-x86/win-arm64/linux-x64/linux-arm64/osx-x64/osx-arm64 原生发布运行，linux-la64 经 QEMU 以 JIT 覆盖）。
using System.ComponentModel;
using System.Globalization;
using System.Text;
using NewLife;
using NewLife.Office;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Ole2;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using NewLife.Office.Xps;

var failures = new List<String>();
var workDir = Path.Combine(Path.GetTempPath(), "office-aot-smoke");
Directory.CreateDirectory(workDir);

void Check(String name, Boolean ok)
{
    Console.WriteLine($"[{(ok ? "PASS" : "FAIL")}] {name}");
    if (!ok) failures.Add(name);
}

void Run(String name, Func<Boolean> func)
{
    try
    {
        Check(name, func());
    }
    catch (Exception ex)
    {
        var lines = ex.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Console.WriteLine($"[FAIL] {name} -> {String.Join(Environment.NewLine, lines.Take(8))}");
        failures.Add(name);
    }
}

String PathOf(String name) => Path.Combine(workDir, name);

// 不变全球化模式下具名区域性不可用（GetCultureInfo 抛 CultureNotFoundException），据此判定
Boolean IsInvariantGlobalization()
{
    try { return CultureInfo.GetCultureInfo("zh-Hans").Name.Length == 0; }
    catch (CultureNotFoundException) { return true; }
}

Console.WriteLine($"IsDynamicCodeSupported={System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported}");

// 0. 编码提供程序（PdfReader/Markdown 等依赖 1252 等代码页）
Run("0.注册 CodePages 编码提供程序并取 1252", () =>
{
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    return Encoding.GetEncoding(1252).CodePage == 1252;
});

// 1. Markdown 解析 → 回写 → 再解析（文本一致）→ HTML 转换
Run("1.Markdown 解析/回写/HTML", () =>
{
    const String md = "# 标题\n\n段落 **粗体** 与 `代码`。\n\n- 项目A\n- 项目B\n\n| A | B |\n|---|---|\n| 1 | 2 |\n";
    var doc = MarkdownDocument.Parse(md);
    var doc2 = MarkdownDocument.Parse(doc.ToMarkdown());
    var html = doc.ToHtml();
    return doc.ExtractText() == doc2.ExtractText() && html.Contains("标题") && html.Contains("<table");
});

// 2. Markdown → PDF
Run("2.Markdown→PDF 生成", () =>
{
    var pdf = MarkdownDocument.Parse("# AOT\n\n正文段落。").ToPdf();
    return pdf.Length > 200 && pdf[0] == (Byte)'%';
});

// 3. xlsx 写入 → 读回（值一致）
Run("3.xlsx 写入/读回", () =>
{
    var path = PathOf("smoke.xlsx");
    using (var w = new ExcelWriter(path))
    {
        w.WriteHeader("Sheet1", ["名称", "数量"]);
        w.WriteRow("Sheet1", new Object?[] { "测试项", 42 });
    }
    using var r = new ExcelReader(path);
    var rows = r.ReadRows("Sheet1").ToList();
    return rows.Count == 2 && (String?)rows[1][0] == "测试项" && rows[1][1].ToInt() == 42;
});

// 4. ExcelHelper.SaveAs<T>/Query<T>（泛型对象映射，覆盖 DAM 标注反射路径）
Run("4.ExcelHelper 对象映射 SaveAs/Query", () =>
{
    var path = PathOf("smoke_objects.xlsx");
    var data = new List<SmokeOrder>
    {
        new() { Id = 1, Customer = "张三", Amount = 99.5 },
        new() { Id = 2, Customer = "李四", Amount = 199 },
    };
    ExcelHelper.SaveAs(path, data);
    var back = ExcelHelper.Query<SmokeOrder>(path).ToList();
    return back.Count == 2 && back[1].Id == 2 && back[1].Customer == "李四" && Math.Abs(back[0].Amount - 99.5) < 1e-6;
});

// 5. xls（BIFF8）写入 → 读回 → ReadObjects<T>
Run("5.xls BIFF 写入/读回/对象映射", () =>
{
    using var w = new BiffWriter();
    w.WriteHeader(["Name", "Age"]);
    w.WriteRow(new Object?[] { "Alice", 30 });
    var bytes = w.ToBytes();
    using var r = new BiffReader(new MemoryStream(bytes));
    var rows = r.ReadSheet().ToList();
    using var r2 = new BiffReader(new MemoryStream(bytes));
    var objs = r2.ReadObjects<SmokePerson>().ToList();
    return rows.Count == 2 && (String?)rows[1][0] == "Alice" && objs.Count == 1 && objs[0].Name == "Alice" && objs[0].Age == 30;
});

// 6. xlsx 模板填充（占位符替换）
Run("6.xlsx 模板填充占位符", () =>
{
    var tpl = PathOf("tpl.xlsx");
    using (var w = new ExcelWriter(tpl))
    {
        w.WriteRow("Sheet1", new Object?[] { "{{Name}}", "{{Count}}" });
    }
    var outPath = PathOf("tpl_out.xlsx");
    new ExcelTemplate(tpl).Fill(outPath, new Dictionary<String, Object> { ["Name"] = "填充值", ["Count"] = 7 });
    using var r = new ExcelReader(outPath);
    var rows = r.ReadRows("Sheet1").ToList();
    return rows.Count >= 1 && (String?)rows[0][0] == "填充值" && rows[0][1].ToInt() == 7;
});

// 7. docx 写入 → 读回（段落/表格）→ ReadObjects<T>
Run("7.docx 写入/读回/对象映射", () =>
{
    var path = PathOf("smoke.docx");
    using (var w = new WordWriter())
    {
        w.AppendHeading("AOT 冒烟", 1);
        w.AppendParagraph("正文段落内容。");
        w.AppendTable([["Name", "Age"], ["Alice", "30"]], firstRowHeader: true);
        w.Save(path);
    }
    using var r = new WordReader(path);
    var paras = r.ReadParagraphs().ToList();
    var tables = r.ReadTables().ToList();
    using var r2 = new WordReader(path);
    var objs = r2.ReadObjects<SmokePerson>().ToList();
    return paras.Any(p => p.Contains("正文段落内容")) && tables.Count == 1 && tables[0].Length == 2
        && objs.Count == 1 && objs[0].Name == "Alice" && objs[0].Age == 30;
});

// 8. docx 模板填充（字典 + 对象模型两条路径）
Run("8.docx 模板填充字典/对象", () =>
{
    var tpl = PathOf("tpl.docx");
    using (var w = new WordWriter())
    {
        w.AppendParagraph("姓名：{{Name}}，年龄：{{Age}}");
        w.Save(tpl);
    }
    var out1 = PathOf("tpl_out1.docx");
    new WordTemplate(tpl).Fill(out1, new Dictionary<String, Object?> { ["Name"] = "张三", ["Age"] = 30 });
    var out2 = PathOf("tpl_out2.docx");
    new WordTemplate(tpl).Fill(out2, new SmokePerson { Name = "李四", Age = 25 });
    using var r1 = new WordReader(out1);
    using var r2 = new WordReader(out2);
    return r1.ReadFullText().Contains("姓名：张三") && r2.ReadFullText().Contains("姓名：李四");
});

// 9. WordWriter.WriteObjects<T> 嵌套对象展开（PropertyInfo 链路径）
Run("9.docx 对象表格（含嵌套展开）", () =>
{
    // NativeAOT 下嵌套展开依赖运行时发现的属性类型；消费方以 DynamicDependency 显式保留嵌套模型成员（库内 CollectProperties 的降级约定配套）
    AotAnchors.PreserveNestedModel();
    var path = PathOf("objects.docx");
    using (var w = new WordWriter())
    {
        w.WriteObjects(new List<SmokeNestedOrder>
        {
            new() { Code = "A1", Customer = new SmokeCustomer { Name = "王五", City = "北京" } },
        }, firstRowHeader: true, style: null, maxDepth: 1);
        w.Save(path);
    }
    using var r = new WordReader(path);
    var tables = r.ReadTables().ToList();
    var header = tables.Count > 0 ? String.Join("|", tables[0][0]) : "";
    var body = tables.Count > 0 && tables[0].Length > 1 ? String.Join("|", tables[0][1]) : "";
    return header.Contains("Customer.Name") && body.Contains("王五") && body.Contains("北京");
});

// 10. pptx 创建 → 读回 → 对象写入
Run("10.pptx 创建/读回/对象写入", () =>
{
    var path = PathOf("smoke.pptx");
    using (var w = new PptxWriter())
    {
        w.AddSlide();
        w.AddTextBox(0, "AOT 冒烟标题", 2, 2, 20, 3, 28, true);
        w.WriteObjects(0, new List<SmokePerson> { new() { Name = "赵六", Age = 40 } });
        w.Save(path);
    }
    using var r = new PptxReader(path);
    var text = r.GetSlideText(0) ?? "";
    return r.GetSlideCount() == 1 && text.Contains("AOT 冒烟标题") && text.Contains("赵六");
});

// 11. PDF 写入 → 读回文本
Run("11.PDF 写入/读回文本", () =>
{
    using var ms = new MemoryStream();
    using (var w = new PdfWriter())
    {
        w.BeginPage();
        w.AppendLine("Hello AOT PDF");
        w.EndPage();
        w.Save(ms);
    }
    ms.Position = 0;
    using var r = new PdfReader(ms);
    var text = r.ExtractText();
    return r.GetPageCount() == 1 && text.Contains("Hello AOT PDF");
});

// 12. PDF Fluent + AddTable<T>（泛型反射路径）
Run("12.PDF Fluent 表格（对象映射）", () =>
{
    using var ms = new MemoryStream();
    using (var doc = new PdfDocumentBuilder())
    {
        doc.AddText("Fluent 冒烟", 16);
        doc.AddTable(new List<SmokePerson> { new() { Name = "孙七", Age = 50 } });
        doc.Save(ms);
    }
    ms.Position = 0;
    using var r = new PdfReader(ms);
    return r.ExtractText().Contains("孙七");
});

// 13. RTF 写入 → 解析
Run("13.RTF 写入/解析", () =>
{
    var writer = new RtfWriter();
    writer.AddParagraph("Hello RTF");
    var doc = RtfDocument.Parse(writer.ToString());
    return doc.GetPlainText().Contains("Hello RTF");
});

// 14. ODS 写入 → 读回 + 对象映射
Run("14.ODS 写入/读回/对象映射", () =>
{
    var ms = new MemoryStream();
    var w = new OdsWriter();
    w.AddSheet("Sheet1", new[] { new[] { "Name", "Age" }, new[] { "周八", "60" } });
    w.AddSheet("Sheet2", new List<SmokePerson> { new() { Name = "吴九", Age = 70 } });
    w.Save(ms);
    ms.Position = 0;
    var sheets = OdsReader.Read(ms);
    ms.Position = 0;
    var objs = OdsReader.ReadObjects<SmokePerson>(ms, "Sheet2").ToList();
    return sheets.Count == 2 && sheets[0].Rows[1][0] == "周八" && objs.Count == 1 && objs[0].Age == 70;
});

// 15. EML 邮件写入 → 解析
Run("15.EML 写入/解析", () =>
{
    var path = PathOf("smoke.eml");
    var msg = new Message { From = "a@b.com", Subject = "AOT 邮件", TextBody = "Hello" };
    msg.To.Add("c@d.com");
    new EmlWriter().Write(msg, path);
    var back = new EmlReader().Read(path);
    return back.Subject == "AOT 邮件" && (back.TextBody ?? "").Contains("Hello");
});

// 16. iCalendar 写入 → 解析
Run("16.ICS 写入/解析", () =>
{
    var path = PathOf("smoke.ics");
    var doc = new ICalDocument();
    doc.Events.Add(new ICalEvent
    {
        Uid = "aot-001",
        Summary = "冒烟日程",
        Start = new DateTimeOffset(2026, 10, 16, 9, 0, 0, TimeSpan.Zero),
        End = new DateTimeOffset(2026, 10, 16, 10, 0, 0, TimeSpan.Zero),
    });
    new ICalWriter().Write(doc, path);
    var back = new ICalReader().Read(path);
    return back.Events.Count == 1 && back.Events[0].Summary == "冒烟日程";
});

// 17. vCard 写入 → 解析
Run("17.vCard 写入/解析", () =>
{
    var path = PathOf("smoke.vcf");
    var contact = new Contact { FullName = "郑十" };
    contact.Emails.Add(new Email { Address = "z@newlife.org", Type = "WORK" });
    new VCardWriter().Write(contact, path);
    var contacts = new VCardReader().ReadAll(path);
    return contacts.Count == 1 && contacts[0].FullName == "郑十";
});

// 18. EPUB 写入 → 解析
Run("18.EPUB 写入/解析", () =>
{
    var path = PathOf("smoke.epub");
    var doc = new EpubDocument { Title = "AOT 电子书", Author = "测试" };
    new EpubWriter().Write(doc, path);
    var back = new EpubReader().Read(path);
    return back.Title == "AOT 电子书";
});

// 19. XPS 写入 → 解析
Run("19.XPS 写入/解析", () =>
{
    var ms = new MemoryStream();
    var w = new XpsWriter();
    w.AddPage(816, 1056, new[] { ("Hello XPS", 96.0, 96.0, 16.0) });
    w.Save(ms);
    ms.Position = 0;
    var pages = new XpsReader().Read(ms);
    return pages.Count == 1 && pages[0].Text.Contains("Hello XPS");
});

// 20. OLE2/CFB 容器写入 → 读取
Run("20.OLE2 容器写入/读取", () =>
{
    using var doc = new CfbDocument();
    doc.PutStream("AotSmoke", Encoding.UTF8.GetBytes("cfb-payload"));
    using var ms = new MemoryStream();
    doc.Save(ms);
    ms.Position = 0;
    using var doc2 = CfbDocument.Open(ms);
    var data = doc2.GetStreamData("AotSmoke");
    return data != null && Encoding.UTF8.GetString(data) == "cfb-payload";
});

// 21. OfficeFactory 内容识别与路由
Run("21.OfficeFactory 内容识别", () =>
{
    var xlsx = PathOf("smoke.xlsx");
    var pdf = new MemoryStream();
    using (var w = new PdfWriter())
    {
        w.BeginPage();
        w.AppendLine("detect");
        w.EndPage();
        w.Save(pdf);
    }
    pdf.Position = 0;
    var rtf = new MemoryStream(Encoding.UTF8.GetBytes(@"{\rtf1 detect}"));
    var format = OfficeFactory.Detect(xlsx);
    var format2 = OfficeFactory.Detect(pdf);
    var format3 = OfficeFactory.Detect(rtf);
    var reader = OfficeFactory.CreateReader(xlsx);
    (reader as IDisposable)?.Dispose();
    return format == "xlsx" && format2 == "pdf" && format3 == "rtf" && reader != null;
});

// 22. 多语言：31 种内置语言表在 AOT 下可用（不变全球化模式下跳过，仍验证英文回退）
Run("22.多语言 全部内置语言表可用", () =>
{
    const String key = "No output location was specified.";
    var saved = Localization.Culture;
    try
    {
        // 不变全球化模式下具名区域性不可用，跳过语种遍历，仅验证英文回退
        if (IsInvariantGlobalization()) return Localization.Translate(key) == key;

        Localization.Culture = CultureInfo.GetCultureInfo("en-US");
        if (Localization.Translate(key) != key) return false;

        var cultures = new[]
        {
            "zh-Hans", "zh-Hant", "ja", "ko", "vi", "th", "id", "ms", "hi", "bn",
            "ar", "he", "fa", "de", "fr", "es", "pt", "it", "nl", "pl",
            "cs", "sv", "da", "nb", "fi", "el", "ro", "hu", "ru", "uk", "tr",
        };
        var missing = new List<String>();
        foreach (var name in cultures)
        {
            Localization.Culture = CultureInfo.GetCultureInfo(name);
            var text = Localization.Translate(key);
            if (text == key || text.Length == 0) missing.Add(name);
        }
        if (missing.Count > 0) throw new Exception("以下语言缺少译文: " + String.Join(", ", missing));
        return true;
    }
    finally
    {
        Localization.Culture = saved;
    }
});

Console.WriteLine();
if (failures.Count > 0)
{
    Console.WriteLine($"❌ AOT 冒烟失败 {failures.Count} 项：{String.Join(", ", failures)}");
    return 1;
}

Console.WriteLine("✅ AOT 冒烟全部通过");
return 0;

/// <summary>冒烟用对象模型（公开属性，覆盖 DisplayName 映射与泛型反射路径）</summary>
public class SmokeOrder
{
    [DisplayName("订单号")]
    public Int32 Id { get; set; }

    [DisplayName("客户")]
    public String? Customer { get; set; }

    [DisplayName("金额")]
    public Double Amount { get; set; }
}

/// <summary>冒烟用对象模型（读写对象映射）</summary>
public class SmokePerson
{
    public String? Name { get; set; }

    public Int32 Age { get; set; }
}

/// <summary>嵌套对象模型（WordWriter 嵌套展开）</summary>
public class SmokeNestedOrder
{
    public String? Code { get; set; }

    public SmokeCustomer? Customer { get; set; }
}

/// <summary>嵌套对象模型（子对象）</summary>
public class SmokeCustomer
{
    public String? Name { get; set; }

    public String? City { get; set; }
}

/// <summary>AOT 保留锚点：演示消费方如何为嵌套展开模型保留成员</summary>
public static class AotAnchors
{
    /// <summary>保留嵌套模型 SmokeCustomer 的公开属性（供嵌套展开动态读取）</summary>
    [System.Diagnostics.CodeAnalysis.DynamicDependency(
        System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.PublicProperties, typeof(SmokeCustomer))]
    public static void PreserveNestedModel() { }
}
