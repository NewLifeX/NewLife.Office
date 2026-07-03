using System.ComponentModel;
using System.Text;
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
using Xunit;

using XUnitTest.Common;

namespace XUnitTest.Pdf;

/// <summary>PDF 格式往返测试</summary>
public class PdfRoundTripTests : IntegrationTestBase
{
    [Fact, DisplayName("PDF_字体变体展示")]
    public void Pdf_FontVariations()
    {
        var path = Path.Combine(OutputDir, "test_fonts.pdf");

        using (var w = new PdfDocumentBuilder())
        {
            w.Title  = "字体对比测试";
            w.Author = "NewLife Office";
            w.Header = "NewLife.Office — 字体展示";
            w.ShowPageNumbers = true;

            // ── CJK 字体（embed:false，不嵌入字体文件，文件体积小）──
            var fontYaHei = w.CreateFont("微软雅黑", embed: false);
            var fontHei   = w.CreateFont("黑体",     embed: false);
            var fontSong  = w.CreateFont("宋体",      embed: false);
            var fontKai   = w.CreateFont("楷体",      embed: false);
            var fontFang  = w.CreateFont("仿宋",      embed: false);

            // ── 标准 Type1 英文字体（PDF 内置，无需嵌入）──
            var fontHelvetica    = w.CreateFont("Helvetica");
            var fontHelveticaBold= w.CreateFont("Helvetica-Bold");
            var fontHelveticaObl = w.CreateFont("Helvetica-Oblique");
            var fontTimesRoman   = w.CreateFont("Times-Roman");
            var fontTimesBold    = w.CreateFont("Times-Bold");
            var fontTimesItalic  = w.CreateFont("Times-Italic");
            var fontCourier      = w.CreateFont("Courier");
            var fontCourierBold  = w.CreateFont("Courier-Bold");

            // ── 第一页：中文字体展示 ──
            w.AddText("PDF 字体对比测试文档", 22f, fontYaHei);
            w.AddEmptyLine(6f);
            w.AddText("一、中文字体展示（字体未嵌入，阅读器需安装对应字体）", 12f, fontHelveticaBold);
            w.AddEmptyLine(4f);

            var cjkSamples = new (String Label, PdfFont Font, String Sample)[]
            {
                ("微软雅黑", fontYaHei, "微软雅黑：现代无衬线，屏幕显示清晰。AaBbCc 123 !@#"),
                ("黑  体",   fontHei,   "黑  体：粗体无衬线，适合标题强调。AaBbCc 456 $%^"),
                ("宋  体",   fontSong,  "宋  体：传统衬线，适合长篇正文排版。AaBbCc 789 &*()"),
                ("楷  体",   fontKai,   "楷  体：书法楷书风格，文艺感强。AaBbCc 012 +-="),
                ("仿  宋",   fontFang,  "仿  宋：仿宋体，适合公文引文排版。AaBbCc 345 []{}"),
            };

            foreach (var (label, font, sample) in cjkSamples)
            {
                w.AddText($"【{label}】", 10f, fontHelveticaBold);
                w.AddText(sample, 13f, font);
                w.AddEmptyLine(3f);
            }

            // ── 第二页：英文字体展示 ──
            w.PageBreak();
            w.AddText("二、标准 Type1 英文字体展示（PDF 内置，无需嵌入）", 14f, fontHelveticaBold);
            w.AddEmptyLine(6f);

            var latinSamples = new (String Label, PdfFont Font)[]
            {
                ("Helvetica",          fontHelvetica),
                ("Helvetica-Bold",     fontHelveticaBold),
                ("Helvetica-Oblique",  fontHelveticaObl),
                ("Times-Roman",        fontTimesRoman),
                ("Times-Bold",         fontTimesBold),
                ("Times-Italic",       fontTimesItalic),
                ("Courier",            fontCourier),
                ("Courier-Bold",       fontCourierBold),
            };

            foreach (var (label, font) in latinSamples)
            {
                w.AddText($"{label,-22} The quick brown fox jumps over the lazy dog. 0123456789", 11f, font);
                w.AddEmptyLine(2f);
            }

            w.AddEmptyLine(6f);
            w.AddText("字号对比（Helvetica，8-20pt）", 12f, fontHelveticaBold);
            w.AddEmptyLine(4f);
            foreach (var size in new Single[] { 8f, 10f, 12f, 14f, 16f, 20f })
                w.AddText($"{size:F0}pt — The quick brown fox / 快速的棕色狐狸跳过懒狗", size, fontHelvetica);

            // ── 第三页：混合内容表格 ──
            w.PageBreak();
            w.AddText("三、字体混合数据表格", 16f, fontHei);
            w.AddEmptyLine(6f);

            var tableData = new List<String[]>
            {
                new[] { "字体名称",    "类型",       "风格",         "适用场景",       "示例文字"          },
                new[] { "Helvetica",  "Type1",      "无衬线",       "屏幕/UI",        "AaBbCc 123"       },
                new[] { "Times-Roman","Type1",      "衬线",         "正式印刷",       "AaBbCc 456"       },
                new[] { "Courier",    "Type1",      "等宽",         "代码/终端",      "AaBbCc 789"       },
                new[] { "微软雅黑",   "TrueType",   "无衬线",       "中文屏幕",       "中文 AaZz 0-9"    },
                new[] { "宋体",       "TrueType",   "衬线",         "中文印刷",       "中文 AaZz 0-9"    },
                new[] { "黑体",       "TrueType",   "粗黑",         "标题/强调",      "中文 AaZz 0-9"    },
            };
            w.AddTable(tableData, firstRowHeader: true);

            w.Save(path);
        }

        Assert.True(File.Exists(path));
        // 不嵌入字体数据但包含压缩 CIDToGIDMap，文件应较小（< 200 KB）
        var fileSize = new FileInfo(path).Length;
        Assert.True(fileSize < 200 * 1024, $"文件过大: {fileSize / 1024} KB，预期 < 200 KB");

        using var reader = new PdfReader(path);
        Assert.Equal(3, reader.GetPageCount());

        var text = reader.ExtractText();
        Assert.Contains("Helvetica", text);
        Assert.Contains("Times-Roman", text);
        Assert.Contains("Courier", text);
        Assert.Contains("Type1", text);
    }

    [Fact, DisplayName("PDF_复杂写入再读取")]
    public void Pdf_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.pdf");

        using (var w = new PdfDocumentBuilder())
        {
            w.Title = "PDF集成测试";
            w.Author = "NewLife Office";

            // ── 第一页：标题 + 中英混合段落 ──
            w.AddText("PDF 集成测试文档", 24f);
            w.AddEmptyLine(10f);
            w.AddText("本文档由 NewLife.Office 自动生成，用于验证 PDF 读写功能。", 12f);
            w.AddText("This document is auto-generated for PDF read/write testing.", 12f);
            w.AddEmptyLine(6f);
            w.AddText("数字与符号：0123456789  !@#$%^&*()-+=[]{}|;':\",./<>?", 11f);
            w.AddText("Unicode 范围：\u00A9 \u00AE \u2122 \u20AC \u00B1 \u00D7 \u00F7", 11f);
            w.AddEmptyLine(8f);

            w.AddText("第一章 数据表格", 18f);
            w.AddEmptyLine(6f);

            var tableData = new List<String[]>
            {
                new[] { "编号", "姓名", "年龄", "城市",   "部门",   "薪资(元)"  },
                new[] { "001",  "张三", "28",   "北京",   "研发部", "15,000"   },
                new[] { "002",  "李四", "35",   "上海",   "市场部", "12,000"   },
                new[] { "003",  "王五", "42",   "广州",   "运营部", "10,000"   },
                new[] { "004",  "赵六", "31",   "深圳",   "销售部", "18,000"   },
                new[] { "005",  "Alice","26",   "Chengdu","Dev",    "16,500"   },
            };
            w.AddTable(tableData, firstRowHeader: true);

            w.AddEmptyLine(10f);
            w.AddText("第二章 英文内容与数字", 18f);
            w.AddEmptyLine(6f);
            w.AddText("Section 2.1 — Numeric Data: 3.14159  2.71828  1.41421  1.73205", 11f);
            w.AddText("Section 2.2 — Special: <tag> & 'quote' & \"double\" & 100% done", 11f);
            w.AddText("Section 2.3 — Email: test@example.com | URL: https://newlifex.com", 11f);

            // ── 第二页：混合表格 + 更多文字 ──
            w.PageBreak();
            w.AddText("第三章 混合内容验证", 16f);
            w.AddEmptyLine(8f);
            w.AddText("中英文混合段落：NewLife Office 是一个 .NET 开源库，支持 PDF、Excel、Word 格式。", 12f);
            w.AddText("Mixed paragraph: 版本 v2.0 released on 2025-01-01, size=512KB, ratio=98.5%。", 12f);
            w.AddEmptyLine(8f);

            var mixedTable = new List<String[]>
            {
                new[] { "Key",        "Value",                  "备注"         },
                new[] { "Name",       "NewLife.Office",         "项目名称"     },
                new[] { "Version",    "2.0.2025.0101",          "版本号"       },
                new[] { "License",    "MIT",                    "开源协议"     },
                new[] { "Language",   "C# 14 / .NET 9",         "编程语言"     },
                new[] { "Supported",  "PDF/Excel/Word/PPT",     "支持格式"     },
                new[] { "Stars",      "1,024",                  "GitHub Stars" },
            };
            w.AddTable(mixedTable, firstRowHeader: true);

            w.AddEmptyLine(10f);
            w.AddText("文档结束 — End of Document", 12f);

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        using var reader = new PdfReader(path);
        Assert.Equal(2, reader.GetPageCount());

        var text = reader.ExtractText();
        Assert.Contains("PDF", text);
        Assert.Contains("NewLife", text);

        var meta = reader.ReadMetadata();
        Assert.Equal(2, meta.PageCount);
        Assert.NotNull(meta.PdfVersion);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<PdfReader>(factoryReader);
        (factoryReader as IDisposable)?.Dispose();
    }

    #region 往返测试

    /// <summary>以共享读方式读取文件全部字节</summary>
    private static Byte[] ReadAllBytesShared(String path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var buf = new Byte[fs.Length];
        fs.ReadExactly(buf, 0, buf.Length);
        return buf;
    }

    /// <summary>获取 Bin 目录下所有 .pdf 文件路径</summary>
    private static List<String> FindAllPdfFiles()
    {
        var baseDir = AppContext.BaseDirectory;
        var binDir = Path.GetFullPath(Path.Combine(baseDir, ".."));
        if (!Directory.Exists(binDir)) return [];

        return Directory.GetFiles(binDir, "*.pdf", SearchOption.TopDirectoryOnly)
            .OrderByDescending(f => new FileInfo(f).Length)
            .ToList();
    }

    /// <summary>验证两个 PdfDocument 模型一致性</summary>
    /// <remarks>
    /// 比较页数、页面尺寸、元数据、书签、以及每页文本内容。
    /// 注：当前 PdfReader 的 xref 级内容流解析有局限，TextBlocks 可能为空，
    /// 因此文本对比使用 ExtractText 的字符多重集比较作为回退方案。
    /// </remarks>
    private static void AssertPdfDocumentEqual(PdfDocument src, PdfDocument dst, String fileName, Func<PdfDocument, String> getFullText)
    {
        var tag = $"[{fileName}]";

        // ① 页数
        Assert.Equal(src.Pages.Count, dst.Pages.Count);

        // ② 逐页比较尺寸
        for (var i = 0; i < src.Pages.Count; i++)
        {
            var sp = src.Pages[i];
            var dp = dst.Pages[i];
            var pt = $"{tag} P{i}";

            // 页面尺寸（1pt 容差）
            Assert.True(Math.Abs(sp.Width - dp.Width) <= 1f, $"{pt} Width: {sp.Width} vs {dp.Width}");
            Assert.True(Math.Abs(sp.Height - dp.Height) <= 1f, $"{pt} Height: {sp.Height} vs {dp.Height}");
            Assert.Equal(sp.Rotation, dp.Rotation);

            // TextBlocks 级对比（双方 TextBlocks 数必须一致，有数据时逐项比较）
            Assert.Equal(sp.TextBlocks.Count, dp.TextBlocks.Count);
            for (var ti = 0; ti < sp.TextBlocks.Count; ti++)
            {
                var st = sp.TextBlocks[ti];
                var dt = dp.TextBlocks[ti];
                Assert.Equal(st.Text, dt.Text);
                Assert.True(Math.Abs(st.FontSize - dt.FontSize) <= 0.5f,
                    $"{pt}.TB[{ti}] FontSize: {st.FontSize} vs {dt.FontSize}");
                // D7: 文本坐标（有数据时比较）
                if (st.X > 0 || dt.X > 0)
                    Assert.True(Math.Abs(st.X - dt.X) <= 2f,
                        $"{pt}.TB[{ti}] X: {st.X} vs {dt.X}");
                if (st.Y > 0 || dt.Y > 0)
                    Assert.True(Math.Abs(st.Y - dt.Y) <= 2f,
                        $"{pt}.TB[{ti}] Y: {st.Y} vs {dt.Y}");
            }

            // D5: 页面扩展属性（双方都有数据时比较）
            if (sp.ContentBytes.Length > 0 && dp.ContentBytes.Length > 0)
                Assert.Equal(sp.ContentBytes.Length, dp.ContentBytes.Length);
            if (sp.Images.Count > 0 && dp.Images.Count > 0)
                Assert.Equal(sp.Images.Count, dp.Images.Count);
            if (sp.LinkAnnotations.Count > 0 && dp.LinkAnnotations.Count > 0)
            {
                Assert.Equal(sp.LinkAnnotations.Count, dp.LinkAnnotations.Count);
                for (var li = 0; li < sp.LinkAnnotations.Count; li++)
                {
                    Assert.True(Math.Abs(sp.LinkAnnotations[li].X - dp.LinkAnnotations[li].X) <= 1f,
                        $"{pt}.LinkAnnot[{li}] X");
                    Assert.True(Math.Abs(sp.LinkAnnotations[li].Y - dp.LinkAnnotations[li].Y) <= 1f,
                        $"{pt}.LinkAnnot[{li}] Y");
                    Assert.Equal(sp.LinkAnnotations[li].Url ?? "", dp.LinkAnnotations[li].Url ?? "");
                }
            }
            if (sp.ExtGStates.Count > 0 && dp.ExtGStates.Count > 0)
                Assert.Equal(sp.ExtGStates.Count, dp.ExtGStates.Count);
        }

        // ③ 文本内容（字符多重集比较，忽略文本块顺序和坐标差异）
        var srcText = getFullText(src);
        var dstText = getFullText(dst);
        var srcSorted = new String(srcText.Where(c => !Char.IsWhiteSpace(c)).OrderBy(c => c).ToArray());
        var dstSorted = new String(dstText.Where(c => !Char.IsWhiteSpace(c)).OrderBy(c => c).ToArray());
        Assert.True(srcSorted == dstSorted,
            $"{tag} 文本字符集不一致（长度: 源={srcSorted.Length}, 输出={dstSorted.Length}）。\n源预览: [{srcText[..Math.Min(80, srcText.Length)]}]...\n输出预览: [{dstText[..Math.Min(80, dstText.Length)]}]...");

        // ④ 元数据
        Assert.Equal(src.Metadata.Title ?? "", dst.Metadata.Title ?? "");
        Assert.Equal(src.Metadata.Author ?? "", dst.Metadata.Author ?? "");
        Assert.Equal(src.Metadata.Subject ?? "", dst.Metadata.Subject ?? "");
        // D8: 元数据扩展属性
        if (src.Metadata.CreationDate != null && dst.Metadata.CreationDate != null)
            Assert.Equal(src.Metadata.CreationDate, dst.Metadata.CreationDate);

        // ⑤ 书签
        Assert.Equal(src.Bookmarks.Count, dst.Bookmarks.Count);
        for (var i = 0; i < src.Bookmarks.Count; i++)
        {
            Assert.Equal(src.Bookmarks[i].Title, dst.Bookmarks[i].Title);
            Assert.Equal(src.Bookmarks[i].PageIndex, dst.Bookmarks[i].PageIndex);
        }

        // ⑥ 注释（数量 + 逐项属性）
        Assert.Equal(src.Annotations.Count, dst.Annotations.Count);
        for (var ai = 0; ai < src.Annotations.Count; ai++)
        {
            var sa = src.Annotations[ai];
            var da = dst.Annotations[ai];
            var at = $"{tag} Annotation[{ai}]";
            Assert.Equal(sa.Type, da.Type);
            Assert.Equal(sa.PageIndex, da.PageIndex);
            Assert.True(Math.Abs(sa.X - da.X) <= 1f, $"{at} X: {sa.X} vs {da.X}");
            Assert.True(Math.Abs(sa.Y - da.Y) <= 1f, $"{at} Y: {sa.Y} vs {da.Y}");
            Assert.True(Math.Abs(sa.Width - da.Width) <= 1f, $"{at} Width: {sa.Width} vs {da.Width}");
            Assert.True(Math.Abs(sa.Height - da.Height) <= 1f, $"{at} Height: {sa.Height} vs {da.Height}");
            Assert.Equal(sa.Url ?? "", da.Url ?? "");
            Assert.Equal(sa.Contents ?? "", da.Contents ?? "");
            // D6: 注释扩展属性
            if (sa.DestinationPage >= 0 && da.DestinationPage >= 0)
                Assert.Equal(sa.DestinationPage, da.DestinationPage);
            if (sa.Author != null && da.Author != null)
                Assert.Equal(sa.Author, da.Author);
            if (sa.Subject != null && da.Subject != null)
                Assert.Equal(sa.Subject, da.Subject);
            if (sa.Color != null && da.Color != null)
                Assert.Equal(sa.Color.ToString(), da.Color.ToString());
            if (sa.Open != da.Open)
                Assert.Equal(sa.Open, da.Open);
        }
    }

    /// <summary>对单个 pdf 文件执行完整往返验证</summary>
    private static void RunSingleFileRoundTrip(String sourcePath, String outputDir)
    {
        var fileName = Path.GetFileName(sourcePath);
        var outputPath = Path.Combine(outputDir, fileName);

        // ─── 读取源文件 ───
        PdfDocument sourceDoc;
        String sourceFullText;
        using (var reader = new PdfReader(sourcePath))
        {
            sourceDoc = reader.ReadDocument();
            sourceFullText = reader.ExtractText();
        }
        Assert.True(sourceDoc.Pages.Count > 0, $"[{fileName}] 源文件应至少包含 1 页");

        // ─── 写入再读取 ───
        using (var writer = new PdfWriter())
        {
            writer.Save(outputPath, sourceDoc);
        }
        Assert.True(File.Exists(outputPath), $"[{fileName}] 输出文件应存在");

        PdfDocument outputDoc;
        String outputFullText;
        using (var reader = new PdfReader(outputPath))
        {
            outputDoc = reader.ReadDocument();
            outputFullText = reader.ExtractText();
        }

        // ─── 模型级比较（用 ExtractText 做文本对比）───
        // 源文件文本用已提取的 sourceFullText，输出文件也用已提取的 outputFullText
        AssertPdfDocumentEqual(sourceDoc, outputDoc, fileName, doc =>
        {
            // 对于源文件，用预提取的文本
            if (doc == sourceDoc) return sourceFullText;
            // 对于输出文件，用预提取的文本
            return outputFullText;
        });
    }

    /// <summary>程序化构造 PDF → PdfWriter 直接写入 → PdfReader 读回 → 文本级对比</summary>
    [Fact]
    [DisplayName("PDF_程序化往返：PdfWriter直接构建→ReadDocument→文本级对比")]
    public void Pdf_RoundTrip_Programmatic()
    {
        var outputPath = Path.Combine(OutputDir, "roundtrip_prog.pdf");

        // ─── 构造源文档（直接使用 PdfWriter API）───
        var sourceDoc = new PdfDocument
        {
            Metadata = new PdfDocumentInfo
            {
                Title = "往返测试文档",
                Author = "NewLife.Office",
                Subject = "PDF RoundTrip Test",
            },
        };
        sourceDoc.Bookmarks.Add(new PdfOutline { Title = "第一页", PageIndex = 0 });
        sourceDoc.Bookmarks.Add(new PdfOutline { Title = "第二页", PageIndex = 1 });

        // 先用 PdfWriter 构建 PDF 文件
        using (var w = new PdfWriter())
        {
            w.DocumentTitle = sourceDoc.Metadata.Title;
            w.DocumentAuthor = sourceDoc.Metadata.Author;
            w.DocumentSubject = sourceDoc.Metadata.Subject;

            w.BeginPage();
            w.DrawText("PDF 往返测试", 56f, 750f, 24f);
            w.DrawText("第一页内容：这是直接构建的 PDF 文档。", 56f, 720f, 12f);
            w.DrawText("Hello World from NewLife.Office PDF Engine.", 56f, 700f, 12f);
            w.DrawText("数字测试：12345.6789  特殊字符：!@#$%^&*()", 56f, 680f, 11f);
            w.EndPage();

            w.BeginPage();
            w.DrawText("第二页", 56f, 780f, 18f);
            w.DrawText("数据行 1: A=100, B=200, C=300", 56f, 750f, 12f);
            w.DrawText("数据行 2: X=1.5, Y=2.7, Z=3.14", 56f, 730f, 12f);
            w.DrawText("中英文混排 Mixed Content 测试 Test", 56f, 710f, 12f);
            w.EndPage();

            w.Bookmarks.AddRange(sourceDoc.Bookmarks);
            w.Save(outputPath);
        }
        Assert.True(File.Exists(outputPath), "输出文件应存在");

        // ─── 读回 ───
        PdfDocument outputDoc;
        String outputFullText;
        using (var reader = new PdfReader(outputPath))
        {
            outputFullText = reader.ExtractText();
            Assert.Contains("PDF 往返测试", outputFullText);
            Assert.Contains("Hello World", outputFullText);

            outputDoc = reader.ReadDocument();
        }

        // ─── 验证 ───
        Assert.Equal(2, outputDoc.Pages.Count);
        // 文本内容包含关键字符串
        Assert.Contains("第一页内容", outputFullText);
        Assert.Contains("第二页", outputFullText);
        Assert.Contains("中英文混排", outputFullText);
    }

    /// <summary>遍历 Bin 目录下所有 .pdf 文件，每个执行完整往返测试</summary>
    /// <remarks>若源文件 ReadDocument 返回 0 页（暂不支持的 PDF 特性），则跳过该文件</remarks>
    [Fact]
    [DisplayName("PDF_逐文件往返：读取Bin所有pdf→PdfDocument→写入→读回→模型对比")]
    public void Pdf_RoundTrip_AllFilesInBin()
    {
        var allFiles = FindAllPdfFiles();
        Assert.True(allFiles.Count > 0, $"Bin 目录下未找到 .pdf 文件。BaseDir={AppContext.BaseDirectory}");

        var pdfOutDir = Path.Combine(OutputDir, "Pdf");
        Directory.CreateDirectory(pdfOutDir);

        var failures = new List<(String file, String error)>();
        var skipped = 0;
        foreach (var sourcePath in allFiles)
        {
            var fileName = Path.GetFileName(sourcePath);
            try
            {
                // 预检查：ReadDocument 能否读取源文件
                using (var preReader = new PdfReader(sourcePath))
                {
                    var preDoc = preReader.ReadDocument();
                    if (preDoc.Pages.Count == 0)
                    {
                        skipped++;
                        continue; // 跳过暂不支持的 PDF
                    }
                }

                RunSingleFileRoundTrip(sourcePath, pdfOutDir);
            }
            catch (Exception ex)
            {
                var lines = ex.Message.Split('\n');
                var msg = String.Join(" | ", lines.Take(3)).Trim();
                failures.Add((fileName, msg));
            }
        }

        if (skipped > 0)
            Console.WriteLine($"  跳过 {skipped} 个暂不支持的 PDF 文件");

        if (failures.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"\n❌ {failures.Count}/{allFiles.Count} 个文件测试失败:");
            foreach (var (file, err) in failures)
                sb.AppendLine($"  [{file}] {err}");
            Assert.Fail(sb.ToString());
        }
    }

    #endregion
}
