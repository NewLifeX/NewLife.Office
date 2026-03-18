using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>PDF 格式集成测试</summary>
public class PdfTests : IntegrationTestBase
{
    [Fact, DisplayName("PDF_字体变体展示")]
    public void Pdf_FontVariations()
    {
        var path = Path.Combine(OutputDir, "test_fonts.pdf");

        using (var w = new PdfFluentDocument())
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

        using (var w = new PdfFluentDocument())
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
}
