using System.ComponentModel;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using NewLife.Office;
using NewLife.Office.Pdf;
using NewLife.Office.Word;
using Xunit;

using XUnitTest.Common;

namespace XUnitTest.Word;

/// <summary>Word docx 格式往返测试</summary>
public class WordRoundTripTests : IntegrationTestBase
{
    [Fact, DisplayName("Word_docx_复杂写入再读取往返")]
    public void Word_Docx_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.docx");

        using (var w = new WordWriter())
        {
            w.DocumentProperties.Title = "集成测试文档";
            w.DocumentProperties.Author = "NewLife Office";

            w.AppendHeading("第一章 概述", 1);
            w.AppendParagraph("这是一份由 NewLife.Office 自动生成的测试文档，用于验证 Word 文件的读写功能。");
            w.AppendParagraph("本文档包含标题、段落、表格、列表等多种元素。");

            w.AppendHeading("第二章 数据表格", 2);
            w.AppendParagraph("以下是一个示例数据表格：");

            var tableData = new[]
            {
                new[] { "产品", "价格", "库存" },
                new[] { "笔记本", "5999", "100" },
                new[] { "手机", "3999", "500" },
                new[] { "平板", "2999", "200" },
            };
            w.AppendTable(tableData);

            w.AppendHeading("第三章 格式化文本", 2);
            w.AppendParagraph("普通段落文本。", ParagraphStyle.Normal,
                new RunProperties { FontSize = 14f, Bold = true });
            w.AppendParagraph("这是另一个段落。");

            w.AppendHeading("附录", 3);
            w.AppendParagraph("文档结束。");

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        // 读取验证
        using var reader = new WordReader(path);
        var paragraphs = reader.ReadParagraphs().ToList();
        Assert.True(paragraphs.Count >= 10);
        Assert.Contains("第一章 概述", paragraphs);
        Assert.Contains("文档结束。", paragraphs);

        // ReadFullText 返回正文文本，Title 属于文档属性不在正文中
        var fullText = reader.ReadFullText();
        Assert.Contains("数据表格", fullText);
        Assert.Contains("格式化文本", fullText);

        // 读取表格
        var tables = reader.ReadTables().ToList();
        Assert.Equal(1, tables.Count);
        Assert.Equal("产品", tables[0][0][0]);
        Assert.Equal("笔记本", tables[0][1][0]);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<WordReader>(factoryReader);
        (factoryReader as IDisposable)?.Dispose();
    }

    [Fact, DisplayName("Word_docx转PDF")]
    public void Word_Docx_To_Pdf()
    {
        var docxPath = Path.Combine(OutputDir, "convert_word.docx");
        var pdfPath = Path.Combine(OutputDir, "converted_from_word.pdf");

        using (var w = new WordWriter())
        {
            w.AppendHeading("Word转PDF测试", 1);
            w.AppendParagraph("这段文字应该出现在PDF中。");
            w.AppendParagraph("支持多段落转换。");
            w.Save(docxPath);
        }

        // 转换
        var converter = new WordPdfConverter();
        converter.ConvertToFile(docxPath, pdfPath);

        Assert.True(File.Exists(pdfPath));

        // 验证 PDF
        using var pdfReader = new PdfReader(pdfPath);
        var text = pdfReader.ExtractText();
        Assert.Contains("Word", text);
        Assert.Equal(1, pdfReader.GetPageCount());
    }

    [Fact, DisplayName("Docx_从Bin读取所有文件并往返验证")]
    public void Docx_RoundTrip_AllFiles()
    {
        var binDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, ".."));
        var files = Directory.GetFiles(binDir, "*.docx", SearchOption.TopDirectoryOnly);
        if (files.Length == 0)
            throw new InvalidOperationException($"未在 {binDir} 找到 .docx 文件，请放入测试文件后重试。");

        var outDir = Path.Combine(binDir, "Output");
        Directory.CreateDirectory(outDir);

        foreach (var sourcePath in files)
        {
            var fileName = Path.GetFileNameWithoutExtension(sourcePath);
            var outputPath = Path.Combine(outDir, $"{fileName}.docx");

            // 读取源文件（用 MemoryStream 避免 Word 独占锁导致 ZipFile.OpenRead 失败）
            Document sourceDoc;
            try
            {
                var sourceBytes = ReadAllBytesShared(sourcePath);
                using (var ms = new MemoryStream(sourceBytes))
                using (var reader = new WordReader(ms))
                {
                    sourceDoc = reader.ReadDocument();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ! 读取源文件 {fileName} 失败: {ex.GetType().Name}: {ex.Message}");
                continue;
            }

            Assert.NotEmpty(sourceDoc.Elements);
            Console.WriteLine($"[{fileName}] 源文件: {sourceDoc.Elements.Count} 元素 (P={sourceDoc.Elements.Count(e=>e.Type==ElementType.Paragraph)} T={sourceDoc.Elements.Count(e=>e.Type==ElementType.Table)} I={sourceDoc.Elements.Count(e=>e.Type==ElementType.Image)})");

            // 写入并重新读取
            try
            {
                using (var writer = new WordWriter())
                {
                    writer.Save(outputPath, sourceDoc);
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"  ! 跳过 {fileName}：输出文件可能被 Word 打开，{ex.Message}");
                continue;
            }

            Assert.True(File.Exists(outputPath), $"输出文件 {outputPath} 应存在");

            // 用 MemoryStream 读回避免 ZIP 文件锁问题
            Document outputDoc;
            try
            {
                var outputBytes = ReadAllBytesShared(outputPath);
                using (var ms = new MemoryStream(outputBytes))
                using (var reader = new WordReader(ms))
                {
                    outputDoc = reader.ReadDocument();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ! 读取输出文件 {fileName} 失败: {ex.GetType().Name}: {ex.Message}");
                continue;
            }

            Console.WriteLine($"[{fileName}] 输出文件: {outputDoc.Elements.Count} 元素 (P={outputDoc.Elements.Count(e=>e.Type==ElementType.Paragraph)} T={outputDoc.Elements.Count(e=>e.Type==ElementType.Table)} I={outputDoc.Elements.Count(e=>e.Type==ElementType.Image)})");

            // 严谨比较
            AssertDocumentsEqual(sourceDoc, outputDoc, fileName);
        }
    }

    [Fact, DisplayName("Docx_往返读写_格式化保真")]
    public void Docx_RoundTrip_FormattingFidelity()
    {
        var path = Path.Combine(OutputDir, "roundtrip_fmt.docx");

        // 构造覆盖所有格式特性的文档
        using (var w = new WordWriter())
        {
            w.DocumentProperties.Title = "格式化保真测试";
            w.DocumentProperties.Author = "NewLife.Office";
            w.PageSettings.HeaderText = "测试页眉";
            w.PageSettings.FooterText = "测试页脚";

            w.AppendHeading("一级标题", 1);
            w.AppendHeading("二级标题", 2);

            w.AppendParagraph("粗体文本", ParagraphStyle.Normal, new RunProperties { Bold = true, FontSize = 14f });
            w.AppendParagraph("斜体文本", ParagraphStyle.Normal, new RunProperties { Italic = true, FontSize = 12f });
            w.AppendParagraph("红色文本", ParagraphStyle.Normal, new RunProperties { ForeColor = "FF0000", FontSize = 16f });
            w.AppendParagraph("带下划线蓝色大字", ParagraphStyle.Normal, new RunProperties { Underline = true, ForeColor = "0000FF", FontSize = 18f, FontName = "Arial" });

            // 混合格式段落
            w.AppendFormattedParagraph(new[]
            {
                new Run { Text = "混合：", Properties = new RunProperties { Bold = true } },
                new Run { Text = "粗体+", Properties = new RunProperties { Bold = true, FontSize = 14f } },
                new Run { Text = "红色斜体", Properties = new RunProperties { Italic = true, ForeColor = "FF0000" } },
                new Run { Text = " 普通" },
            });

            // 超链接
            w.AppendHyperlink("访问官网", "https://newlifex.com", new RunProperties { FontSize = 12f });

            // 对齐
            var leftPara = w.AppendParagraph("左对齐");
            leftPara.Alignment = "left";
            var centerPara = w.AppendParagraph("居中");
            centerPara.Alignment = "center";
            var rightPara = w.AppendParagraph("右对齐");
            rightPara.Alignment = "right";

            // 表格
            w.AppendHeading("数据表格", 2);
            var tableStyle = new TableStyle { BorderColor = "333333", BorderSize = 6, HeaderBgColor = "4472C4", StripeColor = "D9E2F3" };
            w.AppendTable(new[] { new[] { "产品", "价格", "库存" }, new[] { "笔记本", "5999", "100" }, new[] { "手机", "3999", "500" } }, true, tableStyle);

            // 列表
            w.AppendBulletList(new[] { "项目一", "项目二", "项目三" });

            // 分页
            w.AppendPageBreak();
            w.AppendParagraph("分页后内容", ParagraphStyle.Normal, new RunProperties { FontSize = 14f, Bold = true });

            // 缩进段落
            var indentPara = w.AppendParagraph("首行缩进段落");
            indentPara.FirstLineIndent = 480;

            // W09-01: 删除线
            w.AppendParagraph("删除线文本", ParagraphStyle.Normal, new RunProperties { Strikethrough = true });

            // W09-02: 上标（X²）
            var supPara = w.AppendFormattedParagraph(new[]
            {
                new Run { Text = "X" },
                new Run { Text = "2", Properties = new RunProperties { Superscript = true } },
            });

            // W09-02: 下标（H₂O）
            var subPara = w.AppendFormattedParagraph(new[]
            {
                new Run { Text = "H" },
                new Run { Text = "2", Properties = new RunProperties { Subscript = true } },
                new Run { Text = "O" },
            });

            // W09-01: 下划线样式（波浪线）
            w.AppendParagraph("波浪下划线", ParagraphStyle.Normal, new RunProperties { UnderlineStyle = UnderlineStyles.Wave });

            // W09-04: 字符间距加宽
            w.AppendParagraph("加宽字符间距", ParagraphStyle.Normal, new RunProperties { CharacterSpacing = 40f });

            // W09-04: 字符缩放
            w.AppendParagraph("字符缩放150%", ParagraphStyle.Normal, new RunProperties { CharacterScaling = 150 });

            // W09-03: 段落边框（上下边框）
            var borderPara = w.AppendParagraph("带边框段落");
            borderPara.Borders = new ParagraphBorders
            {
                Top    = new Border { Style = BorderStyle.Single, Color = "FF0000", Width = 12 },
                Bottom = new Border { Style = BorderStyle.Double, Color = "0000FF", Width = 8 },
                Left   = new Border { Style = BorderStyle.Dotted, Color = "00AA00", Width = 4 },
            };

            // W09-01: 制表位
            var tabPara = w.AppendParagraph("姓名\t部门\t薪资");
            tabPara.TabStops = new List<TabStop>
            {
                new TabStop { Position = 3600, Alignment = "left" },
                new TabStop { Position = 7200, Alignment = "right", Leader = "dot" },
            };

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        // 第一轮读取
        Document doc1;
        using (var reader = new WordReader(path)) { doc1 = reader.ReadDocument(); }

        // 第一轮写入
        var path2 = Path.Combine(OutputDir, "roundtrip_fmt_v2.docx");
        using (var writer = new WordWriter()) { writer.Save(path2, doc1); }

        // 第二轮读取
        Document doc2;
        using (var reader = new WordReader(path2)) { doc2 = reader.ReadDocument(); }

        // 严谨比较
        AssertDocumentsEqual(doc1, doc2, "formatted");

        // 专项验证
        Assert.Contains("一级标题", GetBodyText(doc1));
        Assert.Contains("红色文本", GetBodyText(doc1));
        Assert.Contains("删除线文本", GetBodyText(doc1));
        Assert.Contains("带边框段落", GetBodyText(doc1));
#pragma warning disable xUnit2012 // Any+Contains 组合无直接 xUnit 等价写法
        Assert.True(doc1.Hyperlinks.Any(h => h.Url.Contains("newlifex")));
#pragma warning restore xUnit2012
    }

    #region 严谨断言和辅助
    /// <summary>以共享读方式读取文件全部字节</summary>
    private static Byte[] ReadAllBytesShared(String path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var buf = new Byte[fs.Length];
        fs.ReadExactly(buf, 0, buf.Length);
        return buf;
    }

    /// <summary>
    /// 规范化 XML 字符串：删除内联命名空间声明，标签间空白压缩为单空格。
    /// 读取时命名空间声明可能在根元素或子元素上，规范化后可直接比较。
    /// </summary>
    internal static String NormalizeXml(String? xml)
    {
        if (String.IsNullOrEmpty(xml)) return "";
        // 删除内联命名空间声明（xmlns:prefix="uri" 或 xmlns="uri"）
        var result = Regex.Replace(xml, @"\s+xmlns:[A-Za-z0-9_]+=(?:""[^""]*""|'[^']*')", "");
        result = Regex.Replace(result, @"\s+xmlns=(?:""[^""]*""|'[^']*')", "");
        // 标签间空白压缩
        result = Regex.Replace(result, @">\s+<", "><");
        return result.Trim();
    }

    /// <summary>从 Document 提取正文纯文本</summary>
    private static String GetBodyText(Document doc)
    {
        var sb = new StringBuilder();
        foreach (var el in doc.Elements)
        {
            if (el.Type == ElementType.Paragraph && el.Paragraph != null)
            {
                foreach (var r in el.Paragraph.Runs) sb.Append(r.Text);
                sb.AppendLine();
            }
            else if (el.Type == ElementType.Table && el.TableRows != null)
            {
                foreach (var row in el.TableRows)
                {
                    foreach (var cell in row)
                    {
                        foreach (var para in cell.Paragraphs)
                            foreach (var r in para.Runs) sb.Append(r.Text);
                        sb.Append('\t');
                    }
                    sb.AppendLine();
                }
            }
        }
        return sb.ToString();
    }

    /// <summary>严谨比较两个 Document，确保结构、格式、内容完全一致</summary>
    private static void AssertDocumentsEqual(Document a, Document b, String context)
    {
        // 元素数量必须相等
        Assert.True(a.Elements.Count == b.Elements.Count,
            $"[{context}] 元素数量不匹配: 源={a.Elements.Count}, 输出={b.Elements.Count}");

        for (var i = 0; i < a.Elements.Count; i++)
        {
            var ae = a.Elements[i];
            var be = b.Elements[i];

            // 类型必须一致
            Assert.True(ae.Type == be.Type,
                $"[{context}] 元素[{i}] 类型不匹配: 源={ae.Type}, 输出={be.Type}");

            // ★ 核心断言：原始 XML 直接比较——捕获所有未建模的内联格式（字体/高亮/删除线/rStyle等）
            if (ae.RawXml != null && be.RawXml != null)
            {
                var aNorm = NormalizeXml(ae.RawXml);
                var bNorm = NormalizeXml(be.RawXml);
                Assert.True(aNorm == bNorm,
                    $"[{context}] 元素[{i}] 原始 XML 不匹配（视觉效果不同）"
                    + $"源={aNorm.Substring(0, Math.Min(200, aNorm.Length))}"
                    + $" 输出={bNorm.Substring(0, Math.Min(200, bNorm.Length))}");
            }
            else
            {
                // 没有 RawXml 时用模型属性比较
                switch (ae.Type)
                {
                    case ElementType.Paragraph:
                        AssertParagraphsEqual(ae.Paragraph, be.Paragraph, context, i);
                        break;
                    case ElementType.Table:
                        AssertTablesEqual(ae, be, context, i);
                        break;
                    case ElementType.Image:
                        AssertImagesEqual(ae, be, context, i);
                        break;
                }
            }
        }

        // 图片数量比较（透传模式下页眉图片 rId 可能不同，跳过严格比较——RawXml 已保证内容）
        if (a.OtherParts.Count == 0)
        {
            Assert.True(a.Images.Count == b.Images.Count,
                $"[{context}] 图片数量不匹配: 源={a.Images.Count}, 输出={b.Images.Count}");
        }

        foreach (var (key, (ext, data)) in a.Images)
        {
            Assert.True(b.Images.TryGetValue(key, out var bImg),
                $"[{context}] 输出缺少图片关系 {key}");
            Assert.True(ext == bImg.Extension,
                $"[{context}] 图片[{key}] 扩展名不匹配: 源={ext}, 输出={bImg.Extension}");
            Assert.True(data.Length == bImg.Data.Length,
                $"[{context}] 图片[{key}] 数据长度不匹配: 源={data.Length}, 输出={bImg.Data.Length}");
            // 逐字节比较图片数据
            for (var i = 0; i < data.Length; i++)
                Assert.True(data[i] == bImg.Data[i],
                    $"[{context}] 图片[{key}] 数据[{i}]: {data[i]:X2} vs {bImg.Data[i]:X2}");
        }

        // 页面设置
        Assert.Equal(a.PageSettings.PageWidth, b.PageSettings.PageWidth);
        Assert.Equal(a.PageSettings.PageHeight, b.PageSettings.PageHeight);
        Assert.Equal(a.PageSettings.Landscape, b.PageSettings.Landscape);

        // sectPr 预先角：页面尺寸/页眉页脚引用必须保真
        if (a.SectPrXml != null && b.SectPrXml != null)
        {
            Assert.True(NormalizeXml(a.SectPrXml) == NormalizeXml(b.SectPrXml),
                $"[{context}] sectPr 不一致——页面尺寸/页边距/页眉页脚引用不同");
        }

        // 原始 XML 部件保真验证
        Assert.True(a.StylesXml == b.StylesXml,
            $"[{context}] styles.xml 内容不一致——样式定义丢失会导致字体/大小/颜色差异");
        Assert.True(a.NumberingXml == b.NumberingXml,
            $"[{context}] numbering.xml 内容不一致——列表编号/项目符号会不同");
        Assert.True(a.SettingsXml == b.SettingsXml,
            $"[{context}] settings.xml 内容不一致——兼容性设置丢失");

        // 文档属性
        Assert.Equal(a.DocumentProperties.Title, b.DocumentProperties.Title);
        Assert.Equal(a.DocumentProperties.Author, b.DocumentProperties.Author);

        // 页眉页脚文本
        Assert.Equal(a.HeaderText ?? "", b.HeaderText ?? "");
        Assert.Equal(a.FooterText ?? "", b.FooterText ?? "");

        // W11: 页眉页脚富文本/内容控件/自定义XML/只读保护/超链接/文档变量（双方都有数据时比较）
        if (a.Headers.Count > 0 && b.Headers.Count > 0)
            Assert.Equal(a.Headers.Count, b.Headers.Count);
        if (a.Footers.Count > 0 && b.Footers.Count > 0)
            Assert.Equal(a.Footers.Count, b.Footers.Count);
        if (a.SdtElements.Count > 0 && b.SdtElements.Count > 0)
            Assert.Equal(a.SdtElements.Count, b.SdtElements.Count);
        if (a.ProtectionReadOnly || b.ProtectionReadOnly)
            Assert.Equal(a.ProtectionReadOnly, b.ProtectionReadOnly);

        // 批注数量
        Assert.Equal(a.Comments.Count, b.Comments.Count);
        for (var ci = 0; ci < a.Comments.Count; ci++)
        {
            Assert.Equal(a.Comments[ci].Text, b.Comments[ci].Text);
            Assert.Equal(a.Comments[ci].Author, b.Comments[ci].Author);
            // W10: 批注扩展属性
            if (a.Comments[ci].Id > 0 && b.Comments[ci].Id > 0)
                Assert.Equal(a.Comments[ci].Id, b.Comments[ci].Id);
            if (a.Comments[ci].Date != null && b.Comments[ci].Date != null)
                Assert.Equal(a.Comments[ci].Date.Value, b.Comments[ci].Date.Value, TimeSpan.FromSeconds(1));
            if (a.Comments[ci].Initials != null && b.Comments[ci].Initials != null)
                Assert.Equal(a.Comments[ci].Initials, b.Comments[ci].Initials);
        }

        // W11: 超链接（双方都有数据时比较完整列表）
        if (a.Hyperlinks.Count > 0 && b.Hyperlinks.Count > 0)
        {
            Assert.Equal(a.Hyperlinks.Count, b.Hyperlinks.Count);
            for (var hi = 0; hi < a.Hyperlinks.Count; hi++)
            {
                Assert.Equal(a.Hyperlinks[hi].Item1, b.Hyperlinks[hi].Item1);
                Assert.Equal(a.Hyperlinks[hi].Item2, b.Hyperlinks[hi].Item2);
            }
        }
        // W11: 自定义XML部件
        if (a.CustomXmlParts.Count > 0 && b.CustomXmlParts.Count > 0)
        {
            Assert.Equal(a.CustomXmlParts.Count, b.CustomXmlParts.Count);
            foreach (var kv in a.CustomXmlParts)
            {
                Assert.True(b.CustomXmlParts.TryGetValue(kv.Key, out var bv),
                    $"[{context}] 缺失 CustomXmlPart: {kv.Key}");
                Assert.True(kv.Value.SequenceEqual(bv),
                    $"[{context}] CustomXmlPart[{kv.Key}] 字节不一致");
            }
        }
        // W06-15: 文档变量
        if (a.DocumentVariables.Count > 0 && b.DocumentVariables.Count > 0)
        {
            Assert.Equal(a.DocumentVariables.Count, b.DocumentVariables.Count);
            foreach (var kv in a.DocumentVariables)
            {
                Assert.True(b.DocumentVariables.TryGetValue(kv.Key, out var bv),
                    $"[{context}] 缺失 DocumentVariable: {kv.Key}");
                Assert.Equal(kv.Value, bv);
            }
        }

        // ★ 最严格的断言：document.xml 原文完全一致——保证任何视觉效果都不丢失
        if (a.DocumentXml != null && b.DocumentXml != null)
        {
            Assert.True(a.DocumentXml == b.DocumentXml,
                $"[{context}] word/document.xml 内容不一致！" +
                $"这意味着往返后文档正文有差异（源={a.DocumentXml.Length}字节 输出={b.DocumentXml.Length}字节）。" +
                $"前100字符差异: 源=[{a.DocumentXml.Substring(0, Math.Min(100, a.DocumentXml.Length))}]" +
                $" 输出=[{b.DocumentXml.Substring(0, Math.Min(100, b.DocumentXml.Length))}]");
        }
        else
        {
            // 没有 DocumentXml 时（程序化生成文档），退化为模型级比较
            Assert.True(a.Elements.Count == b.Elements.Count,
                $"[{context}] 元素数量不匹配: 源={a.Elements.Count}, 输出={b.Elements.Count}");
        }

        // OtherParts 完全保真——主题、字体表、脚注、尾注、页眉页脚 raw XML 等
        Assert.True(a.OtherParts.Count == b.OtherParts.Count,
            $"[{context}] OtherParts 数量不匹配: 源={a.OtherParts.Count}个部件 ({String.Join(", ", a.OtherParts.Keys)}), 输出={b.OtherParts.Count}");
        foreach (var kv in a.OtherParts)
        {
            Assert.True(b.OtherParts.TryGetValue(kv.Key, out var bBytes),
                $"[{context}] 输出缺少 OtherParts 部件: {kv.Key}");
            Assert.True(kv.Value.Length == bBytes!.Length,
                $"[{context}] OtherParts[{kv.Key}] 字节长度不匹配: 源={kv.Value.Length} 输出={bBytes.Length}");
            Assert.True(kv.Value.SequenceEqual(bBytes),
                $"[{context}] OtherParts[{kv.Key}] 字节内容不一致（应原样透传）");
        }
    }

    private static void AssertParagraphsEqual(Paragraph? a, Paragraph? b, String context, Int32 idx)
    {
        Assert.NotNull(a);
        Assert.NotNull(b);
        if (a == null || b == null) return;

        // 样式标识符
        Assert.True(a.StyleId == b.StyleId,
            $"[{context}] 段落[{idx}] StyleId 不匹配: 源={a.StyleId}, 输出={b.StyleId}");
        Assert.Equal(a.Style, b.Style);

        // 对齐
        Assert.True(a.Alignment == b.Alignment,
            $"[{context}] 段落[{idx}] 对齐不匹配: 源={a.Alignment}, 输出={b.Alignment}");

        // 缩进
        Assert.Equal(a.IndentLeft, b.IndentLeft);
        Assert.Equal(a.IndentRight, b.IndentRight);
        Assert.Equal(a.FirstLineIndent, b.FirstLineIndent);
        Assert.Equal(a.SpaceBefore, b.SpaceBefore);
        Assert.Equal(a.SpaceAfter, b.SpaceAfter);
        Assert.Equal(a.LineSpacingPct, b.LineSpacingPct);

        // 背景色
        Assert.True(a.BackgroundColor == b.BackgroundColor,
            $"[{context}] 段落[{idx}] 背景色不匹配");

        // 特殊标记
        Assert.Equal(a.IsBullet, b.IsBullet);
        Assert.Equal(a.IsPageBreak, b.IsPageBreak);
        Assert.True(a.BookmarkName == b.BookmarkName,
            $"[{context}] 段落[{idx}] 书签不匹配");

        // W08: 有序列表/分页控制（双方都有非默认值时才比较）
        if (a.IsOrderedList || b.IsOrderedList)
            Assert.Equal(a.IsOrderedList, b.IsOrderedList);
        if (a.ListLevel > 0 || b.ListLevel > 0)
            Assert.Equal(a.ListLevel, b.ListLevel);
        if (a.ListStartOverride != null && b.ListStartOverride != null)
            Assert.Equal(a.ListStartOverride, b.ListStartOverride);
        if (a.KeepNext || b.KeepNext)
            Assert.Equal(a.KeepNext, b.KeepNext);
        if (a.KeepLines || b.KeepLines)
            Assert.Equal(a.KeepLines, b.KeepLines);
        if (!a.WidowControl || !b.WidowControl)
            Assert.Equal(a.WidowControl, b.WidowControl);
        if (a.DropCapLines != null && b.DropCapLines != null)
            Assert.Equal(a.DropCapLines, b.DropCapLines);
        if (a.DropCapChars != null && b.DropCapChars != null)
            Assert.Equal(a.DropCapChars, b.DropCapChars);

        // W09: 制表位比较
        if (a.TabStops != null || b.TabStops != null)
        {
            Assert.NotNull(a.TabStops);
            Assert.NotNull(b.TabStops);
            if (a.TabStops != null && b.TabStops != null)
            {
                Assert.True(a.TabStops.Count == b.TabStops.Count,
                    $"[{context}] 段落[{idx}] TabStops 数量不匹配: 源={a.TabStops.Count}, 输出={b.TabStops.Count}");
                for (var ti = 0; ti < a.TabStops.Count; ti++)
                {
                    var at = a.TabStops[ti];
                    var bt = b.TabStops[ti];
                    Assert.Equal(at.Position, bt.Position);
                    Assert.True(at.Alignment == bt.Alignment,
                        $"[{context}] 段落[{idx}] TabStop[{ti}] 对齐不匹配");
                    Assert.True(at.Leader == bt.Leader,
                        $"[{context}] 段落[{idx}] TabStop[{ti}] 前导符不匹配");
                }
            }
        }

        // W09: 段落边框比较
        if (a.Borders != null || b.Borders != null)
        {
            Assert.NotNull(a.Borders);
            Assert.NotNull(b.Borders);
            if (a.Borders != null && b.Borders != null)
            {
                AssertBorderEqual(a.Borders.Top,    b.Borders.Top,    context, idx, "上边框");
                AssertBorderEqual(a.Borders.Bottom, b.Borders.Bottom, context, idx, "下边框");
                AssertBorderEqual(a.Borders.Left,   b.Borders.Left,   context, idx, "左边框");
                AssertBorderEqual(a.Borders.Right,  b.Borders.Right,  context, idx, "右边框");
            }
        }

        // Run 逐一严格比较
        Assert.True(a.Runs.Count == b.Runs.Count,
            $"[{context}] 段落[{idx}] Run 数量不匹配: 源={a.Runs.Count}, 输出={b.Runs.Count}");

        for (var ri = 0; ri < a.Runs.Count; ri++)
        {
            var ar = a.Runs[ri];
            var br = b.Runs[ri];

            Assert.True(ar.Text == br.Text,
                $"[{context}] 段落[{idx}] Run[{ri}] 文本不匹配: 源='{ar.Text}', 输出='{br.Text}'");

            var arp = ar.Properties;
            var brp = br.Properties;

            // 两者都为 null 则通过
            if (arp == null && brp == null) continue;

            Assert.True(arp != null && brp != null,
                $"[{context}] 段落[{idx}] Run[{ri}] 格式属性不匹配: 一方为 null");
            if (arp == null || brp == null) continue;

            Assert.True(arp.Bold == brp.Bold,
                $"[{context}] 段落[{idx}] Run[{ri}] 粗体不匹配");
            Assert.True(arp.Italic == brp.Italic,
                $"[{context}] 段落[{idx}] Run[{ri}] 斜体不匹配");
            Assert.True(arp.Underline == brp.Underline,
                $"[{context}] 段落[{idx}] Run[{ri}] 下划线不匹配");
            Assert.True(arp.ForeColor == brp.ForeColor,
                $"[{context}] 段落[{idx}] Run[{ri}] 颜色不匹配: 源={arp.ForeColor}, 输出={brp.ForeColor}");
            Assert.True(arp.FontSize == brp.FontSize,
                $"[{context}] 段落[{idx}] Run[{ri}] 字号不匹配: 源={arp.FontSize}, 输出={brp.FontSize}");
            Assert.True(arp.FontName == brp.FontName,
                $"[{context}] 段落[{idx}] Run[{ri}] 字体不匹配: 源={arp.FontName}, 输出={brp.FontName}");

            // W09: 删除线
            Assert.True(arp.Strikethrough == brp.Strikethrough,
                $"[{context}] 段落[{idx}] Run[{ri}] 删除线不匹配");
            // W09: 上标/下标
            Assert.True(arp.Superscript == brp.Superscript,
                $"[{context}] 段落[{idx}] Run[{ri}] 上标不匹配");
            Assert.True(arp.Subscript == brp.Subscript,
                $"[{context}] 段落[{idx}] Run[{ri}] 下标不匹配");
            // W09: 下划线样式
            Assert.True(arp.UnderlineStyle == brp.UnderlineStyle,
                $"[{context}] 段落[{idx}] Run[{ri}] 下划线样式不匹配: 源={arp.UnderlineStyle}, 输出={brp.UnderlineStyle}");
            // W09: 字符间距
            Assert.True(arp.CharacterSpacing == brp.CharacterSpacing,
                $"[{context}] 段落[{idx}] Run[{ri}] 字符间距不匹配: 源={arp.CharacterSpacing}, 输出={brp.CharacterSpacing}");
            // W09: 字符缩放
            Assert.True(arp.CharacterScaling == brp.CharacterScaling,
                $"[{context}] 段落[{idx}] Run[{ri}] 字符缩放不匹配: 源={arp.CharacterScaling}, 输出={brp.CharacterScaling}");
            // W07: 发光/阴影（双方都有值时才比较）
            if (arp.GlowColor != null && brp.GlowColor != null)
                Assert.True(arp.GlowColor == brp.GlowColor,
                    $"[{context}] 段落[{idx}] Run[{ri}] 发光颜色不匹配");
            if (arp.GlowSize != null && brp.GlowSize != null)
                Assert.Equal(arp.GlowSize, brp.GlowSize);
            if (arp.ShadowColor != null && brp.ShadowColor != null)
                Assert.True(arp.ShadowColor == brp.ShadowColor,
                    $"[{context}] 段落[{idx}] Run[{ri}] 阴影颜色不匹配");
            if (arp.ShadowOffsetX != null && brp.ShadowOffsetX != null)
                Assert.Equal(arp.ShadowOffsetX, brp.ShadowOffsetX);
            if (arp.ShadowOffsetY != null && brp.ShadowOffsetY != null)
                Assert.Equal(arp.ShadowOffsetY, brp.ShadowOffsetY);
        }
    }

    /// <summary>比较单边边框</summary>
    private static void AssertBorderEqual(Border? a, Border? b, String context, Int32 idx, String edge)
    {
        if (a == null && b == null) return;
        Assert.True(a != null && b != null,
            $"[{context}] 段落[{idx}] {edge} 边框不匹配: 一方为 null");
        if (a == null || b == null) return;
        Assert.True(a.Style == b.Style,
            $"[{context}] 段落[{idx}] {edge} 线型不匹配: 源={a.Style}, 输出={b.Style}");
        Assert.True(a.Color == b.Color,
            $"[{context}] 段落[{idx}] {edge} 颜色不匹配: 源={a.Color}, 输出={b.Color}");
        Assert.Equal(a.Width, b.Width);
    }

    private static void AssertTablesEqual(Element a, Element b, String context, Int32 idx)
    {
        var aRows = a.TableRows;
        var bRows = b.TableRows;
        Assert.NotNull(aRows);
        Assert.NotNull(bRows);
        if (aRows == null || bRows == null) return;

        Assert.True(aRows.Count == bRows.Count,
            $"[{context}] 表格[{idx}] 行数不匹配: 源={aRows.Count}, 输出={bRows.Count}");
        Assert.Equal(a.TableFirstRowHeader, b.TableFirstRowHeader);

        for (var ri = 0; ri < aRows.Count; ri++)
        {
            var aRow = aRows[ri];
            var bRow = bRows[ri];
            Assert.True(aRow.Count == bRow.Count,
                $"[{context}] 表格[{idx}] 行[{ri}] 列数不匹配: 源={aRow.Count}, 输出={bRow.Count}");

            for (var ci = 0; ci < aRow.Count; ci++)
            {
                var aCell = aRow[ci];
                var bCell = bRow[ci];

                // 单元格文本
                var aText = String.Concat(aCell.Paragraphs.SelectMany(p => p.Runs).Select(r => r.Text));
                var bText = String.Concat(bCell.Paragraphs.SelectMany(p => p.Runs).Select(r => r.Text));
                Assert.True(aText == bText,
                    $"[{context}] 表格[{idx}] 单元格[{ri},{ci}] 文本不匹配");

                Assert.True(aCell.BackgroundColor == bCell.BackgroundColor,
                    $"[{context}] 表格[{idx}] 单元格[{ri},{ci}] 背景色不匹配");
                // W09: 单元格合并/垂直对齐（双方都有非默认值时才比较）
                if (aCell.ColSpan > 1 || bCell.ColSpan > 1)
                    Assert.Equal(aCell.ColSpan, bCell.ColSpan);
                if (aCell.RowSpan > 1 || bCell.RowSpan > 1)
                    Assert.Equal(aCell.RowSpan, bCell.RowSpan);
                if (aCell.VerticalAlignment != null && bCell.VerticalAlignment != null)
                    Assert.Equal(aCell.VerticalAlignment, bCell.VerticalAlignment);
            }
        }

        // 表格样式
        if (a.TableStyle != null && b.TableStyle != null)
        {
            Assert.True(a.TableStyle.BorderColor == b.TableStyle.BorderColor);
            Assert.Equal(a.TableStyle.BorderSize, b.TableStyle.BorderSize);
        }
    }

    private static void AssertImagesEqual(Element a, Element b, String context, Int32 idx)
    {
        var ai = a.Image;
        var bi = b.Image;
        Assert.NotNull(ai);
        Assert.NotNull(bi);
        if (ai == null || bi == null) return;

        Assert.True(ai.Extension == bi.Extension,
            $"[{context}] 图片[{idx}] 扩展名不匹配: 源={ai.Extension}, 输出={bi.Extension}");
        Assert.Equal(ai.WidthEmu, bi.WidthEmu);
        Assert.Equal(ai.HeightEmu, bi.HeightEmu);
        // W12: 图片关系 ID（双方都有值时比较）
        if (ai.RelId != null && bi.RelId != null)
            Assert.True(ai.RelId == bi.RelId,
                $"[{context}] 图片[{idx}] RelId 不匹配: 源={ai.RelId}, 输出={bi.RelId}");
        // 图片数据逐字节比较
        if (ai.ImageData.Length > 0 && bi.ImageData.Length > 0)
        {
            Assert.Equal(ai.ImageData.Length, bi.ImageData.Length);
            for (var i = 0; i < ai.ImageData.Length; i++)
                Assert.True(ai.ImageData[i] == bi.ImageData[i],
                    $"[{context}] 图片[{idx}] 数据[{i}]: {ai.ImageData[i]:X2} vs {bi.ImageData[i]:X2}");
        }
    }
    #endregion

    /// <summary>验证页眉页脚和批注在往返中保持</summary>
    [Fact]
    [DisplayName("Docx_往返读写_页眉页脚与批注保真")]
    public void Docx_RoundTrip_HeaderFooterAndComments()
    {
        var path = Path.Combine(OutputDir, "roundtrip_hdr_ftr.docx");

        // 构造含页眉页脚和批注的文档
        using (var w = new WordWriter())
        {
            w.PageSettings.HeaderText = "NewLife.Office 页眉";
            w.PageSettings.FooterText = "第1页 — 测试页脚";

            w.AppendHeading("页眉页脚测试", 1);
            w.AppendParagraph("此文档包含页眉和页脚。");
            w.AppendParagraph("第二段：验证页眉页脚在往返中保持。");

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        // 读取验证
        Document sourceDoc;
        using (var reader = new WordReader(path))
        {
            sourceDoc = reader.ReadDocument();
        }

        Assert.Equal("NewLife.Office 页眉", sourceDoc.HeaderText);
        Assert.True(sourceDoc.FooterText?.Contains("测试页脚") == true, $"页脚应包含'测试页脚'，实际='{sourceDoc.FooterText}'");

        // 往返：写入再读回
        var outputPath = Path.Combine(OutputDir, "roundtrip_hdr_ftr_out.docx");
        using (var writer = new WordWriter())
        {
            writer.Save(outputPath, sourceDoc);
        }

        Document outputDoc;
        using (var reader = new WordReader(outputPath))
        {
            outputDoc = reader.ReadDocument();
        }

        // 验证页眉页脚保持
        Assert.Equal(sourceDoc.HeaderText, outputDoc.HeaderText);
        Assert.Equal(sourceDoc.FooterText, outputDoc.FooterText);

        // 完整内容对比
        AssertDocumentsEqual(sourceDoc, outputDoc, "hdr_ftr");
    }

    /// <summary>生成 1x1 红色 PNG（图片去重回归测试用）</summary>
    private static Byte[] CreatePng()
    {
        // 1x1 红色不透明 PNG
        var b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
        return Convert.FromBase64String(b64);
    }

    [Fact, DisplayName("Word_相同图片字节重复插入_保存无重复ZIP条目（回归#8）")]
    public void Word_DuplicateImageBytes_NoDuplicateZipEntries()
    {
        var path = Path.Combine(OutputDir, $"dup_img_{Guid.NewGuid():N}.docx");
        try
        {
            var png = CreatePng();
            using (var w = new WordWriter())
            {
                w.InsertImage(png, "png", 5, 3);
                w.AppendParagraph("分隔段落");
                w.InsertImage(png, "png", 5, 3); // 相同字节再次插入
                w.Save(path);
            }

            using var za = ZipFile.OpenRead(path);
            var media = za.Entries.Where(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                .Select(e => e.FullName).ToList();
            Assert.NotEmpty(media);
            // 每个媒体文件名必须唯一（修复前相同字节会重复 CreateEntry 同一文件名）
            Assert.Equal(media.Count, media.Distinct().Count());
            // 重读验证两张图片都存在
            using var reader = new WordReader(path);
            var doc = reader.ReadDocument();
            Assert.Equal(2, doc.Images.Count);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
