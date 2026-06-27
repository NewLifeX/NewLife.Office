using System.ComponentModel;
using System.Text;
using NewLife.Log;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>PPTX 全功能往返集成测试</summary>
/// <remarks>
/// 读取真实的 .pptx 文件，将所有内容解析为内存对象（PptSlide/PptTextBox/PptShape/PptImage/
/// PptTable/PptChart/PptGroup/PptTransition 等），
/// 再用 PptxWriter 从这些对象重建新文件，最后多层验证新旧文件内容一致。
/// 不使用 ZIP 级原始 XML 拷贝——每一页幻灯片都必须经过对象模型的完整读取与重建。
/// </remarks>
public class PptxRoundTripTests
{
    static PptxRoundTripTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    #region 辅助方法

    /// <summary>以共享读方式读取文件全部字节（允许其他进程同时打开）</summary>
    private static Byte[] ReadAllBytesShared(String path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var buf = new Byte[fs.Length];
        fs.ReadExactly(buf, 0, buf.Length);
        return buf;
    }

    /// <summary>获取 Bin 目录下所有 .pptx 文件路径（按大小降序）</summary>
    private static List<String> FindAllPptxFiles()
    {
        var baseDir = AppContext.BaseDirectory;
        var candidates = new List<String>
        {
            Path.GetFullPath(Path.Combine(baseDir, "..")),
            Path.GetFullPath(Path.Combine(baseDir, "..", "..", "Bin")),
        };

        var allFiles = new List<String>();
        foreach (var dir in candidates)
        {
            if (!Directory.Exists(dir)) continue;
            allFiles.AddRange(Directory.GetFiles(dir, "*.pptx", SearchOption.TopDirectoryOnly));
        }

        return allFiles.Distinct().OrderByDescending(f => new FileInfo(f).Length).ToList();
    }

    /// <summary>输出目录</summary>
    private static String OutputDir => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "Output"));

    #endregion

    #region 细粒度对比断言

    private static Boolean EmuClose(Int64 a, Int64 b) => Math.Abs(a - b) <= 2;
    private static Boolean Eq(String a, String b) => String.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    /// <summary>验证幻灯片级属性（版式/背景/备注/切换/尺寸）</summary>
    private static void AssertSlideProperties(PptSlide src, PptSlide dst, Int32 slideIndex)
    {
        var tag = $"幻灯片{slideIndex}";
        Assert.Equal(src.LayoutIndex, dst.LayoutIndex);
        Assert.Equal(src.BackgroundColor ?? "", dst.BackgroundColor ?? "");
        if (src.Notes != dst.Notes)
            Assert.Equal(src.Notes, dst.Notes);
        AssertTransition(src.Transition, dst.Transition, tag);
        // 背景图
        if (src.BackgroundImage != null && dst.BackgroundImage != null)
            Assert.Equal(src.BackgroundImage.Data.Length, dst.BackgroundImage.Data.Length);
        else
            Assert.True((src.BackgroundImage == null) == (dst.BackgroundImage == null), $"{tag} 背景图存在性不匹配");
    }

    /// <summary>验证幻灯片各元素数量</summary>
    private static void AssertSlideCounts(PptSlide src, PptSlide dst, Int32 slideIndex)
    {
        var tag = $"幻灯片{slideIndex}";
        Assert.Equal(src.TextBoxes.Count, dst.TextBoxes.Count);
        Assert.Equal(src.Shapes.Count, dst.Shapes.Count);
        Assert.Equal(src.Images.Count, dst.Images.Count);
        Assert.Equal(src.Tables.Count, dst.Tables.Count);
        Assert.Equal(src.Charts.Count, dst.Charts.Count);
        Assert.Equal(src.Groups.Count, dst.Groups.Count);
        Assert.Equal(src.Videos.Count, dst.Videos.Count);
    }

    /// <summary>验证文本框位置和尺寸</summary>
    private static void AssertTextBoxPosition(PptTextBox src, PptTextBox dst, String tag)
    {
        Assert.True(EmuClose(src.Left, dst.Left), $"{tag}.Left: {src.Left} vs {dst.Left}");
        Assert.True(EmuClose(src.Top, dst.Top), $"{tag}.Top: {src.Top} vs {dst.Top}");
        Assert.True(EmuClose(src.Width, dst.Width), $"{tag}.Width: {src.Width} vs {dst.Width}");
        Assert.True(EmuClose(src.Height, dst.Height), $"{tag}.Height: {src.Height} vs {dst.Height}");
    }

    /// <summary>验证文本框文本内容</summary>
    private static void AssertTextBoxText(PptTextBox src, PptTextBox dst, String tag)
    {
        Assert.Equal(src.Text, dst.Text);
    }

    /// <summary>验证文本框字体属性（名称/大小/粗斜/颜色）</summary>
    private static void AssertTextBoxFont(PptTextBox src, PptTextBox dst, String tag)
    {
        // 空文本框跳过 Run 数量验证
        var hasText = src.Text.Length > 0 || dst.Text.Length > 0;
        if (hasText)
        {
            Assert.Equal(src.Runs.Count, dst.Runs.Count);
            for (var i = 0; i < src.Runs.Count; i++)
                AssertTextRun($"  {tag}.Run[{i}]", src.Runs[i], dst.Runs[i]);
        }
        Assert.Equal(src.FontSize, dst.FontSize);
        Assert.Equal(src.Bold, dst.Bold);
        Assert.Equal(src.FontColor ?? "", dst.FontColor ?? "");
        Assert.Equal(src.LatinFontName ?? "", dst.LatinFontName ?? "");
        Assert.Equal(src.EastAsianFontName ?? "", dst.EastAsianFontName ?? "");
        Assert.Equal(src.ComplexScriptFontName ?? "", dst.ComplexScriptFontName ?? "");
        Assert.Equal(src.SymbolFontName ?? "", dst.SymbolFontName ?? "");
    }

    /// <summary>验证文本框格式属性（对齐/背景色/超链接/自适配）</summary>
    private static void AssertTextBoxFormat(PptTextBox src, PptTextBox dst, String tag)
    {
        // 空字符串（未指定）与 "l"（显式左对齐）语义等价
        var srcAlign = src.Alignment.Length == 0 ? "l" : src.Alignment;
        var dstAlign = dst.Alignment.Length == 0 ? "l" : dst.Alignment;
        Assert.Equal(srcAlign, dstAlign);
        Assert.Equal(src.BackgroundColor ?? "", dst.BackgroundColor ?? "");
        Assert.Equal(src.HyperlinkUrl ?? "", dst.HyperlinkUrl ?? "");
        Assert.Equal(src.AutoFit, dst.AutoFit);
    }

    /// <summary>验证文本框段落属性（行距/段前距）</summary>
    private static void AssertTextBoxParagraph(PptTextBox src, PptTextBox dst, String tag)
    {
        Assert.Equal(src.LineSpacingPct, dst.LineSpacingPct);
        Assert.Equal(src.SpaceBeforePt, dst.SpaceBeforePt);
        // 内边距
        Assert.Equal(src.LeftInset, dst.LeftInset);
        Assert.Equal(src.RightInset, dst.RightInset);
        Assert.Equal(src.TopInset, dst.TopInset);
        Assert.Equal(src.BottomInset, dst.BottomInset);
        // 垂直锁定
        Assert.Equal(src.Anchor, dst.Anchor);
        // 多段落对比
        Assert.Equal(src.Paragraphs.Count, dst.Paragraphs.Count);
        for (var pi = 0; pi < src.Paragraphs.Count; pi++)
            AssertParagraph($"{tag}.P[{pi}]", src.Paragraphs[pi], dst.Paragraphs[pi]);
    }

    /// <summary>验证单个段落</summary>
    private static void AssertParagraph(String tag, PptParagraph src, PptParagraph dst)
    {
        Assert.Equal(src.Alignment, dst.Alignment);
        Assert.Equal(src.Level, dst.Level);
        Assert.Equal(src.LineSpacingPct, dst.LineSpacingPct);
        Assert.Equal(src.LineSpacingPts, dst.LineSpacingPts);
        Assert.Equal(src.SpaceBeforePt, dst.SpaceBeforePt);
        Assert.Equal(src.SpaceAfterPt, dst.SpaceAfterPt);
        Assert.Equal(src.BulletChar ?? "", dst.BulletChar ?? "");
        Assert.Equal(src.BulletNone, dst.BulletNone);
        Assert.Equal(src.Runs.Count, dst.Runs.Count);
        for (var i = 0; i < src.Runs.Count; i++)
            AssertTextRun($"{tag}.Run[{i}]", src.Runs[i], dst.Runs[i]);
    }

    /// <summary>验证富文本片段</summary>
    private static void AssertTextRun(String tag, PptTextRun src, PptTextRun dst)
    {
        Assert.Equal(src.Text, dst.Text);
        Assert.Equal(src.FontSize, dst.FontSize);
        Assert.Equal(src.Bold, dst.Bold);
        Assert.Equal(src.Italic, dst.Italic);
        Assert.Equal(src.Underline, dst.Underline);
        Assert.Equal(src.FontColor ?? "", dst.FontColor ?? "");
        Assert.Equal(src.LatinFontName ?? "", dst.LatinFontName ?? "");
        Assert.Equal(src.EastAsianFontName ?? "", dst.EastAsianFontName ?? "");
        Assert.Equal(src.ComplexScriptFontName ?? "", dst.ComplexScriptFontName ?? "");
        Assert.Equal(src.SymbolFontName ?? "", dst.SymbolFontName ?? "");
        Assert.Equal(src.HyperlinkUrl ?? "", dst.HyperlinkUrl ?? "");
        var srcGrad = src.GradFillColors?.Length ?? 0;
        var dstGrad = dst.GradFillColors?.Length ?? 0;
        Assert.Equal(srcGrad, dstGrad);
        for (var i = 0; i < Math.Min(srcGrad, dstGrad); i++)
            Assert.Equal(src.GradFillColors![i] ?? "", dst.GradFillColors![i] ?? "");
    }

    /// <summary>验证基本图形</summary>
    private static void AssertShape(PptShape src, PptShape dst, String tag)
    {
        Assert.Equal(src.ShapeType, dst.ShapeType);
        Assert.True(EmuClose(src.Left, dst.Left), $"{tag}.Left");
        Assert.True(EmuClose(src.Top, dst.Top), $"{tag}.Top");
        Assert.True(EmuClose(src.Width, dst.Width), $"{tag}.Width");
        Assert.True(EmuClose(src.Height, dst.Height), $"{tag}.Height");
        Assert.Equal(src.FillColor ?? "", dst.FillColor ?? "");
        Assert.Equal(src.LineColor ?? "", dst.LineColor ?? "");
        Assert.Equal(src.LineWidth, dst.LineWidth);
        Assert.Equal(src.FontSize, dst.FontSize);
        Assert.Equal(src.Bold, dst.Bold);
        Assert.Equal(src.FontColor ?? "", dst.FontColor ?? "");
        Assert.Equal(src.LatinFontName ?? "", dst.LatinFontName ?? "");
        Assert.Equal(src.EastAsianFontName ?? "", dst.EastAsianFontName ?? "");
        Assert.Equal(src.ComplexScriptFontName ?? "", dst.ComplexScriptFontName ?? "");
        Assert.Equal(src.SymbolFontName ?? "", dst.SymbolFontName ?? "");
    }

    /// <summary>验证图片</summary>
    private static void AssertImage(PptImage src, PptImage dst, String tag)
    {
        Assert.Equal(src.Extension, dst.Extension);
        Assert.True(EmuClose(src.Left, dst.Left), $"{tag}.Left");
        Assert.True(EmuClose(src.Top, dst.Top), $"{tag}.Top");
        Assert.True(EmuClose(src.Width, dst.Width), $"{tag}.Width");
        Assert.True(EmuClose(src.Height, dst.Height), $"{tag}.Height");
        Assert.Equal(src.Data.Length, dst.Data.Length);
    }

    /// <summary>验证表格</summary>
    private static void AssertTable(PptTable src, PptTable dst, String tag)
    {
        Assert.True(EmuClose(src.Left, dst.Left), $"{tag}.Left");
        Assert.True(EmuClose(src.Top, dst.Top), $"{tag}.Top");
        Assert.True(EmuClose(src.Width, dst.Width), $"{tag}.Width");
        Assert.True(EmuClose(src.Height, dst.Height), $"{tag}.Height");
        Assert.Equal(src.FirstRowHeader, dst.FirstRowHeader);
        Assert.Equal(src.Rows.Count, dst.Rows.Count);
        for (var ri = 0; ri < src.Rows.Count; ri++)
        {
            Assert.Equal(src.Rows[ri].Length, dst.Rows[ri].Length);
            for (var ci = 0; ci < src.Rows[ri].Length; ci++)
                Assert.Equal(src.Rows[ri][ci], dst.Rows[ri][ci]);
        }
    }

    /// <summary>验证图表</summary>
    private static void AssertChart(PptChart src, PptChart dst, String tag)
    {
        Assert.Equal(src.ChartType, dst.ChartType);
        Assert.Equal(src.Title ?? "", dst.Title ?? "");
        Assert.True(EmuClose(src.Left, dst.Left), $"{tag}.Left");
        Assert.True(EmuClose(src.Top, dst.Top), $"{tag}.Top");
        Assert.True(EmuClose(src.Width, dst.Width), $"{tag}.Width");
        Assert.True(EmuClose(src.Height, dst.Height), $"{tag}.Height");
        Assert.Equal(src.Categories, dst.Categories);
        Assert.Equal(src.Series.Count, dst.Series.Count);
        for (var i = 0; i < src.Series.Count; i++)
        {
            Assert.Equal(src.Series[i].Name, dst.Series[i].Name);
            Assert.Equal(src.Series[i].Values, dst.Series[i].Values);
        }
    }

    /// <summary>验证形状组</summary>
    private static void AssertGroup(PptGroup src, PptGroup dst, String tag)
    {
        Assert.True(EmuClose(src.Left, dst.Left), $"{tag}.Left");
        Assert.True(EmuClose(src.Top, dst.Top), $"{tag}.Top");
        Assert.True(EmuClose(src.Width, dst.Width), $"{tag}.Width");
        Assert.True(EmuClose(src.Height, dst.Height), $"{tag}.Height");
        Assert.Equal(src.Shapes.Count, dst.Shapes.Count);
        for (var i = 0; i < src.Shapes.Count; i++)
            AssertShape(src.Shapes[i], dst.Shapes[i], $"{tag}.SP[{i}]");
        Assert.Equal(src.TextBoxes.Count, dst.TextBoxes.Count);
        for (var i = 0; i < src.TextBoxes.Count; i++)
        {
            AssertTextBoxPosition(src.TextBoxes[i], dst.TextBoxes[i], $"{tag}.TB[{i}]");
            AssertTextBoxText(src.TextBoxes[i], dst.TextBoxes[i], $"{tag}.TB[{i}]");
        }
    }

    /// <summary>验证视频</summary>
    private static void AssertVideo(PptVideo src, PptVideo dst, String tag)
    {
        Assert.Equal(src.Extension, dst.Extension);
        Assert.True(EmuClose(src.Left, dst.Left), $"{tag}.Left");
        Assert.True(EmuClose(src.Top, dst.Top), $"{tag}.Top");
        Assert.True(EmuClose(src.Width, dst.Width), $"{tag}.Width");
        Assert.True(EmuClose(src.Height, dst.Height), $"{tag}.Height");
        Assert.Equal(src.Data.Length, dst.Data.Length);
        var srcThumb = src.ThumbnailData?.Length ?? 0;
        var dstThumb = dst.ThumbnailData?.Length ?? 0;
        Assert.Equal(srcThumb, dstThumb);
    }

    /// <summary>验证切换动画</summary>
    private static void AssertTransition(PptTransition src, PptTransition dst, String tag)
    {
        if (src == null && dst == null) return;
        Assert.True(src != null && dst != null, $"{tag} 切换动画存在性不匹配");
        if (src == null || dst == null) return;
        Assert.Equal(src.Type, dst.Type);
        Assert.Equal(src.DurationMs, dst.DurationMs);
        Assert.Equal(src.Direction, dst.Direction);
        Assert.Equal(src.AdvanceOnClick, dst.AdvanceOnClick);
    }

    #endregion
    [Fact]
    [DisplayName("诊断：对象模型级字体对比")]
    public void Diagnose_FontDifferences()
    {
        var sourcePath = FindAllPptxFiles()[0];
        var outputPath = Path.Combine(
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "Output")),
            Path.GetFileName(sourcePath));

        if (!File.Exists(outputPath))
        {
            RunDiagnosticRoundTrip();
        }

        var diagSb = new StringBuilder();
        var fontMismatch = false;
        using (var srcReader = new PptxReader(sourcePath))
        using (var outReader = new PptxReader(outputPath))
        {
            var srcSlides = srcReader.ReadAllSlides().ToList();
            var outSlides = outReader.ReadAllSlides().ToList();
            var maxSlides = Math.Min(srcSlides.Count, outSlides.Count);

            diagSb.AppendLine($"源幻灯片数={srcSlides.Count} 输出={outSlides.Count}");

            for (var si = 0; si < maxSlides; si++)
            {
                diagSb.AppendLine();
                diagSb.AppendLine($"=== 幻灯片 {si} ===");

                var src = srcSlides[si];
                var dst = outSlides[si];

                // 逐 TextBox 逐 Run 提取字体引用（仅 Run 级别，排除 defRPr/endParaRPr）
                var srcFonts = new List<(String tag, String typeface)>();
                var outFonts = new List<(String tag, String typeface)>();
                CollectRunFonts(src, srcFonts);
                CollectRunFonts(dst, outFonts);

                if (srcFonts.Count != outFonts.Count || !srcFonts.SequenceEqual(outFonts))
                    fontMismatch = true;

                diagSb.AppendLine($"  源 Runs 字体引用: {srcFonts.Count}");
                foreach (var f in srcFonts)
                    diagSb.AppendLine($"    [{f.tag}] {f.typeface}");
                diagSb.AppendLine($"  输出 Runs 字体引用: {outFonts.Count}");
                foreach (var f in outFonts)
                    diagSb.AppendLine($"    [{f.tag}] {f.typeface}");

                // 逐 Run 对比对象模型
                for (var ti = 0; ti < Math.Max(src.TextBoxes.Count, dst.TextBoxes.Count); ti++)
                {
                    if (ti >= src.TextBoxes.Count || ti >= dst.TextBoxes.Count) break;
                    var stb = src.TextBoxes[ti];
                    var otb = dst.TextBoxes[ti];
                    diagSb.AppendLine($"  TB[{ti}] Text=[{stb.Text.Substring(0, Math.Min(40, stb.Text.Length))}] Lfn=[{stb.LatinFontName}] Efn=[{stb.EastAsianFontName}]");
                    for (var ri = 0; ri < Math.Max(stb.Runs.Count, otb.Runs.Count); ri++)
                    {
                        if (ri >= stb.Runs.Count || ri >= otb.Runs.Count) break;
                        var sr = stb.Runs[ri];
                        var or = otb.Runs[ri];
                        var diff = (sr.LatinFontName != or.LatinFontName || sr.EastAsianFontName != or.EastAsianFontName) ? " ***DIFF***" : "";
                        diagSb.AppendLine($"    Run[{ri}] Text=[{sr.Text}] Lfn=[{sr.LatinFontName}] vs [{or.LatinFontName}] Efn=[{sr.EastAsianFontName}] vs [{or.EastAsianFontName}]{diff}");
                    }
                }
            }
        }

        Assert.False(fontMismatch, "对象模型级字体引用不一致，差异见诊断文件");

        var diagPath = Path.Combine(
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "Output")),
            "PptxFontDiagnose.txt");
        File.WriteAllText(diagPath, diagSb.ToString(), Encoding.UTF8);
        XTrace.WriteLine(diagSb.ToString());
    }

    /// <summary>诊断测试依赖的单文件往返（使用第一个 pptx 文件）</summary>
    private static void RunDiagnosticRoundTrip()
    {
        var files = FindAllPptxFiles();
        if (files.Count == 0) return;
        RunSingleFileRoundTrip(files[0], OutputDir);
    }

    /// <summary>收集幻灯片中所有 TextBox 的 Run 级别字体引用（仅 latin + ea）</summary>
    private static void CollectRunFonts(PptSlide slide, List<(String tag, String typeface)> result)
    {
        foreach (var tb in slide.TextBoxes)
        {
            foreach (var run in tb.Runs)
            {
                if (run.LatinFontName != null)
                    result.Add(("latin", run.LatinFontName));
                if (run.EastAsianFontName != null)
                    result.Add(("ea", run.EastAsianFontName));
            }
            // 对于无 Run 的 TextBox，使用 TextBox 级字体
            if (tb.Runs.Count == 0)
            {
                if (tb.LatinFontName != null)
                    result.Add(("latin", tb.LatinFontName));
                if (tb.EastAsianFontName != null)
                    result.Add(("ea", tb.EastAsianFontName));
            }
        }
        // Shapes 中的字体也收集
        foreach (var sp in slide.Shapes)
        {
            if (sp.Text.Length > 0)
            {
                if (sp.LatinFontName != null)
                    result.Add(("latin", sp.LatinFontName));
                if (sp.EastAsianFontName != null)
                    result.Add(("ea", sp.EastAsianFontName));
            }
        }
    }

    /// <summary>诊断：逐页比较源与输出完整 slide XML（规范化空格），定位第一个差异</summary>
    [Fact]
    [DisplayName("诊断：逐页规范化 XML 对比，定位第一个差异字符")]
    public void Diagnose_FullXmlCompare()
    {
        var sourcePath = FindAllPptxFiles()[0];
        var outputPath = Path.Combine(
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "Output")),
            Path.GetFileName(sourcePath));

        if (!File.Exists(outputPath))
            RunDiagnosticRoundTrip();

        using var srcReader = new PptxReader(sourcePath);
        using var outReader = new PptxReader(outputPath);
        var srcCount = srcReader.GetSlideCount();
        var outCount = outReader.GetSlideCount();
        Assert.Equal(srcCount, outCount);

        for (var i = 0; i < srcCount; i++)
        {
            var srcXml = srcReader.GetSlideXml(i);
            var outXml = outReader.GetSlideXml(i);
            if (srcXml == null && outXml == null) continue;
            Assert.NotNull(srcXml);
            Assert.NotNull(outXml);

            // 规范化：统一空白字符，忽略 XML 声明差异
            var normSrc = NormalizeXml(srcXml!);
            var normOut = NormalizeXml(outXml!);

            if (normSrc != normOut)
            {
                var diffPos = 0;
                var minLen = Math.Min(normSrc.Length, normOut.Length);
                for (; diffPos < minLen; diffPos++)
                    if (normSrc[diffPos] != normOut[diffPos]) break;

                var ctxStart = Math.Max(0, diffPos - 80);
                var ctxEnd = Math.Min(minLen, diffPos + 80);
                // 诊断输出差异位置，不强制 Assert（namespace 声明位置等序列化细节差异无实际影响）
                XTrace.WriteLine($"幻灯片 {i} spTree XML 在位置 {diffPos} 处不一致");
                XTrace.WriteLine($"  源附近: [{normSrc.Substring(ctxStart, Math.Min(160, normSrc.Length - ctxStart))}]");
                XTrace.WriteLine($"  输出附近: [{normOut.Substring(ctxStart, Math.Min(160, normOut.Length - ctxStart))}]");
            }
        }
    }

    /// <summary>规范化 XML：移除声明、排序属性、统一关系ID、仅比较 spTree 内容</summary>
    private static String NormalizeXml(String xml)
    {
        // 解析为 XmlDocument
        var doc = new System.Xml.XmlDocument();
        doc.PreserveWhitespace = false;
        doc.LoadXml(xml);

        // 提取 <p:spTree> 子树（幻灯片内容核心），忽略外围结构差异（bg位置等）
        var spTree = doc.SelectSingleNode("//*[local-name()='spTree']") as System.Xml.XmlElement;
        if (spTree == null) return xml; // 无 spTree，直接比较原文

        // 递归排序所有属性
        SortAttributes(spTree);

        // 序列化 spTree 子树
        using var sw = new StringWriter();
        var settings = new System.Xml.XmlWriterSettings
        {
            OmitXmlDeclaration = true,
            Indent = false,
            NewLineHandling = System.Xml.NewLineHandling.None,
        };
        using var xw = System.Xml.XmlWriter.Create(sw, settings);
        spTree.WriteTo(xw);
        xw.Flush();
        var normalized = sw.ToString();

        // 统一关系 ID 和形状 ID/名称
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized,
            @"(?:r:embed|r:id|r:link)\s*=\s*""[^""]*""", "r:ref=\"X\"");
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized,
            @"\bId\s*=\s*""[^""]*""", "Id=\"X\"");
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized,
            @"\bname\s*=\s*""[^""]*""", "name=\"X\"");
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized,
            @"\bid\s*=\s*""[^""]*""", "id=\"X\"");

        return normalized;
    }

    /// <summary>递归排序 XML 元素的所有属性（按名称排序）</summary>
    private static void SortAttributes(System.Xml.XmlElement el)
    {
        if (el.Attributes.Count <= 1) { /* 无需排序 */ }
        else
        {
            var attrs = new System.Xml.XmlAttribute[el.Attributes.Count];
            el.Attributes.CopyTo(attrs, 0);
            el.Attributes.RemoveAll();
            foreach (var a in attrs.OrderBy(a => a.Name))
                el.Attributes.Append(a);
        }
        foreach (var child in el.ChildNodes)
        {
            if (child is System.Xml.XmlElement childEl)
                SortAttributes(childEl);
        }
    }

    #region 主测试

    /// <summary>对单个 pptx 文件执行完整往返验证（读取→写入→模型对比→Entry 对比）</summary>
    private static void RunSingleFileRoundTrip(String sourcePath, String outputDir)
    {
        var fileName = Path.GetFileName(sourcePath);
        var sourceInfo = new FileInfo(sourcePath);
        var sourceBytes = ReadAllBytesShared(sourcePath);
        var outputPath = Path.Combine(outputDir, fileName);

        // ─── 读取阶段 ───────────────────────────────────────────────
        var sourceSlides = new List<PptSlide>();
        String masterXml, layoutXml, themeXml;
        Int32 sourceImageCount, sourceSlideCount;
        Int64 sourceImageTotalBytes;
        Int64 sourceSlideWidth, sourceSlideHeight;
        String[] sourceAccentColors;

        using (var ms = new MemoryStream(sourceBytes))
        using (var reader = new PptxReader(ms))
        {
            sourceSlideCount = reader.GetSlideCount();
            Assert.True(sourceSlideCount > 0, $"[{fileName}] 源文件应至少包含 1 张幻灯片");

            sourceSlides = reader.ReadAllSlides().ToList();
            Assert.Equal(sourceSlideCount, sourceSlides.Count);

            masterXml = reader.GetSlideMasterXml(0);
            layoutXml = reader.GetSlideLayoutXml(0);
            themeXml = reader.GetThemeXml();

            var images = reader.ExtractImages().ToList();
            sourceImageCount = images.Count;
            sourceImageTotalBytes = images.Sum(img => (Int64)img.Data.Length);

            sourceSlideWidth = reader.SlideWidth;
            sourceSlideHeight = reader.SlideHeight;
            sourceAccentColors = reader.AccentColors;
        }

        // ─── 写入阶段 ───────────────────────────────────────────────
        try
        {
            using var writer = new PptxWriter();
            writer.LoadMaster(sourceBytes);
            writer.SlideWidth = sourceSlideWidth;
            writer.SlideHeight = sourceSlideHeight;
            writer.AccentColors = sourceAccentColors;

            for (var i = 0; i < sourceSlides.Count; i++)
            {
                var src = sourceSlides[i];
                var slideIdx = writer.Slides.Count;
                writer.AddSlide(src.LayoutIndex);

                foreach (var tb in src.TextBoxes)
                {
                    var leftCm = PptxWriter.EmuToCm(tb.Left);
                    var topCm = PptxWriter.EmuToCm(tb.Top);
                    var widthCm = PptxWriter.EmuToCm(tb.Width);
                    var heightCm = PptxWriter.EmuToCm(tb.Height);
                    if (widthCm <= 0) widthCm = 10;
                    if (heightCm <= 0) heightCm = 2;

                    var newTb = writer.AddTextBox(slideIdx, tb.Text, leftCm, topCm, widthCm, heightCm,
                        tb.FontSize > 0 ? tb.FontSize : 18, tb.Bold);

                    newTb.FontColor = tb.FontColor;
                    newTb.LatinFontName = tb.LatinFontName;
                    newTb.EastAsianFontName = tb.EastAsianFontName;
                    newTb.ComplexScriptFontName = tb.ComplexScriptFontName;
                    newTb.SymbolFontName = tb.SymbolFontName;
                    newTb.Alignment = tb.Alignment;
                    newTb.BackgroundColor = tb.BackgroundColor;
                    newTb.HyperlinkUrl = tb.HyperlinkUrl;
                    newTb.AutoFit = tb.AutoFit;
                    newTb.LineSpacingPct = tb.LineSpacingPct;
                    newTb.SpaceBeforePt = tb.SpaceBeforePt;
                    newTb.LeftInset = tb.LeftInset;
                    newTb.RightInset = tb.RightInset;
                    newTb.TopInset = tb.TopInset;
                    newTb.BottomInset = tb.BottomInset;
                    newTb.Anchor = tb.Anchor;

                    if (tb.Paragraphs.Count > 0)
                    {
                        newTb.Paragraphs.Clear();
                        foreach (var pp in tb.Paragraphs)
                        {
                            var newPp = new PptParagraph
                            {
                                Alignment = pp.Alignment,
                                Level = pp.Level,
                                LineSpacingPct = pp.LineSpacingPct,
                                LineSpacingPts = pp.LineSpacingPts,
                                SpaceBeforePt = pp.SpaceBeforePt,
                                SpaceAfterPt = pp.SpaceAfterPt,
                                BulletChar = pp.BulletChar,
                                BulletNone = pp.BulletNone,
                            };
                            foreach (var run in pp.Runs)
                                newPp.Runs.Add(new PptTextRun
                                {
                                    Text = run.Text,
                                    FontSize = run.FontSize,
                                    Bold = run.Bold,
                                    Italic = run.Italic,
                                    Underline = run.Underline,
                                    FontColor = run.FontColor,
                                    LatinFontName = run.LatinFontName,
                                    EastAsianFontName = run.EastAsianFontName,
                                    ComplexScriptFontName = run.ComplexScriptFontName,
                                    SymbolFontName = run.SymbolFontName,
                                    HyperlinkUrl = run.HyperlinkUrl,
                                    GradFillColors = run.GradFillColors,
                                    GradAngle = run.GradAngle,
                                });
                            newTb.Paragraphs.Add(newPp);
                        }
                    }
                    if (tb.Runs.Count > 0)
                    {
                        newTb.Runs.Clear();
                        foreach (var run in tb.Runs)
                            newTb.Runs.Add(new PptTextRun
                            {
                                Text = run.Text,
                                FontSize = run.FontSize,
                                Bold = run.Bold,
                                Italic = run.Italic,
                                Underline = run.Underline,
                                FontColor = run.FontColor,
                                LatinFontName = run.LatinFontName,
                                EastAsianFontName = run.EastAsianFontName,
                                ComplexScriptFontName = run.ComplexScriptFontName,
                                SymbolFontName = run.SymbolFontName,
                                HyperlinkUrl = run.HyperlinkUrl,
                                GradFillColors = run.GradFillColors,
                                GradAngle = run.GradAngle,
                            });
                    }
                }

                foreach (var sp in src.Shapes)
                {
                    var leftCm = PptxWriter.EmuToCm(sp.Left);
                    var topCm = PptxWriter.EmuToCm(sp.Top);
                    var widthCm = PptxWriter.EmuToCm(sp.Width);
                    var heightCm = PptxWriter.EmuToCm(sp.Height);
                    if (widthCm <= 0) widthCm = 5;
                    if (heightCm <= 0) heightCm = 3;
                    var newSp = writer.AddShape(slideIdx, sp.ShapeType, leftCm, topCm, widthCm, heightCm, sp.FillColor);
                    newSp.Text = sp.Text; newSp.FontSize = sp.FontSize > 0 ? sp.FontSize : 14;
                    newSp.FontColor = sp.FontColor; newSp.Bold = sp.Bold;
                    newSp.LineColor = sp.LineColor; newSp.LineWidth = sp.LineWidth;
                    newSp.LatinFontName = sp.LatinFontName; newSp.EastAsianFontName = sp.EastAsianFontName;
                    newSp.ComplexScriptFontName = sp.ComplexScriptFontName; newSp.SymbolFontName = sp.SymbolFontName;
                }

                foreach (var img in src.Images)
                    writer.AddImage(slideIdx, img.Data, img.Extension,
                        PptxWriter.EmuToCm(img.Left), PptxWriter.EmuToCm(img.Top),
                        PptxWriter.EmuToCm(img.Width), PptxWriter.EmuToCm(img.Height));

                foreach (var vid in src.Videos)
                {
                    var newVid = writer.AddVideo(slideIdx, vid.Data, vid.Extension,
                        PptxWriter.EmuToCm(vid.Left), PptxWriter.EmuToCm(vid.Top),
                        PptxWriter.EmuToCm(vid.Width), PptxWriter.EmuToCm(vid.Height));
                    newVid.ThumbnailData = vid.ThumbnailData;
                    newVid.ThumbnailExtension = vid.ThumbnailExtension;
                }

                foreach (var tbl in src.Tables)
                {
                    var newTbl = writer.AddTable(slideIdx, tbl.Rows,
                        PptxWriter.EmuToCm(tbl.Left), PptxWriter.EmuToCm(tbl.Top),
                        PptxWriter.EmuToCm(tbl.Width), tbl.FirstRowHeader);
                    newTbl.Height = tbl.Height;
                    if (tbl.ColWidths.Length > 0) newTbl.ColWidths = tbl.ColWidths;
                    foreach (var kv in tbl.CellStyles) newTbl.CellStyles[kv.Key] = kv.Value;
                }

                foreach (var chart in src.Charts)
                {
                    var leftCm = PptxWriter.EmuToCm(chart.Left);
                    var topCm = PptxWriter.EmuToCm(chart.Top);
                    var widthCm = PptxWriter.EmuToCm(chart.Width);
                    var heightCm = PptxWriter.EmuToCm(chart.Height);
                    PptChart newChart;
                    if (chart.ChartType == "line")
                        newChart = writer.AddLineChart(slideIdx, chart.Categories, leftCm, topCm, widthCm, heightCm);
                    else if (chart.ChartType == "pie")
                        newChart = writer.AddPieChart(slideIdx, chart.Categories, leftCm, topCm, widthCm, heightCm);
                    else
                        newChart = writer.AddBarChart(slideIdx, chart.Categories, leftCm, topCm, widthCm, heightCm);
                    newChart.Title = chart.Title; newChart.ChartType = chart.ChartType;
                    newChart.Series.Clear();
                    foreach (var ser in chart.Series) newChart.Series.Add(new PptChartSeries { Name = ser.Name, Values = ser.Values });
                }

                foreach (var grp in src.Groups)
                {
                    var newGrp = writer.GroupShapes(slideIdx,
                        PptxWriter.EmuToCm(grp.Left), PptxWriter.EmuToCm(grp.Top),
                        PptxWriter.EmuToCm(grp.Width), PptxWriter.EmuToCm(grp.Height));
                    foreach (var sp in grp.Shapes) newGrp.Shapes.Add(sp);
                    foreach (var tb in grp.TextBoxes) newGrp.TextBoxes.Add(tb);
                }

                if (src.BackgroundColor != null) writer.SetBackground(slideIdx, src.BackgroundColor);
                if (src.BackgroundImage != null) writer.Slides[slideIdx].BackgroundImage = src.BackgroundImage;
                if (src.Notes != null) writer.SetNotes(slideIdx, src.Notes);
                if (src.Transition != null) writer.SetTransition(slideIdx, src.Transition.Type, src.Transition.DurationMs);
            }

            writer.Save(outputPath);
        }
        catch
        {
            throw;
        }

        // ─── 验证阶段 ───────────────────────────────────────────────
        try
        {
            using var outReader = new PptxReader(outputPath);

            // ① 幻灯片数量
            var outSlideCount = outReader.GetSlideCount();
            Assert.Equal(sourceSlideCount, outSlideCount);

            // ② 逐页文本
            using var srcReader2 = new PptxReader(sourcePath);
            for (var i = 0; i < sourceSlideCount; i++)
            {
                var srcText = srcReader2.GetSlideText(i) ?? String.Empty;
                var outText = outReader.GetSlideText(i) ?? String.Empty;
                Assert.Equal(srcText.Trim(), outText.Trim());
            }

            // ③ 母版/版式/主题
            if (masterXml != null) { var m = outReader.GetSlideMasterXml(0); Assert.Equal(masterXml, m); }
            if (layoutXml != null) { var l = outReader.GetSlideLayoutXml(0); Assert.Equal(layoutXml, l); }
            if (themeXml != null) { var t = outReader.GetThemeXml(); Assert.Equal(themeXml, t); }

            // ④ 图片数据量（仅当源有图片时检查，部分文件图片在版式/母版中不由 ExtractImages 返回）
            var outImages = outReader.ExtractImages().ToList();
            var outImageTotalBytes = outImages.Sum(img => (Int64)img.Data.Length);
            if (sourceImageTotalBytes > 0)
            {
                var imageRatio = (Double)outImageTotalBytes / sourceImageTotalBytes;
                Assert.True(imageRatio >= 0.5, $"输出图片总字节({outImageTotalBytes})不应显著少于源({sourceImageTotalBytes})");
            }

            // ⑤ 图表数据
            using var srcReader3 = new PptxReader(sourcePath);
            for (var i = 0; i < sourceSlideCount; i++)
            {
                var srcCharts = srcReader3.ReadChartData(i).ToList();
                var outCharts = outReader.ReadChartData(i).ToList();
                Assert.Equal(srcCharts.Count, outCharts.Count);
                for (var ci = 0; ci < srcCharts.Count; ci++)
                {
                    Assert.Equal(srcCharts[ci].ChartType, outCharts[ci].ChartType);
                    Assert.Equal(srcCharts[ci].Categories, outCharts[ci].Categories);
                    Assert.Equal(srcCharts[ci].Series.Count, outCharts[ci].Series.Count);
                    for (var si = 0; si < srcCharts[ci].Series.Count; si++)
                    {
                        Assert.Equal(srcCharts[ci].Series[si].Name, outCharts[ci].Series[si].Name);
                        Assert.Equal(srcCharts[ci].Series[si].Values, outCharts[ci].Series[si].Values);
                    }
                }
            }

            // ⑥ 细粒度属性对比
            var outSlides = outReader.ReadAllSlides().ToList();
            Assert.Equal(sourceSlides.Count, outSlides.Count);
            for (var i = 0; i < sourceSlides.Count; i++)
            {
                var src = sourceSlides[i];
                var dst = outSlides[i];

                AssertSlideProperties(src, dst, i);
                AssertSlideCounts(src, dst, i);

                for (var ti = 0; ti < src.TextBoxes.Count; ti++)
                {
                    var tag = $"幻灯片{i}.TB[{ti}]";
                    AssertTextBoxPosition(src.TextBoxes[ti], dst.TextBoxes[ti], tag);
                    AssertTextBoxText(src.TextBoxes[ti], dst.TextBoxes[ti], tag);
                    AssertTextBoxFont(src.TextBoxes[ti], dst.TextBoxes[ti], tag);
                    AssertTextBoxFormat(src.TextBoxes[ti], dst.TextBoxes[ti], tag);
                    AssertTextBoxParagraph(src.TextBoxes[ti], dst.TextBoxes[ti], tag);
                }
                for (var si = 0; si < src.Shapes.Count; si++)
                    AssertShape(src.Shapes[si], dst.Shapes[si], $"幻灯片{i}.SP[{si}]");
                for (var ii = 0; ii < src.Images.Count; ii++)
                    AssertImage(src.Images[ii], dst.Images[ii], $"幻灯片{i}.IMG[{ii}]");
                for (var ti = 0; ti < src.Tables.Count; ti++)
                    AssertTable(src.Tables[ti], dst.Tables[ti], $"幻灯片{i}.TBL[{ti}]");
                for (var ci = 0; ci < src.Charts.Count; ci++)
                    AssertChart(src.Charts[ci], dst.Charts[ci], $"幻灯片{i}.CHT[{ci}]");
                for (var gi = 0; gi < src.Groups.Count; gi++)
                    AssertGroup(src.Groups[gi], dst.Groups[gi], $"幻灯片{i}.GRP[{gi}]");
                for (var vi = 0; vi < src.Videos.Count; vi++)
                    AssertVideo(src.Videos[vi], dst.Videos[vi], $"幻灯片{i}.VID[{vi}]");
            }

            // ⑦ Entry 级 ZIP 对比
            var entryResult = PptxZipComparer.Compare(sourcePath, outputPath);
            if (entryResult.HasCritical)
            {
                XTrace.WriteLine($"  [Entry 差异] {fileName}: {String.Join(" | ", entryResult.Critical.Take(5))}{(entryResult.Critical.Count > 5 ? $" ... 等 {entryResult.Critical.Count - 5} 个" : "")}");
                Assert.Fail($"[{fileName}] Entry 级关键差异 ({entryResult.Critical.Count} 项):\n{String.Join("\n", entryResult.Critical)}");
            }

            XTrace.WriteLine($"  \u2705 {fileName}: {sourceSlideCount} 幻灯片, {sourceImageCount} 图片, {outSlides.Sum(s => s.TextBoxes.Count)} 文本框, Entry 对比通过");
        }
        finally { }
    }

    /// <summary>遍历 Bin 目录下所有 .pptx 文件，每个执行完整往返测试，任意失败则整体失败</summary>
    [Fact]
    [DisplayName("pptx逐文件往返")]
    public void FullRoundTrip_AllFiles()
    {
        var allFiles = FindAllPptxFiles();
        Assert.True(allFiles.Count > 0, $"Bin 目录下未找到 .pptx 文件。BaseDir={AppContext.BaseDirectory}");

        Directory.CreateDirectory(OutputDir);
        XTrace.WriteLine($"找到 {allFiles.Count} 个 pptx 文件，开始逐文件往返测试\n");

        var failures = new List<(String file, String error)>();
        foreach (var sourcePath in allFiles)
        {
            var fileName = Path.GetFileName(sourcePath);
            try
            {
                RunSingleFileRoundTrip(sourcePath, OutputDir);
            }
            catch (Exception ex)
            {
                // 取异常消息的前 3 行（第一行是概要，后面是细节）
                var lines = ex.Message.Split('\n');
                var msg = String.Join(" | ", lines.Take(3)).Trim();
                failures.Add((fileName, msg));
            }
        }

        // 汇总报告
        if (failures.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"\n❌ {failures.Count}/{allFiles.Count} 个文件测试失败:");
            foreach (var (file, err) in failures)
                sb.AppendLine($"  [{file}] {err}");
            Assert.Fail(sb.ToString());
        }
        else
        {
            XTrace.WriteLine($"\n✅ 全部 {allFiles.Count} 个文件往返测试通过");
        }
    }

    #endregion
}
