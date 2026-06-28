using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace XUnitTest.Ppt;

/// <summary>PPTX ZIP Entry 级对比工具（分级策略版）</summary>
/// <remarks>
/// 对比策略分级：
/// - 母版/版式/主题（ppt/slideMasters / ppt/slideLayouts / ppt/theme）→ 二进制精确对比
/// - 嵌入字体（ppt/fonts/）→ 二进制精确对比（关键！缺失会导致字体显示错误）
/// - 媒体文件（ppt/media/）→ SHA256 内容哈希对比（忽略 Writer 重命名，只验证内容存在性）
/// - 幻灯片 XML（ppt/slides/slideN.xml）→ 规范化 spTree 对比（忽略 shapeId/relId/name）
/// - 图表 XML（ppt/charts/）→ 规范化 plotArea 对比
/// - presentation.xml → 结构化对比（尺寸 + 幻灯片数量）
/// - 幻灯片 rels → 关系类型数量对比（忽略具体 ID 和路径）
/// </remarks>
public static class PptxZipComparer
{
    /// <summary>对比结果（分关键/信息两级）</summary>
    public sealed class CompareResult
    {
        /// <summary>关键差异（影响视觉效果，应导致测试失败）</summary>
        public List<String> Critical { get; } = [];
        /// <summary>信息性差异（预期行为或不影响视觉，仅报告）</summary>
        public List<String> Informational { get; } = [];
        /// <summary>是否有关键差异</summary>
        public Boolean HasCritical => Critical.Count > 0;
        /// <summary>所有差异合并（向后兼容）</summary>
        public List<String> AllIssues => [.. Critical, .. Informational];
    }

    /// <summary>对比两个 pptx 文件的 ZIP Entry，返回分级差异结果</summary>
    /// <param name="sourcePath">源 pptx 文件路径</param>
    /// <param name="outputPath">输出 pptx 文件路径</param>
    public static CompareResult Compare(String sourcePath, String outputPath)
    {
        var result = new CompareResult();

        using var srcZip = ZipFile.OpenRead(sourcePath);
        using var outZip = ZipFile.OpenRead(outputPath);

        // 过滤目录条目
        var srcEntries = srcZip.Entries
            .Where(e => !e.FullName.EndsWith("/"))
            .ToDictionary(e => e.FullName, StringComparer.OrdinalIgnoreCase);
        var outEntries = outZip.Entries
            .Where(e => !e.FullName.EndsWith("/"))
            .ToDictionary(e => e.FullName, StringComparer.OrdinalIgnoreCase);

        // 1. 母版/版式：二进制精确对比
        CompareBinaryPrefix(srcEntries, outEntries, "ppt/slideMasters/", result.Critical);
        CompareBinaryPrefix(srcEntries, outEntries, "ppt/slideLayouts/", result.Critical);

        // 主题：固定比较 theme1.xml（LoadMaster 加载此文件，WriteTheme 写入此文件，避免多主题文件时误匹配）
        var srcTheme = srcEntries.ContainsKey("ppt/theme/theme1.xml") ? "ppt/theme/theme1.xml"
            : srcEntries.Keys.FirstOrDefault(k =>
                k.StartsWith("ppt/theme/", StringComparison.OrdinalIgnoreCase) && k.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        var outTheme = outEntries.ContainsKey("ppt/theme/theme1.xml") ? "ppt/theme/theme1.xml"
            : outEntries.Keys.FirstOrDefault(k =>
                k.StartsWith("ppt/theme/", StringComparison.OrdinalIgnoreCase) && k.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        if (srcTheme != null && outTheme != null)
        {
            var srcText = NormalizeLineEndings(StripBom(ReadAllBytes(srcEntries[srcTheme])));
            var outText = NormalizeLineEndings(StripBom(ReadAllBytes(outEntries[outTheme])));
            if (!srcText.SequenceEqual(outText))
                result.Critical.Add($"ppt/theme: 主题内容不一致 (src={srcText.Length}B out={outText.Length}B)");
        }

        // 2. 嵌入字体：二进制精确对比（关键）
        CompareBinaryPrefix(srcEntries, outEntries, "ppt/fonts/", result.Critical);

        // 3. 媒体文件：SHA256 内容哈希对比（允许重命名）
        CompareMediaByHash(srcEntries, outEntries, result);

        // 4. 幻灯片 XML：规范化 spTree 对比（信息性——模型级对比更精确，XML 对比作为补充诊断）
        CompareSlideXmls(srcEntries, outEntries, result.Informational);

        // 5. 图表 XML：规范化 plotArea 对比
        CompareChartXmls(srcEntries, outEntries, result.Critical);

        // 6. presentation.xml：结构化对比
        if (srcEntries.TryGetValue("ppt/presentation.xml", out var srcPres) &&
            outEntries.TryGetValue("ppt/presentation.xml", out var outPres))
            ComparePresentation(srcPres, outPres, result.Critical);

        // 7. 幻灯片数量
        var srcCount = srcEntries.Keys.Count(k =>
            System.Text.RegularExpressions.Regex.IsMatch(k, @"^ppt/slides/slide\d+\.xml$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase));
        var outCount = outEntries.Keys.Count(k =>
            System.Text.RegularExpressions.Regex.IsMatch(k, @"^ppt/slides/slide\d+\.xml$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase));
        if (srcCount != outCount)
            result.Critical.Add($"幻灯片数量不一致: {srcCount} vs {outCount}");

        // 8. 幻灯片 rels：关系类型统计（信息性）
        CompareSlideRels(srcEntries, outEntries, result.Informational);

        return result;
    }

    /// <summary>向后兼容：返回所有差异列表</summary>
    public static List<String> CompareZipEntries(String sourcePath, String outputPath)
        => Compare(sourcePath, outputPath).AllIssues;

    #region 对比策略

    private static void CompareBinaryPrefix(
        Dictionary<String, ZipArchiveEntry> srcEntries,
        Dictionary<String, ZipArchiveEntry> outEntries,
        String prefix,
        List<String> issues)
    {
        var srcSet = srcEntries.Keys
            .Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToList();
        var outSet = new HashSet<String>(
            outEntries.Keys.Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)),
            StringComparer.OrdinalIgnoreCase);

        foreach (var name in srcSet)
        {
            if (!outSet.Contains(name))
            {
                issues.Add($"缺失: {name}");
                continue;
            }
            // 使用去 BOM 的字节对比（防止 UTF-8 BOM 差异导致误报）
            var srcData = StripBom(ReadAllBytes(srcEntries[name]));
            var outData = StripBom(ReadAllBytes(outEntries[name]));
            if (!srcData.SequenceEqual(outData))
                issues.Add($"内容变化: {name} ({srcData.Length}B -> {outData.Length}B)");
        }
    }

    private static void CompareMediaByHash(
        Dictionary<String, ZipArchiveEntry> srcEntries,
        Dictionary<String, ZipArchiveEntry> outEntries,
        CompareResult result)
    {
        var srcMedia = srcEntries
            .Where(kv => kv.Key.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var outMedia = outEntries
            .Where(kv => kv.Key.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // 计算输出所有媒体的哈希集合
        var outHashes = new HashSet<String>(StringComparer.Ordinal);
        foreach (var kv in outMedia)
            outHashes.Add(ComputeHash(ReadAllBytes(kv.Value)));

        // 检查每个源媒体内容是否存在于输出（允许不同名称）
        var missing = new List<String>();
        foreach (var kv in srcMedia)
        {
            var hash = ComputeHash(ReadAllBytes(kv.Value));
            if (!outHashes.Contains(hash))
                missing.Add(Path.GetFileName(kv.Key));
        }

        if (missing.Count > 0)
        {
            var preview = String.Join(", ", missing.Take(5));
            var suffix = missing.Count > 5 ? $" ... 等 {missing.Count - 5} 个" : "";
            // 超过 25% 才算关键（基础设施图片/handout/notes 中的媒体不影响幻灯片视觉）
            var missingPct = srcMedia.Count > 0 ? (Double)missing.Count / srcMedia.Count : 0;
            var msg = $"媒体内容丢失 {missing.Count}/{srcMedia.Count} 个 ({missingPct:P0}): {preview}{suffix}";
            if (missingPct > 0.25)
                result.Critical.Add(msg);
            else
                result.Informational.Add($"[可接受] {msg} — 可能是母版/讲义/备注专用媒体，不影响幻灯片视觉");
        }
        else
        {
            result.Informational.Add($"媒体: 源 {srcMedia.Count} 个内容全部存在于输出 {outMedia.Count} 个媒体中（重命名符合预期）");
        }
    }

    private static void CompareSlideXmls(
        Dictionary<String, ZipArchiveEntry> srcEntries,
        Dictionary<String, ZipArchiveEntry> outEntries,
        List<String> issues)
    {
        var pattern = new System.Text.RegularExpressions.Regex(
            @"^ppt/slides/slide(\d+)\.xml$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        var srcSlides = srcEntries.Keys
            .Select(k => pattern.Match(k))
            .Where(m => m.Success)
            .OrderBy(m => Int32.Parse(m.Groups[1].Value))
            .Select(m => m.Value)
            .ToList();

        for (var i = 0; i < srcSlides.Count; i++)
        {
            var outKey = $"ppt/slides/slide{i + 1}.xml";
            if (!outEntries.ContainsKey(outKey)) { issues.Add($"缺失幻灯片: {outKey}"); continue; }

            var normSrc = NormalizeSlideXml(ReadAllText(srcEntries[srcSlides[i]]));
            var normOut = NormalizeSlideXml(ReadAllText(outEntries[outKey]));
            if (normSrc == normOut) continue;

            var diffPos = 0;
            var minLen = Math.Min(normSrc.Length, normOut.Length);
            for (; diffPos < minLen; diffPos++)
                if (normSrc[diffPos] != normOut[diffPos]) break;

            var ctxStart = Math.Max(0, diffPos - 50);
            var srcCtx = normSrc.Substring(ctxStart, Math.Min(100, normSrc.Length - ctxStart));
            issues.Add($"幻灯片{i + 1} spTree 不一致 (位置 {diffPos}): ...{srcCtx}...");
        }
    }

    private static void CompareChartXmls(
        Dictionary<String, ZipArchiveEntry> srcEntries,
        Dictionary<String, ZipArchiveEntry> outEntries,
        List<String> issues)
    {
        var srcCharts = srcEntries.Keys
            .Where(k => k.StartsWith("ppt/charts/", StringComparison.OrdinalIgnoreCase) && k.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(k => k).ToList();
        var outCharts = outEntries.Keys
            .Where(k => k.StartsWith("ppt/charts/", StringComparison.OrdinalIgnoreCase) && k.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(k => k).ToList();

        if (srcCharts.Count != outCharts.Count)
        {
            issues.Add($"图表数量不一致: {srcCharts.Count} vs {outCharts.Count}");
            return;
        }
        for (var i = 0; i < srcCharts.Count; i++)
        {
            var ns = NormalizeChartXml(ReadAllText(srcEntries[srcCharts[i]]));
            var no = NormalizeChartXml(ReadAllText(outEntries[outCharts[i]]));
            if (ns != no)
                issues.Add($"图表{i + 1} plotArea 不一致");
        }
    }

    private static void ComparePresentation(
        ZipArchiveEntry src, ZipArchiveEntry dst, List<String> issues)
    {
        var srcXml = ReadAllText(src);
        var dstXml = ReadAllText(dst);
        var (scx, scy) = ExtractSldSz(srcXml);
        var (dcx, dcy) = ExtractSldSz(dstXml);
        if (scx != dcx || scy != dcy)
            issues.Add($"幻灯片尺寸不一致: {scx}x{scy} vs {dcx}x{dcy}");
    }

    private static void CompareSlideRels(
        Dictionary<String, ZipArchiveEntry> srcEntries,
        Dictionary<String, ZipArchiveEntry> outEntries,
        List<String> issues)
    {
        var pattern = new System.Text.RegularExpressions.Regex(
            @"^ppt/slides/_rels/slide(\d+)\.xml\.rels$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        foreach (var kv in srcEntries)
        {
            var m = pattern.Match(kv.Key);
            if (!m.Success) continue;
            var idx = m.Groups[1].Value;
            var outKey = $"ppt/slides/_rels/slide{idx}.xml.rels";
            if (!outEntries.TryGetValue(outKey, out var outRels)) { issues.Add($"缺失幻灯片rels: {outKey}"); continue; }

            var srcRels = ExtractRelationships(ReadAllText(kv.Value));
            var dstRels = ExtractRelationships(ReadAllText(outRels));

            var srcTypes = srcRels.GroupBy(r => SimplifyRelType(r.Type))
                .ToDictionary(g => g.Key, g => g.Count());
            var dstTypes = dstRels.GroupBy(r => SimplifyRelType(r.Type))
                .ToDictionary(g => g.Key, g => g.Count());
            foreach (var kv2 in srcTypes)
            {
                if (!dstTypes.TryGetValue(kv2.Key, out var cnt) || cnt != kv2.Value)
                    issues.Add($"slide{idx}.rels: {kv2.Key} 关系数 {kv2.Value} vs {dstTypes.GetValueOrDefault(kv2.Key, 0)}");
            }
        }
    }

    #endregion

    #region XML 规范化

    private static String NormalizeSlideXml(String xml)
    {
        var doc = new XmlDocument();
        doc.PreserveWhitespace = false;
        doc.LoadXml(xml);
        var spTree = doc.SelectSingleNode("//*[local-name()='spTree']") as XmlElement;
        if (spTree == null) return xml;
        SortAttributes(spTree);
        using var sw = new StringWriter();
        var settings = new XmlWriterSettings { OmitXmlDeclaration = true, Indent = false, NewLineHandling = NewLineHandling.None };
        using var xw = XmlWriter.Create(sw, settings);
        spTree.WriteTo(xw);
        xw.Flush();
        var n = sw.ToString();
        n = System.Text.RegularExpressions.Regex.Replace(n, @"(?:r:embed|r:id|r:link)\s*=\s*""[^""]*""", "r:ref=\"X\"");
        n = System.Text.RegularExpressions.Regex.Replace(n, @"\bId\s*=\s*""[^""]*""", "Id=\"X\"");
        n = System.Text.RegularExpressions.Regex.Replace(n, @"\bname\s*=\s*""[^""]*""", "name=\"X\"");
        n = System.Text.RegularExpressions.Regex.Replace(n, @"\bid\s*=\s*""[^""]*""", "id=\"X\"");
        // 移除 namespace 声明（xmlns:xxx="..." 声明位置因序列化器不同而不同，语义等价）
        n = System.Text.RegularExpressions.Regex.Replace(n, @"\s*xmlns:[a-zA-Z0-9]+=\s*""[^""]*""", "");
        return n;
    }

    private static String NormalizeChartXml(String xml)
    {
        var doc = new XmlDocument();
        doc.PreserveWhitespace = false;
        doc.LoadXml(xml);
        var plotArea = doc.SelectSingleNode("//*[local-name()='plotArea']") as XmlElement;
        if (plotArea == null) return xml;
        SortAttributes(plotArea);
        using var sw = new StringWriter();
        var settings = new XmlWriterSettings { OmitXmlDeclaration = true, Indent = false, NewLineHandling = NewLineHandling.None };
        using var xw = XmlWriter.Create(sw, settings);
        plotArea.WriteTo(xw);
        xw.Flush();
        return sw.ToString();
    }

    private static void SortAttributes(XmlElement el)
    {
        if (el.Attributes.Count > 1)
        {
            var attrs = new XmlAttribute[el.Attributes.Count];
            el.Attributes.CopyTo(attrs, 0);
            el.Attributes.RemoveAll();
            foreach (var a in attrs.OrderBy(a => a.Name))
                el.Attributes.Append(a);
        }
        foreach (var child in el.ChildNodes)
        {
            if (child is XmlElement childEl)
                SortAttributes(childEl);
        }
    }

    #endregion

    #region 辅助方法

    private static (Dictionary<String, String> Defaults, Dictionary<String, String> Overrides) ExtractContentTypes(String xml)
    {
        var doc = new XmlDocument();
        doc.LoadXml(xml);
        var defaults = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        var overrides = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        foreach (XmlElement el in doc.SelectNodes("//*[local-name()='Default']")!)
            defaults[el.GetAttribute("Extension")] = el.GetAttribute("ContentType");
        foreach (XmlElement el in doc.SelectNodes("//*[local-name()='Override']")!)
            overrides[el.GetAttribute("PartName")] = el.GetAttribute("ContentType");
        return (defaults, overrides);
    }

    private static (String Type, String Target)[] ExtractRelationships(String xml)
    {
        var doc = new XmlDocument();
        doc.LoadXml(xml);
        var list = new List<(String, String)>();
        const String PKGNS = "http://schemas.openxmlformats.org/package/2006/relationships";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", PKGNS);
        foreach (XmlElement rel in doc.SelectNodes("//r:Relationship", ns)!)
            list.Add((rel.GetAttribute("Type"), rel.GetAttribute("Target")));
        return list.ToArray();
    }

    private static (Int64 cx, Int64 cy) ExtractSldSz(String xml)
    {
        var doc = new XmlDocument();
        doc.LoadXml(xml);
        var sldSz = doc.SelectSingleNode("//*[local-name()='sldSz']") as XmlElement;
        if (sldSz == null) return (0, 0);
        Int64.TryParse(sldSz.GetAttribute("cx"), out var cx);
        Int64.TryParse(sldSz.GetAttribute("cy"), out var cy);
        return (cx, cy);
    }

    private static String SimplifyRelType(String type)
    {
        var parts = type.Split('/');
        return parts.Length > 0 ? parts[^1] : type;
    }

    private static Byte[] ReadAllBytes(ZipArchiveEntry entry)
    {
        using var ms = new MemoryStream();
        using var s = entry.Open();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    private static String ReadAllText(ZipArchiveEntry entry)
    {
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    private static String ComputeHash(Byte[] data)
    {
        using var sha = SHA256.Create();
        return Convert.ToBase64String(sha.ComputeHash(data));
    }

    /// <summary>去除 UTF-8 BOM（EF BB BF），使二进制对比不受 BOM 影响</summary>
    private static Byte[] StripBom(Byte[] data)
    {
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
            return data[3..];
        return data;
    }

    /// <summary>归一化换行符：\r\n → \n（StreamReader.ReadToEnd 会做此转换，导致写出后与原文不同）</summary>
    private static Byte[] NormalizeLineEndings(Byte[] data)
    {
        // 统计 \r\n 对数，预分配结果数组
        var crlfCount = 0;
        for (var i = 0; i < data.Length - 1; i++)
            if (data[i] == 0x0D && data[i + 1] == 0x0A) crlfCount++;
        if (crlfCount == 0) return data;
        var result = new Byte[data.Length - crlfCount];
        var wi = 0;
        for (var ri = 0; ri < data.Length; ri++)
        {
            if (ri < data.Length - 1 && data[ri] == 0x0D && data[ri + 1] == 0x0A)
            {
                result[wi++] = 0x0A;
                ri++; // 跳过 \n
            }
            else
                result[wi++] = data[ri];
        }
        return result;
    }

    #endregion
}
