using System.IO.Compression;
using System.Text;

namespace NewLife.Office;

partial class PptxWriter
{
    #region 保存方法
    /// <summary>保存到文件</summary>
    /// <param name="path">输出路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs);
    }

    /// <summary>从 PptDocument 模型保存到文件</summary>
    /// <param name="path">输出路径</param>
    /// <param name="document">演示文稿数据模型</param>
    public void Save(String path, PptDocument document)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs, document);
    }

    /// <summary>从 PptDocument 模型保存到流</summary>
    /// <param name="stream">目标流</param>
    /// <param name="document">演示文稿数据模型</param>
    public void Save(Stream stream, PptDocument document)
    {
        // 将文档模型属性应用到 writer
        SlideWidth  = document.SlideWidth;
        SlideHeight = document.SlideHeight;
        if (document.AccentColors is { Length: > 0 })
            SetAccentColors(document.AccentColors);
        if (document.Master != null)
        {
            _progMasters.Clear();
            _progMasters.Add(document.Master);
        }
        // 幻灯片：清空现有后逐一添加
        Slides.Clear();
        foreach (var slide in document.Slides)
            Slides.Add(slide);
        if (!document.Properties.Password.IsNullOrEmpty())
            SetProtection(document.Properties.Password);
        _documentProperties = document.Properties;
        HeaderFooter = document.HeaderFooter;
        Sections = document.Sections;
        Save(stream);
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">目标流</param>
    public void Save(Stream stream)
    {
        using var za = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteContentTypes(za);
        WriteRels(za);
        WritePresentation(za);
        WritePresentationRels(za);
        WriteSlideLayout(za);
        WriteSlideMaster(za);
        WriteInfraMedia(za);
        for (var i = 0; i < Slides.Count; i++)
        {
            WriteSlide(za, i, Slides[i]);
        }
        // 写入跨文件复制的原始幻灯片（S10-04）
        var totalSlides = Slides.Count;
        for (var ri = 0; ri < _rawSlides.Count; ri++)
        {
            var rawIdx = totalSlides + ri;
            WriteZipEntryText(za, $"ppt/slides/slide{rawIdx + 1}.xml", _rawSlides[ri].SlideXml);
            WriteZipEntryText(za, $"ppt/slides/_rels/slide{rawIdx + 1}.xml.rels", _rawSlides[ri].RelsXml);
        }
        // 写入原始幻灯片的媒体文件
        foreach (var (name, data) in _rawSlideMedia)
        {
            var entry = za.CreateEntry($"ppt/media/{name}", CompressionLevel.Fastest);
            using var es = entry.Open();
            es.Write(data, 0, data.Length);
        }
        // 写入嵌入字体（ppt/fonts/）
        foreach (var kv in _embeddedFonts)
        {
            var fontEntry = za.CreateEntry($"ppt/fonts/{kv.Key}", CompressionLevel.Fastest);
            using var es = fontEntry.Open();
            es.Write(kv.Value, 0, kv.Value.Length);
        }
        WriteTheme(za);
        WriteDocProps(za);
        WriteComments(za);
    }
    #endregion

    #region 私有方法
    private PptSlide EnsureSlide(Int32 idx)
    {
        while (Slides.Count <= idx)
        {
            Slides.Add(new PptSlide());
        }
        return Slides[idx];
    }

    /// <summary>厘米转换为 EMU（English Metric Units）</summary>
    /// <param name="cm">厘米值</param>
    /// <returns>EMU 值（1 cm = 360000 EMU）</returns>
    public static Int64 CmToEmu(Double cm) => (Int64)(cm * 360000);

    /// <summary>EMU 转换为厘米</summary>
    /// <param name="emu">EMU 值</param>
    /// <returns>厘米值（1 cm = 360000 EMU）</returns>
    public static Double EmuToCm(Int64 emu) => emu / 360000.0;

    /// <summary>磅（点/pt）转换为 EMU</summary>
    /// <param name="pt">磅值</param>
    /// <returns>EMU 值（1 pt = 12700 EMU）</returns>
    public static Int64 PtToEmu(Double pt) => (Int64)(pt * 12700);

    /// <summary>EMU 转换为磅（点/pt）</summary>
    /// <param name="emu">EMU 值</param>
    /// <returns>磅值（1 pt = 12700 EMU）</returns>
    public static Double EmuToPt(Int64 emu) => emu / 12700.0;

    private void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), Encoding.UTF8);
        sw.Write(content);
    }

    private static void WriteZipEntryText(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), Encoding.UTF8);
        sw.Write(content);
    }

    private void WriteContentTypes(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>");
        var masterCount = Math.Max(1, _masterContents.Count);
        for (var i = 1; i <= masterCount; i++)
            sb.Append($"<Override PartName=\"/ppt/slideMasters/slideMaster{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>");
        var layoutCount = Math.Max(1, _layoutContents.Count);
        if (_layoutContents.Count == 0 && _progMasters.Count > 0)
            layoutCount = Math.Max(1, _progMasters.Sum(m => m.Layouts.Count));
        for (var i = 1; i <= layoutCount; i++)
            sb.Append($"<Override PartName=\"/ppt/slideLayouts/slideLayout{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>");
        sb.Append("<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
        for (var i = 0; i < Slides.Count; i++)
        {
            sb.Append($"<Override PartName=\"/ppt/slides/slide{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
        }
        // 原始幻灯片内容类型（S10-04）
        for (var i = 0; i < _rawSlides.Count; i++)
        {
            sb.Append($"<Override PartName=\"/ppt/slides/slide{Slides.Count + i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
        }
        // chart types
        foreach (var slide in Slides)
        {
            foreach (var chart in slide.Charts)
            {
                sb.Append($"<Override PartName=\"/ppt/charts/chart{chart.ChartNumber}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
            }
        }
        // image types
        var addedExt = new HashSet<String>();
        foreach (var slide in Slides)
        {
            foreach (var img in slide.Images)
            {
                if (addedExt.Add(img.Extension))
                {
                    var ct = img.Extension is "jpg" or "jpeg" ? "image/jpeg"
                           : img.Extension == "svg" ? "image/svg+xml"
                           : "image/png";
                    sb.Append($"<Default Extension=\"{img.Extension}\" ContentType=\"{ct}\"/>");
                }
            }
        }
        // 嵌入字体类型
        if (_embeddedFonts.Count > 0)
            sb.Append("<Default Extension=\"fntdata\" ContentType=\"application/vnd.ms-office.activeX+xml\"/>");

        // docProps（S14）
        if (_documentProperties != null)
        {
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        }

        // comments（S13-01）
        if (Slides.Any(s => s.Comments.Count > 0))
        {
            sb.Append("<Override PartName=\"/ppt/comments/commentAuthors.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml\"/>");
            for (var i = 0; i < Slides.Count; i++)
            {
                if (Slides[i].Comments.Count > 0)
                    sb.Append($"<Override PartName=\"/ppt/comments/comment{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.comments+xml\"/>");
            }
        }

        sb.Append("</Types>");
        WriteEntry(za, "[Content_Types].xml", sb.ToString());
    }

    private void WriteRels(ZipArchive za) =>
        WriteEntry(za, "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>" +
            "</Relationships>");

    private void WritePresentation(ZipArchive za)
    {
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        var hasSections = Sections != null && Sections.Count > 0;
        var nsExt = hasSections ? " xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"" : "";
        sb.Append($"<p:presentation xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\"{nsExt} saveSubsetFonts=\"1\">");
        sb.Append($"<p:sldSz cx=\"{SlideWidth}\" cy=\"{SlideHeight}\"/>");
        var masterCount = Math.Max(1, _masterContents.Count);
        if (_masterContents.Count == 0 && _progMasters.Count > 0)
            masterCount = Math.Max(1, _progMasters.Count);
        sb.Append("<p:sldMasterIdLst>");
        for (var i = 0; i < masterCount; i++)
            sb.Append($"<p:sldMasterId id=\"{2147483648 + i}\" r:id=\"rMaster{i + 1}\"/>");
        sb.Append("</p:sldMasterIdLst>");
        sb.Append("<p:sldIdLst>");
        for (var i = 0; i < Slides.Count; i++)
        {
            sb.Append($"<p:sldId id=\"{256 + i}\" r:id=\"rSlide{i + 1}\"/>");
        }
        // 原始幻灯片（S10-04）
        for (var i = 0; i < _rawSlides.Count; i++)
        {
            sb.Append($"<p:sldId id=\"{256 + Slides.Count + i}\" r:id=\"rSlide{Slides.Count + i + 1}\"/>");
        }
        sb.Append("</p:sldIdLst>");

        // Section 列表（S13-04）
        if (hasSections)
        {
            sb.Append("<p:extLst><p:ext uri=\"{521415D9-36F7-43E2-AB2F-B90AF26B5E84}\">");
            sb.Append("<p14:sectionLst>");
            for (var si = 0; si < Sections!.Count; si++)
            {
                var sec = Sections[si];
                sb.Append($"<p14:section name=\"{EscXml(sec.Name)}\"><p14:sldIdLst>");
                foreach (var idx in sec.SlideIndices)
                    sb.Append($"<p14:sldId id=\"{256 + idx}\"/>");
                sb.Append("</p14:sldIdLst></p14:section>");
            }
            sb.Append("</p14:sectionLst></p:ext></p:extLst>");
        }
        // 全局页眉页脚（S13-03）
        var hf = HeaderFooter;
        if (hf != null && (hf.ShowFooter || hf.ShowPageNumber || hf.ShowDate))
        {
            sb.Append("<p:hf");
            if (hf.ShowFooter) sb.Append($" footer=\"{EscXml(hf.FooterText ?? String.Empty)}\"");
            if (hf.ShowPageNumber) sb.Append(" showSlideNum=\"1\"");
            if (hf.ShowDate)
            {
                sb.Append(" dt=\"1\"");
                if (!hf.DateAutomatic && hf.FixedDate != null)
                    sb.Append($" fdt=\"{EscXml(hf.FixedDate)}\"");
                if (hf.DateFormat != null)
                    sb.Append($" dfmt=\"{EscXml(hf.DateFormat)}\"");
            }
            sb.Append("/>");
        }

        // 演示文稿保护（S07-04）
        if (_protectionHash != null)
            sb.Append($"<p:modifyVerifier algorithmName=\"SHA-512\" hashData=\"{_protectionHash}\" saltData=\"{_protectionSalt}\" spinCount=\"100000\"/>");
        sb.Append("</p:presentation>");
        WriteEntry(za, "ppt/presentation.xml", sb.ToString());
    }

    private void WritePresentationRels(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        var masterCount = Math.Max(1, _masterContents.Count);
        var hasExternalMaster = _masterContents.Count > 0 || _progMasters.Count > 0;
        if (_masterContents.Count == 0 && _progMasters.Count > 0)
            masterCount = Math.Max(1, _progMasters.Count);
        for (var i = 1; i <= masterCount; i++)
            sb.Append($"<Relationship Id=\"rMaster{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster{i}.xml\"/>");
        // 无模板且无编程母版时从 presentation rels 引用 theme；否则 theme 由 master rels 引用
        if (!hasExternalMaster)
            sb.Append("<Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>");
        for (var i = 0; i < Slides.Count; i++)
        {
            sb.Append($"<Relationship Id=\"rSlide{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i + 1}.xml\"/>");
        }
        // 原始幻灯片关系（S10-04）
        for (var i = 0; i < _rawSlides.Count; i++)
        {
            sb.Append($"<Relationship Id=\"rSlide{Slides.Count + i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{Slides.Count + i + 1}.xml\"/>");
        }
        sb.Append("</Relationships>");
        WriteEntry(za, "ppt/_rels/presentation.xml.rels", sb.ToString());
    }

    private void WriteSlide(ZipArchive za, Int32 idx, PptSlide slide)
    {
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var shapeId = 2;
        // 收集超链接 relId → url（用于 rels 文件）
        var hlinkMap = new Dictionary<String, String>();

        // 自动分配 RelId：对未设置 RelId 的图片/视频/背景图统一补全
        foreach (var img in slide.Images)
        {
            if (img.RelId.IsNullOrEmpty()) img.RelId = $"rImg{_imgGlobal++}";
        }
        foreach (var vid in slide.Videos)
        {
            if (vid.RelId.IsNullOrEmpty()) vid.RelId = $"rVid{_videoGlobal++}";
        }

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<p:sld xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\"");
        if (slide.Hidden) sb.Append(" show=\"0\"");
        sb.Append('>');

        // background
        if (slide.BackgroundImage != null)
        {
            var bgImg = slide.BackgroundImage;
            sb.Append("<p:bg><p:bgPr>");
            sb.Append("<a:blipFill dpi=\"0\" rotWithShape=\"1\">");
            sb.Append($"<a:blip r:embed=\"rBg{idx + 1}\"/>");
            sb.Append("<a:stretch><a:fillRect/></a:stretch></a:blipFill>");
            sb.Append("<a:effectLst/></p:bgPr></p:bg>");
            // 背景图片 rel 和媒体写在后面
        }
        else if (slide.BackgroundColor != null)
        {
            sb.Append("<p:bg><p:bgPr>");
            sb.Append($"<a:solidFill><a:srgbClr val=\"{slide.BackgroundColor.TrimStart('#')}\"/></a:solidFill>");
            sb.Append("<a:effectLst/></p:bgPr></p:bg>");
        }
        else if (slide.BackgroundGradientType != null && slide.BackgroundGradientColor1 != null && slide.BackgroundGradientColor2 != null)
        {
            sb.Append("<p:bg><p:bgPr>");
            sb.Append($"<a:gradFill><a:gsLst>");
            sb.Append($"<a:gs pos=\"0\"><a:srgbClr val=\"{slide.BackgroundGradientColor1.TrimStart('#')}\"/></a:gs>");
            sb.Append($"<a:gs pos=\"100000\"><a:srgbClr val=\"{slide.BackgroundGradientColor2.TrimStart('#')}\"/></a:gs>");
            sb.Append($"</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
            sb.Append("<a:effectLst/></p:bgPr></p:bg>");
        }

        sb.Append("<p:cSld><p:spTree>");
        sb.Append("<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
        sb.Append("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");

        // text boxes
        foreach (var tb in slide.TextBoxes)
        {
            // 处理超链接
            String? hlRelId = null;
            if (tb.HyperlinkUrl != null)
            {
                hlRelId = $"rHlk{_hlinkGlobal++}";
                hlinkMap[hlRelId] = tb.HyperlinkUrl;
            }
            var descr = tb.AltText != null ? $" descr=\"{EscXml(tb.AltText)}\"" : "";
            sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"TextBox\"{descr}/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{tb.Left}\" y=\"{tb.Top}\"/><a:ext cx=\"{tb.Width}\" cy=\"{tb.Height}\"/>");
            if (tb.Rotation != 0) sb.Append($" rot=\"{tb.Rotation}\"");
            sb.Append("</a:xfrm>");
            sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            if (tb.BackgroundColor != null)
                sb.Append($"<a:solidFill><a:srgbClr val=\"{tb.BackgroundColor.TrimStart('#')}\"/></a:solidFill>");
            else
                sb.Append("<a:noFill/>");
            sb.Append("</p:spPr>");
            // bodyPr auto-fit + 内边距 + anchor
            var fitTag = tb.AutoFit switch { 1 => "<a:spAutoFit/>", 2 => "<a:noAutofit/>", _ => "<a:normAutofit/>" };
            var insets = new StringBuilder();
            if (tb.LeftInset > 0) insets.Append($" lIns=\"{tb.LeftInset}\"");
            if (tb.RightInset > 0) insets.Append($" rIns=\"{tb.RightInset}\"");
            if (tb.TopInset > 0) insets.Append($" tIns=\"{tb.TopInset}\"");
            if (tb.BottomInset > 0) insets.Append($" bIns=\"{tb.BottomInset}\"");
            var anchorAttr = tb.Anchor.Length > 0 ? $" anchor=\"{tb.Anchor}\"" : "";
            var vertAttr = tb.TextDirection != null ? $" vert=\"{tb.TextDirection}\"" : "";
            sb.Append($"<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"{insets}{anchorAttr}{vertAttr}>{fitTag}</a:bodyPr><a:lstStyle/>");

            // 段落写入：优先使用 Paragraphs（多段结构），回退到 Runs（单段兼容）
            if (tb.Paragraphs.Count > 0)
            {
                foreach (var pp in tb.Paragraphs)
                {
                    var hasPPrChildren = pp.LineSpacingPct > 0 || pp.LineSpacingPts > 0 || pp.SpaceBeforePt > 0 || pp.SpaceAfterPt > 0 || pp.BulletChar != null || pp.BulletNone;
                    sb.Append($"<a:p><a:pPr{(pp.Alignment.Length > 0 && pp.Alignment != "l" ? " algn=\"" + pp.Alignment + "\"" : "")}");
                    if (pp.Level > 0)
                        sb.Append($" lvl=\"{pp.Level}\"");
                    if (hasPPrChildren)
                    {
                        sb.Append('>');
                        if (pp.LineSpacingPct > 0)
                            sb.Append($"<a:lnSpc><a:spcPct val=\"{pp.LineSpacingPct}\"/></a:lnSpc>");
                        else if (pp.LineSpacingPts > 0)
                            sb.Append($"<a:lnSpc><a:spcPts val=\"{pp.LineSpacingPts}\"/></a:lnSpc>");
                        if (pp.SpaceBeforePt > 0)
                            sb.Append($"<a:spcBef><a:spcPts val=\"{pp.SpaceBeforePt * 100}\"/></a:spcBef>");
                        if (pp.SpaceAfterPt > 0)
                            sb.Append($"<a:spcAft><a:spcPts val=\"{pp.SpaceAfterPt * 100}\"/></a:spcAft>");
                        if (pp.BulletChar != null)
                            sb.Append($"<a:buChar char=\"{EscXml(pp.BulletChar)}\"/>");
                        else if (pp.BulletNone)
                            sb.Append("<a:buNone/>");
                        sb.Append("</a:pPr>");
                    }
                    else
                        sb.Append("/>");

                    foreach (var run in pp.Runs)
                    {
                        WriteTextRun(sb, run, tb, ref _hlinkGlobal, hlinkMap);
                    }
                    sb.Append("</a:p>");
                }
            }
            else
            {
                // 向后兼容：Runs 单段模式
                var hasPPrChildren = tb.LineSpacingPct > 0 || tb.SpaceBeforePt > 0;
                sb.Append($"<a:p><a:pPr{(tb.Alignment.Length > 0 ? " algn=\"" + tb.Alignment + "\"" : "")}");
                if (hasPPrChildren)
                {
                    sb.Append('>');
                    if (tb.LineSpacingPct > 0)
                        sb.Append($"<a:lnSpc><a:spcPct val=\"{tb.LineSpacingPct}\"/></a:lnSpc>");
                    if (tb.SpaceBeforePt > 0)
                        sb.Append($"<a:spcBef><a:spcPts val=\"{tb.SpaceBeforePt * 100}\"/></a:spcBef>");
                    sb.Append("</a:pPr>");
                }
                else
                    sb.Append("/>");
                if (tb.Runs.Count > 0)
                {
                    foreach (var run in tb.Runs)
                    {
                        WriteTextRun(sb, run, tb, ref _hlinkGlobal, hlinkMap);
                    }
                }
                else
                {
                    WriteSingleLineTextRun(sb, tb, hlRelId);
                }
                sb.Append("</a:p>");
            }
            sb.Append("</p:txBody></p:sp>");
        }

        // shapes（基本图形）
        foreach (var sp in slide.Shapes)
        {
            var descr = sp.AltText != null ? $" descr=\"{EscXml(sp.AltText)}\"" : "";
            sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"Shape\"{descr}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{sp.Left}\" y=\"{sp.Top}\"/><a:ext cx=\"{sp.Width}\" cy=\"{sp.Height}\"/>");
            if (sp.Rotation != 0) sb.Append($" rot=\"{sp.Rotation}\"");
            if (sp.FlipHorizontal) sb.Append(" flipH=\"1\"");
            if (sp.FlipVertical) sb.Append(" flipV=\"1\"");
            sb.Append("</a:xfrm>");
            // 圆角矩形：写入 adj 调整值（CornerRadius=EMU，转换为 OOXML adj 值占宽度百分比*50000）
            if (sp.ShapeType == "roundRect" && sp.CornerRadius > 0)
            {
                var adjVal = sp.CornerRadius * 50000 / sp.Width;
                sb.Append($"<a:prstGeom prst=\"roundRect\"><a:avLst><a:gd name=\"adj\" fmla=\"val {adjVal}\"/></a:avLst></a:prstGeom>");
            }
            else
                sb.Append($"<a:prstGeom prst=\"{sp.ShapeType}\"><a:avLst/></a:prstGeom>");
            if (sp.FillImage != null && sp.FillImage.Length > 0)
            {
                sp.ShapeImageRelId ??= "rImgFill1";
                sb.Append($"<a:blipFill dpi=\"0\" rotWithShape=\"1\">");
                sb.Append($"<a:blip r:embed=\"{sp.ShapeImageRelId}\"/>");
                sb.Append("<a:stretch><a:fillRect/></a:stretch></a:blipFill>");
            }
            else if (sp.GradientType != null && sp.GradientColor1 != null && sp.GradientColor2 != null)
            {
                var angle = sp.GradientAngle * 60000; // 度 → DrawingML 单位（1/60000 度）
                var isLinear = sp.GradientType == "linear";
                sb.Append($"<a:gradFill rotWithShape=\"1\"><a:gsLst>");
                sb.Append($"<a:gs pos=\"0\"><a:srgbClr val=\"{sp.GradientColor1.TrimStart('#')}\"/></a:gs>");
                sb.Append($"<a:gs pos=\"100000\"><a:srgbClr val=\"{sp.GradientColor2.TrimStart('#')}\"/></a:gs>");
                sb.Append("</a:gsLst>");
                if (isLinear)
                    sb.Append($"<a:lin ang=\"{angle}\" scaled=\"1\"/>");
                else
                    sb.Append("<a:rad sr=\"50000\"/>");
                sb.Append("<a:tileRect/></a:gradFill>");
            }
            else if (sp.FillColor != null)
                sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FillColor.TrimStart('#')}\"/></a:solidFill>");
            else
                sb.Append("<a:noFill/>");
            if (sp.LineColor != null)
                sb.Append($"<a:ln w=\"{sp.LineWidth}\"><a:solidFill><a:srgbClr val=\"{sp.LineColor.TrimStart('#')}\"/></a:solidFill>");
            else
                sb.Append("<a:ln><a:noFill/>");
            if (sp.DashStyle != null)
                sb.Append($"<a:prstDash val=\"{sp.DashStyle}\"/>");
            sb.Append("</a:ln>");
            sb.Append("</p:spPr>");
            if (sp.Text != null)
            {
                // bodyPr with text direction, autofit, and insets
                var vertAttr = sp.TextDirection != null ? $" vert=\"{sp.TextDirection}\"" : "";
                var anchorAttr = sp.Anchor != null ? $" anchor=\"{sp.Anchor}\"" : "";
                var insetL = sp.LeftInset ?? 25400;
                var insetR = sp.RightInset ?? 25400;
                var insetT = sp.TopInset ?? 12700;
                var insetB = sp.BottomInset ?? 12700;
                var autoFitXml = sp.TextAutoFit == "norm" ? "<a:normAutofit/>" : sp.TextAutoFit == "noAuto" ? "<a:noAutofit/>" : "";
                sb.Append($"<p:txBody><a:bodyPr lIns=\"{insetL}\" rIns=\"{insetR}\" tIns=\"{insetT}\" bIns=\"{insetB}\"{vertAttr}{anchorAttr}>{autoFitXml}</a:bodyPr><a:lstStyle/><a:p><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{sp.FontSize * 100}\"{(sp.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                WriteFontColor(sb, sp.FontColor);
                WriteFontElements(sb, sp.LatinFontName, sp.EastAsianFontName, sp.ComplexScriptFontName, sp.SymbolFontName);
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(sp.Text)}</a:t>");
                sb.Append("</a:r></a:p></p:txBody>");
            }
            sb.Append("</p:sp>");
        }

        // images
        foreach (var img in slide.Images)
        {
            sb.Append($"<p:pic><p:nvPicPr><p:cNvPr id=\"{shapeId++}\" name=\"Image\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>");
            sb.Append("<p:blipFill>");
            sb.Append($"<a:blip r:embed=\"{img.RelId}\"");
            if (img.IsSvg)
            {
                // SVG: 在 <a:blip> 内嵌套 <asvg:svgBlip>，SEO/兼容性更好
                sb.Append($" xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\"><asvg:svgBlip r:embed=\"{img.RelId}\"/></a:blip>");
            }
            else
            {
                sb.Append("/>");
            }
            sb.Append("<a:stretch><a:fillRect/></a:stretch></p:blipFill>");
            sb.Append("<p:spPr>");
            var rotAttr = img.Rotation != 0 ? $" rot=\"{img.Rotation}\"" : "";
            sb.Append($"<a:xfrm{rotAttr}><a:off x=\"{img.Left}\" y=\"{img.Top}\"/><a:ext cx=\"{img.Width}\" cy=\"{img.Height}\"/></a:xfrm>");
            if (img.CornerRadius > 0)
                sb.Append($"<a:prstGeom prst=\"roundRect\"><a:avLst><a:gd name=\"adj\" fmla=\"val {img.CornerRadius}\"/></a:avLst></a:prstGeom></p:spPr></p:pic>");
            else
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
        }

        // videos
        foreach (var vid in slide.Videos)
        {
            // 确保视频有缩略图 RelId（若 AddVideo 未分配则补充）
            if (vid.ThumbnailRelId.Length == 0)
                vid.ThumbnailRelId = $"rVidThumb{_videoGlobal}";
            // 若无缩略图数据，生成最小 PNG 占位
            if (vid.ThumbnailData == null || vid.ThumbnailData.Length == 0)
            {
                vid.ThumbnailData = MinimalPng;
                vid.ThumbnailExtension = "png";
            }
            sb.Append($"<p:pic><p:nvPicPr><p:cNvPr id=\"{shapeId++}\" name=\"Video\"><a:hlinkClick action=\"ppaction://media\"/></p:cNvPr>");
            sb.Append("<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>");
            sb.Append($"<p:nvPr><a:videoFile r:link=\"{vid.RelId}\"/></p:nvPr></p:nvPicPr>");
            sb.Append($"<p:blipFill><a:blip r:embed=\"{vid.ThumbnailRelId}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{vid.Left}\" y=\"{vid.Top}\"/><a:ext cx=\"{vid.Width}\" cy=\"{vid.Height}\"/></a:xfrm>");
            sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
        }

        // tables
        foreach (var tbl in slide.Tables)
        {
            BuildPptTableXml(sb, tbl, ref shapeId);
        }

        // charts
        foreach (var chart in slide.Charts)
        {
            sb.Append($"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{shapeId++}\" name=\"Chart\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>");
            sb.Append($"<p:xfrm><a:off x=\"{chart.Left}\" y=\"{chart.Top}\"/><a:ext cx=\"{chart.Width}\" cy=\"{chart.Height}\"/></p:xfrm>");
            sb.Append($"<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">");
            sb.Append($"<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"{chart.RelId}\"/>");
            sb.Append("</a:graphicData></a:graphic></p:graphicFrame>");
        }

        // groups（形状组，S07-02）
        foreach (var grp in slide.Groups)
        {
            sb.Append($"<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"{shapeId++}\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
            sb.Append("<p:grpSpPr>");
            sb.Append($"<a:xfrm><a:off x=\"{grp.Left}\" y=\"{grp.Top}\"/><a:ext cx=\"{grp.Width}\" cy=\"{grp.Height}\"/>");
            sb.Append($"<a:chOff x=\"{grp.Left}\" y=\"{grp.Top}\"/><a:chExt cx=\"{grp.Width}\" cy=\"{grp.Height}\"/></a:xfrm>");
            sb.Append("</p:grpSpPr>");
            // shapes inside group
            foreach (var sp in grp.Shapes)
            {
                sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"GrpShape\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>");
                sb.Append("<p:spPr>");
                sb.Append($"<a:xfrm><a:off x=\"{sp.Left}\" y=\"{sp.Top}\"/><a:ext cx=\"{sp.Width}\" cy=\"{sp.Height}\"/></a:xfrm>");
                sb.Append($"<a:prstGeom prst=\"{sp.ShapeType}\"><a:avLst/></a:prstGeom>");
                if (sp.FillColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FillColor.TrimStart('#')}\"/></a:solidFill>");
                else
                    sb.Append("<a:noFill/>");
                sb.Append("</p:spPr>");
                if (sp.Text != null)
                {
                    sb.Append("<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>");
                    sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{sp.FontSize * 100}\"{(sp.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                    WriteFontColor(sb, sp.FontColor);
                    WriteFontElements(sb, sp.LatinFontName, sp.EastAsianFontName, sp.ComplexScriptFontName, sp.SymbolFontName);
                    sb.Append("</a:rPr>");
                    sb.Append($"<a:t>{EscXml(sp.Text)}</a:t>");
                    sb.Append("</a:r></a:p></p:txBody>");
                }
                sb.Append("</p:sp>");
            }
            // text boxes inside group
            foreach (var tb in grp.TextBoxes)
            {
                sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"GrpTextBox\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>");
                sb.Append("<p:spPr>");
                sb.Append($"<a:xfrm><a:off x=\"{tb.Left}\" y=\"{tb.Top}\"/><a:ext cx=\"{tb.Width}\" cy=\"{tb.Height}\"/></a:xfrm>");
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/>");
                sb.Append("</p:spPr>");
                sb.Append("<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"><a:normAutofit/></a:bodyPr><a:lstStyle/>");
                sb.Append($"<a:p><a:pPr algn=\"{tb.Alignment}\"/><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{tb.FontSize * 100}\"{(tb.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                WriteFontColor(sb, tb.FontColor);
                WriteFontElements(sb, tb.LatinFontName, tb.EastAsianFontName, tb.ComplexScriptFontName, tb.SymbolFontName);
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(tb.Text)}</a:t>");
                sb.Append("</a:r></a:p></p:txBody></p:sp>");
            }
            sb.Append("</p:grpSp>");
        }

        // connectors（连接器，S13-02）
        foreach (var cn in slide.Connectors)
        {
            sb.Append($"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id=\"{shapeId++}\" name=\"Connector\"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{cn.Left}\" y=\"{cn.Top}\"/><a:ext cx=\"{cn.Width}\" cy=\"{cn.Height}\"/></a:xfrm>");
            var lineColor = cn.LineColor ?? "000000";
            var prstDash = cn.DashStyle != null ? $"<a:prstDash val=\"{cn.DashStyle}\"/>" : "";
            var tailEnd = cn.StartArrow != null ? $"<a:tailEnd type=\"{cn.StartArrow}\"/>" : "";
            var headEnd = cn.EndArrow != null ? $"<a:headEnd type=\"{cn.EndArrow}\"/>" : "";
            sb.Append($"<a:prstGeom prst=\"{cn.ConnectorType}Connector1\"><a:avLst/></a:prstGeom>");
            sb.Append($"<a:ln w=\"{cn.LineWidth}\">{prstDash}<a:solidFill><a:srgbClr val=\"{lineColor}\"/></a:solidFill>{tailEnd}{headEnd}</a:ln>");
            sb.Append("</p:spPr></p:cxnSp>");
        }

        sb.Append("</p:spTree></p:cSld>");

        // notes
        if (slide.Notes != null)
        {
            sb.Append("<p:notes><p:cSld><p:spTree>");
            sb.Append("<p:sp><p:nvSpPr><p:cNvPr id=\"1\" name=\"notes\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\"/></p:nvPr></p:nvSpPr>");
            sb.Append("<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>");
            sb.Append($"<a:p><a:r><a:rPr lang=\"zh-CN\" dirty=\"0\"/><a:t>{EscXml(slide.Notes)}</a:t></a:r></a:p>");
            sb.Append("</p:txBody></p:sp></p:spTree></p:cSld></p:notes>");
        }

        // 转场动画
        if (slide.Transition != null)
        {
            var t = slide.Transition;
            sb.Append($"<p:transition dur=\"{t.DurationMs}\" {(t.AdvanceOnClick ? "advClick=\"1\"" : "advClick=\"0\"")}>");
            sb.Append(t.Type switch
            {
                "fade" => "<p:fade/>",
                "push" => $"<p:push dir=\"{t.Direction}\"/>",
                "wipe" => $"<p:wipe dir=\"{t.Direction}\"/>",
                "zoom" => "<p:zoom/>",
                "split" => "<p:split/>",
                "cut" => "<p:cut/>",
                "morph" => "<p:morph option=\"byObject\"/>",
                _ => "<p:fade/>",
            });
            sb.Append("</p:transition>");
        }

        // 元素动画（S12-01）
        if (slide.Animations.Count > 0)
        {
            sb.Append("<p:timing><p:tnLst><p:par><p:cTn id=\"1\" dur=\"indefinite\" fill=\"hold\">");
            sb.Append("<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn><p:childTnLst>");
            foreach (var anim in slide.Animations.OrderBy(a => a.Order))
            {
                var nodeName = anim.Category switch
                {
                    PptAnimationCategory.Entrance => "animEffect",
                    PptAnimationCategory.Emphasis => "animEmph",
                    PptAnimationCategory.Exit => "animEffect",
                    PptAnimationCategory.MotionPath => "animMotion",
                    _ => "animEffect",
                };
                var animType = anim.Category switch
                {
                    PptAnimationCategory.Entrance => "entr",
                    PptAnimationCategory.Emphasis => "emph",
                    PptAnimationCategory.Exit => "exit",
                    _ => "entr",
                };
                var trigger = anim.Trigger switch
                {
                    PptAnimationTrigger.WithPrevious => "withPrev",
                    PptAnimationTrigger.AfterPrevious => "afterPrev",
                    _ => "onClick",
                };
                sb.Append($"<p:par><p:cTn nodeType=\"{trigger}\" fill=\"hold\">");
                sb.Append("<p:stCondLst>");
                sb.Append(anim.Trigger switch
                {
                    PptAnimationTrigger.OnClick => "<p:cond delay=\"indefinite\"/>",
                    PptAnimationTrigger.WithPrevious => "<p:cond delay=\"0\"/>",
                    PptAnimationTrigger.AfterPrevious => "<p:cond delay=\"0\"/>",
                    _ => "<p:cond delay=\"indefinite\"/>",
                });
                sb.Append("</p:stCondLst></p:cTn><p:childTnLst>");
                sb.Append($"<p:{nodeName} filter=\"{EscXml(anim.Effect)}\"{(anim.Category == PptAnimationCategory.Exit ? " exit=\"1\"" : "")}>");
                sb.Append("<p:cBhvr><p:cTn");
                if (anim.DurationMs > 0)
                    sb.Append($" dur=\"{anim.DurationMs}\"");
                if (anim.DelayMs > 0)
                    sb.Append($" st=\"{(anim.Trigger == PptAnimationTrigger.AfterPrevious ? anim.DelayMs : 0)}\"");
                sb.Append($">");
                if (anim.DelayMs > 0)
                    sb.Append($"<p:stCondLst><p:cond delay=\"{anim.DelayMs}\"/></p:stCondLst>");
                sb.Append("</p:cTn>");
                // 目标元素：按 TargetType 和 TargetIndex 定位
                sb.Append("<p:tgtEl>");
                var spTargetType = anim.TargetType switch { "shape" => "spTgt", "chart" => "chartTgt", "table" => "oleChartTgt", _ => "spTgt" };
                var targetId = anim.TargetIndex + 2; // shapeId 从 2 开始
                sb.Append($"<p:{spTargetType} spid=\"{targetId}\"/>");
                sb.Append("</p:tgtEl>");
                sb.Append($"<p:attrNameLst><p:attrName>style.{animType}Effect</p:attrName></p:attrNameLst>");
                sb.Append("</p:cBhvr></p:animEffect>");
                sb.Append("</p:childTnLst></p:par>");
            }
            sb.Append("</p:childTnLst></p:par></p:tnLst></p:timing>");
        }

        sb.Append("</p:sld>");
        WriteEntry(za, $"ppt/slides/slide{idx + 1}.xml", sb.ToString());

        // slide rels
        var relsSb = new StringBuilder();
        relsSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        var layoutCount = _layoutContents.Count > 0
            ? _layoutContents.Count
            : Math.Max(1, _progMasters.Sum(m => m.Layouts.Count));
        var layoutIdx = Math.Max(0, Math.Min(slide.LayoutIndex, layoutCount - 1));
        var layoutNum = layoutIdx + 1;
        relsSb.Append($"<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout{layoutNum}.xml\"/>");
        foreach (var img in slide.Images)
        {
            relsSb.Append($"<Relationship Id=\"{img.RelId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/{img.RelId}.{img.Extension}\"/>");
        }
        foreach (var hlEntry in hlinkMap)
        {
            relsSb.Append($"<Relationship Id=\"{hlEntry.Key}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{EscXml(hlEntry.Value)}\" TargetMode=\"External\"/>");
        }
        foreach (var chart in slide.Charts)
        {
            relsSb.Append($"<Relationship Id=\"{chart.RelId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{chart.ChartNumber}.xml\"/>");
        }
        foreach (var vid in slide.Videos)
        {
            relsSb.Append($"<Relationship Id=\"{vid.RelId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/video\" Target=\"../media/{vid.RelId}.{vid.Extension}\"/>");
            // 视频缩略图关系
            if (vid.ThumbnailRelId.Length > 0)
                relsSb.Append($"<Relationship Id=\"{vid.ThumbnailRelId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/{vid.ThumbnailRelId}.{vid.ThumbnailExtension}\"/>");
        }
        if (slide.BackgroundImage != null)
        {
            relsSb.Append($"<Relationship Id=\"rBg{idx + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/bg{idx + 1}.{slide.BackgroundImage.Extension}\"/>");
        }
        relsSb.Append("</Relationships>");
        WriteEntry(za, $"ppt/slides/_rels/slide{idx + 1}.xml.rels", relsSb.ToString());

        // write image media
        foreach (var img in slide.Images)
        {
            using var entry = za.CreateEntry($"ppt/media/{img.RelId}.{img.Extension}").Open();
            entry.Write(img.Data, 0, img.Data.Length);
        }

        // write video media
        foreach (var vid in slide.Videos)
        {
            using (var entry = za.CreateEntry($"ppt/media/{vid.RelId}.{vid.Extension}").Open())
            {
                entry.Write(vid.Data, 0, vid.Data.Length);
            }
            // 写入缩略图（在视频流关闭后）
            if (vid.ThumbnailData != null && vid.ThumbnailData.Length > 0 && vid.ThumbnailRelId.Length > 0)
            {
                using var thumbEntry = za.CreateEntry($"ppt/media/{vid.ThumbnailRelId}.{vid.ThumbnailExtension}").Open();
                thumbEntry.Write(vid.ThumbnailData, 0, vid.ThumbnailData.Length);
            }
        }

        // write background image media
        if (slide.BackgroundImage != null)
        {
            using var entry = za.CreateEntry($"ppt/media/bg{idx + 1}.{slide.BackgroundImage.Extension}").Open();
            entry.Write(slide.BackgroundImage.Data, 0, slide.BackgroundImage.Data.Length);
        }

        // write chart XMLs
        foreach (var chart in slide.Charts)
        {
            WriteChartXml(za, chart);
        }
    }

    /// <summary>写入文档属性（docProps/core.xml 和 docProps/app.xml），S14-01/S14-02</summary>
    private void WriteDocProps(ZipArchive za)
    {
        var props = _documentProperties;
        if (props == null) return;
        var hasTitle = !props.Title.IsNullOrEmpty();
        var hasAuthor = !props.Author.IsNullOrEmpty();
        var hasSubject = !props.Subject.IsNullOrEmpty();
        var hasDesc = !props.Description.IsNullOrEmpty();
        if (!hasTitle && !hasAuthor && !hasSubject && !hasDesc) return;

        // core.xml
        var coreSb = new StringBuilder();
        coreSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        coreSb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"");
        coreSb.Append(" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"");
        coreSb.Append(" xmlns:dcterms=\"http://purl.org/dc/terms/\"");
        coreSb.Append(" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        if (hasTitle)
            coreSb.Append($"<dc:title>{EscXml(props.Title!)}</dc:title>");
        if (hasAuthor)
            coreSb.Append($"<dc:creator>{EscXml(props.Author!)}</dc:creator>");
        if (hasSubject)
            coreSb.Append($"<dc:subject>{EscXml(props.Subject!)}</dc:subject>");
        if (hasDesc)
            coreSb.Append($"<dc:description>{EscXml(props.Description!)}</dc:description>");
        var now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
        coreSb.Append($"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created>");
        coreSb.Append($"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified>");
        coreSb.Append("</cp:coreProperties>");
        WriteZipEntryText(za, "docProps/core.xml", coreSb.ToString());

        // app.xml
        var appSb = new StringBuilder();
        appSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        appSb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");
        var totalSlides = Slides.Count + _rawSlides.Count;
        appSb.Append($"<Slides>{totalSlides}</Slides>");
        if (!props.Author.IsNullOrEmpty())
            appSb.Append($"<Company>{EscXml(props.Author!)}</Company>");
        appSb.Append("</Properties>");
        WriteZipEntryText(za, "docProps/app.xml", appSb.ToString());
    }

    /// <summary>写入批注（ppt/comments/commentAuthors.xml 和各幻灯片 comments XML），S13-01</summary>
    private void WriteComments(ZipArchive za)
    {
        // 收集所有幻灯片的批注，确保每个幻灯片内 Index 从 1 开始
        var allCommentSlides = new List<Int32>();
        var allCommentItems = new List<PptComment>();
        var uniqueAuthors = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < Slides.Count; i++)
        {
            var comments = Slides[i].Comments;
            for (var ci = 0; ci < comments.Count; ci++)
            {
                var c = comments[ci];
                c.Index = ci + 1;
                allCommentSlides.Add(i);
                allCommentItems.Add(c);
                var name = c.Author ?? "unknown";
                var id = c.AuthorId ?? name;
                if (!uniqueAuthors.ContainsKey(name))
                    uniqueAuthors[name] = id;
            }
        }
        if (allCommentItems.Count == 0) return;

        // commentAuthors.xml — 全局唯一，记录所有批注作者
        var authSb = new StringBuilder();
        authSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        authSb.Append("<p:cmAuthorLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">");
        var authorIdMap = new Dictionary<String, Int32>(StringComparer.OrdinalIgnoreCase);
        var authId = 0;
        foreach (var kv in uniqueAuthors)
        {
            var name = kv.Key;
            var uid = kv.Value;
            authorIdMap[name] = authId;
            var initial = name.Length > 2 ? name.Substring(0, 2) : name;
            authSb.Append($"<p:cmAuthor id=\"{authId}\" name=\"{EscXml(name)}\" initials=\"{EscXml(initial)}\"");
            authSb.Append($" uid=\"{EscXml(uid)}\" lastIdx=\"{allCommentItems.Count}\"/>");
            authId++;
        }
        authSb.Append("</p:cmAuthorLst>");
        WriteZipEntryText(za, "ppt/comments/commentAuthors.xml", authSb.ToString());

        // 为每个有批注的幻灯片生成 comments{N}.xml — 按 slideIdx 分组
        var slideMap = new Dictionary<Int32, List<PptComment>>();
        for (var i = 0; i < allCommentSlides.Count; i++)
        {
            var idx = allCommentSlides[i];
            if (!slideMap.TryGetValue(idx, out var list))
                slideMap[idx] = list = [];
            list.Add(allCommentItems[i]);
        }
        foreach (var kv in slideMap)
        {
            var slideIdx = kv.Key;
            var slideCommentList = kv.Value;
            var csb = new StringBuilder();
            csb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            csb.Append("<p:cmLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">");
            foreach (var comment in slideCommentList)
            {
                authorIdMap.TryGetValue(comment.Author ?? "unknown", out var ai);
                var dateStr = (comment.Date ?? DateTime.UtcNow).ToString("yyyy-MM-ddTHH:mm:ssZ");
                csb.Append($"<p:cm authorId=\"{ai}\" dt=\"{dateStr}\" idx=\"{comment.Index}\">");
                csb.Append("<p:pos ");
                csb.Append($"x=\"{(Int32)(comment.X * SlideWidth)}\" ");
                csb.Append($"y=\"{(Int32)(comment.Y * SlideHeight)}\"/>");
                csb.Append($"<p:text>{EscXml(comment.Text ?? String.Empty)}</p:text>");
                csb.Append("</p:cm>");
            }
            csb.Append("</p:cmLst>");
            WriteZipEntryText(za, $"ppt/comments/comment{slideIdx + 1}.xml", csb.ToString());
        }
    }

    private void WriteChartXml(ZipArchive za, PptChart chart)
    {
        const String C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<c:chartSpace xmlns:c=\"{C}\" xmlns:a=\"{A}\">");
        sb.Append("<c:date1904 val=\"0\"/>");
        sb.Append("<c:chart>");
        if (chart.Title != null)
        {
            sb.Append("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>");
            sb.Append($"<a:p><a:r><a:t>{EscXml(chart.Title)}</a:t></a:r></a:p>");
            sb.Append("</c:rich></c:tx><c:overlay val=\"0\"/></c:title>");
        }
        sb.Append("<c:autoTitleDeleted val=\"0\"/>");
        sb.Append("<c:plotArea>");

        var chartElem = chart.ChartType switch
        {
            "line" => "lineChart",
            "pie" => "pieChart",
            "area" => "areaChart",
            "scatter" => "scatterChart",
            "bubble" => "bubbleChart",
            "radar" => "radarChart",
            "stock" => "stockChart",
            "surface" => "surfaceChart",
            _ => "barChart",
        };
        sb.Append($"<c:{chartElem}>");
        if (chart.ChartType == "bar")
            sb.Append("<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>");
        if (chart.ChartType == "scatter")
            sb.Append("<c:scatterStyle val=\"lineMarker\"/>");
        if (chart.ChartType == "bubble")
            sb.Append("<c:bubbleScale val=\"100\"/>");
        if (chart.ChartType == "radar")
            sb.Append("<c:radarStyle val=\"marker\"/>");

        var serColors = new[] { "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646" };
        for (var si = 0; si < chart.Series.Count; si++)
        {
            var ser = chart.Series[si];
            var color = serColors[si % serColors.Length];
            sb.Append("<c:ser>");
            sb.Append($"<c:idx val=\"{si}\"/><c:order val=\"{si}\"/>");
            sb.Append($"<c:tx><c:strRef><c:f/><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{EscXml(ser.Name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>");
            sb.Append($"<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr>");
            // categories — for scatter/bubble, skip and use per-series xVal
            var isXy = chart.ChartType is "scatter" or "bubble";
            if (!isXy && chart.Categories.Length > 0)
            {
                sb.Append("<c:cat><c:strRef><c:f/><c:strCache>");
                sb.Append($"<c:ptCount val=\"{chart.Categories.Length}\"/>");
                for (var ci = 0; ci < chart.Categories.Length; ci++)
                {
                    sb.Append($"<c:pt idx=\"{ci}\"><c:v>{EscXml(chart.Categories[ci])}</c:v></c:pt>");
                }
                sb.Append("</c:strCache></c:strRef></c:cat>");
            }
            // xVal for scatter/bubble
            if (isXy && ser.XValues is { Length: > 0 })
            {
                sb.Append("<c:xVal><c:numRef><c:f/><c:numCache>");
                sb.Append($"<c:ptCount val=\"{ser.XValues.Length}\"/>");
                for (var xi = 0; xi < ser.XValues.Length; xi++)
                    sb.Append($"<c:pt idx=\"{xi}\"><c:v>{ser.XValues[xi]}</c:v></c:pt>");
                sb.Append("</c:numCache></c:numRef></c:xVal>");
            }
            // values
            sb.Append("<c:val><c:numRef><c:f/><c:numCache>");
            sb.Append($"<c:ptCount val=\"{ser.Values.Length}\"/>");
            for (var vi = 0; vi < ser.Values.Length; vi++)
            {
                sb.Append($"<c:pt idx=\"{vi}\"><c:v>{ser.Values[vi]}</c:v></c:pt>");
            }
            sb.Append("</c:numCache></c:numRef></c:val>");
            sb.Append("</c:ser>");
        }
        if (chart.ChartType != "pie")
        {
            sb.Append("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
            sb.Append($"</c:{chartElem}>");
            // category axis
            sb.Append("<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>");
            // value axis
            sb.Append("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/>");
            if (chart.AxisMinValue.HasValue)
                sb.Append($"<c:min val=\"{chart.AxisMinValue.Value}\"/>");
            if (chart.AxisMaxValue.HasValue)
                sb.Append($"<c:max val=\"{chart.AxisMaxValue.Value}\"/>");
            sb.Append("</c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>");
        }
        else
        {
            sb.Append($"</c:{chartElem}>");
        }
        sb.Append("</c:plotArea>");
        sb.Append("<c:legend><c:legendPos val=\"b\"/></c:legend>");
        sb.Append("</c:chart></c:chartSpace>");
        WriteEntry(za, $"ppt/charts/chart{chart.ChartNumber}.xml", sb.ToString());
    }

    /// <summary>写入单个富文本 Run（供 Paragraphs 和 Runs 模式共用）</summary>
    private void WriteTextRun(StringBuilder sb, PptTextRun run, PptTextBox tb, ref Int32 hlinkGlobal, Dictionary<String, String> hlinkMap)
    {
        String? runHlRelId = null;
        if (run.HyperlinkUrl != null)
        {
            runHlRelId = $"rHlk{hlinkGlobal++}";
            hlinkMap[runHlRelId] = run.HyperlinkUrl;
        }
        var runSz = run.FontSize > 0 ? run.FontSize : tb.FontSize;
        var runFc = run.FontColor ?? tb.FontColor;
        var runLatinFn = run.LatinFontName ?? tb.LatinFontName;
        var runEaFn = run.EastAsianFontName ?? tb.EastAsianFontName;
        var runCsFn = run.ComplexScriptFontName ?? tb.ComplexScriptFontName;
        var runSymFn = run.SymbolFontName ?? tb.SymbolFontName;
        if (runLatinFn == null && runEaFn == null)
        {
            var fn = run.FontName ?? tb.FontName;
            runLatinFn = fn;
            runEaFn = fn;
        }
        sb.Append("<a:r>");
        var baseline = run.Superscript ? " baseline=\"30000\"" : run.Subscript ? " baseline=\"-25000\"" : "";
        sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"{runSz * 100}\"{(run.Bold ? " b=\"1\"" : "")}{(run.Italic ? " i=\"1\"" : "")}{(run.Underline ? " u=\"sng\"" : "")}{baseline} dirty=\"0\">");
        if (run.GradFillColors?.Length >= 2)
            WriteGradFill(sb, run.GradFillColors, run.GradAngle);
        else
            WriteFontColor(sb, runFc);
        if (runLatinFn != null)
            sb.Append($"<a:latin typeface=\"{EscXml(runLatinFn)}\"/>");
        if (runEaFn != null)
            sb.Append($"<a:ea typeface=\"{EscXml(runEaFn)}\"/>");
        if (runCsFn != null)
            sb.Append($"<a:cs typeface=\"{EscXml(runCsFn)}\"/>");
        if (runSymFn != null)
            sb.Append($"<a:sym typeface=\"{EscXml(runSymFn)}\"/>");
        if (runHlRelId != null)
            sb.Append($"<a:hlinkClick r:id=\"{runHlRelId}\"/>");
        sb.Append("</a:rPr>");
        sb.Append($"<a:t>{EscXml(run.Text)}</a:t>");
        sb.Append("</a:r>");
    }

    /// <summary>写入单行非 Run 模式文本（向后兼容）</summary>
    private void WriteSingleLineTextRun(StringBuilder sb, PptTextBox tb, String? hlRelId)
    {
        sb.Append("<a:r>");
        sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"{tb.FontSize * 100}\"{(tb.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
        WriteFontColor(sb, tb.FontColor);
        var tbLatinFn = tb.LatinFontName;
        var tbEaFn = tb.EastAsianFontName;
        var tbCsFn = tb.ComplexScriptFontName;
        var tbSymFn = tb.SymbolFontName;
        if (tbLatinFn == null && tbEaFn == null)
        {
            var fn = tb.FontName;
            tbLatinFn = fn;
            tbEaFn = fn;
        }
        if (tbLatinFn != null)
            sb.Append($"<a:latin typeface=\"{EscXml(tbLatinFn)}\"/>");
        if (tbEaFn != null)
            sb.Append($"<a:ea typeface=\"{EscXml(tbEaFn)}\"/>");
        if (tbCsFn != null)
            sb.Append($"<a:cs typeface=\"{EscXml(tbCsFn)}\"/>");
        if (tbSymFn != null)
            sb.Append($"<a:sym typeface=\"{EscXml(tbSymFn)}\"/>");
        if (hlRelId != null)
            sb.Append($"<a:hlinkClick r:id=\"{hlRelId}\"/>");
        sb.Append("</a:rPr>");
        sb.Append($"<a:t>{EscXml(tb.Text)}</a:t>");
        sb.Append("</a:r>");
    }

    private static void BuildPptTableXml(StringBuilder sb, PptTable tbl, ref Int32 shapeId)
    {
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        sb.Append($"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{shapeId++}\" name=\"Table\"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>");
        sb.Append($"<p:xfrm><a:off x=\"{tbl.Left}\" y=\"{tbl.Top}\"/><a:ext cx=\"{tbl.Width}\" cy=\"{tbl.Height}\"/></p:xfrm>");
        sb.Append($"<a:graphic xmlns:a=\"{A}\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">");
        var tableStyleId = tbl.TableStyleGuid ?? "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
        sb.Append($"<a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"><a:tableStyleId>{tableStyleId}</a:tableStyleId></a:tblPr>");
        // columns
        var colCount = tbl.Rows.Count > 0 ? tbl.Rows[0].Length : 1;
        var autoColW = colCount > 0 ? tbl.Width / colCount : tbl.Width;
        sb.Append("<a:tblGrid>");
        for (var c = 0; c < colCount; c++)
        {
            var cw = tbl.ColWidths.Length > c ? tbl.ColWidths[c] : autoColW;
            sb.Append($"<a:gridCol w=\"{cw}\"/>");
        }
        sb.Append("</a:tblGrid>");
        for (var ri = 0; ri < tbl.Rows.Count; ri++)
        {
            var row = tbl.Rows[ri];
            var isHeaderRow = ri == 0 && tbl.FirstRowHeader;
            sb.Append("<a:tr h=\"370840\">");
            for (var ci = 0; ci < row.Length; ci++)
            {
                tbl.CellStyles.TryGetValue((ri, ci), out var cs);
                tbl.MergedCells.TryGetValue((ri, ci), out var merge);
                tbl.CellBorders.TryGetValue((ri, ci), out var border);
                var isBold = isHeaderRow || (cs?.Bold ?? false);
                var cellSz = (cs?.FontSize ?? 0) > 0 ? cs!.FontSize : 0;
                var cellFc = cs?.FontColor;
                var cellBg = cs?.BackgroundColor;
                sb.Append("<a:tc");
                if (merge.ColSpan > 1) sb.Append($" gridSpan=\"{merge.ColSpan}\"");
                if (merge.RowSpan > 1) sb.Append($" rowSpan=\"{merge.RowSpan}\"");
                // 被合并的后续单元格：RowSpan 向上合并的隐藏单元格（vmMerge=1）
                var vmMerge = false;
                for (var pr = 0; pr < ri; pr++)
                {
                    tbl.MergedCells.TryGetValue((pr, ci), out var prevMerge);
                    if (prevMerge.RowSpan > 1 && ri < pr + prevMerge.RowSpan)
                    {
                        vmMerge = true;
                        break;
                    }
                }
                if (vmMerge) sb.Append(" vMerge=\"1\"");
                sb.Append("><a:txBody><a:bodyPr/><a:lstStyle/>");
                sb.Append("<a:p><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\"{(isBold ? " b=\"1\"" : "")}{(cellSz > 0 ? $" sz=\"{cellSz * 100}\"" : "")} dirty=\"0\">");
                if (cellFc != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{cellFc.TrimStart('#')}\"/></a:solidFill>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(row[ci])}</a:t>");
                sb.Append("</a:r></a:p></a:txBody>");
                if (cellBg != null || border != null || vmMerge)
                {
                    sb.Append("<a:tcPr");
                    if (vmMerge) sb.Append(" vMerge=\"1\"");
                    sb.Append('>');
                    if (cellBg != null)
                        sb.Append($"<a:solidFill><a:srgbClr val=\"{cellBg.TrimStart('#')}\"/></a:solidFill>");
                    if (border != null)
                    {
                        if (border.LeftColor != null)
                            sb.Append($"<a:lnL w=\"{border.LeftWidth}\"><a:solidFill><a:srgbClr val=\"{border.LeftColor.TrimStart('#')}\"/></a:solidFill></a:lnL>");
                        if (border.RightColor != null)
                            sb.Append($"<a:lnR w=\"{border.RightWidth}\"><a:solidFill><a:srgbClr val=\"{border.RightColor.TrimStart('#')}\"/></a:solidFill></a:lnR>");
                        if (border.TopColor != null)
                            sb.Append($"<a:lnT w=\"{border.TopWidth}\"><a:solidFill><a:srgbClr val=\"{border.TopColor.TrimStart('#')}\"/></a:solidFill></a:lnT>");
                        if (border.BottomColor != null)
                            sb.Append($"<a:lnB w=\"{border.BottomWidth}\"><a:solidFill><a:srgbClr val=\"{border.BottomColor.TrimStart('#')}\"/></a:solidFill></a:lnB>");
                    }
                    sb.Append("</a:tcPr>");
                }
                else
                    sb.Append("<a:tcPr/>");
                sb.Append("</a:tc>");
            }
            sb.Append("</a:tr>");
        }
        sb.Append("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
    }

    private void WriteSlideLayout(ZipArchive za)
    {
        if (_layoutContents.Count > 0)
        {
            for (var i = 0; i < _layoutContents.Count; i++)
            {
                var c = _layoutContents[i];
                WriteZipEntryText(za, $"ppt/slideLayouts/slideLayout{i + 1}.xml", c.Xml);
                if (c.RelsXml.Length > 0)
                    WriteZipEntryText(za, $"ppt/slideLayouts/_rels/slideLayout{i + 1}.xml.rels", c.RelsXml);
            }
            return;
        }
        if (_progMasters.Count > 0)
        {
            WriteProgLayouts(za);
            return;
        }
        // 无模板时写出默认空白版式（向后兼容）
        WriteEntry(za, "ppt/slideLayouts/slideLayout1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" type=\"blank\" preserve=\"1\">" +
            "<p:cSld name=\"Blank\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
            "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>" +
            "</p:spTree></p:cSld></p:sldLayout>");
        WriteEntry(za, "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/>" +
            "</Relationships>");
    }

    private void WriteSlideMaster(ZipArchive za)
    {
        if (_masterContents.Count > 0)
        {
            for (var i = 0; i < _masterContents.Count; i++)
            {
                var c = _masterContents[i];
                WriteZipEntryText(za, $"ppt/slideMasters/slideMaster{i + 1}.xml", c.Xml);
                if (c.RelsXml.Length > 0)
                    WriteZipEntryText(za, $"ppt/slideMasters/_rels/slideMaster{i + 1}.xml.rels", c.RelsXml);
            }
            return;
        }
        if (_progMasters.Count > 0)
        {
            WriteProgMasters(za);
            return;
        }
        // 无模板时写出默认空白母版（向后兼容）
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        WriteEntry(za, "ppt/slideMasters/slideMaster1.xml",
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<p:sldMaster xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\">" +
            "<p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>" +
            "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
            "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>" +
            "</p:spTree></p:cSld>" +
            "<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>" +
            "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rLayout1\"/></p:sldLayoutIdLst>" +
            "</p:sldMaster>");
        WriteEntry(za, "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>" +
            "<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
            "</Relationships>");
    }

    private void WriteTheme(ZipArchive za)
    {
        // 优先使用原始字节（避免 StreamReader/StreamWriter 导致的换行符归一化差异）
        if (_templateThemeBytes != null)
        {
            using var es = za.CreateEntry("ppt/theme/theme1.xml").Open();
            es.Write(_templateThemeBytes, 0, _templateThemeBytes.Length);
            return;
        }
        if (_templateThemeXml != null)
        {
            WriteZipEntryText(za, "ppt/theme/theme1.xml", _templateThemeXml);
            return;
        }
        WriteEntry(za, "ppt/theme/theme1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
            "<a:themeElements><a:clrScheme name=\"Office\">" +
            "<a:dk1><a:sysClr lastClr=\"000000\" val=\"windowText\"/></a:dk1>" +
            "<a:lt1><a:sysClr lastClr=\"FFFFFF\" val=\"window\"/></a:lt1>" +
            "<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>" +
            "<a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>" +
            AccentColorXml(0) + AccentColorXml(1) + AccentColorXml(2) + AccentColorXml(3) + AccentColorXml(4) + AccentColorXml(5) +
            "<a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink>" +
            "<a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>" +
            "</a:clrScheme>" +
            "<a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
            "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
            "<a:fmtScheme name=\"Office\"><a:fillStyleLst><a:noFill/><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:noFill/></a:fillStyleLst>" +
            "<a:lnStyleLst><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>" +
            "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
            "<a:bgFillStyleLst><a:noFill/><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:noFill/></a:bgFillStyleLst>" +
            "</a:fmtScheme></a:themeElements></a:theme>");
    }

    private void WriteInfraMedia(ZipArchive za)
    {
        foreach (var kv in _infraMedia)
        {
            var entry = za.CreateEntry($"ppt/media/{kv.Key}", CompressionLevel.Fastest);
            using var es = entry.Open();
            es.Write(kv.Value, 0, kv.Value.Length);
        }
    }

    /// <summary>根据 XML 条目路径计算对应的 rels 条目路径</summary>
    private static String GetRelsEntryPath(String xmlPath)
    {
        var dir = (Path.GetDirectoryName(xmlPath) ?? String.Empty).Replace('\\', '/');
        return $"{dir}/_rels/{Path.GetFileName(xmlPath)}.rels";
    }

    /// <summary>从版式 XML 提取显示名称：优先取 cSld/@name，其次取 root/@type</summary>
    private static String ExtractLayoutDisplayName(String layoutXml)
    {
        try
        {
            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(layoutXml);
            var cSld = doc.DocumentElement?.SelectSingleNode("//*[local-name()='cSld']") as System.Xml.XmlElement;
            if (cSld != null)
            {
                var name = cSld.GetAttribute("name");
                if (name.Length > 0) return name;
            }
            var typeAttr = doc.DocumentElement?.GetAttribute("type") ?? String.Empty;
            if (typeAttr.Length > 0) return typeAttr;
        }
        catch
        {
            // 版式 XML 可能格式不符预期或包含非法字符，提取名称失败时静默返回空字符串，
            // 不影响 pptx 文件的正常保存和 PowerPoint 打开
        }
        return String.Empty;
    }

    /// <summary>写出编程式创建的母版 XML</summary>
    private void WriteProgMasters(ZipArchive za)
    {
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        for (var mi = 0; mi < _progMasters.Count; mi++)
        {
            var master = _progMasters[mi];
            var masterId = mi + 1;
            var layoutOffset = GetProgLayoutOffset(mi);

            // 母版 XML
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<p:sldMaster xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\">");
            sb.Append("<p:cSld>");
            // 背景
            if (master.BackgroundColor != null)
            {
                sb.Append("<p:bg><p:bgPr>");
                sb.Append($"<a:solidFill><a:srgbClr val=\"{master.BackgroundColor.TrimStart('#')}\"/></a:solidFill>");
                sb.Append("<a:effectLst/></p:bgPr></p:bg>");
            }
            else
            {
                sb.Append("<p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>");
            }
            // spTree（含母版形状）
            sb.Append("<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
            sb.Append("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");
            var mShapeId = 2;
            foreach (var sp in master.Shapes)
            {
                WriteShapeXml(sb, sp, ref mShapeId, A);
            }
            sb.Append("</p:spTree></p:cSld>");
            // txStyles
            sb.Append("<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>");
            // 版式 ID 列表
            sb.Append("<p:sldLayoutIdLst>");
            for (var li = 0; li < master.Layouts.Count; li++)
            {
                var layoutGlobalIdx = layoutOffset + li + 1;
                sb.Append($"<p:sldLayoutId id=\"{2147483649 + layoutGlobalIdx}\" r:id=\"rLayout{layoutGlobalIdx}\"/>");
            }
            sb.Append("</p:sldLayoutIdLst>");
            sb.Append("</p:sldMaster>");
            WriteZipEntryText(za, $"ppt/slideMasters/slideMaster{masterId}.xml", sb.ToString());

            // 母版 rels
            var relsSb = new StringBuilder();
            relsSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            relsSb.Append("<Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>");
            for (var li = 0; li < master.Layouts.Count; li++)
            {
                var layoutGlobalIdx = layoutOffset + li + 1;
                relsSb.Append($"<Relationship Id=\"rLayout{layoutGlobalIdx}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout{layoutGlobalIdx}.xml\"/>");
            }
            relsSb.Append("</Relationships>");
            WriteZipEntryText(za, $"ppt/slideMasters/_rels/slideMaster{masterId}.xml.rels", relsSb.ToString());
        }
    }

    /// <summary>计算指定母版之前所有母版的版式累计数量</summary>
    private Int32 GetProgLayoutOffset(Int32 masterIndex)
    {
        var offset = 0;
        for (var i = 0; i < masterIndex; i++)
        {
            offset += _progMasters[i].Layouts.Count;
        }
        return offset;
    }

    /// <summary>写出编程式创建的版式 XML（由 WriteSlideLayout 调用）</summary>
    private void WriteProgLayouts(ZipArchive za)
    {
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var globalIdx = 0;
        for (var mi = 0; mi < _progMasters.Count; mi++)
        {
            var master = _progMasters[mi];
            for (var li = 0; li < master.Layouts.Count; li++)
            {
                globalIdx++;
                var layout = master.Layouts[li];
                var sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append($"<p:sldLayout xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\"");
                sb.Append($" type=\"{EscXml(layout.LayoutType)}\" preserve=\"1\">");
                sb.Append($"<p:cSld name=\"{EscXml(layout.Name)}\">");
                sb.Append("<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
                sb.Append("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");
                var shapeId = 2;
                foreach (var sp in layout.Shapes)
                {
                    WriteShapeXml(sb, sp, ref shapeId, A);
                }
                foreach (var tb in layout.TextBoxes)
                {
                    WriteTextBoxXml(sb, tb, ref shapeId, A);
                }
                sb.Append("</p:spTree></p:cSld></p:sldLayout>");
                WriteZipEntryText(za, $"ppt/slideLayouts/slideLayout{globalIdx}.xml", sb.ToString());

                // 版式 rels
                var masterId = mi + 1;
                var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    $"<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster{masterId}.xml\"/>" +
                    "</Relationships>";
                WriteZipEntryText(za, $"ppt/slideLayouts/_rels/slideLayout{globalIdx}.xml.rels", rels);
            }
        }
    }

    /// <summary>生成 p:sp 形状 XML 片段（复用于母版、版式、幻灯片）</summary>
    private static void WriteShapeXml(StringBuilder sb, PptShape sp, ref Int32 shapeId, String aNs)
    {
        sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"Shape\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>");
        sb.Append("<p:spPr>");
        sb.Append($"<a:xfrm><a:off x=\"{sp.Left}\" y=\"{sp.Top}\"/><a:ext cx=\"{sp.Width}\" cy=\"{sp.Height}\"/></a:xfrm>");
        if (!String.IsNullOrEmpty(sp.ShapeType))
            sb.Append($"<a:prstGeom prst=\"{sp.ShapeType}\"><a:avLst/></a:prstGeom>");
        if (sp.FillColor != null)
            sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FillColor.TrimStart('#')}\"/></a:solidFill>");
        else
            sb.Append("<a:noFill/>");
        if (sp.LineColor != null)
            sb.Append($"<a:ln w=\"{sp.LineWidth}\"><a:solidFill><a:srgbClr val=\"{sp.LineColor.TrimStart('#')}\"/></a:solidFill></a:ln>");
        sb.Append("</p:spPr>");
        if (!String.IsNullOrEmpty(sp.Text))
        {
            sb.Append("<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>");
            sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{sp.FontSize * 100}\"{(sp.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
            WriteFontColor(sb, sp.FontColor);
            WriteFontElements(sb, sp.LatinFontName, sp.EastAsianFontName, sp.ComplexScriptFontName, sp.SymbolFontName);
            sb.Append("</a:rPr>");
            sb.Append($"<a:t>{EscXml(sp.Text)}</a:t>");
            sb.Append("</a:r></a:p></p:txBody>");
        }
        sb.Append("</p:sp>");
    }

    /// <summary>生成 p:sp 文本框 XML 片段（复用于版式、幻灯片）</summary>
    private static void WriteTextBoxXml(StringBuilder sb, PptTextBox tb, ref Int32 shapeId, String aNs)
    {
        sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"TextBox\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>");
        sb.Append("<p:spPr>");
        sb.Append($"<a:xfrm><a:off x=\"{tb.Left}\" y=\"{tb.Top}\"/><a:ext cx=\"{tb.Width}\" cy=\"{tb.Height}\"/></a:xfrm>");
        sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
        if (tb.BackgroundColor != null)
            sb.Append($"<a:solidFill><a:srgbClr val=\"{tb.BackgroundColor.TrimStart('#')}\"/></a:solidFill>");
        else
            sb.Append("<a:noFill/>");
        sb.Append("</p:spPr>");
        sb.Append("<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"><a:normAutofit/></a:bodyPr><a:lstStyle/>");
        sb.Append($"<a:p><a:pPr algn=\"{tb.Alignment}\"/><a:r>");
        sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"{tb.FontSize * 100}\"{(tb.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
        WriteFontColor(sb, tb.FontColor);
        WriteFontElements(sb, tb.LatinFontName, tb.EastAsianFontName, tb.ComplexScriptFontName, tb.SymbolFontName);
        sb.Append("</a:rPr>");
        sb.Append($"<a:t>{EscXml(tb.Text)}</a:t>");
        sb.Append("</a:r></a:p></p:txBody></p:sp>");
    }

    /// <summary>生成单个强调色 XML 片段</summary>
    private String AccentColorXml(Int32 index)
    {
        var color = index < AccentColors.Length ? AccentColors[index] : "4F81BD";
        return $"<a:accent{index + 1}><a:srgbClr val=\"{color}\"/></a:accent{index + 1}>";
    }

    private static String EscXml(String s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
         .Replace("\"", "&quot;").Replace("'", "&apos;")
         .Replace("\r", "&#xD;");

    /// <summary>写出渐变填充</summary>
    private static void WriteGradFill(StringBuilder sb, String[] colors, Int32 angle)
    {
        sb.Append("<a:gradFill><a:gsLst>");
        var step = colors.Length > 1 ? 100000 / (colors.Length - 1) : 0;
        for (var i = 0; i < colors.Length; i++)
        {
            var pos = i == colors.Length - 1 ? 100000 : step * i;
            sb.Append($"<a:gs pos=\"{pos}\"><a:srgbClr val=\"{colors[i].TrimStart('#')}\"/></a:gs>");
        }
        sb.Append($"</a:gsLst><a:lin ang=\"{angle}\" scaled=\"0\"/></a:gradFill>");
    }

    /// <summary>写出字体颜色（srgbClr 或 schemeClr）</summary>
    private static void WriteFontColor(StringBuilder sb, String? fontColor)
    {
        if (fontColor == null) return;
        if (fontColor.StartsWith("scheme:", StringComparison.OrdinalIgnoreCase))
        {
            var val = fontColor.Substring(7);
            sb.Append($"<a:solidFill><a:schemeClr val=\"{val}\"/></a:solidFill>");
        }
        else
        {
            sb.Append($"<a:solidFill><a:srgbClr val=\"{fontColor.TrimStart('#')}\"/></a:solidFill>");
        }
    }

    /// <summary>写出字体元素（latin/ea/cs/sym）</summary>
    /// <param name="sb">输出构建器</param>
    /// <param name="latinFn">拉丁/西文字体名称</param>
    /// <param name="eaFn">东亚/中文字体名称</param>
    /// <param name="csFn">复杂脚本字体名称</param>
    /// <param name="symFn">符号字体名称</param>
    private static void WriteFontElements(StringBuilder sb, String? latinFn, String? eaFn, String? csFn, String? symFn)
    {
        if (latinFn != null) sb.Append($"<a:latin typeface=\"{EscXml(latinFn)}\"/>");
        if (eaFn != null) sb.Append($"<a:ea typeface=\"{EscXml(eaFn)}\"/>");
        if (csFn != null) sb.Append($"<a:cs typeface=\"{EscXml(csFn)}\"/>");
        if (symFn != null) sb.Append($"<a:sym typeface=\"{EscXml(symFn)}\"/>");
    }
    #endregion
}
