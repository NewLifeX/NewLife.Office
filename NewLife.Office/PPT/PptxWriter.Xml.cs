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
        WriteTheme(za);
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
        sb.Append("<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>");
        sb.Append("<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>");
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
                    var ct = img.Extension is "jpg" or "jpeg" ? "image/jpeg" : "image/png";
                    sb.Append($"<Default Extension=\"{img.Extension}\" ContentType=\"{ct}\"/>");
                }
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
        sb.Append($"<p:presentation xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\" saveSubsetFonts=\"1\">");
        sb.Append($"<p:sldSz cx=\"{SlideWidth}\" cy=\"{SlideHeight}\"/>");
        sb.Append("<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rMaster1\"/></p:sldMasterIdLst>");
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
        sb.Append("<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>");
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
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<p:sld xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\">");

        // background
        if (slide.BackgroundColor != null)
        {
            sb.Append("<p:bg><p:bgPr>");
            sb.Append($"<a:solidFill><a:srgbClr val=\"{slide.BackgroundColor.TrimStart('#')}\"/></a:solidFill>");
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
            sb.Append($"<a:p><a:pPr algn=\"{tb.Alignment}\"/>");
            if (tb.Runs.Count > 0)
            {
                foreach (var run in tb.Runs)
                {
                    String? runHlRelId = null;
                    if (run.HyperlinkUrl != null)
                    {
                        runHlRelId = $"rHlk{_hlinkGlobal++}";
                        hlinkMap[runHlRelId] = run.HyperlinkUrl;
                    }
                    var runSz = run.FontSize > 0 ? run.FontSize : tb.FontSize;
                    var runFc = run.FontColor ?? tb.FontColor;
                    sb.Append("<a:r>");
                    sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"{runSz * 100}\"{(run.Bold ? " b=\"1\"" : "")}{(run.Italic ? " i=\"1\"" : "")} dirty=\"0\">");
                    if (runFc != null)
                        sb.Append($"<a:solidFill><a:srgbClr val=\"{runFc.TrimStart('#')}\"/></a:solidFill>");
                    if (runHlRelId != null)
                        sb.Append($"<a:hlinkClick r:id=\"{runHlRelId}\"/>");
                    sb.Append("</a:rPr>");
                    sb.Append($"<a:t>{EscXml(run.Text)}</a:t>");
                    sb.Append("</a:r>");
                }
            }
            else
            {
                sb.Append("<a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"{tb.FontSize * 100}\"{(tb.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                if (tb.FontColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{tb.FontColor.TrimStart('#')}\"/></a:solidFill>");
                if (hlRelId != null)
                    sb.Append($"<a:hlinkClick r:id=\"{hlRelId}\"/>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(tb.Text)}</a:t>");
                sb.Append("</a:r>");
            }
            sb.Append("</a:p></p:txBody></p:sp>");
        }

        // shapes（基本图形）
        foreach (var sp in slide.Shapes)
        {
            sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"Shape\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{sp.Left}\" y=\"{sp.Top}\"/><a:ext cx=\"{sp.Width}\" cy=\"{sp.Height}\"/></a:xfrm>");
            sb.Append($"<a:prstGeom prst=\"{sp.ShapeType}\"><a:avLst/></a:prstGeom>");
            if (sp.FillColor != null)
                sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FillColor.TrimStart('#')}\"/></a:solidFill>");
            else
                sb.Append("<a:noFill/>");
            if (sp.LineColor != null)
                sb.Append($"<a:ln w=\"{sp.LineWidth}\"><a:solidFill><a:srgbClr val=\"{sp.LineColor.TrimStart('#')}\"/></a:solidFill></a:ln>");
            else
                sb.Append("<a:ln><a:noFill/></a:ln>");
            sb.Append("</p:spPr>");
            if (sp.Text != null)
            {
                sb.Append("<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{sp.FontSize * 100}\"{(sp.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                if (sp.FontColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FontColor.TrimStart('#')}\"/></a:solidFill>");
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
            sb.Append($"<a:blip r:embed=\"{img.RelId}\"/>");
            sb.Append("<a:stretch><a:fillRect/></a:stretch></p:blipFill>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{img.Left}\" y=\"{img.Top}\"/><a:ext cx=\"{img.Width}\" cy=\"{img.Height}\"/></a:xfrm>");
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
                    sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{sp.FontSize * 100}\" dirty=\"0\"/>");
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
                if (tb.FontColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{tb.FontColor.TrimStart('#')}\"/></a:solidFill>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(tb.Text)}</a:t>");
                sb.Append("</a:r></a:p></p:txBody></p:sp>");
            }
            sb.Append("</p:grpSp>");
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
                _ => "<p:fade/>",
            });
            sb.Append("</p:transition>");
        }

        sb.Append("</p:sld>");
        WriteEntry(za, $"ppt/slides/slide{idx + 1}.xml", sb.ToString());

        // slide rels
        var relsSb = new StringBuilder();
        relsSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        relsSb.Append("<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>");
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
        relsSb.Append("</Relationships>");
        WriteEntry(za, $"ppt/slides/_rels/slide{idx + 1}.xml.rels", relsSb.ToString());

        // write image media
        foreach (var img in slide.Images)
        {
            using var entry = za.CreateEntry($"ppt/media/{img.RelId}.{img.Extension}").Open();
            entry.Write(img.Data, 0, img.Data.Length);
        }

        // write chart XMLs
        foreach (var chart in slide.Charts)
        {
            WriteChartXml(za, chart);
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
            _ => "barChart",
        };
        sb.Append($"<c:{chartElem}>");
        if (chart.ChartType == "bar")
            sb.Append("<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>");

        var serColors = new[] { "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646" };
        for (var si = 0; si < chart.Series.Count; si++)
        {
            var ser = chart.Series[si];
            var color = serColors[si % serColors.Length];
            sb.Append("<c:ser>");
            sb.Append($"<c:idx val=\"{si}\"/><c:order val=\"{si}\"/>");
            sb.Append($"<c:tx><c:strRef><c:f/><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{EscXml(ser.Name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>");
            sb.Append($"<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr>");
            // categories
            if (chart.Categories.Length > 0)
            {
                sb.Append("<c:cat><c:strRef><c:f/><c:strCache>");
                sb.Append($"<c:ptCount val=\"{chart.Categories.Length}\"/>");
                for (var ci = 0; ci < chart.Categories.Length; ci++)
                {
                    sb.Append($"<c:pt idx=\"{ci}\"><c:v>{EscXml(chart.Categories[ci])}</c:v></c:pt>");
                }
                sb.Append("</c:strCache></c:strRef></c:cat>");
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
            sb.Append("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>");
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

    private static void BuildPptTableXml(StringBuilder sb, PptTable tbl, ref Int32 shapeId)
    {
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        sb.Append($"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{shapeId++}\" name=\"Table\"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>");
        sb.Append($"<p:xfrm><a:off x=\"{tbl.Left}\" y=\"{tbl.Top}\"/><a:ext cx=\"{tbl.Width}\" cy=\"{tbl.Height}\"/></p:xfrm>");
        sb.Append($"<a:graphic xmlns:a=\"{A}\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">");
        sb.Append("<a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>");
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
                var isBold = isHeaderRow || (cs?.Bold ?? false);
                var cellSz = (cs?.FontSize ?? 0) > 0 ? cs!.FontSize : 0;
                var cellFc = cs?.FontColor;
                var cellBg = cs?.BackgroundColor;
                sb.Append("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>");
                sb.Append("<a:p><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\"{(isBold ? " b=\"1\"" : "")}{(cellSz > 0 ? $" sz=\"{cellSz * 100}\"" : "")} dirty=\"0\">");
                if (cellFc != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{cellFc.TrimStart('#')}\"/></a:solidFill>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(row[ci])}</a:t>");
                sb.Append("</a:r></a:p></a:txBody>");
                if (cellBg != null)
                    sb.Append($"<a:tcPr><a:solidFill><a:srgbClr val=\"{cellBg.TrimStart('#')}\"/></a:solidFill></a:tcPr>");
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

    private void WriteTheme(ZipArchive za) =>
        WriteEntry(za, "ppt/theme/theme1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
            "<a:themeElements><a:clrScheme name=\"Office\">" +
            "<a:dk1><a:sysClr lastClr=\"000000\" val=\"windowText\"/></a:dk1>" +
            "<a:lt1><a:sysClr lastClr=\"FFFFFF\" val=\"window\"/></a:lt1>" +
            "<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>" +
            "<a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>" +
            $"<a:accent1><a:srgbClr val=\"{AccentColors[0]}\"/></a:accent1>" +
            $"<a:accent2><a:srgbClr val=\"{AccentColors[1]}\"/></a:accent2>" +
            $"<a:accent3><a:srgbClr val=\"{AccentColors[2]}\"/></a:accent3>" +
            $"<a:accent4><a:srgbClr val=\"{AccentColors[3]}\"/></a:accent4>" +
            $"<a:accent5><a:srgbClr val=\"{AccentColors[4]}\"/></a:accent5>" +
            $"<a:accent6><a:srgbClr val=\"{AccentColors[5]}\"/></a:accent6>" +
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

    private static String EscXml(String s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
         .Replace("\"", "&quot;").Replace("'", "&apos;");
    #endregion
}
