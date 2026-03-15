using System.Data;
using System.IO.Compression;
using System.Security;

namespace NewLife.Office;

partial class ExcelWriter
{
    #region 样式管理
    /// <summary>根据用户样式和数字格式，查找或创建 XF 条目并返回索引</summary>
    private Int32 GetOrCreateXf(CellStyle cs, Int32 numFmtId)
    {
        // 找或创建字体
        var font = new FontEntry(cs.FontName, cs.FontSize, cs.Bold, cs.Italic, cs.Underline, cs.FontColor);
        var fontId = FindOrAdd(_fonts, font);

        // 找或创建填充
        var fillId = 0;
        if (!cs.BackgroundColor.IsNullOrEmpty())
        {
            var fill = new FillEntry(cs.BackgroundColor, "solid");
            fillId = FindOrAdd(_fills, fill);
        }

        // 找或创建边框
        var borderId = 0;
        if (cs.Border != CellBorderStyle.None)
        {
            var border = new BorderEntry(cs.Border, cs.BorderColor);
            borderId = FindOrAdd(_borders, border);
        }

        // 复合键去重
        var key = $"{numFmtId}-{fontId}-{fillId}-{borderId}-{(Int32)cs.HAlign}-{(Int32)cs.VAlign}-{(cs.WrapText ? 1 : 0)}";
        if (_xfCache.TryGetValue(key, out var idx)) return idx;

        var xf = new XfEntry(numFmtId, fontId, fillId, borderId, cs.HAlign, cs.VAlign, cs.WrapText);
        idx = _xfEntries.Count;
        _xfEntries.Add(xf);
        _xfCache[key] = idx;
        return idx;
    }

    /// <summary>获取或创建自定义数字格式</summary>
    private Int32 GetOrCreateNumFmt(String formatCode)
    {
        if (_numFmtMap.TryGetValue(formatCode, out var id)) return id;
        id = _nextNumFmtId++;
        _numFmtMap[formatCode] = id;
        return id;
    }

    private static Int32 FindOrAdd<T>(List<T> list, T item) where T : notnull
    {
        for (var i = 0; i < list.Count; i++)
        {
            if (list[i].Equals(item)) return i;
        }
        list.Add(item);
        return list.Count - 1;
    }

    /// <summary>解析单元格引用（如 "A1"）返回 (行0基, 列0基)</summary>
    private static (Int32 Row, Int32 Col) ParseCellRef(String cellRef)
    {
        var colLen = 0;
        for (var i = 0; i < cellRef.Length; i++)
        {
            var ch = cellRef[i];
            if (ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z') colLen++;
            else break;
        }

        var colIndex = 0;
        for (var i = 0; i < colLen; i++)
        {
            var ch = cellRef[i];
            if (ch is >= 'a' and <= 'z') ch = (Char)(ch - 'a' + 'A');
            colIndex = colIndex * 26 + (ch - 'A' + 1);
        }
        colIndex--; // 转 0 基

        var rowStr = cellRef.Substring(colLen);
        var rowIndex = Int32.Parse(rowStr) - 1; // 转 0 基

        return (rowIndex, colIndex);
    }

    /// <summary>生成单元格引用（如 "A1"），行列均为 0 基</summary>
    private static String MakeCellRef(Int32 row, Int32 col) => GetColumnName(col) + (row + 1);

    /// <summary>获取边框 OOXML 样式名</summary>
    private static String GetBorderStyleName(CellBorderStyle style) => style switch
    {
        CellBorderStyle.Thin => "thin",
        CellBorderStyle.Medium => "medium",
        CellBorderStyle.Thick => "thick",
        CellBorderStyle.Dashed => "dashed",
        CellBorderStyle.Dotted => "dotted",
        CellBorderStyle.DoubleLine => "double",
        _ => "thin",
    };
    #endregion

    #region 保存
    /// <summary>保存到文件或目标流</summary>
    public void Save()
    {
        // 若未写任何 sheet，创建一个空的默认工作表，避免生成非法 workbook
        if (_sheetNames.Count == 0) EnsureSheet(SheetName);

        var target = Stream;
        if (target == null)
        {
            if (FileName.IsNullOrEmpty()) throw new InvalidOperationException("未指定输出位置");

            var file = FileName.EnsureDirectory(true).GetFullPath();
            target = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
        }

        // 判断哪些 sheet 有图片
        var sheetsWithImages = new HashSet<Int32>();
        var globalImageIndex = 0;
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetImages.TryGetValue(_sheetNames[i], out var imgs) && imgs.Count > 0)
                sheetsWithImages.Add(i);
        }

        // 判断哪些 sheet 有超链接
        var sheetsWithHyperlinks = new HashSet<Int32>();
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetHyperlinks.TryGetValue(_sheetNames[i], out var links) && links.Count > 0)
                sheetsWithHyperlinks.Add(i);
        }

        // 判断哪些 sheet 需要打印标题行
        var sheetsWithPrintTitles = new HashSet<Int32>();
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetPageSetups.TryGetValue(_sheetNames[i], out var ps) && ps.PrintTitleStartRow > 0)
                sheetsWithPrintTitles.Add(i);
        }

        // 判断哪些 sheet 有批注
        var sheetsWithComments = new HashSet<Int32>();
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetComments.TryGetValue(_sheetNames[i], out var cmts) && cmts.Count > 0)
                sheetsWithComments.Add(i);
        }

        using var za = new ZipArchive(target, ZipArchiveMode.Create, leaveOpen: Stream != null, entryNameEncoding: Encoding);

        // _rels/.rels
        using (var sw = new StreamWriter(za.CreateEntry("_rels/.rels").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        }

        // [Content_Types].xml
        using (var sw = new StreamWriter(za.CreateEntry("[Content_Types].xml").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sw.Write("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                sw.Write("<Override PartName=\"/xl/worksheets/sheet");
                sw.Write(i + 1);
                sw.Write(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }
            if (_shared.Count > 0)
            {
                sw.Write("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            }
            sw.Write("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            // 图片类型
            var imageExts = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in _sheetImages)
            {
                foreach (var img in kv.Value)
                {
                    imageExts.Add(img.Extension);
                }
            }
            foreach (var ext in imageExts)
            {
                var mime = ext == "png" ? "image/png" : ext == "jpeg" || ext == "jpg" ? "image/jpeg" : ext == "gif" ? "image/gif" : "image/png";
                sw.Write($"<Default Extension=\"{ext}\" ContentType=\"{mime}\"/>");
            }
            // Drawing
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                if (sheetsWithImages.Contains(i))
                    sw.Write($"<Override PartName=\"/xl/drawings/drawing{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
            }
            // 批注
            if (sheetsWithComments.Count > 0)
                sw.Write("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>");
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                if (sheetsWithComments.Contains(i))
                    sw.Write($"<Override PartName=\"/xl/comments{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>");
            }
            sw.Write("</Types>");
        }

        // workbook.xml
        using (var sw = new StreamWriter(za.CreateEntry("xl/workbook.xml").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            for (var i = 0; i < _sheetNames.Count; i++)
            {
                var name = SecurityElement.Escape(_sheetNames[i]) ?? _sheetNames[i];
                sw.Write($"<sheet name=\"{name}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
            }
            sw.Write("</sheets>");
            // 打印标题行的 definedNames
            if (sheetsWithPrintTitles.Count > 0)
            {
                sw.Write("<definedNames>");
                foreach (var si in sheetsWithPrintTitles)
                {
                    var ps = _sheetPageSetups[_sheetNames[si]];
                    var sn = SecurityElement.Escape(_sheetNames[si]) ?? _sheetNames[si];
                    sw.Write($"<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"{si}\">'{sn}'!${ ps.PrintTitleStartRow}:${ps.PrintTitleEndRow}</definedName>");
                }
                sw.Write("</definedNames>");
            }
            sw.Write("</workbook>");
        }

        // workbook 关系
        using (var sw = new StreamWriter(za.CreateEntry("xl/_rels/workbook.xml.rels").Open(), Encoding))
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (var i = 0; i < _sheetNames.Count; i++) sw.Write($"<Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
            var nextId = _sheetNames.Count + 1;
            sw.Write($"<Relationship Id=\"rId{nextId++}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            if (_shared.Count > 0) sw.Write($"<Relationship Id=\"rId{nextId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            sw.Write("</Relationships>");
        }

        // styles.xml（完整版：numFmts + fonts + fills + borders + cellXfs）
        WriteStylesXml(za);

        // sharedStrings.xml
        if (_shared.Count > 0)
        {
            using var sw = new StreamWriter(za.CreateEntry("xl/sharedStrings.xml").Open(), Encoding);
            sw.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{_sharedCount}\" uniqueCount=\"{_shared.Count}\">");
            foreach (var kv in _shared.OrderBy(e => e.Value))
            {
                var txt = SecurityElement.Escape(kv.Key) ?? String.Empty;
                sw.Write("<si><t>");
                sw.Write(txt);
                sw.Write("</t></si>");
            }
            sw.Write("</sst>");
        }

        // worksheets
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            var sheet = _sheetNames[i];
            var entry = za.CreateEntry($"xl/worksheets/sheet{i + 1}.xml");
            using var sw = new StreamWriter(entry.Open(), Encoding);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:etc=\"http://www.wps.cn/officeDocument/2017/etCustomData\">");

            // sheetViews（冻结窗格）
            if (_sheetFreezes.TryGetValue(sheet, out var freeze) && (freeze.Rows > 0 || freeze.Cols > 0))
            {
                sw.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">");
                var topLeft = MakeCellRef(freeze.Rows, freeze.Cols);
                if (freeze.Rows > 0 && freeze.Cols > 0)
                {
                    sw.Write($"<pane xSplit=\"{freeze.Cols}\" ySplit=\"{freeze.Rows}\" topLeftCell=\"{topLeft}\" activePane=\"bottomRight\" state=\"frozen\"/>");
                }
                else if (freeze.Rows > 0)
                {
                    sw.Write($"<pane ySplit=\"{freeze.Rows}\" topLeftCell=\"{topLeft}\" activePane=\"bottomLeft\" state=\"frozen\"/>");
                }
                else
                {
                    sw.Write($"<pane xSplit=\"{freeze.Cols}\" topLeftCell=\"{topLeft}\" activePane=\"topRight\" state=\"frozen\"/>");
                }
                sw.Write("</sheetView></sheetViews>");
            }

            // cols（列宽）
            if (AutoFitColumnWidth && _sheetColWidths.TryGetValue(sheet, out var widths) && widths.Count > 0)
            {
                if (widths.Any(e => e > 0))
                {
                    sw.Write("<cols>");
                    for (var c = 0; c < widths.Count; c++)
                    {
                        var w = widths[c];
                        if (w <= 0) continue;
                        sw.Write($"<col min=\"{c + 1}\" max=\"{c + 1}\" width=\"{w:0.##}\" customWidth=\"1\"/>");
                    }
                    sw.Write("</cols>");
                }
            }

            // sheetData（带行高注入）
            sw.Write("<sheetData>");
            if (_sheetRows.TryGetValue(sheet, out var list))
            {
                var hasHeights = _sheetRowHeights.TryGetValue(sheet, out var heights) && heights.Count > 0;
                var rowNum = 1;
                foreach (var r in list)
                {
                    if (hasHeights && heights!.TryGetValue(rowNum, out var ht))
                    {
                        sw.Write(r.Replace($"<row r=\"{rowNum}\"", $"<row r=\"{rowNum}\" ht=\"{ht:0.##}\" customHeight=\"1\""));
                    }
                    else
                    {
                        sw.Write(r);
                    }
                    rowNum++;
                }
            }
            sw.Write("</sheetData>");

            // sheetProtection
            if (_sheetProtection.TryGetValue(sheet, out var pwd))
            {
                sw.Write("<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"");
                if (!pwd.IsNullOrEmpty())
                {
                    var hash = ComputeSheetProtectionHash(pwd);
                    sw.Write($" password=\"{hash}\"");
                }
                sw.Write("/>");
            }

            // autoFilter
            if (_sheetAutoFilters.TryGetValue(sheet, out var filter))
            {
                sw.Write($"<autoFilter ref=\"{filter}\"/>");
            }

            // mergeCells
            if (_sheetMerges.TryGetValue(sheet, out var merges) && merges.Count > 0)
            {
                sw.Write($"<mergeCells count=\"{merges.Count}\">");
                foreach (var (sr, sc, er, ec) in merges)
                {
                    sw.Write($"<mergeCell ref=\"{MakeCellRef(sr, sc)}:{MakeCellRef(er, ec)}\"/>");
                }
                sw.Write("</mergeCells>");
            }

            // conditionalFormatting
            if (_sheetCondFormats.TryGetValue(sheet, out var conds) && conds.Count > 0)
            {
                var priority = 1;
                foreach (var cf in conds)
                {
                    sw.Write($"<conditionalFormatting sqref=\"{cf.Range}\">");
                    switch (cf.Type)
                    {
                        case ConditionalFormatType.GreaterThan:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"greaterThan\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.LessThan:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"lessThan\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.Equal:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"equal\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.Between:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"between\"><formula>{SecurityElement.Escape(cf.Value)}</formula><formula>{SecurityElement.Escape(cf.Value2)}</formula></cfRule>");
                            break;
                        case ConditionalFormatType.DataBar:
                            sw.Write($"<cfRule type=\"dataBar\" priority=\"{priority++}\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF{cf.Color ?? "4472C4"}\"/></dataBar></cfRule>");
                            break;
                        case ConditionalFormatType.ColorScale:
                            sw.Write($"<cfRule type=\"colorScale\" priority=\"{priority++}\"><colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFFFFFFF\"/><color rgb=\"FF{cf.Color ?? "4472C4"}\"/></colorScale></cfRule>");
                            break;
                    }
                    sw.Write("</conditionalFormatting>");
                }
            }

            // dataValidations
            if (_sheetValidations.TryGetValue(sheet, out var validations) && validations.Count > 0)
            {
                sw.Write($"<dataValidations count=\"{validations.Count}\">");
                foreach (var v in validations)
                {
                    if (v.Items != null)
                    {
                        var formula = "\"" + String.Join(",", v.Items.Select(e => SecurityElement.Escape(e))) + "\"";
                        sw.Write($"<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"{v.CellRange}\"><formula1>{formula}</formula1></dataValidation>");
                    }
                    else if (!v.ValidationType.IsNullOrEmpty())
                    {
                        var op = v.Operator ?? "between";
                        sw.Write($"<dataValidation type=\"{v.ValidationType}\" operator=\"{op}\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"{v.CellRange}\">");
                        sw.Write($"<formula1>{SecurityElement.Escape(v.Formula1 ?? "0")}</formula1>");
                        if (!v.Formula2.IsNullOrEmpty()) sw.Write($"<formula2>{SecurityElement.Escape(v.Formula2!)}</formula2>");
                        sw.Write("</dataValidation>");
                    }
                }
                sw.Write("</dataValidations>");
            }

            // hyperlinks
            if (sheetsWithHyperlinks.Contains(i) && _sheetHyperlinks.TryGetValue(sheet, out var hyperlinks))
            {
                sw.Write("<hyperlinks>");
                for (var h = 0; h < hyperlinks.Count; h++)
                {
                    var hl = hyperlinks[h];
                    var cellRef = MakeCellRef(hl.Row - 1, hl.Col);
                    sw.Write($"<hyperlink ref=\"{cellRef}\" r:id=\"rHl{h + 1}\"");
                    if (!hl.Display.IsNullOrEmpty()) sw.Write($" display=\"{SecurityElement.Escape(hl.Display)}\"");
                    sw.Write("/>");
                }
                sw.Write("</hyperlinks>");
            }

            // pageMargins + pageSetup + headerFooter
            if (_sheetPageSetups.TryGetValue(sheet, out var pageSetup))
            {
                sw.Write($"<pageMargins left=\"{pageSetup.MarginLeft:0.##}\" right=\"{pageSetup.MarginRight:0.##}\" top=\"{pageSetup.MarginTop:0.##}\" bottom=\"{pageSetup.MarginBottom:0.##}\" header=\"0.3\" footer=\"0.3\"/>");
                var orient = pageSetup.Orientation == PageOrientation.Landscape ? "landscape" : "portrait";
                sw.Write($"<pageSetup orientation=\"{orient}\"");
                if (pageSetup.PaperSize != PaperSize.Default) sw.Write($" paperSize=\"{(Int32)pageSetup.PaperSize}\"");
                sw.Write("/>");
                if (!pageSetup.HeaderText.IsNullOrEmpty() || !pageSetup.FooterText.IsNullOrEmpty())
                {
                    sw.Write("<headerFooter>");
                    if (!pageSetup.HeaderText.IsNullOrEmpty()) sw.Write($"<oddHeader>{SecurityElement.Escape(pageSetup.HeaderText)}</oddHeader>");
                    if (!pageSetup.FooterText.IsNullOrEmpty()) sw.Write($"<oddFooter>{SecurityElement.Escape(pageSetup.FooterText)}</oddFooter>");
                    sw.Write("</headerFooter>");
                }
            }

            // drawing（图片引用）
            if (sheetsWithImages.Contains(i))
            {
                sw.Write($"<drawing r:id=\"rDr1\"/>");
            }

            // legacyDrawing（批注 VML 引用）
            if (sheetsWithComments.Contains(i))
            {
                sw.Write($"<legacyDrawing r:id=\"rVml1\"/>");
            }

            sw.Write("</worksheet>");
            sw.Dispose();

            // sheet rels（超链接 + 图片 drawing + 批注关系）
            if (sheetsWithHyperlinks.Contains(i) || sheetsWithImages.Contains(i) || sheetsWithComments.Contains(i))
            {
                var relEntry = za.CreateEntry($"xl/worksheets/_rels/sheet{i + 1}.xml.rels");
                using var rsw = new StreamWriter(relEntry.Open(), Encoding);
                rsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                if (sheetsWithHyperlinks.Contains(i) && _sheetHyperlinks.TryGetValue(sheet, out var rels))
                {
                    for (var h = 0; h < rels.Count; h++)
                    {
                        rsw.Write($"<Relationship Id=\"rHl{h + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{SecurityElement.Escape(rels[h].Url)}\" TargetMode=\"External\"/>");
                    }
                }
                if (sheetsWithImages.Contains(i))
                {
                    rsw.Write($"<Relationship Id=\"rDr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing{i + 1}.xml\"/>");
                }
                if (sheetsWithComments.Contains(i))
                {
                    rsw.Write($"<Relationship Id=\"rVml1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing{i + 1}.vml\"/>");
                    rsw.Write($"<Relationship Id=\"rCmt1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments{i + 1}.xml\"/>");
                }
                rsw.Write("</Relationships>");
            }
        }

        // Drawings 和媒体文件
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (!sheetsWithImages.Contains(i)) continue;
            var sheet = _sheetNames[i];
            var images = _sheetImages[sheet];

            // drawing{i+1}.xml
            var drawEntry = za.CreateEntry($"xl/drawings/drawing{i + 1}.xml");
            using (var dsw = new StreamWriter(drawEntry.Open(), Encoding))
            {
                dsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                for (var j = 0; j < images.Count; j++)
                {
                    var img = images[j];
                    var emuW = (Int64)(img.Width * 9525); // px → EMU
                    var emuH = (Int64)(img.Height * 9525);
                    dsw.Write($"<xdr:twoCellAnchor><xdr:from><xdr:col>{img.Col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{img.Row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>");
                    dsw.Write($"<xdr:to><xdr:col>{img.Col + 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{img.Row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>");
                    dsw.Write($"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{j + 2}\" name=\"Image{globalImageIndex + 1}\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>");
                    dsw.Write($"<xdr:blipFill><a:blip r:embed=\"rImg{j + 1}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
                    dsw.Write($"<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{emuW}\" cy=\"{emuH}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>");
                    globalImageIndex++;
                }
                dsw.Write("</xdr:wsDr>");
            }

            // drawing rels
            var drawRelEntry = za.CreateEntry($"xl/drawings/_rels/drawing{i + 1}.xml.rels");
            using (var drsw = new StreamWriter(drawRelEntry.Open(), Encoding))
            {
                drsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (var j = 0; j < images.Count; j++)
                {
                    drsw.Write($"<Relationship Id=\"rImg{j + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{globalImageIndex - images.Count + j + 1}.{images[j].Extension}\"/>");
                }
                drsw.Write("</Relationships>");
            }

            // 媒体文件
            for (var j = 0; j < images.Count; j++)
            {
                var img = images[j];
                var mediaEntry = za.CreateEntry($"xl/media/image{globalImageIndex - images.Count + j + 1}.{img.Extension}");
                using var ms2 = mediaEntry.Open();
                ms2.Write(img.Data, 0, img.Data.Length);
            }
        }

        // 批注文件：xl/commentsN.xml + xl/drawings/vmlDrawingN.vml
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (!sheetsWithComments.Contains(i)) continue;
            var sheet = _sheetNames[i];
            var comments = _sheetComments[sheet];

            // 收集所有不同作者（保持插入顺序，用 List 去重）
            var authors = new List<String>();
            foreach (var c in comments)
            {
                if (!authors.Contains(c.Author)) authors.Add(c.Author);
            }

            // xl/commentsN.xml
            using (var csw = new StreamWriter(za.CreateEntry($"xl/comments{i + 1}.xml").Open(), Encoding))
            {
                csw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                csw.Write("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                csw.Write("<authors>");
                foreach (var a in authors) csw.Write($"<author>{SecurityElement.Escape(a)}</author>");
                csw.Write("</authors><commentList>");
                foreach (var c in comments)
                {
                    var cellRef = MakeCellRef(c.Row - 1, c.Col);
                    var authorId = authors.IndexOf(c.Author);
                    csw.Write($"<comment ref=\"{cellRef}\" authorId=\"{authorId}\">");
                    csw.Write($"<text><r><t xml:space=\"preserve\">{SecurityElement.Escape(c.Text)}</t></r></text>");
                    csw.Write("</comment>");
                }
                csw.Write("</commentList></comments>");
            }

            // xl/drawings/vmlDrawingN.vml
            using (var vsw = new StreamWriter(za.CreateEntry($"xl/drawings/vmlDrawing{i + 1}.vml").Open(), Encoding))
            {
                vsw.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                vsw.Write("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
                vsw.Write("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m0,0l0,21600,21600,21600,21600,0xe\">");
                vsw.Write("<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");
                for (var j = 0; j < comments.Count; j++)
                {
                    var c = comments[j];
                    vsw.Write($"<v:shape id=\"_x0000_s{1025 + j}\" type=\"#_x0000_t202\" " +
                              "style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" " +
                              "fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
                    vsw.Write("<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
                    vsw.Write("<v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>");
                    vsw.Write("<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/>");
                    vsw.Write($"<x:Row>{c.Row - 1}</x:Row><x:Column>{c.Col}</x:Column>");
                    vsw.Write("</x:ClientData></v:shape>");
                }
                vsw.Write("</xml>");
            }
        }

        target.Flush();
    }

    /// <summary>生成完整的 styles.xml</summary>
    private void WriteStylesXml(ZipArchive za)
    {
        using var sw = new StreamWriter(za.CreateEntry("xl/styles.xml").Open(), Encoding);
        sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        // numFmts（自定义）
        if (_numFmtMap.Count > 0)
        {
            sw.Write($"<numFmts count=\"{_numFmtMap.Count}\">");
            foreach (var kv in _numFmtMap)
            {
                sw.Write($"<numFmt numFmtId=\"{kv.Value}\" formatCode=\"{SecurityElement.Escape(kv.Key)}\"/>");
            }
            sw.Write("</numFmts>");
        }

        // fonts
        sw.Write($"<fonts count=\"{_fonts.Count}\">");
        foreach (var f in _fonts)
        {
            sw.Write("<font>");
            if (f.Bold) sw.Write("<b/>");
            if (f.Italic) sw.Write("<i/>");
            if (f.Underline) sw.Write("<u/>");
            if (f.Size > 0) sw.Write($"<sz val=\"{f.Size}\"/>");
            if (!f.Color.IsNullOrEmpty()) sw.Write($"<color rgb=\"FF{f.Color}\"/>");
            if (!f.Name.IsNullOrEmpty()) sw.Write($"<name val=\"{SecurityElement.Escape(f.Name)}\"/>");
            sw.Write("</font>");
        }
        sw.Write("</fonts>");

        // fills
        sw.Write($"<fills count=\"{_fills.Count}\">");
        foreach (var f in _fills)
        {
            sw.Write("<fill>");
            if (f.PatternType == "none")
                sw.Write("<patternFill patternType=\"none\"/>");
            else if (f.PatternType == "gray125")
                sw.Write("<patternFill patternType=\"gray125\"/>");
            else
                sw.Write($"<patternFill patternType=\"solid\"><fgColor rgb=\"FF{f.BgColor}\"/></patternFill>");
            sw.Write("</fill>");
        }
        sw.Write("</fills>");

        // borders
        sw.Write($"<borders count=\"{_borders.Count}\">");
        foreach (var b in _borders)
        {
            if (b.Style == CellBorderStyle.None)
            {
                sw.Write("<border><left/><right/><top/><bottom/><diagonal/></border>");
            }
            else
            {
                var sn = GetBorderStyleName(b.Style);
                var ca = b.Color.IsNullOrEmpty() ? "" : $"<color rgb=\"FF{b.Color}\"/>";
                sw.Write($"<border><left style=\"{sn}\">{ca}</left><right style=\"{sn}\">{ca}</right><top style=\"{sn}\">{ca}</top><bottom style=\"{sn}\">{ca}</bottom><diagonal/></border>");
            }
        }
        sw.Write("</borders>");

        // cellXfs
        sw.Write($"<cellXfs count=\"{_xfEntries.Count}\">");
        foreach (var xf in _xfEntries)
        {
            sw.Write($"<xf numFmtId=\"{xf.NumFmtId}\" fontId=\"{xf.FontId}\" fillId=\"{xf.FillId}\" borderId=\"{xf.BorderId}\"");
            if (xf.FontId > 0) sw.Write(" applyFont=\"1\"");
            if (xf.FillId > 0) sw.Write(" applyFill=\"1\"");
            if (xf.BorderId > 0) sw.Write(" applyBorder=\"1\"");
            if (xf.NumFmtId > 0) sw.Write(" applyNumberFormat=\"1\"");
            if (xf.HAlign != HorizontalAlignment.General || xf.VAlign != VerticalAlignment.Top || xf.WrapText)
            {
                sw.Write(" applyAlignment=\"1\"><alignment");
                if (xf.HAlign != HorizontalAlignment.General) sw.Write($" horizontal=\"{xf.HAlign.ToString().ToLower()}\"");
                if (xf.VAlign != VerticalAlignment.Top) sw.Write($" vertical=\"{xf.VAlign.ToString().ToLower()}\"");
                if (xf.WrapText) sw.Write(" wrapText=\"1\"");
                sw.Write("/></xf>");
            }
            else
            {
                sw.Write("/>");
            }
        }
        sw.Write("</cellXfs>");

        // 条件格式需要的 dxf（差异格式）
        var totalDxf = 0;
        foreach (var kv in _sheetCondFormats)
        {
            foreach (var cf in kv.Value)
            {
                if (cf.Type < ConditionalFormatType.DataBar) totalDxf++;
            }
        }
        if (totalDxf > 0)
        {
            sw.Write($"<dxfs count=\"{totalDxf}\">");
            foreach (var kv in _sheetCondFormats)
            {
                foreach (var cf in kv.Value)
                {
                    if (cf.Type >= ConditionalFormatType.DataBar) continue;
                    sw.Write("<dxf>");
                    if (!cf.Color.IsNullOrEmpty())
                        sw.Write($"<fill><patternFill><bgColor rgb=\"FF{cf.Color}\"/></patternFill></fill>");
                    sw.Write("</dxf>");
                }
            }
            sw.Write("</dxfs>");
        }

        sw.Write("</styleSheet>");
    }

    /// <summary>计算工作表保护密码哈希（Excel 传统算法）</summary>
    private static String ComputeSheetProtectionHash(String password)
    {
        var hash = 0;
        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash ^= password[i];
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        }
        hash ^= password.Length;
        hash ^= 0xCE4B;
        return hash.ToString("X4");
    }
    #endregion
}