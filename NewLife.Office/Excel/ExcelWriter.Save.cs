using System.Data;
using System.IO.Compression;
using System.Security;

namespace NewLife.Office;

partial class ExcelWriter
{
    #region 样式管理
    /// <summary>根据用户样式和数字格式，查找或创建 XF 条目并返回索引</summary>
    private Int32 GetOrCreateXf(Office.ExcelCellStyle cs, Int32 numFmtId)
    {
        // 找或创建字体
        var font = new FontEntry(cs.FontName, cs.FontSize, cs.Bold, cs.Italic, cs.Underline, cs.FontColor, cs.Strike, cs.VerticalAlign);
        var fontId = FindOrAdd(_fonts, font);

        // 找或创建填充
        var fillId = 0;
        if (!cs.BackgroundColor.IsNullOrEmpty())
        {
            var fill = new FillEntry(cs.BackgroundColor, "solid");
            fillId = FindOrAdd(_fills, fill);
        }
        else if (!cs.GradientColor1.IsNullOrEmpty() && !cs.GradientColor2.IsNullOrEmpty())
        {
            var gradType = cs.GradientType.EqualIgnoreCase("radial") ? "radial" : "linear";
            var fill = new FillEntry(null, "gradient", gradType, cs.GradientColor1, cs.GradientColor2);
            fillId = FindOrAdd(_fills, fill);
        }
        else if (!cs.PatternType.IsNullOrEmpty())
        {
            var fill = new FillEntry(null, "pattern", PatternFgColor: cs.PatternFgColor, PatternTypeName: cs.PatternType);
            fillId = FindOrAdd(_fills, fill);
        }

        // 找或创建边框：单边属性优先，回退到全局 Border
        var borderId = 0;
        var leftStyle   = cs.LeftBorder   != ExcelCellBorderStyle.None ? cs.LeftBorder   : cs.Border;
        var rightStyle  = cs.RightBorder  != ExcelCellBorderStyle.None ? cs.RightBorder  : cs.Border;
        var topStyle    = cs.TopBorder    != ExcelCellBorderStyle.None ? cs.TopBorder    : cs.Border;
        var bottomStyle = cs.BottomBorder != ExcelCellBorderStyle.None ? cs.BottomBorder : cs.Border;
        var leftColor   = cs.LeftBorderColor   ?? cs.BorderColor;
        var rightColor  = cs.RightBorderColor  ?? cs.BorderColor;
        var topColor    = cs.TopBorderColor    ?? cs.BorderColor;
        var bottomColor = cs.BottomBorderColor ?? cs.BorderColor;
        var diagonalStyle = cs.DiagonalBorder;
        var diagonalColor = cs.DiagonalBorderColor;
        if (leftStyle != ExcelCellBorderStyle.None || rightStyle != ExcelCellBorderStyle.None ||
            topStyle  != ExcelCellBorderStyle.None || bottomStyle != ExcelCellBorderStyle.None ||
            diagonalStyle != ExcelCellBorderStyle.None)
        {
            var border = new BorderEntry(leftStyle, leftColor, rightStyle, rightColor, topStyle, topColor, bottomStyle, bottomColor, diagonalStyle, diagonalColor);
            borderId = FindOrAdd(_borders, border);
        }

        // 复合键去重
        var key = $"{numFmtId}-{fontId}-{fillId}-{borderId}-{(Int32)cs.HAlign}-{(Int32)cs.VAlign}-{(cs.WrapText ? 1 : 0)}-{cs.TextRotation}-{cs.Indent}-{(cs.ShrinkToFit ? 1 : 0)}";
        if (_xfCache.TryGetValue(key, out var idx)) return idx;

        var xf = new XfEntry(numFmtId, fontId, fillId, borderId, cs.HAlign, cs.VAlign, cs.WrapText, cs.TextRotation, cs.Indent, cs.ShrinkToFit);
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

        var rowStr = cellRef[colLen..];
        var rowIndex = Int32.Parse(rowStr) - 1; // 转 0 基

        return (rowIndex, colIndex);
    }

    /// <summary>生成单元格引用（如 "A1"），行列均为 0 基</summary>
    private static String MakeCellRef(Int32 row, Int32 col) => GetColumnName(col) + (row + 1);

    /// <summary>生成结构化表格 XML（xl/tables/tableN.xml）</summary>
    private void WriteTableXml(StreamWriter sw, ExcelTableInfo tbl, Int32 tableId)
    {
        sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sw.Write("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
        var eName = SecurityElement.Escape(tbl.Name) ?? tbl.Name;
        sw.Write($" id=\"{tableId}\" name=\"{eName}\" displayName=\"{eName}\" ref=\"{tbl.Range}\">");

        // autoFilter（仅当有筛选按钮时）
        if (tbl.ShowFilterButton)
            sw.Write($"<autoFilter ref=\"{tbl.Range}\"/>");

        // 解析列数
        var colCount = ParseRangeColumnCount(tbl.Range);
        sw.Write($"<tableColumns count=\"{colCount}\">");
        for (var c = 0; c < colCount; c++)
        {
            var colName = tbl.ColumnNames != null && c < tbl.ColumnNames.Length
                ? SecurityElement.Escape(tbl.ColumnNames[c]) ?? tbl.ColumnNames[c]
                : $"Column{c + 1}";
            sw.Write($"<tableColumn id=\"{c + 1}\" name=\"{colName}\"/>");
        }
        sw.Write("</tableColumns>");

        // 样式
        var styleName = tbl.StyleName.IsNullOrEmpty() ? "TableStyleMedium9" : SecurityElement.Escape(tbl.StyleName) ?? tbl.StyleName;
        var first  = tbl.ShowFirstColumn ? "1" : "0";
        var last   = tbl.ShowLastColumn  ? "1" : "0";
        var rowStr = tbl.ShowRowStripes  ? "1" : "0";
        var colStr = tbl.ShowColumnStripes ? "1" : "0";
        sw.Write($"<tableStyleInfo name=\"{styleName}\" showFirstColumn=\"{first}\" showLastColumn=\"{last}\" showRowStripes=\"{rowStr}\" showColumnStripes=\"{colStr}\"/>");

        sw.Write("</table>");
    }

    /// <summary>从 Excel 范围字符串（如 "A1:E10" 或 "B3:D8"）解析列数</summary>
    private static Int32 ParseRangeColumnCount(String range)
    {
        if (range.IsNullOrEmpty()) return 1;
        var sep = range.IndexOf(':');
        if (sep < 0) return 1;
        var (_, startCol) = ParseCellRef(range[..sep]);
        var (_, endCol)   = ParseCellRef(range[(sep + 1)..]);
        return Math.Max(1, endCol - startCol + 1);
    }

    /// <summary>生成图表 XML（xl/charts/chartN.xml）</summary>
    private static void WriteChartXml(StreamWriter sw, ExcelChart chart, Int32 chartId)
    {
        const String C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sw.Write($"<c:chartSpace xmlns:c=\"{C}\" xmlns:a=\"{A}\">");
        sw.Write("<c:date1904 val=\"0\"/>");
        sw.Write("<c:chart>");
        if (!chart.Title.IsNullOrEmpty())
        {
            var et = SecurityElement.Escape(chart.Title!) ?? chart.Title;
            sw.Write($"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"zh-CN\"/><a:t>{et}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>");
        }
        sw.Write("<c:autoTitleDeleted val=\"0\"/>");
        sw.Write("<c:plotArea>");

        var chartElem = chart.Type switch
        {
            "line" => "lineChart",
            "pie"  => "pieChart",
            "area" => "areaChart",
            "scatter" => "scatterChart",
            _ => "barChart",
        };
        sw.Write($"<c:{chartElem}>");
        if (chart.Type == "bar") sw.Write("<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>");
        else if (chart.Type == "line") sw.Write("<c:grouping val=\"standard\"/>");

        var serColors = new[] { "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646" };
        var categories = chart.Categories ?? [];
        for (var si = 0; si < chart.Series.Count; si++)
        {
            var ser = chart.Series[si];
            var color = serColors[si % serColors.Length];
            sw.Write("<c:ser>");
            sw.Write($"<c:idx val=\"{si}\"/><c:order val=\"{si}\"/>");
            var eName = SecurityElement.Escape(ser.Name) ?? ser.Name ?? "";
            sw.Write($"<c:tx><c:strRef><c:f/><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{eName}</c:v></c:pt></c:strCache></c:strRef></c:tx>");
            sw.Write($"<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr>");
            if (categories.Length > 0)
            {
                sw.Write("<c:cat><c:strRef><c:f/><c:strCache>");
                sw.Write($"<c:ptCount val=\"{categories.Length}\"/>");
                for (var ci = 0; ci < categories.Length; ci++)
                    sw.Write($"<c:pt idx=\"{ci}\"><c:v>{SecurityElement.Escape(categories[ci]) ?? categories[ci]}</c:v></c:pt>");
                sw.Write("</c:strCache></c:strRef></c:cat>");
            }
            sw.Write("<c:val><c:numRef><c:f/><c:numCache>");
            sw.Write($"<c:ptCount val=\"{ser.Data.Length}\"/>");
            for (var vi = 0; vi < ser.Data.Length; vi++)
                sw.Write($"<c:pt idx=\"{vi}\"><c:v>{ser.Data[vi]}</c:v></c:pt>");
            sw.Write("</c:numCache></c:numRef></c:val>");
            sw.Write("</c:ser>");
        }
        if (chart.Type != "pie")
        {
            sw.Write("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
            sw.Write($"</c:{chartElem}>");
            sw.Write("<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>");
            sw.Write("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>");
        }
        else
        {
            sw.Write($"</c:{chartElem}>");
        }
        sw.Write("</c:plotArea>");
        sw.Write("<c:legend><c:legendPos val=\"b\"/></c:legend>");
        sw.Write("</c:chart></c:chartSpace>");
    }

    /// <summary>写出单边边框 XML（style 为 None 时输出自关闭空元素）</summary>
    private static void WriteBorderSide(StreamWriter sw, String tag, ExcelCellBorderStyle style, String? color)
    {
        if (style == ExcelCellBorderStyle.None)
        {
            sw.Write($"<{tag}/>");
        }
        else
        {
            var sn = GetBorderStyleName(style);
            sw.Write($"<{tag} style=\"{sn}\">");
            WriteColorXml(sw, color);
            sw.Write($"</{tag}>");
        }
    }

    /// <summary>获取边框 OOXML 样式名</summary>
    private static String GetBorderStyleName(ExcelCellBorderStyle style) => style switch
    {
        ExcelCellBorderStyle.Thin => "thin",
        ExcelCellBorderStyle.Medium => "medium",
        ExcelCellBorderStyle.Thick => "thick",
        ExcelCellBorderStyle.Dashed => "dashed",
        ExcelCellBorderStyle.Dotted => "dotted",
        ExcelCellBorderStyle.DoubleLine => "double",
        _ => "thin",
    };

    /// <summary>将颜色字符串（RGB 六位 或 "theme:N"）写入 StreamWriter 为 &lt;color .../&gt;</summary>
    private static void WriteColorXml(StreamWriter sw, String? color)
    {
        if (color.IsNullOrEmpty()) return;
        sw.Write($"<color {FormatColorAttr(color)}/>");
    }

    /// <summary>将颜色字符串格式化为 color 元素的属性片段（不含尖括号）</summary>
    private static String FormatColorAttr(String? color)
    {
        if (color.IsNullOrEmpty()) return String.Empty;
        if (color!.StartsWith("theme:", StringComparison.Ordinal))
            return $"theme=\"{color[6..]}\"";
        return $"rgb=\"FF{color}\"";
    }
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

        // 判断哪些 sheet 有表格
        var sheetsWithTables = new Dictionary<Int32, List<ExcelTableInfo>>();
        var globalTableIndex = 0; // 全局表格编号（1基）
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetTables.TryGetValue(_sheetNames[i], out var tbls) && tbls.Count > 0)
                sheetsWithTables[i] = tbls;
        }

        // 判断哪些 sheet 有图表
        var sheetsWithCharts = new Dictionary<Int32, List<ExcelChart>>();
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            if (_sheetCharts.TryGetValue(_sheetNames[i], out var chs) && chs.Count > 0)
                sheetsWithCharts[i] = chs;
        }
        var globalChartId = 0; // 全局图表编号（1基）

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
            // 结构化表格
            foreach (var kv in sheetsWithTables)
            {
                foreach (var tbl in kv.Value)
                    sw.Write($"<Override PartName=\"/xl/tables/table{++globalTableIndex}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
            }
            globalTableIndex = 0; // 重置
            // 图表
            foreach (var kv in sheetsWithCharts)
            {
                foreach (var c in kv.Value)
                    sw.Write($"<Override PartName=\"/xl/charts/chart{++globalChartId}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
            }
            globalChartId = 0; // 重置
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
            // definedNames（打印标题行 + 用户自定义命名范围）
            var hasDefinedNames = sheetsWithPrintTitles.Count > 0 || _definedNames.Count > 0;
            if (hasDefinedNames)
            {
                sw.Write("<definedNames>");
                foreach (var si in sheetsWithPrintTitles)
                {
                    var ps = _sheetPageSetups[_sheetNames[si]];
                    var sn = SecurityElement.Escape(_sheetNames[si]) ?? _sheetNames[si];
                    sw.Write($"<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"{si}\">'{sn}'!${ps.PrintTitleStartRow}:${ps.PrintTitleEndRow}</definedName>");
                }
                foreach (var (dnName, dnFormula) in _definedNames)
                {
                    var en = SecurityElement.Escape(dnName) ?? dnName;
                    var ef = SecurityElement.Escape(dnFormula) ?? dnFormula;
                    sw.Write($"<definedName name=\"{en}\">{ef}</definedName>");
                }
                sw.Write("</definedNames>");
            }
            // 工作簿保护
            if (_workbookProtectionHash != null)
            {
                sw.Write("<workbookProtection");
                if (_workbookLockStructure) sw.Write(" lockStructure=\"1\"");
                if (_workbookLockWindows) sw.Write(" lockWindows=\"1\"");
                if (_workbookProtectionHash.Length > 0) sw.Write($" workbookPassword=\"{_workbookProtectionHash}\"");
                sw.Write("/>");
            }
            // 计算选项（确保 Excel 打开时自动重算）
            sw.Write("<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>");
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

            // sheetPr（标签颜色 + 大纲属性）
            var hasTabColor = _sheetTabColors.TryGetValue(sheet, out var tabColor) && !tabColor.IsNullOrEmpty();
            var hasColOl = _sheetColOutlines.ContainsKey(sheet);
            var hasRowOl = _sheetRowOutlines.ContainsKey(sheet);
            if (hasTabColor || hasColOl || hasRowOl)
            {
                sw.Write("<sheetPr>");
                if (hasTabColor) sw.Write($"<tabColor rgb=\"FF{tabColor}\"/>");
                if (hasColOl || hasRowOl) sw.Write("<outlinePr summaryBelow=\"1\" summaryRight=\"1\"/>");
                sw.Write("</sheetPr>");
            }

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

            // cols（列宽 + 列大纲级别）——有自定义列宽或列分组时始终写入
            var hasColOutlines = _sheetColOutlines.TryGetValue(sheet, out var colOutlines) && colOutlines.Count > 0;
            if ((_sheetColWidths.TryGetValue(sheet, out var widths) && widths.Any(e => e > 0)) || hasColOutlines)
            {
                // 收集所有需要输出的列索引
                var colSet = new SortedSet<Int32>();
                if (widths != null)
                    for (var c = 0; c < widths.Count; c++) if (widths[c] > 0) colSet.Add(c);
                if (hasColOutlines)
                    foreach (var c in colOutlines!.Keys) colSet.Add(c);

                if (colSet.Count > 0)
                {
                    sw.Write("<cols>");
                    foreach (var c in colSet)
                    {
                        var w = (widths != null && c < widths.Count) ? widths[c] : 0;
                        var outline = (hasColOutlines && colOutlines!.TryGetValue(c, out var co)) ? co : (Level: 0, Collapsed: false);
                        sw.Write($"<col min=\"{c + 1}\" max=\"{c + 1}\"");
                        if (w > 0) sw.Write($" width=\"{w:0.##}\" customWidth=\"1\"");
                        if (outline.Level > 0) sw.Write($" outlineLevel=\"{outline.Level}\"");
                        if (outline.Collapsed) sw.Write(" collapsed=\"1\"");
                        sw.Write("/>");
                    }
                    sw.Write("</cols>");
                }
            }

            // sheetData（带行高和行大纲级别注入）
            sw.Write("<sheetData>");
            if (_sheetRows.TryGetValue(sheet, out var list))
            {
                var hasHeights = _sheetRowHeights.TryGetValue(sheet, out var heights) && heights.Count > 0;
                var hasRowOutlines = _sheetRowOutlines.TryGetValue(sheet, out var rowOutlines) && rowOutlines.Count > 0;
                var rowNum = 1;
                foreach (var r in list)
                {
                    var rowTag = $"<row r=\"{rowNum}\"";
                    var replacement = rowTag;
                    if (hasHeights && heights!.TryGetValue(rowNum, out var ht))
                        replacement = $"<row r=\"{rowNum}\" ht=\"{ht:0.##}\" customHeight=\"1\"";
                    if (hasRowOutlines && rowOutlines!.TryGetValue(rowNum, out var ro) && ro.Level > 0)
                    {
                        var tag = replacement == rowTag
                            ? $"<row r=\"{rowNum}\""
                            : replacement.TrimEnd('"');
                        replacement = tag + $" outlineLevel=\"{ro.Level}\"" + (ro.Collapsed ? " collapsed=\"1\"" : "") + (replacement == rowTag ? "" : "\"");
                    }
                    sw.Write(r.Replace(rowTag, replacement));
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
                        case ExcelConditionalFormatType.GreaterThan:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"greaterThan\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ExcelConditionalFormatType.LessThan:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"lessThan\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ExcelConditionalFormatType.Equal:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"equal\"><formula>{SecurityElement.Escape(cf.Value)}</formula></cfRule>");
                            break;
                        case ExcelConditionalFormatType.Between:
                            sw.Write($"<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"{priority++}\" operator=\"between\"><formula>{SecurityElement.Escape(cf.Value)}</formula><formula>{SecurityElement.Escape(cf.Value2)}</formula></cfRule>");
                            break;
                        case ExcelConditionalFormatType.DataBar:
                            sw.Write($"<cfRule type=\"dataBar\" priority=\"{priority++}\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF{cf.Color ?? "4472C4"}\"/></dataBar></cfRule>");
                            break;
                        case ExcelConditionalFormatType.ColorScale:
                            sw.Write($"<cfRule type=\"colorScale\" priority=\"{priority++}\"><colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFFFFFFF\"/><color rgb=\"FF{cf.Color ?? "4472C4"}\"/></colorScale></cfRule>");
                            break;
                        case ExcelConditionalFormatType.IconSet:
                            {
                                var its = cf.IconSetType ?? "3Arrows";
                                var count = its[0] - '0'; // 取前缀数字
                                if (count < 3 || count > 5) count = 3;
                                sw.Write($"<cfRule type=\"iconSet\" priority=\"{priority++}\"><iconSet iconSet=\"{SecurityElement.Escape(its)}\">");
                                for (var p = 0; p < count; p++)
                                {
                                    var pct = p == 0 ? 0 : (Int32)Math.Round(100.0 * p / count);
                                    sw.Write($"<cfvo type=\"percent\" val=\"{pct}\"/>");
                                }
                                sw.Write("</iconSet></cfRule>");
                                break;
                            }
                        case ExcelConditionalFormatType.Expression:
                            {
                                var esc = SecurityElement.Escape(cf.Formula) ?? cf.Formula ?? String.Empty;
                                if (cf.Color.IsNullOrEmpty())
                                    sw.Write($"<cfRule type=\"expression\" dxfId=\"0\" priority=\"{priority++}\"><formula>{esc}</formula></cfRule>");
                                else
                                    sw.Write($"<cfRule type=\"expression\" dxfId=\"0\" priority=\"{priority++}\"><formula>{esc}</formula></cfRule>");
                                break;
                            }
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
                var orient = pageSetup.Orientation == ExcelPageOrientation.Landscape ? "landscape" : "portrait";
                sw.Write($"<pageSetup orientation=\"{orient}\"");
                if (pageSetup.PaperSize != ExcelPaperSize.Default) sw.Write($" paperSize=\"{(Int32)pageSetup.PaperSize}\"");
                sw.Write("/>");
                if (!pageSetup.HeaderText.IsNullOrEmpty() || !pageSetup.FooterText.IsNullOrEmpty())
                {
                    sw.Write("<headerFooter>");
                    if (!pageSetup.HeaderText.IsNullOrEmpty()) sw.Write($"<oddHeader>{SecurityElement.Escape(pageSetup.HeaderText)}</oddHeader>");
                    if (!pageSetup.FooterText.IsNullOrEmpty()) sw.Write($"<oddFooter>{SecurityElement.Escape(pageSetup.FooterText)}</oddFooter>");
                    sw.Write("</headerFooter>");
                }
            }

            // drawing（图片 + 图表引用）
            var hasDrawing = sheetsWithImages.Contains(i) || sheetsWithCharts.ContainsKey(i);
            if (hasDrawing)
            {
                sw.Write($"<drawing r:id=\"rDr1\"/>");
            }

            // legacyDrawing（批注 VML 引用）
            if (sheetsWithComments.Contains(i))
            {
                sw.Write($"<legacyDrawing r:id=\"rVml1\"/>");
            }

            // tableParts（结构化表格引用）
            if (sheetsWithTables.TryGetValue(i, out var shTables))
            {
                sw.Write($"<tableParts count=\"{shTables.Count}\">");
                for (var t = 0; t < shTables.Count; t++)
                    sw.Write($"<tablePart r:id=\"rTbl{t + 1}\"/>");
                sw.Write("</tableParts>");
            }

            sw.Write("</worksheet>");
            sw.Dispose();

            // sheet rels（超链接 + 图片 drawing + 批注 + 表格 + 图表关系）
            var needSheetRels = sheetsWithHyperlinks.Contains(i) || sheetsWithImages.Contains(i) ||
                                sheetsWithComments.Contains(i) || sheetsWithTables.ContainsKey(i) || sheetsWithCharts.ContainsKey(i);
            if (needSheetRels)
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
                if (sheetsWithTables.TryGetValue(i, out var tblRels))
                {
                    for (var t = 0; t < tblRels.Count; t++)
                    {
                        rsw.Write($"<Relationship Id=\"rTbl{t + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{globalTableIndex + t + 1}.xml\"/>");
                    }
                    globalTableIndex += tblRels.Count;
                }
                // 图表关系
                if (sheetsWithCharts.TryGetValue(i, out var chartRels))
                {
                    for (var c = 0; c < chartRels.Count; c++)
                    {
                        rsw.Write($"<Relationship Id=\"rCh{c + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{globalChartId + c + 1}.xml\"/>");
                    }
                    globalChartId += chartRels.Count;
                }
                rsw.Write("</Relationships>");
            }
        }

        // Drawings、媒体文件和图表
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            var hasImages = sheetsWithImages.Contains(i);
            var hasCharts = sheetsWithCharts.ContainsKey(i);
            if (!hasImages && !hasCharts) continue;
            var sheet = _sheetNames[i];

            // drawing{i+1}.xml（包含图片和图表锚点）
            var drawEntry = za.CreateEntry($"xl/drawings/drawing{i + 1}.xml");
            using (var dsw = new StreamWriter(drawEntry.Open(), Encoding))
            {
                dsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                // 图片锚点
                if (hasImages && _sheetImages.TryGetValue(sheet, out var images))
                {
                    for (var j = 0; j < images.Count; j++)
                    {
                        var img = images[j];
                        var emuW = (Int64)(img.Width * 9525);
                        var emuH = (Int64)(img.Height * 9525);
                        var editAs = img.EditAs.IsNullOrEmpty() ? "oneCell" : img.EditAs;
                        dsw.Write($"<xdr:twoCellAnchor editAs=\"{editAs}\">");
                        dsw.Write($"<xdr:from><xdr:col>{img.Col}</xdr:col><xdr:colOff>{img.FromColOff}</xdr:colOff><xdr:row>{img.Row}</xdr:row><xdr:rowOff>{img.FromRowOff}</xdr:rowOff></xdr:from>");
                        if (img.ToRow >= 0 && img.ToCol >= 0)
                            dsw.Write($"<xdr:to><xdr:col>{img.ToCol}</xdr:col><xdr:colOff>{img.ToColOff}</xdr:colOff><xdr:row>{img.ToRow}</xdr:row><xdr:rowOff>{img.ToRowOff}</xdr:rowOff></xdr:to>");
                        else
                            dsw.Write($"<xdr:to><xdr:col>{img.Col + 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{img.Row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>");
                        dsw.Write($"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{j + 2}\" name=\"Image{globalImageIndex + 1}\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>");
                        dsw.Write($"<xdr:blipFill><a:blip r:embed=\"rImg{j + 1}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
                        dsw.Write($"<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{emuW}\" cy=\"{emuH}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>");
                        globalImageIndex++;
                    }
                }
                // 图表锚点
                if (hasCharts && sheetsWithCharts.TryGetValue(i, out var charts))
                {
                    for (var c = 0; c < charts.Count; c++)
                    {
                        var chart = charts[c];
                        var emuW = (Int64)(chart.WidthPx * 9525);
                        var emuH = (Int64)(chart.HeightPx * 9525);
                        dsw.Write("<xdr:twoCellAnchor>");
                        dsw.Write($"<xdr:from><xdr:col>{chart.AnchorCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{chart.AnchorRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>");
                        dsw.Write($"<xdr:to><xdr:col>{chart.AnchorCol + 8}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{chart.AnchorRow + 16}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>");
                        var cNvPrId = 2 + (hasImages && _sheetImages.TryGetValue(sheet, out var ims) ? ims.Count : 0) + c;
                        dsw.Write($"<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{cNvPrId}\" name=\"Chart {c + 1}\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>");
                        dsw.Write($"<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{emuW}\" cy=\"{emuH}\"/></xdr:xfrm>");
                        dsw.Write($"<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rCh{c + 1}\"/></a:graphicData></a:graphic>");
                        dsw.Write("</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>");
                    }
                }
                dsw.Write("</xdr:wsDr>");
            }

            // drawing rels
            {
                _sheetImages.TryGetValue(sheet, out var drawImgs);
                var hasDrawImgs = drawImgs != null && drawImgs.Count > 0;
                var hasDrawCharts = sheetsWithCharts.TryGetValue(i, out var dCharts) && dCharts.Count > 0;
                if (hasDrawImgs || hasDrawCharts)
                {
                    var drawRelEntry = za.CreateEntry($"xl/drawings/_rels/drawing{i + 1}.xml.rels");
                    using var drsw = new StreamWriter(drawRelEntry.Open(), Encoding);
                    drsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    if (hasDrawImgs)
                    {
                        for (var j = 0; j < drawImgs!.Count; j++)
                            drsw.Write($"<Relationship Id=\"rImg{j + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{globalImageIndex - drawImgs.Count + j + 1}.{drawImgs[j].Extension}\"/>");
                    }
                    if (hasDrawCharts)
                    {
                        for (var c = 0; c < dCharts!.Count; c++)
                            drsw.Write($"<Relationship Id=\"rCh{c + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{globalChartId - dCharts.Count + c + 1}.xml\"/>");
                    }
                    drsw.Write("</Relationships>");
                }
            }

            // 媒体文件（仅图片）
            if (hasImages && _sheetImages.TryGetValue(sheet, out var mediaImgs))
            {
                for (var j = 0; j < mediaImgs.Count; j++)
                {
                    var img = mediaImgs[j];
                    var mediaEntry = za.CreateEntry($"xl/media/image{globalImageIndex - mediaImgs.Count + j + 1}.{img.Extension}");
                    using var ms2 = mediaEntry.Open();
                    ms2.Write(img.Data, 0, img.Data.Length);
                }
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

        // 写入 OtherParts（Reader 收集的原始 ZIP 部件，确保往返不丢内容）
        WriteOtherParts(za);

        // 结构化表格 XML 文件
        var tblIndex = 0;
        foreach (var kv in sheetsWithTables)
        {
            foreach (var tbl in kv.Value)
            {
                tblIndex++;
                using var tsw = new StreamWriter(za.CreateEntry($"xl/tables/table{tblIndex}.xml").Open(), Encoding);
                WriteTableXml(tsw, tbl, tblIndex);
            }
        }

        // 图表 XML 文件
        var cIndex = 0;
        foreach (var kv in sheetsWithCharts)
        {
            foreach (var chart in kv.Value)
            {
                cIndex++;
                using var csw = new StreamWriter(za.CreateEntry($"xl/charts/chart{cIndex}.xml").Open(), Encoding);
                WriteChartXml(csw, chart, cIndex);
                // chart rels
                var cRelEntry = za.CreateEntry($"xl/charts/_rels/chart{cIndex}.xml.rels");
                using var crsw = new StreamWriter(cRelEntry.Open(), Encoding);
                crsw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                crsw.Write("</Relationships>");
            }
        }

        target.Flush();
    }

    /// <summary>写入 OtherParts 中未被显式处理过的部件</summary>
    private void WriteOtherParts(ZipArchive za)
    {
        if (_otherParts.Count == 0) return;

        // 已由 Writer 显式生成的部件
        var generated = new HashSet<String>(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        };
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            generated.Add($"xl/worksheets/sheet{i + 1}.xml");
            // 超链接/图片/批注产生的 rels 也跳过
            if (_sheetHyperlinks.ContainsKey(_sheetNames[i]) ||
                _sheetImages.ContainsKey(_sheetNames[i]) ||
                _sheetComments.ContainsKey(_sheetNames[i]))
            {
                generated.Add($"xl/worksheets/_rels/sheet{i + 1}.xml.rels");
            }
            // 图片 drawing 和 rels
            if (_sheetImages.TryGetValue(_sheetNames[i], out var imgs) && imgs.Count > 0)
            {
                generated.Add($"xl/drawings/drawing{i + 1}.xml");
                generated.Add($"xl/drawings/_rels/drawing{i + 1}.xml.rels");
            }
            // 批注 comments 和 vml
            if (_sheetComments.TryGetValue(_sheetNames[i], out var cmts) && cmts.Count > 0)
            {
                generated.Add($"xl/comments{i + 1}.xml");
                generated.Add($"xl/drawings/vmlDrawing{i + 1}.vml");
            }
        }

        foreach (var kv in _otherParts)
        {
            if (generated.Contains(kv.Key)) continue;
            // 跳过媒体文件（Writer 已写入）
            if (kv.Key.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase)) continue;

            using var e = za.CreateEntry(kv.Key).Open();
            e.Write(kv.Value, 0, kv.Value.Length);
        }
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
            if (f.Strike) sw.Write("<strike/>");
            if (!f.VerticalAlign.IsNullOrEmpty()) sw.Write($"<vertAlign val=\"{f.VerticalAlign}\"/>");
            if (f.Size > 0) sw.Write($"<sz val=\"{f.Size}\"/>");
            WriteColorXml(sw, f.Color);
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
            else if (f.PatternType == "solid")
                sw.Write($"<patternFill patternType=\"solid\"><fgColor {FormatColorAttr(f.BgColor)}/></patternFill>");
            else if (f.PatternType == "gradient" && f.GradientType != null)
            {
                sw.Write($"<gradientFill {(f.GradientType == "radial" ? "type=\"path\"" : "type=\"linear\"")} degree=\"{(f.GradientType == "radial" ? 0 : 90)}\">");
                sw.Write($"<stop position=\"0\"><color {FormatColorAttr(f.GradientColor1)}/></stop>");
                sw.Write($"<stop position=\"1\"><color {FormatColorAttr(f.GradientColor2)}/></stop>");
                sw.Write("</gradientFill>");
            }
            else if (f.PatternType == "pattern" && f.PatternTypeName != null)
            {
                sw.Write($"<patternFill patternType=\"{f.PatternTypeName}\">");
                if (!f.PatternFgColor.IsNullOrEmpty()) sw.Write($"<fgColor {FormatColorAttr(f.PatternFgColor)}/>");
                if (!f.BgColor.IsNullOrEmpty()) sw.Write($"<bgColor {FormatColorAttr(f.BgColor)}/>");
                sw.Write("</patternFill>");
            }
            else
                sw.Write($"<patternFill patternType=\"solid\"><fgColor {FormatColorAttr(f.BgColor)}/></patternFill>");
            sw.Write("</fill>");
        }
        sw.Write("</fills>");

        // borders
        sw.Write($"<borders count=\"{_borders.Count}\">");
        foreach (var b in _borders)
        {
            var hasAny = b.Left != ExcelCellBorderStyle.None || b.Right != ExcelCellBorderStyle.None ||
                         b.Top  != ExcelCellBorderStyle.None || b.Bottom != ExcelCellBorderStyle.None ||
                         b.Diagonal != ExcelCellBorderStyle.None;
            if (!hasAny)
            {
                sw.Write("<border><left/><right/><top/><bottom/><diagonal/></border>");
            }
            else
            {
                sw.Write("<border>");
                WriteBorderSide(sw, "left",   b.Left,   b.LeftColor);
                WriteBorderSide(sw, "right",  b.Right,  b.RightColor);
                WriteBorderSide(sw, "top",    b.Top,    b.TopColor);
                WriteBorderSide(sw, "bottom", b.Bottom, b.BottomColor);
                if (b.Diagonal != ExcelCellBorderStyle.None)
                {
                    var diagStyle = b.Diagonal switch
                    {
                        ExcelCellBorderStyle.Thin => "thin",
                        ExcelCellBorderStyle.Medium => "medium",
                        ExcelCellBorderStyle.Thick => "thick",
                        ExcelCellBorderStyle.Dashed => "dashed",
                        ExcelCellBorderStyle.Dotted => "dotted",
                        ExcelCellBorderStyle.DoubleLine => "double",
                        _ => "thin"
                    };
                    sw.Write($"<diagonal style=\"{diagStyle}\">");
                    if (!b.DiagonalColor.IsNullOrEmpty()) sw.Write($"<color rgb=\"FF{b.DiagonalColor}\"/>");
                    sw.Write("</diagonal>");
                }
                else
                    sw.Write("<diagonal/>");
                sw.Write("</border>");
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
            var needAlignment = xf.HAlign != ExcelHorizontalAlignment.General || xf.VAlign != ExcelVerticalAlignment.Top ||
                               xf.WrapText || xf.TextRotation != 0 || xf.Indent > 0 || xf.ShrinkToFit;
            if (needAlignment)
            {
                sw.Write(" applyAlignment=\"1\"><alignment");
                if (xf.HAlign != ExcelHorizontalAlignment.General) sw.Write($" horizontal=\"{xf.HAlign.ToString().ToLower()}\"");
                if (xf.VAlign != ExcelVerticalAlignment.Top) sw.Write($" vertical=\"{xf.VAlign.ToString().ToLower()}\"");
                if (xf.WrapText) sw.Write(" wrapText=\"1\"");
                if (xf.TextRotation != 0) sw.Write($" textRotation=\"{xf.TextRotation}\"");
                if (xf.Indent > 0) sw.Write($" indent=\"{xf.Indent}\"");
                if (xf.ShrinkToFit) sw.Write(" shrinkToFit=\"1\"");
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
                if (cf.Type < ExcelConditionalFormatType.DataBar ||
                    (cf.Type == ExcelConditionalFormatType.Expression && !cf.Color.IsNullOrEmpty())) totalDxf++;
            }
        }
        if (totalDxf > 0)
        {
            sw.Write($"<dxfs count=\"{totalDxf}\">");
            foreach (var kv in _sheetCondFormats)
            {
                foreach (var cf in kv.Value)
                {
                    if (cf.Type >= ExcelConditionalFormatType.DataBar &&
                        !(cf.Type == ExcelConditionalFormatType.Expression && !cf.Color.IsNullOrEmpty())) continue;
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