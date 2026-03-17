using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;

namespace NewLife.Office;

/// <summary>轻量级 Excel 数据透视表生成器</summary>
/// <remarks>
/// 独立于 ExcelWriter，接受源数据（DataTable 或对象集合），
/// 在指定目标工作表生成含 PivotCache + PivotTable 的 xlsx 文件。
/// 当前支持行字段、数据字段、简单汇总，适用于无需动态刷新的静态报表场景。
/// </remarks>
public class ExcelPivotBuilder
{
    #region 属性
    /// <summary>透视表名称</summary>
    public String Name { get; set; } = "PivotTable1";

    /// <summary>源数据工作表名称</summary>
    public String SourceSheet { get; set; } = "Sheet1";

    /// <summary>透视表放置的目标工作表名称</summary>
    public String TargetSheet { get; set; } = "Pivot";

    /// <summary>透视表左上角单元格（如 "A1"）</summary>
    public String TargetCell { get; set; } = "A1";

    /// <summary>字段配置列表</summary>
    public List<PivotField> Fields { get; } = [];

    /// <summary>文本编码</summary>
    public Encoding Encoding { get; set; } = Encoding.UTF8;
    #endregion

    #region 数据
    private String[]? _headers;
    private List<Object?[]>? _rows;
    #endregion

    #region 构造
    /// <summary>实例化透视表生成器</summary>
    public ExcelPivotBuilder() { }
    #endregion

    #region 方法
    /// <summary>设置源数据（二维数组形式，首行为表头）</summary>
    /// <param name="headers">列头数组</param>
    /// <param name="rows">数据行集合</param>
    public void SetSourceData(String[] headers, IEnumerable<Object?[]> rows)
    {
        _headers = headers ?? throw new ArgumentNullException(nameof(headers));
        _rows = rows?.ToList() ?? throw new ArgumentNullException(nameof(rows));
    }

    /// <summary>添加行字段</summary>
    /// <param name="fieldName">字段名称</param>
    public ExcelPivotBuilder AddRowField(String fieldName)
    {
        Fields.Add(new PivotField { Name = fieldName, IsRowField = true });
        return this;
    }

    /// <summary>添加列字段</summary>
    /// <param name="fieldName">字段名称</param>
    public ExcelPivotBuilder AddColumnField(String fieldName)
    {
        Fields.Add(new PivotField { Name = fieldName, IsColumnField = true });
        return this;
    }

    /// <summary>添加数据字段</summary>
    /// <param name="fieldName">字段名称</param>
    /// <param name="func">汇总函数</param>
    /// <param name="caption">显示标题（可空，空时自动生成）</param>
    public ExcelPivotBuilder AddDataField(String fieldName, PivotSummaryFunction func = PivotSummaryFunction.Sum, String? caption = null)
    {
        Fields.Add(new PivotField { Name = fieldName, IsDataField = true, SummaryFunction = func, Caption = caption });
        return this;
    }

    /// <summary>生成包含数据源和透视表的 xlsx 文件</summary>
    /// <param name="outputPath">输出文件路径</param>
    public void Save(String outputPath)
    {
        using var fs = new FileStream(outputPath.EnsureDirectory(true).GetFullPath(), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
        Save(fs);
    }

    /// <summary>生成包含数据源和透视表的 xlsx 到流</summary>
    /// <param name="stream">可写目标流</param>
    public void Save(Stream stream)
    {
        if (_headers == null || _rows == null)
            throw new InvalidOperationException("尚未设置源数据，请先调用 SetSourceData()");

        using var za = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true, entryNameEncoding: Encoding);

        var dataSheetIdx = 0;
        var pivotSheetIdx = 1;
        var cacheId = 1;

        // _rels/.rels
        WriteEntry(za, "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");

        // [Content_Types].xml
        WriteEntry(za, "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            $"<Override PartName=\"/xl/worksheets/sheet{dataSheetIdx + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            $"<Override PartName=\"/xl/worksheets/sheet{pivotSheetIdx + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            $"<Override PartName=\"/xl/pivotCache/pivotCacheDefinition{cacheId}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>" +
            $"<Override PartName=\"/xl/pivotCache/pivotCacheRecords{cacheId}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml\"/>" +
            $"<Override PartName=\"/xl/pivotTables/pivotTable{cacheId}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
            "</Types>");

        // workbook.xml
        WriteEntry(za, "xl/workbook.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<sheets>" +
            $"<sheet name=\"{Escape(SourceSheet)}\" sheetId=\"{dataSheetIdx + 1}\" r:id=\"rId{dataSheetIdx + 1}\"/>" +
            $"<sheet name=\"{Escape(TargetSheet)}\" sheetId=\"{pivotSheetIdx + 1}\" r:id=\"rId{pivotSheetIdx + 1}\"/>" +
            "</sheets></workbook>");

        // workbook rels
        WriteEntry(za, "xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId{dataSheetIdx + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{dataSheetIdx + 1}.xml\"/>" +
            $"<Relationship Id=\"rId{pivotSheetIdx + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{pivotSheetIdx + 1}.xml\"/>" +
            $"<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            $"<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition{cacheId}.xml\"/>" +
            "</Relationships>");

        // styles.xml（最小化）
        WriteEntry(za, "xl/styles.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
            "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
            "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>" +
            "</styleSheet>");

        // 源数据工作表
        WriteDataSheet(za, dataSheetIdx);

        // 透视表工作表（空白，透视表叠加在上）
        WriteEntry(za, $"xl/worksheets/sheet{pivotSheetIdx + 1}.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<sheetData/></worksheet>");

        // pivot sheet rels
        WriteEntry(za, $"xl/worksheets/_rels/sheet{pivotSheetIdx + 1}.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable{cacheId}.xml\"/>" +
            "</Relationships>");

        // pivotCacheDefinition
        WritePivotCacheDefinition(za, dataSheetIdx, cacheId);

        // pivotCacheRecords
        WritePivotCacheRecords(za, cacheId);

        // pivotTable
        WritePivotTable(za, cacheId);
    }

    private void WriteDataSheet(ZipArchive za, Int32 sheetIdx)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        sb.Append("<sheetData>");

        // header row
        sb.Append("<row r=\"1\">");
        for (var c = 0; c < _headers!.Length; c++)
        {
            var col = GetColName(c) + "1";
            sb.Append($"<c r=\"{col}\" t=\"inlineStr\"><is><t>{Escape(_headers[c])}</t></is></c>");
        }
        sb.Append("</row>");

        // data rows
        for (var r = 0; r < _rows!.Count; r++)
        {
            var row = _rows[r];
            sb.Append($"<row r=\"{r + 2}\">");
            for (var c = 0; c < row.Length && c < _headers.Length; c++)
            {
                if (row[c] == null) continue;
                var col = GetColName(c) + (r + 2);
                var val = row[c]!;
                if (val is String s)
                    sb.Append($"<c r=\"{col}\" t=\"inlineStr\"><is><t>{Escape(s)}</t></is></c>");
                else
                    sb.Append($"<c r=\"{col}\"><v>{SecurityElement.Escape(Convert.ToString(val, CultureInfo.InvariantCulture))}</v></c>");
            }
            sb.Append("</row>");
        }

        sb.Append("</sheetData></worksheet>");
        WriteEntry(za, $"xl/worksheets/sheet{sheetIdx + 1}.xml", sb.ToString());
    }

    private void WritePivotCacheDefinition(ZipArchive za, Int32 dataSheetIdx, Int32 cacheId)
    {
        var totalRows = (_rows?.Count ?? 0) + 1; // data + header
        var lastCol = _headers != null && _headers.Length > 0 ? GetColName(_headers.Length - 1) : "A";
        var sourceRef = $"A1:{lastCol}{totalRows}";

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
        sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
        sb.Append($"refreshOnLoad=\"1\" recordCount=\"{_rows?.Count ?? 0}\" r:id=\"rCR1\">");
        sb.Append("<cacheSource type=\"worksheet\">");
        sb.Append($"<worksheetSource ref=\"{sourceRef}\" sheet=\"{Escape(SourceSheet)}\"/>");
        sb.Append("</cacheSource>");
        sb.Append($"<cacheFields count=\"{_headers!.Length}\">");
        for (var h = 0; h < _headers.Length; h++)
        {
            sb.Append($"<cacheField name=\"{Escape(_headers[h])}\" numFmtId=\"0\">");
            // 收集唯一值
            var uniques = _rows!.Select(r => r.Length > h ? r[h] : null)
                                .Where(v => v != null)
                                .Select(v => Convert.ToString(v, CultureInfo.InvariantCulture)!)
                                .Distinct()
                                .OrderBy(x => x)
                                .ToList();
            if (uniques.Count <= 1000) // 过多则不展开
            {
                sb.Append($"<sharedItems count=\"{uniques.Count}\">");
                foreach (var u in uniques)
                    sb.Append($"<s v=\"{Escape(u)}\"/>");
                sb.Append("</sharedItems>");
            }
            else
            {
                sb.Append("<sharedItems containsMixedTypes=\"1\"/>");
            }
            sb.Append("</cacheField>");
        }
        sb.Append("</cacheFields></pivotCacheDefinition>");

        WriteEntry(za, $"xl/pivotCache/pivotCacheDefinition{cacheId}.xml", sb.ToString());

        // pivotCacheDefinition rels
        WriteEntry(za, $"xl/pivotCache/_rels/pivotCacheDefinition{cacheId}.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rCR1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" Target=\"pivotCacheRecords{cacheId}.xml\"/>" +
            "</Relationships>");
    }

    private void WritePivotCacheRecords(ZipArchive za, Int32 cacheId)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{_rows!.Count}\">");
        foreach (var row in _rows)
        {
            sb.Append("<r>");
            for (var c = 0; c < _headers!.Length; c++)
            {
                var v = c < row.Length ? row[c] : null;
                if (v == null || v is String)
                    sb.Append($"<s v=\"{Escape(v?.ToString() ?? String.Empty)}\"/>");
                else
                    sb.Append($"<n v=\"{SecurityElement.Escape(Convert.ToString(v, CultureInfo.InvariantCulture))}\"/>");
            }
            sb.Append("</r>");
        }
        sb.Append("</pivotCacheRecords>");
        WriteEntry(za, $"xl/pivotCache/pivotCacheRecords{cacheId}.xml", sb.ToString());
    }

    private void WritePivotTable(ZipArchive za, Int32 cacheId)
    {
        var rowFields = Fields.Where(f => f.IsRowField).ToList();
        var colFields = Fields.Where(f => f.IsColumnField).ToList();
        var dataFields = Fields.Where(f => f.IsDataField).ToList();

        // 字段索引映射
        var fieldIndex = new Dictionary<String, Int32>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < _headers!.Length; i++) fieldIndex[_headers[i]] = i;

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
        sb.Append($"name=\"{Escape(Name)}\" cacheId=\"{cacheId}\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" ");
        sb.Append("applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" ");
        sb.Append("dataCaption=\"Values\" updatedVersion=\"3\" minRefreshableVersion=\"3\" showCalcMbrs=\"0\" useAutoFormatting=\"1\" >");
        sb.Append($"<location ref=\"{TargetCell}\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\"/>");
        sb.Append($"<pivotFields count=\"{_headers.Length}\">");
        for (var h = 0; h < _headers.Length; h++)
        {
            var field = Fields.FirstOrDefault(f => String.Equals(f.Name, _headers[h], StringComparison.OrdinalIgnoreCase));
            if (field?.IsRowField == true)
            {
                sb.Append("<pivotField axis=\"axisRow\" showAll=\"0\"><items count=\"1\"><item t=\"default\"/></items></pivotField>");
            }
            else if (field?.IsDataField == true)
            {
                sb.Append("<pivotField dataField=\"1\" showAll=\"0\"><items count=\"0\"/></pivotField>");
            }
            else
            {
                sb.Append("<pivotField showAll=\"0\"/>");
            }
        }
        sb.Append("</pivotFields>");

        if (rowFields.Count > 0)
        {
            sb.Append($"<rowFields count=\"{rowFields.Count}\">");
            foreach (var rf in rowFields)
            {
                if (fieldIndex.TryGetValue(rf.Name, out var idx)) sb.Append($"<field x=\"{idx}\"/>");
            }
            sb.Append("</rowFields>");
        }

        if (colFields.Count > 0)
        {
            sb.Append($"<colFields count=\"{colFields.Count}\">");
            foreach (var cf in colFields)
            {
                if (fieldIndex.TryGetValue(cf.Name, out var idx)) sb.Append($"<field x=\"{idx}\"/>");
            }
            sb.Append("</colFields>");
        }

        if (dataFields.Count > 0)
        {
            sb.Append($"<dataFields count=\"{dataFields.Count}\">");
            foreach (var df in dataFields)
            {
                if (!fieldIndex.TryGetValue(df.Name, out var idx)) continue;
                var caption = Escape(df.Caption ?? $"{FuncName(df.SummaryFunction)} of {df.Name}");
                var fn = FuncName(df.SummaryFunction).ToLower();
                sb.Append($"<dataField name=\"{caption}\" fld=\"{idx}\" subtotal=\"{fn}\" baseField=\"0\" baseItem=\"0\"/>");
            }
            sb.Append("</dataFields>");
        }

        sb.Append("</pivotTableDefinition>");
        WriteEntry(za, $"xl/pivotTables/pivotTable{cacheId}.xml", sb.ToString());

        // pivotTable rels
        WriteEntry(za, $"xl/pivotTables/_rels/pivotTable{cacheId}.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition{cacheId}.xml\"/>" +
            "</Relationships>");
    }
    #endregion

    #region 辅助
    private void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), Encoding);
        sw.Write(content);
    }

    private static String Escape(String? s) =>
        s == null ? String.Empty : (SecurityElement.Escape(s) ?? s);

    private static String GetColName(Int32 index)
    {
        index++;
        var sb = new StringBuilder();
        while (index > 0)
        {
            var mod = (index - 1) % 26;
            sb.Insert(0, (Char)('A' + mod));
            index = (index - 1) / 26;
        }
        return sb.ToString();
    }

    private static String FuncName(PivotSummaryFunction f) => f switch
    {
        PivotSummaryFunction.Count => "Count",
        PivotSummaryFunction.Average => "Average",
        PivotSummaryFunction.Max => "Max",
        PivotSummaryFunction.Min => "Min",
        _ => "Sum",
    };
    #endregion
}
