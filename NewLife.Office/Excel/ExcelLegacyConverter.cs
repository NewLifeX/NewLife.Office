namespace NewLife.Office.Excel;

/// <summary>旧格式 xls（BIFF8）读取结果到 xlsx 工作簿模型的转换器</summary>
/// <remarks>
/// 供统一 AI 接口复用：xls 的结构快照、区域读取、搜索与"xls → xlsx"转换都需要 <see cref="ExcelDocument"/> 模型，
/// 而 <see cref="BiffReader"/> 只提供按行读取与零散的样式/合并/批注接口，故集中在此转换。
/// <para>转换范围：单元格值、单元格样式、合并单元格、列宽、批注、页面设置与定义名；图片、超链接与公式文本不转换。</para>
/// </remarks>
public static class ExcelLegacyConverter
{
    /// <summary>把 BIFF 读取器转换为工作簿模型</summary>
    /// <param name="reader">xls 读取器</param>
    /// <returns>工作簿模型</returns>
    public static ExcelDocument ToDocument(BiffReader reader)
    {
        if (reader == null) throw new ArgumentNullException(nameof(reader));

        var document = new ExcelDocument();
        foreach (var name in reader.SheetNames)
        {
            var sheet = new Worksheet { Name = name };

            // ReadSheet 按稠密行输出（含中间空行），行号即数组下标，无需 ActualRowNumbers
            foreach (var row in reader.ReadSheet(name)) sheet.Rows.Add(row);

            foreach (var kv in reader.ReadCellFormats(name)) sheet.CellFormats[kv.Key] = kv.Value;
            foreach (var merge in reader.GetMerges(name)) sheet.Merges.Add(merge);
            foreach (var kv in reader.GetComments(name)) sheet.Comments[kv.Key] = kv.Value;

            // 列宽由最近一次 ReadSheet 解析得到，必须在读取该表之后取
            foreach (var kv in reader.GetColumnWidths()) sheet.ColumnWidths[kv.Key] = kv.Value;

            var setup = reader.GetPageSetup(name);
            sheet.Orientation = setup.Landscape ? PageOrientation.Landscape : PageOrientation.Portrait;
            if (setup.PaperSize > 0) sheet.PaperSize = (PaperSize)setup.PaperSize;
            sheet.HeaderText = setup.Header;
            sheet.FooterText = setup.Footer;

            document.Worksheets.Add(sheet);
        }

        foreach (var kv in reader.GetDefinedNames()) document.DefinedNames[kv.Key] = kv.Value;

        return document;
    }
}
