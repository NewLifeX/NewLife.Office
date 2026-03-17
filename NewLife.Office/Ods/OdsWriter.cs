using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml;

namespace NewLife.Office.Ods;

/// <summary>ODS 电子表格写入器</summary>
/// <remarks>
/// 生成 OpenDocument Spreadsheet（.ods）格式文件。
/// 支持多工作表、字符串、数值、日期、布尔值、公式类型单元格、基本样式，以及对象集合导出。
/// </remarks>
public sealed class OdsWriter
{
    #region 属性
    /// <summary>文档标题</summary>
    public String Title { get; set; } = "";

    /// <summary>文档作者</summary>
    public String Author { get; set; } = "";

    /// <summary>工作表列表</summary>
    public List<OdsSheet> Sheets { get; } = [];
    #endregion

    #region 方法 — 添加数据
    /// <summary>添加工作表（字符串二维数据，每行为 String 数组）</summary>
    /// <param name="name">工作表名称</param>
    /// <param name="rows">行数据（每行为 String[]）</param>
    /// <returns>当前写入器（链式调用）</returns>
    public OdsWriter AddSheet(String name, IEnumerable<String[]> rows)
    {
        var sheet = new OdsSheet { Name = name };
        foreach (var row in rows)
        {
            sheet.Rows.Add(row ?? []);
        }
        Sheets.Add(sheet);
        return this;
    }

    /// <summary>添加工作表（字符串二维数据）</summary>
    /// <param name="name">工作表名称</param>
    /// <param name="rows">行数据</param>
    /// <returns>当前写入器（链式调用）</returns>
    public OdsWriter AddSheet(String name, IEnumerable<IEnumerable<String>> rows)
    {
        var sheet = new OdsSheet { Name = name };
        foreach (var row in rows)
        {
            var cells = new List<String>();
            foreach (var cell in row) cells.Add(cell ?? "");
            sheet.Rows.Add([.. cells]);
        }
        Sheets.Add(sheet);
        return this;
    }

    /// <summary>添加工作表对象</summary>
    /// <param name="sheet">工作表对象</param>
    /// <returns>当前写入器（链式调用）</returns>
    public OdsWriter AddSheet(OdsSheet sheet)
    {
        Sheets.Add(sheet);
        return this;
    }

    /// <summary>添加工作表（对象集合，使用反射读取属性名作列头）</summary>
    /// <typeparam name="T">列表元素类型</typeparam>
    /// <param name="name">工作表名称</param>
    /// <param name="items">数据集合</param>
    /// <returns>当前写入器（链式调用）</returns>
    public OdsWriter AddSheet<T>(String name, IEnumerable<T> items)
    {
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead).ToArray();

        var sheet = new OdsSheet { Name = name };

        // 生成表头行
        var headers = new String[props.Length];
        for (var i = 0; i < props.Length; i++)
        {
            var dn = props[i].GetCustomAttribute<DisplayNameAttribute>();
            if (dn != null) { headers[i] = dn.DisplayName; continue; }
            var desc = props[i].GetCustomAttribute<DescriptionAttribute>();
            if (desc != null) { headers[i] = desc.Description; continue; }
            headers[i] = props[i].Name;
        }
        sheet.Rows.Add(headers);

        // 数据行
        foreach (var item in items)
        {
            if (item == null) continue;
            var row = new String[props.Length];
            for (var i = 0; i < props.Length; i++)
            {
                var val = props[i].GetValue(item);
                row[i] = val == null ? "" : Convert.ToString(val, CultureInfo.InvariantCulture) ?? "";
            }
            sheet.Rows.Add(row);
        }

        Sheets.Add(sheet);
        return this;
    }
    #endregion

    #region 方法 — 保存
    /// <summary>保存到文件</summary>
    /// <param name="path">文件路径</param>
    public void Save(String path)
    {
        using var fs = File.Create(path);
        Save(fs);
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">输出流</param>
    public void Save(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteMimetype(zip);
        WriteManifest(zip);
        WriteMeta(zip);
        WriteStyles(zip);
        WriteContent(zip);
    }
    #endregion

    #region ZIP 内容写入
    private static void WriteMimetype(ZipArchive zip)
    {
        var entry = zip.CreateEntry("mimetype", CompressionLevel.NoCompression);
        using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        w.Write("application/vnd.oasis.opendocument.spreadsheet");
    }

    private static void WriteManifest(ZipArchive zip)
    {
        var entry = zip.CreateEntry("META-INF/manifest.xml");
        using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        w.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
        w.WriteLine(@"<manifest:manifest xmlns:manifest=""urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"">");
        w.WriteLine(@"  <manifest:file-entry manifest:full-path=""/"" manifest:media-type=""application/vnd.oasis.opendocument.spreadsheet""/>");
        w.WriteLine(@"  <manifest:file-entry manifest:full-path=""content.xml"" manifest:media-type=""text/xml""/>");
        w.WriteLine(@"  <manifest:file-entry manifest:full-path=""styles.xml"" manifest:media-type=""text/xml""/>");
        w.WriteLine(@"  <manifest:file-entry manifest:full-path=""meta.xml"" manifest:media-type=""text/xml""/>");
        w.Write("</manifest:manifest>");
    }

    private void WriteMeta(ZipArchive zip)
    {
        var entry = zip.CreateEntry("meta.xml");
        using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        w.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
        w.WriteLine(@"<office:document-meta xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0"" xmlns:meta=""urn:oasis:names:tc:opendocument:xmlns:meta:1.0"" xmlns:dc=""http://purl.org/dc/elements/1.1/"">");
        w.WriteLine(@"  <office:meta>");
        if (!String.IsNullOrEmpty(Title))
            w.WriteLine($"    <dc:title>{XmlEncode(Title)}</dc:title>");
        if (!String.IsNullOrEmpty(Author))
            w.WriteLine($"    <dc:creator>{XmlEncode(Author)}</dc:creator>");
        w.WriteLine(@"  </office:meta>");
        w.Write("</office:document-meta>");
    }

    private static void WriteStyles(ZipArchive zip)
    {
        var entry = zip.CreateEntry("styles.xml");
        using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        w.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?><office:document-styles xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0""></office:document-styles>");
    }

    private void WriteContent(ZipArchive zip)
    {
        // 收集所有唯一单元格样式，分配名称 ce1, ce2, ...
        var styleKeyMap = new Dictionary<String, String>();
        var styleList = new List<(String StyleName, OdsCellStyle Style)>();
        foreach (var sheet in Sheets)
        {
            foreach (var style in sheet.CellStyles.Values)
            {
                var key = GetStyleKey(style);
                if (!styleKeyMap.ContainsKey(key))
                {
                    var sname = $"ce{styleList.Count + 1}";
                    styleKeyMap[key] = sname;
                    styleList.Add((sname, style));
                }
            }
        }

        var entry = zip.CreateEntry("content.xml");
        using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        w.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
        w.WriteLine(@"<office:document-content");
        w.WriteLine(@"  xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0""");
        w.WriteLine(@"  xmlns:table=""urn:oasis:names:tc:opendocument:xmlns:table:1.0""");
        w.WriteLine(@"  xmlns:text=""urn:oasis:names:tc:opendocument:xmlns:text:1.0""");
        w.WriteLine(@"  xmlns:fo=""urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0""");
        w.WriteLine(@"  xmlns:style=""urn:oasis:names:tc:opendocument:xmlns:style:1.0""");
        w.WriteLine(@"  office:version=""1.2"">");

        // 写入自动样式块
        if (styleList.Count > 0)
        {
            w.WriteLine(@"  <office:automatic-styles>");
            foreach (var (sname, style) in styleList)
            {
                WriteStyleEntry(w, sname, style);
            }
            w.WriteLine(@"  </office:automatic-styles>");
        }

        w.WriteLine(@"  <office:body>");
        w.WriteLine(@"    <office:spreadsheet>");

        foreach (var sheet in Sheets)
        {
            w.WriteLine($@"      <table:table table:name=""{XmlEncode(sheet.Name)}"">");
            for (var ri = 0; ri < sheet.Rows.Count; ri++)
            {
                var row = sheet.Rows[ri];
                w.WriteLine(@"        <table:table-row>");
                for (var ci = 0; ci < row.Length; ci++)
                {
                    String? sname = null;
                    if (sheet.CellStyles.TryGetValue((ri, ci), out var cellStyle))
                        sname = styleKeyMap[GetStyleKey(cellStyle)];
                    WriteCell(w, row[ci], sname);
                }
                w.WriteLine(@"        </table:table-row>");
            }
            w.WriteLine(@"      </table:table>");
        }

        w.WriteLine(@"    </office:spreadsheet>");
        w.WriteLine(@"  </office:body>");
        w.Write("</office:document-content>");
    }

    private static void WriteStyleEntry(StreamWriter w, String styleName, OdsCellStyle style)
    {
        w.WriteLine($@"    <style:style style:name=""{styleName}"" style:family=""table-cell"">");
        var hasTxtProp = style.FontBold || style.FontItalic || style.FontSize > 0 || !String.IsNullOrEmpty(style.FontColor);
        if (hasTxtProp)
        {
            var sb = new StringBuilder();
            if (style.FontBold) sb.Append(@" fo:font-weight=""bold""");
            if (style.FontItalic) sb.Append(@" fo:font-style=""italic""");
            if (style.FontSize > 0) sb.Append($@" fo:font-size=""{style.FontSize:F1}pt""");
            if (!String.IsNullOrEmpty(style.FontColor)) sb.Append($@" fo:color=""{style.FontColor}""");
            w.WriteLine($@"      <style:text-properties{sb}/>");
        }
        if (!String.IsNullOrEmpty(style.BackgroundColor))
            w.WriteLine($@"      <style:table-cell-properties fo:background-color=""{style.BackgroundColor}""/>");
        if (!String.IsNullOrEmpty(style.HAlign))
            w.WriteLine($@"      <style:paragraph-properties fo:text-align=""{style.HAlign}""/>");
        w.WriteLine(@"    </style:style>");
    }

    private static void WriteCell(StreamWriter w, String value, String? styleName = null)
    {
        var sAttr = styleName != null ? $@" table:style-name=""{styleName}""" : String.Empty;
        if (String.IsNullOrEmpty(value))
        {
            w.WriteLine($@"          <table:table-cell{sAttr}/>");
            return;
        }

        // 公式：以 = 开头
        if (value.Length > 1 && value[0] == '=')
        {
            // OpenDocument 公式前缀 of:
            var formula = "of:" + value;
            w.WriteLine($@"          <table:table-cell{sAttr} table:formula=""{XmlEncode(formula)}"" office:value-type=""formula""><text:p>{XmlEncode(value)}</text:p></table:table-cell>");
            return;
        }

        if (Double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var num))
        {
            w.WriteLine($@"          <table:table-cell{sAttr} office:value-type=""float"" office:value=""{num.ToString(CultureInfo.InvariantCulture)}""><text:p>{XmlEncode(value)}</text:p></table:table-cell>");
            return;
        }

        if (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            var boolVal = value.ToLowerInvariant();
            w.WriteLine($@"          <table:table-cell{sAttr} office:value-type=""boolean"" office:boolean-value=""{boolVal}""><text:p>{XmlEncode(value)}</text:p></table:table-cell>");
            return;
        }

        // 默认字符串
        w.WriteLine($@"          <table:table-cell{sAttr} office:value-type=""string""><text:p>{XmlEncode(value)}</text:p></table:table-cell>");
    }
    #endregion

    #region 辅助
    private static String GetStyleKey(OdsCellStyle s) =>
        $"{(s.FontBold ? 1 : 0)}|{(s.FontItalic ? 1 : 0)}|{s.FontSize:F1}|{s.FontColor ?? ""}|{s.BackgroundColor ?? ""}|{s.HAlign ?? ""}";

    private static String XmlEncode(String text)
    {
        if (String.IsNullOrEmpty(text)) return text;
        return text.Replace("&", "&amp;")
                   .Replace("<", "&lt;")
                   .Replace(">", "&gt;")
                   .Replace("\"", "&quot;")
                   .Replace("'", "&apos;");
    }
    #endregion
}
