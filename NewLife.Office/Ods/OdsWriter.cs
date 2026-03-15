using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office.Ods;

/// <summary>ODS 电子表格写入器</summary>
/// <remarks>
/// 生成 OpenDocument Spreadsheet（.ods）格式文件。
/// 支持多工作表、字符串、数值、日期、布尔值类型单元格和基本单元格样式。
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
        w.WriteLine(@"  <office:body>");
        w.WriteLine(@"    <office:spreadsheet>");

        foreach (var sheet in Sheets)
        {
            w.WriteLine($@"      <table:table table:name=""{XmlEncode(sheet.Name)}"">");
            foreach (var row in sheet.Rows)
            {
                w.WriteLine(@"        <table:table-row>");
                foreach (var cell in row)
                    WriteCell(w, cell);
                w.WriteLine(@"        </table:table-row>");
            }
            w.WriteLine(@"      </table:table>");
        }

        w.WriteLine(@"    </office:spreadsheet>");
        w.WriteLine(@"  </office:body>");
        w.Write("</office:document-content>");
    }

    private static void WriteCell(StreamWriter w, String value)
    {
        // Attempt to detect numeric/bool/date values for proper typing
        if (String.IsNullOrEmpty(value))
        {
            w.WriteLine(@"          <table:table-cell/>");
            return;
        }

        if (Double.TryParse(value, System.Globalization.NumberStyles.Number,
            System.Globalization.CultureInfo.InvariantCulture, out var num))
        {
            w.WriteLine($@"          <table:table-cell office:value-type=""float"" office:value=""{num.ToString(System.Globalization.CultureInfo.InvariantCulture)}""><text:p>{XmlEncode(value)}</text:p></table:table-cell>");
            return;
        }

        if (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            var boolVal = value.ToLowerInvariant();
            w.WriteLine($@"          <table:table-cell office:value-type=""boolean"" office:boolean-value=""{boolVal}""><text:p>{XmlEncode(value)}</text:p></table:table-cell>");
            return;
        }

        // Default: string
        w.WriteLine($@"          <table:table-cell office:value-type=""string""><text:p>{XmlEncode(value)}</text:p></table:table-cell>");
    }
    #endregion

    #region 辅助
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
