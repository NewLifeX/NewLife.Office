using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office.Ods;

/// <summary>ODS 电子表格读取器</summary>
/// <remarks>
/// 读取 OpenDocument Spreadsheet（.ods）格式文件，提取工作表名称及单元格数据。
/// ODS 是基于 ZIP 的 XML 格式，核心内容在 content.xml 中。
/// </remarks>
public sealed class OdsReader
{
    #region 常量
    private const String NsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private const String NsTable  = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private const String NsText   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    #endregion

    #region 方法 — 读取
    /// <summary>从文件路径读取所有工作表数据</summary>
    /// <param name="path">ODS 文件路径</param>
    /// <returns>工作表列表</returns>
    public static List<OdsSheet> ReadFile(String path)
    {
        using var fs = File.OpenRead(path);
        return Read(fs);
    }

    /// <summary>从流读取所有工作表数据</summary>
    /// <param name="stream">ODS 输入流</param>
    /// <returns>工作表列表</returns>
    public static List<OdsSheet> Read(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var entry = zip.GetEntry("content.xml");
        if (entry == null) return [];

        using var contentStream = entry.Open();
        return ParseContentXml(contentStream);
    }

    /// <summary>从文件路径读取第一张工作表的数据行</summary>
    /// <param name="path">ODS 文件路径</param>
    /// <returns>行列表，每行为字符串数组</returns>
    public static List<String[]> ReadRows(String path)
    {
        var sheets = ReadFile(path);
        return sheets.Count > 0 ? sheets[0].Rows : [];
    }

    /// <summary>从流读取第一张工作表的数据行</summary>
    /// <param name="stream">ODS 输入流</param>
    /// <returns>行列表，每行为字符串数组</returns>
    public static List<String[]> ReadRows(Stream stream)
    {
        var sheets = Read(stream);
        return sheets.Count > 0 ? sheets[0].Rows : [];
    }
    #endregion

    #region XML 解析
    private static List<OdsSheet> ParseContentXml(Stream xmlStream)
    {
        var result = new List<OdsSheet>();
        var settings = new XmlReaderSettings { IgnoreWhitespace = false, IgnoreComments = true };
        using var reader = XmlReader.Create(xmlStream, settings);

        OdsSheet? currentSheet = null;
        List<String>? currentRow = null;
        StringBuilder? cellText = null;
        var inTextP = false;
        var cellRepeat = 1;
        var rowRepeat = 1;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                var ns = reader.NamespaceURI;
                var name = reader.LocalName;

                if (ns == NsTable && name == "table")
                {
                    var sheetName = reader.GetAttribute("name", NsTable) ?? "";
                    currentSheet = new OdsSheet { Name = sheetName };
                    result.Add(currentSheet);
                    continue;
                }

                if (ns == NsTable && name == "table-row")
                {
                    rowRepeat = GetRepeatAttr(reader, NsTable, "number-rows-repeated");
                    currentRow = [];
                    continue;
                }

                if (ns == NsTable && name == "table-cell")
                {
                    cellRepeat = GetRepeatAttr(reader, NsTable, "number-columns-repeated");
                    cellText = new StringBuilder();
                    inTextP = false;
                    continue;
                }

                if (ns == NsText && name == "p")
                {
                    inTextP = true;
                    continue;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                var ns = reader.NamespaceURI;
                var name = reader.LocalName;

                if (ns == NsText && name == "p")
                {
                    inTextP = false;
                    continue;
                }

                if (ns == NsTable && name == "table-cell")
                {
                    if (currentRow != null && cellText != null)
                    {
                        var val = cellText.ToString();
                        for (var i = 0; i < cellRepeat; i++)
                            currentRow.Add(val);
                    }
                    cellText = null;
                    inTextP = false;
                    continue;
                }

                if (ns == NsTable && name == "table-row")
                {
                    if (currentSheet != null && currentRow != null)
                    {
                        var trimmedRow = TrimTrailingEmpty(currentRow);
                        if (trimmedRow != null || rowRepeat == 1)
                        {
                            var arr = trimmedRow ?? [];
                            for (var i = 0; i < rowRepeat; i++)
                                currentSheet.Rows.Add(arr);
                        }
                    }
                    currentRow = null;
                    rowRepeat = 1;
                    continue;
                }
            }
            else if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.SignificantWhitespace)
            {
                if (inTextP && cellText != null)
                    cellText.Append(reader.Value);
            }
        }

        return result;
    }

    private static Int32 GetRepeatAttr(XmlReader reader, String ns, String localName)
    {
        var val = reader.GetAttribute(localName, ns);
        return val != null && Int32.TryParse(val, out var n) ? n : 1;
    }

    private static String[]? TrimTrailingEmpty(List<String> row)
    {
        var end = row.Count - 1;
        while (end >= 0 && String.IsNullOrEmpty(row[end])) end--;
        if (end < 0) return null; // all empty
        var arr = new String[end + 1];
        for (var i = 0; i <= end; i++) arr[i] = row[i];
        return arr;
    }
    #endregion
}

/// <summary>ODS 工作表数据</summary>
public sealed class OdsSheet
{
    /// <summary>工作表名称</summary>
    public String Name { get; set; } = "";

    /// <summary>数据行列表（每行为字符串数组）</summary>
    public List<String[]> Rows { get; } = [];
}
