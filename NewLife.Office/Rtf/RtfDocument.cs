using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NewLife.Office.Rtf;

/// <summary>RTF 文档对象模型</summary>
/// <remarks>
/// 表示解析后的 RTF 文档，提供从文件/字符串解析，
/// 以及访问段落、表格和提取纯文本的能力。
/// </remarks>
public sealed class RtfDocument
{
    #region 属性
    /// <summary>顶层块：RTF 段落与表格的混合列表（RtfParagraph | RtfTable）</summary>
    public List<Object> Blocks { get; } = [];

    /// <summary>文档标题</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>作者</summary>
    public String Author { get; set; } = String.Empty;

    /// <summary>主题</summary>
    public String Subject { get; set; } = String.Empty;

    /// <summary>所有段落（包括表格内段落）</summary>
    public IEnumerable<RtfParagraph> Paragraphs
    {
        get
        {
            foreach (var block in Blocks)
            {
                if (block is RtfParagraph para)
                    yield return para;
                else if (block is RtfTable table)
                    foreach (var row in table.Rows)
                        foreach (var cell in row.Cells)
                            foreach (var p in cell.Paragraphs)
                                yield return p;
            }
        }
    }

    /// <summary>所有顶层表格</summary>
    public IEnumerable<RtfTable> Tables
    {
        get
        {
            foreach (var block in Blocks)
                if (block is RtfTable t) yield return t;
        }
    }
    #endregion

    #region 解析
    /// <summary>从文件解析 RTF 文档</summary>
    /// <param name="path">文件路径</param>
    /// <returns>已解析的文档对象</returns>
    public static RtfDocument ParseFile(String path)
    {
        var text = File.ReadAllText(path, Encoding.Default);
        return Parse(text);
    }

    /// <summary>从流解析 RTF 文档</summary>
    /// <param name="stream">输入流</param>
    /// <returns>已解析的文档对象</returns>
    public static RtfDocument Parse(Stream stream)
    {
        using var reader = new StreamReader(stream, Encoding.Default, detectEncodingFromByteOrderMarks: true);
        return Parse(reader.ReadToEnd());
    }

    /// <summary>从字符串解析 RTF 文档</summary>
    /// <param name="rtf">RTF 文本</param>
    /// <returns>已解析的文档对象</returns>
    public static RtfDocument Parse(String rtf)
    {
        if (String.IsNullOrEmpty(rtf)) return new RtfDocument();
        return new RtfReader().Read(rtf);
    }
    #endregion

    #region 输出
    /// <summary>提取全文纯文本（段落间以换行分隔）</summary>
    /// <returns>纯文本字符串</returns>
    public String GetPlainText()
    {
        var sb = new StringBuilder();
        foreach (var block in Blocks)
        {
            if (block is RtfParagraph para)
            {
                sb.AppendLine(para.GetPlainText());
            }
            else if (block is RtfTable table)
            {
                foreach (var row in table.Rows)
                {
                    sb.AppendLine(row.GetPlainText());
                }
            }
        }
        return sb.ToString();
    }
    #endregion
}
