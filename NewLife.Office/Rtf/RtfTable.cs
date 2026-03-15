using System;
using System.Collections.Generic;

namespace NewLife.Office.Rtf;

/// <summary>RTF 表格单元格</summary>
public sealed class RtfTableCell
{
    #region 属性
    /// <summary>单元格段落列表</summary>
    public List<RtfParagraph> Paragraphs { get; } = [];

    /// <summary>单元格右边界位置（twips，相对页面左边距）</summary>
    public Int32 RightBoundary { get; set; }
    #endregion

    #region 方法
    /// <summary>获取单元格纯文本</summary>
    /// <returns>纯文本字符串</returns>
    public String GetPlainText()
    {
        var sb = new System.Text.StringBuilder();
        for (var i = 0; i < Paragraphs.Count; i++)
        {
            if (i > 0) sb.Append(" ");
            sb.Append(Paragraphs[i].GetPlainText());
        }
        return sb.ToString();
    }

    /// <inheritdoc/>
    public override String ToString() => GetPlainText();
    #endregion
}

/// <summary>RTF 表格行</summary>
public sealed class RtfTableRow
{
    #region 属性
    /// <summary>单元格列表</summary>
    public List<RtfTableCell> Cells { get; } = [];
    #endregion

    #region 方法
    /// <summary>获取行纯文本（Tab 分隔）</summary>
    /// <returns>纯文本字符串</returns>
    public String GetPlainText()
    {
        var parts = new List<String>();
        foreach (var cell in Cells) parts.Add(cell.GetPlainText());
        return String.Join("\t", parts);
    }

    /// <inheritdoc/>
    public override String ToString() => GetPlainText();
    #endregion
}

/// <summary>RTF 表格</summary>
public sealed class RtfTable
{
    #region 属性
    /// <summary>行列表</summary>
    public List<RtfTableRow> Rows { get; } = [];
    #endregion

    #region 方法
    /// <inheritdoc/>
    public override String ToString() => $"Table({Rows.Count} rows)";
    #endregion
}
