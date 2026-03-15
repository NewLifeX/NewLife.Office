using System;
using System.Collections.Generic;

namespace NewLife.Office.Rtf;

/// <summary>RTF 段落</summary>
/// <remarks>
/// 表示文档中的一个逻辑段落，包含若干文本片段（Run）和段落格式信息。
/// </remarks>
public sealed class RtfParagraph
{
    #region 属性
    /// <summary>段落内的文本片段列表</summary>
    public List<RtfRun> Runs { get; } = [];

    /// <summary>段落对齐方式</summary>
    public RtfAlignment Alignment { get; set; } = RtfAlignment.Left;

    /// <summary>左缩进（twips，1440 twips = 1 英寸）</summary>
    public Int32 LeftIndent { get; set; }

    /// <summary>右缩进（twips）</summary>
    public Int32 RightIndent { get; set; }

    /// <summary>首行缩进（twips，负值为悬挂缩进）</summary>
    public Int32 FirstLineIndent { get; set; }

    /// <summary>段前间距（twips）</summary>
    public Int32 SpaceBefore { get; set; }

    /// <summary>段后间距（twips）</summary>
    public Int32 SpaceAfter { get; set; }

    /// <summary>行距（twips，0 = 单倍行距自动）</summary>
    public Int32 LineSpacing { get; set; }

    /// <summary>是否为表格内容段落</summary>
    public Boolean InTable { get; set; }
    #endregion

    #region 方法
    /// <summary>获取段落纯文本内容</summary>
    /// <returns>纯文本字符串</returns>
    public String GetPlainText()
    {
        var sb = new System.Text.StringBuilder();
        foreach (var run in Runs)
        {
            if (run.IsLineBreak)
                sb.AppendLine();
            else
                sb.Append(run.Text);
        }
        return sb.ToString();
    }

    /// <inheritdoc/>
    public override String ToString() => GetPlainText();
    #endregion
}
