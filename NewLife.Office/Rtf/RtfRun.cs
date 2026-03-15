using System;
using System.Collections.Generic;

namespace NewLife.Office.Rtf;

/// <summary>RTF 文本片段（Run）</summary>
/// <remarks>
/// 表示具有统一格式的一段连续文本，是 RTF 文档的最小排版单元。
/// </remarks>
public sealed class RtfRun
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字体名称（null 表示使用继承值）</summary>
    public String? FontName { get; set; }

    /// <summary>字体大小（半磅，0 表示继承）</summary>
    public Int32 FontSize { get; set; }

    /// <summary>是否粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>是否斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>是否下划线</summary>
    public Boolean Underline { get; set; }

    /// <summary>是否删除线</summary>
    public Boolean Strikethrough { get; set; }

    /// <summary>前景色（RGB，-1 表示默认）</summary>
    public Int32 ForeColor { get; set; } = -1;

    /// <summary>背景色（RGB，-1 表示默认）</summary>
    public Int32 BackColor { get; set; } = -1;

    /// <summary>是否为硬换行（不新起段落）</summary>
    public Boolean IsLineBreak { get; set; }
    #endregion

    #region 方法
    /// <inheritdoc/>
    public override String ToString() => $"{Text.Substring(0, Math.Min(Text.Length, 40))}{(Text.Length > 40 ? "..." : "")}";
    #endregion
}
