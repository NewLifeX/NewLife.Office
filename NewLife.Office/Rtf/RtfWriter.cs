using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office.Rtf;

/// <summary>RTF 文档写入器</summary>
/// <remarks>
/// 用于创建 RTF 格式文档，支持：段落格式、字符格式（粗体/斜体/下划线/颜色）、表格、文档属性和模板填充。
/// 写入后通过 ToString() 或 Save() 输出。
/// </remarks>
public sealed class RtfWriter
{
    #region 属性
    /// <summary>文档标题</summary>
    public String Title { get; set; } = "";

    /// <summary>文档作者</summary>
    public String Author { get; set; } = "";

    /// <summary>文档主题</summary>
    public String Subject { get; set; } = "";

    /// <summary>默认字体名称</summary>
    public String DefaultFont { get; set; } = "Times New Roman";

    /// <summary>默认字体大小（磅，内部以半磅存储）</summary>
    public Double DefaultFontSize { get; set; } = 12.0;

    // 内容块列表（RtfParagraph 或 RtfTable）
    private readonly List<Object> _blocks = [];
    // 字体表：字体名 → 索引
    private readonly Dictionary<String, Int32> _fontMap = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<String> _fonts = [];
    // 颜色表：RGB → 索引（索引0保留为auto，实际用1起）
    private readonly Dictionary<Int32, Int32> _colorMap = [];
    private readonly List<Int32> _colors = [];
    #endregion

    #region 方法 — 添加内容
    /// <summary>添加纯文本段落</summary>
    /// <param name="text">段落文本</param>
    /// <returns>当前写入器（链式调用）</returns>
    public RtfWriter AddParagraph(String text)
    {
        var para = new RtfParagraph();
        if (!String.IsNullOrEmpty(text))
            para.Runs.Add(new RtfRun { Text = text });
        _blocks.Add(para);
        return this;
    }

    /// <summary>添加带格式的段落对象</summary>
    /// <param name="para">段落对象</param>
    /// <returns>当前写入器（链式调用）</returns>
    public RtfWriter AddParagraph(RtfParagraph para)
    {
        _blocks.Add(para);
        return this;
    }

    /// <summary>添加表格（字符串二维数组）</summary>
    /// <param name="rows">行数据，每行为列文本数组</param>
    /// <returns>当前写入器（链式调用）</returns>
    public RtfWriter AddTable(IEnumerable<IEnumerable<String>> rows)
    {
        var table = new RtfTable();
        foreach (var row in rows)
        {
            var tableRow = new RtfTableRow();
            foreach (var cellText in row)
            {
                var cell = new RtfTableCell();
                var para = new RtfParagraph { InTable = true };
                if (!String.IsNullOrEmpty(cellText))
                    para.Runs.Add(new RtfRun { Text = cellText });
                cell.Paragraphs.Add(para);
                tableRow.Cells.Add(cell);
            }
            if (tableRow.Cells.Count > 0)
                table.Rows.Add(tableRow);
        }
        if (table.Rows.Count > 0)
            _blocks.Add(table);
        return this;
    }

    /// <summary>添加 RtfTable 对象</summary>
    /// <param name="table">表格对象</param>
    /// <returns>当前写入器（链式调用）</returns>
    public RtfWriter AddTable(RtfTable table)
    {
        _blocks.Add(table);
        return this;
    }
    #endregion

    #region 方法 — 输出
    /// <summary>生成 RTF 字符串</summary>
    /// <returns>RTF 格式字符串</returns>
    public override String ToString()
    {
        // Two-pass: pass1 collect fonts/colors, pass2 emit
        CollectFontsAndColors();
        return Emit();
    }

    /// <summary>保存到文件</summary>
    /// <param name="path">文件路径</param>
    public void Save(String path)
    {
        var rtf = ToString();
        File.WriteAllText(path, rtf, new UTF8Encoding(false));
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">目标流</param>
    public void Save(Stream stream)
    {
        var rtf = ToString();
        var bytes = new UTF8Encoding(false).GetBytes(rtf);
        stream.Write(bytes, 0, bytes.Length);
    }
    #endregion

    #region 静态 — 模板填充
    /// <summary>将 RTF 模板中的 {{Key}} 占位符替换为实际值</summary>
    /// <param name="rtf">RTF 模板文本</param>
    /// <param name="values">键值对字典</param>
    /// <returns>替换后的 RTF 文本</returns>
    public static String FillTemplate(String rtf, IDictionary<String, String> values)
    {
        if (String.IsNullOrEmpty(rtf) || values == null || values.Count == 0) return rtf;
        return Regex.Replace(rtf, @"\{\{(\w+)\}\}", m =>
        {
            var key = m.Groups[1].Value;
            return values.TryGetValue(key, out var v) ? EscapeRtf(v) : m.Value;
        });
    }
    #endregion

    #region 核心：字体/颜色收集
    private void CollectFontsAndColors()
    {
        _fontMap.Clear(); _fonts.Clear();
        _colorMap.Clear(); _colors.Clear();

        EnsureFont(DefaultFont);

        foreach (var block in _blocks)
        {
            if (block is RtfParagraph para)
                CollectFromPara(para);
            else if (block is RtfTable table)
            {
                foreach (var row in table.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var cp in cell.Paragraphs)
                            CollectFromPara(cp);
            }
        }
    }

    private void CollectFromPara(RtfParagraph para)
    {
        foreach (var run in para.Runs)
        {
            if (run.FontName != null) EnsureFont(run.FontName);
            if (run.ForeColor >= 0) EnsureColor(run.ForeColor);
            if (run.BackColor >= 0) EnsureColor(run.BackColor);
        }
    }

    private Int32 EnsureFont(String name)
    {
        if (_fontMap.TryGetValue(name, out var idx)) return idx;
        idx = _fonts.Count;
        _fonts.Add(name);
        _fontMap[name] = idx;
        return idx;
    }

    private Int32 EnsureColor(Int32 rgb)
    {
        if (_colorMap.TryGetValue(rgb, out var idx)) return idx;
        idx = _colors.Count + 1; // 1-based (0=auto)
        _colors.Add(rgb);
        _colorMap[rgb] = idx;
        return idx;
    }
    #endregion

    #region 核心：RTF 输出
    private String Emit()
    {
        var sb = new StringBuilder(4096);
        // RTF header
        sb.Append("{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat");
        // Font table
        sb.Append("\r\n{\\fonttbl");
        for (var i = 0; i < _fonts.Count; i++)
        {
            var fname = _fonts[i];
            var family = GuessFamily(fname);
            sb.Append($"{{\\f{i}\\f{family}\\fcharset0 {EscapeRtf(fname)};}}");
        }
        sb.Append('}');
        // Color table
        if (_colors.Count > 0)
        {
            sb.Append("\r\n{\\colortbl ;");
            foreach (var rgb in _colors)
                sb.Append($"\\red{(rgb >> 16) & 0xFF}\\green{(rgb >> 8) & 0xFF}\\blue{rgb & 0xFF};");
            sb.Append('}');
        }
        // Document info
        if (!String.IsNullOrEmpty(Title) || !String.IsNullOrEmpty(Author) || !String.IsNullOrEmpty(Subject))
        {
            sb.Append("\r\n{\\info");
            if (!String.IsNullOrEmpty(Title)) sb.Append($"{{\\title {EscapeRtf(Title)}}}");
            if (!String.IsNullOrEmpty(Author)) sb.Append($"{{\\author {EscapeRtf(Author)}}}");
            if (!String.IsNullOrEmpty(Subject)) sb.Append($"{{\\subject {EscapeRtf(Subject)}}}");
            sb.Append('}');
        }
        sb.Append("\r\n\\viewkind4\\uc1\r\n");

        // Content
        foreach (var block in _blocks)
        {
            if (block is RtfParagraph para)
                EmitParagraph(sb, para);
            else if (block is RtfTable table)
                EmitTable(sb, table);
        }

        sb.Append('}');
        return sb.ToString();
    }

    private void EmitParagraph(StringBuilder sb, RtfParagraph para)
    {
        sb.Append("\\pard");
        // Alignment
        switch (para.Alignment)
        {
            case RtfAlignment.Center: sb.Append("\\qc"); break;
            case RtfAlignment.Right: sb.Append("\\qr"); break;
            case RtfAlignment.Justify: sb.Append("\\qj"); break;
        }
        // Indentation / spacing
        if (para.LeftIndent != 0) sb.Append($"\\li{para.LeftIndent}");
        if (para.RightIndent != 0) sb.Append($"\\ri{para.RightIndent}");
        if (para.FirstLineIndent != 0) sb.Append($"\\fi{para.FirstLineIndent}");
        if (para.SpaceBefore != 0) sb.Append($"\\sb{para.SpaceBefore}");
        if (para.SpaceAfter != 0) sb.Append($"\\sa{para.SpaceAfter}");
        if (para.LineSpacing != 0) sb.Append($"\\sl{para.LineSpacing}\\slmult1");
        if (para.InTable) sb.Append("\\intbl");
        sb.Append(' ');
        foreach (var run in para.Runs)
            EmitRun(sb, run);
        sb.Append("\\par\r\n");
    }

    private void EmitRun(StringBuilder sb, RtfRun run)
    {
        if (run.IsLineBreak) { sb.Append("\\line "); return; }
        var hasFmt = run.Bold || run.Italic || run.Underline || run.Strikethrough
                     || run.ForeColor >= 0 || run.BackColor >= 0
                     || run.FontName != null || run.FontSize > 0;
        if (hasFmt) sb.Append('{');

        // Font
        if (run.FontName != null && _fontMap.TryGetValue(run.FontName, out var fi))
            sb.Append($"\\f{fi}");
        // Font size
        if (run.FontSize > 0) sb.Append($"\\fs{run.FontSize}");
        // Format
        if (run.Bold) sb.Append("\\b");
        if (run.Italic) sb.Append("\\i");
        if (run.Underline) sb.Append("\\ul");
        if (run.Strikethrough) sb.Append("\\strike");
        // Color
        if (run.ForeColor >= 0 && _colorMap.TryGetValue(run.ForeColor, out var fcIdx))
            sb.Append($"\\cf{fcIdx}");
        if (run.BackColor >= 0 && _colorMap.TryGetValue(run.BackColor, out var bcIdx))
            sb.Append($"\\highlight{bcIdx}");
        if (hasFmt) sb.Append(' ');
        sb.Append(EscapeRtf(run.Text ?? ""));
        if (hasFmt) sb.Append('}');
    }

    private void EmitTable(StringBuilder sb, RtfTable table)
    {
        // Calculate default cell widths (equal width, 10080 twips = 7 inches)
        const Int32 totalWidth = 10080;
        foreach (var row in table.Rows)
        {
            var cols = row.Cells.Count;
            if (cols == 0) continue;
            sb.Append("\\trowd\\trgaph108\\trrh0");
            for (var i = 0; i < cols; i++)
            {
                var rightBound = row.Cells[i].RightBoundary > 0
                    ? row.Cells[i].RightBoundary
                    : totalWidth * (i + 1) / cols;
                sb.Append($"\\cellx{rightBound}");
            }
            sb.Append("\r\n");
            foreach (var cell in row.Cells)
                EmitCell(sb, cell);
            sb.Append("\\row\r\n");
        }
    }

    private void EmitCell(StringBuilder sb, RtfTableCell cell)
    {
        foreach (var para in cell.Paragraphs)
        {
            sb.Append("\\pard\\intbl ");
            foreach (var run in para.Runs)
                EmitRun(sb, run);
        }
        sb.Append("\\cell\r\n");
    }
    #endregion

    #region 辅助
    private static String EscapeRtf(String text)
    {
        if (String.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length + 8);
        foreach (var ch in text)
        {
            if (ch == '\\') sb.Append("\\\\");
            else if (ch == '{') sb.Append("\\{");
            else if (ch == '}') sb.Append("\\}");
            else if (ch > 127)
            {
                // Unicode escape: \uN?
                sb.Append($"\\u{(Int32)ch}?");
            }
            else sb.Append(ch);
        }
        return sb.ToString();
    }

    private static String GuessFamily(String fontName)
    {
        if (fontName == null) return "roman";
        var lc = fontName.ToLowerInvariant();
        if (lc.Contains("courier") || lc.Contains("mono") || lc.Contains("console") || lc.Contains("code"))
            return "modern";
        if (lc.Contains("arial") || lc.Contains("helvetica") || lc.Contains("calibri") || lc.Contains("gothic"))
            return "swiss";
        return "roman";
    }
    #endregion
}
