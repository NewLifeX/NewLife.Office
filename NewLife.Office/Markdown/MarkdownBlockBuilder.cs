using System.Collections.Generic;

namespace NewLife.Office.Markdown;

/// <summary>Markdown 块构建辅助（供各格式前端复用）</summary>
/// <remarks>
/// 对标 anydoc shared 层：各格式前端统一经此构建块模型，
/// 保证文本清洗与表格构造行为一致，避免每格式各自拼字符串。
/// </remarks>
public static class MarkdownBlockBuilder
{
    /// <summary>创建段落块（自动清洗文本）</summary>
    /// <param name="text">文本，可为 null</param>
    /// <returns>段落块</returns>
    public static MarkdownBlock Paragraph(String? text)
    {
        var clean = MarkdownTextCleaner.Clean(text);
        return MarkdownBlock.CreateParagraph([MarkdownInline.CreateText(clean)]);
    }

    /// <summary>创建标题块（自动清洗文本）</summary>
    /// <param name="level">等级（1-6）</param>
    /// <param name="text">文本，可为 null</param>
    /// <returns>标题块</returns>
    public static MarkdownBlock Heading(Int32 level, String? text)
    {
        var clean = MarkdownTextCleaner.Clean(text);
        return MarkdownBlock.CreateHeading(level, [MarkdownInline.CreateText(clean)]);
    }

    /// <summary>创建 GFM 表格块（二维数组，首行可选为表头）</summary>
    /// <param name="rows">行数据（每行为一维字符串数组）</param>
    /// <param name="firstRowHeader">首行是否表头，默认 true</param>
    /// <returns>表格块，空输入返回空表格</returns>
    public static MarkdownBlock Table(String[][] rows, Boolean firstRowHeader = true)
    {
        var table = new TableBlock();
        if (rows == null || rows.Length == 0) return table;

        for (var r = 0; r < rows.Length; r++)
        {
            var row = rows[r] ?? [];
            var rowBlock = new TableRowBlock();
            for (var c = 0; c < row.Length; c++)
            {
                var isHeader = firstRowHeader && r == 0;
                // GFM 表格单元格文本（管道符由写入器在序列化时统一转义，避免双重转义）
                var cellText = MarkdownTextCleaner.CleanCell(row[c]);
                var cell = new TableCellBlock([MarkdownInline.CreateText(cellText)], isHeader);
                rowBlock.Children.Add(cell);
            }
            table.Children.Add(rowBlock);
        }
        return table;
    }

    /// <summary>创建引用块（默认用于备注、注释等）</summary>
    /// <param name="text">文本，可为 null</param>
    /// <returns>引用块</returns>
    public static MarkdownBlock BlockQuote(String? text)
    {
        return MarkdownBlock.CreateBlockQuote([Paragraph(text)]);
    }
}
