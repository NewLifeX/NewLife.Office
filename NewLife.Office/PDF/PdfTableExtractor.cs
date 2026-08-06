using System;
using System.Collections.Generic;
using System.Linq;

namespace NewLife.Office.Pdf;

/// <summary>PDF 结构化表格提取器（P09，对标 iText 7/Aspose.PDF）</summary>
/// <remarks>
/// 基于文本块坐标聚类的表格识别：
/// 1. 按 Y 坐标聚合文本块为行
/// 2. 聚类所有行的 X 起点识别列边界（连续多行共享多列 → 表格区域）
/// 3. 按列边界将行内文本块分配到单元格
/// <para>纯文本坐标启发式，无边框表格也可识别；适合规则列对齐的表格。</para>
/// </remarks>
public static class PdfTableExtractor
{
    /// <summary>提取 PDF 中所有页的结构化表格</summary>
    /// <param name="reader">PDF 读取器</param>
    /// <param name="yTolerance">Y 坐标容差（点），同一行内文本的 Y 最大差，默认 5</param>
    /// <param name="xGap">X 列间距阈值（点），超过视为新列，默认 15</param>
    /// <param name="minRows">最小行数（含表头），低于此不识别为表格，默认 2</param>
    /// <returns>表格列表（含页索引与行列数据）</returns>
    public static List<PdfTableData> Extract(PdfReader reader, Single yTolerance = 5f, Single xGap = 15f, Int32 minRows = 2)
    {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        var result = new List<PdfTableData>();

        // 按页提取文本块（xref 正常时每页 PageIndex 正确；xref 失败时回退合并 PageIndex=0）
        foreach (var (pageIdx, blocks) in GetPageBlocks(reader))
        {
            var table = ExtractFromBlocks(blocks, pageIdx, yTolerance, xGap, minRows);
            if (table != null)
                result.Add(table);
        }
        return result;
    }

    /// <summary>按页获取文本块；所有页为空（xref 不可靠）时回退 ExtractTextWithPositions 合并</summary>
    private static List<(Int32 Page, List<PdfText> Blocks)> GetPageBlocks(PdfReader reader)
    {
        var result = new List<(Int32, List<PdfText>)>();
        var pageCount = reader.GetPageCount();
        for (var pi = 0; pi < pageCount; pi++)
        {
            var blocks = reader.GetPageTextBlocks(pi);
            if (blocks is { Count: > 0 })
                result.Add((pi, blocks));
        }
        // 回退：全部文本块合并到第 0 页
        if (result.Count == 0)
        {
            var all = reader.ExtractTextWithPositions().ToList();
            if (all.Count > 0)
                result.Add((0, all));
        }
        return result;
    }

    /// <summary>从单页文本块提取表格（Y 分组 → X 聚类 → 单元格分配）</summary>
    private static PdfTableData? ExtractFromBlocks(List<PdfText> blocks, Int32 pageIdx,
        Single yTolerance, Single xGap, Int32 minRows)
    {
        if (blocks.Count < 4) return null;

        // 按 Y 分组为行
        var rows = GroupByY(blocks, yTolerance);
        if (rows.Count < minRows) return null;

        // 聚类列边界
        var bounds = DetectColumnBounds(rows, xGap);
        if (bounds.Count < 2) return null;

        // 筛选参与表格的行（行内存在多列分布的行）
        var tableRows = rows.Where(r => CountColumns(r, bounds) >= 2).ToList();
        if (tableRows.Count < minRows) return null;

        var table = new PdfTableData { PageIndex = pageIdx };
        foreach (var row in tableRows.OrderByDescending(r => r[0].Y))
        {
            var tr = new PdfTableRowData();
            var sorted = row.OrderBy(t => t.X).ToList();
            for (var ci = 0; ci < bounds.Count; ci++)
            {
                var start = bounds[ci];
                var end = ci + 1 < bounds.Count ? bounds[ci + 1] : Single.MaxValue;
                var cellTexts = sorted.Where(t => t.X >= start - 1 && t.X < end).Select(t => t.Text).ToList();
                tr.Cells.Add(String.Join(" ", cellTexts).Trim());
            }
            // 去掉尾部空单元格
            while (tr.CellCount > 0 && String.IsNullOrEmpty(tr.Cells[^1]))
                tr.Cells.RemoveAt(tr.Cells.Count - 1);
            if (tr.CellCount > 0)
                table.Rows.Add(tr);
        }
        return table.RowCount >= minRows ? table : null;
    }

    /// <summary>按 Y 坐标分组为行（行内保留全部文本块）</summary>
    private static List<List<PdfText>> GroupByY(List<PdfText> blocks, Single yTolerance)
    {
        var rows = new List<List<PdfText>>();
        foreach (var item in blocks.OrderByDescending(t => t.Y))
        {
            var added = false;
            foreach (var row in rows)
            {
                if (Math.Abs(row[0].Y - item.Y) <= yTolerance)
                {
                    row.Add(item);
                    added = true;
                    break;
                }
            }
            if (!added) rows.Add([item]);
        }
        return rows;
    }

    /// <summary>聚类所有文本块的 X 起点为列边界（间隔超过阈值视为新列）</summary>
    private static List<Single> DetectColumnBounds(List<List<PdfText>> rows, Single xGap)
    {
        var xs = rows.SelectMany(r => r.Select(t => t.X)).OrderBy(x => x).ToList();
        var bounds = new List<Single>();
        foreach (var x in xs)
        {
            if (bounds.Count == 0 || x - bounds[^1] > xGap)
                bounds.Add(x);
            else if (x < bounds[^1])
                bounds[^1] = x;
        }
        return bounds;
    }

    /// <summary>统计行内文本块覆盖的列数（X 落在不同列边界内的数量）</summary>
    private static Int32 CountColumns(List<PdfText> row, List<Single> bounds)
    {
        var cols = new HashSet<Int32>();
        foreach (var t in row)
        {
            for (var ci = 0; ci < bounds.Count; ci++)
            {
                var start = bounds[ci];
                var end = ci + 1 < bounds.Count ? bounds[ci + 1] : Single.MaxValue;
                if (t.X >= start - 1 && t.X < end)
                {
                    cols.Add(ci);
                    break;
                }
            }
        }
        return cols.Count;
    }
}
