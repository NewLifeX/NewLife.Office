namespace NewLife.Office.Ppt;

/// <summary>演示文稿模型深拷贝器</summary>
/// <remarks>
/// 克隆幻灯片时元素实例必须独立：否则对克隆页按锚点改写会同时改掉源页（补丁仍报成功，属静默数据破坏）。
/// 图片/视频的字节数组只读共享，其余集合与字典均重建。
/// </remarks>
public static class PptxCloner
{
    /// <summary>深拷贝幻灯片</summary>
    /// <param name="source">源幻灯片</param>
    /// <returns>元素实例独立的副本</returns>
    public static Slide CloneSlide(Slide source)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));

        var clone = (Slide)source.ShallowClone();

        clone.TextBoxes = [];
        foreach (var item in source.TextBoxes) clone.TextBoxes.Add(CloneTextBox(item));

        clone.Images = [];
        foreach (var item in source.Images) clone.Images.Add((Picture)item.ShallowClone());

        clone.Tables = [];
        foreach (var item in source.Tables) clone.Tables.Add(CloneTable(item));

        clone.Shapes = [];
        foreach (var item in source.Shapes) clone.Shapes.Add((Shape)item.ShallowClone());

        clone.Charts = [];
        foreach (var item in source.Charts) clone.Charts.Add(CloneChart(item));

        clone.Groups = [];
        foreach (var item in source.Groups) clone.Groups.Add(CloneGroup(item));

        clone.Connectors = [];
        foreach (var item in source.Connectors) clone.Connectors.Add((ConnectionShape)item.ShallowClone());

        clone.Videos = [];
        foreach (var item in source.Videos) clone.Videos.Add((Video)item.ShallowClone());

        clone.Animations = [];
        foreach (var item in source.Animations) clone.Animations.Add((Animation)item.ShallowClone());

        clone.Comments = [];
        foreach (var item in source.Comments) clone.Comments.Add((Comment)item.ShallowClone());

        if (source.BackgroundImage != null) clone.BackgroundImage = (Picture)source.BackgroundImage.ShallowClone();
        if (source.Transition != null) clone.Transition = (Transition)source.Transition.ShallowClone();

        return clone;
    }

    private static TextBox CloneTextBox(TextBox source)
    {
        var clone = (TextBox)source.ShallowClone();

        clone.Runs = [];
        foreach (var run in source.Runs) clone.Runs.Add(CloneRun(run));

        clone.Paragraphs = [];
        foreach (var paragraph in source.Paragraphs) clone.Paragraphs.Add(CloneParagraph(paragraph));

        return clone;
    }

    private static Paragraph CloneParagraph(Paragraph source)
    {
        var clone = (Paragraph)source.ShallowClone();

        clone.Runs = [];
        foreach (var run in source.Runs) clone.Runs.Add(CloneRun(run));

        return clone;
    }

    private static Run CloneRun(Run source)
    {
        var clone = (Run)source.ShallowClone();
        if (source.GradFillColors != null) clone.GradFillColors = (String[])source.GradFillColors.Clone();
        return clone;
    }

    private static Table CloneTable(Table source)
    {
        var clone = (Table)source.ShallowClone();

        clone.Rows = [];
        foreach (var row in source.Rows) clone.Rows.Add((String[])row.Clone());

        clone.ColWidths = (Int64[])source.ColWidths.Clone();

        clone.CellStyles = [];
        foreach (var kv in source.CellStyles) clone.CellStyles[kv.Key] = (CellStyle)kv.Value.ShallowClone();

        clone.CellBorders = [];
        foreach (var kv in source.CellBorders) clone.CellBorders[kv.Key] = (CellBorder)kv.Value.ShallowClone();

        clone.MergedCells = new Dictionary<(Int32, Int32), (Int32, Int32)>(source.MergedCells);

        return clone;
    }

    private static Chart CloneChart(Chart source)
    {
        var clone = (Chart)source.ShallowClone();

        clone.Categories = (String[])source.Categories.Clone();

        clone.Series = [];
        foreach (var series in source.Series)
        {
            var item = (ChartSeries)series.ShallowClone();
            item.Values = (Double[])series.Values.Clone();
            if (series.XValues != null) item.XValues = (Double[])series.XValues.Clone();
            clone.Series.Add(item);
        }

        return clone;
    }

    private static GroupShape CloneGroup(GroupShape source)
    {
        var clone = (GroupShape)source.ShallowClone();

        clone.Shapes = [];
        foreach (var item in source.Shapes) clone.Shapes.Add((Shape)item.ShallowClone());

        clone.TextBoxes = [];
        foreach (var item in source.TextBoxes) clone.TextBoxes.Add(CloneTextBox(item));

        clone.Images = [];
        foreach (var item in source.Images) clone.Images.Add((Picture)item.ShallowClone());

        return clone;
    }
}
