namespace NewLife.Office;

/// <summary>PPT 幻灯片自动排版引擎</summary>
/// <remarks>
/// 根据 <see cref="PptSlide.Layout"/> 策略和各元素的 <see cref="PptTextBox.Role"/>，
/// 自动计算 16:9 画布上每个元素的坐标，并直接写入 Left/Top/Width/Height（EMU 单位）。
/// <para>画布尺寸：33.87cm × 19.05cm（16:9）。边距：左/右 2cm，顶部 1.5cm。</para>
/// <example>
/// <code>
/// var slide = new PptSlide { Layout = "title_content" };
/// slide.TextBoxes.Add(new PptTextBox { Text = "标题", Role = "title" });
/// slide.TextBoxes.Add(new PptTextBox { Text = "正文内容", Role = "body" });
/// // 柱状图
/// var chart = new PptChart { ChartType = "bar" };
/// slide.Charts.Add(chart);
///
/// LayoutEngine.Apply(slide);  // 自动计算并写入所有元素坐标
/// </code>
/// </example>
/// </remarks>
public static class LayoutEngine
{
    #region 常量
    /// <summary>幻灯片画布宽度（cm），16:9 标准</summary>
    public const Double SlideW = 33.87;

    /// <summary>幻灯片画布高度（cm）</summary>
    public const Double SlideH = 19.05;

    /// <summary>左/右边距（cm）</summary>
    public const Double MarginL = 2.0;

    /// <summary>顶部边距（cm）</summary>
    public const Double MarginT = 1.5;

    /// <summary>内容区宽度（cm），等于 SlideW - MarginL * 2</summary>
    public const Double ContentW = SlideW - MarginL * 2; // 29.87

    private const Double CmToEmu = 360000.0;
    #endregion

    #region 公共方法
    /// <summary>根据 slide.Layout 策略自动计算所有元素坐标，并写入各元素的 Left/Top/Width/Height</summary>
    /// <param name="slide">目标幻灯片</param>
    public static void Apply(PptSlide slide)
    {
        var strategy = (slide.Layout ?? "title_content").ToLowerInvariant();
        switch (strategy)
        {
            case "title_only":  ApplyTitleOnly(slide);  break;
            case "two_column":  ApplyTwoColumn(slide);  break;
            case "chart_only":  ApplyChartOnly(slide);  break;
            case "blank":       ApplyBlank(slide);      break;
            default:            ApplyTitleContent(slide); break;
        }
    }
    #endregion

    #region 布局策略
    // ── title_content（默认）──────────────────────────────────────────
    private static void ApplyTitleContent(PptSlide slide)
    {
        var curY = MarginT;

        // 标题
        var title = FindRole(slide, "title");
        SetRect(title, MarginL, curY, ContentW, 2.2);
        curY += 2.4;

        // 副标题
        var subtitle = FindRole(slide, "subtitle");
        if (subtitle != null)
        {
            SetRect(subtitle, MarginL, curY, ContentW, 1.2);
            curY += 1.5;
        }

        // KPI 行（最多并排）
        var kpis = slide.TextBoxes.Where(tb => tb.Role == "kpi").ToList();
        if (kpis.Count > 0)
        {
            var kpiW = ContentW / kpis.Count;
            for (var i = 0; i < kpis.Count; i++)
                SetRect(kpis[i], MarginL + i * kpiW, curY, kpiW - 0.3, 3.5);
            curY += 4.0;
        }

        // 正文
        var body = FindRole(slide, "body");
        if (body != null && kpis.Count == 0)
        {
            SetRect(body, MarginL, curY, ContentW, 5.5);
            curY += 6.0;
        }

        // 图表 + 表格共存：图表右半，表格左半
        var chart = slide.Charts.FirstOrDefault();
        var table = slide.Tables.FirstOrDefault();

        if (chart != null && table != null)
        {
            var halfW = (ContentW - 0.8) / 2;
            SetRectTable(table, MarginL, curY, halfW, 8.0);
            SetRectChart(chart, MarginL + halfW + 0.8, curY, halfW, 8.0);
            curY += 8.5;
        }
        else if (chart != null)
        {
            SetRectChart(chart, MarginL, curY, ContentW, 9.0);
            curY += 9.5;
        }
        else if (table != null)
        {
            SetRectTable(table, MarginL, curY, ContentW, 6.5);
            curY += 7.0;
        }

        // 图片
        var image = slide.Images.FirstOrDefault();
        if (image != null)
        {
            SetRectImage(image, MarginL, curY, ContentW, 7.0);
            curY += 7.5;
        }

        // 注释（固定在底部）
        var caption = FindRole(slide, "caption");
        if (caption != null)
            SetRect(caption, MarginL, SlideH - 1.5, ContentW, 1.0);

        // 兜底：其余未分配的正文文本框（非 kpi/caption/title/subtitle/body）
        foreach (var tb in slide.TextBoxes)
        {
            if (ReferenceEquals(tb, title) || ReferenceEquals(tb, subtitle) ||
                ReferenceEquals(tb, body)  || ReferenceEquals(tb, caption)  ||
                kpis.Any(k => ReferenceEquals(k, tb))) continue;
            SetRect(tb, MarginL, curY, ContentW, 2.0);
            curY += 2.5;
        }
    }

    // ── title_only（封面）────────────────────────────────────────────
    private static void ApplyTitleOnly(PptSlide slide)
    {
        var title = FindRole(slide, "title") ?? slide.TextBoxes.FirstOrDefault();
        SetRect(title, MarginL, SlideH / 2 - 2.0, ContentW, 4.0);

        var subtitle = FindRole(slide, "subtitle");
        if (subtitle != null)
            SetRect(subtitle, MarginL, SlideH / 2 + 2.2, ContentW, 1.5);
    }

    // ── two_column（双栏）────────────────────────────────────────────
    private static void ApplyTwoColumn(PptSlide slide)
    {
        var colW = (ContentW - 1.0) / 2;
        var curY = MarginT;

        var title = FindRole(slide, "title");
        SetRect(title, MarginL, curY, ContentW, 2.0);
        curY += 2.5;

        var titleObj = (Object?)title;
        var nonTitles = slide.TextBoxes.Where(tb => !ReferenceEquals(tb, titleObj)).Cast<Object>()
            .Concat(slide.Tables.Cast<Object>())
            .Concat(slide.Charts.Cast<Object>())
            .Concat(slide.Images.Cast<Object>())
            .ToList();

        var half  = nonTitles.Count / 2;
        var left  = nonTitles.Take(half + nonTitles.Count % 2).ToList();
        var right = nonTitles.Skip(half + nonTitles.Count % 2).ToList();

        var lyL = curY;
        foreach (var elem in left)
        {
            var h = GetHeight(elem);
            SetRectObj(elem, MarginL, lyL, colW, h);
            lyL += h + 0.5;
        }
        var lyR = curY;
        foreach (var elem in right)
        {
            var h = GetHeight(elem);
            SetRectObj(elem, MarginL + colW + 1, lyR, colW, h);
            lyR += h + 0.5;
        }
    }

    // ── chart_only（全幅图表）────────────────────────────────────────
    private static void ApplyChartOnly(PptSlide slide)
    {
        var title = FindRole(slide, "title");
        if (title != null)
            SetRect(title, MarginL, MarginT, ContentW, 1.5);

        var chart = slide.Charts.FirstOrDefault();
        if (chart != null)
        {
            var topY = title != null ? MarginT + 2.0 : MarginT;
            SetRectChart(chart, MarginL, topY, ContentW, SlideH - topY - 1.5);
        }
    }

    // ── blank（流式布局）─────────────────────────────────────────────
    private static void ApplyBlank(PptSlide slide)
    {
        var curY = MarginT;
        foreach (Object elem in slide.TextBoxes.Cast<Object>()
            .Concat(slide.Tables.Cast<Object>())
            .Concat(slide.Charts.Cast<Object>())
            .Concat(slide.Images.Cast<Object>()))
        {
            var h = GetHeight(elem);
            SetRectObj(elem, MarginL, curY, ContentW, h);
            curY += h + 0.5;
        }
    }
    #endregion

    #region 辅助方法
    private static PptTextBox? FindRole(PptSlide slide, String role) =>
        slide.TextBoxes.FirstOrDefault(tb => tb.Role == role);

    /// <summary>设置 PptTextBox 坐标（cm → EMU）</summary>
    private static void SetRect(PptTextBox? tb, Double x, Double y, Double w, Double h)
    {
        if (tb == null) return;
        tb.Left   = (Int64)(x * CmToEmu);
        tb.Top    = (Int64)(y * CmToEmu);
        tb.Width  = (Int64)(w * CmToEmu);
        tb.Height = (Int64)(h * CmToEmu);
    }

    private static void SetRectTable(PptTable table, Double x, Double y, Double w, Double h)
    {
        table.Left   = (Int64)(x * CmToEmu);
        table.Top    = (Int64)(y * CmToEmu);
        table.Width  = (Int64)(w * CmToEmu);
        table.Height = (Int64)(h * CmToEmu);
    }

    private static void SetRectChart(PptChart chart, Double x, Double y, Double w, Double h)
    {
        chart.Left   = (Int64)(x * CmToEmu);
        chart.Top    = (Int64)(y * CmToEmu);
        chart.Width  = (Int64)(w * CmToEmu);
        chart.Height = (Int64)(h * CmToEmu);
    }

    private static void SetRectImage(PptImage image, Double x, Double y, Double w, Double h)
    {
        image.Left   = (Int64)(x * CmToEmu);
        image.Top    = (Int64)(y * CmToEmu);
        image.Width  = (Int64)(w * CmToEmu);
        image.Height = (Int64)(h * CmToEmu);
    }

    private static void SetRectObj(Object elem, Double x, Double y, Double w, Double h)
    {
        switch (elem)
        {
            case PptTextBox tb: SetRect(tb, x, y, w, h);         break;
            case PptTable t:    SetRectTable(t, x, y, w, h);     break;
            case PptChart c:    SetRectChart(c, x, y, w, h);     break;
            case PptImage img:  SetRectImage(img, x, y, w, h);   break;
        }
    }

    private static Double GetHeight(Object elem) => elem switch
    {
        PptTable   => 6.0,
        PptChart   => 9.0,
        PptImage   => 7.0,
        PptTextBox tb => tb.Role == "kpi" ? 3.5 : 2.5,
        _          => 2.5,
    };
    #endregion
}
