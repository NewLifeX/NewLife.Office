using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>幻灯片自动排版引擎测试</summary>
public class LayoutEngineTests
{
    private const Double CmToEmu = 360000.0;

    [Fact(DisplayName = "LayoutEngine—KPI行之后的正文仍被排版（避免保持临时坐标）")]
    public void TitleContent_BodyAfterKpis_Positioned()
    {
        var slide = new Slide { Layout = "title_content" };
        slide.TextBoxes.Add(new TextBox { Text = "标题", Role = "title" });
        slide.TextBoxes.Add(new TextBox { Text = "128,400", Role = "kpi" });
        slide.TextBoxes.Add(new TextBox { Text = "说明文字", Role = "body" });

        LayoutEngine.Apply(slide);

        var body = slide.TextBoxes.First(tb => tb.Role == "body");
        // 未被排版的特征：宽度保持新增时的初始值/临时坐标；这里断言宽度已展开、左缩进为页面边距、位于 KPI 行之下
        Assert.Equal((Int64)(LayoutEngine.MarginL * CmToEmu), body.Left);
        Assert.Equal((Int64)(LayoutEngine.ContentW * CmToEmu), body.Width);
        Assert.True(body.Top > (Int64)(LayoutEngine.MarginT * CmToEmu));
    }

    [Fact(DisplayName = "LayoutEngine—无KPI时正文与标题均被排版")]
    public void TitleContent_BodyWithoutKpis_Positioned()
    {
        var slide = new Slide { Layout = "title_content" };
        slide.TextBoxes.Add(new TextBox { Text = "标题", Role = "title" });
        slide.TextBoxes.Add(new TextBox { Text = "正文", Role = "body" });

        LayoutEngine.Apply(slide);

        var title = slide.TextBoxes.First(tb => tb.Role == "title");
        var body = slide.TextBoxes.First(tb => tb.Role == "body");
        Assert.Equal((Int64)(LayoutEngine.MarginT * CmToEmu), title.Top);
        Assert.True(body.Top > title.Top);
    }
}
