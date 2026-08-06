using System.ComponentModel;
using System.IO;
using NewLife.Office.Pdf;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>PDF Fluent 声明式布局测试（P07，审计缺口）</summary>
/// <remarks>
/// 覆盖链式组件布局（P07-01）、自动分页（P07-02）、组件复用（P07-03）。
/// </remarks>
public class FluentTests
{
    static FluentTests() => System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

    [Fact, DisplayName("P07-01 Fluent：链式组件布局生成有效 PDF")]
    public void Fluent_ComponentLayout()
    {
        using var ms = new MemoryStream();
        using (var doc = new PdfDocumentBuilder())
        {
            doc.AddText("Fluent 布局标题", 18)
               .AddEmptyLine()
               .AddText("第一段正文内容，验证链式调用。")
               .AddTable(new[] { new[] { "列A", "列B" }, new[] { "1", "2" } }, firstRowHeader: true)
               .Save(ms);
        }
        ms.Position = 0;

        using var reader = new PdfReader(ms);
        Assert.True(reader.GetPageCount() >= 1);
        var text = reader.ExtractText();
        Assert.Contains("Fluent 布局标题", text);
        Assert.Contains("第一段正文内容", text);
    }

    [Fact, DisplayName("P07-02 Fluent：内容超高自动分页")]
    public void Fluent_AutoPageBreak()
    {
        using var ms = new MemoryStream();
        using (var doc = new PdfDocumentBuilder())
        {
            // 每行 14pt，A4 可用高度约 700pt → 60 行必然跨页
            for (var i = 0; i < 60; i++)
                doc.AddText($"自动分页测试行 {i}", 12);
            doc.Save(ms);
        }
        ms.Position = 0;

        using var reader = new PdfReader(ms);
        Assert.True(reader.GetPageCount() > 1, $"页数 {reader.GetPageCount()} 应大于 1");
        var text = reader.ExtractText();
        Assert.Contains("自动分页测试行 59", text);
    }

    [Fact, DisplayName("P07-03 Fluent：组件复用（UseComponent）")]
    public void Fluent_ReusableComponent()
    {
        using var ms = new MemoryStream();
        using (var doc = new PdfDocumentBuilder())
        {
            // 定义复用组件：页头 + 分隔线
            void Header(PdfDocumentBuilder d)
            {
                d.AddText("复用页眉", 14);
                d.AddRule(0.5f, "333333");
                d.AddEmptyLine(6);
            }

            Header(doc);
            doc.AddText("第一页内容。");
            doc.PageBreak();
            Header(doc);
            doc.AddText("第二页内容。");
            doc.Save(ms);
        }
        ms.Position = 0;

        using var reader = new PdfReader(ms);
        Assert.True(reader.GetPageCount() >= 2, $"页数 {reader.GetPageCount()} 应 >= 2");
        var text = reader.ExtractText();
        Assert.Contains("复用页眉", text);
        Assert.Contains("第一页内容", text);
        Assert.Contains("第二页内容", text);
    }
}
