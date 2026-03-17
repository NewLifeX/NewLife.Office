using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>PowerPoint pptx 格式集成测试</summary>
public class PptxTests : IntegrationTestBase
{
    [Fact, DisplayName("PPT_pptx_复杂写入再读取往返")]
    public void Pptx_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.pptx");

        using (var w = new PptxWriter())
        {
            // 幻灯片1: 标题页
            var slide1 = w.AddSlide();
            w.AddTextBox(0, "NewLife.Office 演示文稿", 2, 2, 20, 3, fontSize: 36, bold: true);
            w.AddTextBox(0, "自动化测试生成", 2, 6, 20, 2, fontSize: 18);

            // 幻灯片2: 内容页
            var slide2 = w.AddSlide();
            w.AddTextBox(1, "功能列表", 1, 1, 22, 2, fontSize: 28, bold: true);
            w.AddTextBox(1, "1. Excel 读写\n2. Word 文档\n3. PDF 生成\n4. PPT 演示\n5. 更多格式支持...", 1, 3.5, 22, 8, fontSize: 16);

            // 幻灯片3: 表格
            var slide3 = w.AddSlide();
            w.AddTextBox(2, "数据统计", 1, 1, 22, 2, fontSize: 28, bold: true);
            var table = new PptTable
            {
                Left = 360000 * 2,
                Top = 360000 * 4,
                Width = 360000 * 20,
                Height = 360000 * 6,
            };
            table.Rows.Add(new[] { "格式", "读取", "写入" });
            table.Rows.Add(new[] { "XLSX", "✓", "✓" });
            table.Rows.Add(new[] { "DOCX", "✓", "✓" });
            slide3.Tables.Add(table);

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        // 读取验证
        using var reader = new PptxReader(path);
        Assert.Equal(3, reader.GetSlideCount());

        var allText = reader.ReadAllText();
        Assert.Contains("NewLife.Office", allText);
        Assert.Contains("功能列表", allText);

        var slideText0 = reader.GetSlideText(0);
        var slideText1 = reader.GetSlideText(1);
        var slideText2 = reader.GetSlideText(2);
        Assert.Contains("演示文稿", slideText0);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<PptxReader>(factoryReader);
        (factoryReader as IDisposable)?.Dispose();
    }
}
