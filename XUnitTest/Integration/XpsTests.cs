using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>XPS 格式集成测试</summary>
public class XpsTests : IntegrationTestBase
{
    [Fact, DisplayName("XPS_复杂写入再读取往返")]
    public void Xps_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.xps");

        var writer = new XpsWriter();
        writer.SetProperties(new XpsProperties
        {
            Title = "XPS集成测试",
            Creator = "NewLife Office",
            Subject = "自动化测试",
        });

        writer.AddPage(816, 1056, new[]
        {
            ("XPS 集成测试文档", 96.0, 96.0, 24.0),
            ("第一页内容。", 96.0, 160.0, 14.0),
            ("这是一个多行文本的测试。", 96.0, 200.0, 12.0),
        });

        writer.AddPage(816, 1056, new[]
        {
            ("第二页", 96.0, 96.0, 20.0),
            ("第二页的内容。", 96.0, 160.0, 14.0),
        });

        writer.Save(path);

        Assert.True(File.Exists(path));

        // 读取验证
        var reader = new XpsReader();
        var pages = reader.Read(path);
        Assert.Equal(2, pages.Count);
        Assert.Contains("XPS", pages[0].Text);
        Assert.Contains("第二页", pages[1].Text);
        Assert.True(pages[0].Width > 0);
        Assert.True(pages[0].Height > 0);

        // 属性
        var props = reader.ReadProperties(path);
        Assert.Equal("XPS集成测试", props?.Title);
        Assert.Equal("NewLife Office", props?.Creator);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<XpsReader>(factoryReader);
    }
}
