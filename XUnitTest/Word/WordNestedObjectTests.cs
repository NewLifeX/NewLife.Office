using NewLife.Office;
using Xunit;

namespace XUnitTest.Word;

/// <summary>WordWriter.WriteObjects 嵌套对象映射单元测试 — W10-05</summary>
public class WordNestedObjectTests
{
    #region 测试模型
    public class Customer
    {
        [System.ComponentModel.DisplayName("客户名称")]
        public String Name { get; set; } = "";

        public Address? BillingAddress { get; set; }
    }

    public class Address
    {
        [System.ComponentModel.DisplayName("城市")]
        public String City { get; set; } = "";

        public String Street { get; set; } = "";
    }
    #endregion

    #region 嵌套对象
    [Fact(DisplayName = "嵌套对象—深度0仅扁平属性")]
    public void WriteObjects_Depth0_FlatOnly()
    {
        var data = new[]
        {
            new Customer { Name = "张三", BillingAddress = new Address { City = "北京", Street = "长安街1号" } },
        };

        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.WriteObjects(data, firstRowHeader: true, maxDepth: 0);
        writer.Save(ms);

        ms.Position = 0;
        var text = OfficeFactory.ReadText(ms, ".docx");
        Assert.NotNull(text);
        Assert.Contains("张三", text);
        // 深度0不展开嵌套对象
        Assert.DoesNotContain("北京", text);
    }

    [Fact(DisplayName = "嵌套对象—深度1展开嵌套属性")]
    public void WriteObjects_Depth1_Expand()
    {
        var data = new[]
        {
            new Customer { Name = "张三", BillingAddress = new Address { City = "北京", Street = "长安街1号" } },
        };

        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.WriteObjects(data, firstRowHeader: true, maxDepth: 1);
        writer.Save(ms);

        ms.Position = 0;
        var text = OfficeFactory.ReadText(ms, ".docx");
        Assert.NotNull(text);
        Assert.Contains("张三", text);
        Assert.Contains("北京", text);
        Assert.Contains("长安街1号", text);
    }

    [Fact(DisplayName = "嵌套对象—DisplayName列名优先")]
    public void WriteObjects_Nested_DisplayName()
    {
        var data = new[]
        {
            new Customer { Name = "测试", BillingAddress = new Address { City = "上海", Street = "南京路" } },
        };

        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.WriteObjects(data, firstRowHeader: true, maxDepth: 1);
        writer.Save(ms);

        ms.Position = 0;
        var text = OfficeFactory.ReadText(ms, ".docx");
        Assert.NotNull(text);
        // 列名应为 Parent.LeafDisplayName 格式
        Assert.Contains("客户名称", text);
        Assert.Contains("BillingAddress.城市", text);
    }
    #endregion
}
