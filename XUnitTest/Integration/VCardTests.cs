using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>VCard 格式集成测试</summary>
public class VCardTests : IntegrationTestBase
{
    [Fact, DisplayName("VCard_复杂写入再读取往返")]
    public void VCard_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.vcf");

        var contact1 = new VCardContact
        {
            FullName = "张三",
            Name = new VCardName { Family = "张", Given = "三" },
            Organization = "新生命科技",
            Title = "高级工程师",
            Birthday = new DateTime(1990, 5, 15),
            Note = "这是一个测试联系人",
            Url = "https://example.com/zhangsan",
        };
        contact1.Phones.Add(new VCardPhone { Number = "+86-138-0000-0001", Type = "CELL" });
        contact1.Phones.Add(new VCardPhone { Number = "+86-010-12345678", Type = "WORK" });
        contact1.Emails.Add(new VCardEmail { Address = "zhangsan@example.com", Type = "WORK" });
        contact1.Emails.Add(new VCardEmail { Address = "zhangsan@personal.com", Type = "HOME" });
        contact1.Addresses.Add(new VCardAddress
        {
            Street = "中关村大街1号",
            City = "北京",
            Region = "北京",
            PostalCode = "100080",
            Country = "中国",
            Type = "WORK",
        });

        var contact2 = new VCardContact
        {
            FullName = "李四",
            Name = new VCardName { Family = "李", Given = "四" },
            Organization = "新生命科技",
            Title = "产品经理",
        };
        contact2.Phones.Add(new VCardPhone { Number = "+86-139-0000-0002", Type = "CELL" });
        contact2.Emails.Add(new VCardEmail { Address = "lisi@example.com", Type = "WORK" });

        var writer = new VCardWriter();
        writer.WriteAll(new[] { contact1, contact2 }, path);

        Assert.True(File.Exists(path));

        // 读取验证
        var reader = new VCardReader();
        var contacts = reader.ReadAll(path);
        Assert.Equal(2, contacts.Count);

        Assert.Equal("张三", contacts[0].FullName);
        Assert.Equal("新生命科技", contacts[0].Organization);
        Assert.Equal(2, contacts[0].Phones.Count);
        Assert.Equal(2, contacts[0].Emails.Count);
        Assert.True(contacts[0].Addresses.Count >= 1);
        Assert.Equal("zhangsan@example.com", contacts[0].Emails[0].Address);

        Assert.Equal("李四", contacts[1].FullName);
        Assert.Equal(1, contacts[1].Phones.Count);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<VCardReader>(factoryReader);
    }
}
