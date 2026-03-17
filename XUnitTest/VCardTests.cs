using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>vCard 联系人格式读写测试</summary>
public class VCardTests
{
    #region 读取测试

    [Fact]
    [DisplayName("解析基本联系人（FN/TEL/EMAIL）")]
    public void Parse_BasicContact_ReadsAllFields()
    {
        var vcf = "BEGIN:VCARD\r\nVERSION:4.0\r\n"
                + "FN:John Doe\r\n"
                + "N:Doe;John;;;\r\n"
                + "TEL;TYPE=WORK:+1-555-100-1234\r\n"
                + "EMAIL;TYPE=WORK:john.doe@example.com\r\n"
                + "END:VCARD\r\n";

        var reader = new VCardReader();
        var contacts = reader.ParseAll(vcf);

        Assert.Single(contacts);
        var c = contacts[0];
        Assert.Equal("John Doe", c.FullName);
        Assert.Equal("Doe", c.Name?.Family);
        Assert.Equal("John", c.Name?.Given);
        Assert.Single(c.Phones);
        Assert.Equal("+1-555-100-1234", c.Phones[0].Number);
        Assert.Equal("WORK", c.Phones[0].Type);
        Assert.Single(c.Emails);
        Assert.Equal("john.doe@example.com", c.Emails[0].Address);
    }

    [Fact]
    [DisplayName("解析多个联系人（多 VCARD 块）")]
    public void Parse_MultipleContacts_AllParsed()
    {
        var vcf = "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Alice\r\nEND:VCARD\r\n"
                + "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Bob\r\nEND:VCARD\r\n"
                + "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Charlie\r\nEND:VCARD\r\n";

        var reader = new VCardReader();
        var contacts = reader.ParseAll(vcf);

        Assert.Equal(3, contacts.Count);
        Assert.Equal("Alice", contacts[0].FullName);
        Assert.Equal("Charlie", contacts[2].FullName);
    }

    [Fact]
    [DisplayName("解析地址（ADR）字段")]
    public void Parse_Address_Parsed()
    {
        var vcf = "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Test\r\n"
                + "ADR;TYPE=HOME:;;123 Main St;Springfield;IL;62701;USA\r\n"
                + "END:VCARD\r\n";

        var reader = new VCardReader();
        var contacts = reader.ParseAll(vcf);

        Assert.Single(contacts[0].Addresses);
        var adr = contacts[0].Addresses[0];
        Assert.Equal("123 Main St", adr.Street);
        Assert.Equal("Springfield", adr.City);
        Assert.Equal("IL", adr.Region);
        Assert.Equal("USA", adr.Country);
        Assert.Equal("HOME", adr.Type);
    }

    [Fact]
    [DisplayName("解析生日（BDAY）字段")]
    public void Parse_Birthday_Parsed()
    {
        var vcf = "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Test\r\nBDAY:19900515\r\nEND:VCARD\r\n";

        var reader = new VCardReader();
        var contacts = reader.ParseAll(vcf);

        Assert.NotNull(contacts[0].Birthday);
        Assert.Equal(1990, contacts[0].Birthday!.Value.Year);
        Assert.Equal(5, contacts[0].Birthday!.Value.Month);
        Assert.Equal(15, contacts[0].Birthday!.Value.Day);
    }

    [Fact]
    [DisplayName("解析 ORG 和 TITLE")]
    public void Parse_OrgAndTitle_Parsed()
    {
        var vcf = "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Jane Smith\r\n"
                + "ORG:Acme Corporation\r\nTITLE:Senior Engineer\r\nEND:VCARD\r\n";

        var reader = new VCardReader();
        var contacts = reader.ParseAll(vcf);

        Assert.Equal("Acme Corporation", contacts[0].Organization);
        Assert.Equal("Senior Engineer", contacts[0].Title);
    }

    [Fact]
    [DisplayName("解析折行属性（续行合并）")]
    public void Parse_FoldedLines_Unfolded()
    {
        var vcf = "BEGIN:VCARD\r\nVERSION:4.0\r\n"
                + "FN:Very Long Na\r\n me That Continues\r\n"
                + "END:VCARD\r\n";

        var reader = new VCardReader();
        var contacts = reader.ParseAll(vcf);

        Assert.Equal("Very Long Name That Continues", contacts[0].FullName);
    }

    #endregion

    #region 写入测试

    [Fact]
    [DisplayName("写入基本联系人")]
    public void Write_BasicContact_ContainsRequiredFields()
    {
        var contact = new VCardContact
        {
            FullName = "Jane Doe",
            Organization = "NewLife",
            Title = "Developer",
        };
        contact.Phones.Add(new VCardPhone { Number = "+86-10-12345678", Type = "WORK" });
        contact.Emails.Add(new VCardEmail { Address = "jane@newlife.org", Type = "WORK" });

        var writer = new VCardWriter();
        var vcf = writer.Build(contact);

        Assert.Contains("BEGIN:VCARD", vcf);
        Assert.Contains("FN:Jane Doe", vcf);
        Assert.Contains("ORG:NewLife", vcf);
        Assert.Contains("TEL;TYPE=WORK:+86-10-12345678", vcf);
        Assert.Contains("EMAIL;TYPE=WORK:jane@newlife.org", vcf);
        Assert.Contains("END:VCARD", vcf);
    }

    [Fact]
    [DisplayName("写入姓名分量（N 属性）")]
    public void Write_WithName_NamePropertyGenerated()
    {
        var contact = new VCardContact
        {
            FullName = "Wang Wei",
            Name = new VCardName { Family = "Wang", Given = "Wei" },
        };

        var writer = new VCardWriter();
        var vcf = writer.Build(contact);

        Assert.Contains("N:Wang\\;Wei\\;\\;\\;", vcf);
    }

    [Fact]
    [DisplayName("写入多联系人到同一文件")]
    public void WriteAll_MultipleContacts_AllInFile()
    {
        var contacts = new List<VCardContact>
        {
            new VCardContact { FullName = "Alice" },
            new VCardContact { FullName = "Bob" },
        };

        var writer = new VCardWriter();
        using var ms = new MemoryStream();
        writer.Write(contacts[0], ms);
        var bytes0 = ms.ToArray();

        var vcf0 = System.Text.Encoding.UTF8.GetString(bytes0);
        Assert.Contains("Alice", vcf0);
    }

    [Fact]
    [DisplayName("往返测试：写入后读取还原联系人")]
    public void RoundTrip_WriteAndRead()
    {
        var original = new VCardContact
        {
            FullName = "Round Trip",
            Organization = "TestOrg",
            Note = "Test note",
            Birthday = new DateTime(1985, 3, 20),
        };
        original.Phones.Add(new VCardPhone { Number = "1234567890", Type = "CELL" });
        original.Emails.Add(new VCardEmail { Address = "rt@test.com" });

        var writer = new VCardWriter();
        var vcf = writer.Build(original);

        var reader = new VCardReader();
        var parsed = reader.ParseAll(vcf);

        Assert.Single(parsed);
        var c = parsed[0];
        Assert.Equal("Round Trip", c.FullName);
        Assert.Equal("TestOrg", c.Organization);
        Assert.Equal("Test note", c.Note);
        Assert.NotNull(c.Birthday);
        Assert.Equal(1985, c.Birthday!.Value.Year);
        Assert.Single(c.Phones);
        Assert.Equal("1234567890", c.Phones[0].Number);
    }

    #endregion

    #region 集成测试

    [Fact]
    [DisplayName("集成：写入 vcf 文件并读取")]
    public void Integration_WriteFile_ThenReadFile()
    {
        var dir = Path.Combine("Bin", "UnitTest", "Artifacts");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "test_contacts.vcf");

        var contacts = new List<VCardContact>
        {
            new VCardContact
            {
                FullName = "张三",
                Organization = "新生命团队",
                Title = "高级工程师",
                Birthday = new DateTime(1990, 5, 20),
                Note = "NewLife.Office 集成测试联系人",
            },
            new VCardContact { FullName = "李四" },
        };
        contacts[0].Phones.Add(new VCardPhone { Number = "+86-138-0000-0001", Type = "CELL" });
        contacts[0].Emails.Add(new VCardEmail { Address = "zhangsan@newlife.org", Type = "WORK" });
        contacts[0].Addresses.Add(new VCardAddress
        {
            Street = "中关村大街 1 号",
            City = "北京",
            Country = "中国",
            Type = "WORK",
        });

        var writer = new VCardWriter();
        writer.WriteAll(contacts, path);

        Assert.True(File.Exists(path));

        var reader = new VCardReader();
        var parsed = reader.ReadAll(path);

        Assert.Equal(2, parsed.Count);
        Assert.Equal("张三", parsed[0].FullName);
        Assert.Equal("新生命团队", parsed[0].Organization);
        Assert.Single(parsed[0].Phones);
        Assert.Single(parsed[0].Emails);
        Assert.Single(parsed[0].Addresses);
        Assert.Equal("李四", parsed[1].FullName);
    }

    #endregion
}
