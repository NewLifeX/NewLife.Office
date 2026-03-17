using System.ComponentModel;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>EML 邮件格式集成测试</summary>
public class EmlTests : IntegrationTestBase
{
    [Fact, DisplayName("EML_复杂写入再读取往返")]
    public void Eml_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.eml");

        var msg = new EmlMessage
        {
            From = "sender@example.com",
            Subject = "NewLife.Office 集成测试邮件",
            Date = new DateTimeOffset(2024, 7, 1, 10, 30, 0, TimeSpan.FromHours(8)),
            TextBody = "这是纯文本正文。\r\n\r\n包含多行内容。\r\n第三行。",
            HtmlBody = "<html><body><h1>HTML正文</h1><p>这是<b>HTML</b>格式的邮件正文。</p><p>支持多种标签。</p></body></html>",
        };
        msg.To.Add("recipient1@example.com");
        msg.To.Add("recipient2@example.com");
        msg.Cc.Add("cc@example.com");

        // 附件
        msg.Attachments.Add(new EmlAttachment
        {
            FileName = "test.txt",
            ContentType = "text/plain",
            Data = Encoding.UTF8.GetBytes("这是附件内容"),
        });

        new EmlWriter().Write(msg, path);

        Assert.True(File.Exists(path));

        // 读取验证
        var reader = new EmlReader();
        var readMsg = reader.Read(path);
        Assert.Equal("sender@example.com", readMsg.From);
        Assert.Contains("集成测试邮件", readMsg.Subject);
        Assert.True(readMsg.To.Count >= 1);
        Assert.NotNull(readMsg.TextBody);
        Assert.Contains("纯文本正文", readMsg.TextBody);
        Assert.NotNull(readMsg.HtmlBody);
        Assert.Contains("HTML正文", readMsg.HtmlBody);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<EmlReader>(factoryReader);
    }
}
