using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>PDF 写入器测试</summary>
public class PdfWriterTests
{
    static PdfWriterTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    #region 基础输出
    [Fact, DisplayName("生成基础 PDF 文件结构正确")]
    public void SavePdf_BasicStructure()
    {
        using var ms = new MemoryStream();
        var writer = new PdfWriter();
        writer.BeginPage();
        writer.DrawText("Hello PDF", 56, 780, 12);
        writer.Save(ms);

        ms.Position = 0;
        var bytes = ms.ToArray();
        var text = Encoding.Latin1.GetString(bytes);

        Assert.StartsWith("%PDF-1.4", text);
        Assert.Contains("%%EOF", text);
        Assert.Contains("/Type /Catalog", text);
        Assert.Contains("/Type /Pages", text);
    }

    [Fact, DisplayName("生成带 Info 元数据的 PDF")]
    public void SavePdf_WithInfoDict()
    {
        using var ms = new MemoryStream();
        var writer = new PdfWriter
        {
            DocumentTitle = "TestDoc",
            DocumentAuthor = "NewLife",
        };
        writer.BeginPage();
        writer.DrawText("Info test", 56, 780, 12);
        writer.Save(ms);

        var text = Encoding.Latin1.GetString(ms.ToArray());
        Assert.Contains("/Title", text);
        Assert.Contains("/Author", text);
        Assert.Contains("TestDoc", text);
    }
    #endregion

    #region 加密
    [Fact, DisplayName("设置用户密码后 PDF 包含加密字典")]
    public void SavePdf_WithUserPassword_HasEncryptDict()
    {
        using var ms = new MemoryStream();
        var writer = new PdfWriter { UserPassword = "open123" };
        writer.BeginPage();
        writer.DrawText("Encrypted content", 56, 780, 12);
        writer.Save(ms);

        var text = Encoding.Latin1.GetString(ms.ToArray());
        Assert.Contains("/Filter /Standard", text);
        Assert.Contains("/V 2", text);
        Assert.Contains("/R 3", text);
        Assert.Contains("/Length 128", text);
        Assert.Contains("/Encrypt", text);
        // trailer 应包含 /ID
        Assert.Contains("/ID [<", text);
    }

    [Fact, DisplayName("不设置密码时 PDF 不含加密字典")]
    public void SavePdf_NoPassword_NoEncryptDict()
    {
        using var ms = new MemoryStream();
        var writer = new PdfWriter();
        writer.BeginPage();
        writer.DrawText("Plain content", 56, 780, 12);
        writer.Save(ms);

        var text = Encoding.Latin1.GetString(ms.ToArray());
        Assert.DoesNotContain("/Filter /Standard", text);
        Assert.DoesNotContain("/Encrypt", text);
    }

    [Fact, DisplayName("设置所有者密码后 PDF 包含加密字典")]
    public void SavePdf_WithOwnerPassword_HasEncryptDict()
    {
        using var ms = new MemoryStream();
        var writer = new PdfWriter
        {
            UserPassword = "",
            OwnerPassword = "owner456",
            Permissions = -3904, // 允许打印，禁止修改
        };
        writer.BeginPage();
        writer.DrawText("Permission-controlled PDF", 56, 780, 12);
        writer.Save(ms);

        var text = Encoding.Latin1.GetString(ms.ToArray());
        Assert.Contains("/Filter /Standard", text);
        Assert.Contains("/P -3904", text);
        Assert.Contains("/O <", text);
        Assert.Contains("/U <", text);
    }

    [Fact, DisplayName("加密 PDF 输出大小合理（包含额外加密字典对象）")]
    public void SavePdf_WithPassword_LargerThanPlain()
    {
        using var msPlain = new MemoryStream();
        using var msEnc = new MemoryStream();

        var plain = new PdfWriter();
        plain.BeginPage();
        plain.DrawText("Test", 56, 780, 12);
        plain.Save(msPlain);

        var enc = new PdfWriter { UserPassword = "pw" };
        enc.BeginPage();
        enc.DrawText("Test", 56, 780, 12);
        enc.Save(msEnc);

        // 加密版本应大于纯文本版本（额外的加密字典对象）
        Assert.True(msEnc.Length > msPlain.Length,
            $"Encrypted PDF ({msEnc.Length}) should be larger than plain ({msPlain.Length})");
    }
    #endregion
}
