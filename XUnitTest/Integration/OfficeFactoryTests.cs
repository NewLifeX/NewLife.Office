using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>OfficeFactory 工厂类自身功能测试</summary>
public class OfficeFactoryTests : IntegrationTestBase
{
    [Fact, DisplayName("IsSupported_支持的后缀返回true")]
    public void IsSupported_Returns_True_For_All_Supported()
    {
        var extensions = new[] { ".xlsx", ".xls", ".docx", ".doc", ".pptx", ".ppt", ".pdf", ".rtf", ".ods", ".epub", ".vcf", ".eml", ".ics", ".md", ".xps" };
        foreach (var ext in extensions)
        {
            Assert.True(OfficeFactory.IsSupported(ext), $"应支持 {ext}");
        }
    }

    [Fact, DisplayName("IsSupported_不带点号也返回true")]
    public void IsSupported_WithoutDot_Returns_True()
    {
        Assert.True(OfficeFactory.IsSupported("xlsx"));
        Assert.True(OfficeFactory.IsSupported("pdf"));
        Assert.True(OfficeFactory.IsSupported("md"));
    }

    [Fact, DisplayName("IsSupported_不支持的后缀返回false")]
    public void IsSupported_Returns_False_For_Unsupported()
    {
        Assert.False(OfficeFactory.IsSupported(".txt"));
        Assert.False(OfficeFactory.IsSupported(".csv"));
        Assert.False(OfficeFactory.IsSupported(".zip"));
        Assert.False(OfficeFactory.IsSupported(""));
        Assert.False(OfficeFactory.IsSupported(null!));
    }

    [Fact, DisplayName("IsSupported_大小写不敏感")]
    public void IsSupported_CaseInsensitive()
    {
        Assert.True(OfficeFactory.IsSupported(".XLSX"));
        Assert.True(OfficeFactory.IsSupported(".Pdf"));
        Assert.True(OfficeFactory.IsSupported("DOCX"));
    }

    [Fact, DisplayName("CreateReader_文件不存在抛FileNotFoundException")]
    public void CreateReader_FileNotFound_Throws()
    {
        Assert.Throws<FileNotFoundException>(() => OfficeFactory.CreateReader("not_exist.xlsx"));
    }

    [Fact, DisplayName("CreateReader_空路径抛ArgumentNullException")]
    public void CreateReader_NullPath_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => OfficeFactory.CreateReader(null!));
        Assert.Throws<ArgumentNullException>(() => OfficeFactory.CreateReader(""));
        Assert.Throws<ArgumentNullException>(() => OfficeFactory.CreateReader("   "));
    }

    [Fact, DisplayName("CreateReader_不支持格式抛NotSupportedException")]
    public void CreateReader_UnsupportedFormat_Throws()
    {
        var path = Path.Combine(OutputDir, "test.txt");
        File.WriteAllText(path, "hello");
        Assert.Throws<NotSupportedException>(() => OfficeFactory.CreateReader(path));
    }

    [Fact, DisplayName("SupportedExtensions_包含15种格式")]
    public void SupportedExtensions_Has15()
    {
        Assert.Equal(15, OfficeFactory.SupportedExtensions.Count);
    }
}
