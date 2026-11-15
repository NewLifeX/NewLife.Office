using System;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using NewLife.Office;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>docProps/core.xml 文档属性读取测试</summary>
/// <remarks>openpyxl、Apache POI 等生成器把 dc/dcterms 命名空间声明在使用它的子元素上，
/// 此时按根元素前缀查找会得到 null 并抛 ArgumentNullException，导致整个 xlsx 无法解析。</remarks>
public class ExcelCorePropertiesTests
{
    static ExcelCorePropertiesTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    [DisplayName("core.xml 命名空间声明在子元素时仍能读取表格")]
    public void ReadRows_ChildNamespaceCoreProperties()
    {
        var path = CreateXlsx(childNamespaceCore: true);
        try
        {
            using (var reader = new ExcelReader(path))
            {
                Assert.Contains("销售明细", reader.Sheets!);
                var rows = reader.ReadRows("销售明细").ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("北京一店", Convert.ToString(rows[1][1]));

                // 文档属性同样要能读到，而不是整体失败
                Assert.Equal("openpyxl", reader.DocumentProperties?.Creator);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    [DisplayName("core.xml 命名空间声明在根元素时保持原有读取结果")]
    public void ReadRows_RootNamespaceCoreProperties()
    {
        var path = CreateXlsx(childNamespaceCore: false);
        try
        {
            using (var reader = new ExcelReader(path))
            {
                Assert.Contains("销售明细", reader.Sheets!);
                Assert.Equal("北京一店", Convert.ToString(reader.ReadRows("销售明细").Last()[1]));
                Assert.Equal("openpyxl", reader.DocumentProperties?.Creator);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    [DisplayName("OfficeFactory 按内容识别 core.xml 子元素命名空间的 xlsx")]
    public void ReadMarkdown_ChildNamespaceCoreProperties()
    {
        var path = CreateXlsx(childNamespaceCore: true);
        try
        {
            String? md;
            using (var fs = File.OpenRead(path))
                md = OfficeFactory.ReadMarkdown(fs, ".xlsx");

            Assert.False(String.IsNullOrEmpty(md));
            Assert.Contains("北京一店", md);
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>用 ExcelWriter 生成 xlsx；childNamespaceCore 为 true 时把 core.xml 改写成
    /// openpyxl 风格（命名空间声明在子元素上）</summary>
    private static String CreateXlsx(Boolean childNamespaceCore)
    {
        var path = Path.Combine(Path.GetTempPath(), $"office_core_{Guid.NewGuid():N}.xlsx");
        using (var fs = File.Create(path))
        {
            var writer = new ExcelWriter(fs);
            writer.WriteHeader("销售明细", ["日期", "门店"]);
            writer.WriteRow("销售明细", ["2026-11-01", "北京一店"]);
            writer.Save();
        }

        var coreXml = childNamespaceCore
            ? "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\">"
              + "<dc:creator xmlns:dc=\"http://purl.org/dc/elements/1.1/\">openpyxl</dc:creator>"
              + "<dcterms:created xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" "
              + "xsi:type=\"dcterms:W3CDTF\">2026-11-15T12:59:13Z</dcterms:created></cp:coreProperties>"
            : "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
              + "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
              + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
              + "<dc:creator>openpyxl</dc:creator></cp:coreProperties>";

        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        zip.GetEntry("docProps/core.xml")?.Delete();
        var entry = zip.CreateEntry("docProps/core.xml");
        using var stream = entry.Open();
        using var writer2 = new StreamWriter(stream, new UTF8Encoding(false));
        writer2.Write(coreXml);

        return path;
    }
}
