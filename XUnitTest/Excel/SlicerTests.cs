using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml.Linq;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>Excel 切片器测试（M28，对标 EPPlus/Aspose.Cells）</summary>
/// <remarks>
/// 覆盖切片器的写入（slicerCache + slicer 部件 + sheet 引用 + 关系）、
/// 读取还原（名称/标题/列名/表格名）与参数校验。
/// </remarks>
public class SlicerTests
{
    /// <summary>构造含切片器的 xlsx 字节</summary>
    private static Byte[] BuildXlsx()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.SheetName = "Sheet1";
            writer.WriteHeader("Sheet1", new[] { "产品", "销量" });
            writer.WriteRows("Sheet1", new Object?[][]
            {
                ["苹果", 10],
                ["香蕉", 20],
                ["苹果", 30],
            });
            writer.AddTable("Sheet1", "A1:B4", "销售表", "TableStyleMedium9", new[] { "产品", "销量" });
            writer.AddSlicer("Sheet1", "销售表", "产品");
            writer.Save();
        }
        return ms.ToArray();
    }

    [Fact, DisplayName("切片器：写入与读取还原（列名/表格名/标题）")]
    public void Slicer_WriteRead()
    {
        var bytes = BuildXlsx();

        using var reader = new ExcelReader(new MemoryStream(bytes), Encoding.UTF8);
        var slicers = reader.ReadSlicers("Sheet1");

        Assert.Single(slicers);
        var s = slicers[0];
        Assert.Equal("产品", s.ColumnName);
        Assert.Equal("销售表", s.TableName);
        Assert.Equal("切片器_产品", s.Name);
        Assert.Equal("产品", s.Caption);
        Assert.Equal("SlicerStyleLight1", s.Style);
        Assert.False(String.IsNullOrEmpty(s.CacheName));
    }

    [Fact, DisplayName("切片器：部件齐全（slicerCache/slicer/引用/关系/类型）")]
    public void Slicer_PartsComplete()
    {
        var bytes = BuildXlsx();
        using var ms = new MemoryStream(bytes);
        using var za = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read);

        // slicerCache 部件存在且含唯一值 items
        var cacheEntry = za.GetEntry("xl/slicerCaches/slicerCache1.xml");
        Assert.NotNull(cacheEntry);
        using (var cs = cacheEntry!.Open())
        {
            var doc = XDocument.Load(cs);
            Assert.NotNull(doc.Root);
            var items = doc.Root!.Descendants().FirstOrDefault(e => e.Name.LocalName == "items");
            Assert.NotNull(items);
            // 产品列唯一值：苹果/香蕉 → 2 项
            Assert.Equal("2", items!.Attribute("count")?.Value);
            var column = doc.Root!.Descendants().FirstOrDefault(e => e.Name.LocalName == "column");
            Assert.Equal("产品", column?.Attribute("name")?.Value);
        }

        // slicer 部件存在
        var slicerEntry = za.GetEntry("xl/slicers/slicer1.xml");
        Assert.NotNull(slicerEntry);
        using (var ss = slicerEntry!.Open())
        {
            var doc = XDocument.Load(ss);
            Assert.NotNull(doc.Root);
            Assert.Equal("切片器_产品", doc.Root!.Attribute("name")?.Value);
            Assert.Equal("SlicerStyleLight1", doc.Root!.Attribute("style")?.Value);
        }

        // sheet extLst 含 x14:slicerList
        var sheetEntry = za.GetEntry("xl/worksheets/sheet1.xml");
        Assert.NotNull(sheetEntry);
        using (var shs = sheetEntry!.Open())
        {
            var text = new StreamReader(shs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("x14:slicerList", text);
            Assert.Contains("A8765BA9-456A-4DAB-B4F3-ACF838C121DE", text);
        }

        // sheet rels 含 slicer 关系
        var relEntry = za.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
        Assert.NotNull(relEntry);
        using (var rs = relEntry!.Open())
        {
            var text = new StreamReader(rs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("/relationships/slicer", text);
        }

        // workbook rels 含 slicerCache 关系
        var wbRelEntry = za.GetEntry("xl/_rels/workbook.xml.rels");
        Assert.NotNull(wbRelEntry);
        using (var ws = wbRelEntry!.Open())
        {
            var text = new StreamReader(ws, Encoding.UTF8).ReadToEnd();
            Assert.Contains("relationships/slicerCache", text);
        }

        // Content_Types 注册
        var ctEntry = za.GetEntry("[Content_Types].xml");
        Assert.NotNull(ctEntry);
        using (var cts = ctEntry!.Open())
        {
            var text = new StreamReader(cts, Encoding.UTF8).ReadToEnd();
            Assert.Contains("vnd.ms-excel.slicer+xml", text);
            Assert.Contains("vnd.ms-excel.slicerCache+xml", text);
        }
    }

    [Fact, DisplayName("切片器：参数校验与名称去重")]
    public void Slicer_Validation()
    {
        using var ms = new MemoryStream();
        using var writer = new ExcelWriter(ms);
        writer.SheetName = "Sheet1";

        Assert.Throws<ArgumentNullException>(() => writer.AddSlicer("Sheet1", "", "产品"));
        Assert.Throws<ArgumentNullException>(() => writer.AddSlicer("Sheet1", "销售表", ""));

        // 同名切片器自动去重
        writer.AddTable("Sheet1", "A1:B4", "销售表", "TableStyleMedium9", new[] { "产品", "销量" });
        var s1 = writer.AddSlicer("Sheet1", "销售表", "产品");
        var s2 = writer.AddSlicer("Sheet1", "销售表", "产品");
        Assert.NotEqual(s1.Name, s2.Name);
    }
}
