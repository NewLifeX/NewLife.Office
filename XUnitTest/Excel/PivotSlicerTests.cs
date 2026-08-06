using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>数据透视表切片器测试（M28，关联透视表字段）</summary>
/// <remarks>
/// 覆盖 PivotBuilder.AddSlicer 的透视表关联写入（slicerCache pivot 模式 + slicer 部件 + sheet 引用）、
/// 读取还原与参数校验。
/// </remarks>
public class PivotSlicerTests
{
    /// <summary>构造含切片器的透视表 xlsx 字节</summary>
    private static Byte[] BuildPivotXlsx()
    {
        using var ms = new MemoryStream();
        var builder = new PivotBuilder();
        builder.SetSourceData(
            new[] { "Region", "Product", "Sales" },
            new List<Object?[]>
            {
                new Object?[] { "East", "A", 100 },
                new Object?[] { "West", "B", 200 },
                new Object?[] { "East", "C", 150 },
            });
        builder.AddRowField("Product");
        builder.AddDataField("Sales", PivotSummaryFunction.Sum);
        builder.AddSlicer("Region");
        builder.Save(ms);
        return ms.ToArray();
    }

    [Fact, DisplayName("透视表切片器：写入与读取还原（列名/透视表名/标题）")]
    public void PivotSlicer_WriteRead()
    {
        var bytes = BuildPivotXlsx();

        using var reader = new ExcelReader(new MemoryStream(bytes), Encoding.UTF8);
        var slicers = reader.ReadSlicers("Pivot");

        Assert.Single(slicers);
        var s = slicers[0];
        Assert.Equal("Region", s.ColumnName);
        Assert.Equal("PivotTable1", s.TableName);
        Assert.Equal("切片器_Region", s.Name);
        Assert.Equal("Region", s.Caption);
        Assert.Equal("SlicerStyleLight1", s.Style);
        Assert.False(String.IsNullOrEmpty(s.CacheName));
    }

    [Fact, DisplayName("透视表切片器：部件齐全（pivot 模式 slicerCache/slicer/引用/类型）")]
    public void PivotSlicer_PartsComplete()
    {
        var bytes = BuildPivotXlsx();
        using var ms = new MemoryStream(bytes);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);

        // slicerCache 部件存在且为透视表模式（pivot="1" + pivotTables 引用）
        var cacheEntry = za.GetEntry("xl/slicerCaches/slicerCache1.xml");
        Assert.NotNull(cacheEntry);
        using (var cs = cacheEntry!.Open())
        {
            var doc = XDocument.Load(cs);
            Assert.NotNull(doc.Root);
            Assert.Equal("1", doc.Root!.Attribute("pivot")?.Value ?? doc.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "data")?.Attribute("pivot")?.Value);
            var pt = doc.Root!.Descendants().FirstOrDefault(e => e.Name.LocalName == "pivotTable");
            Assert.NotNull(pt);
            Assert.Equal("PivotTable1", pt!.Attribute("name")?.Value);
            var column = doc.Root!.Descendants().FirstOrDefault(e => e.Name.LocalName == "column");
            Assert.Equal("Region", column?.Attribute("name")?.Value);
            var items = doc.Root!.Descendants().FirstOrDefault(e => e.Name.LocalName == "items");
            Assert.NotNull(items);
            Assert.Equal("2", items!.Attribute("count")?.Value);
        }

        // slicer 部件存在
        var slicerEntry = za.GetEntry("xl/slicers/slicer1.xml");
        Assert.NotNull(slicerEntry);
        using (var ss = slicerEntry!.Open())
        {
            var doc = XDocument.Load(ss);
            Assert.NotNull(doc.Root);
            Assert.Equal("切片器_Region", doc.Root!.Attribute("name")?.Value);
        }

        // 透视表工作表（sheet2）extLst 含 x14:slicerList
        var sheetEntry = za.GetEntry("xl/worksheets/sheet2.xml");
        Assert.NotNull(sheetEntry);
        using (var shs = sheetEntry!.Open())
        {
            var text = new StreamReader(shs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("x14:slicerList", text);
            Assert.Contains("A8765BA9-456A-4DAB-B4F3-ACF838C121DE", text);
        }

        // 透视表工作表 rels 含 slicer 与 slicerCache 关系
        var relEntry = za.GetEntry("xl/worksheets/_rels/sheet2.xml.rels");
        Assert.NotNull(relEntry);
        using (var rs = relEntry!.Open())
        {
            var text = new StreamReader(rs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("/relationships/slicer", text);
            Assert.Contains("/relationships/slicerCache", text);
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

    [Fact, DisplayName("透视表切片器：参数校验与名称去重")]
    public void PivotSlicer_Validation()
    {
        var builder = new PivotBuilder();
        Assert.Throws<ArgumentNullException>(() => builder.AddSlicer(""));

        builder.SetSourceData(
            new[] { "Region", "Product", "Sales" },
            new List<Object?[]>
            {
                new Object?[] { "East", "A", 100 },
                new Object?[] { "West", "B", 200 },
            });
        var s1 = builder.AddSlicer("Region");
        var s2 = builder.AddSlicer("Region");
        Assert.NotEqual(s1.Name, s2.Name);
    }
}
