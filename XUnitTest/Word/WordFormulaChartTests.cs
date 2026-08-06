using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>Word 公式与图表创建测试（W13，对标 Open XML SDK/Aspose.Words）</summary>
/// <remarks>
/// 覆盖 OMML 数学公式（分数/上下标/根式/积分）与 DrawingML 图表（柱状/折线/饼图）的写入。
/// </remarks>
public class WordFormulaChartTests
{
    /// <summary>生成 docx 字节，返回 (docx, document.xml 文本)</summary>
    private static (Byte[] Docx, String DocumentXml) BuildDocx(Action<WordWriter> build)
    {
        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
        {
            build(writer);
            writer.Save(ms);
        }
        var bytes = ms.ToArray();
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entry = za.GetEntry("word/document.xml");
        Assert.NotNull(entry);
        using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
        return (bytes, sr.ReadToEnd());
    }

    [Fact, DisplayName("OMML 公式：分数/上下标/根式/积分写入")]
    public void Omml_WriteStructures()
    {
        var (_, xml) = BuildDocx(w =>
        {
            w.AppendParagraph("公式示例");
            w.AppendFormula(OmmlFormula.Fraction("1", "2"));
            w.AppendFormula(OmmlFormula.SuperScript("x", "2"));
            w.AppendFormula(OmmlFormula.SubScript("a", "1"));
            w.AppendFormula(OmmlFormula.Radical("x + 1"));
            w.AppendFormula(OmmlFormula.Integral("x", "0", "1"));
        });

        // oMath 块存在
        Assert.Contains("<m:oMathPara>", xml);
        Assert.Contains("<m:oMath>", xml);
        // 各类结构
        Assert.Contains("<m:f>", xml);       // 分数
        Assert.Contains("<m:sSup>", xml);    // 上标
        Assert.Contains("<m:sSub>", xml);    // 下标
        Assert.Contains("<m:rad>", xml);     // 根式
        Assert.Contains("<m:nary>", xml);    // 积分
        // 文本内容
        Assert.Contains("x + 1", xml);
    }

    [Fact, DisplayName("OMML 公式：参数校验")]
    public void Omml_Validation()
    {
        using var writer = new WordWriter();
        Assert.Throws<ArgumentNullException>(() => writer.AppendFormula(null!));
    }

    [Fact, DisplayName("图表：柱状图写入含 chartSpace 与嵌入式数据")]
    public void Chart_ColumnWrite()
    {
        var (bytes, documentXml) = BuildDocx(w =>
        {
            w.AppendParagraph("销售图表");
            w.AppendChart(WordChart.Column("月度销售", ["一月", "二月", "三月"], "销售额", [100, 150, 200]));
        });

        // document.xml 引用 chart 部件
        Assert.Contains("c:chart", documentXml);
        Assert.Contains("rChart1", documentXml);

        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);

        // chart1.xml 存在且为 barChart（column）
        var chartEntry = za.GetEntry("word/charts/chart1.xml");
        Assert.NotNull(chartEntry);
        using (var cs = chartEntry!.Open())
        {
            var xml = new StreamReader(cs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("<c:barChart>", xml);
            Assert.Contains("月度销售", xml);
            Assert.Contains("一月", xml);
            Assert.Contains("externalData", xml);
        }

        // 嵌入式 xlsx 存在
        var xlsxEntry = za.GetEntry("word/embeddings/Microsoft_Excel_Worksheet1.xlsx");
        Assert.NotNull(xlsxEntry);
        Assert.True(xlsxEntry!.Length > 0);

        // chart 关系存在
        var relEntry = za.GetEntry("word/charts/_rels/chart1.xml.rels");
        Assert.NotNull(relEntry);
        using (var rs = relEntry!.Open())
        {
            var xml = new StreamReader(rs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("Microsoft_Excel_Worksheet1.xlsx", xml);
        }

        // document.xml.rels 含 chart 关系
        var docRelEntry = za.GetEntry("word/_rels/document.xml.rels");
        Assert.NotNull(docRelEntry);
        using (var drs = docRelEntry!.Open())
        {
            var xml = new StreamReader(drs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("relationships/chart", xml);
        }

        // Content_Types 注册
        var ctEntry = za.GetEntry("[Content_Types].xml");
        Assert.NotNull(ctEntry);
        using (var cts = ctEntry!.Open())
        {
            var xml = new StreamReader(cts, Encoding.UTF8).ReadToEnd();
            Assert.Contains("drawingml.chart+xml", xml);
            Assert.Contains("Extension=\"xlsx\"", xml);
        }
    }

    [Fact, DisplayName("图表：折线/饼图类型")]
    public void Chart_LineAndPie()
    {
        var (_, xml) = BuildDocx(w =>
        {
            w.AppendChart(WordChart.Line("趋势", ["Q1", "Q2", "Q3"], "指标", [1, 2, 3]));
            w.AppendChart(WordChart.Pie("占比", ["A", "B"], [60, 40]));
        });

        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.AppendChart(WordChart.Line("趋势", ["Q1", "Q2", "Q3"], "指标", [1, 2, 3]));
        writer.AppendChart(WordChart.Pie("占比", ["A", "B"], [60, 40]));
        writer.Save(ms);
        var bytes = ms.ToArray();

        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using (var cs = za.GetEntry("word/charts/chart1.xml")!.Open())
        {
            var chartXml = new StreamReader(cs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("<c:lineChart>", chartXml);
        }
        using (var cs = za.GetEntry("word/charts/chart2.xml")!.Open())
        {
            var chartXml = new StreamReader(cs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("<c:pieChart>", chartXml);
        }
    }
}
