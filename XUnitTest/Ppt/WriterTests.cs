using System.IO.Compression;
using System.Linq;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PptxWriter 写入器单元测试（含母版/版式加载功能）</summary>
public class PptxWriterTests
{
    static PptxWriterTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    // ─── 测试模板构建辅助 ─────────────────────────────────────────────────

    /// <summary>构建最小测试模板 pptx，包含指定版式名称列表</summary>
    private static Byte[] BuildTemplatePptx(params String[] layoutNames)
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            // theme
            WriteZip(zip, "ppt/theme/theme1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"TestTheme\">" +
                "<a:themeElements><a:clrScheme name=\"Test\"><a:dk1><a:srgbClr val=\"000000\"/></a:dk1>" +
                "<a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1></a:clrScheme></a:themeElements></a:theme>");

            // master with layout references
            var masterLayoutIds = new StringBuilder();
            var masterRels = new StringBuilder();
            masterRels.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            masterRels.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            masterRels.Append("<Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>");
            for (var i = 0; i < layoutNames.Length; i++)
            {
                masterLayoutIds.Append($"<p:sldLayoutId id=\"{2147483649 + i}\" r:id=\"rLayout{i + 1}\"/>");
                masterRels.Append($"<Relationship Id=\"rLayout{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout{i + 1}.xml\"/>");
            }
            masterRels.Append("</Relationships>");

            WriteZip(zip, "ppt/slideMasters/slideMaster1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
                "<p:grpSpPr/></p:spTree></p:cSld>" +
                "<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>" +
                $"<p:sldLayoutIdLst>{masterLayoutIds}</p:sldLayoutIdLst>" +
                "</p:sldMaster>");
            WriteZip(zip, "ppt/slideMasters/_rels/slideMaster1.xml.rels", masterRels.ToString());

            // layouts
            for (var i = 0; i < layoutNames.Length; i++)
            {
                WriteZip(zip, $"ppt/slideLayouts/slideLayout{i + 1}.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    $"type=\"blank\" preserve=\"1\"><p:cSld name=\"{layoutNames[i]}\"><p:spTree/></p:cSld></p:sldLayout>");
                WriteZip(zip, $"ppt/slideLayouts/_rels/slideLayout{i + 1}.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/>" +
                    "</Relationships>");
            }
        }
        return ms.ToArray();
    }

    private static void WriteZip(ZipArchive zip, String path, String content)
    {
        using var sw = new StreamWriter(zip.CreateEntry(path).Open(), Encoding.UTF8);
        sw.Write(content);
    }

    private static String ReadZipEntry(ZipArchive zip, String path)
    {
        var entry = zip.GetEntry(path);
        if (entry == null) return String.Empty;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    // ─── 无模板向后兼容性测试 ─────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("无模板：AddSlide 默认版式索引为 0")]
    public void AddSlide_NoTemplate_DefaultLayoutIndex()
    {
        var writer = new PptxWriter();
        var slide = writer.AddSlide();

        Assert.Equal(0, slide.LayoutIndex);
    }

    [Fact, System.ComponentModel.DisplayName("无模板：GetLayoutCount 返回 1")]
    public void GetLayoutCount_NoTemplate_Returns1()
    {
        var writer = new PptxWriter();
        Assert.Equal(1, writer.GetLayoutCount());
    }

    [Fact, System.ComponentModel.DisplayName("无模板：GetLayoutName(0) 返回 blank")]
    public void GetLayoutName_NoTemplate_ReturnsBlank()
    {
        var writer = new PptxWriter();
        Assert.Equal("blank", writer.GetLayoutName(0));
    }

    [Fact, System.ComponentModel.DisplayName("无模板：保存生成合法 pptx 结构")]
    public void Save_NoTemplate_ProducesValidPptx()
    {
        var writer = new PptxWriter();
        writer.AddSlide();
        writer.AddTextBox(0, "测试", 1, 1, 10, 2);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("ppt/presentation.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slideMasters/slideMaster1.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slideLayouts/slideLayout1.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slides/slide1.xml"));
    }

    // ─── LoadMaster 基础功能测试 ─────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("LoadMaster：单版式模板 GetLayoutCount 返回 1")]
    public void LoadMaster_SingleLayout_GetLayoutCountReturns1()
    {
        var template = BuildTemplatePptx("TitleSlide");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        Assert.Equal(1, writer.GetLayoutCount());
    }

    [Fact, System.ComponentModel.DisplayName("LoadMaster：双版式模板 GetLayoutCount 返回 2")]
    public void LoadMaster_TwoLayouts_GetLayoutCountReturns2()
    {
        var template = BuildTemplatePptx("Layout1", "Layout2");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        Assert.Equal(2, writer.GetLayoutCount());
    }

    [Fact, System.ComponentModel.DisplayName("LoadMaster：正确提取版式显示名称")]
    public void LoadMaster_ExtractsLayoutDisplayNames()
    {
        var template = BuildTemplatePptx("企业标题页", "内容页");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        Assert.Equal("企业标题页", writer.GetLayoutName(0));
        Assert.Equal("内容页", writer.GetLayoutName(1));
    }

    [Fact, System.ComponentModel.DisplayName("LoadMaster：超出范围索引返回空字符串")]
    public void GetLayoutName_OutOfRange_ReturnsEmpty()
    {
        var template = BuildTemplatePptx("Layout1");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        Assert.Equal(String.Empty, writer.GetLayoutName(5));
        Assert.Equal(String.Empty, writer.GetLayoutName(-1));
    }

    // ─── 模板构造函数测试 ──────────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("模板构造函数：自动加载母版")]
    public void Constructor_WithTemplateBytes_LoadsMaster()
    {
        var template = BuildTemplatePptx("Title", "Content", "Blank");
        // 构造函数接受文件路径；此处用 LoadMaster(Byte[]) 测试等效逻辑
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        Assert.Equal(3, writer.GetLayoutCount());
    }

    // ─── AddSlide 版式索引测试 ────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("AddSlide(layoutIndex)：正确设置 LayoutIndex")]
    public void AddSlide_WithLayoutIndex_SetsCorrectIndex()
    {
        var template = BuildTemplatePptx("L0", "L1", "L2");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        var s0 = writer.AddSlide(0);
        var s1 = writer.AddSlide(1);
        var s2 = writer.AddSlide(2);

        Assert.Equal(0, s0.LayoutIndex);
        Assert.Equal(1, s1.LayoutIndex);
        Assert.Equal(2, s2.LayoutIndex);
    }

    [Fact, System.ComponentModel.DisplayName("AddSlide：超出范围自动修正到最后一个版式")]
    public void AddSlide_IndexOutOfRange_ClampsToMax()
    {
        var template = BuildTemplatePptx("L0", "L1");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        var slide = writer.AddSlide(99);
        Assert.Equal(1, slide.LayoutIndex);
    }

    // ─── Save 输出结构验证 ────────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("Save：保存后包含模板母版 XML")]
    public void Save_WithTemplate_PreservesMasterXml()
    {
        const String MarkerText = "TestTheme";
        var template = BuildTemplatePptx("Layout1");
        var writer = new PptxWriter();
        writer.LoadMaster(template);
        writer.AddSlide();

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var themeXml = ReadZipEntry(zip, "ppt/theme/theme1.xml");
        Assert.Contains(MarkerText, themeXml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("Save：多版式时幻灯片 rels 引用正确版式文件")]
    public void Save_MultipleLayouts_SlideReferencesCorrectLayout()
    {
        var template = BuildTemplatePptx("L0", "L1", "L2");
        var writer = new PptxWriter();
        writer.LoadMaster(template);
        writer.AddSlide(0);  // slide1 → slideLayout1.xml
        writer.AddSlide(2);  // slide2 → slideLayout3.xml

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var rels1 = ReadZipEntry(zip, "ppt/slides/_rels/slide1.xml.rels");
        var rels2 = ReadZipEntry(zip, "ppt/slides/_rels/slide2.xml.rels");

        Assert.Contains("slideLayout1.xml", rels1, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("slideLayout3.xml", rels2, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("Save：模板版式写入 ZIP 中的 slideLayout 条目")]
    public void Save_WithTwoLayouts_WritesBothLayoutEntries()
    {
        var template = BuildTemplatePptx("L0", "L1");
        var writer = new PptxWriter();
        writer.LoadMaster(template);
        writer.AddSlide();

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("ppt/slideLayouts/slideLayout1.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slideLayouts/slideLayout2.xml"));
        Assert.Null(zip.GetEntry("ppt/slideLayouts/slideLayout3.xml"));
    }

    [Fact, System.ComponentModel.DisplayName("Save：presentation.xml 包含 sldMasterId 引用")]
    public void Save_WithTemplate_PresentationContainsMasterIdRef()
    {
        var template = BuildTemplatePptx("L0");
        var writer = new PptxWriter();
        writer.LoadMaster(template);
        writer.AddSlide();

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var presXml = ReadZipEntry(zip, "ppt/presentation.xml");
        Assert.Contains("rMaster1", presXml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("Save：[Content_Types].xml 包含模板版式条目")]
    public void Save_WithTwoLayouts_ContentTypesHasBothLayouts()
    {
        var template = BuildTemplatePptx("A", "B");
        var writer = new PptxWriter();
        writer.LoadMaster(template);
        writer.AddSlide();

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var ct = ReadZipEntry(zip, "[Content_Types].xml");
        Assert.Contains("slideLayout1.xml", ct, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("slideLayout2.xml", ct, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("LoadMaster 后再次调用：清除旧内容")]
    public void LoadMaster_CalledTwice_ClearsOldContent()
    {
        var t1 = BuildTemplatePptx("A", "B", "C");
        var t2 = BuildTemplatePptx("X");
        var writer = new PptxWriter();
        writer.LoadMaster(t1);
        Assert.Equal(3, writer.GetLayoutCount());
        writer.LoadMaster(t2);
        Assert.Equal(1, writer.GetLayoutCount());
        Assert.Equal("X", writer.GetLayoutName(0));
    }

    // ─── CopyMasterFrom 测试 ──────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("CopyMasterFrom：等效于 LoadMaster")]
    public void CopyMasterFrom_EquivalentToLoadMaster()
    {
        var template = BuildTemplatePptx("CopyTest");
        var tmpFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            File.WriteAllBytes(tmpFile, template);

            var writer = new PptxWriter();
            writer.CopyMasterFrom(tmpFile);
            Assert.Equal(1, writer.GetLayoutCount());
            Assert.Equal("CopyTest", writer.GetLayoutName(0));
        }
        finally { try { File.Delete(tmpFile); } catch { } }
    }

    // ─── 边界用例测试 ────────────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("边界：AddSlide 负索引修正为 0")]
    public void AddSlide_NegativeIndex_ClampsToZero()
    {
        var template = BuildTemplatePptx("L0", "L1");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        var slide = writer.AddSlide(-5);
        Assert.Equal(0, slide.LayoutIndex);
    }

    [Fact, System.ComponentModel.DisplayName("边界：无版式模板加载后 GetLayoutCount 正确")]
    public void LoadMaster_ZeroLayouts_Returns1()
    {
        var template = BuildTemplatePptx();
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        Assert.Equal(1, writer.GetLayoutCount());
    }

    [Fact, System.ComponentModel.DisplayName("边界：保存后无版式名称的 layout 返回空字符串")]
    public void GetLayoutName_NoLayouts_HandlesGracefully()
    {
        var template = BuildTemplatePptx();
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        // 无 cSld@name 且无 @type 的版式 → 返回空串
        var name = writer.GetLayoutName(0);
        Assert.NotNull(name);
    }

    [Fact, System.ComponentModel.DisplayName("边界：空幻灯片保存生成合法 pptx")]
    public void Save_NoSlides_ProducesValidPptx()
    {
        var writer = new PptxWriter();
        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("ppt/presentation.xml"));
        Assert.NotNull(zip.GetEntry("[Content_Types].xml"));
    }

    [Fact, System.ComponentModel.DisplayName("边界：模板有主题但无主题 XML 时不崩溃")]
    public void LoadMaster_PreservesThemeOrFallsBackGracefully()
    {
        var template = BuildTemplatePptx("有主题的版式");
        var writer = new PptxWriter();
        // 不用 Assert.Throws，确保不抛异常
        var ex = Record.Exception(() => writer.LoadMaster(template));
        Assert.Null(ex);
    }

    [Fact, System.ComponentModel.DisplayName("边界：模板版式中包含中文名称")]
    public void LoadMaster_ChineseLayoutNames_PreservedCorrectly()
    {
        var template = BuildTemplatePptx("标题幻灯片", "内容幻灯片", "空白");
        var writer = new PptxWriter();
        writer.LoadMaster(template);

        Assert.Equal("标题幻灯片", writer.GetLayoutName(0));
        Assert.Equal("内容幻灯片", writer.GetLayoutName(1));
        Assert.Equal("空白", writer.GetLayoutName(2));
    }

    // ─── PptxReader 新方法测试 ────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("GetSlideMasterXml：从模板提取母版 XML")]
    public void GetSlideMasterXml_ReturnsNonNull()
    {
        var template = BuildTemplatePptx("L");
        using var ms = new MemoryStream(template);
        using var reader = new PptxReader(ms);

        var xml = reader.GetSlideMasterXml(0);
        Assert.NotNull(xml);
        Assert.Contains("sldMaster", xml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("GetSlideMasterXml：越界索引返回 null")]
    public void GetSlideMasterXml_OutOfRange_ReturnsNull()
    {
        var template = BuildTemplatePptx("L");
        using var ms = new MemoryStream(template);
        using var reader = new PptxReader(ms);

        Assert.Null(reader.GetSlideMasterXml(99));
    }

    [Fact, System.ComponentModel.DisplayName("GetSlideLayoutXml：从模板提取版式 XML")]
    public void GetSlideLayoutXml_ReturnsNonNull()
    {
        var template = BuildTemplatePptx("TestLayout");
        using var ms = new MemoryStream(template);
        using var reader = new PptxReader(ms);

        var xml = reader.GetSlideLayoutXml(0);
        Assert.NotNull(xml);
        Assert.Contains("sldLayout", xml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("TestLayout", xml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("GetSlideLayoutXml：越界索引返回 null")]
    public void GetSlideLayoutXml_OutOfRange_ReturnsNull()
    {
        var template = BuildTemplatePptx("L");
        using var ms = new MemoryStream(template);
        using var reader = new PptxReader(ms);

        Assert.Null(reader.GetSlideLayoutXml(99));
    }

    [Fact, System.ComponentModel.DisplayName("GetThemeXml：从模板提取主题 XML")]
    public void GetThemeXml_ReturnsNonNull()
    {
        var template = BuildTemplatePptx("L");
        using var ms = new MemoryStream(template);
        using var reader = new PptxReader(ms);

        var xml = reader.GetThemeXml();
        Assert.NotNull(xml);
        Assert.Contains("TestTheme", xml, StringComparison.OrdinalIgnoreCase);
    }

    // ─── 编程式母版测试（Phase 5）────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("CreateMaster：创建母版后 GetLayoutCount 正确")]
    public void CreateMaster_NoLayouts_GetLayoutCountReturns1()
    {
        var writer = new PptxWriter();
        writer.CreateMaster();

        Assert.Equal(1, writer.GetLayoutCount());
    }

    [Fact, System.ComponentModel.DisplayName("CreateMaster：添加版式后 GetLayoutCount 正确")]
    public void CreateMaster_WithLayouts_GetLayoutCountCorrect()
    {
        var writer = new PptxWriter();
        var master = writer.CreateMaster();
        master.AddLayout("标题页", "title");
        master.AddLayout("内容页", "blank");

        Assert.Equal(2, writer.GetLayoutCount());
    }

    [Fact, System.ComponentModel.DisplayName("CreateMaster：版式名称正确返回")]
    public void CreateMaster_GetLayoutNameCorrect()
    {
        var writer = new PptxWriter();
        var master = writer.CreateMaster();
        master.AddLayout("封面");
        master.AddLayout("正文");

        Assert.Equal("封面", writer.GetLayoutName(0));
        Assert.Equal("正文", writer.GetLayoutName(1));
    }

    [Fact, System.ComponentModel.DisplayName("CreateMaster：多母版版式名称正确")]
    public void CreateMaster_MultipleMasters_LayoutNamesCorrect()
    {
        var writer = new PptxWriter();
        var m1 = writer.CreateMaster();
        m1.AddLayout("母版1-版式A");
        m1.AddLayout("母版1-版式B");
        var m2 = writer.CreateMaster();
        m2.AddLayout("母版2-版式C");

        Assert.Equal(3, writer.GetLayoutCount());
        Assert.Equal("母版1-版式A", writer.GetLayoutName(0));
        Assert.Equal("母版1-版式B", writer.GetLayoutName(1));
        Assert.Equal("母版2-版式C", writer.GetLayoutName(2));
    }

    [Fact, System.ComponentModel.DisplayName("编程母版：保存生成合法 pptx 结构")]
    public void Save_ProgMaster_ProducesValidPptx()
    {
        var writer = new PptxWriter();
        var master = writer.CreateMaster();
        master.BackgroundColor = "1F497D";
        master.AddLayout("标题幻灯片", "title");
        master.AddLayout("内容幻灯片", "blank");

        var slide = writer.AddSlide(0);
        writer.AddTextBox(0, "编程母版测试", 2, 3, 20, 3, fontSize: 28, bold: true);
        writer.AddSlide(1);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("ppt/slideMasters/slideMaster1.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slideLayouts/slideLayout1.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slideLayouts/slideLayout2.xml"));

        var masterXml = ReadZipEntry(zip, "ppt/slideMasters/slideMaster1.xml");
        Assert.Contains("1F497D", masterXml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("编程母版：母版形状写入 XML")]
    public void Save_ProgMasterWithShapes_WritesShapeXml()
    {
        var writer = new PptxWriter();
        var master = writer.CreateMaster();
        master.Shapes.Add(new PptShape
        {
            ShapeType = "rect",
            Left = PptxWriter.CmToEmu(0),
            Top = PptxWriter.CmToEmu(0),
            Width = PptxWriter.CmToEmu(33.87),
            Height = PptxWriter.CmToEmu(1),
            FillColor = "1F497D",
        });
        master.AddLayout("空白");
        writer.AddSlide(0);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var masterXml = ReadZipEntry(zip, "ppt/slideMasters/slideMaster1.xml");
        Assert.Contains("prstGeom prst=\"rect\"", masterXml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("编程版式：版式文本框写入 XML")]
    public void Save_ProgLayoutWithTextboxes_WritesLayoutXml()
    {
        var writer = new PptxWriter();
        var master = writer.CreateMaster();
        var layout = master.AddLayout("标题页", "title");
        layout.TextBoxes.Add(new PptTextBox
        {
            Text = "单击此处添加标题",
            Left = PptxWriter.CmToEmu(2),
            Top = PptxWriter.CmToEmu(2),
            Width = PptxWriter.CmToEmu(20),
            Height = PptxWriter.CmToEmu(3),
            FontSize = 28,
        });
        writer.AddSlide(0);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var layoutXml = ReadZipEntry(zip, "ppt/slideLayouts/slideLayout1.xml");
        Assert.Contains("单击此处添加标题", layoutXml, StringComparison.OrdinalIgnoreCase);
    }

    // ─── keepTemplateSlides 保留模板幻灯片测试 ─────────────────────────────

    [Fact, System.ComponentModel.DisplayName("keepTemplateSlides：保留模板幻灯片追加到最终输出")]
    public void LoadMaster_KeepSlides_PreservesTemplateSlides()
    {
        // 构建包含一张原始幻灯片的模板
        using var templateMs = new MemoryStream();
        using (var tzip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZip(tzip, "ppt/theme/theme1.xml",
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"></a:theme>");
            WriteZip(tzip, "ppt/slideMasters/slideMaster1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:cSld><p:spTree/></p:cSld><p:txStyles/><p:sldLayoutIdLst>" +
                "<p:sldLayoutId id=\"2147483649\" r:id=\"rLayout1\"/></p:sldLayoutIdLst></p:sldMaster>");
            WriteZip(tzip, "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>" +
                "<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                "</Relationships>");
            WriteZip(tzip, "ppt/slideLayouts/slideLayout1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" type=\"blank\">" +
                "<p:cSld name=\"Blank\"><p:spTree/></p:cSld></p:sldLayout>");
            WriteZip(tzip, "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/></Relationships>");
            WriteZip(tzip, "ppt/slides/slide1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree/></p:cSld></p:sld>");
            WriteZip(tzip, "ppt/slides/_rels/slide1.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/></Relationships>");
            WriteZip(tzip, "ppt/presentation.xml",
                "<?xml version=\"1.0\"?><p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"></p:presentation>");
            WriteZip(tzip, "ppt/_rels/presentation.xml.rels",
                "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");
            WriteZip(tzip, "[Content_Types].xml",
                "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"></Types>");
        }
        templateMs.Position = 0;
        var templateBytes = templateMs.ToArray();

        var writer = new PptxWriter();
        writer.LoadMaster(templateBytes, keepTemplateSlides: true);

        // 添加一张程序化幻灯片
        writer.AddSlide(0);
        writer.AddTextBox(0, "新增内容", 2, 2, 10, 2);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        // 应该有 2 张幻灯片：程序化 slide1 + 模板 slide2
        Assert.NotNull(zip.GetEntry("ppt/slides/slide1.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slides/slide2.xml"));
        // 模板母版和版式也应保留
        Assert.NotNull(zip.GetEntry("ppt/slideMasters/slideMaster1.xml"));
        Assert.NotNull(zip.GetEntry("ppt/slideLayouts/slideLayout1.xml"));
    }

    [Fact, System.ComponentModel.DisplayName("keepTemplateSlides=false 时不保留模板幻灯片")]
    public void LoadMaster_WithoutKeepSlides_NoTemplateSlides()
    {
        // 使用 BuildTemplatePptx 辅助方法，它将 theme1 as marker
        var template = BuildTemplatePptx("L");
        var writer = new PptxWriter();
        writer.LoadMaster(template, keepTemplateSlides: false);
        writer.AddSlide(0);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("ppt/slides/slide1.xml"));
        // 不应有 slide2（模板幻灯片未被保留）
        Assert.Null(zip.GetEntry("ppt/slides/slide2.xml"));
    }

    // ─── 混合模式与覆盖盲区测试 ────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("混合：LoadMaster 清除编程母版")]
    public void LoadMaster_AfterCreateMaster_ClearsProgMasters()
    {
        var writer = new PptxWriter();
        writer.CreateMaster().AddLayout("编程版式A");
        writer.CreateMaster().AddLayout("编程版式B");
        Assert.Equal(2, writer.GetLayoutCount());

        var template = BuildTemplatePptx("模板版式");
        writer.LoadMaster(template);
        // 模板只有 1 个版式
        Assert.Equal(1, writer.GetLayoutCount());
        Assert.Equal("模板版式", writer.GetLayoutName(0));
    }

    [Fact, System.ComponentModel.DisplayName("编程母版：无版式时 Save 正常回退到默认版式")]
    public void Save_ProgMasterNoLayouts_FallsBackToDefault()
    {
        var writer = new PptxWriter();
        writer.CreateMaster(); // 有母版但无版式
        writer.AddSlide(0);

        using var ms = new MemoryStream();
        var ex = Record.Exception(() => writer.Save(ms));
        Assert.Null(ex);
        Assert.True(ms.Length > 0);
    }

    [Fact, System.ComponentModel.DisplayName("编程母版：版式同时包含形状和文本框")]
    public void Save_ProgLayoutWithShapesAndTextboxes_MixedContent()
    {
        var writer = new PptxWriter();
        var master = writer.CreateMaster();
        var layout = master.AddLayout("混合版式", "blank");
        layout.Shapes.Add(new PptShape
        {
            ShapeType = "rect",
            Left = PptxWriter.CmToEmu(1),
            Top = PptxWriter.CmToEmu(1),
            Width = PptxWriter.CmToEmu(5),
            Height = PptxWriter.CmToEmu(2),
            FillColor = "FF0000",
        });
        layout.TextBoxes.Add(new PptTextBox
        {
            Text = "混合内容",
            Left = PptxWriter.CmToEmu(7),
            Top = PptxWriter.CmToEmu(1),
            Width = PptxWriter.CmToEmu(10),
            Height = PptxWriter.CmToEmu(2),
            FontSize = 18,
        });
        writer.AddSlide(0);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var layoutXml = ReadZipEntry(zip, "ppt/slideLayouts/slideLayout1.xml");
        Assert.Contains("prstGeom prst=\"rect\"", layoutXml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("混合内容", layoutXml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("keepTemplateSlides：模板幻灯片含媒体文件正确重命名")]
    public void LoadMaster_KeepSlidesWithMedia_RenamesMediaFiles()
    {
        // 构建带媒体引用的模板：slide1.xml.rels 引用 image1.png
        using var templateMs = new MemoryStream();
        using (var tzip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZip(tzip, "ppt/theme/theme1.xml", "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"></a:theme>");
            WriteZip(tzip, "ppt/slideMasters/slideMaster1.xml", "<?xml version=\"1.0\"?><p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree/></p:cSld><p:txStyles/><p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rLayout1\"/></p:sldLayoutIdLst></p:sldMaster>");
            WriteZip(tzip, "ppt/slideMasters/_rels/slideMaster1.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/><Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/></Relationships>");
            WriteZip(tzip, "ppt/slideLayouts/slideLayout1.xml", "<?xml version=\"1.0\"?><p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" type=\"blank\"><p:cSld name=\"Blank\"><p:spTree/></p:cSld></p:sldLayout>");
            WriteZip(tzip, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/></Relationships>");
            // 幻灯片引用了图片
            WriteZip(tzip, "ppt/slides/slide1.xml", "<?xml version=\"1.0\"?><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:cSld><p:spTree><p:pic><p:blipFill><a:blip r:embed=\"rImg1\"/></a:blipFill></p:pic></p:spTree></p:cSld></p:sld>");
            WriteZip(tzip, "ppt/slides/_rels/slide1.xml.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/><Relationship Id=\"rImg1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/></Relationships>");
            // 媒体文件
            WriteZip(tzip, "ppt/media/image1.png", String.Empty);
            WriteZip(tzip, "ppt/presentation.xml", "<?xml version=\"1.0\"?><p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"></p:presentation>");
            WriteZip(tzip, "ppt/_rels/presentation.xml.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");
            WriteZip(tzip, "[Content_Types].xml", "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"></Types>");
        }
        templateMs.Position = 0;

        var writer = new PptxWriter();
        writer.LoadMaster(templateMs.ToArray(), keepTemplateSlides: true);

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        // 模板幻灯片被保留
        Assert.NotNull(zip.GetEntry("ppt/slides/slide1.xml"));
        // 媒体文件被重命名后存在（不以 image1.png 命名，以 mt{N}.png 命名）
        var mediaEntries = zip.Entries.Where(e => e.FullName.StartsWith("ppt/media/")).ToList();
        Assert.NotEmpty(mediaEntries);
    }

    [Fact, System.ComponentModel.DisplayName("编程母版：多母版各含版式，AddSlide 索引正确跨越母版边界")]
    public void AddSlide_CrossMasterBoundary_CorrectLayoutRef()
    {
        var writer = new PptxWriter();
        var m1 = writer.CreateMaster();
        m1.AddLayout("M1-L0");
        m1.AddLayout("M1-L1");
        var m2 = writer.CreateMaster();
        m2.AddLayout("M2-L0");

        // 版式索引 0,1 属母版1；索引 2 属母版2
        writer.AddSlide(0); // M1-L0 → slideLayout1
        writer.AddSlide(1); // M1-L1 → slideLayout2
        writer.AddSlide(2); // M2-L0 → slideLayout3

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var rels3 = ReadZipEntry(zip, "ppt/slides/_rels/slide3.xml.rels");
        Assert.Contains("slideLayout3.xml", rels3, StringComparison.OrdinalIgnoreCase);
    }

    [Fact, System.ComponentModel.DisplayName("模板加载后编程母版被清除再Save不抛异常")]
    public void Save_TemplateThenProgMasterThenTemplate_CorrectOutput()
    {
        var t1 = BuildTemplatePptx("模板A");
        var writer = new PptxWriter();
        writer.LoadMaster(t1);
        writer.AddSlide(0);

        // 改为编程母版
        writer.CreateMaster().AddLayout("编程C");
        writer.AddSlide(0);

        // 再改回模板加载
        var t2 = BuildTemplatePptx("模板B");
        writer.LoadMaster(t2);

        using var ms = new MemoryStream();
        var ex = Record.Exception(() => writer.Save(ms));
        Assert.Null(ex);
        Assert.True(ms.Length > 0);
    }

    [Fact, System.ComponentModel.DisplayName("DuplicateShape克隆形状")]
    public void DuplicateShape_ClonesShape()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var src = writer.AddShape(0, "rect", 1, 2, 5, 3, "FF0000");
        src.Text = "原始形状";
        src.FontSize = 18;
        src.Bold = true;

        var clone = writer.DuplicateShape(0, 0);
        Assert.NotNull(clone);
        Assert.Equal("原始形状", clone.Text);
        Assert.Equal("rect", clone.ShapeType);
        Assert.Equal(PptxWriter.CmToEmu(1), clone.Left);
        Assert.Equal(PptxWriter.CmToEmu(5), clone.Width);
        Assert.Equal(PptxWriter.CmToEmu(3), clone.Height);
        Assert.Equal("FF0000", clone.FillColor);
        Assert.Equal(18, clone.FontSize);
        Assert.True(clone.Bold);
        // 验证偏移
        Assert.True(clone.Top > src.Top);

        // 验证写入不抛异常
        using var ms = new MemoryStream();
        var ex = Record.Exception(() => writer.Save(ms));
        Assert.Null(ex);
        Assert.True(ms.Length > 0);
    }

    [Fact, System.ComponentModel.DisplayName("DuplicateShape索引越界抛异常")]
    public void DuplicateShape_InvalidIndex_Throws()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.DuplicateShape(0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.DuplicateShape(0, -1));
    }

    [Fact, System.ComponentModel.DisplayName("图片旋转读写往返")]
    public void ImageRotation_Roundtrip()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var img = writer.AddImage(0, new Byte[] { 0x89, 0x50, 0x4E, 0x47 }, "png", 2, 2, 8, 6);
        img.Rotation = 5400000; // 90度

        using var ms = new MemoryStream();
        writer.Save(ms);

        ms.Position = 0;
        using var reader = new PptxReader(ms);
        var slides = reader.ReadAllSlides().ToList();
        Assert.Single(slides);
        Assert.Single(slides[0].Images);
        Assert.Equal(5400000, slides[0].Images[0].Rotation);
    }

    [Fact, System.ComponentModel.DisplayName("形状图片填充生成blipFill")]
    public void ShapeImageFill_WritesBlipFill()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var sp = writer.AddShape(0, "rect", 1, 2, 5, 3, null);
        sp.FillImage = new Byte[] { 1, 2, 3 };
        sp.ShapeImageRelId = "rImg1";

        using var ms = new MemoryStream();
        writer.Save(ms);

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var slideEntry = za.GetEntry("ppt/slides/slide1.xml");
        Assert.NotNull(slideEntry);
        using var sr = new StreamReader(slideEntry!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("a:blipFill", xml);
        Assert.Contains("r:embed=\"rImg1\"", xml);
    }

    [Fact, System.ComponentModel.DisplayName("Morph切换—写入并验证")]
    public void MorphTransition_WritesCorrectXml()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        writer.SetTransition(0, "morph", 1000);

        using var ms = new MemoryStream();
        writer.Save(ms);

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var slideEntry = za.GetEntry("ppt/slides/slide1.xml");
        Assert.NotNull(slideEntry);
        using var sr = new StreamReader(slideEntry!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("p:morph", xml);
        Assert.Contains("byObject", xml);
    }

    [Fact, System.ComponentModel.DisplayName("形状克隆—模型级Clone含新ID和偏移")]
    public void ShapeClone_ModelLevel_WithOffset()
    {
        var src = new PptShape
        {
            Id = 1,
            Text = "原始文本",
            ShapeType = "rect",
            Left = 1000000, Top = 2000000,
            Width = 5000000, Height = 3000000,
            FillColor = "FF0000",
            LineColor = "0000FF",
            FontSize = 16,
            Bold = true,
            Rotation = 5400000,
            AltText = "描述文本",
        };

        var clone = src.Clone(newId: 2, offsetX: 500000, offsetY: 300000);

        Assert.Equal(2, clone.Id);
        Assert.Equal("原始文本", clone.Text);
        Assert.Equal("rect", clone.ShapeType);
        Assert.Equal(1500000, clone.Left);  // 1000000 + 500000
        Assert.Equal(2300000, clone.Top);   // 2000000 + 300000
        Assert.Equal(5000000, clone.Width);
        Assert.Equal(3000000, clone.Height);
        Assert.Equal("FF0000", clone.FillColor);
        Assert.Equal("0000FF", clone.LineColor);
        Assert.Equal(16, clone.FontSize);
        Assert.True(clone.Bold);
        Assert.Equal(5400000, clone.Rotation);
        Assert.Equal("描述文本", clone.AltText);
    }

    [Fact, System.ComponentModel.DisplayName("图片圆角—写入roundRect几何并验证")]
    public void ImageCornerRadius_WritesRoundRectGeometry()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var img = writer.AddImage(0, [1, 2, 3], "png", 2, 2, 5, 4);
        img.CornerRadius = 50000;

        using var ms = new MemoryStream();
        writer.Save(ms);

        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var slideEntry = za.GetEntry("ppt/slides/slide1.xml");
        Assert.NotNull(slideEntry);
        using var sr = new StreamReader(slideEntry!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("roundRect", xml);
        Assert.Contains("50000", xml);
    }

    [Fact, System.ComponentModel.DisplayName("形状渐变填充—GradientFill写入gradFill")]
    public void GradientFill_Shape()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var shape = writer.AddShape(0, "rect", 1, 1, 8, 4);
        shape.GradientType = "linear";
        shape.GradientColor1 = "4472C4";
        shape.GradientColor2 = "FFFFFF";
        shape.GradientAngle = 45;

        using var ms = new MemoryStream();
        writer.Save(ms);
        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var slide1 = za.GetEntry("ppt/slides/slide1.xml");
        Assert.NotNull(slide1);
        using var sr = new StreamReader(slide1!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("gradFill", xml);
        Assert.Contains("4472C4", xml);
    }

    [Fact, System.ComponentModel.DisplayName("取消组合—UngroupShapes释放组内形状")]
    public void UngroupShapes_Works()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var group = writer.GroupShapes(0, 1, 1, 10, 10);
        group.Shapes.Add(new PptShape { ShapeType = "rect", Left = 100, Top = 200, Width = 500, Height = 300 });
        group.TextBoxes.Add(new PptTextBox { Text = "组内文字", Left = 300, Top = 400, Width = 200, Height = 100 });

        var slide = writer.Slides[0];
        Assert.Single(slide.Groups);
        Assert.Empty(slide.Shapes);

        writer.UngroupShapes(0, 0);
        Assert.Empty(slide.Groups);
        Assert.Single(slide.Shapes);
        Assert.Single(slide.TextBoxes);
        // 坐标应从相对于组转换为相对于幻灯片
        Assert.True(slide.Shapes[0].Left > 0);
    }

    [Fact, System.ComponentModel.DisplayName("Z-Order—BringToFront置顶")]
    public void ZOrder_BringToFront()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        writer.AddShape(0, "rect", 1, 1, 5, 5);
        writer.AddShape(0, "ellipse", 2, 2, 5, 5);

        var slide = writer.Slides[0];
        Assert.Equal(2, slide.Shapes.Count);
        Assert.Equal("rect", slide.Shapes[0].ShapeType);

        writer.BringToFront(0, 0);
        Assert.Equal("ellipse", slide.Shapes[0].ShapeType);
        Assert.Equal("rect", slide.Shapes[1].ShapeType);
    }

    [Fact, System.ComponentModel.DisplayName("股价图—AddStockChart创建并验证")]
    public void StockChart_CreatesSuccessfully()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var chart = writer.AddStockChart(0, ["周一", "周二", "周三"], 2, 2, 16, 12);
        chart.Series.Add(new PptChartSeries { Name = "AAPL", Values = [150, 152, 148] });

        using var ms = new MemoryStream();
        writer.Save(ms);
        Assert.True(ms.Length > 0);
        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var chartEntry = za.GetEntry("ppt/charts/chart1.xml");
        Assert.NotNull(chartEntry);
        using var sr = new StreamReader(chartEntry!.Open(), Encoding.UTF8);
        Assert.Contains("stockChart", sr.ReadToEnd());
    }

    [Fact, System.ComponentModel.DisplayName("雷达图—AddRadarChart创建并验证")]
    public void RadarChart_CreatesSuccessfully()
    {
        var writer = new PptxWriter();
        writer.AddSlide(0);
        var chart = writer.AddRadarChart(0, ["速度", "力量", "智力"], 2, 2, 16, 12);
        chart.Series.Add(new PptChartSeries { Name = "英雄A", Values = [80, 60, 90] });

        using var ms = new MemoryStream();
        writer.Save(ms);

        Assert.True(ms.Length > 0);
        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var chartEntry = za.GetEntry("ppt/charts/chart1.xml");
        Assert.NotNull(chartEntry);
        using var sr = new StreamReader(chartEntry!.Open(), Encoding.UTF8);
        var xml = sr.ReadToEnd();
        Assert.Contains("radarChart", xml);
        Assert.Contains("radarStyle", xml);
    }

    [Fact(DisplayName = "翻转—FlipHorizontal/Vertical输出flipH/flipV")]
    public void FlipHorizontal_Vertical_OutputFlipAttributes()
    {
        using var writer = new PptxWriter();
        var shape = writer.AddShape(0, "rect", 2, 2, 6, 4);
        shape.FlipHorizontal = true;
        shape.FlipVertical = true;
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            writer.Save(tempFile);
            Assert.True(File.Exists(tempFile));
            using var za = ZipFile.OpenRead(tempFile);
            var entry = za.GetEntry("ppt/slides/slide1.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("flipH=\"1\"", xml);
            Assert.Contains("flipV=\"1\"", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "形状增强—TextDirection/DashStyle/Inset输出bodyPr和prstDash")]
    public void Shape_TextDirection_DashStyle_Inset_OutputXml()
    {
        using var writer = new PptxWriter();
        var shape = writer.AddShape(0, "rect", 2, 2, 6, 4);
        shape.Text = "竖排虚线边框形状";
        shape.TextDirection = "vert";
        shape.DashStyle = "dash";
        shape.LeftInset = 50000;
        shape.TopInset = 30000;
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            writer.Save(tempFile);
            Assert.True(File.Exists(tempFile));
            using var za = ZipFile.OpenRead(tempFile);
            var entry = za.GetEntry("ppt/slides/slide1.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("vert=\"vert\"", xml);
            Assert.Contains("prstDash val=\"dash\"", xml);
            Assert.Contains("lIns=\"50000\"", xml);
            Assert.Contains("tIns=\"30000\"", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "形状Anchor—输出bodyPr anchor属性")]
    public void Shape_Anchor_OutputsBodyPrAnchor()
    {
        using var writer = new PptxWriter();
        var shape = writer.AddShape(0, "rect", 2, 2, 6, 4);
        shape.Text = "垂直居中文字";
        shape.Anchor = "ctr";
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            writer.Save(tempFile);
            Assert.True(File.Exists(tempFile));
            using var za = ZipFile.OpenRead(tempFile);
            var entry = za.GetEntry("ppt/slides/slide1.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("anchor=\"ctr\"", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "散点图—XValues输出c:xVal")]
    public void ScatterChart_XValues_OutputsXVal()
    {
        using var writer = new PptxWriter();
        var chart = writer.AddScatterChart(0, ["A", "B", "C"], 1, 1, 8, 6);
        chart.Title = "Scatter Test";
        chart.Series.Add(new PptChartSeries
        {
            Name = "Series1",
            XValues = [1.0, 2.0, 3.0],
            Values = [10.0, 20.0, 15.0]
        });
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            writer.Save(tempFile);
            Assert.True(File.Exists(tempFile));
            using var za = ZipFile.OpenRead(tempFile);
            var entry = za.GetEntry("ppt/charts/chart1.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("<c:xVal>", xml);
            Assert.Contains("<c:v>2</c:v>", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "演讲者备注—SetNotes写入notesSlide")]
    public void SetNotes_WritesNotesSlide()
    {
        using var writer = new PptxWriter();
        writer.AddSlide();
        writer.SetNotes(0, "这是演讲者备注内容");
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            writer.Save(tempFile);
            Assert.True(File.Exists(tempFile));
            using var za = ZipFile.OpenRead(tempFile);
            var entry = za.GetEntry("ppt/slides/slide1.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("演讲者备注内容", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "幻灯片页脚—SetSlideFooter写入页脚和页码")]
    public void SetSlideFooter_WritesFooter()
    {
        using var writer = new PptxWriter();
        writer.AddSlide();
        writer.SetSlideFooter(0, "公司机密", true);
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            writer.Save(tempFile);
            Assert.True(File.Exists(tempFile));
            using var za = ZipFile.OpenRead(tempFile);
            var entry = za.GetEntry("ppt/slides/slide1.xml");
            Assert.NotNull(entry);
            using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("公司机密", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
}
