using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>JSON 补丁测试：解析、按锚点批量修改、失败即整体不保存</summary>
public class PatchApiTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #region 工具
    private static Byte[] BuildDocx(String bodyInnerXml)
    {
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>");
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static Document Read(Document doc)
    {
        using var ms = new MemoryStream(Save(doc));
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static Document ReadXml(String bodyInnerXml)
    {
        using var ms = new MemoryStream(BuildDocx(bodyInnerXml));
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static Byte[] Save(Document doc)
    {
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.Save(ms, doc);
        return ms.ToArray();
    }

    private static Document ReadBack(Byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static String BodyText(Document doc)
    {
        var sb = new StringBuilder();
        foreach (var el in doc.Elements)
        {
            if (el.Paragraph == null) continue;
            foreach (var r in el.Paragraph.Runs) sb.Append(r.Text);
            sb.Append('|');
        }
        return sb.ToString();
    }

    /// <summary>构建含 标题 + 两段正文 + 表格 的样张</summary>
    private static Document BuildSample()
    {
        var doc = new Document();
        doc.AddHeading("第一章", 1);
        doc.AddParagraph("这是正文一。");
        doc.AddParagraph("这是正文二。");
        doc.AddTable([["姓名", "部门"], ["张三", "研发"]]);
        return doc;
    }
    #endregion

    #region 解析
    [Fact(DisplayName = "补丁-解析 ops 数组与字段")]
    public void ParseJson_ReadsOps()
    {
        const String json = """
        { "ops": [
            { "op": "set_text", "anchor": "p#1", "text": "新内容" },
            { "op": "replace_text", "find": "旧", "replace": "新", "regex": false },
            { "op": "set_cell_text", "anchor": "tbl#3", "row": 1, "column": 1, "text": "市场部" }
        ]}
        """;

        var patch = DocxPatch.ParseJson(json);
        Assert.Equal(3, patch.Ops.Count);
        Assert.Equal("set_text", patch.Ops[0].Op);
        Assert.Equal("p#1", patch.Ops[0].Anchor);
        Assert.Equal("新内容", patch.Ops[0].Text);
        Assert.False(patch.Ops[1].Regex);
        Assert.Equal(1, patch.Ops[2].Row);
        Assert.Equal(1, patch.Ops[2].Column);
    }

    [Fact(DisplayName = "补丁-非法 JSON 抛 FormatException")]
    public void ParseJson_Invalid_Throws()
    {
        Assert.Throws<FormatException>(() => DocxPatch.ParseJson("{oops}"));
        Assert.Empty(DocxPatch.ParseJson("{}").Ops);
        Assert.Throws<FormatException>(() => DocxPatch.ParseJson("{{\"foo\":1}}"));
        Assert.Throws<ArgumentNullException>(() => DocxPatch.ParseJson(null!));
    }
    #endregion

    #region 执行
    [Fact(DisplayName = "补丁-set_text 按锚点改写段落")]
    public void Apply_SetText()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"set_text","anchor":"p#1","text":"改后的正文一"}]}""");
        var result = DocxPatch.Apply(doc, patch);

        Assert.True(result.Success, String.Join(";", result.Failed));
        Assert.Equal("第一章|改后的正文一|这是正文二。|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "补丁-set_style 设置标题样式")]
    public void Apply_SetStyle()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"set_style","anchor":"p#1","style":"Heading2"}]}""");
        Assert.True(DocxPatch.Apply(doc, patch).Success);

        var back = ReadBack(Save(doc));
        Assert.Equal(ParagraphStyle.Heading2, back.AllParagraphs()[1].Style);
    }

    [Fact(DisplayName = "补丁-insert_paragraph 在锚点后插入")]
    public void Apply_InsertParagraph()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"insert_paragraph","anchor":"p#1","position":"after","text":"插入的新段"}]}""");
        Assert.True(DocxPatch.Apply(doc, patch).Success);

        Assert.Equal("第一章|这是正文一。|插入的新段|这是正文二。|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "补丁-delete 删除锚点元素")]
    public void Apply_Delete()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"delete","anchor":"p#2"}]}""");
        Assert.True(DocxPatch.Apply(doc, patch).Success);

        Assert.Equal("第一章|这是正文一。|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "补丁-set_cell_text 改写单元格")]
    public void Apply_SetCellText()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"set_cell_text","anchor":"tbl#3","row":1,"column":1,"text":"市场部"}]}""");
        Assert.True(DocxPatch.Apply(doc, patch).Success);

        var back = ReadBack(Save(doc));
        var table = Assert.Single(back.AllTables());
        Assert.Equal("市场部", table.GetCellText(1, 1));
    }

    [Fact(DisplayName = "补丁-insert_table 按 cells 矩阵插入表格")]
    public void Apply_InsertTable()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"insert_table","anchor":"p#2","position":"after","cells":[["A","B"],["1","2"]]}]}""");
        Assert.True(DocxPatch.Apply(doc, patch).Success);

        var back = ReadBack(Save(doc));
        Assert.Equal(2, back.AllTables().Count);
        Assert.Contains(back.AllTables(), t => t.GetCellText(1, 0) == "1");
    }

    [Fact(DisplayName = "补丁-add_comment 给段落加批注")]
    public void Apply_AddComment()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"add_comment","anchor":"p#1","text":"这里需要数据","author":"AI"}]}""");
        Assert.True(DocxPatch.Apply(doc, patch).Success);

        var back = ReadBack(Save(doc));
        var comment = Assert.Single(back.Comments);
        Assert.Equal("这里需要数据", comment.Text);
        Assert.Equal("AI", comment.Author);
    }

    [Fact(DisplayName = "补丁-replace_text 正则替换")]
    public void Apply_ReplaceRegex()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"replace_text","find":"正文(一|二)","replace":"段落$1","regex":true}]}""");
        Assert.True(DocxPatch.Apply(doc, patch).Success);

        Assert.Equal("第一章|这是段落一。|这是段落二。|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "补丁-多操作按顺序执行")]
    public void Apply_MultipleOps()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        const String json = """
        { "ops": [
            { "op": "set_text", "anchor": "p#1", "text": "第一段改" },
            { "op": "insert_paragraph", "anchor": "p#1", "position": "after", "text": "新增段" },
            { "op": "set_style", "anchor": "p#0", "style": "Heading2" }
        ]}
        """;
        var result = DocxPatch.Apply(doc, DocxPatch.ParseJson(json));

        Assert.True(result.Success, String.Join(";", result.Failed));
        Assert.Equal(3, result.Applied.Count);

        var back = ReadBack(Save(doc));
        Assert.Equal("第一章|第一段改|新增段|这是正文二。|", BodyText(back));
        Assert.Equal(ParagraphStyle.Heading2, back.AllParagraphs()[0].Style);
    }
    #endregion

    #region 失败语义
    [Fact(DisplayName = "补丁-锚点不存在记录失败且不抛异常")]
    public void Apply_BadAnchor_RecordsFailure()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var patch = DocxPatch.ParseJson("""{"ops":[{"op":"set_text","anchor":"p#99","text":"x"}]}""");
        var result = DocxPatch.Apply(doc, patch);

        Assert.False(result.Success);
        var failed = Assert.Single(result.Failed);
        Assert.Contains("set_text", failed);
    }

    [Fact(DisplayName = "补丁-不支持的操作记录失败")]
    public void Apply_UnknownOp_RecordsFailure()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var result = DocxPatch.Apply(doc, DocxPatch.ParseJson("""{"ops":[{"op":"fly_to_moon"}]}"""));
        Assert.False(result.Success);
        Assert.Contains("不支持的操作", result.Failed[0]);
    }

    [Fact(DisplayName = "补丁-replace_text 未匹配视为失败")]
    public void Apply_NoMatch_RecordsFailure()
    {
        var doc = BuildSample();
        doc.BeginEdit();

        var result = DocxPatch.Apply(doc, DocxPatch.ParseJson("""{"ops":[{"op":"replace_text","find":"不存在的词","replace":"x"}]}"""));
        Assert.False(result.Success);
        Assert.Contains("未匹配", result.Failed[0]);
    }

    [Fact(DisplayName = "补丁-失败时调用方不保存，原文档不受影响")]
    public void Apply_PartialFailure_DocumentNotSaved()
    {
        var original = BuildSample();
        var originalBytes = Save(original);
        var beforeText = BodyText(ReadBack(originalBytes));

        // 先成功的操作改内存模型，第二条失败
        var doc = ReadBack(originalBytes);
        doc.BeginEdit();
        var result = DocxPatch.Apply(doc, DocxPatch.ParseJson(
            """{"ops":[{"op":"set_text","anchor":"p#1","text":"改过"},{"op":"delete","anchor":"p#99"}]}"""));

        Assert.False(result.Success);
        // 调用方按契约不保存 → 重读原字节仍是原文
        Assert.Equal(beforeText, BodyText(ReadBack(originalBytes)));
    }

    [Fact(DisplayName = "补丁-部分失败不回滚（内存模型已被改动，必须丢弃重读）")]
    public void Apply_PartialFailure_DoesNotRollback()
    {
        var doc = BuildSample();
        doc.BeginEdit();
        var result = DocxPatch.Apply(doc, DocxPatch.ParseJson(
            """{"ops":[{"op":"set_text","anchor":"p#1","text":"已改"},{"op":"delete","anchor":"p#99"}]}"""));

        Assert.False(result.Success);
        Assert.Single(result.Applied);
        // 契约：无回滚——成功的那条仍留在内存模型上，调用方须丢弃该 Document 重新读取
        Assert.Equal("已改", doc.Elements[1].Paragraph!.Runs[0].Text);
    }

    [Fact(DisplayName = "补丁-结构性操作会使后续锚点偏移（须排在最后）")]
    public void Apply_InsertShiftsFollowingAnchors()
    {
        var doc = BuildSample();
        doc.BeginEdit();
        // p#0 前插入新段 → 原 p#1 变成 p#2，因此这里用 p#1 会命中“原 p#0”
        var result = DocxPatch.Apply(doc, DocxPatch.ParseJson(
            """{"ops":[{"op":"insert_paragraph","anchor":"p#0","position":"before","text":"插入"},{"op":"set_text","anchor":"p#1","text":"改"}]}"""));

        Assert.True(result.Success);
        Assert.Equal("插入", doc.Elements[0].Paragraph!.Runs[0].Text);
        // 命中偏移后的元素（原 p#0），而非原始的 p#1——这正是要求结构性操作排最后的原因
        Assert.Equal("改", doc.Elements[1].Paragraph!.Runs[0].Text);
    }
    #endregion
}
