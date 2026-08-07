using System.ComponentModel;
using System.Text;
using NewLife.Office;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Ole2;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using NewLife.Office.Xps;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>BiffReader xls BIFF8 读取器单元测试</summary>
public class BiffReaderTests
{
    // ─── BIFF8 字节构建辅助 ─────────────────────────────────────────────────

    /// <summary>通过 CfbDocument 将 BIFF8 流打包为 xls OLE2 字节</summary>
    private static Byte[] BuildXls(Byte[] workbookStream)
    {
        var doc = new CfbDocument();
        doc.Root.AddStream("Workbook", workbookStream);
        return doc.ToBytes();
    }

    /// <summary>构建 BIFF8 记录：type(2) + len(2) + data</summary>
    private static Byte[] Rec(UInt16 type, Byte[] data)
    {
        var buf = new Byte[4 + data.Length];
        buf[0] = (Byte)(type & 0xFF);
        buf[1] = (Byte)(type >> 8);
        buf[2] = (Byte)(data.Length & 0xFF);
        buf[3] = (Byte)(data.Length >> 8);
        Array.Copy(data, 0, buf, 4, data.Length);
        return buf;
    }

    private static Byte[] Concat(params Byte[][] parts)
    {
        var total = 0;
        foreach (var p in parts) total += p.Length;
        var result = new Byte[total];
        var pos = 0;
        foreach (var p in parts) { Array.Copy(p, 0, result, pos, p.Length); pos += p.Length; }
        return result;
    }

    private static Byte[] LE2(UInt16 v) => [(Byte)(v & 0xFF), (Byte)(v >> 8)];
    private static Byte[] LE4(UInt32 v) => [(Byte)(v & 0xFF), (Byte)((v >> 8) & 0xFF), (Byte)((v >> 16) & 0xFF), (Byte)(v >> 24)];
    private static Byte[] LE8(Double v) => BitConverter.GetBytes(v);

    // 构建 BIFF8 BOF 记录 (type=0x0809，vers=0x0600, dt=0x0005 Globals 或 0x0010 Worksheet)
    private static Byte[] Bof(UInt16 dt = 0x0005) =>
        Rec(0x0809, Concat(LE2(0x0600), LE2(dt), LE4(0), LE4(0)));

    private static Byte[] Eof() => Rec(0x000A, Array.Empty<Byte>());

    // SST 记录 (0x00FC)
    private static Byte[] SstRec(String[] strings)
    {
        var body = new List<Byte>();
        body.AddRange(LE4((UInt32)strings.Length));  // total refs（随意填 count）
        body.AddRange(LE4((UInt32)strings.Length));  // unique count

        foreach (var s in strings)
        {
            var chars = Encoding.Unicode.GetBytes(s);
            body.AddRange(LE2((UInt16)s.Length));  // cch
            body.Add(0x01);                         // fHighByte=1 (UTF-16LE)
            body.AddRange(chars);
        }

        return Rec(0x00FC, body.ToArray());
    }

    // BOUNDSHEET 记录 (0x0085) — lbPlyPos 为工作表 BOF 在工作簿流中的字节偏移
    private static Byte[] BoundSheet(UInt32 bofOffset, String name)
    {
        var body = new List<Byte>();
        body.AddRange(LE4(bofOffset));               // lbPlyPos (4 bytes)
        body.Add(0x00);                              // grbit low  (hidden=0)
        body.Add(0x00);                              // grbit high (type=worksheet)
        body.Add((Byte)name.Length);                 // cch
        body.Add(0x00);                              // fHighByte=0 (Latin-1)
        body.AddRange(Encoding.ASCII.GetBytes(name));
        return Rec(0x0085, body.ToArray());
    }

    // 计算 BoundSheet 记录大小（Latin-1 名称）
    private static Int32 BoundSheetSize(String name) =>
        4 + 4 + 1 + 1 + 1 + 1 + name.Length; // header + lbPlyPos + grbit*2 + cch + fHighByte + name

    // LABELSST 记录 (0x00FD): row(2)+col(2)+xf(2)+sstIndex(4) = 10 bytes data
    private static Byte[] LabelSst(UInt16 row, UInt16 col, UInt32 sstIdx) =>
        Rec(0x00FD, Concat(LE2(row), LE2(col), LE2(0), LE4(sstIdx)));

    // NUMBER 记录 (0x0203): row+col+xf+double
    private static Byte[] Number(UInt16 row, UInt16 col, Double value) =>
        Rec(0x0203, Concat(LE2(row), LE2(col), LE2(0), LE8(value)));

    // BOOLERR 记录 (0x0205): row(2)+col(2)+xf(2)+boolOrErr(1)+isError(1) = 8 bytes data
    private static Byte[] BoolErr(UInt16 row, UInt16 col, Boolean val) =>
        Rec(0x0205, Concat(LE2(row), LE2(col), LE2(0), new Byte[] { (Byte)(val ? 1 : 0), 0x00 }));

    /// <summary>构建一个最小单工作表 xls</summary>
    private static Byte[] BuildMinimalXls(String sheetName, Byte[] sheetBody, String[] sst = null)
    {
        // Globals 段
        var globalsParts = new List<Byte[]> { Bof(0x0005) };
        if (sst != null) globalsParts.Add(SstRec(sst));

        var eof = Eof();
        // 工作表 BOF 偏移 = 当前 globals 大小 + BoundSheet 记录大小 + EOF 大小
        var wsOffset = (UInt32)(globalsParts.Sum(p => p.Length) + BoundSheetSize(sheetName) + eof.Length);

        globalsParts.Add(BoundSheet(wsOffset, sheetName));
        globalsParts.Add(eof);

        // 工作表段
        var sheetFull = Concat(Bof(0x0010), sheetBody, Eof());

        return Concat(globalsParts.Concat(new[] { sheetFull }).ToArray());
    }

    // ─── 工作表名称测试 ────────────────────────────────────────────────────

    [Fact, DisplayName("读取工作表名称列表")]
    public void SheetNames_AreReadCorrectly()
    {
        var stream = BuildMinimalXls("MySheet", []);
        var xls = BuildXls(stream);
        using var reader = new BiffReader(new MemoryStream(xls));
        Assert.Single(reader.SheetNames);
        Assert.Equal("MySheet", reader.SheetNames[0]);
    }

    // ─── 字符串单元格（LABELSST）─────────────────────────────────────────

    [Fact, DisplayName("读取 LABELSST 字符串单元格")]
    public void ReadSheet_StringCell_FromSst()
    {
        var sst = new[] { "Hello", "World" };
        // 第0行第0列 = SST[0]="Hello", 第0行第1列 = SST[1]="World"
        var body = Concat(LabelSst(0, 0, 0), LabelSst(0, 1, 1));
        var stream = BuildMinimalXls("Sheet1", body, sst);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        Assert.Equal("Hello", rows[0][0]);
        Assert.Equal("World", rows[0][1]);
    }

    // ─── 数值单元格（NUMBER）─────────────────────────────────────────────

    [Fact, DisplayName("读取 NUMBER 数值单元格")]
    public void ReadSheet_NumberCell()
    {
        var body = Number(0, 0, 3.14);
        var stream = BuildMinimalXls("Sheet1", body);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        var val = Assert.IsType<Double>(rows[0][0]);
        Assert.Equal(3.14, val, 6);
    }

    // ─── 布尔单元格（BOOLERR）────────────────────────────────────────────

    [Fact, DisplayName("读取 BOOLERR 布尔单元格（true）")]
    public void ReadSheet_BoolCell_True()
    {
        var body = BoolErr(0, 0, true);
        var stream = BuildMinimalXls("Sheet1", body);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        var bVal = Assert.IsType<Boolean>(rows[0][0]);
        Assert.True(bVal);
    }

    [Fact, DisplayName("读取 BOOLERR 布尔单元格（false）")]
    public void ReadSheet_BoolCell_False()
    {
        var body = BoolErr(0, 0, false);
        var stream = BuildMinimalXls("Sheet1", body);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows = reader.ReadSheet().ToList();
        Assert.Single(rows);
        var bVal = Assert.IsType<Boolean>(rows[0][0]);
        Assert.False(bVal);
    }

    // ─── 多行数据 ──────────────────────────────────────────────────────────

    [Fact, DisplayName("读取多行字符串数据（2行2列）")]
    public void ReadSheet_MultipleRows()
    {
        var sst = new[] { "Name", "Age", "Alice", "30", "Bob", "25" };
        var body = Concat(
            LabelSst(0, 0, 0), LabelSst(0, 1, 1),
            LabelSst(1, 0, 2), LabelSst(1, 1, 3),
            LabelSst(2, 0, 4), LabelSst(2, 1, 5));
        var stream = BuildMinimalXls("Data", body, sst);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows = reader.ReadSheet("Data").ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal("Name", rows[0][0]);
        Assert.Equal("Alice", rows[1][0]);
        Assert.Equal("25", rows[2][1]);
    }

    // ─── 按名称读取工作表 ─────────────────────────────────────────────────

    [Fact, DisplayName("按工作表名称读取，大小写不敏感")]
    public void ReadSheet_ByName_CaseInsensitive()
    {
        var sst = new[] { "Data" };
        var body = LabelSst(0, 0, 0);
        var stream = BuildMinimalXls("MyData", body, sst);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows1 = reader.ReadSheet("mydata").ToList();
        var rows2 = reader.ReadSheet("MYDATA").ToList();
        Assert.Single(rows1);
        Assert.Single(rows2);
    }

    // ─── 按名读取不存在的工作表返回空 ────────────────────────────────────

    [Fact, DisplayName("按不存在的名称读取工作表返回空序列")]
    public void ReadSheet_NonExistentName_ReturnsEmpty()
    {
        var stream = BuildMinimalXls("Sheet1", []);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows = reader.ReadSheet("NoSuchSheet").ToList();
        Assert.Empty(rows);
    }

    // ─── ReadObjects 泛型映射 ─────────────────────────────────────────────

    private class Person
    {
        public String Name { get; set; } = String.Empty;
        public String Age { get; set; } = String.Empty;
    }

    [Fact, DisplayName("ReadObjects 将工作表行映射为对象列表")]
    public void ReadObjects_MapsToTypedObjects()
    {
        var sst = new[] { "Name", "Age", "Alice", "30" };
        var body = Concat(
            LabelSst(0, 0, 0), LabelSst(0, 1, 1),
            LabelSst(1, 0, 2), LabelSst(1, 1, 3));
        var stream = BuildMinimalXls("Sheet1", body, sst);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var people = reader.ReadObjects<Person>().ToList();
        Assert.Single(people);
        Assert.Equal("Alice", people[0].Name);
        Assert.Equal("30", people[0].Age);
    }

    // ─── 空工作表 ─────────────────────────────────────────────────────────

    [Fact, DisplayName("空工作表（无单元格）ReadSheet 返回空序列")]
    public void ReadSheet_EmptySheet_ReturnsEmpty()
    {
        var stream = BuildMinimalXls("Empty", []);
        var xls = BuildXls(stream);

        using var reader = new BiffReader(new MemoryStream(xls));
        var rows = reader.ReadSheet().ToList();
        Assert.Empty(rows);
    }

    // ─── 无效文件抛出异常 ─────────────────────────────────────────────────

    [Fact, DisplayName("非 OLE2 内容抛出 InvalidDataException")]
    public void InvalidOle2_ThrowsInvalidDataException()
    {
        var fakeData = new Byte[512];
        Encoding.ASCII.GetBytes("PK\x03\x04").CopyTo(fakeData, 0);
        using var ms = new MemoryStream(fakeData);
        Assert.Throws<InvalidDataException>(() => new BiffReader(ms));
    }

    // ─── 批注（NOTE + TXO）───────────────────────────────────────────────

    [Fact, DisplayName("读取批注（NOTE 作者 + TXO 文本）")]
    public void ReadComments_NoteAndTxo()
    {
        // NOTE：row=0,col=1,objectId=1,flags=0,cch=3,fHighByte=1,作者="李四"
        var noteData = Concat(
            LE2(0), LE2(1),                 // row, col
            LE2(1), LE2(0),                 // objectId, flags
            LE2(3),                         // cch = 3
            new Byte[] { 0x01 },            // fHighByte = UTF-16
            Encoding.Unicode.GetBytes("李四"));
        var noteRec = Rec(0x001C, noteData);

        // TXO：grbit=0,rot=0,cchText=2,cbRuns=0,ifntEmpty=0,reserved=0,文本="你好"
        var txoData = Concat(
            LE2(0), LE2(0),                 // grbit, rot
            LE2(2),                         // cchText = 2
            LE2(0), LE2(0), LE2(0),         // cbRuns, ifntEmpty, reserved
            Encoding.Unicode.GetBytes("你好"));
        var txoRec = Rec(0x01B6, txoData);

        var sheetBody = Concat(LabelSst(0, 0, 0), noteRec, txoRec);
        var xls = BuildXls(BuildMinimalXls("Sheet1", sheetBody, new[] { "值" }));

        using var reader = new BiffReader(new MemoryStream(xls));
        var comments = reader.GetComments();
        Assert.Single(comments);
        Assert.True(comments.TryGetValue((0, 1), out var c));
        Assert.Equal("你好", c.Text);
        Assert.Equal("李四", c.Author);
    }

    [Fact, DisplayName("无批注时 GetComments 返回空")]
    public void ReadComments_None()
    {
        var xls = BuildXls(BuildMinimalXls("Sheet1", LabelSst(0, 0, 0), new[] { "值" }));
        using var reader = new BiffReader(new MemoryStream(xls));
        Assert.Empty(reader.GetComments());
    }

    // ─── 条件格式（CONDFMT + CF）─────────────────────────────────────────

    [Fact, DisplayName("读取条件格式（cellIs 大于/小于 + 多规则）")]
    public void ReadConditionalFormats_CellIs()
    {
        // CONDFMT：count=2, r1=0, r2=1, c1=0, c2=0, flags=0 → A1:A2 两条规则
        var condFmtData = Concat(LE2(2), LE2(0), LE2(1), LE2(0), LE2(0), LE2(0));
        var condFmtRec = Rec(0x01B0, condFmtData);

        // CF1：type=1(cellIs)|op=4(greaterThan) → flags=0x41，f1=100
        var cf1Data = Concat(LE2(0x41), LE8(100.0), LE8(0.0));
        var cf1Rec = Rec(0x01B1, cf1Data);

        // CF2：type=1(cellIs)|op=5(lessThan) → flags=0x51，f1=50
        var cf2Data = Concat(LE2(0x51), LE8(50.0), LE8(0.0));
        var cf2Rec = Rec(0x01B1, cf2Data);

        var sheetBody = Concat(LabelSst(0, 0, 0), condFmtRec, cf1Rec, cf2Rec);
        var xls = BuildXls(BuildMinimalXls("Sheet1", sheetBody, new[] { "值" }));

        using var reader = new BiffReader(new MemoryStream(xls));
        var cfs = reader.GetConditionalFormats();
        Assert.Equal(2, cfs.Count);
        Assert.Equal("A1:A2", cfs[0].Range);
        Assert.Equal(ConditionalFormatValues.GreaterThan, cfs[0].Type);
        Assert.Equal("100", cfs[0].Value);
        Assert.Equal(ConditionalFormatValues.LessThan, cfs[1].Type);
        Assert.Equal("50", cfs[1].Value);
    }

    [Fact, DisplayName("读取条件格式（between 双值）")]
    public void ReadConditionalFormats_Between()
    {
        // CONDFMT：count=1, r1=2, r2=2, c1=3, c2=3 → D3:D3 一条规则
        var condFmtData = Concat(LE2(1), LE2(2), LE2(2), LE2(3), LE2(3), LE2(0));
        var condFmtRec = Rec(0x01B0, condFmtData);

        // CF：type=1|op=0(between) → flags=0x01，f1=10，f2=20
        var cfData = Concat(LE2(0x01), LE8(10.0), LE8(20.0));
        var cfRec = Rec(0x01B1, cfData);

        var sheetBody = Concat(LabelSst(0, 0, 0), condFmtRec, cfRec);
        var xls = BuildXls(BuildMinimalXls("Sheet1", sheetBody, new[] { "值" }));

        using var reader = new BiffReader(new MemoryStream(xls));
        var cfs = reader.GetConditionalFormats();
        Assert.Single(cfs);
        Assert.Equal("D3:D3", cfs[0].Range);
        Assert.Equal(ConditionalFormatValues.Between, cfs[0].Type);
        Assert.Equal("10", cfs[0].Value);
        Assert.Equal("20", cfs[0].Value2);
    }

    [Fact, DisplayName("无条件格式时 GetConditionalFormats 返回空")]
    public void ReadConditionalFormats_None()
    {
        var xls = BuildXls(BuildMinimalXls("Sheet1", LabelSst(0, 0, 0), new[] { "值" }));
        using var reader = new BiffReader(new MemoryStream(xls));
        Assert.Empty(reader.GetConditionalFormats());
    }

    [Fact, DisplayName("读取条件格式图案前景色（pattern fgColor 调色板还原）")]
    public void ReadConditionalFormats_Color()
    {
        // CONDFMT：count=1, r1=0, r2=0, c1=0, c2=0 → A1:A1 一条规则
        var condFmtData = Concat(LE2(1), LE2(0), LE2(0), LE2(0), LE2(0), LE2(0));
        var condFmtRec = Rec(0x01B0, condFmtData);

        // CF：flags=0x41 (cellIs|greaterThan)，f1=100，f2=0
        // font(4)+border(4)+pattern(4)+temp(2)：pattern 段 pattern=1(solid)+fgColor=2(红)+bgColor=0
        var cfData = Concat(
            LE2(0x41), LE8(100.0), LE8(0.0),
            LE2(0), LE2(0),                       // font(4)
            LE2(0), LE2(0),                       // border(4)
            new Byte[] { 0x01, 0x02, 0x00, 0x00 }, // pattern(4): pattern=1, fgColor=2, bgColor=0
            LE2(0));                              // temp(2)
        var cfRec = Rec(0x01B1, cfData);

        var sheetBody = Concat(LabelSst(0, 0, 0), condFmtRec, cfRec);
        var xls = BuildXls(BuildMinimalXls("Sheet1", sheetBody, new[] { "值" }));

        using var reader = new BiffReader(new MemoryStream(xls));
        var cfs = reader.GetConditionalFormats();
        Assert.Single(cfs);
        Assert.Equal(ConditionalFormatValues.GreaterThan, cfs[0].Type);
        Assert.Equal("100", cfs[0].Value);
        Assert.Equal("FF0000", cfs[0].Color);
    }
}
