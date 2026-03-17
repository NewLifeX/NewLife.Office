using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>DocReader .doc 二进制格式读取器单元测试</summary>
public class DocReaderTests
{
    // ─── 构建最小 .doc 文件辅助 ─────────────────────────────────────────

    private static Byte[] LE2(UInt16 v) => new Byte[] { (Byte)(v & 0xFF), (Byte)(v >> 8) };
    private static Byte[] LE4(UInt32 v) =>
        new Byte[] { (Byte)(v & 0xFF), (Byte)((v >> 8) & 0xFF), (Byte)((v >> 16) & 0xFF), (Byte)(v >> 24) };

    private static Byte[] Concat(params Byte[][] parts)
    {
        var total = 0;
        foreach (var p in parts) total += p.Length;
        var buf = new Byte[total];
        var pos = 0;
        foreach (var p in parts) { Array.Copy(p, 0, buf, pos, p.Length); pos += p.Length; }
        return buf;
    }

    /// <summary>
    /// 构建最小有效 .doc WordDocument 流（Word 97 格式）。
    /// 文本以 ANSI 压缩形式内嵌在 WordDocument 流中。
    /// </summary>
    /// <param name="text">要写入的文本（段落之间以 \r 分隔，\r 自动转为 Word 段落符）</param>
    private static Byte[] BuildWordDocStream(String text)
    {
        // ─── 目标文本：将 \n/\r 换为 0x0D（Word 段落符）─────────────────
        var chars = new List<Byte>();
        foreach (var ch in text)
        {
            if (ch == '\r' || ch == '\n')
                chars.Add(0x0D);
            else
                chars.Add((Byte)ch);
        }
        // Word 文档的最后一个段落符：每个 .doc 都以 0x0D 结尾
        chars.Add(0x0D);

        var textBytes = chars.ToArray();          // ANSI 文本
        var textLen = textBytes.Length;           // 字符数（= 字节数，因为 ANSI）

        // ─── FIB 布局参数（Word 97 标准值）──────────────────────────────
        // csw = 14, cslw = 22
        // FIB base (32) + csw(2) + FibRgW97(28) + cslw(2) + FibRgLw97(88) + cbRgFcLcb(2)
        //   = 32 + 2 + 28 + 2 + 88 + 2 = 154 → FibRgFcLcb97 starts here
        // CLX entry (index 13) at 154 + 13*8 = 258
        const Int32 FibSize = 400;                // 留足够空间放 FIB
        const Int32 FcClxInFib = 258;             // 固定偏移
        const Int32 LcbClxInFib = 262;

        // CLX 紧跟在 FIB 后面
        const Int32 ClxStart = FibSize;

        // ANSI 文本紧跟在 CLX 后
        // CLX 大小：1(clxt) + 4(lcb) + PlcPcd大小
        // PlcPcd：2 个 CP(= 2*4=8字节) + 1 个 PCD(= 8字节) = 16字节
        const Int32 PlcPcdSize = 2 * 4 + 1 * 8;  // (n+1)*4 + n*8 with n=1
        // CLX = PCDT: clxt(1) + lcb(4) + PlcPcd(PlcPcdSize) = 21 bytes
        // （此处 LcbClx 仅用于注释说明，实际用 ClxBlockSize 计算）

        const Int32 ClxBlockSize = 1 + 4 + PlcPcdSize;

        // 文本从 ClxStart+ClxBlockSize 开始，但 ANSI 压缩偏移 = textOffset * 2！
        // 这是因为 fCompressed=1 时 fc = byteOffset * 2
        var textOffset = ClxStart + ClxBlockSize;
        var fcValue = textOffset * 2;             // 压缩存储，fc = byteOffset * 2

        // ─── 构建 FIB ─────────────────────────────────────────────────
        var fib = new Byte[FibSize];

        // wIdent = 0xA5EC
        fib[0] = 0xEC; fib[1] = 0xA5;
        // nFib = 0x00C1 (Word 97)
        fib[2] = 0xC1; fib[3] = 0x00;

        // csw = 14 at offset 32
        fib[32] = 14; fib[33] = 0;
        // FibRgW97: 28 bytes of zeros (offset 34..61)
        // cslw = 22 at offset 62
        fib[62] = 22; fib[63] = 0;
        // FibRgLw97: 88 bytes of zeros (offset 64..151)
        // ccbRgFcLcb = 74 at offset 152
        fib[152] = 74; fib[153] = 0;
        // FibRgFcLcb97 at offset 154: 74 entries * 8 bytes = 592 bytes (only need index 13)
        // Entry 13 (fcClx at 258, lcbClx at 262)
        var fcClxBytes = LE4((UInt32)ClxStart);
        var lcbClxBytes = LE4((UInt32)ClxBlockSize);
        fib[FcClxInFib] = fcClxBytes[0]; fib[FcClxInFib + 1] = fcClxBytes[1];
        fib[FcClxInFib + 2] = fcClxBytes[2]; fib[FcClxInFib + 3] = fcClxBytes[3];
        fib[LcbClxInFib] = lcbClxBytes[0]; fib[LcbClxInFib + 1] = lcbClxBytes[1];
        fib[LcbClxInFib + 2] = lcbClxBytes[2]; fib[LcbClxInFib + 3] = lcbClxBytes[3];

        // ─── 构建 CLX（仅 PCDT，无 PRC）─────────────────────────────────
        // PCDT: clxt(0x02) + lcb(4) + PlcPcd
        var aCP0 = LE4(0);                              // CP 起点 = 0
        var aCP1 = LE4((UInt32)textLen);                // CP 终点 = textLen
        var pcdClsPcd = LE2(0);                         // clsPcd（无 fNoParaMark）
        // FcCompressed: fc | (fCompressed=1 << 30)
        var fcComp = (UInt32)(fcValue | (1 << 30));     // bit30=1 → ANSI
        var pcdFc = LE4(fcComp);
        var pcdPrm = LE2(0);
        var pcd = Concat(pcdClsPcd, pcdFc, pcdPrm);    // 8 bytes

        var plcPcd = Concat(aCP0, aCP1, pcd);           // (n+1)*4 + n*8

        var lcbPlcPcdBytes = LE4((UInt32)plcPcd.Length);
        var clxBlock = Concat(new Byte[] { 0x02 }, lcbPlcPcdBytes, plcPcd);

        // ─── 拼合 WordDocument 流 ──────────────────────────────────────
        return Concat(fib, clxBlock, textBytes);
    }

    /// <summary>将 WordDocument 流打包进 OLE2/CFB 容器为 .doc 字节</summary>
    private static Stream BuildDoc(String text)
    {
        var wordDocBytes = BuildWordDocStream(text);
        var doc = new CfbDocument();
        doc.Root.AddStream("WordDocument", wordDocBytes);
        return new MemoryStream(doc.ToBytes());
    }

    // ─── 测试 ─────────────────────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("ReadFullText 返回单行文本")]
    public void ReadFullText_SingleLine()
    {
        using var stream = BuildDoc("Hello World");
        using var reader = new DocReader(stream);
        var text = reader.ReadFullText();
        Assert.Contains("Hello World", text);
    }

    [Fact, System.ComponentModel.DisplayName("ReadParagraphs 按段落分隔返回")]
    public void ReadParagraphs_MultipleParagraphs()
    {
        // Word 段落符 0x0D 在流中内嵌；我们用 \r 表示段落边界
        using var stream = BuildDoc("First\rSecond\rThird");
        using var reader = new DocReader(stream);
        var paras = reader.ReadParagraphs().ToList();
        Assert.True(paras.Count >= 3);
        Assert.Contains("First", paras);
        Assert.Contains("Second", paras);
        Assert.Contains("Third", paras);
    }

    [Fact, System.ComponentModel.DisplayName("ReadFullText 内容含特殊符号不崩溃")]
    public void ReadFullText_WithDigitsAndPunctuation()
    {
        const String expected = "Price: $9.99 (2024)";
        using var stream = BuildDoc(expected);
        using var reader = new DocReader(stream);
        var text = reader.ReadFullText();
        Assert.Contains("Price", text);
        Assert.Contains("9.99", text);
    }

    [Fact, System.ComponentModel.DisplayName("ReadParagraphs 空文档不返回段落")]
    public void ReadParagraphs_EmptyContent()
    {
        // 仅含段落符，无实际文字
        using var stream = BuildDoc(String.Empty);
        using var reader = new DocReader(stream);
        // 空内容只有一个隐含的 0x0D → 文本为空或只含换行 → 段落数 0
        var paras = reader.ReadParagraphs().ToList();
        Assert.Empty(paras);
    }

    [Fact, System.ComponentModel.DisplayName("非 OLE2 格式抛出 InvalidDataException")]
    public void InvalidOle2_Throws()
    {
        var fakeData = new Byte[512];
        Encoding.ASCII.GetBytes("PK\x03\x04").CopyTo(fakeData, 0);
        using var ms = new MemoryStream(fakeData);
        Assert.Throws<InvalidDataException>(() => new DocReader(ms));
    }

    [Fact, System.ComponentModel.DisplayName("文件中存在 WordDocument 流但 wIdent 无效时抛出")]
    public void InvalidWIdent_Throws()
    {
        // 构建一个 WordDocument 流但 wIdent = 0xFFFF（无效）
        var wordDoc = new Byte[512];
        wordDoc[0] = 0xFF; wordDoc[1] = 0xFF;  // wIdent = 0xFFFF
        var cfb = new CfbDocument();
        cfb.Root.AddStream("WordDocument", wordDoc);
        using var ms = new MemoryStream(cfb.ToBytes());
        Assert.Throws<InvalidDataException>(() => new DocReader(ms));
    }

    [Fact, System.ComponentModel.DisplayName("ReadTables 识别含 0x07 标记的表格行")]
    public void ReadTables_BasicTable()
    {
        // 在 Word 二进制格式中，0x07 是表格单元格结束符
        // 注意：\u0007 是固定4位转义，避免 \x07B 被解析成 0x7B='{'  
        // 模拟一个 2 行 3 列的表格：每行以 \u0007 分隔单元格，行末跟 \r（段落符）
        const String Sep = "\u0007";
        var tableText = "A" + Sep + "B" + Sep + "C" + Sep + "\rX" + Sep + "Y" + Sep + "Z" + Sep + "\r";
        using var stream = BuildDoc(tableText);
        using var reader = new DocReader(stream);

        var tables = reader.ReadTables().ToList();
        Assert.True(tables.Count >= 1, "应识别出至少一张表格");

        var tbl = tables[0];
        Assert.Equal(2, tbl.Length);
        Assert.Equal(3, tbl[0].Length);
        Assert.Equal("A", tbl[0][0]);
        Assert.Equal("B", tbl[0][1]);
        Assert.Equal("C", tbl[0][2]);
        Assert.Equal("X", tbl[1][0]);
        Assert.Equal("Y", tbl[1][1]);
        Assert.Equal("Z", tbl[1][2]);
    }

    [Fact, System.ComponentModel.DisplayName("ReadTables 文档无表格时返回空序列")]
    public void ReadTables_NoTable_ReturnsEmpty()
    {
        using var stream = BuildDoc("Just plain text\rAnother paragraph");
        using var reader = new DocReader(stream);

        var tables = reader.ReadTables().ToList();
        Assert.Empty(tables);
    }

    [Fact, System.ComponentModel.DisplayName("ReadTables 表格与普通段落混合时正确分组")]
    public void ReadTables_MixedContent()
    {
        // 普通段落 → 表格 → 普通段落
        const String Sep = "\u0007";
        var text = "Title\rA" + Sep + "B" + Sep + "\rC" + Sep + "D" + Sep + "\rFooter";
        using var stream = BuildDoc(text);
        using var reader = new DocReader(stream);

        var tables = reader.ReadTables().ToList();
        Assert.Single(tables);
        Assert.Equal(2, tables[0].Length);
        Assert.Equal(2, tables[0][0].Length);
    }
}
