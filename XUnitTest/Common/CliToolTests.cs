using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office.Cli;
using NewLife.Office.Markdown;
using NewLife.Office.Pdf;
using Xunit;

namespace XUnitTest.Common;

/// <summary>CLI 工具链测试（GEN-2，对标 anydoc CLI/Agent Skill）</summary>
/// <remarks>
/// 覆盖 convert/info 命令的退出码与输出，验证命令行入口可用。
/// </remarks>
public class CliToolTests
{
    private static String GetTempDir() => Path.Combine(Path.GetTempPath(), "cli_" + Guid.NewGuid().ToString("N"));

    [Fact, DisplayName("CLI：convert 输出到文件")]
    public void Convert_ToFile()
    {
        var dir = GetTempDir();
        Directory.CreateDirectory(dir);
        try
        {
            var input = Path.Combine(dir, "input.md");
            File.WriteAllText(input, "# 标题\n\n正文内容", Encoding.UTF8);
            var output = Path.Combine(dir, "output.md");

            var code = Program.Main(["convert", input, output]);

            Assert.Equal(0, code);
            Assert.True(File.Exists(output));
            Assert.Contains("标题", File.ReadAllText(output));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("CLI：convert 输出到 stdout（无 output 参数）")]
    public void Convert_ToStdout()
    {
        var dir = GetTempDir();
        Directory.CreateDirectory(dir);
        try
        {
            var input = Path.Combine(dir, "input.md");
            File.WriteAllText(input, "# 标题\n\n正文", Encoding.UTF8);
            var code = Program.Main(["convert", input]);
            Assert.Equal(0, code);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("CLI：convert 不存在的文件返回 1")]
    public void Convert_FileNotFound()
    {
        var code = Program.Main(["convert", @"C:\nonexistent_file_xyz.docx"]);
        Assert.Equal(1, code);
    }

    [Fact, DisplayName("CLI：info 命令返回格式")]
    public void Info_Format()
    {
        var dir = GetTempDir();
        Directory.CreateDirectory(dir);
        try
        {
            var input = Path.Combine(dir, "input.md");
            File.WriteAllText(input, "hello", Encoding.UTF8);
            var code = Program.Main(["info", input]);
            Assert.Equal(0, code);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("CLI：convert 扩展名错误时内容识别兜底（GEN-2b）")]
    public void Convert_ContentDetectFallback()
    {
        var dir = GetTempDir();
        Directory.CreateDirectory(dir);
        try
        {
            // 用 PdfWriter 生成 PDF，但扩展名写 .dat（内容识别应兜底）
            var pdfBytes = BuildSimplePdf();
            var input = Path.Combine(dir, "data.dat");
            File.WriteAllBytes(input, pdfBytes);
            var output = Path.Combine(dir, "out.md");

            var code = Program.Main(["convert", input, output]);

            Assert.Equal(0, code);
            Assert.True(File.Exists(output));
            Assert.Contains("Hello", File.ReadAllText(output));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    /// <summary>生成含英文文本的最小 PDF（PdfWriter）</summary>
    private static Byte[] BuildSimplePdf()
    {
        using var ms = new MemoryStream();
        using (var writer = new PdfWriter())
        {
            writer.BeginPage();
            writer.DrawText("Hello PDF", 56, 780, 12);
            writer.Save(ms);
        }
        return ms.ToArray();
    }

    [Fact, DisplayName("CLI：doc2docx 转换 doc 为 docx")]
    public void Doc2Docx_Convert()
    {
        var dir = GetTempDir();
        Directory.CreateDirectory(dir);
        try
        {
            var input = Path.Combine(dir, "input.doc");
            File.WriteAllBytes(input, BuildMinimalDoc("Hello Doc"));
            var output = Path.Combine(dir, "output.docx");

            var code = Program.Main(["doc2docx", input, output]);

            Assert.Equal(0, code);
            Assert.True(File.Exists(output));

            // 输出 docx 可被 WordReader 读取
            using var reader = new NewLife.Office.Word.WordReader(output);
            Assert.Contains("Hello Doc", reader.ReadFullText());
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("CLI：doc2docx 参数不足返回 1")]
    public void Doc2Docx_MissingArgs()
    {
        var code = Program.Main(["doc2docx", "only_input.doc"]);
        Assert.Equal(1, code);
    }

    /// <summary>构建最小 .doc（WordDocument 流 + OLE2 容器）</summary>
    private static Byte[] BuildMinimalDoc(String text)
    {
        static Byte[] LE2(UInt16 v) => new Byte[] { (Byte)(v & 0xFF), (Byte)(v >> 8) };
        static Byte[] LE4(UInt32 v) =>
            new Byte[] { (Byte)(v & 0xFF), (Byte)((v >> 8) & 0xFF), (Byte)((v >> 16) & 0xFF), (Byte)(v >> 24) };

        var chars = new List<Byte>();
        foreach (var ch in text) chars.Add((Byte)ch);
        chars.Add(0x0D);
        var textBytes = chars.ToArray();
        var textLen = textBytes.Length;

        const Int32 FibSize = 400;
        const Int32 PlcPcdSize = 2 * 4 + 1 * 8;
        const Int32 ClxBlockSize = 1 + 4 + PlcPcdSize;
        var ClxStart = FibSize;
        var textOffset = ClxStart + ClxBlockSize;
        var fcValue = textOffset * 2;

        var fib = new Byte[FibSize];
        fib[0] = 0xEC; fib[1] = 0xA5;
        fib[2] = 0xC1; fib[3] = 0x00;
        fib[32] = 14; fib[33] = 0;
        fib[62] = 22; fib[63] = 0;
        fib[152] = 74; fib[153] = 0;
        Array.Copy(LE4((UInt32)ClxStart), 0, fib, 258, 4);
        Array.Copy(LE4((UInt32)ClxBlockSize), 0, fib, 262, 4);

        var plcPcd = Concat(LE4(0), LE4((UInt32)textLen), Concat(LE2(0), LE4((UInt32)(fcValue | (1 << 30))), LE2(0)));
        var clxBlock = Concat(new Byte[] { 0x02 }, LE4((UInt32)plcPcd.Length), plcPcd);

        var wordDoc = Concat(fib, clxBlock, textBytes);
        var cfb = new NewLife.Office.Ole2.CfbDocument();
        cfb.Root.AddStream("WordDocument", wordDoc);
        return cfb.ToBytes();

        static Byte[] Concat(params Byte[][] parts)
        {
            var total = 0;
            foreach (var p in parts) total += p.Length;
            var buf = new Byte[total];
            var pos = 0;
            foreach (var p in parts) { Array.Copy(p, 0, buf, pos, p.Length); pos += p.Length; }
            return buf;
        }
    }

    [Fact, DisplayName("CLI：无参数返回 1")]
    public void NoArgs_Returns1()
    {
        var code = Program.Main([]);
        Assert.Equal(1, code);
    }
}
