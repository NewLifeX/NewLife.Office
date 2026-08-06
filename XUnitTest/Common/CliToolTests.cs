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

    [Fact, DisplayName("CLI：无参数返回 1")]
    public void NoArgs_Returns1()
    {
        var code = Program.Main([]);
        Assert.Equal(1, code);
    }
}
