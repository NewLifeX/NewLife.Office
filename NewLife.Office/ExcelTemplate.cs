using System.IO.Compression;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace NewLife.Office;

/// <summary>Excel模板填充器</summary>
/// <remarks>
/// 基于xlsx模板文件，替换占位符（{{Key}}）并可选扩展数据行。
/// 仅修改工作表中的共享字符串和单元格值，不改变原有样式/格式。
/// </remarks>
public class ExcelTemplate
{
    #region 属性
    /// <summary>模板文件路径</summary>
    public String TemplatePath { get; }

    /// <summary>文本编码</summary>
    public Encoding Encoding { get; set; } = Encoding.UTF8;
    #endregion

    #region 构造
    /// <summary>实例化模板填充器</summary>
    /// <param name="templatePath">模板xlsx文件路径</param>
    public ExcelTemplate(String templatePath)
    {
        if (templatePath.IsNullOrEmpty()) throw new ArgumentNullException(nameof(templatePath));
        TemplatePath = templatePath.GetFullPath();
    }
    #endregion

    #region 方法
    /// <summary>填充模板并保存到目标路径</summary>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="data">占位符数据（键为占位符名称，值为替换内容）</param>
    public void Fill(String outputPath, IDictionary<String, Object> data)
    {
        if (outputPath.IsNullOrEmpty()) throw new ArgumentNullException(nameof(outputPath));
        if (data == null) throw new ArgumentNullException(nameof(data));

        outputPath = outputPath.EnsureDirectory(true).GetFullPath();

        // 复制模板到输出位置
        File.Copy(TemplatePath, outputPath, true);

        // 打开输出文件进行修改
        using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
        using var za = new ZipArchive(fs, ZipArchiveMode.Update, false, Encoding);

        // 替换共享字符串中的占位符
        ReplaceInSharedStrings(za, data);

        // 替换各工作表中内联文本的占位符
        foreach (var entry in za.Entries.ToArray())
        {
            if (entry.FullName.StartsWithIgnoreCase("xl/worksheets/") && entry.Name.EndsWithIgnoreCase(".xml"))
            {
                ReplaceInSheet(za, entry, data);
            }
        }
    }

    /// <summary>填充模板到流</summary>
    /// <param name="output">输出流</param>
    /// <param name="data">占位符数据</param>
    public void Fill(Stream output, IDictionary<String, Object> data)
    {
        if (output == null) throw new ArgumentNullException(nameof(output));
        if (data == null) throw new ArgumentNullException(nameof(data));

        // 将模板复制到输出流
        using (var templateFs = new FileStream(TemplatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            templateFs.CopyTo(output);
        }
        output.Position = 0;

        using var za = new ZipArchive(output, ZipArchiveMode.Update, true, Encoding);

        ReplaceInSharedStrings(za, data);

        foreach (var entry in za.Entries.ToArray())
        {
            if (entry.FullName.StartsWithIgnoreCase("xl/worksheets/") && entry.Name.EndsWithIgnoreCase(".xml"))
            {
                ReplaceInSheet(za, entry, data);
            }
        }
    }

    private void ReplaceInSharedStrings(ZipArchive za, IDictionary<String, Object> data)
    {
        var entry = za.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return;

        String xml;
        using (var sr = new StreamReader(entry.Open(), Encoding))
        {
            xml = sr.ReadToEnd();
        }

        var modified = ReplacePlaceholders(xml, data);
        if (modified == xml) return;

        // 删除旧条目并创建新条目
        var fullName = entry.FullName;
        entry.Delete();
        var newEntry = za.CreateEntry(fullName);
        using var sw = new StreamWriter(newEntry.Open(), Encoding);
        sw.Write(modified);
    }

    private void ReplaceInSheet(ZipArchive za, ZipArchiveEntry entry, IDictionary<String, Object> data)
    {
        String xml;
        using (var sr = new StreamReader(entry.Open(), Encoding))
        {
            xml = sr.ReadToEnd();
        }

        var modified = ReplacePlaceholders(xml, data);
        if (modified == xml) return;

        var fullName = entry.FullName;
        entry.Delete();
        var newEntry = za.CreateEntry(fullName);
        using var sw = new StreamWriter(newEntry.Open(), Encoding);
        sw.Write(modified);
    }

    private static String ReplacePlaceholders(String xml, IDictionary<String, Object> data)
    {
        // 匹配 {{Key}} 占位符
        return Regex.Replace(xml, @"\{\{(\w+)\}\}", m =>
        {
            var key = m.Groups[1].Value;
            if (data.TryGetValue(key, out var val))
            {
                return SecurityElement.Escape(val?.ToString() ?? "") ?? "";
            }
            return m.Value; // 未匹配的占位符保持原样
        });
    }
    #endregion
}
