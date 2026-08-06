using System.Text;

namespace NewLife.Office.Markdown;

/// <summary>Markdown 文本规范化工具</summary>
/// <remarks>
/// 对标 anydoc shared/text.rs：去除控制字符与布局不可见字符，转换 NBSP 为普通空格，
/// 剥离软连字符；保留 ZWNJ/ZWJ（阿拉伯/印度语形变与 emoji 序列需要）。
/// 所有格式前端在产出文本块前应统一调用清洗，保证输出洁净一致。
/// </remarks>
public static class MarkdownTextCleaner
{
    /// <summary>清洗文本：去控制字符、NBSP→空格、剥离软连字符，保留 ZWNJ/ZWJ</summary>
    /// <param name="text">原始文本，可为 null</param>
    /// <returns>清洗后文本，空输入返回空字符串</returns>
    public static String Clean(String? text)
    {
        if (String.IsNullOrEmpty(text)) return String.Empty;

        var sb = new StringBuilder(text!.Length);
        foreach (var ch in text!)
        {
            // 控制字符：仅保留制表/换行/回车
            if (ch < 0x20)
            {
                if (ch == '\t' || ch == '\n' || ch == '\r') sb.Append(ch);
                continue;
            }
            // NBSP → 普通空格
            if (ch == '\u00A0')
            {
                sb.Append(' ');
                continue;
            }
            // 软连字符（布局不可见）剥离
            if (ch == '\u00AD') continue;
            // DEL 与 C1 控制字符剔除
            if (ch == '\u007F' || ch >= '\u0080' && ch <= '\u009F') continue;
            sb.Append(ch);
        }
        return sb.ToString();
    }

    /// <summary>将单元格文本清洗为单行（换行转空格，供 GFM 表格单元格使用）</summary>
    /// <param name="text">原始文本，可为 null</param>
    /// <returns>单行文本</returns>
    public static String CleanCell(String? text)
    {
        var value = Clean(text);
        return value.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
    }
}
