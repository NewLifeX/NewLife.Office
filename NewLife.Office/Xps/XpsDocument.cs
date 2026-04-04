using System.Text;

namespace NewLife.Office;

/// <summary>XPS 文档包装，封装页面列表并提供文本/Markdown 提取能力</summary>
public class XpsDocument : ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>页面列表</summary>
    public List<XpsPage> Pages { get; }

    /// <summary>文档属性</summary>
    public XpsProperties? Properties { get; set; }
    #endregion

    #region 构造
    /// <summary>实例化 XPS 文档包装</summary>
    /// <param name="pages">页面列表</param>
    public XpsDocument(List<XpsPage> pages) => Pages = pages ?? [];

    /// <summary>实例化 XPS 文档包装</summary>
    /// <param name="pages">页面列表</param>
    /// <param name="properties">文档属性</param>
    public XpsDocument(List<XpsPage> pages, XpsProperties? properties)
    {
        Pages = pages ?? [];
        Properties = properties;
    }
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（各页文本拼接）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText()
    {
        if (Pages == null || Pages.Count == 0) return null;

        var sb = new StringBuilder();
        for (var i = 0; i < Pages.Count; i++)
        {
            if (i > 0) sb.AppendLine();
            sb.Append(Pages[i].Text);
        }
        return sb.ToString();
    }

    /// <summary>提取 Markdown 格式（页间用分隔线分隔）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown()
    {
        if (Pages == null || Pages.Count == 0) return null;

        var sb = new StringBuilder();
        for (var i = 0; i < Pages.Count; i++)
        {
            if (i > 0)
            {
                sb.AppendLine();
                sb.AppendLine("---");
                sb.AppendLine();
            }
            sb.Append(Pages[i].Text);
        }
        return sb.ToString();
    }
    #endregion
}
