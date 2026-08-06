using System.Text;

namespace NewLife.Office.Word;

/// <summary>OMML 数学公式（W13，对标 Open XML SDK/Aspose.Words）</summary>
/// <remarks>
/// 表示 Word OfficeMath（OMML）公式。提供分数/上下标/根式/积分等常用结构的
/// 静态工厂方法，也可直接传入原始 OMML 内部 XML。
/// </remarks>
public class OmmlFormula
{
    /// <summary>OMML 内部 XML（如 <c>&lt;m:f&gt;...&lt;/m:f&gt;</c>，不含 m:oMath 包裹）</summary>
    public String Xml { get; set; } = String.Empty;

    /// <summary>生成简单文本（m:r + m:t）</summary>
    /// <param name="text">文本内容</param>
    /// <returns>公式对象</returns>
    public static OmmlFormula Text(String text)
    {
        return new OmmlFormula { Xml = $"<m:r><m:t xml:space=\"preserve\">{Esc(text)}</m:t></m:r>" };
    }

    /// <summary>生成分数</summary>
    /// <param name="numerator">分子（公式或文本）</param>
    /// <param name="denominator">分母（公式或文本）</param>
    /// <returns>公式对象</returns>
    public static OmmlFormula Fraction(String numerator, String denominator)
    {
        return new OmmlFormula
        {
            Xml = $"<m:f><m:num>{ToOmml(numerator)}</m:num><m:den>{ToOmml(denominator)}</m:den></m:f>",
        };
    }

    /// <summary>生成上标（如 x²）</summary>
    /// <param name="baseExpr">底数（公式或文本）</param>
    /// <param name="sup">上标（公式或文本）</param>
    /// <returns>公式对象</returns>
    public static OmmlFormula SuperScript(String baseExpr, String sup)
    {
        return new OmmlFormula
        {
            Xml = $"<m:sSup><m:e>{ToOmml(baseExpr)}</m:e><m:sup>{ToOmml(sup)}</m:sup></m:sSup>",
        };
    }

    /// <summary>生成下标（如 x₁）</summary>
    /// <param name="baseExpr">底数（公式或文本）</param>
    /// <param name="sub">下标（公式或文本）</param>
    /// <returns>公式对象</returns>
    public static OmmlFormula SubScript(String baseExpr, String sub)
    {
        return new OmmlFormula
        {
            Xml = $"<m:sSub><m:e>{ToOmml(baseExpr)}</m:e><m:sub>{ToOmml(sub)}</m:sub></m:sSub>",
        };
    }

    /// <summary>生成根式（平方根或 n 次根）</summary>
    /// <param name="radicand">被开方数（公式或文本）</param>
    /// <param name="index">根次数（null 为平方根）</param>
    /// <returns>公式对象</returns>
    public static OmmlFormula Radical(String radicand, String? index = null)
    {
        if (String.IsNullOrEmpty(index))
        {
            return new OmmlFormula
            {
                Xml = $"<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/><m:e>{ToOmml(radicand)}</m:e></m:rad>",
            };
        }
        return new OmmlFormula
        {
            Xml = $"<m:rad><m:deg>{ToOmml(index)}</m:deg><m:e>{ToOmml(radicand)}</m:e></m:rad>",
        };
    }

    /// <summary>生成积分（∫）</summary>
    /// <param name="integrand">被积函数（公式或文本）</param>
    /// <param name="lower">下限（可空）</param>
    /// <param name="upper">上限（可空）</param>
    /// <returns>公式对象</returns>
    public static OmmlFormula Integral(String integrand, String? lower = null, String? upper = null)
    {
        var sub = String.IsNullOrEmpty(lower) ? "<m:sub/>" : $"<m:sub>{ToOmml(lower)}</m:sub>";
        var sup = String.IsNullOrEmpty(upper) ? "<m:sup/>" : $"<m:sup>{ToOmml(upper)}</m:sup>";
        return new OmmlFormula
        {
            Xml = $"<m:nary><m:naryPr><m:chr m:val=\"∫\"/><m:limLoc m:val=\"subSup\"/></m:naryPr>{sub}{sup}<m:e>{ToOmml(integrand)}</m:e></m:nary>",
        };
    }

    /// <summary>直接使用原始 OMML 内部 XML</summary>
    /// <param name="innerXml">OMML 内部 XML</param>
    /// <returns>公式对象</returns>
    public static OmmlFormula Raw(String innerXml) => new() { Xml = innerXml };

    /// <summary>将简单文本转换为 OMML 运行元素（已含 &lt;m:r&gt; 的按原样返回）</summary>
    private static String ToOmml(String expr)
    {
        if (expr.IsNullOrEmpty()) return "<m:r/>";
        // 若已包含 OMML 结构标记（m: 前缀），视为已格式化
        if (expr.Contains("<m:", StringComparison.OrdinalIgnoreCase)) return expr;
        return $"<m:r><m:t xml:space=\"preserve\">{Esc(expr)}</m:t></m:r>";
    }

    /// <summary>XML 转义</summary>
    private static String Esc(String text) => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
