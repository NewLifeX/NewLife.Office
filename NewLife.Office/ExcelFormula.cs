using System.Globalization;

namespace NewLife.Office;

/// <summary>Excel 公式封装，用于在 WriteRow 中嵌入公式单元格</summary>
/// <remarks>
/// 将 ExcelFormula 实例作为值传入 WriteRow / WriteObjects 的数据行，
/// ExcelWriter 会生成 &lt;f&gt;...&lt;/f&gt; 公式单元格而非普通 &lt;v&gt; 值单元格。
/// CachedValue 可选填上次计算结果，供不计算公式的读取方（如 NPOI 只读模式）快速显示。
/// </remarks>
public class ExcelFormula
{
    #region 属性
    /// <summary>公式文本（不含等号，如 "SUM(A1:A10)"）</summary>
    public String Formula { get; }

    /// <summary>缓存值（可空，用于读取时快速显示，不参与实际计算）</summary>
    public Object? CachedValue { get; }
    #endregion

    #region 构造
    /// <summary>实例化公式封装</summary>
    /// <param name="formula">公式文本（不含等号）</param>
    /// <param name="cachedValue">缓存的上次计算结果（可空）</param>
    public ExcelFormula(String formula, Object? cachedValue = null)
    {
        if (formula.IsNullOrEmpty()) throw new ArgumentNullException(nameof(formula));
        Formula = formula;
        CachedValue = cachedValue;
    }
    #endregion

    /// <summary>返回公式描述</summary>
    public override String ToString() => $"={Formula}";
}
