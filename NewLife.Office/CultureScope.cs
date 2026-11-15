using System.Globalization;

namespace NewLife.Office;

/// <summary>临时把当前线程区域设置为固定区域，用于文档序列化</summary>
/// <remarks>
/// PDF/XPS/OOXML 等格式的数值必须使用 "." 作为小数点，与宿主线程区域无关；
/// 序列化期间切换到固定区域，退出后恢复原区域，不影响调用方其他逻辑。
/// </remarks>
internal readonly struct CultureScope : IDisposable
{
    private readonly CultureInfo _original;

    private CultureScope(CultureInfo culture)
    {
        _original = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = culture;
    }

    /// <summary>切换到固定区域（小数点为 "."）</summary>
    /// <returns>离开时恢复原区域的令牌</returns>
    public static CultureScope Invariant() => new(CultureInfo.InvariantCulture);

    /// <summary>恢复原区域</summary>
    public void Dispose() => CultureInfo.CurrentCulture = _original;
}
