using System.Globalization;

namespace XUnitTest.Common;

/// <summary>文化相关测试辅助：globalization-invariant 运行时没有具体文化，相关用例应明确跳过</summary>
internal static class CultureTestHelper
{
    /// <summary>当前运行时是否支持具体文化（非 invariant 模式）</summary>
    public static Boolean HasSpecificCultures { get; } = Probe();

    private static Boolean Probe()
    {
        try
        {
            _ = CultureInfo.GetCultureInfo("en-US");
            return true;
        }
        catch (CultureNotFoundException)
        {
            return false;
        }
    }

    /// <summary>invariant 模式下输出跳过原因并返回 true，调用方直接 return</summary>
    /// <returns>是否应跳过当前用例</returns>
    public static Boolean SkipIfInvariant()
    {
        if (HasSpecificCultures) return false;

        Console.WriteLine("跳过：当前运行时为 globalization-invariant 模式，不支持具体文化。");
        return true;
    }
}
