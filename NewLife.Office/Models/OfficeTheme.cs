namespace NewLife.Office;

/// <summary>Office 主题色板工具</summary>
/// <remarks>
/// 内置 11 套主题色板，每套 6 个强调色（Accent1~6，16进制 RGB 无 # 前缀）。
/// 通过可选的 customResolver 委托支持外部注入自定义色板（如从数据库读取）。
/// <example>
/// <code>
/// // 获取内置色板
/// var colors = OfficeTheme.Get("blue");         // ["2563EB", "1D4ED8", ...]
/// var primary = OfficeTheme.GetPrimary("blue"); // "2563EB"
/// var light   = OfficeTheme.GetLight("blue");   // "DBEAFE"
///
/// // 注入外部色板（如从数据库读取）
/// var colors = OfficeTheme.Get("my-theme", name =>
/// {
///     var palette = CardStyleService.GetColors(name);
///     return palette != null ? [palette.Primary, palette.Secondary, ...] : null;
/// });
/// </code>
/// </example>
/// </remarks>
public static class OfficeTheme
{
    #region 公共方法
    /// <summary>获取主题强调色数组（Accent1~6，16进制 RGB 无 # 前缀）</summary>
    /// <param name="theme">主题名称（不区分大小写），null 或未知时返回 blue 色板</param>
    /// <param name="customResolver">自定义色板解析器，返回 null 时回退到内置色板</param>
    /// <returns>6 个颜色字符串数组</returns>
    public static String[] Get(String? theme, Func<String, String[]?>? customResolver = null)
    {
        if (!theme.IsNullOrEmpty() && customResolver != null)
        {
            var custom = customResolver(theme!);
            if (custom is { Length: > 0 })
                return custom;
        }
        return GetBuiltin(theme);
    }

    /// <summary>获取主题主色（Accent1，表头背景等深色场景）</summary>
    /// <param name="theme">主题名称</param>
    /// <param name="customResolver">自定义色板解析器</param>
    /// <returns>16进制 RGB 字符串（无 # 前缀）</returns>
    public static String GetPrimary(String? theme, Func<String, String[]?>? customResolver = null)
        => Get(theme, customResolver)[0];

    /// <summary>获取主题浅色（Accent6，斑马纹/卡片背景等浅色场景）</summary>
    /// <param name="theme">主题名称</param>
    /// <param name="customResolver">自定义色板解析器</param>
    /// <returns>16进制 RGB 字符串（无 # 前缀）</returns>
    public static String GetLight(String? theme, Func<String, String[]?>? customResolver = null)
        => Get(theme, customResolver)[5];

    /// <summary>获取主题默认幻灯片背景色</summary>
    /// <param name="theme">主题名称</param>
    /// <returns>16进制 RGB 字符串（无 # 前缀）</returns>
    public static String GetBackground(String? theme) =>
        (theme?.ToLowerInvariant() ?? String.Empty) switch
        {
            "dark" or "ocean" or "sunset" or "forest" or "slate" or "amber" => "0F172A",
            _ => "FFFFFF",
        };
    #endregion

    #region 内置色板
    private static String[] GetBuiltin(String? theme) =>
        (theme?.ToLowerInvariant() ?? String.Empty) switch
        {
            "blue"      => ["2563EB", "1D4ED8", "60A5FA", "93C5FD", "1E40AF", "DBEAFE"],
            "dark"      => ["6366F1", "4F46E5", "818CF8", "A5B4FC", "3730A3", "E0E7FF"],
            "corporate" => ["374151", "1F2937", "6B7280", "9CA3AF", "111827", "F3F4F6"],
            "warm"      => ["EA580C", "C2410C", "FB923C", "FED7AA", "9A3412", "FFF7ED"],
            "green"     => ["16A34A", "15803D", "4ADE80", "BBF7D0", "14532D", "DCFCE7"],
            "minimal"   => ["18181B", "27272A", "71717A", "A1A1AA", "09090B", "FAFAFA"],
            "ocean"     => ["0EA5E9", "0284C7", "38BDF8", "7DD3FC", "0369A1", "BAE6FD"],
            "sunset"    => ["F97316", "C026D3", "FB923C", "E879F9", "7C3AED", "F0ABFC"],
            "forest"    => ["059669", "065F46", "34D399", "A7F3D0", "064E3B", "D1FAE5"],
            "slate"     => ["64748B", "475569", "94A3B8", "CBD5E1", "1E293B", "F8FAFC"],
            "amber"     => ["F59E0B", "D97706", "FCD34D", "FDE68A", "92400E", "FEF3C7"],
            _           => ["2563EB", "1D4ED8", "60A5FA", "93C5FD", "1E40AF", "DBEAFE"],
        };
    #endregion
}
