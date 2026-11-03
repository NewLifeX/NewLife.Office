namespace NewLife.Office.Word;

/// <summary>长度单位值，用于在厘米/英寸/磅/缇/EMU 之间安全换算</summary>
/// <remarks>
/// OOXML 中不同属性使用不同单位：页面尺寸与缩进用缇（twips，1/1440 英寸），
/// 图片尺寸用 EMU（1/914400 英寸），字号用半磅。手写换算极易出错，本类型统一以 EMU 为内部基准。
/// <para>单位换算关系：1 英寸 = 914400 EMU = 1440 缇 = 72 磅 = 2.54 厘米。</para>
/// <example>
/// <code>
/// var w = Length.FromCm(10);            // 10 厘米
/// var twips = w.Twips;                  // → 5669.29…（段落缩进/单元格宽度用）
/// var emu = w.Emu;                      // → 3600000（图片尺寸用）
///
/// Length l = Length.Parse("2.5cm");     // 支持 cm/mm/pt/in/twip/emu 后缀
/// </code>
/// </example>
/// </remarks>
public readonly struct Length : IEquatable<Length>
{
    #region 常量
    /// <summary>每英寸 EMU 数</summary>
    public const Double EmuPerInch = 914400;

    /// <summary>每磅 EMU 数</summary>
    public const Double EmuPerPoint = 12700;

    /// <summary>每缇 EMU 数</summary>
    public const Double EmuPerTwip = 635;

    /// <summary>每厘米 EMU 数</summary>
    public const Double EmuPerCm = 360000;
    #endregion

    #region 属性
    /// <summary>以 EMU 表示的长度（内部基准单位）</summary>
    public Double Emu { get; }

    /// <summary>以缇（twips，1/1440 英寸）表示的长度</summary>
    public Double Twips => Emu / EmuPerTwip;

    /// <summary>以磅（point）表示的长度</summary>
    public Double Points => Emu / EmuPerPoint;

    /// <summary>以厘米表示的长度</summary>
    public Double Cm => Emu / EmuPerCm;

    /// <summary>以英寸表示的长度</summary>
    public Double Inches => Emu / EmuPerInch;
    #endregion

    #region 构造
    /// <summary>用 EMU 值构造</summary>
    /// <param name="emu">EMU 值</param>
    public Length(Double emu) => Emu = emu;

    /// <summary>从 EMU 构造</summary>
    /// <param name="value">EMU 值</param>
    /// <returns>长度值</returns>
    public static Length FromEmu(Double value) => new(value);

    /// <summary>从缇（twips）构造</summary>
    /// <param name="value">缇值</param>
    /// <returns>长度值</returns>
    public static Length FromTwips(Double value) => new(value * EmuPerTwip);

    /// <summary>从磅（point）构造</summary>
    /// <param name="value">磅值</param>
    /// <returns>长度值</returns>
    public static Length FromPoints(Double value) => new(value * EmuPerPoint);

    /// <summary>从厘米构造</summary>
    /// <param name="value">厘米值</param>
    /// <returns>长度值</returns>
    public static Length FromCm(Double value) => new(value * EmuPerCm);

    /// <summary>从毫米构造</summary>
    /// <param name="value">毫米值</param>
    /// <returns>长度值</returns>
    public static Length FromMm(Double value) => new(value * EmuPerCm / 10);

    /// <summary>从英寸构造</summary>
    /// <param name="value">英寸值</param>
    /// <returns>长度值</returns>
    public static Length FromInches(Double value) => new(value * EmuPerInch);

    /// <summary>转换为缇的整数值（OOXML 需要整数缇的属性使用）</summary>
    /// <returns>四舍五入后的缇值</returns>
    public Int32 ToTwipsInt32() => (Int32)Math.Round(Twips);

    /// <summary>转换为 EMU 的整数值（OOXML 需要整数 EMU 的属性使用）</summary>
    /// <returns>四舍五入后的 EMU 值</returns>
    public Int64 ToEmuInt64() => (Int64)Math.Round(Emu);
    #endregion

    #region 解析与输出
    /// <summary>解析带单位后缀的长度文本，如 "2.5cm"、"12pt"、"1in"、"720twip"、"914400emu"</summary>
    /// <param name="text">长度文本（无后缀时按缇处理，与 OOXML 默认一致）</param>
    /// <returns>长度值</returns>
    /// <exception cref="ArgumentNullException">text 为空</exception>
    /// <exception cref="FormatException">文本为空或数字/单位无法识别</exception>
    public static Length Parse(String text)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var s = text.Trim();
        if (s.Length == 0) throw new FormatException(L.T("The length text is empty."));

        // 找出数字结束位置，后面是单位
        var i = 0;
        if (i < s.Length && (s[i] == '-' || s[i] == '+')) i++;
        while (i < s.Length && (Char.IsDigit(s[i]) || s[i] == '.')) i++;
        if (!Double.TryParse(s[..i], out var value))
            throw new FormatException(L.T("Invalid length: {0}", text));

        var unit = s[i..].Trim().ToLowerInvariant();
        return unit switch
        {
            "" or "twip" or "twips" => FromTwips(value),
            "emu" => FromEmu(value),
            "pt" or "point" or "points" => FromPoints(value),
            "mm" => FromMm(value),
            "cm" => FromCm(value),
            "in" or "inch" or "inches" => FromInches(value),
            _ => throw new FormatException(L.T("Unknown length unit: {0}", unit)),
        };
    }

    /// <summary>输出带单位的最短文本表示（优先 cm，其次 pt/twip）</summary>
    /// <returns>如 "2.54cm"</returns>
    public override String ToString() => $"{Math.Round(Cm, 4)}cm";

    /// <summary>判断两个长度是否等值</summary>
    /// <param name="other">另一个长度</param>
    /// <returns>EMU 差值小于 0.5 视为相等</returns>
    public Boolean Equals(Length other) => Math.Abs(Emu - other.Emu) < 0.5;

    /// <summary>判断对象是否为等值长度</summary>
    /// <param name="obj">待比较对象</param>
    /// <returns>等值返回 true</returns>
    public override Boolean Equals(Object? obj) => obj is Length other && Equals(other);

    /// <summary>获取哈希码</summary>
    /// <returns>哈希值</returns>
    public override Int32 GetHashCode() => Emu.GetHashCode();
    #endregion
}
