namespace NewLife.Office;

/// <summary>颜色值对象，跨 Word/PPT/Excel/PDF 四格式通用</summary>
/// <remarks>
/// 轻量值对象，持有 RGBA 四通道分量，并提供 16 进制字符串与分量之间的互转。
/// <example>
/// <code>
/// var red   = OfficeColor.Red;                  // 预定义
/// var blue  = OfficeColor.FromHex("0000FF");    // 从 hex 解析
/// var color = new OfficeColor(0, 128, 255);     // 从分量构造
/// Console.WriteLine(color.Hex);                 // "0080FF"
/// </code>
/// </example>
/// </remarks>
public class OfficeColor
{
    #region 属性
    /// <summary>红色分量（0-255）</summary>
    public Byte R { get; set; }

    /// <summary>绿色分量（0-255）</summary>
    public Byte G { get; set; }

    /// <summary>蓝色分量（0-255）</summary>
    public Byte B { get; set; }

    /// <summary>不透明度分量（0=全透明，255=全不透明），默认 255</summary>
    public Byte A { get; set; } = 255;

    /// <summary>6 位 16 进制 RGB 字符串（无 # 前缀），如 "FF0000"</summary>
    public String Hex => $"{R:X2}{G:X2}{B:X2}";

    /// <summary>8 位 16 进制 ARGB 字符串（无 # 前缀），如 "FFFF0000"</summary>
    public String HexArgb => $"{A:X2}{R:X2}{G:X2}{B:X2}";

    /// <summary>是否全透明</summary>
    public Boolean IsTransparent => A == 0;
    #endregion

    #region 构造
    /// <summary>实例化空白颜色（默认黑色不透明）</summary>
    public OfficeColor() { }

    /// <summary>从 RGB 分量实例化</summary>
    /// <param name="r">红色分量（0-255）</param>
    /// <param name="g">绿色分量（0-255）</param>
    /// <param name="b">蓝色分量（0-255）</param>
    /// <param name="a">不透明度（0-255），默认 255</param>
    public OfficeColor(Byte r, Byte g, Byte b, Byte a = 255)
    {
        R = r;
        G = g;
        B = b;
        A = a;
    }
    #endregion

    #region 工厂方法
    /// <summary>从 16 进制字符串解析颜色</summary>
    /// <param name="hex">支持格式：RRGGBB / #RRGGBB / AARRGGBB / #AARRGGBB</param>
    /// <returns>解析失败时返回 <see cref="Black"/></returns>
    public static OfficeColor FromHex(String? hex)
    {
        if (hex.IsNullOrEmpty()) return Black;
        hex = hex!.TrimStart('#');
        try
        {
            if (hex.Length == 6)
                return new OfficeColor(
                    Convert.ToByte(hex.Substring(0, 2), 16),
                    Convert.ToByte(hex.Substring(2, 2), 16),
                    Convert.ToByte(hex.Substring(4, 2), 16));
            if (hex.Length == 8)
                return new OfficeColor(
                    Convert.ToByte(hex.Substring(2, 2), 16),
                    Convert.ToByte(hex.Substring(4, 2), 16),
                    Convert.ToByte(hex.Substring(6, 2), 16),
                    Convert.ToByte(hex.Substring(0, 2), 16));
        }
        catch { }
        return Black;
    }

    /// <summary>返回同一颜色但指定不透明度的新实例</summary>
    /// <param name="a">不透明度（0-255）</param>
    public OfficeColor WithAlpha(Byte a) => new(R, G, B, a);
    #endregion

    #region 预定义颜色
    /// <summary>黑色 #000000</summary>
    public static readonly OfficeColor Black = new(0, 0, 0);

    /// <summary>白色 #FFFFFF</summary>
    public static readonly OfficeColor White = new(255, 255, 255);

    /// <summary>红色 #FF0000</summary>
    public static readonly OfficeColor Red = new(255, 0, 0);

    /// <summary>绿色 #008000</summary>
    public static readonly OfficeColor Green = new(0, 128, 0);

    /// <summary>蓝色 #0000FF</summary>
    public static readonly OfficeColor Blue = new(0, 0, 255);

    /// <summary>黄色 #FFFF00</summary>
    public static readonly OfficeColor Yellow = new(255, 255, 0);

    /// <summary>橙色 #FFA500</summary>
    public static readonly OfficeColor Orange = new(255, 165, 0);

    /// <summary>灰色 #808080</summary>
    public static readonly OfficeColor Gray = new(128, 128, 128);

    /// <summary>浅灰 #D3D3D3</summary>
    public static readonly OfficeColor LightGray = new(211, 211, 211);

    /// <summary>深灰 #404040</summary>
    public static readonly OfficeColor DarkGray = new(64, 64, 64);

    /// <summary>全透明</summary>
    public static readonly OfficeColor Transparent = new(0, 0, 0, 0);
    #endregion

    /// <summary>返回 Hex 字符串表示</summary>
    public override String ToString() => A == 255 ? $"#{Hex}" : $"#{HexArgb}";
}
