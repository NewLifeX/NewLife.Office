using SkiaSharp;

namespace NewLife.Office.Rendering.Pdf;

/// <summary>PDF 图形状态</summary>
/// <remarks>维护 PDF 内容流执行过程中的图形栈，对应 PDF 规范 §4.3。</remarks>
internal sealed class PdfGraphicsState
{
    #region 矩阵
    /// <summary>当前变换矩阵 (CTM)，从用户空间到设备空间</summary>
    public SKMatrix CTM { get; set; } = SKMatrix.Identity;

    /// <summary>文本矩阵</summary>
    public SKMatrix TextMatrix { get; set; } = SKMatrix.Identity;

    /// <summary>文本行矩阵（文本行起始位置）</summary>
    public SKMatrix TextLineMatrix { get; set; } = SKMatrix.Identity;

    /// <summary>初始 CTM（页面缩放+翻转 Y）</summary>
    public SKMatrix PageMatrix { get; set; } = SKMatrix.Identity;
    #endregion

    #region 文本状态
    /// <summary>当前字体名称（PDF 字体名，如 "F1"）</summary>
    public String? FontKey { get; set; }

    /// <summary>字体大小（pt）</summary>
    public Single FontSize { get; set; } = 12f;

    /// <summary>字符间距（Tc），默认 0</summary>
    public Single CharSpacing { get; set; }

    /// <summary>单词间距（Tw），默认 0</summary>
    public Single WordSpacing { get; set; }

    /// <summary>水平缩放（Tz），默认 100（百分比）</summary>
    public Single HorizontalScaling { get; set; } = 100f;

    /// <summary>行距（TL），默认 0</summary>
    public Single Leading { get; set; }

    /// <summary>文本渲染模式（Tr），默认 0（填充）</summary>
    public Int32 TextRenderMode { get; set; }

    /// <summary>文本上升（Ts），默认 0</summary>
    public Single TextRise { get; set; }
    #endregion

    #region 颜色
    /// <summary>填充色（缓存为 SKColor）</summary>
    public SKColor FillColor { get; set; } = SKColors.Black;

    /// <summary>描边色</summary>
    public SKColor StrokeColor { get; set; } = SKColors.Black;
    #endregion

    #region 图形状态
    /// <summary>线宽</summary>
    public Single LineWidth { get; set; } = 1f;

    /// <summary>线帽样式</summary>
    public SKStrokeCap LineCap { get; set; } = SKStrokeCap.Butt;

    /// <summary>线连接样式</summary>
    public SKStrokeJoin LineJoin { get; set; } = SKStrokeJoin.Miter;

    /// <summary>斜接限制</summary>
    public Single MiterLimit { get; set; } = 10f;

    /// <summary>虚线模式（null = 实线）</summary>
    public Single[]? DashPattern { get; set; }

    /// <summary>虚线相位</summary>
    public Single DashPhase { get; set; }

    /// <summary>当前路径</summary>
    public SKPath? CurrentPath { get; set; }

    /// <summary>文本位置（用于 BT 块内定位）</summary>
    public Single TextX { get; set; }

    /// <summary>文本位置 Y</summary>
    public Single TextY { get; set; }
    #endregion

    #region 文本 Y 翻转相关
    /// <summary>页面高度，用于 Y 轴翻转（PDF Y 向上 → 设备 Y 向下）</summary>
    public Single PageHeight { get; set; }

    /// <summary>将 PDF 坐标 (Y 向上) 转换为设备坐标 (Y 向下)</summary>
    public Single ToDeviceY(Single pdfY) => PageHeight - pdfY;

    /// <summary>获取文本绘制位置（已翻转 Y）</summary>
    public SKPoint GetTextPosition()
    {
        var tm = TextMatrix;
        return new SKPoint(tm.TransX, ToDeviceY(tm.TransY));
    }
    #endregion

    #region 克隆
    /// <summary>深拷贝当前状态（用于 q 操作符保存）</summary>
    public PdfGraphicsState Clone()
    {
        var clone = (PdfGraphicsState)MemberwiseClone();
        clone.CurrentPath = CurrentPath != null ? new SKPath(CurrentPath) : null;
        if (DashPattern != null)
        {
            clone.DashPattern = new Single[DashPattern.Length];
            Array.Copy(DashPattern, clone.DashPattern, DashPattern.Length);
        }
        return clone;
    }
    #endregion
}
