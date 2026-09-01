using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NewLife.Office.Pdf;
using SkiaSharp;

namespace NewLife.Office.Rendering.Pdf;

/// <summary>PDF 页面渲染器（单页）</summary>
/// <remarks>
/// 核心渲染管线：读取 PDF 页面内容流 → 解析操作符 → 执行图形状态机 → SkiaSharp Canvas 绘制。
/// 支持文本（Tj/TJ）、路径（m/l/re/S/f）、图片（Do/Image XObject）、颜色（G/g/RG/rg）等常用操作符。
/// </remarks>
internal sealed class PdfPageRenderer
{
    #region 字段
    private Byte[] _data = [];
    private PdfXRefTable _xref = null!;
    private PdfGraphicsState _state = null!;
    private SKCanvas _canvas = null!;
    private PdfFontMapper _fontMapper = null!;
    private readonly Encoding _latin1 = Encoding.GetEncoding(28591);
    #endregion

    #region 渲染入口
    /// <summary>渲染 PDF 单页到指定画布</summary>
    /// <param name="data">PDF 原始字节</param>
    /// <param name="xref">交叉引用表</param>
    /// <param name="pageIndex">页面索引（0-based）</param>
    /// <param name="canvas">SkiaSharp 画布</param>
    /// <param name="width">画布宽度</param>
    /// <param name="height">画布高度</param>
    /// <param name="dpi">分辨率（DPI）</param>
    public void Render(Byte[] data, PdfXRefTable xref, Int32 pageIndex, SKCanvas canvas, Int32 width, Int32 height, Int32 dpi)
    {
        _data = data;
        _xref = xref;
        _canvas = canvas;
        _fontMapper = new PdfFontMapper();

        _state = new PdfGraphicsState
        {
            PageHeight = height,
            PageMatrix = SKMatrix.CreateScaleTranslation(dpi / 72f, dpi / 72f, 0, 0),
        };

        canvas.Clear(SKColors.White);

        // 获取页面对象号并读取内容流
        var pageObjNums = GetPageObjectNumbers();
        if (pageIndex >= pageObjNums.Count) return;

        var objNum = pageObjNums[pageIndex];
        var pageObj = PdfObjectParser.ReadObject(_data, _xref, objNum);
        if (pageObj is not DictObj pageDictObj) return;

        var contentsData = GetPageContents(pageDictObj.Value);
        if (contentsData != null && contentsData.Length > 0)
        {
            var contentStream = new PdfContentStream(contentsData);
            ExecuteOperators(contentStream.Operators);
        }
    }
    #endregion

    #region 页面对象号与内容流
    private List<Int32> GetPageObjectNumbers()
    {
        var result = new List<Int32>();
        var text = _latin1.GetString(_data);
        var pos = 0;
        while (true)
        {
            var found = text.IndexOf("/Type", pos, StringComparison.Ordinal);
            if (found < 0) break;

            var afterType = text.IndexOf("/Page", found, StringComparison.Ordinal);
            if (afterType >= 0 && afterType - found < 80)
            {
                var objStart = text.LastIndexOf("obj", found);
                if (objStart >= 0)
                {
                    // 对象定义形如 "5 0 obj"，从 obj 前回溯对象号；obj 前可能有空格（如 "5 0 obj"）
                    var numStart = objStart - 1;
                    // 跳过 obj 前的空白
                    while (numStart >= 0 && Char.IsWhiteSpace(text[numStart])) numStart--;
                    // 回退数字（对象号）
                    while (numStart >= 0 && Char.IsDigit(text[numStart])) numStart--;
                    numStart++;
                    // 防御：obj 前必须存在对象号（numStart < objStart），否则 Substring 负长度崩溃
                    if (objStart > numStart)
                    {
                        var numStr = text.Substring(numStart, objStart - numStart - 1).Trim();
                        if (Int32.TryParse(numStr, out var num) && !result.Contains(num))
                            result.Add(num);
                    }
                }
                pos = afterType + 5;
            }
            else
            {
                pos = found + 5;
            }
        }
        return result;
    }

    private Byte[]? GetPageContents(PdfDict pageDict)
    {
        if (!pageDict.TryGetValue("Contents", out var contentsVal)) return null;
        if (contentsVal is Ref contentsRef)
            return GetStreamData(contentsRef.ObjNum);
        if (contentsVal is PdfArray contentsArr)
        {
            var result = new List<Byte>();
            foreach (var item in contentsArr.Items)
            {
                if (item is Ref r)
                {
                    var part = GetStreamData(r.ObjNum);
                    if (part != null) result.AddRange(part);
                }
            }
            return result.Count > 0 ? result.ToArray() : null;
        }
        return null;
    }

    private Byte[]? GetStreamData(Int32 objNum)
    {
        var obj = PdfObjectParser.ReadObject(_data, _xref, objNum);
        if (obj is DictObj dictObj && dictObj.Value.TryGetValue("_stream", out var streamVal))
        {
            if (streamVal is PdfString ps)
                return System.Text.Encoding.GetEncoding(28591).GetBytes(ps.Value);
            var rawObj = (Object)streamVal;
            if (rawObj is Byte[] sb)
                return sb;
        }
        return null;
    }
    #endregion

    #region 操作符执行
    private void ExecuteOperators(List<Operator> operators)
    {
        foreach (var op in operators)
        {
            ExecuteOperator(op);
        }
    }

    private void ExecuteOperator(Operator op)
    {
        switch (op.Name)
        {
            // ──── 图形状态 ────
            case "q": SaveState(); break;
            case "Q": RestoreState(); break;
            case "cm": Op_cm(op); break;
            case "w": Op_w(op); break;
            case "J": Op_J(op); break;
            case "j": Op_j(op); break;
            case "M": Op_M(op); break;
            case "d": Op_d(op); break;
            case "gs": break; // 扩展图形状态，暂跳过

            // ──── 颜色 ────
            case "G": Op_G(op); break;
            case "g": Op_g(op); break;
            case "RG": Op_RG(op); break;
            case "rg": Op_rg(op); break;
            case "K": Op_K(op); break;
            case "k": Op_k(op); break;

            // ──── 文本 ────
            case "BT": break; // 开始文本对象
            case "ET": break; // 结束文本对象
            case "Tf": Op_Tf(op); break;
            case "Td": Op_Td(op); break;
            case "TD": Op_TD(op); break;
            case "Tm": Op_Tm(op); break;
            case "T*": Op_Tstar(); break;
            case "Tj": Op_Tj(op); break;
            case "TJ": Op_TJ(op); break;
            case "Tc": Op_Tc(op); break;
            case "Tw": Op_Tw(op); break;
            case "Tz": Op_Tz(op); break;
            case "TL": Op_TL(op); break;
            case "Ts": Op_Ts(op); break;
            case "Tr": Op_Tr(op); break;
            case "'": Op_Tick(); break;

            // ──── 路径构造 ────
            case "m": Op_m(op); break;
            case "l": Op_l(op); break;
            case "re": Op_re(op); break;
            case "h": break; // 闭合路径

            // ──── 路径绘制 ────
            case "S": Op_S(); break;
            case "s": Op_s(); break;
            case "f": case "F": Op_f(); break;
            case "f*": Op_f(); break; // 偶奇规则填充，简化处理
            case "B": Op_B(); break;
            case "b": Op_b(); break;
            case "B*": Op_B(); break;
            case "n": Op_n(); break;

            // ──── XObject ────
            // case "Do": 暂不支持（需解析 XObject 资源字典）

            // 忽略未知操作符
            default: break;
        }
    }
    #endregion

    #region 图形状态栈
    private readonly Stack<PdfGraphicsState> _stateStack = new();

    private void SaveState()
    {
        _stateStack.Push(_state.Clone());
    }

    private void RestoreState()
    {
        if (_stateStack.Count > 0)
            _state = _stateStack.Pop();
    }
    #endregion

    #region 操作符实现 - 图形
    private void Op_cm(Operator op)
    {
        if (op.Operands.Count < 6) return;
        var a = ToFloat(op.Operands[0]);
        var b = ToFloat(op.Operands[1]);
        var c = ToFloat(op.Operands[2]);
        var d = ToFloat(op.Operands[3]);
        var e = ToFloat(op.Operands[4]);
        var f = ToFloat(op.Operands[5]);

        var m = new SKMatrix { ScaleX = a, SkewX = c, TransX = e, SkewY = b, ScaleY = d, TransY = f, Persp2 = 1 };
        _state.CTM = _state.CTM.PreConcat(m);
    }

    private void Op_w(Operator op) { if (op.Operands.Count > 0) _state.LineWidth = ToFloat(op.Operands[0]); }
    private void Op_J(Operator op)
    {
        if (op.Operands.Count > 0)
        {
            _state.LineCap = (Int32)ToFloat(op.Operands[0]) switch
            {
                1 => SKStrokeCap.Round,
                2 => SKStrokeCap.Square,
                _ => SKStrokeCap.Butt,
            };
        }
    }
    private void Op_j(Operator op)
    {
        if (op.Operands.Count > 0)
        {
            _state.LineJoin = (Int32)ToFloat(op.Operands[0]) switch
            {
                1 => SKStrokeJoin.Round,
                2 => SKStrokeJoin.Bevel,
                _ => SKStrokeJoin.Miter,
            };
        }
    }
    private void Op_M(Operator op) { if (op.Operands.Count > 0) _state.MiterLimit = ToFloat(op.Operands[0]); }

    private void Op_d(Operator op)
    {
        if (op.Operands.Count < 2) return;
        var arrIdx = 0;
        // 操作数结构：[[dashArray] dashPhase]
        var dashCount = op.Operands.Count - 1;
        _state.DashPattern = new Single[dashCount];
        for (var i = 0; i < dashCount; i++)
            _state.DashPattern[i] = ToFloat(op.Operands[i]);
        _state.DashPhase = ToFloat(op.Operands[^1]);
        arrIdx++;
    }
    #endregion

    #region 操作符实现 - 颜色
    private void Op_G(Operator op)  { var g = ToByte(op.Operands[0]); _state.StrokeColor = new SKColor(g, g, g); }
    private void Op_g(Operator op)  { var g = ToByte(op.Operands[0]); _state.FillColor = new SKColor(g, g, g); }
    private void Op_RG(Operator op) { if (op.Operands.Count >= 3) _state.StrokeColor = new SKColor(ToByte(op.Operands[0]), ToByte(op.Operands[1]), ToByte(op.Operands[2])); }
    private void Op_rg(Operator op) { if (op.Operands.Count >= 3) _state.FillColor = new SKColor(ToByte(op.Operands[0]), ToByte(op.Operands[1]), ToByte(op.Operands[2])); }
    private void Op_K(Operator op) { if (op.Operands.Count >= 4) { var r = (Byte)(255 * (1 - ToFloat(op.Operands[0])) * (1 - ToFloat(op.Operands[3]))); var g = (Byte)(255 * (1 - ToFloat(op.Operands[1])) * (1 - ToFloat(op.Operands[3]))); var b = (Byte)(255 * (1 - ToFloat(op.Operands[2])) * (1 - ToFloat(op.Operands[3]))); _state.StrokeColor = new SKColor(r, g, b); } }
    private void Op_k(Operator op) { if (op.Operands.Count >= 4) { var r = (Byte)(255 * (1 - ToFloat(op.Operands[0])) * (1 - ToFloat(op.Operands[3]))); var g = (Byte)(255 * (1 - ToFloat(op.Operands[1])) * (1 - ToFloat(op.Operands[3]))); var b = (Byte)(255 * (1 - ToFloat(op.Operands[2])) * (1 - ToFloat(op.Operands[3]))); _state.FillColor = new SKColor(r, g, b); } }
    #endregion

    #region 操作符实现 - 文本
    private void Op_Tf(Operator op)
    {
        if (op.Operands.Count < 2) return;
        _state.FontKey = op.Operands[0] is String s ? s : op.Operands[0].ToString();
        _state.FontSize = ToFloat(op.Operands[1]);
    }

    private void Op_Td(Operator op)
    {
        if (op.Operands.Count < 2) return;
        var tx = ToFloat(op.Operands[0]);
        var ty = ToFloat(op.Operands[1]);
        var m = SKMatrix.CreateTranslation(tx, ty);
        _state.TextLineMatrix = _state.TextLineMatrix.PreConcat(m);
        _state.TextMatrix = _state.TextLineMatrix;
    }

    private void Op_TD(Operator op)
    {
        if (op.Operands.Count < 2) return;
        _state.Leading = -ToFloat(op.Operands[1]);
        Op_Td(op);
    }

    private void Op_Tm(Operator op)
    {
        if (op.Operands.Count < 6) return;
        var a = ToFloat(op.Operands[0]);
        var b = ToFloat(op.Operands[1]);
        var c = ToFloat(op.Operands[2]);
        var d = ToFloat(op.Operands[3]);
        var e = ToFloat(op.Operands[4]);
        var f = ToFloat(op.Operands[5]);
        _state.TextMatrix = new SKMatrix { ScaleX = a, SkewX = c, TransX = e, SkewY = b, ScaleY = d, TransY = f, Persp2 = 1 };
        _state.TextLineMatrix = _state.TextMatrix;
    }

    private void Op_Tstar()
    {
        var m = SKMatrix.CreateTranslation(0, -_state.Leading);
        _state.TextLineMatrix = _state.TextLineMatrix.PreConcat(m);
        _state.TextMatrix = _state.TextLineMatrix;
    }

    private void Op_Tj(Operator op)
    {
        if (op.Operands.Count < 1) return;
        var text = GetStringValue(op.Operands[0]);
        if (String.IsNullOrEmpty(text)) return;
        DrawText(text);
    }

    private void Op_TJ(Operator op)
    {
        // TJ 的操作数是数组：[ str num str num ... ]
        if (op.Operands.Count == 0) return;

        foreach (var operand in op.Operands)
        {
            if (operand is Object[] arr)
            {
                foreach (var item in arr)
                {
                    ProcessTJElement(item);
                }
            }
            else
            {
                ProcessTJElement(operand);
            }
        }
    }

    private void ProcessTJElement(Object element)
    {
        if (element is String str)
        {
            DrawText(str);
        }
        else
        {
            // PdfTextOperand（internal 类型）或数值调整
            var type = element.GetType();
            if (type.Name == "PdfTextOperand")
            {
                var valProp = type.GetProperty("Value");
                if (valProp != null)
                {
                    var inner = valProp.GetValue(element);
                    if (inner != null) ProcessTJElement(inner);
                }
            }
            else
            {
                var adjust = ToFloat(element);
                if (adjust < 0)
                {
                    var m = SKMatrix.CreateTranslation(-adjust * _state.FontSize / 1000f, 0);
                    _state.TextMatrix = _state.TextMatrix.PreConcat(m);
                }
            }
        }
    }

    private void Op_Tc(Operator op) { if (op.Operands.Count > 0) _state.CharSpacing = ToFloat(op.Operands[0]); }
    private void Op_Tw(Operator op) { if (op.Operands.Count > 0) _state.WordSpacing = ToFloat(op.Operands[0]); }
    private void Op_Tz(Operator op) { if (op.Operands.Count > 0) _state.HorizontalScaling = ToFloat(op.Operands[0]); }
    private void Op_TL(Operator op) { if (op.Operands.Count > 0) _state.Leading = ToFloat(op.Operands[0]); }
    private void Op_Ts(Operator op) { if (op.Operands.Count > 0) _state.TextRise = ToFloat(op.Operands[0]); }
    private void Op_Tr(Operator op) { if (op.Operands.Count > 0) _state.TextRenderMode = (Int32)ToFloat(op.Operands[0]); }
    private void Op_Tick() { Op_Tstar(); DrawText("\n"); }
    #endregion

    #region 操作符实现 - 路径
    private void Op_m(Operator op)
    {
        if (op.Operands.Count < 2) return;
        _state.CurrentPath ??= new SKPath();
        var x = ToFloat(op.Operands[0]);
        var y = _state.ToDeviceY(ToFloat(op.Operands[1]));
        _state.CurrentPath.MoveTo(x, y);
    }

    private void Op_l(Operator op)
    {
        if (op.Operands.Count < 2) return;
        _state.CurrentPath ??= new SKPath();
        var x = ToFloat(op.Operands[0]);
        var y = _state.ToDeviceY(ToFloat(op.Operands[1]));
        _state.CurrentPath.LineTo(x, y);
    }

    private void Op_re(Operator op)
    {
        if (op.Operands.Count < 4) return;
        var x = ToFloat(op.Operands[0]);
        var y = _state.ToDeviceY(ToFloat(op.Operands[1]));
        var w = ToFloat(op.Operands[2]);
        var h = ToFloat(op.Operands[3]);
        _state.CurrentPath ??= new SKPath();
        _state.CurrentPath.AddRect(new SKRect(x, y - h, x + w, y));
    }

    private void Op_S() { StrokePath(); }
    private void Op_s() { _state.CurrentPath?.Close(); StrokePath(); }
    private void Op_f() { FillPath(); }
    private void Op_B() { FillPath(); StrokePath(); }
    private void Op_b() { _state.CurrentPath?.Close(); FillPath(); StrokePath(); }
    private void Op_n() { _state.CurrentPath = null; }
    #endregion

    #region 绘制辅助
    private void DrawText(String text)
    {
        if (String.IsNullOrEmpty(_state.FontKey)) return;

        var typeface = _fontMapper.GetTypeface(_state.FontKey!, _state.FontKey, null);
        var pos = _state.GetTextPosition();

        using var font = new SKFont(typeface, _state.FontSize, 1, 0);
        using var paint = new SKPaint
        {
            IsAntialias = true,
            Color = _state.FillColor,
        };

        _canvas.DrawText(text, pos.X, pos.Y, font, paint);

        var advance = font.MeasureText(text, paint) + _state.CharSpacing * _state.FontSize / 1000f;
        var m = SKMatrix.CreateTranslation(advance, 0);
        _state.TextMatrix = _state.TextMatrix.PreConcat(m);
    }

    private void StrokePath()
    {
        var path = _state.CurrentPath;
        if (path == null) return;

        using var paint = new SKPaint
        {
            Style = SKPaintStyle.Stroke, Color = _state.StrokeColor,
            StrokeWidth = _state.LineWidth, StrokeCap = _state.LineCap,
            StrokeJoin = _state.LineJoin, StrokeMiter = _state.MiterLimit, IsAntialias = true,
        };
        if (_state.DashPattern != null)
            paint.PathEffect = SKPathEffect.CreateDash(_state.DashPattern, _state.DashPhase);

        _canvas.DrawPath(path, paint);
        _state.CurrentPath = null;
    }

    private void FillPath()
    {
        var path = _state.CurrentPath;
        if (path == null) return;

        using var paint = new SKPaint { Style = SKPaintStyle.Fill, Color = _state.FillColor, IsAntialias = true };
        _canvas.DrawPath(path, paint);
    }
    #endregion

    #region 工具方法
    private static Single ToFloat(Object value) => value switch
    {
        Single f => f, Double d => (Single)d, Int32 i => i, Int64 l => l,
        String s => Single.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var r) ? r : 0,
        _ => 0f,
    };

    private static Byte ToByte(Object value) { var f = ToFloat(value); var clamped = f * (f <= 1 ? 255 : 1); if (clamped < 0) clamped = 0; if (clamped > 255) clamped = 255; return (Byte)clamped; }

    private static String GetStringValue(Object value)
    {
        if (value is String s) return s;
        // PdfTextOperand 是 internal 类型，通过反射获取 Value
        var type = value.GetType();
        if (type.Name == "PdfTextOperand")
        {
            var valProp = type.GetProperty("Value");
            if (valProp != null)
            {
                var inner = valProp.GetValue(value);
                if (inner != null) return GetStringValue(inner);
            }
        }
        return value.ToString() ?? String.Empty;
    }
    #endregion
}
