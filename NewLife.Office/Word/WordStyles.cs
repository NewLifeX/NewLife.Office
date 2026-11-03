using System.Xml;

namespace NewLife.Office.Word;

/// <summary>样式对象 API：按样式 ID 读取/新增/修改/删除 styles.xml 中的样式定义</summary>
/// <remarks>
/// 直接操作 <see cref="Document.StylesXml"/> 的 XML DOM，改动后置 <see cref="Document.StylesModified"/>，
/// 透传模式下 Writer 会用模型样式表覆盖源 styles.xml。
/// <para>样式表为空时以 Writer 的默认样式表为基线（docDefaults + Normal + Heading1~6 + TableGrid）。</para>
/// <example>
/// <code>
/// using var reader = new WordReader("source.docx");
/// var doc = reader.ReadDocument();
/// doc.BeginEdit();
///
/// doc.SetStyleFont("Normal", "微软雅黑", 10.5, color: "333333");
/// doc.SetStyleParagraph("Normal", lineSpacingPct: 150, spaceAfter: 120);
/// doc.AddStyle("QuoteStyle", "引用", basedOn: "Normal");
/// doc.SetStyleFont("QuoteStyle", italic: true, color: "666666");
/// </code>
/// </example>
/// </remarks>
public static class WordStyles
{
    private const String Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #region 读取
    /// <summary>读取全部样式定义</summary>
    /// <param name="doc">文档</param>
    /// <returns>样式快照列表（按 styles.xml 中的顺序）</returns>
    public static List<WordStyleInfo> GetStyles(this Document doc)
    {
        var list = new List<WordStyleInfo>();
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        var xml = Load(doc);
        var ns = Ns(xml);
        foreach (XmlElement el in xml.SelectNodes("//w:style", ns)!)
        {
            var info = ReadStyle(el, ns);
            if (info != null) list.Add(info);
        }
        return list;
    }

    /// <summary>按样式 ID 读取样式定义</summary>
    /// <param name="doc">文档</param>
    /// <param name="styleId">样式 ID（不区分大小写）</param>
    /// <returns>样式快照；不存在返回 null</returns>
    public static WordStyleInfo? GetStyle(this Document doc, String styleId)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (String.IsNullOrEmpty(styleId)) return null;

        var xml = Load(doc);
        var ns = Ns(xml);
        var el = FindStyle(xml, ns, styleId);
        return el == null ? null : ReadStyle(el, ns);
    }
    #endregion

    #region 增删改
    /// <summary>新增样式（已存在同 ID 时返回现有样式，不覆盖）</summary>
    /// <param name="doc">文档</param>
    /// <param name="styleId">样式 ID</param>
    /// <param name="name">显示名（为 null 时用 styleId）</param>
    /// <param name="type">样式类型：paragraph / character / table</param>
    /// <param name="basedOn">基础样式 ID（如 Normal）</param>
    /// <returns>新增或现有样式</returns>
    /// <exception cref="ArgumentNullException">doc 为空</exception>
    /// <exception cref="ArgumentException">styleId 为空</exception>
    public static WordStyleInfo AddStyle(this Document doc, String styleId, String? name = null, String type = "paragraph", String? basedOn = "Normal")
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (String.IsNullOrEmpty(styleId)) throw new ArgumentException(L.T("The style id is empty."), nameof(styleId));

        var xml = Load(doc);
        var ns = Ns(xml);

        var existing = FindStyle(xml, ns, styleId);
        if (existing != null) return ReadStyle(existing, ns)!;

        var root = xml.DocumentElement!;
        var style = xml.CreateElement("w", "style", Wns);
        style.SetAttribute("type", Wns, type);
        style.SetAttribute("styleId", Wns, styleId);

        var nameEl = xml.CreateElement("w", "name", Wns);
        nameEl.SetAttribute("val", Wns, name ?? styleId);
        style.AppendChild(nameEl);

        if (!String.IsNullOrEmpty(basedOn))
        {
            var basedEl = xml.CreateElement("w", "basedOn", Wns);
            basedEl.SetAttribute("val", Wns, basedOn);
            style.AppendChild(basedEl);
        }

        root.AppendChild(style);
        Save(doc, xml);
        return ReadStyle(style, ns)!;
    }

    /// <summary>删除样式</summary>
    /// <param name="doc">文档</param>
    /// <param name="styleId">样式 ID</param>
    /// <returns>删除成功返回 true；样式不存在返回 false</returns>
    public static Boolean RemoveStyle(this Document doc, String styleId)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (String.IsNullOrEmpty(styleId)) return false;

        var xml = Load(doc);
        var ns = Ns(xml);
        var el = FindStyle(xml, ns, styleId);
        if (el == null) return false;

        el.ParentNode?.RemoveChild(el);
        Save(doc, xml);
        return true;
    }

    /// <summary>设置样式字符格式（字体/字号/粗斜体/颜色）</summary>
    /// <param name="doc">文档</param>
    /// <param name="styleId">样式 ID</param>
    /// <param name="fontName">西文字体名（同时设置 hAnsi）</param>
    /// <param name="sizePt">字号（磅）</param>
    /// <param name="bold">是否加粗（null 表示不改）</param>
    /// <param name="italic">是否斜体（null 表示不改）</param>
    /// <param name="color">文字颜色（16 进制，可带 #）</param>
    /// <returns>修改后的样式快照；样式不存在返回 null</returns>
    public static WordStyleInfo? SetStyleFont(this Document doc, String styleId, String? fontName = null, Double? sizePt = null, Boolean? bold = null, Boolean? italic = null, String? color = null)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        var xml = Load(doc);
        var ns = Ns(xml);
        var style = FindStyle(xml, ns, styleId);
        if (style == null) return null;

        var rPr = GetOrCreateChild(xml, style, "rPr", "name", "basedOn", "next", "pPr", "qFormat");
        if (fontName != null)
        {
            var rFonts = GetOrCreateChild(xml, rPr, "rFonts", "rStyle");
            rFonts.SetAttribute("ascii", Wns, fontName);
            rFonts.SetAttribute("hAnsi", Wns, fontName);
        }
        if (sizePt.HasValue)
        {
            var sz = GetOrCreateChild(xml, rPr, "sz", "rStyle", "rFonts", "b", "i", "color", "spacing", "w");
            sz.SetAttribute("val", Wns, ((Int32)Math.Round(sizePt.Value * 2)).ToString());
        }
        if (bold.HasValue) SetSwitch(xml, rPr, "b", bold);
        if (italic.HasValue) SetSwitch(xml, rPr, "i", italic);
        if (color != null)
        {
            var colorEl = GetOrCreateChild(xml, rPr, "color", "rStyle", "rFonts", "b", "i");
            colorEl.SetAttribute("val", Wns, color.TrimStart('#'));
        }

        Save(doc, xml);
        return ReadStyle(style, ns);
    }

    /// <summary>设置样式段落格式（对齐/间距/行距/缩进）</summary>
    /// <param name="doc">文档</param>
    /// <param name="styleId">样式 ID</param>
    /// <param name="alignment">对齐方式：left/center/right/both</param>
    /// <param name="spaceBefore">段前间距（twips）</param>
    /// <param name="spaceAfter">段后间距（twips）</param>
    /// <param name="lineSpacingPct">行距（百分值，100=单倍）</param>
    /// <param name="indentLeft">左缩进（twips）</param>
    /// <param name="firstLineIndent">首行缩进（twips，负值=悬挂）</param>
    /// <returns>修改后的样式快照；样式不存在返回 null</returns>
    public static WordStyleInfo? SetStyleParagraph(this Document doc, String styleId, String? alignment = null,
        Int32? spaceBefore = null, Int32? spaceAfter = null, Int32? lineSpacingPct = null,
        Int32? indentLeft = null, Int32? firstLineIndent = null)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        var xml = Load(doc);
        var ns = Ns(xml);
        var style = FindStyle(xml, ns, styleId);
        if (style == null) return null;

        var pPr = GetOrCreateChild(xml, style, "pPr", "name", "basedOn", "next", "qFormat");
        if (alignment != null)
        {
            var jc = GetOrCreateChild(xml, pPr, "jc", "pStyle", "keepNext", "keepLines", "numPr", "pBdr", "shd", "tabs", "spacing", "ind");
            jc.SetAttribute("val", Wns, alignment);
        }
        if (spaceBefore.HasValue || spaceAfter.HasValue || lineSpacingPct.HasValue)
        {
            var spacing = GetOrCreateChild(xml, pPr, "spacing", "pStyle", "keepNext", "keepLines", "numPr", "pBdr", "shd", "tabs");
            if (spaceBefore.HasValue) spacing.SetAttribute("before", Wns, spaceBefore.Value.ToString());
            if (spaceAfter.HasValue) spacing.SetAttribute("after", Wns, spaceAfter.Value.ToString());
            if (lineSpacingPct.HasValue)
            {
                spacing.SetAttribute("line", Wns, (lineSpacingPct.Value * 240 / 100).ToString());
                spacing.SetAttribute("lineRule", Wns, "auto");
            }
        }
        if (indentLeft.HasValue || firstLineIndent.HasValue)
        {
            var ind = GetOrCreateChild(xml, pPr, "ind", "pStyle", "keepNext", "keepLines", "numPr", "pBdr", "shd", "tabs", "spacing");
            if (indentLeft.HasValue) ind.SetAttribute("left", Wns, indentLeft.Value.ToString());
            if (firstLineIndent.HasValue)
            {
                if (firstLineIndent.Value >= 0) ind.SetAttribute("firstLine", Wns, firstLineIndent.Value.ToString());
                else ind.SetAttribute("hanging", Wns, (-firstLineIndent.Value).ToString());
            }
        }

        Save(doc, xml);
        return ReadStyle(style, ns);
    }
    #endregion

    #region 辅助
    /// <summary>加载样式表 XML（为空时以默认样式表为基线）</summary>
    private static XmlDocument Load(Document doc)
    {
        var xml = new XmlDocument();
        xml.LoadXml(doc.StylesXml ?? WordWriter.BuildDefaultStylesXml());
        return xml;
    }

    /// <summary>写回文档并标记样式已修改</summary>
    private static void Save(Document doc, XmlDocument xml)
    {
        doc.StylesXml = xml.OuterXml;
        doc.StylesModified = true;
        doc.MarkModified(body: false);
    }

    private static XmlNamespaceManager Ns(XmlDocument xml)
    {
        var ns = new XmlNamespaceManager(xml.NameTable);
        ns.AddNamespace("w", Wns);
        return ns;
    }

    private static XmlElement? FindStyle(XmlDocument xml, XmlNamespaceManager ns, String styleId)
    {
        foreach (XmlElement el in xml.SelectNodes("//w:style", ns)!)
        {
            var id = Attr(el, "styleId");
            if (String.Equals(id, styleId, StringComparison.OrdinalIgnoreCase)) return el;
        }
        return null;
    }

    /// <summary>读取 XML 属性，兼容带前缀/不带前缀的命名空间差异</summary>
    /// <param name="el">元素</param>
    /// <param name="localName">不带前缀的本地名（如 styleId、val、ascii）</param>
    /// <returns>属性值；不存在返回 null</returns>
    private static String? Attr(XmlElement? el, String localName)
    {
        if (el == null) return null;

        var value = el.GetAttribute(localName, Wns);
        if (!String.IsNullOrEmpty(value)) return value;
        value = el.GetAttribute("w:" + localName);
        if (!String.IsNullOrEmpty(value)) return value;
        value = el.GetAttribute(localName);
        return String.IsNullOrEmpty(value) ? null : value;
    }

    private static WordStyleInfo? ReadStyle(XmlElement style, XmlNamespaceManager ns)
    {
        var id = Attr(style, "styleId");
        if (String.IsNullOrEmpty(id)) return null;

        var info = new WordStyleInfo
        {
            StyleId = id!,
            Type = Attr(style, "type"),
            Name = Attr(style.SelectSingleNode("w:name", ns) as XmlElement, "val"),
            BasedOn = Attr(style.SelectSingleNode("w:basedOn", ns) as XmlElement, "val"),
        };

        var rPr = style.SelectSingleNode("w:rPr", ns) as XmlElement;
        if (rPr != null)
        {
            info.FontName = Attr(rPr.SelectSingleNode("w:rFonts", ns) as XmlElement, "ascii");
            if (Double.TryParse(Attr(rPr.SelectSingleNode("w:sz", ns) as XmlElement, "val"), out var half)) info.FontSize = half / 2;
            info.Bold = ReadSwitch(rPr.SelectSingleNode("w:b", ns) as XmlElement);
            info.Italic = ReadSwitch(rPr.SelectSingleNode("w:i", ns) as XmlElement);
            info.Color = Attr(rPr.SelectSingleNode("w:color", ns) as XmlElement, "val");
        }

        var pPr = style.SelectSingleNode("w:pPr", ns) as XmlElement;
        if (pPr != null)
        {
            info.Alignment = Attr(pPr.SelectSingleNode("w:jc", ns) as XmlElement, "val");
            var spacing = pPr.SelectSingleNode("w:spacing", ns) as XmlElement;
            if (Int32.TryParse(Attr(spacing, "before"), out var before)) info.SpaceBefore = before;
            if (Int32.TryParse(Attr(spacing, "after"), out var after)) info.SpaceAfter = after;
            if (Int32.TryParse(Attr(spacing, "line"), out var line) && line > 0) info.LineSpacingPct = line * 100 / 240;

            var ind = pPr.SelectSingleNode("w:ind", ns) as XmlElement;
            if (Int32.TryParse(Attr(ind, "left"), out var left)) info.IndentLeft = left;
            if (Int32.TryParse(Attr(ind, "firstLine"), out var first)) info.FirstLineIndent = first;
            else if (Int32.TryParse(Attr(ind, "hanging"), out var hanging)) info.FirstLineIndent = -hanging;
        }
        return info;
    }

    /// <summary>获取或创建子元素（按 schema 顺序：插到最后一个前驱元素之后，无前驱则插到最前）</summary>
    private static XmlElement GetOrCreateChild(XmlDocument xml, XmlElement parent, String tag, params String[] predecessors)
    {
        var existing = parent.SelectSingleNode($"w:{tag}", Ns(xml)) as XmlElement;
        if (existing != null) return existing;

        var child = xml.CreateElement("w", tag, Wns);
        XmlNode? lastPredecessor = null;
        foreach (XmlNode node in parent.ChildNodes)
        {
            if (node is not XmlElement el) continue;
            if (Array.IndexOf(predecessors, el.LocalName) >= 0) lastPredecessor = node;
        }
        if (lastPredecessor != null)
            parent.InsertAfter(child, lastPredecessor);
        else
            parent.PrependChild(child);
        return child;
    }

    private static Boolean? ReadSwitch(XmlElement? el)
    {
        if (el == null) return null;
        var val = Attr(el, "val");
        return val == null || val is "1" or "true" or "on";
    }

    private static void SetSwitch(XmlDocument xml, XmlElement parent, String tag, Boolean? value)
    {
        var existing = parent.SelectSingleNode($"w:{tag}", Ns(xml)) as XmlElement;
        if (value == null)
        {
            existing?.ParentNode?.RemoveChild(existing);
            return;
        }

        var el = existing ?? GetOrCreateChild(xml, parent, tag, "rStyle", "rFonts");
        if (value == true)
            el.RemoveAttribute("val", Wns);
        else
            el.SetAttribute("val", Wns, "0");
    }

    private static String? NullIfEmpty(String? value) => String.IsNullOrEmpty(value) ? null : value;
    #endregion
}
