using System.IO.Compression;
using System.Xml;

var path = @"D:\X\NewLife.Office\Bin\A4工业计算机_v2.0.docx";
using var za = ZipFile.OpenRead(path);

Console.WriteLine("=== ZIP entries ===");
foreach (var e in za.Entries) Console.WriteLine("  " + e.FullName);

var docEntry = za.GetEntry("word/document.xml")!;
var xml = new XmlDocument();
using (var s = docEntry.Open()) xml.Load(s);
var ns = new XmlNamespaceManager(xml.NameTable);
ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
ns.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
ns.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
ns.AddNamespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
ns.AddNamespace("v", "urn:schemas-microsoft-com:vml");
ns.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

Console.WriteLine("\n=== document.xml statistics ===");
Console.WriteLine("w:p: " + xml.SelectNodes("//w:p", ns)!.Count);
Console.WriteLine("w:tbl: " + xml.SelectNodes("//w:tbl", ns)!.Count);
Console.WriteLine("w:drawing: " + xml.SelectNodes("//w:drawing", ns)!.Count);
Console.WriteLine("wp:inline: " + xml.SelectNodes("//wp:inline", ns)!.Count);
Console.WriteLine("a:blip: " + xml.SelectNodes("//a:blip", ns)!.Count);
Console.WriteLine("w:sectPr: " + xml.SelectNodes("//w:sectPr", ns)!.Count);
Console.WriteLine("w:bookmarkStart: " + xml.SelectNodes("//w:bookmarkStart", ns)!.Count);
Console.WriteLine("w:pPr/w:shd (para bg): " + xml.SelectNodes("//w:pPr/w:shd", ns)!.Count);
Console.WriteLine("w:rPr/w:b (bold): " + xml.SelectNodes("//w:rPr/w:b", ns)!.Count);
Console.WriteLine("w:rPr/w:color: " + xml.SelectNodes("//w:rPr/w:color", ns)!.Count);
Console.WriteLine("w:rPr/w:sz: " + xml.SelectNodes("//w:rPr/w:sz", ns)!.Count);
Console.WriteLine("w:rPr/w:rFonts: " + xml.SelectNodes("//w:rPr/w:rFonts", ns)!.Count);
Console.WriteLine("w:rPr/w:u: " + xml.SelectNodes("//w:rPr/w:u", ns)!.Count);
Console.WriteLine("w:pStyle: " + xml.SelectNodes("//w:pStyle", ns)!.Count);
Console.WriteLine("w:hyperlink: " + xml.SelectNodes("//w:hyperlink", ns)!.Count);
Console.WriteLine("w:tblGrid: " + xml.SelectNodes("//w:tblGrid", ns)!.Count);
Console.WriteLine("w:tcBorders: " + xml.SelectNodes("//w:tcBorders", ns)!.Count);

// 节属性详情
var sectPr = xml.SelectSingleNode("//w:sectPr", ns) as XmlElement;
if (sectPr != null) {
    Console.WriteLine("\n=== sectPr details ===");
    foreach (XmlNode c in sectPr.ChildNodes)
        if (c is XmlElement el)
            Console.WriteLine("  " + el.Name + ": " + (el.OuterXml.Length > 150 ? el.OuterXml[..150] + "..." : el.OuterXml));
}

// 检查页眉
var hdrNames = new[] { "word/header2.xml", "word/header1.xml", "word/header3.xml" };
foreach (var hn in hdrNames) {
    var hdrEntry = za.GetEntry(hn);
    if (hdrEntry == null) continue;
    var hxml = new XmlDocument();
    using (var s = hdrEntry.Open()) hxml.Load(s);
    Console.WriteLine($"\n=== {hn} ===");
    Console.WriteLine("w:p: " + hxml.SelectNodes("//w:p", ns)!.Count);
    Console.WriteLine("w:tbl: " + hxml.SelectNodes("//w:tbl", ns)!.Count);
    Console.WriteLine("w:drawing: " + hxml.SelectNodes("//w:drawing", ns)!.Count);
    Console.WriteLine("a:blip: " + hxml.SelectNodes("//a:blip", ns)!.Count);
    Console.WriteLine("v:imagedata: " + hxml.SelectNodes("//v:imagedata", ns)!.Count);
}

// 页脚
foreach (var fn in new[] { "word/footer2.xml", "word/footer1.xml" }) {
    var ftrEntry = za.GetEntry(fn);
    if (ftrEntry == null) continue;
    var fxml = new XmlDocument();
    using (var s = ftrEntry.Open()) fxml.Load(s);
    Console.WriteLine($"\n=== {fn} ===");
    Console.WriteLine("w:p: " + fxml.SelectNodes("//w:p", ns)!.Count);
    Console.WriteLine("a:blip: " + fxml.SelectNodes("//a:blip", ns)!.Count);
}

Console.WriteLine("\n=== Done ===");
