using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace NewLife.Office;

/// <summary>EPUB 电子书写入器，生成 EPUB 3 格式</summary>
public class EpubWriter
{
    #region 写入方法

    /// <summary>将文档写入文件</summary>
    /// <param name="doc">EPUB 文档</param>
    /// <param name="filePath">输出文件路径</param>
    public void Write(EpubDocument doc, String filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
        Write(doc, fs);
    }

    /// <summary>将文档写入流</summary>
    /// <param name="doc">EPUB 文档</param>
    /// <param name="stream">输出流</param>
    public void Write(EpubDocument doc, Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteEpub(doc, zip);
    }

    #endregion

    #region 内部生成

    private static void WriteEpub(EpubDocument doc, ZipArchive zip)
    {
        // mimetype（必须第一个且不压缩）
        var mimetypeEntry = zip.CreateEntry("mimetype", CompressionLevel.NoCompression);
        using (var s = mimetypeEntry.Open())
        using (var sw = new StreamWriter(s, new UTF8Encoding(false)))
            sw.Write("application/epub+zip");

        // META-INF/container.xml
        WriteText(zip, "META-INF/container.xml", BuildContainer());

        // 样式表
        var css = String.IsNullOrEmpty(doc.StyleSheet) ? DefaultCss() : doc.StyleSheet;
        WriteText(zip, "OEBPS/style.css", css);

        // 封面
        var hasCover = doc.Cover != null && doc.Cover.Length > 0;
        if (hasCover)
            WriteBinary(zip, "OEBPS/cover" + GetImageExt(doc.CoverMediaType), doc.Cover!);

        // 封面 XHTML
        if (hasCover)
            WriteText(zip, "OEBPS/cover.xhtml", BuildCoverXhtml(doc));

        // 章节
        var chapters = PadChapters(doc.Chapters);
        foreach (var ch in chapters)
            WriteText(zip, "OEBPS/" + ch.FileName, BuildChapterXhtml(ch));

        // 导航 nav.xhtml (EPUB3)
        WriteText(zip, "OEBPS/nav.xhtml", BuildNav(doc, chapters));

        // OPF
        WriteText(zip, "OEBPS/content.opf", BuildOpf(doc, chapters, hasCover));
    }

    private static List<EpubChapter> PadChapters(List<EpubChapter> chapters)
    {
        var result = new List<EpubChapter>();
        for (var i = 0; i < chapters.Count; i++)
        {
            var ch = chapters[i];
            if (String.IsNullOrEmpty(ch.FileName))
                ch.FileName = "chapter" + (i + 1).ToString("D2") + ".xhtml";
            result.Add(ch);
        }
        return result;
    }

    private static String BuildContainer()
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
             + "<container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\">\n"
             + "  <rootfiles>\n"
             + "    <rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/>\n"
             + "  </rootfiles>\n"
             + "</container>";
    }

    private static String BuildOpf(EpubDocument doc, List<EpubChapter> chapters, Boolean hasCover)
    {
        var id = String.IsNullOrEmpty(doc.Identifier) ? Guid.NewGuid().ToString() : doc.Identifier;
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        sb.AppendLine("<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"uid\">");
        sb.AppendLine("  <metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">");
        sb.AppendLine("    <dc:identifier id=\"uid\">" + XmlEscape(id) + "</dc:identifier>");
        sb.AppendLine("    <dc:title>" + XmlEscape(doc.Title) + "</dc:title>");
        sb.AppendLine("    <dc:creator>" + XmlEscape(doc.Author) + "</dc:creator>");
        sb.AppendLine("    <dc:language>" + XmlEscape(doc.Language) + "</dc:language>");
        if (!String.IsNullOrEmpty(doc.Publisher))
            sb.AppendLine("    <dc:publisher>" + XmlEscape(doc.Publisher) + "</dc:publisher>");
        if (!String.IsNullOrEmpty(doc.Description))
            sb.AppendLine("    <dc:description>" + XmlEscape(doc.Description) + "</dc:description>");
        if (!String.IsNullOrEmpty(doc.PublishDate))
            sb.AppendLine("    <dc:date>" + XmlEscape(doc.PublishDate) + "</dc:date>");
        sb.AppendLine("    <meta property=\"dcterms:modified\">" + DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") + "</meta>");
        if (hasCover)
            sb.AppendLine("    <meta name=\"cover\" content=\"cover-image\"/>");
        sb.AppendLine("  </metadata>");

        sb.AppendLine("  <manifest>");
        sb.AppendLine("    <item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/>");
        sb.AppendLine("    <item id=\"css\" href=\"style.css\" media-type=\"text/css\"/>");
        if (hasCover)
        {
            var ext = GetImageExt(doc.CoverMediaType);
            sb.AppendLine("    <item id=\"cover-image\" href=\"cover" + ext + "\" media-type=\"" + doc.CoverMediaType + "\" properties=\"cover-image\"/>");
            sb.AppendLine("    <item id=\"cover-page\" href=\"cover.xhtml\" media-type=\"application/xhtml+xml\"/>");
        }

        foreach (var ch in chapters)
            sb.AppendLine("    <item id=\"" + ItemId(ch.FileName) + "\" href=\"" + ch.FileName + "\" media-type=\"application/xhtml+xml\"/>");
        sb.AppendLine("  </manifest>");

        sb.AppendLine("  <spine>");
        if (hasCover)
            sb.AppendLine("    <itemref idref=\"cover-page\"/>");
        sb.AppendLine("    <itemref idref=\"nav\"/>");
        foreach (var ch in chapters)
            sb.AppendLine("    <itemref idref=\"" + ItemId(ch.FileName) + "\"/>");
        sb.AppendLine("  </spine>");
        sb.AppendLine("</package>");
        return sb.ToString();
    }

    private static String BuildNav(EpubDocument doc, List<EpubChapter> chapters)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\" lang=\"" + doc.Language + "\">");
        sb.AppendLine("<head><meta charset=\"UTF-8\"/><title>" + XmlEscape(doc.Title) + "</title></head>");
        sb.AppendLine("<body>");
        sb.AppendLine("  <nav epub:type=\"toc\" id=\"toc\">");
        sb.AppendLine("    <h1>目录</h1>");
        sb.AppendLine("    <ol>");
        foreach (var ch in chapters)
            sb.AppendLine("      <li><a href=\"" + ch.FileName + "\">" + XmlEscape(ch.Title) + "</a></li>");
        sb.AppendLine("    </ol>");
        sb.AppendLine("  </nav>");
        sb.AppendLine("</body>");
        sb.Append("</html>");
        return sb.ToString();
    }

    private static String BuildCoverXhtml(EpubDocument doc)
    {
        var ext = GetImageExt(doc.CoverMediaType);
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
             + "<!DOCTYPE html>\n"
             + "<html xmlns=\"http://www.w3.org/1999/xhtml\">\n"
             + "<head><meta charset=\"UTF-8\"/><title>封面</title>\n"
             + "<style>body{margin:0;padding:0;text-align:center;} img{max-width:100%;}</style></head>\n"
             + "<body><img src=\"cover" + ext + "\" alt=\"封面\"/></body>\n"
             + "</html>";
    }

    private static String BuildChapterXhtml(EpubChapter ch)
    {
        // 如果内容已经是完整 XHTML 直接使用，否则包装
        if (ch.Content.TrimStart().StartsWith("<?xml") || ch.Content.TrimStart().StartsWith("<!DOCTYPE"))
            return ch.Content;

        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
             + "<!DOCTYPE html>\n"
             + "<html xmlns=\"http://www.w3.org/1999/xhtml\">\n"
             + "<head><meta charset=\"UTF-8\"/><title>" + XmlEscape(ch.Title) + "</title>\n"
             + "<link rel=\"stylesheet\" href=\"style.css\" type=\"text/css\"/></head>\n"
             + "<body>\n"
             + "<h1>" + XmlEscape(ch.Title) + "</h1>\n"
             + ch.Content + "\n"
             + "</body>\n"
             + "</html>";
    }

    private static String DefaultCss() =>
        "body { font-family: serif; font-size: 1em; line-height: 1.6; margin: 1em; }\n"
      + "h1, h2, h3 { font-weight: bold; }\n"
      + "p { text-indent: 2em; margin: 0; }\n";

    private static void WriteText(ZipArchive zip, String path, String content)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var s = entry.Open();
        using var sw = new StreamWriter(s, new UTF8Encoding(false));
        sw.Write(content);
    }

    private static void WriteBinary(ZipArchive zip, String path, Byte[] data)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var s = entry.Open();
        s.Write(data, 0, data.Length);
    }

    private static String GetImageExt(String mediaType)
    {
        if (mediaType.Contains("png")) return ".png";
        if (mediaType.Contains("gif")) return ".gif";
        if (mediaType.Contains("svg")) return ".svg";
        return ".jpg";
    }

    private static String ItemId(String fileName) =>
        "item-" + Path.GetFileNameWithoutExtension(fileName).Replace(" ", "-");

    private static String XmlEscape(String s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    #endregion
}
