using System.Globalization;
using System.Runtime.CompilerServices;
using NewLife;
using NewLife.Office;
using NewLife.Office.Ole2;
using Xunit;

namespace XUnitTest.Common;

/// <summary>多语言本地化测试：31 种内置语言表与英文回退</summary>
[Collection("Localization")]
public class LocalizationTests
{
    static LocalizationTests() => RuntimeHelpers.RunModuleConstructor(typeof(OfficeFactory).Module.ModuleHandle);

    /// <summary>本库内置的 31 种语言代码</summary>
    public static readonly String[] Cultures =
    [
        "zh-Hans", "zh-Hant", "ja", "ko", "vi", "th", "id", "ms", "hi", "bn",
        "ar", "he", "fa", "de", "fr", "es", "pt", "it", "nl", "pl",
        "cs", "sv", "da", "nb", "fi", "el", "ro", "hu", "ru", "uk", "tr",
    ];

    /// <summary>全部已登记词条（62 条运行时消息 + 61 条文档输出标签）</summary>
    public static readonly String[] AllKeys =
    [
        "Workbook stream not found; the file may not be a valid xls file.",
        "BIFF8 Globals BOF record not found; the file may be corrupted.",
        "The range must be in A1:F1 format.",
        "The rich-text segments cannot be empty.",
        "No output location was specified.",
        "Source data has not been set; call SetSourceData() first.",
        "The stream content cannot be recognized or the format is not supported.",
        "Rendering PDF pages to images requires a rendering library, which this version does not support.",
        "The signature field is missing the /Contents placeholder.",
        "The /Contents placeholder is malformed.",
        "The certificate does not contain a usable RSA private key.",
        "PowerPoint Document stream not found; the file may not be a valid .ppt file.",
        "File not found: {0}",
        "The stream content cannot be recognized or the format is not supported: {0}",
        "The text is too long ({0} bytes); the maximum is 106 bytes. Please shorten the text.",
        "Signature field '{0}' was not found.",
        "Signature length {0} exceeds the reserved {1} characters.",
        "Slide index {0} is out of range (total {1} slides).",
        "The number of persisted objects {0} exceeds the per-block persistence directory limit of 4095.",
        "Image index {0} is out of range (total {1} images).",
        "Slide {0} does not exist in the source file.",
        "The encrypted PPTX requires an open password.",
        "Slide {0} does not exist.",
        "WordDocument stream not found; the file may not be a valid .doc file.",
        "Unsupported document format: wIdent = 0x{0:X4}. Only Word 97-2003 binary format is supported.",
        "The EncryptionInfo stream is missing from the CFB container.",
        "The EncryptedPackage stream is missing from the CFB container.",
        "The EncryptedPackage data is too short.",
        "The password is incorrect or the document is corrupted.",
        "The EncryptionInfo data is too short.",
        "The XML declaration is missing from EncryptionInfo.",
        "saltValue is missing from EncryptionInfo.",
        "The saltValue format in EncryptionInfo is invalid.",
        "Invalid docx file: word/document.xml is missing.",
        "Unsupported file format: {0}",
        "Scaling failed.",
        "Unable to decode the image data.",
        "Invalid image dimensions.",
        "Failed to generate the QR code.",
        "Page index {0} is out of range [0, {1}].",
        "Failed to convert Word to PDF.",
        "Error: {0}",
        "Usage: convert <input> [output]",
        "Unable to convert: {0}",
        "Converted: {0} → {1}",
        "Usage: doc2docx <input.doc> <output.docx>",
        "Usage: info <input>",
        "File: {0}",
        "Format: {0}",
        "Unknown",
        "Size: {0} bytes",
        "NewLife.Office CLI tool",
        "Usage:",
        "  convert <input> [output]   Convert office documents to Markdown (output defaults to stdout)",
        "  doc2docx <in.doc> <out.docx>   Convert doc (97-2003) to docx (W20)",
        "  info <input>               Show file format information",
        "Supported formats: docx/xlsx/pptx/pdf/md/ods/rtf/epub, etc.",
        "Unable to find worksheet",
        "Stream must be seekable.",
        "Not a valid CFB file: magic number mismatch.",
        "Unexpected end of CFB stream.",
        "Rendering to images requires a rendering library such as SkiaSharp, which this version does not support. Convert to PDF first, then use PdfDocument.RenderToImages.",
        "Start: {0:yyyy-MM-dd HH:mm}",
        "End: {0:yyyy-MM-dd HH:mm}",
        "Location: {0}",
        "Description: {0}",
        "Status: {0}",
        "[To-do] {0}",
        "Due: {0:yyyy-MM-dd HH:mm}",
        "Untitled event",
        "- **Start**: {0:yyyy-MM-dd HH:mm}",
        "- **End**: {0:yyyy-MM-dd HH:mm}",
        "- **Location**: {0}",
        "- **Description**: {0}",
        "- **Status**: {0}",
        "- **Organizer**: {0}",
        "- **Attendees**: {0}",
        "Untitled to-do",
        "- **Due**: {0:yyyy-MM-dd HH:mm}",
        "- **Progress**: {0}%",
        "Contents",
        "Cover",
        "Slicer_{0}",
        "(Blank)",
        "From: {0}",
        "To: {0}",
        "Cc: {0}",
        "Subject: {0}",
        "Date: {0:yyyy-MM-dd HH:mm:ss}",
        "- **From**: {0}",
        "- **To**: {0}",
        "- **Cc**: {0}",
        "- **Subject**: {0}",
        "- **Date**: {0:yyyy-MM-dd HH:mm:ss}",
        "Page {0}",
        "Slide {0}",
        "--- Slide break ---",
        "Default section",
        "Organization: {0}",
        "Job title: {0}",
        "Birthday: {0:yyyy-MM-dd}",
        "Phone: {0}",
        "Email: {0}",
        "Address: {0}",
        "Website: {0}",
        "Note: {0}",
        "Unknown contact",
        "- **Organization**: {0}",
        "- **Job title**: {0}",
        "- **Birthday**: {0:yyyy-MM-dd}",
        "- **Phone**: {0}",
        "- **Email**: {0}",
        "- **Address**: {0}",
        "- **Website**: {0}",
        "- **Note**: {0}",
        "Category",
        "Click or tap to enter text",
        "Click or tap to enter a date",
        "Choose an item",
        "Choose or enter an item",
        "(author {0})",
        "Footnotes {0}: {1}",
        "Endnotes {0}: {1}",
        "Image {0}",
        "Image",
        "Table({0} rows)",
    ];

    [Fact(DisplayName = "31 种语言已翻译全部词条")]
    public void AllKeysTranslated()
    {
        var saved = Localization.Culture;
        try
        {
            foreach (var name in Cultures)
            {
                Localization.Culture = CultureInfo.GetCultureInfo(name);
                foreach (var key in AllKeys)
                {
                    var text = Localization.Translate(key);
                    Assert.False(String.IsNullOrEmpty(text), $"{name} 缺少译文：{key}");
                    Assert.False(text == key, $"{name} 缺少译文：{key}");
                }
            }
        }
        finally
        {
            Localization.Culture = saved;
        }
    }

    [Fact(DisplayName = "程序集标题与描述按当前语言翻译")]
    public void AssemblyMetadataIsLocalized()
    {
        const String title = "Office Automation Library";
        const String description = "Office automation library supporting read/write, template filling and format conversion for Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB and XPS, with zero external dependencies and an MIT license.";
        const String renderingTitle = "Office Document Rendering Extension";
        const String renderingDescription = "Rendering extension for NewLife.Office, providing document page to image rendering, image processing, watermark overlay and QR code generation, powered by SkiaSharp for cross-platform 2D graphics.";
        var saved = Localization.Culture;
        try
        {
            Localization.Culture = CultureInfo.GetCultureInfo("zh-Hans");
            Assert.Equal("办公自动化库", Localization.Translate(title));
            Assert.StartsWith("办公自动化库", Localization.Translate(description));
            Assert.Equal("办公文档渲染扩展", Localization.Translate(renderingTitle));
            Assert.StartsWith("NewLife.Office 的渲染扩展包", Localization.Translate(renderingDescription));

            Localization.Culture = CultureInfo.GetCultureInfo("en-US");
            Assert.Equal(title, Localization.Translate(title));
        }
        finally
        {
            Localization.Culture = saved;
        }
    }

    [Fact(DisplayName = "译文格式化与占位符保持一致")]
    public void TranslateWithArguments()
    {
        var saved = Localization.Culture;
        try
        {
            Localization.Culture = CultureInfo.GetCultureInfo("zh-Hans");
            Assert.Equal("文件未找到：abc.txt", Localization.Translate("File not found: {0}", "abc.txt"));

            Localization.Culture = CultureInfo.GetCultureInfo("en-US");
            Assert.Equal("File not found: abc.txt", Localization.Translate("File not found: {0}", "abc.txt"));
        }
        finally
        {
            Localization.Culture = saved;
        }
    }

    [Fact(DisplayName = "运行时异常消息按当前区域翻译")]
    public void ExceptionMessagesAreLocalized()
    {
        var saved = Localization.Culture;
        try
        {
            Localization.Culture = CultureInfo.GetCultureInfo("zh-Hans");
            var ex = Assert.Throws<ArgumentException>(() => CfbDocument.Open(new NonSeekableStream()));
            Assert.Contains("流必须支持查找。", ex.Message);

            Localization.Culture = CultureInfo.GetCultureInfo("en-US");
            ex = Assert.Throws<ArgumentException>(() => CfbDocument.Open(new NonSeekableStream()));
            Assert.Contains("Stream must be seekable.", ex.Message);
        }
        finally
        {
            Localization.Culture = saved;
        }
    }

    [Fact(DisplayName = "英文与未登记词条回退原文")]
    public void FallbackToEnglish()
    {
        var saved = Localization.Culture;
        try
        {
            Localization.Culture = CultureInfo.GetCultureInfo("en-US");
            Assert.Equal(AllKeys[0], Localization.Translate(AllKeys[0]));

            Localization.Culture = CultureInfo.GetCultureInfo("zh-Hans");
            Assert.Equal("Not a registered message.", Localization.Translate("Not a registered message."));
        }
        finally
        {
            Localization.Culture = saved;
        }
    }

    private sealed class NonSeekableStream : MemoryStream
    {
        public override Boolean CanSeek => false;
    }
}

/// <summary>本地化测试集合：Localization.Culture 是进程级全局状态，禁止与其他测试并行</summary>
[CollectionDefinition("Localization", DisableParallelization = true)]
public class LocalizationCollection;
