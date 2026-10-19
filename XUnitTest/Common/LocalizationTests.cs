using System.Globalization;
using System.Runtime.CompilerServices;
using NewLife;
using NewLife.Office;
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

    /// <summary>已登记的运行时词条（英文源文）</summary>
    public static readonly String[] Keys =
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
    ];

    /// <summary>全部已登记词条（12 条核心词条 + 45 条新增异常/CLI 词条）</summary>
    public static readonly String[] AllKeys =
    [
        .. Keys,
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
    ];

    [Fact(DisplayName = "31 种语言均已翻译核心词条")]
    public void CoreKeysTranslated()
    {
        var saved = Localization.Culture;
        try
        {
            foreach (var name in Cultures)
            {
                Localization.Culture = CultureInfo.GetCultureInfo(name);
                foreach (var key in Keys)
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

    [Fact(DisplayName = "简繁中文已翻译全部词条")]
    public void ChineseTranslationsComplete()
    {
        var saved = Localization.Culture;
        try
        {
            foreach (var name in new[] { "zh-Hans", "zh-Hant" })
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

    [Fact(DisplayName = "英文与未登记词条回退原文")]
    public void FallbackToEnglish()
    {
        var saved = Localization.Culture;
        try
        {
            Localization.Culture = CultureInfo.GetCultureInfo("en-US");
            Assert.Equal(Keys[0], Localization.Translate(Keys[0]));

            Localization.Culture = CultureInfo.GetCultureInfo("zh-Hans");
            Assert.Equal("Not a registered message.", Localization.Translate("Not a registered message."));
        }
        finally
        {
            Localization.Culture = saved;
        }
    }
}

/// <summary>本地化测试集合：Localization.Culture 是进程级全局状态，禁止与其他测试并行</summary>
[CollectionDefinition("Localization", DisableParallelization = true)]
public class LocalizationCollection;
