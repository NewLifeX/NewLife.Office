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

    [Fact(DisplayName = "全部语言表均已翻译全部词条")]
    public void AllKeysTranslated()
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
