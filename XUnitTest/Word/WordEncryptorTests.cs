using NewLife.Office;
using Xunit;

namespace XUnitTest.Word;

/// <summary>WordEncryptor 单元测试 — AES 加密 docx 文档</summary>
public class WordEncryptorTests
{
    [Fact(DisplayName = "Word加密—加密解密往返")]
    public void EncryptDecrypt_RoundTrip()
    {
        // 创建简单 docx
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.AppendParagraph("加密测试内容");
        writer.AppendParagraph("第二段文字123");
        writer.Save(ms);

        var originalBytes = ms.ToArray();
        var password = "TestPassword123";

        // 加密
        var encrypted = WordEncryptor.Encrypt(originalBytes, password);
        Assert.NotNull(encrypted);
        Assert.True(encrypted.Length > 0);
        Assert.NotEqual(originalBytes, encrypted);

        // 解密
        var decrypted = WordEncryptor.Decrypt(encrypted, password);
        Assert.NotNull(decrypted);
        Assert.Equal(originalBytes, decrypted);
    }

    [Fact(DisplayName = "Word加密—错误密码解密失败")]
    public void Decrypt_WrongPassword()
    {
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.AppendParagraph("test");
        writer.Save(ms);

        var encrypted = WordEncryptor.Encrypt(ms.ToArray(), "correct");

        Assert.Throws<System.Security.Cryptography.CryptographicException>(() =>
            WordEncryptor.Decrypt(encrypted, "wrong"));
    }

    [Fact(DisplayName = "Word加密—加密后内容不可读为明文")]
    public void Encrypt_ContentNotPlaintext()
    {
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.AppendParagraph("敏感信息：机密文档");
        writer.Save(ms);

        var encrypted = WordEncryptor.Encrypt(ms.ToArray(), "pwd");
        var encryptedText = System.Text.Encoding.UTF8.GetString(encrypted);

        Assert.DoesNotContain("敏感信息", encryptedText);
        Assert.DoesNotContain("机密文档", encryptedText);
    }

    [Fact(DisplayName = "Word加密—中文字符密码")]
    public void EncryptDecrypt_ChinesePassword()
    {
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.AppendParagraph("中文测试内容");
        writer.Save(ms);

        var originalBytes = ms.ToArray();
        var password = "中文密码测试123";

        var encrypted = WordEncryptor.Encrypt(originalBytes, password);
        var decrypted = WordEncryptor.Decrypt(encrypted, password);

        Assert.Equal(originalBytes, decrypted);
    }

    [Fact(DisplayName = "Word加密—空文档加密")]
    public void EncryptDecrypt_EmptyDocument()
    {
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.Save(ms);

        var encrypted = WordEncryptor.Encrypt(ms.ToArray(), "pwd");
        var decrypted = WordEncryptor.Decrypt(encrypted, "pwd");

        Assert.NotNull(decrypted);
        Assert.Equal(ms.ToArray(), decrypted);
    }
}
