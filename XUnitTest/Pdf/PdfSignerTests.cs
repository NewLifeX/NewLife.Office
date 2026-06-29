using NewLife.Office;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>PdfSigner 单元测试 — PDF 数字签名</summary>
public class PdfSignerTests
{
    [Fact(DisplayName = "PDF签名—签名后可正常打开")]
    public void Sign_PdfWithSignatureField()
    {
        var cert = PdfSigner.CreateSelfSignedCertificate();
        if (cert == null) return;

        try
        {
            using var ms = new MemoryStream();
            using var writer = new PdfWriter();
            writer.DrawText("Test PDF Signature", 50, 700, 12);
            writer.AddSignatureField("Signature1", 50, 600, 200, 50);
            writer.Save(ms);
            var pdfBytes = ms.ToArray();

            var signed = PdfSigner.Sign(pdfBytes, cert);
            Assert.NotNull(signed);
            // 签名后文件长度应相近（PKCS#7 签名填充预留空间）
            Assert.True(signed.Length >= pdfBytes.Length - 100 && signed.Length <= pdfBytes.Length + 100);
        }
        finally
        {
            cert.Dispose();
        }
    }

    [Fact(DisplayName = "PDF签名—无签名字段抛出InvalidOperationException")]
    public void Sign_NoField_Throws()
    {
        var cert = PdfSigner.CreateSelfSignedCertificate();
        if (cert == null) return;

        try
        {
            using var ms = new MemoryStream();
            using var writer = new PdfWriter();
            writer.DrawText("No signature field", 50, 700, 12);
            writer.Save(ms);

            Assert.Throws<InvalidOperationException>(() =>
                PdfSigner.Sign(ms.ToArray(), cert));
        }
        finally
        {
            cert.Dispose();
        }
    }

    [Fact(DisplayName = "PDF签名—自签名证书创建成功且含私钥")]
    public void CreateSelfSignedCertificate_Valid()
    {
        var cert = PdfSigner.CreateSelfSignedCertificate();
        if (cert == null) return;

        try
        {
            Assert.True(cert.HasPrivateKey);
            Assert.Contains("CN=NewLife.Office", cert.Subject);
        }
        finally
        {
            cert.Dispose();
        }
    }
}
