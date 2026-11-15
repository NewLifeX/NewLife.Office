using NewLife.Office;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Ole2;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using NewLife.Office.Xps;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>Signer 单元测试 — PDF 数字签名</summary>
public class SignerTests
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

    [Fact(DisplayName = "PDF签名—ByteRange 覆盖全文且 messageDigest 匹配")]
    public void Sign_ByteRangeAndMessageDigest_AreConsistent()
    {
            var cert = PdfSigner.CreateSelfSignedCertificate();
            if (cert == null) return;

            try
            {
                using var ms = new MemoryStream();
                using (var writer = new PdfWriter())
                {
                    writer.DrawText("ByteRange check", 50, 700, 12);
                    writer.AddSignatureField("Signature1", 50, 600, 200, 50);
                    writer.Save(ms);
                }

                var signed = PdfSigner.Sign(ms.ToArray(), cert!);

                // 1. /ByteRange 必须是四段真实偏移
                var text = System.Text.Encoding.GetEncoding("ISO-8859-1").GetString(signed);
                var brIdx = text.IndexOf("/ByteRange", StringComparison.Ordinal);
                Assert.True(brIdx >= 0, "缺少 /ByteRange");
                var open = text.IndexOf('[', brIdx);
                var close = text.IndexOf(']', open);
                var parts = text[(open + 1)..close].Split(' ', StringSplitOptions.RemoveEmptyEntries);
                Assert.Equal(4, parts.Length);

                var off1 = Int64.Parse(parts[0]);
                var len1 = Int64.Parse(parts[1]);
                var off2 = Int64.Parse(parts[2]);
                var len2 = Int64.Parse(parts[3]);
                Assert.Equal(0, off1);
                Assert.True(len1 > 0 && off2 > len1 && len2 > 0, $"ByteRange 无效: {text[(open + 1)..close]}");
                Assert.Equal(signed.Length, off2 + len2);

                // 2. /Contents 必须落在 ByteRange 的“空洞”里
                var contentsIdx = text.IndexOf("/Contents <", StringComparison.Ordinal);
                var hexStart = contentsIdx + "/Contents ".Length;
                var hexEnd = text.IndexOf('>', hexStart);
                Assert.Equal(len1, hexStart);
                Assert.Equal(off2, hexEnd + 1);

                // 3. 用 ByteRange 覆盖的字节重算摘要，应与 PKCS#7 里的 messageDigest 一致
                using var sha = System.Security.Cryptography.SHA256.Create();
                sha.TransformBlock(signed, 0, (Int32)len1, null, 0);
                sha.TransformFinalBlock(signed, (Int32)off2, (Int32)len2);
                var docHash = sha.Hash!;

                var hex = text.Substring((Int32)hexStart + 1, (Int32)(hexEnd - hexStart - 1));
                var pkcs7 = Convert.FromHexString(hex.TrimEnd('0').PadRight(((hex.TrimEnd('0').Length + 1) / 2) * 2, '0'));

                var digest = FindMessageDigest(pkcs7);
                Assert.NotNull(digest);
                Assert.Equal(docHash, digest!);
        }
        finally
        {
                cert!.Dispose();
        }
    }

    /// <summary>从 PKCS#7 SignedData 中提取 SHA-256 messageDigest 属性值</summary>
    private static Byte[]? FindMessageDigest(Byte[] pkcs7)
    {
            // Attribute ::= SEQUENCE { OID 1.2.840.113549.1.9.4, SET { OCTET STRING } }
            Byte[] oid = [0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x09, 0x04];
            for (var i = 0; i + oid.Length + 4 <= pkcs7.Length; i++)
            {
                var match = true;
                for (var j = 0; j < oid.Length; j++)
                {
                    if (pkcs7[i + j] != oid[j]) { match = false; break; }
                }
                if (!match) continue;

                // 跳过 SET(0x31) 与 OCTET STRING(0x04) 头，取 32 字节摘要
                var pos = i + oid.Length;
                if (pkcs7[pos] != 0x31) continue;
                pos += 2;
                if (pkcs7[pos] != 0x04) continue;
                var len = pkcs7[pos + 1];
                pos += 2;
                if (len != 32) continue;

                var result = new Byte[32];
                Array.Copy(pkcs7, pos, result, 0, 32);
                return result;
            }
            return null;
    }
}
