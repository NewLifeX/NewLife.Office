using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using NewLife;

namespace NewLife.Office.Pdf;

/// <summary>PDF 数字签名器，实现 PKCS#7 分离签名（adbe.pkcs7.detached）</summary>
/// <remarks>
/// 符合 ISO 32000-1 §12.8 数字签名规范。
/// 使用手动构建的 ASN.1 DER 编码 PKCS#7 SignedData 结构，无外部加密库依赖。
/// 
/// 使用流程：
/// 1. 用 PdfWriter.AddSignatureField() 创建签名字段预留区
/// 2. 调用 Signer.Sign() 计算签名并填入 /Contents
/// </remarks>
public static class PdfSigner
{
    #region 签名入口
    /// <summary>对 PDF 字节数组进行数字签名</summary>
    /// <param name="pdfBytes">PDF 原始字节（需已包含签名字段占位）</param>
    /// <param name="certificate">X.509 证书（需含 RSA 私钥）</param>
    /// <param name="signatureFieldName">签名字段名称（默认 "Signature1"）</param>
    /// <returns>签名后的 PDF 字节</returns>
    public static Byte[] Sign(Byte[] pdfBytes, X509Certificate2 certificate, String signatureFieldName = "Signature1")
    {
        if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));
        if (certificate == null) throw new ArgumentNullException(nameof(certificate));

        // 不修改调用方缓冲区
        pdfBytes = (Byte[])pdfBytes.Clone();

        // 1. 定位签名字段中的 /Contents 占位符
        var pdfText = Encoding.GetEncoding("ISO-8859-1").GetString(pdfBytes);
        var fieldPattern = $"/T ({EscapePdfString(signatureFieldName)})";
        var fieldIdx = pdfText.IndexOf(fieldPattern, StringComparison.Ordinal);
        if (fieldIdx < 0)
            throw new InvalidOperationException(L.T("Signature field '{0}' was not found.", signatureFieldName));

        var contentsIdx = pdfText.IndexOf("/Contents <", fieldIdx, StringComparison.Ordinal);
        if (contentsIdx < 0)
            throw new InvalidOperationException(L.T("The signature field is missing the /Contents placeholder."));

        var hexStart = contentsIdx + "/Contents ".Length;   // '<' 的位置
        var valStart = hexStart + 1;                        // 十六进制数据起点
        var valEnd = pdfText.IndexOf('>', valStart);
        if (valEnd < 0)
            throw new InvalidOperationException(L.T("The /Contents placeholder is malformed."));

        // 2. ByteRange：签名覆盖除 <...> 之外的全部字节
        var range1Len = hexStart;
        var range2Start = valEnd + 1;
        var range2Len = pdfBytes.Length - range2Start;

        // 3. 先回填 ByteRange（定宽占位，替换后长度不变），再计算摘要，保证摘要覆盖的是最终字节
        ReplaceByteRange(pdfBytes, fieldIdx, range1Len, range2Start, range2Len, pdfText);

        Byte[] docHash;
        using (var sha = SHA256.Create())
        {
            sha.TransformBlock(pdfBytes, 0, range1Len, null, 0);
            sha.TransformFinalBlock(pdfBytes, range2Start, range2Len);
            docHash = sha.Hash!;
        }

        // 4. 构建 PKCS#7 签名（含 messageDigest 等签名属性）
        var placeholderLen = valEnd - valStart;
        var pkcs7 = BuildPkcs7(docHash, certificate);
        var sigHex = BitConverter.ToString(pkcs7).Replace("-", "");
        if (sigHex.Length > placeholderLen)
            throw new InvalidOperationException(L.T("Signature length {0} exceeds the reserved {1} characters.", sigHex.Length, placeholderLen));

        // 5. 写入 /Contents
        var paddedSig = Encoding.ASCII.GetBytes(sigHex.PadRight(placeholderLen, '0'));
        Array.Copy(paddedSig, 0, pdfBytes, valStart, paddedSig.Length);

        return pdfBytes;
    }

    /// <summary>把 /ByteRange 的定宽占位替换为真实偏移</summary>
    private static void ReplaceByteRange(Byte[] pdfBytes, Int32 fieldIdx, Int32 range1Len, Int32 range2Start, Int32 range2Len, String pdfText)
    {
        // /ByteRange 与 /Contents 在同一个签名字典内，从字段起点向后找即可
        var brIdx = pdfText.IndexOf("/ByteRange", fieldIdx, StringComparison.Ordinal);
        if (brIdx < 0) return;

        var open = pdfText.IndexOf('[', brIdx);
        var close = open < 0 ? -1 : pdfText.IndexOf(']', open);
        if (open < 0 || close < 0) return;

        var actual = $"[0 {range1Len:D10} {range2Start:D10} {range2Len:D10}]";
        var slot = pdfText.Substring(open, close - open + 1);

        // 定宽占位符才能原地替换；长度不同说明占位符不是本库写出的格式
        if (slot.Length != actual.Length) return;

        var bytes = Encoding.ASCII.GetBytes(actual);
        Array.Copy(bytes, 0, pdfBytes, open, bytes.Length);
    }

    /// <summary>对 PDF 文件签名并保存</summary>
    public static void Sign(String inputPath, String outputPath, X509Certificate2 certificate,
        String signatureFieldName = "Signature1")
    {
        var pdfBytes = File.ReadAllBytes(inputPath.GetFullPath());
        var signed = Sign(pdfBytes, certificate, signatureFieldName);
        File.WriteAllBytes(outputPath.GetFullPath(), signed);
    }
    #endregion

    #region PKCS#7 构建（手动 ASN.1 DER）
    private static Byte[] BuildPkcs7(Byte[] docHash, X509Certificate2 cert)
    {
        // 签名属性：contentType + messageDigest + signingTime；RSA 签名作用于整个 SignedAttrs 的 DER
        // DER 要求 SET OF 的元素按编码字节升序排列，否则严格的解析器会报 InvalidSetOrdering
        var attrs = new List<Byte[]>
        {
            BuildSequence(BuildOid("1.2.840.113549.1.9.3"), BuildSet(BuildOid("1.2.840.113549.1.7.1"))),
            BuildSequence(BuildOid("1.2.840.113549.1.9.4"), BuildSet(BuildOctetString(docHash))),
            BuildSequence(BuildOid("1.2.840.113549.1.9.5"), BuildSet(BuildUtcTime(DateTime.UtcNow))),
        };
        attrs.Sort(CompareDer);
        var signedAttrs = BuildSet([.. attrs]);

        Byte[] attrHash;
        using (var sha = SHA256.Create())
            attrHash = sha.ComputeHash(signedAttrs);

        var signature = SignHash(attrHash, cert);
        var certBytes = cert.RawData;
        var issuerBytes = cert.IssuerName.RawData;
        var serialBytes = NormalizeSerial(cert.GetSerialNumber());

        // 构建底层元素
        var digestAlgoId = BuildSequence(
            BuildOid("2.16.840.1.101.3.4.2.1"), // SHA-256
            BuildNull());

        var digestAlgos = BuildSet(digestAlgoId);

        var contentInfo = BuildSequence(
            BuildOid("1.2.840.113549.1.7.1")); // data

        // SignerInfo
        var issuerSerial = BuildSequence(
            issuerBytes,
            BuildIntegerBytes(serialBytes));

        var signerDigestAlgo = BuildSequence(
            BuildOid("2.16.840.1.101.3.4.2.1"),
            BuildNull());

        var signerEncryptAlgo = BuildSequence(
            BuildOid("1.2.840.113549.1.1.1"), // RSA
            BuildNull());

        var encryptedDigest = BuildOctetString(signature);

        // SignedAttrs 以 [0] IMPLICIT 形式内嵌（替换 SET 的标签，内容保持 DER 原样）
        var headerLen = 1 + LengthBytesCount(signedAttrs);
        var attrContent = new Byte[signedAttrs.Length - headerLen];
        Array.Copy(signedAttrs, headerLen, attrContent, 0, attrContent.Length);
        var signedAttrsField = BuildTlv(0xA0, attrContent);

        var signerInfo = BuildSequence(
            BuildInteger(1),
            issuerSerial,
            signerDigestAlgo,
            signedAttrsField,
            signerEncryptAlgo,
            encryptedDigest);

        var signerInfos = BuildSet(signerInfo);

        // SignedData 内容
        var signedDataContent = Concat(
            BuildInteger(1),
            digestAlgos,
            contentInfo,
            certBytes.Length > 0 ? BuildImplicitTag(0, certBytes) : [],
            signerInfos);

        var signedData = BuildSequence(signedDataContent);

        // 外层 ContentInfo
        var outerContent = Concat(
            BuildOid("1.2.840.113549.1.7.2"), // signedData
            BuildExplicitTag(0, signedData));

        return BuildSequence(outerContent);
    }

    private static Byte[] SignHash(Byte[] hash, X509Certificate2 cert)
    {
        // .NET Framework 4.5 路径：使用 RSACryptoServiceProvider
#if NETFRAMEWORK
        var csp = cert.PrivateKey as RSACryptoServiceProvider;
        if (csp != null)
            return csp.SignHash(hash, CryptoConfig.MapNameToOID("SHA256"));

        // 尝试 RSA.Create()
        try
        {
            var rsa = cert.PrivateKey as RSA;
            if (rsa != null)
            {
                var param = rsa.ExportParameters(true);
                var csp2 = new RSACryptoServiceProvider();
                csp2.ImportParameters(param);
                return csp2.SignHash(hash, CryptoConfig.MapNameToOID("SHA256"));
            }
        }
        catch { }
#else
        var rsa = cert.GetRSAPrivateKey();
        if (rsa != null)
            return rsa.SignHash(hash, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
#endif

        throw new InvalidOperationException(L.T("The certificate does not contain a usable RSA private key."));
    }

    private static Byte[] NormalizeSerial(Byte[] serial)
    {
        var start = 0;
        while (start < serial.Length - 1 && serial[start] == 0)
            start++;
        var result = new Byte[serial.Length - start];
        Array.Copy(serial, start, result, 0, result.Length);
        return result;
    }
    #endregion

    #region ASN.1 DER 构建原子方法（返回完整 TLV 字节数组）
    private static Byte[] BuildSequence(params Byte[][] items) => BuildTlv(0x30, Concat(items));
    private static Byte[] BuildSet(params Byte[][] items) => BuildTlv(0x31, Concat(items));
    private static Byte[] BuildOid(String oid) => BuildTlv(0x06, EncodeOidBytes(oid));
    private static Byte[] BuildNull() => [0x05, 0x00];
    private static Byte[] BuildInteger(Int32 value) => BuildTlv(0x02, EncodeIntBytes(value));
    private static Byte[] BuildIntegerBytes(Byte[] value) => BuildTlv(0x02, value);
    private static Byte[] BuildOctetString(Byte[] value) => BuildTlv(0x04, value);
    private static Byte[] BuildUtcTime(DateTime utc) => BuildTlv(0x17, Encoding.ASCII.GetBytes(utc.ToString("yyMMddHHmmss'Z'", System.Globalization.CultureInfo.InvariantCulture)));

    /// <summary>返回 DER 长度字段占用的字节数（不含标签字节）</summary>
    private static Int32 LengthBytesCount(Byte[] tlv)
    {
        if (tlv.Length < 2) return 1;
        var first = tlv[1];
        if (first < 0x80) return 1;
        return 1 + (first & 0x7F);
    }

    /// <summary>按无符号字节序比较两段 DER 编码，用于 SET OF 排序</summary>
    private static Int32 CompareDer(Byte[] x, Byte[] y)
    {
        var len = Math.Min(x.Length, y.Length);
        for (var i = 0; i < len; i++)
        {
            if (x[i] != y[i]) return x[i] < y[i] ? -1 : 1;
        }
        return x.Length.CompareTo(y.Length);
    }
    private static Byte[] BuildImplicitTag(Int32 tagNum, Byte[] value) => BuildTlv((Byte)(0xA0 | (tagNum & 0x1F)), value);
    private static Byte[] BuildExplicitTag(Int32 tagNum, Byte[] value) => BuildTlv((Byte)(0xA0 | (tagNum & 0x1F)), value);

    private static Byte[] BuildTlv(Byte tag, Byte[] value)
    {
        var lenBytes = EncodeLength(value.Length);
        var result = new Byte[1 + lenBytes.Length + value.Length];
        result[0] = tag;
        Array.Copy(lenBytes, 0, result, 1, lenBytes.Length);
        Array.Copy(value, 0, result, 1 + lenBytes.Length, value.Length);
        return result;
    }

    private static Byte[] Concat(params Byte[][] arrays)
    {
        var total = 0;
        foreach (var a in arrays) total += a.Length;
        var result = new Byte[total];
        var offset = 0;
        foreach (var a in arrays)
        {
            Array.Copy(a, 0, result, offset, a.Length);
            offset += a.Length;
        }
        return result;
    }

    private static Byte[] EncodeLength(Int32 length)
    {
        if (length < 128) return [(Byte)length];
        if (length < 0x100) return [0x81, (Byte)length];
        if (length < 0x10000) return [0x82, (Byte)(length >> 8), (Byte)length];
        if (length < 0x1000000) return [0x83, (Byte)(length >> 16), (Byte)(length >> 8), (Byte)length];
        return [0x84, (Byte)(length >> 24), (Byte)(length >> 16), (Byte)(length >> 8), (Byte)length];
    }

    private static Byte[] EncodeOidBytes(String oid)
    {
        var parts = oid.Split('.');
        var result = new List<Byte> { (Byte)(40 * Int32.Parse(parts[0]) + Int32.Parse(parts[1])) };
        for (var i = 2; i < parts.Length; i++)
        {
            var val = Int64.Parse(parts[i]);
            var stack = new Stack<Byte>();
            stack.Push((Byte)(val & 0x7F));
            val >>= 7;
            while (val > 0)
            {
                stack.Push((Byte)((val & 0x7F) | 0x80));
                val >>= 7;
            }
            while (stack.Count > 0) result.Add(stack.Pop());
        }
        return result.ToArray();
    }

    private static Byte[] EncodeIntBytes(Int32 value)
    {
        if (value == 0) return [0];
        var bytes = new List<Byte>();
        var v = value;
        while (v != 0) { bytes.Insert(0, (Byte)(v & 0xFF)); v >>= 8; }
        if ((bytes[0] & 0x80) != 0) bytes.Insert(0, 0);
        return bytes.ToArray();
    }
    #endregion

    #region 工具
    /// <summary>创建自签名证书（用于测试）</summary>
    /// <returns>含 RSA 私钥的自签名证书；当前平台不支持时返回 null</returns>
    public static X509Certificate2? CreateSelfSignedCertificate(String subjectName = "CN=NewLife.Office PDF Signer")
    {
#if NETFRAMEWORK || NETSTANDARD2_0
        return null;
#else
        try
        {
            using var rsa = RSA.Create(2048);
            var subject = new X500DistinguishedName(subjectName);
            var req = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            req.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
            var cert = req.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddYears(5));
            return new X509Certificate2(cert.Export(X509ContentType.Pfx));
        }
        catch
        {
            return null;
        }
#endif
    }

    private static String EscapePdfString(String s)
    {
        var sb = new StringBuilder();
        foreach (var c in s)
        {
            switch (c)
            {
                case '(': sb.Append("\\("); break;
                case ')': sb.Append("\\)"); break;
                case '\\': sb.Append("\\\\"); break;
                default: sb.Append(c); break;
            }
        }
        return sb.ToString();
    }
    #endregion
}
