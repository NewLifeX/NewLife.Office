using System.Security.Cryptography;
using System.Text;
using System.Xml;
using NewLife.Office.Ole2;

namespace NewLife.Office.Word;

/// <summary>Office 文档加密器（ECMA-376 Agile Encryption）</summary>
/// <remarks>
/// 按 MS-OFFCRYPTO 2.3.4 的 Agile 方案实现：PBKDF2 风格迭代哈希（SHA-512）、AES-256-CBC 分段加密、
/// 数据完整性 HMAC，输出标准 OLE2 复合文档（EncryptionInfo + EncryptedPackage）。
/// 生成的文件可被 Microsoft Word / WPS / msoffcrypto-tool 直接打开。
/// </remarks>
public static class WordEncryptor
{
    #region 常量
    private const Int32 SaltSize = 16;
    private const Int32 BlockSize = 16;
    private const Int32 KeyBits = 256;
    private const Int32 HashSize = 64;
    private const Int32 SpinCount = 100000;
    private const Int32 SegmentLength = 4096;

    /// <summary>加密描述符版本（Agile）</summary>
    private static readonly Byte[] VersionInfo = [0x04, 0x00, 0x04, 0x00];

    /// <summary>加密描述符保留字段</summary>
    private static readonly Byte[] Reserved = [0x40, 0x00, 0x00, 0x00];

    private static readonly Byte[] BlockKeyVerifierInput = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
    private static readonly Byte[] BlockKeyVerifierHash = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
    private static readonly Byte[] BlockKeyValue = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
    private static readonly Byte[] BlockKeyIntegrity1 = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
    private static readonly Byte[] BlockKeyIntegrity2 = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];
    #endregion

    #region 加密
    /// <summary>加密 docx/xlsx/pptx 字节数组，返回 OLE2 CFB 格式的加密文档</summary>
    /// <param name="docxBytes">普通 OOXML ZIP 字节</param>
    /// <param name="password">打开密码</param>
    /// <returns>CFB 封装的加密文档字节</returns>
    public static Byte[] Encrypt(Byte[] docxBytes, String password)
    {
        if (docxBytes == null) throw new ArgumentNullException(nameof(docxBytes));
        if (password == null) throw new ArgumentNullException(nameof(password));

        var pwd = password;

        // 1. 校验器（verifier）：随机 16 字节，其哈希用于验证密码是否正确
        var verifierSalt = RandomBytes(SaltSize);
        var h = IteratedHash(pwd, verifierSalt);
        var verifierInput = RandomBytes(SaltSize);
        var encryptedVerifierInput = AesCbc(verifierInput, DeriveKey(h, BlockKeyVerifierInput), verifierSalt);
        var verifierHash = AesCbc(Sha512(verifierInput), DeriveKey(h, BlockKeyVerifierHash), verifierSalt);

        // 2. 文档密钥：随机 16 字节，用 0x36 补齐到 32 字节
        var secretKey = new Byte[KeyBits / 8];
        var rawKey = RandomBytes(SaltSize);
        Array.Copy(rawKey, secretKey, rawKey.Length);
        for (var i = rawKey.Length; i < secretKey.Length; i++) secretKey[i] = 0x36;

        var encryptedKeyValue = AesCbc(secretKey, DeriveKey(h, BlockKeyValue), verifierSalt);

        // 3. 正文分段加密
        var keyDataSalt = RandomBytes(SaltSize);
        var encryptedPackage = EncryptPackage(docxBytes, secretKey, keyDataSalt);

        // 4. 数据完整性：HMAC-SHA512 覆盖密文
        var hmacKey = RandomBytes(HashSize);
        var iv1 = IntegrityIv(keyDataSalt, BlockKeyIntegrity1);
        var iv2 = IntegrityIv(keyDataSalt, BlockKeyIntegrity2);
        var encryptedHmacKey = AesCbc(hmacKey, secretKey, iv1);

        Byte[] hmacValue;
        using (var hmac = new HMACSHA512(hmacKey))
            hmacValue = hmac.ComputeHash(encryptedPackage);
        var encryptedHmacValue = AesCbc(hmacValue, secretKey, iv2);

        // 5. 组装 EncryptionInfo 与 CFB 容器
        var xml = BuildDescriptor(keyDataSalt, verifierSalt, encryptedVerifierInput, verifierHash, encryptedKeyValue, encryptedHmacKey, encryptedHmacValue);
        var info = new Byte[VersionInfo.Length + Reserved.Length + xml.Length];
        Array.Copy(VersionInfo, 0, info, 0, VersionInfo.Length);
        Array.Copy(Reserved, 0, info, VersionInfo.Length, Reserved.Length);
        Array.Copy(xml, 0, info, VersionInfo.Length + Reserved.Length, xml.Length);

        return CreateCfb(info, encryptedPackage);
    }

    /// <summary>加密并保存到文件</summary>
    /// <param name="outputPath">输出路径</param>
    /// <param name="docxBytes">普通 OOXML ZIP 字节</param>
    /// <param name="password">打开密码</param>
    public static void Save(String outputPath, Byte[] docxBytes, String password)
    {
        var encrypted = Encrypt(docxBytes, password);
        File.WriteAllBytes(outputPath.GetFullPath(), encrypted);
    }
    #endregion

    #region 解密
    /// <summary>解密 ECMA-376 Agile 加密的 Office 文档</summary>
    /// <param name="encryptedBytes">CFB 加密文档字节</param>
    /// <param name="password">打开密码</param>
    /// <returns>解密后的 OOXML ZIP 字节</returns>
    /// <exception cref="CryptographicException">密码错误或文档损坏</exception>
    public static Byte[] Decrypt(Byte[] encryptedBytes, String password)
    {
        if (encryptedBytes == null) throw new ArgumentNullException(nameof(encryptedBytes));
        if (password == null) throw new ArgumentNullException(nameof(password));

        using var ms = new MemoryStream(encryptedBytes);
        var root = new CfbReader(ms).Parse();

        var info = root.GetStream("EncryptionInfo")?.Data;
        var package = root.GetStream("EncryptedPackage")?.Data;
        if (info == null || info.Length == 0)
            throw new InvalidOperationException(L.T("The EncryptionInfo stream is missing from the CFB container."));
        if (package == null || package.Length == 0)
            throw new InvalidOperationException(L.T("The EncryptedPackage stream is missing from the CFB container."));

        var descriptor = ParseDescriptor(info);

        // 1. 用密码派生密钥并解出文档密钥
        var h = IteratedHash(password, descriptor.VerifierSalt);
        var secretKey = AesCbc(descriptor.EncryptedKeyValue, DeriveKey(h, BlockKeyValue), descriptor.VerifierSalt, encrypt: false);

        // 2. 校验密码：解出校验器与校验器哈希
        var verifierInput = AesCbc(descriptor.EncryptedVerifierInput, DeriveKey(h, BlockKeyVerifierInput), descriptor.VerifierSalt, encrypt: false);
        var verifierHash = AesCbc(descriptor.EncryptedVerifierHashValue, DeriveKey(h, BlockKeyVerifierHash), descriptor.VerifierSalt, encrypt: false);
        var expected = Sha512(verifierInput);
        if (!EqualBytes(expected, verifierHash))
            throw new CryptographicException(L.T("The document could not be decrypted with this password."));

        // 3. 分段解密正文
        return DecryptPackage(package, secretKey, descriptor.KeyDataSalt);
    }
    #endregion

    #region 正文加解密
    private static Byte[] EncryptPackage(Byte[] plain, Byte[] key, Byte[] keyDataSalt)
    {
        using var output = new MemoryStream();
        var header = new Byte[8];
        WriteInt64(header, 0, plain.Length);
        output.Write(header, 0, header.Length);

        for (var offset = 0; offset < plain.Length || offset == 0; offset += SegmentLength)
        {
            var count = Math.Min(SegmentLength, plain.Length - offset);
            var chunk = new Byte[count];
            if (count > 0) Array.Copy(plain, offset, chunk, 0, count);

            // 末段补齐到块大小；整段 4096 字节本身即块对齐
            var padded = PadToBlock(chunk);
            var iv = SegmentIv(keyDataSalt, offset / SegmentLength);
            var enc = AesCbc(padded, key, iv);
            output.Write(enc, 0, enc.Length);

            if (plain.Length == 0) break;
        }

        return output.ToArray();
    }

    private static Byte[] DecryptPackage(Byte[] encrypted, Byte[] key, Byte[] keyDataSalt)
    {
        if (encrypted.Length < 8) throw new InvalidDataException(L.T("The EncryptedPackage data is too short."));

        var totalSize = ReadInt64(encrypted, 0);
        if (totalSize < 0) throw new InvalidDataException(L.T("The EncryptedPackage declares an invalid length."));

        using var output = new MemoryStream();
        var remaining = totalSize;
        var index = 0;
        for (var offset = 8; offset < encrypted.Length && remaining > 0; offset += SegmentLength, index++)
        {
            var count = Math.Min(SegmentLength, encrypted.Length - offset);
            // 末段可能不足一个分组（个别第三方写入器不补齐），按分组对齐后解密
            count -= count % BlockSize;
            if (count <= 0) break;

            var chunk = new Byte[count];
            Array.Copy(encrypted, offset, chunk, 0, count);

            var iv = SegmentIv(keyDataSalt, index);
            var dec = AesCbc(chunk, key, iv, encrypt: false);
            var take = (Int32)Math.Min(remaining, dec.Length);
            output.Write(dec, 0, take);
            remaining -= take;
        }

        return output.ToArray();
    }

    /// <summary>段 IV = Hash(keyDataSalt + 段序号)[0..16]</summary>
    private static Byte[] SegmentIv(Byte[] keyDataSalt, Int32 index)
    {
        var buf = new Byte[keyDataSalt.Length + 4];
        Array.Copy(keyDataSalt, buf, keyDataSalt.Length);
        buf[keyDataSalt.Length] = (Byte)index;
        buf[keyDataSalt.Length + 1] = (Byte)(index >> 8);
        buf[keyDataSalt.Length + 2] = (Byte)(index >> 16);
        buf[keyDataSalt.Length + 3] = (Byte)(index >> 24);
        return Sha512(buf).AsSpan(0, BlockSize).ToArray();
    }

    /// <summary>完整性块 IV = Hash(keyDataSalt + blockKey)[0..16]</summary>
    private static Byte[] IntegrityIv(Byte[] keyDataSalt, Byte[] blockKey)
    {
        var buf = new Byte[keyDataSalt.Length + blockKey.Length];
        Array.Copy(keyDataSalt, buf, keyDataSalt.Length);
        Array.Copy(blockKey, 0, buf, keyDataSalt.Length, blockKey.Length);
        return Sha512(buf).AsSpan(0, BlockSize).ToArray();
    }

    private static Byte[] PadToBlock(Byte[] data)
    {
        var extra = data.Length % BlockSize;
        if (extra == 0) return data;

        var padded = new Byte[data.Length + (BlockSize - extra)];
        Array.Copy(data, padded, data.Length);
        return padded;
    }
    #endregion

    #region 密钥派生
    /// <summary>PBKDF2 风格迭代哈希：Hash(salt + pwd)，再迭代 Hash(LE32(i) + H)</summary>
    private static Byte[] IteratedHash(String password, Byte[] salt)
    {
        var pwdBytes = Encoding.Unicode.GetBytes(password);
        var buf = new Byte[salt.Length + pwdBytes.Length];
        Array.Copy(salt, buf, salt.Length);
        Array.Copy(pwdBytes, 0, buf, salt.Length, pwdBytes.Length);

        var h = Sha512(buf);
        var iter = new Byte[4];
        for (var i = 0; i < SpinCount; i++)
        {
            iter[0] = (Byte)i;
            iter[1] = (Byte)(i >> 8);
            iter[2] = (Byte)(i >> 16);
            iter[3] = (Byte)(i >> 24);

            var next = new Byte[4 + h.Length];
            Array.Copy(iter, next, 4);
            Array.Copy(h, 0, next, 4, h.Length);
            h = Sha512(next);
        }
        return h;
    }

    /// <summary>最终密钥 = Hash(H + blockKey) 截断到密钥长度</summary>
    private static Byte[] DeriveKey(Byte[] h, Byte[] blockKey)
    {
        var buf = new Byte[h.Length + blockKey.Length];
        Array.Copy(h, buf, h.Length);
        Array.Copy(blockKey, 0, buf, h.Length, blockKey.Length);

        var hash = Sha512(buf);
        var key = new Byte[KeyBits / 8];
        Array.Copy(hash, key, key.Length);
        return key;
    }
    #endregion

    #region 描述符
    private sealed class Descriptor
    {
        public Byte[] KeyDataSalt = [];
        public Byte[] VerifierSalt = [];
        public Byte[] EncryptedVerifierInput = [];
        public Byte[] EncryptedVerifierHashValue = [];
        public Byte[] EncryptedKeyValue = [];
    }

    private static Byte[] BuildDescriptor(Byte[] keyDataSalt, Byte[] verifierSalt, Byte[] encVerifierInput,
        Byte[] encVerifierHashValue, Byte[] encKeyValue, Byte[] encHmacKey, Byte[] encHmacValue)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        sb.Append("<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" ");
        sb.Append("xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">");
        sb.Append("<keyData saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" ");
        sb.Append("cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" ");
        sb.Append($"saltValue=\"{Convert.ToBase64String(keyDataSalt)}\"/>");
        sb.Append("<dataIntegrity ");
        sb.Append($"encryptedHmacKey=\"{Convert.ToBase64String(encHmacKey)}\" ");
        sb.Append($"encryptedHmacValue=\"{Convert.ToBase64String(encHmacValue)}\"/>");
        sb.Append("<keyEncryptors><keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">");
        sb.Append("<p:encryptedKey spinCount=\"100000\" saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" ");
        sb.Append("cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" ");
        sb.Append($"saltValue=\"{Convert.ToBase64String(verifierSalt)}\" ");
        sb.Append($"encryptedVerifierHashInput=\"{Convert.ToBase64String(encVerifierInput)}\" ");
        sb.Append($"encryptedVerifierHashValue=\"{Convert.ToBase64String(encVerifierHashValue)}\" ");
        sb.Append($"encryptedKeyValue=\"{Convert.ToBase64String(encKeyValue)}\"/>");
        sb.Append("</keyEncryptor></keyEncryptors></encryption>");

        return new UTF8Encoding(false).GetBytes(sb.ToString());
    }

    private static Descriptor ParseDescriptor(Byte[] info)
    {
        // 头部：版本(4) + 保留(4)，随后是 XML 描述符
        var xmlStart = -1;
        for (var i = 0; i + 1 < info.Length; i++)
        {
            if (info[i] == (Byte)'<' && info[i + 1] == (Byte)'?') { xmlStart = i; break; }
        }
        if (xmlStart < 0) throw new InvalidOperationException(L.T("The XML declaration is missing from EncryptionInfo."));

        var xml = Encoding.UTF8.GetString(info, xmlStart, info.Length - xmlStart).TrimEnd('\0');
        var doc = OfficeXml.LoadXml(xml);

        var result = new Descriptor
        {
            KeyDataSalt = ReadSalt(doc, "keyData", "saltValue"),
            VerifierSalt = ReadSalt(doc, "encryptedKey", "saltValue"),
            EncryptedVerifierInput = ReadSalt(doc, "encryptedKey", "encryptedVerifierHashInput"),
            EncryptedVerifierHashValue = ReadSalt(doc, "encryptedKey", "encryptedVerifierHashValue"),
            EncryptedKeyValue = ReadSalt(doc, "encryptedKey", "encryptedKeyValue"),
        };
        return result;
    }

    private static Byte[] ReadSalt(XmlDocument doc, String localName, String attribute)
    {
        var node = doc.GetElementsByTagName("*");
        foreach (XmlNode item in node)
        {
            if (item.LocalName != localName) continue;
            if (item.Attributes?[attribute]?.Value is not { Length: > 0 } value) continue;

            try
            {
                return Convert.FromBase64String(value);
            }
            catch (FormatException)
            {
                throw new InvalidOperationException(L.T("The saltValue format in EncryptionInfo is invalid."));
            }
        }

        throw new InvalidOperationException(L.T("saltValue is missing from EncryptionInfo."));
    }
    #endregion

    #region 辅助
    private static Byte[] CreateCfb(Byte[] encryptionInfo, Byte[] encryptedPackage)
    {
        var root = new CfbStorage { Name = "Root Entry" };
        root.Children.Add(new CfbStream { Name = "EncryptionInfo", Data = encryptionInfo });
        root.Children.Add(new CfbStream { Name = "EncryptedPackage", Data = encryptedPackage });

        using var ms = new MemoryStream();
        new CfbWriter().Write(root, ms);
        return ms.ToArray();
    }

    private static Byte[] AesCbc(Byte[] data, Byte[] key, Byte[] iv, Boolean encrypt = true)
    {
        using var aes = Aes.Create();
        aes.Key = key.AsSpan(0, KeyBits / 8).ToArray();
        aes.IV = iv.AsSpan(0, BlockSize).ToArray();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;

        using var tf = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor();
        return tf.TransformFinalBlock(data, 0, data.Length);
    }

    private static Byte[] Sha512(Byte[] data)
    {
        using var sha = SHA512.Create();
        return sha.ComputeHash(data);
    }

    private static Byte[] RandomBytes(Int32 count)
    {
        var buf = new Byte[count];
        using var rng = RandomNumberGenerator.Create();
        rng.GetBytes(buf);
        return buf;
    }

    private static Boolean EqualBytes(Byte[] x, Byte[] y)
    {
        if (x.Length != y.Length) return false;
        for (var i = 0; i < x.Length; i++)
            if (x[i] != y[i]) return false;

        return true;
    }

    private static void WriteInt64(Byte[] buf, Int32 offset, Int64 value)
    {
        for (var i = 0; i < 8; i++)
            buf[offset + i] = (Byte)(value >> (8 * i));
    }

    private static Int64 ReadInt64(Byte[] buf, Int32 offset)
    {
        Int64 value = 0;
        for (var i = 0; i < 8; i++)
            value |= (Int64)buf[offset + i] << (8 * i);

        return value;
    }
    #endregion
}
