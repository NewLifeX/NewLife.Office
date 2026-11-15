using System.Security.Cryptography;
using System.Text;
using NewLife.Buffers;

namespace NewLife.Office.Pdf;

/// <summary>PDF 加密修订版本</summary>
public enum CipherRevision
{
    /// <summary>RC4 40-bit（PDF 1.1-1.3，已废弃）</summary>
    Rc4_40 = 2,

    /// <summary>RC4 128-bit（PDF 1.4-1.5，默认）</summary>
    Rc4_128 = 3,

    /// <summary>AES-128（PDF 1.6-1.7 ExtensionLevel 3）</summary>
    Aes_128 = 4,

    /// <summary>AES-256（PDF 2.0）</summary>
    Aes_256 = 6,
}

internal sealed class PdfEncryptor
{
    #region 属性
    private static readonly Byte[] _padding =
    [
        0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
        0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
        0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
        0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A,
    ];

    private readonly Byte[] _key; // 全局加密密钥（RC4: MD5→16字节; AES-128: MD5→16字节; AES-256: SHA-256→32字节）

    /// <summary>Owner 密钥条目</summary>
    public Byte[] OEntry { get; }

    /// <summary>User 密钥条目</summary>
    public Byte[] UEntry { get; }

    /// <summary>加密权限标志（写入加密字典 /P）</summary>
    public Int32 EncPermissions { get; }

    /// <summary>加密修订版本</summary>
    public CipherRevision Revision { get; }

    /// <summary>OE 条目（AES-256 所有者验证密钥，/R 6 专用）</summary>
    public Byte[]? OEEntry { get; }

    /// <summary>UE 条目（AES-256 用户验证密钥，/R 6 专用）</summary>
    public Byte[]? UEEntry { get; }

    /// <summary>Perms 条目（AES-256 加密的权限，/R 6 专用）</summary>
    public Byte[]? PermsEntry { get; }
    #endregion

    #region 构造
    /// <summary>实例化 PDF 加密器</summary>
    /// <param name="userPwd">用户密码（打开密码），null 表示空密码</param>
    /// <param name="ownerPwd">所有者密码（权限密码）</param>
    /// <param name="permissions">权限标志位</param>
    /// <param name="fileId">文件标识符（16 字节 MD5）</param>
    /// <param name="revision">加密修订版本</param>
    public PdfEncryptor(String? userPwd, String? ownerPwd, Int32 permissions, Byte[] fileId, CipherRevision revision = CipherRevision.Rc4_128)
    {
        EncPermissions = permissions;
        Revision = revision;
        var fid = fileId.Length >= 16 ? fileId.AsSpan(0, 16).ToArray() : fileId;

        switch (revision)
        {
            case CipherRevision.Aes_256:
                (_key, OEntry, UEntry, OEEntry, UEEntry, PermsEntry) = ComputeAes256(userPwd, ownerPwd, permissions, fid);
                break;
            case CipherRevision.Aes_128:
                (_key, OEntry, UEntry) = ComputeAes128(userPwd, ownerPwd, permissions, fid);
                break;
            default:
                (_key, OEntry, UEntry) = ComputeRc4(userPwd, ownerPwd, permissions, fid);
                break;
        }
    }
    #endregion

    #region 方法
    /// <summary>加密字节数组</summary>
    /// <param name="data">原始字节</param>
    /// <param name="objNum">PDF 对象号</param>
    /// <param name="genNum">PDF 代数号</param>
    /// <returns>加密后字节</returns>
    public Byte[] EncryptBytes(Byte[] data, Int32 objNum, Int32 genNum)
    {
        return Revision switch
        {
            CipherRevision.Aes_128 => EncryptAes128(data, objNum, genNum),
            CipherRevision.Aes_256 => EncryptAes256(data, objNum, genNum),
            _ => ComputeRc4(ObjKey(objNum, genNum), data),
        };
    }

    /// <summary>加密字符串，返回 PDF 十六进制字符串格式 &lt;hex&gt;</summary>
    public String EncryptString(String s, Int32 objNum, Int32 genNum)
    {
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
        {
            if (ch >= 32 && ch < 256) sb.Append(ch);
            else if (ch >= 256) sb.Append('?');
        }
        var bytes = Encoding.GetEncoding(28591).GetBytes(sb.ToString());
        var encrypted = EncryptBytes(bytes, objNum, genNum);
        return "<" + BitConverter.ToString(encrypted).Replace("-", "") + ">";
    }
    #endregion

    #region RC4 128-bit 计算（/R 3）
    private static (Byte[] Key, Byte[] O, Byte[] U) ComputeRc4(String? userPwd, String? ownerPwd, Int32 permissions, Byte[] fid)
    {
        var uPass = PadPwd(userPwd ?? String.Empty);
        var oPass = PadPwd(ownerPwd ?? (userPwd ?? String.Empty));

        // 算法 3.3：计算 O 条目
        var ownerKey = ComputeMd5(oPass);
        for (var i = 0; i < 50; i++) ownerKey = ComputeMd5(ownerKey);
        var oStep = ComputeRc4(ownerKey, uPass);
        for (var i = 1; i <= 19; i++)
        {
            var k = new Byte[ownerKey.Length];
            for (var j = 0; j < k.Length; j++) k[j] = (Byte)(ownerKey[j] ^ i);
            oStep = ComputeRc4(k, oStep);
        }
        var o = oStep;

        // 算法 3.2：计算全局加密密钥
        var buf = new Byte[32 + 32 + 4 + fid.Length];
        var bw = new SpanWriter(buf, 0, buf.Length);
        bw.Write(uPass);
        bw.Write(o);
        bw.Write(permissions);
        bw.Write(fid);
        var keyHash = ComputeMd5(buf);
        for (var i = 0; i < 50; i++) keyHash = ComputeMd5(keyHash);
        var key = keyHash;

        // 算法 3.5：计算 U 条目
        var uBuf = new Byte[_padding.Length + fid.Length];
        var ubw = new SpanWriter(uBuf, 0, uBuf.Length);
        ubw.Write(_padding);
        ubw.Write(fid);
        var uStep = ComputeRc4(key, ComputeMd5(uBuf));
        for (var i = 1; i <= 19; i++)
        {
            var k = new Byte[key.Length];
            for (var j = 0; j < k.Length; j++) k[j] = (Byte)(key[j] ^ i);
            uStep = ComputeRc4(k, uStep);
        }
        var u = new Byte[32];
        Array.Copy(uStep, u, uStep.Length);
        return (key, o, u);
    }
    #endregion

    #region AES-128 计算（/R 4）
    private static (Byte[] Key, Byte[] O, Byte[] U) ComputeAes128(String? userPwd, String? ownerPwd, Int32 permissions, Byte[] fid)
    {
        var uPass = PadPwd(userPwd ?? String.Empty);
        var oPass = PadPwd(ownerPwd ?? (userPwd ?? String.Empty));

        // 与 RC4 相同的密钥派生（算法 3.2），但用 AES-128 加密
        var (key, o, u) = ComputeRc4(userPwd, ownerPwd, permissions, fid);

        // AES-128 需要加密元数据标志
        return (key, o, u);
    }

    private Byte[] EncryptAes128(Byte[] data, Int32 objNum, Int32 genNum)
    {
        var objKey = ObjKey(objNum, genNum);
        using var aes = Aes.Create();
        aes.Key = objKey;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.GenerateIV();

        var iv = aes.IV; // 16 字节随机 IV

        using var encryptor = aes.CreateEncryptor();
        using var ms = new MemoryStream();
        ms.Write(iv, 0, iv.Length);
        using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
        {
            cs.Write(data, 0, data.Length);
        }
        return ms.ToArray(); // IV + 密文
    }
    #endregion

    #region AES-256 计算（/R 6）
    private static (Byte[] Key, Byte[] O, Byte[] U, Byte[]? OE, Byte[]? UE, Byte[]? Perms) ComputeAes256(String? userPwd, String? ownerPwd, Int32 permissions, Byte[] fid)
    {
        // ISO 32000-2：文件加密密钥随机生成，与密码无关；密码只用于包装（OE/UE）
        var fileKey = new Byte[32];
        using (var rng = RandomNumberGenerator.Create())
            rng.GetBytes(fileKey);

        var uPwd = userPwd ?? String.Empty;
        var oPwd = ownerPwd ?? uPwd;

        // U = hash2B(userPwd, uValidationSalt, "") + uValidationSalt + uKeySalt
        var uValidationSalt = RandomBytes(8);
        var uKeySalt = RandomBytes(8);
        var u = new Byte[48];
        Array.Copy(Hash2B(uPwd, uValidationSalt, []), 0, u, 0, 32);
        Array.Copy(uValidationSalt, 0, u, 32, 8);
        Array.Copy(uKeySalt, 0, u, 40, 8);

        // UE = AES-256(no IV, no padding) 加密 fileKey，密钥为 hash2B(userPwd, uKeySalt, "")
        var ue = AesNoPadding(Hash2B(uPwd, uKeySalt, []), fileKey, encrypt: true, useIv: false);

        // O = hash2B(ownerPwd, oValidationSalt, U) + oValidationSalt + oKeySalt
        var oValidationSalt = RandomBytes(8);
        var oKeySalt = RandomBytes(8);
        var o = new Byte[48];
        Array.Copy(Hash2B(oPwd, oValidationSalt, u), 0, o, 0, 32);
        Array.Copy(oValidationSalt, 0, o, 32, 8);
        Array.Copy(oKeySalt, 0, o, 40, 8);

        // OE = AES-256(no IV, no padding) 加密 fileKey，密钥为 hash2B(ownerPwd, oKeySalt, U)
        var oe = AesNoPadding(Hash2B(oPwd, oKeySalt, u), fileKey, encrypt: true, useIv: false);

        // Perms：4 字节扩展权限 + 4 字节权限 + 'T' 'a' 'd' 'b' + 4 字节随机
        var permsData = new Byte[16];
        var pw = new SpanWriter(permsData, 0, 16);
        pw.Write(0xFFFFFFFFu);
        pw.Write((UInt32)permissions);
        pw.Write((Byte)'T');
        pw.Write((Byte)'a');
        pw.Write((Byte)'d');
        pw.Write((Byte)'b');
        var tail = RandomBytes(4);
        pw.Write(tail);
        var perms = AesNoPadding(fileKey, permsData, encrypt: true, useIv: false);

        return (fileKey, o, u, oe, ue, perms);
    }

    /// <summary>ISO 32000-2 Algorithm 2.B 密码哈希</summary>
    private static Byte[] Hash2B(String password, Byte[] salt, Byte[] uData)
    {
        // 密码按 UTF-8 处理（PDF 2.0 规定），SASLprep 归一在本库不做
        var pwd = Encoding.UTF8.GetBytes(password);
        var k = ComputeSha256(Concat(pwd, salt, uData));
        var e = new Byte[64];
        var round = 0;
        do
        {
            // K1 = (password + K + uData) 重复 64 次
            var block = Concat(pwd, k, uData);
            var k1 = new Byte[block.Length * 64];
            for (var i = 0; i < 64; i++)
                Array.Copy(block, 0, k1, i * block.Length, block.Length);

            // E = AES-128-CBC(无填充) 加密 K1，密钥 K[0..16]，IV K[16..32]
            using var aes = Aes.Create();
            aes.Key = k.AsSpan(0, 16).ToArray();
            aes.IV = k.AsSpan(16, 16).ToArray();
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            using (var enc = aes.CreateEncryptor())
                e = enc.TransformFinalBlock(k1, 0, k1.Length);

            var mod = 0;
            for (var i = 0; i < 16; i++) mod += e[i];
            mod %= 3;

            k = mod switch
            {
                0 => ComputeSha256(e),
                1 => ComputeSha384(e),
                _ => ComputeSha512(e),
            };
            round++;
        }
        while (round < 64 || e[e.Length - 1] > round - 32);

        return k.AsSpan(0, 32).ToArray();
    }

    /// <summary>AES-256 加解密；useIv=false 时使用全零 IV（密钥包装场景）</summary>
    private static Byte[] AesNoPadding(Byte[] key, Byte[] data, Boolean encrypt, Boolean useIv)
    {
        using var aes = Aes.Create();
        aes.Key = key.AsSpan(0, 32).ToArray();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.IV = useIv ? RandomBytes(16) : new Byte[16];

        using var tf = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor();
        var result = tf.TransformFinalBlock(data, 0, data.Length);
        if (!encrypt || !useIv) return result;

        var output = new Byte[16 + result.Length];
        Array.Copy(aes.IV, 0, output, 0, 16);
        Array.Copy(result, 0, output, 16, result.Length);
        return output;
    }

    private static Byte[] RandomBytes(Int32 count)
    {
        var buf = new Byte[count];
        using var rng = RandomNumberGenerator.Create();
        rng.GetBytes(buf);
        return buf;
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

    private Byte[] EncryptAes256(Byte[] data, Int32 objNum, Int32 genNum)
    {
        // ISO 32000-2：AESV3 直接用文件密钥，随机 IV，PKCS#7 填充
        using var aes = Aes.Create();
        aes.Key = _key;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.GenerateIV();

        var iv = aes.IV;
        using var encryptor = aes.CreateEncryptor();
        using var ms = new MemoryStream();
        ms.Write(iv, 0, iv.Length);
        using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
        {
            cs.Write(data, 0, data.Length);
        }
        return ms.ToArray();
    }
    #endregion

    #region 辅助
    private Byte[] ObjKey(Int32 objNum, Int32 genNum)
    {
        // AESV3（/R 6）不使用对象密钥，直接使用 32 字节文件密钥
        if (Revision >= CipherRevision.Aes_256) return _key;

        // AESV2（/R 4）需要在密钥派生末尾追加 "sAlT"，否则解出的内容是乱码
        var salt = Revision == CipherRevision.Aes_128 ? new Byte[] { 0x73, 0x41, 0x6C, 0x54 } : [];

        var buf = new Byte[_key.Length + 5 + salt.Length];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write(_key);
        writer.Write((Byte)objNum);
        writer.Write((Byte)(objNum >> 8));
        writer.Write((Byte)(objNum >> 16));
        writer.Write((Byte)genNum);
        writer.Write((Byte)(genNum >> 8));
        if (salt.Length > 0) writer.Write(salt);

        var hash = ComputeMd5(buf);
        var keyLen = Math.Min(hash.Length, _key.Length + 5);
        var result = new Byte[keyLen];
        Array.Copy(hash, result, keyLen);
        return result;
    }

    private static Byte[] PadPwd(String pwd)
    {
        var raw = Encoding.GetEncoding(28591).GetBytes(pwd);
        var r = new Byte[32];
        var copyLen = Math.Min(raw.Length, 32);
        Array.Copy(raw, r, copyLen);
        Array.Copy(_padding, 0, r, copyLen, 32 - copyLen);
        return r;
    }

    private static Byte[] ComputeMd5(Byte[] data)
    {
        using var md5 = MD5.Create();
        return md5.ComputeHash(data);
    }

    private static Byte[] ComputeSha256(Byte[] data)
    {
        using var sha = SHA256.Create();
        return sha.ComputeHash(data);
    }

    private static Byte[] ComputeSha384(Byte[] data)
    {
        using var sha = SHA384.Create();
        return sha.ComputeHash(data);
    }

    private static Byte[] ComputeSha512(Byte[] data)
    {
        using var sha = SHA512.Create();
        return sha.ComputeHash(data);
    }

    private static Byte[] ComputeRc4(Byte[] key, Byte[] data)
    {
        var s = new Byte[256];
        for (var i = 0; i < 256; i++) s[i] = (Byte)i;
        var j = 0;
        for (var i = 0; i < 256; i++)
        {
            j = (j + s[i] + key[i % key.Length]) & 0xFF;
            var tmp = s[i]; s[i] = s[j]; s[j] = tmp;
        }
        var result = new Byte[data.Length];
        var x = 0; j = 0;
        for (var k = 0; k < data.Length; k++)
        {
            x = (x + 1) & 0xFF;
            j = (j + s[x]) & 0xFF;
            var tmp = s[x]; s[x] = s[j]; s[j] = tmp;
            result[k] = (Byte)(data[k] ^ s[(s[x] + s[j]) & 0xFF]);
        }
        return result;
    }
    #endregion
}