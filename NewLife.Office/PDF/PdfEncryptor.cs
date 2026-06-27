using System.Security.Cryptography;
using System.Text;
using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>PDF 加密修订版本</summary>
public enum PdfCipherRevision
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
    public PdfCipherRevision Revision { get; }

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
    public PdfEncryptor(String? userPwd, String? ownerPwd, Int32 permissions, Byte[] fileId, PdfCipherRevision revision = PdfCipherRevision.Rc4_128)
    {
        EncPermissions = permissions;
        Revision = revision;
        var fid = fileId.Length >= 16 ? fileId.AsSpan(0, 16).ToArray() : fileId;

        switch (revision)
        {
            case PdfCipherRevision.Aes_256:
                (_key, OEntry, UEntry, OEEntry, UEEntry, PermsEntry) = ComputeAes256(userPwd, ownerPwd, permissions, fid);
                break;
            case PdfCipherRevision.Aes_128:
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
            PdfCipherRevision.Aes_128 => EncryptAes128(data, objNum, genNum),
            PdfCipherRevision.Aes_256 => EncryptAes256(data, objNum, genNum),
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
        var bytes = Encoding.GetEncoding(1252).GetBytes(sb.ToString());
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
        var uPass = PadPwd(userPwd ?? String.Empty);
        var oPass = PadPwd(ownerPwd ?? (userPwd ?? String.Empty));

        // 生成 256 位加密密钥（SHA-256）
        var key = ComputeSha256(uPass.Concat(fid).Concat(oPass).ToArray());

        // 计算 U 条目：SHA-256(userPassword + validationSalt) 前 32 字节
        var validationSalt = key.AsSpan(32, 8).ToArray();
        var uHash = ComputeSha256(uPass.Concat(validationSalt).ToArray());
        var u = new Byte[48];
        Array.Copy(uHash, 0, u, 0, 32);
        Array.Copy(validationSalt, 0, u, 32, 8);
        Array.Copy(key, 40, u, 40, 8); // keySalt

        // 计算 O 条目：SHA-256(ownerPassword + validationSalt + u) 前 32 字节
        var ownerHash = ComputeSha256(oPass.Concat(validationSalt).Concat(u).ToArray());
        var o = new Byte[48];
        Array.Copy(ownerHash, 0, o, 0, 32);
        Array.Copy(validationSalt, 0, o, 32, 8);
        Array.Copy(key, 40, o, 40, 8);

        // 计算 OE 条目：AES-256-CBC 加密 ownerPassword 的 SHA-256
        Byte[]? oe = null;
        if (!String.IsNullOrEmpty(ownerPwd))
        {
            var oeData = ComputeSha256(oPass);
            oe = EncryptAes256(key, oeData);
        }

        // 计算 UE 条目：AES-256-CBC 加密 userPassword 的 SHA-256
        Byte[]? ue = null;
        if (!String.IsNullOrEmpty(userPwd))
        {
            var ueData = ComputeSha256(uPass);
            ue = EncryptAes256(key, ueData);
        }

        // 计算 Perms 条目：AES-256-CBC 加密的权限字节
        var permsData = new Byte[16];
        var pw = new SpanWriter(permsData, 0, 16);
        pw.Write(permissions);
        pw.Write((Byte)0xFF); pw.Write((Byte)0xFF); pw.Write((Byte)0xFF); pw.Write((Byte)0xFF);
        pw.Write(fid[0] == 0 ? (Byte)0x54 : fid[0]); // 'T'
        pw.Write(fid[1] == 0 ? (Byte)0x61 : fid[1]); // 'a'
        pw.Write(fid[2] == 0 ? (Byte)0x64 : fid[2]); // 'd'
        pw.Write(fid[3] == 0 ? (Byte)0x62 : fid[3]); // 'b'
        var perms = EncryptAes256(key, permsData);

        return (key, o, u, oe, ue, perms);
    }

    private static Byte[] EncryptAes256(Byte[] key, Byte[] data)
    {
        using var aes = Aes.Create();
        aes.Key = key;
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

    private Byte[] EncryptAes256(Byte[] data, Int32 objNum, Int32 genNum)
    {
        // AES-256 内容加密：直接使用 _key 加密
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
        var buf = new Byte[_key.Length + 5];
        var writer = new SpanWriter(buf, 0, buf.Length);
        writer.Write(_key);
        writer.Write((Byte)objNum);
        writer.Write((Byte)(objNum >> 8));
        writer.Write((Byte)(objNum >> 16));
        writer.Write((Byte)genNum);
        writer.Write((Byte)(genNum >> 8));
        var hash = ComputeMd5(buf);
        var keyLen = Math.Min(hash.Length, _key.Length + 5);
        // AES-256 时密钥长度固定为 32
        if (Revision >= PdfCipherRevision.Aes_256) keyLen = 32;
        var result = new Byte[keyLen];
        Array.Copy(hash, result, keyLen);
        return result;
    }

    private static Byte[] PadPwd(String pwd)
    {
        var raw = Encoding.GetEncoding(1252).GetBytes(pwd);
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