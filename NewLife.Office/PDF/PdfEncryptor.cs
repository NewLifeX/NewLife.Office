using System.Security.Cryptography;
using System.Text;
using NewLife.Buffers;

namespace NewLife.Office;

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

    private readonly Byte[] _key; // 128 位全局密钥（MD5 输出，16 字节）

    /// <summary>Owner 密钥条目（32 字节，写入加密字典 /O）</summary>
    public Byte[] OEntry { get; }

    /// <summary>User 密钥条目（32 字节，写入加密字典 /U）</summary>
    public Byte[] UEntry { get; }

    /// <summary>加密权限标志（写入加密字典 /P）</summary>
    public Int32 EncPermissions { get; }
    #endregion

    #region 构造
    /// <summary>实例化 PDF 加密器，按 PDF 1.4 算法 3.2/3.3/3.5 计算密钥和授权条目</summary>
    /// <param name="userPwd">用户密码（打开密码），null 表示空密码</param>
    /// <param name="ownerPwd">所有者密码（权限密码）</param>
    /// <param name="permissions">权限标志位（PDF 规范 Table 3.20）</param>
    /// <param name="fileId">文件标识符（16 字节 MD5）</param>
    public PdfEncryptor(String? userPwd, String? ownerPwd, Int32 permissions, Byte[] fileId)
    {
        EncPermissions = permissions;
        var uPass = PadPwd(userPwd ?? String.Empty);
        var oPass = PadPwd(ownerPwd ?? (userPwd ?? String.Empty));

        // 算法 3.3：计算 O 条目（修订版 3）
        var ownerKey = ComputeMd5(oPass);
        for (var i = 0; i < 50; i++) ownerKey = ComputeMd5(ownerKey);
        var oStep = ComputeRc4(ownerKey, uPass);
        for (var i = 1; i <= 19; i++)
        {
            var k = new Byte[ownerKey.Length];
            for (var j = 0; j < k.Length; j++) k[j] = (Byte)(ownerKey[j] ^ i);
            oStep = ComputeRc4(k, oStep);
        }
        OEntry = oStep; // 32 字节

        // 算法 3.2：计算全局加密密钥
        var fid = fileId.Length >= 16 ? fileId.AsSpan(0, 16).ToArray() : fileId;
        var buf = new Byte[32 + 32 + 4 + fid.Length];
        var bw = new SpanWriter(buf, 0, buf.Length);
        bw.Write(uPass);                                        // 32 字节：用户密码
        bw.Write(OEntry);                                       // 32 字节：O 条目
        bw.Write(permissions);                                  // 4 字节：权限（小端）
        bw.Write(fid);                                          // 文件 ID
        var keyHash = ComputeMd5(buf);
        for (var i = 0; i < 50; i++) keyHash = ComputeMd5(keyHash);
        _key = keyHash; // 16 字节

        // 算法 3.5：计算 U 条目（修订版 3）
        var uBuf = new Byte[_padding.Length + fid.Length];
        var ubw = new SpanWriter(uBuf, 0, uBuf.Length);
        ubw.Write(_padding);
        ubw.Write(fid);
        var uStep = ComputeRc4(_key, ComputeMd5(uBuf));
        for (var i = 1; i <= 19; i++)
        {
            var k = new Byte[_key.Length];
            for (var j = 0; j < k.Length; j++) k[j] = (Byte)(_key[j] ^ i);
            uStep = ComputeRc4(k, uStep);
        }
        UEntry = new Byte[32];
        Array.Copy(uStep, UEntry, uStep.Length);
    }
    #endregion

    #region 方法
    /// <summary>加密字节数组（RC4，基于对象号派生子密钥，算法 3.1）</summary>
    /// <param name="data">原始字节</param>
    /// <param name="objNum">PDF 对象号</param>
    /// <param name="genNum">PDF 代数号</param>
    /// <returns>加密后字节（长度与原始相同）</returns>
    public Byte[] EncryptBytes(Byte[] data, Int32 objNum, Int32 genNum) => ComputeRc4(ObjKey(objNum, genNum), data);

    /// <summary>加密字符串，返回 PDF 十六进制字符串格式 &lt;hex&gt;</summary>
    /// <param name="s">待加密文本（非 Latin-1 字符自动替换为 ?）</param>
    /// <param name="objNum">PDF 对象号</param>
    /// <param name="genNum">PDF 代数号</param>
    /// <returns>十六进制字符串，格式如 &lt;AABB...&gt;</returns>
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