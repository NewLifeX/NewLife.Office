using System.Security.Cryptography;
using System.Text;

namespace NewLife.Office;

/// <summary>PDF 字体定义</summary>
public class PdfFont
{
    #region 属性
    /// <summary>字体资源名（如 F1）</summary>
    public String Name { get; }

    /// <summary>基础字体名（Type1 标准字体或嵌入 TrueType 名）</summary>
    public String BaseFont { get; }

    /// <summary>是否中文字体（使用 Identity-H 编码）</summary>
    public Boolean IsCjk { get; }
    #endregion

    #region 构造
    /// <summary>实例化字体</summary>
    /// <param name="name">资源名</param>
    /// <param name="baseFont">基础字体名</param>
    /// <param name="isCjk">是否中文字体</param>
    public PdfFont(String name, String baseFont, Boolean isCjk = false)
    {
        Name = name;
        BaseFont = baseFont;
        IsCjk = isCjk;
    }
    #endregion
}

/// <summary>PDF 页面对象（记录每页内容）</summary>
public class PdfPage
{
    #region 属性
    /// <summary>页面宽度（点，1 pt = 1/72 英寸）</summary>
    public Single Width { get; set; } = 595f; // A4

    /// <summary>页面高度（点）</summary>
    public Single Height { get; set; } = 842f; // A4

    /// <summary>内容流字节</summary>
    public Byte[] ContentBytes { get; set; } = [];

    /// <summary>此页引用的图片 XObject 名称→数据</summary>
    public Dictionary<String, (Byte[] Data, Int32 Width, Int32 Height, Boolean IsJpeg)> Images { get; } = [];

    /// <summary>页面旋转角度（0/90/180/270）</summary>
    public Int32 Rotation { get; set; } = 0;

    /// <summary>页面超链接注释列表（PDF 坐标：原点在左下角）</summary>
    public List<(Single X, Single Y, Single W, Single H, String Url)> LinkAnnotations { get; } = [];

    /// <summary>PDF 对象号（catalog=1, pages=2, page=3...）</summary>
    internal Int32 PageObjId { get; set; }

    /// <summary>内容流对象号</summary>
    internal Int32 ContentObjId { get; set; }
    #endregion
}

/// <summary>PDF 文档书签</summary>
public class PdfBookmark
{
    #region 属性
    /// <summary>书签标题</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>目标页面索引（0起始）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>子书签</summary>
    public List<PdfBookmark> Children { get; } = [];
    #endregion
}


/// <summary>PDF 元数据</summary>
public class PdfMetadata
{
    #region 属性
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>创建时间字符串（PDF 格式 D:YYYYMMDDHHmmss）</summary>
    public String? CreationDate { get; set; }

    /// <summary>PDF 版本（如 1.4）</summary>
    public String? PdfVersion { get; set; }

    /// <summary>总页数</summary>
    public Int32 PageCount { get; set; }
    #endregion
}

/// <summary>PDF 文本项，包含文本内容和近似坐标</summary>
/// <remarks>
/// 坐标系以页面左下角为原点，单位为 PDF 用户空间单位（通常约等于磅/pt）。
/// </remarks>
public class PdfTextItem
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>近似 X 坐标（PDF 用户空间单位）</summary>
    public Single X { get; set; }

    /// <summary>近似 Y 坐标（PDF 用户空间单位）</summary>
    public Single Y { get; set; }

    /// <summary>字体大小（通过 Tf 操作符获取，0 表示未知）</summary>
    public Single FontSize { get; set; }
    #endregion
}

/// <summary>PDF 嵌入图片流</summary>
public class PdfImageStream
{
    #region 属性
    /// <summary>图片在文档中的顺序索引（从 0 开始）</summary>
    public Int32 Index { get; set; }

    /// <summary>图片宽度（像素）</summary>
    public Int32 Width { get; set; }

    /// <summary>图片高度（像素）</summary>
    public Int32 Height { get; set; }

    /// <summary>编码过滤器名称，如 DCTDecode、FlateDecode 等</summary>
    public String Filter { get; set; } = String.Empty;

    /// <summary>原始流字节（未解压缩），对 JPEG 可直接使用</summary>
    public Byte[] RawData { get; set; } = [];

    /// <summary>是否为 JPEG（DCTDecode）图片，可直接将 RawData 保存为 .jpg</summary>
    public Boolean IsJpeg => Filter.IndexOf("DCTDecode", StringComparison.OrdinalIgnoreCase) >= 0;
    #endregion
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
        var fid = fileId.Length >= 16 ? fileId.Take(16).ToArray() : fileId;
        var buf = new List<Byte>(84);
        buf.AddRange(uPass);                                    // 32 字节：用户密码
        buf.AddRange(OEntry);                                   // 32 字节：O 条目
        buf.Add((Byte)permissions);                             // 4 字节：权限（小端）
        buf.Add((Byte)(permissions >> 8));
        buf.Add((Byte)(permissions >> 16));
        buf.Add((Byte)(permissions >> 24));
        buf.AddRange(fid);                                      // 16 字节：文件 ID
        var keyHash = ComputeMd5(buf.ToArray());
        for (var i = 0; i < 50; i++) keyHash = ComputeMd5(keyHash);
        _key = keyHash; // 16 字节

        // 算法 3.5：计算 U 条目（修订版 3）
        var uBuf = new List<Byte>(_padding);
        uBuf.AddRange(fid);
        var uStep = ComputeRc4(_key, ComputeMd5(uBuf.ToArray()));
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
        _key.CopyTo(buf, 0);
        buf[_key.Length] = (Byte)objNum;
        buf[_key.Length + 1] = (Byte)(objNum >> 8);
        buf[_key.Length + 2] = (Byte)(objNum >> 16);
        buf[_key.Length + 3] = (Byte)genNum;
        buf[_key.Length + 4] = (Byte)(genNum >> 8);
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
