namespace NewLife.Office.Markdown;

/// <summary>文档转换错误类型（对标 anydoc ConvertError 变体）</summary>
public enum ConvertErrorType
{
    /// <summary>未知格式或无法转换的格式（如纯图 PDF）</summary>
    Unsupported,

    /// <summary>结构不可用，无法提取有意义内容</summary>
    Malformed,

    /// <summary>加密或密码保护</summary>
    Encrypted,

    /// <summary>超出固定安全限制（解压/嵌套/节点数）</summary>
    ResourceLimit,

    /// <summary>缺少产出有意义输出所必需的部件</summary>
    MissingPart,

    /// <summary>文件无法读取（仅文件入口）</summary>
    Io,
}

/// <summary>文档转换异常（对标 anydoc ConvertError）</summary>
/// <remarks>
/// 仅当无法产出有意义 Markdown 时抛出：输入不可读、结构不可用、加密、
/// 或超出固定安全限制。可恢复的文档瑕疵应跳过并继续转换，不抛此异常。
/// </remarks>
public class ConvertErrorException : Exception
{
    /// <summary>错误类型</summary>
    public ConvertErrorType Type { get; }

    /// <summary>实例化</summary>
    /// <param name="type">错误类型</param>
    /// <param name="message">错误消息</param>
    public ConvertErrorException(ConvertErrorType type, String message) : base(message) => Type = type;

    /// <summary>实例化</summary>
    /// <param name="type">错误类型</param>
    /// <param name="message">错误消息</param>
    /// <param name="innerException">内部异常</param>
    public ConvertErrorException(ConvertErrorType type, String message, Exception innerException) : base(message, innerException) => Type = type;

    /// <summary>创建 Unsupported 异常</summary>
    /// <param name="message">错误消息</param>
    /// <returns>异常对象</returns>
    public static ConvertErrorException Unsupported(String message) => new(ConvertErrorType.Unsupported, message);

    /// <summary>创建 Malformed 异常</summary>
    /// <param name="message">错误消息</param>
    /// <returns>异常对象</returns>
    public static ConvertErrorException Malformed(String message) => new(ConvertErrorType.Malformed, message);

    /// <summary>创建 ResourceLimit 异常</summary>
    /// <param name="message">错误消息</param>
    /// <returns>异常对象</returns>
    public static ConvertErrorException ResourceLimit(String message) => new(ConvertErrorType.ResourceLimit, message);
}
