namespace NewLife.Office;

/// <summary>文档属性</summary>
public class WordDocumentProperties
{
    #region 属性
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>描述</summary>
    public String? Description { get; set; }

    /// <summary>自定义属性集合（Key=属性名, Value=属性值和类型）</summary>
    /// <remarks>
    /// 自定义属性写入 docProps/custom.xml，支持的数据类型：
    /// <c>lpwstr</c>（字符串）、<c>i4</c>（整数）、<c>r8</c>（浮点）、<c>bool</c>（布尔）、<c>date</c>（日期）
    /// </remarks>
    public Dictionary<String, (String Value, String Type)> CustomProperties { get; set; } = [];
    #endregion
}
