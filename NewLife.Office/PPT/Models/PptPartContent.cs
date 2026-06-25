namespace NewLife.Office;

/// <summary>PPT 文档片段原始内容（母版/版式等）</summary>
/// <remarks>以原始 XML 形式存储，避免完整模型化 OOXML schema 带来的维护负担。</remarks>
internal sealed class PptPartContent
{
    #region 属性
    /// <summary>部件 XML 内容</summary>
    public String Xml { get; set; } = String.Empty;

    /// <summary>部件关系 XML 内容</summary>
    public String RelsXml { get; set; } = String.Empty;
    #endregion
}
