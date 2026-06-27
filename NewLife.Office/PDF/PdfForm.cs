namespace NewLife.Office;

/// <summary>PDF AcroForm 表单</summary>
/// <remarks>
/// 描述交互式表单结构，支持文本框、复选框、单选按钮、选择框和签名字段。
/// 通过 PdfWriter 写入表单字典和字段树，通过 PdfReader 读取表单字段和当前值。
/// </remarks>
public class PdfForm
{
    #region 属性
    /// <summary>表单字段列表</summary>
    public List<PdfFormField> Fields { get; set; } = [];

    /// <summary>是否需要外观流（/NeedAppearances），默认 true</summary>
    public Boolean NeedAppearances { get; set; } = true;
    #endregion
}

/// <summary>PDF 表单字段类型</summary>
public enum PdfFormFieldType
{
    /// <summary>按钮（含复选框、单选按钮）</summary>
    Btn,

    /// <summary>文本框</summary>
    Tx,

    /// <summary>选择框（下拉列表/列表框）</summary>
    Ch,

    /// <summary>签名字段</summary>
    Sig,
}

/// <summary>PDF 表单字段标志（位掩码）</summary>
[Flags]
public enum PdfFormFieldFlags
{
    /// <summary>无标志</summary>
    None = 0,

    /// <summary>只读</summary>
    ReadOnly = 1 << 0,

    /// <summary>必填</summary>
    Required = 1 << 1,

    /// <summary>不可导出</summary>
    NoExport = 1 << 2,

    // ── 文本字段特有 ──
    /// <summary>多行</summary>
    Multiline = 1 << 12,

    /// <summary>密码（显示为星号）</summary>
    Password = 1 << 13,

    /// <summary>文件选择</summary>
    FileSelect = 1 << 20,

    /// <summary>不拼写检查</summary>
    DoNotSpellCheck = 1 << 22,

    /// <summary>不滚动</summary>
    DoNotScroll = 1 << 23,

    /// <summary>梳状（等宽字符间距）</summary>
    Comb = 1 << 24,

    /// <summary>富文本</summary>
    RichText = 1 << 25,

    // ── 按钮字段特有 ──
    /// <summary>无切换关闭（单选按钮必选一项）</summary>
    NoToggleToOff = 1 << 14,

    /// <summary>单选按钮</summary>
    Radio = 1 << 15,

    /// <summary>按下时执行</summary>
    Pushbutton = 1 << 16,

    // ── 选择字段特有 ──
    /// <summary>支持多选（复选框列表）</summary>
    MultiSelect = 1 << 21,

    /// <summary>提交时包含选中值</summary>
    CommitOnSelChange = 1 << 26,
}

/// <summary>PDF 表单字段</summary>
public class PdfFormField
{
    #region 属性
    /// <summary>字段完全限定名（如 "form1[0].#subform[0].txtName[0]"）</summary>
    public String FullName { get; set; } = String.Empty;

    /// <summary>字段类型</summary>
    public PdfFormFieldType FieldType { get; set; } = PdfFormFieldType.Tx;

    /// <summary>字段值</summary>
    public String? Value { get; set; }

    /// <summary>默认值</summary>
    public String? DefaultValue { get; set; }

    /// <summary>字段标志</summary>
    public PdfFormFieldFlags Flags { get; set; }

    /// <summary>页面索引（0起始）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>字段矩形区域（PDF 坐标，从左下角量起，单位磅）</summary>
    public Single X { get; set; }

    /// <summary>字段矩形区域下边界</summary>
    public Single Y { get; set; }

    /// <summary>字段宽度</summary>
    public Single Width { get; set; } = 100f;

    /// <summary>字段高度</summary>
    public Single Height { get; set; } = 18f;

    /// <summary>字体名称（默认 Helvetica）</summary>
    public String? FontName { get; set; }

    /// <summary>字号（默认 12）</summary>
    public Single FontSize { get; set; } = 12f;

    /// <summary>最大字符数（0=无限制）</summary>
    public Int32 MaxLength { get; set; }

    /// <summary>工具提示文本</summary>
    public String? Tooltip { get; set; }

    /// <summary>下拉/列表选项（仅 Ch 类型）</summary>
    public List<String> Options { get; set; } = [];

    /// <summary>子字段（如单选按钮组内的各个选项）</summary>
    public List<PdfFormField> Kids { get; set; } = [];
    #endregion
}
