using System.Text;
using NewLife.Buffers;
using NewLife.Office.Ole2;

namespace NewLife.Office.Ppt;

/// <summary>PowerPoint 97-2003 二进制（.ppt）演示文稿写入器</summary>
/// <remarks>
/// 生成 Microsoft PowerPoint 97-2003 二进制格式（MS-PPT）的 .ppt 文件，
/// 打包在 OLE2/CFB 容器中，零外部依赖。
/// <para>当前为最小可用版本：创建包含纯文本幻灯片的演示文稿，每张幻灯片可含多行文本。
/// 遵循 MS-PPT 记录结构（Document / SlideContainer / SlideListWithText / 持久化目录），
/// 产物可被 <see cref="PptReader"/> 读回，也可由 PowerPoint / WPS 打开。</para>
/// <para>写入示例：</para>
/// <code>
/// using var writer = new PptWriter();
/// writer.AddSlide("标题", "第一行", "第二行");
/// writer.AddSlide("第二张幻灯片");
/// writer.Save("slides.ppt");
/// </code>
/// </remarks>
public sealed class PptWriter : IDisposable
{
    #region 常量
    // MS-PPT 记录类型（见 [MS-PPT] 2.2 与 POI RecordTypes）
    private const UInt16 RecDocument = 0x03E8;              // DocumentContainer（持久化对象 1）
    private const UInt16 RecDocumentAtom = 0x03E9;          // DocumentAtom
    private const UInt16 RecEndDocument = 0x03EA;           // EndDocument
    private const UInt16 RecSlide = 0x03EE;                 // SlideContainer（持久化对象 2..N+1）
    private const UInt16 RecSlideAtom = 0x03EF;             // SlideAtom
    private const UInt16 RecEnvironment = 0x03F2;           // Environment
    private const UInt16 RecSlidePersistAtom = 0x03F3;      // SlidePersistAtom
    private const UInt16 RecSlideListWithText = 0x0FF0;     // SlideListWithText
    private const UInt16 RecTextHeaderAtom = 0x0F9F;        // TextHeaderAtom
    private const UInt16 RecTextCharsAtom = 0x0FA0;         // TextCharsAtom（UTF-16LE）
    private const UInt16 RecUserEditAtom = 0x0FF5;          // UserEditAtom
    private const UInt16 RecCurrentUserAtom = 0x0FF6;       // CurrentUserAtom（独立 "Current User" 流）
    private const UInt16 RecPersistPtrIncrementalBlock = 0x1772; // 持久化目录

    // 单个原子记录数据上限（MS-PPT 2.1.2：记录数据不得超过 8190 字节）
    private const Int32 MaxAtomDataSize = 8190 - 8;

    // 文本类型（TextHeaderAtom.textType）
    private const Int32 TextTypeTitle = 0;
    private const Int32 TextTypeBody = 2;
    #endregion

    #region 属性
    /// <summary>幻灯片宽度（EMU，默认 10 英寸 = 9144000）</summary>
    public Int32 SlideWidth { get; set; } = 9144000;

    /// <summary>幻灯片高度（EMU，默认 7.5 英寸 = 6858000）</summary>
    public Int32 SlideHeight { get; set; } = 6858000;

    /// <summary>最后编辑用户名（写入 "Current User" 流），null 使用默认值</summary>
    public String? Author { get; set; }

    /// <summary>幻灯片数量</summary>
    public Int32 SlideCount => _slides.Count;
    #endregion

    #region 私有字段
    private readonly List<List<String>> _slides = [];
    #endregion

    #region 构造
    /// <summary>实例化写入器</summary>
    public PptWriter() { }

    /// <summary>释放资源</summary>
    public void Dispose() { GC.SuppressFinalize(this); }
    #endregion

    #region 方法
    /// <summary>添加一张幻灯片（每行文本一个段落）</summary>
    /// <param name="texts">幻灯片文本行</param>
    public void AddSlide(params String[] texts) => _slides.Add([.. texts]);

    /// <summary>添加一张幻灯片（每行文本一个段落）</summary>
    /// <param name="texts">幻灯片文本行</param>
    public void AddSlide(IEnumerable<String> texts) => _slides.Add([.. texts]);

    /// <summary>保存为 .ppt 文件</summary>
    /// <param name="path">目标文件路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs);
    }

    /// <summary>将 .ppt 数据写入流（OLE2 容器）</summary>
    /// <param name="stream">可写输出流</param>
    public void Save(Stream stream)
    {
        var (ppt, userEditOffset) = BuildPptStream();
        var doc = new CfbDocument();
        doc.PutStream("PowerPoint Document", ppt);
        doc.PutStream("Current User", BuildCurrentUserAtom(userEditOffset));
        doc.Save(stream);
    }

    /// <summary>将 .ppt 数据序列化为字节数组</summary>
    /// <returns>.ppt 格式字节数组</returns>
    public Byte[] ToBytes()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }
    #endregion

    #region MS-PPT 流构建
    /// <summary>构建 "PowerPoint Document" 流字节及其 UserEditAtom 偏移</summary>
    /// <returns>（流字节，UserEditAtom 绝对偏移）</returns>
    private (Byte[] Stream, Int32 UserEditOffset) BuildPptStream()
    {
        var total = _slides.Count + 1;
        if (total > 4095)
            throw new InvalidOperationException($"持久化对象数量 {total} 超出单块持久化目录上限 4095。");

        // 1. Document（偏移 0）与各 SlideContainer，记录绝对偏移
        var doc = BuildDocument();
        var slides = new List<Byte[]>(_slides.Count);
        var slideOffsets = new Int32[_slides.Count];
        var offset = doc.Length;
        for (var i = 0; i < _slides.Count; i++)
        {
            slideOffsets[i] = offset;
            var slide = BuildSlide(_slides[i]);
            slides.Add(slide);
            offset += slide.Length;
        }

        // 2. UserEditAtom 固定 36 字节（8 头 + 28 体），随后紧跟持久化目录
        var userEditOffset = offset;
        var persistPtrOffset = userEditOffset + 36;
        var userEdit = BuildUserEditAtom(
            _slides.Count > 0 ? slideOffsets[^1] : 0,
            persistPtrOffset,
            total);
        var persistPtr = BuildPersistPtr(slideOffsets);

        using var ms = new MemoryStream();
        ms.Write(doc, 0, doc.Length);
        foreach (var s in slides) ms.Write(s, 0, s.Length);
        ms.Write(userEdit, 0, userEdit.Length);
        ms.Write(persistPtr, 0, persistPtr.Length);
        return (ms.ToArray(), userEditOffset);
    }

    /// <summary>构建 Document 容器（首子记录必须是 DocumentAtom）</summary>
    private Byte[] BuildDocument()
    {
        return Container(RecDocument,
            Atom(RecDocumentAtom, 1, BuildDocumentAtomData()),
            BuildOutlineSlideListWithText(),
            Container(RecEnvironment),
            Atom(RecEndDocument, 0, []));
    }

    /// <summary>DocumentAtom 数据（48 字节）</summary>
    private Byte[] BuildDocumentAtomData()
    {
        var buf = new Byte[48];
        var w = new SpanWriter(buf, 0, buf.Length);
        w.Write((UInt32)SlideWidth);   // slideSizeX
        w.Write((UInt32)SlideHeight);  // slideSizeY
        w.Write(6858000u);             // notesSizeX
        w.Write(9144000u);             // notesSizeY
        w.Write(10000u);               // serverZoomFrom（100%）
        w.Write(10000u);               // serverZoomTo
        w.Write(0u);                   // notesMasterPersist
        w.Write(0u);                   // handoutMasterPersist
        w.Write((UInt16)1);            // firstSlideNum
        w.Write((UInt16)0);            // slideSizeType（ON_SCREEN）
        // 4 布尔字节 + 8 保留字节（默认 0）
        return buf;
    }

    /// <summary>构建幻灯片容器（SlideAtom + 正文文本）</summary>
    /// <param name="lines">幻灯片文本行</param>
    private Byte[] BuildSlide(List<String> lines)
    {
        return Container(RecSlide,
            Atom(RecSlideAtom, 2, BuildSlideAtomData()),
            BuildSlideListWithText(lines));
    }

    /// <summary>SlideAtom 数据（24 字节）</summary>
    private static Byte[] BuildSlideAtomData()
    {
        var buf = new Byte[24];
        var w = new SpanWriter(buf, 0, buf.Length);
        w.Write(new Byte[12]);         // SSlideLayoutAtom（BLANK_SLIDE，全 0）
        w.Write(0x80000000u);          // masterID = USES_MASTER_SLIDE_ID
        w.Write(0u);                   // notesID
        w.Write((UInt16)0x0007);       // flags：跟随母版对象/配色/背景
        w.Write((UInt16)0);            // 保留
        return buf;
    }

    /// <summary>构建幻灯片自身的 SlideListWithText（BODY 文本，多行以 \r 分隔）</summary>
    /// <param name="lines">幻灯片文本行</param>
    private Byte[] BuildSlideListWithText(List<String> lines)
    {
        return Container(RecSlideListWithText,
            Atom(RecTextHeaderAtom, 0, I32(TextTypeBody)),
            BuildTextCharsAtom(String.Join("\r", lines)));
    }

    /// <summary>构建大纲 SlideListWithText：每张幻灯片一组 SlidePersistAtom + 标题/正文</summary>
    private Byte[] BuildOutlineSlideListWithText()
    {
        var parts = new List<Byte[]>();
        for (var i = 0; i < _slides.Count; i++)
        {
            var lines = _slides[i];
            parts.Add(BuildSlidePersistAtom(i + 2, i + 256));
            if (lines.Count > 0)
            {
                parts.Add(Atom(RecTextHeaderAtom, 0, I32(TextTypeTitle)));
                parts.Add(BuildTextCharsAtom(lines[0]));
            }
            if (lines.Count > 1)
            {
                parts.Add(Atom(RecTextHeaderAtom, 0, I32(TextTypeBody)));
                parts.Add(BuildTextCharsAtom(String.Join("\r", lines.Skip(1))));
            }
        }
        return Container(RecSlideListWithText, [.. parts]);
    }

    /// <summary>SlidePersistAtom 数据（20 字节）</summary>
    /// <param name="persistId">持久化对象 ID（2 起始）</param>
    /// <param name="slideIdentifier">幻灯片内部标识（256 起始）</param>
    private static Byte[] BuildSlidePersistAtom(Int32 persistId, Int32 slideIdentifier)
    {
        var buf = new Byte[20];
        var w = new SpanWriter(buf, 0, buf.Length);
        w.Write((UInt32)persistId);        // refID：持久化 ID
        w.Write(4u);                       // flags：HAS_SHAPES_OTHER_THAN_PLACEHOLDERS
        w.Write(1u);                       // numPlaceholderTexts
        w.Write((UInt32)slideIdentifier);  // slideIdentifier（256+）
        // 4 字节保留（默认 0）
        return buf;
    }

    /// <summary>TextCharsAtom（UTF-16LE），超长自动拆分为续记录（recInstance=0xF）</summary>
    /// <param name="text">文本内容</param>
    private static Byte[] BuildTextCharsAtom(String text)
    {
        var bytes = Encoding.Unicode.GetBytes(text);
        if (bytes.Length <= MaxAtomDataSize)
            return Atom(RecTextCharsAtom, 0, bytes);

        var parts = new List<Byte[]>();
        var offset = 0;
        var first = true;
        while (offset < bytes.Length)
        {
            var chunk = Math.Min(MaxAtomDataSize, bytes.Length - offset);
            var body = new Byte[chunk];
            Array.Copy(bytes, offset, body, 0, chunk);
            parts.Add(Atom(RecTextCharsAtom, first ? (UInt16)0 : (UInt16)0x0F, body));
            offset += chunk;
            first = false;
        }
        return Concat([.. parts]);
    }

    /// <summary>UserEditAtom 数据（28 字节）</summary>
    /// <param name="lastSlideOffset">最后一张幻灯片的绝对偏移</param>
    /// <param name="persistPtrOffset">持久化目录的绝对偏移</param>
    /// <param name="maxPersist">持久化对象总数（含 Document）</param>
    private static Byte[] BuildUserEditAtom(Int32 lastSlideOffset, Int32 persistPtrOffset, Int32 maxPersist)
    {
        var buf = new Byte[28];
        var w = new SpanWriter(buf, 0, buf.Length);
        w.Write((UInt32)lastSlideOffset);   // lastViewedSlideID：最后幻灯片偏移
        w.Write(0x00000A02u);               // pptVersion
        w.Write(0u);                        // lastUserEditAtomOffset（首次编辑为 0）
        w.Write((UInt32)persistPtrOffset);  // persistPointersOffset
        w.Write(0u);                        // docPersistRef（Document 位于偏移 0）
        w.Write((UInt32)maxPersist);        // maxPersistWritten
        w.Write((UInt16)1);                 // lastViewType（SLIDE_VIEW）
        w.Write((UInt16)0);                 // unused
        return buf;
    }

    /// <summary>持久化目录（PersistPtrIncrementalBlock）：Document(1) + 幻灯片(2..N+1)</summary>
    /// <param name="slideOffsets">各幻灯片绝对偏移</param>
    private static Byte[] BuildPersistPtr(Int32[] slideOffsets)
    {
        var count = slideOffsets.Length + 1;
        var buf = new Byte[4 + count * 4];
        var w = new SpanWriter(buf, 0, buf.Length);
        // info 块：低 20 位起始持久化 ID（1），高 12 位条目数
        w.Write((UInt32)((count << 20) | 1));
        w.Write(0u); // Document 偏移 0
        for (var i = 0; i < slideOffsets.Length; i++)
            w.Write((UInt32)slideOffsets[i]);
        return Atom(RecPersistPtrIncrementalBlock, (UInt16)count, buf);
    }

    /// <summary>构建 "Current User" 流（CurrentUserAtom，独立 OLE2 流）</summary>
    /// <param name="userEditOffset">PowerPoint Document 流中 UserEditAtom 的绝对偏移</param>
    private Byte[] BuildCurrentUserAtom(Int32 userEditOffset)
    {
        var name = Author ?? "NewLife.Office";
        var ascii = Encoding.ASCII.GetBytes(name);
        var unicode = Encoding.Unicode.GetBytes(name);

        using var ms = new MemoryStream();
        ms.WriteByte(0x00);                          // recVer
        ms.WriteByte(0x00);                          // recInstance
        ms.Write(I16(RecCurrentUserAtom), 0, 2);     // recType = 0x0FF6
        ms.Write(I32(24 + ascii.Length), 0, 4);      // recLen
        ms.Write(I32(20), 0, 4);                     // size（细节区大小）
        ms.Write(I32(unchecked((Int32)0xE391C05F)), 0, 4); // headerToken（非加密魔数）
        ms.Write(I32(userEditOffset), 0, 4);         // currentEditOffset
        ms.Write(I16(ascii.Length), 0, 2);           // 用户名长度（ASCII）
        ms.Write(I16(0x03F4), 0, 2);                 // docFinalVersion
        ms.WriteByte(3);                             // docMajorNo
        ms.WriteByte(0);                             // docMinorNo
        ms.Write(I16(0), 0, 2);                      // 保留
        ms.Write(ascii, 0, ascii.Length);            // ASCII 用户名
        ms.Write(I32(8), 0, 4);                      // releaseVersion
        ms.Write(unicode, 0, unicode.Length);        // Unicode 用户名
        return ms.ToArray();
    }
    #endregion

    #region 记录构建辅助
    /// <summary>构建标准 8 字节记录头 + 数据</summary>
    /// <param name="recVer">记录版本（容器 0xF，原子 0）</param>
    /// <param name="recInstance">记录实例</param>
    /// <param name="recType">记录类型</param>
    /// <param name="data">记录数据</param>
    private static Byte[] Record(UInt16 recVer, UInt16 recInstance, UInt16 recType, Byte[] data)
    {
        var buf = new Byte[8 + data.Length];
        var w = new SpanWriter(buf, 0, buf.Length);
        w.Write((UInt16)(((recInstance & 0x0FFF) << 4) | (recVer & 0x0F)));
        w.Write(recType);
        w.Write((UInt32)data.Length);
        if (data.Length > 0) Array.Copy(data, 0, buf, 8, data.Length);
        return buf;
    }

    /// <summary>容器记录（recVer = 0xF）</summary>
    /// <param name="recType">记录类型</param>
    /// <param name="children">子记录</param>
    private static Byte[] Container(UInt16 recType, params Byte[][] children)
        => Record(0x0F, 0, recType, Concat(children));

    /// <summary>原子记录（recVer = 0）</summary>
    /// <param name="recType">记录类型</param>
    /// <param name="recInstance">记录实例</param>
    /// <param name="data">记录数据</param>
    private static Byte[] Atom(UInt16 recType, UInt16 recInstance, Byte[] data)
        => Record(0, recInstance, recType, data);

    /// <summary>Int32 → 4 字节小端</summary>
    private static Byte[] I32(Int32 v)
    {
        var buf = new Byte[4];
        new SpanWriter(buf).Write((UInt32)v);
        return buf;
    }

    /// <summary>Int32 → 2 字节小端</summary>
    private static Byte[] I16(Int32 v)
    {
        var buf = new Byte[2];
        new SpanWriter(buf).Write((UInt16)v);
        return buf;
    }

    /// <summary>拼接字节数组</summary>
    private static Byte[] Concat(params Byte[][] parts)
    {
        var total = 0;
        foreach (var p in parts) total += p.Length;
        var buf = new Byte[total];
        var pos = 0;
        foreach (var p in parts)
        {
            Array.Copy(p, 0, buf, pos, p.Length);
            pos += p.Length;
        }
        return buf;
    }
    #endregion
}
