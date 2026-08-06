using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml.Linq;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>Excel 线程化批注测试（M27，对标 EPPlus/Aspose.Cells）</summary>
/// <remarks>
/// 覆盖线程化批注的写入（personList + commentList 部件 + sheet 引用）、
/// 读取还原（作者/时间/正文/回复关系）、同一作者共享 person 与参数校验。
/// </remarks>
public class ThreadedCommentTests
{
    /// <summary>构造含线程化批注的 xlsx 字节</summary>
    private static Byte[] BuildXlsx()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.SheetName = "Sheet1";
            var top = writer.AddThreadedComment("Sheet1", 1, 0, "这是顶级批注", "张三", "zhangsan@example.com");
            writer.ReplyThreadedComment(top, "回复一：同意", "李四", "lisi@example.com");
            writer.ReplyThreadedComment(top, "回复二：补充说明", "张三", "zhangsan@example.com");
            writer.AddThreadedComment("Sheet1", 2, 0, "第二个单元格批注", "王五", null, time: new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc));
            writer.Save();
        }
        return ms.ToArray();
    }

    [Fact, DisplayName("线程化批注：写入与读取还原（含回复线程）")]
    public void ThreadedComment_WriteRead()
    {
        var bytes = BuildXlsx();

        using var reader = new ExcelReader(new MemoryStream(bytes), Encoding.UTF8);
        var comments = reader.ReadThreadedComments("Sheet1");

        Assert.Equal(4, comments.Count);

        // 顶级批注
        var top = comments.FirstOrDefault(c => c.Row == 1 && c.Col == 0 && c.ParentId == null);
        Assert.NotNull(top);
        Assert.Equal("这是顶级批注", top!.Text);
        Assert.Equal("张三", top.Author);
        Assert.Equal("zhangsan@example.com", top.UserId);
        Assert.False(String.IsNullOrEmpty(top.Id));

        // 回复
        var replies = comments.Where(c => c.ParentId == top.Id).ToList();
        Assert.Equal(2, replies.Count);
        Assert.Contains(replies, r => r.Author == "李四" && r.Text == "回复一：同意");
        Assert.Contains(replies, r => r.Author == "张三" && r.Text == "回复二：补充说明");

        // 第二单元格批注时间还原
        var second = comments.FirstOrDefault(c => c.Row == 2 && c.Col == 0);
        Assert.NotNull(second);
        Assert.Equal("王五", second!.Author);
        Assert.Equal(2024, second.Time.Year);
    }

    [Fact, DisplayName("线程化批注：同一作者共享 person 且部件齐全")]
    public void ThreadedComment_SameAuthorSharedPerson()
    {
        var bytes = BuildXlsx();

        // 用 ZipArchive 流解析
        var ms = new MemoryStream(bytes);
        using var za = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read);

        // person.xml 存在
        var personEntry = za.GetEntry("xl/persons/person.xml");
        Assert.NotNull(personEntry);
        using (var ps = personEntry!.Open())
        {
            var doc = XDocument.Load(ps);
            Assert.NotNull(doc.Root);
            // 张三、李四、王五 3 个不同 person
            Assert.Equal(3, doc.Root!.Elements().Count(e => e.Name.LocalName == "person"));
        }

        // threadedComment 部件存在且含 4 条
        var tcEntry = za.GetEntry("xl/threadedComments/threadedComment1.xml");
        Assert.NotNull(tcEntry);
        using (var cs = tcEntry!.Open())
        {
            var doc = XDocument.Load(cs);
            Assert.NotNull(doc.Root);
            Assert.Equal(4, doc.Root!.Elements().Count(e => e.Name.LocalName == "threadedComment"));
        }

        // sheet 引用：extLst + x14:commentList
        var sheetEntry = za.GetEntry("xl/worksheets/sheet1.xml");
        Assert.NotNull(sheetEntry);
        using (var ss = sheetEntry!.Open())
        {
            var text = new StreamReader(ss, Encoding.UTF8).ReadToEnd();
            Assert.Contains("x14:commentList", text);
            Assert.Contains("t=\"threaded\"", text);
        }

        // sheet rels 含 threadedComment 关系
        var relEntry = za.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
        Assert.NotNull(relEntry);
        using (var rs = relEntry!.Open())
        {
            var text = new StreamReader(rs, Encoding.UTF8).ReadToEnd();
            Assert.Contains("threadedComment", text);
        }
    }

    [Fact, DisplayName("线程化批注：参数校验")]
    public void ThreadedComment_Validation()
    {
        using var ms = new MemoryStream();
        using var writer = new ExcelWriter(ms);
        writer.SheetName = "Sheet1";

        Assert.Throws<ArgumentNullException>(() => writer.AddThreadedComment("Sheet1", 1, 0, "", "作者"));
        Assert.Throws<ArgumentNullException>(() => writer.AddThreadedComment("Sheet1", 1, 0, "文本", ""));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.AddThreadedComment("Sheet1", 0, 0, "文本", "作者"));
    }
}
