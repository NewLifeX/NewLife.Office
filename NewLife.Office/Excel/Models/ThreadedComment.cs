using System;

namespace NewLife.Office.Excel;

/// <summary>线程化批注（Threaded Comment，M27，对标 EPPlus/Aspose.Cells）</summary>
/// <remarks>
/// Excel 2016+ 的对话式批注，支持多用户回复形成对话线程。
/// 与普通批注（M26）不同，线程化批注使用独立的 threadedComment 部件与 person 部件，
/// 每个批注有全局唯一 Id（关联 sheet 中 comment 元素的 guid），回复通过 ParentId 形成树。
/// </remarks>
public class ThreadedComment
{
    /// <summary>行号（1-based）</summary>
    public Int32 Row { get; set; }

    /// <summary>列号（0-based）</summary>
    public Int32 Col { get; set; }

    /// <summary>批注正文</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>作者显示名（写入 person.xml）</summary>
    public String Author { get; set; } = String.Empty;

    /// <summary>作者用户标识（邮箱/账号，写入 person.xml userId）</summary>
    public String? UserId { get; set; }

    /// <summary>批注时间（UTC）</summary>
    public DateTime Time { get; set; }

    /// <summary>父批注 Id（回复时设置，顶级批注为 null）</summary>
    public String? ParentId { get; set; }

    /// <summary>批注唯一 Id（GUID 花括号格式，关联 sheet comment guid）</summary>
    public String Id { get; set; } = Guid.NewGuid().ToString("B").ToUpper();

    /// <summary>关联 person 的 Id（GUID 花括号格式）</summary>
    public String PersonId { get; set; } = Guid.NewGuid().ToString("B").ToUpper();

    /// <summary>所属工作表名称</summary>
    public String Sheet { get; set; } = String.Empty;

    /// <summary>快速创建顶级批注</summary>
    /// <param name="sheet">工作表名称</param>
    /// <param name="row">行号（1-based）</param>
    /// <param name="col">列号（0-based）</param>
    /// <param name="text">批注正文</param>
    /// <param name="author">作者显示名</param>
    /// <param name="userId">作者用户标识</param>
    /// <returns>线程化批注实例</returns>
    public static ThreadedComment Create(String sheet, Int32 row, Int32 col, String text, String author, String? userId = null)
    {
        return new ThreadedComment
        {
            Sheet = sheet,
            Row = row,
            Col = col,
            Text = text,
            Author = author,
            UserId = userId,
            Time = DateTime.UtcNow,
        };
    }

    /// <summary>创建回复批注（关联父批注 Id）</summary>
    /// <param name="parent">父批注</param>
    /// <param name="text">回复正文</param>
    /// <param name="author">作者显示名</param>
    /// <param name="userId">作者用户标识</param>
    /// <returns>线程化批注实例</returns>
    public static ThreadedComment Reply(ThreadedComment parent, String text, String author, String? userId = null)
    {
        return new ThreadedComment
        {
            Sheet = parent.Sheet,
            Row = parent.Row,
            Col = parent.Col,
            Text = text,
            Author = author,
            UserId = userId,
            Time = DateTime.UtcNow,
            ParentId = parent.Id,
        };
    }
}
