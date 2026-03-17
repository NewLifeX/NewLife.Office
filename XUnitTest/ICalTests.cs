using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>iCal 日历格式读写测试</summary>
public class ICalTests
{
    #region 辅助

    private static String BuildSimpleIcs(String summary, String dtStart, String dtEnd)
    {
        return "BEGIN:VCALENDAR\r\n"
             + "VERSION:2.0\r\n"
             + "PRODID:-//Test//Test//EN\r\n"
             + "BEGIN:VEVENT\r\n"
             + "UID:test-uid-001\r\n"
             + $"SUMMARY:{summary}\r\n"
             + $"DTSTART:{dtStart}\r\n"
             + $"DTEND:{dtEnd}\r\n"
             + "END:VEVENT\r\n"
             + "END:VCALENDAR\r\n";
    }

    #endregion

    #region 读取测试

    [Fact]
    [DisplayName("解析单个 VEVENT")]
    public void Parse_SingleEvent_Read()
    {
        var ics = BuildSimpleIcs("Team Meeting", "20240601T100000Z", "20240601T110000Z");
        var reader = new ICalReader();
        var doc = reader.Parse(ics);

        Assert.Single(doc.Events);
        var evt = doc.Events[0];
        Assert.Equal("Team Meeting", evt.Summary);
        Assert.Equal("test-uid-001", evt.Uid);
        Assert.NotNull(evt.Start);
        Assert.NotNull(evt.End);
        Assert.Equal(2024, evt.Start!.Value.Year);
        Assert.Equal(6, evt.Start!.Value.Month);
    }

    [Fact]
    [DisplayName("解析全天事件（VALUE=DATE）")]
    public void Parse_AllDayEvent_AllDayFlagSet()
    {
        var ics = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:allday-001\r\n"
                + "SUMMARY:Birthday\r\nDTSTART;VALUE=DATE:20240715\r\nDTEND;VALUE=DATE:20240716\r\n"
                + "END:VEVENT\r\nEND:VCALENDAR\r\n";

        var reader = new ICalReader();
        var doc = reader.Parse(ics);

        Assert.Single(doc.Events);
        Assert.True(doc.Events[0].AllDay);
        Assert.Equal(7, doc.Events[0].Start!.Value.Month);
        Assert.Equal(15, doc.Events[0].Start!.Value.Day);
    }

    [Fact]
    [DisplayName("解析多个 VEVENT")]
    public void Parse_MultipleEvents_AllRead()
    {
        var ics = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\n"
                + "BEGIN:VEVENT\r\nUID:e1\r\nSUMMARY:Event1\r\nDTSTART:20240101T090000Z\r\nEND:VEVENT\r\n"
                + "BEGIN:VEVENT\r\nUID:e2\r\nSUMMARY:Event2\r\nDTSTART:20240201T090000Z\r\nEND:VEVENT\r\n"
                + "BEGIN:VEVENT\r\nUID:e3\r\nSUMMARY:Event3\r\nDTSTART:20240301T090000Z\r\nEND:VEVENT\r\n"
                + "END:VCALENDAR\r\n";

        var reader = new ICalReader();
        var doc = reader.Parse(ics);

        Assert.Equal(3, doc.Events.Count);
        Assert.Equal("Event1", doc.Events[0].Summary);
        Assert.Equal("Event3", doc.Events[2].Summary);
    }

    [Fact]
    [DisplayName("解析 VTODO")]
    public void Parse_Todo_Read()
    {
        var ics = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\n"
                + "BEGIN:VTODO\r\nUID:todo-001\r\nSUMMARY:Write report\r\n"
                + "STATUS:NEEDS-ACTION\r\nPRIORITY:2\r\nPERCENT-COMPLETE:30\r\n"
                + "DUE:20240630T170000Z\r\nEND:VTODO\r\n"
                + "END:VCALENDAR\r\n";

        var reader = new ICalReader();
        var doc = reader.Parse(ics);

        Assert.Single(doc.Todos);
        var todo = doc.Todos[0];
        Assert.Equal("Write report", todo.Summary);
        Assert.Equal("NEEDS-ACTION", todo.Status);
        Assert.Equal(2, todo.Priority);
        Assert.Equal(30, todo.PercentComplete);
    }

    [Fact]
    [DisplayName("解析转义文本")]
    public void Parse_EscapedText_UnescapedCorrectly()
    {
        var ics = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\n"
                + "BEGIN:VEVENT\r\nUID:esc-001\r\n"
                + @"SUMMARY:Meeting\, Room 101" + "\r\n"
                + @"DESCRIPTION:Line1\nLine2" + "\r\n"
                + "DTSTART:20240101T090000Z\r\nEND:VEVENT\r\n"
                + "END:VCALENDAR\r\n";

        var reader = new ICalReader();
        var doc = reader.Parse(ics);

        Assert.Equal("Meeting, Room 101", doc.Events[0].Summary);
        Assert.Contains("\n", doc.Events[0].Description);
    }

    [Fact]
    [DisplayName("解析折行属性（RFC 5545 行折叠）")]
    public void Parse_FoldedLines_Unfolded()
    {
        var ics = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\n"
                + "BEGIN:VEVENT\r\nUID:fold-001\r\n"
                + "SUMMARY:This is a very long summa\r\n ry that spans multiple lines\r\n"
                + "DTSTART:20240101T090000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n";

        var reader = new ICalReader();
        var doc = reader.Parse(ics);

        Assert.Equal("This is a very long summary that spans multiple lines", doc.Events[0].Summary);
    }

    #endregion

    #region 写入测试

    [Fact]
    [DisplayName("写入包含基本 VEVENT 的日历")]
    public void Write_BasicEvent_ContainsRequiredFields()
    {
        var doc = new ICalDocument { CalendarName = "My Calendar" };
        doc.Events.Add(new ICalEvent
        {
            Uid = "write-test-001",
            Summary = "Sprint Review",
            Start = new DateTimeOffset(2024, 6, 1, 14, 0, 0, TimeSpan.Zero),
            End = new DateTimeOffset(2024, 6, 1, 15, 0, 0, TimeSpan.Zero),
            Location = "Conference Room A",
        });

        var writer = new ICalWriter();
        var ics = writer.Build(doc);

        Assert.Contains("BEGIN:VCALENDAR", ics);
        Assert.Contains("BEGIN:VEVENT", ics);
        Assert.Contains("SUMMARY:Sprint Review", ics);
        Assert.Contains("UID:write-test-001", ics);
        Assert.Contains("LOCATION:Conference Room A", ics);
        Assert.Contains("END:VEVENT", ics);
        Assert.Contains("END:VCALENDAR", ics);
    }

    [Fact]
    [DisplayName("写入全天事件（VALUE=DATE）")]
    public void Write_AllDayEvent_UsesDateFormat()
    {
        var doc = new ICalDocument();
        doc.Events.Add(new ICalEvent
        {
            Summary = "Holiday",
            AllDay = true,
            Start = new DateTimeOffset(2024, 12, 25, 0, 0, 0, TimeSpan.Zero),
            End = new DateTimeOffset(2024, 12, 26, 0, 0, 0, TimeSpan.Zero),
        });

        var writer = new ICalWriter();
        var ics = writer.Build(doc);

        Assert.Contains("DTSTART;VALUE=DATE:20241225", ics);
        Assert.Contains("DTEND;VALUE=DATE:20241226", ics);
    }

    [Fact]
    [DisplayName("写入 VTODO")]
    public void Write_Todo_ContainsRequiredFields()
    {
        var doc = new ICalDocument();
        doc.Todos.Add(new ICalTodo
        {
            Summary = "Fix bug",
            Status = "IN-PROCESS",
            Priority = 1,
            PercentComplete = 50,
            Due = new DateTimeOffset(2024, 6, 30, 17, 0, 0, TimeSpan.Zero),
        });

        var writer = new ICalWriter();
        var ics = writer.Build(doc);

        Assert.Contains("BEGIN:VTODO", ics);
        Assert.Contains("SUMMARY:Fix bug", ics);
        Assert.Contains("PRIORITY:1", ics);
        Assert.Contains("PERCENT-COMPLETE:50", ics);
        Assert.Contains("END:VTODO", ics);
    }

    [Fact]
    [DisplayName("往返测试：写入后读取还原日历")]
    public void RoundTrip_WriteAndRead()
    {
        var doc = new ICalDocument { CalendarName = "Test Cal" };
        doc.Events.Add(new ICalEvent
        {
            Uid = "rt-001",
            Summary = "Round Trip Event",
            Description = "Line1\nLine2",
            Start = new DateTimeOffset(2024, 3, 15, 9, 0, 0, TimeSpan.Zero),
            End = new DateTimeOffset(2024, 3, 15, 10, 0, 0, TimeSpan.Zero),
        });

        var writer = new ICalWriter();
        var ics = writer.Build(doc);

        var reader = new ICalReader();
        var parsed = reader.Parse(ics);

        Assert.Single(parsed.Events);
        Assert.Equal("Round Trip Event", parsed.Events[0].Summary);
        Assert.Equal("rt-001", parsed.Events[0].Uid);
        Assert.Equal(2024, parsed.Events[0].Start!.Value.Year);
    }

    #endregion

    #region 集成测试

    [Fact]
    [DisplayName("集成：写入 ics 文件并读取")]
    public void Integration_WriteFile_ThenReadFile()
    {
        var dir = Path.Combine("Bin", "UnitTest", "Artifacts");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "test_output.ics");

        var doc = new ICalDocument
        {
            CalendarName = "NewLife 测试日历",
            TimeZone = "Asia/Shanghai",
        };
        doc.Events.Add(new ICalEvent
        {
            Uid = "integ-001@newlife.org",
            Summary = "项目启动会议",
            Description = "讨论项目计划\\n确定里程碑",
            Location = "会议室 A",
            Start = new DateTimeOffset(2024, 7, 1, 9, 0, 0, TimeSpan.Zero),
            End = new DateTimeOffset(2024, 7, 1, 11, 0, 0, TimeSpan.Zero),
            Status = "CONFIRMED",
        });
        doc.Todos.Add(new ICalTodo
        {
            Summary = "准备演示材料",
            Status = "NEEDS-ACTION",
            Due = new DateTimeOffset(2024, 6, 28, 18, 0, 0, TimeSpan.Zero),
            Priority = 1,
        });

        var writer = new ICalWriter();
        writer.Write(doc, path);

        Assert.True(File.Exists(path));

        var reader = new ICalReader();
        var parsed = reader.Read(path);

        Assert.Single(parsed.Events);
        Assert.Single(parsed.Todos);
        Assert.Equal("项目启动会议", parsed.Events[0].Summary);
        Assert.Equal("会议室 A", parsed.Events[0].Location);
        Assert.Equal("准备演示材料", parsed.Todos[0].Summary);
    }

    #endregion
}
