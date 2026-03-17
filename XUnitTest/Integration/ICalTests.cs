using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>iCal 日历格式集成测试</summary>
public class ICalTests : IntegrationTestBase
{
    [Fact, DisplayName("iCal_复杂写入再读取往返")]
    public void ICal_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.ics");

        var doc = new ICalDocument
        {
            CalendarName = "集成测试日历",
            TimeZone = "Asia/Shanghai",
            Method = "PUBLISH",
        };

        doc.Events.Add(new ICalEvent
        {
            Uid = "event-001@newlife.office",
            Summary = "项目启动会",
            Description = "讨论新项目的技术方案和人员分工。",
            Location = "会议室A301",
            Start = new DateTimeOffset(2024, 7, 1, 9, 0, 0, TimeSpan.FromHours(8)),
            End = new DateTimeOffset(2024, 7, 1, 11, 0, 0, TimeSpan.FromHours(8)),
            Status = "CONFIRMED",
            Organizer = "manager@example.com",
        });

        doc.Events.Add(new ICalEvent
        {
            Uid = "event-002@newlife.office",
            Summary = "全天培训",
            Description = "新员工培训日。",
            Start = new DateTimeOffset(2024, 7, 5, 0, 0, 0, TimeSpan.FromHours(8)),
            End = new DateTimeOffset(2024, 7, 5, 23, 59, 59, TimeSpan.FromHours(8)),
            AllDay = true,
        });

        doc.Todos.Add(new ICalTodo
        {
            Uid = "todo-001@newlife.office",
            Summary = "完成集成测试",
            Description = "编写并运行所有格式的集成测试。",
            Due = new DateTimeOffset(2024, 7, 10, 18, 0, 0, TimeSpan.FromHours(8)),
            Priority = 1,
            Status = "IN-PROCESS",
            PercentComplete = 50,
        });

        new ICalWriter().Write(doc, path);

        Assert.True(File.Exists(path));

        // 读取验证
        var reader = new ICalReader();
        var readDoc = reader.Read(path);
        Assert.Equal("2.0", readDoc.Version);
        Assert.Equal("集成测试日历", readDoc.CalendarName);

        Assert.Equal(2, readDoc.Events.Count);
        Assert.Equal("项目启动会", readDoc.Events[0].Summary);
        Assert.Equal("会议室A301", readDoc.Events[0].Location);
        Assert.Equal("全天培训", readDoc.Events[1].Summary);

        Assert.True(readDoc.Todos.Count >= 1);
        Assert.Equal("完成集成测试", readDoc.Todos[0].Summary);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<ICalReader>(factoryReader);
    }
}
