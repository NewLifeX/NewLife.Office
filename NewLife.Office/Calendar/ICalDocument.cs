namespace NewLife.Office;

/// <summary>iCal 日历文档（RFC 5545）</summary>
/// <remarks>包含一组日历事件（VEVENT）和待办（VTODO）</remarks>
public class ICalDocument
{
    #region 属性

    /// <summary>日历版本，默认 2.0</summary>
    public String Version { get; set; } = "2.0";

    /// <summary>产品标识（PRODID）</summary>
    public String ProdId { get; set; } = "-//NewLife//NewLife.Office//ZH";

    /// <summary>日历名称（X-WR-CALNAME）</summary>
    public String? CalendarName { get; set; }

    /// <summary>时区标识（X-WR-TIMEZONE）</summary>
    public String? TimeZone { get; set; }

    /// <summary>日历方法（METHOD），如 PUBLISH/REQUEST/REPLY</summary>
    public String? Method { get; set; }

    /// <summary>事件列表</summary>
    public List<ICalEvent> Events { get; } = [];

    /// <summary>待办列表</summary>
    public List<ICalTodo> Todos { get; } = [];

    #endregion
}

/// <summary>iCal 日历事件（VEVENT，RFC 5545 第 3.6.1 节）</summary>
public class ICalEvent
{
    #region 属性

    /// <summary>唯一标识（UID）</summary>
    public String? Uid { get; set; }

    /// <summary>摘要/标题（SUMMARY）</summary>
    public String? Summary { get; set; }

    /// <summary>描述（DESCRIPTION）</summary>
    public String? Description { get; set; }

    /// <summary>地点（LOCATION）</summary>
    public String? Location { get; set; }

    /// <summary>开始时间（DTSTART）</summary>
    public DateTimeOffset? Start { get; set; }

    /// <summary>结束时间（DTEND）</summary>
    public DateTimeOffset? End { get; set; }

    /// <summary>持续时长（DURATION），格式如 P1DT2H3M</summary>
    public String? Duration { get; set; }

    /// <summary>全天事件标志</summary>
    public Boolean AllDay { get; set; }

    /// <summary>创建时间（CREATED）</summary>
    public DateTimeOffset? Created { get; set; }

    /// <summary>最后修改时间（LAST-MODIFIED）</summary>
    public DateTimeOffset? LastModified { get; set; }

    /// <summary>重复规则（RRULE），如 FREQ=WEEKLY;BYDAY=MO,WE,FR</summary>
    public String? Recurrence { get; set; }

    /// <summary>状态（STATUS），如 CONFIRMED/TENTATIVE/CANCELLED</summary>
    public String? Status { get; set; }

    /// <summary>组织者（ORGANIZER）</summary>
    public String? Organizer { get; set; }

    /// <summary>参与者列表（ATTENDEE）</summary>
    public List<String> Attendees { get; } = [];

    /// <summary>扩展属性</summary>
    public Dictionary<String, String> ExtraProps { get; } = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);

    #endregion
}

/// <summary>iCal 待办（VTODO，RFC 5545 第 3.6.2 节）</summary>
public class ICalTodo
{
    #region 属性

    /// <summary>唯一标识（UID）</summary>
    public String? Uid { get; set; }

    /// <summary>摘要/标题（SUMMARY）</summary>
    public String? Summary { get; set; }

    /// <summary>描述（DESCRIPTION）</summary>
    public String? Description { get; set; }

    /// <summary>开始时间（DTSTART）</summary>
    public DateTimeOffset? Start { get; set; }

    /// <summary>截止时间（DUE）</summary>
    public DateTimeOffset? Due { get; set; }

    /// <summary>完成时间（COMPLETED）</summary>
    public DateTimeOffset? Completed { get; set; }

    /// <summary>完成百分比（PERCENT-COMPLETE，0-100）</summary>
    public Int32? PercentComplete { get; set; }

    /// <summary>优先级（PRIORITY，0-9，1 最高）</summary>
    public Int32? Priority { get; set; }

    /// <summary>状态（STATUS），如 NEEDS-ACTION/IN-PROCESS/COMPLETED/CANCELLED</summary>
    public String? Status { get; set; }

    #endregion
}
