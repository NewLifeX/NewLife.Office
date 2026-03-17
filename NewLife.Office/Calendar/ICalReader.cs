using System.Globalization;
using System.Text;

namespace NewLife.Office;

/// <summary>iCal 日历文件读取器（RFC 5545）</summary>
/// <remarks>
/// 支持解析 .ics 文件中的 VCALENDAR、VEVENT、VTODO 组件。
/// 支持 DTSTART/DTEND 的日期时间及全天日期格式（VALUE=DATE）。
/// </remarks>
public class ICalReader
{
    #region 读取方法

    /// <summary>从文件读取 iCal 文档</summary>
    /// <param name="path">iCal 文件路径（.ics）</param>
    /// <returns>解析后的日历文档</returns>
    public ICalDocument Read(String path)
    {
        var text = File.ReadAllText(path, Encoding.UTF8);
        return Parse(text);
    }

    /// <summary>从流读取 iCal 文档</summary>
    /// <param name="stream">可读流</param>
    /// <returns>解析后的日历文档</returns>
    public ICalDocument Read(Stream stream)
    {
        using var sr = new StreamReader(stream, Encoding.UTF8);
        return Parse(sr.ReadToEnd());
    }

    /// <summary>从字符串解析 iCal 文档</summary>
    /// <param name="text">iCal 文本内容</param>
    /// <returns>解析后的日历文档</returns>
    public ICalDocument Parse(String text)
    {
        var doc = new ICalDocument();
        var lines = UnfoldLines(text);
        var props = ParseProperties(lines);

        ICalEvent? currentEvent = null;
        ICalTodo? currentTodo = null;
        var inCalendar = false;

        foreach (var (name, param, value) in props)
        {
            var nameLow = name.ToLowerInvariant();
            switch (nameLow)
            {
                case "begin":
                    switch (value.ToUpperInvariant())
                    {
                        case "VCALENDAR": inCalendar = true; break;
                        case "VEVENT": currentEvent = new ICalEvent(); break;
                        case "VTODO": currentTodo = new ICalTodo(); break;
                    }
                    break;
                case "end":
                    switch (value.ToUpperInvariant())
                    {
                        case "VCALENDAR": inCalendar = false; break;
                        case "VEVENT":
                            if (currentEvent != null) { doc.Events.Add(currentEvent); currentEvent = null; }
                            break;
                        case "VTODO":
                            if (currentTodo != null) { doc.Todos.Add(currentTodo); currentTodo = null; }
                            break;
                    }
                    break;
                default:
                    if (currentEvent != null)
                        ApplyEventProp(currentEvent, name, param, value);
                    else if (currentTodo != null)
                        ApplyTodoProp(currentTodo, name, param, value);
                    else if (inCalendar)
                        ApplyCalendarProp(doc, nameLow, value);
                    break;
            }
        }

        return doc;
    }

    #endregion

    #region 私有方法

    private static List<String> UnfoldLines(String text)
    {
        // RFC 5545: 行以 CRLF 或 LF 结束，续行以空格/TAB 开头
        var lines = new List<String>();
        var sb = new StringBuilder();
        foreach (var line in text.Split('\n'))
        {
            var l = line.TrimEnd('\r');
            if (l.Length > 0 && (l[0] == ' ' || l[0] == '\t'))
                sb.Append(l[1..]);
            else
            {
                if (sb.Length > 0) lines.Add(sb.ToString());
                sb.Clear();
                sb.Append(l);
            }
        }
        if (sb.Length > 0) lines.Add(sb.ToString());
        return lines;
    }

    /// <summary>解析属性行，返回 (名称, 参数串, 值) 三元组</summary>
    private static List<(String Name, String Param, String Value)> ParseProperties(List<String> lines)
    {
        var result = new List<(String, String, String)>();
        foreach (var line in lines)
        {
            if (String.IsNullOrWhiteSpace(line)) continue;
            var colonIdx = line.IndexOf(':');
            if (colonIdx <= 0) continue;
            var left = line[..colonIdx];
            var val = line[(colonIdx + 1)..];
            var semiIdx = left.IndexOf(';');
            var name = semiIdx >= 0 ? left[..semiIdx] : left;
            var param = semiIdx >= 0 ? left[(semiIdx + 1)..] : String.Empty;
            result.Add((name.Trim(), param.Trim(), val.Trim()));
        }
        return result;
    }

    private static void ApplyCalendarProp(ICalDocument doc, String name, String value)
    {
        switch (name)
        {
            case "version": doc.Version = value; break;
            case "prodid": doc.ProdId = value; break;
            case "method": doc.Method = value; break;
            case "x-wr-calname": doc.CalendarName = UnescapeText(value); break;
            case "x-wr-timezone": doc.TimeZone = value; break;
        }
    }

    private static void ApplyEventProp(ICalEvent evt, String name, String param, String value)
    {
        switch (name.ToLowerInvariant())
        {
            case "uid": evt.Uid = value; break;
            case "summary": evt.Summary = UnescapeText(value); break;
            case "description": evt.Description = UnescapeText(value); break;
            case "location": evt.Location = UnescapeText(value); break;
            case "rrule": evt.Recurrence = value; break;
            case "status": evt.Status = value; break;
            case "organizer": evt.Organizer = value; break;
            case "attendee": evt.Attendees.Add(value); break;
            case "dtstart":
                evt.AllDay = param.IndexOf("VALUE=DATE", StringComparison.OrdinalIgnoreCase) >= 0
                             && !param.Contains("DATE-TIME");
                evt.Start = ParseDateTime(value, evt.AllDay);
                break;
            case "dtend":
                var allDayEnd = param.IndexOf("VALUE=DATE", StringComparison.OrdinalIgnoreCase) >= 0
                                && !param.Contains("DATE-TIME");
                evt.End = ParseDateTime(value, allDayEnd);
                break;
            case "duration": evt.Duration = value; break;
            case "created": evt.Created = ParseDateTime(value, false); break;
            case "last-modified": evt.LastModified = ParseDateTime(value, false); break;
            default:
                evt.ExtraProps[name] = value;
                break;
        }
    }

    private static void ApplyTodoProp(ICalTodo todo, String name, String param, String value)
    {
        switch (name.ToLowerInvariant())
        {
            case "uid": todo.Uid = value; break;
            case "summary": todo.Summary = UnescapeText(value); break;
            case "description": todo.Description = UnescapeText(value); break;
            case "status": todo.Status = value; break;
            case "dtstart": todo.Start = ParseDateTime(value, false); break;
            case "due": todo.Due = ParseDateTime(value, false); break;
            case "completed": todo.Completed = ParseDateTime(value, false); break;
            case "percent-complete":
                if (Int32.TryParse(value, out var pct)) todo.PercentComplete = pct;
                break;
            case "priority":
                if (Int32.TryParse(value, out var pri)) todo.Priority = pri;
                break;
        }
    }

    /// <summary>解析 iCal 日期/时间格式：20240101T120000Z 或 20240101</summary>
    private static DateTimeOffset? ParseDateTime(String value, Boolean allDay)
    {
        if (String.IsNullOrEmpty(value)) return null;
        if (allDay || value.Length == 8)
        {
            // YYYYMMDD
            if (DateTime.TryParseExact(value, "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var d))
                return new DateTimeOffset(d, TimeSpan.Zero);
        }
        else
        {
            // YYYYMMDDTHHMISS or YYYYMMDDTHHMISSZ
            var fmt = value.EndsWith("Z", StringComparison.OrdinalIgnoreCase) ? "yyyyMMddTHHmmssZ" : "yyyyMMddTHHmmss";
            if (DateTime.TryParseExact(value.TrimEnd('Z', 'z'), "yyyyMMddTHHmmss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal, out var dt))
                return new DateTimeOffset(dt.ToUniversalTime(), TimeSpan.Zero);
        }
        return null;
    }

    private static String UnescapeText(String value)
    {
        return value.Replace("\\n", "\n").Replace("\\N", "\n")
                    .Replace("\\,", ",").Replace("\\;", ";")
                    .Replace("\\\\", "\\");
    }

    #endregion
}
