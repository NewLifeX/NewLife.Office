using System.Text;

namespace NewLife.Office;

/// <summary>iCal 日历文件写入器（RFC 5545）</summary>
/// <remarks>
/// 生成符合 RFC 5545 规范的 .ics 文件。
/// 行长度超过 75 字节时自动折行（CRLF + 空格）。
/// </remarks>
public class ICalWriter
{
    #region 写入方法

    /// <summary>将日历文档写入文件</summary>
    /// <param name="document">日历文档</param>
    /// <param name="path">输出文件路径（.ics）</param>
    public void Write(ICalDocument document, String path)
    {
        var content = Build(document);
        File.WriteAllText(path, content, new UTF8Encoding(false));
    }

    /// <summary>将日历文档写入流</summary>
    /// <param name="document">日历文档</param>
    /// <param name="stream">可写输出流</param>
    public void Write(ICalDocument document, Stream stream)
    {
        var content = Build(document);
        var bytes = new UTF8Encoding(false).GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>将日历文档序列化为字符串</summary>
    /// <param name="document">日历文档</param>
    /// <returns>iCal 格式字符串</returns>
    public String Build(ICalDocument document)
    {
        var sb = new StringBuilder();
        AppendLine(sb, "BEGIN:VCALENDAR");
        AppendLine(sb, "VERSION:" + document.Version);
        AppendLine(sb, "PRODID:" + document.ProdId);
        if (document.Method != null) AppendLine(sb, "METHOD:" + document.Method);
        if (document.CalendarName != null) AppendLine(sb, "X-WR-CALNAME:" + EscapeText(document.CalendarName));
        if (document.TimeZone != null) AppendLine(sb, "X-WR-TIMEZONE:" + document.TimeZone);

        foreach (var evt in document.Events)
        {
            WriteEvent(sb, evt);
        }

        foreach (var todo in document.Todos)
        {
            WriteTodo(sb, todo);
        }

        AppendLine(sb, "END:VCALENDAR");
        return sb.ToString();
    }

    #endregion

    #region 私有方法

    private static void WriteEvent(StringBuilder sb, ICalEvent evt)
    {
        AppendLine(sb, "BEGIN:VEVENT");
        AppendLine(sb, "UID:" + (evt.Uid ?? Guid.NewGuid().ToString()));
        if (evt.Summary != null) AppendLine(sb, "SUMMARY:" + EscapeText(evt.Summary));
        if (evt.Description != null) AppendLine(sb, "DESCRIPTION:" + EscapeText(evt.Description));
        if (evt.Location != null) AppendLine(sb, "LOCATION:" + EscapeText(evt.Location));
        if (evt.Organizer != null) AppendLine(sb, "ORGANIZER:" + evt.Organizer);
        foreach (var att in evt.Attendees)
        {
            AppendLine(sb, "ATTENDEE:" + att);
        }
        if (evt.Status != null) AppendLine(sb, "STATUS:" + evt.Status);
        if (evt.Recurrence != null) AppendLine(sb, "RRULE:" + evt.Recurrence);
        if (evt.Duration != null) AppendLine(sb, "DURATION:" + evt.Duration);
        if (evt.Start.HasValue)
        {
            if (evt.AllDay)
                AppendLine(sb, "DTSTART;VALUE=DATE:" + evt.Start.Value.ToString("yyyyMMdd"));
            else
                AppendLine(sb, "DTSTART:" + FormatDateTime(evt.Start.Value));
        }
        if (evt.End.HasValue)
        {
            if (evt.AllDay)
                AppendLine(sb, "DTEND;VALUE=DATE:" + evt.End.Value.ToString("yyyyMMdd"));
            else
                AppendLine(sb, "DTEND:" + FormatDateTime(evt.End.Value));
        }
        if (evt.Created.HasValue) AppendLine(sb, "CREATED:" + FormatDateTime(evt.Created.Value));
        if (evt.LastModified.HasValue) AppendLine(sb, "LAST-MODIFIED:" + FormatDateTime(evt.LastModified.Value));
        foreach (var kv in evt.ExtraProps)
        {
            AppendLine(sb, kv.Key + ":" + kv.Value);
        }
        AppendLine(sb, "END:VEVENT");
    }

    private static void WriteTodo(StringBuilder sb, ICalTodo todo)
    {
        AppendLine(sb, "BEGIN:VTODO");
        AppendLine(sb, "UID:" + (todo.Uid ?? Guid.NewGuid().ToString()));
        if (todo.Summary != null) AppendLine(sb, "SUMMARY:" + EscapeText(todo.Summary));
        if (todo.Description != null) AppendLine(sb, "DESCRIPTION:" + EscapeText(todo.Description));
        if (todo.Status != null) AppendLine(sb, "STATUS:" + todo.Status);
        if (todo.Priority.HasValue) AppendLine(sb, "PRIORITY:" + todo.Priority.Value);
        if (todo.PercentComplete.HasValue) AppendLine(sb, "PERCENT-COMPLETE:" + todo.PercentComplete.Value);
        if (todo.Start.HasValue) AppendLine(sb, "DTSTART:" + FormatDateTime(todo.Start.Value));
        if (todo.Due.HasValue) AppendLine(sb, "DUE:" + FormatDateTime(todo.Due.Value));
        if (todo.Completed.HasValue) AppendLine(sb, "COMPLETED:" + FormatDateTime(todo.Completed.Value));
        AppendLine(sb, "END:VTODO");
    }

    /// <summary>添加折行支持的属性行（RFC 5545 要求内容行最长 75 字节）</summary>
    private static void AppendLine(StringBuilder sb, String line)
    {
        var bytes = Encoding.UTF8.GetBytes(line);
        if (bytes.Length <= 75)
        {
            sb.Append(line).Append("\r\n");
            return;
        }

        // 折行：每 75 字节截断，续行以空格开头
        var pos = 0;
        var first = true;
        while (pos < line.Length)
        {
            if (!first) sb.Append(" ");
            // 贪心取不超过 75 字节的字符
            var take = TakeChars(line, pos, first ? 75 : 74);
            sb.Append(line, pos, take);
            sb.Append("\r\n");
            pos += take;
            first = false;
        }
    }

    /// <summary>从 pos 开始取最多 maxBytes 字节对应的字符数（避免截断 UTF-8 多字节字符）</summary>
    private static Int32 TakeChars(String text, Int32 pos, Int32 maxBytes)
    {
        var bytes = 0;
        var count = 0;
        while (pos + count < text.Length)
        {
            var c = text[pos + count];
            var cb = Encoding.UTF8.GetByteCount(new Char[] { c });
            if (bytes + cb > maxBytes) break;
            bytes += cb;
            count++;
        }
        return count > 0 ? count : 1;  // 至少取 1 个字符防止死循环
    }

    private static String FormatDateTime(DateTimeOffset dt)
    {
        return dt.ToUniversalTime().ToString("yyyyMMddTHHmmssZ");
    }

    private static String EscapeText(String value)
    {
        return value.Replace("\\", "\\\\").Replace(";", "\\;")
                    .Replace(",", "\\,").Replace("\n", "\\n").Replace("\r", "");
    }

    #endregion
}
