using System.ComponentModel;
using System.Text;
using System.Text.Json;
using System.Reflection;
using NewLife.Office.Word;
using Xunit;
using WDocument = NewLife.Office.Word.Document;

namespace XUnitTest.Common;

/// <summary>MCP 服务协议测试：直接驱动 JSON-RPC 分发，验证工具列表与调用契约</summary>
/// <remarks>
/// stdio 循环无法在测试进程内重定向，因此把 <see cref="McpServer.Handle"/> 开放为 internal，
/// 由本类直接喂入请求行并校验响应。覆盖 initialize / tools/list / tools/call / 未知方法 /
/// 非法 JSON / 通知无响应 / 工具内部失败（不得让异常穿出）。
/// </remarks>
public class McpServerTests
{
    /// <summary>构造一行 JSON-RPC 请求</summary>
    /// <param name="id">请求编号</param>
    /// <param name="method">方法名</param>
    /// <param name="params">参数 JSON（可为 null）</param>
    /// <returns>请求文本</returns>
    private static String Request(Int32 id, String method, String? @params = null)
        => @params == null
            ? $"{{\"jsonrpc\":\"2.0\",\"id\":{id},\"method\":\"{method}\"}}"
            : $"{{\"jsonrpc\":\"2.0\",\"id\":{id},\"method\":\"{method}\",\"params\":{@params}}}";

    /// <summary>MCP 服务类型（internal，通过反射驱动以避免为测试扩大可见性）</summary>
    private static readonly Type McpType = Type.GetType("NewLife.Office.Cli.McpServer, NewLife.Office.Cli")!;

    /// <summary>协议分发方法（private static Handle）</summary>
    private static readonly MethodInfo HandleMethod = McpType.GetMethod("Handle", BindingFlags.NonPublic | BindingFlags.Static)!;

    /// <summary>驱动一次协议分发</summary>
    /// <param name="line">请求文本</param>
    /// <returns>响应文本；通知返回 null</returns>
    private static String? Handle(String line)
    {
        try
        {
            return (String?)HandleMethod.Invoke(null, [line]);
        }
        catch (TargetInvocationException ex) when (ex.InnerException != null)
        {
            throw ex.InnerException;
        }
    }

    /// <summary>调用处理并解析响应</summary>
    /// <param name="line">请求文本</param>
    /// <returns>响应的 JSON 文档（调用方负责释放）</returns>
    private static JsonDocument Call(String line)
    {
        var response = Handle(line);
        Assert.NotNull(response);
        return JsonDocument.Parse(response!);
    }

    private static String CreateDocx(String dir)
    {
        var path = Path.Combine(dir, "in.docx");
        var doc = new WDocument();
        doc.AddParagraph("第一段");
        doc.AddParagraph("第二段");
        new WordWriter().Save(path, doc);
        return path;
    }

    [Fact, DisplayName("MCP：initialize 返回协议版本与服务信息")]
    public void Initialize_ReturnsServerInfo()
    {
        using var doc = Call(Request(1, "initialize", "{\"protocolVersion\":\"2024-11-05\",\"capabilities\":{}}"));
        var root = doc.RootElement;
        Assert.Equal("2.0", root.GetProperty("jsonrpc").GetString());
        Assert.Equal(1, root.GetProperty("id").GetInt32());
        var result = root.GetProperty("result");
        Assert.False(String.IsNullOrEmpty(result.GetProperty("protocolVersion").GetString()));
        Assert.Equal("newlife-office", result.GetProperty("serverInfo").GetProperty("name").GetString());
    }

    [Fact, DisplayName("MCP：tools/list 返回四格式与统一工具且带 inputSchema")]
    public void ToolsList_ReturnsTools()
    {
        using var doc = Call(Request(2, "tools/list"));
        var tools = doc.RootElement.GetProperty("result").GetProperty("tools");
        Assert.True(tools.GetArrayLength() >= 14);

        var names = tools.EnumerateArray().Select(e => e.GetProperty("name").GetString()).ToList();
        Assert.Contains("docx_read_outline", names);
        Assert.Contains("docx_apply_patch", names);
        Assert.Contains("office_read_outline", names);
        Assert.Contains("office_read_text", names);
        Assert.Contains("office_apply_patch", names);
        Assert.Contains("xlsx_read_outline", names);
        Assert.Contains("xlsx_read_range", names);
        Assert.Contains("xlsx_apply_patch", names);
        Assert.Contains("pptx_read_outline", names);
        Assert.Contains("pptx_apply_patch", names);
        Assert.Contains("pdf_read_outline", names);
        Assert.Contains("pdf_search", names);
        Assert.Contains("pdf_apply_patch", names);

        foreach (var tool in tools.EnumerateArray())
        {
            Assert.Equal("object", tool.GetProperty("inputSchema").GetProperty("type").GetString());
            Assert.True(tool.TryGetProperty("description", out _));
        }
    }

    [Fact, DisplayName("MCP：通知（无 id）不产生响应")]
    public void Notification_NoResponse()
    {
        var response = Handle("{\"jsonrpc\":\"2.0\",\"method\":\"notifications/initialized\"}");
        Assert.Null(response);
    }

    [Fact, DisplayName("MCP：未知方法返回 -32601")]
    public void UnknownMethod_Returns32601()
    {
        using var doc = Call(Request(3, "no/such/method"));
        Assert.Equal(-32601, doc.RootElement.GetProperty("error").GetProperty("code").GetInt32());
    }

    [Fact, DisplayName("MCP：非法 JSON 抛出可被外层捕获的异常")]
    public void InvalidJson_Throws()
    {
        // Run() 会捕获并回 -32603，这里验证 Handle 本身明确失败而非静默返回 null
        Assert.ThrowsAny<JsonException>(() => Handle("{ not json "));
    }

    [Fact, DisplayName("MCP：未知工具名返回 -32602")]
    public void UnknownTool_Returns32602()
    {
        using var doc = Call(Request(4, "tools/call", "{\"name\":\"no_such_tool\",\"arguments\":{}}"));
        Assert.Equal(-32602, doc.RootElement.GetProperty("error").GetProperty("code").GetInt32());
    }

    [Fact, DisplayName("MCP：docx_read_outline 返回结构文本")]
    public void ReadOutlineTool_ReturnsText()
    {
        var dir = Path.Combine(Path.GetTempPath(), "mcp_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = CreateDocx(dir);
            var args = JsonSerializer.Serialize(new { path });
            using var doc = Call(Request(5, "tools/call", $"{{\"name\":\"docx_read_outline\",\"arguments\":{args}}}"));
            var result = doc.RootElement.GetProperty("result");
            Assert.False(result.GetProperty("isError").GetBoolean());
            var text = result.GetProperty("content")[0].GetProperty("text").GetString();
            Assert.Contains("第一段", text);
            Assert.Contains("第二段", text);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("MCP：docx_read_outline 文件不存在时返回 isError 而非崩溃")]
    public void ReadOutlineTool_MissingFile_ReturnsError()
    {
        var ghost = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        var args = JsonSerializer.Serialize(new { path = ghost });
        using var doc = Call(Request(6, "tools/call", $"{{\"name\":\"docx_read_outline\",\"arguments\":{args}}}"));
        var result = doc.RootElement.GetProperty("result");
        Assert.True(result.GetProperty("isError").GetBoolean());
        Assert.False(String.IsNullOrWhiteSpace(result.GetProperty("content")[0].GetProperty("text").GetString()));
    }

    [Fact, DisplayName("MCP：docx_apply_patch 缺少参数时返回 isError")]
    public void ApplyPatchTool_MissingArguments_ReturnsError()
    {
        using var doc = Call(Request(7, "tools/call", "{\"name\":\"docx_apply_patch\",\"arguments\":{\"path\":\"a.docx\"}}"));
        var result = doc.RootElement.GetProperty("result");
        Assert.True(result.GetProperty("isError").GetBoolean());
    }

    [Fact, DisplayName("MCP：docx_apply_patch 成功应用后输出文件存在")]
    public void ApplyPatchTool_Success()
    {
        var dir = Path.Combine(Path.GetTempPath(), "mcp_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = CreateDocx(dir);
            var output = Path.Combine(dir, "out.docx");
            // 锚点为 0 基索引：p#0 即第一个段落
            var request = "{\"jsonrpc\":\"2.0\",\"id\":8,\"method\":\"tools/call\",\"params\":{\"name\":\"docx_apply_patch\",\"arguments\":{"
                + $"\"path\":{JsonSerializer.Serialize(path)},\"output\":{JsonSerializer.Serialize(output)},"
                + "\"patch\":{\"ops\":[{\"op\":\"set_text\",\"anchor\":\"p#0\",\"text\":\"改后段落\"}]}}}}";
            using var doc = Call(request);
            var result = doc.RootElement.GetProperty("result");
            Assert.False(result.GetProperty("isError").GetBoolean());
            Assert.True(File.Exists(output));

            using var reader = new WordReader(output);
            var outline = reader.ReadOutline().ToJson();
            Assert.Contains("改后段落", outline);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("MCP：docx_apply_patch 失败时不产生输出文件")]
    public void ApplyPatchTool_Failure_NoOutput()
    {
        var dir = Path.Combine(Path.GetTempPath(), "mcp_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = CreateDocx(dir);
            var output = Path.Combine(dir, "out.docx");
            var request = "{\"jsonrpc\":\"2.0\",\"id\":9,\"method\":\"tools/call\",\"params\":{\"name\":\"docx_apply_patch\",\"arguments\":{"
                + $"\"path\":{JsonSerializer.Serialize(path)},\"output\":{JsonSerializer.Serialize(output)},"
                + "\"patch\":{\"ops\":[{\"op\":\"replace_text\",\"anchor\":\"不存在的锚点\",\"text\":\"x\"}]}}}}";
            using var doc = Call(request);
            Assert.True(doc.RootElement.GetProperty("result").GetProperty("isError").GetBoolean());
            Assert.False(File.Exists(output), "补丁失败时不应留下输出文件");
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("MCP：ping 返回空结果")]
    public void Ping_ReturnsEmptyResult()
    {
        using var doc = Call(Request(10, "ping"));
        Assert.Equal(JsonValueKind.Object, doc.RootElement.GetProperty("result").ValueKind);
        Assert.Equal(0, doc.RootElement.GetProperty("result").EnumerateObject().Count());
    }
}
