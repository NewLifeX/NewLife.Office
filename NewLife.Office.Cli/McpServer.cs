using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using NewLife.Office.Word;

namespace NewLife.Office.Cli;

/// <summary>极简 MCP（Model Context Protocol）stdio 服务端：把 docx 读写能力暴露给 Agent</summary>
/// <remarks>
/// 走 MCP 的 stdio 传输：每行一条 JSON-RPC 2.0 消息（请求/响应/通知），stdin 读、stdout 写。
/// 只实现最小可用集合：<c>initialize</c> / <c>tools/list</c> / <c>tools/call</c> / <c>ping</c>。
/// <para>工具：<c>docx_read_outline</c>（读结构快照）、<c>docx_apply_patch</c>（按补丁改文档）。</para>
/// <para>零外部依赖：只用 BCL 的 System.Text.Json 手写协议层，不引入 MCP SDK。</para>
/// </remarks>
internal static class McpServer
{
    private const String ProtocolVersion = "2024-11-05";

    /// <summary>启动 stdio 循环（由 CLI 的 mcp 命令调用）</summary>
    /// <returns>退出码</returns>
    public static Int32 Run()
    {
        using var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };

        String? line;
        while ((line = stdin.ReadLine()) != null)
        {
            line = line.Trim();
            if (line.Length == 0) continue;

            String? response;
            try
            {
                response = Handle(line);
            }
            catch (Exception ex)
            {
                response = Error(null, -32603, ex.Message);
            }
            if (response != null) stdout.WriteLine(response);
        }
        return 0;
    }

    #region 协议
    /// <summary>处理一行 JSON-RPC 请求，返回响应文本（通知返回 null）</summary>
    /// <param name="line">一行 JSON-RPC 请求文本</param>
    /// <returns>响应 JSON 文本；通知类消息返回 null</returns>
    private static String? Handle(String line)
    {
        using var doc = JsonDocument.Parse(line);
        var root = doc.RootElement;

        var method = root.TryGetProperty("method", out var methodEl) ? methodEl.GetString() : null;
        JsonElement? id = root.TryGetProperty("id", out var idEl) ? idEl : null;

        // 通知（无 id）不需要响应
        if (method == null || id == null) return null;

        switch (method)
        {
            case "initialize":
                return Result(id.Value, new JsonObject
                {
                    ["protocolVersion"] = ProtocolVersion,
                    ["capabilities"] = new JsonObject
                    {
                        ["tools"] = new JsonObject(),
                    },
                    ["serverInfo"] = new JsonObject
                    {
                        ["name"] = "newlife-office-docx",
                        ["version"] = "1.0.0",
                    },
                });

            case "tools/list":
                return Result(id.Value, new JsonObject { ["tools"] = BuildTools() });

            case "tools/call":
                return CallTool(id.Value, root);

            case "ping":
                return Result(id.Value, new JsonObject());

            default:
                return Error(id.Value, -32601, $"Method not found: {method}");
        }
    }

    private static JsonArray BuildTools()
    {
        var read = new JsonObject
        {
            ["name"] = "docx_read_outline",
            ["description"] = "读取 docx 结构快照（块级视图，带 p#3/tbl#5 锚点、标题层级、表格矩阵），返回 JSON 文本",
            ["inputSchema"] = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["path"] = new JsonObject { ["type"] = "string", ["description"] = "docx 文件路径" },
                },
                ["required"] = new JsonArray("path"),
            },
        };

        var apply = new JsonObject
        {
            ["name"] = "docx_apply_patch",
            ["description"] = "按锚点批量修改 docx：replace_text/set_text/insert_paragraph/delete/set_style/set_cell_text/insert_table/add_comment。任一条失败则不写文件。注意：锚点指 read 时刻的元素下标，insert/delete 会让后续锚点偏移，请把结构性操作排在补丁最后",
            ["inputSchema"] = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["path"] = new JsonObject { ["type"] = "string", ["description"] = "输入 docx 路径" },
                    ["output"] = new JsonObject { ["type"] = "string", ["description"] = "输出 docx 路径" },
                    ["patch"] = new JsonObject { ["type"] = "object", ["description"] = "补丁对象，形如 {\"ops\":[...]}" },
                },
                ["required"] = new JsonArray("path", "output", "patch"),
            },
        };

        return [read, apply];
    }

    private static String CallTool(JsonElement id, JsonElement root)
    {
        if (!root.TryGetProperty("params", out var ps) || !ps.TryGetProperty("name", out var nameEl))
            return Error(id, -32602, "Missing params.name");

        var name = nameEl.GetString();
        var args = ps.TryGetProperty("arguments", out var a) ? a : default;

        try
        {
            return name switch
            {
                "docx_read_outline" => ReadOutline(id, args),
                "docx_apply_patch" => ApplyPatch(id, args),
                _ => Error(id, -32602, $"Unknown tool: {name}"),
            };
        }
        catch (Exception ex)
        {
            return ToolText(id, $"失败：{ex.Message}", isError: true);
        }
    }
    #endregion

    #region 工具实现
    private static String ReadOutline(JsonElement id, JsonElement args)
    {
        var path = Program.ResolvePath(GetString(args, "path") ?? throw new ArgumentException("缺少 path"));
        using var reader = new WordReader(path);
        return ToolText(id, reader.ReadOutline().ToJson());
    }

    private static String ApplyPatch(JsonElement id, JsonElement args)
    {
        var path = Program.ResolvePath(GetString(args, "path") ?? throw new ArgumentException("缺少 path"));
        var output = Program.ResolvePath(GetString(args, "output") ?? throw new ArgumentException("缺少 output"));
        if (!args.TryGetProperty("patch", out var patchEl)) throw new ArgumentException("缺少 patch");

        var patch = DocxPatch.ParseJson(patchEl.GetRawText());

        using var reader = new WordReader(path);
        var doc = reader.ReadDocument();
        doc.BeginEdit();

        var result = DocxPatch.Apply(doc, patch);
        if (!result.Success)
            return ToolText(id, "失败，未写出文件：\n" + String.Join("\n", result.Failed), isError: true);

        using (var writer = new WordWriter())
            writer.Save(output, doc);

        var text = $"已应用 {result.Applied.Count} 条操作 → {output}";
        // 告警不影响成功状态，但要显式告知，避免“内容被静默丢弃”
        if (result.Warnings.Count > 0)
            text += "\n告警：\n" + String.Join("\n", result.Warnings);

        return ToolText(id, text);
    }

    private static String? GetString(JsonElement args, String name)
        => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.String
            ? el.GetString()
            : null;
    #endregion

    #region 响应构造
    private static String ToolText(JsonElement id, String text, Boolean isError = false)
    {
        var content = new JsonArray
        {
            new JsonObject
            {
                ["type"] = "text",
                ["text"] = text,
            },
        };

        return Result(id, new JsonObject
        {
            ["content"] = content,
            ["isError"] = isError,
        });
    }

    private static String Result(JsonElement id, JsonNode result)
        => new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = JsonNode.Parse(id.GetRawText()),
            ["result"] = result,
        }.ToJsonString();

    private static String Error(JsonElement? id, Int32 code, String message)
        => new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id.HasValue ? JsonNode.Parse(id.Value.GetRawText()) : null,
            ["error"] = new JsonObject
            {
                ["code"] = code,
                ["message"] = message,
            },
        }.ToJsonString();
    #endregion
}
