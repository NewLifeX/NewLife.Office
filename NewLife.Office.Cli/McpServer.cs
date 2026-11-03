using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using NewLife.Office.Word;
using NewLife;

namespace NewLife.Office.Cli;

/// <summary>极简 MCP（Model Context Protocol）stdio 服务端：把 docx/xlsx/pptx/pdf 读写能力暴露给 Agent</summary>
/// <remarks>
/// 走 MCP 的 stdio 传输：每行一条 JSON-RPC 2.0 消息（请求/响应/通知），stdin 读、stdout 写。
/// 只实现最小可用集合：<c>initialize</c> / <c>tools/list</c> / <c>tools/call</c> / <c>ping</c>。
/// <para>工具：四格式的结构快照/补丁，以及统一的 <c>office_read_outline</c> / <c>office_read_text</c> / <c>office_search</c> / <c>office_apply_patch</c>。</para>
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
                        ["name"] = "newlife-office",
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
        var read = Tool("docx_read_outline", "读取 docx 结构快照（块级视图，带 p#3/tbl#5 锚点、标题层级、表格矩阵），返回 JSON 文本",
            new JsonObject { ["path"] = PathProp("docx 文件路径") }, "path");

        var apply = Tool("docx_apply_patch", "按锚点批量修改 docx：replace_text/set_text/insert_paragraph/delete/set_style/set_cell_text/insert_table/add_comment。任一条失败则不写文件。注意：锚点指 read 时刻的元素下标，insert/delete 会让后续锚点偏移，请把结构性操作排在补丁最后",
            new JsonObject
            {
                ["path"] = PathProp("输入 docx 路径"),
                ["output"] = PathProp("输出 docx 路径"),
                ["patch"] = new JsonObject { ["type"] = "object", ["description"] = "补丁对象，形如 {\"ops\":[...]}" },
            }, "path", "output", "patch");

        var officeOutline = Tool("office_read_outline", "读取任意受支持文档（docx/xlsx/pptx/pdf）的结构快照，按扩展名或内容自动识别格式，返回 JSON 文本",
            new JsonObject { ["path"] = PathProp("文档路径") }, "path");

        var officeText = Tool("office_read_text", "分段读取任意文档的纯文本：maxChars 控制本次字符数，cursor 传上次返回的 nextCursor 续读",
            new JsonObject
            {
                ["path"] = PathProp("文档路径"),
                ["maxChars"] = IntProp("本次最多返回字符数（0 不限）"),
                ["cursor"] = IntProp("续读游标（上次的 nextCursor）"),
            }, "path");

        var officeSearch = Tool("office_search", "在工作簿（文本/公式/正则）或 PDF 中搜索，返回命中位置（Sheet!A1 或 页码 + 矩形坐标）",
            new JsonObject
            {
                ["path"] = PathProp("文档路径"),
                ["text"] = StrProp("搜索内容"),
                ["regex"] = new JsonObject { ["type"] = "boolean", ["description"] = "是否按正则表达式" },
                ["maxResults"] = IntProp("最大命中数（默认 100）"),
            }, "path", "text");

        var officeApply = Tool("office_apply_patch", "对任意受支持文档应用 JSON 补丁：docx（段落/表格）、xlsx（A1 区域/样式/行）、pptx（形状/表格/页）、pdf（页/标注/表单/元数据）。任一条失败则不写文件",
            new JsonObject
            {
                ["path"] = PathProp("输入文件路径"),
                ["output"] = PathProp("输出文件路径"),
                ["patch"] = new JsonObject { ["type"] = "object", ["description"] = "补丁对象，形如 {\"ops\":[...]}" },
            }, "path", "output", "patch");

        var xlsxOutline = Tool("xlsx_read_outline", "读取 xlsx/xls 结构快照（工作表、已用区域、表头、前 5 行预览、表格对象、图表/批注统计）",
            new JsonObject { ["path"] = PathProp("工作簿路径") }, "path");

        var xlsxRange = Tool("xlsx_read_range", "按 A1 区域读取工作簿数据：maxRows 控制本次行数，cursor 传上次的 nextCursor 续读",
            new JsonObject
            {
                ["path"] = PathProp("工作簿路径"),
                ["range"] = StrProp("A1 区域（如 Sheet1!A1:D100）"),
                ["maxRows"] = IntProp("本次最多返回行数（0 不限）"),
                ["cursor"] = IntProp("续读游标（上次的 nextCursor）"),
            }, "path", "range");

        var xlsxApply = Tool("xlsx_apply_patch", "按 A1 坐标批量修改工作簿：set_cells/set_formula/append_rows/insert_rows/delete_rows/merge/unmerge/set_style/set_number_format/set_column_width/set_row_height/freeze/autofilter/add_sheet/rename_sheet/delete_sheet/add_table/add_chart。任一条失败则不写文件",
            new JsonObject
            {
                ["path"] = PathProp("输入工作簿路径"),
                ["output"] = PathProp("输出工作簿路径"),
                ["patch"] = new JsonObject { ["type"] = "object", ["description"] = "补丁对象，形如 {\"ops\":[...]}" },
            }, "path", "output", "patch");

        var pptxOutline = Tool("pptx_read_outline", "读取 pptx 结构快照（幻灯片、文本框、形状、图片、表格、图表、备注），带 s#2/tb#0 锚点",
            new JsonObject { ["path"] = PathProp("演示文稿路径") }, "path");

        var pptxApply = Tool("pptx_apply_patch", "按锚点批量修改演示文稿：set_shape_text/set_table_cell/update_chart/replace_image/set_notes/set_title/add_slide/delete_slide/move_slide/clone_slide。任一条失败则不写文件",
            new JsonObject
            {
                ["path"] = PathProp("输入演示文稿路径"),
                ["output"] = PathProp("输出演示文稿路径"),
                ["patch"] = new JsonObject { ["type"] = "object", ["description"] = "补丁对象，形如 {\"ops\":[...]}" },
            }, "path", "output", "patch");

        var pdfOutline = Tool("pdf_read_outline", "读取 PDF 结构快照（页数与尺寸、书签树、表单字段、注释、附件、字体、加密/签名状态、每页首行预览）",
            new JsonObject { ["path"] = PathProp("PDF 路径") }, "path");

        var pdfSearch = Tool("pdf_search", "在 PDF 中搜索文本，返回页码与矩形坐标（可直接用于 add_highlight/add_note）",
            new JsonObject
            {
                ["path"] = PathProp("PDF 路径"),
                ["text"] = StrProp("搜索内容"),
                ["regex"] = new JsonObject { ["type"] = "boolean", ["description"] = "是否按正则表达式" },
                ["maxResults"] = IntProp("最大命中数（默认 100）"),
            }, "path", "text");

        var pdfApply = Tool("pdf_apply_patch", "按对象级补丁修改 PDF：delete_pages/extract_pages/move_page/rotate_page/set_metadata/add_text/add_watermark/add_highlight/add_note/fill_form/merge。任一条失败则不写文件",
            new JsonObject
            {
                ["path"] = PathProp("输入 PDF 路径"),
                ["output"] = PathProp("输出 PDF 路径"),
                ["patch"] = new JsonObject { ["type"] = "object", ["description"] = "补丁对象，形如 {\"ops\":[...]}" },
            }, "path", "output", "patch");

        return [read, apply, officeOutline, officeText, officeSearch, officeApply, xlsxOutline, xlsxRange, xlsxApply, pptxOutline, pptxApply, pdfOutline, pdfSearch, pdfApply];
    }

    private static JsonObject Tool(String name, String description, JsonObject properties, params String[] required)
        => new()
        {
            ["name"] = name,
            ["description"] = description,
            ["inputSchema"] = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = properties,
                ["required"] = new JsonArray(required.Select(e => (JsonNode?)JsonValue.Create(e)).ToArray()),
            },
        };

    private static JsonObject StrProp(String description) => new() { ["type"] = "string", ["description"] = description };

    private static JsonObject IntProp(String description) => new() { ["type"] = "integer", ["description"] = description };

    private static JsonObject PathProp(String description) => StrProp(description);

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
                "docx_read_outline" or "xlsx_read_outline" or "pptx_read_outline" or "pdf_read_outline" or "office_read_outline"
                    => ReadOutline(id, args),
                "xlsx_read_range" => ReadRange(id, args),
                "office_read_text" => ReadText(id, args),
                "office_search" or "pdf_search" or "xlsx_search" => Search(id, args),
                "docx_apply_patch" or "xlsx_apply_patch" or "pptx_apply_patch" or "pdf_apply_patch" or "office_apply_patch"
                    => ApplyPatch(id, args),
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
        var path = Program.ResolvePath(GetString(args, "path") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "path")));
        var result = OfficeDocumentTool.ReadOutline(path);
        return result.Success ? ToolText(id, result.Content!) : ToolError(id, result);
    }

    private static String ReadRange(JsonElement id, JsonElement args)
    {
        var path = Program.ResolvePath(GetString(args, "path") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "path")));
        var range = GetString(args, "range") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "range"));
        var maxRows = GetInt32(args, "maxRows") ?? 0;
        var cursor = GetInt32(args, "cursor") ?? 0;

        var result = OfficeDocumentTool.ReadRange(path, range, maxRows, cursor);
        return result.Success ? ToolText(id, result.Content!) : ToolError(id, result);
    }

    private static String ReadText(JsonElement id, JsonElement args)
    {
        var path = Program.ResolvePath(GetString(args, "path") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "path")));
        var maxChars = GetInt32(args, "maxChars") ?? 0;
        var cursor = GetInt32(args, "cursor") ?? 0;

        var result = OfficeDocumentTool.ReadText(path, maxChars, cursor);
        if (!result.Success) return ToolError(id, result);

        var text = result.Content ?? String.Empty;
        if (result.Truncated) text += $"\n\n--- 已截断，续读请传 cursor={result.NextCursor} ---";
        return ToolText(id, text);
    }

    private static String Search(JsonElement id, JsonElement args)
    {
        var path = Program.ResolvePath(GetString(args, "path") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "path")));
        var text = GetString(args, "text") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "text"));
        var regex = GetBoolean(args, "regex") ?? false;
        var maxResults = GetInt32(args, "maxResults") ?? 100;

        var result = OfficeDocumentTool.Search(path, text, regex, maxResults);
        return result.Success ? ToolText(id, result.Content!) : ToolError(id, result);
    }

    private static String ApplyPatch(JsonElement id, JsonElement args)
    {
        var path = Program.ResolvePath(GetString(args, "path") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "path")));
        var output = Program.ResolvePath(GetString(args, "output") ?? throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "output")));
        if (!args.TryGetProperty("patch", out var patchEl)) throw new ArgumentException(Localization.Translate("Missing required argument: '{0}'.", "patch"));

        var result = OfficeDocumentTool.ApplyPatch(path, patchEl.GetRawText(), output);
        if (!result.Success) return ToolError(id, result);

        var text = $"已应用补丁 → {output}";
        if (!result.Content.IsNullOrEmpty()) text += "\n" + result.Content;
        if (result.Warnings.Count > 0) text += "\n告警：\n" + String.Join("\n", result.Warnings);
        return ToolText(id, text);
    }

    private static String ToolError(JsonElement id, OfficeToolResult result)
        => ToolText(id, $"失败（{result.Code}）：{result.Message}\n未写出文件。", isError: true);

    private static Int32? GetInt32(JsonElement args, String name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var el)) return null;
        if (el.ValueKind == JsonValueKind.Number) return el.GetInt32();
        if (el.ValueKind == JsonValueKind.String && Int32.TryParse(el.GetString(), out var parsed)) return parsed;
        return null;
    }

    private static Boolean? GetBoolean(JsonElement args, String name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var el)) return null;
        if (el.ValueKind == JsonValueKind.True) return true;
        if (el.ValueKind == JsonValueKind.False) return false;
        return null;
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
