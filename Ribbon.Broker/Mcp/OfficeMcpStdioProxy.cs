using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Ribbon.Broker.Infrastructure;
using Ribbon.Contracts;

namespace Ribbon.Broker.Mcp;

internal sealed class OfficeMcpStdioProxy
{
    private readonly string _preferredHostId;
    private readonly BrokerLog _log;
    private readonly Dictionary<string, OfficeToolDefinition> _tools = new(StringComparer.OrdinalIgnoreCase);

    public OfficeMcpStdioProxy(string preferredHostId, BrokerLog log)
    {
        _preferredHostId = preferredHostId;
        _log = log;
    }

    public async Task<int> RunAsync(CancellationToken cancellationToken)
    {
        using var pipe = new NamedPipeClientStream(".", RibbonProtocol.PipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
        await pipe.ConnectAsync(10000, cancellationToken).ConfigureAwait(false);
        await using var broker = new PipePeer(pipe, _log);
        broker.Start((_, request, _) => Task.FromResult<RpcEnvelope?>(RpcEnvelope.Failure(request, "The MCP proxy does not accept broker requests.")), cancellationToken);

        using var input = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        using var output = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };

        while (!cancellationToken.IsCancellationRequested)
        {
            var line = await input.ReadLineAsync(cancellationToken).ConfigureAwait(false);
            if (line == null) break;

            JsonElement message;
            try
            {
                using var document = JsonDocument.Parse(line);
                message = document.RootElement.Clone();
            }
            catch (Exception exception)
            {
                await WriteErrorAsync(output, null, -32700, exception.Message, cancellationToken).ConfigureAwait(false);
                continue;
            }

            if (!message.TryGetProperty("method", out var methodElement)) continue;
            var method = methodElement.GetString() ?? string.Empty;
            var hasId = message.TryGetProperty("id", out var id);
            var parameters = message.TryGetProperty("params", out var paramsElement) ? paramsElement : default;

            if (!hasId)
            {
                continue;
            }

            try
            {
                object result = method switch
                {
                    "initialize" => Initialize(parameters),
                    "ping" => new { },
                    "tools/list" => await ListToolsAsync(broker, cancellationToken).ConfigureAwait(false),
                    "tools/call" => await CallToolAsync(broker, parameters, cancellationToken).ConfigureAwait(false),
                    _ => throw new McpMethodNotFoundException(method)
                };
                await WriteResultAsync(output, id, result, cancellationToken).ConfigureAwait(false);
            }
            catch (McpMethodNotFoundException exception)
            {
                await WriteErrorAsync(output, id, -32601, exception.Message, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                _log.Error($"Office MCP method '{method}' failed.", exception);
                await WriteErrorAsync(output, id, -32603, exception.Message, cancellationToken).ConfigureAwait(false);
            }
        }

        return 0;
    }

    private static object Initialize(JsonElement parameters)
    {
        var requestedVersion = parameters.ValueKind == JsonValueKind.Object
            && parameters.TryGetProperty("protocolVersion", out var version)
            ? version.GetString()
            : null;
        return new
        {
            protocolVersion = requestedVersion ?? "2025-06-18",
            capabilities = new { tools = new { listChanged = false } },
            serverInfo = new { name = "ribbon-office", title = "Ribbon for Microsoft Office", version = RibbonProtocol.ProductVersion },
            instructions = "Use these tools to inspect and modify the user's currently connected Microsoft Office applications. Prefer precise reads before writes and clearly summarize user-visible changes. In Excel, use excel_write_formulas for formulas, treat excel_format_range as a patch that preserves unspecified styling, and create charts or tables only after inspecting their source range. In Word, inspect context and headings before structural edits, treat word_format_range as a patch, and refresh character positions after mutations because later text moves as the document changes. In PowerPoint, inspect the slide outline and target slide before editing, use returned shape names for later mutations, treat powerpoint_format_shape as a patch, and refresh slide numbers after moving, duplicating, or deleting slides."
        };
    }

    private async Task<object> ListToolsAsync(PipePeer broker, CancellationToken cancellationToken)
    {
        var response = await broker.RequestAsync(RibbonProtocol.ListTools, JsonCodec.Serialize(new HostIdRequest { HostId = _preferredHostId }), cancellationToken).ConfigureAwait(false);
        var definitions = JsonCodec.Deserialize<List<OfficeToolDefinition>>(response.Payload);
        _tools.Clear();
        foreach (var definition in definitions)
        {
            _tools[definition.Name] = definition;
        }

        var tools = new JsonArray();
        foreach (var definition in definitions)
        {
            tools.Add(new JsonObject
            {
                ["name"] = definition.Name,
                ["description"] = definition.Description,
                ["inputSchema"] = JsonNode.Parse(string.IsNullOrWhiteSpace(definition.InputSchemaJson) ? "{\"type\":\"object\"}" : definition.InputSchemaJson),
                ["annotations"] = new JsonObject
                {
                    ["destructiveHint"] = definition.Destructive,
                    ["readOnlyHint"] = !definition.Destructive
                }
            });
        }
        return new JsonObject { ["tools"] = tools };
    }

    private async Task<object> CallToolAsync(PipePeer broker, JsonElement parameters, CancellationToken cancellationToken)
    {
        var name = parameters.GetProperty("name").GetString()
            ?? throw new InvalidDataException("MCP tools/call requires a tool name.");
        if (!_tools.TryGetValue(name, out var definition))
        {
            await ListToolsAsync(broker, cancellationToken).ConfigureAwait(false);
            if (!_tools.TryGetValue(name, out definition))
            {
                throw new InvalidOperationException($"Office tool '{name}' is not available.");
            }
        }
        var argumentsJson = parameters.TryGetProperty("arguments", out var arguments) ? arguments.GetRawText() : "{}";
        var invocation = new OfficeToolInvocation
        {
            HostId = definition.HostId,
            ToolName = name,
            ArgumentsJson = argumentsJson
        };
        var response = await broker.RequestAsync(RibbonProtocol.InvokeTool, JsonCodec.Serialize(invocation), cancellationToken).ConfigureAwait(false);
        var result = JsonCodec.Deserialize<OfficeToolResult>(response.Payload);
        var contentText = result.Success ? result.ContentJson ?? "{}" : result.Error ?? "Office tool failed.";
        JsonNode? structured = null;
        if (result.Success)
        {
            try { structured = JsonNode.Parse(contentText); } catch { }
        }
        return new JsonObject
        {
            ["content"] = new JsonArray(new JsonObject { ["type"] = "text", ["text"] = contentText }),
            ["structuredContent"] = structured,
            ["isError"] = !result.Success
        };
    }

    private static Task WriteResultAsync(StreamWriter output, JsonElement id, object result, CancellationToken cancellationToken)
    {
        return WriteAsync(output, new Dictionary<string, object?>
        {
            ["jsonrpc"] = "2.0",
            ["id"] = IdValue(id),
            ["result"] = result
        }, cancellationToken);
    }

    private static Task WriteErrorAsync(StreamWriter output, JsonElement? id, int code, string message, CancellationToken cancellationToken)
    {
        return WriteAsync(output, new Dictionary<string, object?>
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id.HasValue ? IdValue(id.Value) : null,
            ["error"] = new { code, message }
        }, cancellationToken);
    }

    private static async Task WriteAsync(StreamWriter output, object value, CancellationToken cancellationToken)
    {
        await output.WriteLineAsync(JsonSerializer.Serialize(value, JsonCodec.Options).AsMemory(), cancellationToken).ConfigureAwait(false);
        await output.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static object? IdValue(JsonElement id)
    {
        return id.ValueKind switch
        {
            JsonValueKind.String => id.GetString(),
            JsonValueKind.Number => id.TryGetInt64(out var number) ? number : id.GetDouble(),
            _ => null
        };
    }

    private sealed class McpMethodNotFoundException : Exception
    {
        public McpMethodNotFoundException(string method) : base($"Unknown MCP method '{method}'.") { }
    }
}
