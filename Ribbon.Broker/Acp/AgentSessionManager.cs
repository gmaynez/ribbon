using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using Ribbon.Broker.Infrastructure;
using Ribbon.Broker.Registry;
using Ribbon.Broker.Server;
using Ribbon.Contracts;

namespace Ribbon.Broker.Acp;

internal sealed class AgentSessionManager : IAsyncDisposable
{
    private readonly BrokerPaths _paths;
    private readonly BrokerLog _log;
    private readonly InstalledAgentStore _store;
    private readonly HostRegistry _hosts;
    private readonly ConcurrentDictionary<string, AgentRuntime> _runtimes = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, SessionContext> _sessions = new(StringComparer.Ordinal);

    public AgentSessionManager(BrokerPaths paths, BrokerLog log, InstalledAgentStore store, HostRegistry hosts)
    {
        _paths = paths;
        _log = log;
        _store = store;
        _hosts = hosts;
    }

    public async Task<SessionStartResponse> StartAsync(SessionStartRequest request, PipePeer client, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.AgentId) || string.IsNullOrWhiteSpace(request.HostId))
        {
            throw new ArgumentException("An agent and Office host are required.");
        }

        var installed = await _store.FindAsync(request.AgentId, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException($"ACP agent '{request.AgentId}' is not installed.");
        var runtime = _runtimes.GetOrAdd(installed.Id, _ => CreateRuntime(installed));
        await runtime.InitializeAsync(cancellationToken).ConfigureAwait(false);

        var host = _hosts.Get(request.HostId);
        var workingDirectory = CreateSessionDirectory(request, host.Registration);
        JsonElement result;
        try
        {
            result = await runtime.Connection.RequestAsync("session/new", new
            {
                cwd = workingDirectory,
                mcpServers = new[]
                {
                    new
                    {
                        name = "ribbon-office",
                        command = Environment.ProcessPath ?? throw new InvalidOperationException("Unable to locate Ribbon.Broker.exe."),
                        args = new[] { "--mcp-stdio", "--host-id", request.HostId },
                        env = Array.Empty<object>()
                    }
                }
            }, cancellationToken).ConfigureAwait(false);
        }
        catch (AcpRpcException exception) when (runtime.AuthenticationMethods.Count > 0)
        {
            _log.Info($"Agent {installed.Id} requires authentication before creating a session: {exception.Message}");
            return new SessionStartResponse
            {
                SessionId = string.Empty,
                AgentName = installed.Name,
                AuthenticationMethods = runtime.AuthenticationMethods.ToList()
            };
        }

        var sessionId = result.GetProperty("sessionId").GetString()
            ?? throw new InvalidDataException("The ACP agent did not return a session id.");
        _sessions[sessionId] = new SessionContext(sessionId, request.HostId, client, runtime);
        return new SessionStartResponse
        {
            SessionId = sessionId,
            AgentName = installed.Name,
            AuthenticationMethods = runtime.AuthenticationMethods.ToList()
        };
    }

    public async Task AuthenticateAsync(AgentAuthenticationRequest request, CancellationToken cancellationToken)
    {
        if (!_runtimes.TryGetValue(request.AgentId, out var runtime))
        {
            var installed = await _store.FindAsync(request.AgentId, cancellationToken).ConfigureAwait(false)
                ?? throw new InvalidOperationException($"ACP agent '{request.AgentId}' is not installed.");
            runtime = _runtimes.GetOrAdd(installed.Id, _ => CreateRuntime(installed));
        }
        await runtime.InitializeAsync(cancellationToken).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(request.MethodId))
        {
            throw new ArgumentException("An authentication method is required.");
        }
        await runtime.Connection.RequestAsync("authenticate", new { methodId = request.MethodId }, cancellationToken).ConfigureAwait(false);
    }

    public async Task PromptAsync(SessionPromptRequest request, CancellationToken cancellationToken)
    {
        var session = GetSession(request.SessionId);
        if (string.IsNullOrWhiteSpace(request.Text))
        {
            throw new ArgumentException("A prompt is required.");
        }

        var result = await session.Runtime.Connection.RequestAsync("session/prompt", new
        {
            sessionId = request.SessionId,
            prompt = new[] { new { type = "text", text = request.Text } }
        }, cancellationToken).ConfigureAwait(false);
        var stopReason = result.TryGetProperty("stopReason", out var stop) ? stop.GetString() : "end_turn";
        await SendUpdateAsync(session, new SessionUpdateMessage
        {
            SessionId = request.SessionId,
            UpdateKind = "turn_complete",
            Status = stopReason ?? "end_turn",
            RawJson = result.GetRawText()
        }, cancellationToken).ConfigureAwait(false);
    }

    public Task CancelAsync(SessionCancelRequest request, CancellationToken cancellationToken)
    {
        var session = GetSession(request.SessionId);
        return session.Runtime.Connection.NotifyAsync("session/cancel", new { sessionId = request.SessionId }, cancellationToken);
    }

    public async ValueTask DisposeAsync()
    {
        foreach (var runtime in _runtimes.Values)
        {
            await runtime.DisposeAsync().ConfigureAwait(false);
        }
        _runtimes.Clear();
        _sessions.Clear();
    }

    private AgentRuntime CreateRuntime(InstalledAgentRecord installed)
    {
        var workingDirectory = Path.Combine(_paths.Sessions, "agents", installed.Id);
        Directory.CreateDirectory(workingDirectory);
        var runtime = new AgentRuntime(installed, workingDirectory, _log);
        runtime.Connection.NotificationReceived = (method, parameters, cancellationToken) =>
            HandleNotificationAsync(runtime, method, parameters, cancellationToken);
        runtime.Connection.RequestReceived = (method, parameters, cancellationToken) =>
            HandleRequestAsync(runtime, method, parameters, cancellationToken);
        return runtime;
    }

    private async Task HandleNotificationAsync(AgentRuntime runtime, string method, JsonElement parameters, CancellationToken cancellationToken)
    {
        if (!string.Equals(method, "session/update", StringComparison.Ordinal))
        {
            return;
        }
        var sessionId = parameters.TryGetProperty("sessionId", out var sessionElement) ? sessionElement.GetString() : null;
        if (string.IsNullOrWhiteSpace(sessionId) || !_sessions.TryGetValue(sessionId, out var session))
        {
            return;
        }
        var update = parameters.GetProperty("update");
        var kind = update.TryGetProperty("sessionUpdate", out var kindElement) ? kindElement.GetString() ?? "unknown" : "unknown";
        var message = new SessionUpdateMessage
        {
            SessionId = sessionId,
            UpdateKind = kind,
            RawJson = update.GetRawText()
        };

        if ((kind == "agent_message_chunk" || kind == "agent_thought_chunk" || kind == "user_message_chunk")
            && update.TryGetProperty("content", out var content))
        {
            message.Text = ExtractContentText(content);
        }
        else if (kind == "tool_call" || kind == "tool_call_update")
        {
            message.ToolName = update.TryGetProperty("title", out var title) ? title.GetString() : null;
            message.Status = update.TryGetProperty("status", out var status) ? status.GetString() : null;
            if (string.IsNullOrWhiteSpace(message.ToolName) && update.TryGetProperty("toolCallId", out var toolCallId))
            {
                message.ToolName = toolCallId.GetString();
            }
        }
        await SendUpdateAsync(session, message, cancellationToken).ConfigureAwait(false);
    }

    private async Task<object?> HandleRequestAsync(AgentRuntime runtime, string method, JsonElement parameters, CancellationToken cancellationToken)
    {
        if (!string.Equals(method, "session/request_permission", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Ribbon does not expose ACP client method '{method}'.");
        }

        var sessionId = parameters.GetProperty("sessionId").GetString()
            ?? throw new InvalidDataException("The permission request did not contain a session id.");
        var session = GetSession(sessionId);
        var toolCall = parameters.GetProperty("toolCall");
        var prompt = new PermissionPrompt
        {
            SessionId = sessionId,
            ToolCallId = toolCall.TryGetProperty("toolCallId", out var id) ? id.GetString() ?? string.Empty : string.Empty,
            Title = toolCall.TryGetProperty("title", out var title) ? title.GetString() ?? "Agent action" : "Agent action",
            RawJson = toolCall.GetRawText(),
            Options = parameters.GetProperty("options").EnumerateArray().Select(option => new PermissionChoice
            {
                OptionId = option.GetProperty("optionId").GetString() ?? string.Empty,
                Name = option.GetProperty("name").GetString() ?? string.Empty,
                Kind = option.GetProperty("kind").GetString() ?? string.Empty
            }).ToList()
        };

        var response = await session.Client.RequestAsync(RibbonProtocol.PermissionRequest, JsonCodec.Serialize(prompt), cancellationToken).ConfigureAwait(false);
        var decision = JsonCodec.Deserialize<PermissionDecision>(response.Payload);
        return decision.Cancelled
            ? new { outcome = new { outcome = "cancelled" } }
            : new { outcome = new { outcome = "selected", optionId = decision.OptionId } };
    }

    private string CreateSessionDirectory(SessionStartRequest request, HostRegistration host)
    {
        var root = Path.Combine(_paths.Sessions, Sanitize(host.HostKind), Sanitize(host.HostId));
        Directory.CreateDirectory(root);
        var guidance = new StringBuilder()
            .AppendLine("# Ribbon Office Session")
            .AppendLine()
            .AppendLine($"You are operating Microsoft {host.HostKind} through the Ribbon Office MCP server.")
            .AppendLine("Use the ribbon-office MCP tools for Office inspection and changes.")
            .AppendLine("Do not attempt to automate Office through shell scripts, UI automation, or direct file-format editing when a Ribbon tool is available.")
            .AppendLine("Treat write, insert, delete, formatting, and structural operations as user-visible changes and describe them precisely.")
            .AppendLine("If the required Office application or document is unavailable, explain that clearly instead of guessing.")
            .ToString();
        File.WriteAllText(Path.Combine(root, "AGENTS.md"), guidance);
        return root;
    }

    private SessionContext GetSession(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !_sessions.TryGetValue(sessionId, out var session))
        {
            throw new InvalidOperationException("The Ribbon agent session was not found.");
        }
        return session;
    }

    private static Task SendUpdateAsync(SessionContext session, SessionUpdateMessage message, CancellationToken cancellationToken)
    {
        return session.Client.NotifyAsync(RibbonProtocol.SessionUpdate, JsonCodec.Serialize(message), cancellationToken);
    }

    private static string ExtractContentText(JsonElement content)
    {
        if (content.ValueKind == JsonValueKind.Object
            && content.TryGetProperty("type", out var type)
            && type.GetString() == "text"
            && content.TryGetProperty("text", out var text))
        {
            return text.GetString() ?? string.Empty;
        }
        return string.Empty;
    }

    private static string Sanitize(string value)
    {
        return string.Concat((value ?? string.Empty).Select(character => char.IsLetterOrDigit(character) || character == '-' ? character : '_'));
    }

    private sealed record SessionContext(string SessionId, string HostId, PipePeer Client, AgentRuntime Runtime);
}

internal sealed class AgentRuntime : IAsyncDisposable
{
    private readonly SemaphoreSlim _initializeGate = new(1, 1);
    private bool _initialized;

    public AgentRuntime(InstalledAgentRecord agent, string workingDirectory, BrokerLog log)
    {
        Agent = agent;
        Connection = new AcpProcessConnection(agent, workingDirectory, log);
    }

    public InstalledAgentRecord Agent { get; }
    public AcpProcessConnection Connection { get; }
    public IReadOnlyList<AgentAuthenticationMethod> AuthenticationMethods { get; private set; } = Array.Empty<AgentAuthenticationMethod>();

    public async Task InitializeAsync(CancellationToken cancellationToken)
    {
        if (_initialized) return;
        await _initializeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_initialized) return;
            var result = await Connection.RequestAsync("initialize", new
            {
                protocolVersion = 1,
                clientCapabilities = new
                {
                    fs = new { readTextFile = false, writeTextFile = false },
                    terminal = false
                },
                clientInfo = new { name = "ribbon", title = "Ribbon for Microsoft Office", version = RibbonProtocol.ProductVersion }
            }, cancellationToken).ConfigureAwait(false);
            var negotiated = result.GetProperty("protocolVersion").GetInt32();
            if (negotiated != 1)
            {
                throw new NotSupportedException($"Agent '{Agent.Name}' negotiated unsupported ACP protocol version {negotiated}.");
            }
            AuthenticationMethods = result.TryGetProperty("authMethods", out var methods) && methods.ValueKind == JsonValueKind.Array
                ? methods.EnumerateArray().Select(method => new AgentAuthenticationMethod
                {
                    Id = method.GetProperty("id").GetString() ?? string.Empty,
                    Name = method.GetProperty("name").GetString() ?? string.Empty,
                    Description = method.TryGetProperty("description", out var description) ? description.GetString() ?? string.Empty : string.Empty
                }).ToList()
                : Array.Empty<AgentAuthenticationMethod>();
            _initialized = true;
        }
        finally
        {
            _initializeGate.Release();
        }
    }

    public async ValueTask DisposeAsync()
    {
        await Connection.DisposeAsync().ConfigureAwait(false);
        _initializeGate.Dispose();
    }
}
