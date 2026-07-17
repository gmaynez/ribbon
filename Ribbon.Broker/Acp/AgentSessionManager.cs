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
    private readonly SemaphoreSlim _runtimeGate = new(1, 1);
    private int _disposed;

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
        await _runtimeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
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
            catch (AcpRpcException exception) when (IsAuthenticationRequired(exception, runtime.AuthenticationMethods.Count))
            {
                _log.Info($"Agent {installed.Id} requires authentication before creating a session: {exception.Message}");
                return new SessionStartResponse
                {
                    SessionId = string.Empty,
                    AgentName = installed.Name,
                    WorkingDirectory = workingDirectory,
                    SupportsLoad = runtime.SupportsSessionLoad,
                    SupportsResume = runtime.SupportsSessionResume,
                    SupportsList = runtime.SupportsSessionList,
                    AuthenticationMethods = runtime.AuthenticationMethods.ToList()
                };
            }

            var sessionId = result.GetProperty("sessionId").GetString()
                ?? throw new InvalidDataException("The ACP agent did not return a session id.");
            var configOptions = AcpSessionConfig.Parse(result);
            var sessionContext = new SessionContext(sessionId, request.HostId, client, runtime);
            sessionContext.ReplaceConfigOptions(configOptions);
            _sessions[sessionId] = sessionContext;
            return new SessionStartResponse
            {
                SessionId = sessionId,
                AgentName = installed.Name,
                WorkingDirectory = workingDirectory,
                SupportsLoad = runtime.SupportsSessionLoad,
                SupportsResume = runtime.SupportsSessionResume,
                SupportsList = runtime.SupportsSessionList,
                AuthenticationMethods = runtime.AuthenticationMethods.ToList(),
                ConfigOptions = configOptions
            };
        }
        finally
        {
            _runtimeGate.Release();
        }
    }

    public async Task<SessionResumeResponse> ResumeAsync(SessionResumeRequest request, PipePeer client, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.AgentId)
            || string.IsNullOrWhiteSpace(request.HostId)
            || string.IsNullOrWhiteSpace(request.SessionId))
        {
            throw new ArgumentException("An agent, Office host, and saved ACP session are required.");
        }

        var installed = await _store.FindAsync(request.AgentId, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException($"ACP agent '{request.AgentId}' is not installed.");
        await _runtimeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
            var runtime = _runtimes.GetOrAdd(installed.Id, _ => CreateRuntime(installed));
            await runtime.InitializeAsync(cancellationToken).ConfigureAwait(false);
            var workingDirectory = ValidateSessionDirectory(request.WorkingDirectory);
            _hosts.Get(request.HostId);
            var method = runtime.SupportsSessionResume
                ? "session/resume"
                : runtime.SupportsSessionLoad ? "session/load" : null;
            if (method == null)
            {
                return ResumeResponse(runtime, installed, request, workingDirectory, false, "unsupported",
                    "This ACP agent does not support restoring previous sessions.", []);
            }

            var context = new SessionContext(request.SessionId, request.HostId, client, runtime)
            {
                SuppressUpdates = string.Equals(method, "session/load", StringComparison.Ordinal)
            };
            if (!_sessions.TryAdd(request.SessionId, context))
            {
                return ResumeResponse(runtime, installed, request, workingDirectory, false, "already_active",
                    "The saved ACP session is already active.", []);
            }

            try
            {
                var result = await runtime.Connection.RequestAsync(method, new
                {
                    sessionId = request.SessionId,
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
                var configOptions = result.ValueKind == JsonValueKind.Object ? AcpSessionConfig.Parse(result) : [];
                if (configOptions.Count == 0) configOptions = context.GetConfigOptions().ToList();
                context.ReplaceConfigOptions(configOptions);
                context.SuppressUpdates = false;
                return ResumeResponse(runtime, installed, request, workingDirectory, true,
                    string.Equals(method, "session/load", StringComparison.Ordinal) ? "loaded" : "resumed",
                    null, configOptions);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                _sessions.TryRemove(new KeyValuePair<string, SessionContext>(request.SessionId, context));
                context.PendingPermissions.CancelAll();
                throw;
            }
            catch (Exception exception)
            {
                _sessions.TryRemove(new KeyValuePair<string, SessionContext>(request.SessionId, context));
                context.PendingPermissions.CancelAll();
                _log.Error($"Agent {installed.Id} could not restore ACP session {request.SessionId}.", exception);
                return ResumeResponse(runtime, installed, request, workingDirectory, false, "unavailable",
                    exception.GetBaseException().Message, []);
            }
        }
        finally
        {
            _runtimeGate.Release();
        }
    }

    public async Task<AgentSessionListResponse> ListSessionsAsync(AgentSessionListRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.AgentId)) throw new ArgumentException("An ACP agent is required.");
        var installed = await _store.FindAsync(request.AgentId, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException($"ACP agent '{request.AgentId}' is not installed.");
        await _runtimeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
            var runtime = _runtimes.GetOrAdd(installed.Id, _ => CreateRuntime(installed));
            await runtime.InitializeAsync(cancellationToken).ConfigureAwait(false);
            if (!runtime.SupportsSessionList)
            {
                return new AgentSessionListResponse { Supported = false, Complete = false, Sessions = [] };
            }

            var workingDirectory = string.IsNullOrWhiteSpace(request.WorkingDirectory)
                ? null
                : ValidateSessionDirectory(request.WorkingDirectory);
            var sessions = new List<AgentSessionSummary>();
            var seenCursors = new HashSet<string>(StringComparer.Ordinal);
            var complete = false;
            string? cursor = null;
            for (var page = 0; page < 100; page++)
            {
                var parameters = new Dictionary<string, object?>();
                if (workingDirectory != null) parameters["cwd"] = workingDirectory;
                if (cursor != null) parameters["cursor"] = cursor;
                var result = await runtime.Connection.RequestAsync("session/list", parameters, cancellationToken).ConfigureAwait(false);
                if (result.TryGetProperty("sessions", out var items) && items.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in items.EnumerateArray())
                    {
                        var id = OptionalString(item, "sessionId");
                        if (string.IsNullOrWhiteSpace(id)) continue;
                        sessions.Add(new AgentSessionSummary
                        {
                            SessionId = id,
                            WorkingDirectory = OptionalString(item, "cwd"),
                            Title = OptionalString(item, "title"),
                            UpdatedAt = OptionalString(item, "updatedAt")
                        });
                    }
                }
                cursor = OptionalString(result, "nextCursor");
                if (string.IsNullOrWhiteSpace(cursor))
                {
                    complete = true;
                    break;
                }
                if (!seenCursors.Add(cursor)) break;
            }
            return new AgentSessionListResponse { Supported = true, Complete = complete, Sessions = sessions };
        }
        finally
        {
            _runtimeGate.Release();
        }
    }

    public async Task AuthenticateAsync(AgentAuthenticationRequest request, CancellationToken cancellationToken)
    {
        await _runtimeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
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
        finally
        {
            _runtimeGate.Release();
        }
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

    public async Task CancelAsync(SessionCancelRequest request, CancellationToken cancellationToken)
    {
        var session = GetSession(request.SessionId);
        session.PendingPermissions.CancelAll();
        await session.Runtime.Connection.NotifyAsync("session/cancel", new { sessionId = request.SessionId }, cancellationToken).ConfigureAwait(false);
    }

    public async Task CloseAsync(SessionCancelRequest request, CancellationToken cancellationToken)
    {
        var session = GetSession(request.SessionId);
        if (!_sessions.TryRemove(new KeyValuePair<string, SessionContext>(request.SessionId, session))) return;
        session.PendingPermissions.CancelAll();

        var closedByAgent = false;
        if (session.Runtime.SupportsSessionClose)
        {
            try
            {
                await session.Runtime.Connection.RequestAsync(
                    "session/close",
                    new { sessionId = request.SessionId },
                    cancellationToken).ConfigureAwait(false);
                closedByAgent = true;
            }
            catch (Exception exception)
            {
                _log.Error($"Agent {session.Runtime.Agent.Id} failed to close ACP session {request.SessionId}; Ribbon will retire the runtime when possible.", exception);
            }
        }
        else
        {
            try
            {
                await session.Runtime.Connection.NotifyAsync(
                    "session/cancel",
                    new { sessionId = request.SessionId },
                    cancellationToken).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                _log.Error($"Agent {session.Runtime.Agent.Id} failed to cancel ACP session {request.SessionId} before retirement.", exception);
            }
        }

        if (closedByAgent) return;
        await RetireUnusedRuntimeAsync(session.Runtime).ConfigureAwait(false);
    }

    public async Task<SessionConfigOptionsResponse> SetConfigOptionAsync(SessionConfigOptionRequest request, CancellationToken cancellationToken)
    {
        var session = GetSession(request.SessionId);
        AcpSessionConfig.RequireSelectValue(session.GetConfigOptions(), request.ConfigId, request.Value);

        var result = await session.Runtime.Connection.RequestAsync("session/set_config_option", new
        {
            sessionId = request.SessionId,
            configId = request.ConfigId,
            value = request.Value
        }, cancellationToken).ConfigureAwait(false);
        var configOptions = AcpSessionConfig.Parse(result);
        if (configOptions.Count == 0)
        {
            throw new InvalidDataException("The ACP agent did not return the complete session configuration state.");
        }
        session.ReplaceConfigOptions(configOptions);
        return new SessionConfigOptionsResponse { ConfigOptions = configOptions };
    }

    public async ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        await _runtimeGate.WaitAsync().ConfigureAwait(false);
        try
        {
            foreach (var session in _sessions.Values)
            {
                session.PendingPermissions.CancelAll();
            }
            _sessions.Clear();
            foreach (var runtime in _runtimes.Values)
            {
                await runtime.DisposeAsync().ConfigureAwait(false);
            }
            _runtimes.Clear();
        }
        finally
        {
            _runtimeGate.Release();
        }
    }

    public async Task ReleaseClientAsync(PipePeer client)
    {
        if (client == null || Volatile.Read(ref _disposed) != 0) return;

        var releasedSessions = 0;
        await _runtimeGate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (Volatile.Read(ref _disposed) != 0) return;
            foreach (var pair in _sessions.Where(pair => ReferenceEquals(pair.Value.Client, client)).ToArray())
            {
                if (_sessions.TryRemove(pair))
                {
                    pair.Value.PendingPermissions.CancelAll();
                    releasedSessions++;
                }
            }

            var activeRuntimes = _sessions.Values.Select(session => session.Runtime).ToHashSet();
            foreach (var pair in _runtimes.ToArray())
            {
                if (activeRuntimes.Contains(pair.Value) || !_runtimes.TryRemove(pair)) continue;
                await pair.Value.DisposeAsync().ConfigureAwait(false);
            }
        }
        finally
        {
            _runtimeGate.Release();
        }

        if (releasedSessions > 0)
        {
            _log.Info($"Released {releasedSessions} ACP session(s) for a disconnected Office client.");
        }
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
        var message = ParseSessionUpdate(sessionId, update);

        if (message.UpdateKind == "config_option_update")
        {
            session.ReplaceConfigOptions(message.ConfigOptions ?? []);
        }
        else if (message.UpdateKind == "tool_call" || message.UpdateKind == "tool_call_update")
        {
            session.HydrateToolCall(message);
        }
        if (!session.SuppressUpdates)
        {
            await SendUpdateAsync(session, message, cancellationToken).ConfigureAwait(false);
        }
    }

    internal static SessionUpdateMessage ParseSessionUpdate(string sessionId, JsonElement update)
    {
        var kind = update.TryGetProperty("sessionUpdate", out var kindElement)
            ? kindElement.GetString() ?? "unknown"
            : "unknown";
        var message = new SessionUpdateMessage
        {
            SessionId = sessionId,
            UpdateKind = kind,
            MessageId = update.TryGetProperty("messageId", out var messageId) ? messageId.GetString() : null,
            RawJson = update.GetRawText()
        };

        if (kind == "config_option_update")
        {
            message.ConfigOptions = AcpSessionConfig.Parse(update);
        }
        else if ((kind == "agent_message_chunk" || kind == "agent_thought_chunk" || kind == "user_message_chunk")
            && update.TryGetProperty("content", out var content))
        {
            message.Text = ExtractContentText(content);
        }
        else if (kind == "tool_call" || kind == "tool_call_update")
        {
            message.ToolCallId = update.TryGetProperty("toolCallId", out var toolCallId) ? toolCallId.GetString() : null;
            message.ToolName = update.TryGetProperty("title", out var title) ? title.GetString() : null;
            message.ToolKind = update.TryGetProperty("kind", out var toolKind) ? toolKind.GetString() : null;
            message.Status = update.TryGetProperty("status", out var status) ? status.GetString() : null;
            message.Text = ExtractToolCallText(update);
        }
        else if (kind == "plan" && update.TryGetProperty("entries", out var entries) && entries.ValueKind == JsonValueKind.Array)
        {
            message.PlanEntries = entries.EnumerateArray().Select(entry => new SessionPlanEntry
            {
                Content = entry.TryGetProperty("content", out var entryContent) ? entryContent.GetString() ?? string.Empty : string.Empty,
                Priority = entry.TryGetProperty("priority", out var priority) ? priority.GetString() : null,
                Status = entry.TryGetProperty("status", out var entryStatus) ? entryStatus.GetString() : null
            }).ToList();
        }
        else if (kind == "session_info_update")
        {
            message.Title = OptionalString(update, "title");
            message.UpdatedAt = OptionalString(update, "updatedAt");
        }

        return message;
    }

    private async Task<object?> HandleRequestAsync(AgentRuntime runtime, string method, JsonElement parameters, CancellationToken cancellationToken)
    {
        if (!string.Equals(method, "session/request_permission", StringComparison.Ordinal))
        {
            throw new AcpRpcException(-32601, $"Ribbon does not expose ACP client method '{method}'.", null);
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

        using var pendingPermission = session.PendingPermissions.Register(cancellationToken);
        try
        {
            var response = await session.Client.RequestAsync(
                RibbonProtocol.PermissionRequest,
                JsonCodec.Serialize(prompt),
                pendingPermission.Token).ConfigureAwait(false);
            var decision = JsonCodec.Deserialize<PermissionDecision>(response.Payload);
            return decision.Cancelled
                ? new { outcome = new { outcome = "cancelled" } }
                : new { outcome = new { outcome = "selected", optionId = decision.OptionId } };
        }
        catch (OperationCanceledException) when (pendingPermission.CancelledByClient)
        {
            return new { outcome = new { outcome = "cancelled" } };
        }
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

    private string ValidateSessionDirectory(string workingDirectory)
    {
        if (string.IsNullOrWhiteSpace(workingDirectory))
        {
            throw new InvalidOperationException("The saved conversation does not contain its ACP working directory.");
        }
        var root = Path.GetFullPath(_paths.Sessions).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        var path = Path.GetFullPath(workingDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase) || !Directory.Exists(path))
        {
            throw new InvalidOperationException("The saved ACP working directory is unavailable or outside Ribbon's session storage.");
        }
        var rootPath = root.TrimEnd(Path.DirectorySeparatorChar);
        for (var current = new DirectoryInfo(path); current != null; current = current.Parent)
        {
            if ((current.Attributes & FileAttributes.ReparsePoint) != 0)
            {
                throw new InvalidOperationException("The saved ACP working directory cannot traverse a symbolic link or junction.");
            }
            if (string.Equals(current.FullName.TrimEnd(Path.DirectorySeparatorChar), rootPath, StringComparison.OrdinalIgnoreCase)) break;
        }
        return path;
    }

    private static SessionResumeResponse ResumeResponse(
        AgentRuntime runtime,
        InstalledAgentRecord installed,
        SessionResumeRequest request,
        string workingDirectory,
        bool resumed,
        string kind,
        string? error,
        IList<SessionConfigOption> configOptions)
    {
        return new SessionResumeResponse
        {
            Resumed = resumed,
            ResumeKind = kind,
            Error = error,
            SessionId = resumed ? request.SessionId : string.Empty,
            AgentName = installed.Name,
            WorkingDirectory = workingDirectory,
            SupportsLoad = runtime.SupportsSessionLoad,
            SupportsResume = runtime.SupportsSessionResume,
            SupportsList = runtime.SupportsSessionList,
            AuthenticationMethods = runtime.AuthenticationMethods.ToList(),
            ConfigOptions = configOptions
        };
    }

    private SessionContext GetSession(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !_sessions.TryGetValue(sessionId, out var session))
        {
            throw new InvalidOperationException("The Ribbon agent session was not found.");
        }
        return session;
    }

    private void ThrowIfDisposed()
    {
        if (Volatile.Read(ref _disposed) != 0)
        {
            throw new ObjectDisposedException(nameof(AgentSessionManager));
        }
    }

    private async Task RetireUnusedRuntimeAsync(AgentRuntime runtime)
    {
        await _runtimeGate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (_sessions.Values.Any(session => ReferenceEquals(session.Runtime, runtime))) return;
            if (_runtimes.TryRemove(new KeyValuePair<string, AgentRuntime>(runtime.Agent.Id, runtime)))
            {
                await runtime.DisposeAsync().ConfigureAwait(false);
            }
        }
        finally
        {
            _runtimeGate.Release();
        }
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

    private static string ExtractToolCallText(JsonElement update)
    {
        if (!update.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array)
        {
            return string.Empty;
        }

        var text = new StringBuilder();
        foreach (var item in content.EnumerateArray())
        {
            JsonElement block;
            if (item.ValueKind == JsonValueKind.Object
                && item.TryGetProperty("type", out var itemType)
                && itemType.GetString() == "content"
                && item.TryGetProperty("content", out var nested))
            {
                block = nested;
            }
            else
            {
                block = item;
            }

            var value = ExtractContentText(block);
            if (string.IsNullOrWhiteSpace(value)) continue;
            if (text.Length > 0) text.AppendLine();
            text.Append(value);
        }
        return text.ToString();
    }

    private static string OptionalString(JsonElement element, string propertyName)
    {
        return element.ValueKind == JsonValueKind.Object
            && element.TryGetProperty(propertyName, out var property)
            && property.ValueKind == JsonValueKind.String
                ? property.GetString() ?? string.Empty
                : string.Empty;
    }

    private static string Sanitize(string value)
    {
        return string.Concat((value ?? string.Empty).Select(character => char.IsLetterOrDigit(character) || character == '-' ? character : '_'));
    }

    internal static bool IsAuthenticationRequired(AcpRpcException exception, int authenticationMethodCount)
    {
        return exception.Code == -32000 && authenticationMethodCount > 0;
    }

    private sealed class SessionContext
    {
        private readonly object _configGate = new();
        private readonly object _toolGate = new();
        private List<SessionConfigOption> _configOptions = [];
        private readonly Dictionary<string, ToolCallState> _toolCalls = new(StringComparer.Ordinal);

        public SessionContext(string sessionId, string hostId, PipePeer client, AgentRuntime runtime)
        {
            SessionId = sessionId;
            HostId = hostId;
            Client = client;
            Runtime = runtime;
        }

        public string SessionId { get; }
        public string HostId { get; }
        public PipePeer Client { get; }
        public AgentRuntime Runtime { get; }
        public PendingPermissionRegistry PendingPermissions { get; } = new();
        public bool SuppressUpdates { get; set; }

        public IReadOnlyList<SessionConfigOption> GetConfigOptions()
        {
            lock (_configGate)
            {
                return _configOptions.ToList();
            }
        }

        public void ReplaceConfigOptions(IEnumerable<SessionConfigOption> configOptions)
        {
            lock (_configGate)
            {
                _configOptions = configOptions?.ToList() ?? [];
            }
        }

        public void HydrateToolCall(SessionUpdateMessage message)
        {
            if (message == null || string.IsNullOrWhiteSpace(message.ToolCallId)) return;
            lock (_toolGate)
            {
                if (!_toolCalls.TryGetValue(message.ToolCallId, out var state))
                {
                    state = new ToolCallState();
                    _toolCalls[message.ToolCallId] = state;
                }

                if (!string.IsNullOrWhiteSpace(message.ToolName)) state.Title = message.ToolName;
                if (!string.IsNullOrWhiteSpace(message.ToolKind)) state.Kind = message.ToolKind;
                if (!string.IsNullOrWhiteSpace(message.Status)) state.Status = message.Status;
                message.ToolName = state.Title ?? message.ToolCallId;
                message.ToolKind = state.Kind;
                message.Status = state.Status;
            }
        }

        private sealed class ToolCallState
        {
            public string? Title { get; set; }
            public string? Kind { get; set; }
            public string? Status { get; set; }
        }
    }
}

internal sealed class PendingPermissionRegistry
{
    private readonly ConcurrentDictionary<long, Registration> _registrations = new();
    private long _nextId;

    public Registration Register(CancellationToken cancellationToken)
    {
        var id = Interlocked.Increment(ref _nextId);
        var registration = new Registration(this, id, cancellationToken);
        if (!_registrations.TryAdd(id, registration))
        {
            registration.Dispose();
            throw new InvalidOperationException("Unable to track an ACP permission request.");
        }
        return registration;
    }

    public void CancelAll()
    {
        foreach (var registration in _registrations.Values)
        {
            registration.CancelByClient();
        }
    }

    private void Remove(long id, Registration registration)
    {
        _registrations.TryRemove(new KeyValuePair<long, Registration>(id, registration));
    }

    internal sealed class Registration : IDisposable
    {
        private readonly object _gate = new();
        private readonly PendingPermissionRegistry _owner;
        private readonly long _id;
        private readonly CancellationTokenSource _cancellation;
        private bool _disposed;
        private bool _cancelledByClient;

        public Registration(PendingPermissionRegistry owner, long id, CancellationToken cancellationToken)
        {
            _owner = owner;
            _id = id;
            _cancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        }

        public CancellationToken Token => _cancellation.Token;
        public bool CancelledByClient
        {
            get
            {
                lock (_gate)
                {
                    return _cancelledByClient;
                }
            }
        }

        public void CancelByClient()
        {
            lock (_gate)
            {
                if (_disposed) return;
                _cancelledByClient = true;
                _cancellation.Cancel();
            }
        }

        public void Dispose()
        {
            lock (_gate)
            {
                if (_disposed) return;
                _disposed = true;
                _owner.Remove(_id, this);
                _cancellation.Dispose();
            }
        }
    }
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
    public bool SupportsSessionClose { get; private set; }
    public bool SupportsSessionLoad { get; private set; }
    public bool SupportsSessionResume { get; private set; }
    public bool SupportsSessionList { get; private set; }

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
            SupportsSessionClose = SupportsSessionCloseCapability(result);
            SupportsSessionLoad = SupportsSessionLoadCapability(result);
            SupportsSessionResume = SupportsSessionCapability(result, "resume");
            SupportsSessionList = SupportsSessionCapability(result, "list");
            _initialized = true;
        }
        finally
        {
            _initializeGate.Release();
        }
    }

    internal static bool SupportsSessionCloseCapability(JsonElement initializeResult)
    {
        return SupportsSessionCapability(initializeResult, "close");
    }

    internal static bool SupportsSessionLoadCapability(JsonElement initializeResult)
    {
        return initializeResult.TryGetProperty("agentCapabilities", out var capabilities)
            && capabilities.ValueKind == JsonValueKind.Object
            && capabilities.TryGetProperty("loadSession", out var load)
            && load.ValueKind == JsonValueKind.True;
    }

    internal static bool SupportsSessionCapability(JsonElement initializeResult, string capabilityName)
    {
        return initializeResult.TryGetProperty("agentCapabilities", out var capabilities)
            && capabilities.ValueKind == JsonValueKind.Object
            && capabilities.TryGetProperty("sessionCapabilities", out var sessions)
            && sessions.ValueKind == JsonValueKind.Object
            && sessions.TryGetProperty(capabilityName, out var capability)
            && (capability.ValueKind == JsonValueKind.Object || capability.ValueKind == JsonValueKind.True);
    }

    public async ValueTask DisposeAsync()
    {
        await Connection.DisposeAsync().ConfigureAwait(false);
        _initializeGate.Dispose();
    }
}
