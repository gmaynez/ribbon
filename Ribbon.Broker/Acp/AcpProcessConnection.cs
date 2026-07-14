using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text.Json;
using Ribbon.Broker.Infrastructure;
using Ribbon.Broker.Registry;

namespace Ribbon.Broker.Acp;

internal sealed class AcpProcessConnection : IAsyncDisposable
{
    private readonly InstalledAgentRecord _agent;
    private readonly string _workingDirectory;
    private readonly BrokerLog _log;
    private readonly ConcurrentDictionary<string, TaskCompletionSource<JsonElement>> _pending = new(StringComparer.Ordinal);
    private readonly SemaphoreSlim _writeGate = new(1, 1);
    private readonly CancellationTokenSource _closed = new();
    private Process? _process;
    private Task? _readLoop;
    private long _nextId;

    public AcpProcessConnection(InstalledAgentRecord agent, string workingDirectory, BrokerLog log)
    {
        _agent = agent;
        _workingDirectory = workingDirectory;
        _log = log;
    }

    public Func<string, JsonElement, CancellationToken, Task<object?>>? RequestReceived { get; set; }
    public Func<string, JsonElement, CancellationToken, Task>? NotificationReceived { get; set; }

    public void Start()
    {
        if (_process != null)
        {
            return;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = _agent.Command,
            WorkingDirectory = _workingDirectory,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (var argument in _agent.Arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }
        foreach (var pair in _agent.Environment)
        {
            startInfo.Environment[pair.Key] = pair.Value;
        }

        _process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Unable to start ACP agent '{_agent.Name}'.");
        _readLoop = ReadLoopAsync(_closed.Token);
        _ = PumpStandardErrorAsync(_process, _closed.Token);
    }

    public async Task<JsonElement> RequestAsync(string method, object parameters, CancellationToken cancellationToken)
    {
        Start();
        var id = Interlocked.Increment(ref _nextId).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var completion = new TaskCompletionSource<JsonElement>(TaskCreationOptions.RunContinuationsAsynchronously);
        if (!_pending.TryAdd(id, completion))
        {
            throw new InvalidOperationException("Unable to allocate an ACP request id.");
        }

        try
        {
            await WriteAsync(new { jsonrpc = "2.0", id = long.Parse(id), method, @params = parameters }, cancellationToken).ConfigureAwait(false);
            return await completion.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _pending.TryRemove(id, out _);
        }
    }

    public Task NotifyAsync(string method, object parameters, CancellationToken cancellationToken)
    {
        Start();
        return WriteAsync(new { jsonrpc = "2.0", method, @params = parameters }, cancellationToken);
    }

    public async ValueTask DisposeAsync()
    {
        _closed.Cancel();
        if (_process != null)
        {
            try
            {
                if (!_process.HasExited)
                {
                    _process.Kill(true);
                    await _process.WaitForExitAsync().ConfigureAwait(false);
                }
            }
            catch
            {
            }
            _process.Dispose();
        }
        if (_readLoop != null)
        {
            try { await _readLoop.ConfigureAwait(false); } catch { }
        }
        _writeGate.Dispose();
        _closed.Dispose();
    }

    private async Task ReadLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested && _process != null)
            {
                var line = await _process.StandardOutput.ReadLineAsync(cancellationToken).ConfigureAwait(false);
                if (line == null)
                {
                    break;
                }

                JsonElement message;
                try
                {
                    using var document = JsonDocument.Parse(line);
                    message = document.RootElement.Clone();
                }
                catch (Exception exception)
                {
                    _log.Error($"Ignored invalid JSON from ACP agent {_agent.Id}: {line}", exception);
                    continue;
                }

                if (message.TryGetProperty("method", out var methodElement))
                {
                    var method = methodElement.GetString() ?? string.Empty;
                    var parameters = message.TryGetProperty("params", out var paramsElement) ? paramsElement.Clone() : EmptyObject();
                    if (message.TryGetProperty("id", out var requestId))
                    {
                        _ = HandleIncomingRequestAsync(requestId.Clone(), method, parameters, cancellationToken);
                    }
                    else if (NotificationReceived != null)
                    {
                        _ = NotificationReceived(method, parameters, cancellationToken);
                    }
                    continue;
                }

                if (!message.TryGetProperty("id", out var responseId))
                {
                    continue;
                }
                var key = IdKey(responseId);
                if (!_pending.TryGetValue(key, out var completion))
                {
                    continue;
                }
                if (message.TryGetProperty("error", out var error))
                {
                    var code = error.TryGetProperty("code", out var codeElement) ? codeElement.GetInt32() : -32603;
                    var text = error.TryGetProperty("message", out var messageElement) ? messageElement.GetString() : "ACP request failed.";
                    var data = error.TryGetProperty("data", out var dataElement) ? dataElement.GetRawText() : null;
                    completion.TrySetException(new AcpRpcException(code, text ?? "ACP request failed.", data));
                }
                else
                {
                    completion.TrySetResult(message.TryGetProperty("result", out var result) ? result.Clone() : EmptyObject());
                }
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
        catch (Exception exception)
        {
            _log.Error($"ACP agent '{_agent.Id}' disconnected unexpectedly.", exception);
        }
        finally
        {
            foreach (var completion in _pending.Values)
            {
                completion.TrySetException(new IOException($"ACP agent '{_agent.Name}' disconnected."));
            }
        }
    }

    private async Task HandleIncomingRequestAsync(JsonElement id, string method, JsonElement parameters, CancellationToken cancellationToken)
    {
        try
        {
            if (RequestReceived == null)
            {
                await WriteAsync(new { jsonrpc = "2.0", id = JsonValue(id), error = new { code = -32601, message = $"Unsupported client method '{method}'." } }, cancellationToken).ConfigureAwait(false);
                return;
            }
            var result = await RequestReceived(method, parameters, cancellationToken).ConfigureAwait(false);
            await WriteAsync(new { jsonrpc = "2.0", id = JsonValue(id), result = result ?? new { } }, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception exception)
        {
            await WriteAsync(new { jsonrpc = "2.0", id = JsonValue(id), error = new { code = -32603, message = exception.Message } }, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task WriteAsync(object message, CancellationToken cancellationToken)
    {
        if (_process == null || _process.HasExited)
        {
            throw new IOException($"ACP agent '{_agent.Name}' is not running.");
        }
        var json = JsonSerializer.Serialize(message, JsonCodec.Options);
        await _writeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await _process.StandardInput.WriteLineAsync(json.AsMemory(), cancellationToken).ConfigureAwait(false);
            await _process.StandardInput.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _writeGate.Release();
        }
    }

    private async Task PumpStandardErrorAsync(Process process, CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                var line = await process.StandardError.ReadLineAsync(cancellationToken).ConfigureAwait(false);
                if (line == null) break;
                _log.Info($"[{_agent.Id}] {line}");
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
    }

    private static string IdKey(JsonElement id)
    {
        return id.ValueKind == JsonValueKind.String ? id.GetString() ?? string.Empty : id.GetRawText();
    }

    private static object JsonValue(JsonElement element)
    {
        return element.ValueKind == JsonValueKind.String ? element.GetString() ?? string.Empty : element.GetInt64();
    }

    private static JsonElement EmptyObject()
    {
        using var document = JsonDocument.Parse("{}");
        return document.RootElement.Clone();
    }
}
