using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class BrokerClient : IDisposable
    {
        private readonly IOfficeHost _host;
        private readonly SynchronizationContext _ui;
        private readonly ConcurrentDictionary<string, TaskCompletionSource<RpcEnvelope>> _pending =
            new ConcurrentDictionary<string, TaskCompletionSource<RpcEnvelope>>(StringComparer.Ordinal);
        private readonly SemaphoreSlim _writeGate = new SemaphoreSlim(1, 1);
        private readonly SemaphoreSlim _permissionDialogGate = new SemaphoreSlim(1, 1);
        private readonly object _permissionGate = new object();
        private readonly HashSet<string> _rememberedPermissions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly CancellationTokenSource _closed = new CancellationTokenSource();
        private NamedPipeClientStream _pipe;
        private StreamReader _reader;
        private StreamWriter _writer;
        private Task _readLoop;
        private string _activeSessionId;
        private int _disposed;

        public BrokerClient(IOfficeHost host, SynchronizationContext ui)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _ui = ui ?? throw new ArgumentNullException(nameof(ui));
        }

        public event EventHandler<SessionUpdateMessage> SessionUpdate;

        public void SetActiveSession(string sessionId)
        {
            lock (_permissionGate)
            {
                if (string.Equals(_activeSessionId, sessionId, StringComparison.Ordinal)) return;
                _activeSessionId = sessionId ?? string.Empty;
                _rememberedPermissions.Clear();
            }
        }

        public async Task ConnectAsync(CancellationToken cancellationToken)
        {
            Exception lastError = null;
            for (var attempt = 0; attempt < 12; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var pipe = new NamedPipeClientStream(".", RibbonProtocol.PipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                try
                {
                    await pipe.ConnectAsync(attempt == 0 ? 250 : 1000, cancellationToken).ConfigureAwait(false);
                    _pipe = pipe;
                    break;
                }
                catch (Exception exception) when (!(exception is OperationCanceledException))
                {
                    lastError = exception;
                    pipe.Dispose();
                    if (attempt == 0)
                    {
                        BrokerLocator.StartBroker();
                    }
                    await Task.Delay(250, cancellationToken).ConfigureAwait(false);
                }
            }

            if (_pipe == null)
            {
                throw new IOException("Unable to connect to Ribbon Broker.", lastError);
            }

            _reader = new StreamReader(_pipe, new UTF8Encoding(false), false, 8192, true);
            _writer = new StreamWriter(_pipe, new UTF8Encoding(false), 8192, true) { AutoFlush = true };
            _readLoop = ReadLoopAsync(_closed.Token);
            await RequestAsync(RibbonProtocol.RegisterHost, JsonCodec.Serialize(_host.Registration), cancellationToken).ConfigureAwait(false);
        }

        public async Task<T> RequestAsync<T>(string method, object payload, CancellationToken cancellationToken)
        {
            var response = await RequestAsync(method, JsonCodec.Serialize(payload), cancellationToken).ConfigureAwait(false);
            return JsonCodec.Deserialize<T>(response.Payload);
        }

        public async Task RequestAsync(string method, object payload, CancellationToken cancellationToken)
        {
            await RequestAsync(method, JsonCodec.Serialize(payload), cancellationToken).ConfigureAwait(false);
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            _closed.Cancel();
            var ownsWriteGate = false;
            try { ownsWriteGate = _writeGate.Wait(1000); } catch { }
            try { _pipe?.Dispose(); } catch { }
            finally
            {
                if (ownsWriteGate) _writeGate.Release();
            }
            try { _readLoop?.GetAwaiter().GetResult(); } catch { }
            // Pending request continuations can still exit WriteAsync during Office shutdown.
            // The pipe owns the OS handle; do not dispose wrappers or synchronization
            // primitives out from under those continuations.
        }

        private async Task<RpcEnvelope> RequestAsync(string method, string payload, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var request = RpcEnvelope.Request(method, payload);
            var completion = new TaskCompletionSource<RpcEnvelope>(TaskCreationOptions.RunContinuationsAsynchronously);
            if (!_pending.TryAdd(request.Id, completion))
            {
                throw new InvalidOperationException("Unable to allocate a Ribbon broker request id.");
            }

            using (cancellationToken.Register(() => completion.TrySetCanceled()))
            {
                try
                {
                    await WriteAsync(request, cancellationToken).ConfigureAwait(false);
                    var response = await completion.Task.ConfigureAwait(false);
                    if (!response.Success)
                    {
                        throw new InvalidOperationException(response.Error ?? "Ribbon Broker request failed.");
                    }
                    return response;
                }
                finally
                {
                    _pending.TryRemove(request.Id, out _);
                }
            }
        }

        private async Task ReadLoopAsync(CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    var line = await _reader.ReadLineAsync().ConfigureAwait(false);
                    if (line == null) break;
                    var envelope = JsonCodec.Deserialize<RpcEnvelope>(line);
                    if (envelope.Version != RibbonProtocol.Version) continue;
                    if (string.Equals(envelope.Kind, "response", StringComparison.Ordinal))
                    {
                        if (_pending.TryGetValue(envelope.Id ?? string.Empty, out var completion))
                        {
                            completion.TrySetResult(envelope);
                        }
                    }
                    else
                    {
                        _ = ProcessIncomingAsync(envelope, cancellationToken);
                    }
                }
            }
            catch (Exception exception) when (cancellationToken.IsCancellationRequested || exception is IOException || exception is ObjectDisposedException)
            {
            }
            finally
            {
                foreach (var completion in _pending.Values)
                {
                    completion.TrySetException(new IOException("The Ribbon Broker connection closed."));
                }
            }
        }

        private async Task ProcessIncomingAsync(RpcEnvelope envelope, CancellationToken cancellationToken)
        {
            RpcEnvelope response = null;
            try
            {
                switch (envelope.Method)
                {
                    case RibbonProtocol.SessionUpdate:
                        SessionUpdate?.Invoke(this, JsonCodec.Deserialize<SessionUpdateMessage>(envelope.Payload));
                        break;
                    case RibbonProtocol.ListTools:
                        response = RpcEnvelope.Response(envelope, JsonCodec.Serialize(_host.GetTools()));
                        break;
                    case RibbonProtocol.InvokeTool:
                        var invocation = JsonCodec.Deserialize<OfficeToolInvocation>(envelope.Payload);
                        OfficeToolResult result;
                        if (!await AuthorizeDestructiveToolAsync(invocation).ConfigureAwait(false))
                        {
                            result = new OfficeToolResult
                            {
                                Success = false,
                                Error = "The user did not allow this document-changing Office action."
                            };
                        }
                        else
                        {
                            result = await _host.InvokeAsync(invocation, cancellationToken).ConfigureAwait(false);
                        }
                        response = RpcEnvelope.Response(envelope, JsonCodec.Serialize(result));
                        break;
                    case RibbonProtocol.PermissionRequest:
                        var prompt = JsonCodec.Deserialize<PermissionPrompt>(envelope.Payload);
                        var decision = await RequestPermissionAsync(prompt).ConfigureAwait(false);
                        response = RpcEnvelope.Response(envelope, JsonCodec.Serialize(decision));
                        break;
                    default:
                        response = RpcEnvelope.Failure(envelope, "Unknown VSTO host method '" + envelope.Method + "'.");
                        break;
                }
            }
            catch (Exception exception)
            {
                response = RpcEnvelope.Failure(envelope, exception.GetBaseException().Message);
            }

            if (string.Equals(envelope.Kind, "request", StringComparison.Ordinal) && response != null)
            {
                try { await WriteAsync(response, cancellationToken).ConfigureAwait(false); } catch { }
            }
        }

        private async Task<PermissionDecision> RequestPermissionAsync(PermissionPrompt prompt)
        {
            var options = prompt.Options ?? new List<PermissionChoice>();
            var allow = options.FirstOrDefault(option =>
                option.Kind != null && option.Kind.StartsWith("allow", StringComparison.OrdinalIgnoreCase));
            var permissionKey = PermissionKey("acp", prompt.Title);
            if (allow != null && IsRemembered(permissionKey))
            {
                return new PermissionDecision { OptionId = allow.OptionId, Cancelled = false, RememberForSession = true };
            }

            await _permissionDialogGate.WaitAsync(_closed.Token).ConfigureAwait(false);
            try
            {
                var completion = new TaskCompletionSource<PermissionDecision>(TaskCreationOptions.RunContinuationsAsynchronously);
                _ui.Post(_ =>
                {
                    try
                    {
                        var decision = PermissionDialog.ShowAcp(Form.ActiveForm, prompt, RibbonPalette.Detect());
                        if (!decision.Cancelled && decision.RememberForSession) Remember(permissionKey);
                        completion.TrySetResult(decision);
                    }
                    catch (Exception exception)
                    {
                        completion.TrySetException(exception);
                    }
                }, null);
                using (_closed.Token.Register(() => completion.TrySetCanceled()))
                {
                    return await completion.Task.ConfigureAwait(false);
                }
            }
            finally
            {
                _permissionDialogGate.Release();
            }
        }

        private async Task<bool> AuthorizeDestructiveToolAsync(OfficeToolInvocation invocation)
        {
            var definition = _host.GetTools().FirstOrDefault(tool =>
                string.Equals(tool.Name, invocation.ToolName, StringComparison.OrdinalIgnoreCase));
            if (definition == null || !definition.Destructive) return true;

            var permissionKey = PermissionKey("office", invocation.ToolName);
            if (IsRemembered(permissionKey)) return true;

            await _permissionDialogGate.WaitAsync(_closed.Token).ConfigureAwait(false);
            try
            {
                var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                _ui.Post(_ =>
                {
                    try
                    {
                        var decision = PermissionDialog.ShowDestructiveTool(
                            Form.ActiveForm,
                            invocation.ToolName,
                            invocation.ArgumentsJson,
                            RibbonPalette.Detect());
                        if (!decision.Cancelled && decision.RememberForSession) Remember(permissionKey);
                        completion.TrySetResult(!decision.Cancelled);
                    }
                    catch (Exception exception)
                    {
                        completion.TrySetException(exception);
                    }
                }, null);
                using (_closed.Token.Register(() => completion.TrySetCanceled()))
                {
                    return await completion.Task.ConfigureAwait(false);
                }
            }
            finally
            {
                _permissionDialogGate.Release();
            }
        }

        private string PermissionKey(string category, string action)
        {
            lock (_permissionGate)
            {
                return (_activeSessionId ?? string.Empty) + "|" + category + "|" + (action ?? string.Empty);
            }
        }

        private bool IsRemembered(string key)
        {
            lock (_permissionGate) return _rememberedPermissions.Contains(key);
        }

        private void Remember(string key)
        {
            lock (_permissionGate) _rememberedPermissions.Add(key);
        }

        private async Task WriteAsync(RpcEnvelope envelope, CancellationToken cancellationToken)
        {
            await _writeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                await _writer.WriteLineAsync(JsonCodec.Serialize(envelope)).ConfigureAwait(false);
                await _writer.FlushAsync().ConfigureAwait(false);
            }
            finally
            {
                _writeGate.Release();
            }
        }
    }
}
