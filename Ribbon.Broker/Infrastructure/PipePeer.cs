using System.Collections.Concurrent;
using System.Text;
using Ribbon.Contracts;

namespace Ribbon.Broker.Infrastructure;

internal sealed class PipePeer : IAsyncDisposable
{
    private readonly Stream _stream;
    private readonly StreamReader _reader;
    private readonly StreamWriter _writer;
    private readonly BrokerLog _log;
    private readonly SemaphoreSlim _writeGate = new(1, 1);
    private readonly ConcurrentDictionary<string, TaskCompletionSource<RpcEnvelope>> _pending = new(StringComparer.Ordinal);
    private readonly CancellationTokenSource _closed = new();
    private Task? _readLoop;

    public PipePeer(Stream stream, BrokerLog log)
    {
        _stream = stream;
        _log = log;
        _reader = new StreamReader(stream, new UTF8Encoding(false), false, 8192, true);
        _writer = new StreamWriter(stream, new UTF8Encoding(false), 8192, true) { AutoFlush = true };
    }

    public bool IsConnected => !_closed.IsCancellationRequested;
    public event Action<PipePeer>? Closed;

    public void Start(Func<PipePeer, RpcEnvelope, CancellationToken, Task<RpcEnvelope?>> handler, CancellationToken cancellationToken)
    {
        if (_readLoop != null)
        {
            throw new InvalidOperationException("The pipe peer has already started.");
        }

        _readLoop = ReadLoopAsync(handler, cancellationToken);
    }

    public async Task<RpcEnvelope> RequestAsync(string method, string payload, CancellationToken cancellationToken)
    {
        var request = RpcEnvelope.Request(method, payload);
        var completion = new TaskCompletionSource<RpcEnvelope>(TaskCreationOptions.RunContinuationsAsynchronously);
        if (!_pending.TryAdd(request.Id, completion))
        {
            throw new InvalidOperationException("Unable to allocate a broker request id.");
        }

        try
        {
            await WriteAsync(request, cancellationToken).ConfigureAwait(false);
            var response = await completion.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
            if (!response.Success)
            {
                throw new InvalidOperationException(response.Error ?? "The broker request failed.");
            }
            return response;
        }
        finally
        {
            _pending.TryRemove(request.Id, out _);
        }
    }

    public Task NotifyAsync(string method, string payload, CancellationToken cancellationToken)
    {
        return WriteAsync(RpcEnvelope.Notification(method, payload), cancellationToken);
    }

    public async ValueTask DisposeAsync()
    {
        _closed.Cancel();
        try { _stream.Dispose(); } catch { }
        if (_readLoop != null)
        {
            try { await _readLoop.ConfigureAwait(false); } catch { }
        }
        try { _reader.Dispose(); } catch { }
        try { _writer.Dispose(); } catch { }
        try { _writeGate.Dispose(); } catch { }
        try { _closed.Dispose(); } catch { }
    }

    private async Task ReadLoopAsync(Func<PipePeer, RpcEnvelope, CancellationToken, Task<RpcEnvelope?>> handler, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _closed.Token);
        try
        {
            while (!linked.IsCancellationRequested)
            {
                var line = await _reader.ReadLineAsync(linked.Token).ConfigureAwait(false);
                if (line == null)
                {
                    break;
                }

                RpcEnvelope envelope;
                try
                {
                    envelope = JsonCodec.Deserialize<RpcEnvelope>(line);
                }
                catch (Exception exception)
                {
                    _log.Error("Ignored an invalid broker pipe message.", exception);
                    continue;
                }

                if (envelope.Version != RibbonProtocol.Version)
                {
                    _log.Error($"Ignored broker protocol version {envelope.Version}.");
                    continue;
                }

                if (string.Equals(envelope.Kind, "response", StringComparison.Ordinal))
                {
                    if (!string.IsNullOrWhiteSpace(envelope.Id) && _pending.TryGetValue(envelope.Id, out var completion))
                    {
                        completion.TrySetResult(envelope);
                    }
                    continue;
                }

                _ = ProcessIncomingAsync(envelope, handler, linked.Token);
            }
        }
        catch (OperationCanceledException) when (linked.IsCancellationRequested)
        {
        }
        catch (Exception exception)
        {
            _log.Error("A broker pipe connection closed unexpectedly.", exception);
        }
        finally
        {
            _closed.Cancel();
            foreach (var completion in _pending.Values)
            {
                completion.TrySetException(new IOException("The broker pipe connection closed."));
            }
            Closed?.Invoke(this);
        }
    }

    private async Task ProcessIncomingAsync(RpcEnvelope envelope, Func<PipePeer, RpcEnvelope, CancellationToken, Task<RpcEnvelope?>> handler, CancellationToken cancellationToken)
    {
        try
        {
            var response = await handler(this, envelope, cancellationToken).ConfigureAwait(false);
            if (string.Equals(envelope.Kind, "request", StringComparison.Ordinal) && response != null)
            {
                await WriteAsync(response, cancellationToken).ConfigureAwait(false);
            }
        }
        catch (Exception exception)
        {
            if (string.Equals(envelope.Kind, "request", StringComparison.Ordinal))
            {
                try { await WriteAsync(RpcEnvelope.Failure(envelope, exception.Message), cancellationToken).ConfigureAwait(false); } catch { }
            }
            _log.Error($"Broker request '{envelope.Method}' failed.", exception);
        }
    }

    private async Task WriteAsync(RpcEnvelope envelope, CancellationToken cancellationToken)
    {
        var line = JsonCodec.Serialize(envelope);
        await _writeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await _writer.WriteLineAsync(line.AsMemory(), cancellationToken).ConfigureAwait(false);
            await _writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _writeGate.Release();
        }
    }
}
