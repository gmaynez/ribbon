using System.Collections.Concurrent;
using System.IO.Pipes;
using Ribbon.Broker.Acp;
using Ribbon.Broker.Infrastructure;
using Ribbon.Broker.Registry;
using Ribbon.Contracts;

namespace Ribbon.Broker.Server;

internal sealed class BrokerServer
{
    private readonly BrokerLog _log;
    private readonly HostRegistry _hosts;
    private readonly AgentRegistryClient _registry;
    private readonly InstalledAgentStore _installed;
    private readonly AgentInstaller _installer;
    private readonly AgentSessionManager _sessions;
    private readonly ConcurrentDictionary<PipePeer, byte> _peers = new();

    public BrokerServer(BrokerPaths paths, BrokerLog log)
    {
        _log = log;
        _hosts = new HostRegistry();
        _registry = new AgentRegistryClient(paths, log);
        _installed = new InstalledAgentStore(paths);
        _installer = new AgentInstaller(paths, log, _registry, _installed);
        _sessions = new AgentSessionManager(paths, log, _installed, _hosts);
    }

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        _log.Info("Ribbon Broker started.");
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                var pipe = new NamedPipeServerStream(
                    RibbonProtocol.PipeName,
                    PipeDirection.InOut,
                    NamedPipeServerStream.MaxAllowedServerInstances,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
                try
                {
                    await pipe.WaitForConnectionAsync(cancellationToken).ConfigureAwait(false);
                }
                catch
                {
                    pipe.Dispose();
                    throw;
                }

                var peer = new PipePeer(pipe, _log);
                _peers.TryAdd(peer, 0);
                peer.Closed += OnPeerClosed;
                peer.Start(HandleAsync, cancellationToken);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
        finally
        {
            await _sessions.DisposeAsync().ConfigureAwait(false);
            foreach (var peer in _peers.Keys)
            {
                await peer.DisposeAsync().ConfigureAwait(false);
            }
            _installer.Dispose();
            _registry.Dispose();
            _log.Info("Ribbon Broker stopped.");
        }
    }

    private async Task<RpcEnvelope?> HandleAsync(PipePeer peer, RpcEnvelope envelope, CancellationToken cancellationToken)
    {
        switch (envelope.Method)
        {
            case RibbonProtocol.RegisterHost:
                _hosts.Register(JsonCodec.Deserialize<HostRegistration>(envelope.Payload), peer);
                return RpcEnvelope.Response(envelope, "{}");

            case RibbonProtocol.UnregisterHost:
                _hosts.Remove(peer);
                return RpcEnvelope.Response(envelope, "{}");

            case RibbonProtocol.ListTools:
                {
                    var hostId = JsonCodec.Deserialize<HostIdRequest>(envelope.Payload).HostId;
                    var tools = new List<OfficeToolDefinition>();
                    var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var host in _hosts.List(hostId))
                    {
                        var response = await host.Peer.RequestAsync(RibbonProtocol.ListTools, "{}", cancellationToken).ConfigureAwait(false);
                        foreach (var tool in JsonCodec.Deserialize<List<OfficeToolDefinition>>(response.Payload))
                        {
                            if (!names.Add(tool.Name)) continue;
                            tool.HostId = host.Registration.HostId;
                            tool.HostKind = host.Registration.HostKind;
                            tools.Add(tool);
                        }
                    }
                    return RpcEnvelope.Response(envelope, JsonCodec.Serialize(tools));
                }

            case RibbonProtocol.InvokeTool:
                {
                    var invocation = JsonCodec.Deserialize<OfficeToolInvocation>(envelope.Payload);
                    var response = await _hosts.Get(invocation.HostId).Peer.RequestAsync(RibbonProtocol.InvokeTool, envelope.Payload, cancellationToken).ConfigureAwait(false);
                    return RpcEnvelope.Response(envelope, response.Payload);
                }

            case RibbonProtocol.ListInstalledAgents:
                return RpcEnvelope.Response(envelope, JsonCodec.Serialize(await _installer.ListInstalledAsync(cancellationToken).ConfigureAwait(false)));

            case RibbonProtocol.ListRegistryAgents:
                return RpcEnvelope.Response(envelope, JsonCodec.Serialize(await _installer.ListRegistryAsync(true, cancellationToken).ConfigureAwait(false)));

            case RibbonProtocol.InstallAgent:
                {
                    var request = JsonCodec.Deserialize<AgentIdRequest>(envelope.Payload);
                    await _installer.InstallAsync(request.AgentId, cancellationToken).ConfigureAwait(false);
                    return RpcEnvelope.Response(envelope, "{}");
                }

            case RibbonProtocol.UninstallAgent:
                {
                    var request = JsonCodec.Deserialize<AgentIdRequest>(envelope.Payload);
                    await _installer.UninstallAsync(request.AgentId, cancellationToken).ConfigureAwait(false);
                    return RpcEnvelope.Response(envelope, "{}");
                }

            case RibbonProtocol.AuthenticateAgent:
                await _sessions.AuthenticateAsync(JsonCodec.Deserialize<AgentAuthenticationRequest>(envelope.Payload), cancellationToken).ConfigureAwait(false);
                return RpcEnvelope.Response(envelope, "{}");

            case RibbonProtocol.StartSession:
                {
                    var response = await _sessions.StartAsync(JsonCodec.Deserialize<SessionStartRequest>(envelope.Payload), peer, cancellationToken).ConfigureAwait(false);
                    return RpcEnvelope.Response(envelope, JsonCodec.Serialize(response));
                }

            case RibbonProtocol.PromptSession:
                await _sessions.PromptAsync(JsonCodec.Deserialize<SessionPromptRequest>(envelope.Payload), cancellationToken).ConfigureAwait(false);
                return RpcEnvelope.Response(envelope, "{}");

            case RibbonProtocol.CancelSession:
                await _sessions.CancelAsync(JsonCodec.Deserialize<SessionCancelRequest>(envelope.Payload), cancellationToken).ConfigureAwait(false);
                return RpcEnvelope.Response(envelope, "{}");

            default:
                throw new InvalidOperationException($"Unknown Ribbon broker method '{envelope.Method}'.");
        }
    }

    private void OnPeerClosed(PipePeer peer)
    {
        _hosts.Remove(peer);
        _peers.TryRemove(peer, out _);
    }
}
