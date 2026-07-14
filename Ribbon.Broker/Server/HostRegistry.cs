using System.Collections.Concurrent;
using Ribbon.Broker.Infrastructure;
using Ribbon.Contracts;

namespace Ribbon.Broker.Server;

internal sealed class HostConnection
{
    public HostConnection(HostRegistration registration, PipePeer peer)
    {
        Registration = registration;
        Peer = peer;
    }

    public HostRegistration Registration { get; }
    public PipePeer Peer { get; }
}

internal sealed class HostRegistry
{
    private readonly ConcurrentDictionary<string, HostConnection> _hosts = new(StringComparer.OrdinalIgnoreCase);

    public int Count => _hosts.Count;

    public void Register(HostRegistration registration, PipePeer peer)
    {
        if (string.IsNullOrWhiteSpace(registration.HostId) || string.IsNullOrWhiteSpace(registration.HostKind))
        {
            throw new ArgumentException("A host id and host kind are required.");
        }
        _hosts[registration.HostId] = new HostConnection(registration, peer);
    }

    public HostConnection Get(string hostId)
    {
        if (!_hosts.TryGetValue(hostId, out var host) || !host.Peer.IsConnected)
        {
            throw new InvalidOperationException("The Office host is no longer connected to Ribbon.");
        }
        return host;
    }

    public IReadOnlyList<HostConnection> List(string preferredHostId)
    {
        return _hosts.Values
            .Where(host => host.Peer.IsConnected)
            .OrderByDescending(host => string.Equals(host.Registration.HostId, preferredHostId, StringComparison.OrdinalIgnoreCase))
            .ThenBy(host => host.Registration.HostKind, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public void Remove(PipePeer peer)
    {
        foreach (var pair in _hosts.Where(pair => ReferenceEquals(pair.Value.Peer, peer)).ToArray())
        {
            _hosts.TryRemove(pair.Key, out _);
        }
    }
}
