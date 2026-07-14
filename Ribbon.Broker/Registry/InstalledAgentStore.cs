using Ribbon.Broker.Infrastructure;

namespace Ribbon.Broker.Registry;

internal sealed class InstalledAgentStore
{
    private readonly BrokerPaths _paths;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public InstalledAgentStore(BrokerPaths paths)
    {
        _paths = paths;
    }

    public async Task<IReadOnlyList<InstalledAgentRecord>> ListAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await ReadUnsafeAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<InstalledAgentRecord?> FindAsync(string id, CancellationToken cancellationToken)
    {
        return (await ListAsync(cancellationToken).ConfigureAwait(false))
            .FirstOrDefault(record => string.Equals(record.Id, id, StringComparison.OrdinalIgnoreCase));
    }

    public async Task UpsertAsync(InstalledAgentRecord record, CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var records = (await ReadUnsafeAsync(cancellationToken).ConfigureAwait(false)).ToList();
            records.RemoveAll(item => string.Equals(item.Id, record.Id, StringComparison.OrdinalIgnoreCase));
            records.Add(record);
            records.Sort((left, right) => string.Compare(left.Name, right.Name, StringComparison.OrdinalIgnoreCase));
            await WriteUnsafeAsync(records, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task RemoveAsync(string id, CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var records = (await ReadUnsafeAsync(cancellationToken).ConfigureAwait(false)).ToList();
            records.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
            await WriteUnsafeAsync(records, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task<List<InstalledAgentRecord>> ReadUnsafeAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_paths.InstalledAgentsFile))
        {
            return new List<InstalledAgentRecord>();
        }

        var json = await File.ReadAllTextAsync(_paths.InstalledAgentsFile, cancellationToken).ConfigureAwait(false);
        return JsonCodec.Deserialize<List<InstalledAgentRecord>>(json);
    }

    private async Task WriteUnsafeAsync(IReadOnlyList<InstalledAgentRecord> records, CancellationToken cancellationToken)
    {
        var temporary = _paths.InstalledAgentsFile + ".tmp";
        await File.WriteAllTextAsync(temporary, JsonCodec.Serialize(records), cancellationToken).ConfigureAwait(false);
        File.Move(temporary, _paths.InstalledAgentsFile, true);
    }
}
