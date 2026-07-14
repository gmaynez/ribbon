using Ribbon.Broker.Infrastructure;

namespace Ribbon.Broker.Registry;

internal sealed class AgentRegistryClient : IDisposable
{
    private static readonly Uri RegistryUri = new("https://cdn.agentclientprotocol.com/registry/v1/latest/registry.json");
    private readonly BrokerPaths _paths;
    private readonly BrokerLog _log;
    private readonly HttpClient _httpClient;

    public AgentRegistryClient(BrokerPaths paths, BrokerLog log)
    {
        _paths = paths;
        _log = log;
        _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Ribbon-Broker/0.1");
    }

    public async Task<RegistryDocument> GetAsync(bool refresh, CancellationToken cancellationToken)
    {
        if (!refresh && File.Exists(_paths.RegistryCacheFile))
        {
            return ReadCache();
        }

        try
        {
            var json = await _httpClient.GetStringAsync(RegistryUri, cancellationToken).ConfigureAwait(false);
            var document = JsonCodec.Deserialize<RegistryDocument>(json);
            Validate(document);
            await File.WriteAllTextAsync(_paths.RegistryCacheFile, json, cancellationToken).ConfigureAwait(false);
            return document;
        }
        catch (Exception exception) when (File.Exists(_paths.RegistryCacheFile) && exception is not OperationCanceledException)
        {
            _log.Error("Unable to refresh the ACP Registry; using the cached copy.", exception);
            return ReadCache();
        }
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }

    private RegistryDocument ReadCache()
    {
        var document = JsonCodec.Deserialize<RegistryDocument>(File.ReadAllText(_paths.RegistryCacheFile));
        Validate(document);
        return document;
    }

    private static void Validate(RegistryDocument document)
    {
        if (document.Agents == null || document.Agents.Count == 0)
        {
            throw new InvalidDataException("The ACP Registry did not contain any agents.");
        }

        foreach (var agent in document.Agents)
        {
            if (string.IsNullOrWhiteSpace(agent.Id) || string.IsNullOrWhiteSpace(agent.Name) || string.IsNullOrWhiteSpace(agent.Version))
            {
                throw new InvalidDataException("The ACP Registry contained an invalid agent entry.");
            }
        }
    }
}
