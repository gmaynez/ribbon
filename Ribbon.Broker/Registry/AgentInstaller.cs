using System.Runtime.InteropServices;
using Ribbon.Broker.Infrastructure;
using Ribbon.Contracts;

namespace Ribbon.Broker.Registry;

internal sealed class AgentInstaller : IDisposable
{
    private readonly BrokerPaths _paths;
    private readonly BrokerLog _log;
    private readonly AgentRegistryClient _registry;
    private readonly InstalledAgentStore _store;
    private readonly HttpClient _httpClient;
    private readonly NodeRuntimeManager _node;

    public AgentInstaller(BrokerPaths paths, BrokerLog log, AgentRegistryClient registry, InstalledAgentStore store)
    {
        _paths = paths;
        _log = log;
        _registry = registry;
        _store = store;
        _httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Ribbon-Broker/0.1");
        _node = new NodeRuntimeManager(paths, log, _httpClient);
    }

    public async Task<IReadOnlyList<AgentSummary>> ListInstalledAsync(CancellationToken cancellationToken)
    {
        var installed = await _store.ListAsync(cancellationToken).ConfigureAwait(false);
        RegistryDocument? registry = null;
        try
        {
            registry = await _registry.GetAsync(false, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            _log.Error("Unable to load the ACP Registry while listing installed agents.", exception);
        }
        return installed.Select(record => ToSummary(record, registry?.Agents.FirstOrDefault(item =>
            string.Equals(item.Id, record.Id, StringComparison.OrdinalIgnoreCase)))).ToList();
    }

    public async Task<IReadOnlyList<AgentSummary>> ListRegistryAsync(bool refresh, CancellationToken cancellationToken)
    {
        var registry = await _registry.GetAsync(refresh, cancellationToken).ConfigureAwait(false);
        var installed = await _store.ListAsync(cancellationToken).ConfigureAwait(false);
        return registry.Agents
            .Where(IsWindowsCompatible)
            .OrderBy(agent => agent.Name, StringComparer.OrdinalIgnoreCase)
            .Select(agent => ToSummary(agent, installed.FirstOrDefault(item =>
                string.Equals(item.Id, agent.Id, StringComparison.OrdinalIgnoreCase))))
            .ToList();
    }

    public async Task<InstalledAgentRecord> InstallAsync(string id, CancellationToken cancellationToken)
    {
        var registry = await _registry.GetAsync(true, cancellationToken).ConfigureAwait(false);
        var agent = registry.Agents.FirstOrDefault(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Agent '{id}' was not found in the ACP Registry.");

        InstalledAgentRecord record;
        var platform = RuntimeInformation.OSArchitecture == Architecture.Arm64 ? "windows-aarch64" : "windows-x86_64";
        if (agent.Distribution.Binary != null && agent.Distribution.Binary.TryGetValue(platform, out var binary))
        {
            record = await InstallBinaryAsync(agent, binary, cancellationToken).ConfigureAwait(false);
        }
        else if (agent.Distribution.Npx != null)
        {
            record = await InstallNpxAsync(agent, agent.Distribution.Npx, cancellationToken).ConfigureAwait(false);
        }
        else if (agent.Distribution.Uvx != null)
        {
            var uvx = FindOnPath("uvx.exe") ?? FindOnPath("uvx");
            if (string.IsNullOrWhiteSpace(uvx))
            {
                throw new InvalidOperationException("This agent requires uvx. Install uv from https://docs.astral.sh/uv/ and retry.");
            }

            record = CreateRecord(agent, uvx, new[] { agent.Distribution.Uvx.Package }.Concat(agent.Distribution.Uvx.Args), agent.Distribution.Uvx.Env, "uvx");
        }
        else
        {
            throw new InvalidOperationException($"Agent '{agent.Name}' has no compatible Windows distribution.");
        }

        await _store.UpsertAsync(record, cancellationToken).ConfigureAwait(false);
        _log.Info($"Installed ACP agent {record.Id} {record.Version} using {record.DistributionType}.");
        return record;
    }

    public async Task UninstallAsync(string id, CancellationToken cancellationToken)
    {
        var root = Path.Combine(_paths.Agents, SanitizeId(id));
        if (Directory.Exists(root))
        {
            Directory.Delete(root, true);
        }
        await _store.RemoveAsync(id, cancellationToken).ConfigureAwait(false);
        _log.Info($"Uninstalled ACP agent {id}.");
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }

    private async Task<InstalledAgentRecord> InstallBinaryAsync(RegistryAgent agent, RegistryBinaryTarget target, CancellationToken cancellationToken)
    {
        var uri = RequireHttps(target.Archive);
        var agentRoot = Path.Combine(_paths.Agents, SanitizeId(agent.Id), agent.Version);
        var temporaryDirectory = agentRoot + ".installing-" + Guid.NewGuid().ToString("N");
        var temporaryArchive = Path.Combine(_paths.Cache, $"{agent.Id}-{agent.Version}-{Guid.NewGuid():N}.download");

        using (var response = await _httpClient.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
        {
            response.EnsureSuccessStatusCode();
            await using var source = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
            await using var targetStream = File.Create(temporaryArchive);
            await source.CopyToAsync(targetStream, cancellationToken).ConfigureAwait(false);
        }

        try
        {
            if (uri.AbsolutePath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
            {
                ArchiveUtilities.ExtractZipSafely(temporaryArchive, temporaryDirectory);
            }
            else
            {
                throw new NotSupportedException("Ribbon currently supports ZIP binary distributions on Windows.");
            }

            var command = ArchiveUtilities.ResolveContainedPath(temporaryDirectory, target.Cmd.TrimStart('.', '/', '\\'));
            if (!File.Exists(command))
            {
                throw new FileNotFoundException("The installed archive did not contain its registry command.", command);
            }

            Directory.CreateDirectory(Path.GetDirectoryName(agentRoot)!);
            if (Directory.Exists(agentRoot))
            {
                Directory.Delete(agentRoot, true);
            }
            Directory.Move(temporaryDirectory, agentRoot);
            command = ArchiveUtilities.ResolveContainedPath(agentRoot, target.Cmd.TrimStart('.', '/', '\\'));
            return CreateRecord(agent, command, target.Args, target.Env, "binary");
        }
        finally
        {
            if (File.Exists(temporaryArchive))
            {
                File.Delete(temporaryArchive);
            }
            if (Directory.Exists(temporaryDirectory))
            {
                Directory.Delete(temporaryDirectory, true);
            }
        }
    }

    private async Task<InstalledAgentRecord> InstallNpxAsync(RegistryAgent agent, RegistryPackageTarget target, CancellationToken cancellationToken)
    {
        var npx = await _node.EnsureNpxAsync(cancellationToken).ConfigureAwait(false);
        var command = npx;
        var arguments = new List<string>();
        var nodeDirectory = Path.GetDirectoryName(npx);
        var node = nodeDirectory == null ? null : Path.Combine(nodeDirectory, "node.exe");
        var npxCli = nodeDirectory == null ? null : Path.Combine(nodeDirectory, "node_modules", "npm", "bin", "npx-cli.js");
        if (node != null && npxCli != null && File.Exists(node) && File.Exists(npxCli))
        {
            command = node;
            arguments.Add(npxCli);
        }
        else if (npx.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Ribbon found npx.cmd but could not locate its node.exe and npx-cli.js runtime files.");
        }

        arguments.Add("-y");
        arguments.Add(target.Package);
        arguments.AddRange(target.Args);
        return CreateRecord(agent, command, arguments, target.Env, "npx");
    }

    private static InstalledAgentRecord CreateRecord(RegistryAgent agent, string command, IEnumerable<string> arguments, IReadOnlyDictionary<string, string> environment, string distributionType)
    {
        return new InstalledAgentRecord
        {
            Id = agent.Id,
            Name = agent.Name,
            Version = agent.Version,
            Description = agent.Description,
            Command = command,
            Arguments = arguments.ToList(),
            Environment = environment.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase),
            License = agent.License,
            Website = string.IsNullOrWhiteSpace(agent.Website) ? agent.Repository : agent.Website,
            DistributionType = distributionType,
            InstalledAt = DateTimeOffset.UtcNow
        };
    }

    private static AgentSummary ToSummary(InstalledAgentRecord installed, RegistryAgent? registry)
    {
        return new AgentSummary
        {
            Id = installed.Id,
            Name = installed.Name,
            Version = installed.Version,
            Description = installed.Description,
            Command = installed.Command,
            Arguments = installed.Arguments,
            Installed = true,
            UpdateAvailable = registry != null && !string.Equals(registry.Version, installed.Version, StringComparison.OrdinalIgnoreCase),
            License = installed.License,
            Website = installed.Website,
            DistributionType = installed.DistributionType
        };
    }

    private static AgentSummary ToSummary(RegistryAgent registry, InstalledAgentRecord? installed)
    {
        return new AgentSummary
        {
            Id = registry.Id,
            Name = registry.Name,
            Version = registry.Version,
            Description = registry.Description,
            Command = installed?.Command ?? string.Empty,
            Arguments = installed?.Arguments ?? new List<string>(),
            Installed = installed != null,
            UpdateAvailable = installed != null && !string.Equals(registry.Version, installed.Version, StringComparison.OrdinalIgnoreCase),
            License = registry.License,
            Website = string.IsNullOrWhiteSpace(registry.Website) ? registry.Repository : registry.Website,
            DistributionType = installed?.DistributionType ?? GetDistributionType(registry)
        };
    }

    private static bool IsWindowsCompatible(RegistryAgent agent)
    {
        var platform = RuntimeInformation.OSArchitecture == Architecture.Arm64 ? "windows-aarch64" : "windows-x86_64";
        return agent.Distribution.Binary?.ContainsKey(platform) == true
            || agent.Distribution.Npx != null
            || agent.Distribution.Uvx != null;
    }

    private static string GetDistributionType(RegistryAgent agent)
    {
        var platform = RuntimeInformation.OSArchitecture == Architecture.Arm64 ? "windows-aarch64" : "windows-x86_64";
        if (agent.Distribution.Binary?.ContainsKey(platform) == true) return "binary";
        if (agent.Distribution.Npx != null) return "npx";
        if (agent.Distribution.Uvx != null) return "uvx";
        return string.Empty;
    }

    private static Uri RequireHttps(string value)
    {
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
        {
            throw new InvalidDataException("Registry downloads must use HTTPS.");
        }
        return uri;
    }

    private static string SanitizeId(string id)
    {
        if (string.IsNullOrWhiteSpace(id) || id.Any(character => !(char.IsLetterOrDigit(character) || character == '-')))
        {
            throw new ArgumentException("The agent id is invalid.", nameof(id));
        }
        return id;
    }

    private static string? FindOnPath(string fileName)
    {
        foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(Path.PathSeparator))
        {
            try
            {
                var candidate = Path.Combine(directory.Trim(), fileName);
                if (File.Exists(candidate)) return candidate;
            }
            catch
            {
            }
        }
        return null;
    }
}
