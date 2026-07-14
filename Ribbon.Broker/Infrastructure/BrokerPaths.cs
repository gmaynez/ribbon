namespace Ribbon.Broker.Infrastructure;

internal sealed class BrokerPaths
{
    public BrokerPaths()
    {
        Root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Ribbon");
        Agents = Path.Combine(Root, "agents");
        Runtimes = Path.Combine(Root, "runtimes");
        Sessions = Path.Combine(Root, "sessions");
        Cache = Path.Combine(Root, "cache");
        Logs = Path.Combine(Root, "logs");

        Directory.CreateDirectory(Root);
        Directory.CreateDirectory(Agents);
        Directory.CreateDirectory(Runtimes);
        Directory.CreateDirectory(Sessions);
        Directory.CreateDirectory(Cache);
        Directory.CreateDirectory(Logs);
    }

    public string Root { get; }
    public string Agents { get; }
    public string Runtimes { get; }
    public string Sessions { get; }
    public string Cache { get; }
    public string Logs { get; }
    public string RegistryCacheFile => Path.Combine(Cache, "registry.json");
    public string InstalledAgentsFile => Path.Combine(Root, "installed-agents.json");
}
