namespace Ribbon.Broker.Registry;

internal sealed class RegistryDocument
{
    public string Version { get; set; } = string.Empty;
    public List<RegistryAgent> Agents { get; set; } = new();
}

internal sealed class RegistryAgent
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Repository { get; set; } = string.Empty;
    public string Website { get; set; } = string.Empty;
    public List<string> Authors { get; set; } = new();
    public string License { get; set; } = string.Empty;
    public string Icon { get; set; } = string.Empty;
    public RegistryDistribution Distribution { get; set; } = new();
}

internal sealed class RegistryDistribution
{
    public Dictionary<string, RegistryBinaryTarget>? Binary { get; set; }
    public RegistryPackageTarget? Npx { get; set; }
    public RegistryPackageTarget? Uvx { get; set; }
}

internal sealed class RegistryBinaryTarget
{
    public string Archive { get; set; } = string.Empty;
    public string Cmd { get; set; } = string.Empty;
    public List<string> Args { get; set; } = new();
    public Dictionary<string, string> Env { get; set; } = new();
}

internal sealed class RegistryPackageTarget
{
    public string Package { get; set; } = string.Empty;
    public List<string> Args { get; set; } = new();
    public Dictionary<string, string> Env { get; set; } = new();
}

internal sealed class InstalledAgentRecord
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Command { get; set; } = string.Empty;
    public List<string> Arguments { get; set; } = new();
    public Dictionary<string, string> Environment { get; set; } = new();
    public string License { get; set; } = string.Empty;
    public string Website { get; set; } = string.Empty;
    public string DistributionType { get; set; } = string.Empty;
    public DateTimeOffset InstalledAt { get; set; }
}
