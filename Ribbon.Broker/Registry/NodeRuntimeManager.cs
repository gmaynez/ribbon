using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using Ribbon.Broker.Infrastructure;

namespace Ribbon.Broker.Registry;

internal sealed class NodeRuntimeManager
{
    private static readonly Uri IndexUri = new("https://nodejs.org/dist/index.json");
    private readonly BrokerPaths _paths;
    private readonly BrokerLog _log;
    private readonly HttpClient _httpClient;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public NodeRuntimeManager(BrokerPaths paths, BrokerLog log, HttpClient httpClient)
    {
        _paths = paths;
        _log = log;
        _httpClient = httpClient;
    }

    public async Task<string> EnsureNpxAsync(CancellationToken cancellationToken)
    {
        var existing = FindOnPath("npx.cmd");
        if (!string.IsNullOrWhiteSpace(existing))
        {
            return existing;
        }

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            existing = Directory.Exists(Path.Combine(_paths.Runtimes, "node"))
                ? Directory.EnumerateFiles(Path.Combine(_paths.Runtimes, "node"), "npx.cmd", SearchOption.AllDirectories).FirstOrDefault()
                : null;
            if (!string.IsNullOrWhiteSpace(existing))
            {
                return existing;
            }

            return await InstallLatestLtsAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task<string> InstallLatestLtsAsync(CancellationToken cancellationToken)
    {
        var indexJson = await _httpClient.GetStringAsync(IndexUri, cancellationToken).ConfigureAwait(false);
        using var index = JsonDocument.Parse(indexJson);
        var release = index.RootElement.EnumerateArray().FirstOrDefault(item =>
            item.TryGetProperty("lts", out var lts) && lts.ValueKind == JsonValueKind.String);
        if (release.ValueKind == JsonValueKind.Undefined)
        {
            throw new InvalidOperationException("Node.js did not publish an LTS runtime in its release index.");
        }

        var version = release.GetProperty("version").GetString()
            ?? throw new InvalidDataException("The Node.js LTS entry did not contain a version.");
        var architecture = RuntimeInformation.OSArchitecture == Architecture.Arm64 ? "arm64" : "x64";
        var fileName = $"node-{version}-win-{architecture}.zip";
        var releaseBase = new Uri($"https://nodejs.org/dist/{version}/");
        var checksums = await _httpClient.GetStringAsync(new Uri(releaseBase, "SHASUMS256.txt"), cancellationToken).ConfigureAwait(false);
        var expectedHash = checksums.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            .Where(parts => parts.Length >= 2 && string.Equals(parts[^1], fileName, StringComparison.Ordinal))
            .Select(parts => parts[0])
            .FirstOrDefault();
        if (string.IsNullOrWhiteSpace(expectedHash))
        {
            throw new InvalidDataException($"Node.js did not publish a checksum for {fileName}.");
        }

        var temporaryArchive = Path.Combine(_paths.Cache, fileName + ".download");
        await DownloadAsync(new Uri(releaseBase, fileName), temporaryArchive, cancellationToken).ConfigureAwait(false);
        string actualHash;
        await using (var archiveStream = File.OpenRead(temporaryArchive))
        {
            actualHash = Convert.ToHexString(await SHA256.HashDataAsync(archiveStream, cancellationToken).ConfigureAwait(false));
        }
        if (!string.Equals(expectedHash, actualHash, StringComparison.OrdinalIgnoreCase))
        {
            File.Delete(temporaryArchive);
            throw new InvalidDataException("The downloaded Node.js runtime failed checksum verification.");
        }

        var nodeRoot = Path.Combine(_paths.Runtimes, "node", version);
        var temporaryDirectory = nodeRoot + ".installing-" + Guid.NewGuid().ToString("N");
        ArchiveUtilities.ExtractZipSafely(temporaryArchive, temporaryDirectory);
        File.Delete(temporaryArchive);

        var npx = Directory.EnumerateFiles(temporaryDirectory, "npx.cmd", SearchOption.AllDirectories).FirstOrDefault()
            ?? throw new InvalidDataException("The Node.js archive did not contain npx.cmd.");
        Directory.CreateDirectory(Path.GetDirectoryName(nodeRoot)!);
        if (Directory.Exists(nodeRoot))
        {
            Directory.Delete(nodeRoot, true);
        }
        Directory.Move(temporaryDirectory, nodeRoot);
        npx = Directory.EnumerateFiles(nodeRoot, "npx.cmd", SearchOption.AllDirectories).First();
        _log.Info($"Installed managed Node.js runtime {version}.");
        return npx;
    }

    private async Task DownloadAsync(Uri uri, string destination, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        await using var source = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        await using var target = File.Create(destination);
        await source.CopyToAsync(target, cancellationToken).ConfigureAwait(false);
    }

    private static string? FindOnPath(string fileName)
    {
        foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                continue;
            }

            try
            {
                var candidate = Path.Combine(directory.Trim(), fileName);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            catch
            {
            }
        }

        return null;
    }
}
