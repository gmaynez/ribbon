using Ribbon.Broker.Infrastructure;
using Ribbon.Broker.Mcp;
using Ribbon.Broker.Server;
using Ribbon.Contracts;

namespace Ribbon.Broker;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        var paths = new BrokerPaths();
        var log = new BrokerLog(paths);

        try
        {
            if (args.Any(value => string.Equals(value, "--mcp-stdio", StringComparison.OrdinalIgnoreCase)))
            {
                var hostId = ReadArgument(args, "--host-id")
                    ?? throw new ArgumentException("--mcp-stdio requires --host-id.");
                return await new OfficeMcpStdioProxy(hostId, log).RunAsync(CancellationToken.None).ConfigureAwait(false);
            }

            using var mutex = new Mutex(true, RibbonProtocol.BrokerMutexName, out var ownsMutex);
            if (!ownsMutex)
            {
                return 0;
            }

            using var shutdown = new CancellationTokenSource();
            Console.CancelKeyPress += (_, eventArgs) =>
            {
                eventArgs.Cancel = true;
                shutdown.Cancel();
            };

            var server = new BrokerServer(paths, log);
            await server.RunAsync(shutdown.Token).ConfigureAwait(false);
            return 0;
        }
        catch (Exception exception)
        {
            log.Error("Ribbon Broker terminated unexpectedly.", exception);
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static string? ReadArgument(IReadOnlyList<string> args, string name)
    {
        for (var index = 0; index < args.Count - 1; index++)
        {
            if (string.Equals(args[index], name, StringComparison.OrdinalIgnoreCase))
            {
                return args[index + 1];
            }
        }

        return null;
    }
}
