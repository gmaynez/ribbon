using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Grid.Constants;
using Grid.Tools;

namespace Grid.Mcp
{
    internal sealed class GridMcpSession : IAsyncDisposable
    {
        private readonly CancellationTokenSource _cancellationTokenSource;
        private readonly StreamableHttpServerTransport _transport;
        private readonly McpServer _server;
        private readonly Task _serverTask;

        private GridMcpSession(string sessionId, StreamableHttpServerTransport transport, McpServer server, CancellationTokenSource cancellationTokenSource)
        {
            SessionId = sessionId;
            _transport = transport;
            _server = server;
            _cancellationTokenSource = cancellationTokenSource;
            _serverTask = _server.RunAsync(_cancellationTokenSource.Token);
        }

        public string SessionId { get; private set; }

        public static GridMcpSession Create(GridToolCatalog toolCatalog, string sessionId)
        {
            StreamableHttpServerTransport transport;
            McpServerOptions options;
            CancellationTokenSource cancellationTokenSource;
            McpServer server;

            cancellationTokenSource = new CancellationTokenSource();
            transport = new StreamableHttpServerTransport(NullLoggerFactory.Instance)
            {
                SessionId = sessionId
            };
            options = CreateServerOptions(toolCatalog);
            server = McpServer.Create(transport, options, NullLoggerFactory.Instance, null);
            return new GridMcpSession(sessionId, transport, server, cancellationTokenSource);
        }

        public Task<bool> HandlePostAsync(JsonRpcMessage message, Stream responseStream, CancellationToken cancellationToken)
        {
            return _transport.HandlePostRequestAsync(message, responseStream, cancellationToken);
        }

        public Task HandleGetAsync(Stream responseStream, CancellationToken cancellationToken)
        {
            return _transport.HandleGetRequestAsync(responseStream, cancellationToken);
        }

        public async ValueTask DisposeAsync()
        {
            try
            {
                _cancellationTokenSource.Cancel();
            }
            catch
            {
            }

            try
            {
                await _transport.DisposeAsync().ConfigureAwait(false);
            }
            catch
            {
            }

            try
            {
                await _server.DisposeAsync().ConfigureAwait(false);
            }
            catch
            {
            }

            try
            {
                await _serverTask.ConfigureAwait(false);
            }
            catch
            {
            }

            _cancellationTokenSource.Dispose();
        }

        private static McpServerOptions CreateServerOptions(GridToolCatalog toolCatalog)
        {
            McpServerOptions options;

            options = new McpServerOptions();
            options.ServerInfo = new Implementation
            {
                Name = GridConstants.ServerName,
                Version = GridConstants.ExtensionVersion,
                Title = "Grid"
            };
            options.ServerInstructions = "Use the tools to automate Excel, Word, and PowerPoint. Prefer small, verifiable actions and inspect context before making changes.";
            options.Capabilities = new ServerCapabilities
            {
                Tools = new ToolsCapability
                {
                    ListChanged = false
                }
            };
            options.ToolCollection = toolCatalog.CreateMcpToolCollection();
            return options;
        }
    }
}
