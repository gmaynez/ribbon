using System;
using System.Threading;
using System.Windows.Forms;
using Grid.Chat;
using Grid.Configuration;
using Grid.Mcp;
using Grid.Office;
using Grid.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace Grid.Runtime
{
    internal sealed class GridRuntime : IDisposable
    {
        private readonly OfficeDispatcher _dispatcher;
        private bool _disposed;

        public GridRuntime(Excel.Application application, SynchronizationContext synchronizationContext)
        {
            if (application == null)
            {
                throw new ArgumentNullException(nameof(application));
            }

            Settings = GridSettings.Default;
            Settings.Normalize();
            _dispatcher = new OfficeDispatcher(synchronizationContext ?? new WindowsFormsSynchronizationContext());
            OfficeAutomation = new OfficeAutomationService(application, _dispatcher);
            ToolHandlers = new GridToolHandlers(OfficeAutomation);
            ToolCatalog = new GridToolCatalog(ToolHandlers);
            ConversationService = new GridConversationService(Settings, ToolCatalog);
            McpServerHost = new GridMcpServerHost(ToolCatalog);
        }

        public GridSettings Settings { get; private set; }

        public OfficeAutomationService OfficeAutomation { get; private set; }

        public GridToolHandlers ToolHandlers { get; private set; }

        public GridToolCatalog ToolCatalog { get; private set; }

        public GridConversationService ConversationService { get; private set; }

        public GridMcpServerHost McpServerHost { get; private set; }

        public void Start()
        {
            Settings.SaveNormalized();
            if (Settings.McpEnabled)
            {
                McpServerHost.StartAsync(Settings.McpPort, CancellationToken.None).GetAwaiter().GetResult();
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            ConversationService.Dispose();
            McpServerHost.Dispose();
        }
    }
}
