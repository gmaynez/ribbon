using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;
using Ribbon.Vsto;
using Excel = Microsoft.Office.Interop.Excel;

namespace Grid.Office
{
    internal sealed class GridOfficeHost : IOfficeHost
    {
        private readonly Excel.Application _application;
        private readonly ExcelAutomationService _automation;
        private readonly string _hostId = "excel-" + Guid.NewGuid().ToString("N");

        public GridOfficeHost(Excel.Application application, SynchronizationContext context)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _automation = new ExcelAutomationService(application, new OfficeDispatcher(context));
        }

        public HostRegistration Registration
        {
            get
            {
                string path = null;
                try { path = _application.ActiveWorkbook?.FullName; } catch { }
                return new HostRegistration
                {
                    HostId = _hostId,
                    HostKind = "Excel",
                    DisplayName = "Microsoft Excel",
                    ProcessId = Process.GetCurrentProcess().Id,
                    DocumentPath = path,
                    Version = _application.Version
                };
            }
        }

        public IList<OfficeToolDefinition> GetTools()
        {
            return new List<OfficeToolDefinition>
            {
                Tool("excel_get_context", "Get the active Excel workbook, worksheet, and cell.", "{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}", false),
                Tool("excel_list_sheets", "List worksheets and their used ranges in the active workbook.", "{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}", false),
                Tool("excel_read_range", "Read values from an Excel range in the active workbook.", "{\"type\":\"object\",\"properties\":{\"sheet_name\":{\"type\":\"string\"},\"address\":{\"type\":\"string\"}},\"required\":[\"address\"],\"additionalProperties\":false}", false),
                Tool("excel_write_range", "Write a rectangular array of values to an Excel range.", "{\"type\":\"object\",\"properties\":{\"sheet_name\":{\"type\":\"string\"},\"address\":{\"type\":\"string\"},\"values\":{\"type\":\"array\",\"items\":{\"type\":\"array\"}}},\"required\":[\"address\",\"values\"],\"additionalProperties\":false}", true)
            };
        }

        public async Task<OfficeToolResult> InvokeAsync(OfficeToolInvocation invocation, CancellationToken cancellationToken)
        {
            try
            {
                object result;
                switch (invocation.ToolName)
                {
                    case "excel_get_context":
                        result = await _automation.GetContextAsync(cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_list_sheets":
                        result = await _automation.ListSheetsAsync(cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_read_range":
                        var read = JsonCodec.Deserialize<RangeRequest>(invocation.ArgumentsJson);
                        result = await _automation.ReadRangeAsync(read.sheet_name, read.address, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_write_range":
                        var write = JsonCodec.Deserialize<WriteRangeRequest>(invocation.ArgumentsJson);
                        result = await _automation.WriteRangeAsync(write.sheet_name, write.address, write.values, cancellationToken).ConfigureAwait(false);
                        break;
                    default:
                        throw new InvalidOperationException("Unknown Excel tool '" + invocation.ToolName + "'.");
                }
                return new OfficeToolResult { Success = true, ContentJson = JsonCodec.Serialize(result) };
            }
            catch (Exception exception)
            {
                return new OfficeToolResult { Success = false, Error = exception.GetBaseException().Message };
            }
        }

        private static OfficeToolDefinition Tool(string name, string description, string schema, bool destructive)
        {
            return new OfficeToolDefinition { Name = name, Description = description, InputSchemaJson = schema, Destructive = destructive, HostKind = "Excel" };
        }

        private class RangeRequest
        {
            public string sheet_name { get; set; }
            public string address { get; set; }
        }

        private sealed class WriteRangeRequest : RangeRequest
        {
            public List<List<object>> values { get; set; }
        }
    }
}
