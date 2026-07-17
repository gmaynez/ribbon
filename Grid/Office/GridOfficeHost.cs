using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
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
        private readonly ExcelCheckpointService _checkpoints;
        private readonly string _hostId = "excel-" + Guid.NewGuid().ToString("N");

        public GridOfficeHost(Excel.Application application, SynchronizationContext context)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            var dispatcher = new OfficeDispatcher(context);
            _automation = new ExcelAutomationService(application, dispatcher);
            _checkpoints = new ExcelCheckpointService(application, dispatcher);
        }

        public HostRegistration Registration
        {
            get
            {
                string path = null;
                string documentId = null;
                try
                {
                    var workbook = _application.ActiveWorkbook;
                    path = workbook?.FullName;
                    var windowHandle = WorkbookWindowHandle(workbook);
                    documentId = OfficeDocumentIdentity.Get("excel", windowHandle == 0
                        ? path
                        : Process.GetCurrentProcess().Id + "|" + windowHandle);
                }
                catch { }
                return new HostRegistration
                {
                    HostId = _hostId,
                    HostKind = "Excel",
                    DisplayName = "Microsoft Excel",
                    ProcessId = Process.GetCurrentProcess().Id,
                    DocumentId = documentId,
                    DocumentPath = path,
                    Version = _application.Version
                };
            }
        }

        private static long WorkbookWindowHandle(Excel.Workbook workbook)
        {
            if (workbook == null) return 0;
            Excel.Windows windows = null;
            Excel.Window window = null;
            try
            {
                windows = workbook.Windows;
                if (windows == null || windows.Count == 0) return 0;
                window = windows[1];
                return window?.Hwnd ?? 0;
            }
            catch { return 0; }
            finally
            {
                if (window != null) Marshal.ReleaseComObject(window);
                if (windows != null) Marshal.ReleaseComObject(windows);
            }
        }

        public IList<OfficeToolDefinition> GetTools()
        {
            return new List<OfficeToolDefinition>
            {
                Tool("excel_get_context", "Inspect the active workbook, worksheet, selected range, active cell, and used range. Call this first when the user's target is ambiguous.", ExcelToolSchemas.Empty, false),
                Tool("excel_list_sheets", "List worksheets, visibility, and used ranges in the active workbook.", ExcelToolSchemas.Empty, false),
                Tool("excel_read_range", "Read a bounded Excel range as JSON-safe values, with formulas by default and optional number formats. Prefer reading existing data before changing it.", ExcelToolSchemas.ReadRange, false),
                Tool("excel_write_range", "Write a rectangular matrix of literal values beginning at an A1 address. Existing cells in the resized target are replaced.", ExcelToolSchemas.WriteRange, true),
                Tool("excel_write_formulas", "Write a rectangular matrix of A1-style Excel formulas. Each non-null formula must begin with '='.", ExcelToolSchemas.WriteFormulas, true),
                Tool("excel_clear_range", "Clear contents, formatting, or everything from an Excel range.", ExcelToolSchemas.ClearRange, true),
                Tool("excel_format_range", "Apply only the specified formatting properties to a range; unspecified formatting remains unchanged. Supports fonts, fills, number formats, alignment, borders, sizing, and AutoFit.", ExcelToolSchemas.FormatRange, true),
                Tool("excel_add_sheet", "Add a named worksheet before or after the active sheet, or at the end of the workbook.", ExcelToolSchemas.AddSheet, true),
                Tool("excel_create_table", "Convert an existing range into a styled Excel table with optional headers, table name, and built-in table style.", ExcelToolSchemas.CreateTable, true),
                Tool("excel_create_chart", "Create an embedded chart from an existing source range and place it at a target cell or immediately to the right of the data.", ExcelToolSchemas.CreateChart, true)
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
                        var read = JsonCodec.Deserialize<ReadRangeRequest>(invocation.ArgumentsJson);
                        result = await _automation.ReadRangeAsync(read, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_write_range":
                        var write = JsonCodec.Deserialize<WriteRangeRequest>(invocation.ArgumentsJson);
                        result = await _automation.WriteRangeAsync(write, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_write_formulas":
                        var formulas = JsonCodec.Deserialize<WriteFormulaRequest>(invocation.ArgumentsJson);
                        result = await _automation.WriteFormulasAsync(formulas, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_clear_range":
                        var clear = JsonCodec.Deserialize<ClearRangeRequest>(invocation.ArgumentsJson);
                        result = await _automation.ClearRangeAsync(clear, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_format_range":
                        var format = JsonCodec.Deserialize<FormatRangeRequest>(invocation.ArgumentsJson);
                        result = await _automation.FormatRangeAsync(format, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_add_sheet":
                        var addSheet = JsonCodec.Deserialize<AddSheetRequest>(invocation.ArgumentsJson);
                        result = await _automation.AddSheetAsync(addSheet, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_create_table":
                        var table = JsonCodec.Deserialize<CreateTableRequest>(invocation.ArgumentsJson);
                        result = await _automation.CreateTableAsync(table, cancellationToken).ConfigureAwait(false);
                        break;
                    case "excel_create_chart":
                        var chart = JsonCodec.Deserialize<CreateChartRequest>(invocation.ArgumentsJson);
                        result = await _automation.CreateChartAsync(chart, cancellationToken).ConfigureAwait(false);
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

        public Task<DocumentCheckpoint> CreateCheckpointAsync(string label, CancellationToken cancellationToken)
        {
            return _checkpoints.CreateAsync(Registration, label, cancellationToken);
        }

        public Task RestoreCheckpointAsync(DocumentCheckpoint checkpoint, CancellationToken cancellationToken)
        {
            return _checkpoints.RestoreAsync(Registration, checkpoint, cancellationToken);
        }

        private static OfficeToolDefinition Tool(string name, string description, string schema, bool destructive)
        {
            return new OfficeToolDefinition { Name = name, Description = description, InputSchemaJson = schema, Destructive = destructive, HostKind = "Excel" };
        }
    }
}
