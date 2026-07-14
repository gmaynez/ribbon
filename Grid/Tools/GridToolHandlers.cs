using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Grid.Constants;
using Grid.Office;

namespace Grid.Tools
{
    [McpServerToolType]
    internal sealed class GridToolHandlers
    {
        private readonly OfficeAutomationService _officeAutomation;

        public GridToolHandlers(OfficeAutomationService officeAutomation)
        {
            _officeAutomation = officeAutomation;
        }

        [McpServerTool(Name = "office_list_running_apps")]
        [Description("Lists whether Excel, Word, and PowerPoint are currently running and reachable for automation.")]
        public Task<Dictionary<string, object>> OfficeListRunningApps(CancellationToken cancellationToken)
        {
            return _officeAutomation.ListRunningAppsAsync(cancellationToken);
        }

        [McpServerTool(Name = "office_get_active_context")]
        [Description("Returns the current active workbook, worksheet, and cell in Excel plus active document or presentation context if Word or PowerPoint are running.")]
        public Task<Dictionary<string, object>> OfficeGetActiveContext(CancellationToken cancellationToken)
        {
            return _officeAutomation.GetActiveContextAsync(cancellationToken);
        }

        [McpServerTool(Name = "excel_list_sheets")]
        [Description("Lists worksheets in the active Excel workbook.")]
        public Task<Dictionary<string, object>> ExcelListSheets(CancellationToken cancellationToken)
        {
            return _officeAutomation.Excel.ListSheetsAsync(cancellationToken);
        }

        [McpServerTool(Name = "excel_read_range")]
        [Description("Reads cell values from an address in the active Excel workbook.")]
        public Task<Dictionary<string, object>> ExcelReadRange(
            [Description("Worksheet name. Leave empty to use the active sheet.")] string sheet_name,
            [Description("Excel address like A1, B2:D5, or NamedRange.")] string address,
            CancellationToken cancellationToken)
        {
            return _officeAutomation.Excel.ReadRangeAsync(sheet_name, address, cancellationToken);
        }

        [McpServerTool(Name = "excel_write_range", Destructive = true)]
        [Description("Writes a two-dimensional array of values into Excel starting at the specified address.")]
        public Task<Dictionary<string, object>> ExcelWriteRange(
            [Description("Worksheet name. Leave empty to use the active sheet.")] string sheet_name,
            [Description("Top-left Excel address where writing should start, for example A1.")] string address,
            [Description("Two-dimensional array of row values, for example [[1,2],[3,4]].")] List<List<object>> values,
            CancellationToken cancellationToken)
        {
            return _officeAutomation.Excel.WriteRangeAsync(sheet_name, address, values, cancellationToken);
        }

        [McpServerTool(Name = "word_get_document_text")]
        [Description("Gets text from the active Word document.")]
        public Task<Dictionary<string, object>> WordGetDocumentText(
            [Description("Maximum number of characters to return.")] int max_characters = GridConstants.DefaultWordTextLimit,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            return _officeAutomation.Word.GetDocumentTextAsync(max_characters, cancellationToken);
        }

        [McpServerTool(Name = "word_insert_text", Destructive = true)]
        [Description("Inserts text into Word at the current selection or document end.")]
        public Task<Dictionary<string, object>> WordInsertText(
            [Description("Text to insert into Word.")] string text,
            [Description("When true, replace the current selection instead of inserting after it.")] bool replace_selection = false,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            return _officeAutomation.Word.InsertTextAsync(text, replace_selection, cancellationToken);
        }

        [McpServerTool(Name = "powerpoint_list_slides")]
        [Description("Lists slides in the active PowerPoint presentation and returns the text found on each slide.")]
        public Task<Dictionary<string, object>> PowerPointListSlides(CancellationToken cancellationToken = default(CancellationToken))
        {
            return _officeAutomation.PowerPoint.ListSlidesAsync(cancellationToken);
        }

        [McpServerTool(Name = "powerpoint_get_slide_text")]
        [Description("Gets the text content from a single PowerPoint slide.")]
        public Task<Dictionary<string, object>> PowerPointGetSlideText(
            [Description("One-based slide number.")] int slide_number,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            return _officeAutomation.PowerPoint.GetSlideTextAsync(slide_number, cancellationToken);
        }

        [McpServerTool(Name = "powerpoint_add_slide", Destructive = true)]
        [Description("Adds a new title and body slide to the active PowerPoint presentation.")]
        public Task<Dictionary<string, object>> PowerPointAddSlide(
            [Description("Slide title text.")] string title,
            [Description("Slide body text.")] string body_text,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            return _officeAutomation.PowerPoint.AddSlideAsync(title, body_text, cancellationToken);
        }
    }
}
