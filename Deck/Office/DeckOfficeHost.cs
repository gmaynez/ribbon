using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;
using Ribbon.Vsto;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Deck.Office
{
    internal sealed class DeckOfficeHost : ICheckpointHost
    {
        private readonly PowerPoint.Application _application;
        private readonly PowerPointAutomationService _automation;
        private readonly PowerPointCheckpointService _checkpoints;
        private readonly string _hostId = "powerpoint-" + Guid.NewGuid().ToString("N");

        public DeckOfficeHost(PowerPoint.Application application, SynchronizationContext context)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            var dispatcher = new OfficeDispatcher(context);
            _automation = new PowerPointAutomationService(application, dispatcher);
            _checkpoints = new PowerPointCheckpointService(application, dispatcher);
        }

        public HostRegistration Registration
        {
            get
            {
                string path = null;
                string documentId = null;
                string documentName = null;
                try
                {
                    var presentation = _application.ActivePresentation;
                    path = presentation?.FullName;
                    documentName = presentation?.Name;
                    var windowHandle = PresentationWindowHandle(presentation);
                    documentId = OfficeDocumentIdentity.Get("powerpoint", windowHandle == 0
                        ? path
                        : Process.GetCurrentProcess().Id + "|" + windowHandle);
                }
                catch { }
                return new HostRegistration
                {
                    HostId = _hostId,
                    HostKind = "PowerPoint",
                    DisplayName = "Microsoft PowerPoint",
                    ProcessId = Process.GetCurrentProcess().Id,
                    DocumentId = documentId,
                    DocumentPath = path,
                    Version = _application.Version,
                    ContextKind = "document",
                    ContextId = documentId,
                    ContextName = documentName,
                    SupportsCheckpoints = true
                };
            }
        }

        private static long PresentationWindowHandle(PowerPoint.Presentation presentation)
        {
            if (presentation == null) return 0;
            PowerPoint.DocumentWindows windows = null;
            PowerPoint.DocumentWindow window = null;
            try
            {
                windows = presentation.Windows;
                if (windows == null || windows.Count == 0) return 0;
                window = windows[1];
                return window?.HWND ?? 0;
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
                Tool("powerpoint_get_context", "Inspect the active presentation, page size, selected slide, and selected shapes. Call this first when the user's target is ambiguous.", PowerPointToolSchemas.Empty, false),
                Tool("powerpoint_list_slides", "List a bounded presentation outline with slide titles, layouts, shape counts, and text summaries.", PowerPointToolSchemas.ListSlides, false),
                Tool("powerpoint_read_slide", "Read one slide as structured shapes with stable shape names, geometry, text, table metadata, and optional speaker notes.", PowerPointToolSchemas.ReadSlide, false),
                Tool("powerpoint_add_slide", "Insert a slide using a supported PowerPoint layout and optionally populate its title and body.", PowerPointToolSchemas.AddSlide, true),
                Tool("powerpoint_delete_slide", "Delete one slide by its current one-based slide number.", PowerPointToolSchemas.Slide, true),
                Tool("powerpoint_duplicate_slide", "Duplicate one slide immediately after its source and return the new slide identity.", PowerPointToolSchemas.Slide, true),
                Tool("powerpoint_move_slide", "Move one slide to a new one-based position in the active presentation.", PowerPointToolSchemas.MoveSlide, true),
                Tool("powerpoint_set_slide_title", "Set a slide title, creating a title text box when the layout has no title placeholder.", PowerPointToolSchemas.SetSlideTitle, true),
                Tool("powerpoint_add_textbox", "Add a positioned text box in points with optional font, alignment, fill, and line formatting.", PowerPointToolSchemas.AddTextBox, true),
                Tool("powerpoint_add_shape", "Add a supported diagram shape in points with optional text and formatting.", PowerPointToolSchemas.AddShape, true),
                Tool("powerpoint_format_shape", "Patch only the supplied geometry, appearance, text, text formatting, or z-order properties of a named shape.", PowerPointToolSchemas.FormatShape, true),
                Tool("powerpoint_delete_shape", "Delete a shape identified by the shape_name returned from powerpoint_read_slide or a creation tool.", PowerPointToolSchemas.Shape, true),
                Tool("powerpoint_add_image", "Place an existing local image on a slide. The path must be absolute; Ribbon does not download images in the Office process.", PowerPointToolSchemas.AddImage, true),
                Tool("powerpoint_add_table", "Add a populated PowerPoint table from a rectangular value matrix with optional header, fill, and text formatting.", PowerPointToolSchemas.AddTable, true),
                Tool("powerpoint_add_chart", "Add a native PowerPoint chart from bounded categories and numeric series with a supported chart type, title, and legend placement.", PowerPointToolSchemas.AddChart, true),
                Tool("powerpoint_set_speaker_notes", "Replace the speaker-notes body for one slide.", PowerPointToolSchemas.SetSpeakerNotes, true),
                Tool("powerpoint_set_slide_background", "Set a solid RGB background color on one slide and detach it from the master background.", PowerPointToolSchemas.SetSlideBackground, true),
                Tool("powerpoint_find_replace", "Find and replace literal text across slide shapes, optionally limited to one slide and speaker notes, with a bounded replacement count.", PowerPointToolSchemas.FindReplace, true)
            };
        }

        public async Task<OfficeToolResult> InvokeAsync(OfficeToolInvocation invocation, CancellationToken cancellationToken)
        {
            try
            {
                object result;
                switch (invocation.ToolName)
                {
                    case "powerpoint_get_context": result = await _automation.GetContextAsync(cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_list_slides": result = await _automation.ListSlidesAsync(JsonCodec.Deserialize<ListSlidesRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_read_slide": result = await _automation.ReadSlideAsync(JsonCodec.Deserialize<ReadSlideRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_add_slide": result = await _automation.AddSlideAsync(JsonCodec.Deserialize<AddSlideRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_delete_slide": result = await _automation.DeleteSlideAsync(JsonCodec.Deserialize<SlideRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_duplicate_slide": result = await _automation.DuplicateSlideAsync(JsonCodec.Deserialize<SlideRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_move_slide": result = await _automation.MoveSlideAsync(JsonCodec.Deserialize<MoveSlideRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_set_slide_title": result = await _automation.SetSlideTitleAsync(JsonCodec.Deserialize<SetSlideTitleRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_add_textbox": result = await _automation.AddTextBoxAsync(JsonCodec.Deserialize<AddTextBoxRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_add_shape": result = await _automation.AddShapeAsync(JsonCodec.Deserialize<AddShapeRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_format_shape": result = await _automation.FormatShapeAsync(JsonCodec.Deserialize<FormatShapeRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_delete_shape": result = await _automation.DeleteShapeAsync(JsonCodec.Deserialize<ShapeRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_add_image": result = await _automation.AddImageAsync(JsonCodec.Deserialize<AddImageRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_add_table": result = await _automation.AddTableAsync(JsonCodec.Deserialize<AddTableRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_add_chart": result = await _automation.AddChartAsync(JsonCodec.Deserialize<AddChartRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_set_speaker_notes": result = await _automation.SetSpeakerNotesAsync(JsonCodec.Deserialize<SetSpeakerNotesRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_set_slide_background": result = await _automation.SetSlideBackgroundAsync(JsonCodec.Deserialize<SetSlideBackgroundRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    case "powerpoint_find_replace": result = await _automation.FindReplaceAsync(JsonCodec.Deserialize<FindReplaceRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false); break;
                    default: throw new InvalidOperationException("Unknown PowerPoint tool '" + invocation.ToolName + "'.");
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
            return new OfficeToolDefinition { Name = name, Description = description, InputSchemaJson = schema, Destructive = destructive, HostKind = "PowerPoint" };
        }
    }
}
