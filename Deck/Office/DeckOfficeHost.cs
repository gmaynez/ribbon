using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Ribbon.Contracts;
using Ribbon.Vsto;

namespace Deck.Office
{
    internal sealed class DeckOfficeHost : IOfficeHost
    {
        private readonly PowerPoint.Application _application;
        private readonly Ribbon.Vsto.OfficeDispatcher _dispatcher;
        private readonly string _hostId = "powerpoint-" + Guid.NewGuid().ToString("N");

        public DeckOfficeHost(PowerPoint.Application application, SynchronizationContext context)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = new Ribbon.Vsto.OfficeDispatcher(context);
        }

        public HostRegistration Registration
        {
            get
            {
                string path = null;
                try { path = _application.ActivePresentation?.FullName; } catch { }
                return new HostRegistration
                {
                    HostId = _hostId,
                    HostKind = "PowerPoint",
                    DisplayName = "Microsoft PowerPoint",
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
                Tool("powerpoint_get_context", "Get the active presentation and selected slide.", "{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}", false),
                Tool("powerpoint_list_slides", "List slides and their visible text in the active presentation.", "{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}", false),
                Tool("powerpoint_read_slide", "Read all text from one slide.", "{\"type\":\"object\",\"properties\":{\"slide_number\":{\"type\":\"integer\",\"minimum\":1}},\"required\":[\"slide_number\"],\"additionalProperties\":false}", false),
                Tool("powerpoint_add_slide", "Add a title-and-content slide to the active presentation.", "{\"type\":\"object\",\"properties\":{\"title\":{\"type\":\"string\"},\"body\":{\"type\":\"string\"}},\"required\":[\"title\",\"body\"],\"additionalProperties\":false}", true)
            };
        }

        public async Task<OfficeToolResult> InvokeAsync(OfficeToolInvocation invocation, CancellationToken cancellationToken)
        {
            try
            {
                var result = await _dispatcher.RunAsync(() => InvokeOnOfficeThread(invocation), cancellationToken).ConfigureAwait(false);
                return new OfficeToolResult { Success = true, ContentJson = JsonCodec.Serialize(result) };
            }
            catch (Exception exception)
            {
                return new OfficeToolResult { Success = false, Error = exception.GetBaseException().Message };
            }
        }

        private object InvokeOnOfficeThread(OfficeToolInvocation invocation)
        {
            var presentation = _application.ActivePresentation ?? throw new InvalidOperationException("PowerPoint does not have an active presentation.");
            switch (invocation.ToolName)
            {
                case "powerpoint_get_context":
                    var selectedSlide = 0;
                    try { selectedSlide = _application.ActiveWindow.View.Slide.SlideIndex; } catch { }
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["path"] = presentation.FullName, ["slide_count"] = presentation.Slides.Count, ["selected_slide"] = selectedSlide };
                case "powerpoint_list_slides":
                    var slides = new List<Dictionary<string, object>>();
                    foreach (PowerPoint.Slide slide in presentation.Slides)
                    {
                        slides.Add(new Dictionary<string, object> { ["slide_number"] = slide.SlideIndex, ["text"] = ReadSlideText(slide) });
                    }
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["slides"] = slides };
                case "powerpoint_read_slide":
                    var read = JsonCodec.Deserialize<SlideRequest>(invocation.ArgumentsJson);
                    if (read.slide_number < 1 || read.slide_number > presentation.Slides.Count) throw new ArgumentOutOfRangeException("slide_number");
                    var requested = presentation.Slides[read.slide_number];
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["slide_number"] = requested.SlideIndex, ["text"] = ReadSlideText(requested) };
                case "powerpoint_add_slide":
                    var add = JsonCodec.Deserialize<AddSlideRequest>(invocation.ArgumentsJson);
                    var created = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
                    created.Shapes.Title.TextFrame.TextRange.Text = add.title ?? string.Empty;
                    if (created.Shapes.Placeholders.Count >= 2) created.Shapes.Placeholders[2].TextFrame.TextRange.Text = add.body ?? string.Empty;
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["slide_number"] = created.SlideIndex };
                default:
                    throw new InvalidOperationException("Unknown PowerPoint tool '" + invocation.ToolName + "'.");
            }
        }

        private static string ReadSlideText(PowerPoint.Slide slide)
        {
            var text = new List<string>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                    && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    text.Add(shape.TextFrame.TextRange.Text);
                }
            }
            return string.Join("\n", text);
        }

        private static OfficeToolDefinition Tool(string name, string description, string schema, bool destructive)
        {
            return new OfficeToolDefinition { Name = name, Description = description, InputSchemaJson = schema, Destructive = destructive, HostKind = "PowerPoint" };
        }

        private sealed class SlideRequest { public int slide_number { get; set; } }
        private sealed class AddSlideRequest { public string title { get; set; } public string body { get; set; } }
    }
}
