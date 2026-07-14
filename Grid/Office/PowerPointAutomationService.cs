using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;

namespace Grid.Office
{
    internal sealed class PowerPointAutomationService
    {
        private const int PpLayoutText = 2;
        private readonly OfficeDispatcher _dispatcher;

        public PowerPointAutomationService(OfficeDispatcher dispatcher)
        {
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<Dictionary<string, object>> GetContextAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                dynamic app;
                dynamic presentation;

                app = GetApplication(false);
                presentation = app != null ? app.ActivePresentation : null;

                return new Dictionary<string, object>
                {
                    ["running"] = app != null,
                    ["active_presentation"] = presentation != null ? (string)presentation.Name : null,
                    ["slide_count"] = presentation != null ? (int)presentation.Slides.Count : 0
                };
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ListSlidesAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                dynamic presentation;
                List<Dictionary<string, object>> slides;
                int index;

                presentation = RequireActivePresentation(false);
                slides = new List<Dictionary<string, object>>();

                for (index = 1; index <= (int)presentation.Slides.Count; index++)
                {
                    dynamic slide;
                    slide = presentation.Slides[index];
                    slides.Add(new Dictionary<string, object>
                    {
                        ["slide_number"] = index,
                        ["text"] = ExtractSlideText(slide)
                    });
                }

                return new Dictionary<string, object>
                {
                    ["presentation"] = (string)presentation.Name,
                    ["slides"] = slides
                };
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> GetSlideTextAsync(int slideNumber, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                dynamic presentation;
                dynamic slide;

                presentation = RequireActivePresentation(false);
                if (slideNumber <= 0 || slideNumber > (int)presentation.Slides.Count)
                {
                    throw new McpException("Parameter 'slide_number' is out of range.");
                }

                slide = presentation.Slides[slideNumber];
                return new Dictionary<string, object>
                {
                    ["presentation"] = (string)presentation.Name,
                    ["slide_number"] = slideNumber,
                    ["text"] = ExtractSlideText(slide)
                };
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddSlideAsync(string title, string bodyText, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                dynamic presentation;
                dynamic slide;
                int slideNumber;

                if (string.IsNullOrWhiteSpace(title) && string.IsNullOrWhiteSpace(bodyText))
                {
                    throw new McpException("At least one of 'title' or 'body_text' must be provided.");
                }

                presentation = EnsurePresentation();
                slideNumber = (int)presentation.Slides.Count + 1;
                slide = presentation.Slides.Add(slideNumber, PpLayoutText);

                SetPlaceholderText(slide, 1, title ?? string.Empty);
                SetPlaceholderText(slide, 2, bodyText ?? string.Empty);

                return new Dictionary<string, object>
                {
                    ["presentation"] = (string)presentation.Name,
                    ["slide_number"] = slideNumber,
                    ["title"] = title ?? string.Empty
                };
            }, cancellationToken);
        }

        private static dynamic GetApplication(bool create)
        {
            Type type;
            object application;

            try
            {
                return Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch
            {
                if (!create)
                {
                    return null;
                }
            }

            type = Type.GetTypeFromProgID("PowerPoint.Application");
            if (type == null)
            {
                throw new McpException("PowerPoint is not available on this machine.");
            }

            application = Activator.CreateInstance(type);
            type.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, application, new object[] { true });
            return application;
        }

        private static dynamic RequireActivePresentation(bool create)
        {
            dynamic app;
            dynamic presentation;

            app = GetApplication(create);
            if (app == null)
            {
                throw new McpException("PowerPoint is not currently running.");
            }

            presentation = app.ActivePresentation;
            if (presentation == null)
            {
                throw new McpException("PowerPoint does not have an active presentation.");
            }

            return presentation;
        }

        private static dynamic EnsurePresentation()
        {
            dynamic app;
            dynamic presentation;

            app = GetApplication(true);
            presentation = app.ActivePresentation;
            if (presentation != null)
            {
                return presentation;
            }

            return app.Presentations.Add(true);
        }

        private static void SetPlaceholderText(dynamic slide, int placeholderIndex, string text)
        {
            try
            {
                slide.Shapes.Placeholders[placeholderIndex].TextFrame.TextRange.Text = text;
            }
            catch
            {
            }
        }

        private static string ExtractSlideText(dynamic slide)
        {
            List<string> textParts;
            int index;

            textParts = new List<string>();

            for (index = 1; index <= (int)slide.Shapes.Count; index++)
            {
                dynamic shape;
                string text;

                shape = slide.Shapes[index];
                try
                {
                    if (shape.HasTextFrame == -1 && shape.TextFrame.HasText == -1)
                    {
                        text = (string)shape.TextFrame.TextRange.Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            textParts.Add(text.Trim());
                        }
                    }
                }
                catch
                {
                }
            }

            return string.Join(Environment.NewLine, textParts.ToArray());
        }
    }
}
