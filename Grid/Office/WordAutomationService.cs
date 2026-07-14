using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;

namespace Grid.Office
{
    internal sealed class WordAutomationService
    {
        private readonly OfficeDispatcher _dispatcher;

        public WordAutomationService(OfficeDispatcher dispatcher)
        {
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<Dictionary<string, object>> GetContextAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                dynamic app;
                dynamic document;
                string documentName;

                app = GetApplication(false);
                document = app != null ? app.ActiveDocument : null;
                documentName = document != null ? (string)document.Name : null;

                return new Dictionary<string, object>
                {
                    ["running"] = app != null,
                    ["active_document"] = documentName
                };
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> GetDocumentTextAsync(int maxCharacters, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                dynamic app;
                dynamic document;
                string originalText;
                string text;

                if (maxCharacters <= 0)
                {
                    throw new McpException("Parameter 'max_characters' must be greater than zero.");
                }

                app = RequireApplication(false);
                document = RequireActiveDocument(app);
                originalText = document.Content != null ? (string)document.Content.Text : string.Empty;
                originalText = originalText ?? string.Empty;
                text = originalText;

                if (text.Length > maxCharacters)
                {
                    text = text.Substring(0, maxCharacters);
                }

                return new Dictionary<string, object>
                {
                    ["document"] = (string)document.Name,
                    ["text"] = text,
                    ["truncated"] = originalText.Length > maxCharacters
                };
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> InsertTextAsync(string text, bool replaceSelection, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                dynamic app;
                dynamic document;
                dynamic selection;

                if (string.IsNullOrWhiteSpace(text))
                {
                    throw new McpException("Parameter 'text' is required.");
                }

                app = RequireApplication(true);
                document = EnsureDocument(app);
                selection = app.Selection;

                if (selection != null)
                {
                    if (replaceSelection)
                    {
                        selection.Text = text;
                    }
                    else
                    {
                        selection.InsertAfter(text);
                    }
                }
                else
                {
                    document.Content.InsertAfter(text);
                }

                return new Dictionary<string, object>
                {
                    ["document"] = (string)document.Name,
                    ["inserted_characters"] = text.Length,
                    ["replaced_selection"] = replaceSelection
                };
            }, cancellationToken);
        }

        private static dynamic GetApplication(bool create)
        {
            Type type;
            object application;

            try
            {
                return Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                if (!create)
                {
                    return null;
                }
            }

            type = Type.GetTypeFromProgID("Word.Application");
            if (type == null)
            {
                throw new McpException("Word is not available on this machine.");
            }

            application = Activator.CreateInstance(type);
            type.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, application, new object[] { true });
            return application;
        }

        private static dynamic RequireApplication(bool create)
        {
            dynamic app;

            app = GetApplication(create);
            if (app == null)
            {
                throw new McpException("Word is not currently running.");
            }

            return app;
        }

        private static dynamic RequireActiveDocument(dynamic app)
        {
            dynamic document;

            document = app.ActiveDocument;
            if (document == null)
            {
                throw new McpException("Word does not have an active document.");
            }

            return document;
        }

        private static dynamic EnsureDocument(dynamic app)
        {
            dynamic document;

            document = app.ActiveDocument;
            if (document != null)
            {
                return document;
            }

            return app.Documents.Add();
        }
    }
}
