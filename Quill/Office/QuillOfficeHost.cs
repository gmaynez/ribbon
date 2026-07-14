using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;
using Ribbon.Vsto;
using Word = Microsoft.Office.Interop.Word;

namespace Quill.Office
{
    internal sealed class QuillOfficeHost : IOfficeHost
    {
        private readonly Word.Application _application;
        private readonly Ribbon.Vsto.OfficeDispatcher _dispatcher;
        private readonly string _hostId = "word-" + Guid.NewGuid().ToString("N");

        public QuillOfficeHost(Word.Application application, SynchronizationContext context)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = new Ribbon.Vsto.OfficeDispatcher(context);
        }

        public HostRegistration Registration
        {
            get
            {
                string path = null;
                try { path = _application.ActiveDocument?.FullName; } catch { }
                return new HostRegistration
                {
                    HostId = _hostId,
                    HostKind = "Word",
                    DisplayName = "Microsoft Word",
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
                Tool("word_get_context", "Get the active Word document and current selection.", "{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}", false),
                Tool("word_read_document", "Read text from the active Word document.", "{\"type\":\"object\",\"properties\":{\"max_characters\":{\"type\":\"integer\",\"minimum\":1,\"maximum\":200000}},\"additionalProperties\":false}", false),
                Tool("word_replace_selection", "Replace the current Word selection with text.", "{\"type\":\"object\",\"properties\":{\"text\":{\"type\":\"string\"}},\"required\":[\"text\"],\"additionalProperties\":false}", true),
                Tool("word_append_text", "Append text at the end of the active Word document.", "{\"type\":\"object\",\"properties\":{\"text\":{\"type\":\"string\"}},\"required\":[\"text\"],\"additionalProperties\":false}", true)
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
            var document = _application.ActiveDocument ?? throw new InvalidOperationException("Word does not have an active document.");
            switch (invocation.ToolName)
            {
                case "word_get_context":
                    var selection = _application.Selection;
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["path"] = document.FullName,
                        ["selection_start"] = selection.Start,
                        ["selection_end"] = selection.End,
                        ["selection_text"] = selection.Text
                    };
                case "word_read_document":
                    var read = JsonCodec.Deserialize<ReadRequest>(invocation.ArgumentsJson);
                    var maximum = read.max_characters <= 0 ? 50000 : Math.Min(read.max_characters, 200000);
                    var text = document.Content.Text ?? string.Empty;
                    var truncated = text.Length > maximum;
                    if (truncated) text = text.Substring(0, maximum);
                    return new Dictionary<string, object> { ["document"] = document.Name, ["text"] = text, ["truncated"] = truncated, ["character_count"] = document.Content.End };
                case "word_replace_selection":
                    var replace = JsonCodec.Deserialize<TextRequest>(invocation.ArgumentsJson);
                    if (replace.text == null) throw new ArgumentException("Parameter 'text' is required.");
                    var current = _application.Selection;
                    current.Text = replace.text;
                    return new Dictionary<string, object> { ["document"] = document.Name, ["inserted_characters"] = replace.text.Length };
                case "word_append_text":
                    var append = JsonCodec.Deserialize<TextRequest>(invocation.ArgumentsJson);
                    if (append.text == null) throw new ArgumentException("Parameter 'text' is required.");
                    var range = document.Content;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    range.InsertAfter(append.text);
                    return new Dictionary<string, object> { ["document"] = document.Name, ["appended_characters"] = append.text.Length };
                default:
                    throw new InvalidOperationException("Unknown Word tool '" + invocation.ToolName + "'.");
            }
        }

        private static OfficeToolDefinition Tool(string name, string description, string schema, bool destructive)
        {
            return new OfficeToolDefinition { Name = name, Description = description, InputSchemaJson = schema, Destructive = destructive, HostKind = "Word" };
        }

        private sealed class ReadRequest { public int max_characters { get; set; } }
        private sealed class TextRequest { public string text { get; set; } }
    }
}
