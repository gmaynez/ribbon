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
        private readonly WordAutomationService _automation;
        private readonly WordCheckpointService _checkpoints;
        private readonly string _hostId = "word-" + Guid.NewGuid().ToString("N");

        public QuillOfficeHost(Word.Application application, SynchronizationContext context)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            var dispatcher = new Ribbon.Vsto.OfficeDispatcher(context);
            _automation = new WordAutomationService(application, dispatcher);
            _checkpoints = new WordCheckpointService(application, dispatcher);
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
                Tool("word_get_context", "Inspect the active Word document, current selection, story, and document statistics. Call this first when the user's target is ambiguous.", WordToolSchemas.Empty, false),
                Tool("word_read_document", "Read a bounded slice of the main document story using stable Word character positions.", WordToolSchemas.ReadDocument, false),
                Tool("word_list_headings", "List heading text, outline levels, and character positions to understand document structure before editing.", WordToolSchemas.ListHeadings, false),
                Tool("word_replace_selection", "Replace the current Word selection with text. Newlines create Word paragraphs.", WordToolSchemas.Text, true),
                Tool("word_append_text", "Append text immediately before the final paragraph mark in the active document.", WordToolSchemas.Text, true),
                Tool("word_insert_text", "Insert text before or after the selection, or at the start or end of the document, without replacing existing text.", WordToolSchemas.InsertText, true),
                Tool("word_replace_range", "Replace or delete an exact character range previously obtained from a Word read or structure tool.", WordToolSchemas.ReplaceRange, true),
                Tool("word_find_replace", "Find and replace literal text in the main document story with bounded replacement count and optional case or whole-word matching.", WordToolSchemas.FindReplace, true),
                Tool("word_format_range", "Apply only the specified style, font, paragraph, or highlight properties to explicit character positions or the current selection.", WordToolSchemas.FormatRange, true),
                Tool("word_insert_heading", "Insert one paragraph and apply a built-in Heading 1 through Heading 9 style.", WordToolSchemas.InsertHeading, true),
                Tool("word_insert_list", "Insert a bulleted or numbered list from structured items.", WordToolSchemas.InsertList, true),
                Tool("word_insert_table", "Insert a populated Word table from a rectangular value matrix with optional header styling and AutoFit behavior.", WordToolSchemas.InsertTable, true),
                Tool("word_add_comment", "Add a review comment to explicit character positions or the current selection.", WordToolSchemas.AddComment, true),
                Tool("word_insert_page_break", "Insert a page break at the selection, document start, or document end.", WordToolSchemas.InsertPageBreak, true)
            };
        }

        public async Task<OfficeToolResult> InvokeAsync(OfficeToolInvocation invocation, CancellationToken cancellationToken)
        {
            try
            {
                object result;
                switch (invocation.ToolName)
                {
                    case "word_get_context":
                        result = await _automation.GetContextAsync(cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_read_document":
                        result = await _automation.ReadDocumentAsync(JsonCodec.Deserialize<ReadDocumentRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_list_headings":
                        result = await _automation.ListHeadingsAsync(JsonCodec.Deserialize<ListHeadingsRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_replace_selection":
                        result = await _automation.ReplaceSelectionAsync(JsonCodec.Deserialize<TextRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_append_text":
                        result = await _automation.AppendTextAsync(JsonCodec.Deserialize<TextRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_insert_text":
                        result = await _automation.InsertTextAsync(JsonCodec.Deserialize<InsertTextRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_replace_range":
                        result = await _automation.ReplaceRangeAsync(JsonCodec.Deserialize<ReplaceRangeRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_find_replace":
                        result = await _automation.FindReplaceAsync(JsonCodec.Deserialize<FindReplaceRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_format_range":
                        result = await _automation.FormatRangeAsync(JsonCodec.Deserialize<FormatWordRangeRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_insert_heading":
                        result = await _automation.InsertHeadingAsync(JsonCodec.Deserialize<InsertHeadingRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_insert_list":
                        result = await _automation.InsertListAsync(JsonCodec.Deserialize<InsertListRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_insert_table":
                        result = await _automation.InsertTableAsync(JsonCodec.Deserialize<InsertTableRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_add_comment":
                        result = await _automation.AddCommentAsync(JsonCodec.Deserialize<AddCommentRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    case "word_insert_page_break":
                        result = await _automation.InsertPageBreakAsync(JsonCodec.Deserialize<InsertPageBreakRequest>(invocation.ArgumentsJson), cancellationToken).ConfigureAwait(false);
                        break;
                    default:
                        throw new InvalidOperationException("Unknown Word tool '" + invocation.ToolName + "'.");
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
            return new OfficeToolDefinition { Name = name, Description = description, InputSchemaJson = schema, Destructive = destructive, HostKind = "Word" };
        }
    }
}
