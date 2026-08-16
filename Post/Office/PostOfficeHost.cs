using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;
using Ribbon.Vsto;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Post.Office
{
    /// <summary>
    /// Outlook (Classic) host slice. Outlook is an item host, not a document host:
    /// its context anchor is the default mailbox store, it offers no checkpoints,
    /// and sending mail is an irreversible operation that always requires an
    /// explicit user confirmation.
    /// </summary>
    internal sealed class PostOfficeHost : IOfficeHost
    {
        private readonly Outlook.Application _application;
        private readonly OutlookAutomationService _automation;
        private readonly string _hostId = "outlook-" + Guid.NewGuid().ToString("N");

        public PostOfficeHost(Outlook.Application application, SynchronizationContext context)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _automation = new OutlookAutomationService(application, new OfficeDispatcher(context));
        }

        public HostRegistration Registration
        {
            get
            {
                string contextId = null;
                string contextName = null;
                try
                {
                    var session = _application.Session;
                    Outlook.Store store = null;
                    try
                    {
                        store = session.DefaultStore;
                        contextId = OfficeDocumentIdentity.Get("outlook", store.StoreID, "store");
                        contextName = "Mailbox · " + store.DisplayName;
                    }
                    finally
                    {
                        ComUtilities.TryRelease(store);
                        ComUtilities.TryRelease(session);
                    }
                }
                catch { }
                return new HostRegistration
                {
                    HostId = _hostId,
                    HostKind = "Outlook",
                    DisplayName = "Microsoft Outlook",
                    ProcessId = Process.GetCurrentProcess().Id,
                    DocumentId = contextId,
                    Version = _application.Version,
                    ContextKind = "store",
                    ContextId = contextId,
                    ContextName = contextName,
                    SupportsCheckpoints = false
                };
            }
        }

        public IList<OfficeToolDefinition> GetTools()
        {
            return new List<OfficeToolDefinition>
            {
                Tool("outlook_get_context", "Inspect the active Outlook profile, default mailbox, open folder, and selection. Call this first when the mailbox or target folder is ambiguous.", OutlookToolSchemas.Empty, false, false),
                Tool("outlook_list_folders", "List the folders of the default mailbox with paths and item counts, bounded to 150 folders.", OutlookToolSchemas.Empty, false, false),
                Tool("outlook_list_items", "List a bounded window of items in one folder with subjects, senders, times, and entry_id values for follow-up reads.", OutlookToolSchemas.ListItems, false, false),
                Tool("outlook_read_item", "Read one email message by entry_id, including a bounded plain-text body. Prefer targeted reads over bulk folder dumps.", OutlookToolSchemas.ReadItem, false, false),
                Tool("outlook_create_draft", "Create a plain-text draft email in the Drafts folder. The draft is never sent by this tool.", OutlookToolSchemas.CreateDraft, true, false),
                Tool("outlook_update_draft", "Patch one draft in the Drafts folder: recipients, subject, and body (replace or append). Omitted fields stay unchanged.", OutlookToolSchemas.UpdateDraft, true, false),
                Tool("outlook_delete_draft", "Permanently delete one draft from the Drafts folder.", OutlookToolSchemas.EntryId, true, false),
                Tool("outlook_send_draft", "Send one draft from the Drafts folder to its recipients. Irreversible: delivered mail cannot be recalled or restored.", OutlookToolSchemas.EntryId, true, true)
            };
        }

        public async Task<OfficeToolResult> InvokeAsync(OfficeToolInvocation invocation, CancellationToken cancellationToken)
        {
            try
            {
                object result;
                switch (invocation.ToolName)
                {
                    case "outlook_get_context":
                        result = await _automation.GetContextAsync(cancellationToken).ConfigureAwait(false);
                        break;
                    case "outlook_list_folders":
                        result = await _automation.ListFoldersAsync(cancellationToken).ConfigureAwait(false);
                        break;
                    case "outlook_list_items":
                        var list = JsonCodec.Deserialize<ListItemsRequest>(invocation.ArgumentsJson);
                        result = await _automation.ListItemsAsync(list, cancellationToken).ConfigureAwait(false);
                        break;
                    case "outlook_read_item":
                        var read = JsonCodec.Deserialize<ReadItemRequest>(invocation.ArgumentsJson);
                        result = await _automation.ReadItemAsync(read, cancellationToken).ConfigureAwait(false);
                        break;
                    case "outlook_create_draft":
                        var create = JsonCodec.Deserialize<CreateDraftRequest>(invocation.ArgumentsJson);
                        result = await _automation.CreateDraftAsync(create, cancellationToken).ConfigureAwait(false);
                        break;
                    case "outlook_update_draft":
                        var update = JsonCodec.Deserialize<UpdateDraftRequest>(invocation.ArgumentsJson);
                        result = await _automation.UpdateDraftAsync(update, cancellationToken).ConfigureAwait(false);
                        break;
                    case "outlook_delete_draft":
                        var delete = JsonCodec.Deserialize<EntryIdRequest>(invocation.ArgumentsJson);
                        result = await _automation.DeleteDraftAsync(delete, cancellationToken).ConfigureAwait(false);
                        break;
                    case "outlook_send_draft":
                        var send = JsonCodec.Deserialize<EntryIdRequest>(invocation.ArgumentsJson);
                        result = await _automation.SendDraftAsync(send, cancellationToken).ConfigureAwait(false);
                        break;
                    default:
                        throw new InvalidOperationException("Unknown Outlook tool '" + invocation.ToolName + "'.");
                }
                return new OfficeToolResult { Success = true, ContentJson = JsonCodec.Serialize(result) };
            }
            catch (Exception exception)
            {
                return new OfficeToolResult { Success = false, Error = exception.GetBaseException().Message };
            }
        }

        private static OfficeToolDefinition Tool(string name, string description, string schema, bool destructive, bool irreversible)
        {
            return new OfficeToolDefinition
            {
                Name = name,
                Description = description,
                InputSchemaJson = schema,
                Destructive = destructive,
                Irreversible = irreversible,
                HostKind = "Outlook"
            };
        }
    }
}
