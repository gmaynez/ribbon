using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Vsto;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Post.Office
{
    internal sealed class OutlookAutomationService
    {
        private const int DefaultMaximumListedItems = 25;
        private const int HardMaximumListedItems = 200;
        private const int DefaultMaximumBodyCharacters = 8000;
        private const int HardMaximumBodyCharacters = 20000;
        private const int HardMaximumDraftBodyCharacters = 100000;
        private const int HardMaximumListedFolders = 150;

        private readonly Outlook.Application _application;
        private readonly OfficeDispatcher _dispatcher;

        public OutlookAutomationService(Outlook.Application application, OfficeDispatcher dispatcher)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<Dictionary<string, object>> GetContextAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Outlook.NameSpace session = null;
                Outlook.Store store = null;
                Outlook.Explorer explorer = null;
                Outlook.MAPIFolder folder = null;
                Outlook.Selection selection = null;
                try
                {
                    session = _application.Session;
                    store = session.DefaultStore;
                    explorer = _application.ActiveExplorer();
                    folder = explorer?.CurrentFolder;
                    selection = explorer?.Selection;
                    var storeId = store?.StoreID ?? string.Empty;
                    var storeName = store?.DisplayName ?? string.Empty;
                    return new Dictionary<string, object>
                    {
                        ["running"] = true,
                        ["profile"] = SafeText(() => session.CurrentProfileName),
                        ["version"] = SafeText(() => _application.Version),
                        ["mailbox_name"] = storeName,
                        ["context_id"] = OfficeDocumentIdentity.Get("outlook", storeId, "store"),
                        ["active_folder"] = SafeText(() => folder?.FolderPath),
                        ["selected_item_count"] = selection?.Count ?? 0,
                        ["supports_checkpoints"] = false
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(selection);
                    ComUtilities.TryRelease(folder);
                    ComUtilities.TryRelease(explorer);
                    ComUtilities.TryRelease(store);
                    ComUtilities.TryRelease(session);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ListFoldersAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Outlook.NameSpace session = null;
                Outlook.Store store = null;
                Outlook.MAPIFolder root = null;
                try
                {
                    session = _application.Session;
                    store = session.DefaultStore;
                    root = store.GetRootFolder();
                    var folders = new List<Dictionary<string, object>>();
                    CollectFolders(root, root.FolderPath, 1, folders);
                    return new Dictionary<string, object>
                    {
                        ["mailbox"] = store.DisplayName,
                        ["root_path"] = root.FolderPath,
                        ["folder_count"] = folders.Count,
                        ["folders"] = folders
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(root);
                    ComUtilities.TryRelease(store);
                    ComUtilities.TryRelease(session);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ListItemsAsync(ListItemsRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            var maximum = request.max_items ?? DefaultMaximumListedItems;
            if (maximum < 1 || maximum > HardMaximumListedItems)
            {
                throw new ArgumentOutOfRangeException("max_items", "max_items must be between 1 and 200.");
            }
            return _dispatcher.RunAsync(delegate
            {
                Outlook.NameSpace session = null;
                Outlook.MAPIFolder folder = null;
                Outlook.Items items = null;
                try
                {
                    session = _application.Session;
                    folder = ResolveFolder(session, request.folder_path);
                    items = folder.Items;
                    try { items.Sort("[LastModificationTime]", true); } catch { }

                    var listed = new List<Dictionary<string, object>>();
                    var total = Math.Min(items.Count, maximum);
                    for (var index = 1; index <= total; index++)
                    {
                        object raw = null;
                        Outlook.MailItem mail = null;
                        try
                        {
                            raw = items[index];
                            mail = raw as Outlook.MailItem;
                            var entry = new Dictionary<string, object>
                            {
                                ["entry_id"] = SafeText(() => ((dynamic)raw).EntryID),
                                ["subject"] = SafeText(() => ((dynamic)raw).Subject),
                                ["message_class"] = SafeText(() => ((dynamic)raw).MessageClass),
                                ["item_kind"] = ItemKind(SafeText(() => ((dynamic)raw).MessageClass)),
                                ["modified_utc"] = SafeDate(() => ((dynamic)raw).LastModificationTime)
                            };
                            if (mail != null)
                            {
                                entry["sender"] = SafeText(() => mail.SenderName);
                                entry["received_utc"] = SafeDate(() => mail.ReceivedTime);
                                entry["unread"] = mail.UnRead;
                                entry["has_attachments"] = mail.Attachments?.Count > 0;
                            }
                            listed.Add(entry);
                        }
                        finally
                        {
                            ComUtilities.TryRelease(mail);
                            ComUtilities.TryRelease(raw);
                        }
                    }

                    return new Dictionary<string, object>
                    {
                        ["folder"] = folder.FolderPath,
                        ["item_count"] = items.Count,
                        ["returned_count"] = listed.Count,
                        ["items"] = listed
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(items);
                    ComUtilities.TryRelease(folder);
                    ComUtilities.TryRelease(session);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ReadItemAsync(ReadItemRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.entry_id)) throw new ArgumentException("entry_id is required.");
            var maximum = request.max_characters ?? DefaultMaximumBodyCharacters;
            if (maximum < 100 || maximum > HardMaximumBodyCharacters)
            {
                throw new ArgumentOutOfRangeException("max_characters", "max_characters must be between 100 and 20000.");
            }
            return _dispatcher.RunAsync(delegate
            {
                Outlook.NameSpace session = null;
                Outlook.MailItem mail = null;
                Outlook.Attachments attachments = null;
                Outlook.MAPIFolder parent = null;
                try
                {
                    session = _application.Session;
                    mail = ResolveMailItem(session, request.entry_id);
                    attachments = mail.Attachments;
                    parent = mail.Parent as Outlook.MAPIFolder;
                    var body = mail.Body ?? string.Empty;
                    var truncated = body.Length > maximum;
                    if (truncated) body = body.Substring(0, maximum);
                    return new Dictionary<string, object>
                    {
                        ["entry_id"] = mail.EntryID,
                        ["subject"] = mail.Subject,
                        ["sender"] = SafeText(() => mail.SenderName),
                        ["to"] = SafeText(() => mail.To),
                        ["cc"] = SafeText(() => mail.CC),
                        ["received_utc"] = SafeDate(() => mail.ReceivedTime),
                        ["unread"] = mail.UnRead,
                        ["sent"] = mail.Sent,
                        ["importance"] = ImportanceText(mail.Importance),
                        ["categories"] = SafeText(() => mail.Categories),
                        ["folder"] = SafeText(() => parent?.FolderPath),
                        ["attachment_count"] = attachments?.Count ?? 0,
                        ["body"] = body,
                        ["body_truncated"] = truncated,
                        ["body_length"] = (mail.Body ?? string.Empty).Length
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(parent);
                    ComUtilities.TryRelease(attachments);
                    ComUtilities.TryRelease(mail);
                    ComUtilities.TryRelease(session);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> CreateDraftAsync(CreateDraftRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (BodyTooLong(request.body)) throw new ArgumentOutOfRangeException("body", "The body must not exceed 100000 characters.");
            return _dispatcher.RunAsync(delegate
            {
                Outlook.MailItem draft = null;
                try
                {
                    draft = (Outlook.MailItem)_application.CreateItem(Outlook.OlItemType.olMailItem);
                    draft.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                    if (request.to != null) draft.To = JoinRecipients(request.to);
                    if (request.cc != null) draft.CC = JoinRecipients(request.cc);
                    draft.Subject = request.subject ?? string.Empty;
                    draft.Body = request.body ?? string.Empty;
                    draft.Save();
                    return new Dictionary<string, object>
                    {
                        ["entry_id"] = draft.EntryID,
                        ["subject"] = draft.Subject,
                        ["to"] = SafeText(() => draft.To),
                        ["cc"] = SafeText(() => draft.CC),
                        ["sent"] = false
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(draft);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> UpdateDraftAsync(UpdateDraftRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.entry_id)) throw new ArgumentException("entry_id is required.");
            if (BodyTooLong(request.body)) throw new ArgumentOutOfRangeException("body", "The body must not exceed 100000 characters.");
            return _dispatcher.RunAsync(delegate
            {
                Outlook.NameSpace session = null;
                Outlook.MailItem draft = null;
                try
                {
                    session = _application.Session;
                    draft = ResolveMailItem(session, request.entry_id);
                    RequireDraft(session, draft);
                    if (request.to != null) draft.To = JoinRecipients(request.to);
                    if (request.cc != null) draft.CC = JoinRecipients(request.cc);
                    if (request.subject != null) draft.Subject = request.subject;
                    if (request.body != null)
                    {
                        draft.Body = string.Equals(request.body_mode, "append", StringComparison.OrdinalIgnoreCase)
                            ? (draft.Body ?? string.Empty) + request.body
                            : request.body;
                    }
                    draft.Save();
                    return new Dictionary<string, object>
                    {
                        ["entry_id"] = draft.EntryID,
                        ["subject"] = draft.Subject,
                        ["to"] = SafeText(() => draft.To),
                        ["cc"] = SafeText(() => draft.CC),
                        ["body_length"] = (draft.Body ?? string.Empty).Length,
                        ["sent"] = false
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(draft);
                    ComUtilities.TryRelease(session);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> DeleteDraftAsync(EntryIdRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.entry_id)) throw new ArgumentException("entry_id is required.");
            return _dispatcher.RunAsync(delegate
            {
                Outlook.NameSpace session = null;
                Outlook.MailItem draft = null;
                try
                {
                    session = _application.Session;
                    draft = ResolveMailItem(session, request.entry_id);
                    RequireDraft(session, draft);
                    var summary = new Dictionary<string, object>
                    {
                        ["entry_id"] = draft.EntryID,
                        ["subject"] = draft.Subject,
                        ["to"] = SafeText(() => draft.To),
                        ["deleted"] = true
                    };
                    draft.Delete();
                    return summary;
                }
                finally
                {
                    ComUtilities.TryRelease(draft);
                    ComUtilities.TryRelease(session);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> SendDraftAsync(EntryIdRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.entry_id)) throw new ArgumentException("entry_id is required.");
            return _dispatcher.RunAsync(delegate
            {
                Outlook.NameSpace session = null;
                Outlook.MailItem draft = null;
                Outlook.Recipients recipients = null;
                try
                {
                    session = _application.Session;
                    draft = ResolveMailItem(session, request.entry_id);
                    RequireDraft(session, draft);
                    recipients = draft.Recipients;
                    if (recipients == null || recipients.Count == 0)
                    {
                        throw new InvalidOperationException("The draft has no recipients. Add at least one recipient before sending.");
                    }
                    if (string.IsNullOrWhiteSpace(draft.Subject) && string.IsNullOrWhiteSpace(draft.Body))
                    {
                        throw new InvalidOperationException("The draft has no subject and no body. Update it before sending.");
                    }
                    var summary = new Dictionary<string, object>
                    {
                        ["entry_id"] = draft.EntryID,
                        ["subject"] = draft.Subject,
                        ["to"] = SafeText(() => draft.To),
                        ["cc"] = SafeText(() => draft.CC)
                    };
                    draft.Send();
                    summary["sent"] = true;
                    summary["irreversible"] = true;
                    return summary;
                }
                finally
                {
                    ComUtilities.TryRelease(recipients);
                    ComUtilities.TryRelease(draft);
                    ComUtilities.TryRelease(session);
                }
            }, cancellationToken);
        }

        private static void CollectFolders(Outlook.MAPIFolder folder, string rootPath, int depth, List<Dictionary<string, object>> collected)
        {
            if (folder == null || collected.Count >= HardMaximumListedFolders) return;
            Outlook.Folders children = null;
            try
            {
                var entry = new Dictionary<string, object>
                {
                    ["path"] = SafeText(() => folder.FolderPath),
                    ["name"] = SafeText(() => folder.Name),
                    ["item_count"] = SafeCount(() => folder.Items.Count)
                };
                collected.Add(entry);
                if (depth >= 4) return;
                children = folder.Folders;
                for (var index = 1; index <= children.Count && collected.Count < HardMaximumListedFolders; index++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = children[index];
                        CollectFolders(child, rootPath, depth + 1, collected);
                    }
                    catch
                    {
                        // A folder that cannot be enumerated must not break the whole listing.
                    }
                    finally
                    {
                        ComUtilities.TryRelease(child);
                    }
                }
            }
            finally
            {
                ComUtilities.TryRelease(children);
            }
        }

        private Outlook.MAPIFolder ResolveFolder(Outlook.NameSpace session, string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                var explorer = _application.ActiveExplorer();
                try
                {
                    var current = explorer != null ? explorer.CurrentFolder : null;
                    if (current != null) return current;
                    return GetDefaultFolderSafe(session, Outlook.OlDefaultFolders.olFolderInbox);
                }
                finally
                {
                    ComUtilities.TryRelease(explorer);
                }
            }

            var path = folderPath.Trim().TrimStart('\\');
            if (path.Length == 0)
            {
                Outlook.Store store = null;
                Outlook.MAPIFolder root = null;
                try
                {
                    store = session.DefaultStore;
                    root = store.GetRootFolder();
                    return root;
                }
                finally
                {
                    // root and store RCWs are owned by the caller after return; the
                    // caller releases the resolved folder, which releases its children.
                    ComUtilities.TryRelease(store);
                }
            }

            var segments = path.Split('\\');
            Outlook.MAPIFolder currentFolder = ResolveFirstSegment(session, segments[0]);
            try
            {
                for (var index = 1; index < segments.Length; index++)
                {
                    var segment = segments[index].Trim();
                    if (segment.Length == 0) continue;
                    var next = FindChild(currentFolder, segment);
                    ComUtilities.TryRelease(currentFolder);
                    currentFolder = next;
                    if (currentFolder == null)
                    {
                        throw new InvalidOperationException("No Outlook folder matches '" + folderPath + "'. Use outlook_list_folders to discover valid paths.");
                    }
                }
                var resolved = currentFolder;
                currentFolder = null;
                return resolved;
            }
            finally
            {
                ComUtilities.TryRelease(currentFolder);
            }
        }

        private Outlook.MAPIFolder ResolveFirstSegment(Outlook.NameSpace session, string segment)
        {
            var trimmed = segment.Trim();
            if (string.Equals(trimmed, "Inbox", StringComparison.OrdinalIgnoreCase))
            {
                return GetDefaultFolderSafe(session, Outlook.OlDefaultFolders.olFolderInbox);
            }
            if (string.Equals(trimmed, "Drafts", StringComparison.OrdinalIgnoreCase))
            {
                return GetDefaultFolderSafe(session, Outlook.OlDefaultFolders.olFolderDrafts);
            }

            Outlook.Store store = null;
            Outlook.MAPIFolder root = null;
            Outlook.MAPIFolder child = null;
            try
            {
                store = session.DefaultStore;
                root = store.GetRootFolder();
                child = FindChild(root, trimmed);
                if (child == null)
                {
                    throw new InvalidOperationException("No Outlook folder matches '" + segment + "'. Use outlook_list_folders to discover valid paths.");
                }
                var resolved = child;
                child = null;
                return resolved;
            }
            finally
            {
                ComUtilities.TryRelease(child);
                ComUtilities.TryRelease(root);
                ComUtilities.TryRelease(store);
            }
        }

        private static Outlook.MAPIFolder FindChild(Outlook.MAPIFolder folder, string name)
        {
            Outlook.Folders children = null;
            try
            {
                children = folder.Folders;
                for (var index = 1; index <= children.Count; index++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = children[index];
                        if (string.Equals(child.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            var resolved = child;
                            child = null;
                            return resolved;
                        }
                    }
                    finally
                    {
                        ComUtilities.TryRelease(child);
                    }
                }
                return null;
            }
            finally
            {
                ComUtilities.TryRelease(children);
            }
        }

        private Outlook.MAPIFolder GetDefaultFolderSafe(Outlook.NameSpace session, Outlook.OlDefaultFolders folder)
        {
            Outlook.MAPIFolder resolved = null;
            try
            {
                resolved = session.GetDefaultFolder(folder);
                var result = resolved;
                resolved = null;
                return result;
            }
            finally
            {
                ComUtilities.TryRelease(resolved);
            }
        }

        private Outlook.MailItem ResolveMailItem(Outlook.NameSpace session, string entryId)
        {
            Outlook.Store store = null;
            object raw = null;
            Outlook.MailItem mail = null;
            try
            {
                store = session.DefaultStore;
                raw = session.GetItemFromID(entryId, store.StoreID);
                mail = raw as Outlook.MailItem;
                if (mail == null)
                {
                    throw new InvalidOperationException("The item is not an email message. Only mail items are supported.");
                }
                var resolved = mail;
                mail = null;
                raw = null;
                return resolved;
            }
            catch (InvalidOperationException)
            {
                ComUtilities.TryRelease(mail);
                throw;
            }
            catch (Exception exception)
            {
                ComUtilities.TryRelease(mail);
                throw new InvalidOperationException("No Outlook item matches entry_id '" + entryId + "'. Refresh it with outlook_list_items.", exception);
            }
            finally
            {
                ComUtilities.TryRelease(raw);
                ComUtilities.TryRelease(store);
            }
        }

        private static void RequireDraft(Outlook.NameSpace session, Outlook.MailItem mail)
        {
            if (mail.Sent)
            {
                throw new InvalidOperationException("The item has already been sent and can no longer be modified through the draft tools.");
            }
            Outlook.MAPIFolder drafts = null;
            Outlook.MAPIFolder parent = null;
            try
            {
                parent = mail.Parent as Outlook.MAPIFolder;
                drafts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
                var parentPath = SafeText(() => parent?.FolderPath);
                var draftsPath = SafeText(() => drafts.FolderPath);
                if (!string.Equals(parentPath, draftsPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("Only drafts stored in the Drafts folder can be changed, deleted, or sent by Ribbon tools.");
                }
            }
            finally
            {
                ComUtilities.TryRelease(drafts);
                ComUtilities.TryRelease(parent);
            }
        }

        private static string JoinRecipients(IList<string> recipients)
        {
            var joined = new List<string>();
            foreach (var recipient in recipients ?? new List<string>())
            {
                if (!string.IsNullOrWhiteSpace(recipient)) joined.Add(recipient.Trim());
            }
            return string.Join("; ", joined);
        }

        private static bool BodyTooLong(string body)
        {
            return body != null && body.Length > HardMaximumDraftBodyCharacters;
        }

        private static string ImportanceText(Outlook.OlImportance importance)
        {
            if (importance == Outlook.OlImportance.olImportanceHigh) return "high";
            if (importance == Outlook.OlImportance.olImportanceLow) return "low";
            return "normal";
        }

        private static string ItemKind(string messageClass)
        {
            var value = (messageClass ?? string.Empty).ToLowerInvariant();
            if (value.StartsWith("ipm.note", StringComparison.Ordinal)) return "mail";
            if (value.StartsWith("ipm.appointment", StringComparison.Ordinal)) return "appointment";
            if (value.StartsWith("ipm.schedule.meeting", StringComparison.Ordinal)) return "meeting";
            if (value.StartsWith("ipm.task", StringComparison.Ordinal)) return "task";
            if (value.StartsWith("ipm.contact", StringComparison.Ordinal)) return "contact";
            if (value.StartsWith("ipm.post", StringComparison.Ordinal)) return "post";
            if (value.StartsWith("ipm.sticky", StringComparison.Ordinal)) return "note";
            if (value.Length == 0) return "item";
            return "other";
        }

        private static string SafeText(Func<string> reader)
        {
            try { return reader() ?? string.Empty; }
            catch { return string.Empty; }
        }

        private static int SafeCount(Func<int> reader)
        {
            try { return reader(); }
            catch { return -1; }
        }

        private static string SafeDate(Func<DateTime> reader)
        {
            try
            {
                return reader().ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }
    }
}
