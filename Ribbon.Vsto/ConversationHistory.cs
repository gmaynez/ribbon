using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    public sealed class ConversationRecord
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string AgentId { get; set; }
        public string AgentName { get; set; }
        public string ModelName { get; set; }
        public string AcpSessionId { get; set; }
        public string AcpWorkingDirectory { get; set; }
        public bool SupportsLoad { get; set; }
        public bool SupportsResume { get; set; }
        public bool SupportsList { get; set; }
        public string HostId { get; set; }
        public string HostKind { get; set; }
        public string DocumentId { get; set; }
        public string DocumentKey { get; set; }
        public string DocumentName { get; set; }
        public string DocumentPath { get; set; }
        public string ContinuedFromId { get; set; }
        public DateTime CreatedAtUtc { get; set; }
        public DateTime UpdatedAtUtc { get; set; }
        public IList<ConversationTranscriptEntry> Entries { get; set; }

        public string DisplayTitle => string.IsNullOrWhiteSpace(Title) ? "Untitled conversation" : Title;
        public string DisplayDocument => string.IsNullOrWhiteSpace(DocumentName) ? HostKind : DocumentName;
        public bool MayResumeNatively => !string.IsNullOrWhiteSpace(AcpSessionId)
            && !string.IsNullOrWhiteSpace(AcpWorkingDirectory)
            && (SupportsResume || SupportsLoad);
    }

    public sealed class ConversationTranscriptEntry
    {
        public string Text { get; set; }
        public string Tone { get; set; }
        public string Style { get; set; }
    }

    public static class ConversationHistoryStorage
    {
        private const int MaximumConversations = 200;
        private static readonly object Gate = new object();
        internal static string StorageRootOverride { get; set; }

        public static ConversationRecord Create(HostRegistration registration, AgentSummary agent, string title)
        {
            if (registration == null) throw new ArgumentNullException(nameof(registration));
            if (agent == null || string.IsNullOrWhiteSpace(agent.Id)) throw new ArgumentException("An ACP agent is required.", nameof(agent));
            var now = DateTime.UtcNow;
            var record = new ConversationRecord
            {
                Id = Guid.NewGuid().ToString("N"),
                Title = NormalizeTitle(title),
                AgentId = agent.Id,
                AgentName = agent.Name ?? agent.Id,
                CreatedAtUtc = now,
                UpdatedAtUtc = now,
                Entries = new List<ConversationTranscriptEntry>()
            };
            UpdateDocumentBinding(record, registration);
            return record;
        }

        public static void UpdateDocumentBinding(ConversationRecord record, HostRegistration registration)
        {
            if (record == null || registration == null) return;
            var path = registration.DocumentPath ?? string.Empty;
            record.HostId = registration.HostId ?? record.HostId ?? string.Empty;
            record.HostKind = registration.HostKind ?? record.HostKind ?? string.Empty;
            record.DocumentId = registration.ContextId ?? registration.DocumentId ?? record.DocumentId ?? string.Empty;
            record.DocumentPath = path;
            record.DocumentName = DocumentName(path, registration.ContextName, record.HostKind);
            record.DocumentKey = GetDocumentKey(registration);
        }

        public static bool MatchesCurrentDocument(ConversationRecord record, HostRegistration registration)
        {
            if (record == null || registration == null) return false;
            var contextId = registration.ContextId ?? registration.DocumentId;
            if (string.Equals(record.HostId, registration.HostId, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(record.DocumentId)
                && string.Equals(record.DocumentId, contextId, StringComparison.Ordinal))
            {
                return true;
            }
            return string.Equals(record.DocumentKey, GetDocumentKey(registration), StringComparison.OrdinalIgnoreCase);
        }

        public static string NormalizeTitle(string title)
        {
            var value = string.Join(" ", (title ?? string.Empty)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)).Trim();
            return value.Length > 160 ? value.Substring(0, 157) + "…" : value;
        }

        public static string GetDocumentKey(HostRegistration registration)
        {
            if (registration == null) return string.Empty;
            var stablePath = StableDocumentPath(registration.DocumentPath);
            if (!string.IsNullOrWhiteSpace(stablePath))
            {
                return (registration.HostKind ?? string.Empty) + "|path|" + stablePath;
            }
            var contextId = registration.ContextId ?? registration.DocumentId ?? registration.HostId ?? string.Empty;
            var contextKind = string.IsNullOrWhiteSpace(registration.ContextKind) ? "document" : registration.ContextKind;
            return (registration.HostKind ?? string.Empty) + "|" + contextKind + "|" + contextId;
        }

        public static IList<ConversationRecord> LoadAll()
        {
            lock (Gate)
            {
                var directory = GetDirectory();
                if (!Directory.Exists(directory)) return new List<ConversationRecord>();
                var serializer = Serializer();
                var records = new List<ConversationRecord>();
                foreach (var file in Directory.GetFiles(directory, "*.json"))
                {
                    try
                    {
                        var record = serializer.Deserialize<ConversationRecord>(File.ReadAllText(file, Encoding.UTF8));
                        if (record == null || string.IsNullOrWhiteSpace(record.Id)) continue;
                        record.Entries = record.Entries ?? new List<ConversationTranscriptEntry>();
                        records.Add(record);
                    }
                    catch
                    {
                        // A damaged history entry must not prevent Office or the remaining history from opening.
                    }
                }
                return records.OrderByDescending(item => item.UpdatedAtUtc).ToList();
            }
        }

        public static void Save(ConversationRecord record)
        {
            if (record == null || string.IsNullOrWhiteSpace(record.Id)) return;
            lock (Gate)
            {
                var directory = GetDirectory();
                Directory.CreateDirectory(directory);
                record.UpdatedAtUtc = DateTime.UtcNow;
                record.Entries = record.Entries ?? new List<ConversationTranscriptEntry>();
                var path = RecordPath(record.Id);
                var temporary = path + ".tmp";
                File.WriteAllText(temporary, Serializer().Serialize(record), new UTF8Encoding(false));
                if (File.Exists(path))
                {
                    try { File.Replace(temporary, path, null); }
                    catch
                    {
                        File.Delete(path);
                        File.Move(temporary, path);
                    }
                }
                else
                {
                    File.Move(temporary, path);
                }
                Prune(directory);
            }
        }

        public static void Delete(ConversationRecord record)
        {
            if (record == null || string.IsNullOrWhiteSpace(record.Id)) return;
            lock (Gate)
            {
                try
                {
                    var path = RecordPath(record.Id);
                    if (File.Exists(path)) File.Delete(path);
                }
                catch
                {
                }
            }
        }

        private static string StableDocumentPath(string documentPath)
        {
            if (string.IsNullOrWhiteSpace(documentPath)) return string.Empty;
            try
            {
                if (!Path.IsPathRooted(documentPath) || !File.Exists(documentPath)) return string.Empty;
                return Path.GetFullPath(documentPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string DocumentName(string path, string contextName, string hostKind)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path)) return Path.GetFileName(path);
            }
            catch
            {
            }
            if (!string.IsNullOrWhiteSpace(contextName)) return contextName;
            return string.IsNullOrWhiteSpace(hostKind) ? "Office document" : "Unsaved " + hostKind + " document";
        }

        private static JavaScriptSerializer Serializer()
        {
            return new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024, RecursionLimit = 128 };
        }

        private static string RecordPath(string id)
        {
            var safeId = string.Concat((id ?? string.Empty).Where(character => char.IsLetterOrDigit(character) || character == '-' || character == '_'));
            if (string.IsNullOrWhiteSpace(safeId)) throw new InvalidOperationException("The conversation id is invalid.");
            var root = Path.GetFullPath(GetDirectory()) + Path.DirectorySeparatorChar;
            var path = Path.GetFullPath(Path.Combine(GetDirectory(), safeId + ".json"));
            if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The conversation path is outside Ribbon's history directory.");
            }
            return path;
        }

        private static void Prune(string directory)
        {
            try
            {
                foreach (var file in new DirectoryInfo(directory).GetFiles("*.json")
                    .OrderByDescending(item => item.LastWriteTimeUtc)
                    .Skip(MaximumConversations))
                {
                    file.Delete();
                }
            }
            catch
            {
            }
        }

        private static string GetDirectory()
        {
            if (!string.IsNullOrWhiteSpace(StorageRootOverride)) return Path.GetFullPath(StorageRootOverride);
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Ribbon",
                "Conversations");
        }
    }
}
