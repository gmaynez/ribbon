using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal enum ConversationRestoreKind
    {
        Resumed,
        UnknownSession,
        Failed
    }

    internal sealed class ConversationRestoreResult
    {
        public ConversationRestoreKind Kind { get; private set; }
        public string ResumeKind { get; private set; }
        public string Error { get; private set; }

        public static ConversationRestoreResult Resumed(string resumeKind)
        {
            return new ConversationRestoreResult
            {
                Kind = ConversationRestoreKind.Resumed,
                ResumeKind = resumeKind
            };
        }

        public static ConversationRestoreResult UnknownSession()
        {
            return new ConversationRestoreResult { Kind = ConversationRestoreKind.UnknownSession };
        }

        public static ConversationRestoreResult Failed(string error)
        {
            return new ConversationRestoreResult
            {
                Kind = ConversationRestoreKind.Failed,
                Error = error
            };
        }
    }

    internal sealed class ConversationWorkspace
    {
        private readonly VstoHostRuntime _runtime;
        private readonly Action _scheduleSave;

        public ConversationWorkspace(VstoHostRuntime runtime, Action scheduleSave)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            _scheduleSave = scheduleSave ?? throw new ArgumentNullException(nameof(scheduleSave));
        }

        public ConversationRecord Current { get; private set; }
        public bool HistoryReadOnly { get; private set; }
        public bool HasCurrent => Current != null;

        public bool IsOpen(ConversationRecord record)
        {
            return Current != null && record != null
                && string.Equals(Current.Id, record.Id, StringComparison.OrdinalIgnoreCase);
        }

        public void Clear()
        {
            Current = null;
            HistoryReadOnly = false;
        }

        public void OpenReadOnly()
        {
            Current = null;
            HistoryReadOnly = true;
        }

        public void Adopt(ConversationRecord record)
        {
            Current = record;
            HistoryReadOnly = false;
        }

        public void Ensure(AgentSummary agent, string firstPrompt, SidebarSession session)
        {
            if (Current != null) return;
            Current = ConversationHistoryStorage.Create(_runtime.Registration, agent, TitleFromPrompt(firstPrompt));
            ApplySession(session, null);
        }

        public void ApplySession(SidebarSession session, string modelName)
        {
            if (Current == null || session == null) return;
            Current.AcpSessionId = session.SessionId ?? string.Empty;
            Current.AcpWorkingDirectory = session.WorkingDirectory ?? string.Empty;
            Current.SupportsLoad = session.SupportsLoad;
            Current.SupportsResume = session.SupportsResume;
            Current.SupportsList = session.SupportsList;
            if (!string.IsNullOrWhiteSpace(modelName)) Current.ModelName = modelName;
            _scheduleSave();
        }

        public void ApplyTitle(string title)
        {
            if (Current == null || string.IsNullOrWhiteSpace(title)) return;
            Current.Title = ConversationHistoryStorage.NormalizeTitle(title);
            _scheduleSave();
        }

        public void ApplyModelName(string modelName)
        {
            if (Current == null || string.IsNullOrWhiteSpace(modelName)) return;
            Current.ModelName = modelName;
            _scheduleSave();
        }

        public void Capture(ConversationTranscriptEntry entry)
        {
            if (Current == null || entry == null || string.IsNullOrEmpty(entry.Text)) return;
            Current.Entries.Add(entry);
            _scheduleSave();
        }

        public bool MatchesCurrentDocument()
        {
            return MatchesDocument(Current);
        }

        public bool MatchesDocument(ConversationRecord record)
        {
            return ConversationHistoryStorage.MatchesCurrentDocument(record, _runtime.Registration);
        }

        public ConversationRecord CreateFreshContinuation(ConversationRecord source, AgentSummary agent)
        {
            var continuation = ConversationHistoryStorage.Create(
                _runtime.Registration,
                agent,
                source.DisplayTitle + " · continued");
            continuation.ContinuedFromId = source.Id;
            continuation.Entries = (source.Entries ?? new List<ConversationTranscriptEntry>()).Select(entry => new ConversationTranscriptEntry
            {
                Text = entry.Text,
                Tone = entry.Tone,
                Style = entry.Style
            }).ToList();
            return continuation;
        }

        public void Persist()
        {
            if (Current == null) return;
            try
            {
                var registration = _runtime.Registration;
                if (ConversationHistoryStorage.MatchesCurrentDocument(Current, registration))
                {
                    ConversationHistoryStorage.UpdateDocumentBinding(Current, registration);
                }
                ConversationHistoryStorage.Save(Current);
            }
            catch
            {
            }
        }

        public async Task<ConversationRestoreResult> TryRestoreAsync(
            ConversationRecord record,
            AgentSummary agent,
            SidebarSession session,
            Action<string> status)
        {
            if (record.SupportsList)
            {
                try
                {
                    status?.Invoke("Checking saved agent session…");
                    var listed = await _runtime.ListAgentSessionsAsync(agent.Id, record.AcpWorkingDirectory);
                    if (listed.Supported)
                    {
                        var known = (listed.Sessions ?? new List<AgentSessionSummary>())
                            .FirstOrDefault(item => string.Equals(item.SessionId, record.AcpSessionId, StringComparison.Ordinal));
                        if (known == null && listed.Complete)
                        {
                            return ConversationRestoreResult.UnknownSession();
                        }

                        if (known != null && !string.IsNullOrWhiteSpace(known.Title))
                        {
                            record.Title = ConversationHistoryStorage.NormalizeTitle(known.Title);
                        }
                    }
                }
                catch
                {
                }
            }

            var resume = await session.ResumeAsync(agent, record, status);
            return resume.Resumed
                ? ConversationRestoreResult.Resumed(resume.ResumeKind)
                : ConversationRestoreResult.Failed(resume.Error);
        }

        private static string TitleFromPrompt(string prompt)
        {
            var title = string.Join(" ", (prompt ?? string.Empty)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)).Trim();
            if (title.Length > 68) title = title.Substring(0, 65) + "…";
            return string.IsNullOrWhiteSpace(title) ? "New Ribbon conversation" : title;
        }
    }
}
