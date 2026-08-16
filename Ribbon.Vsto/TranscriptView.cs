using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class TranscriptView : IDisposable
    {
        private readonly RichTextBox _box;
        private readonly RibbonPalette _palette;
        private readonly string _hostKind;
        private readonly Dictionary<string, string> _toolStatuses = new Dictionary<string, string>(StringComparer.Ordinal);
        private Font _regular;
        private Font _bold;
        private Font _italic;
        private bool _suppressCapture;
        private bool _hasConversation;
        private bool _activityVisible;
        private bool _thoughtVisible;
        private bool _responseVisible;
        private string _planSignature;

        public TranscriptView(RichTextBox box, RibbonPalette palette, string hostKind)
        {
            _box = box ?? throw new ArgumentNullException(nameof(box));
            _palette = palette ?? throw new ArgumentNullException(nameof(palette));
            _hostKind = string.IsNullOrWhiteSpace(hostKind) ? "Office" : hostKind;
        }

        public event Action<ConversationTranscriptEntry> EntryCaptured;

        public void CreateFonts(Font family)
        {
            _regular = new Font(family.FontFamily, 9.5f, FontStyle.Regular);
            _bold = new Font(family.FontFamily, 9.25f, FontStyle.Bold);
            _italic = new Font(family.FontFamily, 9.25f, FontStyle.Italic);
            _box.Font = _regular;
        }

        public void ShowWelcome()
        {
            _box.Clear();
            Append("Ready to work in " + _hostKind + ".\n", _palette.Text, FontStyle.Bold);
            Append(
                "Ask an agent to inspect, explain, or update the open " + RibbonProductIdentity.GetDocumentNoun(_hostKind) + ".\n\n",
                _palette.MutedText,
                FontStyle.Regular);
            Append(RibbonProductIdentity.GetExamplePrompt(_hostKind), _palette.Accent, FontStyle.Regular);
            _hasConversation = false;
        }

        public void Render(ConversationRecord record)
        {
            _suppressCapture = true;
            try
            {
                _box.Clear();
                foreach (var entry in record?.Entries ?? new List<ConversationTranscriptEntry>())
                {
                    Append(entry.Text ?? string.Empty, ColorForTone(entry.Tone), FontStyleFor(entry.Style));
                }
                _hasConversation = record != null && record.Entries != null && record.Entries.Count > 0;
            }
            finally
            {
                _suppressCapture = false;
            }
        }

        public void BeginTurn(string text, string agentName)
        {
            if (!_hasConversation)
            {
                _box.Clear();
                _hasConversation = true;
            }
            else
            {
                Append(Environment.NewLine + Environment.NewLine, _palette.Text, FontStyle.Regular);
            }

            Append("You\n", _palette.MutedText, FontStyle.Bold);
            Append(text + Environment.NewLine + Environment.NewLine, _palette.Text, FontStyle.Regular);
            Append(agentName + "\n", _palette.Accent, FontStyle.Bold);
            _toolStatuses.Clear();
            _activityVisible = false;
            _thoughtVisible = false;
            _responseVisible = false;
            _planSignature = null;
        }

        public void EndTurn()
        {
            Append(Environment.NewLine + Environment.NewLine, _palette.Text, FontStyle.Regular);
        }

        public bool TryApply(SessionUpdateMessage update)
        {
            if (update.UpdateKind == "agent_message_chunk" && !string.IsNullOrEmpty(update.Text))
            {
                if (_activityVisible && !_responseVisible)
                {
                    Append(Environment.NewLine + "Response\n", _palette.Accent, FontStyle.Bold);
                    _responseVisible = true;
                }
                Append(update.Text, _palette.Text, FontStyle.Regular);
                return true;
            }

            if (update.UpdateKind == "agent_thought_chunk" && !string.IsNullOrEmpty(update.Text))
            {
                if (!_thoughtVisible)
                {
                    Append("Thinking\n", _palette.MutedText, FontStyle.Bold);
                    _thoughtVisible = true;
                    _activityVisible = true;
                }
                Append(update.Text, _palette.MutedText, FontStyle.Italic);
                return true;
            }

            if (update.UpdateKind == "tool_call" || update.UpdateKind == "tool_call_update")
            {
                AppendToolActivity(update);
                return true;
            }

            if (update.UpdateKind == "plan")
            {
                AppendPlan(update.PlanEntries);
                return true;
            }

            return false;
        }

        public void AppendAutoApproval(string label)
        {
            Append(Environment.NewLine + "Auto-approved · " + label + "\n", _palette.Accent, FontStyle.Italic);
        }

        public void AppendCheckpointRestored(string displayName)
        {
            Append(Environment.NewLine + Environment.NewLine + "Checkpoint restored\n", _palette.Success, FontStyle.Bold);
            Append("The document is back at " + displayName + ". A fresh agent session will be used for the next turn.\n", _palette.MutedText, FontStyle.Regular);
        }

        public void AppendFreshContinuationNotice()
        {
            Append(Environment.NewLine + Environment.NewLine + "Fresh agent context\n", _palette.Accent, FontStyle.Bold);
            Append("The saved transcript is shown above, but this agent session cannot see its earlier messages. Include any needed context in your next prompt.\n",
                _palette.MutedText, FontStyle.Regular);
        }

        public static string ToolStatusText(SessionUpdateMessage update)
        {
            var status = FriendlyToolStatus(update.Status);
            return string.IsNullOrWhiteSpace(update.ToolName)
                ? "Using Office tools…"
                : update.ToolName + (string.IsNullOrWhiteSpace(status) ? string.Empty : " · " + status);
        }

        public static bool IsFailedTool(SessionUpdateMessage update)
        {
            return string.Equals(update.Status, "failed", StringComparison.OrdinalIgnoreCase);
        }

        public void Dispose()
        {
            _bold?.Dispose();
            _regular?.Dispose();
            _italic?.Dispose();
        }

        private void AppendPlan(IList<SessionPlanEntry> entries)
        {
            if (entries == null || entries.Count == 0) return;
            var signature = string.Join("|", entries.Select(entry =>
                (entry.Content ?? string.Empty) + ":" + (entry.Status ?? string.Empty)));
            if (string.Equals(signature, _planSignature, StringComparison.Ordinal)) return;
            _planSignature = signature;
            Append((_activityVisible ? Environment.NewLine : string.Empty) + "Plan\n", _palette.MutedText, FontStyle.Bold);
            foreach (var entry in entries)
            {
                var marker = string.Equals(entry.Status, "completed", StringComparison.OrdinalIgnoreCase) ? "✓"
                    : string.Equals(entry.Status, "in_progress", StringComparison.OrdinalIgnoreCase) ? "→"
                    : "•";
                var color = marker == "✓" ? _palette.Success : marker == "→" ? _palette.Accent : _palette.MutedText;
                Append(marker + " " + (entry.Content ?? string.Empty) + "\n", color, FontStyle.Regular);
            }
            _activityVisible = true;
        }

        private void AppendToolActivity(SessionUpdateMessage update)
        {
            var key = string.IsNullOrWhiteSpace(update.ToolCallId) ? update.ToolName ?? Guid.NewGuid().ToString("N") : update.ToolCallId;
            _toolStatuses.TryGetValue(key, out var previousStatus);
            var status = update.Status ?? string.Empty;
            var first = !_toolStatuses.ContainsKey(key);
            if (first && !_activityVisible)
            {
                Append("Activity\n", _palette.MutedText, FontStyle.Bold);
            }

            if (first)
            {
                Append("• " + (update.ToolName ?? "Office action") + "\n", _palette.Text, FontStyle.Regular);
                _activityVisible = true;
            }

            if (!string.Equals(previousStatus, status, StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(status, "completed", StringComparison.OrdinalIgnoreCase))
                {
                    Append("  ✓ Completed\n", _palette.Success, FontStyle.Regular);
                }
                else if (string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase))
                {
                    Append("  × Failed\n", _palette.Danger, FontStyle.Regular);
                }
                else if (!first && string.Equals(status, "in_progress", StringComparison.OrdinalIgnoreCase))
                {
                    Append("  → In progress\n", _palette.Accent, FontStyle.Regular);
                }
                _toolStatuses[key] = status;
            }

            if (!string.IsNullOrWhiteSpace(update.Text))
            {
                Append("  " + update.Text.Trim() + "\n", _palette.MutedText, FontStyle.Regular);
            }
        }

        private void Append(string text, Color color, FontStyle style)
        {
            if (!_suppressCapture && !string.IsNullOrEmpty(text))
            {
                var captured = EntryCaptured;
                if (captured != null)
                {
                    captured(new ConversationTranscriptEntry
                    {
                        Text = text,
                        Tone = ToneFor(color),
                        Style = StyleFor(style)
                    });
                }
            }

            _box.SelectionStart = _box.TextLength;
            _box.SelectionLength = 0;
            _box.SelectionColor = color;
            _box.SelectionFont = style == FontStyle.Bold
                ? _bold
                : style == FontStyle.Italic ? _italic : _regular;
            _box.AppendText(text);
            _box.SelectionStart = _box.TextLength;
            _box.ScrollToCaret();
        }

        private string ToneFor(Color color)
        {
            if (color.ToArgb() == _palette.Accent.ToArgb()) return "accent";
            if (color.ToArgb() == _palette.Success.ToArgb()) return "success";
            if (color.ToArgb() == _palette.Danger.ToArgb()) return "danger";
            if (color.ToArgb() == _palette.MutedText.ToArgb()) return "muted";
            return "text";
        }

        private Color ColorForTone(string tone)
        {
            switch (tone)
            {
                case "accent": return _palette.Accent;
                case "success": return _palette.Success;
                case "danger": return _palette.Danger;
                case "muted": return _palette.MutedText;
                default: return _palette.Text;
            }
        }

        private static string StyleFor(FontStyle style)
        {
            return style == FontStyle.Bold ? "bold" : style == FontStyle.Italic ? "italic" : "regular";
        }

        private static FontStyle FontStyleFor(string style)
        {
            return string.Equals(style, "bold", StringComparison.OrdinalIgnoreCase) ? FontStyle.Bold
                : string.Equals(style, "italic", StringComparison.OrdinalIgnoreCase) ? FontStyle.Italic
                : FontStyle.Regular;
        }

        private static string FriendlyToolStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status)) return string.Empty;
            return status.Replace('_', ' ');
        }
    }
}
