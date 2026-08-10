using System;

namespace Ribbon.Vsto
{
    internal enum ApprovalMode
    {
        // Prompt before each destructive Office action and ACP permission request (default).
        Ask,

        // Auto-approve destructive Office tools and ACP permission requests for the active
        // agent session. The mode resets to Ask when the session changes or Office restarts,
        // and every auto-approved action is surfaced to the task pane for auditability.
        Auto
    }

    internal sealed class AutoApprovalRecord : EventArgs
    {
        public AutoApprovalRecord(string category, string action, string argumentsJson)
        {
            Category = category ?? string.Empty;
            Action = action ?? string.Empty;
            ArgumentsJson = argumentsJson ?? string.Empty;
        }

        // "office" for destructive Office MCP tools; "acp" for relayed ACP permission requests.
        public string Category { get; }

        // Tool name for Office actions, or the prompt title for ACP requests.
        public string Action { get; }

        // The raw arguments/payload that were auto-approved, for the audit line.
        public string ArgumentsJson { get; }
    }
}
