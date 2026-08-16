using System;
using System.Drawing;

namespace Ribbon.Vsto
{
    public static class RibbonProductIdentity
    {
        public static string GetProductName(string hostKind)
        {
            if (string.Equals(hostKind, "Excel", StringComparison.OrdinalIgnoreCase)) return "Grid";
            if (string.Equals(hostKind, "Word", StringComparison.OrdinalIgnoreCase)) return "Quill";
            if (string.Equals(hostKind, "PowerPoint", StringComparison.OrdinalIgnoreCase)) return "Deck";
            if (string.Equals(hostKind, "Outlook", StringComparison.OrdinalIgnoreCase)) return "Post";
            return "Office";
        }

        public static string GetTaskPaneTitle(string hostKind)
        {
            var host = string.IsNullOrWhiteSpace(hostKind) ? "Office" : hostKind;
            return "Ribbon " + GetProductName(host) + " for " + host;
        }

        public static string GetMark(string hostKind)
        {
            var product = GetProductName(hostKind);
            return product.Length == 0 ? "R" : product.Substring(0, 1).ToUpperInvariant();
        }

        public static Color GetBrandColor(string hostKind)
        {
            if (string.Equals(hostKind, "Excel", StringComparison.OrdinalIgnoreCase)) return Color.FromArgb(33, 115, 70);
            if (string.Equals(hostKind, "Word", StringComparison.OrdinalIgnoreCase)) return Color.FromArgb(43, 87, 154);
            if (string.Equals(hostKind, "PowerPoint", StringComparison.OrdinalIgnoreCase)) return Color.FromArgb(210, 71, 38);
            if (string.Equals(hostKind, "Outlook", StringComparison.OrdinalIgnoreCase)) return Color.FromArgb(15, 108, 189);
            return Color.FromArgb(99, 102, 241);
        }

        public static string GetDocumentNoun(string hostKind)
        {
            if (string.Equals(hostKind, "Excel", StringComparison.OrdinalIgnoreCase)) return "workbook";
            if (string.Equals(hostKind, "PowerPoint", StringComparison.OrdinalIgnoreCase)) return "presentation";
            if (string.Equals(hostKind, "Outlook", StringComparison.OrdinalIgnoreCase)) return "mailbox";
            return "document";
        }

        public static string GetExamplePrompt(string hostKind)
        {
            if (string.Equals(hostKind, "Excel", StringComparison.OrdinalIgnoreCase))
                return "Try: “Summarize this sheet and create a chart.”";
            if (string.Equals(hostKind, "Word", StringComparison.OrdinalIgnoreCase))
                return "Try: “Polish this document and apply heading styles.”";
            if (string.Equals(hostKind, "PowerPoint", StringComparison.OrdinalIgnoreCase))
                return "Try: “Turn these notes into a clear slide.”";
            if (string.Equals(hostKind, "Outlook", StringComparison.OrdinalIgnoreCase))
                return "Try: “Summarize today's unread email and draft replies.”";
            return "Try asking the agent to inspect the open document first.";
        }
    }
}
