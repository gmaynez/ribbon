using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class PermissionDialog : Form
    {
        private readonly RibbonPalette _palette;
        private readonly CheckBox _remember;
        private string _selectedOptionId;
        private bool _cancelled = true;

        private PermissionDialog(
            RibbonPalette palette,
            string heading,
            string description,
            string details,
            IList<PermissionChoice> options,
            bool offerSessionMemory)
        {
            _palette = palette ?? RibbonPalette.Detect();
            _remember = new CheckBox
            {
                Text = "Don't ask again for this action during this agent session",
                AutoSize = true,
                ForeColor = _palette.Text,
                BackColor = _palette.Surface,
                Visible = offerSessionMemory,
                Margin = new Padding(0, 6, 0, 0)
            };

            Text = "Ribbon · Permission";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(560, offerSessionMemory ? 328 : 302);
            BackColor = _palette.Background;
            ForeColor = _palette.Text;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(18),
                BackColor = _palette.Background
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.Controls.Add(BuildHeader(heading, description), 0, 0);
            root.Controls.Add(BuildBody(details), 0, 1);
            root.Controls.Add(BuildActions(options), 0, 2);
            Controls.Add(root);
        }

        public static PermissionDecision ShowAcp(IWin32Window owner, PermissionPrompt prompt, RibbonPalette palette)
        {
            var options = (prompt?.Options ?? new List<PermissionChoice>()).ToList();
            var hasAlwaysChoice = options.Any(option => IsAlways(option?.Kind));
            using (var dialog = new PermissionDialog(
                palette,
                prompt?.Title ?? "Agent action",
                "The active ACP agent is requesting permission before it continues.",
                PermissionDetails(prompt?.RawJson),
                options,
                !hasAlwaysChoice && options.Any(option => IsAllow(option?.Kind))))
            {
                dialog.ShowDialog(owner);
                var selected = options.FirstOrDefault(option => string.Equals(option.OptionId, dialog._selectedOptionId, StringComparison.Ordinal));
                return new PermissionDecision
                {
                    Cancelled = dialog._cancelled,
                    OptionId = dialog._selectedOptionId ?? string.Empty,
                    RememberForSession = !dialog._cancelled && IsAllow(selected?.Kind) && dialog._remember.Checked
                };
            }
        }

        public static PermissionDecision ShowDestructiveTool(
            IWin32Window owner,
            string toolName,
            string argumentsJson,
            RibbonPalette palette)
        {
            var options = new List<PermissionChoice>
            {
                new PermissionChoice { OptionId = "allow", Name = "Allow", Kind = "allow_once" },
                new PermissionChoice { OptionId = "reject", Name = "Cancel", Kind = "reject_once" }
            };
            using (var dialog = new PermissionDialog(
                palette,
                FriendlyToolName(toolName),
                "This Office action changes the open document.",
                FormatArguments(argumentsJson),
                options,
                true))
            {
                dialog.ShowDialog(owner);
                return new PermissionDecision
                {
                    Cancelled = dialog._cancelled || dialog._selectedOptionId != "allow",
                    OptionId = dialog._selectedOptionId ?? string.Empty,
                    RememberForSession = !dialog._cancelled && dialog._selectedOptionId == "allow" && dialog._remember.Checked
                };
            }
        }

        private Control BuildHeader(string heading, string description)
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = _palette.Background };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 42));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(_palette, "!", _palette.Danger) { Margin = new Padding(0, 3, 10, 0) };
            var text = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = _palette.Background };
            text.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            text.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            text.Controls.Add(Label(heading, 12f, FontStyle.Bold, _palette.Text), 0, 0);
            text.Controls.Add(Label(description, 8.5f, FontStyle.Regular, _palette.MutedText), 0, 1);
            layout.Controls.Add(mark, 0, 0);
            layout.Controls.Add(text, 1, 0);
            return layout;
        }

        private Control BuildBody(string details)
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 8), Padding = new Padding(14, 12, 14, 10) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(Label("ACTION DETAILS", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);
            var detailText = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = ScrollBars.Vertical,
                BackColor = _palette.Surface,
                ForeColor = _palette.Text,
                Text = string.IsNullOrWhiteSpace(details) ? "No additional details were provided." : details,
                AccessibleName = "Permission action details"
            };
            layout.Controls.Add(detailText, 0, 1);
            layout.Controls.Add(_remember, 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildActions(IList<PermissionChoice> options)
        {
            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = _palette.Background,
                Padding = new Padding(0, 6, 0, 0)
            };

            foreach (var option in options ?? new List<PermissionChoice>())
            {
                var choice = option;
                var allow = IsAllow(choice?.Kind);
                var button = new RibbonButton(_palette, allow ? RibbonButtonKind.Primary : RibbonButtonKind.Secondary)
                {
                    Text = string.IsNullOrWhiteSpace(choice?.Name) ? (allow ? "Allow" : "Reject") : choice.Name,
                    Width = 112,
                    Margin = new Padding(8, 0, 0, 0),
                    AccessibleName = string.IsNullOrWhiteSpace(choice?.Name) ? "Permission choice" : choice.Name
                };
                button.Click += (sender, args) =>
                {
                    _selectedOptionId = choice?.OptionId ?? string.Empty;
                    _cancelled = false;
                    DialogResult = allow ? DialogResult.OK : DialogResult.Cancel;
                    Close();
                };
                actions.Controls.Add(button);
            }

            if (actions.Controls.Count == 0)
            {
                var close = new RibbonButton(_palette) { Text = "Close", Width = 92 };
                close.Click += (sender, args) => Close();
                actions.Controls.Add(close);
            }
            return actions;
        }

        private Label Label(string text, float size, FontStyle style, Color color)
        {
            return new Label
            {
                Text = text ?? string.Empty,
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = color,
                BackColor = _palette.Background,
                Font = new Font(Font.FontFamily, size, style),
                Margin = new Padding(0)
            };
        }

        private static bool IsAllow(string kind)
        {
            return !string.IsNullOrWhiteSpace(kind) && kind.StartsWith("allow", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsAlways(string kind)
        {
            return !string.IsNullOrWhiteSpace(kind) && kind.EndsWith("always", StringComparison.OrdinalIgnoreCase);
        }

        private static string FriendlyToolName(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName)) return "Change the Office document";
            var words = toolName.Replace("excel_", string.Empty).Replace("word_", string.Empty).Replace("powerpoint_", string.Empty).Replace('_', ' ');
            return char.ToUpperInvariant(words[0]) + words.Substring(1);
        }

        private static string PermissionDetails(string rawJson)
        {
            return string.IsNullOrWhiteSpace(rawJson) ? string.Empty : rawJson;
        }

        private static string FormatArguments(string argumentsJson)
        {
            return string.IsNullOrWhiteSpace(argumentsJson) || argumentsJson == "{}"
                ? "The action does not require additional parameters."
                : argumentsJson;
        }
    }
}
