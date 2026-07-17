using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class PermissionDialog : Form
    {
        private readonly RibbonPalette _palette;
        private readonly CheckBox _remember;
        private readonly int _collapsedClientHeight;
        private TextBox _technicalDetails;
        private LinkLabel _technicalToggle;
        private RowStyle _technicalDetailsRow;
        private bool _technicalDetailsExpanded;
        private string _selectedOptionId;
        private bool _cancelled = true;

        private PermissionDialog(
            RibbonPalette palette,
            string heading,
            string description,
            string details,
            IList<PermissionChoice> options,
            bool offerSessionMemory,
            OfficePermissionContent officeContent = null)
        {
            _palette = palette ?? RibbonPalette.Detect();
            _collapsedClientHeight = officeContent == null ? (offerSessionMemory ? 328 : 302) : 338;
            _remember = new CheckBox
            {
                Text = officeContent == null
                    ? "Don't ask again for this action during this agent session"
                    : "Allow this type of document edit for the rest of this agent session",
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
            ClientSize = new Size(560, _collapsedClientHeight);
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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.Controls.Add(BuildHeader(heading, description), 0, 0);
            root.Controls.Add(officeContent == null ? BuildAcpBody(details) : BuildOfficeBody(officeContent), 0, 1);
            root.Controls.Add(BuildActions(options), 0, 2);
            Controls.Add(root);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            RibbonWindowChrome.Apply(this, _palette);
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
            var content = BuildOfficePermissionContent(toolName, argumentsJson);
            var options = new List<PermissionChoice>
            {
                new PermissionChoice { OptionId = "allow", Name = content.ConfirmationLabel, Kind = "allow_once" },
                new PermissionChoice { OptionId = "reject", Name = "Cancel", Kind = "reject_once" }
            };
            using (var dialog = new PermissionDialog(
                palette,
                content.Heading,
                "Review this document change before it runs.",
                content.TechnicalDetails,
                options,
                true,
                content))
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
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(_palette, "!", _palette.Danger) { Margin = new Padding(0, 3, 10, 0) };
            var text = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = _palette.Background };
            text.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            text.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            text.Controls.Add(Label(heading, 12f, FontStyle.Bold, _palette.Text, _palette.Background), 0, 0);
            text.Controls.Add(Label(description, 8.5f, FontStyle.Regular, _palette.MutedText, _palette.Background), 0, 1);
            layout.Controls.Add(mark, 0, 0);
            layout.Controls.Add(text, 1, 0);
            return layout;
        }

        private Control BuildAcpBody(string details)
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 8), Padding = new Padding(14, 12, 14, 10) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(Label("ACTION DETAILS", 7.5f, FontStyle.Bold, _palette.MutedText, _palette.Surface), 0, 0);
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
            RibbonNativeTheme.ApplyDarkScrollBars(detailText, _palette);
            layout.Controls.Add(detailText, 0, 1);
            layout.Controls.Add(_remember, 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildOfficeBody(OfficePermissionContent content)
        {
            var surface = new RibbonSurface(_palette)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 8, 0, 8),
                Padding = new Padding(14, 10, 14, 9)
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                BackColor = _palette.Surface
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 27));
            _technicalDetailsRow = new RowStyle(SizeType.Absolute, 0);
            layout.RowStyles.Add(_technicalDetailsRow);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var summary = new Label
            {
                Text = content.Summary,
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                ForeColor = _palette.Text,
                BackColor = _palette.Surface,
                Font = new Font(Font.FontFamily, 9.25f, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 4),
                AccessibleName = "Proposed document change"
            };
            layout.Controls.Add(summary, 0, 0);

            var checkpoint = new Label
            {
                Text = "Checkpoint saved before this turn. You can restore it from the task pane.",
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                ForeColor = _palette.Success,
                BackColor = _palette.Surface,
                Font = new Font(Font.FontFamily, 8.4f, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0),
                AccessibleName = "Checkpoint available"
            };
            layout.Controls.Add(checkpoint, 0, 1);

            _technicalToggle = new LinkLabel
            {
                Text = "+ Show technical details",
                Dock = DockStyle.Fill,
                AutoSize = false,
                LinkColor = _palette.Accent,
                ActiveLinkColor = _palette.AccentHover,
                VisitedLinkColor = _palette.Accent,
                BackColor = _palette.Surface,
                Font = new Font(Font.FontFamily, 8.25f, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleLeft,
                LinkBehavior = LinkBehavior.HoverUnderline,
                Margin = new Padding(0),
                AccessibleName = "Show technical details"
            };
            _technicalToggle.LinkClicked += (sender, args) => ToggleTechnicalDetails();
            layout.Controls.Add(_technicalToggle, 0, 2);

            _technicalDetails = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = ScrollBars.Vertical,
                BackColor = _palette.SurfaceRaised,
                ForeColor = _palette.Text,
                Text = string.IsNullOrWhiteSpace(content.TechnicalDetails)
                    ? "The action does not require additional parameters."
                    : content.TechnicalDetails,
                Visible = false,
                AccessibleName = "Technical action details",
                Margin = new Padding(0, 2, 0, 6)
            };
            RibbonNativeTheme.ApplyDarkScrollBars(_technicalDetails, _palette);
            layout.Controls.Add(_technicalDetails, 0, 3);
            layout.Controls.Add(_remember, 0, 4);
            surface.Controls.Add(layout);
            return surface;
        }

        private void ToggleTechnicalDetails()
        {
            _technicalDetailsExpanded = !_technicalDetailsExpanded;
            _technicalDetails.Visible = _technicalDetailsExpanded;
            _technicalDetailsRow.Height = _technicalDetailsExpanded ? 118 : 0;
            _technicalToggle.Text = _technicalDetailsExpanded ? "- Hide technical details" : "+ Show technical details";
            _technicalToggle.AccessibleName = _technicalDetailsExpanded ? "Hide technical details" : "Show technical details";
            ClientSize = new Size(ClientSize.Width, _collapsedClientHeight + (_technicalDetailsExpanded ? 118 : 0));
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
                    Width = Math.Min(180, Math.Max(112, TextRenderer.MeasureText(
                        string.IsNullOrWhiteSpace(choice?.Name) ? (allow ? "Allow" : "Reject") : choice.Name,
                        Font).Width + 30)),
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

        private Label Label(string text, float size, FontStyle style, Color color, Color background)
        {
            return new Label
            {
                Text = text ?? string.Empty,
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = color,
                BackColor = background,
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

        private static OfficePermissionContent BuildOfficePermissionContent(string toolName, string argumentsJson)
        {
            var heading = FriendlyToolName(toolName);
            var arguments = ParseArguments(argumentsJson);
            var summary = BuildFriendlySummary(toolName, arguments, heading);
            return new OfficePermissionContent
            {
                Heading = heading,
                ConfirmationLabel = heading,
                Summary = summary,
                TechnicalDetails = PrettyPrintJson(argumentsJson)
            };
        }

        private static string BuildFriendlySummary(string toolName, IDictionary<string, object> arguments, string heading)
        {
            var name = toolName ?? string.Empty;
            var sheet = StringValue(arguments, "sheet_name");
            var address = StringValue(arguments, "address");
            var target = ExcelTarget(sheet, address);

            switch (name.ToLowerInvariant())
            {
                case "excel_write_range":
                    return "Write " + MatrixSize(arguments, "values") + " of values to " + target + ".";
                case "excel_write_formulas":
                    return "Write " + MatrixSize(arguments, "formulas") + " of formulas to " + target + ".";
                case "excel_clear_range":
                    return "Clear " + target + FormatMode(arguments, "mode") + ".";
                case "excel_format_range":
                    return "Apply formatting to " + target + ". Existing values stay in place.";
                case "excel_add_sheet":
                    return "Add a worksheet named " + Quoted(StringValue(arguments, "name"), "a new worksheet") + ".";
                case "excel_create_table":
                    return "Create " + Quoted(StringValue(arguments, "table_name"), "a table") + " in " + target
                        + FormatStyle(arguments) + ".";
                case "excel_create_chart":
                    return "Create " + FriendlyEnum(StringValue(arguments, "chart_type"), "a chart") + " from "
                        + ExcelTarget(sheet, StringValue(arguments, "source_range")) + FormatTitle(arguments) + ".";
                case "word_replace_selection":
                    return "Replace the current selection with " + TextDescription(arguments, "text") + ".";
                case "word_append_text":
                    return "Append " + TextDescription(arguments, "text") + " to the end of the document.";
                case "word_insert_text":
                    return "Insert " + TextDescription(arguments, "text") + " at "
                        + FriendlyEnum(StringValue(arguments, "position"), "the current selection") + ".";
                case "word_replace_range":
                    return "Replace characters " + RangeDescription(arguments) + " with " + TextDescription(arguments, "text") + ".";
                case "word_find_replace":
                case "powerpoint_find_replace":
                    return "Replace " + Quoted(StringValue(arguments, "find_text"), "matching text") + " with "
                        + Quoted(StringValue(arguments, "replace_text"), "new text") + FindReplaceScope(arguments, name) + ".";
                case "word_format_range":
                    return "Apply formatting to " + RangeDescription(arguments) + ". Existing text stays in place.";
                case "word_insert_heading":
                    return "Insert " + Quoted(StringValue(arguments, "text"), "a heading") + " as Heading "
                        + StringValue(arguments, "level", "1") + ".";
                case "word_insert_list":
                    return "Insert a " + FriendlyEnum(StringValue(arguments, "list_type"), "list") + " with "
                        + CollectionCount(arguments, "items", "item", "items") + ".";
                case "word_insert_table":
                    return "Insert a table with " + MatrixSize(arguments, "values") + ".";
                case "word_add_comment":
                    return "Add a review comment to " + RangeDescription(arguments) + ".";
                case "word_insert_page_break":
                    return "Insert a page break at " + FriendlyEnum(StringValue(arguments, "position"), "the current selection") + ".";
                case "powerpoint_add_slide":
                    return "Add a " + FriendlyEnum(StringValue(arguments, "layout"), "title and content") + " slide"
                        + AtPosition(arguments) + FormatTitle(arguments) + ".";
                case "powerpoint_delete_slide":
                    return "Delete slide " + StringValue(arguments, "slide_number", "?") + ".";
                case "powerpoint_duplicate_slide":
                    return "Duplicate slide " + StringValue(arguments, "slide_number", "?") + ".";
                case "powerpoint_move_slide":
                    return "Move slide " + StringValue(arguments, "slide_number", "?") + " to position "
                        + StringValue(arguments, "position", "?") + ".";
                case "powerpoint_set_slide_title":
                    return "Set the title of slide " + StringValue(arguments, "slide_number", "?") + " to "
                        + Quoted(StringValue(arguments, "title"), "new text") + ".";
                case "powerpoint_delete_shape":
                    return "Delete " + Quoted(StringValue(arguments, "shape_name"), "a shape") + " from slide "
                        + StringValue(arguments, "slide_number", "?") + ".";
                case "powerpoint_format_shape":
                    return "Update " + Quoted(StringValue(arguments, "shape_name"), "a shape") + " on slide "
                        + StringValue(arguments, "slide_number", "?") + ".";
                case "powerpoint_set_speaker_notes":
                    return "Replace the speaker notes on slide " + StringValue(arguments, "slide_number", "?") + ".";
                case "powerpoint_set_slide_background":
                    return "Set slide " + StringValue(arguments, "slide_number", "?") + " background to "
                        + StringValue(arguments, "color", "a new color") + ".";
                default:
                    var slide = StringValue(arguments, "slide_number");
                    return heading + (string.IsNullOrWhiteSpace(slide) ? " in the open document." : " on slide " + slide + ".");
            }
        }

        private static IDictionary<string, object> ParseArguments(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var value = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(json);
                var source = value as IDictionary<string, object>;
                return source == null
                    ? new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    : new Dictionary<string, object>(source, StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static string StringValue(IDictionary<string, object> values, string key, string fallback = "")
        {
            object value;
            return values != null && values.TryGetValue(key, out value) && value != null
                ? Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
                : fallback;
        }

        private static string ExcelTarget(string sheet, string address)
        {
            if (!string.IsNullOrWhiteSpace(sheet) && !string.IsNullOrWhiteSpace(address)) return sheet + " · " + address;
            if (!string.IsNullOrWhiteSpace(address)) return "range " + address;
            if (!string.IsNullOrWhiteSpace(sheet)) return "worksheet " + sheet;
            return "the active worksheet";
        }

        private static string Quoted(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : "\"" + Truncate(value, 80) + "\"";
        }

        private static string FriendlyEnum(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : value.Replace('_', ' ');
        }

        private static string MatrixSize(IDictionary<string, object> values, string key)
        {
            object raw;
            if (values == null || !values.TryGetValue(key, out raw)) return "a block";
            var rows = raw as ICollection;
            if (rows == null || rows.Count == 0) return "a block";
            var first = rows.Cast<object>().FirstOrDefault() as ICollection;
            return rows.Count + " row" + (rows.Count == 1 ? string.Empty : "s") + (first == null ? string.Empty : " × " + first.Count + " column" + (first.Count == 1 ? string.Empty : "s"));
        }

        private static string CollectionCount(IDictionary<string, object> values, string key, string singular, string plural)
        {
            object raw;
            var collection = values != null && values.TryGetValue(key, out raw) ? raw as ICollection : null;
            var count = collection?.Count ?? 0;
            return count + " " + (count == 1 ? singular : plural);
        }

        private static string TextDescription(IDictionary<string, object> values, string key)
        {
            var text = StringValue(values, key);
            if (string.IsNullOrEmpty(text)) return "empty text";
            return text.Length <= 55 ? Quoted(text.Replace("\r", " ").Replace("\n", " "), "text") : text.Length + " characters of text";
        }

        private static string RangeDescription(IDictionary<string, object> values)
        {
            var start = StringValue(values, "start");
            var end = StringValue(values, "end");
            return string.IsNullOrWhiteSpace(start) || string.IsNullOrWhiteSpace(end) ? "the current selection" : start + "–" + end;
        }

        private static string FormatMode(IDictionary<string, object> values, string key)
        {
            var mode = StringValue(values, key);
            return string.IsNullOrWhiteSpace(mode) ? string.Empty : " (" + FriendlyEnum(mode, mode) + ")";
        }

        private static string FormatStyle(IDictionary<string, object> values)
        {
            var style = StringValue(values, "style");
            if (string.IsNullOrWhiteSpace(style)) return string.Empty;
            if (style.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase)) style = style.Substring("TableStyle".Length);
            var split = style.StartsWith("Medium", StringComparison.OrdinalIgnoreCase) ? "Medium " + style.Substring("Medium".Length) : style;
            return " using the " + split + " style";
        }

        private static string FormatTitle(IDictionary<string, object> values)
        {
            var title = StringValue(values, "title");
            return string.IsNullOrWhiteSpace(title) ? string.Empty : " titled " + Quoted(title, title);
        }

        private static string AtPosition(IDictionary<string, object> values)
        {
            var position = StringValue(values, "position");
            return string.IsNullOrWhiteSpace(position) ? string.Empty : " at position " + position;
        }

        private static string FindReplaceScope(IDictionary<string, object> values, string toolName)
        {
            if (!toolName.StartsWith("powerpoint_", StringComparison.OrdinalIgnoreCase)) return " in the document";
            var slide = StringValue(values, "slide_number");
            return string.IsNullOrWhiteSpace(slide) ? " across the presentation" : " on slide " + slide;
        }

        private static string Truncate(string value, int length)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= length) return value;
            return value.Substring(0, Math.Max(0, length - 1)) + "…";
        }

        private static string PrettyPrintJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json.Trim() == "{}") return string.Empty;
            var output = new StringBuilder();
            var indent = 0;
            var quoted = false;
            var escaped = false;
            foreach (var character in json.Trim())
            {
                if (quoted)
                {
                    output.Append(character);
                    if (escaped) escaped = false;
                    else if (character == '\\') escaped = true;
                    else if (character == '"') quoted = false;
                    continue;
                }

                switch (character)
                {
                    case '"': quoted = true; output.Append(character); break;
                    case '{':
                    case '[': output.Append(character).AppendLine(); indent++; output.Append(' ', indent * 2); break;
                    case '}':
                    case ']': output.AppendLine(); indent = Math.Max(0, indent - 1); output.Append(' ', indent * 2).Append(character); break;
                    case ',': output.Append(character).AppendLine(); output.Append(' ', indent * 2); break;
                    case ':': output.Append(": "); break;
                    default: if (!char.IsWhiteSpace(character)) output.Append(character); break;
                }
            }
            return output.ToString();
        }

        private sealed class OfficePermissionContent
        {
            public string Heading { get; set; }
            public string ConfirmationLabel { get; set; }
            public string Summary { get; set; }
            public string TechnicalDetails { get; set; }
        }
    }
}
