using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class ConversationHistoryDialog : Form
    {
        private readonly HostRegistration _registration;
        private readonly RibbonPalette _palette;
        private readonly ListView _list = new ListView();
        private readonly TextBox _search = new TextBox();
        private readonly CheckBox _showAll = new CheckBox();
        private readonly Label _details = new Label();
        private readonly Label _status = new Label();
        private readonly RibbonStatusDot _statusDot = new RibbonStatusDot();
        private readonly RibbonButton _open;
        private readonly RibbonButton _delete;
        private readonly RibbonButton _close;
        private readonly ImageList _rowHeight = new ImageList();
        private IList<ConversationRecord> _records = new List<ConversationRecord>();

        public ConversationHistoryDialog(HostRegistration registration, RibbonPalette palette)
        {
            _registration = registration ?? throw new ArgumentNullException(nameof(registration));
            _palette = palette ?? RibbonPalette.Detect();
            _open = new RibbonButton(_palette, RibbonButtonKind.Primary) { Text = "Open", Width = 92, Enabled = false };
            _delete = new RibbonButton(_palette, RibbonButtonKind.Danger) { Text = "Delete", Glyph = RibbonGlyph.Remove, Width = 96, Enabled = false };
            _close = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "Close", Width = 84, DialogResult = DialogResult.Cancel };

            Text = "Ribbon · Conversation History";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 500);
            Size = new Size(920, 620);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = _palette.Background;
            ForeColor = _palette.Text;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            DoubleBuffered = true;
            ResizeRedraw = true;
            BuildLayout();
        }

        public ConversationRecord SelectedConversation { get; private set; }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _rowHeight.Dispose();
            base.Dispose(disposing);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            RibbonWindowChrome.Apply(this, _palette);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _records = ConversationHistoryStorage.LoadAll();
            ApplyFilter();
            FitColumns();
        }

        private void BuildLayout()
        {
            var root = new RibbonLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(18, 16, 18, 14),
                BackColor = _palette.Background
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildFilters(), 0, 1);
            root.Controls.Add(BuildList(), 0, 2);
            root.Controls.Add(BuildDetails(), 0, 3);
            root.Controls.Add(BuildFooter(), 0, 4);
            Controls.Add(root);
            AcceptButton = _open;
            CancelButton = _close;
        }

        private Control BuildHeader()
        {
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = _palette.Background };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(_palette) { Margin = new Padding(2, 4, 10, 0) };
            var titles = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = _palette.Background };
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 29));
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            titles.Controls.Add(LabelFor("Conversation History", 14f, FontStyle.Bold, _palette.Text, _palette.Background), 0, 0);
            titles.Controls.Add(LabelFor("Reopen Ribbon chats or review work from another Office document", 9f, FontStyle.Regular, _palette.MutedText, _palette.Background), 0, 1);
            header.Controls.Add(mark, 0, 0);
            header.Controls.Add(titles, 1, 0);
            return header;
        }

        private Control BuildFilters()
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = _palette.Background };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            var searchSurface = new RibbonSurface(_palette)
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 7, 12, 5),
                CornerRadius = 7,
                UseRaisedBackground = true
            };
            _search.Dock = DockStyle.Fill;
            _search.BorderStyle = BorderStyle.None;
            _search.BackColor = _palette.SurfaceRaised;
            _search.ForeColor = _palette.Text;
            _search.TextChanged += (sender, args) => ApplyFilter();
            _search.Enter += (sender, args) => searchSurface.EmphasizeBorder = true;
            _search.Leave += (sender, args) => searchSurface.EmphasizeBorder = false;
            _search.AccessibleName = "Search Ribbon conversation history";
            RibbonCue.Set(_search, "Search by title, agent, model, or document");
            searchSurface.Controls.Add(_search);
            _showAll.Text = "Show conversations from all documents";
            _showAll.AutoSize = true;
            _showAll.ForeColor = _palette.MutedText;
            _showAll.BackColor = _palette.Background;
            _showAll.CheckedChanged += (sender, args) => ApplyFilter();
            layout.Controls.Add(searchSurface, 0, 0);
            layout.Controls.Add(_showAll, 0, 1);
            return layout;
        }

        private Control BuildList()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Padding = new Padding(1), Margin = new Padding(0, 2, 0, 10), CornerRadius = 8 };
            _list.Dock = DockStyle.Fill;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.HideSelection = false;
            _list.MultiSelect = false;
            _list.BorderStyle = BorderStyle.None;
            _list.BackColor = _palette.Surface;
            _list.ForeColor = _palette.Text;
            _list.Font = new Font(Font.FontFamily, 9f);
            _list.OwnerDraw = true;
            _rowHeight.ImageSize = new Size(1, 30);
            _rowHeight.ColorDepth = ColorDepth.Depth8Bit;
            _list.SmallImageList = _rowHeight;
            _list.Columns.Add("Conversation", 260);
            _list.Columns.Add("Agent", 140);
            _list.Columns.Add("Document", 180);
            _list.Columns.Add("Updated", 130);
            _list.Columns.Add("Continuity", 110);
            _list.DrawColumnHeader += DrawColumnHeader;
            _list.DrawItem += (sender, args) => { };
            _list.DrawSubItem += DrawSubItem;
            _list.SelectedIndexChanged += (sender, args) => UpdateSelection();
            _list.DoubleClick += (sender, args) => OpenSelected();
            _list.Resize += (sender, args) => FitColumns();
            _list.AccessibleName = "Saved Ribbon conversations";
            RibbonNativeTheme.ApplyDarkScrollBars(_list, _palette);
            surface.Controls.Add(_list);
            return surface;
        }

        private Control BuildDetails()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Padding = new Padding(12, 8, 12, 7), Margin = new Padding(0, 0, 0, 8), CornerRadius = 8 };
            _details.Dock = DockStyle.Fill;
            _details.AutoEllipsis = true;
            _details.Text = "Select a conversation to see where it belongs and how it can be reopened.";
            _details.ForeColor = _palette.MutedText;
            _details.BackColor = _palette.Surface;
            _details.TextAlign = ContentAlignment.MiddleLeft;
            surface.Controls.Add(_details);
            return surface;
        }

        private Control BuildFooter()
        {
            var footer = new RibbonLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, BackColor = _palette.Background };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 18));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _statusDot.Anchor = AnchorStyles.Left;
            _statusDot.DotColor = _palette.Success;
            _status.Dock = DockStyle.Fill;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.ForeColor = _palette.MutedText;
            _delete.Click += (sender, args) => DeleteSelected();
            _open.Click += (sender, args) => OpenSelected();
            _delete.Dock = DockStyle.Fill;
            _delete.Margin = new Padding(4, 6, 0, 5);
            _close.Dock = DockStyle.Fill;
            _close.Margin = new Padding(4, 6, 0, 5);
            _open.Dock = DockStyle.Fill;
            _open.Margin = new Padding(4, 6, 0, 5);
            footer.Controls.Add(_statusDot, 0, 0);
            footer.Controls.Add(_status, 1, 0);
            footer.Controls.Add(_delete, 2, 0);
            footer.Controls.Add(_close, 3, 0);
            footer.Controls.Add(_open, 4, 0);
            return footer;
        }

        private void ApplyFilter()
        {
            if (!IsHandleCreated) return;
            var query = (_search.Text ?? string.Empty).Trim();
            var filtered = _records.Where(record =>
                (_showAll.Checked || ConversationHistoryStorage.MatchesCurrentDocument(record, _registration))
                && (query.Length == 0
                    || Contains(record.Title, query)
                    || Contains(record.AgentName, query)
                    || Contains(record.ModelName, query)
                    || Contains(record.DocumentName, query)))
                .OrderByDescending(record => record.UpdatedAtUtc)
                .ToList();
            _list.BeginUpdate();
            try
            {
                _list.Items.Clear();
                foreach (var record in filtered)
                {
                    var current = ConversationHistoryStorage.MatchesCurrentDocument(record, _registration);
                    var continuity = current
                        ? record.MayResumeNatively ? "Native" : "Transcript"
                        : "Read only";
                    var item = new ListViewItem(record.DisplayTitle) { Tag = record };
                    item.SubItems.Add(string.IsNullOrWhiteSpace(record.AgentName) ? record.AgentId : record.AgentName);
                    item.SubItems.Add(record.DisplayDocument);
                    item.SubItems.Add(record.UpdatedAtUtc.ToLocalTime().ToString("g"));
                    item.SubItems.Add(continuity);
                    _list.Items.Add(item);
                }
            }
            finally
            {
                _list.EndUpdate();
            }
            _status.Text = filtered.Count + (filtered.Count == 1 ? " conversation" : " conversations");
            UpdateSelection();
        }

        private void UpdateSelection()
        {
            var record = SelectedRecord();
            _open.Enabled = record != null;
            _delete.Enabled = record != null;
            if (record == null)
            {
                _details.Text = "Select a conversation to see where it belongs and how it can be reopened.";
                return;
            }
            var current = ConversationHistoryStorage.MatchesCurrentDocument(record, _registration);
            var behavior = !current
                ? "This belongs to another document and will open read-only."
                : record.MayResumeNatively
                    ? "Ribbon will ask the agent to restore its native ACP context."
                    : "The saved transcript is available; the agent may need a fresh continuation.";
            _details.Text = record.DisplayDocument + " · " + (record.AgentName ?? record.AgentId)
                + (string.IsNullOrWhiteSpace(record.ModelName) ? string.Empty : " · " + record.ModelName)
                + Environment.NewLine + behavior;
        }

        private void OpenSelected()
        {
            var record = SelectedRecord();
            if (record == null) return;
            SelectedConversation = record;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void DeleteSelected()
        {
            var record = SelectedRecord();
            if (record == null) return;
            var answer = MessageBox.Show(this,
                "Delete this saved Ribbon conversation?\r\n\r\n" + record.DisplayTitle + "\r\n\r\nThis does not change the Office document.",
                "Delete conversation", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.OK) return;
            ConversationHistoryStorage.Delete(record);
            _records = _records.Where(item => !string.Equals(item.Id, record.Id, StringComparison.OrdinalIgnoreCase)).ToList();
            ApplyFilter();
        }

        private ConversationRecord SelectedRecord()
        {
            return _list.SelectedItems.Count == 0 ? null : _list.SelectedItems[0].Tag as ConversationRecord;
        }

        private void FitColumns()
        {
            if (_list.Columns.Count != 5 || _list.ClientSize.Width <= 0) return;
            var available = Math.Max(540, _list.ClientSize.Width);
            _list.Columns[4].Width = 105;
            _list.Columns[3].Width = 132;
            _list.Columns[2].Width = Math.Max(145, available / 5);
            _list.Columns[1].Width = Math.Max(120, available / 6);
            _list.Columns[0].Width = Math.Max(190, available - _list.Columns[1].Width - _list.Columns[2].Width - _list.Columns[3].Width - _list.Columns[4].Width);
        }

        private void DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            using (var background = new SolidBrush(_palette.SurfaceRaised)) e.Graphics.FillRectangle(background, e.Bounds);
            using (var headerFont = new Font(Font.FontFamily, 8.25f, FontStyle.Bold))
            {
                TextRenderer.DrawText(e.Graphics, e.Header.Text, headerFont, Rectangle.Inflate(e.Bounds, -9, 0), _palette.MutedText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            }
            using (var pen = new Pen(_palette.Border)) e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }

        private void DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            var row = e.Item.Index % 2 == 0
                ? _palette.Surface
                : RibbonDrawing.Blend(_palette.SurfaceRaised, _palette.Surface, _palette.IsDark ? 0.58f : 0.72f);
            var backgroundColor = e.Item.Selected
                ? RibbonDrawing.Blend(_palette.Accent, _palette.Surface, _palette.IsDark ? 0.72f : 0.86f)
                : row;
            using (var background = new SolidBrush(backgroundColor)) e.Graphics.FillRectangle(background, e.Bounds);
            var textColor = e.ColumnIndex == 4
                ? e.SubItem.Text == "Native" ? _palette.Success : e.SubItem.Text == "Read only" ? _palette.MutedText : _palette.Accent
                : _palette.Text;
            var font = e.ColumnIndex == 0 ? new Font(Font.FontFamily, 9f, FontStyle.Bold) : Font;
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, font, Rectangle.Inflate(e.Bounds, -9, 0), textColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
            if (e.ColumnIndex == 0) font.Dispose();
            using (var pen = new Pen(RibbonDrawing.Blend(_palette.Border, _palette.Surface, 0.45f)))
            {
                e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            }
        }

        private static bool Contains(string value, string query)
        {
            return !string.IsNullOrWhiteSpace(value) && value.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static Label LabelFor(string text, float size, FontStyle style, Color color, Color background)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = color,
                BackColor = background,
                Font = new Font("Segoe UI", size, style),
                Margin = new Padding(0)
            };
        }
    }
}
