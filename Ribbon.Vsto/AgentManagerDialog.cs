using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class AgentManagerDialog : Form
    {
        private readonly VstoHostRuntime _runtime;
        private readonly RibbonPalette _palette;
        private readonly ListView _list = new ListView();
        private readonly TextBox _search = new TextBox();
        private readonly RibbonButton _toggle;
        private readonly RibbonButton _refresh;
        private readonly RibbonButton _close;
        private readonly Label _status = new Label();
        private readonly RibbonStatusDot _statusDot = new RibbonStatusDot();
        private readonly Label _detailTitle = new Label();
        private readonly Label _description = new Label();
        private readonly Label _metadata = new Label();
        private readonly List<AgentSummary> _allAgents = new List<AgentSummary>();
        private readonly ImageList _rowHeight = new ImageList();
        private readonly ToolTip _toolTip = new ToolTip { InitialDelay = 450, ReshowDelay = 100, AutoPopDelay = 8000 };

        public AgentManagerDialog(VstoHostRuntime runtime, RibbonPalette palette = null)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            _palette = palette ?? RibbonPalette.Detect();
            _toggle = new RibbonButton(_palette, RibbonButtonKind.Primary) { Text = "Install", Glyph = RibbonGlyph.Download, Width = 112, Enabled = false };
            _refresh = new RibbonButton(_palette, RibbonButtonKind.Secondary) { Text = "Refresh", Glyph = RibbonGlyph.Refresh, Width = 104 };
            _close = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "Close", Width = 84, DialogResult = DialogResult.Cancel };

            Text = "Ribbon · ACP Agent Registry";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(780, 520);
            Size = new Size(940, 640);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = _palette.Background;
            ForeColor = _palette.Text;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            BuildLayout();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _rowHeight.Dispose();
                _toolTip.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            RibbonWindowChrome.Apply(this, _palette);
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            FitColumns();
            await ReloadAsync();
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(18, 16, 18, 14),
                BackColor = _palette.Background
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 92));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildToolbar(), 0, 1);
            root.Controls.Add(BuildList(), 0, 2);
            root.Controls.Add(BuildDetails(), 0, 3);
            root.Controls.Add(BuildFooter(), 0, 4);
            Controls.Add(root);
            CancelButton = _close;
        }

        private Control BuildHeader()
        {
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = _palette.Background };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(_palette) { Margin = new Padding(2, 4, 10, 0) };
            var titles = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = _palette.Background };
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 29));
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            titles.Controls.Add(LabelFor("Agent Registry", 14f, FontStyle.Bold, _palette.Text, _palette.Background), 0, 0);
            titles.Controls.Add(LabelFor("Discover and manage ACP agents for every Ribbon workspace", 9f, FontStyle.Regular, _palette.MutedText, _palette.Background), 0, 1);
            header.Controls.Add(mark, 0, 0);
            header.Controls.Add(titles, 1, 0);
            return header;
        }

        private Control BuildToolbar()
        {
            var toolbar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, BackColor = _palette.Background, Margin = new Padding(0, 0, 0, 8) };
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var searchSurface = new RibbonSurface(_palette)
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 7, 12, 5),
                Margin = new Padding(0, 3, 0, 3),
                CornerRadius = 7,
                UseRaisedBackground = true
            };
            _search.Dock = DockStyle.Fill;
            _search.BorderStyle = BorderStyle.None;
            _search.BackColor = _palette.SurfaceRaised;
            _search.ForeColor = _palette.Text;
            _search.Font = new Font(Font.FontFamily, 9.25f, FontStyle.Regular);
            _search.TextChanged += (sender, args) => ApplyFilter();
            _search.KeyDown += SearchOnKeyDown;
            _search.Enter += (sender, args) => searchSurface.EmphasizeBorder = true;
            _search.Leave += (sender, args) => searchSurface.EmphasizeBorder = false;
            _search.AccessibleName = "Search ACP agents";
            RibbonCue.Set(_search, "Search by agent name, id, or description");
            searchSurface.Controls.Add(_search);
            toolbar.Controls.Add(searchSurface, 0, 0);
            return toolbar;
        }

        private Control BuildList()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Padding = new Padding(1), Margin = new Padding(0, 0, 0, 10), CornerRadius = 8 };
            _list.Dock = DockStyle.Fill;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.HideSelection = false;
            _list.MultiSelect = false;
            _list.BorderStyle = BorderStyle.None;
            _list.BackColor = _palette.Surface;
            _list.ForeColor = _palette.Text;
            _list.Font = new Font(Font.FontFamily, 9f, FontStyle.Regular);
            _list.OwnerDraw = true;
            _rowHeight.ImageSize = new Size(1, 28);
            _rowHeight.ColorDepth = ColorDepth.Depth8Bit;
            _list.SmallImageList = _rowHeight;
            _list.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            _list.Columns.Add("Agent", 180);
            _list.Columns.Add("Version", 80);
            _list.Columns.Add("Distribution", 90);
            _list.Columns.Add("Status", 110);
            _list.Columns.Add("Description", 320);
            _list.DrawColumnHeader += DrawColumnHeader;
            _list.DrawItem += (sender, args) => { };
            _list.DrawSubItem += DrawSubItem;
            _list.SelectedIndexChanged += (sender, args) => UpdateSelection();
            _list.DoubleClick += async (sender, args) => { if (_toggle.Enabled) await ToggleAsync(); };
            _list.Resize += (sender, args) => FitColumns();
            _list.AccessibleName = "Compatible ACP agents";
            RibbonNativeTheme.ApplyDarkScrollBars(_list, _palette);
            surface.Controls.Add(_list);
            return surface;
        }

        private Control BuildDetails()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Padding = new Padding(12, 8, 12, 7), Margin = new Padding(0, 0, 0, 8), CornerRadius = 8 };
            var details = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            details.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            details.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            details.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            _detailTitle.Dock = DockStyle.Fill;
            _detailTitle.AutoEllipsis = true;
            _detailTitle.Text = "Agent details";
            _detailTitle.ForeColor = _palette.Text;
            _detailTitle.BackColor = _palette.Surface;
            _detailTitle.Font = new Font(Font.FontFamily, 9.25f, FontStyle.Bold);
            _detailTitle.TextAlign = ContentAlignment.MiddleLeft;
            _description.Dock = DockStyle.Fill;
            _description.AutoEllipsis = true;
            _description.Text = "Select an agent to see its details.";
            _description.ForeColor = _palette.Text;
            _description.BackColor = _palette.Surface;
            _description.TextAlign = ContentAlignment.MiddleLeft;
            _metadata.Dock = DockStyle.Fill;
            _metadata.AutoEllipsis = true;
            _metadata.ForeColor = _palette.MutedText;
            _metadata.BackColor = _palette.Surface;
            _metadata.Font = new Font(Font.FontFamily, 8f, FontStyle.Regular);
            _metadata.TextAlign = ContentAlignment.MiddleLeft;
            details.Controls.Add(_detailTitle, 0, 0);
            details.Controls.Add(_description, 0, 1);
            details.Controls.Add(_metadata, 0, 2);
            surface.Controls.Add(details);
            return surface;
        }

        private Control BuildFooter()
        {
            var footer = new RibbonLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, BackColor = _palette.Background };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 18));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _statusDot.Anchor = AnchorStyles.Left;
            _statusDot.DotColor = _palette.MutedText;
            _status.Dock = DockStyle.Fill;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.AutoEllipsis = true;
            _status.ForeColor = _palette.MutedText;
            _status.Font = new Font(Font.FontFamily, 8.5f, FontStyle.Regular);
            _refresh.Dock = DockStyle.Fill;
            _refresh.Margin = new Padding(0, 6, 8, 5);
            _refresh.Click += async (sender, args) => await ReloadAsync();
            _toggle.Dock = DockStyle.Fill;
            _toggle.Margin = new Padding(0, 6, 0, 5);
            _toggle.Click += async (sender, args) => await ToggleAsync();
            _close.Dock = DockStyle.Fill;
            _close.Margin = new Padding(0, 6, 8, 5);
            _refresh.AccessibleName = "Refresh ACP Registry";
            _toggle.AccessibleName = "Install, update, or uninstall selected agent";
            _close.AccessibleName = "Close Agent Registry";
            _toolTip.SetToolTip(_refresh, "Download the latest compatible agents from the ACP Registry.");
            _toolTip.SetToolTip(_close, "Close the Agent Registry.");
            footer.Controls.Add(_statusDot, 0, 0);
            footer.Controls.Add(_status, 1, 0);
            footer.Controls.Add(_refresh, 2, 0);
            footer.Controls.Add(_close, 3, 0);
            footer.Controls.Add(_toggle, 4, 0);
            return footer;
        }

        private async Task ReloadAsync()
        {
            SetBusy(true, "Loading ACP Registry…");
            try
            {
                var agents = await _runtime.GetRegistryAgentsAsync();
                var orderedAgents = agents.OrderBy(item => item.Name).ToList();
                RibbonUiThread.Run(this, () =>
                {
                    _allAgents.Clear();
                    _allAgents.AddRange(orderedAgents);
                    ApplyFilter();
                    SetStatusCore(_allAgents.Count + " compatible agents", _palette.Success);
                });
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
            finally
            {
                RibbonUiThread.Run(this, () =>
                {
                    SetBusyCore(false, _status.Text);
                    UpdateSelection();
                });
            }
        }

        private void ApplyFilter()
        {
            var selectedId = SelectedAgent()?.Id;
            var query = _search.Text.Trim();
            var filtered = string.IsNullOrWhiteSpace(query)
                ? _allAgents
                : _allAgents.Where(agent => Contains(agent.Name, query) || Contains(agent.Description, query) || Contains(agent.Id, query)).ToList();
            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var agent in filtered)
            {
                var row = new ListViewItem(agent.Name ?? agent.Id) { Tag = agent };
                row.SubItems.Add(agent.Version ?? string.Empty);
                row.SubItems.Add(agent.DistributionType ?? string.Empty);
                row.SubItems.Add(agent.Installed ? (agent.UpdateAvailable ? "Update available" : "Installed") : "Available");
                row.SubItems.Add(agent.Description ?? string.Empty);
                _list.Items.Add(row);
                if (agent.Id == selectedId) row.Selected = true;
            }
            _list.EndUpdate();
            if (_list.SelectedItems.Count == 0 && _list.Items.Count > 0) _list.Items[0].Selected = true;
            SetStatus(filtered.Count + " of " + _allAgents.Count + " compatible agents", _palette.Success);
        }

        private async Task ToggleAsync()
        {
            var agent = SelectedAgent();
            if (agent == null) return;
            SetBusy(true, agent.UpdateAvailable
                ? "Updating " + agent.Name + "…"
                : agent.Installed ? "Uninstalling " + agent.Name + "…" : "Installing " + agent.Name + "…");
            try
            {
                if (agent.UpdateAvailable) await _runtime.InstallAgentAsync(agent.Id);
                else if (agent.Installed) await _runtime.UninstallAgentAsync(agent.Id);
                else await _runtime.InstallAgentAsync(agent.Id);
                await ReloadAsync();
            }
            catch (Exception exception)
            {
                ShowError(exception);
                RibbonUiThread.Run(this, () => SetBusyCore(false, _status.Text));
            }
        }

        private void UpdateSelection()
        {
            var agent = SelectedAgent();
            _toggle.Enabled = agent != null;
            _toggle.Text = agent != null && agent.UpdateAvailable ? "Update" : agent != null && agent.Installed ? "Uninstall" : "Install";
            _toggle.Glyph = agent != null && agent.UpdateAvailable ? RibbonGlyph.Refresh : agent != null && agent.Installed ? RibbonGlyph.Remove : RibbonGlyph.Download;
            _toggle.Kind = agent != null && agent.Installed && !agent.UpdateAvailable ? RibbonButtonKind.Danger : RibbonButtonKind.Primary;
            _detailTitle.Text = agent?.Name ?? "Agent details";
            _description.Text = agent?.Description ?? "Select an agent to see its details.";
            _metadata.Text = agent == null
                ? string.Empty
                : string.Join("   ·   ", new[] { agent.Id, "v" + agent.Version, agent.DistributionType, agent.License }.Where(value => !string.IsNullOrWhiteSpace(value)));
            _toolTip.SetToolTip(_toggle, agent == null
                ? "Select an agent first."
                : agent.UpdateAvailable ? "Update " + agent.Name + " to version " + agent.Version + "."
                : agent.Installed ? "Uninstall " + agent.Name + "." : "Install " + agent.Name + ".");
        }

        private AgentSummary SelectedAgent()
        {
            return _list.SelectedItems.Count == 0 ? null : _list.SelectedItems[0].Tag as AgentSummary;
        }

        private void SetBusy(bool busy, string status)
        {
            RibbonUiThread.Run(this, () => SetBusyCore(busy, status));
        }

        private void SetBusyCore(bool busy, string status)
        {
            _list.Enabled = !busy;
            _search.Enabled = !busy;
            _toggle.Enabled = !busy && SelectedAgent() != null;
            _refresh.Enabled = !busy;
            _close.Enabled = !busy;
            UseWaitCursor = busy;
            SetStatusCore(status, busy ? _palette.Accent : _statusDot.DotColor);
        }

        private void SetStatus(string text, Color color)
        {
            RibbonUiThread.Run(this, () => SetStatusCore(text, color));
        }

        private void SetStatusCore(string text, Color color)
        {
            _status.Text = text;
            _statusDot.DotColor = color;
            _statusDot.Invalidate();
        }

        private void ShowError(Exception exception)
        {
            var message = exception.GetBaseException().Message;
            RibbonUiThread.Run(this, () =>
            {
                SetStatusCore(message, _palette.Danger);
                MessageBox.Show(this, message, "ACP Agent Registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            });
        }

        private void DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            using (var background = new SolidBrush(_palette.SurfaceRaised)) e.Graphics.FillRectangle(background, e.Bounds);
            using (var headerFont = new Font(Font.FontFamily, 8.25f, FontStyle.Bold))
            {
                TextRenderer.DrawText(e.Graphics, e.Header.Text, headerFont,
                    Rectangle.Inflate(e.Bounds, -9, 0), _palette.MutedText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            }
            using (var pen = new Pen(_palette.Border)) e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }

        private void DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            var selected = e.Item.Selected;
            var rowColor = e.Item.Index % 2 == 0
                ? _palette.Surface
                : RibbonDrawing.Blend(_palette.SurfaceRaised, _palette.Surface, _palette.IsDark ? 0.58f : 0.72f);
            var backgroundColor = selected ? RibbonDrawing.Blend(_palette.Accent, _palette.Surface, _palette.IsDark ? 0.72f : 0.86f) : rowColor;
            using (var background = new SolidBrush(backgroundColor)) e.Graphics.FillRectangle(background, e.Bounds);
            var textColor = _palette.Text;
            if (e.ColumnIndex == 3)
            {
                textColor = e.SubItem.Text == "Installed" ? _palette.Success : e.SubItem.Text == "Update available" ? _palette.Accent : _palette.MutedText;
            }
            var font = e.ColumnIndex == 0 ? new Font(Font.FontFamily, 9f, FontStyle.Bold) : Font;
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, font, Rectangle.Inflate(e.Bounds, -9, 0), textColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
            if (e.ColumnIndex == 0) font.Dispose();
            using (var pen = new Pen(RibbonDrawing.Blend(_palette.Border, _palette.Surface, 0.45f)))
            {
                e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            }
        }

        private void FitColumns()
        {
            if (_list.Columns.Count < 5) return;
            var available = _list.ClientSize.Width;
            if (available < 300) return;
            _list.Columns[0].Width = available * 23 / 100;
            _list.Columns[1].Width = available * 10 / 100;
            _list.Columns[2].Width = available * 12 / 100;
            _list.Columns[3].Width = available * 15 / 100;
            _list.Columns[4].Width = available
                - _list.Columns[0].Width
                - _list.Columns[1].Width
                - _list.Columns[2].Width
                - _list.Columns[3].Width;
        }

        private void SearchOnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Escape || _search.TextLength == 0) return;
            _search.Clear();
            e.SuppressKeyPress = true;
            e.Handled = true;
        }

        private static bool Contains(string value, string query)
        {
            return !string.IsNullOrWhiteSpace(value) && value.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private Label LabelFor(string text, float size, FontStyle style, Color foreground, Color background)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = foreground,
                BackColor = background,
                Font = new Font(Font.FontFamily, size, style),
                Margin = new Padding(0)
            };
        }
    }
}
