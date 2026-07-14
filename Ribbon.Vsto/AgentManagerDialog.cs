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
        private readonly Label _description = new Label();
        private readonly Label _metadata = new Label();
        private readonly List<AgentSummary> _allAgents = new List<AgentSummary>();
        private readonly ImageList _rowHeight = new ImageList();

        public AgentManagerDialog(VstoHostRuntime runtime, RibbonPalette palette = null)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            _palette = palette ?? RibbonPalette.Detect();
            _toggle = new RibbonButton(_palette, RibbonButtonKind.Primary) { Text = "Install", Glyph = RibbonGlyph.Download, Width = 102, Enabled = false };
            _refresh = new RibbonButton(_palette, RibbonButtonKind.Secondary) { Text = "Refresh", Glyph = RibbonGlyph.Refresh, Width = 100 };
            _close = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "Close", Width = 84, DialogResult = DialogResult.OK };

            Text = "Ribbon · ACP Agent Registry";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 500);
            Size = new Size(920, 610);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = _palette.Background;
            ForeColor = _palette.Text;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            BuildLayout();
        }

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

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 78));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildToolbar(), 0, 1);
            root.Controls.Add(BuildList(), 0, 2);
            root.Controls.Add(BuildDetails(), 0, 3);
            root.Controls.Add(BuildFooter(), 0, 4);
            Controls.Add(root);
            AcceptButton = _close;
        }

        private Control BuildHeader()
        {
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, BackColor = _palette.Background };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(_palette) { Margin = new Padding(2, 3, 10, 0) };
            var titles = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = _palette.Background };
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            titles.Controls.Add(LabelFor("Agent Registry", 15f, FontStyle.Bold, _palette.Text, _palette.Background), 0, 0);
            titles.Controls.Add(LabelFor("Discover and manage ACP agents available to every Office application", 8.75f, FontStyle.Regular, _palette.MutedText, _palette.Background), 0, 1);
            header.Controls.Add(mark, 0, 0);
            header.Controls.Add(titles, 1, 0);
            return header;
        }

        private Control BuildToolbar()
        {
            var toolbar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, BackColor = _palette.Background, Margin = new Padding(0, 0, 0, 8) };
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));

            var searchSurface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Padding = new Padding(10, 7, 10, 5), Margin = new Padding(0, 3, 12, 3), CornerRadius = 7 };
            _search.Dock = DockStyle.Fill;
            _search.BorderStyle = BorderStyle.None;
            _search.BackColor = _palette.Surface;
            _search.ForeColor = _palette.Text;
            _search.Font = new Font(Font.FontFamily, 9.25f, FontStyle.Regular);
            _search.TextChanged += (sender, args) => ApplyFilter();
            RibbonCue.Set(_search, "Search by agent name, id, or description");
            searchSurface.Controls.Add(_search);
            toolbar.Controls.Add(searchSurface, 0, 0);

            var count = LabelFor("PUBLIC ACP REGISTRY", 7.5f, FontStyle.Bold, _palette.Accent, _palette.Background);
            count.TextAlign = ContentAlignment.MiddleRight;
            toolbar.Controls.Add(count, 1, 0);
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
            _list.Columns.Add("Agent", 190);
            _list.Columns.Add("Version", 86);
            _list.Columns.Add("Distribution", 96);
            _list.Columns.Add("Status", 122);
            _list.Columns.Add("Description", 330);
            _list.DrawColumnHeader += DrawColumnHeader;
            _list.DrawItem += (sender, args) => { };
            _list.DrawSubItem += DrawSubItem;
            _list.SelectedIndexChanged += (sender, args) => UpdateSelection();
            _list.DoubleClick += async (sender, args) => { if (_toggle.Enabled) await ToggleAsync(); };
            _list.Resize += (sender, args) => FitDescriptionColumn();
            surface.Controls.Add(_list);
            return surface;
        }

        private Control BuildDetails()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Padding = new Padding(12, 9, 12, 8), Margin = new Padding(0, 0, 0, 8), CornerRadius = 8 };
            var details = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = _palette.Surface };
            details.RowStyles.Add(new RowStyle(SizeType.Percent, 62));
            details.RowStyles.Add(new RowStyle(SizeType.Percent, 38));
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
            details.Controls.Add(_description, 0, 0);
            details.Controls.Add(_metadata, 0, 1);
            surface.Controls.Add(details);
            return surface;
        }

        private Control BuildFooter()
        {
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, BackColor = _palette.Background };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 18));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            _statusDot.Anchor = AnchorStyles.Left;
            _statusDot.DotColor = _palette.MutedText;
            _status.Dock = DockStyle.Fill;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.AutoEllipsis = true;
            _status.ForeColor = _palette.MutedText;
            _status.Font = new Font(Font.FontFamily, 8.5f, FontStyle.Regular);
            _refresh.Dock = DockStyle.Fill;
            _refresh.Margin = new Padding(4, 6, 4, 5);
            _refresh.Click += async (sender, args) => await ReloadAsync();
            _toggle.Dock = DockStyle.Fill;
            _toggle.Margin = new Padding(4, 6, 4, 5);
            _toggle.Click += async (sender, args) => await ToggleAsync();
            _close.Dock = DockStyle.Fill;
            _close.Margin = new Padding(4, 6, 0, 5);
            footer.Controls.Add(_statusDot, 0, 0);
            footer.Controls.Add(_status, 1, 0);
            footer.Controls.Add(_refresh, 2, 0);
            footer.Controls.Add(_toggle, 3, 0);
            footer.Controls.Add(_close, 4, 0);
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
            SetBusy(true, agent.Installed ? "Uninstalling " + agent.Name + "…" : "Installing " + agent.Name + "…");
            try
            {
                if (agent.Installed) await _runtime.UninstallAgentAsync(agent.Id);
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
            _toggle.Text = agent != null && agent.Installed ? "Uninstall" : "Install";
            _toggle.Glyph = agent != null && agent.Installed ? RibbonGlyph.Remove : RibbonGlyph.Download;
            _toggle.Kind = agent != null && agent.Installed ? RibbonButtonKind.Danger : RibbonButtonKind.Primary;
            _description.Text = agent?.Description ?? "Select an agent to see its details.";
            _metadata.Text = agent == null
                ? string.Empty
                : string.Join("   ·   ", new[] { agent.Id, "v" + agent.Version, agent.DistributionType, agent.License }.Where(value => !string.IsNullOrWhiteSpace(value)));
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
            var backgroundColor = selected ? RibbonDrawing.Blend(_palette.Accent, _palette.Surface, _palette.IsDark ? 0.72f : 0.86f) : _palette.Surface;
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

        private void FitDescriptionColumn()
        {
            if (_list.Columns.Count < 5) return;
            var fixedWidth = _list.Columns[0].Width + _list.Columns[1].Width + _list.Columns[2].Width + _list.Columns[3].Width;
            _list.Columns[4].Width = Math.Max(180, _list.ClientSize.Width - fixedWidth - 4);
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
